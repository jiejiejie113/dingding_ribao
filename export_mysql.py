"""将 MySQL 数据库的所有基础表逐表导出为 Excel。"""

import argparse
import base64
import getpass
import logging
import math
import re
import sys
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from pathlib import Path

try:
    import pymysql
    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell
except ImportError:
    print("缺少依赖，请先运行：python -m pip install -r requirements.txt")
    sys.exit(1)


MAX_DATA_ROWS = 1_048_575  # 每个工作表预留一行表头
TABLE_FILENAMES = {
    "dept_follow_up": "跟单",
    "dept_alloc": "纸样分配",
    "dept_pattern": "纸样",
    "dept_cut": "裁床",
    "dept_sew": "车位",
}
ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


def quote_identifier(name):
    return "`" + name.replace("`", "``") + "`"


def safe_filename(name):
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return (cleaned or "table")[:100]


def excel_value(value, stats):
    if isinstance(value, bytes):
        value = "base64:" + base64.b64encode(value).decode("ascii")
    elif isinstance(value, Decimal):
        value = str(value)  # 保留金额等精确小数
    elif isinstance(value, int) and abs(value) >= 10**15:
        value = str(value)  # 避免 Excel 的 15 位数字精度限制
    elif isinstance(value, float) and not math.isfinite(value):
        value = str(value)
    elif isinstance(value, (datetime, date, time, timedelta)):
        value = str(value)
    elif value is not None and not isinstance(value, (str, int, float, bool)):
        value = str(value)
    if isinstance(value, str):
        cleaned = ILLEGAL_XML.sub("", value)
        if cleaned != value:
            stats["cleaned"] += 1
        if len(cleaned) > 32767:
            stats["truncated"] += 1
            cleaned = cleaned[:32767]
        return cleaned
    return value


def append_row(sheet, values, stats):
    cells = []
    for value in values:
        cell = WriteOnlyCell(sheet, value=excel_value(value, stats))
        if isinstance(cell.value, str):
            # 所有数据库字符串按文本写入，避免以 = 开头的内容变成公式。
            cell.data_type = "s"
        cells.append(cell)
    sheet.append(cells)


def export_table(connection, table, destination, batch_size):
    workbook = Workbook(write_only=True)
    stats = {"rows": 0, "sheets": 0, "cleaned": 0, "truncated": 0}
    partial = destination.with_suffix(".partial.xlsx")
    try:
        with connection.cursor(pymysql.cursors.SSCursor) as cursor:
            cursor.execute("SELECT * FROM " + quote_identifier(table))
            headers = [column[0] for column in cursor.description]
            if len(headers) > 16384:
                raise ValueError("列数超过 Excel 的 16384 列限制")

            def new_sheet():
                stats["sheets"] += 1
                sheet = workbook.create_sheet("数据_" + str(stats["sheets"]))
                sheet.freeze_panes = "A2"
                append_row(sheet, headers, stats)
                return sheet

            sheet = new_sheet()
            sheet_rows = 0
            while True:
                rows = cursor.fetchmany(batch_size)
                if not rows:
                    break
                for row in rows:
                    if sheet_rows == MAX_DATA_ROWS:
                        sheet = new_sheet()
                        sheet_rows = 0
                    append_row(sheet, row, stats)
                    sheet_rows += 1
                    stats["rows"] += 1
                if stats["rows"] % (batch_size * 20) == 0:
                    logging.info("  %s 已读取 %s 行", table, stats["rows"])
        workbook.save(partial)
        partial.replace(destination)
        return stats
    finally:
        # write_only 模式下，异常时也结束 XML 写入器，释放临时资源。
        for sheet in workbook.worksheets:
            if not sheet.closed:
                sheet.close()
        workbook.close()


def parse_args():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--host", default="192.168.1.248")
    parser.add_argument("--port", type=int, default=3306)
    parser.add_argument("--database", default="yida_data")
    parser.add_argument("--user", help="数据库用户名；未指定时交互输入")
    parser.add_argument("--output", type=Path, default=Path(__file__).resolve().parent / "excel-dingding")
    parser.add_argument("--batch-size", type=int, default=2000)
    parser.add_argument("--read-timeout", type=int, default=600, help="单次读取超时秒数")
    parser.add_argument("--process", action="store_true", help="导出成功后合并业务表并生成透视表")
    args = parser.parse_args()
    if args.batch_size <= 0 or args.read_timeout <= 0:
        parser.error("batch-size 和 read-timeout 必须大于 0")
    return args


def main():
    args = parse_args()
    user = args.user or input("MySQL 用户名：").strip()
    if not user:
        print("用户名不能为空。")
        return 1
    password = get_saved_password(args.host, args.port, user)
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(output / "export.log", encoding="utf-8"), logging.StreamHandler()],
    )
    options = dict(
        host=args.host, port=args.port, user=user, password=password,
        database=args.database, charset="utf8mb4", connect_timeout=15,
        read_timeout=args.read_timeout, write_timeout=60, autocommit=True,
    )
    logging.info("连接 %s:%s / %s", args.host, args.port, args.database)
    try:
        with pymysql.connect(**options) as connection:
            with connection.cursor() as cursor:
                cursor.execute(
                    "SELECT TABLE_NAME FROM information_schema.TABLES "
                    "WHERE TABLE_SCHEMA = %s AND TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME",
                    (args.database,),
                )
                tables = [row[0] for row in cursor.fetchall()]
    except pymysql.MySQLError as exc:
        logging.error("连接或获取表清单失败：%s", exc)
        logging.error("请检查内网连接、端口、账号密码和数据库 SELECT 权限。日志：%s", output)
        return 1

    logging.info("发现 %s 张基础表，输出目录：%s", len(tables), output)
    results = []
    for index, table in enumerate(tables, 1):
        filename = TABLE_FILENAMES.get(table, f"{index:04d}_{safe_filename(table)}") + ".xlsx"
        logging.info("[%s/%s] 导出 %s", index, len(tables), table)
        try:
            # 每张表独立连接，某张表断线后不影响后续表尝试导出。
            with pymysql.connect(**options) as connection:
                stats = export_table(connection, table, output / filename, args.batch_size)
            results.append([table, "成功", stats["rows"], stats["sheets"], filename,
                            stats["cleaned"], stats["truncated"], ""])
            logging.info("完成 %s：%s 行，%s 个工作表", table, stats["rows"], stats["sheets"])
            if stats["cleaned"] or stats["truncated"]:
                logging.warning("%s：%s 个单元格移除非法字符，%s 个单元格因 Excel 长度限制截断",
                                table, stats["cleaned"], stats["truncated"])
        except Exception as exc:
            logging.exception("导出失败 %s", table)
            results.append([table, "失败", None, None, "", None, None, str(exc)])

    summary = Workbook(write_only=True)
    sheet = summary.create_sheet("导出结果")
    summary_stats = {"cleaned": 0, "truncated": 0}
    append_row(sheet, ["表名", "状态", "数据行数", "工作表数", "文件名", "清理字符单元格数", "截断单元格数", "错误信息"], summary_stats)
    for row in results:
        append_row(sheet, row, summary_stats)
    summary.save(output / "导出汇总.xlsx")
    failed = sum(row[1] == "失败" for row in results)
    logging.info("导出结束：成功 %s 张，失败 %s 张。目录：%s", len(results) - failed, failed, output)
    if not tables:
        logging.warning("未发现可见的基础表，请确认数据库名及账号权限。")
    if args.process and not failed and tables:
        missing = set(TABLE_FILENAMES) - set(tables)
        if missing:
            logging.error("本次未导出以下业务表，停止合并以避免使用旧文件：%s", ", ".join(sorted(missing)))
            return 1
        from process_excel import process
        process(output)
    return 1 if failed else 0


def get_saved_password(host, port, user):
    """通过 Windows 凭据管理器为当前用户保存和读取密码。"""
    import win32cred
    import pywintypes
    target = f"DingdingExcel/MySQL/{host}:{port}/{user}"
    try:
        credential = win32cred.CredRead(target, win32cred.CRED_TYPE_GENERIC)
        blob = credential["CredentialBlob"]
        return blob.decode("utf-16-le") if isinstance(blob, bytes) else blob
    except pywintypes.error as exc:
        if exc.winerror != 1168:
            raise
    password = getpass.getpass("MySQL 密码（首次输入后保存到 Windows 凭据管理器）：")
    win32cred.CredWrite({
        "Type": win32cred.CRED_TYPE_GENERIC,
        "TargetName": target,
        "UserName": user,
        "CredentialBlob": password,
        "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
    })
    return password


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n用户已取消导出。")
        sys.exit(130)
    except Exception:
        logging.exception("执行失败")
        sys.exit(1)
