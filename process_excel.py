"""按数据透视表1.7月份版的业务规则合并五张表，并刷新 Excel 模板。"""
import argparse
import ast
import importlib.util
import logging
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from report_period import report_period

ROOT = Path(__file__).resolve().parent
REFERENCE = ROOT / "1.7数据透视表（月份）.py"
TEMPLATE = ROOT / "2026-10-16 模板.xlsx"


def load_rules():
    spec = importlib.util.spec_from_file_location("pivot_rules_17", REFERENCE)
    rules = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(rules)
    # 直接读取参考程序的列清单，避免两份业务字段定义逐渐不一致。
    tree = ast.parse(REFERENCE.read_text(encoding="utf-8-sig"))
    main = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "main")
    columns = {}
    for node in main.body:
        if isinstance(node, ast.Assign) and isinstance(node.value, ast.List):
            for target in node.targets:
                if isinstance(target, ast.Name) and (target.id.endswith("_columns") or target.id == "merge_keys"):
                    columns[target.id] = ast.literal_eval(node.value)
    return rules, columns


def merge_files(folder, rules, columns):
    frames = []
    for filename, key in [
        ("跟单", "follow_up"), ("纸样分配", "allocation"),
        ("纸样", "pattern_status"), ("裁床", "cutting"), ("车位", "sewing"),
    ]:
        path = folder / (filename + ".xlsx")
        # 读取全部分片工作表，保留标识字段的文本形式。
        sheets = pd.read_excel(path, sheet_name=None, dtype=object)
        frame = pd.concat(sheets.values(), ignore_index=True)
        frame.columns = frame.columns.str.strip()
        frame = frame.rename(columns={
            "品牌/部门": "品牌&部门", "计划件数": "计划制衣件数",
            "计划裁剪件数": "计划制衣件数", "纸样完成状态": "纸样完成进度",
            "裁床完成状态": "裁床完成进度", "裁剪数量": "裁剪总数",
            "modified_time": "车办改转态日期",
        })
        required = columns[key + "_columns"]
        missing = set(required) - set(frame.columns)
        if missing:
            raise ValueError(f"{filename} 缺少字段：{sorted(missing)}")
        frame = rules.normalize_staff_column(frame[required].copy())
        # 与参考代码最终按业务键保留首条相同，提前去重避免重复键乘法膨胀。
        before = len(frame)
        frame = frame.drop_duplicates(columns["merge_keys"], keep="first")
        logging.info("%s：读取 %s 行，业务键去重后 %s 行", filename, before, len(frame))
        frames.append(frame)
    merged = frames[0]
    for frame in frames[1:]:
        # 关键合并失败必须抛出错误，不使用参考函数的静默回退。
        keys = columns["merge_keys"]
        overlap = (set(merged.columns) & set(frame.columns)) - set(keys)
        merged = pd.merge(merged, frame, on=keys, how="left",
                          suffixes=("_left", "_right"), validate="one_to_one")
        for col in overlap:
            merged[col] = merged[col + "_left"].combine_first(merged[col + "_right"])
            merged = merged.drop(columns=[col + "_left", col + "_right"])
    merged = merged.reindex(columns=columns["output_columns"])
    for col in ["跟办人员", "产品类型", "款式类型"]:
        merged[col] = merged[col].fillna("")
    today = pd.Timestamp(datetime.now().date())
    days = (today - pd.to_datetime(merged["办单提交纸样部日期"], errors="coerce")).dt.days
    remove = merged["制办单分配进度"].isna() & (days > 3) & (merged["是否经过纸样部"] == "需要")
    removed = merged[remove].copy()
    removed["超时天数"] = days[remove]
    removed.to_excel(folder / "剔除数据表.xlsx", index=False)
    merged = merged[~remove].copy()
    merged = rules.add_duration_columns(merged)
    merged = rules.add_date_only_columns(merged)
    for col in ["车办实际总件数", "裁剪总数"]:
        merged[col] = pd.to_numeric(merged[col], errors="coerce").fillna(0)
    abnormal = merged["车办实际总件数"] > merged["裁剪总数"]
    merged = merged.assign(
        _abnormal=abnormal,
        _submitted=pd.to_datetime(merged["办单提交纸样部日期"], errors="coerce"),
    ).sort_values(["_abnormal", "_submitted"], ascending=[False, False], kind="stable", na_position="last")
    merged = merged[columns["output_columns"]].reset_index(drop=True)
    path = folder / "合并数据_完整提取.xlsx"
    merged.to_excel(path, index=False, sheet_name="数据")
    workbook = load_workbook(path)
    sheet = workbook.active
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    fill = PatternFill("solid", fgColor="D4D8DC")
    for row in range(2, int(abnormal.sum()) + 2):
        for cell in sheet[row]:
            cell.fill = fill
    workbook.save(path)
    workbook.close()
    logging.info("合并完成：保留 %s 行，剔除 %s 行，件数异常 %s 行", len(merged), len(removed), int(abnormal.sum()))
    return path, merged


def refresh_report(source, frame, template, output, rules):
    if rules.win32com is None:
        raise RuntimeError("缺少 pywin32，请安装 requirements.txt 中的依赖")
    partial = output.with_name(output.stem + ".partial.xlsx")
    # 仅修正缓存 XML 中的引用，保留模板的透视表、图表和布局。
    from openpyxl.utils import get_column_letter
    with zipfile.ZipFile(template) as original, zipfile.ZipFile(partial, "w", zipfile.ZIP_DEFLATED) as updated:
        for entry in original.infolist():
            content = original.read(entry.filename)
            if entry.filename.startswith("xl/pivotCache/pivotCacheDefinition") and entry.filename.endswith(".xml"):
                # 字节级替换避免重新序列化时改变扩展命名空间。
                import re
                reference = f"A1:{get_column_letter(len(frame.columns))}{len(frame) + 1}"
                content = re.sub(rb'(<worksheetSource\b[^>]*\bref=")[^"]*(")',
                                 lambda match: match[1] + reference.encode() + match[2], content)
            updated.writestr(entry, content)
    excel = rules.win32com.client.DispatchEx("Excel.Application")
    import win32process
    pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
    rules.CURRENT_EXCEL_PID = pid
    workbook = None
    retry = rules.com_retry
    try:
        retry(lambda: setattr(excel, "DisplayAlerts", False))
        retry(lambda: setattr(excel, "Visible", False))
        if not rules.wait_excel_ready(excel):
            raise RuntimeError("Excel 未能就绪")
        workbook = retry(excel.Workbooks.Open, str(partial.resolve()), UpdateLinks=0)
        sheet = retry(lambda: workbook.Worksheets("数据"))
        retry(lambda: sheet.UsedRange.Clear())
        data = frame.copy()
        for col in data.columns:
            if pd.api.types.is_datetime64_any_dtype(data[col]):
                data[col] = data[col].dt.strftime("%Y-%m-%d %H:%M:%S")
        data = data.astype(object).where(pd.notna(data), None)
        for index, col in enumerate(data.columns, 1):
            if "日期" in col or "时间" in col or col in {
                "东明款号", "办名称", "重做版次", "品牌&部门", "设计周",
            }:
                retry(lambda index=index: setattr(sheet.Columns(index), "NumberFormat", "@"))
        values = [list(data.columns)] + data.values.tolist()
        for start in range(0, len(values), 2000):
            chunk = values[start:start + 2000]
            retry(lambda start=start, chunk=chunk: setattr(
                sheet.Range(sheet.Cells(start + 1, 1), sheet.Cells(start + len(chunk), len(data.columns))),
                "Value", chunk))
        # 模板缓存原本引用整列，改为完整字段和当前实际行数。
        caches = retry(lambda: workbook.PivotCaches())
        for index in range(1, retry(lambda: caches.Count) + 1):
            cache = retry(lambda index=index: caches.Item(index))
            retry(lambda: setattr(cache, "MissingItemsLimit", 0))
            retry(cache.Refresh)
        rebuild_staff_pivot(workbook, sheet, len(data), len(data.columns), retry)
        for ws in rules.get_worksheets(workbook):
            for pivot in retry(lambda ws=ws: list(ws.PivotTables())):
                retry(pivot.RefreshTable)
        set_yesterday_filters(workbook, retry)
        fill_brief(workbook, frame, retry)
        abnormal_count = int((frame["车办实际总件数"] > frame["裁剪总数"]).sum())
        if abnormal_count:
            retry(lambda: setattr(sheet.Range(sheet.Cells(2, 1), sheet.Cells(abnormal_count + 1, len(data.columns))).Interior,
                                  "Color", 14473428))
        rules.wait_excel_idle(excel)
        retry(workbook.Save)
        retry(lambda: workbook.Close(False))
        workbook = None
    finally:
        rules.cleanup_excel(excel, workbook, pid)
    output = publish_report(partial, output)
    logging.info("透视表已生成：%s", output)
    return output


def rebuild_staff_pivot(workbook, data_sheet, row_count, column_count, retry):
    """独立重建师傅透视表，移除旧布局及固定J列公式造成的扩展冲突。"""
    staff = retry(lambda: workbook.Worksheets("纸样师傅角度"))
    for pivot in retry(lambda: list(staff.PivotTables())):
        retry(lambda pivot=pivot: pivot.TableRange2.Clear())
    retry(lambda: staff.Cells.UnMerge())
    retry(lambda: staff.Cells.Clear())
    source = f"'数据'!R1C1:R{row_count + 1}C{column_count}"
    cache = retry(lambda: workbook.PivotCaches().Create(SourceType=1, SourceData=source))
    pivot = retry(lambda: cache.CreatePivotTable(TableDestination=staff.Range("A4"), TableName="StaffStatusSummary"))
    row_field = retry(lambda: pivot.PivotFields("纸样师傅"))
    retry(lambda: setattr(row_field, "Orientation", 1))
    col_field = retry(lambda: pivot.PivotFields("纸样完成进度"))
    retry(lambda: setattr(col_field, "Orientation", 2))
    retry(lambda: pivot.AddDataField(pivot.PivotFields("东明款号"), "计数项:东明款号", -4112))
    retry(pivot.RefreshTable)
    # 未完成口径与简报一致，用明细聚合，不再引用易变的B/C/D列。
    last_column = retry(lambda: pivot.TableRange2.Column + pivot.TableRange2.Columns.Count)
    retry(lambda: setattr(staff.Cells(4, last_column + 1), "Value", "师傅"))
    retry(lambda: setattr(staff.Cells(4, last_column + 2), "Value", "未完成纸样（进行中、未处理）"))
    values = retry(lambda: data_sheet.Range(data_sheet.Cells(1, 1), data_sheet.Cells(row_count + 1, column_count)).Value)
    frame = pd.DataFrame(values[1:], columns=values[0])
    status = frame["纸样完成进度"].fillna("").astype(str).str.replace(r"^\d+", "", regex=True).str.strip()
    counts = frame.loc[status.isin(["进行中", "未处理"])].groupby("纸样师傅", dropna=False).size()
    for index, (name, count) in enumerate(counts.items(), 5):
        retry(lambda index=index, name=name: setattr(staff.Cells(index, last_column + 1), "Value", "未指定" if pd.isna(name) else str(name)))
        retry(lambda index=index, count=count: setattr(staff.Cells(index, last_column + 2), "Value", int(count)))
    retry(lambda: staff.UsedRange.Columns.AutoFit())


def publish_report(partial, output):
    try:
        partial.replace(output)
    except PermissionError:
        output = output.with_name(output.stem + datetime.now().strftime("_%H%M%S_%f") + output.suffix)
        partial.replace(output)
        logging.warning("原报表被占用，已另存为：%s", output)
    return output


def fill_brief(workbook, frame, retry):
    """业务流量取已筛选透视表总计；手头存量取当前合并明细。"""
    def totals(sheet_name):
        sheet = retry(lambda: workbook.Worksheets(sheet_name))
        result = []
        for pivot in retry(lambda: list(sheet.PivotTables())):
            values = retry(lambda pivot=pivot: pivot.TableRange1.Value)
            total = next((row for row in reversed(values) if row[0] == "总计"), None)
            if total is None:
                raise ValueError(f"{sheet_name} 未找到透视表总计，停止填写简报")
            result.append((retry(lambda pivot=pivot: pivot.TableRange1.Column), total))
        return [row for _, row in sorted(result)]

    changes = totals("办单数量变化")
    if len(changes) != 2:
        raise ValueError("办单数量变化应包含新增办单、纸样完成两张透视表")
    cutting = totals("裁床完成")[0]
    sewing = totals("车衣完成")[0]
    status = frame["纸样完成进度"].fillna("").astype(str).str.strip().str.replace(r"^\d+", "", regex=True)
    unfinished = status.isin(["进行中", "未处理"])
    today = pd.Timestamp(datetime.now().date())
    submitted = pd.to_datetime(frame["办单提交纸样部日期"], errors="coerce").dt.normalize()
    overdue = unfinished & ((today - submitted).dt.days > 10)
    _, label, weekly = report_period()

    def number(value):
        return f"{float(value or 0):g}"

    lines = [
        f"{label}新增办单：{number(changes[0][1])}个",
        f"完成纸样数量：{number(changes[1][1])}个",
        f"纸样师傅手头上未完成纸样：{int(unfinished.sum())}个（进行中、未处理），其中超过10天：{int(overdue.sum())}个（截至{today.strftime('%Y-%m-%d')}）",
        f"裁床裁剪款数：{number(cutting[1])}款；裁剪数量：{number(cutting[2])}件；计划件数：{number(cutting[3])}件",
        f"车衣款数：{number(sewing[1])}款；车衣数量：{number(sewing[2])}件；计划件数：{number(sewing[3])}件；裁剪件数：{number(sewing[4])}件",
    ]
    sheet = retry(lambda: workbook.Worksheets("透视表简报"))
    for index, text in enumerate(lines, 1):
        retry(lambda index=index, text=text: setattr(sheet.Cells(index, 1), "Value", text))
        logging.info("简报：%s", text)


DATE_FILTERS = {
    "裁床完成": ["裁衣完成并移交日期（不含时间）"],
    "办单数量变化": ["办单提交纸样部日期（不含时间）", "纸样完成并转至排唛日期（不含时间）"],
    "车衣完成": ["车办改转态日期（不含时间）"],
}


def set_yesterday_filters(workbook, retry):
    dates, label, weekly = report_period()
    for sheet_name, names in DATE_FILTERS.items():
        sheet = retry(lambda: workbook.Worksheets(sheet_name))
        pivots = retry(lambda: list(sheet.PivotTables()))
        if len(pivots) != len(names):
            raise ValueError(f"{sheet_name} 透视表数量与日期配置不符")
        pivots.sort(key=lambda pivot: retry(lambda: pivot.TableRange1.Column))
        for pivot, name in zip(pivots, names):
            field = retry(lambda: pivot.PivotFields(name))
            retry(lambda: setattr(field, "Orientation", 3))
            retry(field.ClearAllFilters)
            if weekly:
                retry(lambda: setattr(pivot, "ManualUpdate", True))
                retry(lambda: setattr(field, "EnableMultiplePageItems", True))
                items = retry(lambda: list(field.PivotItems()))
                matching = [item for item in items if retry(lambda item=item: item.Name) in dates]
                if not matching:
                    raise ValueError(f"{sheet_name} 上周没有可用日期，停止生成，避免错误统计全部数据")
                for item in matching:
                    retry(lambda item=item: setattr(item, "Visible", True))
                for item in items:
                    if retry(lambda item=item: item.Name) not in dates:
                        retry(lambda item=item: setattr(item, "Visible", False))
                # Excel 切换多选时可能重新显示空白项，按名称重新获取后再次隐藏。
                for _ in range(2):
                    for item in retry(lambda: list(field.PivotItems())):
                        item_name = retry(lambda item=item: item.Name)
                        if item_name not in dates and retry(lambda item=item: item.Visible):
                            retry(lambda item_name=item_name: setattr(field.PivotItems(item_name), "Visible", False))
                retry(lambda: setattr(pivot, "ManualUpdate", False))
                retry(pivot.Update)
                selected = {retry(lambda item=item: item.Name) for item in retry(lambda: list(field.PivotItems())) if retry(lambda item=item: item.Visible)}
                expected = {retry(lambda item=item: item.Name) for item in matching}
                if selected - {"(blank)", "(空白)"} != expected:
                    raise ValueError(f"{sheet_name} 上周日期筛选失败：{selected}")
            else:
                retry(lambda: setattr(field, "EnableMultiplePageItems", False))
                retry(lambda: setattr(field, "CurrentPage", dates[0]))
                if retry(lambda: field.CurrentPage.Name) != dates[0]:
                    raise ValueError(f"{sheet_name} 昨日筛选失败")
            logging.info("日期筛选：%s / %s = %s", sheet_name, name, label)


def process(folder, template=TEMPLATE, merge_only=False):
    rules, columns = load_rules()
    source, frame = merge_files(folder, rules, columns)
    if not merge_only and not frame.empty:
        _, _, weekly = report_period()
        report_name = " 上周周报.xlsx" if weekly else " 数据透视表.xlsx"
        report = refresh_report(source, frame, template, folder / (datetime.now().strftime("%Y-%m-%d") + report_name), rules)
        import json
        manifest = folder / "latest_report.json"
        temporary = manifest.with_suffix(".tmp")
        temporary.write_text(json.dumps({"report": str(report.resolve()), "created_at": datetime.now().isoformat()}, ensure_ascii=False), encoding="utf-8")
        temporary.replace(manifest)
    elif not merge_only:
        raise ValueError("合并结果为空，未生成本次透视表；目录中原有报表不代表本次结果")
    return source


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=ROOT / "excel-dingding")
    parser.add_argument("--template", type=Path, default=TEMPLATE)
    parser.add_argument("--merge-only", action="store_true", help="仅合并，不刷新透视表")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    process(args.input.resolve(), args.template.resolve(), args.merge_only)
