"""每日重新生成报表，使用 dinging-message-api 的 dws 封装发送群文件和简报。"""
import argparse
import hashlib
import json
import logging
import os
from pathlib import Path
import subprocess
import sys
import time
import uuid
from datetime import datetime, timezone, timedelta

ROOT = Path(__file__).resolve().parent
GROUP = "cidwD23obEJ75VGz/+2Fej91g=="
BEIJING = timezone(timedelta(hours=8))


def save(path, state):
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--scheduled", action="store_true", help="提前准备，等待北京时间09:00发送")
    parser.add_argument("--once-at", help="一次性发送的北京时间，格式 YYYY-MM-DDTHH:MM")
    args = parser.parse_args()
    if args.once_at:
        args.once_at = datetime.fromisoformat(args.once_at).replace(tzinfo=BEIJING)
        if args.once_at.date() != datetime.now(BEIJING).date():
            parser.error("一次性发送日期必须为今天，防止发送过期报表")
    os.chdir(ROOT)
    logs = ROOT / "send-logs"
    logs.mkdir(exist_ok=True)
    today = datetime.now(BEIJING).date().isoformat()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                        handlers=[logging.FileHandler(logs / (today + ".log"), encoding="utf-8"), logging.StreamHandler()])
    import msvcrt
    lock = (logs / "job.lock").open("a+b")
    lock.seek(0)
    if lock.read(1) == b"":
        lock.write(b"0")
        lock.flush()
    lock.seek(0)
    msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
    try:
        return run(args, logs, today)
    finally:
        lock.seek(0)
        msvcrt.locking(lock.fileno(), msvcrt.LK_UNLCK, 1)
        lock.close()


def run(args, logs, today):
    group_key = hashlib.sha256(GROUP.encode()).hexdigest()[:10]
    batch = today + (args.once_at.strftime("-once-%H%M") if args.once_at else "") + "-" + group_key
    state_path = logs / (batch + ("-preview.json" if args.dry_run else ".json"))
    state = json.loads(state_path.read_text(encoding="utf-8")) if state_path.exists() else {}
    if state.get("file_sent") and state.get("brief_sent"):
        logging.info("今天已发送完成，跳过重复任务")
        return 0
    sys.path.insert(0, str(ROOT / "dinging-message-api"))
    from app.dingtalk import send_message, run_dws
    from openpyxl import load_workbook
    auth = run_dws(["auth", "status", "--format", "json"])
    if not auth.get("authenticated") or not auth.get("token_valid"):
        raise RuntimeError("dws登录失效，请运行 dws auth login --device --no-browser --recommend")
    if not state.get("report"):
        started = datetime.now()
        subprocess.run([sys.executable, "-X", "utf8", str(ROOT / "export_mysql.py"), "--user", "yida_ro", "--process"], cwd=ROOT, check=True)
        manifest = json.loads((ROOT / "excel-dingding/latest_report.json").read_text(encoding="utf-8"))
        if datetime.fromisoformat(manifest["created_at"]) < started:
            raise RuntimeError("本次没有生成新的有效报表")
        report = Path(manifest["report"])
        from verify_report import verify
        verify(ROOT / "excel-dingding", report)
        workbook = load_workbook(report, read_only=True, data_only=True)
        try:
            lines = [row[0] for row in workbook["透视表简报"].values if row[0]]
            if len(lines) != 5 or not all(isinstance(line, str) for line in lines):
                raise RuntimeError("简报内容不完整")
        finally:
            workbook.close()
        # 固定本次附件，发送重试时不受重新导出影响。
        import shutil
        attachment = logs / (batch + " 数据透视表.xlsx")
        shutil.copy2(report, attachment)
        state = {"report": str(attachment), "brief": "\n".join(lines),
                 "sha256": hashlib.sha256(attachment.read_bytes()).hexdigest(), "group": GROUP}
        save(state_path, state)
    report = Path(state["report"])
    if hashlib.sha256(report.read_bytes()).hexdigest() != state["sha256"]:
        raise RuntimeError("待发送附件已改变，停止发送")
    if args.scheduled or args.once_at:
        now = datetime.now(BEIJING)
        target = args.once_at or now.replace(hour=9, minute=0, second=0, microsecond=0)
        if now < target:
            logging.info("报表准备完成，等待北京时间%s发送", target.strftime("%H:%M"))
            time.sleep((target - now).total_seconds())
        if datetime.now(BEIJING).date().isoformat() != today:
            raise RuntimeError("跨日任务不发送过期报表")
    for kind in ["file", "brief"]:
        if state.get(kind + "_sent"):
            continue
        key = str(uuid.uuid5(uuid.NAMESPACE_URL, f"yida-report/{batch}/{GROUP}/{kind}/{args.dry_run}"))
        result = None
        from format_brief import format_brief
        content = report.name if kind == "file" else format_brief(state["brief"].splitlines())
        for attempt in range(3):
            result = send_message(target_type="group", target=GROUP,
                msg_type="file" if kind == "file" else "markdown",
                content=content,
                file_path=str(report.relative_to(ROOT)) if kind == "file" else None,
                idempotency_key=key, dry_run=args.dry_run, timeout=600)
            if result.success and result.raw.get("success") is not False:
                break
            logging.error("%s发送尝试%s失败：%s", kind, attempt + 1, result.error or result.raw)
            if attempt < 2:
                time.sleep(10)
        else:
            raise RuntimeError(f"{kind}发送失败：{result.error}")
        state[kind + "_sent"] = True
        state[kind + "_result"] = result.raw
        save(state_path, state)
        logging.info("%s %s完成，目标：数据科技部（4人群）", kind, "预览" if args.dry_run else "发送")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        logging.exception("每日报表任务失败")
        sys.exit(1)
