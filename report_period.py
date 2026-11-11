"""周一统计上个完整自然周，其余日期统计昨日。"""
from datetime import datetime, timedelta


def report_period(now=None):
    today = (now or datetime.now()).date()
    end = today - timedelta(days=1)
    weekly = today.weekday() == 0
    start = end - timedelta(days=6) if weekly else end
    dates = [(start + timedelta(days=index)).isoformat() for index in range((end - start).days + 1)]
    label = f"{start.month}月{start.day}日至{end.month}月{end.day}日" if weekly else f"{end.month}月{end.day}日"
    if start.year != end.year:
        label = f"{start.year}年{start.month}月{start.day}日至{end.year}年{end.month}月{end.day}日"
    return dates, label, weekly
