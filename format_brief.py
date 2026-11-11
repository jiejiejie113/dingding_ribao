"""将报表简报格式化为钉钉 Markdown 生产日报。"""
import re


def format_brief(lines):
    if len(lines) != 5:
        raise ValueError("生产日报需要五行完整的原始简报")

    def extract(pattern, text):
        match = re.search(pattern, text)
        if not match:
            raise ValueError(f"无法解析简报：{text}")
        return match.groups()

    label, new = extract(r"^(.+?)新增办单：([\d.]+)个", lines[0])
    completed, = extract(r"完成纸样数量：([\d.]+)个", lines[1])
    pending, overdue, year, month, day = extract(
        r"未完成纸样：([\d.]+)个.*超过10天：([\d.]+)个（截至(\d{4})-(\d{2})-(\d{2})）", lines[2])
    cutting, cut_qty, cut_plan = extract(
        r"裁床裁剪款数：([\d.]+)款；裁剪数量：([\d.]+)件；计划件数：([\d.]+)件", lines[3])
    sewing, sewn_qty, sew_plan, sewn_cut = extract(
        r"车衣款数：([\d.]+)款；车衣数量：([\d.]+)件；计划件数：([\d.]+)件；裁剪件数：([\d.]+)件", lines[4])
    label = re.sub(r"0(\d)日", r"\1日", label)
    return (
        f"## {label}生产日报\n\n"
        f"### 办单与纸样\n\n"
        f"- 新增办单：{new}个\n"
        f"- 完成纸样：{completed}个\n"
        f"- 待完成纸样：{pending}个（含进行中、未处理）\n"
        f"  - 其中，超过10天未完成：{overdue}个（截至{year}年{int(month)}月{int(day)}日）\n\n"
        f"### 裁床裁剪\n\n"
        f"- 裁剪款数：{cutting}款\n"
        f"- 裁剪数量：{cut_qty}件\n"
        f"- 计划件数：{cut_plan}件\n\n"
        f"### 车衣生产\n\n"
        f"- 车衣款数：{sewing}款\n"
        f"- 车衣数量：{sewn_qty}件\n"
        f"- 计划件数：{sew_plan}件\n"
        f"- 裁剪件数：{sewn_cut}件"
    )
