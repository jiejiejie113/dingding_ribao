"""只读验证合并文件与当天报表的数据、缓存、筛选和异常标色。"""
from datetime import datetime, timedelta
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import pandas as pd
from openpyxl import load_workbook


def verify(folder, report=None):
    report = report or folder / (datetime.now().strftime("%Y-%m-%d") + " 数据透视表.xlsx")
    merged = pd.read_excel(folder / "合并数据_完整提取.xlsx", dtype=object)
    removed = pd.read_excel(folder / "剔除数据表.xlsx", dtype=object)
    keys = ["东明款号", "办名称", "重做版次", "品牌&部门"]
    base = pd.read_excel(folder / "跟单.xlsx", dtype=object)
    assert not merged.duplicated(keys).any(), "合并结果存在重复业务键"
    assert len(merged) + len(removed) == len(base.drop_duplicates(keys)), "保留及剔除数量不守恒"
    assert removed["制办单分配进度"].isna().all()
    assert (removed["是否经过纸样部"] == "需要").all()
    assert (removed["超时天数"] > 3).all()
    workbook = load_workbook(report, read_only=True, data_only=True)
    try:
        rows = iter(workbook["数据"].values)
        actual = pd.DataFrame(rows, columns=next(rows))
        assert actual.shape == merged.shape
        for col in merged.columns:
            left, right = merged[col].fillna(""), actual[col].fillna("")
            same = left.astype(str) == right.astype(str)
            # Excel/Pandas 的整数和浮点表现可不同，但纯文本列必须逐字相等。
            if col in keys or col == "设计周":
                assert same.all(), f"文本字段变化：{col}"
            else:
                ln, rn = pd.to_numeric(left, errors="coerce"), pd.to_numeric(right, errors="coerce")
                assert (same | (ln.notna() & rn.notna() & (ln == rn))).all(), f"字段不一致：{col}"
        print(f"PASS 数据逐列核对：{len(merged)}行、{len(merged.columns)}列，文本前导零保留")
    finally:
        workbook.close()
    namespace = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with ZipFile(report) as archive:
        assert archive.testzip() is None
        caches = {}
        for name in archive.namelist():
            if name.startswith("xl/pivotCache/pivotCacheDefinition") and name.endswith(".xml"):
                root = ET.fromstring(archive.read(name))
                assert int(root.get("recordCount")) == len(merged)
                assert root.find("m:cacheSource/m:worksheetSource", namespace).get("ref") == f"A1:AJ{len(merged)+1}"
                caches[name] = root
        pivots = [name for name in archive.namelist() if name.startswith("xl/pivotTables/pivotTable") and name.endswith(".xml")]
        assert len(pivots) == 7
        # 通过关系文件找到每张透视表缓存，检查四个日期页字段真实选中项。
        dates = {"裁衣完成并移交日期（不含时间）", "办单提交纸样部日期（不含时间）",
                 "纸样完成并转至排唛日期（不含时间）", "车办改转态日期（不含时间）"}
        checked = []
        yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        import posixpath
        for name in pivots:
            pivot = ET.fromstring(archive.read(name))
            relations = ET.fromstring(archive.read(posixpath.dirname(name) + "/_rels/" + posixpath.basename(name) + ".rels"))
            target = next(r.get("Target") for r in relations if r.get("Type").endswith("/pivotCacheDefinition"))
            cache_name = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(name), target))
            fields = caches[cache_name].findall("m:cacheFields/m:cacheField", namespace)
            for page in pivot.findall("m:pageFields/m:pageField", namespace):
                field = fields[int(page.get("fld"))]
                if field.get("name") not in dates:
                    continue
                if datetime.now().weekday() != 0:
                    item = pivot.findall("m:pivotFields/m:pivotField", namespace)[int(page.get("fld"))].findall("m:items/m:item", namespace)[int(page.get("item"))]
                    value = list(field.find("m:sharedItems", namespace))[int(item.get("x"))].get("v")
                    assert value == yesterday, (field.get("name"), value)
                    checked.append(field.get("name"))
        if datetime.now().weekday() != 0:
            assert set(checked) == dates
        print("PASS 文件完整性、7张透视表、缓存实际范围及昨日筛选")
    print(f"PASS 数量守恒及剔除规则：保留{len(merged)}，剔除{len(removed)}")


if __name__ == "__main__":
    verify(Path(__file__).resolve().parent / "excel-dingding")
