import sys
try:
    if sys.stdout and sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

import pandas as pd
import os
import math
import re
import shutil
import subprocess
import time
import ctypes
from ctypes import wintypes
import warnings
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import win32com.client
except ImportError:
    win32com = None
    print("⚠️  未安装 pywin32 库，透视表刷新功能将不可用。请运行：pip install pywin32")

try:
    import xlwings as xw
except ImportError:
    xw = None
    print("⚠️  未安装 xlwings 库，部分功能可能不可用。请运行：pip install xlwings")

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ================= Excel COM 忙碌自动重试 =================
VBA_E_IGNORE = -2146777998  # 0x800AC472：Excel 忙时拒绝 OLE 调用
BUSY_HRESULTS = {
    -2146777998,  # 0x800AC472 VBA_E_IGNORE
    -2147418111,  # 0x80010001 RPC_E_CALL_REJECTED
    -2147417846,  # 0x8001010A RPC_E_SERVERCALL_RETRYLATER
}
CURRENT_EXCEL_PID = None  # 当前自动化 Excel 进程 PID（用于自动关闭阻塞弹窗）

def close_blocking_dialogs(pid):
    """关闭指定 Excel 进程里阻塞自动化的模态弹窗（如「Microsoft Office 激活向导」）"""
    if not pid:
        return []
    closed = []
    try:
        user32 = ctypes.windll.user32
        enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        def _cb(hwnd, lparam):
            try:
                p = wintypes.DWORD()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(p))
                if p.value != pid:
                    return True
                title = ctypes.create_unicode_buffer(512)
                user32.GetWindowTextW(hwnd, title, 512)
                t = title.value
                if ('激活向导' in t) or ('Activation Wizard' in t) \
                        or ('正在等待另一个应用程序' in t) or ('waiting for another application' in t):
                    user32.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE
                    closed.append(t or hex(hwnd))
            except Exception:
                pass
            return True
        user32.EnumWindows(enum_proc(_cb), 0)
    except Exception:
        pass
    return closed

def is_excel_busy_error(e):
    try:
        if isinstance(e, AttributeError):
            # pywin32 动态分派会把 Excel 忙碌的 COM 错误吞掉，伪装成 "<unknown>.属性名" / "Item.属性名"
            return bool(re.fullmatch(r'[^.\s]+\.\w+', str(e)))
        if not e.args:
            return False
        if e.args[0] in BUSY_HRESULTS:
            return True
        # 有些调用会把忙碌错误包在 DISP_E_EXCEPTION 的 excepinfo 里（如 Application.Quit）
        excepinfo = e.args[2] if len(e.args) > 2 else None
        return isinstance(excepinfo, (tuple, list)) and bool(excepinfo) and excepinfo[-1] in BUSY_HRESULTS
    except Exception:
        return False

def com_retry(func, *args, retries=15, delay=1.0, label=None, **kwargs):
    """执行 Excel COM 调用，遇到「Excel 忙」错误时自动等待重试"""
    name = label or getattr(func, '__name__', 'Excel 操作')
    for attempt in range(1, retries + 1):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            if not is_excel_busy_error(e) or attempt == retries:
                raise
            print(f"  ⏳ [{name}] Excel 忙，等待 {delay:g} 秒后重试（第 {attempt} 次）...")
            closed = close_blocking_dialogs(CURRENT_EXCEL_PID)
            if closed:
                print(f"  🧹 已自动关闭阻塞弹窗：{closed}（建议尽快激活 Office，否则每次启动都会弹出）")
            time.sleep(delay)

def retry_on_excel_busy(retries=5, delay=1.0, label=None):
    """函数级重试装饰器：整个函数因 Excel 忙失败时自动重跑"""
    def decorator(func):
        def wrapper(*args, **kwargs):
            return com_retry(func, *args, retries=retries, delay=delay,
                             label=label or func.__name__, **kwargs)
        return wrapper
    return decorator

def wait_excel_idle(excel, timeout=120, interval=0.5):
    """等待 Excel 完成启动/后台计算刷新；遇到「忙」继续等，不能当作空闲"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            ready = excel.Ready
            calc = excel.CalculationState
        except Exception as e:
            if is_excel_busy_error(e):
                time.sleep(interval)
                continue
            return True  # 属性不存在等非忙异常，视为空闲
        if ready and calc == 0:
            return True
        time.sleep(interval)
    return False

def get_worksheets(workbook):
    """安全获取工作表列表：逐项 Item(i) 获取，避免枚举器在 Excel 忙时返回残缺对象"""
    count = com_retry(lambda: workbook.Worksheets.Count, label='获取工作表数')
    sheets = []
    for i in range(1, count + 1):
        sheets.append(com_retry(lambda i=i: workbook.Worksheets.Item(i), label='获取工作表'))
    return sheets

def list_excel_pids():
    """列出当前所有 EXCEL.EXE 进程 PID"""
    try:
        out = subprocess.check_output(['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE', '/FO', 'CSV', '/NH'],
                                      encoding='gbk', errors='replace')
        return {int(p) for p in re.findall(r'"EXCEL\.EXE","(\d+)"', out, re.IGNORECASE)}
    except Exception:
        return set()

def kill_process(pid):
    """强制结束指定 PID 的进程（仅用于清理我们自己的 Excel 实例）"""
    try:
        subprocess.run(['taskkill', '/F', '/PID', str(pid)],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception:
        return False

def find_hidden_excel_pids():
    """找出没有可见窗口的后台 Excel 进程（残留的自动化实例）"""
    try:
        out = subprocess.check_output(['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE', '/V', '/FO', 'CSV', '/NH'],
                                      encoding='gbk', errors='replace')
    except Exception:
        return []
    hidden = []
    for line in out.splitlines():
        fields = re.findall(r'"([^"]*)"', line)
        if len(fields) < 2 or fields[0].upper() != 'EXCEL.EXE':
            continue
        try:
            pid = int(fields[1])
        except ValueError:
            continue
        title = fields[-1] if fields else ''
        if (not title) or ('HardwareMonitorWindow' in title) or title.strip() in ('N/A', '无'):
            hidden.append(pid)
    return hidden

def wait_excel_ready(excel, timeout=180, interval=1.0, show_after=30):
    """等待 Excel 实例真正就绪；超过 show_after 秒仍不响应就把窗口显示出来便于排查"""
    start = time.time()
    shown = False
    tick = 0
    while time.time() - start < timeout:
        tick += 1
        try:
            com_retry(lambda: excel.Workbooks.Count, retries=1, label='就绪探测')
            print(f"✅ Excel 已就绪（等待 {time.time() - start:.0f} 秒）")
            return True
        except Exception as e:
            if not is_excel_busy_error(e):
                raise
        closed = close_blocking_dialogs(CURRENT_EXCEL_PID)
        if closed:
            print(f"  🧹 已自动关闭阻塞弹窗：{closed}")
        waited = time.time() - start
        if not shown and waited >= show_after:
            print("⚠️  Excel 启动很慢，已显示 Excel 窗口；若有弹窗（登录/加载项/文档恢复），请点掉它...")
            try:
                com_retry(lambda: setattr(excel, 'Visible', True), retries=3, label='显示 Excel')
            except Exception:
                pass
            shown = True
        if tick % 5 == 0:
            print(f"  ⏳ 等待 Excel 就绪...（已等 {waited:.0f}/{timeout} 秒）")
        time.sleep(interval)
    return False

def cleanup_excel(excel, wb=None, our_pid=None):
    """退出 Excel：先关工作簿再 Quit；若仍残留则强制结束我们自己的进程，避免僵尸堆积"""
    if wb is not None:
        try:
            com_retry(lambda: wb.Close(False), retries=5, delay=1.0, label='关闭工作簿')
        except Exception:
            pass
    try:
        com_retry(excel.Quit, retries=10, delay=1.0, label='退出 Excel')
    except Exception:
        pass
    if our_pid:
        # 给 Excel 最多 15 秒自行退出，避免误杀正在关闭的进程
        for _ in range(15):
            if our_pid not in list_excel_pids():
                break
            time.sleep(1)
        if our_pid in list_excel_pids():
            print(f"⚠️  Excel 未能正常退出，强制结束进程 PID {our_pid}")
            kill_process(our_pid)
    global CURRENT_EXCEL_PID
    CURRENT_EXCEL_PID = None

# ================= 配置区域 =================
DEFAULT_TEMPLATE_PATH = r"C:\Users\caohanjie\Desktop\合并转透视表\2026-10-16 模板.xlsx"
DEFAULT_OUTPUT_DIR = r"C:\Users\caohanjie\Desktop\合并转透视表"

FILE_TYPE_KEYWORDS = {
    'follow_up': ['跟单', 'follow', 'tracking', 'merchandiser'],
    'allocation': ['分配', 'allocation', 'assign', '派工'],
    'pattern_status': ['状态', 'status', '纸样', 'pattern'],
    'cutting': ['裁床', 'cutting', '裁剪', '裁断'],
    'sewing': ['车办', '车衣', 'sewing', '车缝']
}

NAME_MAPPING = {
    'wei rongling': '韦荣领',
    'weirongling': '韦荣领',
}

# ================= 通用辅助函数 =================
def select_folder(title="选择文件夹"):
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory(title=title); root.destroy()
    return folder_path

def select_multiple_files(title, filetypes=None):
    if filetypes is None: filetypes = [("Excel 文件", "*.xlsx *.xls")]
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    file_paths = filedialog.askopenfilenames(title=title, filetypes=filetypes); root.destroy()
    return list(file_paths)

def select_file(title, filetypes=None):
    if filetypes is None: filetypes = [("Excel 文件", "*.xlsx *.xls")]
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes); root.destroy()
    return file_path if file_path else None

def identify_file_type(file_path):
    file_name = os.path.basename(file_path).lower()
    for file_type, keywords in FILE_TYPE_KEYWORDS.items():
        for keyword in keywords:
            if keyword.lower() in file_name: return file_type
    return None

def get_available_filename(base_path):
    if not os.path.exists(base_path): return base_path
    try:
        os.remove(base_path); return base_path
    except PermissionError:
        timestamp = datetime.now().strftime('%H%M%S')
        name, ext = os.path.splitext(base_path)
        return f"{name}_{timestamp}{ext}"

def wait_for_file_closed(filepath, max_wait=10):
    for i in range(max_wait):
        try:
            with open(filepath, 'r+b') as f: pass
            return True
        except PermissionError:
            print(f"⏳ 文件被占用，等待 {i+1}/{max_wait} 秒...")
            time.sleep(1)
    return False

def normalize_name(name):
    if pd.isna(name) or not isinstance(name, str):
        return name
    name = name.strip()
    lower_name = name.lower()
    if lower_name in NAME_MAPPING:
        return NAME_MAPPING[lower_name]
    import re
    m = re.search(r'dm[-—–]\s*([\u4e00-\u9fff]+)', name, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r'([\u4e00-\u9fff]+)', name)
    if m:
        return m.group(1)
    return name

def normalize_staff_column(df):
    if '跟办人员' in df.columns:
        df['跟办人员'] = df['跟办人员'].apply(normalize_name)
    if '分配师傅' in df.columns:
        df['分配师傅'] = df['分配师傅'].apply(normalize_name)
    if '纸样师傅' in df.columns:
        df['纸样师傅'] = df['纸样师傅'].apply(normalize_name)
    return df

# ================= 第一阶段：数据合并逻辑 =================
def read_and_extract_data(file_path, required_columns, file_name):
    if not file_path or not os.path.exists(file_path):
        print(f"⚠️  {file_name} 文件不存在或未选择"); return pd.DataFrame()
    try:
        df = pd.read_excel(file_path)
        if df.empty: print(f"⚠️  {file_name} 文件为空"); return pd.DataFrame()
        
        df.columns = df.columns.str.strip()
        column_mapping = {
            '品牌/部门': '品牌&部门', '计划件数': '计划制衣件数', '计划裁剪件数': '计划制衣件数',
            '纸样完成状态': '纸样完成进度', '裁床完成状态': '裁床完成进度', '裁剪数量': '裁剪总数',
            '修改时间': '车办改转态日期'
        }
        df = df.rename(columns=column_mapping)
        
        available_columns = [col for col in required_columns if col in df.columns]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns: print(f"⚠️  {file_name} 缺少列：{missing_columns}")
        
        if available_columns:
            df = df[available_columns].copy()
            print(f"✅ {file_name} 成功读取 {len(df)} 行数据，{len(available_columns)} 列")
        else:
            df = pd.DataFrame(); print(f"❌ {file_name} 无法提取所需数据")
        return df
    except Exception as e:
        print(f"❌ 读取 {file_name} 失败：{str(e)}"); return pd.DataFrame()

def merge_dataframes_safe(df_base, df_add, merge_keys, how='left'):
    if df_base.empty and df_add.empty: return pd.DataFrame()
    if df_base.empty: return df_add.copy()
    if df_add.empty: return df_base.copy()
    
    available_keys = [key for key in merge_keys if key in df_base.columns and key in df_add.columns]
    if not available_keys:
        print(f"⚠️  缺少共同合并键，可用键：{available_keys}"); return df_base
    try:
        common_cols = set(df_base.columns) & set(df_add.columns)
        duplicate_cols = common_cols - set(available_keys)
        merged = pd.merge(df_base, df_add, on=available_keys, how=how, suffixes=('_left', '_right'))
        for col in duplicate_cols:
            left_col, right_col = f'{col}_left', f'{col}_right'
            if left_col in merged.columns and right_col in merged.columns:
                merged[col] = merged[left_col].combine_first(merged[right_col])
                merged.drop(columns=[left_col, right_col], inplace=True)
        print(f"✅ 合并后数据量：{len(merged)} 行 ({how}连接)"); return merged
    except Exception as e:
        print(f"❌ 合并失败：{str(e)}"); return df_base

def calculate_duration_ceiling(df, end_col, start_col, new_col_name):
    if end_col not in df.columns or start_col not in df.columns:
        df[new_col_name] = None; return df
    df[end_col] = pd.to_datetime(df[end_col], errors='coerce')
    df[start_col] = pd.to_datetime(df[start_col], errors='coerce')
    duration_raw = (df[end_col] - df[start_col]).dt.total_seconds() / (24 * 3600)
    def ceiling_func(x):
        if pd.isna(x): return None
        if x <= 0: return 0
        return math.ceil(x)
    df[new_col_name] = duration_raw.apply(ceiling_func); return df

def add_duration_columns(df):
    if df.empty: return df
    today = pd.Timestamp(datetime.now().date())
    df = calculate_duration_ceiling(df, '纸样完成并转至排唛日期', '办单提交纸样部日期', '纸样处理时长 (天)')
    df = calculate_duration_ceiling(df, '裁衣完成并移交日期', '纸样完成并转至排唛日期', '裁床处理时长 (天)')
    
    df['纸样完成并转至排唛日期'] = pd.to_datetime(df['纸样完成并转至排唛日期'], errors='coerce')
    duration_cutting_current = (today - df['纸样完成并转至排唛日期']).dt.total_seconds() / (24 * 3600)
    df['裁床耗时 (天) - 当前'] = duration_cutting_current.apply(lambda x: math.ceil(x) if pd.notna(x) and x > 0 else (0 if pd.notna(x) else None))
    
    df['办单提交纸样部日期'] = pd.to_datetime(df['办单提交纸样部日期'], errors='coerce')
    duration_total = (today - df['办单提交纸样部日期']).dt.total_seconds() / (24 * 3600)
    df['总耗时 (天)'] = duration_total.apply(lambda x: math.ceil(x) if pd.notna(x) and x > 0 else (0 if pd.notna(x) else None))
    return df

def add_date_only_columns(df):
    if df.empty: return df
    if '办单提交纸样部日期' in df.columns:
        temp_dt = pd.to_datetime(df['办单提交纸样部日期'], errors='coerce')
        df['办单提交纸样部日期（不含时间）'] = temp_dt.dt.strftime('%Y-%m-%d')
        df['办单提交纸样部日期（年月）'] = temp_dt.dt.strftime('%Y-%m')
        del temp_dt
    if '纸样完成并转至排唛日期' in df.columns:
        temp_dt = pd.to_datetime(df['纸样完成并转至排唛日期'], errors='coerce')
        df['纸样完成并转至排唛日期（不含时间）'] = temp_dt.dt.strftime('%Y-%m-%d')
        df['纸样完成并转至排唛日期（年月）'] = temp_dt.dt.strftime('%Y-%m')
        del temp_dt
    if '裁衣完成并移交日期' in df.columns:
        temp_dt = pd.to_datetime(df['裁衣完成并移交日期'], errors='coerce')
        df['裁衣完成并移交日期（不含时间）'] = temp_dt.dt.strftime('%Y-%m-%d')
        df['裁衣完成并移交日期（年月）'] = temp_dt.dt.strftime('%Y-%m')
        del temp_dt
        
    # ✅ AB列严格保持原格式，仅提取纯日期放入AC列，并生成只含年月的列
    if '车办改转态日期' in df.columns:
        temp_dt = pd.to_datetime(df['车办改转态日期'], errors='coerce')
        df['车办改转态日期（不含时间）'] = temp_dt.dt.strftime('%Y-%m-%d')
        df['车办改转态日期（年月）'] = temp_dt.dt.strftime('%Y-%m')
        del temp_dt
    print("✅ 不含时间日期列已添加（AB列原格式保留，AC列生成纯日期，表尾新增各日期的年月列）")
    return df

def create_empty_template(output_path):
    columns = [
        '东明款号', '办名称', '重做版次', '品牌&部门', '跟办人员', '是否需要做办', '产品类型',
        '是否经过纸样部', '款式类型', '设计周',
        '办单提交纸样部日期', '办单提交纸样部日期（不含时间）', '计划制衣件数',
        '制办单分配进度', '办单分配日期', '分配师傅', '纸样师傅', '纸样完成进度', 
        '开始处理纸样日期', '纸样完成并转至排唛日期', '纸样完成并转至排唛日期（不含时间）',
        '裁床完成进度', '开始裁衣日期', '裁衣完成并移交日期', '裁衣完成并移交日期（不含时间）',
        '裁剪总数', '车衣进度', '开始车衣时间', '组别', '车办实际总件数',
        '车办改转态日期', '车办改转态日期（不含时间）',
        '纸样处理时长 (天)', '裁床处理时长 (天)', '裁床耗时 (天) - 当前', '总耗时 (天)',
        '办单提交纸样部日期（年月）', '纸样完成并转至排唛日期（年月）',
        '裁衣完成并移交日期（年月）', '车办改转态日期（年月）'
    ]
    df_empty = pd.DataFrame(columns=columns)
    df_empty.to_excel(output_path, index=False)
    print(f"✅ 空模板已创建：{output_path}")
    return columns

# ================= 第二阶段：透视表刷新逻辑 =================
@retry_on_excel_busy(retries=30, delay=1.0)
def _set_pivot_page_filters(workbook, yesterday):
    checks = [
        ("裁床完成", "裁衣完成并移交日期（不含时间）"),
        ("办单数量变化", "办单提交纸样部日期（不含时间）"),
        ("办单数量变化", "纸样完成并转至排唛日期（不含时间）"),
        ("车衣完成", "车办改转态日期（不含时间）"),
    ]
    success_count = 0
    for sheet_name, field_name in checks:
        try:
            ws = workbook.Worksheets(sheet_name)
            field_found = False
            for pt in ws.PivotTables():
                for i in range(1, pt.PivotFields().Count + 1):
                    pf = pt.PivotFields().Item(i)
                    if round(pf.Orientation) == 3 and pf.Name == field_name:
                        field_found = True
                        items = pf.PivotItems()
                        target_item = None
                        for j in range(1, items.Count + 1):
                            item = items.Item(j)
                            if item.Name == yesterday:
                                target_item = item
                                break
                        if target_item is None:
                            print(f"  ⚠️ {sheet_name} / {field_name}: 透视表缓存中无 [{yesterday}] 的数据")
                            continue
                        if not target_item.Visible:
                            target_item.Visible = True
                        pf.EnableMultiplePageItems = False
                        pf.CurrentPage = yesterday
                        print(f"  ✅ 已筛选 {sheet_name} → {field_name} = [{yesterday}]")
                        success_count += 1
            if not field_found:
                print(f"  ⚠️ {sheet_name}: 未找到页字段 [{field_name}]")
        except Exception as e:
            if is_excel_busy_error(e):
                raise
            print(f"  ⚠️ {sheet_name} / {field_name} 筛选失败: {str(e)[:60]}")
    return success_count > 0

@retry_on_excel_busy(retries=30, delay=1.0)
def _set_pivot_page_filters_weekend(workbook, saturday, sunday):
    checks = [
        ("裁床完成", "裁衣完成并移交日期（不含时间）"),
        ("办单数量变化", "办单提交纸样部日期（不含时间）"),
        ("办单数量变化", "纸样完成并转至排唛日期（不含时间）"),
        ("车衣完成", "车办改转态日期（不含时间）"),
    ]
    target = {saturday, sunday}
    for sheet_name, field_name in checks:
        try:
            ws = workbook.Worksheets(sheet_name)
            field_found = False
            for pt in ws.PivotTables():
                for i in range(1, pt.PivotFields().Count + 1):
                    pf = pt.PivotFields().Item(i)
                    if round(pf.Orientation) == 3 and pf.Name == field_name:
                        field_found = True
                        all_names = [pf.PivotItems().Item(j).Name for j in range(1, pf.PivotItems().Count + 1)]
                        missing = target - set(all_names)
                        if missing:
                            print(f"  ⚠️ {sheet_name} / {field_name}: 透视表缓存中无 {sorted(missing)} 的数据")
                        pf.EnableMultiplePageItems = True
                        pf.ClearAllFilters()
                        for j in range(1, pf.PivotItems().Count + 1):
                            item = pf.PivotItems().Item(j)
                            item.Visible = (item.Name in target)
                        visible = [pf.PivotItems().Item(j).Name for j in range(1, pf.PivotItems().Count + 1) if pf.PivotItems().Item(j).Visible]
                        print(f"  ✅ 已筛选 {sheet_name} → {field_name} = {visible}")
            if not field_found:
                print(f"  ⚠️ {sheet_name}: 未找到页字段 [{field_name}]")
        except Exception as e:
            if is_excel_busy_error(e):
                raise
            print(f"  ⚠️ {sheet_name} / {field_name} 筛选失败: {str(e)[:60]}")
    return True

@retry_on_excel_busy(retries=3, delay=5.0)
def refresh_pivot_with_excel(source_path, template_path, output_path):
    if win32com is None:
        print("❌ 无法刷新透视表：缺少 pywin32 库"); return False

    source_path = os.path.abspath(os.path.normpath(source_path))
    template_path = os.path.abspath(os.path.normpath(template_path))
    output_path = os.path.abspath(os.path.normpath(output_path))

    print("🚀 正在启动 Excel 刷新数据透视表...")
    if os.path.exists(output_path):
        print(f"⚠️  输出文件已存在：{output_path}")
        if not wait_for_file_closed(output_path, max_wait=5):
            output_path = get_available_filename(output_path)
    
    try:
        shutil.copy2(template_path, output_path)
        print(f"✅ 模板已复制：{os.path.basename(output_path)}")
    except PermissionError as e:
        print(f"❌ 无法复制模板文件，请关闭所有 Excel 文件后重试")
        output_path = get_available_filename(output_path)
        shutil.copy2(template_path, output_path)
    
    pids_before = list_excel_pids()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
    except Exception as e:
        print(f"❌ 无法启动 Excel 进程：{str(e)}"); return False
    our_excel_pid = next(iter(list_excel_pids() - pids_before), None)
    if our_excel_pid:
        print(f"ℹ️  本次自动化 Excel 进程 PID：{our_excel_pid}")
    global CURRENT_EXCEL_PID
    CURRENT_EXCEL_PID = our_excel_pid

    wb = None
    try:
        com_retry(lambda: setattr(excel, 'Visible', False), retries=10, delay=1.0, label='隐藏 Excel')
        com_retry(lambda: setattr(excel, 'DisplayAlerts', False), retries=10, delay=1.0, label='关闭提示')
        # 等待 Excel 实例真正就绪（加载项多时启动会很慢，期间所有调用都会被拒绝）
        if not wait_excel_ready(excel, timeout=180, show_after=30):
            raise Exception("Excel 启动后长时间无响应（可能被加载项或弹窗卡住）；请查看 Excel 窗口或禁用加载项后重试")

        wb = com_retry(excel.Workbooks.Open, output_path, retries=90, delay=2.0, label='打开工作簿')
        if wb is None: raise Exception("Excel Workbooks.Open 返回 None")

        # 打开后先探测一次，等 Excel 完成加载（大文件/pivot缓存加载期间会拒绝外部调用）
        com_retry(lambda: wb.Worksheets.Count, retries=60, delay=1.0, label='等待工作簿加载')
        wait_excel_idle(excel, timeout=60)

        sheet_names = [com_retry(lambda ws=ws: ws.Name, label='读取工作表名') for ws in get_worksheets(wb)]
        if '数据' not in sheet_names:
            new_ws = com_retry(lambda: wb.Worksheets.Add(wb.Worksheets.Item(1)), label='新建数据表')
            com_retry(lambda: setattr(new_ws, 'Name', '数据'), label='命名数据表')
        ws_data = com_retry(lambda: wb.Worksheets('数据'), label='获取数据表')

        df = pd.read_excel(source_path)

        # ✅ 智能格式化日期：跳过"不含时间"/"年月"列，且保留AB列(车办改转态日期)原格式
        skip_format_keywords = ['不含时间', '车办改转态日期', '（年月）']
        for col in df.columns:
            if any(kw in col for kw in skip_format_keywords): continue
            if '日期' in col or '时间' in col:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')
            elif df[col].dtype == 'datetime64[ns]':
                df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')

        max_row = max(2, len(df) + 2)
        # 🧹 按 UsedRange 全量清空「数据」表旧数据，避免上次运行残留行混入透视表
        try:
            used = com_retry(lambda: ws_data.UsedRange, label='获取已用区域')
            used_last_row = int(com_retry(lambda: used.Row, label='已用区域行')) + \
                            int(com_retry(lambda: used.Rows.Count, label='已用区域行数')) - 1
            used_last_col = int(com_retry(lambda: used.Column, label='已用区域列')) + \
                            int(com_retry(lambda: used.Columns.Count, label='已用区域列数')) - 1
        except Exception:
            used_last_row, used_last_col = 1, 1
        clear_last_row = max(max_row, used_last_row)
        clear_last_col = max(len(df.columns), used_last_col)
        print(f"🧹 清空「数据」表旧数据（第2~{clear_last_row}行，{clear_last_col}列）")
        com_retry(lambda: ws_data.Range(ws_data.Cells(2, 1), ws_data.Cells(clear_last_row, clear_last_col)).Clear(), label='清空旧数据')
        for col_idx, col_name in enumerate(df.columns, 1):
            com_retry(lambda: setattr(ws_data.Cells(1, col_idx), 'Value', str(col_name)))

        print(f"📝 正在写入 {len(df)} 条数据到模板...")
        col_names = list(df.columns)
        date_only_keywords = ['不含时间']
        original_format_keywords = ['车办改转态日期']
        data_array = []
        for _, row in df.iterrows():
            row_values = []
            for ci, value in enumerate(row):
                col_name = col_names[ci]
                if pd.isna(value):
                    row_values.append("")
                elif isinstance(value, (pd.Timestamp, datetime)):
                    if any(kw in col_name for kw in date_only_keywords):
                        row_values.append(value.strftime('%Y-%m-%d'))
                    elif any(kw in col_name for kw in original_format_keywords):
                        if value.hour == 0 and value.minute == 0 and value.second == 0:
                            row_values.append(value.strftime('%Y-%m-%d'))
                        else:
                            row_values.append(value.strftime('%Y-%m-%d %H:%M:%S'))
                    else:
                        row_values.append(value.strftime('%Y-%m-%d %H:%M:%S'))
                else:
                    row_values.append(value)
            data_array.append(row_values)

        if data_array:
            num_rows = len(data_array)
            num_cols = len(df.columns)
            # 将日期列设为文本格式，防止 Excel 自动转换日期字符串导致透视表格式不匹配
            for ci, col_name in enumerate(df.columns, 1):
                if '日期' in col_name or '时间' in col_name:
                    com_retry(lambda: setattr(ws_data.Columns(ci), 'NumberFormat', '@'), label='设置日期列格式')
            dest_range = com_retry(lambda: ws_data.Range(
                ws_data.Cells(2, 1),
                ws_data.Cells(1 + num_rows, num_cols)
            ), label='定位写入区域')
            com_retry(lambda: setattr(dest_range, 'Value', data_array), retries=30, label='写入数据')
        
        print(f"✅ 数据写入完成")

        yesterday = get_yesterday_date()

        # 清空旧缓存 + 刷新透视表（年月列放在表格末尾，前36列位置不变，透视缓存无需改动）
        pivot_caches = com_retry(lambda: wb.PivotCaches(), label='获取透视缓存')
        for pc_idx in range(1, com_retry(lambda: pivot_caches.Count, label='缓存数量') + 1):
            try:
                com_retry(lambda pc_idx=pc_idx: setattr(pivot_caches.Item(pc_idx), 'MissingItemsLimit', 0), label='清空缓存')
            except:
                pass
        for ws in get_worksheets(wb):
            if com_retry(lambda ws=ws: ws.Name, label='读取工作表名') != '数据':
                for pt in com_retry(lambda ws=ws: list(ws.PivotTables()), label='获取透视表'):
                    try:
                        com_retry(pt.RefreshTable, retries=30, label='刷新透视表'); print(f"✅ 已刷新透视表：{ws.Name} - {pt.Name}")
                    except Exception as e:
                        print(f"⚠️ 刷新失败 {ws.Name} - {pt.Name}: {str(e)}")

        # 设置透视表筛选器
        from datetime import timedelta as _td
        _today = datetime.now()
        if _today.weekday() == 0:
            _sun = (_today - _td(days=1)).strftime('%Y-%m-%d')
            _sat = (_today - _td(days=2)).strftime('%Y-%m-%d')
            print(f"  📅 周一运行：透视表筛选器设为周六({_sat})+周日({_sun})")
            _set_pivot_page_filters_weekend(wb, _sat, _sun)
        else:
            print(f"📅 设置透视表筛选器...")
            _set_pivot_page_filters(wb, yesterday)

        # 验证筛选结果
        print(f"📋 筛选器验证结果：")
        verify_checks = [
            ('裁床完成', '裁衣完成并移交日期（不含时间）'),
            ('办单数量变化', '办单提交纸样部日期（不含时间）'),
            ('办单数量变化', '纸样完成并转至排唛日期（不含时间）'),
            ('车衣完成', '车办改转态日期（不含时间）'),
        ]
        for sheet_name, field_name in verify_checks:
            try:
                ws = com_retry(lambda: wb.Worksheets(sheet_name), label='验证取表')
                field_found = False
                for pt in com_retry(lambda: list(ws.PivotTables()), label='验证取透视表'):
                    fields_obj = pt.PivotFields()
                    for i in range(1, fields_obj.Count + 1):
                        pf = fields_obj.Item(i)
                        if round(pf.Orientation) == 3 and pf.Name == field_name:
                            field_found = True
                            if pf.CurrentPage is not None:
                                cp_name = pf.CurrentPage.Name
                                status = "✅" if cp_name == yesterday else "⚠️"
                                print(f"  {status} {sheet_name} → {field_name} = [{cp_name}]")
                            else:
                                print(f"  ⚠️ {sheet_name} → {field_name} 当前页为空（可能启用了多选页）")
                if not field_found:
                    print(f"  ⚠️ {sheet_name}: 未找到页字段 [{field_name}]")
            except Exception as e:
                print(f"  ⚠️ {sheet_name} 验证失败: {str(e)[:40]}")

        # 等待 Excel 后台刷新/计算完成后再保存，避免「Excel 忙」导致保存失败
        wait_excel_idle(excel, timeout=120)
        com_retry(wb.RefreshAll, retries=60, delay=1.0, label='刷新全部')
        wait_excel_idle(excel, timeout=180)
        com_retry(wb.Save, retries=60, delay=1.0, label='保存')
        com_retry(wb.Close, retries=60, delay=1.0, label='关闭工作簿')
        wb = None
        print(f"✅ 文件已保存：{output_path}"); return True
    except Exception as e:
        print(f"❌ 错误：{str(e)}")
        raise
    finally:
        cleanup_excel(excel, wb, our_excel_pid)

# ================= 文件自动识别与分类 =================
def scan_folder_for_files(folder_path):
    files_by_type = {k: [] for k in FILE_TYPE_KEYWORDS.keys()}
    if not os.path.exists(folder_path): return files_by_type
    for file_name in os.listdir(folder_path):
        if file_name.endswith(('.xlsx', '.xls')) and not file_name.startswith('~'):
            file_path = os.path.join(folder_path, file_name)
            file_type = identify_file_type(file_path)
            if file_type:
                files_by_type[file_type].append(file_path)
                print(f"📁 识别：{file_name} → {file_type}")
    return files_by_type

def classify_selected_files(file_paths):
    files_by_type = {k: [] for k in FILE_TYPE_KEYWORDS.keys()}
    for file_path in file_paths:
        file_type = identify_file_type(file_path)
        if file_type:
            files_by_type[file_type].append(file_path)
            print(f"📁 识别：{os.path.basename(file_path)} → {file_type}")
        else:
            print(f"⚠️  无法识别文件类型：{os.path.basename(file_path)}")
    return files_by_type

# ================= 筛选器功能 =================
def get_yesterday_date():
    from datetime import timedelta
    return (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')

def filter_data_by_yesterday(df, date_columns):
    yesterday = get_yesterday_date()
    print(f"\n📅 日期筛选器：筛选前一天数据 = {yesterday}")

    if df.empty:
        print("⚠️  数据为空，跳过筛选")
        return df

    for col in date_columns:
        if col in df.columns:
            temp_col = pd.to_datetime(df[col], errors='coerce')
            date_only = temp_col.dt.strftime('%Y-%m-%d')
            df = df[date_only == yesterday]

            filtered_count = len(df)
            print(f"  ✅ 已按 '{col}' 筛选：保留 {filtered_count} 条记录")
            break
    return df

# ================= 主程序入口 =================
def main():
    print("=" * 70)
    print("📊 全能数据处理系统：以跟单部为基准 + 左连接合并 + 透视表刷新")
    print("=" * 70)
    
    print("\n【阶段 1/2】数据合并处理")
    print("-" * 70)
    
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    choice = messagebox.askyesno("选择文件方式", "请选择文件获取方式：\n\n【是】选择文件夹（自动扫描）\n【否】手动多选文件")
    root.destroy()
    
    files_by_type = {k: [] for k in FILE_TYPE_KEYWORDS.keys()}
    save_directory = DEFAULT_OUTPUT_DIR
    
    if choice:
        folder_path = select_folder("请选择包含数据文件的文件夹")
        if not folder_path: print("❌ 未选择文件夹，程序退出"); return
        save_directory = folder_path; print(f"\n📂 扫描文件夹：{folder_path}")
        files_by_type = scan_folder_for_files(folder_path)
    else:
        file_paths = select_multiple_files("请选择所有数据文件（可多选）")
        if not file_paths: print("❌ 未选择文件，程序退出"); return
        save_directory = os.path.dirname(file_paths[0]); print(f"\n📂 已选择 {len(file_paths)} 个文件")
        files_by_type = classify_selected_files(file_paths)
    
    print("\n" + "=" * 70); print("📋 文件识别结果："); print("-" * 70)
    type_names = {'follow_up': '跟单部', 'allocation': '分配表', 'pattern_status': '纸样部状态表', 'cutting': '裁床部', 'sewing': '车办部'}
    for file_type, file_list in files_by_type.items():
        count = len(file_list); name = type_names[file_type]
        if count > 0: print(f"✅ {name}: {count} 个文件")
        else: print(f"⚠️  {name}: 未找到文件")
    
    follow_up_columns = ['东明款号', '办名称', '重做版次', '产品类型', '是否经过纸样部', '款式类型', '设计周', '办单提交纸样部日期', '计划制衣件数', '品牌&部门', '跟办人员', '是否需要做办']
    allocation_columns = ['东明款号', '办名称', '重做版次', '制办单分配进度', '办单分配日期', '分配师傅', '品牌&部门', '产品类型', '款式类型', '跟办人员']
    pattern_status_columns = ['东明款号', '办名称', '重做版次', '纸样师傅', '纸样完成进度', '开始处理纸样日期', '纸样完成并转至排唛日期', '品牌&部门', '产品类型', '款式类型', '跟办人员', '计划制衣件数']
    cutting_columns = ['东明款号', '办名称', '重做版次', '裁床完成进度', '开始裁衣日期', '裁衣完成并移交日期', '裁剪总数', '品牌&部门', '产品类型', '款式类型', '跟办人员', '计划制衣件数']
    sewing_columns = ['东明款号', '办名称', '重做版次', '品牌&部门', '车衣进度', '开始车衣时间', '组别', '车办实际总件数', '车办改转态日期']
    merge_keys = ['东明款号', '办名称', '重做版次', '品牌&部门']
    
    print("\n📖 正在读取数据..."); print("-" * 70)
    
    df_follow_up = pd.DataFrame()
    for f in files_by_type['follow_up']:
        df_temp = read_and_extract_data(f, follow_up_columns, os.path.basename(f))
        df_follow_up = pd.concat([df_follow_up, df_temp], ignore_index=True) if not df_temp.empty else df_follow_up
    df_follow_up = normalize_staff_column(df_follow_up)

    df_allocation = pd.DataFrame()
    for f in files_by_type['allocation']:
        df_temp = read_and_extract_data(f, allocation_columns, os.path.basename(f))
        df_allocation = pd.concat([df_allocation, df_temp], ignore_index=True) if not df_temp.empty else df_allocation
    df_allocation = normalize_staff_column(df_allocation)

    df_pattern_status = pd.DataFrame()
    for f in files_by_type['pattern_status']:
        df_temp = read_and_extract_data(f, pattern_status_columns, os.path.basename(f))
        df_pattern_status = pd.concat([df_pattern_status, df_temp], ignore_index=True) if not df_temp.empty else df_pattern_status
    df_pattern_status = normalize_staff_column(df_pattern_status)

    df_cutting = pd.DataFrame()
    for f in files_by_type['cutting']:
        df_temp = read_and_extract_data(f, cutting_columns, os.path.basename(f))
        df_cutting = pd.concat([df_cutting, df_temp], ignore_index=True) if not df_temp.empty else df_cutting
    df_cutting = normalize_staff_column(df_cutting)

    df_sewing = pd.DataFrame()
    for f in files_by_type['sewing']:
        df_temp = read_and_extract_data(f, sewing_columns, os.path.basename(f))
        df_sewing = pd.concat([df_sewing, df_temp], ignore_index=True) if not df_temp.empty else df_sewing
    df_sewing = normalize_staff_column(df_sewing)
    
    print(f"\n📊 各表数据量统计：")
    print(f"  跟单部：{len(df_follow_up)} 行 | 分配表：{len(df_allocation)} 行 | 状态表：{len(df_pattern_status)} 行")
    print(f"  裁床部：{len(df_cutting)} 行 | 车办部：{len(df_sewing)} 行")
    
    df_merged = pd.DataFrame()
    if not df_follow_up.empty:
        df_merged = df_follow_up.copy()
        print(f"\n📌 初始基准数据（跟单部）：{len(df_merged)} 行")
    else:
        print("⚠️  未找到【跟单部】数据，将使用首个有数据的表作为基准")
        for df_temp, name in [(df_allocation, '分配表'), (df_pattern_status, '状态表'), (df_cutting, '裁床部'), (df_sewing, '车办部')]:
            if not df_temp.empty:
                df_merged = df_temp.copy(); print(f"📌 使用 {name} 作为基准：{len(df_merged)} 行"); break

    if not df_merged.empty:
        if not df_allocation.empty: df_merged = merge_dataframes_safe(df_merged, df_allocation, merge_keys, how='left')
        if not df_pattern_status.empty: df_merged = merge_dataframes_safe(df_merged, df_pattern_status, merge_keys, how='left')
        if not df_cutting.empty: df_merged = merge_dataframes_safe(df_merged, df_cutting, merge_keys, how='left')
        if not df_sewing.empty: df_merged = merge_dataframes_safe(df_merged, df_sewing, merge_keys, how='left')

    # ✅ 严格对齐模板列顺序（前36列顺序不变，保证透视缓存范围不动；表尾新增4个年月列）
    output_columns = [
        '东明款号', '办名称', '重做版次', '品牌&部门', '跟办人员', '是否需要做办', '产品类型',
        '是否经过纸样部', '款式类型', '设计周',
        '办单提交纸样部日期', '办单提交纸样部日期（不含时间）', '计划制衣件数',
        '制办单分配进度', '办单分配日期', '分配师傅', '纸样师傅', '纸样完成进度', 
        '开始处理纸样日期', '纸样完成并转至排唛日期', '纸样完成并转至排唛日期（不含时间）',
        '裁床完成进度', '开始裁衣日期', '裁衣完成并移交日期', '裁衣完成并移交日期（不含时间）',
        '裁剪总数', '车衣进度', '开始车衣时间', '组别', '车办实际总件数',
        '车办改转态日期', '车办改转态日期（不含时间）',
        '纸样处理时长 (天)', '裁床处理时长 (天)', '裁床耗时 (天) - 当前', '总耗时 (天)',
        '办单提交纸样部日期（年月）', '纸样完成并转至排唛日期（年月）',
        '裁衣完成并移交日期（年月）', '车办改转态日期（年月）'
    ]
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    if not df_merged.empty:
        for col in output_columns:
            if col not in df_merged.columns: df_merged[col] = None
        df_merged = df_merged[output_columns]
        df_merged = df_merged.drop_duplicates(subset=merge_keys, keep='first')
        
        for col in ['跟办人员', '产品类型', '款式类型']:
            if col in df_merged.columns: df_merged[col] = df_merged[col].fillna('')
        
        # 剔除「分配进度为空 且 跑代码当天-办单提交纸样部日期>3天」的数据
        if '制办单分配进度' in df_merged.columns and '办单提交纸样部日期' in df_merged.columns:
            before = len(df_merged)
            today = pd.Timestamp(datetime.now().date())
            submission_date = pd.to_datetime(df_merged['办单提交纸样部日期'], errors='coerce')
            days_diff = (today - submission_date).dt.days
            need_pattern = df_merged['是否经过纸样部'] == '需要'
            mask_keep = ~(df_merged['制办单分配进度'].isna() & (days_diff > 3) & need_pattern)
            df_removed = df_merged[~mask_keep].copy()
            df_merged = df_merged[mask_keep].copy()
            removed = before - len(df_merged)
            if removed > 0:
                df_removed['超时天数'] = days_diff[~mask_keep]
                print(f"🗑️ 已剔除分配进度为空且超3天未分配的数据：{removed} 条")
                print(f"\n{'='*50}")
                print(f"📋 剔除数据明细如下：")
                print(f"{'='*50}")
                pd.set_option('display.max_columns', None)
                pd.set_option('display.width', 200)
                pd.set_option('display.max_colwidth', 20)
                print(df_removed[['东明款号', '办名称', '办单提交纸样部日期', '制办单分配进度', '超时天数']].to_string(index=False))
                pd.reset_option('display.max_columns')
                pd.reset_option('display.width')
                pd.reset_option('display.max_colwidth')
                # 保存剔除数据到文件
                removed_file = os.path.join(save_directory, f"剔除数据表_{timestamp}.xlsx")
                df_removed.to_excel(removed_file, index=False)
                print(f"📁 剔除数据已保存至：{removed_file}")

        df_merged = add_duration_columns(df_merged)
        df_merged = add_date_only_columns(df_merged)
        
        # 车办实际总件数比裁剪总数大的数据置顶；其余按办单提交纸样部日期由新到旧
        if '车办实际总件数' in df_merged.columns and '裁剪总数' in df_merged.columns:
            df_merged['车办实际总件数'] = pd.to_numeric(df_merged['车办实际总件数'], errors='coerce').fillna(0)
            df_merged['裁剪总数'] = pd.to_numeric(df_merged['裁剪总数'], errors='coerce').fillna(0)
            df_merged['_件数异常'] = df_merged['车办实际总件数'] > df_merged['裁剪总数']
            sort_cols, sort_ascending = ['_件数异常'], [False]
            if '办单提交纸样部日期' in df_merged.columns:
                df_merged['_提交时间'] = pd.to_datetime(df_merged['办单提交纸样部日期'], errors='coerce')
                sort_cols.append('_提交时间')
                sort_ascending.append(False)
            df_merged = df_merged.sort_values(by=sort_cols, ascending=sort_ascending,
                                              kind='stable', na_position='last')
            df_merged = df_merged.drop(columns=[c for c in ('_件数异常', '_提交时间') if c in df_merged.columns])
            df_merged = df_merged.reset_index(drop=True)
            abnormal_count = (df_merged['车办实际总件数'] > df_merged['裁剪总数']).sum()
            if abnormal_count > 0:
                print(f"⚠️  车办实际总件数 > 裁剪总数的数据：{abnormal_count} 条，已置顶；其余按办单提交纸样部日期由新到旧排序")

        df_merged = df_merged[output_columns]

    merged_data_file = os.path.join(save_directory, f"合并数据_完整提取_{timestamp}.xlsx")
    
    if df_merged.empty:
        create_empty_template(merged_data_file)
    else:
        df_merged.to_excel(merged_data_file, index=False)
        print(f"\n✅ 中间合并文件已保存：{merged_data_file}")
    
    print("\n" + "=" * 70); print("【阶段 1】完成！准备进入透视表刷新..."); print("=" * 70)
    
    if df_merged.empty:
        print("⚠️  合并数据为空，跳过透视表刷新步骤。")
        messagebox.showinfo("完成", f"数据合并完成（无数据）。\n文件已保存至：\n{merged_data_file}")
        return

    print("\n【阶段 2/2】透视表刷新处理"); print("-" * 70)

    # 启动前检测并清理上次残留的后台 Excel 进程（会锁住输出文件、干扰本次运行）
    hidden_pids = find_hidden_excel_pids()
    if hidden_pids:
        print(f"\n⚠️  检测到 {len(hidden_pids)} 个残留的后台 Excel 进程：{hidden_pids}")
        kill_them = False
        try:
            _root = tk.Tk(); _root.withdraw(); _root.attributes('-topmost', True)
            kill_them = messagebox.askyesno(
                "残留 Excel 进程",
                f"检测到 {len(hidden_pids)} 个上次运行残留的后台 Excel 进程：\n{hidden_pids}\n\n"
                "它们可能锁住输出文件或干扰本次运行。\n是否现在结束这些进程？")
            _root.destroy()
        except Exception:
            pass
        if kill_them:
            for pid in hidden_pids:
                kill_process(pid)
            print(f"✅ 已结束残留进程：{hidden_pids}")
            time.sleep(1)
    
    template_path = DEFAULT_TEMPLATE_PATH
    if not os.path.exists(template_path):
        print(f"⚠️  默认模板路径不存在：{template_path}")
        template_path = select_file("请选择【数据透视表模板】Excel 文件", [("Excel 模板", "*.xlsx")])
        if not template_path:
            messagebox.showinfo("完成", f"数据合并完成，但未选择模板。\n合并文件：{merged_data_file}")
            return
        save_directory = os.path.dirname(template_path)
    
    report_date = datetime.now().strftime('%Y-%m-%d')
    pivot_output_file = os.path.join(save_directory, f"{report_date} 数据透视表.xlsx")
    
    print(f"📁 模板文件：{template_path}"); print(f"📁 源数据文件：{merged_data_file}"); print(f"📁 输出文件：{pivot_output_file}")
    print("\n⚠️  重要提示：请关闭所有 Excel 文件，特别是目标输出文件！")
    
    try:
        success = refresh_pivot_with_excel(merged_data_file, template_path, pivot_output_file)
        if success:
            messagebox.showinfo("全部完成", f"✅ 数据合并与透视表刷新均已完成！\n\n📊 共处理 {len(df_merged)} 条记录\n📁 中间数据：{os.path.basename(merged_data_file)}\n📈 最终报告：{os.path.basename(pivot_output_file)}\n\n文件已保存至：\n{save_directory}")
        else:
            messagebox.showwarning("警告", "数据合并成功，但透视表刷新失败。请检查 pywin32 库或 Excel 环境。")
    except Exception as e:
        print(f"\n❌ 发生错误：{str(e)}")
        import traceback; traceback.print_exc()
        msg = f"处理过程中发生错误：\n{str(e)}"
        if is_excel_busy_error(e):
            msg += "\n\n提示：Excel 正忙或卡在弹窗/编辑状态。\n请关闭所有 Excel 窗口（含后台进程）后重试。"
        messagebox.showerror("错误", msg)

if __name__ == "__main__":
    main()