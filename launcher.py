"""Windows 双击启动入口，中文提示放在 Python 中以避免 CMD 编码问题。"""
import importlib
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def main():
    check_only = "--check" in sys.argv
    mode = "merge" if "merge" in sys.argv else "export"
    print("启动检查：Python", sys.version.split()[0], flush=True)
    missing = []
    for name in ["pymysql", "openpyxl", "cryptography", "pandas", "win32com.client"]:
        try:
            importlib.import_module(name)
        except ImportError:
            missing.append(name)
    if missing:
        print("正在安装缺少的依赖：" + ", ".join(missing), flush=True)
        result = subprocess.run([sys.executable, "-m", "pip", "install", "-r", str(ROOT / "requirements.txt")])
        if result.returncode:
            print("依赖安装失败，请检查网络和上方错误。")
            return result.returncode
    print("启动检查通过。", flush=True)
    if check_only:
        return 0
    script = "process_excel.py" if mode == "merge" else "export_mysql.py"
    command = [sys.executable, "-X", "utf8", str(ROOT / script)]
    if mode == "export":
        command += ["--process", "--user", "yida_ro"]
        print("即将导出并合并，自动读取 Windows 中保存的数据库密码。", flush=True)
    result = subprocess.run(command, cwd=ROOT)
    print("处理完成。" if result.returncode == 0 else "执行未完成，请查看上方错误信息。", flush=True)
    return result.returncode


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        import traceback
        traceback.print_exc()
        sys.exit(1)
