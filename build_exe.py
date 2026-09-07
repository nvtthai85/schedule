"""
Script đóng gói ứng dụng Lập Kế Hoạch FECT thành file .exe độc lập chạy trên Windows bằng PyInstaller.
Chỉ cần chạy: python build_exe.py
"""
import os
import sys
import subprocess

def build():
    print("=== BẮT ĐẦU ĐÓNG GÓI ỨNG DỤNG LẬP KẾ HOẠCH FECT THÀNH FILE .EXE ===")
    
    try:
        import PyInstaller
    except ImportError:
        print("[*] Đang cài đặt PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    sep = os.pathsep
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=Lap_Ke_Hoach_FECT",
        "--onedir",
        "--noconfirm",
        "--clean",
        f"--add-data=templates{sep}templates",
        f"--add-data=template_Ke hoach kiem soat, giam sat_team QAFECT.xlsx{sep}.",
        f"--add-data=2026_MA TRAN GIAM SAT TAI FECT_2026_thaotdp_update 04.9.2026.xlsx{sep}.",
        "--hidden-import=uvicorn.logging",
        "--hidden-import=uvicorn.loops",
        "--hidden-import=uvicorn.loops.auto",
        "--hidden-import=uvicorn.protocols",
        "--hidden-import=uvicorn.protocols.http",
        "--hidden-import=uvicorn.protocols.http.auto",
        "--hidden-import=uvicorn.protocols.websockets",
        "--hidden-import=uvicorn.protocols.websockets.auto",
        "--hidden-import=uvicorn.lifespans",
        "--hidden-import=uvicorn.lifespans.on",
        "--hidden-import=openpyxl",
        "--hidden-import=jinja2",
        "--hidden-import=multipart",
        "app.py"
    ]
    
    print("[*] Đang thực thi lệnh đóng gói...")
    subprocess.check_call(cmd)
    
    print("======================================================================")
    print("[✓] ĐÓNG GÓI HOÀN TẤT THÀNH CÔNG!")
    print("Thư mục ứng dụng đóng gói: dist/Lap_Ke_Hoach_FECT/")
    print("👉 Người dùng Windows chỉ cần click đúp vào: Lap_Ke_Hoach_FECT.exe")
    print("======================================================================")

if __name__ == '__main__':
    build()
