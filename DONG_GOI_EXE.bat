@echo off
chcp 65001 >nul
title Đóng Gói Ứng Dụng Thành File EXE (Windows)
echo ======================================================================
echo           ĐÓNG GÓI ỨNG DỤNG LẬP KẾ HOẠCH FECT THÀNH FILE .EXE
echo ======================================================================
echo.
python build_exe.py
pause
