@echo off
chcp 65001 >nul
title HỆ THỐNG LẬP KẾ HOẠCH KIỂM SOÁT FECT (QA FECT)

echo ======================================================================
echo          HỆ THỐNG TỰ ĐỘNG HÓA LẬP KẾ HOẠCH KIỂM SOÁT FECT
echo ======================================================================
echo.

:: 1. Kiểm tra nếu có sẵn thư mục Python Portable nhúng
if exist "python_windows\python.exe" (
    echo [✓] Tìm thấy môi trường Python Portable tích hợp sẵn.
    echo [*] Đang khởi chạy ứng dụng và tự động mở trình duyệt...
    echo 🌐 Truy cập tại: http://localhost:8000
    echo.
    python_windows\python.exe app.py
    goto :eof
)

:: 2. Kiểm tra nếu máy Windows đã cài đặt Python
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo [✓] Đã phát hiện Python trên máy tính.
    if not exist ".venv" (
        echo [*] Thiết lập môi trường chạy lần đầu tiên (chỉ mất ~30 giây)...
        python -m venv .venv
        call .venv\Scriptsctivate.bat
        python -m pip install --upgrade pip >nul 2>&1
        pip install -r requirements.txt
    ) else (
        call .venv\Scriptsctivate.bat
    )
    echo [*] Đang khởi chạy ứng dụng...
    echo 🌐 Truy cập tại: http://localhost:8000
    python app.py
    goto :eof
)

:: 3. Trường hợp máy chưa cài đặt Python: Tự động tải Python Portable về máy
echo [!] Máy tính của bạn chưa cài đặt Python.
echo [*] Hệ thống đang tự động tải và cấu hình Python Portable siêu nhẹ (~15MB)...
echo     (Không yêu cầu quyền Admin, không can thiệp vào hệ điều hành của bạn)
echo.

powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Write-Host '  -> Đang tải Python Portable...'; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip' -OutFile 'python_temp.zip'; Write-Host '  -> Đang giải nén...'; Expand-Archive -Path 'python_temp.zip' -DestinationPath 'python_windows' -Force; Remove-Item 'python_temp.zip'}"

if exist "python_windows\python.exe" (
    echo [✓] Đã tải xong Python Portable!
    echo [*] Đang cấu hình pip và các thư viện cần thiết...
    
    :: Cấu hình file ._pth để nhận diện site-packages
    powershell -Command "& { = Get-Item 'python_windows\*._pth'; (Get-Content ) -replace '#import site', 'import site' | Set-Content }"

    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Write-Host '  -> Đang cài đặt thư viện hỗ trợ...'; Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile 'get-pip.py'; .\python_windows\python.exe get-pip.py --no-warn-script-location; Remove-Item 'get-pip.py'}"

    echo [*] Đang tải các thư viện nghiệp vụ (openpyxl, fastapi, uvicorn)...
    python_windows\python.exe -m pip install -r requirements.txt --no-warn-script-location

    echo.
    echo ======================================================================
    echo [✓] Cấu hình hoàn tất 100%! Lần sau bạn chỉ cần click là chạy ngay lập tức.
    echo ======================================================================
    echo.
    python_windows\python.exe app.py
    goto :eof
)

echo.
echo [X] Không thể tự động kết nối tải Python Portable.
echo Bạn có thể cài đặt Python từ https://www.python.org (nhớ tích chọn 'Add Python to PATH').
pause
