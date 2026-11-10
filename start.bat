@echo off
chcp 65001 >nul
title Feng Mix - Trộn đề thi trắc nghiệm Online
setlocal enabledelayedexpansion

echo.
echo ========================================================
echo   🎓 FENG MIX - HỆ THỐNG TRỘN ĐỀ THI TRẮC NGHIỆM 2025
echo ========================================================
echo.

cd /d "%~dp0"

:: 1. Kiểm tra Python
set "PYTHON_CMD="

where python >nul 2>nul
if %errorlevel% equ 0 (
    set "PYTHON_CMD=python"
) else (
    where py >nul 2>nul
    if %errorlevel% equ 0 (
        set "PYTHON_CMD=py"
    )
)

if "%PYTHON_CMD%"=="" (
    echo [X] KHÔNG TÌM THẤY PYTHON TRÊN HỆ THỐNG!
    echo.
    echo Vui lòng cài đặt Python (phiên bản 3.9 trở lên).
    echo LƯU Ý KHI CÀI ĐẶT: Nhớ tích chọn ô "Add python.exe to PATH".
    echo.
    echo Đang mở trình duyệt để tải Python...
    start https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [✓] Đã tìm thấy Python:
%PYTHON_CMD% --version
echo.

:: 2. Kiểm tra hoặc tạo môi trường ảo .venv
if not exist ".venv\Scripts\activate.bat" (
    echo [1/3] Đang khởi tạo môi trường ảo Python (.venv)...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo [!] Không thể tạo .venv tự động, sẽ dùng Python trực tiếp của hệ thống.
        set "VENV_PYTHON=%PYTHON_CMD%"
    ) else (
        echo [✓] Đã tạo môi trường ảo .venv thành công.
        set "VENV_PYTHON=.venv\Scripts\python.exe"
    )
) else (
    set "VENV_PYTHON=.venv\Scripts\python.exe"
)

:: 3. Tự động kiểm tra và cài đặt thư viện cần thiết
echo [2/3] Kiểm tra và tự động cài đặt các thư viện cần thiết...
if exist "requirements.txt" (
    set "REQ_FILE=requirements.txt"
) else if exist "backend\requirements.txt" (
    set "REQ_FILE=backend\requirements.txt"
) else (
    set "REQ_FILE="
)

if not "%REQ_FILE%"=="" (
    "%VENV_PYTHON%" -m pip install --upgrade pip -q >nul 2>nul
    "%VENV_PYTHON%" -m pip install -r "%REQ_FILE%" -q
    if errorlevel 1 (
        echo [!] Đang thử lại cài đặt thư viện...
        "%VENV_PYTHON%" -m pip install -r "%REQ_FILE%"
    )
    echo [✓] Đã cài đặt và kiểm tra xong toàn bộ thư viện cần thiết.
) else (
    echo [!] Không tìm thấy file requirements. Đang cài đặt các thư viện mặc định...
    "%VENV_PYTHON%" -m pip install fastapi uvicorn[standard] python-docx openpyxl qrcode[pil] Pillow python-multipart
)

:: 4. Khởi động ứng dụng và tự động mở trình duyệt
echo.
echo [3/3] Đang khởi động máy chủ Feng Mix tại http://localhost:8000...
echo       Trình duyệt web sẽ tự động mở sau vài giây.
echo       (Nhấn tổ hợp phím Ctrl + C trong cửa sổ này nếu muốn dừng máy chủ)
echo.

start "" "http://localhost:8000"

cd backend
"%VENV_PYTHON%" -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload

pause
