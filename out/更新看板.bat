@echo off
chcp 65001 >nul
rem 日三牙醫體系 · 班表看板月更新（Windows：按兩下即可執行）
cd /d "%~dp0"

echo ===============================================
echo   日三牙醫體系 . 班表看板 月更新
echo ===============================================
echo.

where python >nul 2>&1
if errorlevel 1 (
  echo 找不到 Python。請先到 https://www.python.org/downloads/ 安裝，
  echo 安裝時記得勾選「Add Python to PATH」，裝好後再按兩下這個檔案。
  echo.
  pause
  exit /b 1
)

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
  echo 正在安裝需要的套件 openpyxl...
  python -m pip install --quiet openpyxl
  if errorlevel 1 (
    echo 安裝失敗。請把這個畫面截圖給負責人。
    echo.
    pause
    exit /b 1
  )
)

set /p PERIOD="要產生哪一個月？格式 YYYY-MM（例如 2026-11）： "
if "%PERIOD%"=="" (
  echo 沒有輸入月份。
  echo.
  pause
  exit /b 1
)

echo.
echo 產生 %PERIOD% 的 Excel 與看板...
echo.
set GEMRAY_PERIOD=%PERIOD%
python build_roster.py                || goto :fail
python build_roster.py --no-assistant  || goto :fail
python build_site.py                   || goto :fail

echo.
echo ===============================================
echo   完成。要上傳的是這個檔案：
echo   %cd%\班表看板.html
echo.
echo   下一步：到 Cloudflare Pages 的專案頁面，
echo   用 Create deployment 把上面那個檔案拖進去。
echo ===============================================
echo.
pause
exit /b 0

:fail
echo.
echo 產生失敗。請把這個畫面截圖給負責人。
echo.
pause
exit /b 1
