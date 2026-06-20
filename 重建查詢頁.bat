@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在從 DRUGT.xml 重建 領藥查詢.html ...
python build.py
if errorlevel 1 (
  echo.
  echo 失敗：請確認此資料夾有 DRUGT.xml，且已安裝 Python 3。
)
echo.
pause
