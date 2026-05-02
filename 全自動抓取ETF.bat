@echo off
setlocal
chcp 65001 >nul
title Yahoo ETF 全自動抓取工具

echo ======================================================
echo          Yahoo ETF 全自動抓取工具 (背景執行版)
echo ======================================================
echo.
echo 準備執行：
echo 1. 程式將在後台自動啟動瀏覽器 (不會顯示視窗)。
echo 2. 正在自動抓取所有 ETF 最新資料 (約 330 筆)。
echo 3. 請稍候約 30 秒，完成後會自動產出 Excel 檔案。
echo.
echo 正在執行中，請不要關閉此視窗...
echo.

:: 執行 Python Selenium 程式
python selenium_get_etf.py

if %errorlevel% neq 0 (
    color 0C
    echo.
    echo ❌ 執行過程中發生錯誤。
    echo 請確認是否已執行：pip install selenium webdriver-manager pandas openpyxl
) else (
    color 0A
    echo.
    echo ✅ 任務成功完成！
    echo 您的 Excel 檔案已經產生在資料夾中了。
)

echo.
echo ======================================================
echo 按任意鍵關閉此視窗...
pause >nul
