@echo off
setlocal
chcp 65001 >nul
title Yahoo ETF 全自動抓取工具 (含自動更新)

echo ======================================================
echo          Yahoo ETF 全自動抓取工具 (版本檢查中)
echo ======================================================
echo.

:: 1. 檢查是否有 Git 並嘗試更新程式碼
where git >nul 2>nul
if %errorlevel% equ 0 (
    if exist ".git" (
        echo [更新檢查] 正在檢查 GitHub 是否有新版本...
        git fetch --quiet
        for /f %%i in ('git rev-list HEAD...origin/main --count') do set CHANGES=%%i
        
        if not "%CHANGES%"=="0" (
            echo [發現更新] 正在自動升級至最新版本...
            git pull origin main --quiet
            echo [更新完成] 程式碼已成功更新。
            echo.
        ) else (
            echo [最新狀態] 您使用的是最新版本。
            echo.
        )
    )
)

echo 準備執行：
echo 1. 程式將在後台自動啟動瀏覽器 (不會顯示視窗)。
echo 2. 正在自動抓取所有 ETF 最新資料 (約 330 筆)。
echo 3. 請稍候約 30-50 秒，完成後會自動產出 Excel 檔案。
echo.
echo 正在執行中，請不要關閉此視窗...
echo.

:: 2. 執行 Python Selenium 程式
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
