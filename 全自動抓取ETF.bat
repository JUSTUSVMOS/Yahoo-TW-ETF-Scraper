@echo off
setlocal
chcp 65001 >nul
title Yahoo ETF 全自動抓取工具 (含自動更新)

echo ======================================================
echo          Yahoo ETF 全自動抓取工具 (系統檢查中)
echo ======================================================
echo.

:: 1. 檢查並更新程式碼 (Git)
where git >nul 2>nul
if %errorlevel% equ 0 (
    if exist ".git" (
        echo [1/3] 正在檢查版本更新...
        git fetch --quiet
        for /f %%i in ('git rev-list HEAD...origin/main --count') do set CHANGES=%%i
        if not "%CHANGES%"=="0" (
            echo [發現新版本] 正在下載更新...
            git pull origin main --quiet
        ) else (
            echo [版本檢查] 已是最新版本。
        )
    )
) else (
    echo [提示] 未安裝 Git，跳過自動更新。
)

:: 2. 自動安裝/更新必要套件 (Pip)
echo [2/3] 正在檢查環境套件 (這可能需要一點時間)...
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt --quiet --no-warn-script-location
echo [環境檢查] 套件已準備就緒。

:: 3. 執行抓取程式
echo [3/3] 正在啟動全自動抓取 (背景執行)...
echo      (約需 30-50 秒，請不要關閉此視窗)
echo.

python selenium_get_etf.py

if %errorlevel% neq 0 (
    color 0C
    echo.
    echo ❌ 執行過程中發生錯誤。
    echo 請檢查網路連線或 Python 環境。
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
