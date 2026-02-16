@echo off
chcp 65001 >nul
echo ============================================
echo   MS-MDConverter セットアップ
echo ============================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo エラー: Pythonがインストールされていません。
    echo https://www.python.org/downloads/ からインストールしてください。
    pause
    exit /b 1
)

REM Create venv
if not exist "venv" (
    echo 仮想環境を作成中...
    python -m venv venv
    echo 仮想環境を作成しました。
) else (
    echo 仮想環境は既に存在します。
)

REM Proxy setting
echo.
echo プロキシを使用しますか？
echo   [1] 使用する
echo   [2] 使用しない
echo.
set /p PROXY_CHOICE=番号を入力 (1 or 2):

set PROXY_OPT=
if "%PROXY_CHOICE%"=="1" (
    echo.
    echo プロキシURLを入力してください。
    echo   例: http://proxy.example.com:8080
    set /p PROXY_URL=URL:
)
if "%PROXY_CHOICE%"=="1" (
    set PROXY_OPT=--proxy %PROXY_URL%
)

REM Activate and install
echo.
echo 依存パッケージをインストール中...
call venv\Scripts\activate.bat
pip install %PROXY_OPT% -r requirements.txt
echo.
echo セットアップ完了！
echo 「run.bat」をダブルクリックして実行できます。
pause
