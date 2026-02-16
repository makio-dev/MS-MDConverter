@echo off
chcp 65001 >nul

if not exist "venv\Scripts\activate.bat" (
    echo エラー: 仮想環境が見つかりません。
    echo 先に setup.bat を実行してください。
    pause
    exit /b 1
)

call venv\Scripts\activate.bat
python md_converter.py
pause
