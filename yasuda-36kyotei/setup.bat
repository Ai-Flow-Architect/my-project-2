@echo off
chcp 65001 > nul
echo ========================================
echo  36協定自動化ツール セットアップ
echo ========================================
echo.

echo [1/2] Pythonバージョン確認...
python --version
if errorlevel 1 (
    echo エラー: Pythonがインストールされていません。
    echo https://www.python.org/downloads/ からインストールしてください。
    pause
    exit /b 1
)

echo.
echo [2/2] 必要パッケージをインストール中...
pip install -r requirements.txt
if errorlevel 1 (
    echo エラー: パッケージのインストールに失敗しました。
    pause
    exit /b 1
)

echo.
echo ========================================
echo  セットアップ完了！
echo  run.bat をダブルクリックして実行してください。
echo ========================================
pause
