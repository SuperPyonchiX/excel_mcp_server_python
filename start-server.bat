@echo off
REM Excel MCP Server (Python版) 起動スクリプト

echo Excel MCP Server (Python版) を起動しています...

REM 現在のディレクトリをプロジェクトルートに変更
cd /d "%~dp0"

REM 依存関係をインストール
echo 依存関係を確認しています...
python -m pip install -r requirements.txt

REM サーバーを起動
echo サーバーを起動しています...
python src\index.py

pause
