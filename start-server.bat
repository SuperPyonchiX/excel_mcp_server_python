@echo off
REM Excel MCP Server (Python版) 起動スクリプト (uv対応版)

echo Excel MCP Server (Python版) を起動しています...

REM 現在のディレクトリをプロジェクトルートに変更
cd /d "%~dp0"

REM uvがインストールされているかチェック
where uv >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo エラー: uv が見つかりません。
    echo uvをインストールしてください: pip install uv
    pause
    exit /b 1
)

REM 仮想環境の同期（依存関係のインストール）
echo 仮想環境と依存関係を同期しています...
uv sync

REM サーバーを起動
echo サーバーを起動しています...
uv run excel-mcp-server

pause
