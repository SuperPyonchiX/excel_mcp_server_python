@echo off
REM Excel MCP Server 開発環境セットアップスクリプト (uv版)

echo Excel MCP Server 開発環境をセットアップしています...

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

REM 仮想環境の初期化と依存関係のインストール
echo 仮想環境を初期化し、依存関係をインストールしています...
uv sync --dev

echo 開発環境のセットアップが完了しました！

echo.
echo 使用可能なコマンド:
echo   uv run excel-mcp-server     # サーバーを起動
echo   uv run python -m pytest    # テストを実行
echo   uv run black src/          # コードフォーマット
echo   uv run ruff src/           # リンターを実行
echo   uv shell                   # 仮想環境に入る

pause
