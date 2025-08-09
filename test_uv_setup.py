#!/usr/bin/env python3
"""
uv対応テストスクリプト
"""

import subprocess
import sys
from pathlib import Path

def run_command(command, description):
    """コマンドを実行して結果を返す"""
    print(f"\n🔄 {description}")
    print(f"   コマンド: {command}")
    
    try:
        result = subprocess.run(
            command, 
            shell=True, 
            check=True, 
            capture_output=True, 
            text=True,
            encoding='utf-8',
            errors='replace'
        )
        print(f"✅ 成功: {description}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 失敗: {description}")
        print(f"   エラー: {e.stderr}")
        return False
    except Exception as e:
        print(f"❌ エラー: {e}")
        return False

def main():
    """メイン関数"""
    print("🚀 Excel MCP Server (uv対応版) テスト開始")
    
    # プロジェクトルートに移動
    project_root = Path(__file__).parent
    print(f"📂 プロジェクトルート: {project_root}")
    
    tests = [
        ("uv --version", "uvのバージョン確認"),
        ("uv sync", "依存関係の同期"),
        ("uv run black --check src/", "コードフォーマット確認"),
        ("uv run ruff check src/", "リンターチェック"),
        ("uv run python -c \"import excel_mcp_server; print('Import OK')\"", "パッケージインポート確認"),
        ("uv run python -c \"from excel_mcp_server.index import mcp; print('MCP instance OK')\"", "MCPインスタンス確認"),
    ]
    
    success_count = 0
    
    for command, description in tests:
        if run_command(command, description):
            success_count += 1
    
    print(f"\n📊 テスト結果: {success_count}/{len(tests)} 成功")
    
    if success_count == len(tests):
        print("🎉 すべてのテストが成功しました！")
        print("\n🎯 次のステップ:")
        print("   1. uv run excel-mcp-server でサーバー起動")
        print("   2. uv run python -m pytest でテスト実行")
        print("   3. uv shell で仮想環境に入る")
        return 0
    else:
        print("⚠️  一部のテストが失敗しました。")
        return 1

if __name__ == "__main__":
    sys.exit(main())
