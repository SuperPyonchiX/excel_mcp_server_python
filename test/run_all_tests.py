#!/usr/bin/env python3
"""
統合テストランナー
全てのテストを実行するためのスクリプト
"""

import subprocess
import sys
import os
from pathlib import Path

# Windows環境での文字化け対策
if sys.platform == "win32":
    os.environ["PYTHONIOENCODING"] = "utf-8"


def run_command(command, description, cwd=None):
    """コマンドを実行して結果を返す"""
    print(f"\n[実行] {description}")
    print(f"   コマンド: {command}")
    
    # プロジェクトルートに移動してからコマンドを実行
    working_dir = cwd if cwd else os.getcwd()
    
    try:
        result = subprocess.run(
            command, 
            shell=True, 
            check=True, 
            capture_output=True, 
            text=True,
            encoding='utf-8',
            errors='replace',
            cwd=working_dir
        )
        print(f"[成功] {description}")
        if result.stdout:
            print(f"   出力: {result.stdout.strip()}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"[失敗] {description}")
        if e.stderr:
            print(f"   エラー: {e.stderr.strip()}")
        if e.stdout:
            print(f"   出力: {e.stdout.strip()}")
        return False
    except Exception as e:
        print(f"[エラー] {e}")
        return False


def main():
    """メイン関数"""
    print("Excel MCP Server テストスイート実行")
    
    # プロジェクトルートを取得
    project_root = Path(__file__).parent.parent
    test_dir = project_root / "test"
    
    print(f"プロジェクトルート: {project_root}")
    print(f"テストディレクトリ: {test_dir}")
    
    tests = [
        # uvセットアップテスト（プロジェクトルートから実行）
        ("uv run python test/test_uv_setup.py", "uvセットアップテスト", project_root),
        
        # pytestでのテスト実行（プロジェクトルートから実行）
        ("uv run python -m pytest -v", "pytestテスト実行", project_root),
        
        # 個別テストファイル実行（テストディレクトリから実行）
        ("uv run python excel_integration_test.py", "Excel統合テスト", test_dir),
        ("uv run python test_tools_list.py", "ツールリストテスト", test_dir),
    ]
    
    success_count = 0
    
    for command, description, working_dir in tests:
        if run_command(command, description, working_dir):
            success_count += 1
    
    print(f"\nテスト結果: {success_count}/{len(tests)} 成功")
    
    if success_count == len(tests):
        print("すべてのテストが成功しました！")
        return 0
    else:
        print("一部のテストが失敗しました。")
        return 1


if __name__ == "__main__":
    sys.exit(main())
