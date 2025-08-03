#!/usr/bin/env python3
"""
Excel MCP Server 管理スクリプト (Python版)
"""

import argparse
import os
import sys
import subprocess
from pathlib import Path

# プロジェクトルートディレクトリ
PROJECT_ROOT = Path(__file__).parent.parent
SRC_DIR = PROJECT_ROOT / "src"
SERVER_SCRIPT = SRC_DIR / "index.py"

def start_server():
    """MCPサーバーを起動"""
    print("📚 Excel MCP Server (Python版) を起動しています...")
    
    if not SERVER_SCRIPT.exists():
        print(f"❌ サーバースクリプトが見つかりません: {SERVER_SCRIPT}")
        return False
    
    try:
        # サーバーを起動
        cmd = [sys.executable, str(SERVER_SCRIPT)]
        print(f"🚀 実行コマンド: {' '.join(cmd)}")
        
        # 標準入出力でサーバーを実行
        process = subprocess.Popen(
            cmd,
            cwd=PROJECT_ROOT,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        print("✅ Excel MCP Server が起動しました")
        print("📝 標準入出力でMCPプロトコルが動作しています")
        print("🛑 Ctrl+C で停止できます")
        
        try:
            # プロセスを待機
            process.wait()
        except KeyboardInterrupt:
            print("\n🛑 サーバーを停止しています...")
            process.terminate()
            process.wait()
            print("✅ サーバーが停止しました")
        
        return True
        
    except Exception as e:
        print(f"❌ サーバー起動エラー: {e}")
        return False


def check_status():
    """サーバーの状態確認"""
    print("🔍 Excel MCP Server (Python版) の状態確認")
    print(f"📁 プロジェクトルート: {PROJECT_ROOT}")
    print(f"📄 サーバースクリプト: {SERVER_SCRIPT}")
    print(f"📄 サーバースクリプト存在: {SERVER_SCRIPT.exists()}")
    
    if SERVER_SCRIPT.exists():
        print("✅ サーバースクリプトが見つかりました")
    else:
        print("❌ サーバースクリプトが見つかりません")
        return False
    
    # 依存関係の確認
    print("\n📦 依存関係の確認:")
    required_packages = ["fastmcp", "openpyxl", "pandas"]
    
    for package in required_packages:
        try:
            __import__(package)
            print(f"   ✅ {package}: インストール済み")
        except ImportError:
            print(f"   ❌ {package}: 未インストール")
    
    return True


def install_dependencies():
    """依存関係をインストール"""
    print("📦 依存関係をインストールしています...")
    
    requirements_file = PROJECT_ROOT / "requirements.txt"
    if not requirements_file.exists():
        print(f"❌ requirements.txtが見つかりません: {requirements_file}")
        return False
    
    try:
        cmd = [sys.executable, "-m", "pip", "install", "-r", str(requirements_file)]
        print(f"🚀 実行コマンド: {' '.join(cmd)}")
        
        result = subprocess.run(cmd, cwd=PROJECT_ROOT, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ 依存関係のインストールが完了しました")
            return True
        else:
            print(f"❌ インストールエラー:\n{result.stderr}")
            return False
            
    except Exception as e:
        print(f"❌ インストールエラー: {e}")
        return False


def run_tests():
    """テストを実行"""
    print("🧪 テストを実行しています...")
    
    test_dir = PROJECT_ROOT / "test"
    test_files = [
        test_dir / "fastmcp_test.py"
    ]
    
    success_count = 0
    total_count = len(test_files)
    
    for test_file in test_files:
        if not test_file.exists():
            print(f"⚠️ テストファイルが見つかりません: {test_file}")
            continue
        
        print(f"\n🔍 実行中: {test_file.name}")
        try:
            cmd = [sys.executable, str(test_file)]
            result = subprocess.run(cmd, cwd=PROJECT_ROOT, capture_output=True, text=True)
            
            if result.returncode == 0:
                print(f"✅ {test_file.name}: 成功")
                print(result.stdout)
                success_count += 1
            else:
                print(f"❌ {test_file.name}: 失敗")
                print(f"標準出力:\n{result.stdout}")
                print(f"エラー出力:\n{result.stderr}")
        
        except Exception as e:
            print(f"❌ {test_file.name}: 実行エラー - {e}")
    
    print(f"\n📊 テスト結果: {success_count}/{total_count} 成功")
    return success_count == total_count


def main():
    """メイン関数"""
    parser = argparse.ArgumentParser(description="Excel MCP Server 管理ツール (Python版)")
    parser.add_argument("command", choices=["start", "status", "install", "test"], 
                       help="実行するコマンド")
    
    args = parser.parse_args()
    
    if args.command == "start":
        start_server()
    elif args.command == "status":
        check_status()
    elif args.command == "install":
        install_dependencies()
    elif args.command == "test":
        run_tests()


if __name__ == "__main__":
    main()
