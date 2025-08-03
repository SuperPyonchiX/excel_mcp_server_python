#!/usr/bin/env python3
"""FastMCP サーバーのツール一覧を確認するテストスクリプト"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from src.index import mcp
import inspect

def test_tools_list():
    """登録されているツール一覧を表示"""
    print("=" * 50)
    print("FastMCP Server - Tools List Test")
    print("=" * 50)
    
    # FastMCPインスタンスの属性を確認
    print("\nFastMCP インスタンス属性:")
    for attr in dir(mcp):
        if not attr.startswith('_'):
            print(f"  - {attr}")
    
    # ツール関数を直接確認
    print("\n登録されたツール関数:")
    
    # グローバル変数からツール関数を探す
    from src import index
    tool_count = 0
    
    for name in dir(index):
        obj = getattr(index, name)
        if hasattr(obj, '__call__') and hasattr(obj, '_mcp_tool'):
            tool_count += 1
            print(f"\n{tool_count}. {name}")
            if obj.__doc__:
                # docstringの最初の行を表示
                doc_lines = obj.__doc__.strip().split('\n')
                print(f"   説明: {doc_lines[0]}")
            
            # 関数のシグネチャを表示
            try:
                sig = inspect.signature(obj)
                print(f"   パラメータ: {sig}")
            except Exception as e:
                print(f"   パラメータ取得エラー: {e}")
    
    print(f"\n合計ツール数: {tool_count}")
    
    # FastMCPのメソッドも確認
    print("\nFastMCPメソッド:")
    for method_name in ['list_tools', 'get_tool', 'tools']:
        if hasattr(mcp, method_name):
            print(f"  - {method_name}: 利用可能")
        else:
            print(f"  - {method_name}: 利用不可")

if __name__ == "__main__":
    test_tools_list()
