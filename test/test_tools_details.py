#!/usr/bin/env python3
"""
FastMCP tools/list レスポンス詳細確認
"""

import sys
import asyncio
from pathlib import Path
import json

# プロジェクトルートをPythonパスに追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

print("=== FastMCP tools/list Response Details ===")
print()

async def test_tools_list_details():
    try:
        from src.index import mcp
        
        # 非同期でlist_toolsメソッドを呼び出し
        print("1. Calling await mcp.list_tools()...")
        tools_list = await mcp.list_tools()
        
        print("2. Tools list details:")
        print(f"   Response type: {type(tools_list)}")
        print(f"   Tools count: {len(tools_list)}")
        print()
        
        if len(tools_list) > 0:
            print("3. First tool details:")
            first_tool = tools_list[0]
            print(f"   First tool type: {type(first_tool)}")
            print(f"   First tool attributes: {[attr for attr in dir(first_tool) if not attr.startswith('_')]}")
            
            if hasattr(first_tool, '__dict__'):
                print(f"   First tool data: {first_tool.__dict__}")
            print()
            
            print("4. All tools summary:")
            for i, tool in enumerate(tools_list):
                print(f"   {i+1:2d}. Type: {type(tool)}")
                if hasattr(tool, 'name'):
                    print(f"       Name: {tool.name}")
                if hasattr(tool, 'description'):
                    print(f"       Description: {tool.description}")
                if hasattr(tool, '__dict__'):
                    print(f"       Data: {tool.__dict__}")
                print()
        
        # JSON化を試行
        print("5. JSON serialization attempts:")
        try:
            # Pydanticモデルの場合
            if all(hasattr(tool, 'model_dump') for tool in tools_list):
                json_data = [tool.model_dump() for tool in tools_list]
                print("   Using model_dump():")
                print(json.dumps(json_data, indent=2, ensure_ascii=False))
            elif all(hasattr(tool, 'dict') for tool in tools_list):
                json_data = [tool.dict() for tool in tools_list]
                print("   Using dict():")
                print(json.dumps(json_data, indent=2, ensure_ascii=False))
            else:
                print("   Direct list serialization:")
                print(json.dumps(tools_list, default=str, indent=2, ensure_ascii=False))
        except Exception as e:
            print(f"   JSON serialization error: {e}")
        
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

# メイン実行
if __name__ == "__main__":
    asyncio.run(test_tools_list_details())
    print("\n=== Test Complete ===")
