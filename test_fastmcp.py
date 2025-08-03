#!/usr/bin/env python3
"""
FastMCPのツール定義をテストするスクリプト
"""

from fastmcp import FastMCP

# FastMCPサーバーインスタンスを作成
mcp = FastMCP("Test Excel MCP Server")

@mcp.tool()
def test_create_workbook(file_path: str) -> str:
    """
    新しいExcelワークブックを作成します（テスト版）
    
    Args:
        file_path: 作成するExcelファイルの絶対パス
    """
    print(f"Would create workbook at: {file_path}")
    return f"Test: Excelワークブック '{file_path}' を作成しました。"

if __name__ == "__main__":
    print("FastMCP Test Server starting...")
    print("Available tools:")
    # ツール情報を表示
    for tool_name in mcp._tools:
        tool = mcp._tools[tool_name]
        print(f"  - {tool_name}: {tool}")
    mcp.run()
