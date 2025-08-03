#!/usr/bin/env python3
"""
Excel操作統合テスト
実際のExcelファイル操作機能を包括的にテストします
"""

import asyncio
import time
import sys
import csv
from pathlib import Path

# パスを設定
TEST_DIR = Path(__file__).parent / "output"
TEST_DIR.mkdir(exist_ok=True)

sys.path.append(str(Path(__file__).parent.parent / "src"))

# FastMCPベースのツール関数をインポート
from index import (
    create_workbook, open_workbook, add_worksheet, set_cell_value, 
    get_cell_value, set_range_values, get_range_values, format_cell,
    add_formula, find_data, export_to_csv, close_workbook
)

# テスト設定
TEST_CONFIG = {
    'filePath': str(TEST_DIR / 'test-workbook.xlsx'),
    'csvPath': str(TEST_DIR / 'test-export.csv'),
    'sheetName': 'TestSheet',
    'timeout': 1.0  # seconds
}

# テストデータ定義
TEST_DATA = {
    'sampleData': [
        ['商品名', '価格', '在庫', '売上'],
        ['商品A', 1000, 50, '=B2*C2'],
        ['商品B', 1500, 30, '=B3*C3'],
        ['商品C', 800, 75, '=B4*C4']
    ],
    
    'headerFormat': {
        'font': {
            'bold': True,
            'size': 12,
            'color': 'FF000080'
        },
        'fill': {
            'type': 'pattern',
            'pattern': 'solid',
            'fgColor': 'FFE0E0E0'
        }
    }
}


async def delay(seconds: float):
    """待機ユーティリティ"""
    await asyncio.sleep(seconds)


async def execute_test_with_delay(test_name: str, test_func, test_id: int):
    """
    テストを実行し結果を表示する
    """
    print(f"\n📤 [{test_id}] {test_name} を実行中...")
    try:
        result = await asyncio.get_event_loop().run_in_executor(None, test_func)
        message = str(result)
        print(f"✅ [{test_id}] {message}")
        return True
    except Exception as e:
        print(f"❌ [{test_id}] エラー: {str(e)}")
        return False


async def run_excel_integration_test():
    """Excel操作の統合テストを実行"""
    print('=== Excel操作統合テスト ===\n')
    print(f'📁 テストファイル: {TEST_CONFIG["filePath"]}')
    print(f'📁 CSV出力先: {TEST_CONFIG["csvPath"]}')
    
    # 既存ファイルを削除
    for file_path in [TEST_CONFIG['filePath'], TEST_CONFIG['csvPath']]:
        path_obj = Path(file_path)
        if path_obj.exists():
            path_obj.unlink()
    
    completed_tests = 0
    total_tests = 10
    
    print('\n🚀 テスト開始\n')
    
    # テストシーケンス実行
    test_results = await run_test_sequence()
    completed_tests = sum(test_results)
    
    # テスト結果サマリー
    await print_test_summary(completed_tests, total_tests)
    
    return completed_tests == total_tests


async def run_test_sequence():
    """テストシーケンスを実行"""
    test_results = []
    test_id = 1
    
    # 1. ワークブック作成
    result = await execute_test_with_delay(
        'ワークブック作成', 
        lambda: create_workbook(filePath=TEST_CONFIG['filePath']),
        test_id
    )
    test_results.append(result)
    test_id += 1
    await delay(TEST_CONFIG['timeout'])
    
    # 2. ワークシート追加
    result = await execute_test_with_delay(
        'ワークシート追加',
        lambda: add_worksheet(
            filePath=TEST_CONFIG['filePath'],
            sheetName=TEST_CONFIG['sheetName']
        ),
        test_id
    )
    test_results.append(result)
    test_id += 1
    await delay(TEST_CONFIG['timeout'])
    
    # 3. タイトル設定
    result = await execute_test_with_delay(
        'タイトル設定',
        lambda: set_cell_value(
            filePath=TEST_CONFIG['filePath'],
            sheetName=TEST_CONFIG['sheetName'],
            cell='A1',
            value='Excel MCP 統合テスト レポート'
        ),
        test_id
    )
    test_results.append(result)
    test_id += 1
    await delay(TEST_CONFIG['timeout'])
    
    # 4. サンプルデータ入力
    result = await execute_test_with_delay(
        'サンプルデータ入力',
        lambda: set_range_values(
            filePath=TEST_CONFIG['filePath'],
            sheetName=TEST_CONFIG['sheetName'],
            startCell='A3',
            values=TEST_DATA['sampleData']
        ),
        test_id
    )
    test_results.append(result)
    test_id += 1
    await delay(TEST_CONFIG['timeout'])
    
    # 5. ヘッダー書式設定
    result = await execute_test_with_delay(
        'ヘッダー書式設定',
        lambda: format_cell(
            filePath=TEST_CONFIG['filePath'],
            sheetName=TEST_CONFIG['sheetName'],
            cell='A3',
            formatSpec=TEST_DATA['headerFormat']
        ),
        test_id
    )
    test_results.append(result)
    test_id += 1
    await delay(TEST_CONFIG['timeout'])
    
    # 6. 合計数式追加
    result = await execute_test_with_delay(
        '合計数式追加',
        lambda: add_formula(
            filePath=TEST_CONFIG['filePath'],
            sheetName=TEST_CONFIG['sheetName'],
            cell='E7',
            formula='=SUM(E4:E6)'
        ),
        test_id
    )
    test_results.append(result)
    test_id += 1
    await delay(TEST_CONFIG['timeout'])
    
    # 7. セル値取得テスト
    result = await execute_test_with_delay(
        'セル値取得テスト',
        lambda: get_cell_value(
            filePath=TEST_CONFIG['filePath'],
            sheetName=TEST_CONFIG['sheetName'],
            cell='A1'
        ),
        test_id
    )
    test_results.append(result)
    test_id += 1
    await delay(TEST_CONFIG['timeout'])
    
    # 8. 範囲データ取得テスト
    result = await execute_test_with_delay(
        '範囲データ取得テスト',
        lambda: get_range_values(
            filePath=TEST_CONFIG['filePath'],
            sheetName=TEST_CONFIG['sheetName'],
            rangeAddr='A3:D6'
        ),
        test_id
    )
    test_results.append(result)
    test_id += 1
    await delay(TEST_CONFIG['timeout'])
    
    # 9. データ検索テスト
    result = await execute_test_with_delay(
        'データ検索テスト',
        lambda: find_data(
            filePath=TEST_CONFIG['filePath'],
            sheetName=TEST_CONFIG['sheetName'],
            searchValue='商品A'
        ),
        test_id
    )
    test_results.append(result)
    test_id += 1
    await delay(TEST_CONFIG['timeout'])
    
    # 10. CSV出力テスト
    result = await execute_test_with_delay(
        'CSV出力テスト',
        lambda: export_to_csv(
            filePath=TEST_CONFIG['filePath'],
            sheetName=TEST_CONFIG['sheetName'],
            csvPath=TEST_CONFIG['csvPath']
        ),
        test_id
    )
    test_results.append(result)
    
    return test_results


async def print_test_summary(completed_tests: int, total_tests: int):
    """テスト結果サマリーを出力"""
    print('\n' + '=' * 50)
    print('🎉 Excel操作統合テスト完了')
    print('=' * 50)
    
    try:
        # ファイル存在確認
        excel_path = Path(TEST_CONFIG['filePath'])
        csv_path = Path(TEST_CONFIG['csvPath'])
        
        print('� 生成されたファイル:')
        
        if excel_path.exists():
            excel_size = excel_path.stat().st_size
            print(f'   📈 Excel: {excel_path}')
            print(f'      サイズ: {excel_size} bytes')
        else:
            print(f'   ❌ Excel: {excel_path} (見つかりません)')
        
        if csv_path.exists():
            csv_size = csv_path.stat().st_size
            print(f'   📋 CSV: {csv_path}')
            print(f'      サイズ: {csv_size} bytes')
            
            # CSV内容の確認
            with open(csv_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()[:5]  # 最初の5行
            
            print(f'\n� CSV出力内容プレビュー:')
            print(''.join(lines).rstrip())
        else:
            print(f'   ❌ CSV: {csv_path} (見つかりません)')
        
    except Exception as error:
        print(f'⚠️  ファイル確認中にエラーが発生しました: {error}')
    
    if completed_tests == total_tests:
        print('\n✅ 全ての機能が正常に動作しました！')
    else:
        print(f'\n⚠️  {total_tests}個中{completed_tests}個のテストが完了しました')


# テスト実行
if __name__ == "__main__":
    try:
        success = asyncio.run(run_excel_integration_test())
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"❌ テスト実行エラー: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
