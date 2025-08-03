# Excel MCP Server (Python版 - FastMCP使用)

AIエージェントがExcelを自由に操作できるModel Context Protocol (MCP) サーバーのPython実装版です。FastMCPフレームワークを使用してシンプルで使いやすいAPIを提供します。

## 機能

このMCPサーバーは以下のExcel操作機能を提供します：

### ワークブック・ワークシート操作
- `create_workbook` - 新しいExcelワークブックを作成
- `open_workbook` - 既存のExcelワークブックを開く
- `get_workbook_info` - ワークブックの詳細情報を取得
- `add_worksheet` - ワークシートを追加
- `close_workbook` - ワークブックを閉じる
- `list_open_workbooks` - 開いているワークブック一覧を表示

### セル・範囲操作
- `set_cell_value` - セルに値を設定
- `get_cell_value` - セルの値を取得
- `set_range_values` - 範囲に2次元配列データを設定
- `get_range_values` - 範囲のデータを取得

### 書式設定
- `format_cell` - セルの書式（フォント、塗りつぶし、罫線）を設定

### 数式・計算
- `add_formula` - セルに数式を追加

### データ操作
- `find_data` - ワークシート内でデータを検索

### 出力
- `export_to_csv` - ワークシートをCSVファイルにエクスポート

## 必要条件

- Python 3.8以上
- pip (Pythonパッケージマネージャー)

## クイックスタート

### Windows環境

1. **依存関係をインストール**:
   ```bash
   pip install -r requirements.txt
   ```

2. **サーバーを起動**:
   ```bash
   python src/index.py
   ```
   
   または、バッチファイルを使用:
   ```bash
   start-server.bat
   ```

3. **テストを実行**:
   ```bash
   python scripts/server_manager.py test
   ```
   
   または、バッチファイルを使用:
   ```bash
   run-tests.bat
   ```

### Linux/Mac環境

1. **依存関係をインストール**:
   ```bash
   pip install -r requirements.txt
   ```

2. **サーバーを起動**:
   ```bash
   python src/index.py
   ```

3. **テストを実行**:
   ```bash
   python scripts/server_manager.py test
   ```

## 管理コマンド

サーバー管理スクリプトを使用できます：

```bash
# サーバー起動  
python scripts/server_manager.py start

# 状態確認
python scripts/server_manager.py status

# 依存関係インストール
python scripts/server_manager.py install

# テスト実行
python scripts/server_manager.py test
```

### MCPクライアントからの使用

FastMCPを使用して作成されたサーバーは、標準的なMCPクライアントから呼び出せます：

```python
# サーバーを起動
python src/index.py

# 別のターミナルでクライアントから呼び出し
# 新しいワークブックを作成
{
  "tool": "create_workbook",
  "arguments": {
    "file_path": "C:/path/to/workbook.xlsx"
  }
}

# ワークシートを追加
{
  "tool": "add_worksheet", 
  "arguments": {
    "file_path": "C:/path/to/workbook.xlsx",
    "sheet_name": "Sheet1"
  }
}

# セルに値を設定
{
  "tool": "set_cell_value",
  "arguments": {
    "file_path": "C:/path/to/workbook.xlsx",
    "sheet_name": "Sheet1",
    "cell": "A1",
    "value": "Hello, Excel!"
  }
}

# 範囲にデータを設定
{
  "tool": "set_range_values",
  "arguments": {
    "file_path": "C:/path/to/workbook.xlsx",
    "sheet_name": "Sheet1",
    "start_cell": "A1",
    "values": [
      ["名前", "年齢", "職業"],
      ["田中", 30, "エンジニア"],
      ["佐藤", 25, "デザイナー"]
    ]
  }
}

# セルの書式を設定
{
  "tool": "format_cell",
  "arguments": {
    "file_path": "C:/path/to/workbook.xlsx",
    "sheet_name": "Sheet1",
    "cell": "A1",
    "format_spec": {
      "font": {
        "bold": true,
        "size": 14,
        "color": "FF0000FF"
      },
      "fill": {
        "type": "pattern",
        "pattern": "solid",
        "fgColor": "FFFF00"
      }
    }
  }
}
```
```

## 開発

### テストの実行

```bash
pytest
```

### コードフォーマット

```bash
black src/ test/
```

### 型チェック

```bash
mypy src/
```

## ファイル構造

```
excel_mcp_server_python/
├── src/
│   └── index.py          # メインサーバー実装
├── test/                 # テストファイル
├── scripts/              # ユーティリティスクリプト
├── requirements.txt      # 依存関係
├── pyproject.toml       # プロジェクト設定
└── README.md            # このファイル
```

## 技術仕様

- **言語**: Python 3.8+
- **MCPフレームワーク**: FastMCP
- **Excelライブラリ**: openpyxl
- **データ処理**: pandas

## TypeScript版との違い

このPython版は、TypeScript版と同等の機能を提供しますが、以下の違いがあります：

- **言語**: Python 3.8+を使用
- **MCPフレームワーク**: FastMCPを使用（シンプルなデコレーターベース）
- **Excelライブラリ**: openpyxlを使用（TypeScript版はExcelJS）
- **非同期処理**: 不要（FastMCPが内部で処理）

## FastMCPの利点

- **シンプル**: デコレーターベースでツールを簡単に定義
- **自動スキーマ生成**: 関数のドキュメントから自動でスキーマ生成
- **型安全**: Python標準の型ヒントを活用
- **軽量**: 最小限の依存関係で動作

## ライセンス

ISC

## 貢献

プルリクエストやイシューの報告を歓迎します。
