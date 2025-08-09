# Excel MCP Server (Python版 - FastMCP使用)

AIエージェントがExcelを自由に操作できるModel Context Protocol (MCP) サーバーのPython実装版です。FastMCPフレームワークを使用してシンプルで使いやすいAPIを提供します。

## 機能

このMCPサーバーは以下のExcel操作機能を提供します：

### ワークブック・ワークシート操作
- `create_workbook` - 新しいExcelワークブックを作成
- `get_workbook_info` - ワークブックの詳細情報を取得
- `add_worksheet` - ワークシートを追加

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

- Python 3.10以上
- uv (推奨) または pip

## クイックスタート

### uv使用（推奨）

1. **uvのインストール** (まだの場合):
   ```bash
   pip install uv
   ```

2. **開発環境のセットアップ**:
   ```bash
   uv sync --dev
   ```
   
   または、バッチファイルを使用（Windows）:
   ```bash
   setup-dev.bat
   ```

3. **サーバーを起動**:
   ```bash
   uv run excel-mcp-server
   ```
   
   または、バッチファイルを使用（Windows）:
   ```bash
   start-server.bat
   ```

4. **テストを実行**:
   ```bash
   uv run python -m pytest
   ```

### 従来のpip使用

1. **仮想環境の作成と依存関係をインストール**:
   ```bash
   python -m venv venv
   # Windows
   venv\Scripts\activate
   # Linux/Mac
   source venv/bin/activate
   
   pip install -r requirements.txt
   ```

2. **サーバーを起動**:
   ```bash
   python src/excel_mcp_server/index.py
   ```

## 開発コマンド（uv使用）

```bash
# 仮想環境に入る
uv shell

# サーバー起動
uv run excel-mcp-server

# テスト実行
uv run python -m pytest

# コードフォーマット
uv run black src/

# リンター実行
uv run ruff src/

# 型チェック
uv run mypy src/

# 開発用依存関係を含む同期
uv sync --dev

# プロダクション用のみ同期
uv sync
```

### MCPクライアントからの使用

FastMCPを使用して作成されたサーバーは、標準的なMCPクライアントから呼び出せます：

```python
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
│   ├── excel_mcp_server/    # メインパッケージ
│   │   ├── __init__.py      # パッケージ初期化
│   │   └── main.py          # メインサーバー実装
│   └── main.py              # 従来形式の実行ファイル（互換性用）
├── test/                    # テストファイル
│   ├── test_uv_setup.py     # uvセットアップテスト
│   ├── excel_integration_test.py  # Excel統合テスト
│   ├── test_tools_list.py   # ツールリストテスト
│   └── output/              # テスト出力ファイル
├── scripts/                 # ユーティリティスクリプト
├── pyproject.toml          # プロジェクト設定（uv対応）
├── uv.lock                 # uvロックファイル
├── requirements.txt        # 従来の依存関係
├── setup-dev.bat          # 開発環境セットアップ（Windows）
├── start-server.bat       # サーバー起動（Windows）
└── README.md              # このファイル
```

## uv（推奨）の利点

- **高速**: Rustで書かれたパッケージマネージャーで非常に高速
- **信頼性**: ロックファイルによる確実な依存関係管理
- **シンプル**: プロジェクト管理が簡潔
- **互換性**: pipと互換性を保ちながら、より良いユーザー体験

## 技術仕様

- **言語**: Python 3.10+
- **パッケージマネージャー**: uv（推奨）またはpip
- **MCPフレームワーク**: FastMCP
- **Excelライブラリ**: openpyxl
- **データ処理**: pandas
- **開発ツール**: pytest, black, ruff, mypy

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
