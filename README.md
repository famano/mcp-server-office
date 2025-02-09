# MCP Server Office

A Model Context Protocol (MCP) server providing tools to read/write Microsoft Word (docx) files.

## Usage

Install with pip:

```bash
pip install mcp-server-office
```

Then, start the MCP server:

```bash
mcp-server-office
```

Or using uv, just:

```bash
uvx mcp-server-office
```

### Available Tools

1. `read_docx`: Read complete contents of a docx file including tables and images.

   - Input: `path` (string) - Absolute path to the target file
   - Note: Images are converted to [Image] placeholders, and track changes are not shown
2. `write_docx`: Create a new docx file with given content.

   - Input:
     - `path` (string) - Absolute path to target file
     - `content` (string) - Content to write to the file
   - Note: Use double line breaks for new paragraphs, and [Table] tag with | separators for tables
3. `edit_docx_paragraph`: Make text replacements in specified paragraphs of a docx file.

   - Input:
     - `path` (string) - Absolute path to file to edit
     - `edits` (array) - List of dictionaries containing search/replace text and paragraph index
       - `paragraph_index` (number) - 0-based index of the paragraph to edit
       - `search` (string) - Text to find within the specified paragraph
       - `replace` (string) - Text to replace with
   - Note: Each search string must match exactly once within the specified paragraph
4. `edit_docx_insert`: Insert new paragraphs into a docx file.

   - Input:
     - `path` (string) - Absolute path to file to edit
     - `inserts` (array) - List of dictionaries containing text and optional paragraph index
       - `text` (string) - Text to insert as a new paragraph
       - `paragraph_index` (number, optional) - 0-based index of the paragraph before which to insert. If not specified, insert at the end.

## Requirements

- Python >= 3.12
- Dependencies:
  - mcp[cli] >= 1.2.0
  - python-docx >= 1.1.2

---

# MCP Server Office (日本語)

Microsoft Word (docx) ファイルの読み書きを提供するModel Context Protocol (MCP) サーバーです。

## 使用方法

pipを使用してインストール:

```bash
pip install mcp-server-office
```

MCPサーバーの起動:

```bash
mcp-server-office
```

または、uvを使う場合:

```bash
uvx mcp-server-office
```

### 利用可能なツール

1. `read_docx`: docxファイルの内容を表やイメージを含めて完全に読み取ります。

   - 入力: `path` (文字列) - 対象ファイルの絶対パス
   - 注意: 画像は[Image]というプレースホルダーに変換され、変更履歴は表示されません
2. `write_docx`: 新しいdocxファイルを指定された内容で作成します。

   - 入力:
     - `path` (文字列) - 作成するファイルの絶対パス
     - `content` (文字列) - ファイルに書き込む内容
   - 注意: 段落は2つの改行で区切り、表は[Table]タグと|区切りを使用します
3. `edit_docx_paragraph`: docxファイル内の指定された段落のテキストを置換します。

   - 入力:
     - `path` (文字列) - 編集するファイルの絶対パス
     - `edits` (配列) - 検索/置換テキストと段落インデックスを含む辞書のリスト
       - `paragraph_index` (数値) - 編集する段落の0ベースのインデックス
       - `search` (文字列) - 指定された段落内で検索するテキスト
       - `replace` (文字列) - 置換するテキスト
   - 注意: 各検索文字列は指定された段落内で一度だけマッチする必要があります
4. `edit_docx_insert`: docxファイルに新しい段落を挿入します。

   - 入力:
     - `path` (文字列) - 編集するファイルの絶対パス
     - `inserts` (配列) - テキストとオプションの段落インデックスを含む辞書のリスト
       - `text` (文字列) - 新しい段落として挿入するテキスト
       - `paragraph_index` (数値, オプション) - 挿入する位置の段落の0ベースのインデックス。指定しない場合は末尾に挿入されます。

## 動作要件

- Python >= 3.12
- 依存パッケージ:
  - mcp[cli] >= 1.2.0
  - python-docx >= 1.1.2
