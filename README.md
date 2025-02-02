# MCP Server Office

A Model Context Protocol (MCP) server providing tools to read/write Microsoft Word (docx) files.

## Installation

```bash
pip install mcp-server-office==0.1.1
```

## Usage

Start the MCP server:

```bash
mcp-server-office
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
3. `edit_docx`: Make multiple text replacements in a docx file.

   - Input:
     - `path` (string) - Absolute path to file to edit
     - `edits` (array) - List of search/replace pairs
   - Note: Each search string must match exactly once in the document. The formatting of the output file will be deleted.

## Requirements

- Python >= 3.12
- Dependencies:
  - mcp[cli] >= 1.2.0
  - python-docx >= 1.1.2

---

# MCP Server Office (日本語)

Microsoft Word (docx) ファイルの読み書きを提供するModel Context Protocol (MCP) サーバーです。

## インストール

```bash
pip install mcp-server-office==0.1.1
```

## 使用方法

MCPサーバーの起動:

```bash
mcp-server-office
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
3. `edit_docx`: docxファイル内の複数のテキストを置換します。

   - 入力:
     - `path` (文字列) - 編集するファイルの絶対パス
     - `edits` (配列) - 検索/置換のペアのリスト
   - 注意: 各検索文字列はドキュメント内で一度だけマッチする必要があります。また、保存したファイルからは書式がなくなります。

## 動作要件

- Python >= 3.12
- 依存パッケージ:
  - mcp[cli] >= 1.2.0
  - python-docx >= 1.1.2
