# MS-MDConverter

MarkdownファイルをExcel（.xlsx）またはWord（.docx）に変換するコマンドラインツールです。

## 対応するMarkdown要素

- 見出し（h1〜h6）
- テーブル（表）
- 箇条書きリスト（順序なし・順序あり）
- コードブロック
- 段落テキスト
- 水平線
- インライン書式（太字・斜体・コード・リンク → プレーンテキストに変換）

## セットアップ

Python 3.10以上が必要です。

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate

pip install -r requirements.txt
```

プロキシ環境の場合:

```bash
pip install --proxy http://proxy.example.com:8080 -r requirements.txt
```

## 使い方

```bash
python md_converter.py
```

1. 出力形式を選択（Excel / Word）
2. 変換するMarkdownファイルのパスを入力
3. 同じフォルダに変換後のファイルが出力されます

## 出力例

### Excel
- 見出しはレベルに応じた色付きヘッダー行
- テーブルは罫線付きのセルに展開
- コードブロックはConsolas + グレー背景

### Word
- 見出しはWordの見出しスタイル（Heading 1〜4）
- テーブルは Table Grid スタイル
- コードブロックはConsolas + グレー背景

## ライセンス

MIT
