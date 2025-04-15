# イベントデータ処理ツール

テキストファイル（CSV形式）からイベントデータを読み込み、重複を除去してExcelファイルに保存・更新するツールです。

## 概要

このツールは、以下の3つのPythonスクリプトで構成されています。

1.  `txt_set.py`: 入力テキストファイル (`csv_data.txt`) から完全に同一の行を除去し、新しいファイル (`new_csv_data.txt`) を作成します。
2.  `CsvCheangeExcel.py`: `new_csv_data.txt` を読み込み、イベントデータを処理します。
    *   多様な日付形式（例: `YYYY年M月D日 HH:MM`, `YYYY-MM-DD HH:MM`）を正規化します。
    *   既存のExcelファイル (`event_data.xlsx`) が存在する場合は読み込み、既存データ内の重複（イベント名＋開催日時）も除去します。
    *   新規データと既存データを結合し、重複（イベント名＋開催日時）を除去してExcelファイルに保存・更新します。
    *   日付が不明なデータは、元のテキストを保持して表示します。
    *   オプションで、開催日時によるソート機能も利用可能です（`main.py` からは現在未使用）。
3.  `main.py`: 上記の処理フローを実行するメインスクリプトです。`txt_set.py` を実行した後、`CsvCheangeExcel.py` を実行します。

## 必要なライブラリ

このツールを実行するには、以下のPythonライブラリが必要です。

*   **pandas**: データ処理とExcelファイルの読み書きに使用します。
*   **openpyxl**: pandasがExcelファイル (.xlsx) を扱うために必要です。

以下のコマンドでインストールできます。

```bash
pip install pandas openpyxl
```

## ファイル構成

```
.
├── csv_data.txt         # 入力データファイル (CSV形式を想定)
├── new_csv_data.txt     # 重複行除去後の一時ファイル (txt_set.py が生成)
├── event_data.xlsx      # 出力/更新されるExcelファイル
├── txt_set.py           # テキストファイルの重複行を除去するスクリプト
├── CsvCheangeExcel.py   # Excel処理を行うメインロジック
├── main.py              # 処理全体を実行するスクリプト
└── README.md            # このファイル
```

## 実行方法

1.  必要なライブラリをインストールします (`pip install pandas openpyxl`)。
2.  `csv_data.txt` ファイルを準備します（イベントデータが含まれていることを想定）。
3.  以下のコマンドを実行します。

```bash
python main.py
```

これにより、`csv_data.txt` が処理され、結果が `event_data.xlsx` に保存または更新されます。
実行中の各ステップの状況はコンソールに出力されます。