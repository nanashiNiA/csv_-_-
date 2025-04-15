import os
import sys

# モジュール検索パスにカレントディレクトリを追加 (環境によっては不要)
# sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from txt_set import remove_duplicate_lines
    # モジュール名が CsvCheangeExcel なので、そのままインポート
    from CsvCheangeExcel import process_event_data
except ImportError as e:
    print(f"必要なモジュールのインポートに失敗しました: {e}")
    print("txt_set.py と CsvCheangeExcel.py が main.py と同じディレクトリにあるか確認してください。")
    sys.exit(1)

def main():
    """
    メイン処理フロー:
    1. txt_set.py を使って入力テキストの重複行を除去
    2. CsvCheangeExcel.py を使って結果をExcelに反映
    """
    input_txt_path = 'csv_data.txt'       # 元のデータファイル
    intermediate_txt_path = 'new_csv_data.txt' # 重複除去後の一時ファイル
    output_excel_path = 'event_data.xlsx'  # 最終的なExcelファイル

    print(f"--- ステップ1: '{input_txt_path}' の重複行を除去 ---")
    try:
        remove_duplicate_lines(input_txt_path, intermediate_txt_path)
        print(f"重複除去後のデータを '{intermediate_txt_path}' に保存しました。")
    except FileNotFoundError:
        print(f"エラー: 入力ファイル '{input_txt_path}' が見つかりません。")
        return # 処理を中断
    except Exception as e:
        print(f"ステップ1の処理中にエラーが発生しました: {e}")
        return # 処理を中断

    print(f"\n--- ステップ2: '{intermediate_txt_path}' のデータをExcel ('{output_excel_path}') に反映 ---")
    try:
        # 中間ファイルを読み込む
        with open(intermediate_txt_path, 'r', encoding='utf-8') as f:
            event_data_text = f.read()

        # event_data_text が空でないことを確認
        if not event_data_text.strip():
            print(f"警告: '{intermediate_txt_path}' が空または空白行のみです。Excel処理をスキップします。")
            return

        # CsvCheangeExcel.py の関数を実行
        # ここでソートなどのオプションを指定可能
        # 例: process_event_data(event_data_text, output_excel_path, sort_by_date=True, ascending=False)
        print(f"'{output_excel_path}' の更新処理を開始します...")
        process_event_data(event_data_text, output_excel_path)
        print(f"'{output_excel_path}' の更新処理が完了しました。")

    except FileNotFoundError:
        # remove_duplicate_lines が成功していれば通常ここには来ないはず
        print(f"エラー: 中間ファイル '{intermediate_txt_path}' が見つかりません。")
    except Exception as e:
        print(f"ステップ2のExcel反映処理中にエラーが発生しました: {e}")

    # オプション: 中間ファイルを削除する場合 (コメントアウトされています)
    # try:
    #     if os.path.exists(intermediate_txt_path):
    #         os.remove(intermediate_txt_path)
    #         print(f"\n中間ファイル '{intermediate_txt_path}' を削除しました。")
    # except Exception as e:
    #     print(f"中間ファイル '{intermediate_txt_path}' の削除中にエラーが発生しました: {e}")

if __name__ == "__main__":
    print("処理を開始します...")
    main()
    print("\n全ての処理が完了しました。")
