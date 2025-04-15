import pandas as pd
import os
import re
from datetime import datetime

def normalize_datetime_str(dt_str):
    """多様な日付時刻文字列を'YYYY-MM-DD HH:MM'形式に正規化する"""
    if pd.isna(dt_str) or not isinstance(dt_str, str):
        return None

    dt_str = dt_str.strip()
    if dt_str.lower() == 'none' or dt_str == '':
        return None

    # YYYY年M月D日（曜日）HH:MM～HH:MM 形式
    match = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日(?:（.+?）)?\s*(\d{1,2}):(\d{1,2})(?:[～ー-].*)?', dt_str)
    if match:
        year, month, day, hour, minute = map(int, match.groups())
        try:
            return datetime(year, month, day, hour, minute)
        except ValueError:
            return None # 不正な日付（例: 2月30日）

    # YYYY-MM-DD HH:MM-HH:MM または YYYY-MM-DD HH:MM 形式
    match = re.match(r'(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2})(?:[:|-].*)?', dt_str)
    if match:
        year, month, day, hour, minute = map(int, match.groups())
        try:
            return datetime(year, month, day, hour, minute)
        except ValueError:
            return None

    # YYYY年M月D日 形式 (時刻なし)
    match = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日', dt_str)
    if match:
        year, month, day = map(int, match.groups())
        try:
            # 時刻がない場合は00:00とする
            return datetime(year, month, day, 0, 0)
        except ValueError:
            return None

    # YYYY-MM-DD 形式 (時刻なし)
    match = re.match(r'(\d{4})-(\d{1,2})-(\d{1,2})', dt_str)
    if match:
        year, month, day = map(int, match.groups())
        try:
            return datetime(year, month, day, 0, 0)
        except ValueError:
            return None

    # その他の日付形式を試す (Pandasに任せる)
    try:
        parsed_date = pd.to_datetime(dt_str, errors='coerce')
        return parsed_date if pd.notna(parsed_date) else None
    except Exception:
        return None # パース失敗

def process_event_data(event_data_text, excel_path="event_data.xlsx", sort_by_date=False, ascending=True):
    """
    イベントデータをExcelファイルに保存・更新する関数

    Parameters:
    event_data_text (str): イベントデータのテキスト（CSVフォーマット）
    excel_path (str): 保存するExcelファイルのパス
    sort_by_date (bool): 開催日時でソートするかどうか (デフォルト: False)
    ascending (bool): sort_by_dateがTrueの場合、昇順でソートするかどうか (デフォルト: True)
    """
    # テキストデータから行を分割
    lines = event_data_text.strip().splitlines()

    # linesが空でないかチェック
    if not lines:
        print("エラー: 入力データが空です。")
        return None # または空のDataFrameを返すなど

    # ヘッダー行とデータ行を分離
    header = lines[0].split(',')
    header_str = lines[0]
    # print(f"ヘッダー: {header}") # デバッグコード削除
    data_rows = []

    for i in range(1, len(lines)):
        # 空行やヘッダー行と同じ行はスキップ
        if not lines[i].strip() or lines[i] == header_str:
            continue

        fields = lines[i].split(',')

        # 列数がヘッダーと一致しない場合の処理を追加
        num_columns = len(header)
        if len(fields) < num_columns:
            # 足りない分を空文字で埋める
            fields.extend([''] * (num_columns - len(fields)))
        elif len(fields) > num_columns:
            # 多い分は切り捨てる（またはエラー処理）
            fields = fields[:num_columns]

        data_rows.append(fields)

    # データフレームを作成
    df_new = pd.DataFrame(data_rows, columns=header)
    # print("--- df_new ---") # 前回のデバッグコードは削除
    # print(df_new.head())    # 前回のデバッグコードは削除
    # print(f"df_newの行数: {len(df_new)}") # 前回のデバッグコードは削除
    # print("---------------") # 前回のデバッグコードは削除

    # '開催日時' 列をdatetimeオブジェクトに変換（エラーはNaTにする）
    df_new['開催日時'] = df_new['開催日時'].apply(normalize_datetime_str)
    df_new['申し込み締切日'] = df_new['申し込み締切日'].apply(normalize_datetime_str)

    # 元の文字列を保持する列を追加
    df_new['開催日時_original'] = df_new['開催日時'].fillna('').astype(str)
    df_new['申し込み締切日_original'] = df_new['申し込み締切日'].fillna('').astype(str)

    # 正規化とdatetime型への変換
    df_new['開催日時'] = pd.to_datetime(df_new['開催日時'].apply(normalize_datetime_str), errors='coerce')
    df_new['申し込み締切日'] = pd.to_datetime(df_new['申し込み締切日'].apply(normalize_datetime_str), errors='coerce')

    # 既存のExcelファイルが存在する場合は読み込む
    if os.path.exists(excel_path):
        try:
            # 既存のデータを読み込み
            df_existing = pd.read_excel(excel_path)
            # 元の文字列を保持 (Excelから読み込んだ時点のものを保持)
            # 欠損値(NaN)は空文字列などに変換しておく方が安全
            df_existing['開催日時_original'] = df_existing['開催日時'].fillna('').astype(str)
            df_existing['申し込み締切日_original'] = df_existing['申し込み締切日'].fillna('').astype(str)

            # 既存ファイル内の日付列も正規化を試みる
            if '開催日時' in df_existing.columns:
                df_existing['開催日時'] = df_existing['開催日時'].apply(lambda x: normalize_datetime_str(str(x)) if pd.notna(x) else None)
            if '申し込み締切日' in df_existing.columns:
                df_existing['申し込み締切日'] = df_existing['申し込み締切日'].apply(lambda x: normalize_datetime_str(str(x)) if pd.notna(x) else None)
            # datetime型に変換 (変換できないものはNaTになる)
            df_existing['開催日時'] = pd.to_datetime(df_existing['開催日時'], errors='coerce')
            df_existing['申し込み締切日'] = pd.to_datetime(df_existing['申し込み締切日'], errors='coerce')

            # --- 既存データ内の重複を削除 ---
            if not df_existing.empty:
                # NaT を考慮した重複削除 (開催日時がNaTでもイベント名が同じなら重複とみなす)
                placeholder_nat = pd.Timestamp.min # または他の適切な代替値
                df_existing['開催日時_filled_for_dup'] = df_existing['開催日時'].fillna(placeholder_nat)
                df_existing.drop_duplicates(subset=['イベント名', '開催日時_filled_for_dup'], keep='first', inplace=True)
                df_existing.drop(columns=['開催日時_filled_for_dup'], inplace=True)

            # --- SET を使った重複キー管理 ---
            existing_keys = set()
            if not df_existing.empty:
                # 既存データのキー (イベント名, 開催日時) をセットに追加
                # 開催日時が NaT の場合は None として扱う
                for _, row in df_existing.iterrows():
                    # イベント名から前後の空白とダブルクォートを除去してキーを生成
                    event_name_key = str(row['イベント名']).strip().strip('\"')
                    key_datetime = row['開催日時'] if pd.notna(row['開催日時']) else None
                    existing_keys.add((event_name_key, key_datetime))
            # --- SET 管理ここまで ---

            # --- 新規データの日付補完ロジック ---
            event_to_first_valid_date = {}
            if not df_existing.empty:
                # 既存データから「イベント名」:「最初の有効な開催日時」の辞書を作成
                for _, row in df_existing.iterrows():
                    event_name = row['イベント名']
                    event_date = row['開催日時']
                    # まだ辞書になく、かつ日付が有効なら追加
                    if event_name not in event_to_first_valid_date and pd.notna(event_date):
                        event_to_first_valid_date[event_name] = event_date

                def fill_missing_date(row):
                    # 新規データの開催日時がNaTで、かつ既存データに同じイベント名の有効な日付がある場合
                    if pd.isna(row['開催日時']):
                        existing_date = event_to_first_valid_date.get(row['イベント名'])
                        if existing_date: # 辞書にキーが存在し、値(日付)が取得できた場合
                            return existing_date
                    return row['開催日時']

                # 新規データの開催日時を補完
                df_new['開催日時'] = df_new.apply(fill_missing_date, axis=1)
                # 必要なら申し込み締切日も同様に補完するロジックを追加
            # --- 日付補完ロジックここまで ---

            # --- 新規データから重複を除外 (SETを使用) ---
            new_data_to_add = []
            processed_new_keys = set() # 新規データ内での重複も防ぐため
            for index, row in df_new.iterrows():
                key_datetime = row['開催日時'] if pd.notna(row['開催日時']) else None
                # イベント名から前後の空白とダブルクォートを除去してキーを生成
                event_name_key = str(row['イベント名']).strip().strip('\"')
                current_key = (event_name_key, key_datetime)

                # 既存キーになく、かつ新規データ内でもまだ処理していないキーのみ追加
                if current_key not in existing_keys and current_key not in processed_new_keys:
                    new_data_to_add.append(row)
                    processed_new_keys.add(current_key)

            df_new_filtered = pd.DataFrame(new_data_to_add, columns=df_new.columns)
            # --- 重複除外ここまで ---

            # 既存データとフィルタリングされた新規データを結合
            if not df_existing.empty:
                df_combined = pd.concat([df_existing, df_new_filtered], ignore_index=True)
            else:
                df_combined = df_new_filtered # 既存データがない場合はフィルタリング後新規データのみ

            print(f"既存ファイルを更新しました。合計 {len(df_combined)} 件のイベントデータ（新規 {len(df_new_filtered)} 件追加）")
        except Exception as e:
            print(f"既存ファイルの読み込みまたは処理中にエラーが発生しました: {e}")
            # エラー発生時は新規データのみで処理を試みるか、中断するか選択
            # ここでは新規データのみで続行する例 (エラーハンドリング改善の余地あり)
            df_combined = df_new
            existing_keys = set() # エラー時は既存キーセットを空にする
            print("エラーのため新規データのみで処理します。")
    else:
        # 新規ファイル作成 (重複チェックは不要)
        df_combined = df_new
        print(f"新規ファイルを作成しました。合計 {len(df_combined)} 件のイベントデータ")

    # 日付補完後、重複削除後のデータに対してソートを実行
    if sort_by_date:
        # NaTを末尾に配置してソート
        df_combined = df_combined.sort_values(by='開催日時', ascending=ascending, na_position='last')

    # Excel形式で保存
    try:
        # データを保存
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            # 保存前にdatetimeオブジェクトをフォーマットする
            df_to_save = df_combined.copy()

            # --- NaTを元の文字列で埋める処理 ---
            # '開催日時'がNaTの行のインデックスを取得 (オリジナル列が存在するか確認)
            if '開催日時_original' in df_to_save.columns:
                nat_date_indices = df_to_save[df_to_save['開催日時'].isna()].index
                # NaTの箇所を対応する元の文字列で置き換える
                df_to_save.loc[nat_date_indices, '開催日時'] = df_to_save.loc[nat_date_indices, '開催日時_original']

            # '申し込み締切日'も同様に処理 (オリジナル列が存在するか確認)
            if '申し込み締切日_original' in df_to_save.columns:
                nat_deadline_indices = df_to_save[df_to_save['申し込み締切日'].isna()].index
                df_to_save.loc[nat_deadline_indices, '申し込み締切日'] = df_to_save.loc[nat_deadline_indices, '申し込み締切日_original']
            # --- 埋める処理ここまで ---

            # 保存前にdatetimeオブジェクトをフォーマット、文字列はそのまま
            df_to_save['開催日時'] = df_to_save['開催日時'].apply(format_date)
            df_to_save['申し込み締切日'] = df_to_save['申し込み締切日'].apply(format_date)

            # 不要になった _original 列を削除して保存
            df_to_save.drop(columns=['開催日時_original', '申し込み締切日_original'], errors='ignore').to_excel(writer, index=False)

        print(f"データを '{excel_path}' に保存しました。")

        # 基本的な統計情報を表示
        print_statistics(df_combined)

        return df_combined
    except Exception as e:
        print(f"ファイル保存中にエラーが発生しました: {e}")
        return None

def format_date(value):
    """datetimeオブジェクトをフォーマット、NaTは空文字、その他(文字列など)はそのまま返す"""
    if isinstance(value, datetime):
        try:
            return value.strftime('%Y-%m-%d %H:%M')
        except ValueError: # 年が範囲外など
            return str(value) # エラーの場合は文字列として返す
    elif pd.isna(value): # NaT の場合
        return ''
    else: # 文字列やその他の型
        return str(value)

def print_statistics(df):
    """データフレームの基本的な統計情報を表示する関数"""
    print("\n--- イベントデータ統計情報 ---")
    print(f"総イベント数: {len(df)}")

    # 開催形式ごとのカウント
    format_counts = df['開催形式'].value_counts()
    print("\n開催形式:")
    for format_type, count in format_counts.items():
        print(f"  {format_type}: {count}件")

    # 主催者情報（イベント内容詳細から抽出）
    organizers = df['イベント内容詳細'].str.extract(r'主催[:：]([^。]+)')
    if not organizers.empty:
        organizer_counts = organizers[0].value_counts()
        print("\n主催者（上位5件）:")
        for org, count in organizer_counts.head(5).items():
            if pd.notna(org):
                print(f"  {org.strip()}: {count}件")

# 以下は実行例 (スクリプト直接実行時)
if __name__ == "__main__":
    # スクリプトが直接実行された場合、常に 'new_csv_data.txt' を処理する
    input_file = 'new_csv_data.txt'
    excel_output = 'event_data.xlsx' # 出力先も明示
    print(f"--- CsvCheangeExcel.py 直接実行 ---")
    print(f"入力ファイル: '{input_file}'")
    print(f"出力ファイル: '{excel_output}'")

    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            event_data = f.read()

        if event_data.strip(): # ファイルが空でないか確認
            # process_event_data を実行 (出力ファイルパスも指定)
            # ソートなどのオプションが必要な場合はここに追加:
            # process_event_data(event_data, excel_path=excel_output, sort_by_date=True)
            process_event_data(event_data, excel_path=excel_output)
            print(f"処理が完了しました。結果は '{excel_output}' を確認してください。")
        else:
            print(f"警告: 入力ファイル '{input_file}' が空です。処理をスキップします。")

    except FileNotFoundError:
        print(f"エラー: 入力ファイル '{input_file}' が見つかりません。")
        print(f"'{input_file}' が存在するか、または main.py が先に実行されているか確認してください。")
    except ImportError:
         # このファイル自体に必要なライブラリ (pandasなど) がない場合
         print(f"エラー: 必要なライブラリ (pandas等) が見つかりません。インストールされているか確認してください。")
    except Exception as e:
        import traceback
        print(f"処理中に予期せぬエラーが発生しました: {e}")
        print("--- トレースバック ---")
        traceback.print_exc()
        print("--------------------")

# 元のコマンドライン引数を処理する if/else ブロックは上記で置き換えられました