def remove_duplicate_lines(input_path, output_path):
    # 重複を避けるためのセット
    seen = set()
    unique_lines = []

    # 元のファイルを読み込む
    with open(input_path, 'r', encoding='utf-8') as infile:
        for line in infile:
            if line not in seen:
                seen.add(line)
                unique_lines.append(line)

    # 新しいファイルに書き込む
    with open(output_path, 'w', encoding='utf-8') as outfile:
        outfile.writelines(unique_lines)

# 使用例
input_file = 'csv_data.txt'
output_file = 'new_csv_data.txt'
remove_duplicate_lines(input_file, output_file)
