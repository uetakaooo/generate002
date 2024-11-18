import os

# フォルダパスの指定
folder_path = r'C:\workspace\inputPDF\venv\file'

# ファイルパスの初期化
before_file_path = None
after_file_path = None
merge_file_path = os.path.join(folder_path, 'response_json_merge.txt')  # マージ結果を保存するファイル

# response_json_beforeで始まるファイルとresponse_json_afterで始まるファイルを探す
for file in os.listdir(folder_path):
    if file.startswith('response_json_before') and file.endswith('.txt'):
        before_file_path = os.path.join(folder_path, file)
    elif file.startswith('response_json_after') and file.endswith('.txt'):
        after_file_path = os.path.join(folder_path, file)

# ファイルが見つからない場合はエラーメッセージ
if not before_file_path or not after_file_path:
    print("ファイルが見つかりません。")
    exit()

# response_json_before.txtの内容を読み込む
with open(before_file_path, 'r', encoding='utf-8') as before_file:
    before_data = before_file.readlines()

# response_json_after.txtの内容を読み込む
with open(after_file_path, 'r', encoding='utf-8') as after_file:
    after_data = after_file.readlines()

# ファイルがタブ区切りかコロン区切りかを判定
def detect_delimiter(data):
    for line in data:
        if '\t\t' in line:
            return '\t\t'  # タブ区切り
        elif ':' in line:
            return ':'  # コロン区切り
    return None

# beforeとafterで同じ区切り文字を使用しているか確認
before_delimiter = detect_delimiter(before_data)
after_delimiter = detect_delimiter(after_data)

if before_delimiter != after_delimiter:
    print("エラー: ファイル間で異なる区切り文字が使用されています。")
    exit()

# キーと値を格納する辞書を準備
before_entries = {}
after_entries = {}

# キーと値を解析する関数
def parse_key_value(line, delimiter):
    key_value = line.strip().split(delimiter)
    if len(key_value) == 2:
        return key_value[0], key_value[1]
    return None, None  # 無効な行の場合

# beforeファイルを解析
for line in before_data:
    key, value = parse_key_value(line, before_delimiter)
    if key and value:
        if key not in before_entries:
            before_entries[key] = []
        before_entries[key].append(value)

# afterファイルを解析
for line in after_data:
    key, value = parse_key_value(line, after_delimiter)
    if key and value:
        if key not in after_entries:
            after_entries[key] = []
        after_entries[key].append(value)

# マージされたエントリを保存するリスト
final_entries = []

# afterを優先して、beforeのデータを追加する
for key, after_values in after_entries.items():
    for value in after_values:
        final_entries.append(f"{key}{after_delimiter}{value}")

    # beforeに存在し、値が異なる場合のみ追加
    if key in before_entries:
        for before_value in before_entries[key]:
            if before_value not in after_values:
                final_entries.append(f"{key}{before_delimiter}{before_value}")

# beforeにしかないキーを追加
for key, before_values in before_entries.items():
    if key not in after_entries:
        for before_value in before_values:
            final_entries.append(f"{key}{before_delimiter}{before_value}")

# マージ結果をresponse_json_merge.txtに書き込む
with open(merge_file_path, 'w', encoding='utf-8') as merge_file:
    for entry in final_entries:
        merge_file.write(f"{entry}\n")

print(f"マージファイルが作成されました: {merge_file_path}")
