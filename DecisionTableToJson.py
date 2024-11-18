import pandas as pd
import os
import json
import logging
from pathlib import Path
import configparser
import re
from collections import defaultdict
import warnings
import traceback
from distutils.util import strtobool
import ctypes
import shutil

# ロギング設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', filemode='w', filename='execute.log')
logger = logging.getLogger()
warnings.simplefilter(action='ignore', category=UserWarning)

#メッセージダイアログ
def outputMsgBox(msg, title):
    msgBox = ctypes.winDLL.user32.MessageBoxW
    msgBox(0, msg, title, 0x00000000|0x00000010|0x00040000)

def set_nested_value(d, path, value):
    """
    ドットで区切られたパスに基づいて、入れ子になった辞書またはリストに値を設定します。

    Parameters:
    d (dict): 値を設定する対象の辞書。
    path (str): 値を設定する位置を示すドットで区切られたパス。
    value (str): 設定する値。
    """

    if value == "\"\"":
        value = ""
    keys = path.split('.')
    if keys[len(keys)-1] == "":
        return "Key name is not set correctly. "
    
    try:
        for key in keys[:-1]:
            if '[' in key and ']' in key:
                array_name, index_str = re.match(r'(.+)\[(\d+)]', key).groups()
                index = int(index_str)
                d = d.setdefault(array_name, [])
                while len(d) <= index:
                    d.append({})
                d = d[index]
            else:
                d = d.setdefault(key,{})
    except Exception as e:
        return f"{e}"

    final_key = keys[-1]

    try:
        if '[' in final_key and ']' in final_key:
            array_name, index_str = re.match(r'(.+)\[(\d+)]', final_key).groups()
            index = int(index_str)
            d = d.setdefault(array_name, [])
            while len(d) <= index:
                d.append({})
            d = d[index]
        else:
            d = d.setdefault(key, {})
    except Exception as e:
        return f"{e}"

    return ""


def process_column(full_df, df, start_col, check_value, process_row_func):
    """
    特定の条件に基づいてDataFrameの各列を処理し、条件に合致する各行に関数を適用します。
    「テストケース」行を基準にしてヘッダー行のインデックスを動的に決定します。

    Parameters:
    full_df (DataFrame): 処理対象の完全なDataFrame。
    df (DataFrame): 処理する関連する列を含むDataFrame。
    start_col (int): 処理を開始する最初の列のインデックス。
    check_value (str): 各列でチェックする値。
    process_row_func (function): 条件に合致する各行に適用する関数。

    Returns:
    tuple: 結果の辞書とログメッセージのリストを含むタプル。
    """
    results = {}
    log_messages = []
    errMsgList = []

    # 「テストケース」が含まれる行のインデックスを特定する
    test_case_row_index = full_df.index[full_df.iloc[:, 2].astype(str) == "テストケース"].tolist()
    if not test_case_row_index:
        log_messages.append("The row 'テストケース' could not be found.")
        return results, log_messages
    test_case_row_index = test_case_row_index[0] + 1  # テストケースが結合セルのため、の次の行がヘッダー行

    # ヘッダー行は「テストケース」行にある
    header_row_index = test_case_row_index

    for col_index, column in enumerate(full_df.columns[start_col:], start=start_col):
        # ヘッダー行が欠けているかどうかを確認
        if pd.isna(full_df.iloc[header_row_index, col_index]):
            log_messages.append(f"Column header is missing for column index {col_index}. Skipping this column.")
            continue
        checkmark_mask = df.iloc[:, col_index] == check_value
        if not checkmark_mask.any():
            log_messages.append(f"No '{check_value}' found in the test case column '{column}'. Skipping this column.")
            continue
        marked_rows_df = df[checkmark_mask]
        result = {}
        for _, row in marked_rows_df.iterrows():
            errMsg = ""
            errMsg = process_row_func(result, row)
            if errMsg != None and errMsg != "":
                errMsgList.append(str(errMsg) + "Line " + str(row.name +1))

        if "" in result:
            results[column] = result[""]
        else:
            results[column] = result
    return results, log_messages, errMsgList


def find_merged_cell_ranges(df, column_index):
    """
    特定の列における結合されたセルの開始行と終了行のインデックスを見つけます。

    パラメータ:
    df (DataFrame): 検索対象のDataFrame。
    column_index (int): 結合されたセルを検索する列のインデックス。

    戻り値:
    list of tuples: 各タプルが結合されたセルの開始行と終了行のインデックスを含むリスト。
    """
    merged_ranges = []
    current_start = None
    for i, value in enumerate(df.iloc[:, column_index]):
        if pd.notna(value) and current_start is None:
            # Start of a new merged cell
            current_start = i
        elif pd.isna(value) and current_start is not None:
            # Inside a merged cell
            continue
        elif pd.notna(value) and current_start is not None:
            # End of a merged cell
            merged_ranges.append((current_start, i - 1))
            current_start = i
    # Ensure the last merged cell range is added
    if current_start is not None:
        merged_ranges.append((current_start, len(df) - 1))
    return merged_ranges


def get_last_row_of_merged_cell(df, column_index, search_value):
    """
    特定の列において、指定した値を含む結合されたセルの最後の行のインデックスを取得します。

    パラメータ:
    df (DataFrame): 検索対象のDataFrame。
    column_index (int): 結合されたセルを検索する列のインデックス。
    search_value (str): 結合されたセル内で検索する値。

    戻り値:
    int: 検索した値を含む結合されたセルの最後の行のインデックス、見つからない場合は-1。
    """
    merged_ranges = find_merged_cell_ranges(df, column_index)
    for start, end in merged_ranges:
        if df.iloc[start, column_index] == search_value:
            return end
    return -1


def find_last_rows(df, column_index, search_value):
    """
    特定の列において、指定した値を含む結合されたセルの最後の行のインデックスをすべて取得します。

    パラメータ:
    df (DataFrame): 検索対象のDataFrame。
    column_index (int): 結合されたセルを検索する列のインデックス。
    search_value (str): 結合されたセル内で検索する値。

    戻り値:
    list: 検索した値を含む結合されたセルの最後の行のインデックスのリスト。
    """
    merged_ranges = find_merged_cell_ranges(df, column_index)
    last_rows = []
    for start, end in merged_ranges:
        if df.iloc[start, column_index] == search_value:
            last_rows.append(end)
    return last_rows


def create_json_objects(df):
    """
    DataFrameからJSONオブジェクトを生成します。複数の「Condition」と「Expected result」のペアを考慮します。

    パラメータ:
    df (DataFrame): JSONオブジェクトを生成する元となるDataFrame。

    戻り値:
    dict: 生成されたJSONオブジェクト。
    """

    def process_row_for_json(result, row):

        try:
            json_path = row.iloc[3].strip('$')  # キーは4列目
        except Exception as e:
            return "No value set for key name. "

        value = row.iloc[5]  # 値は6列目
        # NaNをNoneに変換
        value = None if pd.isna(value) else value
        # "True"と"False"をbooleanに変換
        if value == "True":
            value = True
        elif value == "False":
            value = False
        elif value == "{}":
            value = {}
        return set_nested_value(result, json_path, value if value != '[]' else [])

    # Find the starting and ending row indices for each 'Expected result' section
    response_param_row_indices = df.index[df.iloc[:, 2].astype(str) == "レスポンスパラメータ"].tolist()
    last_row_of_expected_results = find_last_rows(df, 0, 'Expected result')

    log_messages = []  # ログメッセージを格納するリスト
    json_objects = {}
    errMsgsList = []
    for i in range(len(response_param_row_indices)):
        response_params_df = df.iloc[response_param_row_indices[i] + 1:last_row_of_expected_results[i] + 1]
        result, msgs = process_column(df, response_params_df, 5, '●', process_row_for_json)
        json_objects[f"ExpectedResult_{i + 1}"] = result
        log_messages.extend(msgs)  # ログメッセージを追加

    return json_objects, log_messages, errMsgsList  # タプルとして返す


def create_properties_json(df):
    """
    DataFrameからプロパティオブジェクトを生成します。複数の「Condition」と「Expected result」のペアを考慮します。

    パラメータ:
    df (DataFrame): プロパティオブジェクトを生成する元となるDataFrame。

    戻り値:
    dict: 生成されたプロパティオブジェクト。
    """

    def process_row_for_properties(result, row):
        key = row.iloc[3]  # キーは4列目
        value = row.iloc[5]  # 値は6列目
        if pd.isna(key):
            # key名がnanの場合
            return "key name is not set. "
        
        if pd.isna(value):
            # 値がnanの場合、空文字を設定
            value = ""

        if key in result:
            # キーが既に存在する場合、既存の値に新しい値をカンマ区切りで追加
            result[key] = f"{result[key]},{value}"
        else:
            # キーが存在しない場合、新しいキーと値を追加
            result[key] = value if value != '[]' else []
        return ""


    # Find the starting and ending row indices for each 'Condition' section
    request_param_row_indices = df.index[df.iloc[:, 2].astype(str) == "リクエストパラメータ"].tolist()
    last_row_of_conditions = find_last_rows(df, 0, 'Condition')

    log_messages = []  # ログメッセージを格納するリスト
    properties_objects = {}
    errMsgsList = []
    for i in range(len(request_param_row_indices)):
        params_df = df.iloc[request_param_row_indices[i] + 1:last_row_of_conditions[i] + 1]
        result, msgs = process_column(df, params_df, 5, '●', process_row_for_properties)
        properties_objects[f"Condition_{i + 1}"] = result
        log_messages.extend(msgs)  # ログメッセージを追加

    return properties_objects, log_messages, errMsgsList  # タプルとして返す

def write_to_files(objects, directory, asta_id_values, column_to_test_case_no_map, file_extension, process_item_func):
    """
    指定されたディレクトリにオブジェクトをファイルとして書き込みます。

    Parameters:
    objects (dict): ファイルに書き込むオブジェクト。
    directory (str): ファイルを書き込むディレクトリ。
    asta_id_values (list): ファイル名に含めるASTA IDのリスト。
    column_to_test_case_no_map (dict): 列のインデックスからテストケース番号へのマッピング。
    file_extension (str): 作成するファイルの拡張子。
    process_item_func (function): ファイルに書き込む前に各アイテムを処理する関数。
    """
    if file_extension == "properties":
        rename_file(0,asta_id_values)
    else:
        rename_file(1, asta_id_values)

    if directFlg == True:
        #ファイルの出力先を変更
        reg = re.compile(r'astaid_\d{4}',re.IGNORECASE)
        #ASTAIDを抽出
        temp = re.search(reg, asta_id_values[0])
        temp = temp.group()
        directory = os.path.join(directory,temp)

    try:
        Path(directory).mkdir(parents=True, exist_ok=True)

        # 各ASTAIDとテストケース番号の組み合わせごとに出現回数をカウント
        occurrence_counts = defaultdict(int)
        for obj_key in objects.keys():
            condition_num = int(obj_key.split('_')[-1])
            asta_id_value = asta_id_values[condition_num - 1]
            for column in objects[obj_key].keys():
                test_case_no = column_to_test_case_no_map.get(column, 'default')
                file_base_name = f"{asta_id_value}_{test_case_no:03}"
                occurrence_counts[file_base_name] += 1

        # ファイル名のベース部分ごとにファイル数をカウント
        file_name_counts = defaultdict(int)
        for obj_key, item in objects.items():
            condition_num = int(obj_key.split('_')[-1])
            asta_id_value = asta_id_values[condition_num - 1]
            for column, content in item.items():
                test_case_no = column_to_test_case_no_map.get(column, 'default')
                file_base_name = f"{asta_id_value}_{test_case_no:03}"
                file_name_counts[file_base_name] += 1
                file_num = file_name_counts[file_base_name]
                # 重複がある場合は連番をつける
                filename = f"{file_base_name}_{file_num:02}.{file_extension}" if occurrence_counts[
                                                                                     file_base_name] > 1 else f"{file_base_name}.{file_extension}"
                filename = filename.replace('/', '-').replace('\\', '-').strip()
                file_path = os.path.join(directory, filename)
                defaultEncod = 'utf-8'
                if file_path.endswith(".properties"):
                    defaultEncod = 'MS932'

                with open(file_path, 'w', encoding=defaultEncod) as file:
                    process_item_func(file, content)

                # Jsonを出力した際、勝手に特定の文字がエスケープされてしまう。
                # ※￥nの文字列が￥￥nで保存されてしまう
                # 保存したJsonファイルをテキストとして読み込み置換して「￥n」を「￥￥”」を「￥”」に変更する
                if (file_path.endswith(".json")):
                    with open(file_path, encoding=defaultEncod) as file:
                        data = file.read()

                    data = data.replace("\\\\n","\\n")
                    data = data.replace("\\\\","\\")

                    with open(file_path, 'w', encoding=defaultEncod) as file:
                        file.write(data)
                else:
                    # propertiesを出力した際、配列項目で空文字を設定している場合、ダブルクォーテーションが入ってしまう
                    # ※「"",1,2,"",""」のように空文字の箇所にダブルクォーテーションが入る
                    # 保存したJsonファイルをテキストとして読み込み置換して「"",」を「,」、「""\n」を「\n」に変換する
                    with open(file_path, 'w', encoding=defaultEncod) as file:
                        data = file.read()

                    data = re.sub('"",', ',', data)
                    data = re.sub('""\n', '\\n', data)

                    with open(file_path, 'w', encoding=defaultEncod) as file:
                        file.write(data)

        logger.info(f"{file_extension.upper()} files have been successfully written to the directory: {directory}")
    except Exception as e:
        logger.error(f"Failed to write {file_extension.upper()} files: {e}")
        raise

def write_json_to_files(json_objects, directory, asta_id_values, column_to_test_case_no_map):
    """
    JSONオブジェクトをファイルとして書き込みます。

    Parameters:
    json_objects (dict): ファイルに書き込むJSONオブジェクト。
    directory (str): ファイルを書き込むディレクトリ。
    asta_id_values (list): ファイル名に含めるASTA IDのリスト。
    column_to_test_case_no_map (dict): 列のインデックスからテストケース番号へのマッピング。
    """

    def process_json(file, json_obj):
        json.dump(json_obj, file, ensure_ascii=False, indent=4)

    logger.info('json_dir')
    logger.info(directory)
    write_to_files(json_objects, directory, asta_id_values, column_to_test_case_no_map, 'json', process_json)


def write_properties_to_files(properties_objects, directory, asta_id_values, column_to_test_case_no_map):
    """
    プロパティオブジェクトをファイルとして書き込みます。

    Parameters:
    properties_objects (dict): ファイルに書き込むプロパティオブジェクト。
    directory (str): ファイルを書き込むディレクトリ。
    asta_id_values (list): ファイル名に含めるASTA IDのリスト。
    column_to_test_case_no_map (dict): 列のインデックスからテストケース番号へのマッピング。
    """

    def process_properties(file, properties):
        for key, value in properties.items():
            if pd.isna(value) or value == '""':
                file.write(f"{key}=\n")
            else:
                file.write(f"{key}={value}\n")

    logger.info('properties_dir')
    logger.info('directory')
    write_to_files(properties_objects, directory, asta_id_values, column_to_test_case_no_map, 'properties',
                   process_properties)

def rename_file(type, asta_id_values):
    """
    decison_tableに記載したASTAIDをチェックし、
    requestDataは「ASTAID_XXXX」、
    responseDataは「Result_ASTAID_XXXX」になるようにリネーム
    """
    logger.info("before:"+(", ").join(asta_id_values))
    #大小関係なしで「ASTAID_XXXX」「ASTAID_XXXX_XX」「ASTAID_XXXX_XXX」「ASTAID_XXXX_XXX_XXX」を検索出来る正規表現
    reg = re.compile(r'astaid_\d{4}(_\d{1,4})?(_\d{1,3})?', re.IGNORECASE)

    for i, tempVal in enumerate(asta_id_values):
        #regで指定した文字列を検索
        temp = re.search(reg, tempVal)
        #検索にHITした文字を抽出
        temp = temp.group()

        #値がNull or ""以外の場合
        if not (not temp):
            #idを取り出し
            temp = temp[7:]

            if type == 0:
                #「ASTAID＿」と結合
                asta_id_values[i] = "ASTAID_" + temp
            else:
                #「Result_ASTAID_と結合」
                asta_id_values[i] = "Result_ASTAID_" + temp
    logger.info("after:" + (", ").join(asta_id_values))

def find_asta_ids(df):
    """
    DataFrame内のすべてのASTAIDを見つけてリストとして返します。

    Parameters:
    df (DataFrame): 検索対象のDataFrame。

    Returns:
    list: 発見されたASTAIDのリスト。
    """
    asta_id_row_indices = df.index[df.iloc[:, 2].astype(str).str.contains("ASTAID", na=False)].tolist()
    return [df.iloc[idx, 3] for idx in asta_id_row_indices]


def process_file(file_path, output_directory):
    """
    単一のExcelファイルを処理し、その内容に基づいてJSONおよびプロパティファイルを生成します。
    ASTAIDのリストの数とConditionとExpected resultのリストの数が一致しない場合はエラーを返します。

    パラメータ:
    file_path (str): 処理するExcelファイルのパス。
    output_directory (str): 出力ファイルを保存するディレクトリ。
    """
    try:
        xls = pd.ExcelFile(file_path)
        first_sheet_name = xls.sheet_names[0]
        data = pd.read_excel(xls, first_sheet_name, header=None)
        test_case_no_row_index = data.index[data.iloc[:, 2].astype(str) == "テストケースNo."].tolist()
        if not test_case_no_row_index:
            logger.error("The row 'テストケースNo.' could not be found.")
            return
        test_case_no_row_index = test_case_no_row_index[0]
        test_case_no_row = data.iloc[test_case_no_row_index]
        column_to_test_case_no_map = {col_index: int(test_case_no) for col_index, test_case_no in
                                      enumerate(test_case_no_row[5:], start=5) if pd.notnull(test_case_no)}

        json_objects, creation_logs_json, resErrMsgs = create_json_objects(data)

        errFlg = False
        logger.info("==========Cell with error in response param ==================")
        for logList in resErrMsgs:
            for log in logList:
                logger.error(log)
                if errFlg == False:errFlg = True
        logger.info("==============================================================")
        logger.info("")
        logger.info("")

        properties_objects, creation_logs_properties = create_properties_json(data)

        logger.info("")
        logger.info("")

        for log in creation_logs_json + creation_logs_properties:
            logger.info(log)

        asta_id_values = find_asta_ids(data)  # ASTAIDを取得
        if not asta_id_values:
            logger.error("ASTAID not found in the third column.")
            return

        # ASTAIDリストの数とCondition/Expected resultリストの数をチェック
        if len(asta_id_values) != len(json_objects) or len(asta_id_values) != len(properties_objects):
            logger.error(
                f"The number of ASTAIDs ({len(asta_id_values)}) does not match the number of conditions ({len(properties_objects)}) or expected results ({len(json_objects)}).")
            return

        # If no errors, create the file
        global directDir, directFlg
        subDir1 = 'json'
        subDir2 = 'properties'
        if errFlg == False:
            if bool(directFlg) == True:
                output_directory = directDir
                jsonDir = os.path.join(output_directory,'responseData')
                jsonDir = os.path.join(jsonDir,'LT')
                propDir = os.path.join(output_directory,'requestData')
                propDir = os.path.join(propDir,'LT')
            else:
                jsonDir = os.path.join(output_directory, 'json')
                propDir = os.path.join(output_directory, 'properties')


            write_json_to_files(json_objects,
                                jsonDir,
                                asta_id_values,
                                column_to_test_case_no_map)
            write_properties_to_files(properties_objects,
                                      propDir,
                                      asta_id_values,
                                      column_to_test_case_no_map)
        else:
            outputMsgBox("ディシジョンテーブルの作成に失敗しました。\r\nexecute.logを確認してください！！","")
    except Exception as e:
        logger.error(f"An error occurred while processing {file_path}: {e}")
        logger.error(traceback.format_exc())

directFlg = ""
directDir = ""

def change_file_extension(file_path, new_extension):
    #ファイル拡張子をmhtファイルに変更し、新たにファイルを作成します
    base = os.path.splitext(file_path)[0]
    new_file_path = base + new_extension
    shutil.copy(file_path, new_file_path)

config = configparser.ConfigParser()
config.read('config.ini')
folder_path = config['DEFAULT']['FolderPath']

new_extension = ".mht"
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)
        change_file_extension(file_path, new_extension)
def main():
    """
    入力ディレクトリ内のすべてのExcelファイルを処理し、出力ファイルを生成するメイン関数。
    """
    config = configparser.ConfigParser()
    config.read('config.ini')
    input_directory = config['DEFAULT']['InputDirectory']
    output_directory = config['DEFAULT']['OutputDirectory']

    global directDir,directFlg
    directFlg = strtobool(config['DEFAULT']['DirectOutputFlg'])
    directDir = config['DEFAULT']['DirectOutputDir']
    try:
        for file_name in os.listdir(input_directory):
            if file_name.startswith('decision_table') and (file_name.endswith('.xlsx') or file_name.endswith('xls')):
                file_path = os.path.join(input_directory, file_name)
                logger.info(f"Processing file: {file_path}")
                process_file(file_path, output_directory)
                logger.info(f"Successfully processed file: {file_path}")
    except Exception as e:
        logger.error(f"An error occurred in main: {e}")


if __name__ == "__main__":
    main()
