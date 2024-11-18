import os
import re
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '.')))

import ctypes
import numpy as np
import traceback
from collections import defaultdict

# 自作のクラス
from logger import logger
from ExcelConst import ExcelConst as const
from DecisionExcelUtils import DecisionExcelUtils
from LoadInputFile import LoadInputFile

# メッセージダイアログ
def outputMsgBox(msg, title):
    msgbox = ctypes.windll.user32.MessageBoxW
    msgbox(0, msg, title, 0x00000000 | 0x00000010 | 0x00040000)

def create_decision_excel(folder):
    # フォルダ名の取得
    asta_id = os.path.basename(folder).lower()
    
    # フォルダ内に既にディシジョンテーブルがあるかチェック
    exists_flag = False
    if LoadInputFile.check_exists_decision(folder, f'decision_table_{asta_id}.xlsx'):
        exists_flag = True
    
    if exists_flag:
        path = os.path.join(folder, f'decison_table_{asta_id}.xlsx')
    else:
        # テンプレートの読み込み
        template_path = os.path.abspath(os.path.join(folder, os.pardir))
        template_path = os.path.join(template_path, const.TEMPLATE_FOLDER_NAME)
        path = os.path.join(template_path, const.TEMPLATE_FILE_NAME)
    
    excel = DecisionExcelUtils(path, const.SHEET_NAME, exists_flag)

    # フォルダ内にあるファイルの命名規則をチェック
    if not LoadInputFile.check_filename(folder):
        outputMsgBox('エラーがあります、ログを確認してください', 'エラー')
        return
    
    # リクエスト用、レスポンス用のファイルからデータを取得
    request_data, response_data, case_num, api_num, message = LoadInputFile.create_data(folder)
    if message != '':
        outputMsgBox(message,'エラー')
        return
    
    request_list = [ defaultdict(list) for i in range(api_num)]
    response_list = [[ defaultdict(list) for j in range(case_num)] for i in range(api_num)]
    if exists_flag:
        # ディシジョンテーブルに記載されているケース数のチェックと取得
        orginal_case_num = excel.check_case_count(case_num)
        # ディシジョンテーブルに記載されているリクエストパラメータを取得
        request_data = [[{} for j in range(api_num)] for i in range(case_num)]
        request_data = excel.get_param_list(case_num, api_num, request_data, const.AREA_NAME_REQUEST)
        # ディシジョンテーブルに記載されているレスポンスパラメータを取得
        response_data = [['' for j in range(api_num)] for i in range(case_num)]
        response_data = excel.get_param_list(case_num, api_num, response_data, const.AREA_NAME_RESPONSE)
        # ファイルから取得したパラメータの追加
        request_data, response_data = LoadInputFile.load_data(folder, orginal_case_num, case_num, api_num, request_data, response_data)

    # ケース数分、ループ処理
    for i, case_object in enumerate(request_data, start=0):
        # APIの呼び出し回数分、ループ処理
        for j, api_object in enumerate(case_object, start=0):
            if api_object is None or len(api_object) == 0:
                continue

            # リクエストパラメータ(Key項目)、ループ処理
            for (key, list_value) in api_object.items():
                for idx, value in enumerate(list_value, start=0):
                    if key not in request_list[j] or (len(request_list[j][key]) == 0 or value not in request_list[j][key]):
                        request_list[j][key].append(value)

    # ケース数分、ループ処理
    for i , case_object in enumerate(response_data, start=0):
        # APIの呼び出し回数分、ループ処理
        for j, api_object in enumerate(case_object, start=0):
            if api_object is None or len(api_object) == 0:
                continue

            # リクエストパラメータ(Key項目数分)、ループ処理
            for(key, value) in api_object.items():
                    response_list[j][i][key].append(value)

    if api_num > 1 and not exists_flag:
        # APIの呼び出し回数が1回より多い場合、かつ新規作成の場合にASTAIDやURL等の行を複製
        excel.add_title_area(api_num)

    # ケース数分、●付けをするエリアを追加
    excel.edit_case_input_area(case_num)

    # 各APIのリクエストパラメータを記載する行の行数を調整する
    excel.increase_decrease_named_area_lines(request_list, const.AREA_NAME_REQUEST, 1)
    # 各APIのレスポンスを記載する行の行数を調整する
    excel.increase_decrease_named_area_lines(response_list, const.AREA_NAME_RESPONSE, 1)
    # ●付けエリアを初期化
    target_list = []
    target_list.append(const.AREA_NAME_REQUEST)
    target_list.append(const.AREA_NAME_RESPONSE)
    #target_list.append(const.AREA_NAME_DECISION)
    excel.clear_decision_area(target_list)
    # リクエストパラメータをエクセルに設定
    excel.edit_request_param(request_list, request_data, const.AREA_NAME_REQUEST)
    # レスポンスパラメータをエクセルに設定
    excel.edit_response_param(response_list, const.AREA_NAME_RESPONSE)

    # ケース番号とスクリプト番号の入力
    excel.input_case_no(case_num, api_num)
    # 強調表示ルールの定義を作成
    excel.create_case_condition_rules(case_num, api_num)

    # ファイルを保存
    excel.save_book(f'{folder}\decision_table_{asta_id}.xlsx')

def main():
    '''
    入力ディレクトリ内にあるフォルダとそのフォルダ内にあるファイルを読み取り
    decision-table.xlsxを作成する
    '''
    input_directory = os.getcwd()
    input_directory = f'{input_directory}\decision_tables'

    # 対象フォルダ
    target_name = 'astaid_\d{4}'
    # 正規表現パターン
    pattern = re.compile(target_name, re.IGNORECASE)

    # フォルダ数分、ループ処理を行う
    for temp in os.listdir(input_directory):
        path = os.path.join(input_directory, temp)
        if os.path.isdir(path) and pattern.search(temp):
            try:
                create_decision_excel(path)
            except Exception as e:
                logger.error(str(e))
                logger.error(f'対象：{path}')
                logger.error('----------------------------------')
                logger.error(traceback.format_exc())
                logger.error('----------------------------------')
                logger.error('')
                logger.error('')
                logger.error('')
                logger.error('')
                logger.error('')
                logger.error('')
                outputMsgBox('エラーがあります、ログを確認してください', 'エラー')
                continue

if __name__ == '__main__':
    main()