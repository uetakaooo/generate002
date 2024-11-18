import os 
import re
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '.')))

from pip._vendor.requests.help import chardet
from typing import Set, Dict

import ConvertJsonToDecision as converter
from logger import logger

class LoadInputFile:
    '''
    ディシジョンテーブルが既に存在しているかチェックする関数
    path
    対象のパス
    file_name
    ディシジョンテーブル名
    '''
    @staticmethod
    def check_exists_decision(path, file_name):
        if not os.path.exists(os.path.join(path, file_name)):
            return False
        return True
    '''
    ファイル名をチェックする関数
    folder
    フォルダのパス
    '''
    @staticmethod
    def check_filename(folder):
        # フォルダ内にあるファイルをチェックする
        err_flag = False
        file_list = os.listdir(folder)
        for file in file_list:
            if os.path.isfile(os.path.join(folder, file)) and file.lower().startswith('request_'):
                # ファイルの場合、かつファイル名が「request_」から始まる場合
                result = LoadInputFile.check_request_file(file)
            elif os.path.isfile(os.path.join(folder, file)) and file.lower().startswith('response_'):
                # ファイルの場合、かつファイル名が「response_」から始まる場合
                result = LoadInputFile.check_response_file(file)

                if result:
                    err_flag = True
        
        if err_flag:
            # チェックエラーあり
            return False
        # チェックエラーなし
        return True
    
    '''
    リクエストを定義するファイルの名称をチェックする関数

    ファイル名が特定の命名規則で入力されているかチェック
        ファイル名:request_X1_X2_X3.txt
        X1:ケース番号
            ※必ず2桁で入力
        X2:参照系の場合「R」、登録・更新・削除系の場合「C or U or D」のどちらかが入力されていること
        X3:何回目のリクエストパラメータなのかを記載
            ※必ず1桁入力

        例) 参照APIの場合
            request_01_R_1.txt
        例) 参照APIの場合   ※テスト対象のAPIを呼ぶ前に別のAPIを呼ぶ必要がある場合(セッションに特定の値が必要な場合等)
            request_01_R_1.txt、request_R_01_2.txt

        例) 登録APIの場合   ※ASTAでは対象APIでデータが登録されているか確認するため、「参照⇒登録⇒参照」を行う
            request_01_R_1.txt、request_C_01_2.txt、response_01_R_3.txt
            ※更新APIの場合は「request_01_U_2.txt」、削除APIの場合は「request_01_D_2.txt」
        ※拡張子は「.txt」「.json」のどちらか

        file_name
        ファイル名
    '''
    @staticmethod
    def check_request_file(file_name):
        # 拡張子のチェック
        if not(file_name.endswith('.txt') or file_name.endswith('.properties')):
            # 拡張子が「.txt」以外の場合はエラーとする
            logger.error(f'リクエストパラメータのファイルで拡張子が.txt以外のファイルが存在します:{file_name}')
            return True
        
        # 命名規則のチェック
        target_name = r'request_\d{2}_(C|R|U|D)_\d{1}.(txt|properties)'
        # 正規表現パターン
        pattern = re.compile(target_name)
        if not pattern.search(file_name):
            # 命名規則を満たしていないファイルの場合
            logger.error(f'リクエストパラメータのファイルで命名規則を満たしていないファイルが存在します:{file_name}')
            return True
        
        return False
    
    '''
    レスポンスを定義するファイルの名称をチェックする関数

    ファイル名が特定の命名規則で入力されているかチェック
        ファイル名:response_X1_X2_X3.txt
        X1:ケース番号
            ※必ず2桁で入力
        X2:参照系の場合「R」、登録・更新・削除系の場合「C or U or D」のどちらかが入力されていること
        x3:何回目のリクエストパラメータなのかを記載
            ※必ず1桁で入力
        
        例) 参照APIの場合
            response_01_R_1.txt
        例) 参照APIの場合   ※テスト対象のAPIを呼ぶ前に別のAPIを呼ぶ必要がある場合(セッションに特定の値が必要な場合等)
            response_01_R_1.txt、response_R_01_2.txt

        例) 登録APIの場合   ※ASTAでは対象のAPIで登録されているか確認するため、「参照⇒登録⇒参照」を行う
            response_01_R_1.txt、response_01_C_2.txt、response_01_R_3.txt
            ※更新APIの場合は「request_01_U_2.txt」、削除APIの場合は「request_01_D_2.txt」
        ※拡張子は「.txt」「.json」のどちらか

        file_name
        ファイル名

    '''
    @staticmethod
    def check_response_file(file_name):
        # 拡張子のチェック
        if not(file_name.endswith('.txt')or file_name.endswith('.json')):
            # 拡張子が「.txt」以外の場合はエラーとする
            logger.error(f'レスポンスパラメータのファイル拡張子が.txt・.json以外のファイルが存在します:{file_name}')
            return True
        
        # 命名規則のチェック
        target_name = r'response_\d{2}_(C|R|U|D)_\d{1}.(txt|json)'
        # 正規表現のパターン
        pattern = re.compile(target_name)
        if not pattern.search(file_name):
            # 命名規則を満たしていないファイルの場合
            logger.error(f'レスポンスパラメータのファイルで命名規則を満たしていないファイルが存在します:{file_name}')
            return True
        return False
    
    '''
    プロパティ用のデータとレスポンス用のデータを作成する関数
    folder
    フォルダのパス
    '''
    @staticmethod
    def create_data(folder):
        '''
        ケース数とAPIの呼び出し回数を元にして、多次元配列を作成(array[ケース数][何回目のAPIか])
        
        '''
        file_list = os.listdir(folder)

        # リクエストとレスポンスのファイル名をそれぞれ取得
        request_list = [ temp for temp in file_list if temp.startswith('request_')]
        response_list = [ temp for temp in file_list if temp.startswith('response_')]

        # APIの呼び出し回数とケース数の最大数を取得
        api_num, case_num = LoadInputFile.count_maxcase_maxapi(request_list)

        result_request = [[{} for j in range(api_num)] for i in range(case_num)]
        result_response = [['' for j in range(api_num)] for i in range(case_num)]

        # ケース数分、ループ処理
        for i in range(case_num):
            # APIの呼び出し回数分、ループ処理
            for j in range(api_num):
                # ファイル名を生成
                case = str(i+1).zfill(2)
                regex = f'request_{case}_(C|R|U|D)_{j+1}.(txt|properties)'

                # リクエスト用のファイルが存在しているかチェック
                file_name = [temp for temp in request_list if re.match(regex, temp)]
                if len(file_name) == 0:
                    '''
                    特定のケースではセッションに値を設定した後、テスト対象のAPIを実行するケース等が存在する
                    APIの呼び出し回数分ループ処理を行うので特定のケースは1回目のリクエストと2回目のリクエストファイルが存在し、
                    特定のケースでは2回目のリクエストファイルしか存在しない場合があるので、
                    その際にリクエストファイルが存在しないものに関しては、データを「None」として設定する
                    '''
                    # 対応しているファイルがない場合、値に「None」を設定
                    result_request[i][j] = None
                    continue
                # リクエスト用のファイルを読み込んでDict型のデータに編集
                data = LoadInputFile.load_properties(folder, file_name[0])
                result_request[i][j] = data

                # レスポンス用のファイルが存在しているかチェック
                file_name = []
                regex = f'response_{case}_(C|R|U|D)_{j+1}.(txt|json)'
                file_name = [temp for temp in response_list if re.match(regex, temp)]
                if len(file_name) == 0:
                    '''
                    特定のケースではセッションに値を設定した後、テスト対象のAPIを実行するケース等が存在する
                    APIの呼び出し回数分ループ処理を行うので特定のケースは1回目のリクエストと2回目のリクエストファイルが存在し、
                    特定のケースでは2回目のリクエストファイルしか存在しない場合があるので、
                    その際にリクエストファイルが存在しないものに関しては、データを「None」として設定する
                    '''
                    # 対応しているファイルがない場合、値に「None」を設定
                    result_response[i][j] = None
                    continue

                # レスポンス用のファイルを読み込み、ディシジョンテーブルに貼り付けられる形式にデータを変換する
                data = []
                data, message = converter.conv_response_to_decision(folder, file_name[0])
                if message != '':
                    return None, None, None, None, message
                
                result_response[i][j] = data
        
        return result_request, result_response, case_num, api_num, ''
    

    '''
    '''
    @staticmethod
    def load_data(folder, original_case_num, case_num, api_num, request_data, response_data):
        result = LoadInputFile.create_data(folder)
        # ケース数文、ループ処理
        for case_idx in range(0, case_num):
            if case_idx > original_case_num - 1:
                # case_idがディシジョンテーブルに記載されているケース数より多い場合に以下の処理を行う
                # ※ディシジョンテーブルに記載されているパラメータは取得済みのため
                for api_idx in range(0, api_num):
                    if result[0][case_idx][api_idx] is None:
                        continue
                    request_data[case_idx][api_idx] = result[0][case_idx][api_idx]

                    if result[1][case_idx][api_idx] is None:
                        continue
                    response_data[case_idx][api_idx] = result[1][case_idx][api_idx]
                return request_data, response_data
            
    '''
    プロパティファイルをロードする関数
    path
    パス
    file
    ファイル名
    '''
    @staticmethod
    def load_properties(path, file):
        path = f'{path}\{file}'

        # ファイルの文字コードを検出
        with open(path, 'rb') as file:
            raw_data = file.read()
            result = chardet.detect(raw_data)
            encoding = result['encoding']

        # 読み込み
        with open(path, 'r', encoding=encoding) as file:
            lines = file.readlines()

        # 読み込んだプロパティ用のファイルのデータを辞書型に変換
        data: Dict[str, Set[int]] = {}
        for line in lines:
            line = line.strip()
            if '=' in line:
                key, value = line.split('=', 1)
                if key not in data:
                    data[key] = set()
                data[key].add(value)

        return data
    
    '''
    ファイルの一覧からAPIの呼び出し回数とケース数を取得する関数
    file_list
    読み込んだファイルの一覧
    '''
    @staticmethod
    def count_maxcase_maxapi(file_list):
        # APIを何回呼び出すかをファイル名から取得
        api_num = 0
        case_num = 0
        for temp in file_list:
            api_num_match = re.search(r'(request_\d{2}_(C|R|U|D)_)(\d{1}).(txt|properties)', temp)
            case_num_match = re.search(r'(request_)(\d{2})(_(C|R|U|D)_\d{1}).(txt|properties)', temp)
            if not (api_num_match or case_num_match):
                logger.error('ファイル名に不備があります')
                err_flag = True
            
            # apiの呼び出し回数を更新
            if api_num < int(api_num_match.group(3)):
                api_num = int(api_num_match.group(3))

            # ケース数を更新
            if case_num < int(case_num_match.group(2)):
                case_num = int(case_num_match.group(2))
        
        return api_num, case_num
