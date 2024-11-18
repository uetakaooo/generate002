import os
import json
import warnings
import traceback

from logger import logger
from pandas.core.reshape import encoding
from pip._vendor.requests.help import chardet

def conv_data(data, root_key='$'):
    '''
    テキストファイルから読み込んだデータの変換処理を行う
    ディシジョンテーブルに貼り付けるための形式に変換
    '''
    def recursive_json(data, root_key='$'):
        '''
        JSON形式のデータを再帰的に処理する
        '''
        for key, value in data.items():
            # 親階層から順番にkey名を結合
            temp_key = f'{root_key}.{key}'
            if isinstance(value, dict):
                recursive_json(value, temp_key)
            elif isinstance(value, list):
                if len(value) > 0:
                    for index, item in enumerate(value):
                        array_key = f'{temp_key}[{index}]'
                        if isinstance(item, dict):
                            recursive_json(item, array_key)
                        else:
                            output_data[f'{temp_key}[{index}]'] = item
                else:
                    output_data[temp_key] = '[]'
            
            else:
                if value is None:
                    value = 'null'
                elif value is False:
                    value = 'false'

                if value is '':
                    value = '""'
                    # 空文字の場合、エクセルに値を貼り付けた際に「''」になる値を設定

                if temp_key == '$.result.headDate':
                    value = 'YYYY/MM/DD'
                    output_data[temp_key] = value
                elif temp_key != '$.result.headtime':
                    output_data[temp_key] = value
        
        return output_data
    
    output_data = {}
    return recursive_json(data)

def conv_response_to_decision(path, file_name):
    '''
    テキストファイルを読み込み、json形式の内容から
    ディシジョンテーブルに貼り付けるための形式に変換処理を行う

    パラメータ
    path(str)       :処理するファイルのパス
    file_name (str):出力ファイルを保持するディレクトリ
    '''
    file_path = f'{path}\{file_name}'
    try:
        # ファイルの文字コードを検出
        with open (file_path, 'rb') as file:
            raw_data = file.read()
            result = chardet.detect(raw_data)
            encoding = result['encoding']
        
        # ファイルの読み込み
        with open(file_path, 'r', encoding=encoding) as file:
            file_data = file.read()
            
        # JSON形式に変換
        data = json.loads(file_data)
    except json.JSONDecodeError as e:
        # 読み込んだJSON形式のデータがおかしい場合
        logger.error(f'An error occured while processing {file_path}: {e}')
        logger.error(traceback.format_exc())
        return None, 'テキストファイルの内容をJSON形式のデータに変換する処理でエラーが発生しました。\r\nエラーログを確認してください'
    except Exception as e:
        # その他エラー
        logger.error(f'An error occured while processing {file_path}: {e}')
        logger.error(traceback.format_exc())
        return None, 'ファイル読み込みでエラーが発生しました'
    
    # データ変換
    output_data = conv_data(data)

    if output_data is None or len(output_data) == 0:
        logger.error(f'データ変換に失敗しました')
        return None, 'エラーが発生しました'
    
    return output_data, ''
