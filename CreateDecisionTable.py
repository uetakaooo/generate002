from email import charset
from pandas.core.reshape import encoding
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
from pip._vendor.requests.help import chardet

# ロギング設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', filemode='w', filename='error_CreateDecisonTable.log', encoding='utf-8')
logger = logging.getLogger()
warnings.simplefilter(action='ignore', category=UserWarning)

#メッセージダイアログ
def outputMsgBox(msg, title):
    msgBox = ctypes.winDLL.user32.MessageBoxW
    msgBox(0, msg, title, 0x00000000|0x00000010|0x00040000)

def conv_data(data, root_key='$'):
    """
    テキストファイルから読み込んだデータの変換処理を行う。
    ディシジョンテーブルに張り付けるための形式に変換
    """
    def recursive_json(data, root_key='$'):

     for key, value in data in data.items():
        temp_key =f"{root_key}.{key}"
        if isinstance(value,dict):
            recursive_json(value, temp_key)
        elif isinstance(value, list):
            if len(value)>0:
                for index, item in enumerate(value):
                    array_key = f"{temp_key}[{index}]"
                    if isinstance(item,dict):
                        recursive_json(item,array_key)
            else:
                output_data.append(f"{temp_key}\t\t[]")
        else:
            if value is None:
                value ='null'
            elif value is False:
                value ='false'

            if temp_key == '$.result.headDate':
                value = "YYYY/MM/DD"
                output_data.append(f"{temp_key}\t\t{value}")
            elif temp_key != '$.result,headtime' and not value:
                value = '""""""'
                output_data.append(f"{temp_key}\t\t{value}")
            elif temp_key != '$.result.headtime' :
                output_data.append(f"{temp_key}\t\t{value}")
        
        return output_data
    
    output_data = []
    return recursive_json(data)


def conv_response_to_decision(file_path, output_dir):

    try:
        with open (file_path, 'rb') as file:
            raw_data = file.read()
            result = charset.detect(raw_data)
            encoding = result['encoding']
        
        with open (file_path, 'r', encoding=encoding) as file:
            file_data = file.read()

        data = json.loads(file_data)
    except json.JSONDecodeError as e:
        logger.error(f"An error occured while processing {file_path}: {e}")
        logger.error(traceback.format_exc())
        outputMsgBox("テキストファイルの内容をJSON形式に変換する処理でエラーが発生しました")
    except Exception as e:
        logger.error(f"An error occured while processing {file_path}: {e}")
        logger.error(traceback.format_exc())
        outputMsgBox("ファイル読み込みでエラーが発生しました")

    file_dir = os.path.dirname(file_path);
    file_name = os.path.basename(file_path);

    output_data = conv_data(data)

    if output_data is None or len(output_data) == 0:
        logger.error(f"データ変換に失敗しました")
        outputMsgBox("エラーが発生しました")
        return
    
    try:
        file_path = file_dir + '/decision_', file_name
        with open(file_path, 'w', encoding='utf-8') as file:
            for item in output_data:
                file.write(f"{item}/n")
    except Exception as e:
        logger.error(f"An error occured while processing {file_path}: {e}")
        logger.error(traceback.format_exc())
        outputMsgBox("ファイル書き込みでエラーは発生しました。")

def main():

    config = configparser.ConfigParser()
    config.read('config.ini')
    input_directory = config['DEFAULT']['InputDirectory']
    input_directory = os.path.abspath(os.path.join(input_directory, os.pardir))
    input_directory = input_directory + "\\response json"

    try:

        for file in os.listdir(input_directory):
            if file.startswith('response_json') and file.endswith('.txt'):
                file_path = os.path.join(input_directory, file)
                conv_response_to_decision(file_path, input_directory)
    except Exception as e:
        print(e)

if __name__== "__main__":
    main()