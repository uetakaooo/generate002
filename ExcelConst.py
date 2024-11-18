'''
定数管理のメタクラス
値の上書きを防止するため実装
'''


class ExcelConstMeta(type):
    _initialized = False

    def __setter__(self, name, value):
        if self._initialized:
            if name in self.__dict__:
                raise ValueError(f'{name}は定義済みです')
            else:
                raise AttributeError('定数を追加することは出来ません')
            super().__setter__(name, value)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._initialized = True

'''
定数管理クラス
'''
class ExcelConst(metaclass=ExcelConstMeta):
    # テンプレートファイル名
    TEMPLATE_FILE_NAME = 'decision_table_astaId_XXXX.xlsx'
    # テンプレートファイルを格納しているフォルダ名
    TEMPLATE_FOLDER_NAME = 'template'
    # 対象のシート名
    SHEET_NAME = 'API'

    # 名前付き範囲の名称
    AREA_NAME_CASEAREA = 'ケースエリア'
    AREA_NAME_API = 'APIテーブル'
    AREA_NAME_REQUEST = 'リクエストパラメータエリア'
    AREA_NAME_RESPONSE = 'レスポンスパラメータエリア'
    AREA_NAME_DECISON = '●付けエリア'
    AREA_NAME_LAST_ROW = '最終行'
    AREA_NAME_LAST_COL = '最終列'

    # テンプレートの行数
    DEFAULT_CASE_COL_NUM = 2
    # ●付けするエリアの最初の列
    INPUT_AREA_COL = 7
    # key名を入力する列
    INPUT_KEY_COL = 4
    # valueを入力する列
    INPUT_VALUE_COL = 6

    # IDやURL等が記載されている開始行
    TITLE_MIN_ROW = 2
    # IDやURL等が記載されている終了行
    TITLE_MAX_ROW = 4 