'実行Bookのリンク
Public Const EXE_COMMON_NAME As String = "検査表_共通ﾏｸﾛ.xlsm"

'ﾍｯﾀﾞｰ
Public Const RANGE_CUSTOMER = "M1"
Public Const RANGE_ARTICLE_1 = "M2"
Public Const RANGE_ARTICLE_2 = "N3"
Public Const RANGE_INSPECTION_QUANTITY = "M3"
Public Const RANGE_LOT_NO = "M4"
Public Const RANGE_INSPECTION_DATE = "M5"

'ﾌｯﾀｰ
Public Const RANGE_CERAMIC_BODY = "C10015"
Public Const RANGE_METALLIZATION = "C10016"
Public Const RANGE_GOLD_PLATING = "C10017"
Public Const RANGE_DIMENSION = "K10015"
Public Const RANGE_PLATTING_THICKNESS = "K10016"
Public Const RANGE_VISUAL = "K10017"
Public Const RANGE_ELECTRICAL = "K10018"
Public Const RANGE_DRAWING_NO = "N10014"
Public Const RANGE_SPEC_NO = "N10015"
Public Const RANGE_NOTE_1 = "B10020"
Public Const RANGE_NOTE_2 = "B10021"
Public Const RANGE_NOTE_3 = "B10022"
Public Const RANGE_NOTE_4 = "B10023"
Public Const RANGE_INSPECTED_BY = "M10020"
Public Const RANGE_APPROVED_BY = "M10023"

'明細部
Public Const COL_ITEM = "B"
Public Const COL_SPC = "C"
Public Const COL_SPEC_L = "D"
Public Const COL_SPEC_C = "E"
Public Const COL_SPEC_R = "F"
Public Const COL_UNIT = "G"
Public Const COL_MAX = "H"
Public Const COL_MIN = "I"
Public Const COL_AVG = "J"
Public Const COL_SD = "K"
Public Const COL_CPK = "L"
Public Const COL_RN = "M"
Public Const COL_RESULT = "N"
Public Const COL_AVI = "O"

'明細の向き
Public Const DETAIL_DIRECTION As Boolean = True
'ﾒｲﾝｼｰﾄ名
Public Const SHEET_NAME_MAIN = "MAIN"
'ﾊﾟﾗﾒｰﾀｼｰﾄ名
Public Const SHEET_NAME_PARAM = "ﾊﾟﾗﾒｰﾀｼｰﾄ"
'ﾍｯﾀﾞｰの客先品名文字ﾊﾞｲﾄ数1
Public Const PARAM_ARTICLE_BYTE1 = "C2"
'ﾍｯﾀﾞｰの客先品名文字ﾊﾞｲﾄ数2
Public Const PARAM_ARTICLE_BYTE2 = "C3"
'文字数限界時の考え
Public Const PARAM_LENGTH_LIMIT = "C4"
'明細として使用可能な高さ(pt)
Public Const PARAM_ROW_HEIGHT = "C5"
'1ﾍﾟｰｼﾞに入る列数
Public Const PARAM_COLUMN_LIMIT = "C6"
'1行時の高さ(標準)(pt)
Public Const PARAM_ROW = "C7"
 '詰めｻｲｽﾞ(pt)
Public Const PARAM_PADDING_SIZE = "C8"
'2行以降の追加高さ(pt)
Public Const PARAM_ADD_HEIGHT = "C9"
'ﾍｯﾀﾞｰの開始行
Public Const PARAM_HEADER_START = "C10"
'ﾍｯﾀﾞｰの終了行
Public Const PARAM_HEADER_END = "C11"
'ﾌｯﾀｰの開始行
Public Const PARAM_FOOTER_START = "C12"
'ﾌｯﾀｰの終了行
Public Const PARAM_FOOTER_END = "C13"
'改行可能記号(沢山書くと遅くなる)
Public Const PARAM_LINEFEED_SYMBOL = "C14"
'ｻﾌﾞﾌｫｰﾏｯﾄﾏｸﾛ保存場所
Public Const PARAM_SUB_PATH = "C16"
'CSVﾌｧｲﾙ保存場所
Public Const PARAM_CSV_PATH = "C15"
'ｴﾋﾞﾃﾞﾝｽ保存場所
Public Const PARAM_EVI_PATH = "C17"
'出力ﾌｧｲﾙ名
Public Const PARAM_OUT_FILENAME = "C18"
'CSVﾌｧｲﾙ1
Public Const PARAM_CSV1 = "C19"
'CSVﾌｧｲﾙ2(明細)
Public Const PARAM_CSV2 = "C20"
'CSVﾌｧｲﾙ3(複数ﾛｯﾄ)
Public Const PARAM_CSV3 = "C21"
'CSVﾌｧｲﾙ3(複数ﾛｯﾄ)存在ﾁｪｯｸ
Public Const PARAM_CSV3_CHECK = "C22"
'ﾃﾞｰﾀｼｰﾄ1
Public Const PARAM_DATA_SHEET1 = "C23"
'ﾃﾞｰﾀｼｰﾄ2(明細)
Public Const PARAM_DATA_SHEET2 = "C24"
'ﾃﾞｰﾀｼｰﾄ3(複数ﾛｯﾄ)
Public Const PARAM_DATA_SHEET3 = "C25"
'出力部数
Public Const PARAM_OUT_COPIES = "C26"
'読込対象ﾌｧｲﾙ(共通)
'1番目出力
Public Const PARAM_TARGET_FORMAT1 = "C29"
'顧客名/Customer
Public Const PARAM_CUSTOMER = "C40"

Public Const TARGET_MAIN = "検査表"
'読込対象ﾌｧｲﾙの範囲
Public Const PARAM_TARGET_FORMAT_RANGE = "C29:C37"

'印刷時用紙サイズ
Public Const PAPER_SIZE = xlPaperB5
'暫時的措置
'印刷方向
Public Const PAGE_ORIENTATION = xlPortrait

