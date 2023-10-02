'実行Bookのリンク
Public Const EXE_COMMON_NAME As String = "検査表_共通ﾏｸﾛ.xlsm"

'ﾍｯﾀﾞｰ
Public Const HEADER_RANGE_CUSTOMER = "D2"
Public Const HEADER_RANGE_ARTICLE_1 = "D3"
Public Const HEADER_RANGE_INSPECTION_QUANTITY = "D4"
Public Const HEADER_RANGE_LOT_NO = "D5"
Public Const HEADER_RANGE_INSPECTION_DATE = "D6"
Public Const HEADER_RANGE_DRAWING_NO = "D7"
Public Const HEADER_RANGE_TAPE_LOT = "D8"
Public Const HEADER_RANGE_7_DIGIT_LOT = "D9"

'ﾌｯﾀｰ
Public Const RANGE_CERAMIC_BODY = "D10015"
Public Const RANGE_METALLIZATION = "D10016"
Public Const RANGE_GOLD_PLATING = "D10017"
Public Const RANGE_DIMENSION = "H10015"
Public Const RANGE_PLATTING_THICKNESS = "H10016"
Public Const RANGE_VISUAL = "H10017"
Public Const RANGE_ELECTRICAL = "H10018"
Public Const RANGE_SPEC_NO = "L10015"

Public Const RANGE_INSPECTED_BY = "Q10015"
Public Const RANGE_APPROVED_BY = "Q10017"

'明細部
Public Const COL_ITEM = "B"
Public Const COL_SPC = "C"
Public Const COL_USL = "D"
Public Const COL_LSL = "E"
Public Const COL_UNIT = "F"
Public Const COL_NUM_SOKUTEICHI_START = 7
Public Const COL_NUM_SOKUTEICHI_END = 70
Public Const COL_MAX = "BS"
Public Const COL_MIN = "BT"
Public Const COL_AVG = "BU"
Public Const COL_SD = "BV"
Public Const COL_CPK = "BW"
Public Const COL_RESULT = "BX"
Public Const COL_AVI = "BY"

Public Const SOKUTEICHI_MAX = 64


'----------一時作成定数------------------
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

'
Public Const TARGET_MAIN = "検査表"
'特例的に読込対象ﾌｧｲﾙの範囲
Public Const PARAM_TARGET_FORMAT_RANGE = "C29:C37"



'暫時的措置
'印刷時用紙サイズ
Public Const PAPER_SIZE = xlPaperB5
'暫時的措置
'印刷方向
Public Const PAGE_ORIENTATION = xlPortrait



