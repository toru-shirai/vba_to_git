'���sBook�̃����N
Public Const EXE_COMMON_NAME As String = "�����\_����ϸ�.xlsm"

'ͯ�ް
Public Const HEADER_RANGE_CUSTOMER = "D2"
Public Const HEADER_RANGE_ARTICLE_1 = "D3"
Public Const HEADER_RANGE_INSPECTION_QUANTITY = "D4"
Public Const HEADER_RANGE_LOT_NO = "D5"
Public Const HEADER_RANGE_INSPECTION_DATE = "D6"
Public Const HEADER_RANGE_DRAWING_NO = "D7"
Public Const HEADER_RANGE_TAPE_LOT = "D8"
Public Const HEADER_RANGE_7_DIGIT_LOT = "D9"

'̯��
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

'���ו�
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

'���ׂ̌���
Public Const DETAIL_DIRECTION As Boolean = True
'Ҳݼ�Ė�
Public Const SHEET_NAME_MAIN = "MAIN"
'���Ұ���Ė�
Public Const SHEET_NAME_PARAM = "���Ұ����"
'ͯ�ް�̋q��i�������޲Đ�1
Public Const PARAM_ARTICLE_BYTE1 = "C2"
'ͯ�ް�̋q��i�������޲Đ�2
Public Const PARAM_ARTICLE_BYTE2 = "C3"
'���������E���̍l��
Public Const PARAM_LENGTH_LIMIT = "C4"
'���ׂƂ��Ďg�p�\�ȍ���(pt)
Public Const PARAM_ROW_HEIGHT = "C5"
'1�߰�ނɓ����
Public Const PARAM_COLUMN_LIMIT = "C6"
'1�s���̍���(�W��)(pt)
Public Const PARAM_ROW = "C7"
 '�l�߻���(pt)
Public Const PARAM_PADDING_SIZE = "C8"
'2�s�ȍ~�̒ǉ�����(pt)
Public Const PARAM_ADD_HEIGHT = "C9"
'ͯ�ް�̊J�n�s
Public Const PARAM_HEADER_START = "C10"
'ͯ�ް�̏I���s
Public Const PARAM_HEADER_END = "C11"
'̯���̊J�n�s
Public Const PARAM_FOOTER_START = "C12"
'̯���̏I���s
Public Const PARAM_FOOTER_END = "C13"
'���s�\�L��(��R�����ƒx���Ȃ�)
Public Const PARAM_LINEFEED_SYMBOL = "C14"
'���̫�ϯ�ϸەۑ��ꏊ
Public Const PARAM_SUB_PATH = "C16"
'CSV̧�ٕۑ��ꏊ
Public Const PARAM_CSV_PATH = "C15"
'�����ݽ�ۑ��ꏊ
Public Const PARAM_EVI_PATH = "C17"
'�o��̧�ٖ�
Public Const PARAM_OUT_FILENAME = "C18"
'CSV̧��1
Public Const PARAM_CSV1 = "C19"
'CSV̧��2(����)
Public Const PARAM_CSV2 = "C20"
'CSV̧��3(����ۯ�)
Public Const PARAM_CSV3 = "C21"
'CSV̧��3(����ۯ�)��������
Public Const PARAM_CSV3_CHECK = "C22"
'�ް����1
Public Const PARAM_DATA_SHEET1 = "C23"
'�ް����2(����)
Public Const PARAM_DATA_SHEET2 = "C24"
'�ް����3(����ۯ�)
Public Const PARAM_DATA_SHEET3 = "C25"
'�o�͕���
Public Const PARAM_OUT_COPIES = "C26"
'�Ǎ��Ώ�̧��(����)
'1�Ԗڏo��
Public Const PARAM_TARGET_FORMAT1 = "C29"
'�ڋq��/Customer
Public Const PARAM_CUSTOMER = "C40"

'�����\Ҳݼ��
Public Const TARGET_MAIN = "�����\"
'�Ǎ��Ώ�̧�ق͈̔�
Public Const PARAM_TARGET_FORMAT_RANGE = "C29:C37"