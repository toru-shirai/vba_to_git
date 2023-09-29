'ﾌｯﾀｰ開始行
Public currentFooterStartRows As Long
'総明細行数
Public allRowsCount As Long
'最終行
Public endRows As Long
'ﾒｲﾝｼｰﾄ名格納箇所
Public mainSheetName As String
'出力ファイル名
Public outputFileName As String
'出力部数
Public copies As String

'印刷種別:直接印刷時の「用紙サイズ」、「印刷の向き（縦・横）」
'ｻﾌﾞﾌｫｰﾏｯﾄ印刷ｻｲｽﾞ
Public paperSize As Collection
Public pageOrientation As Collection

'test conflict4