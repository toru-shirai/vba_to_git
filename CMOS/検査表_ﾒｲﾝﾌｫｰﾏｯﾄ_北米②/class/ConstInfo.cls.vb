'自身のBookLink
Public formLink As String

'明細開始行
Public detailStartRow As Long
Public rangeCustomer As String
Public rengeArticle1 As String
Public rangeArticle2 As String
Public rangeInspectionQuantity As String
Public rangeLotNo As String
Public rangeInspectionDate As String
'ﾌｯﾀｰ
Public rangeCeramicBody As String
Public rangeMetallization As String
Public rangeGoldPlating As String
Public rangeDimension As String
Public rangePlattingThickness As String
Public rangeVisual As String
Public rangeElectrical As String
Public rangeDrawingNo As String
Public rangeSpecNo As String
Public rangeNote1 As String
Public rangeNote2 As String
Public rangeNote3 As String
Public rangeNote4 As String
Public rangeInspectedBy As String
Public rangeApprovedBy As String

'明細部
Public colItem As String
Public colSpc As String
Public colSpecL As String
Public colSpecC As String
Public colSpecR As String
Public colUnit As String
Public colMax As String
Public colMin As String
Public colAvg As String
Public colSd As String
Public colCpk As String
Public colRn As String
Public colResult As String
Public colAvi As String
   
'明細の向き
Public detailDirection As Boolean
'ﾒｲﾝｼｰﾄ名
Public sheetNameMain As String
'ﾊﾟﾗﾒｰﾀｼｰﾄ名
Public sheetNameParam As String
'ﾍｯﾀﾞｰの客先品名文字ﾊﾞｲﾄ数1
Public paramArticleByte1 As String
'ﾍｯﾀﾞｰの客先品名文字ﾊﾞｲﾄ数2
Public paramArticleByte2 As String
'文字数限界時の考え
Public paramLengthLimit As String
'明細として使用可能な高さ(pt)
Public paramRowHeight As String
'1ﾍﾟｰｼﾞに入る列数
Public paramColumnLimit As String
'1行時の高さ(標準)(pt)
Public paramRow As String
'詰めｻｲｽﾞ(pt)
Public paramPaddingSize As String
'2行以降の追加高さ(pt)
Public paramAddHeight As String
'ﾍｯﾀﾞｰの開始行
Public paramHeaderStart As String
'ﾍｯﾀﾞｰの終了行
Public paramHeaderEnd As String
'ﾌｯﾀｰの開始行
Public paramFooterStart As String
'ﾌｯﾀｰの終了行
Public paramFooterEnd As String
'改行可能記号(沢山書くと遅くなる)
Public paramLinefeedSymbol As String
'ｻﾌﾞﾌｫｰﾏｯﾄﾏｸﾛ保存場所
Public paramSubPath As String
'CSVﾌｧｲﾙ保存場所
Public paramCsvPath As String
'ｴﾋﾞﾃﾞﾝｽ保存場所
Public paramEviPath As String
'出力ﾌｧｲﾙ名
Public paramOutFilename As String
'CSVﾌｧｲﾙ1
Public paramCsv1 As String
'CSVﾌｧｲﾙ2(明細)
Public paramCsv2 As String
'CSVﾌｧｲﾙ3(複数ﾛｯﾄ)
Public paramCsv3 As String
'CSVﾌｧｲﾙ3(複数ﾛｯﾄ)存在ﾁｪｯｸ
Public paramCsv3Check As String
'ﾃﾞｰﾀｼｰﾄ1
Public paramDataSheet1 As String
'ﾃﾞｰﾀｼｰﾄ2(明細)
Public paramDataSheet2 As String
'ﾃﾞｰﾀｼｰﾄ3(複数ﾛｯﾄ)
Public paramDataSheet3 As String
'出力部数
Public paramOutCopies As String
'読込対象ﾌｧｲﾙ(共通)
'1番目出力
Public paramTargetFormat1 As String
'顧客名/Customer
Public paramCustomer As String
   
'
Public targetMain As String
'特例的に読込対象ﾌｧｲﾙの範囲
Public paramTargetFormatRange As String
   
'印刷時用紙サイズ
Public paperSize As Long
'印刷方向
Public pageOrientation As Long

