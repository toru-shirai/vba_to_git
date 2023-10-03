Option Explicit

Sub MainExec()
    
    Dim exeResult As Boolean
    Dim conInfo As ConstInfo
    Set conInfo = New ConstInfo
    
    '副作用利用でセット
    Call SetConst(conInfo)
    '共通処理呼び出し
    exeResult = Application.Run("'" & ThisWorkbook.Path & "\" & SETTING_CONST.EXE_COMMON_NAME & "'!MAIN.StartPrintPdf", ThisWorkbook, conInfo)
    Workbooks(SETTING_CONST.EXE_COMMON_NAME).Close
    
    If exeResult Then
        Err.Raise Number:=-1, Description:="ｴﾗｰ内容"
    End If

End Sub

'-------------------------------------------------------------------------------
' メソッド      SetConst
' 機能          必要な定数やMAIN帳票情報をクラスに変換
' 機能説明　　  必要な定数やMAIN帳票情報を共通関数へ引き渡すため、クラスに変換
'-------------------------------------------------------------------------------
Sub SetConst(ByRef conInfo As ConstInfo)
    
    '変更不要
    conInfo.formLink = "'" & ThisWorkbook.FullName & "'"
    conInfo.detailStartRow = CLng(ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_PARAM).Range(PARAM_HEADER_END).Value) + 1
    Dim a As String
    a = "test"
    
    '定数に修正が有る場合のみ対応
    
    conInfo.rangeCustomer = SETTING_CONST.RANGE_CUSTOMER
    conInfo.rengeArticle1 = SETTING_CONST.RANGE_ARTICLE_1
    conInfo.rangeArticle2 = SETTING_CONST.RANGE_ARTICLE_2
    conInfo.rangeInspectionQuantity = SETTING_CONST.RANGE_INSPECTION_QUANTITY
    conInfo.rangeLotNo = SETTING_CONST.RANGE_LOT_NO
    conInfo.rangeInspectionDate = SETTING_CONST.RANGE_INSPECTION_DATE
    conInfo.rangeCeramicBody = SETTING_CONST.RANGE_CERAMIC_BODY
    conInfo.rangeMetallization = SETTING_CONST.RANGE_METALLIZATION
    conInfo.rangeGoldPlating = SETTING_CONST.RANGE_GOLD_PLATING
    conInfo.rangeDimension = SETTING_CONST.RANGE_DIMENSION
    conInfo.rangePlattingThickness = SETTING_CONST.RANGE_PLATTING_THICKNESS
    conInfo.rangeVisual = SETTING_CONST.RANGE_VISUAL
    conInfo.rangeElectrical = SETTING_CONST.RANGE_ELECTRICAL
    conInfo.rangeDrawingNo = SETTING_CONST.RANGE_DRAWING_NO
    conInfo.rangeSpecNo = SETTING_CONST.RANGE_SPEC_NO
    conInfo.rangeNote1 = SETTING_CONST.RANGE_NOTE_1
    conInfo.rangeNote2 = SETTING_CONST.RANGE_NOTE_2
    conInfo.rangeNote3 = SETTING_CONST.RANGE_NOTE_3
    conInfo.rangeNote4 = SETTING_CONST.RANGE_NOTE_4
    conInfo.rangeInspectedBy = SETTING_CONST.RANGE_INSPECTED_BY
    conInfo.rangeApprovedBy = SETTING_CONST.RANGE_APPROVED_BY
    conInfo.colItem = SETTING_CONST.COL_ITEM
    conInfo.colSpc = SETTING_CONST.COL_SPC
    conInfo.colSpecL = SETTING_CONST.COL_SPEC_L
    conInfo.colSpecC = SETTING_CONST.COL_SPEC_C
    conInfo.colSpecR = SETTING_CONST.COL_SPEC_R
    conInfo.colUnit = SETTING_CONST.COL_UNIT
    conInfo.colMax = SETTING_CONST.COL_MAX
    conInfo.colMin = SETTING_CONST.COL_MIN
    conInfo.colAvg = SETTING_CONST.COL_AVG
    conInfo.colSd = SETTING_CONST.COL_SD
    conInfo.colCpk = SETTING_CONST.COL_CPK
    conInfo.colRn = SETTING_CONST.COL_RN
    conInfo.colResult = SETTING_CONST.COL_RESULT
    conInfo.colAvi = SETTING_CONST.COL_AVI
    conInfo.detailDirection = SETTING_CONST.DETAIL_DIRECTION
    conInfo.sheetNameMain = SETTING_CONST.SHEET_NAME_MAIN
    conInfo.sheetNameParam = SETTING_CONST.SHEET_NAME_PARAM
    conInfo.paramArticleByte1 = SETTING_CONST.PARAM_ARTICLE_BYTE1
    conInfo.paramArticleByte2 = SETTING_CONST.PARAM_ARTICLE_BYTE2
    conInfo.paramLengthLimit = SETTING_CONST.PARAM_LENGTH_LIMIT
    conInfo.paramRowHeight = SETTING_CONST.PARAM_ROW_HEIGHT
    conInfo.paramColumnLimit = SETTING_CONST.PARAM_COLUMN_LIMIT
    conInfo.paramRow = SETTING_CONST.PARAM_ROW
    conInfo.paramPaddingSize = SETTING_CONST.PARAM_PADDING_SIZE
    conInfo.paramAddHeight = SETTING_CONST.PARAM_ADD_HEIGHT
    conInfo.paramHeaderStart = SETTING_CONST.PARAM_HEADER_START
    conInfo.paramHeaderEnd = SETTING_CONST.PARAM_HEADER_END
    conInfo.paramFooterStart = SETTING_CONST.PARAM_FOOTER_START
    conInfo.paramFooterEnd = SETTING_CONST.PARAM_FOOTER_END
    conInfo.paramLinefeedSymbol = SETTING_CONST.PARAM_LINEFEED_SYMBOL
    conInfo.paramSubPath = SETTING_CONST.PARAM_SUB_PATH
    conInfo.paramCsvPath = SETTING_CONST.PARAM_CSV_PATH
    conInfo.paramEviPath = SETTING_CONST.PARAM_EVI_PATH
    conInfo.paramOutFilename = SETTING_CONST.PARAM_OUT_FILENAME
    conInfo.paramCsv1 = SETTING_CONST.PARAM_CSV1
    conInfo.paramCsv2 = SETTING_CONST.PARAM_CSV2
    conInfo.paramCsv3 = SETTING_CONST.PARAM_CSV3
    conInfo.paramCsv3Check = SETTING_CONST.PARAM_CSV3_CHECK
    conInfo.paramDataSheet1 = SETTING_CONST.PARAM_DATA_SHEET1
    conInfo.paramDataSheet2 = SETTING_CONST.PARAM_DATA_SHEET2
    conInfo.paramDataSheet3 = SETTING_CONST.PARAM_DATA_SHEET3
    conInfo.paramOutCopies = SETTING_CONST.PARAM_OUT_COPIES
    conInfo.paramTargetFormat1 = SETTING_CONST.PARAM_TARGET_FORMAT1
    conInfo.paramCustomer = SETTING_CONST.PARAM_CUSTOMER
    conInfo.targetMain = SETTING_CONST.TARGET_MAIN
    conInfo.paramTargetFormatRange = SETTING_CONST.PARAM_TARGET_FORMAT_RANGE
    conInfo.paperSize = SETTING_CONST.PAPER_SIZE
    conInfo.pageOrientation = SETTING_CONST.PAGE_ORIENTATION


End Sub


'-------------------------------------------------------------------------------
' メソッド      writeElement
' 機能          本帳票に値を書き込む
' 機能説明　　  CSVの値を本帳票に値を書き込む
'-------------------------------------------------------------------------------
Public Function writeElement(ByRef csvInfo As Variant, ByRef argInfo As Variant) As Boolean
    
    Dim i As Long
    Dim items As Variant
    Dim detailStartRow As Long
    Dim paramWs As Worksheet        'ﾊﾟﾗﾒｰﾀｼｰﾄObj
    Dim customerParam As String     '顧客名/Customer
    Dim articleCount1 As Long       '製品名区切り位置1
    Dim articleCount2 As Long       '製品名区切り位置2
    Dim articleStr1 As String       '製品名1
    Dim articleStr2 As String       '製品名2
    Dim searchLine As Long          '検索結果一時保存
    Dim searchLineFeedCount As Long '区切り文字位置
    Dim lineFeedStr As String       '改行文字
    Dim lineFeedArray() As String   '複数改行文字分割後
    Dim lfCount As Long             '改行文字数
    Dim kikakuType As String        '規格ﾀｲﾌﾟ
    Dim addRowCount As Long         '分類表記追加のカウント用
    Dim writingRow As Long          '記入行
    
    Dim outHeader As Cls_OutputHeader '出力項目(ﾍｯﾀﾞｰ)
    Dim outDetail As Cls_OutputDetail '出力項目(明細)
    Dim outFooter As Cls_OutputFooter '出力項目(ﾌｯﾀｰ)
    
'    On Error GoTo ErrorCatch
    
    '出力項目一時保存
    Set outHeader = New Cls_OutputHeader

    'ﾊﾟﾗﾒｰﾀｼｰﾄObj取得
    Set paramWs = ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_PARAM)

    '明細開始行計算
    detailStartRow = CLng(ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_PARAM).Range(PARAM_HEADER_END).Value) + 1

    'Customer
    'ﾊﾟﾗﾒｰﾀの顧客名/Customer参照
    customerParam = paramWs.Range(SETTING_CONST.PARAM_CUSTOMER).Value
    '1の場合、(20)取引先略式名
    If customerParam = "1" Then
        outHeader.Customer = csvInfo.commonInfo.FurimukesakiName
    '2の場合、(79)取引先名
    ElseIf customerParam = "2" Then
        outHeader.Customer = csvInfo.commonInfo.TorihikisakiSeishikiNameJyuchuzan
    End If

    'Article：(21)KC品名
    articleCount1 = paramWs.Range(SETTING_CONST.PARAM_ARTICLE_BYTE1).Value
    articleCount2 = paramWs.Range(SETTING_CONST.PARAM_ARTICLE_BYTE2).Value
    lineFeedStr = paramWs.Range(SETTING_CONST.PARAM_LINEFEED_SYMBOL).Value
    searchLineFeedCount = 0
    If Not IsEmpty(lineFeedStr) Then
        'lineFeedStrを配列へ分割
        lineFeedArray = OneStringSplit(lineFeedStr)
        For i = 0 To UBound(lineFeedArray)
            searchLine = InStrRev(csvInfo.commonInfo.KcHinmei, lineFeedArray(i))
            'TODO 【●条件確認中】 予想：改行可能文字がKC品名に存在する、かつヘッダーの客先品名文字長1より短い、かつ過去の検索結果より後に存在する場合
            If Not searchLine = 0 _
            And searchLineFeedCount < searchLine _
            And searchLine < articleCount1 Then
                searchLineFeedCount = searchLine
                Exit For
            End If
        Next i
    End If
    '文字に応じて分割表記
    'ヘッダーの客先品名文字長1以下
    If LenB(csvInfo.commonInfo.KcHinmei) <= articleCount1 Then
        articleStr1 = csvInfo.commonInfo.KcHinmei
    'ヘッダーの客先品名文字長1以上かつ区切り文字無し
    ElseIf IsEmpty(lineFeedStr) Then
        articleStr1 = LeftB(csvInfo.commonInfo.KcHinmei, articleCount1)
        articleStr2 = MidB$(csvInfo.commonInfo.KcHinmei, LenB(articleStr1) + 1, articleCount2)
        'Article (2)：(21)KC品名
        outHeader.Article2 = articleStr2
    'ヘッダーの客先品名文字長1以上かつ区切り文字有り TODO 【●条件確認中】 予想：改行可能文字がKC品名に存在する、かつヘッダーの客先品名文字長1より短い、かつ過去の検索結果より後に存在する場合
    ElseIf Not searchLineFeedCount = 0 Then
        articleStr1 = LeftB(csvInfo.commonInfo.KcHinmei, searchLineFeedCount)
        articleStr2 = MidB$(csvInfo.commonInfo.KcHinmei, LenB(articleStr1) + 1, articleCount2)
        'Article (2)：(21)KC品名 2行目
        outHeader.Article2 = articleStr2
    End If
    'Article (1):(21)KC品名
    outHeader.Article1 = articleStr1
    'Inspection Quantity：(18)出荷数
    outHeader.InspectionQuantity = csvInfo.commonInfo.ShukkaSu
    'Lot No.：(16)出荷ﾛｯﾄNo
    outHeader.LotNo = csvInfo.commonInfo.ShukkaLotNo
    'Inspection Date：(14)出荷日
    outHeader.InspectionDate = "=UPPER(TEXT(" & csvInfo.commonInfo.ShukkaDate & ",""mmm dd,yyyy""))"

    Call WriteHeader(outHeader)
    
    'csv行数と実記入行の差分
    addRowCount = detailStartRow

    '2. 明細部分へ値をｾｯﾄ
    For i = 0 To csvInfo.DetailInfoList.Count - 1
    
        '出力項目クラス
        Set outDetail = New Cls_OutputDetail
        '対象CSV項目
        Set items = csvInfo.DetailInfoList(i + 1)
        '記入行数
        writingRow = i + detailStartRow
        
        '[1:分類]の場合は分類を記載
        If items.itemKbn = "1" Then
            outDetail.ItemName = items.ItemName

        '[2:測定項目名]の場合は測定項目名と結果を記載
        ElseIf items.itemKbn = "2" Then
            If Left(paramWs.Range(SETTING_CONST.PARAM_LENGTH_LIMIT), 1) = "1" Then
                'ITEM(測定項目名)：(104)測定項目名
                outDetail.ItemName = items.ItemName
            ElseIf Left(paramWs.Range(SETTING_CONST.PARAM_LENGTH_LIMIT), 1) = "2" Then
                '改行回数に応じて変更 ※LenBとすると、高さ計算結果が倍になるため注意
                lfCount = Len(items.ItemName) - Len(Replace(items.ItemName, vbLf, ""))
                outDetail.ItemNameLfCnt = lfCount
                'ITEM(測定項目名)：(104)測定項目名
                outDetail.ItemName = items.ItemName
            End If
            'SPC：(123)SPC
            outDetail.SpcValue = items.Spc1

            kikakuType = Left(items.KikakuHantei, 1)
            '規格ﾀｲﾌﾟに応じて表記を変更
            Select Case kikakuType
            Case "0"
                'SPEC(左)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecL = items.Usl
                'SPEC(中)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecC = "-"
                'SPEC(右)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecR = items.Lsl
            Case "1"
                'SPEC(左)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecL = items.Usl
                'SPEC(中)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecC = "-"
                'SPEC(右)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecR = "MAX"
            Case "3"
                'SPEC(左)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecL = items.Lsl
                'SPEC(中)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecC = "-"
                'SPEC(右)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecR = "MIN"
            Case "4"
                'SPEC(左)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecL = "(" & items.Usl & ")"
                'SPEC(中)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecC = "-"
                'SPEC(右)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecL = "(" & items.Lsl & ")"
            Case "5"
                'SPEC(左)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecL = items.Usl
                'SPEC(中)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecC = "-"
                'SPEC(右)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecR = items.Lsl
            Case "6"
                'SPEC(左)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecL = ""
                'SPEC(中)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecC = items.Ondo & "℃ - " & items.Jikan & items.JikanTani & " OVEN"
                'SPEC(右)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecR = ""
            Case "8"
                'SPEC(左)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecL = ""
                'SPEC(中)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecC = "-"
                'SPEC(右)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecR = ""
            Case "9"
                'SPEC(左)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecL = items.Usl
                'SPEC(中)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecC = "-"
                'SPEC(右)：ｼｰﾄ「機能定義書(PXJDO301)」　「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
                outDetail.SpecR = items.Lsl
            End Select

            'ﾃﾞｰﾀ表示方法に応じて表記を変更


            '規格ﾀｲﾌﾟに応じて表記を変更
            Select Case kikakuType
            '「0:MAX/MIN/AVG」の場合：「MAX/MIN/AVG」の項目がある場合は、値をｾｯﾄする。「標準偏差/Cpk」の項目がある場合は、中央揃えに設定して「-」をｾｯﾄする。
            Case "6"
                Call Kikaku6Or8(outDetail)
            Case "8"
                Call Kikaku6Or8(outDetail)
            Case Else
                'UNIT：(107)単位
                outDetail.Unit = items.TaniKensa
                'MAX：(108)最大値
                outDetail.MaxN = items.MaxN
                outDetail.MaxCenterd = False
                'MIN：(109)最小値
                outDetail.MinN = items.MinN
                outDetail.MinCenterd = False
                'AVG：(110)平均値
                outDetail.Ave = items.Ave
                outDetail.AveCenterd = False
                'SD：(111)標準偏差
                outDetail.Sigma = items.Sigma
                outDetail.SigmaCenterd = False
                'Cpk：(112)cpk
                outDetail.Cpk = items.Cpk
                outDetail.CpkCenterd = False
            End Select


            'r/n：(113)抜取数(不合格数)+"/"+(114)抜取数(母数)
            outDetail.RN = items.NukitoriSuFugoukakuSu & "/" & items.NukitoriSuBosu
            'RESULT：(115)判定
            outDetail.result = ResultString(items.Hantei)
            '100% AVI applied：(128)画像検査
            outDetail.AviApplied = AviAppliedString(items.GazoKensa)
        Else
            addRowCount = addRowCount + 1
        End If
    
        Call WriteDetail(outDetail, writingRow)
    
    Next i
    
    '最終行を保存
    argInfo.endRows = i + addRowCount - 1

    '3. ﾌｯﾀｰ部分へ値をｾｯﾄ
    Set outFooter = New Cls_OutputFooter
    
    'MATERIAL Ceramic Body：(33)CERAMIC
    outFooter.Ceramic = csvInfo.commonInfo.Ceramic
    'MATERIAL Metallization：(34)METALIZE
    outFooter.Metalize = csvInfo.commonInfo.Metalize
    'MATERIAL Gold Plating：(36)PLATING
    outFooter.Plating = csvInfo.commonInfo.Plating
    '<SAMPLING PLAN> Dimension：(62)SAMPLING_PLAN2
    outFooter.SamplingPlan2 = csvInfo.commonInfo.SamplingPlan2
    '<SAMPLING PLAN> Platting Thickness：(63)SAMPLING_PLAN3
    outFooter.SamplingPlan3 = csvInfo.commonInfo.SamplingPlan3
    '<SAMPLING PLAN> VISUAL：(64)SAMPLING_PLAN4
    outFooter.SamplingPlan4 = csvInfo.commonInfo.SamplingPlan4
    '<SAMPLING PLAN> Electrical(Open/Short)：(65)SAMPLING_PLAN5
    outFooter.SamplingPlan5 = csvInfo.commonInfo.SamplingPlan5
    '<SAMPLING PLAN> Drawing No：(22)KC図番
    outFooter.KcZuban = csvInfo.commonInfo.KcZubanJyuchuzan
    '<SAMPLING PLAN> Spec No：(24)SPEC
    outFooter.Spec = csvInfo.commonInfo.SpecJuchuzan
    'NOTE(1)：(37)備考1
    outFooter.Biko1 = csvInfo.commonInfo.Biko1
    'NOTE(2)：(38)備考2
    outFooter.Biko2 = csvInfo.commonInfo.Biko2
    'NOTE(3)：(39)備考3
    outFooter.Biko3 = csvInfo.commonInfo.Biko3
    'NOTE(4)：(40)備考4
    outFooter.Biko4 = csvInfo.commonInfo.Biko4
    'Inspected by：(53)作成者名
    outFooter.InspectedBy = csvInfo.commonInfo.KensahyoHakkoshaName
    'Approved by：(54)承認者名
    outFooter.ApprovedBy = csvInfo.commonInfo.KensahyoShoninshaName
    
    Call WriteFooter(outFooter)
    
    Exit Function

ErrorCatch:
    writeElement = True

End Function

'-------------------------------------------------------------------------------
' メソッド      OneStringSplit
' 機能          文字列を分割する
' 機能説明　　  文字列を1文字ずつ分割して配列を返す
'-------------------------------------------------------------------------------
Function OneStringSplit(ByVal target As String)

    Dim i As Long
    Dim returnArray() As String
    ReDim returnArray(Len(target) - 1)
    For i = 0 To Len(target) - 1
        returnArray(i) = Left(target, 1)
        target = Right(target, Len(target) - 1)
    Next i

    '戻り値
    OneStringSplit = returnArray

End Function

'-------------------------------------------------------------------------------
' メソッド      CellEdit
' 機能          セルの編集
' 機能説明　　  対象Bookのシートかつ対象セルの値・書式の編集
'-------------------------------------------------------------------------------
Sub CellEdit(pasteData As String, pasteCell As String, Optional centered As Boolean = False, Optional kaigyo As Long = 0)

    Dim defaultHight As Long
    Dim plusHight As Long
    Dim ws As Worksheet

    defaultHight = ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_PARAM).Range(SETTING_CONST.PARAM_ROW).Value
    plusHight = ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_PARAM).Range(SETTING_CONST.PARAM_ADD_HEIGHT).Value
    Set ws = ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_MAIN)

    'セルへ記載
    ws.Range(pasteCell).Value = pasteData
    '書式設定を中央揃えに変更
    If (centered) Then
        ws.Range(pasteCell).HorizontalAlignment = xlCenter
    End If
    '値の改行回数が1以上の時
    If kaigyo > 0 Then
        ws.Range(pasteCell).RowHeight = defaultHight + (plusHight * kaigyo)
    End If
End Sub

'-------------------------------------------------------------------------------
' メソッド      ResultString
' 機能          結果の数値を文字列に変換する
' 機能説明      カスタマイズを楽にするための関数
'-------------------------------------------------------------------------------
Function ResultString(ByVal target) As String

    Select Case target
    Case "0"
        '戻り値
        ResultString = "REJECT"
    Case "1"
        '戻り値
        ResultString = "ACCEPT"
    End Select
End Function


'-------------------------------------------------------------------------------
' メソッド      AviApplied
' 機能          結果の数値を文字列に変換する
' 機能説明      カスタマイズを楽にするための関数
'-------------------------------------------------------------------------------
Function AviAppliedString(ByVal target) As String

    Select Case target
    Case "0"
        '戻り値
        AviAppliedString = ""
    Case "1"
        '戻り値
        AviAppliedString = "○"
    End Select
End Function

'-------------------------------------------------------------------------------
' メソッド      WriteHeader
' 機能          Headerを記述する
' 機能説明      Headerに値を与える
'-------------------------------------------------------------------------------
Sub WriteHeader(outHeader As Cls_OutputHeader)

    '取引先
    Call CellEdit(outHeader.Customer, SETTING_CONST.RANGE_CUSTOMER)
    'KC品名
    Call CellEdit(outHeader.Article1, SETTING_CONST.RANGE_ARTICLE_1)
    'KC品名 2行目
    Call CellEdit(outHeader.Article2, SETTING_CONST.RANGE_ARTICLE_2)
    '出荷数
    Call CellEdit(outHeader.InspectionQuantity, SETTING_CONST.RANGE_INSPECTION_QUANTITY)
    '出荷ﾛｯﾄNo
    Call CellEdit(outHeader.LotNo, SETTING_CONST.RANGE_LOT_NO)
    '出荷日
    Call CellEdit(outHeader.InspectionDate, SETTING_CONST.RANGE_INSPECTION_DATE)

End Sub

'-------------------------------------------------------------------------------
' メソッド      Kikaku6Or8
' 機能          規格ﾀｲﾌﾟが6,8の場合の処理を関数化
' 機能説明      規格ﾀｲﾌﾟが6,8の場合にUNIT～CPKまでの箇所が"-"になるため関数化
'-------------------------------------------------------------------------------
Sub Kikaku6Or8(ByRef outDetail As Cls_OutputDetail)
    
    '(116)規格ﾀｲﾌﾟが「6:耐熱性、8:外観」の場合は"-"
    'UNIT：(107)単位
    outDetail.Unit = "-"
    'MAX：(108)最大値
    outDetail.MaxN = "-"
    outDetail.MaxCenterd = True
    'MIN：(109)最小値
    outDetail.MinN = "-"
    outDetail.MinCenterd = True
    'AVG：(110)平均値
    outDetail.Ave = "-"
    outDetail.AveCenterd = True
    'SD：(111)標準偏差
    outDetail.Sigma = "-"
    outDetail.SigmaCenterd = True
    'Cpk：(112)cpk
    outDetail.Cpk = "-"
    outDetail.CpkCenterd = True

End Sub

'-------------------------------------------------------------------------------
' メソッド      WriteDetail
' 機能          明細部を記述する
' 機能説明      明細部に値を与える
'-------------------------------------------------------------------------------
Sub WriteDetail(outDetail As Cls_OutputDetail, writingRow As Long)

    'ITEM(測定項目名)：(20)取引先略式名 / (79)取引先名
    Call CellEdit(outDetail.ItemName, COL_ITEM & CStr(writingRow), False, outDetail.ItemNameLfCnt)
    'SPC：(123)SPC
    Call CellEdit(outDetail.SpcValue, COL_SPC & CStr(writingRow))
    '「D.SPECの表現方法（SPEC(左)(中)(右)の場合)」参照
    'SPEC：ｼｰﾄ「機能定義書(PXJDO301)」
    Call CellEdit(outDetail.SpecL, COL_SPEC_L & CStr(writingRow))
    Call CellEdit(outDetail.SpecC, COL_SPEC_C & CStr(writingRow))
    Call CellEdit(outDetail.SpecR, COL_SPEC_R & CStr(writingRow))
    'UNIT：(107)単位
    Call CellEdit(outDetail.Unit, COL_UNIT & CStr(writingRow))
    'MAX：(108)最大値
    Call CellEdit(outDetail.MaxN, COL_MAX & CStr(writingRow), outDetail.MaxCenterd)
    'MIN：(109)最小値
    Call CellEdit(outDetail.MinN, COL_MIN & CStr(writingRow), outDetail.MinCenterd)
    'AVG：(110)平均値
    Call CellEdit(outDetail.Ave, COL_AVG & CStr(writingRow), outDetail.AveCenterd)
    'SD：(111)標準偏差
    Call CellEdit(outDetail.Sigma, COL_SD & CStr(writingRow), outDetail.SigmaCenterd)
    'Cpk：(112)cpk
    Call CellEdit(outDetail.Cpk, COL_CPK & CStr(writingRow), outDetail.CpkCenterd)
    'r/n：(113)抜取数(不合格数)+"/"+(114)抜取数(母数)
    Call CellEdit(outDetail.RN, COL_RN & CStr(writingRow))
    'RESULT：(115)判定
    Call CellEdit(outDetail.result, COL_RESULT & CStr(writingRow))
    '100% AVI applied：(128)画像検査
    Call CellEdit(outDetail.AviApplied, COL_AVI & CStr(writingRow))

End Sub


'-------------------------------------------------------------------------------
' メソッド      WriteFooter
' 機能          Footerを記述する
' 機能説明      Footerに値を与える
'-------------------------------------------------------------------------------
Sub WriteFooter(outFooter As Cls_OutputFooter)

    'MATERIAL Ceramic Body：(33)CERAMIC
    Call CellEdit(outFooter.Ceramic, SETTING_CONST.RANGE_CERAMIC_BODY)
    'MATERIAL Metallization：(34)METALIZE
    Call CellEdit(outFooter.Metalize, SETTING_CONST.RANGE_METALLIZATION)
    'MATERIAL Gold Plating：(36)PLATING
    Call CellEdit(outFooter.Plating, SETTING_CONST.RANGE_GOLD_PLATING)
    '<SAMPLING PLAN> Dimension：(62)SAMPLING_PLAN2
    Call CellEdit(outFooter.SamplingPlan2, SETTING_CONST.RANGE_DIMENSION)
    '<SAMPLING PLAN> Platting Thickness：(63)SAMPLING_PLAN3
    Call CellEdit(outFooter.SamplingPlan3, SETTING_CONST.RANGE_PLATTING_THICKNESS)
    '<SAMPLING PLAN> VISUAL：(64)SAMPLING_PLAN4
    Call CellEdit(outFooter.SamplingPlan4, SETTING_CONST.RANGE_VISUAL)
    '<SAMPLING PLAN> Electrical(Open/Short)：(65)SAMPLING_PLAN5
    Call CellEdit(outFooter.SamplingPlan5, SETTING_CONST.RANGE_ELECTRICAL)
    '<SAMPLING PLAN> Drawing No：(22)KC図番
    Call CellEdit(outFooter.KcZuban, SETTING_CONST.RANGE_DRAWING_NO)
    '<SAMPLING PLAN> Spec No：(24)SPEC
    Call CellEdit(outFooter.Spec, SETTING_CONST.RANGE_SPEC_NO)
    'NOTE(1)：(37)備考1
    Call CellEdit(outFooter.Biko1, SETTING_CONST.RANGE_NOTE_1)
    'NOTE(2)：(38)備考2
    Call CellEdit(outFooter.Biko2, SETTING_CONST.RANGE_NOTE_2)
    'NOTE(3)：(39)備考3
    Call CellEdit(outFooter.Biko3, SETTING_CONST.RANGE_NOTE_3)
    'NOTE(4)：(40)備考4
    Call CellEdit(outFooter.Biko4, SETTING_CONST.RANGE_NOTE_4)
    'Inspected by：(53)作成者名
    Call CellEdit(outFooter.InspectedBy, SETTING_CONST.RANGE_INSPECTED_BY)
    'Approved by：(54)承認者名
    Call CellEdit(outFooter.ApprovedBy, SETTING_CONST.RANGE_APPROVED_BY)

End Sub

'-------------------------------------------------------------------------------
' メソッド      MainSheetNameCreate
' 機能          各帳票のﾒｲﾝｼｰﾄ名を取得
' 機能説明　　  帳票名が各帳票毎に違うため、カスタマイズを楽にするための関数
'-------------------------------------------------------------------------------
Function MainSheetNameCreate(csvInfo)
    
    Dim result(2)
    
    On Error GoTo ErrorCatch
    
    '修正不要
    result(0) = "False"
    
    'ｼｰﾄ名：各帳票ごとに修正予定
    result(1) = csvInfo.commonInfo.KcZubanJyuchuzan
    MainSheetNameCreate = result
    
    Exit Function
    
ErrorCatch:
    
    result(0) = "True"
    
    MainSheetNameCreate = result
    
End Function



