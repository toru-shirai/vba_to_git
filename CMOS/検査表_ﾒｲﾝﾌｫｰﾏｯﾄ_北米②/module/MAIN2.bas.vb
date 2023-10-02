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

Sub SetConst(ByRef conInfo As ConstInfo)
    
    '変更不要
    conInfo.formLink = "'" & ThisWorkbook.FullName & "'"
    conInfo.detailStartRow = CLng(ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_PARAM).Range(PARAM_HEADER_END).Value) + 1
    
    '定数に修正が有る場合のみ対応
    
    conInfo.rangeCustomer = SETTING_CONST.HEADER_RANGE_CUSTOMER
    conInfo.rengeArticle1 = SETTING_CONST.HEADER_RANGE_ARTICLE_1
    conInfo.rangeInspectionQuantity = SETTING_CONST.HEADER_RANGE_INSPECTION_QUANTITY
    conInfo.rangeLotNo = SETTING_CONST.HEADER_RANGE_LOT_NO
    conInfo.rangeInspectionDate = SETTING_CONST.HEADER_RANGE_INSPECTION_DATE
    conInfo.rangeCeramicBody = SETTING_CONST.RANGE_CERAMIC_BODY
    conInfo.rangeMetallization = SETTING_CONST.RANGE_METALLIZATION
    conInfo.rangeGoldPlating = SETTING_CONST.RANGE_GOLD_PLATING
    conInfo.rangeDimension = SETTING_CONST.RANGE_DIMENSION
    conInfo.rangePlattingThickness = SETTING_CONST.RANGE_PLATTING_THICKNESS
    conInfo.rangeVisual = SETTING_CONST.RANGE_VISUAL
    conInfo.rangeElectrical = SETTING_CONST.RANGE_ELECTRICAL
    conInfo.rangeSpecNo = SETTING_CONST.RANGE_SPEC_NO
    conInfo.rangeInspectedBy = SETTING_CONST.RANGE_INSPECTED_BY
    conInfo.rangeApprovedBy = SETTING_CONST.RANGE_APPROVED_BY
    conInfo.colItem = SETTING_CONST.COL_ITEM
    conInfo.colSpc = SETTING_CONST.COL_SPC
    conInfo.colUnit = SETTING_CONST.COL_UNIT
    conInfo.colMax = SETTING_CONST.COL_MAX
    conInfo.colMin = SETTING_CONST.COL_MIN
    conInfo.colAvg = SETTING_CONST.COL_AVG
    conInfo.colSd = SETTING_CONST.COL_SD
    conInfo.colCpk = SETTING_CONST.COL_CPK
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
    outHeader.Article = csvInfo.commonInfo.KcHinmei
    'Quantity：(18)出荷数
    outHeader.Quantity = csvInfo.commonInfo.ShukkaSu
    'Lot No.：(16)出荷ﾛｯﾄNo
    outHeader.LotNo = csvInfo.commonInfo.ShukkaLotNo
    'Date：(14)出荷日
    outHeader.InspectionDate = "=UPPER(TEXT(" & csvInfo.commonInfo.ShukkaDate & ",""mmm dd,yyyy""))"
    'Drawing No：(22)KC図番
    outHeader.DrawingNo = csvInfo.commonInfo.KcZubanJyuchuzan
    'Tape Lot：(57)ﾃｰﾌﾟﾛｯﾄNo
    outHeader.TapeLot = csvInfo.commonInfo.TapeLotNo
    '7 Digit Lot：(58)北米ﾛｯﾄNo
    outHeader.SevenDigitLot = csvInfo.commonInfo.AbclotNo

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

            kikakuType = items.KikakuHantei

            'USL
            Dim outputUSL As String       'USL
            outputUSL = ""
            Select Case kikakuType
            Case "0", "1", "2", "9"
                '(116)規格ﾀｲﾌﾟが"0,1,2,9"の場合 (121)USL
                outputUSL = items.Usl
            Case "3", "8"
                '(116)規格ﾀｲﾌﾟが"3"の場合 空欄
                '(116)規格ﾀｲﾌﾟが「8:外観」の場合 空欄
                outputUSL = ""
            Case "6"
                '(116)規格ﾀｲﾌﾟが「6:耐熱性」の場合 (132)耐熱温度+"℃"+" -"+(133)耐熱時間+(134)耐熱時間単位+"Oven"
                outputUSL = items.Ondo & "℃ -" & items.Jikan & items.JikanTani & " Oven"
                '※USL/LSLのｾﾙを結合する。
                Dim mergeRange As String
                mergeRange = COL_USL & CStr(writingRow) & ":" & COL_LSL & CStr(writingRow)
                ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_MAIN).Range(mergeRange).Merge
            End Select
            outDetail.Usl = outputUSL
            
            'LSL
            Dim outputLSL As String
            outputLSL = ""
            Select Case kikakuType
            Case "0", "3", "9"
                '(116)規格ﾀｲﾌﾟが「0,3,9」の場合 (122)LSL
                outputLSL = items.Lsl
            Case "1", "2"
                '(116)規格ﾀｲﾌﾟが"1,2"の場合 空欄
                outputLSL = ""
            End Select
            outDetail.Lsl = outputLSL
            
            'UNIT：(107)単位
            outDetail.Unit = BlankToHyphen(items.TaniKensa)
            '測定値1～64
            Set outDetail.SokuteiValueCollection = items.SokuteiValueCollection
            
            '規格ﾀｲﾌﾟに応じて表記を変更
            Select Case kikakuType
            '「0:MAX/MIN/AVG」の場合：「MAX/MIN/AVG」の項目がある場合は、値をｾｯﾄする。「標準偏差/Cpk」の項目がある場合は、中央揃えに設定して「-」をｾｯﾄする。
            Case "6"
                Call Kikaku6Or8(outDetail)
            Case "8"
                Call Kikaku6Or8(outDetail)
            Case Else
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
    
    'Ceramic Body：(33)CERAMIC
    outFooter.Ceramic = csvInfo.commonInfo.Ceramic
    'Metallization：(34)METALIZE
    outFooter.Metalize = csvInfo.commonInfo.Metalize
    'Gold Plating：(36)PLATING
    outFooter.Plating = csvInfo.commonInfo.Plating
    'Dimension：(62)SAMPLING_PLAN2
    outFooter.SamplingPlan2 = csvInfo.commonInfo.SamplingPlan2
    'Platting Thickness：(63)SAMPLING_PLAN3
    outFooter.SamplingPlan3 = csvInfo.commonInfo.SamplingPlan3
    'VISUAL：(64)SAMPLING_PLAN4
    outFooter.SamplingPlan4 = csvInfo.commonInfo.SamplingPlan4
    'Electrical(Open/Short)：(65)SAMPLING_PLAN5
    outFooter.SamplingPlan5 = csvInfo.commonInfo.SamplingPlan5
    'Spec No：(24)SPEC
    outFooter.Spec = csvInfo.commonInfo.SpecJuchuzan
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
' メソッド      WriteHeader
' 機能          Headerを記述する
' 機能説明      Headerに値を与える
'-------------------------------------------------------------------------------
Sub WriteHeader(outHeader As Cls_OutputHeader)

    Call CellEdit(outHeader.Customer, SETTING_CONST.HEADER_RANGE_CUSTOMER)
    Call CellEdit(outHeader.Article, SETTING_CONST.HEADER_RANGE_ARTICLE_1)
    Call CellEdit(outHeader.Quantity, SETTING_CONST.HEADER_RANGE_INSPECTION_QUANTITY)
    Call CellEdit(outHeader.LotNo, SETTING_CONST.HEADER_RANGE_LOT_NO)
    Call CellEdit(outHeader.InspectionDate, SETTING_CONST.HEADER_RANGE_INSPECTION_DATE)
    Call CellEdit(outHeader.DrawingNo, SETTING_CONST.HEADER_RANGE_DRAWING_NO)
    Call CellEdit(outHeader.TapeLot, SETTING_CONST.HEADER_RANGE_TAPE_LOT)
    Call CellEdit(outHeader.SevenDigitLot, SETTING_CONST.HEADER_RANGE_7_DIGIT_LOT)

End Sub

'-------------------------------------------------------------------------------
' メソッド      WriteDetail
' 機能          明細部を記述する
' 機能説明      明細部に値を与える
'-------------------------------------------------------------------------------
Sub WriteDetail(outDetail As Cls_OutputDetail, writingRow As Long)

    Call CellEdit(outDetail.ItemName, COL_ITEM & CStr(writingRow), False, outDetail.ItemNameLfCnt)
    Call CellEdit(outDetail.SpcValue, COL_SPC & CStr(writingRow))
    Call CellEdit(outDetail.Usl, COL_USL & CStr(writingRow))
    Call CellEdit(outDetail.Lsl, COL_LSL & CStr(writingRow))
    Call CellEdit(outDetail.Unit, COL_UNIT & CStr(writingRow))
    Call CellEditSokuteichi(outDetail.SokuteiValueCollection, writingRow)
    Call CellEdit(outDetail.MaxN, COL_MAX & CStr(writingRow), outDetail.MaxCenterd)
    Call CellEdit(outDetail.MinN, COL_MIN & CStr(writingRow), outDetail.MinCenterd)
    Call CellEdit(outDetail.Ave, COL_AVG & CStr(writingRow), outDetail.AveCenterd)
    Call CellEdit(outDetail.Sigma, COL_SD & CStr(writingRow), outDetail.SigmaCenterd)
    Call CellEdit(outDetail.Cpk, COL_CPK & CStr(writingRow), outDetail.CpkCenterd)
    Call CellEdit(outDetail.result, COL_RESULT & CStr(writingRow))
    Call CellEdit(outDetail.AviApplied, COL_AVI & CStr(writingRow))

End Sub

'-------------------------------------------------------------------------------
' メソッド      WriteFooter
' 機能          Footerを記述する
' 機能説明      Footerに値を与える
'-------------------------------------------------------------------------------
Sub WriteFooter(outFooter As Cls_OutputFooter)

    Call CellEdit(outFooter.Ceramic, SETTING_CONST.RANGE_CERAMIC_BODY)
    Call CellEdit(outFooter.Metalize, SETTING_CONST.RANGE_METALLIZATION)
    Call CellEdit(outFooter.Plating, SETTING_CONST.RANGE_GOLD_PLATING)
    Call CellEdit(outFooter.SamplingPlan2, SETTING_CONST.RANGE_DIMENSION)
    Call CellEdit(outFooter.SamplingPlan3, SETTING_CONST.RANGE_PLATTING_THICKNESS)
    Call CellEdit(outFooter.SamplingPlan4, SETTING_CONST.RANGE_VISUAL)
    Call CellEdit(outFooter.SamplingPlan5, SETTING_CONST.RANGE_ELECTRICAL)
    Call CellEdit(outFooter.Spec, SETTING_CONST.RANGE_SPEC_NO)
    Call CellEdit(outFooter.InspectedBy, SETTING_CONST.RANGE_INSPECTED_BY)
    Call CellEdit(outFooter.ApprovedBy, SETTING_CONST.RANGE_APPROVED_BY)

End Sub

'-------------------------------------------------------------------------------
' メソッド      CellEdit
' 機能          セルの編集
' 機能説明　　  対象Bookのシートかつ対象セルの値・書式の編集
'-------------------------------------------------------------------------------
Sub CellEdit(pasteData As String, pasteCell As String, Optional centered As Boolean = False, Optional kaigyo As Long = 0)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_MAIN)

    'セルへ記載
    ws.Range(pasteCell).Value = pasteData
    '書式設定を中央揃えに変更
    If (centered) Then
        ws.Range(pasteCell).HorizontalAlignment = xlCenter
    End If
    '値の改行回数が1以上の時
    If kaigyo > 0 Then
        Dim defaultHight As Long
        Dim plusHight As Long
        defaultHight = ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_PARAM).Range(SETTING_CONST.PARAM_ROW).Value
        plusHight = ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_PARAM).Range(SETTING_CONST.PARAM_ADD_HEIGHT).Value
        
        ws.Range(pasteCell).RowHeight = defaultHight + (plusHight * kaigyo)
    End If
End Sub

'-------------------------------------------------------------------------------
' メソッド      CellEditSokuteichi
' 機能          セルの編集
' 機能説明　　  対象Bookのシートかつ対象セルの値・書式の編集
'-------------------------------------------------------------------------------
Sub CellEditSokuteichi(SokuteiValueCollection As Collection, currentRow As Long)

    Dim sokuteiIndex As Long
    Dim sokuteiValueArray() As Variant
    If SokuteiValueCollection Is Nothing Then
        Exit Sub
    End If
    sokuteiValueArray = SokuteichiCollectionToArray(SokuteiValueCollection)
    For sokuteiIndex = 0 To UBound(sokuteiValueArray)
        sokuteiValueArray(sokuteiIndex) = BlankToHyphen(CStr(sokuteiValueArray(sokuteiIndex)))
    Next sokuteiIndex

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_MAIN)
    If UBound(sokuteiValueArray) <= 0 Then
        Exit Sub
    End If
    
    With ws
        .Range(.Cells(currentRow, SETTING_CONST.COL_NUM_SOKUTEICHI_START), .Cells(currentRow, SETTING_CONST.COL_NUM_SOKUTEICHI_END)).Value = sokuteiValueArray
    End With

End Sub

'-------------------------------------------------------------------------------
' メソッド      BlankToHyphen
' 機能          値無しは-を返す
'-------------------------------------------------------------------------------
Private Function BlankToHyphen(targetString As String) As String
    If targetString = "" Then
        BlankToHyphen = "-"
    Else
        BlankToHyphen = targetString
    End If
End Function

'-------------------------------------------------------------------------------
' メソッド      SokuteichiCollectionToArray
' 機能          測定値のCollectionを配列にする
'-------------------------------------------------------------------------------
Private Function SokuteichiCollectionToArray(ByVal targetCollection As Collection) As Variant
 
    Dim resultArray(SOKUTEICHI_MAX)

    Dim index
    Dim val
     
    ' indexの初期値を設定する
    index = 0
    For Each val In targetCollection
     
        resultArray(index) = val
        index = index + 1
        If SOKUTEICHI_MAX <= index Then
            Exit For
        End If
    Next
  
    ' 戻り値に設定する
    SokuteichiCollectionToArray = resultArray
 
End Function

'-------------------------------------------------------------------------------
' メソッド      Kikaku6Or8
' 機能          規格ﾀｲﾌﾟが6,8の場合の処理を関数化
' 機能説明      規格ﾀｲﾌﾟが6,8の場合にUNIT～CPKまでの箇所が"-"になるため関数化
'-------------------------------------------------------------------------------
Sub Kikaku6Or8(ByRef outDetail As Cls_OutputDetail)
    
    '(116)規格ﾀｲﾌﾟが「6:耐熱性、8:外観」の場合は"-"
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
' メソッド      ResultString
' 機能          結果の数値を文字列に変換する
' 機能説明      カスタマイズを楽にするための関数
'-------------------------------------------------------------------------------
Private Function ResultString(ByVal target)

    Select Case target
    Case "0"
        '戻り値
        ResultString = "NG"
    Case "1"
        '戻り値
        ResultString = "OK"
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