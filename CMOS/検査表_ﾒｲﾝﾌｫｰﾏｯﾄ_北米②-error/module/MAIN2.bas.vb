Option Explicit

Sub MainExec()
    
    Dim exeResult As Boolean
    Dim conInfo As ConstInfo
    Set conInfo = New ConstInfo
    
    '����p���p�ŃZ�b�g
    Call SetConst(conInfo)
    '���ʏ����Ăяo��
    exeResult = Application.Run("'" & ThisWorkbook.Path & "\" & SETTING_CONST.EXE_COMMON_NAME & "'!MAIN.StartPrintPdf", ThisWorkbook, conInfo)
    Workbooks(SETTING_CONST.EXE_COMMON_NAME).Close
    
    If exeResult Then
        Err.Raise Number:=-1, Description:="�װ���e"
    End If

End Sub

Sub SetConst(ByRef conInfo As ConstInfo)
    
    '�ύX�s�v
    conInfo.formLink = "'" & ThisWorkbook.FullName & "'"
    conInfo.detailStartRow = CLng(ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_PARAM).Range(PARAM_HEADER_END).Value) + 1
    
    '�萔�ɏC�����L��ꍇ�̂ݑΉ�
    
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

End Sub

'-------------------------------------------------------------------------------
' ���\�b�h      writeElement
' �@�\          �{���[�ɒl����������
' �@�\�����@�@  CSV�̒l��{���[�ɒl����������
'-------------------------------------------------------------------------------
Public Function writeElement(ByRef csvInfo As Variant, ByRef argInfo As Variant) As Boolean
    
    Dim i As Long
    Dim items As Variant
    Dim detailStartRow As Long
    Dim paramWs As Worksheet        '���Ұ����Obj
    Dim customerParam As String     '�ڋq��/Customer
    Dim lfCount As Long             '���s������
    Dim kikakuType As String        '�K�i����
    Dim addRowCount As Long         '���ޕ\�L�ǉ��̃J�E���g�p
    Dim writingRow As Long          '�L���s
    
    Dim outHeader As Cls_OutputHeader '�o�͍���(ͯ�ް)
    Dim outDetail As Cls_OutputDetail '�o�͍���(����)
    Dim outFooter As Cls_OutputFooter '�o�͍���(̯��)
    
'    On Error GoTo ErrorCatch
    
    '�o�͍��ڈꎞ�ۑ�
    Set outHeader = New Cls_OutputHeader

    '���Ұ����Obj�擾
    Set paramWs = ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_PARAM)

    '���׊J�n�s�v�Z
    detailStartRow = CLng(ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_PARAM).Range(PARAM_HEADER_END).Value) + 1

    'Customer
    '���Ұ��̌ڋq��/Customer�Q��
    customerParam = paramWs.Range(SETTING_CONST.PARAM_CUSTOMER).Value
    '1�̏ꍇ�A(20)����旪����
    If customerParam = "1" Then
        outHeader.Customer = csvInfo.commonInfo.FurimukesakiName
    '2�̏ꍇ�A(79)����於
    ElseIf customerParam = "2" Then
        outHeader.Customer = csvInfo.commonInfo.TorihikisakiSeishikiNameJyuchuzan
    End If
    'Article�F(21)KC�i��
    outHeader.Article = csvInfo.commonInfo.KcHinmei
    'Quantity�F(18)�o�א�
    outHeader.Quantity = csvInfo.commonInfo.ShukkaSu
    'Lot No.�F(16)�o��ۯ�No
    outHeader.LotNo = csvInfo.commonInfo.ShukkaLotNo
    'Date�F(14)�o�ד�
    outHeader.InspectionDate = "=UPPER(TEXT(""" & csvInfo.commonInfo.ShukkaDate & """,""mmm dd,yyyy""))"
    'Drawing No�F(22)KC�}��
    outHeader.DrawingNo = csvInfo.commonInfo.KcZubanJyuchuzan
    'Tape Lot�F(57)ð��ۯ�No
    outHeader.TapeLot = csvInfo.commonInfo.TapeLotNo
    '7 Digit Lot�F(58)�k��ۯ�No
    outHeader.SevenDigitLot = csvInfo.commonInfo.AbclotNo

    Call WriteHeader(outHeader)
    
    'csv�s���Ǝ��L���s�̍���
    addRowCount = detailStartRow

    '2. ���ו����֒l���
    For i = 0 To csvInfo.DetailInfoList.Count - 1
    
        '�o�͍��ڃN���X
        Set outDetail = New Cls_OutputDetail
        '�Ώ�CSV����
        Set items = csvInfo.DetailInfoList(i + 1)
        '�L���s��
        writingRow = i + detailStartRow
        
        '[1:����]�̏ꍇ�͕��ނ��L��
        If items.itemKbn = "1" Then
            outDetail.ItemName = items.ItemName

        '[2:���荀�ږ�]�̏ꍇ�͑��荀�ږ��ƌ��ʂ��L��
        ElseIf items.itemKbn = "2" Then
            If Left(paramWs.Range(SETTING_CONST.PARAM_LENGTH_LIMIT), 1) = "1" Then
                'ITEM(���荀�ږ�)�F(104)���荀�ږ�
                outDetail.ItemName = items.ItemName
            ElseIf Left(paramWs.Range(SETTING_CONST.PARAM_LENGTH_LIMIT), 1) = "2" Then
                '���s�񐔂ɉ����ĕύX ��LenB�Ƃ���ƁA�����v�Z���ʂ��{�ɂȂ邽�ߒ���
                lfCount = Len(items.ItemName) - Len(Replace(items.ItemName, vbLf, ""))
                outDetail.ItemNameLfCnt = lfCount
                'ITEM(���荀�ږ�)�F(104)���荀�ږ�
                outDetail.ItemName = items.ItemName
            End If
            'SPC�F(123)SPC
            outDetail.SpcValue = items.Spc1

            kikakuType = items.KikakuHantei

            'USL
            Dim outputUSL As String       'USL
            outputUSL = ""
            Select Case kikakuType
            Case "0", "1", "2", "9"
                '(116)�K�i���߂�"0,1,2,9"�̏ꍇ (121)USL
                outputUSL = items.Usl
            Case "3", "8"
                '(116)�K�i���߂�"3"�̏ꍇ ��
                '(116)�K�i���߂��u8:�O�ρv�̏ꍇ ��
                outputUSL = ""
            Case "6"
                '(116)�K�i���߂��u6:�ϔM���v�̏ꍇ (132)�ϔM���x+"��"+" -"+(133)�ϔM����+(134)�ϔM���ԒP��+"Oven"
                outputUSL = items.Ondo & "�� -" & items.Jikan & items.JikanTani & " Oven"
                '��USL/LSL�̾ق���������B
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
                '(116)�K�i���߂��u0,3,9�v�̏ꍇ (122)LSL
                outputLSL = items.Lsl
            Case "1", "2"
                '(116)�K�i���߂�"1,2"�̏ꍇ ��
                outputLSL = ""
            End Select
            outDetail.Lsl = outputLSL
            
            'UNIT�F(107)�P��
            outDetail.Unit = BlankToHyphen(items.TaniKensa)
            '����l1�`64
            Set outDetail.SokuteiValueCollection = items.SokuteiValueCollection
            
            '�K�i���߂ɉ����ĕ\�L��ύX
            Select Case kikakuType
            '�u0:MAX/MIN/AVG�v�̏ꍇ�F�uMAX/MIN/AVG�v�̍��ڂ�����ꍇ�́A�l��Ă���B�u�W���΍�/Cpk�v�̍��ڂ�����ꍇ�́A���������ɐݒ肵�āu-�v��Ă���B
            Case "6"
                Call Kikaku6Or8(outDetail)
            Case "8"
                Call Kikaku6Or8(outDetail)
            Case Else
                'MAX�F(108)�ő�l
                outDetail.MaxN = items.MaxN
                outDetail.MaxCenterd = False
                'MIN�F(109)�ŏ��l
                outDetail.MinN = items.MinN
                outDetail.MinCenterd = False
                'AVG�F(110)���ϒl
                outDetail.Ave = items.Ave
                outDetail.AveCenterd = False
                'SD�F(111)�W���΍�
                outDetail.Sigma = items.Sigma
                outDetail.SigmaCenterd = False
                'Cpk�F(112)cpk
                outDetail.Cpk = items.Cpk
                outDetail.CpkCenterd = False
            End Select

            'RESULT�F(115)����
            outDetail.result = ResultString(items.Hantei)
            '100% AVI applied�F(128)�摜����
            outDetail.AviApplied = AviAppliedString(items.GazoKensa)
        Else
            addRowCount = addRowCount + 1
        End If
    
        Call WriteDetail(outDetail, writingRow)
    
    Next i
    
    '�ŏI�s��ۑ�
    argInfo.endRows = i + addRowCount - 1

    '3. ̯�������֒l���
    Set outFooter = New Cls_OutputFooter
    
    'Ceramic Body�F(33)CERAMIC
    outFooter.Ceramic = csvInfo.commonInfo.Ceramic
    'Metallization�F(34)METALIZE
    outFooter.Metalize = csvInfo.commonInfo.Metalize
    'Gold Plating�F(36)PLATING
    outFooter.Plating = csvInfo.commonInfo.Plating
    'Dimension�F(62)SAMPLING_PLAN2
    outFooter.SamplingPlan2 = csvInfo.commonInfo.SamplingPlan2
    'Platting Thickness�F(63)SAMPLING_PLAN3
    outFooter.SamplingPlan3 = csvInfo.commonInfo.SamplingPlan3
    'VISUAL�F(64)SAMPLING_PLAN4
    outFooter.SamplingPlan4 = csvInfo.commonInfo.SamplingPlan4
    'Electrical(Open/Short)�F(65)SAMPLING_PLAN5
    outFooter.SamplingPlan5 = csvInfo.commonInfo.SamplingPlan5
    'Spec No�F(24)SPEC
    outFooter.Spec = csvInfo.commonInfo.SpecJuchuzan
    'Inspected by�F(53)�쐬�Җ�
    outFooter.InspectedBy = csvInfo.commonInfo.KensahyoHakkoshaName
    'Approved by�F(54)���F�Җ�
    outFooter.ApprovedBy = csvInfo.commonInfo.KensahyoShoninshaName
    
    Call WriteFooter(outFooter)
    
    Exit Function

ErrorCatch:
    writeElement = True

End Function

'-------------------------------------------------------------------------------
' ���\�b�h      WriteHeader
' �@�\          Header���L�q����
' �@�\����      Header�ɒl��^����
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
' ���\�b�h      WriteDetail
' �@�\          ���ו����L�q����
' �@�\����      ���ו��ɒl��^����
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
' ���\�b�h      WriteFooter
' �@�\          Footer���L�q����
' �@�\����      Footer�ɒl��^����
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
' ���\�b�h      CellEdit
' �@�\          �Z���̕ҏW
' �@�\�����@�@  �Ώ�Book�̃V�[�g���ΏۃZ���̒l�E�����̕ҏW
'-------------------------------------------------------------------------------
Sub CellEdit(pasteData As String, pasteCell As String, Optional centered As Boolean = False, Optional kaigyo As Long = 0)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_MAIN)

    '�Z���֋L��
    ws.Range(pasteCell).Value = pasteData
    '�����ݒ�𒆉������ɕύX
    If (centered) Then
        ws.Range(pasteCell).HorizontalAlignment = xlCenter
    End If
    '�l�̉��s�񐔂�1�ȏ�̎�
    If kaigyo > 0 Then
        Dim defaultHight As Long
        Dim plusHight As Long
        defaultHight = ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_PARAM).Range(SETTING_CONST.PARAM_ROW).Value
        plusHight = ThisWorkbook.Worksheets(SETTING_CONST.SHEET_NAME_PARAM).Range(SETTING_CONST.PARAM_ADD_HEIGHT).Value
        
        ws.Range(pasteCell).RowHeight = defaultHight + (plusHight * kaigyo)
    End If
End Sub

'-------------------------------------------------------------------------------
' ���\�b�h      CellEditSokuteichi
' �@�\          �Z���̕ҏW
' �@�\�����@�@  �Ώ�Book�̃V�[�g���ΏۃZ���̒l�E�����̕ҏW
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
' ���\�b�h      BlankToHyphen
' �@�\          �l������-��Ԃ�
'-------------------------------------------------------------------------------
Private Function BlankToHyphen(targetString As String) As String
    If targetString = "" Then
        BlankToHyphen = "-"
    Else
        BlankToHyphen = targetString
    End If
End Function

'-------------------------------------------------------------------------------
' ���\�b�h      SokuteichiCollectionToArray
' �@�\          ����l��Collection��z��ɂ���
'-------------------------------------------------------------------------------
Private Function SokuteichiCollectionToArray(ByVal targetCollection As Collection) As Variant
 
    Dim resultArray(SOKUTEICHI_MAX)

    Dim index
    Dim val
     
    ' index�̏����l��ݒ肷��
    index = 0
    For Each val In targetCollection
     
        resultArray(index) = val
        index = index + 1
        If SOKUTEICHI_MAX <= index Then
            Exit For
        End If
    Next
  
    ' �߂�l�ɐݒ肷��
    SokuteichiCollectionToArray = resultArray
 
End Function

'-------------------------------------------------------------------------------
' ���\�b�h      Kikaku6Or8
' �@�\          �K�i���߂�6,8�̏ꍇ�̏������֐���
' �@�\����      �K�i���߂�6,8�̏ꍇ��UNIT�`CPK�܂ł̉ӏ���"-"�ɂȂ邽�ߊ֐���
'-------------------------------------------------------------------------------
Sub Kikaku6Or8(ByRef outDetail As Cls_OutputDetail)
    
    '(116)�K�i���߂��u6:�ϔM���A8:�O�ρv�̏ꍇ��"-"
    'MAX�F(108)�ő�l
    outDetail.MaxN = "-"
    outDetail.MaxCenterd = True
    'MIN�F(109)�ŏ��l
    outDetail.MinN = "-"
    outDetail.MinCenterd = True
    'AVG�F(110)���ϒl
    outDetail.Ave = "-"
    outDetail.AveCenterd = True
    'SD�F(111)�W���΍�
    outDetail.Sigma = "-"
    outDetail.SigmaCenterd = True
    'Cpk�F(112)cpk
    outDetail.Cpk = "-"
    outDetail.CpkCenterd = True

End Sub

'-------------------------------------------------------------------------------
' ���\�b�h      ResultString
' �@�\          ���ʂ̐��l�𕶎���ɕϊ�����
' �@�\����      �J�X�^�}�C�Y���y�ɂ��邽�߂̊֐�
'-------------------------------------------------------------------------------
Private Function ResultString(ByVal target)

    Select Case target
    Case "0"
        '�߂�l
        ResultString = "NG"
    Case "1"
        '�߂�l
        ResultString = "OK"
    End Select
End Function

'-------------------------------------------------------------------------------
' ���\�b�h      AviApplied
' �@�\          ���ʂ̐��l�𕶎���ɕϊ�����
' �@�\����      �J�X�^�}�C�Y���y�ɂ��邽�߂̊֐�
'-------------------------------------------------------------------------------
Function AviAppliedString(ByVal target) As String

    Select Case target
    Case "0"
        '�߂�l
        AviAppliedString = ""
    Case "1"
        '�߂�l
        AviAppliedString = "�Z"
    End Select
End Function

'-------------------------------------------------------------------------------
' ���\�b�h      MainSheetNameCreate
' �@�\          �e���[��Ҳݼ�Ė����擾
' �@�\�����@�@  ���[�����e���[���ɈႤ���߁A�J�X�^�}�C�Y���y�ɂ��邽�߂̊֐�
'-------------------------------------------------------------------------------
Function MainSheetNameCreate(csvInfo)
    
    Dim result(2)
    
    On Error GoTo ErrorCatch
    
    '�C���s�v
    result(0) = "False"
    
    '��Ė��F�e���[���ƂɏC���\��
    result(1) = csvInfo.commonInfo.KcZubanJyuchuzan
    MainSheetNameCreate = result
    
    Exit Function
    
ErrorCatch:
    
    result(0) = "True"
    
    MainSheetNameCreate = result
    
End Function