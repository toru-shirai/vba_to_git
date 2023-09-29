Option Explicit
Public grb_csvInfo As Cls_CsvInfo
Public grb_argInfo As Cls_ArgInfo
Public grb_targetWb As Workbook
Public grb_constInfo As Variant

Public Function StartPrintPdf(wb As Workbook, constInfo As Variant) As Boolean

'    On Error GoTo Errored
    
    StartPrintPdf = False

    '対象帳票のBookObjを取得
    '対象フォーマットの定数一覧をグローバル変数へ代入 ※定数ではないため変更しない様にする
    Set grb_targetWb = wb
    Set grb_constInfo = constInfo

    '画面更新停止
    Application.ScreenUpdating = False
    'メッセージを非表示
    Application.DisplayAlerts = False

    '(1)-1 CSVﾌｧｲﾙ読込み
    Call ImportCSV

    '(1)-2 ｻﾌﾞ帳票ﾌｫｰﾏｯﾄ取込 ※サブモジュールの関数名はSubExecに統一予定 変更時はExecSubを修正
    Call ExecSub

    '(1)-3 ﾒｲﾝｼｰﾄｺﾋﾟｰ
    Call ExecMain

    '(1)-4 改ﾍﾟｰｼﾞ組み込み
    If (grb_csvInfo.commonInfo.InsatsuType0 = "1" _
    Or grb_csvInfo.commonInfo.InsatsuType1 = "1" _
    Or grb_csvInfo.commonInfo.InsatsuType2 = "1" _
    ) Then
        Call MainPageBreak
    End If

    '(1)-5 ｴﾋﾞﾃﾞﾝｽ用ｴｸｾﾙ保存
    Call SaveBook
    
    '(1)-6 不要ｼｰﾄ削除
    Call DeleteSheets

    '(1)-7 印刷実行
    Call PrintOut

    '画面更新停止解除
    Application.ScreenUpdating = True
    'メッセージを表示
    Application.DisplayAlerts = True

    Exit Function

Errored:

    '画面更新停止解除
    Application.ScreenUpdating = True
    'メッセージを表示
    Application.DisplayAlerts = True

    'ｴﾗｰが発生したことを呼び出し元へ伝える
    StartPrintPdf = True

End Function

'-------------------------------------------------------------------------------
' メソッド      ImportCSV
' 機能          CSVを配列にし、ﾃﾞｰﾀｼｰﾄに貼り付ける。存在チェックに問題があればFalseを返す
' 機能説明　　  複数の必要な変数や関数呼び出しとクラス変換をまとめたラップ関数
'-------------------------------------------------------------------------------
Sub ImportCSV()

    Dim paramWs As Worksheet
    Dim defaultArray As Variant     '配列：CSVファイル1
    Dim detailArray As Variant      '配列：CSVファイル2(明細)
    Dim multiLotArray As Variant    '配列：CSVファイル3(複数ロット)
    Dim csvFolderPath As String     'CSVファイル保存場所
    Dim csvFilePath1 As String      'CSVファイル1
    Dim csvFilePath2 As String      'CSVファイル2(明細)
    Dim csvFilePath3 As String      'CSVファイル3(複数ロット)
    Dim csvFile3Exist As Boolean    'CSVファイル3(複数ロット)の要否
    Dim csvPasteSheet1 As String    'CSVファイル1の貼り付け先シート名
    Dim csvPasteSheet2 As String    'CSVファイル2(明細)の貼り付け先シート名
    Dim csvPasteSheet3 As String    'CSVファイル3(複数ロット)の貼り付け先シート名

    Set paramWs = grb_targetWb.Worksheets(grb_constInfo.sheetNameParam)

    csvFolderPath = paramWs.Range(grb_constInfo.paramCsvPath)
    csvFilePath1 = csvFolderPath & "\" & paramWs.Range(grb_constInfo.paramCsv1)
    csvFilePath2 = csvFolderPath & "\" & paramWs.Range(grb_constInfo.paramCsv2)
    csvFilePath3 = csvFolderPath & "\" & paramWs.Range(grb_constInfo.paramCsv3)
   'Boolean型
    csvFile3Exist = (Left(paramWs.Range(grb_constInfo.paramCsv3Check).Value, 1) = "1")
    csvPasteSheet1 = paramWs.Range(grb_constInfo.paramDataSheet1)
    csvPasteSheet2 = paramWs.Range(grb_constInfo.paramDataSheet2)
    csvPasteSheet3 = paramWs.Range(grb_constInfo.paramDataSheet3)

    'CSVの配列化と貼り付け処理 ※-1は改行コード指定 LF(10) CR(13) CRLF(-1)
    defaultArray = CsvToArray(csvFilePath1, "UTF-8", -1, csvPasteSheet1)
    detailArray = CsvToArray(csvFilePath2, "UTF-8", -1, csvPasteSheet2)
    '2次パース 測定値を更に細分化 カンマ区切りで項目を持っている
    'Call ParseSecond(detailArray, csvPasteSheet2)
    If (csvFile3Exist) Then
        multiLotArray = CsvToArray(csvFilePath3, "UTF-8", -1, csvPasteSheet3)
    End If
    '配列をグローバルクラスへ変換
    Call ArrayToClass(defaultArray, detailArray, multiLotArray, csvFile3Exist)
    '2次パース 測定値を更に細分化 カンマ区切りで項目を持っている
    Call ParseSokuteichi(grb_csvInfo.DetailInfoList, csvPasteSheet2)
    '同タイミングにてargInfoも生成
    Set grb_argInfo = New Cls_ArgInfo

End Sub

'-------------------------------------------------------------------------------
' メソッド      CsvToArray
' 機能          CSVを配列にし、ﾃﾞｰﾀｼｰﾄに貼り付ける
' 機能説明　　  CSVﾌｧｲﾙを読込み配列に変換し、セルへの貼り付け用関数を呼び出す
'               lineSeparator: 改行コード指定 LF(10) CR(13) CRLF(-1)
'               charset: 文字コード指定
'-------------------------------------------------------------------------------
Function CsvToArray(targetPath As String, charset As String, lineSeparator As Long, sheetName As String) As String()

    Dim i As Long
    Dim j As Long
    Dim max_n As Long
    Dim rowCount As Long
    Dim buf As String           'バッファ
    Dim tmp() As String         '一時保持用変数
    Dim a_sArLine() As String   '戻り値

    'FileSystemObjectの生成
    With CreateObject("Scripting.FileSystemObject").OpenTextFile(targetPath, 8)
        max_n = .Line 'ﾌｧｲﾙ行数取得
        .Close
    End With

    i = 0

    'CSVをﾃｷｽﾄ読み込み
    With CreateObject("ADODB.Stream")
        .charset = charset
        .lineSeparator = lineSeparator
        .Open
        .LoadFromFile targetPath
        Do Until .EOS
            buf = .ReadText(-2)
            tmp = AsSplit(buf)
            If i = 0 Then
                '取得した行数で2次元配列の再定義
                ReDim a_sArLine(max_n, UBound(tmp))
            End If
            For j = 0 To UBound(tmp)
                '分割した内容を配列の項目へ入れる
                a_sArLine(i, j) = tmp(j)
            Next j
            i = i + 1
        Loop
        .Close
    End With

    Call ArrayPaste(a_sArLine, sheetName)

    '戻り値
    CsvToArray = a_sArLine

End Function

'-------------------------------------------------------------------------------
' メソッド      AsSplit
' 機能          CSVの1行を分割した配列を返す
' 機能説明　　  全項目がダブルクォーテーションで囲まれているCSVの1行を分割した配列を返す
'-------------------------------------------------------------------------------
Function AsSplit(ByVal buf As String) As String()

    Dim result() As String  '戻り値
    Dim str As String       '文字列
    Dim i As Long

    '行の先頭と末尾の1文字を除外する ※["]を除外する処理
    buf = Mid(buf, 2, Len(buf) - 2)
    '[","]を対象にSplitする
    AsSplit = Split(buf, """,""")

End Function

'-------------------------------------------------------------------------------
' メソッド      parseSecond
' 機能          測定値の値取得
' 機能説明　　  測定値に入っているCSV形式のデータ配列化し、セルへの貼り付け用関数を呼び出す
'-------------------------------------------------------------------------------
Sub ParseSecond(targetArray As Variant, sheetName As String)

    Dim i As Long
    Dim j As Long
    Dim colCount As Long
    Dim tmp() As String
    Dim innerArray() As Variant

    colCount = UBound(targetArray) - 1

    For i = 0 To colCount
        tmp = Split(targetArray(i + 1, detailEnum.CSV_IDX_SOKUTEI_Value - 1), ",")
        If i = 0 Then
            ReDim innerArray(colCount, UBound(tmp))
        End If
        For j = 0 To UBound(tmp)
            innerArray(i, j) = tmp(j)
        Next j
    Next i

    Call ArrayPaste(innerArray, sheetName, UBound(innerArray), UBound(innerArray, 2))

End Sub

'-------------------------------------------------------------------------------
' メソッド      parseSecond
' 機能          測定値の値取得
' 機能説明　　  測定値に入っているCSV形式のデータ配列化し、セルへの貼り付け用関数を呼び出す
'-------------------------------------------------------------------------------
Sub ParseSokuteichi(targetArray As Collection, sheetName As String)

    Dim i As Long
    Dim j As Long
    Dim colCount As Long
    Dim tmp() As String
    Dim innerArray() As Variant
    Dim maxSokuteichi As Long

    colCount = targetArray.Count - 1
    maxSokuteichi = -1

    For i = 0 To colCount
        tmp = Split(targetArray(i + 1).SokuteiValue, ",")
        If maxSokuteichi < UBound(tmp) Then
            ReDim Preserve innerArray(colCount, UBound(tmp))
            maxSokuteichi = UBound(tmp)
        End If

        Set targetArray(i + 1).SokuteiValueCollection = New Collection
        For j = 0 To UBound(tmp)
            targetArray(i + 1).SokuteiValueCollection.Add (tmp(j))
            innerArray(i, j) = tmp(j)
        Next j
    Next i

    Call ArrayPaste(innerArray, sheetName, , 200)
End Sub
'-------------------------------------------------------------------------------
' メソッド      ArrayPaste
' 機能          セルへの配列の貼り付け処理
' 機能説明      セルへの配列の貼り付け処理
'-------------------------------------------------------------------------------
Sub ArrayPaste(pasteArray As Variant, sheetName As String, Optional rowNum As Long = 1, Optional columnNum As Long = 1)

    Dim endRow As Long
    Dim endColumn As Long

    endRow = UBound(pasteArray) + rowNum
    endColumn = AsUBound(pasteArray, 2) + columnNum

    'pasteArrayが1次配列になるのは測定値を持ってきたときのみ
    With grb_targetWb.Worksheets(sheetName)
        .Range(.Cells(rowNum, columnNum), .Cells(endRow, endColumn)).Value = pasteArray
    End With

End Sub
'-------------------------------------------------------------------------------
' メソッド      AsUBound
' 機能          UBound関数のラップ関数
' 機能説明      次元数違反でエラーを起こさずに1を返す
'-------------------------------------------------------------------------------
Public Function AsUBound(var, dimension As Long) As Long
    AsUBound = 1
    On Error Resume Next
    '戻り値
    AsUBound = UBound(var, dimension)
End Function

'-------------------------------------------------------------------------------
' メソッド      ArrayToClass
' 機能          配列をクラスに設定する
' 機能説明      引数の配列をグローバル変数のクラスに設定する
'-------------------------------------------------------------------------------
Sub ArrayToClass(defItem, detailItem, multiLotItem, multiExist)

    Dim i As Long

    Set grb_csvInfo = New Cls_CsvInfo

    'FSJD005 検査表出力ﾃﾞｰﾀ
    Set grb_csvInfo.commonInfo = ConvertCommonInfo(defItem)

    'FSJD006 検査表出力ﾃﾞｰﾀ(明細)
    Set grb_csvInfo.DetailInfoList = New Collection
    '明細情報を設定しコレクションに追加
    For i = 1 To UBound(detailItem)
        If Not detailItem(i, 0) = "" Then
            Call grb_csvInfo.DetailInfoList.Add(ConvertDetailInfo(detailItem, i))
        End If
    Next i

    If (multiExist) Then
        'FSJD007 検査表出力ﾃﾞｰﾀ(複数ﾛｯﾄ)
        Set grb_csvInfo.MultiLotInfoList = New Collection
        '明細情報を設定しコレクションに追加
        For i = 1 To UBound(multiLotItem)
            If Not multiLotItem(i, 0) = "" Then
                Call grb_csvInfo.MultiLotInfoList.Add(ConvertMultiLotInfo(multiLotItem, i))
            End If
        Next i
    End If


End Sub

'-------------------------------------------------------------------------------
' メソッド      ConvertCommonInfo
' 機能          配列をクラスに設定する(FSJD005 検査表出力ﾃﾞｰﾀ)
' 機能説明      引数の配列をグローバル変数のクラス(Cls_Common)に設定する
'-------------------------------------------------------------------------------
Function ConvertCommonInfo(var) As Cls_Common

    Dim commonInfo As Cls_Common
    Set commonInfo = New Cls_Common

    '印刷種別:直接印刷
    commonInfo.InsatsuType0 = var(1, headerEnum.CSV_IDX_INSATSU_TYPE0 - 1)
    '印刷種別:PDF
    commonInfo.InsatsuType1 = var(1, headerEnum.CSV_IDX_INSATSU_TYPE1 - 1)
    '印刷種別:Excel
    commonInfo.InsatsuType2 = var(1, headerEnum.CSV_IDX_INSATSU_TYPE2 - 1)
    '印刷種別:CSV
    commonInfo.InsatsuType3 = var(1, headerEnum.CSV_IDX_INSATSU_TYPE3 - 1)
    '印刷種別:TSV
    commonInfo.InsatsuType5 = var(1, headerEnum.CSV_IDX_INSATSU_TYPE5 - 1)
    '出力先(PDF)
    commonInfo.Shutsuryokusaki1 = var(1, headerEnum.CSV_IDX_SHUTSURYOKUSAKI1 - 1)
    '出力先(Excel)
    commonInfo.Shutsuryokusaki2 = var(1, headerEnum.CSV_IDX_SHUTSURYOKUSAKI2 - 1)
    '出力先(CSV)
    commonInfo.Shutsuryokusaki3 = var(1, headerEnum.CSV_IDX_SHUTSURYOKUSAKI3 - 1)
    '出力先(TSV)
    commonInfo.Shutsuryokusaki5 = var(1, headerEnum.CSV_IDX_SHUTSURYOKUSAKI5 - 1)
    '出荷日
    commonInfo.ShukkaDate = var(1, headerEnum.CSV_IDX_SHUKKA_DATE - 1)
    '製伝No
    commonInfo.SeidenNo = var(1, headerEnum.CSV_IDX_SEIDEN_NO - 1)
    '出荷ﾛｯﾄNo
    commonInfo.ShukkaLotNo = var(1, headerEnum.CSV_IDX_SHUKKA_LOT_NO - 1)
    '出荷数
    commonInfo.ShukkaSu = var(1, headerEnum.CSV_IDX_SHUKKA_SU - 1)
    '取引先ｺｰﾄﾞ
    commonInfo.TorihikisakiCd = var(1, headerEnum.CSV_IDX_TORIHIKISAKI_CD - 1)
    '取引先略式名
    commonInfo.FurimukesakiName = var(1, headerEnum.CSV_IDX_FURIMUKESAKI_NAME - 1)
    'KC品名
    commonInfo.KcHinmei = var(1, headerEnum.CSV_IDX_KC_HINMEI - 1)
    'KC図番
    commonInfo.KcZubanJyuchuzan = var(1, headerEnum.CSV_IDX_KC_ZUBAN_JYUCHUZAN - 1)
    '客先PN
    commonInfo.KyakusakiPartsNoJyuchuzan = var(1, headerEnum.CSV_IDX_KYAKUSAKI_PARTS_NO_JYUCHUZAN - 1)
    'SPEC
    commonInfo.SpecJuchuzan = var(1, headerEnum.CSV_IDX_SPEC_JUCHUZAN - 1)
    '注文番号
    commonInfo.ChumonNo = var(1, headerEnum.CSV_IDX_CHUMON_NO - 1)
    '品目ｺｰﾄﾞ
    commonInfo.HinmokuCd = var(1, headerEnum.CSV_IDX_HINMOKU_CD - 1)
    'ｻﾌﾞｺｰﾄﾞ
    commonInfo.SubCd = var(1, headerEnum.CSV_IDX_SUB_CD - 1)
    '営業所ｺｰﾄﾞ
    commonInfo.EigyoshoCd = var(1, headerEnum.CSV_IDX_EIGYOSHO_CD - 1)
    '材質ｺｰﾄﾞ
    commonInfo.ZaishitsuCd = var(1, headerEnum.CSV_IDX_ZAISHITSU_CD - 1)
    'CERAMIC
    commonInfo.Ceramic = var(1, headerEnum.CSV_IDX_CERAMIC - 1)
    'METALIZE
    commonInfo.Metalize = var(1, headerEnum.CSV_IDX_METALIZE - 1)
    'METAL
    commonInfo.Metal = var(1, headerEnum.CSV_IDX_METAL - 1)
    'PLATING
    commonInfo.Plating = var(1, headerEnum.CSV_IDX_PLATING - 1)
    '備考1
    commonInfo.Biko1 = var(1, headerEnum.CSV_IDX_BIKO_1 - 1)
    '備考2
    commonInfo.Biko2 = var(1, headerEnum.CSV_IDX_BIKO_2 - 1)
    '備考3
    commonInfo.Biko3 = var(1, headerEnum.CSV_IDX_BIKO_3 - 1)
    '備考4
    commonInfo.Biko4 = var(1, headerEnum.CSV_IDX_BIKO_4 - 1)
    '備考5
    commonInfo.Biko5 = var(1, headerEnum.CSV_IDX_BIKO_5 - 1)
    'めっき日
    commonInfo.MekkiDate = var(1, headerEnum.CSV_IDX_MEKKI_DATE - 1)
    '作成者名
    commonInfo.KensahyoHakkoshaName = var(1, headerEnum.CSV_IDX_KENSAHYO_HAKKOSHA_NAME - 1)
    '承認者名
    commonInfo.KensahyoShoninshaName = var(1, headerEnum.CSV_IDX_KENSAHYO_SHONINSHA_NAME - 1)
    '総合判定
    commonInfo.SogoHantei = var(1, headerEnum.CSV_IDX_SOGO_HANTEI - 1)
    'ﾃｰﾌﾟﾛｯﾄNo
    commonInfo.TapeLotNo = var(1, headerEnum.CSV_IDX_TAPE_LOT_NO - 1)
    'ABCﾛｯﾄNo
    commonInfo.AbclotNo = var(1, headerEnum.CSV_IDX_ABCLOT_NO - 1)
    '有効期限
    commonInfo.YukoDate = var(1, headerEnum.CSV_IDX_YUKO_DATE - 1)
    '検印者名
    commonInfo.KensahyoKeninshaName = var(1, headerEnum.CSV_IDX_KENSAHYO_KENINSHA_NAME - 1)
    'SAMPLING_PLAN1
    commonInfo.SamplingPlan1 = var(1, headerEnum.CSV_IDX_SAMPLING_PLAN1 - 1)
    'SAMPLING_PLAN2
    commonInfo.SamplingPlan2 = var(1, headerEnum.CSV_IDX_SAMPLING_PLAN2 - 1)
    'SAMPLING_PLAN3
    commonInfo.SamplingPlan3 = var(1, headerEnum.CSV_IDX_SAMPLING_PLAN3 - 1)
    'SAMPLING_PLAN4
    commonInfo.SamplingPlan4 = var(1, headerEnum.CSV_IDX_SAMPLING_PLAN4 - 1)
    'SAMPLING_PLAN5
    commonInfo.SamplingPlan5 = var(1, headerEnum.CSV_IDX_SAMPLING_PLAN5 - 1)
    '箱番号(明細ｼｰｹﾝｽで代用)
    commonInfo.HakoNo = var(1, headerEnum.CSV_IDX_HAKO_NO - 1)
    'PROGRAM
    commonInfo.Program = var(1, headerEnum.CSV_IDX_PROGRAM - 1)
    'MCO
    commonInfo.Mco = var(1, headerEnum.CSV_IDX_MCO - 1)
    'MI
    commonInfo.Mi = var(1, headerEnum.CSV_IDX_MI - 1)
    'CONFIG
    commonInfo.Config = var(1, headerEnum.CSV_IDX_CONFIG - 1)
    'MS
    commonInfo.Ms = var(1, headerEnum.CSV_IDX_MS - 1)
    '磁器ﾁｪｯｸﾗﾝｸ
    commonInfo.JikiCheckRank = var(1, headerEnum.CSV_IDX_JIKI_CHECK_RANK - 1)
    '磁器ﾁｪｯｸﾗﾝｸXTop
    commonInfo.JikiCheckRankXtop = var(1, headerEnum.CSV_IDX_JIKI_CHECK_RANK_XTOP - 1)
    '磁器ﾁｪｯｸﾗﾝｸXMiddle
    commonInfo.JikiCheckRankXmiddle = var(1, headerEnum.CSV_IDX_JIKI_CHECK_RANK_XMIDDLE - 1)
    '磁器ﾁｪｯｸﾗﾝｸXBottom
    commonInfo.JikiCheckRankXbottom = var(1, headerEnum.CSV_IDX_JIKI_CHECK_RANK_XBOTTOM - 1)
    '磁器ﾁｪｯｸﾗﾝｸYRight
    commonInfo.JikiCheckRankYright = var(1, headerEnum.CSV_IDX_JIKI_CHECK_RANK_YRIGHT - 1)
    '磁器ﾁｪｯｸﾗﾝｸYMiddle
    commonInfo.JikiCheckRankYmiddle = var(1, headerEnum.CSV_IDX_JIKI_CHECK_RANK_YMIDDLE - 1)
    '磁器ﾁｪｯｸﾗﾝｸYLeft
    commonInfo.JikiCheckRankYleft = var(1, headerEnum.CSV_IDX_JIKI_CHECK_RANK_YLEFT - 1)
    '物流発送日
    commonInfo.ButsuryuHassoDate = var(1, headerEnum.CSV_IDX_BUTSURYU_HASSO_DATE - 1)
    '取引先名
    commonInfo.TorihikisakiSeishikiNameJyuchuzan = var(1, headerEnum.CSV_IDX_TORIHIKISAKI_SEISHIKI_NAME_JYUCHUZAN - 1)

    '戻り値
    Set ConvertCommonInfo = commonInfo

End Function


'-------------------------------------------------------------------------------
' メソッド      ConvertDetailInfo
' 機能          配列をクラスに設定する(FSJD006 検査表出力ﾃﾞｰﾀ(明細))
' 機能説明      引数の配列をグローバル変数のクラス(Cls_Detail)に設定する
'-------------------------------------------------------------------------------
Function ConvertDetailInfo(var, i) As Cls_Detail

    Dim detailInfo As Cls_Detail
    Set detailInfo = New Cls_Detail

    '測定項目区分
    detailInfo.itemKbn = var(i, detailEnum.CSV_IDX_ITEM_KBN - 1)
    '測定項目(ｱｲﾃﾑｺｰﾄﾞ)
    detailInfo.ItemCd = var(i, detailEnum.CSV_IDX_ITEM_CD - 1)
    '測定項目名
    detailInfo.ItemName = var(i, detailEnum.CSV_IDX_ITEM_NAME - 1)
    '項目順
    detailInfo.InsatsuNo = var(i, detailEnum.CSV_IDX_INSATSU_NO - 1)
    '規格補足
    detailInfo.KikakuHosoku = var(i, detailEnum.CSV_IDX_KIKAKU_HOSOKU - 1)
    '単位
    detailInfo.TaniKensa = var(i, detailEnum.CSV_IDX_TANI_KENSA - 1)
    '最大値
    detailInfo.MaxN = var(i, detailEnum.CSV_IDX_MAX_N - 1)
    '最小値
    detailInfo.MinN = var(i, detailEnum.CSV_IDX_MIN_N - 1)
    '平均値
    detailInfo.Ave = var(i, detailEnum.CSV_IDX_AVE - 1)
    '標準偏差
    detailInfo.Sigma = var(i, detailEnum.CSV_IDX_SIGMA - 1)
    'cpk
    detailInfo.Cpk = var(i, detailEnum.CSV_IDX_CPK - 1)
    '抜取数(不合格数)
    detailInfo.NukitoriSuFugoukakuSu = var(i, detailEnum.CSV_IDX_NUKITORI_SU_FUGOUKAKU_SU - 1)
    '抜取数(母数)
    detailInfo.NukitoriSuBosu = var(i, detailEnum.CSV_IDX_NUKITORI_SU_BOSU - 1)
    '判定
    detailInfo.Hantei = var(i, detailEnum.CSV_IDX_HANTEI - 1)
    '規格ﾀｲﾌﾟ
    detailInfo.KikakuHantei = var(i, detailEnum.CSV_IDX_KIKAKU_HANTEI - 1)
    '測定値
    detailInfo.SokuteiValue = var(i, detailEnum.CSV_IDX_SOKUTEI_Value - 1)
    'AQL(水)(検査表用)
    detailInfo.AqlSuijunKensa = var(i, detailEnum.CSV_IDX_AQL_SUIJUN_KENSA - 1)
    'USL
    detailInfo.Usl = var(i, detailEnum.CSV_IDX_USL - 1)
    'LSL
    detailInfo.Lsl = var(i, detailEnum.CSV_IDX_LSL - 1)
    'SPC
    detailInfo.Spc1 = var(i, detailEnum.CSV_IDX_SPC - 1)
    '規格基準値
    detailInfo.KikakuChushin = var(i, detailEnum.CSV_IDX_KIKAKU_CHUSHIN - 1)
    '規格上限差
    detailInfo.PlusKosa = var(i, detailEnum.CSV_IDX_PLUS_KOSA - 1)
    '規格下限差
    detailInfo.MinusKosa = var(i, detailEnum.CSV_IDX_MINUS_KOSA - 1)
    '測定装置
    detailInfo.SokuteiSouchi = var(i, detailEnum.CSV_IDX_SOKUTEI_SOUCHI - 1)
    '画像検査
    detailInfo.GazoKensa = var(i, detailEnum.CSV_IDX_GAZO_KENSA - 1)
    '出荷ﾛｯﾄ備考
    detailInfo.ShukkaLotBiko = var(i, detailEnum.CSV_IDX_SHUKKA_LOT_BIKO - 1)
    'AQL(抜)(検査表用)
    detailInfo.AqlNukiKensa = var(i, detailEnum.CSV_IDX_AQL_NUKI_KENSA - 1)
    '耐熱温度
    detailInfo.Ondo = var(i, detailEnum.CSV_IDX_ONDO - 1)
    '耐熱時間
    detailInfo.Jikan = var(i, detailEnum.CSV_IDX_JIKAN - 1)
    '耐熱時間単位
    detailInfo.JikanTani = var(i, detailEnum.CSV_IDX_JIKAN_TANI - 1)
    '耐熱方法
    detailInfo.Tainetsu = var(i, detailEnum.CSV_IDX_TAINETSU - 1)
    'AQL S/S
    detailInfo.AqlSs = var(i, detailEnum.CSV_IDX_AQL_SS - 1)
    'AQL A/R
    detailInfo.AqlAr = var(i, detailEnum.CSV_IDX_AQL_AR - 1)
    'MODE-C=0件数
    detailInfo.ModeCZeroCnt = var(i, detailEnum.CSV_IDX_MODE_C_ZERO_CNT - 1)
    'ﾃﾞｰﾀ表示方法
    detailInfo.DataHyoji = var(i, detailEnum.CSV_IDX_DATA_HYOJI - 1)
    '分類
    detailInfo.Bunrui = var(i, detailEnum.CSV_IDX_BUNRUI - 1)
    'MODE-C件数
    detailInfo.ModeCCnt = var(i, detailEnum.CSV_IDX_MODE_C_CNT - 1)

    '戻り値
    Set ConvertDetailInfo = detailInfo

End Function

'-------------------------------------------------------------------------------
' メソッド      ConvertMultiLotInfo
' 機能          配列をクラスに設定する(FSJD007 検査表出力ﾃﾞｰﾀ(複数ﾛｯﾄ))
' 機能説明      引数の配列をグローバル変数のクラス(Cls_MultiLot)に設定する
'-------------------------------------------------------------------------------
Function ConvertMultiLotInfo(var, i) As Cls_MultiLot

    Dim multiLotInfo As Cls_MultiLot
    Set multiLotInfo = New Cls_MultiLot

    '出荷ﾛｯﾄNo(明細)
    multiLotInfo.ShukkaLotNoMeisai = var(i, multiEnum.CSV_IDX_SHUKKA_LOT_NO_MEISAI - 1)
    '出荷数(明細)
    multiLotInfo.ShukkaSuMeisai = var(i, multiEnum.CSV_IDX_SHUKKA_SU_MEISAI - 1)
    'めっき日(明細)
    multiLotInfo.MekkiDateMeisai = var(i, multiEnum.CSV_IDX_MEKKI_DATE_MEISAI - 1)
    '電気工程実績日(明細)
    multiLotInfo.DenkiKoteiJissekiDateMeisai = var(i, multiEnum.CSV_IDX_DENKI_KOTEI_JISSEKI_DATE_MEISAI - 1)
    '外観工程実績計上日(明細)
    multiLotInfo.GaikanKoteiJissekiDateMeisai = var(i, multiEnum.CSV_IDX_GAIKAN_KOTEI_JISSEKI_DATE_MEISAI - 1)

    '戻り値
    Set ConvertMultiLotInfo = multiLotInfo

End Function

'-------------------------------------------------------------------------------
' メソッド      ExecSub
' 機能          ｻﾌﾞ帳票ﾌｫｰﾏｯﾄ取込
' 機能説明      ｻﾌﾞ帳票ﾌｫｰﾏｯﾄ取込
'-------------------------------------------------------------------------------
Private Sub ExecSub()

    Dim i As Long
    Dim subFolder As String
    Dim subFiles As Variant
    Dim subBook As Workbook
    Dim sheetName As String
    Dim resultArray As Variant

    subFolder = grb_targetWb.Worksheets(grb_constInfo.sheetNameParam).Range(grb_constInfo.paramSubPath).Value
    subFiles = grb_targetWb.Worksheets(grb_constInfo.sheetNameParam).Range(grb_constInfo.paramTargetFormat1).Resize(9, 2).Value
    
    'ｸﾗｽ生成
    Set grb_argInfo.paperSize = New Collection
    Set grb_argInfo.pageOrientation = New Collection

    For i = 1 To UBound(subFiles)
        If (Not subFiles(i, 1) = grb_constInfo.targetMain And Not IsEmpty(subFiles(i, 1))) Then
            Set subBook = Workbooks.Open(subFolder & "\" & subFiles(i, 1))
            If (i = 1) Then
                subBook.Worksheets(subFiles(i, 2)).Copy After:=grb_targetWb.Worksheets(grb_targetWb.Worksheets.Count)
            ElseIf (IsEmpty(subFiles(i - 1, 2))) Then
                subBook.Worksheets(subFiles(i, 2)).Copy After:=grb_targetWb.Worksheets(subFiles(i - 2, 2))
            Else
            '1：ファイル名　2：シート名
            subBook.Worksheets(subFiles(i, 2)).Copy After:=grb_targetWb.Worksheets(subFiles(i - 1, 2))
            End If
            resultArray = Application.Run(subFiles(i, 1) & "!" & "SUB_MODULE.SubExec", grb_targetWb, grb_csvInfo)
            
            If resultArray(0) = "True" Then
                Err.Raise Number:=-1, Description:="ｴﾗｰ内容"
            End If
            
            'ｻﾌﾞﾌｫｰﾏｯﾄの印刷設定を保持する
            grb_argInfo.paperSize.Add (resultArray(1))
            grb_argInfo.pageOrientation.Add (resultArray(2))
            Call subBook.Close
        ElseIf (Not subFiles(i, 1) = "" Or IsEmpty(subFiles(i, 1))) Then
            'ﾒｲﾝの検査表の印刷設定を保持する
            grb_argInfo.paperSize.Add (grb_constInfo.paperSize)
            grb_argInfo.pageOrientation.Add (grb_constInfo.pageOrientation)
        End If
    Next i


End Sub

'-------------------------------------------------------------------------------
' メソッド      ExecMain
' 機能          帳票を作成する
' 機能説明　　  CSVﾌｧｲﾙを元に帳票を作成する
'-------------------------------------------------------------------------------
Public Function ExecMain() As String

    Dim nowDisplayAlerts As Boolean
    Dim i As Long
    Dim sheetNameResult As Variant
    Dim afterSheetName As String    'ｺﾋﾟｰ先指定に利用
    Dim beforeSheetName As String   'ｺﾋﾟｰ先指定に利用
    Dim detailStartRow As Long
    Dim longArray(10000) As String


    '警告非表示
    nowDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    '記述処理
    If Application.Run(grb_constInfo.formLink & "!MAIN2.writeElement", grb_csvInfo, grb_argInfo) Then
        Err.Raise Number:=-1, Description:="ｴﾗｰ内容"
    End If
    
    'ｺﾋﾟｰ後、前になるｼｰﾄ名取得
    afterSheetName = SearchKensahyo.Offset(1, 1).Value
    'ｺﾋﾟｰ後、後になるｼｰﾄ名取得
    beforeSheetName = SearchKensahyo.Offset(-1, 1).Value
    
    'ﾒｲﾝｼｰﾄ名取得
    sheetNameResult = Application.Run(grb_constInfo.formLink & "!MainSheetNameCreate", grb_csvInfo)
    If sheetNameResult(0) = "True" Then
        Err.Raise Number:=-1, Description:="ｴﾗｰ内容"
    End If
    grb_argInfo.mainSheetName = sheetNameResult(1)
    
    '1. 非印刷ﾒｲﾝｼｰﾄを検査表印刷用ﾒｲﾝｼｰﾄへｺﾋﾟｰする(ﾊﾟﾗﾒｰﾀの読込対象ﾌｧｲﾙの出力順番を考慮する)。
    If (beforeSheetName = "ｼｰﾄ名") Then
        grb_targetWb.Worksheets(grb_constInfo.sheetNameMain).Copy Before:=Sheets(afterSheetName)
        grb_targetWb.ActiveSheet.Name = grb_argInfo.mainSheetName
    Else
        grb_targetWb.Worksheets(grb_constInfo.sheetNameMain).Copy After:=Sheets(beforeSheetName)
        grb_targetWb.ActiveSheet.Name = grb_argInfo.mainSheetName
    End If
    
    '測定項目縦並び時のｹｰｽ
    If (grb_constInfo.detailDirection) Then
        '総行数の保持
        grb_argInfo.allRowsCount = _
            grb_targetWb.Worksheets(grb_constInfo.sheetNameMain).Range(grb_constInfo.colItem & CStr(grb_targetWb.Worksheets(grb_constInfo.sheetNameParam).Range(grb_constInfo.paramFooterStart).Value)).End(xlUp).Row _
            - grb_targetWb.Worksheets(grb_constInfo.sheetNameMain).Range(grb_constInfo.colItem & CStr(grb_constInfo.detailStartRow)).Row + 1
    End If

End Function

'-------------------------------------------------------------------------------
' メソッド      searchKensahyo
' 機能          ﾒｲﾝｼｰﾄのRangeを返す
' 機能説明　　  ﾊﾟﾗﾒｰﾀｼｰﾄに有る読み込み対象ﾌｧｲﾙのﾌｧｲﾙ名を検索し、ｼｰﾄ名のRangeを返す
'-------------------------------------------------------------------------------
Function SearchKensahyo() As Range

    Set SearchKensahyo = _
            grb_targetWb.Worksheets(grb_constInfo.sheetNameParam).Range(grb_constInfo.paramTargetFormatRange).Find( _
            What:=grb_constInfo.targetMain, _
            LookIn:=xlValues, _
            lookat:=xlWhole, _
            SearchOrder:=xlByColumns, _
            MatchByte:=False)

End Function

'-------------------------------------------------------------------------------
' メソッド      PageBreak
' 機能          改ページ処理
' 機能説明　　  改ページ処理
'-------------------------------------------------------------------------------
Sub MainPageBreak()
    Dim paramWs As Worksheet        'ﾊﾟﾗﾒｰﾀｼｰﾄObj
    Dim editWS As Worksheet         '編集対象ｼｰﾄObj
    Dim detailStartCell As Range    '明細開始セル
    Dim detailRange As Range        '明細範囲
    Dim emptyStartCell As Range     '記入されなかった明細部の開始セル
    Dim emptyRange As Range         '空白範囲
    Dim headerRows As Long          'ﾍｯﾀﾞｰ行数
    Dim footerRows As Long          'ﾌｯﾀｰ行数
    Dim currentRow As Long          '現編集行数
    Dim breakCount As Long          '改行回数
    Dim currentRowsHeight As Long    '行高さの一時保存用
    Dim pageHightLimit As Long      '1ページの最大pt数
    Dim defaultHeight As Long       '標準の高さ(pt)
    Dim afterSheetName As String    'ｺﾋﾟｰ先指定に利用
    Dim beforeSheetName As String   'ｺﾋﾟｰ先指定に利用
    Dim allPages As Long            '総ページ数
    Dim i As Long
    Dim subtitleAjustCount As Long  '分類位置調整回数

    'ﾊﾟﾗﾒｰﾀｼｰﾄObj取得
    Set paramWs = grb_targetWb.Worksheets(grb_constInfo.sheetNameParam)
    '編集対象ワークシート取得
    Set editWS = grb_targetWb.Worksheets(grb_argInfo.mainSheetName)
    'ﾍｯﾀﾞｰ行数取得
    headerRows = paramWs.Range(grb_constInfo.paramHeaderEnd).Value - paramWs.Range(grb_constInfo.paramHeaderStart).Value + 1
    'ﾌｯﾀｰ行数取得
    footerRows = paramWs.Range(grb_constInfo.paramFooterEnd).Value - paramWs.Range(grb_constInfo.paramFooterStart).Value + 1

    '標準の行の高さを取得
    defaultHeight = paramWs.Range(grb_constInfo.paramRow).Value
    
    '測定項目縦並び時のｹｰｽ
    If (grb_constInfo.detailDirection) Then
        '未入力行の削除
        Set emptyStartCell = editWS.Range("B" & CStr(grb_argInfo.allRowsCount + grb_constInfo.detailStartRow))
        Set emptyRange = Range(emptyStartCell, editWS.Range(grb_constInfo.colItem & CStr(grb_targetWb.Worksheets(grb_constInfo.sheetNameParam).Range(grb_constInfo.paramFooterStart).Value)).Offset(-1, 0))
        emptyRange.EntireRow.Delete
    End If

    Dim paramColumnLimit As Long
    paramColumnLimit = paramWs.Range(grb_constInfo.paramColumnLimit).Value
    'ﾌｫｰﾏｯﾄ行のｺﾋﾟｰ元開始行
    Dim formatCopySourceStartRow As Long
    'ﾌｫｰﾏｯﾄ行のｺﾋﾟｰ元開始列
    Dim formatCopySourceStartColumn As Long
    'ﾌｫｰﾏｯﾄ行のｺﾋﾟｰ元終了行
    Dim formatCopySourceEndRow As Long
    'ﾌｫｰﾏｯﾄ行のｺﾋﾟｰ元終了列
    Dim formatCopySourceEndColumn As Long

    '測定項目横並び時のｹｰｽ
    If Not (grb_constInfo.detailDirection) Then
        'ﾌｫｰﾏｯﾄ行のｺﾋﾟｰ元開始行：ﾊﾟﾗﾒｰﾀｼｰﾄ.ﾍｯﾀﾞｰの開始行
        formatCopySourceStartRow = paramWs.Range(grb_constInfo.paramHeaderStart).Value
        'ﾌｫｰﾏｯﾄ行のｺﾋﾟｰ元開始列：ﾍｯﾀﾞｰの開始列（ﾊﾟﾗﾒｰﾀｼｰﾄにはない項目）
        formatCopySourceStartColumn = grb_constInfo.heaederStartColumn
        'ﾌｫｰﾏｯﾄ行のｺﾋﾟｰ元終了行：ﾊﾟﾗﾒｰﾀｼｰﾄ.ﾌｯﾀｰの終了行
        formatCopySourceEndRow = paramWs.Range(grb_constInfo.paramFooterEnd).Value
        'ﾌｫｰﾏｯﾄ行のｺﾋﾟｰ元終了列：ﾌｯﾀｰの終了列（ﾊﾟﾗﾒｰﾀｼｰﾄにはない項目）
        formatCopySourceEndColumn = grb_constInfo.footerEndColumn
        
        Dim endRow As Long
        If grb_csvInfo.DetailInfoList.Count <= paramColumnLimit Then
            endRow = formatCopySourceEndColumn
        Else
            endRow = formatCopySourceEndColumn + (grb_csvInfo.DetailInfoList.Count - paramColumnLimit)
        End If
        
        editWS.PageSetup.PrintArea = editWS.Range( _
            editWS.Cells(formatCopySourceStartRow, formatCopySourceStartColumn), _
            editWS.Cells(formatCopySourceEndRow, endRow)).Address
    End If

    If (IsCopy) Then
        'ｼｰﾄｺﾋﾟｰ
        editWS.Copy After:=editWS
    End If

    '測定項目縦並び時のｹｰｽ
    If (grb_constInfo.detailDirection) Then
        '2. 改ﾍﾟｰｼﾞあり検査表の場合、1ﾍﾟｰｼﾞに収まるかﾁｪｯｸする。
        '取得：1ページの最大pt数
        pageHightLimit = paramWs.Range(grb_constInfo.paramRowHeight)

        '2-1. 明細行の合計の高さを算出する。
        Set detailStartCell = editWS.Range(grb_constInfo.colItem & CStr(grb_constInfo.detailStartRow))
        Set detailRange = Range(detailStartCell, editWS.Range(grb_constInfo.colItem & CStr(grb_argInfo.endRows)))
        '2-2. 明細行の高さが1ﾍﾟｰｼﾞに収まるかの判断をする。
        If (detailRange.Height <= pageHightLimit) Then
            currentRowsHeight = detailRange.Height
        Else
            '3. 印刷種別が「直接印刷 or PDF出力」 and 2-2で改ﾍﾟｰｼﾞが必要と判断した場合、
            '3-1. 明細の先頭から高さを計算し、改ﾍﾟｰｼﾞを入れる行を探す
            '各行から「ﾊﾟﾗﾒｰﾀ.詰めｻｲｽﾞ」を引いたｻｲｽﾞの明細行の高さ合計を計算する(ここではｾﾙ内改行は考慮しない)。
            If (detailRange.Height - (detailRange.Rows.Count * CLng(paramWs.Range(grb_constInfo.paramPaddingSize).Value))) <= pageHightLimit Then
                '1ページに収まるので、各行のサイズを詰める
                For i = 0 To detailRange.Rows.Count - 1
                        editWS.Range(grb_constInfo.colItem & CStr(grb_constInfo.detailStartRow + i)).RowHeight = editWS.Range(grb_constInfo.colItem & CStr(grb_constInfo.detailStartRow + i)).Height - paramWs.Range(grb_constInfo.paramPaddingSize).Value
                Next i
            currentRowsHeight = detailRange.Height
            '1ページに収まらない場合
            Else
                '値の初期化
                breakCount = 0
                subtitleAjustCount = 0
                currentRowsHeight = 0
                'emptyLine = 0
                '行数分ループ
                For i = 0 To grb_argInfo.allRowsCount - 1
                    '現在の行数を計算:明細開始行 + 何行目のﾃﾞｰﾀ + (改行回数 * (ヘッダ総行数+フッタ総行数)
                    currentRow = grb_constInfo.detailStartRow + i + subtitleAjustCount + (breakCount * (headerRows + footerRows))
                    '3-1-1. 明細の高さ合計　＝　明細の高さの合計　＋　ｶﾚﾝﾄ行の高さ
                    currentRowsHeight = currentRowsHeight + editWS.Range(grb_constInfo.colItem & CStr(currentRow)).Height
                    '3-1-2. 明細の高さ合計　＞　「ﾊﾟﾗﾒｰﾀ.明細として使用可能な高さ」 の場合
                    If (currentRowsHeight > pageHightLimit) Then
                        currentRowsHeight = editWS.Range(grb_constInfo.colItem & CStr(currentRow)).Height
                        '現行の区分を確認して分類の場合に行追加
                        If grb_csvInfo.DetailInfoList.Item(i).itemKbn = "1" Then
                            editWS.Rows(currentRow - 1).Insert
                            editWS.Rows(currentRow - 1).RowHeight = defaultHeight
                            subtitleAjustCount = subtitleAjustCount + 1
                        End If
                        Call PageBreakInsert(editWS, currentRow)
                        breakCount = breakCount + 1
                    End If
                '3-1-3. 次の行へ
                Next i
            End If
        End If
        
        '3. 最終ページの場合かつ、挿入後の高さが[変数]明細の高さ合計 ＜「ﾊﾟﾗﾒｰﾀ.明細として使用可能な高さ」の間、空白行を挿入する。
        While currentRowsHeight + defaultHeight < pageHightLimit
            currentRow = grb_constInfo.detailStartRow + i + subtitleAjustCount + (breakCount * (headerRows + footerRows))
            editWS.Rows(currentRow).Insert
            '行の書式を前行からｺﾋﾟｰ&ﾍﾟｰｽﾄ
            editWS.Rows(currentRow - 1).Copy
            editWS.Rows(currentRow).PasteSpecial (xlPasteFormats)
            Application.CutCopyMode = False
            '行の高さを標準に設定
            If (Not editWS.Rows(currentRow - 1).RowHeight = defaultHeight) Then
                editWS.Rows(currentRow).RowHeight = defaultHeight
            End If
            '明細の高さ合計更新
            currentRowsHeight = currentRowsHeight + editWS.Rows(currentRow).RowHeight
            i = i + 1
        Wend

    Else
        '測定項目よこ並び時のｹｰｽ
        Dim totalPage As Long
        totalPage = (grb_csvInfo.DetailInfoList.Count \ paramColumnLimit) + (grb_csvInfo.DetailInfoList.Count Mod paramColumnLimit)

        'ﾌｫｰﾏｯﾄ行のｺﾋﾟｰ先開始行
        Dim formatCopyDestinationStartRow As Long
        'ﾌｫｰﾏｯﾄ行のｺﾋﾟｰ先開始列
        Dim formatCopyDestinationStartColumn As Long
        'ﾌｫｰﾏｯﾄ行のｺﾋﾟｰ先終了行
        Dim formatCopyDestinationEndRow As Long
        'ﾌｫｰﾏｯﾄ行のｺﾋﾟｰ先終了列
        Dim formatCopyDestinationEndColumn As Long
        
        formatCopyDestinationStartRow = formatCopySourceStartRow
        formatCopyDestinationStartColumn = formatCopySourceStartColumn
        formatCopyDestinationEndRow = formatCopySourceEndRow
        formatCopyDestinationEndColumn = formatCopySourceEndColumn

        '明細行のｺﾋﾟｰ元開始行
        Dim detailCopySourceStartRow As Long
        '明細行のｺﾋﾟｰ元開始列
        Dim detailCopySourceStartColumn As Long
        '明細行のｺﾋﾟｰ元終了行
        Dim detailCopySourceEndRow As Long
        '明細行のｺﾋﾟｰ元終了列
        Dim detailCopySourceEndColumn As Long

        '明細行のｺﾋﾟｰ先開始行
        Dim detailCopyDestinationStartRow As Long
        '明細行のｺﾋﾟｰ先開始列
        Dim detailCopyDestinationStartColumn As Long
        '明細行のｺﾋﾟｰ先終了行
        Dim detailCopyDestinationEndRow As Long
        '明細行のｺﾋﾟｰ先終了列
        Dim detailCopyDestinationEndColumn As Long
        
        '明細行のｺﾋﾟｰ元開始行：ﾊﾟﾗﾒｰﾀｼｰﾄ.ﾍｯﾀﾞｰの終了行 + 1
        detailCopySourceStartRow = paramWs.Range(grb_constInfo.paramHeaderEnd).Value + 1
        '明細行のｺﾋﾟｰ元開始列
        detailCopySourceStartColumn = 0
        '明細行のｺﾋﾟｰ元終了行：ﾊﾟﾗﾒｰﾀｼｰﾄ.ﾌｯﾀｰの終了行 - 1
        detailCopySourceEndRow = paramWs.Range(grb_constInfo.paramFooterStart).Value - 1
        '明細行のｺﾋﾟｰ元終了列
        detailCopySourceEndColumn = formatCopySourceEndColumn
        
        '明細行のｺﾋﾟｰ先開始行
        detailCopyDestinationStartRow = detailCopySourceStartRow
        '明細行のｺﾋﾟｰ先開始列：明細ﾃﾞｰﾀの開始列（ﾊﾟﾗﾒｰﾀｼｰﾄにはない項目）
        detailCopyDestinationStartColumn = grb_constInfo.detailDataStartColumn
        '明細行のｺﾋﾟｰ先終了行
        detailCopyDestinationEndRow = detailCopySourceEndRow
        '明細行のｺﾋﾟｰ先終了列
        detailCopyDestinationEndColumn = formatCopySourceEndColumn
        
        For i = 1 To totalPage
            '1ページ目は処理しない
            If i <> 1 Then
                '1ページの行数をたす
                formatCopyDestinationStartRow = formatCopyDestinationStartRow + formatCopySourceEndRow
                formatCopyDestinationEndRow = formatCopyDestinationEndRow + formatCopySourceEndRow
    
                '1ページ目のフォーマットを次のページへコピーする
                Call editWS.Range(editWS.Cells(formatCopySourceStartRow, formatCopySourceStartColumn), editWS.Cells(formatCopySourceEndRow, formatCopySourceEndColumn)) _
                    .Copy(editWS.Range(editWS.Cells(formatCopyDestinationStartRow, formatCopyDestinationStartColumn), editWS.Cells(formatCopyDestinationEndRow, formatCopyDestinationEndColumn)))
    
                '明細エリアのコピー
                Dim detailArray As Variant
                detailCopySourceStartColumn = detailCopySourceEndColumn + 1
                detailCopySourceEndColumn = detailCopySourceStartColumn + paramColumnLimit - 1
                detailArray = editWS.Range(editWS.Cells(detailCopySourceStartRow, detailCopySourceStartColumn), editWS.Cells(detailCopySourceEndRow, detailCopySourceEndColumn))
    
                detailCopyDestinationStartRow = detailCopyDestinationStartRow + formatCopySourceEndRow
                detailCopyDestinationEndRow = detailCopyDestinationEndRow + formatCopySourceEndRow
                editWS.Range(editWS.Cells(detailCopyDestinationStartRow, detailCopyDestinationStartColumn), editWS.Cells(detailCopyDestinationEndRow, detailCopyDestinationEndColumn)) = detailArray
            End If
        Next i

        editWS.PageSetup.PrintArea = editWS.Range(editWS.Cells(formatCopySourceStartRow, formatCopySourceStartColumn), editWS.Cells(formatCopyDestinationEndRow, formatCopyDestinationEndColumn)).Address
        DoEvents
        Dim pageBreakRow As Long
        For i = 1 To totalPage - 1
            pageBreakRow = (formatCopySourceEndRow * i) + 1
            editWS.Rows(pageBreakRow).PageBreak = xlPageBreakManual
        Next i
    End If
    
End Sub
    
'-------------------------------------------------------------------------------
' メソッド      IsCopy
' 機能          ﾒｲﾝｼｰﾄが2数あるかをBooleanで返す
' 機能説明　　  印刷用ﾒｲﾝｼｰﾄが2つ存在するかのﾁｪｯｸ用
'-------------------------------------------------------------------------------
Function IsCopy() As Boolean
    IsCopy = (grb_csvInfo.commonInfo.InsatsuType0 = "1" _
    Or grb_csvInfo.commonInfo.InsatsuType1 = "1" _
    Or grb_csvInfo.commonInfo.InsatsuType2 = "1") _
    And _
    (grb_csvInfo.commonInfo.InsatsuType3 = "1" _
    Or grb_csvInfo.commonInfo.InsatsuType5 = "1")

End Function

'-------------------------------------------------------------------------------
' メソッド      PageBreakInsert
' 機能          ﾌｯﾀｰ・改ページ・ﾍｯﾀﾞｰの挿入
' 機能説明　　  ﾌｯﾀｰ・改ページ・ﾍｯﾀﾞｰの挿入
'-------------------------------------------------------------------------------
Sub PageBreakInsert(ByVal ws As Worksheet, ByVal currentRow As Long)

    Dim headerStartRow As Long
    Dim headerEndRow As Long
    Dim headerRows As Long
    Dim footerStartRow As Long
    Dim footerEndRow As Long
    Dim footerRows As Long
    Dim copyWs As Worksheet
    Dim paramWs As Worksheet

    Set copyWs = grb_targetWb.Worksheets(grb_constInfo.sheetNameMain)
    Set paramWs = grb_targetWb.Worksheets(grb_constInfo.sheetNameParam)
    headerStartRow = paramWs.Range(grb_constInfo.paramHeaderStart).Value
    headerEndRow = paramWs.Range(grb_constInfo.paramHeaderEnd).Value
    headerRows = headerEndRow - headerStartRow + 1
    footerStartRow = paramWs.Range(grb_constInfo.paramFooterStart).Value
    footerEndRow = paramWs.Range(grb_constInfo.paramFooterEnd).Value
    footerRows = footerEndRow - footerStartRow + 1

    'ﾌｯﾀｰ挿入
    Range(copyWs.Rows(footerStartRow), copyWs.Rows(footerEndRow)).Copy
    Range(ws.Rows(currentRow), ws.Rows(currentRow + footerRows)).Insert
    Application.CutCopyMode = False
    'ﾍｯﾀﾞｰ挿入
    Range(copyWs.Rows(headerStartRow), copyWs.Rows(headerEndRow)).Copy
    Range(ws.Rows(currentRow + footerRows), ws.Rows(currentRow + footerRows + headerRows)).Insert
    Application.CutCopyMode = False
    '改ﾍﾟｰｼﾞ組み込み
    ws.Rows(currentRow + footerRows).PageBreak = xlPageBreakManual

End Sub


'-------------------------------------------------------------------------------
' メソッド      SaveBook
' 機能          ｴﾋﾞﾃﾞﾝｽ用ｴｸｾﾙ保存
' 機能説明      1. 「ﾊﾟﾗﾒｰﾀ.ｴﾋﾞﾃﾞﾝｽ保存場所」に、「ﾊﾟﾗﾒｰﾀ.出力ﾌｧｲﾙ名」でﾏｸﾛﾌｧｲﾙを保存する。
'-------------------------------------------------------------------------------
Private Sub SaveBook()

    Dim full_path As String  '作成するフォルダーのフルパス
    full_path = MakeDir
    
    '出力ファイル名取得
    GetOutputFileName
    
    'エクセル
    grb_targetWb.SaveAs (CreateNewFilePath(full_path & "\" & grb_argInfo.outputFileName & ".xlsm"))

    'FSJD005
    Call BackUpFileCopy(full_path, "FSJD005.CSV")
    'FSJD006
    Call BackUpFileCopy(full_path, "FSJD006.CSV")
    'FSJD007
    Call BackUpFileCopy(full_path, "FSJD007.CSV")

End Sub

'-------------------------------------------------------------------------------
' メソッド      MakeDir
' 機能          ディレクトリ作成
' 機能説明　　  エビデンス保存先のディレクトリを作成する
'-------------------------------------------------------------------------------
Function MakeDir() As String

    Dim full_path As String     '作成するフォルダーのフルパス

    full_path = grb_targetWb.Worksheets(grb_constInfo.sheetNameParam).Range(grb_constInfo.paramEviPath).Value
    full_path = full_path & "\" & Format(Date, "yyyy")
    full_path = full_path & "\" & Format(Date, "mm")
    full_path = full_path & "\" & Format(Date, "dd")

    MakeDir = CommonMakeDir(full_path)

End Function

'-------------------------------------------------------------------------------
' メソッド      BackUpFileCopy
' 機能          バックファイルコピー
' 機能説明　　  CSVファイルを保存先にバックファイルのコピーをする
'-------------------------------------------------------------------------------
Private Sub BackUpFileCopy(dirPath As String, fileName As String)
    Dim sourceFile As String    'コピー元のファイルパス
    Dim backUpFile As String    'コピー先のファイルパス
    
    'CSV保存場所のフルパス
    sourceFile = grb_targetWb.Worksheets(grb_constInfo.sheetNameParam).Range(grb_constInfo.paramCsvPath).Value
    sourceFile = sourceFile & "\" & fileName
    
    If Dir(sourceFile) = "" Then
        'コピー元の対象ファイルが存在しない場合、処理を終了する
        Debug.Print "sourceFile:" & sourceFile
        Exit Sub
    End If
    
    '保存先のフルパス
    backUpFile = dirPath & "\" & grb_argInfo.outputFileName & fileName
    backUpFile = CreateNewFilePath(backUpFile)
    
    FileCopy sourceFile, backUpFile

End Sub

'-------------------------------------------------------------------------------
' メソッド      CreateNewFilePath
' 機能          新規ファイルパス生成
' 機能説明　　  同一ファイル名のファイルが存在する場合はファイル名に"_"+連番値を付与する
'               ※保存ファイルについて保存先に同一ファイル名のファイルが存在する場合はファイル名に"_"+連番値を付与すること。
'               (ﾌｧｲﾙ重複がなくなるまで連番値をｶｳﾝﾄｱｯﾌﾟすること。)
'-------------------------------------------------------------------------------
Function CreateNewFilePath(ByVal filePath As String) As String
    Dim newFilePath As String '新ファイルパス格納用変数
    
    If (Dir(filePath) = "") Then
        '同名ファイルが存在しない場合
        newFilePath = filePath
    Else
        '同名ファイルが存在する場合
        '拡張子と拡張子を除いたファイルパス取得
        Dim extensionPosition As Long           '拡張子の位置
        Dim exceptExtensionFilePath As String   '拡張子を除いたファイルパス格納用変数
        Dim extension As String                 '拡張子格納用変数
        Dim i As Long
        Dim num As String
        
        extensionPosition = InStrRev(filePath, ".")
    
        If (0 < extensionPosition) Then
            extension = Right(filePath, Len(filePath) - extensionPosition)
            exceptExtensionFilePath = Left(filePath, extensionPosition - 1)
        Else
            extension = ""
            exceptExtensionFilePath = filePath
        End If
        
        '連番文字列を生成
        i = 1
        num = i
        
        '連番付きの新しいパスを生成
        newFilePath = exceptExtensionFilePath & "_" & num & "." & extension
        
        '同名のの連番付ファイル名が存在しなくなるまでループ
        Do While (Dir(newFilePath) <> "")
            i = i + 1
            num = i
            newFilePath = exceptExtensionFilePath & "_" & num & "." & extension
        Loop
    End If
    
    CreateNewFilePath = newFilePath
End Function

'-------------------------------------------------------------------------------
' メソッド      GetOutputFileName
' 機能          出力ファイル名取得
' 機能説明　　  出力ファイル名を取得する
'-------------------------------------------------------------------------------
Private Sub GetOutputFileName()
    Dim param As String     'パラメーター
    
    '「ﾊﾟﾗﾒｰﾀ.出力ﾌｧｲﾙ名」
    param = grb_targetWb.Worksheets(grb_constInfo.sheetNameParam).Range(grb_constInfo.paramOutFilename).Value
    param = Split(param, ":")(0) ' 出力ファイル名の区分値取得
    
    'CSV1から出荷日、製伝No、出荷ﾛｯﾄを取得する
    Select Case param
        Case "0"
            '0:出荷日(yyyymmdd)+製伝No+出荷ﾛｯﾄ+出荷数
            grb_argInfo.outputFileName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Format(grb_csvInfo.commonInfo.ShukkaDate, "yyyymmdd") & grb_csvInfo.commonInfo.SeidenNo & grb_csvInfo.commonInfo.ShukkaLotNo & grb_csvInfo.commonInfo.ShukkaSu, _
                                        "\", " "), "/", " "), ":", " "), "*", " "), "?", " "), """", " "), "<", " "), ">", " "), "|", " ")

        Case "1"
            '1:検査表作成日(yyyymmdd)+製伝No+出荷ﾛｯﾄ+出荷数
            grb_argInfo.outputFileName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace( _
                                        Format(Date, "yyyymmdd") & grb_csvInfo.commonInfo.SeidenNo & grb_csvInfo.commonInfo.ShukkaLotNo & grb_csvInfo.commonInfo.ShukkaSu, _
                                        "\", " "), "/", " "), ":", " "), "*", " "), "?", " "), """", " "), "<", " "), ">", " "), "|", " ")
        Case Else
    End Select

End Sub

'-------------------------------------------------------------------------------
' メソッド      DeleteSheets
' 機能          不要ｼｰﾄ削除
' 機能説明      1. 非提出用ｼｰﾄを削除する。
'-------------------------------------------------------------------------------
Private Sub DeleteSheets()

    Dim targetSheetName As String   '1番目出力のｼｰﾄ名
    Dim lastSheetName As String     '最終ｼｰﾄ名
    Dim objSheet As Worksheet       'ｵﾌﾞｼﾞｪｸﾄｼｰﾄ
    
        
    '「ﾊﾟﾗﾒｰﾀ.出力部数」取得
    grb_argInfo.copies = grb_targetWb.Worksheets(grb_constInfo.sheetNameParam).Range(grb_constInfo.paramOutCopies).Value
    
    '「ﾊﾟﾗﾒｰﾀ.読込対象ﾌｧｲﾙ」の1番目出力のｼｰﾄ名
    targetSheetName = grb_targetWb.Worksheets(grb_constInfo.sheetNameParam).Range(grb_constInfo.paramTargetFormat1).Offset(, 1).Value
    'ｼｰﾄ名が""の場合、ﾒｲﾝｼｰﾄ名を取得
    If (targetSheetName = "") Then
        targetSheetName = grb_argInfo.mainSheetName
    End If
    
    '「ﾊﾟﾗﾒｰﾀ.読込対象ﾌｧｲﾙ」の先頭に定義しているｼｰﾄより左のｼｰﾄを不要ｼｰﾄとして削除対象とする。
    lastSheetName = Worksheets(Worksheets.Count).Name
    For Each objSheet In grb_targetWb.Worksheets
        If (objSheet.Name = targetSheetName _
        Or objSheet.Name = lastSheetName) Then
            '提出用ｼｰﾄの場合、ｼｰﾄ削除処理を終了する
            Exit For
        End If
        
        '非提出用ｼｰﾄを削除する。
        grb_targetWb.Worksheets(objSheet.Name).Delete
    Next

End Sub

'-------------------------------------------------------------------------------
' メソッド      PrintOut
' 機能          印刷実行
' 機能説明      1. CSV1「印刷種別:XXXX」の設定内容に応じて、以下の処理を実行する。
'-------------------------------------------------------------------------------
Private Sub PrintOut()
    
    If grb_csvInfo.commonInfo.InsatsuType0 = "1" Then
        '印刷種別:直接印刷
        DirectPrintOut
    End If
    
    If grb_csvInfo.commonInfo.InsatsuType1 = "1" Then
        '印刷種別:PDF
        CreatePdf
    End If
    
    If grb_csvInfo.commonInfo.InsatsuType3 = "1" Then
        '印刷種別:CSV
        CreateCsv
    End If
    
    If grb_csvInfo.commonInfo.InsatsuType5 = "1" Then
        '印刷種別:TSV
        CreateTsv
    End If
   
    If grb_csvInfo.commonInfo.InsatsuType2 = "1" Then
        '印刷種別:Excel
        CreateExcel
    End If
   
End Sub

'-------------------------------------------------------------------------------
' メソッド      DirectPrintOut
' 機能          直接印刷
' 機能説明　　  印刷対象(ﾊﾟﾗﾒｰﾀ.読込対象ﾌｧｲﾙ)を標準ﾌﾟﾘﾝﾀで「ﾊﾟﾗﾒｰﾀ.出力部数」分印刷実行する。
'-------------------------------------------------------------------------------
Private Sub DirectPrintOut()
    Dim targetSheets As Variant '直接印刷対象シート
    Dim i As Long
    
    '印刷対象シートを抽出
    targetSheets = GetTargetPrintSheets
    
    If Not IsEmpty(targetSheets) Then
        

'印刷種別:直接印刷時の「用紙サイズ」、「印刷の向き（縦・横）」（SJD-1732対応）
        For i = 0 To UBound(targetSheets)
            
            Sheets(targetSheets(i)).PageSetup.Orientation = grb_argInfo.pageOrientation.Item(i + 1) '印刷の向き
            Sheets(targetSheets(i)).PageSetup.paperSize = grb_argInfo.paperSize.Item(i + 1)         '用紙サイズ
       
        Next i

        'デフォルトプリンターに直接印刷
        Sheets(targetSheets).PrintOut copies:=grb_argInfo.copies, Preview:=False
    
    End If

End Sub

'-------------------------------------------------------------------------------
' メソッド      CreatePdf
' 機能          PDF作成
' 機能説明　　  印刷対象(ﾊﾟﾗﾒｰﾀ.読込対象ﾌｧｲﾙ)をﾊﾟﾗﾒｰﾀ,ﾌﾟﾘﾝﾀ名で、CSV1.発行部数分印刷実行する(1ﾌｧｲﾙに全て含める)。
'               ※保存先に同一ファイル名のファイルが存在する場合は上書きをすること｡
'               ※最小サイズ設定でPDF変換し､保存すること｡
'               出力先DIR          CSV1:出力先(PDF)
'               出力先ファイル名   ﾊﾟﾗﾒｰﾀ.出力ﾌｧｲﾙ名+".pdf"
'-------------------------------------------------------------------------------
Private Sub CreatePdf()
    Dim filePathPdf As String   '保存先フォルダパス
    Dim targetSheets As Variant 'PDF出力対象シート
    Dim i As Long
    
    'PDF出力対象シートを抽出
    targetSheets = GetTargetPrintSheets
    
    'PDF出力したいシートを選択する(複数)
    Worksheets(targetSheets).Select
    
    '保存先フォルダパス作成
    filePathPdf = MakeOutputDir(grb_csvInfo.commonInfo.Shutsuryokusaki1)
    
    'PDF出力
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=filePathPdf & "\" & grb_argInfo.outputFileName, _
        Quality:=xlQualityMinimum

End Sub

'-------------------------------------------------------------------------------
' メソッド      CreateExcel
' 機能          Excel作成
' 機能説明　　  ｴｸｾﾙ形式でﾌｧｲﾙを保存する。
'               ※保存先に同一ﾌｧｲﾙ名のﾌｧｲﾙが存在する場合は上書きをすること。
'               FileFormat:=xlOpenXMLWorkbook
'               出力先DIR      CSV1:出力先(Excel)
'               出力先ﾌｧｲﾙ名   ﾊﾟﾗﾒｰﾀ.出力ﾌｧｲﾙ名+".xlsx"   ※ﾏｸﾛは含めない
'-------------------------------------------------------------------------------
Private Sub CreateExcel()
    Dim filePathExcel As String '保存先フォルダパス
    
    '保存先フォルダパス作成
    filePathExcel = MakeOutputDir(grb_csvInfo.commonInfo.Shutsuryokusaki2)
    
    If (IsCopy) Then
        'PDF出力用ｼｰﾄを削除する。
        'grb_targetWb.Worksheets(grb_argInfo.mainSheetName).Delete
        grb_targetWb.Worksheets(MainSheetNameExcel).Delete
    End If
    
    'Excel作成
    grb_targetWb.SaveAs fileName:=filePathExcel & "\" & grb_argInfo.outputFileName, _
        FileFormat:=xlOpenXMLWorkbook

End Sub

'-------------------------------------------------------------------------------
' メソッド      CreateCsv
' 機能          Csv作成
' 機能説明　　  ﾒｲﾝｼｰﾄをCSV形式でﾌｧｲﾙを保存する。
'               For Append     ※「なければ新規、あれば追記」である必要がある
'               出力先DIR      CSV1:出力先(CSV)
'               出力先ﾌｧｲﾙ名   ﾊﾟﾗﾒｰﾀ.出力ﾌｧｲﾙ名+".csv"
'-------------------------------------------------------------------------------
Private Sub CreateCsv()
    Dim filePathCsv As String   '保存先フォルダパス
    
    '保存先フォルダパス作成
    filePathCsv = MakeOutputDir(grb_csvInfo.commonInfo.Shutsuryokusaki3)
    
    'CSV出力
    Call CreateTxt(filePathCsv, True)

End Sub

'-------------------------------------------------------------------------------
' メソッド      CreateTsv
' 機能          Tsv作成
' 機能説明　　  ﾒｲﾝｼｰﾄをﾀﾌﾞ区切りでﾌｧｲﾙを保存する。
'               For Append     ※「なければ新規、あれば追記」である必要がある
'               出力先DIR      CSV1:出力先(TSV)
'               出力先ﾌｧｲﾙ名   ﾊﾟﾗﾒｰﾀ.出力ﾌｧｲﾙ名+".tsv"
'-------------------------------------------------------------------------------
Private Sub CreateTsv()
    Dim filePathTsv As String   '保存先フォルダパス
    
    '保存先フォルダパス作成
    filePathTsv = MakeOutputDir(grb_csvInfo.commonInfo.Shutsuryokusaki5)
    
    'TSV出力
    Call CreateTxt(filePathTsv, False)

End Sub

'-------------------------------------------------------------------------------
' メソッド      MakeOutputDir
' 機能          ディレクトリ作成
' 機能説明　　  出力先のディレクトリを作成する
'-------------------------------------------------------------------------------
Function MakeOutputDir(dirPath As String) As String

    MakeOutputDir = CommonMakeDir(dirPath)

End Function

'-------------------------------------------------------------------------------
' メソッド      CommonMakeDir
' 機能          ディレクトリ作成（共通）
' 機能説明　　  ディレクトリを作成する
'-------------------------------------------------------------------------------
Function CommonMakeDir(path As String) As String

    Dim tmp_path As String      '一時的なパス
    Dim arr() As String         '配列
    Dim cnt As Long             '開始位置
    Dim i As Long

    arr = Split(path, "\")
    ' ファイルパスの末尾から存在チェックを実施する
    For i = UBound(arr) To 1 Step -1
      tmp_path = CreatePath(arr, i)
      If Not Dir(tmp_path, vbDirectory) = "" Then
        ' ディレクトリが存在した場合、処理を終了
        ' 開始位置を保持する（存在しない時の位置が必要な為、1加算する）
        cnt = i + 1
        Exit For
      End If
    Next i
    
    ' 存在しないファイルパス（開始位置）からディレクトリ作成を実施する
    For i = cnt To UBound(arr)
      tmp_path = tmp_path & "\" & arr(i)
      If Dir(tmp_path, vbDirectory) = "" Then
        MkDir tmp_path
      End If
    Next i

    CommonMakeDir = path

End Function

'-------------------------------------------------------------------------------
' メソッド      CreatePath
' 機能          パス作成
' 機能説明　　  パスを作成する
'-------------------------------------------------------------------------------
Function CreatePath(arr() As String, cnt) As String

    Dim tmp_path As String
    Dim i As Long

    tmp_path = arr(0)  ' 初期値
    For i = 1 To cnt
      tmp_path = tmp_path & "\" & arr(i)
    Next i

    CreatePath = tmp_path

End Function

'-------------------------------------------------------------------------------
' メソッド      GetTargetPrintSheets
' 機能          直接印刷・PDF対象シート取得
' 機能説明　　  直接印刷・PDF対象シートを取得する
'-------------------------------------------------------------------------------
Function GetTargetPrintSheets() As Variant
    Dim targetSheets As Variant '直接印刷・PDF対象シート
    Dim i As Long
    
    '直接印刷・PDF出力対象シートを抽出
    For i = 1 To Worksheets.Count
        If Worksheets(i).Visible = xlSheetVisible Then
            
            If (IsCopy) _
            And Worksheets(i).Name = (MainSheetNameExcel) Then
                'Excel出力用のシート名の場合、処理をスキップ
                GoTo Continue
            End If
            
            If IsEmpty(targetSheets) Then
                ReDim targetSheets(0)
            Else
                ReDim Preserve targetSheets(UBound(targetSheets) + 1)
            End If
            targetSheets(UBound(targetSheets)) = Worksheets(i).Name
        End If
Continue:
    Next i
    
    GetTargetPrintSheets = targetSheets

End Function

'-------------------------------------------------------------------------------
' メソッド      CreateTxt
' 機能          txtﾌｧｲﾙ作成
' 機能説明　　  CSV、TSVﾌｧｲﾙを作成する
'-------------------------------------------------------------------------------
Private Sub CreateTxt(filePath As String, isCsv As Boolean)
    Dim targetSheet As Worksheet    '対象シート
    Dim targetRange As Range        '対象のRange
    Dim outputAry As Variant        '要素配列
    Dim outputFile As String        'txtファイルのパス
    Dim delimiter As String         '区切り文字
    Dim targetNum As Long           'ファイル番号
    
    '出力対象シートを定義
    Set targetSheet = grb_targetWb.Worksheets(MainSheetNameExcel)
    
    '1列目から最終列の1行目から最終行までを定義
    targetSheet.Select
    Set targetRange = targetSheet.Range(ActiveSheet.PageSetup.PrintArea)
    
    '対象範囲の値を配列に格納
    outputAry = targetRange.Value
    
    '保存ファイル名を定義
    outputFile = filePath & "\" & grb_argInfo.outputFileName
    If (isCsv) Then
        'CSVの場合
        outputFile = outputFile & ".csv"
        delimiter = ","
    Else
        'TSVの場合
        outputFile = outputFile & ".tsv"
        delimiter = vbTab
    End If
    
    '空番号を取得
    targetNum = FreeFile
    
    '書き込みのためにファイルを開く（ファイルがなければ作成される）
    Open outputFile For Append As #targetNum
        
        Dim i As Long
        Dim j As Long
        Dim targetVal As String
        
        '行方向の要素数分ループ
        For i = LBound(outputAry, 1) To UBound(outputAry, 1)
            '列方向の要素数分ループ
            For j = LBound(outputAry, 2) To UBound(outputAry, 2)
                'シートの値を配列から定義
                targetVal = outputAry(i, j)
                
                '値をファイルに書き込み
                If j = UBound(outputAry, 2) Then
                    '最終列なら、「;」をつけない（「"；"」をつけると、改行なしで書き込み）
                    Print #targetNum, targetVal
                Else
                    '最終列でなければ、値の後に「","」or「tab」末尾に「";"」をつける（「"；"」をつけると、改行なしで書き込み）
                    Print #targetNum, targetVal & delimiter;
                End If
            Next
        Next
    
    Close #targetNum

End Sub

'-------------------------------------------------------------------------------
' メソッド      MainSheetNameExcel
' 機能          Excel・CSV・TSV対象のシート名取得
' 機能説明　　  Excel・CSV・TSV対象のシート名を取得する
'-------------------------------------------------------------------------------
Function MainSheetNameExcel() As String
    Dim sheetNameExcel As String 'Excel・CSV・TSV対象のシート名
    
    If (IsCopy) Then
        'PDF出力対象のシートがある場合、コピー後のシート名を取得する
        sheetNameExcel = grb_argInfo.mainSheetName & " (2)"
    Else
        'PDF出力対象のシートがない場合、メインシート名をそのまま取得する
        sheetNameExcel = grb_argInfo.mainSheetName
    End If
    
    MainSheetNameExcel = sheetNameExcel

End Function

Public Function New_PXJDO301() As PXJDO301
    Set New_PXJDO301 = New PXJDO301
End Function



