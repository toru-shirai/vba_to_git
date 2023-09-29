Option Explicit

'-------------------------------------------------------------------------------
' E.SPECの表現方法(「USL,LSL,上限値,下限値」分離枠の場合)
'-------------------------------------------------------------------------------
'Nominal Dim, Nominal,target
Function SpecBunriwakuNominal(ByVal cellValue, ByVal kikakuType As String)

    Select Case kikakuType
    Case "0", "2", "9"
        SpecBunriwakuNominal = cellValue
    Case "6", "8"
        SpecBunriwakuNominal = "-"
    Case "1", "3"
        SpecBunriwakuNominal = ""
    End Select
End Function

'Tol Max (+), Up Tol
Function SpecBunriwakuTolMax(ByVal cellValue, ByVal kikakuType As String)
    'SpecBunriwakuNominalと同じ
    SpecBunriwakuTolMax = SpecBunriwakuNominal(cellValue, kikakuType)
End Function

'Tol Min (-), Low Tol
Function SpecBunriwakuTolMin(ByVal cellValue, ByVal kikakuType As String)
    'SpecBunriwakuNominalと同じ
    SpecBunriwakuTolMin = SpecBunriwakuNominal(cellValue, kikakuType)
End Function

'USL, Target Max, Upper
Function SpecBunriwakuUsl(ByVal cellValue, ByVal kikakuType As String)

    Select Case kikakuType
    Case "0", "1", "2", "9"
        SpecBunriwakuUsl = cellValue
    Case "6", "8"
        SpecBunriwakuUsl = "-"
    Case "3"
        SpecBunriwakuUsl = ""
    End Select
End Function

'LSL, Target Min, Lower
Function SpecBunriwakuLsl(ByVal cellValue, ByVal kikakuType As String)

    Select Case kikakuType
    Case "0", "2", "3", "9"
        SpecBunriwakuLsl = cellValue
    Case "6", "8"
        SpecBunriwakuLsl = "-"
    Case "1"
        SpecBunriwakuLsl = ""
    End Select
End Function

