VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_BTF_MonthData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit       ' CBA_BTF_MonthData @CBA_BTF

Private plCurMonth As Long
Private plCurYear As Long
''Private CBIS As Variant
Private Fore As Variant
Private Data(1 To 70, 0 To 40) As CBA_BTF_ForecastTypes
Private LvLdic As Scripting.Dictionary
Sub formulate(ByVal lYearNo As Long, ByVal lMonthNo As Long, ByVal PClass As Long, ByRef Fore, Optional ByVal incCBIS As Boolean = True)
    Dim dates As Date, dateE As Date, strSQL As String, a As Long
    Dim DefaultLevelWhenNoForecast As String, thisForecastLevel As String '', Val As String
    Dim this
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    plCurMonth = lMonthNo: plCurYear = lYearNo
    dates = DateSerial(plCurYear, plCurMonth, 1)
    dateE = DateSerial(plCurYear, plCurMonth + 1, 0)
    
    
    Set LvLdic = New Scripting.Dictionary
    DefaultLevelWhenNoForecast = "CG"
    If Fore(0, 0) <> 0 Then
        For a = LBound(Fore, 2) To UBound(Fore, 2)
            If Fore(3, a) = 0 Then
                If LvLdic.Exists(Fore(2, a)) = False Then LvLdic.Add Fore(2, a), "CG"
                'DefaultLevelWhenNoForecast = "CG"
            ElseIf Fore(3, a) > 0 Then
                If LvLdic.Exists(Fore(2, a)) = False Then LvLdic.Add Fore(2, a), "SCG"
                'DefaultLevelWhenNoForecast = "SCG"
            Else
                If LvLdic.Exists(Fore(2, a)) = False Then LvLdic.Add Fore(2, a), "None"
                'this = "None"
            End If
        Next
    End If
    
    If Fore(0, 0) <> 0 Then
        For a = LBound(Fore, 2) To UBound(Fore, 2)

            If NZ(Fore(3, a), "") = "" Or Fore(3, a) = "-" Then Fore(3, a) = 0
            If Data(Fore(2, a), Fore(3, a)) Is Nothing Then Set Data(Fore(2, a), Fore(3, a)) = New CBA_BTF_ForecastTypes
            If Fore(0, a) = lYearNo And Fore(1, a) = lMonthNo And Fore(4, a) = PClass Then
                If Fore(3, a) = 0 Then
                    this = "CG"
                ElseIf Fore(3, a) > 0 Then
                    this = "SCG"
                Else
                    this = "None"
                End If
'                If Fore(2, a) = 5 And Fore(3, a) = 1 Then
'                a = a
'                End If
                ' Fore format is.... YearNo(0), MonthNo(1), CG(2), SCG(3), PC(4), FRetailO(5), FRetailN(6), FMarginO(7), FMarginN(8)
                If Data(Fore(2, a), 0) Is Nothing Then Set Data(Fore(2, a), 0) = New CBA_BTF_ForecastTypes
                Data(Fore(2, a), 0).setLevel this: thisForecastLevel = this     ' @RWFC Should this be (Fore(2, a), Fore(3, a)) as are the others
                Data(Fore(2, a), Fore(3, a)).Sales NZ(Fore(5, a), 0), NZ(Fore(6, a), 0)
                Data(Fore(2, a), Fore(3, a)).MarginDol NZ(Fore(5, a), 0) * NZ(Fore(7, a), 0), NZ(Fore(6, a), 0) * NZ(Fore(8, a), 0)
                Data(Fore(2, a), Fore(3, a)).MarginPer NZ(Fore(7, a), 0), NZ(Fore(8, a), 0)
                Data(Fore(2, a), Fore(3, a)).Cost NZ(Fore(5, a), 0) * (1 - NZ(Fore(7, a), 0)), NZ(Fore(6, a), 0) * (1 - NZ(Fore(8, a), 0))
                'added code
                If this = "SCG" Then
                    Data(Fore(2, a), 0).SalesAcc NZ(Fore(5, a), 0), NZ(Fore(6, a), 0)
                    Data(Fore(2, a), 0).MarginDolAcc NZ(Fore(5, a), 0) * NZ(Fore(7, a), 0), NZ(Fore(6, a), 0) * NZ(Fore(8, a), 0)
'                    data(Fore(2, a), 0).MarginPer NZ(Fore(7, a), 0), NZ(Fore(8, a), 0)
                    Data(Fore(2, a), 0).CostAcc NZ(Fore(5, a), 0) * (1 - NZ(Fore(7, a), 0)), NZ(Fore(6, a), 0) * (1 - NZ(Fore(8, a), 0))
                End If
                'end of added code
            End If
        Next
    End If
    
    'If thisForecastLevel = "" Then thisForecastLevel = DefaultLevelWhenNoForecast
    If incCBIS = True Then
        getCBISData lYearNo, lMonthNo, PClass ', thisForecastLevel
        If CBA_COM_SKU_CBISRS.State = 1 Then CBA_COM_SKU_CBISRS.Close
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-Formulate", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Function getCurMonth() As Date
    getCurMonth = plCurMonth
End Function
Function getCurYear() As Date
    getCurYear = plCurYear
End Function
Property Get CGData(ByVal CG As Long, ByVal scg As Long) As CBA_BTF_ForecastTypes
    If Data(CG, scg) Is Nothing Then Set Data(CG, scg) = New CBA_BTF_ForecastTypes
    Set CGData = Data(CG, scg)
End Property
Private Function getCBISData(ByVal lYearNo As Long, ByVal lMonthNo As Long, ByVal PClass As Long)
    Dim strSQL As String
    Dim a As Long
    Dim CG
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    For Each CG In LvLdic
        Set CBA_COM_SKU_CBISRS = New ADODB.Recordset
        If LvLdic(CG) = "SCG" Then strSQL = "select cgno, SCGno, productclass, retailnet, cost, retail, posretail from #OP" & Chr(10)
        If LvLdic(CG) = "CG" Then strSQL = "select cgno, 0 as SCG, productclass, sum(retailnet) as retailnet, sum(cost) as cost, sum(retail) as retail, sum(posretail) as posretail  from #OP" & Chr(10)
        strSQL = strSQL & "where yearno = " & lYearNo & Chr(10)
        strSQL = strSQL & "and monthno = " & lMonthNo & Chr(10)
        strSQL = strSQL & "and productclass = " & PClass & Chr(10)
        If LvLdic(CG) = "CG" Then strSQL = strSQL & "Group by cgno, productclass"
        'Debug.Print strSQL
        CBA_COM_SKU_CBISRS.Open strSQL, CBA_COM_SKU_CBISCN
        Do Until CBA_COM_SKU_CBISRS.EOF
            If Data(CBA_COM_SKU_CBISRS.Fields(0), CBA_COM_SKU_CBISRS.Fields(1)) Is Nothing Then
                Set Data(CBA_COM_SKU_CBISRS.Fields(0), CBA_COM_SKU_CBISRS.Fields(1)) = New CBA_BTF_ForecastTypes
            End If
            Data(CBA_COM_SKU_CBISRS.Fields(0), CBA_COM_SKU_CBISRS.Fields(1)).Sales , , CBA_COM_SKU_CBISRS.Fields(6)
            If CBA_COM_SKU_CBISRS.Fields(5) > 0 Then
                Data(CBA_COM_SKU_CBISRS.Fields(0), CBA_COM_SKU_CBISRS.Fields(1)).MarginPer , , (CBA_COM_SKU_CBISRS.Fields(3) - CBA_COM_SKU_CBISRS.Fields(4)) / CBA_COM_SKU_CBISRS.Fields(5)
                Data(CBA_COM_SKU_CBISRS.Fields(0), CBA_COM_SKU_CBISRS.Fields(1)).MarginDol , , ((CBA_COM_SKU_CBISRS.Fields(3) - CBA_COM_SKU_CBISRS.Fields(4)) / CBA_COM_SKU_CBISRS.Fields(5)) * CBA_COM_SKU_CBISRS.Fields(6)
            End If
            Set Data(CBA_COM_SKU_CBISRS.Fields(0), CBA_COM_SKU_CBISRS.Fields(1)).RCVRetailNet = CBA_COM_SKU_CBISRS.Fields(3)
            Set Data(CBA_COM_SKU_CBISRS.Fields(0), CBA_COM_SKU_CBISRS.Fields(1)).RCVRetail = CBA_COM_SKU_CBISRS.Fields(5)
            Data(CBA_COM_SKU_CBISRS.Fields(0), CBA_COM_SKU_CBISRS.Fields(1)).Cost , , CBA_COM_SKU_CBISRS.Fields(4)
            CBA_COM_SKU_CBISRS.MoveNext
        Loop
    Next
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-getCBISData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

