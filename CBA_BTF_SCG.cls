VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_BTF_SCG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit       ' CBA_BTF_SCG @CBA_BTF Changed 181209
 
Private Type CBA_BTF_SCGData
    scg As Long
    MonthNo As Long
    YearNo As Long
    CG As Long
    BD As String
    GBD As String
    Prods As Collection
End Type

Private CBA_BTF_CGProdData() As CBA_BTF_ProdData
Private tSCG As CBA_BTF_SCGData

Sub CreateSCG(ByVal CG As Long, ByVal scg As Long)
    Dim strSQL As String, Desc As String
    Dim curProd As Long, pcount As Long
    Dim RS As ADODB.Recordset
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    tSCG.MonthNo = Month(Date)
    tSCG.YearNo = Year(Date)
    tSCG.scg = scg
    tSCG.CG = CG
    ' Will now bring back as many records as there are from the current year + all records from the prior year regardless
    Set RS = New ADODB.Recordset
    
    ' THIS NEEDS TO BE RECREATED
    
    RS.Open strSQL, CBA_COM_CBISCN
    CBA_bFCast_NoDataReturned = False
    If RS.EOF Then
        CBA_bFCast_NoDataReturned = True
        MsgBox "No Data Returned"
        Exit Sub
    Else
        With RS
            Set tSCG.Prods = New Collection
            ReDim CBA_BTF_CGProdData(1 To 36, 0 To 0)
            curProd = 0: pcount = -1
            Do Until .EOF
                If curProd <> .Fields(0) Then
                    pcount = pcount + 1
                    curProd = .Fields(0)
                    tSCG.Prods.Add .Fields(0)
                    ReDim Preserve CBA_BTF_CGProdData(1 To 36, 0 To pcount)
                End If
'                If curProd = 43842 Then
'                    a = a
'                End If
                Set CBA_BTF_CGProdData(.Fields(2), pcount) = New CBA_BTF_ProdData
                If .Fields(0) = 99999 Then Desc = "Other" Else Desc = .Fields(13)
                CBA_BTF_CGProdData(.Fields(2), pcount).Generate .Fields(0), .Fields(1), .Fields(2), .Fields(3), .Fields(4), .Fields(5), .Fields(6), .Fields(7), .Fields(8), .Fields(9), .Fields(10), .Fields(11), .Fields(12), Desc, NZ(.Fields(14), 0)
                .MoveNext
            Loop
        End With
        RS.Close
        Set RS = Nothing
    End If
    ' Add any data that has already been forecasted. There are three scenarios:-
        ' Scenario 1 : Product is in the array, and has data for month                              ; Overwrite the old forecasted data with the new
        ' Scenario 2 : Product is in the array, but has no data for month (Array Sub-El is null)    ; Create Sub-El and write the forecasted data into it
        ' Scenario 3 : Product is not in the array, (Array El is null)                              ; Create El & Sub-El and write the forecasted data into it
    Call getForecastedProducts
        
    Set RS = New ADODB.Recordset
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
        strSQL = strSQL & "DECLARE @ENTEREDDATE date = '" & Year(Date) & "-" & Month(Date) & "-01'" & Chr(10)
        strSQL = strSQL & "DECLARE @CG int = " & CG & Chr(10)
        strSQL = strSQL & "DECLARE @SCG int = " & scg & Chr(10)
        strSQL = strSQL & "DECLARE @SDATE date = @ENTEREDDATE" & Chr(10)
        strSQL = strSQL & "DECLARE @EDATE date = dateadd(D,-1,dateadd(M,1,@SDATE))" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select bd.firstname + ' ' + bd.name as BD, gbd.firstname + ' ' + gbd.name as GBD from (" & Chr(10)
        strSQL = strSQL & "select p.empno, row_number()  over (order by count(p.empno) desc) as row from cbis599p.dbo.pos pos" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.product p on p.productcode = pos.productcode" & Chr(10)
        strSQL = strSQL & "where pos.posdate >= dateadd(D,1,dateadd(M,-1,@EDATE)) and pos.posdate <= @EDATE and p.CGno = @CG and p.SCGNO = @SCG" & Chr(10)
        strSQL = strSQL & "group by p.empno) a  left join cbis599p.dbo.employee as bd on bd.empno = a.empno" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.employee as gbd on gbd.EmpNo = bd.EmpNo_Grp"
    RS.Open strSQL, CBA_COM_CBISCN
    If RS.EOF Then
    Else
        tSCG.BD = RS.Fields(0)
        tSCG.GBD = RS.Fields(1)
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CreateSCG", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Property Get CG() As Long
    CG = tSCG.CG
End Property
Property Get scg() As Long
    scg = tSCG.scg
End Property
Property Get BD() As String
    BD = tSCG.BD
End Property
Property Get GBD() As String
    GBD = tSCG.GBD
End Property
Property Get getSCGData() 'As CBA_BTF_ProdData()
    getSCGData = CBA_BTF_CGProdData
End Property

Private Sub getForecastedProducts()
    
    ' This routine will load the data from the Forecast table ProductData and overwrite the data in the arrays
    Dim sSQL As String, lProdLastWritten As Long, lElNo As Long, dtDateSaved As Date
    Dim RS As ADODB.Recordset, CN As ADODB.Connection, lLblNo As Long, lTimes As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    Set CN = New ADODB.Connection
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ForeCast") & ";"
    Set RS = New ADODB.Recordset

    ' Query will get the last Product Data of each Product (OrderBY is in the qry
    sSQL = "SELECT * FROM qry_LatestProductForecast WHERE CG = " & tSCG.CG & " AND  SCG = " & tSCG.scg & "; "
    RS.Open sSQL, CN
    lTimes = 0
    Do While Not RS.EOF
        lTimes = lTimes + 1
        ' On change of Product Code
        If lProdLastWritten <> RS!productcode Then
            ' Find the element it is going into
            lElNo = CBA_BTN_TestProductInArray(RS!productcode, 0, CStr(RS!ForecastDate), True)
            lProdLastWritten = RS!productcode
            dtDateSaved = RS!SubmitDate
        ElseIf dtDateSaved > RS!SubmitDate Then
            GoTo SkipToNext
        End If

        lLblNo = Month(RS!ForecastDate)
        If CBA_BTF_CGProdData(lLblNo, lElNo) Is Nothing Then Call CBA_BTN_TestProductInArray(RS!productcode, lElNo, CStr(RS!ForecastDate))
        CBA_BTF_CGProdData(lLblNo, lElNo).UpdateValue "ForeAvgPrice", IIf(NZ(RS!ReForeAvgPrice, 0) > 0, NZ(RS!ReForeAvgPrice, 0), NZ(RS!ForeAvgPrice, 0))
        CBA_BTF_CGProdData(lLblNo, lElNo).UpdateValue "ForeCPU", IIf(NZ(RS!ReForeAvgUnitCost, 0) > 0, NZ(RS!ReForeAvgUnitCost, 0), NZ(RS!ForeAvgUnitCost, 0))
        CBA_BTF_CGProdData(lLblNo, lElNo).UpdateValue "ForeQTY", IIf(NZ(RS!ReForeAvgPOSQty, 0) > 0, NZ(RS!ReForeAvgPOSQty, 0), NZ(RS!ForeAvgPOSQty, 0))
        CBA_BTF_CGProdData(lLblNo, lElNo).UpdateValue "ForeRET", IIf(NZ(RS!ReForeAvgPOSRET, 0) > 0, NZ(RS!ReForeAvgPOSRET, 0), NZ(RS!ForeAvgPOSRET, 0))
SkipToNext:
        RS.MoveNext
    Loop
    
Exit_Routine:
    On Error Resume Next
    Set RS = Nothing
    Set CN = Nothing
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-getForecastedProducts", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub


Private Function CBA_BTN_TestProductInArray(ByVal lProdCode As Long, ByRef lHaveArrayEl As Long, _
                                            Optional ByVal sMthReq As String = "", _
                                            Optional ByVal bAddIfDoesntExist As Boolean = False) As Long
    ' PLEASE NOTE: THE SUB-ELEMENTS CONFORM TO THE MONTHS - IE b 1 = MONTH 1...
    
    ' Get the product b number and return it in a long. Fill in any missing b's or c's if req.
    Dim b As Long, c As Long, bProdFound As Boolean, lbFrom As Long, lbTo As Long, lNewArrayEl As Long
    Dim lMthReq As Long, lYrReq As Long, sDesc As String
    Dim sSQL As String, bPassOK As Boolean, lPCls As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    If sMthReq > "" Then lMthReq = Month(CDate(sMthReq)): lYrReq = Year(CDate(sMthReq))
    lbFrom = LBound(CBA_BTF_CGProdData, 2): lbTo = UBound(CBA_BTF_CGProdData, 2)   ' Set the default b from/to
''    lNewArrayEl = 0
    ' If the element has already been found...
    If lHaveArrayEl > 0 Then lbFrom = lHaveArrayEl: lbTo = lHaveArrayEl
    ' For each b element
    For b = lbFrom To lbTo
        For c = LBound(CBA_BTF_CGProdData, 1) To UBound(CBA_BTF_CGProdData, 1)
            If CBA_BTF_CGProdData(c, b) Is Nothing Then
            ElseIf CBA_BTF_CGProdData(c, b).productcode = lProdCode Then
                lNewArrayEl = b
                bProdFound = True
                sDesc = CBA_BTF_CGProdData(c, b).Description
                lPCls = CBA_BTF_CGProdData(c, b).ProdClass
                If lMthReq > 0 Then
                    If CBA_BTF_CGProdData(c, b).Month = lMthReq Then GoTo Exit_Routine
                Else
                    If bAddIfDoesntExist = False Then GoTo Exit_Routine
                End If
            End If
        Next c
    Next b
    ' Add the b Element, if it isn't in there - Note: Els will exist but be nothing at this stage
    If bProdFound = False Then
        ReDim Preserve CBA_BTF_CGProdData(LBound(CBA_BTF_CGProdData, 1) To UBound(CBA_BTF_CGProdData, 1), LBound(CBA_BTF_CGProdData, 2) To UBound(CBA_BTF_CGProdData, 2) + 1)
        lNewArrayEl = UBound(CBA_BTF_CGProdData, 2)
    End If
    ' Add the Month Element
    b = lNewArrayEl
    If lMthReq > 0 Then
        For c = LBound(CBA_BTF_CGProdData, 1) To UBound(CBA_BTF_CGProdData, 1)
            If CBA_BTF_CGProdData(c, b) Is Nothing Then
                If c = lMthReq Then
                    ' Get the product details....
                    If sDesc = "" Then
                        sSQL = "SELECT Description, ProductClass " & _
                               "FROM cbis599p.dbo.Product " & _
                               "WHERE ProductCode=" & lProdCode & " AND CG = " & tSCG.CG & " AND  SCG = " & tSCG.scg & ";"
                        CBA_DBtoQuery = 599
                        bPassOK = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CBIS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", sSQL, 120, , , False)
                        If bPassOK = True Then
                            sDesc = CBA_CBISarr(0, 0)
                            lPCls = CBA_CBISarr(0, 1)
                        End If
                    End If
                    Set CBA_BTF_CGProdData(lMthReq, b) = New CBA_BTF_ProdData
                    CBA_BTF_CGProdData(lMthReq, b).Generate lProdCode, lYrReq, lMthReq, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, sDesc, lPCls
                End If
            End If
        Next c
    End If
Exit_Routine:
    lHaveArrayEl = lNewArrayEl
    CBA_BTN_TestProductInArray = lNewArrayEl
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_BTN_TestProductInArray", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

