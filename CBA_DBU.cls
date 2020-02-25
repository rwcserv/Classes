VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_DBU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit                     'CBA_DBU 190731
'''Each Class module contains all the relevent DBU data for one month in a year
Private Forecast As Scripting.Dictionary
Private Sales As Scripting.Dictionary
Private CGSales As Scripting.Dictionary
Private SCGSales As Scripting.Dictionary
Private pcSales As Scripting.Dictionary
Private CGPCSales As Scripting.Dictionary
Private RCV As Scripting.Dictionary
Private Wastage As DBU_Pclass
Private RCVMargin As DBU_Pclass
Private POSMargin As DBU_Pclass
Private RCVCont As DBU_Pclass
Private POSCont As DBU_Pclass
Private PaL As DBU_PaLCosts
Private MonthNo As Byte
Private YearNo As Integer
Private DBU As String
Private ForeCastData(1 To 4) As CBA_BTF_MonthData
Private sumSalesOriFore As Single, sumSalesReFore As Single, sumMarginOriFore As Single
Private sumMarginReFore As Single, sumContOriFore As Single, sumContReFore As Single
Private TotBus As Single
Private TotBusReforecast As Single
Private TotBusOriginal As Single
Private TotbusRCV As Single
Property Get TotalBusinessRetail()
    TotalBusinessRetail = TotBus
End Property
Property Get TotalBusinessRetailRCV()
    TotalBusinessRetailRCV = TotbusRCV
End Property
Property Get TotalBusinessOFore()
    TotalBusinessOFore = TotBusOriginal
End Property
Property Get TotalBusinessRetailReFore()
    TotalBusinessRetailReFore = TotBusReforecast
End Property
Property Get RetailOriginalForecast()
    RetailOriginalForecast = sumSalesOriFore
End Property
Property Get RetailReForecast()
    RetailReForecast = sumSalesReFore
End Property
Property Get MarginOriginalForecast()
    MarginOriginalForecast = sumMarginOriFore
End Property
Property Get MarginReForecast()
    MarginReForecast = sumMarginReFore
End Property
Property Get ContributionOriginalForecast()
    ContributionOriginalForecast = sumContOriFore
End Property
Property Get ContributionReForecast()
    ContributionReForecast = sumContReFore
End Property
Property Get RCVData(ByVal ValuetoReturn As String, Optional ByVal CGno = -1, Optional ByVal SCGno = -1, Optional ProductClass = -1)    'As Single
    On Error GoTo hand
    If ValuetoReturn = "" Then RCVData = 0: On Error GoTo 0: Exit Property
    If CGno = -1 Then RCVData = RCV(ValuetoReturn): On Error GoTo 0: Exit Property
    If SCGno = -1 Then RCVData = RCV(CGno)(ValuetoReturn): On Error GoTo 0: Exit Property
    If ProductClass = -1 Then RCVData = RCV(CGno)(SCGno)(ValuetoReturn): On Error GoTo 0: Exit Property
    RCVData = RCV(CGno)(SCGno)(ProductClass)(ValuetoReturn)
    On Error GoTo 0: Exit Property
hand:
    RCVData = 0: On Error GoTo 0: Err.Clear: Exit Property
End Property
Property Get DBUForecast(ByVal ProductClass As Byte) As CBA_BTF_MonthData
    Set DBUForecast = ForeCastData(ProductClass)
End Property
Property Get SalesData(ByVal ValuetoReturn As String, Optional ByVal CGno = -1, Optional ByVal SCGno = -1, Optional ProductClass = -1)    'As Single
    On Error GoTo hand
    If ValuetoReturn = "" Then SalesData = 0: On Error GoTo 0: Exit Property
    If CGno = -1 Then SalesData = Sales(ValuetoReturn): On Error GoTo 0: Exit Property
    If SCGno = -1 Then SalesData = Sales(CGno)(ValuetoReturn): On Error GoTo 0: Exit Property
    If ProductClass = -1 Then SalesData = Sales(CGno)(SCGno)(ValuetoReturn): On Error GoTo 0: Exit Property
    SalesData = Sales(CGno)(SCGno)(ProductClass)(ValuetoReturn)
    On Error GoTo 0: Exit Property
hand:
    SalesData = 0: On Error GoTo 0: Err.Clear: Exit Property
End Property
Property Get MonthNum() As Byte
    MonthNum = MonthNo
End Property
Property Get YearNum() As Integer
    YearNum = YearNo
End Property
Sub formulate(ByVal MnthNo As Byte, ByVal YrNo As Integer, ByVal DBURef As String, ByRef ActiveCGList As Variant, ByVal Totalbus As Single, ByVal TotalForebus As Single, ByVal TotalReForebus As Single, ByVal TotalRCVBus As Single)
Dim TempArr
Dim CGallocate As Scripting.Dictionary, CGPCCheck As Scripting.Dictionary
Dim sumCGPCQTY(1 To 4) As Single, sumCGPCRetail(1 To 4) As Single, sumCGPCCost(1 To 4) As Single, sumCGPCRetNet(1 To 4) As Single
Dim DBUCGList As Variant, arrCG As Variant
Dim strSQL As String, PCstr  As String, strCheck As String
Dim CBA_COM_CBISRS As ADODB.Recordset
Dim curActiveCG As Byte
Dim atCG As Boolean, added As Boolean
Dim a As Long, CGcnt As Long, lCurCG As Long, lCurSCG As Long, curPC As Long, Pc As Long, c As Long
Dim TotQTY As Single, TotRetail As Single, sumSCGRetail As Single, sumSCGQTY As Single, sumCGRetail As Single, sumCGQTY As Single
Dim sumCGCost As Single, sumCGRetNet As Single, TotCost As Single, TotRetNet As Single, sumSCGCost As Single, sumSCGRetNet As Single

    MonthNo = MnthNo
    YearNo = YrNo
    DBUCGList = CBA_DBU_Runtime.GetDBUCGList(DBURef)
    On Error Resume Next
    If DBUCGList(1, 1) = 0 Then
    Err.Clear
    On Error GoTo 0
        MsgBox "No CG Data Returned"
        Exit Sub
    End If
    On Error GoTo 0
    
    
    TotBusReforecast = TotalReForebus
    TotBusOriginal = TotalForebus
    TotBus = Totalbus
    TotbusRCV = TotalRCVBus
           
    Set CGallocate = New Scripting.Dictionary
    curActiveCG = 0: atCG = True
    For a = LBound(DBUCGList, 2) To UBound(DBUCGList, 2)
        If CByte(DBUCGList(1, a)) <> curActiveCG Then
            curActiveCG = CByte(DBUCGList(1, a))
            CGallocate.Add curActiveCG, CBA_DBU_Runtime.checkMultiDBUCG(curActiveCG)
        End If
    Next a
    
    
    
   
    
    
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
    strSQL = strSQL & "select p.cgno, p.scgno, p.productclass, isnull(sum(pos.retail),0) as Retail, sum(Quantity) As QTY" & Chr(10)
    strSQL = strSQL & "from cbis599p.dbo.pos pos" & Chr(10)
    strSQL = strSQL & "inner join cbis599p.dbo.product p on p.productcode = pos.productcode" & Chr(10)
    strSQL = strSQL & "where posdate >= '" & YearNo & "-" & MonthNo & "-01' and posdate <= '" & Format(DateSerial(YearNo, MonthNo + 1, 0), "YYYY-MM-DD") & "' and (" & Chr(10)
    curActiveCG = DBUCGList(1, 1)
    added = False
    For a = LBound(DBUCGList, 2) To UBound(DBUCGList, 2)
        If curActiveCG <> DBUCGList(1, a) Then
            curActiveCG = DBUCGList(1, a)
            added = False
        End If
        If CGallocate(DBUCGList(1, a)) = True And added = False Then
            If curActiveCG <> DBUCGList(1, UBound(DBUCGList, 2)) Then
                strSQL = strSQL & "(Cgno = " & DBUCGList(1, a) & ") or " & Chr(10)
            Else
                strSQL = strSQL & "(Cgno = " & DBUCGList(1, a) & "))" & Chr(10)
            End If
            added = True
        End If
        If CGallocate(DBUCGList(1, a)) = False Then
            If a < UBound(DBUCGList, 2) Then
                strSQL = strSQL & "(Cgno = " & DBUCGList(1, a) & " and scgno = " & DBUCGList(2, a) & " and Productclass = " & DBUCGList(3, a) & ") or " & Chr(10)
            Else
                strSQL = strSQL & "(Cgno = " & DBUCGList(1, a) & " and scgno = " & DBUCGList(2, a) & " and Productclass = " & DBUCGList(3, a) & "))" & Chr(10)
            End If
        End If
    Next
    strSQL = strSQL & "group by p.cgno, p.scgno, p.productclass" & Chr(10)
    strSQL = strSQL & "order by p.cgno, p.scgno, p.productclass" & Chr(10)
    
    Set CBA_COM_CBISRS = New ADODB.Recordset
    CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN
    If CBA_COM_CBISRS.EOF = False Then
        arrCG = CBA_COM_CBISRS.GetRows()
        Set Sales = New Scripting.Dictionary
        CGcnt = 0: lCurCG = 0: lCurSCG = -1: curPC = 0
        For a = LBound(arrCG, 2) To UBound(arrCG, 2)
'        If arrCG(0, a) = 14 Then
'        a = a
'        End If
            TotQTY = TotQTY + arrCG(4, a)
            TotRetail = TotRetail + arrCG(3, a)
            If arrCG(0, a) <> lCurCG Then
                If CGcnt > 0 Then
                    SCGSales.Add "Retail", sumSCGRetail
                    SCGSales.Add "Quantity", sumSCGQTY
                    sumSCGRetail = 0: sumSCGQTY = 0
                    CGSales.Add CByte(lCurSCG), SCGSales
                    Set SCGSales = New Scripting.Dictionary
                    CGSales.Add "Retail", sumCGRetail
                    CGSales.Add "Quantity", sumCGQTY
                    sumCGQTY = 0: sumCGRetail = 0
                    Sales.Add CByte(lCurCG), CGSales
                    Set CGSales = New Scripting.Dictionary
                Else
                    Set CGSales = New Scripting.Dictionary
                    Set SCGSales = New Scripting.Dictionary
                    Set pcSales = New Scripting.Dictionary
                End If
                CGcnt = CGcnt + 1
                lCurCG = arrCG(0, a)
            End If
            sumCGQTY = sumCGQTY + arrCG(4, a)
            sumCGRetail = sumCGRetail + arrCG(3, a)
            If arrCG(1, a) <> lCurSCG Or a = UBound(arrCG, 2) Then
                If CGcnt > 0 And lCurSCG <> -1 Then
                    If a > LBound(arrCG, 2) Then
                    If arrCG(0, a - 1) = lCurCG Then
                        SCGSales.Add "Retail", sumSCGRetail
                        SCGSales.Add "Quantity", sumSCGQTY
                        sumSCGRetail = 0: sumSCGQTY = 0
                        CGSales.Add CByte(lCurSCG), SCGSales
                        Set SCGSales = New Scripting.Dictionary
                    End If
                    End If
                End If
                lCurSCG = NZ(arrCG(1, a), 0) '#TP added NZ that may be a bad idea
            End If
            sumSCGQTY = sumSCGQTY + arrCG(4, a)
            sumSCGRetail = sumSCGRetail + arrCG(3, a)
            Set pcSales = New Scripting.Dictionary
            pcSales.Add "Retail", arrCG(3, a)
            pcSales.Add "Quantity", arrCG(4, a)
            SCGSales.Add CByte(arrCG(2, a)), pcSales
            If a = UBound(arrCG, 2) Then
                CGSales.Add "Retail", sumCGRetail
                CGSales.Add "Quantity", sumCGQTY
                sumCGQTY = 0: sumCGRetail = 0
                Sales.Add CByte(arrCG(0, a)), CGSales
                Set CGSales = New Scripting.Dictionary
            End If
        Next
        Sales.Add "Quantity", TotQTY
        Sales.Add "Retail", TotRetail
    Else
        Set Sales = New Scripting.Dictionary
        Sales.Add "Quantity", 0
        Sales.Add "Retail", 0
    End If
    

    
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
    strSQL = strSQL & "select p.cgno, p.scgno, p.productclass, sum(rcv.retail) as Retail, sum(rcv.Quantity) as Quantity, sum(rcv.retailnet) as RetailNet, sum(rcv.Cost)As Cost" & Chr(10)
    strSQL = strSQL & "from cbis599p.dbo.RECEIVING rcv" & Chr(10)
    strSQL = strSQL & "inner join cbis599p.dbo.product p on p.productcode = rcv.productcode" & Chr(10)
    strSQL = strSQL & "where dayenddate >= '" & YearNo & "-" & MonthNo & "-01' and dayenddate <= '" & Format(DateSerial(YearNo, MonthNo + 1, 0), "YYYY-MM-DD") & "' and (" & Chr(10)
    curActiveCG = DBUCGList(1, 1)
    added = False
    For a = LBound(DBUCGList, 2) To UBound(DBUCGList, 2)
        If curActiveCG <> DBUCGList(1, a) Then
            curActiveCG = DBUCGList(1, a)
            added = False
        End If
        If CGallocate(DBUCGList(1, a)) = True And added = False Then
            If curActiveCG <> DBUCGList(1, UBound(DBUCGList, 2)) Then
                strSQL = strSQL & "(Cgno = " & DBUCGList(1, a) & ") or " & Chr(10)
            Else
                strSQL = strSQL & "(Cgno = " & DBUCGList(1, a) & "))" & Chr(10)
            End If
            added = True
        End If
        If CGallocate(DBUCGList(1, a)) = False Then
            If a < UBound(DBUCGList, 2) Then
                strSQL = strSQL & "(Cgno = " & DBUCGList(1, a) & " and scgno = " & DBUCGList(2, a) & " and Productclass = " & DBUCGList(3, a) & ") or " & Chr(10)
            Else
                strSQL = strSQL & "(Cgno = " & DBUCGList(1, a) & " and scgno = " & DBUCGList(2, a) & " and Productclass = " & DBUCGList(3, a) & "))" & Chr(10)
            End If
        End If
    Next
    strSQL = strSQL & "group by p.cgno, p.scgno, p.productclass" & Chr(10)
    strSQL = strSQL & "order by p.cgno, p.scgno, p.productclass" & Chr(10)
    
    If IsEmpty(arrCG) Then Else Erase arrCG
    Set CBA_COM_CBISRS = New ADODB.Recordset
    CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN
    
    
    If CBA_COM_CBISRS.EOF = False Then
        arrCG = CBA_COM_CBISRS.GetRows()
        Set RCV = New Scripting.Dictionary
        CGcnt = 0: lCurCG = 0: lCurSCG = -1: curPC = 0
        sumCGQTY = 0: sumCGRetail = 0: sumCGCost = 0: sumCGRetNet = 0
        TotQTY = 0: TotRetail = 0: TotCost = 0: TotRetNet = 0
        sumCGQTY = 0: sumCGRetail = 0: sumCGCost = 0: sumCGRetNet = 0
        sumSCGRetail = 0: sumSCGQTY = 0: sumSCGCost = 0: sumSCGRetNet = 0
        For a = LBound(arrCG, 2) To UBound(arrCG, 2)
            TotQTY = TotQTY + arrCG(4, a)
            TotRetail = TotRetail + arrCG(3, a)
            TotCost = TotCost + arrCG(6, a)
            TotRetNet = TotRetNet + arrCG(5, a)
'            If arrCG(0, a) = 24 Then
'            a = a
'            End If
            
            
            If arrCG(0, a) <> lCurCG Then
                'Debug.Print lCurCG
                If CGcnt > 0 Then
                    SCGSales.Add "Retail", sumSCGRetail
                    SCGSales.Add "Quantity", sumSCGQTY
                    SCGSales.Add "Cost", sumSCGCost
                    SCGSales.Add "RetailNet", sumSCGRetNet
                    If sumSCGRetail = 0 Then
                        SCGSales.Add "Margin", 0: SCGSales.Add "Contribution", 0
                    Else
                        SCGSales.Add "Margin", ((sumSCGRetNet - sumSCGCost) / sumSCGRetail)
                        SCGSales.Add "Contribution", (sumSCGRetNet - sumSCGCost)
                    End If
                    sumSCGRetail = 0: sumSCGQTY = 0: sumSCGCost = 0: sumSCGRetNet = 0
                    CGSales.Add CByte(lCurSCG), SCGSales
                    Set SCGSales = New Scripting.Dictionary
                    CGSales.Add "Retail", sumCGRetail
                    CGSales.Add "Quantity", sumCGQTY
                    CGSales.Add "Cost", sumCGCost
                    CGSales.Add "RetailNet", sumCGRetNet
                    If sumCGRetail = 0 Then
                        CGSales.Add "Margin", 0: CGSales.Add "Contribution", 0
                    Else
                        CGSales.Add "Margin", ((sumCGRetNet - sumCGCost) / sumCGRetail)
                        CGSales.Add "Contribution", (sumCGRetNet - sumCGCost)
                    End If
                    For Pc = 1 To 4
                        Set CGPCSales = New Scripting.Dictionary
                        CGPCSales.Add "Quantity", sumCGPCQTY(Pc)
                        CGPCSales.Add "Retail", sumCGPCRetail(Pc)
                        CGPCSales.Add "Cost", sumCGPCCost(Pc)
                        CGPCSales.Add "RetailNet", sumCGPCRetNet(Pc)
                        PCstr = "PC" & Format(Pc, "00")
                        CGSales.Add PCstr, CGPCSales
                        sumCGPCQTY(Pc) = 0: sumCGPCRetail(Pc) = 0: sumCGPCCost(Pc) = 0: sumCGPCRetNet(Pc) = 0
                    Next
                    sumCGQTY = 0: sumCGRetail = 0: sumCGCost = 0: sumCGRetNet = 0
                    RCV.Add CByte(lCurCG), CGSales
                    Set CGSales = New Scripting.Dictionary
                Else
                    Set CGSales = New Scripting.Dictionary
                    Set SCGSales = New Scripting.Dictionary
                    Set pcSales = New Scripting.Dictionary
                End If
                CGcnt = CGcnt + 1
                lCurCG = arrCG(0, a)
                lCurSCG = NZ(arrCG(1, a), 0)
    '            If lCurCG = 4 Then
    '            a = a
    '            End If
            End If
            sumCGQTY = sumCGQTY + arrCG(4, a)
            sumCGRetail = sumCGRetail + arrCG(3, a)
            sumCGCost = sumCGCost + arrCG(6, a)
            sumCGRetNet = sumCGRetNet + arrCG(5, a)
            If NZ(arrCG(1, a), 0) <> lCurSCG Or a = UBound(arrCG, 2) Then
                If CGcnt > 0 And lCurSCG <> -1 Then 'And (sumSCGQTY <> 0 Or sumSCGRetail <> 0 Or sumSCGCost <> 0 Or sumSCGRetNet <> 0) Then
                    SCGSales.Add "Retail", sumSCGRetail
                    SCGSales.Add "Quantity", sumSCGQTY
                    SCGSales.Add "Cost", sumSCGCost
                    SCGSales.Add "RetailNet", sumSCGRetNet
                    If sumSCGRetail = 0 Then
                        SCGSales.Add "Margin", 0: SCGSales.Add "Contribution", 0
                    Else
                        SCGSales.Add "Margin", ((sumSCGRetNet - sumSCGCost) / sumSCGRetail)
                        SCGSales.Add "Contribution", (sumSCGRetNet - sumSCGCost)
                    End If
                    sumSCGRetail = 0: sumSCGQTY = 0: sumSCGCost = 0: sumSCGRetNet = 0
                    CGSales.Add CByte(lCurSCG), SCGSales
                    Set SCGSales = New Scripting.Dictionary
                End If
                lCurSCG = NZ(arrCG(1, a), 0)
    '            If lCurSCG = 3 Then
    '            a = a
    '            End If
            End If
            sumSCGQTY = sumSCGQTY + arrCG(4, a)
            sumSCGRetail = sumSCGRetail + arrCG(3, a)
            sumSCGCost = sumSCGCost + arrCG(6, a)
            sumSCGRetNet = sumSCGRetNet + arrCG(5, a)
            Set pcSales = New Scripting.Dictionary
            pcSales.Add "Retail", arrCG(3, a)
            pcSales.Add "Quantity", arrCG(4, a)
            pcSales.Add "Cost", arrCG(6, a)
            pcSales.Add "RetailNet", arrCG(5, a)
            If arrCG(3, a) = 0 Then
                pcSales.Add "Margin", 0: pcSales.Add "Contribution", 0
            Else
                pcSales.Add "Margin", ((arrCG(5, a) - arrCG(6, a)) / arrCG(3, a))
                pcSales.Add "Contribution", (arrCG(5, a) - arrCG(6, a))
            End If
'            If arrCG(2, a) = 3 And lCurCG = 1 And lCurSCG = 1 Then
'            a = a
'            End If
            
            SCGSales.Add CByte(arrCG(2, a)), pcSales
            sumCGPCQTY(arrCG(2, a)) = sumCGPCQTY(arrCG(2, a)) + arrCG(4, a)
            sumCGPCRetail(arrCG(2, a)) = sumCGPCRetail(arrCG(2, a)) + arrCG(3, a)
            sumCGPCCost(arrCG(2, a)) = sumCGPCCost(arrCG(2, a)) + arrCG(6, a)
            sumCGPCRetNet(arrCG(2, a)) = sumCGPCRetNet(arrCG(2, a)) + arrCG(5, a)

            If a = UBound(arrCG, 2) Then
                CGSales.Add "Retail", sumCGRetail
                CGSales.Add "Quantity", sumCGQTY
                CGSales.Add "Cost", sumCGCost
                CGSales.Add "RetailNet", sumCGRetNet
                If sumCGRetail = 0 Then
                    CGSales.Add "Margin", 0: CGSales.Add "Contribution", 0
                Else
                    CGSales.Add "Margin", ((sumCGRetNet - sumCGCost) / sumCGRetail)
                    CGSales.Add "Contribution", (sumCGRetNet - sumCGCost)
                End If
                sumCGQTY = 0: sumCGRetail = 0: sumCGCost = 0: sumCGRetNet = 0
                For Pc = 1 To 4
                    Set CGPCSales = New Scripting.Dictionary
                    CGPCSales.Add "Quantity", sumCGPCQTY(Pc)
                    CGPCSales.Add "Retail", sumCGPCRetail(Pc)
                    CGPCSales.Add "Cost", sumCGPCCost(Pc)
                    CGPCSales.Add "RetailNet", sumCGPCRetNet(Pc)
                    PCstr = "PC" & Format(Pc, "00")
                    CGSales.Add PCstr, CGPCSales
                    sumCGPCQTY(Pc) = 0: sumCGPCRetail(Pc) = 0: sumCGPCCost(Pc) = 0: sumCGPCRetNet(Pc) = 0
                Next
                RCV.Add CByte(arrCG(0, a)), CGSales
                Set CGSales = New Scripting.Dictionary
            End If
        Next
        RCV.Add "Quantity", TotQTY
        RCV.Add "Retail", TotRetail
        RCV.Add "Cost", TotCost
        RCV.Add "RetailNet", TotRetNet
        If TotRetail = 0 Then
            RCV.Add "Margin", 0: RCV.Add "Contribution", 0
        Else
            RCV.Add "Margin", ((TotRetNet - TotCost) / TotRetail)
            RCV.Add "Contribution", (TotRetNet - TotCost)
        End If
    Else
        Set RCV = New Scripting.Dictionary
        RCV.Add "Quantity", 0
        RCV.Add "Retail", 0
        RCV.Add "Cost", 0
        RCV.Add "RetailNet", 0
        RCV.Add "Margin", 0
        RCV.Add "Contribution", 0
    End If
    
'    For Each k In RCV.Keys
'        Debug.Print k
'    Next
    
    
    Set Forecast = New Scripting.Dictionary
    Set CGSales = New Scripting.Dictionary
    Set SCGSales = New Scripting.Dictionary
    Set pcSales = New Scripting.Dictionary
    CGcnt = 0: lCurCG = 0: lCurSCG = -1: curPC = 0
    sumCGQTY = 0: sumCGRetail = 0: sumCGCost = 0: sumCGRetNet = 0
    TotQTY = 0: TotRetail = 0: TotCost = 0: TotRetNet = 0
    sumCGQTY = 0: sumCGRetail = 0: sumCGCost = 0: sumCGRetNet = 0
    sumSCGRetail = 0: sumSCGQTY = 0: sumSCGCost = 0: sumSCGRetNet = 0
    
    Set CGPCCheck = New Scripting.Dictionary
    sumSalesOriFore = 0: sumSalesReFore = 0: sumMarginOriFore = 0: sumMarginReFore = 0: sumContOriFore = 0: sumContReFore = 0
    For c = LBound(DBUCGList, 2) To UBound(DBUCGList, 2)
        If FCbM(YearNo, MonthNo, DBUCGList(3, c)).CGData(DBUCGList(1, c), 0).Level = "CG" Then
            strCheck = DBUCGList(1, c) & "-" & Format(DBUCGList(3, c), "00")
            If CGPCCheck.Exists(strCheck) = False Then
                If CGallocate(DBUCGList(1, c)) = 0 Then
                    Debug.Print "Forecast for CG: " & DBUCGList(1, c) & "-PC: " & DBUCGList(3, c) & " at CG Level; but DBU SQL asking for at SCG"
                End If
                sumSalesOriFore = sumSalesOriFore + FCbM(YearNo, MonthNo, DBUCGList(3, c)).CGData(DBUCGList(1, c), 0).Sales.OriginalForecast
                sumSalesReFore = sumSalesReFore + FCbM(YearNo, MonthNo, DBUCGList(3, c)).CGData(DBUCGList(1, c), 0).Sales.ReForecast
                If RCV("Retail") <> 0 And RCV.Exists(DBUCGList(1, c)) Then
                    sumMarginOriFore = sumMarginOriFore + (FCbM(YearNo, MonthNo, DBUCGList(3, c)).CGData(DBUCGList(1, c), 0).MarginPer.OriginalForecast * RCV(DBUCGList(1, c))("PC" & Format(DBUCGList(3, c), "00"))("Retail"))
                    sumMarginReFore = sumMarginReFore + (FCbM(YearNo, MonthNo, DBUCGList(3, c)).CGData(DBUCGList(1, c), 0).MarginPer.ReForecast * RCV(DBUCGList(1, c))("PC" & Format(DBUCGList(3, c), "00"))("Retail"))
                End If
                sumContOriFore = sumContOriFore + FCbM(YearNo, MonthNo, DBUCGList(3, c)).CGData(DBUCGList(1, c), 0).MarginDol.OriginalForecast
                sumContReFore = sumContReFore + FCbM(YearNo, MonthNo, DBUCGList(3, c)).CGData(DBUCGList(1, c), 0).MarginDol.ReForecast
                CGPCCheck.Add strCheck, 1
            End If
        Else
            sumSalesOriFore = sumSalesOriFore + FCbM(YearNo, MonthNo, DBUCGList(3, c)).CGData(DBUCGList(1, c), DBUCGList(2, c)).Sales.OriginalForecast
            sumSalesReFore = sumSalesReFore + FCbM(YearNo, MonthNo, DBUCGList(3, c)).CGData(DBUCGList(1, c), DBUCGList(2, c)).Sales.ReForecast
            If RCV.Exists(DBUCGList(1, c)) = True Then
                If RCV(DBUCGList(1, c)).Exists(DBUCGList(2, c)) = True Then
                    If RCV(DBUCGList(1, c))(DBUCGList(2, c)).Exists(DBUCGList(3, c)) = True Then
                        sumMarginOriFore = sumMarginOriFore + (FCbM(YearNo, MonthNo, DBUCGList(3, c)).CGData(DBUCGList(1, c), DBUCGList(2, c)).MarginPer.OriginalForecast * RCV(DBUCGList(1, c))(DBUCGList(2, c))(DBUCGList(3, c))("Retail"))
                        sumMarginReFore = sumMarginReFore + (FCbM(YearNo, MonthNo, DBUCGList(3, c)).CGData(DBUCGList(1, c), DBUCGList(2, c)).MarginPer.ReForecast * RCV(DBUCGList(1, c))(DBUCGList(2, c))(DBUCGList(3, c))("Retail"))
                    End If
                End If
            End If
            sumContOriFore = sumContOriFore + FCbM(YearNo, MonthNo, DBUCGList(3, c)).CGData(DBUCGList(1, c), DBUCGList(2, c)).MarginDol.OriginalForecast
            sumContReFore = sumContReFore + FCbM(YearNo, MonthNo, DBUCGList(3, c)).CGData(DBUCGList(1, c), DBUCGList(2, c)).MarginDol.ReForecast
        End If
    Next
    
    
End Sub






