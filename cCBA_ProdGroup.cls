VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCBA_ProdGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit                         ' cCBA_ProdGroup
Private Const NATIONAL_DB As Long = 599
Private Const RDIVFROM  As Long = 501
Private Const RDIVTO As Long = 509
Private Const DebugProdGroupObjects As Boolean = False
''Private Enum e_DataCont
''    [_First]
''    eContractData = 0
''    eBraketData = 1
''    eIncoData = 2
''    ePricedata = 3
''    [_Last]
''End Enum

''Private Enum e_DataWaste
''    [_First]
''    eMarkdown = 0
''    eLoss = 1
''    [_Last]
''End Enum
'Enum eQtyOrRetail
'    eQty = 0
'    eRetail = 1
'End Enum
''Public Enum e_USWorUSWpp
''    eUSW = 0
''    eUSWP = 1
''End Enum
''Public Enum e_RCVMarginType
''    eRCVMarginPercent = 0
''    eRCVContributionDollar = 1
''End Enum


''Private pdblCarryOverValue As double
Private pbBDataGenerated As Boolean
Private psdTotBus As Scripting.Dictionary
Private pcnCCN As ADODB.Connection
Private pcnSCN As ADODB.Connection
Private pcnSMCN As ADODB.Connection
Private pcnCRCN As ADODB.Connection
Private pcnRCN(RDIVFROM To RDIVTO) As ADODB.Connection

Private prsSalesRS As ADODB.Recordset
Private prsBaseProdRS As ADODB.Recordset
Private prsContRS(RDIVFROM To RDIVTO) As ADODB.Recordset

Private psdRCost(RDIVFROM To RDIVTO) As Scripting.Dictionary
Private psdRegDivMiss As Scripting.Dictionary
''Private psdUserList As Scripting.Dictionary
''Private psdEtoNUList As Scripting.Dictionary
Private psdCOMdic As Scripting.Dictionary
Private psdCG_PDic As Scripting.Dictionary
Private psdProductLabelDic As Scripting.Dictionary
Private psdWastageIndex As Scripting.Dictionary

''Private pcolCGSelection As Collection
Private pcolCGs As Collection

Private pvCGList() As Variant
Private pvACGList() As Variant
Private pbACG As Boolean
''Private pbB_RdyToRun As Boolean
Private pdtMnthFrom As Date
Private pdtMnthTo As Date

Private psBuildCode As String
''Private pTimeTEST As double
Private plPcode As Long
Private bRunFordoubleProd As Boolean
Private psdAlcoholStoreNos As Scripting.Dictionary
Private pbHasAlcoholStores As Boolean
Private ProdToMSegData() As Variant
Private pbHasProdToMSegData As Boolean
Private psdSalesIndex As Scripting.Dictionary
Private psdSalesDivIndex As Scripting.Dictionary
Private psdUSWIndex As Scripting.Dictionary
Private psdUSWDivIndex As Scripting.Dictionary
Private psdSTIndex As Scripting.Dictionary
Private pbSalesBool As Boolean
Private pbSalesDivBool As Boolean
Private pbSalesUSWBool As Boolean
Private pbSalesUSWDivBool As Boolean
Private pbSalesSTBool As Boolean
Private plDoc_ID As Long



Private Enum e_CBA_PGQ
    [_First]
    eRCVQTY = 0
    eRCVRetail = 1
    eRCVRetailNet = 2
    eRCVCost = 3
    eRCVMargin = 4
    eRCVContribution = 5
    eRCVUSW = 6
    ePOSQTY = 7
    ePOSRetail = 8
    ePOSUSW = 9
    eWastage = 10
    eWeeksInPeriodRCV = 11
    eAdjRCVMargin = 12
    eAdjRCVContribution = 13
    eProdCountbyWeekNo = 14
    eMaxNoOfProducts = 15
    eWeeksInPeriodPOS = 16
    ePOSUPSbyWeek = 17
    ePOSUPSbyWeekpProduct = 18
    ePOSUSWpProd = 19
    eSellThrough = 20
    'ePOSUSWforaUniqueProd = 20
    [_Last]
End Enum
Public Function getProductState(ByVal productcode As Long, Optional ByVal shouldProdBeIncluded As Boolean = True) As Boolean
    If sdCG_PDic.Exists(CStr(productcode)) = shouldProdBeIncluded Then getProductState = True Else getProductState = False
End Function
Public Function getRCVMargin(ByVal RCVMarginType As e_RCVMarginType, ByVal DateFrom As Date, ByVal DateTo As Date, Optional ByVal PClass As Long, Optional DivNo As Long = 599, Optional CGno As Long, Optional SCGno As Long) As Double
    If RCVMarginType = e_RCVMarginType.eRCVMarginPercent Then
        getRCVMargin = QueryData(e_CBA_PGQ.eRCVMargin, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateFrom, DateTo)
    Else
        getRCVMargin = QueryData(e_CBA_PGQ.eRCVContribution, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateFrom, DateTo)
    End If
End Function
Public Function getRCV(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal QtyCostRetailOrNet As e_RCVQTYCostRetailNet, Optional ByVal PClass As Long, Optional DivNo As Long = 599, Optional CGno As Long, Optional SCGno As Long) As Double
    If QtyCostRetailOrNet = e_RCVQTYCostRetailNet.eRCVQTY Then getRCV = QueryData(e_CBA_PGQ.eRCVQTY, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateFrom, DateTo)
    If QtyCostRetailOrNet = e_RCVQTYCostRetailNet.eRCVRetail Then getRCV = QueryData(e_CBA_PGQ.eRCVRetail, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateFrom, DateTo)
    If QtyCostRetailOrNet = e_RCVQTYCostRetailNet.eRCVRetailNet Then getRCV = QueryData(e_CBA_PGQ.eRCVRetailNet, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateFrom, DateTo)
    If QtyCostRetailOrNet = e_RCVQTYCostRetailNet.eRCVCost Then getRCV = QueryData(e_CBA_PGQ.eRCVCost, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateFrom, DateTo)
End Function
Public Function getPOS(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal QtyOrRetail As e_RetailorQTY, Optional ByVal PClass As Long, Optional DivNo As Long = 599, Optional CGno As Long, Optional SCGno As Long) As Double
    If QtyOrRetail = e_RetailorQTY.eQTY Then getPOS = QueryData(e_CBA_PGQ.ePOSQTY, , , CGno, SCGno, PClass, DivNo, , , , , , , QtyOrRetail, , , , , , DateFrom, DateTo)
    If QtyOrRetail = e_RetailorQTY.eRetail Then getPOS = QueryData(e_CBA_PGQ.ePOSRetail, , , CGno, SCGno, PClass, DivNo, , , , , , , QtyOrRetail, , , , , , DateFrom, DateTo)
End Function
Public Function getUSW(ByVal USWType As e_USWorUSWpp, ByVal DateFrom As Date, ByVal DateTo As Date, ByVal QtyOrRetail As e_RetailorQTY, Optional ByVal PClass As Long, Optional DivNo As Long = 599, Optional CGno As Long, Optional SCGno As Long) As Double
Dim USW_Type As e_CBA_PGQ, asQty As Boolean
    If DivNo <> 599 And InStr(1, sBuildCode, "Ur") = 0 Then Debug.Print "Cannot Query USW by region without 'Ur' dataset": Exit Function
    If USWType = e_USWorUSWpp.eUSW Then USW_Type = e_CBA_PGQ.ePOSUSW Else USW_Type = e_CBA_PGQ.ePOSUSWpProd
    If QtyOrRetail = e_RetailorQTY.eQTY Then asQty = True Else asQty = False
    getUSW = QueryData(USW_Type, , , CGno, SCGno, PClass, DivNo, , , , , , , asQty, , , , , , DateFrom, DateTo)
End Function
Public Function getTotalBusiness(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal ValueNeeded As String) ', Optional ByVal Pclass As Long = 0, Optional DivNo As Long = 599) As Double
Dim d As Double
Dim yr As Long, mth As Long
    For yr = Year(DateFrom) To Year(DateTo)
        For mth = 1 To 12
            If DateSerial(yr, mth + 1, 0) >= DateFrom And DateSerial(yr, mth, 1) <= DateTo Then
                If sdTotBus(CStr(yr))(CStr(mth)).Exists(ValueNeeded) Then d = d + sdTotBus(CStr(yr))(CStr(mth))(ValueNeeded)
            End If
        Next
    Next
    getTotalBusiness = d
End Function
Public Function getSellThrough(ByVal DateFrom As Date, ByVal DateTo As Date, Optional ByVal PClass As Long, Optional CGno As Long, Optional SCGno As Long) As Double
    getSellThrough = QueryData(eSellThrough, , , , , PClass)
End Function
Private Function QueryData(ByVal QueryType As e_CBA_PGQ, Optional ByVal LongVal1 As Long, Optional ByVal LongVal2 As Long _
    , Optional ByVal CGno As Long, Optional ByVal SCGno As Long, Optional ByVal PClass As Long, Optional ByVal DivNo As Long = 599 _
    , Optional ByVal StringVal1 As String, Optional ByVal StringVal2 As String, Optional ByVal StringVal3 As String, Optional ByVal StringVal4 As String _
    , Optional ByVal StringVal5 As String, Optional ByVal StringVal6 As String, Optional ByVal BooleanVal1 As Boolean, Optional ByVal BooleanVal2 As Boolean _
    , Optional ByVal BooleanVal3 As Boolean, Optional ByVal BooleanVal4 As Boolean, Optional ByVal BooleanVal5 As Boolean, Optional ByVal BooleanVal6 As Boolean _
    , Optional ByVal DateVal1 As Date, Optional ByVal DateVal2 As Date, Optional ByVal DateVal3 As Date, Optional ByRef ObjectVal1 As Object, Optional ByRef ObjectVal2 As Object) As Double
    Dim v As Variant
    Dim OutPutVal As Double
    Dim PD As cCBA_Prod
    Dim cnt As Long, wkcnt As Long, wkfrom As Long, wkto As Long, yrfrom As Long, yrto As Long, yr As Long, wk As Long, pwkfrom As Long, pwkto As Long
    Dim S1 As Double, S2 As Double
    Dim tempno As Double
    On Error GoTo Err_Routine
    CBA_Error = ""
    cnt = 0
    If DivNo = 0 Then DivNo = 599
    For Each v In sdCG_PDic
'        If v = "59062" Then
'            Stop
'        End If
        'If v <> StringVal1 And QueryType = ePOSUSWforaUniqueProd Then GoTo NextOne
        Set PD = sdCG_PDic(v)
        If PClass > 0 Then
            If PClass = 5 Then
                'food and non food specials
                If PD.lPClass = 2 Or PD.lPClass = 3 Then Else GoTo NextOne
            Else
                If PD.lPClass <> PClass Then GoTo NextOne
            End If
        End If
        If CGno > 0 Then If PD.lLegCG <> CGno Then GoTo NextOne
        If CGno > 0 And SCGno > 0 Then If PD.lLegSCG <> SCGno Then GoTo NextOne
        
        Select Case QueryType
            Case 0
                OutPutVal = OutPutVal + PD.getRCVdata(DateVal1, DateVal2, "QTY", DivNo)
            Case 1
                OutPutVal = OutPutVal + PD.getRCVdata(DateVal1, DateVal2, "Retail", DivNo)
            Case 2
                OutPutVal = OutPutVal + PD.getRCVdata(DateVal1, DateVal2, "RetailNet", DivNo)
            Case 3
                OutPutVal = OutPutVal + PD.getRCVdata(DateVal1, DateVal2, "Cost", DivNo)
            Case 6
                tempno = PD.getRCVdata(DateVal1, DateVal2, "USWALL", DivNo)
                If tempno <> 0 Then cnt = cnt + 1
                OutPutVal = OutPutVal + tempno
            Case 7
                OutPutVal = OutPutVal + PD.getPOSdata(DateVal1, DateVal2, True, , , , DivNo)
                If QueryType = 20 And tempno <> 0 Then cnt = cnt + 1
            Case 8
                OutPutVal = OutPutVal + PD.getPOSdata(DateVal1, DateVal2, False, , , , DivNo)
            
            Case 10
                OutPutVal = OutPutVal + PD.getWastageData(e_RetailorQTY.eRetail, e_InvMDALL.eBoth, DateVal1, DateVal2, DivNo) '@TP NEED TO FIX
            Case 11
                OutPutVal = PD.getRCVdata(DateVal1, DateVal2, "USWCNT", DivNo)
                If OutPutVal <> 0 Then Exit For
            Case 14
                OutPutVal = OutPutVal + PD.getPOSdata(DateVal1, DateVal2, True, e_POSUSWTypes.eUSWisActive, LongVal1, LongVal2, DivNo)
            Case 16
                OutPutVal = PD.getPOSdata(DateVal1, DateVal2, True, e_POSUSWTypes.eUSWCNT, , , DivNo)
                If OutPutVal <> 0 Then Exit For
            Case 17
                tempno = PD.getPOSdata(DateVal1, DateVal2, BooleanVal1, e_POSUSWTypes.eUSWALL, LongVal1, LongVal2, DivNo)
                OutPutVal = OutPutVal + tempno
            Case 18
                tempno = PD.getPOSdata(DateVal1, DateVal2, BooleanVal1, e_POSUSWTypes.eUSWALL, LongVal1, LongVal2, DivNo)
                If tempno > 0.9 Then cnt = cnt + 1
                OutPutVal = OutPutVal + tempno
            Case 20
''                PD.getSellThrough DateVal1, DateVal2, s1, s2
            Case 21
            
            Case 22
            
            Case 23
            
            Case 24
            
            Case 25
            
            Case 26
            
            Case 27
            
            Case 28
            
            Case 29
            
            Case 30
            
            Case 31
            
            Case 32
        
        End Select
NextOne:
    Next
    
    Select Case QueryType
        Case 4
            tempno = QueryData(e_CBA_PGQ.eRCVRetail, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateVal1, DateVal2)
            If tempno <> 0 Then
                OutPutVal = QueryData(e_CBA_PGQ.eRCVContribution, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateVal1, DateVal2) / tempno
            Else
                OutPutVal = 0
            End If
        Case 5
            tempno = QueryData(e_CBA_PGQ.eRCVRetailNet, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateVal1, DateVal2) - QueryData(e_CBA_PGQ.eRCVCost, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateVal1, DateVal2)
            OutPutVal = tempno
        'Case 6, 9
        '    wkcnt = QueryData(eWeeksInPeriodPOS, , , , , , , , , , , , , , , , , , , DateVal1, DateVal2)
        '   If cnt = 0 Or wkcnt = 0 Then OutPutVal = 0 Else OutPutVal = OutPutVal / (cnt * wkcnt)
        Case 12
            tempno = QueryData(e_CBA_PGQ.eRCVRetail, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateVal1, DateVal2)
            If tempno <> 0 Then
                OutPutVal = QueryData(e_CBA_PGQ.eAdjRCVContribution, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateVal1, DateVal2) / tempno
            Else
                OutPutVal = 0
            End If
        Case 13
            OutPutVal = QueryData(e_CBA_PGQ.eRCVContribution, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateVal1, DateVal2) + QueryData(e_CBA_PGQ.eWastage, , , CGno, SCGno, PClass, DivNo, , , , , , , , , , , , , DateVal1, DateVal2)
        Case 9, 15, 19
            wkfrom = CBA_BasicFunctions.CBA_IsoWeekNumber(DateVal1)
            If DateVal1 >= CBA_BasicFunctions.CBA_IsoYearStart(Year(DateVal1)) Then yrfrom = Year(DateVal1) Else yrfrom = Year(DateVal1) - 1
            wkto = CBA_BasicFunctions.CBA_IsoWeekNumber(DateVal2)
            If DateVal2 < CBA_BasicFunctions.CBA_IsoYearStart(Year(DateVal2)) Then yrto = Year(DateVal2) + 1 Else yrto = Year(DateVal2)
            If Month(DateVal2) > 10 And wkto < 4 Then yrto = yrto + 1
            For yr = yrfrom To yrto
                If yrfrom = yrto Then
                    pwkfrom = wkfrom: pwkto = wkto
                Else
                    pwkfrom = 1: pwkto = 53
                    If yr = yrfrom Then pwkfrom = wkfrom
                    If yr = yrto Then pwkto = wkto
                End If
                For wk = pwkfrom To pwkto
                    If QueryType = 9 Or QueryType = 19 Then
                        If QueryType = 9 Then
                            If StringVal1 <> "" Then
                                tempno = QueryData(e_CBA_PGQ.ePOSUPSbyWeek, wk, yr, CGno, SCGno, PClass, DivNo, StringVal1, , , , , , BooleanVal1, , , , , , DateVal1, DateVal2)
                            Else
                                tempno = QueryData(e_CBA_PGQ.ePOSUPSbyWeek, wk, yr, CGno, SCGno, PClass, DivNo, , , , , , , BooleanVal1, , , , , , DateVal1, DateVal2)
                            End If
                        End If
                        If QueryType = 19 Then tempno = QueryData(e_CBA_PGQ.ePOSUPSbyWeekpProduct, wk, yr, CGno, SCGno, PClass, DivNo, , , , , , , BooleanVal1, , , , , , DateVal1, DateVal2)
                        If tempno <> 0 Then OutPutVal = OutPutVal + tempno: cnt = cnt + 1
                    End If
                    If QueryType = 15 Then
                        tempno = QueryData(e_CBA_PGQ.eProdCountbyWeekNo, wk, yr, CGno, SCGno, PClass, DivNo, , , , , , , BooleanVal1, , , , , , DateVal1, DateVal2)
                        If tempno > OutPutVal Then OutPutVal = tempno
                    End If
                Next
            Next
            If (QueryType = 9 Or QueryType = 19) And OutPutVal <> 0 And IsEmpty(OutPutVal) = False Then OutPutVal = NZ(OutPutVal / cnt, 0)
        Case 18 ', 20
            If cnt > 0 Then OutPutVal = OutPutVal / cnt
    End Select
    If QueryType = eSellThrough Then
        If S1 = 0 Then QueryData = 0 Else QueryData = S2 / S1
    Else
        QueryData = NZ(OutPutVal, 0)
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_ProdGroup.QueryData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Com", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Function getProdObject(ByVal lPcode As String, Optional ByVal AddUSWDivData As Boolean = False) As cCBA_Prod
Dim v As Variant
Dim P As cCBA_Prod
Dim strSQL As String
Dim CN As cCBA_Connect
Dim curProd As Long, a As Long, b As Long, curYear As Long, cnt As Long, totcnt As Long
Dim yearDic As Scripting.Dictionary
Dim PeriodDic As Scripting.Dictionary
Dim psdSalesIndex As Scripting.Dictionary
Dim FieldDic As Scripting.Dictionary
Dim SalesDic As Scripting.Dictionary
Dim DivDic As Scripting.Dictionary
Dim curDiv As Long
Dim RS As ADODB.Recordset
Dim SD As cCBA_ProdPeriodSalesData
'    For Each v In sdCG_PDic
'        Debug.Print v
'    Next
    If sdCG_PDic.Exists(lPcode) Then
        If InStr(1, sdCG_PDic(lPcode).sBldCode, "Ur") = 0 And AddUSWDivData = True Then
            Set CN = New cCBA_Connect
            If CN.SetConnection("CBIS") = False Then Exit Function
            strSQL = "SET NOCOUNT ON" & Chr(10)
            strSQL = strSQL & "SET ANSI_WARNINGS OFF" & Chr(10)
            strSQL = strSQL & "declare @YRST nvarchar(4) = " & Year(Date) - 3 & Chr(10)
            strSQL = strSQL & "" & Chr(10)
            strSQL = strSQL & "select productcode into #TEMP from cbis599p.dbo.product where productcode = " & lPcode & Chr(10)
            strSQL = strSQL & "" & Chr(10)
            strSQL = strSQL & "select YEAR(DATEADD(day, 26 - DATEPART(isoww, pdate), pdate)) as yrno," & Chr(10)
            strSQL = strSQL & "DatePart(ISOWK, PDATE) as Wkno , count(pdate) as daycnt into #DATEWK  " & Chr(10)
            strSQL = strSQL & "from (SELECT  TOP (DATEDIFF(D, @YRST + '-01-01', EOMONTH(getdate(),-1)) + 1)" & Chr(10)
            strSQL = strSQL & "    Pdate = convert(date,DATEADD(d, ROW_NUMBER() OVER(ORDER BY a.object_id) - 1, @YRST + '-01-01'))" & Chr(10)
            strSQL = strSQL & "    FROM    sys.all_objects a CROSS JOIN sys.all_objects b ) a" & Chr(10)
            strSQL = strSQL & "group by YEAR(DATEADD(day, 26 - DATEPART(isoww, pdate), pdate)), DatePart(ISOWK, PDATE)" & Chr(10)
            strSQL = strSQL & "order by yrno, wkno" & Chr(10)
            strSQL = strSQL & "" & Chr(10)
            strSQL = strSQL & "select productcode into #PList from (select distinct rcv.productcode from cbis599p.dbo.receiving rcv inner join #TEMP p on p.ProductCode = rcv.ProductCode where year(DayEndDate) >= @YRST" & Chr(10)
            strSQL = strSQL & "union select distinct pos.productcode from cbis599p.dbo.pos pos inner join #TEMP p on p.ProductCode = pos.ProductCode where year(posdate) >= @YRST) a" & Chr(10)
            strSQL = strSQL & "" & Chr(10)
            strSQL = strSQL & "select p.productcode, p.CGNo as LCG, p.scgno as LSCG, acg.Category, acg.CatNo, ACG.CGNo, ACG.SCGNo, ProductClass, p.TaxID, P.EmpNo,P.Con_ProductCode" & Chr(10)
            strSQL = strSQL & ",P.isRandomWeight, P.Packsize, P.PriceCrdBrand, P.SeasonId,P.Trademarkno" & Chr(10)
            strSQL = strSQL & "into #P from cbis599p.dbo.product p" & Chr(10)
            strSQL = strSQL & "inner join #PList PL on pl.ProductCode = p.productcode" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.tf_ACGMap() acg on acg.ACGEntityID = p.ACGEntityID" & Chr(10)
            strSQL = strSQL & "" & Chr(10)
            strSQL = strSQL & "select Divno into #DIV from cbis599p.dbo.DIVISION where divno > '500' and divno < '510'" & Chr(10)
            strSQL = strSQL & "" & Chr(10)
            strSQL = strSQL & "select pos.DivNo, year(posdate) as YRno, datepart(ISOWK,posdate) as WKno, pos.productcode, sum(retail) as retail, sum(quantity) as QTY" & Chr(10)
            strSQL = strSQL & "into #POSWKDIV from cbis599p.dbo.pos pos" & Chr(10)
            strSQL = strSQL & "inner join #Plist P on p.ProductCode = pos.ProductCode" & Chr(10)
            strSQL = strSQL & "where YEAR(posdate) >= @YRST" & Chr(10)
            strSQL = strSQL & "group by pos.DivNo,year(posdate) , datepart(ISOWK,posdate) , pos.productcode" & Chr(10)
            strSQL = strSQL & "" & Chr(10)
            strSQL = strSQL & "select rcv.DivNo, year(dayenddate) as YRno, datepart(ISOWK,dayenddate) as Wkno, rcv.productcode , sum(Quantity) as Quantity , sum(Pieces) as Pieces" & Chr(10)
            strSQL = strSQL & "into #RCVQTYWKDIV from cbis599p.dbo.receiving rcv" & Chr(10)
            strSQL = strSQL & "inner join #PList p on p.ProductCode = rcv.ProductCode" & Chr(10)
            strSQL = strSQL & "where rcv.recordid = '001' and YEAR(dayenddate) >= @YRST" & Chr(10)
            strSQL = strSQL & "group by rcv.DivNo, year(dayenddate), datepart(ISOWK,dayenddate), rcv.productcode, rcv.recordid" & Chr(10)
            strSQL = strSQL & "" & Chr(10)
            strSQL = strSQL & "select rcv.DivNo, YEAR(dayenddate) as YRno, datepart(ISOWK,dayenddate) as Wkno, rcv.productcode , sum(retail) as Retail, sum(retailnet) as RetailNet, sum(cost) as Cost" & Chr(10)
            strSQL = strSQL & "into #RCVRETWKDIV from cbis599p.dbo.RECEIVING rcv" & Chr(10)
            strSQL = strSQL & "inner join #PList p on p.ProductCode = rcv.ProductCode" & Chr(10)
            strSQL = strSQL & "where YEAR(dayenddate) >= @YRST" & Chr(10)
            strSQL = strSQL & "group by rcv.DivNo, year(dayenddate),  datepart(ISOWK,dayenddate), rcv.productcode" & Chr(10)
            strSQL = strSQL & "" & Chr(10)
            strSQL = strSQL & "set datefirst 1" & Chr(10)
            strSQL = strSQL & "select Stores.DivNo, datepart(ISOWK,validfrom) as WeekNo, datepart(YEAR,validfrom) as YearNo" & Chr(10)
            strSQL = strSQL & ", dateadd(D,-2,validfrom) as WKStart, dateadd(D,4,validfrom) as WKEnd, sum(NoOfStores)  as NoOfStores" & Chr(10)
            strSQL = strSQL & "into #STORESWKDIV from cbis599p.portfolio.Stores where datepart(YEAR,validfrom) >= @YRST" & Chr(10)
            strSQL = strSQL & "group by Stores.DivNo, datepart(ISOWK,validfrom) , datepart(YEAR,validfrom), dateadd(D,-2,validfrom), dateadd(D,4,validfrom)" & Chr(10)
            strSQL = strSQL & "" & Chr(10)
            strSQL = strSQL & "select d.YRno , d.Wkno, p.productcode" & Chr(10)
            strSQL = strSQL & ", isnull(R.Retail / nullif(s.noofstores,0),0) as RCVRetailpwk, isnull(RetailNet/ nullif(s.noofstores,0),0) as RCVRetailNetpwk, isnull(Cost/ nullif(s.noofstores,0),0) as RCVCostpwk" & Chr(10)
            strSQL = strSQL & ", isnull(RC.Quantity/ nullif(s.noofstores,0), 0) as RCVQTYpwk, isnull(RC.Pieces/ nullif(s.noofstores,0), 0) as RCVPiecespwk, isnull(pos.QTY/ nullif(s.noofstores,0),0) as POSQTYpwk" & Chr(10)
            strSQL = strSQL & ", isnull(pos.Retail/ nullif(s.noofstores,0),0) as POSRetailpwk, isnull(d.daycnt,0) as DayCnt, isnull(s.noofstores,0) as NoOfStores,isnull(pr.CGNo,0) as CGno, isnull(div.divno,0) as Divno" & Chr(10)
            strSQL = strSQL & "from #DATEWK d cross join #PList p cross join #DIV div" & Chr(10)
            strSQL = strSQL & "left join #RCVRETWKDIV R on R.productcode = p.productcode and R.YRno = d.Yrno and R.wkno = d.wkno and r.DivNo =div.DivNo" & Chr(10)
            strSQL = strSQL & "left join #RCVQTYWKDIV RC on RC.ProductCode = p.ProductCode and RC.wkno = d.wkno and RC.YRno = d.YRno and rc.DivNo = div.DivNo" & Chr(10)
            strSQL = strSQL & "left join #POSWKDIV pos on pos.ProductCode = p.ProductCode and pos.wkno = d.wkno and pos.YRno = d.Yrno and pos.DivNo = div.DivNo" & Chr(10)
            strSQL = strSQL & "left join #STORESWKDIV s on s.WeekNo = d.wkno and s.YearNo = d.Yrno and s.DivNo = div.DivNo" & Chr(10)
            strSQL = strSQL & "left join #P pr on pr.productcode = p.ProductCode" & Chr(10)
            strSQL = strSQL & "order by div.divno, YRno, wkno, productcode" & Chr(10)
            strSQL = strSQL & "" & Chr(10)
            strSQL = strSQL & "drop table #DIV, #RCVQTYWKDIV, #RCVRETWKDIV, #POSWKDIV" & Chr(10)
            strSQL = strSQL & "drop table #PList, #DATEWK, #P, #STORESWKDIV, #TEMP" & Chr(10)
            If CN.Query(strSQL, "ProdUSWDivData" & lPcode) = False Then Exit Function
            curProd = 0
            Set RS = CN.op("ProdUSWDivData" & lPcode)
            Set SD = New cCBA_ProdPeriodSalesData
            SD.lPcode = lPcode
            a = 0
            Do Until RS.EOF
                a = a + 1
                RS.MoveNext
            Loop
            RS.MoveFirst
            SD.Build RS, a, "USWDiv", IIf(bHasAlcoholStores, sdAlcoholStoreNos, Nothing)
            Set sdCG_PDic(lPcode).cUSWDiv = SD
            sdCG_PDic(lPcode).sBldCode = sdCG_PDic(lPcode).sBldCode & "Ur"
        End If
        Set getProdObject = sdCG_PDic(lPcode)
    Else
        Set getProdObject = Nothing
    End If
End Function
Function getProdListing() As Collection
Dim col As Collection
Dim v As Variant
    Set col = New Collection
    For Each v In sdCG_PDic
        col.Add v
    Next
    Set getProdListing = col
End Function
Function getWastageData(ByVal isRetail As e_RetailorQTY, ByVal InventoryMarkdownBoth As e_InvMDALL, Optional ByVal dteFrom As Date, Optional ByVal dteTo As Date, Optional ByVal PCode As Long, Optional ByVal ProductClass As Long, Optional ByVal DivNo As Long = 599) As Double '@TP NEED TO FIX
Dim P As cCBA_Prod
Dim s As Double
Dim v As Variant
    s = 0
    If PCode > 0 Then
        Set P = sdCG_PDic(PCode)
        s = P.getWastageData(isRetail, InventoryMarkdownBoth, dteFrom, dteTo, DivNo)
    Else
        For Each v In sdCG_PDic
            If IsEmpty(v) Then GoTo NextOneNow
            Set P = sdCG_PDic(v)
            If ProductClass > 0 And P.lPClass <> ProductClass Then GoTo NextOneNow
            If InStr(1, P.sBldCode, "W") = 0 Then GoTo NextOneNow
            s = s + P.getWastageData(eRetail, InventoryMarkdownBoth, dteFrom, dteTo, DivNo)
NextOneNow:
        Next
    End If
    getWastageData = s
End Function
Function RunDataGeneration(ByVal dDateFrom As Date, ByVal dDateTo As Date, ByVal ACG As Boolean, ByRef cCGDist As Collection, ByVal BuildCode As String, Optional ByVal PCode As Long = 0) As Boolean
Dim v As Variant
    Set colCGs = cCGDist
    bACG = ACG
    dtMnthFrom = DateSerial(Year(dDateFrom), Month(dDateFrom) - 1, 1)
    dtMnthTo = DateSerial(Year(dDateTo), Month(dDateTo) + 1, 0)
    sBuildCode = BuildCode
    lPcode = PCode
    If InStr(1, sBuildCode, "As") = 0 And lPcode = 0 Then
        For Each v In colCGs
            If CLng(Mid(v, 1, 3)) < 5 Then AddToBuildCode "As": Exit For
        Next
    End If
    
    RunDataGeneration = GenProductDic
End Function
Private Function AddToBuildCode(ByVal RefVal As String) As Boolean
    If InStr(1, sBuildCode, RefVal) = 0 Then sBuildCode = sBuildCode & RefVal
End Function
Public Function setMSegAllocation(ByVal PCode As String, ByVal ScanDataNotManual As Boolean, ByVal MSegName As String) As Boolean
    If PCode = 0 Then Exit Function
    If sdCG_PDic.Exists(PCode) Then
        If ScanDataNotManual = True Then sdCG_PDic(PCode).ScanDataMSeg = MSegName Else sdCG_PDic(PCode).ManualMSeg 1, MSegName
    Else
        MsgBox "error: Product Not in ProdGroup Data"
        If DebugProdGroupObjects = True Then Debug.Print "error: Product Not in ProdGroup Data"
    End If
End Function
Public Function SaveMSegAllocationtoDB() As Boolean
Dim CN As cCBA_Connect
Dim v As Variant
Dim P As cCBA_Prod
Dim RS As ADODB.Recordset
Dim ManDic As Scripting.Dictionary
Dim ScanDic As Scripting.Dictionary
Dim bHasData As Boolean
Dim fArr(0 To 3) As Variant
Dim vArr(0 To 3) As Variant


    Set ManDic = New Scripting.Dictionary: Set ScanDic = New Scripting.Dictionary: Set CN = New cCBA_Connect
    fArr(0) = "PA_SegmentationMethod": fArr(1) = "PA_MSegName": fArr(2) = "PA_CAT_ID": fArr(3) = "PA_ProductCode"
    
    If CN.SetConnection("CAMERA", True) = False Then SaveMSegAllocationtoDB = False: Exit Function
HaveAnotherGo:
    If CN.Query("Select * from L1_ProdToMSegAllocation where PA_CAT_ID = " & mCAM_Runtime.cRibbonData.sdCategoryID(mCAM_Runtime.ActiveDO.sCategoryName), "UpdQueryALLMSegForcategoryID", True) = True Then
        bHasData = True
        Set RS = CN.op("UpdQueryALLMSegForcategoryID")
        Do Until RS.EOF
            If sdCG_PDic.Exists(CStr(RS(4))) Then
                Set P = sdCG_PDic(CStr(RS(4)))
                If RS(1) = 2 Then
                    If RS(2) <> P.sScanDataMSeg Then
                        If P.sScanDataMSeg = "" Then
                            RS.Delete
                        Else
                            RS(2) = P.sScanDataMSeg
                            ScanDic.Add CStr(RS(4)), CStr(RS(2))
                        End If
                    Else
                        ScanDic.Add CStr(RS(4)), CStr(RS(2))
                    End If
                ElseIf RS(1) = 3 Then
                    If RS(2) <> P.ManualMSeg(1) Then
                        If P.ManualMSeg(1) = "" Then
                            RS.Delete
                        Else
                            RS(2) = P.ManualMSeg(1)
                            ManDic.Add CStr(RS(4)), CStr(RS(2))
                        End If
                    Else
                        ManDic.Add CStr(RS(4)), CStr(RS(2))
                    End If
                End If
            End If
            RS.MoveNext
        Loop
        For Each v In sdCG_PDic
            Set P = sdCG_PDic(v)
            If P.ManualMSeg(1) <> "" And ManDic.Exists(CStr(P.lPcode)) = False Then
                vArr(0) = 3: vArr(1) = P.ManualMSeg(1): vArr(2) = mCAM_Runtime.cRibbonData.sdCategoryID(mCAM_Runtime.ActiveDO.sCategoryName): vArr(3) = P.lPcode
                RS.AddNew fArr, vArr
            End If
            If P.sScanDataMSeg <> "" And ScanDic.Exists(P.lPcode) = False Then
                vArr(0) = 2: vArr(1) = P.sScanDataMSeg: vArr(2) = mCAM_Runtime.cRibbonData.sdCategoryID(mCAM_Runtime.ActiveDO.sCategoryName): vArr(3) = P.lPcode
                RS.AddNew fArr, vArr
            End If
        Next
        RS.UpdateBatch adAffectAllChapters
    End If
    If bHasData = False Then
        For Each v In sdCG_PDic
            Set P = sdCG_PDic(v)
            If P.ManualMSeg(1) <> "" Then
                CN.Query "Insert Into L1_ProdToMSegAllocation (PA_SegmentationMethod,PA_MSegName,PA_CAT_ID,PA_ProductCode) Values(" & 3 & ",'" & P.ManualMSeg(1) & "'," & mCAM_Runtime.cRibbonData.sdCategoryID(mCAM_Runtime.ActiveDO.sCategoryName) & "," & P.lPcode & ")", "InsertProdAllocateWhereThereIsNone", True: GoTo HaveAnotherGo
            ElseIf P.sScanDataMSeg <> "" Then
                CN.Query "Insert Into L1_ProdToMSegAllocation (PA_SegmentationMethod,PA_MSegName,PA_CAT_ID,PA_ProductCode) Values(" & 2 & ",'" & P.sScanDataMSeg & "'," & mCAM_Runtime.cRibbonData.sdCategoryID(mCAM_Runtime.ActiveDO.sCategoryName) & "," & P.lPcode & ")", "InsertProdAllocateWhereThereIsNone", True: GoTo HaveAnotherGo
            End If
        Next
    End If

End Function
Public Function CompareBuildCode(ByVal bcode As String) As Boolean
Dim s As String, a As Long
    For a = 1 To Len(bcode)
        If Mid(bcode, a + 1, 1) = LCase(Mid(bcode, a + 1, 1)) Then s = Mid(bcode, a, 2) Else s = Mid(bcode, a, 1)
        If InStr(1, sBuildCode, s) = 0 Then Exit Function
    Next
    CompareBuildCode = True
End Function


Private Function GenProductDic() As Boolean
    '''''''''''''''''''''''''''''''***********************
    ''''''''''''NOTES FROM ARCHITECTURE DESIGN''''''''''''
    '   The Market Segemnt (MSEG) allocation must be pulled as part of the buildcode fucntions
    '   the Line allocation should be able to be calculated on generaion
    '
    '
    '
    '''''''''''''''''''''''''''''''***********************
    
    
    
    '''     BUILDCODE OPTIONS
    '''     B = Base Product Data (Description, Packsize, CG, SCG, ACAT, ACG, ASCG, BD, GBD) ' Mandatory
    '''     C = Contract Data ()
    '''     R = Receiving Data
    '''     P = POS Data
    '''     Ur = USW data at a Rgeional Level
    '''     M = POS Margin (Cost) Data
    '''     T = Total Business Metrics (Share)
    '''     Cm = COMRADE data
    '''     F = Forecasting Data
    '''     L = Product Label Dictionary
    '''     W = Wastage
    '''     Ms = Market Segment Allocation
    '''     As = Alcohol Store Numbers included (automatically added)


    Dim YearStart As Long, a As Long, div As Long, totprods As Long
    Dim totbusBool As Boolean, SalesBool As Boolean, POSMarginBool As Boolean, BaseBool As Boolean, ContractBool As Boolean, ComradeBool As Boolean, ForecastBool As Boolean, PLabelBool As Boolean, WastageBool As Boolean, ProdToMsegBool As Boolean
    Dim CRS As ADODB.Recordset, EmptyRS As ADODB.Recordset
    Dim strSQL As String
    Dim C_P As cCBA_Prod
    Dim FirstName  As String, LastName As String, strPlist As String
    Dim PList() As Variant
    Dim ProductDic As Scripting.Dictionary, EmptyDic As Scripting.Dictionary
    Dim curRnd As Double, ter As Double, Rate As Double
    Dim col As Collection, EmptyCol As Collection
    Dim curPCode As Long
    Dim TempArr As Variant, EmptyArr(0 To 0, 0 To 0) As Variant
    Dim ContData(0 To 3) As Variant
    Dim ContIndex(0 To 3) As Scripting.Dictionary
    Dim cont As Variant
    Dim TempRS As ADODB.Recordset
    Dim TimeTEST As Double
    Dim sdS As Scripting.Dictionary, sdU As Scripting.Dictionary, sdSD As Scripting.Dictionary
    Dim sdUD As Scripting.Dictionary, sdST As Scripting.Dictionary
    Dim CMcol As Collection
    Dim cWaste As cCBA_WastageData
    Dim sContents As String, sDesc As String

    
    '''''''-----------TEST VARIABLES---------------------
    'Variables for Development
    'Dim ACG As Boolean
    'Dim CGDist(0 To 1, 0 To 2) As Long ' 0 = CGno , 1 = SCGNo
    'Dim NATIONAL_DB As Long, RDivFrom  As Long, RDivto As Long
    'Dim sdRegDivMiss As Scripting.Dictionary
    
    
    'ACG = False
    'Set CGDist(0) = New Scripting.Dictionary
    'CGDist(0).Add 51, 3
    'CGDist(0).Add 53, 4
    
    'ACG = True
    'CGDist(0, 0) = 74: CGDist(1, 0) = 1
    'CGDist(0, 1) = 1: CGDist(1, 1) = 2
    'CGDist(0, 2) = 74: CGDist(1, 2) = 2
    'CGDist(0, 3) = 51: CGDist(1, 3) = 2
    'NATIONAL_DB = 599
    'RDivFrom = 501
    'RDivto = 509
    On Error GoTo Err_Routine
    CBA_Error = ""
    Set sdRegDivMiss = New Scripting.Dictionary
    sdRegDivMiss.Add 508, 508
    
    
    'BuildCode = "BCTRMP"
    'BuildCode = "BC"
    'End of test variables
    '''''''-----------END TEST VARIABLES---------------------
    EmptyArr(0, 0) = 0
    Set EmptyDic = New Scripting.Dictionary
    Set EmptyRS = New ADODB.Recordset
    Set EmptyCol = New Collection
    YearStart = Year(Date) - 3
    If InStr(1, sBuildCode, "B") = 0 Then
        MsgBox "Base Data Not Requested in buildcode"
        Exit Function
    End If
    
    'Set cnSCN = New ADODB.Connection
    'With cnSCN
    '    .ConnectionTimeout = 50
    '    .CommandTimeout = 50
    '    .Open "Provider=SQLNCLI10;DATA SOURCE=599DBL12;;INTEGRATED SECURITY=sspi;"
    'End With
    'Set cnSMCN = New ADODB.Connection
    'With cnSMCN
    '    .ConnectionTimeout = 50
    '    .CommandTimeout = 50
    '    .Open "Provider=SQLNCLI10;DATA SOURCE=599DBL11;;INTEGRATED SECURITY=sspi;"
    'End With
    If RunConnectionSetups(True, False, False, False, False) = False Then Exit Function
    
    If InStr(1, sBuildCode, "B") > 0 Then BaseBool = buildBaseData
    Set CRS = New ADODB.Recordset
    strSQL = "select distinct productcode from #BASEPRODDATA order by productcode"
    CRS.Open strSQL, cnCCN
    If CRS.EOF Then
        MsgBox "NoBaseDataReturned"
        GenProductDic = False
        Exit Function
    Else
        PList = CRS.GetRows()
    End If
        
    '    ReDim CBA_Prods(1 To totprods)
    Set CRS = Nothing
    
    
    If InStr(1, sBuildCode, "M") > 0 Or InStr(1, sBuildCode, "c") > 0 Or InStr(1, sBuildCode, "W") > 0 Then
        If RunConnectionSetups(False, True, False, False, False) = False Then Exit Function
    End If
    
    If InStr(1, sBuildCode, "T") > 0 And (InStr(1, sBuildCode, "R") > 0 Or InStr(1, sBuildCode, "P") > 0) Then totbusBool = buildTotBusData(YearStart, sBuildCode): DoEvents

    'TimeTEST = Timer
    If InStr(1, sBuildCode, "c") > 0 Then ContractBool = buildContractData(YearStart): DoEvents
    If ContractBool Then
        For a = 0 To 3
            ContData(a) = pullALLContData(a): DoEvents
            Set ContIndex(a) = prepareALLContIndex(ContData(a), a): DoEvents
        Next
    End If
    If InStr(1, sBuildCode, "W") > 0 Then WastageBool = buildWastageData(YearStart): DoEvents
    
'    If WastageBool Then
'        For a = 0 To 1
'            vWastageData(a) = pullALLWastageData(a): DoEvents
'            Set sdWastageIndex(a) = prepareALLWastageIndex(ContData(a), a): DoEvents
'        Next
'    End If
    If DebugProdGroupObjects = True Then Debug.Print cont; TimeTEST - Timer
    If InStr(1, sBuildCode, "Cm") > 0 Then ComradeBool = buildComradeData(PList): DoEvents
    If ComradeBool = False Then Set sdCOMdic = New Scripting.Dictionary
    If InStr(1, sBuildCode, "L") = 0 Then Set sdProductLabelDic = New Scripting.Dictionary Else PLabelBool = buildProductLabelData: DoEvents
    If InStr(1, sBuildCode, "Ms") > 0 Then
        If RunConnectionSetups(True, False, False, False, True) = False Then Exit Function
        ProdToMsegBool = buildProdToMSegData(PList): DoEvents
    End If
    
    Set psdCG_PDic = New Scripting.Dictionary
    curRnd = 0: If DebugProdGroupObjects = True Then ter = Timer
    For a = LBound(PList, 2) To UBound(PList, 2)
        If DebugProdGroupObjects = True Then TimeTEST = Timer

        'If UBound(PList, 2) > 0 Then If Round((a / UBound(PList, 2)), 2) > curRnd Then Rate = (Timer - ter): curRnd = Round((a / UBound(PList, 2)), 2): ter = Timer
        

        If a / Round(UBound(PList, 2) / 20, 0) = Round(a / Round(UBound(PList, 2) / 20, 0), 0) Then If CBA_BasicFunctions.isRunningSheetDisplayed = True And UBound(PList, 2) > 0 Then CBA_BasicFunctions.RunningSheetAddComment 1, 1, "Generating ALDI Product DataCube: " & Round((a / UBound(PList, 2)), 2) * 100 & "%" '''  : Est. Time to Complete: " & CStr(Round((Rate * (1 - Round((a / UBound(PList, 2)), 2))) * 100 / 60, 2)) & " minutes": DoEvents

        Set C_P = New cCBA_Prod
        
        'If InStr(1, sBuildCode, "c") > 0 Then If IsEmpty(sdCOMdic(CStr(PList(0, a)))) Then Set col = Nothing Else Set col = sdCOMdic(CStr(PList(0, a)))
        
'        If PList(0, a) = 76002 Then
'        a = a
'        End If
        
'Private psdSalesIndex As Scripting.Dictionary
'Private psdSalesDivIndex As Scripting.Dictionary
'Private psdUSWIndex As Scripting.Dictionary
'Private psdUSWDivIndex As Scripting.Dictionary
'Private psdSTIndex As Scripting.Dictionary
        
        
        
'        If psdSalesIndex.Exists(PList(0, a)) Then Set sdS = psdSalesIndex(PList(0, a)) Else Set sdS = Nothing
        
        If psdCOMdic Is Nothing Then
            Set CMcol = EmptyCol
        ElseIf psdCOMdic.Exists(CStr(PList(0, a))) Then
            Set CMcol = psdCOMdic(CStr(PList(0, a)))
        Else
            Set CMcol = EmptyCol
        End If
        
        If psdWastageIndex Is Nothing Then
            Set cWaste = New cCBA_WastageData
        ElseIf psdWastageIndex.Exists(CLng(PList(0, a))) Then
            Set cWaste = psdWastageIndex(CLng(PList(0, a)))
            Set TempRS = pullBaseData(PList(0, a), psBuildCode)
            If TempRS(18) = 4 Then
                If IsNull(TempRS(6)) Then sContents = "" Else sContents = TempRS(6)
                If IsNull(TempRS(1)) Then sDesc = "" Else sDesc = TempRS(1)
                If InStr(1, sDesc, "24pk") > 0 Or InStr(1, sContents, "24pk") > 0 Or InStr(1, sContents, "24x") > 0 Or InStr(1, sContents, "24 x") > 0 Then cWaste.bis24PackBeer = True
                If InStr(1, sDesc, "6pk") > 0 Or InStr(1, sContents, "6pk") > 0 Or InStr(1, sContents, "6x") > 0 Or InStr(1, sContents, "6 x") > 0 Then cWaste.bis6PackBeer = True
            End If
        Else
            Set cWaste = New cCBA_WastageData
        End If
        
        
        
        
        If DebugProdGroupObjects = True Then Debug.Print "Start Build Prod Object: " & Timer - TimeTEST
        
        C_P.Build _
            PList(0, a), _
            psBuildCode, _
            IIf(totbusBool, psdTotBus, EmptyDic), _
            IIf(BaseBool, pullBaseData(PList(0, a), psBuildCode), EmptyRS), _
            IIf(ContractBool, psdRCost, EmptyDic), _
            IIf(ContractBool, psdRegDivMiss, EmptyDic), _
            NATIONAL_DB, _
            RDIVFROM, _
            RDIVTO, _
            IIf(ComradeBool, CMcol, EmptyCol), _
            IIf(PLabelBool, psdProductLabelDic(CStr(PList(0, a))), ""), _
            IIf(ContractBool, ContData, EmptyArr), _
            IIf(ContractBool, ContIndex, EmptyDic), _
            IIf(WastageBool, cWaste, Nothing), _
            IIf(ProdToMsegBool, ProdToMSegData, EmptyArr)


        
        'IIf(psdCOMdic.Exists(CStr(PList(0, a))), psdCOMdic(CStr(PList(0, a))), EmptyCol) '
        
        If DebugProdGroupObjects = True Then Debug.Print "Completed Not added to Dic Prod Object: " & Timer - TimeTEST

        psdCG_PDic.Add CStr(PList(0, a)), C_P
        
        If DebugProdGroupObjects = True Then Debug.Print "Completed Prod Object: " & Timer - TimeTEST
        
    Next
    
    If InStr(1, sBuildCode, "R") > 0 Or InStr(1, sBuildCode, "P") > 0 Then
'        SalesBool = BuildSalesDataRS(YearStart, sBuildCode)
'        DoEvents
'    End If
'    'If InStr(1, sBuildCode, "M") > 0 Then POSMarginBool = True: buildPOSMarginData YearStart, colCGs, ACG: DoEvents
'    If SalesBool = False Then
'        'Handle no sales data returned
'    Else
    SalesBool = buildSalesDataRS(YearStart, sBuildCode, "Sales")
    For a = LBound(PList, 2) To UBound(PList, 2)
        If a / Round(UBound(PList, 2) / 20, 0) = Round(a / Round(UBound(PList, 2) / 20, 0), 0) Then If CBA_BasicFunctions.isRunningSheetDisplayed = True And UBound(PList, 2) > 0 Then CBA_BasicFunctions.RunningSheetAddComment 1, 1, "Distributing National Sales DataCube: " & Round((a / UBound(PList, 2)), 2) * 100 & "%": DoEvents
        If pbSalesBool And IsEmpty(psdSalesIndex(PList(0, a))) = False Then Set psdCG_PDic(CStr(PList(0, a))).cSales = psdSalesIndex(PList(0, a))
    Next
    Set psdSalesIndex = Nothing
    
    SalesBool = buildSalesDataRS(YearStart, sBuildCode, "USW")
    For a = LBound(PList, 2) To UBound(PList, 2)
        If a / Round(UBound(PList, 2) / 20, 0) = Round(a / Round(UBound(PList, 2) / 20, 0), 0) Then If CBA_BasicFunctions.isRunningSheetDisplayed = True And UBound(PList, 2) > 0 Then CBA_BasicFunctions.RunningSheetAddComment 1, 1, "Distributing National USW DataCube: " & Round((a / UBound(PList, 2)), 2) * 100 & "%": DoEvents
        If pbSalesUSWBool And IsEmpty(psdUSWIndex(PList(0, a))) = False Then Set psdCG_PDic(CStr(PList(0, a))).cUSW = psdUSWIndex(PList(0, a))
    Next
    Set psdUSWIndex = Nothing
    
    SalesBool = buildSalesDataRS(YearStart, sBuildCode, "SalesDiv")
    For a = LBound(PList, 2) To UBound(PList, 2)
        If a / Round(UBound(PList, 2) / 20, 0) = Round(a / Round(UBound(PList, 2) / 20, 0), 0) Then If CBA_BasicFunctions.isRunningSheetDisplayed = True And UBound(PList, 2) > 0 Then CBA_BasicFunctions.RunningSheetAddComment 1, 1, "Distributing Regional Sales DataCube: " & Round((a / UBound(PList, 2)), 2) * 100 & "%": DoEvents
        If pbSalesDivBool And IsEmpty(psdSalesDivIndex(PList(0, a))) = False Then Set psdCG_PDic(CStr(PList(0, a))).cSalesDiv = psdSalesDivIndex(PList(0, a))
    Next
    Set psdSalesDivIndex = Nothing

    SalesBool = buildSalesDataRS(YearStart, sBuildCode, "USWDiv")
    For a = LBound(PList, 2) To UBound(PList, 2)
        If a / Round(UBound(PList, 2) / 20, 0) = Round(a / Round(UBound(PList, 2) / 20, 0), 0) Then If CBA_BasicFunctions.isRunningSheetDisplayed = True And UBound(PList, 2) > 0 Then CBA_BasicFunctions.RunningSheetAddComment 1, 1, "Distributing Regional USW DataCube: " & Round((a / UBound(PList, 2)), 2) * 100 & "%": DoEvents
        If pbSalesUSWDivBool And IsEmpty(psdUSWDivIndex(PList(0, a))) = False Then Set psdCG_PDic(CStr(PList(0, a))).cUSWDiv = psdUSWDivIndex(PList(0, a))
    Next
    Set psdUSWDivIndex = Nothing

    SalesBool = buildSalesDataRS(YearStart, sBuildCode, "ST")
    For a = LBound(PList, 2) To UBound(PList, 2)
        If a / Round(UBound(PList, 2) / 20, 0) = Round(a / Round(UBound(PList, 2) / 20, 0), 0) Then If CBA_BasicFunctions.isRunningSheetDisplayed = True And UBound(PList, 2) > 0 Then CBA_BasicFunctions.RunningSheetAddComment 1, 1, "Distributing Sell Through DataCube: " & Round((a / UBound(PList, 2)), 2) * 100 & "%": DoEvents
        If pbSalesSTBool And IsEmpty(psdSTIndex(PList(0, a))) = False Then Set psdCG_PDic(CStr(PList(0, a))).c14DayST = psdSTIndex(PList(0, a))
    Next
    Set psdSTIndex = Nothing

    End If
    
    
    
    
    
    'Set GenProductDic = ProductDic
    If sdCG_PDic Is Nothing Then Else GenProductDic = True
    'If InStr(1, sBuildCode, "Cm") > 0 Then Erase CBA_COM_Match
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_ProdGroup.GenProductDic", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Com", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Function GetCGListing(ByVal ACG As Boolean) As Variant
    If ACG = True Then GetCGListing = vACGList Else GetCGListing = vCGList
End Function
Private Function CreateCGListing() As Boolean
    Dim strSQL As String
    Dim RS As ADODB.Recordset
    On Error GoTo Err_Routine
    CBA_Error = ""
    If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Creating CGListing..."
    If RunConnectionSetups(True, False, False, False, False) = False Then Exit Function
    Set RS = New ADODB.Recordset
    strSQL = "select cg.cgno, cg.description as CGDesc, isnull(scgno,0), isnull(scg.description,cg.Description) as SCGDesc" & Chr(10)
    strSQL = strSQL & "from cbis599p.dbo.COMMODITYGROUP cg left join cbis599p.dbo.SUBCOMMODITYGROUP scg on scg.cgno = cg.cgno"
    RS.Open strSQL, cnCCN
    If RS.EOF Then CreateCGListing = False: Exit Function
    pvCGList = RS.GetRows()
    Set RS = New ADODB.Recordset
    strSQL = "select * from cbis599p.dbo.tf_ACGMap() order by CGno"
    RS.Open strSQL, cnCCN
    If RS.EOF Then CreateCGListing = False: Exit Function
    pvACGList = RS.GetRows()
    Set RS = Nothing
    CreateCGListing = True
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_ProdGroup.createCGListing", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Com", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Function pullData(ByVal QueryFor As String, ByVal lPcode As Long, ByVal BuildCode As String, Optional ByVal div As Long) As ADODB.Recordset
    Dim RS As ADODB.Recordset
    Dim strSQL As String
    Dim tim As Double
    tim = Timer
    Set RS = New ADODB.Recordset
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
    If InStr(1, BuildCode, "c") > 0 And div <> 0 Then
        If QueryFor = "Contract" Then strSQL = strSQL & "select * from #CONT where productcode = " & lPcode & " Order by deliveryfrom"
        If QueryFor = "Braket" Then strSQL = strSQL & "select * from #BV where productcode = " & lPcode & " Order by BKVfrom"
        If QueryFor = "Price" Then strSQL = strSQL & "select * from #RET where productcode = " & lPcode & " Order by RVfrom"
        If QueryFor = "IncoTerm" Then strSQL = strSQL & "select * from #DCH where productcode = " & lPcode & " Order by Validfrom"
        RS.Open strSQL, pcnRCN(div)
        If DebugProdGroupObjects = True Then Debug.Print "(PullData)Pull" & QueryFor & " Data: " & Timer - tim

    ElseIf InStr(1, BuildCode, "B") > 0 And QueryFor = "Base" Then
        strSQL = strSQL & "select *  from #BASEPRODDATA where productcode = " & lPcode & Chr(10)
        RS.Open strSQL, cnCCN
        If DebugProdGroupObjects = True Then Debug.Print "(PullData)PullBaseData: " & Timer - tim
    ElseIf (InStr(1, BuildCode, "R") > 0 Or InStr(1, BuildCode, "P") > 0) Then
        If QueryFor = "Sales" Then
            strSQL = strSQL & "select * into #SQRY from #SALES where productcode = " & lPcode & Chr(10)
            strSQL = strSQL & "select *, (select count(productcode) from #SQRY) as cnt from #SQRY drop table #SQRY" & Chr(10)
            RS.Open strSQL, cnCCN
            If DebugProdGroupObjects = True Then Debug.Print "(PullData)PullSalesData: " & Timer - tim
        ElseIf QueryFor = "USW" Then
            strSQL = strSQL & "select * from #SALESWK where productcode = " & lPcode & " order by yrno, wkno, productcode"
            RS.Open strSQL, cnCCN
            If DebugProdGroupObjects = True Then Debug.Print "(PullData)PullUSWData: " & Timer - tim
        ElseIf QueryFor = "SalesDiv" Then
            strSQL = strSQL & "select * from #SALESDIV where productcode = " & lPcode & " order by DivNo,yrno, mno"
            RS.Open strSQL, cnCCN
            If DebugProdGroupObjects = True Then Debug.Print "(PullData)PullSalesDivData: " & Timer - tim
        ElseIf QueryFor = "USWDiv" Then
            strSQL = strSQL & "select * from #SALESWKDIV where productcode = " & lPcode & " order by divno ,yrno, wkno, productcode"
            RS.Open strSQL, cnCCN
            If DebugProdGroupObjects = True Then Debug.Print "(PullData)PullUSWDivData: " & Timer - tim
        ElseIf QueryFor = "SellThrough" Then
            strSQL = "select * from  #SALESSP14 where productcode = " & lPcode & " order by productcode, OSD "
            RS.Open strSQL, cnCCN
            If DebugProdGroupObjects = True Then Debug.Print "(PullData)PullSTData: " & Timer - tim

        End If
    Else
        Set pullData = Nothing
        'Set pullData = RS
        Exit Function
    End If
    If RS.EOF Then Set pullData = Nothing Else Set pullData = RS
    'Set pullData = RS

End Function
Private Function prepareALLContIndex(ByVal vData As Variant, ByVal DataSet As e_DataCont) As Scripting.Dictionary
    Dim a As Long, b As Long
    Dim ProdCol As String
    Dim dic As Scripting.Dictionary
    Dim col As Collection
    Dim curProd As String
    Dim PrdDic As Scripting.Dictionary
    Dim P As Variant
    Dim tim As Double
    tim = Timer
    If vData(0, 0) = 0 Then Set prepareALLContIndex = Nothing: Exit Function
    Set dic = New Scripting.Dictionary
    If DataSet = e_DataCont.eContractData Then ProdCol = 2 Else ProdCol = 1
    Set PrdDic = New Scripting.Dictionary
    For a = LBound(vData, 2) To UBound(vData, 2)
        If PrdDic.Exists(CStr(vData(ProdCol, a))) = False Then PrdDic.Add CStr(vData(ProdCol, a)), 0
    Next
    For Each P In PrdDic
        Set col = New Collection
        For a = LBound(vData, 2) To UBound(vData, 2)
            If CStr(vData(ProdCol, a)) = P Then col.Add a
        Next
        dic.Add CStr(P), col
    Next
    Set prepareALLContIndex = dic
    If DebugProdGroupObjects = True Then Debug.Print "PullALLContData: " & Timer - tim
    
End Function
Private Function pullALLContData(ByVal DataSet As e_DataCont) As Variant
    Dim RS As ADODB.Recordset
    Dim arr(RDIVFROM To RDIVTO) As Variant, ContArr() As Variant
    Dim TotRows As Long, TotCols As Long, cnt As Long, rnd As Long
    Dim div As Long
    Dim a As Long, b As Long
    Dim strSQL As String
    Dim tim As Double
    tim = Timer
    On Error GoTo Err_Routine
    CBA_Error = ""
    TotRows = 0: TotCols = 0

        tim = Timer
        If DataSet = e_DataCont.eContractData Then strSQL = "Select * from #CONT order by productcode": If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Preparing Contract Data..."
        If DataSet = e_DataCont.eBraketData Then strSQL = "Select * from #BV order by productcode": If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Preparing Braket Data..."
        If DataSet = e_DataCont.ePricedata Then strSQL = "Select * from #RET order by productcode": If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Preparing Retail Data..."
        If DataSet = e_DataCont.eIncoData Then strSQL = "Select * from #DCH order by productcode": If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Preparing IncoTerm Data..."
        
        For div = RDIVFROM To RDIVTO
            Set RS = New ADODB.Recordset
            If sdRegDivMiss.Exists(div) = False Then
                RS.Open strSQL, pcnRCN(div)
                If RS.EOF Then
                Else
                    arr(div) = RS.GetRows()
                    TotRows = TotRows + (UBound(arr(div), 2) + 1)
                    If TotCols < UBound(arr(div), 1) Then TotCols = UBound(arr(div), 1)
                End If
                If DebugProdGroupObjects = True Then Debug.Print tim - Timer
            End If
        Next
        TotRows = TotRows - 1
        If TotRows = -1 Then
            ReDim ContArr(0, 0)
            ContArr(0, 0) = 0
            pullALLContData = ContArr
            Exit Function
        End If
        ReDim ContArr(0 To TotCols, 0 To TotRows)
        cnt = -1
        For div = RDIVFROM To RDIVTO
            If sdRegDivMiss.Exists(div) = False And IsEmpty(arr(div)) = False Then
                For a = LBound(arr(div), 2) To UBound(arr(div), 2)
                    If arr(div)(0, 0) = 0 Then
                    Else
                        cnt = cnt + 1
                        For b = LBound(arr(div), 1) To UBound(arr(div), 1)
                            ContArr(b, cnt) = arr(div)(b, a)
                        Next
                    End If
                Next
            End If
        Next
        pullALLContData = ContArr
Exit_Routine:
    If DebugProdGroupObjects = True Then Debug.Print "PullContData: " & Timer - tim
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_ProdGroup.pullALLContData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Com", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Function pullBaseData(ByVal lPcode As Long, ByVal BuildCode As String) As ADODB.Recordset
    Dim tim As Double
    tim = Timer
    Set pullBaseData = pullData("Base", lPcode, BuildCode)
    If DebugProdGroupObjects = True Then Debug.Print "Pull Base Data: " & Timer - tim
    
End Function
Private Function getAlcoholStoreNos() As Boolean
Dim tim As Double
Dim RS As ADODB.Recordset
Dim CN As ADODB.Connection
Dim yearDic As Scripting.Dictionary
Dim monthDic As Scripting.Dictionary
Dim regionDic As Scripting.Dictionary
Dim strSQ As String
Dim curMonth As Long, curYear As Long

    tim = Timer
    Set CN = New ADODB.Connection
    With CN
        .ConnectionTimeout = 50
        .CommandTimeout = 50
        .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & CBA_BSA & "LIVE DATABASES\ABI.accdb"
    End With
    Set RS = New ADODB.Recordset
    RS.Open "select * from ABI_AlcoholStores", CN
    
    Set yearDic = New Scripting.Dictionary
    
    Set regionDic = New Scripting.Dictionary
    
    
    Do Until RS.EOF
        If curYear <> Year(RS(0)) And curYear > 0 Then
            yearDic.Add CLng(curYear), monthDic
            Set monthDic = New Scripting.Dictionary
            curYear = Year(RS(0))
        ElseIf curYear = 0 Then
            curYear = Year(RS(0))
            Set monthDic = New Scripting.Dictionary
        End If
        Set regionDic = New Scripting.Dictionary
        regionDic.Add 501, CLng(RS(1))
        regionDic.Add 502, CLng(RS(2))
        regionDic.Add 503, CLng(RS(3))
        regionDic.Add 504, CLng(RS(4))
        regionDic.Add 505, CLng(RS(5))
        regionDic.Add 506, CLng(RS(6))
        regionDic.Add 507, CLng(RS(7))
        regionDic.Add 509, CLng(RS(8))
        regionDic.Add 599, CLng(RS(9))
        monthDic.Add Month(RS(0)), regionDic
        RS.MoveNext
    Loop
    yearDic.Add CLng(curYear), monthDic
    Set sdAlcoholStoreNos = yearDic
    CN.Close
    Set CN = Nothing
    If DebugProdGroupObjects = True Then Debug.Print "Pull Alcohol Store Data: " & Timer - tim
    
End Function
Private Function buildProdToMSegData(ByRef PList As Variant) As Boolean
Dim tim As Double
Dim strPlist As String
Dim a As Long, b As Long
Dim strSQL As String
Dim RS As ADODB.Recordset


    strPlist = ""
    For a = LBound(PList, 2) To UBound(PList, 2)
        If strPlist = "" Then
            strPlist = PList(0, a)
        Else
            If a / 20 = Round(a / 20, 0) Then strPlist = strPlist & Chr(10)
            strPlist = strPlist & ", " & PList(0, a)
        End If
    Next
    strSQL = "select * from L1_ProdToMSegAllocation where PA_Productcode in (" & strPlist & ")" & Chr(10) & "Order by PA_Productcode"
    Set RS = New ADODB.Recordset
    RS.Open strSQL, cnCRCN
    If RS.EOF Then
        ReDim ProdToMSegData(0 To 0, 0 To 0)
        ProdToMSegData(0, 0) = 0
        bHasProdToMSegData = False
    Else
        ProdToMSegData = RS.GetRows()
        bHasProdToMSegData = True
    End If
    buildProdToMSegData = bHasProdToMSegData
    If DebugProdGroupObjects = True Then Debug.Print "Pull ProdToMseg Data: " & Timer - tim
    
End Function
Private Function buildComradeData(ByRef PList As Variant) As Boolean
    Dim strPlist As String
    Dim a As Long
    Dim col As Collection
    Dim curPCode As Long
    Dim c As Long
    On Error GoTo Err_Routine
    CBA_Error = ""
    strPlist = ""
    For a = LBound(PList, 2) To UBound(PList, 2)
        If strPlist = "" Then strPlist = PList(0, a) Else strPlist = strPlist & ", " & PList(0, a)
    Next
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 0, 0, "Connecting to COMRADE..."
    If CBA_COM_SetupMatchArray.CBA_SetupMatchArray(False, dtMnthFrom, dtMnthTo, , , strPlist, True) = False Then
        buildComradeData = False
        'MsgBox "COMRADE Data not generated correctly", vbOKOnly
        Exit Function
    Else
        Set sdCOMdic = New Scripting.Dictionary
        For a = LBound(PList, 2) To UBound(PList, 2)
            Set col = New Collection
            For c = LBound(CBA_COM_Match) To UBound(CBA_COM_Match)
                If CBA_COM_Match(c).AldiPCode = PList(0, a) Then
                    col.Add CBA_COM_Match(c)
                End If
            Next
            'If col.Count > 0 Then Debug.Print PList(0, a) & "- No of Matches:" & col.Count
            If col.Count > 0 Then sdCOMdic.Add CStr(PList(0, a)), col
        Next
        buildComradeData = True
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_ProdGroup.buildComradeData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Com", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Function buildWastageData(ByVal StartYear As Long) As Boolean
    Dim strSQL As String
    Dim div As Long, a As Long, b As Long, rw As Long, TotRows As Long
    Dim arr(501 To 509) As Variant
    Dim tme As Variant
    Dim col As Collection
    Dim TempArr() As Variant
    Dim ProdDic As Scripting.Dictionary
    Dim WD As cCBA_WastageData
    Dim cnt As Long
    Dim totcnt As Long
    Dim curProd As Long
    Dim st As Long
    Dim CN(501 To 509) As cCBA_Connect
    Dim RefDic As Scripting.Dictionary
    Dim RS As ADODB.Recordset
    Dim v As Variant
    On Error GoTo Err_Routine
    
    CBA_Error = ""
    
    tme = Timer
    
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
    strSQL = strSQL & "declare @startdate as date = '" & StartYear & "-01-01'" & Chr(10)
    strSQL = strSQL & "declare @enddate as date = convert(date,getdate())" & Chr(10)
    strSQL = strSQL & "select productcode into #PS from purchase.dbo.product p " & Chr(10)
    If bACG = True Then strSQL = strSQL & "left join purchase.dbo.tf_ACGMap() as acg on acg.ACGEntityID = p.ACGEntityID" & Chr(10)
    strSQL = strSQL & "where " & buildstrCGList(True) & Chr(10)
    strSQL = strSQL & "select Mm.StoreNoFrom as StoreNo, Mm.MemoDate, MD.ProductCode" & Chr(10)
    strSQL = strSQL & ",MD.totalpieces as TotalPieces, MD.totalretail as TotalRetail" & Chr(10)
    strSQL = strSQL & "into #MS from Purchase.dbo.MemoDetail as MD WITH(INDEX(MEMOD_PRODMEMO))" & Chr(10)
    strSQL = strSQL & "Inner Join Purchase.dbo.MemoMaster as MM on MM.MemoNo = MD.MemoNo" & Chr(10)
    strSQL = strSQL & "inner join #PS p on p.ProductCode = MD.ProductCode" & Chr(10)
    strSQL = strSQL & "where MemoDate >= @startdate and memodate <= @enddate and MM.memotypeid in (8, 15, 16)" & Chr(10)
    strSQL = strSQL & "select StoreNo, MemoDate, ProductCode, sum(totalpieces) as TotalPieces, sum(totalretail) as TotalRetail" & Chr(10)
    strSQL = strSQL & "into #MMS from #MS group by StoreNo, Memodate, ProductCode" & Chr(10)
    strSQL = strSQL & "select i.storeno, i.InventoryDate, i.productcode, sum(i.invdiff) as qtydiff" & Chr(10)
    strSQL = strSQL & "into #INVS  from purchase.dbo.StoreInvDiff i" & Chr(10)
    strSQL = strSQL & "inner join #PS p on p.ProductCode = i.ProductCode" & Chr(10)
    strSQL = strSQL & "where i.InventoryDate >= @startdate" & Chr(10)
    strSQL = strSQL & "group by i.storeno,i.InventoryDate, i.productcode" & Chr(10)
    strSQL = strSQL & "select r.productcode, retail, validfrom, isnull(validto,convert(date,getdate())) as validto into #RETSS from purchase.dbo.Retail r" & Chr(10)
    strSQL = strSQL & "inner join #PS p on p.ProductCode = r.ProductCode" & Chr(10)
    strSQL = strSQL & "select i.*, i.qtydiff * r.Retail as RetailDiff into #IS" & Chr(10)
    strSQL = strSQL & "from #INVS i left join #RETSS r on r.ProductCode = i.ProductCode and r.ValidFrom <= i.InventoryDate and r.validto >= i.InventoryDate" & Chr(10)
    strSQL = strSQL & "select isnull(m.memodate, i.Inventorydate) as WDate, isnull(m.productcode, i.productcode) as productcode" & Chr(10)
    strSQL = strSQL & ", isnull(m.TotalPieces,0) as TotalPieces, isnull(m.TotalRetail,0) as TotalRetail" & Chr(10)
    strSQL = strSQL & ", isnull(i.qtydiff,0) as QTYDiff, isnull(i.RetailDiff,0) as RetailDiff, isnull(m.StoreNo, i.StoreNo) as  storeno" & Chr(10)
    strSQL = strSQL & "into #W from #MMS m full outer join #IS i on i.InventoryDate = m.MemoDate and i.ProductCode = m.ProductCode and i.StoreNo = m.StoreNo" & Chr(10)
    strSQL = strSQL & "select isnull(m.productcode, i.productcode) as productcode into #PRODS from #MMS m full outer join #IS i on i.InventoryDate = m.MemoDate and i.ProductCode = m.ProductCode and i.StoreNo = m.StoreNo group by isnull(m.productcode, i.productcode) order by isnull(m.productcode, i.productcode)" & Chr(10)
    strSQL = strSQL & "drop table #PS, #MS,  #INVS, #RETSS, #IS, #MMS" & Chr(10)
    Set psdWastageIndex = New Scripting.Dictionary
    Set ProdDic = New Scripting.Dictionary
    For div = RDIVFROM To RDIVTO
        If div = 508 Then div = 509
        If sdRegDivMiss.Exists(div) = False Then
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Generating " & CBA_BasicFunctions.CBA_DivtoReg(div) & " Wastage  Data..."
            Set CN(div) = New cCBA_Connect
            CN(div).SetConnection CStr(div)
            CN(div).Query strSQL, "MainWasteQuery"
            CN(div).Query "select productcode from #PRODS", "ProdsQuery"
            Set RS = CN(div).op("ProdsQuery")
            Do Until RS.EOF
                If psdWastageIndex.Exists(CLng(RS(0))) = False Then
                    Set WD = New cCBA_WastageData
                    WD.lPcode = CLng(RS(0))
                    psdWastageIndex.Add CLng(RS(0)), WD
                End If
                RS.MoveNext
            Loop
            CN(div).Query "select count(distinct productcode) from #W ", "ProdCount"
            CN(div).Query "select * from #W order by productcode, Wdate", "FullQuery"
        End If
    Next
    For div = RDIVFROM To RDIVTO
        If div = 508 Then div = 509
           If CBA_BasicFunctions.isRunningSheetDisplayed = True And totcnt > 0 Then CBA_BasicFunctions.RunningSheetAddComment 1, 1, "Packaging " & CBA_BasicFunctions.CBA_DivtoReg(div) & " Wastage DataCube: ...preparing..."
        'cnt = 0
        Set RS = CN(div).op("FullQuery")
        'If CN(div).Query("select count(productcode) from #W", "ProdCount") = True Then totcnt = CN(div).op("ProdCount")(0)
        Set RefDic = New Scripting.Dictionary
        Set RS = CN(div).op("FullQuery")
        a = 0: curProd = CLng(RS(1)): st = 0
        Do Until RS.EOF
            If curProd <> CLng(RS(1)) Then
                RefDic.Add curProd, a
                curProd = CLng(RS(1))
                a = 0
            End If
            a = a + 1
            RS.MoveNext
        Loop
        If curProd > 0 Then RefDic.Add curProd, a
        
        
        RS.MoveFirst
'        CurProd = 0
'        a = -1
        totcnt = CN(div).op("ProdCount")(0)
        cnt = 0
        Do Until RS.EOF
'            a = a + 1
'            If CurProd <> CLng(RS(1)) Then
'                RefDic.Add CLng(RS(1)), a
'                CurProd = CLng(RS(1))
'            End If
'        RS.MoveFirst
'        For Each v In psdWastageIndex
            cnt = cnt + 1
            If cnt / Round(totcnt / 20, 0) = Round(cnt / Round(totcnt / 20, 0), 0) Then If CBA_BasicFunctions.isRunningSheetDisplayed = True And totcnt > 0 Then CBA_BasicFunctions.RunningSheetAddComment 1, 1, "Packaging " & CBA_BasicFunctions.CBA_DivtoReg(div) & " Wastage DataCube: " & Round((cnt / totcnt), 2) * 100 & "%": DoEvents
            psdWastageIndex(CLng(RS(1))).addData RS, div, RefDic(CLng(RS(1)))
'        Next
            'RS.MoveNext
        Loop
    
    Next
    buildWastageData = True
If DebugProdGroupObjects = True Then Debug.Print Timer - tme

Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_ProdGroup.buildComradeData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Com", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Function buildContractData(ByVal StartYear As Long) As Boolean
Dim strSQL As String
Dim div As Long, a As Long
    On Error GoTo Err_Routine
    CBA_Error = ""

    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
    strSQL = strSQL & "select distinct p.productcode, TaxID into #PROD" & Chr(10)
    strSQL = strSQL & "from purchase.dbo.product p" & Chr(10)
    strSQL = strSQL & "left join purchase.dbo.tf_ACGMap() acg on acg.ACGEntityID = p.ACGEntityID" & Chr(10)
    strSQL = strSQL & "where" & buildstrCGList(True) & Chr(10)
    strSQL = strSQL & "select distinct substring(@@SERVERNAME,2,3) as ServerName ,c.contractno, c.productcode, c.DeliveryFrom, c.deliveryto, c.CaseHeight, c.CaseLength, c.CaseWidth" & Chr(10)
    strSQL = strSQL & ", c.CountryOfOrigin, c.PayTermDays, c.AdvertisingDate, c.returnsallowed, c.leadtimedelivery, c.ValidInDiv, c.RCVTempFrom, c.RCVTempTo" & Chr(10)
    strSQL = strSQL & ", c.TemperatureType, c.ExpirationTime, c.supplierno, s.Name1, p.TaxID" & Chr(10)
    strSQL = strSQL & "into #CONT" & Chr(10)
    strSQL = strSQL & "from purchase.dbo.Contract c" & Chr(10)
    strSQL = strSQL & "inner join #PROD p on p.productcode = c.productcode left join purchase.dbo.supplier s on s.supplierno = c.supplierno" & Chr(10)
    strSQL = strSQL & "where isnull(c.DeliveryTo, getdate()) > '" & StartYear & "-01-01'" & Chr(10)
    strSQL = strSQL & "order by c.ProductCode, c.deliveryfrom" & Chr(10)
    strSQL = strSQL & "" & Chr(10)
    strSQL = strSQL & "select substring(@@SERVERNAME,2,3) as ServerName , p.productcode , bv.contractno, bv.BracketNo, bv.ValidFrom as BKVFrom, bv.ValidTo as BKVTo, bv.ExWorks, bv.FOB, bv.DAT, bv.CurrencyCode, bv.Freehouse, bv.Freight" & Chr(10)
    strSQL = strSQL & "into #BV" & Chr(10)
    strSQL = strSQL & "from purchase.dbo.BracketValue bv" & Chr(10)
    strSQL = strSQL & "inner join (select distinct Contractno from #CONT) c on c.ContractNo = bv.ContractNo" & Chr(10)
    strSQL = strSQL & "left join #CONT ct on ct.ContractNo = bv.ContractNo" & Chr(10)
    strSQL = strSQL & "left join #PROD p on p.ProductCode = ct.ProductCode" & Chr(10)
    strSQL = strSQL & "where isnull(bv.ValidTo, getdate()) > '" & StartYear & "-01-01'" & Chr(10)
    strSQL = strSQL & "order by c.contractno, bv.validfrom" & Chr(10)
    strSQL = strSQL & "" & Chr(10)
    strSQL = strSQL & "select substring(@@SERVERNAME,2,3) as ServerName ,r.productcode, r.retail, r.ValidFrom as RVFrom, r.ValidTo as RVTo" & Chr(10)
    strSQL = strSQL & "into #RET" & Chr(10)
    strSQL = strSQL & "from purchase.dbo.retail r" & Chr(10)
    strSQL = strSQL & "inner join #PROD p on p.productcode = r.productcode" & Chr(10)
    strSQL = strSQL & "where isnull(r.ValidTo, getdate()) > '" & StartYear & "-01-01'" & Chr(10)
    strSQL = strSQL & "order by r.ProductCode, r.ValidFrom" & Chr(10)
    strSQL = strSQL & "select substring(@@SERVERNAME,2,3) as ServerName ,c.ProductCode, ch.contractno, ch.divno, ch.validfrom, ch.validto, ch.sendtodiv" & Chr(10)
    strSQL = strSQL & ", case when pricebase = 0 then 'DDP' when PriceBase = 2 then 'ExWorks' when PriceBase = 3 then 'FOB' end as incoterm" & Chr(10)
    strSQL = strSQL & "into #DCH from purchase.dbo.ContractHis ch" & Chr(10)
    strSQL = strSQL & "inner join #CONT c on c.contractno = ch.contractno" & Chr(10)
    strSQL = strSQL & "where isnull(ch.ValidTo, getdate()) > '" & StartYear & "-01-01'" & Chr(10)
    'Debug.Print strSQL
    For div = RDIVFROM To RDIVTO
        If sdRegDivMiss.Exists(div) = False Then
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Generating " & CBA_BasicFunctions.CBA_DivtoReg(div) & " Contract  Data..."
            Set prsContRS(div) = New ADODB.Recordset
            prsContRS(div).Open strSQL, pcnRCN(div)
            buildContractData = True
            DoEvents
        End If
    Next
        
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_ProdGroup.buildComradeData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Com", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function buildBaseData() As Boolean
Dim strSQL As String
Dim div As Long, a As Long
    On Error GoTo Err_Routine
    CBA_Error = ""

    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Generating Base Data..."
    Set rsBaseProdRS = New ADODB.Recordset
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
    strSQL = strSQL & "select p.productcode, p.description, p.ProductClass,p.SeasonID, p.Packsize,p.Con_ProductCode, p.Contents, p.EmpNo, p.isRandomWeight,p.PriceCrdBrand, p.PriceCrdDesc" & Chr(10)
    strSQL = strSQL & ", p.CGno as LCGno, cg.Description as LCGDesc" & Chr(10)
    strSQL = strSQL & ", isnull(p.scgno,0) as LSCG, isnull(scg.Description,cg.Description) as LSCGDESC" & Chr(10)
    strSQL = strSQL & ", convert(nvarchar(4),convert(int,p.cgno)) + '-' + convert(nvarchar(4),convert(int,isnull(p.scgno,0))) as LCGmap" & Chr(10)
    strSQL = strSQL & ", acg.CatNo as ACatNo, acg.Category, acg.CGNo as ACGno, acg.CommodityGroup as ACGDesc" & Chr(10)
    strSQL = strSQL & ", acg.SCGNo as ASCGNo, acg.SubCommodityGroup as ASCGDesc" & Chr(10)
    strSQL = strSQL & ", convert(nvarchar(4),acg.CatNo) + '-' + convert(nvarchar(4),acg.CGno) + '-' + convert(nvarchar(4),acg.SCGno) as ACGmap, p.taxid, p.Trademarkno" & Chr(10)
    strSQL = strSQL & "into #BASEPRODDATA  from cbis599p.dbo.product p" & Chr(10)
    strSQL = strSQL & "left join cbis599p.dbo.COMMODITYGROUP cg on cg.cgno = p.cgno" & Chr(10)
    strSQL = strSQL & "left join cbis599p.dbo.SUBCOMMODITYGROUP scg on scg.cgno = p.cgno and scg.scgno = p.scgno" & Chr(10)
    strSQL = strSQL & "left join cbis599p.dbo.tf_ACGMap() acg on acg.ACGEntityID = p.ACGEntityID" & Chr(10)
    strSQL = strSQL & "where" & buildstrCGList(True) & Chr(10)
    strSQL = strSQL & "order by p.productcode"
    rsBaseProdRS.Open strSQL, cnCCN
    buildBaseData = True

Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_ProdGroup.buildComradeData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Com", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Function buildProductLabelData() As Boolean
Dim strSQL As String, str As String
Dim div As Long, a As Long
Dim RS As ADODB.Recordset
Dim curProd As String
    On Error GoTo Err_Routine
    CBA_Error = ""


    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Generating Product Label Data..."
    Set RS = New ADODB.Recordset
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
    strSQL = strSQL & "select pa.productcode, pat.Description, rank() over (Partition by pa.Productcode order by pat.Description)" & Chr(10)
    strSQL = strSQL & "from cbis599p.dbo.ProductAttribute pa inner join #BASEPRODDATA b on b.productcode = pa.ProductCode left join cbis599p.dbo.ProductAttributeType pat on pat.ProductAttributeTypeID = pa.ProductAttributeTypeID" & Chr(10)
    strSQL = strSQL & "order by pa.ProductCode , pat.Description desc" & Chr(10)
    RS.Open strSQL, cnCCN
    a = 0
    If RS.EOF = False Then curProd = RS(0): Set sdProductLabelDic = New Scripting.Dictionary: str = "": buildProductLabelData = True
    Do Until RS.EOF
        If curProd <> RS(0) Then
            sdProductLabelDic.Add curProd, str
            curProd = RS(0)
            str = ""
        End If
        If str = "" Then str = RS(1) Else str = str & ", " & RS(1)
        RS.MoveNext
    Loop
    If Not sdProductLabelDic Is Nothing Or curProd <> "" Then sdProductLabelDic.Add curProd, str
        
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_ProdGroup.buildComradeData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Com", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Function buildstrCGList(Optional ByVal FrontDescRqd As Boolean = False) As String
    Dim cnt As Long, a As Long
    Dim scg As Variant
    Dim strCGList As String, strCGL As String
    On Error GoTo Err_Routine
    CBA_Error = ""
    strCGList = ""
    cnt = 0
    If bRunFordoubleProd = True Then
        If FrontDescRqd = True Then
            buildstrCGList = " p.productcode = " & lPcode
        Else
            buildstrCGList = " productcode = " & lPcode
        End If
    Else
        If Not colCGs Is Nothing Then
            For Each scg In colCGs
                If strCGList = "" Then
                    If bACG = False Then
                        If FrontDescRqd = False Then
                            strCGL = "(Cgno = " & Left(scg, 3): If Right(scg, 2) > 0 Then strCGL = strCGL & " and SCGno = " & Right(scg, 2) & ")" Else strCGL = strCGL & ")"
                        Else
                            strCGL = "(p.Cgno = " & Left(scg, 3): If Right(scg, 2) > 0 Then strCGL = strCGL & " and p.SCGno = " & Right(scg, 2) & ")" Else strCGL = strCGL & ")"
                        End If
                    Else
                        strCGL = "(acg.Cgno = " & Left(scg, 3): If Right(scg, 2) > 0 Then strCGL = strCGL & " and acg.SCGno = " & Right(scg, 2) & ")" Else strCGL = strCGL & ")"
                    End If
                Else
                    If bACG = False Then
                        If FrontDescRqd = False Then
                            strCGL = " or (Cgno = " & Left(scg, 3): If Right(scg, 2) > 0 Then strCGL = strCGL & " and SCGno = " & Right(scg, 2) & ")" Else strCGL = strCGL & ")"
                        Else
                            strCGL = " or (p.Cgno = " & Left(scg, 3): If Right(scg, 2) > 0 Then strCGL = strCGL & " and p.SCGno = " & Right(scg, 2) & ")" Else strCGL = strCGL & ")"
                        End If
                    Else
                        strCGL = " or (acg.Cgno = " & Left(scg, 3): If Right(scg, 2) > 0 Then strCGL = strCGL & " and acg.SCGno = " & Right(scg, 2) & ")" Else strCGL = strCGL & ")"
                    End If
                End If
                cnt = cnt + Len(strCGL)
                If cnt > 99 Then
                    strCGList = strCGList & strCGL & Chr(10)
                    cnt = 0
                Else
                    strCGList = strCGList & strCGL
                End If
            Next
        End If
        buildstrCGList = strCGList
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f_CBA_ProdGroup.buildstrCGList", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Com", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function buildSalesDataRS(ByVal StartYear As Long, ByVal BuildCode As String, Optional ByVal SalesDatatoBuild As String) As Boolean
    Dim strSQL As String
    Dim strCGList As String
    Dim v As Variant
    Dim a As Long, curProd As Long
    Dim RS As ADODB.Recordset
    Dim PD As cCBA_ProdPeriodSalesData
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Generating Sales Data..."

    If bSalesBool = False And bSalesDivBool = False And bSalesUSWBool = False And bSalesUSWDivBool = False And bSalesSTBool = False Then
        strCGList = buildstrCGList
        Set rsSalesRS = New ADODB.Recordset
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
        strSQL = strSQL & "declare @YRST nvarchar(4) = " & StartYear & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        If bACG = False Then
            strSQL = strSQL & "select productcode into #TEMP  from cbis599p.dbo.product where " & Chr(10)
        Else
            strSQL = strSQL & "select p.productcode into #TEMP from cbis599p.dbo.product p" & Chr(10)
            strSQL = strSQL & "inner join cbis599p.dbo.tf_ACGMap() acg on acg.ACGEntityID = p.ACGEntityID where " & Chr(10)
        End If
        strSQL = strSQL & strCGList & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        
        If strCGList = "" Then
            buildSalesDataRS = False
            Exit Function
        End If
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select year(PDATE) as Yrno, Month(PDATE) as Mno into #DATE  from (" & Chr(10)
        strSQL = strSQL & "SELECT  TOP (DATEDIFF(MONTH, @YRST + '-01-01', EOMONTH(getdate(),-1)) + 1)" & Chr(10)
        strSQL = strSQL & "    Pdate = convert(date,DATEADD(MONTH, ROW_NUMBER() OVER(ORDER BY a.object_id) - 1, @YRST + '-01-01'))" & Chr(10)
        strSQL = strSQL & "    FROM    sys.all_objects a CROSS JOIN sys.all_objects b" & Chr(10)
        strSQL = strSQL & ") a" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select YEAR(DATEADD(day, 26 - DATEPART(isoww, pdate), pdate)) as yrno," & Chr(10)
        strSQL = strSQL & "DatePart(ISOWK, PDATE) as Wkno , count(pdate) as daycnt into #DATEWK  " & Chr(10)
        strSQL = strSQL & "from (SELECT  TOP (DATEDIFF(D, @YRST + '-01-01', EOMONTH(getdate(),-1)) + 1)" & Chr(10)
        strSQL = strSQL & "    Pdate = convert(date,DATEADD(d, ROW_NUMBER() OVER(ORDER BY a.object_id) - 1, @YRST + '-01-01'))" & Chr(10)
        strSQL = strSQL & "    FROM    sys.all_objects a CROSS JOIN sys.all_objects b ) a" & Chr(10)
        strSQL = strSQL & "group by YEAR(DATEADD(day, 26 - DATEPART(isoww, pdate), pdate)), DatePart(ISOWK, PDATE)" & Chr(10)
        strSQL = strSQL & "order by yrno, wkno" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select productcode into #PList from (select distinct rcv.productcode from cbis599p.dbo.receiving rcv inner join #TEMP p on p.ProductCode = rcv.ProductCode where year(DayEndDate) >= @YRST" & Chr(10)
        strSQL = strSQL & "union select distinct pos.productcode from cbis599p.dbo.pos pos inner join #TEMP p on p.ProductCode = pos.ProductCode where year(posdate) >= @YRST) a" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select p.productcode, p.CGNo as LCG, p.scgno as LSCG, acg.Category, acg.CatNo, ACG.CGNo, ACG.SCGNo, ProductClass, p.TaxID, P.EmpNo,P.Con_ProductCode" & Chr(10)
        strSQL = strSQL & ",P.isRandomWeight, P.Packsize, P.PriceCrdBrand, P.SeasonId,P.Trademarkno" & Chr(10)
        strSQL = strSQL & "into #P from cbis599p.dbo.product p" & Chr(10)
        strSQL = strSQL & "inner join #PList PL on pl.ProductCode = p.productcode" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.tf_ACGMap() acg on acg.ACGEntityID = p.ACGEntityID" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select productcode into #SP from #P where ProductClass in (2,3)" & Chr(10)
        strSQL = strSQL & "select p.productcode, c.contractno into #SC from cbis599p.dbo.contract c inner join #SP p on p.ProductCode = c.ProductCode group by p.productcode, c.contractno" & Chr(10)
        strSQL = strSQL & "select c.productcode, c.advertisingdate as OSD, dateadd(DAY,14,c.advertisingdate) as EOD" & Chr(10)
        strSQL = strSQL & "into #SPDTE from cbis599p.dbo.CONTRACT c" & Chr(10)
        strSQL = strSQL & "inner join #SC sc on sc.ProductCode = c.ProductCode and sc.contractno = c.contractno" & Chr(10)
        strSQL = strSQL & "where year(c.advertisingdate) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by c.productcode, c.advertisingdate , dateadd(DAY,14,c.advertisingdate)" & Chr(10)
        strSQL = strSQL & "select dayenddate, p.productcode, Quantity into #SRC from cbis599p.dbo.RECEIVING rcv inner join #SP p on p.ProductCode = rcv.ProductCode where recordid = '001' and year(dayenddate) >= @YRST" & Chr(10)
        strSQL = strSQL & "select s.productcode, s.OSD, sum(quantity) as QTY into #SPRCV from #SRC rcv" & Chr(10)
        strSQL = strSQL & "inner join #SPDTE s on s.ProductCode = rcv.ProductCode and rcv.DayEndDate >= dateadd(DAY,-45,s.EOD) and rcv.DayEndDate <= s.EOD group by s.productcode, s.OSD" & Chr(10)
        strSQL = strSQL & "select s.productcode, s.OSD, s.EOD, isnull(sum(retail),0) as retail, isnull(sum(Quantity),0) as QTY, isnull(rcv.QTY,0) as RCVQTY, isnull(sum(Quantity) / nullif(rcv.QTY,0),0) as ST  into #SALESSP14 from #SPDTE s" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.pos pos on pos.ProductCode = s.ProductCode and pos.posdate >= s.OSD and pos.posdate <= s.EOD" & Chr(10)
        strSQL = strSQL & "left join #SPRCV rcv on rcv.ProductCode = s.ProductCode and rcv.OSD = s.OSD" & Chr(10)
        strSQL = strSQL & "where s.OSD is not null group by s.productcode, s.OSD, s.EOD, rcv.QTY" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select year(posdate) as YRno, datepart(ISOWK,posdate) as WKno, pos.productcode, sum(retail) as retail, sum(quantity) as QTY" & Chr(10)
        strSQL = strSQL & "into #POSWK from cbis599p.dbo.pos pos" & Chr(10)
        strSQL = strSQL & "inner join #Plist P on p.ProductCode = pos.ProductCode" & Chr(10)
        strSQL = strSQL & "where YEAR(posdate) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by year(posdate) , datepart(ISOWK,posdate) , pos.productcode" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select year(posdate) as YRno, month(posdate) as Mno, pos.productcode, sum(retail) as retail, sum(quantity) as QTY" & Chr(10)
        strSQL = strSQL & "into #POS from cbis599p.dbo.pos pos" & Chr(10)
        strSQL = strSQL & "inner join #Plist P on p.ProductCode = pos.ProductCode" & Chr(10)
        strSQL = strSQL & "where YEAR(posdate) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by year(posdate) , month(posdate) , pos.productcode" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select year(dayenddate) as YRno, datepart(ISOWK,dayenddate) as Wkno, rcv.productcode , sum(Quantity) as Quantity , sum(Pieces) as Pieces" & Chr(10)
        strSQL = strSQL & "into #RCVQTYWK from cbis599p.dbo.receiving rcv" & Chr(10)
        strSQL = strSQL & "inner join #PList p on p.ProductCode = rcv.ProductCode" & Chr(10)
        strSQL = strSQL & "where rcv.recordid = '001' and YEAR(dayenddate) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by  year(dayenddate), datepart(ISOWK,dayenddate), rcv.productcode, rcv.recordid" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select YEAR(dayenddate) as YRno, datepart(ISOWK,dayenddate) as Wkno, rcv.productcode , sum(retail) as Retail, sum(retailnet) as RetailNet, sum(cost) as Cost" & Chr(10)
        strSQL = strSQL & "into #RCVRETWK from cbis599p.dbo.RECEIVING rcv" & Chr(10)
        strSQL = strSQL & "inner join #PList p on p.ProductCode = rcv.ProductCode" & Chr(10)
        strSQL = strSQL & "where YEAR(dayenddate) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by  year(dayenddate),  datepart(ISOWK,dayenddate), rcv.productcode" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select year(dayenddate) as YRno, month(dayenddate) as Mno, rcv.productcode , sum(Quantity) as Quantity , sum(Pieces) as Pieces" & Chr(10)
        strSQL = strSQL & "into #RCVQTY from cbis599p.dbo.receiving rcv" & Chr(10)
        strSQL = strSQL & "inner join #PList p on p.ProductCode = rcv.ProductCode" & Chr(10)
        strSQL = strSQL & "where rcv.recordid = '001' and YEAR(dayenddate) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by  year(dayenddate),month(dayenddate), rcv.productcode, rcv.recordid" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select YEAR(dayenddate) as YRno, Month(dayenddate) as Mno, rcv.productcode , sum(retail) as Retail, sum(retailnet) as RetailNet, sum(cost) as Cost" & Chr(10)
        strSQL = strSQL & "into #RCVRET from cbis599p.dbo.RECEIVING rcv" & Chr(10)
        strSQL = strSQL & "inner join #PList p on p.ProductCode = rcv.ProductCode" & Chr(10)
        strSQL = strSQL & "where YEAR(dayenddate) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by  year(dayenddate), month(dayenddate), rcv.productcode" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "set datefirst 1" & Chr(10)
        strSQL = strSQL & "select datepart(ISOWK,validfrom) as WeekNo, datepart(YEAR,validfrom) as YearNo" & Chr(10)
        strSQL = strSQL & ", dateadd(D,-2,validfrom) as WKStart, dateadd(D,4,validfrom) as WKEnd, sum(NoOfStores)  as NoOfStores" & Chr(10)
        strSQL = strSQL & "into #STORESWK from cbis599p.portfolio.Stores where datepart(YEAR,validfrom) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by datepart(ISOWK,validfrom) , datepart(YEAR,validfrom), dateadd(D,-2,validfrom), dateadd(D,4,validfrom)" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select d.YRno , d.Mno, p.productcode, isnull(R.Retail,0) as RCVRetail, isnull(RetailNet,0) as RCVRetailNet, isnull(Cost,0) as RCVCost" & Chr(10)
        strSQL = strSQL & ", isnull(RC.Quantity, 0) as RCVQTY, isnull(RC.Pieces, 0) as RCVPieces, isnull(pos.QTY,0) as POSQTY, isnull(pos.Retail,0) as POSRetail" & Chr(10)
        strSQL = strSQL & "into #SALES from #DATE d cross join #PList p" & Chr(10)
        strSQL = strSQL & "left join #RCVRET R on R.productcode = p.productcode and R.YRno = d.Yrno and R.Mno = d.Mno" & Chr(10)
        strSQL = strSQL & "left join #RCVQTY RC on RC.ProductCode = p.ProductCode and RC.Mno = d.Mno and RC.YRno = d.YRno" & Chr(10)
        strSQL = strSQL & "left join #POS pos on pos.ProductCode = p.ProductCode and pos.Mno = d.Mno and pos.YRno = d.Yrno" & Chr(10)
        strSQL = strSQL & "order by productcode, YRno, MNo" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select d.YRno , d.Wkno, p.productcode" & Chr(10)
        strSQL = strSQL & ", isnull(R.Retail / nullif(s.noofstores,0),0) as RCVRetailpwk, isnull(RetailNet/ nullif(s.noofstores,0),0) as RCVRetailNetpwk, isnull(Cost/ nullif(s.noofstores,0),0) as RCVCostpwk" & Chr(10)
        strSQL = strSQL & ", isnull(RC.Quantity/ nullif(s.noofstores,0), 0) as RCVQTYpwk, isnull(RC.Pieces/ nullif(s.noofstores,0), 0) as RCVPiecespwk, isnull(pos.QTY/ nullif(s.noofstores,0),0) as POSQTYpwk" & Chr(10)
        strSQL = strSQL & ", isnull(pos.Retail/ nullif(s.noofstores,0),0) as POSRetailpwk, isnull(d.daycnt,0) as DayCnt, isnull(s.noofstores,0) as NoOfStores, isnull(pr.CGNo,0) as CGno" & Chr(10)
        strSQL = strSQL & "into #SALESWK " & Chr(10)
        strSQL = strSQL & "from #DATEWK d cross join #PList p" & Chr(10)
        strSQL = strSQL & "left join #RCVRETWK R on R.productcode = p.productcode and R.YRno = d.Yrno and R.wkno = d.wkno" & Chr(10)
        strSQL = strSQL & "left join #RCVQTYWK RC on RC.ProductCode = p.ProductCode and RC.wkno = d.wkno and RC.YRno = d.YRno" & Chr(10)
        strSQL = strSQL & "left join #POSWK pos on pos.ProductCode = p.ProductCode and pos.wkno = d.wkno and pos.YRno = d.Yrno" & Chr(10)
        strSQL = strSQL & "left join #STORESWK s on s.WeekNo = d.wkno and s.YearNo = d.Yrno" & Chr(10)
        strSQL = strSQL & "left join #P pr on pr.productcode = p.ProductCode" & Chr(10)
        strSQL = strSQL & "order by productcode, YRno, wkno" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select Divno into #DIV from cbis599p.dbo.DIVISION where divno > '500' and divno < '510'" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        If InStr(1, sBuildCode, "Ur") > 0 Then
        strSQL = strSQL & "select pos.DivNo, year(posdate) as YRno, datepart(ISOWK,posdate) as WKno, pos.productcode, sum(retail) as retail, sum(quantity) as QTY" & Chr(10)
        strSQL = strSQL & "into #POSWKDIV from cbis599p.dbo.pos pos" & Chr(10)
        strSQL = strSQL & "inner join #Plist P on p.ProductCode = pos.ProductCode" & Chr(10)
        strSQL = strSQL & "where YEAR(posdate) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by pos.DivNo,year(posdate) , datepart(ISOWK,posdate) , pos.productcode" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        End If
        strSQL = strSQL & "select pos.DivNo, year(posdate) as YRno, month(posdate) as Mno, pos.productcode, sum(retail) as retail, sum(quantity) as QTY" & Chr(10)
        strSQL = strSQL & "into #POSDIV from cbis599p.dbo.pos pos" & Chr(10)
        strSQL = strSQL & "inner join #Plist P on p.ProductCode = pos.ProductCode" & Chr(10)
        strSQL = strSQL & "where YEAR(posdate) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by pos.DivNo, year(posdate) , month(posdate) , pos.productcode" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        If InStr(1, sBuildCode, "Ur") > 0 Then
        strSQL = strSQL & "select rcv.DivNo, year(dayenddate) as YRno, datepart(ISOWK,dayenddate) as Wkno, rcv.productcode , sum(Quantity) as Quantity , sum(Pieces) as Pieces" & Chr(10)
        strSQL = strSQL & "into #RCVQTYWKDIV from cbis599p.dbo.receiving rcv" & Chr(10)
        strSQL = strSQL & "inner join #PList p on p.ProductCode = rcv.ProductCode" & Chr(10)
        strSQL = strSQL & "where rcv.recordid = '001' and YEAR(dayenddate) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by rcv.DivNo, year(dayenddate), datepart(ISOWK,dayenddate), rcv.productcode, rcv.recordid" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select rcv.DivNo, YEAR(dayenddate) as YRno, datepart(ISOWK,dayenddate) as Wkno, rcv.productcode , sum(retail) as Retail, sum(retailnet) as RetailNet, sum(cost) as Cost" & Chr(10)
        strSQL = strSQL & "into #RCVRETWKDIV from cbis599p.dbo.RECEIVING rcv" & Chr(10)
        strSQL = strSQL & "inner join #PList p on p.ProductCode = rcv.ProductCode" & Chr(10)
        strSQL = strSQL & "where YEAR(dayenddate) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by rcv.DivNo, year(dayenddate),  datepart(ISOWK,dayenddate), rcv.productcode" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        End If
        strSQL = strSQL & "select rcv.DivNo, year(dayenddate) as YRno, month(dayenddate) as Mno, rcv.productcode , sum(Quantity) as Quantity , sum(Pieces) as Pieces" & Chr(10)
        strSQL = strSQL & "into #RCVQTYDIV from cbis599p.dbo.receiving rcv" & Chr(10)
        strSQL = strSQL & "inner join #PList p on p.ProductCode = rcv.ProductCode" & Chr(10)
        strSQL = strSQL & "where rcv.recordid = '001' and YEAR(dayenddate) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by rcv.DivNo,year(dayenddate),month(dayenddate), rcv.productcode, rcv.recordid" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        strSQL = strSQL & "select rcv.DivNo, YEAR(dayenddate) as YRno, Month(dayenddate) as Mno, rcv.productcode , sum(retail) as Retail, sum(retailnet) as RetailNet, sum(cost) as Cost" & Chr(10)
        strSQL = strSQL & "into #RCVRETDIV from cbis599p.dbo.RECEIVING rcv" & Chr(10)
        strSQL = strSQL & "inner join #PList p on p.ProductCode = rcv.ProductCode" & Chr(10)
        strSQL = strSQL & "where YEAR(dayenddate) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by rcv.DivNo, year(dayenddate), month(dayenddate), rcv.productcode" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        If InStr(1, sBuildCode, "Ur") > 0 Then
        strSQL = strSQL & "set datefirst 1" & Chr(10)
        strSQL = strSQL & "select Stores.DivNo, datepart(ISOWK,validfrom) as WeekNo, datepart(YEAR,validfrom) as YearNo" & Chr(10)
        strSQL = strSQL & ", dateadd(D,-2,validfrom) as WKStart, dateadd(D,4,validfrom) as WKEnd, sum(NoOfStores)  as NoOfStores" & Chr(10)
        strSQL = strSQL & "into #STORESWKDIV from cbis599p.portfolio.Stores where datepart(YEAR,validfrom) >= @YRST" & Chr(10)
        strSQL = strSQL & "group by Stores.DivNo, datepart(ISOWK,validfrom) , datepart(YEAR,validfrom), dateadd(D,-2,validfrom), dateadd(D,4,validfrom)" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        End If
        strSQL = strSQL & "select d.YRno , d.Mno, p.productcode, isnull(R.Retail,0) as RCVRetail, isnull(RetailNet,0) as RCVRetailNet, isnull(Cost,0) as RCVCost" & Chr(10)
        strSQL = strSQL & ", isnull(RC.Quantity, 0) as RCVQTY, isnull(RC.Pieces, 0) as RCVPieces, isnull(pos.QTY,0) as POSQTY, isnull(pos.Retail,0) as POSRetail , isnull(div.divno,0) as DivNo" & Chr(10)
        strSQL = strSQL & "into #SALESDIV from #DATE d cross join #PList p cross join #DIV div" & Chr(10)
        strSQL = strSQL & "left join #RCVRETDIV R on R.productcode = p.productcode and R.YRno = d.Yrno and R.Mno = d.Mno and r.DivNo = div.divno" & Chr(10)
        strSQL = strSQL & "left join #RCVQTYDIV RC on RC.ProductCode = p.ProductCode and RC.Mno = d.Mno and RC.YRno = d.YRno and rc.divno  = div.divno" & Chr(10)
        strSQL = strSQL & "left join #POSDIV pos on pos.ProductCode = p.ProductCode and pos.Mno = d.Mno and pos.YRno = d.Yrno and pos.DivNo = div.divno" & Chr(10)
        strSQL = strSQL & "order by productcode, YRno, MNo,div.divno" & Chr(10)
        strSQL = strSQL & "" & Chr(10)
        If InStr(1, sBuildCode, "Ur") > 0 Then
            strSQL = strSQL & "select d.YRno , d.Wkno, p.productcode" & Chr(10)
            strSQL = strSQL & ", isnull(R.Retail / nullif(s.noofstores,0),0) as RCVRetailpwk, isnull(RetailNet/ nullif(s.noofstores,0),0) as RCVRetailNetpwk, isnull(Cost/ nullif(s.noofstores,0),0) as RCVCostpwk" & Chr(10)
            strSQL = strSQL & ", isnull(RC.Quantity/ nullif(s.noofstores,0), 0) as RCVQTYpwk, isnull(RC.Pieces/ nullif(s.noofstores,0), 0) as RCVPiecespwk, isnull(pos.QTY/ nullif(s.noofstores,0),0) as POSQTYpwk" & Chr(10)
            strSQL = strSQL & ", isnull(pos.Retail/ nullif(s.noofstores,0),0) as POSRetailpwk, isnull(d.daycnt,0) as DayCnt, isnull(s.noofstores,0) as NoOfStores,isnull(pr.CGNo,0) as CGno, isnull(div.divno,0) as Divno" & Chr(10)
            strSQL = strSQL & "into #SALESWKDIV " & Chr(10)
            strSQL = strSQL & "from #DATEWK d cross join #PList p cross join #DIV div" & Chr(10)
            strSQL = strSQL & "left join #RCVRETWKDIV R on R.productcode = p.productcode and R.YRno = d.Yrno and R.wkno = d.wkno and r.DivNo =div.DivNo" & Chr(10)
            strSQL = strSQL & "left join #RCVQTYWKDIV RC on RC.ProductCode = p.ProductCode and RC.wkno = d.wkno and RC.YRno = d.YRno and rc.DivNo = div.DivNo" & Chr(10)
            strSQL = strSQL & "left join #POSWKDIV pos on pos.ProductCode = p.ProductCode and pos.wkno = d.wkno and pos.YRno = d.Yrno and pos.DivNo = div.DivNo" & Chr(10)
            strSQL = strSQL & "left join #STORESWKDIV s on s.WeekNo = d.wkno and s.YearNo = d.Yrno and s.DivNo = div.DivNo" & Chr(10)
            strSQL = strSQL & "left join #P pr on pr.productcode = p.ProductCode" & Chr(10)
            strSQL = strSQL & "order by productcode, YRno, wkno, div.divno" & Chr(10)
            strSQL = strSQL & "" & Chr(10)
        End If
        rsSalesRS.Open strSQL, cnCCN
    End If
    If SalesDatatoBuild = "Sales" Then
        If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 0, 0, "Packaging National Sales Data": DoEvents
        Set RS = New ADODB.Recordset
        strSQL = "Select * from #SALES order by productcode,YRno , Mno"
        RS.Open strSQL, cnCCN
        If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 0, 0, "Indexing National Sales Data": DoEvents
        bSalesBool = BuildSalesDictionary(RS, psdSalesIndex)
        If bSalesBool = False Then Set psdSalesIndex = New Scripting.Dictionary
    End If

    If SalesDatatoBuild = "USW" Then
        If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 0, 0, "Packaging National USW Data": DoEvents
        Set RS = New ADODB.Recordset
        strSQL = "Select * from #SALESWK  order by productcode,yrno, wkno"
        RS.Open strSQL, cnCCN
        bSalesUSWBool = BuildSalesDictionary(RS, psdUSWIndex)
        If bSalesUSWBool = False Then Set psdUSWIndex = New Scripting.Dictionary
'        For Each v In psdUSWIndex(2390)
'            Set PD = psdUSWIndex(2390)(v)
'            Debug.Print PD.dblPOSQTYpwk
'        Next
        
    End If

    If SalesDatatoBuild = "SalesDiv" Then
        If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 0, 0, "Packaging Regional Sales Data": DoEvents
        Set RS = New ADODB.Recordset
        strSQL = "Select * from #SALESDIV order by productcode, YRno , Mno,divno "
        RS.Open strSQL, cnCCN
        bSalesDivBool = BuildSalesDictionary(RS, psdSalesDivIndex)
        If bSalesDivBool = False Then Set psdSalesDivIndex = New Scripting.Dictionary
    End If

    If SalesDatatoBuild = "USWDiv" Then
        If InStr(1, sBuildCode, "Ur") > 0 Then
            If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 0, 0, "Packaging Regional USW Data": DoEvents
            Set RS = New ADODB.Recordset
            strSQL = "Select * from #SALESWKDIV order by productcode, divno, yrno, wkno"
            RS.Open strSQL, cnCCN
            bSalesUSWDivBool = BuildSalesDictionary(RS, psdUSWDivIndex)
            If bSalesUSWDivBool = False Then Set psdUSWDivIndex = New Scripting.Dictionary
        Else
            Set psdUSWDivIndex = New Scripting.Dictionary
        End If
    End If

    If SalesDatatoBuild = "ST" Then
        If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 0, 0, "Packaging Sell Through Data": DoEvents
        Set RS = New ADODB.Recordset
        strSQL = "Select * from #SALESSP14 where retail is not null order by ProductCode, OSD"
        RS.Open strSQL, cnCCN
        bSalesSTBool = BuildSalesDictionary(RS, psdSTIndex)
        If bSalesSTBool = False Then Set psdSTIndex = New Scripting.Dictionary
    End If
    buildSalesDataRS = True
    
End Function
Private Function BuildSalesDictionary(ByRef RS As ADODB.Recordset, ByRef ProdDic As Scripting.Dictionary) As Boolean
Dim curProd As Long, a As Long, b As Long, curYear As Long, cnt As Long, totcnt As Long
Dim PointDic As Scripting.Dictionary
Dim PD As cCBA_ProdPeriodSalesData
Dim ASTO As Boolean
Dim lcol As Long
    On Error GoTo Err_Routine
    CBA_Error = ""

    
    If RS.Fields(0).Name = "productcode" Then lcol = 0 Else lcol = 2
    Set ProdDic = New Scripting.Dictionary: Set PointDic = New Scripting.Dictionary
    curProd = 0
    RS.MoveFirst
    a = 0: cnt = -1: curProd = CLng(RS(lcol))
    Do Until RS.EOF
        cnt = cnt + 1
        If CLng(RS(lcol)) <> curProd Then
            Set PD = New cCBA_ProdPeriodSalesData
            PD.lPcode = curProd
            ProdDic.Add curProd, PD
            PointDic.Add curProd, a
            a = 0
            curProd = CLng(RS(lcol))
        End If
        a = a + 1
        RS.MoveNext
    Loop
    If curProd > 0 Then
        Set PD = New cCBA_ProdPeriodSalesData
        PD.lPcode = curProd
        ProdDic.Add curProd, PD
        PointDic.Add curProd, a
    End If
    totcnt = cnt: cnt = -1
    
    If InStr(1, sBuildCode, "As") > 0 Then If bHasAlcoholStores = False Then getAlcoholStoreNos
    
    
    RS.MoveFirst
    Do Until RS.EOF
        Set PD = ProdDic(CLng(RS(lcol)))
        If RS.Fields.Count = 11 Then
            If RS.Fields(10).Name = "DivNo" Then
                If RS.Fields(1).Name = "Wkno" Then
                    PD.Build RS, PointDic(CLng(RS(lcol))), "USWDiv", IIf(InStr(1, sBuildCode, "As") > 0, sdAlcoholStoreNos, Nothing)
                Else
                    PD.Build RS, PointDic(CLng(RS(lcol))), "SalesDiv", IIf(InStr(1, sBuildCode, "As") > 0, sdAlcoholStoreNos, Nothing)
                End If
            End If
        ElseIf RS.Fields.Count = 14 Then
            If RS.Fields(1).Name = "Wkno" Then
                PD.Build RS, PointDic(CLng(RS(lcol))), "USWDiv", IIf(InStr(1, sBuildCode, "As") > 0, sdAlcoholStoreNos, Nothing)
            Else
                PD.Build RS, PointDic(CLng(RS(lcol))), "SalesDiv", IIf(InStr(1, sBuildCode, "As") > 0, sdAlcoholStoreNos, Nothing)
            End If
        
        Else
            If RS.Fields(0).Name = "productcode" Then
                PD.Build RS, PointDic(CLng(RS(lcol))), "ST"
            ElseIf RS.Fields(1).Name = "Wkno" Then
                PD.Build RS, PointDic(CLng(RS(lcol))), "USW", IIf(InStr(1, sBuildCode, "As") > 0, sdAlcoholStoreNos, Nothing)
            Else
                PD.Build RS, PointDic(CLng(RS(lcol))), "Sales", IIf(InStr(1, sBuildCode, "As") > 0, sdAlcoholStoreNos, Nothing)
            End If
        End If
    Loop
    
    BuildSalesDictionary = True
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_ProdGroup.GenProductDic", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Com", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next


End Function
'Private Function BuildSalesDictionary(ByRef RS As ADODB.Recordset, ByRef ProdDic As Scripting.Dictionary) As Boolean
'Dim curProd As Long, a As Long, b As Long, curYear As Long, cnt As Long, totcnt As Long
'Dim SalesDic As Scripting.Dictionary
'Dim PD As cCBA_ProdPeriodSalesData
'Dim ASTO As Boolean
'Dim lCol As Long
'    On Error GoTo Err_Routine
'    CBA_Error = ""
'
'
'    If RS.Fields(0).Name = "productcode" Then lCol = 0 Else lCol = 2
'
'    curProd = 0
'    RS.MoveFirst
'    cnt = -1
'    Do Until RS.EOF
'        cnt = cnt + 1
'        RS.MoveNext
'    Loop
'    totcnt = cnt: cnt = -1
'
'    If InStr(1, sBuildCode, "As") > 0 Then If bHasAlcoholStores = False Then getAlcoholStoreNos
'
'
'        ''Productcode | Year | Month/Week | Field
'    Set ProdDic = New Scripting.Dictionary
'    RS.MoveFirst
'    Do Until RS.EOF
'        If curProd <> RS(lCol) Then
'            If curProd = 0 Then
'                curProd = CLng(RS(lCol))
'            Else
'                ProdDic.Add curProd, SalesDic
'                curProd = CLng(RS(lCol))
'            End If
'            Set SalesDic = New Scripting.Dictionary
'        End If
'        cnt = cnt + 1
'        If cnt / Round(totcnt / 20, 0) = Round(cnt / Round(totcnt / 20, 0), 0) Then If CBA_BasicFunctions.isRunningSheetDisplayed = True And totcnt > 0 Then CBA_BasicFunctions.RunningSheetAddComment 1, 1, "Packaging National Sales/USW DataCube: " & Round((cnt / totcnt), 2) * 100 & "%"
'
'        Set PD = New cCBA_ProdPeriodSalesData
'        For a = 0 To RS.Fields.Count - 1
'            If InStr(1, sBuildCode, "As") > 0 Then
'                ASTO = 0
'                If sdAlcoholStoreNos.Exists(RS(0)) Then
'                    If sdAlcoholStoreNos(RS(0)).Exists((IIf(RS.Fields(1).Name = "Wkno", Month(CBA_BasicFunctions.GetDayFromWeekNumber(RS(0), RS(1))), RS(1)))) Then
'                        If sdAlcoholStoreNos(RS(0))(IIf(RS.Fields(1).Name = "Wkno", Month(CBA_BasicFunctions.GetDayFromWeekNumber(RS(0), RS(1))), RS(1)))(599) Then
'                            ASTO = sdAlcoholStoreNos(RS(0))(IIf(RS.Fields(1).Name = "Wkno", Month(CBA_BasicFunctions.GetDayFromWeekNumber(RS(0), RS(1))), RS(1)))(599)
'                        End If
'                    End If
'                End If
'            Else
'                Select Case LCase(RS.Fields(a).Name)
'                    Case "yrno": PD.lYRno = CLng(RS(a))
'                    Case "wkno": PD.lWkno = CLng(RS(a))
'                    Case "mno": PD.lMNo = CLng(RS(a))
'                    Case "productcode": PD.lProductcode = CLng(RS(a))
'                    Case "rcvretailpwk": If ASTO > 0 Then PD.dblRCVRetailpwk = (RS(a) * RS(11)) / ASTO Else PD.dblRCVRetailpwk = Cdbl(RS(a))
'                    Case "rcvretailnetpwk": If ASTO > 0 Then PD.dblRCVRetailNetpwk = (RS(a) * RS(11)) / ASTO Else PD.dblRCVRetailNetpwk = Cdbl(RS(a))
'                    Case "rcvcostpwk": If ASTO > 0 Then PD.dblRCVCostpwk = (RS(a) * RS(11)) / ASTO Else PD.dblRCVCostpwk = Cdbl(RS(a))
'                    Case "rcvqtypwk": If ASTO > 0 Then PD.dblRCVQTYpwk = (RS(a) * RS(11)) / ASTO Else PD.dblRCVQTYpwk = Cdbl(RS(a))
'                    Case "rcvpiecespwk": If ASTO > 0 Then PD.dblRCVPiecespwk = (RS(a) * RS(11)) / ASTO Else PD.dblRCVPiecespwk = Cdbl(RS(a))
'                    Case "posqtypwk": If ASTO > 0 Then PD.dblPOSQTYpwk = (RS(a) * RS(11)) / ASTO Else PD.dblPOSQTYpwk = Cdbl(RS(a))
'                    Case "posretailpwk": If ASTO > 0 Then PD.dblPOSRetailpwk = (RS(a) * RS(11)) / ASTO Else PD.dblPOSRetailpwk = Cdbl(RS(a))
'                    Case "daycnt": PD.lDayCnt = CLng(RS(a))
'                    Case "noofstores": PD.lNoOfStores = CLng(RS(a))
'                    Case "cgno": PD.lCGno = CLng(RS(a))
'                    Case "divno": PD.lDivno = CLng(RS(a))
'                    Case "rcvretail": If ASTO > 0 Then PD.dblRCVRetail = (RS(a) * RS(11)) / ASTO Else PD.dblRCVRetail = Cdbl(RS(a))
'                    Case "rcvretailnet": If ASTO > 0 Then PD.dblRCVRetailNet = (RS(a) * RS(11)) / ASTO Else PD.dblRCVRetailNet = Cdbl(RS(a))
'                    Case "rcvcost": If ASTO > 0 Then PD.dblRCVCost = (RS(a) * RS(11)) / ASTO Else PD.dblRCVCost = Cdbl(RS(a))
'                    Case "rcvqty": If ASTO > 0 Then PD.dblRCVQTY = (RS(a) * RS(11)) / ASTO Else PD.dblRCVQTY = Cdbl(RS(a))
'                    Case "rcvpieces": If ASTO > 0 Then PD.dblRCVPieces = (RS(a) * RS(11)) / ASTO Else PD.dblRCVPieces = Cdbl(RS(a))
'                    Case "posqty": If ASTO > 0 Then PD.dblPOSQTY = (RS(a) * RS(11)) / ASTO Else PD.dblPOSQTY = Cdbl(RS(a))
'                    Case "posretail": If ASTO > 0 Then PD.dblPOSRetail = (RS(a) * RS(11)) / ASTO Else PD.dblPOSRetail = Cdbl(RS(a))
'                    Case "osd": PD.dOSD = CDate(RS(a))
'                    Case "eod": PD.dEOD = CDate(RS(a))
'                End Select
'            End If
'        Next
'        If lCol = 0 Then
'            SalesDic.Add "OSD" & Format(PD.dOSD, "YYYYMMDD"), PD
'        ElseIf PD.lDivno > 0 Then
'            If PD.lWkno = 0 Then
'                SalesDic.Add "DMY" & Format(PD.lDivno, "000") & Format(PD.lMNo, "00") & "-" & Format(PD.lYRno, "00"), PD
'            Else
'                SalesDic.Add "DWY" & Format(PD.lDivno, "000") & Format(PD.lWkno, "00") & "-" & Format(PD.lYRno, "00"), PD
'            End If
'        Else
'            If PD.lWkno = 0 Then
'                SalesDic.Add "MY" & Format(PD.lMNo, "00") & "-" & Format(PD.lYRno, "00"), PD
'            Else
'                SalesDic.Add "WY" & Format(PD.lWkno, "00") & "-" & Format(PD.lYRno, "00"), PD
'            End If
'        End If
'        RS.MoveNext
'    Loop
'    ProdDic.Add curProd, SalesDic
'    If ProdDic Is Nothing Then BuildSalesDictionary = False Else BuildSalesDictionary = True
'
'Exit_Routine:
'
'    On Error Resume Next
'    Exit Function
'
'Err_Routine:
'    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_ProdGroup.GenProductDic", 3)
'    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
'    Debug.Print CBA_Error
'    Stop            ' ^RW Camera + next line
'    Resume Next
'    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
'    Call g_Write_Err_Table(Err, CBA_Error, "Com", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
'    GoTo Exit_Routine
'    Resume Next
'
'
'End Function
Private Function BuildDivSalesDictionary(ByRef RS As ADODB.Recordset) As Scripting.Dictionary
Dim curProd As Long, a As Long, b As Long, curYear As Long, cnt As Long, totcnt As Long
Dim yearDic As Scripting.Dictionary
Dim PeriodDic As Scripting.Dictionary
Dim psdSalesIndex As Scripting.Dictionary
Dim FieldDic As Scripting.Dictionary
Dim SalesDic As Scripting.Dictionary
Dim DivDic As Scripting.Dictionary
Dim curDiv As Long
    curProd = 0
    RS.MoveFirst
    cnt = -1
    Do Until RS.EOF
        cnt = cnt + 1
        RS.MoveNext
    Loop
    totcnt = cnt: cnt = -1

        ''Productcode | Year | Month/Week | Field
    Set SalesDic = New Scripting.Dictionary
    Set DivDic = New Scripting.Dictionary
    RS.MoveFirst
    Do Until RS.EOF
        cnt = cnt + 1
        If cnt / Round(totcnt / 20, 0) = Round(cnt / Round(totcnt / 20, 0), 0) Then If CBA_BasicFunctions.isRunningSheetDisplayed = True And totcnt > 0 Then CBA_BasicFunctions.RunningSheetAddComment 1, 1, "Packaging Regional Sales/USW DataCubes: " & Round((cnt / totcnt), 2) * 100 & "%"

        If RS(2) <> curProd Then
            If curProd > 0 Then
                If yearDic.Exists(curYear) Then
                    MsgBox "Issue with Sales Data not going into Dictionary"
                    Stop
                Else
                    If curProd <> 0 Then yearDic.Add curYear, PeriodDic: DivDic.Add curDiv, yearDic: SalesDic.Add curProd, DivDic
                End If
            End If
            curProd = CLng(RS(2))
            Set DivDic = New Scripting.Dictionary
            Set yearDic = New Scripting.Dictionary
            Set PeriodDic = New Scripting.Dictionary
            curDiv = CLng(RS(10))
            curYear = CLng(RS(0))
            DoEvents
        ElseIf RS(10) <> curDiv Then
            If DivDic.Exists(curDiv) Then
                MsgBox "Issue with Sales Data not going into Dictionary"
                Stop
            Else
                If curDiv <> 0 Then yearDic.Add curYear, PeriodDic: DivDic.Add curDiv, yearDic
            End If
            curDiv = CLng(RS(10))
            Set yearDic = New Scripting.Dictionary
            Set PeriodDic = New Scripting.Dictionary
            curYear = CLng(RS(0))
        ElseIf RS(0) <> curYear Then
            If yearDic.Exists(curYear) Then
                MsgBox "Issue with Sales Data not going into Dictionary"
                Stop
            Else
                If curYear <> 0 Then yearDic.Add curYear, PeriodDic
            End If
            Set PeriodDic = New Scripting.Dictionary
            curYear = CLng(RS(0))
        End If
        Set FieldDic = New Scripting.Dictionary
        For a = 3 To RS.Fields.Count - 1
            If FieldDic.Exists(CStr(RS.Fields(a).Name)) Then
                MsgBox "Issue with Sales Data not going into Dictionary"
                Stop
            Else
                If InStr(1, sBuildCode, "As") > 0 Then
                    If sdAlcoholStoreNos.Exists(RS(0)) Then
                        If sdAlcoholStoreNos(RS(0)).Exists((IIf(RS.Fields(1).Name = "Wkno", Month(CBA_BasicFunctions.GetDayFromWeekNumber(RS(0), RS(1))), RS(1)))) Then
                            If sdAlcoholStoreNos(RS(0))(IIf(RS.Fields(1).Name = "Wkno", Month(CBA_BasicFunctions.GetDayFromWeekNumber(RS(0), RS(1))), RS(1)))(RS(13)) Then
                                If a = 11 Then
                                    FieldDic.Add CStr(RS.Fields(a).Name), CStr(sdAlcoholStoreNos(RS(0))(IIf(RS.Fields(1).Name = "Wkno", Month(CBA_BasicFunctions.GetDayFromWeekNumber(RS(0), RS(1))), RS(1)))(RS(13)))
                                Else
                                    FieldDic.Add CStr(RS.Fields(a).Name), CStr((RS(a) * RS(11)) / sdAlcoholStoreNos(RS(0))(IIf(RS.Fields(1).Name = "Wkno", Month(CBA_BasicFunctions.GetDayFromWeekNumber(RS(0), RS(1))), RS(1)))(RS(13)))
                                End If
                            End If
                        End If
                    End If
                Else
                    FieldDic.Add CStr(RS.Fields(a).Name), CDbl(RS(a))
                End If
            End If
        Next
        If PeriodDic.Exists(CLng(RS(1))) Then
            MsgBox "Issue with Sales Data not going into Dictionary"
            Stop
        Else
            PeriodDic.Add CLng(RS(1)), FieldDic
        End If
        RS.MoveNext
        If RS.EOF Then yearDic.Add curYear, PeriodDic: DivDic.Add curDiv, yearDic: SalesDic.Add curProd, DivDic
    Loop
Set BuildDivSalesDictionary = SalesDic
Set SalesDic = Nothing
Set yearDic = Nothing
Set PeriodDic = Nothing
Set FieldDic = Nothing
Set DivDic = Nothing
End Function
Private Function BuildSellThroughSalesDictionary(ByRef RS As ADODB.Recordset) As Scripting.Dictionary
Dim curProd As Long, a As Long, b As Long, curOSD As Date, cnt As Long, totcnt As Long
Dim OSDDic As Scripting.Dictionary
Dim PeriodDic As Scripting.Dictionary
Dim psdSalesIndex As Scripting.Dictionary
Dim FieldDic As Scripting.Dictionary
Dim SalesDic As Scripting.Dictionary
    curProd = 0
    RS.MoveFirst
    cnt = -1
    Do Until RS.EOF
        cnt = cnt + 1
        RS.MoveNext
    Loop
    totcnt = cnt: cnt = -1
        ''Productcode | Year | Month/Week | Field
    Set SalesDic = New Scripting.Dictionary
    RS.MoveFirst
    Do Until RS.EOF
        cnt = cnt + 1
        If cnt / Round(totcnt / 20, 0) = Round(cnt / Round(totcnt / 20, 0), 0) Then If CBA_BasicFunctions.isRunningSheetDisplayed = True And totcnt > 0 Then CBA_BasicFunctions.RunningSheetAddComment 1, 1, "Packaging SellThrough DataCube: " & Round((cnt / totcnt), 2) * 100 & "%"

        If RS(0) <> curProd Then
            If curProd > 0 Then
                If OSDDic.Exists(curOSD) Then
                    MsgBox "Issue with Sales Data not going into Dictionary"
                    Stop
                Else
                    If SalesDic.Exists(curProd) Then
                        MsgBox "Issue with Sales Data not going into Dictionary"
                        Stop
                    Else
                        If curProd <> 0 Then
                            OSDDic.Add curOSD, FieldDic
                            SalesDic.Add curProd, OSDDic
                        End If
                    End If
                End If
            End If
'            If CurProd = 1025 Then
'            Stop
'            End If
            
            If IsNull(RS(0)) Then
                Stop
            End If
            curProd = CLng(RS(0))
            Set OSDDic = New Scripting.Dictionary
            Set PeriodDic = New Scripting.Dictionary
            If IsNull(RS(1)) Then
                Stop
            End If
            curOSD = CLng(RS(1))
            DoEvents
        ElseIf CLng(RS(1)) <> curOSD Then
            If OSDDic.Exists(curOSD) Then
                MsgBox "Issue with Sales Data not going into Dictionary"
                Stop
            Else
                If curProd <> 0 Then OSDDic.Add curOSD, FieldDic
            End If
            If IsNull(RS(1)) Then
                Stop
            End If
            curOSD = CLng(RS(1))
        End If
        Set FieldDic = New Scripting.Dictionary
        For a = 3 To RS.Fields.Count - 1
            If FieldDic.Exists(CStr(RS.Fields(a).Name)) Then
                MsgBox "Issue with Sales Data not going into Dictionary"
                Stop
            Else
                If IsNull(RS(a)) Then
                    Stop
                End If
                FieldDic.Add CStr(RS.Fields(a).Name), CDbl(RS(a))
            End If
        Next
        RS.MoveNext
        If RS.EOF Then
            OSDDic.Add curOSD, FieldDic
            SalesDic.Add curProd, OSDDic
        End If
    Loop
Set BuildSellThroughSalesDictionary = SalesDic


End Function

Private Function buildTotBusData(ByVal StartYear As Long, ByVal BuildCode As String) As Boolean
Dim strSQL As String
Dim Ydic As Scripting.Dictionary
Dim Mdic As Scripting.Dictionary
Dim lCurYear As Long
Dim h As Long
Dim TempRS As ADODB.Recordset

If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Generating Total Business Data..."

Set sdTotBus = New Scripting.Dictionary
Set TempRS = New ADODB.Recordset
strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
strSQL = strSQL & "declare @YRST nvarchar(4) = " & StartYear & Chr(10)
strSQL = strSQL & "" & Chr(10)
strSQL = strSQL & "select year(PDATE) as Yrno, Month(PDATE) as Mno into #DATE  from (" & Chr(10)
strSQL = strSQL & "SELECT  TOP (DATEDIFF(MONTH, @YRST + '-01-01', EOMONTH(getdate(),-1)) + 1)" & Chr(10)
strSQL = strSQL & "    Pdate = convert(date,DATEADD(MONTH, ROW_NUMBER() OVER(ORDER BY a.object_id) - 1, @YRST + '-01-01'))" & Chr(10)
strSQL = strSQL & "    FROM    sys.all_objects a CROSS JOIN sys.all_objects b" & Chr(10)
strSQL = strSQL & ") a" & Chr(10)
strSQL = strSQL & "" & Chr(10)
If InStr(1, BuildCode, "P") > 0 Then
    strSQL = strSQL & "select year(posdate) as YRno, month(posdate) as Mno, sum(retail) as retail, sum(quantity) as QTY" & Chr(10)
    strSQL = strSQL & "into #POS from cbis599p.dbo.pos pos" & Chr(10)
    strSQL = strSQL & "where YEAR(posdate) >= @YRST" & Chr(10)
    strSQL = strSQL & "group by year(posdate) , month(posdate)" & Chr(10)
    strSQL = strSQL & "" & Chr(10)
End If
If InStr(1, BuildCode, "R") > 0 Then
    strSQL = strSQL & "select year(dayenddate) as YRno, month(dayenddate) as Mno, sum(Quantity) as Quantity , sum(Pieces) as Pieces" & Chr(10)
    strSQL = strSQL & "into #RCVQTY from cbis599p.dbo.receiving rcv" & Chr(10)
    strSQL = strSQL & "where rcv.recordid = '001' and YEAR(dayenddate) >= @YRST" & Chr(10)
    strSQL = strSQL & "group by  year(dayenddate),month(dayenddate), rcv.recordid" & Chr(10)
    strSQL = strSQL & "" & Chr(10)
    strSQL = strSQL & "select YEAR(dayenddate) as YRno, Month(dayenddate) as Mno, sum(retail) as Retail, sum(retailnet) as RetailNet, sum(cost) as Cost" & Chr(10)
    strSQL = strSQL & "into #RCVRET from cbis599p.dbo.RECEIVING rcv" & Chr(10)
    strSQL = strSQL & "where YEAR(dayenddate) >= @YRST" & Chr(10)
    strSQL = strSQL & "group by  year(dayenddate), month(dayenddate)" & Chr(10)
    strSQL = strSQL & "" & Chr(10)
End If
strSQL = strSQL & "select d.YRno , d.Mno" & Chr(10)
If InStr(1, BuildCode, "R") > 0 Then strSQL = strSQL & ", isnull(R.Retail,0) as RCVRetail, isnull(RetailNet,0) as RCVRetailNet, isnull(Cost,0) as RCVCost" & Chr(10)
If InStr(1, BuildCode, "R") > 0 Then strSQL = strSQL & ", isnull(RC.Quantity, 0) as RCVQTY, isnull(RC.Pieces, 0) as RCVPieces" & Chr(10)
If InStr(1, BuildCode, "P") > 0 Then strSQL = strSQL & ", isnull(pos.QTY,0) as POSQTY, isnull(pos.Retail,0) as POSRetail" & Chr(10)
strSQL = strSQL & "into #SALES from #DATE d" & Chr(10)
If InStr(1, BuildCode, "R") > 0 Then strSQL = strSQL & "left join #RCVRET R on R.YRno = d.Yrno and R.Mno = d.Mno" & Chr(10)
If InStr(1, BuildCode, "R") > 0 Then strSQL = strSQL & "left join #RCVQTY RC on RC.Mno = d.Mno and RC.YRno = d.YRno" & Chr(10)
If InStr(1, BuildCode, "P") > 0 Then strSQL = strSQL & "left join #POS pos on pos.Mno = d.Mno and pos.YRno = d.Yrno" & Chr(10)
strSQL = strSQL & "order by YRno, MNo" & Chr(10)
strSQL = strSQL & "" & Chr(10)
If InStr(1, BuildCode, "R") > 0 Then strSQL = strSQL & "drop table #RCVQTY, #RCVRET" & Chr(10)
If InStr(1, BuildCode, "P") > 0 Then strSQL = strSQL & "drop table #POS" & Chr(10)
strSQL = strSQL & "" & Chr(10)
strSQL = strSQL & "select * from #SALES order by YRno, MNo" & Chr(10)
strSQL = strSQL & "" & Chr(10)
strSQL = strSQL & "drop table #DATE, #SALES" & Chr(10)
TempRS.Open strSQL, cnCCN

If TempRS.EOF Then
    Debug.Print "Error in Generation of Total business Data"
    buildTotBusData = False
Else
    Set sdTotBus = New Scripting.Dictionary
    Set Ydic = New Scripting.Dictionary
    lCurYear = TempRS.Fields(0) '': curMth = 0
    With TempRS
        Do Until .EOF
            If lCurYear <> .Fields(0) Then
                sdTotBus.Add CStr(lCurYear), Ydic
                Set Ydic = New Scripting.Dictionary
                lCurYear = .Fields(0)
            End If
            Set Mdic = New Scripting.Dictionary
            For h = 0 To .Fields.Count - 1
                If .Fields(h).Name <> "YRno" And .Fields(h).Name <> "Mno" Then
                    Mdic.Add CStr(.Fields(h).Name), CStr(.Fields(h))
                End If
            Next
            Ydic.Add CStr(.Fields(1)), Mdic
            .MoveNext
        Loop
        sdTotBus.Add CStr(lCurYear), Ydic
    End With
    buildTotBusData = True
End If


End Function
Private Function RunConnectionSetups(ByVal CBIS As Boolean, ByVal MMS As Boolean, ByVal Spaceman As Boolean, _
    ByVal COMRADE_STAR As Boolean, ByVal CAMERA As Boolean) As Boolean
Dim div As Long
'''''''-----------TEST VARIABLES---------------------
'Variables for Development
Set sdRegDivMiss = New Scripting.Dictionary
sdRegDivMiss.Add 508, 508
'End of test variables
'''''''-----------END TEST VARIABLES---------------------
RunConnectionSetups = True
If COMRADE_STAR = True Then
    If cnSCN Is Nothing Then
        Set cnSCN = New ADODB.Connection
        With cnSCN
            .ConnectionTimeout = 50
            .CommandTimeout = 50
            .Open "Provider=SQLNCLI10;DATA SOURCE=" & CBA_BasicFunctions.TranslateServerName("599DBL12", Date) & ";;INTEGRATED SECURITY=sspi;"
        End With
    End If
    If cnSCN.State <> 1 Then RunConnectionSetups = False
End If
If Spaceman = True Then
    If cnSMCN Is Nothing Then
        Set cnSMCN = New ADODB.Connection
        With cnSMCN
            .ConnectionTimeout = 50
            .CommandTimeout = 50
            .Open "Provider=SQLNCLI10;DATA SOURCE=599DBL11;;INTEGRATED SECURITY=sspi;"
        End With
    End If
    If cnSMCN.State <> 1 Then RunConnectionSetups = False
End If
If CBIS = True Then
    If cnCCN Is Nothing Then
        Set cnCCN = New ADODB.Connection
        With cnCCN
            .ConnectionTimeout = 50
            .CommandTimeout = 50
'            .Open "Provider=SQLOLEDB;DATA SOURCE=" & NATIONAL_DB & "DBL01;;INTEGRATED SECURITY=sspi;DataTypeCompatibility=80"
            .Open "Provider=SQLNCLI10;DATA SOURCE=" & NATIONAL_DB & "DBL01;;INTEGRATED SECURITY=sspi;"
        End With
    End If
    If cnCCN.State <> 1 Then RunConnectionSetups = False
End If
If MMS = True Then
    For div = RDIVFROM To RDIVTO
        If pcnRCN(div) Is Nothing Then
            If sdRegDivMiss.Exists(div) = False Then
                Set pcnRCN(div) = New ADODB.Connection
                With pcnRCN(div)
                    .ConnectionTimeout = 50
                    .CommandTimeout = 50
                    .Open "Provider=SQLNCLI10;DATA SOURCE=0" & div & "Z0IDBSRVL02;;INTEGRATED SECURITY=sspi;"
                End With
            End If
        End If
        If sdRegDivMiss.Exists(div) = False Then If pcnRCN(div).State <> 1 Then RunConnectionSetups = False
    Next
End If
If CAMERA = True Then
    If cnCRCN Is Nothing Then
        Set cnCRCN = New ADODB.Connection
        With cnCRCN
            .ConnectionTimeout = 50
            .CommandTimeout = 50
            .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Cam")
        End With
    End If
    If cnCRCN.State <> 1 Then RunConnectionSetups = False
End If


End Function
Private Function CloseConnections(ByVal CBIS As Boolean, ByVal MMS As Boolean, ByVal Spaceman As Boolean, _
    ByVal COMRADE_STAR As Boolean, ByVal CAMERA As Boolean) As Boolean
Dim div As Long
'    If CBIS = False And MMS = False And Spaceman = False And COMRADE_STAR = False And CATREV = False Then
'        CBIS = True: MMS = True: Spaceman = True: COMRADE_STAR = True: CAMERA = True
'    End If

    If Spaceman = True Then If Not cnSMCN Is Nothing Then cnSMCN.Close: Set cnSMCN = Nothing
    If CBIS = True Then If Not cnCCN Is Nothing Then cnCCN.Close: Set cnCCN = Nothing
    If COMRADE_STAR = True Then If Not cnSCN Is Nothing Then cnSCN.Close: Set cnSCN = Nothing
    If CAMERA = True Then If Not cnCRCN Is Nothing Then cnCRCN.Close: Set cnCRCN = Nothing
    If MMS = True Then
        For div = RDIVFROM To RDIVTO
            If Not pcnRCN(div) Is Nothing Then pcnRCN(div).Close: Set pcnRCN(div) = Nothing
        Next
    End If
    CloseConnections = True
End Function
Private Function SetProdGroupDates(ByVal MonthFrom As Date, ByVal MonthTo As Date) As Boolean
    If bBDataGenerated = False Then
        If MonthFrom > MonthTo Then
            'Month To date is before Month Start date!! Silly!
            SetProdGroupDates = False
        Else
            dtMnthTo = MonthTo
            dtMnthFrom = MonthFrom
            SetProdGroupDates = True
        End If
    Else
        'Cannot set new dates when data is already built! just create a new instance of the class!
        SetProdGroupDates = False
    End If
End Function
Private Property Get cnCCN() As ADODB.Connection: Set cnCCN = pcnCCN: End Property
Private Property Set cnCCN(ByVal objNewValue As ADODB.Connection): Set pcnCCN = objNewValue: End Property
Private Property Get cnSCN() As ADODB.Connection: Set cnSCN = pcnSCN: End Property
Private Property Set cnSCN(ByVal objNewValue As ADODB.Connection): Set pcnSCN = objNewValue: End Property
Private Property Get cnSMCN() As ADODB.Connection: Set cnSMCN = pcnSMCN: End Property
Private Property Set cnSMCN(ByVal objNewValue As ADODB.Connection): Set pcnSMCN = objNewValue: End Property
Private Property Get cnCRCN() As ADODB.Connection: Set cnCRCN = pcnCRCN: End Property
Private Property Set cnCRCN(ByVal objNewValue As ADODB.Connection): Set pcnCRCN = objNewValue: End Property
Private Property Get rsSalesRS() As ADODB.Recordset: Set rsSalesRS = prsSalesRS: End Property
Private Property Set rsSalesRS(ByVal objNewValue As ADODB.Recordset): Set prsSalesRS = objNewValue: End Property
Private Property Get rsBaseProdRS() As ADODB.Recordset: Set rsBaseProdRS = prsBaseProdRS: End Property
Private Property Set rsBaseProdRS(ByVal objNewValue As ADODB.Recordset): Set prsBaseProdRS = objNewValue: End Property
Public Property Get sdTotBus() As Scripting.Dictionary: Set sdTotBus = psdTotBus: End Property
Private Property Set sdTotBus(ByVal objNewValue As Scripting.Dictionary): Set psdTotBus = objNewValue: End Property
Private Property Get sdRCost() As Scripting.Dictionary(): sdRCost = psdRCost: End Property
Private Property Get sdRegDivMiss() As Scripting.Dictionary: Set sdRegDivMiss = psdRegDivMiss: End Property
Private Property Set sdRegDivMiss(ByVal objNewValue As Scripting.Dictionary): Set psdRegDivMiss = objNewValue: End Property
Public Property Get sdCOMdic() As Scripting.Dictionary: Set sdCOMdic = psdCOMdic: End Property
Private Property Set sdCOMdic(ByVal objNewValue As Scripting.Dictionary): Set psdCOMdic = objNewValue: End Property
Private Property Get sdCG_PDic() As Scripting.Dictionary: Set sdCG_PDic = psdCG_PDic: End Property
Private Property Set sdCG_PDic(ByVal objNewValue As Scripting.Dictionary): Set psdCG_PDic = objNewValue: End Property
Private Property Get sdProductLabelDic() As Scripting.Dictionary: Set sdProductLabelDic = psdProductLabelDic: End Property
Private Property Set sdProductLabelDic(ByVal objNewValue As Scripting.Dictionary): Set psdProductLabelDic = objNewValue: End Property
Private Property Get sdWastageIndex() As Scripting.Dictionary: Set sdWastageIndex = psdWastageIndex: End Property
Private Property Set sdWastageIndex(ByVal objNewValue As Scripting.Dictionary): Set psdWastageIndex = objNewValue: End Property
Public Property Get colCGs() As Collection: Set colCGs = pcolCGs: End Property
Private Property Set colCGs(ByVal objNewValue As Collection): Set pcolCGs = objNewValue: End Property
Private Property Get vCGList() As Variant(): vCGList = pvCGList: End Property
Private Property Let vCGList(NewValue() As Variant): pvCGList = NewValue: End Property
Private Property Get vACGList() As Variant(): vACGList = pvACGList: End Property
Private Property Let vACGList(NewValue() As Variant): pvACGList = NewValue: End Property
Public Property Get bACG() As Boolean: bACG = pbACG: End Property
Private Property Let bACG(ByVal bNewValue As Boolean): pbACG = bNewValue: End Property
Public Property Get dtMnthFrom() As Date: dtMnthFrom = pdtMnthFrom: End Property
Private Property Let dtMnthFrom(ByVal dtNewValue As Date): pdtMnthFrom = dtNewValue: End Property
Public Property Get dtMnthTo() As Date: dtMnthTo = pdtMnthTo: End Property
Private Property Let dtMnthTo(ByVal dtNewValue As Date): pdtMnthTo = dtNewValue: End Property
Public Property Get sBuildCode() As String: sBuildCode = psBuildCode: End Property
Private Property Let sBuildCode(ByVal sNewValue As String): psBuildCode = sNewValue: End Property
Public Property Get bBDataGenerated() As Boolean: bBDataGenerated = pbBDataGenerated: End Property
Private Property Let lPcode(ByVal lNewValue As Long)
    plPcode = lNewValue
    If lNewValue > 0 Then bRunFordoubleProd = True Else bRunFordoubleProd = False
End Property
Public Property Get lPcode() As Long: lPcode = plPcode: End Property
Public Property Get bHasAlcoholStores() As Boolean: bHasAlcoholStores = pbHasAlcoholStores: End Property
Private Property Let bHasAlcoholStores(ByVal bNewValue As Boolean): pbHasAlcoholStores = bNewValue: End Property
Private Property Get bHasProdToMSegData() As Boolean: bHasProdToMSegData = pbHasProdToMSegData: End Property
Private Property Let bHasProdToMSegData(ByVal NewValue As Boolean): pbHasProdToMSegData = NewValue: End Property
Public Property Get sdAlcoholStoreNos() As Scripting.Dictionary: Set sdAlcoholStoreNos = psdAlcoholStoreNos: End Property
Private Property Set sdAlcoholStoreNos(ByVal NewValue As Scripting.Dictionary): Set psdAlcoholStoreNos = NewValue: bHasAlcoholStores = True: End Property
Public Property Get bSalesBool() As Boolean: bSalesBool = pbSalesBool: End Property
Private Property Let bSalesBool(ByVal NewValue As Boolean): pbSalesBool = NewValue: End Property
Public Property Get bSalesDivBool() As Boolean: bSalesDivBool = pbSalesDivBool: End Property
Private Property Let bSalesDivBool(ByVal NewValue As Boolean): pbSalesDivBool = NewValue: End Property
Public Property Get bSalesUSWBool() As Boolean: bSalesUSWBool = pbSalesUSWBool: End Property
Private Property Let bSalesUSWBool(ByVal NewValue As Boolean): pbSalesUSWBool = NewValue: End Property
Public Property Get bSalesUSWDivBool() As Boolean: bSalesUSWDivBool = pbSalesUSWDivBool: End Property
Private Property Let bSalesUSWDivBool(ByVal NewValue As Boolean): pbSalesUSWDivBool = NewValue: End Property
Public Property Get bSalesSTBool() As Boolean: bSalesSTBool = pbSalesSTBool: End Property
Private Property Let bSalesSTBool(ByVal NewValue As Boolean): pbSalesSTBool = NewValue: End Property
Private Sub Class_Terminate()
    Set psdCOMdic = Nothing
    Set psdCG_PDic = Nothing
    Set psdProductLabelDic = Nothing
    Set psdWastageIndex = Nothing
End Sub
Private Property Get lDoc_ID() As Long: lDoc_ID = plDoc_ID: End Property
Public Property Let lDoc_ID(ByVal NewValue As Long)
Dim v As Variant
plDoc_ID = NewValue
For Each v In psdCG_PDic
    psdCG_PDic(v).lDoc_ID = NewValue
Next

End Property
