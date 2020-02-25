VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_AST_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private cdtPromoEndDt As Date                           ' Date the Promotion Ends (One Year Ago)
Private cdtFromDt As Date                               ' Date the Promotion Starts (One Year Ago)
Private cdtToDt As Date                                 ' Date the Current Promotion Ends
Private clPCode As Long                                 ' The current Product Code
Private clPDID As Long                                  ' The Product ID  - This will be 999999999 if an Add (not saved to DB)
Private cUnit_Cost As Currency                          ' The National Unit Cost
Private cUnit_CostDiv As Scripting.Dictionary           ' The Regional Unit Cost
Private cSupplier_Cost_Support As Currency              ' The National Supplier Cost Support (per unit val)
Private cSupplier_Cost_SupportDiv As Scripting.Dictionary 'the Regional Supplier Cost Support (per unit val)
Private cPrior_Sales As Long                            ' The National POS QTY Sales for the Promotion (One Year Ago) Date Range
Private cPrior_SalesDiv As Scripting.Dictionary         ' The Regional POS QTY Sales for the Promotion (One Year Ago) Date Range
Private cPrior_TotalRetailDiv As Scripting.Dictionary   ' The Regional POS Retail Sales for the Promotion (One Year Ago) Date Range
Private cOrigEstSales As Long                           ' The National Estimated (Store Adjusted) Total POS QTY Sales
Private cEstSalesDiv As Scripting.Dictionary            ' The Regional Estimated (Store Adjusted) Total POS QTY Sales
Private cOrigCalcSales As Long                          ' The National Calculated (Entered by User) Total POS QTY Sales
Private cCalcSalesDiv As Scripting.Dictionary           ' The Regional  Calculated (Entered by User) Total POS QTY Sales
Private cUPSPW As Long                                  ' The National USW Value for the current/future SS Promotion
Private cUPSPWDiv As Scripting.Dictionary               ' The Regional USW Value for the current/future SS Promotion
Private clWoS As Long                                   ' The Weeks of Sale for all all product date ranges
Private cFill_Qty As Long                               ' The National Merch Fill Quantity (Extrapolated from division)
Private cFill_QtyDiv As Scripting.Dictionary            ' The Regional Merch Fill Quantity (Entered by User)
Private cCurr_Retail_Price As Currency                  ' The Retail Price (e.g. $1.99) for 501 region currently
Private cCurr_Retail_PriceDiv As Scripting.Dictionary   ' The Regional Price(e.g. $1.99) (from CBIS)
Private cRetail_Price As Currency                       ' The National Expected Retail Price (e.g. $1.99)
Private cRetail_PriceDiv As Scripting.Dictionary        ' The Regional Expected Retail Price (e.g. $1.99) (Entered by User)
Private cEstMultiplier As Single                        ' The National Multiplier (Avg??)
Private cEstMultiplierDiv As Scripting.Dictionary       ' The Regional Multipler (entered by USer)
Private cPrStrNoDiv As Scripting.Dictionary             'The Total Stores for the Prior Promotion Period by Region
''Private sngPCDiv As Scripting.Dictionary              ' Ration/Divisble for calcuation (Not used)
Private cStatusDiv As Scripting.Dictionary              ' The Status for the Region
Private cRegionDiv As Scripting.Dictionary              ' The Region
Private WoS As Long                                     'The Weeks of Sale
Private caCurRegStoNo As Scripting.Dictionary
Private sngPCDiv As Scripting.Dictionary
'Function cRecalculate(ByRef cVal As Variant, ByRef cDic As Scripting.Dictionary, ByVal CalcType As String) As Boolean
'Dim DivNo  As Long
'Dim tempVal
'    If CalcType = "USW" Then
'        tempVal = 0
'        For DivNo = 501 To 509
'            If DivNo = 508 Then DivNo = 509
'
'
'        Next
'    Else
'        tempVal = 0
'        For DivNo = 501 To 509
'            If DivNo = 508 Then DivNo = 509
'            tempVal = tempVal + cDic(DivNo)
'        Next
'    End If
'    cVal = tempVal
'End Function
Property Get pStatusDiv(ByVal DivNo As Long) As Long
    pStatusDiv = cStatusDiv(DivNo)
End Property
Property Get pdtPromoEndDt() As Date
     pdtPromoEndDt = cdtPromoEndDt
End Property
Property Get pdtFromDt() As Date
    pdtFromDt = cdtFromDt
End Property
Property Get pdtToDt() As Date
     pdtToDt = cdtToDt
End Property
Property Get plPcode() As Long
    plPcode = clPCode
End Property
Property Get plPDID() As Long
    plPDID = clPDID
End Property
Property Get pUnit_Cost() As Currency
    pUnit_Cost = cUnit_Cost
End Property
Property Get pUnit_CostDiv(ByVal DivNo As Long) As Currency
    pUnit_CostDiv = cUnit_CostDiv(DivNo)
End Property
Property Get pSupplier_Cost_Support() As Currency
    pSupplier_Cost_Support = cSupplier_Cost_Support
End Property
Property Get pSupplier_Cost_SupportDiv(ByVal DivNo As Long) As Currency
    pSupplier_Cost_SupportDiv = cSupplier_Cost_SupportDiv(DivNo)
End Property
Property Get pPrior_Sales() As Long
    pPrior_Sales = cPrior_Sales
End Property
Property Get pExp_Sales() As Long
    pExp_Sales = cOrigEstSales
End Property
Property Get pWoS() As Long
    pWoS = WoS
End Property
Property Get pOrigCalcSales() As Long
    pOrigCalcSales = cOrigCalcSales
End Property
Property Get pCalcSalesDiv(ByVal DivNo As Long) As Long
    pCalcSalesDiv = cCalcSalesDiv(DivNo)
End Property
Property Get pUPSPW() As Long
    pUPSPW = cUPSPW
End Property
''Property Get pFill_Qty() As Long
''    pFill_Qty = cFill_Qty
''End Property
Property Get pFill_QtyDiv(ByVal DivNo As Long) As Long
    pFill_QtyDiv = cFill_QtyDiv(DivNo)
End Property
Property Get pCurr_Retail_Price() As Currency
    pCurr_Retail_Price = cCurr_Retail_Price
End Property
Property Get pCurr_Retail_PriceDiv(ByVal DivNo As Long) As Currency
    pCurr_Retail_PriceDiv = cCurr_Retail_PriceDiv(DivNo)
End Property
Property Get pRetail_Price() As Currency
    pRetail_Price = cRetail_Price
End Property
Property Get pRetail_PriceDiv(ByVal DivNo As Long) As Currency
    pRetail_PriceDiv = cRetail_PriceDiv(DivNo)
End Property
Property Get pEstMultiplier() As Single
    pEstMultiplier = cEstMultiplier
End Property
Property Get pEstMultiplierDiv(ByVal DivNo As Long) As Single
    pEstMultiplierDiv = cEstMultiplierDiv(DivNo)
End Property
''Property Get plWoS() As Long
''    plWoS = clWoS
''End Property
Property Get pPrStrNoDiv(ByVal DivNo As Long) As Long
    pPrStrNoDiv = cPrStrNoDiv(DivNo)
End Property
Property Get pcUPSPWDiv(ByVal DivNo As Long) As Long
    pcUPSPWDiv = cUPSPWDiv(DivNo)
End Property
Property Get pcPrior_TotalRetailDiv(ByVal DivNo As Long) As Long
    pcPrior_TotalRetailDiv = cPrior_TotalRetailDiv(DivNo)
End Property
''Property Get psngPCDiv(ByVal DivNo As Long) As Long
''    psngPCDiv = sngPCDiv(DivNo)
''End Property
Property Get pcPrior_SalesDiv(ByVal DivNo As Long) As Long
    pcPrior_SalesDiv = cPrior_SalesDiv(DivNo)
End Property
Property Get pcEstSalesDiv(ByVal DivNo As Long) As Long
    pcEstSalesDiv = cEstSalesDiv(DivNo)
End Property
Property Get pcaCurRegStoNoDiv(ByVal DivNo As Long) As Long
    pcaCurRegStoNoDiv = caCurRegStoNo(DivNo)
End Property

Sub formulate(ByRef P_ID As Long)

Dim lIdx As Long, sSQL As String, aRegs() As String, sReg As String
Dim dtFromDt As Date, dtToDt As Date, lPcode As Long, lPDID As Long, lPVID As Long, bPassOK As Boolean ', sngMult As Single
Dim CN As ADODB.Connection, RS As ADODB.Recordset, RSI As ADODB.Recordset
Dim WoSChange As Boolean
Dim a_Regs As String
Dim DivNo As Long, a As Long



    a_Regs = "Min,Der,Stp,Pre,Dan,Bre,Rgy,xxx,Jkt"
    'On Error GoTo Err_Routine
    
    clPDID = P_ID
    CBA_ErrTag = ""
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    Set RSI = New ADODB.Recordset
    ' Split the Regions
    aRegs = Split(a_Regs, ",")
    
    If clPDID < CBA_LongHiVal Then
        If CBA_AST_frm_Products.cboWeeksOfSale <> Trim(Val(Mid(CBA_AST_frm_Products.cboWeeksOfSale.Tag, InStr(1, CBA_AST_frm_Products.cboWeeksOfSale.Tag, "~") + 1, 9))) _
            Or CDate(Trim(Left(Mid(CBA_AST_frm_Products.txtOnSaleDate.Tag, InStr(1, CBA_AST_frm_Products.txtOnSaleDate.Tag, " ") + 1, 19), Len(Mid(CBA_AST_frm_Products.txtOnSaleDate.Tag, InStr(1, CBA_AST_frm_Products.txtOnSaleDate.Tag, " ") + 1, 19))))) <> CDate(g_FixDate(CBA_AST_frm_Products.txtOnSaleDate)) Then WoSChange = True
    End If
    If clPDID = CBA_LongHiVal Or WoSChange = True Then
        WoS = CBA_AST_frm_Products.cboWeeksOfSale
        cdtPromoEndDt = g_FixDate(CBA_AST_frm_Products.txtEndDate)
        cdtFromDt = DateAdd("m", -12, g_FixDate(CBA_AST_frm_Products.txtOnSaleDate))
        cdtToDt = DateAdd("m", -12, cdtPromoEndDt)
        clPCode = CBA_AST_frm_Products.txtProductCode
        cUnit_Cost = NZ(CBA_AST_frm_Products.txtUnitCost, 0)
        cSupplier_Cost_Support = NZ(CBA_AST_frm_Products.txtSupplierCostSupport, 0)
        cPrior_Sales = NZ(CBA_AST_frm_Products.txtPriorSales, 0)
        cOrigEstSales = NZ(CBA_AST_frm_Products.txtExpectedSales, 0)
        cOrigCalcSales = NZ(CBA_AST_frm_Products.txtCalculatedSales, 0)
        cUPSPW = NZ(CBA_AST_frm_Products.txtUPSPW, 0)
        cFill_Qty = NZ(0, 0)
        cCurr_Retail_Price = NZ(CBA_AST_frm_Products.txtCurrRetailPrice, 0)
        cRetail_Price = NZ(CBA_AST_frm_Products.txtRetailPrice, 0)
        cEstMultiplier = NZ(CBA_AST_frm_Products.txtSalesMultiplier, 0)
        clWoS = CBA_AST_frm_Products.cboWeeksOfSale
        CBA_ErrTag = "Data"
        CBA_SQL_Queries.CBA_GenPullSQL "CBA_AST_POSbDiv", cdtFromDt, cdtToDt, clPCode
        Set cPrior_SalesDiv = New Scripting.Dictionary
        Set cPrior_TotalRetailDiv = New Scripting.Dictionary
        If CBA_CBISarr(0, 0) = 0 Then
            bPassOK = False
            For DivNo = 501 To 509
                If DivNo = 508 Then DivNo = 509
                cPrior_SalesDiv.Add DivNo, 0
                cPrior_TotalRetailDiv.Add DivNo, 0
            Next
        Else
            bPassOK = True
            For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                cPrior_SalesDiv.Add CLng(CBA_CBISarr(0, a)), CLng(CBA_CBISarr(1, a))
                cPrior_TotalRetailDiv.Add CLng(CBA_CBISarr(0, a)), CBA_CBISarr(2, a)
            Next
        End If
        CBA_SQL_Queries.CBA_GenPullSQL "CBA_AST_StoreNos", cdtToDt, cdtToDt
        Set cPrStrNoDiv = New Scripting.Dictionary
        If CBA_CBISarr(0, 0) = 0 Then
            bPassOK = False
            For DivNo = 501 To 509
                If DivNo = 508 Then DivNo = 509
                cPrStrNoDiv.Add DivNo, 0
            Next
        Else
            bPassOK = True
            For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                cPrStrNoDiv.Add CLng(CBA_CBISarr(0, a)), CLng(CBA_CBISarr(1, a))
            Next
        End If
        
        CBA_SQL_Queries.CBA_GenPullSQL "CBA_AST_StoreNos", cdtPromoEndDt, cdtPromoEndDt
        Set caCurRegStoNo = New Scripting.Dictionary
        If CBA_CBISarr(0, 0) = 0 Then
            bPassOK = False
            For DivNo = 501 To 509
                If DivNo = 508 Then DivNo = 509
                caCurRegStoNo.Add DivNo, 0
            Next
        Else
            bPassOK = True
            For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                caCurRegStoNo.Add CLng(CBA_CBISarr(0, a)), CLng(CBA_CBISarr(1, a))
            Next
        End If
        CBA_SQL_Queries.CBA_GenPullSQL "CBA_AST_CurrentPrice", , , clPCode
        Set cRetail_PriceDiv = New Scripting.Dictionary
        Set cCurr_Retail_PriceDiv = New Scripting.Dictionary
        If CBA_CBISarr(0, 0) = 0 Then
            bPassOK = False
            For DivNo = 501 To 509
                If DivNo = 508 Then DivNo = 509
                cRetail_PriceDiv.Add DivNo, 0
                cCurr_Retail_PriceDiv.Add DivNo, 0
            Next
        Else
            bPassOK = True
            For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                cRetail_PriceDiv.Add CLng(CBA_CBISarr(0, a)), CCur(CBA_CBISarr(1, a))
                cCurr_Retail_PriceDiv.Add CLng(CBA_CBISarr(0, a)), CCur(CBA_CBISarr(1, a))
            Next
        End If
        Set cUPSPWDiv = New Scripting.Dictionary
        Set sngPCDiv = New Scripting.Dictionary
        Set cEstSalesDiv = New Scripting.Dictionary
        Set cCalcSalesDiv = New Scripting.Dictionary
        Set cUnit_CostDiv = New Scripting.Dictionary
        Set cSupplier_Cost_SupportDiv = New Scripting.Dictionary
        Set cFill_QtyDiv = New Scripting.Dictionary
        Set cEstMultiplierDiv = New Scripting.Dictionary
        Set cStatusDiv = New Scripting.Dictionary
        Set cRegionDiv = New Scripting.Dictionary
        If bPassOK = True Then
            For lIdx = 501 To 509
                If lIdx <> 508 Then
                    cUnit_CostDiv.Add lIdx, cUnit_Cost
                    cSupplier_Cost_SupportDiv.Add lIdx, cSupplier_Cost_Support
                    cFill_QtyDiv.Add lIdx, cFill_Qty
                    cEstMultiplierDiv.Add lIdx, cEstMultiplier
                    cStatusDiv.Add lIdx, 1
                    cRegionDiv.Add lIdx, lIdx
                    If cPrStrNoDiv.Exists(lIdx) And cPrior_SalesDiv.Exists(lIdx) And NZ(clWoS, 0) <> 0 Then
                        cUPSPWDiv.Add lIdx, NZ(cPrior_SalesDiv(lIdx) / cPrStrNoDiv(lIdx) / clWoS, 0)
                        sngPCDiv.Add lIdx, 0
                        cEstSalesDiv.Add lIdx, (cPrior_SalesDiv(lIdx) / cPrStrNoDiv(lIdx)) * caCurRegStoNo(lIdx)
                        cCalcSalesDiv.Add lIdx, (cPrior_SalesDiv(lIdx) / cPrStrNoDiv(lIdx)) * caCurRegStoNo(lIdx)
                    End If
                End If
            Next
        End If
    Else
        WoS = CBA_AST_frm_Products.cboWeeksOfSale
        cdtPromoEndDt = g_FixDate(CBA_AST_frm_Products.txtEndDate)
        cdtFromDt = DateAdd("m", -12, g_FixDate(CBA_AST_frm_Products.txtOnSaleDate))
        cdtToDt = DateAdd("m", -12, cdtPromoEndDt)
        clPCode = CBA_AST_frm_Products.txtProductCode
        cUnit_Cost = NZ(CBA_AST_frm_Products.txtUnitCost, 0)
        cSupplier_Cost_Support = NZ(CBA_AST_frm_Products.txtSupplierCostSupport, 0)
        cPrior_Sales = NZ(CBA_AST_frm_Products.txtPriorSales, 0)
        cOrigEstSales = NZ(CBA_AST_frm_Products.txtExpectedSales, 0)
        cOrigCalcSales = NZ(CBA_AST_frm_Products.txtCalculatedSales, 0)
        cUPSPW = NZ(CBA_AST_frm_Products.txtUPSPW, 0)
        cFill_Qty = NZ(0, 0)
        cCurr_Retail_Price = NZ(CBA_AST_frm_Products.txtCurrRetailPrice, 0)
        cRetail_Price = NZ(CBA_AST_frm_Products.txtRetailPrice, 0)
        cEstMultiplier = NZ(CBA_AST_frm_Products.txtSalesMultiplier, 0)
        clWoS = CBA_AST_frm_Products.cboWeeksOfSale
        CBA_ErrTag = "Data"
        CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";"
        CBA_ErrTag = "SQL"
        sSQL = "SELECT * FROM L3_ProductRegions WHERE PV_PD_ID=" & P_ID
        RSI.Open sSQL, CN
        'RSI.Close
        Set cUPSPWDiv = New Scripting.Dictionary
        Set cEstSalesDiv = New Scripting.Dictionary
        Set cCalcSalesDiv = New Scripting.Dictionary
        Set cUnit_CostDiv = New Scripting.Dictionary
        Set cSupplier_Cost_SupportDiv = New Scripting.Dictionary
        Set cFill_QtyDiv = New Scripting.Dictionary
        Set cEstMultiplierDiv = New Scripting.Dictionary
        Set cStatusDiv = New Scripting.Dictionary
        Set cRegionDiv = New Scripting.Dictionary
        Set cPrior_SalesDiv = New Scripting.Dictionary
        Set cPrior_TotalRetailDiv = New Scripting.Dictionary
        Set cRetail_PriceDiv = New Scripting.Dictionary
        Set cCurr_Retail_PriceDiv = New Scripting.Dictionary
        Do While Not RSI.EOF
            'POPULATE THE CLASS MODULE VALUES FROM THE DATABASE
            CBA_ErrTag = "Data"
            cPrior_SalesDiv.Add CLng(NZ(RSI!PV_Region, 0)), CLng(NZ(RSI!PV_Prior_Sales, 0))
            cPrior_TotalRetailDiv.Add CLng(NZ(RSI!PV_Region, 0)), 0
            cRetail_PriceDiv.Add CLng(NZ(RSI!PV_Region, 0)), CCur(NZ(RSI!PV_Retail_Price, 0))
            cCurr_Retail_PriceDiv.Add CLng(NZ(RSI!PV_Region, 0)), CCur(NZ(RSI!PV_Curr_Retail_Price, 0))
            cUnit_CostDiv.Add CLng(NZ(RSI!PV_Region, 0)), CCur(NZ(RSI!PV_Unit_Cost, 0))
            cSupplier_Cost_SupportDiv.Add CLng(NZ(RSI!PV_Region, 0)), CCur(NZ(RSI!PV_Supplier_Cost_Support, 0))
            cFill_QtyDiv.Add CLng(NZ(RSI!PV_Region, 0)), CLng(NZ(RSI!PV_Fill_Qty, 0))
            cEstMultiplierDiv.Add CLng(NZ(RSI!PV_Region, 0)), CSng(NZ(RSI!PV_OrigEstMultiplier, 0))
            cStatusDiv.Add CLng(NZ(RSI!PV_Region, 0)), CLng(NZ(RSI!PV_Status, 0))
            cRegionDiv.Add CLng(NZ(RSI!PV_Region, 0)), CLng(NZ(RSI!PV_Region, 0))
            cUPSPWDiv.Add CLng(NZ(RSI!PV_Region, 0)), CLng(NZ(RSI!PV_UPSPW, 0))
            cEstSalesDiv.Add CLng(NZ(RSI!PV_Region, 0)), CLng(NZ(RSI!PV_OrigEstSales, 0))
            cCalcSalesDiv.Add CLng(NZ(RSI!PV_Region, 0)), CLng(NZ(RSI!PV_OrigCalcSales, 0))
            RSI.MoveNext
        Loop
        CBA_SQL_Queries.CBA_GenPullSQL "CBA_AST_StoreNos", cdtPromoEndDt, cdtPromoEndDt
        Set caCurRegStoNo = New Scripting.Dictionary
        If CBA_CBISarr(0, 0) = 0 Then
            bPassOK = False
            For DivNo = 501 To 509
                If DivNo = 508 Then DivNo = 509
                caCurRegStoNo.Add DivNo, 0
            Next
        Else
            bPassOK = True
            For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                caCurRegStoNo.Add CLng(CBA_CBISarr(0, a)), CLng(CBA_CBISarr(1, a))
            Next
        End If
    End If

End Sub
Function SetpStatusDiv(ByVal ValToSet As Long, ByVal DivNo As Long)
    cStatusDiv(DivNo) = ValToSet
End Function
'Function SetpdtPromoEndDt(ByVal ValToSet As Date)
'     cdtPromoEndDt = ValToSet
'End Property
'Function SetpdtFromDt(ByVal ValToSet As Date)
'    cdtFromDt = ValToSet
'End Property
'Function SetpdtToDt(ByVal ValToSet As Date)
'     cdtToDt = ValToSet
'End Property
'Function SetplPCode(ByVal ValToSet As Long)
'    clPCode = ValToSet
'End Property
'Function SetplPDID(ByVal ValToSet)
'    clPDID = ValToSet
'End Property
Function SetpUnit_Cost(ByVal ValToSet As Currency)
    cUnit_Cost = ValToSet
End Function
Function SetpUnit_CostDiv(ByVal ValToSet As Currency, ByVal DivNo As Long)
    cUnit_CostDiv(DivNo) = ValToSet
End Function
Function SetpSupplier_Cost_Support(ByVal ValToSet As Currency)
    cSupplier_Cost_Support = ValToSet
End Function
Function SetpSupplier_Cost_SupportDiv(ByVal ValToSet As Currency, ByVal DivNo As Long)
    cSupplier_Cost_SupportDiv(DivNo) = ValToSet
End Function
''Function SetpPrior_Sales(ByVal ValToSet As Long)
''    pPrior_Sales = cPrior_Sales
''End Function
''Function SetpOrigEstSales(ByVal ValToSet As Long)
''    cOrigEstSales = ValToSet
''End Function
Function SetpOrigCalcSales(ByVal ValToSet As Long)
    cOrigCalcSales = ValToSet
End Function
Function SetpCalcSalesDiv(ByVal ValToSet As Long, ByVal DivNo As Long)
    cCalcSalesDiv(DivNo) = ValToSet
End Function
Function SetpUPSPW(ByVal ValToSet As Long)
    cUPSPW = ValToSet
End Function
Function SetpFill_Qty(ByVal ValToSet As Long)
    cFill_Qty = ValToSet
End Function
Function SetpFill_QtyDiv(ByVal ValToSet As Long, ByVal DivNo As Long)
    cFill_QtyDiv(DivNo) = ValToSet
End Function
''Function SetpCurr_Retail_Price(ByVal ValToSet As Currency)
''    cCurr_Retail_Price = ValToSet
''End Function
Function SetpCurr_Retail_PriceDiv(ByVal ValToSet As Currency, ByVal DivNo As Long)
    cCurr_Retail_PriceDiv(DivNo) = ValToSet
    
End Function
Function SetpRetail_Price(ByVal ValToSet As Currency)
    cRetail_Price = ValToSet
End Function
Function SetpRetail_PriceDiv(ByVal ValToSet As Currency, ByVal DivNo As Long)
    cRetail_PriceDiv(DivNo) = ValToSet
End Function
Function SetpEstMultiplier(ByVal ValToSet As Single)
    cEstMultiplier = ValToSet
End Function
Function SetpEstMultiplierDiv(ByVal ValToSet As Single, ByVal DivNo As Long)
    cEstMultiplierDiv(DivNo) = ValToSet
End Function
''Function SetplWoS(ByVal ValToSet As Long)
''    clWoS = ValToSet
''End Function
'Function SetpPrStrNoDiv(ByVal ValToSet As Long, ByVal Divno As Long)
'    cPrStrNoDiv(Divno) = ValToSet
'End Property
Function SetpcUPSPWDiv(ByVal ValToSet As Long, ByVal DivNo As Long)
    cUPSPWDiv(DivNo) = ValToSet
End Function
Function SetpcPrior_TotalRetailDiv(ByVal ValToSet As Currency, ByVal DivNo As Long)
    cPrior_TotalRetailDiv(DivNo) = ValToSet
End Function
'Function SetpsngPCDiv(ByVal ValToSet As Long, ByVal Divno As Long)
'    sngPCDiv(Divno) = ValToSet
'End Property
Function SetpcPrior_SalesDiv(ByVal ValToSet As Long, ByVal DivNo As Long)
    cPrior_SalesDiv(DivNo) = ValToSet
End Function
Function SetpcEstSalesDiv(ByVal ValToSet As Long, ByVal DivNo As Long)
    cEstSalesDiv(DivNo) = ValToSet
End Function
