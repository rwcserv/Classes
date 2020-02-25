VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCBA_Prod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit                                 ' cCBA_Prod
Private Const DebugProdGroupObjects As Boolean = False
Private Type t_ContD
    Region  As Long
    ContractNo  As Long
    DeliveryFrom As Date
    DeliveryTo As Date
    CaseHeight As Double
    CaseLength As Double
    CaseWidth As Double
    CountryOfOrigin As String
    PayTermDays As Long
    OSD As Date
    ReturnsAllowed As Boolean
    LeadTimeDelivery As Long
    ValidInDiv As Boolean
    RCVTempFrom As Long
    RCVTempTo As Long
    TemperatureType As Long
    ExpirationTime As String
    SupplierNumber As Long
    SupplierName As String
    TaxID As Byte
End Type
Const NoOfVarsInContD As Long = 20
Private Type t_BrkD
    Region As Long
    ContractNo  As Long
    ValidFrom As Date
    ValidTo As Date
    ExWorks As Double
    FOB As Double
    DAT As Double
    DDP As Double
    Freight As Double
    Currency As String
End Type
Const NoOfVarsInBrkD As Long = 10
Private Type t_RetD
    Region As Long
    ValidFrom As Date
    ValidTo As Date
    Price As Double
End Type
Const NoOfVarsInRetD As Long = 4
Private Type t_IncoD
    Region As Long
    ContractNo As Long
    ValidFrom As Date
    ValidTo As Date
    SendToDiv As Boolean
    IncoTerm As String
End Type
Const NoOfVarsInIncoD As Long = 6
''Private Enum DataCont
''    ContractData = 0
''    BraketData = 1
''    IncoData = 2
''    Pricedata = 3
''End Enum
''Public Enum RetailorQTY
''    Retail = 0
''    QTY = 1
''End Enum
''Public Enum e_InvMDALL
''    eInventoryDifference = 0
''    eMarkdowns = 1
''    eBoth = 2
''End Enum
''Public Enum e_POSUSWTypes
''    eNotUSW = 0
''    eUSWisActive = 1
''    eUSWALL = 2
''    eProductLevel = 3
''    eUSWCNT = 4
''End Enum
Private plPcode As Long, plPClass As Long, plPksize As Long, plCon_PCode As Long, plEmpNo As Long, plTaxID As Long
Private plIsRW As Long, plLegCG As Long, plLegSCG As Long, plACatNum As Long, plACGnum As Long
Private plASCGNum As Long, plSeaID As Long, psBldCode As String, psProdDesc As String, psContent As String
Private psPCrdBnd As String, psPCrdDesc As String, psLegCGDesc As String, psLegSCGDesc As String
Private psLegCGmap As String, psACat As String, psACGDesc As String, psASCGDesc As String
Private a_SalesData() As Variant, a_RCVData() As Variant
Private pcSales As cCBA_ProdPeriodSalesData
Private pcUSW As cCBA_ProdPeriodSalesData
Private pc14DayST As cCBA_ProdPeriodSalesData
Private pcSalesDiv As cCBA_ProdPeriodSalesData
Private pcUSWDiv As cCBA_ProdPeriodSalesData
Private a_ContractData() As t_ContD
Private a_BraketData() As t_BrkD
Private a_IncoData() As t_IncoD
Private a_PriceData() As t_RetD
Private pcWastage As cCBA_WastageData
Private psdTotalBus As Scripting.Dictionary
Private psdPLabels As Scripting.Dictionary
Private pcolComrade As Collection
Private pbContractDataAvailable As Boolean, pbBraketDataAvailable As Boolean, pbPriceDataAvailable As Boolean, pbIncoDataAvailable As Boolean
Private pbGetsLetsOnly As Boolean
Private psScanDataMSeg As String
Private psManualMSeg(1 To 3) As String
Private plTrademarkNo As Long
Private plDoc_ID As Long

Function Build(ByVal ProdCode As Long, ByRef BuildCode As String, Optional ByRef TotBus As Scripting.Dictionary, _
        Optional ByRef ProdRS As ADODB.Recordset, Optional ByRef RCost As Variant, _
        Optional ByRef RDivM As Scripting.Dictionary, Optional ByRef NDB As Long, Optional ByRef RDIVFROM As Long, Optional ByRef RDIVTO As Long, _
        Optional ByRef Comrade As Collection, Optional ByVal ProdLabels As String, Optional ByRef ContData As Variant, _
        Optional ByRef ContIndex As Variant, Optional ByRef WasteData As cCBA_WastageData, _
        Optional ByRef ProdToMSegarr As Variant) As Boolean
        ', Optional ByRef SalesDic As Scripting.Dictionary
        'Optional ByRef USWDic As Scripting.Dictionary, Optional ByRef ProdToMSegarr As Variant, Optional ByRef SalesDIVDic As Scripting.Dictionary, _
        'Optional ByRef USWDIVDic As Scripting.Dictionary, Optional ByRef SellThrough14DaysDic As Scripting.Dictionary) As Boolean

    Dim RS As ADODB.Recordset
    Dim strSQL As String
    Dim a As Long, b As Long, Scnt As Long, Rcnt As Long, divcol As Long, div As Long, r As Long, c As Long
    Dim tim As Double
    Dim PLabelArr As Variant, P As Variant
    Dim COM As CBA_COM_COMMatch
    Dim Wcol As Collection
    Dim timtop As Double
'''     BUILDCODE OPTIONS
'''     B = Base Product Data (Description, Packsize, CG, SCG, ACAT, ACG, ASCG, BD, GBD) ' Mandatory
'''     c = Contract Data () MUST BE LITTLE c
'''     R = Receiving Data
'''     P = POS Data
'''     M = POS Margin (Cost) Data
'''     T = Total Business Metrics (Share)
'''     Cm = COMRADE data
'''     F = Forecasting Data
'''     L = Product Label Dictionary
'''     W = Wastage
'''     As = Alcohol Store Numbers included (automatically added)
'''     Ms = Market Segment Allocation
''      Ur = UnitsPerStore by Region (LARGE DATA SET) - Consider running product by product

'    On Error GoTo Err_Routine
'    CBA_Error = ""
    tim = Timer
    timtop = Timer
    If BuildCode = "GetsLetsOnly" Then bGetsLetsOnly = True: Exit Function
    lPcode = ProdCode
    sBldCode = BuildCode
    If InStr(1, BuildCode, "B") > 0 Then
        sProdDesc = NZ(ProdRS(1), 0): lPClass = NZ(ProdRS(2), 0): lPackSize = NZ(ProdRS(4), 0)
        plCon_PCode = NZ(ProdRS(5), 0): lEmpNo = NZ(ProdRS(7), 0): lIsRW = NZ(ProdRS(8), 0)
        lLegCG = NZ(ProdRS(11), 0): lLegSCG = NZ(ProdRS(13), 0): lACatNum = NZ(ProdRS(16), 0)
        lACGNum = NZ(ProdRS(18), 0): lASCGNum = NZ(ProdRS(20), 0): lSeaID = NZ(ProdRS(3), 0)
        sContent = NZ(ProdRS(6), ""): sPCrdBnd = NZ(ProdRS(9), ""): sPCrdDesc = NZ(ProdRS(10), "")
        sLegCGDesc = NZ(ProdRS(12), ""): sLegSCGDesc = NZ(ProdRS(14), ""): sLegCGMap = NZ(ProdRS(15), "")
        sACat = NZ(ProdRS(17), ""): sACGDesc = NZ(ProdRS(19), ""): sASCGDesc = NZ(ProdRS(21), "")
        lTaxID = NZ(ProdRS(23), ""): lTrademarkNo = NZ(ProdRS(24), 0)
    End If
    If DebugProdGroupObjects = True Then Debug.Print "Base Complete in: " & Timer - tim
    'tim = Timer
    If InStr(1, BuildCode, "L") > 0 And ProdLabels <> "" Then
        Set psdPLabels = New Scripting.Dictionary
        PLabelArr = Split(CStr(ProdLabels), ",")
        For a = LBound(PLabelArr) To UBound(PLabelArr)
            psdPLabels.Add PLabelArr(a), PLabelArr(a)
        Next
    End If
    If DebugProdGroupObjects = True Then Debug.Print "Labels Complete in: " & Timer - tim
    'tim = Timer
'    If (InStr(1, BuildCode, "P") > 0 Or InStr(1, BuildCode, "R") > 0) Then
'        Set psdSales = SalesDic
'        Set psdUSW = USWDic
'        Set psd14DayST = SellThrough14DaysDic
'        Set psdUSWDiv = USWDIVDic
'        Set psdSalesDiv = SalesDIVDic
'    End If
'    If (InStr(1, BuildCode, "P") > 0 Or InStr(1, BuildCode, "R") > 0) And IsMissing(SalesArr) = False And IsMissing(SalesArrDIV) = False Then
'        sdSales = SalesDic
'        sdUSW = USWDic
'        sdSalesDiv = SalesDIVDic
'        sdUSWDiv = USWDIVDic
'        c14DayST = SellThrough14DaysDic
'
''        Set psdSalesDiv = New Scripting.Dictionary: divcol = 10
''        For div = 501 To 509
''            If div = 508 Then div = 509
''            Rcnt = -1
''            If SalesArrDIV(0, 0) <> 0 Then
''                For a = LBound(SalesArrDIV, 2) To UBound(SalesArrDIV, 2)
''                    If SalesArrDIV(divcol, a) = div Then Rcnt = Rcnt + 1
''                    If SalesArrDIV(divcol, a) > div Then Exit For
''                Next
''            End If
''            If Rcnt > -1 Then
''                ReDim arr(0 To UBound(SalesArrDIV, 1), 0 To Rcnt)
''                r = -1
''                For a = LBound(SalesArrDIV, 2) To UBound(SalesArrDIV, 2)
''                    If SalesArrDIV(divcol, a) > div Then Exit For
''                    If SalesArrDIV(divcol, a) = CStr(div) Then
''                        r = r + 1: c = -1
''                        For b = LBound(SalesArrDIV, 1) To UBound(SalesArrDIV, 1) - 1
''                            c = c + 1
''                            arr(c, r) = SalesArrDIV(b, a)
''                        Next
''                    End If
''                Next
''            Else
''                ReDim arr(0 To 0, 0 To 0): arr(0, 0) = 0
''            End If
''            psdSalesDiv.Add CStr(div), arr
''        Next
''        If DebugProdGroupObjects = True Then Debug.Print "Regional Sales Complete in: " & Timer - tim:        tim = Timer
''
''        Set psdUSWDiv = New Scripting.Dictionary: divcol = 13
''        For div = 501 To 509
''            If div = 508 Then div = 509
''            Rcnt = -1
''            If USWarrDIV(0, 0) <> 0 Then
''                For a = LBound(USWarrDIV, 2) To UBound(USWarrDIV, 2)
''                    If USWarrDIV(divcol, a) = div Then Rcnt = Rcnt + 1
''                    If USWarrDIV(divcol, a) > div Then Exit For
''                Next
''            End If
''            If Rcnt > -1 Then
''                ReDim arr(0 To UBound(USWarrDIV, 1), 0 To Rcnt)
''                r = -1
''                For a = LBound(USWarrDIV, 2) To UBound(USWarrDIV, 2)
''                    If USWarrDIV(divcol, a) > div Then Exit For
''                    If USWarrDIV(divcol, a) = div Then
''                        r = r + 1: c = -1
''                        For b = LBound(USWarrDIV, 1) To UBound(USWarrDIV, 1) - 1
''                            c = c + 1
''                            arr(c, r) = USWarrDIV(b, a)
''                        Next
''                    End If
''                Next
''            Else
''                ReDim arr(0 To 0, 0 To 0): arr(0, 0) = 0
''            End If
''            psdUSWDiv.Add CStr(div), arr
''        Next
''        If DebugProdGroupObjects = True Then Debug.Print "Regional USW Complete in: " & Timer - tim:        tim = Timer
'
'    End If

    If InStr(1, BuildCode, "T") > 0 Then
        Set psdTotalBus = TotBus
        If DebugProdGroupObjects = True Then Debug.Print "Totbus Complete in: " & Timer - tim ': tim = Timer
    End If
    If InStr(1, BuildCode, "Cm") > 0 Then
        Set pcolComrade = Comrade
        If DebugProdGroupObjects = True Then Debug.Print "Comrade Complete in: " & Timer - tim ': tim = Timer
    End If
    
'    b = 0
'    For a = LBound(ContData(0), 2) To UBound(ContData(0), 2)
'        If ContData(0)(2, a) = 3946 Then
'            b = b + 1
'        End If
'    Next
    
    
    
    If InStr(1, BuildCode, "c") > 0 Then
'    If InStr(1, BuildCode, "Cm") = InStr(1, BuildCode, "C") > 0 Then
        If ContIndex(e_DataCont.eContractData) Is Nothing Then
            pbContractDataAvailable = False
        ElseIf ContIndex(e_DataCont.eContractData).Exists(CStr(ProdCode)) = False Then
            pbContractDataAvailable = False
        Else
            pbContractDataAvailable = True
            ReDim a_ContractData(0 To ContIndex(e_DataCont.eContractData)(CStr(ProdCode)).Count - 1)
            a = -1
            For Each P In ContIndex(e_DataCont.eContractData)(CStr(ProdCode))
                a = a + 1
                a_ContractData(a).Region = NZ(ContData(e_DataCont.eContractData)(0, P), 0)
                a_ContractData(a).ContractNo = NZ(ContData(e_DataCont.eContractData)(1, P), 0)
                a_ContractData(a).DeliveryFrom = NZ(ContData(e_DataCont.eContractData)(3, P), 0)
                a_ContractData(a).DeliveryTo = NZ(ContData(e_DataCont.eContractData)(4, P), #12/31/2050#)
                a_ContractData(a).CaseHeight = NZ(ContData(e_DataCont.eContractData)(5, P), 0)
                a_ContractData(a).CaseLength = NZ(ContData(e_DataCont.eContractData)(6, P), 0)
                a_ContractData(a).CaseWidth = NZ(ContData(e_DataCont.eContractData)(7, P), 0)
                a_ContractData(a).CountryOfOrigin = NZ(ContData(e_DataCont.eContractData)(8, P), "")
                a_ContractData(a).PayTermDays = NZ(ContData(e_DataCont.eContractData)(9, P), 0)
                a_ContractData(a).OSD = NZ(ContData(e_DataCont.eContractData)(10, P), 0)
                a_ContractData(a).ReturnsAllowed = NZ(ContData(e_DataCont.eContractData)(11, P), False)
                a_ContractData(a).LeadTimeDelivery = NZ(ContData(e_DataCont.eContractData)(12, P), 0)
                a_ContractData(a).ValidInDiv = NZ(ContData(e_DataCont.eContractData)(13, P), False)
                a_ContractData(a).RCVTempFrom = NZ(ContData(e_DataCont.eContractData)(14, P), 0)
                a_ContractData(a).RCVTempTo = NZ(ContData(e_DataCont.eContractData)(15, P), 0)
                a_ContractData(a).TemperatureType = NZ(ContData(e_DataCont.eContractData)(16, P), 0)
                a_ContractData(a).ExpirationTime = NZ(ContData(e_DataCont.eContractData)(17, P), "")
                a_ContractData(a).SupplierNumber = NZ(ContData(e_DataCont.eContractData)(18, P), "")
                a_ContractData(a).SupplierName = NZ(ContData(e_DataCont.eContractData)(19, P), "")
                a_ContractData(a).TaxID = NZ(ContData(e_DataCont.eContractData)(20, P), "")
                
            Next
        End If
        If ContIndex(e_DataCont.eBraketData) Is Nothing Then
            pbBraketDataAvailable = False
        ElseIf ContIndex(e_DataCont.eBraketData).Exists(CStr(ProdCode)) = False Then
            pbBraketDataAvailable = False
        Else
            pbBraketDataAvailable = True
            ReDim a_BraketData(0 To ContIndex(e_DataCont.eBraketData)(CStr(ProdCode)).Count - 1): a = -1
            For Each P In ContIndex(e_DataCont.eBraketData)(CStr(ProdCode))
                a = a + 1
                a_BraketData(a).Region = NZ(ContData(e_DataCont.eBraketData)(0, P), 0)
                a_BraketData(a).ContractNo = NZ(ContData(e_DataCont.eBraketData)(2, P), 0)
                a_BraketData(a).ValidFrom = NZ(ContData(e_DataCont.eBraketData)(4, P), 0)
                a_BraketData(a).ValidTo = NZ(ContData(e_DataCont.eBraketData)(5, P), #12/31/2050#)
                a_BraketData(a).ExWorks = NZ(ContData(e_DataCont.eBraketData)(6, P), 0)
                a_BraketData(a).FOB = NZ(ContData(e_DataCont.eBraketData)(7, P), 0)
                a_BraketData(a).DAT = NZ(ContData(e_DataCont.eBraketData)(8, P), 0)
                a_BraketData(a).DDP = NZ(ContData(e_DataCont.eBraketData)(10, P), 0)
                a_BraketData(a).Freight = NZ(ContData(e_DataCont.eBraketData)(11, P), 0)
                a_BraketData(a).Currency = NZ(ContData(e_DataCont.eBraketData)(9, P), "")
            Next
        End If
        If ContIndex(e_DataCont.eIncoData) Is Nothing Then
            pbIncoDataAvailable = False
        ElseIf ContIndex(e_DataCont.eIncoData).Exists(CStr(ProdCode)) = False Then
            pbIncoDataAvailable = False
        Else
            pbIncoDataAvailable = True
            ReDim a_IncoData(0 To ContIndex(e_DataCont.eIncoData)(CStr(ProdCode)).Count - 1): a = -1
            For Each P In ContIndex(e_DataCont.eIncoData)(CStr(ProdCode))
                a = a + 1
                a_IncoData(a).Region = NZ(ContData(e_DataCont.eIncoData)(0, P), 0)
                a_IncoData(a).ContractNo = NZ(ContData(e_DataCont.eIncoData)(2, P), 0)
                a_IncoData(a).ValidFrom = NZ(ContData(e_DataCont.eIncoData)(4, P), 0)
                a_IncoData(a).ValidTo = NZ(ContData(e_DataCont.eIncoData)(5, P), #12/31/2050#)
                a_IncoData(a).SendToDiv = NZ(ContData(e_DataCont.eIncoData)(6, P), False)
                a_IncoData(a).IncoTerm = NZ(ContData(e_DataCont.eIncoData)(7, P), "")
            Next
        End If
        If ContIndex(e_DataCont.ePricedata) Is Nothing Then
            pbPriceDataAvailable = False
        ElseIf ContIndex(e_DataCont.ePricedata).Exists(CStr(ProdCode)) = False Then
            pbPriceDataAvailable = False
        Else
            pbPriceDataAvailable = True
            ReDim a_PriceData(0 To ContIndex(e_DataCont.ePricedata)(CStr(ProdCode)).Count - 1): a = -1
            For Each P In ContIndex(e_DataCont.ePricedata)(CStr(ProdCode))
                a = a + 1
                a_PriceData(a).Region = NZ(ContData(e_DataCont.ePricedata)(0, P), 0)
                a_PriceData(a).ValidFrom = NZ(ContData(e_DataCont.ePricedata)(3, P), 0)
                a_PriceData(a).ValidTo = NZ(ContData(e_DataCont.ePricedata)(4, P), #12/31/2050#)
                a_PriceData(a).Price = NZ(ContData(e_DataCont.ePricedata)(2, P), 0)
            Next
        End If
        If DebugProdGroupObjects = True Then Debug.Print "Contract Complete in: " & Timer - tim
        'tim = Timer
    End If
    
    If InStr(1, BuildCode, "W") > 0 Then
        Set pcWastage = WasteData
'        If Not WasteIndex Is Nothing Then
'            If WasteIndex.Exists(CStr(ProdCode)) Then
'                ReDim a_WastageData(0 To 7, 0 To WasteIndex(CStr(ProdCode)).Count - 1): a = -1
'                Set Wcol = WasteIndex(CStr(ProdCode))
'                For Each P In Wcol
'                    'Wcol = WasteIndex(CStr(ProdCode))(P)
'                    a = a + 1
'                    For b = 0 To 7
'                        a_WastageData(b, a) = WasteData(b, P)
'                    Next
'                Next
'            Else
'                ReDim a_WastageData(0 To 7, 0 To 0)
'                a_WastageData(0, 0) = 0: a_WastageData(1, 0) = 0: a_WastageData(2, 0) = 0: a_WastageData(3, 0) = 0
'                a_WastageData(4, 0) = 0: a_WastageData(5, 0) = 0: a_WastageData(6, 0) = 0:  a_WastageData(7, 0) = 0
'            End If
'        End If
        If DebugProdGroupObjects = True Then Debug.Print "Wastage Complete in: " & Timer - tim ': tim = Timer
    End If


    If InStr(1, BuildCode, "Ms") > 0 And IsMissing(ProdToMSegarr) = False Then
        If ProdToMSegarr(0, 0) = 0 Then
        Else
            For a = LBound(ProdToMSegarr, 2) To UBound(ProdToMSegarr, 2)
                If ProdToMSegarr(4, a) > ProdCode Then Exit For
                If ProdToMSegarr(4, a) = ProdCode Then
                    If ProdToMSegarr(1, a) = 3 Then
                        Me.ManualMSeg 1, CStr(ProdToMSegarr(2, a))
                    ElseIf ProdToMSegarr(1, a) = 2 Then
                        Me.sScanDataMSeg = CStr(ProdToMSegarr(2, a))
                    End If
                End If
            Next
            If DebugProdGroupObjects = True Then Debug.Print "Market Segment Complete in: " & Timer - tim ': tim = Timer
        End If
    End If


'Exit_Routine:
    If DebugProdGroupObjects = True Then Debug.Print "Prod build Method Complete in: " & Timer - timtop

'    On Error Resume Next
'    Exit Function
'
'Err_Routine:
'    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_Prod.Build", 3)
'    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
'    Debug.Print CBA_Error
'    Stop            ' ^RW Camera + next line
'    Resume Next
'    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
'    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
'    GoTo Exit_Routine
'    Resume Next

End Function

Private Function ReturnData(ByVal DataSet As e_DataCont, ByVal Field As String, ByVal ArrayPosit As Long) As Variant

    On Error GoTo Err_Routine
    CBA_Error = ""
    If DataSet = e_DataCont.eContractData Then
        ''' ReturnCode Values '''
        ''' U = UniqueValue (Distinct)  ''' R = Region          ''' C = ContractNo      ''' Df = DeliveryFrom
        ''' Dt = Deliveryto             ''' H = CaseHeight      ''' L = CaseLength      ''' W = CaseWidth
        ''' Oo = CountryOfOrigin        ''' P = PayTermDays     ''' O = OSD             ''' Ra = ReturnsAllowed
        ''' Lt = LeadTimeDelivery       ''' V = ValidInDiv      ''' Tf = RCVTempFrom    ''' Tt = RCVTempTo
        ''' Ty = TemperatureType        ''' E = ExpirationTime  ''' Sn = SupplierName   ''' Tx = lTaxID
        ''' So = SupplierNumber
        If Field = "R" Then ReturnData = a_ContractData(ArrayPosit).Region: Exit Function
        If Field = "c" Then ReturnData = a_ContractData(ArrayPosit).ContractNo: Exit Function
        If Field = "Df" Then ReturnData = a_ContractData(ArrayPosit).DeliveryFrom: Exit Function
        If Field = "Dt" Then ReturnData = a_ContractData(ArrayPosit).DeliveryTo: Exit Function
        If Field = "H" Then ReturnData = a_ContractData(ArrayPosit).CaseHeight: Exit Function
        If Field = "L" Then ReturnData = a_ContractData(ArrayPosit).CaseLength: Exit Function
        If Field = "W" Then ReturnData = a_ContractData(ArrayPosit).CaseWidth: Exit Function
        If Field = "Oo" Then ReturnData = a_ContractData(ArrayPosit).CountryOfOrigin: Exit Function
        If Field = "P" Then ReturnData = a_ContractData(ArrayPosit).PayTermDays: Exit Function
        If Field = "O" Then ReturnData = a_ContractData(ArrayPosit).OSD: Exit Function
        If Field = "Ra" Then ReturnData = a_ContractData(ArrayPosit).ReturnsAllowed: Exit Function
        If Field = "Lt" Then ReturnData = a_ContractData(ArrayPosit).LeadTimeDelivery: Exit Function
        If Field = "V" Then ReturnData = a_ContractData(ArrayPosit).ValidInDiv: Exit Function
        If Field = "Tf" Then ReturnData = a_ContractData(ArrayPosit).RCVTempFrom: Exit Function
        If Field = "Tt" Then ReturnData = a_ContractData(ArrayPosit).RCVTempTo: Exit Function
        If Field = "Ty" Then ReturnData = a_ContractData(ArrayPosit).TemperatureType: Exit Function
        If Field = "E" Then ReturnData = a_ContractData(ArrayPosit).ExpirationTime: Exit Function
        If Field = "Sn" Then ReturnData = a_ContractData(ArrayPosit).SupplierName: Exit Function
        If Field = "So" Then ReturnData = a_ContractData(ArrayPosit).SupplierNumber: Exit Function
        If Field = "Tx" Then ReturnData = a_ContractData(ArrayPosit).TaxID: Exit Function
        
        
        
    ElseIf DataSet = e_DataCont.eBraketData Then
        ''' ReturnCode Values '''
        ''' U = UniqueValue (Distinct)  ''' R = Region          ''' C = ContractNo      ''' Vf = ValidFrom
        ''' Vt = ValidTo                ''' E = ExWorks         ''' F = FOB             ''' T = DAT
        ''' D = DDP                     ''' Fr = Freight         ''' C = Currency
        If Field = "R" Then ReturnData = a_BraketData(ArrayPosit).Region: Exit Function
        If Field = "c" Then ReturnData = a_BraketData(ArrayPosit).ContractNo: Exit Function
        If Field = "Vf" Then ReturnData = a_BraketData(ArrayPosit).ValidFrom: Exit Function
        If Field = "Vt" Then ReturnData = a_BraketData(ArrayPosit).ValidTo: Exit Function
        If Field = "E" Then ReturnData = a_BraketData(ArrayPosit).ExWorks: Exit Function
        If Field = "F" Then ReturnData = a_BraketData(ArrayPosit).FOB: Exit Function
        If Field = "T" Then ReturnData = a_BraketData(ArrayPosit).DAT: Exit Function
        If Field = "D" Then ReturnData = a_BraketData(ArrayPosit).DDP: Exit Function
        If Field = "Fr" Then ReturnData = a_BraketData(ArrayPosit).Freight: Exit Function
        If Field = "c" Then ReturnData = a_BraketData(ArrayPosit).Currency: Exit Function
    
    ElseIf DataSet = e_DataCont.eIncoData Then
        ''' ReturnCode Values '''
        ''' U = UniqueValue (Distinct)  ''' R = Region          ''' C = ContractNo      ''' Vf = ValidFrom
        ''' Vt = ValidTo                ''' S = SendToDiv       ''' I = IncoTerm
        If Field = "R" Then ReturnData = a_IncoData(ArrayPosit).Region: Exit Function
        If Field = "c" Then ReturnData = a_IncoData(ArrayPosit).ContractNo: Exit Function
        If Field = "Vf" Then ReturnData = a_IncoData(ArrayPosit).ValidFrom: Exit Function
        If Field = "Vt" Then ReturnData = a_IncoData(ArrayPosit).ValidTo: Exit Function
        If Field = "S" Then ReturnData = a_IncoData(ArrayPosit).SendToDiv: Exit Function
        If Field = "I" Then ReturnData = a_IncoData(ArrayPosit).IncoTerm: Exit Function
    
    ElseIf DataSet = e_DataCont.ePricedata Then
        ''' ReturnCode Values '''
        ''' U = UniqueValue (Distinct)  ''' R = Region          ''' Vf = ValidFrom      ''' Vt = ValidTo
        ''' P = Price
        If Field = "R" Then ReturnData = a_PriceData(ArrayPosit).Region: Exit Function
        If Field = "Vf" Then ReturnData = a_PriceData(ArrayPosit).ValidFrom: Exit Function
        If Field = "Vt" Then ReturnData = a_PriceData(ArrayPosit).ValidTo: Exit Function
        If Field = "P" Then ReturnData = a_PriceData(ArrayPosit).Price: Exit Function
    
    End If
    
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_Prod.ReturnData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Function TranslateReturnCode(ByVal ReturnCode As String) As Collection
    Dim ReturnCol As Collection
    Dim a As Long
    On Error GoTo Err_Routine
    CBA_Error = ""
    If Mid(ReturnCode, 1, 1) <> "U" And InStr(1, ReturnCode, "U") > 0 Then
            MsgBox "The U must be placed at the start of the ReturnCode when querying a unique value", vbOKOnly: ReDim arr(0 To 0, 0 To 0): arr(0, 0) = 0: Exit Function
    End If
    If Mid(ReturnCode, 1, 1) = "U" Then
        If Len(ReturnCode) > 3 Or (Len(ReturnCode) = 3 And UCase(Mid(ReturnCode, 3, 1)) = Mid(ReturnCode, 3, 1)) Then
            MsgBox "Only One field can be returned when querying a unique value", vbOKOnly: ReDim arr(0 To 0, 0 To 0): arr(0, 0) = 0: Exit Function
        End If
        Set ReturnCol = New Collection
        If Len(ReturnCode) = 3 Then
            ReturnCol.Add Mid(ReturnCode, 2, 2)
        Else
            ReturnCol.Add Mid(ReturnCode, 2, 1)
        End If
    Else
        Set ReturnCol = New Collection
        For a = 1 To Len(ReturnCode)
            If a + 1 > Len(ReturnCode) Then
                ReturnCol.Add Mid(ReturnCode, a, 1)
            Else
                If LCase(Mid(ReturnCode, a + 1, 1)) = Mid(ReturnCode, a + 1, 1) Then
                    ReturnCol.Add Mid(ReturnCode, a, 2)
                    a = a + 1
                Else
                    ReturnCol.Add Mid(ReturnCode, a, 1)
                End If
            End If
        Next
    End If
    Set TranslateReturnCode = ReturnCol
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_Prod.TranslateReturnCode", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Function getContractData(ByVal ReturnCode As String, Optional ByVal LatestFromDate As Date, _
    Optional ByVal EarliestToDate As Date, Optional ContractNo As Long, Optional div As Long = 599) As Variant
    If bGetsLetsOnly = True Then Exit Function
    getContractData = getDataContData(e_DataCont.eContractData, ReturnCode, LatestFromDate, EarliestToDate, ContractNo, div)
End Function
Function getPriceData(ByVal ReturnCode As String, Optional ByVal LatestValidFrom As Date, Optional ByVal EarliestValidTo As Date, Optional div As Long = 599) As Variant
    If bGetsLetsOnly = True Then Exit Function
    getPriceData = getDataContData(e_DataCont.ePricedata, ReturnCode, LatestValidFrom, EarliestValidTo, 0, div)
End Function
Function getIncoTermData(ByVal ReturnCode As String, Optional ByVal LatestFromDate As Date, _
    Optional ByVal EarliestToDate As Date, Optional ContractNo As Long, Optional div As Long = 599) As Variant
    If bGetsLetsOnly = True Then Exit Function
    getIncoTermData = getDataContData(e_DataCont.eIncoData, ReturnCode, LatestFromDate, EarliestToDate, ContractNo, div)
End Function
Function getBraketData(ByVal ReturnCode As String, Optional ByVal LatestFromDate As Date, _
    Optional ByVal EarliestToDate As Date, Optional ContractNo As Long, Optional div As Long = 599) As Variant
    If bGetsLetsOnly = True Then Exit Function
    getBraketData = getDataContData(e_DataCont.eBraketData, ReturnCode, LatestFromDate, EarliestToDate, ContractNo, div)
End Function
Private Function getDataContData(ByVal DC As e_DataCont, ByVal ReturnCode As String, Optional ByVal LatestFromDate As Date, _
    Optional ByVal EarliestToDate As Date, Optional ContractNo As Long, Optional div As Long = 599) As Variant
    Dim a As Long, b As Long, cnt As Long, Ccnt As Long
    Dim arr As Variant, v As Variant, fld As Variant, vl As Variant
    Dim col As Collection, ReturnCol As Collection
    Dim UDic As Scripting.Dictionary
    Dim UseThis As Boolean, UniqueQuery As Boolean
    On Error GoTo Err_Routine
    CBA_Error = ""
    
    If (DC = e_DataCont.eIncoData And pbIncoDataAvailable = False) Or (DC = e_DataCont.eBraketData And pbBraketDataAvailable = False) Or _
        (DC = e_DataCont.eContractData And pbContractDataAvailable = False) Or (DC = e_DataCont.ePricedata And pbPriceDataAvailable = False) Or _
            ReturnCode = "" Then
            If ReturnCode = "" Then Debug.Print "Return Code is Required"
            ReDim arr(0 To 0, 0 To 0): arr(0, 0) = 0: getDataContData = arr: Exit Function
    End If
    Set ReturnCol = TranslateReturnCode(ReturnCode)
    If ReturnCol.Count = -1 Then ReDim arr(0 To 0, 0 To 0): arr(0, 0) = 0: getDataContData = arr: Exit Function
    If Mid(ReturnCode, 1, 1) = "U" Then UniqueQuery = True
    
    
    If LatestFromDate = 0 And EarliestToDate = 0 And ContractNo = 0 And div = 599 Then
        Set col = New Collection
        Select Case DC
            Case e_DataCont.eBraketData
                For a = LBound(a_BraketData) To UBound(a_BraketData)
                    col.Add a
                Next
            Case e_DataCont.eContractData
                For a = LBound(a_ContractData) To UBound(a_ContractData)
                    If a_ContractData(a).ValidInDiv = True Then col.Add a
                Next
            Case e_DataCont.ePricedata
                For a = LBound(a_PriceData) To UBound(a_PriceData)
                    col.Add a
                Next
            Case e_DataCont.eIncoData
                For a = LBound(a_IncoData) To UBound(a_IncoData)
                    If a_IncoData(a).SendToDiv = True Then col.Add a
                Next
        End Select
    Else
        UseThis = True: Set col = New Collection
        Select Case DC
            Case e_DataCont.eBraketData
                For a = LBound(a_BraketData) To UBound(a_BraketData)
                    UseThis = True
                    If LatestFromDate <> 0 And LatestFromDate < a_BraketData(a).ValidFrom Then UseThis = False
                    If EarliestToDate <> 0 And EarliestToDate > a_BraketData(a).ValidTo Then UseThis = False
                    If ContractNo <> 0 And ContractNo <> a_BraketData(a).ContractNo Then UseThis = False
                    If div <> 599 And div <> a_BraketData(a).Region Then UseThis = False
                    If UseThis = True Then col.Add a
                Next
            Case e_DataCont.eContractData
                For a = LBound(a_ContractData) To UBound(a_ContractData)
                    UseThis = True
                    If LatestFromDate <> 0 And LatestFromDate < a_ContractData(a).DeliveryFrom Then UseThis = False
                    If EarliestToDate <> 0 And EarliestToDate > a_ContractData(a).DeliveryTo Then UseThis = False
                    If ContractNo <> 0 And ContractNo <> a_ContractData(a).ContractNo Then UseThis = False
                    If div <> 599 And div <> a_ContractData(a).Region Then UseThis = False
                    If UseThis = True Then col.Add a
                Next
            Case e_DataCont.ePricedata
                For a = LBound(a_PriceData) To UBound(a_PriceData)
                    UseThis = True
                    If LatestFromDate <> 0 And LatestFromDate < a_PriceData(a).ValidFrom Then UseThis = False
                    If EarliestToDate <> 0 And EarliestToDate > a_PriceData(a).ValidTo Then UseThis = False
                    If div <> 599 And div <> a_PriceData(a).Region Then UseThis = False
                    If UseThis = True Then col.Add a
                Next
            Case e_DataCont.eIncoData
                For a = LBound(a_IncoData) To UBound(a_IncoData)
                    If a_IncoData(a).SendToDiv = False Then UseThis = False Else UseThis = True
                    If LatestFromDate <> 0 And LatestFromDate < a_IncoData(a).ValidFrom Then UseThis = False
                    If EarliestToDate <> 0 And EarliestToDate > a_IncoData(a).ValidTo Then UseThis = False
                    If ContractNo <> 0 And ContractNo <> a_IncoData(a).ContractNo Then UseThis = False
                    If div <> 599 And div <> a_IncoData(a).Region Then UseThis = False
                    If UseThis = True Then col.Add a
                Next
        End Select
    End If
    If col.Count = 0 Then ReDim arr(0 To 0, 0 To 0): arr(0, 0) = 0: getDataContData = arr: Exit Function
    
    
    If UniqueQuery = True Then
        Set UDic = New Scripting.Dictionary
        For Each v In col
            For Each fld In ReturnCol
                vl = ReturnData(DC, fld, v)
                If UDic.Exists(CStr(vl)) = False Then UDic.Add CStr(vl), CStr(vl)
            Next
        Next
        ReDim arr(0 To ReturnCol.Count - 1, 0 To UDic.Count)
        cnt = -1
        For Each vl In UDic
            cnt = cnt + 1: arr(0, cnt) = vl
        Next
    Else
        ReDim arr(0 To ReturnCol.Count - 1, 0 To col.Count)
        cnt = -1
        For Each v In col
            cnt = cnt + 1: Ccnt = -1
            For Each fld In ReturnCol
                Ccnt = Ccnt + 1: arr(Ccnt, cnt) = ReturnData(DC, fld, v)
            Next
        Next
    End If
    getDataContData = arr
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_Prod.GetDataContData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Function getWastageData(ByVal isRetail As e_RetailorQTY, ByVal InventoryMarkdownBoth As e_InvMDALL, Optional ByVal DFrom As Date, Optional Dto As Date, Optional div As Long = 599) As Double '@TP NEED TO FIX UP
    
    If bGetsLetsOnly = True Then Exit Function
    If DFrom = 0 Or Dto = 0 Then
        'need to handle
        Exit Function
    End If
    If div = 0 Then div = 599
    If cWastage Is Nothing Then Exit Function
    getWastageData = cWastage.getWastageData(isRetail, InventoryMarkdownBoth, DFrom, Dto, div)
    
End Function
Function getTotBusData(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal QueryType As String) As Double
    Dim mth As Variant, yr As Variant
    Dim s As Double
    If bGetsLetsOnly = True Then Exit Function
    If psdTotalBus Is Nothing Then getTotBusData = 0: Exit Function
    s = 0
    For yr = Year(DateFrom) To Year(DateTo)
        For mth = 1 To 12
            If DateFrom <= DateSerial(yr, mth, 1) And DateTo >= DateSerial(yr, mth, 1) Then
                s = s + psdTotalBus(CStr(yr))(CStr(mth))(QueryType)
            End If
        Next
    Next
    getTotBusData = s
End Function
Function getRCVShare(ByVal DateFrom As Date, ByVal DateTo As Date, Optional ByVal RetailnotQTY As Boolean) As Double
    If bGetsLetsOnly = True Then Exit Function
    If RetailnotQTY = True Then
        If getTotBusData(DateFrom, DateTo, "RCVRetail") = 0 Then getRCVShare = 0 Else getRCVShare = getRCVdata(DateFrom, DateTo, "Retail") / getTotBusData(DateFrom, DateTo, "RCVRetail")
    Else
        If getTotBusData(DateFrom, DateTo, "RCVQTY") = 0 Then getRCVShare = 0 Else getRCVShare = getRCVdata(DateFrom, DateTo, "QTY") / getTotBusData(DateFrom, DateTo, "RCVQTY")
    End If
End Function
Function getPOSShare(ByVal DateFrom As Date, ByVal DateTo As Date, Optional ByVal RetailnotQTY As Boolean) As Double
    If bGetsLetsOnly = True Then Exit Function
    If RetailnotQTY = True Then
        If getTotBusData(DateFrom, DateTo, "POSRetail") = 0 Then getPOSShare = 0 Else getPOSShare = CDbl(getPOSdata(DateFrom, DateTo, False) / getTotBusData(DateFrom, DateTo, "POSRetail"))
    Else
        If getTotBusData(DateFrom, DateTo, "POSQTY") = 0 Then getPOSShare = 0 Else getPOSShare = CDbl(getPOSdata(DateFrom, DateTo, True) / getTotBusData(DateFrom, DateTo, "POSQTY"))
    End If
End Function
Function getRCVMargin(ByVal DateFrom As Date, ByVal DateTo As Date) As Double
Dim denom As Double
    If bGetsLetsOnly = True Then Exit Function
    denom = getRCVdata(DateFrom, DateTo, "Retail")
    If denom = 0 Then getRCVMargin = 0: Exit Function
    getRCVMargin = getRCVContribution(DateFrom, DateTo) / denom
End Function
Function getRCVContribution(ByVal DateFrom As Date, ByVal DateTo As Date) As Double
    If bGetsLetsOnly = True Then Exit Function
    getRCVContribution = (getRCVdata(DateFrom, DateTo, "RetailNet") - getRCVdata(DateFrom, DateTo, "Cost"))
End Function
Function getPOSMargin(ByVal DateFrom As Date, ByVal DateTo As Date) As Double
    If bGetsLetsOnly = True Then Exit Function
    getPOSMargin = (getPOSdata(DateFrom, DateTo, False) - getRCVdata(DateFrom, DateTo, "Cost")) / getPOSdata(DateFrom, DateTo, False) * 100
End Function
Function getSellThrough(ByVal DateFrom As Date, ByVal DateTo As Date, ByRef SumOfRCVQTY As Double, ByRef SumOfPOSQTY As Double) As Boolean '@TP NEED TO FIX UP
    If c14DayST Is Nothing Then Exit Function
    getSellThrough = c14DayST.getSellThrough(DateFrom, DateTo, SumOfRCVQTY, SumOfPOSQTY)
End Function
Function CalculateTrialPeriod() As Variant
    If InStr(1, Me.sBldCode, "Ur") = 0 Then Exit Function
    CalculateTrialPeriod = cUSWDiv.CalculateTrialPeriod
End Function
Private Function getdata(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal QTYnotRetail As Boolean, Optional RCVnotPOS As Boolean = False, Optional ByVal USWType As e_POSUSWTypes = 0 _
    , Optional WkNo As Long = 0, Optional YrNo As Long = 0, Optional DivNo As Long = 599, Optional NeededValue As String) As Double
Dim a As Long, cntwks As Long, yr As Long
Dim wkfrom As Integer, wkto As Integer, yrfrom As Integer, yrto As Integer, pwkfrom As Integer, pwkto As Integer
Dim s As Double
Dim arr As Variant
Dim vYrno As Variant, vMno As Variant, vDiv As Variant, vWkno As Variant
Dim ValueNeeded As String
Dim v As Variant
Dim PD As cCBA_ProdPeriodSalesData

On Error GoTo Err_Routine
CBA_Error = ""

    If InStr(1, Me.sBldCode, "P") = 0 And InStr(1, Me.sBldCode, "R") = 0 Then getdata = 0: Exit Function
    If bGetsLetsOnly = True Then Exit Function
    If USWType = e_POSUSWTypes.eNotUSW Then
        If DivNo = 599 Then
            If cSales Is Nothing Then getdata = 0 Else getdata = cSales.getdata(DateFrom, DateTo, QTYnotRetail, RCVnotPOS, USWType, WkNo, YrNo, DivNo, NeededValue)
        Else
            If cSalesDiv Is Nothing Then getdata = 0 Else getdata = cSalesDiv.getdata(DateFrom, DateTo, QTYnotRetail, RCVnotPOS, USWType, WkNo, YrNo, DivNo, NeededValue)
        End If
    ElseIf USWType = e_POSUSWTypes.eUSWisActive Or USWType = e_POSUSWTypes.eUSWALL Then
        If DivNo = 599 Then
            If cUSW Is Nothing Then getdata = 0 Else getdata = cUSW.getdata(DateFrom, DateTo, QTYnotRetail, RCVnotPOS, USWType, WkNo, YrNo, DivNo, NeededValue)
        Else
            If cUSWDiv Is Nothing Then getdata = 0 Else getdata = cUSWDiv.getdata(DateFrom, DateTo, QTYnotRetail, RCVnotPOS, USWType, WkNo, YrNo, DivNo, NeededValue)
        End If
    Else
        If DivNo = 599 Then
            If cUSW Is Nothing Then getdata = 0 Else getdata = cUSW.getdata(DateFrom, DateTo, QTYnotRetail, RCVnotPOS, USWType, WkNo, YrNo, DivNo, NeededValue)
        Else
            If cUSWDiv Is Nothing Then getdata = 0 Else getdata = cUSWDiv.getdata(DateFrom, DateTo, QTYnotRetail, RCVnotPOS, USWType, WkNo, YrNo, DivNo, NeededValue)
        End If
    End If
    
    

Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_Prod.getData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function
'Private Function getdata(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal QTYnotRetail As Boolean, Optional RCVnotPOS As Boolean = False, Optional ByVal USWType As e_POSUSWTypes = 0 _
'    , Optional Wkno As Long = 0, Optional YRno As Long = 0, Optional Divno As Long = 599, Optional NeededValue As String)
'Dim a As Long, cntwks As Long, yr As Long
'Dim wkfrom As Integer, wkto As Integer, yrfrom As Integer, yrto As Integer, pwkfrom As Integer, pwkto As Integer
'Dim s As double
'Dim arr As Variant
'Dim vYrno As Variant, vMno As Variant, vDiv As Variant, vWkno As Variant
'Dim ValueNeeded As String
'Dim v As Variant
'Dim PD As cCBA_ProdPeriodSalesData
'
'On Error GoTo Err_Routine
'CBA_Error = ""
'
'    If InStr(1, Me.sBldCode, "P") = 0 And InStr(1, Me.sBldCode, "R") = 0 Then getdata = 0: Exit Function
'    If bGetsLetsOnly = True Then Exit Function
'    s = 0
'    If USWType = e_POSUSWTypes.eNotUSW Then
'        For vYrno = Year(DateFrom) To Year(DateTo)
'            For vMno = 1 To 12
'                For Each v In sdSales
'                    Set PD = sdSales(v)
'                    If PD.lMNo = vMno And PD.lYRno = vYrno Then
'                        Select Case NeededValue
'                            Case "RCVCost"
'                                s = s + PD.dblRCVCost
'                            Case "RCVRetailNet"
'                                s = s + PD.dblRCVRetailNet
'                            Case "RCVRetail"
'                                s = s + PD.dblRCVRetail
'                            Case "RCVCost"
'                                s = s + PD.dblRCVCost
'                            Case Else
'                                If RCVnotPOS = False Then
'                                    If QTYnotRetail = True Then
'                                        s = s + PD.dblPOSQTY
'                                    Else
'                                        s = s + PD.dblPOSRetail
'                                    End If
'                                Else
'                                    If QTYnotRetail = True Then
'                                        s = s + PD.dblRCVQTY
'                                    Else
'                                        s = s + PD.dblRCVRetail
'                                    End If
'                                End If
'                        End Select
'                    End If
'                Next
'            Next
'        Next
'    ElseIf USWType = e_POSUSWTypes.eUSWisActive Or USWType = e_POSUSWTypes.eUSWALL Then
'        For vYrno = Year(DateFrom) To Year(DateTo)
'            For vMno = 1 To 12
'                If Divno < 599 Then
'                    For Each v In sdUSWDiv
'                        Set PD = sdUSWDiv(v)
'                        If USWType = e_POSUSWTypes.eUSWisActive And PD.lDivno = Divno And PD.dblPOSQTY > 0 Then getdata = 1: Exit Function
'                        If PD.lDivno = Divno Then
'                            Select Case NeededValue
'                                Case "RCVCost"
'                                    s = s + PD.dblRCVCost
'                                Case "RCVRetailNet"
'                                    s = s + PD.dblRCVRetailNet
'                                Case "RCVRetail"
'                                    s = s + PD.dblRCVRetail
'                                Case "RCVCost"
'                                    s = s + PD.dblRCVCost
'                                Case Else
'                                    If RCVnotPOS = False Then
'                                        If QTYnotRetail = True Then
'                                            s = s + PD.dblPOSQTY
'                                        Else
'                                            s = s + PD.dblPOSRetail
'                                        End If
'                                    Else
'                                        If QTYnotRetail = True Then
'                                            s = s + PD.dblRCVQTY
'                                        Else
'                                            s = s + PD.dblRCVRetail
'                                        End If
'                                    End If
'                            End Select
'                        End If
'                    Next
'                Else
'                    For Each v In sdUSW
'                        Set PD = sdUSW(v)
'                        If USWType = e_POSUSWTypes.eUSWisActive And PD.dblPOSQTY > 0 Then getdata = 1: Exit Function
'                        Select Case NeededValue
'                            Case "RCVCostpwk"
'                                s = s + PD.dblRCVCostpwk
'                            Case "RCVRetailNetpwk"
'                                s = s + PD.dblRCVRetailNetpwk
'                            Case "RCVRetailpwk"
'                                s = s + PD.dblRCVRetailpwk
'                            Case "RCVCostpwk"
'                                s = s + PD.dblRCVCostpwk
'                            Case Else
'                                If RCVnotPOS = False Then
'                                    If QTYnotRetail = True Then
'                                        s = s + PD.dblPOSQTYpwk
'                                    Else
'                                        s = s + PD.dblPOSRetailpwk
'                                    End If
'                                Else
'                                    If QTYnotRetail = True Then
'                                        s = s + PD.dblRCVQTYpwk
'                                    Else
'                                        s = s + PD.dblRCVRetailpwk
'                                    End If
'                                End If
'                        End Select
'                    Next
'                End If
'            Next
'        Next
'    Else
'        wkfrom = CBA_BasicFunctions.CBA_IsoWeekNumber(DateFrom)
'        If DateFrom >= CBA_BasicFunctions.CBA_IsoYearStart(Year(DateFrom)) Then yrfrom = Year(DateFrom) Else yrfrom = Year(DateFrom) - 1
'        wkto = CBA_BasicFunctions.CBA_IsoWeekNumber(DateTo)
'        If DateTo < CBA_BasicFunctions.CBA_IsoYearStart(Year(DateTo)) Then yrto = Year(DateTo) + 1 Else yrto = Year(DateTo)
'        If wkto < 4 And Month(DateTo) > 10 And yrto = Year(DateTo) Then yrto = yrto + 1
'        pwkfrom = 0: pwkto = 0
'        cntwks = 0: s = 0
'        For yr = yrfrom To yrto
'            If yrfrom = yrto Then
'                pwkfrom = wkfrom: pwkto = wkto
'            Else
'                pwkfrom = 1: pwkto = 53
'                If yr = yrfrom Then pwkfrom = wkfrom
'                If yr = yrto Then pwkto = wkto
'            End If
'            For vWkno = pwkfrom To pwkto
'                If Divno = 599 Then
'                    For Each v In sdUSW
'                        Set PD = sdUSW(v)
'                        If PD.lWkno >= vWkno And PD.lYRno = yr Then
'                            If PD.lDayCnt > 3 And PD.dblPOSQTYpwk > 1 Then
'                                Select Case NeededValue
'                                    Case "RCVCostpwk"
'                                        s = s + PD.dblRCVCostpwk: cntwks = cntwks + 1
'                                    Case "RCVRetailNetpwk"
'                                        s = s + PD.dblRCVRetailNetpwk: cntwks = cntwks + 1
'                                    Case "RCVRetailpwk"
'                                        s = s + PD.dblRCVRetailpwk: cntwks = cntwks + 1
'                                    Case "RCVCostpwk"
'                                        s = s + PD.dblRCVCostpwk: cntwks = cntwks + 1
'                                    Case Else
'                                        If RCVnotPOS = False Then
'                                            If QTYnotRetail = True Then
'                                                s = s + PD.dblPOSQTYpwk: cntwks = cntwks + 1
'                                            Else
'                                                s = s + PD.dblPOSRetailpwk: cntwks = cntwks + 1
'                                            End If
'                                        Else
'                                            If QTYnotRetail = True Then
'                                                s = s + PD.dblRCVQTYpwk: cntwks = cntwks + 1
'                                            Else
'                                                s = s + PD.dblRCVRetailpwk: cntwks = cntwks + 1
'                                            End If
'                                        End If
'                                End Select
'                            End If
'                        End If
'                    Next
'                Else
'                    For Each v In sdUSWDiv
'                        Set PD = sdUSWDiv(v)
'                        If PD.lWkno >= vWkno And PD.lYRno = yr And PD.lDivno = Divno Then
'                            If PD.lDayCnt > 3 And PD.dblPOSQTYpwk > 1 Then
'                                Select Case NeededValue
'                                    Case "RCVCostpwk"
'                                        s = s + PD.dblRCVCostpwk: cntwks = cntwks + 1
'                                    Case "RCVRetailNetpwk"
'                                        s = s + PD.dblRCVRetailNetpwk: cntwks = cntwks + 1
'                                    Case "RCVRetailpwk"
'                                        s = s + PD.dblRCVRetailpwk: cntwks = cntwks + 1
'                                    Case "RCVCostpwk"
'                                        s = s + PD.dblRCVCostpwk: cntwks = cntwks + 1
'                                    Case Else
'                                        If RCVnotPOS = False Then
'                                            If QTYnotRetail = True Then
'                                                s = s + PD.dblPOSQTYpwk: cntwks = cntwks + 1
'                                            Else
'                                                s = s + PD.dblPOSRetailpwk: cntwks = cntwks + 1
'                                            End If
'                                        Else
'                                            If QTYnotRetail = True Then
'                                                s = s + PD.dblRCVQTYpwk: cntwks = cntwks + 1
'                                            Else
'                                                s = s + PD.dblRCVRetailpwk: cntwks = cntwks + 1
'                                            End If
'                                        End If
'                                End Select
'                            End If
'                        End If
'                    Next
'                End If
'            Next
'        Next
'        If USWType = e_POSUSWTypes.eProductLevel And cntwks > 0 Then s = s / cntwks Else s = 0
'        If USWType = e_POSUSWTypes.eUSWCNT Then s = cntwks
'    End If
'    getdata = s
'
'
'
'Exit_Routine:
'
'    On Error Resume Next
'    Exit Function
'
'Err_Routine:
'    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_Prod.getData", 3)
'    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
'    Debug.Print CBA_Error
'    Stop            ' ^RW Camera + next line
'    Resume Next
'    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
'    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
'    GoTo Exit_Routine
'    Resume Next
'
'End Function
Function getPOSdata(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal QTYnotRetail As Boolean, Optional ByVal USWType As e_POSUSWTypes = 0 _
    , Optional WkNo As Long = 0, Optional YrNo As Long = 0, Optional DivNo As Long = 599) As Double
    If bGetsLetsOnly = True Then Exit Function
    getPOSdata = getdata(DateFrom, DateTo, QTYnotRetail, False, USWType, WkNo, YrNo, DivNo)
End Function
Function getRCVdata(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal QueryType As String, Optional DivNo As Long = 599) As Double
    If bGetsLetsOnly = True Then Exit Function
    Select Case LCase(QueryType)
        Case "retail": getRCVdata = getdata(DateFrom, DateTo, False, True, eNotUSW, , , DivNo)
        Case "retailnet": getRCVdata = getdata(DateFrom, DateTo, False, True, eNotUSW, , , DivNo, "RCVRetailNet")
        Case "cost":  getRCVdata = getdata(DateFrom, DateTo, False, True, eNotUSW, , , DivNo, "RCVCost")
        Case "qty": getRCVdata = getdata(DateFrom, DateTo, False, True, eNotUSW, , , DivNo)
        Case "pieces": getRCVdata = getdata(DateFrom, DateTo, False, True, eNotUSW, , , DivNo, "RCVPieces")
        Case "usw": getRCVdata = getdata(DateFrom, DateTo, True, True, eProductLevel, , , DivNo)
        Case "uswall": getRCVdata = getdata(DateFrom, DateTo, True, True, eUSWALL, , , DivNo)
        Case "uswcnt": getRCVdata = getdata(DateFrom, DateTo, True, True, eUSWCNT, , , DivNo)
    End Select
End Function
Function getProdLabel(ByVal ProductLabelToSearchFor As String) As Boolean
Dim d As Variant
    If bGetsLetsOnly = True Then Exit Function
    If sdPLabels Is Nothing Then getProdLabel = False: Exit Function
    If sdPLabels.Exists(ProductLabelToSearchFor) Then getProdLabel = True
'    For d = 0 To sdPLabels.Count
'        Debug.Print d
'    Next
End Function

Public Property Get lPcode() As Long: lPcode = plPcode: End Property
Private Property Let lPcode(ByVal lNewValue As Long): plPcode = lNewValue: End Property
Public Property Get lPClass() As Long: lPClass = plPClass: End Property
Public Property Let lPClass(ByVal lNewValue As Long): plPClass = lNewValue: End Property
Public Property Get lPackSize() As Long: lPackSize = plPksize: End Property
Private Property Let lPackSize(ByVal lNewValue As Long): plPksize = lNewValue: End Property
Public Property Get lCon_PCode() As Long: lCon_PCode = plCon_PCode: End Property
Private Property Let lCon_PCode(ByVal lNewValue As Long): plCon_PCode = lNewValue: End Property
Public Property Get lEmpNo() As Long: lEmpNo = plEmpNo: End Property
Private Property Let lEmpNo(ByVal lNewValue As Long): plEmpNo = lNewValue: End Property
Public Property Get lIsRW() As Long: lIsRW = plIsRW: End Property
Private Property Let lIsRW(ByVal lNewValue As Long): plIsRW = lNewValue: End Property
Public Property Get lLegCG() As Long: lLegCG = plLegCG: End Property
Public Property Let lLegCG(ByVal lNewValue As Long): plLegCG = lNewValue: End Property
Public Property Get lLegSCG() As Long: lLegSCG = plLegSCG: End Property
Public Property Let lLegSCG(ByVal lNewValue As Long): plLegSCG = lNewValue: End Property
Public Property Get lACatNum() As Long: lACatNum = plACatNum: End Property
Public Property Let lACatNum(ByVal lNewValue As Long): plACatNum = lNewValue: End Property
Public Property Get lACGNum() As Long: lACGNum = plACGnum: End Property
Public Property Let lACGNum(ByVal lNewValue As Long): plACGnum = lNewValue: End Property
Public Property Get lASCGNum() As Long: lASCGNum = plASCGNum: End Property
Public Property Let lASCGNum(ByVal lNewValue As Long): plASCGNum = lNewValue: End Property
Public Property Get lSeaID() As Long: lSeaID = plSeaID: End Property
Private Property Let lSeaID(ByVal lNewValue As Long): plSeaID = lNewValue: End Property
Public Property Get sBldCode() As String: sBldCode = psBldCode: End Property
Public Property Let sBldCode(ByVal sNewValue As String): psBldCode = sNewValue: End Property
Public Property Get sProdDesc() As String: sProdDesc = psProdDesc: End Property
Private Property Let sProdDesc(ByVal sNewValue As String): psProdDesc = sNewValue: End Property
Public Property Get sContent() As String: sContent = psContent: End Property
Private Property Let sContent(ByVal sNewValue As String): psContent = sNewValue: End Property
Public Property Get sPCrdBnd() As String: sPCrdBnd = psPCrdBnd: End Property
Private Property Let sPCrdBnd(ByVal sNewValue As String): psPCrdBnd = sNewValue: End Property
Public Property Get sPCrdDesc() As String: sPCrdDesc = psPCrdDesc: End Property
Private Property Let sPCrdDesc(ByVal sNewValue As String): psPCrdDesc = sNewValue: End Property
Public Property Get sLegCGDesc() As String: sLegCGDesc = psLegCGDesc: End Property
Public Property Let sLegCGDesc(ByVal sNewValue As String): psLegCGDesc = sNewValue: End Property
Public Property Get sLegSCGDesc() As String: sLegSCGDesc = psLegSCGDesc: End Property
Public Property Let sLegSCGDesc(ByVal sNewValue As String): psLegSCGDesc = sNewValue: End Property
Public Property Get sLegCGMap() As String: sLegCGMap = psLegCGmap: End Property
Private Property Let sLegCGMap(ByVal sNewValue As String): psLegCGmap = sNewValue: End Property
Public Property Get sACat() As String: sACat = psACat: End Property
Public Property Let sACat(ByVal sNewValue As String): psACat = sNewValue: End Property
Public Property Get sACGDesc() As String: sACGDesc = psACGDesc: End Property
Public Property Let sACGDesc(ByVal sNewValue As String): psACGDesc = sNewValue: End Property
Public Property Get sASCGDesc() As String: sASCGDesc = psASCGDesc: End Property
Public Property Let sASCGDesc(ByVal sNewValue As String): psASCGDesc = sNewValue: End Property
Public Property Get lTaxID() As Long: lTaxID = plTaxID: End Property
Private Property Let lTaxID(ByVal lNewValue As Long): plTaxID = lNewValue: End Property
Public Property Get colComrade() As Collection: Set colComrade = pcolComrade: End Property
Public Property Set sdTotalBus(ByVal objNewValue As Scripting.Dictionary): Set psdTotalBus = objNewValue: End Property
Public Property Get bGetsLetsOnly() As Boolean: bGetsLetsOnly = pbGetsLetsOnly: End Property
Private Property Let bGetsLetsOnly(ByVal bNewValue As Boolean): pbGetsLetsOnly = bNewValue: End Property
Public Property Get sScanDataMSeg() As String: sScanDataMSeg = psScanDataMSeg: End Property
Public Property Let sScanDataMSeg(ByVal sNewValue As String): psScanDataMSeg = sNewValue: End Property
Public Function ManualMSeg(ByVal ManualSegmentNumber As Long, Optional setNewValue As String = "") As String
    If setNewValue <> "" Then
        If setNewValue = "0" Then psManualMSeg(ManualSegmentNumber) = "" Else psManualMSeg(ManualSegmentNumber) = setNewValue
    End If
    ManualMSeg = psManualMSeg(ManualSegmentNumber)
End Function
Public Property Get sdPLabels() As Scripting.Dictionary: Set sdPLabels = psdPLabels: End Property
Private Property Set sdPLabels(ByVal objNewValue As Scripting.Dictionary): Set psdPLabels = objNewValue: End Property
Public Property Get lTrademarkNo() As Long: lTrademarkNo = plTrademarkNo: End Property
Private Property Let lTrademarkNo(ByVal lNewValue As Long): plTrademarkNo = lNewValue: End Property
'Public Property Get sdWastage() As Scripting.Dictionary: Set sdWastage = psdWastage: End Property
'Private Property Set sdWastage(ByVal NewValue As Scripting.Dictionary): Set psdWastage = NewValue: End Property
Private Property Get cSales() As cCBA_ProdPeriodSalesData: Set cSales = pcSales: End Property
Public Property Set cSales(ByVal NewValue As cCBA_ProdPeriodSalesData): Set pcSales = NewValue: End Property
Private Property Get cUSW() As cCBA_ProdPeriodSalesData: Set cUSW = pcUSW: End Property
Public Property Set cUSW(ByVal NewValue As cCBA_ProdPeriodSalesData): Set pcUSW = NewValue: End Property
Private Property Get c14DayST() As cCBA_ProdPeriodSalesData: Set c14DayST = pc14DayST: End Property
Public Property Set c14DayST(ByVal NewValue As cCBA_ProdPeriodSalesData): Set pc14DayST = NewValue: End Property
Private Property Get cSalesDiv() As cCBA_ProdPeriodSalesData: Set cSalesDiv = pcSalesDiv: End Property
Public Property Set cSalesDiv(ByVal NewValue As cCBA_ProdPeriodSalesData): Set pcSalesDiv = NewValue: End Property
Private Property Get cUSWDiv() As cCBA_ProdPeriodSalesData: Set cUSWDiv = pcUSWDiv: End Property
Public Property Set cUSWDiv(ByVal NewValue As cCBA_ProdPeriodSalesData): Set pcUSWDiv = NewValue: End Property
Public Property Get cWastage() As cCBA_WastageData: Set cWastage = pcWastage: End Property
Private Property Set cWastage(ByVal NewValue As cCBA_WastageData): Set pcWastage = NewValue: End Property
Private Property Get lDoc_ID() As Long: lDoc_ID = plDoc_ID: End Property
Public Property Let lDoc_ID(ByVal NewValue As Long)
Dim v As Variant
plDoc_ID = NewValue
If Not cSales Is Nothing Then cSales.lDoc_ID = NewValue
If Not cSalesDiv Is Nothing Then cSalesDiv.lDoc_ID = NewValue
If Not cUSW Is Nothing Then cUSW.lDoc_ID = NewValue
If Not cUSWDiv Is Nothing Then cUSWDiv.lDoc_ID = NewValue
If Not c14DayST Is Nothing Then c14DayST.lDoc_ID = NewValue
End Property
