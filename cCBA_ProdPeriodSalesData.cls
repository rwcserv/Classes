VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCBA_ProdPeriodSalesData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type PPSD
    lYRno As Long
    lWkno As Long
    lProductcode As Long
    lMNo As Long
    dblRCVRetailpwk As Double
    dblRCVRetailNetpwk As Double
    dblRCVCostpwk As Double
    dblRCVQTYpwk As Double
    dblRCVPiecespwk As Double
    dblPOSQTYpwk As Double
    dblPOSRetailpwk As Double
    lDayCnt As Long
    lNoOfStores As Long
    lCGno As Long
    lDivno As Long
    dblRCVRetail As Double
    dblRCVRetailNet As Double
    dblRCVCost As Double
    dblRCVQTY As Double
    dblRCVPieces As Double
    dblPOSQTY As Double
    dblPOSRetail As Double
    dOSD As Date
    dEOD As Date
    dblSellThrough As Double
End Type
Private pArrSales() As PPSD
Private plPcode As Long
Private plDoc_ID As Long
Public Function Build(ByRef RS As ADODB.Recordset, ByVal no As Long, ByVal DataSetType As String, Optional ByRef ASTO As Scripting.Dictionary) As Boolean

Dim cnt As Long, totcnt As Long, a As Long
Dim SD As PPSD
Dim pnts As Variant
Dim TempArr() As PPSD
Dim lcol As Long
Dim ASVal As Long
If no = 0 Then Exit Function

    If RS.Fields(0).Name = "productcode" Then lcol = 0 Else lcol = 2

    On Error Resume Next
    If UBound(pArrSales) < 0 Then
    ReDim TempArr(0 To no - 1)
    Else
    ReDim TempArr(0 To UBound(pArrSales) + no)
        For a = 0 To UBound(pArrSales)
            TempArr(a).lYRno = pArrSales(a).lYRno
            TempArr(a).lWkno = pArrSales(a).lWkno
            TempArr(a).lProductcode = pArrSales(a).lProductcode
            TempArr(a).lMNo = pArrSales(a).lMNo
            TempArr(a).dblRCVRetailpwk = pArrSales(a).dblRCVRetailpwk
            TempArr(a).dblRCVRetailNetpwk = pArrSales(a).dblRCVRetailNetpwk
            TempArr(a).dblRCVCostpwk = pArrSales(a).dblRCVCostpwk
            TempArr(a).dblRCVQTYpwk = pArrSales(a).dblRCVQTYpwk
            TempArr(a).dblRCVPiecespwk = pArrSales(a).dblRCVPiecespwk
            TempArr(a).dblPOSQTYpwk = pArrSales(a).dblPOSQTYpwk
            TempArr(a).dblPOSRetailpwk = pArrSales(a).dblPOSRetailpwk
            TempArr(a).lDayCnt = pArrSales(a).lDayCnt
            TempArr(a).lNoOfStores = pArrSales(a).lNoOfStores
            TempArr(a).lCGno = pArrSales(a).lCGno
            TempArr(a).lDivno = pArrSales(a).lDivno
            TempArr(a).dblRCVRetail = pArrSales(a).dblRCVRetail
            TempArr(a).dblRCVRetailNet = pArrSales(a).dblRCVRetailNet
            TempArr(a).dblRCVCost = pArrSales(a).dblRCVCost
            TempArr(a).dblRCVQTY = pArrSales(a).dblRCVQTY
            TempArr(a).dblRCVPieces = pArrSales(a).dblRCVPieces
            TempArr(a).dblPOSQTY = pArrSales(a).dblPOSQTY
            TempArr(a).dblPOSRetail = pArrSales(a).dblPOSRetail
            TempArr(a).dOSD = pArrSales(a).dOSD
            TempArr(a).dEOD = pArrSales(a).dEOD
        Next
    End If
    Err.Clear
    On Error GoTo 0
    a = a - 1
    Do Until CLng(RS(lcol)) <> lPcode
        a = a + 1
        
        If InStr(1, DataSetType, "Sales") > 0 Then
            SD.lYRno = CLng(RS(0))
            SD.lMNo = CLng(RS(1))
            SD.lProductcode = CLng(RS(2))
            SD.dblRCVRetail = CDbl(RS(3))
            SD.dblRCVRetailNet = CDbl(RS(4))
            SD.dblRCVCost = CDbl(RS(5))
            SD.dblRCVQTY = CDbl(RS(6))
            SD.dblRCVPieces = CDbl(RS(7))
            SD.dblPOSQTY = CDbl(RS(8))
            SD.dblPOSRetail = CDbl(RS(9))
            If InStr(1, DataSetType, "Div") > 0 Then SD.lDivno = CLng(RS(10))
        End If
        If ASTO Is Nothing Then
            If InStr(1, DataSetType, "USW") > 0 Then
                SD.lYRno = CLng(RS(0))
                SD.lWkno = CLng(RS(1))
                SD.lProductcode = CLng(RS(2))
                SD.dblRCVRetailpwk = CDbl(RS(3))
                SD.dblRCVRetailNetpwk = CDbl(RS(4))
                SD.dblRCVCostpwk = CDbl(RS(5))
                SD.dblRCVQTYpwk = CDbl(RS(6))
                SD.dblRCVPiecespwk = CDbl(RS(7))
                SD.dblPOSQTYpwk = CDbl(RS(8))
                SD.dblPOSRetailpwk = CDbl(RS(9))
                SD.lDayCnt = CLng(RS(10))
                SD.lNoOfStores = CLng(RS(11))
                SD.lCGno = CLng(RS(12))
                If InStr(1, DataSetType, "Div") > 0 Then SD.lDivno = CLng(RS(13))
            End If
        Else
            If InStr(1, DataSetType, "USW") > 0 Then
                SD.lYRno = CLng(RS(0))
                SD.lWkno = CLng(RS(1))
                SD.lProductcode = CLng(RS(2))
                SD.lNoOfStores = CLng(RS(11))
                If InStr(1, DataSetType, "Div") > 0 Then SD.lDivno = CLng(RS(13))
                If ASTO.Exists(CLng(RS(0))) Then
                    If ASTO(CLng(RS(0))).Exists(CLng((IIf(RS.Fields(1).Name = "Wkno", Month(CBA_BasicFunctions.GetDayFromWeekNumber(RS(0), RS(1))), RS(1))))) Then
                        If ASTO(CLng(RS(0)))(CLng(IIf(RS.Fields(1).Name = "Wkno", Month(CBA_BasicFunctions.GetDayFromWeekNumber(RS(0), RS(1))), RS(1))))(IIf(SD.lDivno = 0, 599, SD.lDivno)) Then
                            ASVal = ASTO(CLng(RS(0)))(CLng(IIf(RS.Fields(1).Name = "Wkno", Month(CBA_BasicFunctions.GetDayFromWeekNumber(RS(0), RS(1))), RS(1))))(IIf(SD.lDivno = 0, 599, SD.lDivno))
                        End If
                    End If
                End If
                If ASVal > 0 Then
                    SD.dblRCVRetailpwk = CDbl(RS(3)) * SD.lNoOfStores / ASVal
                    SD.dblRCVRetailNetpwk = CDbl(RS(4)) * SD.lNoOfStores / ASVal
                    SD.dblRCVCostpwk = CDbl(RS(5)) * SD.lNoOfStores / ASVal
                    SD.dblRCVQTYpwk = CDbl(RS(6)) * SD.lNoOfStores / ASVal
                    SD.dblRCVPiecespwk = CDbl(RS(7)) * SD.lNoOfStores / ASVal
                    SD.dblPOSQTYpwk = CDbl(RS(8)) * SD.lNoOfStores / ASVal
                    SD.dblPOSRetailpwk = CDbl(RS(9)) * SD.lNoOfStores / ASVal
                End If
                SD.lDayCnt = CLng(RS(10))
                SD.lCGno = CLng(RS(12))
            End If
        End If
        
        
        If DataSetType = "ST" Then
            SD.dOSD = CDate(RS(1))
            SD.dEOD = CDate(RS(2))
            SD.dblPOSQTY = CDbl(RS(3))
            SD.dblPOSRetail = CDbl(RS(4))
            SD.dblRCVQTY = CDbl(RS(5))
            SD.dblSellThrough = CDbl(RS(6))
        End If
        
        TempArr(a) = SD
        RS.MoveNext
        If RS.EOF Then Exit Do
    Loop
    pArrSales = TempArr
    Erase TempArr
        
End Function
Public Function CalculateTrialPeriod() As Variant
Dim a As Long
Dim div As Long
Dim arr As Variant
Dim yearDic As Scripting.Dictionary, WeekDic As Scripting.Dictionary
Dim va As Variant, v As Variant
Dim EndDate As Date, StartDate As Date
Dim TempDic As Scripting.Dictionary

Set yearDic = New Scripting.Dictionary
Set WeekDic = New Scripting.Dictionary

For div = 501 To 509
    If div = 508 Then div = 509
    For a = LBound(pArrSales) To UBound(pArrSales)
        If pArrSales(a).lDivno = div Then
            If yearDic.Exists(CLng(pArrSales(a).lYRno)) = False Then
                Set WeekDic = New Scripting.Dictionary
                yearDic.Add CLng(pArrSales(a).lYRno), WeekDic
            End If
            If yearDic(CLng(pArrSales(a).lYRno)).Exists(CLng(pArrSales(a).lWkno)) = False Then
                If pArrSales(a).dblPOSQTYpwk > 0 And pArrSales(a + 1).dblPOSQTYpwk > 0 And pArrSales(a + 2).dblPOSQTYpwk > 0 And pArrSales(a + 3).dblPOSQTYpwk > 0 Then
                    yearDic(CLng(pArrSales(a).lYRno)).Add CLng(pArrSales(a).lWkno), CStr(div)
                End If
            Else
                If pArrSales(a).dblPOSQTYpwk > 0 And pArrSales(a + 1).dblPOSQTYpwk > 0 And pArrSales(a + 2).dblPOSQTYpwk > 0 And pArrSales(a + 3).dblPOSQTYpwk > 0 Then
                    yearDic(CLng(pArrSales(a).lYRno))(CLng(pArrSales(a).lWkno)) = yearDic(CLng(pArrSales(a).lYRno))(CLng(pArrSales(a).lWkno)) & "-" & CStr(div)
                End If
            End If
        End If
    Next
Next

For Each v In yearDic
    For Each va In yearDic(v)
        'Debug.Print YearDic(v)(va) & " | Year: " & v & " Wk: " & va
        If Len(yearDic(v)(va)) >= 7 And StartDate = 0 Then
            StartDate = CBA_BasicFunctions.GetDayFromWeekNumber(CInt(v), CInt(va))
        End If
        If Len(yearDic(v)(va)) >= 19 And EndDate = 0 Then
            EndDate = CBA_BasicFunctions.GetDayFromWeekNumber(CInt(v), CInt(va))
        End If
    Next
Next
ReDim arr(0 To 1)
arr(0) = StartDate
arr(1) = EndDate
CalculateTrialPeriod = arr

End Function
Public Function getSellThrough(ByVal DateFrom As Date, ByVal DateTo As Date, ByRef SumOfRCVQTY As Double, ByRef SumOfPOSQTY As Double) As Boolean '@TP NEED TO FIX UP
Dim a As Long
    For a = LBound(pArrSales) To UBound(pArrSales)
        If pArrSales(a).dOSD <= DateFrom And pArrSales(a).dOSD >= DateFrom Then
            getSellThrough = True
            SumOfRCVQTY = SumOfPOSQTY + pArrSales(a).dblPOSQTY
            SumOfRCVQTY = SumOfRCVQTY + pArrSales(a).dblRCVQTY
        End If
    Next
End Function
Public Function getdata(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal QTYnotRetail As Boolean, Optional RCVnotPOS As Boolean = False, Optional ByVal USWType As e_POSUSWTypes = 0 _
    , Optional WkNo As Long = 0, Optional YrNo As Long = 0, Optional DivNo As Long = 599, Optional NeededValue As String) As Double
Dim a As Long, cntwks As Long, yr As Long
Dim wkfrom As Integer, wkto As Integer, yrfrom As Integer, yrto As Integer, pwkfrom As Integer, pwkto As Integer
Dim s As Double
Dim arr As Variant
Dim vYrno As Variant, vMno As Variant, vDiv As Variant, vWkno As Variant
Dim ValueNeeded As String
Dim v As Variant



    s = 0
    If USWType = e_POSUSWTypes.eNotUSW Then
        For a = 0 To UBound(pArrSales)
            If DateFrom <= DateSerial(pArrSales(a).lYRno, pArrSales(a).lMNo, 1) And DateTo >= DateSerial(pArrSales(a).lYRno, pArrSales(a).lMNo, 1) And pArrSales(a).lDivno = IIf(DivNo = 599, 0, DivNo) Then
                Select Case NeededValue
                    Case "RCVCost"
                        s = s + pArrSales(a).dblRCVCost
                    Case "RCVRetailNet"
                        s = s + pArrSales(a).dblRCVRetailNet
                    Case "RCVRetail"
                        s = s + pArrSales(a).dblRCVRetail
                    Case "RCVCost"
                        s = s + pArrSales(a).dblRCVCost
                    Case Else
                        If RCVnotPOS = False Then
                            If QTYnotRetail = True Then
                                s = s + pArrSales(a).dblPOSQTY
                            Else
                                s = s + pArrSales(a).dblPOSRetail
                            End If
                        Else
                            If QTYnotRetail = True Then
                                s = s + pArrSales(a).dblRCVQTY
                            Else
                                s = s + pArrSales(a).dblRCVRetail
                            End If
                        End If
                End Select
            End If
        Next
    ElseIf USWType = e_POSUSWTypes.eUSWisActive Or USWType = e_POSUSWTypes.eUSWALL Then
        If WkNo = 0 And YrNo = 0 Then getdata = 0: Exit Function
        s = 0
        For a = 0 To UBound(pArrSales)
            If pArrSales(a).lYRno = YrNo And pArrSales(a).lWkno = WkNo And pArrSales(a).dblPOSQTYpwk > 0 And pArrSales(a).lDivno = IIf(DivNo = 599, 0, DivNo) Then
                If USWType = e_POSUSWTypes.eUSWisActive Then getdata = 1: Exit Function
                Select Case NeededValue
                    Case "RCVCostpwk"
                        getdata = ((pArrSales(a).dblRCVCostpwk / pArrSales(a).lDayCnt) * 7)
                    Case "RCVRetailNetpwk"
                        getdata = ((pArrSales(a).dblRCVRetailNetpwk / pArrSales(a).lDayCnt) * 7)
                    Case "RCVRetailpwk"
                        getdata = ((pArrSales(a).dblRCVRetailpwk / pArrSales(a).lDayCnt) * 7)
                    Case "RCVCostpwk"
                        getdata = ((pArrSales(a).dblRCVCostpwk / pArrSales(a).lDayCnt) * 7)
                    Case Else
                        If RCVnotPOS = False Then
                            If QTYnotRetail = True Then
                                getdata = ((pArrSales(a).dblPOSQTYpwk / pArrSales(a).lDayCnt) * 7)
                            Else
                                getdata = ((pArrSales(a).dblPOSRetailpwk / pArrSales(a).lDayCnt) * 7)
                            End If
                        Else
                            If QTYnotRetail = True Then
                                getdata = ((pArrSales(a).dblRCVQTYpwk / pArrSales(a).lDayCnt) * 7)
                            Else
                                getdata = ((pArrSales(a).dblRCVRetailpwk / pArrSales(a).lDayCnt) * 7)
                            End If
                        End If
                End Select
                Exit Function
            End If
        Next
    Else
        wkfrom = CBA_BasicFunctions.CBA_IsoWeekNumber(DateFrom)
        If DateFrom >= CBA_BasicFunctions.CBA_IsoYearStart(Year(DateFrom)) Then yrfrom = Year(DateFrom) Else yrfrom = Year(DateFrom) - 1
        wkto = CBA_BasicFunctions.CBA_IsoWeekNumber(DateTo)
        If DateTo < CBA_BasicFunctions.CBA_IsoYearStart(Year(DateTo)) Then yrto = Year(DateTo) + 1 Else yrto = Year(DateTo)
        If wkto < 4 And Month(DateTo) > 10 And yrto = Year(DateTo) Then yrto = yrto + 1
        pwkfrom = 0: pwkto = 0
        cntwks = 0: s = 0
        For yr = yrfrom To yrto
            If yrfrom = yrto Then
                pwkfrom = wkfrom: pwkto = wkto
            Else
                pwkfrom = 1: pwkto = 53
                If yr = yrfrom Then pwkfrom = wkfrom
                If yr = yrto Then pwkto = wkto
            End If
            For vWkno = pwkfrom To pwkto
                For a = 0 To UBound(pArrSales)
                    If pArrSales(a).lWkno >= vWkno And pArrSales(a).lYRno = yr And pArrSales(a).lDivno = IIf(DivNo = 599, 0, DivNo) Then
                        If pArrSales(a).lDayCnt > 3 And pArrSales(a).dblPOSQTYpwk > 1 Then
                            Select Case NeededValue
                                Case "RCVCostpwk"
                                    s = s + ((pArrSales(a).dblRCVCostpwk / pArrSales(a).lDayCnt) * 7): cntwks = cntwks + 1
                                Case "RCVRetailNetpwk"
                                    s = s + ((pArrSales(a).dblRCVRetailNetpwk / pArrSales(a).lDayCnt) * 7): cntwks = cntwks + 1
                                Case "RCVRetailpwk"
                                    s = s + ((pArrSales(a).dblRCVRetailpwk / pArrSales(a).lDayCnt) * 7): cntwks = cntwks + 1
                                Case "RCVCostpwk"
                                    s = s + ((pArrSales(a).dblRCVCostpwk / pArrSales(a).lDayCnt) * 7): cntwks = cntwks + 1
                                Case Else
                                    If RCVnotPOS = False Then
                                        If QTYnotRetail = True Then
                                            s = s + ((pArrSales(a).dblPOSQTYpwk / pArrSales(a).lDayCnt) * 7): cntwks = cntwks + 1
                                        Else
                                            s = s + ((pArrSales(a).dblPOSRetailpwk / pArrSales(a).lDayCnt) * 7): cntwks = cntwks + 1
                                        End If
                                    Else
                                        If QTYnotRetail = True Then
                                            s = s + ((pArrSales(a).dblRCVQTYpwk / pArrSales(a).lDayCnt) * 7): cntwks = cntwks + 1
                                        Else
                                            s = s + ((pArrSales(a).dblRCVRetailpwk / pArrSales(a).lDayCnt) * 7): cntwks = cntwks + 1
                                        End If
                                    End If
                            End Select
                        End If
                    End If
                Next
            Next
        Next
        If USWType = e_POSUSWTypes.eProductLevel And cntwks > 0 Then s = s / cntwks Else s = 0
        If USWType = e_POSUSWTypes.eUSWCNT Then s = cntwks
    End If
    getdata = s





End Function















''''------------------ORIGINAL---------------------
''Private plYRno As Long
''Private plWkno As Long
''Private plProductcode As Long
''Private plMNo As Long
''Private pdblRCVRetailpwk As double
''Private pdblRCVRetailNetpwk As double
''Private pdblRCVCostpwk As double
''Private pdblRCVQTYpwk As double
''Private pdblRCVPiecespwk As double
''Private pdblPOSQTYpwk As double
''Private pdblPOSRetailpwk As double
''Private plDayCnt As Long
''Private plNoOfStores As Long
''Private plCGno As Long
''Private plDivno As Long
''Private pdblRCVRetail As double
''Private pdblRCVRetailNet As double
''Private pdblRCVCost As double
''Private pdblRCVQTY As double
''Private pdblRCVPieces As double
''Private pdblPOSQTY As double
''Private pdblPOSRetail As double
''Private pdOSD As Date
''Private pdEOD As Date
''Public Property Get lYRno() As Long: lYRno = plYRno: End Property
''Public Property Let lYRno(ByVal NewValue As Long): plYRno = NewValue: End Property
''Public Property Get lWkno() As Long: lWkno = plWkno: End Property
''Public Property Let lWkno(ByVal NewValue As Long): plWkno = NewValue: End Property
''Public Property Get lProductcode() As Long: lProductcode = plProductcode: End Property
''Public Property Let lProductcode(ByVal NewValue As Long): plProductcode = NewValue: End Property
''Public Property Get dblRCVRetailpwk() As double: dblRCVRetailpwk = pdblRCVRetailpwk: End Property
''Public Property Let dblRCVRetailpwk(ByVal NewValue As double): pdblRCVRetailpwk = NewValue: End Property
''Public Property Get dblRCVRetailNetpwk() As double: dblRCVRetailNetpwk = pdblRCVRetailNetpwk: End Property
''Public Property Let dblRCVRetailNetpwk(ByVal NewValue As double): pdblRCVRetailNetpwk = NewValue: End Property
''Public Property Get dblRCVCostpwk() As double: dblRCVCostpwk = pdblRCVCostpwk: End Property
''Public Property Let dblRCVCostpwk(ByVal NewValue As double): pdblRCVCostpwk = NewValue: End Property
''Public Property Get dblRCVQTYpwk() As double: dblRCVQTYpwk = pdblRCVQTYpwk: End Property
''Public Property Let dblRCVQTYpwk(ByVal NewValue As double): pdblRCVQTYpwk = NewValue: End Property
''Public Property Get dblRCVPiecespwk() As double: dblRCVPiecespwk = pdblRCVPiecespwk: End Property
''Public Property Let dblRCVPiecespwk(ByVal NewValue As double): pdblRCVPiecespwk = NewValue: End Property
''Public Property Get dblPOSQTYpwk() As double: dblPOSQTYpwk = pdblPOSQTYpwk: End Property
''Public Property Let dblPOSQTYpwk(ByVal NewValue As double): pdblPOSQTYpwk = NewValue: End Property
''Public Property Get dblPOSRetailpwk() As double: dblPOSRetailpwk = pdblPOSRetailpwk: End Property
''Public Property Let dblPOSRetailpwk(ByVal NewValue As double): pdblPOSRetailpwk = NewValue: End Property
''Public Property Get lNoOfStores() As Long: lNoOfStores = plNoOfStores: End Property
''Public Property Let lNoOfStores(ByVal NewValue As Long): plNoOfStores = NewValue: End Property
''Public Property Get lCGno() As Long: lCGno = plCGno: End Property
''Public Property Let lCGno(ByVal NewValue As Long): plCGno = NewValue: End Property
''Public Property Get lDivno() As Long: lDivno = plDivno: End Property
''Public Property Let lDivno(ByVal NewValue As Long): plDivno = NewValue: End Property
''Public Property Get lDayCnt() As Long: lDayCnt = plDayCnt: End Property
''Public Property Let lDayCnt(ByVal NewValue As Long): plDayCnt = NewValue: End Property
''Public Property Get dblRCVRetail() As double: dblRCVRetail = pdblRCVRetail: End Property
''Public Property Let dblRCVRetail(ByVal NewValue As double): pdblRCVRetail = NewValue: End Property
''Public Property Get dblRCVRetailNet() As double: dblRCVRetailNet = pdblRCVRetailNet: End Property
''Public Property Let dblRCVRetailNet(ByVal NewValue As double): pdblRCVRetailNet = NewValue: End Property
''Public Property Get dblRCVCost() As double: dblRCVCost = pdblRCVCost: End Property
''Public Property Let dblRCVCost(ByVal NewValue As double): pdblRCVCost = NewValue: End Property
''Public Property Get dblRCVQTY() As double: dblRCVQTY = pdblRCVQTY: End Property
''Public Property Let dblRCVQTY(ByVal NewValue As double): pdblRCVQTY = NewValue: End Property
''Public Property Get dblRCVPieces() As double: dblRCVPieces = pdblRCVPieces: End Property
''Public Property Let dblRCVPieces(ByVal NewValue As double): pdblRCVPieces = NewValue: End Property
''Public Property Get dblPOSQTY() As double: dblPOSQTY = pdblPOSQTY: End Property
''Public Property Let dblPOSQTY(ByVal NewValue As double): pdblPOSQTY = NewValue: End Property
''Public Property Get dblPOSRetail() As double: dblPOSRetail = pdblPOSRetail: End Property
''Public Property Let dblPOSRetail(ByVal NewValue As double): pdblPOSRetail = NewValue: End Property
''Public Property Get lMNo() As Long: lMNo = plMNo: End Property
''Public Property Let lMNo(ByVal NewValue As Long): plMNo = NewValue: End Property
''Public Property Get dOSD() As Date: dOSD = pdOSD: End Property
''Public Property Let dOSD(ByVal NewValue As Date): pdOSD = NewValue: End Property
''Public Property Get dEOD() As Date: dEOD = pdEOD: End Property
''Public Property Let dEOD(ByVal NewValue As Date): pdEOD = NewValue: End Property

Public Property Get lPcode() As Long: lPcode = plPcode: End Property

Public Property Let lPcode(ByVal NewValue As Long): plPcode = NewValue: End Property

Private Property Get lDoc_ID() As Long: lDoc_ID = plDoc_ID: End Property

Public Property Let lDoc_ID(ByVal NewValue As Long): plDoc_ID = NewValue: End Property
