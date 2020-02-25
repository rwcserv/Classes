VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCBA_Forecasting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit       ' cCBA_Forcasting Changed 200216
'''Option Private Module       ' Excel users cannot access procedures

''Private Type CBA_BTF_ForecastMetrics
''    OriginalForecast As Single
''    ReForecast As Single
''    Actual As Single
''End Type

''Private Type CBA_BTF_CGDataDetail
''    Year As Long
''    Month As String
''    POSQTY As Long
''    POSPYQTY As Long
''    POSYOYQTY As Single
''    POSRET As Single
''    POSPYRET As Single
''    POSYOYRET As Single
''    RCVMargin As Single
''    RCVPYMargin As Single
''    CPU As Single
''    CPUYOY As Single
''    CPUPY As Single
''    ForeCPU As Single
''    ForePrice As Single
''    ForeQTY As Single
''    ForeRetail As Single
''    PForeRetail As Single
''    ForePYRetail As Single
''    ForePYQTY As Single
''    ForeCost As Single
''    ForeRetailNet As Single
''    ForePYCost As Single
''    ForePYRetailNet As Single
''    ForeRCVMargin As Single
''    PForeRCVMargin As Single
''    PUplift As Single
''    Uplift As Single
''
''End Type

''Private Type CBA_BTF_SaMData
''    Year As Long
''    Month As String
''    TotSales As CBA_BTF_ForecastMetrics
''    CRSales As CBA_BTF_ForecastMetrics
''    FVSales As CBA_BTF_ForecastMetrics
''    MeatSales As CBA_BTF_ForecastMetrics
''    ChilSales As CBA_BTF_ForecastMetrics
''    SpecialSales As CBA_BTF_ForecastMetrics
''    SeasonalSales As CBA_BTF_ForecastMetrics
''End Type


''Private basedatapulled As Boolean
''Private FMarginP  As Variant
''Private FReMarginP As Variant
''Private FReSales As Variant
''Private FSales As Variant
''Private CBA_SCGData() As CBA_BTF_SCG
''Private Type CBA_BTF_ReportParamaters
''    ReportName As String
'Private plPSMonth As Long
Private plPSYear As Long
'Private plPEMonth As Long
Private plPEYear As Long
Private psBD As String
Private psGBD As String
Private plCG As Long
Private plSCG As Long
Private plProdClass As Long
Private parrCGs() As Variant
Private pvRtnArr() As Variant
Private pcFCbM() As CBA_BTF_MonthData
Private Enum eArrPos
    [_First]
    eCSRF = 2
    eCSSA = 3
    eCMRF = 4
    eCMSA = 5
    eCSYoY = 6
    eCMYoY = 7
    [_Last]
End Enum
Const CDFT0 = 0, CDFT1 = 7, CDFTL = 16, CDFTT = 25, CELMS = 34


''Private RCVRetNet As Single, RCVCost As Single, RCVRet As Single
''Private TotRCVRetNet As Single, TotRCVRet As Single, TotRCVCost As Single
''Private SummaryData() As Single

Public Function GenForecasting(PSYear As Long, PEYear As Long, BD As String, GBD As String, CG As Long, scg As Long, ProdClass As Long) As Boolean
    Dim lIdx As Long
    lPSYear = PSYear        ' , CGs() As Variant
    lPEYear = PEYear
    sBD = BD
    sGBD = GBD
    lCG = CG
    lSCG = scg
    lProdClass = ProdClass
    Call CBA_COM_SQLQueries.CBA_COM_GenPullSQL("SCGListCG", , , CG)
    parrCGs = CBA_CBISarr
    For lIdx = 0 To UBound(parrCGs, 2)
        CBA_CBISarr(1, lIdx) = Left(CBA_CBISarr(1, lIdx), 2)
    Next
    Call CBA_BTF_Runtime.setCGs(CBA_CBISarr)
    Call SetForecastArray
End Function

Public Function GetForecastValue(lRowNo As Long, sE As String) As Variant
    Static lRowCount As Long
    If lRowCount = 0 Then
        lRowCount = pvRtnArr(0, 0)
    End If
    Select Case sE
        Case "Count"
            GetForecastValue = lRowCount                                         ' Record count
        Case "Desc"
            GetForecastValue = pvRtnArr(lRowNo, 1)                               ' The Desc for the line
        Case "SalesRF"
            GetForecastValue = pvRtnArr(lRowNo, CDFT0 + eCSRF)                   ' Return this one into Sales Forecast
        Case "MarginRF"
            GetForecastValue = pvRtnArr(lRowNo, CDFT0 + eCMRF)                   ' Return this one into Margin Forecast
        Case "SalesYoY"
            GetForecastValue = pvRtnArr(lRowNo, CDFT0 + eCSYoY)                  ' Return this one into Sales YoY %
        Case "MarginYoY"
            GetForecastValue = pvRtnArr(lRowNo, CDFT0 + eCMYoY)                  ' Return this one into Margin YoY %
        Case "TSalesRF"
            GetForecastValue = pvRtnArr(lRowCount, CDFTT + eCSRF)                ' Return this one into Sales Total Forecast
        Case "TMarginRF"
            GetForecastValue = pvRtnArr(lRowCount, CDFTT + eCMRF)                ' Return this one into Margin Total Forecast
        Case "TSalesYoY"
            GetForecastValue = pvRtnArr(lRowCount, CDFTT + eCSYoY)               ' Return this one into Sales Total YoY %
        Case "TMarginYoY"
            GetForecastValue = pvRtnArr(lRowCount, CDFTT + eCMYoY)               ' Return this one into Margin Total YoY %
        End Select
End Function
Public Function getForecast(ByVal YrNo As Long, ByVal MthNo As Long, ByVal PClass As Long, ByVal CGno As Long, SCGno As Long) As CBA_BTF_ForecastTypes
    Set getForecast = pcFCbM(YrNo, MthNo, PClass).CGData(CGno, SCGno)
End Function

Private Function SetForecastArray() As Variant
    Dim Data() As Variant
    Dim RP As CBA_BTF_ReportParamaters
    Dim NoOfMonths As Long, Rounder As Long, yr As Long, Mnth As Long, MTick As Long, Pc As Long, lCG As Long, SalesFC As Long
    Dim SalesM As Long, oprow As Long, Mtype As Long, opcol As Long, mont As Long, YearFrom  As Long, YearTo As Long
    Dim MonthFrom As Long, MonthTo As Long, curM As Long
    Dim b As Long, lRowNo As Long, col
    Dim PCL As Long, TotalMonths As Long, StartY As Long, EndY As Long
    'Dim RCVRetNet As Single, RCVRet As Single, RCVCost As Single
    Dim SaMData() As CBA_BTF_SaMData
    Dim SaMContData() As CBA_BTF_SaMData
    Dim strClass As String
    Dim PClass As Long, lCurCG As Long, StartRow As Long
    Dim FLvl() As String, isFirstRun As Boolean
    Dim strDBUtoForecast As String
    
    'On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    YearFrom = lPSYear - 2
    YearTo = lPSYear
    MonthFrom = 1
    MonthTo = 12
    
    CBA_BasicFunctions.CBA_Running "CG/SCG Sales and Margin Forecast Generation"
''    If CBA_BTF_SetupForecastArray.CBA_BTF_SetupForecastArray(YearFrom, YearTo, MonthFrom, MonthTo) = True Then
    If True = True Then
        
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Preparing Report..."
        
        On Error Resume Next
        ''If UBound(parrCGs, 2) < 0 Then
''             sGBD = "ALL"
''            CreateCGsArray RP
        ''End If
        On Error GoTo Err_Routine
''    CreateCGsArray RP                 ' RWCam will have to define cCGs
''        Unload CBA_BTF_frm_Reporting
''        CBA_BasicFunctions.CBA_Running "CG/SCG Sales and Margin Generation"
        If CBA_BTF_SetupForecastArray.CBA_BTF_SetupForecastArray(YearFrom, YearTo, MonthFrom, MonthTo) = True Then
                   
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Preparing Stats..."
            
            On Error Resume Next
            If UBound(parrCGs, 2) < 0 Then
                sGBD = "ALL"
''                CreateCGsArray RP ' RWCam will have to define cCGs
            End If
            On Error GoTo Err_Routine
            ReDim pvRtnArr(0 To UBound(parrCGs, 2) + 1, 0 To CELMS)
            
            For PClass = 1 To 1
                If (lProdClass > 0 And lProdClass = PClass) Or lProdClass = 0 Then
                    pvRtnArr(0, 0) = UBound(parrCGs, 2)
                    For lCG = LBound(parrCGs, 2) To UBound(parrCGs, 2)
                        ''lRowNo = lRowNo + 1
                        lRowNo = lCG + 1
                        If parrCGs(0, lCG) <> lCurCG Then
''                            If lCurCG <> 0 Then
''                                If FCbM(lPSYear, 1, PClass).CGData(lCurCG, 0).Level = "SCG" Or FCbM(lPSYear, 1, PClass).CGData(lCurCG, 0).Level = "CG" Then
''                                    setCGTotal wks(PClass), StartRow, lCurCG, FCbM(lPSYear, 1, PClass).CGData(lCurCG, 0).Level, lRowNo, PClass, False
''                                End If
''                            End If
                            lCurCG = parrCGs(0, lCG)
                            isFirstRun = True
''                            StartRow = lRowNo
''                            TotRCVRetNet = 0: TotRCVRet = 0: TotRCVCost = 0
                        Else
                            If FCbM(lPSYear, 1, PClass).CGData(lCurCG, 0).Level = "CG" Or (FCbM(lPSYear, 1, PClass).CGData(lCurCG, 0).Level = "SCG" And Val(Left(parrCGs(1, lCG), 2)) = 0) Then
                                lRowNo = 1
                                pvRtnArr(0, 0) = 1
''                                GoTo letsdothenextlineintheCGarray
                            End If
                        End If

                        If FCbM(lPSYear, 1, PClass).CGData(lCurCG, 0).Level = "CG" Then
                            pvRtnArr(lRowNo, 1) = "CG: " & parrCGs(0, lCG)
                        ElseIf FCbM(lPSYear, 1, PClass).CGData(lCurCG, 0).Level = "SCG" Then
                            pvRtnArr(lRowNo, 1) = parrCGs(0, lCG) & "/" & parrCGs(1, lCG)
                        Else
                            For b = 1 To 12
                                FCbM(lPSYear, b, PClass).CGData(lCurCG, 0).setLevel "SCG"
                            Next
                            pvRtnArr(lRowNo, 1) = "CG: " & parrCGs(0, lCG) & "-" & parrCGs(1, lCG)

'                                    lRowNo = lRowNo - 1
'                                    GoTo letsdothenextlineintheCGarray
                        End If
                        For Mnth = 1 To 12
                            If lPSYear >= Year(Date) Then
                                If FCbM(lPSYear, Mnth, PClass).CGData(lCurCG, 0).Level = "None" Then
                                      lCG = lCG
                                ElseIf FCbM(lPSYear, Mnth, PClass).CGData(lCurCG, 0).Level = "CG" Then
                                      Call PutDataIntoArray(pvRtnArr, PClass, Mnth, lPSYear, "CG", lCurCG, lRowNo)
                                ElseIf FCbM(lPSYear, Mnth, PClass).CGData(lCurCG, 0).Level = "SCG" Then
                                      parrCGs(1, lCG) = Val(parrCGs(1, lCG))
                                      Call PutDataIntoArray(pvRtnArr, PClass, Mnth, lPSYear, "SCG", lCurCG, lRowNo, parrCGs(1, lCG))
                                End If
                            End If
                        Next Mnth
letsdothenextlineintheCGarray:
                    Next lCG
                End If
            Next PClass
        End If
''        For b = 0 To 10
'''            Debug.Print FCbM(2020, 1, 1).CGData(3, b).Sales().ReForecast & ",";
''            Debug.Print FCbM(2019, 1, 1).CGData(2, b).Level & ",";
''
''        Next
        pcFCbM = FCbM
        Erase FCbM
        If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.CBA_Close_Running
    End If



exitTheReport:
If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
''Erase CBA_SCGData

Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-SetForecastArray", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function PutDataIntoArray(ByRef arr As Variant, ByVal PClass As Long, ByVal Mnth As Long, ByVal reportYear As Long, ByVal FLevel As String, ByVal CG As Long, ByVal RowNum As Long, Optional ByVal scg As Long)
    Dim CBA_Proc As String
    Dim sng0SalesReFC As Single, sng0SalesActual As Single, sng0MarginReFC As Single, sng0MarginActual As Single
    Dim sng1SalesReFC As Single, sng1SalesActual As Single, sng1MarginReFC As Single, sng1MarginActual As Single
''    Dim sngTSalesReFC As Single, sngTSalesActual As Single, sngTMarginReFC As Single, sngTMarginActual As Single
''    Dim sngLSalesReFC As Single, sngLSalesActual As Single, sngLMarginReFC As Single, sngLMarginActual As Single
    Static lRowCount As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    If lRowCount = 0 Then
        lRowCount = pvRtnArr(0, 0)
    End If
    ' Have left out the original format unless its goes into the format figure...
    If FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.ReForecast = 0 Then
        sng0SalesReFC = FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.OriginalForecast
    Else
        sng0SalesReFC = FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.ReForecast
    End If
    If FCbM(reportYear, Mnth, PClass).CGData(CG, scg).MarginDol.ReForecast = 0 Then
        sng0MarginReFC = FCbM(reportYear, Mnth, PClass).CGData(CG, scg).MarginDol.OriginalForecast
    Else
        sng0MarginReFC = FCbM(reportYear, Mnth, PClass).CGData(CG, scg).MarginDol.ReForecast
    End If
    If FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).Sales.ReForecast = 0 Then
        sng1SalesReFC = FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).Sales.OriginalForecast
    Else
        sng1SalesReFC = FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).Sales.ReForecast
    End If
    If FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).MarginDol.ReForecast = 0 Then
        sng1MarginReFC = FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).MarginDol.OriginalForecast
    Else
        sng1MarginReFC = FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).MarginDol.ReForecast
    End If
    
    sng0SalesActual = FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.Actual
    sng0MarginActual = FCbM(reportYear, Mnth, PClass).CGData(CG, scg).MarginDol.Actual
    sng1SalesActual = FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).Sales.Actual
    sng1MarginActual = FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).MarginDol.Actual

  

''
''ForecastRatherThanActuals:
''    ' Single year data
    arr(RowNum, CDFT0 + eCSRF) = arr(RowNum, CDFT0 + eCSRF) + sng0SalesReFC                   ' Return this one into Sales Forecast
    arr(RowNum, CDFT0 + eCSSA) = arr(RowNum, CDFT0 + eCSSA) + sng0SalesActual
    arr(RowNum, CDFT0 + eCMRF) = arr(RowNum, CDFT0 + eCMRF) + sng0MarginReFC                  ' Return this one into Margin Forecast
    arr(RowNum, CDFT0 + eCMSA) = arr(RowNum, CDFT0 + eCMSA) + sng0MarginActual
    arr(RowNum, CDFT1 + eCSRF) = arr(RowNum, CDFT1 + eCSRF) + sng1SalesReFC
    arr(RowNum, CDFT1 + eCSSA) = arr(RowNum, CDFT1 + eCSSA) + sng1SalesActual
    arr(RowNum, CDFT1 + eCMRF) = arr(RowNum, CDFT1 + eCMRF) + sng1MarginReFC
    arr(RowNum, CDFT1 + eCMSA) = arr(RowNum, CDFT1 + eCMSA) + sng1MarginActual

''    ' Total of all Product SCGs
    arr(lRowCount, CDFTT + eCSRF) = arr(lRowCount, CDFTT + eCSRF) + sng0SalesReFC
    arr(lRowCount, CDFTT + eCSSA) = arr(lRowCount, CDFTT + eCSSA) + sng0SalesActual
    arr(lRowCount, CDFTT + eCMRF) = arr(lRowCount, CDFTT + eCMRF) + sng0MarginReFC
    arr(lRowCount, CDFTT + eCMSA) = arr(lRowCount, CDFTT + eCMSA) + sng0MarginActual
    
    arr(lRowCount, CDFTL + eCSRF) = arr(lRowCount, CDFTL + eCSRF) + sng1SalesReFC
    arr(lRowCount, CDFTL + eCSSA) = arr(lRowCount, CDFTL + eCSSA) + sng1SalesActual
    arr(lRowCount, CDFTL + eCMRF) = arr(lRowCount, CDFTL + eCMRF) + sng1MarginReFC
    arr(lRowCount, CDFTL + eCMSA) = arr(lRowCount, CDFTL + eCMSA) + sng1MarginActual


    arr(RowNum, CDFT0 + eCSYoY) = g_DivZero((arr(RowNum, CDFT0 + eCSSA) - arr(RowNum, CDFT1 + eCSSA)), arr(RowNum, CDFT1 + eCSSA))       ' Return this one into Sales YoY %
    arr(RowNum, CDFT0 + eCMYoY) = g_DivZero((arr(RowNum, CDFT0 + eCMSA) - arr(RowNum, CDFT1 + eCMSA)), arr(RowNum, CDFT1 + eCMSA))       ' Return this one into Margin YoY %
    
    arr(lRowCount, CDFTT + eCSYoY) = g_DivZero((arr(lRowCount, CDFTT + eCSSA) - arr(lRowCount, CDFTL + eCSSA)), arr(lRowCount, CDFTL + eCSSA))    ' Return this one into Sales Total YoY %
    arr(lRowCount, CDFTT + eCMYoY) = g_DivZero((arr(lRowCount, CDFTT + eCMSA) - arr(lRowCount, CDFTL + eCMSA)), arr(lRowCount, CDFTL + eCMSA))    ' Return this one into Margin Total YoY %
'
''    arr(RowNum, CDFT + eCSYoY) = g_DivZero((arr(RowNum, CDFT + eCSSA) - arr(RowNum, CDFT + eCSSA)), arr(RowNum, CDFT + eCSSA))                          ' YoY Actual Sales???
''    arr(RowNum, CDFT + eCMYoY) = g_DivZero((arr(RowNum, CDFT + eCMSA) - arr(RowNum, CDFT + eCMSA)), arr(RowNum, CDFT + eCMSA))                          ' YoY Actual Margin???
''
   
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-PutDataIntoArray", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Property Get lPSYear() As Long: lPSYear = plPSYear: End Property
Private Property Let lPSYear(ByVal lNewValue As Long): plPSYear = lNewValue: End Property

Private Property Get lPEYear() As Long: lPEYear = plPEYear: End Property
Private Property Let lPEYear(ByVal lNewValue As Long): plPEYear = lNewValue: End Property

Private Property Get sBD() As String: sBD = psBD: End Property
Private Property Let sBD(ByVal sNewValue As String): psBD = sNewValue: End Property

Private Property Get sGBD() As String: sGBD = psGBD: End Property
Private Property Let sGBD(ByVal sNewValue As String): psGBD = sNewValue: End Property

Private Property Get lCG() As Long: lCG = plCG: End Property
Private Property Let lCG(ByVal lNewValue As Long): plCG = lNewValue: End Property

Private Property Get lSCG() As Long: lSCG = plSCG: End Property
Private Property Let lSCG(ByVal lNewValue As Long): plSCG = lNewValue: End Property

Private Property Get lProdClass() As Long: lProdClass = plProdClass: End Property

Private Property Let lProdClass(ByVal lNewValue As Long): plProdClass = lNewValue: End Property
