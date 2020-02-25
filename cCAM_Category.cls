VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCAM_Category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit                     ' cCAM_Category
Implements iCBA_Doc_Data
Private pcPGrp As cCBA_ProdGroup
Private pcFCast As cCBA_Forecasting
Private parrFcase() As Variant
Private psdCategoryList As Scripting.Dictionary
Private psdLineCountDic As Scripting.Dictionary
Private psdUserList As Scripting.Dictionary
Private psdEtoNUList As Scripting.Dictionary
Private psdMSegs As Scripting.Dictionary
Private psdNData As Scripting.Dictionary

Private pvBaseUsers() As Variant
Private psCategoryName As String
Private pbACG As Boolean
Private pbIsActive As Boolean
Private pbCRTrialPerf As Boolean
Private pbLineCount As Boolean
Private pbCatStrat As Boolean
Private pbCorpResp As Boolean
Private pbExecSum As Boolean
Private pbCRPerf As Boolean
Private pbForecast As Boolean
Private pbSuppBase As Boolean
Private pbCatPlano As Boolean
Private pbCatPerfOver As Boolean
Private pbSeasPerf As Boolean
Private pbInterComp As Boolean
Private pbCRPerfProds As Boolean
Private pbCommUpd As Boolean
Private pbTotCatPerf As Boolean
Private pbSpecPerf As Boolean
Private pbCompAnal As Boolean
Private pbQA As Boolean
Private pbKPIs As Boolean
Private pbMarkOver As Boolean
Private pbWastage As Boolean
Private pbCatRevActions As Boolean
Private pbCRTrialProdList As Boolean
Private pbRdyToRun As Boolean
Private pdtDateFrom As Date
Private pdtDateTo As Date
Private plDoc_ID As Long
Private peDocumentType As e_DocuType
Private psBuildCode As String
Private fMSS As fCAM_MSegSelect
Private fPTM As fCAM_ProdToMSeg
Private fLC As fCAM_LineCount
Private Const lMINUS_YY As Long = 3
Private bChartsBuilt(1 To 5) As Boolean ' this is naughty and we need to fix it
Private peDefaultMSegtype As e_MSegType

    ' Holds specific system category data relevant for the non reporting part of CAMERA.
    ' That includes instances of other data objects, such as cCBA_ProdGroup etc.
    ' This is the trigger that not only houses the data for the document but allows it to have bespoke functionailty embedded into the CBA_Document direclty,
    '   effectivly eliminating the need for the Runtime to be called for the link between the document specific data and the processing of that data.
    ' The runtime will still be needed for non document specific information such as full category lists etc,
    '   but relating to what should appear in a cell or button functionality etc that can all sit in the Data Object

'Implement Relevent Code                                        ' NOTE - HAS TO BE iCBA_Doc_Data AS THIS IS THE PREFIX AND THE NAME OF THE CLASS
Public Function iCBA_Doc_Data_DocumentType() As e_DocuType
    'i think this may need to refer to the system rather than the document in the system, but lets leave it as docutype for the min
    iCBA_Doc_Data_DocumentType = eDocumentType
End Function
Public Function iCBA_Doc_Data_SetDocID(ByVal DocID As Long) As Boolean
Dim v As Variant
    lDoc_ID = DocID
    If Not sdLineCountDic Is Nothing Then
        For Each v In sdLineCountDic
            sdLineCountDic(v).lDoc_ID = DocID
        Next
    End If
    cPGrp.lDoc_ID = DocID
    iCBA_Doc_Data_SetDocID = True
End Function


Public Function iCBA_Doc_Data_GetVBACode(Optional ByVal sAdditionalCode As String) As Scripting.Dictionary
    Dim dic As Scripting.Dictionary, sAddinName As String
    Dim strCode As String, DocName As String
    On Error GoTo Err_Routine
    CBA_Error = ""
    Set dic = New Scripting.Dictionary
    sAddinName = mCAM_Runtime.cDocHolder.sAddinName
    
    'WORKBOOK CODE (always needed to unlink the worksheet from the VB Code)
    strCode = "Private Const CBAddin = """ & sAddinName & """" & Chr(10)
    strCode = strCode & "Private Const PRC = """ & "!mCAM_Runtime.ControlObject" & """" & Chr(10)
    strCode = strCode & "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & Chr(10)
    strCode = strCode & "Dim sht As Worksheet" & Chr(10)
    strCode = strCode & "   For Each sht In ThisWorkbook.Worksheets" & Chr(10)
    strCode = strCode & "       Run CBAddin & PRC , ThisWorkbook.Name & ""-"" & sht.CodeName, , ""Close""" & Chr(10)
    strCode = strCode & "   Next" & Chr(10)
    strCode = strCode & "   ThisWorkbook.Saved = True" & Chr(10)
    strCode = strCode & "End Sub" & Chr(10)
    
    'WORKBOOK CODE Addition
    dic.Add "ThisWorkbook", CStr(strCode)
    
    'WORKSHEET CODE Additions
    If eDocumentType = eForecast Or eDocumentType = eCoreRangeCategoryReview Then
        strCode = "Private Const CBAddin = """ & sAddinName & """" & Chr(10)
        strCode = strCode & "Private Const PRC = """ & "!mCAM_Runtime.ControlObject" & """" & Chr(10)
        strCode = strCode & "Sub ActionButton(lRow As Long, lCol As Long)" & Chr(10)
        strCode = strCode & "    Dim sError As String, btn As Button" & Chr(10)
        strCode = strCode & "    On Error GoTo Err_Routine" & Chr(10)
        strCode = strCode & "    Set btn = ActiveSheet.Buttons(Application.Caller)" & Chr(10)
        strCode = strCode & "    Run CBAddin & PRC  , ""WorkbookRef.Name-WorksheetRef.CodeName"",, ""RefreshFC""" & Chr(10)
        strCode = strCode & "    On Error Resume Next" & Chr(10)
        strCode = strCode & "    Exit Sub" & Chr(10)
        strCode = strCode & "Err_Routine:" & Chr(10)
        strCode = strCode & "    sError = Err.Number & ""-"" & Err.Description" & Chr(10)
        strCode = strCode & "    Debug.Print sError" & Chr(10)
        strCode = strCode & "    If Run(CBAddin & ""!CBA_BasicFunctions.g_Get_Dev_Sts"", ""DevIP"") = ""Y"" Then" & Chr(10)
        strCode = strCode & "         Stop" & Chr(10)
        strCode = strCode & "         Resume Next" & Chr(10)
        strCode = strCode & "     Else" & Chr(10)
        strCode = strCode & "         MsgBox ""This worksheet is no longer attached to the Application; Please close this worksheet and restart the application""" & Chr(10)
        strCode = strCode & "     End If" & Chr(10)
        strCode = strCode & "End Sub" & Chr(10)
        dic.Add "ThisWorksheet", CStr(strCode)
    End If
    If eDocumentType = eCoreRangeCategoryReview Then
        strCode = " Private Sub Worksheet_Change(ByVal Target As Range)" & Chr(10)
        strCode = strCode & "    On Error GoTo Err_Routine" & Chr(10)
        strCode = strCode & "    Dim bfound As Boolean, lRet As Long, lIter As Long, lRow As Long, lCol As Long, vValue, sValue As String" & Chr(10)
        strCode = strCode & "    Dim sError As String, sRtn1 as String, lUT_AF as Long, lUT_MH as Long, lUT_WC as Long, rngR As Range" & Chr(10)
        strCode = strCode & "    Static bIsRunning as Boolean" & Chr(10)
        strCode = strCode & "    If bIsRunning Then Exit Sub" & Chr(10)
        strCode = strCode & "    bIsRunning = True" & Chr(10)
        strCode = strCode & "    If Run(CBAddin & ""!CBA_BasicFunctions.g_SetupIP"" ,""RenderDoc"", 0) = False Then" & Chr(10)
        strCode = strCode & "        Call Run(CBAddin & PRC  ,  ""WorkbookRef.Name-WorksheetRef.CodeName"",, ""Unprotect"")" & Chr(10)
        strCode = strCode & "        lRow = Target.Row" & Chr(10)
        strCode = strCode & "        lCol = Target.Column" & Chr(10)
        strCode = strCode & "        vValue = Target.Value" & Chr(10)
        strCode = strCode & "        sValue = CStr(vValue)" & Chr(10)
        strCode = strCode & "        If sValue = ""Unknown Value"" Then" & Chr(10)
        strCode = strCode & "            On Error Resume Next" & Chr(10)
        strCode = strCode & "            sValue = Target.Cells(1, 1).Value" & Chr(10)
        strCode = strCode & "            On Error GoTo Err_Routine" & Chr(10)
        strCode = strCode & "        End If" & Chr(10)
        strCode = strCode & "        lUT_MH = Run(CBAddin & PRC  ,  ""WorkbookRef.Name-WorksheetRef.CodeName"",, ""UT_Merge_Height"",,lRow,lCol)" & Chr(10)
        strCode = strCode & "        If lUT_MH < Target.Rows.Count Then" & Chr(10)
        strCode = strCode & "            For lIter = lRow To (lRow + Target.Rows.Count - 1) Step lUT_MH" & Chr(10)
        strCode = strCode & "                call Run(CBAddin & PRC  , ""WorkbookRef.Name-WorksheetRef.CodeName"",, ""UpdateRC"",sValue,lIter,lCol)" & Chr(10)
        strCode = strCode & "            Next" & Chr(10)
        strCode = strCode & "        Else" & Chr(10)
        strCode = strCode & "            lUT_AF = Run(CBAddin & PRC  ,  ""WorkbookRef.Name-WorksheetRef.CodeName"",, ""AutoFit"",,lRow,lCol)" & Chr(10)
        strCode = strCode & "            lUT_WC = Run(CBAddin & PRC  ,  ""WorkbookRef.Name-WorksheetRef.CodeName"",, ""UpdateRC"",sValue,lRow,lCol)" & Chr(10)
        strCode = strCode & "        End If" & Chr(10)
        strCode = strCode & "        If lUT_AF > 0 Then" & Chr(10)
        strCode = strCode & "            If lUT_WC > 0 Then Target.Font.ColorIndex = 3 Else Target.Font.ColorIndex = 1" & Chr(10)
        strCode = strCode & "            Set rngR = Range(""AFVAL"" & lRow)" & Chr(10)
        strCode = strCode & "            rngR.Value = sValue" & Chr(10)
        strCode = strCode & "            rngR.WrapText = True" & Chr(10)
        strCode = strCode & "            rngR.EntireRow.AutoFit" & Chr(10)
        strCode = strCode & "        End If" & Chr(10)
        strCode = strCode & "        Call Run(CBAddin & PRC  ,  ""WorkbookRef.Name-WorksheetRef.CodeName"",, ""Protect"")" & Chr(10)
        strCode = strCode & "    End If" & Chr(10)
        strCode = strCode & "    On Error Resume Next" & Chr(10)
        strCode = strCode & "    bIsRunning = False" & Chr(10)
        strCode = strCode & "    Exit Sub" & Chr(10)
        strCode = strCode & "Err_Routine:" & Chr(10)
        strCode = strCode & "    sError = Err.Number & ""-"" & Err.Description" & Chr(10)
        strCode = strCode & "    Debug.Print sError" & Chr(10)
        strCode = strCode & "    If Run(CBAddin & ""!CBA_BasicFunctions.g_Get_Dev_Sts"", ""DevIP"") = ""Y"" Then" & Chr(10)
        strCode = strCode & "         Stop" & Chr(10)
        strCode = strCode & "         Resume Next" & Chr(10)
        strCode = strCode & "     Else" & Chr(10)
        strCode = strCode & "         MsgBox ""This worksheet is no longer attached to the Application; Please close this worksheet and restart the application""" & Chr(10)
        strCode = strCode & "     End If" & Chr(10)
        strCode = strCode & "End Sub" & Chr(10)
        dic.Add "ThisWorksheet1", CStr(strCode)
    End If

Exit_Routine:

    On Error Resume Next
    Set iCBA_Doc_Data_GetVBACode = dic
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.getDocVBACode", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function


Public Function iCBA_Doc_Data_GetGrpNo(ByVal ProcedureName As String) As Long
    Dim cLC As cCAM_LineCount_Data
    Dim v As Variant
    Dim S1 As Single, S2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    Dim N As cCBA_NielsenData
    Dim MSegType  As e_MSegType
    Dim P As cCBA_Prod
    Dim col As Collection
    
    Select Case ProcedureName
        Case "CoreRangeTrialProductList"
            Set cLC = sdLineCountDic(eMAT)
            S1 = 0
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eCurTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eBrandSucTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eSucTrial Then
                    S1 = S1 + 1
                End If
            Next
        Case "CoreRangePerfByProduct"
''            Set col = cPGrp.getProdListing
''            For Each v In col
''                Set P = cPGrp.getProdObject(Mid(v, 1, InStr(1, v, "-") - 1))
''                If P.lPClass = 1 Then If P.getPOSdata(cPGrp.dtMnthFrom, cPGrp.dtMnthTo, True, eNotUSW) > 1000 Then S1 = S1 + 1
''            Next
            S1 = 10                             ' @RWCam @TP T will fix up
        Case "MarketOV1", "MarketOV2"
            If ProcedureName = "MarketOV1" Then
                Set sdNData = mCBA_Nielsen.GetNielsenData(Year(dtDateTo), Month(dtDateTo), IIf(eDefaultMSegtype = eScanData, True, False), cPGrp.colCGs, bACG)
                For S1 = 1 To 11
                    If sdNData Is Nothing Then Set sdNData = mCBA_Nielsen.GetNielsenData(Year(DateAdd("M", -S1, dtDateTo)), Month(DateAdd("M", -S1, dtDateTo)), IIf(eDefaultMSegtype = eScanData, True, False), cPGrp.colCGs, bACG)
                Next
            End If
            MSegType = eDefaultMSegtype
            S1 = 0
            For Each v In psdMSegs
                Set N = psdMSegs(v)
                If MSegType = eHomescan And N.IsHomescan Then S1 = S1 + 1
                If MSegType = eScanData And N.isScanData Then S1 = S1 + 1
                If MSegType = eManual1 And N.IsManual Then S1 = S1 + 1
            Next
    End Select
    iCBA_Doc_Data_GetGrpNo = CLng(S1)
End Function

Public Function iCBA_Doc_Data_GetUDTData(ByRef cUDT As cCBA_UDT, Optional ByVal ReqType As String) As Boolean
    ' Will process the UT_Procedure and UT_Name UDT fields to return values that are calculated or obtained from a data object
    Dim sDates As String, sString As String, CP As cCBA_Prod, v As Variant, sSep As String
    Select Case ReqType
    Case "1stFormat"
        Select Case cUDT.sUT_Procedure
        Case "OAHeading"
            sDates = " for " & g_FixDate(dtDateFrom) & " to " & g_FixDate(dtDateTo)
            Select Case eDocumentType
            Case e_DocuType.eToplineCategoryPerformance                                             ' Items 2 (chart) and 3
                cUDT.sUT_Default_Value = UCase("Topline Category Performance Report") & sDates
            Case e_DocuType.eMarketOverview                                                         ' Items 4/1 and 4/2
                cUDT.sUT_Default_Value = UCase("Market Overview Report") & sDates
            Case e_DocuType.eCoreRangeCategoryReview
                cUDT.sUT_Default_Value = UCase("Core Range Category Review") & sDates
            Case e_DocuType.eCoreRangePerformance                                                   ' Items 5, 6 and 9
                cUDT.sUT_Default_Value = UCase("Core Range Performance Report") & sDates
            Case e_DocuType.eSpecSeasPerformance                                                    ' Items 7 and 8
                cUDT.sUT_Default_Value = UCase("Specials and Seasonal Performance") & sDates
            Case e_DocuType.eLineCountOverviewReport                                                ' Item 10
                cUDT.sUT_Default_Value = UCase("Line Count Overview Report") & sDates
            Case e_DocuType.eForecast                                                               ' Item 17
                cUDT.sUT_Default_Value = UCase("Forecast Report") & sDates
            Case e_DocuType.eCoreRangeProductListing                                                ' Items 20 and 21
                cUDT.sUT_Default_Value = UCase("Core Range Product Listing") & sDates
            End Select
        Case "CatHeading"
            sString = IIf(bACG, " AGC ", " LCG "): sSep = ""
            For Each v In sdCategoryList.Keys
                Set CP = sdCategoryList(v)
                If CP.sACat = "" And bACG = False Then
                    sString = sString & sSep & CStr(Format(CP.lLegCG, "000") & "|" & Format(CP.lLegSCG, "000"))
                    sSep = ", "
                ElseIf CP.lACatNum > 0 And bACG = True Then
                    sString = sString & sSep & CStr(Format(CP.lACGNum, "000") & "|" & Format(CP.lASCGNum, "000"))
                    sSep = ", "
                End If
            Next
            cUDT.sUT_Default_Value = "Category - Sub Category" & sString
        Case "CGName"
            cUDT.sUT_Default_Value = "CATEGORY REVIEW: (TBC)"
        Case "UserName+Date"
            cUDT.sUT_Default_Value = CBA_User & " as of  *" & g_FixDate(Now(), CBA_DMYHN & " am/pm") & "*"
        Case "Year"
            If cUDT.lUT_Grp_No > 0 Then
                cUDT.sUT_Default_Value = CStr(Year(dtDateTo) - lMINUS_YY + cUDT.lUT_Grp_No)
            Else
                cUDT.sUT_Default_Value = "MAT " & Format(DateSerial(Year(dtDateTo), Month(dtDateTo), Day(dtDateTo)), "mm/yyyy")
            End If
        Case "CatPerfOverview"                                                                          ' @2. Category Performance Overview (in the Topline Category Performance Report)
            cUDT.sUT_Default_Value = CStr(f_CatPerfOverview(cUDT.sUT_Name, cUDT.lUT_Grp_No))
        Case "TotalCatPerf", "TotalCatPerfT", "TotalCatPerfTT"                                          ' @3. Total Category Performance (in the Topline Category Performance Report)
            cUDT.sUT_Default_Value = CStr(f_TotalCatPerf(cUDT.sUT_Name, cUDT.lUT_Grp_No, IIf(cUDT.sUT_Procedure = "TotalCatPerfTT", "TT", "")))
        Case "MarketOV1", "MarketOV1T", "MarketOV1PY"                                                   ' @4.1. Market Overview, Table 1: Market vs ALDI overall
            cUDT.sUT_Default_Value = CStr(f_MarketOV1(cUDT.sUT_Name, cUDT.lUT_Grp_No, Right(cUDT.sUT_Procedure, 2)))
        Case "MarketOV2", "MarketOV2T", "MarketOV2PY"                                                   ' @4.2. Market Overview, Table 2: Market Breakdown
            cUDT.sUT_Default_Value = CStr(f_MarketOV2(cUDT.sUT_Name, cUDT.lUT_Grp_No, Right(cUDT.sUT_Procedure, 2)))
        Case "CoreRangePerf", "CoreRangePerfT"                                                          ' @5. Core Range Performance
            cUDT.sUT_Default_Value = CStr(f_CoreRangePerf(cUDT.sUT_Name, cUDT.lUT_Grp_No))
        Case "CoreRangeTrialPerf", "CoreRangeTrialPerfT"                                                ' @6. Core Range Trial Performance
            cUDT.sUT_Default_Value = CStr(f_CoreRangeTrialPerf(cUDT.sUT_Name, cUDT.lUT_Grp_No))
        Case "SeasPerf", "SeasPerfT"                                                                    ' @7. Seasonal Performance
            cUDT.sUT_Default_Value = CStr(f_SeasPerf(cUDT.sUT_Name, cUDT.lUT_Grp_No))
        Case "SpecPerf", "SpecPerfT"                                                                    ' @8. Specials Performance
            cUDT.sUT_Default_Value = CStr(f_SpecPerf(cUDT.sUT_Name, cUDT.lUT_Grp_No))
        Case "LCOverView", "LCOverViewT"
            cUDT.sUT_Default_Value = CStr(f_LCOverView(cUDT.sUT_Name, cUDT.lUT_Grp_No))
        Case "Wastage", "WastageT"                                                                      ' @9. Wastage
            cUDT.sUT_Default_Value = CStr(f_Wastage(cUDT.sUT_Name, cUDT.lUT_Grp_No))
        Case "Forecast", "ForecastT"                                                                    ' @17. Forecasting
            cUDT.sUT_Default_Value = CStr(f_Forecast(cUDT.sUT_Name, cUDT.lUT_Grp_No))
        Case "CoreRangeTrialProductList"                                                                ' @20. Core Range Trial Product List
            cUDT.sUT_Default_Value = CStr(f_CoreRangeTrialProductList(cUDT.sUT_Name, cUDT.lUT_Grp_No))
        Case "CoreRangePerfByProduct"                                                                   ' @21. Core Range Performance by Product
            cUDT.sUT_Default_Value = CStr(f_CoreRangePerfByProduct(cUDT.sUT_Name, cUDT.lUT_Grp_No))
        Case ""
        Case ""
        End Select
    Case ""
    Case ""
    Case ""
    Case ""
    Case ""
    Case ""
    End Select
    iCBA_Doc_Data_GetUDTData = True
End Function

Private Function f_CatPerfOverview(ByVal sName As String, ByVal lGrp_No As Long, Optional sSuff As String = "")
    ' Handle @2. Category Performance Overview (in the Topline Category Performance Report)
    Dim CP As cCBA_ProdGroup, cLC As cCAM_LineCount_Data, dtTo As Date, dtFrom As Date, dtToM As Date, dtFromM As Date, lLC As Long, v As Variant
    Dim S1 As Single, S2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    Dim dt1 As Date
    Dim a As Long, grpNo As Long, pcs As Long
    Dim CGDic As Scripting.Dictionary
    Dim SCGDic As Scripting.Dictionary
    On Error GoTo Err_Routine
    CBA_Error = ""
    sName = g_Left(sName, Len(CStr(lGrp_No)))
    If lGrp_No = 13 And sName = "LineCntRev" Then
        BuildCatPerfOverViewCharts "toplineandforecast", bChartsBuilt(1)
        BuildCatPerfOverViewCharts "cyrevenue", bChartsBuilt(2)
        BuildCatPerfOverViewCharts "linecount", bChartsBuilt(3)
        BuildCatPerfOverViewCharts "cygrossprofit", bChartsBuilt(4)
        BuildCatPerfOverViewCharts "pygrossprofit", bChartsBuilt(5)
    End If
    If lGrp_No > 13 Then Exit Function
    dtFrom = dtDateFrom: dtTo = dtDateTo
    If lGrp_No = 1 Then grpNo = 99 Else grpNo = lGrp_No - 1 'This is here as the rist row needs to have the headers, just like the row to the left needs to be the labels
    Set CP = cPGrp
    Select Case sName
        Case "PClassLabel"
            If grpNo = 1 Or grpNo = 5 Or grpNo = 9 Then
                f_CatPerfOverview = "Core Range"
            ElseIf grpNo = 2 Or grpNo = 6 Or grpNo = 10 Then
                f_CatPerfOverview = "Specials"
            ElseIf grpNo = 3 Or grpNo = 7 Or grpNo = 11 Then
                f_CatPerfOverview = "Seasonal"
            Else
                f_CatPerfOverview = ""
            End If
        Case "LineCntRev"
            If grpNo = 99 Then f_CatPerfOverview = "Line Count": Exit Function
            If grpNo = 4 Then f_CatPerfOverview = "CY Revenue": Exit Function
            If grpNo = 8 Then f_CatPerfOverview = "PY Revenue": Exit Function
            
            If grpNo <= 3 And grpNo > 0 Then
                Set cLC = sdLineCountDic(eMAT)
                For Each v In cLC.sdProductAllocation
                    Select Case grpNo
                        Case 1
                            If cLC.sdProductAllocation(v) = e_LineCountType.eBranded Or cLC.sdProductAllocation(v) = e_LineCountType.eBrandSucTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eCore Or cLC.sdProductAllocation(v) = e_LineCountType.eSucTrial Then S1 = S1 + 1
                        Case 2
                            If cLC.sdProductAllocation(v) = e_LineCountType.eSeasonal Then S1 = S1 + 1
                        Case 3
                            If cLC.sdProductAllocation(v) = e_LineCountType.eSpecial Then S1 = S1 + 1
                    End Select
                Next
                f_CatPerfOverview = S1
            ElseIf grpNo <= 7 And grpNo >= 5 Then
                If grpNo = 5 Then s3 = 1
                If grpNo = 6 Then s3 = 5
                If grpNo = 7 Then s3 = 4
                S1 = CP.getPOS(dtFrom, dtTo, eRetail, CLng(s3))
                S2 = CP.getPOS(dtFrom, dtTo, eRetail)
                If S2 = 0 Then S1 = 0 Else S1 = S1 / S2
            ElseIf grpNo <= 11 And grpNo >= 9 Then
                If grpNo = 9 Then s3 = 1
                If grpNo = 10 Then s3 = 5
                If grpNo = 11 Then s3 = 4
                S1 = CP.getPOS(dtFrom, dtTo, eRetail, CLng(s3))
                S2 = CP.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eRetail, CLng(s3))
                If S2 = 0 Then S1 = 0 Else S1 = (S1 - S2) / S2
            End If
            f_CatPerfOverview = S1
        Case "Labels"
            If grpNo < 99 Then
                S2 = Year(dtFrom)
                S1 = Month(dtFrom) + grpNo - 1
                If S1 > 12 Then S1 = 1: S2 = S2 + 1
                f_CatPerfOverview = CStr(Format(DateSerial(S2, S1, 1), "MMM"))
            End If
            
        Case "Stat1MM"
            If grpNo = 99 Then f_CatPerfOverview = "CY POS $": Exit Function
            S2 = Year(dtFrom)
            S1 = Month(dtFrom) + grpNo - 1
            If S1 > 12 Then S1 = 1: S2 = S2 + 1
            f_CatPerfOverview = CP.getPOS(DateSerial(S2, S1, 1), DateSerial(S2, S1 + 1, 0), eRetail)
        Case "Stat2MM"
            If grpNo = 99 Then f_CatPerfOverview = "PY POS $": Exit Function
            S2 = Year(dtFrom) - 1
            S1 = Month(dtFrom) + grpNo - 1
            If S1 > 12 Then S1 = 1: S2 = S2 + 1
            f_CatPerfOverview = CP.getPOS(DateSerial(S2, S1, 1), DateSerial(S2, S1 + 1, 0), eRetail)
        Case "Stat3MM"
            If grpNo = 99 Then f_CatPerfOverview = "MAT Forecast": Exit Function
            If bACG = True Then Debug.Print "ACG not supported in Forecasting system":  f_CatPerfOverview = "": Exit Function
            Set CGDic = New Scripting.Dictionary
            Set SCGDic = New Scripting.Dictionary
            S2 = Year(dtFrom)
            S1 = Month(dtFrom) + grpNo - 1
            If S1 > 12 Then S1 = 1: S2 = S2 + 1
            'So: for each group no is a month
            'now we know the year and month, we can address the forecast data
            'we need to only pull the product classes that are forecasted at CG level once, but each SCG must be included.
            'To achieve this we create 2 dictionaries. If it goes in the CG dic, then its just grabbing the numkbebr there and then
            'If its SCG then it will need to grab iteach time it gors round the CategoryList Loop
            For Each v In psdCategoryList
               'Debug.Print CLng(Mid(v, 1, 3)), CLng(Mid(v, 4, 3))
                If CGDic.Exists(CLng(Mid(v, 1, 3))) = False And SCGDic.Exists(CLng(Mid(v, 1, 3))) = False Then
                    For pcs = 1 To 4
                        If cFCast.getForecast(S2, S1, pcs, CLng(Mid(v, 1, 3)), 0).Level = "SCG" Then
                            If SCGDic.Exists(CLng(Mid(v, 1, 3))) Then SCGDic(CLng(Mid(v, 1, 3))) = SCGDic(CLng(Mid(v, 1, 3))) & "|" & CStr(pcs) Else SCGDic.Add CLng(Mid(v, 1, 3)), CStr(pcs)
                        Else
                            If CGDic.Exists(CLng(Mid(v, 1, 3))) Then CGDic(CLng(Mid(v, 1, 3))) = CGDic(CLng(Mid(v, 1, 3))) & "|" & CStr(pcs) Else CGDic.Add CLng(Mid(v, 1, 3)), CStr(pcs)
                            s3 = s3 + cFCast.getForecast(S2, S1, pcs, CLng(Mid(v, 1, 3)), 0).Sales.ReForecast
                        End If
                    Next
                End If
                For pcs = 1 To 4
                    If InStr(1, SCGDic(CLng(Mid(v, 1, 3))), CStr(pcs)) > 0 Then s3 = s3 + cFCast.getForecast(S2, S1, pcs, CLng(Mid(v, 1, 3)), CLng(Mid(v, 4, 3))).Sales.ReForecast
                Next
            Next
            f_CatPerfOverview = s3
        Case "Stat4MM"
            If grpNo = 99 Then f_CatPerfOverview = "PY MAT Forecast": Exit Function
            If bACG = True Then Debug.Print "ACG not supported in Forecasting system":  f_CatPerfOverview = "": Exit Function
            Set CGDic = New Scripting.Dictionary
            Set SCGDic = New Scripting.Dictionary
            S2 = Year(dtFrom) - 1
            S1 = Month(dtFrom) + grpNo - 1
            If S1 > 12 Then S1 = 1: S2 = S2 + 1
            'So: same as Stat3MM
            For Each v In psdCategoryList
                If CGDic.Exists(CLng(Mid(v, 1, 3))) = False And SCGDic.Exists(CLng(Mid(v, 1, 3))) = False Then
                    For pcs = 1 To 4
                        If cFCast.getForecast(S2, S1, pcs, CLng(Mid(v, 1, 3)), 0).Level = "SCG" Then
                            If SCGDic.Exists(CLng(Mid(v, 1, 3))) Then SCGDic(CLng(Mid(v, 1, 3))) = SCGDic(CLng(Mid(v, 1, 3))) & "|" & CStr(pcs) Else SCGDic.Add CLng(Mid(v, 1, 3)), CStr(pcs)
                        Else
                            If CGDic.Exists(CLng(Mid(v, 1, 3))) Then CGDic(CLng(Mid(v, 1, 3))) = CGDic(CLng(Mid(v, 1, 3))) & "|" & CStr(pcs) Else CGDic.Add CLng(Mid(v, 1, 3)), CStr(pcs)
                            s3 = s3 + cFCast.getForecast(S2, S1, pcs, CLng(Mid(v, 1, 3)), 0).Sales.ReForecast
                        End If
                    Next
                End If
                For pcs = 1 To 4
                    If InStr(1, SCGDic(CLng(Mid(v, 1, 3))), CStr(pcs)) > 0 Then s3 = s3 + cFCast.getForecast(S2, S1, pcs, CLng(Mid(v, 1, 3)), CLng(Mid(v, 4, 3))).Sales.ReForecast
                Next
            Next
            f_CatPerfOverview = s3
        Case "GrossProfit"
            If grpNo = 99 Then
                f_CatPerfOverview = "CY Gross Profit": Exit Function
            ElseIf grpNo = 4 Then
                f_CatPerfOverview = "PY Gross Profit": Exit Function
            ElseIf grpNo <= 3 And grpNo > 0 Then
                If grpNo = 1 Then s3 = 1
                If grpNo = 2 Then s3 = 5
                If grpNo = 3 Then s3 = 4
                S1 = CP.getRCVMargin(eRCVMarginPercent, dtFrom, dtTo, CLng(s3))
                S1 = S1
            ElseIf grpNo <= 7 And grpNo >= 5 Then
                If grpNo = 5 Then s3 = 1
                If grpNo = 6 Then s3 = 5
                If grpNo = 7 Then s3 = 4
                S1 = CP.getRCVMargin(eRCVMarginPercent, DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), CLng(s3))
            End If
            f_CatPerfOverview = S1
        Case Else
            f_CatPerfOverview = "Not Found"
    End Select

Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.f_CatPerfOverview", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function
Private Sub BuildCatPerfOverViewCharts(ByVal ChartToBuild As String, ByRef bRefresh As Boolean)
Dim RangeIncLabels As Range, RangeDataOnly As Range
Dim cht As Chart
Dim shp As Shape
Dim v As Variant
Dim ColW As Single
Dim RowH As Single
Dim pt As Point
Dim ad As Long

    If eDocumentType = e_DocuType.eCoreRangeCategoryReview Then ad = 29 Else ad = 0
    ColW = Columns(1).Width
    RowH = Rows(1).Height
    Select Case LCase(ChartToBuild)
        Case "toplineandforecast"
            If bRefresh = True Then Exit Sub
            
            Set RangeIncLabels = Range(Cells(12 + ad, 1), Cells(24 + ad, 5))
            Set RangeDataOnly = Range(Cells(13 + ad, 2), Cells(24 + ad, 5))
            Set cht = mCAM_Runtime.cDocHolder.AddChart(13, 12 + ad, "G:\Central_Buying\General\4_Administration\Z - Buying Systems Analyst\LIVE EXCEL CHART TEMPLATES\cht_CAM_CategoryLines.crtx", lDoc_ID, RangeIncLabels, RangeDataOnly, "MMM", True, "#", False, 11.1627907 * ColW, 12.53571427 * RowH)
            cht.Legend.Top = cht.PlotArea.Height - (cht.PlotArea.Height * 0.01)
        Case "cyrevenue"
            If bRefresh = True Then Exit Sub
            'Set RangeIncLabels = Range(Cells(16, 6), Cells(19, 7))
            Set RangeIncLabels = Range(Range(Cells(16 + ad, 6), Cells(19 + ad, 6)).Address & "," & Range(Cells(16 + ad, 8), Cells(19 + ad, 8)).Address)
            Set cht = mCAM_Runtime.cDocHolder.AddChart(1, 12 + ad, "G:\Central_Buying\General\4_Administration\Z - Buying Systems Analyst\LIVE EXCEL CHART TEMPLATES\cht_CAM_Revenue.crtx", lDoc_ID, RangeIncLabels, , , , , , 12.1317829 * ColW, 12.43511902 * RowH)
            'cht.Select
            cht.SetElement msoElementChartTitleCenteredOverlay
            cht.ChartTitle.Caption = "CY Revenue"
            cht.ChartTitle.Characters.Font.ColorIndex = 0
            cht.ChartTitle.Characters.Font.Size = 14
            v = cht.PlotArea.Left
            cht.PlotArea.Left = -4
            cht.PlotArea.Top = -2
            cht.PlotArea.Width = 396.25
            cht.PlotArea.Height = 163.25
            cht.ChartTitle.Left = cht.PlotArea.Width * 0.5 - (cht.ChartTitle.Width / 2)
            cht.ChartTitle.Top = cht.PlotArea.Height * 0.7561
            cht.PlotArea.Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
            cht.ApplyDataLabels xlDataLabelsShowLabelAndPercent ', , , , True
'            cht.FullSeriesCollection(1).DataLabels.Position = xlLabelPositionLeft
            cht.FullSeriesCollection(1).DataLabels.Position = xlLabelPositionOutsideEnd
            cht.FullSeriesCollection(1).DataLabels.Font.Bold = True
'            For Each pt In cht.FullSeriesCollection(1).Points
'                pt.DataLabel.Select
'                'pt.DataLabel.Orientation =
'            Next
            
            
            
        Case "linecount"
            If bRefresh = True Then Exit Sub
            Set RangeIncLabels = Range(Range(Cells(13 + ad, 6), Cells(15 + ad, 6)).Address & "," & Range(Cells(13 + ad, 8), Cells(15 + ad, 8)).Address)
            Set cht = mCAM_Runtime.cDocHolder.AddChart(3, 13 + ad, "G:\Central_Buying\General\4_Administration\Z - Buying Systems Analyst\LIVE EXCEL CHART TEMPLATES\cht_CAM_LCMini.crtx", lDoc_ID, RangeIncLabels, , , , , , 8.21798918968 * ColW, 8.332478379 * RowH)
            cht.ChartTitle.Caption = "Line Count"
            cht.ChartTitle.Characters.Font.ColorIndex = 0
            cht.FullSeriesCollection(1).DataLabels.Position = xlLabelPositionCenter
'            With cht.FullSeriesCollection(1)
'                .DataLabels.Position = xlLabelPositionCenter
'                For Each v In .Points
''                    v.DataLabel.Select
'                    Debug.Print v.DataLabel.Top, v.DataLabel.Left
'                    Debug.Print v.Name
'                    If (v.DataLabel.Top / cht.ChartArea.Height) < 0.5 Then
'                        If (v.DataLabel.Left / cht.ChartArea.Width) < 0.5 Then
'                            Set shp = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, v.DataLabel.Left - 75, v.DataLabel.Top - 40, 75, 40)
'                        Else
'                            Set shp = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, v.DataLabel.Left + v.DataLabel.Width, v.DataLabel.Top - 40, 75, 40)
'                        End If
'                    Else
'                        If (v.DataLabel.Left / cht.ChartArea.Width) < 0.5 Then
'                            Set shp = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, v.DataLabel.Left - 75, v.DataLabel.Top, 75, 40)
'                        Else
'                            Set shp = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, v.DataLabel.Left + v.DataLabel.Width, v.DataLabel.Top, 75, 40)
'                        End If
'                    End If
'                    shp.TextFrame2.TextRange.Characters.Font.Size = 8
'                    If v.Name = "S1P1" Then
'                        shp.TextFrame2.TextRange.Characters.text = "Core Range Revenue 55 YoY growth 2.1%"
'                        shp.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
'                    ElseIf v.Name = "S1P2" Then
'                        shp.TextFrame2.TextRange.Characters.text = "Specials Range Revenue 55 YoY growth 2.1%"
'                    Else
'                        shp.TextFrame2.TextRange.Characters.text = "Seasonal Range Revenue 55 YoY growth 2.1%"
'                    End If
'                'Debug.Print v.DataLabel.Name
'                Next
'            End With
        Case "cygrossprofit"
            If bRefresh = True Then Exit Sub
            Set RangeIncLabels = Range(Cells(12 + ad, 6), Cells(15 + ad, 7))
            Set cht = mCAM_Runtime.cDocHolder.AddChart(24, 12 + ad, "G:\Central_Buying\General\4_Administration\Z - Buying Systems Analyst\LIVE EXCEL CHART TEMPLATES\cht_CAM_CYGrossProfit.crtx", lDoc_ID, RangeIncLabels, , , , , , 10.023252975109 * ColW, 12.76785713 * RowH)
            v = cht.PlotArea.Left
            cht.PlotArea.Left = -3
            cht.PlotArea.Top = -3
            cht.PlotArea.Width = 322.249921259842
            cht.PlotArea.Height = 164
            cht.SetElement msoElementChartTitleCenteredOverlay
            cht.ChartTitle.Caption = "CY Gross Profit"
            cht.ChartTitle.Characters.Font.ColorIndex = 0
            cht.ChartTitle.Characters.Font.Size = 14
            cht.ChartTitle.Top = cht.PlotArea.Height * 0.7561
            cht.PlotArea.Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Case "pygrossprofit"
            If bRefresh = True Then Exit Sub
            Set RangeIncLabels = Range(Cells(16 + ad, 6), Cells(19 + ad, 7))
            Set cht = mCAM_Runtime.cDocHolder.AddChart(25, 13 + ad, "G:\Central_Buying\General\4_Administration\Z - Buying Systems Analyst\LIVE EXCEL CHART TEMPLATES\cht_CAM_PYGrossProfit.crtx", lDoc_ID, RangeIncLabels, , , , , , 8.21798918968 * ColW, 8.332478379 * RowH)
            cht.ChartTitle.Caption = "PY Gross Profit"
            cht.ChartTitle.Characters.Font.ColorIndex = 0
            cht.ChartTitle.Characters.Font.Size = 14
            cht.ChartTitle.Top = cht.PlotArea.Height * 0.5 - (cht.ChartTitle.Height / 2)
            cht.ChartTitle.Left = cht.PlotArea.Width * 0.5 - (cht.ChartTitle.Width / 2)
            cht.FullSeriesCollection(1).DataLabels.Position = xlLabelPositionInsideEnd
            For Each pt In cht.FullSeriesCollection(1).Points
                If pt.DataLabel.Caption = "Add text" Then pt.DataLabel.AutoText = True
            Next
            
            
    End Select
    bRefresh = True
End Sub
''Private Sub BuildCatPerfOverViewCharts(ByVal ChartToBuild As String, ByRef bRefresh As Boolean)
''Dim RangeIncLabels As Range, RangeDataOnly As Range
''Dim cht As Chart
''Dim shp As Shape
''Dim v As Variant
''Dim ColW As Single
''Dim RowH As Single
''Dim pt As Point
''
''    ColW = Columns(1).Width
''    RowH = Rows(1).Height
''    Select Case LCase(ChartToBuild)
''        Case "toplineandforecast"
''            If bRefresh = True Then Exit Sub
''            Set RangeIncLabels = Range(Cells(16, 1), Cells(28, 5))
''            Set RangeDataOnly = Range(Cells(17, 2), Cells(28, 5))
''            Set cht = mCAM_Runtime.cDocHolder.AddChart(13, 16, "G:\Central_Buying\General\4_Administration\Z - Buying Systems Analyst\LIVE EXCEL CHART TEMPLATES\cht_CAM_CategoryLines.crtx", lDoc_ID, RangeIncLabels, RangeDataOnly, "MMM", True, "#", False, 11.1627907 * ColW, 27 * RowH)
''            cht.Legend.Top = cht.PlotArea.Height - (cht.PlotArea.Height * 0.01)
''        Case "cyrevenue"
''            If bRefresh = True Then Exit Sub
''            'Set RangeIncLabels = Range(Cells(16, 6), Cells(19, 7))
''            Set RangeIncLabels = Range(Range(Cells(20, 6), Cells(23, 6)).Address & "," & Range(Cells(20, 8), Cells(23, 8)).Address)
''            Set cht = mCAM_Runtime.cDocHolder.AddChart(1, 16, "G:\Central_Buying\General\4_Administration\Z - Buying Systems Analyst\LIVE EXCEL CHART TEMPLATES\cht_CAM_Revenue.crtx", lDoc_ID, RangeIncLabels, , , , , , 12.1317829 * ColW, 26.70833333 * RowH)
''            'cht.Select
''            cht.SetElement msoElementChartTitleCenteredOverlay
''            cht.ChartTitle.Caption = "CY Revenue"
''            cht.ChartTitle.Characters.Font.ColorIndex = 0
''            cht.ChartTitle.Characters.Font.Size = 14
''            v = cht.PlotArea.Left
''            cht.PlotArea.Left = -4
''            cht.PlotArea.Top = -2
''            cht.PlotArea.Width = 396.25
''            cht.PlotArea.Height = 163.25
''            cht.ChartTitle.Left = cht.PlotArea.Width * 0.5 - (cht.ChartTitle.Width / 2)
''            cht.ChartTitle.Top = cht.PlotArea.Height * 0.7561
''            cht.PlotArea.Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
''            cht.ApplyDataLabels xlDataLabelsShowLabelAndPercent ', , , , True
'''            cht.FullSeriesCollection(1).DataLabels.Position = xlLabelPositionLeft
''            cht.FullSeriesCollection(1).DataLabels.Position = xlLabelPositionOutsideEnd
''            cht.FullSeriesCollection(1).DataLabels.Font.Bold = True
'''            For Each pt In cht.FullSeriesCollection(1).Points
'''                pt.DataLabel.Select
'''                'pt.DataLabel.Orientation =
'''            Next
''
''
''
''        Case "linecount"
''            If bRefresh = True Then Exit Sub
''            Set RangeIncLabels = Range(Range(Cells(17, 6), Cells(19, 6)).Address & "," & Range(Cells(17, 8), Cells(19, 8)).Address)
''            Set cht = mCAM_Runtime.cDocHolder.AddChart(3, 18, "G:\Central_Buying\General\4_Administration\Z - Buying Systems Analyst\LIVE EXCEL CHART TEMPLATES\cht_CAM_LCMini.crtx", lDoc_ID, RangeIncLabels, , , , , , 8.21798918968 * ColW, 17.946876525879 * RowH)
''            cht.ChartTitle.Caption = "Line Count"
''            cht.ChartTitle.Characters.Font.ColorIndex = 0
''            cht.FullSeriesCollection(1).DataLabels.Position = xlLabelPositionCenter
'''            With cht.FullSeriesCollection(1)
'''                .DataLabels.Position = xlLabelPositionCenter
'''                For Each v In .Points
''''                    v.DataLabel.Select
'''                    Debug.Print v.DataLabel.Top, v.DataLabel.Left
'''                    Debug.Print v.Name
'''                    If (v.DataLabel.Top / cht.ChartArea.Height) < 0.5 Then
'''                        If (v.DataLabel.Left / cht.ChartArea.Width) < 0.5 Then
'''                            Set shp = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, v.DataLabel.Left - 75, v.DataLabel.Top - 40, 75, 40)
'''                        Else
'''                            Set shp = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, v.DataLabel.Left + v.DataLabel.Width, v.DataLabel.Top - 40, 75, 40)
'''                        End If
'''                    Else
'''                        If (v.DataLabel.Left / cht.ChartArea.Width) < 0.5 Then
'''                            Set shp = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, v.DataLabel.Left - 75, v.DataLabel.Top, 75, 40)
'''                        Else
'''                            Set shp = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, v.DataLabel.Left + v.DataLabel.Width, v.DataLabel.Top, 75, 40)
'''                        End If
'''                    End If
'''                    shp.TextFrame2.TextRange.Characters.Font.Size = 8
'''                    If v.Name = "S1P1" Then
'''                        shp.TextFrame2.TextRange.Characters.text = "Core RangeRevenue 55%YoY growth 2.1%"
'''                        shp.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
'''                    ElseIf v.Name = "S1P2" Then
'''                        shp.TextFrame2.TextRange.Characters.text = "Specials RangeRevenue 55%YoY growth 2.1%"
'''                    Else
'''                        shp.TextFrame2.TextRange.Characters.text = "Seasonal RangeRevenue 55%YoY growth 2.1%"
'''                    End If
'''                'Debug.Print v.DataLabel.Name
'''                Next
'''            End With
''        Case "cygrossprofit"
''            If bRefresh = True Then Exit Sub
''            Set RangeIncLabels = Range(Cells(16, 6), Cells(19, 7))
''            Set cht = mCAM_Runtime.cDocHolder.AddChart(24, 16, "G:\Central_Buying\General\4_Administration\Z - Buying Systems Analyst\LIVE EXCEL CHART TEMPLATES\cht_CAM_CYGrossProfit.crtx", lDoc_ID, RangeIncLabels, , , , , , 10.023252975109 * ColW, 27.5 * RowH)
''            v = cht.PlotArea.Left
''            cht.PlotArea.Left = -3
''            cht.PlotArea.Top = -3
''            cht.PlotArea.Width = 322.249921259842
''            cht.PlotArea.Height = 164
''            cht.SetElement msoElementChartTitleCenteredOverlay
''            cht.ChartTitle.Caption = "CY Gross Profit"
''            cht.ChartTitle.Characters.Font.ColorIndex = 0
''            cht.ChartTitle.Characters.Font.Size = 14
''            cht.ChartTitle.Top = cht.PlotArea.Height * 0.7561
''            cht.PlotArea.Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
''        Case "pygrossprofit"
''            If bRefresh = True Then Exit Sub
''            Set RangeIncLabels = Range(Cells(20, 6), Cells(23, 7))
''            Set cht = mCAM_Runtime.cDocHolder.AddChart(25, 18, "G:\Central_Buying\General\4_Administration\Z - Buying Systems Analyst\LIVE EXCEL CHART TEMPLATES\cht_CAM_PYGrossProfit.crtx", lDoc_ID, RangeIncLabels, , , , , , 8.21798918968 * ColW, 17.946876525879 * RowH)
''            cht.ChartTitle.Caption = "PY Gross Profit"
''            cht.ChartTitle.Characters.Font.ColorIndex = 0
''            cht.ChartTitle.Characters.Font.Size = 14
''            cht.ChartTitle.Top = cht.PlotArea.Height * 0.5 - (cht.ChartTitle.Height / 2)
''            cht.ChartTitle.Left = cht.PlotArea.Width * 0.5 - (cht.ChartTitle.Width / 2)
''            cht.FullSeriesCollection(1).DataLabels.Position = xlLabelPositionInsideEnd
''            For Each pt In cht.FullSeriesCollection(1).Points
''                If pt.DataLabel.Caption = "Add text" Then pt.DataLabel.AutoText = True
''            Next
''
''
''    End Select
''    bRefresh = True
''End Sub
Private Function f_TotalCatPerf(ByVal sName As String, ByVal lGrp_No As Long, Optional sSuff As String = "")
    ' Handle Total Category Performance Report fields - @3. Total Category Performance  Note if sSuff>"" then is the Total Business f_TotalCatPerf
    Dim CP As cCBA_ProdGroup, cLC As cCAM_LineCount_Data, dtTo As Date, dtFrom As Date, dtToM As Date, dtFromM As Date, lLC As Long, v As Variant
    Dim S1 As Single, S2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    On Error GoTo Err_Routine
    CBA_Error = ""
    If lGrp_No > 0 And sName <> "Year3" Then
        dtFrom = CDate("01/01/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        dtTo = CDate("31/12/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        sName = g_Left(sName, Len(CStr(lGrp_No)))
    ElseIf sName = "Year3" Then
         f_TotalCatPerf = "Total Business MAT " & Format(DateSerial(Year(dtDateTo), Month(dtDateTo), Day(dtDateTo)), "mm/yyyy")
         Exit Function
    Else
        dtFrom = dtDateFrom: dtTo = dtDateTo
'        dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
'        dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo)))
'        dtTo = dtTo - 1
    End If
'    If IsNumeric(Right(sName, 1)) Then sName = g_Left(sName, 1)
    If lGrp_No = 1 Then
        Set cLC = sdLineCountDic(ePrevPriorCalendarYr)
    ElseIf lGrp_No = 2 Then
        Set cLC = sdLineCountDic(ePriorCalendarYr)
    Else
        Set cLC = sdLineCountDic(eMAT)
    End If

    Set CP = cPGrp
    Select Case sName
        Case "TotalQtyUnits"
            If sSuff > "" Then                  ' If Total Business
                f_TotalCatPerf = CP.getTotalBusiness(dtFrom, dtTo, "POSQTY")
            Else
                f_TotalCatPerf = CP.getPOS(dtFrom, dtTo, eQTY)
            End If
        Case "TotalQtyUnits+/-%PY"
            If sSuff > "" Then                  ' If Total Business
                S1 = CP.getTotalBusiness(dtFrom, dtTo, "POSQTY")
                S2 = CP.getTotalBusiness(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), "POSQTY")
                If S2 = 0 Then f_TotalCatPerf = "-" Else f_TotalCatPerf = ((S1 - S2) / S2) * 100
            Else
                S1 = CP.getPOS(dtFrom, dtTo, eQTY)
                S2 = CP.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eQTY)
                If S2 = 0 Then f_TotalCatPerf = "-" Else f_TotalCatPerf = ((S1 - S2) / S2) * 100
            End If
        Case "TotalRetailValue"
            If sSuff > "" Then                  ' If Total Business
                f_TotalCatPerf = CP.getTotalBusiness(dtFrom, dtTo, "POSRetail")
            Else
                f_TotalCatPerf = CP.getPOS(dtFrom, dtTo, eRetail)
            End If
        Case "TotalRetailValue+/-%PY"
            If sSuff > "" Then                  ' If Total Business
                S1 = CP.getTotalBusiness(dtFrom, dtTo, "POSRetail")
                S2 = CP.getTotalBusiness(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), "POSRetail")
                If S2 = 0 Then f_TotalCatPerf = "-" Else f_TotalCatPerf = ((S1 - S2) / S2) * 100
            Else
                S1 = CP.getPOS(dtFrom, dtTo, eRetail)
                S2 = CP.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eRetail)
                If S2 = 0 Then f_TotalCatPerf = "-" Else f_TotalCatPerf = ((S1 - S2) / S2) * 100
            End If
        Case "Margin%"
            If sSuff > "" Then                  ' If Total Business
                f_TotalCatPerf = (CP.getTotalBusiness(dtFrom, dtTo, "RCVRetailNet") - CP.getTotalBusiness(dtFrom, dtTo, "RCVCost")) / CP.getTotalBusiness(dtFrom, dtTo, "RCVRetail")
            Else
                f_TotalCatPerf = CP.getRCVMargin(eRCVMarginPercent, dtFrom, dtTo)
            End If
        Case "Margin%+/-%PY"
            If sSuff > "" Then                  ' If Total Business
                S1 = (CP.getTotalBusiness(dtFrom, dtTo, "RCVRetailNet") - CP.getTotalBusiness(dtFrom, dtTo, "RCVCost")) / CP.getTotalBusiness(dtFrom, dtTo, "RCVRetail")
                S2 = (CP.getTotalBusiness(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), "RCVRetailNet") - CP.getTotalBusiness(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), "RCVCost")) / CP.getTotalBusiness(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), "RCVRetail")
                If S2 = 0 Then f_TotalCatPerf = "-" Else f_TotalCatPerf = ((S1 - S2) / S2) * 100
            Else
                S1 = CP.getRCVMargin(eRCVMarginPercent, dtFrom, dtTo)
                S2 = CP.getRCVMargin(eRCVMarginPercent, DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo))
                If S2 = 0 Then f_TotalCatPerf = "-" Else f_TotalCatPerf = ((S1 - S2) / S2) * 100
            End If
        Case "Share"
            If sSuff > "" Then                  ' If Total Business
                f_TotalCatPerf = "100"
            Else
                S1 = CP.getPOS(dtFrom, dtTo, eRetail)
                S2 = CP.getTotalBusiness(dtFrom, dtTo, "POSRetail")
                If S2 = 0 Then f_TotalCatPerf = "-" Else f_TotalCatPerf = Round((S1 / S2), 2)
            End If
        Case "ContriCont."
            If sSuff > "" Then                  ' If Total Business
                f_TotalCatPerf = (CP.getTotalBusiness(dtFrom, dtTo, "RCVRetailNet") - CP.getTotalBusiness(dtFrom, dtTo, "RCVCost")) / CP.getTotalBusiness(dtFrom, dtTo, "RCVRetail")
            Else
                S1 = CP.getRCVMargin(eRCVContributionDollar, dtFrom, dtTo)
                S2 = (CP.getTotalBusiness(dtFrom, dtTo, "RCVRetailNet") - CP.getTotalBusiness(dtFrom, dtTo, "RCVCost"))
                If S2 = 0 Then f_TotalCatPerf = "-" Else f_TotalCatPerf = Round((S1 / S2), 2)
            End If
        Case "LineCount"
            If sSuff > "" Then                  ' If Total Business
                f_TotalCatPerf = "N/A"
            Else
                For Each v In cLC.sdProductAllocation
                    If cLC.sdProductAllocation(v) = e_LineCountType.eCurTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eDeleted Or cLC.sdProductAllocation(v) = e_LineCountType.eRegional Then
                    Else
                        S1 = S1 + 1
                    End If
                Next
                f_TotalCatPerf = S1
            End If
        Case Else
            f_TotalCatPerf = "Not Found"
    End Select
    
Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.f_TotalCatPerf", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_MarketOV1(ByVal sName As String, ByVal lGrp_No As Long, Optional sSuff As String = "", Optional MSegType As e_MSegType = eHomescan)
    ' Handle  ' @4.1. Market Overview, Table 1: Market vs ALDI overall
    Dim CP As cCBA_ProdGroup, cLC As cCAM_LineCount_Data, dtTo As Date, dtFrom As Date, dtToM As Date, dtFromM As Date, lLC As Long, v As Variant
    Dim S1 As Single, S2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    Dim N As cCBA_NielsenData
    Dim v2 As Variant
    On Error GoTo Err_Routine
    CBA_Error = ""
    If lGrp_No > 0 Then sName = g_Left(sName, Len(CStr(lGrp_No)))
'    If lGrp_No > 0 Then
'        dtFrom = CDate("01/01/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
'        dtTo = CDate("31/12/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
'
'    ElseIf sSuff = "PY" Then
'        dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 2))
'        dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
'        dtTo = dtTo - 1
'    Else
        dtFrom = dtDateFrom
        dtTo = dtDateTo
'    End If
'    If IsNumeric(Right(sName, 1)) Then sName = g_Left(sName, 1)
    Set CP = cPGrp
    If lGrp_No > 0 Then
        Set N = GetMSeg(GetAllSelMSegNames(MSegType, True)(lGrp_No - 1), MSegType)
        Set N = sdNData(CStr(N.CGno))(CStr(N.SCGno))
    End If
    Select Case sName
        Case "MarketSegment"
            f_MarketOV1 = GetAllSelMSegNames(MSegType)(lGrp_No - 1)
        Case "TotMktVol"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                f_MarketOV1 = N.QTY - cPGrp.getPOS(dtFrom, dtTo, eQTY, , , N.CGno, N.SCGno)
            ElseIf sSuff = "PY" Then            ' If Prior Year
                For Each v In sdNData
                    For Each v2 In sdNData(v)
                        S1 = (sdNData(v)(v2).QTY * sdNData(v)(v2).YOYQTY)
                        S2 = cPGrp.getPOS(dtFrom, dtTo, eQTY, , , sdNData(v)(v2).CGno, sdNData(v)(v2).SCGno) - cPGrp.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eQTY, , , sdNData(v)(v2).CGno, sdNData(v)(v2).SCGno)
                        s3 = s3 + ((sdNData(v)(v2).QTY - cPGrp.getPOS(dtFrom, dtTo, eQTY, , , sdNData(v)(v2).CGno, sdNData(v)(v2).SCGno)) - (S1 - S2)) ' THIS IS LAST YEARS NON ALDI TOTAL MARKET QTY
                    Next
                Next
                f_MarketOV1 = s3
            Else                                ' If This Year
                For Each v In sdNData
                    For Each v2 In sdNData(v)
                        S1 = S1 + sdNData(v)(v2).QTY
                    Next
                Next
                f_MarketOV1 = S1
            End If
        Case "TotMktVolChg%"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                S1 = N.QTY * N.YOYQTY
                S2 = cPGrp.getPOS(dtFrom, dtTo, eQTY, , , N.CGno, N.SCGno) - cPGrp.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eQTY, , , N.CGno, N.SCGno)
                s3 = (N.QTY - cPGrp.getPOS(dtFrom, dtTo, eQTY, , , N.CGno, N.SCGno)) - (S1 - S2) ' THIS IS LAST YEARS NON ALDI TOTAL MARKET QTY
                If s3 = 0 Then f_MarketOV1 = 0 Else f_MarketOV1 = ((N.QTY - cPGrp.getPOS(dtFrom, dtTo, eQTY, , , N.CGno, N.SCGno)) - s3) / s3
            ElseIf sSuff = "PY" Then            ' If Prior Year
                f_MarketOV1 = "N/A"
            Else                                ' If This Year
                For Each v In sdNData
                    For Each v2 In sdNData(v)
                        S1 = S1 + (sdNData(v)(v2).YOYQTY * sdNData(v)(v2).QTY)
                        S2 = S2 + sdNData(v)(v2).QTY
                    Next
                Next
                f_MarketOV1 = S1 / S2
            End If
        Case "TotMktValue"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                f_MarketOV1 = N.Retail - cPGrp.getPOS(dtFrom, dtTo, eRetail, , , N.CGno, N.SCGno)
            ElseIf sSuff = "PY" Then            ' If Prior Year
                For Each v In sdNData
                    For Each v2 In sdNData(v)
                        S1 = (sdNData(v)(v2).Retail * sdNData(v)(v2).YOYRetail)
                        S2 = cPGrp.getPOS(dtFrom, dtTo, eRetail, , , sdNData(v)(v2).CGno, sdNData(v)(v2).SCGno) - cPGrp.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eRetail, , , sdNData(v)(v2).CGno, sdNData(v)(v2).SCGno)
                        s3 = s3 + ((sdNData(v)(v2).QTY - cPGrp.getPOS(dtFrom, dtTo, eRetail, , , sdNData(v)(v2).CGno, sdNData(v)(v2).SCGno)) - (S1 - S2)) ' THIS IS LAST YEARS NON ALDI TOTAL MARKET RETAIL
                    Next
                Next
                f_MarketOV1 = s3
            Else                                ' If This Year
                For Each v In sdNData
                    For Each v2 In sdNData(v)
                        S1 = S1 + sdNData(v)(v2).Retail
                    Next
                Next
                f_MarketOV1 = S1
            End If
        Case "TotMkt+/-% PY"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                S1 = N.Retail * N.YOYRetail
                S2 = cPGrp.getPOS(dtFrom, dtTo, eRetail, , , N.CGno, N.SCGno) - cPGrp.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eRetail, , , N.CGno, N.SCGno)
                s3 = (N.Retail - cPGrp.getPOS(dtFrom, dtTo, eRetail, , , N.CGno, N.SCGno)) - (S1 - S2) ' THIS IS LAST YEARS NON ALDI TOTAL MARKET RETAIL
                If s3 = 0 Then f_MarketOV1 = 0 Else f_MarketOV1 = ((N.Retail - cPGrp.getPOS(dtFrom, dtTo, eRetail, , , N.CGno, N.SCGno)) - s3) / s3
            ElseIf sSuff = "PY" Then            ' If Prior Year
                f_MarketOV1 = "N/A"
            Else                                ' If This Year
                For Each v In sdNData
                    For Each v2 In sdNData(v)
                        S1 = S1 + (sdNData(v)(v2).YOYRetail * sdNData(v)(v2).Retail)
                        S2 = S2 + sdNData(v)(v2).Retail
                    Next
                Next
                f_MarketOV1 = S1 / S2
            End If
        Case "AldiVol"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                f_MarketOV1 = cPGrp.getPOS(dtFrom, dtTo, eQTY, , , N.CGno, N.SCGno)
            ElseIf sSuff = "PY" Then            ' If Prior Year
                f_MarketOV1 = cPGrp.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eQTY)
            Else                                ' If This Year
                f_MarketOV1 = cPGrp.getPOS(dtFrom, dtTo, eQTY)
            End If
        Case "AldiVolChg%"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                S1 = cPGrp.getPOS(dtFrom, dtTo, eQTY, , , N.CGno, N.SCGno)
                S2 = cPGrp.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eQTY, , , N.CGno, N.SCGno)
                If S2 = 0 Then f_MarketOV1 = 0 Else f_MarketOV1 = (S1 - S2) / S2
            ElseIf sSuff = "PY" Then            ' If Prior Year
                S1 = cPGrp.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eQTY)
                S2 = cPGrp.getPOS(DateAdd("YYYY", -2, dtFrom), DateAdd("YYYY", -2, dtTo), eQTY)
                If S2 = 0 Then f_MarketOV1 = 0 Else f_MarketOV1 = (S1 - S2) / S2
            Else                                ' If This Year
                S1 = cPGrp.getPOS(dtFrom, dtTo, eQTY)
                S2 = cPGrp.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eQTY)
                If S2 = 0 Then f_MarketOV1 = 0 Else f_MarketOV1 = (S1 - S2) / S2
            End If
        Case "AldiValue"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                f_MarketOV1 = cPGrp.getPOS(dtFrom, dtTo, eRetail, , , N.CGno, N.SCGno)
            ElseIf sSuff = "PY" Then            ' If Prior Year
                f_MarketOV1 = cPGrp.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eRetail)
            Else                                ' If This Year
                f_MarketOV1 = cPGrp.getPOS(dtFrom, dtTo, eRetail)
            End If
        Case "Aldi+/-% PY"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                S1 = cPGrp.getPOS(dtFrom, dtTo, eRetail, , , N.CGno, N.SCGno)
                S2 = cPGrp.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eRetail, , , N.CGno, N.SCGno)
                If S2 = 0 Then f_MarketOV1 = 0 Else f_MarketOV1 = (S1 - S2) / S2
            ElseIf sSuff = "PY" Then            ' If Prior Year
                S1 = cPGrp.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eRetail)
                S2 = cPGrp.getPOS(DateAdd("YYYY", -2, dtFrom), DateAdd("YYYY", -2, dtTo), eRetail)
                If S2 = 0 Then f_MarketOV1 = 0 Else f_MarketOV1 = (S1 - S2) / S2
            Else                                ' If This Year
                S1 = cPGrp.getPOS(dtFrom, dtTo, eRetail)
                S2 = cPGrp.getPOS(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eRetail)
                If S2 = 0 Then f_MarketOV1 = 0 Else f_MarketOV1 = (S1 - S2) / S2
            End If
        Case "AldiMargin%"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                f_MarketOV1 = cPGrp.getRCVMargin(eRCVMarginPercent, dtFrom, dtTo, , , N.CGno, N.SCGno)
            ElseIf sSuff = "PY" Then            ' If Prior Year
                f_MarketOV1 = cPGrp.getRCVMargin(eRCVMarginPercent, DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo))
            Else                                ' If This Year
                f_MarketOV1 = cPGrp.getRCVMargin(eRCVMarginPercent, dtFrom, dtTo)
            End If
        
        Case Else
            f_MarketOV1 = "Not Found"
    End Select
    
Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.f_MarketOV1", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_MarketOV2(ByVal sName As String, ByVal lGrp_No As Long, Optional sSuff As String = "", Optional MSegType As e_MSegType = eHomescan)
    ' Handle  @4.2. Market Overview, Table 2: Market Breakdown
    Dim CP As cCBA_ProdGroup, cLC As cCAM_LineCount_Data, dtTo As Date, dtFrom As Date, dtToM As Date, dtFromM As Date, lLC As Long, v As Variant
    Dim S1 As Single, S2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    Dim N As cCBA_NielsenData
    Dim v2 As Variant
    On Error GoTo Err_Routine
    CBA_Error = ""
'    If lGrp_No > 0 Then
'        dtFrom = CDate("01/01/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
'        dtTo = CDate("31/12/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
'        sName = g_Left(sName, Len(CStr(lGrp_No)))
'    ElseIf sSuff = "PY" Then
'        dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 2))
'        dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
'        dtTo = dtTo - 1
'    Else
'        dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
'        dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo)))
'        dtTo = dtTo - 1
'    End If
'    If IsNumeric(Right(sName, 1)) Then sName = g_Left(sName, 1)
    If lGrp_No > 0 Then sName = g_Left(sName, Len(CStr(lGrp_No)))
    dtFrom = dtDateFrom
    dtTo = dtDateTo
    Set CP = cPGrp
    If lGrp_No > 0 Then
        Set N = GetMSeg(GetAllSelMSegNames(MSegType, True)(lGrp_No - 1), MSegType)
        Set N = sdNData(CStr(N.CGno))(CStr(N.SCGno))
    End If
    Select Case sName
        Case "MarketSegment"
            f_MarketOV2 = GetAllSelMSegNames(MSegType)(lGrp_No - 1)
        Case "TotalValue%"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                f_MarketOV2 = N.CategoryMarketShare
            ElseIf sSuff = "PY" Then            ' If Prior Year
                f_MarketOV2 = "TBC"
            Else                                ' If This Year
                f_MarketOV2 = "TBC"
            End If
        Case "ALDI%ofCategory"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                S1 = cPGrp.getPOS(dtFrom, dtTo, eRetail, , , N.CGno, N.SCGno)
                S2 = N.Retail
                If S2 = 0 Then f_MarketOV2 = 0 Else f_MarketOV2 = S1 / S2
            ElseIf sSuff = "PY" Then            ' If Prior Year
                f_MarketOV2 = "TBC"
            Else                                ' If This Year
                f_MarketOV2 = "TBC"
            End If
        Case "Value($000)"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                f_MarketOV2 = N.MarketPLShare * N.Retail
            ElseIf sSuff = "PY" Then            ' If Prior Year
                f_MarketOV2 = "TBC"
            Else                                ' If This Year
                f_MarketOV2 = "TBC"
            End If
        Case "MarketShare%"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                f_MarketOV2 = N.MarketPLShare
            ElseIf sSuff = "PY" Then            ' If Prior Year
                f_MarketOV2 = "TBC"
            Else                                ' If This Year
                f_MarketOV2 = "TBC"
            End If
        Case "TotalMarketVolume%"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                If N.QTY = 0 Then f_MarketOV2 = 0 Else f_MarketOV2 = cPGrp.getPOS(dtFrom, dtTo, eQTY, , , N.CGno, N.SCGno) / N.QTY
            ElseIf sSuff = "PY" Then            ' If Prior Year
                f_MarketOV2 = "TBC"
            Else                                ' If This Year
                f_MarketOV2 = "TBC"
            End If
        Case "TotalMarketValue%"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                If N.Retail = 0 Then f_MarketOV2 = 0 Else f_MarketOV2 = cPGrp.getPOS(dtFrom, dtTo, eRetail, , , N.CGno, N.SCGno) / N.Retail
            ElseIf sSuff = "PY" Then            ' If Prior Year
                f_MarketOV2 = "TBC"
            Else                                ' If This Year
                f_MarketOV2 = "TBC"
            End If
        Case "LineCount"
            If lGrp_No > 0 Then                 ' If Grouped Lines
                f_MarketOV2 = "TBC"
            ElseIf sSuff = "PY" Then            ' If Prior Year
                f_MarketOV2 = "TBC"
            Else                                ' If This Year
                f_MarketOV2 = "TBC"
            End If
        Case Else
            f_MarketOV2 = "Not Found"
    End Select
    
Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.f_MarketOV2", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_CoreRangePerf(ByVal sName As String, ByVal lGrp_No As Long)
    ' Handle Core Range Performance Report fields - @5. Core Range Performance
    Dim CP As cCBA_ProdGroup, cLC As cCAM_LineCount_Data, dtTo As Date, dtFrom As Date, dtToM As Date, dtFromM As Date, lLC As Long, v As Variant
    Dim S1 As Single, S2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    On Error GoTo Err_Routine
    CBA_Error = ""
    ''If IsNumeric(Right(sName, 1)) Then sName = g_Left(sName, 1)
    If lGrp_No > 0 Then
        dtFrom = CDate("01/01/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        dtTo = CDate("31/12/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        sName = g_Left(sName, Len(CStr(lGrp_No)))
    Else
        dtFrom = dtDateFrom
        dtTo = dtDateTo
        'dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
        'dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo)))
        'dtTo = dtTo - 1
    End If
    Set CP = cPGrp
    Select Case sName
        Case "TotalQtyUnits"
            f_CoreRangePerf = CStr(Format(CP.getPOS(dtFrom, dtTo, e_RetailorQTY.eRetail, 1), "#,0"))
        Case "TotalRetail"
            f_CoreRangePerf = CStr(Format(CP.getPOS(dtFrom, dtTo, e_RetailorQTY.eRetail, 1), "$#,0"))
        Case "TRV+/-%PY"
            S1 = CP.getPOS(dtFrom, dtTo, e_RetailorQTY.eRetail, 1)
            S2 = CP.getPOS(DateAdd("yyyy", -1, dtFrom), DateAdd("yyyy", -1, dtTo), e_RetailorQTY.eRetail, 1)
            If S1 > 0 Then
                f_CoreRangePerf = CStr(Round((S1 - S2) / S1, 4))
            Else
                f_CoreRangePerf = "0.0%"
            End If
        Case "Margin"
            f_CoreRangePerf = CStr(CP.getRCVMargin(e_RCVMarginType.eRCVMarginPercent, dtFrom, dtTo))
        Case "Mgn+/-%PY"
            f_CoreRangePerf = CP.getRCVMargin(e_RCVMarginType.eRCVMarginPercent, dtFrom, dtTo) - CP.getRCVMargin(e_RCVMarginType.eRCVMarginPercent, DateAdd("yyyy", -1, dtFrom), DateAdd("yyyy", -1, dtTo))
        Case "Share"
            s4 = 0: s5 = 0
            S2 = Year(dtFrom)
            s3 = Month(dtFrom)
            For S1 = 1 To 12
                s4 = CP.sdTotBus(CStr(S2))(CStr(s3))("POSRetail")
                If s3 < 12 Then s3 = s3 + 1 Else s3 = 1: S2 = S2 + 1
            Next
            s5 = CP.getPOS(dtFrom, dtTo, e_RetailorQTY.eRetail, 1)
            If s4 = 0 Then f_CoreRangePerf = 0 Else f_CoreRangePerf = Format((s5 / s4) * 10, "0.00")
        Case "U/S/WperProd"
            f_CoreRangePerf = CStr(Round(CP.getUSW(e_USWorUSWpp.eUSWp, dtFrom, dtTo, e_RetailorQTY.eRetail, 1), 0))
        Case "%ShareofTotalCat"
            S1 = CP.getPOS(dtFrom, dtTo, e_RetailorQTY.eRetail, 1)
            S2 = CP.getPOS(dtFrom, dtTo, e_RetailorQTY.eRetail)
            If S2 = 0 Then f_CoreRangePerf = 0 Else f_CoreRangePerf = Format(S1 / S2, "0.00%")
        Case "NetCont.$/S/WperProd"
            S1 = CP.getUSW(eUSWp, dtFrom, dtTo, eQTY, 1)
            S2 = CP.getRCV(dtFrom, dtTo, eRCVRetailNet, 1)
            s3 = CP.getRCV(dtFrom, dtTo, eRCVCost, 1)
            s4 = CP.getWastageData(eRetail, eBoth, dtFrom, dtTo, , 1)
            s5 = CP.getRCV(dtFrom, dtTo, eRCVRetail, 1)
            s6 = CP.getRCV(dtFrom, dtTo, eRCVQTY, 1)
            If s5 = 0 Or s6 = 0 Then
                f_CoreRangePerf = 0
            Else
                f_CoreRangePerf = Format(S1 * ((S2 - s3 + s4) / s5) * (s5 / s6), "$#,0.00")
            End If
        Case "POSU/S/WperProd"
            f_CoreRangePerf = CP.getUSW(eUSWp, dtFrom, dtTo, eQTY, 1)
        Case "GrossU/S/WperProd"
            S1 = CP.getUSW(eUSWp, dtFrom, dtTo, eQTY, 1)
            S2 = CP.getRCV(dtFrom, dtTo, eRCVRetailNet, 1)
            s3 = CP.getRCV(dtFrom, dtTo, eRCVCost, 1)
            s4 = CP.getWastageData(eRetail, eBoth, dtFrom, dtTo, , 1)
            s5 = CP.getRCV(dtFrom, dtTo, eRCVRetail, 1)
            s6 = CP.getRCV(dtFrom, dtTo, eRCVQTY, 1)
            If s5 = 0 Or s6 = 0 Then
                f_CoreRangePerf = 0
            Else
                f_CoreRangePerf = Format(S1 * ((S2 - s3) / s5) * (s5 / s6), "$#,0.00")
            End If
        Case "LineCount"
            If lGrp_No = 1 Then
                Set cLC = sdLineCountDic(ePrevPriorCalendarYr)
            ElseIf lGrp_No = 2 Then
                Set cLC = sdLineCountDic(ePriorCalendarYr)
            Else
                Set cLC = sdLineCountDic(eMAT)
            End If
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eSeasonal Then
                    lLC = lLC + 1
                End If
            Next
            f_CoreRangePerf = CStr(lLC)
''        Case ""
''            f_CoreRangePerf = ""
        Case Else
''            f_CoreRangePerf = "Not Found"
    End Select

Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.f_CoreRangePerf", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_CoreRangeTrialPerf(ByVal sName As String, ByVal lGrp_No As Long)
    ' Handle Core Range Trial Performance Report fields @6. Core Range Trial Performance
    Dim CP As cCBA_ProdGroup, cLC As cCAM_LineCount_Data, dtTo As Date, dtFrom As Date, dtToM As Date, dtFromM As Date, lLC As Long, v As Variant
    Dim S1 As Single, S2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    Dim P As cCBA_Prod
    Dim arr As Variant
    Dim DtF As Date, DtT As Date
    On Error GoTo Err_Routine
    CBA_Error = ""
    ''If IsNumeric(Right(sName, 1)) Then sName = g_Left(sName, 1)
    If lGrp_No > 0 Then
        dtFrom = CDate("01/01/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        dtTo = CDate("31/12/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
    sName = g_Left(sName, Len(CStr(lGrp_No)))
    Else
        dtFrom = dtDateFrom
        dtTo = dtDateTo
'        dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
'        dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo)))
'        dtTo = dtTo - 1
    End If
    Set CP = cPGrp
    If lGrp_No = 1 Then
        Set cLC = sdLineCountDic(ePrevPriorCalendarYr)
    ElseIf lGrp_No = 2 Then
        Set cLC = sdLineCountDic(ePriorCalendarYr)
    Else
        Set cLC = sdLineCountDic(eMAT)
    End If
    Select Case sName
        Case "TotNoOfCatCRTrials"
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eCurTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eBrandSucTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eSucTrial Then
                    lLC = lLC + 1
                End If
            Next
            f_CoreRangeTrialPerf = CStr(lLC)
        Case "TrialRetail%"
            'GOING TO NEED TO BUILD A FUNCTION THAT LOOKS A THE REGION USW AND IDENTIFIES WHEN A PRODUCT GOES NATIONAL VS THE TRIAL
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eCurTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eBrandSucTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eSucTrial Then
                    If IsNumeric(Trim(Mid(v, 1, InStr(1, v, "-") - 1))) Then
                        Set P = CP.getProdObject(Trim(Mid(v, 1, InStr(1, v, "-") - 1)))
                        arr = P.CalculateTrialPeriod
                        If arr(0) < dtFrom Or arr(1) = arr(0) Then DtF = dtFrom Else DtF = arr(0)
                        If arr(1) > dtTo Or arr(1) = 0 Or arr(1) = arr(0) Then DtT = dtTo Else DtT = arr(1)
                        S1 = S1 + P.getPOSdata(DtF, DtT, False)
                    End If
                End If
            Next
            S2 = CP.getPOS(dtFrom, dtTo, eRetail, 1)
            If S2 = 0 Then f_CoreRangeTrialPerf = "0.00%" Else f_CoreRangeTrialPerf = S1 / S2
        Case "Margin%"
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eCurTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eBrandSucTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eSucTrial Then
                    If IsNumeric(Trim(Mid(v, 1, InStr(1, v, "-") - 1))) Then
                        Set P = CP.getProdObject(Trim(Mid(v, 1, InStr(1, v, "-") - 1)))
                        arr = P.CalculateTrialPeriod
                        If arr(0) < dtFrom Or arr(1) = arr(0) Then DtF = dtFrom Else DtF = arr(0)
                        If arr(1) > dtTo Or arr(1) = 0 Or arr(1) = arr(0) Then DtT = dtTo Else DtT = arr(1)
                        S1 = S1 + P.getRCVdata(DtF, DtT, "RCVRetailNet")
                        S2 = S2 + P.getRCVdata(DtF, DtT, "RCVCost")
                        s3 = s3 + P.getRCVdata(DtF, DtT, "RCVRetail")
                    End If
                End If
            Next
            If s3 = 0 Then f_CoreRangeTrialPerf = 0 Else f_CoreRangeTrialPerf = (S1 - S2) / s3
        Case "U/S/WperProd"
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eCurTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eBrandSucTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eSucTrial Then
                    If IsNumeric(Trim(Mid(v, 1, InStr(1, v, "-") - 1))) Then
                        Set P = CP.getProdObject(Trim(Mid(v, 1, InStr(1, v, "-") - 1)))
                        arr = P.CalculateTrialPeriod
                        If arr(0) < dtFrom Or arr(1) = arr(0) Then DtF = dtFrom Else DtF = arr(0)
                        If arr(1) > dtTo Or arr(1) = 0 Or arr(1) = arr(0) Then DtT = dtTo Else DtT = arr(1)
                        S1 = S1 + P.getPOSdata(DtF, DtT, True, eProductLevel)
                        S2 = S2 + 1
                    End If
                End If
            Next
            If S2 = 0 Then f_CoreRangeTrialPerf = 0 Else f_CoreRangeTrialPerf = Format(S1 / S2, "0.0")
        Case "GrossCont$/S/WperProd"
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eCurTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eBrandSucTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eSucTrial Then
                    If IsNumeric(Trim(Mid(v, 1, InStr(1, v, "-") - 1))) Then
                        Set P = CP.getProdObject(Trim(Mid(v, 1, InStr(1, v, "-") - 1)))
                        arr = P.CalculateTrialPeriod
                        If arr(0) < dtFrom Or arr(1) = arr(0) Then DtF = dtFrom Else DtF = arr(0)
                        If arr(1) > dtTo Or arr(1) = 0 Or arr(1) = arr(0) Then DtT = dtTo Else DtT = arr(1)
                        S1 = S1 + P.getPOSdata(DtF, DtT, False, eProductLevel)
                        S2 = S2 + 1
                    End If
                End If
            Next
            If S2 = 0 Then f_CoreRangeTrialPerf = 0 Else f_CoreRangeTrialPerf = Format(S1 / S2, "0.0")
        Case "%OfSuccess"
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eCurTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eBrandSucTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eSucTrial Then
                    S1 = S1 + 1
                End If
                If cLC.sdProductAllocation(v) = e_LineCountType.eCurTrial Then
                    S2 = S2 + 1
                End If
            Next
            If S1 = 0 Then f_CoreRangeTrialPerf = 0 Else f_CoreRangeTrialPerf = Format(S2 / S1, "0")
        Case Else
''            f_CoreRangeTrialPerf = "Not Found"
    End Select

Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.f_CoreRangeTrialPerf", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_SeasPerf(ByVal sName As String, ByVal lGrp_No As Long)
    ' Handle Seasonal Performance fields                        ' @7. Seasonal Performance
    Dim CP As cCBA_ProdGroup, cLC As cCAM_LineCount_Data, dtTo As Date, dtFrom As Date, lLC As Long, v As Variant

    Dim S1 As Single, S2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    On Error GoTo Err_Routine
    CBA_Error = ""
''    If IsNumeric(Right(sName, 1)) Then sName = g_Left(sName, 1)
    If lGrp_No > 0 Then
''        If sName = "Year" Then
''            f_SeasPerf = CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No)
''            GoTo Exit_Routine
''        End If
        dtFrom = CDate("01/01/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        dtTo = CDate("31/12/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        sName = g_Left(sName, Len(CStr(lGrp_No)))
    Else
''        If sName = "Year" Then
''                f_SeasPerf = "MAT " & Format(DateSerial(Year(dtDateTo), Month(dtDateTo), Day(dtDateTo)), "mm/yyyy")
''            GoTo Exit_Routine
''        End If
        dtFrom = dtDateFrom
        dtTo = dtDateTo

'        dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
'        dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo)))
'        dtTo = dtTo - 1
    End If
    Set CP = cPGrp
    Select Case sName
        Case "TotalRetailValue"
            f_SeasPerf = CStr(Format(CP.getPOS(dtFrom, dtTo, e_RetailorQTY.eRetail, 4), "$#,0"))
        Case "TRV+/-%PY"
            S1 = CP.getPOS(dtFrom, dtTo, e_RetailorQTY.eRetail, 4)
            S2 = CP.getPOS(DateAdd("yyyy", -1, dtFrom), DateAdd("yyyy", -1, dtTo), e_RetailorQTY.eRetail, 4)
            If S1 Then
                f_SeasPerf = CStr(Round((S1 - S2) / S1, 4))
            Else
                f_SeasPerf = "0%"
            End If
        Case "Margin"
            f_SeasPerf = CStr(CP.getRCVMargin(e_RCVMarginType.eRCVMarginPercent, dtFrom, dtTo))
        Case "Mgn+/-%PY"
            f_SeasPerf = CStr(CP.getRCVMargin(e_RCVMarginType.eRCVMarginPercent, dtFrom, dtTo) - CP.getRCVMargin(e_RCVMarginType.eRCVMarginPercent, DateAdd("yyyy", -1, dtFrom), DateAdd("yyyy", -1, dtTo)))
        Case "Markdown/Losses%Sales"
            S1 = CStr(CP.getPOS(dtFrom, dtTo, e_RetailorQTY.eRetail, 4))
            S2 = CP.getWastageData(eRetail, eBoth, dtFrom, dtTo, , 4)
            If S1 = 0 Then f_SeasPerf = 0 Else f_SeasPerf = Abs(S2) / S1
        Case "No.Items"
            If lGrp_No = 1 Then
                Set cLC = sdLineCountDic(ePrevPriorCalendarYr)
            ElseIf lGrp_No = 2 Then
                Set cLC = sdLineCountDic(ePriorCalendarYr)
            Else
                Set cLC = sdLineCountDic(eMAT)
            End If
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eSeasonal Then lLC = lLC + 1
            Next
            f_SeasPerf = CStr(lLC)
        Case "U/S/WperProd"
            f_SeasPerf = CStr(Round(CP.getUSW(e_USWorUSWpp.eUSWp, dtFrom, dtTo, e_RetailorQTY.eQTY, 4), 0))
        Case "%ShareofTotalCat"
            S1 = CP.getPOS(dtFrom, dtTo, e_RetailorQTY.eRetail, 4)
            S2 = CP.getPOS(dtFrom, dtTo, e_RetailorQTY.eRetail)
            If S2 = 0 Then f_SeasPerf = 0 Else f_SeasPerf = Format(S1 / S2, "0.00%")
        Case "NetCont.$/S/WperProd"

            S1 = CP.getUSW(eUSWp, dtFrom, dtTo, eQTY, 4)
            S2 = CP.getRCV(dtFrom, dtTo, eRCVRetailNet, 4)
            s3 = CP.getRCV(dtFrom, dtTo, eRCVCost, 4)
            s4 = CP.getWastageData(eRetail, eBoth, dtFrom, dtTo, , 4)
            s5 = CP.getRCV(dtFrom, dtTo, eRCVRetail, 4)
            s6 = CP.getRCV(dtFrom, dtTo, eRCVQTY, 4)

            If s5 = 0 Or s6 = 0 Then
                f_SeasPerf = 0
            Else
                f_SeasPerf = Format(S1 * ((S2 - s3 + s4) / s5) * (s5 / s6), "$#,0.00")
            End If
        Case Else
            f_SeasPerf = "Not Found"
    End Select

Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.f_SeasPerf", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function f_SpecPerf(ByVal sName As String, ByVal lGrp_No As Long) ''CoreRangePerf
    ' Handle Specials Performance fields                    ' @8. Specials Performance
    Dim CP As cCBA_ProdGroup, cLC As cCAM_LineCount_Data, dtTo As Date, dtFrom As Date, dtToM As Date, dtFromM As Date, lLC As Long, v As Variant
    Dim S1 As Single, S2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    On Error GoTo Err_Routine
    CBA_Error = ""
''    If IsNumeric(Right(sName, 1)) Then sName = g_Left(sName, 1)
    If lGrp_No > 0 Then
''        If sName = "Year" Then
''            f_SpecPerf = CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No)
''            GoTo Exit_Routine
''        End If
        dtFrom = CDate("01/01/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        dtTo = CDate("31/12/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        sName = g_Left(sName, Len(CStr(lGrp_No)))
    Else
''        If sName = "Year" Then
''                f_SpecPerf = "MAT " & Format(DateSerial(Year(dtDateTo), Month(dtDateTo), Day(dtDateTo)), "mm/yyyy")
''            GoTo Exit_Routine
''        End If
        dtFrom = dtDateFrom
        dtTo = dtDateTo

'        dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
'        dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo)))
'        dtTo = dtTo - 1
    End If
    Set CP = cPGrp
    Select Case sName
        Case "TotalRetailValue"
            f_SpecPerf = CStr(Format(CP.getPOS(dtFrom, dtTo, e_RetailorQTY.eRetail, 5), "$#,0"))
        Case "TRV+/-%PY"
            S1 = CP.getPOS(dtFrom, dtTo, e_RetailorQTY.eRetail, 5)
            S2 = CP.getPOS(DateAdd("yyyy", -1, dtFrom), DateAdd("yyyy", -1, dtTo), e_RetailorQTY.eRetail, 5)
            If S1 > 0 Then
                If ((S1 - S2) / S1) < -1 Then
                    f_SpecPerf = "<-100%"
                ElseIf ((S1 - S2) / S1) > 1 Then
                    f_SpecPerf = ">100%"
                Else
                    f_SpecPerf = CStr(Format(Round((S1 - S2) / S1, 4), "#,0.00%"))
                End If
            Else
                f_SpecPerf = "0.00%"
            End If
        Case "Margin"
            S1 = CP.getRCV(dtFrom, dtTo, eRCVRetailNet, 5)
            S2 = CP.getRCV(dtFrom, dtTo, eRCVCost, 5)
            s3 = CP.getRCV(dtFrom, dtTo, eRCVRetail, 5)
            If s3 = 0 Then f_SpecPerf = 0 Else f_SpecPerf = CStr(Format((S1 - S2) / s3, "#,0.00%"))
        Case "Mgn+/-%PY"
            S1 = CP.getRCV(dtFrom, dtTo, eRCVRetailNet, 5)
            S2 = CP.getRCV(dtFrom, dtTo, eRCVCost, 5)
            s3 = CP.getRCV(dtFrom, dtTo, eRCVRetail, 5)
            s4 = CP.getRCV(DateAdd("yyyy", -1, dtFrom), DateAdd("yyyy", -1, dtTo), eRCVRetailNet, 5)
            s5 = CP.getRCV(DateAdd("yyyy", -1, dtFrom), DateAdd("yyyy", -1, dtTo), eRCVCost, 5)
            s6 = CP.getRCV(DateAdd("yyyy", -1, dtFrom), DateAdd("yyyy", -1, dtTo), eRCVRetail, 5)
            If s3 = 0 Then S1 = 0 Else S1 = (S1 - S2) / s3
            If s6 = 0 Then s4 = 0 Else s4 = (s4 - s5) / s6
            f_SpecPerf = CStr(Format((S1 - s4), "#,0.00%"))
        Case "Markdown/Losses%Sales"
            S1 = CP.getWastageData(eRetail, eBoth, dtFrom, dtTo, , 5)
            S2 = CP.getPOS(dtFrom, dtTo, eRetail, 5)
            If S2 = 0 Then
                f_SpecPerf = "0.00%"
            Else
                f_SpecPerf = Format((Abs(S1) / S2), "#,0.00%")
            End If
        Case "No.Items"
            If lGrp_No = 1 Then
                Set cLC = sdLineCountDic(ePrevPriorCalendarYr)
            ElseIf lGrp_No = 2 Then
                Set cLC = sdLineCountDic(ePriorCalendarYr)
            Else
                Set cLC = sdLineCountDic(eMAT)
            End If
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eSeasonal Then
                    lLC = lLC + 1
                End If
            Next
            f_SpecPerf = CStr(lLC)
        Case "U/S/WperProd"
            f_SpecPerf = CStr(Format(CP.getUSW(eUSWp, dtFrom, dtTo, e_RetailorQTY.eQTY, 5), "#,0.0"))
        Case "%ShareofTotalCat"
            S1 = CP.getPOS(dtFrom, dtTo, eRetail, 5)
            S2 = CP.getPOS(dtFrom, dtTo, eRetail)
            If S2 = 0 Then f_SpecPerf = "0.0%" Else f_SpecPerf = Format(S1 / S2, "#,0.0%")
        Case "NetCont.$/S/WperProd"
            S1 = CP.getUSW(eUSWp, dtFrom, dtTo, eQTY, 5)
            S2 = CP.getRCV(dtFrom, dtTo, eRCVRetailNet, 5)
            s3 = CP.getRCV(dtFrom, dtTo, eRCVCost, 5)
            s4 = CP.getWastageData(eRetail, eBoth, dtFrom, dtTo, , 5)
            s5 = CP.getRCV(dtFrom, dtTo, eRCVRetail, 5)
            s6 = CP.getRCV(dtFrom, dtTo, eRCVQTY, 5)
            If s5 = 0 Or s6 = 0 Then
                f_SpecPerf = 0
            Else
                f_SpecPerf = Format(S1 * ((S2 - s3 + s4) / s5) * (s5 / s6), "$#,0.00")
            End If                   'CStr(cP.getUSW(eUSW, dtFrom, dtTo, eRetail))
        Case "Sellthrough%14days"
            f_SpecPerf = CP.getSellThrough(dtFrom, dtTo, 5)
''        Case ""
''            f_SpecPerf = ""
        Case Else
            f_SpecPerf = "Not Found"
    End Select

Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.f_SpecPerf", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_LCOverView(ByVal sName As String, ByVal lGrp_No As Long)
    ' Handle Line Count fields
    Dim cLC As cCAM_LineCount_Data, lLC As Long, lLC1 As Long, v As Variant '', dtTo As Date, dtFrom As Date, dtToM As Date, dtFromM As Date,
    Dim S1 As Single, S2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    On Error GoTo Err_Routine
    CBA_Error = ""
''   If IsNumeric(Right(sName, 1)) Then sName = g_Left(sName, 1)
    If lGrp_No > 0 Then sName = g_Left(sName, Len(CStr(lGrp_No)))
    If lGrp_No = 1 Then
        Set cLC = sdLineCountDic(ePrevPriorCalendarYr)
    ElseIf lGrp_No = 2 Then
        Set cLC = sdLineCountDic(ePriorCalendarYr)
    Else
        Set cLC = sdLineCountDic(eMAT)
    End If
    Select Case sName
        Case "CORERANGETotal"
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eCore Or cLC.sdProductAllocation(v) = e_LineCountType.eBrandSucTrial Or _
                            cLC.sdProductAllocation(v) = e_LineCountType.eBranded Or cLC.sdProductAllocation(v) = e_LineCountType.eSucTrial Then
                    lLC = lLC + 1
                End If
            Next
            f_LCOverView = CStr(lLC)
        Case "Branded"
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eBrandSucTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eBranded Or _
                            cLC.sdProductAllocation(v) = e_LineCountType.eSucTrial Then
                    lLC = lLC + 1
                End If
            Next
            f_LCOverView = CStr(lLC)
        Case "PrivateLabel"
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eBrandSucTrial Or cLC.sdProductAllocation(v) = e_LineCountType.eBranded Or _
                            cLC.sdProductAllocation(v) = e_LineCountType.eSucTrial Then
                    lLC1 = lLC1 ''+ 1
                ElseIf cLC.sdProductAllocation(v) = e_LineCountType.eCore Then
                    lLC = lLC + 1
                End If
            Next
            f_LCOverView = CStr(lLC - lLC1)
        Case "Healthy"
'            For Each v In cLC.sdProductAllocation
'                If cLC.sdProductAllocation(v) = e_LineCountType.eh Then
'                    lLC = lLC + cLC.sdProductAllocation.Count
'                End If
'            Next
'            f_LCOverView = CStr(lLC)
'            f_LCOverView = "TBC"  'CStr(lLC)
        Case "TRIALS"
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eCurTrial Then
                    lLC = lLC + 1
                End If
            Next
            f_LCOverView = CStr(lLC)
        Case "REGIONAL"
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eRegional Then
                    lLC = lLC + 1
                End If
            Next
            f_LCOverView = CStr(lLC)
        Case "SEASONAL"
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eSeasonal Then
                    lLC = lLC + 1
                End If
            Next
            f_LCOverView = CStr(lLC)
        Case "SPECIALS"
            For Each v In cLC.sdProductAllocation
                If cLC.sdProductAllocation(v) = e_LineCountType.eSpecial Then
                    lLC = lLC + 1
                End If
            Next
            f_LCOverView = CStr(lLC)
        
''            f_LCOverView = "TBC"
''        Case ""
''            f_LCOverView = ""
        Case Else
            f_LCOverView = "Not Found"
    End Select
    
Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.f_LCOverView", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_Wastage(ByVal sName As String, ByVal lGrp_No As Long)
    ' Handle Core Range Performance Report fields   @9. Wastage
    Dim CP As cCBA_ProdGroup, cLC As cCAM_LineCount_Data, dtTo As Date, dtFrom As Date, dtToM As Date, dtFromM As Date, lLC As Long, v As Variant
    Dim S1 As Single, S2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    Dim PClass As Long
    On Error GoTo Err_Routine
    CBA_Error = ""
    PClass = 1

''    If IsNumeric(Right(sName, 1)) Then sName = g_Left(sName, 1)
    If lGrp_No > 0 Then
        dtFrom = CDate("01/01/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        dtTo = CDate("31/12/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        sName = g_Left(sName, Len(CStr(lGrp_No)))
    Else
        dtFrom = dtDateFrom
        dtTo = dtDateTo
'        dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
'        dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo)))
'        dtTo = dtTo - 1
    End If
    Set CP = cPGrp
    Select Case sName
        Case "CatWastageIn%ofCat"
            S1 = CP.getRCV(dtFrom, dtTo, eRCVRetail, PClass)
            S2 = CP.getWastageData(eRetail, eBoth, dtFrom, dtTo, , PClass)
            'f_Wastage = CStr(CP.getw(dtFrom, dtTo, e_RetailorQTY.eRetail))
            If S1 = 0 Then f_Wastage = "0.00%" Else f_Wastage = Format(Abs(S2) / S1, "0.00%")
        Case "CatWastageIn%+/-%PY"
            S1 = CP.getRCV(dtFrom, dtTo, eRCVRetail, PClass)
            S2 = CP.getWastageData(eRetail, eBoth, dtFrom, dtTo, , PClass)
            s3 = CP.getRCV(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eRCVRetail, PClass)
            s4 = CP.getWastageData(eRetail, eBoth, DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), , PClass)
            'f_Wastage = CStr(CP.getw(dtFrom, dtTo, e_RetailorQTY.eRetail))
            If S1 = 0 Then S1 = 0 Else S1 = Abs(S2) / S1
            If s3 = 0 Then s3 = 0 Else s3 = Abs(s4) / s3
            s5 = S1 - s3
            f_Wastage = Format(s5, "0.00%")
        Case "Wastage$perStore"
              f_Wastage = Format(Abs(CP.getWastageData(eRetail, eStores, dtFrom, dtTo, , PClass)), "$0#,0.00")
        Case "Wastage$perStore+/-%PY"
            S1 = CP.getWastageData(eRetail, eStores, DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), , PClass)
            S2 = CP.getWastageData(eRetail, eStores, dtFrom, dtTo, , PClass)
            If S1 = 0 Then f_Wastage = "0.00%" Else f_Wastage = Format((S2 - S1) / S1, "0.00%")
        Case "NetMargin"
            S1 = CP.getRCV(dtFrom, dtTo, eRCVRetailNet, PClass)
            S2 = CP.getRCV(dtFrom, dtTo, eRCVCost, PClass)
            s3 = CP.getRCV(dtFrom, dtTo, eRCVRetail, PClass)
            s4 = CP.getWastageData(eRetail, eBoth, dtFrom, dtTo, , PClass)
            If s3 = 0 Then f_Wastage = "0.00%" Else f_Wastage = Format((S1 - S2 + s4) / s3, "#,0.00%")
        Case "NetMargin+/-%PY"

            S1 = CP.getRCV(dtFrom, dtTo, eRCVRetailNet, PClass) - CP.getRCV(dtFrom, dtTo, eRCVCost, PClass) + CP.getWastageData(eRetail, eBoth, dtFrom, dtTo, , PClass)
            S2 = CP.getRCV(dtFrom, dtTo, eRCVRetail, PClass)
            s3 = CP.getRCV(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eRCVRetailNet, PClass) - CP.getRCV(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eRCVCost, PClass) + CP.getWastageData(eRetail, eBoth, DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), , PClass)
            s4 = CP.getRCV(DateAdd("YYYY", -1, dtFrom), DateAdd("YYYY", -1, dtTo), eRCVRetail, PClass)
            If S2 = 0 Then S1 = 0 Else S1 = S1 / S2
            If s4 = 0 Then s3 = 0 Else s3 = s3 / s4
            f_Wastage = Format((S1 - s3), "#,0.00%")
        Case "NetCont.$/S/WperProd"
            f_Wastage = "TBC"
        Case "NetCont.$/S/WperProd+/-%PY"
            f_Wastage = "TBC"
''        Case ""
''            f_Wastage = ""
        Case Else
''            f_Wastage = "Not Found"
    End Select

Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.f_Wastage", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function


Private Function f_Forecast(ByVal sName As String, ByVal lGrp_No As Long, Optional sSuff As String = "")
    ' Handle @17. Forecasting
''    Dim CP As cCBA_ProdGroup, cLC As cCAM_LineCount_Data, dtTo As Date, dtFrom As Date, dtToM As Date, dtFromM As Date, lLC As Long, v As Variant
''    Dim s1 As Single, s2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    On Error GoTo Err_Routine
    CBA_Error = ""
''    GoTo Exit_Routine
    If lGrp_No > cFCast.GetForecastValue(lGrp_No, "Count") + 1 Then GoTo Exit_Routine
    If lGrp_No > 0 Then
''        dtFrom = CDate("01/01/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
''        dtTo = CDate("31/12/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        sName = g_Left(sName, Len(CStr(lGrp_No)))
    Else
''        dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
''        dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo)))
''        dtTo = dtTo - 1
    End If
'    If IsNumeric(Right(sName, 1)) Then sName = g_Left(sName, 1)
''    Set CP = cPGrp
    Select Case sName
        Case "Category"
            If lGrp_No > 0 Then
                f_Forecast = cFCast.GetForecastValue(lGrp_No, "Desc")
            Else
''                f_Forecast = ""
            End If
        Case "SalesFC"
            If lGrp_No > 0 Then
                f_Forecast = Round(cFCast.GetForecastValue(lGrp_No, "SalesRF"), 0)              ' Return this one into Sales Forecast
            Else
                f_Forecast = Round(cFCast.GetForecastValue(0, "TSalesRF"), 0)                   ' Return this one into Sales Total Forecast
            End If
        Case "SalesYoY"
            If lGrp_No > 0 Then
                f_Forecast = Round((cFCast.GetForecastValue(lGrp_No, "SalesYoY") / 100), 3)       ' Return this one into Sales YoY %
            Else
                f_Forecast = Round((cFCast.GetForecastValue(0, "TSalesYoY") / 100), 3)            ' Return this one into Sales Total YoY %
            End If
        Case "MarginFC"
            If lGrp_No > 0 Then
                f_Forecast = CStr(Round(cFCast.GetForecastValue(lGrp_No, "MarginRF"), 0))       ' Return this one into Margin Forecast
            Else
                f_Forecast = Round(cFCast.GetForecastValue(0, "TMarginRF"), 0)                  ' Return this one into Margin Total Forecast
            End If
        Case "MarginYoY"
            If lGrp_No > 0 Then
                f_Forecast = Round((cFCast.GetForecastValue(lGrp_No, "MarginYoY") / 100), 3)      ' Return this one into Margin YoY %
            Else
                f_Forecast = Round((cFCast.GetForecastValue(0, "TMarginYoY") / 100), 3)           ' Return this one into Margin Total YoY %
            End If
        Case Else
            f_Forecast = "Not Found"
    End Select
    
Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.f_Forecast", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_CoreRangeTrialProductList(ByVal sName As String, ByVal lGrp_No As Long, Optional sSuff As String = "")
    ' Handle  ' @20. Core Range Trial Product List
    Dim CP As cCBA_ProdGroup, cLC As cCAM_LineCount_Data, dtTo As Date, dtFrom As Date, dtToM As Date, dtFromM As Date, lLC As Long, v As Variant
    Dim S1 As Single, S2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    On Error GoTo Err_Routine
    CBA_Error = ""
    If lGrp_No > 0 Then
        dtFrom = CDate("01/01/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        dtTo = CDate("31/12/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        sName = g_Left(sName, Len(CStr(lGrp_No)))
    ElseIf sSuff = "PY" Then
        dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 2))
        dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
        dtTo = dtTo - 1
    Else
        dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
        dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo)))
        dtTo = dtTo - 1
    End If
    'If IsNumeric(Right(sName, 1)) Then sName = g_Left(sName, 1)
    Set CP = cPGrp
    Select Case sName
        Case "ProdDesc"
            f_CoreRangeTrialProductList = "TBC"
        Case "Successful/Unsuccessful"
            f_CoreRangeTrialProductList = "TBC"
        Case "DateTrialCommencement"
            f_CoreRangeTrialProductList = "TBC"
        Case "DateofNatLaunchorDeletion"
            f_CoreRangeTrialProductList = "TBC"
        Case "Margin"
            f_CoreRangeTrialProductList = "TBC"
        Case "Share"
            f_CoreRangeTrialProductList = "TBC"
        Case "Markdown/Losses%Sales"
            f_CoreRangeTrialProductList = "TBC"
        Case "U/S/WperProd"
            f_CoreRangeTrialProductList = "TBC"
        Case "NetCont.$/S/WperProd"
            f_CoreRangeTrialProductList = "TBC"
        Case ""
            f_CoreRangeTrialProductList = "TBC"
        Case Else
            f_CoreRangeTrialProductList = "Not Found"
    End Select
    
Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.f_CoreRangeTrialProductList", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_CoreRangePerfByProduct(ByVal sName As String, ByVal lGrp_No As Long, Optional sSuff As String = "")
    ' Handle  ' @21. Core Range Performance by Product
    Dim CP As cCBA_ProdGroup, cLC As cCAM_LineCount_Data, dtTo As Date, dtFrom As Date, dtToM As Date, dtFromM As Date, lLC As Long, v As Variant
    Dim S1 As Single, S2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single
    On Error GoTo Err_Routine
    CBA_Error = ""
    If lGrp_No > 0 Then
        dtFrom = CDate("01/01/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        dtTo = CDate("31/12/" & CStr(Year(dtDateTo) - lMINUS_YY + lGrp_No))
        sName = g_Left(sName, Len(CStr(lGrp_No)))
    ElseIf sSuff = "PY" Then
        dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 2))
        dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
        dtTo = dtTo - 1
    Else
        dtFrom = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo) - 1))
        dtTo = CDate("01/" & CStr(Month(dtDateTo)) & "/" & CStr(Year(dtDateTo)))
        dtTo = dtTo - 1
    End If
'    If IsNumeric(Right(sName, 1)) Then sName = g_Left(sName, 1)
    Set CP = cPGrp
    Select Case sName
        Case "ProdCODE"
            f_CoreRangePerfByProduct = "TBC"
        Case "ProdCODE"
            f_CoreRangePerfByProduct = "TBC"
        Case "ProdDesc"
            f_CoreRangePerfByProduct = "TBC"
        Case "COUNTRY"
            f_CoreRangePerfByProduct = "TBC"
        Case "SUPPLIERS"
            f_CoreRangePerfByProduct = "TBC"
        Case "UNITRETAIL$"
            f_CoreRangePerfByProduct = "TBC"
        Case "3M+/-%"
            f_CoreRangePerfByProduct = "TBC"
        Case "12M+/-%"
            f_CoreRangePerfByProduct = "TBC"
        Case "MARGIN"
            f_CoreRangePerfByProduct = "TBC"
        Case "SHARE"
            f_CoreRangePerfByProduct = "TBC"
        Case "USW"
            f_CoreRangePerfByProduct = "TBC"
        
        Case Else
            f_CoreRangePerfByProduct = "Not Found"
    End Select
    
Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.f_CoreRangePerfByProduct", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function
'System Relevent Code
Private Sub Class_Initialize()


End Sub

Public Function SaveSegmentsToDB() As Boolean
    Dim CN As cCBA_Connect
    Dim N As cCBA_NielsenData
    Dim v As Variant
    Dim bfound As Boolean
    Dim RS As ADODB.Recordset
    On Error GoTo Err_Routine
    CBA_Error = ""

    Set CN = New cCBA_Connect
    If CN.SetConnection("CAMERA", True) = False Then MsgBox ("Error in SaveSegmentsToDB - cCBA_Connect.Setconnection"): Exit Function
    If Not sdMSegs Is Nothing Then
        For Each v In sdMSegs
            If IsEmpty(v) = False Then
                Set N = sdMSegs(v)
                If N.Category_ID = 0 Then N.Category_ID = mCAM_Runtime.cRibbonData.sdCategoryID(Me.sCategoryName)
                If N.IsManual = True Then
                    bfound = False
                    If CN.Query("select MS_CN_ID, MS_ManualMSegName, MS_IsActive, MS_CN_ID from L1_ManualSegments where MS_CN_ID = " & mCAM_Runtime.cRibbonData.sdCategoryID(Me.sCategoryName), "UpdQManualMSegs" & Me.sCategoryName) = True Then
                        Set RS = CN.op("UpdQManualMSegs" & Me.sCategoryName)
                        RS.MoveFirst
                        Do Until RS.EOF
                            If RS(1) = N.MSegDescription Then
                                bfound = True
                                Exit Do
                            End If
                            RS.MoveNext
                        Loop
                        If bfound = True And RS(2) = False Then RS(2) = True: RS.Update
                        If bfound = False Then
                            CN.Query "INSERT INTO L1_ManualSegments (MS_CN_ID, MS_ManualMSegName, MS_IsActive) Values(" & N.Category_ID & ",'" & N.MSegDescription & "',TRUE)", "IManualMSeg" & N.Category & N.MSegDescription
                        End If
                    Else
                        If bfound = False Then
                            CN.Query "INSERT INTO L1_ManualSegments (MS_CN_ID, MS_ManualMSegName, MS_IsActive) Values(" & N.Category_ID & ",'" & N.MSegDescription & "',TRUE)", "IManualMSeg" & N.Category & N.MSegDescription
                        End If
                    End If
                ElseIf N.isScanData = True Then
                    bfound = False
                    If CN.Query("select SA.SA_SH_ID, SA.SA_CN_ID, SDH.SH_Desc from L1_ScanDataAllocation SA left join N0_ScanDataHeaders SDH on SDH.SH_ID = SA.SA_SH_ID where SA.SA_CN_ID = " & mCAM_Runtime.cRibbonData.sdCategoryID(Me.sCategoryName), "QScanDataMSeg" & Me.sCategoryName) = True Then
                        Set RS = CN.op("QScanDataMSeg" & Me.sCategoryName)
                        RS.MoveFirst
                        Do Until RS.EOF
                            If RS(2) = N.MSegDescription Then
                                bfound = True
                                Exit Do
                            End If
                            RS.MoveNext
                        Loop
                    Else
                        If bfound = False Then
                            If CN.Query("Select SH_ID from N0_ScanDataHeaders where SH_Desc = '" & N.MSegDescription & "'", "SH_ID" & N.MSegDescription) = True Then
                                'If N.Category_ID = 0 Then N.Category_ID = mCAM_Runtime.cRibbonData.sdCategoryID(Me.sCategoryName)
                                CN.Query "INSERT INTO L1_ScanDataAllocation (SA_SH_ID,SA_CN_ID) Values (" & CN.op("SH_ID" & N.MSegDescription)(0) & ", " & N.Category_ID & ")", "IScanDataMSeg" & N.Category & N.MSegDescription
                            End If
                        End If
                    End If
                End If
            End If
        Next

    'now loop through the recordset to look for the msegs that need to be inactivated
        If CN.Query("select MS_CN_ID, MS_ManualMSegName, MS_IsActive, MS_CN_ID from L1_ManualSegments where MS_CN_ID = " & mCAM_Runtime.cRibbonData.sdCategoryID(Me.sCategoryName), "QManualMSegs" & Me.sCategoryName, True) = True Then
            Set RS = CN.op("QManualMSegs" & Me.sCategoryName)
            RS.MoveFirst
            bfound = False
            Do Until RS.EOF
                bfound = False
                For Each v In sdMSegs
                    Set N = sdMSegs(v)
                    If N.IsManual = True Then
                        If RS(1) = N.MSegDescription Then
                            bfound = True
                            Exit For
                        End If
                    End If
                Next
                If bfound = False Then CN.Query "Update L1_ManualSegments set MS_IsActive = False where MS_ManualMSegName = '" & RS(1) & "' and MS_CN_ID = " & RS(0), "UInactiveManualMSeg" & RS(0) & "-" & RS(1)
                RS.MoveNext
            Loop
        'NEED TO ADD SCAN DATA INACTIVATION PROCESS
        End If
        If CN.Query("select SA.SA_SH_ID, SA.SA_CN_ID, SDH.SH_Desc from L1_ScanDataAllocation SA left join N0_ScanDataHeaders SDH on SDH.SH_ID = SA.SA_SH_ID where SA.SA_CN_ID = " & mCAM_Runtime.cRibbonData.sdCategoryID(Me.sCategoryName), "QScanDataMSeg" & Me.sCategoryName, True) = True Then
            Set RS = CN.op("QScanDataMSeg" & Me.sCategoryName)
            RS.MoveFirst
            bfound = False
            Do Until RS.EOF
                bfound = False
                For Each v In sdMSegs
                    Set N = sdMSegs(v)
                    If N.isScanData = True Then
                        If RS(2) = N.MSegDescription Then
                            bfound = True
                            Exit For
                        End If
                    End If
                Next
                If bfound = False Then CN.Query "Delete From L1_ScanDataAllocation where SA_SH_ID = " & RS(0) & " and SA_CN_ID = " & RS(1), "DELInactiveManualMSeg" & RS(0) & "-" & RS(1)
                RS.MoveNext
            Loop
        End If
    End If
SaveSegmentsToDB = True
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.SaveSegmentsToDB", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next


End Function

    ' This should be pull functionality other than sCategoryName?
Public Function Construct(ByVal Category As String, ByVal DFrom As Date, ByVal Dto As Date, ByVal byACG As Boolean, Optional ByVal DocumentType As e_DocuType = 0, Optional ByVal BuildCode As String = "BCRPTLWMs") As Boolean
    ' Tells the object to construct itself and how to do that.
    ' It will need to return back to the runtime for data pertaining to the category specifically, and potentially some other data
        ' however I believe these 3 variables will be all that is required to complete constuction of the necessary data objects        @RWCam????
    Dim P As cCBA_Prod
    Dim col As Collection
    Dim v As Variant
    On Error GoTo Err_Routine
    CBA_Error = ""
    dtDateFrom = DFrom
    dtDateTo = Dto
    bACG = byACG
    sCategoryName = Category
    eDocumentType = DocumentType
    sBuildCode = BuildCode
    Set sdCategoryList = mCAM_Runtime.cRibbonData.GetCategoryList(sCategoryName, bACG)
    If sdCategoryList Is Nothing Then Construct = False: Exit Function
    Set col = New Collection
    For Each v In sdCategoryList.Keys
        Set P = sdCategoryList(v)
        If P.sACat = "" And bACG = False Then
            col.Add CStr(Format(P.lLegCG, "000") & Format(P.lLegSCG, "000"))
        ElseIf P.lACatNum > 0 And bACG = True Then
            col.Add CStr(Format(P.lACGNum, "000") & Format(P.lASCGNum, "000"))
        End If
    Next
    Set cPGrp = New cCBA_ProdGroup
    'Ribbondata does not need COMRADE, bForecast. Could do without Wastage etc if no prod detail function?? @TP @RWCam
    cPGrp.RunDataGeneration dtDateFrom, dtDateTo, bACG, col, sBuildCode
    Set sdMSegs = mCAM_Runtime.cRibbonData.PullCategoryMSegs(sCategoryName)
    eDefaultMSegtype = eHomescan
    CreateAllLineObjectsForAllPeriods
    Select Case DocumentType
        Case eForecast, e_DocuType.eToplineCategoryPerformance, e_DocuType.eCoreRangeCategoryReview
            Set cFCast = New cCBA_Forecasting
            Call cFCast.GenForecasting(Year(Dto), Year(Dto), "", "", Val(P.lLegCG), 0, 0)
    End Select
    Construct = True
    bIsActive = True
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.Construct", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function
Private Function IsMSegDicActive() As Boolean
    IsMSegDicActive = bIsActive
End Function
Public Function GetMSeg(ByVal MSegName As String, ByVal eMSegType As e_MSegType) As cCBA_NielsenData
Dim v As Variant
Dim N As cCBA_NielsenData
    For Each v In sdMSegs
        Set N = sdMSegs(v)
        If ((N.IsManual = True And eMSegType = eManual1) Or (N.IsHomescan = True And eMSegType = eHomescan) _
            Or (N.isScanData = True And eMSegType = eScanData)) And N.MSegDescription = MSegName Then
                Set GetMSeg = N
                Exit Function
        End If
    Next
    Set GetMSeg = Nothing
End Function
Public Function AmendMSeg(ByVal AmendNotDelete As Boolean, ByRef ND As cCBA_NielsenData) As Boolean
Dim v As Variant
Dim N As cCBA_NielsenData

'this could bug when you add a manual segment that has the same name as a scandata segment?? No Apparently not!
    If ND Is Nothing Or ND.MSegDescription = "" Then AmendMSeg = False: Exit Function
    For Each v In sdMSegs
        Set N = sdMSegs(v)
        If N.isScanData = ND.isScanData And N.IsHomescan = ND.IsHomescan And N.IsManual = ND.IsManual And _
            N.Category_ID = ND.Category_ID And N.MSegDescription = ND.MSegDescription Then
                If AmendNotDelete = True Then
                    Set sdMSegs(v) = ND: AmendMSeg = True: Exit Function
                Else
                    sdMSegs.Remove (v): AmendMSeg = True: Exit Function
                End If

        End If
    Next
    sdMSegs.Add CStr(ND.MSegDescription), ND: AmendMSeg = True
End Function
Public Function GetAllSelMSegNames(ByVal eMSegType As e_MSegType, Optional HomeScanNaturalName As Boolean = False) As Variant
Dim col As Collection
Dim N As cCBA_NielsenData
Dim arr As Variant
Dim v As Variant
Dim a As Long
Dim strHS As String
    Set col = New Collection
    For Each v In sdMSegs
        Set N = sdMSegs(v)
        If (N.IsManual = True And eMSegType = eManual1) Or (N.IsHomescan = True And eMSegType = eHomescan) _
            Or (N.isScanData = True And eMSegType = eScanData) Then
            If N.IsHomescan = False Or HomeScanNaturalName = True Then
                col.Add CStr(N.MSegDescription)
            Else
                If mCAM_Runtime.cRibbonData.byACGSelected = True And N.ACG = True Then
                    strHS = mCAM_Runtime.cRibbonData.GetCGListing(eACatNum, True, True, N.Category_ID, N.CGno, N.SCGno, True)(0)
                    strHS = strHS & "|" & mCAM_Runtime.cRibbonData.GetCGListing(eACGnum, True, True, N.Category_ID, N.CGno, N.SCGno, True)(0)
                    strHS = strHS & "|" & mCAM_Runtime.cRibbonData.GetCGListing(eASCGnum, True, True, N.Category_ID, N.CGno, N.SCGno, True)(0)
                    col.Add strHS
                ElseIf mCAM_Runtime.cRibbonData.byACGSelected = 99 And N.ACG = False Then
                    strHS = mCAM_Runtime.cRibbonData.GetCGListing(eLegCG, False, True, N.Category_ID, N.CGno, N.SCGno, True)(0)
                    strHS = strHS & "|" & mCAM_Runtime.cRibbonData.GetCGListing(eLegSCG, False, True, N.Category_ID, N.CGno, N.SCGno, True)(0)
                    col.Add strHS
                End If
            End If
        End If
    Next
    ReDim arr(0 To IIf(col.Count > 0, col.Count - 1, 0))
    a = -1
    For Each v In col
        a = a + 1: arr(a) = CStr(v)
    Next
    GetAllSelMSegNames = arr
End Function
Public Function RequestProdData(ByVal PCode As String) As cCBA_Prod
    Dim P As cCBA_Prod, v As Variant
    For Each v In cPGrp
        Set P = cPGrp(v)
        If P.lPcode = PCode Then
            Set RequestProdData = P
            Exit Function
        End If
    Next
    Set RequestProdData = Nothing
End Function
Private Function CompareMeToAnother() ' @TP Put Back In

End Function
Private Function GetMSegToProdData() ' @TP Put Back In

End Function
Private Function AmendMSegToProdData() ' @TP Put Back In

End Function
Public Function SetLineCountObjectReset(ByVal thisPeriod As e_LineCountPeriod) As Boolean
    SetLineCountObjectReset = sdLineCountDic(thisPeriod).LineCountReset
End Function
Public Function AmendLineCountData(ByVal thisPeriod As e_LineCountPeriod, ByVal PCode As String, ByVal thisAllocation As e_LineCountType) As Boolean
    ' Amends a Line Count Object
Dim P As cCBA_Prod
    On Error GoTo Err_Routine
    CBA_Error = ""
    AmendLineCountData = sdLineCountDic(thisPeriod).Allocate(PCode, thisAllocation)
    If lDoc_ID > 0 Then
        If Not cPGrp Is Nothing Then
            Set P = cPGrp.getProdObject(Mid(PCode, 1, InStr(1, PCode, "-") - 1), False)
            If P.lPClass = 4 And (thisAllocation = eCore Or thisAllocation = eBranded Or thisAllocation = eBrandSucTrial) Then P.lPClass = 1
            If P.lPClass = 1 And (thisAllocation = eSeasonal) Then P.lPClass = 4
            'Debug.Print P.lPClass
            'Debug.Print e_LineCountType.eDeleted
        End If
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.amendLineCountData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Function AddLineCountObject(ByRef LC As cCAM_LineCount_Data, ByVal thisPeriod As e_LineCountPeriod, Optional ByRef CN As ADODB.Connection) As Boolean
    ' Adds a Line Count Object
    On Error GoTo Err_Routine
    CBA_Error = ""
    If sdLineCountDic Is Nothing Then Set sdLineCountDic = New Scripting.Dictionary
    If sdLineCountDic.Exists(thisPeriod) Then Exit Function
    sdLineCountDic.Add thisPeriod, LC
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.addLineCountObject", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function PullLineCountAllocation(ByVal thisPeriod As e_LineCountPeriod, ByVal thisType As e_LineCountType) As Variant
    ' This needs to query the bLineCount Object for other line count allocations that have happened in the past against this
    '  category for the period specified, if nothing found then it will provide the calculated LineCountAllocation
    Dim LCo As Scripting.Dictionary, col As Collection, v As Variant, arr() As Variant, a As Long, bValid As Boolean
    On Error GoTo Err_Routine
    CBA_Error = ""
    If sdLineCountDic.Exists(thisPeriod) = False Then AddLineCountObject CreateLineCountObject(thisPeriod), thisPeriod
    Set col = New Collection
    Set LCo = sdLineCountDic(thisPeriod).sdProductAllocation
    For Each v In LCo
        bValid = False
        If LCo(v) = thisType Then
            bValid = True
        ElseIf thisType = eCore And (LCo(v) = eBranded Or LCo(v) = eSucTrial Or LCo(v) = eBrandSucTrial) Then
            bValid = True
        ElseIf thisType = eBranded And LCo(v) = eBrandSucTrial Then
            bValid = True
        ElseIf thisType = eSucTrial And LCo(v) = eBrandSucTrial Then
            bValid = True
        End If

        If bValid = True Then
''            If CLng(v) = 45379 Then
''                a = a
''            End If
            col.Add CStr(v)
        End If
    Next
    ReDim arr(0 To IIf(col.Count > 0, col.Count - 1, 0))
    a = -1
    For Each v In col
        a = a + 1
        arr(a) = v
    Next
    PullLineCountAllocation = arr
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.pullLineCountAllocation", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Public Function SaveLineCountAllocation(Optional ByVal SendingToMaster As Boolean = False) As Boolean
    ' This needs to query the bLineCount Object for other line count allocations that have happened in the past against this
    '  category for the period specified, if nothing found then it will provide the calculated LineCountAllocation
'    Dim LCo As Scripting.Dictionary, col As Collection, v As Variant, arr() As Variant, a As Long, bValid As Boolean
'    On Error GoTo Err_Routine
'    CBA_Error = ""
'    SaveLineCountAllocation = False
'    If sdLineCountDic.Exists(thisPeriod) = False Then GoTo Exit_Routine
'    SaveLineCountAllocation = True
    ' Need to fill this out

    Dim v As Variant
    Dim LC As cCAM_LineCount_Data

    For Each v In sdLineCountDic
        Set LC = sdLineCountDic(v)
        LC.SaveLineCountToDB SendingToMaster
    Next

Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.SaveLineCountAllocation", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
'''Public Function PullCalculatedLineCountAllocation(ByVal thisPeriod As e_LineCountPeriod, ByVal thisType As e_LineCountType) As Variant   ' @RWCam Doesn't work...
'''    Dim LCo As Scripting.Dictionary
'''    Dim col As Collection
'''    Dim v As Variant
'''    Dim arr() As Variant, a As Long
'''    On Error GoTo Err_Routine
'''    CBA_Error = ""
'''    If sdLineCountDic.Exists(thisPeriod) = False Then addLineCountObject CreateLineCountObject(thisPeriod), thisPeriod
'''    Set col = New Collection
'''    Set LCo = sdLineCountDic          ''(thisPeriod).sdCalculatedProductAllocation
'''    For Each v In LCo
'''        If LCo(v) = thisType Then col.Add CLng(v)
'''    Next
'''    ReDim arr(0 To col.Count - 1)
'''    a = -1
'''    For Each v In col
'''        a = a + 1
'''        arr(a) = v
'''    Next
'''    ''pullLineCountAllocation = arr
'''    PullCalculatedLineCountAllocation = arr
'''Exit_Routine:
'''
'''    On Error Resume Next
'''    Exit Function
'''
'''Err_Routine:
'''    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.pullCalculatedLineCountAllocation", 3)
'''    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
'''    Debug.Print CBA_Error
'''    Stop            ' ^RW Camera + next line
'''    Resume Next
'''    Call g_FileWrite(g_getDB("Gen", True), CBA_Error, , , True, True)
'''    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
'''    GoTo Exit_Routine
'''    Resume Next
'''End Function
Private Function CreateLineCountObject(ByVal thisPeriod As e_LineCountPeriod, Optional ByRef CN As ADODB.Connection) As cCAM_LineCount_Data
    ' Creates an individual cCAM_bLineCount Object
    Dim LC As cCAM_LineCount_Data
    Dim SingleRunOnly As Boolean
    On Error GoTo Err_Routine
    CBA_Error = ""
    ' If more than one are being created then a Connection can be passed in, if not one is generated.
    If CN Is Nothing Then
        SingleRunOnly = True
        Set CN = New ADODB.Connection
        With CN
            .ConnectionTimeout = 50
            .CommandTimeout = 50
            .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Cam")
        End With
    End If

        Set LC = New cCAM_LineCount_Data
        If cPGrp Is Nothing Then Exit Function
        LC.BuildObject sCategoryName, dtDateTo, cPGrp, thisPeriod, CN, lDoc_ID
        Set CreateLineCountObject = LC

    If SingleRunOnly = True Then
        If CN.State <> 0 Then CN.Close
        Set CN = Nothing
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.createLineCountObject", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Function CreateAllLineObjectsForAllPeriods() As Boolean
    ' Creates all the line count objects for all periods for a Category Object. Runs CreateLineCountObject for each Period in e_LineCountPeriod
    Dim CN As ADODB.Connection
    On Error GoTo Err_Routine
    CBA_Error = ""
    Set CN = New ADODB.Connection
    With CN
        .ConnectionTimeout = 50
        .CommandTimeout = 50
        .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Cam")
    End With
    If sdLineCountDic Is Nothing Then
        Set sdLineCountDic = New Scripting.Dictionary
    End If
    sdLineCountDic.Add eMAT, CreateLineCountObject(eMAT, CN)
    sdLineCountDic.Add ePriorMAT, CreateLineCountObject(ePriorMAT, CN)
    sdLineCountDic.Add eYTD, CreateLineCountObject(eYTD, CN)
    sdLineCountDic.Add ePriorYTD, CreateLineCountObject(ePriorYTD, CN)
    sdLineCountDic.Add eQTRTD, CreateLineCountObject(eQTRTD, CN)
    sdLineCountDic.Add ePriorQTR, CreateLineCountObject(ePriorQTR, CN)
    'sdLineCountDic.add eCurrentMonth, createLineCountObject(eCurrentMonth, CN)
    'sdLineCountDic.add ePriorMonth, createLineCountObject(ePriorMonth, CN)
    sdLineCountDic.Add ePriorCalendarYr, CreateLineCountObject(ePriorCalendarYr, CN)
    sdLineCountDic.Add ePrevPriorCalendarYr, CreateLineCountObject(ePrevPriorCalendarYr, CN)
    If CN.State <> 0 Then CN.Close
    Set CN = Nothing
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Category.createAllLineObjectsForAllPeriods", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Function CreateUserList() As Boolean    ' @RWCam - Took from runtime
    Dim CN As ADODB.Connection, RS As New ADODB.Recordset  ' @RWCam - Needs CN handling
    Dim strSQL As String
    On Error GoTo Err_Routine
    CBA_Error = ""

    If psdUserList Is Nothing Then
''        If RunConnectionSetups(True, False, False, False, False) = False Then Exit Function
        strSQL = "select empno, firstname + ' ' + name as Name--, PositionNo" & Chr(10)
        strSQL = strSQL & ",case when positionno = 43 then 'GBD'" & Chr(10)
        strSQL = strSQL & "when positionno = 5 then 'BD'" & Chr(10)
        strSQL = strSQL & "when positionno = 6 then 'BA'" & Chr(10)
        strSQL = strSQL & "when positionno = 27 then 'BM'" & Chr(10)
        strSQL = strSQL & "when positionno = 24 then 'QAA'" & Chr(10)
        strSQL = strSQL & "when positionno = 30 then 'MktM'" & Chr(10)
        strSQL = strSQL & "when positionno = 22 then 'CRD/QAD'" & Chr(10)
        strSQL = strSQL & "when positionno = 25 then 'CRM'" & Chr(10)
        strSQL = strSQL & "when positionno = 31 then 'MktA'" & Chr(10)
        strSQL = strSQL & "when positionno = 42 then 'MerchA'" & Chr(10)
        strSQL = strSQL & "when positionno = 41 then 'MerchM'" & Chr(10)
        strSQL = strSQL & "when positionno = 28 then 'CRA'" & Chr(10)
        strSQL = strSQL & "end As Role" & Chr(10)
        strSQL = strSQL & "from cbis599p.dbo.EMPLOYEE" & Chr(10)
        strSQL = strSQL & "where (IsInactive = 0 And empsign Is Not Null And PositionNo Is Not Null And empno_grp Is Not Null And EmpNo <> 999999)" & Chr(10)
        strSQL = strSQL & "or (positionno = 43 and IsInactive = 0 and empsign is not null )" & Chr(10)
        strSQL = strSQL & "order by firstname" & Chr(10)
        Set RS = New ADODB.Recordset
        RS.Open strSQL, CN
        If RS.EOF Then
        Else
            Set psdUserList = New Scripting.Dictionary
            Set psdEtoNUList = New Scripting.Dictionary
            Do Until RS.EOF
                If psdUserList.Exists(CStr(RS(1))) = False Then psdUserList.Add CStr(RS(1)), CStr(RS(0))
                If psdEtoNUList.Exists(CStr(RS(0))) = False Then psdEtoNUList.Add CStr(RS(0)), CStr(RS(1))
                RS.MoveNext
            Loop
        End If
        Set RS = Nothing
    End If

    If psdUserList.Count > -1 And psdEtoNUList.Count > -1 Then CreateUserList = True
Exit_Routine:

    On Error Resume Next
    CN.Close
    Set CN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("sf-mCam_Runtime.CreateUserList", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Function ChangePermissions(ByRef Names As Scripting.Dictionary, lCurB_ID As Long) As Boolean
    ' Change Permissions on the fCam_Permissions form
    Dim U As Variant
    Dim strSQL As String
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim a As Long
    On Error GoTo Err_Routine
    CBA_Error = ""
    With CN
        .ConnectionTimeout = 50
        .CommandTimeout = 50
        .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Cam")
    End With

    If lCurB_ID > -1 Then
        ChangePermissions = True
''        If RunConnectionSetups(False, False, False, False, True) = True Then        ' @RWCam - Took Out
            For Each U In Names.Keys
                If Names(U) = 1 Then
                    If psdUserList(U) <> "" Then
                        Set RS = New ADODB.Recordset
                        strSQL = "UPDATE BASE set CR_EmpNoOwner = " & psdUserList(U) & " where ID = " & lCurB_ID
                        RS.Open strSQL, CN
                        Set RS = Nothing
                    Else
                        MsgBox "EmpNo not found for Owner, Changes not made", vbOKOnly: ChangePermissions = False: Exit Function
                    End If
                    Exit For
                End If
            Next

            Set RS = New ADODB.Recordset
            strSQL = "select * from USR_ASSIGN where ID = " & lCurB_ID
            RS.Open strSQL, CN, adOpenDynamic, adLockOptimistic
            If RS.EOF Then
                MsgBox "No Record returned for Assigned Personel, Changes not made", vbOKOnly: ChangePermissions = False: Exit Function
            Else
                For a = 2 To RS.Fields.Count - 1
                    RS.Update a, Null
                Next
                a = 1
                For Each U In Names.Keys
                    If Names(U) = 0 Then
                        a = a + 1
                        RS.Update a, psdUserList(U)
                        Debug.Print RS(a)
                    End If
                Next
            End If
            Set RS = New ADODB.Recordset
            strSQL = "SELECT BASE.*, USR_ASSIGN.* FROM BASE INNER JOIN USR_ASSIGN ON BASE.ID = USR_ASSIGN.B_ID;"
            RS.Open strSQL, CN
            pvBaseUsers = RS.GetRows()
            ''CloseConnections False, False, False, False, True
''        End If
    End If
Exit_Routine:

    On Error Resume Next
    CN.Close
    Set CN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("sf-mCam_Runtime.ChangePermissions", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function


Public Property Get cPGrp() As cCBA_ProdGroup: Set cPGrp = pcPGrp: End Property
Private Property Set cPGrp(ByVal objNewValue As cCBA_ProdGroup): Set pcPGrp = objNewValue: End Property

Private Property Get sdMSegs() As Scripting.Dictionary: Set sdMSegs = psdMSegs: End Property
Private Property Set sdMSegs(ByVal objNewValue As Scripting.Dictionary): Set psdMSegs = objNewValue: End Property

Public Property Get sCategoryName() As String: sCategoryName = psCategoryName: End Property
Private Property Let sCategoryName(ByVal sNewValue As String): psCategoryName = sNewValue: End Property

Public Property Get bACG() As Boolean: bACG = pbACG: End Property
Private Property Let bACG(ByVal bNewValue As Boolean): pbACG = bNewValue: End Property

Private Property Get bCRTrialPerf() As Boolean: bCRTrialPerf = pbCRTrialPerf: End Property
Private Property Let bCRTrialPerf(ByVal bNewValue As Boolean): pbCRTrialPerf = bNewValue: End Property

Private Property Get bLineCount() As Boolean: bLineCount = pbLineCount: End Property
Private Property Let bLineCount(ByVal bNewValue As Boolean): pbLineCount = bNewValue: End Property

Private Property Get bCatStrat() As Boolean: bCatStrat = pbCatStrat: End Property
Private Property Let bCatStrat(ByVal bNewValue As Boolean): pbCatStrat = bNewValue: End Property

Private Property Get bCorpResp() As Boolean: bCorpResp = pbCorpResp: End Property
Private Property Let bCorpResp(ByVal bNewValue As Boolean): pbCorpResp = bNewValue: End Property

Private Property Get bExecSum() As Boolean: bExecSum = pbExecSum: End Property
Private Property Let bExecSum(ByVal bNewValue As Boolean): pbExecSum = bNewValue: End Property

Private Property Get bCRPerf() As Boolean: bCRPerf = pbCRPerf: End Property
Private Property Let bCRPerf(ByVal bNewValue As Boolean): pbCRPerf = bNewValue: End Property

Private Property Get bForecast() As Boolean: bForecast = pbForecast: End Property
Private Property Let bForecast(ByVal bNewValue As Boolean): pbForecast = bNewValue: End Property

Private Property Get bSuppBase() As Boolean: bSuppBase = pbSuppBase: End Property
Private Property Let bSuppBase(ByVal bNewValue As Boolean): pbSuppBase = bNewValue: End Property

Private Property Get bCatPlano() As Boolean: bCatPlano = pbCatPlano: End Property
Private Property Let bCatPlano(ByVal bNewValue As Boolean): pbCatPlano = bNewValue: End Property

Private Property Get bCatPerfOver() As Boolean: bCatPerfOver = pbCatPerfOver: End Property
Private Property Let bCatPerfOver(ByVal bNewValue As Boolean): pbCatPerfOver = bNewValue: End Property

Private Property Get bSeasPerf() As Boolean: bSeasPerf = pbSeasPerf: End Property
Private Property Let bSeasPerf(ByVal bNewValue As Boolean): pbSeasPerf = bNewValue: End Property

Private Property Get bInterComp() As Boolean: bInterComp = pbInterComp: End Property
Private Property Let bInterComp(ByVal bNewValue As Boolean): pbInterComp = bNewValue: End Property

Private Property Get bCRPerfProds() As Boolean: bCRPerfProds = pbCRPerfProds: End Property
Private Property Let bCRPerfProds(ByVal bNewValue As Boolean): pbCRPerfProds = bNewValue: End Property

Private Property Get bCommUpd() As Boolean: bCommUpd = pbCommUpd: End Property
Private Property Let bCommUpd(ByVal bNewValue As Boolean): pbCommUpd = bNewValue: End Property

Private Property Get bTotCatPerf() As Boolean: bTotCatPerf = pbTotCatPerf: End Property
Private Property Let bTotCatPerf(ByVal bNewValue As Boolean): pbTotCatPerf = bNewValue: End Property

Private Property Get bSpecPerf() As Boolean: bSpecPerf = pbSpecPerf: End Property
Private Property Let bSpecPerf(ByVal bNewValue As Boolean): pbSpecPerf = bNewValue: End Property

Private Property Get bCompAnal() As Boolean: bCompAnal = pbCompAnal: End Property
Private Property Let bCompAnal(ByVal bNewValue As Boolean): pbCompAnal = bNewValue: End Property

Private Property Get bQA() As Boolean: bQA = pbQA: End Property
Private Property Let bQA(ByVal bNewValue As Boolean): pbQA = bNewValue: End Property

Private Property Get bKPIs() As Boolean: bKPIs = pbKPIs: End Property
Private Property Let bKPIs(ByVal bNewValue As Boolean): pbKPIs = bNewValue: End Property

Private Property Get bMarkOver() As Boolean: bMarkOver = pbMarkOver: End Property
Private Property Let bMarkOver(ByVal bNewValue As Boolean): pbMarkOver = bNewValue: End Property

Private Property Get bWastage() As Boolean: bWastage = pbWastage: End Property
Private Property Let bWastage(ByVal bNewValue As Boolean): pbWastage = bNewValue: End Property

Private Property Get bCatRevActions() As Boolean: bCatRevActions = pbCatRevActions: End Property
Private Property Let bCatRevActions(ByVal bNewValue As Boolean): pbCatRevActions = bNewValue: End Property

Private Property Get bCRTrialProdList() As Boolean: bCRTrialProdList = pbCRTrialProdList: End Property
Private Property Let bCRTrialProdList(ByVal bNewValue As Boolean): pbCRTrialProdList = bNewValue: End Property

Private Property Get bRdyToRun() As Boolean: bRdyToRun = pbRdyToRun: End Property
Private Property Let bRdyToRun(ByVal bNewValue As Boolean): pbRdyToRun = bNewValue: End Property

Public Property Get dtDateFrom() As Date: dtDateFrom = pdtDateFrom: End Property
Private Property Let dtDateFrom(ByVal dtNewValue As Date): pdtDateFrom = dtNewValue: End Property

Public Property Get dtDateTo() As Date: dtDateTo = pdtDateTo: End Property
Private Property Let dtDateTo(ByVal dtNewValue As Date): pdtDateTo = dtNewValue: End Property

Public Property Get sdLineCountDic() As Scripting.Dictionary: Set sdLineCountDic = psdLineCountDic: End Property
Private Property Set sdLineCountDic(ByVal objNewValue As Scripting.Dictionary): Set psdLineCountDic = objNewValue: End Property

Public Property Get lDoc_ID() As Long: lDoc_ID = plDoc_ID: End Property
Private Property Let lDoc_ID(ByVal lNewValue As Long): plDoc_ID = lNewValue: End Property

Private Property Get sdCategoryList() As Scripting.Dictionary: Set sdCategoryList = psdCategoryList: End Property
Private Property Set sdCategoryList(ByVal objNewValue As Scripting.Dictionary): Set psdCategoryList = objNewValue: End Property

Public Property Get eDocumentType() As e_DocuType: eDocumentType = peDocumentType: End Property
Private Property Let eDocumentType(ByVal eNewValue As e_DocuType): peDocumentType = eNewValue: End Property

Public Property Get vBaseUsers() As Variant(): vBaseUsers = pvBaseUsers: End Property
Private Property Let vBaseUsers(NewValue() As Variant): pvBaseUsers = NewValue: End Property

Public Property Get sdUserList() As Scripting.Dictionary: Set sdUserList = psdUserList: End Property
Private Property Set sdUserList(ByVal objNewValue As Scripting.Dictionary): Set psdUserList = objNewValue: End Property

Public Property Get sdEToNUList() As Scripting.Dictionary: Set sdEToNUList = psdEtoNUList: End Property
Private Property Set sdEToNUList(ByVal objNewValue As Scripting.Dictionary): Set psdEtoNUList = objNewValue: End Property

Private Property Get bIsActive() As Boolean: bIsActive = pbIsActive: End Property
Private Property Let bIsActive(ByVal bNewValue As Boolean): pbIsActive = bNewValue: End Property

Public Property Get sBuildCode() As String: sBuildCode = psBuildCode: End Property

Private Property Let sBuildCode(ByVal sNewValue As String): psBuildCode = sNewValue: End Property

Private Property Get cFCast() As cCBA_Forecasting: Set cFCast = pcFCast: End Property

Private Property Set cFCast(ByVal objNewValue As cCBA_Forecasting): Set pcFCast = objNewValue: End Property

Private Property Get eDefaultMSegtype() As e_MSegType: eDefaultMSegtype = peDefaultMSegtype: End Property

Private Property Let eDefaultMSegtype(ByVal NewValue As e_MSegType): peDefaultMSegtype = NewValue: End Property

Private Property Get sdNData() As Scripting.Dictionary: Set sdNData = psdNData: End Property

Private Property Set sdNData(ByVal NewValue As Scripting.Dictionary): Set psdNData = NewValue: End Property
