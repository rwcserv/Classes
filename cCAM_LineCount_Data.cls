VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCAM_LineCount_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit                                         ' cCAM_LineCount_Data
Private psdProductAllocation As Scripting.Dictionary
Private psdCalculatedProductAllocation As Scripting.Dictionary
Private pbIsChangedNotSaved As Boolean
Private pdtSnapShotDate As Date
Private peLCPeriod As e_LineCountPeriod
Private pbIsDocumentNotRibbonObject As Boolean
Private plDoc_ID As Long
Public Function Copy(ByRef orig As cCAM_LineCount_Data) As Boolean
Dim v As Variant
    Set psdProductAllocation = New Scripting.Dictionary
    For Each v In orig.sdProductAllocation
        psdProductAllocation.Add CStr(v), CLng(orig.sdProductAllocation(v))
    Next
    Set psdCalculatedProductAllocation = New Scripting.Dictionary
    For Each v In orig.sdCalculatedProductAllocation
        psdCalculatedProductAllocation.Add CStr(v), CLng(orig.sdCalculatedProductAllocation(v))
    Next
    pbIsChangedNotSaved = orig.IsChangedNotSaved
    pdtSnapShotDate = orig.dtSnapShotDate
    peLCPeriod = orig.eLCPeriod
    pbIsDocumentNotRibbonObject = orig.IsDocumentNotRibbonObject
    If Err.Number = 0 Then Copy = True
End Function
Public Function BuildObject(ByVal CategoryName As String, ByVal FinalDateOfPeriod As Date, ByRef PG As cCBA_ProdGroup, ByVal thisPeriod As e_LineCountPeriod, Optional ByRef CN As ADODB.Connection, Optional ByVal DocID As Long = 0) As Boolean
' Tells the object what it is and allows it get the data and create itself
'  It is understood that this method where snapshot date and e_linecountperiod could produce duplicate objects (i.e 2018 PYMAT is same as 2017 MAT)
'  but its considered a minimal performace issue. Needs to be corrected if proven otherwise
    eLCPeriod = thisPeriod
    dtSnapShotDate = FinalDateOfPeriod
    If CN Is Nothing Then CN = CreateConnection
    BuildObject = GetLineCountData(CategoryName, PG, CN, DocID)
End Function
Private Function CreateConnection() As ADODB.Connection
Dim CN As ADODB.Connection
Set CN = New ADODB.Connection
With CN
    .ConnectionTimeout = 50
    .CommandTimeout = 50
    .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Cam")
End With
Set CreateConnection = CN
End Function
Function GetLineCountData(ByVal CategoryName As String, ByRef PG As cCBA_ProdGroup, ByRef CN As ADODB.Connection, ByVal DocID As Long) As Boolean
    Dim RS As ADODB.Recordset, strSQL As String, DateForQuery As Date, sPcodeAndDesc As String, lLCType As e_LineCountType
    On Error GoTo Err_Routine
    CBA_Error = ""
    If CategoryName = "" Then GetLineCountData = False: Exit Function
    If DocID > 0 Then IsDocumentNotRibbonObject = True
    Set sdCalculatedProductAllocation = CreateInitalLineCount(PG)
    Set RS = New ADODB.Recordset
''    strSQL = "select productcode, linecounttype from CR_LineCountallocation LCA" & Chr(10)
''    strSQL = strSQL & "left join g_CategoryName CN ON LCA.CAT_ID = CN.CAT_ID" & Chr(10)
''    strSQL = strSQL & "where LCA.CATREV_ID = " & DocID & Chr(10)
''    strSQL = strSQL & "and CN.CategoryName = '" & CategoryName & "'" & Chr(10)
    strSQL = "SELECT LA_ProductCode, LA_LineCountType FROM L1_LineCountAllocation LCA" & Chr(10)
    strSQL = strSQL & "LEFT JOIN L0_CategoryName CN ON LCA.LA_CAT_ID = CN.CN_ID" & Chr(10)
    strSQL = strSQL & "WHERE LCA.LA_CATREV_ID = " & DocID & Chr(10)
    strSQL = strSQL & "AND CN.CN_CategoryName = '" & CategoryName & "'" & Chr(10)
    ''DocID is the ID in the CAMERA system of the specific category review. (i.e. the Document) this may end up being a duplicate of the DOC_ID somewhere or an alternated ID. This will probably have to be addressed
    ''as there may be a different line count setup as part of a document to the default, this is necessary
    '' if CATREV_ID =0 then its looking get the default
    Select Case eLCPeriod
        Case eMAT
            strSQL = strSQL & "and LCA.LA_SnapShotDate <= #" & Format(dtSnapShotDate, "MM/DD/YYYY") & "#" & Chr(10)
            strSQL = strSQL & "and LCA.LA_SnapShotDate >= #" & Format(DateAdd("M", -8, dtSnapShotDate), "MM/DD/YYYY") & "#" & Chr(10)
        Case ePriorMAT
            strSQL = strSQL & "and LCA.LA_SnapShotDate <= #" & Format(DateAdd("YYYY", -1, dtSnapShotDate), "MM/DD/YYYY") & "#" & Chr(10)
            strSQL = strSQL & "and LCA.LA_SnapShotDate >= #" & Format(DateAdd("M", -8, DateAdd("YYYY", -1, dtSnapShotDate)), "MM/DD/YYYY") & "#" & Chr(10)
        Case eQTRTD
            DateForQuery = DateSerial(Year(dtSnapShotDate), Int((Month(dtSnapShotDate) - 1) / 3) * 3 + 1, 1)
            strSQL = strSQL & "and LCA.LA_SnapShotDate <= #" & Format(dtSnapShotDate, "MM/DD/YYYY") & "#" & Chr(10)
            strSQL = strSQL & "and LCA.LA_SnapShotDate >= #" & Format(DateForQuery, "MM/DD/YYYY") & "#" & Chr(10)
        Case ePriorQTR
            DateForQuery = DateSerial(Year(dtSnapShotDate), Int((Month(dtSnapShotDate) - 1) / 3) * 3 + 1, 0)
            strSQL = strSQL & "and LCA.LA_SnapShotDate <= #" & Format(DateForQuery, "MM/DD/YYYY") & "#" & Chr(10)
            strSQL = strSQL & "and LCA.LA_SnapShotDate >= #" & Format(DateSerial(Year(DateForQuery), Int((Month(DateForQuery) - 1) / 3) * 3 + 1, 1), "MM/DD/YYYY") & "#" & Chr(10)
        Case eCurrentMonth, ePriorMonth
'            'really this needs to find the most recent which will envolve a max statement e.t.c. can look at later
'            If eLCPeriod = eCurrentMonth Then DateForQuery = Date Else DateForQuery = dtSnapShotDate
'            strSQL = strSQL & "and Max(LCA.LA_SnapShotDate) <= #" & Format(DateForQuery, "MM/DD/YYYY") & "#" & Chr(10)
'            strSQL = strSQL & "and Max(LCA.LA_SnapShotDate) >= #" & Format(DateAdd("M", -6, DateForQuery), "MM/DD/YYYY") & "#" & Chr(10)
        Case ePriorCalendarYr, ePrevPriorCalendarYr
            strSQL = strSQL & "and LCA.LA_SnapShotDate <= #" & Format(DateSerial(Year(dtSnapShotDate) - IIf(eLCPeriod = ePriorCalendarYr, 1, 2), 12, 31), "MM/DD/YYYY") & "#" & Chr(10)
            strSQL = strSQL & "and LCA.LA_SnapShotDate >= #" & Format(DateSerial(Year(dtSnapShotDate) - IIf(eLCPeriod = ePriorCalendarYr, 1, 2), 1, 1), "MM/DD/YYYY") & "#" & Chr(10)
    End Select
    strSQL = strSQL & "ORDER BY LA_ProductCode ;"
    RS.Open strSQL, CN
    If sdProductAllocation Is Nothing Then Set sdProductAllocation = New Scripting.Dictionary
    If RS.EOF Then Set sdProductAllocation = sdCalculatedProductAllocation
    Do Until RS.EOF
        sPcodeAndDesc = CStr(RS(0))
        If sdProductAllocation.Exists(sPcodeAndDesc) Then
            lLCType = sdProductAllocation(sPcodeAndDesc)
            If lLCType = eBranded And CLng(RS(1)) = eSucTrial Then              ' If it is already eBranded
                sdProductAllocation(sPcodeAndDesc) = eBrandSucTrial
            ElseIf lLCType = eSucTrial And CLng(RS(1)) = eBranded Then          ' If it is already eSucTrial
                sdProductAllocation(sPcodeAndDesc) = eBrandSucTrial
            ElseIf lLCType = eBranded Or lLCType = eSucTrial Or lLCType = eBrandSucTrial Then
            Else
                sdProductAllocation(sPcodeAndDesc) = CLng(RS(1))
            End If
        Else
            sdProductAllocation.Add sPcodeAndDesc, CLng(RS(1))
        End If
        RS.MoveNext
    Loop
    GetLineCountData = True
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_LineCount_Data.GetLineCountData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error & vbCrLf & strSQL
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function
Public Function Allocate(ByVal sPcodeAndDesc As String, ByVal LineCountType As e_LineCountType) As Boolean
    Dim lLCType As e_LineCountType
    On Error GoTo Err_Routine
    CBA_Error = ""
    ' Allocates a product to a LCType. If the product is already allocated to a type in the dictionary then it is changed, else it is added.
    If sdProductAllocation.Exists(sPcodeAndDesc) Then
        If sdProductAllocation(sPcodeAndDesc) = LineCountType Then Allocate = False: Exit Function
''        sdProductAllocation(sPCodeAndDesc) = LineCountType
        lLCType = sdProductAllocation(sPcodeAndDesc)
        If lLCType = eBranded And LineCountType = eSucTrial Then              ' If it is already eBranded
            sdProductAllocation(sPcodeAndDesc) = eBrandSucTrial
        ElseIf lLCType = eSucTrial And LineCountType = eBranded Then          ' If it is already eSucTrial
            sdProductAllocation(sPcodeAndDesc) = eBrandSucTrial
        '''ElseIf lLCType = eBranded Or lLCType = eSucTrial Or lLCType = eBrandSucTrial Then
        ElseIf lLCType = eBrandSucTrial And (LineCountType = eBranded Or LineCountType = eSucTrial) Then           ' If it is already eBrandSucTrial
            'sdProductAllocation(sPcodeAndDesc) = eBrandSucTrial
        Else
            sdProductAllocation(sPcodeAndDesc) = LineCountType
        End If

        IsChangedNotSaved = True
    Else
        sdProductAllocation.Add sPcodeAndDesc, LineCountType
        IsChangedNotSaved = True
    End If
    Allocate = True
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_LineCount_Data.Allocate", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
''    Debug.Print CBA_Error & vbCrLf & strSQL
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Public Function LineCountReset() As Boolean
    IsChangedNotSaved = False
    LineCountReset = IsChangedNotSaved
End Function
Public Function SaveLineCountToDB(Optional ByVal SendToMaster As Boolean = False) As Boolean
    Dim PCodeAndDesc As String
    Dim CN As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim strSQL As String
    Dim CN_ID As Long
    Dim CurAll As Scripting.Dictionary
    Dim a As Long, b As Long
    Dim RsToBeUpdated As Boolean
    Dim v As Variant
    On Error GoTo Err_Routine
    CBA_Error = ""
    Set CurAll = New Scripting.Dictionary
    Set CN = CreateConnection
    Set RS = New ADODB.Recordset
    strSQL = "SELECT CN_ID from L0_CategoryName where CN_isActive = TRUE and CN_CategoryName = '" & mCAM_Runtime.CategoryObject(lDoc_ID).sCategoryName & "'"
    RS.Open strSQL, CN
    If RS.EOF Then MsgBox "Error in cCAM_LineCount_Data: SaveLineCountToDB cannot find CN_ID": SaveLineCountToDB = False: Exit Function
    CN_ID = CLng(RS(0))
    Set RS = New ADODB.Recordset
    strSQL = "select * from L1_LineCountAllocation where LA_CAT_ID = " & CN_ID & Chr(10)
    If SendToMaster = False Then strSQL = strSQL & "and LA_CATREV_ID = " & IIf(lDoc_ID < 200, 0, lDoc_ID) & Chr(10) Else strSQL = strSQL & "and LA_CATREV_ID = 0" & Chr(10)
    strSQL = strSQL & "and LA_Period_ID = " & eLCPeriod & Chr(10)
    strSQL = strSQL & "and LA_SnapShotDate = #" & Format(dtSnapShotDate, "MM/DD/YYYY") & "#" & Chr(10)
    RS.Open strSQL, CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
        a = 0
        Do Until RS.EOF
            a = a + 1
            CurAll.Add CStr(RS(3)), CLng(a)
            RS.MoveNext
        Loop
    End If
    If IsChangedNotSaved = True Then
        For Each v In sdProductAllocation
            If InStr(1, v, "-") > 0 Then
                PCodeAndDesc = CStr(v)
            ElseIf v = "" Then
                GoTo NextPallocation
            Else
                MsgBox "Error found in sdProductAllocation Dictionary in cCAM_LineCount_Data", vbOKOnly
                Stop
            End If
            If CurAll.Exists(PCodeAndDesc) Then
                RS.MoveFirst
                For a = 1 To CurAll(PCodeAndDesc)
                    If a < CurAll(PCodeAndDesc) Then RS.MoveNext
                Next
                If RS(4) <> sdProductAllocation(v) Then
                    RS(4) = CLng(sdProductAllocation(v))
                    RsToBeUpdated = True
                End If
            Else
                RS.AddNew Array("LA_CAT_ID", "LA_CATREV_ID", "LA_Productcode", "LA_LineCountType", "LA_Period_ID", "LA_SnapShotDate"), Array(CN_ID, IIf(lDoc_ID < 200, 0, lDoc_ID), PCodeAndDesc, CLng(sdProductAllocation(v)), eLCPeriod, dtSnapShotDate)
                RsToBeUpdated = True
            End If
NextPallocation:
        Next
        If RsToBeUpdated = True Then RS.Update
        IsChangedNotSaved = False
        Set RS = Nothing
        CN.Close
        Set CN = Nothing
        SaveLineCountToDB = True
    End If
    
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_LineCount_Data.SaveLineCountToDB", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error & vbCrLf & strSQL
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next


End Function


Function CreateInitalLineCount(ByRef PG As cCBA_ProdGroup) As Scripting.Dictionary
    ' Looks at the ProdGroup and allocates out the line count as per base logic. The allocation sits in a Dic(Key(Pcode-Pdescription),Item:e_LineCountType)
    Dim prod As Variant, arr As Variant
    Dim CP As cCBA_Prod
    Dim dic As Scripting.Dictionary
    Dim a As Long
    Dim s_Val As Single
    Dim col As Collection
    Dim DateFrom As Date, DateTo As Date
    On Error GoTo Err_Routine
    CBA_Error = ""
    
    Set dic = New Scripting.Dictionary
    Set col = PG.getProdListing
    For Each prod In col
        Set CP = PG.getProdObject(prod)
        Select Case eLCPeriod
            Case e_LineCountPeriod.eMAT
                DateTo = dtSnapShotDate
                DateFrom = DateAdd("M", -8, dtSnapShotDate)
            Case e_LineCountPeriod.ePriorMAT
                DateTo = DateAdd("YYYY", -1, dtSnapShotDate)
                DateFrom = DateAdd("M", -8, DateTo)
            Case e_LineCountPeriod.eYTD
                DateTo = dtSnapShotDate
                DateFrom = DateSerial(Year(DateTo), 1, 1)
            Case e_LineCountPeriod.ePriorYTD
                DateTo = DateAdd("YYYY", -1, dtSnapShotDate)
                DateFrom = DateSerial(Year(DateTo) - 1, 1, 1)
            Case e_LineCountPeriod.eQTRTD
                DateTo = dtSnapShotDate
                DateFrom = DateSerial(Year(dtSnapShotDate), Int((Month(dtSnapShotDate) - 1) / 3) * 3 + 1, 1)
            Case e_LineCountPeriod.ePriorQTR
                DateTo = DateSerial(Year(dtSnapShotDate), Int((Month(dtSnapShotDate) - 1) / 3) * 3 + 1, 0)
                DateFrom = DateSerial(Year(DateTo), Int((Month(DateTo) - 1) / 3) * 3 + 1, 0)
            Case e_LineCountPeriod.eCurrentMonth, e_LineCountPeriod.ePriorMonth
                'really this needs to find the most recent which will envolve a max statement e.t.c. can look at later
                If eLCPeriod = e_LineCountPeriod.eCurrentMonth Then
                    DateTo = Date
                    DateFrom = DateSerial(Year(Date), Int((Month(Date) - 1) / 3) * 3 + 1, 1)
                Else
                    DateTo = dtSnapShotDate
                    DateFrom = DateSerial(Year(dtSnapShotDate), Int((Month(dtSnapShotDate) - 1) / 3) * 3 + 1, 1)
                End If
            Case e_LineCountPeriod.ePriorCalendarYr, e_LineCountPeriod.ePrevPriorCalendarYr
                DateTo = DateSerial(Year(dtSnapShotDate) - IIf(eLCPeriod = e_LineCountPeriod.ePriorCalendarYr, 1, 2), 12, 31)
                DateFrom = DateSerial(Year(dtSnapShotDate) - IIf(eLCPeriod = e_LineCountPeriod.ePriorCalendarYr, 1, 2), 1, 1)
        End Select
        If CP.lPClass = 2 Or CP.lPClass = 3 Then
            arr = CP.getContractData("UDt", , DateAdd("M", -1, DateFrom)) ' Asking if there ws an active contract in the period (moving the datefrom back 1 month due to specials deliveryfrom dates
            If arr(0, 0) > 0 Then
                ''If so and there POS retail over 5K AUD and its not a consolidated product
                If CP.getPOSdata(DateFrom, DateTo, False) > 5000 And CP.lCon_PCode = 0 Then dic.Add prod & "-" & CP.sProdDesc, e_LineCountType.eSpecial
            End If
        Else
            arr = CP.getContractData("UDt", , DateFrom) ' Asking if there ws an active contract in the period
            If arr(0, 0) = 0 Then 'if not then
                If CP.lCon_PCode = 0 Then dic.Add prod & "-" & CP.sProdDesc, e_LineCountType.eDeleted
            Else
                If CP.getPOSdata(DateFrom, DateTo, True) < 500 Then
                    dic.Add prod & "-" & CP.sProdDesc, e_LineCountType.eDeleted
                ElseIf CP.getProdLabel("Regional") = True Then
                    dic.Add prod & "-" & CP.sProdDesc, e_LineCountType.eRegional
                ElseIf CP.getProdLabel("Trial") = True Then
                    dic.Add prod & "-" & CP.sProdDesc, e_LineCountType.eCurTrial
                ElseIf CP.getProdLabel("Branded") = True Then
                    dic.Add prod & "-" & CP.sProdDesc, e_LineCountType.eBranded
                ElseIf CP.lPClass = 1 Then
                    dic.Add prod & "-" & CP.sProdDesc, e_LineCountType.eCore
                ElseIf CP.lPClass = 4 Then
                    dic.Add prod & "-" & CP.sProdDesc, e_LineCountType.eSeasonal
                End If
            End If
        End If
TryTheNextOne:
    Next
    Set CreateInitalLineCount = dic
Exit_Routine:
    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_LineCount_Data.CreateInitalLineCount", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Public Property Get sdProductAllocation() As Scripting.Dictionary: Set sdProductAllocation = psdProductAllocation: End Property
Private Property Set sdProductAllocation(ByVal objNewValue As Scripting.Dictionary): Set psdProductAllocation = objNewValue: End Property
Public Property Get sdCalculatedProductAllocation() As Scripting.Dictionary: Set sdCalculatedProductAllocation = psdCalculatedProductAllocation: End Property
Private Property Set sdCalculatedProductAllocation(ByVal objNewValue As Scripting.Dictionary): Set psdCalculatedProductAllocation = objNewValue: End Property
Public Property Get IsChangedNotSaved() As Boolean: IsChangedNotSaved = pbIsChangedNotSaved: End Property
Public Property Let IsChangedNotSaved(ByVal bNewValue As Boolean): pbIsChangedNotSaved = bNewValue: End Property
Public Property Get dtSnapShotDate() As Date: dtSnapShotDate = pdtSnapShotDate: End Property
Private Property Let dtSnapShotDate(ByVal dtNewValue As Date): pdtSnapShotDate = dtNewValue: End Property
Public Property Get eLCPeriod() As e_LineCountPeriod: eLCPeriod = peLCPeriod: End Property
Private Property Let eLCPeriod(ByVal eNewValue As e_LineCountPeriod): peLCPeriod = eNewValue: End Property
Public Property Get IsDocumentNotRibbonObject() As Boolean: IsDocumentNotRibbonObject = pbIsDocumentNotRibbonObject: End Property
Private Property Let IsDocumentNotRibbonObject(ByVal bNewValue As Boolean): pbIsDocumentNotRibbonObject = bNewValue: End Property
Public Property Get lDoc_ID() As Long: lDoc_ID = plDoc_ID: End Property
Public Property Let lDoc_ID(ByVal NewValue As Long): plDoc_ID = NewValue: End Property
