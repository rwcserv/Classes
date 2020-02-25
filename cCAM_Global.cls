VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCAM_Global"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit                 ' cCAM_Global

Private Const NATIONAL_DB As Integer = 599
Private psdCategory As Scripting.Dictionary       'Dic(Key:CategoryName,Item:Collection of cCBA_Prod)
Private psdLCGList As Scripting.Dictionary        'Dic(Key:Format(CGno,"000")&Format(SCGNo,"000"),Item:cCBA_Prod)
Private psdACGList As Scripting.Dictionary        'Dic(Key:Format(ACGno,"000")&Format(CGno,"000")&Format(SCGNo,"000"),Item:cCBA_Prod)
Private pbIsCGListActive As Boolean
Private psdMSegDic As Scripting.Dictionary        'Dic(Key:CategoryName,Item:Collection of cCBA_Nielsen)
Private psActiveRibbonCategory As String
Private pdtDateTo As Date
Private pbIsActive As Boolean
Private pbCategoryDicIsChanged As Boolean
Private pbyACGSelected As Byte
Private plCurCatIndexSelected As Long
Private psdCategoryID As Scripting.Dictionary
    ' Holds data specific to the CAMERA System in general.
    ' Where there are global dictionaries etc that are required then they are stored here.
    ' If it is Category Specific, i.e. for one category, then it is stored in the CAMERA_Data object(this includes for the selected category in the ribbon, this is why there is no ProdtoMSeg Stored in this object)

Public Function GetCategoryList(ByVal CategoryName As String, ByVal ACG As Boolean) As Scripting.Dictionary
    ' Returns the collection of cCBA_Prod Objects (getsLetsOnly) for that category
    Dim sdDic As Scripting.Dictionary
    Dim P As cCBA_Prod
    On Error GoTo Err_Routine
    CBA_Error = ""
    If sdCategoryDic.Exists(CategoryName) Then
        Set sdDic = New Scripting.Dictionary
        For Each P In sdCategoryDic(CategoryName)
            If ACG And P.lACGNum > 0 Then
                sdDic.Add Format(P.lACatNum, "000") & Format(P.lACGNum, "000") & Format(P.lASCGNum, "000"), P
            ElseIf ACG = False And P.lACGNum = 0 Then
                sdDic.Add Format(P.lLegCG, "000") & Format(P.lLegSCG, "000"), P
            End If
        Next
        If sdDic.Count > 0 Then Set GetCategoryList = sdDic Else Set GetCategoryList = Nothing
    Else
        Set GetCategoryList = Nothing
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Global.getCategoryList", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function GetAllCategoryNames() As Variant
    ' This steps through the sdCategoryDic, takes all the keys (CategoryName(S)) and returns a One-Dimentional Array
    Dim arr() As String
    Dim a As Long
    Dim v As Variant
    Dim P As cCBA_Prod
    On Error GoTo Err_Routine
    CBA_Error = ""
    If sdCategoryDic Is Nothing Then
        ReDim arr(0 To 0): arr(0) = False: GetAllCategoryNames = arr
    Else
        a = -1
        For Each v In sdCategoryDic.Keys
            If sdCategoryDic(v).Count > 0 Then a = a + 1
        Next
        ReDim arr(0 To CLng(a))
        a = -1
        For Each v In sdCategoryDic.Keys
            If sdCategoryDic(v).Count > 0 Then a = a + 1: arr(a) = CStr(v)
        Next
    End If
    GetAllCategoryNames = arr
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Global.getAllCategoryNames", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

''Public Function NoOfCategoryNames() As Long
''    If sdCategoryDic Is Nothing Then NoOfCategoryNames = 0 Else NoOfCategoryNames = sdCategoryDic.Count + 1
''End Function

Public Function GetCGListing(ByVal ColNo As pe_CGListingColName, ByVal ACG As Boolean, Optional ByVal UniqueValues As Boolean = False, Optional CatNo As Long = 0 _
    , Optional ByVal CGno As Long = 0, Optional SCGno As Long = 0, Optional ByVal IncludeDescriptionAsOneValue = False) As Variant()
    ' Return an array with all the CG-SCG listings in an array. This array will be dimensioned differently dependent on the ACG or LCG setups.
    ' If arrays are created the isCGListActive boolean is set to True

    'IF LCG then CGNo / CGDesc / SCGNo / SCGDesc
    'IF ACG then CatNo / CatDesc/ CGNo / CGDesc / SCGNo / SCGDesc
    Dim P As cCBA_Prod
    Dim arr() As Variant
    Dim narr() As Variant
    Dim a As Long
    Dim v As Variant
'    Dim sdDic As Scripting.Dictionary
    On Error GoTo Err_Routine
    CBA_Error = ""

    ReDim arr(0 To 0, 0 To 0)
    arr(0, 0) = "False"
    If ACG = True Then
        If sdACGList Is Nothing Then
        Else
            If ColNo = eALL Then
                ReDim arr(0 To 5, 0 To sdACGList.Count - 1)
            Else
                ReDim arr(0 To sdACGList.Count - 1)
            End If
            a = -1
            For Each v In sdACGList.Keys
                Set P = sdACGList(v)
                If CatNo > 0 And Format(P.lACatNum, "00") <> CatNo Then GoTo NextACG
                If CGno > 0 And P.lACGNum <> CGno Then GoTo NextACG
                If SCGno > 0 And P.lASCGNum <> SCGno Then GoTo NextACG
                If UniqueValues = True And a > -1 Then
                    Select Case ColNo
                        Case eACatNum
                            If arr(a) = CStr(Format(P.lACatNum, "00")) Or arr(a) = CStr(Format(P.lACatNum, "00")) & "-" & CStr(P.sACat) Then GoTo NextACG
                        Case eACat
                            If arr(a) = CStr(P.sACat) Then GoTo NextACG
                        Case eACGnum
                            If arr(a) = CStr(P.lACGNum) Or arr(a) = CStr(P.lACGNum) & "-" & CStr(P.sACGDesc) Then GoTo NextACG
                        Case eACGdesc
                            If arr(a) = arr(a) = CStr(P.sACGDesc) Then GoTo NextACG
                        Case eASCGnum
                            If arr(a) = arr(a) = CStr(P.lASCGNum) Or arr(a) = CStr(P.lASCGNum) & "-" & CStr(P.sASCGDesc) Then GoTo NextACG
                        Case eASCGdesc
                            If arr(a) = arr(a) = CStr(P.sASCGDesc) Then GoTo NextACG
                    End Select
                End If
                a = a + 1
                Select Case ColNo
                    Case eALL
                        arr(0, a) = CStr(Format(P.lACatNum, "00"))
                        arr(1, a) = CStr(P.sACat)
                        arr(2, a) = CStr(P.lACGNum)
                        arr(3, a) = CStr(P.sACGDesc)
                        arr(4, a) = CStr(P.lASCGNum)
                        arr(5, a) = CStr(P.sASCGDesc)
                    Case eACatNum
                        If IncludeDescriptionAsOneValue = False Then arr(a) = CStr(Format(P.lACatNum, "00")) Else arr(a) = CStr(Format(P.lACatNum, "00")) & "-" & CStr(P.sACat)
                    Case eACat
                        arr(a) = CStr(P.sACat)
                    Case eACGnum
                        If IncludeDescriptionAsOneValue = False Then arr(a) = CStr(P.lACGNum) Else arr(a) = CStr(P.lACGNum) & "-" & CStr(P.sACGDesc)
                    Case eACGdesc
                        arr(a) = CStr(P.sACGDesc)
                    Case eASCGnum
                        If IncludeDescriptionAsOneValue = False Then arr(a) = CStr(P.lASCGNum) Else arr(a) = CStr(P.lASCGNum) & "-" & CStr(P.sASCGDesc)
                    Case eASCGdesc
                        arr(a) = CStr(P.sASCGDesc)
                End Select
NextACG:
            Next
        End If
    Else
        If sdLCGList Is Nothing Then
        Else
            If ColNo = eALL Then
                ReDim arr(0 To 3, 0 To sdLCGList.Count - 1)
            Else
                ReDim arr(0 To sdLCGList.Count - 1)
            End If
            a = -1
            For Each v In sdLCGList.Keys
                Set P = sdLCGList(v)
                If CGno > 0 And P.lLegCG <> CGno Then GoTo NextLCG
                If SCGno > 0 And P.lLegSCG <> SCGno Then GoTo NextLCG
                If UniqueValues = True And a > -1 Then
                    Select Case ColNo
                        Case eLegCG
                            If arr(a) = CStr(P.lLegCG) Or arr(a) = CStr(P.lLegCG) & "-" & CStr(P.sLegCGDesc) Then GoTo NextLCG
                        Case eLegCGdesc
                            If arr(a) = CStr(P.sLegCGDesc) Then GoTo NextLCG
                        Case eLegSCG
                            If arr(a) = CStr(P.lLegSCG) Or arr(a) = CStr(P.lLegSCG) & "-" & CStr(P.sLegSCGDesc) Then GoTo NextLCG
                        Case eLegSCGdesc
                            If arr(a) = CStr(P.sLegSCGDesc) Then GoTo NextLCG
                    End Select
                End If
                a = a + 1
                Select Case ColNo
                    Case eALL
                        arr(0, a) = CStr(P.lLegCG)
                        arr(1, a) = CStr(P.sLegCGDesc)
                        arr(2, a) = CStr(P.lLegSCG)
                        arr(3, a) = CStr(P.sLegSCGDesc)
                    Case eLegCG
                        If IncludeDescriptionAsOneValue = False Then arr(a) = CStr(P.lLegCG) Else arr(a) = CStr(P.lLegCG) & "-" & CStr(P.sLegCGDesc)
                    Case eLegCGdesc
                        arr(a) = CStr(P.sLegCGDesc)
                    Case eLegSCG
                        If IncludeDescriptionAsOneValue = False Then arr(a) = CStr(P.lLegSCG) Else arr(a) = CStr(P.lLegSCG) & "-" & CStr(P.sLegSCGDesc)
                    Case eLegSCGdesc
                        arr(a) = CStr(P.sLegSCGDesc)
                End Select
NextLCG:
            Next
        End If
    End If
    If ColNo <> eALL And a > -1 Then
        ReDim narr(0 To a)
        For a = 0 To a
            narr(a) = arr(a)
        Next
        arr = narr
    End If
    GetCGListing = arr
Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Global.getCGListing", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function AmendCategoryDic(ByVal DeleteNotAdd As Boolean, ByVal CategoryName As String, Optional ByRef Prd As cCBA_Prod) As Boolean
    ' Takes an amended collection item (consisting of cCBA_Prod objects) and places it into the sdCategoryDic
    Dim P As cCBA_Prod
    Dim col As Collection
    Dim idx As Long
    On Error GoTo Err_Routine
    CBA_Error = ""
        If sdCategoryDic.Exists(CategoryName) Then
            'Set col = sdCategoryDic(CategoryName)
            idx = 0
            For Each P In sdCategoryDic(CategoryName)
                idx = idx + 1
                If P.lACatNum = Prd.lACatNum And P.lACGNum = Prd.lACGNum And P.lASCGNum = Prd.lASCGNum And P.lLegCG = Prd.lLegCG And P.lLegSCG = Prd.lLegSCG Then
                    If DeleteNotAdd = True Then sdCategoryDic(CategoryName).Remove idx: bCategoryDicIsChanged = True: AmendCategoryDic = True
                    Exit Function
                End If
            Next
            sdCategoryDic(CStr(CategoryName)).Add Prd
            If DeleteNotAdd = False Then AmendCategoryDic = True: bCategoryDicIsChanged = True
        Else
            Set col = New Collection
            col.Add Prd
            sdCategoryDic.Add CStr(CategoryName), col
            If DeleteNotAdd = False Then AmendCategoryDic = True: bCategoryDicIsChanged = True
        End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Global.amendCategoryDic", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function IsMSegDicActive() As Boolean
    ' Asks if the Dic already exisits, if no then generates the sdMSegDic by utilizing the Neilsen Database connection and pulling the full list of sdMSegs and their setups
    IsMSegDicActive = False
End Function

Public Function PullCategoryMSegs(ByVal CategoryName As String) As Scripting.Dictionary
    ' Pulls all the Msegs for the requested Category
    Dim N As cCBA_NielsenData
    Dim v As Variant
    Dim sdDic As Scripting.Dictionary
    On Error GoTo Err_Routine
    CBA_Error = ""
    Set sdDic = New Scripting.Dictionary
    For Each v In sdMSegDic
        Set N = sdMSegDic(v)
        If N.SelectedForCategory = CategoryName Then sdDic.Add CStr(N.MSegDescription), N
    Next
    Set PullCategoryMSegs = sdDic
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Global.PullCategoryMSegs", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Public Function GetMSeg(ByVal MSegName As String, ByVal eMSegType As e_MSegType) As cCBA_NielsenData
Dim v As Variant
Dim N As cCBA_NielsenData
    For Each v In sdMSegDic
        Set N = sdMSegDic(v)
        If ((N.IsManual = True And eMSegType = eManual1) Or (N.IsHomescan = True And eMSegType = eHomescan) _
            Or (N.isScanData = True And eMSegType = eScanData)) And N.MSegDescription = MSegName Then
                Set GetMSeg = N
                Exit Function
        End If
    Next
    Set GetMSeg = Nothing
End Function
Public Function AmendMSeg(ByVal AmendNotDelete As Boolean, ByRef ND As cCBA_NielsenData) As Boolean
    ' Takes an amended CBA_Neilsen object and places it into the sdMSegDic, or sends instruction to remove an object from the sdMSegDic.
    ' Saves to the Database on amendment (note, this is triggered when a userform is closed)
Dim v As Variant
Dim N As cCBA_NielsenData
    If ND Is Nothing Then AmendMSeg = False: Exit Function
    For Each v In sdMSegDic
        Set N = sdMSegDic(v)
        If N.isScanData = ND.isScanData And N.IsHomescan = ND.IsHomescan And N.IsManual = ND.IsManual And _
            N.Category_ID = ND.Category_ID And N.MSegDescription = ND.MSegDescription Then
                If AmendNotDelete = True Then
                    Set sdMSegDic(v) = ND: AmendMSeg = True: Exit Function
                Else
                    sdMSegDic.Remove (v): AmendMSeg = True: Exit Function
                End If

        End If
    Next
    sdMSegDic.Add ND.MSegDescription, ND: AmendMSeg = True
End Function

Public Function GetAllMSegNames(ByVal eMSegType As e_MSegType) As Variant
    ' This steps through the sdMSegDic, takes all the keys (MSegName(S)) relevant fo the MsegType and returns a one-dimentional Array
    Dim arr() As String
    Dim N As cCBA_NielsenData
    Dim P As cCBA_Prod
    Dim a As Long
    Dim v As Variant
    Dim col As Collection
    On Error GoTo Err_Routine
    CBA_Error = ""
    If sdMSegDic Is Nothing Then
        ReDim arr(0 To 0)
        arr(0) = "False"
    Else
        Set col = New Collection
        For Each v In sdMSegDic
            Set N = sdMSegDic(v)
            If (N.IsHomescan And eMSegType = eHomescan) Or (N.isScanData And eMSegType = eScanData) Or (N.IsManual And (eMSegType = eManual1 Or eMSegType = eManual2 Or eMSegType = eManual3)) Then
                col.Add CStr(N.MSegDescription)
            End If
        Next
        ReDim arr(0 To IIf(col.Count > 0, col.Count - 1, 0))

        a = -1
        For Each v In col
            a = a + 1: arr(a) = CStr(v)
        Next

        GetAllMSegNames = arr
    End If
Exit_Routine:

    On Error Resume Next

    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Global.getAllMSegNames", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Function CreateMSegDic(ByVal CompelToGenerate As Boolean) As Boolean
    ' Create Market Segments dictionaries
    On Error GoTo Err_Routine
    CBA_Error = ""
    If sdMSegDic Is Nothing Or CompelToGenerate = True Then Set sdMSegDic = mCBA_Nielsen.GetNielsenSegmentationData: CreateMSegDic = True
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Global.createMSegDic", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function CreateCGListing(ByVal CompelToGenerate As Boolean) As Boolean
    ' If sdACGList and sdLCGList are nothing then creates the sdACGList and sdLCGList arrays.
    Dim strSQL As String
    Dim RS As ADODB.Recordset
    Dim v As Variant
    Dim P As cCBA_Prod
    Dim CBISCN As ADODB.Connection
    On Error GoTo Err_Routine
    CBA_Error = ""

    If CompelToGenerate = True Or bIsCGListActive = False Then
        Set CBISCN = New ADODB.Connection
        With CBISCN
            .ConnectionTimeout = 50
            .CommandTimeout = 50
            .Open "Provider=SQLNCLI10;DATA SOURCE=" & NATIONAL_DB & "DBL01;;INTEGRATED SECURITY=sspi;"
        End With

        Set RS = New ADODB.Recordset
        strSQL = "select cg.cgno, cg.description as CGDesc, isnull(scgno,0), isnull(scg.description,cg.Description) as SCGDesc" & Chr(10)
        strSQL = strSQL & "from cbis599p.dbo.COMMODITYGROUP cg left join cbis599p.dbo.SUBCOMMODITYGROUP scg on scg.cgno = cg.cgno"
        RS.Open strSQL, CBISCN
        If RS.EOF Then CreateCGListing = False: Exit Function
        Set sdLCGList = New Scripting.Dictionary
        Do
            Set P = New cCBA_Prod
            P.Build 0, "GetsLetsOnly"
            P.lACatNum = 0
            P.lLegCG = CLng(RS(0))
            P.sLegCGDesc = CStr(RS(1))
            P.lLegSCG = CLng(RS(2))
            P.sLegSCGDesc = CStr(RS(3))
            sdLCGList.Add CStr(Format(CStr(RS(0)), "000") & Format(CStr(RS(2)), "000")), P
            RS.MoveNext
        Loop Until RS.EOF
        Set RS = New ADODB.Recordset
        strSQL = "select * from cbis599p.dbo.tf_ACGMap() order by CGno"
        RS.Open strSQL, CBISCN
        If RS.EOF Then CreateCGListing = False: Exit Function
        Set sdACGList = New Scripting.Dictionary
        Do
            Set P = New cCBA_Prod
            P.Build 0, "GetsLetsOnly"
            P.lLegCG = 0
            P.lLegSCG = 0
            P.lACatNum = CLng(RS(1))
            P.sACat = CStr(RS(2))
            P.lACGNum = CLng(RS(3))
            P.sACGDesc = CStr(RS(4))
            P.lASCGNum = CLng(RS(5))
            P.sASCGDesc = CStr(RS(6))
            sdACGList.Add CStr(Format(CStr(RS(1)), "000") & Format(CStr(RS(3)), "000") & Format(CStr(RS(5)), "000")), P
            RS.MoveNext
        Loop Until RS.EOF
        CBISCN.Close
        Set CBISCN = Nothing
        Set RS = Nothing
        CreateCGListing = True
        bIsCGListActive = True
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Global.createCGListing", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Function SetCategoryDic(ByVal CompelToGenerate As Boolean) As Boolean
    ' Sets the sdCategoryDic by utilizing the CATREV Database connection and pulling the current Category Listing.
    ' Then generates all the nested collections of cCBA_Prod objects.
    Dim RS As ADODB.Recordset
    Dim P As cCBA_Prod
    Dim curCat As String
    Dim col As Collection
    Dim CAMERACN As ADODB.Connection
    On Error GoTo Err_Routine
    CBA_Error = ""
    Set CAMERACN = New ADODB.Connection
    If CompelToGenerate = True Or sdCategoryDic Is Nothing Then
        With CAMERACN
            .ConnectionTimeout = 50
            .CommandTimeout = 50
            .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Cam")
        End With
        Set sdCategoryDic = New Scripting.Dictionary
        Set sdCategoryID = New Scripting.Dictionary
        Set RS = New ADODB.Recordset
        RS.Open "select * from qry_L0_CategoryDic", CAMERACN  ' Updated query in CAMERA_Ext
        Do Until RS.EOF
            Set P = New cCBA_Prod
            P.Build 0, "GetsLetsOnly"
            P.lACatNum = CLng(RS(1))
            P.lACGNum = CLng(RS(2))
            P.lASCGNum = CLng(RS(3))
            P.lLegCG = CLng(RS(4))
            P.lLegSCG = CLng(RS(5))
            If sdCategoryDic.Exists(CStr(RS(0))) Then
                sdCategoryDic(CStr(RS(0))).Add P
            Else
                Set col = New Collection
                col.Add P
                sdCategoryDic.Add CStr(RS(0)), col
            End If
            If sdCategoryID.Exists(CStr(RS(0))) Then
                sdCategoryID(CStr(RS(0))) = CLng(RS(6))
            Else
                sdCategoryID.Add CStr(RS(0)), CLng(RS(6))
            End If
            RS.MoveNext
        Loop
        Set RS = Nothing
        CAMERACN.Close
        Set CAMERACN = Nothing
    End If
    If sdCategoryDic Is Nothing Then SetCategoryDic = False Else SetCategoryDic = True
Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Global.setCategoryDic", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Function InterfaceCategoryDicChangesToDB() As Boolean
    ' Update any sdCategoryDic changes into the database
    Dim CN As ADODB.Connection
    Dim AddRS As ADODB.Recordset
    Dim Pc As cCBA_Prod, Pm As cCBA_Prod
    Dim vc As Variant, vcc As Variant, vm As Variant, vmc As Variant
    Dim bfound As Boolean, catfound As Boolean
    Dim strSQL As String
    Dim lCAT_ID As Long
    Dim MainDic As Scripting.Dictionary, TempDic As Scripting.Dictionary
    Dim CatCnt As Long, MainCnt As Long
    Dim CatCol As Scripting.Dictionary, MainCol As Scripting.Dictionary
    Dim arr As Variant
    Dim col As Collection

    'THIS IS CLOSE BUT ITS NOT GOING TO WORK
    ' HAVE TO CONSIDER IF CATEGORYDIC HAS THINGS NOT IN THE DB AND IF IT HAS STUFF THAT ISNT.
    ' SHOULD THE RS.EOF LOOP HAPPEN IN THE CATEGORYDIC LOOP INSTEAD??? HMMM
    On Error GoTo Err_Routine
    CBA_Error = ""
    If bCategoryDicIsChanged = True Then

        'Stop

        Set TempDic = sdCategoryDic
        SetCategoryDic True
        Set MainDic = sdCategoryDic
        Set sdCategoryDic = TempDic
        Set TempDic = Nothing

        'MainDic is how the Data is currently in the Database
        'sdCategoryDic is the current updated and active Dictionary

        For Each vc In sdCategoryDic
            bfound = False
            For Each vm In MainDic
                If vc = vm Then bfound = True: Exit For
            Next
            If bfound = False Then
                Set col = New Collection
                MainDic.Add vc, col
            End If
        Next

        For Each vc In sdCategoryDic
            Set CatCol = New Scripting.Dictionary
            Set MainCol = New Scripting.Dictionary
            catfound = False
            For Each vm In MainDic
                If vc = vm Then
                    catfound = True
                    CatCnt = 0
                    For Each vcc In sdCategoryDic(vc)
                        CatCnt = CatCnt + 1
                        bfound = False
                        Set Pc = vcc
                        MainCnt = 0
                        For Each vmc In MainDic(vm)
                            MainCnt = MainCnt + 1
                            Set Pm = vmc
                            'Debug.Print Pm.lLegCG & Pm.lLegSCG, Pc.lLegCG & Pc.lLegSCG
                            If Pc.lACatNum = Pm.lACatNum And Pc.lACGNum = Pm.lACGNum And Pc.lASCGNum = Pm.lASCGNum And Pc.lLegCG = Pm.lLegCG And Pc.lLegSCG = Pm.lLegSCG Then
                                bfound = True
                                Exit For
                            End If
                        Next
                        If bfound = False Then
                            CatCol.Add CStr(vc) & "|" & CStr(CatCnt), CLng(CatCnt)
                        End If
                    Next

                    MainCnt = 0
                    For Each vmc In MainDic(vm)
                        MainCnt = MainCnt + 1
                        bfound = False
                        Set Pm = vmc
                        CatCnt = 0
                        For Each vcc In sdCategoryDic(vc)
                            CatCnt = CatCnt + 1
                            Set Pc = vcc
                            'Debug.Print Pm.lLegCG & Pm.lLegSCG, Pc.lLegCG & Pc.lLegSCG
                            If Pc.lACatNum = Pm.lACatNum And Pc.lACGNum = Pm.lACGNum And Pc.lASCGNum = Pm.lASCGNum And Pc.lLegCG = Pm.lLegCG And Pc.lLegSCG = Pm.lLegSCG Then
                                bfound = True
                                Exit For
                            End If
                        Next
                        If bfound = False Then
                            MainCol.Add CStr(vm) & "|" & CStr(MainCnt), CLng(MainCnt)
                        End If
                    Next
                End If
            Next

            If CatCol.Count > 0 Or MainCol.Count > 0 Then
                Set CN = New ADODB.Connection: CN.ConnectionTimeout = 50: CN.CommandTimeout = 50
                CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Cam")
            End If
            For Each vcc In CatCol
                'needs to be added
                arr = Split(vcc, "|")
Rpt1:
                Set AddRS = New ADODB.Recordset
                strSQL = "select CN_ID, CN_isActive from L0_CategoryName where CN_CategoryName = '" & Trim(arr(0)) & "'"
                AddRS.Open strSQL, CN
                If AddRS.EOF = False Then
                    lCAT_ID = AddRS(0)
                    If AddRS(1) = False Then
                        Set AddRS = New ADODB.Recordset
                        strSQL = "Update L0_CategoryName" & Chr(10) & "set CN_isActive = True where CN_ID = " & lCAT_ID
                        AddRS.Open strSQL, CN
                        GoTo Rpt1
                    End If
                Else
                    Set AddRS = New ADODB.Recordset
                    strSQL = "Insert into L0_CategoryName(CN_CategoryName,CN_isActive)" & Chr(10)
                    strSQL = strSQL & "VALUES('" & Trim(arr(0)) & "', True)"
                    AddRS.Open strSQL, CN
                    GoTo Rpt1
                End If
                Set AddRS = New ADODB.Recordset
                Set Pc = sdCategoryDic(arr(0))(CatCol(vcc))
                strSQL = "Insert into L0_CategoryAllocations"
                If Pc.lACatNum > 0 Then
                    strSQL = strSQL & "(CA_CAT_ID,CA_ACG,CA_ACGCat,CA_ACGCGNo,CA_ACGSCGNo)" & Chr(10)
                    strSQL = strSQL & "VALUES (" & lCAT_ID & ",True," & Pc.lACatNum & ", " & Pc.lACGNum & ", " & Pc.lASCGNum & ")"
                Else
                    strSQL = strSQL & "(CA_CAT_ID,CA_ACG,CA_LCGCGNo,CA_LCGSCGNo)" & Chr(10)
                    strSQL = strSQL & "VALUES (" & lCAT_ID & ",False," & Pc.lLegCG & ", " & Pc.lLegSCG & ")"
                End If
                AddRS.Open strSQL, CN
                lCAT_ID = 0
            Next
            For Each vmc In MainCol
                'needs to be deleted
                If IsEmpty(vmc) = False Then
                    arr = Split(vmc, "|")
                    Set AddRS = New ADODB.Recordset
                    strSQL = "select CN_ID, CN_isActive from L0_CategoryName where CN_CategoryName = '" & CStr(Trim(arr(0))) & "'"
                    AddRS.Open strSQL, CN
                    If AddRS.EOF = False Then
                        lCAT_ID = AddRS(0)
                    Else
                        'Dont want add here really and the inactive is handled later
                        Exit For
                    End If
                    Set AddRS = New ADODB.Recordset
                    Set Pm = MainDic(arr(0))(MainCol(vmc))
                    strSQL = "Delete from L0_CategoryAllocations where"
                    If Pm.lACatNum > 0 Then
                        strSQL = strSQL & " CA_CAT_ID = " & lCAT_ID & " and CA_ACGCat = " & Pm.lACatNum & " And CA_ACGCGNo = " & Pm.lACGNum & " and CA_ACGSCGNo = " & Pm.lASCGNum & Chr(10)
                    Else
                        strSQL = strSQL & " CA_CAT_ID = " & lCAT_ID & " and CA_LCGCGNo = " & Pm.lLegCG & " And CA_LCGSCGNo = " & Pm.lLegSCG & Chr(10)
                    End If
                    AddRS.Open strSQL, CN
                End If
                lCAT_ID = 0
            Next
        Next

        For Each vm In MainDic
            bfound = False
            For Each vc In sdCategoryDic
                If vm = vc And sdCategoryDic(vc).Count > 0 Then bfound = True: Exit For
            Next
            If bfound = False Then
                Set AddRS = New ADODB.Recordset
                If CN.State = 0 Then
                    Set CN = New ADODB.Connection: CN.ConnectionTimeout = 50: CN.CommandTimeout = 50
                    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Cam")
                End If
                strSQL = "update L0_CategoryName" & Chr(10) & "set CN_isActive = FALSE where CN_CategoryName = '" & vm & "'"
                AddRS.Open strSQL, CN
            End If
        Next
        bCategoryDicIsChanged = False
    End If
Exit_Routine:

    On Error Resume Next
    CN.Close
    Set CN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCAM_Global.InterfaceCategoryDicChangesToDB", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function Get_eSysDocTypeCols(Optional ByVal eDocType As e_DocuType = 0, Optional ByVal DocMetric As String = "SH_Desc", Optional ByVal DocString As String = "") As Variant
    ' Will get the 'System Document Type' data from the appropriate UT_DB Hdr table
    Dim sReturn As String, sSQL As String, RS As ADODB.Recordset, lIdx As Long, bfound As Boolean
    Dim CN As ADODB.Connection, arr()
    Static arrDB(), bActive As Boolean

    On Error GoTo Err_Routine
    If bActive = False Then
        If CN Is Nothing Then
            Set CN = New ADODB.Connection
            CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("UDT", , , , , "Test") & ";"
        End If
        Set RS = New ADODB.Recordset
        sSQL = "SELECT SH_ID, SH_Desc, SH_BuildCode FROM C1_Seg_Template_Hdrs WHERE SH_SysType ='C' and SH_Valid='Y' ORDER BY SH_Desc;"
''        sSQL = "SELECT SH_ID, SH_Desc, SH_BuildCode FROM C1_Seg_Template_Hdrs WHERE SH_SysType ='C' ORDER BY SH_ID;"
        RS.Open sSQL, CN
        arrDB = RS.GetRows
        bActive = True
    End If
    ReDim arr(0)
    For lIdx = 0 To UBound(arrDB, 2)
        If DocMetric = "List" Then
            ReDim Preserve arr(lIdx)
            arr(lIdx) = arrDB(1, lIdx)
        Else
            ' Find the field required
            If UCase(DocMetric) = "SH_ID" Then
                If DocString = arrDB(1, lIdx) Then
                    bfound = True
                    Exit For
                End If
            Else
                If eDocType = arrDB(0, lIdx) Then
                    bfound = True
                    Exit For
                End If
            End If
        End If
    Next
    If DocMetric = "List" Then
        Get_eSysDocTypeCols = arr
        GoTo Exit_Routine
    End If
    If bfound = False Then
        MsgBox eDocType & " Document Type not found"
        Stop
    End If
    ' Find the field required
    Select Case UCase(DocMetric)
        Case Is = "SH_ID"
            Get_eSysDocTypeCols = arrDB(0, lIdx)
        Case Is = "SH_DESC"
            Get_eSysDocTypeCols = arrDB(1, lIdx)
        Case Is = "BUILDCODE"
            Get_eSysDocTypeCols = NZ(arrDB(2, lIdx), "")
        Case Else
            MsgBox DocMetric & " not found"
            Stop
    End Select

Exit_Routine:
    On Error Resume Next
    Set RS = Nothing
    Set CN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_UDT_ETL.Get_eSysDocTypeCols", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Sub Class_Initialize()
    'needs to contruct the ribbon data object - should hold only global list values not category specific
    On Error GoTo Err_Routine
    CBA_Error = ""
    lCurCatIndexSelected = -1
    byACGSelected = 99
    If SetCategoryDic(True) And CreateCGListing(True) And CreateMSegDic(True) Then bIsActive = True
Exit_Routine:

    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-cCAM_Global.Class_Initialize", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Private Property Get sdCategoryDic() As Scripting.Dictionary: Set sdCategoryDic = psdCategory: End Property
Private Property Set sdCategoryDic(ByVal objNewValue As Scripting.Dictionary): Set psdCategory = objNewValue: End Property
Public Property Get sdLCGList() As Scripting.Dictionary: Set sdLCGList = psdLCGList: End Property
Private Property Set sdLCGList(ByVal objNewValue As Scripting.Dictionary): Set psdLCGList = objNewValue: End Property
Public Property Get sdACGList() As Scripting.Dictionary: Set sdACGList = psdACGList: End Property
Private Property Set sdACGList(ByVal objNewValue As Scripting.Dictionary): Set psdACGList = objNewValue: End Property
Private Property Get bIsCGListActive() As Boolean: bIsCGListActive = pbIsCGListActive: End Property
Private Property Let bIsCGListActive(ByVal bNewValue As Boolean): pbIsCGListActive = bNewValue: End Property
Private Property Get sdMSegDic() As Scripting.Dictionary: Set sdMSegDic = psdMSegDic: End Property
Private Property Set sdMSegDic(ByVal objNewValue As Scripting.Dictionary): Set psdMSegDic = objNewValue: End Property
Private Property Get sActiveRibbonCategory() As String: sActiveRibbonCategory = psActiveRibbonCategory: End Property
Private Property Let sActiveRibbonCategory(ByVal sNewValue As String): psActiveRibbonCategory = sNewValue: End Property
Private Property Get dtDateTo() As Date: dtDateTo = pdtDateTo: End Property
Private Property Let dtDateTo(ByVal dtNewValue As Date): pdtDateTo = dtNewValue: End Property
Public Property Get bIsActive() As Boolean: bIsActive = pbIsActive: End Property
Private Property Let bIsActive(ByVal bNewValue As Boolean): pbIsActive = bNewValue: End Property
Public Property Get bCategoryDicIsChanged() As Boolean: bCategoryDicIsChanged = pbCategoryDicIsChanged: End Property
Private Property Let bCategoryDicIsChanged(ByVal bNewValue As Boolean): pbCategoryDicIsChanged = bNewValue: End Property
Public Property Get lCurCatIndexSelected() As Long: lCurCatIndexSelected = plCurCatIndexSelected: End Property
Public Property Let lCurCatIndexSelected(ByVal lNewValue As Long): plCurCatIndexSelected = lNewValue: End Property
Public Property Get byACGSelected() As Byte: byACGSelected = pbyACGSelected: End Property
Public Property Let byACGSelected(ByVal NewValue As Byte): pbyACGSelected = NewValue: End Property
Public Property Get sdCategoryID() As Scripting.Dictionary: Set sdCategoryID = psdCategoryID: End Property
Private Property Set sdCategoryID(ByVal NewValue As Scripting.Dictionary): Set psdCategoryID = NewValue: End Property

