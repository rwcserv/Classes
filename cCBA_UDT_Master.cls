VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCBA_UDT_Master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit                                             ' cCBA_UDT_Master Changed 200103

Private psdSH As Scripting.Dictionary                       ' Template Hdrs from C1_Seg_Template Hdrs - e.g NDP, Retender and MSO Core Range
Private psdSL As Scripting.Dictionary                       ' Template Lines from C2_Seg_Template_Lines
Private psdUDT As Scripting.Dictionary                      ' Template Lines from C3_Seg_Segment_UDTs
Private plCurrUDTIdx As Long                                ' psdUDT current Index
    ' This module will contain all the Master UDT data for the System Document Type (at this time C or T)
    ' Using the GET_UDT_SD routine, it will return the scripting dictionary for the individual e_DocuType required

Public Function Generate_UDT(ByVal DocType As e_DocuType) As Boolean
    ' Will Generate the UDT Master Class SDs
    Set sdSH = New Scripting.Dictionary
    Set sdSL = New Scripting.Dictionary
    Set sdUDT = New Scripting.Dictionary
    Generate_UDT = Get_UDT_DB(DocType)
End Function

Public Function Get_UDT_SD(ByVal DocType As e_DocuType) As Scripting.Dictionary
    ' Will return the Master Class Scripting.Dictionary for the e_DocuType (Doc Type) required
    Dim Init_Next As Long, sdUT As Scripting.Dictionary, cUDT As cCBA_UDT, sUT_ID As String
    On Error GoTo Err_Routine
    CBA_Error = ""
    Set sdUT = New Scripting.Dictionary
    Call Get_UDT_Idx(DocType, -1)
    Init_Next = 1
    Do While Init_Next = 1
        Set cUDT = Get_UDT_Idx(DocType, Init_Next)
        If Not cUDT Is Nothing Then
            sUT_ID = cUDT.sUT_ID
            sdUT.Add sUT_ID, cUDT
        End If
    Loop
Exit_Routine:

    On Error Resume Next
    If Not sdUT Is Nothing Then
        Set Get_UDT_SD = sdUT
    Else
        Set Get_UDT_SD = Nothing
    End If
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_UDT_Master.Get_UDT_SD", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function Get_UDT_Idx(ByVal DocType As e_DocuType, Optional ByRef Init_Next As Long = 1, Optional ByVal SysType As String = "T") As cCBA_UDT
    ' This routine will get the next index for the DocType selected, and return its cCBA_UDT
    Static stcLastLT As String
    Dim lSeg As Long, lSegIdx As Long, sSL_Mandatory As String, sSL_SH_ID As String, cSL As cTEN4_Segments, cUDT As cCBA_UDT
    On Error GoTo Err_Routine
    CBA_Error = ""

    Set Get_UDT_Idx = Nothing
    ' If initialising...
    If Init_Next < 0 Then
        stcLastLT = SysType
        lCurrUDTIdx = -1
''        Init_Next = 0
        GoTo Exit_Routine
    End If
    ' Ensure input parameters haven't changed
    If stcLastLT <> SysType Then
        MsgBox SysType & " Load Type is different to initiated"
        Init_Next = -1
        Exit Function
    End If
ReStartUT:
    ' Get next position relative to the last
    If lCurrUDTIdx + Init_Next < sdUDT.Count Then                       ' Until End point of SD
        lCurrUDTIdx = lCurrUDTIdx + Init_Next                           ' Increment the current ID
        ' The following 'Clone' Code won't now be done here, but done in the ETL itself to cater for multiple Group_No lines
''        Set cUDT = New cCBA_UDT
''        If cUDT.Clone(sdUDT(sdUDT.Keys(lCurrUDTIdx))) = False Then
''            MsgBox "UST Failed to Clone"
''            Set Get_UDT_Idx = Nothing
''            Exit Function
''        End If
        ' Instead the following code which will just ref the master class of UDTs
        Set cUDT = sdUDT(sdUDT.Keys(lCurrUDTIdx))

        lSeg = cUDT.lUT_SL_ID                                           ' Get the Segment ID from the UDT Class array
        Set cSL = sdSL(CStr(lSeg))
        sSL_Mandatory = cSL.sSL_Mandatory                               ' Get the Mandatory flag
        If sSL_Mandatory = "Y" Or sSL_Mandatory = "U" Then
            sSL_SH_ID = cSL.sSL_SH_ID                                   ' See if the segment is valid for this Template
            If InStr(1, sSL_SH_ID, "," & DocType & ",") = 0 Then sSL_Mandatory = "N" ' Not valid
        End If
        ' Is the UDT valid to include?
        If sSL_Mandatory = "Y" Or sSL_Mandatory = "U" Then              ' sSL_Mandatory = "Y" is a Mandatory Segment - sSL_Mandatory = "U"  is a User Included Segment
            Set Get_UDT_Idx = cUDT
        Else
            Init_Next = 1
            GoTo ReStartUT
        End If
    Else    ' Else is last one
        Init_Next = -1
    End If
    
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_UDT_Master.Get_UDT_Idx", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function Get_UDT_DB(ByVal DocType As e_DocuType) As Boolean
    ' This routine will get the Template Headers, Segment Lines, UDT Lines; SysType is the System Doc Type - Currently C[ategory] or T[en]
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim sSQL As String, sSH_ID As e_DocuType, lGrp_No As Long, lSavedGrp_No As Long, lGrpIdx As Long, lIdxIdx As Long
    Dim SH_Desc As String, UT_Default_Value As String, sUT_ID As String, sUT_ID_Ext As String
    Dim SL_ID As Long, UT_Link_ID As Long, SL_Seq As Long, SL_Desc As String, SL_SH_ID As String, SL_Mandatory As String, SL_Valid As String, SL_UpdDate As String, SL_UpdUser As String
    Dim UT_Pos_Top As Long, UT_Pos_Left As Long, UT_Merge_Width As Long, lTotal As Long, bValidLine As Boolean
    Dim cUDT As cCBA_UDT, cSH As cTEN3_Templates, cSL As cTEN4_Segments, lSH_Idx As Long, lSL_Idx As Long, lUT_Idx As Long, SysType As String
    
    On Error GoTo Err_Routine
    Get_UDT_DB = False
    CBA_ErrTag = "SQL"
    Set CN = New ADODB.Connection
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("UDT", , , , , "Test") & ";"    ' At this stage Test or Live should be established for whatever application is being used
    ' Get the OA Doc Type
    SysType = Get_SysType(DocType, , CN)
    ' Get the correct database
''    If SysType = "T" Then
''        CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("UDT", , , , , "Ten") & ";"
''    ElseIf SysType = "C" Then
''        CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("UDT", , , , , "Cam") & ";"
''    Else
''        MsgBox SysType & " database key not found"
''        Exit Function
''    End If
    ' Get the Template Hdrs
    Set RS = New ADODB.Recordset
    sSQL = "SELECT * FROM C1_Seg_Template_Hdrs WHERE SH_SysType='" & SysType & "' ORDER BY SH_ID;"
    Set RS = New ADODB.Recordset
    RS.Open sSQL, CN
    Do While Not RS.EOF
        CBA_ErrTag = ""
        sSH_ID = CStr(RS!SH_ID)
        SH_Desc = NZ(RS!SH_Desc, 0)
        ' Add a Header
        lSH_Idx = lSH_Idx + 1
        If sdSH.Exists(CStr(sSH_ID)) Then
            MsgBox sSH_ID & " Tender Hdr already exists"
            Exit Function
        End If
        Set cSH = New cTEN3_Templates
        Call cSH.Add_Class_SH(lSH_Idx, sSH_ID, SH_Desc)
        sdSH.Add sSH_ID, cSH
        RS.MoveNext
    Loop
'''    CN.Close

   
    ' Get the Segment records and insert into the class module arrays / dictionaries (note: brings in SL_SH_ID=,#, (Template ID) records too. i.e. records that are for all templates)
    Set RS = New ADODB.Recordset
    sSQL = "SELECT * FROM qry_C1_Tenders WHERE SL_Sts_ID=1 AND UT_Sts_ID=1 AND SL_SysType='" & SysType & "' ORDER BY SL_Seq, UT_Pos_Top, UT_Pos_Left"
    RS.Open sSQL, CN
    Do While Not RS.EOF
        Get_UDT_DB = True
        CBA_ErrTag = ""
        If SL_ID <> NZ(RS!SL_ID, 0) Then
            UT_Pos_Top = 0
            SL_ID = NZ(RS!SL_ID, 0): SL_SH_ID = NZ(RS!SL_SH_ID, "")
            SL_Seq = NZ(RS!SL_Seq, 0): SL_Desc = NZ(RS!SL_Desc, ""):
            SL_Mandatory = NZ(RS!SL_Mandatory, "N"): SL_Valid = NZ(RS!SL_Valid, "Y")
            ' Add a Segment Class Line
            If sdSL.Exists(CStr(SL_ID)) Then
                MsgBox SL_ID & " Tender Line already exists"
                Exit Function
            End If
            lSL_Idx = lSL_Idx + 1
            Set cSL = New cTEN4_Segments
            Call cSL.Add_Class_SL(lSL_Idx, SL_ID, SL_Seq, SL_Desc, SL_SH_ID, SL_Mandatory, 0, SL_Valid)
            sdSL.Add CStr(SL_ID), cSL
        End If
        
''        ' Add the row to the udt array
        sUT_ID = CStr(RS!UT_ID)
        sUT_ID_Ext = g_Fmt_2_IDs(sUT_ID, 0, e_UTFldFmt.eUT_ID)
        If RS!UT_Link_ID > 0 Then UT_Link_ID = g_Fmt_2_IDs(RS!UT_Link_ID, lGrp_No, e_UTFldFmt.eUT_Link) Else UT_Link_ID = "0"
        ' Set up the ID and it's array index to the dictionary
        lUT_Idx = lUT_Idx + 1
        If sdUDT.Exists(sUT_ID) Then
            MsgBox sUT_ID & " UDT already exists"
            Exit Function
        End If
        ' Add the ID and it's array index to the master dictionary
        Set cUDT = New cCBA_UDT
        UT_Default_Value = NZ(RS!UT_Default_Value, "")
        Call cUDT.Add_Class_UT(lUT_Idx, sUT_ID, sUT_ID_Ext, RS!UT_SL_ID, RS!UT_Pos_Left, RS!UT_Pos_Top, RS!UT_Grp_No, RS!UT_Merge_Width, RS!UT_Merge_Height, UT_Link_ID, NZ(RS!UT_Name, ""), _
                            RS!UT_Hdg_Type, UT_Default_Value, RS!UT_Wrap_Text, RS!UT_Font_Size, RS!UT_Locked, RS!UT_BG_Color, RS!UT_FG_Color, RS!UT_Bold, RS!UT_Italic, RS!UT_Underline, RS!UT_StrikeThrough, RS!UT_TextAlign, _
                            RS!UT_Field_Type, NZ(RS!UT_Formula, ""), NZ(RS!UT_Procedure, ""), NZ(RS!UT_Line_Type, ""), NZ(RS!UT_Border, ""), NZ(RS!UT_Cond_Format, ""), NZ(RS!UT_Hyperlink, ""), _
                            NZ(RS!UT_Image, ""), NZ(RS!UT_Source, ""), "000000", RS!UT_Grp_No)
        sdUDT.Add sUT_ID, cUDT
        ' Next record
        RS.MoveNext
    Loop
        
Exit_Routine:
    On Error Resume Next
''    Call SaveIP_UT("Set", False)
    CN.Close
    Set RS = Nothing
    Set CN = Nothing
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("cCBA_UDT_Master-f-Get_UDT_DB", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function Get_SysType(ByVal Doc_ID As Long, Optional SysType As String = "SysType", Optional CN As ADODB.Connection) As String
    ' Will get the 'System Document Type' (at the moment C ot T) or the 'Document Type Desc' from the appropriate UT_DB table
    Dim sReturn As String, sSQL As String, RS As ADODB.Recordset, lIdx As Long
    Static arrDB(), bActive As Boolean
    
    On Error GoTo Err_Routine
    If bActive = False Then
        If CN Is Nothing Then
            Set CN = New ADODB.Connection
            CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("UDT", , , , , "Test") & ";"
        End If
        Set RS = New ADODB.Recordset
        sSQL = "SELECT SH_ID,SH_SysType, SH_Desc FROM C1_Seg_Template_Hdrs"
        RS.Open sSQL, CN
        arrDB = RS.GetRows
        bActive = True
    End If
    ' Run through the Doc Types to find the type required
    For lIdx = 0 To UBound(arrDB, 2)
        If Doc_ID = arrDB(0, lIdx) Then
            If SysType = "SysType" Then
                Get_SysType = arrDB(1, lIdx)
            Else
                Get_SysType = arrDB(2, lIdx)
            End If
            Exit For
        End If
    Next
    
Exit_Routine:
    On Error Resume Next
    Set RS = Nothing
    ''Set CN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-Get_SysType", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function

Private Property Get sdSH() As Scripting.Dictionary: Set sdSH = psdSH: End Property
Private Property Set sdSH(ByVal objNewValue As Scripting.Dictionary): Set psdSH = objNewValue: End Property

Private Property Get sdSL() As Scripting.Dictionary: Set sdSL = psdSL: End Property
Private Property Set sdSL(ByVal objNewValue As Scripting.Dictionary): Set psdSL = objNewValue: End Property

Private Property Get lCurrUDTIdx() As Long: lCurrUDTIdx = plCurrUDTIdx: End Property
Private Property Let lCurrUDTIdx(ByVal lNewValue As Long): plCurrUDTIdx = lNewValue: End Property

Private Property Get sdUDT() As Scripting.Dictionary: Set sdUDT = psdUDT: End Property
Private Property Set sdUDT(ByVal objNewValue As Scripting.Dictionary): Set psdUDT = objNewValue: End Property

