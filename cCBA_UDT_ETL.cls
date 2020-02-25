VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCBA_UDT_ETL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit                                             ' cCBA_UDT_ETL Changed 200103

Private psdUDT As Scripting.Dictionary                      ' Template Lines from cCBA_UDT And cCBA_UDT_Master
Private plCurrUDTIdx As Long                                ' psdUDT current Index
Private psSysType As String                                 ' System Template Type
'Private peDocType As e_DocuType                            ' Template Required
'Private plDoc_ID As Long                                   ' Document Header ID
'Private psDoc_Name As String                               ' Document Description
Private psngCol_Width As Single
Private psngRow_Height As Single
Private psngNo_Of_Cols As Single
Private psngPage_Margin As Single
''Private Const cnst0_ID As Long = 200                        ' If the Doc_ID is less than this figure, then it is treated as a NEW ID i.e. it is added; else it is updated
' Change the below as required
''Private Const cstFIELDTYPE As String = "txt", cstLINETYPE As String = "T1R1B1L1"

Public Function Generate_UDT_ETL(ByVal DocType As e_DocuType, Optional ByVal lDoc_ID As Long = -1, Optional sdSubMst As Scripting.Dictionary) As Boolean
''Public Function Generate_UDT_ETL(ByVal DocType As e_DocuType, Optional ByVal DocName As String, Optional ByVal DocID As Long = -1, Optional cMast As cCBA_UDT_Master) As Boolean
    ' Will Generate the UDT Document Class SD
    Dim CN As ADODB.Connection
    'eDocType = DocType ': lDoc_ID = DocID ': sDoc_Name = DocName
    sSysType = Get_SysType(DocType)
    COL_WIDTH = Get_SysType(DocType, "Col_Width")
    ROW_HEIGHT = Get_SysType(DocType, "Row_Height")
    NO_OF_COLS = Get_SysType(DocType, "No_Of_Cols")
    PAGE_MARGIN = Get_SysType(DocType, "Page_Margin")
    Set sdUDT = New Scripting.Dictionary
    If lDoc_ID < CBA_200_ID Then
        Generate_UDT_ETL = Copy_SubMst_UDT(sdSubMst)
    Else
        Generate_UDT_ETL = Get_UDT_DB(lDoc_ID)
    End If
End Function
Public Function MassEdit(ByVal TopEditNotLeft As Boolean, ByVal StartVal As Long, AddVal As Long, Optional ByVal EndVal As Long = 0) As Boolean '@TP
    Dim v As Variant, v2 As Variant, v3 As Variant
    Dim U As cCBA_UDT, U2(1 To 9999) As cCBA_UDT                '@RWCam @TP array will need fixing...
    Dim newRowDic As Scripting.Dictionary, sdNewUDT As Scripting.Dictionary
    Dim lMaxMergeHeight As Long, a As Long, b As Long
    On Error GoTo Err_Routine
    CBA_Error = ""
    Set sdNewUDT = New Scripting.Dictionary
    Set newRowDic = New Scripting.Dictionary
    If TopEditNotLeft = True Then
        For Each v In sdUDT
            Set U = sdUDT(v)
            If U.lUT_Pos_Top = StartVal And U.lUT_Pos_Left = 1 Then
                For Each v2 In sdUDT
                    If sdUDT(v2).lUT_Pos_Top = StartVal Then
                        a = a + 1: Set U2(a) = New cCBA_UDT: U2(a).CloneETL sdUDT(v2)
                        newRowDic.Add CStr(Format(sdUDT(v2).lUT_Pos_Top, "00000") & Format(sdUDT(v2).lUT_Pos_Left, "00000")), U2(a)
                        If sdUDT(v2).lUT_Merge_Height > lMaxMergeHeight Then lMaxMergeHeight = sdUDT(v2).lUT_Merge_Height
                    ElseIf sdUDT(v2).lUT_Pos_Top > StartVal Then
                        Exit For
                    End If
                Next
                For v2 = 1 To AddVal
                    For Each v3 In newRowDic
                        a = a + 1: Set U2(a) = New cCBA_UDT: U2(a).CloneETL newRowDic(v3)
                         U2(a).lUT_Pos_Top = U2(a).lUT_Pos_Top + (U2(a).lUT_Merge_Height * v2)
                         U2(a).lUT_Grp_No = U2(a).lUT_Grp_No + v2
                        sdNewUDT.Add CStr(Format(U2(a).lUT_Pos_Top, "00000") & Format(U2(a).lUT_Pos_Left, "00000")), U2(a)
                    Next
                Next
            End If
            If U.lUT_Pos_Top <= StartVal Then
                a = a + 1: Set U2(a) = New cCBA_UDT: U2(a).CloneETL U
                sdNewUDT.Add CStr(v), U2(a)
            ElseIf U.lUT_Pos_Top > StartVal Then
                U.lUT_Pos_Top = U.lUT_Pos_Top + (lMaxMergeHeight * (AddVal))
                a = a + 1: Set U2(a) = New cCBA_UDT: U2(a).CloneETL U
                sdNewUDT.Add CStr(Format(U.lUT_Pos_Top, "00000") & Format(U.lUT_Pos_Left, "00000")), U2(a)
            End If
        Next
        Set sdUDT = sdNewUDT
    Else
        Debug.Print "This function is not built yet"
        Stop
    '    For Each v In sdUDT
    '        Set U = sdUDT(v)
    '        If U.lUT_Pos_Left >= StartVal And IIf(EndVal = 0, False, U.lUT_Pos_Left < EndVal) Then U.lUT_Pos_Top = U.lUT_Pos_Top + AddVal
    '    Next
    End If
    If Err.Number = 0 Then MassEdit = True
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_UT_ETL.MassEdit", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

''
''Private Function Get_UDT_Idx(Optional ByVal Init_Next As Long = 1) As cCBA_UDT
''    ' This routine will get and return the next index for the document
''    Dim cUDT As cCBA_UDT
''    On Error GoTo Err_Routine
''    CBA_Error = ""
''
''    Set Get_UDT_Idx = Nothing
''    ' If initialising...
''    If Init_Next < 0 Then
''        lCurrUDTIdx = 0
''        Init_Next = 0
''    End If
''    ' Get next position relative to the last
''    If lCurrUDTIdx + Init_Next <= sdUDT.Count Then                        ' End point of for this SD
''        lCurrUDTIdx = lCurrUDTIdx + Init_Next                             ' Increment the current ID
''        Set cUDT = sdUDT.Keys(lCurrUDTIdx)
''    End If
''Exit_Routine:
''
''    On Error Resume Next
''    Exit Function
''
''Err_Routine:
''    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_UDT_ETL.Get_UDT_Idx", 3)
''    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
''    Debug.Print CBA_Error
''    Stop            ' ^RW Camera + next line
''    Resume Next
''    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
''    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
''    GoTo Exit_Routine
''    Resume Next
''End Function

Private Function Copy_SubMst_UDT(sdMst As Scripting.Dictionary) As Boolean
    ' This routine will reformat the UDT Lines from the copied sub-master UDT class
'    Dim UT_Pos_Top As Long, UT_Pos_Left As Long, UT_Merge_Width As Long, lTotal As Long, bValidLine As Boolean
    Dim cUDT As cCBA_UDT, lDL_Idx As Long, sDL_ID As String, Init_Next As Long, lMHeight As Long ''', DL_Data As String, lTop As Long, lHeight As Long
    'dim cMst As cCBA_UDT_SubMst
    ''Dim v As Variant, v2 As Variant
    Dim bNextSeg As Boolean, lAccTop As Long, lmaxTop As Long, lCurSL_ID As Long, lVNo As Long, lRecNo As Long, lTotGrpNo As Long, lGrpNo As Long, lUDTIdx As Long
    On Error GoTo Err_Routine
    Copy_SubMst_UDT = False: lDL_Idx = 0: lVNo = 0: Init_Next = 1: lUDTIdx = -1: lTotGrpNo = 0: lGrpNo = 0: lMHeight = 0 'lTop = 1:
    ' Get next position relative to the last
    Do While lUDTIdx + Init_Next < sdMst.Count                      ' End point of Copied Sub-Master UDT
        lUDTIdx = lUDTIdx + Init_Next                               ' Increment the current Idx
RedoUDT:
        ' The following 'Clone' Code will now be done here, to cater for multiple Group_No lines
        Set cUDT = New cCBA_UDT
        If cUDT.CloneETL(sdMst(sdMst.Keys(lUDTIdx))) = False Then
            MsgBox lUDTIdx & " indexed UDT, Failed to Clone"
            Stop
            Exit Function
        End If
''        Set cUDT = sdMst(v)
        If Not cUDT Is Nothing Then
            'If cUDT.lUT_Grp_No = 1 And cUDT.lUT_Pos_Left = 1 Then
''            If cUDT.lUT_Grp_No = 1 Then
''                lVNo = lVNo
''            Else
''                lVNo = lVNo
''            End If
            If lVNo = 0 And lTotGrpNo = 0 And cUDT.lUT_Grp_No > 0 Then      ' If New a Grouped Line section
                lVNo = lUDTIdx
                If cUDT.lUT_Grp_No = 1 Then
                    lTotGrpNo = 10 ''iDocDataObject.GetGrpNo(cUDT.sUT_Procedure)
                    If lTotGrpNo < 1 Then
                        MsgBox "No Group Numbers found"
                        lTotGrpNo = 1
                    End If
                Else
                    lTotGrpNo = cUDT.lUT_Grp_No
                End If
                lMHeight = cUDT.lUT_Merge_Height
                lGrpNo = 1
            ElseIf lVNo > 0 And cUDT.lUT_Grp_No = 0 Then                    ' If the end of the grouped lines...
                If lGrpNo + 1 > lTotGrpNo Then
                    lVNo = 0: lGrpNo = 0: lTotGrpNo = 0
                Else                                                        ' If the next set of grouped lines...
                    lGrpNo = lGrpNo + 1: lUDTIdx = lVNo
                    lAccTop = lAccTop + lMHeight
'                    lCurSL_ID = 0
                    GoTo RedoUDT
                End If
            End If
            If lCurSL_ID <> cUDT.lUT_SL_ID Then
                lCurSL_ID = cUDT.lUT_SL_ID
                If bNextSeg = False Then
                    bNextSeg = True
                Else
                    lAccTop = lmaxTop - 1
                End If
            End If
            
            ' Set up the ID
            lDL_Idx = lDL_Idx + 1
            ' Add to the Top Position and change the UDT to reflect number and GroupNo
            cUDT.lUT_Pos_Top = cUDT.lUT_Pos_Top + lAccTop: cUDT.lUT_No = lDL_Idx: cUDT.lUT_Grp_No = lGrpNo
            If lGrpNo > 0 And cUDT.sUT_Name > "" Then cUDT.sUT_Name = cUDT.sUT_Name & CStr(lGrpNo)
            sDL_ID = g_Fmt_2_IDs(cUDT.lUT_Pos_Top, cUDT.lUT_Pos_Left, e_UTFldFmt.eUT_TopLeft, e_UTFldFmt.eUT_TopLeft)
            If sdUDT.Exists(sDL_ID) Then
                CBA_Error = sDL_ID & " (UDT_ID=" & cUDT.sUT_ID & ") UDT already exists"
                MsgBox CBA_Error
                Debug.Print CBA_Error                   ' RWCam Fix up when ready
''                Copy_SubMst_UDT = False
'                Stop
                GoTo NextUDT
                GoTo Exit_Routine
            End If
            ' Add the SubMst class altered data to the the new class
            'Call cUDT.Add_Class_Ltd_UT(lDL_Idx, cUDT.sUT_ID, 0, cUDT.lUT_Pos_Left, lTop, 0, cUDT.lUT_Merge_Width, cUDT.lUT_Merge_Height, NZ(cUDT.sUT_Name, ""), _
                                cUDT.sUT_Hdg_Type, NZ(cUDT.sUT_Data, ""), cstWRAPTEXT, cUDT.lUT_Font_Size, cstLOCKED, cUDT.lUT_BG_Color, cUDT.lUT_FG_Color, cUDT.sUT_BIUS, cstTEXTALIGN, cstFIELDTYPE, cstLINETYPE, NZ(cUDT.sUT_Border, ""))
            ' Increment the Top position so that a new unique SD key is formed
'            If cUDT.lUT_Pos_Left = 1 Then
'                lTop = lTop + cUDT.lUT_Merge_Height
'                lHeight = cUDT.lUT_Merge_Height
'            End If

            If lmaxTop < cUDT.lUT_Pos_Top + cUDT.lUT_Merge_Height Then lmaxTop = cUDT.lUT_Pos_Top + cUDT.lUT_Merge_Height
            sdUDT.Add sDL_ID, cUDT
        Else
            CBA_Error = "Why is the UDT Nothing"
            MsgBox CBA_Error
            Debug.Print CBA_Error                   ' RWCam Fix up when ready
            Stop
        End If
NextUDT:
    Loop
    Copy_SubMst_UDT = True

        
Exit_Routine:
    On Error Resume Next

    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_UDT_ETL.Copy_SubMst_UDT", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

''Private Function Copy_SubMst_UDT(sdMst As Scripting.Dictionary) As Boolean
''    ' This routine will reformat the UDT Lines from the copied sub-master UDT class
''    ' Will copy lines of UDTs if needed
''''    Dim UT_Pos_Top As Long, UT_Pos_Left As Long, UT_Merge_Width As Long, lTotal As Long, bValidLine As Boolean
''    Dim cUDT As cCBA_UDT, lDL_Idx As Long, sDL_ID As String '', DL_Data As String, lTop As Long, lHeight As Long
''    'dim cMst As cCBA_UDT_SubMst
''    Dim v As Variant, bNextSeg As Boolean, lAccTop As Long, lMinTop As Long, lMaxBot As Long, lCurSL_ID As Long, lLastTop As Long, lCurrTop As Long, lTempBot As Long, lTempTop As Long
''    On Error GoTo Err_Routine
''    Copy_SubMst_UDT = False: lDL_Idx = 0: lMinTop = 0: lMaxBot = 0: lCurrTop = 0 '': lInit_Next = 1
''    ' For each UDT in the SD...
''    For Each v In sdMst
''        Set cUDT = sdMst(v)
''        If Not cUDT Is Nothing Then
''            If lCurSL_ID <> cUDT.lUT_SL_ID Then
''                lCurSL_ID = cUDT.lUT_SL_ID
''                lCurrTop = 0: lLastTop = 0
''            End If
''
''            If lLastTop <= cUDT.lUT_Pos_Top Then
''                lLastTop = lCurrTop
''                lCurrTop = cUDT.lUT_Pos_Top
''                lTempTop = cUDT.lUT_Pos_Top
''                lMinTop = lMinTop + lCurrTop - lLastTop: lMaxBot = lMinTop + cUDT.lUT_Merge_Height - 1
''                lAccTop = lMinTop
''            End If
''            If lTempTop <> cUDT.lUT_Pos_Top Then
''                lTempTop = cUDT.lUT_Pos_Top
''                lAccTop = lMinTop + lTempTop - lCurrTop
''                lTempTop = lCurrTop
''            End If
''''            If lTempBot <> cUDT.lUT_Merge_Height - 1 Then
''''                lTempBot = cUDT.lUT_Merge_Height - 1
''''            End If
''
''            ' Set up the ID
''            lDL_Idx = lDL_Idx + 1
''            ' Add to the Top Position and change the UDT to reflect
''            cUDT.lUT_Pos_Top = lAccTop: cUDT.lUT_No = lDL_Idx
''            sDL_ID = g_Fmt_2_IDs(cUDT.lUT_Pos_Top, cUDT.lUT_Pos_Left, e_UTFldFmt.eUT_TopLeft, e_UTFldFmt.eUT_TopLeft)
''            If sdUDT.Exists(sDL_ID) Then
''                CBA_Error = sDL_ID & " (UDT_ID=" & cUDT.sUT_ID & ") UDT already exists"
''                MsgBox CBA_Error
''                Debug.Print CBA_Error                   ' RWCam Fix up when ready
''''                Copy_SubMst_UDT = False
'''                Stop
''                GoTo NextUDT
''                GoTo Exit_Routine
''            End If
''            ' Add the SubMst class altered data to the the new class
''            'Call cUDT.Add_Class_Ltd_UT(lDL_Idx, cUDT.sUT_ID, 0, cUDT.lUT_Pos_Left, lTop, 0, cUDT.lUT_Merge_Width, cUDT.lUT_Merge_Height, NZ(cUDT.sUT_Name, ""), _
''                                cUDT.sUT_Hdg_Type, NZ(cUDT.sUT_Data, ""), cstWRAPTEXT, cUDT.lUT_Font_Size, cstLOCKED, cUDT.lUT_BG_Color, cUDT.lUT_FG_Color, cUDT.sUT_BIUS, cstTEXTALIGN, cstFIELDTYPE, cstLINETYPE, NZ(cUDT.sUT_Border, ""))
''            ' Increment the Top position so that a new unique SD key is formed
'''            If cUDT.lUT_Pos_Left = 1 Then
'''                lTop = lTop + cUDT.lUT_Merge_Height
'''                lHeight = cUDT.lUT_Merge_Height
'''            End If
''
''''            If lMinTop < cUDT.lUT_Pos_Top + cUDT.lUT_Merge_Height Then lMinTop = cUDT.lUT_Pos_Top + cUDT.lUT_Merge_Height
''''            If iCBA_Doc_Data_GetUDTData(cUDT, "1stFormat") = False Then
''''                CBA_Error = sDL_ID & " (UDT_ID=" & cUDT.sUT_ID & ") UDT didn't format"
''''                MsgBox CBA_Error
''''                Debug.Print CBA_Error                   ' RWCam Fix up when ready
''''                Stop
''''            End If
''            sdUDT.Add sDL_ID, cUDT
''        Else
''            CBA_Error = "Why is the UDT Nothing"
''            MsgBox CBA_Error
''            Debug.Print CBA_Error                   ' RWCam Fix up when ready
''            Stop
''
''        End If
''NextUDT:
''    Next
''    Copy_SubMst_UDT = True
''    'Loop
''
''Exit_Routine:
''
''    On Error Resume Next
''
''    Exit Function
''Err_Routine:
''    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_UDT_ETL.Copy_SubMst_UDT", 3)
''    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
''    Debug.Print CBA_Error
''    Stop            ' ^RW Camera + next line
''    Resume Next
''    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
''    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
''    GoTo Exit_Routine
''    Resume Next
''End Function
''
Private Function Get_UDT_DB(ByVal lDoc_ID As Long) As Boolean
    ' This routine will get the existing doc, UDT Lines from the UT_Db; sSysType is the System Doc Type - Currently C[ategory] or T[en]
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim sSQL As String, cUDT As cCBA_UDT, DH_ID As Long, lDL_Idx As Long, sDL_ID As String, DL_Data As String
    On Error GoTo Err_Routine
    Get_UDT_DB = False
    CBA_ErrTag = "SQL"
    Set CN = New ADODB.Connection
    If sSysType = "T" Then
        CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("UDT", , , , , "Ten") & ";"
    ElseIf sSysType = "C" Then
        CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("UDT", , , , , "Cam") & ";"
    Else
        MsgBox sSysType & " database key not found"
        Stop
        GoTo Exit_Routine
    End If

    ' Get the Template Hdrs
    Set RS = New ADODB.Recordset
    sSQL = "SELECT * FROM qryL2_Doc_Lines WHERE DH_ID=" & lDoc_ID & " ORDER BY DL_Pos_Top, DL_Pos_Left;"
    Set RS = New ADODB.Recordset
    RS.Open sSQL, CN
    Do While Not RS.EOF
        Get_UDT_DB = True
        CBA_ErrTag = ""
''        If lDoc_ID < CBA_200_ID Then
''            lDoc_ID = NZ(RS!DH_ID, 0): sDoc_Name = NZ(RS!DH_Doc_Name, "")
''        End If
        'If sDoc_Name = "" Then sDoc_Name = NZ(RS!DH_Doc_Name, "")
        ' Set up the ID and it's array index to the dictionary
        lDL_Idx = lDL_Idx + 1
        sDL_ID = g_Fmt_2_IDs(RS!DL_Pos_Top, RS!DL_Pos_Left, e_UTFldFmt.eUT_TopLeft, e_UTFldFmt.eUT_TopLeft)
        If sdUDT.Exists(sDL_ID) Then
            MsgBox sDL_ID & " UDT already exists"
            GoTo Exit_Routine
        End If
        ' Add the ID and it's array index to the Doc UDT dictionary
        Set cUDT = New cCBA_UDT
        DL_Data = NZ(RS!DL_Data, "")
        Call cUDT.Add_Class_Ltd_UT(lDL_Idx, RS!DL_UT_ID, 0, RS!DL_Pos_Left, RS!DL_Pos_Top, 0, RS!DL_Merge_Width, RS!DL_Merge_Height, NZ(RS!DL_Name, ""), _
                            DL_Data, RS!DL_Font_Size, RS!DL_BG_Color, RS!DL_FG_Color, RS!DL_BIUS, NZ(RS!DL_Border, ""))
        sdUDT.Add sDL_ID, cUDT
        ' Next record
        RS.MoveNext
    Loop
    Get_UDT_DB = True
        
Exit_Routine:
    On Error Resume Next
''    Call SaveIP_UT("Set", False)
    CN.Close
    Set RS = Nothing
    Set CN = Nothing
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_UDT_ETL.Get_UDT_DB", 3)
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


''Public Function Copy_UDT_ETL(cUDT_IN As cCBA_UDT, lCopy_To_Top As Long) As Boolean
''    ' Will copy the Individual cCBA_UDT_ETL to the lCopy_To_Top Position
''    Dim lUTTop As Long, cUDT As cCBA_UDT
''    On Error GoTo Err_Routine
''    CBA_Error = "": Copy_UDT_ETL = False
''    lUTTop = cUDT_IN.lUT_Pos_Top
''    ' For the whole document
''    For Each cUDT In sdUDT
''        If lUTTop = cUDT.lUT_Pos_Top Then
''            cUDT_IN.lUT_Pos_Top = lCopy_To_Top
''            GoTo GotoEnd
''        End If
''    Next
''    MsgBox "'To' - Top Position not found - document not updated - please try again", vbOKOnly
''    GoTo Exit_Routine
''GotoEnd:
''    ' Add the line...
''    Copy_UDT_ETL = Upd_UDT_ETL(cUDT_IN, "New")
''Exit_Routine:
''
''    On Error Resume Next
''    Exit Function
''
''Err_Routine:
''    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-cCBA_UDT_ETL.Copy_UDT_ETL", 3)
''    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
''    Debug.Print CBA_Error
''    Stop            ' ^RW Camera + next line
''    Resume Next
''    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
''    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
''    GoTo Exit_Routine
''    Resume Next
''
''
''End Function

Public Function Upd_UDT_ETL(cUDT_IN As cCBA_UDT, ByVal sUpd_Del_New As String) As Boolean
    ' Will update the Individual cCBA_UDT_ETL lines - Will instigate re-indexing if sdUDT needs it
    Dim lUTTop As Long, sdic As Scripting.Dictionary, cUDT As cCBA_UDT, lAddTop As Long, sTopLeft As String
    On Error GoTo Err_Routine
    CBA_Error = "": Upd_UDT_ETL = False
    lUTTop = cUDT_IN.lUT_Pos_Top: lAddTop = 0
    Set sdic = New Scripting.Dictionary
    ' For the whole document
    For Each cUDT In sdUDT
        If lUTTop = cUDT.lUT_Pos_Top Then
            If sUpd_Del_New = "Upd" Then
                lAddTop = cUDT.lUT_Merge_Height - cUDT_IN.lUT_Merge_Height
                sTopLeft = g_Fmt_2_IDs(cUDT_IN.lUT_Pos_Top, cUDT_IN.lUT_Pos_Left, e_UTFldFmt.eUT_ID, e_UTFldFmt.eUT_ID)
                ' Create new SD
                sdic.Add sTopLeft, cUDT_IN
                GoTo GotoNext
            ElseIf sUpd_Del_New = "New" Then
                lAddTop = cUDT_IN.lUT_Merge_Height
                sTopLeft = g_Fmt_2_IDs(cUDT_IN.lUT_Pos_Top, cUDT_IN.lUT_Pos_Left, e_UTFldFmt.eUT_ID, e_UTFldFmt.eUT_ID)
                ' Create new SD
                sdic.Add sTopLeft, cUDT_IN
            ElseIf sUpd_Del_New = "Del" Then
                lAddTop = cUDT_IN.lUT_Merge_Height * -1
                GoTo GotoNext
            End If
        End If
        ' If lAddTop > 0...
        If lAddTop <> 0 Then cUDT.lUT_Pos_Top = cUDT.lUT_Pos_Top + lAddTop
        sTopLeft = g_Fmt_2_IDs(cUDT.lUT_Pos_Top, cUDT.lUT_Pos_Left, e_UTFldFmt.eUT_ID, e_UTFldFmt.eUT_ID)
        ' Create new SD
        sdic.Add sTopLeft, cUDT
GotoNext:
    Next
    ' If the SD has changed...
    If lAddTop <> 0 Then
        Set sdUDT = sdic
        Upd_UDT_ETL = True                   ' Will have to be re-displayed
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-cCBA_UDT_ETL.Upd_UDT_ETL", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
   
End Function

Public Function Write_UDT_ETL(ByRef lDoc_ID As Long, ByVal eDocType As e_DocuType) As Boolean
    ' This routine will delete any existing document lines and write the new Class Document & Lines to the UT_Db database
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim sSQL As String, lUT_Idx As Long, cUDT As cCBA_UDT, bHdrWrite As Boolean, lRecs As Long
    Dim v As Variant
    On Error GoTo Err_Routine
    
    ''CBA_BasicFunctions.CBA_Running "Saving Data for " & psSH_Desc
''    Application.ScreenUpdating = False
    ''If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Saving " & IIf(psNPD_R_MSO = "R", "Retender", psNPD_R_MSO) & " Data..."
    Write_UDT_ETL = False
    CBA_ErrTag = "SQL"
    bHdrWrite = False
''    Application.Cursor = xlWait
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("UDT") & ";"
    ' Set SQL to delete any lines that exist
    If lDoc_ID >= CBA_200_ID Then
        sSQL = "DELETE FROM L2_Doc_Lines WHERE DL_DH_ID=" & lDoc_ID
        RS.Open sSQL, CN
    End If
    CBA_ErrTag = "Lines"
    ' Get the UDT data
    For Each v In sdUDT
        Set cUDT = sdUDT(v)
        CBA_ErrTag = "SQL"
        ' Don't Add or Update headings
''        If cUDT.sUT_Hdg_Type = "Y" Then GoTo GoNext
        ' Add the header if not existing on the first record
        If lDoc_ID < CBA_200_ID Then
            'If sDoc_Name = "" Then sDoc_Name = Format(lDoc_ID, e_UTFldFmt.eUT_TopLeft)
            sSQL = "INSERT INTO L1_Doc_Hdrs (DH_SH_ID, DH_UpdUser, DH_CrtUser)" & Chr(10)
            sSQL = sSQL & "VALUES(" & eDocType & ",'" & CBA_User() & "','" & CBA_User() & "')"
            RS.Open sSQL, CN
            lDoc_ID = g_DLookup("DH_ID", "L1_Doc_Hdrs", "DH_ID > 0", "DH_ID DESC", g_GetDB("UDT"), 0)
            bHdrWrite = True
            If lDoc_ID = 0 Then
                MsgBox "DH_ID is 0 "
                Stop
            End If
        ' Else Update on the first record
        ElseIf bHdrWrite = False Then
            sSQL = "UPDATE L1_Doc_Hdrs SET "
'            If IsNumeric(sDoc_Name) = True And sDoc_Name <> Format(lDoc_ID, e_UTFldFmt.eUT_TopLeft) Then
'                sDoc_Name = Format(lDoc_ID, e_UTFldFmt.eUT_TopLeft)
'                sSQL = sSQL & "DH_Doc_Name='" & sDoc_Name & "',"
'            End If
            sSQL = sSQL & "DH_UpdUser='" & CBA_User & "', DH_UpdDate=" & g_GetSQLDate(Now(), CBA_DMYHN)
            sSQL = sSQL & " WHERE DH_ID=" & lDoc_ID
            RS.Open sSQL, CN
            bHdrWrite = True
        End If
        ' Add the Document
        sSQL = "INSERT INTO L2_Doc_Lines ( DL_DH_ID, DL_UT_ID, DL_Pos_Top, DL_Pos_Left, DL_Merge_Width, DL_Merge_Height, DL_IsGrp, DL_Name, DL_Data, DL_Font_Size, DL_BG_Color, DL_FG_Color, DL_BIUS, DL_Border)" & Chr(10)
        sSQL = sSQL & "VALUES ( "
        sSQL = sSQL & lDoc_ID & ","
        sSQL = sSQL & cUDT.sUT_ID & ","
        sSQL = sSQL & cUDT.lUT_Pos_Top & ","
        sSQL = sSQL & cUDT.lUT_Pos_Left & ","
        sSQL = sSQL & cUDT.lUT_Merge_Width & ","
        sSQL = sSQL & cUDT.lUT_Merge_Height & ","
        sSQL = sSQL & IIf(cUDT.lUT_Grp_No > 0, True, False) & ","
        sSQL = sSQL & "'" & cUDT.sUT_Name & "',"
        sSQL = sSQL & "'" & Replace(cUDT.sUT_Data, "'", "`") & "',"
        sSQL = sSQL & cUDT.lUT_Font_Size & ","
        sSQL = sSQL & cUDT.lUT_BG_Color & ","
        sSQL = sSQL & cUDT.lUT_FG_Color & ","
        sSQL = sSQL & "'" & cUDT.sUT_BIUS & "',"
        sSQL = sSQL & "'" & cUDT.sUT_Border & "'"
        sSQL = sSQL & " )"
        CBA_ErrTag = sSQL
        RS.Open sSQL, CN
        lRecs = lRecs + 1
GoNext:
    Next
    Write_UDT_ETL = True

Exit_Routine:

    On Error Resume Next
    CN.Close
    Set RS = Nothing
    Set CN = Nothing
''    ' Close the worksheet
''    Application.Cursor = xlDefault
''    Application.ScreenUpdating = True
''    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    Exit Function
    
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_UDT_ETL.Write_UDT_ETL", 3)
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

Public Function Get_SysType(ByVal DocType As e_DocuType, Optional ByVal DocMetric As String = "SH_SysType", Optional CN As ADODB.Connection) As Variant
    ' Will get the 'System Document Type', or various other data
    Dim sReturn As String, sSQL As String, RS As ADODB.Recordset, lIdx As Long
    Static arrDB(), bActive As Boolean
        
    On Error GoTo Err_Routine
    If bActive = False Then
        If CN Is Nothing Then
            Set CN = New ADODB.Connection
            CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("UDT", , , , , "Test") & ";"
        End If
        Set RS = New ADODB.Recordset
        sSQL = "SELECT * FROM C1_Seg_Template_Hdrs WHERE SH_ID =" & DocType & ";"
        RS.Open sSQL, CN
        arrDB = RS.GetRows
        bActive = True
    End If
    ' Find the field required
    Select Case DocMetric
        Case Is = "SH_Desc"
            Get_SysType = arrDB(1, 0)
        Case Is = "SysType", "SH_SysType"
            Get_SysType = arrDB(2, 0)
        Case Is = "Col_Width"
            Get_SysType = arrDB(4, 0)
        Case Is = "Row_Height"
            Get_SysType = arrDB(5, 0)
        Case Is = "No_Of_Cols"
            Get_SysType = arrDB(6, 0)
        Case Is = "Page_Margin"
            Get_SysType = arrDB(7, 0)
        Case Is = "Password"
            Get_SysType = NZ(arrDB(8, 0), "")
        Case Else
            MsgBox DocMetric & " not found"
    End Select

Exit_Routine:
    On Error Resume Next
    Set RS = Nothing
    ''Set CN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_UDT_ETL.Get_SysType", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function

Private Property Get lCurrUDTIdx() As Long: lCurrUDTIdx = plCurrUDTIdx: End Property
Private Property Let lCurrUDTIdx(ByVal lNewValue As Long): plCurrUDTIdx = lNewValue: End Property

Public Property Get sdUDT() As Scripting.Dictionary: Set sdUDT = psdUDT: End Property
Private Property Set sdUDT(ByVal objNewValue As Scripting.Dictionary): Set psdUDT = objNewValue: End Property

Private Property Get sSysType() As String: sSysType = psSysType: End Property
Private Property Let sSysType(ByVal sNewValue As String): psSysType = sNewValue: End Property
'
'Private Property Get eDocType() As Long: eDocType = peDocType: End Property
'Private Property Let eDocType(ByVal lNewValue As Long): peDocType = lNewValue: End Property

'Private Property Get lDoc_ID() As Long: lDoc_ID = plDoc_ID: End Property
'Private Property Let lDoc_ID(ByVal lNewValue As Long): plDoc_ID = lNewValue: End Property
'
'Private Property Get sDoc_Name() As String: sDoc_Name = psDoc_Name: End Property
'Private Property Let sDoc_Name(ByVal sNewValue As String): psDoc_Name = sNewValue: End Property

Public Property Get COL_WIDTH() As Single: COL_WIDTH = psngCol_Width: End Property
Private Property Let COL_WIDTH(ByVal NewValue As Single): psngCol_Width = NewValue: End Property
Public Property Get ROW_HEIGHT() As Single: ROW_HEIGHT = psngRow_Height: End Property
Private Property Let ROW_HEIGHT(ByVal NewValue As Single): psngRow_Height = NewValue: End Property
Public Property Get NO_OF_COLS() As Single: NO_OF_COLS = psngNo_Of_Cols: End Property
Private Property Let NO_OF_COLS(ByVal NewValue As Single): psngNo_Of_Cols = NewValue: End Property
Public Property Get PAGE_MARGIN() As Single: PAGE_MARGIN = psngPage_Margin: End Property
Private Property Let PAGE_MARGIN(ByVal NewValue As Single): psngPage_Margin = NewValue: End Property



