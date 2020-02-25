VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTEN6_Tenders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit       ' cTEN6_Tenders Changed 190808
Public plTH_ID As Long                 ' Level ID - This is the overall table/number that ties the Tender/TLines/Version/ProductCode/Contract and ATP Document together
Private pTH_Ary() As Variant
Private plTH_Idx As Long, plTU_Idx As Long
Private Const lcTH_ID = 0, lcTH_LV_ID = 1, lcTH_Filename = 2, lcTH_SU_ID = 3, lcTH_Upd = 4               ' Make the Upd Const the largest element number to dimension arrays etc
Private pTU_Ary() As Variant
Const ALL_FLDS = "TU_ID,TH_ID,TH_LV_ID,TH_Filename,TH_SU_ID,TU_TH_ID,TU_Value,TU_UpdDate,TU_UpdUser,TH_UpdDate,TH_CrtUser"
Private Const lcTU_ID = 0, lcTU_UT_ID = 1, lcTU_Value = 2, lcTU_UpdUser = 3, lcTU_UpdDate = 4, lcTU_Upd = 5

Private psdTenders As Scripting.Dictionary
Private psdTLines As Scripting.Dictionary

Public Sub Generate(ByVal TH_ID As Long, sPath As String, sFile As String)
    Dim lIdx As Long
    ' Generate the Class
    plTH_Idx = -1: plTU_Idx = -1
    plTH_ID = TH_ID
    Set psdTenders = New Scripting.Dictionary
    Set psdTLines = New Scripting.Dictionary
''    Set psdStatuses = New Scripting.Dictionary
    ReDim pTH_Ary(0 To lcTH_Upd, 0 To 0)
    ReDim pTU_Ary(0 To lcTU_Upd, 0 To 0)
    If TH_ID > 0 Then
        Call Get_db_Class
    Else
        Call Add_Class_TH(TH_ID, "", 1, "", "A")
        Call Add_Class_TL(0, 0, "", CBA_User, "", "A")
    End If
End Sub

Public Function Get_db_Class() As String
    ' This routine will get the Class Lines
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim sSQL As String, lRecs As Long
    Dim TH_ID As Long, TH_LV_ID As Long, TH_Filename As String, TH_SU_ID As Long
    Dim TU_ID As Long, TU_TH_ID As Long, TU_UT_ID As Long, TU_Value As String
    Dim TU_UpdDate As String, TU_UpdUser As String
    Dim vVar, lID As Long

    On Error GoTo Err_Routine
    
    CBA_ErrTag = "SQL": lRecs = 0
    Set CN = New ADODB.Connection
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Ten") & ";"
       
    ' Get the Tender records and insert into the class module arrays / dictionaries
    Set RS = New ADODB.Recordset
    sSQL = "SELECT * FROM qry_L9_TDoc_UDTs WHERE TH_ID=" & plTH_ID
    RS.Open sSQL, CN
    Do While Not RS.EOF
        If lRecs = 0 Then
            TH_ID = plTH_ID: TH_LV_ID = NZ(RS!TH_LV_ID, 0): TH_Filename = NZ(RS!TH_Filename, "")
            TH_SU_ID = NZ(RS!TH_SU_ID, 0)
            Call Add_Class_TH(TH_ID, TH_LV_ID, TH_Filename, TH_SU_ID, "")
        End If
        lRecs = lRecs + 1
        TU_ID = NZ(RS!TU_ID, 0)
        TU_UT_ID = NZ(RS!TU_UT_ID, 0): TU_Value = NZ(RS!TU_Value, ""):
        TU_UpdUser = NZ(RS!TU_UpdUser, ""): TU_UpdDate = NZ(RS!TU_UpdDate, "")
        Call Add_Class_TL(TU_ID, TU_UT_ID, TU_Value, TU_UpdUser, TU_UpdDate, "")
        RS.MoveNext
    Loop
    
Exit_Routine:
    On Error Resume Next
    CN.Close
    Set RS = Nothing
    Set CN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("clsTenders-f-Get_db_Class", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Sub Add_Class_TH(TH_ID As Long, TH_LV_ID As Long, TH_Filename As String, TH_SU_ID As Long, ByVal TH_Upd As String)
    ' Add a Header
    plTH_Idx = plTH_Idx + 1
    If TH_ID < 0 Then TH_ID = plTH_Idx
    If psdTenders.Exists(TH_ID) Then
        MsgBox TH_ID & " already exists"
        Exit Sub
    End If
    psdTenders.Add TH_ID, plTH_Idx
    ReDim Preserve pTH_Ary(0 To lcTH_Upd, 0 To plTH_Idx)
    pTH_Ary(lcTH_ID, plTH_Idx) = TH_ID                                      ' Class Auto-number - (If 0, hasn't been created yet)
    pTH_Ary(lcTH_LV_ID, plTH_Idx) = TH_LV_ID                                  ' Class File Desc
    pTH_Ary(lcTH_Filename, plTH_Idx) = TH_Filename
    pTH_Ary(lcTH_SU_ID, plTH_Idx) = TH_SU_ID
    pTH_Ary(lcTH_Upd, plTH_Idx) = p_UpdFlag(CStr(pTH_Ary(lcTH_Upd, plTH_Idx)), TH_Upd)
End Sub


Sub Add_Class_TL(TU_ID As Long, TU_UT_ID As Long, TU_Value As String, TU_UpdUser As String, TU_UpdDate As String, ByVal TU_Upd As String)
    Dim lIdx As Long, lSIdx As Long, bfound As Boolean
    Static lSupp_ID As Long
    ' Add a Class Line
    plTU_Idx = plTU_Idx + 1
    If TU_ID < 0 Then TU_ID = plTU_Idx
    If psdTLines.Exists(TU_ID) Then
        MsgBox TU_ID & " already exists"
        Exit Sub
    End If
    psdTLines.Add TU_ID, plTU_Idx
    ReDim Preserve pTU_Ary(0 To lcTU_Upd, 0 To plTU_Idx)
    ' Add the parts of the line to the array
    pTU_Ary(lcTU_ID, plTU_Idx) = TU_ID                                                  ' Line Auto-number
    pTU_Ary(lcTU_UT_ID, plTU_Idx) = TU_UT_ID
    pTU_Ary(lcTU_Value, plTU_Idx) = TU_Value
    pTU_Ary(lcTU_UpdDate, plTU_Idx) = TU_UpdDate
    pTU_Ary(lcTU_UpdUser, plTU_Idx) = TU_UpdUser

    pTU_Ary(lcTU_Upd, plTU_Idx) = p_UpdFlag(CStr(pTU_Ary(lcTU_Upd, plTU_Idx)), TU_Upd)    ' Update Flag
    '''pTU_Ary(lcTU_Sts_ID, plTU_Idx) = Get_Class_TL(TU_ID, "TU_Sts_ID")                  ' Stat
End Sub

Sub Upd_Class_TH(ByVal TH_ID As Long, sField As String, ByVal NewValue, ByRef TH_Upd As String)
    ' Upd Class Header
    Dim lIdx As Long, lIdx1 As Long, sUpd As String
    ' Update the value
    If sField = "TH_LV_ID" Then
        pTH_Ary(lcTH_LV_ID, lIdx) = NewValue                                                 ' Class line number
        pTH_Ary(lcTH_Upd, lIdx) = p_UpdFlag(CStr(pTH_Ary(lcTH_Upd, lIdx)), TH_Upd)          ' Upd Flag "A" for add / "U" for Upd / "N" for Not updated
    ElseIf sField = "TH_Filename" Then
        pTH_Ary(lcTH_Filename, lIdx) = NewValue
        pTH_Ary(lcTH_Upd, lIdx) = p_UpdFlag(CStr(pTH_Ary(lcTH_Upd, lIdx)), TH_Upd)
    ElseIf sField = "TH_SU_ID" Then
        pTH_Ary(lcTH_SU_ID, lIdx) = NewValue
        pTH_Ary(lcTH_Upd, lIdx) = p_UpdFlag(CStr(pTH_Ary(lcTH_Upd, lIdx)), TH_Upd)
    Else
        MsgBox "Field " & sField & " not found"
    End If
End Sub

Sub Upd_Class_TL(ByVal TU_ID As Long, sField As String, ByVal NewValue, ByRef TU_Upd As String)
    Dim lIDPosx As Long, sFlag As String, lIDPos As Long
    ' Get the Index to use
    lIDPosx = Get_Line_Idx(TU_ID)
    ' Update the values
    If sField = "TU_UT_ID" Then
        pTU_Ary(lcTU_UT_ID, plTU_Idx) = NZ(NewValue, 0)                                        ' Line Foreign key to the Version
        lIDPos = lcTU_UT_ID
    ElseIf sField = "TU_Value" Then
        pTU_Ary(lcTU_Value, plTU_Idx) = NZ(NewValue, "")
        lIDPos = lcTU_Value
    ElseIf sField = "TU_UpdDate" Then
        pTU_Ary(lcTU_UpdDate, plTU_Idx) = NZ(NewValue, "")
        lIDPos = -1
    ElseIf sField = "TU_UpdUser" Then
        pTU_Ary(lcTU_UpdUser, plTU_Idx) = NZ(NewValue, "")
        lIDPos = -1
    Else
        MsgBox "Field " & sField & " not found"
        Exit Sub
    End If
    ' If an update...
    If TU_Upd > "" And lIDPos > -1 Then
        sFlag = p_StsUpdFlag(CStr(pTU_Ary(lcTU_Upd, lIDPosx)), lIDPos)                      ' Upd Flag "A" for add / "U" for Upd / "N" for Not updated
        Call p_StsUpdFlag(CStr(pTU_Ary(lcTU_Upd, lIDPosx)), lIDPos, sFlag)
    End If
End Sub

Function Get_Class_TH(ByVal TH_ID As Long, sField As String, Optional bAdd As Boolean = False, Optional bSQL As Boolean = False, Optional sTablePrefix As String = "")
    ' Get a Header value
    Dim lIdx As Long, lIdx1 As Long
    If Not psdTenders.Exists(TH_ID) Then
        MsgBox "Hdr Line " & TH_ID & " doesn't exist"
        Exit Function
    Else
        lIdx = psdTenders.Item(TH_ID)
    End If
    ' Get the value
    If sField = "TH_LV_ID" Then                                                        ' Class File SEQ# - may change if lines are incremented
        Get_Class_TH = pTH_Ary(lcTH_LV_ID, lIdx)
    ElseIf sField = "TH_Filename" Then
        Get_Class_TH = pTH_Ary(lcTH_Filename, lIdx)
    ElseIf sField = "TH_SU_ID" Then
        Get_Class_TH = pTH_Ary(lcTH_SU_ID, lIdx)
    Else
        MsgBox "Field " & sField & " not found"
    End If
    ' If required for SQL...
    If bSQL Then
        Get_Class_TH = p_SQL_Field(sField, Get_Class_TH, bAdd, bSQL, sTablePrefix)
    End If
    
End Function

Function Get_Class_TL(ByVal TU_ID As Long, sField As String, Optional bAdd As Boolean = False, Optional bSQL As Boolean = False, Optional sTablePrefix As String = "")
      Dim lIdx As Long, bfound As Boolean
    ' Get the Index to use
    lIdx = Get_Line_Idx(TU_ID)
    ' Get the value
    If sField = "TU_ID" Then
        Get_Class_TL = pTU_Ary(lcTU_ID, lIdx)                                         ' Line Level ID
    ElseIf sField = "TU_UT_ID" Then
        Get_Class_TL = pTU_Ary(lcTU_UT_ID, lIdx)                                         ' Line Foreign key to the TLines
    ElseIf sField = "TU_Value" Then
        Get_Class_TL = pTU_Ary(lcTU_Value, lIdx)
    ElseIf sField = "TU_UpdDate" Then
        Get_Class_TL = pTU_Ary(lcTU_UpdDate, lIdx)
    ElseIf sField = "TU_UpdUser" Then
        Get_Class_TL = pTU_Ary(lcTU_UpdUser, lIdx)
    Else
        MsgBox "Field " & sField & " not found"
        Exit Function
    End If
  ' If required for SQL...
    If bSQL Then
        Get_Class_TL = p_SQL_Field(sField, Get_Class_TL, bAdd, bSQL, sTablePrefix)
    End If
End Function

Function Get_Line_Idx(ByVal TU_ID As Long) As Long
    ' Will get the line array index number
    If Not psdTLines.Exists(TU_ID) Then
        Get_Line_Idx = TU_ID                            ' Will be this one before the records are written
    Else
        Get_Line_Idx = psdTLines.Item(TU_ID)        ' Will be the TU_ID key after the record is written
    End If
End Function

Function p_SQL_Field(sField As String, vVal, bAdd As Boolean, bSQL As Boolean, Optional sTablePrefix As String = "")
    ' Procees the SQL Values as an Update field ("Field1=[formatted field]) or an Add field ("[formatted field])
    If bSQL Then
        If NZ(vVal, "") = "" Or vVal = "NULL" Then
            vVal = "NULL"
        ElseIf IsNumeric(vVal) Then
        ElseIf g_IsDate(vVal, True) Then
            If g_IsDate(vVal) = False Then vVal = g_FixDate(vVal, CBA_DMYHN)
            vVal = g_GetSQLDate(vVal, CBA_DMYHN)
        Else
            vVal = "'" & Replace(vVal, "'", "`") & "'"
        End If
    End If
    If Not bAdd And bSQL Then
        p_SQL_Field = sTablePrefix & sField & "=" & vVal
    Else
        p_SQL_Field = vVal
    End If
    
End Function

Private Function p_StsUpdFlag(sInputFlags As String, ByVal lPosNo As Long, Optional sInputFlag As String = "") As String
    ' This function will set or get the various Update flags.
    If lPosNo < 0 Then
        p_StsUpdFlag = "NNNNNNNN"
    ElseIf sInputFlag = "" Then
        lPosNo = lPosNo + 1
        p_StsUpdFlag = Mid(sInputFlags, lPosNo, 1)
    ElseIf sInputFlag > "" Then
        lPosNo = lPosNo + 1
        Mid(sInputFlags, lPosNo, 1) = sInputFlag
        p_StsUpdFlag = sInputFlags
    End If
End Function

Private Function p_UpdFlag(sArray_Upd As String, sUpdFlag As String) As String
    ' Set the Update Flag (for the entire record)
    p_UpdFlag = sArray_Upd
    If sUpdFlag = "" Then
    ElseIf sUpdFlag = "A" And sArray_Upd <> "A" Then
        p_UpdFlag = sUpdFlag
    ElseIf sUpdFlag = "U" And sArray_Upd = "N" Then
        p_UpdFlag = sUpdFlag
    End If
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    Call g_EraseAry(pTH_Ary)
    Call g_EraseAry(pTU_Ary)
    Set psdTLines = Nothing
    Set psdTenders = Nothing
End Sub




