VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCBA_UDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit        ' cCBA_UDT Changed 191106
    ' This class will hold all the UDT (User Defined Types) for the Excel area covered (Note: all valid UDTs are loaded into memory)
Private m_lUT_No As Long
Private m_sUT_ID As String
Private m_sUT_ID_Ext As String
Private m_lUT_SL_ID As Long
Private m_lUT_Pos_Left As Long
Private m_lUT_Pos_Top As Long
Private m_lUT_Grp_No As Long
Private m_lUT_Orig_Grp_No As Long
Private m_lUT_Merge_Width As Long
Private m_lUT_Merge_Height As Long
Private m_lUT_Link_ID As Long
Private m_sUT_Name As String
Private m_sUT_Hdg_Type As String
Private m_sUT_Default_Value As String
Private m_sUT_Original_Value As String
Private m_sUT_Wrap_Text As String
Private m_lUT_Font_Size As Long
Private m_sUT_Locked As String
Private m_lUT_BG_Color As Long
Private m_lUT_FG_Color As Long
Private m_sUT_Bold As String
Private m_sUT_Italic As String
Private m_sUT_Underline As String
Private m_sUT_StrikeThrough As String
Private m_sUT_TextAlign As String
Private m_sUT_Field_Type As String
Private m_sUT_Formula As String
Private m_sUT_Procedure As String
Private m_sUT_Border As String
Private m_sUT_Line_Type As String
Private m_sUT_Cond_Format As String
Private m_sUT_Hyperlink As String
Private m_sUT_Image As String
Private m_sUT_Source As String ' This will hold the source of the line - S=Come from CBIs, T=Comes from Table, D=Comes from Default, E=User (or Application) Entry - see p_Source for more data
Private m_sUT_Addit As String
Private m_sUT_Upd As String
Public Function CloneETL(ByRef cUDT As cCBA_UDT) As Boolean '@TP
m_lUT_No = CLng(cUDT.lUT_No)
m_sUT_ID = CStr(cUDT.sUT_ID)
m_sUT_ID_Ext = CStr(cUDT.sUT_ID_Ext)
m_lUT_SL_ID = CLng(cUDT.lUT_SL_ID)
m_lUT_Pos_Left = CLng(cUDT.lUT_Pos_Left)
m_lUT_Pos_Top = CLng(cUDT.lUT_Pos_Top)
m_lUT_Grp_No = CLng(cUDT.lUT_Grp_No)
m_lUT_Orig_Grp_No = CLng(cUDT.lUT_Orig_Grp_No)
m_lUT_Merge_Width = CLng(cUDT.lUT_Merge_Width)
m_lUT_Merge_Height = CLng(cUDT.lUT_Merge_Height)
m_lUT_Link_ID = CLng(cUDT.lUT_Link_ID)
m_sUT_Name = CStr(cUDT.sUT_Name)
m_sUT_Hdg_Type = CStr(cUDT.sUT_Hdg_Type)
m_sUT_Default_Value = CStr(cUDT.sUT_Default_Value)
m_sUT_Original_Value = CStr(cUDT.sUT_Original_Value)
m_sUT_Wrap_Text = CStr(cUDT.sUT_Wrap_Text)
m_lUT_Font_Size = CLng(cUDT.lUT_Font_Size)
m_sUT_Locked = CStr(cUDT.sUT_Locked)
m_lUT_BG_Color = CLng(cUDT.lUT_BG_Color)
m_lUT_FG_Color = CLng(cUDT.lUT_FG_Color)
m_sUT_Bold = CStr(cUDT.sUT_Bold)
m_sUT_Italic = CStr(cUDT.sUT_Italic)
m_sUT_Underline = CStr(cUDT.sUT_Underline)
m_sUT_StrikeThrough = CStr(cUDT.sUT_StrikeThrough)
m_sUT_TextAlign = CStr(cUDT.sUT_TextAlign)
m_sUT_Field_Type = CStr(cUDT.sUT_Field_Type)
m_sUT_Formula = CStr(cUDT.sUT_Formula)
m_sUT_Procedure = CStr(cUDT.sUT_Procedure)
m_sUT_Border = CStr(cUDT.sUT_Border)
m_sUT_Line_Type = CStr(cUDT.sUT_Line_Type)
m_sUT_Cond_Format = CStr(cUDT.sUT_Cond_Format)
m_sUT_Hyperlink = CStr(cUDT.sUT_Hyperlink)
m_sUT_Image = CStr(cUDT.sUT_Image)
m_sUT_Source = CStr(cUDT.sUT_Source)
m_sUT_Addit = CStr(cUDT.sUT_Addit)
m_sUT_Upd = CStr(cUDT.sUT_Upd)

If Err.Number = 0 Then CloneETL = True
End Function
Public Sub Add_Class_UT(UT_No As Long, UT_ID As String, UT_ID_Ext As String, UT_SL_ID As Long, UT_Pos_Left As Long, UT_Pos_Top As Long, UT_Grp_No As Long, UT_Merge_Width As Long, UT_Merge_Height As Long, _
                UT_Link_ID As Long, UT_Name As String, UT_Hdg_Type As String, UT_Default_Value As String, UT_Wrap_Text As String, UT_Font_Size As Long, UT_Locked As String, UT_BG_Color As Long, _
                UT_FG_Color As Long, UT_Bold As String, UT_Italic As String, UT_Underline As String, UT_StrikeThrough As String, UT_TextAlign As String, UT_Field_Type As String, UT_Formula As String, _
                UT_Procedure As String, UT_Line_Type As String, UT_Border As String, UT_Cond_Format As String, UT_Hyperlink As String, UT_Image As String, UT_Source As String, UT_Addit As String, UT_Orig_Grp_No As Long)

    ' Add the parts of the line to the array
    lUT_No = UT_No
    sUT_ID = UT_ID
    sUT_ID_Ext = UT_ID_Ext
    lUT_SL_ID = UT_SL_ID
    lUT_Grp_No = UT_Grp_No
    lUT_Orig_Grp_No = UT_Orig_Grp_No
    lUT_Pos_Left = UT_Pos_Left
    lUT_Pos_Top = UT_Pos_Top
    lUT_Merge_Width = UT_Merge_Width
    lUT_Merge_Height = UT_Merge_Height
    lUT_Link_ID = UT_Link_ID
    sUT_Hdg_Type = UT_Hdg_Type
    sUT_Name = UT_Name
    sUT_Default_Value = UT_Default_Value
    sUT_Original_Value = UT_Default_Value     ' Save the original value - this can be restored
    sUT_Wrap_Text = UT_Wrap_Text
    lUT_Font_Size = UT_Font_Size
    sUT_Locked = UT_Locked
    lUT_BG_Color = UT_BG_Color
    lUT_FG_Color = UT_FG_Color
    sUT_Bold = UT_Bold
    sUT_Italic = UT_Italic
    sUT_Underline = UT_Underline
    sUT_StrikeThrough = UT_StrikeThrough
    sUT_TextAlign = UT_TextAlign
    sUT_Field_Type = UT_Field_Type
    sUT_Formula = UT_Formula
    sUT_Procedure = UT_Procedure
    sUT_Line_Type = UT_Line_Type
    sUT_Border = UT_Border
    sUT_Cond_Format = UT_Cond_Format
    sUT_Hyperlink = UT_Hyperlink
    sUT_Image = UT_Image
    If UT_Source = "" Then UT_Source = "DNNN"
    sUT_Source = UT_Source
    sUT_Addit = UT_Addit
    If sUT_Upd = "" Then
        If UT_Hdg_Type = "N" Or UT_Locked = "N" Then
            sUT_Upd = "A"                           ' Make an Add for now - the app will check to see if there are existing records and set it to N if there is
        Else
            sUT_Upd = "N"
        End If
    End If
       
End Sub

Public Sub Add_Class_Ltd_UT(UT_No As Long, UT_ID As String, UT_SL_ID As Long, UT_Pos_Left As Long, UT_Pos_Top As Long, UT_Grp_No As Long, UT_Merge_Width As Long, UT_Merge_Height As Long, _
                UT_Name As String, UT_Data As String, UT_Font_Size As Long, UT_BG_Color As Long, UT_FG_Color As Long, UT_BUIS, UT_Border As String)

    ' Add the parts of the line to the array
    lUT_No = UT_No
    sUT_ID = UT_ID
''    sUT_ID_Ext = UT_ID_Ext
    lUT_SL_ID = UT_SL_ID
    lUT_Grp_No = UT_Grp_No
    lUT_Orig_Grp_No = UT_Grp_No
    lUT_Pos_Left = UT_Pos_Left
    lUT_Pos_Top = UT_Pos_Top
    lUT_Merge_Width = UT_Merge_Width
    lUT_Merge_Height = UT_Merge_Height
    lUT_Link_ID = 0
    sUT_Name = UT_Name
    Let sUT_Data = UT_Data
    sUT_Original_Value = sUT_Default_Value
''    sUT_Wrap_Text = UT_Wrap_Text
    lUT_Font_Size = UT_Font_Size
''    sUT_Locked = UT_Locked
    lUT_BG_Color = UT_BG_Color
    lUT_FG_Color = UT_FG_Color
    Call LetsUT_BIUS(UT_BUIS)
''    sUT_TextAlign = UT_TextAlign
''    sUT_Field_Type = UT_Field_Type
    Call LetsUT_Source("DNNNN")
    If sUT_Upd = "" Then
        If sUT_Hdg_Type = "N" Or sUT_Locked = "N" Then
            sUT_Upd = "A"                           ' Make an Add for now - the app will check to see if there are existing records and set it to N if there is
        Else
            sUT_Upd = "N"
        End If
    End If

End Sub

Public Sub Get_Class_UT(UT_No As Long, UT_ID As String, UT_ID_Ext As String, UT_SL_ID As Long, UT_Pos_Left As Long, UT_Pos_Top As Long, UT_Grp_No As Long, UT_Merge_Width As Long, UT_Merge_Height As Long, UT_Link_ID As Long, _
                 UT_Name As String, UT_Hdg_Type As String, UT_Default_Value As String, UT_Wrap_Text As String, UT_Font_Size As Long, UT_Locked As String, UT_BG_Color As Long, UT_FG_Color As Long, _
                 UT_Bold As String, UT_Italic As String, UT_Underline As String, UT_StrikeThrough As String, UT_TextAlign As String, UT_Field_Type As String, UT_Formula As String, _
                 UT_Procedure As String, UT_Line_Type As String, UT_Border As String, UT_Cond_Format As String, UT_Hyperlink As String, UT_Image As String, UT_Source As String, UT_Addit As String, UT_Orig_Grp_No As Long, UT_Upd As String)

    ' Add the parts of the line to the array
    UT_No = lUT_No
    UT_ID = sUT_ID
    UT_ID_Ext = sUT_ID_Ext
    UT_SL_ID = lUT_SL_ID
    UT_Grp_No = lUT_Grp_No
    UT_Orig_Grp_No = lUT_Orig_Grp_No
    UT_Pos_Left = lUT_Pos_Left
    UT_Pos_Top = lUT_Pos_Top
    UT_Merge_Width = lUT_Merge_Width
    UT_Merge_Height = lUT_Merge_Height
    UT_Link_ID = lUT_Link_ID
    UT_Hdg_Type = sUT_Hdg_Type
    UT_Name = sUT_Name
    UT_Default_Value = sUT_Default_Value
    UT_Wrap_Text = sUT_Wrap_Text
    UT_Font_Size = lUT_Font_Size
    UT_Locked = sUT_Locked
    UT_BG_Color = lUT_BG_Color
    UT_FG_Color = lUT_FG_Color
    UT_Bold = sUT_Bold
    UT_Italic = sUT_Italic
    UT_Underline = sUT_Underline
    UT_StrikeThrough = sUT_StrikeThrough
    UT_TextAlign = sUT_TextAlign
    UT_Field_Type = sUT_Field_Type
    UT_Formula = sUT_Formula
    UT_Procedure = sUT_Procedure
    UT_Line_Type = sUT_Line_Type
    UT_Border = sUT_Border
    UT_Cond_Format = sUT_Cond_Format
    UT_Hyperlink = sUT_Hyperlink
    UT_Image = sUT_Image
    UT_Source = sUT_Source
    UT_Addit = sUT_Addit
    UT_Upd = sUT_Upd
End Sub

Public Sub Get_Class_Ltd_UT(UT_Pos_Left As Long, UT_Pos_Top As Long, UT_SL_ID As Long, UT_Grp_No As Long, UT_Hdg_Type As String, UT_Default_Value As String, _
                      UT_Field_Type As String, UT_Source As String, UT_Addit As String, UT_Orig_Grp_No As Long, UT_Upd As String)

    ' Add the parts of the line to the array
    UT_SL_ID = lUT_SL_ID
    UT_Grp_No = lUT_Grp_No
    UT_Orig_Grp_No = lUT_Orig_Grp_No
    UT_Pos_Left = lUT_Pos_Left
    UT_Pos_Top = lUT_Pos_Top
    UT_Hdg_Type = sUT_Hdg_Type
    UT_Default_Value = sUT_Default_Value
    UT_Field_Type = sUT_Field_Type
    UT_Source = sUT_Source
    UT_Addit = sUT_Addit
    UT_Upd = sUT_Upd
End Sub

Public Sub Rest_Class_UT()
    Dim lIDPosx As Long, sFlag As String, lIDPos As Long
    ' Will restore the Original_Value into the Default_Value
    sUT_Default_Value = sUT_Original_Value
    If sUT_Hdg_Type = "N" Or sUT_Locked = "N" Then
        sUT_Upd = "A"                           ' Make an Add for now - the app will check to see if there are existing records in the table and set it to N if there is
    Else
        sUT_Upd = "N"
    End If
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
''    Set UT_Rec = Nothing
''    Call g_EraseAry(pSH_Ary)
''    Call g_EraseAry(pSL_Ary)
''    Call g_EraseAry(pUT_Ary)
''    Set psdUT = Nothing
''    Set psdSL = Nothing
''    Set psdSH = Nothing
End Sub

Public Property Get lUT_No() As Long: lUT_No = m_lUT_No: End Property
Public Property Let lUT_No(ByVal lNewValue As Long): m_lUT_No = lNewValue: End Property

Public Property Get sUT_ID() As String: sUT_ID = m_sUT_ID: End Property
Public Property Let sUT_ID(ByVal sNewValue As String): m_sUT_ID = sNewValue: End Property

Public Property Get sUT_ID_Ext() As String: sUT_ID_Ext = m_sUT_ID_Ext: End Property
Public Property Let sUT_ID_Ext(ByVal sNewValue As String): m_sUT_ID_Ext = sNewValue: End Property

Public Property Get lUT_SL_ID() As Long: lUT_SL_ID = m_lUT_SL_ID: End Property
Public Property Let lUT_SL_ID(ByVal lNewValue As Long): m_lUT_SL_ID = lNewValue: End Property

Public Property Get lUT_Pos_Left() As Long: lUT_Pos_Left = m_lUT_Pos_Left: End Property
Public Property Let lUT_Pos_Left(ByVal lNewValue As Long): m_lUT_Pos_Left = lNewValue: End Property

Public Property Get lUT_Pos_Top() As Long: lUT_Pos_Top = m_lUT_Pos_Top: End Property
Public Property Let lUT_Pos_Top(ByVal lNewValue As Long): m_lUT_Pos_Top = lNewValue: End Property

Public Property Get lUT_Grp_No() As Long: lUT_Grp_No = m_lUT_Grp_No: End Property
Public Property Let lUT_Grp_No(ByVal lNewValue As Long): m_lUT_Grp_No = lNewValue: End Property
Public Property Get bytUT_IsGrp() As Byte
    bytUT_IsGrp = IIf(lUT_Grp_No, 1, 0)
End Property

Public Property Get lUT_Orig_Grp_No() As Long: lUT_Orig_Grp_No = m_lUT_Orig_Grp_No: End Property
Public Property Let lUT_Orig_Grp_No(ByVal lNewValue As Long): m_lUT_Orig_Grp_No = lNewValue: End Property

Public Property Get lUT_Merge_Width() As Long: lUT_Merge_Width = m_lUT_Merge_Width: End Property
Public Property Let lUT_Merge_Width(ByVal lNewValue As Long): m_lUT_Merge_Width = lNewValue: End Property

Public Property Get lUT_Merge_Height() As Long: lUT_Merge_Height = m_lUT_Merge_Height: End Property
Public Property Let lUT_Merge_Height(ByVal lNewValue As Long): m_lUT_Merge_Height = lNewValue: End Property

Public Property Get lUT_Link_ID() As Long: lUT_Link_ID = m_lUT_Link_ID: End Property
Public Property Let lUT_Link_ID(ByVal lNewValue As Long): m_lUT_Link_ID = lNewValue: End Property

Public Property Get sUT_Name() As String: sUT_Name = m_sUT_Name: End Property
Public Property Let sUT_Name(ByVal sNewValue As String): m_sUT_Name = sNewValue: End Property

Public Property Get sUT_Hdg_Type() As String: sUT_Hdg_Type = m_sUT_Hdg_Type: End Property
Public Property Let sUT_Hdg_Type(ByVal sNewValue As String): m_sUT_Hdg_Type = sNewValue: End Property

Public Property Get sUT_Default_Value() As String: sUT_Default_Value = m_sUT_Default_Value: End Property
Public Property Let sUT_Default_Value(ByVal sNewValue As String): m_sUT_Default_Value = sNewValue: End Property

Public Property Get sUT_Original_Value() As String: sUT_Original_Value = m_sUT_Original_Value: End Property
Public Property Let sUT_Original_Value(ByVal sNewValue As String): m_sUT_Original_Value = sNewValue: End Property

Public Property Get sUT_Wrap_Text() As String: sUT_Wrap_Text = m_sUT_Wrap_Text: End Property
Public Property Let sUT_Wrap_Text(ByVal sNewValue As String): m_sUT_Wrap_Text = sNewValue: End Property

Public Property Get lUT_Font_Size() As Long: lUT_Font_Size = m_lUT_Font_Size: End Property
Public Property Let lUT_Font_Size(ByVal lNewValue As Long): m_lUT_Font_Size = lNewValue: End Property

Public Property Get sUT_Locked() As String: sUT_Locked = m_sUT_Locked: End Property
Public Property Let sUT_Locked(ByVal sNewValue As String): m_sUT_Locked = sNewValue: End Property

Public Property Get lUT_BG_Color() As Long: lUT_BG_Color = m_lUT_BG_Color: End Property
Public Property Let lUT_BG_Color(ByVal lNewValue As Long): m_lUT_BG_Color = lNewValue: End Property

Public Property Get lUT_FG_Color() As Long: lUT_FG_Color = m_lUT_FG_Color: End Property
Public Property Let lUT_FG_Color(ByVal lNewValue As Long): m_lUT_FG_Color = lNewValue: End Property

Public Property Get sUT_Bold() As String: sUT_Bold = m_sUT_Bold: End Property
Public Property Let sUT_Bold(ByVal sNewValue As String): m_sUT_Bold = sNewValue: End Property

Public Property Get sUT_Italic() As String: sUT_Italic = m_sUT_Italic: End Property
Public Property Let sUT_Italic(ByVal sNewValue As String): m_sUT_Italic = sNewValue: End Property

Public Property Get sUT_Underline() As String: sUT_Underline = m_sUT_Underline: End Property
Public Property Let sUT_Underline(ByVal sNewValue As String): m_sUT_Underline = sNewValue: End Property

Public Property Get sUT_StrikeThrough() As String: sUT_StrikeThrough = m_sUT_StrikeThrough: End Property
Public Property Let sUT_StrikeThrough(ByVal sNewValue As String): m_sUT_StrikeThrough = sNewValue: End Property
Public Property Get sUT_BIUS() As String
    sUT_BIUS = sUT_Bold & sUT_Italic & sUT_Underline & sUT_StrikeThrough & sUT_Hdg_Type & sUT_TextAlign & sUT_Wrap_Text & sUT_Locked
End Property

Public Sub LetsUT_BIUS(ByVal sNewValue As String)
    sUT_Bold = Mid(sNewValue, 1, 1)
    sUT_Italic = Mid(sNewValue, 2, 1)
    sUT_Underline = Mid(sNewValue, 3, 1)
    sUT_StrikeThrough = Mid(sNewValue, 4, 1)
    sUT_Hdg_Type = Mid(sNewValue, 5, 1)
    sUT_TextAlign = Mid(sNewValue, 6, 2)
    sUT_Wrap_Text = Mid(sNewValue, 9, 2)
    sUT_Locked = Mid(sNewValue, 10, 2)
End Sub

Public Property Get sUT_TextAlign() As String: sUT_TextAlign = m_sUT_TextAlign: End Property
Public Property Let sUT_TextAlign(ByVal sNewValue As String): m_sUT_TextAlign = sNewValue: End Property

Public Property Get sUT_Field_Type() As String: sUT_Field_Type = m_sUT_Field_Type: End Property
Public Property Let sUT_Field_Type(ByVal sNewValue As String): m_sUT_Field_Type = sNewValue: End Property

Public Property Get sUT_Formula() As String: sUT_Formula = m_sUT_Formula: End Property
Public Property Let sUT_Formula(ByVal sNewValue As String): m_sUT_Formula = sNewValue: End Property

Public Property Get sUT_Procedure() As String: sUT_Procedure = m_sUT_Procedure: End Property
Public Property Let sUT_Procedure(ByVal sNewValue As String): m_sUT_Procedure = sNewValue: End Property

Public Property Get sUT_Data() As String
    sUT_Data = "<v>" & sUT_Default_Value & "</v><f>" & sUT_Field_Type & "</f><p>" & sUT_Procedure & "</p><c>" & sUT_Cond_Format & "</c><h>" & sUT_Hyperlink & _
               "</h><i>" & sUT_Image & "</i><fm>" & sUT_Formula & "</fm>" & "<lt>" & sUT_Line_Type & "</lt" & "<ft>" & sUT_Field_Type & "</ft"
End Property
Public Property Let sUT_Data(ByVal sNewValue As String)
    sUT_Default_Value = g_RtnStrBetween(sNewValue, "<v>", "</v")
    sUT_Field_Type = g_RtnStrBetween(sNewValue, "<f>", "</f")
    sUT_Procedure = g_RtnStrBetween(sNewValue, "<p>", "</p")
    sUT_Cond_Format = g_RtnStrBetween(sNewValue, "<c>", "</c")
    sUT_Hyperlink = g_RtnStrBetween(sNewValue, "<h>", "</h")
    sUT_Image = g_RtnStrBetween(sNewValue, "<i>", "</i")
    sUT_Formula = g_RtnStrBetween(sNewValue, "<fm>", "</fm")
    sUT_Line_Type = g_RtnStrBetween(sNewValue, "<lt>", "</lt")
    sUT_Field_Type = g_RtnStrBetween(sNewValue, "<ft>", "</ft")
End Property

Public Property Get sUT_Border() As String: sUT_Border = m_sUT_Border: End Property
Public Property Let sUT_Border(ByVal sNewValue As String): m_sUT_Border = sNewValue: End Property

Public Property Get sUT_Line_Type() As String: sUT_Line_Type = m_sUT_Line_Type: End Property
Public Property Let sUT_Line_Type(ByVal sNewValue As String): m_sUT_Line_Type = sNewValue: End Property

Public Property Get sUT_Cond_Format() As String: sUT_Cond_Format = m_sUT_Cond_Format: End Property
Public Property Let sUT_Cond_Format(ByVal sNewValue As String): m_sUT_Cond_Format = sNewValue: End Property

Public Property Get sUT_Hyperlink() As String: sUT_Hyperlink = m_sUT_Hyperlink: End Property
Public Property Let sUT_Hyperlink(ByVal sNewValue As String): m_sUT_Hyperlink = sNewValue: End Property

Public Property Get sUT_Image() As String: sUT_Image = m_sUT_Image: End Property
Public Property Let sUT_Image(ByVal sNewValue As String): m_sUT_Image = sNewValue: End Property

Public Function GetsUT_Source(ByVal lNo As Long) As String
    If lNo = 0 Then
        GetsUT_Source = m_sUT_Source
    Else
        GetsUT_Source = Mid(m_sUT_Source & "YYYY", lNo, 1)
    End If
End Function
Private Sub LetsUT_Source(ByVal sNewValue As String)
    m_sUT_Source = sNewValue
End Sub
Public Property Get sUT_Source() As String: sUT_Source = m_sUT_Source: End Property
Private Property Let sUT_Source(ByVal sNewValue As String): m_sUT_Source = sNewValue: End Property

Public Function GetsUT_Addit(ByVal lNo As Long) As String
    Dim aAry() As String
    If lNo = 0 Then
        GetsUT_Addit = m_sUT_Addit
    Else
        aAry = Split(m_sUT_Addit, ",")
        GetsUT_Addit = aAry(lNo - 1)
    End If
End Function
Public Sub LetsUT_Addit(ByVal sNewValue As String, ByVal lNo As Long)
'''    m_sUT_Addit = sNewValue
    Dim aAry() As String
    If lNo = 0 Then
        m_sUT_Addit = sNewValue
    Else
        aAry = Split(m_sUT_Addit, ",")
        aAry(lNo - 1) = Format(sNewValue, "000000")
        m_sUT_Addit = Join(aAry, ",")
    End If
End Sub
Public Property Get sUT_Addit() As String: sUT_Addit = m_sUT_Addit: End Property
Private Property Let sUT_Addit(ByVal sNewValue As String): m_sUT_Addit = sNewValue: End Property

Public Property Get sUT_Upd() As String: sUT_Upd = m_sUT_Upd: End Property
Public Property Let sUT_Upd(ByVal sNewValue As String): m_sUT_Upd = sNewValue: End Property

