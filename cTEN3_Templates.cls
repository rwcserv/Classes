VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTEN3_Templates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit       ' cTEN3_Templates Changed 190828
' This class will hold all the Templates for the Excel area covered (Note: all valid Templates are loaded into memory)
Private m_lSH_ID As Long
Private m_lSH_No As Long
Private m_sSH_Desc As String

Public Sub Add_Class_SH(SH_No As Long, SH_ID As Long, SH_Desc As String)
    m_lSH_No = SH_No                                      ' Class Array number
    m_lSH_ID = SH_ID                                      ' Class Auto-number
    m_sSH_Desc = SH_Desc                                  ' Class File Desc
End Sub


''Function Get_Class_SH(sField As String, Optional bAdd As Boolean = False, Optional bSQL As Boolean = False, Optional sTablePrefix As String = "")
''    ' Get a Header value
''    If sField = "SH_Desc" Then
''        Get_Class_SH = m_sSH_Desc
''    ElseIf sField = "SH_No" Then
''        Get_Class_SH = m_lSH_No
''    ElseIf sField = "SH_ID" Then
''        Get_Class_SH = m_lSH_ID
''    Else
''        MsgBox "Field " & sField & " not found"
''    End If
''
''End Function


Private Sub Class_Terminate()
    On Error Resume Next
End Sub

Public Property Get lSH_ID() As Long: lSH_ID = m_lSH_ID: End Property
Public Property Let lSH_ID(ByVal lNewValue As Long): m_lSH_ID = lNewValue: End Property

Public Property Get lSH_No() As Long: lSH_No = m_lSH_No: End Property
Public Property Let lSH_No(ByVal lNewValue As Long): m_lSH_No = lNewValue: End Property

Public Property Get sSH_Desc() As String: sSH_Desc = m_sSH_Desc: End Property
Public Property Let sSH_Desc(ByVal sNewValue As String): m_sSH_Desc = sNewValue: End Property
