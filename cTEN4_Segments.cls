VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTEN4_Segments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit       ' cTEN4_Segments Changed 191104
    ' This class will hold all the Segments for the Excel area covered (Note: all valid Segments are loaded into memory)
Private m_lSL_No As Long
Private m_lSL_ID As Long
Private m_lSL_Seq As Long
Private m_sSL_SH_ID As String
Private m_sSL_Desc As String
Private m_sSL_Mandatory As String
Private m_lSL_Template_No As Long
Private m_sSL_Valid As String

Public Sub Add_Class_SL(SL_No As Long, SL_ID As Long, SL_Seq As Long, SL_Desc As String, SL_SH_ID As String, SL_Mandatory As String, SL_Template_No As Long, SL_Valid As String)
    ' Add the parts of the line
    m_lSL_No = SL_No                                                  ' Line Array Number
    m_lSL_ID = SL_ID                                                  ' Line Auto-number
    m_lSL_Seq = SL_Seq                                                ' Line Foreign key to the TLLine
    m_sSL_SH_ID = SL_SH_ID                                            ' Line Foreign key to the Status table
    m_sSL_Desc = SL_Desc
    m_sSL_Mandatory = SL_Mandatory
    m_lSL_Template_No = SL_Template_No
    m_sSL_Valid = SL_Valid

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
End Sub

Public Property Get lSL_No() As Long: lSL_No = m_lSL_No: End Property
Public Property Let lSL_No(ByVal lNewValue As Long): m_lSL_No = lNewValue: End Property

Public Property Get lSL_ID() As Long: lSL_ID = m_lSL_ID: End Property
Public Property Let lSL_ID(ByVal lNewValue As Long): m_lSL_ID = lNewValue: End Property

Public Property Get lSL_Seq() As Long: lSL_Seq = m_lSL_Seq: End Property
Public Property Let lSL_Seq(ByVal lNewValue As Long): m_lSL_Seq = lNewValue: End Property

Public Property Get sSL_SH_ID() As String: sSL_SH_ID = m_sSL_SH_ID: End Property
Public Property Let sSL_SH_ID(ByVal sNewValue As String): m_sSL_SH_ID = sNewValue: End Property

Public Property Get sSL_Desc() As String: sSL_Desc = m_sSL_Desc: End Property
Public Property Let sSL_Desc(ByVal sNewValue As String): m_sSL_Desc = sNewValue: End Property

Public Property Get sSL_Mandatory() As String: sSL_Mandatory = m_sSL_Mandatory: End Property
Public Property Let sSL_Mandatory(ByVal sNewValue As String): m_sSL_Mandatory = sNewValue: End Property

Public Property Get lSL_Template_No() As Long: lSL_Template_No = m_lSL_Template_No: End Property
Public Property Let lSL_Template_No(ByVal lNewValue As Long): m_lSL_Template_No = lNewValue: End Property

Public Property Get sSL_Valid() As String: sSL_Valid = m_sSL_Valid: End Property
Public Property Let sSL_Valid(ByVal sNewValue As String): m_sSL_Valid = sNewValue: End Property
