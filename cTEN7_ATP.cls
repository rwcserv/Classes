VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTEN7_ATP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit       ' cTEN7_ATP Changed 191112
Private m_lTH_ID As Long
Private m_sTH_Alt    As String
Private m_sTH_Name As String
Private m_sTH_Upd As String

Sub Add_Class_ATP(TH_ID As Long, TH_Alt As String, TH_Name As String, TH_Upd As String)
    m_lTH_ID = TH_ID
    m_sTH_Alt = TH_Alt
    m_sTH_Name = TH_Name
    m_sTH_Upd = TH_Upd
End Sub

Public Property Get TH_ID() As Long: TH_ID = m_lTH_ID: End Property
Public Property Let TH_ID(ByVal lNewValue As Long): m_lTH_ID = lNewValue: End Property

Public Property Get TH_Alt() As String: TH_Alt = m_sTH_Alt: End Property
Public Property Let TH_Alt(ByVal sNewValue As String): m_sTH_Alt = sNewValue: End Property

Public Property Get TH_Name() As String: TH_Name = m_sTH_Name: End Property
Public Property Let TH_Name(ByVal sNewValue As String): m_sTH_Name = sNewValue: End Property

Public Property Get TH_Upd() As String: TH_Upd = m_sTH_Upd: End Property
Public Property Let TH_Upd(ByVal sNewValue As String): m_sTH_Upd = sNewValue: End Property

