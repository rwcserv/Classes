VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTEN8_Lines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit         ' cTEN8_Lines Changed 191104

' This class will hold all the UDT Lines by Row for the tender document
' The reason for this class module is that the Grouped lines need to be flagged as hidden or not
Private m_lTL_Row As Long
Private m_lTL_Height As Long
Private m_sTL_Flag As String            ' Will be "N"otHidden, "H"idden or "L"inesNotShown only
Private m_sTL_Upd As String             ' Will be "N"oUpd or "U"pd only

Sub Add_Class_TL(TL_Row As Long, TL_Height As Long, TL_Flag As String, TL_Upd As String)
    ' Add the parts of the line to the array
    m_lTL_Row = TL_Row
    m_lTL_Height = TL_Height
    m_sTL_Flag = TL_Flag
    m_sTL_Upd = TL_Upd
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
End Sub

Public Property Get lTL_Row() As Long: lTL_Row = m_lTL_Row: End Property
Public Property Let lTL_Row(ByVal lNewValue As Long): m_lTL_Row = lNewValue: End Property

Public Property Get lTL_Height() As Long: lTL_Height = m_lTL_Height: End Property
Public Property Let lTL_Height(ByVal lNewValue As Long): m_lTL_Height = lNewValue: End Property

Public Property Get sTL_Flag() As String: sTL_Flag = m_sTL_Flag: End Property
Public Property Let sTL_Flag(ByVal sNewValue As String): m_sTL_Flag = sNewValue: End Property

Public Property Get sTL_Upd() As String: sTL_Upd = m_sTL_Upd: End Property
Public Property Let sTL_Upd(ByVal sUpdFlag As String): m_sTL_Upd = sUpdFlag: End Property

