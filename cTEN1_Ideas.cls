VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTEN1_Ideas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit       ' cTEN1_Ideas Changed 191104
Private m_lID_ID As Long
Private m_sID_Desc As String
Private m_lID_Sts_ID As Long
Private m_sSts_Desc As String
Private m_sID_BA As String
Private m_sID_BD As String
Private m_lID_BA_Emp_No As Long
Private m_lID_BD_Emp_No As Long
Private m_lID_GD_Emp_No As Long
Private m_sID_UpdDate As String
Private m_sID_UpdUser As String
Private m_sID_CrtUser As String
Private m_sID_Upd As String

Sub Add_Class_ID(ID_ID As Long, ID_Desc As String, ID_Sts_ID As Long, Sts_Desc As String, ID_BA As String, ID_BD As String, ID_BA_Emp_No As Long, ID_BD_Emp_No As Long, ID_GD_Emp_No As Long, ID_UpdDate As String, ID_UpdUser As String, ID_CrtUser As String, ByVal ID_Upd As String)
    m_lID_ID = ID_ID                                      ' Class Auto-number - (If 0, hasn't been created yet)
    m_sID_Desc = ID_Desc                                  ' Class File Desc
    m_lID_Sts_ID = ID_Sts_ID                              ' Pointer to the Idea Status
    m_sSts_Desc = Sts_Desc
    m_sID_BA = ID_BA
    m_sID_BD = ID_BD
    m_lID_BA_Emp_No = ID_BA_Emp_No
    m_lID_BD_Emp_No = ID_BD_Emp_No
    m_lID_GD_Emp_No = ID_GD_Emp_No
    m_sID_UpdDate = ID_UpdDate
    m_sID_UpdUser = ID_UpdUser
    m_sID_CrtUser = ID_CrtUser
    m_sID_Upd = ID_Upd
End Sub

Public Property Get lID_ID() As Long: lID_ID = m_lID_ID: End Property
Public Property Let lID_ID(ByVal lNewValue As Long): m_lID_ID = lNewValue: End Property

Public Property Get sID_Desc() As String: sID_Desc = m_sID_Desc: End Property
Public Property Let sID_Desc(ByVal sNewValue As String): m_sID_Desc = sNewValue: End Property

Public Property Get lID_Sts_ID() As Long: lID_Sts_ID = m_lID_Sts_ID: End Property
Public Property Let lID_Sts_ID(ByVal lNewValue As Long): m_lID_Sts_ID = lNewValue: End Property

Public Property Get sSts_Desc() As String: sSts_Desc = m_sSts_Desc: End Property
Public Property Let sSts_Desc(ByVal sNewValue As String): m_sSts_Desc = sNewValue: End Property

Public Property Get sID_BA() As String: sID_BA = m_sID_BA: End Property
Public Property Let sID_BA(ByVal sNewValue As String): m_sID_BA = sNewValue: End Property

Public Property Get sID_BD() As String: sID_BD = m_sID_BD: End Property
Public Property Let sID_BD(ByVal sNewValue As String): m_sID_BD = sNewValue: End Property

Public Property Get lID_BA_Emp_No() As Long: lID_BA_Emp_No = m_lID_BA_Emp_No: End Property
Public Property Let lID_BA_Emp_No(ByVal lNewValue As Long): m_lID_BA_Emp_No = lNewValue: End Property

Public Property Get lID_BD_Emp_No() As Long: lID_BD_Emp_No = m_lID_BD_Emp_No: End Property
Public Property Let lID_BD_Emp_No(ByVal lNewValue As Long): m_lID_BD_Emp_No = lNewValue: End Property

Public Property Get lID_GD_Emp_No() As Long: lID_GD_Emp_No = m_lID_GD_Emp_No: End Property
Public Property Let lID_GD_Emp_No(ByVal lNewValue As Long): m_lID_GD_Emp_No = lNewValue: End Property

Public Property Get sID_UpdDate() As String: sID_UpdDate = m_sID_UpdDate: End Property
Public Property Let sID_UpdDate(ByVal sNewValue As String): m_sID_UpdDate = sNewValue: End Property

Public Property Get sID_UpdUser() As String: sID_UpdUser = m_sID_UpdUser: End Property
Public Property Let sID_UpdUser(ByVal sNewValue As String): m_sID_UpdUser = sNewValue: End Property

Public Property Get sID_CrtUser() As String: sID_CrtUser = m_sID_CrtUser: End Property
Public Property Let sID_CrtUser(ByVal sNewValue As String): m_sID_CrtUser = sNewValue: End Property

Public Property Get sID_Upd() As String: sID_Upd = m_sID_Upd: End Property
Public Property Let sID_Upd(ByVal sNewValue As String): m_sID_Upd = sNewValue: End Property
