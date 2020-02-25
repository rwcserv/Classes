VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTEN2_Levels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit       ' cTEN2_Levels Changed 191104
Private m_lID_ID As Long                                        ' ID of the Idea
Private m_lLV_ID As Long                                        '
Private m_sLV_Portfolio_No As String                            ' The Portfolio number - if blank or 0, no number has yet been applied
Private m_sLV_Portfolio_Desc As String                          ' Description of the Portfolio
Private m_lLV_Sts_ID As Long                                    '
Private m_sLV_Version_No As String                              ' The number of the Version - 1 to ...
Private m_sLV_Version_Desc As String                            ' Description of the Portfolio Version
Private m_sLV_Product_Code As String                            '
Private m_sLV_Prior_Product_Code As String                      ' Prior Product Code (will be able to base the extracted details upon this Product Code, if the actual Prod Code is missing)
Private m_sLV_Contract_No As String                             ' The Contract Number
Private m_sLV_TH_Docs As String                                 ' Tender Doc created NPT(New Product Dev);Retender;MSO(Management Sign-off) - will be in the format of  'NPT-R-MSO'
Private m_sLV_ATPFile As String                                 ' The file name input for the Comparision Report inport (No path kept)
Private m_lLV_TH_ID As Long                                     ' Tender Document (L7_TDoc_Hdrs) Header ID (Report produced for GBD Signoff)
Private m_lLV_AH_ID As Long                                     ' Aldi Tender Platform (L4_ATP_Hdrs) Header ID (ATP Comparison Document)
Private m_sLV_UpdUser As String                                 '
Private m_sLV_UpdDate As String                                 '
Private m_sLV_Upd As String                                     '

Sub Add_Class_LV(ID_ID As Long, LV_ID As Long, LV_Portfolio_No As String, LV_Portfolio_Desc As String, LV_Sts_ID As Long, _
                 LV_Version_No As String, LV_Version_Desc As String, LV_Product_Code As String, LV_Prior_Product_Code As String, LV_Contract_No As String, LV_TH_Docs As String, _
                 LV_ATPFile As String, LV_TH_ID As Long, LV_AH_ID As Long, LV_UpdUser As String, LV_UpdDate As String, ByVal LV_Upd As String)
    ' Add the partlv_rec.ts of the line to the array
    m_lID_ID = ID_ID                                                  ' Header ID
    m_lLV_ID = LV_ID                                                  ' Line Auto-number
    m_lLV_Sts_ID = LV_Sts_ID                                          ' Line Foreign key to the Status table

    m_sLV_Portfolio_No = LV_Portfolio_No                              ' Text fields
    m_sLV_Portfolio_Desc = LV_Portfolio_Desc
    m_sLV_Version_No = LV_Version_No
    m_sLV_Version_Desc = LV_Version_Desc
    m_sLV_Product_Code = LV_Product_Code
    m_sLV_Prior_Product_Code = LV_Prior_Product_Code
    m_sLV_Contract_No = LV_Contract_No
    m_sLV_TH_Docs = LV_TH_Docs
    m_sLV_ATPFile = LV_ATPFile
    m_lLV_TH_ID = LV_TH_ID
    m_lLV_AH_ID = LV_AH_ID
''    If LV_AH_ID > 0 Then
''        LV_AH_ID = LV_AH_ID
''    End If
    m_sLV_UpdDate = LV_UpdDate
    m_sLV_Upd = LV_Upd
    m_sLV_UpdUser = LV_UpdUser

    m_sLV_Upd = LV_Upd    ' Update Flag
End Sub

Private Sub Class_Terminate()
''    On Error Resume Next
''    Call g_EraseAry(pID_Ary)
''    Call g_EraseAry(pLV_Ary)
''    Set psdLV = Nothing
''    Set psdID = Nothing
End Sub


Public Property Get lID_ID() As Long: lID_ID = m_lID_ID: End Property
Public Property Let lID_ID(ByVal lNewValue As Long): m_lID_ID = lNewValue: End Property

Public Property Get lLV_ID() As Long: lLV_ID = m_lLV_ID: End Property
Public Property Let lLV_ID(ByVal lNewValue As Long): m_lLV_ID = lNewValue: End Property

Public Property Get sLV_Portfolio_No() As String: sLV_Portfolio_No = m_sLV_Portfolio_No: End Property
Public Property Let sLV_Portfolio_No(ByVal sNewValue As String): m_sLV_Portfolio_No = sNewValue: End Property

Public Property Get sLV_Portfolio_Desc() As String: sLV_Portfolio_Desc = m_sLV_Portfolio_Desc: End Property
Public Property Let sLV_Portfolio_Desc(ByVal sNewValue As String): m_sLV_Portfolio_Desc = sNewValue: End Property

Public Property Get lLV_Sts_ID() As Long: lLV_Sts_ID = m_lLV_Sts_ID: End Property
Public Property Let lLV_Sts_ID(ByVal lNewValue As Long): m_lLV_Sts_ID = lNewValue: End Property

Public Property Get sLV_Version_No() As String: sLV_Version_No = m_sLV_Version_No: End Property
Public Property Let sLV_Version_No(ByVal sNewValue As String): m_sLV_Version_No = sNewValue: End Property

Public Property Get sLV_Version_Desc() As String: sLV_Version_Desc = m_sLV_Version_Desc: End Property
Public Property Let sLV_Version_Desc(ByVal sNewValue As String): m_sLV_Version_Desc = sNewValue: End Property

Public Property Get sLV_Product_Code() As String: sLV_Product_Code = m_sLV_Product_Code: End Property
Public Property Let sLV_Product_Code(ByVal sNewValue As String): m_sLV_Product_Code = sNewValue: End Property

Public Property Get sLV_Prior_Product_Code() As String: sLV_Prior_Product_Code = m_sLV_Prior_Product_Code: End Property
Public Property Let sLV_Prior_Product_Code(ByVal sNewValue As String): m_sLV_Prior_Product_Code = sNewValue: End Property

Public Property Get sLV_Contract_No() As String: sLV_Contract_No = m_sLV_Contract_No: End Property
Public Property Let sLV_Contract_No(ByVal sNewValue As String): m_sLV_Contract_No = sNewValue: End Property

Public Property Get sLV_TH_Docs() As String: sLV_TH_Docs = m_sLV_TH_Docs: End Property
Public Property Let sLV_TH_Docs(ByVal sNewValue As String): m_sLV_TH_Docs = sNewValue: End Property

Public Property Get sLV_ATPFile() As String: sLV_ATPFile = m_sLV_ATPFile: End Property
Public Property Let sLV_ATPFile(ByVal sNewValue As String): m_sLV_ATPFile = sNewValue: End Property

Public Property Get lLV_TH_ID() As Long: lLV_TH_ID = m_lLV_TH_ID: End Property
Public Property Let lLV_TH_ID(ByVal lNewValue As Long): m_lLV_TH_ID = lNewValue: End Property

Public Property Get lLV_AH_ID() As Long: lLV_AH_ID = m_lLV_AH_ID: End Property
Public Property Let lLV_AH_ID(ByVal lNewValue As Long): m_lLV_AH_ID = lNewValue: End Property

Public Property Get sLV_UpdUser() As String: sLV_UpdUser = m_sLV_UpdUser: End Property
Public Property Let sLV_UpdUser(ByVal sNewValue As String): m_sLV_UpdUser = sNewValue: End Property

Public Property Get sLV_UpdDate() As String: sLV_UpdDate = m_sLV_UpdDate: End Property
Public Property Let sLV_UpdDate(ByVal sNewValue As String): m_sLV_UpdDate = sNewValue: End Property

Public Property Get sLV_Upd() As String: sLV_Upd = m_sLV_Upd: End Property
Public Property Let sLV_Upd(ByVal sNewValue As String): m_sLV_Upd = sNewValue: End Property
