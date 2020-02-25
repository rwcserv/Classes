VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_PfVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sPPortfolioID As String, m_sPPfVersionID As String
Private ContDic As Scripting.Dictionary
Private ProdDic As Scripting.Dictionary
Private m_LCGCGno As Byte
Private m_LCGSCGno As Byte
Private m_ACGCGno As Byte
Private m_ACGSCGno As Byte
Private m_ACGCatNo As Byte
Private m_SeasonID As Byte
Private m_ProductClass As Byte
Private m_sBDFName As String
Private m_sBDName As String
Private m_sBAFName As String
Private m_sBAName As String
Private m_sACGCategory As String
Private m_sACGCommodityGroup As String
Private m_sACGSubCommodityGroup As String
Private m_sLCGCommodityGroup As String
Private m_sLCGSubCommodityGroup As String
Private m_iPackSize As Integer
Private m_CatNo As Byte
Private m_sDescription As String
Private m_sCompPCode As String

Private Type PfVersionTypeData
    productcode As String
    ContractNo As String
    OSD As Date
    ACGEntityID As Integer
    Cost1 As Single
    Retail1 As Single
    NetRetail1 As Single
    Quantity1 As Single
    LPfStatusID As Integer
    LSpecialCatergoryno As Long
    DeliveryFrom As Date
    DeliveryTo As Date
    lTaxID As Byte
    Packunit As String
    MinOrderQty As Single
    Cost As Single
    CostCurrency As String
    Retail As Single
    Quantity As Single
    SuggestedQty As Single
    Leadtime As String
    CostPickup As Single
    PriceBase As Byte
    LRegionID As Single
    CurrencyCode As String
    Comp1PortfolioID As String
    Comp1PfVersionID As String
    Comp2PortfolioID As String
    Comp2PfVersionID As String
End Type
Private PfVersionData() As PfVersionTypeData

Function Generate(ByRef arr As Variant, ByVal PortfolioID As String, ByVal pfVersionID As Byte) As Boolean
    Dim cColl As Collection
    Dim curPCode As String
    Dim DummyContractNo As Long
    Dim lNoOfContracts As Long
    Dim bfound As Boolean, a As Long
    On Error GoTo Err_Routine
    CBA_Error = ""
    
    lNoOfContracts = -1
    m_sPPortfolioID = PortfolioID
    m_sPPfVersionID = pfVersionID
    Set ContDic = New Scripting.Dictionary
    Set ProdDic = New Scripting.Dictionary
    For a = LBound(arr, 2) To UBound(arr, 2)
        If bfound = True And arr(1, a) > pfVersionID Then Exit For
        If arr(0, a) = PortfolioID And arr(1, a) = pfVersionID Then
            bfound = True
            If curPCode = CStr(NZ(arr(3, a), "0")) And curPCode <> "" Then
                cColl.Add CStr(NZ(arr(4, a), ""))
            Else
                If curPCode <> "" Then
                    If ProdDic.Exists(curPCode) Then
                        ProdDic(curPCode).Add CStr(NZ(arr(4, a), ""))
                    Else
                        ProdDic.Add CStr(curPCode), cColl
                    End If
                End If
                Set cColl = New Collection
                curPCode = CStr(NZ(arr(3, a), "0"))
                cColl.Add CStr(NZ(arr(4, a), ""))
            End If
            lNoOfContracts = lNoOfContracts + 1
            ReDim Preserve PfVersionData(0 To lNoOfContracts)
            'Set pPD = New CBA_PfVersionContract
            If lNoOfContracts = 0 Then
                If m_sDescription = "" Then m_sDescription = NZ(arr(2, a), "")
                If m_LCGCGno = 0 Then m_LCGCGno = NZ(arr(6, a), 0)
                If m_LCGSCGno = 0 Then m_LCGSCGno = NZ(arr(7, a), 0)
                'PfVersionData(lNoOfContracts).ACGEntityID = NZ(RS(8), 0)
                If m_CatNo = 0 Then m_CatNo = NZ(arr(9, a), 0)
                If m_sACGCategory = "" Then m_sACGCategory = NZ(arr(10, a), "")
                If m_ACGCGno = 0 Then m_ACGCGno = NZ(arr(11, a), 0)
                If m_sACGCommodityGroup = "" Then m_sACGCommodityGroup = NZ(arr(12, a), "")
                If m_ACGSCGno = 0 Then m_ACGSCGno = NZ(arr(13, a), 0)
                If m_sACGSubCommodityGroup = "" Then m_sACGSubCommodityGroup = NZ(arr(14, a), "")
                If m_SeasonID = 0 Then m_SeasonID = NZ(arr(15, a), 0)
                If m_ProductClass = 0 Then m_ProductClass = NZ(arr(16, a), 0)
                If m_iPackSize = 0 Then m_iPackSize = NZ(arr(17, a), 0)
                If m_sBDName = "" Then m_sBDName = NZ(arr(23, a), "")
                If m_sBDFName = "" Then m_sBDFName = NZ(arr(24, a), "")
                If m_sBAName = "" Then m_sBAName = NZ(arr(25, a), "")
                If m_sBAFName = "" Then m_sBAFName = NZ(arr(26, a), "")
                If m_sLCGCommodityGroup = "" Then m_sLCGCommodityGroup = NZ(arr(47, a), "")
                If m_sLCGSubCommodityGroup = "" Then m_sLCGSubCommodityGroup = NZ(arr(48, a), "")
                If m_sCompPCode = "" Then m_sCompPCode = NZ(arr(49, a), "")
            End If
            PfVersionData(lNoOfContracts).productcode = NZ(arr(3, a), "0")
            PfVersionData(lNoOfContracts).ContractNo = NZ(arr(4, a), "")
            PfVersionData(lNoOfContracts).OSD = NZ(arr(5, a), 0)
            PfVersionData(lNoOfContracts).Cost1 = NZ(arr(18, a), 0)
            PfVersionData(lNoOfContracts).Retail1 = NZ(arr(19, a), 0)
            PfVersionData(lNoOfContracts).NetRetail1 = NZ(arr(20, a), 0)
            PfVersionData(lNoOfContracts).Quantity1 = NZ(arr(21, a), 0)
            PfVersionData(lNoOfContracts).LPfStatusID = NZ(arr(22, a), 0)
            PfVersionData(lNoOfContracts).LSpecialCatergoryno = NZ(arr(27, a), 0)
            PfVersionData(lNoOfContracts).DeliveryFrom = NZ(arr(28, a), 0)
            PfVersionData(lNoOfContracts).DeliveryTo = NZ(arr(29, a), 0)
            PfVersionData(lNoOfContracts).lTaxID = NZ(arr(30, a), 0)
            PfVersionData(lNoOfContracts).Packunit = NZ(arr(31, a), "")
            PfVersionData(lNoOfContracts).MinOrderQty = NZ(arr(32, a), 0)
            PfVersionData(lNoOfContracts).Cost = NZ(arr(33, a), 0)
            PfVersionData(lNoOfContracts).CostCurrency = NZ(arr(34, a), 0)
            PfVersionData(lNoOfContracts).Retail = NZ(arr(35, a), 0)
            PfVersionData(lNoOfContracts).Quantity = NZ(arr(36, a), 0)
            PfVersionData(lNoOfContracts).SuggestedQty = NZ(arr(37, a), 0)
            PfVersionData(lNoOfContracts).Leadtime = NZ(arr(38, a), 0)
            PfVersionData(lNoOfContracts).CostPickup = NZ(arr(39, a), 0)
            PfVersionData(lNoOfContracts).PriceBase = NZ(arr(40, a), 0)
            PfVersionData(lNoOfContracts).LRegionID = NZ(arr(41, a), 0)
            PfVersionData(lNoOfContracts).CurrencyCode = NZ(arr(42, a), "")
            PfVersionData(lNoOfContracts).Comp1PortfolioID = NZ(arr(43, a), "")
            PfVersionData(lNoOfContracts).Comp1PfVersionID = NZ(arr(44, a), "")
            PfVersionData(lNoOfContracts).Comp2PortfolioID = NZ(arr(45, a), "")
            PfVersionData(lNoOfContracts).Comp2PfVersionID = NZ(arr(46, a), "")
            If NZ(arr(4, a), "") = "" Then
                DummyContractNo = DummyContractNo + 1
                ContDic.Add CStr(DummyContractNo), CStr(lNoOfContracts)
            Else
                If ContDic.Exists(CStr(NZ(arr(4, a), ""))) Then
                    ContDic(CStr(NZ(arr(4, a), ""))) = CStr(lNoOfContracts)
                Else
                    ContDic.Add CStr(NZ(arr(4, a), "")), CStr(lNoOfContracts)
                End If
            End If
        End If
    Next
    If ProdDic.Exists(curPCode) = False Then ProdDic.Add curPCode, cColl
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_PfVersion.Generate", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Property Get PortfolioID() As String
    PortfolioID = m_sPPortfolioID
End Property
Public Property Let PPortfolioID(ByVal sNewValue As String)
    m_sPPortfolioID = sNewValue
End Property
Public Property Get pfVersionID() As String
    pfVersionID = m_sPPfVersionID
End Property
Public Property Let PPfVersionID(ByVal sNewValue As String)
    m_sPPfVersionID = sNewValue
End Property
Private Function getSpecificDataFromProdDic(ByVal FieldRqd As String, ByVal ContNoOrPCode As Long) As Scripting.Dictionary
Dim dic As Scripting.Dictionary
Dim i As Variant
Dim col As Collection
Set dic = New Scripting.Dictionary
    If ContNoOrPCode = 0 Then Set getSpecificDataFromProdDic = getSpecificDataFromContDic(FieldRqd, 0): Exit Function
    
    Set col = ProdDic(Format(ContNoOrPCode, "000000"))
    For Each i In ProdDic(Format(ContNoOrPCode, "000000"))
        Select Case FieldRqd
            Case "Cost"
                dic.Add ProdDic(CStr(ContNoOrPCode))(i), PfVersionData(ContDic(ProdDic(CStr(ContNoOrPCode))(i))).Cost1
            Case "Retail"
                dic.Add ProdDic(CStr(ContNoOrPCode))(i), PfVersionData(ContDic(ProdDic(CStr(ContNoOrPCode))(i))).Retail1
            Case "NetRetail"
                dic.Add ProdDic(CStr(ContNoOrPCode))(i), PfVersionData(ContDic(ProdDic(CStr(ContNoOrPCode))(i))).NetRetail1
            Case "Quantity"
                dic.Add ProdDic(CStr(ContNoOrPCode))(i), PfVersionData(ContDic(ProdDic(CStr(ContNoOrPCode))(i))).Quantity1
            Case "Currency"
                dic.Add ProdDic(CStr(ContNoOrPCode))(i), PfVersionData(ContDic(ProdDic(CStr(ContNoOrPCode))(i))).CostCurrency
        End Select
    Next
    Set getSpecificDataFromProdDic = dic
End Function
Private Function getSpecificDataFromContDic(ByVal FieldRqd As String, ByVal ContNoOrPCode As Long) As Scripting.Dictionary
Dim dic As Scripting.Dictionary
Dim i As Variant
Set dic = New Scripting.Dictionary
    If ContNoOrPCode = 0 Then
        For Each i In ContDic
            Select Case FieldRqd
                Case "Cost"
                    dic.Add i, PfVersionData(ContDic(i)).Cost1
                Case "Retail"
                    dic.Add i, PfVersionData(ContDic(i)).Retail1
                Case "NetRetail"
                    dic.Add i, PfVersionData(ContDic(i)).NetRetail1
                Case "Quantity"
                    dic.Add i, PfVersionData(ContDic(i)).Quantity1
                Case "Currency"
                    dic.Add i, PfVersionData(ContDic(i)).CostCurrency
            End Select
        Next
    Else
        Select Case FieldRqd
            Case "Cost"
                dic.Add CStr(Format(ContNoOrPCode, "000000")), PfVersionData(ContDic(CStr(Format(ContNoOrPCode, "000000")))).Cost1
            Case "Retail"
                dic.Add CStr(Format(ContNoOrPCode, "000000")), PfVersionData(ContDic(CStr(Format(ContNoOrPCode, "000000")))).Retail1
            Case "NetRetail"
                dic.Add CStr(Format(ContNoOrPCode, "000000")), PfVersionData(ContDic(CStr(Format(ContNoOrPCode, "000000")))).NetRetail1
            Case "Quantity"
                dic.Add CStr(Format(ContNoOrPCode, "000000")), PfVersionData(ContDic(CStr(Format(ContNoOrPCode, "000000")))).Quantity1
            Case "Currency"
                dic.Add CStr(Format(ContNoOrPCode, "000000")), PfVersionData(ContDic(CStr(Format(ContNoOrPCode, "000000")))).CostCurrency
        End Select
    End If
    Set getSpecificDataFromContDic = dic
End Function
Function getCost(Optional ByVal ContNoOrPCode As Long, Optional ByVal isContractnoNotProductCode As Boolean = True) As Scripting.Dictionary
    If isContractnoNotProductCode = False Then Set getCost = getSpecificDataFromProdDic("Cost", ContNoOrPCode) Else Set getCost = getSpecificDataFromContDic("Cost", ContNoOrPCode)
End Function
Function getRetail(Optional ByVal ContNoOrPCode As Long, Optional ByVal isContractnoNotProductCode As Boolean = True) As Scripting.Dictionary
    If isContractnoNotProductCode = False Then Set getRetail = getSpecificDataFromProdDic("Retail", ContNoOrPCode) Else Set getRetail = getSpecificDataFromContDic("Retail", ContNoOrPCode)
End Function
Function getNetRetail(Optional ByVal ContNoOrPCode As Long, Optional ByVal isContractnoNotProductCode As Boolean = True) As Scripting.Dictionary
    If isContractnoNotProductCode = False Then Set getNetRetail = getSpecificDataFromProdDic("NetRetail", ContNoOrPCode) Else Set getNetRetail = getSpecificDataFromContDic("NetRetail", ContNoOrPCode)
End Function
Function getMargin(Optional ByVal ContNoOrPCode As Long, Optional ByVal isContractnoNotProductCode As Boolean = True) As Scripting.Dictionary
    Dim Cost As Scripting.Dictionary
    Dim RetNet As Scripting.Dictionary
    Dim Ret As Scripting.Dictionary
    Dim dic As Scripting.Dictionary
    Dim rObj As Object
    Set dic = New Scripting.Dictionary
    Set Cost = New Scripting.Dictionary
    Set RetNet = New Scripting.Dictionary
    Set Ret = New Scripting.Dictionary
    If isContractnoNotProductCode = False Then
        Set Cost = getSpecificDataFromProdDic("Cost", ContNoOrPCode)
        Set RetNet = getSpecificDataFromProdDic("NetRetail", ContNoOrPCode)
        Set Ret = getSpecificDataFromProdDic("Retail", ContNoOrPCode)
    Else
        Set Cost = getSpecificDataFromContDic("Cost", ContNoOrPCode)
        Set RetNet = getSpecificDataFromContDic("NetRetail", ContNoOrPCode)
        Set Ret = getSpecificDataFromContDic("Retail", ContNoOrPCode)
    End If

    For Each rObj In RetNet
        If Ret(rObj) > 0 Then dic.Add rObj, (RetNet(rObj) - Cost(rObj)) / Ret(rObj)
    Next
    Set getMargin = dic
End Function
Function getQuantity(Optional ByVal ContNoOrPCode As Long, Optional ByVal isContractnoNotProductCode As Boolean = True) As Scripting.Dictionary
    If isContractnoNotProductCode = False Then Set getQuantity = getSpecificDataFromProdDic("Quantity", ContNoOrPCode) Else Set getQuantity = getSpecificDataFromContDic("Quantity", ContNoOrPCode)
End Function
Function getCurrency(Optional ByVal ContNoOrPCode As Long, Optional ByVal isContractnoNotProductCode As Boolean = True) As Scripting.Dictionary
    If isContractnoNotProductCode = False Then Set getCurrency = getSpecificDataFromProdDic("Currency", ContNoOrPCode) Else Set getCurrency = getSpecificDataFromContDic("Currency", ContNoOrPCode)
End Function
Public Property Get LCGCGno() As Byte: LCGCGno = m_LCGCGno: End Property
Public Property Get LCGSCGno() As Byte: LCGSCGno = m_LCGSCGno: End Property
Public Property Get ACGCGno() As Byte: ACGCGno = m_ACGCGno: End Property
Public Property Get ACGSCGno() As Byte: ACGSCGno = m_ACGSCGno: End Property
Public Property Get ACGCatNo() As Byte: ACGCatNo = m_ACGCatNo: End Property
Public Property Get Seasonid() As Byte: Seasonid = m_SeasonID: End Property
Public Property Get BDFName() As String: BDFName = m_sBDFName: End Property
Public Property Get BDName() As String: BDName = m_sBDName: End Property
Public Property Get BAFName() As String: BAFName = m_sBAFName: End Property
Public Property Get BAName() As String: BAName = m_sBAName: End Property
Public Property Get ACGCategory() As String: ACGCategory = m_sACGCategory: End Property
Public Property Get ACGCommodityGroup() As String: ACGCommodityGroup = m_sACGCommodityGroup: End Property
Public Property Get Packsize() As Integer: Packsize = m_iPackSize: End Property
Public Property Get CatNo() As Byte: CatNo = m_CatNo: End Property
Public Property Get Description() As String: Description = m_sDescription: End Property
Public Property Get LCGCommodityGroup() As String: LCGCommodityGroup = m_sLCGCommodityGroup: End Property
Public Property Get LCGSubCommodityGroup() As String: LCGSubCommodityGroup = m_sLCGSubCommodityGroup: End Property

Public Property Get CompPCode() As String: CompPCode = m_sCompPCode: End Property
