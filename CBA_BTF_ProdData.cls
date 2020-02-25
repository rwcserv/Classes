VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_BTF_ProdData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit       ' CBA_BTF_ProdData @CBA_BTF Changed 181209
Private pProductCode As Long
Private pYear As String
Private pMonth As String
Private pPOSQTY As Long
Private pForeQTY As Long
Private pPOSPYQTY As Long
Private pPOSYOTQTY As Long
Private pPOSRET As Single
Private pForeAvgPrice As Single
Private pPOSYOYRET As Single
Private pRCVMargin As Single
Private pCPU As Single
Private pForeCPU As Single
Private pCPUPY As Single
Private pCPUYOY As Single
Private PDesc As String
Private pPOSPYRET As Single
Private pForeRET As Single
Private pProdClass As Long

Sub Generate(ByVal productcode As Long, ByVal Year As Long, ByVal Month As String, _
    ByVal POSQTY As Long, ByVal POSPYQTY As Long, ByVal POSYOTQTY As Long, ByVal POSRET As Single, _
    ByVal POSYOYRET As Single, ByVal RCVMargin As Single, ByVal CPU As Single, ByVal CPUYOY As Single, _
    ByVal POSPYRET As Single, ByVal CPUPY As Single, ByVal Description As String, ByVal ProdClass As Long)

    pProductCode = productcode
    pYear = Year
    pMonth = Month
    pPOSQTY = POSQTY
    pPOSPYQTY = POSPYQTY
    pPOSYOTQTY = POSYOTQTY
    pPOSRET = POSRET
    pPOSPYRET = POSPYRET
    pPOSYOYRET = POSYOYRET
    pRCVMargin = RCVMargin
    pCPU = CPU
    pCPUYOY = CPUYOY
    PDesc = Description
    pCPUPY = CPUPY
    pProdClass = ProdClass
    ''default forecast values
    pForeCPU = pCPU
    If pPOSQTY = 0 Then pForeAvgPrice = 0 Else pForeAvgPrice = pPOSRET / pPOSQTY
    pForeQTY = pPOSQTY
    pForeRET = pPOSRET

End Sub
Property Get productcode() As Long
    productcode = pProductCode
End Property
Property Get Year() As Long
    Year = pYear
End Property
Property Get Month() As String
    Month = pMonth
End Property
Property Get ProdClass() As String
    ProdClass = pProdClass
End Property
Property Get POSQTY() As Long
    POSQTY = pPOSQTY
End Property
Property Get POSPYQTY() As Long
    POSPYQTY = pPOSPYQTY
End Property
Property Get POSYOTQTY() As Single
    POSYOTQTY = pPOSYOTQTY
End Property
Property Get POSRET() As Single
    POSRET = pPOSRET
End Property
Property Get POSYOYRET() As Single
    POSYOYRET = pPOSYOYRET
End Property
Property Get RCVMargin() As Single
    RCVMargin = pRCVMargin
End Property
Property Get CPU() As Single
    CPU = pCPU
End Property
Property Get CPUYOY() As Single
    CPUYOY = pCPUYOY
End Property
Property Get Description() As String
    Description = PDesc
End Property
Property Get POSPYRET() As String
    POSPYRET = pPOSPYRET
End Property
Property Get CPUPY() As Single
    CPUPY = pCPUPY
End Property
Property Get ForeCPU() As Single
    ForeCPU = pForeCPU
End Property
Property Get ForeAvgPrice() As Single
    ForeAvgPrice = pForeAvgPrice
End Property
Property Get ForeQTY() As Long
    ForeQTY = pForeQTY
End Property
Property Get ForeRET() As Long
    ForeRET = pForeRET
End Property
Sub UpdateForecast(ByRef obj As Control)
    Dim Mno As Long
    If Mid(obj.Name, 1, 1) = "M" Then
        If Mid(obj.Name, 3, 1) = "_" Then Mno = Mid(obj.Name, 2, 1) Else Mno = Mid(obj.Name, 2, 2)
        If InStr(1, LCase(obj.Name), "_avgprice") > 0 Then
            pForeAvgPrice = obj.Value
        ElseIf InStr(1, LCase(obj.Name), "_avgunitcost") > 0 Then
            pForeCPU = obj.Value
        ElseIf InStr(1, LCase(obj.Name), "_fposqty") > 0 Then
            pForeQTY = obj.Value
        ElseIf InStr(1, LCase(obj.Name), "_fposret") > 0 Then
            pForeRET = obj.Value
        End If
    End If
End Sub

Sub UpdateValue(sField As String, ByVal NewValue)
    ' Will update the values into the array
    If sField = "Uplift" Then
        pForeRET = Round(pPOSRET * (1 + (NewValue / 100)), 4)
        pForeQTY = Round(pPOSQTY * (1 + (NewValue / 100)), 4)
    ElseIf sField = "ForeAvgPrice" Then
        pForeAvgPrice = NewValue
    ElseIf sField = "ForeCPU" Then
        pForeCPU = NewValue
    ElseIf sField = "ForeQTY" Then
        pForeQTY = NewValue
    ElseIf sField = "ForeRET" Then
        pForeRET = NewValue
    End If
End Sub

