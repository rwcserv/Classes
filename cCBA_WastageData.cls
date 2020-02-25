VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCBA_WastageData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type WData
    pdWdate As Date
    pdblTotalPieces As Double
    pdblTotalRetail As Double
    pdblQTYDiff As Double
    pdblRetailDiff As Double
    plStoreno As Long
    plDivno As Long
End Type
Private pArrWaste() As WData
Private plPcode As Long
Private pbIsActive As Boolean
Private pbis24PackBeer As Boolean
Private pbis6PackBeer As Boolean

Function addData(ByRef RS As ADODB.Recordset, ByVal DivNo As Long, ByVal no As Long) As Boolean
'Dim RS As ADODB.Recordset
Dim cnt As Long, totcnt As Long, a As Long
Dim WD As WData
Dim pnts As Variant
Dim TempArr() As WData

If no = 0 Then Exit Function


    On Error Resume Next
    If UBound(pArrWaste) < 0 Then
    ReDim TempArr(0 To no - 1)
    Else
    ReDim TempArr(0 To UBound(pArrWaste) + no)
        For a = 0 To UBound(pArrWaste)
            TempArr(a).pdWdate = pArrWaste(a).pdWdate
            TempArr(a).plDivno = pArrWaste(a).plDivno
            TempArr(a).plStoreno = pArrWaste(a).plStoreno
            TempArr(a).pdblQTYDiff = pArrWaste(a).pdblQTYDiff
            TempArr(a).pdblRetailDiff = pArrWaste(a).pdblRetailDiff
            TempArr(a).pdblTotalPieces = pArrWaste(a).pdblTotalPieces
            TempArr(a).pdblTotalRetail = pArrWaste(a).pdblTotalRetail
        Next
    End If
    Err.Clear
    On Error GoTo 0
    a = a - 1
    Do Until CLng(RS(1)) <> lPcode
        a = a + 1
        WD.pdWdate = CDate(RS(0))
        WD.plDivno = DivNo
        WD.plStoreno = CLng(RS(6))
        WD.pdblTotalPieces = CDbl(RS(2))
        WD.pdblTotalRetail = CDbl(RS(3))
        WD.pdblQTYDiff = CDbl(RS(4))
        WD.pdblRetailDiff = CDbl(RS(5))
        TempArr(a) = WD
        RS.MoveNext
        If RS.EOF Then Exit Do
    Loop
    pArrWaste = TempArr
    Erase TempArr


End Function
Function getWastageData(ByVal isRetail As e_RetailorQTY, ByVal InventoryMarkdownBoth As e_InvMDALL, Optional ByVal DFrom As Date, Optional Dto As Date, Optional div As Long = 599) As Double '@TP NEED TO FIX UP
Dim Dts As Date
Dim OpVal As Double
Dim a As Long
Dim StoDic As Scripting.Dictionary
    
    On Error Resume Next
    If UBound(pArrWaste) < 0 Then
        Err.Clear: On Error GoTo 0
        getWastageData = 0
    Else
        On Error GoTo 0
        Set StoDic = New Scripting.Dictionary
        For a = LBound(pArrWaste) To UBound(pArrWaste)
            If pArrWaste(a).pdWdate >= DFrom And pArrWaste(a).pdWdate <= Dto And (pArrWaste(a).plDivno = div Or div = 599) Then
                If InventoryMarkdownBoth = e_InvMDALL.eStores Or InventoryMarkdownBoth = e_InvMDALL.eInventoryDifference Or InventoryMarkdownBoth = e_InvMDALL.eBoth Then
                    If (bis6PackBeer = True And pArrWaste(a).pdblQTYDiff < 0) Or (bis24PackBeer = False And bis6PackBeer = False) Then
                        If isRetail = eQTY Then OpVal = OpVal + pArrWaste(a).pdblQTYDiff Else OpVal = OpVal + pArrWaste(a).pdblRetailDiff
                    End If
                End If
                If InventoryMarkdownBoth = e_InvMDALL.eStores Or InventoryMarkdownBoth = e_InvMDALL.eMarkdowns Or InventoryMarkdownBoth = e_InvMDALL.eBoth Then
                    If isRetail = eQTY Then OpVal = OpVal - pArrWaste(a).pdblTotalPieces Else OpVal = OpVal + pArrWaste(a).pdblTotalRetail
                End If
                If InventoryMarkdownBoth = e_InvMDALL.eStores Then
                    If StoDic.Exists(CLng(pArrWaste(a).plStoreno)) = False Then
                        StoDic.Add (CLng(pArrWaste(a).plStoreno)), 1
                    End If
                End If
            End If
        Next
        If InventoryMarkdownBoth = eStores Then If StoDic.Count > 0 Then OpVal = OpVal / StoDic.Count Else OpVal = 0
        getWastageData = OpVal
    End If
End Function
Private Sub checkIsActive()
    If lPcode > 0 Then bIsActive = True Else bIsActive = False
End Sub
Public Property Get lPcode() As Long: lPcode = plPcode: End Property
Public Property Let lPcode(ByVal NewValue As Long): plPcode = NewValue: checkIsActive: End Property
Public Property Get bIsActive() As Boolean: bIsActive = pbIsActive: End Property
Private Property Let bIsActive(ByVal NewValue As Boolean): pbIsActive = NewValue: End Property
Private Property Get bis24PackBeer() As Boolean: bis24PackBeer = pbis24PackBeer: End Property
Public Property Let bis24PackBeer(ByVal NewValue As Boolean): pbis24PackBeer = NewValue: End Property
Private Property Get bis6PackBeer() As Boolean: bis6PackBeer = pbis6PackBeer: End Property
Public Property Let bis6PackBeer(ByVal NewValue As Boolean): pbis6PackBeer = NewValue: End Property
