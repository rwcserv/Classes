VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_COM_COMCompSKU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private CBA_COM_SKU_Ccode As String
Private CBA_COM_SKU_CCodeName  As String
Private CBA_COM_SKU_CPack As String
Private CBA_COM_SKU_Comp As String
Private LastSeen As Date
Private Type CBA_COM_SKUData
    Pricecol As Collection
    Highest As Single
    Lowest As Single
    Median As Single
    Mode As Single
    Mean As Single
    MeanCount As Long
    MetaData As String
    MultiBuy As Single
End Type
Private Recent(1 To 6) As CBA_COM_SKUData
Private Promo(1 To 6) As CBA_COM_SKUData
Private NonPromo(1 To 6) As CBA_COM_SKUData
Private MultiBuy(1 To 6) As CBA_COM_SKUData
Private ALL(1 To 6) As CBA_COM_SKUData
''Private CBA_COM_SKUCOMRS As ADODB.Recordset
''Private CBA_COM_SKUCBISRS As ADODB.Recordset
Public Sub CBA_COM_SKUFormulate(ByVal compet As String, ByVal CCode As String, ByVal Cname As String, ByVal cPack As String, ByVal LastSeen As Date)
    Dim a As Long
    Dim vDiv As Single, divpromo As Single
    Dim rec As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    CBA_COM_SKU_Comp = compet
    CBA_COM_SKU_Ccode = CCode
    CBA_COM_SKU_CCodeName = Cname
    CBA_COM_SKU_CPack = cPack
    
If CBA_COM_SKU_COMRS.State = 0 Then Exit Sub
    
    For a = 1 To 6
        Recent(a).Lowest = 999999
        Recent(a).Highest = 0
        Set Recent(a).Pricecol = New Collection
        Promo(a).Lowest = 999999
        Promo(a).Highest = 0
        Set Promo(a).Pricecol = New Collection
        NonPromo(a).Lowest = 999999
        NonPromo(a).Highest = 0
        Set NonPromo(a).Pricecol = New Collection
        MultiBuy(a).Lowest = 999999
        MultiBuy(a).Highest = 0
        Set MultiBuy(a).Pricecol = New Collection
        ALL(a).Lowest = 999999
        ALL(a).Highest = 0
        Set ALL(a).Pricecol = New Collection
    Next

    With CBA_COM_SKU_COMRS
        Do Until .EOF
            If .Fields(0) = CCode Then
                If IsNull(.Fields(4)) Then
                    If .Fields(1) = LastSeen Then rec = True Else rec = False
                    AllocatePricing "6", .Fields(3), .Fields(2), rec, NZ(.Fields(5), 0)
                Else
                    If .Fields(1) = LastSeen Then rec = True Else rec = False
'                    If InStr(1, .Fields(3), "Promo") > 0 Then
'                    a = a
'                    End If
                    
                    
                    AllocatePricing .Fields(4), .Fields(3), .Fields(2), rec, NZ(.Fields(5), 0)
                End If
            ElseIf .Fields(0) <> CCode Then
                Exit Do
            End If
            .MoveNext
        Loop

        For a = 1 To 6
            If Promo(a).MeanCount > 0 Then Promo(a).Mean = Promo(a).Mean / Promo(a).MeanCount
            If NonPromo(a).MeanCount > 0 Then NonPromo(a).Mean = NonPromo(a).Mean / NonPromo(a).MeanCount
            If Recent(a).MeanCount > 0 Then Recent(a).Mean = Recent(a).Mean / Recent(a).MeanCount
            If MultiBuy(a).MeanCount > 0 Then MultiBuy(a).Mean = MultiBuy(a).Mean / MultiBuy(a).MeanCount
            If ALL(a).MeanCount > 0 Then ALL(a).Mean = ALL(a).Mean / ALL(a).MeanCount
            NonPromo(a).Mode = Mode(NonPromo(a).Pricecol)
            Promo(a).Mode = Mode(Promo(a).Pricecol)
            Recent(a).Mode = Mode(Recent(a).Pricecol)
            MultiBuy(a).Mode = Mode(MultiBuy(a).Pricecol)
            ALL(a).Mode = Mode(ALL(a).Pricecol)
            vDiv = (NonPromo(a).Pricecol.Count + 1) / 2
            If vDiv = Round((NonPromo(a).Pricecol.Count + 1) / 2, 0) And vDiv > 1 Then
                NonPromo(a).Median = (NonPromo(a).Pricecol(vDiv) + NonPromo(a).Pricecol(vDiv + 1)) / 2
            Else
                If vDiv = 0.5 Then
                    NonPromo(a).Median = 0
                Else
                    On Error Resume Next
                    If vDiv = 1 Then
                        NonPromo(a).Median = NonPromo(a).Pricecol(vDiv)
                    Else
                        NonPromo(a).Median = NonPromo(a).Pricecol(vDiv + 0.5)
                    End If
                    On Error GoTo Err_Routine
                End If
            End If
            divpromo = (Promo(a).Pricecol.Count + 1) / 2
            If divpromo = Round((Promo(a).Pricecol.Count + 1) / 2, 0) And divpromo > 1 Then
                Promo(a).Median = (Promo(a).Pricecol(divpromo) + Promo(a).Pricecol(divpromo + 1)) / 2
            Else
                If divpromo = 0.5 Then
                    Promo(a).Median = 0
                Else
                    On Error Resume Next
                    If divpromo = 1 Then
                        NonPromo(a).Median = NonPromo(a).Pricecol(divpromo)
                        Promo(a).Median = Promo(a).Pricecol(divpromo)
                    Else
                        NonPromo(a).Median = NonPromo(a).Pricecol(divpromo + 0.5)
                        Promo(a).Median = Promo(a).Pricecol(divpromo + 0.5)
                    End If
                    On Error GoTo Err_Routine
                End If
            End If
            vDiv = (Recent(a).Pricecol.Count + 1) / 2
            If vDiv = Round((Recent(a).Pricecol.Count + 1) / 2, 0) And vDiv > 1 Then
                Recent(a).Median = (Recent(a).Pricecol(vDiv) + Recent(a).Pricecol(vDiv + 1)) / 2
            Else
                If vDiv = 0.5 Then
                    Recent(a).Median = 0
                Else
                    On Error Resume Next
                    If vDiv = 1 Then
                        Recent(a).Median = Recent(a).Pricecol(vDiv)
                    Else
                        Recent(a).Median = Recent(a).Pricecol(vDiv + 0.5)
                    End If
                    On Error GoTo Err_Routine
                End If
            End If
            vDiv = (MultiBuy(a).Pricecol.Count + 1) / 2
            If vDiv = Round((MultiBuy(a).Pricecol.Count + 1) / 2, 0) And vDiv > 1 Then
                MultiBuy(a).Median = (MultiBuy(a).Pricecol(vDiv) + MultiBuy(a).Pricecol(vDiv + 1)) / 2
            Else
                If vDiv = 0.5 Then
                    MultiBuy(a).Median = 0
                Else
                    On Error Resume Next
                    If vDiv = 1 Then
                        MultiBuy(a).Median = MultiBuy(a).Pricecol(vDiv)
                    Else
                        MultiBuy(a).Median = MultiBuy(a).Pricecol(vDiv + 0.5)
                    End If
                    On Error GoTo Err_Routine
                End If
            End If
            vDiv = (ALL(a).Pricecol.Count + 1) / 2
            If vDiv = Round((ALL(a).Pricecol.Count + 1) / 2, 0) And vDiv > 1 Then
                ALL(a).Median = (ALL(a).Pricecol(vDiv) + ALL(a).Pricecol(vDiv + 1)) / 2
            Else
                If vDiv = 0.5 Then
                    ALL(a).Median = 0
                Else
                    On Error Resume Next
                    If vDiv = 1 Then
                        ALL(a).Median = ALL(a).Pricecol(vDiv)
                    Else
                        ALL(a).Median = ALL(a).Pricecol(vDiv + 0.5)
                    End If
                    On Error GoTo Err_Routine
                End If
            End If
        Next

    End With
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_COM_SKUFormaulate", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Private Function AllocatePricing(ByVal State As String, ByVal isSpecial As String, ByVal Price As Single, ByVal isRecent As Boolean, ByVal MultiB As String)
    Dim Stateno As Long
    Dim ThisMB As Single, useprice As Single
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    Select Case State
        Case "NSW"
            Stateno = 1
        Case "ACT"
            Stateno = 1
        Case "VIC"
            Stateno = 2
        Case "QLD"
            Stateno = 3
        Case "SA"
            Stateno = 4
        Case "WA"
            Stateno = 5
        Case "NAT", "6", "Unknown"
            Stateno = 6
    End Select
        
    If InStr(1, isSpecial, "Promo") > 0 Then
        If MultiB <> "" And CBA_COM_SKU_Comp = "DM" Then
            ThisMB = DecypherMultibuy(MultiB)
            If ThisMB < Price Then useprice = ThisMB Else useprice = Price
        Else
            useprice = Price
        End If
        If Stateno < 6 Then
            If Promo(Stateno).Highest < useprice Then Promo(Stateno).Highest = useprice
            If Promo(Stateno).Lowest > useprice Then Promo(Stateno).Lowest = useprice
            Promo(Stateno).Pricecol.Add useprice
            Promo(Stateno).Mean = Promo(Stateno).Mean + useprice
            Promo(Stateno).MeanCount = Promo(Stateno).MeanCount + 1
            If InStr(1, Promo(Stateno).MetaData, isSpecial) > 0 Then Else If Promo(Stateno).MetaData = "" Then Promo(Stateno).MetaData = isSpecial Else Promo(Stateno).MetaData = Promo(Stateno).MetaData & ", " & isSpecial
        End If
        If Promo(6).Highest < useprice Then Promo(6).Highest = useprice
        If Promo(6).Lowest > useprice Then Promo(6).Lowest = useprice
        Promo(6).Pricecol.Add useprice
        Promo(6).Mean = Promo(6).Mean + useprice
        Promo(6).MeanCount = Promo(6).MeanCount + 1
        If InStr(1, Promo(6).MetaData, isSpecial) > 0 Then Else If Promo(6).MetaData = "" Then Promo(6).MetaData = isSpecial Else Promo(6).MetaData = Promo(6).MetaData & ", " & isSpecial
    Else
        If Stateno < 6 Then
            If NonPromo(Stateno).Highest < Price Then NonPromo(Stateno).Highest = Price
            If NonPromo(Stateno).Lowest > Price Then NonPromo(Stateno).Lowest = Price
            NonPromo(Stateno).Pricecol.Add Price
            NonPromo(Stateno).Mean = NonPromo(Stateno).Mean + Price
            NonPromo(Stateno).MeanCount = NonPromo(Stateno).MeanCount + 1
            If InStr(1, NonPromo(Stateno).MetaData, isSpecial) > 0 Then Else If NonPromo(Stateno).MetaData = "" Then NonPromo(Stateno).MetaData = isSpecial Else NonPromo(Stateno).MetaData = NonPromo(Stateno).MetaData & ", " & isSpecial
        End If
        If NonPromo(6).Highest < Price Then NonPromo(6).Highest = Price
        If NonPromo(6).Lowest > Price Then NonPromo(6).Lowest = Price
        NonPromo(6).Pricecol.Add Price
        NonPromo(6).Mean = NonPromo(6).Mean + Price
        NonPromo(6).MeanCount = NonPromo(6).MeanCount + 1
            If InStr(1, NonPromo(6).MetaData, isSpecial) > 0 Then Else If NonPromo(6).MetaData = "" Then NonPromo(6).MetaData = isSpecial Else NonPromo(6).MetaData = NonPromo(6).MetaData & ", " & isSpecial
    End If
    
    If isRecent = True Then
        If Stateno < 6 Then
            If Recent(Stateno).Highest < Price Then Recent(Stateno).Highest = Price
            If Recent(Stateno).Lowest > Price Then Recent(Stateno).Lowest = Price
            Recent(Stateno).Pricecol.Add Price
            Recent(Stateno).Mean = Recent(Stateno).Mean + Price
            Recent(Stateno).MeanCount = Recent(Stateno).MeanCount + 1
            If InStr(1, Recent(Stateno).MetaData, isSpecial) > 0 Then Else If Recent(Stateno).MetaData = "" Then Recent(Stateno).MetaData = isSpecial Else Recent(Stateno).MetaData = Recent(Stateno).MetaData & ", " & isSpecial
        End If
        If Recent(6).Highest < Price Then Recent(6).Highest = Price
        If Recent(6).Lowest > Price Then Recent(6).Lowest = Price
        Recent(6).Pricecol.Add Price
        Recent(6).Mean = Recent(6).Mean + Price
        Recent(6).MeanCount = Recent(6).MeanCount + 1
        If InStr(1, Recent(6).MetaData, isSpecial) > 0 Then Else If Recent(6).MetaData = "" Then Recent(6).MetaData = isSpecial Else Recent(6).MetaData = Recent(6).MetaData & ", " & isSpecial
    End If
    If MultiB <> "" Then
        ThisMB = DecypherMultibuy(MultiB)
        If Stateno < 6 Then
            If MultiBuy(Stateno).Highest < ThisMB Then MultiBuy(Stateno).Highest = ThisMB
            If MultiBuy(Stateno).Lowest > ThisMB Then MultiBuy(Stateno).Lowest = ThisMB
            MultiBuy(Stateno).Pricecol.Add ThisMB
            MultiBuy(Stateno).Mean = MultiBuy(Stateno).Mean + ThisMB
            MultiBuy(Stateno).MeanCount = MultiBuy(Stateno).MeanCount + 1
            If InStr(1, MultiBuy(Stateno).MetaData, isSpecial) > 0 Then Else If MultiBuy(Stateno).MetaData = "" Then MultiBuy(Stateno).MetaData = isSpecial Else MultiBuy(Stateno).MetaData = MultiBuy(Stateno).MetaData & ", " & isSpecial
        End If
        If MultiBuy(6).Highest < ThisMB Then MultiBuy(6).Highest = ThisMB
        If MultiBuy(6).Lowest > ThisMB Then MultiBuy(6).Lowest = ThisMB
        MultiBuy(6).Pricecol.Add ThisMB
        MultiBuy(6).Mean = MultiBuy(6).Mean + ThisMB
        MultiBuy(6).MeanCount = MultiBuy(6).MeanCount + 1
        If InStr(1, MultiBuy(6).MetaData, isSpecial) > 0 Then Else If MultiBuy(6).MetaData = "" Then MultiBuy(6).MetaData = isSpecial Else MultiBuy(6).MetaData = MultiBuy(6).MetaData & ", " & isSpecial
    End If
    If InStr(1, isSpecial, "Promo") > 0 And MultiB <> "" And CBA_COM_SKU_Comp = "DM" Then
        ThisMB = DecypherMultibuy(MultiB)
        If ThisMB < Price Then useprice = ThisMB Else useprice = Price
    Else
        useprice = Price
    End If
    If Stateno < 6 Then
        If ALL(Stateno).Highest < useprice Then ALL(Stateno).Highest = useprice
        If ALL(Stateno).Lowest > useprice Then ALL(Stateno).Lowest = useprice
        ALL(Stateno).Pricecol.Add useprice
        ALL(Stateno).Mean = ALL(Stateno).Mean + useprice
        ALL(Stateno).MeanCount = ALL(Stateno).MeanCount + 1
        If InStr(1, ALL(Stateno).MetaData, isSpecial) > 0 Then Else If ALL(Stateno).MetaData = "" Then ALL(Stateno).MetaData = isSpecial Else ALL(Stateno).MetaData = ALL(Stateno).MetaData & ", " & isSpecial
    End If
    If ALL(6).Highest < useprice Then ALL(6).Highest = useprice
    If ALL(6).Lowest > useprice Then ALL(6).Lowest = useprice
    ALL(6).Pricecol.Add useprice
    ALL(6).Mean = ALL(6).Mean + useprice
    ALL(6).MeanCount = ALL(6).MeanCount + 1
    If InStr(1, ALL(6).MetaData, isSpecial) > 0 Then Else If ALL(6).MetaData = "" Then ALL(6).MetaData = isSpecial Else ALL(6).MetaData = ALL(6).MetaData & ", " & isSpecial

Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AllocatePricing", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function Mode(ByVal col As Collection) As Single
    Dim Marr() As Single, most As Single
    Dim mArrcnt As Long
    Dim lNum As Long, mostcnt As Long
    Dim m As Variant
    Dim bfound As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    mArrcnt = 0
    ReDim Marr(1 To 2, 1 To 1)
    For Each m In col
        If mArrcnt = 0 Then
            mArrcnt = mArrcnt + 1
            Marr(1, mArrcnt) = m
            Marr(2, mArrcnt) = 1
        Else
            bfound = False
            For lNum = 1 To mArrcnt
                If Marr(1, lNum) = m Then
                    bfound = True
                    Marr(2, lNum) = Marr(2, lNum) + 1
                    Exit For
                End If
            Next
            If bfound = False Then
                mArrcnt = mArrcnt + 1
                ReDim Preserve Marr(1 To 2, 1 To mArrcnt)
                Marr(1, mArrcnt) = m
                Marr(2, mArrcnt) = 1
            End If
        End If
    Next
    most = 0
    For lNum = 1 To mArrcnt
        If Marr(2, lNum) > mostcnt Then
            most = Marr(1, lNum)
            mostcnt = Marr(2, lNum)
        End If
    Next
    Mode = most
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("CBA_COM_COMCompSKU;f-Mode", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function
Property Get CBA_COM_SKU_compcode() As String
    CBA_COM_SKU_compcode = CBA_COM_SKU_Ccode
End Property
Property Get CBA_COM_SKU_CompProdName() As String
    CBA_COM_SKU_CompProdName = CBA_COM_SKU_CCodeName
End Property
Property Get CBA_COM_SKU_CompPacksize() As String
    CBA_COM_SKU_CompPacksize = CBA_COM_SKU_CPack
End Property
Property Get CBA_COM_SKU_Competitor() As String
    CBA_COM_SKU_Competitor = CBA_COM_SKU_Comp
End Property

Property Get getPriceData(ByVal State As String, ByVal pdtype As String, ByVal pstype As String) As Single
    Dim Stateno As Long
    Select Case State
        Case "NSW"
            Stateno = 1
        Case "VIC"
            Stateno = 2
        Case "QLD"
            Stateno = 3
        Case "SA"
            Stateno = 4
        Case "WA"
            Stateno = 5
        Case "National"
            Stateno = 6
    End Select
    
    Select Case pdtype
        Case "Mean"
            If pstype = "Promo" Then
                getPriceData = Round(Promo(Stateno).Mean, 2)
            ElseIf pstype = "NonPromo" Then
                getPriceData = Round(NonPromo(Stateno).Mean, 2)
            ElseIf pstype = "Recent" Then
                getPriceData = Round(Recent(Stateno).Mean, 2)
            ElseIf pstype = "MultiBuy" Then
                getPriceData = Round(MultiBuy(Stateno).Mean, 2)
            ElseIf pstype = "ALL" Then
                 getPriceData = Round(ALL(Stateno).Mean, 2)
            End If
        Case "Mode"
            If pstype = "Promo" Then
                getPriceData = Round(Promo(Stateno).Mode, 2)
            ElseIf pstype = "NonPromo" Then
                getPriceData = Round(NonPromo(Stateno).Mode, 2)
            ElseIf pstype = "Recent" Then
                getPriceData = Round(Recent(Stateno).Mode, 2)
            ElseIf pstype = "MultiBuy" Then
                getPriceData = Round(MultiBuy(Stateno).Mode, 2)
            ElseIf pstype = "ALL" Then
                 getPriceData = Round(ALL(Stateno).Mode, 2)
            End If
        Case "Median"
            If pstype = "Promo" Then
                getPriceData = Round(Promo(Stateno).Median, 2)
            ElseIf pstype = "NonPromo" Then
                getPriceData = Round(NonPromo(Stateno).Median, 2)
            ElseIf pstype = "Recent" Then
                getPriceData = Round(Recent(Stateno).Median, 2)
            ElseIf pstype = "MultiBuy" Then
                getPriceData = Round(MultiBuy(Stateno).Median, 2)
            ElseIf pstype = "ALL" Then
                 getPriceData = Round(ALL(Stateno).Median, 2)
            End If
        Case "Highest"
            If pstype = "Promo" Then
                getPriceData = Round(Promo(Stateno).Highest, 2)
            ElseIf pstype = "NonPromo" Then
                getPriceData = Round(NonPromo(Stateno).Highest, 2)
            ElseIf pstype = "Recent" Then
                getPriceData = Round(Recent(Stateno).Highest, 2)
            ElseIf pstype = "MultiBuy" Then
                getPriceData = Round(MultiBuy(Stateno).Highest, 2)
            ElseIf pstype = "ALL" Then
                 getPriceData = Round(ALL(Stateno).Highest, 2)
            End If
        Case "Lowest"
            If pstype = "Promo" Then
                If CBA_COM_SKU_Comp = "DM" Then
                    If Promo(Stateno).Lowest = 999999 Then getPriceData = 0 Else If Round(MultiBuy(Stateno).Lowest, 2) < Round(Promo(Stateno).Lowest, 2) Then getPriceData = Round(MultiBuy(Stateno).Lowest, 2) Else getPriceData = Round(Promo(Stateno).Lowest, 2)
                Else
                    If Promo(Stateno).Lowest = 999999 Then getPriceData = 0 Else getPriceData = Round(Promo(Stateno).Lowest, 2)
                End If
            ElseIf pstype = "NonPromo" Then
                If NonPromo(Stateno).Lowest = 999999 Then getPriceData = 0 Else getPriceData = Round(NonPromo(Stateno).Lowest, 2)
            ElseIf pstype = "Recent" Then
                If Recent(Stateno).Lowest = 999999 Then getPriceData = 0 Else getPriceData = Round(Recent(Stateno).Lowest, 2)
            ElseIf pstype = "MultiBuy" Then
                If MultiBuy(Stateno).Lowest = 999999 Then getPriceData = 0 Else getPriceData = Round(MultiBuy(Stateno).Lowest, 2)
            ElseIf pstype = "ALL" Then
                 getPriceData = Round(ALL(Stateno).Lowest, 2)
            End If
    
    End Select

End Property
Property Get getMetaData(ByVal State As String, ByVal pstype As String) As String
    Dim Stateno As Long
    Select Case State
        Case "NSW"
            Stateno = 1
        Case "VIC"
            Stateno = 2
        Case "QLD"
            Stateno = 3
        Case "SA"
            Stateno = 4
        Case "WA"
            Stateno = 5
        Case "National"
            Stateno = 6
    End Select
    
    If pstype = "Promo" Then
        getMetaData = Promo(Stateno).MetaData
    ElseIf pstype = "NonPromo" Then
        getMetaData = NonPromo(Stateno).MetaData
    ElseIf pstype = "Recent" Then
        getMetaData = Recent(Stateno).MetaData
    ElseIf pstype = "MultiBuy" Then
        getMetaData = MultiBuy(Stateno).MetaData
    ElseIf pstype = "ALL" Then
         getMetaData = ALL(Stateno).MetaData
    End If


End Property
Function DecypherMultibuy(ByVal MB As String) As Single
    Dim a As Long, b As Long, c As Long, start As Long, ender As Long
    Dim QTY As Long
    Dim Price As Single
    Dim ident As String
    Dim atend As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    '    If MB = "21 per pack of 6" Then
    '    a = a
    '    End If
        
    If CBA_COM_SKU_Comp = "DM" Then
        If InStr(1, LCase(MB), "in any six") > 0 Then
            DecypherMultibuy = Trim(Mid(MB, 1, InStr(1, MB, "in any six") - 1))
        ElseIf InStr(1, LCase(MB), "per") > 0 Then
            Price = Trim(Mid(MB, 1, InStr(1, LCase(MB), "per") - 1))
            ender = 0
            For a = Len(MB) To 1 Step -1
                If ender = 0 And IsNumeric(Mid(MB, a, 1)) = True Then
                    ender = a
                ElseIf ender > 0 And (IsNumeric(Mid(MB, a, 1)) = False And Mid(MB, a, 1) <> ".") Then
                    start = a + 1
                    QTY = Mid(MB, start, ender - start + 1)
                    If QTY <> 0 Then DecypherMultibuy = Price / QTY Else DecypherMultibuy = 0
                    Exit Function
                End If
            Next
        ElseIf InStr(1, LCase(MB), "for") > 0 Then
            If InStr(1, LCase(MB), "cases") > 0 Then
                ident = "cases": If InStr(1, LCase(MB), "cases") > InStr(1, LCase(MB), "for") Then atend = True Else atend = False
            ElseIf InStr(1, LCase(MB), "bottle") > 0 Then
                ident = "bottle": If InStr(1, LCase(MB), "bottle") > InStr(1, LCase(MB), "for") Then atend = True Else atend = False
            ElseIf InStr(1, LCase(MB), "pack") > 0 Then
                ident = "pack": If InStr(1, LCase(MB), "pack") > InStr(1, LCase(MB), "for") Then atend = True Else atend = False
            ElseIf InStr(1, LCase(MB), "bundle") > 0 Then
                ident = "bundle": If InStr(1, LCase(MB), "bundle") > InStr(1, LCase(MB), "for") Then atend = True Else atend = False
            End If
            If atend = True Then
                start = 0: ender = 0
                For a = 1 To Len(MB)
                    If IsNumeric(Mid(MB, a, 1)) = True And start = 0 Then
                        start = a
                    ElseIf (IsNumeric(Mid(MB, a, 1)) = False And Mid(MB, a, 1) <> ".") And start > 0 Then
                        ender = a - 1
                        Price = Mid(MB, start, ender - start + 1)
                        Exit For
                    End If
                Next
                start = 0: ender = 0
                For a = Len(MB) To 1 Step -1
                    If IsNumeric(Mid(MB, a, 1)) = True And ender = 0 Then
                        ender = a
                    ElseIf (IsNumeric(Mid(MB, a, 1)) = False And Mid(MB, a, 1) <> ".") And ender > 0 Then
                        start = a + 1
                        QTY = Mid(MB, start, ender - start + 1)
                        Exit For
                    End If
                Next
                If QTY <> 0 Then DecypherMultibuy = Price / QTY Else DecypherMultibuy = 0
            
            Else
                start = 0: ender = 0
                For a = 1 To Len(MB)
                    If IsNumeric(Mid(MB, a, 1)) = True And start = 0 Then
                        start = a
                    ElseIf IsNumeric(Mid(MB, a, 1)) = False And start > 0 Then
                        ender = a - 1
                        QTY = Mid(MB, start, ender - start + 1)
                        Exit For
                    End If
                Next
                start = 0: ender = 0
                For a = Len(MB) To 1 Step -1
                    If IsNumeric(Mid(MB, a, 1)) = True And ender = 0 Then
                        ender = a
                    ElseIf IsNumeric(Mid(MB, a, 1)) = False And ender > 0 Then
                        start = a + 1
                        Price = Mid(MB, start, ender - start + 1)
                        Exit For
                    End If
                Next
                If QTY <> 0 Then DecypherMultibuy = Price / QTY Else DecypherMultibuy = 0
            End If
        End If
   
    Else
        
        For a = 1 To Len(MB)
        
            If LCase(Mid(MB, a + 1, 4)) = " for" And IsNumeric(Mid(MB, a, 1)) Then
                If Mid(MB, a + 6, 1) = "$" Then b = a + 7 Else b = a + 6
                For c = a To 1 Step -1
                    If IsNumeric(Mid(MB, c, 1)) = False Or Mid(MB, c, 1) = "" Then
                        start = c + 1
                        Exit For
                    End If
                Next
                If start = 0 Then start = 1
                QTY = Trim(Mid(MB, start, (a + 1) - start))
                
                For c = b To Len(MB)
                    
                    If IsNumeric(Mid(MB, c, 1)) = True Or Mid(MB, c, 1) = "." Then
                    Else
                        Price = Trim(Mid(MB, b, c - b))
                        Exit For
                    End If
                    If Price = 0 And c = Len(MB) Then
                        Price = Trim(Mid(MB, b, Len(MB) - (b - 1)))
                    End If
                Next
                If QTY <> 0 Then DecypherMultibuy = Price / QTY Else DecypherMultibuy = 0
                Exit For
            End If
    
        Next
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-decypherMultibuy", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function
