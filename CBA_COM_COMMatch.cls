VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_COM_COMMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private CBA_COM_dbbeingqueried  As Long, CBA_COM_Ccode As String, CBA_COM_Acode As Long, CBA_COM_ACG As Long, CBA_COM_ASCG As Long
Private CBA_COM_mtype As String, CBA_COM_Competitor As String, CBA_COM_Comp As String, CBA_COM_CcodeName As String, CBA_COM_CPack As String
Private CBA_COM_Ccupunit As String, CBA_COM_pricetouse As String, CBA_COM_Prorataby As String, CBA_COM_ADesc As String
Private CBA_COM_Ccupmes As Single, CBA_COM_Compdiv As Single, CBA_COM_Compmult As Single, CBA_COM_OrigCompPack As Single
Private CBA_COM_dtefrom As Date, CBA_COM_dteto As Date
Private Type CBA_COM_Tdpm
    ScrapeDate As Date
    State As String
    Shelf As Single
    Was As Single
    Discount As Single
    Permeasure As Single
    Special As Boolean
    ProRata As Single
    NonSpecialProRata As Single
    A_Retail As Single
    Diff As String
    Count As Long
End Type
Private CBA_COM_DatePriceMatrix() As CBA_COM_Tdpm
Private CBA_COM_COMRS As ADODB.Recordset
Private CBA_COM_CBISRS As ADODB.Recordset

Public Sub CBA_COM_Formulate(ByVal CompCode As String, ByVal A_Code As Long, ByVal MatchType As String, ByVal DateFrom As Date, Optional ByVal DateTo As Date, Optional ByVal ModeOnly As Boolean)
    Dim lNum As Long, a As Long, b As Long, c As Long, starter As Long, j As Long
    Dim shelfcount As Long
    Dim curDatescraped As Date
    Dim findstate As String
    Dim thistime As Single
    Dim noaldiret As Boolean, shelffound As Boolean
    Dim minAlidpricedate As Date, maxAlidpricedate As Date
    Dim shelfcnt() As Variant
    Dim rw As Boolean
    Dim tR As CBAR_Report
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    'ReDim CBISarr(0 To 21, 0 To 0)
    ReDim CBA_COMarr(0 To 10, 0 To 0)
    
    'check their packsize on 790069
    If ModeOnly = True Then ReDim shelfcnt(1 To 3, 1 To 1)
    
    If DateTo = 0 Then CBA_COM_dteto = Date Else CBA_COM_dteto = DateTo
    
    CBA_COM_dtefrom = DateFrom
    CBA_COM_Ccode = CompCode
    CBA_COM_Acode = A_Code
    CBA_COM_mtype = MatchType
    
    If Mid(MatchType, 1, 2) = "B_" Then
        On Error Resume Next
        If UBound(CBA_COM_ADarr, 2) < 0 Then
            Err.Clear
            On Error GoTo Err_Routine
            Me.CBA_COM_buildSQL "CBIS_ClassComAldiData"
            CBA_COM_ADarr = CBA_CBISarr
        End If
        Err.Clear
        On Error GoTo Err_Routine
        MatchType = CMM_getComp2Find(MatchType, CBA_COM_ADarr(1, 0))
    End If
    
    If InStr(1, LCase(MatchType), "coles") > 0 Then
        CBA_COM_Comp = "C"
        CBA_COM_Competitor = "Coles"
    ElseIf InStr(1, LCase(MatchType), "ww") > 0 Then
        CBA_COM_Comp = "WW"
        CBA_COM_Competitor = "Woolworths"
    ElseIf InStr(1, LCase(MatchType), "dm") > 0 Then
        CBA_COM_Comp = "DM"
        CBA_COM_Competitor = "Dan Murphy's"
    
    ElseIf InStr(1, LCase(MatchType), "fc") > 0 Then
        CBA_COM_Comp = "FC"
        CBA_COM_Competitor = "First Choice"
    End If
    'if CBA_COM_COmp = "" then CMM_getComp2Find(
    CBA_COM_dbbeingqueried = CBA_DBtoQuery
    
 
    
    On Error Resume Next
    If A_Code <> CBA_COM_AParr(0, 0) Or (A_Code = CBA_COM_AParr(0, 0) And (CBA_COM_AParr(2, 0) > DateFrom Or CBA_COM_AParr(3, UBound(CBA_COM_AParr, 2)) < DateTo)) Then
        Err.Clear
        On Error GoTo Err_Routine
        Me.CBA_COM_buildSQL "CBIS_ClassComAldiData"
        CBA_COM_ADarr = CBA_CBISarr
        Erase CBA_CBISarr
        ReDim CBA_COM_AParr(0 To 4, 0 To 0)
        If CBA_COM_ADarr(0, 0) <> 0 Then
        If CBA_COM_ADarr(1, 0) = "58" Then
            For b = 0 To UBound(CBA_COM_APParr, 2)
'            Range(wks_Data.Cells(1, 52), wks_Data.Cells(1, 56)).EntireColumn.ClearContents
'            For j = 0 To 4
'                wks_Data.Cells(b + 1, 52 + j) = APParr(j, b)
'            Next
'
'            If APParr(1, b) = "WA" Then
'            a = a
'            End If
'
            
            
            
            If CBA_COM_APParr(0, b) = CBA_COM_ADarr(0, 0) Then
'                If b = 400 Then
'                testing = APParr(1, b)
'                End If
                If CBA_COM_AParr(0, 0) > 0 Then ReDim Preserve CBA_COM_AParr(0 To 4, 0 To UBound(CBA_COM_AParr, 2) + 1)
                For j = 0 To 4
                    CBA_COM_AParr(j, UBound(CBA_COM_AParr, 2)) = CBA_COM_APParr(j, b)
                Next
            End If
            Next

        Else
            Me.CBA_COM_buildSQL "CBIS_ClassComAldiRetails"
            'Line to pull most recent AldiRetail if no retail exists in query period.
            'This is to cover for products in transition to new productcodes. Limited to Top150 Report and Active Match Report.
            If CBA_CBISarr(0, 0) = 0 Then
                tR = CBAR_Runtime.getActiveReport
                If tR.ReportNo = 8 Or tR.ReportNo = 1 Then Me.CBA_COM_buildSQL "CBIS_ClassComAldiRetailsFuture"
            End If
            CBA_COM_AParr = CBA_CBISarr
            Erase CBA_CBISarr
        End If
        End If
    End If
    On Error GoTo Err_Routine

    
    If CBA_COM_ADarr(0, 0) <> 0 Then
        CBA_COM_ADesc = CBA_COM_ADarr(4, LBound(CBA_COM_ADarr, 2))
        CBA_COM_ACG = CBA_COM_ADarr(1, LBound(CBA_COM_ADarr, 2))
        
        If CBA_COM_ACG = 2 Or CBA_COM_ACG = 3 Or CBA_COM_ACG = 1 Then
            CBA_COM_CPack = "per bottle"
        ElseIf CBA_COM_ACG = 4 Then
            CBA_COM_CPack = "6"
        End If
        If IsNull(CBA_COM_ADarr(2, LBound(CBA_COM_ADarr, 2))) Then CBA_COM_ASCG = 0 Else CBA_COM_ASCG = CBA_COM_ADarr(2, LBound(CBA_COM_ADarr, 2))
        If CBA_COM_ADarr(9, 0) = 1 Or CBA_COM_ACG = 58 Then rw = True
    Else
    'a = a
    End If
    thistime = Round(Timer, 4)
    
    If ModeOnly = True Then Me.CBA_COM_buildSQL "COM_ClassComProdDataMode" Else Me.CBA_COM_buildSQL "COM_ClassComProdData"
    
    
    'thattime = Round(Timer - thistime, 4)
    'thistime = Round(Timer, 4)
    If CBA_COMarr(0, 0) <> 0 Then
    
    
    'testlong = UBound(CBA_COMarr, 2)
    CBA_COM_CcodeName = CBA_COMarr(1, 0)
    If CBA_COM_Comp <> "DM" Then CBA_COM_CPack = CBA_COMarr(2, 0)
'    If CBA_COM_ACG = "58" Then
'        'ReDim packarrproduce(0 To 0, 0 To 0)
'        packarrproduce(0, 0) = CBA_COMarr(2, 0)
'        CBA_COM_PackDecode.decodepack True, packarrproduce(0, 0)
'    End If
    If IsNull(CBA_COMarr(6, 0)) Then Else CBA_COM_Ccupunit = CBA_COMarr(6, 0)
    If IsNull(CBA_COMarr(7, 0)) Then Else CBA_COM_Ccupmes = CBA_COMarr(7, 0)
    starter = 1

    If CBA_COM_ADarr(0, 0) <> 0 Then CBA_COM_resolvethiscomp CBA_COM_Ccode, 0, DateFrom, rw
 


        On Error Resume Next
        If CBA_COM_AParr(0, 0) = 0 Then
            noaldiret = True
        End If
        Err.Clear
        On Error GoTo Err_Routine
        
            lNum = 0
            curDatescraped = CBA_COMarr(0, 0)
        If CBA_COM_ACG <> 58 Then
            If noaldiret = True Then
                minAlidpricedate = 0
                maxAlidpricedate = 0
            Else
                minAlidpricedate = CBA_COM_AParr(2, 0)
                maxAlidpricedate = CBA_COM_AParr(3, 0)
                For a = LBound(CBA_COM_AParr, 2) To UBound(CBA_COM_AParr, 2)
                    If CBA_COM_AParr(2, a) < minAlidpricedate Then
                        minAlidpricedate = CBA_COM_AParr(2, a)
                    End If
                    If CBA_COM_AParr(3, a) > maxAlidpricedate Then
                        maxAlidpricedate = CBA_COM_AParr(3, a)
                    End If
                Next
            End If
        End If
                
        'THIS IS WHERE THE LOOPING STARTS!!!
        For a = 0 To UBound(CBA_COMarr, 2)
            If a < UBound(CBA_COMarr, 2) + 1 Then
                If CBA_COMarr(2, a) <> CBA_COM_CPack And CBA_COM_ADarr(0, 0) <> 0 Then
                    CBA_COM_CPack = CBA_COMarr(2, a)
                    CBA_COM_resolvethiscomp CBA_COM_Ccode, a, CBA_COMarr(0, a), rw
                End If
            End If
            If CDate(CBA_COMarr(0, a)) <> CDate(curDatescraped) Then
                lNum = lNum + 1
                ReDim Preserve CBA_COM_DatePriceMatrix(1 To lNum)
                With CBA_COM_DatePriceMatrix(lNum)
                    .Shelf = 99999999
                    For b = starter To lNum - 1
                        If ModeOnly = True Then
                            shelffound = False
                            For c = LBound(shelfcnt, 2) To UBound(shelfcnt, 2)
                                If shelfcnt(1, c) = "" And UBound(shelfcnt, 2) = 1 Then
                                    shelfcnt(1, c) = CBA_COM_DatePriceMatrix(b).Shelf
                                    shelfcnt(2, c) = b
                                    shelfcnt(3, c) = 1
                                    shelffound = True
                                Else
                                    If shelfcnt(1, c) = CBA_COM_DatePriceMatrix(b).Shelf Then
                                        shelffound = True
                                        shelfcnt(3, c) = shelfcnt(3, c) + 1
                                        Exit For
                                    End If
                                End If
                            Next
                            If shelffound = False Then
                                ReDim Preserve shelfcnt(1 To 3, 1 To UBound(shelfcnt, 2) + 1)
                                shelfcnt(1, UBound(shelfcnt, 2)) = CBA_COM_DatePriceMatrix(b).Shelf
                                shelfcnt(2, UBound(shelfcnt, 2)) = b
                                shelfcnt(3, UBound(shelfcnt, 2)) = 1
                            End If
                        Else
                            If CBA_COM_DatePriceMatrix(b).Shelf < .Shelf Then
                                 .ScrapeDate = CBA_COM_DatePriceMatrix(b).ScrapeDate
                                 .State = "national"
                                 .Shelf = CBA_COM_DatePriceMatrix(b).Shelf
                                 .Was = CBA_COM_DatePriceMatrix(b).Was
                                 .Discount = CBA_COM_DatePriceMatrix(b).Discount
                                 .Permeasure = CBA_COM_DatePriceMatrix(b).Permeasure
                                 .Special = CBA_COM_DatePriceMatrix(b).Special
                                 '.Count = CBA_COM_DatePriceMatrix(b).Count
                                 If CBA_COM_ADarr(0, 0) <> 0 Then .NonSpecialProRata = CBA_COM_DatePriceMatrix(b).NonSpecialProRata
                                 If CBA_COM_ADarr(0, 0) <> 0 Then .ProRata = CBA_COM_DatePriceMatrix(b).ProRata
                                 If CBA_COM_AParr(0, 0) = 0 Then
                                     .A_Retail = 0
                                     .Diff = 0
                                 Else
                                     For c = 0 To UBound(CBA_COM_AParr, 2)
                                         If .State <> "national" Or CBA_COM_ACG = 58 Then
                                             If .State = "national" And CBA_COM_ACG = 58 Then findstate = "nsw" Else findstate = .State
                                             If LCase(CBA_COM_AParr(1, c)) = findstate Then
                                                 If CBA_COM_AParr(2, c) <= .ScrapeDate And CBA_COM_AParr(3, c) >= .ScrapeDate Then
                                                     .A_Retail = CBA_COM_AParr(4, c)
                                                     If .ProRata <> 0 Then .Diff = (.ProRata - .A_Retail) / .ProRata
                                                     Exit For
                                                 End If
                                             End If
                                         ElseIf .State = "national" Then
                                             If LCase(CBA_COM_AParr(5, c)) = 1 Then
                                                 If (CBA_COM_AParr(2, c) <= .ScrapeDate And CBA_COM_AParr(3, c) >= .ScrapeDate) Then
                                                     .A_Retail = CBA_COM_AParr(4, c)
                                                     If .ProRata <> 0 Then .Diff = (.ProRata - .A_Retail) / .ProRata
                                                     Exit For
                                                 ElseIf minAlidpricedate <> "12:00:00 AM" And minAlidpricedate > .ScrapeDate And CBA_COM_AParr(2, c) = minAlidpricedate Then
                                                     .A_Retail = CBA_COM_AParr(4, c)
                                                     If .ProRata <> 0 Then .Diff = (.ProRata - .A_Retail) / .ProRata
                                                     Exit For
                                                 ElseIf maxAlidpricedate <> "12:00:00 AM" And maxAlidpricedate < .ScrapeDate And CBA_COM_AParr(3, c) = maxAlidpricedate Then
                                                     .A_Retail = CBA_COM_AParr(4, c)
                                                     If .ProRata <> 0 Then .Diff = (.ProRata - .A_Retail) / .ProRata
                                                     Exit For
                                                 End If
                                             End If
                                         End If
                                     Next
                                 End If
                             End If
                        End If
                    Next
                    
                    'Code added to get mode on national price
                    If ModeOnly = True Then
                        shelfcount = 0
                        For c = LBound(shelfcnt, 2) To UBound(shelfcnt, 2)
                            If shelfcnt(3, c) > shelfcount Then
                                shelfcount = shelfcnt(3, c)
                                b = shelfcnt(2, c)
                            End If
                        Next
                         .ScrapeDate = CBA_COM_DatePriceMatrix(b).ScrapeDate
                         .State = "national"
                         .Shelf = CBA_COM_DatePriceMatrix(b).Shelf
                         .Was = CBA_COM_DatePriceMatrix(b).Was
                         .Discount = CBA_COM_DatePriceMatrix(b).Discount
                         .Permeasure = CBA_COM_DatePriceMatrix(b).Permeasure
                         .Special = CBA_COM_DatePriceMatrix(b).Special
                         '.Count = CBA_COM_DatePriceMatrix(b).Count
                         If CBA_COM_ADarr(0, 0) <> 0 Then .NonSpecialProRata = CBA_COM_DatePriceMatrix(b).NonSpecialProRata
                         If CBA_COM_ADarr(0, 0) <> 0 Then .ProRata = CBA_COM_DatePriceMatrix(b).ProRata
                         If CBA_COM_AParr(0, 0) = 0 Then
                             .A_Retail = 0
                             .Diff = 0
                         Else
                             For c = 0 To UBound(CBA_COM_AParr, 2)
                                 If .State <> "national" Or CBA_COM_ACG = 58 Then
                                     If .State = "national" And CBA_COM_ACG = 58 Then findstate = "nsw" Else findstate = .State
                                     If LCase(CBA_COM_AParr(1, c)) = findstate Then
                                         If CBA_COM_AParr(2, c) <= .ScrapeDate And CBA_COM_AParr(3, c) >= .ScrapeDate Then
                                             .A_Retail = CBA_COM_AParr(4, c)
                                             If .ProRata <> 0 Then .Diff = (.ProRata - .A_Retail) / .ProRata
                                             Exit For
                                         End If
                                     End If
                                 ElseIf .State = "national" Then
                                     If LCase(CBA_COM_AParr(5, c)) = 1 Then
                                         If (CBA_COM_AParr(2, c) <= .ScrapeDate And CBA_COM_AParr(3, c) >= .ScrapeDate) Then
                                             .A_Retail = CBA_COM_AParr(4, c)
                                             If .ProRata <> 0 Then .Diff = (.ProRata - .A_Retail) / .ProRata
                                             Exit For
                                         ElseIf minAlidpricedate <> "12:00:00 AM" And minAlidpricedate > .ScrapeDate And CBA_COM_AParr(2, c) = minAlidpricedate Then
                                             .A_Retail = CBA_COM_AParr(4, c)
                                             If .ProRata <> 0 Then .Diff = (.ProRata - .A_Retail) / .ProRata
                                             Exit For
                                         ElseIf maxAlidpricedate <> "12:00:00 AM" And maxAlidpricedate < .ScrapeDate And CBA_COM_AParr(3, c) = maxAlidpricedate Then
                                             .A_Retail = CBA_COM_AParr(4, c)
                                             If .ProRata <> 0 Then .Diff = (.ProRata - .A_Retail) / .ProRata
                                             Exit For
                                         End If
                                     End If
                                 End If
                             Next
                         End If
                     End If
                    'End of Code added to get mode on national price
                End With
                If CDate(CBA_COMarr(0, a)) < DateFrom Then Exit For
                starter = lNum + 1
                curDatescraped = CDate(CBA_COMarr(0, a))
                If ModeOnly = True Then ReDim shelfcnt(1 To 3, 1 To 1)
            End If
            lNum = lNum + 1

            ReDim Preserve CBA_COM_DatePriceMatrix(1 To lNum)
            With CBA_COM_DatePriceMatrix(lNum)
                .ScrapeDate = CBA_COMarr(0, a)
                .State = LCase(CBA_COMarr(9, a))
                .Shelf = CBA_COMarr(3, a)
                .Was = CBA_COMarr(4, a)
                If IsNull(CBA_COMarr(5, a)) Then .Discount = 0 Else .Discount = CBA_COMarr(5, a)
                '.Count = CBA_COMarr(11, a)
                
                If (CBA_COM_ACG = 62 Or CBA_COM_ACG = 64) And (IsNull(CBA_COMarr(8, a)) Or CBA_COMarr(8, a) = 0) Then
                        For b = 1 To UBound(CBA_COM_arrOutput, 2)
                            If CBA_COM_arrOutput(3, b) = CBA_COMarr(3, a) + CBA_COMarr(5, a) And CBA_COM_arrOutput(5, b) = "g" Then
                                .Permeasure = (CBA_COMarr(3, a) / CBA_COM_arrOutput(4, b)) * 1000
                                Exit For
                            End If
                        Next
                Else
                    If IsNull(CBA_COMarr(8, a)) Then .Permeasure = 0 Else .Permeasure = CBA_COMarr(8, a)
                End If
                
                
                If CBA_COMarr(10, a) = "Special" Or (.Discount > 0 And CBA_COMarr(10, a) <> "X") Then .Special = True
                If CBA_COM_ADarr(0, 0) <> 0 And CBA_COM_Compdiv <> 0 Then
                    If CBA_COM_pricetouse = "shelf" Then
                        .ProRata = (.Shelf / CBA_COM_Compdiv) * CBA_COM_Compmult
                            .NonSpecialProRata = ((.Shelf + .Discount) / CBA_COM_Compdiv) * CBA_COM_Compmult
                            If .Special = True Then
                                For b = 1 To UBound(CBA_COMarr, 2)
                                    If CBA_COMarr(0, b) = .ScrapeDate Then
                                    If CBA_COMarr(0, b) = .ScrapeDate And CBA_COMarr(3, b) + CBA_COMarr(5, b) < .Shelf + .Discount And (IsNull(CBA_COMarr(10, b)) Or CBA_COMarr(10, b) = "") And CBA_COMarr(5, b) = 0 Then
                                        .Discount = CBA_COMarr(3, b) - .Shelf
                                        .NonSpecialProRata = ((.Shelf + .Discount) / CBA_COM_Compdiv) * CBA_COM_Compmult
                                    ElseIf CBA_COMarr(0, b) > .ScrapeDate Then
                                        Exit For
                                    End If
                                    End If
                                Next
                            End If
                    Else
                        .ProRata = (.Permeasure / CBA_COM_Compdiv) * CBA_COM_Compmult
                        If .Shelf <> 0 And .Permeasure <> 0 Then .NonSpecialProRata = ((.Shelf + .Discount) / ((.Shelf / .Permeasure) * CBA_COM_Compdiv)) * CBA_COM_Compmult
                    End If
                End If
                If CBA_COM_AParr(0, 0) = 0 Then
                    .A_Retail = 0
                    .Diff = 0
                Else
                    For b = 0 To UBound(CBA_COM_AParr, 2)
                        If LCase(CBA_COM_AParr(1, b)) = .State Then
                            If CBA_COM_AParr(2, b) <= .ScrapeDate And CBA_COM_AParr(3, b) >= .ScrapeDate Then
                                .A_Retail = CBA_COM_AParr(4, b)
                                If .A_Retail <> 0 And .ProRata <> 0 Then .Diff = Format(((.ProRata - .A_Retail) / .ProRata), "0.00%")
                                Exit For
                            End If
                        End If
                    Next
                End If
            End With
        Next
                
        'for the last entry of the night
        lNum = lNum + 1
        ReDim Preserve CBA_COM_DatePriceMatrix(1 To lNum)
        With CBA_COM_DatePriceMatrix(lNum)
            .Shelf = 99999999
            For b = starter To lNum - 1
               If CBA_COM_DatePriceMatrix(b).Shelf < .Shelf Then
                    .ScrapeDate = CBA_COM_DatePriceMatrix(b).ScrapeDate
                    .State = "national"
                    .Shelf = CBA_COM_DatePriceMatrix(b).Shelf
                    .Was = CBA_COM_DatePriceMatrix(b).Was
                    .Discount = CBA_COM_DatePriceMatrix(b).Discount
                    .Permeasure = CBA_COM_DatePriceMatrix(b).Permeasure
                    .Special = CBA_COM_DatePriceMatrix(b).Special
                    '.Count = CBA_COM_DatePriceMatrix(b).Count
                    If CBA_COM_ADarr(0, 0) <> 0 Then .NonSpecialProRata = CBA_COM_DatePriceMatrix(b).NonSpecialProRata
                    If CBA_COM_ADarr(0, 0) <> 0 Then .ProRata = CBA_COM_DatePriceMatrix(b).ProRata
                    If CBA_COM_AParr(0, 0) = 0 Then
                        .A_Retail = 0
                        .Diff = 0
                    Else
                        For c = 0 To UBound(CBA_COM_AParr, 2)
                            If .State <> "national" Or CBA_COM_ACG = 58 Then
                                If .State = "national" And CBA_COM_ACG = 58 Then findstate = "nsw" Else findstate = .State
                                If LCase(CBA_COM_AParr(1, c)) = findstate Then
                                    If CBA_COM_AParr(2, c) <= .ScrapeDate And CBA_COM_AParr(3, c) >= .ScrapeDate Then
                                        .A_Retail = CBA_COM_AParr(4, c)
                                        If .ProRata <> 0 Then .Diff = (.ProRata - .A_Retail) / .ProRata
                                        Exit For
                                    End If
                                End If
                            ElseIf .State = "national" Then
                                If LCase(CBA_COM_AParr(5, c)) = 1 Then
                                    If (CBA_COM_AParr(2, c) <= .ScrapeDate And CBA_COM_AParr(3, c) >= .ScrapeDate) Then
                                        .A_Retail = CBA_COM_AParr(4, c)
                                        If .ProRata <> 0 Then .Diff = (.ProRata - .A_Retail) / .ProRata
                                        Exit For
                                    ElseIf minAlidpricedate <> "12:00:00 AM" And minAlidpricedate > .ScrapeDate And CBA_COM_AParr(2, c) = minAlidpricedate Then
                                        .A_Retail = CBA_COM_AParr(4, c)
                                        If .ProRata <> 0 Then .Diff = (.ProRata - .A_Retail) / .ProRata
                                        Exit For
                                    ElseIf maxAlidpricedate <> "12:00:00 AM" And maxAlidpricedate < .ScrapeDate And CBA_COM_AParr(3, c) = maxAlidpricedate Then
                                        .A_Retail = CBA_COM_AParr(4, c)
                                        If .ProRata <> 0 Then .Diff = (.ProRata - .A_Retail) / .ProRata
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        End With
    End If
    'thattime = Round(Timer - thistime, 4)
    CBA_DBtoQuery = CBA_COM_dbbeingqueried
    
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_COM_Formulate", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & strSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Property Get compet() As String
    compet = CBA_COM_Comp
End Property
Property Get MatchType() As String
    MatchType = CBA_COM_mtype
End Property
Property Get CompCode() As String
    CompCode = CBA_COM_Ccode
End Property
Property Get CompProdName() As String
    CompProdName = CBA_COM_CcodeName
End Property
Property Get CompPacksize() As String
    CompPacksize = CBA_COM_CPack
End Property
Property Get HowCompvalue() As String
    HowCompvalue = CBA_COM_pricetouse
End Property
Property Get HowComp() As String
    HowComp = CBA_COM_Prorataby
End Property
Property Get AldiPCode() As Long
    AldiPCode = CBA_COM_Acode
End Property
Property Get AldiPName() As String
    AldiPName = CBA_COM_ADesc
End Property
Property Get AldiPCG() As Long
    AldiPCG = CBA_COM_ACG
End Property
Property Get AldiPSCG() As Long
    AldiPSCG = CBA_COM_ASCG
End Property
Property Get Compcupunit() As String
    Compcupunit = CBA_COM_Ccupunit
End Property
Property Get Compcupmeasure() As String
    Compcupmeasure = CBA_COM_Ccupmes
End Property
Property Get DateFrom() As Date
    DateFrom = CBA_COM_dtefrom
End Property
Property Get DateTo() As Date
    DateTo = CBA_COM_dteto
End Property
Property Get Competitor() As String
    Competitor = CBA_COM_Competitor
End Property
Property Get CompDivideby() As String
    CompDivideby = CBA_COM_Compdiv
End Property
Property Get CompMultby() As String
    CompMultby = CBA_COM_Compmult
End Property
Property Get CompOriginalPack() As Long
    CompOriginalPack = CBA_COM_OrigCompPack
End Property

Property Get getdates() As Date()
    Dim TempArr() As Date
    Dim a As Long
    Dim lNum As Long
    On Error Resume Next
    
    If isDate(CBA_COM_DatePriceMatrix(1).ScrapeDate) = False Then
        ReDim getdates(1)
        Exit Property
    Else
    End If
    Err.Clear
    On Error GoTo 0
    lNum = 1
    ReDim TempArr(1 To lNum)
    TempArr(lNum) = CBA_COM_DatePriceMatrix(lNum).ScrapeDate
    For a = 1 To UBound(CBA_COM_DatePriceMatrix)
        If CBA_COM_DatePriceMatrix(a).ScrapeDate <> TempArr(lNum) Then
        lNum = lNum + 1
        ReDim Preserve TempArr(1 To lNum)
        TempArr(lNum) = CBA_COM_DatePriceMatrix(a).ScrapeDate
        End If
    Next
    
    getdates = TempArr
End Property
Property Get Pricedata(ByVal datefor As Date, ByVal WhatPrice As String, ByVal State As String, Optional byCurrentOnly As Boolean) As Variant
    Dim a As Long
    
    On Error Resume Next
    If UBound(CBA_COM_DatePriceMatrix, 1) = -1 Then
        Err.Clear
        On Error GoTo 0
        Pricedata = 0
    Else
        On Error GoTo 0
        If byCurrentOnly = True Then
            For a = 1 To UBound(CBA_COM_DatePriceMatrix)
                If CBA_COM_DatePriceMatrix(a).ScrapeDate <= datefor Then
                    If CBA_COM_DatePriceMatrix(a).State = LCase(State) Then
                        If LCase(WhatPrice) = "shelf" Then Pricedata = CBA_COM_DatePriceMatrix(a).Shelf
                        If LCase(WhatPrice) = "was" Then Pricedata = CBA_COM_DatePriceMatrix(a).Was
                        If LCase(WhatPrice) = "discount" Then Pricedata = CBA_COM_DatePriceMatrix(a).Discount
                        If LCase(WhatPrice) = "permeasure" Then Pricedata = CBA_COM_DatePriceMatrix(a).Permeasure
                        If LCase(WhatPrice) = "special" Then Pricedata = CBA_COM_DatePriceMatrix(a).Special
                        If LCase(WhatPrice) = "prorata" Then Pricedata = CBA_COM_DatePriceMatrix(a).ProRata
                        If LCase(WhatPrice) = "nonpromoprorata" Then Pricedata = CBA_COM_DatePriceMatrix(a).NonSpecialProRata
                        If LCase(WhatPrice) = "isspecial" Then Pricedata = CBA_COM_DatePriceMatrix(a).Special
                        If LCase(WhatPrice) = "aldiretail" Then Pricedata = CBA_COM_DatePriceMatrix(a).A_Retail
                        If LCase(WhatPrice) = "delta" Then Pricedata = CBA_COM_DatePriceMatrix(a).Diff
                        If LCase(WhatPrice) = "state" Then Pricedata = CBA_COM_DatePriceMatrix(a).State
                Exit For
                    End If
                End If
            Next
        Else
            For a = 1 To UBound(CBA_COM_DatePriceMatrix)
                If CBA_COM_DatePriceMatrix(a).ScrapeDate = datefor Then
                    If CBA_COM_DatePriceMatrix(a).State = LCase(State) Then
                        If LCase(WhatPrice) = "shelf" Then Pricedata = CBA_COM_DatePriceMatrix(a).Shelf
                        If LCase(WhatPrice) = "was" Then Pricedata = CBA_COM_DatePriceMatrix(a).Was
                        If LCase(WhatPrice) = "discount" Then Pricedata = CBA_COM_DatePriceMatrix(a).Discount
                        If LCase(WhatPrice) = "permeasure" Then Pricedata = CBA_COM_DatePriceMatrix(a).Permeasure
                        If LCase(WhatPrice) = "special" Then Pricedata = CBA_COM_DatePriceMatrix(a).Special
                        If LCase(WhatPrice) = "prorata" Then Pricedata = CBA_COM_DatePriceMatrix(a).ProRata
                        If LCase(WhatPrice) = "nonpromoprorata" Then Pricedata = CBA_COM_DatePriceMatrix(a).NonSpecialProRata
                        If LCase(WhatPrice) = "isspecial" Then Pricedata = CBA_COM_DatePriceMatrix(a).Special
                        If LCase(WhatPrice) = "aldiretail" Then Pricedata = CBA_COM_DatePriceMatrix(a).A_Retail
                        If LCase(WhatPrice) = "delta" Then Pricedata = CBA_COM_DatePriceMatrix(a).Diff
                        If LCase(WhatPrice) = "state" Then Pricedata = CBA_COM_DatePriceMatrix(a).State
                    Exit For
                    End If
                End If
            Next
        End If
    End If
    'If IsEmpty(Pricedata) Then Pricedata = -151
End Property
Property Get RetailHistoryArray(ByVal WhatSection As String) As Variant()
Dim TempArr As Variant
Dim a As Long
On Error Resume Next
If IsEmpty(UBound(CBA_COM_DatePriceMatrix)) Then
    ReDim TempArr(1 To 3, 1 To 1): TempArr(1, 1) = 0
    RetailHistoryArray = TempArr
    Err.Clear
    On Error GoTo 0
    Exit Property
Else
    ReDim TempArr(1 To 3, 1 To UBound(CBA_COM_DatePriceMatrix))
End If


'tester = CBA_COM_Acode
        For a = LBound(CBA_COM_DatePriceMatrix) To UBound(CBA_COM_DatePriceMatrix)
            TempArr(1, a) = CBA_COM_DatePriceMatrix(a).ScrapeDate
'            If Month(temparr(1, a)) = 9 Then
'            a = a
'            End If
            TempArr(2, a) = CBA_COM_DatePriceMatrix(a).State
            If WhatSection = "prorata" Then
                TempArr(3, a) = CBA_COM_DatePriceMatrix(a).ProRata
            ElseIf WhatSection = "nonpromoprorata" Then TempArr(3, a) = CBA_COM_DatePriceMatrix(a).NonSpecialProRata
            ElseIf WhatSection = "shelf" Then TempArr(3, a) = CBA_COM_DatePriceMatrix(a).Shelf
            ElseIf WhatSection = "was" Then TempArr(3, a) = CBA_COM_DatePriceMatrix(a).Was
            ElseIf WhatSection = "discount" Then TempArr(3, a) = CBA_COM_DatePriceMatrix(a).Discount
            ElseIf WhatSection = "permeasure" Then TempArr(3, a) = CBA_COM_DatePriceMatrix(a).Permeasure
            ElseIf WhatSection = "aldiretail" Then TempArr(3, a) = CBA_COM_DatePriceMatrix(a).A_Retail
            End If
        Next
        RetailHistoryArray = TempArr
    
End Property
Property Get RetailsArray() As Variant()
    Dim TempArr As Variant
    Dim a As Long
    On Error Resume Next
    If IsEmpty(UBound(CBA_COM_DatePriceMatrix)) Then
        ReDim TempArr(1 To 12, 1 To 1): TempArr(1, 1) = 0
        RetailsArray = TempArr
        Err.Clear
        On Error GoTo 0
        Exit Property
    Else
        ReDim TempArr(1 To 12, 1 To UBound(CBA_COM_DatePriceMatrix))
    End If

    For a = LBound(CBA_COM_DatePriceMatrix) To UBound(CBA_COM_DatePriceMatrix)
    '            If Month(CBA_COM_DatePriceMatrix(a).scrapedate) = 10 And CBA_COM_DatePriceMatrix(a).state = "national" Then
    '            a = a
    '            End If
        TempArr(1, a) = CBA_COM_DatePriceMatrix(a).ScrapeDate
        TempArr(2, a) = CBA_COM_DatePriceMatrix(a).State
        TempArr(3, a) = CBA_COM_DatePriceMatrix(a).Shelf
        TempArr(4, a) = CBA_COM_DatePriceMatrix(a).Was
        TempArr(5, a) = CBA_COM_DatePriceMatrix(a).Discount
        TempArr(6, a) = CBA_COM_DatePriceMatrix(a).Permeasure
        TempArr(7, a) = CBA_COM_DatePriceMatrix(a).NonSpecialProRata
        TempArr(8, a) = CBA_COM_DatePriceMatrix(a).ProRata
        TempArr(9, a) = CBA_COM_DatePriceMatrix(a).Special
        TempArr(10, a) = CBA_COM_DatePriceMatrix(a).A_Retail
        TempArr(11, a) = CBA_COM_DatePriceMatrix(a).Diff
        TempArr(12, a) = CBA_COM_DatePriceMatrix(a).Count
    Next
    RetailsArray = TempArr
End Property

Private Function CBA_COM_resolvethiscomp(ByVal CompCode As String, ByVal rowinCBA_COMarr As Long, ByVal DFrom As Date, Optional ByVal RandomWeight As Boolean)
    Dim g As Boolean, l As Boolean, m As Boolean, PK As Boolean, compG As Boolean, compL As Boolean, compPK As Boolean, compM As Boolean, compP As Boolean, compS As Boolean, s As Boolean, P As Boolean
    Dim Gval As Single, lVal As Single, Pkval As Single, Mval As Single, Pval As Single, sVal As Single, compGval As Single, compLval As Single, compPKval As Single, compMval As Single, compPVal As Single, compSval As Single, compGret As Single, compLret As Single, compPKret As Single, compMret As Single, compPret As Single, compSret As Single
    Dim inArray() As Variant
    Dim a As Long, b As Long
    Dim bOutput As Boolean
    Dim varno As String
    
    '1 = Gram , 2 = Volume, 3 = Meters, 4 = Pack.
    'Cancel = False
    
    g = False
    l = False
    PK = False
    m = False
    P = False
    s = False
    compG = False
    compL = False
    compPK = False
    compM = False
    compP = False
    compS = False
    compGret = False
    compLret = False
    compPKret = False
    compMret = False
    compPret = False
    compSret = False
            
            
    inArray = CBA_COM_ADarr
    If CBA_COM_CBIS_PackDecode.CBA_COM_genCBISoutput(True, inArray) = True Then
    
        For a = 1 To UBound(CBA_COM_CBISarrOutput, 2)
            For b = 3 To UBound(CBA_COM_CBISarrOutput, 1)
                If IsEmpty(CBA_COM_CBISarrOutput(b, a)) Then
                    Exit For
                ElseIf LCase(CBA_COM_CBISarrOutput(b, a)) = "pieces" Then
                    PK = True
                    If CBA_COM_CBISarrOutput(b - 1, a) > Pkval Then Pkval = CBA_COM_CBISarrOutput(b - 1, a)
                ElseIf CBA_COM_CBISarrOutput(b, a) = "g" And CBA_COM_CBISarrOutput(1, a) <> 6305 And CBA_COM_CBISarrOutput(1, a) <> 39812 And CBA_COM_CBISarrOutput(1, a) <> 56170 Then
                    g = True
                    Gval = CBA_COM_CBISarrOutput(b - 1, a)
                ElseIf CBA_COM_CBISarrOutput(b, a) = "ml" Then
                    l = True
                    lVal = CBA_COM_CBISarrOutput(b - 1, a)
                ElseIf CBA_COM_CBISarrOutput(b, a) = "mm" Then
                    m = True
                    Mval = CBA_COM_CBISarrOutput(b - 1, a)
                ElseIf LCase(CBA_COM_CBISarrOutput(b, a)) = "pages" Then
                    P = True
                    Pval = CBA_COM_CBISarrOutput(b - 1, a)
                ElseIf LCase(CBA_COM_CBISarrOutput(b, a)) = "sheets" Then
                    s = True
                    sVal = CBA_COM_CBISarrOutput(b - 1, a)
                End If
            Next
            Exit For
        Next
    
    End If
    'Erase CBISarr
    
    
    
    Erase inArray
    ReDim inArray(0 To 14, 0 To 0)
    
    If CBA_COMarr(3, rowinCBA_COMarr) = 0 Then
        For a = 0 To UBound(CBA_COMarr, 2)
            If CBA_COMarr(3, a) <> 0 Then
                rowinCBA_COMarr = a
                Exit For
            End If
        Next
    End If
    inArray(0, 0) = CompCode
    inArray(1, 0) = CBA_COMarr(2, rowinCBA_COMarr)
    inArray(2, 0) = Null
    inArray(3, 0) = CBA_COMarr(3, rowinCBA_COMarr)
    inArray(4, 0) = CBA_COMarr(4, rowinCBA_COMarr)
    inArray(5, 0) = CBA_COMarr(5, rowinCBA_COMarr)
    inArray(6, 0) = CBA_COMarr(6, rowinCBA_COMarr)
    inArray(7, 0) = CBA_COMarr(7, rowinCBA_COMarr)
    inArray(8, 0) = CBA_COMarr(8, rowinCBA_COMarr)
    inArray(9, 0) = Null
    inArray(10, 0) = Null
    inArray(11, 0) = Null
    If CBA_COMarr(9, rowinCBA_COMarr) = 1 Then inArray(12, 0) = "NSW"
    If CBA_COMarr(9, rowinCBA_COMarr) = 2 Then inArray(12, 0) = "QLD"
    If CBA_COMarr(9, rowinCBA_COMarr) = 3 Then inArray(12, 0) = "VIC"
    If CBA_COMarr(9, rowinCBA_COMarr) = 4 Then inArray(12, 0) = "SA"
    If CBA_COMarr(9, rowinCBA_COMarr) = 5 Then inArray(12, 0) = "WA"
    inArray(13, 0) = CBA_COMarr(1, rowinCBA_COMarr)
    inArray(14, 0) = Null
    bOutput = CBA_COM_PackDecode.DecodePack(True, inArray)
    Erase inArray
   
    If InStr(1, LCase(Me.CompProdName), "pack") > 0 _
        Or InStr(1, LCase(Me.CompProdName), "pk") > 0 _
            Or InStr(1, LCase(Me.CompProdName), "kg") > 0 And CompCode <> "3223704P" _
                Or InStr(1, LCase(Me.CompProdName), "g") > 0 _
                    Then
                        CBA_COM_comDepackProdName (CBA_COMarr(3, rowinCBA_COMarr))
    End If
    
    If bOutput = True Then
        For a = 1 To UBound(CBA_COM_arrOutput, 2)
                If IsEmpty(CBA_COM_arrOutput(5, a)) Then
                    Exit For
                ElseIf LCase(CBA_COM_arrOutput(5, a)) = "pieces" Then
                    compPK = True
                    If CBA_COM_arrOutput(5 - 1, a) > compPKval Then
                        compPKval = CBA_COM_arrOutput(5 - 1, a)
                        compPKret = CBA_COM_arrOutput(5 - 2, a)
                    End If
                ElseIf CBA_COM_arrOutput(5, a) = "g" Then
                    If compG = False Then
                        compG = True
                        compGval = CBA_COM_arrOutput(5 - 1, a)
                        compGret = CBA_COM_arrOutput(5 - 2, a)
                    Else
                        varno = compGret / compGval
    '                   '''' If ((compGret / compGval) / (CBA_COM_arrOutput(5 - 2, a) / CBA_COM_arrOutput(5 - 1, a))) / (CBA_COM_arrOutput(5 - 2, a) / CBA_COM_arrOutput(5 - 1, a)) >= 0.05 Then
                        If Len((varno)) < Len(CBA_COM_arrOutput(5 - 2, a) / CBA_COM_arrOutput(5 - 1, a)) Or RandomWeight = True Then
                            If CBA_COM_arrOutput(5 - 1, a) = "1000" And compGval <> "1000" And InStr(1, LCase(CBA_COMarr(2, rowinCBA_COMarr)), "approx") = 0 Then
                                CBA_COM_OrigCompPack = compGval
                            Else
                                CBA_COM_OrigCompPack = 0
                            End If
                            compGval = CBA_COM_arrOutput(5 - 1, a)
                            compGret = CBA_COM_arrOutput(5 - 2, a)
                        End If
                    End If
                ElseIf CBA_COM_arrOutput(5, a) = "ml" And compL = False Then
                    compL = True
                    compLval = CBA_COM_arrOutput(5 - 1, a)
                    compLret = CBA_COM_arrOutput(5 - 2, a)
                ElseIf CBA_COM_arrOutput(5, a) = "mm" Then
                    If compM = False Then
                    compM = True
                    compMval = CBA_COM_arrOutput(5 - 1, a)
                    compMret = CBA_COM_arrOutput(5 - 2, a)
                    Else
                        varno = compMret / compMval
                        If Len((varno)) < Len(CBA_COM_arrOutput(5 - 2, a) / CBA_COM_arrOutput(5 - 1, a)) Or RandomWeight = True Then
                            compMval = CBA_COM_arrOutput(5 - 1, a)
                            compMret = CBA_COM_arrOutput(5 - 2, a)
                        End If
                    End If
                ElseIf LCase(CBA_COM_arrOutput(5, a)) = "pages" Then
                    compP = True
                    compPVal = CBA_COM_arrOutput(5 - 1, a)
                    compPret = CBA_COM_arrOutput(5 - 2, a)
                ElseIf LCase(CBA_COM_arrOutput(5, a)) = "sheets" Then
                    compS = True
                    compSval = CBA_COM_arrOutput(5 - 1, a)
                    compSret = CBA_COM_arrOutput(5 - 2, a)
                End If
        Next
    End If
    
    If g = True And compG = True And Not (InStr(1, LCase(CBA_COM_CBISarrOutput(2, 1)), "nappy") > 0 Or InStr(1, LCase(CBA_COM_CBISarrOutput(2, 1)), "nappies") > 0 Or CBA_COM_CBISarrOutput(1, 1) = 2277) Then
        CBA_COM_Prorataby = "G"
        CBA_COM_Compdiv = compGval
        CBA_COM_Compmult = Gval
        If CSng(NZ(CBA_COMarr(8, rowinCBA_COMarr), 0)) = CSng(compGret) Then CBA_COM_pricetouse = "permeasure" Else CBA_COM_pricetouse = "shelf"
    ElseIf l = True And compL = True Then
        'need to hard code some drinks lines so the pack no and ml value are multiplied
        If CompCode = "8464810P" Or CompCode = "1849293P" Or CompCode = "1849318P" Or CompCode = "1333574P" Or CompCode = "365023" Then
            If compPKval <> 0 Then CBA_COM_Compdiv = compLval * compPKval Else CBA_COM_Compdiv = compLval
        Else
            CBA_COM_Compdiv = compLval
        End If
        CBA_COM_Prorataby = "L"
        CBA_COM_Compmult = lVal
        If CSng(NZ(CBA_COMarr(8, rowinCBA_COMarr), 0)) = CSng(compLret) Then CBA_COM_pricetouse = "permeasure" Else CBA_COM_pricetouse = "shelf"
    ElseIf m = True And compM = True Then
        CBA_COM_Prorataby = "M"
        CBA_COM_Compdiv = compMval
        CBA_COM_Compmult = Mval
        If CSng(NZ(CBA_COMarr(8, rowinCBA_COMarr), 0)) = CSng(compMret) Then CBA_COM_pricetouse = "permeasure" Else CBA_COM_pricetouse = "shelf"
    ElseIf P = True And compPK = True Then
        CBA_COM_Prorataby = "PG"
        CBA_COM_Compdiv = compPVal
        CBA_COM_Compmult = Pval
        If CSng(NZ(CBA_COMarr(8, rowinCBA_COMarr), 0)) = CSng(compPret) Then CBA_COM_pricetouse = "permeasure" Else CBA_COM_pricetouse = "shelf"
    ElseIf s = True And compPK = True Then
        CBA_COM_Prorataby = "SHT"
        CBA_COM_Compdiv = compSval
        CBA_COM_Compmult = sVal
        If CSng(NZ(CBA_COMarr(8, rowinCBA_COMarr), 0)) = CSng(compSret) Then CBA_COM_pricetouse = "permeasure" Else CBA_COM_pricetouse = "shelf"
    ElseIf PK = True And compPK = True Then
        CBA_COM_Prorataby = "PK"
        CBA_COM_Compdiv = compPKval
        CBA_COM_Compmult = Pkval
        If CSng(NZ(CBA_COMarr(8, rowinCBA_COMarr), 0)) = CSng(compPKret) Then CBA_COM_pricetouse = "permeasure" Else CBA_COM_pricetouse = "shelf"
    End If

    
End Function

Function CBA_COM_buildSQL(ByVal SQLType As String, Optional ByVal MMSREG As Long) As Boolean
    Dim strSQL  As String
    Dim WedDates As Scripting.Dictionary
    Dim num As Long, a As Long
    Dim DT As Date
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    Set WedDates = New Scripting.Dictionary
    num = 0
    For DT = DateFrom To DateTo
        If WeekDay(DT, vbWednesday) = 1 Then
            num = num + 1
            WedDates.Add num, DT
        End If
    Next
    
    CBA_ErrTag = "SQL"
    If SQLType = "COM_ClassComProdData" Then
    
        Set CBA_COM_COMRS = New ADODB.Recordset
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) '& "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
        If Me.compet = "C" Then
            strSQL = strSQL & "select distinct convert(date,datescraped), brand + ' ' + name as brandname, packsize, price, priceprevious, pricesaving, cupmeasureunits, cupmeasurevalue" & Chr(10)
            strSQL = strSQL & ", cupsize, case when urlscan_storeseotoken = 'a-national' then 'NAT' else upper(substring(urlscan_storeseotoken,3,charindex('-',urlscan_storeseotoken,3)-3)) end as State, multiitemprice" & Chr(10)
            strSQL = strSQL & "from tools.dbo.com_c_prod where colesproductid in ('" & Me.CompCode & "') " & Chr(10)
            If CBA_COM_ACG = 58 Then
                strSQL = strSQL & "and convert(date,datescraped) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
                strSQL = strSQL & "and convert(date,datescraped) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            Else
                For a = 1 To num
                    If a = 1 Then
                        strSQL = strSQL & "and (convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD")
                        If num = 1 Then strSQL = strSQL & "')" & Chr(10) Else strSQL = strSQL & "'" & Chr(10)
                    ElseIf a = num Then
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "')" & Chr(10)
                    Else
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
                    End If
                Next
            End If
            strSQL = strSQL & "order by convert(date,datescraped)" & Chr(10)
        ElseIf Me.compet = "WW" Then
            strSQL = strSQL & "select distinct convert(date,datescraped), name, packagesize, price, wasprice, savingsAmount, cupmeasureunits, cupmeasurevalue, cupPrice" & Chr(10)
            strSQL = strSQL & ", case when substring(convert(nvarchar(max), s.Postcode),1,1) = '0' then 'NT' else case when substring(convert(nvarchar(max), s.Postcode),1,1) = '1' then 'NSW' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '2' then 'NSW' else case when substring(convert(nvarchar(max), s.Postcode),1,1) = '3' then 'VIC' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '4' then 'QLD' else case when substring(convert(nvarchar(max), s.Postcode),1,1) = '5' then 'SA' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '6' then 'WA' else case when substring(convert(nvarchar(max), s.Postcode),1,1) = '7' then 'TAS' else " & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '8' then 'VIC' else case when substring(convert(nvarchar(max), s.Postcode),1,1) = '9' then 'QLD' end end end end end end end end end end as [State] " & Chr(10)
            strSQL = strSQL & ", case when isEdrSpecial = 1 then 'Special' else null end as isSpecial from tools.dbo.com_w_prod p" & Chr(10)
            strSQL = strSQL & "inner join tools.dbo.com_w_stores s on s.addressid = p.addressid where stockcode in ('" & Me.CompCode & "')" & Chr(10)
            If CBA_COM_ACG = 58 Then
                strSQL = strSQL & "and convert(date,datescraped) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
                strSQL = strSQL & "and convert(date,datescraped) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            Else
                For a = 1 To num
                    If a = 1 Then
                        strSQL = strSQL & "and (convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD")
                        If num = 1 Then strSQL = strSQL & "')" & Chr(10) Else strSQL = strSQL & "'" & Chr(10)
                    ElseIf a = num Then
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "')" & Chr(10)
                    Else
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
                    End If
                Next
            End If
            strSQL = strSQL & "order by convert(date,datescraped)" & Chr(10)
        ElseIf Me.compet = "DM" Then
            strSQL = strSQL & "select distinct convert(date,datescraped), brandname, CPack, price, wasprice, savingsamount, unit,cupmeasurevalue, cupPrice,  State, Special from (" & Chr(10)
            strSQL = strSQL & "select datescraped, brand + ' ' + p.name as brandname, SizeDescription1 as CPack, price1 as price, price1 as wasprice, 0 as savingsamount" & Chr(10)
            strSQL = strSQL & ", null as unit, null as cupmeasurevalue, null as cupPrice" & Chr(10)
            strSQL = strSQL & ", case when s.storescan_storestate = 'ACT' then 'NSW' else storescan_storestate end as State, null as Special" & Chr(10)
            strSQL = strSQL & ", row_number() over (Partition by datescraped, SizeDescription1, case when s.storescan_storestate = 'ACT' then 'NSW' else s.storescan_storestate end order by price1 desc) as row" & Chr(10)
            strSQL = strSQL & "from tools.dbo.com_dm_prod p inner join tools.dbo.com_dm_stores s on p.t0_Stores_StoreScan_StoreNo = s.StoreScan_StoreNo" & Chr(10)
            strSQL = strSQL & "where productid in ('" & Me.CompCode & "') " & Chr(10)
            If Me.CompPacksize = "per bottle" Then
                strSQL = strSQL & "and SizeDescription1 = 'per bottle'" & Chr(10)
            Else
                strSQL = strSQL & "and (charindex('six',lower(SizeDescription1),1) > 0 or charindex('6',lower(SizeDescription1),1) > 0 )"
            End If
            strSQL = strSQL & "and SizeDescription1 = '" & Me.CompPacksize & "'" & Chr(10)
            strSQL = strSQL & "and convert(date,datescraped) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "and convert(date,datescraped) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "union select datescraped, brand + ' ' + p.name, SizeDescription2, price2, price2, 0, null, null, null" & Chr(10)
            strSQL = strSQL & ", case when s.storescan_storestate = 'ACT' then 'NSW' else storescan_storestate end as State, null" & Chr(10)
            strSQL = strSQL & ", row_number() over (Partition by datescraped, SizeDescription2, case when s.storescan_storestate = 'ACT' then 'NSW' else s.storescan_storestate end order by price2 desc) as row" & Chr(10)
            strSQL = strSQL & "from tools.dbo.com_dm_prod p inner join tools.dbo.com_dm_stores s on p.t0_Stores_StoreScan_StoreNo = s.StoreScan_StoreNo" & Chr(10)
            strSQL = strSQL & "where productid in ('" & Me.CompCode & "') " & Chr(10)
            If Me.CompPacksize = "per bottle" Then
                strSQL = strSQL & "and SizeDescription2 = 'per bottle'" & Chr(10)
            Else
                strSQL = strSQL & "and (charindex('six',lower(SizeDescription2),1) > 0 or charindex('6',lower(SizeDescription2),1) > 0 )"
            End If
            If CBA_COM_ACG = 58 Then
                strSQL = strSQL & "and convert(date,datescraped) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
                strSQL = strSQL & "and convert(date,datescraped) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            Else
                For a = 1 To num
                    If a = 1 Then
                        strSQL = strSQL & "and (convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD")
                        If num = 1 Then strSQL = strSQL & "')" & Chr(10) Else strSQL = strSQL & "'" & Chr(10)
                    ElseIf a = num Then
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "')" & Chr(10)
                    Else
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
                    End If
                Next
            End If
            strSQL = strSQL & "union select datescraped, brand + ' ' + p.name, SizeDescription3, price3, price3, 0, null, null, null" & Chr(10)
            strSQL = strSQL & ", case when s.storescan_storestate = 'ACT' then 'NSW' else storescan_storestate end as State, null" & Chr(10)
            strSQL = strSQL & ", row_number() over (Partition by datescraped, SizeDescription3, case when s.storescan_storestate = 'ACT' then 'NSW' else s.storescan_storestate end order by price3 desc) as row" & Chr(10)
            strSQL = strSQL & "from tools.dbo.com_dm_prod p inner join tools.dbo.com_dm_stores s on p.t0_Stores_StoreScan_StoreNo = s.StoreScan_StoreNo" & Chr(10)
            strSQL = strSQL & "where  productid in ('" & Me.CompCode & "') " & Chr(10)
            If Me.CompPacksize = "per bottle" Then
                strSQL = strSQL & "and SizeDescription3 = 'per bottle'" & Chr(10)
            Else
                strSQL = strSQL & "and (charindex('six',lower(SizeDescription3),1) > 0 or charindex('6',lower(SizeDescription3),1) > 0 )"
            End If
            If CBA_COM_ACG = 58 Then
                strSQL = strSQL & "and convert(date,datescraped) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
                strSQL = strSQL & "and convert(date,datescraped) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            Else
                For a = 1 To num
                    If a = 1 Then
                        strSQL = strSQL & "and (convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD")
                        If num = 1 Then strSQL = strSQL & "')" & Chr(10) Else strSQL = strSQL & "'" & Chr(10)
                    ElseIf a = num Then
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "')" & Chr(10)
                    Else
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
                    End If
                Next
            End If
            strSQL = strSQL & ") a  where a.row = 1 order by convert(date,datescraped),brandname, state" & Chr(10)
        ElseIf Me.compet = "FC" Then
            strSQL = strSQL & "select DateScraped, StoreNo, ProductId, Description,    Brand, PriceEach,PriceAcrossAny,PriceRange" & Chr(10)
            strSQL = strSQL & "into #PROD from tools.dbo.com_fc_prod where productid in (" & Me.CompCode & ") and convert(date,datescraped) >= dateadd(D,-7,convert(date,'" & Format(Me.DateFrom, "YYYY-MM-DD") & "')) and convert(date,datescraped) <= convert(date,'" & Format(Me.DateTo, "YYYY-MM-DD") & "')" & Chr(10)
            strSQL = strSQL & "and Storeno in (7023,7297,7024,7242,7207,8662,7200,7211,7210,7252,7262,7232,7070,7062,7064,7256,7268,7237,7236,7285)" & Chr(10)
            strSQL = strSQL & "select distinct convert(date,fc.datescraped), fc.Brand + ' ' + fc.Description as brandname, 'each', fc.priceeach, isnull(fc2.PriceEach, fc.priceeach)" & Chr(10)
            strSQL = strSQL & ", isnull(fc2.PriceEach - fc.priceeach,0), 'each',1, fc.priceeach, case when fs.State = 'ACT' then 'NSW' else fs.State end as State, fc.PriceAcrossAny" & Chr(10)
            strSQL = strSQL & "from #PROD fc left join tools.dbo.Com_FC_Stores fs on fs.StoreNo = fc.StoreNo" & Chr(10)
            strSQL = strSQL & "left join #PROD fc2 on fc2.productid = fc.productid and fc2.storeno = fc.storeno and dateadd(D,-7,fc2.DateScraped) = fc.DateScraped " & Chr(10)
            strSQL = strSQL & "where fc.productid in ('" & Me.CompCode & "') " & Chr(10)
            If CBA_COM_ACG = 58 Then
                strSQL = strSQL & "and convert(date,datescraped) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
                strSQL = strSQL & "and convert(date,datescraped) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            Else
                For a = 1 To num
                    If a = 1 Then
                        strSQL = strSQL & "and (convert(date,fc.datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD")
                        If num = 1 Then strSQL = strSQL & "')" & Chr(10) Else strSQL = strSQL & "'" & Chr(10)
                    ElseIf a = num Then
                        strSQL = strSQL & "or convert(date,fc.datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "')" & Chr(10)
                    Else
                        strSQL = strSQL & "or convert(date,fc.datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
                    End If
                Next
            End If
            strSQL = strSQL & "order by convert(date,fc.datescraped) drop table #PROD" & Chr(10)
        End If
        'Debug.Print strSQL
    
            'Debug.Print strSQL_TBLNAME
        CBA_COM_COMRS.Open strSQL, CBA_COM_COMCN ', , , "adAsyncExecute"
        If CBA_COM_COMRS.EOF Then
            ReDim CBA_COMarr(0, 0)
            CBA_COMarr(0, 0) = 0
        Else
            CBA_COMarr = CBA_COM_COMRS.GetRows()
        End If
        CBA_COM_COMRS.Close
        Set CBA_COM_COMRS = Nothing
    
    End If
    
    
    If SQLType = "COM_ClassComProdDataMode" Then
        Set CBA_COM_COMRS = New ADODB.Recordset
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) '& "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
        strSQL = strSQL & "select datescraped, brandname, packsize, price, wasprice, pricesaving, cupunit, cupvalue, cupPrice, [State], flags from (" & Chr(10)
        strSQL = strSQL & "select *, row_number() over (partition by datescraped, [state] order by datescraped, [state], cnt desc) as row from (" & Chr(10)
        strSQL = strSQL & "Select distinct *, count(price) over (partition by datescraped, [State], price) as cnt  from (" & Chr(10)
        If Me.compet = "C" Then
            strSQL = strSQL & "select convert(date,datescraped) as datescraped, brand + ' ' + name as brandname, packsize, price, priceprevious as wasprice, pricesaving, cupmeasureunits as cupunit, cupmeasurevalue as cupvalue" & Chr(10)
            strSQL = strSQL & ", cupsize as cupPrice, case when urlscan_storeseotoken = 'a-national' then 'NAT' else upper(substring(urlscan_storeseotoken,3,charindex('-',urlscan_storeseotoken,3)-3)) end as [State], multiitemprice as flags" & Chr(10)
            strSQL = strSQL & "from tools.dbo.com_c_prod where colesproductid in ('" & Me.CompCode & "') " & Chr(10)
            If CBA_COM_ACG = 58 Then
                strSQL = strSQL & "and convert(date,datescraped) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
                strSQL = strSQL & "and convert(date,datescraped) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            Else
                For a = 1 To num
                    If a = 1 Then
                        strSQL = strSQL & "and (convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
                        If num = 1 Then strSQL = strSQL & ")"
                    ElseIf a = num Then
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "')" & Chr(10)
                    Else
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
                    End If
                Next
            End If
    
        ElseIf Me.compet = "WW" Then
            strSQL = strSQL & "select convert(date,datescraped) as datescraped, name as brandname, packagesize as packsize, price, wasprice, savingsAmount as pricesaving, cupmeasureunits as cupunit, cupmeasurevalue as cupvalue, cupPrice" & Chr(10)
            strSQL = strSQL & ", case when substring(convert(nvarchar(max), s.Postcode),1,1) = '0' then 'NT' else case when substring(convert(nvarchar(max), s.Postcode),1,1) = '1' then 'NSW' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '2' then 'NSW' else case when substring(convert(nvarchar(max), s.Postcode),1,1) = '3' then 'VIC' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '4' then 'QLD' else case when substring(convert(nvarchar(max), s.Postcode),1,1) = '5' then 'SA' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '6' then 'WA' else case when substring(convert(nvarchar(max), s.Postcode),1,1) = '7' then 'TAS' else " & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '8' then 'VIC' else case when substring(convert(nvarchar(max), s.Postcode),1,1) = '9' then 'QLD' end end end end end end end end end end as [State] " & Chr(10)
            strSQL = strSQL & ", case when isEdrSpecial = 1 then 'Special' else null end as flags from tools.dbo.com_w_prod p" & Chr(10)
            strSQL = strSQL & "inner join tools.dbo.com_w_stores s on s.addressid = p.addressid where stockcode in ('" & Me.CompCode & "')" & Chr(10)
            If CBA_COM_ACG = 58 Then
                strSQL = strSQL & "and convert(date,datescraped) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
                strSQL = strSQL & "and convert(date,datescraped) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            Else
                For a = 1 To num
                    If a = 1 Then
                        strSQL = strSQL & "and (convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
                        If num = 1 Then strSQL = strSQL & ")"
                    ElseIf a = num Then
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "')" & Chr(10)
                    Else
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
                    End If
                Next
            End If
        ElseIf Me.compet = "DM" Then
            strSQL = strSQL & "select convert(date,datescraped) as datescraped, brandname as brandname, CPack as packsize, price, wasprice, savingsamount as pricesaving, unit as cupunit ,cupmeasurevalue as cupvalue, cupPrice,  [State], Special as flags from (" & Chr(10)
            strSQL = strSQL & "select datescraped, brand + ' ' + p.name as brandname, SizeDescription1 as CPack, price1 as price, price1 as wasprice, 0 as savingsamount" & Chr(10)
            strSQL = strSQL & ", null as unit, null as cupmeasurevalue, null as cupPrice" & Chr(10)
            strSQL = strSQL & ", case when s.storescan_storestate = 'ACT' then 'NSW' else storescan_storestate end as [State], null as Special" & Chr(10)
            strSQL = strSQL & ", row_number() over (Partition by datescraped, SizeDescription1, case when s.storescan_storestate = 'ACT' then 'NSW' else s.storescan_storestate end order by price1 desc) as row" & Chr(10)
            strSQL = strSQL & "from tools.dbo.com_dm_prod p inner join tools.dbo.com_dm_stores s on p.t0_Stores_StoreScan_StoreNo = s.StoreScan_StoreNo" & Chr(10)
            strSQL = strSQL & "where productid in ('" & Me.CompCode & "') " & Chr(10)
            If Me.CompPacksize = "per bottle" Then
                strSQL = strSQL & "and SizeDescription1 = 'per bottle'" & Chr(10)
            Else
                strSQL = strSQL & "and (charindex('six',lower(SizeDescription1),1) > 0 or charindex('6',lower(SizeDescription1),1) > 0 )"
            End If
            strSQL = strSQL & "and convert(date,datescraped) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "and convert(date,datescraped) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "union select datescraped, brand + ' ' + p.name, SizeDescription2, price2, price2, 0, null, null, null" & Chr(10)
            strSQL = strSQL & ", case when s.storescan_storestate = 'ACT' then 'NSW' else storescan_storestate end as [State], null" & Chr(10)
            strSQL = strSQL & ", row_number() over (Partition by datescraped, SizeDescription2, case when s.storescan_storestate = 'ACT' then 'NSW' else s.storescan_storestate end order by price2 desc) as row" & Chr(10)
            strSQL = strSQL & "from tools.dbo.com_dm_prod p inner join tools.dbo.com_dm_stores s on p.t0_Stores_StoreScan_StoreNo = s.StoreScan_StoreNo" & Chr(10)
            strSQL = strSQL & "where productid in ('" & Me.CompCode & "') " & Chr(10)
            If Me.CompPacksize = "per bottle" Then
                strSQL = strSQL & "and SizeDescription2 = 'per bottle'" & Chr(10)
            Else
                strSQL = strSQL & "and (charindex('six',lower(SizeDescription2),1) > 0 or charindex('6',lower(SizeDescription2),1) > 0 )"
            End If
            strSQL = strSQL & "and convert(date,datescraped) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "and convert(date,datescraped) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "union select datescraped, brand + ' ' + p.name, SizeDescription3, price3, price3, 0, null, null, null" & Chr(10)
            strSQL = strSQL & ", case when s.storescan_storestate = 'ACT' then 'NSW' else storescan_storestate end as State, null" & Chr(10)
            strSQL = strSQL & ", row_number() over (Partition by datescraped, SizeDescription3, case when s.storescan_storestate = 'ACT' then 'NSW' else s.storescan_storestate end order by price3 desc) as row" & Chr(10)
            strSQL = strSQL & "from tools.dbo.com_dm_prod p inner join tools.dbo.com_dm_stores s on p.t0_Stores_StoreScan_StoreNo = s.StoreScan_StoreNo" & Chr(10)
            strSQL = strSQL & "where  productid in ('" & Me.CompCode & "') " & Chr(10)
            If Me.CompPacksize = "per bottle" Then
                strSQL = strSQL & "and SizeDescription3 = 'per bottle'" & Chr(10)
            Else
                strSQL = strSQL & "and (charindex('six',lower(SizeDescription3),1) > 0 or charindex('6',lower(SizeDescription3),1) > 0 )"
            End If
            strSQL = strSQL & "and convert(date,datescraped) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
            If CBA_COM_ACG = 58 Then
                strSQL = strSQL & "and convert(date,datescraped) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
                strSQL = strSQL & "and convert(date,datescraped) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            Else
                For a = 1 To num
                    If a = 1 Then
                        strSQL = strSQL & "and (convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
                        If num = 1 Then strSQL = strSQL & ")"
                    ElseIf a = num Then
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "')" & Chr(10)
                    Else
                        strSQL = strSQL & "or convert(date,datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
                    End If
                Next
            End If
    
            strSQL = strSQL & ") a  where a.row = 1 " & Chr(10)
        End If
        strSQL = strSQL & ") a ) b ) c where c.row = 1 order by convert(date,datescraped),brandname, [state], price"
        
        If Me.compet = "FC" Then
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) '& "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            strSQL = strSQL & "select DateScraped, StoreNo, ProductId, Description,    Brand, PriceEach,PriceAcrossAny,PriceRange" & Chr(10)
            strSQL = strSQL & "into #PROD from tools.dbo.com_fc_prod where productid in (" & Me.CompCode & ") and convert(date,datescraped) >= dateadd(D,-7,convert(date,'" & Format(Me.DateFrom, "YYYY-MM-DD") & "')) and convert(date,datescraped) <= convert(date,'" & Format(Me.DateTo, "YYYY-MM-DD") & "')" & Chr(10)
            strSQL = strSQL & "and Storeno in (7023,7297,7024,7242,7207,8662,7200,7211,7210,7252,7262,7232,7070,7062,7064,7256,7268,7237,7236,7285)" & Chr(10)
            strSQL = strSQL & "select distinct convert(date,fc.datescraped), fc.Brand + ' ' + fc.Description as brandname, 'each', fc.priceeach, isnull(fc2.PriceEach, fc.priceeach)" & Chr(10)
            strSQL = strSQL & ", isnull(fc2.PriceEach - fc.priceeach,0), 'each',1, fc.priceeach, case when fs.State = 'ACT' then 'NSW' else fs.State end as State, fc.PriceAcrossAny" & Chr(10)
            strSQL = strSQL & "from #PROD fc left join tools.dbo.Com_FC_Stores fs on fs.StoreNo = fc.StoreNo" & Chr(10)
            strSQL = strSQL & "left join #PROD fc2 on fc2.productid = fc.productid and fc2.storeno = fc.storeno and dateadd(D,-7,fc2.DateScraped) = fc.DateScraped " & Chr(10)
            strSQL = strSQL & "where fc.productid in ('" & Me.CompCode & "') " & Chr(10)
            If CBA_COM_ACG = 58 Then
                strSQL = strSQL & "and convert(date,fc.datescraped) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
                strSQL = strSQL & "and convert(date,fc.datescraped) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            Else
                For a = 1 To num
                    If a = 1 Then
                        strSQL = strSQL & "and (convert(date,fc.datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
                        If num = 1 Then strSQL = strSQL & ")"
                    ElseIf a = num Then
                        strSQL = strSQL & "or convert(date,fc.datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "')" & Chr(10)
                    Else
                        strSQL = strSQL & "or convert(date,fc.datescraped) = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
                    End If
                Next
            End If
    
            strSQL = strSQL & "order by convert(date,fc.datescraped) drop table #PROD" & Chr(10)
        End If
        CBA_COM_COMRS.Open strSQL, CBA_COM_COMCN ', , , "adAsyncExecute"
        If CBA_COM_COMRS.EOF Then
            ReDim CBA_COMarr(0, 0)
            CBA_COMarr(0, 0) = 0
        Else
            CBA_COMarr = CBA_COM_COMRS.GetRows()
        End If
        CBA_COM_COMRS.Close
        Set CBA_COM_COMRS = Nothing
    End If
    
    
    If SQLType = "CBIS_ClassComAldiData" Then
        Set CBA_COM_CBISRS = New ADODB.Recordset
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10)
        strSQL = strSQL & "select distinct p.productcode, p.cgno, p.scgno, p.productclass, p.description, p.packsize, p.unitcode_wght, p.unitcode_Vol, p.unitcode_Pack," & Chr(10)
        strSQL = strSQL & "p.israndomweight, p.iscoldstorage, p.isfrozen, p.isfood, p.isdirect, p.weightvolume, p.contents, p.weight, p.weightKG," & Chr(10)
        strSQL = strSQL & "p.volume , p.pricecrddesc, p.pricecrddoubleinf, unitpricebase  from cbis599p.dbo.product as P where p.productcode in (" & Me.AldiPCode & ")" & Chr(10)
        'Debug.Print strSQL
        CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN ', , , "adAsyncExecute"
        If CBA_COM_CBISRS.EOF Then
            ReDim CBA_CBISarr(0, 0)
            CBA_CBISarr(0, 0) = 0
        Else
            CBA_CBISarr = CBA_COM_CBISRS.GetRows()
        End If
        CBA_COM_CBISRS.Close
        Set CBA_COM_CBISRS = Nothing
    End If
    
    If SQLType = "CBIS_ClassComAldiRetails" Then
        Set CBA_COM_CBISRS = New ADODB.Recordset
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) '& "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
        strSQL = strSQL & "select productcode, State, validfrom, validto, Retail, ROW_NUMBER() over (Partition by productcode, validfrom order by Retail) as row from (" & Chr(10)
        strSQL = strSQL & "select productcode, case when divno in (501,504) then 'NSW' else case when divno in (502,505) then 'VIC' else case when divno in (503,506) " & Chr(10)
        strSQL = strSQL & "then 'QLD' else case when divno = 507 then 'SA' else 'WA' end end end end as State  " & Chr(10)
        strSQL = strSQL & ", row_number() over (Partition by case when divno in (501,504) then 'NSW' else case when divno in (502,505) then 'VIC' else case when divno in (503,506) " & Chr(10)
        strSQL = strSQL & "then 'QLD' else case when divno = 507 then 'SA' else 'WA' end end end end , validfrom order by retail) as row" & Chr(10)
        strSQL = strSQL & ", convert(date,validfrom) as validfrom , isnull(convert(date,validto),convert(date,getdate())) as validto,  Retail from cbis599p.dbo.divretail" & Chr(10)
        strSQL = strSQL & "where isnull(convert(date,validto),convert(date,getdate())) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
        strSQL = strSQL & "and isnull(convert(date,validfrom),convert(date,getdate())) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
        strSQL = strSQL & "and productcode = " & Me.AldiPCode & Chr(10)
        strSQL = strSQL & ") a where a.row = 1 and retail <> 0" & Chr(10)
        strSQL = strSQL & "order by productcode, state,validfrom" & Chr(10)
        'Debug.Print strSQL
        CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN
        If CBA_COM_CBISRS.EOF Then
            ReDim CBA_CBISarr(0, 0)
            CBA_CBISarr(0, 0) = 0
        Else
            CBA_CBISarr = CBA_COM_CBISRS.GetRows()
        End If
        CBA_COM_CBISRS.Close
        Set CBA_COM_CBISRS = Nothing
    End If
    
    If SQLType = "CBIS_ClassComAldiRetailsFuture" Then
        Set CBA_COM_CBISRS = New ADODB.Recordset
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) '& "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
        strSQL = strSQL & "select productcode, State, validfrom, validto, Retail, ROW_NUMBER() over (Partition by productcode, validfrom order by Retail) as row from (" & Chr(10)
        strSQL = strSQL & "select productcode, State, validfrom, validto, Retail, ROW_NUMBER() over (Partition by productcode, state order by validfrom desc) as row from (" & Chr(10)
        strSQL = strSQL & "select productcode, case when divno in (501,504) then 'NSW' else case when divno in (502,505) then 'VIC' else case when divno in (503,506) " & Chr(10)
        strSQL = strSQL & "then 'QLD' else case when divno = 507 then 'SA' else 'WA' end end end end as State  " & Chr(10)
        strSQL = strSQL & ", row_number() over (Partition by case when divno in (501,504) then 'NSW' else case when divno in (502,505) then 'VIC' else case when divno in (503,506) " & Chr(10)
        strSQL = strSQL & "then 'QLD' else case when divno = 507 then 'SA' else 'WA' end end end end , validfrom order by retail) as row" & Chr(10)
        strSQL = strSQL & ", convert(date,validfrom) as validfrom , isnull(convert(date,validto),convert(date,getdate())) as validto,  Retail from cbis599p.dbo.divretail" & Chr(10)
        strSQL = strSQL & "where productcode = " & Me.AldiPCode & Chr(10)
        strSQL = strSQL & ") a where a.row = 1 and retail <> 0" & Chr(10)
        strSQL = strSQL & ") b where b.row = 1 " & Chr(10)
        strSQL = strSQL & "order by productcode, state,validfrom" & Chr(10)
    '  '  Debug.Print strSQL
        CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN
        If CBA_COM_CBISRS.EOF Then
            ReDim CBA_CBISarr(0, 0)
            CBA_CBISarr(0, 0) = 0
        Else
            CBA_CBISarr = CBA_COM_CBISRS.GetRows()
        End If
        CBA_COM_CBISRS.Close
        Set CBA_COM_CBISRS = Nothing
    End If
    
    If SQLType = "CBIS_ClassComAldiRetailsNational" Then
    '    If Me.AldiPCode = 9372 Then
    '    a = a
    '    End If
        Set CBA_COM_CBISRS = New ADODB.Recordset
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) '& "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
        strSQL = strSQL & "select productcode, State, validfrom, validto, Retail from (" & Chr(10)
        strSQL = strSQL & "select productcode, State, validfrom, validto, Retail, ROW_NUMBER() over (partition by productcode, validfrom order by Retail) as row from (" & Chr(10)
        strSQL = strSQL & "select productcode, case when divno in (501,504) then 'NSW' else case when divno in (502,505) then 'VIC' else case when divno in (503,506)" & Chr(10)
        strSQL = strSQL & "then 'QLD' else case when divno = 507 then 'SA' else 'WA' end end end end as State" & Chr(10)
        strSQL = strSQL & ", row_number() over (Partition by case when divno in (501,504) then 'NSW' else case when divno in (502,505) then 'VIC' else case when divno in (503,506)" & Chr(10)
        strSQL = strSQL & "then 'QLD' else case when divno = 507 then 'SA' else 'WA' end end end end , validfrom order by retail) as row" & Chr(10)
        strSQL = strSQL & ", convert(date,validfrom) as validfrom , isnull(convert(date,validto),convert(date,getdate())) as validto,  Retail from cbis599p.dbo.divretail" & Chr(10)
        strSQL = strSQL & "where isnull(convert(date,validto),convert(date,getdate())) >= '" & Format(Me.DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
        strSQL = strSQL & "and isnull(convert(date,validfrom),convert(date,getdate())) <= '" & Format(Me.DateTo, "YYYY-MM-DD") & "'" & Chr(10)
        strSQL = strSQL & "and productcode = " & Me.AldiPCode & Chr(10)
        strSQL = strSQL & ") a where a.row = 1 and retail <> 0" & Chr(10)
        strSQL = strSQL & ") b where b.row = 1 and retail <> 0" & Chr(10)
        strSQL = strSQL & "order by productcode, state,validfrom" & Chr(10)
        'Debug.Print strSQL
        CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN
        If CBA_COM_CBISRS.EOF Then
            ReDim CBA_CBISarr(0, 0)
            CBA_CBISarr(0, 0) = 0
        Else
            CBA_CBISarr = CBA_COM_CBISRS.GetRows()
        End If
        CBA_COM_CBISRS.Close
        Set CBA_COM_CBISRS = Nothing
    End If

Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_COM_buildSQL", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & strSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Function CBA_COM_comDepackProdName(ByVal Shelf As Single)
    Dim colPacks As Collection
    Dim packtype As Variant, a As Long, b As Long, topofCaO As String, thisprodname As String, thisprodno As String
    Dim cntspace As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    Set colPacks = New Collection
    colPacks.Add "pack"
    colPacks.Add "pk"
    colPacks.Add "kg"
    colPacks.Add "g"
    thisprodname = Me.CompProdName
    thisprodno = Me.CompCode
    cntspace = 0
    For Each packtype In colPacks
        If InStr(1, LCase(thisprodname), packtype) > 0 Then
            a = Len(thisprodname) - 2
            Do
                a = a - 1
                'strTEST = LCase(Mid(thisprodname, a, 4))
                If LCase(Mid(thisprodname, a, Len(packtype))) = packtype Or LCase(Right(thisprodname, Len(packtype))) = packtype Then
                    If LCase(Right(thisprodname, Len(packtype))) = packtype Then
                        a = Len(thisprodname) - 1
                        b = Len(thisprodname) - 2
                    Else
                        b = a
                    End If
                    If b = 1 Then Exit Do
                    Do
                        b = b - 1
                        If Mid(thisprodname, b, 1) = " " Then cntspace = cntspace + 1
                        If Mid(thisprodname, b, 1) = " " Or IsNumeric(Mid(thisprodname, b, 1)) And cntspace < 2 Then
                        Else
                            'lkjlkjl = Mid(thisprodname, b + 1, a - (b + 1))
                            If IsNumeric(Trim(Mid(thisprodname, b + 1, a - (b + 1)))) = True Then
                                On Error Resume Next
                                If UBound(CBA_COM_arrOutput, 2) > 0 Then
                                    On Error GoTo Err_Routine
                                    topofCaO = (UBound(CBA_COM_arrOutput, 2) + 1)
                                    ReDim Preserve CBA_COM_arrOutput(1 To 6, 1 To topofCaO)
                                Else
                                    Err.Clear
                                    On Error GoTo Err_Routine
                                    topofCaO = 1
                                    ReDim CBA_COM_arrOutput(1 To 5, 1 To topofCaO)
                                End If
                            
                            
                                'strTEST = Trim(Mid(thisprodname, b + 1, a - (b + 1)))
                                CBA_COM_arrOutput(1, topofCaO) = thisprodno
                                CBA_COM_arrOutput(3, topofCaO) = Shelf
                                If packtype = "pack" Or packtype = "pk" Then
                                    CBA_COM_arrOutput(4, topofCaO) = Trim(Mid(thisprodname, b + 1, a - (b + 1)))
                                    CBA_COM_arrOutput(5, topofCaO) = "Pieces"
                                ElseIf packtype = "kg" Then
                                    CBA_COM_arrOutput(4, topofCaO) = Trim(Mid(thisprodname, b + 1, a - (b + 1))) * 1000
                                    CBA_COM_arrOutput(5, topofCaO) = "g"
                                ElseIf packtype = "g" Then
                                    CBA_COM_arrOutput(4, topofCaO) = Trim(Mid(thisprodname, b + 1, a - (b + 1)))
                                    CBA_COM_arrOutput(5, topofCaO) = "g"
                                End If
                                
                            End If
                            Exit Do
                        End If
                    Loop Until b = 1
                    Exit Do
                End If
            Loop Until a = 1
            
        End If
    Next
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("CBA_COM_comDepackProdName", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    

End Function
