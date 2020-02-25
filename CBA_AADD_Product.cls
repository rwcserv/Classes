VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_AADD_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Data() As Single
Private CG As Byte
Private scg As Byte
Private CGDesc As String
Private SCGDesc As String
Private productcode As Long
Private ProductDescription As String
Private PromoStart As Date
Private PromoEnd As Date

Function BuildIt(ByVal PCode As Long, ByVal YearStart As Byte, ByVal monthstart As Byte, _
                 ByVal daystart As Byte, ByVal dayduration As Integer, ByVal ProdDesc As String) As Boolean
    Dim CGSCG() As String
    Dim a As Long, vDiv As Integer, totrow As Long
    Dim RS As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    If IsNumeric(PCode) And PCode <> 0 Then
        productcode = PCode
        ProductDescription = ProdDesc
        PromoStart = DateSerial(YearStart, monthstart, daystart)
        PromoEnd = DateAdd("D", dayduration, DateSerial(YearStart, monthstart, daystart))
        If CBA_SQL_Queries.CBA_GenPullSQL("CBIS_CG", , , productcode) = True Then
            CG = CBA_CBISarr(0, 0)
            scg = CBA_CBISarr(1, 0)
            Erase CBA_CBISarr
            CGSCG = CBA_AADD_Runtime.CBA_AADD_getCGSCGList
            For a = LBound(CGSCG, 2) To UBound(CGSCG, 2)
                If Mid(CGSCG(0, a), 1, InStr(1, CGSCG(0, a), " - ")) = CG Then
                    CGDesc = CGSCG(0, a)
                    SCGDesc = CGSCG(1, a)
                    Erase CGSCG
                    Exit For
                End If
            Next
            For vDiv = 501 To 509
                If vDiv <> 508 Then
                    Set RS = New ADODB.Recordset
                    strSQL = "select storeno, salesdate, retail, quantity from purchase.dbo.storesales" & Chr(10)
                    strSQL = strSQL & "where salesdate >= '" & Format(DateSerial(YearStart, monthstart, daystart), "YYYY-MM-DD") & "'" & Chr(10)
                    strSQL = strSQL & "and salesdate <= '" & Format(DateAdd("D", dayduration, DateSerial(YearStart, monthstart, daystart)), "YYYY-MM-DD") & "'" & Chr(10)
                    strSQL = strSQL & "and productcode = " & productcode & Chr(10)
                    strSQL = strSQL & "order by storeno, salesdate" & Chr(10)
                    RS.Open strSQL, CBA_AADD_MMSCN(vDiv)
                    Do Until RS.EOF
                        totrow = totrow + 1
                        ReDim Preserve Data(1 To 5, 1 To totrow)
                        For a = 1 To 5
                            If a = 1 Then
                                Data(a, totrow) = vDiv
                            Else
                                Data(a, totrow) = RS.Fields(a - 2)
                            End If
                        Next
                        RS.MoveNext
                    Loop
                    Set RS = Nothing
                End If
            Next
            BuildIt = True
        End If
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-BuildIt", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Property Get CGno()
    CGno = CG
End Property
Property Get SCGno()
    SCGno = scg
End Property
Property Get CGDescription()
    CGDescription = CGDesc
End Property
Property Get SCGDescription()
    SCGDescription = SCGDesc
End Property
Property Get Product()
    Product = productcode
End Property
Property Get getdata()
    getdata = Data
End Property
Property Get ProductDesc()
    ProductDesc = ProductDescription
End Property
Property Get PromoStartDate()
    PromoStartDate = PromoStart
End Property
Property Get PromoEndDate()
    PromoEndDate = PromoEnd
End Property

