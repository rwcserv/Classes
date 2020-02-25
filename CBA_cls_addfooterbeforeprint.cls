VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_cls_addfooterbeforeprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents appevent As Application
Attribute appevent.VB_VarHelpID = -1

Private Sub appevent_WorkbookBeforePrint(ByVal WB As Workbook, Cancel As Boolean)
    Dim lRet As Long
    Dim strFileName As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    If WB.ActiveSheet.PageSetup.LeftFooter = "" Then
        If ActiveSheet.CodeName = "CBA_TEN_WBK" Then
            If strFileName = "" Then strFileName = ActiveSheet.Cells(1, 10).Value & "-" & ActiveSheet.Cells(1, 14).Value
        Else
            If strFileName = "" Then strFileName = Application.ActiveWorkbook.FullName
        End If
        Select Case Application.WorksheetFunction.RoundDown(Len(strFileName) / 125, 0)
            Case 0
                ActiveSheet.PageSetup.LeftFooter = "&9CORP BUYING, per: " & Format(Date, "DD/MM/YYYY") & Chr(10) & strFileName
            Case 1
                ActiveSheet.PageSetup.LeftFooter = "&9CORP BUYING, per: " & Format(Date, "DD/MM/YYYY") & Chr(10) & Mid(strFileName, 1, 125) & Chr(10) & Mid(strFileName, 125, 125)
            Case 1
                ActiveSheet.PageSetup.LeftFooter = "&9CORP BUYING, per: " & Format(Date, "DD/MM/YYYY") & Chr(10) & Mid(strFileName, 1, 125) & Chr(10) & Mid(strFileName, 125, 125) & Chr(10) & Mid(strFileName, 250, 999999)
        End Select
        ActiveSheet.PageSetup.RightFooter = "&P of &N"
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-appevent_WorkbookBeforePrint", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    If Err.Number = 1004 Then GoTo Exit_Routine '' Resume Next                                                      ' #RW 200108 - uncommented and pointed to Exit_Routine
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
