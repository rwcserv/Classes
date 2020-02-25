VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCBA_Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit             ' cCBA_Document
'Private psdUTs As Scripting.Dictionary         ' Dictionary containing UTObjects(Key: Top "00000" & Left "00000") Item: cCBA_UDT-Copied from the ETL Class when the document is created
Private psDocumentName As String                ' Name of the document
Private plDocumentID As Long                    ' ID of Document
Private peDocumentType As e_DocuType            ' Type of document
Private psLinkDoc As String                     ' Workbook.Name & Worksheet.Name
Private pwbWorkbookRef As Workbook
Private pwsWorksheetRef As Worksheet
Private pbIsWorkbookSet As Boolean
Private pbIsWorksheetSet As Boolean
Private pbIsSaved As Boolean
Private pbIsLinked As Boolean
Private pbIsRendered As Boolean
Private pbReadyForRender As Boolean
Private pcUDT_Data As cCBA_UDT_ETL                          ' This is already created by the calling class or module
Private piDocDataObject As iCBA_Doc_Data
Private Const cAFCOL = 60


' Is a virtual document. All the information that exists on a screen, exists in virtual fashion in this object.
' The object can be linked to a workbook and rendered from a specific spot at any time. UDT objects are stored in an ETL Class.
' On creation, only the DocumentName and DocumentType are required, however this class cannot render unless the necessary information is available.
' If UDT data is received, all top values may be re-allocated and the dictionary re-indexed.  Any VBA Code in any of the worksheets, workbooks or created modules are received in a dictionary.

Public Function CloneDoc(ByVal DocObject As cCBA_Document) As Boolean ' This will not work properly. deep cloning is needed.. especially on those UDT and category Objects
    On Error GoTo Err_Routine
    CBA_Error = ""
    sDocumentName = DocObject.sDocumentName
    lDocumentID = DocObject.lDocumentID
    eDocumentType = DocObject.eDocumentType
    sLinkDoc = DocObject.sLinkDoc
    WorkbookRef = DocObject.WorkbookRef
    WorksheetRef = DocObject.WorksheetRef
    bIsWorkbookSet = DocObject.bIsWorkbookSet
    bIsWorksheetSet = DocObject.bIsWorksheetSet
    bIsSaved = DocObject.bIsSaved
    bIsLinked = DocObject.bIsLinked
    bReadyForRender = DocObject.bReadyForRender
    'This might need to be edited so that the dictionary is copied!
    Set cUDT_Data = DocObject.cUDT_Data
    Set iDocDataObject = DocObject.iDocDataObject
    If Err.Number = 0 Then CloneDoc = True
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_Document.CloneDoc", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Public Function BuildDocumentObject(ByVal DocName As String, ByVal DocID As Long, ByVal DocType As e_DocuType, _
    ByRef cETL As cCBA_UDT_ETL, Optional ByRef iDoc_Object As iCBA_Doc_Data, Optional ByVal lTopOffset As Long = 0, _
        Optional lLeftOffset As Long = 0) As Boolean ' @TP
    Dim UT As cCBA_UDT
    Dim v1 As Variant
    Dim v2 As Variant
    Dim sdMaxGrpNos As Scripting.Dictionary
    Dim lMaxGrpNo As Long
    Dim lmaxTop As Long
    Dim lRH As Long
    Dim sdNewUDT As Scripting.Dictionary
    On Error GoTo Err_Routine
    CBA_Error = ""
    
    Set iDocDataObject = iDoc_Object
    Set cUDT_Data = cETL
    
    
    Set sdMaxGrpNos = New Scripting.Dictionary
backtothestartbaby:
    For Each v1 In cUDT_Data.sdUDT
        Set UT = cUDT_Data.sdUDT(v1)
        If UT.bytUT_IsGrp > 0 Then
            If sdMaxGrpNos.Exists(CStr(UT.sUT_Procedure)) = False Then
                sdMaxGrpNos.Add CStr(UT.sUT_Procedure), CLng(UT.lUT_Grp_No)
                lmaxTop = 0: lRH = UT.lUT_Merge_Height
                For Each v2 In cUDT_Data.sdUDT
                    If cUDT_Data.sdUDT(v2).sUT_Procedure = UT.sUT_Procedure Then
                        If lmaxTop < cUDT_Data.sdUDT(v2).lUT_Pos_Top Then lmaxTop = cUDT_Data.sdUDT(v2).lUT_Pos_Top
                        If cUDT_Data.sdUDT(v2).lUT_Grp_No > sdMaxGrpNos(CStr(UT.sUT_Procedure)) Then sdMaxGrpNos(CStr(UT.sUT_Procedure)) = cUDT_Data.sdUDT(v2).lUT_Grp_No
                    End If
                Next
                lMaxGrpNo = iDocDataObject.GetGrpNo(CStr(UT.sUT_Procedure))
                If lMaxGrpNo > sdMaxGrpNos(CStr(UT.sUT_Procedure)) Then
                    Call cUDT_Data.MassEdit(True, lmaxTop, lMaxGrpNo - sdMaxGrpNos(CStr(UT.sUT_Procedure)))
                   GoTo backtothestartbaby
                End If
            End If
        End If
    Next
    
    sDocumentName = DocName
    eDocumentType = DocType
    lDocumentID = DocID
    
    'If lTopOffset > 0 Then cUDT_Data.MassEdit True, 1, lTopOffset
    'If lLeftOffset > 0 Then cUDT_Data.MassEdit False, 1, lLeftOffset
    
    
    lDocumentID = DocID
    BuildDocumentObject = True
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_Document.BuildDocumentObject", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Sub DefineLinkedStatus()
    If Not cUDT_Data Is Nothing And bIsWorkbookSet = True And bIsWorksheetSet = True Then bIsLinked = True Else bIsLinked = False
End Sub
Private Function TranslateDbValues(ByVal DbFieldName As String, ByVal DvVal As String) As Variant
    Dim vVal As Variant
    On Error GoTo Err_Routine
    CBA_Error = ""
    Select Case DbFieldName
        Case "UT_Wrap_Text", "UT_Locked", "UT_Italic", "UT_Bold", "UT_Underline"
            If DvVal = "Y" Then vVal = True Else vVal = False
        Case "UT_TextAlign"
            If DvVal = "C" Then vVal = xlCenter
            If DvVal = "L" Then vVal = xlLeft
            If DvVal = "R" Then vVal = xlRight
            If DvVal = "T" Then vVal = xlTop
            If DvVal = "B" Then vVal = xlBottom
        Case "UT_Field_Type"
            If DvVal = "txt" Then vVal = xlGeneral
        Case "UT_Border"
            If DvVal = 0 Then vVal = xlNone
            If DvVal = 1 Then vVal = xlThin
            If DvVal = 2 Then vVal = xlMedium
            If DvVal = 3 Then vVal = xlThick
    End Select
    TranslateDbValues = vVal
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_Document.TranslateDbValues", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function CrtAndRenderDoc(Optional ByRef WB As Workbook, Optional ByRef WS As Worksheet) As String
    On Error GoTo Err_Routine
    CBA_Error = ""
    If bReadyForRender = False Then
        'If sdUTs Is Nothing Then Set sdUTs = cUDT_Data.sdUDT '@@@@UT
        If Not cUDT_Data.sdUDT Is Nothing Then
            If WB Is Nothing Then Set WorkbookRef = Application.Workbooks.Add Else Set WorkbookRef = WB
            If WB Is Nothing Then
                Set WorksheetRef = ActiveSheet
            ElseIf WS Is Nothing Then
                Set WorksheetRef = WorkbookRef.Worksheets.Add
            Else
                Set WorksheetRef = WS
            End If
            If eDocumentType = eCoreRangeCategoryReview Then WorksheetRef.Name = "Cat Rev"
            sLinkDoc = WorkbookRef.Name & "-" & WorksheetRef.CodeName
            CrtAndRenderDoc = sLinkDoc
            Call RenderDoc
        Else
            MsgBox "Cannot RenderDoc as the necessary data is not available"
        End If
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_Document.CrtAndRenderDoc", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function RenderDoc(Optional bChangesOnly As Boolean = False) As Boolean
    ' Will Render the UDT_Data obj to the worksheet
    Dim cUDT As cCBA_UDT
    Dim v As Variant, v1 As Variant
    Dim s As String
    Dim rng As Range
    Dim lChkNo As Long
    Dim lBtnNo As Long
    Dim sdVBACode As Scripting.Dictionary, bFmt As Boolean
    Dim sDefVal As String, sVBA As String
    Dim bButtonRendered As Boolean
    On Error GoTo Err_Routine
    CBA_Error = "": RenderDoc = False: bButtonRendered = bIsRendered '''': lRecNo = 0
    
    If bIsRendered = True Then
        If cUDT_Data.Get_SysType(eDocumentType, "Password") > "" Then Workbooks(WorkbookRef.Name).Worksheets(WorksheetRef.Name).Unprotect cUDT_Data.Get_SysType(eDocumentType, "Password")
    End If
    If bChangesOnly = True Then
        Call g_SetupIP("RenderDoc", 1, True)
        For Each v In cUDT_Data.sdUDT
            Set cUDT = cUDT_Data.sdUDT(v)
            sDefVal = cUDT.sUT_Default_Value
            bFmt = iDocDataObject.GetUDTData(cUDT, "1stFormat")
            If bFmt = True And sDefVal <> cUDT.sUT_Default_Value Then
                Set rng = Workbooks(WorkbookRef.Name).Worksheets(WorksheetRef.Name).Cells(cUDT.lUT_Pos_Top, cUDT.lUT_Pos_Left)
                rng.Value = cUDT.sUT_Default_Value
             End If
        Next
        RenderDoc = True
        Call g_SetupIP("RenderDoc", 1, False)
        GoTo ReRender
    End If
    
    If cUDT_Data Is Nothing Then
        MsgBox "No cUDT Data Available to Render"
        RenderDoc = False
    Else
        If bIsRendered = False Then
            Set sdVBACode = iDocDataObject.GetVBACode                   ' @RWCam Will need password included and put in VBACode - hard coded at the moment
        End If
        Call g_SetupIP("RenderDoc", 1, True)
        Application.ScreenUpdating = False
        Range(WorksheetRef.Cells(, 1), WorksheetRef.Cells(, cUDT_Data.NO_OF_COLS)).ColumnWidth = cUDT_Data.COL_WIDTH
        Range(WorksheetRef.Cells(, cAFCOL + 1), WorksheetRef.Cells(, cAFCOL + 1)).ColumnWidth = (cUDT_Data.COL_WIDTH + 0.3) * cUDT_Data.NO_OF_COLS
        Range(WorksheetRef.Cells(, cAFCOL + 1), WorksheetRef.Cells(, cAFCOL + 1)).Locked = True
        For Each v In cUDT_Data.sdUDT
            Set cUDT = cUDT_Data.sdUDT(v)
            With Range(WorksheetRef.Cells(cUDT.lUT_Pos_Top, cUDT.lUT_Pos_Left), WorksheetRef.Cells(cUDT.lUT_Pos_Top + cUDT.lUT_Merge_Height - 1, cUDT.lUT_Pos_Left + cUDT.lUT_Merge_Width - 1))
                bFmt = iDocDataObject.GetUDTData(cUDT, "1stFormat")
                If bFmt = False Then
                    CBA_Error = g_Fmt_2_IDs(cUDT.lUT_Pos_Top, cUDT.lUT_Pos_Left, e_UTFldFmt.eUT_TopLeft, e_UTFldFmt.eUT_TopLeft) & " (UDT_ID=" & cUDT.sUT_ID & ") UDT didn't format"
                    MsgBox CBA_Error
                    Debug.Print CBA_Error                   ' RWCam Fix up when ready
                    Stop
                End If
                .RowHeight = cUDT_Data.ROW_HEIGHT
                .Merge
                If cUDT.sUT_Locked = "Y" Then
                    .Locked = True
                Else
                    .Locked = False
                End If
                If cUDT.sUT_Image > "" Then
''                    .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\" & cUDT.sUT_Image   ' @RWCam Can't get this working for now, even though it works in TEN
                End If
                .WrapText = TranslateDbValues("UT_Wrap_Text", cUDT.sUT_Wrap_Text)
                .Font.Size = cUDT.lUT_Font_Size
                .Interior.ColorIndex = cUDT.lUT_BG_Color
                .Font.ColorIndex = cUDT.lUT_FG_Color
                .Font.Bold = TranslateDbValues("UT_Bold", cUDT.sUT_Bold)
                .Font.Italic = TranslateDbValues("UT_Italic", cUDT.sUT_Italic)
                .Font.Underline = TranslateDbValues("UT_Underline", cUDT.sUT_Underline)
                .Font.Strikethrough = TranslateDbValues("UT_StrikeThrough", cUDT.sUT_StrikeThrough)
                '.Select
                If Mid(cUDT.sUT_TextAlign, 1, 1) <> "N" And Mid(cUDT.sUT_TextAlign, 1, 1) <> "X" Then
                    If InStr(1, cUDT.sUT_TextAlign, "L") > 0 Or InStr(1, cUDT.sUT_TextAlign, "C") > 0 Or InStr(1, cUDT.sUT_TextAlign, "R") > 0 Then
                    .HorizontalAlignment = TranslateDbValues("UT_TextAlign", Mid(cUDT.sUT_TextAlign, 1, 1))
                    End If
                    If InStr(1, cUDT.sUT_TextAlign, "T") > 0 Or InStr(1, cUDT.sUT_TextAlign, "CC") > 0 Or InStr(1, cUDT.sUT_TextAlign, "B") > 0 Then
                        .VerticalAlignment = TranslateDbValues("UT_TextAlign", Mid(cUDT.sUT_TextAlign, 2, 1))
                    End If
                End If
                If cUDT.sUT_Formula <> "" Then
                    .Formula = cUDT.sUT_Formula
                Else
                    .Value = cUDT.sUT_Default_Value
                End If
                .NumberFormat = TranslateDbValues("UT_Field_Type", cUDT.sUT_Field_Type)
                If InStr(1, cUDT.sUT_Border, "T") > 0 And Mid(cUDT.sUT_Border, InStr(1, cUDT.sUT_Border, "T") + 1, 1) > 0 Then _
                    .Borders(xlEdgeTop).Weight = TranslateDbValues("UT_Border", Mid(cUDT.sUT_Border, InStr(1, cUDT.sUT_Border, "T") + 1, 1))
                If InStr(1, cUDT.sUT_Border, "B") > 0 And Mid(cUDT.sUT_Border, InStr(1, cUDT.sUT_Border, "B") + 1, 1) > 0 Then _
                    .Borders(xlEdgeBottom).Weight = TranslateDbValues("UT_Border", Mid(cUDT.sUT_Border, InStr(1, cUDT.sUT_Border, "B") + 1, 1))
                If InStr(1, cUDT.sUT_Border, "L") > 0 And Mid(cUDT.sUT_Border, InStr(1, cUDT.sUT_Border, "L") + 1, 1) > 0 Then _
                    .Borders(xlEdgeLeft).Weight = TranslateDbValues("UT_Border", Mid(cUDT.sUT_Border, InStr(1, cUDT.sUT_Border, "L") + 1, 1))
                If InStr(1, cUDT.sUT_Border, "R") > 0 And Mid(cUDT.sUT_Border, InStr(1, cUDT.sUT_Border, "R") + 1, 1) > 0 Then _
                    .Borders(xlEdgeRight).Weight = TranslateDbValues("UT_Border", Mid(cUDT.sUT_Border, InStr(1, cUDT.sUT_Border, "R") + 1, 1))
'                If InStr(1, cUDT.sUT_Line_Type, "T") > 0 And Mid(cUDT.sUT_Line_Type, InStr(1, cUDT.sUT_Line_Type, "T") + 1, 1) > 0 Then _
'                    .Borders(xlEdgeTop).LineStyle = Mid(cUDT.sUT_Line_Type, InStr(1, cUDT.sUT_Line_Type, "T") + 1, 1)
'                If InStr(1, cUDT.sUT_Line_Type, "B") > 0 And Mid(cUDT.sUT_Line_Type, InStr(1, cUDT.sUT_Line_Type, "B") + 1, 1) > 0 Then _
'                    .Borders(xlEdgeBottom).LineStyle = Mid(cUDT.sUT_Line_Type, InStr(1, cUDT.sUT_Line_Type, "B") + 1, 1)
'                If InStr(1, cUDT.sUT_Line_Type, "L") > 0 And Mid(cUDT.sUT_Line_Type, InStr(1, cUDT.sUT_Line_Type, "L") + 1, 1) > 0 Then _
'                    .Borders(xlEdgeLeft).LineStyle = Mid(cUDT.sUT_Line_Type, InStr(1, cUDT.sUT_Line_Type, "L") + 1, 1)
'                If InStr(1, cUDT.sUT_Line_Type, "R") > 0 And Mid(cUDT.sUT_Line_Type, InStr(1, cUDT.sUT_Line_Type, "R") + 1, 1) > 0 Then _
'                    .Borders(xlEdgeRight).LineStyle = Mid(cUDT.sUT_Line_Type, InStr(1, cUDT.sUT_Line_Type, "R") + 1, 1)

                ' Other parts set-up
                If cUDT.sUT_Formula > "" Then
                    .Formula = cUDT.sUT_Formula
                End If
                If cUDT.sUT_Hyperlink > "" Then
                    .Hyperlink = cUDT.sUT_Hyperlink
                End If
                If cUDT.sUT_Cond_Format > "" Then

                End If


                ' Reset the Default Value if it is affected by the Procedure ...    Note: lSL_Template_No=0 and UT_Upd="A" if no data has been gotton from the table for that segment
                'If cUDT.sUT_Procedure > "" And cUDT.lSL_Template_No = 0 And cUDT.UT_Upd = "A" Then
                    'cUDT.UT_Default_Value = mTEN_Runtime.Set_TD_TD_Value(cUDT.sUT_Procedure, cUDT.sUT_Name, cUDT.sUT_Source, cUDT.sUT_Default_Value, cUDT.lUT_Pos_Top, cUDT.lUT_Pos_Left, cUDT.lUT_Grp_No)
                'End If


                ' Field Type
                If Len(cUDT.sUT_Field_Type) >= 3 Then
                    Select Case Left(cUDT.sUT_Field_Type, 3)
                        Case "taf"
                            .NumberFormat = xlGeneral
''                            cUDT.sUT_Default_Value = ChrW(9643) & cUDT.sUT_Default_Value    '@RW was meant to be bullets...
''                            .EntireRow.AutoFit
                            If cUDT.sUT_Wrap_Text = "Y" Then
                                Set rng = Range(g_GetExcelCell(cUDT.lUT_Pos_Top, cAFCOL))
                                rng.WrapText = False
                                rng.Value = cUDT.sUT_Default_Value
                                rng.WrapText = True
                                rng.EntireRow.AutoFit
                            End If
                            If g_WordCount(cUDT.sUT_Default_Value) > Val(g_Right(cUDT.sUT_Field_Type & " ", 3)) Then .Font.ColorIndex = 3
                        Case "txt"
                            .NumberFormat = xlGeneral
                        Case "txa"
                            .NumberFormat = xlGeneral
                            cUDT.sUT_Default_Value = "'" & cUDT.sUT_Default_Value
                        Case "cur", "per", "dte", "num", "now"
                            If Len(cUDT.sUT_Field_Type) > 3 Then
                                If Left(cUDT.sUT_Field_Type, 3) = "now" Then
                                    cUDT.sUT_Default_Value = Format(Date, CBA_DMY)
''                                        cUDT.Upd_Class_UT "UT_Default_Value", cUDT.sUT_Default_Value, "Upd"
                                End If
                                .NumberFormat = g_Right(cUDT.sUT_Field_Type, 3)
                            End If
                        Case "chk"              ' Create a CheckBox where specified
                            lChkNo = lChkNo + 1
                            Call AddCheckbox(lChkNo, IIf(cUDT.sUT_Default_Value = "Y", True, False), cUDT.lUT_Pos_Top, cUDT.lUT_Pos_Left, g_GetExcelCell(cUDT.lUT_Pos_Top, cUDT.lUT_Pos_Left))
                        Case "ddb"              ' Create a DropDownBox where specified
                            .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=cUDT.sUT_Formula
                            DoEvents
                        Case "btn"              ' Create a button at the end of the line where specified
                            lBtnNo = lBtnNo + 1
                            If bButtonRendered = False Then
                                If Not sdVBACode Is Nothing Then
                                    For Each v1 In sdVBACode
                                        If v1 = "ThisWorkbook" Then
                                            '''WorkbookRef.VBProject.VBComponents(v1).CodeModule.AddFromString sdVBACode(v1)  WorkbookRef.Name & - & WorksheetRef.CodeName
                                        ElseIf v1 = "ThisWorksheet" Then
                                            GoSub VBAReplace
                                            WorkbookRef.VBProject.VBComponents(WorksheetRef.CodeName).CodeModule.AddFromString sVBA
                                        End If
                                    Next
                                End If
                                DoEvents
                                bButtonRendered = True
                            End If
                            Call AddButton(lBtnNo, g_Right(cUDT.sUT_Field_Type, 3), cUDT.lUT_Pos_Top, cUDT.lUT_Pos_Left, g_GetExcelCell(cUDT.lUT_Pos_Top, cUDT.lUT_Pos_Left), g_GetExcelCell(cUDT.lUT_Pos_Top - 1 + cUDT.lUT_Merge_Height, cUDT.lUT_Pos_Left))
                    End Select
                End If
''                    .id = cUDT.sUT_ID                       ' @RWCam Can't get Cell.IDs to work
            End With
        Next
    End If
    If bIsRendered = False Then
        If Not sdVBACode Is Nothing Then
            For Each v1 In sdVBACode
                If v1 = "ThisWorkbook" Then
                    GoSub VBAReplace
                    WorkbookRef.VBProject.VBComponents(v1).CodeModule.AddFromString sVBA
                ElseIf v1 = "ThisWorksheet1" Then
                    GoSub VBAReplace
                    WorkbookRef.VBProject.VBComponents(WorksheetRef.CodeName).CodeModule.AddFromString sVBA
                    bButtonRendered = True         ' @RWCam - will have to make the "ThisWorksheet1" name a bit more universal eventually
                End If
            Next
        End If
    End If
    With WorksheetRef
        .Activate
        RenderDoc = True
        Application.PrintCommunication = False
        'Range(.Cells(1, 1), .Cells(lUdtBot, NO_OF_COLS + 1)).Select
        .PageSetup.FitToPagesTall = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.RightFooter = "Sheet: &A"
        .PageSetup.LeftMargin = cUDT_Data.PAGE_MARGIN
        .PageSetup.RightMargin = cUDT_Data.PAGE_MARGIN
        .PageSetup.TopMargin = cUDT_Data.PAGE_MARGIN
        .PageSetup.BottomMargin = cUDT_Data.PAGE_MARGIN
        .PageSetup.HeaderMargin = 1         '''' PAGE_MARGIN
        .PageSetup.FooterMargin = cUDT_Data.PAGE_MARGIN
        '.PageSetup.PrintArea = "$A$1:" & g_GetExcelCell(lUdtBot, NO_OF_COLS + 2, "$")

        'Application.PrintCommunication = True
        '.Cells(lFreezePos + 1, NO_OF_COLS + 1).Select
        'ActiveWindow.FreezePanes = True
    End With
ReRender:           ' Workbooks(WorkbookRef.Name).Worksheets(WorksheetRef.Name).
    If cUDT_Data.Get_SysType(eDocumentType, "Password") > "" Then Workbooks(WorkbookRef.Name).Worksheets(WorksheetRef.Name).Protect cUDT_Data.Get_SysType(eDocumentType, "Password")
    bIsRendered = True
    Application.ScreenUpdating = True

Exit_Routine:

    On Error Resume Next
    Call g_SetupIP("RenderDoc", 1, False)
    Exit Function
VBAReplace:
    sVBA = sdVBACode(v1)
    sVBA = Replace(sVBA, "WorkbookRef.Name", WorkbookRef.Name)
    sVBA = Replace(sVBA, "WorksheetRef.CodeName", WorksheetRef.CodeName)
    sVBA = Replace(sVBA, "AFVAL", Left(g_GetExcelCell(20, cAFCOL), 2))
    Return

Err_Routine:
''    If Err.Number = 1004 Then Resume Next
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_Document.RenderDoc", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function AddUDTLines(Optional ByVal sVal1 As String, Optional ByVal lVal1 As Long, Optional ByVal sVal2 As String, Optional ByVal lVal2 As Long) As Boolean
    ' Will Add Whole UDT Lines (multiple UDTs with each line) to the UDT_Data obj
    Dim cUDT As cCBA_UDT, cUDT2 As cCBA_UDT, cUDT3 As cCBA_UDT
    Dim v As Variant, dic As Scripting.Dictionary, lRec As Long, bIn1stPass As Boolean, bGotAllPass As Boolean, lCapLinF As Long, lCapLinT As Long
    Dim lNewTop As Long, lLastTop As Long, lLastHgt As Long, lCapNo As Long, lIterNo As Long, bHasBeenAdded As Boolean, lIdxNo As Long
''    Dim s As String
''    Dim rng As Range
''    Dim lChkNo As Long
''    Dim lBtnNo As Long
''    Dim sdVBACode As Scripting.Dictionary, bFmt As Boolean
''    Dim sDefVal As String
''    Dim lRecNo As Long
    On Error GoTo Err_Routine
    CBA_Error = "": AddUDTLines = False

    If cUDT_Data Is Nothing Then
        MsgBox "No cUDT Data Available to AddUDTLines"
        AddUDTLines = False
    Else
        lRec = 0
DoRest:
        Do While lRec < cUDT_Data.sdUDT.Count - 1
            Set cUDT = cUDT_Data.sdUDT.Keys(lRec)
            If cUDT.sUT_Procedure = sVal1 Then
                If bIn1stPass = False Then
                    lCapLinF = lRec: bIn1stPass = True: lLastTop = cUDT.lUT_Pos_Top: lLastHgt = cUDT.lUT_Merge_Height
                Else
                    If lCapLinT = 0 And lRec = cUDT_Data.sdUDT.Count - 1 Then
                        lCapLinT = lRec: bGotAllPass = True
                    End If
                End If
            Else
                If lCapLinT = 0 Then
                    lCapLinT = lRec - 1: bGotAllPass = True
                End If
                lIdxNo = cUDT.lUT_No
            End If
            If bHasBeenAdded = True Then
                lIdxNo = lIdxNo + 1
                cUDT.lUT_No = lIdxNo
                cUDT.lUT_Pos_Top = cUDT.lUT_Pos_Top + lNewTop
                dic.Add cUDT_Data.sdUDT.Keys(lRec), cUDT
            ElseIf bGotAllPass = False Then
                dic.Add cUDT_Data.sdUDT.Keys(lRec), cUDT
            ElseIf bGotAllPass = True And lIterNo = 0 Then
                For lIterNo = 1 To lVal1
                    For lCapNo = lCapLinF To lCapLinT
                        Set cUDT2 = New cCBA_UDT
                        If cUDT.CloneETL(cUDT_Data(cUDT_Data.sdUDT.Keys(lCapNo))) = False Then
                            MsgBox lCapNo & " failed to index/clone"
                            Stop
                            GoTo Exit_Routine
                        End If
''                        Set cUDT2 = cUDT_Data.CloneETL(cUDT_Data.sdUDT.Keys(lCapNo))
''                        Set cUDT3 = cUDT.Clone(cUDT_Data(cUDT_Data.sdUDT.Keys(lCapNo)))
                        If lCapNo = lCapLinF Then lNewTop = lNewTop + lLastHgt
                        lIdxNo = lIdxNo + 1
                        cUDT.lUT_No = lIdxNo
                        cUDT.lUT_Grp_No = lIterNo + 1
                        cUDT.lUT_Pos_Top = cUDT.lUT_Pos_Top + lNewTop
                        cUDT.sUT_Name = g_Left(cUDT.sUT_Name, 1) & (lIterNo + 1)
                        dic.Add cUDT_Data.sdUDT.Keys(lCapNo), cUDT2
                    Next
                Next
                bHasBeenAdded = True
                AddUDTLines = True
            End If
        Loop
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    If Err.Number = 1004 Then Resume Next
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_Document.AddUDTLines", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
''Private Function AddUDTLines(Optional ByVal sValue1 As String, Optional ByVal lValue1 As Long) As Boolean
''    ' Will Add Whole UDT Lines (multiple UDTs with each line) to the UDT_Data obj
''    Dim cUDT As cCBA_UDT, cUDT2 As cCBA_UDT, cUDT3 As cCBA_UDT
''    Dim v As Variant, dic As Scripting.Dictionary, lRec As Long, bIn1stPass As Boolean, bGotAllPass As Boolean, lCapLinF As Long, lCapLinT As Long
''    Dim lNewTop As Long, lLastTop As Long, lLastHgt As Long, lCapNo As Long, lIterNo As Long, bHasBeenAdded As Boolean, lIdxNo As Long
''''    Dim s As String
''''    Dim rng As Range
''''    Dim lChkNo As Long
''''    Dim lBtnNo As Long
''''    Dim sdVBACode As Scripting.Dictionary, bFmt As Boolean
''''    Dim sDefVal As String
''''    Dim lRecNo As Long
''    On Error GoTo Err_Routine
''    CBA_Error = "": AddUDTLines = False
''
''    If cUDT_Data Is Nothing Then
''        MsgBox "No cUDT Data Available to AddUDTLines"
''        AddUDTLines = False
''    Else
''        lRec = 0
''DoRest:
''        Do While lRec < cUDT_Data.sdUDT.Count - 1
''            Set cUDT = cUDT_Data.sdUDT.Keys(lRec)
''            If cUDT.sUT_Procedure = sValue1 Then
''                If bIn1stPass = False Then
''                    lCapLinF = lRec: bIn1stPass = True: lLastTop = cUDT.lUT_Pos_Top: lLastHgt = cUDT.lUT_Merge_Height
''                Else
''                    If lCapLinT = 0 And lRec = cUDT_Data.sdUDT.Count - 1 Then
''                        lCapLinT = lRec: bGotAllPass = True
''                    End If
''                End If
''            Else
''                If lCapLinT = 0 Then
''                    lCapLinT = lRec - 1: bGotAllPass = True
''                End If
''                lIdxNo = cUDT.lUT_No
''            End If
''            If bHasBeenAdded = True Then
''                lIdxNo = lIdxNo + 1
''                cUDT.lUT_No = lIdxNo
''                cUDT.lUT_Pos_Top = cUDT.lUT_Pos_Top + lNewTop
''                dic.Add cUDT_Data.sdUDT.Keys(lRec), cUDT
''            ElseIf bGotAllPass = False Then
''                dic.Add cUDT_Data.sdUDT.Keys(lRec), cUDT
''            ElseIf bGotAllPass = True And lIterNo = 0 Then
''                For lIterNo = 1 To lValue1
''                    For lCapNo = lCapLinF To lCapLinT
''                        Set cUDT2 = New cCBA_UDT
''                        If cUDT.CloneETL(cUDT_Data(cUDT_Data.sdUDT.Keys(lCapNo))) = False Then
''                            MsgBox lCapNo & " failed to index/clone"
''                            Stop
''                            GoTo Exit_Routine
''                        End If
''''                        Set cUDT2 = cUDT_Data.CloneETL(cUDT_Data.sdUDT.Keys(lCapNo))
''''                        Set cUDT3 = cUDT.Clone(cUDT_Data(cUDT_Data.sdUDT.Keys(lCapNo)))
''                        If lCapNo = lCapLinF Then lNewTop = lNewTop + lLastHgt
''                        lIdxNo = lIdxNo + 1
''                        cUDT.lUT_No = lIdxNo
''                        cUDT.lUT_Grp_No = lIterNo + 1
''                        cUDT.lUT_Pos_Top = cUDT.lUT_Pos_Top + lNewTop
''                        cUDT.sUT_Name = g_Left(cUDT.sUT_Name, 1) & (lIterNo + 1)
''                        dic.Add cUDT_Data.sdUDT.Keys(lCapNo), cUDT2
''                    Next
''                Next
''                bHasBeenAdded = True
''                AddUDTLines = True
''            End If
''        Loop
''    End If
''Exit_Routine:
''
''    On Error Resume Next
''    Exit Function
''
''Err_Routine:
''    If Err.Number = 1004 Then Resume Next
''    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-cCBA_Document.AddUDTLines", 3)
''    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
''    Debug.Print CBA_Error
''    Stop            ' ^RW Camera + next line
''    Resume Next
''    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
''    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
''    GoTo Exit_Routine
''    Resume Next
''End Function
''
Sub AddButton(lNo As Long, sName As String, lRow As Long, lcol As Long, sSPos As String, sEPos As String)

    Range(sSPos, sEPos).Select
    With WorksheetRef.Buttons.Add(Selection.Left, Selection.Top, Selection.Width, Selection.Height)
        .Name = "Button" & lNo
        .Caption = sName
        ''.OnAction = ThisWorkbook.Name & "!'CBA_TEN_WBK.ActionButton  " & lRow & ", " & lCol & ", """ & sAddit & """'"   ''' Original
        ''.OnAction = WorksheetRef.Name & "!'" & WorksheetRef.CodeName & ".ActionButton  " & lRow & ", " & lCol & "'"
''        .OnAction = "!'" & WorksheetRef.CodeName & ".ActionButton  " & lRow & ", " & lCol & "'"
''        .OnAction = "!'" & WorksheetRef.Name & ".ActionButton  " & lRow & ", " & lCol & "'"
''        .OnAction = WorksheetRef.Name & ".ActionButton  " & lRow & ", " & lCol '''& "'"
        .OnAction = WorkbookRef.Name & "!'" & WorksheetRef.CodeName & ".ActionButton  " & lRow & ", " & lcol & "'"
    End With

End Sub
Sub AddCheckbox(lNo As Long, bOnOff As Boolean, lRow As Long, lcol As Long, sPos As String)
    Range(sPos).Select
    With WorksheetRef.CheckBoxes.Add(Selection.Left, Selection.Top, Selection.Width, Selection.Height)
        .Name = "CheckBox" & lNo
        .Caption = ""
        .LinkedCell = sPos
        If Not bOnOff Then
            .Value = xlOff
        Else
            .Value = xlOn
        End If
        '.OnAction = WorksheetRef.Name & "!'CBA_TEN_WBK.ActionCheckBox CheckBox" & lNo & ", " & lRow & ", " & lCol & "'" '    & ",'" & sPos & "''"
    End With
End Sub
Public Function AddChart(ByVal LeftCellPosition As Long, ByVal TopCellPosition As Long, ByVal sChartTemplatePath As String, ByVal DocID As String, _
    ByRef RangeIncLabels As Range, Optional ByRef RangeDataOnly As Range, Optional ByVal sXAxisLabelFormat As String, Optional ByVal bXAxisReversePlotOrder As Boolean, _
        Optional ByVal sYAxisLabelFormat As String, Optional ByVal bYAxisReversePlotOrder As Boolean, Optional ByVal ChartWidth As Double, Optional ChartHeight As Double) As Chart '@TP
    '*************************************************************************************************
    '******************GENERAL CHART INFO TO HELP BUILD THE CHART YOU WANT'***************************
    '*************************************************************************************************
    '***************************CHART TYPES***********************************************************
    'xl3DArea = 3-D Area                            xl3DAreaStacked = 3-D Stacked Area              xl3DAreaStacked100 = 3-D Stacked Area
    'xl3DBarClustered = 3-D Clustered Bar           xl3DBarStacked = 3-D Stacked Bar                xl3DBarStacked100 = 3-D 100% Stacked Bar
    'xl3DColumn = 3-D Column                        xl3DColumnClustered = 3-D Clustered Column      xl3DColumnStacked = 3-D Stacked Column
    'xl3DColumnStacked100 = 3-D 100% Stacked Column  xl3DLine = 3-D Line                            xl3DPie = 3-D Pie
    'xl3DPieExploded = Exploded 3-D Pie             xlArea = Area                                   xlAreaStacked  = Stacked Area
    'xlAreaStacked100 = 100% Stacked Area           xlBarClustered = Clustered Bar                  xlBarOfPie = Bar of Pie
    'xlBarStacked = Stacked Bar                     xlBarStacked100 = 100% Stacked Bar              xlBubble = Bubble
    'xlBubble3DEffect = Bubble with 3-D Effects     xlColumnClustered = Clustered Column            xlColumnStacked = Stacked Column
    'xlColumnStacked100 = 100% Stacked Column       xlConeBarClustered = Clustered Cone Bar         xlConeBarStacked = Stacked Cone Bar
    'xlConeBarStacked100 = 100% Stacked Cone Bar    xlConeCol = 3-D Cone Column                     xlConeColClustered = Clustered Cone Column
    'xlConeColStacked = Stacked Cone Column         xlConeColStacked100 = 100% Stacked Cone Column  xlCylinderBarStacked = Stacked Cylinder Bar
    'xlDoughnut = Doughnut                          xlDoughnutExploded = Exploded Doughnut          xlLineMarkers  = Line with Data Markers
    'xlLine = Line                                  xlLineStacked = Stacked Line                    xlPie = Pie
    'xlPieExploded = Exploded Pie                   xlPieOfPie = Pie of Pie                         xlPyramidBarClustered = Clustered Pyramid Bar
    'xlRadar = Radar                                xlRadarFilled = Filled Radar                    xlRadarMarkers = Radar with Data Markers
    'xlStockHLC = High-Low-Close                    xlStockOHLC = Open-High-Low-Close               xlStockVHLC = Volume-High-Low-Close
    'xlStockVOHLC = Volume-Open-High-Low-Close      xlSurface = 3-D Surface                         xlSurfaceTopView = Surface (Top View)
    'xlSurfaceWireframe = 3-D Surface(wire-frame)   xlXYScatter = Scatter                           xlXYScatterLines = Scatter with Lines
    'xlPyramidCol = 3-D Pyramid Column              xlPyramidColStacked = Stacked Pyramid Column    xlPyramidColClustered = Clustered Pyramid Column
    'xlCylinderCol = 3-D Cylinder Column                        xlCylinderColStacked = Stacked Cylinder Column
    'xlCylinderBarClustered = Clustered Cylinder Bar            xlCylinderBarStacked100 = 100% Stacked Cylinder Bar
    'xlCylinderColClustered = Clustered Cylinder Column         xlCylinderColStacked100 = 100% Stacked Cylinder Column
    'xlLineMarkersStacked100 = 100% Stacked Line w Markers      xlLineStacked100 = 100% Stacked Line
    'xlSurfaceTopViewWireframe = Surface (Top View wire-frame)  xlLineMarkersStacked = Stacked Line w Data Markers
    'xlPyramidBarStacked = Stacked Pyramid Bar                  xlPyramidBarStacked100 = 100% Stacked Pyramid Bar
    'xlPyramidColStacked100 = 100% Stacked Pyramid Column       xlXYScatterSmoothNoMarkers = Scatter with Smoothed Lines and No Data Markers
    'xlXYScatterSmooth = Scatter with SmoothedLines             xlXYScatterLinesNoMarkers = Scatter with Lines and No Data Markers
    '************************************************************************************************************************
Dim cht As Chart

    Set cht = WorksheetRef.Shapes.AddChart2.Chart
    With cht
        .ApplyChartTemplate sChartTemplatePath
        .SetSourceData Source:=RangeIncLabels
        If sXAxisLabelFormat <> "" Then .Axes(xlCategory).TickLabels.NumberFormat = sXAxisLabelFormat
        If sYAxisLabelFormat <> "" Then .Axes(xlValue).TickLabels.NumberFormat = sYAxisLabelFormat
        If bXAxisReversePlotOrder = True Then .Axes(xlCategory).ReversePlotOrder = True
        If bYAxisReversePlotOrder = True Then .Axes(xlValue).ReversePlotOrder = True
        If Not RangeDataOnly Is Nothing Then
            .Axes(xlValue).MinimumScale = Application.WorksheetFunction.Min(RangeDataOnly)
            .Axes(xlValue).MaximumScale = Application.WorksheetFunction.Max(RangeDataOnly)
        End If
        If ChartWidth <> 0 Then .ChartArea.Width = ChartWidth
        If ChartHeight <> 0 Then .ChartArea.Height = ChartHeight
        .ChartArea.Left = WorksheetRef.Cells(1, LeftCellPosition).Left
        .ChartArea.Top = WorksheetRef.Cells(TopCellPosition, 1).Top
    End With
    If Not cht Is Nothing Then Set AddChart = cht
End Function
Public Function ControlDocument(ByRef Target As Range, ByVal Command As String, Optional ByRef DocID As Long, Optional sVal1 As String, Optional lVal1 As Long, Optional lVal2 As Long) As Long
    Dim cUDT As cCBA_UDT, sUT_ID As String
    'Dim tempDocID As Long
    'Command = "rendercell"
    ControlDocument = True
    Select Case LCase(Command)
        Case "save"
            If cUDT_Data.Write_UDT_ETL(DocID, eDocumentType) = False Then
                MsgBox "Failure in Saving Document"
                ControlDocument = False
                Exit Function
            End If
            lDocumentID = DocID
            bIsSaved = True
        Case "protect"
            If cUDT_Data.Get_SysType(eDocumentType, "Password") > "" Then Workbooks(WorkbookRef.Name).Worksheets(WorksheetRef.Name).Protect cUDT_Data.Get_SysType(eDocumentType, "Password")
        Case "unprotect"
            If cUDT_Data.Get_SysType(eDocumentType, "Password") > "" Then Workbooks(WorkbookRef.Name).Worksheets(WorksheetRef.Name).Unprotect cUDT_Data.Get_SysType(eDocumentType, "Password")
        Case "addrow"                       ' Add a row to an unrendered UDT
            If AddUDTLines(sVal1, lVal1, lVal2) = False Then
                MsgBox "Failure in Adding Rows to Document"
                ControlDocument = False
                Exit Function
            End If
        Case "addcolumn"
        
        Case "addoleobject"
        
        Case "deletecolumn"
        
        Case "deleterow"
        
        Case "copyrow"
    
        Case "copycolumn"
    
        Case "copycell"
    
        Case "close"
        
        Case "deletesheet"
        
        Case "copysheet"
        
        Case "renderdoc"
            If RenderDoc = False Then
                MsgBox "Failure in Rendering Document"
                ControlDocument = False
                Exit Function
            End If
        Case "rendercell"
            RenderDoc True
        
        Case "refreshfc"
            If RenderDoc(True) = True Then
                MsgBox "Refreshed"
            Else
                MsgBox "Failure in Refreshing Document"
                ControlDocument = False
                Exit Function
            End If
        Case "ut_merge_height"
            GoSub GSRowCol                      ' Get the UDT concerned
            ControlDocument = cUDT.lUT_Merge_Height
        Case "updaterc"
            GoSub GSRowCol                      ' Get the UDT concerned
            cUDT.sUT_Default_Value = sVal1
            If Left(cUDT.sUT_Field_Type, 3) = "taf" Then
                ControlDocument = IIf(g_WordCount(sVal1) > Val(g_Right(cUDT.sUT_Field_Type & " ", 3)), 1, 0)
            End If
        Case "autofit"
            GoSub GSRowCol                      ' Get the UDT concerned
            ControlDocument = IIf(Left(cUDT.sUT_Field_Type, 3) = "taf", 1, 0)
        Case ""
        
        Case ""
        
        Case ""
        
        Case ""
        
        Case ""
    
    
    End Select
''    ControlDocument = True
    Exit Function
GSRowCol:
    sUT_ID = g_Fmt_2_IDs(lVal1, lVal2, e_UTFldFmt.eUT_TopLeft, e_UTFldFmt.eUT_TopLeft)
    If cUDT_Data.sdUDT.Exists(sUT_ID) = False Then
        MsgBox sUT_ID & " doesn't exist in " & Command
        ControlDocument = False
        Exit Function
    Else
        Set cUDT = cUDT_Data.sdUDT(sUT_ID)
    End If
    Return
End Function
''Public Function ControlDocument(ByRef Target As Range, ByVal Command As String, Optional ByRef DocID As Long, _
''    Optional ByVal sValue1 As String, Optional ByVal lValue2 As Long) As Boolean
''    'Dim tempDocID As Long
''    'Command = "rendercell"
''    Select Case LCase(Command)
''        Case "save"
''            If cUDT_Data.Write_UDT_ETL(DocID, eDocumentType) = False Then
''                MsgBox "Failure in Saving Document"
''                ControlDocument = False
''                Exit Function
''            End If
''            lDocumentID = DocID
''            bIsSaved = True
''
''        Case "protect"
''
''        Case "unprotect"
''
''        Case "addrow"                       ' Add a row to an unrendered UDT
''            If AddUDTLines(sValue1, lValue2) = False Then
''                MsgBox "Failure in Adding Rows to Document"
''                ControlDocument = False
''                Exit Function
''            End If
''        Case "addcolumn"
''
''        Case "addoleobject"
''
''        Case "deletecolumn"
''
''        Case "deleterow"
''
''        Case "copyrow"
''
''        Case "copycolumn"
''
''        Case "copycell"
''
''        Case "close"
''
''        Case "deletesheet"
''
''        Case "copysheet"
''
''        Case "renderdoc"
''            If RenderDoc = False Then
''                MsgBox "Failure in Rendering Document"
''                ControlDocument = False
''                Exit Function
''            End If
''        Case "rendercell"
''            RenderDoc True
''
''        Case "refreshfc"
''            If RenderDoc(True) = True Then
''                MsgBox "Refreshed"
''            Else
''                MsgBox "Failure in Refreshing Document"
''                ControlDocument = False
''                Exit Function
''            End If
''        Case ""
''
''        Case ""
''
''        Case ""
''
''        Case ""
''
''        Case ""
''
''        Case ""
''
''        Case ""
''
''        Case ""
''
''
''    End Select
''    ControlDocument = True
''End Function
'Private Function IsIndexedUT() As Boolean
'''    If cUDT_Data.sdUDT Is Nothing Or GenerateDocumentDics = False Then
'    If cUDT_Data.sdUDT Is Nothing Then
'        IsIndexedUT = False
'    Else
'        bReadyForRender = True
'        IsIndexedUT = True
'    End If
'End Function
Private Property Set WorkbookRef(ByVal objNewValue As Workbook)
    Set pwbWorkbookRef = objNewValue
    If objNewValue Is Nothing Then bIsWorkbookSet = False Else bIsWorkbookSet = True
    DefineLinkedStatus
End Property

Private Property Set WorksheetRef(ByVal objNewValue As Worksheet)
    Set pwsWorksheetRef = objNewValue
    If objNewValue Is Nothing Then bIsWorksheetSet = False Else bIsWorksheetSet = True
    DefineLinkedStatus
End Property

Public Property Get sDocumentName() As String: sDocumentName = psDocumentName: End Property
Private Property Let sDocumentName(ByVal sNewValue As String): psDocumentName = sNewValue:  End Property

Public Property Get eDocumentType() As e_DocuType: eDocumentType = peDocumentType: End Property
Private Property Let eDocumentType(ByVal eNewValue As e_DocuType): peDocumentType = eNewValue: DefineLinkedStatus: End Property

Public Property Get WorkbookRef() As Workbook: Set WorkbookRef = pwbWorkbookRef: End Property
Public Property Get WorksheetRef() As Worksheet: Set WorksheetRef = pwsWorksheetRef: End Property

Public Property Get bIsLinked() As Boolean: bIsLinked = pbIsLinked: End Property
Private Property Let bIsLinked(ByVal bNewValue As Boolean): pbIsLinked = bNewValue: End Property

Public Property Get bIsWorkbookSet() As Boolean: bIsWorkbookSet = pbIsWorkbookSet: End Property
Private Property Let bIsWorkbookSet(ByVal bNewValue As Boolean): pbIsWorkbookSet = bNewValue: End Property
Public Property Get bIsWorksheetSet() As Boolean: bIsWorksheetSet = pbIsWorksheetSet: End Property
Private Property Let bIsWorksheetSet(ByVal bNewValue As Boolean): pbIsWorksheetSet = bNewValue: End Property

Public Property Get bReadyForRender() As Boolean
    'you shouldnt need this line:: if indexSegDic = True then
    bReadyForRender = pbReadyForRender
End Property
Private Property Let bReadyForRender(ByVal bNewValue As Boolean): pbReadyForRender = bNewValue: End Property

Private Sub Class_Initialize()
'bIsWorkbookSet = False
'bIsWorksheetSet = False
End Sub

Public Property Get cUDT_Data() As cCBA_UDT_ETL: Set cUDT_Data = pcUDT_Data: End Property
Private Property Set cUDT_Data(ByVal objNewValue As cCBA_UDT_ETL): Set pcUDT_Data = objNewValue: End Property

Public Property Get bIsSaved() As Boolean: bIsSaved = pbIsSaved: End Property
Public Property Let bIsSaved(ByVal bNewValue As Boolean): pbIsSaved = bNewValue: End Property

Public Property Get iDocDataObject() As iCBA_Doc_Data: Set iDocDataObject = piDocDataObject: End Property
Private Property Set iDocDataObject(ByVal objNewValue As iCBA_Doc_Data): Set piDocDataObject = objNewValue: End Property

Public Property Get lDocumentID() As Long: lDocumentID = plDocumentID: End Property
Public Property Let lDocumentID(ByVal lNewValue As Long)
Dim UF As Object
    plDocumentID = lNewValue
    If Not iDocDataObject Is Nothing Then Call iDocDataObject.SetDocID(lNewValue)
End Property

Public Property Get sLinkDoc() As String: sLinkDoc = psLinkDoc: End Property
Private Property Let sLinkDoc(ByVal sNewValue As String): psLinkDoc = sNewValue: End Property

Private Property Get bIsRendered() As Boolean: bIsRendered = pbIsRendered: End Property
Private Property Let bIsRendered(ByVal bNewValue As Boolean): pbIsRendered = bNewValue: End Property
