VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_AST_clsProdRows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit   ' ASYST Product Rows Class Module 190402
' 190220 RW@ Now make it to do AssocFld and ProdRow forms

''' Use a class module to create a collection of textboxes for the AST ProductRows and AssocFlds form
' For Product Rows
Private WithEvents tProductCode As MSForms.TextBox
Attribute tProductCode.VB_VarHelpID = -1
Private WithEvents tProductDesc As MSForms.TextBox
Attribute tProductDesc.VB_VarHelpID = -1
Private WithEvents tFutureProdCode As MSForms.TextBox
Attribute tFutureProdCode.VB_VarHelpID = -1
Private WithEvents tOnSaleDate As MSForms.TextBox
Attribute tOnSaleDate.VB_VarHelpID = -1
Private WithEvents tWeeksOfSale As MSForms.TextBox
Attribute tWeeksOfSale.VB_VarHelpID = -1
Private WithEvents tEndCatDate As MSForms.TextBox
Attribute tEndCatDate.VB_VarHelpID = -1

Private WithEvents tTableMerch As MSForms.CheckBox
Attribute tTableMerch.VB_VarHelpID = -1
Private WithEvents tEndCapMerch As MSForms.CheckBox
Attribute tEndCapMerch.VB_VarHelpID = -1

Private WithEvents tUCostExpMSupp As MSForms.TextBox
Attribute tUCostExpMSupp.VB_VarHelpID = -1
Private WithEvents tSuppCostActMSupp As MSForms.TextBox
Attribute tSuppCostActMSupp.VB_VarHelpID = -1
Private WithEvents tCurRetFreq As MSForms.TextBox
Attribute tCurRetFreq.VB_VarHelpID = -1
Private WithEvents tNewRetOOH As MSForms.TextBox
Attribute tNewRetOOH.VB_VarHelpID = -1
Private WithEvents tDiscStandee As MSForms.TextBox
Attribute tDiscStandee.VB_VarHelpID = -1
Private WithEvents tSalesMultMedRegs As MSForms.TextBox
Attribute tSalesMultMedRegs.VB_VarHelpID = -1
Private WithEvents tUPSPWPress As MSForms.TextBox
Attribute tUPSPWPress.VB_VarHelpID = -1
Private WithEvents tPriorSalesTV As MSForms.TextBox
Attribute tPriorSalesTV.VB_VarHelpID = -1
Private WithEvents tCalcRadio As MSForms.TextBox
Attribute tCalcRadio.VB_VarHelpID = -1

Private WithEvents tProductApprovalDate As MSForms.TextBox
Attribute tProductApprovalDate.VB_VarHelpID = -1
Private WithEvents tGBDMApprovalDate As MSForms.TextBox
Attribute tGBDMApprovalDate.VB_VarHelpID = -1
' For AssocFlds
Private WithEvents tUnitCost As MSForms.TextBox
Attribute tUnitCost.VB_VarHelpID = -1
Private WithEvents tSupplierCostSupport As MSForms.TextBox
Attribute tSupplierCostSupport.VB_VarHelpID = -1
Private WithEvents tPriorSales As MSForms.TextBox
Attribute tPriorSales.VB_VarHelpID = -1
Private WithEvents tExpectedSales As MSForms.TextBox
Attribute tExpectedSales.VB_VarHelpID = -1
Private WithEvents tCalculatedSales As MSForms.TextBox
Attribute tCalculatedSales.VB_VarHelpID = -1
Private WithEvents tUPSPW As MSForms.TextBox
Attribute tUPSPW.VB_VarHelpID = -1
Private WithEvents tFillQty As MSForms.TextBox
Attribute tFillQty.VB_VarHelpID = -1
Private WithEvents tCurrRetailPrice As MSForms.TextBox
Attribute tCurrRetailPrice.VB_VarHelpID = -1
Private WithEvents tRetailPrice As MSForms.TextBox
Attribute tRetailPrice.VB_VarHelpID = -1
Private WithEvents tSalesMultiplier As MSForms.TextBox
Attribute tSalesMultiplier.VB_VarHelpID = -1

' For both
Private WithEvents tUpd As MSForms.TextBox
Attribute tUpd.VB_VarHelpID = -1
Private WithEvents tID As MSForms.TextBox
Attribute tID.VB_VarHelpID = -1
Private tMeTop As Long, tMeLeft As Long, plAuth As Long, psForm As String, psMrk_Buy As String, psField As String, plTimes As Long

Public Sub FormInit(AllCtls As MSForms.Controls, lNo As Long, MeTop As Long, MeLeft As Long, sForm As String, Optional BorM As String = "M")
    ' Will add a control array and handle event thereof
    Dim ctl As MSForms.Control, sMsgF As String, sMsgN As String, sSepF As String, sSepN As String, bfound As Boolean, sPath As String, bIsNum As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    tMeTop = MeTop: tMeLeft = MeLeft: psForm = sForm
    sMsgF = "": sSepF = "": sMsgN = "": sSepN = "": psMrk_Buy = "B"
    ' Set authorities
    If psForm = "AssocFlds" Then
        plAuth = CBA_lAuthority
        psMrk_Buy = "B"
    Else
        If BorM = "M" Then
            psMrk_Buy = "M"
        Else
            psMrk_Buy = "B"
        End If
    End If
    ' Set up each control on the form (Note: the whole form will be checked but only the line no concerned will be taken)
    For Each ctl In AllCtls
        bfound = False: bIsNum = False
        If Left(ctl.Name, 3) = "txt" Or Left(ctl.Name, 3) = "chk" Then
            If ctl.Name = "txt" & Format(lNo, "00") & "ProductCode" Then
                Set tProductCode = ctl
                bfound = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "ProductDesc" Then
                Set tProductDesc = ctl
                bfound = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "FutureProdCode" Then
                Set tFutureProdCode = ctl
                bfound = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "OnSaleDate" Then
                Set tOnSaleDate = ctl
                bfound = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "WeeksOfSale" Then
                Set tWeeksOfSale = ctl
                bfound = True
            ElseIf ctl.Name = "chk" & Format(lNo, "00") & "TableMerch" Then
                Set tTableMerch = ctl
                bfound = True
            ElseIf ctl.Name = "chk" & Format(lNo, "00") & "EndCapMerch" Then
                Set tEndCapMerch = ctl
                bfound = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "EndCatDate" Then
                Set tEndCatDate = ctl
                bfound = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "UCostExpMSupp" Then
                Set tUCostExpMSupp = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "SuppCostActMSupp" Then
                Set tSuppCostActMSupp = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "PriorSalesTV" Then
                Set tPriorSalesTV = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "DiscStandee" Then
                Set tDiscStandee = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "CalcRadio" Then
                Set tCalcRadio = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "UPSPWPress" Then
                Set tUPSPWPress = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "CurRetFreq" Then
                Set tCurRetFreq = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "NewRetOOH" Then
                Set tNewRetOOH = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "SalesMultMedRegs" Then
                Set tSalesMultMedRegs = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "ProductApprovalDate" Then
                Set tProductApprovalDate = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "GBDMApprovalDate" Then
                Set tGBDMApprovalDate = ctl
                bfound = True
                plAuth = CBA_lAuthority
                If plAuth <> 1 And g_IsDate(tGBDMApprovalDate, True) = True Then plAuth = 0
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "UnitCost" Then
                Set tUnitCost = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "SupplierCostSupport" Then
                Set tSupplierCostSupport = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "PriorSales" Then
                Set tPriorSales = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "ExpectedSales" Then
                Set tExpectedSales = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "CalculatedSales" Then
                Set tCalculatedSales = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "UPSPW" Then
                Set tUPSPW = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "FillQty" Then
                Set tFillQty = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "CurrRetailPrice" Then
                Set tCurrRetailPrice = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "RetailPrice" Then
                Set tRetailPrice = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "SalesMultiplier" Then
                Set tSalesMultiplier = ctl
                bfound = True: bIsNum = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "Upd" Then
                Set tUpd = ctl
                bfound = True
            ElseIf ctl.Name = "txt" & Format(lNo, "00") & "ID" Then
                Set tID = ctl
                bfound = True
            End If
        End If
        If bfound = True Then
            sMsgF = sMsgF & sSepF & ctl.Name
            sSepF = ";"
        Else
            sMsgN = sMsgN & sSepN & ctl.Name
            sSepN = ";"
        End If
    Next
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("CBA_AST_clsProdRows-FormInit", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub

Private Sub p_RegionClick(sOld As String)
    Dim sFields As String, sFFields As String, sSQL As String, bPassOK As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    If psMrk_Buy = "M" Then Exit Sub
    CBA_lProduct_ID = tID
    ''Call AST_CrtProdRegionRecs(CBA_lProduct_ID)   No longer needed as should be created
    CBA_lAuth = plAuth
    CBA_strAldiMsg = "RegionRows~" & sOld
    CBA_AST_frm_AssocFld.Show vbModal
    ' Global flag says regions have changed
    If CBA_bDataChg = True Then
        Call AST_AveProdRegionRecs(CBA_lProduct_ID, CBA_strAldiMsg)
        If CBA_strAldiMsg > "" Then
            CBA_DBtoQuery = 1
            sFields = "PD_Unit_Cost,PD_Supplier_Cost_Support,PD_Curr_Retail_Price,PD_Retail_Price,PD_Retail_Discount,PD_Sales_Multiplier,PD_UPSPW  ,PD_Prior_Sales,PD_Calculated_Sales"
            sFFields = "UCostExpMSupp,SuppCostActMSupp      ,CurRetFreq          ,NewRetOOH      ,DiscStandee,       SalesMultMedRegs   ,UPSPWPress,PriorSalesTV  ,CalcRadio          "
            sSQL = "SELECT " & sFields & " FROM [L2_Products] WHERE PD_ID=" & CBA_lProduct_ID & ";"
            bPassOK = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "L2_Products", g_GetDB("ASYST"), CBA_MSAccess, sSQL, 120, , , False)
            If bPassOK = True Then
                ''Call g_SetupIP(psForm, 1, True)                     ' Don't pull the update as the table is already updated
                tUCostExpMSupp.Value = NZ(CBA_ABIarr(0, 0), 0)
                Call p_UpdIP(tUCostExpMSupp, , True, False)
                tSuppCostActMSupp.Value = NZ(CBA_ABIarr(1, 0), 0)
                Call p_UpdIP(tSuppCostActMSupp, , True, False)
                tCurRetFreq.Value = NZ(CBA_ABIarr(2, 0), 0)
                Call p_UpdIP(tCurRetFreq, , True, False)
                tNewRetOOH.Value = NZ(CBA_ABIarr(3, 0), 0)
                Call p_UpdIP(tNewRetOOH, , True, False)
                tDiscStandee.Value = NZ(CBA_ABIarr(4, 0), 0)
                Call p_UpdIP(tDiscStandee, , True, False)
                tSalesMultMedRegs.Value = NZ(CBA_ABIarr(5, 0), 0)
                Call p_UpdIP(tSalesMultMedRegs, , True, False)
                tUPSPWPress.Value = NZ(CBA_ABIarr(6, 0), 0)
                Call p_UpdIP(tUPSPWPress, , True, False)
                tPriorSalesTV.Value = NZ(CBA_ABIarr(7, 0), 0)
                Call p_UpdIP(tPriorSalesTV, , True, False)
                tCalcRadio.Value = NZ(CBA_ABIarr(8, 0), 0)
                Call p_UpdIP(tCalcRadio, , True, False)
                ''Call g_SetupIP(psForm, 1, False)
            End If
        End If
        ''Call p_UpdIP(Me.txtRegion)
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("CBA_AST_clsProdRows-p_RegionClick", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    ''If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Sub p_UpdIP(cActCtl As Control, Optional bIsDate As Boolean = False, Optional bFormat As Boolean = False, Optional bLockTable As Boolean = True)
    ' Flag the record as being updated
    Static bRecursed As Boolean
    Dim sFormat As String, sFlag As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    If bRecursed = True Then Exit Sub
    
    Call CBA_ProcI("s-p_UpdIP", 3)
    bRecursed = True
    ' If a format, get it...
    If bFormat Then sFormat = AST_FillTagArrays(cActCtl.Tag, CBA_lFrmID, plAuth, "Format")
    ' If a date, check to see if it has changed, if not then hop out of update
    If bIsDate = True Then
        If varCal.bCalValReturned = False Then GoTo Exit_Routine
    End If
    ' Process the update of the field/line
    If g_SetupIP(psForm) = False Then
        If psForm = "ProdRows" And Val(Mid(cActCtl.Name, 4, 2)) > 0 Then
    ''        Debug.Print Mid(cActCtl.Name, 4, 2) & ".";
            If tUpd.Value = "N" Or tUpd.Value = "x" Then
                ' Update the flag to show it has been updated - This should cause the field in the form module to fire....
                If bLockTable = True Then
                    sFlag = AST_Lock_Table("ASYST", "PD_", "L2_Products", "PD_ID=" & tID, plAuth, "Get_SetY")
                    If sFlag = "AlreadyLocked" Then
                        sFlag = "E"
                        tProductCode.BackColor = CBA_Red
                    Else
                        sFlag = "Y"
                        tProductCode.BackColor = CBA_Green
                    End If
                    ' Update the maint values into the Upd field
                    tUpd.Value = sFlag
                    CBA_PRa(0, AST_TF("Upd"), Val(tUpd.Tag)) = sFlag
                End If
            End If
            ' Update the array with the new value
            If Val(tUpd.Tag) > 0 And g_Right(cActCtl.Name, 5) <> "Upd" Then CBA_PRa(0, AST_TF(cActCtl.Tag), Val(tUpd.Tag)) = cActCtl.Value
        ElseIf psForm = "AssocFlds" Then
            If Val(Mid(cActCtl.Name, 4, 2)) > 0 Then
                varFldVars.sField3 = cActCtl.Tag
                If tUpd.Value = "Y" Then
                    tUpd.Value = "y"
                Else
                    tUpd.Value = "Y"
                End If
                sFlag = "Y"
            End If
        End If
    End If
    ' Apply the format
    If bFormat Then cActCtl.Value = Format(cActCtl.Value, sFormat)
Exit_Routine:
    bRecursed = False
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("CBA_AST_clsProdRows-p_UpdIP", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Private Sub tProductDesc_Change()
    Static bRecurse As Boolean
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    Call p_UpdIP(tProductDesc)
''    tProductDesc.Value = tProductDesc.Value & "[[["
''  ' Debug.Print tProductDesc.Name & ":" & tProductDesc.Value & ";";
    bRecurse = False
End Sub

Private Sub tFutureProdCode_Change()
    Static bRecurse As Boolean
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    Call p_UpdIP(tFutureProdCode)
''    tFutureProdCode.Value = tFutureProdCode.Value & "[[["
''  ' Debug.Print tFutureProdCode.Name & ":" & tFutureProdCode.Value & ";";
    bRecurse = False
End Sub

Private Sub tTableMerch_Change()           '_Click()
    Static bRecurse As Boolean
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    Call p_UpdIP(tTableMerch)
''    tTableMerch.Value = tTableMerch.Value & "[[["
''  ' Debug.Print tTableMerch.Name & ":" & tTableMerch.Value & ";";
    bRecurse = False
End Sub

Private Sub tEndCapMerch_Change()           'Click()
    Static bRecurse As Boolean
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    Call p_UpdIP(tEndCapMerch)
''    tEndCapMerch.Value = tEndCapMerch.Value & "[[["
''  ' Debug.Print tEndCapMerch.Name & ":" & tEndCapMerch.Value & ";";
    bRecurse = False
End Sub

Private Sub tOnSaleDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    If bRecurse = True Or psMrk_Buy = "M" Then Exit Sub
    If Mid(tOnSaleDate.Name, 4, 2) = "00" Then Exit Sub
    Call CBA_BT_CalendarShow(tMeTop, tMeLeft, tOnSaleDate, CBA_D2DMY, False)
    Call AST_setEndDate1(psForm, tOnSaleDate, tWeeksOfSale, tEndCatDate, CBA_D2DMY)
    'Call p_UpdIP(tOnSaleDate, True)
    bRecurse = False
End Sub

Private Sub tOnSaleDate_Change()
    Static bRecurse As Boolean
    If bRecurse = True Then Exit Sub
    Call p_UpdIP(tOnSaleDate, True)
    bRecurse = False
End Sub

Private Sub tUpd_Change()
    Static bRecurse As Boolean
    If bRecurse = True Then Exit Sub
    If tUpd.Value <> "x" Then Exit Sub
    bRecurse = True
    If g_SetupIP(psForm) = False Then
        Call p_UpdIP(tUpd)
        ''Debug.Print Mid(tUpd.Name, 4, 2) & "clstupd;";
    End If
    bRecurse = False
End Sub

Private Sub tWeeksOfSale_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    If bRecurse = True Or psMrk_Buy = "M" Then Exit Sub
    If Mid(tWeeksOfSale.Name, 4, 2) = "00" Then Exit Sub
    bRecurse = True
    varFldVars.lFldWidth = 80
    varFldVars.lFldHeight = 0
    varFldVars.lFrmLeft = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
    ''varFldVars.lFrmwidth = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
    varFldVars.lFrmTop = g_PosForm(0, 0, 0, "Top")
    varFldVars.sHdg = "Enter Weeks Of Sales"
    varFldVars.sSQL = "AST_WeeksOfSale"
    varFldVars.sDB = "ASYST"
    varFldVars.bAllowNullOfField = False
    varFldVars.lCols = 1
    varFldVars.sType = "ComboBox"
    CBA_frmEntryField.Show vbModal
    If CBA_bDataChg Then
        tWeeksOfSale.Value = varFldVars.sField1
        Call AST_setEndDate1(psForm, tOnSaleDate, tWeeksOfSale, tEndCatDate, CBA_D2DMY)
        ''Call p_UpdIP(tWeeksOfSale)
    End If
    bRecurse = False
End Sub
Private Sub tWeeksOfSale_Change()
    Static bRecurse As Boolean
    If bRecurse = True Then Exit Sub
    Call p_UpdIP(tWeeksOfSale)
    bRecurse = False
End Sub

Private Sub tProductApprovalDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If psMrk_Buy = "M" Then Exit Sub
    If Mid(tProductApprovalDate.Name, 4, 2) = "00" Then Exit Sub
''    If Val(Mid(tProductApprovalDate.Name, 4, 2)) = 0 Then Exit Sub
    ' Test to see if the date can be applied - some fields may not be filled in the l2_products table....
''    If g_IsDate(tProductApprovalDate.Value, True) = False Then
''        If AST_UpdGBDMDate(g_FixDate(tProductApprovalDate.Value), CBA_lPromotion_ID, True, CStr(tID.Value), "Fatal", "Product Approval Date") <> "True" Then Exit Sub
''    End If
    Call CBA_BT_CalendarShow(tMeTop, tMeLeft, tProductApprovalDate, CBA_D2DMY)
    If varCal.bCalValReturned = False Then Exit Sub
''    If bHasDates = True Then AST_UpdGBDMDate(g_FixDate(tProductApprovalDate.Value), CBA_lPromotion_ID, tID.Value, , "Product Approval Date") <> "True" Then Exit Sub
''    Call p_UpdIP(tProductApprovalDate, True)
End Sub
Private Sub tProductApprovalDate_Change()
    Static bRecurse As Boolean
    If bRecurse = True Then Exit Sub
    If Mid(tProductApprovalDate.Name, 4, 2) = "00" Then Exit Sub
    Call p_UpdIP(tProductApprovalDate)
    bRecurse = False
End Sub

Private Sub tEndCatDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not psMrk_Buy = "M" Then Exit Sub
    If Mid(tEndCatDate.Name, 4, 2) = "00" Then Exit Sub
    Call CBA_BT_CalendarShow(tMeTop, tMeLeft, tEndCatDate, CBA_D2DMY)
    Call p_UpdIP(tEndCatDate, True)
End Sub
Private Sub tEndCatDate_Change()
    Static bRecurse As Boolean
    If bRecurse = True Then Exit Sub
    Call p_UpdIP(tEndCatDate)
    bRecurse = False
End Sub

Private Sub tGBDMApprovalDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim bHasDate As Boolean
    If psMrk_Buy = "M" Then Exit Sub
    If Mid(tGBDMApprovalDate.Name, 4, 2) = "00" Then Exit Sub
        ' Test to see if the date can be applied - some fields may not be filled in the l2_products table, some will be filled as they are encountered ...
    If g_IsDate(tGBDMApprovalDate.Value, True) = False Then
        If AST_UpdGBDMDate(Date, CBA_lPromotion_ID, True, CStr(tID.Value), "Fatal") <> "True" Then Exit Sub
    Else
        bHasDate = True
    End If
    ' Get the date to apply
    Call CBA_BT_CalendarShow(tMeTop, tMeLeft, tGBDMApprovalDate, CBA_D2DMY)
    If varCal.bCalValReturned = False Then Exit Sub
    If g_IsDate(tProductApprovalDate.Value, True) = False And g_IsDate(tGBDMApprovalDate.Value, True) = True Then
        tProductApprovalDate.Value = tGBDMApprovalDate.Value
        Call p_UpdIP(tProductApprovalDate, True)
    End If
    If bHasDate = False Then Call AST_UpdGBDMDate(g_FixDate(tGBDMApprovalDate.Value), CBA_lPromotion_ID, False, CStr(tID.Value), "All", , False)
    Call p_UpdIP(tGBDMApprovalDate, True)
End Sub

Private Sub tGBDMApprovalDate_Change()
    Static bRecurse As Boolean
    If bRecurse = True Then Exit Sub
    If Mid(tGBDMApprovalDate.Name, 4, 2) = "00" Then Exit Sub
    Call p_UpdIP(tGBDMApprovalDate)
    bRecurse = False
End Sub

Private Sub tNewRetOOH_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean, sFormat As String, sTagName As String
    Dim sOld As String
    If Mid(tNewRetOOH.Name, 4, 2) = "00" Then Exit Sub
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        '''If tNewRetOOH.BackColor = CBA_Grey Then Exit Sub
        sOld = tNewRetOOH
        CBA_strAldiMsg = "Out_Of_Home_Dates~" & sOld
        varCal.sDate = g_FixDate(tOnSaleDate)
        CBA_AST_frm_AssocFld.Show vbModal
        If CBA_strAldiMsg > "" Then tNewRetOOH.Value = CBA_strAldiMsg
        If sOld <> tNewRetOOH Then Call p_UpdIP(tNewRetOOH)
    Else
        Call p_RegionClick(tNewRetOOH)
    End If
    bRecurse = False
End Sub

Private Sub tSuppCostActMSupp_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean, sFormat As String, sTagName As String
    Dim sOld As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    If Mid(tSuppCostActMSupp.Name, 4, 2) = "00" Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
''        '''If tSuppCostActMSupp.BackColor = CBA_Grey Then Exit Sub
        varFldVars.lFldWidth = 80
        varFldVars.lFldHeight = 0
        varFldVars.lFrmLeft = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
        ''varFldVars.lFrmwidth = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
        varFldVars.lFrmTop = g_PosForm(0, 0, 0, "Top")
        varFldVars.sHdg = "Enter Actual Marketing Support"
        varFldVars.sSQL = "SELECT * FROM [L0_Tiers] ORDER BY TR_ID"
        varFldVars.sDB = "ASYST"
        varFldVars.bAllowNullOfField = False
        varFldVars.lCols = 2
        varFldVars.sType = "ComboBox"
        CBA_frmEntryField.Show vbModal
        If CBA_bDataChg Then
            tSuppCostActMSupp.Value = varFldVars.sField2
            Call p_UpdIP(tSuppCostActMSupp)
        End If
    Else
        Call p_RegionClick(tSuppCostActMSupp)
    End If
    
    bRecurse = False

End Sub

Private Sub tUCostExpMSupp_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean, sFormat As String, sTagName As String
    Dim sOld As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    If Mid(tUCostExpMSupp.Name, 4, 2) = "00" Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        ' This field is not changeable...
        Exit Sub
    Else
        Call p_RegionClick(tUCostExpMSupp)
    End If
    
    bRecurse = False

End Sub

Private Sub tPriorSalesTV_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean, sFormat As String, sTagName As String
    Dim sOld As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    If Mid(tPriorSalesTV.Name, 4, 2) = "00" Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        '''If tPriorSalesTV.BackColor = CBA_Grey Then Exit Sub
        sOld = tPriorSalesTV
        CBA_strAldiMsg = "TV_Dates~" & sOld
        varCal.sDate = g_FixDate(tOnSaleDate)
        CBA_AST_frm_AssocFld.Show vbModal
        If CBA_strAldiMsg > "" Then tPriorSalesTV.Value = CBA_strAldiMsg
        If sOld <> tPriorSalesTV Then Call p_UpdIP(tPriorSalesTV)
    Else
        Call p_RegionClick(tPriorSalesTV)
    End If
    
    bRecurse = False
End Sub

Private Sub tDiscStandee_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean, sFormat As String, sTagName As String
    Dim sOld As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    If Mid(tDiscStandee.Name, 4, 2) = "00" Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        '''If tDiscStandee.BackColor = CBA_Grey Then Exit Sub
        sOld = tDiscStandee
        CBA_strAldiMsg = "Standee_Dates~" & sOld
        varCal.sDate = g_FixDate(tOnSaleDate)
        CBA_AST_frm_AssocFld.Show vbModal
        If CBA_strAldiMsg > "" Then tDiscStandee.Value = CBA_strAldiMsg
        If sOld <> tDiscStandee Then Call p_UpdIP(tDiscStandee)
    Else
        Call p_RegionClick(tDiscStandee)
    End If
    
    bRecurse = False

End Sub

Private Sub tCalcRadio_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean, sFormat As String, sTagName As String
    Dim sOld As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    If Mid(tCalcRadio.Name, 4, 2) = "00" Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        '''If tCalcRadio.BackColor = CBA_Grey Then Exit Sub
        sOld = tCalcRadio
        CBA_strAldiMsg = "Radio_Dates~" & sOld
        varCal.sDate = g_FixDate(tOnSaleDate)
        CBA_AST_frm_AssocFld.Show vbModal
        If CBA_strAldiMsg > "" Then tCalcRadio.Value = CBA_strAldiMsg
        If sOld <> tCalcRadio Then Call p_UpdIP(tCalcRadio)
    Else
        Call p_RegionClick(tCalcRadio)
    End If
    
    bRecurse = False
End Sub

Private Sub tUPSPWPress_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean, sFormat As String, sTagName As String
    Dim sOld As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    If Mid(tUPSPWPress.Name, 4, 2) = "00" Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        '''If tUPSPWPress.BackColor = CBA_Grey Then Exit Sub
        sOld = tUPSPWPress
        CBA_strAldiMsg = "Press_Dates_(WA_only)~" & sOld
        varCal.sDate = g_FixDate(tOnSaleDate)
        CBA_AST_frm_AssocFld.Show vbModal
        If CBA_strAldiMsg > "" Then tUPSPWPress.Value = CBA_strAldiMsg
        If sOld <> tUPSPWPress Then Call p_UpdIP(tUPSPWPress)
    Else
        Call p_RegionClick(tUPSPWPress)
    End If
    
    bRecurse = False

End Sub

Private Sub tCurRetFreq_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean, sFormat As String, sTagName As String
    Dim sOld As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    If Mid(tCurRetFreq.Name, 4, 2) = "00" Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
'        '''If tCurRetFreq.BackColor = CBA_Grey Then Exit Sub
        varFldVars.lFldWidth = 80
        varFldVars.lFldHeight = 0
        varFldVars.lFrmLeft = g_PosForm(0, (varFldVars.lFldWidth * 2), 0, "Left")
        varFldVars.lFrmTop = g_PosForm(0, 0, 0, "Top")
        varFldVars.sHdg = "Enter Frequency / Items Per Basket"
        varFldVars.sSQL = "SELECT * FROM [L0_Freq_Items] ORDER BY FI_ID"
        varFldVars.sDB = "ASYST"
        varFldVars.bAllowNullOfField = False
        varFldVars.lCols = 2
        varFldVars.sType = "ComboBox"
        CBA_frmEntryField.Show vbModal
        If CBA_bDataChg Then
            tCurRetFreq.Value = varFldVars.sField2
            Call p_UpdIP(tCurRetFreq)
        End If


'        sOld = tCurRetFreq
'        CBA_strAldiMsg = "Out_Of_Home_Dates~" & sOld
'        varCal.sDate = g_FixDate(tOnSaleDate)
'        CBA_AST_frm_AssocFld.Show vbModal
'        If CBA_strAldiMsg > "" Then tCurRetFreq.Value = CBA_strAldiMsg
'        If sOld <> tCurRetFreq Then Call p_UpdIP(tCurRetFreq)
    Else
        Call p_RegionClick(tCurRetFreq)
    End If
    
    bRecurse = False
End Sub

Private Sub tSalesMultMedRegs_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean, sFormat As String, sTagName As String
    Dim sOld As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    If Mid(tSalesMultMedRegs.Name, 4, 2) = "00" Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        '''If tSalesMultMedRegs.BackColor = CBA_Grey Then Exit Sub
        sOld = tSalesMultMedRegs
        CBA_strAldiMsg = "Media_Regions~" & sOld
        varCal.sDate = g_FixDate(tOnSaleDate)
        CBA_AST_frm_AssocFld.Show vbModal
        If CBA_strAldiMsg > "" Then tSalesMultMedRegs.Value = CBA_strAldiMsg
        If sOld <> tSalesMultMedRegs Then Call p_UpdIP(tSalesMultMedRegs)
    Else
        Call p_RegionClick(tSalesMultMedRegs)
    End If
    bRecurse = False
End Sub

Private Sub tRetailPrice_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Static bRecurse As Boolean, sFormat As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    ' Reset the end of field vars
    If psField <> tRetailPrice.Name Then
        psField = tRetailPrice.Name
        plTimes = 0
    End If
    ' Test the entry
    If g_GetKeyType(KeyCode, Shift) = "Exit" Then   ' Tab or Enter...
        If plTimes > 0 Then
            If sFormat = "" Then sFormat = AST_FillTagArrays(tRetailPrice.Tag, CBA_lFrmID, plAuth, "Format")
            tRetailPrice.Value = Format(tRetailPrice.Value, sFormat)
            Call p_UpdIP(tRetailPrice)
        End If
    ElseIf g_GetKeyType(KeyCode, Shift) <> "Num" Then        ' Not a number
        ''Debug.Print "kc" & KeyCode;
        KeyCode = 0
    Else
        plTimes = plTimes + 1
    End If
    ''Debug.Print tRetailPrice.Name & "=" & sFormat & ";";
    bRecurse = False

End Sub

Private Sub tSupplierCostSupport_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Static bRecurse As Boolean, sFormat As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    ' Reset the end of field vars
    If psField <> tSupplierCostSupport.Name Then
        psField = tSupplierCostSupport.Name
        plTimes = 0
    End If

   'Debug.Print ".";
    ' Test the entry
    If g_GetKeyType(KeyCode, Shift) = "Exit" Then   ' Tab or Enter...
        If plTimes > 0 Then
            If sFormat = "" Then sFormat = AST_FillTagArrays(tSupplierCostSupport.Tag, CBA_lFrmID, plAuth, "Format")
            tSupplierCostSupport.Value = Format(tSupplierCostSupport.Value, sFormat)
            Call p_UpdIP(tSupplierCostSupport)
        End If
    ElseIf g_GetKeyType(KeyCode, Shift) <> "Num" Then        ' Not a number
        ''Debug.Print "kc" & KeyCode;
        KeyCode = 0
    Else
        plTimes = plTimes + 1
    End If
    ''Debug.Print tSupplierCostSupport.Name & "=" & sFormat & ";";
    bRecurse = False

End Sub

Private Sub tUnitCost_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Static bRecurse As Boolean, sFormat As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    ' Reset the end of field vars
    If psField <> tUnitCost.Name Then
        psField = tUnitCost.Name
        plTimes = 0
    End If
    ' Test the entry
    If g_GetKeyType(KeyCode, Shift) = "Exit" Then   ' Tab or Enter...
        If plTimes > 0 Then
            If sFormat = "" Then sFormat = AST_FillTagArrays(tUnitCost.Tag, CBA_lFrmID, plAuth, "Format")
            tUnitCost.Value = Format(tUnitCost.Value, sFormat)
            Call p_UpdIP(tUnitCost)
        End If
    ElseIf g_GetKeyType(KeyCode, Shift) <> "Num" Then        ' Not a number
        ''Debug.Print "kc" & KeyCode;
        KeyCode = 0
    Else
        plTimes = plTimes + 1
    End If
    ''Debug.Print tUnitCost.Name & "=" & sFormat & ";";
    bRecurse = False

End Sub

Private Sub tPriorSales_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Static bRecurse As Boolean, sFormat As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    ' Reset the end of field vars
    If psField <> tPriorSales.Name Then
        psField = tPriorSales.Name
        plTimes = 0
    End If
    ' Test the entry
    If g_GetKeyType(KeyCode, Shift) = "Exit" Then   ' Tab or Enter...
        If plTimes > 0 Then
            If sFormat = "" Then sFormat = AST_FillTagArrays(tPriorSales.Tag, CBA_lFrmID, plAuth, "Format")
            tPriorSales.Value = Format(tPriorSales.Value, sFormat)
            Call p_UpdIP(tPriorSales)
        End If
    ElseIf g_GetKeyType(KeyCode, Shift) <> "Num" Then        ' Not a number
        ''Debug.Print "kc" & KeyCode;
        KeyCode = 0
    Else
        plTimes = plTimes + 1
    End If
    ''Debug.Print tPriorSales.Name & "=" & sFormat & ";";
    bRecurse = False

End Sub

Private Sub tExpectedSales_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Static bRecurse As Boolean, sFormat As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    ' Reset the end of field vars
    If psField <> tExpectedSales.Name Then
        psField = tExpectedSales.Name
        plTimes = 0
    End If
    ' Test the entry
    If g_GetKeyType(KeyCode, Shift) = "Exit" Then   ' Tab or Enter...
        If plTimes > 0 Then
            If sFormat = "" Then sFormat = AST_FillTagArrays(tExpectedSales.Tag, CBA_lFrmID, plAuth, "Format")
            tExpectedSales.Value = Format(tExpectedSales.Value, sFormat)
            Call p_UpdIP(tExpectedSales)
        End If
    ElseIf g_GetKeyType(KeyCode, Shift) <> "Num" Then        ' Not a number
        ''Debug.Print "kc" & KeyCode;
        KeyCode = 0
    Else
        plTimes = plTimes + 1
    End If
    ''Debug.Print tExpectedSales.Name & "=" & sFormat & ";";
    bRecurse = False

End Sub

Private Sub tCalculatedSales_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Static bRecurse As Boolean, sFormat As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    ' Reset the end of field vars
    If psField <> tCalculatedSales.Name Then
        psField = tCalculatedSales.Name
        plTimes = 0
    End If
    ' Test the entry
    If g_GetKeyType(KeyCode, Shift) = "Exit" Then   ' Tab or Enter...
        If plTimes > 0 Then
            If sFormat = "" Then sFormat = AST_FillTagArrays(tCalculatedSales.Tag, CBA_lFrmID, plAuth, "Format")
            tCalculatedSales.Value = Format(tCalculatedSales.Value, sFormat)
            Call p_UpdIP(tCalculatedSales)
        End If
    ElseIf g_GetKeyType(KeyCode, Shift) <> "Num" Then        ' Not a number
        ''Debug.Print "kc" & KeyCode;
        KeyCode = 0
    Else
        plTimes = plTimes + 1
    End If
    ''Debug.Print tCalculatedSales.Name & "=" & sFormat & ";";
    bRecurse = False

End Sub

Private Sub tUPSPW_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Static bRecurse As Boolean, sFormat As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    ' Reset the end of field vars
    If psField <> tUPSPW.Name Then
        psField = tUPSPW.Name
        plTimes = 0
    End If
    ' Test the entry
    If g_GetKeyType(KeyCode, Shift) = "Exit" Then   ' Tab or Enter...
        If plTimes > 0 Then
            If sFormat = "" Then sFormat = AST_FillTagArrays(tUPSPW.Tag, CBA_lFrmID, plAuth, "Format")
            tUPSPW.Value = Format(tUPSPW.Value, sFormat)
            Call p_UpdIP(tUPSPW)
        End If
    ElseIf g_GetKeyType(KeyCode, Shift) <> "Num" Then        ' Not a number
        ''Debug.Print "kc" & KeyCode;
        KeyCode = 0
    Else
        plTimes = plTimes + 1
    End If
    ''Debug.Print tUPSPW.Name & "=" & sFormat & ";";
    bRecurse = False
End Sub

Private Sub tFillQty_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Static bRecurse As Boolean, sFormat As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    ' Reset the end of field vars
    If psField <> tFillQty.Name Then
        psField = tFillQty.Name
        plTimes = 0
    End If
    ' Test the entry
    If g_GetKeyType(KeyCode, Shift) = "Exit" Then   ' Tab or Enter...
        If plTimes > 0 Then
            If sFormat = "" Then sFormat = AST_FillTagArrays(tFillQty.Tag, CBA_lFrmID, plAuth, "Format")
            tFillQty.Value = Format(tFillQty.Value, sFormat)
            Call p_UpdIP(tFillQty)
        End If
    ElseIf g_GetKeyType(KeyCode, Shift) <> "Num" Then        ' Not a number
        ''Debug.Print "kc" & KeyCode;
        KeyCode = 0
    Else
        plTimes = plTimes + 1
    End If
    ''Debug.Print tFillQty.Name & "=" & sFormat & ";";
    bRecurse = False
End Sub

Private Sub tCurrRetailPrice_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Static bRecurse As Boolean, sFormat As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    ' Reset the end of field vars
    If psField <> tCurrRetailPrice.Name Then
        psField = tCurrRetailPrice.Name
        plTimes = 0
    End If
    ' Test the entry
    If g_GetKeyType(KeyCode, Shift) = "Exit" Then   ' Tab or Enter...
        If plTimes > 0 Then
            If sFormat = "" Then sFormat = AST_FillTagArrays(tCurrRetailPrice.Tag, CBA_lFrmID, plAuth, "Format")
            tCurrRetailPrice.Value = Format(tCurrRetailPrice.Value, sFormat)
            Call p_UpdIP(tCurrRetailPrice)
        End If
    ElseIf g_GetKeyType(KeyCode, Shift) <> "Num" Then        ' Not a number
        ''Debug.Print "kc" & KeyCode;
        KeyCode = 0
    Else
        plTimes = plTimes + 1
    End If
    ''Debug.Print tCurrRetailPrice.Name & "=" & sFormat & ";";
    bRecurse = False
End Sub

Private Sub tSalesMultiplier_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Static bRecurse As Boolean, sFormat As String
    ' If setup or already being processed, hop out
    If g_SetupIP(psForm) Or bRecurse Then Exit Sub
    bRecurse = True
    ' Reset the end of field vars
    If psField <> tRetailPrice.Name Then
        psField = tRetailPrice.Name
        plTimes = 0
    End If
    ' Test the entry
    If g_GetKeyType(KeyCode, Shift) = "Exit" Then   ' Tab or Enter...
        If plTimes > 0 Then
            If sFormat = "" Then sFormat = AST_FillTagArrays(tSalesMultiplier.Tag, CBA_lFrmID, plAuth, "Format")
            tSalesMultiplier.Value = Format(tSalesMultiplier.Value, sFormat)
            Call p_UpdIP(tSalesMultiplier)
        End If
    ElseIf g_GetKeyType(KeyCode, Shift) <> "Num" Then        ' Not a number
        ''Debug.Print "kc" & KeyCode;
        KeyCode = 0
    Else
        plTimes = plTimes + 1
    End If
    ''Debug.Print tSalesMultiplier.Name & "=" & sFormat & ";";
    bRecurse = False

End Sub

