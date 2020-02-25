VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_BTF_SCGNonProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit       ' @CBA_BTF_SCGNonProd @CBA_BTF Changed 181211

Private CBA_BTF_CGD(1 To 36) As CBA_BTF_CGDataDetail
Private CGno As Long, SCGno As Long, PClassNo As Long
Private SCGDesc As String
Private BDName As String
Private GBDName As String
Private ContainsData As Boolean

Sub CreateSCG(ByVal CG As Long, ByVal scg As Long, ByVal lPClass As Long, dtDate As Date)
    Dim strSQL As String, bFCNet As Boolean, lM As Long, lMn As Long, lY As Long, lYIdx As Long, lv As Long, lYY As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    Dim RS As ADODB.Recordset
    CGno = CG
    SCGno = scg
    PClassNo = lPClass
    ' Get the POS and Ret data from the current year and the last year  ' @RWFC 200107 - Date changed to input date (dtDate)
    lYY = Year(dtDate) - 1
    Set RS = New ADODB.Recordset
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
    strSQL = strSQL & "DECLARE @ENTEREDDATE date = '" & Year(dtDate) & "-12-31'" & Chr(10)  ' @RWFC 200107 - Date changed to input date (dtDate)
    strSQL = strSQL & "DECLARE @ENTEREDDTE date = dateadd(m,-25,@ENTEREDDATE)" & Chr(10)
    strSQL = strSQL & "DECLARE @SYEAR int = Year(@ENTEREDDTE)" & Chr(10)
    strSQL = strSQL & "DECLARE @CG int = " & CGno & Chr(10)
    strSQL = strSQL & "DECLARE @SCG int = " & SCGno & Chr(10)
    strSQL = strSQL & "DECLARE @PCLASS int = " & PClassNo & Chr(10)
    strSQL = strSQL & "DECLARE @SDATE date = @ENTEREDDATE" & Chr(10)
    strSQL = strSQL & "DECLARE @EDATE date = dateadd(D,-1,dateadd(M,1,@SDATE))" & Chr(10)
    strSQL = strSQL & "select month, year, sum(retail) as retail,sum(retailnet) as retailnet, sum(cost) as cost,sum(isnull(Quantity,0)) as Quantity into #RCV from (" & Chr(10)
    strSQL = strSQL & "select month(dayenddate) as month, year(dayenddate) as year ,sum(rcv.retail) as Retail, sum(rcv.retailnet) as Retailnet" & Chr(10)
    strSQL = strSQL & ", sum(rcv.cost) as Cost, case when rcv.RecordID = '001' then sum(quantity) end as Quantity" & Chr(10)
    strSQL = strSQL & "from (select dayenddate, isnull(p.con_productcode,p.productcode) as Pcode, retail, RetailNet, cost, quantity, RecordID" & Chr(10)
    strSQL = strSQL & "from cbis599p.dbo.receiving rcv " & Chr(10)
    strSQL = strSQL & "left join cbis599p.dbo.product p on p.productcode = rcv.productcode where rcv.dayenddate >= dateadd(M,-37,@SDATE) and rcv.dayenddate <= dateadd(D,-1,@SDATE) and p.CGno = @CG and p.SCGno = @SCG" & Chr(10)
    strSQL = strSQL & ") rcv left join cbis599p.dbo.product p on p.productcode = rcv.pcode where rcv.dayenddate >= dateadd(M,-37,@SDATE) and rcv.dayenddate <= dateadd(D,-1,@SDATE) and p.CGno = @CG and p.SCGno = @SCG and p.ProductClass = @PCLASS" & Chr(10)
    strSQL = strSQL & "group by month(dayenddate) , year(dayenddate), rcv.RecordID) a group by month, year order by year desc, month desc" & Chr(10)
    strSQL = strSQL & "select pos.ProductCode as Pcode,  posdate, pos.retail, pos.quantity into #P from cbis599p.dbo.pos pos" & Chr(10)
    strSQL = strSQL & "left join cbis599p.dbo.product p on p.productcode = pos.productcode where pos.posdate >= dateadd(D,1,dateadd(M,-37,@EDATE)) and pos.posdate <= @EDATE and p.CGno = @CG and p.SCGno = @SCG and p.ProductClass = @PCLASS" & Chr(10)
    strSQL = strSQL & "select year(posdate) as year, month(posdate) as month, sum(pos.retail) as Retail, sum(pos.quantity) as Quantity" & Chr(10)
    strSQL = strSQL & "into #pos from #P pos group by  month(posdate) ,year(posdate) order by  month(posdate) ,year(posdate)" & Chr(10)
    strSQL = strSQL & "select pos.year, pos.month, pos.Quantity as POSQTY, pos.retail as POSRET, pospy.Quantity as POSPYQTY, pospy.retail as POSPYRET" & Chr(10)
    strSQL = strSQL & ", rcv.retail as RCVRET, rcv.retailnet as RCVRETNET, rcv.cost as RCVCOST, rcv.Quantity as RCVQTY" & Chr(10)
    strSQL = strSQL & ", rcvpy.retail as RCVPYRET, rcvpy.cost as RCVPYCOST, rcvpy.Quantity as RCVPYQTY, rcvpy.retailnet as RCVPYRETNET" & Chr(10)
    strSQL = strSQL & "into #BASE from #pos pos full outer join #RCV rcv on rcv.year = pos.year and rcv.month = pos.month" & Chr(10)
    strSQL = strSQL & "left join #pos pospy on pospy.month = pos.month and pospy.year = pos.year -1" & Chr(10)
    strSQL = strSQL & "left join #rcv rcvpy on rcvpy.month = pos.month and rcvpy.year = pos.year -1" & Chr(10)
    strSQL = strSQL & "where pos.year > @SYEAR" & Chr(10)
    strSQL = strSQL & "order by pos.year desc, pos.month desc" & Chr(10)
    strSQL = strSQL & "select year, month, POSQTY, pospyqty, case when isnull(POSPYQTY,0) = 0  then 0 else (isnull(POSQTY,0) - POSPYQTY) / POSPYQTY end as POSQTYYOY" & Chr(10)
    strSQL = strSQL & ", POSRET, case when isnull(POSPYRET,0) = 0  then 0 else (isnull(POSRET,0) - POSPYRET) / POSPYRET end as POSRETYOY" & Chr(10)
    strSQL = strSQL & ", case when RCVRET = 0  then 0 else (RCVRETNET - RCVCOST) / RCVRET end as RCVMargin" & Chr(10)
    strSQL = strSQL & ", case when isnull(RCVQTY,0) = 0 then 0 else RCVCOST / RCVQTY end CPU" & Chr(10)
    strSQL = strSQL & ", case when isnull(RCVPYQTY,0) = 0 or isnull(RCVQTY,0) = 0 then 0 else ((RCVCOST / RCVQTY) - (RCVPYCOST / RCVPYQTY)) / (RCVPYCOST / RCVPYQTY) end CPUYOY,isnull(POSPYRET,0) as POSPYRET" & Chr(10)
    strSQL = strSQL & ", case when isnull(RCVPYQTY,0) = 0 or isnull(RCVQTY,0) = 0 then 0 else(RCVPYCOST / RCVPYQTY) end as CPUPY" & Chr(10)
    strSQL = strSQL & ", case when RCVPYRET = 0  then 0 else (RCVPYRETNET - RCVPYCOST) / RCVPYRET end as RCVPYMargin" & Chr(10)
    strSQL = strSQL & ", (Select description from cbis599p.dbo.subcommoditygroup where cgno = @CG and scgno = @SCG)" & Chr(10)
    strSQL = strSQL & "from #BASE  order by year desc, month desc" & Chr(10)
    strSQL = strSQL & "drop table #RCV,#P, #pos, #BASE" & Chr(10)

    CBA_ErrTag = "SQL"
    RS.Open strSQL, CBA_COM_CBISCN
    If RS.EOF Then
        ContainsData = False
        Exit Sub
    Else
        ContainsData = True
        CBA_ErrTag = ""
        With RS
            Do Until .EOF
                lYIdx = BTF_GetElDate(NZ(.Fields(0), 0))
                lMn = NZ(.Fields(1), 0)
                If SCGDesc = "" Then SCGDesc = NZ(.Fields(13), 0)
                CBA_BTF_CGD((lYIdx * 12) + lMn).Year = NZ(.Fields(0), 0)
                CBA_BTF_CGD((lYIdx * 12) + lMn).Month = NZ(.Fields(1), 0)
                CBA_BTF_CGD((lYIdx * 12) + lMn).POSQTY = NZ(.Fields(2), 0)
                CBA_BTF_CGD((lYIdx * 12) + lMn).POSPYQTY = NZ(.Fields(3), 0)
                CBA_BTF_CGD((lYIdx * 12) + lMn).POSYOYQTY = NZ(.Fields(4), 0)
                CBA_BTF_CGD((lYIdx * 12) + lMn).POSRET = NZ(.Fields(5), 0)
                CBA_BTF_CGD((lYIdx * 12) + lMn).POSYOYRET = NZ(.Fields(6), 0)
                CBA_BTF_CGD((lYIdx * 12) + lMn).RCVMargin = NZ(.Fields(7), 0)
                CBA_BTF_CGD((lYIdx * 12) + lMn).CPU = NZ(.Fields(8), 0)
                CBA_BTF_CGD((lYIdx * 12) + lMn).CPUYOY = NZ(.Fields(9), 0)
                CBA_BTF_CGD((lYIdx * 12) + lMn).POSPYRET = NZ(.Fields(10), 0)
                CBA_BTF_CGD((lYIdx * 12) + lMn).CPUPY = NZ(.Fields(11), 0)
                CBA_BTF_CGD((lYIdx * 12) + lMn).RCVPYMargin = NZ(.Fields(12), 0)
                
                ''default forecast values
                If PClassNo = 1 Or PClassNo = 4 Then
                    CBA_BTF_CGD(NZ((lYIdx * 12) + lMn, 0)).ForeCPU = NZ(.Fields(8), 0)
                    If NZ(.Fields(2), 0) = 0 Then CBA_BTF_CGD(NZ((lYIdx * 12) + lMn, 0)).ForePrice = 0 Else CBA_BTF_CGD(NZ((lYIdx * 12) + lMn, 0)).ForePrice = NZ(.Fields(5), 0) / .Fields(2)
                    CBA_BTF_CGD(NZ((lYIdx * 12) + lMn, 0)).ForeQTY = NZ(.Fields(2), 0)
                End If
                
                
                .MoveNext
            Loop
        End With
        RS.Close
        Set RS = Nothing
    End If
    
    Set RS = New ADODB.Recordset
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
    strSQL = strSQL & "DECLARE @ENTEREDDATE date = '" & Year(dtDate) & "-" & Month(dtDate) & "-01'" & Chr(10)
    strSQL = strSQL & "DECLARE @CG int = " & CGno & Chr(10)
    strSQL = strSQL & "DECLARE @SDATE date = @ENTEREDDATE" & Chr(10)
    strSQL = strSQL & "DECLARE @EDATE date = dateadd(D,-1,dateadd(M,1,@SDATE))" & Chr(10)
    strSQL = strSQL & "" & Chr(10)
    strSQL = strSQL & "select bd.firstname + ' ' + bd.name as BD, gbd.firstname + ' ' + gbd.name as GBD from (" & Chr(10)
    strSQL = strSQL & "select p.empno, row_number()  over (order by count(p.empno) desc) as row from cbis599p.dbo.pos pos" & Chr(10)
    strSQL = strSQL & "left join cbis599p.dbo.product p on p.productcode = pos.productcode" & Chr(10)
    strSQL = strSQL & "where pos.posdate >= dateadd(D,1,dateadd(M,-1,@EDATE)) and pos.posdate <= @EDATE and p.CGno = @CG " & Chr(10)
    strSQL = strSQL & "group by p.empno) a  left join cbis599p.dbo.employee as bd on bd.empno = a.empno" & Chr(10)
    strSQL = strSQL & "left join cbis599p.dbo.employee as gbd on gbd.EmpNo = bd.EmpNo_Grp" & Chr(10)
    CBA_ErrTag = "SQL"
    RS.Open strSQL, CBA_COM_CBISCN
    If RS.EOF Then
    Else
        BDName = NZ(RS.Fields(0), "")
        GBDName = NZ(RS.Fields(1), "")
    End If
    ' Get the POS and Ret sales data from the current year and the last year
    If PClassNo = 2 Or PClassNo = 3 Then
        CBA_ErrTag = ""
        For lY = Year(dtDate) - 1 To Year(dtDate)               ' @RWFC 200107 - Date changed to input date (dtDate)
            For lMn = 1 To 12
                lM = lMn
                Set RS = New ADODB.Recordset
                strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
                strSQL = strSQL & "DECLARE @DTEFR DATE = '" & lY & "-" & lM & "-01'" & Chr(10)
                strSQL = strSQL & "DECLARE @DTETO DATE = dateadd(D,-1,dateadd(M,1,@DTEFR))" & Chr(10)
                strSQL = strSQL & "DECLARE @CG INT = " & CGno & Chr(10)
                strSQL = strSQL & "DECLARE @SCG INT = " & SCGno & Chr(10)
                strSQL = strSQL & "DECLARE @PCLASS INT = " & PClassNo & Chr(10)
                strSQL = strSQL & "select cgno, scgno, sum(Retail) as retail, sum(Cost) as Cost, sum(NetRetail) as RetailNet, sum(IM75) + sum(IM20) as IM, (sum(Cost) / sum(Retail)) * (sum(IM75) + sum(IM20)) as IMCost" & Chr(10)
                strSQL = strSQL & ", sum(RNIM75) + sum(RNIM20) as IMRN, sum(NM75) + sum(NM20) as NM, (sum(Cost) / sum(Retail)) * (sum(NM75) + sum(NM20)) as NMCost, sum(RNNM75) + sum(RNNM20) as NMRN,  sum(Retail)/20 as MDLoss" & Chr(10)
                strSQL = strSQL & ",round((sum(QTY) / sum(Retail)) * (sum(IM75) + sum(IM20)),0) as IMQTY, round((sum(QTY) / sum(Retail)) * (sum(NM75) + sum(NM20)),0) as NMQTY" & Chr(10)
                strSQL = strSQL & "from (SELECT dos.bd as BD, dos.BD_name as BD_Name, dos.cgno, dos.scgno, DOS.OSD, DOS.PRODUCTCODE,DOS.DESCR, SUM(DOS.SALES) AS 'Retail', sum (dos.cost) as Cost, sum(dos.netretail) as NetRetail, sum(dos.netretail-dos.cost) as Contr, (dos.netretail-dos.cost)/nullif(dos.sales2,0) as Margin, dos.theme, dos.retail1 as Price, Sum(QTY) as QTY" & Chr(10)
                strSQL = strSQL & ",(((SUM(DOS.SALES) / 4)*3) / 14) * case when DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) > 14 then 14 else DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) end as IM75" & Chr(10)
                strSQL = strSQL & ",(((SUM(DOS.SALES) / 4)*3) / 14) * case when DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) < 14 then 14 - DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) else 0 end as NM75" & Chr(10)
                strSQL = strSQL & ",((SUM(DOS.SALES) / 5) / 16) * (case when DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) > 30 then 30 else DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) end - case when DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) > 14 then 14 else DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) end) as IM20" & Chr(10)
                strSQL = strSQL & ",((SUM(DOS.SALES) / 5) / 16) * (case when DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) < 30 then 30 - DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) else 0 end - case when DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) < 14 then 14 - DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) else 0 end) as NM20" & Chr(10)
                strSQL = strSQL & ",(((SUM(dos.netretail) / 4) * 3) / 14) * case when DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) > 14 then 14 else DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) end as RNIM75" & Chr(10)
                strSQL = strSQL & ",(((SUM(dos.netretail) / 4) * 3) / 14) * case when DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) < 14 then 14 - DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) else 0 end as RNNM75" & Chr(10)
                strSQL = strSQL & ",((SUM(dos.netretail) / 5) / 16) * (case when DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) > 30 then 30 else DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) end - case when DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) > 14 then 14 else DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) end) as RNIM20" & Chr(10)
                strSQL = strSQL & ",((SUM(dos.netretail) / 5) / 16) * (case when DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) < 30 then 30 - DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) else 0 end - case when DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) < 14 then 14 - DATEDIFF(DAY,DOS.OSD, EOMONTH(DOS.OSD,0)) else 0 end) as RNNM20" & Chr(10)
                strSQL = strSQL & "FROM (SELECT em.empsign as BD, em.name as BD_Name,sptf.cgno as cgno, sptf.scgno as scgno, SRVR.GroupAdvdate AS OSD, SRVR.PRODUCTCODE,ppl.description AS DESCR, ROW_NUMBER() OVER (partition by srvr.groupadvdate,srvr.portfolioid, srvr.productcode, srvr.quantity1,srvr.retail1,ppl.description order by srvr.groupadvdate,srvr.portfolioid, srvr.productcode, srvr.quantity1,srvr.retail1,ppl.description) AS rowNumber, (SRVR.QUANTITY1 * SRVR.RETAIL1) AS 'SALES',(case when srvr.cost1>0 then SRVR.QUANTITY1 * SRVR.retail1 else 0 end) AS 'Sales2', (case when srvr.cost1>0 then SRVR.QUANTITY1 * SRVR.COST1 else 0 end) AS 'COST', (case when srvr.cost1>0 then SRVR.QUANTITY1 * SRVR.netRETAIL1 else 0 end) AS 'NetRetail', sscn.description as theme, srvr.retail1 as Retail1, SRVR.QUANTITY1 as QTY" & Chr(10)
                strSQL = strSQL & "FROM [cbis599p].[portfolio].[Rep_PfVersionReg] as SRVR LEFT JOIN [CBIS599p].[portfolio].[Portfolio] as SPTF ON SRVR.PortfolioID = SPTF.PortfolioID" & Chr(10)
                strSQL = strSQL & "LEFT JOIN [cbis599p].[portfolio].[Pfversion] as PFV ON SRVR.PortfolioID = PFV.PortfolioID AND SRVR.PFVERSIONID = PFV.PFVERSIONID" & Chr(10)
                strSQL = strSQL & "LEFT JOIN [cbis599p].[portfolio].[PfVersionMapping] AS PFVM ON SRVR.PortfolioID = PFVM.PortfolioID AND SRVR.PFVERSIONID = PFVM.PFVERSIONID" & Chr(10)
                strSQL = strSQL & "LEFT JOIN [cbis599p].[portfolio].[PfVersionReg] AS PFVR ON SRVR.PortfolioID = PFVR.PortfolioID AND SRVR.PFVERSIONID = PFVR.PFVERSIONID" & Chr(10)
                strSQL = strSQL & "LEFT JOIN [CBIS599p].[Dbo].[SPECIALCATEGORY] as SSCN ON PFVR.SpecialCatergoryno = SSCN.SpecialCatergoryNo" & Chr(10)
                strSQL = strSQL & "LEFT JOIN [cbis599p].[portfolio].[pfversionlng] AS PPL ON SRVR.PortfolioID = ppl.PortfolioID  AND SRVR.PFVERSIONID = ppl.PFVERSIONID and ppl.LANGUAGEID=0" & Chr(10)
                strSQL = strSQL & "LEFT JOIN [cbis599p].[dbo].[employee] AS EM ON SPTF.empno = Em.empno WHERE pfvr.advertisingdate >= @DTEFR AND pfvr.advertisingdate <= @DTETO" & Chr(10)
                strSQL = strSQL & "AND PFV.PFSTATUSID IS NOT NULL AND PFVR.PRODUCTCLASS = @PCLASS AND SPTF.CGNo = @CG and SPTF.SCGNo = @SCG" & Chr(10)
                strSQL = strSQL & "GROUP BY em.empsign, em.name, sptf.cgno, sptf.scgno, SRVR.GroupAdvdate, SSCN.dESCRIPTION, PFVM.PRODUCTCODE, ppl.description, srvr.portfolioid, SRVR.Productcode , SRVR.QUANTITY1 , SRVR.RETAIL1, srvr.cost1, srvr.netretail1, sscn.description, srvr.retail1) AS DOS" & Chr(10)
                strSQL = strSQL & "where Rownumber = 1 GROUP BY dos.bd,  dos.BD_name, dos.cgno, dos.scgno, DOS.OSD, DOS.PRODUCTCODE, DOS.DESCR, dos.theme, (dos.netretail-dos.cost)/nullif(dos.sales2,0), dos.retail1) a group by cgno, scgno" & Chr(10)
                CBA_ErrTag = "SQL"
                RS.Open strSQL, CBA_COM_CBISCN
                If RS.EOF Then
                Else
                    CBA_ErrTag = ""
                    With RS
                        lYIdx = BTF_GetElDate(lY)
                        If lMn > 1 Then
                            CBA_BTF_CGD((lYIdx * 12) + lM).ForeQTY = CBA_BTF_CGD((lYIdx * 12) + lM).ForeQTY + NZ(.Fields(11), 0)
                            CBA_BTF_CGD((lYIdx * 12) + lM).ForeRetail = CBA_BTF_CGD((lYIdx * 12) + lM).ForeRetail + NZ(.Fields(5), 0)
                            CBA_BTF_CGD((lYIdx * 12) + lM).ForeCost = CBA_BTF_CGD((lYIdx * 12) + lM).ForeCost + NZ(.Fields(6), 0)
                            CBA_BTF_CGD((lYIdx * 12) + lM).ForeRetailNet = CBA_BTF_CGD((lYIdx * 12) + lM).ForeRetailNet + NZ(.Fields(7), 0)
                        End If
                        If lM + 1 = 13 Then lv = 1 Else lv = lM + 1
                        If lMn < 13 Then
                            CBA_BTF_CGD((lYIdx * 12) + lv).ForeQTY = CBA_BTF_CGD((lYIdx * 12) + lv).ForeQTY + NZ(.Fields(12), 0)
                            CBA_BTF_CGD((lYIdx * 12) + lv).ForeRetail = CBA_BTF_CGD((lYIdx * 12) + lv).ForeRetail + NZ(.Fields(8), 0)
                            CBA_BTF_CGD((lYIdx * 12) + lv).ForeCost = CBA_BTF_CGD((lYIdx * 12) + lv).ForeCost + NZ(.Fields(9), 0)
                            CBA_BTF_CGD((lYIdx * 12) + lv).ForeRetailNet = CBA_BTF_CGD((lYIdx * 12) + lv).ForeRetailNet + NZ(.Fields(10), 0)
                        End If
                    End With
                End If
            Next
        Next
        bFCNet = True
        For lYIdx = 0 To 2
            For lMn = 1 To 12
                If CBA_BTF_CGD((lYIdx * 12) + lMn).ForeRetail <> 0 Then
                    CBA_BTF_CGD((lYIdx * 12) + lMn).ForeRCVMargin = (CBA_BTF_CGD((lYIdx * 12) + lMn).ForeRetailNet - CBA_BTF_CGD((lYIdx * 12) + lMn).ForeCost) / CBA_BTF_CGD((lYIdx * 12) + lMn).ForeRetail
                End If
            Next
        Next
    End If
    CBA_ErrTag = ""
    ' Get any existing forecast data
    bFCNet = getForecastedSCGs(lYY)
    ' If Product Class = 1 or 4, and the values are 0, copy the values from the prior year
    If PClassNo = 1 Or PClassNo = 4 Then
        If Not bFCNet Then
            For lYIdx = 0 To 2
                For lMn = 1 To 12
                    If NZ(CBA_BTF_CGD((lYIdx * 12) + lMn).ForeRetail, 0) = 0 Then
                        CBA_BTF_CGD((lYIdx * 12) + lMn).ForeRetail = NZ(CBA_BTF_CGD((lYIdx * 12) + lMn).POSRET, 0)
                        Call UpdateValue("ForeRetail", CBA_BTF_CGD((lYIdx * 12) + lMn).ForeRetail, lMn, lYIdx)
                    End If
                    If NZ(CBA_BTF_CGD((lYIdx * 12) + lMn).ForeRCVMargin, 0) = 0 Then CBA_BTF_CGD((lYIdx * 12) + lMn).ForeRCVMargin = NZ(CBA_BTF_CGD((lYIdx * 12) + lMn).RCVMargin, 0)
                    Call UpdateValue("ForeRetail", CBA_BTF_CGD((lYIdx * 12) + lMn).ForeRetail, lMn, lYIdx)
                Next
            Next
        End If
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CreateSCG", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & strSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , , True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub

Private Function getForecastedSCGs(lYY As Long) As Boolean

    ' This routine will load the data from the Forecast table CGData and overwrite the data in the arrays
    Dim sSQL As String, dtDateSaved As Date, lNoOfDates As Long, bFCast As Boolean, bReFCast As Boolean, lYear As Long '',bDiffers As Boolean
    Dim RS As ADODB.Recordset, CN As ADODB.Connection, lMthNo As Long, lYYNo As Long, lTimes As Long, lElNo As Long, bSkip As Boolean
    Dim curFRetail As Currency, curFReRetail As Currency, curT1Retail As Currency, curT2Retail As Currency '', sReturn As String
    Dim curARetail(1 To 12) As Currency, curAMargin(1 To 12) As Single, curAUplift(1 To 12) As Single
    On Error GoTo Err_Routine
    CBA_ErrTag = "SQL"
    
    Set CN = New ADODB.Connection
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ForeCast") & ";"
    Set RS = New ADODB.Recordset
    getForecastedSCGs = False
    ' Query will get the last Product Data of each Product (OrderBY is in the qry
    sSQL = "SELECT * FROM SCGData WHERE  CG = " & CGno & " AND  SCG = " & SCGno & " AND ProductClass = " & PClassNo & " And YearNo > " & lYY & _
            " ORDER BY CG, SCG, ProductClass, YearNo DESC, DateTimeSubmitted DESC, MonthNo DESC;"
    RS.Open sSQL, CN
    lTimes = 0: lNoOfDates = 0: bFCast = False: bReFCast = False: lYear = 0: curT1Retail = 0: curT2Retail = 0
    Do While Not RS.EOF
        ' On change of Date...
        If dtDateSaved <> RS!DateTimeSubmitted Then
            dtDateSaved = RS!DateTimeSubmitted
            getForecastedSCGs = True
            ' If the year has changed
            If lYear <> RS!YearNo Then
                If lYY = Year(dtDateSaved) And bFCast = True And bReFCast = True Then GoTo Exit_Routine
                bFCast = False: bReFCast = False: lTimes = 0: lNoOfDates = 0
                lYear = RS!YearNo
                If lYear = 2020 Then
                    lYear = lYear
                End If
            End If
            ' What type of record is it? Forecast or Reforecast
            bSkip = False
            curFReRetail = NZ(RS!FReRetail, 0)
            curFRetail = NZ(RS!FRetail, 0)
            ' If the type of record is uncertain, do further tests...
            If curFReRetail + curFRetail = 0 Then
                sSQL = "CG = " & CGno & " AND  SCG = " & SCGno & " AND ProductClass = " & PClassNo & " And DateTimeSubmitted =" & g_GetSQLDate(RS!DateTimeSubmitted, CBA_DMYHN)
                curFRetail = g_DLookup("SumRetail", "qry_SCGTots", sSQL, "DateTimeSubmitted", g_GetDB("ForeCast"), 0)
                If curFRetail = 0 Then curFReRetail = 1
            End If
            ' If a reforecast...
            If curFReRetail > 0 Then
                If bReFCast = True Then     ' If a reforecast has already been done for the year, flag to skip the set of records
                    bSkip = True
                Else
                    bReFCast = True         ' Flag to capture values
                End If
            End If
            ' If a forecast...
            If curFRetail > 0 Then
                If bFCast = True Then       ' If a forecast has already been done for the year, flag to skip the set of records
                    bSkip = True
                Else
                    bFCast = True           ' Flag to capture values
                End If
            End If
            lNoOfDates = lNoOfDates + 1
            If lTimes > 0 Then
                If lTimes / 12 <> lTimes \ 12 Then
                    CBA_Error = "Non fatal error-Forcast has changed but records are not a multiple of 12 - SQL follows... " & vbCrLf & sSQL
                    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
                    Debug.Print CBA_Error
                End If
            End If
''            If lNoOfDates > 2 And bDiffers = True Then
''                If lYY + 1 = Year(dtDateSaved) Then GoTo Exit_Routine
''                lTimes = 0: lNoOfDates = 0: bDiffers = False
''            End If
        End If
        If bSkip Then GoTo SkipToNext
        lTimes = lTimes + 1: CBA_Error = ""
        lMthNo = Month(RS!ForecastDate)
        lYYNo = Year(DateAdd("YYYY", -1, RS!ForecastDate))
        lYYNo = BTF_GetElDate(lYYNo)
        lElNo = (lYYNo * 12) + lMthNo
        If (bReFCast = True And bFCast = False) Or (bReFCast = False And bFCast = True) Then
            If bFCast = True Then
                curT1Retail = curT1Retail + NZ(RS!FRetail, 0)
                curARetail(lMthNo) = NZ(RS!FRetail, 0)
                curAUplift(lMthNo) = NZ(RS!FUplift, 0)
                curAMargin(lMthNo) = NZ(RS!FMarginP, 0)
            End If
            Call UpdateValue("ForeRetail+", IIf(NZ(RS!FReRetail, 0) > 0, NZ(RS!FReRetail, 0), NZ(RS!FRetail, 0)), lMthNo, lYYNo)
            Call UpdateValue("Uplift+", IIf(NZ(RS!FReUplift, 0) > 0, NZ(RS!FReUplift, 0), NZ(RS!FUplift, 0)), lMthNo, lYYNo)
            Call UpdateValue("ForeRCVMargin", IIf(NZ(RS!FReMarginP, 0) > 0, NZ(RS!FReMarginP, 0), NZ(RS!FMarginP, 0)), lMthNo, lYYNo)
''            If NZ(RS!FReRetail, 0) = 0 And NZ(RS!FRetail, 0) > 0 Then
''                Call UpdateValue("PForeRetail", NZ(RS!FRetail, 0), lMthNo, lYYNo)
''                Call UpdateValue("PForeRCVMargin", NZ(RS!FMarginP, 0), lMthNo, lYYNo)
''                Call UpdateValue("PUplift", NZ(RS!FUplift, 0), lMthNo, lYYNo)
''            End If
        ElseIf bFCast = True Then
            curT2Retail = curT2Retail + NZ(RS!FRetail, 0)
            Call UpdateValue("PForeRetail", NZ(RS!FRetail, 0), lMthNo, lYYNo)
            Call UpdateValue("PForeRCVMargin", NZ(RS!FMarginP, 0), lMthNo, lYYNo)
            Call UpdateValue("PUplift", NZ(RS!FUplift, 0), lMthNo, lYYNo)
''            If Me.Differs(lElNo) Then bDiffers = True
        End If

SkipToNext:
        RS.MoveNext
    Loop
    ' If there are no prior forecasted values, then insert them
    If curT1Retail > 0 And curT2Retail = 0 Then
        For lMthNo = 1 To 12
            Call UpdateValue("PForeRetail", curARetail(lMthNo), lMthNo, lYYNo)
            Call UpdateValue("PForeRCVMargin", curAMargin(lMthNo), lMthNo, lYYNo)
            Call UpdateValue("PUplift", curAUplift(lMthNo), lMthNo, lYYNo)
        Next
    End If
    
Exit_Routine:
    On Error Resume Next
    Set RS = Nothing
    Set CN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-getForecastedSCGs", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function


Property Get CG() As Long
    CG = CGno
End Property
Property Get BD() As String
    BD = BDName
End Property
Property Get GBD() As String
    GBD = GBDName
End Property
Property Get YearNo(ByVal lElNo) As Long
    YearNo = CBA_BTF_CGD(lElNo).Year
End Property
Property Get Description() As String
    Description = SCGDesc
End Property
Property Get POSQTY(ByVal lElNo) As Long
    POSQTY = CBA_BTF_CGD(lElNo).POSQTY
End Property
Property Get POSPYQTY(ByVal lElNo) As Long
    POSPYQTY = CBA_BTF_CGD(lElNo).POSPYQTY
End Property
Property Get POSYOYQTY(ByVal lElNo) As Single
    POSYOYQTY = CBA_BTF_CGD(lElNo).POSYOYQTY
End Property
Property Get POSRET(ByVal lElNo) As Single
    POSRET = CBA_BTF_CGD(lElNo).POSRET
End Property
Property Get POSYOYRET(ByVal lElNo) As Single
    POSYOYRET = CBA_BTF_CGD(lElNo).POSPYRET
End Property
Property Get RCVMargin(ByVal lElNo) As Single
    RCVMargin = CBA_BTF_CGD(lElNo).RCVMargin
End Property
Property Get RCVPYMargin(ByVal lElNo) As Single
    RCVPYMargin = CBA_BTF_CGD(lElNo).RCVPYMargin
End Property
Property Get CPU(ByVal lElNo) As Single
    CPU = CBA_BTF_CGD(lElNo).CPU
End Property
Property Get CPUYOY(ByVal lElNo) As Single
    CPUYOY = CBA_BTF_CGD(lElNo).CPUYOY
End Property
Property Get POSPYRET(ByVal lElNo) As String
    POSPYRET = CBA_BTF_CGD(lElNo).POSPYRET
End Property
Property Get CPUPY(ByVal lElNo) As Single
    CPUPY = CBA_BTF_CGD(lElNo).CPUPY
End Property
Property Get ForeCPU(ByVal lElNo) As Single
    ForeCPU = CBA_BTF_CGD(lElNo).ForeCPU
End Property
Property Get ForeAvgPrice(ByVal lElNo) As Single
    ForeAvgPrice = CBA_BTF_CGD(lElNo).ForePrice
End Property
Property Get ForeQTY(ByVal lElNo) As Long
    ForeQTY = CBA_BTF_CGD(lElNo).ForeQTY
End Property
Property Get ForeRetail(ByVal lElNo) As Single
    ForeRetail = CBA_BTF_CGD(lElNo).ForeRetail
End Property
Property Get PForeRetail(ByVal lElNo) As Single
    PForeRetail = CBA_BTF_CGD(lElNo).PForeRetail
End Property
Property Get ForeRCVMargin(ByVal lElNo) As Single
    ForeRCVMargin = CBA_BTF_CGD(lElNo).ForeRCVMargin
End Property
Property Get PForeRCVMargin(ByVal lElNo) As Single
    PForeRCVMargin = CBA_BTF_CGD(lElNo).PForeRCVMargin
End Property
Property Get PUplift(ByVal lElNo) As Single
    PUplift = CBA_BTF_CGD(lElNo).PUplift
End Property
Property Get Uplift(ByVal lElNo) As Single
    Uplift = CBA_BTF_CGD(lElNo).Uplift
End Property
Property Get isDataContained() As Boolean
    isDataContained = ContainsData
End Property

''' Functions
''Function Differs(ByVal lElNo As Long) As Boolean
''    Differs = False
''    If CBA_BTF_CGD(lElNo).ForeRetail <> CBA_BTF_CGD(lElNo).PForeRetail And CBA_BTF_CGD(lElNo).PForeRetail > 0 Then Differs = True
''    If CBA_BTF_CGD(lElNo).ForeRCVMargin <> CBA_BTF_CGD(lElNo).PForeRCVMargin And CBA_BTF_CGD(lElNo).PForeRCVMargin > 0 Then Differs = True
''End Function

Sub UpdateValue(sField As String, ByVal NewValue, ByVal lMonthNo As Long, ByVal lYearNo As Long)
    Dim lElNo As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    ' As there are 36 elements now, we have to use the year to decide which set of elements are to be updated
    lYearNo = BTF_GetElDate(lYearNo)
    lElNo = (lYearNo * 12) + lMonthNo

    ' Will update the values into the array
    Select Case sField
    Case Is = "POSQTY"
       CBA_BTF_CGD(lElNo).POSQTY = NewValue           '.POSQTY
    Case Is = "POSPYQTY"
       CBA_BTF_CGD(lElNo).POSPYQTY = NewValue         '.POSPYQTY
    Case Is = "POSYOYQTY"
       CBA_BTF_CGD(lElNo).POSYOYQTY = NewValue        '.POSYOYQTY
    Case Is = "POSRET"
       CBA_BTF_CGD(lElNo).POSRET = NewValue           '.POSRET
    Case Is = "POSYOYRET"
       CBA_BTF_CGD(lElNo).POSYOYRET = NewValue        '.POSPYRET
    Case Is = "RCVMargin"
       CBA_BTF_CGD(lElNo).RCVMargin = NewValue        '.RCVMargin
    Case Is = "RCVPYMargin"
       CBA_BTF_CGD(lElNo).RCVPYMargin = NewValue      '.RCVPYMargin
    Case Is = "CPU"
       CBA_BTF_CGD(lElNo).CPU = NewValue              '.CPU
    Case Is = "CPUYOY"
       CBA_BTF_CGD(lElNo).CPUYOY = NewValue           '.CPUYOY
    Case Is = "POSPYRET"
       CBA_BTF_CGD(lElNo).POSPYRET = NewValue         '.POSPYRET
    Case Is = "CPUPY"
       CBA_BTF_CGD(lElNo).CPUPY = NewValue            '.CPUPY
    Case Is = "ForeCPU"
       CBA_BTF_CGD(lElNo).ForeCPU = NewValue          '.ForeCPU
    Case Is = "ForeAvgPrice"
       CBA_BTF_CGD(lElNo).ForePrice = NewValue        '.ForePrice
    Case Is = "ForeQTY"
       CBA_BTF_CGD(lElNo).ForeQTY = NewValue          '.ForeQTY
    Case Is = "ForeRetail"
       CBA_BTF_CGD(lElNo).ForeRetail = NewValue       '.ForeRetail
       If CBA_BTF_CGD(lElNo).POSRET <> 0 Then
            CBA_BTF_CGD(lElNo).Uplift = Round(((CBA_BTF_CGD(lElNo).ForeRetail - CBA_BTF_CGD(lElNo).POSRET) / CBA_BTF_CGD(lElNo).POSRET), 4)
       End If
    Case Is = "ForeRetail+"
       CBA_BTF_CGD(lElNo).ForeRetail = NewValue       '.ForeRetail
    Case Is = "PForeRetail"
       CBA_BTF_CGD(lElNo).PForeRetail = NewValue      '.PForeRetail
    Case Is = "ForeRetailNet"
       CBA_BTF_CGD(lElNo).ForeRetailNet = NewValue    '.ForeRetailNet
    Case Is = "ForeRCVMargin"
       CBA_BTF_CGD(lElNo).ForeRCVMargin = NewValue    '.ForeRCVMargin
    Case Is = "PForeRCVMargin"
       CBA_BTF_CGD(lElNo).PForeRCVMargin = NewValue   '.PForeRCVMargin
    Case Is = "PUplift"
       If CBA_BTF_CGD(lElNo).POSRET <> 0 Then
            CBA_BTF_CGD(lElNo).PUplift = Round(((CBA_BTF_CGD(lElNo).PForeRetail - CBA_BTF_CGD(lElNo).POSRET) / CBA_BTF_CGD(lElNo).POSRET), 4)
       End If
    Case Is = "ForeCost"
       CBA_BTF_CGD(lElNo).ForeCost = NewValue         '.ForeCost
    Case Is = "Uplift"
       CBA_BTF_CGD(lElNo).Uplift = NewValue           '.Uplift
       CBA_BTF_CGD(lElNo).ForeRetail = Round(CBA_BTF_CGD(lElNo).POSRET * (1 + (CBA_BTF_CGD(lElNo).Uplift)), 4)
    Case Is = "Uplift+"
       CBA_BTF_CGD(lElNo).Uplift = NewValue           '.Uplift
    Case Else
        Err.Raise 513, , sField & " not found in SCG UpdateValue"
    End Select
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-UpdateValue (SCG)", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub




