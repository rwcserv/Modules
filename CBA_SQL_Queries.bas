Attribute VB_Name = "CBA_SQL_Queries"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Function CBA_GenPullSQL(ByVal CBA_SQLType As String, Optional ByRef CBA_datefrom, _
    Optional ByRef CBA_Dateto, Optional ByVal CBA_LongVal1 As Long, _
    Optional ByVal CBA_LongVal2 As Long, Optional ByVal CBA_LongVal3 As Long, _
    Optional ByVal CBA_LongVal4 As Long, Optional ByVal CBA_StringVal1 As String, _
    Optional ByVal CBA_StringVal2 As String, Optional ByVal CBA_StringVal3 As String, _
    Optional ByVal CBA_StringVal4 As String, Optional ByVal CBA_BooleanVal1 As Boolean, _
    Optional ByVal CBA_BooleanVal2 As Boolean, Optional ByVal CBA_VariantVal1 As Variant, _
    Optional ByVal CBA_VariantVal2 As Variant, Optional ByVal CBA_DoubleVal1 As Single, _
    Optional ByVal CBA_DoubleVal2 As Single, Optional ByVal CBA_DoubleVal3 As Single, Optional ByVal CBA_DoubleVal4 As Single)
    'CBA_SQLType  - Defines the SQL Query used)
    Dim CBA_strSQL As String
    Dim bOutput As Boolean
    Dim CBA_QueryTimerStart As String, strCGSQL As String
    Dim bytCurCG As Byte
    Dim a As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    If CBA_DBtoQuery = 0 Then CBA_DBtoQuery = 599
    CBA_QueryTimerStart = Now
    
    If Mid(CBA_SQLType, 1, 4) = "CBIS" Then CBA_DBtoQuery = 599
    
    
    
    Select Case CBA_SQLType
    
        Case "CBIS_LDT"
    
        CBA_strSQL = "SET NOCOUNT ON" & Chr(10)
        CBA_strSQL = CBA_strSQL & "SET ANSI_WARNINGS OFF" & Chr(10)
        CBA_strSQL = CBA_strSQL & "SET DATEFIRST 1" & Chr(10)
        CBA_strSQL = CBA_strSQL & "DECLARE @DATES date DECLARE @DATEE date DECLARE @PROD int DECLARE @WEEKS float DECLARE @DAYS int" & Chr(10)
        CBA_strSQL = CBA_strSQL & "SET @PROD =  " & CBA_LongVal1 & Chr(10)
        CBA_strSQL = CBA_strSQL & "SET @DATES = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
        CBA_strSQL = CBA_strSQL & "SET @DATEE = '" & Format(CBA_Dateto, "YYYY-MM-DD") & "'" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "SET @DAYS = DATEDIFF(D,@DATES,dateadd(D,1,@DATEE)) " & Chr(10)
        CBA_strSQL = CBA_strSQL & "SET @WEEKS = convert(float,@DAYS) / 7" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select productcode into #PRODS from cbis599p.dbo.product where isnull(con_productcode, productcode) = @PROD" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select pos.productcode, divno into #DIVS from cbis599p.dbo.pos left join cbis599p.dbo.PRODUCT p on p.productcode = pos.ProductCode " & Chr(10)
        CBA_strSQL = CBA_strSQL & "where posdate >= @DATES and posdate <= @DATEE and p.ProductCode in (select * from #PRODS) group by pos.productcode, divno having sum(pos.Retail) > 5000" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select productcode, description, cgno, scgno, productclass, isRandomWeight, TaxID into #P from cbis599p.dbo.product where productcode in (select * from #PRODS)" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select divno, datepart(ISOWK, validfrom) as Weekno, year(validfrom) as Yearno, NoofStores into #DIVSTORE " & Chr(10)
        CBA_strSQL = CBA_strSQL & "from cbis599p.portfolio.Stores where Validto >= dateadd(M,-1,@DATES) and validfrom <= dateadd(M,1,@DATEE) and divno in (select divno from #DIVS)" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select weekno, sum(noofStores) as noofStores into #NATSTORE from #DIVSTORE group by Weekno" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select divno , case when validfrom < @DATES then @DATES else validfrom end as validfrom" & Chr(10)
        CBA_strSQL = CBA_strSQL & ", case when isnull(validto,getdate()) > @DATEE then @DATEE else isnull(validto,getdate()) end as validto" & Chr(10)
        CBA_strSQL = CBA_strSQL & ", Retail into #RETAIL from cbis599p.dbo.divretail where productcode = @PROD and validfrom <= @DATEE and isnull(validto,getdate()) >= @DATES" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select rcv.divno, isnull(p.con_productcode, p.productcode) as productcode, sum(rcv.Quantity) as Quantity into #RCVQTY from cbis599p.dbo.RECEIVING rcv" & Chr(10)
        CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.product p on p.productcode = rcv.productcode" & Chr(10)
        CBA_strSQL = CBA_strSQL & "inner join #DIVS div on div.DivNo = rcv.DivNo and (div.ProductCode = p.con_productcode or div.productcode = p.productcode) " & Chr(10)
        CBA_strSQL = CBA_strSQL & "where DayEndDate >= dateadd(M,-4,@DATES) and DayEndDate <= @DATEE  and rcv.RecordID = '001' group by rcv.DivNo, isnull(p.con_productcode, p.productcode)" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select rcv.divno, isnull(p.con_productcode, p.productcode) as productcode, sum(Cost) / rcvq.Quantity as Cost into #COSTS from cbis599p.dbo.RECEIVING rcv" & Chr(10)
        CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.product p on p.productcode = rcv.productcode" & Chr(10)
        CBA_strSQL = CBA_strSQL & "inner join #DIVS div on div.DivNo = rcv.DivNo and (div.ProductCode = p.con_productcode or div.productcode = p.productcode) " & Chr(10)
        CBA_strSQL = CBA_strSQL & "left join #RCVQTY rcvq on rcvq.ProductCode = isnull(p.con_productcode, p.productcode) and rcvq.DivNo = rcv.DivNo" & Chr(10)
        CBA_strSQL = CBA_strSQL & "where DayEndDate >= dateadd(M,-4,@DATES) and DayEndDate <= @DATEE group by rcv.DivNo, isnull(p.con_productcode, p.productcode), rcvq.Quantity" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select datepart(ISOWK, posdate) as Weekno, sum(Retail) as TotRetail into #TOTBUS from cbis599p.dbo.pos pos left join cbis599p.dbo.product p on p.Productcode = pos.ProductCode " & Chr(10)
        CBA_strSQL = CBA_strSQL & "where posdate >= @DATES and posdate <= @DATEE " & Chr(10)
        CBA_strSQL = CBA_strSQL & "and divno in (Select divno from #DIVS) " & Chr(10)
        CBA_strSQL = CBA_strSQL & "group by datepart(ISOWK, posdate)" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select pos.divno, pos.productcode, posdate, sum(pos.Retail) as Retail, ds.NoOfStores, case when (select Taxid from #P where ProductCode = @PROD) = '01' then " & Chr(10)
        CBA_strSQL = CBA_strSQL & "case when count(r.retail) = 0 or sum(pos.Quantity) = 0 then 0 else (sum(pos.Retail) / sum(pos.Quantity) / (sum(r.Retail)/count(r.retail))) * sum(pos.Quantity) end Else sum(pos.Quantity) End as QTY, sum(pos.Quantity) as UnitQTY into #DAYPOS from cbis599p.dbo.pos " & Chr(10)
        CBA_strSQL = CBA_strSQL & "left join #DIVSTORE ds on ds.Weekno = datepart(ISOWK,posdate) and ds.DivNo = pos.DivNo inner join #DIVS div on div.productcode = pos.productcode and div.DivNo = pos.DivNo" & Chr(10)
        CBA_strSQL = CBA_strSQL & "left join #RETAIL r on r.DivNo = pos.divno and r.validfrom <= pos.posdate and r.validto >= pos.PosDate where posdate >= @DATES and posdate <= @DATEE group by  posdate, pos.productcode, pos.divno, ds.NoOfStores" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select isnull(p.con_productcode, pos.productcode) as productcode, pos.divno, datepart(ISOWK,posdate) as Weekno, sum(QTY) as QTY, sum(UnitQTY) as UnitQTY, sum(Retail) as Retail , cost ,ds.NoOfStores" & Chr(10)
        CBA_strSQL = CBA_strSQL & ",case when (select Taxid from #P tp where tp.productcode = isnull(p.con_productcode, pos.productcode)) = 1 then (sum(retail) - (sum(QTY) * cost)) / nullif(sum(retail),0) else ((sum(retail)/1.1) - (sum(QTY) * cost)) / nullif(sum(retail),0) end * sum(Retail) as MarginDol" & Chr(10)
        CBA_strSQL = CBA_strSQL & ", tb.TotRetail , sum(QTY) / nullif(ds.NoOfStores,0) as UpS" & Chr(10)
        CBA_strSQL = CBA_strSQL & ",  (case when (select Taxid from #P tp where tp.productcode = isnull(p.con_productcode, pos.productcode)) = 1 then (sum(retail) - (sum(QTY) * cost)) / nullif(sum(retail),0) else ((sum(retail)/1.1) - (sum(QTY) * cost)) / nullif(sum(retail),0) end * sum(Retail) ) / nullif(ds.NoOfStores,0) as MarginDolpS" & Chr(10)
        CBA_strSQL = CBA_strSQL & "into #BASEPOS from #DAYPOS pos" & Chr(10)
        CBA_strSQL = CBA_strSQL & "inner join cbis599p.dbo.product p on p.productcode = pos.productcode" & Chr(10)
        CBA_strSQL = CBA_strSQL & "inner join #DIVS d on d.DivNo = pos.DivNo and (d.ProductCode = pos.ProductCode or d.productcode = p.con_productcode)" & Chr(10)
        CBA_strSQL = CBA_strSQL & "left join #TOTBUS tb on datepart(ISOWK,posdate) = tb.Weekno" & Chr(10)
        CBA_strSQL = CBA_strSQL & "left join #COSTS c on c.DivNo = pos.DivNo and (c.ProductCode = pos.ProductCode or c.productcode = p.con_productcode)" & Chr(10)
        CBA_strSQL = CBA_strSQL & "left join #DIVSTORE ds on ds.Weekno = datepart(ISOWK,posdate) and ds.DivNo = pos.DivNo" & Chr(10)
        CBA_strSQL = CBA_strSQL & "where posdate >= @DATES and posdate <= @DATEE group by isnull(p.con_productcode, pos.productcode),pos.divno, datepart(ISOWK,posdate), tb.TotRetail, Cost, ds.NoOfStores" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select divno, weekno, sum(QTY) as QTY, sum(Retail) as Retail, sum(QTY * cost) as Cost, noofstores, sum(MarginDol) as MarginDol" & Chr(10)
        CBA_strSQL = CBA_strSQL & ", TotRetail, sum(MarginDolpS) as MarginDolpS, sum(MarginDolpS) / @Weeks as MarginDolpSW, sum(Ups) as UpS" & Chr(10)
        CBA_strSQL = CBA_strSQL & ", sum(Ups) / @Weeks as UpSW, @Weeks as Weeks, (select Taxid from #P where productcode = @PROD) as taxid, sum(UnitQTY) as unitQTY from #BASEPOS group by DivNo, weekno, noofstores, TotRetail" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "drop table #PRODS, #DIVS, #P, #DIVSTORE,  #NATSTORE, #RETAIL , #RCVQTY, #COSTS, #TOTBUS, #DAYPOS, #BASEPOS" & Chr(10)
    
    
        Case "MMS_MDLoss"
        CBA_strSQL = "SET NOCOUNT ON" & Chr(10)
        CBA_strSQL = CBA_strSQL & "SET ANSI_WARNINGS OFF" & Chr(10)
        CBA_strSQL = CBA_strSQL & "SET DATEFIRST 1" & Chr(10)
        CBA_strSQL = CBA_strSQL & "DECLARE @DATES date DECLARE @DATEE date DECLARE @PROD int" & Chr(10)
        CBA_strSQL = CBA_strSQL & "SET @PROD = " & CBA_LongVal1 & Chr(10)
        CBA_strSQL = CBA_strSQL & "SET @DATES = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
        CBA_strSQL = CBA_strSQL & "SET @DATEE = '" & Format(CBA_Dateto, "YYYY-MM-DD") & "'" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select productcode into #PRODS from purchase.dbo.product where coalesce(con_productcode,Mainproduct, productcode) = @PROD" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select productcode, description, cgno, scgno, productclass, isRandomWeight, TaxID into #P from purchase.dbo.product where productcode in (select * from #PRODS)" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select case when validfrom < @DATES then @DATES else validfrom end as validfrom , case when isnull(validto,getdate()) > @DATEE then @DATEE else isnull(validto,getdate()) end as validto" & Chr(10)
        CBA_strSQL = CBA_strSQL & ", Retail into #RETAIL from purchase.dbo.retail where productcode = @PROD and validfrom <= @DATEE and isnull(validto,getdate()) >= @DATES" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "Select datepart(ISOWK,pdate) as weekno, datepart(YY, pdate) as Yearno into #WEEKS from (SELECT  TOP (DATEDIFF(DAY, @DATES, @DATEE) + 1)" & Chr(10)
        CBA_strSQL = CBA_strSQL & "Pdate = convert(date,DATEADD(DAY, ROW_NUMBER() OVER(ORDER BY a.object_id) - 1, @DATES))     FROM    sys.all_objects a CROSS JOIN sys.all_objects b) a" & Chr(10)
        CBA_strSQL = CBA_strSQL & "group by datepart(ISOWK,pdate), datepart(YY, pdate) order by datepart(YY, pdate), datepart(ISOWK,pdate)" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "select wk.weekno, isnull(md.MDQTY,0) as MDQTY, isnull(md.MDRetail,0) as MDRetail, isnull(l.LossQTY,0) as LossQTY, isnull(l.LossRetail,0) as LossRetail, isnull(c.invdiff,0) as invdiff from #WEEKS wk" & Chr(10)
        CBA_strSQL = CBA_strSQL & "left join (select weekno, sum(MDRetail) as MDRetail, sum(MDQTY) as MDQTY from (select datepart(ISOWK, md.entrydate) as weekno, sum(md.totalretail) as MDRetail " & Chr(10)
        CBA_strSQL = CBA_strSQL & ", case when (select IsRandomWeight from #P where productcode = @PROD) = '01' then case when sum(md.totalpieces) = 0 then 0 else (abs(sum(md.totalretail)) / sum(md.totalpieces) / (sum(r.Retail)/count(r.retail))) * sum(md.totalpieces) end" & Chr(10)
        CBA_strSQL = CBA_strSQL & "else sum(md.totalpieces) end as MDQTY from Purchase.dbo.MemoDetail md left join #RETAIL r on r.validfrom <= md.EntryDate and r.validto>= md.EntryDate" & Chr(10)
        CBA_strSQL = CBA_strSQL & "where md.retail <> 0 and md.entrydate >= @DATES and md.entrydate <= @DATEE and VoucherTypeID <> 0 and md.ProductCode in (Select productcode from #PRODS) group by datepart(ISOWK, md.entrydate)" & Chr(10)
        CBA_strSQL = CBA_strSQL & ") a group by weekno ) md on md.weekno = wk.weekno left join (select datepart(ISOWK, md.entrydate) as weekno, sum(md.totalretail) as LossRetail " & Chr(10)
        CBA_strSQL = CBA_strSQL & ", case when (select IsRandomWeight from #P where productcode = @PROD) = '01' then case when sum(md.totalpieces) = 0 then 0 else (abs(sum(md.totalretail)) / sum(md.totalpieces) / (sum(r.Retail)/count(r.retail))) * sum(md.totalpieces) end" & Chr(10)
        CBA_strSQL = CBA_strSQL & "else sum(md.totalpieces) end as LossQTY from Purchase.dbo.MemoDetail md left join #RETAIL r on r.validfrom <= md.EntryDate and r.validto>= md.EntryDate where (md.retail = 0 or VoucherTypeID = 0) " & Chr(10)
        CBA_strSQL = CBA_strSQL & "and md.entrydate >= @DATES and md.entrydate <= @DATEE and md.ProductCode in (Select productcode from #PRODS) group by datepart(ISOWK, md.entrydate)) l on l.weekno = wk.weekno" & Chr(10)
        CBA_strSQL = CBA_strSQL & "left join (select datepart(ISOWK, InventoryDate) as weekno, sum(invDiff) as InvDiff from Purchase.dbo.StoreInvDiff where productcode in (Select productcode from #PRODS) " & Chr(10)
        CBA_strSQL = CBA_strSQL & "and InventoryDate >= @DATES and InventoryDate <= @DATEE and storeno <> 100 group by datepart(ISOWK, InventoryDate)) c on c.weekno = wk.weekno" & Chr(10)
        CBA_strSQL = CBA_strSQL & "" & Chr(10)
        CBA_strSQL = CBA_strSQL & "drop table #PRODS, #RETAIL, #P, #WEEKS" & Chr(10)
    
        Case "CBIS_IsTax"
            CBA_strSQL = CBA_strSQL & "select TaxID from cbis599p.dbo.product where productcode = " & CBA_LongVal1
        Case "CBIS_Prodinfo"
            CBA_strSQL = CBA_strSQL & "select description from cbis599p.dbo.product where productcode = " & CBA_LongVal1
        Case "ABI_INSERT"
            CBA_strSQL = CBA_StringVal1
        Case "CBIS_IsAlcohol"
            CBA_strSQL = "Select case when CGno < 5 then 'Yes' else 'No' end from cbis599p.dbo.product where productcode = " & CBA_LongVal1
        Case "ABI_AlcoholStores"
            CBA_strSQL = "select * from ABI_AlcoholStores where [DateQueried] >= " & "#" & Format(DateSerial(Year(CBA_datefrom), Month(CBA_datefrom), 1), "DD/MM/YYYY") & "# and [DateQueried] <= #" & Format(DateSerial(Year(CBA_Dateto), Month(CBA_Dateto) + 1, 0), "DD/MM/YYYY") & "#"
        Case "CBIS_WeekYear"
            CBA_strSQL = CBA_strSQL & "Select datepart(ISOWK,pdate) as weekno, datepart(YY, pdate) as Yearno from (SELECT  TOP (DATEDIFF(DAY, '" & Format(CBA_datefrom, "YYYY-MM-DD") & "', '" & Format(CBA_Dateto, "YYYY-MM-DD") & "') + 1)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "Pdate = convert(date,DATEADD(DAY, ROW_NUMBER() OVER(ORDER BY a.object_id) - 1, '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'))   FROM    sys.all_objects a CROSS JOIN sys.all_objects b) a" & Chr(10)
            CBA_strSQL = CBA_strSQL & "group by datepart(ISOWK,pdate), datepart(YY, pdate) order by datepart(YY, pdate), datepart(ISOWK,pdate)" & Chr(10)
    
        Case "CBIS_POSDATA"
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @PROD int = " & CBA_LongVal1 & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @SDATE date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @EDATE date = '" & Format(CBA_Dateto, "YYYY-MM-DD") & "'" & Chr(10)
            If (CBA_LongVal2 = 0 Or CBA_LongVal2 = 599) Then Else CBA_strSQL = CBA_strSQL & "declare @DIV int = " & CBA_LongVal2 & Chr(10)
            CBA_strSQL = CBA_strSQL & "select productcode into #PROD from cbis599p.dbo.product p" & Chr(10)
            If CBA_BooleanVal1 = False Then
                CBA_strSQL = CBA_strSQL & "where productcode = @PROD" & Chr(10)
            Else
                CBA_strSQL = CBA_strSQL & "where isnull(con_productcode, productcode) = @PROD" & Chr(10)
            End If
            CBA_strSQL = CBA_strSQL & "SELECT  TOP (DATEDIFF(DAY,  @SDATE,  @EDATE) + 1)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "Pdate = convert(date,DATEADD(DAY, ROW_NUMBER() OVER(ORDER BY a.object_id) - 1,  @SDATE))" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #DATE FROM sys.all_objects a CROSS JOIN sys.all_objects b" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select  convert(date,posdate) as posdate,sum(quantity) as QTY, sum(retail) as Retail" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #POS from cbis599p.dbo.pos pos inner join #PROD p  on p.productcode = pos.ProductCode --left join cbis599p.dbo.division d on d.divno = pos.DivNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where pos.PosDate >= @SDATE and pos.PosDate <= @EDATE " & Chr(10)
            If (CBA_LongVal2 = 0 Or CBA_LongVal2 = 599) Then Else CBA_strSQL = CBA_strSQL & "and pos.divno = @DIV " & Chr(10)
            CBA_strSQL = CBA_strSQL & "group by posdate Select d.pdate, isnull(pos.QTY,0) as QTY , isnull(pos.Retail,0) as Retail" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from #DATE d left  join #POS pos on pos.posdate = d.Pdate" & Chr(10)
            CBA_strSQL = CBA_strSQL & "drop table #DATE, #POS, #PROD" & Chr(10)

    Case "CBIS_Specials"
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @PROD int = " & CBA_LongVal1 & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @SDATE date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @EDATE date = '" & Format(CBA_Dateto, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @CG int = (select cgno from cbis599p.dbo.product where productcode = @Prod)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "SELECT DOS.PRODUCTCODE,DOS.OSD,DOS.DESCR,dos.BD_name as BD_Name,SUM(DOS.SALES) AS 'Retail'," & Chr(10)
            CBA_strSQL = CBA_strSQL & "    sum (dos.cost) as Cost, sum(dos.netretail) as NetRetail," & Chr(10)
            CBA_strSQL = CBA_strSQL & "    sum(dos.netretail-dos.cost) as Contr, (dos.netretail-dos.cost)/nullif(dos.sales2,0) as Margin," & Chr(10)
            CBA_strSQL = CBA_strSQL & "    dos.theme, dos.retail1 as Price,dos.bd as BD, dos.cgno, dos.scgno FROM (" & Chr(10)
            CBA_strSQL = CBA_strSQL & "SELECT em.empsign as BD, em.name as BD_Name,sptf.cgno as cgno, sptf.scgno as scgno, SRVR.GroupAdvdate AS OSD," & Chr(10)
            CBA_strSQL = CBA_strSQL & "    SRVR.PRODUCTCODE,ppl.description AS DESCR, ROW_NUMBER() OVER (partition by srvr.groupadvdate,srvr.portfolioid, srvr.productcode," & Chr(10)
            CBA_strSQL = CBA_strSQL & "    srvr.quantity1,srvr.retail1,ppl.description order by srvr.groupadvdate,srvr.portfolioid," & Chr(10)
            CBA_strSQL = CBA_strSQL & "    srvr.productcode, srvr.quantity1,srvr.retail1,ppl.description) AS rowNumber," & Chr(10)
            CBA_strSQL = CBA_strSQL & "    (SRVR.QUANTITY1 * SRVR.RETAIL1) AS 'SALES',(case when srvr.cost1>0 then SRVR.QUANTITY1 * SRVR.retail1 else 0 end) AS 'Sales2'," & Chr(10)
            CBA_strSQL = CBA_strSQL & "    (case when srvr.cost1>0 then SRVR.QUANTITY1 * SRVR.COST1 else 0 end) AS 'COST'," & Chr(10)
            CBA_strSQL = CBA_strSQL & "    (case when srvr.cost1>0 then SRVR.QUANTITY1 * SRVR.netRETAIL1 else 0 end) AS 'NetRetail'," & Chr(10)
            CBA_strSQL = CBA_strSQL & "    sscn.description as theme, srvr.retail1 as Retail1" & Chr(10)
            CBA_strSQL = CBA_strSQL & "FROM [cbis599p].[portfolio].[Rep_PfVersionReg] as SRVR" & Chr(10)
            CBA_strSQL = CBA_strSQL & "LEFT JOIN [CBIS599p].[portfolio].[Portfolio] as SPTF ON SRVR.PortfolioID = SPTF.PortfolioID" & Chr(10)
            CBA_strSQL = CBA_strSQL & "LEFT JOIN [cbis599p].[portfolio].[Pfversion] as PFV ON SRVR.PortfolioID = PFV.PortfolioID AND SRVR.PFVERSIONID = PFV.PFVERSIONID" & Chr(10)
            CBA_strSQL = CBA_strSQL & "LEFT JOIN [cbis599p].[portfolio].[PfVersionMapping] AS PFVM ON SRVR.PortfolioID = PFVM.PortfolioID AND SRVR.PFVERSIONID = PFVM.PFVERSIONID" & Chr(10)
            CBA_strSQL = CBA_strSQL & "LEFT JOIN [cbis599p].[portfolio].[PfVersionReg] AS PFVR ON SRVR.PortfolioID = PFVR.PortfolioID AND SRVR.PFVERSIONID = PFVR.PFVERSIONID" & Chr(10)
            CBA_strSQL = CBA_strSQL & "LEFT JOIN [CBIS599p].[Dbo].[SPECIALCATEGORY] as SSCN ON      PFVR.SpecialCatergoryno = SSCN.SpecialCatergoryNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "LEFT JOIN [cbis599p].[portfolio].[pfversionlng] AS PPL ON " & Chr(10)
            CBA_strSQL = CBA_strSQL & "     SRVR.PortfolioID = ppl.PortfolioID  AND SRVR.PFVERSIONID = ppl.PFVERSIONID and ppl.LANGUAGEID=0" & Chr(10)
            CBA_strSQL = CBA_strSQL & "LEFT JOIN [cbis599p].[dbo].[employee] AS EM ON SPTF.empno = Em.empno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "WHERE pfvr.advertisingdate >= @SDATE AND pfvr.advertisingdate <= @EDATE AND PFV.PFSTATUSID IS NOT NULL" & Chr(10)
            CBA_strSQL = CBA_strSQL & "AND PFVR.PRODUCTCLASS in (2,3)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "AND SPTF.CGNo in (@CG)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "GROUP BY em.empsign, em.name, sptf.cgno, sptf.scgno," & Chr(10)
            CBA_strSQL = CBA_strSQL & "     SRVR.GroupAdvdate, SSCN.dESCRIPTION, PFVM.PRODUCTCODE, ppl.description, srvr.portfolioid, SRVR.Productcode ," & Chr(10)
            CBA_strSQL = CBA_strSQL & "     SRVR.QUANTITY1 , SRVR.RETAIL1, srvr.cost1, srvr.netretail1, sscn.description, srvr.retail1 ) AS DOS" & Chr(10)
            CBA_strSQL = CBA_strSQL & "WHERE Rownumber = 1" & Chr(10)
            CBA_strSQL = CBA_strSQL & "GROUP BY dos.bd,  dos.BD_name, dos.cgno, dos.scgno, DOS.OSD, DOS.PRODUCTCODE," & Chr(10)
            CBA_strSQL = CBA_strSQL & "     DOS.DESCR, dos.theme, (dos.netretail-dos.cost)/nullif(dos.sales2,0), dos.retail1"
                    
    Case "CBIS_Retail and Cost"
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @PROD int = " & CBA_LongVal1 & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @DIV int = 501" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @SDATE date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select productcode, validfrom, retail  into #RET from (" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select productcode, validfrom, Retail , row_number() over (Partition by productcode order by validfrom desc) as row" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from cbis599p.dbo.divretail" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where productcode = @PROD and divno = @DIV  and AuthorizedDate is not null  and validfrom <= @SDATE" & Chr(10)
            CBA_strSQL = CBA_strSQL & ") a where a.row = 1" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select distinct productcode, contractno into #CONTS from cbis599p.dbo.contract" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where productcode = @PROD and deliveryfrom <= @SDATE and isnull(deliveryto, convert(date,@SDATE)) >= convert(date,@SDATE)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select contractno, validfrom ,case when (isnull(FOB,0) + isnull(ExWorks,0)) = 0 then DDP" & Chr(10)
            CBA_strSQL = CBA_strSQL & "         when isnull(FOB,0) > 0 then FOB + isnull(Freight,0)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "         when isnull(Exworks,0) > 0 then ExWorks + isnull(Freight,0)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "end as cost  into #COST from (" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select bv.ContractNo, bv.validfrom , isnull(validto, convert(date,@SDATE)) as validto" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",ROW_NUMBER() over (PArtition by bv.contractno order by validfrom desc) as row" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", bv.DDP, bv.FOB, bv.ExWorks, bv.Freight, bv.FreightAUD" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from cbis599p.dbo.BRAKETVALUE bv inner join #CONTS cont on cont.ContractNo = bv.ContractNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where DivNo = @DIV  and validfrom <= convert(date,@SDATE) and isnull(validto, convert(date,@SDATE)) >= convert(date,@SDATE)" & Chr(10)
            CBA_strSQL = CBA_strSQL & ") a where a.row = 1" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select r.productcode, r.Retail, c.cost/ p.packsize as cost from #RET r" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join #CONTS co on co.ProductCode = r.ProductCode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "inner join #COST c on c.ContractNo = co.ContractNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.product p on p.productcode = r.ProductCode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "drop table #RET, #CONTS, #COST" & Chr(10)
                     
    Case "CBIS_Themes"
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @SDATE date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @EDATE date = '" & Format(CBA_Dateto, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "SELECT DISTINCT sc.Description,min(pfvr.Advertisingdate) " & Chr(10)
            CBA_strSQL = CBA_strSQL & "FROM cbis599p.Portfolio.PfVersionReg PFVR  " & Chr(10)
            CBA_strSQL = CBA_strSQL & "LEFT JOIN cbis599p.dbo.SPECIALCATEGORY SC ON SC.SpecialCatergoryno = pfvr.SpecialCatergoryno " & Chr(10)
            CBA_strSQL = CBA_strSQL & "WHERE pfvr.advertisingdate >= @SDATE AND pfvr.Advertisingdate <= @EDATE AND sc.Description IS NOT NULL " & Chr(10)
            CBA_strSQL = CBA_strSQL & "GROUP BY sc.Description  " & Chr(10)
            CBA_strSQL = CBA_strSQL & "ORDER BY min(pfvr.Advertisingdate)  " & Chr(10)
            
            
        Case "MMS_POSDATA"
            CBA_DBtoQuery = CBA_LongVal2
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @PROD int = " & CBA_LongVal1 & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @SDATE date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @EDATE date = '" & Format(CBA_Dateto, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select productcode INTO #P from purchase.dbo.product" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where productcode in (@PROD)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select ss.salesdate, ss.storeno, ss.productcode, isnull(ss.retail,0) as retail, isnull(ss.quantity,0) as quantity" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #SALES" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from purchase.dbo.storesales ss" & Chr(10)
            CBA_strSQL = CBA_strSQL & "inner join #P as p on p.ProductCode = ss.ProductCode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where SalesDate >= @SDATE" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and salesdate <= @EDATE" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select * into #PDATA from (" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select datepart(iso_week,salesdate) as IsoWk, YEAR(DATEADD(day, 26 - DATEPART(ISOWW, salesdate), salesdate)) as IsoYr" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",productcode, count(distinct storeno) as storecount, sum(quantity) as QTY, sum(Retail) as Ret, count(distinct salesdate) as daycount" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from #SALES" & Chr(10)
            CBA_strSQL = CBA_strSQL & "group by datepart(iso_week,salesdate) , YEAR(DATEADD(day, 26 - DATEPART(ISOWW, salesdate), salesdate)), productcode" & Chr(10)
            CBA_strSQL = CBA_strSQL & ") a" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where a.daycount > 2" & Chr(10)
            CBA_strSQL = CBA_strSQL & "order by productcode,IsoYr, IsoWk" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select productcode,  avg(QTY / nullif(storecount,0)) as USW, avg(Ret / nullif(storecount,0)) as RSW, count(productcode) as wkCnt" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from #PDATA" & Chr(10)
            CBA_strSQL = CBA_strSQL & "group by productcode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "drop table #SALES, #PDATA, #P" & Chr(10)
            
            ' Brings back ... productcode  UnitsPerStorePerWeek   RetailPerStorePerWeek        WeekCount
            '                     1035          425.222602          1100.77071917808             4
            
        
        Case "MMS_StoreSalesDATA"
            CBA_DBtoQuery = CBA_LongVal2
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @PROD int = " & CBA_LongVal1 & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @SDATE date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @EDATE date = '" & Format(CBA_Dateto, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select productcode into #PROD from purchase.dbo.product p" & Chr(10)
            If CBA_BooleanVal1 = False Then
                CBA_strSQL = CBA_strSQL & "where productcode = @PROD" & Chr(10)
            Else
                CBA_strSQL = CBA_strSQL & "where coalesce(mainproduct,con_productcode, productcode) = @PROD" & Chr(10)
            End If
            CBA_strSQL = CBA_strSQL & "SELECT  TOP (DATEDIFF(DAY,  @SDATE,  @EDATE) + 1)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "Pdate = convert(date,DATEADD(DAY, ROW_NUMBER() OVER(ORDER BY a.object_id) - 1,  @SDATE))" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #DATE FROM sys.all_objects a CROSS JOIN sys.all_objects b" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select  convert(date,salesdate) as posdate, storeno, sum(quantity) as QTY, sum(retail) as Retail" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #POS from purchase.dbo.storesales pos inner join #PROD p  on p.productcode = pos.ProductCode --left join cbis599p.dbo.division d on d.divno = pos.DivNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where pos.salesDate >= @SDATE and pos.SalesDate <= @EDATE" & Chr(10)
            CBA_strSQL = CBA_strSQL & "group by salesdate , storeno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "Select substring(@@Servername,2,3) as Div, d.pdate, s.City, isnull(pos.QTY,0) as QTY , isnull(pos.Retail,0) as Retail" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from #DATE d" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left  join #POS pos on pos.posdate = d.Pdate" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join purchase.dbo.Store s on s.StoreNo = pos.StoreNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "order by pos.storeno, d.pdate" & Chr(10)
            CBA_strSQL = CBA_strSQL & "drop table #DATE, #POS, #PROD" & Chr(10)
    
        Case "CBIS_INVDATA"
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @PROD int = " & CBA_LongVal1 & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @SDATE date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @EDATE date = '" & Format(CBA_Dateto, "YYYY-MM-DD") & "'" & Chr(10)
            If (CBA_LongVal2 = 0 Or CBA_LongVal2 = 599) Then Else CBA_strSQL = CBA_strSQL & "declare @DIV int = " & CBA_LongVal2 & Chr(10)
            CBA_strSQL = CBA_strSQL & "select productcode into #PROD from cbis599p.dbo.product p" & Chr(10)
            If CBA_BooleanVal1 = False Then
                CBA_strSQL = CBA_strSQL & "where productcode = @PROD" & Chr(10)
            Else
                CBA_strSQL = CBA_strSQL & "where isnull(con_productcode, productcode) = @PROD" & Chr(10)
            End If
            CBA_strSQL = CBA_strSQL & "SELECT  TOP (DATEDIFF(DAY,  @SDATE,  @EDATE) + 1)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "Pdate = convert(date,DATEADD(DAY, ROW_NUMBER() OVER(ORDER BY a.object_id) - 1,  @SDATE))" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #DATE FROM sys.all_objects a CROSS JOIN sys.all_objects b" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select divno,sh.StockReferenceDate, sum(StoreQuantity) as StoreQTY, sum(Quantity)  as WHQTY" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from cbis599p.dbo.STOCKhis sh inner join #DATE d on d.pdate = sh.StockReferenceDate" & Chr(10)
            CBA_strSQL = CBA_strSQL & "inner join #PROD p on p.ProductCode = sh.ProductCode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "group by sh.StockReferenceDate, sh.divno order by sh.StockReferenceDate, sh.divno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "drop table #PROD, #DATE" & Chr(10)
    
        Case "CBIS_UnrealisedRevenue"
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @Prod int = " & CBA_LongVal1 & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @SDate date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @CG int = (select cgno from cbis599p.dbo.product where productcode = @Prod)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @POSBeforeOSD int declare @RCVBeforeOSD int declare @RCVAfterOSD int" & Chr(10)
            CBA_strSQL = CBA_strSQL & "if @CG < 5  set @POSBeforeOSD = 3 else if @CG = 64 or @CG = 62 set @POSBeforeOSD = 0 else set @POSBeforeOSD = 1" & Chr(10)
            CBA_strSQL = CBA_strSQL & "if @CG < 5  set @RCVBeforeOSD = 14 else if @CG = 64 or @CG = 62 set @RCVBeforeOSD = 3 else set @RCVBeforeOSD = 14" & Chr(10)
            CBA_strSQL = CBA_strSQL & "if @CG < 5  set @RCVAfterOSD = 3 else if @CG = 64 or @CG = 62 set @RCVAfterOSD = 0 else set @RCVAfterOSD = 0" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @3Date date = dateadd(D,2,@Sdate) declare @7Date date = dateadd(D,6,@Sdate)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @14Date date = dateadd(D,13,@Sdate) declare @30Date date = dateadd(D,29,@Sdate) declare @60Date date = dateadd(D,59,@Sdate)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @90Date date = dateadd(D,89,@Sdate) declare @EDate date = dateadd(D,179,@Sdate)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select convert(int,divno) as divno into #DIV  from cbis599p.dbo.division where divno > '500' and divno < '510'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select divno, OSD, case when row = 1 and OSD = @Sdate then 'Start' else 'End' end as OSDType into #OSDDATA from (" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select c.contractno, dch.divno, c.advertisingdate as OSD , row_number() over (Partition by dch.divno order by c.advertisingdate) as row" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from cbis599p.dbo.contract c left join cbis599p.dbo.DIVCONTRACTHIS dch on dch.ContractNo = c.contractno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where c.ProductCode = @Prod and c.advertisingdate >= dateadd(D,-3,@Sdate) and dch.SendToDiv = 1 ) a where a.row in (1,2) order by divno, OSD" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select d.divno,isnull(o.OSD,@EDate) as OSD into #OSDSTART from #DIV d left join #OSDDATA o on o.DivNo = d.divno where OSDType = 'Start'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select d.divno, isnull(o.OSD,@EDate) as OSD into #OSDEND  from #DIV d left join #OSDDATA o on o.DivNo = d.divno and OSDType = 'End'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select divno, sum(Pieces) as rcvQTY into #RCVQty from cbis599p.dbo.receiving rcv left join cbis599p.dbo.PRODUCT p on p.productcode = rcv.ProductCode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where dayenddate >= dateadd(D,-@RCVBeforeOSD,@SDate) and dayenddate <=dateadd(D,@RCVAfterOSD,@Sdate) and isnull(p.con_ProductCode, rcv.productcode) = @prod" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and RecordID = '001' group by DivNo select divno, sum(retail) as rcvRet, sum(cost) as rcvCost, sum(retailnet) as rcvRetNet into #RCVRet from cbis599p.dbo.receiving rcv" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.PRODUCT p on p.productcode = rcv.ProductCode where dayenddate >= dateadd(D,-@RCVBeforeOSD,@SDate)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and dayenddate <= dateadd(D,@RCVAfterOSD,@Sdate) and isnull(p.con_ProductCode, rcv.productcode) = @prod group by DivNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select pos.posdate, pos.divno, case when PosDate <= @3Date then sum(retail) end as pos3Ret, case when PosDate <= @7Date then sum(retail) end as pos7Ret" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when PosDate <= @14Date then sum(retail) end as pos14Ret, case when PosDate <= @30Date then sum(retail) end as pos30Ret" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when PosDate <= @60Date then sum(retail) end as pos60Ret, case when PosDate <= @90Date then sum(retail) end as pos90Ret" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when PosDate <= @EDate then sum(retail) end as pos180Ret" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when PosDate <= @3Date then sum(Quantity) end as pos3QTY" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when PosDate <= @7Date then sum(Quantity) end as pos7QTY" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when PosDate <= @14Date then sum(Quantity) end as pos14QTY" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when PosDate <= @30Date then sum(Quantity) end as pos30QTY, case when PosDate <= @60Date then sum(Quantity) end as pos60QTY" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when PosDate <= @90Date then sum(Quantity) end as pos90QTY, case when PosDate <= @EDate then sum(Quantity) end as pos180QTY" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #POSSTART from cbis599p.dbo.POS left join cbis599p.dbo.PRODUCT p on p.productcode = POS.ProductCode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "inner join #OSDSTART os on os.DivNo = pos.divno and os.OSD <= pos.PosDate inner join #OSDEND oe on oe.DivNo = pos.divno and oe.OSD >= pos.PosDate" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where PosDate >= dateadd(D,-@POSBeforeOSD,@SDate) and PosDate <= @EDate and isnull(p.con_ProductCode, pos.productcode) = @prod group by posdate, pos.DivNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select divno, sum(pos3Ret) as pos3Ret, sum(pos7Ret) as pos7Ret, sum(pos14Ret) as pos14Ret, sum(pos30Ret) as pos30Ret, sum(pos60Ret) as pos60Ret, sum(pos90Ret) as pos90Ret, sum(pos180Ret) as pos180Ret" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", sum(pos3QTY) as pos3QTY, sum(pos7QTY) as pos7QTY, sum(pos14QTY) as pos14QTY, sum(pos30QTY) as pos30QTY, sum(pos60QTY) as pos60QTY, sum(pos90QTY) as pos90QTY, sum(pos180QTY) as pos180QTY into #POS From #POSSTART group by divno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select rcvqty.divno as QTYDiv, rcvqty.rcvQTY" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when  pos3QTY > rcvqty.rcvQty then rcvqty.rcvQty else pos3Qty end as pos3Qty" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when  pos7QTY > rcvqty.rcvQty then rcvqty.rcvQty else pos7Qty end as pos7Qty" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when  pos14QTY > rcvqty.rcvQty then rcvqty.rcvQty else pos14Qty end as pos14Qty" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when pos30QTY > rcvqty.rcvQty then rcvqty.rcvQty else pos30Qty end as pos30Qty, case when pos60QTY > rcvqty.rcvQty then rcvqty.rcvQty else pos60Qty end as pos60Qty" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when pos90QTY > rcvqty.rcvQty then rcvqty.rcvQty else pos90Qty end as pos90Qty, case when pos180QTY > rcvqty.rcvQty then rcvqty.rcvQty else pos180Qty end as pos180Qty" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",rcvqty.divno as RetDiv,  rcvret.rcvRet, RCVRet.rcvRetNet,rcvret.rcvCost" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when pos3Ret > rcvret.rcvRet then rcvret.rcvRet else pos3Ret end as pos3Ret" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when pos7Ret > rcvret.rcvRet then rcvret.rcvRet else pos7Ret end as pos7Ret" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when pos14Ret > rcvret.rcvRet then rcvret.rcvRet else pos14Ret end as pos14Ret" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when pos30Ret > rcvret.rcvRet then rcvret.rcvRet else pos30Ret end as pos30Ret, case when pos60Ret > rcvret.rcvRet then rcvret.rcvRet else pos60Ret end as pos60Ret" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when pos90Ret > rcvret.rcvRet then rcvret.rcvRet else pos90Ret end as pos90Ret, case when pos180Ret > rcvret.rcvRet then rcvret.rcvRet else pos180Ret end as pos180Ret" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #DATA from #RCVQTY rcvqty  left join #RCVRet rcvret on rcvret.divno = rcvqty.divno left join #POS pos on pos.divno = rcvqty.divno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select QTYdiv,rcvQTY ,pos3Qty, pos3Qty  / rcvQTY as ST3,pos7Qty, pos7Qty  / rcvQTY as ST7,pos14Qty, pos14Qty  / rcvQTY as ST14,pos30Qty , pos30Qty  / rcvQTY as ST30,pos60Qty" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",pos60Qty  / rcvQTY as ST60,pos90Qty, pos90Qty  / rcvQTY as ST90,pos180Qty,pos180Qty  / rcvQTY as ST180" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",RetDiv,rcvRet --, rcvRetNet,  rcvCost" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",rcvRet - pos3Ret as URR3, pos3Ret - rcvCost as Cont3 ,rcvRet - pos7Ret as URR7, pos7Ret - rcvCost as Cont7,rcvRet - pos14Ret as URR14, pos14Ret - rcvCost as Cont14" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",rcvRet - pos30Ret as URR30, pos30Ret - rcvCost as Cont30,rcvRet - pos60Ret as URR60, pos60Ret - rcvCost as Cont60 ,rcvRet - pos90Ret as URR90, pos90Ret - rcvCost as Cont90" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",rcvRet - pos180Ret as URR180, pos90Ret - rcvCost as Cont180 ,(rcvRetNet - rcvCost)/rcvRet as GrossMargin" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when rcvRetNet = rcvRet then (pos180Ret - rcvCost) / pos180Ret else ((pos180Ret /1.1) - rcvCost) / pos180Ret end as NetMargin" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from #DATA" & Chr(10)
            CBA_strSQL = CBA_strSQL & "drop table #DIV, #OSDDATA, #OSDSTART, #OSDEND, #RCVqty, #RCVret, #POSSTART, #POS, #DATA" & Chr(10)
        Case "CBIS_CG"
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select cgno, scgno from cbis599p.dbo.product where productcode = " & CBA_LongVal1 & Chr(10)
        Case "CBIS_CGDesc"
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select scgno from cbis599p.dbo.product where productcode = " & CBA_LongVal1 & Chr(10)
        Case "CBIS_SCG"
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select cgno from cbis599p.dbo.product where productcode = " & CBA_LongVal1 & Chr(10)
        Case "CBIS_SCGDesc"
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select description from cbis599p.dbo.subcommoditygroup where CGno = CBA_LongVal1 and SCGno = " & CBA_LongVal2 & Chr(10)
        Case "CBIS_ProductDesc"
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select description from cbis599p.dbo.product where productcode = " & CBA_LongVal1 & Chr(10)
        Case "CBIS_CGSCGList"
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select isnull(cg.cgno,0) + ' - ' + isnull(cg.description,'') as CG, isnull(scg.scgno,1) + ' - ' + isnull(scg.description,cg.description)  as SCG" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from cbis599p.dbo.COMMODITYGROUP cg left join cbis599p.dbo.SUBCOMMODITYGROUP scg on scg.cgno = cg.cgno" & Chr(10)
        Case "CBIS_ForecastQuery"
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @CG int = " & CBA_LongVal1 & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @SCG int = " & CBA_LongVal2 & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @FMONTH int = " & CBA_LongVal3 & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @FYEAR int = " & CBA_LongVal4 & Chr(10)
            CBA_strSQL = CBA_strSQL & "select productclass into #PC from cbis599p.dbo.productclass" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select p.productclass, sum(rcv.retail) as PYRetail  , case when sum(rcv.retailnet) = 0 then 0 else (sum(rcv.retail) - sum(rcv.cost)) / sum(rcv.retailnet) end as PYMargin into #PY" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from cbis599p.dbo.Receiving rcv inner join cbis599p.dbo.product p on p.productcode = rcv.productcode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where p.cgno = @CG and p.scgno = @SCG and month(rcv.DayEnddate) = @FMONTH" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and year(rcv.DayEnddate) = @FYEAR -1 group by p.ProductClass" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select p.productclass, sum(rcv.retail) as CYRetail, (sum(rcv.retail) - sum(rcv.cost)) / sum(rcv.retailnet) as CYMargin into #CY" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from cbis599p.dbo.Receiving rcv inner join cbis599p.dbo.product p on p.productcode = rcv.productcode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where p.cgno = @CG and p.scgno = @SCG and month(rcv.DayEnddate) = @FMONTH" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and year(rcv.DayEnddate) = @FYEAR group by p.ProductClass" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select PC.productclass, isnull(PY.PYRetail,0) ,  isnull(CY.CYRetail,0), isnull(PY.PYMargin,0), isnull(CY.CYMargin,0)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from #PC pc left join #PY py on py.ProductClass = pc.ProductClass" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join #CY cy on cy.ProductClass = pc.ProductClass" & Chr(10)
            CBA_strSQL = CBA_strSQL & "drop table #PC, #PY, #CY" & Chr(10)
        Case "CBForecast_Apply"
            CBA_DBtoQuery = 8
            CBA_strSQL = "INSERT INTO Data ([Productclass], [FMonth], [FYear], [CG], [SCG], [FSales],[FReSales], [FMarginP], [FREMarginP], [SubmitDate], [UserName])" & Chr(10)
            CBA_strSQL = CBA_strSQL & "SELECT " & CBA_LongVal1 & " ," & CBA_LongVal2 & ", " & CBA_LongVal3 & ", " & CBA_LongVal4 & ", " & CBA_VariantVal1 & ", " & CBA_DoubleVal1 & ", " & CBA_DoubleVal2 & ", " & CBA_DoubleVal3 & ", " & CBA_DoubleVal4 & " , '" & Now & "' , '" & Application.UserName & "'"
        Case "CBForecast_SalesOrigQuery"
            CBA_DBtoQuery = 7
            CBA_strSQL = "SELECT top 1 t1.FSales FROM Data as t1" & Chr(10)
            CBA_strSQL = CBA_strSQL & "WHERE (((t1.ProductClass)=" & CBA_LongVal1 & ") AND ((t1.FMonth)=" & CBA_LongVal2 & ") AND ((t1.FYear)=" & CBA_LongVal3 & ") AND ((t1.CG)=" & CBA_LongVal4 & ") AND ((t1.SCG)=" & CBA_VariantVal1 & "))" & Chr(10)
            CBA_strSQL = CBA_strSQL & "order by t1.submitdate desc" & Chr(10)
        Case "CBForecast_SalesReforecastQuery"
            CBA_DBtoQuery = 7
            CBA_strSQL = "SELECT top 1 t1.FReSales FROM Data as t1" & Chr(10)
            CBA_strSQL = CBA_strSQL & "WHERE (((t1.ProductClass)=" & CBA_LongVal1 & ") AND ((t1.FMonth)=" & CBA_LongVal2 & ") AND ((t1.FYear)=" & CBA_LongVal3 & ") AND ((t1.CG)=" & CBA_LongVal4 & ") AND ((t1.SCG)=" & CBA_VariantVal1 & "))" & Chr(10)
            CBA_strSQL = CBA_strSQL & "order by t1.submitdate desc" & Chr(10)
        Case "CBForecast_MarginOrigQuery"
            CBA_DBtoQuery = 7
            CBA_strSQL = "SELECT top 1 t1.FMarginP FROM Data as t1" & Chr(10)
            CBA_strSQL = CBA_strSQL & "WHERE (((t1.ProductClass)=" & CBA_LongVal1 & ") AND ((t1.FMonth)=" & CBA_LongVal2 & ") AND ((t1.FYear)=" & CBA_LongVal3 & ") AND ((t1.CG)=" & CBA_LongVal4 & ") AND ((t1.SCG)=" & CBA_VariantVal1 & "))" & Chr(10)
            CBA_strSQL = CBA_strSQL & "order by t1.submitdate desc" & Chr(10)
        Case "CBForecast_MarginReforecastQuery"
            CBA_DBtoQuery = 7
            CBA_strSQL = "SELECT top 1 t1.FReMarginP FROM Data as t1" & Chr(10)
            CBA_strSQL = CBA_strSQL & "WHERE (((t1.ProductClass)=" & CBA_LongVal1 & ") AND ((t1.FMonth)=" & CBA_LongVal2 & ") AND ((t1.FYear)=" & CBA_LongVal3 & ") AND ((t1.CG)=" & CBA_LongVal4 & ") AND ((t1.SCG)=" & CBA_VariantVal1 & "))" & Chr(10)
            CBA_strSQL = CBA_strSQL & "order by t1.submitdate desc" & Chr(10)
        Case "CBForecast_RepFMarginP", "CBForecast_RepFReMarginP", "CBForecast_RepFReSales", "CBForecast_RepFSales"
            CBA_DBtoQuery = 7
            CBA_strSQL = "SELECT Data.CG, Data.SCG, Data.FMonth, Data.ProductClass, Data.FYear, Last(Data." & Mid(CBA_SQLType, 15, 99) & ") AS " & Mid(CBA_SQLType, 15, 99) & Chr(10)
            CBA_strSQL = CBA_strSQL & "FROM Data" & Chr(10)
            CBA_strSQL = CBA_strSQL & "WHERE (((Data." & Mid(CBA_SQLType, 15, 99) & ") > 0))" & Chr(10)
            CBA_strSQL = CBA_strSQL & "GROUP BY Data.CG, Data.SCG, Data.FMonth, Data.ProductClass, Data.FYear;" & Chr(10)
        Case "CBForecast_CutOff"
            CBA_DBtoQuery = 7
            CBA_strSQL = "SELECT Dayno, MonthNo, YearNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "FROM CutOffDates" & Chr(10)
            CBA_strSQL = CBA_strSQL & "Where YearofForecast = " & CBA_LongVal1 & " And MonthOfForecast = " & CBA_LongVal2 & Chr(10)
        Case "DBU_Map"
            CBA_DBtoQuery = 1
            CBA_strSQL = "select * from DBU_Map order by DBU, CG, SCG, PClass"
        Case "DBU_Map_ByDBU"
            CBA_DBtoQuery = 1
            CBA_strSQL = "select * from DBU_Map where DBU = '" & CBA_StringVal1 & "'  order by DBU, CG, SCG, PClass"
        Case "DBU_List"
            CBA_DBtoQuery = 1
            CBA_strSQL = "select * from DBUList"
        Case "DBU_Map_CGCheck"
            CBA_DBtoQuery = 1
            CBA_strSQL = "select count(DBU) from (SELECT distinct DBU_Map.DBU FROM DBU_Map where DBU_Map.CG = " & CBA_LongVal1 & ") a;"
        Case "CBA_Forecast_TotalMonth"
            CBA_DBtoQuery = 7
            CBA_strSQL = "select sum(FRetailO), sum(FRetailN) from DATAQ where MonthNo = " & CBA_LongVal1 & " and YearNo = " & CBA_LongVal2
        Case "CBA_Forecast_TotalBus"
            CBA_DBtoQuery = 7
            CBA_strSQL = "select MonthNo, YearNo, sum(FRetailO), sum(iif(FRetailN=0,FRetailO,FRetailN)) from DATAQ where (MonthNo >= " & CBA_LongVal1 & " and YearNo >= " & CBA_LongVal2 & ") and (MonthNo <= " & CBA_LongVal3 & " and YearNo <= " & CBA_LongVal4 & ") group by MonthNo, Yearno"
        Case "CBA_CDS_AuditbyCG"
            CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_RetrieveCGsForAudit"
            CBA_DBtoQuery = 599
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select distinct c.contractno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #CL from cbis599p.dbo.CONTRACT c" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.product p  on p.ProductCode = c.productcode" & Chr(10)
            strCGSQL = ""
            If UBound(CBA_CDSarr, 2) = 0 And CBA_CDSarr(0, 0) = 0 Then
                strCGSQL = strCGSQL & "where ((p.cgno = 1 and p.scgno in (1,3,4))" & Chr(10)
                strCGSQL = strCGSQL & "or (p.cgno = 4 and p.scgno in (1,2,3))" & Chr(10)
                strCGSQL = strCGSQL & "or (p.cgno = 5 and p.scgno in (1,2,3,4,5,6,7,8))" & Chr(10)
                strCGSQL = strCGSQL & "or (p.cgno = 50 and p.scgno in (2,3))" & Chr(10)
                strCGSQL = strCGSQL & "or (p.cgno = 51 and p.scgno in (1,4,5,36))" & Chr(10)
                strCGSQL = strCGSQL & "or (p.cgno = 54 and p.scgno in (6)))" & Chr(10)
            Else
                bytCurCG = 0
                For a = LBound(CBA_CDSarr, 2) To UBound(CBA_CDSarr, 2)
                    If bytCurCG <> CBA_CDSarr(0, a) Then
                        If strCGSQL <> "" Then strCGSQL = strCGSQL & "))" & Chr(10)
                        bytCurCG = CBA_CDSarr(0, a)
                        If strCGSQL = "" Then
                            strCGSQL = strCGSQL & "where ((p.cgno = " & bytCurCG & " and p.scgno in ("
                        Else
                            strCGSQL = strCGSQL & "or (p.cgno = " & bytCurCG & " and p.scgno in ("
                        End If
                    Else
                        strCGSQL = strCGSQL & ","
                    End If
                    If a = UBound(CBA_CDSarr, 2) Then
                        strCGSQL = strCGSQL & CBA_CDSarr(1, a) & ")))"
                    Else
                        strCGSQL = strCGSQL & CBA_CDSarr(1, a)
                    End If
                Next
            End If
            CBA_strSQL = CBA_strSQL & strCGSQL
            CBA_strSQL = CBA_strSQL & "and isnull(c.deliveryto,getdate()) >= '2017-01-01'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and c.deliveryfrom <= getdate()" & Chr(10)

            CBA_strSQL = CBA_strSQL & "select distinct Contractno, ps.SpecificationID as SID, ps.Description as SchemeID, c.OptionID, ol.Description as type, count(contractno) over (partition by contractno)  as cnt" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #PS from cbis599p.dbo.contractproductspecification  c" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.GeDa.ProductSpecificationLng ps on ps.SpecificationID = c.SpecificationID and ps.LanguageID = 0" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.GeDa.ProductSpecificationOptionLng ol on ol.SpecificationID = c.SpecificationID and ol.OptionID = c.OptionID and ol.LanguageID = 0" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where ContractNo in (select contractno from #CL)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and ps.SpecificationID in (25,26,13,45)" & Chr(10)

            CBA_strSQL = CBA_strSQL & "select distinct contractno into #CR from #PS a where cnt = 4" & Chr(10)

            CBA_strSQL = CBA_strSQL & "select distinct cl.contractno--, cr.ContractNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #CO from #CL cl" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join #CR cr on cr.ContractNo = cl.ContractNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where cr.ContractNo Is Null" & Chr(10)

            CBA_strSQL = CBA_strSQL & "select c.contractno, c.productcode, p.description, p.cgno, cg.description, p.scgno, scg.description, c.supplierno, s.Name1, p.ProductClass, c.advertisingdate,c.deliveryfrom, c.deliveryto" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when sa.SpecificationID  = 13 then 'Yes' else 'No' end as SA" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when nsw.SpecificationID = 25 then 'Yes' else 'No' end as NSW" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when qld.SpecificationID  = 26 then 'Yes' else 'No' end as QLD" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when act.SpecificationID  = 45 then 'Yes' else 'No' end as ACT" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from cbis599p.dbo.contract c" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.PRODUCT p on p.productcode = c.productcode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.Commoditygroup cg on cg.cgno = p.cgno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.subCommoditygroup scg on scg.cgno = p.cgno and scg.scgno = p.scgno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.supplier s on s.supplierno = c.supplierno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.contractproductspecification sa on sa.ContractNo = c.ContractNo and sa.SpecificationID  = 13" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.contractproductspecification nsw on nsw.ContractNo = c.contractno and nsw.SpecificationID = 25" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.contractproductspecification qld on qld.ContractNo = c.ContractNo and qld.SpecificationID  = 26" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.contractproductspecification act on act.ContractNo = c.ContractNo and act.SpecificationID  = 45" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where c.contractno in (select distinct contractno from #CO)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "order by ContractNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "drop table #CL,#CR, #PS, #CO" & Chr(10)

        Case "CBA_CDS_RetrieveCGsForAudit"
            CBA_DBtoQuery = 9
            CBA_strSQL = CBA_strSQL & "select CGno, SCGno from CG_SCG_CDS" & Chr(10)
        Case "CBA_CDS_AuditCheckData"
            CBA_DBtoQuery = 9
            CBA_strSQL = CBA_strSQL & "select * from CDS_CONT_CHECK order by ContractNo" & Chr(10)
        Case "CBA_CDS_Cont_Map"
            CBA_DBtoQuery = 9
            CBA_strSQL = "select * from Cont_Map" & Chr(10)
        Case "CBA_CDS_Cont_Map_Where"
            CBA_DBtoQuery = 9
            CBA_strSQL = "select * from Cont_Map" & Chr(10)
            CBA_strSQL = CBA_strSQL & "Where Contractno in (" & CBA_StringVal1 & ")" & Chr(10)
            CBA_strSQL = CBA_strSQL & "order by contractno,region,schemeid,containertype,incoterm" & Chr(10)
        Case "CBA_CDS_CONT_CHECK_toInclude"
            CBA_DBtoQuery = 9
            CBA_strSQL = "SELECT Contractno, 501 as Region , 'NSW' as SchemeID FROM CDS_CONT_CHECK WHERE Checked = True and CDS_No_Applies = False and NSW = True" & Chr(10)
            CBA_strSQL = CBA_strSQL & "union SELECT Contractno, 502, 'NSW' FROM CDS_CONT_CHECK WHERE Checked = True and CDS_No_Applies = False and NSW = True" & Chr(10)
            CBA_strSQL = CBA_strSQL & "union SELECT Contractno, 503, 'NSW' FROM CDS_CONT_CHECK WHERE Checked = True and CDS_No_Applies = False and NSW = True" & Chr(10)
            CBA_strSQL = CBA_strSQL & "union SELECT Contractno, 504, 'NSW' FROM CDS_CONT_CHECK WHERE Checked = True and CDS_No_Applies = False and NSW = True" & Chr(10)
            CBA_strSQL = CBA_strSQL & "union SELECT Contractno, 503, 'QLD' FROM CDS_CONT_CHECK WHERE Checked = True and CDS_No_Applies = False and QLD = True" & Chr(10)
            CBA_strSQL = CBA_strSQL & "union SELECT Contractno, 506, 'QLD' FROM CDS_CONT_CHECK WHERE Checked = True and CDS_No_Applies = False and QLD = True" & Chr(10)
            CBA_strSQL = CBA_strSQL & "union SELECT Contractno, 507, 'SA' FROM CDS_CONT_CHECK WHERE Checked = True and CDS_No_Applies = False and SA = True" & Chr(10)
            CBA_strSQL = CBA_strSQL & "union SELECT Contractno, 504, 'ACT' FROM CDS_CONT_CHECK WHERE Checked = True and CDS_No_Applies = False and ACT = True" & Chr(10)
        
        Case "CBA_CDS_CBISQuery"
            CBA_DBtoQuery = 599
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            CBA_strSQL = CBA_strSQL & "SELECT productcode into #BR from cbis599p.dbo.ProductAttribute where ProductAttributeTypeID = 2" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select supplierno, suppadrno, countryCode, Street1, City1, ZipCode" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",case when ISNUMERIC(Zipcode) = 1 and CHARINDEX('.',zipcode,1) = 0 and CountryCode = 'AUS'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "    then case" & Chr(10)
            CBA_strSQL = CBA_strSQL & "    when substring(ZipCode,1,2) >= 10 and substring(ZipCode,1,2) <= 29 then 'NSW'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "    when substring(ZipCode,1,2) >= 02 and substring(ZipCode,1,2) <= 03 then 'ACT'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "    when substring(ZipCode,1,2) >= 30 and substring(ZipCode,1,2) <= 39 then 'VIC'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "    when substring(ZipCode,1,2) >= 80 and substring(ZipCode,1,2) <= 89 then 'VIC'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "    when substring(ZipCode,1,2) >= 40 and substring(ZipCode,1,2) <= 49 then 'QLD'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "    when substring(ZipCode,1,2) >= 90 and substring(ZipCode,1,2) <= 99 then 'QLD'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "    when substring(ZipCode,1,2) >= 50 and substring(ZipCode,1,2) <= 59 then 'SA'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "    when substring(ZipCode,1,2) >= 60 and substring(ZipCode,1,2) <= 69 then 'WA'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "    when substring(ZipCode,1,2) >= 70 and substring(ZipCode,1,2) <= 79 then 'TAS'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "    when substring(ZipCode,1,2) >= 04 and substring(ZipCode,1,2) < 10 then 'NT'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "    end else  '' end as State" & Chr(10)
            CBA_strSQL = CBA_strSQL & " into #SUPA from cbis599p.dbo.SUPPADDRESS" & Chr(10)
            CBA_strSQL = CBA_strSQL & "" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select distinct divno into #DIV from cbis599p.dbo.DIVISION where divno in ('501','502','503','504','505','506','507','509')" & Chr(10)
            CBA_strSQL = CBA_strSQL & "" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select distinct Contractno, ps.SpecificationID as SID, ps.Description as SchemeID, c.OptionID, ol.Description as type" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #CR from cbis599p.dbo.contractproductspecification  c" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.GeDa.ProductSpecificationLng ps on ps.SpecificationID = c.SpecificationID and ps.LanguageID = 0" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.GeDa.ProductSpecificationOptionLng ol on ol.SpecificationID = c.SpecificationID and ol.OptionID = c.OptionID and ol.LanguageID = 0" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where c.SpecificationID in (25,26,13,45)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select * into #CRD from (" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select distinct Contractno, SchemeID, OptionID, type, case  when SID = 13 then 507 when SID = 25 then 501 when SID = 26 then 503 when SID = 45 then 504 end as Divno from #CR" & Chr(10)
            CBA_strSQL = CBA_strSQL & "union select distinct Contractno, SchemeID, OptionID, type, case when SID = 25 then 502 when SID = 26 then 506 end as Divno from #CR" & Chr(10)
            CBA_strSQL = CBA_strSQL & "union select distinct Contractno, SchemeID, OptionID, type, case when SID = 25 then 503 end as Divno from #CR" & Chr(10)
            CBA_strSQL = CBA_strSQL & "union select distinct Contractno, SchemeID, OptionID, type, case when SID = 25 then 504 end as Divno from #CR" & Chr(10)
            CBA_strSQL = CBA_strSQL & ") a where divno is not null" & Chr(10)
            CBA_strSQL = CBA_strSQL & "drop table #CR" & Chr(10)
            CBA_strSQL = CBA_strSQL & "" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select contractno, productcode, c.supplierno, s.name1 into #C" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from cbis599p.dbo.contract c" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.supplier s on s.supplierno = c.supplierno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where contractno in (select contractno from #CRD)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select p.productcode, p.description as PDesc, p.cgno, p.scgno, p.productclass into #P from cbis599p.dbo.product p where productcode in (select distinct productcode from #C)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select distinct c.contractno,dch.divno, productcode, convert(date,deliveryfrom) as deliveryfrom, convert(date,isnull(deliveryto,getdate())) as deliveryto" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #CONT" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from cbis599p.dbo.CONTrACT c" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.DIVCONTRACTHIS dch on dch.ContractNo = c.ContractNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where c.contractno in (select ContractNo from #C)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select ContractNo, divno, convert(date,validfrom) as validfrom, convert(date,isnull(validto,getdate())) as validto, s.SupplierNo, s.SuppAdrNo, s.Street1, s.City1, s.ZipCode, s.State, s.CountryCode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #DCH" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from cbis599p.dbo.DIVCONTRACTHIS dch" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join #SUPA s on s.SupplierNo = dch.PIC_SupplierNo and s.SuppAdrNo = dch.PIC_SuppAdrNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where contractno in (select ContractNo from #C)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select contractno, divno, convert(date,validfrom) as validfrom, convert(date,isnull(validto,getdate())) as validto" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",case when isnull(Exworks,0) > 0 then 'ExWorks' else case when isnull(FOB,0) > 0 then 'FOB' else case when isnull(DDP,0) > 0 then 'DDP' end end end as inco" & Chr(10)
            CBA_strSQL = CBA_strSQL & "into #INCO from cbis599p.dbo.braketvalue" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where contractno in (select ContractNo from #C)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "" & Chr(10)
            CBA_strSQL = CBA_strSQL & "" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select distinct c.contractno, c.ProductCode" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",p.PDesc, p.cgno,p.scgno,p.productclass" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", div.DivNo--, coalesce(crd1.SID,crd2.SID,crd3.SID,crd4.SID) as SID" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",c.supplierno, c.name1" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",crd.SchemeID" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",crd.type" & Chr(10)
            CBA_strSQL = CBA_strSQL & "--, cont.deliveryfrom, cont.deliveryto" & Chr(10)
            CBA_strSQL = CBA_strSQL & "--, dch.validfrom, dch.validto" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",i.inco" & Chr(10)
            CBA_strSQL = CBA_strSQL & "--, i.validfrom, i.validto" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",dch.Street1, dch.City1, dch.State, dch.CountryCode" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",case when isnull(br.ProductCode,0) = c.ProductCode then 'Branded' else '' end as Branded" & Chr(10)
            CBA_strSQL = CBA_strSQL & "" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from #C c" & Chr(10)
            CBA_strSQL = CBA_strSQL & "cross join #DIV div" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join #P p on p.productcode = c.productcode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join #CRD crd on crd.ContractNo = c.ContractNo and crd.Divno = div.DivNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join #CONT cont on cont.ContractNo = c.contractno and cont.DivNo = div.DivNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join #DCH dch on dch.ContractNo = c.ContractNo and dch.DivNo = div.DivNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join #INCO i on i.ContractNo = c.ContractNo and i.DivNo = div.DivNo" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join #BR br on br.ProductCode = c.ProductCode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where cont.deliveryto >= '2017-01-01'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and cont.deliveryfrom <= getdate()" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and i.validto >= '2017-01-01'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and i.validfrom <= getdate()" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and schemeid is not null" & Chr(10)
            CBA_strSQL = CBA_strSQL & "--and (crd.SID" & Chr(10)
            CBA_strSQL = CBA_strSQL & "" & Chr(10)
            CBA_strSQL = CBA_strSQL & "drop table #BR, #SUPA, #DIV, #CRD, #C, #P, #CONT, #DCH, #INCO" & Chr(10)
        
        Case "CBA_CDS_StoreReceiving"
            CBA_DBtoQuery = CBA_LongVal1
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "DECLARE @SDATE date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "DECLARE @EDATE date = '" & Format(CBA_Dateto, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select c.contractno, c.productcode into #C  from purchase.dbo.contract c where contractno in (" & CBA_StringVal1 & ")" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select sr.productcode,sr.ReceivingDate, sr.storeno,  sum(sr.Quantity) as QTY" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from purchase.dbo.StoreReceiving sr " & Chr(10)
            CBA_strSQL = CBA_strSQL & "where sr.ReceivingDate >= @SDATE and sr.ReceivingDate <= @EDATE" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and sr.productcode in (select distinct productcode from #C)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "group by  sr.productcode, sr.ReceivingDate, sr.storeno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "order by  sr.productcode, sr.ReceivingDate, sr.storeno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "drop table #C" & Chr(10)
        Case "CBA_CDS_CBIS_BaseData"
            CBA_DBtoQuery = 599
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select c.contractno, c.productcode, c.supplierno into #C from cbis599p.dbo.contract c" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where contractno in (" & CBA_StringVal1 & ")" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select productcode, 1 as branded into #PA from cbis599p.dbo.ProductAttribute where ProductAttributeTypeID = 2" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select c.contractno, c.productcode, p.description, c.supplierno, s.Name1, isnull(pa.branded, 0) as branded, p.packsize " & Chr(10)
            CBA_strSQL = CBA_strSQL & "from #C c left join #PA pa on pa.ProductCode = c.productcode left join cbis599p.dbo.product p on p.productcode = c.productcode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.SUPPLIER s on s.supplierno = c.supplierno" & Chr(10)
            CBA_strSQL = CBA_strSQL & "drop table #C, #PA" & Chr(10)
        Case "CBA_CDS_CBIS_IncoTerms"
            CBA_DBtoQuery = 599
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "DECLARE @SDATE date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "DECLARE @EDATE date = '" & Format(CBA_Dateto, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select Contractno, divno, case when validfrom < @SDATE then @SDATE else Validfrom end as ValidFrom" & Chr(10)
            CBA_strSQL = CBA_strSQL & ", case when isnull(validto, convert(date,getdate())) > @EDATE then @EDATE else isnull(validto, convert(date,getdate())) end as validto" & Chr(10)
            CBA_strSQL = CBA_strSQL & ",case when FOB is null and ExWorks is null and DDP is not null then 'DDP' when FOB is null and ExWorks is not null then 'ExWorks'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "when FOB is not null and ExWorks is null then 'FOB' else 'Unknown' end as incoterms" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from cbis599p.dbo.BRAKETVALUE where contractno in (" & CBA_StringVal1 & ")" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and isnull(validto, convert(date,getdate())) >= @SDATE and validfrom <= @EDATE" & Chr(10)
            CBA_strSQL = CBA_strSQL & "order by ContractNo, divno" & Chr(10)
        Case "CBA_AST_StoreNos"
            CBA_DBtoQuery = 599
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "DECLARE @SDATE date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select divno, NoOfStores from cbis599p.Portfolio.Stores s" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where s.Validfrom <= @SDATE and s.Validto >= @SDATE" & Chr(10)
        Case "CBA_AST_ALLStoreNos"
            CBA_DBtoQuery = 599       ' #RW 190801 new SQL as the old one could bring back a null
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "DECLARE @TDATE date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "DECLARE @lastdate date = (select max(validto) from cbis599p.portfolio.Stores)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "DECLARE @SDATE date = (select case when @lastdate < @TDATE then @lastdate else @TDATE end)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select sum(NoOfStores) from cbis599p.Portfolio.Stores s" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where s.Validfrom >= @SDATE AND isnull(s.Validto,@SDATE) <= @SDATE" & Chr(10)
        Case "CBA_AST_POSbDiv"
            CBA_DBtoQuery = 599
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @PROD int = " & CBA_LongVal1 & Chr(10)
            CBA_strSQL = CBA_strSQL & "DECLARE @SDATE date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @EDATE date = '" & Format(CBA_Dateto, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select divno, sum(quantity), sum(retail) from cbis599p.dbo.pos pos" & Chr(10)
            CBA_strSQL = CBA_strSQL & "left join cbis599p.dbo.product p on p.productcode = pos.productcode" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where isnull(p.Con_ProductCode,p.productcode) = @PROD" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and pos.posdate >= @SDATE and pos.PosDate <= @EDATE group by Divno" & Chr(10)
        Case "CBA_AST_CurrentPrice"
            CBA_DBtoQuery = 599
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "declare @PROD int = " & CBA_LongVal1 & Chr(10)
            CBA_strSQL = CBA_strSQL & "select divno, Retail from cbis599p.dbo.divretail dr" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where productcode = @PROD and validfrom <= convert(date,getdate())" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and isnull(validto,convert(date,getdate())) >= convert(date,getdate())" & Chr(10)
        Case "CBA_CDS_QuantityCount(COMRADE)"
            CBA_DBtoQuery = 599
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select distinct p.productcode, p.cgno, p.scgno, p.productclass, p.description, p.packsize, p.unitcode_wght, p.unitcode_Vol, p.unitcode_Pack," & Chr(10)
            CBA_strSQL = CBA_strSQL & "p.israndomweight, p.iscoldstorage, p.isfrozen, p.isfood, p.isdirect, p.weightvolume, p.contents, p.weight, p.weightKG," & Chr(10)
            CBA_strSQL = CBA_strSQL & "p.volume , p.pricecrddesc, p.pricecrddoubleinf, unitpricebase  from cbis599p.dbo.product as P where p.productcode in ("
            CBA_strSQL = CBA_strSQL & CBA_StringVal1 & ")" & Chr(10)
        Case "CBA_getUniquePCodesFromContractNo"
            CBA_DBtoQuery = 599
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select distinct productcode from cbis599p.dbo.contract where contractno in (" & CBA_StringVal1 & ")"
        Case "CBA_FilterActiveContracts"
            CBA_DBtoQuery = 599
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "DECLARE @DFROM date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "DECLARE @DTO date = '" & Format(CBA_Dateto, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select distinct contractno from cbis599p.dbo.contract where dateadd(MM,12,isnull(deliveryto,getdate())) >= @DFROM and isnull(deliveryfrom,@DFROM) <= @DTO and contractno in (" & CBA_StringVal1 & ")"
        Case "CBA_RCV15EitherSide"
            CBA_DBtoQuery = 599
            CBA_strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            CBA_strSQL = CBA_strSQL & "DECLARE @DFROM date = '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
            CBA_strSQL = CBA_strSQL & "select contractno, productcode, divno, sum(quantity)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "from cbis599p.dbo.RECEIVING" & Chr(10)
            CBA_strSQL = CBA_strSQL & "where DayEndDate >= dateadd(D,-15,@DFROM)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and dayenddate <= dateadd(D,15,@DFROM)" & Chr(10)
            CBA_strSQL = CBA_strSQL & "and contractno in (" & CBA_StringVal1 & ")" & Chr(10)
            CBA_strSQL = CBA_strSQL & "group by divno, contractno, productcode order by divno, ProductCode, contractno" & Chr(10)
    
    End Select
    'Debug.Print CBA_strSQL
    
    If CBA_strSQL <> "" Then
        If CBA_DBtoQuery = 1 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "ABI_QUERY", CBA_BSA & "LIVE DATABASES\ABI.accdb", CBA_MSAccess, CBA_strSQL, 120, , , False)  'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 2 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "ABI_QUERY", CBA_BSA & "LIVE DATABASES\ABI.accdb", CBA_MSAccess, CBA_strSQL, 120, , , False)  'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 3 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "COMRADE_QUERY", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLNCLI10", CBA_strSQL, 120, , , False) 'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 4 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "COMRADE_QUERY", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLNCLI10", CBA_strSQL, 120, , , False) 'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 5 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "SSDB_QUERY", CBA_BSA & "LIVE DATABASES\SuperSaverStoreStock.accdb", CBA_MSAccess, CBA_strSQL, 120, , , False)   'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 6 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "SSDB_QUERY", CBA_BSA & "LIVE DATABASES\SuperSaverStoreStock.accdb", CBA_MSAccess, CBA_strSQL, 120, , , False)  'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 7 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CBForecast_QUERY", g_GetDB("ForeCast"), CBA_MSAccess, CBA_strSQL, 120, , , False)   'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 8 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "CBForecast_UPDATE", g_GetDB("ForeCast"), CBA_MSAccess, CBA_strSQL, 120, , , False)   'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 9 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CDS_ADM_QUERY", CBA_BSA & "LIVE DATABASES\CDS_ADM.accdb", CBA_MSAccess, CBA_strSQL, 120, , , False)   'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 10 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "CDS_ADM_UPDATE", CBA_BSA & "LIVE DATABASES\CDS_ADM.accdb", CBA_MSAccess, CBA_strSQL, 120, , , False)    'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery > 500 And CBA_DBtoQuery < 510 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "MMS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", CBA_strSQL, 120, , , False)    'Runs DB_Connection module to create connection to dtabase and run query
        Else
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CBIS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", CBA_strSQL, 120, , , False)   'Runs DB_Connection module to create connection to dtabase and run query
        End If
    End If
    
    If bOutput = False Then
        CBA_GenPullSQL = False
    Else
        CBA_DB_Connect.CBA_DB_CBADBUpdate "UPDATE", "CBADB_QUERY", CBA_BSA & "LIVE DATABASES\CBADB.accdb", CBA_MSAccess, _
            "INSERT INTO CBA_QueryLogging ([DateTimeStarted], [DateTimeEnded], [CBA_SQLType], [User], [DBtoQuery])" & Chr(10) & _
            "VALUES ('" & CBA_QueryTimerStart & "' , '" & Now & "', '" & CBA_SQLType & "' , '" & CBA_BasicFunctions.CBA_Current_UserID & "', " & CBA_DBtoQuery & ")"
        CBA_GenPullSQL = True
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_GenPullSQL", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & vbCrLf & CBA_strSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function




