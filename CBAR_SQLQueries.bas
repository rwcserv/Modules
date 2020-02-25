Attribute VB_Name = "CBAR_SQLQueries"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Function CBAR_GenPullSQL(ByRef SQLType, Optional ByRef DateFrom, Optional ByRef DateTo, Optional ByRef cCG As Long, Optional ByRef scg As String, Optional ByRef cType, Optional ByRef Bypass As String)
    Dim Ref As Variant
    'SQLType  - Defines the SQL Query used)
    'DIFFERENT TYPES OF QUERY
    'SHP_QTY            QUANTITY FROM SHIPPING TABLE
    'SHP_RET            RETAIL FROM SHIPPING TABLE
    'SHP_LINE           TOTAL LINES
    'RCV_COST           COST FROM RECIEVING TABLE
    'RCV_RET            RETAIL FROM RECIEVING TABLE
    'RCV_RETN           RETAILNET FROM RECIEVING TABLE
    'PORT_STORE         Number of stores from DateTo less 6 days
    Dim CBA_QueryTimerStart, strSQL As String, strPSQL As String, CG
    Dim sStates As String, sStateLook As String, strProds As String, curCtype, bOutput As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    'wks_Data.Unprotect "thepasswordispassword"
    If CBA_DBtoQuery = 0 Then CBA_DBtoQuery = 599
    If LCase(Mid(SQLType, 1, 4)) = "cbis" Then
        CBA_DBtoQuery = 599
    ElseIf LCase(Mid(SQLType, 1, 3)) = "COM" Then
        CBA_DBtoQuery = 3
    End If
    CBA_QueryTimerStart = Now
    DoEvents
    
    If SQLType = "COM_CheckT150List" Then
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
        strSQL = strSQL & "select productcode, description,cgno into #PROD from cbis599p.dbo.product where isnull(con_productcode, productcode) in (" & Bypass & ")" & Chr(10)
        strSQL = strSQL & "select a.productcode, p.description,p.cgno, sum(retail) as retail into #POS from (" & Chr(10)
        strSQL = strSQL & "select isnull(p.con_productcode,pos.productcode) as productcode , sum(pos.retail) as retail" & Chr(10)
        strSQL = strSQL & "from cbis599p.dbo.pos pos" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.product p on p.productcode = pos.ProductCode" & Chr(10)
        strSQL = strSQL & "inner join #PROD po on po.productcode = pos.ProductCode" & Chr(10)
        strSQL = strSQL & "where pos.posdate >= DateAdd(WK, -2, getdate())" & Chr(10)
        strSQL = strSQL & "group by isnull(p.con_productcode,pos.productcode)" & Chr(10)
        strSQL = strSQL & ") a left join #PROD p on p.productcode = a.productcode group by a.productcode, p.description, p.cgno" & Chr(10)
        strSQL = strSQL & "select pos.productcode, pos.Description, pos.cgno" & Chr(10)
        strSQL = strSQL & "from #POS pos" & Chr(10)
        strSQL = strSQL & "where pos.Retail < 9999" & Chr(10)
        strSQL = strSQL & "drop table #PROD, #POS" & Chr(10)
    End If
    
    
    If SQLType = "CBAR_T150NewPotentials" Then
    'strPSQL = ""
        For Each Ref In cType
            If Ref <> "" Then
                If strPSQL = "" Then strPSQL = "where charindex('" & LCase(Ref) & "',lower(description),1) > 0 " & Chr(10) Else strPSQL = strPSQL & "or charindex('" & LCase(Ref) & "',lower(description),1) > 0 " & Chr(10)
            End If
        Next
        If strPSQL = "" Then ReDim CBA_CBISarr(0 To 0, 0 To 0): CBA_CBISarr(0, 0) = 0: Exit Function
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
        strSQL = strSQL & "select distinct productcode, description, Con_ProductCode, cgno, scgno into #PROD from cbis599p.dbo.product " & Chr(10)
        strSQL = strSQL & strPSQL
        strSQL = strSQL & "select a.productcode, description, p.cgno from (select isnull(p.con_productcode,p.productcode) as productcode , sum(retail) as ret" & Chr(10)
        strSQL = strSQL & "from cbis599p.dbo.pos inner join #PROD p on p.ProductCode = pos.ProductCode" & Chr(10)
        strSQL = strSQL & "where pos.posdate >= dateadd(WK,-2,getdate()) group by isnull(p.con_productcode,p.productcode), description) a" & Chr(10)
        strSQL = strSQL & "left join #PROD p on p.ProductCode = a.productcode where a.ret > 9999" & Chr(10)
        strSQL = strSQL & "drop table #PROD" & Chr(10)
    End If
    
    
    If SQLType = "COM_MMRCore" Then
            CBA_DBtoQuery = 3
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            strSQL = strSQL & "----****Aldi Prod Listing****----" & Chr(10)
            
            strSQL = strSQL & "select stockcode into #W from tools.dbo.com_w_prod group by stockcode" & Chr(10)
            strSQL = strSQL & "select a.stockcode as Ccode, b.packagesize as packsize into #WWProds from #W a" & Chr(10)
            strSQL = strSQL & "left join (select distinct stockcode, packagesize from tools.dbo.com_w_prod where datescraped = '" & Format(DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "group by stockcode, packagesize ) b on a.stockcode = b.stockcode" & Chr(10)
            strSQL = strSQL & "select colesproductid into #C from tools.dbo.com_c_prod group by colesproductid" & Chr(10)
            strSQL = strSQL & "select a.colesproductid as Ccode, b.packsize into #ColesProds from #C a" & Chr(10)
            strSQL = strSQL & "left join (select distinct colesproductid, packsize from tools.dbo.com_c_prod where datescraped = '" & Format(DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "group by colesproductid, packsize ) b on a.colesproductid = b.colesproductid" & Chr(10)
            strSQL = strSQL & "select distinct A_Code, A_product, '' as Buyer, case when C_CCode is null and C_CCode2 is null and C_CCode3 is null and C_CCode4 is null and C_CVCode is null" & Chr(10)
            strSQL = strSQL & "and C_CVCode2 is null and C_CVCode3 is null and C_SBCode is null and C_SBCode2 is null then 'No' else 'Yes' end as [ColesPL]" & Chr(10)
            strSQL = strSQL & ",case when W_HBCode is null and W_HBCode2 is null and W_HBCode3 is null and W_SelCode is null and W_SelCode2 is null" & Chr(10)
            strSQL = strSQL & "and W_SelCode3 is null and W_SelCode4 is null and W_SelCode5 is null then 'No' else 'Yes' end as [WWPL]" & Chr(10)
            strSQL = strSQL & ",case when B_MLCode is null and B_MLCode2 is null and B_MLCode5 is null then 'No' else 'Yes' end as [ColesML]" & Chr(10)
            strSQL = strSQL & ",case when B_MLCode3 is null and B_MLCode4 is null and B_MLCode6 is null then 'No' else 'Yes' end as [WWML]" & Chr(10)
            strSQL = strSQL & ",case when B_CBCode is null then 'No' else 'Yes' end as [ColesCB] ,case when B_CBCode3 is null then 'No' else 'Yes' end as [WWCB]" & Chr(10)
            strSQL = strSQL & ",case when B_PBCode is null then 'No' else 'Yes' end as [ColesPB] ,case when B_PBCode3 is null then 'No' else 'Yes' end as [WWPB]" & Chr(10)
            strSQL = strSQL & "into #Prods from tools.dbo.com_prodmap p" & Chr(10)
            strSQL = strSQL & "left join #Colesprods c on c.Ccode = p.C_CCode or c.Ccode = p.C_CCode2 or c.Ccode = p.C_CCode3 or c.Ccode = p.C_CCode4 or c.Ccode = p.C_CVCode" & Chr(10)
            strSQL = strSQL & "or c.Ccode = p.C_CVCode2 or c.Ccode = p.C_CVCode3 or c.Ccode = p.C_SBCode or c.Ccode = p.C_SBCode2 or c.Ccode = p.B_MLCode" & Chr(10)
            strSQL = strSQL & "or c.Ccode = p.B_MLCode2 or c.Ccode = p.B_MLCode5 or c.Ccode = p.B_CBCode or c.Ccode = p.B_PBCode" & Chr(10)
            strSQL = strSQL & "left join #WWprods w on w.Ccode = p.W_HBCode or w.Ccode = p.W_HBCode2 or w.Ccode = p.W_HBCode3 or w.Ccode = p.W_SelCode" & Chr(10)
            strSQL = strSQL & "or w.Ccode = p.W_SelCode2 or w.Ccode = p.W_SelCode3 or w.Ccode = p.W_SelCode4 or w.Ccode = p.W_SelCode5 or w.Ccode = p.B_MLCode3" & Chr(10)
            strSQL = strSQL & "or w.Ccode = p.B_MLCode4 or w.Ccode = p.B_MLCode6 or w.Ccode = p.B_CBCode3 or w.Ccode = p.B_PBCode3" & Chr(10)
'            strSQL = strSQL & "where A_CG <> 58 and not (A_CG = 51 and A_SCG = 40) and A_CG > 4 and (w.Packsize is not null or c.packsize is not null)" & Chr(10)
            strSQL = strSQL & "where A_CG <> 58 and A_CG > 4 and (w.Packsize is not null or c.packsize is not null)" & Chr(10)
            strSQL = strSQL & "and A_Code in (" & Bypass & ")" & Chr(10)
            strSQL = strSQL & "select * from #Prods where ColesPL = 'No' and WWPL = 'No'" & Chr(10)
            strSQL = strSQL & "drop table #C, #W, #WWProds, #ColesProds,#Prods" & Chr(10)
            'Debug.Print strSQL
    End If
    
    
    If SQLType = "COM_MMRProduce" Then
            CBA_DBtoQuery = 3
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
            strSQL = strSQL & "----****Aldi Produce Prod Listing****----" & Chr(10)
            strSQL = strSQL & "select distinct stockcode as Ccode into #WWProds from tools.dbo.com_w_prod where datescraped = '" & Format(DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "select distinct colesproductid as Ccode into #ColesProds from tools.dbo.com_c_prod where datescraped = '" & Format(DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "select distinct A_Code,A_Product, case when C_Code is null and C_CCode is null and C_CCode2 is null and C_CCode3 is null then 'No' else 'Yes' end as [ColesNational]" & Chr(10)
            strSQL = strSQL & ",case when C_Code2 is null then 'No' else 'Yes' end as [ColesNSW]   ,case when C_Code3 is null then 'No' else 'Yes' end as [ColesVIC]" & Chr(10)
            strSQL = strSQL & ",case when C_Code4 is null then 'No' else 'Yes' end as [ColesQLD]   ,case when C_Code5 is null then 'No' else 'Yes' end as [ColesSA]" & Chr(10)
            strSQL = strSQL & ",case when C_Code6 is null then 'No' else 'Yes' end as [ColesWA]" & Chr(10)
            strSQL = strSQL & ",case when W_HBCode is null and W_HBCode2 is null and W_HBCode3 is null and W_Code is null then 'No' else 'Yes' end as [WWNational]" & Chr(10)
            strSQL = strSQL & ",case when W_Code2 is null then 'No' else 'Yes' end as [WWNSW]  ,case when W_Code3 is null then 'No' else 'Yes' end as [WWVIC]" & Chr(10)
            strSQL = strSQL & ",case when W_Code4 is null then 'No' else 'Yes' end as [WWQLD]  ,case when W_Code5 is null then 'No' else 'Yes' end as [WWSA]" & Chr(10)
            strSQL = strSQL & ",case when W_Code6 is null then 'No' else 'Yes' end as [WWWA] into #PProds from tools.dbo.com_prodmap p" & Chr(10)
            strSQL = strSQL & "inner join #Colesprods c on c.Ccode = p.C_Code or c.Ccode = p.C_CCode  or c.Ccode = p.C_CCode2  or c.Ccode = p.C_CCode3  or c.Ccode = p.C_Code2 or c.Ccode = p.C_Code3  or c.Ccode = p.C_Code4 or c.Ccode = p.C_Code5  or c.Ccode = p.C_Code6 " & Chr(10)
            strSQL = strSQL & "inner join #WWprods w on w.Ccode = p.W_HBCode or w.Ccode = p.W_HBCode2 or w.Ccode = p.W_HBCode3 or w.Ccode = p.W_Code or w.Ccode = p.W_Code2 or w.Ccode = p.W_Code3 or w.Ccode = p.W_Code4 or w.Ccode = p.W_Code5" & Chr(10)
            strSQL = strSQL & "where A_CG in (58,27) and A_Code in (" & Bypass & ")" & Chr(10)
            strSQL = strSQL & "union select distinct A_Code,A_Product, case when C_Code is null and C_CCode is null and C_CCode2 is null and C_CCode3 is null then 'No' else 'Yes' end as [ColesNational]" & Chr(10)
            strSQL = strSQL & ",case when C_Code2 is null then 'No' else 'Yes' end as [ColesNSW]   ,case when C_Code3 is null then 'No' else 'Yes' end as [ColesVIC]" & Chr(10)
            strSQL = strSQL & ",case when C_Code4 is null then 'No' else 'Yes' end as [ColesQLD]   ,case when C_Code5 is null then 'No' else 'Yes' end as [ColesSA]" & Chr(10)
            strSQL = strSQL & ",case when C_Code6 is null then 'No' else 'Yes' end as [ColesWA]" & Chr(10)
            strSQL = strSQL & ",case when W_HBCode is null and W_HBCode2 is null and W_HBCode3 is null and W_Code is null then 'No' else 'Yes' end as [WWNational]" & Chr(10)
            strSQL = strSQL & ",case when W_Code2 is null then 'No' else 'Yes' end as [WWNSW]  ,case when W_Code3 is null then 'No' else 'Yes' end as [WWVIC]" & Chr(10)
            strSQL = strSQL & ",case when W_Code4 is null then 'No' else 'Yes' end as [WWQLD]  ,case when W_Code5 is null then 'No' else 'Yes' end as [WWSA]" & Chr(10)
            strSQL = strSQL & ",case when W_Code6 is null then 'No' else 'Yes' end as [WWWA] from tools.dbo.com_prodmap p" & Chr(10)
            strSQL = strSQL & "where A_CG in (58,27) and A_Code in (" & Bypass & ")" & Chr(10)
            strSQL = strSQL & "and ((p.C_Code is null and p.C_CCode is null and p.C_CCode2 is null and p.C_CCode3 is null and p.C_Code2 is null and p.C_Code3 is null and p.C_Code4 is null and p.C_Code5 is null and p.C_Code6  is null ) " & Chr(10)
            strSQL = strSQL & "or (p.W_HBCode is null and p.W_HBCode2 is null and p.W_HBCode3 is null and p.W_Code is null and p.W_Code2 is null and p.W_Code3 is null and p.W_Code4 is null and p.W_Code5 is null))" & Chr(10)
            
            strSQL = strSQL & "select * from #PProds" & Chr(10)
            strSQL = strSQL & "where ((ColesNational = 'No' and (ColesNSW = 'No' or ColesVIC = 'No' or ColesQLD = 'No' or ColesSA = 'No' or ColesWA = 'No' ))" & Chr(10)
            strSQL = strSQL & "or (WWNational = 'No' and (WWNSW = 'No' or WWVIC = 'No' or WWQLD = 'No' or WWSA = 'No' or WWWA = 'No' )))" & Chr(10)
            strSQL = strSQL & "drop table #WWProds, #ColesProds,#PProds" & Chr(10)
            'Debug.Print strSQL
    End If
    
    
    
    If SQLType = "CBIS_RefreshComradeAldiProdListing" Then
            CBA_DBtoQuery = 599
            strSQL = "select distinct isnull(p.con_productcode,p.productcode) as pcode, emp.firstname + ' ' + emp.name as Buyer" & Chr(10)
            strSQL = strSQL & "from cbis599p.dbo.contract c" & Chr(10)
            strSQL = strSQL & "inner join cbis599p.dbo.product p on p.productcode = c.productcode" & Chr(10)
            strSQL = strSQL & "inner join cbis599p.dbo.divcontracthis dc on dc.contractno = c.contractno" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.employee emp on emp.empno = p.empno" & Chr(10)
            strSQL = strSQL & "where not p.cgno in (58, 61, 27) and dc.sendtodiv = 1" & Chr(10)
            strSQL = strSQL & "and ((p.productclass = 1 and c.deliveryfrom  <= getdate() and isnull(c.deliveryto,getdate()) >= getdate())" & Chr(10)
            strSQL = strSQL & "or (p.productclass = 4 and p.seasonid in (5,6,7) and dateadd(mm,-1,c.deliveryfrom)  <= getdate() and isnull(c.deliveryto,getdate()) >= getdate()))" & Chr(10)
            If CG > 0 Then strSQL = strSQL & "and p.cgno = " & CG & Chr(10)
            If scg <> "0" And scg <> "" Then strSQL = strSQL & "and p.scgno = " & scg & Chr(10)
            If Bypass <> "" Then strSQL = strSQL & "and emp.name  = '" & Bypass & "'" & Chr(10)
            If cType <> "" Then strSQL = strSQL & "and emp.empno_grp  = (Select empno from cbis599p.dbo.employee where firstname + ' ' + name = '" & cType & "')" & Chr(10)
            strSQL = strSQL & "Order by pcode" & Chr(10)
            'Debug.Print strSQL
    End If
    
    If SQLType = "MMS_RefreshComradeAldiProdListing" Then
    
            strSQL = "select distinct substring(@@SERVERNAME,1,3) as Divno, isnull(p.con_productcode, c.ProductCode) as productcode, p.Description from Purchase.PTPImport.ContractHistory c" & Chr(10)
            strSQL = strSQL & "left join purchase.dbo.Product p on p.ProductCode = c.ProductCode where FromDate <= '" & Format(DateTo, "YYYY-MM-DD") & "' and ToDate >= '" & Format(DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            
    '        strSQL = "select distinct substring(@@SERVERNAME,1,3) as Divno, p.productcode from purchase.dbo.product p left join purchase.dbo.contract c on c.productcode = p.productcode" & Chr(10)
    '        strSQL = strSQL & "where p.cgno in (27,58) and c.Active = 1" & Chr(10)
    '      '  Debug.Print strSQL
    End If
    
    If SQLType = "COM_ClassProdMapHis" Then
    
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
        strSQL = strSQL & "select b.datechanged,b.aldiprod, b.comppcode, b.comptype, b.Cgno" & Chr(10)
        strSQL = strSQL & ", row_number() over (Partition by aldiprod, comptype order by datechanged) as row" & Chr(10)
        strSQL = strSQL & ", count(comppcode) over (Partition by aldiprod, comptype) as cnt" & Chr(10)
        strSQL = strSQL & "into #ChangeComps from ( select a.datechanged,a.aldiprod, a.comppcode, a.comptype, a.Cgno" & Chr(10)
        strSQL = strSQL & "from (select datechanged, aldiprod, comppcode, comptype, pm.A_CG as CGno" & Chr(10)
        strSQL = strSQL & ", rank() over (partition by aldiprod, comptype, convert(date, datechanged) order by datechanged desc)  as row" & Chr(10)
        strSQL = strSQL & "from tools.dbo.com_mapchange mc left join tools.dbo.com_prodmap pm on pm.A_Code = mc.Aldiprod" & Chr(10)
        If cCG <> 0 And scg = "0" Then strSQL = strSQL & "where aldiprod in (select A_Code from tools.dbo.com_prodmap where A_CG = " & cCG & ")" & Chr(10)
        If cCG <> 0 And scg <> "0" Then strSQL = strSQL & "where aldiprod in (select A_Code from tools.dbo.com_prodmap where A_CG = " & cCG & " and A_SCG = " & scg & ")" & Chr(10)
        If cCG = 0 And scg = "0" And Bypass <> "" Then strSQL = strSQL & "where aldiprod in (" & Bypass & ")" & Chr(10)
        strSQL = strSQL & ") a where a.row = 1 and convert(date, datechanged) <= '" & Format(DateTo, "YYYY-MM-DD") & "') b" & Chr(10)
        If Not IsMissing(cType) Then
            If cType <> "" Then
                If InStr(1, LCase(strSQL), "where") > 0 Then
                    strSQL = strSQL & "and charindex('" & cType & "',lower(b.comptype),1) > 0 " & Chr(10)
                Else
                    strSQL = strSQL & "where charindex('" & cType & "',lower(b.comptype),1) > 0 " & Chr(10)
                End If
            End If
        End If
        
        strSQL = strSQL & "order by aldiprod,comptype, datechanged desc" & Chr(10)
        strSQL = strSQL & "select Aldiprod, Comptype, Validfrom, Validto, CCode, CGno from (select Aldiprod, comptype,case when row = 1 then '" & Format(DateFrom, "YYYY-MM-DD") & "' else" & Chr(10)
        strSQL = strSQL & "case when convert(date,datechanged) < '" & Format(DateFrom, "YYYY-MM-DD") & "' then '" & Format(DateFrom, "YYYY-MM-DD") & "' else convert(date,datechanged) end end as validfrom" & Chr(10)
        strSQL = strSQL & ",case when row = cnt then '" & Format(DateTo, "YYYY-MM-DD") & "'  else" & Chr(10)
        strSQL = strSQL & "(select dateadd(dd,-1,convert(date,datechanged)) from #ChangeComps c where c.comptype = a.comptype and c.aldiprod = a.aldiprod and c.row = a.row + 1)" & Chr(10)
        strSQL = strSQL & "end as validto ,comppcode as CCode, CGno from (select *" & Chr(10)
        strSQL = strSQL & "from #ChangeComps cc" & Chr(10)
        strSQL = strSQL & ") a ) b where  validto > '" & Format(DateFrom, "YYYY-MM-DD") & "' and CCode <> 'Unassigned' order by Aldiprod" & Chr(10)
        strSQL = strSQL & "drop table  #ChangeComps" & Chr(10)
    '  '  Debug.Print strSQL
    
    End If
    
    If SQLType = "COM_ClassProdMap" Then
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
        If Bypass <> "" Then strSQL = strSQL & "select A_Code into #Prods from tools.dbo.Com_ProdMap where A_Code in (" & Bypass & ")" & Chr(10)
        
        strSQL = strSQL & "select distinct a.A_Code, a.A_CG as CG, a.A_SCG as SCG, a.matchtype, a.CCode, a.cpack from (" & Chr(10)
        strSQL = strSQL & "select cd.A_Code, cd.A_CG, cd.A_SCG, cd.CCode, cd.Matchtype, cd.cpack from (" & Chr(10)
        strSQL = strSQL & "select A_Code, A_CG, A_SCG,  B_CBCode as Ccode, 'ColesCB' as Matchtype, null as Cpack from tools.dbo.Com_ProdMap where B_CBCode is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code,A_CG, A_SCG,  B_MLCode , 'ColesML1', null from tools.dbo.Com_ProdMap where B_MLCode  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code,A_CG, A_SCG,  B_MLCode2 , 'ColesML2', null from tools.dbo.Com_ProdMap where B_MLCode2  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code,A_CG, A_SCG,  B_MLCode5 , 'ColesML3', null from tools.dbo.Com_ProdMap where B_MLCode5  is not null and A_CG > 4" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, B_PBCode , 'ColesPB', null from tools.dbo.Com_ProdMap where B_PBCode  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CCode , 'ColesColes', null from tools.dbo.Com_ProdMap where C_CCode  is not null and A_CG <> 58" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CCode2 , 'ColesColes1', null from tools.dbo.Com_ProdMap where C_CCode2  is not null and A_CG <> 58" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CCode3 , 'ColesColes2', null from tools.dbo.Com_ProdMap where C_CCode3  is not null and A_CG <> 58" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CCode4 , 'ColesColes3', null from tools.dbo.Com_ProdMap where C_CCode4  is not null and A_CG <> 58" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CCode , 'ColesWNAT1', null from tools.dbo.Com_ProdMap where C_CCode  is not null and A_CG = 58" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CCode2 , 'ColesWNAT2', null from tools.dbo.Com_ProdMap where C_CCode2  is not null and A_CG = 58" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CCode3 , 'ColesWNAT3', null from tools.dbo.Com_ProdMap where C_CCode3  is not null and A_CG = 58" & Chr(10)
        'strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CCode4 , 'ColesWNAT4', null from tools.dbo.Com_ProdMap where C_CCode4  is not null and A_CG = 58" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_Code , 'ColesW', null from tools.dbo.Com_ProdMap where C_Code  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_Code2 , 'ColesWNSW', null from tools.dbo.Com_ProdMap where C_Code2  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_Code3 , 'ColesWVIC', null from tools.dbo.Com_ProdMap where C_Code3  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_Code4 , 'ColesWQLD', null from tools.dbo.Com_ProdMap where C_Code4  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_Code5 , 'ColesWSA', null from tools.dbo.Com_ProdMap where C_Code5  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_Code6 , 'ColesWWA', null from tools.dbo.Com_ProdMap where C_Code6  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CVCode , 'ColesVal', null from tools.dbo.Com_ProdMap where C_CVCode  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CVCode2 , 'ColesVal1', null from tools.dbo.Com_ProdMap where C_CVCode2  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CVCode3 , 'ColesVal2', null from tools.dbo.Com_ProdMap where C_CVCode3  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_SBCode , 'ColesSB', null from tools.dbo.Com_ProdMap where C_SBCode  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_SBCode2 , 'ColesSB1', null from tools.dbo.Com_ProdMap where C_SBCode2  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_SBCode3 , 'ColesSB2', null from tools.dbo.Com_ProdMap where C_SBCode3  is not null" & Chr(10)
        If Bypass <> "" Then
            strSQL = strSQL & ") as cd inner join #Prods p on p.A_Code = cd.A_Code union select wd.A_Code,wd.A_CG, wd.A_SCG, wd.CCode, wd.Matchtype, Cpack from (" & Chr(10)
        Else
            strSQL = strSQL & ") as cd union select wd.A_Code,wd.A_CG, wd.A_SCG, wd.CCode, wd.Matchtype, Cpack from (" & Chr(10)
        End If
        strSQL = strSQL & "select A_Code, A_CG, A_SCG,  B_CBCode3 as Ccode, 'WWCB' as Matchtype, null as Cpack from tools.dbo.Com_ProdMap where B_CBCode3  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, B_MLCode3 , 'WWML1', null from tools.dbo.Com_ProdMap where B_MLCode3  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, B_MLCode4 , 'WWML2', null from tools.dbo.Com_ProdMap where B_MLCode4  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, B_MLCode6 , 'WWML3', null from tools.dbo.Com_ProdMap where B_MLCode6  is not null  and A_CG > 4" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, B_PBCode3 , 'WWPB', null from tools.dbo.Com_ProdMap where B_PBCode3  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_Code , 'WWW', null from tools.dbo.Com_ProdMap where W_Code  is not null " & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_Code2 , 'WWWNSW', null from tools.dbo.Com_ProdMap where W_Code2  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_Code3 , 'WWWVIC', null from tools.dbo.Com_ProdMap where W_Code3  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_Code4 , 'WWWQLD', null from tools.dbo.Com_ProdMap where W_Code4  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_Code5 , 'WWWSA', null from tools.dbo.Com_ProdMap where W_Code5  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_Code6 , 'WWWWA', null from tools.dbo.Com_ProdMap where W_Code6  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_HBCode , 'WWHB', null from tools.dbo.Com_ProdMap where W_HBCode  is not null and A_CG <> 58 " & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_HBCode2 , 'WWHB1', null from tools.dbo.Com_ProdMap where W_HBCode2  is not null and A_CG <> 58 " & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_HBCode3 , 'WWHB2', null from tools.dbo.Com_ProdMap where W_HBCode3  is not null and A_CG <> 58 " & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_HBCode , 'WWWNAT1', null from tools.dbo.Com_ProdMap where W_HBCode  is not null and A_CG = 58 " & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_HBCode2 , 'WWWNAT2', null from tools.dbo.Com_ProdMap where W_HBCode2  is not null and A_CG = 58 " & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_HBCode3 , 'WWWNAT3', null from tools.dbo.Com_ProdMap where W_HBCode3  is not null and A_CG = 58 " & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_SelCode , 'WWSelect', null from tools.dbo.Com_ProdMap where W_SelCode  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_SelCode2 , 'WWSelect1', null from tools.dbo.Com_ProdMap where W_SelCode2  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_SelCode3 , 'WWSelect2', null from tools.dbo.Com_ProdMap where W_SelCode3  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_SelCode4 , 'WWSelect3', null from tools.dbo.Com_ProdMap where W_SelCode4  is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_SelCode5 , 'WWSelect4', null from tools.dbo.Com_ProdMap where W_SelCode5  is not null" & Chr(10)
        If Bypass <> "" Then
            strSQL = strSQL & ") as wd inner join #Prods p on p.A_Code = wd.A_Code union select distinct dmd.A_Code,dmd.A_CG, dmd.A_SCG, dmd.CCode, dmd.Matchtype,dmd.cpack from (" & Chr(10)
        Else
            strSQL = strSQL & ") as wd union select distinct dmd.A_Code,dmd.A_CG, dmd.A_SCG, dmd.CCode, dmd.Matchtype,dmd.cpack from (" & Chr(10)
        End If
        strSQL = strSQL & "select A_Code, A_CG, A_SCG,  B_CBCode2 as Ccode, B_CBCode2Pack as Cpack, 'DMCB' as Matchtype from tools.dbo.Com_ProdMap where B_CBCode2 is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG,  DM_Code1 as Ccode, DM_Code1Pack, 'DM1' as Matchtype from tools.dbo.Com_ProdMap where DM_Code1 is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG,  DM_Code2 as Ccode, DM_Code2Pack, 'DM2' as Matchtype from tools.dbo.Com_ProdMap where DM_Code2 is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG,  B_PBCode2 as Ccode, B_PBCode2Pack, 'DMPB' as Matchtype from tools.dbo.Com_ProdMap where B_PBCode2 is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG,  B_MLCode5 as Ccode, B_MLCode5Pack, 'DMML1' as Matchtype from tools.dbo.Com_ProdMap where B_MLCode5 is not null and B_MLCode5Pack is not null" & Chr(10)
        strSQL = strSQL & "union select A_Code, A_CG, A_SCG,  B_MLCode6 as Ccode, B_MLCode6Pack, 'DMML2' as Matchtype from tools.dbo.Com_ProdMap where B_MLCode6 is not null and B_MLCode6Pack is not null) as dmd "
        If Bypass <> "" Then strSQL = strSQL & "inner join #Prods p on p.A_Code = dmd.A_Code" & Chr(10)
        strSQL = strSQL & ") a" & Chr(10) & "where a.CCode <> 'Unassigned'" & Chr(10)
        If Bypass <> "" Then
        ElseIf cCG <> 0 Then
            strSQL = strSQL & "and A_CG in (" & cCG & ")" & Chr(10)
            If scg <> "0" Then strSQL = strSQL & "and A_SCG = " & scg & Chr(10)
        End If
        strSQL = strSQL & "order by a.A_Code" & Chr(10)
        If Bypass <> "" Then strSQL = strSQL & "drop table #Prods" & Chr(10)
    '  '  Debug.Print strSQL
    
    End If
    
    If SQLType = "CBIS_POSQTY" Then
            CBA_DBtoQuery = 599
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 3" & Chr(10)
            strSQL = strSQL & "declare @datefr as date" & Chr(10)
            strSQL = strSQL & "declare @dateto as date" & Chr(10)
            strSQL = strSQL & "set @dateto = '" & Format(DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "set @datefr = '" & Format(DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "set @datefr = (select case when datepart(WEEKDAY,@datefr) = 1 then @datefr else dateadd(DD,7-(datepart(WEEKDAY,@datefr)),@datefr) end )" & Chr(10)
            strSQL = strSQL & "set @dateto = (select case when datepart(WEEKDAY,@dateto) = 1 then @dateto else dateadd(DD,7-(datepart(WEEKDAY,@dateto)),@dateto) end )" & Chr(10)
            strSQL = strSQL & "select isnull(p.con_productcode, pos.productcode)" & Chr(10)
            strSQL = strSQL & ",convert(nvarchar(25),datepart(WEEK,posdate))" & Chr(10)
            strSQL = strSQL & ",sum(Pos.Quantity) as POSQTY" & Chr(10)
            strSQL = strSQL & ", sum(ISNULL((NULLIF(pos.retail,0)/NULLIF(Pos.Quantity,0)/isnull(nullif(dr.retail,0),NULLIF(pos.retail,0)/NULLIF(Pos.Quantity,0)))*nullif(Pos.Quantity,0),0)) as POSCALCQTY, sum(pos.retail) as POSretail--, sum(retailnet) as RCVretail" & Chr(10)
            strSQL = strSQL & "from cbis599p.dbo.pos pos" & Chr(10)
            strSQL = strSQL & "inner join cbis599p.dbo.product p on p.productcode = pos.productcode" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.divretail dr on dr.divno = pos.divno and dr.validfrom <= pos.posdate and isnull(dr.validto,getdate()) >= pos.posdate and dr.productcode = pos.productcode" & Chr(10)
            strSQL = strSQL & "where pos.posdate >= @datefr and pos.posdate < @dateto" & Chr(10)
            strSQL = strSQL & "and isnull(p.con_productcode, pos.productcode) in (" & Bypass & ")" & Chr(10)
            If UCase(scg) = "NSW" Then strSQL = strSQL & "and pos.divno in (501,504)" & Chr(10)
            If UCase(scg) = "VIC" Then strSQL = strSQL & "and pos.divno in (502,505)" & Chr(10)
            If UCase(scg) = "QLD" Then strSQL = strSQL & "and pos.divno in (503,506)" & Chr(10)
            If UCase(scg) = "SA" Then strSQL = strSQL & "and pos.divno in (507)" & Chr(10)
            If UCase(scg) = "WA" Then strSQL = strSQL & "and pos.divno in (509)" & Chr(10)
            strSQL = strSQL & "group by isnull(p.con_productcode, pos.productcode), convert(nvarchar(25),datepart(WEEK,posdate))" & Chr(10)
            strSQL = strSQL & "order by isnull(p.con_productcode, pos.productcode), convert(nvarchar(25),datepart(WEEK,posdate))" & Chr(10)
            'Debug.Print strSQL
    End If
    
    
    If SQLType = "GBDList" Then strSQL = "select firstname + ' ' + name , empno from cbis599p.dbo.employee where positionno in (43) and isinactive = 0 order by firstname"
    If SQLType = "BDList" Then strSQL = "select firstname + ' ' + name , empno from cbis599p.dbo.employee where positionno in (5) and isinactive = 0 order by firstname"
    If SQLType = "CBIS_BuyerInfo" Then strSQL = "select firstname + ' ' + name from cbis599p.dbo.employee e left join cbis599p.dbo.product p on e.empno = p.empno where p.productcode = " & Bypass
    If SQLType = "CGList" Then strSQL = "select CGNo + ' - ' + Description as CG from cbis599p.dbo.commoditygroup"
    If SQLType = "SCGList" Then strSQL = "select CGno, SCGNo + ' - ' + Description as SCG from cbis599p.dbo.subcommoditygroup"
    If SQLType = "CBIS_Pclass" Then strSQL = "select Productclass from cbis599p.dbo.product where productcode = " & Bypass
    If SQLType = "COM_PDesc" Then
        If cType = "CO" Then strSQL = "select top 1 brand + ' ' + name from tools.dbo.com_c_prod where colesproductid = '" & Bypass & "' order by DateScraped desc"
        If cType = "WW" Then strSQL = "select top 1 brand + ' ' + name from tools.dbo.com_w_prod where stockcode = '" & Bypass & "' order by DateScraped desc"
        If cType = "DM" Then strSQL = "select top 1 brand + ' ' + name from tools.dbo.com_dm_prod where productid = '" & Bypass & "' order by DateScraped desc"
        If cType = "FC" Then strSQL = "select top 1 brand + ' ' + description from tools.dbo.com_fc_prod where productid = '" & Bypass & "' order by DateScraped desc"
    End If
    
    
    If SQLType = "COM_2ScrapeDates" Then
    CBA_DBtoQuery = 3
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
            strSQL = strSQL & "select 'C', datescraped, num from (select datescraped, num, row_number() over (order by datescraped desc) as row from (" & Chr(10)
            strSQL = strSQL & "select datescraped, count(datescraped) as num from tools.dbo.com_c_prod group by DateScraped) a) b where row < 57" & Chr(10)
            strSQL = strSQL & "union select 'WW', datescraped, num from (select datescraped, num, row_number() over (order by datescraped desc) as row from (" & Chr(10)
            strSQL = strSQL & "select datescraped, count(datescraped) as num from tools.dbo.com_w_prod group by DateScraped) a) b where row < 57" & Chr(10)
            strSQL = strSQL & "union select 'DM', datescraped, num from (select datescraped, num, row_number() over (order by datescraped desc) as row from (" & Chr(10)
            strSQL = strSQL & "select datescraped, count(datescraped) as num from tools.dbo.com_dm_prod group by DateScraped) a) b where row < 9" & Chr(10)
            strSQL = strSQL & "union select 'FC', datescraped, num from (select datescraped, num, row_number() over (order by datescraped desc) as row from (" & Chr(10)
            strSQL = strSQL & "select datescraped, count(datescraped) as num from tools.dbo.com_fc_prod group by DateScraped) a) b where row < 9" & Chr(10)
    End If
    
    If SQLType = "CBIS_ProdbyEmp" Then
            CBA_DBtoQuery = 599
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
            strSQL = strSQL & "select  p.productcode, p.description, p.cgno, cg.description as cgdesc,p.scgno,scg.description as scgdesc, p.seasonid,s.description as seasondesc, p.productclass,pc.description as productclassdesc, p.empno, emp.firstname + ' ' +  emp.name as Buyer, gbd.firstname + ' ' + gbd.name as GBD" & Chr(10)
            strSQL = strSQL & "from cbis599p.dbo.product p left join cbis599p.dbo.season s on s.seasonid = p.seasonid" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.commoditygroup cg on cg.cgno = p.cgno" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.subcommoditygroup scg on scg.cgno = p.cgno and scg.scgno = p.scgno" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.productclass pc on pc.productclass = p.productclass" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.employee emp on emp.empno = p.empno" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.employee gbd on emp.empno_grp = gbd.empno" & Chr(10)
            strSQL = strSQL & "where p.productclass in (1,4)  order by productcode" & Chr(10)
    
    End If
    If SQLType = "CBIS_ProdbyEmpActive" Then
            CBA_DBtoQuery = 599
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
            strSQL = strSQL & "select  p.productcode, p.description, p.cgno, cg.description as cgdesc,p.scgno,scg.description as scgdesc, p.seasonid,s.description as seasondesc, p.productclass,pc.description as productclassdesc, p.empno, case when p.empno = 999999 then 'Produce' else emp.firstname + ' ' +  emp.name end as Buyer, gbd.firstname + ' ' + gbd.name as GBD" & Chr(10)
            strSQL = strSQL & "from cbis599p.dbo.product p left join cbis599p.dbo.season s on s.seasonid = p.seasonid" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.commoditygroup cg on cg.cgno = p.cgno" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.subcommoditygroup scg on scg.cgno = p.cgno and scg.scgno = p.scgno" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.productclass pc on pc.productclass = p.productclass" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.employee emp on emp.empno = p.empno" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.employee gbd on emp.empno_grp = gbd.empno" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.contract c on c.ProductCode = p.productcode" & Chr(10)
            strSQL = strSQL & "where p.productclass in (1,4) and isnull(c.deliveryto, getdate()) >= getdate() and p.cgno <> 61 " & Chr(10)
            strSQL = strSQL & "group by p.productcode, p.description, p.cgno, cg.description ,p.scgno,scg.description , p.seasonid,s.description , p.productclass,pc.description , p.empno, emp.firstname + ' ' +  emp.name , gbd.firstname + ' ' + gbd.name" & Chr(10)
            strSQL = strSQL & "order by productcode" & Chr(10)
  
    
    End If
    
    If SQLType = "CBIS_ProductShare" Then
    
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
            strSQL = strSQL & "declare @datefrom as date" & Chr(10)
            strSQL = strSQL & "declare @dateto as date" & Chr(10)
            strSQL = strSQL & "declare @totbus as decimal(18,4)" & Chr(10)
            strSQL = strSQL & "set @dateto = '" & Format(DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "set @datefrom = '" & Format(DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "set @totbus = (select sum(retail) from cbis599p.dbo.pos pos inner join cbis599p.dbo.product p on p.productcode = pos.productcode" & Chr(10)
            strSQL = strSQL & "where p.productclass in (1,4) and pos.posdate >= @datefrom and pos.posdate <= @dateto" & Chr(10)
            If sStates = "NSW" Then strSQL = strSQL & "and pos.divno in (501,504)" & Chr(10)
            If sStates = "VIC" Then strSQL = strSQL & "and pos.divno in (502,505)" & Chr(10)
            If sStates = "QLD" Then strSQL = strSQL & "and pos.divno in (503,506)" & Chr(10)
            If sStates = "SA" Then strSQL = strSQL & "and pos.divno in (507)" & Chr(10)
            If sStates = "WA" Then strSQL = strSQL & "and pos.divno in (509)" & Chr(10)
            strSQL = strSQL & ")" & Chr(10)
            strSQL = strSQL & "select pos.productcode,p.con_productcode, sum(retail) / @totbus as share" & Chr(10)
            strSQL = strSQL & "from cbis599p.dbo.pos pos" & Chr(10)
            strSQL = strSQL & "inner join cbis599p.dbo.product p on p.productcode = pos.productcode" & Chr(10)
            strSQL = strSQL & "where p.productclass in (1,4)" & Chr(10)
            strSQL = strSQL & "and pos.posdate >= @datefrom and pos.posdate <= @dateto" & Chr(10)
            If sStates = "NSW" Then strSQL = strSQL & "and pos.divno in (501,504)" & Chr(10)
            If sStates = "VIC" Then strSQL = strSQL & "and pos.divno in (502,505)" & Chr(10)
            If sStates = "QLD" Then strSQL = strSQL & "and pos.divno in (503,506)" & Chr(10)
            If sStates = "SA" Then strSQL = strSQL & "and pos.divno in (507)" & Chr(10)
            If sStates = "WA" Then strSQL = strSQL & "and pos.divno in (509)" & Chr(10)
            If strProds <> "" Then strSQL = strSQL & "and productcode = 1035"
            strSQL = strSQL & "group by pos.productcode, p.con_productcode" & Chr(10)
            strSQL = strSQL & "order by pos.productcode, p.con_productcode" & Chr(10)
            'Debug.Print strSQL
    
    End If
    
    If SQLType = "CBIS_Retails" Then
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
            strSQL = strSQL & "declare @datefrom as date" & Chr(10)
            strSQL = strSQL & "declare @dateto as date" & Chr(10)
            strSQL = strSQL & "declare @totbus as decimal(18,4)" & Chr(10)
            strSQL = strSQL & "set @dateto = '" & Format(DateTo, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "set @datefrom = '" & Format(DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "select productcode, validfrom, validto, retail from (" & Chr(10)
            strSQL = strSQL & "select productcode, convert(date,validfrom) as validfrom, convert(date,isnull(validto,getdate())) as validto,retail" & Chr(10)
            strSQL = strSQL & ",row_number() over (Partition by productcode, validfrom order by retail) as row" & Chr(10)
            strSQL = strSQL & "from cbis599p.dbo.divretail" & Chr(10)
            strSQL = strSQL & "where validfrom <= @dateto" & Chr(10)
            strSQL = strSQL & "and isnull(validto,getdate()) >= @datefrom" & Chr(10)
            If sStates = "NSW" Then strSQL = strSQL & "and pos.divno in (501,504)" & Chr(10)
            If sStates = "VIC" Then strSQL = strSQL & "and pos.divno in (502,505)" & Chr(10)
            If sStates = "QLD" Then strSQL = strSQL & "and pos.divno in (503,506)" & Chr(10)
            If sStates = "SA" Then strSQL = strSQL & "and pos.divno in (507)" & Chr(10)
            If sStates = "WA" Then strSQL = strSQL & "and pos.divno in (509)" & Chr(10)
            If strProds <> "" Then strSQL = strSQL & "and productcode = " & strProds & Chr(10)
            strSQL = strSQL & ") a where a.row = 1" & Chr(10)
            strSQL = strSQL & "order by productcode" & Chr(10)
    End If
    
    
    
    If SQLType = "Chart_Coles" Then
            CBA_DBtoQuery = 3
            strSQL = "declare @PROD as nvarchar(20)" & Chr(10)
            strSQL = strSQL & "set @PROD = '" & cType & "'" & Chr(10)
            strSQL = strSQL & "select distinct cp.CCode,cp.pack,cp.Datescraped, cp.price , cp.pricesaving, cp.price + pricesaving as unpromotedprice, cp.state" & Chr(10)
            strSQL = strSQL & "from (" & Chr(10)
            strSQL = strSQL & "select colesproductid as CCode, Datescraped, packsize as pack, Price, priceprevious as previousprice, pricesaving ," & Chr(10)
            strSQL = strSQL & "row_number() over (partition by datescraped order by price) as row" & Chr(10)
            strSQL = strSQL & ",case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = 'NSW' then 'NSW' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = 'VIC' then 'VIC' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = 'QLD' then 'QLD' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,2) = 'SA' then 'SA' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,2) = 'WA' then 'WA' else 'Unknown'" & Chr(10)
            strSQL = strSQL & "end end end end end as 'State'" & Chr(10)
            strSQL = strSQL & "from tools.dbo.com_c_prod" & Chr(10)
            strSQL = strSQL & "where colesproductid = @PROD" & Chr(10)
            If LCase(Bypass) = "national" Then Else strSQL = strSQL & "and substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = '" & Bypass & "'" & Chr(10)
            strSQL = strSQL & ") as cp" & Chr(10)
            strSQL = strSQL & "where cp.Price > 0" & Chr(10)
            strSQL = strSQL & "and cp.datescraped > dateadd(ww,-12,getdate())" & Chr(10)
            strSQL = strSQL & "and cp.row = 1" & Chr(10)
            strSQL = strSQL & "group by cp.CCode,cp.Datescraped,cp.price, cp.pack,cp.previousprice, cp.pricesaving, cp.state" & Chr(10)
            strSQL = strSQL & "order by pack, datescraped" & Chr(10)
            'Debug.Print strSQL
    End If
    
    If SQLType = "Chart_WW" Then
            CBA_DBtoQuery = 3
            strSQL = "declare @PROD as nvarchar(20)" & Chr(10)
            strSQL = strSQL & "set @PROD = '" & cType & "'" & Chr(10)
            strSQL = strSQL & "select distinct cp.CCode,cp.pack,cp.Datescraped, cp.price , cp.pricesaving, cp.price + pricesaving as unpromotedprice, cp.state" & Chr(10)
            strSQL = strSQL & "from (" & Chr(10)
            strSQL = strSQL & "select stockcode as CCode, Datescraped, packagesize as pack, Price, wasprice as previousprice, savingsamount as pricesaving ," & Chr(10)
            strSQL = strSQL & "row_number() over (partition by datescraped order by price) as row," & Chr(10)
            strSQL = strSQL & "case when convert(int,substring(convert(nvarchar(max), s.Postcode),1,1))in (1,2) then 'NSW' else" & Chr(10)
            strSQL = strSQL & "case when convert(int,substring(convert(nvarchar(max), s.Postcode),1,1))in (3,8) then 'VIC' else" & Chr(10)
            strSQL = strSQL & "case when convert(int,substring(convert(nvarchar(max), s.Postcode),1,1))in (4,9) then 'QLD' else" & Chr(10)
            strSQL = strSQL & "case when convert(int,substring(convert(nvarchar(max), s.Postcode),1,1))in (5)  then 'SA' else" & Chr(10)
            strSQL = strSQL & "case when convert(int,substring(convert(nvarchar(max), s.Postcode),1,1))in (6) then 'WA'" & Chr(10)
            strSQL = strSQL & "end end end end end  as 'State'" & Chr(10)
            strSQL = strSQL & "from tools.dbo.com_w_prod as p" & Chr(10)
            strSQL = strSQL & "left join tools.dbo.com_w_stores as s on s.addressid = p.addressid" & Chr(10)
            strSQL = strSQL & "where stockcode = @PROD" & Chr(10)
            If LCase(Bypass) <> "national" Then
                If sStateLook = "NSW" Then
                    strSQL = strSQL & "and convert(int,substring(convert(nvarchar(max), s.Postcode),1,1))in (1,2)" & Chr(10)
                ElseIf sStateLook = "QLD" Then
                    strSQL = strSQL & "and convert(int,substring(convert(nvarchar(max), s.Postcode),1,1))in (4,9)" & Chr(10)
                ElseIf sStateLook = "VIC" Then
                    strSQL = strSQL & "and convert(int,substring(convert(nvarchar(max), s.Postcode),1,1))in (3,8)" & Chr(10)
                ElseIf sStateLook = "SA" Then
                    strSQL = strSQL & "and convert(int,substring(convert(nvarchar(max), s.Postcode),1,1))in (5)" & Chr(10)
                ElseIf sStateLook = "WA" Then
                    strSQL = strSQL & "and convert(int,substring(convert(nvarchar(max), s.Postcode),1,1))in (6)" & Chr(10)
                End If
            End If
            strSQL = strSQL & ") as cp" & Chr(10)
            strSQL = strSQL & "where cp.Price > 0 And cp.datescraped > DateAdd(ww, -12, getdate()) And cp.Row = 1" & Chr(10)
            strSQL = strSQL & "group by cp.CCode,cp.Datescraped,cp.price, cp.pack,cp.previousprice, cp.pricesaving, cp.state" & Chr(10)
            strSQL = strSQL & "order by pack, datescraped" & Chr(10)
            'Debug.Print strSQL
    End If
    
    
    
    If SQLType = "Chart_DM" Then
            CBA_DBtoQuery = 3
            strSQL = "declare @PROD as nvarchar(20)" & Chr(10)
            strSQL = strSQL & "set @PROD = '" & cType & "'" & Chr(10)
            strSQL = strSQL & "select  c.CCode, c.pack, a.pdate,c.price, 0 as pricesaving, 0 as unpromotedprice, state" & Chr(10)
            strSQL = strSQL & "from (SELECT  TOP (DATEDIFF(DAY, dateadd(ww,-12,getdate()), getdate()) + 1)" & Chr(10)
            strSQL = strSQL & "Pdate = convert(date,DATEADD(DAY, ROW_NUMBER() OVER(ORDER BY a.object_id) - 1, dateadd(ww,-12,getdate())))" & Chr(10)
            strSQL = strSQL & "FROM    sys.all_objects a CROSS JOIN sys.all_objects b) as a" & Chr(10)
            strSQL = strSQL & "left join (" & Chr(10)
            strSQL = strSQL & "select CCode, Datescraped, pack, price, state," & Chr(10)
            strSQL = strSQL & "row_number() over (partition by datescraped, pack order by price) as row" & Chr(10)
            strSQL = strSQL & "from (" & Chr(10)
            strSQL = strSQL & "select productid as CCode, convert(date,Datescraped) as Datescraped, sizedescription1 as pack, Price1 as price, case when s.storescan_storestate = 'ACT' then 'NSW' else s.storescan_storestate end as state from tools.dbo.com_dm_prod" & Chr(10)
            strSQL = strSQL & "left join tools.dbo.com_dm_stores as s on t0_Stores_storescan_storeNo = s.storescan_storeNo" & Chr(10)
            strSQL = strSQL & "where productid = @PROD and datescraped >= dateadd(ww,-12,getdate())" & Chr(10)
            strSQL = strSQL & "Union" & Chr(10)
            strSQL = strSQL & "select productid as CCode, convert(date,Datescraped) as Datescraped, sizedescription2 as pack, Price2 as price, case when s.storescan_storestate = 'ACT' then 'NSW' else s.storescan_storestate end as state from tools.dbo.com_dm_prod" & Chr(10)
            strSQL = strSQL & "left join tools.dbo.com_dm_stores as s on t0_Stores_storescan_storeNo = s.storescan_storeNo" & Chr(10)
            strSQL = strSQL & "where productid = @PROD and datescraped >= dateadd(ww,-12,getdate())" & Chr(10)
            strSQL = strSQL & "Union" & Chr(10)
            strSQL = strSQL & "select productid as CCode, convert(date,Datescraped) as Datescraped, sizedescription3 as pack, Price3 as price, case when s.storescan_storestate = 'ACT' then 'NSW' else s.storescan_storestate end as state from tools.dbo.com_dm_prod" & Chr(10)
            strSQL = strSQL & "left join tools.dbo.com_dm_stores as s on t0_Stores_storescan_storeNo = s.storescan_storeNo" & Chr(10)
            strSQL = strSQL & "where productid = @PROD and datescraped >= dateadd(ww,-12,getdate())" & Chr(10)
            strSQL = strSQL & "Union" & Chr(10)
            strSQL = strSQL & "select productid as CCode, convert(date,Datescraped) as Datescraped, specialsizedescription1 as pack, specialPrice1 as price, case when s.storescan_storestate = 'ACT' then 'NSW' else s.storescan_storestate end as state from tools.dbo.com_dm_prod" & Chr(10)
            strSQL = strSQL & "left join tools.dbo.com_dm_stores as s on t0_Stores_storescan_storeNo = s.storescan_storeNo" & Chr(10)
            strSQL = strSQL & "where productid = @PROD and datescraped >= dateadd(ww,-12,getdate())" & Chr(10)
            strSQL = strSQL & ") as d" & Chr(10)
            strSQL = strSQL & "where d.CCode is not null and d.pack <> ''" & Chr(10)
            If LCase(Bypass) <> "national" Then strSQL = strSQL & "and state = '" & Bypass & "'" & Chr(10)
            strSQL = strSQL & ") as c on c.datescraped = a.Pdate" & Chr(10)
            strSQL = strSQL & "where Row = 1" & Chr(10)
            strSQL = strSQL & "and pack = '" & scg & "'" & Chr(10)
            strSQL = strSQL & "order by pack,pdate" & Chr(10)
            'Debug.Print strSQL
    End If
    
    If SQLType = "CBAR_CompSKUCount" Then
            CBA_DBtoQuery = 3
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
            strSQL = strSQL & "DECLARE @SDATE Date = '" & Format(DateSerial(Year(DateFrom), Month(DateFrom), 1), "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "DECLARE @EDATE Date = dateadd(M,1,@SDATE)" & Chr(10)
            strSQL = strSQL & "select b.datechanged,b.aldiprod, b.comppcode, b.comptype, b.Cgno" & Chr(10)
            strSQL = strSQL & ", row_number() over (Partition by aldiprod, comptype order by datechanged) as row" & Chr(10)
            strSQL = strSQL & ", count(comppcode) over (Partition by aldiprod, comptype) as cnt" & Chr(10)
            strSQL = strSQL & "into #ChangeComps from ( select a.datechanged,a.aldiprod, a.comppcode, a.comptype, a.Cgno" & Chr(10)
            strSQL = strSQL & "from (select datechanged, aldiprod, comppcode, comptype, pm.A_CG as CGno" & Chr(10)
            strSQL = strSQL & ", rank() over (partition by aldiprod, comptype, convert(date, datechanged) order by datechanged desc)  as row" & Chr(10)
            strSQL = strSQL & "from tools.dbo.com_mapchange mc left join tools.dbo.com_prodmap pm on pm.A_Code = mc.Aldiprod" & Chr(10)
            strSQL = strSQL & "" & Chr(10)
            strSQL = strSQL & "where aldiprod in (1034 , 1035 , 1304 , 1332 , 1469 , 1486 , 1507 , 1534 , 2206 , 2277 , 2390 , 2534 , 2541 , 2856 , 3245 , 3372 , 3746 , 3816 , 3910 , 3991 , 3994 , 4030 , 4163 , 4284 , 4577 , 4679 , 5262 , 5492 , 5581 , 5618 , 5658 , 5785 , 5825 , 5845 , 5925 , 6102 , 6305 , 6667 , 6793 , 7355 , 7750 , 8554 , 8575 , 8672 , 9018 , 9035 , 9156 , 9268 , 9314 , 9332 , 9336 , 9372 , 9635 , 9756 , 9933 , 35181 , 36145 , 36480 , 36783 , 37894 , 38066 , 38090 , 38093 , 38916 , 38992 , 39374 , 39812 , 39837 , 40715 , 41550 , 41756 , 41764 , 42028 , 42191 , 42247 , 42335 , 42534 , 42802 , 42935 , 43075 , 43229 , 44438 , 44834 , 44835 , 44951 , 46020 , 46190 , 46310 , 46711 , 46890 , 47086 , 47271 , 47451 , 47699 , 47951 , 48162 , 48999 , 49052 , 49089 , 49115 , 49598 , 49686 , 50180 , 50208 , 50232 , 51545 , 51629 , 51769 , 51854 , 52186 , 52495 , 52620 , 52726 , 53046 , 53078 , 53313 , 54002 , 54309 , 54850" & Chr(10)
            strSQL = strSQL & ", 55175 , 55187 , 56237 , 57246 , 57373 , 57406 , 57777 , 58052 , 58515 , 58599 , 58791 , 59123 , 60001 , 61040" & Chr(10)
            strSQL = strSQL & ", 76014 , 76016 , 76081 , 76082 , 76083 , 76084 , 76085 , 76101 , 76111 , 76112 , 76128 , 76133 , 76134 , 76139 , 76673 , 76679 , 76798 , 77035 , 77036 , 77040 , 77051 , 77052 , 77053 , 77117 , 77118 , 77119 , 77166 , 77254 , 77255 , 77256 , 77257 , 77258 , 77263 , 77281 , 77305 , 77306 , 77307 , 77384 , 77385 , 77595 , 77613)" & Chr(10)
            strSQL = strSQL & ") a where a.row = 1 and convert(date, datechanged) <= dateadd(D,-1,@EDATE)) b" & Chr(10)
            strSQL = strSQL & "order by aldiprod,comptype, datechanged desc" & Chr(10)
            strSQL = strSQL & "            " & Chr(10)
            strSQL = strSQL & "select CCode into #PRODS" & Chr(10)
            strSQL = strSQL & "from (select Aldiprod, comptype,case when row = 1 then @SDATE else" & Chr(10)
            strSQL = strSQL & "case when convert(date,datechanged) < @SDATE then @SDATE else convert(date,datechanged) end end as validfrom" & Chr(10)
            strSQL = strSQL & ",case when row = cnt then dateadd(D,-1,@EDATE)  else" & Chr(10)
            strSQL = strSQL & "(select dateadd(dd,-1,convert(date,datechanged)) from #ChangeComps c where c.comptype = a.comptype and c.aldiprod = a.aldiprod and c.row = a.row + 1)" & Chr(10)
            strSQL = strSQL & "end as validto ,comppcode as CCode, CGno from (select *" & Chr(10)
            strSQL = strSQL & "from #ChangeComps cc" & Chr(10)
            strSQL = strSQL & ") a ) b where  validto > @SDATE and CCode <> 'Unassigned'" & Chr(10)
            strSQL = strSQL & "group by CCode" & Chr(10)
            strSQL = strSQL & "            " & Chr(10)
            strSQL = strSQL & "Select 'Coles' as Chain, colesproductid, p.Brand, p.Name" & Chr(10)
            strSQL = strSQL & ", case when charindex('smart buy',lower(p.Brand)) > 0 then 'Smartbuy' else" & Chr(10)
            strSQL = strSQL & "    case when charindex('coles',lower(p.Brand)) > 0 then 'Coles' else 'Branded' end end as Label" & Chr(10)
            strSQL = strSQL & "into #COLES from tools.dbo.Com_C_Prod p" & Chr(10)
            strSQL = strSQL & "where DateScraped >= @SDATE and DateScraped < @EDATE" & Chr(10)
            strSQL = strSQL & "group by colesproductid, p.Brand, p.Name" & Chr(10)
            strSQL = strSQL & "order by case when charindex('smart buy',lower(p.Brand)) > 0 then 'Smartbuy' else" & Chr(10)
            strSQL = strSQL & "    case when charindex('coles',lower(p.Brand)) > 0 then 'Coles' else 'Branded' end end" & Chr(10)
            strSQL = strSQL & "            " & Chr(10)
            strSQL = strSQL & "Select 'Woolworths' as Chain, Stockcode, p.Brand, p.Name" & Chr(10)
            strSQL = strSQL & ", case when charindex('select',lower(p.Name)) > 0 then 'Woolworths' else" & Chr(10)
            strSQL = strSQL & "    case when charindex('woolworths',lower(p.Name)) > 0 then 'Woolworths' else" & Chr(10)
            strSQL = strSQL & "        case when charindex('essentials',lower(p.Name)) > 0 then 'Woolworths' else" & Chr(10)
            strSQL = strSQL & "            case when charindex('homebrand',lower(p.Name)) > 0 then 'Homebrand' else 'Branded' end end end end as Label" & Chr(10)
            strSQL = strSQL & "into #WW from tools.dbo.Com_W_Prod p" & Chr(10)
            strSQL = strSQL & "where DateScraped >= @SDATE and DateScraped < @EDATE" & Chr(10)
            strSQL = strSQL & "group by stockcode, p.Brand, p.Name" & Chr(10)
            strSQL = strSQL & "order by case when charindex('select',lower(p.Name)) > 0 then 'Woolworths' else" & Chr(10)
            strSQL = strSQL & "    case when charindex('woolworths',lower(p.Name)) > 0 then 'Woolworths' else" & Chr(10)
            strSQL = strSQL & "        case when charindex('essentials',lower(p.Name)) > 0 then 'Woolworths' else" & Chr(10)
            strSQL = strSQL & "            case when charindex('homebrand',lower(p.Name)) > 0 then 'Homebrand' else 'Branded' end end end end" & Chr(10)
            strSQL = strSQL & "            " & Chr(10)
            strSQL = strSQL & "select chain, label, count(label) as SKUs" & Chr(10)
            strSQL = strSQL & "from #COLES c" & Chr(10)
            strSQL = strSQL & "--inner join #PRODS p on p.CCode = c.ColesProductId" & Chr(10)
            strSQL = strSQL & "group by Chain, Label" & Chr(10)
            strSQL = strSQL & "Union" & Chr(10)
            strSQL = strSQL & "select chain, label,count(label) from #WW w" & Chr(10)
            strSQL = strSQL & "--inner join #PRODS p on p.CCode = w.Stockcode" & Chr(10)
            strSQL = strSQL & "group by Chain, Label" & Chr(10)
            strSQL = strSQL & "                        " & Chr(10)
            strSQL = strSQL & "drop table #ChangeComps, #PRODS, #COLES, #WW" & Chr(10)
    
    
    End If
    
    If SQLType = "CBAR_ProdGBDBDEmail" Then
            CBA_DBtoQuery = 599
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
            strSQL = strSQL & "select p.ProductCode, GBD.Empno as GBDEmpno, GBD.Firstname + '.' + GBD.Name  + '@aldi.com.au' as GBDEmail" & Chr(10)
            strSQL = strSQL & ", BD.Empno as BDEmpno, BD.Firstname + '.' + BD.Name  + '@aldi.com.au' as BDEmail" & Chr(10)
            strSQL = strSQL & "from cbis599p.dbo.product p" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.employee BD on BD.empno = p.empno" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.employee GBD on GBD.EmpNo = BD.EmpNo_Grp" & Chr(10)
            strSQL = strSQL & "where BD.PositionNo = 5 And BD.empno <> 999999" & Chr(10)
            strSQL = strSQL & "order by p.productcode" & Chr(10)
    End If
    If SQLType = "CBAR_BDBAEmail" Then
            CBA_DBtoQuery = 599
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
            strSQL = strSQL & "select BD.Empno as BDEmpno, BD.Firstname + '.' + BD.Name  + '@aldi.com.au' as BDEmail" & Chr(10)
            strSQL = strSQL & ", BA.Firstname + ' ' + BA.Name as BA, BA.Firstname + '.' + BA.Name + '@aldi.com.au' as BAEmail" & Chr(10)
            strSQL = strSQL & "from cbis599p.dbo.employee BD" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.employee BA on BA.EmpNo_Grp = BD.EmpNo" & Chr(10)
            strSQL = strSQL & "where BD.PositionNo = 5 And BD.empno <> 999999" & Chr(10)
            strSQL = strSQL & "order by BDEmpno" & Chr(10)
    End If
    
    
    If SQLType = "CBAR_MatchedwDBName" Then
            CBA_DBtoQuery = 3
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
            strSQL = strSQL & "DECLARE @sql_com as nvarchar(Max)" & Chr(10)
            strSQL = strSQL & "DECLARE @listStr Varchar(Max)" & Chr(10)
            strSQL = strSQL & "SELECT @listStr =  COALESCE(@listStr+' ' ,'select A_Code, A_CG, ') + Column_Name + ', ''' + Column_Name + ''' from tools.dbo.com_prodmap where ' + COLUMN_NAME + ' is not null '" & Chr(10)
            If cCG <> 0 Then strSQL = strSQL & "+ ' and A_CG = " & cCG & "'" & Chr(10)
            If cCG <> 0 And scg <> "0" Then strSQL = strSQL & "+ ' and A_SCG = " & scg & "'" & Chr(10)
            'strSQL = strSQL & "+ ' and ' + COLUMN_NAME + ' in (''2463360P'') '" & Chr(10)
            If Bypass > "" Then strSQL = strSQL & "+ ' and A_Code in (" & g_SplitSQLString(Bypass) & ")'" & Chr(10)
            strSQL = strSQL & "+ ' union select A_Code, A_CG, '" & Chr(10)
            strSQL = strSQL & "FROM tools.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'com_prodmap' and substring(Column_Name,1,1) <> 'A'" & Chr(10)
            strSQL = strSQL & "select @sql_com =  substring(@listStr,1,len(@listStr)-26) + ' Order by 3,2'" & Chr(10)
            strSQL = strSQL & "exec(@sql_com)" & Chr(10)
    End If
    
EndSQLType:
    On Error Resume Next
    If cType = "" Then
        Err.Clear
        On Error GoTo Err_Routine
    Else
        curCtype = cType
        On Error GoTo Err_Routine
    End If
    
    
    If strSQL <> "" Then
        If CBA_DBtoQuery = 1 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "ABI_QUERY", CBA_BSA & "LIVE DATABASES\ABI.accdb", CBA_MSAccess, strSQL, 120, , , False)  'Runs CBA_DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 2 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "ABI_QUERY", CBA_BSA & "LIVE DATABASES\ABI.accdb", CBA_MSAccess, strSQL, 120, , , False)  'Runs CBA_DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 3 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "COMRADE_QUERY", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLNCLI10", strSQL, 120, , , False) 'Runs CBA_DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 4 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "COMRADE_QUERY", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLNCLI10", strSQL, 120, , , False) 'Runs CBA_DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery > 500 And CBA_DBtoQuery < 510 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "MMS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", strSQL, 120, , , False)   'Runs CBA_DB_Connection module to create connection to dtabase and run query
        Else
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CBIS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", strSQL, 120, , , False)  'Runs CBA_DB_Connection module to create connection to dtabase and run query
        End If
    End If
    curCtype = ""
    If CBA_DBtoQuery = 1 Then
        If bOutput = False Then
        CBAR_GenPullSQL = False
    '    ElseIf (ABIarr(0, 0) = 0 And bOutput = True) Then
    '    CBAR_GenPullSQL = False
        Else
        CBAR_GenPullSQL = True
        End If
    Else
        If bOutput = False Then
        CBAR_GenPullSQL = False
        'ElseIf (CBISarr(0, 0) = 0 And bOutput = True) Then
        'CBAR_GenPullSQL = False
        Else
        CBAR_GenPullSQL = True
            CBA_DB_Connect.CBA_DB_CBADBUpdate "UPDATE", "CBADB_QUERY", CBA_BSA & "LIVE DATABASES\CBADB.accdb", CBA_MSAccess, _
            "INSERT INTO CBA_QueryLogging ([DateTimeStarted], [DateTimeEnded], [CBA_SQLType], [User], [DBtoQuery])" & Chr(10) & _
            "VALUES ('" & CBA_QueryTimerStart & "' , '" & Now & "', '" & SQLType & "' , '" & CBA_BasicFunctions.CBA_Current_UserID & "', " & CBA_DBtoQuery & ")"
    
        End If
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBAR_GenPullSQL", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function


