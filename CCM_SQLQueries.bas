Attribute VB_Name = "CCM_SQLQueries"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Function CBA_COM_MATCHGenPullSQL(ByRef SQLType, Optional ByRef DateFrom, Optional ByRef DateTo, Optional ByRef cCG As Long, Optional ByRef scg As String, Optional ByRef cType, Optional ByRef Bypass As String)
    'SQLType  - Defines the SQL Query used)
    'DIFFERENT TYPES OF QUERY
    'SHP_QTY            QUANTITY FROM SHIPPING TABLE
    'SHP_RET            RETAIL FROM SHIPPING TABLE
    'SHP_LINE           TOTAL LINES
    'RCV_COST           COST FROM RECIEVING TABLE
    'RCV_RET            RETAIL FROM RECIEVING TABLE
    'RCV_RETN           RETAILNET FROM RECIEVING TABLE
    'PORT_STORE         Number of stores from DateTo less 6 days
    Dim a As Long
    Dim str As String, strSQL As String, assPCode
    Dim cSCG() As Variant, cntSCGs, StartNo, colSCG As Collection, Comp2Find, PCode, curCtype
    Dim lNum As Long, bOutput As Boolean, sStateLook As String, bCurrOnOverWrite As Boolean, strAldiMess As String, GenPullSQL, bOverwriteHappened As Boolean
    Dim strSQLprt01 As String, strSQLprt02 As String, strSQLprt03 As String, strSQLprt04 As String, strSQLprt05 As String, strSQLprt10 As String
    Dim strSQLprt11 As String, strSQLprt12 As String, strSQLprt13 As String, strSQLprt14 As String, strSQLprt20 As String, strSQLprt21 As String, strSQLprt22 As String
    'wks_Data.Unprotect "thepasswordispassword"
    Dim CBA_QueryTimerStart As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    If CBA_DBtoQuery = 0 Then CBA_DBtoQuery = 599
    CBA_QueryTimerStart = Now
    
    DoEvents
    
    If scg <> "" Then
        cntSCGs = 0
        StartNo = 0
        Set colSCG = New Collection
        lNum = InStr(1, scg, ",")
        If lNum = 0 Then
            colSCG.Add scg
            cntSCGs = cntSCGs + 1
        Else
            Do
                If StartNo = 0 Then StartNo = 1
                colSCG.Add Mid(scg, StartNo, lNum - 1)
                cntSCGs = cntSCGs + 1
                StartNo = lNum + 1
                lNum = InStr(StartNo, scg, ",")
            Loop Until lNum = 0
        End If
    Else
    End If
    
    If SQLType = "CCM_MatchMap" Then
        CBA_DBtoQuery = 3
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
        strSQL = strSQL & "DECLARE @sql_com as nvarchar(Max)" & Chr(10)
        strSQL = strSQL & "DECLARE @listStr Varchar(Max)" & Chr(10)
        strSQL = strSQL & "SELECT @listStr = COALESCE(@listStr+',' ,'') + Column_Name" & Chr(10)
        strSQL = strSQL & "FROM tools.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'com_prodmap'" & Chr(10)
        strSQL = strSQL & "and substring(Column_Name,1,1) <> 'A' and substring(Column_Name,1,1) <> 'S'" & Chr(10)
        strSQL = strSQL & "select @sql_com = 'select * from tools.dbo.com_prodmap where coalesce(' + @listStr + ')is not null  order by A_Code'" & Chr(10)
        strSQL = strSQL & "exec(@sql_com)" & Chr(10)

    End If
    
    If SQLType = "getMMData" Then
        CBA_DBtoQuery = 3
        If cType = False Then
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            strSQL = strSQL & "DECLARE @sql_com as nvarchar(Max)" & Chr(10)
            strSQL = strSQL & "DECLARE @listStr Varchar(Max)" & Chr(10)
            strSQL = strSQL & "SELECT @listStr =  COALESCE(@listStr+' ' ,'select ') + 'A_Code, A_CG, A_SCG, ''' + Column_Name + ''' as MatchType,' + Column_Name + ',null as Pack from tools.dbo.com_prodmap where ' + COLUMN_NAME + ' is not null '" & Chr(10)
            If cCG <> 0 Then strSQL = strSQL & "+ ' and A_CG = " & cCG & "'" & Chr(10)
            If scg <> "" And scg <> "0" Then strSQL = strSQL & "+ ' and A_SCG = " & scg & "'" & Chr(10)
            strSQL = strSQL & "+ ' union select'" & Chr(10)
            strSQL = strSQL & "FROM tools.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'com_prodmap'" & Chr(10)
            strSQL = strSQL & "and substring(Column_Name,1,1) <> 'A' and charindex('Pack',Column_Name,1) = 0 and Column_Name <> 'B_AVCode'" & Chr(10)
            strSQL = strSQL & "select @sql_com =  substring(@listStr,1,len(@listStr)-12)  + ' Order by A_Code'" & Chr(10)
            strSQL = strSQL & "exec(@sql_com)" & Chr(10)
        Else
        
        End If
    End If
    
    If SQLType = "isCG" Then
        CBA_DBtoQuery = 599
        strSQL = "SELECT PRD.cgno, prd.scgno FROM CBIS599P.DBO.PRODUCT AS PRD" & Chr(10)
        strSQL = strSQL & "WHERE PRD.PRODUCTCODE = '" & Bypass & "'"
    End If
    
    If SQLType = "FINDPRODDESC" And IsNumeric(cType) Then
        CBA_DBtoQuery = 599
        strSQL = "SELECT PRD.DESCRIPTION, PRD.CON_PRODUCTCODE FROM CBIS599P.DBO.PRODUCT AS PRD" & Chr(10)
        strSQL = strSQL & "WHERE PRD.PRODUCTCODE = '" & cType & "'"
    End If
    
    If SQLType = "getMatchedProducts" Then
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
        strSQL = strSQL & "DECLARE @sql_com as nvarchar(Max)" & Chr(10)
        strSQL = strSQL & "DECLARE @listStr Varchar(Max)" & Chr(10)
        strSQL = strSQL & "SELECT @listStr = COALESCE(@listStr+',' ,'') + Column_Name" & Chr(10)
        strSQL = strSQL & "FROM tools.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'com_prodmap'" & Chr(10)
        strSQL = strSQL & "and substring(Column_Name,1,1) <> 'A'" & Chr(10)
        strSQL = strSQL & "select @sql_com = 'select * from tools.dbo.com_prodmap where coalesce(' + @listStr + ')is not null'" & Chr(10)
        strSQL = strSQL & "exec(@sql_com)" & Chr(10)
    End If
    
    
    If SQLType = "getActiveProdsCGSCGBuyer" Then
        CBA_DBtoQuery = 599
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
        strSQL = strSQL & "select pos.productcode into #PRODUCE from cbis599p.dbo.pos pos" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.product p on p.productcode = pos.ProductCode" & Chr(10)
        strSQL = strSQL & "where p.cgno in (27,58) and pos.posdate >= dateadd(M,-4,getdate())" & Chr(10)
        strSQL = strSQL & "select isnull(con_productcode,p.productcode) as productcode, p.description, p.cgno" & Chr(10)
        strSQL = strSQL & ",cg.Description, isnull(p.scgno,0) ,isnull(scg.Description,''), emp.name--, c.deliveryfrom, c.deliveryto" & Chr(10)
        strSQL = strSQL & "from cbis599p.dbo.PRODUCT p" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.EMPLOYEE emp on emp.empno = p.EmpNo" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.commoditygroup cg on cg.cgno = p.cgno" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.subcommoditygroup scg on scg.cgno = p.cgno and scg.scgno = p.SCGNo" & Chr(10)
        strSQL = strSQL & "inner join cbis599p.dbo.CONTRACT c on c.productcode = p.ProductCode" & Chr(10)
        strSQL = strSQL & "where c.DeliveryFrom <= DateAdd(m, 1, Convert(Date, getdate()))" & Chr(10)
        strSQL = strSQL & "And IsNull(c.Deliveryto, Convert(Date, getdate())) >= Convert(Date, getdate())" & Chr(10)
        strSQL = strSQL & "and p.productclass in (1,4) group by  isnull(con_productcode,p.productcode)" & Chr(10)
        strSQL = strSQL & ", p.description, p.cgno,cg.Description, p.scgno,scg.Description, emp.name" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select isnull(con_productcode,p.productcode) , p.description, p.cgno, cg.Description" & Chr(10)
        strSQL = strSQL & ", isnull(p.scgno,0) ,isnull(scg.Description,''), 'Produce'--, dateadd(M,-4,getdate()), dateadd(W,2,getdate())" & Chr(10)
        strSQL = strSQL & "from #PRODUCE pr" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.product p on p.ProductCode = pr.ProductCode" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.commoditygroup cg on cg.cgno = p.cgno" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.subcommoditygroup scg on scg.cgno = p.cgno and scg.scgno = p.SCGNo" & Chr(10)
        strSQL = strSQL & "group by isnull(con_productcode,p.productcode), p.description, p.cgno" & Chr(10)
        strSQL = strSQL & ", cg.Description, isnull(p.scgno,0) ,isnull(scg.Description,'')" & Chr(10)
        strSQL = strSQL & "drop table #PRODUCE" & Chr(10)
        'Debug.Print strSQL
    End If
    
    If SQLType = "getActiveProdsCGSCGBuyerDetail" Then
        CBA_DBtoQuery = 599
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
        strSQL = strSQL & "select pos.productcode into #PRODUCE from cbis599p.dbo.pos pos" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.product p on p.productcode = pos.ProductCode" & Chr(10)
        strSQL = strSQL & "where p.cgno in (27,58) and pos.posdate >= dateadd(M,-4,getdate())" & Chr(10)
        If Bypass <> "" Then strSQL = strSQL & "and pos.productcode in (" & Bypass & ")" & Chr(10)
        If Bypass = "" And cCG <> 0 Then strSQL = strSQL & "and p.cgno = " & cCG & Chr(10)
        If Bypass = "" And cCG <> 0 And scg <> "0" Then strSQL = strSQL & "and p.scgno = " & scg & Chr(10)
        strSQL = strSQL & "select isnull(con_productcode,p.productcode) as productcode, p.description, p.cgno" & Chr(10)
        strSQL = strSQL & ",cg.Description, isnull(p.scgno,0) ,isnull(scg.Description,''), emp.name--, c.deliveryfrom, c.deliveryto" & Chr(10)
        strSQL = strSQL & "from cbis599p.dbo.PRODUCT p" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.EMPLOYEE emp on emp.empno = p.EmpNo" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.commoditygroup cg on cg.cgno = p.cgno" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.subcommoditygroup scg on scg.cgno = p.cgno and scg.scgno = p.SCGNo" & Chr(10)
        strSQL = strSQL & "inner join cbis599p.dbo.CONTRACT c on c.productcode = p.ProductCode" & Chr(10)
        strSQL = strSQL & "where c.DeliveryFrom <= DateAdd(m, 1, Convert(Date, getdate()))" & Chr(10)
        strSQL = strSQL & "And IsNull(c.Deliveryto, Convert(Date, getdate())) >= Convert(Date, getdate())" & Chr(10)
        strSQL = strSQL & "and p.productclass in (1,4) " & Chr(10)
        If Bypass <> "" Then strSQL = strSQL & "and isnull(con_productcode,p.productcode) in (" & Bypass & ")" & Chr(10)
        If Bypass = "" And cCG <> 0 Then strSQL = strSQL & "and p.cgno = " & cCG & Chr(10)
        If Bypass = "" And cCG <> 0 And scg <> "0" Then strSQL = strSQL & "and p.scgno = " & scg & Chr(10)
        strSQL = strSQL & "group by  isnull(con_productcode,p.productcode)" & Chr(10)
        strSQL = strSQL & ", p.description, p.cgno,cg.Description, p.scgno,scg.Description, emp.name" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select isnull(con_productcode,p.productcode) , p.description, p.cgno, cg.Description" & Chr(10)
        strSQL = strSQL & ", isnull(p.scgno,0) ,isnull(scg.Description,''), 'Produce'--, dateadd(M,-4,getdate()), dateadd(W,2,getdate())" & Chr(10)
        strSQL = strSQL & "from #PRODUCE pr" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.product p on p.ProductCode = pr.ProductCode" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.commoditygroup cg on cg.cgno = p.cgno" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.subcommoditygroup scg on scg.cgno = p.cgno and scg.scgno = p.SCGNo" & Chr(10)
        strSQL = strSQL & "group by isnull(con_productcode,p.productcode), p.description, p.cgno" & Chr(10)
        strSQL = strSQL & ", cg.Description, isnull(p.scgno,0) ,isnull(scg.Description,'')" & Chr(10)
        strSQL = strSQL & "drop table #PRODUCE" & Chr(10)
        'Debug.Print strSQL
    End If
    
    
    If SQLType = "getBuyerlisting" Then
        CBA_DBtoQuery = 599
        strSQL = "select name from cbis599p.dbo.employee where positionno = 05 and isinactive = 0 and EmpNo_Grp is not null"
    End If
    
    
    If SQLType = "getAldiProdsforAllMatched" Then
    
        CBA_DBtoQuery = 3
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
        strSQL = strSQL & "DECLARE @sql_com as nvarchar(Max)" & Chr(10)
        strSQL = strSQL & "DECLARE @listStr Varchar(Max)" & Chr(10)
        strSQL = strSQL & "SELECT @listStr = COALESCE(@listStr+',' ,'') + Column_Name" & Chr(10)
        strSQL = strSQL & "FROM tools.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'com_prodmap'" & Chr(10)
        strSQL = strSQL & "and substring(Column_Name,1,1) <> 'A' " & Chr(10)
        strSQL = strSQL & "select @sql_com = 'select A_Code from tools.dbo.com_prodmap where coalesce(' + @listStr + ') is not null'" & Chr(10)
        strSQL = strSQL & "exec(@sql_com)" & Chr(10)
    
    End If
    
    
    If SQLType = "getAldiProdInfo" Then
        CBA_DBtoQuery = 599
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
        strSQL = strSQL & "select p.productcode as productcode, p.description, p.cgno" & Chr(10)
        strSQL = strSQL & ",cg.Description, isnull(p.scgno,0) ,isnull(scg.Description,''), emp.name--, c.deliveryfrom, c.deliveryto" & Chr(10)
        strSQL = strSQL & "from cbis599p.dbo.PRODUCT p" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.EMPLOYEE emp on emp.empno = p.EmpNo" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.commoditygroup cg on cg.cgno = p.cgno" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.subcommoditygroup scg on scg.cgno = p.cgno and scg.scgno = p.SCGNo" & Chr(10)
        strSQL = strSQL & "inner join cbis599p.dbo.CONTRACT c on c.productcode = p.ProductCode" & Chr(10)
        strSQL = strSQL & "where p.productcode in (" & Bypass & ")" & Chr(10)
        strSQL = strSQL & "group by  p.productcode, p.description, p.cgno,cg.Description, p.scgno,scg.Description, emp.name" & Chr(10)
    
    End If
    
    If SQLType = "getGBDlisting" Then
        CBA_DBtoQuery = 599
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
        strSQL = strSQL & "select Firstname + ' ' + name from cbis599p.dbo.employee where PositionNo = 43"
    End If
    
    
    If SQLType = "COM_ALLCurMatched" Then
    
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_NULLS , ANSI_WARNINGS ON " & Chr(10) & "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10)
    strSQL = strSQL & "select distinct a.A_Code, p.description, a.A_CG as CG, a.A_SCG as SCG, a.matchtype, a.CCode, a.name  from (" & Chr(10)
    strSQL = strSQL & "select cd.A_Code, cd.A_CG, cd.A_SCG, cd.CCode, cd.Matchtype, cp.brand + ' ' + cp.name as name from (" & Chr(10)
    strSQL = strSQL & "select A_Code, A_CG, A_SCG,  B_CBCode as Ccode, 'ColesCB' as Matchtype from tools.dbo.Com_ProdMap where B_CBCode is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code,A_CG, A_SCG,  B_MLCode , 'ColesML1' from tools.dbo.Com_ProdMap where B_MLCode  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code,A_CG, A_SCG,  B_MLCode2 , 'ColesML2' from tools.dbo.Com_ProdMap where B_MLCode2  is not null" & Chr(10)
    If cCG > 4 Then strSQL = strSQL & "union select A_Code,A_CG, A_SCG,  B_MLCode5 , 'ColesML3' from tools.dbo.Com_ProdMap where B_MLCode5  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, B_PBCode , 'ColesPB' from tools.dbo.Com_ProdMap where B_PBCode  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CCode , 'ColesColes' from tools.dbo.Com_ProdMap where C_CCode  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CCode2 , 'ColesColes1' from tools.dbo.Com_ProdMap where C_CCode2  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CCode3 , 'ColesColes2' from tools.dbo.Com_ProdMap where C_CCode3  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CCode4 , 'ColesColes3' from tools.dbo.Com_ProdMap where C_CCode4  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_Code , 'ColesWeb' from tools.dbo.Com_ProdMap where C_Code  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_Code2 , 'ColesWNSW' from tools.dbo.Com_ProdMap where C_Code2  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_Code3 , 'ColesWVIC' from tools.dbo.Com_ProdMap where C_Code3  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_Code4 , 'ColesWQLD' from tools.dbo.Com_ProdMap where C_Code4  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_Code5 , 'ColesWSA' from tools.dbo.Com_ProdMap where C_Code5  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_Code6 , 'ColesWWA' from tools.dbo.Com_ProdMap where C_Code6  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CVCode , 'ColesVal' from tools.dbo.Com_ProdMap where C_CVCode  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CVCode2 , 'ColesVal1' from tools.dbo.Com_ProdMap where C_CVCode2  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_CVCode3 , 'ColesVal2' from tools.dbo.Com_ProdMap where C_CVCode3  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_SBCode , 'ColesSB' from tools.dbo.Com_ProdMap where C_SBCode  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_SBCode2 , 'ColesSB1' from tools.dbo.Com_ProdMap where C_SBCode2  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, C_SBCode3 , 'ColesSB2' from tools.dbo.Com_ProdMap where C_SBCode3  is not null" & Chr(10)
    strSQL = strSQL & ") as cd left join tools.dbo.Com_C_Prod cp on cp.colesproductid = cd.CCode --and cp.datescraped > dateadd(m,-2,getdate())" & Chr(10)
    strSQL = strSQL & "union select wd.A_Code,wd.A_CG, wd.A_SCG, wd.CCode, wd.Matchtype, wp.name from (" & Chr(10)
    strSQL = strSQL & "select A_Code, A_CG, A_SCG,  B_CBCode3 as Ccode, 'WWCB' as Matchtype from tools.dbo.Com_ProdMap where B_CBCode3  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, B_MLCode3 , 'WWML1' from tools.dbo.Com_ProdMap where B_MLCode3  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, B_MLCode4 , 'WWML2' from tools.dbo.Com_ProdMap where B_MLCode4  is not null" & Chr(10)
    If cCG > 4 Then strSQL = strSQL & "union select A_Code, A_CG, A_SCG, B_MLCode6 , 'WWML3' from tools.dbo.Com_ProdMap where B_MLCode6  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, B_PBCode3 , 'WWPB' from tools.dbo.Com_ProdMap where B_PBCode3  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_Code , 'WWWeb' from tools.dbo.Com_ProdMap where W_Code  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_Code2 , 'WWWNSW' from tools.dbo.Com_ProdMap where W_Code2  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_Code3 , 'WWWVIC' from tools.dbo.Com_ProdMap where W_Code3  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_Code4 , 'WWWQLD' from tools.dbo.Com_ProdMap where W_Code4  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_Code5 , 'WWWSA' from tools.dbo.Com_ProdMap where W_Code5  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_Code6 , 'WWWWA' from tools.dbo.Com_ProdMap where W_Code6  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_HBCode , 'WWHB' from tools.dbo.Com_ProdMap where W_HBCode  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_HBCode2 , 'WWHB1' from tools.dbo.Com_ProdMap where W_HBCode2  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_HBCode3 , 'WWHB2' from tools.dbo.Com_ProdMap where W_HBCode3  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_SelCode , 'WWSelect' from tools.dbo.Com_ProdMap where W_SelCode  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_SelCode2 , 'WWSelect1' from tools.dbo.Com_ProdMap where W_SelCode2  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_SelCode3 , 'WWSelect2' from tools.dbo.Com_ProdMap where W_SelCode3  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_SelCode4 , 'WWSelect3' from tools.dbo.Com_ProdMap where W_SelCode4  is not null" & Chr(10)
    strSQL = strSQL & "union select A_Code, A_CG, A_SCG, W_SelCode5 , 'WWSelect4' from tools.dbo.Com_ProdMap where W_SelCode5  is not null" & Chr(10)
    strSQL = strSQL & ") as wd left join tools.dbo.Com_w_Prod wp on wp.stockcode = wd.CCode --and wp.datescraped > dateadd(m,-2,getdate())" & Chr(10)
    If cCG < 5 Then strSQL = strSQL & "union select distinct dmd.A_Code,dmd.A_CG, dmd.A_SCG, dmd.CCode, dmd.Matchtype, dmp.brand + ' ' + dmp.name from (" & Chr(10)
    If cCG < 5 Then strSQL = strSQL & "select A_Code, A_CG, A_SCG,  B_CBCode2 as Ccode, B_CBCode2Pack, 'DMCB' as Matchtype from tools.dbo.Com_ProdMap where B_CBCode2 is not null" & Chr(10)
    If cCG < 5 Then strSQL = strSQL & "union select A_Code, A_CG, A_SCG,  DM_Code1 as Ccode, DM_Code1Pack, 'DM1' as Matchtype from tools.dbo.Com_ProdMap where DM_Code1 is not null" & Chr(10)
    If cCG < 5 Then strSQL = strSQL & "union select A_Code, A_CG, A_SCG,  DM_Code2 as Ccode, DM_Code2Pack, 'DM2' as Matchtype from tools.dbo.Com_ProdMap where DM_Code2 is not null" & Chr(10)
    If cCG < 5 Then strSQL = strSQL & "union select A_Code, A_CG, A_SCG,  B_PBCode2 as Ccode, B_PBCode2Pack, 'DMPB' as Matchtype from tools.dbo.Com_ProdMap where B_PBCode2 is not null" & Chr(10)
    If cCG < 5 Then strSQL = strSQL & "union select A_Code, A_CG, A_SCG,  B_MLCode5 as Ccode, B_MLCode5Pack, 'DMML1' as Matchtype from tools.dbo.Com_ProdMap where B_MLCode5 is not null and B_MLCode5Pack is not null" & Chr(10)
    If cCG < 5 Then strSQL = strSQL & "union select A_Code, A_CG, A_SCG,  B_MLCode6 as Ccode, B_MLCode6Pack, 'DMML2' as Matchtype from tools.dbo.Com_ProdMap where B_MLCode6 is not null and B_MLCode6Pack is not null" & Chr(10)
    If cCG < 5 Then strSQL = strSQL & ") as dmd left join tools.dbo.Com_dm_Prod dmp on dmp.productid = dmd.CCode" & Chr(10)
    strSQL = strSQL & ") a left join cbis599.Cbis599p.dbo.product p on p.productcode = a.a_Code" & Chr(10)
    strSQL = strSQL & "where A_Code in (" & Bypass & ")" & Chr(10)
    strSQL = strSQL & "order by a.A_Code" & Chr(10)
    
    'Debug.Print strSQL
    
    End If
    
    
    
    If SQLType = "COM_Findcurrentlineretail" Then
    CBA_DBtoQuery = 3
                    strSQL = "SET DATEFORMAT dmy" & Chr(10)
                    strSQL = strSQL & "DECLARE @CWDATE as Date" & Chr(10)
                    strSQL = strSQL & "SET @CWDATE = (select max(convert(date, datescraped)) from Tools.dbo.com_c_prod)" & Chr(10)
                    strSQL = strSQL & "select Price from tools.dbo.Com_c_Prod" & Chr(10)
                    strSQL = strSQL & "where Colesproductid = '" & assPCode & "'" & Chr(10)
                    strSQL = strSQL & "and urlscan_Storeseotoken = '" & scg & "'" & Chr(10)
                    strSQL = strSQL & "and datescraped = @CWDATE" & Chr(10)
    '          '  Debug.Print strSQL
    End If
    
    
    If SQLType = "COM_isOnoverwrite" Then
            CBA_DBtoQuery = 3
            strSQL = "select NewRetail, StoreLocation, Oldretail from tools.dbo.Com_RetailChange" & Chr(10)
            strSQL = strSQL & "where CompPcode = '" & assPCode & "'" & Chr(10)
            If scg <> "" Then strSQL = strSQL & "and StoreLocation = '" & scg & "'" & Chr(10)
            strSQL = strSQL & "and validto is null" & Chr(10)
    
          '  Debug.Print strSQL
    
    End If
    
    If SQLType = "COM_CancelOverwrite" Then
                bCurrOnOverWrite = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("COM_isOnoverwrite", , , , scg)
                If bCurrOnOverWrite = True Then
                    CBA_DBtoQuery = 4
                    strSQL = "SET DATEFORMAT dmy" & Chr(10)
                    strSQL = strSQL & "DECLARE @CWDATE as Date" & Chr(10)
                    strSQL = strSQL & "SET @CWDATE = (select max(convert(date, datescraped)) from Tools.dbo.com_c_prod)" & Chr(10)
                    strSQL = strSQL & "Update tools.dbo.com_retailchange" & Chr(10)
                    strSQL = strSQL & "set Validto = @CWDATE" & Chr(10)
                    strSQL = strSQL & "where CompPCode = '" & assPCode & "'" & Chr(10)
                    strSQL = strSQL & "and StoreLocation = '" & scg & "'" & Chr(10)
                    strSQL = strSQL & "Update tools.dbo.com_c_prod" & Chr(10)
                    strSQL = strSQL & "Set Price = " & CBA_COMarr(2, 0) & Chr(10)
                    strSQL = strSQL & "where colesproductid = '" & assPCode & "'" & Chr(10)
                    strSQL = strSQL & "and urlscan_Storeseotoken = '" & scg & "'" & Chr(10)
                    strSQL = strSQL & "and datescraped = @CWDATE" & Chr(10)
    '              '  Debug.Print strSQL
    
                End If
    
    End If
    
    
    
    If SQLType = "COM_WriteOverwrite" Then
            
            bCurrOnOverWrite = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("COM_isOnoverwrite", , , , scg)
            CBA_DBtoQuery = 4
            If bCurrOnOverWrite = False Then
                
                bCurrOnOverWrite = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("COM_Findcurrentlineretail", , , , scg)
                If bCurrOnOverWrite = False Then
                    strAldiMess = "You cannot place an overwrite on this product" & Chr(10) & "There is no data for the Mount Druitt Store"
                    strAldiMess = "When an overwrite is applied, it is applied to the Mount Druitt Store" & Chr(10)
                    strAldiMess = "Typically this issue occurs when you apply an overwite to a new match which has not been matched previously" & Chr(10)
                    strAldiMess = "It is highly likely, now you have applied the match, next Wednesday, the product will avaliable to overwrite" & Chr(10)
                    strAldiMess = "For further information, please contact " & g_Get_Dev_Sts("DevUsers") & Chr(10)
                    MsgBox strAldiMess, vbOKOnly
                    GenPullSQL = False
                    Exit Function
                Else
                    CBA_DBtoQuery = 4
                    strSQL = "SET DATEFORMAT dmy" & Chr(10)
                    strSQL = strSQL & "DECLARE @CWDATE as Date" & Chr(10)
                    strSQL = strSQL & "SET @CWDATE = (select max(convert(date, datescraped)) from Tools.dbo.com_c_prod)" & Chr(10)
                    strSQL = strSQL & "Update tools.dbo.com_c_prod" & Chr(10)
                    strSQL = strSQL & "Set Price = " & cType & Chr(10)
                    strSQL = strSQL & "where colesproductid = '" & assPCode & "'" & Chr(10)
                    strSQL = strSQL & "and urlscan_Storeseotoken = '" & scg & "'" & Chr(10)
                    strSQL = strSQL & "and datescraped = @CWDATE" & Chr(10)
                    strSQL = strSQL & "Insert into tools.dbo.Com_RetailChange (AldiUser, DateChanged, CompType, CompPCode, OldRetail, NewRetail, StoreLocation, Validfrom, Validto)" & Chr(10)
                    strSQL = strSQL & "VALUES ('" & Application.UserName & "',getdate(), 'Coles','" & assPCode & "', '" & CBA_COMarr(0, 0) & "' , " & cType & ", '" & scg & "', @CWDATE, null)" & Chr(10)
                    bOverwriteHappened = True
                End If
            
            Else
            
                strSQL = "SET DATEFORMAT dmy"
                strSQL = strSQL & "DECLARE @CWDATE as Date"
                strSQL = strSQL & "SET @CWDATE = (select max(convert(date, datescraped)) from Tools.dbo.com_c_prod)" & Chr(10)
                strSQL = strSQL & "Update tools.dbo.Com_RetailChange" & Chr(10)
                strSQL = strSQL & "Set Validto = @CWDATE" & Chr(10)
                strSQL = strSQL & "where comppcode = '" & assPCode & "'" & Chr(10)
                strSQL = strSQL & "and CompType = '" & Comp2Find & "'" & Chr(10)
                strSQL = strSQL & "and storelocation = '" & scg & "'" & Chr(10)
                strSQL = strSQL & "and validto is null" & Chr(10)
                strSQL = strSQL & "Update tools.dbo.com_c_prod" & Chr(10)
                strSQL = strSQL & "Set Price = " & cType & Chr(10)
                strSQL = strSQL & "where colesproductid = '" & assPCode & "'" & Chr(10)
                strSQL = strSQL & "and urlscan_Storeseotoken = '" & scg & "'" & Chr(10)
                strSQL = strSQL & "and datescraped = @CWDATE" & Chr(10)
                strSQL = strSQL & "Insert into tools.dbo.Com_RetailChange (AldiUser, DateChanged, CompType, CompPCode, OldRetail, NewRetail, StoreLocation, Validfrom, Validto)" & Chr(10)
                strSQL = strSQL & "VALUES ('" & Application.UserName & "',getdate(),'Coles','" & assPCode & "', " & CBA_COMarr(2, 0) & " , " & cType & ", '" & scg & "', @CWDATE, null)" & Chr(10)
                bOverwriteHappened = True
            End If
            
    '      '  Debug.Print strSQL
    End If
    
    
    If SQLType = "COM_FindCheepestStoreinState" Then
            CBA_DBtoQuery = 3
            strSQL = "SET DATEFORMAT dmy"
            strSQL = strSQL & "DECLARE @CWDATE as Date" & Chr(10)
            strSQL = strSQL & "SET @CWDATE = (select max(convert(date, datescraped)) from Tools.dbo.com_c_prod)" & Chr(10)
            strSQL = strSQL & "select b.UrlScan_StoreSeoToken from (" & Chr(10)
            strSQL = strSQL & "select a.pcode, a.brand, a.name, a.packsize, a.retail, a.Wprice, a.saved, a.cupmeasureunits, a.cupmeasurevalue, a.cupsize,a.UrlScan_StoreSeoToken," & Chr(10)
            strSQL = strSQL & "a.state, row_number() over (partition by a.pcode, a.state order by a.retail) as row from (" & Chr(10)
            strSQL = strSQL & "select distinct colesproductid as 'PCode', isnull(brand,'') as Brand, name, packsize, round(price,2) as Retail, priceprevious as 'WPrice', pricesaving as 'Saved'," & Chr(10)
            strSQL = strSQL & "cupmeasureunits, cupmeasurevalue, cupsize,UrlScan_StoreSeoToken," & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = 'NSW' then 'NSW' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = 'VIC' then 'VIC' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = 'QLD' then 'QLD' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,2) = 'SA' then 'SA' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,2) = 'WA' then 'WA' else 'Unknown'" & Chr(10)
            strSQL = strSQL & "end end end end end as 'State'" & Chr(10)
            strSQL = strSQL & "from Tools.dbo.com_c_prod as CP" & Chr(10)
            strSQL = strSQL & "where Datescraped = @CWDATE" & Chr(10)
            strSQL = strSQL & "and colesproductid = '" & Bypass & "'" & Chr(10)
            strSQL = strSQL & "and UrlScan_StoreSeoToken in (select distinct UrlScan_StoreSeoToken from Tools.dbo.com_c_prod where Datescraped = @CWDATE)" & Chr(10)
            strSQL = strSQL & "group by colesproductid, Brand, name, packsize, priceprevious, pricesaving, price, cupmeasureunits, cupmeasurevalue, cupsize, UrlScan_StoreSeoToken) as a) as b" & Chr(10)
            strSQL = strSQL & "where b.row = 1 and b.state = '" & cType & "'" & Chr(10)
            strSQL = strSQL & "group by b.pcode, b.state, b.brand, b.name, b.packsize, b.retail, b.Wprice, b.cupmeasureunits, b.cupmeasurevalue, b.cupsize, b.saved ,b.UrlScan_StoreSeoToken" & Chr(10)
            
    '      '  Debug.Print strSQL
    
    End If
    
    
    
    
    If SQLType = "COM_FindOverwriteStoreRegionColes" Then
            CBA_DBtoQuery = 3
            strSQL = "select Store from (" & Chr(10)
            strSQL = strSQL & "select urlscan_storeseotoken as Store, Price" & Chr(10)
            strSQL = strSQL & ", row_number() over (partition by '' order by price) as row" & Chr(10)
            strSQL = strSQL & "from dbo.com_c_prod" & Chr(10)
            strSQL = strSQL & "where datescraped = (select max(datescraped) from dbo.com_c_prod )" & Chr(10)
            strSQL = strSQL & "and colesproductid = '" & Bypass & "'" & Chr(10)
            strSQL = strSQL & "and (substring(urlscan_storeseotoken,3,3) = '" & cType & "' or substring(urlscan_storeseotoken,3,2) = '" & cType & "')" & Chr(10)
            strSQL = strSQL & ") a where a.row = 1" & Chr(10)
    
    End If
    
    
    If SQLType = "COM_FindOverwriteStoreColes" Then
            CBA_DBtoQuery = 3
            strSQL = "select top 1 s.urlscan_storeseotoken from tools.dbo.com_c_stores s where storescan_locality = 'MOUNT DRUITT' or storescan_locality = 'ST MARYS'"
    End If
    
    
    If SQLType = "COM_BrandedProdMap" Then
            CBA_DBtoQuery = 3
            strSQL = "select A_Code,B_MLCode, B_CBCode, B_AVCode  from tools.dbo.com_prodmap where A_Code = " & Bypass
    
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
            If Bypass = "National" Then Else strSQL = strSQL & "and substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = '" & Bypass & "'" & Chr(10)
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
            If sStateLook <> "National" Then
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
            If Bypass <> "National" Then strSQL = strSQL & "and state = '" & Bypass & "'" & Chr(10)
            strSQL = strSQL & ") as c on c.datescraped = a.Pdate" & Chr(10)
            strSQL = strSQL & "where Row = 1" & Chr(10)
            strSQL = strSQL & "and pack = '" & scg & "'" & Chr(10)
            strSQL = strSQL & "order by pack,pdate" & Chr(10)
            'Debug.Print strSQL
    End If
    
    If SQLType = "CBIS_RefreshComradeAldiProdListing" Then
            CBA_DBtoQuery = 599
            strSQL = "select distinct DATA.pcode, DATA.Region from (" & Chr(10)
            strSQL = strSQL & "select isnull(p.con_productcode,p.productcode) as pcode," & Chr(10)
            strSQL = strSQL & "case when dc.divno in (501,504) then 'NSW' else" & Chr(10)
            strSQL = strSQL & "case when dc.divno in (503,506) then 'QLD' else" & Chr(10)
            strSQL = strSQL & "case when dc.divno in (502,505) then 'VIC' else" & Chr(10)
            strSQL = strSQL & "case when dc.divno in (509) then 'WA' else" & Chr(10)
            strSQL = strSQL & "case when dc.divno in (507) then 'SA'" & Chr(10)
            strSQL = strSQL & "end end end end end as Region" & Chr(10)
            strSQL = strSQL & "from cbis599p.dbo.contract c" & Chr(10)
            strSQL = strSQL & "inner join cbis599p.dbo.product p on p.productcode = c.productcode" & Chr(10)
            strSQL = strSQL & "inner join cbis599p.dbo.divcontracthis dc on dc.contractno = c.contractno" & Chr(10)
            strSQL = strSQL & "where not p.cgno in (58, 61, 27) and dc.bracketno = 1 and dc.sendtodiv = 1" & Chr(10)
            strSQL = strSQL & "and ((p.productclass = 1 and c.deliveryfrom  <= getdate() and isnull(c.deliveryto,getdate()) >= getdate())" & Chr(10)
            strSQL = strSQL & "or (p.productclass = 4 and p.seasonid in (5,6,7) and dateadd(mm,-1,c.deliveryfrom)  <= getdate() and isnull(c.deliveryto,getdate()) >= getdate()))" & Chr(10)
            strSQL = strSQL & ") as DATA " & Chr(10)
            strSQL = strSQL & "Union" & Chr(10)
            strSQL = strSQL & "select distinct isnull(p.con_productcode,p.productcode) as pcode, 'National' as Region" & Chr(10)
            strSQL = strSQL & "from cbis599p.dbo.contract c" & Chr(10)
            strSQL = strSQL & "inner join cbis599p.dbo.product p on p.productcode = c.productcode" & Chr(10)
            strSQL = strSQL & "inner join cbis599p.dbo.divcontracthis dc on dc.contractno = c.contractno" & Chr(10)
            strSQL = strSQL & "where not p.cgno in (58, 61, 27) and dc.bracketno = 1 and dc.sendtodiv = 1" & Chr(10)
            strSQL = strSQL & "and ((p.productclass = 1 and c.deliveryfrom  <= getdate() and isnull(c.deliveryto,getdate()) >= getdate())" & Chr(10)
            strSQL = strSQL & "or (p.productclass = 4 and p.seasonid in (5,6,7) and dateadd(mm,-1,c.deliveryfrom)  <= getdate() and isnull(c.deliveryto,getdate()) >= getdate()))" & Chr(10)
            strSQL = strSQL & "Order by Region" & Chr(10)
            'Debug.Print strSQL
    End If
    
    
    If SQLType = "COMRADE_DMDataSheet" Then
            strSQL = "SET DATEFORMAT dmy" & Chr(10)
            strSQL = strSQL & "DECLARE @CWDATE as Date" & Chr(10)
            strSQL = strSQL & "DECLARE @LWDATE as Date" & Chr(10)
            strSQL = strSQL & "SET @LWDATE = (select CONVERT(date,Datescraped) AS 'WWPenultimateUpdate' from (" & Chr(10)
            strSQL = strSQL & "select distinct CONVERT(date,Datescraped) as Datescraped, row_number() Over (order by CONVERT(date,Datescraped)) PT from  (" & Chr(10)
            strSQL = strSQL & "select distinct 'WW' as 'Table', CONVERT(date,Datescraped) as Datescraped from tools.dbo.com_dm_Prod" & Chr(10)
            strSQL = strSQL & "group by CONVERT(date,Datescraped) ) a) b" & Chr(10)
            strSQL = strSQL & "where b.PT = (select max(PT) from (" & Chr(10)
            strSQL = strSQL & "select distinct CONVERT(date,Datescraped) as Datescraped, row_number() Over (order by CONVERT(date,Datescraped)) PT from  (" & Chr(10)
            strSQL = strSQL & "select distinct 'WW' as 'Table', CONVERT(date,Datescraped) as Datescraped from tools.dbo.com_dm_Prod" & Chr(10)
            strSQL = strSQL & "group by CONVERT(date,Datescraped) ) a) b )-1)" & Chr(10)
            strSQL = strSQL & "SET @CWDATE = (select max(convert(date, datescraped)) from Tools.dbo.com_dm_prod)" & Chr(10)
            strSQL = strSQL & "select CWDATA.PCode,  CWDATA.packsize,  LWDATA.Retail as 'LastCOMRetail', CWDATA.Retail, CWDATA.Wprice, CWDATA.saved," & Chr(10)
            strSQL = strSQL & "CWDATA.cupunit as CWCupUnit, CWDATA.cupval as CWCupVal, CWDATA.cupretail as CWCupRetail," & Chr(10)
            strSQL = strSQL & "LWDATA.cupunit as LWCupUnit, LWDATA.cupval as LWCupVal, LWDATA.cupretail as LWCupRetail, LWDATA.state, CWDATA.description, CWDATA.Brand" & Chr(10)
            strSQL = strSQL & "from (" & Chr(10)
            strSQL = strSQL & "select b.pcode, b.brand, b.name as 'description', b.packsize, b.retail, b.Wprice, b.saved, cupmeasureunits as 'cupunit', cupmeasurevalue as 'cupval' , cupsize as 'cupretail', b.state from (" & Chr(10)
            strSQL = strSQL & "select a.pcode, a.brand, a.name, a.packsize, a.retail, a.Wprice, a.saved, a.cupmeasureunits, a.cupmeasurevalue, a.cupsize," & Chr(10)
            strSQL = strSQL & "a.state, row_number() over (partition by a.pcode, a.state, a.packsize order by a.retail) as row from (" & Chr(10)
            strSQL = strSQL & "select distinct productid as 'PCode', isnull(brand,'') as Brand, name, sizedescription1 as packsize, round(price1,2) as Retail, NULL as 'WPrice', NULL as 'Saved'," & Chr(10)
            strSQL = strSQL & "null as cupmeasureunits,null as  cupmeasurevalue, null as cupsize, s.StoreScan_Storestate as 'State'--, CP.Name as Description" & Chr(10)
            strSQL = strSQL & "from Tools.dbo.com_dm_prod as CP" & Chr(10)
            strSQL = strSQL & "left join tools.dbo.com_dm_stores as s on s.StoreScan_StoreNo = CP.t0_Stores_StoreScan_StoreNo" & Chr(10)
            strSQL = strSQL & "where Datescraped = @CWDATE" & Chr(10)
            strSQL = strSQL & "and CP.t0_Stores_StoreScan_StoreNo in (select distinct t0_Stores_StoreScan_StoreNo from Tools.dbo.com_dm_prod where Datescraped = @CWDATE)" & Chr(10)
            strSQL = strSQL & "group by productid , name,brand,price1,sizedescription1, s.StoreScan_Storestate" & Chr(10)
            strSQL = strSQL & "union" & Chr(10)
            strSQL = strSQL & "select distinct productid as 'PCode', isnull(brand,'') as Brand, name, sizedescription2 as packsize, round(price2,2) as Retail, NULL as 'WPrice', NULL as 'Saved'," & Chr(10)
            strSQL = strSQL & "null as cupmeasureunits,null as  cupmeasurevalue, null as cupsize, s.StoreScan_Storestate as 'State'--, CP.Name as Description" & Chr(10)
            strSQL = strSQL & "from Tools.dbo.com_dm_prod as CP" & Chr(10)
            strSQL = strSQL & "left join tools.dbo.com_dm_stores as s on s.StoreScan_StoreNo = CP.t0_Stores_StoreScan_StoreNo" & Chr(10)
            strSQL = strSQL & "where Datescraped = @CWDATE" & Chr(10)
            strSQL = strSQL & "and CP.t0_Stores_StoreScan_StoreNo in (select distinct t0_Stores_StoreScan_StoreNo from Tools.dbo.com_dm_prod where Datescraped = @CWDATE)" & Chr(10)
            strSQL = strSQL & "group by productid , name,brand,price2,sizedescription2, s.StoreScan_Storestate" & Chr(10)
            strSQL = strSQL & "union" & Chr(10)
            strSQL = strSQL & "select distinct productid as 'PCode', isnull(brand,'') as Brand, name, sizedescription3 as packsize, round(price3,2) as Retail, NULL as 'WPrice', NULL as 'Saved'," & Chr(10)
            strSQL = strSQL & "null as cupmeasureunits,null as  cupmeasurevalue, null as cupsize, s.StoreScan_Storestate as 'State'--, CP.Name as Description" & Chr(10)
            strSQL = strSQL & "from Tools.dbo.com_dm_prod as CP" & Chr(10)
            strSQL = strSQL & "left join tools.dbo.com_dm_stores as s on s.StoreScan_StoreNo = CP.t0_Stores_StoreScan_StoreNo" & Chr(10)
            strSQL = strSQL & "where Datescraped = @CWDATE" & Chr(10)
            strSQL = strSQL & "and CP.t0_Stores_StoreScan_StoreNo in (select distinct t0_Stores_StoreScan_StoreNo from Tools.dbo.com_dm_prod where Datescraped = @CWDATE)" & Chr(10)
            strSQL = strSQL & "group by productid , name,brand,price3,sizedescription3, s.StoreScan_Storestate" & Chr(10)
            strSQL = strSQL & ") as a where a.retail > 0) as b" & Chr(10)
            strSQL = strSQL & "where b.Row = 1" & Chr(10)
            strSQL = strSQL & "group by b.pcode, b.state, b.brand, b.name, b.packsize, b.retail, b.Wprice, b.cupmeasureunits, b.cupmeasurevalue, b.cupsize, b.saved" & Chr(10)
            strSQL = strSQL & ") as CWDATA" & Chr(10)
            strSQL = strSQL & "left join (" & Chr(10)
            strSQL = strSQL & "select b.pcode, b.brand, b.name, b.packsize, b.retail, b.Wprice, b.saved, cupmeasureunits as 'cupunit', cupmeasurevalue as 'cupval' , cupsize as 'cupretail', b.state from (" & Chr(10)
            strSQL = strSQL & "select a.pcode, a.brand, a.name, a.packsize, a.retail, a.Wprice, a.saved, a.cupmeasureunits, a.cupmeasurevalue, a.cupsize," & Chr(10)
            strSQL = strSQL & "a.state, row_number() over (partition by a.pcode, a.state order by a.retail) as row from (" & Chr(10)
            strSQL = strSQL & "select distinct productid as 'PCode', isnull(brand,'') as Brand, name, sizedescription1 as packsize, round(price1,2) as Retail, NULL as 'WPrice', NULL as 'Saved'," & Chr(10)
            strSQL = strSQL & "null as cupmeasureunits,null as  cupmeasurevalue, null as cupsize, s.StoreScan_Storestate as 'State', CP.Name as 'Description'" & Chr(10)
            strSQL = strSQL & "from Tools.dbo.com_dm_prod as CP" & Chr(10)
            strSQL = strSQL & "left join tools.dbo.com_dm_stores as s on s.StoreScan_StoreNo = CP.t0_Stores_StoreScan_StoreNo" & Chr(10)
            strSQL = strSQL & "where Datescraped = @LWDATE" & Chr(10)
            strSQL = strSQL & "and CP.t0_Stores_StoreScan_StoreNo in (select distinct t0_Stores_StoreScan_StoreNo from Tools.dbo.com_dm_prod where Datescraped = @LWDATE)" & Chr(10)
            strSQL = strSQL & "group by productid , name,brand,price1,sizedescription1, s.StoreScan_Storestate" & Chr(10)
            strSQL = strSQL & ") as a) as b" & Chr(10)
            strSQL = strSQL & "where b.Row = 1" & Chr(10)
            strSQL = strSQL & "group by b.pcode, b.state, b.brand, b.name, b.packsize, b.retail, b.Wprice, b.cupmeasureunits, b.cupmeasurevalue, b.cupsize, b.saved" & Chr(10)
            strSQL = strSQL & ") as LWDATA on LWDATA.Pcode = CWDATA.Pcode and LWDATA.state = CWDATA.state" & Chr(10)
            strSQL = strSQL & "order by Pcode, packsize" & Chr(10)
            'Debug.Print strSQL
    
    End If
    
    If SQLType = "COM_Matchbyline" Then
            CBA_DBtoQuery = 3
            strSQL = "select convert(int,A_Code) as ACode, A_CG, A_SCG, C_Code as 'MatchCode', 'C_Code' as 'MatchType', null from tools.dbo.com_prodmap where C_Code is not null" & Chr(10)
            strSQL = strSQL & "union select convert(int,A_Code) as ACode, A_CG, A_SCG, W_Code as 'MatchCode', 'W_Code' as 'MatchType', null from tools.dbo.com_prodmap where W_Code is not null" & Chr(10)
            strSQL = strSQL & "union select  convert(int,A_Code) as ACode, A_CG, A_SCG, C_SBCode as 'MatchCode', 'C_SBCode' as 'MatchType', null from tools.dbo.com_prodmap where C_SBCode is not null" & Chr(10)
            strSQL = strSQL & "union select  convert(int,A_Code) as ACode, A_CG, A_SCG, W_HBCode as 'MatchCode', 'W_HBCode' as 'MatchType', null from tools.dbo.com_prodmap where W_HBCode is not null" & Chr(10)
            strSQL = strSQL & "union select  convert(int,A_Code) as ACode, A_CG, A_SCG, C_CVCode as 'MatchCode', 'C_CVCode' as 'MatchType', null from tools.dbo.com_prodmap where C_CVCode is not null" & Chr(10)
            strSQL = strSQL & "union select  convert(int,A_Code) as ACode, A_CG, A_SCG, C_CCode as 'MatchCode', 'C_CCode' as 'MatchType', null from tools.dbo.com_prodmap where C_CCode is not null" & Chr(10)
            strSQL = strSQL & "union select  convert(int,A_Code) as ACode, A_CG, A_SCG, W_SelCode as 'MatchCode', 'W_SelCode' as 'MatchType', null from tools.dbo.com_prodmap where W_SelCode is not null" & Chr(10)
            strSQL = strSQL & "union select  convert(int,A_Code) as ACode, A_CG, A_SCG, B_MLCode as 'MatchCode', 'B_MLCode' as 'MatchType', null from tools.dbo.com_prodmap where B_MLCode is not null" & Chr(10)
            strSQL = strSQL & "union select  convert(int,A_Code) as ACode, A_CG, A_SCG, B_CBCode as 'MatchCode', 'B_CBCode' as 'MatchType', null from tools.dbo.com_prodmap where B_CBCode is not null" & Chr(10)
            strSQL = strSQL & "union select  convert(int,A_Code) as ACode, A_CG, A_SCG, B_AVCode as 'MatchCode', 'B_AVCode' as 'MatchType', null from tools.dbo.com_prodmap where B_AVCode is not null" & Chr(10)
            strSQL = strSQL & "union select  convert(int,A_Code) as ACode, A_CG, A_SCG, DM_Code1 as 'MatchCode', 'DM_Code1' as 'MatchType', null from tools.dbo.com_prodmap where DM_Code1 is not null" & Chr(10)
            'strSQL = strSQL & "union select  convert(int,A_Code) as ACode, A_CG, A_SCG, DM_Code1 as 'MatchCode', 'DM_Code1' as 'MatchType', DM_Code1Pack from tools.dbo.com_prodmap where DM_Code1 is not null" & Chr(10)
            strSQL = strSQL & "union select  convert(int,A_Code) as ACode, A_CG, A_SCG, DM_Code2 as 'MatchCode', 'DM_Code2' as 'MatchType', null from tools.dbo.com_prodmap where DM_Code2 is not null" & Chr(10)
            'strSQL = strSQL & "union select  convert(int,A_Code) as ACode, A_CG, A_SCG, DM_Code2 as 'MatchCode', 'DM_Code2' as 'MatchType', DM_Code2Pack from tools.dbo.com_prodmap where DM_Code2 is not null" & Chr(10)
            
            
            strSQL = strSQL & "order by convert(int,A_Code) " & Chr(10)
            'Debug.Print strSQL
    End If
    
    
    If SQLType = "CBIS_Packinfo" Then
            CBA_DBtoQuery = 599
            strSQL = "select distinct p.productcode, p.cgno, p.scgno, p.productclass, p.description, p.packsize, p.unitcode_wght, p.unitcode_Vol, p.unitcode_Pack," & Chr(10)
            strSQL = strSQL & "p.israndomweight, p.iscoldstorage, p.isfrozen, p.isfood, p.isdirect, p.weightvolume, p.contents, p.weight, p.weightKG," & Chr(10)
            strSQL = strSQL & "p.volume , p.pricecrddesc, p.pricecrddoubleinf" & Chr(10)
            strSQL = strSQL & " from cbis599p.dbo.product as P" & Chr(10)
            'Debug.Print strSQL
            
    End If
    
    
    If SQLType = "COMRADE_Mappings" Then
    
            strSQL = "select A_Code,  C_Code ,C_SBCode, C_CVCode, C_CCode, W_HBCode, W_SelCode, W_Code" & Chr(10)
            strSQL = strSQL & ", B_MLCode, B_CBCode, B_AVCode, DM_Code1, DM_Code1pack, DM_Code2, DM_Code2Pack" & Chr(10)
            strSQL = strSQL & ",C_SBCode2,C_SBCode3,C_CCode2,C_CCode3,C_CCode4,C_CVCode,C_CVCode2,C_CVCode3" & Chr(10)
            strSQL = strSQL & ",C_Code2,C_Code3,C_Code4,C_Code5,C_Code6,W_Code2,W_Code3,W_Code4,W_Code5,W_Code6" & Chr(10)
            strSQL = strSQL & ",W_SelCode2,W_SelCode3,W_SelCode4,W_SelCode5,W_HBCode2,W_HBCode3,B_PBCode,B_PBCode2" & Chr(10)
            strSQL = strSQL & ",B_PBCode2Pack,B_MLCode2,B_MLCode3,B_MLCode4,B_MLCode5,B_MLCode6,B_MLCode5Pack" & Chr(10)
            strSQL = strSQL & ",B_MLCode6Pack,B_CBCode2,B_CBCode2Pack,B_CBCode3" & Chr(10)
            strSQL = strSQL & "from tools.dbo.com_Prodmap" & Chr(10)
            strSQL = strSQL & "where C_Code Is Not Null Or C_SBCode Is Not Null Or C_CVCode Is Not Null Or C_CCode Is Not Null" & Chr(10)
            strSQL = strSQL & "or  W_HBCode is not null or  W_SelCode is not null or  W_Code is not null or  B_MLCode is not null" & Chr(10)
            strSQL = strSQL & "or  B_CBCode is not null or  B_AVCode is not null or  DM_Code1 is not null or  DM_Code1pack is not null" & Chr(10)
            strSQL = strSQL & "or  DM_Code2 is not null or  DM_Code2Pack  is not null or C_SBCode2 is not null or C_SBCode3 is not null" & Chr(10)
            strSQL = strSQL & "or C_CCode2 is not null or C_CCode3 is not null or C_CCode4 is not null or C_CVCode is not null" & Chr(10)
            strSQL = strSQL & "or C_CVCode2 is not null or C_CVCode3 is not null or C_Code2 is not null or C_Code3 is not null" & Chr(10)
            strSQL = strSQL & "or C_Code4 is not null or C_Code5 is not null or C_Code6 is not null or W_Code2 is not null" & Chr(10)
            strSQL = strSQL & "or W_Code3 is not null or W_Code4 is not null or W_Code5 is not null or W_Code6 is not null" & Chr(10)
            strSQL = strSQL & "or W_SelCode2 is not null or W_SelCode3 is not null or W_SelCode4 is not null or W_SelCode5 is not null" & Chr(10)
            strSQL = strSQL & "or W_HBCode2 is not null or W_HBCode3 is not null or B_PBCode is not null or B_PBCode2 is not null" & Chr(10)
            strSQL = strSQL & "or B_PBCode2Pack is not null or B_MLCode2 is not null or B_MLCode3 is not null or B_MLCode4 is not null" & Chr(10)
            strSQL = strSQL & "or B_MLCode5 is not null or B_MLCode6 is not null or B_MLCode5Pack is not null or B_MLCode6Pack is not null" & Chr(10)
            strSQL = strSQL & "or B_CBCode2 is not null or B_CBCode2Pack is not null or B_CBCode3 is not null" & Chr(10)
            'Debug.Print strSQL
    
    End If
    
    
    If SQLType = "COMRADE_ColesDataSheet" Then
    
            strSQL = "SET DATEFORMAT dmy" & Chr(10)
            strSQL = strSQL & "DECLARE @CWDATE as Date" & Chr(10)
            strSQL = strSQL & "DECLARE @LWDATE as Date" & Chr(10)
            strSQL = strSQL & "SET @LWDATE = (select CONVERT(date,Datescraped) AS 'WWPenultimateUpdate' from (" & Chr(10)
            strSQL = strSQL & "select distinct CONVERT(date,Datescraped) as Datescraped, row_number() Over (order by CONVERT(date,Datescraped)) PT from  (" & Chr(10)
            strSQL = strSQL & "select distinct 'Coles' as 'Table', CONVERT(date,Datescraped) as Datescraped from tools.dbo.com_C_Prod" & Chr(10)
            strSQL = strSQL & "group by CONVERT(date,Datescraped) ) a) b" & Chr(10)
            strSQL = strSQL & "where b.PT = (select max(PT) from (" & Chr(10)
            strSQL = strSQL & "select distinct CONVERT(date,Datescraped) as Datescraped, row_number() Over (order by CONVERT(date,Datescraped)) PT from  (" & Chr(10)
            strSQL = strSQL & "select distinct 'Coles' as 'Table', CONVERT(date,Datescraped) as Datescraped from tools.dbo.com_C_Prod" & Chr(10)
            strSQL = strSQL & "group by CONVERT(date,Datescraped) ) a) b )-1)" & Chr(10)
            strSQL = strSQL & "SET @CWDATE = (select max(convert(date, datescraped)) from Tools.dbo.com_c_prod)" & Chr(10)
            strSQL = strSQL & "select CWDATA.PCode,  CWDATA.packsize,  LWDATA.Retail as 'LastCOMRetail', CWDATA.Retail, CWDATA.Wprice, CWDATA.saved, " & Chr(10)
            strSQL = strSQL & "CWDATA.cupunit as CWCupUnit, CWDATA.cupval as CWCupVal, CWDATA.cupretail as CWCupRetail," & Chr(10)
            strSQL = strSQL & "LWDATA.cupunit as LWCupUnit, LWDATA.cupval as LWCupVal, LWDATA.cupretail as LWCupRetail, CWDATA.state, CWDATA.Description , CWDATA.Brand" & Chr(10)
            strSQL = strSQL & "from (" & Chr(10)
            strSQL = strSQL & "select b.pcode, b.brand, b.name, b.packsize, b.retail, b.Wprice, b.saved, cupmeasureunits as 'cupunit', cupmeasurevalue as 'cupval' , cupsize as 'cupretail', b.state, b.Description from (" & Chr(10)
            strSQL = strSQL & "select a.pcode, a.brand, a.name, a.packsize, a.retail, a.Wprice, a.saved, a.cupmeasureunits, a.cupmeasurevalue, a.cupsize, a.state, a.Description, row_number() over (partition by a.pcode, a.state order by a.retail) as row from (" & Chr(10)
            strSQL = strSQL & "select distinct colesproductid as 'PCode', isnull(brand,'') as Brand, name, packsize, round(price,2) as Retail, priceprevious as 'WPrice', pricesaving as 'Saved', cupmeasureunits, cupmeasurevalue, cupsize," & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = 'NSW' then 'NSW' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = 'VIC' then 'VIC' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = 'QLD' then 'QLD' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,2) = 'SA' then 'SA' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,2) = 'WA' then 'WA' else 'Unknown'" & Chr(10)
            strSQL = strSQL & "end end end end end as 'State' , CP.Name as 'Description'" & Chr(10)
            strSQL = strSQL & "from Tools.dbo.com_c_prod as CP" & Chr(10)
            strSQL = strSQL & "where Datescraped = @CWDATE" & Chr(10)
            strSQL = strSQL & "and UrlScan_StoreSeoToken in (select distinct UrlScan_StoreSeoToken from Tools.dbo.com_c_prod where Datescraped = @CWDATE)" & Chr(10)
            strSQL = strSQL & "group by colesproductid, Brand, name, packsize, priceprevious, pricesaving, price, cupmeasureunits, cupmeasurevalue, cupsize, UrlScan_StoreSeoToken) as a) as b" & Chr(10)
            strSQL = strSQL & "where b.Row = 1" & Chr(10)
            strSQL = strSQL & "group by b.pcode, b.state, b.brand, b.name, b.packsize, b.retail, b.Wprice, b.cupmeasureunits, b.cupmeasurevalue, b.cupsize, b.saved, b.Description" & Chr(10)
            strSQL = strSQL & ") as CWDATA" & Chr(10)
            strSQL = strSQL & "left join (" & Chr(10)
            strSQL = strSQL & "select b.pcode, b.brand, b.name, b.packsize, b.retail, b.Wprice, b.saved, cupmeasureunits as 'cupunit', cupmeasurevalue as 'cupval' , cupsize as 'cupretail', b.state from (" & Chr(10)
            strSQL = strSQL & "select a.pcode, a.brand, a.name, a.packsize, a.retail, a.Wprice, a.saved, a.cupmeasureunits, a.cupmeasurevalue, a.cupsize, a.state, row_number() over (partition by a.pcode, a.state order by a.retail) as row from (" & Chr(10)
            strSQL = strSQL & "select distinct colesproductid as 'PCode', isnull(brand,'') as Brand, name, packsize, round(price,2) as Retail, priceprevious as 'WPrice', pricesaving as 'Saved', cupmeasureunits, cupmeasurevalue, cupsize," & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = 'NSW' then 'NSW' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = 'VIC' then 'VIC' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,3) = 'QLD' then 'QLD' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,2) = 'SA' then 'SA' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), UrlScan_StoreSeoToken),3,2) = 'WA' then 'WA' else 'Unknown'" & Chr(10)
            strSQL = strSQL & "end end end end end as 'State'" & Chr(10)
            strSQL = strSQL & "from Tools.dbo.com_c_prod as CP" & Chr(10)
            strSQL = strSQL & "where Datescraped = @LWDATE" & Chr(10)
            strSQL = strSQL & "and UrlScan_StoreSeoToken in (select distinct UrlScan_StoreSeoToken from Tools.dbo.com_c_prod where Datescraped = @LWDATE)" & Chr(10)
            strSQL = strSQL & "group by colesproductid, Brand, name, packsize, priceprevious, pricesaving, price, cupmeasureunits, cupmeasurevalue, cupsize, UrlScan_StoreSeoToken) as a) as b" & Chr(10)
            strSQL = strSQL & "where b.Row = 1" & Chr(10)
            strSQL = strSQL & "group by b.pcode, b.state, b.brand, b.name, b.packsize, b.retail, b.Wprice, b.cupmeasureunits, b.cupmeasurevalue, b.cupsize, b.saved" & Chr(10)
            strSQL = strSQL & ") as LWDATA on LWDATA.Pcode = CWDATA.Pcode and LWDATA.state = CWDATA.state" & Chr(10)
    '        strSQL = strSQL & "where  CWDATA.PCode in (select distinct convert(nvarchar(12),COMPM.C_Code) as 'MatchedPCode' from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
    '        strSQL = strSQL & "inner join tools.dbo.com_C_Prod as CC on convert(nvarchar(12),CC.ColesProductId) = convert(nvarchar(12),COMPM.C_Code) " & Chr(10)
    '        strSQL = strSQL & "union select distinct convert(nvarchar(12),COMPM.C_SBCode) as 'MatchedPCode' from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
    '        strSQL = strSQL & "inner join tools.dbo.com_C_Prod as CSB on convert(nvarchar(12),CSB.ColesProductId) = convert(nvarchar(12),COMPM.C_SBCode) " & Chr(10)
    '        strSQL = strSQL & "union select distinct convert(nvarchar(12),COMPM.C_CVCode) as 'MatchedPCode' from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
    '        strSQL = strSQL & "inner join tools.dbo.com_C_Prod as CCV on convert(nvarchar(12),CCV.ColesProductId) = convert(nvarchar(12),COMPM.C_CVCode) " & Chr(10)
    '        strSQL = strSQL & "union select distinct convert(nvarchar(12),COMPM.C_CCode) as 'MatchedPCode' from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
    '        strSQL = strSQL & "inner join tools.dbo.com_C_Prod as CCC on convert(nvarchar(12),CCC.ColesProductId) = convert(nvarchar(12),COMPM.C_CCode) " & Chr(10)
    '        strSQL = strSQL & "union select distinct convert(nvarchar(12),COMPM.B_MLCode) as 'MatchedPCode' from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
    '        strSQL = strSQL & "inner join tools.dbo.com_C_Prod as CBML on convert(nvarchar(12),CBML.ColesProductId) = convert(nvarchar(12),COMPM.B_MLCode) " & Chr(10)
    '        strSQL = strSQL & "union select distinct convert(nvarchar(12),COMPM.B_MLCode) as 'MatchedPCode' from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
    '        strSQL = strSQL & "inner join tools.dbo.com_C_Prod as CBCB on convert(nvarchar(12),CBCB.ColesProductId) = convert(nvarchar(12),COMPM.B_CBCode) )" & Chr(10)
            strSQL = strSQL & "order by Pcode" & Chr(10)
            'Debug.Print strSQL
    
    End If
    
    
    
    If SQLType = "COMRADE_WWDataSheet" Then
    
            strSQL = "SET DATEFORMAT dmy" & Chr(10)
            strSQL = strSQL & "DECLARE @CWDATE as Date" & Chr(10)
            strSQL = strSQL & "DECLARE @LWDATE as Date" & Chr(10)
            strSQL = strSQL & "SET @LWDATE = (select CONVERT(date,Datescraped) AS 'WWPenultimateUpdate' from (" & Chr(10)
            strSQL = strSQL & "select distinct CONVERT(date,Datescraped) as Datescraped, row_number() Over (order by CONVERT(date,Datescraped)) PT from  (" & Chr(10)
            strSQL = strSQL & "select distinct 'WW' as 'Table', CONVERT(date,Datescraped) as Datescraped from tools.dbo.com_w_Prod" & Chr(10)
            strSQL = strSQL & "group by CONVERT(date,Datescraped) ) a) b" & Chr(10)
            strSQL = strSQL & "where b.PT = (select max(PT) from (" & Chr(10)
            strSQL = strSQL & "select distinct CONVERT(date,Datescraped) as Datescraped, row_number() Over (order by CONVERT(date,Datescraped)) PT from  (" & Chr(10)
            strSQL = strSQL & "select distinct 'WW' as 'Table', CONVERT(date,Datescraped) as Datescraped from tools.dbo.com_w_Prod" & Chr(10)
            strSQL = strSQL & "group by CONVERT(date,Datescraped) ) a) b )-1)" & Chr(10)
            strSQL = strSQL & "SET @CWDATE = (select max(convert(date, datescraped)) from Tools.dbo.com_w_prod)" & Chr(10)
            strSQL = strSQL & "select CWDATA.Stockcode as 'PCode', CWDATA.packsize, LWDATA.Retail as 'LastCOMRetail', CWDATA.Retail,  CWDATA.wprice , CWDATA.saved, " & Chr(10)
            strSQL = strSQL & "CWDATA.cupunit as CWCupunit, CWDATA.cupval as CWCupval, CWDATA.cupretail as CWCupretail," & Chr(10)
            strSQL = strSQL & "LWDATA.cupunit as LWCupunit, LWDATA.cupval as LWCupval, LWDATA.cupretail as LWCupretail, CWDATA.state, CWDATA.Description , CWDATA.Brand" & Chr(10)
            strSQL = strSQL & "from (" & Chr(10)
            strSQL = strSQL & "select stockcode, brand, name, addressid,  packsize, Retail, wprice , saved,cupmeasureunits as 'cupunit', cupmeasurevalue as 'cupval' , cupprice as 'cupretail',  state , name as 'Description' from (" & Chr(10)
            strSQL = strSQL & "select stockcode, Retail, addressid, Brand, name, packsize, wprice , saved, cupmeasureunits, cupmeasurevalue, cupprice, state," & Chr(10)
            strSQL = strSQL & "row_number() over (partition by stockcode, state order by Retail ) as row" & Chr(10)
            strSQL = strSQL & "from ( select stockcode, round(price,2) as 'Retail', p.addressid, isnull(brand,'') as 'Brand', name," & Chr(10)
            strSQL = strSQL & "packagesize as 'Packsize', wasprice as 'WPrice', savingsamount as 'Saved'," & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '1' then 'NSW' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '2' then 'NSW' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '3' then 'VIC' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '4' then 'QLD' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '5' then 'SA' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '6' then 'WA' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '8' then 'VIC' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '9' then 'QLD'" & Chr(10)
            strSQL = strSQL & "end end end end end end end end as 'State'," & Chr(10)
            strSQL = strSQL & "cupmeasureunits, cupmeasurevalue, cupprice" & Chr(10)
            strSQL = strSQL & "from Tools.dbo.com_W_prod as p" & Chr(10)
            strSQL = strSQL & "left join Tools.dbo.com_W_Stores as s on p.addressid = s.addressid" & Chr(10)
            strSQL = strSQL & "where datescraped = @CWDATE" & Chr(10)
            strSQL = strSQL & "group by stockcode,  price, s.Postcode, p.addressid, brand, name,  packagesize, wasprice ,cupmeasureunits, cupmeasurevalue, cupprice, savingsamount" & Chr(10)
            strSQL = strSQL & ") as a ) as b where b.Row = 1 ) as CWDATA" & Chr(10)
            strSQL = strSQL & "left join (" & Chr(10)
            strSQL = strSQL & "select stockcode, brand, name, addressid,  packsize, Retail, wprice , saved,cupmeasureunits as 'cupunit', cupmeasurevalue as 'cupval' , cupprice as 'cupretail',  state from (" & Chr(10)
            strSQL = strSQL & "select stockcode, Retail, addressid, Brand, name, packsize, wprice , saved, cupmeasureunits, cupmeasurevalue, cupprice, state," & Chr(10)
            strSQL = strSQL & "row_number() over (partition by stockcode, state order by Retail ) as row" & Chr(10)
            strSQL = strSQL & "from ( select stockcode, round(price,2) as 'Retail', p.addressid, isnull(brand,'') as 'Brand', name," & Chr(10)
            strSQL = strSQL & "packagesize as 'Packsize', wasprice as 'WPrice', savingsamount as 'Saved'," & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '1' then 'NSW' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '2' then 'NSW' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '3' then 'VIC' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '4' then 'QLD' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '5' then 'SA' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '6' then 'WA' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '8' then 'VIC' else" & Chr(10)
            strSQL = strSQL & "case when substring(convert(nvarchar(max), s.Postcode),1,1) = '9' then 'QLD'" & Chr(10)
            strSQL = strSQL & "end end end end end end end end as 'State'," & Chr(10)
            strSQL = strSQL & "cupmeasureunits, cupmeasurevalue, cupprice" & Chr(10)
            strSQL = strSQL & "from Tools.dbo.com_W_prod as p" & Chr(10)
            strSQL = strSQL & "left join Tools.dbo.com_W_Stores as s on p.addressid = s.addressid" & Chr(10)
            strSQL = strSQL & "where datescraped = @LWDATE" & Chr(10)
            strSQL = strSQL & "group by stockcode,  price, s.Postcode, p.addressid, brand, name,  packagesize, wasprice ,cupmeasureunits, cupmeasurevalue, cupprice, savingsamount" & Chr(10)
            strSQL = strSQL & ") as a ) as b where b.Row = 1 ) as LWDATA on LWDATA.Stockcode = CWDATA.Stockcode and LWDATA.state = CWDATA.state" & Chr(10)
    '        strSQL = strSQL & "where CWDATA.Stockcode in (select distinct convert(nvarchar(12),COMPM.W_Code) as 'MatchedPCode'" & Chr(10)
    '        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
    '        strSQL = strSQL & "inner join tools.dbo.com_W_Prod as WC on convert(nvarchar(12),WC.Stockcode) = convert(nvarchar(12),COMPM.W_Code)" & Chr(10)
    '        strSQL = strSQL & "union select distinct convert(nvarchar(12),COMPM.W_HBCode) as 'MatchedPCode'" & Chr(10)
    '        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
    '        strSQL = strSQL & "inner join tools.dbo.com_W_Prod as WHB on convert(nvarchar(12),WHB.Stockcode) = convert(nvarchar(12),COMPM.W_HBCode)" & Chr(10)
    '        strSQL = strSQL & "union select distinct convert(nvarchar(12),COMPM.W_SelCode) as 'MatchedPCode'" & Chr(10)
    '        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
    '        strSQL = strSQL & "inner join tools.dbo.com_W_Prod as WSel on convert(nvarchar(12),WSel.Stockcode) = convert(nvarchar(12),COMPM.W_SelCode)" & Chr(10)
    '        strSQL = strSQL & "union select distinct convert(nvarchar(12),COMPM.B_MLCode) as 'MatchedPCode'" & Chr(10)
    '        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
    '        strSQL = strSQL & "inner join tools.dbo.com_W_Prod as WBML on convert(nvarchar(12),WBML.Stockcode) = convert(nvarchar(12),COMPM.B_MLCode)" & Chr(10)
    '        strSQL = strSQL & "union select distinct convert(nvarchar(12),COMPM.B_CBCode) as 'MatchedPCode'" & Chr(10)
    '        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
    '        strSQL = strSQL & "inner join tools.dbo.com_W_Prod as WBCB on convert(nvarchar(12),WBCB.Stockcode) = convert(nvarchar(12),COMPM.B_CBCode))" & Chr(10)
            strSQL = strSQL & "order by CWDATA.stockcode" & Chr(10)
            'Debug.Print strSQL
    
    End If
    
    
    
    If SQLType = "COMRADE_MATCHEDPRODUCTDETAIL" Then
    
        strSQL = strSQL & "SET DATEFORMAT dmy" & Chr(10)
        strSQL = strSQL & "SELECT distinct COMPDATA.AldiPCode,  COMPDATA.A_CG, COMPDATA.A_SCG,   COMPDATA.AldiDescription,COMPDATA.MatchType, COMPDATA.MatchInfo," & Chr(10)
        strSQL = strSQL & "COMPDATA.MatchedPCode, STORES.Postcode, COMPDATA.MatchedDescription, COMPDATA.PriceDate, COMPDATA.Retail," & Chr(10)
        strSQL = strSQL & "COMPDATA.Packsize" & Chr(10)
        strSQL = strSQL & "from (" & Chr(10)
        strSQL = strSQL & "select distinct CONVERT(nvarchar(10),COMPM.A_Code) as 'AldiPCode',  COMPM.A_CG, COMPM.A_SCG,   COMPM.A_Product as 'AldiDescription','Woolworths General' as 'MatchType', 'WW' as 'MatchInfo'," & Chr(10)
        strSQL = strSQL & "COMPM.W_Code as 'MatchedPCode', cast(cast(WC.AddressID as int) as nvarchar(50)) as 'StoreRef', WC.Name as 'MatchedDescription', convert(date,WC.Datescraped) as 'PriceDate', WC.Price as 'Retail'," & Chr(10)
        strSQL = strSQL & "WC.Packagesize as 'Packsize'" & Chr(10)
        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
        strSQL = strSQL & "inner join tools.dbo.com_W_Prod as WC on WC.Stockcode = COMPM.W_Code" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select distinct CONVERT(nvarchar(10),COMPM.A_Code) as 'AldiPCode',  COMPM.A_CG, COMPM.A_SCG,   COMPM.A_Product as 'AldiDescription','Woolworths HomeBrand' as 'MatchType', 'WW' as 'MatchInfo'," & Chr(10)
        strSQL = strSQL & "COMPM.W_HBCode as 'MatchedPCode', cast(cast(WHB.AddressID as int) as nvarchar(50)) as 'StoreRef', WHB.Name as 'MatchedDescription', convert(date,WHB.Datescraped) as 'PriceDate', WHB.Price as 'Retail'," & Chr(10)
        strSQL = strSQL & "WHB.Packagesize as 'Packsize'" & Chr(10)
        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
        strSQL = strSQL & "inner join tools.dbo.com_W_Prod as WHB on WHB.Stockcode = COMPM.W_HBCode" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select distinct CONVERT(nvarchar(10),COMPM.A_Code) as 'AldiPCode',  COMPM.A_CG, COMPM.A_SCG,   COMPM.A_Product as 'AldiDescription','Woolworths Select' as 'MatchType', 'WW' as 'MatchInfo'," & Chr(10)
        strSQL = strSQL & "COMPM.W_SelCode as 'MatchedPCode', cast(cast(WSel.AddressID as int) as nvarchar(50)) as 'StoreRef', WSel.Name as 'MatchedDescription', convert(date,WSel.Datescraped) as 'PriceDate', WSel.Price as 'Retail'," & Chr(10)
        strSQL = strSQL & "WSel.Packagesize as 'Packsize'" & Chr(10)
        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
        strSQL = strSQL & "inner join tools.dbo.com_W_Prod as WSel on WSel.Stockcode = COMPM.W_SelCode" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select distinct CONVERT(nvarchar(10),COMPM.A_Code) as 'AldiPCode',  COMPM.A_CG, COMPM.A_SCG, COMPM.A_Product as 'AldiDescription','MarketLeader Brand' as 'MatchType', 'WW' as 'MatchInfo'," & Chr(10)
        strSQL = strSQL & "COMPM.B_MLCode as 'MatchedPCode', cast(cast(WBML.AddressID as int) as nvarchar(50)) as 'StoreRef', WBML.Name as 'MatchedDescription', convert(date,WBML.Datescraped) as 'PriceDate', WBML.Price as 'Retail'," & Chr(10)
        strSQL = strSQL & "WBML.Packagesize as 'Packsize'" & Chr(10)
        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
        strSQL = strSQL & "inner join tools.dbo.com_W_Prod as WBML on WBML.Stockcode = COMPM.B_MLCode" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select distinct CONVERT(nvarchar(10),COMPM.A_Code) as 'AldiPCode',  COMPM.A_CG, COMPM.A_SCG, COMPM.A_Product as 'AldiDescription','Control Brand' as 'MatchType', 'WW' as 'MatchInfo'," & Chr(10)
        strSQL = strSQL & "COMPM.B_CBCode as 'MatchedPCode', cast(cast(WBCB.AddressID as int) as nvarchar(50)) as 'StoreRef', WBCB.Name as 'MatchedDescription', convert(date,WBCB.Datescraped) as 'PriceDate', WBCB.Price as 'Retail'," & Chr(10)
        strSQL = strSQL & "WBCB.Packagesize as 'Packsize'" & Chr(10)
        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
        strSQL = strSQL & "inner join tools.dbo.com_W_Prod as WBCB on WBCB.Stockcode = COMPM.B_CBCode" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select distinct CONVERT(nvarchar(10),COMPM.A_Code) as 'AldiPCode',  COMPM.A_CG, COMPM.A_SCG, COMPM.A_Product as 'AldiDescription','Coles General' as 'MatchType', 'Coles' as 'MatchInfo'," & Chr(10)
        strSQL = strSQL & "COMPM.C_Code as 'MatchedPCode', cast(CC.UrlScan_StoreSeoToken as nvarchar(50)) as 'StoreRef', CC.Name as 'MatchedDescription', convert(date,CC.Datescraped) as 'PriceDate', CC.Price as 'Retail'," & Chr(10)
        strSQL = strSQL & "CC.Packsize as 'Packsize'" & Chr(10)
        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
        strSQL = strSQL & "inner join tools.dbo.com_C_Prod as CC on CC.ColesProductId = COMPM.C_Code" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select distinct CONVERT(nvarchar(10),COMPM.A_Code) as 'AldiPCode',  COMPM.A_CG, COMPM.A_SCG, COMPM.A_Product as 'AldiDescription','Coles SmartBuy' as 'MatchType', 'Coles' as 'MatchInfo'," & Chr(10)
        strSQL = strSQL & "COMPM.C_SBCode as 'MatchedPCode', cast(CSB.UrlScan_StoreSeoToken as nvarchar(50)) as 'StoreRef', CSB.Name as 'MatchedDescription', convert(date,CSB.Datescraped) as 'PriceDate', CSB.Price as 'Retail'," & Chr(10)
        strSQL = strSQL & "CSB.Packsize as 'Packsize'" & Chr(10)
        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
        strSQL = strSQL & "inner join tools.dbo.com_C_Prod as CSB on CSB.ColesProductId = COMPM.C_SBCode" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select distinct CONVERT(nvarchar(10),COMPM.A_Code) as 'AldiPCode',  COMPM.A_CG, COMPM.A_SCG, COMPM.A_Product as 'AldiDescription','Coles Value' as 'Match Type', 'Coles' as 'MatchInfo'," & Chr(10)
        strSQL = strSQL & "COMPM.C_CVCode as 'MatchedPCode', cast(CCV.UrlScan_StoreSeoToken as nvarchar(50)) as 'StoreRef', CCV.Name as 'MatchedDescription', convert(date,CCV.Datescraped) as 'PriceDate', CCV.Price as 'Retail'," & Chr(10)
        strSQL = strSQL & "CCV.Packsize as 'Packsize'" & Chr(10)
        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
        strSQL = strSQL & "inner join tools.dbo.com_C_Prod as CCV on CCV.ColesProductId = COMPM.C_CVCode" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select distinct CONVERT(nvarchar(10),COMPM.A_Code) as 'AldiPCode',  COMPM.A_CG, COMPM.A_SCG, COMPM.A_Product as 'AldiDescription','Coles OwnBrand' as 'MatchType', 'Coles' as 'MatchInfo'," & Chr(10)
        strSQL = strSQL & "COMPM.C_CCode as 'MatchedPCode', cast(CCC.UrlScan_StoreSeoToken as nvarchar(50)) as 'StoreRef', CCC.Name as 'MatchedDescription', convert(date,CCC.Datescraped) as 'PriceDate', CCC.Price as 'Retail'," & Chr(10)
        strSQL = strSQL & "CCC.Packsize as 'Packsize'" & Chr(10)
        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
        strSQL = strSQL & "inner join tools.dbo.com_C_Prod as CCC on CCC.ColesProductId = COMPM.C_CCode" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select distinct CONVERT(nvarchar(10),COMPM.A_Code) as 'AldiPCode',  COMPM.A_CG, COMPM.A_SCG, COMPM.A_Product as 'AldiDescription','MarketLeader Brand' as 'MatchType', 'Coles' as 'MatchInfo'," & Chr(10)
        strSQL = strSQL & "COMPM.B_MLCode as 'MatchedPCode', cast(CBML.UrlScan_StoreSeoToken as nvarchar(50)) as 'StoreRef', CBML.Name as 'MatchedDescription', convert(date,CBML.Datescraped) as 'PriceDate', CBML.Price as 'Retail'," & Chr(10)
        strSQL = strSQL & "CBML.Packsize as 'Packsize'" & Chr(10)
        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
        strSQL = strSQL & "inner join tools.dbo.com_C_Prod as CBML on CBML.ColesProductId = COMPM.B_MLCode" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select distinct CONVERT(nvarchar(10),COMPM.A_Code) as 'AldiPCode',  COMPM.A_CG, COMPM.A_SCG, COMPM.A_Product as 'AldiDescription','Control Brand' as 'MatchType', 'Coles' as 'MatchInfo'," & Chr(10)
        strSQL = strSQL & "COMPM.B_MLCode as 'MatchedPCode', cast(CBCB.UrlScan_StoreSeoToken as nvarchar(50)) as 'StoreRef', CBCB.Name as 'MatchedDescription', convert(date,CBCB.Datescraped) as 'PriceDate', CBCB.Price as 'Retail'," & Chr(10)
        strSQL = strSQL & "CBCB.Packsize as 'Packsize'" & Chr(10)
        strSQL = strSQL & "from tools.dbo.Com_ProdMap as COMPM" & Chr(10)
        strSQL = strSQL & "inner join tools.dbo.com_C_Prod as CBCB on CBCB.ColesProductId = COMPM.B_CBCode) AS COMPDATA" & Chr(10)
        strSQL = strSQL & "inner join (" & Chr(10)
        strSQL = strSQL & "select 'WW' as 'Retailer', CONVERT(date,Datescraped) AS 'DatesRQD' from (" & Chr(10)
        strSQL = strSQL & "select distinct datescraped, row_number() Over (order by datescraped) PT from  (" & Chr(10)
        strSQL = strSQL & "select distinct 'WW' as 'Retailer', datescraped from tools.dbo.com_W_Prod" & Chr(10)
        strSQL = strSQL & "group by datescraped ) a) b" & Chr(10)
        strSQL = strSQL & "where b.PT = (select max(PT) from (" & Chr(10)
        strSQL = strSQL & "select distinct datescraped, row_number() Over (order by datescraped) PT from  (" & Chr(10)
        strSQL = strSQL & "select distinct 'WW' as 'Table', datescraped from tools.dbo.com_W_Prod" & Chr(10)
        strSQL = strSQL & "group by datescraped ) a) b )-1" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select 'Coles' as 'Retailer', CONVERT(date,Datescraped) AS 'DatesRQD' from (" & Chr(10)
        strSQL = strSQL & "select distinct datescraped, row_number() Over (order by datescraped) PT from  (" & Chr(10)
        strSQL = strSQL & "select distinct 'Coles' as 'Table', convert(date,datescraped) as datescraped from tools.dbo.com_C_Prod" & Chr(10)
        strSQL = strSQL & "group by datescraped ) a) b" & Chr(10)
        strSQL = strSQL & "where b.PT = (select max(PT) from (" & Chr(10)
        strSQL = strSQL & "select distinct datescraped, row_number() Over (order by datescraped) PT from  (" & Chr(10)
        strSQL = strSQL & "select distinct 'Coles' as 'Table', convert(date,datescraped) as datescraped from tools.dbo.com_C_Prod" & Chr(10)
        strSQL = strSQL & "group by datescraped ) a) b )-1" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select 'Coles' as 'Retailer', MAX(CONVERT(date,datescraped)) as 'DatesRQD'" & Chr(10)
        strSQL = strSQL & "from tools.dbo.com_C_Prod" & Chr(10)
        strSQL = strSQL & "Union" & Chr(10)
        strSQL = strSQL & "select 'WW' as 'Retailer', MAX(CONVERT(date,datescraped)) as 'DatesRQD'" & Chr(10)
        strSQL = strSQL & "from tools.dbo.com_W_Prod" & Chr(10)
        strSQL = strSQL & ") as TABLEDATES on COMPDATA.PriceDate = TABLEDATES.DatesRQD and COMPDATA.MatchInfo = TABLEDATES.Retailer" & Chr(10)
        strSQL = strSQL & "inner join (" & Chr(10)
        strSQL = strSQL & "select cast(UrlScan_StoreSeoToken as nvarchar(50)) as 'StoreRef', StoreScan_Postcode as Postcode from  tools.dbo.com_c_Stores" & Chr(10)
        strSQL = strSQL & "union select cast(cast(AddressID as int)as nvarchar(50)) as 'StoreRef', Postcode from  tools.dbo.com_W_Stores" & Chr(10)
        strSQL = strSQL & ") as STORES ON STORES.StoreRef = COMPDATA.StoreRef" & Chr(10)
        strSQL = strSQL & "where ((COMPDATA.A_CG = 44 And COMPDATA.A_SCG = 3) Or (COMPDATA.A_CG = 54 And COMPDATA.A_SCG = 3))" & Chr(10)
        strSQL = strSQL & "and STORES.Postcode in (4558,4120,4218,4207,4350,2250,2022,2770,2601,2650,3550,3356,3019,3070,3975,5082,5168,6027,6168,6100,4103,2759,2609,3630,3280,3020,3072,5081,5169)" & Chr(10)
        'Debug.Print strSQL
    
    
    End If
    
    If SQLType = "ABI_ALCOHOLSTORENOS" Then
        strSQL = "SELECT AAS.DateQueried, AAS.[501], AAS.[502], AAS.[503], AAS.[504], AAS.[505], AAS.[506], AAS.[507], AAS.[509], AAS.[TOTAL]" & Chr(10)
        strSQL = strSQL & "FROM ABI_AlcoholStores as AAS" & Chr(10)
        strSQL = strSQL & "WHERE Month(AAS.DateQueried) = Month(" & DateTo & ")" & Chr(10)
    
        ''Debug.Print strSQL
    
    End If
    
    
    If SQLType = "FINDNAME" And cCG <> 0 And IsEmpty(cntSCGs) = False Then
                    For a = 1 To cntSCGs
                        If a = 1 Then str = cSCG(a) Else str = str & "," & cSCG(a)
                    Next
    strSQL = "select Firstname, Name from cbis599p.dbo.employee" & Chr(10) & _
                "where EMpNo in (select EmpNo from cbis599p.dbo.product" & Chr(10) & _
                    "where CGNo = " & cCG & " and SCGNo in (" & str & ")" & Chr(10) & _
                        "group by EmpNo ,CGNo )"
    End If
    
    If SQLType = "FINDPRODCLASS" And IsNumeric(cType) Then
    strSQL = "SELECT PRD.PRODUCTCLASS FROM CBIS599P.DBO.PRODUCT AS PRD" & Chr(10) & _
                "WHERE PRD.PRODUCTCODE = '" & cType & "'"
    End If
    
    If SQLType = "FINDLINKEDPROD" And IsNumeric(cType) Then
    strSQL = "SELECT DISTINCT PRODUCTCODE  FROM cbis599p.dbo.PRODUCT AS PRD " & Chr(10) & _
                "WHERE CON_PRODUCTCODE = '" & cType & "'"
    End If
    
    
    
    If SQLType = "FINDSEASONID" And IsNumeric(cType) Then
    strSQL = "SELECT PRD.SEASONID FROM CBIS599P.DBO.PRODUCT AS PRD" & Chr(10) & _
                "WHERE PRD.PRODUCTCODE = '" & cType & "'"
    End If
    
    If SQLType = "PORT_NOSTORE" And Not IsMissing(DateTo) Then
        CBA_DBtoQuery = 599
        strSQLprt01 = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
        strSQLprt02 = "SELECT SUM(PORT.NOOFSTORES) FROM CBIS" & CBA_DBtoQuery & "P.Portfolio.Stores AS PORT " & Chr(10)
        strSQLprt03 = "WHERE PORT.VALIDFROM >= '" & Format(DateSerial(Year(DateTo), Month(DateTo), Day(DateTo) - 6), "yyyy-mm-dd") & "' AND PORT.VALIDFROM <= '" & Format(DateTo, "yyyy-mm-dd") & "'" & Chr(10)
        If Bypass <> "" Then strSQLprt04 = "AND DIVNO IN (" & Bypass & ")"
        
        strSQL = strSQLprt01 & strSQLprt02 & strSQLprt03
        If Bypass <> "" Then strSQL = strSQL & strSQLprt04
        ''Debug.Print strSQL
    End If
    
    If SQLType = "PROD_RETAIL" Then
        If Not IsMissing(DateTo) And Not IsMissing(DateFrom) And cType <> "" Then
            strSQLprt01 = "select  sum(retail) / count(retail) from cbis599p.dbo.infocardtab" & Chr(10)
            strSQLprt02 = "where productcode = " & cType & Chr(10)
            strSQLprt03 = "and validfrom <= '" & Format(DateTo, "yyyy-mm-dd") & "' AND validto >= '" & Format(DateFrom, "yyyy-mm-dd") & "'" & Chr(10)
            
            strSQL = strSQLprt01 & strSQLprt02 & strSQLprt03
            '''Debug.Print strSQL
        End If
    End If
    
    
    If SQLType = "POS_QTY_BY" Then
    
    
    strSQL = "SET NOCOUNT ON" & Chr(10)
    strSQL = strSQL & "SET ANSI_WARNINGS OFF" & Chr(10)
    strSQL = strSQL & "SET DATEFIRST 1" & Chr(10)
    strSQL = strSQL & "SELECT DIVNO,  " & Bypass & Chr(10)
    strSQL = strSQL & "FROM( SELECT Distinct POS.DIVNO," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=1 THEN Quantity ELSE 0 END) AS Pos1," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=2 THEN Quantity ELSE 0 END) AS Pos2," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=3 THEN Quantity ELSE 0 END) AS Pos3," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=4 THEN Quantity ELSE 0 END) AS Pos4," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=5 THEN Quantity ELSE 0 END) AS Pos5," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=6 THEN Quantity ELSE 0 END) AS Pos6," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=7 THEN Quantity ELSE 0 END) AS Pos7," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=8 THEN Quantity ELSE 0 END) AS Pos8," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=9 THEN Quantity ELSE 0 END) AS Pos9," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=10 THEN Quantity ELSE 0 END) AS Pos10," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=11 THEN Quantity ELSE 0 END) AS Pos11," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=12 THEN Quantity ELSE 0 END) AS Pos12," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=13 THEN Quantity ELSE 0 END) AS Pos13," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=14 THEN Quantity ELSE 0 END) AS Pos14," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=15 THEN Quantity ELSE 0 END) AS Pos15," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=16 THEN Quantity ELSE 0 END) AS Pos16," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=17 THEN Quantity ELSE 0 END) AS Pos17," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=18 THEN Quantity ELSE 0 END) AS Pos18," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=19 THEN Quantity ELSE 0 END) AS Pos19," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=20 THEN Quantity ELSE 0 END) AS Pos20," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=21 THEN Quantity ELSE 0 END) AS Pos21," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=22 THEN Quantity ELSE 0 END) AS Pos22," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=23 THEN Quantity ELSE 0 END) AS Pos23," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=24 THEN Quantity ELSE 0 END) AS Pos24," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=25 THEN Quantity ELSE 0 END) AS Pos25," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=26 THEN Quantity ELSE 0 END) AS Pos26," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=27 THEN Quantity ELSE 0 END) AS Pos27," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=28 THEN Quantity ELSE 0 END) AS Pos28," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=29 THEN Quantity ELSE 0 END) AS Pos29," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=30 THEN Quantity ELSE 0 END) AS Pos30," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=31 THEN Quantity ELSE 0 END) AS Pos31," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=32 THEN Quantity ELSE 0 END) AS Pos32," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=33 THEN Quantity ELSE 0 END) AS Pos33," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=34 THEN Quantity ELSE 0 END) AS Pos34," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=35 THEN Quantity ELSE 0 END) AS Pos35," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=36 THEN Quantity ELSE 0 END) AS Pos36," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=37 THEN Quantity ELSE 0 END) AS Pos37," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=38 THEN Quantity ELSE 0 END) AS Pos38," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=39 THEN Quantity ELSE 0 END) AS Pos39," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=40 THEN Quantity ELSE 0 END) AS Pos40," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=41 THEN Quantity ELSE 0 END) AS Pos41," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=42 THEN Quantity ELSE 0 END) AS Pos42," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=43 THEN Quantity ELSE 0 END) AS Pos43," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=44 THEN Quantity ELSE 0 END) AS Pos44," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=45 THEN Quantity ELSE 0 END) AS Pos45," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=46 THEN Quantity ELSE 0 END) AS Pos46," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=47 THEN Quantity ELSE 0 END) AS Pos47," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=48 THEN Quantity ELSE 0 END) AS Pos48," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=49 THEN Quantity ELSE 0 END) AS Pos49," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=50 THEN Quantity ELSE 0 END) AS Pos50," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=51 THEN Quantity ELSE 0 END) AS Pos51," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=52 THEN Quantity ELSE 0 END) AS Pos52," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=53 THEN Quantity ELSE 0 END) AS Pos53" & Chr(10)
    strSQL = strSQL & "FROM CBIS599p.dbo.POS AS POS" & Chr(10)
    strSQL = strSQL & "JOIN cbis599p.dbo.PRODUCT as PRD ON PRD.PRODUCTCODE = POS.PRODUCTCODE" & Chr(10)
    strSQL = strSQL & "WHERE POS.POSDATE >= '" & Format(DateFrom, "YYYY-MM-DD") & "' AND POS.POSDATE <= '" & Format(DateTo, "YYYY-MM-DD") & "'" & Chr(10)
    strSQL = strSQL & "AND (POS.PRODUCTCODE = " & cType & " OR PRD.CON_PRODUCTCODE = " & cType & ")" & Chr(10)
    strSQL = strSQL & "GROUP BY POS.DIVNO" & Chr(10)
    strSQL = strSQL & "UNION" & Chr(10)
    strSQL = strSQL & "SELECT '599' as DIVNO," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=1 THEN Quantity ELSE 0 END) AS Pos1," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=2 THEN Quantity ELSE 0 END) AS Pos2," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=3 THEN Quantity ELSE 0 END) AS Pos3," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=4 THEN Quantity ELSE 0 END) AS Pos4," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=5 THEN Quantity ELSE 0 END) AS Pos5," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=6 THEN Quantity ELSE 0 END) AS Pos6," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=7 THEN Quantity ELSE 0 END) AS Pos7," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=8 THEN Quantity ELSE 0 END) AS Pos8," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=9 THEN Quantity ELSE 0 END) AS Pos9," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=10 THEN Quantity ELSE 0 END) AS Pos10," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=11 THEN Quantity ELSE 0 END) AS Pos11," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=12 THEN Quantity ELSE 0 END) AS Pos12," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=13 THEN Quantity ELSE 0 END) AS Pos13," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=14 THEN Quantity ELSE 0 END) AS Pos14," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=15 THEN Quantity ELSE 0 END) AS Pos15," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=16 THEN Quantity ELSE 0 END) AS Pos16," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=17 THEN Quantity ELSE 0 END) AS Pos17," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=18 THEN Quantity ELSE 0 END) AS Pos18," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=19 THEN Quantity ELSE 0 END) AS Pos19," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=20 THEN Quantity ELSE 0 END) AS Pos20," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=21 THEN Quantity ELSE 0 END) AS Pos21," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=22 THEN Quantity ELSE 0 END) AS Pos22," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=23 THEN Quantity ELSE 0 END) AS Pos23," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=24 THEN Quantity ELSE 0 END) AS Pos24," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=25 THEN Quantity ELSE 0 END) AS Pos25," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=26 THEN Quantity ELSE 0 END) AS Pos26," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=27 THEN Quantity ELSE 0 END) AS Pos27," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=28 THEN Quantity ELSE 0 END) AS Pos28," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=29 THEN Quantity ELSE 0 END) AS Pos29," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=30 THEN Quantity ELSE 0 END) AS Pos30," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=31 THEN Quantity ELSE 0 END) AS Pos31," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=32 THEN Quantity ELSE 0 END) AS Pos32," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=33 THEN Quantity ELSE 0 END) AS Pos33," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=34 THEN Quantity ELSE 0 END) AS Pos34," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=35 THEN Quantity ELSE 0 END) AS Pos35," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=36 THEN Quantity ELSE 0 END) AS Pos36," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=37 THEN Quantity ELSE 0 END) AS Pos37," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=38 THEN Quantity ELSE 0 END) AS Pos38," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=39 THEN Quantity ELSE 0 END) AS Pos39," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=40 THEN Quantity ELSE 0 END) AS Pos40," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=41 THEN Quantity ELSE 0 END) AS Pos41," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=42 THEN Quantity ELSE 0 END) AS Pos42," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=43 THEN Quantity ELSE 0 END) AS Pos43," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=44 THEN Quantity ELSE 0 END) AS Pos44," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=45 THEN Quantity ELSE 0 END) AS Pos45," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=46 THEN Quantity ELSE 0 END) AS Pos46," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=47 THEN Quantity ELSE 0 END) AS Pos47," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=48 THEN Quantity ELSE 0 END) AS Pos48," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=49 THEN Quantity ELSE 0 END) AS Pos49," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=50 THEN Quantity ELSE 0 END) AS Pos50," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=51 THEN Quantity ELSE 0 END) AS Pos51," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=52 THEN Quantity ELSE 0 END) AS Pos52," & Chr(10)
    strSQL = strSQL & "SUM(CASE WHEN DATEPART(""WK"",POS.POSDATE)=53 THEN Quantity ELSE 0 END) AS Pos53" & Chr(10)
    strSQL = strSQL & "FROM CBIS599p.dbo.POS AS POS" & Chr(10)
    strSQL = strSQL & "JOIN cbis599p.dbo.PRODUCT as PRD ON PRD.PRODUCTCODE = POS.PRODUCTCODE" & Chr(10)
    strSQL = strSQL & "WHERE POS.POSDATE >= '" & Format(DateFrom, "YYYY-MM-DD") & "' AND POS.POSDATE <= '" & Format(DateTo, "YYYY-MM-DD") & "'" & Chr(10)
    strSQL = strSQL & "AND (POS.PRODUCTCODE = " & cType & " OR PRD.CON_PRODUCTCODE = " & cType & ")) AS QRY" & Chr(10)
    If cCG > 0 Then strSQL = strSQL & "WHERE DIVNO = '" & cCG & "'" & Chr(10)
    strSQL = strSQL & "ORDER BY QRY.DIVNO ASC" & Chr(10)
    
    ''Debug.Print strSQL
    
    End If
    
    If SQLType = "ABI_SEASONDATES" Then
    
    strSQL = "SELECT SS.[" & cType & "FirstOSD], SS.[" & cType & "FinalOSD]" & Chr(10)
    strSQL = strSQL & "FROM SeasonStartOSD as SS" & Chr(10)
    strSQL = strSQL & "where SeasonYear = " & Year(DateFrom) & Chr(10)
    
    'Debug.Print strSQL
    
    End If
    
    If SQLType = "ABI_SEASONDATES_UPDATE" Then
    If UCase(Bypass) = "FINAL" Then
    strSQL = "UPDATE SeasonStartOSD as SS" & Chr(10)
    strSQL = strSQL & "SET SS.[" & cType & "FinalOSD] = '" & DateTo & "'" & Chr(10)
    strSQL = strSQL & "where SeasonYear = " & Year(DateFrom) & Chr(10)
    ElseIf UCase(Bypass) = "FIRST" Then
    strSQL = "UPDATE SeasonStartOSD as SS" & Chr(10)
    strSQL = strSQL & "SET SS.[" & cType & "FirstOSD] = '" & DateTo & "'" & Chr(10)
    strSQL = strSQL & "where SeasonYear = " & Year(DateFrom) & Chr(10)
    Else
    strSQL = ""
    End If
    'Debug.Print strSQL
    End If
    
    
    
    
    If SQLType = "POS_SCURVE" Then
    
    strSQL = "SELECT sum(Quantity)" & Chr(10)
    strSQL = strSQL & "FROM CBIS599p.dbo.POS AS POS" & Chr(10)
    strSQL = strSQL & "JOIN cbis599p.dbo.PRODUCT as PRD ON PRD.PRODUCTCODE = POS.PRODUCTCODE" & Chr(10)
    strSQL = strSQL & "WHERE POS.POSDATE >= '" & Format(DateFrom, "yyyy-mm-dd") & "' AND POS.POSDATE <= '" & Format(DateTo, "yyyy-mm-dd") & "'" & Chr(10)
    strSQL = strSQL & "AND (POS.PRODUCTCODE = " & cType & " OR PRD.CON_PRODUCTCODE = " & cType & ")" & Chr(10)
    If cCG > 500 And cCG < 510 Then strSQL = strSQL & "AND POS.DIVNO = " & cCG
    
    ''Debug.Print strSQL
    
    End If
    
    
    
    
    If SQLType = "POS_QTY" Then
     strSQL = ""
    
        strSQLprt01 = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
        If Bypass = "DivIdent" Then
        strSQLprt02 = "SELECT distinct DIVNO FROM CBIS" & CBA_DBtoQuery & "p.dbo.POS AS POS " & Chr(10)
        ElseIf IsNumeric(cType) = True Then
        strSQLprt02 = "SELECT Distinct POS.DIVNO, SUM(QUANTITY) AS 'POSQTY' FROM CBIS" & CBA_DBtoQuery & "p.dbo.POS AS POS " & Chr(10)
        Else
        strSQLprt02 = "SELECT SUM(QUANTITY) AS 'POSQTY', SUM(RETAIL) AS 'POSRET' FROM CBIS" & CBA_DBtoQuery & "p.dbo.POS AS POS " & Chr(10)
        End If
        strSQLprt03 = "JOIN cbis" & CBA_DBtoQuery & "p.dbo.PRODUCT as PRD ON PRD.PRODUCTCODE = POS.PRODUCTCODE" & Chr(10)
        strSQLprt04 = "WHERE "
        If DateTo <> "" And DateFrom <> "" Then strSQLprt05 = "POS.POSDATE >= '" & Format(DateFrom, "yyyy-mm-dd") & "' AND POS.POSDATE <= '" & Format(DateTo, "yyyy-mm-dd") & "'" & Chr(10)
        If cCG <> 0 Then strSQLprt10 = "AND PRD.CGNO = " & cCG & Chr(10)
        If cCG <> 0 And scg <> "" Then strSQLprt11 = "AND PRD.SCGNO IN (" & Mid(scg, 1, Len(scg) - 1) & ")" * Chr(10)
        If Mid(cType, 1, 7) = "PCLASS:" Then strSQLprt12 = "AND PRD.PRODUCTCLASS IN (" & Mid(cType, 8, 99) & ")" & ")" & Chr(10)
        If IsNumeric(cType) Then
            If Bypass = "DivIdent" Then
                strSQLprt13 = "AND POS.PRODUCTCODE = " & cType & Chr(10)
            Else
                strSQLprt13 = "AND (POS.PRODUCTCODE = " & cType & " OR PRD.CON_PRODUCTCODE = " & cType & ")" & Chr(10)
            End If
        End If
        If Mid(Bypass, 1, 4) = "DIV:" Then strSQLprt14 = "AND DIVNO IN (" & Mid(Bypass, 2, Len(Bypass) - 1) & ")" & Chr(10)
        strSQLprt20 = "GROUP BY POS.PRODUCTCODE" & Chr(10)
        strSQLprt21 = "GROUP BY POS.DIVNO" & Chr(10)
        strSQLprt22 = "ORDER BY POS.DIVNO ASC" & Chr(10)
        
        
        
        strSQL = strSQLprt01 & strSQLprt02
        If cCG <> 0 Or scg <> "" Or Mid(Bypass, 1, 7) = "PCLASS:" Or Bypass = "DivIdent" Or IsNumeric(PCode) Then strSQL = strSQL & strSQLprt03
        strSQL = strSQL & strSQLprt04
        If DateTo <> "" And DateFrom <> "" Then strSQL = strSQL & strSQLprt05
        If cCG <> 0 Then strSQL = strSQL & strSQLprt10
        If scg <> "" Then strSQL = strSQL & strSQLprt11
        If Mid(cType, 1, 7) = "PCLASS:" Then strSQL = strSQL & strSQLprt12
        If IsNumeric(cType) Then strSQL = strSQL & strSQLprt13 & strSQLprt21 & strSQLprt22
        If Mid(Bypass, 1, 4) = "DIV:" Then strSQL = strSQL & strSQLprt14
        'strSQL = strSQL & strSQLprt20 & strSQLprt21
        'Debug.Print strSQL
    End If
    
    If SQLType = "CBIS_PRODSbCGbSCG" Then
            CBA_DBtoQuery = 599
        If cCG = "58" Or cCG = "27" Then
            strSQL = "select distinct prd.productcode, prd.description" & Chr(10)
            strSQL = strSQL & "from cbis599p.dbo.product as prd" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.contract as cont on cont.productcode = prd.productcode" & Chr(10)
            strSQL = strSQL & "where prd.cgno = " & cCG & Chr(10)
            strSQL = strSQL & "and prd.scgno = " & scg & Chr(10)
            If cType <> "" Then strSQL = strSQL & "and prd.productclass = " & cType & Chr(10)
            strSQL = strSQL & "and isnull(cont.deliveryto,getdate()) > dateadd(""m"",-3,getdate())" & Chr(10)
        Else
            strSQL = "select distinct a.pcode, p.description from (" & Chr(10)
            strSQL = strSQL & "select distinct isnull(prd.con_productcode,prd.productcode) as pcode" & Chr(10)
            strSQL = strSQL & "from cbis599p.dbo.product as prd" & Chr(10)
            strSQL = strSQL & "inner join cbis599p.dbo.contract as cont on cont.productcode = prd.productcode" & Chr(10)
            strSQL = strSQL & "where prd.cgno = " & cCG & Chr(10)
            If scg <> "" Then strSQL = strSQL & "and prd.scgno = " & scg & Chr(10)
            strSQL = strSQL & "and prd.productclass = " & cType & Chr(10)
            strSQL = strSQL & "and isnull(cont.deliveryto,getdate()) > dateadd(m,-3,getdate())" & Chr(10)
            strSQL = strSQL & ") a left join cbis599p.dbo.product p on p.productcode = a.pcode order by a.pcode" & Chr(10)
        End If
            'Debug.Print strSQL
    End If
    
    
    'Debug.Print strSQL
    curCtype = cType
    If strSQL <> "" Then
        If CBA_DBtoQuery = 1 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "ABI_QUERY", CBA_BSA & "LIVE DATABASES\ABI.accdb", CBA_MSAccess, strSQL, 120, , , False)  'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 2 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "ABI_QUERY", CBA_BSA & "LIVE DATABASES\ABI.accdb", CBA_MSAccess, strSQL, 120, , , False)  'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 3 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "COMRADE_QUERY", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLNCLI10", strSQL, 120, , , False) 'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery = 4 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "COMRADE_QUERY", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLNCLI10", strSQL, 120, , , False) 'Runs DB_Connection module to create connection to dtabase and run query
        ElseIf CBA_DBtoQuery > 500 And CBA_DBtoQuery < 510 Then
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "MMS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", strSQL, 120, , , False)   'Runs DB_Connection module to create connection to dtabase and run query
        Else
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CBIS_QUERY", CBA_DBtoQuery & "DBL01", "SQLNCLI10", strSQL, 120, , , False) 'Runs DB_Connection module to create connection to dtabase and run query
        End If
    End If
    curCtype = ""
    If CBA_DBtoQuery = 1 Then
        If bOutput = False Then
            CBA_COM_MATCHGenPullSQL = False
        '    ElseIf (ABIarr(0, 0) = 0 And bOutput = True) Then
        '    GenPullSQL = False
        Else
            CBA_COM_MATCHGenPullSQL = True
        End If
    Else
        If bOutput = False Then
            CBA_COM_MATCHGenPullSQL = False
            'ElseIf (CBISarr(0, 0) = 0 And bOutput = True) Then
            'GenPullSQL = False
        Else
            CBA_COM_MATCHGenPullSQL = True
        End If
    End If
    If bOutput = False Then
        CBA_COM_MATCHGenPullSQL = False
    Else
        CBA_DB_Connect.CBA_DB_CBADBUpdate "UPDATE", "CBADB_QUERY", CBA_BSA & "LIVE DATABASES\CBADB.accdb", CBA_MSAccess, _
            "INSERT INTO CBA_QueryLogging ([DateTimeStarted], [DateTimeEnded], [CBA_SQLType], [User], [DBtoQuery])" & Chr(10) & _
            "VALUES ('" & CBA_QueryTimerStart & "' , '" & Now & "', '" & SQLType & "' , '" & CBA_BasicFunctions.CBA_Current_UserID & "', " & CBA_DBtoQuery & ")"
        CBA_COM_MATCHGenPullSQL = True
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_COM_MATCHGenPullSQL", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function


