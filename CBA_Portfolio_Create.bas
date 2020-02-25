Attribute VB_Name = "CBA_Portfolio_Create"
Option Explicit
Private CBA_PortfolioDictionary As Scripting.Dictionary
Private Const NATIONAL_DB As Long = 599
Private pCGno As Long
Private PPortfolioID As String
Private PPfVersionID As Long
Private bCreated As Boolean

Function get_CBA_PortfolioDictionary() As Scripting.Dictionary
    Set get_CBA_PortfolioDictionary = CBA_PortfolioDictionary
    If Not bCreated Then create_CBA_PortfolioDictionary
    Set get_CBA_PortfolioDictionary = CBA_PortfolioDictionary
End Function
Private Function create_CBA_PortfolioDictionary()
Dim RS As ADODB.Recordset
Dim CN As ADODB.Connection
Dim strSQL As String, PortID  As String
Dim arr As Variant, a As Long, b As Long, c As Long, PfVersionArr
Dim DicPF As Scripting.Dictionary, DicPFV As Scripting.Dictionary, dic As Scripting.Dictionary, DicMap As Scripting.Dictionary
Dim port As CBA_PfVersion, ti, strPFs As String
Dim TotRows As Long, curPF As String, PercentAmountDisplayed, RateOfRun As Long, TopNo As Long

Set dic = New Scripting.Dictionary
Set CN = New ADODB.Connection

If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Building Portfolio Data"
If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Querying CBIS for Portfolio Data"
With CN
    .ConnectionTimeout = 50
    .CommandTimeout = 50
    .Open "Provider=SQLNCLI10;DATA SOURCE=" & NATIONAL_DB & "DBL01;;INTEGRATED SECURITY=sspi;"
End With
strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
If pCGno > 0 Then strSQL = strSQL & "declare @CG nvarchar(max) = '" & pCGno & "'" & Chr(10)
If PPortfolioID <> "" Then strSQL = strSQL & "declare @PFID nvarchar(max) = '" & PPortfolioID & "'" & Chr(10)
'If PPfVersionID > -1 Then strSQL = strSQL & "declare @PFVID int = " & PPfVersionID & Chr(10)
strSQL = strSQL & "select p.PortfolioID, p.PortfolioNo, p.Sp_Productcode, p.EmpNo, bd.Name as BDName, bd.Firstname as bdFName, p.Assistant, ass.Name as BAName, ass.Firstname as BAFName" & Chr(10)
strSQL = strSQL & ", p.CreationDate, p.ModifiedDate,  p.CGNo as LCGCGNo, isnull(p.SCGNo,0) as LCGSCGNo, p.ACGEntityID, acg.CatNo, acg.Category, acg.CGNo as ACGCGno, acg.CommodityGroup" & Chr(10)
strSQL = strSQL & ", acg.SCGNo  as ACGSCGno, acg.SubCommodityGroup, cg.Description as LCGCGDesc, scg.Description as LCGSCGDesc" & Chr(10)
strSQL = strSQL & "into #PORTHEAD from cbis599p.portfolio.portfolio p" & Chr(10)
strSQL = strSQL & "left join cbis599p.dbo.employee as bd on bd.empno = p.EmpNo" & Chr(10)
strSQL = strSQL & "left join cbis599p.dbo.employee as ass on ass.empno = p.Assistant" & Chr(10)
strSQL = strSQL & "left join cbis599p.dbo.tf_ACGMap() as acg on acg.ACGEntityID = p.ACGEntityID" & Chr(10)
strSQL = strSQL & "left join cbis599p.dbo.COMMODITYGROUP as cg on cg.cgno = p.cgno" & Chr(10)
strSQL = strSQL & "left join cbis599p.dbo.SUBCOMMODITYGROUP as scg on scg.CGNo = p.cgno and scg.scgno = p.scgno" & Chr(10)
strSQL = strSQL & " where p.PortfolioID = @PFID" & Chr(10)
strSQL = strSQL & "select pfv.PortfolioID, pfv.PfVersionID, pfv.productcode, pfv.ContractNo into #PFVM from cbis599p.portfolio.PfVersionMapping pfv where pfv.PortfolioID = @PFID" & Chr(10)
strSQL = strSQL & "" & Chr(10)
strSQL = strSQL & "select PortfolioID, PfVersionID, PfStatusID, Advertisingdate, SpecialCatergoryno, Comp1PortfolioID, Comp1PfVersionID, Comp2PortfolioID, Comp2PfVersionID" & Chr(10)
strSQL = strSQL & ", deliveryfrom, Deliveryto,seasonid, empno, taxid, productclass, packsize, packunit, MinOrderQty, cost, CostCurrency, isnull(retailaftermd,retail) as retail" & Chr(10)
strSQL = strSQL & ", quantity, SuggestedQty, leadtime,StorageTempFrom, StorageTempTo,CostPickup, PriceBase, RegionID" & Chr(10)
strSQL = strSQL & "into #PFVS from cbis599p.portfolio.PfVersionReg p" & Chr(10)
If PPortfolioID <> "" Then strSQL = strSQL & "where PortfolioID = @PFID" & Chr(10)
strSQL = strSQL & "" & Chr(10)
strSQL = strSQL & "select s.PortfolioID, s.PfVersionID, s.Comp1PortfolioID, s.Comp1PfVersionID, s.Comp2PortfolioID, s.Comp2PfVersionID into #FILTER" & Chr(10)
strSQL = strSQL & "from #PFVS s left join #PORTHEAD h on h.PortfolioID = s.PortfolioID" & Chr(10)
If pCGno > 0 Then
    strSQL = strSQL & "where  h.LCGCGNo in (@CG)" & Chr(10)
    If PPortfolioID <> "" Then strSQL = strSQL & "and s.PortfolioID = @PFID" & Chr(10)
Else
    If PPortfolioID <> "" Then strSQL = strSQL & "where s.PortfolioID = @PFID" & Chr(10)
End If
strSQL = strSQL & "" & Chr(10)
strSQL = strSQL & "select s.* into #PFVR from #PFVS s inner join #FILTER f on f.PortfolioID = s.PortfolioID and f.PfVersionID = s.PfVersionID" & Chr(10)
strSQL = strSQL & "union select s.* from #PFVS s inner join #FILTER f on f.Comp1PortfolioID = s.PortfolioID and f.Comp1PfVersionID = s.PfVersionID" & Chr(10)
strSQL = strSQL & "union select s.* from #PFVS s inner join #FILTER f on f.Comp2PortfolioID = s.PortfolioID and f.Comp2PfVersionID = s.PfVersionID" & Chr(10)
strSQL = strSQL & "" & Chr(10)
strSQL = strSQL & "select rpfvr.Portfolioid, rpfvr.PfVersionid, rpfvr.RegionID, rpfvr.Currencycode, rpfvr.Productcode, rpfvr.cost1, rpfvr.retail1, rpfvr.NetRetail1, rpfvr.Quantity1, rpfvr.GroupAdvdate" & Chr(10)
strSQL = strSQL & "into #RPFVR from cbis599p.portfolio.Rep_PfVersionReg rpfvr inner join #PFVR p on p.PortfolioID = rpfvr.Portfolioid and p.PfVersionID = rpfvr.PfVersionid" & Chr(10)
strSQL = strSQL & "" & Chr(10)
strSQL = strSQL & "select pfvl.PortfolioID, pfvl.PfVersionID, pfvl.Description into #PFVL from cbis599p.portfolio.PfVersionLng pfvl" & Chr(10)
strSQL = strSQL & "inner join #PFVR p on p.PortfolioID = pfvl.PortfolioID and p.PfVersionID = pfvl.PfVersionID where pfvl.LanguageID = 0" & Chr(10)
strSQL = strSQL & "" & Chr(10)
strSQL = strSQL & "select cp.portfolioid, cp.pfversionid, cp.CmpProductCodeAsString into #CP from cbis599p.portfolio.ComparisonProducts cp inner join #FILTER f on f.PortfolioID = cp.PortfolioID and f.PfVersionID = cp.PfVersionID where cp.ServerID = 5" & Chr(10)
strSQL = strSQL & "" & Chr(10)
strSQL = strSQL & "select pf.PortfolioID,  pf.PfVersionID, pfvl.Description,   pfvm.productcode,   pfvm.ContractNo,    isnull(rpfvr.GroupAdvdate,pfvr.Advertisingdate) as OSD" & Chr(10)
strSQL = strSQL & ", ph.LCGCGNo, ph.LCGSCGNo, ph.ACGEntityID,  ph.CatNo,   ph.Category,    ph.ACGCGno, ph.CommodityGroup,  ph.ACGSCGno,    ph.SubCommodityGroup,   pfvr.seasonid,  pfvr.productclass,  pfvr.packsize" & Chr(10)
strSQL = strSQL & ", rpfvr.cost1, rpfvr.retail1,   rpfvr.NetRetail1,   rpfvr.Quantity1,    pfvr.PfStatusID, ph.BDName, ph.bdFName, ph.BAName,  ph.BAFName" & Chr(10)
strSQL = strSQL & ", pfvr.SpecialCatergoryno,  pfvr.deliveryfrom,  pfvr.Deliveryto,    pfvr.taxid, pfvr.packunit,  pfvr.MinOrderQty,   pfvr.cost,  pfvr.CostCurrency" & Chr(10)
strSQL = strSQL & ", pfvr.retail,  pfvr.quantity,  pfvr.SuggestedQty,  pfvr.leadtime,  pfvr.CostPickup,    pfvr.PriceBase, pfvr.RegionID,  rpfvr.Currencycode" & Chr(10)
strSQL = strSQL & ", pf.Comp1PortfolioID,  pf.Comp1PfVersionID,    pf.Comp2PortfolioID,    pf.Comp2PfVersionID, ph.LCGCGDesc, ph.LCGSCGDesc, cp.CmpProductCodeAsString as CompCode" & Chr(10)
strSQL = strSQL & "into #BASE from #FILTER pf" & Chr(10)
strSQL = strSQL & "left join #PORTHEAD ph on ph.PortfolioID = pf.PortfolioID" & Chr(10)
strSQL = strSQL & "left join #PFVR pfvr on pfvr.PortfolioID = pf.PortfolioID and pfvr.PfVersionID = pf.PfVersionID" & Chr(10)
strSQL = strSQL & "left join #PFVM pfvm on pfvm.PortfolioID = pf.PortfolioID and pfvm.PfVersionID = pfvr.PfVersionID" & Chr(10)
strSQL = strSQL & "left join #RPFVR rpfvr on rpfvr.PfVersionID = pfvr.PfVersionID and rpfvr.PortfolioID = pfvr.PortfolioID" & Chr(10)
strSQL = strSQL & "left join #PFVL pfvl on pfvl.PfVersionID = pfvr.PfVersionID and pfvl.PortfolioID = pfvr.PortfolioID" & Chr(10)
strSQL = strSQL & "left join #CP cp on cp.PfVersionID = pfvr.PfVersionID and cp.PortfolioID = pfvr.PortfolioID" & Chr(10)
If PPortfolioID <> "" Then strSQL = strSQL & "where pf.PortfolioID = @PFID" & Chr(10)
'If PPfVersionID > -1 Then strSQL = strSQL & "and pf.PfVersionID = @PFVID" & Chr(10)
strSQL = strSQL & "order by pf.PortfolioID, pf.PfVersionID, pfvm.productcode,pfvm.ContractNo" & Chr(10)
strSQL = strSQL & "" & Chr(10)
strSQL = strSQL & "select distinct PortfolioID, PfVersionID into #IDX from #BASE order by PortfolioID, PfVersionID" & Chr(10)
strSQL = strSQL & "drop table #PORTHEAD,#PFVM ,#PFVS, #FILTER, #PFVR, #RPFVR, #PFVL" & Chr(10)
Set RS = New ADODB.Recordset
RS.Open strSQL, CN


strSQL = "select PortfolioID, max(isnull(pfVersionID,0)) from #IDX group by PortfolioID  order by PortfolioID"
Set RS = New ADODB.Recordset
RS.Open strSQL, CN
arr = RS.GetRows()

Set DicMap = New Scripting.Dictionary
c = 0: curPF = "": Set DicPF = New Scripting.Dictionary
TotRows = UBound(arr, 2): PercentAmountDisplayed = 0
RateOfRun = 199
ti = Timer
For a = LBound(arr, 2) To TotRows
    'ti = Timer
    If a = c Then
        strPFs = ""
        TopNo = IIf(a + RateOfRun > TotRows, TotRows, a + RateOfRun)
        For b = a To TopNo
            If b = TopNo Then c = b + 1
            If strPFs = "" Then strPFs = "'" & arr(0, b) & "'" Else strPFs = strPFs & ", '" & arr(0, b) & "'"
            DicMap.Add arr(0, b), CStr(arr(1, b))
        Next
        strSQL = "select * from #BASE where PortfolioID in (" & strPFs & ")  order by  PortfolioID, PfVersionID, productcode"
        Set RS = New ADODB.Recordset: RS.Open strSQL, CN
        PfVersionArr = RS.GetRows()
    End If
    Set DicPFV = New Scripting.Dictionary
    For b = 0 To DicMap(arr(0, a))
        Set port = New CBA_PfVersion
        port.Generate PfVersionArr, arr(0, a), CByte(b)
        DicPFV.Add CStr(b), port
    Next
    DicPF.Add arr(0, a), DicPFV
    If TotRows > 0 Then
        If Round(a / TotRows, 2) > PercentAmountDisplayed Then
            PercentAmountDisplayed = Round(a / TotRows, 2)
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Building Portfolio DataCube: " & Format(PercentAmountDisplayed, "0%")
            DoEvents
            ti = Timer
        End If
    End If
Next
'If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running

Set CBA_PortfolioDictionary = DicPF
RS.Close
Set RS = Nothing
CN.Close
Set CN = Nothing
Set dic = Nothing

End Function
Function set_CGno(ByVal l_CGno As Long) As Boolean
    pCGno = l_CGno
    set_CGno = True
End Function
Function set_PortfolioID(ByVal s_PortfolioID As String) As Boolean
    PPortfolioID = s_PortfolioID
    set_PortfolioID = True
End Function
Function set_PfVersionID(ByVal l_PFVersionID As Long) As Boolean
    If l_PFVersionID = -1 Then PPfVersionID = Empty Else PPfVersionID = l_PFVersionID
    set_PfVersionID = True
End Function
Sub TESTERORPORT()
    Dim PfVer As CBA_PfVersion, i
    'set_PortfolioID "{8D04547B-281D-42F3-8AF0-000DCB1464F1}"
    'set_PfVersionID 1
    set_PortfolioID ""
    set_PfVersionID -1
    set_CGno 12
    
    create_CBA_PortfolioDictionary
    If CBA_PortfolioDictionary.Exists("{F3A2894C-5E40-401D-85AE-A9C87DDDAA00}") And IsEmpty(CBA_PortfolioDictionary("{F3A2894C-5E40-401D-85AE-A9C87DDDAA00}")) = False Then
        Set PfVer = CBA_PortfolioDictionary("{F3A2894C-5E40-401D-85AE-A9C87DDDAA00}")("2")
        Debug.Print PfVer.LCGCommodityGroup
        Debug.Print PfVer.LCGSubCommodityGroup
        For Each i In PfVer.getRetail()
            Debug.Print i, PfVer.getRetail(i)(i)
        Next
        Debug.Print PfVer.getRetail("030932")("030932")
        Debug.Print PfVer.getMargin("030932")("030932")
    Else
    
    End If
    
    
End Sub

