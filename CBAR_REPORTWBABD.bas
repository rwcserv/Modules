Attribute VB_Name = "CBAR_REPORTWBABD"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Type SubBasket
    alcohol As Single
    ambientnonfood As Single
    ambientfood As Single
    chilled As Single
    frozen As Single
    produce As Single
    meat As Single
End Type
Private CBA_COM_CBISRS As ADODB.Recordset

Sub WBA(ByVal StateLookup As String, ByVal DFrom As Date, ByVal Dto As Date, ByVal DisplayData As String, ByRef GraphsToMake As Collection, ByVal CGno As Long, ByVal OldM As Boolean)
    Dim curRetail() As Single, curAldiRetail() As Single, bfound As Boolean, prod, totComp, totAldi, strChartname As String, Point, PointRow
    Dim ThisProd As Long, strSQL As String ', dfrom As Date, dto As Date
    Dim Retails() As SubBasket, AldiRetails() As SubBasket, counter() As SubBasket, CompsToInc As New Collection, colexcp() As Collection, colMatchused() As Collection
    Dim CompstoExcludefromBranded As Collection, WeekstoCollect As Collection, compet, Compt, myprod, retARR, Aretarr, ScrapeDate
    Dim totProdShare() As Single, AldiProdsbyPeriod() As Long
    Dim sCurMatchcode As String, sCurMatchDesc As String, sCheckMatchUsed As String
    Dim lNum As Long, NumMonths, NoOfMonths, ProdCntish, a As Long, b As Long, c As Long, j As Long, v As Long, curProd, strAldiProdList, ExProd, bUse As Boolean, ForEntry As Boolean
    Dim wbk_WBA, wks_WBA, wks_output, cht, NumForNextStart, Around, General, dset, arrofDates, RowTopper, AldiProdDesc, Afcnt, Anfcnt, Ccnt, Fcnt, Mecnt
    Dim SourceCellNoEnd, SourceCellNo, CatCellNo, SRng, LRng, CRng, ThisChart
    Dim numWeeks As Long
    Dim dayer  As Long
    Dim arrayWeekStart As Long, arrayWeekend As Long
    Dim WeekDic As Scripting.Dictionary
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Creating Weighted Basket Report"
    
    NumMonths = ((Year(Dto) - Year(DFrom)) * 12) + (Month(Dto) - Month(DFrom)) + 1
    NoOfMonths = Month(Dto) - Month(DFrom) + 1
    numWeeks = 0
    Set WeekDic = New Scripting.Dictionary
    For dayer = 0 To DateDiff("D", DFrom, Dto)
        If WeekDay(DateAdd("D", dayer, DFrom), vbWednesday) = 1 Then
            numWeeks = numWeeks + 1
            WeekDic.Add DateAdd("D", dayer, DFrom), numWeeks
        End If
    Next
    
    If CBA_COM_SetupMatchArray.CBA_SetupMatchArray(OldM, DFrom, Dto, CGno) = True Then
     
        'identify last week in each month for monthly reporting, add to collection
        ProdCntish = 0
        For a = 1 To UBound(CBA_COM_Match)
            If a = 1 Then
                curProd = CBA_COM_Match(a).AldiPCode
                strAldiProdList = curProd
            ElseIf curProd <> CBA_COM_Match(a).AldiPCode Then
                ProdCntish = ProdCntish + 1
                curProd = CBA_COM_Match(a).AldiPCode
                strAldiProdList = strAldiProdList & ", " & curProd
            End If
        Next
        
        Set CBA_COM_CBISCN = New ADODB.Connection
        With CBA_COM_CBISCN
            .ConnectionTimeout = 100
            .CommandTimeout = 300
            .Open "Provider= SQLNCLI10; DATA SOURCE= 599DBL01; ;INTEGRATED SECURITY=sspi;"
        End With
        Set CBA_COM_CBISRS = New ADODB.Recordset
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
        strSQL = strSQL & "declare @datefrom as date" & Chr(10)
        strSQL = strSQL & "declare @dateto as date" & Chr(10)
        strSQL = strSQL & "declare @totbus as decimal(18,2)" & Chr(10)
        strSQL = strSQL & "set @datefrom = '" & Format(DFrom, "YYYY-MM-DD") & "'" & Chr(10)
        strSQL = strSQL & "set @dateto = '" & Format(Dto, "YYYY-MM-DD") & "'" & Chr(10)
        strSQL = strSQL & "select productcode into #Prods from cbis599p.dbo.product where productcode in (" & strAldiProdList & ")  and productclass = 1 order by productcode" & Chr(10)
        If LCase(DisplayData) = "months" Then
            strSQL = strSQL & "select month(posdate) as prd, sum(retail) as totbus into #Totbus from cbis599p.dbo.pos pos where posdate >= @datefrom and posdate <= @dateto group by month(posdate)" & Chr(10)
            strSQL = strSQL & "select month(posdate),isnull(p.con_productcode,p.productcode), round((sum(retail) / tb.totbus) * 100,2) from cbis599p.dbo.pos pos" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.product p on p.productcode = pos.productcode" & Chr(10) & "inner join #Prods pr on pr.productcode = isnull(p.con_productcode,p.productcode)" & Chr(10)
            strSQL = strSQL & "left join #totbus tb on tb.prd = month(posdate)" & Chr(10)
            strSQL = strSQL & "where posdate >= @datefrom and posdate <= @dateto " & Chr(10)
            strSQL = strSQL & "group by  month(posdate), isnull(p.con_productcode,p.productcode), tb.totbus  order by isnull(p.con_productcode,p.productcode)" & Chr(10)
        ElseIf LCase(DisplayData) = "years" Then
            strSQL = strSQL & "select year(posdate) as prd, sum(retail) as totbus into #Totbus from cbis599p.dbo.pos pos where posdate >= @datefrom and posdate <= @dateto group by year(posdate)" & Chr(10)
            strSQL = strSQL & "select year(posdate),isnull(p.con_productcode,p.productcode), round((sum(retail) / tb.totbus) * 100,2) from cbis599p.dbo.pos pos" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.product p on p.productcode = pos.productcode" & Chr(10) & "inner join #Prods pr on pr.productcode = isnull(p.con_productcode,p.productcode)" & Chr(10)
            strSQL = strSQL & "left join #totbus tb on tb.prd = year(posdate)" & Chr(10)
            strSQL = strSQL & "where posdate >= @datefrom and posdate <= @dateto" & Chr(10)
            strSQL = strSQL & "group by  year(posdate), isnull(p.con_productcode,p.productcode), tb.totbus  order by isnull(p.con_productcode,p.productcode)" & Chr(10)
        ElseIf LCase(DisplayData) = "weeks" Then
            strSQL = strSQL & "select datepart(iso_week,posdate) as prd, sum(retail) as totbus, datepart(YEAR,posdate) as yrno into #Totbus from cbis599p.dbo.pos pos where posdate >= @datefrom and posdate <= @dateto group by datepart(iso_week,posdate), datepart(YEAR,posdate)" & Chr(10)
            strSQL = strSQL & "select datepart(iso_week,posdate),isnull(p.con_productcode,p.productcode), round((sum(retail) / tb.totbus) * 100,2) , datepart(YEAR,posdate) from cbis599p.dbo.pos pos" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.product p on p.productcode = pos.productcode" & Chr(10) & "inner join #Prods pr on pr.productcode = isnull(p.con_productcode,p.productcode)" & Chr(10)
            strSQL = strSQL & "left join #totbus tb on tb.prd = datepart(iso_week,posdate) and tb.yrno = datepart(YEAR,posdate)" & Chr(10)
            strSQL = strSQL & "where posdate >= @datefrom and posdate <= @dateto" & Chr(10)
            strSQL = strSQL & "group by  datepart(iso_week,posdate), isnull(p.con_productcode,p.productcode), tb.totbus , datepart(YEAR,posdate) order by isnull(p.con_productcode,p.productcode) " & Chr(10)
        ElseIf LCase(DisplayData) = "total period" Then
            strSQL = strSQL & "select 0 as prd, sum(retail) as totbus into #Totbus from cbis599p.dbo.pos pos where posdate >= @datefrom and posdate <= @dateto " & Chr(10)
            strSQL = strSQL & "select 0 as prds, isnull(p.con_productcode,p.productcode), round((sum(retail) / tb.totbus) * 100,2) from cbis599p.dbo.pos pos" & Chr(10)
            strSQL = strSQL & "left join cbis599p.dbo.product p on p.productcode = pos.productcode" & Chr(10) & "inner join #Prods pr on pr.productcode = isnull(p.con_productcode,p.productcode)" & Chr(10)
            strSQL = strSQL & "left join #totbus tb on tb.prd = prds" & Chr(10)
            strSQL = strSQL & "where posdate >= @datefrom and posdate <= @dateto" & Chr(10)
            strSQL = strSQL & "group by isnull(p.con_productcode,p.productcode), tb.totbus order by isnull(p.con_productcode,p.productcode)" & Chr(10)
        End If
        strSQL = strSQL & "drop table #Prods, #Totbus"
    '  '  Debug.Print strSQL
        CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN
        If CBA_COM_CBISRS.EOF Or CBA_COM_CBISRS.BOF Then
            ReDim CBISarr(0, 0): CBISarr(0, 0) = 0
        Else
            CBISarr = CBA_COM_CBISRS.GetRows
            CBA_COM_CBISRS.Close
        End If
        Set CBA_COM_CBISRS = Nothing
           
        If LCase(DisplayData) = "months" Then
            For a = Month(DFrom) To Month(Dto)
                Set CBA_COM_CBISRS = New ADODB.Recordset
                CBA_COM_CBISRS.CursorLocation = adUseClient
                strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
                strSQL = strSQL & "declare @startdate as date declare @enddate as date set @startdate = '" & Format(DateAdd("m", a - Month(DFrom), DFrom), "YYYY-MM-DD") & "' set @enddate = '" & Format(DateSerial(Year(DateAdd("m", (a + 1) - Month(DFrom), DFrom)), Month(DateAdd("m", (a + 1) - Month(DFrom), DFrom)), -1), "YYYY-MM-DD") & "' select distinct isnull(p.con_productcode,p.productcode) as pcode, p.cgno" & Chr(10)
                strSQL = strSQL & "from cbis599p.dbo.contract c inner join cbis599p.dbo.product p on p.productcode = c.productcode inner join cbis599p.dbo.divcontracthis dc on dc.contractno = c.contractno" & Chr(10)
                strSQL = strSQL & "where not p.cgno in (58, 61, 27) and dc.sendtodiv = 1 and ((p.productclass = 1 and c.deliveryfrom  <= @enddate and isnull(c.deliveryto,getdate()) >=@startdate)" & Chr(10)
                strSQL = strSQL & ") Order by isnull(p.con_productcode,p.productcode)" & Chr(10)
        '      '  Debug.Print strSQL
                CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN
                If CBA_COM_CBISRS.EOF Or CBA_COM_CBISRS.BOF Then
                    ReDim AldiProdsbyPeriod(1, 2999, a): AldiProdsbyPeriod(1, 2999, a) = 0
                Else
                    ReDim Preserve AldiProdsbyPeriod(1 To 2, 1 To 2999, 1 To a)
                    For b = 0 To CBA_COM_CBISRS.RecordCount - 1
                        AldiProdsbyPeriod(1, b + 1, a) = CBA_COM_CBISRS.Fields.Item(0)
                        AldiProdsbyPeriod(2, b + 1, a) = CBA_COM_CBISRS.Fields.Item(1)
                        CBA_COM_CBISRS.MoveNext
                    Next
                End If
                CBA_COM_CBISRS.Close: Set CBA_COM_CBISRS = Nothing
            Next
            ExProd = ""
            For b = Month(DFrom) To Month(Dto)
                For a = 1 To UBound(AldiProdsbyPeriod, 2)
                    If IsNull(AldiProdsbyPeriod(1, a, b)) Then Exit For
                    If ExProd = "" Then ExProd = AldiProdsbyPeriod(1, a, b) Else ExProd = ExProd & ", " & AldiProdsbyPeriod(1, a, b)
                Next
            Next
        ElseIf LCase(DisplayData) = "weeks" Then
        
            arrayWeekStart = 60: arrayWeekend = 0
            For a = LBound(CBISarr, 2) To UBound(CBISarr, 2)
                If Year(DFrom) = CBISarr(3, a) And arrayWeekStart > CBISarr(0, a) Then arrayWeekStart = CBISarr(0, a)
                If Year(Dto) = CBISarr(3, a) And arrayWeekend < CBISarr(0, a) Then arrayWeekend = CBISarr(0, a)
            Next
            
            If Year(Dto) = Year(DFrom) Then
                For a = 1 To 20
                    If (arrayWeekend - arrayWeekStart) >= numWeeks Then
                        arrayWeekStart = arrayWeekStart + 1
                    Else
                        Exit For
                    End If
                Next
            Else
                MsgBox "Please contact Tom or Bob on 9218 as this report is not built to handle straddling years at the moment. Thanks"
                CBA_COM_CBISCN.Close: Set CBA_COM_CBISCN = Nothing
                If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
                Application.DisplayAlerts = False
                ActiveWorkbook.Close
                Application.DisplayAlerts = True
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
            End If
            
            
            For a = 1 To numWeeks
                Set CBA_COM_CBISRS = New ADODB.Recordset
                CBA_COM_CBISRS.CursorLocation = adUseClient
                strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10)
                strSQL = strSQL & "declare @startdate as date declare @enddate as date set @startdate = '" & Format(DateAdd("m", a - Month(DFrom), DFrom), "YYYY-MM-DD") & "' set @enddate = '" & Format(DateSerial(Year(DateAdd("m", (a + 1) - Month(DFrom), DFrom)), Month(DateAdd("m", (a + 1) - Month(DFrom), DFrom)), -1), "YYYY-MM-DD") & "' select distinct isnull(p.con_productcode,p.productcode) as pcode, p.cgno" & Chr(10)
                strSQL = strSQL & "from cbis599p.dbo.contract c inner join cbis599p.dbo.product p on p.productcode = c.productcode inner join cbis599p.dbo.divcontracthis dc on dc.contractno = c.contractno" & Chr(10)
                strSQL = strSQL & "where not p.cgno in (58, 61, 27) and dc.sendtodiv = 1 and ((p.productclass = 1 and c.deliveryfrom  <= @enddate and isnull(c.deliveryto,getdate()) >=@startdate)" & Chr(10)
                strSQL = strSQL & ") Order by isnull(p.con_productcode,p.productcode)" & Chr(10)
        '      '  Debug.Print strSQL
                CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN
                If CBA_COM_CBISRS.EOF Or CBA_COM_CBISRS.BOF Then
                    ReDim AldiProdsbyPeriod(1, 2999, a): AldiProdsbyPeriod(1, 2999, a) = 0
                Else
                    ReDim Preserve AldiProdsbyPeriod(1 To 2, 1 To 2999, 1 To a)
                    For b = 0 To CBA_COM_CBISRS.RecordCount - 1
                        AldiProdsbyPeriod(1, b + 1, a) = CBA_COM_CBISRS.Fields.Item(0)
                        AldiProdsbyPeriod(2, b + 1, a) = CBA_COM_CBISRS.Fields.Item(1)
                        CBA_COM_CBISRS.MoveNext
                    Next
                End If
                CBA_COM_CBISRS.Close: Set CBA_COM_CBISRS = Nothing
            Next
            ExProd = ""
            For b = 1 To numWeeks
                For a = 1 To UBound(AldiProdsbyPeriod, 2)
                    If IsNull(AldiProdsbyPeriod(1, a, b)) Then Exit For
                    If ExProd = "" Then ExProd = AldiProdsbyPeriod(1, a, b) Else ExProd = ExProd & ", " & AldiProdsbyPeriod(1, a, b)
                Next
            Next
        
        
        End If
        
        Set CompstoExcludefromBranded = New Collection
        CompstoExcludefromBranded.Add "colesval"
        CompstoExcludefromBranded.Add "colespl"
        CompstoExcludefromBranded.Add "colescoles"
        CompstoExcludefromBranded.Add "colessb"
        CompstoExcludefromBranded.Add "wwselect"
        CompstoExcludefromBranded.Add "wwww"
        CompstoExcludefromBranded.Add "wwhb"
        CompstoExcludefromBranded.Add "dm1"
        CompstoExcludefromBranded.Add "dm2"
    
        CBAR_WBA.Copy
        Set wbk_WBA = ActiveWorkbook
        Set wks_WBA = wbk_WBA.Sheets(1)
        Set wks_output = wbk_WBA.Worksheets.Add
        wks_output.Name = "output"
        wks_WBA.Activate
        For Each cht In CBAR_WBA.ChartObjects
            cht.Delete
        Next
    
        lNum = 0
        NumForNextStart = 0
        Around = 0
        Range(wks_WBA.Cells(1, 10), wks_WBA.Cells(1, 20)).EntireColumn.NumberFormat = General
        wks_output.Cells.ClearContents
        For Each dset In GraphsToMake
            ''wks_run.Cells(14, 6).Value = "Generating graph: " & dset & "...."
            DoEvents
            Around = Around + 1
            Set CompsToInc = New Collection
            If dset = "ColesVal" Or dset = "ColesCOS" Or dset = "All" Then CompsToInc.Add "colesval"
            If dset = "ColesPL" Or dset = "ColesColes" Or dset = "ColesCOS" Or dset = "ColesS" Or dset = "All" Then CompsToInc.Add "colespl"
            If dset = "ColesPL" Or dset = "ColesColes" Or dset = "ColesCOS" Or dset = "ColesS" Or dset = "All" Then CompsToInc.Add "colescoles"
            If dset = "ColesSB" Or dset = "ColesCOS" Or dset = "All" Then CompsToInc.Add "colessb"
            If dset = "WWSelect" Or dset = "WWCOS" Or dset = "All" Then CompsToInc.Add "wwselect"
            If dset = "WWSelect" Or dset = "WWCOS" Or dset = "All" Then CompsToInc.Add "wwww"
            If dset = "WWSelect" Or dset = "WWCOS" Or dset = "All" Then CompsToInc.Add "wwpl"
            If dset = "WWHB" Or dset = "WWCOS" Or dset = "All" Then CompsToInc.Add "wwhb"
            If dset = "DM1" Or dset = "DMCOS" Or dset = "All" Then CompsToInc.Add "dm1"
            If dset = "DM2" Or dset = "DMCOS" Or dset = "All" Then CompsToInc.Add "dm2"
            If dset = "DMQ" Or dset = "DMCOS" Or dset = "All" Then CompsToInc.Add "dmq"
            If dset = "FC1" Or dset = "FCCOS" Or dset = "All" Then CompsToInc.Add "fc1"
            If dset = "FC2" Or dset = "FCCOS" Or dset = "All" Then CompsToInc.Add "fc2"
            If dset = "FCQ" Or dset = "FCCOS" Or dset = "All" Then CompsToInc.Add "fcq"
            If dset = "ColesMarketLeaders" Or dset = "ColesBranded" Or dset = "Branded" Or dset = "All" Then CompsToInc.Add "colesml"
            If dset = "ColesControlBrands" Or dset = "ColesBranded" Or dset = "Branded" Or dset = "All" Then CompsToInc.Add "colescb"
            If dset = "WWMarketLeaders" Or dset = "WWBranded" Or dset = "Branded" Or dset = "All" Then CompsToInc.Add "wwml"
            If dset = "WWControlBrands" Or dset = "WWBranded" Or dset = "Branded" Or dset = "All" Then CompsToInc.Add "wwcb"
            If dset = "DMMarketLeaders" Or dset = "DMBranded" Or dset = "Branded" Or dset = "All" Then CompsToInc.Add "dmml"
            If dset = "DMControlBrands" Or dset = "DMBranded" Or dset = "Branded" Or dset = "All" Then CompsToInc.Add "dmcb"
            If dset = "ColesPhantoms" Or dset = "Phantoms" Or dset = "All" Then CompsToInc.Add "colespb"
            If dset = "WWPhantoms" Or dset = "Phantoms" Or dset = "All" Then CompsToInc.Add "wwpb"
            If dset = "DMPhantoms" Or dset = "Phantoms" Or dset = "All" Then CompsToInc.Add "dmpb"
        
            ReDim colexcp(1)
            Set colexcp(1) = New Collection
            ReDim colMatchused(1 To 1)
            Set colMatchused(1) = New Collection
            
            If DisplayData = "months" Then
                ReDim Retails(Month(DFrom) To Month(Dto))
                ReDim AldiRetails(Month(DFrom) To Month(Dto))
                ReDim counter(Month(DFrom) To Month(Dto))
            ElseIf DisplayData = "weeks" Then
                ReDim Retails(1 To numWeeks)
                ReDim AldiRetails(1 To numWeeks)
                ReDim counter(1 To numWeeks)
            End If
          
            For a = 1 To UBound(CBA_COM_Match) + 1
        
                If a <= UBound(CBA_COM_Match) Then
                    Set WeekstoCollect = New Collection
                    arrofDates = CBA_COM_Match(a).getdates
                    If DisplayData = "months" Then
                        For j = LBound(arrofDates) To UBound(arrofDates)
                            If UBound(arrofDates) = j And ((Month(arrofDates(j)) <= Month(Dto) And Year(arrofDates(j)) = Year(Dto)) Or (Month(arrofDates(j)) > Month(Dto) And Year(arrofDates(j)) = Year(Dto) - 1)) Then
                                WeekstoCollect.Add arrofDates(j)
                            Else
                                If j > 1 Then
                                    If Month(arrofDates(j)) > Month(arrofDates(j - 1)) And ((Month(arrofDates(j - 1)) <= Month(Dto) And Year(arrofDates(j - 1)) = Year(Dto)) Or (Month(arrofDates(j - 1)) > Month(Dto) And Year(arrofDates(j - 1)) = Year(Dto) - 1)) Then
                                        WeekstoCollect.Add arrofDates(j - 1)
                                    End If
                                End If
                            End If
                        Next
                    ElseIf DisplayData = "weeks" Then
                        For j = LBound(arrofDates) To UBound(arrofDates)
                              If ((Month(arrofDates(j)) <= Month(Dto) And Year(arrofDates(j)) = Year(Dto)) Or (Month(arrofDates(j)) > Month(Dto) And Year(arrofDates(j)) = Year(Dto) - 1)) Then WeekstoCollect.Add arrofDates(j)
                        Next
                    End If
                End If
                
                
                
                'rowtopper = 0
                If a < UBound(CBA_COM_Match) + 1 Then
                    If InStr(1, LCase(CBA_COM_Match(a).MatchType), "ww") > 0 Then
                        RowTopper = RowTopper + 1
                        wks_output.Cells(RowTopper, 1).Value = CBA_COM_Match(a).AldiPCode & " - " & CBA_COM_Match(a).MatchType
                '      '  Debug.Print CBA_COM_Match(a).AldiPCode & " - " & CBA_COM_Match(a).MatchType & " - " & CBA_COM_Match(a).CompCode
                    End If
                End If

                If a = UBound(CBA_COM_Match) + 1 Then v = a - 1 Else v = a
                If ThisProd <> CBA_COM_Match(v).AldiPCode Or a = UBound(CBA_COM_Match) + 1 Then
            
                    If a > 1 Then
                        For j = LBound(arrofDates) To UBound(arrofDates)
                            If curRetail(j) > 0 Then
                                  Set CBA_COM_CBISRS = New ADODB.Recordset
                                  strSQL = "select description from cbis599p.dbo.product where productcode = " & ThisProd
                                  CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN
                                  AldiProdDesc = CBA_COM_CBISRS.Fields.Item(0)
                                  CBA_COM_CBISRS.Close: Set CBA_COM_CBISRS = Nothing
                                  Select Case CGno
                                  
                                  Case 1 To 4
                                  Case 5, 40 To 50, 52 To 57
                                      Retails(j).ambientfood = Retails(j).ambientfood + curRetail(j)
                                      If curRetail(j) > 0 Then counter(j).ambientfood = counter(j).ambientfood + 1
                                      If curRetail(j) > 0 Then AldiRetails(j).ambientfood = AldiRetails(j).ambientfood + curAldiRetail(j)
                                      Afcnt = Afcnt + 1
                                      wks_output.Cells(Afcnt, 11) = j
                                      wks_output.Cells(Afcnt, 12) = ThisProd
                                      wks_output.Cells(Afcnt, 13) = AldiProdDesc
                                      wks_output.Cells(Afcnt, 14) = sCurMatchcode
                                      wks_output.Cells(Afcnt, 15) = sCurMatchDesc
                                      wks_output.Cells(Afcnt, 16) = curRetail(j)
                                      wks_output.Cells(Afcnt, 17) = curAldiRetail(j)
                                      wks_output.Cells(Afcnt, 18) = totProdShare(j)
                                      If (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j) > 0.5 Or (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j) < -0.5 Then wks_output.Cells(Afcnt, 19) = (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j)
                                  Case 6 To 37, 39, 61, 65
                                      Retails(j).ambientnonfood = Retails(j).ambientnonfood + curRetail(j)
                                      If curRetail(j) > 0 Then counter(j).ambientnonfood = counter(j).ambientnonfood + 1
                                      If curRetail(j) > 0 Then AldiRetails(j).ambientnonfood = AldiRetails(j).ambientnonfood + curAldiRetail(j)
                                      Anfcnt = Anfcnt + 1
                                      wks_output.Cells(Anfcnt, 21) = j
                                      wks_output.Cells(Anfcnt, 22) = ThisProd
                                      wks_output.Cells(Anfcnt, 23) = AldiProdDesc
                                      wks_output.Cells(Anfcnt, 24) = sCurMatchcode
                                      wks_output.Cells(Anfcnt, 25) = sCurMatchDesc
                                      wks_output.Cells(Anfcnt, 26) = curRetail(j)
                                      wks_output.Cells(Anfcnt, 27) = curAldiRetail(j)
                                      wks_output.Cells(Anfcnt, 28) = totProdShare(j)
                                      If (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j) > 0.5 Or (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j) < -0.5 Then wks_output.Cells(Anfcnt, 29) = (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j)
                                  Case 51
                                      Retails(j).chilled = Retails(j).chilled + curRetail(j)
                                      If curRetail(j) > 0 Then counter(j).chilled = counter(j).chilled + 1
                                      If curRetail(j) > 0 Then AldiRetails(j).chilled = AldiRetails(j).chilled + curAldiRetail(j)
                                      Ccnt = Ccnt + 1
                                      wks_output.Cells(Ccnt, 31) = j
                                      wks_output.Cells(Ccnt, 32) = ThisProd
                                      wks_output.Cells(Ccnt, 33) = AldiProdDesc
                                      wks_output.Cells(Ccnt, 34) = sCurMatchcode
                                      wks_output.Cells(Ccnt, 35) = sCurMatchDesc
                                      wks_output.Cells(Ccnt, 36) = curRetail(j)
                                      wks_output.Cells(Ccnt, 37) = curAldiRetail(j)
                                      wks_output.Cells(Ccnt, 38) = totProdShare(j)
                                      If (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j) > 0.5 Or (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j) < -0.5 Then wks_output.Cells(Ccnt, 39) = (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j)
                                  Case 62, 64
                                      Retails(j).meat = Retails(j).meat + curRetail(j)
                                      If curRetail(j) > 0 Then counter(j).meat = counter(j).meat + 1
                                      If curRetail(j) > 0 Then AldiRetails(j).meat = AldiRetails(j).meat + curAldiRetail(j)
                                      Mecnt = Mecnt + 1
                                      wks_output.Cells(Mecnt, 41) = j
                                      wks_output.Cells(Mecnt, 42) = ThisProd
                                      wks_output.Cells(Mecnt, 43) = AldiProdDesc
                                      wks_output.Cells(Mecnt, 44) = sCurMatchcode
                                      wks_output.Cells(Mecnt, 45) = sCurMatchDesc
                                      wks_output.Cells(Mecnt, 46) = curRetail(j)
                                      wks_output.Cells(Mecnt, 47) = curAldiRetail(j)
                                      wks_output.Cells(Mecnt, 48) = totProdShare(j)
                                      If (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j) > 0.5 Or (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j) < -0.5 Then wks_output.Cells(Mecnt, 49) = (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j)
                                  Case 38
                                      Retails(j).frozen = Retails(j).frozen + curRetail(j)
                                      If curRetail(j) > 0 Then counter(j).frozen = counter(j).frozen + 1
                                      If curRetail(j) > 0 Then AldiRetails(j).frozen = AldiRetails(j).frozen + curAldiRetail(j)
                                      Fcnt = Fcnt + 1
                                      wks_output.Cells(Fcnt, 51) = j
                                      wks_output.Cells(Fcnt, 52) = ThisProd
                                      wks_output.Cells(Fcnt, 53) = AldiProdDesc
                                      wks_output.Cells(Fcnt, 54) = sCurMatchcode
                                      wks_output.Cells(Fcnt, 55) = sCurMatchDesc
                                      wks_output.Cells(Fcnt, 56) = curRetail(j)
                                      wks_output.Cells(Fcnt, 57) = curAldiRetail(j)
                                      wks_output.Cells(Fcnt, 58) = totProdShare(j)
                                      If (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j) > 0.5 Or (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j) < -0.5 Then wks_output.Cells(Fcnt, 59) = (curRetail(j) - curAldiRetail(j)) / curAldiRetail(j)
                                  Case 58
                                  
                                  End Select
                            End If
                        Next
                    
                    End If
                
                    If a = UBound(CBA_COM_Match) + 1 Then GoTo completereport
                    ThisProd = CBA_COM_Match(v).AldiPCode
                    
                    If DisplayData = "months" Then
                        ReDim curRetail(Month(DFrom) To Month(Dto))
                        ReDim curAldiRetail(Month(DFrom) To Month(Dto))
                        ReDim totProdShare(Month(DFrom) To Month(Dto))
                        ElseIf DisplayData = "weeks" Then
                        ReDim curRetail(1 To numWeeks)
                        ReDim curAldiRetail(1 To numWeeks)
                        ReDim totProdShare(1 To numWeeks)
                    End If
                    
                    CGno = CBA_COM_Match(v).AldiPCG
                End If
                bUse = False
                sCheckMatchUsed = LCase(CBA_COM_Match(a).MatchType)
                For Each compet In CompsToInc
                    If InStr(1, sCheckMatchUsed, compet) > 0 Then
                        If CBA_COM_Match(a).CompProdName = "" Then
                            bUse = False
                            Exit For
                        Else
                            bUse = True
                            Exit For
                        End If
                    End If
                Next
                '''*********Excluding any CBA_COM_Matches that have a private label CBA_COM_Match from branded reporting
                If bUse = True And (dset = "ColesBranded" Or dset = "WWBranded" Or dset = "DMBranded" Or dset = "Branded") Then
                    For j = LBound(CBA_COM_Match) To UBound(CBA_COM_Match)
                        If CBA_COM_Match(j).AldiPCode = ThisProd Then
                            For Each Compt In CompstoExcludefromBranded
                                If InStr(1, LCase(CBA_COM_Match(j).MatchType), Compt) > 0 Then
                                    bUse = False
                                    Exit For
                                End If
                            Next
                            If bUse = False Then Exit For
                        ElseIf CBA_COM_Match(j).AldiPCode > ThisProd Then Exit For
                        End If
                    Next
                End If
                If bUse = True Then
                    'a = a
                    If myprod <> ThisProd Then
                        colMatchused(1).Add ThisProd
                    End If
                    retARR = CBA_COM_Match(a).RetailHistoryArray("nonpromoprorata")
                    Aretarr = CBA_COM_Match(a).RetailHistoryArray("aldiretail")
                    For b = 1 To UBound(retARR, 2)
                        For Each ScrapeDate In WeekstoCollect
                            If retARR(1, b) = ScrapeDate Then
                                If LCase(retARR(2, b)) = LCase(StateLookup) Then
                                    If myprod <> CBA_COM_Match(a).AldiPCode Then
                                        myprod = CBA_COM_Match(a).AldiPCode
                                        ForEntry = False
                                        For c = 1 To UBound(AldiProdsbyPeriod, 2)
                                            If IsNull(AldiProdsbyPeriod(1, c, Month(retARR(1, b)))) Then Exit For
                                            If myprod = AldiProdsbyPeriod(1, c, Month(retARR(1, b))) Then
                                                ForEntry = True
                                                Exit For
                                            End If
                                        Next
                                        For c = 0 To UBound(CBISarr, 2)
                                            'If Mid(CBISarr(1, c), 1, 1) > Mid(myprod, 1, 1) Then Exit For
                                            If DisplayData = "months" Then
                                                If CBISarr(1, c) = myprod Then
                                                    totProdShare(CBISarr(0, c)) = CBISarr(2, c)
                                                    'totProdShare(CBISarr(0, c)) = 1
                                                    Exit For
                                                End If
                                            ElseIf DisplayData = "weeks" Then
                                                If CBISarr(1, c) = myprod And CBISarr(0, c) >= arrayWeekStart And CBISarr(0, c) <= arrayWeekend Then
                                                    totProdShare(CBISarr(0, c) - (arrayWeekStart - 1)) = CBISarr(2, c)
                                                    'totProdShare(CBISarr(0, c)) = 1
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    Else
                                        If DisplayData = "months" Then
                                            If Month(retARR(1, b)) <= Month(Dto) Then
                                                If totProdShare(Month(retARR(1, b))) = 0 Then
                                                    For c = 0 To UBound(CBISarr, 2)
                                                        'If CBISarr(1, c) > myprod Then Exit For
                                                        If CBISarr(1, c) = myprod Then
                                                            totProdShare(CBISarr(0, c)) = CBISarr(2, c)
                                                            'totProdShare(CBISarr(0, c)) = 1
                                                            Exit For
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        ElseIf DisplayData = "weeks" Then
                                            If WeekDic.Item(retARR(1, b)) < numWeeks Then
                                                If totProdShare(WeekDic.Item(retARR(1, b))) = 0 Then
                                                    For c = 0 To UBound(CBISarr, 2)
                                                        'If CBISarr(1, c) > myprod Then Exit For
                                                        If CBISarr(1, c) = myprod Then
                                                            totProdShare(WeekDic.Item(retARR(1, b))) = CBISarr(2, c)
                                                            'totProdShare(CBISarr(0, c)) = 1
                                                            Exit For
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        End If
                                    End If
                                    If ForEntry = True Then
                                        If Aretarr(3, b) > 0 Then
                                            If curRetail(Month(retARR(1, b))) = 0 Or curRetail(Month(retARR(1, b))) > retARR(3, b) * totProdShare(Month(retARR(1, b))) Then
                                                curRetail(Month(retARR(1, b))) = retARR(3, b) * totProdShare(Month(retARR(1, b)))
                                                sCurMatchcode = CBA_COM_Match(a).CompCode
                                                sCurMatchDesc = CBA_COM_Match(a).CompProdName
                                            End If
                                            If curAldiRetail(Month(Aretarr(1, b))) = 0 Or curAldiRetail(Month(Aretarr(1, b))) > Aretarr(3, b) * totProdShare(Month(Aretarr(1, b))) Then
                                                curAldiRetail(Month(Aretarr(1, b))) = Aretarr(3, b) * totProdShare(Month(Aretarr(1, b)))
                                            End If
                                        Else
                                            a = a
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Next
                Else
                    bfound = False
                    For Each prod In colexcp(1)
                        If prod = ThisProd Then
                            bfound = True
                            Exit For
                        End If
                    Next
                    If bfound = False Then colexcp(1).Add ThisProd
                End If
            Next
completereport:

            DoEvents
            With wks_WBA
            
                If lNum = 0 Then
                    .Cells.ClearContents
                    For Each cht In .ChartObjects
                    cht.Delete
                    Next
                End If
                
                If lNum > 0 Then lNum = (Around + ((Around - 1) * (NoOfMonths * 2))) - 1
                lNum = lNum + 1
                .Cells(lNum, 10).Value = dset
                .Cells(lNum, 11).Value = "Alcohol"
                .Cells(lNum, 12).Value = "Ambient Food"
                .Cells(lNum, 13).Value = "Ambient Non Food"
                .Cells(lNum, 14).Value = "Chilled"
                .Cells(lNum, 15).Value = "Meat"
                .Cells(lNum, 16).Value = "Frozen"
                .Cells(lNum, 17).Value = "Produce"
                .Cells(lNum, 18).Value = "Total"
        
                For a = LBound(Retails) To UBound(Retails)
                    lNum = lNum + 1
                    .Cells(lNum, 10).Value = MonthName(a)
                    If counter(a).alcohol > 0 And AldiRetails(a).alcohol <> 0 Then .Cells(lNum, 11).Value = (Retails(a).alcohol - AldiRetails(a).alcohol) / AldiRetails(a).alcohol Else .Cells(lNum, 11).Value = "%0"
                    If counter(a).ambientfood > 0 And AldiRetails(a).ambientfood <> 0 Then .Cells(lNum, 12).Value = (Retails(a).ambientfood - AldiRetails(a).ambientfood) / AldiRetails(a).ambientfood Else .Cells(lNum, 12).Value = "%0"
                    If counter(a).ambientnonfood > 0 And AldiRetails(a).ambientnonfood <> 0 Then .Cells(lNum, 13).Value = (Retails(a).ambientnonfood - AldiRetails(a).ambientnonfood) / AldiRetails(a).ambientnonfood Else .Cells(lNum, 13).Value = "%0"
                    If counter(a).chilled > 0 And AldiRetails(a).chilled <> 0 Then .Cells(lNum, 14).Value = (Retails(a).chilled - AldiRetails(a).chilled) / AldiRetails(a).chilled Else .Cells(lNum, 14).Value = "%0"
                    If counter(a).meat > 0 And AldiRetails(a).meat <> 0 Then .Cells(lNum, 15).Value = (Retails(a).meat - AldiRetails(a).meat) / AldiRetails(a).meat Else .Cells(lNum, 15).Value = "%0"
                    If counter(a).frozen > 0 And AldiRetails(a).frozen <> 0 Then .Cells(lNum, 16).Value = (Retails(a).frozen - AldiRetails(a).frozen) / AldiRetails(a).frozen Else .Cells(lNum, 16).Value = "%0"
                    If counter(a).produce > 0 And AldiRetails(a).produce <> 0 Then .Cells(lNum, 17).Value = (Retails(a).produce - AldiRetails(a).produce) / AldiRetails(a).produce Else .Cells(lNum, 17).Value = "%0"
            
                    totComp = Retails(a).alcohol + Retails(a).ambientfood + Retails(a).ambientnonfood + Retails(a).chilled + Retails(a).frozen + Retails(a).meat + Retails(a).produce
                    totAldi = AldiRetails(a).alcohol + AldiRetails(a).ambientfood + AldiRetails(a).ambientnonfood + AldiRetails(a).chilled + AldiRetails(a).frozen + AldiRetails(a).meat + AldiRetails(a).produce
                    If totAldi <> 0 Then .Cells(lNum, 18).Value = (totComp - totAldi) / totAldi Else .Cells(lNum, 18).Value = 0
                    Range(.Cells(lNum, 11), .Cells(lNum, 18)).NumberFormat = "#.00%"
                Next
                
                For b = LBound(counter) To UBound(counter)
                    lNum = lNum + 1
                    .Cells(lNum, 10).Value = MonthName(b) & " Units"
                    .Cells(lNum, 11).Value = Format(counter(b).alcohol, "0")
                    .Cells(lNum, 12).Value = Format(counter(b).ambientfood, "0")
                    .Cells(lNum, 13).Value = Format(counter(b).ambientnonfood, "0")
                    .Cells(lNum, 14).Value = Format(counter(b).chilled, "0")
                    .Cells(lNum, 15).Value = Format(counter(b).meat, "0")
                    .Cells(lNum, 16).Value = Format(counter(b).frozen, "0")
                    .Cells(lNum, 17).Value = Format(counter(b).produce, "0")
                    .Cells(lNum, 18).Value = Format(Application.WorksheetFunction.Sum(Range(.Cells(lNum, 11), .Cells(lNum, 17))), "0")
                    Range(.Cells(lNum, 11), .Cells(lNum, 18)).NumberFormat = "#0"
                Next
                lNum = 1
            
                If dset = "ColesCOS" Then strChartname = "ALDI vs Coles Cheapest on Show" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "WWCOS" Then strChartname = "ALDI vs Woolworths Cheapest on Show" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "DMCOS" Then strChartname = "ALDI vs Dan Murphys Cheapest on Show" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "ColesVal" Then strChartname = "ALDI vs Coles Value" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "ColesSB" Then strChartname = "ALDI vs Coles SmartBuy" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "ColesPL" Then strChartname = "ALDI vs Coles" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "WWHB" Then strChartname = "ALDI vs Homebrand" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "WWSelect" Then strChartname = "ALDI vs Select" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "DM1" Then strChartname = "ALDI vs Dan Murphys Primary" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "DM2" Then strChartname = "ALDI vs Dan Murphys Secondary" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "Branded" Then strChartname = "ALDI vs Branded" & Chr(10) & "% ALDI Cheaper by (sub basket)" & Chr(10) & "(Where there is no private label competitor)"
                If dset = "Phantoms" Then strChartname = "ALDI vs Phantom Brands" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "WWPhantom" Then strChartname = "ALDI vs Woolworths Phantom Brands" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "DMPhantom" Then strChartname = "ALDI vs Dan Murphy's Phantom Brands" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "ColesPhantom" Then strChartname = "ALDI vs Coles Phantom Brands" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "DMBranded" Then strChartname = "ALDI vs Dan Murphys Branded" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "WWBranded" Then strChartname = "ALDI vs Woolworths Branded" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "ColesBranded" Then strChartname = "ALDI vs Coles Branded" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "ColesPhantoms" Then strChartname = "ALDI vs Coles Phantom Brands" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "WWPhantoms" Then strChartname = "ALDI vs Woolworths Phantom Brands" & Chr(10) & "% ALDI Cheaper by (sub basket)"
                If dset = "All" Then strChartname = "ALDI vs All Competitors" & Chr(10) & "% ALDI Cheaper by (sub basket)"
            
                If Around = 1 Then
                    .Cells(1, 6).Value = "Weighted Basket Analysis"
                    Point = 1: PointRow = 3
                    SourceCellNo = 2
                    SourceCellNoEnd = SourceCellNo + (NoOfMonths * 2) - Around
                    CatCellNo = (1 * Around) + ((Around - 1) * (NoOfMonths * 2))
                
                
                    Set SRng = Range(.Cells(SourceCellNo, 11), .Cells(SourceCellNoEnd, 18))  'SourceDATA
                    Set LRng = Range(.Cells(SourceCellNo, 10), .Cells(SourceCellNoEnd, 10))   'Legend Range
                    Set CRng = Range(.Cells(CatCellNo, 11), .Cells(CatCellNo, 18))  'Category Range
            
                    ThisChart = CBAR_WBA_ChartCreate.WBAChartCreate(.Cells(3, 1).Left, .Cells(3, 1).Top, _
                                SRng, wks_WBA, strChartname, LRng, CRng, , xlColumnClustered, xlColumnClustered, xlRows, , , , 1 + NoOfMonths, True, 102)
                Else
                
                    If Round(Around / 2, 0) = Around / 2 Then
                        Point = 8
                    Else
                        Point = 1
                        PointRow = PointRow + 19
                    End If
                    SourceCellNo = Around + 1 + ((Around - 1) * (NoOfMonths * 2))
                    SourceCellNoEnd = SourceCellNo + (NoOfMonths * 2) - 1
                    CatCellNo = (Around + ((Around - 1) * (NoOfMonths * 2)))
                
                    Set SRng = Range(.Cells(SourceCellNo, 11), .Cells(SourceCellNoEnd, 18))  'SourceDATA
                    Set LRng = Range(.Cells(SourceCellNo, 10), .Cells(SourceCellNoEnd, 10))   'Legend Range
                    Set CRng = Range(.Cells(CatCellNo, 11), .Cells(CatCellNo, 18))  'Category Range
                    
                    ThisChart = CBAR_WBA_ChartCreate.WBAChartCreate(.Cells(PointRow, Point).Left, .Cells(PointRow, Point).Top, _
                                SRng, wks_WBA, strChartname, LRng, CRng, , xlColumnClustered, xlColumnClustered, xlRows, , , , 1 + NoOfMonths, True, 102)
                
                End If
            
            
                wks_WBA.Cells(1, 1).Select
            End With
        
        Next dset
            
        DoEvents
        wks_WBA.Cells(1, 17).EntireColumn.Delete
        wks_WBA.Cells(1, 11).EntireColumn.Delete
        wks_WBA.Activate
        DoEvents
            
            
        With wks_WBA.PageSetup
            .PrintArea = Range(wks_WBA.Cells(1, 1), wks_WBA.Cells(PointRow + 20, 9)).Address
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .LeftFooter = "&9CORP BUYING, Admin, per: " & Format(Date, "DD/MM/YYYY") & Chr(10) & "Weighted Basket Analysis (by Share)"
            .Orientation = xlPortrait
            .PrintGridlines = False
            .PrintTitleRows = ""
            .RightFooter = "&P of &N"
            .PaperSize = xlPaperA4
        End With
            
            
        CBA_COM_CBISCN.Close: Set CBA_COM_CBISCN = Nothing
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
        wks_output.Visible = xlSheetHidden
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-WBA", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub







