Attribute VB_Name = "CBA_BTF_SetupForecastArray"
' option explicit       ' CBA_BTF_SetupForecastArray @CBA_BTF Changed 181205
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Function CBA_BTF_SetupForecastArray(ByVal YearFrom As Long, ByVal YearTo As Long, ByVal MonthFrom As Long, ByVal _
                MonthTo As Long, Optional ByRef ArrayWithProductClassesAsNumbersIncluded As Variant, Optional runForAll As Boolean, _
                Optional ByVal includeCBISData As Boolean = True _
                ) As Boolean
    Dim CGs() As Variant
    Dim strSQL As String, strSQLCG As String, strPClass As String
    Dim Fore() As Variant
    Dim PCStart As Long, PCEnd As Long, a As Long, dates, dateE, lYY As Long, lMM As Long, b As Long, c As Long, lCurCG As Long
    Dim SCGcnt As Long
    Dim isitEmpty As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = "GetConn"
    
    If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Preparing Forecast Data.."
    CGs = CBA_BTF_Runtime.getCGs
    Set CBA_DBCN = New ADODB.Connection
    
    With CBA_DBCN
        .ConnectionTimeout = 100
        .CommandTimeout = 100
        .Open "Provider= " & CBA_MSAccess & "; DATA SOURCE= " & g_GetDB("ForeCast") '; ;INTEGRATED SECURITY=sspi;"
    End With
    
    If includeCBISData = True Then
        Set CBA_COM_SKU_CBISCN = New ADODB.Connection
        With CBA_COM_SKU_CBISCN
            .ConnectionTimeout = 100
            .CommandTimeout = 100
            .Open "Provider= SQLNCLI10; DATA SOURCE= 599DBL01; ;INTEGRATED SECURITY=sspi;"
        End With
    End If
    
    
    ''Replaced with an automated script that runs every hour
    'Dim acObj As New Access.Application
    'acObj.Application.Visible = False
    'acObj.OpenCurrentDatabase cba_bsa & "LIVE DATABASES\CBForecast.accdb", , "ADatabasePassword"
    'acObj.Application.Run "t_TestForecastsPC", True
    'acObj.Quit
    ''End of replacement
    
    dates = DateSerial(YearFrom, MonthFrom, 1)
    dateE = DateSerial(YearTo, MonthTo + 1, 0)
    strSQLCG = ""
    
    On Error Resume Next
    If UBound(CGs, 2) < 0 Then
        isitEmpty = True
        Err.Clear
    End If
    On Error GoTo Err_Routine
    
    If isitEmpty = False Then
        lCurCG = -1
        For a = LBound(CGs, 2) To UBound(CGs, 2)
            If a = LBound(CGs, 2) Then
                strSQLCG = "And (" & Chr(10)
            Else
                If strSQLCG <> "And (" & Chr(10) And Right(strSQLCG, 1) <> "," Then strSQLCG = strSQLCG & " Or " & Chr(10)
            End If
            If IsEmpty(CGs(0, a)) = False Then
                
                If CGs(1, a) = 0 Then
                    strSQLCG = strSQLCG & "(p.CGno = " & CGs(0, a) & ")"
                    If a + 1 < UBound(CGs, 2) Then
                        If CGs(0, a) = CGs(0, a + 1) Then
                            c = 0
                            For b = a + 1 To UBound(CGs, 2)
                                If CGs(0, a) = CGs(0, b) Then
                                    c = c + 1
                                Else
                                    a = a + c
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                Else
                    If a = 0 Then
                        strSQLCG = strSQLCG & "(p.CGno = " & CGs(0, a) & ") or "
    '                ElseIf CGs(0, a) <> CGs(0, a - 1) Then
    '                    strSQLCG = strSQLCG & "(p.CGno = " & CGs(0, a) & ") or "
                    End If
                    
'                    strSQLCG = strSQLCG & "(p.CGno = " & CGs(0, a) & " and p.SCGno = " & CGs(1, a) & ")" & Chr(10)
                    If lCurCG <> CGs(0, a) Then
                        lCurCG = CGs(0, a)
                        SCGcnt = 1
                        If UBound(CGs, 2) = 0 Then                                                            '#RW Added Line
                          strSQLCG = strSQLCG & "(p.CGno = " & CGs(0, a) & " and p.SCGno in (0"               '#RW Added Line
                        Else
                          strSQLCG = strSQLCG & "(p.CGno = " & CGs(0, a) & " and p.SCGno in (0,"
                        End If
                    Else
                        SCGcnt = SCGcnt + 1
                    End If
                    'strSQLCG = strSQLCG & "(p.CGno = " & CGs(0, a) & " and p.SCGno = " & CGs(1, a) & ")" & Chr(10)
''63942                         If Right(strSQLCG, 1) = "," Then strSQLCG = g_Left(strSQLCG, 1)                           '#RW Added Line

                    If a = UBound(CGs, 2) Then
                        strSQLCG = strSQLCG & CGs(1, a) & "))" & Chr(10)
                    ElseIf CGs(0, a + 1) = CGs(0, a) Then
                        strSQLCG = strSQLCG & CGs(1, a) & ","
                    Else
                        strSQLCG = strSQLCG & CGs(1, a) & "))" & Chr(10)
                    End If
                    'End If
                End If
            End If
        Next
        strSQLCG = strSQLCG & ")"
    End If
    CBA_ErrTag = "Array"
    If IsMissing(ArrayWithProductClassesAsNumbersIncluded) = False Then
        PCStart = 4: PCEnd = 1
        For a = LBound(ArrayWithProductClassesAsNumbersIncluded) To UBound(ArrayWithProductClassesAsNumbersIncluded)
            If ArrayWithProductClassesAsNumbersIncluded(a) < PCStart Then PCStart = ArrayWithProductClassesAsNumbersIncluded(a)
            If ArrayWithProductClassesAsNumbersIncluded(a) > PCEnd Then PCEnd = ArrayWithProductClassesAsNumbersIncluded(a)
            If strPClass = "" Then
                strPClass = "and productclass in (" & ArrayWithProductClassesAsNumbersIncluded(a)
            Else
                strPClass = strPClass & "," & ArrayWithProductClassesAsNumbersIncluded(a)
            End If
        Next
        strPClass = strPClass & ")"
    Else
        PCStart = 1: PCEnd = 4
    End If

    If includeCBISData = True Then

        If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Querying CBIS.."
        CBA_ErrTag = "SQL"
        'CBIS Query
        Set CBA_COM_SKU_CBISRS = New ADODB.Recordset
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
        strSQL = strSQL & "declare @SDATE date = '" & Format(dates, "YYYY-MM-DD") & "'" & Chr(10)
        strSQL = strSQL & "declare @EDATE date = '" & Format(dateE, "YYYY-MM-DD") & "'" & Chr(10)
    
        strSQL = strSQL & "Select month(pdate) as Monthno, year(pdate) as Yearno" & Chr(10)
        strSQL = strSQL & "into #DATES" & Chr(10)
        strSQL = strSQL & "from (SELECT  TOP (DATEDIFF(DAY, @SDATE, @EDATE) + 1)" & Chr(10)
        strSQL = strSQL & "Pdate = convert(date,DATEADD(DAY, ROW_NUMBER() OVER(ORDER BY a.object_id) - 1, @SDATE))" & Chr(10)
        strSQL = strSQL & "FROM    sys.all_objects a CROSS JOIN sys.all_objects b) a" & Chr(10)
        strSQL = strSQL & "group by month(pdate), year(pdate) order by  month(pdate), year(pdate)" & Chr(10)
    
        strSQL = strSQL & "select distinct cg.cgno, isnull(scg.scgno,0) as scgno into #CG from cbis599p.dbo.COMMODITYGROUP cg" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.SUBCOMMODITYGROUP scg on scg.cgno = cg.cgno and scg.scgno = scg.scgno" & Chr(10)
        If strSQLCG <> "" Then strSQL = strSQL & Replace(Replace(Replace(strSQLCG, "And", "Where"), "p.CGno", "cg.CGno"), "p.SCGno", "scg.SCGno") & Chr(10)
    
        strSQL = strSQL & "select d.Monthno, d.Yearno, c.CGno, c.SCGno, pc.ProductClass into #BASE from #DATES d cross join #CG c" & Chr(10)
        strSQL = strSQL & "cross join (select distinct productclass from cbis599p.dbo.PRODUCTCLASS" & Chr(10)
        If strPClass <> "" Then strSQL = strSQL & Replace(strPClass, "and", "Where") & Chr(10)
        strSQL = strSQL & ") as pc" & Chr(10)
    
        strSQL = strSQL & "Select year(posdate) as yearno, month(posdate) as monthno, p.CGno, isnull(p.scgno,0) as scgno, p.productclass, sum(isnull(retail,0)) as posretail" & Chr(10)
        strSQL = strSQL & "into #POS from cbis599p.dbo.pos" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.product p on p.productcode = pos.productcode" & Chr(10)
        strSQL = strSQL & "inner join #CG cg on cg.CGNo = p.CGNo and isnull(cg.SCGNo,0) = isnull(p.scgno,0)" & Chr(10)
        strSQL = strSQL & "where pos.posdate >= @SDATE and pos.posdate <= @EDATE" & Chr(10)
        If strPClass <> "" Then strSQL = strSQL & strPClass & Chr(10)
        strSQL = strSQL & "group by year(posdate) , month(posdate),p.CGno, p.SCgno, p.productclass" & Chr(10)
        strSQL = strSQL & "order by CGno, SCgno, p.productclass,year(posdate) , month(posdate)" & Chr(10)
    
        strSQL = strSQL & "select year(rcv.dayenddate) as yearno, month(rcv.dayenddate) as monthno, p.cgno,isnull(p.SCGno,0) as SCGNo, p.productclass , sum(isnull(retailnet,0))  as retailnet , sum(isnull(cost,0)) as cost" & Chr(10)
        strSQL = strSQL & ", sum(isnull(retail,0)) as retail" & Chr(10)
        strSQL = strSQL & "into #RCV from cbis599p.dbo.RECEIVING rcv" & Chr(10)
        strSQL = strSQL & "left join cbis599p.dbo.product p on p.productcode = rcv.productcode" & Chr(10)
        strSQL = strSQL & "inner join #CG cg on cg.CGNo = p.CGNo and isnull(cg.SCGNo,0) = isnull(p.scgno,0)" & Chr(10)
        strSQL = strSQL & "where DayEndDate >= @SDATE and DayEndDate <= @EDATE" & Chr(10)
        If strPClass <> "" Then strSQL = strSQL & strPClass & Chr(10)
        strSQL = strSQL & "group by year(rcv.dayenddate) , month(rcv.dayenddate), p.cgno, p.scgno, p.productclass" & Chr(10)
        strSQL = strSQL & "order by p.cgno, p.scgno" & Chr(10)
    
        strSQL = strSQL & "select b.yearno, b.monthno, b.cgno,b.SCGno, b.productclass , isnull(rcv.retailnet,0) as retailnet, isnull(rcv.cost,0) as cost" & Chr(10)
        strSQL = strSQL & ", isnull(rcv.retail ,0) as retail , isnull(pos.posretail,0) as posretail" & Chr(10)
        strSQL = strSQL & "into #OP" & Chr(10)
        strSQL = strSQL & "from #BASE b" & Chr(10)
        strSQL = strSQL & "left join #POS pos on pos.CGNo = b.cgno and pos.SCGNo = b.SCGNo and pos.ProductClass = b.ProductClass and pos.yearno = b.Yearno and pos.monthno = b.Monthno" & Chr(10)
        strSQL = strSQL & "left join #RCV rcv on rcv.CGNo = b.cgno and rcv.SCGNo = b.SCGNo and rcv.ProductClass = b.ProductClass and rcv.yearno = b.Yearno and rcv.monthno = b.Monthno" & Chr(10)
        strSQL = strSQL & "order by b.yearno, b.monthno, b.cgno,b.SCGno, b.productclass" & Chr(10)
    
        strSQL = strSQL & "drop table #DATES,#CG, #BASE , #RCV, #POS" & Chr(10)
        'Debug.Print strSQL
        CBA_COM_SKU_CBISRS.Open strSQL, CBA_COM_SKU_CBISCN
    End If

    If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Querying Forecast Database.."
    'Forecast Query
    Set CBA_DBRS = New ADODB.Recordset
    strSQL = "Select * from DATAQ "
    If strPClass <> "" Then
        strSQL = strSQL & Replace(strPClass, "and", "Where")
        If strSQLCG <> "" Then strSQL = strSQL & strSQLCG
    Else
        If strSQLCG <> "" Then strSQL = strSQL & Replace(Replace(Replace(Replace(strSQLCG, "And", "Where"), "p.", ""), "CGno", "CG"), "SCGno", "SCG")
    End If
'    strSQL = strSQL & "and YearNo = " & Yearto
    strSQL = strSQL & "Order by CG,SCG,PC"
    'Debug.Print strSQL
    CBA_DBRS.Open strSQL, CBA_DBCN
    If CBA_DBRS.EOF Then
        ReDim Fore(0, 0)
        Fore(0, 0) = 0
    Else
        Fore = CBA_DBRS.GetRows()
    End If
    Set CBA_DBRS = Nothing

    ReDim FCbM(YearFrom To YearTo, MonthFrom To MonthTo, 1 To 4)
    CBA_ErrTag = "Array"
    For lYY = LBound(FCbM, 1) To UBound(FCbM, 1)
        For lMM = LBound(FCbM, 2) To UBound(FCbM, 2)
            DoEvents
            For a = PCStart To PCEnd
                If a = 1 Then strPClass = " Core Range"
                If a = 2 Then strPClass = " FoodSpecial"
                If a = 3 Then strPClass = " NonFoodSpecial"
                If a = 4 Then strPClass = " Seasonal"
                If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Building Object: " & lMM & "-" & lYY & strPClass
                Set FCbM(lYY, lMM, a) = New CBA_BTF_MonthData
'                If lYY = 2019 And lMM = 1 And a = 1 Then
'                a = a
'                End If
                FCbM(lYY, lMM, a).formulate lYY, lMM, a, Fore, includeCBISData
                
            Next
        Next
    Next
    If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Objects Created"
    
    
    If IsMissing(FCbM) = False Then CBA_BTF_SetupForecastArray = True
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_BTF_SetupForecastArray", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & strSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

