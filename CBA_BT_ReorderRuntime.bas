Attribute VB_Name = "CBA_BT_ReorderRuntime"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Private Type POSDataType
    DAT As Date
    QTY As Single
    Retail As Single
End Type

Sub ReorderRuntime(ByVal OSD As Date, ByVal PCode As Long, ByVal SellT As Single, ByVal MaxNum As Single, Optional MultiDrop As Boolean)
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim strSQL As String, strSQLDrop As String, strSQLPOS As String, wks_Reorder As Worksheet
    Dim lRowNo As Long, a As Long, b As Long, casesamt As Long
    Dim sumarr() As Long, tic As Long, rcd As Long, nodiff As Long
    Dim RCell As Range
    Dim POSData() As POSDataType
    Dim vDiv As Long
    Dim divfnd As Long
    Dim wks_MultiDropMAA As Worksheet
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    ReDim sumarr(1 To 2, 501 To 509)
    
    CBA_BasicFunctions.CBA_Running "Preparing Repeat Reorder"
    Application.ScreenUpdating = False
    
    CBA_SQL_Queries.CBA_GenPullSQL "CBIS_CG", , , PCode
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
        strSQL = strSQL & "declare @PROD int = " & PCode & Chr(10)
        strSQL = strSQL & "declare @SDATE date = '" & Format(OSD, "YYYY-MM-DD") & "'" & Chr(10)
    If CBA_CBISarr(0, 0) = 0 Then
        MsgBox "Not a recognised productcode"
        Exit Sub
    ElseIf CBA_CBISarr(0, 0) < 5 Then
            strSQL = strSQL & "declare @EDATE date = dateadd(D, 13, @SDATE)" & Chr(10)
            strSQL = strSQL & "declare @RCVPdays int = 6" & Chr(10)
            strSQL = strSQL & "declare @RCVAdays int = 0" & Chr(10)
    ElseIf CBA_CBISarr(0, 0) = 62 Or CBA_CBISarr(0, 0) = 64 Then
            strSQL = strSQL & "declare @EDATE date = dateadd(D, 6, @SDATE)" & Chr(10)
            strSQL = strSQL & "declare @RCVPdays int = 3" & Chr(10)
            strSQL = strSQL & "declare @RCVAdays int = 3" & Chr(10)
    ElseIf CBA_CBISarr(0, 0) = 51 And (CBA_CBISarr(1, 0) = 22 Or CBA_CBISarr(1, 0) = 17 Or CBA_CBISarr(1, 0) = 23 Or CBA_CBISarr(1, 0) = 25) Then
            strSQL = strSQL & "declare @EDATE date = dateadd(D, 6, @SDATE)" & Chr(10)
            strSQL = strSQL & "declare @RCVPdays int = 3" & Chr(10)
            strSQL = strSQL & "declare @RCVAdays int = 3" & Chr(10)
    Else
        MsgBox "Not a meat or alcohol product"
        Exit Sub
    End If
        strSQL = strSQL & "declare @Pack nvarchar(10) = (select packsize from purchase.dbo.product where productcode = @PROD)" & Chr(10)
        strSQL = strSQL & "declare @Days decimal(4,2) = datediff(D,@SDATE,@EDATE) + 1" & Chr(10)
        strSQL = strSQL & "SELECT  TOP (DATEDIFF(DAY, @SDATE, @EDATE) + 1) Pdate = convert(date,DATEADD(DAY, ROW_NUMBER() OVER(ORDER BY a.object_id) - 1, @SDATE))" & Chr(10)
        strSQL = strSQL & "into #DATE FROM sys.all_objects a CROSS JOIN sys.all_objects b" & Chr(10)
        strSQL = strSQL & "select productcode into #PROD from purchase.dbo.product p where coalesce(mainproduct,con_productcode, productcode) = @PROD" & Chr(10)
        strSQL = strSQL & "select storeno into #STORES from  purchase.dbo.storesales pos where pos.productcode in (select productcode from #PROD)" & Chr(10)
        strSQL = strSQL & "and pos.salesDate >= @SDATE and pos.SalesDate <= @EDATE group by storeno" & Chr(10)
        strSQL = strSQL & "select d.pdate, s.storeno into #BASE from #DATE d cross join #STORES s" & Chr(10)
        strSQL = strSQL & "select salesdate as posdate, storeno, sum(quantity) as QTY, sum(retail) as Retail" & Chr(10)
        strSQL = strSQL & "into #POS from Purchase.dbo.storesales pos" & Chr(10)
        strSQL = strSQL & "inner join #PROD p  on p.productcode = pos.ProductCode" & Chr(10)
        strSQL = strSQL & "where pos.salesDate >= @SDATE and pos.SalesDate <= @EDATE" & Chr(10)
        strSQL = strSQL & "group by salesdate , storeno" & Chr(10)
        strSQL = strSQL & "select  b.Pdate as posdate, b.storeno, p.QTY, p.Retail" & Chr(10)
        strSQL = strSQL & ", case when isnull(p.retail,0) = 0 then 1 else 0 end as cnt" & Chr(10)
        strSQL = strSQL & "into #POSFULL  from #BASE b" & Chr(10)
        strSQL = strSQL & "left join #POS p  on p.posDate = b.Pdate and p.StoreNo = b.storeno" & Chr(10)
        strSQL = strSQL & "select storeno, sum(quantity)  as Stock" & Chr(10)
        strSQL = strSQL & "into #STOCK from purchase.dbo.StoreReceiving sr" & Chr(10)
        strSQL = strSQL & "inner join #PROD p on p.productcode = sr.ProductCode" & Chr(10)
        strSQL = strSQL & "where ReceivingDate >= dateadd(D,-@RCVPdays,@SDATE)" & Chr(10)
        strSQL = strSQL & "and ReceivingDate <= dateadd(D,@RCVAdays,@SDATE) group by storeno" & Chr(10)
        strSQL = strSQL & "Select case when substring(@@Servername,1,1) = 0 then substring(@@Servername,2,3)  else substring(@@Servername,1,3) end as Div , convert(nvarchar(5),pos.storeno) + '-' + s.City, sum(isnull(pos.QTY,0)) as QTY , sum(isnull(pos.Retail,0)) as Retail, @Pack" & Chr(10)
        strSQL = strSQL & ", st.stock as stockallocation, sum(isnull(pos.cnt,0)) / @Days as nosalesdays, st.stock - sum(isnull(pos.QTY,0)) as residualstock" & Chr(10)
        strSQL = strSQL & ", case when st.stock - sum(isnull(pos.QTY,0)) = 0 and sum(isnull(pos.cnt,0)) / @Days > 0 then" & Chr(10)
        strSQL = strSQL & "case when ((sum(isnull(pos.QTY,0)) / (@Days * (1-(sum(isnull(pos.cnt,0)) / @Days)))) * @Days) > (sum(isnull(pos.QTY,0)) * 2) then sum(isnull(pos.QTY,0)) * 2 else (sum(isnull(pos.QTY,0)) / (@Days * (1-(sum(isnull(pos.cnt,0)) / @Days)))) * @Days end" & Chr(10)
        strSQL = strSQL & "else sum(isnull(pos.QTY,0)) end as RevisedQTY from #POSFULL pos" & Chr(10)
        strSQL = strSQL & "left join purchase.dbo.Store s on s.StoreNo = pos.StoreNo" & Chr(10)
        strSQL = strSQL & "left join #STOCK as st on st.storeno = pos.StoreNo" & Chr(10)
        strSQL = strSQL & "group by pos.storeno, s.City, st.stock" & Chr(10)
    
        strSQLDrop = "drop table #PROD,#DATE, #STORES,#BASE,#POS , #POSFULL, #STOCK" & Chr(10)
        strSQLPOS = "select posdate, sum(QTY), sum(Retail) from #POS group by posdate order by posdate  "
 
    
    Application.Workbooks.Add
    Set wks_Reorder = ActiveSheet
    With wks_Reorder
    .Name = "MAAReorder"
    lRowNo = 6
    Range(.Cells(1, 1), .Cells(5, 49)).Interior.ColorIndex = 49
    .Cells(1, 1).Select
    .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
    .Cells(4, 2).Font.Name = "ALDI SUED Office"
    .Cells(4, 2).Font.Size = 24
    .Cells(4, 2).Font.ColorIndex = 2
    .Cells(lRowNo, 1).Value = "Region"
    .Cells(lRowNo, 2).Value = "Store Name"
    .Cells(lRowNo, 3).Value = "POS Units"
    .Cells(lRowNo, 4).Value = "POS Retail"
    .Cells(lRowNo, 5).Value = "Stock Allocation"
    .Cells(lRowNo, 6).Value = "No Sale Days"
    .Cells(lRowNo, 7).Value = "Residual Stock"
    .Cells(lRowNo, 8).Value = "revised QTY"
    .Cells(lRowNo, 9).Value = "Unit Reorder"
    .Cells(lRowNo, 10).Value = "Case Reorder"
    Range(.Cells(lRowNo, 1), .Cells(lRowNo, 17)).Font.Bold = True
    Range(.Cells(lRowNo, 1), .Cells(lRowNo, 17)).Font.Underline = True
        For a = 501 To 509
            If a <> 508 Then
                Set CN = New ADODB.Connection
                Set RS = New ADODB.Recordset
                    With CN
                        .ConnectionTimeout = 100
                        .CommandTimeout = 300
                        .Open "Provider= SQLNCLI10; DATA SOURCE= " & CBA_BasicFunctions.TranslateServerName(a, Date) & "; ;INTEGRATED SECURITY=sspi;"
                    End With
                    RS.Open strSQL, CN
                    Do Until RS.EOF
                        lRowNo = lRowNo + 1
                        .Cells(lRowNo, 1).Value = CBA_BasicFunctions.CBA_DivtoReg(RS(0))
                        For b = 1 To 3
                            .Cells(lRowNo, b + 1).Value = RS(b)
                        Next
                        .Cells(lRowNo, 5).Value = RS(5)
                        .Cells(lRowNo, 6).Value = RS(6)
                        .Cells(lRowNo, 7).Value = RS(7)
                        .Cells(lRowNo, 8).Value = RS(8)
                        .Cells(lRowNo, 9).Value = RS(8) / SellT
                        .Cells(lRowNo, 9).NumberFormat = "#,0.00"
                        casesamt = Application.WorksheetFunction.RoundUp((((RS(8) / SellT) / RS(4))), 0)
                        .Cells(lRowNo, 10).Value = casesamt
                        sumarr(1, RS(0)) = sumarr(1, RS(0)) + casesamt
                        sumarr(2, RS(0)) = sumarr(2, RS(0)) + (casesamt * RS(4))
                        RS.MoveNext
                    Loop
                    
                    If MultiDrop = True Then
                        Set RS = New ADODB.Recordset
                        RS.Open strSQLPOS, CN
                        tic = 0
                        Do Until RS.EOF
                            tic = tic + 1
                            If rcd < tic Then
                                rcd = tic
                                ReDim Preserve POSData(501 To 509, 1 To rcd)
                            End If
                            POSData(a, tic).DAT = RS(0)
                            POSData(a, tic).QTY = RS(1)
                            POSData(a, tic).Retail = RS(2)
                            RS.MoveNext
                        Loop
                        RS.Close
                    End If
                    Set RS = New ADODB.Recordset
                    RS.Open strSQLDrop, CN
                CN.Close
            End If
        Next
    .Cells(1, 6).EntireColumn.NumberFormat = "0%"
    Range(.Cells(6, 1), .Cells(lRowNo, 10)).EntireColumn.AutoFit
    CBA_SQL_Queries.CBA_GenPullSQL "CBIS_ProductDesc", , , PCode
    .Cells(4, 2).Value = "Repeat Reorder" & PCode & ": " & CBA_CBISarr(0, 0) & " OSD:" & OSD
    .Cells(6, 15).Value = "Region": .Cells(6, 16).Value = "Total Cases": .Cells(6, 17).Value = "Total Units"
    .Cells(7, 15).Value = CBA_BasicFunctions.CBA_DivtoReg(501): .Cells(7, 16).Value = sumarr(1, 501): .Cells(7, 17).Value = sumarr(2, 501)
    .Cells(8, 15).Value = CBA_BasicFunctions.CBA_DivtoReg(502): .Cells(8, 16).Value = sumarr(1, 502): .Cells(8, 17).Value = sumarr(2, 502)
    .Cells(9, 15).Value = CBA_BasicFunctions.CBA_DivtoReg(503): .Cells(9, 16).Value = sumarr(1, 503): .Cells(9, 17).Value = sumarr(2, 503)
    .Cells(10, 15).Value = CBA_BasicFunctions.CBA_DivtoReg(504): .Cells(10, 16).Value = sumarr(1, 504): .Cells(10, 17).Value = sumarr(2, 504)
    .Cells(11, 15).Value = CBA_BasicFunctions.CBA_DivtoReg(505): .Cells(11, 16).Value = sumarr(1, 505): .Cells(11, 17).Value = sumarr(2, 505)
    .Cells(12, 15).Value = CBA_BasicFunctions.CBA_DivtoReg(506): .Cells(12, 16).Value = sumarr(1, 506): .Cells(12, 17).Value = sumarr(2, 506)
    .Cells(13, 15).Value = CBA_BasicFunctions.CBA_DivtoReg(507): .Cells(13, 16).Value = sumarr(1, 507): .Cells(13, 17).Value = sumarr(2, 507)
    .Cells(14, 15).Value = CBA_BasicFunctions.CBA_DivtoReg(509): .Cells(14, 16).Value = sumarr(1, 509): .Cells(14, 17).Value = sumarr(2, 509)
    
    
    Range(.Cells(21, 15), .Cells(21, 19)).Merge
    .Cells(21, 15).Value = "REPEAT SUMMARY"
    Range(.Cells(21, 15), .Cells(24, 19)).Font.Bold = True
    
    Range(.Cells(22, 15), .Cells(22, 16)).Merge
    .Cells(22, 15).Value = "Previous Special"
    Range(.Cells(22, 17), .Cells(22, 18)).Merge
    .Cells(22, 17).Value = "Proposed Special"
    .Cells(22, 19).Value = "Variance"
    
    .Cells(23, 15).Value = "Total Units": .Cells(23, 17).Value = "Total Units"
    .Cells(24, 15).Value = "Total POS Value": .Cells(24, 17).Value = "Total POS Value"
    
    .Cells(23, 16).Value = "=SUM(R7C5:R" & lRowNo & "C5)"
    .Cells(23, 16).NumberFormat = "#,0"
    .Cells(24, 16).Value = "=(SUM(R7C4:R" & lRowNo & "C4)/SUM(R7C3:R" & lRowNo & "C3))*R23C16"
    .Cells(24, 16).NumberFormat = "$#,0"
    
    .Cells(23, 18).Value = "=SUM(R7C17:R14C17)"
    .Cells(23, 18).NumberFormat = "#,0"
    
    .Cells(24, 18).Value = "=(SUM(R7C4:R" & lRowNo & "C4)/SUM(R7C3:R" & lRowNo & "C3))*R23C18"
    .Cells(24, 18).NumberFormat = "$#,0"
    
    .Cells(23, 19).Value = "=RC[-1]/RC[-3]-1"
    .Cells(23, 19).NumberFormat = "$0.0%"
    .Cells(24, 19).Value = "=RC[-1]-RC[-3]"
    .Cells(24, 19).NumberFormat = "$#,0"
    Range(.Cells(21, 15), .Cells(21, 19)).Borders.Weight = xlThin
    Range(.Cells(22, 15), .Cells(24, 16)).Borders.Weight = xlThin
    Range(.Cells(22, 17), .Cells(24, 18)).Borders.Weight = xlThin
    Range(.Cells(22, 19), .Cells(24, 19)).Borders.Weight = xlThin
    
    Range(.Cells(21, 15), .Cells(22, 19)).HorizontalAlignment = xlCenter
    
    .Cells(23, 16).Font.Bold = False
    .Cells(23, 18).Font.Bold = False
    
    If MaxNum > 0 Then
        Range(.Cells(4, 16), .Cells(4, 17)).Merge
        Range(.Cells(4, 16), .Cells(4, 17)).Font.Bold = True
        .Cells(4, 15).Value = "Desired Units"
        .Cells(4, 16).Value = MaxNum
        Range(.Cells(4, 15), .Cells(5, 19)).Font.ColorIndex = 0
        Range(.Cells(4, 15), .Cells(5, 19)).Borders.ColorIndex = 0
        Range(.Cells(5, 15), .Cells(5, 19)).Borders(xlEdgeTop).Weight = xlThin
        Range(.Cells(5, 15), .Cells(5, 19)).Borders.Weight = xlThin
        Range(.Cells(5, 18), .Cells(5, 19)).Merge
        .Cells(6, 18).Value = "Total Units"
        .Cells(6, 19).Value = "Total Cases"
        .Cells(5, 18).Value = "Desired"
        .Cells(5, 17).Value = "Reorder"
        .Cells(6, 18).Font.Underline = True
        .Cells(6, 18).Font.Bold = True
        .Cells(6, 11).Value = "Desired Units"
        .Cells(6, 12).Value = "Desired Cases"
        Range(.Cells(6, 11), .Cells(6, 12)).Font.Underline = True
        Range(.Cells(6, 11), .Cells(6, 12)).Font.Bold = True
        For Each RCell In .Columns(11).Cells
            If RCell.Row > 6 And RCell.Offset(0, -1).Value <> "" Then
                RCell.Value = "=(RC[-2]/SUM(C9))*R4C16"
                RCell.Offset(0, 1).Value = "=ROUNDUP((RC[-1]/" & casesamt & "),0)"
            ElseIf RCell.Row > 6 And RCell.Offset(0, -1).Value = "" Then
                Exit For
            End If
        Next
        For a = 7 To 14
            .Cells(a, 18).Value = "=ROUND(IFERROR(SUMIF(C1,RC[-3],C11),0),0)"
            .Cells(a, 19).Value = "=IFERROR(SUMIF(C1,RC[-4],C12),0)"
        Next
        Range(.Cells(7, 18), .Cells(14, 19)).NumberFormat = "#,0"
        Range(.Cells(25, 17), .Cells(25, 18)).Merge
        Range(.Cells(25, 17), .Cells(27, 19)).Borders.Weight = xlThin
        .Cells(25, 17).Value = "Desired Special"
        .Cells(25, 19).Value = "Variance"
        .Cells(26, 17).Value = "Total Units"
        .Cells(27, 17).Value = "Total POS Value"
        .Cells(26, 18).Value = "=SUM(R[-19]C:R[-12]C)"
        .Cells(26, 19).Value = "=(RC[-1]-R[-3]C[-3])/R[-3]C[-3]"
        .Cells(27, 19).Value = "=RC[-1]-R[-3]C[-3]"
        .Cells(27, 18).Value = "=(SUM(R7C4:R315C4)/SUM(R7C3:R315C3))*R26C18"
        Range(.Cells(25, 17), .Cells(25, 19)).Font.Bold = True
        Range(.Cells(25, 17), .Cells(27, 17)).Font.Bold = True
        Range(.Cells(27, 17), .Cells(27, 19)).Font.Bold = True
        
        
        
    End If
    
    If MultiDrop = True Then
        On Error GoTo multidroperror
        CBA_T_ShortlifeMAA.Copy , wks_Reorder
        Set wks_MultiDropMAA = ActiveSheet
        With wks_MultiDropMAA
            Application.EnableEvents = False
            Application.Calculation = xlCalculationManual
            For vDiv = 501 To 509
                If vDiv <> 508 Then
'                    If .Cells(6, 51).Value = "" Then
                        'divfnd = 0
                        For a = 0 To DateDiff("D", POSData(vDiv, LBound(POSData, 2)).DAT, POSData(vDiv, UBound(POSData, 2)).DAT)
                            If .Cells(6, 51 + a).Value = "" Or .Cells(6, 51 + a).Value = DateAdd("D", a, POSData(vDiv, LBound(POSData, 2)).DAT) Then
                                If .Cells(6, 51 + a).Value = "" Then
                                    .Cells(6, 51 + a).Value = DateAdd("D", a, POSData(vDiv, LBound(POSData, 2)).DAT)
                                    .Cells(6, 51 + a).NumberFormat = "DDDD"
                                End If
                                If vDiv = 509 Then
                                    .Cells(14, 51 + a).Value = POSData(vDiv, a + 1).QTY / SellT
                                Else
                                    .Cells(vDiv - 494, 51 + a).Value = POSData(vDiv, a + 1).QTY / SellT
                                End If
                                .Cells(16, 51 + a).Value = .Cells(16, 51 + a).Value + POSData(vDiv, a + 1).Retail
                            ElseIf POSData(vDiv, LBound(POSData, 2)).DAT = "12:00:00 AM" Then
                                If vDiv = 509 Then
                                    .Cells(14, 51 + a).Value = POSData(vDiv, a + 1).QTY / SellT
                                Else
                                    .Cells(vDiv - 494, 51 + a).Value = POSData(vDiv, a + 1).QTY / SellT
                                End If
                            Else
                                MsgBox "There has been an error with what regions are selling what on what day." & Chr(10) & Chr(10) & "Please contact Tom or Bob on 9218"
                                Exit Sub
                            End If
                        Next
                End If
            Next
            Application.EnableEvents = True
            Application.Calculation = xlCalculationAutomatic
            .Cells(8, 11).Value = 1
            .Cells(12, 11).Value = 1
            .Cells(12, 12).Value = 1
            .Cells(1, 1).Select
        End With
        On Error GoTo Err_Routine
    End If
    .Activate
    Range(.Cells(6, 15), .Cells(6, 17)).EntireColumn.AutoFit
    .Cells(1, 15).EntireColumn.ColumnWidth = 18.3
    .Cells(1, 16).EntireColumn.ColumnWidth = 13
    .Cells(1, 17).EntireColumn.ColumnWidth = 18.3
    ActiveWindow.FreezePanes = False
    .Cells(7, 1).Select
    ActiveWindow.FreezePanes = True
    End With
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    Application.ScreenUpdating = True
    
    
Exit Sub

multidroperror:
    Err.Clear
    On Error GoTo Err_Routine
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    Application.ScreenUpdating = True
    MsgBox "There has been an error. This may be due to incorrect values being entered." & Chr(10) & Chr(10) & "Please try again. If the problem persists please call Tom or Bob on 9218"
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-ReorderRuntime", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Sub CBA_BT_ReorderTool(Control As IRibbonControl)
    CBA_BT_ReorderMAA.Show vbModeless
End Sub
