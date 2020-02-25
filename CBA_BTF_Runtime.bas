Attribute VB_Name = "CBA_BTF_Runtime"
Option Explicit       ' CBA_BTF_Runtime @CBA_BTF Changed 181224
'''Option Private Module       ' Excel users cannot access procedures

Type CBA_BTF_ReportParamaters
    ReportName As String
    PSMonth As Long
    PSYear As Long
    PEMonth As Long
    PEYear As Long
    BD As String
    GBD As String
    CG As Long
    scg As Long
    ProductClass As Long
End Type
Type CBA_BTF_ForecastMetrics
    OriginalForecast As Single
    ReForecast As Single
    Actual As Single
End Type


Type CBA_BTF_CGDataDetail
    Year As Long
    Month As String
    POSQTY As Long
    POSPYQTY As Long
    POSYOYQTY As Single
    POSRET As Single
    POSPYRET As Single
    POSYOYRET As Single
    RCVMargin As Single
    RCVPYMargin As Single
    CPU As Single
    CPUYOY As Single
    CPUPY As Single
    ForeCPU As Single
    ForePrice As Single
    ForeQTY As Single
    ForeRetail As Single
    PForeRetail As Single
    ForePYRetail As Single
    ForePYQTY As Single
    ForeCost As Single
    ForeRetailNet As Single
    ForePYCost As Single
    ForePYRetailNet As Single
    ForeRCVMargin As Single
    PForeRCVMargin As Single
    PUplift As Single
    Uplift As Single

End Type

Type CBA_BTF_SaMData
    Year As Long
    Month As String
    TotSales As CBA_BTF_ForecastMetrics
    CRSales As CBA_BTF_ForecastMetrics
    FVSales As CBA_BTF_ForecastMetrics
    MeatSales As CBA_BTF_ForecastMetrics
    ChilSales As CBA_BTF_ForecastMetrics
    SpecialSales As CBA_BTF_ForecastMetrics
    SeasonalSales As CBA_BTF_ForecastMetrics
End Type


Private basedatapulled As Boolean
Private FMarginP  As Variant
Private FReMarginP As Variant
Private FReSales As Variant
Private FSales As Variant
Private CBA_SCGData() As CBA_BTF_SCG
Private CGs() As Variant
Private RCVRetNet As Single, RCVCost As Single, RCVRet As Single
Private TotRCVRetNet As Single, TotRCVRet As Single, TotRCVCost As Single
Private SummaryData() As Single

Public Function CBA_BTF_ChartChange(ByRef wks As Worksheet, ByVal ChartTitle As String, ByVal DataUsed As String, ByVal Lastrow As Long)
    Dim rng As Range, Head As Range, cht As Variant
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    For Each cht In wks.ChartObjects
        cht.Delete
    Next
    If DataUsed = "POS Quantity" Then
        Set rng = wks.Range("F7:G" & Lastrow)
        Set Head = wks.Range("F7:G7")
''        Set Rng = wks.Range("G" & LastRow & ":F7")
''        Set Head = wks.Range("G7:F7")
    ElseIf DataUsed = "Cost per Unit" Then
        Set rng = wks.Range("N7:O" & Lastrow)
        Set Head = wks.Range("N7:O7")
    ElseIf DataUsed = "POS Retail" Then
        Set rng = wks.Range("H7:I" & Lastrow)
        Set Head = wks.Range("H7:I7")
    End If
    cht = CCM_ChartCreate.CBA_BTF_ChartCreate(wks.Cells(7, 12).Left, wks.Cells(7, 12).Top, rng, wks, ChartTitle, Head, Range("D8:E" & Lastrow), xlBottom, xlLine, xlLine, xlColumns, False, 15, 10)
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_BTF_ChartChange", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Public Function CBA_BTF_pclassDecypher(ByVal Product_Class_as_String As String) As Variant
    ' Will return the Product Class variables
    Select Case Product_Class_as_String
        Case "Core Range"
            CBA_BTF_pclassDecypher = CLng(1)
        Case "Food Specials"
            CBA_BTF_pclassDecypher = CLng(2)
        Case "Non-Food Specials"
            CBA_BTF_pclassDecypher = CLng(3)
        Case "Seasonal"
            CBA_BTF_pclassDecypher = CLng(4)
        Case "1"
            CBA_BTF_pclassDecypher = "Core Range"
        Case "2"
            CBA_BTF_pclassDecypher = "Food Specials"
        Case "3"
            CBA_BTF_pclassDecypher = "Non-Food Specials"
        Case "4"
            CBA_BTF_pclassDecypher = "Seasonal"
    End Select

End Function
Private Function getSCGData() As CBA_BTF_SCG()
    getSCGData = CBA_SCGData
End Function
Public Function getCGs() As Variant
    getCGs = CGs
End Function
'Callback for CBA_BTF_Entry onAction
Public Sub CBA_BTF_Apply(Control As IRibbonControl)
    CBA_BTF_frm_Splash.Show
End Sub
'Callback for CBA_BTF_Query onAction
Public Sub CBA_BTF_Query(Control As IRibbonControl)
    CBA_BTF_frm_Reporting.Show
End Sub
Public Sub RunReport(ByRef r As CBA_BTF_ReportParamaters)
    Dim Data() As Variant
    Dim RP As CBA_BTF_ReportParamaters
    Dim wbk As Workbook
    Dim wst As Worksheet, wks(1 To 5) As Worksheet
    Dim NoOfMonths As Long, Rounder As Long, yr As Long, Mnth As Long, MTick As Long, Pc As Long, CG As Long, SalesFC As Long
    Dim SalesM As Long, oprow As Long, Mtype As Long, opcol As Long, mont As Long, YearFrom  As Long, YearTo As Long
    Dim MonthFrom As Long, MonthTo As Long, curM As Long
    Dim a As Long, b As Long, c As Long, lRowNo As Long, col
    Dim PCL As Long, TotalMonths As Long, StartY As Long, EndY As Long
    'Dim RCVRetNet As Single, RCVRet As Single, RCVCost As Single
    Dim SaMData() As CBA_BTF_SaMData
    Dim SaMContData() As CBA_BTF_SaMData
    Dim strClass As String
    Dim PClass As Long, lCurCG As Long, StartRow As Long
    Dim FLvl() As String, isFirstRun As Boolean
    Dim strDBUtoForecast As String
    'On Error GoTo Err_Routine
    CBA_ErrTag = ""
    RP = r
    
    Select Case RP.ReportName
        Case "DBU Period Forecast Report"
            CBA_SQL_Queries.CBA_GenPullSQL "DBU_List"
            strDBUtoForecast = ""
            For a = 0 To CBA_BTF_frm_Reporting.cbx_DBU.ListCount - 1
                If CBA_BTF_frm_Reporting.cbx_DBU.Selected(a) = True Then
                    For b = LBound(CBA_ABIarr, 2) To UBound(CBA_ABIarr, 2)
                        If Len(CBA_ABIarr(2, b)) <> 0 Then
                            If Left(CBA_ABIarr(2, b), 1) = Chr(10) Then
                                If Mid(CBA_ABIarr(2, b), 2, Len(CBA_ABIarr(2, b)) - 1) = CBA_BTF_frm_Reporting.cbx_DBU.List(a) Then
                                    If strDBUtoForecast = "" Then strDBUtoForecast = CBA_ABIarr(1, b) Else strDBUtoForecast = strDBUtoForecast & "," & CBA_ABIarr(1, b)
                                    Exit For
                                End If
                            ElseIf Right(CBA_ABIarr(2, b), 1) = Chr(10) Then
                                If Left(CBA_ABIarr(2, b), Len(CBA_ABIarr(2, b)) - 1) = CBA_BTF_frm_Reporting.cbx_DBU.List(a) Then
                                    If strDBUtoForecast = "" Then strDBUtoForecast = CBA_ABIarr(1, b) Else strDBUtoForecast = strDBUtoForecast & "," & CBA_ABIarr(1, b)
                                    Exit For
                                End If
                            Else
                                If CBA_ABIarr(2, b) = CBA_BTF_frm_Reporting.cbx_DBU.List(a) Then
                                    If strDBUtoForecast = "" Then strDBUtoForecast = CBA_ABIarr(1, b) Else strDBUtoForecast = strDBUtoForecast & "," & CBA_ABIarr(1, b)
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
            Next
            Unload CBA_BTF_frm_Reporting
            If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "DBU Period Forecast Report"
            CBA_DBU_Runtime.RunDBUForecastReport DateSerial(RP.PSYear, RP.PSMonth, 1), DateSerial(RP.PEYear, RP.PEMonth + 1, 0), strDBUtoForecast
        Case "CG/SCG Sales and Margin Forecast Report"
            
            If RP.PSYear = 0 Then
                MsgBox "No Year has been selected"
                Exit Sub
            End If
           
            YearFrom = RP.PSYear - 2
            YearTo = RP.PSYear
            MonthFrom = 1
            MonthTo = 12
            
            CreateCGsArray RP
            Unload CBA_BTF_frm_Reporting
            CBA_BasicFunctions.CBA_Running "CG/SCG Sales and Margin Forecast Report"
            If CBA_BTF_SetupForecastArray.CBA_BTF_SetupForecastArray(YearFrom, YearTo, MonthFrom, MonthTo) = True Then
           
                ReDim SummaryData(1 To 199, 1 To 5, 2 To 32)
                Set wbk = Workbooks.Add
                Set wks(1) = wbk.Worksheets("Sheet1"): wks(1).Name = "Core Range"
                wbk.Worksheets.Add ''''', , wks(1)
                Set wks(2) = wbk.Worksheets("Sheet2"): wks(2).Name = "Food Special"
                wbk.Worksheets.Add ''''', , wks(2)
                Set wks(3) = wbk.Worksheets("Sheet3"): wks(3).Name = "Non-Food Special"
                wbk.Worksheets.Add , wks(3)           ' Was 3
                Set wks(4) = wbk.Worksheets("Sheet4"): wks(4).Name = "Seasonal"
                wbk.Worksheets.Add wks(1)             ' Was 1
                Set wks(5) = wbk.Worksheets("Sheet5"): wks(5).Name = "Summary"

'                Application.ScreenUpdating = False
                
                ''ActiveWindow.WindowState = xlMinimized                    '#RW Added line
                
                If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Preparing Report..."
''                      wbk.Application.mi
''52500                 wbk.Worksheets.Visible = False                        '#RW Added line
                
                On Error Resume Next
                If UBound(CGs, 2) < 0 Then
                    RP.GBD = "ALL"
                    CreateCGsArray RP
                End If
                On Error GoTo Err_Routine
                
                For PClass = 1 To 5
                    If (r.ProductClass > 0 And r.ProductClass = PClass) Or r.ProductClass = 0 Or PClass = 5 Then
                        With wks(PClass)
                            wbk.Activate
                            .Activate
                            Range(.Cells(1, 1), .Cells(5, 79)).Interior.ColorIndex = 49
                            .Cells(1, 1).Select
                            .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
                            .Cells(4, 3).Font.Name = "ALDI SUED Office"
                            .Cells(4, 3).Font.Size = 24
                            .Cells(4, 3).Font.ColorIndex = 2
                            Rows("1:5").Font.ThemeColor = xlThemeColorDark1
                            Range("I1:R1,I2:J2,K2:R2,B10:D10,E10:G10,H10:J10,AE9:AF9,AC9:AD9,AC8:AF8,Z10:AB10,W10:Y10,W9:AB9,Q8:AB8,Q9:V9,T10:V10,Q10:S10,N8:P9,K8:M9,B8:J9").Merge
'                            Range("I2:R3").Borders.Weight = xlThin
'                            Range("I2:R3").Borders.ColorIndex = 2
                            Range("I1").Font.Underline = xlUnderlineStyleSingle
                            Range("I1").Font.Bold = True
                            For a = 2 To 32
                                Select Case a
                                    Case 2, 5, 8
                                        .Cells(11, a).Value = "Original Forecast"
                                    Case 3, 6, 9
                                        .Cells(11, a).Value = "Re-Forecast"
                                    Case 17, 20, 23, 26
                                        .Cells(11, a).Value = "Original Forecast " & RP.PSYear & " vs " & RP.PSYear - 1
                                    Case 18, 21, 24, 27
                                        .Cells(11, a).Value = "Re-Forecast" & RP.PSYear & " vs " & RP.PSYear - 1
                                    Case 29 To 32
                                        .Cells(11, a).Value = "Actual" & RP.PSYear - 1 & " vs " & RP.PSYear - 2
                                    Case Else
                                        .Cells(11, a).Value = "Actual"
                                End Select
                            Next
                            Range("A10:AF12,B8:J12,K8:M12,N8:P12,Q8:AF12").Borders.Weight = xlThin
                            Range("B12:D18,E12:G18,H12:J18,K12:L18,N12:O18,Q12:V18,AC12:AD18").NumberFormat = "$#,0.00"
                            Range("M12:M18,P12:P18,W12:AB18,AE12:AF18").NumberFormat = "0.00%"
                            Rows("8:13").HorizontalAlignment = xlCenter
                            
'                            Range("I1").Value = "Report Paramaters"
'                            Range("I2").Value = "Time Period:"
'                            Range("I3").Value = "CG:"
'                            Range("K3").Value = "SCG:"
'                            Range("M3").Value = "BD:"
'                            Range("O3").Value = "GBD:"
'                            Range("Q3").Value = "ProductClass:"
                            Range("A8").Value = "Overall"
                            Range("A9").Value = "Sales / Margin"
                            Range("A10").Value = "Category"
                            Range("B10,K10,N10").Value = "POS Sales $"
                            Range("E10,L10,O10").Value = "RCV Margin $"
                            Range("H10,M10,P10").Value = "RCV Margin %"
                            Range("Z10,AF10,T10,AD10").Value = "Margin"
                            Range("W10,AE10,Q10,AC10").Value = "Sales"
                            Range("AE9,W9").Value = "Growth %"
                            Range("AC9,Q9").Value = "Growth $"
                                            
                            'YEARS
                            Range("Q8,B8").Value = RP.PSYear
                            Range("AC8,K8").Value = RP.PSYear - 1
                            Range("N8").Value = RP.PSYear - 2
                            
                            
                            If RP.GBD <> "" Then
                                .Cells(3, 16).Value = RP.GBD
                            ElseIf RP.BD <> "" Then
                                .Cells(3, 14).Value = RP.BD
                            ElseIf RP.CG <> 0 Then
                                .Cells(3, 10).Value = RP.CG
                                If RP.scg <> 0 Then .Cells(3, 12).Value = RP.scg
                            End If
'                            If pclass <> 0 Then
'                                If pclass = 1 Then .Cells(3, 18) = "Core"
'                                If pclass = 2 Then .Cells(3, 18) = "FoodSpecial"
'                                If pclass = 3 Then .Cells(3, 18) = "NonFoodSpecial"
'                                If pclass = 4 Then .Cells(3, 18) = "Seasonal"
'                                If pclass = 5 Then .Cells(3, 18) = "Summary"
'                            End If
                            .Cells(2, 11).Value = RP.PSYear - 2 & " through " & RP.PSYear
                            lRowNo = 11
                            lCurCG = 0
                            
                            If PClass = 5 Then
                                
                                For CG = 1 To 199
                                    If SummaryData(CG, 1, 11) <> 0 Or SummaryData(CG, 2, 11) <> 0 Or SummaryData(CG, 3, 11) <> 0 Or SummaryData(CG, 4, 11) <> 0 Or SummaryData(CG, 1, 2) <> 0 Or SummaryData(CG, 2, 2) <> 0 Or SummaryData(CG, 3, 2) <> 0 Or SummaryData(CG, 4, 2) <> 0 Then
                                        StartRow = lRowNo + 1
                                        lCurCG = CG
                                        For PCL = 1 To 4
                                            If SummaryData(CG, PCL, 11) <> 0 Or SummaryData(CG, PCL, 2) <> 0 Then
                                                lRowNo = lRowNo + 1
                                                If PCL = 1 Then
                                                    strClass = ": Core"
                                                ElseIf PCL = 2 Then
                                                    strClass = ": F-Spec"
                                                ElseIf PCL = 3 Then
                                                    strClass = ": NF-Spec"
                                                ElseIf PCL = 4 Then
                                                    strClass = ": Season"
                                                End If
                                            
                                                wks(5).Cells(lRowNo, 1).Value = "CG " & CG & strClass
                                                For b = 2 To 32
                                                    wks(5).Cells(lRowNo, b) = SummaryData(CG, PCL, b)
                                                Next
                                            End If
                                        Next
                                        
                                        lRowNo = lRowNo + 1
                                        wks(5).Cells(lRowNo, 1).Value = "CG " & CG
                                        For a = 2 To 32
                                            b = 0: c = 0
                                            Select Case a
                                                Case 1
                                                    Range(.Cells(12, a), .Cells(lRowNo, a)).NumberFormat = "General"
                                                Case 2 To 7, 11, 12, 14, 15, 17 To 22, 29, 30
                                                    wks(5).Cells(lRowNo, a).Value = Application.WorksheetFunction.Sum(Range(wks(5).Cells(StartRow, a), wks(5).Cells(lRowNo - 1, a)))
                                                    Range(.Cells(StartRow, a), .Cells(lRowNo, a)).NumberFormat = "$#,0"
                                                Case 8, 9, 10
                                                    b = 6: c = 3
                                                    Range(.Cells(StartRow, a), .Cells(lRowNo, a)).NumberFormat = "#,0.0%"
                                                Case 13, 16
                                                    b = 2: c = 1
                                                    Range(.Cells(StartRow, a), .Cells(lRowNo, a)).NumberFormat = "#,0.0%"
                                                Case 23 To 28
                                                    b = 21: c = 6
                                                    Range(.Cells(StartRow, a), .Cells(lRowNo, a)).NumberFormat = "#,0.0%"
                                                Case 31, 32
                                                    b = 17: c = 2
                                                    Range(.Cells(StartRow, a), .Cells(lRowNo, a)).NumberFormat = "#,0.0%"
                                            End Select
                                            If b > 0 Then
                                                If wks(5).Cells(lRowNo, a - b).Value = 0 Then
                                                    wks(5).Cells(lRowNo, a).Value = 0
                                                Else
                                                    wks(5).Cells(lRowNo, a).Value = wks(5).Cells(lRowNo, a - c).Value / wks(5).Cells(lRowNo, a - b).Value
                                                End If
                                            End If
                                        Next
                                        Range(.Cells(StartRow, 1), .Cells(lRowNo, 32)).Borders.Weight = xlThin
                                        'Range(.Cells(lRowNo, a), .Cells(lRowNo, a)).Borders(xlEdgeTop).Weight = xlThick
                                        Range(wks(5).Cells(StartRow, 1), wks(5).Cells(lRowNo - 1, 1)).RowHeight = 0
                                        
                                    End If
                                Next
                                SetCGTotal wks(PClass), 12, lCurCG, "CG", lRowNo, PClass, True
                                .Cells.EntireColumn.AutoFit
                                .Columns(3).ColumnWidth = 15
                                .Cells(4, 3).Value = "CG/SCG Forecast Summary"
                                .Activate
                                .Cells(12, 1).Select
                                ActiveWindow.FreezePanes = True
                                GoTo exitTheReport
                            End If
                            
'                            If pclass = 3 Then
'                            a = a
'                            End If
'
                            
                            
                            
                            For a = LBound(CGs, 2) To UBound(CGs, 2)
                                lRowNo = lRowNo + 1
                                If CGs(0, a) <> lCurCG Then
                                    If lCurCG <> 0 Then
                                        If FCbM(RP.PSYear, 1, PClass).CGData(lCurCG, 0).Level = "SCG" Or FCbM(RP.PSYear, 1, PClass).CGData(lCurCG, 0).Level = "CG" Then
                                            SetCGTotal wks(PClass), StartRow, lCurCG, FCbM(RP.PSYear, 1, PClass).CGData(lCurCG, 0).Level, lRowNo, PClass, False
                                        End If
                                    End If
                                    lCurCG = CGs(0, a)
                                    isFirstRun = True
                                    StartRow = lRowNo
                                    TotRCVRetNet = 0: TotRCVRet = 0: TotRCVCost = 0
                                Else

                                    If FCbM(RP.PSYear, 1, PClass).CGData(lCurCG, 0).Level = "CG" Or (FCbM(RP.PSYear, 1, PClass).CGData(lCurCG, 0).Level = "SCG" And CGs(1, a) = 0) Then
                                        lRowNo = lRowNo - 1
                                        GoTo letsdothenextlineintheCGarray
                                    End If
                                End If

                                If FCbM(RP.PSYear, 1, PClass).CGData(lCurCG, 0).Level = "CG" Then
                                    .Cells(lRowNo, 1).Value = "CG: " & CGs(0, a)
                                ElseIf FCbM(RP.PSYear, 1, PClass).CGData(lCurCG, 0).Level = "SCG" Then
                                    .Cells(lRowNo, 1).Value = "CG: " & CGs(0, a) & "-" & CGs(1, a)
                                Else
                                    For b = 1 To 12
                                        FCbM(RP.PSYear, b, PClass).CGData(lCurCG, 0).setLevel "SCG"
                                    Next
                                    .Cells(lRowNo, 1).Value = "CG: " & CGs(0, a) & "-" & CGs(1, a)

'                                    lRowNo = lRowNo - 1
'                                    GoTo letsdothenextlineintheCGarray
                                End If
                                

                                RCVRetNet = 0: RCVCost = 0: RCVRet = 0
                                For Mnth = 1 To 12
                                    If RP.PSYear >= Year(Date) Then
                                        If FCbM(RP.PSYear, Mnth, PClass).CGData(lCurCG, 0).Level = "None" Then
                                              a = a
                                        ElseIf FCbM(RP.PSYear, Mnth, PClass).CGData(lCurCG, 0).Level = "CG" Then
                                              putDataOnSheet wks(PClass), PClass, Mnth, RP.PSYear, "CG", lCurCG, lRowNo
                                        ElseIf FCbM(RP.PSYear, Mnth, PClass).CGData(lCurCG, 0).Level = "SCG" Then
                                              'If (CGs(1, a) = "0" And isFirstRun = True) Or CGs(1, a) <> 0 Then
                                                  CGs(1, a) = Val(CGs(1, a))
                                                  putDataOnSheet wks(PClass), PClass, Mnth, RP.PSYear, "SCG", lCurCG, lRowNo, CGs(1, a)
                                              'End If
                                        End If
                                    End If
                                Next

                                
                                For c = 1 To 6
                                    .Cells(lRowNo, c + 16).Value = .Cells(lRowNo, 32 + c).Value
                                Next
        
                                If .Cells(lRowNo, 2).Value = 0 Then .Cells(lRowNo, 8).Value = 0 Else .Cells(lRowNo, 8).Value = .Cells(lRowNo, 5).Value / .Cells(lRowNo, 2).Value
                                If .Cells(lRowNo, 3).Value = 0 Then .Cells(lRowNo, 9).Value = 0 Else .Cells(lRowNo, 9).Value = .Cells(lRowNo, 6).Value / .Cells(lRowNo, 3).Value
                                If RCVRet = 0 Then .Cells(lRowNo, 10).Value = 0 Else .Cells(lRowNo, 10).Value = (RCVRetNet - RCVCost) / RCVRet
                                For b = 17 To 22
                                    If .Cells(lRowNo, b).Value = 0 Then .Cells(lRowNo, b + 6).Value = 0 Else .Cells(lRowNo, b + 6).Value = (.Cells(lRowNo, b - 15).Value - .Cells(lRowNo, b).Value) / .Cells(lRowNo, b).Value
                                    .Cells(lRowNo, b).Value = .Cells(lRowNo, b - 15).Value - .Cells(lRowNo, b).Value
                                Next
                                RCVRetNet = 0: RCVRet = 0: RCVCost = 0 ': PYRCVRetNet = 0: PYRCVRet = 0: PYRCVCost = 0
                                For c = 1 To 2
                                    For Mnth = 1 To 12
                                        If FCbM(RP.PSYear, Mnth, PClass).CGData(lCurCG, 0).Level = "SCG" Then
                                            .Cells(lRowNo, 11 + ((c - 1) * 3)).Value = .Cells(lRowNo, 11 + ((c - 1) * 3)).Value + FCbM(RP.PSYear - c, Mnth, PClass).CGData(CLng(CGs(0, a)), CLng(CGs(1, a))).Sales.Actual
                                            .Cells(lRowNo, 12 + ((c - 1) * 3)).Value = .Cells(lRowNo, 12 + ((c - 1) * 3)).Value + FCbM(RP.PSYear - c, Mnth, PClass).CGData(CLng(CGs(0, a)), CLng(CGs(1, a))).MarginDol.Actual
                                            RCVRetNet = RCVRetNet + FCbM(RP.PSYear - c, Mnth, PClass).CGData(CLng(CGs(0, a)), CLng(CGs(1, a))).RCVRetailNet
                                            RCVRet = RCVRet + FCbM(RP.PSYear - c, Mnth, PClass).CGData(CLng(CGs(0, a)), CLng(CGs(1, a))).RCVRetail
                                            RCVCost = RCVCost + FCbM(RP.PSYear - c, Mnth, PClass).CGData(CLng(CGs(0, a)), CLng(CGs(1, a))).Cost.Actual
                                        ElseIf FCbM(RP.PSYear, Mnth, PClass).CGData(lCurCG, 0).Level = "CG" Then
                                           'why no value for 2018 actuals?
                                            .Cells(lRowNo, 11 + ((c - 1) * 3)).Value = .Cells(lRowNo, 11 + ((c - 1) * 3)).Value + FCbM(RP.PSYear - c, Mnth, PClass).CGData(CLng(CGs(0, a)), 0).Sales.Actual
                                            'hmmm?
                                            .Cells(lRowNo, 12 + ((c - 1) * 3)).Value = .Cells(lRowNo, 12 + ((c - 1) * 3)).Value + FCbM(RP.PSYear - c, Mnth, PClass).CGData(CLng(CGs(0, a)), 0).MarginDol.Actual
                                            RCVRetNet = RCVRetNet + FCbM(RP.PSYear - c, Mnth, PClass).CGData(CLng(CGs(0, a)), 0).RCVRetailNet
                                            RCVRet = RCVRet + FCbM(RP.PSYear - c, Mnth, PClass).CGData(CLng(CGs(0, a)), 0).RCVRetail
                                            RCVCost = RCVCost + FCbM(RP.PSYear - c, Mnth, PClass).CGData(CLng(CGs(0, a)), 0).Cost.Actual
                                        ElseIf FCbM(RP.PSYear, Mnth, PClass).CGData(lCurCG, 0).Level = "None" Then
                                            .Cells(lRowNo, 11 + ((c - 1) * 3)).Value = 0: .Cells(lRowNo, 12 + ((c - 1) * 3)).Value = 0: RCVRetNet = 0: RCVRet = 0: RCVCost = 0
                                        End If
                                    Next
                                    If RCVRet = 0 Then .Cells(lRowNo, 13 + ((c - 1) * 3)).Value = 0 Else .Cells(lRowNo, 13 + ((c - 1) * 3)).Value = (RCVRetNet - RCVCost) / RCVRet
                                Next
                                .Cells(lRowNo, 29).Value = .Cells(lRowNo, 11).Value - .Cells(lRowNo, 14).Value
                                If .Cells(lRowNo, 14).Value = 0 Then .Cells(lRowNo, 31).Value = 0 Else .Cells(lRowNo, 31).Value = (.Cells(lRowNo, 11).Value - .Cells(lRowNo, 14).Value) / .Cells(lRowNo, 14).Value
                                .Cells(lRowNo, 30).Value = .Cells(lRowNo, 12).Value - .Cells(lRowNo, 15).Value
                                If .Cells(lRowNo, 15).Value = 0 Then .Cells(lRowNo, 32).Value = 0 Else .Cells(lRowNo, 32).Value = (.Cells(lRowNo, 12).Value - .Cells(lRowNo, 15).Value) / .Cells(lRowNo, 15).Value
                                TotRCVRetNet = TotRCVRetNet + RCVRetNet
                                TotRCVRet = TotRCVRet + RCVRet
                                TotRCVCost = TotRCVCost + RCVCost
letsdothenextlineintheCGarray:
                            Next
                            
                            If (FCbM(RP.PSYear, 1, PClass).CGData(lCurCG, 0).Level = "SCG" Or FCbM(RP.PSYear, 1, PClass).CGData(lCurCG, 0).Level = "CG") And lCurCG > 0 Then
                                lRowNo = lRowNo + 1
                                SetCGTotal wks(PClass), StartRow, lCurCG, FCbM(RP.PSYear, 1, PClass).CGData(lCurCG, 0).Level, lRowNo, PClass, False
                            End If
                            SetCGTotal wks(PClass), 12, lCurCG, FCbM(RP.PSYear, 1, PClass).CGData(lCurCG, 0).Level, lRowNo, PClass, True
                            Range(.Cells(12, 33), .Cells(lRowNo, 38)).Clear
                            Cells.EntireColumn.AutoFit
                            Columns(3).ColumnWidth = 15
                            .Cells(4, 3).Value = "CG/SCG Sales and Margin Forecast Report"
                            .Activate
                            .Cells(12, 1).Select
                            ActiveWindow.FreezePanes = True
                            
                        End With
                Else
                    Application.DisplayAlerts = False
                    wks(PClass).Delete
                    Application.DisplayAlerts = True
                End If
            Next
            Erase FCbM
            Application.ScreenUpdating = True
            If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.CBA_Close_Running
        End If
        
    Case "Admin S&M Report"
        Dim endm As Byte, MT As Byte, startm As Byte
        Dim strFormu As String
            
            If RP.PSYear = 0 Then MsgBox "No Year Start has been selected": Exit Sub
            If RP.PEYear = 0 Then MsgBox "No Year End has been selected": Exit Sub
            If RP.PSMonth = 0 Then MsgBox "No Month Start has been selected": Exit Sub
            If RP.PEMonth = 0 Then MsgBox "No Month End has been selected": Exit Sub

            
           
            YearFrom = RP.PSYear
            YearTo = RP.PEYear
            MonthFrom = RP.PSMonth
            MonthTo = RP.PEMonth
            
            TotalMonths = DateDiff("M", DateSerial(RP.PSYear, RP.PSMonth, 1), DateSerial(RP.PEYear, RP.PEMonth + 1, 0)) + 1
            ReDim SaMData(1 To TotalMonths)
            ReDim SaMContData(1 To TotalMonths)
            
            
            CreateCGsArray RP
            Unload CBA_BTF_frm_Reporting
            CBA_BasicFunctions.CBA_Running "Admin S&M Report"
            
            
            If CBA_BTF_SetupForecastArray.CBA_BTF_SetupForecastArray(YearFrom, YearTo, MonthFrom, MonthTo) = True Then
                Application.ScreenUpdating = False
                Set wbk = Workbooks.Add
                Set wks(1) = ActiveSheet
                With wks(1)
                    curM = 0
                    For yr = YearFrom To YearTo
                        If yr = YearTo Then endm = MonthTo Else endm = 12
                        If yr = YearFrom Then startm = MonthFrom Else startm = 1
                        For MT = startm To endm
                            curM = curM + 1
                            For CG = 1 To 70
                                For PClass = 1 To 4
                                    If PClass = 1 Then
                                        If CG = 58 Then
                                            SaMData(curM).FVSales.Actual = SaMData(curM).FVSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.Actual
                                            SaMData(curM).FVSales.OriginalForecast = SaMData(curM).FVSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.OriginalForecast
                                            SaMData(curM).FVSales.ReForecast = SaMData(curM).FVSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.ReForecast
                                            SaMContData(curM).FVSales.Actual = SaMContData(curM).FVSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.Actual
                                            SaMContData(curM).FVSales.OriginalForecast = SaMContData(curM).FVSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.Actual
                                            SaMContData(curM).FVSales.ReForecast = SaMContData(curM).FVSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.Actual
                                        ElseIf CG = 51 Then
                                            SaMData(curM).ChilSales.Actual = SaMData(curM).ChilSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.Actual
                                            SaMData(curM).ChilSales.OriginalForecast = SaMData(curM).ChilSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.OriginalForecast
                                            SaMData(curM).ChilSales.ReForecast = SaMData(curM).ChilSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.ReForecast
                                            SaMContData(curM).ChilSales.Actual = SaMContData(curM).ChilSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.Actual
                                            SaMContData(curM).ChilSales.OriginalForecast = SaMContData(curM).ChilSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.OriginalForecast
                                            SaMContData(curM).ChilSales.ReForecast = SaMContData(curM).ChilSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.ReForecast
                                        ElseIf CG = 62 Or CG = 63 Then
                                            SaMData(curM).MeatSales.Actual = SaMData(curM).MeatSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.Actual
                                            SaMData(curM).MeatSales.OriginalForecast = SaMData(curM).MeatSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.OriginalForecast
                                            SaMData(curM).MeatSales.ReForecast = SaMData(curM).MeatSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.ReForecast
                                            SaMContData(curM).MeatSales.Actual = SaMContData(curM).MeatSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.Actual
                                            SaMContData(curM).MeatSales.OriginalForecast = SaMContData(curM).MeatSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.OriginalForecast
                                            SaMContData(curM).MeatSales.ReForecast = SaMContData(curM).MeatSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.ReForecast
                                        Else
                                            SaMData(curM).CRSales.Actual = SaMData(curM).CRSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.Actual
                                            SaMData(curM).CRSales.OriginalForecast = SaMData(curM).CRSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.OriginalForecast
                                            SaMData(curM).CRSales.ReForecast = SaMData(curM).CRSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.ReForecast
                                            SaMContData(curM).CRSales.Actual = SaMContData(curM).CRSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.Actual
                                            SaMContData(curM).CRSales.OriginalForecast = SaMContData(curM).CRSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.OriginalForecast
                                            SaMContData(curM).CRSales.ReForecast = SaMContData(curM).CRSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.ReForecast
                                        End If
                                        SaMData(curM).TotSales.Actual = SaMData(curM).TotSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.Actual
                                        SaMData(curM).TotSales.OriginalForecast = SaMData(curM).TotSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.OriginalForecast
                                        SaMData(curM).TotSales.ReForecast = SaMData(curM).TotSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.ReForecast
                                        SaMContData(curM).TotSales.Actual = SaMContData(curM).TotSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.Actual
                                        SaMContData(curM).TotSales.OriginalForecast = SaMContData(curM).TotSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.OriginalForecast
                                        SaMContData(curM).TotSales.ReForecast = SaMContData(curM).TotSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.ReForecast
                                    ElseIf PClass = 2 Or PClass = 3 Then
                                        SaMData(curM).SpecialSales.Actual = SaMData(curM).SpecialSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.Actual
                                        SaMData(curM).SpecialSales.OriginalForecast = SaMData(curM).SpecialSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.OriginalForecast
                                        SaMData(curM).SpecialSales.ReForecast = SaMData(curM).SpecialSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.ReForecast
                                        SaMData(curM).TotSales.Actual = SaMData(curM).TotSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.Actual
                                        SaMData(curM).TotSales.OriginalForecast = SaMData(curM).TotSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.OriginalForecast
                                        SaMData(curM).TotSales.ReForecast = SaMData(curM).TotSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.ReForecast
                                        SaMContData(curM).SpecialSales.Actual = SaMContData(curM).SpecialSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.Actual
                                        SaMContData(curM).SpecialSales.OriginalForecast = SaMContData(curM).SpecialSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.OriginalForecast
                                        SaMContData(curM).SpecialSales.ReForecast = SaMContData(curM).SpecialSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.ReForecast
                                        SaMContData(curM).TotSales.Actual = SaMContData(curM).TotSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.Actual
                                        SaMContData(curM).TotSales.OriginalForecast = SaMContData(curM).TotSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.OriginalForecast
                                        SaMContData(curM).TotSales.ReForecast = SaMContData(curM).TotSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.ReForecast
                                    Else
                                        SaMData(curM).SeasonalSales.Actual = SaMData(curM).SeasonalSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.Actual
                                        SaMData(curM).SeasonalSales.OriginalForecast = SaMData(curM).SeasonalSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.OriginalForecast
                                        SaMData(curM).SeasonalSales.ReForecast = SaMData(curM).SeasonalSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.ReForecast
                                        SaMData(curM).TotSales.Actual = SaMData(curM).TotSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.Actual
                                        SaMData(curM).TotSales.OriginalForecast = SaMData(curM).TotSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.OriginalForecast
                                        SaMData(curM).TotSales.ReForecast = SaMData(curM).TotSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).Sales.ReForecast
                                        SaMContData(curM).SeasonalSales.Actual = SaMContData(curM).SeasonalSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.Actual
                                        SaMContData(curM).SeasonalSales.OriginalForecast = SaMContData(curM).SeasonalSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.OriginalForecast
                                        SaMContData(curM).SeasonalSales.ReForecast = SaMContData(curM).SeasonalSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.ReForecast
                                        SaMContData(curM).TotSales.Actual = SaMContData(curM).TotSales.Actual + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.Actual
                                        SaMContData(curM).TotSales.OriginalForecast = SaMContData(curM).TotSales.OriginalForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.OriginalForecast
                                        SaMContData(curM).TotSales.ReForecast = SaMContData(curM).TotSales.ReForecast + FCbM(yr, MT, PClass).CGData(CG, 0).MarginDol.ReForecast
                                    End If
                                Next
                            Next
                        Next
                    Next
                
                    For a = 8 To 37
                        Select Case a
                            Case 8, 10, 14, 18, 22, 26, 30, 34
                                .Cells(a, 4).Value = "Sales"
                            Case 9, 12, 16, 20, 24, 28, 32, 36
                                .Cells(a, 4).Value = "Margin"
                            Case 11, 15, 19, 23, 27, 31, 35
                                .Cells(a, 4).Value = "Share"
                            Case 13, 17, 21, 25, 29, 33, 37
                                .Cells(a, 4).Value = "Contribution"
                        End Select
                    Next
                    .Cells(8, 3).Value = "Total Business"
                    .Cells(10, 3).Value = "Core Range"
                    .Cells(11, 3).Value = "Including F&V"
                    .Cells(12, 3).Value = "Chilled, Meat"
                    .Cells(14, 3).Value = "Core Range"
                    .Cells(15, 3).Value = "Excluding"
                    .Cells(16, 3).Value = "Chilled, Meat"
                    .Cells(18, 3).Value = "Fruit & Veg"
                    .Cells(19, 3).Value = "58"
                    .Cells(22, 3).Value = "Chilled"
                    .Cells(23, 3).Value = "51"
                    .Cells(26, 3).Value = "Fresh Meat"
                    .Cells(27, 3).Value = "62,64"
                    .Cells(30, 3).Value = "Specials"
                    .Cells(34, 3).Value = "Seasonal"
                    .Cells(35, 3).Value = "Planned"
                    .Range(.Cells(8, 2), .Cells(37, 2)).Font.Bold = True
                    .Cells.Font.Name = "ALDI SUED Office"
                    .Rows(7).Font.Size = 9
                    
                    For a = 3 To 4
                        .Columns(a).Font.Size = 9
                        .Columns(a).Font.Bold = True
                        .Columns(a).ColumnWidth = 11
                    Next
                    
                    Range(.Cells(1, 1), .Cells(4, 199)).Interior.ColorIndex = 49
                    .Cells(1, 1).Select
                    .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
                    .Cells(2, 6).Value = "Sales and Margin Forecast Report (POS)"
                    .Cells(2, 6).Font.Name = "ALDI SUED Office"
                    .Cells(2, 6).Font.Size = 22
                    .Cells(2, 6).Font.ColorIndex = 2
                    StartY = RP.PSYear
                    For a = 1 To TotalMonths + 1
                        If a = TotalMonths + 1 Then
                            strFormu = "=SUM("
                            For b = 5 To (5 + (4 * (a - 1))) - 1 Step 4
                                If strFormu = "=SUM(" Then strFormu = strFormu & "RC[-" & b - 1 & "]" Else strFormu = strFormu & ",RC[-" & b - 1 & "]"
                            Next
                            If strFormu = "=SUM(" Then strFormu = "" Else strFormu = strFormu & ")"
                        End If
                        startm = RP.PSMonth + (a - 1)
                        If a > 1 And startm = 1 Then StartY = StartY + 1
                        With Range(.Cells(6, (5 + (4 * (a - 1)))), .Cells(6, 4 + (4 * a)))
                            .Merge
                            If a = TotalMonths + 1 Then
                                .Value = "TOTAL"
                            Else
                                .Value = Format(DateSerial(StartY, startm, 1), "MMMM-YYYY")
                            End If
                            .Borders.Weight = xlThick
                            .Font.Bold = True
                            .HorizontalAlignment = xlCenter
                        End With
                        'Range(.Cells(7, (5 + (4 * (a - 1)))), .Cells(37, 4 + (4 * a))).Select
                        Range(.Cells(7, (5 + (4 * (a - 1)))), .Cells(37, 4 + (4 * a))).Borders.Weight = xlThin
                        Range(.Cells(7, (5 + (4 * (a - 1)))), .Cells(37, (5 + (4 * (a - 1))))).BorderAround , xlMedium
                        Range(.Cells(7, 1 + (5 + (4 * (a - 1)))), .Cells(37, 1 + (5 + (4 * (a - 1))))).BorderAround , xlMedium
                        Range(.Cells(7, 2 + (5 + (4 * (a - 1)))), .Cells(37, 2 + (5 + (4 * (a - 1))))).BorderAround , xlMedium
                        Range(.Cells(7, 3 + (5 + (4 * (a - 1)))), .Cells(37, 3 + (5 + (4 * (a - 1))))).BorderAround , xlMedium
                        Range(.Cells(8, (5 + (4 * (a - 1)))), .Cells(9, 4 + (4 * a))).BorderAround , xlMedium
                        Range(.Cells(10, (5 + (4 * (a - 1)))), .Cells(13, 4 + (4 * a))).BorderAround , xlMedium
                        Range(.Cells(14, (5 + (4 * (a - 1)))), .Cells(17, 4 + (4 * a))).BorderAround , xlMedium
                        Range(.Cells(18, (5 + (4 * (a - 1)))), .Cells(21, 4 + (4 * a))).BorderAround , xlMedium
                        Range(.Cells(22, (5 + (4 * (a - 1)))), .Cells(25, 4 + (4 * a))).BorderAround , xlMedium
                        Range(.Cells(26, (5 + (4 * (a - 1)))), .Cells(29, 4 + (4 * a))).BorderAround , xlMedium
                        Range(.Cells(30, (5 + (4 * (a - 1)))), .Cells(33, 4 + (4 * a))).BorderAround , xlMedium
                        Range(.Cells(34, (5 + (4 * (a - 1)))), .Cells(37, 4 + (4 * a))).BorderAround , xlMedium
                        
                        Range(.Cells(7, (5 + (4 * (a - 1)))), .Cells(7, 4 + (4 * a))).Font.Bold = True
                        Range(.Cells(7, (5 + (4 * (a - 1)))), .Cells(7, 4 + (4 * a))).BorderAround , xlThick
                        Range(.Cells(7, (5 + (4 * (a - 1)))), .Cells(37, 4 + (4 * a))).HorizontalAlignment = xlCenter
                        'Range(.Cells(8, (5 + (4 * (a - 1)))), .Cells(8, 4 + (2 * a))).Font.bOld = True
                        Range(.Cells(8, (5 + (4 * (a - 1)))), .Cells(37, 4 + (4 * a))).BorderAround , xlThick
                        .Cells(7, (5 + (4 * (a - 1)))).Value = "ORIG 2019" & Chr(10) & "FORECAST"
                        .Cells(7, 1 + (5 + (4 * (a - 1)))).Value = "2019" & Chr(10) & "Re-Forecast"
                        .Cells(7, 2 + (5 + (4 * (a - 1)))).Value = "Actual"
                        .Cells(7, 3 + (5 + (4 * (a - 1)))).Value = "+/-"
                        
                        If a <= TotalMonths Then
                            .Cells(8, (5 + (4 * (a - 1)))).Value = SaMData(a).TotSales.OriginalForecast
                            .Cells(8, 1 + (5 + (4 * (a - 1)))).Value = SaMData(a).TotSales.ReForecast
                            .Cells(8, 2 + (5 + (4 * (a - 1)))).Value = SaMData(a).TotSales.Actual
                            
                            .Cells(9, (5 + (4 * (a - 1)))).Value = g_DivZero((SaMContData(a).TotSales.OriginalForecast), .Cells(8, (5 + (4 * (a - 1)))).Value) * 100
                            .Cells(9, 1 + (5 + (4 * (a - 1)))).Value = g_DivZero((SaMContData(a).TotSales.ReForecast), .Cells(8, 1 + (5 + (4 * (a - 1)))).Value) * 100
                            .Cells(9, 2 + (5 + (4 * (a - 1)))).Value = g_DivZero((SaMContData(a).TotSales.Actual), .Cells(8, 2 + (5 + (4 * (a - 1)))).Value) * 100
                            
                            .Cells(14, (5 + (4 * (a - 1)))).Value = SaMData(a).CRSales.OriginalForecast
                            .Cells(14, 1 + (5 + (4 * (a - 1)))).Value = SaMData(a).CRSales.ReForecast
                            .Cells(14, 2 + (5 + (4 * (a - 1)))).Value = SaMData(a).CRSales.Actual
                            .Cells(16, (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).CRSales.OriginalForecast, SaMData(a).CRSales.OriginalForecast) * 100
                            .Cells(16, 1 + (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).CRSales.ReForecast, SaMData(a).CRSales.ReForecast) * 100
                            .Cells(16, 2 + (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).CRSales.Actual, SaMData(a).CRSales.Actual) * 100
                            .Cells(17, (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).CRSales.OriginalForecast, SaMData(a).TotSales.OriginalForecast) * 100
                            .Cells(17, 1 + (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).CRSales.ReForecast, SaMData(a).TotSales.ReForecast) * 100
                            .Cells(17, 2 + (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).CRSales.Actual, SaMData(a).TotSales.Actual) * 100
                            
                            .Cells(18, (5 + (4 * (a - 1)))).Value = SaMData(a).FVSales.OriginalForecast
                            .Cells(18, 1 + (5 + (4 * (a - 1)))).Value = SaMData(a).FVSales.ReForecast
                            .Cells(18, 2 + (5 + (4 * (a - 1)))).Value = SaMData(a).FVSales.Actual
                            .Cells(20, (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).FVSales.OriginalForecast, SaMData(a).FVSales.OriginalForecast) * 100
                            .Cells(20, 1 + (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).FVSales.ReForecast, SaMData(a).FVSales.ReForecast) * 100
                            .Cells(20, 2 + (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).FVSales.Actual, SaMData(a).FVSales.Actual) * 100
                            .Cells(21, (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).FVSales.OriginalForecast, SaMData(a).TotSales.OriginalForecast) * 100
                            .Cells(21, 1 + (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).FVSales.ReForecast, SaMData(a).TotSales.ReForecast) * 100
                            .Cells(21, 2 + (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).FVSales.Actual, SaMData(a).TotSales.Actual) * 100
                            
                            .Cells(22, (5 + (4 * (a - 1)))).Value = SaMData(a).ChilSales.OriginalForecast
                            .Cells(22, 1 + (5 + (4 * (a - 1)))).Value = SaMData(a).ChilSales.ReForecast
                            .Cells(22, 2 + (5 + (4 * (a - 1)))).Value = SaMData(a).ChilSales.Actual
                            .Cells(24, (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).ChilSales.OriginalForecast, SaMData(a).ChilSales.OriginalForecast) * 100
                            .Cells(24, 1 + (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).ChilSales.ReForecast, SaMData(a).ChilSales.ReForecast) * 100
                            .Cells(24, 2 + (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).ChilSales.Actual, SaMData(a).ChilSales.Actual) * 100
                            .Cells(25, (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).ChilSales.OriginalForecast, SaMData(a).TotSales.OriginalForecast) * 100
                            .Cells(25, 1 + (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).ChilSales.ReForecast, SaMData(a).TotSales.ReForecast) * 100
                            .Cells(25, 2 + (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).ChilSales.Actual, SaMData(a).TotSales.Actual) * 100
                            
                            .Cells(26, (5 + (4 * (a - 1)))).Value = SaMData(a).MeatSales.OriginalForecast
                            .Cells(26, 1 + (5 + (4 * (a - 1)))).Value = SaMData(a).MeatSales.ReForecast
                            .Cells(26, 2 + (5 + (4 * (a - 1)))).Value = SaMData(a).MeatSales.Actual
                            .Cells(28, (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).MeatSales.OriginalForecast, SaMData(a).MeatSales.OriginalForecast) * 100
                            .Cells(28, 1 + (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).MeatSales.ReForecast, SaMData(a).MeatSales.ReForecast) * 100
                            .Cells(28, 2 + (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).MeatSales.Actual, SaMData(a).MeatSales.Actual) * 100
                            .Cells(29, (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).MeatSales.OriginalForecast, SaMData(a).TotSales.OriginalForecast) * 100
                            .Cells(29, 1 + (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).MeatSales.ReForecast, SaMData(a).TotSales.ReForecast) * 100
                            .Cells(29, 2 + (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).MeatSales.Actual, SaMData(a).TotSales.Actual) * 100
                            
                            .Cells(30, (5 + (4 * (a - 1)))).Value = SaMData(a).SpecialSales.OriginalForecast
                            .Cells(30, 1 + (5 + (4 * (a - 1)))).Value = SaMData(a).SpecialSales.ReForecast
                            .Cells(30, 2 + (5 + (4 * (a - 1)))).Value = SaMData(a).SpecialSales.Actual
                            .Cells(32, (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).SpecialSales.OriginalForecast, SaMData(a).SpecialSales.OriginalForecast) * 100
                            .Cells(32, 1 + (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).SpecialSales.ReForecast, SaMData(a).SpecialSales.ReForecast) * 100
                            .Cells(32, 2 + (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).SpecialSales.Actual, SaMData(a).SpecialSales.Actual) * 100
                            .Cells(33, (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).SpecialSales.OriginalForecast, SaMData(a).TotSales.OriginalForecast) * 100
                            .Cells(33, 1 + (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).SpecialSales.ReForecast, SaMData(a).TotSales.ReForecast) * 100
                            .Cells(33, 2 + (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).SpecialSales.Actual, SaMData(a).TotSales.Actual) * 100
                            
                            .Cells(34, (5 + (4 * (a - 1)))).Value = SaMData(a).SeasonalSales.OriginalForecast
                            .Cells(34, 1 + (5 + (4 * (a - 1)))).Value = SaMData(a).SeasonalSales.ReForecast
                            .Cells(34, 2 + (5 + (4 * (a - 1)))).Value = SaMData(a).SeasonalSales.Actual
                            .Cells(36, (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).SeasonalSales.OriginalForecast, SaMData(a).SeasonalSales.OriginalForecast) * 100
                            .Cells(36, 1 + (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).SeasonalSales.ReForecast, SaMData(a).SeasonalSales.ReForecast) * 100
                            .Cells(36, 2 + (5 + (4 * (a - 1)))).Value = g_DivZero(SaMContData(a).SeasonalSales.Actual, SaMData(a).SeasonalSales.Actual) * 100
                            .Cells(37, (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).SeasonalSales.OriginalForecast, SaMData(a).TotSales.OriginalForecast) * 100
                            .Cells(37, 1 + (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).SeasonalSales.ReForecast, SaMData(a).TotSales.ReForecast) * 100
                            .Cells(37, 2 + (5 + (4 * (a - 1)))).Value = CBA_BasicFunctions.g_DivZero(SaMContData(a).SeasonalSales.Actual, SaMData(a).TotSales.Actual) * 100
                        End If
                        
                        
                        For b = 8 To 37
'                            If b = 18 Then
'                            a = a
'                            End If
                            If a = TotalMonths + 1 Then
                                For c = 0 To 2
                                    .Cells(b, c + (5 + (4 * (a - 1)))).Value = strFormu
                                Next
                            End If
                            
                            If .Cells(b, 1 + (5 + (4 * (a - 1)))).Value > 0 Then
                                .Cells(b, 3 + (5 + (4 * (a - 1)))).Value = .Cells(b, 2 + (5 + (4 * (a - 1)))).Value - .Cells(b, 1 + (5 + (4 * (a - 1)))).Value
                            Else
                                .Cells(b, 3 + (5 + (4 * (a - 1)))).Value = .Cells(b, 2 + (5 + (4 * (a - 1)))).Value - .Cells(b, (5 + (4 * (a - 1)))).Value
                            End If
                            If .Cells(b, 2 + (5 + (4 * (a - 1)))).Value = 0 Then .Cells(b, 3 + (5 + (4 * (a - 1)))).Value = "-"
                        Next
                        'ADDING SHARE
                        For c = 8 To 37
                            For b = 0 To 3
                                If b = 3 And .Cells(c, b + (5 + (4 * (a - 1)))).Value < 0 Then .Cells(c, b + (5 + (4 * (a - 1)))).Font.ColorIndex = 3
                                Select Case c
                                
                                    Case 8, 10, 14, 18, 22, 26, 30, 34
                                        'Sales
                                        If a < TotalMonths + 1 Then
                                            If c = 10 Then
                                                .Cells(c, b + (5 + (4 * (a - 1)))).Value = "=IF(SUM(R[4]C,R[8]C,R[12]C,R[16]C)<>0,SUM(R[4]C,R[8]C,R[12]C,R[16]C),"" - "")"
                                            Else
                                                If IsNumeric(.Cells(c, b + (5 + (4 * (a - 1)))).Value) Then .Cells(c, b + (5 + (4 * (a - 1)))).Value = .Cells(c, b + (5 + (4 * (a - 1)))).Value / 1000000
                                            End If
                                        End If
                                        .Cells(c, b + (5 + (4 * (a - 1)))).NumberFormat = "#0.00"
                                    Case 9, 12, 16, 20, 24, 28, 32, 36
                                        If a < TotalMonths + 1 Then
                                            If c = 12 Then .Cells(c, b + (5 + (4 * (a - 1)))).Value = "=IF(IFERROR(R[1]C/R[-2]C,0)*100=0,""-"",IFERROR(R[1]C/R[-2]C,0)*100)"
                                        End If
                                        'Margin
                                        .Cells(c, b + (5 + (4 * (a - 1)))).NumberFormat = "#0.00"
                                    Case 11, 15, 19, 23, 27, 31, 35
                                        'Share
                                        If a < TotalMonths + 1 Then
                                            .Cells(c, b + (5 + (4 * (a - 1)))).Value = "=IF(IFERROR((R[-1]C/R8C)*100,0)=0,""-"",IFERROR((R[-1]C/R8C)*100,0))"
                                        End If
                                        .Cells(c, b + (5 + (4 * (a - 1)))).NumberFormat = "#0.00"
                                    Case 13, 17, 21, 25, 29, 33, 37
                                        'Contribution
                                        If a < TotalMonths + 1 Then
                                            If c = 13 Then
                                                .Cells(c, b + (5 + (4 * (a - 1)))).Value = "=IF(SUM(R[4]C,R[8]C,R[12]C,R[16]C)<>0,SUM(R[4]C,R[8]C,R[12]C,R[16]C),""-"")"
'                                            Else
'                                                If IsNumeric(.Cells(c, b + (5 + (4 * (a - 1)))).Value) Then .Cells(c, b + (5 + (4 * (a - 1)))).Value = .Cells(c, b + (5 + (4 * (a - 1)))).Value / 1000000
                                            End If
                                        End If
                                        .Cells(c, b + (5 + (4 * (a - 1)))).NumberFormat = "#0.00"
                                End Select
                                If .Cells(c, b + (5 + (4 * (a - 1)))).Value = 0 Then .Cells(c, b + (5 + (4 * (a - 1)))).Value = "-"
                            Next
                        Next
                        Range(.Cells(7, (5 + (4 * (a - 1)))), .Cells(37, (5 + (4 * (a - 1))))).Interior.ColorIndex = 20
                        Range(.Cells(7, 1 + (5 + (4 * (a - 1)))), .Cells(37, 1 + (5 + (4 * (a - 1))))).Interior.ColorIndex = 40
                    Next
                Range(.Cells(1, (5 + (4 * ((TotalMonths + 1) - 1)))), .Cells(1, 3 + (5 + (4 * ((TotalMonths + 1) - 1))))).ColumnWidth = 11.8
                Range(.Cells(1, 1), .Cells(1, 2)).EntireColumn.Delete
                End With
                
                Application.ScreenUpdating = True
            End If
    
    Case "Product Level SCG Dynamic Forecasting Report"
        
            
        'Dim endm As Byte, MT As Byte, startm As Byte
        'Dim strFormu As String
            
            If RP.PSYear = 0 Then MsgBox "No Year Start has been selected": Exit Sub
            If RP.PEYear = 0 Then MsgBox "No Year End has been selected": Exit Sub
            If RP.PSMonth = 0 Then MsgBox "No Month Start has been selected": Exit Sub
            If RP.PEMonth = 0 Then MsgBox "No Month End has been selected": Exit Sub

            
           
            YearFrom = RP.PSYear
            YearTo = RP.PEYear
            MonthFrom = RP.PSMonth
            MonthTo = RP.PEMonth
            
            TotalMonths = DateDiff("M", DateSerial(RP.PSYear, RP.PSMonth, 1), DateSerial(RP.PEYear, RP.PEMonth + 1, 0)) + 1
            'ReDim SaMData(1 To totalmonths)
            'ReDim SaMContData(1 To totalmonths)
            
            
            CreateCGsArray RP
            Unload CBA_BTF_frm_Reporting
            'CBA_BasicFunctions.CBA_Running "DBU Forecast Report"
            
            'arrDBU = CBA_SQL_Queries.CBA_GenPullSQL("DBU_Map")
            
            If CBA_BTF_SetupForecastArray.CBA_BTF_SetupForecastArray(YearFrom, YearTo, MonthFrom, MonthTo) = True Then
                Application.ScreenUpdating = False
                Set wbk = Workbooks.Add
                Set wks(1) = ActiveSheet
                With wks(1)
     
                End With
           
            End If
      
End Select


exitTheReport:
If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
Erase CBA_SCGData

Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-RunReport", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Private Function BuildSummaryData(ByRef wks As Worksheet, ByVal PClass As Long, ByVal CG As Long, ByVal RowNum As Long)
Dim RCell As Range
    For Each RCell In wks.Rows(RowNum).Cells
        If RCell.Column > 1 Then
            If RCell.Column = 33 Then Exit For
            SummaryData(CG, PClass, RCell.Column) = RCell.Value
        End If
    Next
End Function

Private Function SetCGTotal(ByRef wks As Worksheet, ByRef StartRow, ByVal CG As Long, ByVal FLevel As String, ByRef lRowNo, ByVal PClass As Long, Optional isfinalGT As Boolean)
    Dim a As Long, b As Long, Pc As Long
    Dim d As Long, e As Long
    Dim c As Long
    Dim strSQL As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    With wks
        If isfinalGT = False Then
            .Cells(lRowNo, 1).Value = "CG Total: " & CG
    
        For a = 2 To 27
            Select Case a
                Case 2, 5, 8, 11, 14, 17, 20, 23, 26
                    If .Cells(lRowNo, a + 1).Value = "" And .Cells(lRowNo, a).Value <> "" Then .Cells(lRowNo, a + 1).Value = .Cells(lRowNo, a).Value
            End Select
        Next
    
        For a = 10 To lRowNo
            .Cells(a, 1).EntireRow.HorizontalAlignment = xlCenter
        Next
        
        For a = 33 To 38
            .Cells(lRowNo, a).Value = Application.WorksheetFunction.Sum(Range(.Cells(StartRow, a), .Cells(lRowNo - 1, a)))
        Next
        
        For a = 1 To 32
            Select Case a
                Case 1
                    Range(.Cells(12, a), .Cells(lRowNo, a)).NumberFormat = "General"
                Case 2 To 7, 11, 14, 17 To 22, 29, 30, 33 To 36
                    Range(.Cells(12, a), .Cells(lRowNo, a)).NumberFormat = "$#,0"
                    .Cells(lRowNo, a).Value = Application.WorksheetFunction.Sum(Range(.Cells(StartRow, a), .Cells(lRowNo - 1, a)))
                Case 8 To 10, 17, 18, 23 To 28, 31, 32
                    Range(.Cells(12, a), .Cells(lRowNo, a)).NumberFormat = "#,0.0%"
                Case 12
                    Set CBA_COM_SKU_CBISRS = New ADODB.Recordset
                    strSQL = "select (sum(retailnet) - sum(cost) ) / sum(retail) , sum(posretail) from #OP" & Chr(10)
                    strSQL = strSQL & "where YearNo = " & wks.Cells(8, 11).MergeArea(1, 1).Value & Chr(10)
                    strSQL = strSQL & "and cgno = " & CG & Chr(10)
                    strSQL = strSQL & "and productclass = " & PClass & Chr(10)
                    CBA_COM_SKU_CBISRS.Open strSQL, CBA_COM_SKU_CBISCN
                    Range(.Cells(12, a), .Cells(lRowNo, a)).NumberFormat = "$#,0"
                    .Cells(lRowNo, a).Value = (CBA_COM_SKU_CBISRS(0) * CBA_COM_SKU_CBISRS(1))
                Case 13
                    Range(.Cells(12, a), .Cells(lRowNo, a)).NumberFormat = "$#,0"
                    .Cells(lRowNo, a).Value = CBA_COM_SKU_CBISRS(0)
                Case 15
                    Set CBA_COM_SKU_CBISRS = New ADODB.Recordset
                    strSQL = "select (sum(retailnet) - sum(cost) ) / sum(retail) , sum(posretail) from #OP" & Chr(10)
                    strSQL = strSQL & "where YearNo = " & wks.Cells(8, 14).MergeArea(1, 1).Value & Chr(10)
                    strSQL = strSQL & "and cgno = " & CG & Chr(10)
                    strSQL = strSQL & "and productclass = " & PClass & Chr(10)
                    CBA_COM_SKU_CBISRS.Open strSQL, CBA_COM_SKU_CBISCN
                    Range(.Cells(12, a), .Cells(lRowNo, a)).NumberFormat = "$#,0"
                    .Cells(lRowNo, a).Value = (CBA_COM_SKU_CBISRS(0) * CBA_COM_SKU_CBISRS(1))
                Case 16
                    Range(.Cells(12, a), .Cells(lRowNo, a)).NumberFormat = "$#,0"
                    .Cells(lRowNo, a).Value = CBA_COM_SKU_CBISRS(0)
    
    '                                .Cells(lRowNo, a).Value = Application.WorksheetFunction.Average(Range(.Cells(12, a), .Cells(lRowNo - 1, a)))
            End Select
            Select Case a
                Case 8, 9
                    If .Cells(lRowNo, a - 6).Value = 0 Then .Cells(lRowNo, a).Value = 0 Else .Cells(lRowNo, a).Value = .Cells(lRowNo, a - 3).Value / .Cells(lRowNo, a - 6).Value
                Case 13, 16
                    If .Cells(lRowNo, a - 2).Value = 0 Then .Cells(lRowNo, a).Value = 0 Else .Cells(lRowNo, a).Value = .Cells(lRowNo, a - 1).Value / .Cells(lRowNo, a - 2).Value
                Case 17, 18
    '                            If .Cells(lRowNo, b).Value = 0 Then .Cells(lRowNo, b + 6).Value = 0 Else .Cells(lRowNo, b + 6).Value = (.Cells(lRowNo, b - 15).Value - .Cells(lRowNo, b).Value) / .Cells(lRowNo, b).Value
    '                                .Cells(lRowNo, b).Value = .Cells(lRowNo, b - 15).Value - .Cells(lRowNo, b).Value
    '                                If .Cells(lRowNo, b).Value = 0 Then .Cells(lRowNo, b + 6).Value = 0 Else .Cells(lRowNo, b + 6).Value = (.Cells(lRowNo, b - 15).Value - .Cells(lRowNo, b).Value) / .Cells(lRowNo, b).Value
            End Select
            Range(.Cells(12, a), .Cells(lRowNo, a)).Borders.Weight = xlThin
            'Range(.Cells(lRowNo, a), .Cells(lRowNo, a)).Borders(xlEdgeTop).Weight = xlThick
        Next
        For b = 17 To 22
            If .Cells(lRowNo, b + 16).Value = 0 Then .Cells(lRowNo, b + 6).Value = 0 Else .Cells(lRowNo, b + 6).Value = (.Cells(lRowNo, b - 15).Value - .Cells(lRowNo, b + 16).Value) / .Cells(lRowNo, b + 16).Value
            '.Cells(lRowNo, b).Value = .Cells(lRowNo, b - 15).Value - .Cells(lRowNo, b).Value
        Next
        If .Cells(lRowNo, 14).Value = 0 Then .Cells(lRowNo, 31).Value = 0 Else .Cells(lRowNo, 31).Value = (.Cells(lRowNo, 11).Value - .Cells(lRowNo, 14).Value) / .Cells(lRowNo, 14).Value
        If .Cells(lRowNo, 15).Value = 0 Then .Cells(lRowNo, 32).Value = 0 Else .Cells(lRowNo, 32).Value = (.Cells(lRowNo, 12).Value - .Cells(lRowNo, 15).Value) / .Cells(lRowNo, 15).Value
        
        If RCVRet = 0 Then
            .Cells(lRowNo, 10).Value = 0
        Else
            .Cells(lRowNo, 10).Value = (TotRCVRetNet - TotRCVCost) / TotRCVRet
        End If
        Rows(StartRow & ":" & lRowNo - 1).Group
        Range(wks.Cells(StartRow, 1), wks.Cells(lRowNo - 1, 1)).RowHeight = 0
        BuildSummaryData wks, PClass, CG, lRowNo
        lRowNo = lRowNo + 1
        
        
        Else
            lRowNo = lRowNo + 2
            .Cells(lRowNo, 1).Value = "Grand Total"
            Range(.Cells(lRowNo, 1), .Cells(lRowNo, 32)).Borders.Weight = xlThin
            Range(.Cells(lRowNo, 1), .Cells(lRowNo, 32)).Borders(xlEdgeTop).Weight = xlThick
            For b = 2 To 32
                    d = 0: e = 0
                For c = LBound(SummaryData, 1) To UBound(SummaryData, 1)
                    Select Case b
                        Case 1
                            Range(.Cells(12, a), .Cells(lRowNo, a)).NumberFormat = "General"
                        Case 2 To 7, 11, 14, 17 To 22, 29, 30
                            For Pc = 1 To 4
                                If PClass = 5 Or Pc = PClass Then .Cells(lRowNo, b).Value = .Cells(lRowNo, b).Value + SummaryData(c, Pc, b)
                            Next
                            Range(.Cells(StartRow, b), .Cells(lRowNo, b)).NumberFormat = "$#,0"
                        Case 8, 9, 10
                            d = 6: e = 3
                            Range(.Cells(StartRow, b), .Cells(lRowNo, b)).NumberFormat = "#,0.0%"
    '                    Case 13, 16
    '                        d = 2: e = 1
    '                        Range(.Cells(startRow, b), .Cells(lRowNo, b)).NumberFormat = "#,0.0%"
                        Case 23 To 28
                            d = 21: e = 6
                            Range(.Cells(StartRow, b), .Cells(lRowNo, b)).NumberFormat = "#,0.0%"
                        Case 31, 32
                            d = 17: e = 2
                            Range(.Cells(StartRow, b), .Cells(lRowNo, b)).NumberFormat = "#,0.0%"
                        Case 12
                            Set CBA_COM_SKU_CBISRS = New ADODB.Recordset
                            strSQL = "select (sum(retailnet) - sum(cost) ) / sum(retail) , sum(posretail) from #OP" & Chr(10)
                            strSQL = strSQL & "where YearNo = " & wks.Cells(8, 11).MergeArea(1, 1).Value & Chr(10)
                            'strSQL = strSQL & "and cgno = " & CG & Chr(10)
                            If PClass = 5 Or PClass = 0 Then Else strSQL = strSQL & "and productclass = " & PClass & Chr(10)
                            CBA_COM_SKU_CBISRS.Open strSQL, CBA_COM_SKU_CBISCN
                            Range(.Cells(12, b), .Cells(lRowNo, b)).NumberFormat = "$#,0"
                            .Cells(lRowNo, b).Value = (CBA_COM_SKU_CBISRS(0) * CBA_COM_SKU_CBISRS(1))
                        Case 13
                            Range(.Cells(StartRow, b), .Cells(lRowNo, b)).NumberFormat = "#,0.0%"
                            .Cells(lRowNo, b).Value = CBA_COM_SKU_CBISRS(0)
                        Case 15
                            Set CBA_COM_SKU_CBISRS = New ADODB.Recordset
                            strSQL = "select (sum(retailnet) - sum(cost) ) / sum(retail) , sum(posretail) from #OP" & Chr(10)
                            strSQL = strSQL & "where YearNo = " & wks.Cells(8, 14).MergeArea(1, 1).Value & Chr(10)
                            'strSQL = strSQL & "and cgno = " & CG & Chr(10)
                            If PClass = 5 Or PClass = 0 Then Else strSQL = strSQL & "and productclass = " & PClass & Chr(10)
                            CBA_COM_SKU_CBISRS.Open strSQL, CBA_COM_SKU_CBISCN
                            Range(.Cells(StartRow, b), .Cells(lRowNo, b)).NumberFormat = "$#,0"
                            .Cells(lRowNo, b).Value = (CBA_COM_SKU_CBISRS(0) * CBA_COM_SKU_CBISRS(1))
                        Case 16
                            Range(.Cells(StartRow, b), .Cells(lRowNo, b)).NumberFormat = "#,0.0%"
                            .Cells(lRowNo, b).Value = CBA_COM_SKU_CBISRS(0)
                            
                    End Select
                    If d > 0 Then
                        If wks.Cells(lRowNo, b - d).Value = 0 Then
                            wks.Cells(lRowNo, b).Value = 0
                        Else
                            wks.Cells(lRowNo, b).Value = wks.Cells(lRowNo, b - e).Value / wks.Cells(lRowNo, b - d).Value
                        End If
                    End If
                Next
            Next
        End If
        
    End With
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-setCGTotal", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function


Private Function putDataOnSheet(ByRef wks As Worksheet, ByVal PClass As Long, ByVal Mnth As Long, ByVal reportYear As Long, ByVal FLevel As String, ByVal CG As Long, ByVal RowNum As Long, Optional ByVal scg As Long)
    Dim CBA_Proc As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    With wks
    
        'If FLevel = "CG" Then SCG = 0
        If Mnth >= Month(Date) Or reportYear > Year(Date) Then
        
                                        
                    If .Cells(11, 4).Value = "Actual" Then .Cells(11, 4).Value = "Actual + Forecast"
                    If FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.ReForecast = 0 Then .Cells(RowNum, 4).Value = .Cells(RowNum, 4).Value + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.OriginalForecast Else .Cells(RowNum, 4).Value = .Cells(RowNum, 4).Value + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.ReForecast
                    
                    If .Cells(11, 7).Value = "Actual" Then .Cells(11, 7).Value = "Actual + Forecast"
                    If FCbM(reportYear, Mnth, PClass).CGData(CG, scg).MarginDol.ReForecast = 0 Then .Cells(RowNum, 7).Value = .Cells(RowNum, 7).Value + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).MarginDol.OriginalForecast Else .Cells(RowNum, 7).Value = .Cells(RowNum, 7).Value + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).MarginDol.ReForecast
                    
                    If .Cells(11, 19).Value = "Actual" Then .Cells(11, 19).Value = "Actual + Forecast"
                    If FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).Sales.ReForecast = 0 Then .Cells(RowNum, 35).Value = .Cells(RowNum, 35).Value + FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).Sales.OriginalForecast Else .Cells(RowNum, 35).Value = .Cells(RowNum, 35).Value + FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).Sales.ReForecast
                    
                    
                    If .Cells(11, 22).Value = "Actual" Then .Cells(11, 22).Value = "Actual + Forecast"
                    If FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).MarginDol.ReForecast = 0 Then .Cells(RowNum, 38).Value = .Cells(RowNum, 38).Value + FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).MarginDol.OriginalForecast Else .Cells(RowNum, 38).Value = .Cells(RowNum, 38).Value + FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).MarginDol.ReForecast
                    
                    
                    If .Cells(11, 10).Value = "Actual" Then .Cells(11, 10).Value = "Actual + Forecast"
                    If FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.ReForecast = 0 Then
                        RCVRetNet = RCVRetNet + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.OriginalForecast
                        RCVRet = RCVRet + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.OriginalForecast
                        RCVCost = RCVCost + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.OriginalForecast - FCbM(reportYear, Mnth, PClass).CGData(CG, scg).MarginDol.OriginalForecast
                    Else
                        RCVRetNet = RCVRetNet + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.ReForecast
                        RCVRet = RCVRet + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.ReForecast
                        RCVCost = RCVCost + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.ReForecast - FCbM(reportYear, Mnth, PClass).CGData(CG, scg).MarginDol.ReForecast
                    End If
                    GoTo ForecastRatherThanActuals
            
        End If
        
        .Cells(RowNum, 4).Value = .Cells(RowNum, 4).Value + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.Actual
        .Cells(RowNum, 7).Value = .Cells(RowNum, 7).Value + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).MarginDol.Actual
        RCVRetNet = RCVRetNet + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).RCVRetailNet
        RCVRet = RCVRet + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).RCVRetail
        RCVCost = RCVCost + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Cost.Actual
ForecastRatherThanActuals:
        .Cells(RowNum, 35).Value = .Cells(RowNum, 35).Value + FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).Sales.Actual
        .Cells(RowNum, 38).Value = .Cells(RowNum, 38).Value + FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).MarginDol.Actual
        .Cells(RowNum, 2).Value = .Cells(RowNum, 2).Value + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.OriginalForecast
        .Cells(RowNum, 3).Value = .Cells(RowNum, 3).Value + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).Sales.ReForecast
        .Cells(RowNum, 5).Value = .Cells(RowNum, 5).Value + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).MarginDol.OriginalForecast
        .Cells(RowNum, 6).Value = .Cells(RowNum, 6).Value + FCbM(reportYear, Mnth, PClass).CGData(CG, scg).MarginDol.ReForecast
        .Cells(RowNum, 33).Value = .Cells(RowNum, 33).Value + FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).Sales.OriginalForecast
        .Cells(RowNum, 34).Value = .Cells(RowNum, 34).Value + FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).Sales.ReForecast
        .Cells(RowNum, 36).Value = .Cells(RowNum, 36).Value + FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).MarginDol.OriginalForecast
        .Cells(RowNum, 37).Value = .Cells(RowNum, 37).Value + FCbM(reportYear - 1, Mnth, PClass).CGData(CG, scg).MarginDol.ReForecast
        
        'If RCVRet = 0 Then putDataOnSheet = 0 Else putDataOnSheet = (RCVRetNet - RCVCost) / RCVRet
        
    
    End With
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-putDataOnSheet", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function pullBaseData()
        
        CBA_SQL_Queries.CBA_GenPullSQL "CBForecast_RepFMarginP"
        FMarginP = CBA_CBFCarr
        CBA_SQL_Queries.CBA_GenPullSQL "CBForecast_RepFReMarginP"
        FReMarginP = CBA_CBFCarr
        CBA_SQL_Queries.CBA_GenPullSQL "CBForecast_RepFReSales"
        FReSales = CBA_CBFCarr
        CBA_SQL_Queries.CBA_GenPullSQL "CBForecast_RepFSales"
        FSales = CBA_CBFCarr
        basedatapulled = True
End Function

Public Function setCGs(ByRef CGarr As Variant)
    CGs = CGarr
End Function

Private Function CreateCGsArray(ByRef RP As CBA_BTF_ReportParamaters)
   Dim a As Long, GBCG As Long, CGno As Long
    Dim yn As Long, CGarrs, GBDCGarrs
    Dim CGDic As Scripting.Dictionary
    
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    If RP.GBD <> "" Then
        If RP.GBD = "ALL" Then RP.GBD = ""
         CBA_COM_SQLQueries.CBA_COM_GenPullSQL "ALLCGCntSCG"
         CGarrs = CBA_CBISarr
         Erase CBA_CBISarr
         CBA_COM_SQLQueries.CBA_COM_GenPullSQL "CBIS_CGsbyGBDnameCount", , , , , , RP.GBD
         GBDCGarrs = CBA_CBISarr
         Erase CBA_CBISarr
         CBA_COM_SQLQueries.CBA_COM_GenPullSQL "CBIS_CGsbyGBDname", , , , , , RP.GBD
         If UBound(CBA_CBISarr, 2) = 0 Then
             If CBA_CBISarr(0, 0) = 0 Then
                 MsgBox "There is no mappings to any CGs for this BD"
                 Exit Function
             End If
         End If
         Set CGDic = New Scripting.Dictionary
         For CGno = LBound(CGarrs, 2) To UBound(CGarrs, 2)
            CGDic.Add CInt(CGarrs(0, CGno)), CInt(CGarrs(1, CGno))
         Next
         For GBCG = LBound(GBDCGarrs, 2) To UBound(GBDCGarrs, 2)
            If CInt(GBDCGarrs(1, GBCG)) = CGDic(CInt(GBDCGarrs(0, GBCG))) Then
                For CGno = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                    If CBA_CBISarr(0, CGno) = GBDCGarrs(0, GBCG) Then
                       CBA_CBISarr(1, CGno) = "0"
                    End If
                    If UBound(CBA_CBISarr, 2) = CGno Then Exit For
                    If CBA_CBISarr(0, CGno + 1) > GBDCGarrs(0, GBCG) Then Exit For
                Next
            End If
         Next
         CGs = CBA_CBISarr
         Erase CBA_CBISarr
         Erase GBDCGarrs
         Erase CGarrs
     ElseIf RP.BD <> "" Then
     ReDim BDarr(0 To 12, 0 To 0)
         CBA_COM_SQLQueries.CBA_COM_GenPullSQL "CBIS_CGsbyEmpname", , , , , , RP.BD
         If UBound(CBA_CBISarr, 2) = 0 Then
             If CBA_CBISarr(0, 0) = 0 Then
                 MsgBox "There is no mappings to any CGs for this BD"
                 Exit Function
             End If
         End If
         CGs = CBA_CBISarr
         Erase CBA_CBISarr
     ElseIf RP.CG <> 0 Then
         If RP.scg <> 0 Then
             ReDim CGs(0 To 1, 0 To 0)
             CGs(0, 0) = RP.CG
             CGs(1, 0) = RP.scg
         ElseIf CBA_BTF_frm_Reporting.cbx_SCG.ListCount = 0 Then
             ReDim CGs(0 To 1, 0 To 0)
             CGs(0, 0) = RP.CG
             CGs(1, 0) = ""
         Else
             ReDim CGs(0 To 1, 0 To CBA_BTF_frm_Reporting.cbx_SCG.ListCount - 1)
             For a = 0 To CBA_BTF_frm_Reporting.cbx_SCG.ListCount - 1
                 CGs(0, a) = RP.CG
                 CGs(1, a) = Trim(Mid(CBA_BTF_frm_Reporting.cbx_SCG.List(a), 1, 2))
             Next
         End If
     Else
          yn = MsgBox("You are about to query for all CG's. Are you sure?", vbYesNo)
          If yn = 6 Then
            RP.GBD = "ALL"
            Erase CGs
          End If
     End If
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CreateCGsArray", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Public Sub CrtForecastingWorkbook(sFormCalled As String)
    ''Dim wbk_Tmp As Workbook
    Dim wks_Tmp As Worksheet, a As Long, sWorksheetName As String
    On Error GoTo Err_Routine
    CBA_ErrTag = "": sWorksheetName = "FCWorksheet"
''    Set wbk_Tmp = Application.Workbooks.Add
    Application.DisplayAlerts = False
    Application.ScreenUpdating = True
    Set wks_Tmp = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ' Delete any worksheet that exists with the same name
    Call g_WorkSheetDelete(ActiveWorkbook, sWorksheetName)
    ' Rename the worksheet
    wks_Tmp.Name = sWorksheetName
    ' Activate the new worksheet
    ActiveWorkbook.Worksheets(sWorksheetName).Select
    ActiveWorkbook.Worksheets(sWorksheetName).Activate
''    ThisWorkbook.Worksheets(sWorksheetName).Hidden = False
    ''Set wks_Tmp = ActiveSheet
    Application.DisplayAlerts = True
    With wks_Tmp
        Range(.Cells(1, 1), .Cells(5, 79)).Interior.ColorIndex = 49
        .Cells(1, 1).Select
        .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
        
        Cells.Font.Name = "ALDI SUED Office"
        .Cells(2, 2).Font.Size = 22
        .Cells(2, 2).Font.ColorIndex = 2
        
        If sFormCalled = "optProduct" Then .Cells(7, 3).Value = "Product Code"
        
        .Cells(7, 4).Value = "Year"
        .Cells(7, 5).Value = "Month"
        .Cells(7, 6).Value = "POS Quantity"
                            '' .Cells(7, 7).Value = "Prior Year POS Quantity"
        .Cells(7, 7).Value = "POS Quantity Year Growth"         '.Cells(7, 8).Value = "POS Quantity Year Growth"
                            ''.Cells(7, 9).Value = "POS Month Growth"
                            ''.Cells(7, 10).Value = "Prior Year POS Month Growth"
        .Cells(7, 8).Value = "POS Retail"                      '.Cells(7, 11).Value = "POS Retail"
        .Cells(7, 9).Value = "POS Retail Year Growth"          '.Cells(7, 12).Value = "POS Retail Year Growth"
        .Cells(7, 10).Value = "Margin% (RCV)"                   '.Cells(7, 13).Value = "Margin% (RCV)"
                                    ''.Cells(7, 14).Value = "Cost Per Unit"
                                    ''.Cells(7, 15).Value = "Cost Per Unit Year Growth"
        Range(.Cells(7, 3), .Cells(7, 10)).Font.Bold = True
        Range(.Cells(7, 3), .Cells(7, 10)).HorizontalAlignment = xlCenter
        Range(.Cells(7, 3), .Cells(7, 10)).VerticalAlignment = xlBottom
        
        If sFormCalled = "optProduct" Then Range(.Cells(7, 3), .Cells(7, 10)).Borders.Weight = 3 Else Range(.Cells(7, 4), .Cells(7, 10)).Borders.Weight = 3
        Range(.Cells(7, 3), .Cells(7, 10)).WrapText = True
        For a = 1 To 12
            .Cells(2, 9 + a).Value = DateAdd("M", a, DateSerial(Year(Date), Month(Date), 1))
            .Cells(2, 9 + a).NumberFormat = "MMM-YY"
        Next
        
        If sFormCalled = "optCG" Then .Cells(2, 22).Value = "CG TOTAL" Else .Cells(2, 22).Value = "SCG TOTAL"
        .Cells(2, 9).Value = "Class TOTAL"
        
        Range(.Cells(3, 9), .Cells(4, 22)).Interior.ColorIndex = False
        Range(.Cells(3, 9), .Cells(4, 22)).Borders.Weight = xlThin
        
'            .Cells(1, 17).Value = "Product Class"
        .Cells(3, 8).Value = "RCV Margin:"
        .Cells(4, 8).Value = "Retail:"
'            With .Cells(2, 17).Validation
'                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
'                    Formula1:="ALL,Core Range,Food Special, Non-Food Special, Seasonal"
'            End With
'            .Cells(2, 17).Interior.ColorIndex = 36
        Range(.Cells(2, 9), .Cells(2, 22)).HorizontalAlignment = xlCenter
        Range(.Cells(2, 9), .Cells(2, 22)).VerticalAlignment = xlBottom
        Range(.Cells(2, 9), .Cells(2, 22)).Font.ColorIndex = 2
        Range(.Cells(2, 9), .Cells(2, 22)).Font.Underline = True
        Range(.Cells(2, 9), .Cells(2, 22)).Font.Bold = True
        Range(.Cells(2, 7), .Cells(2, 22)).EntireColumn.ColumnWidth = 12
''            Range("L1").EntireColumn.ColumnWidth = 13
        Range(.Cells(3, 8), .Cells(4, 9)).Font.ColorIndex = 2
        ' Add comments
''            Dim c As Comment, lIdx As Long
        Range(.Cells(3, 9), .Cells(3, 9)).AddComment "The Consolidated RCV Margin for the Class selected"
        Range(.Cells(4, 9), .Cells(4, 9)).AddComment "The Retail Value for the Class selected"
        Range(.Cells(3, 11), .Cells(3, 11)).AddComment "The Consolidated RCV Margin for the Entire Class, Month by Month"
        Range(.Cells(4, 11), .Cells(4, 11)).AddComment "The Retail Value for the Entire Class, Month by Month"
        Range(.Cells(3, 22), .Cells(3, 22)).AddComment "The Consolidated RCV Margin Total for the Entire Class"
        Range(.Cells(4, 22), .Cells(4, 22)).AddComment "The Retail Total for the Entire Class"
''            Call RePositionComments
''            For Each c In .Comments
''''                Set c = .Comments(lIdx)
''                c.Shape.Left = c.Parent.Offset(0, 1).Left - 1000
''                c.Shape.Top = c.Parent.Offset(0, 1).Top - 10
''            Next
    
        Unload CBA_BTF_frm_Splash
''        Application.ScreenUpdating = True
        If sFormCalled = "optCG" Then
            .Cells(2, 2).Value = " CG Level Forecasting Tool"
            CBA_BTF_frm_CGNo.Show
        ElseIf sFormCalled = "optSCG" Then
''                .Cells(2, 2).Font.Size = 22
            .Cells(2, 2).Value = " SCG Level Forecasting Tool"
            CBA_BTF_frm_SCGNo.Show
        ElseIf sFormCalled = "optProduct" Then
            .Cells(2, 2).Value = " Product Forecasting Tool"
            CBA_BTF_ProductLevel.Show
        End If
    End With
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CrtForecastingWorkbook", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Public Function BTF_LastFCastTest(ByVal lCG As Long, ByVal lPCls As Long, ByVal sCG_SCG_PD As String, ByVal dtFCDate As Date, ByVal TestApply As String, ByRef DspMsg As String) As Boolean
    ' Will determine whether prior (other table) forecasts exist and display messages, if they do
    ' if TestApply="Test" then the function will return with an error - if "Apply" then will display all msgs and prompt for continuance
    
    ' Now will only display message if the forecast being done if not the most recent
    Dim sSQL As String, sData As String, bFCExists As Boolean, strGroup As String, aArry() As String, aDate(0 To 2, 0 To 2) As String, lDIdx As Long
    Dim sGrp As String, bMsgGiven As Boolean, dtLatest As Date, sMsg As String, sLvlIn As String, sPCls As String, bReForecast As Boolean, dtDate As Date
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    Const FLDSRETURNED = "UserName & ""~"" & DateTimeSubmitted", FLDDEFAULT = "x~01/01/1900 00:00:00"
    sLvlIn = sCG_SCG_PD
    bFCExists = False
    dtDate = CDate("01/01/" & Year(dtFCDate))
    sPCls = CBA_BTF_pclassDecypher(CStr(lPCls))
    BTF_LastFCastTest = False                     ' False means that the procedure has no errors
    sSQL = "(CG = " & lCG & ") AND (ProductClass = " & lPCls & ")" _
            & " AND (ForecastDate = #" & dtDate & "#) "        ''AND (ForecastDate <= #" & DateAdd("M", 12, dtFCDate) & "#)"

    ' Get if it is a reformat - it may affect if it can be changed or not
    If CBA_BT_getCutOffDate(dtDate) <> "Format" Then bReForecast = True
        
    lDIdx = 0
    ' Set up the decription of the group
    If sLvlIn = "CG" Then
        strGroup = "CG"
    ElseIf sLvlIn = "SCG" Then
        strGroup = "SCG"
    ElseIf sLvlIn = "PD" Then
        strGroup = "Product "
    End If
    
    ' Test to see if there are data in the various Forcasting Data tables
    sGrp = "CG"
    sData = g_DLookup(FLDSRETURNED, "CGData", sSQL, "DateTimeSubmitted DESC", g_GetDB("ForeCast"), FLDDEFAULT)
    GoSub ChkData
    sGrp = "SCG"
    sData = g_DLookup(FLDSRETURNED, "SCGData", sSQL, "DateTimeSubmitted DESC", g_GetDB("ForeCast"), FLDDEFAULT)
    GoSub ChkData
    sGrp = "PD"
    sData = FLDDEFAULT
    GoSub ChkData
    
Exit_Routine:
    ' If there are forecasting items that exist in whatever table, see if it is the on the form that is calling, if so return true regardless that this is the latest to be returned
    If bFCExists = True Then
        dtLatest = CDate("01/01/1900")
        sGrp = ""
        For lDIdx = 0 To 2
            If dtLatest < CDate(aDate(1, lDIdx)) Then
                dtLatest = CDate(aDate(1, lDIdx))
                Call CBA_getUserShortTitle("", "", aDate(2, lDIdx), sData)
                sGrp = aDate(0, lDIdx)
            End If
        Next
        ' If there are data in the other various Forcasting Data tables
        bFCExists = False
        If sLvlIn <> sGrp Then
            If sGrp = "CG" Or sGrp = "SCG" Or sGrp = "PD" Then
                GoSub MsgChk
            End If
        Else
            DspMsg = ""
        End If
        'Debug.Print sGrp & "-" & dtLatest
        If sGrp <> sLvlIn And TestApply = "Test" Then
            BTF_LastFCastTest = False
        ElseIf sGrp <> sLvlIn And TestApply = "Apply" Then
            BTF_LastFCastTest = bFCExists
        End If
    Else
        DspMsg = ""
    End If
Exit_Routine2:
    On Error Resume Next
    Exit Function

ChkData:
    aArry = Split(sData & "~" & sLvlIn, "~")
    aDate(0, lDIdx) = sGrp
    aDate(1, lDIdx) = aArry(1)
    aDate(2, lDIdx) = aArry(0)
    If aArry(0) <> "x" Then bFCExists = True
    lDIdx = lDIdx + 1
    Return
    
MsgChk:
    sMsg = "This " & IIf(sLvlIn = "PD", "Product", sLvlIn) & " for " & Year(dtFCDate) & ", is currently being forecasted at " & IIf(sGrp = "PD", "Product", sGrp) & _
            "(" & sPCls & " class) level, by " & sData & vbCrLf & vbCrLf
    If (sGrp = "PD" Or sGrp = "SCG") And (sLvlIn = "PD" Or sLvlIn = "SCG") Then bReForecast = False
    If TestApply = "Test" Then
        If bMsgGiven = False And sLvlIn <> sGrp Then
            If bReForecast = True Then
                sMsg = sMsg & "The forecast cannot be changed to a" & IIf(sLvlIn = "SCG", "n ", " ") & sLvlIn & " level for " & Year(dtFCDate) & ", as it is a reforecast"
                bFCExists = True
            Else
                sMsg = sMsg & "If this forecast is applied, only the " & IIf(sLvlIn = "PD", "Product", sLvlIn) & " level forcast for " & Year(dtFCDate) & " will apply"
            End If
            DspMsg = sMsg
            bMsgGiven = True
        End If
    Else
        If bReForecast = True Then
            sMsg = sMsg & "The forecast cannot be changed to a" & IIf(sLvlIn = "SCG", "n ", " ") & IIf(sLvlIn = "PD", "Product", sLvlIn) & " level at this time as it is a reforecast"
            MsgBox sMsg, vbOKOnly
            bFCExists = True
        Else
            sMsg = sMsg & "Do you wish to reapply this forecast at the " & IIf(sLvlIn = "PD", "Product", sLvlIn) & "(" & sPCls & " class) level?"
            If MsgBox(sMsg, vbYesNo) = vbNo Then bFCExists = True
        End If
    End If
    Return
    
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-BTF_LastFCastTest", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine2
    Resume Next
End Function

Public Function UpdateTotals(ByRef wks As Worksheet, ByRef CGData As Variant, ByVal curPClass As Long, ByVal curYYIdx As Long, Optional bFreeze As Boolean = False)
   ' Put the totals for the CG/PClass at the top af the spreadsheet
    Dim TotRet(1 To 12) As Single, TotQTY(1 To 12) As Single, d As Long, c As Long, TTotRet As Single, TTotQTY As Single
    Dim TCQTY As Single, TCRET As Single
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    If bFreeze Then Application.ScreenUpdating = False
    
    ''lElemntNo = Month(Date)
    Erase TotQTY: Erase TotRet: TTotQTY = 0: TTotRet = 0: TCQTY = 0: TCRET = 0

    For d = 1 To 4      ' For each class
            For c = 1 To 12 ' For Each month
                On Error Resume Next
                TotQTY(c) = TotQTY(c) + CGData(d).ForeRCVMargin((curYYIdx * 12) + c) * CGData(d).ForeRetail((curYYIdx * 12) + c)
                TotRet(c) = TotRet(c) + CGData(d).ForeRetail((curYYIdx * 12) + c)
                TTotQTY = TTotQTY + CGData(d).ForeRCVMargin((curYYIdx * 12) + c) * CGData(d).ForeRetail((curYYIdx * 12) + c)
                TTotRet = TTotRet + CGData(d).ForeRetail((curYYIdx * 12) + c)
                Err.Clear
                On Error GoTo Err_Routine
            Next
    Next
    For c = 1 To 12  ' For the current Class
        On Error Resume Next
        TCQTY = TCQTY + CGData(curPClass).ForeRCVMargin((curYYIdx * 12) + c) * CGData(curPClass).ForeRetail((curYYIdx * 12) + c)
        TCRET = TCRET + CGData(curPClass).ForeRetail((curYYIdx * 12) + c)
        Err.Clear
        On Error GoTo Err_Routine
        If NZ(TotRet(c), 0) <> 0 Then
            wks.Cells(3, 9 + c) = Format(TotQTY(c) / TotRet(c), "0.00%")
        Else
            wks.Cells(3, 9 + c) = Format(0, "0.00%")
        End If
        wks.Cells(4, 9 + c) = Format(TotRet(c), "#,0")
    Next
    If NZ(TCRET, 0) <> 0 Then
        wks.Cells(3, 9) = Format(TCQTY / TCRET, "0.00%")
    Else
        wks.Cells(3, 9) = Format(0, "0.00%")
    End If
    wks.Cells(4, 9) = Format(TCRET, "#,0")
    Range(wks.Cells(3, 9), wks.Cells(4, 9)).Font.ColorIndex = 1
    If NZ(TTotRet, 0) <> 0 Then
        wks.Cells(3, 22) = Format(TTotQTY / TTotRet, "0.00%")
    Else
        wks.Cells(3, 22) = Format(0, "0.00%")
    End If
    wks.Cells(4, 22) = Format(TTotRet, "#,0")
    If bFreeze Then Application.ScreenUpdating = True
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-UpdateTotals", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function BTF_Start_Date(ByVal dtDate As Date) As Date

    ' Will return the 'Forecast From' date for the various forecast screens
    'BTF_Start_Date("31/12/2018") '   Hard code the date for now until we are ready to let it go (would usually be today's date)
    If Month(dtDate) < 2 Then
        dtDate = CDate("01/01/" & (Year(dtDate)))
    ElseIf Month(dtDate) > 10 Then
''        dtDate = Date       '''''CDate("01/01/" & (Year(dtDate)) + 1) '@RWFC will have to fix as it is wrong
    Else
        dtDate = CDate("01/07/" & (Year(dtDate)))
    End If
    BTF_Start_Date = dtDate

End Function

Public Function BTF_GetElDate(lInputYY As Long) As Long
    ' As there are 36 elements now, we have to use the year to decide which set of elements are to be updated.
    If lInputYY < 2000 Then
        BTF_GetElDate = lInputYY
    ElseIf lInputYY < Year(Date) Then
        BTF_GetElDate = 0
    ElseIf lInputYY = Year(Date) Then
        BTF_GetElDate = 1
    ElseIf lInputYY = Year(Date) + 1 Then
        BTF_GetElDate = 2
    ElseIf lInputYY = Year(Date) + 2 Then
        BTF_GetElDate = 3
    Else
        MsgBox "Index Error"
    End If
    
End Function


