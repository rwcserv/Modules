Attribute VB_Name = "CBA_DBU_Runtime"
Option Explicit                 ' CBA_DBU_Runtime 190731
Option Private Module          ' Excel users cannot access procedures

Type DBU_Pclass
    CoreRange As Single
    Food As Single
    NonFood As Single
    Seasonal As Single
End Type
Type DBU_PaLCosts
    PAYROLL_EXPENSES As Single
    EMPLOYEE_BENEFITS As Single
    OTHER_PERSONNEL_EXPENSES As Single
    PROPERTY_EXPENSES As Single
    UTILITIES As Single
    MONEY_TRANSFER As Single
    TRANSPORT_EXPENSES As Single
    ADVERTISING As Single
    EXAMINATION_EXPENSES As Single
    IT_EXPENSES As Single
    PROFESSIONAL_FEES As Single
    TRAVEL_EXPENSES As Single
    IC_SERVICES As Single
    TAXES As Single
    UNIFORMS As Single
    COMMUNICATIONS As Single
    MATERIALS As Single
    SECURITY As Single
    WASTE_DISPOSAL As Single
    OTHER_EXPENSES_OTHER As Single
    DEPRECIATION As Single
    Sales As Single
    Cost_of_goods_sold As Single
    ASSET_DEVELOPMENT As Single
    Summary_of_Costs As Single
    Income_IC_Services As Single
End Type
Type DBUnit
    Name As String
    Ref As String
    DBUData() As CBA_DBU
End Type
Private DBUCGList As Variant
Private ActiveDBUs As Scripting.Dictionary
Private ActiveDBUCC As Scripting.Dictionary
Private CBA_DBUarr() As DBUnit
Sub RunDBUForecastReport(ByVal DateFrom As Date, ByVal DateTo As Date, ByRef strDBUforReport As String)
Dim wbkDBU As Workbook
Dim wksDBU As Worksheet
Dim dbs As Long, DBUcnt As Long, rw As Long, a As Long, b As Long, c As Long, col As Long, brw As Long
Dim past As Boolean
Dim TotBus As Single, TotbusOrig As Single, TotRCVOrigBus As Single, TotRCVReBus As Single, TotbusRe As Single, TotCont As Single
Dim strForm As String, strSum As String, strAvg As String, strFinal As String
Dim TotbusRCV As Single

    If DBU_Data_Create(DateFrom, DateTo, strDBUforReport) = True Then
        
        Set wbkDBU = Application.Workbooks.Add
        Set wksDBU = wbkDBU.Sheets(1)
        With wksDBU
            Range(.Cells(1, 1), .Cells(5, 79)).Interior.ColorIndex = 49
            .Cells(1, 1).Select
            .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
            .Cells(4, 3).Value = "DBU Forecast Report"
            .Cells(4, 3).Font.Name = "ALDI SUED Office"
            .Cells(4, 3).Font.Size = 24
            .Cells(4, 3).Font.ColorIndex = 2
            DBUcnt = 0
            rw = 8
            .Cells(rw, 1).Value = "TOTAL"
            Range(.Cells(rw, 1), .Cells(rw + 3, 1)).Merge
            .Cells(rw, 2).Value = "Sales"
            .Cells(rw + 1, 2).Value = "Share"
            .Cells(rw + 2, 2).Value = "Margin"
            .Cells(rw + 3, 2).Value = "Cont.($)"
            TotRCVOrigBus = 0: TotRCVReBus = 0
            For dbs = LBound(CBA_DBUarr) To UBound(CBA_DBUarr)
'                If InStr(1, CBA_DBUarr(dbs).Name, "e-Commerce") > 0 Then
'                a = a
'                End If
                On Error Resume Next
                If UBound(CBA_DBUarr(dbs).DBUData) = -1 Then
                    Err.Clear
                    On Error GoTo 0
                Else
                    On Error GoTo 0
                    rw = rw + 4
                    DBUcnt = DBUcnt + 1
                    .Cells(rw, 1).Value = CBA_DBUarr(dbs).Name

                    Range(.Cells(rw, 1), .Cells(rw + 3, 1)).Merge
                    .Cells(rw, 2).Value = "Sales"
                    .Cells(rw + 1, 2).Value = "Share"
                    .Cells(rw + 2, 2).Value = "Margin"
                    .Cells(rw + 3, 2).Value = "Cont.(%)"
                    For a = LBound(CBA_DBUarr(dbs).DBUData) To UBound(CBA_DBUarr(dbs).DBUData)
                        If CBA_DBUarr(dbs).DBUData(a).MonthNum < Month(Date) And CBA_DBUarr(dbs).DBUData(a).YearNum <= Year(Date) Then past = True Else past = False
                        .Cells(6, 3 + (a * 3)).Value = CBA_DBUarr(dbs).DBUData(a).MonthNum & "-" & CBA_DBUarr(dbs).DBUData(a).YearNum
                        Range(.Cells(6, 3 + (a * 3)), .Cells(6, 2 + ((a + 1) * 3))).Merge
                        .Cells(7, 3 + (a * 3) + 0).Value = "Orig Fcst."
                        .Cells(7, 3 + (a * 3) + 1).Value = "ReFcst."
                        .Cells(7, 3 + (a * 3) + 2).Value = "Actual"
                        'NEED a total business value
                        TotBus = CBA_DBUarr(dbs).DBUData(a).TotalBusinessRetail
                        TotbusOrig = CBA_DBUarr(dbs).DBUData(a).TotalBusinessOFore
                        TotbusRe = CBA_DBUarr(dbs).DBUData(a).TotalBusinessRetailReFore
                        TotbusRCV = CBA_DBUarr(dbs).DBUData(a).TotalBusinessRetailRCV
                        'If CBA_DBUarr(dbs).DBUData(a).MonthNum = 1 Then Debug.Print CBA_DBUarr(dbs).Name & " - " & Format(CBA_DBUarr(dbs).DBUData(a).RCVData("Retail"), "#")
                        '.Cells(Rw, 3 + ((UBound(CBA_DBUarr(dbs).DBUData) + 2) * 3)).Value = .Cells(Rw, 3 + ((UBound(CBA_DBUarr(dbs).DBUData) + 2) * 3)).Value + CBA_DBUarr(dbs).DBUData(a).RCVData("Retail")
                        If past = True Then
                            If DBUcnt = 1 Then
                                TotRCVOrigBus = TotRCVOrigBus + TotbusRCV
                                TotRCVReBus = TotRCVReBus + TotbusRCV
                            End If
                            .Cells(rw, 3 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value = .Cells(rw, 3 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value + CBA_DBUarr(dbs).DBUData(a).RCVData("Retail")
                            .Cells(rw, 4 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value = .Cells(rw, 4 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value + CBA_DBUarr(dbs).DBUData(a).RCVData("Retail")
                            .Cells(rw, 5 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value = .Cells(rw, 5 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value + CBA_DBUarr(dbs).DBUData(a).RCVData("Retail")
                            .Cells(rw + 3, 3 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value = .Cells(rw + 3, 3 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value + CBA_DBUarr(dbs).DBUData(a).MarginOriginalForecast
                            .Cells(rw + 3, 4 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value = .Cells(rw + 3, 4 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value + CBA_DBUarr(dbs).DBUData(a).MarginReForecast
                            .Cells(rw + 3, 5 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value = .Cells(rw + 3, 5 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value + IIf(CBA_DBUarr(dbs).DBUData(a).MarginReForecast = 0, CBA_DBUarr(dbs).DBUData(a).MarginOriginalForecast, CBA_DBUarr(dbs).DBUData(a).MarginReForecast)
                        Else
                            If DBUcnt = 1 Then
                                TotRCVOrigBus = TotRCVOrigBus + TotbusOrig
                                TotRCVReBus = TotRCVReBus + TotbusRe
                            End If
                            .Cells(rw, 3 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value = .Cells(rw, 3 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value + CBA_DBUarr(dbs).DBUData(a).RetailOriginalForecast
                            .Cells(rw, 4 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value = .Cells(rw, 4 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value + CBA_DBUarr(dbs).DBUData(a).RetailReForecast
                            .Cells(rw, 5 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value = .Cells(rw, 5 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value + IIf(CBA_DBUarr(dbs).DBUData(a).RetailReForecast = 0, CBA_DBUarr(dbs).DBUData(a).RetailOriginalForecast, CBA_DBUarr(dbs).DBUData(a).RetailReForecast)
                            .Cells(rw + 3, 3 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value = .Cells(rw + 3, 3 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value + CBA_DBUarr(dbs).DBUData(a).ContributionOriginalForecast
                            .Cells(rw + 3, 4 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value = .Cells(rw + 3, 4 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value + CBA_DBUarr(dbs).DBUData(a).ContributionReForecast
                            .Cells(rw + 3, 5 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value = .Cells(rw + 3, 5 + ((UBound(CBA_DBUarr(dbs).DBUData) + 1) * 3)).Value + IIf(CBA_DBUarr(dbs).DBUData(a).ContributionReForecast = 0, CBA_DBUarr(dbs).DBUData(a).ContributionOriginalForecast, CBA_DBUarr(dbs).DBUData(a).ContributionReForecast)
                        End If
                        
                        For b = 0 To 2
                            TotCont = 0
                            If b = 0 Then
                                .Cells(rw, 3 + (a * 3) + b).Value = CBA_DBUarr(dbs).DBUData(a).RetailOriginalForecast / 1000000
                                If TotbusOrig = 0 Then
                                    .Cells(rw + 1, 3 + (a * 3) + b).Value = 0
                                Else
                                    .Cells(rw + 1, 3 + (a * 3) + b).Value = (CBA_DBUarr(dbs).DBUData(a).RetailOriginalForecast / TotbusOrig) * 100
                                End If
                                If past = True Then
                                    If CBA_DBUarr(dbs).DBUData(a).RCVData("Retail") = 0 Then
                                        .Cells(rw + 2, 3 + (a * 3) + b).Value = 0
                                    Else
                                        .Cells(rw + 2, 3 + (a * 3) + b).Value = (CBA_DBUarr(dbs).DBUData(a).MarginOriginalForecast / CBA_DBUarr(dbs).DBUData(a).RCVData("Retail")) * 100
                                    End If
                                    '.Cells(10, 3 + (a * 3) + b).Value = .Cells(10, 3 + (a * 3) + b).Value + (CBA_DBUarr(dbs).DBUData(a).RCVData("Retail") / 1000000)
                                    TotCont = TotCont + (CBA_DBUarr(dbs).DBUData(a).MarginOriginalForecast / 1000000)
                                    If InStr(1, CBA_DBUarr(dbs).Name, "e-Commerce") > 0 Then .Cells(rw + 2, 3 + (a * 3) + b).Value = 100
                                    If InStr(1, CBA_DBUarr(dbs).Name, "e-Commerce") > 0 Then TotCont = TotCont + (CBA_DBUarr(dbs).DBUData(a).RCVData("Retail") / 1000000)
                                Else
                                    If CBA_DBUarr(dbs).DBUData(a).RetailOriginalForecast = 0 Then
                                        .Cells(rw + 2, 3 + (a * 3) + b).Value = 0
                                        .Cells(11, 3 + (a * 3) + b).Value = .Cells(11, 3 + (a * 3) + b).Value + 0
                                    Else
                                        .Cells(rw + 2, 3 + (a * 3) + b).Value = (CBA_DBUarr(dbs).DBUData(a).ContributionOriginalForecast / CBA_DBUarr(dbs).DBUData(a).RetailOriginalForecast) * 100
                                        TotCont = TotCont + (CBA_DBUarr(dbs).DBUData(a).ContributionOriginalForecast / 1000000)
                                    End If
                                End If
                            
                            ElseIf b = 1 Then
                                .Cells(rw, 3 + (a * 3) + b).Value = CBA_DBUarr(dbs).DBUData(a).RetailReForecast / 1000000
                                If TotbusRe = 0 Then
                                    .Cells(rw + 1, 3 + (a * 3) + b).Value = 0
                                Else
                                    .Cells(rw + 1, 3 + (a * 3) + b).Value = CBA_DBUarr(dbs).DBUData(a).RetailReForecast / TotbusRe * 100
                                End If
                                If past = True Then
                                    If CBA_DBUarr(dbs).DBUData(a).RCVData("Retail") = 0 Then
                                        .Cells(rw + 2, 3 + (a * 3) + b).Value = 0
                                    Else
                                        .Cells(rw + 2, 3 + (a * 3) + b).Value = (CBA_DBUarr(dbs).DBUData(a).MarginReForecast / CBA_DBUarr(dbs).DBUData(a).RCVData("Retail")) * 100
                                    End If
                                    TotCont = TotCont + (CBA_DBUarr(dbs).DBUData(a).MarginReForecast / 1000000)
                                    '.Cells(10, 3 + (a * 3) + b).Value = .Cells(10, 3 + (a * 3) + b).Value + CBA_DBUarr(dbs).DBUData(a).RCVData("Retail")
                                    If InStr(1, CBA_DBUarr(dbs).Name, "e-Commerce") > 0 Then .Cells(rw + 2, 3 + (a * 3) + b).Value = 100
                                    If InStr(1, CBA_DBUarr(dbs).Name, "e-Commerce") > 0 Then TotCont = TotCont + (CBA_DBUarr(dbs).DBUData(a).RCVData("Retail") / 1000000)
                                Else
                                    If CBA_DBUarr(dbs).DBUData(a).RetailReForecast = 0 Then
                                        .Cells(rw + 2, 3 + (a * 3) + b).Value = 0
                                        .Cells(11, 3 + (a * 3) + b).Value = .Cells(11, 3 + (a * 3) + b).Value + 0
                                    Else
                                        .Cells(rw + 2, 3 + (a * 3) + b).Value = (CBA_DBUarr(dbs).DBUData(a).ContributionReForecast / CBA_DBUarr(dbs).DBUData(a).RetailReForecast) * 100
                                        TotCont = TotCont + (CBA_DBUarr(dbs).DBUData(a).ContributionReForecast / 1000000)
                                    End If
                                End If
                            ElseIf b = 2 Then
                                If past = True Then
                                    .Cells(rw, 3 + (a * 3) + b).Value = CBA_DBUarr(dbs).DBUData(a).SalesData("Retail") / 1000000
                                    If TotBus = 0 Then
                                        .Cells(rw + 1, 3 + (a * 3) + b).Value = 0
                                    Else
                                        .Cells(rw + 1, 3 + (a * 3) + b).Value = (CBA_DBUarr(dbs).DBUData(a).SalesData("Retail") / TotBus) * 100
                                    End If
                                    If InStr(1, CBA_DBUarr(dbs).Name, "e-Commerce") > 0 Then
                                        .Cells(rw + 2, 3 + (a * 3) + b).Value = 100
                                        TotCont = TotCont + (CBA_DBUarr(dbs).DBUData(a).RCVData("Retail") / 1000000)
                                    Else
                                        .Cells(rw + 2, 3 + (a * 3) + b).Value = (CBA_DBUarr(dbs).DBUData(a).RCVData("Margin")) * 100
                                        TotCont = TotCont + (CBA_DBUarr(dbs).DBUData(a).RCVData("Contribution") / 1000000)
                                    End If
                                    'If InStr(1, CBA_DBUarr(dbs).Name, "e-Commerce") > 0 Then .Cells(Rw + 2, 3 + (a * 3) + b).Value = 100
                                    'If InStr(1, CBA_DBUarr(dbs).Name, "e-Commerce") > 0 Then TotCont = TotCont + .Cells(Rw, 3 + (a * 3) + b).Value
                                Else
                                    .Cells(rw, 3 + (a * 3) + b).Value = 0
                                    .Cells(rw + 1, 3 + (a * 3) + b).Value = 0
                                    .Cells(rw + 2, 3 + (a * 3) + b).Value = 0
                                    TotCont = TotCont + 0
                                End If
                            End If
                            If CBA_DBUarr(dbs).DBUData(a).YearNum > Year(Date) Or (CBA_DBUarr(dbs).DBUData(a).MonthNum >= Month(Date) And CBA_DBUarr(dbs).DBUData(a).YearNum = Year(Date)) Then
                                If b < 2 Then
                                    If b = 0 Then .Cells(10, 3 + (a * 3) + b).Value = TotbusOrig
                                    If b = 1 Then .Cells(10, 3 + (a * 3) + b).Value = TotbusRe
                                    'If b = 0 Then .Cells(10, 3 + (a * 3) + b).Value = Cells(10, 3 + (a * 3) + b).Value + CBA_DBUarr(dbs).DBUData(a).RetailOriginalForecast
                                    'If b = 1 Then .Cells(10, 3 + (a * 3) + b).Value = Cells(10, 3 + (a * 3) + b).Value + CBA_DBUarr(dbs).DBUData(a).RetailReForecast
                                    If TotCont <> 0 Then .Cells(11, 3 + (a * 3) + b).Value = .Cells(11, 3 + (a * 3) + b).Value + TotCont
                                Else
                                    .Cells(10, 3 + (a * 3) + b).Value = 0
                                    .Cells(11, 3 + (a * 3) + b).Value = 0
                                End If
                                '.Cells(Rw + 3, 3 + (a * 3) + b).Formula = "=((R[-1]C/100)*(R[-2]C/100))*100" '
                                .Cells(rw + 3, 3 + (a * 3) + b).Value = (.Cells(rw + 2, 3 + (a * 3) + b).Value * .Cells(rw + 1, 3 + (a * 3) + b).Value) / 100
                            Else
                                If TotCont <> 0 Then .Cells(11, 3 + (a * 3) + b).Value = .Cells(11, 3 + (a * 3) + b).Value + TotCont
                                '.Cells(10, 3 + (a * 3) + b).Value = Cells(10, 3 + (a * 3) + b).Value + CBA_DBUarr(dbs).DBUData(a).RCVData("Retail")
                                '''.Cells(Rw + 3, 3 + (a * 3) + b).Formula = "=((R[-1]C/100)*(R[-2]C/100))*100"
                                ''''.Cells(10, 3 + (a * 3) + b).Value = Totbus
                                .Cells(rw + 3, 3 + (a * 3) + b).Value = (CBA_DBUarr(dbs).DBUData(a).RCVData("Retail") * .Cells(rw + 2, 3 + (a * 3) + b).Value) / CBA_DBUarr(dbs).DBUData(a).TotalBusinessRetailRCV
                                'If InStr(1, CBA_DBUarr(dbs).Name, "e-Commerce") > 0 Then .Cells(Rw + 3, 3 + (a * 3) + b).Value = .Cells(Rw + 1, 3 + (a * 3) + b).Value
                                .Cells(10, 3 + (a * 3) + b).Value = CBA_DBUarr(dbs).DBUData(a).TotalBusinessRetailRCV
                            End If
                        Next
                    Next
                End If
            Next
            
    
            col = ((DateDiff("M", DateFrom, DateTo) + 1) * 3) + 2
            brw = 12 + (DBUcnt * 4)
            Range(.Cells(8, 3), .Cells(9, col)).Formula = "=SUMIFS(R12C:R" & brw + 4 & "C,R12C2:R" & brw + 4 & "C2,RC2)"
            For a = 3 To col
                If .Cells(10, a).Value <> 0 Then .Cells(10, a).Value = (.Cells(11, a).Value / (.Cells(10, a).Value / 1000000)) * 100
            Next
            Range(.Cells(8, 1), .Cells(brw - 1, col + 3)).Borders.Weight = xlThin
            For a = 1 To (DateDiff("M", DateFrom, DateTo) + 2)
                Range(.Cells(8, 3 * a), .Cells(brw - 1, (3 * a) + 2)).BorderAround , xlMedium
            Next
            For a = 1 To DBUcnt + 1
                Range(.Cells(8 + ((a - 1) * 4), 1), .Cells(8 + ((a - 1) * 4) + 3, col + 3)).BorderAround , xlMedium
            Next
            Range(.Cells(6, col + 1), .Cells(6, col + 3)).Merge
            .Cells(6, col + 1).Value = "TOTAL"
            .Cells(7, col + 1).Value = "Orig Fcst."
            .Cells(7, col + 2).Value = "ReFcst."
            .Cells(7, col + 3).Value = "Actual (+Fcst.)"
            
            strForm = ""
            strFinal = ""
            strSum = "=SUM("
            strAvg = "=AVERAGE("
            b = -3
            For a = col - 2 To 3 Step b
                If Year(.Cells(6, (col - a) + 1).MergeArea(1, 1).Value) > Year(Date) Or (Year(.Cells(6, (col - a) + 1).MergeArea(1, 1).Value) = Year(Date) And Month(.Cells(6, (col - a) + 1).MergeArea(1, 1).Value) >= Month(Date)) Then
                    If a = 3 Then strFinal = strFinal & "RC[-" & a + 1 & "])" Else strFinal = strFinal & "RC[-" & a + 1 & "],"
                Else
                    If a = 3 Then strFinal = strFinal & "RC[-" & a & "])" Else strFinal = strFinal & "RC[-" & a & "],"
                End If
                If a = 3 Then strForm = strForm & "RC[-" & a & "])" Else strForm = strForm & "RC[-" & a & "],"
            Next
            
            
            For a = 0 To DBUcnt
                If a > 0 Then
                    For b = 1 To 3
                        .Cells(10 + (a * 4), col + b).Value = (.Cells(11 + (a * 4), col + b).Value / .Cells(8 + (a * 4), col + b).Value) * 100
                        If b = 1 Then .Cells(11 + (a * 4), col + b).Value = .Cells(10 + (a * 4), col + b).Value * (.Cells(8 + (a * 4), col + b).Value / TotRCVOrigBus)
                        If b = 2 Then .Cells(11 + (a * 4), col + b).Value = .Cells(10 + (a * 4), col + b).Value * (.Cells(8 + (a * 4), col + b).Value / TotRCVReBus)
                        If b = 3 Then
                            If InStr(1, .Cells(11 + (a * 4), 1).MergeArea(1, 1).Value, "e-Commerce") > 0 Then
                            Else
                                .Cells(11 + (a * 4), col + b).Value = strAvg & strFinal
                                .Cells(10 + (a * 4), col + b).Value = strAvg & strFinal
                            End If
                        End If
                    Next
                End If
                Range(.Cells(8 + (a * 4), col + 1), .Cells(8 + (a * 4), col + 2)).Formula = strSum & strForm
                Range(.Cells(8 + (a * 4), col + 3), .Cells(8 + (a * 4), col + 3)).Formula = strSum & strFinal
                Range(.Cells(9 + (a * 4), col + 1), .Cells(9 + (a * 4), col + 2)).Formula = strAvg & strForm
                Range(.Cells(9 + (a * 4), col + 3), .Cells(9 + (a * 4), col + 3)).Formula = strAvg & strFinal
                If InStr(1, .Cells(11 + (a * 4), 1).MergeArea(1, 1).Value, "e-Commerce") > 0 Then
                    Range(.Cells(10 + (a * 4), col + 1), .Cells(10 + (a * 4), col + 3)).Value = 100
                    .Cells(11 + (a * 4), col + 1).Value = .Cells(9 + (a * 4), col + 1).Value
                    .Cells(11 + (a * 4), col + 2).Value = .Cells(9 + (a * 4), col + 2).Value
                    .Cells(11 + (a * 4), col + 3).Value = .Cells(9 + (a * 4), col + 3).Value
                End If
                If a = 0 Then
                    Range(.Cells(10 + (a * 4), col + 1), .Cells(10 + (a * 4), col + 2)).Formula = strAvg & strForm
                    Range(.Cells(10 + (a * 4), col + 3), .Cells(10 + (a * 4), col + 3)).Formula = strAvg & strFinal
                    'Range(.Cells(10 + (a * 4), col + 1), .Cells(10 + (a * 4), col + 3)).Formula = "=iferror((R[1]C/R[-2]C)*100,0)"
                    Range(.Cells(11 + (a * 4), col + 1), .Cells(11 + (a * 4), col + 2)).Formula = strSum & strForm
                    Range(.Cells(11 + (a * 4), col + 3), .Cells(11 + (a * 4), col + 3)).Formula = strSum & strFinal
                End If
            Next
            
            
            
            
            Range(.Cells(6, col + 1), .Cells(7, col + 3)).Borders.Weight = xlThin
            Range(.Cells(6, 3), .Cells(7, col)).Borders.Weight = xlThin
            Range(.Cells(6, 1), .Cells(brw + 4, col + 3)).HorizontalAlignment = xlCenter
            Range(.Cells(6, 1), .Cells(brw + 4, col + 3)).VerticalAlignment = xlCenter
            Range(.Cells(8, 3), .Cells(brw + 4, col + 3)).NumberFormat = "#,0.00"
            Range(.Cells(8, 1), .Cells(brw + 4, 1)).EntireRow.RowHeight = 14.25
        End With
    End If
    If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.CBA_Close_Running
End Sub
Function DBU_Data_Create(ByVal DateFrom As Date, ByVal DateTo As Date, Optional ByVal DBUs As String) As Boolean
Dim strUnusedDBURefs As String
Dim DBUU() As String, strSQL As String, strformess As String
Dim DBUD As Variant, ActiveCGList As Variant
Dim DBUDesc As Scripting.Dictionary, CGallocate As Scripting.Dictionary
Dim CntUnUsed As Long, d As Long, a As Long, b As Long, c As Long, lAdd As Long, dbs As Long, lCurYear As Long
Dim r As Long, cnt As Long, totMnth As Long, mnhn As Long, Y As Long
Dim CBA_COM_CBISRS As ADODB.Recordset
Dim Listing As Variant
Dim curActiveCG As Byte, bytCurCG As Byte, m As Byte, bytCurMonth As Byte, iCurSCG As Integer
Dim atCG As Boolean, bAdded As Boolean
Dim totBusForeCalc As Scripting.Dictionary, totBusCalc As Scripting.Dictionary, subdic As Scripting.Dictionary
Dim SubSubDic As Scripting.Dictionary, breakDic As Scripting.Dictionary, SubbreakDic As Scripting.Dictionary
Dim totBusRCVCalc As Scripting.Dictionary

', Optional ByRef ForeRP As CBA_BTF_ReportParamaters

If isDate(DateFrom) = False Or isDate(DateTo) = False Then
    MsgBox "DateFrom or DateTo is not date"
    DBU_Data_Create = False
    Exit Function
End If

If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Gathering DBU Mappings.."
CBA_SQL_Queries.CBA_GenPullSQL "DBU_List"
Set DBUDesc = New Scripting.Dictionary
For r = LBound(CBA_ABIarr, 2) To UBound(CBA_ABIarr, 2)
    DBUDesc.Add CBA_ABIarr(1, r), CBA_ABIarr(2, r)
Next






If DBUs <> "" Then
    'If InStr(1, DBUs, ",") = 0 Then DBUs = DBUs & ","
    DBUD = Split(DBUs, ",")
    CntUnUsed = 0
    For d = LBound(DBUD) To UBound(DBUD)
        If checkDBUs(DBUD(d)) = "" Then
            If strUnusedDBURefs = "" Then strUnusedDBURefs = DBUD(d) Else strUnusedDBURefs = strUnusedDBURefs & ", " & DBUD(d)
            DBUD(d) = ""
            CntUnUsed = CntUnUsed + 1
            'this process looks at what the active DBUs are, and if a live one has been requested it is excluded from the setup of CBA_DBU's
        End If
    Next
    If UBound(DBUD) - CntUnUsed >= 0 Then
    ReDim DBUU(0 To UBound(DBUD) - CntUnUsed)
    lAdd = -1
    For d = LBound(DBUD) To UBound(DBUD)
        If DBUD(d) <> "" Then
            lAdd = lAdd + 1
            DBUU(lAdd) = DBUD(d)
            'this now adds all DBUs that are to be created
        End If
    Next
    Else
        MsgBox "No Active DBUs to query"
        DBU_Data_Create = False
        Exit Function
    End If
Else
    CBA_SQL_Queries.CBA_GenPullSQL "DBU_List"
    cnt = 0
    ReDim DBUU(0 To UBound(CBA_ABIarr, 2))
    For a = LBound(CBA_ABIarr, 2) To UBound(CBA_ABIarr, 2)
        DBUU(a) = CBA_ABIarr(1, a)
    Next
    Erase CBA_ABIarr
End If


totMnth = DateDiff("M", DateFrom, DateTo)
'If totMnth = 0 Then totMnth = 1
If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Gathering CBIS Data.."

Set CBA_COM_CBISCN = New ADODB.Connection
With CBA_COM_CBISCN
    .ConnectionTimeout = 50
    .CommandTimeout = 50
    .Open "Provider= SQLNCLI10; DATA SOURCE= 599DBL01; ;INTEGRATED SECURITY=sspi;"
End With

strSQL = strSQL & "select p.cgno, isnull(p.scgno,0) as scgno, p.productclass" & Chr(10)
strSQL = strSQL & "--select acg.CatNo as ACGCat, acg.CGNo as ACGCG, acg.SCGNo as ACGSCG, p.productclass" & Chr(10)
strSQL = strSQL & "from cbis599p.dbo.product p" & Chr(10)
strSQL = strSQL & "--left join cbis599p.dbo.tf_ACGMap() acg on acg.ACGEntityID = p.ACGEntityID" & Chr(10)
strSQL = strSQL & "group by p.cgno, isnull(p.scgno,0) , p.productclass" & Chr(10)
strSQL = strSQL & "--group by acg.CatNo , acg.CGNo , acg.SCGNo , p.productclass" & Chr(10)
strSQL = strSQL & "order by p.cgno, isnull(p.scgno,0) , p.productclass" & Chr(10)
strSQL = strSQL & "--order by acg.CatNo , acg.CGNo , acg.SCGNo , p.productclass" & Chr(10)
strSQL = strSQL & "--ACG links at productcode, so you will need to convert to ACG or not at all, this can be done by swapping out the commented lines of code" & Chr(10)
strSQL = strSQL & "--Please be aware that subsequent lines of VBA code will also need to be updated for ACG functionality" & Chr(10)


Set CBA_COM_CBISRS = New ADODB.Recordset
CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN
ActiveCGList = CBA_COM_CBISRS.GetRows()

CBA_SQL_Queries.CBA_GenPullSQL "CBA_Forecast_TotalBus", , , Month(DateFrom), Year(DateFrom), Month(DateTo), Year(DateTo)
Set totBusForeCalc = New Scripting.Dictionary
Set subdic = New Scripting.Dictionary: Set SubSubDic = New Scripting.Dictionary
Set breakDic = New Scripting.Dictionary: Set SubbreakDic = New Scripting.Dictionary
lCurYear = CBA_CBFCarr(1, 0)
bytCurMonth = CBA_CBFCarr(0, 0)
For a = LBound(CBA_CBFCarr, 2) To UBound(CBA_CBFCarr, 2)
    If CBA_CBFCarr(1, a) <> lCurYear Or a = UBound(CBA_CBFCarr, 2) Then
        If a = UBound(CBA_CBFCarr, 2) Then
            subdic.Add CBA_CBFCarr(0, a), CBA_CBFCarr(2, a)
            SubSubDic.Add CBA_CBFCarr(0, a), CBA_CBFCarr(3, a)
        End If
        breakDic.Add lCurYear, subdic
        SubbreakDic.Add lCurYear, SubSubDic
        totBusForeCalc.Add "Original", breakDic
        totBusForeCalc.Add "Reforecast", SubbreakDic
        If a = UBound(CBA_CBFCarr, 2) Then Exit For
        Set subdic = New Scripting.Dictionary: Set SubSubDic = New Scripting.Dictionary
        Set breakDic = New Scripting.Dictionary: Set SubbreakDic = New Scripting.Dictionary
        lCurYear = CBA_CBFCarr(1, a)
    End If
    subdic.Add CBA_CBFCarr(0, a), CBA_CBFCarr(2, a)
    SubSubDic.Add CBA_CBFCarr(0, a), CBA_CBFCarr(3, a)
Next

strSQL = "select year(posdate), month(posdate), sum(retail) from cbis599p.dbo.pos where posdate >= '" & Format(DateFrom, "YYYY-MM-DD") & "' and posdate <= '" & Format(DateTo, "YYYY-MM-DD") & "' group by year(posdate), month(posdate) order by year(posdate), month(posdate)"
Set CBA_COM_CBISRS = New ADODB.Recordset
CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN
If CBA_COM_CBISRS.EOF = False Then
    CBA_CBISarr = CBA_COM_CBISRS.GetRows()
    Set totBusCalc = New Scripting.Dictionary
    Set subdic = New Scripting.Dictionary
    lCurYear = CBA_CBISarr(0, 0)
    bytCurMonth = CBA_CBISarr(1, 0)
    For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
        If CBA_CBISarr(0, a) <> lCurYear Or a = UBound(CBA_CBISarr, 2) Then
            If a = UBound(CBA_CBISarr, 2) Then subdic.Add CStr(CBA_CBISarr(1, a)), CStr(CBA_CBISarr(2, a))
            totBusCalc.Add CStr(lCurYear), subdic
            Set subdic = New Scripting.Dictionary
            lCurYear = CBA_CBISarr(0, a)
        End If
        subdic.Add CStr(CBA_CBISarr(1, a)), CStr(CBA_CBISarr(2, a))
    Next
Else
    Set totBusCalc = New Scripting.Dictionary
    Set subdic = New Scripting.Dictionary
    For Y = Year(DateFrom) To Year(DateTo)
        For a = 1 To 12
            subdic.Add CStr(a), CStr(0)
        Next
        totBusCalc.Add CStr(Y), subdic
        Set subdic = New Scripting.Dictionary
    Next
End If



strSQL = "select year(dayenddate), month(dayenddate), sum(retail) from cbis599p.dbo.receiving where dayenddate >= '" & Format(DateFrom, "YYYY-MM-DD") & "' and dayenddate <= '" & Format(DateTo, "YYYY-MM-DD") & "' group by year(dayenddate), month(dayenddate) order by year(dayenddate), month(dayenddate)"
Set CBA_COM_CBISRS = New ADODB.Recordset
CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN
If CBA_COM_CBISRS.EOF = False Then
    CBA_CBISarr = CBA_COM_CBISRS.GetRows()
    Set totBusRCVCalc = New Scripting.Dictionary
    Set subdic = New Scripting.Dictionary
    lCurYear = CBA_CBISarr(0, 0)
    bytCurMonth = CBA_CBISarr(1, 0)
    For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
        If CBA_CBISarr(0, a) <> lCurYear Or a = UBound(CBA_CBISarr, 2) Then
            If a = UBound(CBA_CBISarr, 2) Then subdic.Add CStr(CBA_CBISarr(1, a)), CStr(CBA_CBISarr(2, a))
            totBusRCVCalc.Add CStr(lCurYear), subdic
            Set subdic = New Scripting.Dictionary
            lCurYear = CBA_CBISarr(0, a)
        End If
        subdic.Add CStr(CBA_CBISarr(1, a)), CStr(CBA_CBISarr(2, a))
    Next
Else
    Set totBusRCVCalc = New Scripting.Dictionary
    Set subdic = New Scripting.Dictionary
    For Y = Year(DateFrom) To Year(DateTo)
        For a = 1 To 12
            subdic.Add CStr(a), CStr(0)
        Next
        totBusRCVCalc.Add CStr(Y), subdic
        Set subdic = New Scripting.Dictionary
    Next
End If






    ReDim CBA_DBUarr(0 To UBound(DBUU))
    For dbs = LBound(DBUU) To UBound(DBUU)
        Listing = GetDBUCGList(DBUU(dbs))
        On Error Resume Next
        If Listing(1, 1) = -1 Then
            Err.Clear
            On Error GoTo 0
            GoTo nextDBU
        End If
        On Error GoTo 0
        Set CGallocate = New Scripting.Dictionary
        curActiveCG = CByte(Listing(1, 1)): atCG = True
        For a = LBound(Listing, 2) To UBound(Listing, 2)
            If CByte(Listing(1, a)) <> curActiveCG Then
                curActiveCG = CByte(Listing(1, a))
                CGallocate.Add curActiveCG, checkMultiDBUCG(curActiveCG)
            End If
        Next a
        bytCurCG = Listing(1, 1): bAdded = False: cnt = 0: iCurSCG = -1
        For a = LBound(Listing, 2) To UBound(Listing, 2)
            If Listing(1, a) <> bytCurCG Then
                bAdded = False: bytCurCG = Listing(1, a)
            End If
            If CGallocate(Listing(1, a)) = True And bAdded = False Then
                cnt = cnt + 1
                bAdded = True
            ElseIf bAdded = True Then
            
            Else
                If Listing(2, a) <> iCurSCG Then
                    iCurSCG = Listing(2, a)
                    cnt = cnt + 1
                End If
            End If
        Next
        ReDim TempArr(0 To 1, 0 To cnt)
        bytCurCG = Listing(1, 1): bAdded = False: cnt = 0: iCurSCG = -1
        For a = LBound(Listing, 2) To UBound(Listing, 2)
            If Listing(1, a) <> bytCurCG Then bAdded = False: bytCurCG = Listing(1, a)
'            If bytCurCG = 58 Then
'            a = a
'            End If
            If CGallocate(Listing(1, a)) = True And bAdded = False Then
                cnt = cnt + 1
                TempArr(0, cnt) = Listing(1, a)
                TempArr(1, cnt) = 0
                bAdded = True
            ElseIf bAdded = True Then
            
            Else
                If Listing(2, a) <> iCurSCG Then
                    iCurSCG = Listing(2, a)
                    cnt = cnt + 1
                    TempArr(0, cnt) = Listing(1, a)
                    TempArr(1, cnt) = Listing(2, a)
                End If
            End If
        Next
        
        
        CBA_BTF_Runtime.setCGs TempArr
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Building Forecast Data Objects.."
        CBA_BTF_SetupForecastArray.CBA_BTF_SetupForecastArray Year(DateFrom), Year(DateTo), 1, 12, , , False
        
        'ReDim totBusForeCalc(0 To totMnth)
        'ReDim totbusCalc(0 To totMnth)
        

        


        
        
        
        
        
        ReDim TempArr(0 To 1, 0 To UBound(Listing, 2) - 1)
        For a = LBound(Listing, 2) To UBound(Listing, 2)
            TempArr(0, a - 1) = Listing(1, a)
            TempArr(1, a - 1) = Listing(2, a)
        Next
        CBA_BTF_Runtime.setCGs TempArr
        ReDim CBA_DBUarr(dbs).DBUData(0 To totMnth)
        mnhn = Month(DateFrom) - 1
        If Right(DBUDesc(DBUU(dbs)), 1) = Chr(10) Then
            strformess = Left(DBUDesc(DBUU(dbs)), Len(DBUDesc(DBUU(dbs)) - 1))
        ElseIf Left(DBUDesc(DBUU(dbs)), 1) = Chr(10) Then
            strformess = Mid(DBUDesc(DBUU(dbs)), 2, 99)
        Else
            strformess = DBUDesc(DBUU(dbs))
        End If
        For m = 0 To totMnth
            If mnhn = 12 Then mnhn = 1 Else mnhn = mnhn + 1
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Creating " & strformess & " DBU Object Data: Month: " & mnhn & " Year: " & Year(DateAdd("M", m, DateFrom))
            Set CBA_DBUarr(dbs).DBUData(m) = New CBA_DBU
            CBA_DBUarr(dbs).Ref = DBUU(dbs)
            CBA_DBUarr(dbs).Name = DBUDesc(DBUU(dbs))
'            If mnhn = 7 Then
'            a = a
'            End If
            
            CBA_DBUarr(dbs).DBUData(m).formulate CByte(mnhn), Year(DateAdd("M", m, DateFrom)), DBUU(dbs), ActiveCGList, NZ(totBusCalc(CStr(Year(DateAdd("M", m, DateFrom))))(CStr(mnhn))), NZ(totBusForeCalc("Original")(Year(DateAdd("M", m, DateFrom)))(mnhn)), NZ(totBusForeCalc("Reforecast")(Year(DateAdd("M", m, DateFrom)))(mnhn)), NZ(totBusRCVCalc(CStr(Year(DateAdd("M", m, DateFrom))))(CStr(mnhn)))

        Next m
        Erase FCbM
nextDBU:
    Next dbs
    
CBA_COM_CBISCN.Close
Set CBA_COM_CBISCN = Nothing
DBU_Data_Create = True








End Function




Function GetDBUCGList(ByVal DBURef As String) As Variant
Dim cnt As Long, a As Long

    On Error Resume Next
    If DBUCGList(1, 1) = 0 Then
        On Error GoTo 0
        Err.Clear
        CBA_SQL_Queries.CBA_GenPullSQL "DBU_Map"
        DBUCGList = CBA_ABIarr
        Erase CBA_ABIarr
    End If
    On Error GoTo 0

    cnt = 0
    For a = LBound(DBUCGList, 2) To UBound(DBUCGList, 2)
        If DBUCGList(1, a) = DBURef Then
            cnt = cnt + 1
        End If
    Next
        If cnt > 0 Then
            ReDim TempArr(1 To 3, 1 To cnt)
            cnt = 0
            For a = LBound(DBUCGList, 2) To UBound(DBUCGList, 2)
                If DBUCGList(1, a) = DBURef Then
                    cnt = cnt + 1
'                    If DBUCGList(2, a) = 58 Then
'                    a = a
'                    End If
                    
                    
                    TempArr(1, cnt) = DBUCGList(2, a)
                    TempArr(2, cnt) = DBUCGList(3, a)
                    TempArr(3, cnt) = DBUCGList(4, a)
                End If
            Next
        End If
    
    If IsEmpty(TempArr) Then
        GetDBUCGList = 0
    Else
        GetDBUCGList = TempArr
    End If
    
End Function
Function checkDBUs(ByVal DBUName As String) As String
Dim arrDBU As Variant
Dim a As Long, b As Long, c As Long
Dim TempArr() As Variant
    If ActiveDBUs Is Nothing Then
        Set ActiveDBUs = New Scripting.Dictionary
        Set ActiveDBUCC = New Scripting.Dictionary
        CBA_SQL_Queries.CBA_GenPullSQL "DBU_List"
        'TempArr = CBA_ABIarr
        'arrDBU = CBA_BasicFunctions.CBA_TransposeArray(TempArr)
        arrDBU = CBA_ABIarr
        For a = LBound(arrDBU, 2) To UBound(arrDBU, 2)
            ActiveDBUs.Add arrDBU(1, a), arrDBU(2, a)
            ActiveDBUCC.Add arrDBU(1, a), arrDBU(3, a)
        Next
    End If
    checkDBUs = ActiveDBUs(DBUName)
End Function
Function checkMultiDBUCG(ByVal CGno As Long) As Boolean
    CBA_SQL_Queries.CBA_GenPullSQL "DBU_Map_CGCheck", , , CGno
    If CBA_ABIarr(0, 0) = 0 Then
        checkMultiDBUCG = False
    Else
        If CBA_ABIarr(0, 0) = 1 Then
            checkMultiDBUCG = True
        Else
            checkMultiDBUCG = False
        End If
    End If
    
End Function
Sub importSATPransactions()
Dim wbkT As Workbook
Dim wshtT As Worksheet, wks As Worksheet
Dim sFile As String, sPath As String
Dim bfound As Boolean
Dim startcell As Long, endcell As Long
Dim RCell As Range
Dim colDic As Scripting.Dictionary
Dim cnt As Long, a As Long
Dim strSQL As String

    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Title = "Select the excel file to import transactions from"
        .Filters.Add "Excel Files", "*.xls?", 1
        .AllowMultiSelect = False
        If .Show = True Then
            sPath = .SelectedItems(1)           ' Get the complete file path.
            sFile = Dir(.SelectedItems(1))      ' Get the file name.
        End If
    End With
    
    If sFile <> "" Then
        Set wbkT = Workbooks.Open(sFile)     ' Open the Excel file.
    Else
        Exit Sub
    End If
    bfound = False
    For Each wks In wbkT.Worksheets
        Debug.Print wks.Name
        If wks.Name = "SAP Transactions" Then
            bfound = True
            Set wshtT = wks
            Exit For
        End If
    Next
    If bfound = False Then Exit Sub
    
    startcell = 0
    For Each RCell In wshtT.Columns(19).Cells
        If RCell.Value <> "" Then
            If startcell = 0 Then startcell = RCell.Row
        Else
            If startcell > 0 Then
                endcell = RCell.Row
                Exit For
            End If
        End If
    Next
    
    Set colDic = New Scripting.Dictionary
    For Each RCell In wshtT.Rows(startcell).Cells
        If RCell.Value = "Amount in local currency" Then colDic.Add RCell.Value, RCell.Column
        If RCell.Value = "Text" Then colDic.Add RCell.Value, RCell.Column
        If RCell.Value = "Account" Then colDic.Add RCell.Value, RCell.Column
        If RCell.Value = "Cost Center" Then colDic.Add RCell.Value, RCell.Column
        If RCell.Value = "Reference Key 1" Then colDic.Add RCell.Value, RCell.Column
        If RCell.Value = "Supplier Name" Then colDic.Add RCell.Value, RCell.Column
        If RCell.Value = "GL Transaction Month" Then colDic.Add RCell.Value, RCell.Column
        If RCell.Value = "Comment" Then colDic.Add RCell.Value, RCell.Column
        If RCell.Value = "P&L Reference" Then colDic.Add RCell.Value, RCell.Column
        If RCell.Value = "GL Amount" Then
            colDic.Add RCell.Value, RCell.Column
            Exit For
        End If
        If RCell.Value = "" And RCell.Column > 50 Then Exit For
    Next
    
    ReDim Tdata(1 To 10, 1 To endcell - startcell)
    cnt = 0
    For Each RCell In wshtT.Columns(colDic("Amount in local currency")).Cells
        If RCell.Row > startcell Then
            cnt = cnt + 1
            If cnt >= endcell - startcell Then Exit For
            Tdata(1, cnt) = wshtT.Cells(RCell.Row, colDic("Amount in local currency")).Value
            Tdata(2, cnt) = wshtT.Cells(RCell.Row, colDic("Text")).Value
            Tdata(3, cnt) = wshtT.Cells(RCell.Row, colDic("Account")).Value
            Tdata(4, cnt) = wshtT.Cells(RCell.Row, colDic("Cost Center")).Value
            Tdata(5, cnt) = wshtT.Cells(RCell.Row, colDic("Reference Key 1")).Value
            Tdata(5, cnt) = Replace(Tdata(5, cnt), "'", "")
            Tdata(6, cnt) = wshtT.Cells(RCell.Row, colDic("Supplier Name")).Value
            Tdata(6, cnt) = Replace(Tdata(6, cnt), "'", "")
            Tdata(7, cnt) = wshtT.Cells(RCell.Row, colDic("GL Transaction Month")).Value
            Tdata(8, cnt) = wshtT.Cells(RCell.Row, colDic("Comment")).Value
            Tdata(8, cnt) = Replace(Tdata(8, cnt), "'", "")
            Tdata(9, cnt) = wshtT.Cells(RCell.Row, colDic("P&L Reference")).Value
            Tdata(10, cnt) = wshtT.Cells(RCell.Row, colDic("GL Amount")).Value
            'If cnt >= endcell - startcell Then Exit For
        End If
    Next
    
    Set CBA_DBCN = New ADODB.Connection
    CBA_DBCN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=G:\Central_Buying\General\4_Administration\Z - Buying Systems Analyst\LIVE DATABASES\ABI.accdb;"
    For a = LBound(Tdata, 2) To UBound(Tdata, 2)
        If IsEmpty(Tdata(1, a)) = False Then
            Set CBA_DBRS = New ADODB.Recordset
            strSQL = ""
            strSQL = strSQL & "insert into DBU_SAP_Trans([AMT_LCURR],[TXT],[ACC],[CC],[REF_KEY],[SUPP_NAME],[GL_DATE],[COMMENT],[PALREF],[GL_AMT])" & Chr(10)
            strSQL = strSQL & "VALUES (" & Tdata(1, a) & ",'" & Tdata(2, a) & "'," & Tdata(3, a) & "," & Tdata(4, a) & ",'" _
                & Tdata(5, a) & "','" & Tdata(6, a) & "',#" & Format(Tdata(7, a), "MM-DD-YYYY") & "#,'" & Tdata(8, a) & "','" _
                & Tdata(9, a) & "'," & Tdata(10, a) & ")" & Chr(10)
            CBA_DBRS.Open strSQL, CBA_DBCN
        End If
    Next
    
    CBA_DBCN.Close
    Set CBA_DBCN = Nothing

    wbkT.Close

End Sub

