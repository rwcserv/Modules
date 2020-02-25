Attribute VB_Name = "CBAR_PPHistory"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Function CBAR_PPHistoryRuntime()
    Dim lRet As Long
    Dim tR As CBAR_Report
    Dim matchsetup As Boolean
    Dim strProds As String
    Dim wks_PPH As Worksheet, wbk_PPH As Workbook, Chart, arrPPH
    Dim arrPOS As Variant, chartmade, lBline As Long, lBRow As Long, lRCol As Long
    Dim bOutput As Boolean, prod, a As Long, b As Long, bfound As Boolean, lcol As Long, lLine As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Preparing to run the Price & Promotion History Report..."
    DoEvents
    tR = CBAR_Runtime.getActiveReport
    
    If Not tR.AldiProds Is Nothing Then
            For Each prod In tR.AldiProds
                If strProds = "" Then
                    strProds = prod
                Else
                    strProds = strProds & ", " & prod
                End If
            Next
    ElseIf tR.BD <> "" Then
        bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("CBIS_ProdbyEmp")
        If bOutput = True Then
            For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                If InStr(1, CBA_CBISarr(11, a), tR.BD) > 0 Then
                    If strProds = "" Then strProds = CBA_CBISarr(0, a) Else strProds = strProds & ", " & CBA_CBISarr(0, a)
                End If
            Next
            Erase CBA_CBISarr
        Else
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
            MsgBox "There has been an error in querying CBIS" & Chr(10) & Chr(10) & "Please try again later or contact " & g_Get_Dev_Sts("DevUsers")
            Exit Function
        End If
    
    ElseIf tR.GBD <> "" Then
        bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("CBIS_ProdbyEmp")
        If bOutput = True Then
            For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                If InStr(1, CBA_CBISarr(12, a), tR.GBD) > 0 Then If strProds = "" Then strProds = CBA_CBISarr(0, a) Else strProds = strProds & ", " & CBA_CBISarr(0, a)
            Next
            Erase CBA_CBISarr
        Else
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
            MsgBox "There has been an error in querying CBIS" & Chr(10) & Chr(10) & "Please try again later or contact " & g_Get_Dev_Sts("DevUsers")
            Exit Function
        End If
    End If
    matchsetup = CBA_COM_SetupMatchArray.CBA_SetupMatchArray(False, tR.DateFrom, tR.DateTo, tR.CG, tR.scg, strProds, True)
    
    If strProds = "" Then
        Set tR.AldiProds = New Collection
        For a = LBound(CBA_COM_Match, 1) To UBound(CBA_COM_Match, 1)
            bfound = False
            For Each prod In tR.AldiProds
                If prod = CBA_COM_Match(a).AldiPCode Then
                    bfound = True
                    Exit For
                End If
            Next
            If bfound = False Then
                tR.AldiProds.Add CBA_COM_Match(a).AldiPCode
                If strProds = "" Then
                    strProds = CBA_COM_Match(a).AldiPCode
                Else
                    strProds = strProds & ", " & CBA_COM_Match(a).AldiPCode
                End If
            End If
        Next
    End If
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 7, 5, "Querying CBIS Data..."
    
    bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("CBIS_POSQTY", tR.DateFrom, tR.DateTo, , tR.State, , strProds)
    arrPOS = CBA_CBISarr
    Erase CBA_CBISarr
    
    Application.ScreenUpdating = False
    CBAR_PPH.Copy
    Set wks_PPH = ActiveSheet
    Set wbk_PPH = ActiveWorkbook
    With wks_PPH
'            .Visible = xlSheetVisible
'            .Activate
        .Cells.NumberFormat = "General"
        If matchsetup = True Then
            '.Cells.Format =
            .Cells.ClearContents
            .Cells(1, 3).Value = "PRICE & PROMOTION HISTORY REPORT"
            .Cells.UnMerge
            .Cells.RowHeight = 13.5
            .Cells.ColumnWidth = 9.38
            .Cells(1, 1).RowHeight = 63
            .Cells(25, 1).RowHeight = 57.75
            .Cells.Borders.LineStyle = xlNone
            For Each Chart In .ChartObjects
                Chart.Delete
            Next
            'code to create the chart
            If UBound(CBA_COM_Match, 1) > 0 Then
                If UBound(CBA_COM_Match, 1) > 40 Then
                    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 7, 5, "Requesting consent to continue..."
                    Application.ScreenUpdating = True
                    lRet = MsgBox("This query will create " & UBound(CBA_COM_Match, 1) & " graphs. Are you sure you wish to proceed?", vbYesNo)
                    If lRet = 7 Then
                        Application.DisplayAlerts = False
                        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
                        wbk_PPH.Close
                        Application.DisplayAlerts = True
                        Exit Function
                    End If
                    
                End If
                '.Cells(1, 6).Value = "Pricing and Promotion History Report"
                'Range(.Cells(1, 6), .Cells(1, 11)).Merge
            lcol = 1
            lLine = 25
            For a = 1 To UBound(CBA_COM_Match, 1)
            
                arrPPH = CBA_COM_Match(a).RetailsArray
                If UBound(arrPPH, 2) = 1 And arrPPH(1, 1) = 0 Then GoTo NextMatch
                If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 7, 5, "Building chart " & a & " of " & UBound(CBA_COM_Match, 1)
                
                If a > 1 Then
                    lcol = lcol + 6
                    lLine = 25
                End If
                    .Cells(lLine, lcol + 0).Value = CBA_COM_Match(a).Competitor & " Product: " & Chr(10) & CBA_COM_Match(a).CompCode & " - " & CBA_COM_Match(a).CompProdName
                    Range(.Cells(lLine, lcol), .Cells(lLine, lcol + 4)).Merge
                    Range(.Cells(lLine, lcol), .Cells(lLine, lcol + 4)).HorizontalAlignment = xlCenter
                    Range(.Cells(lLine, lcol), .Cells(lLine, lcol + 4)).VerticalAlignment = xlCenter
                    lLine = lLine + 5
                    .Cells(lLine, lcol + 0).Value = "Date"
                    .Cells(lLine, lcol + 1).Value = "Price"
                    .Cells(lLine, lcol + 2).Value = "Normal Price"
                    .Cells(lLine, lcol + 3).Value = "Pricesaving"
                        .Cells(lLine, lcol + 4).Value = CBA_COM_Match(a).AldiPCode & " POS"
                
                For b = LBound(arrPPH, 2) To UBound(arrPPH, 2)
                    If arrPPH(2, b) = LCase(tR.State) Then
                        lLine = lLine + 1
                        .Cells(lLine, lcol).Value = arrPPH(1, b)
                        If CBA_COM_Match(a).AldiPCG = 64 Or CBA_COM_Match(a).AldiPCG = 62 Then
                            .Cells(lLine, lcol + 1).Value = arrPPH(6, b)
                            .Cells(lLine, lcol + 2).Value = arrPPH(4, b) * (arrPPH(6, b) / arrPPH(3, b))
                            .Cells(lLine, lcol + 3).Value = arrPPH(5, b) * (arrPPH(6, b) / arrPPH(3, b))
                        Else
                            .Cells(lLine, lcol + 1).Value = arrPPH(3, b)
                            .Cells(lLine, lcol + 2).Value = arrPPH(4, b)
                            .Cells(lLine, lcol + 3).Value = arrPPH(5, b)
                        End If

                    End If
                Next
                lLine = 30
                bfound = False
                For b = LBound(arrPOS, 2) To UBound(arrPOS, 2)
                
                    If arrPOS(0, b) = CBA_COM_Match(a).AldiPCode Then
                        bfound = True
                        lLine = lLine + 1
                        .Cells(lLine, lcol + 4).Value = arrPOS(2, b)
                        If DateAdd("D", 7, .Cells(lLine, lcol).Value) > Date Then
                            Application.DisplayCommentIndicator = xlCommentIndicatorOnly
                            .Cells(lLine, lcol + 4).Interior.ColorIndex = 22
                            .Cells(lLine, lcol + 4).AddComment "not a full week of POS data avaliable"
                        End If
                    ElseIf bfound = True Then
                        Exit For
                    End If
                Next
                
                
                
                chartmade = CBAR_PPHChartCreate.ChartCreate(.Cells(3, lcol).Left, .Cells(3, lcol).Top, .Range(.Cells(31, lcol + 1), .Cells(lLine, lcol + 4)), wks_PPH, .Cells(25, lcol).Value, Range(.Cells(30, lcol + 1), .Cells(30, lcol + 4)), Range(.Cells(31, lcol), .Cells(lLine, lcol)), xlBottom, xlLine, xlLine, xlColumns, , Range(.Cells(30, lcol), .Cells(30, lcol + 4)).Width / 28.34645669, 10)
                lBline = lLine
                ''Average Price Calculation
                lLine = 26
                .Cells(lLine, lcol).Value = "Average Retail inc Promotion:"
                Range(.Cells(lLine, lcol), .Cells(lLine, lcol + 2 + 1)).Merge
                .Cells(lLine, lcol + 3 + 1).Value = "=AVERAGE(" & Range(.Cells(31, lcol + 1), .Cells(lBline, lcol + 1)).Address & ")"
                .Cells(lLine, lcol + 3 + 1).Value = Format(.Cells(lLine, lcol + 3 + 1).Value, "$0.00")

                ''Average Full Price
                lLine = lLine + 1
                .Cells(lLine, lcol).Value = "Average Full Retail:"
                Range(.Cells(lLine, lcol), .Cells(lLine, lcol + 2 + 1)).Merge
                .Cells(lLine, lcol + 3 + 1).Value = "=AVERAGE(" & Range(.Cells(31, lcol + 2), .Cells(lBline, lcol + 2)).Address & ")"
                .Cells(lLine, lcol + 3 + 1).Value = Format(.Cells(lLine, lcol + 3 + 1).Value, "$0.00")

                ''Average Price when Promoted
                lLine = lLine + 1
                .Cells(lLine, lcol).Value = "Average Promotion Retail:"
                Range(.Cells(lLine, lcol), .Cells(lLine, lcol + 2 + 1)).Merge
                .Cells(lLine, lcol + 3 + 1).Value = "=IFERROR(SUMIF(" & Range(.Cells(31, lcol + 3), .Cells(lBline, lcol + 3)).Address & "," & """<>""" & "&0" & "," & Range(.Cells(31, lcol + 1), .Cells(lBline, lcol + 1)).Address & ")/COUNTIF(" & Range(.Cells(31, lcol + 3), .Cells(lBline, lcol + 3)).Address & "," & """<>""" & "&0" & "),0)"
                .Cells(lLine, lcol + 3 + 1).Value = Format(.Cells(lLine, lcol + 3 + 1).Value, "$0.00")

                ''Percent of weeks on Promotion
                lLine = lLine + 1
                .Cells(lLine, lcol).Value = "Percent of weeks on Promotion:"
                Range(.Cells(lLine, lcol), .Cells(lLine, lcol + 2 + 1)).Merge
                .Cells(lLine, lcol + 3 + 1).Value = "=COUNTIF(" & Range(.Cells(31, lcol + 3), .Cells(lBline, lcol + 3)).Address & "," & """<>""" & "&0" & ")  / COUNT(" & Range(.Cells(31, lcol + 3), .Cells(lBline, lcol + 3)).Address & ")"
                .Cells(lLine, lcol + 3 + 1).Value = Format(.Cells(lLine, lcol + 3 + 1).Value, "0.0%")

                .Range(.Cells(25, lcol), .Cells(29, lcol + 3 + 1)).Borders.LineStyle = xlContinuous

NextMatch:
            Next

            End If
        Else
            Application.DisplayAlerts = False
            wbk_PPH.Close
            Application.DisplayAlerts = True
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
            Application.ScreenUpdating = True
            MsgBox "No active matches were found"
            Exit Function
        End If
    End With

    With wks_PPH.PageSetup
    lBRow = wks_PPH.Cells(9999, 1).End(xlUp).Row
    lRCol = wks_PPH.Cells(lBRow, 9999).End(xlToLeft).Column
    
        If lBRow = 1 And lRCol = 1 Then Else .PrintArea = Range(wks_PPH.Cells(1, 1), wks_PPH.Cells(lBRow, lRCol)).Address
        On Error Resume Next
        .Zoom = False
        .FitToPagesWide = False
        .FitToPagesTall = 1
        On Error GoTo Err_Routine
        Select Case Application.WorksheetFunction.RoundDown(Len(Application.ActiveWorkbook.FullName) / 125, 0)
            Case 0
                .LeftFooter = "&9CORP BUYING, Admin, per: " & Format(Date, "DD/MM/YYYY") & Chr(10) & Application.ActiveWorkbook.FullName
            Case 1
                .LeftFooter = "&9CORP BUYING, Admin, per: " & Format(Date, "DD/MM/YYYY") & Chr(10) & Mid(Application.ActiveWorkbook.FullName, 1, 125) & Chr(10) & Mid(Application.ActiveWorkbook.FullName, 125, 125)
            Case 1
                .LeftFooter = "&9CORP BUYING, Admin, per: " & Format(Date, "DD/MM/YYYY") & Chr(10) & Mid(Application.ActiveWorkbook.FullName, 1, 125) & Chr(10) & Mid(Application.ActiveWorkbook.FullName, 125, 125) & Chr(10) & Mid(Application.ActiveWorkbook.FullName, 250, 999999)
        End Select
        '.LeftFooter = "&9CORP BUYING, Admin, per: " & Format(Date, "DD/MM/YYYY") & Chr(10) & Application.ActiveWorkbook.FullName
        .RightFooter = "&P of &N"
        .Orientation = xlLandscape
        .PrintGridlines = False
    End With

    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    Application.ScreenUpdating = True
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBAR_PPHistoryRuntime", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
