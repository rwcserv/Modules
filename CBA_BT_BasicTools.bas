Attribute VB_Name = "CBA_BT_BasicTools"
Option Explicit
Option Private Module          ' Excel users cannot access procedures

Private CBA_BT_ICcursht As Worksheet
Private CBA_BT_ICcurwbk As Workbook

Sub CBA_FXCalculator(Control As IRibbonControl)
    Dim user As String, AdminUser As Boolean
    user = Application.UserName
    If InStr(1, LCase(user), "lentini, david") > 0 Or InStr(1, LCase(user), "pearce, tom") > 0 Or InStr(1, LCase(user), "hollier, moira") > 0 Or InStr(1, LCase(user), "white, robert") > 0 _
        Or InStr(1, LCase(user), "collett, sarah") > 0 Or InStr(1, LCase(user), "venkatesan , sang") > 0 Or InStr(1, LCase(user), "woods, mark") > 0 _
            Or InStr(1, LCase(user), "sanders, justine") > 0 Or InStr(1, LCase(user), "hillsmith, caroline") > 0 Or InStr(1, LCase(user), "baines, stuart") > 0 Then
        AdminUser = True
    Else
        AdminUser = False
    End If
    If AdminUser = True Then
        If MsgBox("Run FX Monthly Report?", vbYesNo) = 6 Then
'            CBA_FX.RunMonthlyReport
            MsgBox "Not setup yet"
        Else
            CBA_FX_PortfolioSelector.Show vbModeless
        End If
    Else
        CBA_FX_PortfolioSelector.Show vbModeless
    End If

End Sub

Sub CBA_ApprovedLDT(Control As IRibbonControl)
    CBA_BT_frm_ALDT.Show vbModeless
End Sub

Sub CBA_fixInsightCentreExcelExport(Control As IRibbonControl)
    Dim bRun As Boolean, wbk, lRet As Long
    On Error GoTo Err_Routine
    
    bRun = False: CBA_ErrTag = "Wbk"
    For Each wbk In Application.Workbooks
        If Windows(wbk.Name).Visible = True Then
            bRun = True
            Exit For
        End If
    Next
FileNameHandler:     ' #SB - Added to deal with differing filenames upon excel recovery
    CBA_ErrTag = ""
    If bRun = True Then
        Set CBA_BT_ICcurwbk = ActiveWorkbook
        If CBA_BT_ICcurwbk Is Nothing Then
            MsgBox "Please have a workbook open and selected on the sheet you wish to correct formatting for"
            Exit Sub
        End If
        Set CBA_BT_ICcursht = ActiveWorkbook.ActiveSheet
        lRet = MsgBox("Do you want to fix the Sheet named: " & CBA_BT_ICcursht.Name & " in the workbook: " & CBA_BT_ICcurwbk.Name, vbYesNo)
        If lRet = 6 Then CBA_BT_ICFix.Show vbModeless Else MsgBox "Please make the sheet you wish to correct the selected sheet"
    Else
        MsgBox "No Active Workbook"
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    If Err.Number = 9 And CBA_ErrTag = "Wbk" Then bRun = True: GoTo FileNameHandler ' #SB - Added to deal with differing filenames upon excel recovery
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_fixInsightCentreExcelExport", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Function CBA_ICFix(ByRef valtobtn() As Long)
    Dim a As Long, b As Long, c As Long, bfound As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    With CBA_BT_ICcursht
        .Activate
        .Cells.UnMerge
        c = 0
        For a = 1 To valtobtn(2)
            bfound = False
            For b = valtobtn(1) + 1 To valtobtn(3)
                If .Cells(b, a).Value <> "" Then
                    bfound = True
                    Exit For
                End If
            Next
            If bfound = False Then
                    CBA_BT_ICcursht.Columns(a).Delete
                    a = a - 1
                    c = c + 1
            End If
            If c + a > valtobtn(2) Then Exit For
        Next
        Cells.EntireColumn.AutoFit
        On Error Resume Next
        Range(.Cells(valtobtn(1), 1), .Cells(valtobtn(1), valtobtn(2) - c + 1)).AutoFilter
    End With
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_ICFix", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Sub CBA_BT_UnrealisedRevenue(Control As IRibbonControl)
    Dim a As Long, b As Long, wks_URR, lRow As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    CBA_BasicFunctions.CBA_Running ("Querying CBIS...")
    Application.ScreenUpdating = False
    CBA_SQL_Queries.CBA_GenPullSQL "CBIS_UnrealisedRevenue", CBA_POSQuery.getCBA_OSDSelected, , CBA_POSQuery.CBA_getPOSProductcode
    
    If CBA_CBISarr(0, 0) = 0 Then
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
        Application.ScreenUpdating = True
        MsgBox "No Data Returned. Please check Productcode and OSD"
        Exit Sub
    End If
    
    Workbooks.Add
    Set wks_URR = ActiveWorkbook.Worksheets(1)
    
    With wks_URR
        Range(.Cells(1, 1), .Cells(4, 49)).Interior.ColorIndex = 49
        .Cells(1, 1).Select
        .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
        
        .Cells.Font.Name = "ALDI SUED Office"
        .Cells(2, 3).Font.Size = 24
        .Cells(2, 3).Font.ColorIndex = 2
        .Cells(5, 1).Value = "Region"
        .Cells(5, 2).Value = "RCV QTY"
        .Cells(5, 3).Value = "POS 3"
        .Cells(5, 4).Value = "SellT% 3"
        .Cells(5, 5).Value = "POS 7"
        .Cells(5, 6).Value = "SellT% 7"
        .Cells(5, 7).Value = "POS 14"
        .Cells(5, 8).Value = "SellT% 14"
        .Cells(5, 9).Value = "POS 30"
        .Cells(5, 10).Value = "SellT% 30"
        .Cells(5, 11).Value = "POS 60"
        .Cells(5, 12).Value = "SellT% 60"
        .Cells(5, 13).Value = "POS 90"
        .Cells(5, 14).Value = "SellT% 90"
        .Cells(5, 15).Value = "POS 180"
        .Cells(5, 16).Value = "SellT% 180"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 1).Value = "Region"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 2).Value = "RCV Retail"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 3).Value = "URR 3"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 4).Value = "Cont$ 3"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 5).Value = "URR 7"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 6).Value = "Cont$ 7"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 7).Value = "URR 14"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 8).Value = "Cont$ 14"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 9).Value = "URR 30"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 10).Value = "Cont$ 30"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 11).Value = "URR 60"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 12).Value = "Cont$ 60"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 13).Value = "URR 90"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 14).Value = "Cont$ 90"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 15).Value = "URR 180"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 16).Value = "Cont$ 180"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 17).Value = "Gross Margin"
        .Cells(UBound(CBA_CBISarr, 2) + 8, 18).Value = "Net Margin"
    
        lRow = 5
        For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
    '        If a > 0 Then If CBA_CBISarr(0, a) < CBA_CBISarr(0, a - 1) Then lRow = lRow + UBound(CBA_CBISarr, 2) + 1
            lRow = lRow + 1
            For b = 0 To 15
                If b = 0 Then
                    .Cells(lRow, b + 1).Value = CBA_BasicFunctions.CBA_DivtoReg(CBA_CBISarr(b, a))
                Else
                    .Cells(lRow, b + 1).Value = CBA_CBISarr(b, a)
                End If
            Next
            For b = 16 To UBound(CBA_CBISarr, 1)
                If b = 16 Then
                    .Cells(lRow + UBound(CBA_CBISarr, 2) + 3, b - 15).Value = CBA_BasicFunctions.CBA_DivtoReg(CBA_CBISarr(b, a))
                ElseIf b >= UBound(CBA_CBISarr, 1) Then
                    If CBA_CBISarr(b, a) < 0 Then .Cells(lRow + UBound(CBA_CBISarr, 2) + 3, b - 15).Value = "-" Else .Cells(lRow + UBound(CBA_CBISarr, 2) + 3, b - 15).Value = CBA_CBISarr(b, a)
                Else
                    .Cells(lRow + UBound(CBA_CBISarr, 2) + 3, b - 15).Value = CBA_CBISarr(b, a)
                End If
            Next
        Next
    
        For a = 4 To 16 Step 2
            Range(.Cells(6, a), .Cells(6 + UBound(CBA_CBISarr, 2), a)).NumberFormat = "0.0%"
        Next
        Range(.Cells(6, 2), .Cells(6 + UBound(CBA_CBISarr, 2), 2)).NumberFormat = "#,0"
        For a = 3 To 15 Step 2
            Range(.Cells(6, a), .Cells(6 + UBound(CBA_CBISarr, 2), a)).NumberFormat = "#,0"
        Next
        For a = 2 To 18
            Select Case a
                Case 2 To 16
                    Range(.Cells(UBound(CBA_CBISarr, 2) + 9, a), .Cells((UBound(CBA_CBISarr, 2) * 2) + 9, a)).NumberFormat = "$#,0.00"
                Case 17, 18
                    Range(.Cells(UBound(CBA_CBISarr, 2) + 9, a), .Cells((UBound(CBA_CBISarr, 2) * 2) + 9, a)).NumberFormat = "0.0%"
            End Select
        Next
        Range(.Cells(5, 1), .Cells(5, 12)).EntireColumn.AutoFit
        
        CBA_SQL_Queries.CBA_GenPullSQL "CBIS_ProductDesc", , , CBA_POSQuery.CBA_getPOSProductcode
        If CBA_CBISarr(0, 0) = 0 Then .Cells(2, 3).Value = "Unrealised Reveune Report" Else .Cells(2, 3).Value = "Unrealised Reveune Report: " & CBA_CBISarr(0, 0)
    
    
    End With
    
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    Application.ScreenUpdating = True
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_BT_UnrealisedRevenue", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Public Function CBA_BT_FwdBwd(ByVal sDate As String, ByRef lFwdBwd As Long) As Boolean
    ' This routine will deliver back a true or false as to whether the year can be moved foreward or backward
    Dim strSQL As String, sReturn As String, b As Long, c As Long, bReForecast As Boolean, dtDate As Date
    CBA_BT_FwdBwd = False
    dtDate = CDate(g_FixDate(sDate))
    dtDate = DateAdd("m", 12 * lFwdBwd, dtDate)
    sReturn = CBA_BT_getCutOffDate(dtDate)
    If sReturn = "Format" Or sReturn = "ReFormat" Or sReturn = "NoSave" Then
        CBA_BT_FwdBwd = True
    End If
End Function

Public Function CBA_BT_getCutOffDate(ByVal dtForecast As Date, Optional dtCurrent As Date = "00:00") As String
    ' Get the Forecasting Cutoff dates into an array, and deliver a status back as required
    Static aDates(), bInit As Boolean
    Dim lRow As Long, sSQL As String, bPassOK As Boolean, lYear As Long, lCYear As Long, dtCapDate As Date, bIsData As Boolean
    ' Produces the following in 2019:-
'''        Print CBA_BT_getCutOffDate("02/01/2020")=Format
'''        Print CBA_BT_getCutOffDate("02/12/2019")=ReFormat
'''        Print CBA_BT_getCutOffDate("01/11/2019")=ReFormat
'''        Print CBA_BT_getCutOffDate("01/10/2019")=ReFormat
'''        Print CBA_BT_getCutOffDate("01/09/2019")=ReFormat
'''        Print CBA_BT_getCutOffDate("01/07/2019")=ReFormat
'''        Print CBA_BT_getCutOffDate("01/01/2018")=NoData

    If g_IsDate(dtCurrent) = False Then dtCurrent = Date
    ' Capture the current table values, when first used and store in an array...
    If bInit = False Then
        CBA_DBtoQuery = 1
        sSQL = "SELECT * FROM CutOffDates " & _
                "WHERE CO_Year > 0 ORDER BY Co_Year;"
        bPassOK = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CutOffDates", g_GetDB("ForeCast"), CBA_MSAccess, sSQL, 120, , , False)
        If bPassOK = True Then
            aDates = CBA_ABIarr
        End If
        CBA_DBtoQuery = 3
        bInit = True
    End If
    ' Return the correct cut off date text
    CBA_BT_getCutOffDate = "": lYear = Year(dtForecast): lCYear = Year(dtCurrent): dtCapDate = "00:00": bIsData = False
    For lRow = 0 To UBound(aDates, 2)
        ' Find out where the date is in the array...
        If dtCurrent < aDates(1, lRow) Then
            dtCapDate = dtCurrent
        ElseIf dtCurrent >= aDates(1, lRow) And dtCurrent < aDates(2, lRow) Then
            dtCapDate = dtCurrent
        ElseIf dtCurrent >= aDates(2, lRow) And lYear < lCYear Then
            bIsData = True
        End If
        ' If the date has been found...
        If g_IsDate(dtCapDate) = True Then
            If lYear + 1 = lCYear And bIsData = True Then
                CBA_BT_getCutOffDate = "NoSave"
            ElseIf lYear = lCYear And dtCurrent < aDates(1, lRow) Then
                CBA_BT_getCutOffDate = "Format"
            ElseIf lYear = lCYear And dtCurrent < aDates(2, lRow) Then
                CBA_BT_getCutOffDate = "ReFormat"
            ElseIf lYear = lCYear + 1 And dtCurrent < aDates(1, lRow) Then
                CBA_BT_getCutOffDate = "Format"
            ElseIf lYear = lCYear + 2 And dtCurrent < aDates(2, lRow) Then
                CBA_BT_getCutOffDate = "Format"
            End If
        End If
        If CBA_BT_getCutOffDate > "" Then GoTo Exit_Routine
    Next
    ' If no match has been found then either we are b4 2018 or we are more than a year ahead
    CBA_BT_getCutOffDate = "NoData"
Exit_Routine:
''    Debug.Print CBA_BT_getCutOffDate
    Exit Function

End Function

Public Function CBA_BT_FmtCellVals(ByRef cCell As Range, ByVal vValIn, Optional sFmt As String = "0.0%") As String
    ' Will format a cell as per the value that is going into it...
    If vValIn >= 9.99 Then
        cCell.Value = "999%"
        cCell.HorizontalAlignment = xlRight
        cCell.Interior.ColorIndex = 3
    Else
        cCell.Value = Format(vValIn, sFmt)
    End If

End Function

Public Sub CBA_BT_CalendarShow(MeTop As Long, MeLeft As Long, cActCtl As Control, sDateFmt As String, _
                                Optional bAllowNullDate As Boolean = True, Optional bTakeDateFromCtl As Boolean = True)
    ' Use this to show and process the results from the CBA_frmCalendar
    Dim lOldClr As Long, vActCtlVal
    If bTakeDateFromCtl Then varCal.sDate = g_FixDate(cActCtl.Value & "")
    lOldClr = cActCtl.BackColor
    cActCtl.BackColor = CBA_Pink
    vActCtlVal = varCal.sDate
    varCal.bAllowNullOfDate = bAllowNullDate
    If (MeTop + cActCtl.Top) > CBA_CalHeight Then
        varCal.lCalTop = MeTop + cActCtl.Top '- CBA_CalHeight
    ElseIf (MeTop + cActCtl.Top) > 400 Then
        varCal.lCalTop = MeTop + cActCtl.Top '- CBA_CalHeight
    Else
        varCal.lCalTop = MeTop + cActCtl.Top + cActCtl.Height
    End If
'    varCal.lCalLeft = cActCtl.Left - cActCtl.Width ''+ cActCtl.Left  ''' + (cactctl.Width / 2)
    If cActCtl.Left - CBA_CalWidth < CBA_CalWidth Then
        varCal.lCalLeft = (MeLeft + cActCtl.Left + cActCtl.Width) + CBA_CalWidth
    Else
        varCal.lCalLeft = (MeLeft + cActCtl.Left) - CBA_CalWidth
    End If
    CBA_frmCalendar.Show vbModal
    If varCal.bCalValReturned = True Then
        If g_IsDate(varCal.sDate, True) = True Then
            cActCtl.Value = g_FixDate(varCal.sDate, sDateFmt)
        Else
            cActCtl.Value = varCal.sDate
        End If
    End If
    cActCtl.BackColor = lOldClr
''    varcal.bCalValReturned = False
    DoEvents
End Sub

