Attribute VB_Name = "CBA_COM_T_BaseMatchData"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Sub CBA_COM_T_BaseMatchDataRun(Control As IRibbonControl)
    Dim PCode As String, lWeeks As Long, wks_Datasht, lIdx As Long, bData As Boolean, bPivot As Boolean
    Dim strPcode As String, strWeeks As String, ErrorCodes As String, wks_Piv, tR As CBA_BTF_ReportParamaters
    Dim dtEDate As Date, dtSDate As Date, lRow As Long, a As Long, b As Long, c As Long, d As Long, arr
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    If CBA_getVersionStatus(g_GetDB("Gen"), CBA_COM_Ver, "Comrade", "COM", True) = "Exit" Then Exit Sub

    strPcode = InputBox("Aldi Product Code - multiple product codes can be entered (separated by a comma)", "")
    If strPcode = "" Then Exit Sub
    arr = Split(strPcode, ",")
    For a = 0 To UBound(arr)
        If IsNumeric(arr(a)) And arr(a) <> "" Then
            If Len(arr(a)) > 7 Or Len(arr(a)) < 4 Then
                If ErrorCodes = "" Then ErrorCodes = arr(a) Else ErrorCodes = ErrorCodes & ", " & arr(a)
                arr(a) = 0
            Else
                If PCode = "" Then PCode = arr(a) Else PCode = PCode & ", " & arr(a)
            End If
        Else
            If ErrorCodes = "" Then ErrorCodes = arr(a) Else ErrorCodes = ErrorCodes & ", " & arr(a)
            arr(a) = 0
        End If
    Next
    If ErrorCodes <> "" Then
        MsgBox ErrorCodes & " are not a valid Product Codes"
        If PCode = "" Then Exit Sub
    End If
    
    strWeeks = InputBox("Weeks of Data (1 to 52):", "")
    If IsNumeric(strWeeks) And strWeeks <> "" Then
        If strWeeks > 52 Or strWeeks = 0 Then
            MsgBox "A number between 1 and 52 is expected"
            Exit Sub
        End If
        lWeeks = CLng(strWeeks)
    Else
        MsgBox "Not a valid Number"
        Exit Sub
    End If
    
    CBA_BasicFunctions.CBA_Running "Loading Data"
    Application.ScreenUpdating = False
'    CBA_WedDate = Date
    If tR.BD = "Produce" Or tR.CG = 58 Then
        dtEDate = DateAdd(d, -1, Date)
    Else
        dtEDate = CBA_COM_Runtime.CBA_getWedDate(CStr(Date))
    End If

'    dtEDate = DateAdd("d", -Weekday(Date, vbThursday), Date)
    dtSDate = DateAdd("d", 1, DateAdd("d", -lWeeks * 7, dtEDate))

    If CBA_COM_SetupMatchArray.CBA_SetupMatchArray(True, dtSDate, dtEDate, , , PCode) = True Then

        Application.Workbooks.Add
        Set wks_Datasht = ActiveSheet
        With wks_Datasht
            lRow = 1
            .Cells(lRow, 1).Value = "AldiProd"
            .Cells(lRow, 2).Value = "AldiPDesc"
            .Cells(lRow, 3).Value = "CG"
            .Cells(lRow, 4).Value = "SCG"
            .Cells(lRow, 5).Value = "Competitor"
            .Cells(lRow, 6).Value = "MatchType"
            .Cells(lRow, 7).Value = "CompCode"
            .Cells(lRow, 8).Value = "CompDesc"
            .Cells(lRow, 9).Value = "CompPackOriginal"
            .Cells(lRow, 10).Value = "CompPack"
            .Cells(lRow, 11).Value = "ScrapedDate"
            .Cells(lRow, 12).Value = "State"
            .Cells(lRow, 13).Value = "ShelfPrice"
            .Cells(lRow, 14).Value = "was"
            .Cells(lRow, 15).Value = "Discount"
            .Cells(lRow, 16).Value = "perMeasure"
            .Cells(lRow, 17).Value = "nonSpecialProRata"
            .Cells(lRow, 18).Value = "ProRata"
            .Cells(lRow, 19).Value = "Special"
            .Cells(lRow, 20).Value = "AldiRetail"
            .Cells(lRow, 21).Value = "diff%"
            .Cells(lRow, 22).Value = "Count"
            For a = LBound(CBA_COM_Match) To UBound(CBA_COM_Match)
                arr = CBA_COM_Match(a).RetailsArray
                For b = LBound(arr, 2) To UBound(arr, 2)
                    lRow = lRow + 1
                    .Cells(lRow, 1).Value = CBA_COM_Match(a).AldiPCode
                    .Cells(lRow, 2).Value = CBA_COM_Match(a).AldiPName
                    .Cells(lRow, 3).Value = CBA_COM_Match(a).AldiPCG
                    .Cells(lRow, 4).Value = CBA_COM_Match(a).AldiPSCG
                    .Cells(lRow, 5).Value = CBA_COM_Match(a).Competitor
                    .Cells(lRow, 6).Value = CBA_COM_Match(a).MatchType
                    .Cells(lRow, 7).Value = CBA_COM_Match(a).CompCode
                    .Cells(lRow, 8).Value = CBA_COM_Match(a).CompProdName
                    .Cells(lRow, 9).Value = CBA_COM_Match(a).CompOriginalPack
                    .Cells(lRow, 10).Value = CBA_COM_Match(a).CompPacksize
                    For c = LBound(arr, 1) To UBound(arr, 1)
                        .Cells(lRow, 10 + c) = arr(c, b)
                    Next
                Next

            Next
            Set wks_Piv = ActiveWorkbook.Sheets.Add
            ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range(wks_Datasht.Cells(1, 1), wks_Datasht.Cells(lRow, 22)), Version:=xlPivotTableVersion14).CreatePivotTable _
                TableDestination:=wks_Piv.Cells(3, 1), TableName:="BaseMatchData" & Format(Date, "YYYY-MM-DD"), DefaultVersion:=xlPivotTableVersion14


            'Range("B7").Select
            With wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields("ScrapedDate")
                .Orientation = xlPageField
                .Position = 1
            End With
            With wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields("State")
                .Orientation = xlPageField
                .Position = 1
            End With
            wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields("State").ClearAllFilters
            wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields("State").CurrentPage = "national"
            With wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields("CompDesc")
                .Orientation = xlRowField
                .Position = 1
            End With
            With wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields("MatchType")
                .Orientation = xlRowField
                .Position = 2
            End With
            wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).AddDataField wks_Piv. _
                PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields("AldiRetail"), _
                "Count of AldiRetail", xlCount
            With wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields("Count of AldiRetail")
                .Caption = "AldiRetail "
                .Function = xlAverage
                .NumberFormat = "$#,##0.00"
            End With
            wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).AddDataField wks_Piv. _
                PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields("nonSpecialProRata"), _
                "Count of nonSpecialProRata", xlCount
            With wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields( _
                "Count of nonSpecialProRata")
                .Caption = "ProRata (excl. Promotion)"
                .Function = xlAverage
                .NumberFormat = "$#,##0.00"
            End With
            wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).AddDataField wks_Piv. _
                PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields("ProRata"), _
                "Count of ProRata", xlCount
            With wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields( _
                "Count of ProRata")
                .Caption = "ProRata "
                .Function = xlAverage
                .NumberFormat = "$#,##0.00"
            End With
            wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).AddDataField wks_Piv. _
                PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields("ShelfPrice"), _
                "Count of ShelfPrice", xlCount
            With wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields("Count of ShelfPrice")
                .Caption = "Shelf "
                .Function = xlAverage
                .NumberFormat = "$#,##0.00"
            End With
            wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).AddDataField wks_Piv. _
                PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields("diff%"), "Count of diff%", _
                xlCount
            With wks_Piv.PivotTables("BaseMatchData" & Format(Date, "YYYY-MM-DD")).PivotFields( _
                "Count of diff%")
                .Caption = "Diff% "
                .Function = xlAverage
                .NumberFormat = "0.0%"
            End With
        End With
        '#RW 191018 Fix Win10 worksheet issue
        For lIdx = 1 To 4
            If g_WorkSheetExists(ActiveWorkbook, "Sheet" & lIdx) Then
                If bData = False Then
                    ActiveWorkbook.Worksheets("Sheet" & lIdx).Name = "Data"
                    bData = True
                ElseIf bPivot = False Then
                    ActiveWorkbook.Worksheets("Sheet" & lIdx).Name = "Pivot"
                    bPivot = True
                Else
                    Call g_WorkSheetDelete(ActiveWorkbook, "Sheet" & lIdx)
                End If
            End If
        Next
        
''        Call g_WorkSheetDelete(ActiveWorkbook.Worksheets, "Sheet1")
        
''        Application.DisplayAlerts = False
''        ActiveWorkbook.Worksheets("Sheet2").Delete
''        ActiveWorkbook.Worksheets("Sheet3").Delete
''        Application.DisplayAlerts = True
''        ActiveWorkbook.Worksheets("Sheet1").Name = "Data"
''        ActiveWorkbook.Worksheets("Sheet4").Name = "Pivot"

    Else
        MsgBox "No Data Found"
    End If


    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running

    Application.ScreenUpdating = True
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_COM_T_BaseMatchDataRun", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

