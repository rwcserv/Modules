Attribute VB_Name = "CBAR_ReportPromoAnalysis"
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

''Private Sub RunMultipleDataSets()
''    Dim sht, cell
''
''    Set sht = ActiveSheet
''    With sht
''
''    For Each cell In .Columns(5).Cells
''        'cell.Select
''        If cell.row > 105 Then Exit For
''        If cell.Value <> "" Then
''
''            CBAR_Runtime.createActiveReport "Promotional Analysis Report"
''            CBAR_Runtime.setActiveReportParamater "datefrom", , , Format(cell.Offset(0, -3).Value, "YYYY-MM-DD")
''            CBAR_Runtime.setActiveReportParamater "dateto", , , Format(cell.Offset(0, -2).Value, "YYYY-MM-DD")
''
''            CBAR_Runtime.setActiveReportParamater "state", "National"
''            CBAR_Runtime.setActiveReportParamater "competitor", "ALL Competitors"
''
''            Call CBAR_PAREport
''
''        End If
''
''    Next
''
''    End With
''
''End Sub
''



Sub CBAR_PAREport()
    Dim tR As CBAR_Report
    Dim DFrom As Date, Dto As Date, a As Long, b As Long, c As Long
    Dim sStates As String, wks_PA, wbk, wks_PAsum, prod, ThisProd
    Dim CGs As Long, SCGs As Long, thisPclass As Long, rets, rng
    Dim prodlist As String, Competitors, wkdiff, bOutput As Boolean
    Dim thisbuyer
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    tR = CBAR_Runtime.getActiveReport
    DFrom = tR.DateFrom
    Dto = tR.DateTo
    sStates = tR.State
    CGs = tR.CG
    SCGs = tR.scg
    Competitors = tR.Competitor
       
    CBAR_PA.Copy
    Set wks_PA = ActiveSheet
    Set wbk = ActiveWorkbook
    CBAR_PAsum.Copy
    ActiveSheet.Move After:=wbk.Sheets(1)
    Set wks_PAsum = ActiveSheet
    wks_PA.Activate
    
    
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Preparing to run Promotional Analysis Report..."
    'Application.ScreenUpdating = False
    DoEvents
    
    'dto = "2018-01-31"
    'dfrom = DateAdd("M", -4, DateSerial(Year(dto), Month(dto) + 1, 1))
    If WeekDay(DFrom, vbWednesday) > 1 Then DFrom = DateAdd("D", 8 - WeekDay(DFrom, vbWednesday), DFrom)
    If WeekDay(Dto, vbWednesday) > 1 Then Dto = DateAdd("D", -WeekDay(Dto, vbWednesday) + 1, Dto)
    
    wkdiff = DateDiff("WW", DFrom, Dto, vbWednesday)
    
    wks_PA.Cells(2, 100).Value = wkdiff + 1
    
    If tR.GBD <> "" Then
        bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("CBIS_ProdbyEmp")
        If bOutput = True Then
            For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                If InStr(1, CBA_CBISarr(12, a), tR.GBD) > 0 Then If prodlist = "" Then prodlist = CBA_CBISarr(0, a) Else prodlist = prodlist & ", " & CBA_CBISarr(0, a)
            Next
        Else
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
            MsgBox "There has been an error in querying CBIS" & Chr(10) & Chr(10) & "Please try again later or contact " & g_Get_Dev_Sts("DevUsers")
            Exit Sub
        End If
    End If
    If tR.BD <> "" Then
        bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("CBIS_ProdbyEmp")
        If bOutput = True Then
            For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                If InStr(1, CBA_CBISarr(11, a), tR.BD) > 0 Then
                    If prodlist = "" Then prodlist = CBA_CBISarr(0, a) Else prodlist = prodlist & ", " & CBA_CBISarr(0, a)
                End If
            Next
        Else
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
            MsgBox "There has been an error in querying CBIS" & Chr(10) & Chr(10) & "Please try again later or contact " & g_Get_Dev_Sts("DevUsers")
            Exit Sub
        End If
    End If
    
    If tR.AldiProds Is Nothing Then
    Else
                For Each prod In tR.AldiProds
                    If prodlist = "" Then prodlist = prod Else prodlist = prodlist & ", " & prod
                Next
    End If
    
    
    
    'If mod_SetupProdArray.SetupMatchArray(True, dfrom, dto, 51) = True Then
    If CBA_COM_SetupMatchArray.CBA_SetupMatchArray(True, DFrom, Dto, CGs, SCGs, prodlist, True, tR.Competitor) = True Then
    
    If wks_PA.FilterMode = True Then wks_PA.ShowAllData
    Range(wks_PA.Cells(6, 1), wks_PA.Cells(99999, 20)).ClearContents
    c = 5
    For a = LBound(CBA_COM_Match, 1) To UBound(CBA_COM_Match, 1)
        On Error Resume Next
        If rets = Empty Then
        Else
            Erase rets
        End If
        On Error GoTo Err_Routine
        
        If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 7, 5, "Resolving Match for Aldi Productcode: " & CBA_COM_Match(a).AldiPCode & " - " & CBA_COM_Match(a).AldiPName
        If InStr(1, CBA_COM_Match(a).MatchType, "ML") > 0 Or InStr(1, CBA_COM_Match(a).MatchType, "CB") > 0 Or InStr(1, CBA_COM_Match(a).MatchType, "PB") > 0 Then
            rets = CBA_COM_Match(a).RetailsArray
            For b = LBound(rets, 2) To UBound(rets, 2)
                'If (rets(5, b) > 0 Or rets(9, b) = True) And Mid(rets(11, b), 1, 1) = "-" Then
                If rets(9, b) = True Then
                If Mid(rets(11, b), 1, 1) = "-" Then
                    
                    If ThisProd <> CBA_COM_Match(a).AldiPCode Then
                        ThisProd = CBA_COM_Match(a).AldiPCode
                        CBA_DBtoQuery = 599
                        bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("CBIS_BuyerInfo", , , , , , CStr(ThisProd))
                        If bOutput = True Then thisbuyer = CBA_CBISarr(0, 0) Else thisbuyer = ""
                        bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("CBIS_Pclass", , , , , , CStr(ThisProd))
                        If bOutput = True Then thisPclass = CBA_CBISarr(0, 0) Else thisPclass = 0
                    End If
                    If thisPclass = 1 Then
                        c = c + 1
                        wks_PA.Cells(c, 1).Value = CBA_COM_Match(a).AldiPCode
                        wks_PA.Cells(c, 2).Value = CBA_COM_Match(a).AldiPName
                        wks_PA.Cells(c, 3).Value = CBA_COM_Match(a).AldiPCG
                        wks_PA.Cells(c, 4).Value = CBA_COM_Match(a).AldiPSCG
                        wks_PA.Cells(c, 5).Value = CBA_COM_Match(a).CompMultby & LCase(CBA_COM_Match(a).HowComp)
                        'wks_PA.Cells(c, 6).Value = CBA_COM_Match(a).HowComp
                        wks_PA.Cells(c, 6).Value = rets(10, b)
                        wks_PA.Cells(c, 7).Value = CBA_COM_Match(a).Competitor
                        wks_PA.Cells(c, 8).Value = CBA_COM_Match(a).CompCode
                        wks_PA.Cells(c, 9).Value = CBA_COM_Match(a).CompProdName
                        wks_PA.Cells(c, 10).Value = CBA_COM_Match(a).MatchType
                        wks_PA.Cells(c, 11).Value = CBA_COM_Match(a).CompDivideby & LCase(CBA_COM_Match(a).HowComp)
                        wks_PA.Cells(c, 12).Value = rets(3, b)
                        wks_PA.Cells(c, 13).Value = rets(8, b)
                        wks_PA.Cells(c, 14).Value = Format(rets(11, b), "#,0.00%")
                        wks_PA.Cells(c, 15).Value = rets(2, b)
                        wks_PA.Cells(c, 16).Value = rets(1, b)
                        wks_PA.Cells(c, 17).Value = thisbuyer
                    End If
                End If
                End If
            Next
            'wks_output.Cells(b, 1).Value = ""
        End If
    Next
        
    wks_PA.Cells(1, 100).Value = c
    If LCase(tR.State) = "national" Then
        wks_PAsum.Cells(2, 12).Value = "National"
    Else
        wks_PAsum.Cells(2, 12).Value = UCase(tR.State)
    End If
    
    wks_PAsum.Cells(2, 7).Value = Format(DFrom, "DD/MM/YYYY") & " - " & Format(Dto, "DD/MM/YYYY")
    Set rng = Range(wks_PAsum.Cells(6, 1), wks_PAsum.Cells(6, 1).End(xlToRight).End(xlDown))
    rng.Sort rng.Columns(13), xlDescending, , , , , , xlYes
    rng.AutoFilter
    wks_PAsum.Activate
    
    Else
        Application.DisplayAlerts = False
        wbk.Close
        Application.DisplayAlerts = True
    End If
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    Application.ScreenUpdating = True
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBAR_PAREport", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Sub

