Attribute VB_Name = "CBAR_ActiveMatch"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Function AMR()
    Dim tR As CBAR_Report, bOutput As Boolean, lNum As Long, strProds As String, a As Long, b As Long, DFrom, Dto, prod
    Dim wks_AMR, thisAldiPrice, rets, Lowest, Highest, Avg, Avgcnt, thismoveline
    Dim maybedto As Scripting.Dictionary, produce As Scripting.Dictionary, thisDic As Scripting.Dictionary
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    

    
    tR = CBAR_Runtime.getActiveReport
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Preparing to run 'Active Match Report'"
    
    
    strProds = ""
    If tR.BD <> "" Then
        bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("CBIS_ProdbyEmpActive")
        If bOutput = True Then
            For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                If InStr(1, CBA_CBISarr(11, a), tR.BD) > 0 Then
                    If strProds = "" Then strProds = CBA_CBISarr(0, a) Else strProds = strProds & ", " & CBA_CBISarr(0, a)
                End If
            Next
        Else
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
            MsgBox "There has been an error in querying CBIS" & Chr(10) & Chr(10) & "Please try again later or contact " & g_Get_Dev_Sts("DevUsers")
            Exit Function
        End If
    End If
    If tR.GBD <> "" Then
        bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("CBIS_ProdbyEmpActive")
        If bOutput = True Then
            For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                If InStr(1, CBA_CBISarr(12, a), tR.GBD) > 0 Then If strProds = "" Then strProds = CBA_CBISarr(0, a) Else strProds = strProds & ", " & CBA_CBISarr(0, a)
            Next
        Else
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
            MsgBox "There has been an error in querying CBIS" & Chr(10) & Chr(10) & "Please try again later or contact " & g_Get_Dev_Sts("DevUsers")
            Exit Function
        End If
    End If
    If tR.AldiProds Is Nothing Then
    Else
                For Each prod In tR.AldiProds
                    If strProds = "" Then strProds = prod Else strProds = strProds & ", " & prod
                Next
    End If
    
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Preparing to run 'Active Match Report'"
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 7, 4, "Checking Dates..."
    bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("COM_2ScrapeDates")
    If bOutput = False Then
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
        MsgBox "There has been an error in querying COMRADE" & Chr(10) & Chr(10) & "Please try again later or contact " & g_Get_Dev_Sts("DevUsers")
        Exit Function
    Else
        Set maybedto = New Scripting.Dictionary
        Set produce = New Scripting.Dictionary
    End If
    
    
    
    'Looking to find the accurate date to (dto) value. Produce can look at any date, other products only wednesdays.
    'Also identifies where there is more than 100000 values against a date (i.e has a full [non produce] dataset )
    For a = LBound(CBA_COMarr, 2) To UBound(CBA_COMarr, 2)
        If a = LBound(CBA_COMarr, 2) Then
            DFrom = CBA_COMarr(1, a)
            Dto = CBA_COMarr(1, a)
            maybedto.Add CBA_COMarr(0, a), CBA_COMarr(1, a): produce.Add CBA_COMarr(0, a), CBA_COMarr(1, a)
        Else
            If CBA_COMarr(1, a) < DFrom Then DFrom = CBA_COMarr(1, a)
            If CBA_COMarr(1, a) >= produce(CBA_COMarr(0, a)) Or CBA_COMarr(1, a) >= maybedto(CBA_COMarr(0, a)) Then ' And Weekday(CBA_COMarr(1, a), vbWednesday) = 1 Then
                If CBA_COMarr(2, a) > 100000 Then
                    If CBA_COMarr(0, a) = "DM" Or CBA_COMarr(0, a) = "FC" Then
                    Else
                        If produce.Exists(CBA_COMarr(0, a)) = True Then produce.Remove CBA_COMarr(0, a)
                        produce.Add CBA_COMarr(0, a), CBA_COMarr(1, a)
                    End If
                    If maybedto.Exists(CBA_COMarr(0, a)) = True Then maybedto.Remove CBA_COMarr(0, a)
                    maybedto.Add CBA_COMarr(0, a), CBA_COMarr(1, a)
                Else
                    If produce.Exists(CBA_COMarr(0, a)) = True Then produce.Remove CBA_COMarr(0, a): produce.Add CBA_COMarr(0, a), CBA_COMarr(1, a)
                End If
            End If
        End If
    Next
    
    If tR.BD = "Produce" Or tR.CG = 58 Then
        If DateDiff("D", produce("C"), produce("WW")) > 0 Then Dto = produce("WW") Else Dto = produce("C")
        DFrom = DateAdd("D", -8, Dto)
    Else
        Dto = CBA_COM_Runtime.CBA_getWedDate
        'If Weekday(dfrom, vbWednesday) > 1 Then dfrom = DateAdd("D", -(Weekday(dfrom, vbWednesday) - 1), dfrom)
        DFrom = DateAdd("D", -8, Dto)
    End If
    
'    If CBA_COM_SetupMatchArray.CBA_SetupMatchArray(tR.Matchhistory, dfrom, dto, tR.CG, tR.SCG, strProds) = True Then
    If CBA_COM_SetupMatchArray.CBA_SetupMatchArray(False, DFrom, Dto, tR.CG, tR.scg, strProds) = True Then
        
        'Application.ScreenUpdating = False
        CBAR_AMR.Copy
        Set wks_AMR = ActiveSheet
        If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.RunningSheetAddComment 7, 4, "Building Report..."
        Application.ScreenUpdating = True
        DoEvents
        With wks_AMR
            
            Range(.Cells(6, 1), .Cells(99999, 15)).ClearContents
            lNum = 5
            For a = LBound(CBA_COM_Match) To UBound(CBA_COM_Match)
                If a = 6 Then
                a = a
                End If
                If (CBA_COM_Match(a).AldiPCG <> 58 And CBA_COM_Match(a).CompProdName <> "") _
                    Or (CBA_COM_Match(a).AldiPCG = 58 And CBA_COM_Match(a).CompProdName <> "") Then
                    Set thisDic = New Scripting.Dictionary
                    If CBA_COM_Match(a).AldiPCG = 58 Then Set thisDic = produce Else Set thisDic = maybedto
                    lNum = lNum + 1
                    .Cells(lNum, 1).Value = CBA_COM_Match(a).AldiPCG
                    .Cells(lNum, 2).Value = CBA_COM_Match(a).AldiPSCG
                    .Cells(lNum, 3).Value = CBA_COM_Match(a).AldiPCode
                    .Cells(lNum, 4).Value = CBA_COM_Match(a).AldiPName
                    .Cells(lNum, 5).Value = CBA_COM_Match(a).Competitor
                    .Cells(lNum, 6).Value = CBA_COM_Match(a).MatchType
                    .Cells(lNum, 7).Value = CBA_COM_Match(a).CompProdName
                    .Cells(lNum, 8).Value = CBA_COM_Match(a).CompDivideby & LCase(CBA_COM_Match(a).HowComp)
                    .Cells(lNum, 9).Value = CBA_COM_Match(a).CompMultby & LCase(CBA_COM_Match(a).HowComp)
                    .Cells(lNum, 10).Value = CBA_COM_Match(a).CompCode
                    ' #RW Start 191113 - then following NZs to handle states...
                    thisAldiPrice = NZ(CBA_COM_Match(a).Pricedata(thisDic(CBA_COM_Match(a).compet), "AldiRetail", tR.State), 0)
                    .Cells(lNum, 11).Value = thisAldiPrice
                    .Cells(lNum, 12).Value = NZ(CBA_COM_Match(a).Pricedata(thisDic(CBA_COM_Match(a).compet), "nonpromoprorata", tR.State), 0)
                    If .Cells(lNum, 12).Value = 0 Then
                        .Cells(lNum, 13).Value = 0
                    Else
'                        .Cells(lNum, 13).Value = (.Cells(lNum, 12).Value - .Cells(lNum, 11).Value) / .Cells(lNum, 11).Value
                        .Cells(lNum, 13).Value = (.Cells(lNum, 12).Value - .Cells(lNum, 11).Value) / .Cells(lNum, 12).Value
                    End If
                    '.Cells(lNum, 13).Value = CBA_COM_Match(a).Pricedata(thisDic(CBA_COM_Match(a).compet), "delta", tR.State)
                    
                    
                    If .Cells(lNum, 11).Value <> 0 Then .Cells(lNum, 13).Interior.ColorIndex = CBA_COM_ReportFunctions.CBA_COM_isTrafficLight(.Cells(lNum, 13).Value, CBA_COM_Match(a).MatchType)
                    If CBA_COM_Match(a).Pricedata(thisDic(CBA_COM_Match(a).compet), "isspecial", tR.State) = True Then
                        .Cells(lNum, 14).Value = "TRUE"
                    ElseIf NZ(CBA_COM_Match(a).Pricedata(thisDic(CBA_COM_Match(a).compet), "discount", tR.State), 0) > 0 Then
                        .Cells(lNum, 14).Value = "TRUE"
                    Else
                        .Cells(lNum, 14).Value = "FALSE"
                    End If
                    
                    .Cells(lNum, 15).Value = NZ(CBA_COM_Match(a).Pricedata(thisDic(CBA_COM_Match(a).compet), "Shelf", tR.State), 0)
                    .Cells(lNum, 16).Value = NZ(CBA_COM_Match(a).Pricedata(DateAdd("WW", -1, thisDic(CBA_COM_Match(a).compet)), "Shelf", tR.State), 0)
                    '.Cells(lNum, 13).Value = CBA_COM_Match(a).CompPacksize
                    '.Cells(lNum, 14).Value = CBA_COM_Match(a).Pricedata(thisDic(CBA_COM_Match(a).compet), "delta", tR.State)
                    rets = CBA_COM_Match(a).RetailHistoryArray("prorata")
                    Lowest = 0: Highest = 0: Avg = 0: Avgcnt = 0
                    For b = LBound(rets, 2) To UBound(rets, 2)
                        If rets(2, b) = LCase(tR.State) Then
                            If Lowest = 0 Or Lowest > rets(3, b) Then Lowest = rets(3, b)
                            If Highest = 0 Or Highest < rets(3, b) Then Highest = rets(3, b)
                            Avg = Avg + rets(3, b)
                            Avgcnt = Avgcnt + 1
                        End If
                    Next
                    .Cells(lNum, 17).Value = Lowest
                    .Cells(lNum, 18).Value = Highest
                    If Avgcnt > 0 Then
                        .Cells(lNum, 19).Value = Avg / Avgcnt
                        thismoveline = MoveRetailintoLine(CBA_COM_Match(a).MatchType, Avg / Avgcnt)
                    Else
                        .Cells(lNum, 19).Value = 0
                    End If
                    ' #RW End 191113
                    If thismoveline <> 0 Then
                        .Cells(lNum, 20).Value = thismoveline
                        If thisAldiPrice = 0 Then .Cells(lNum, 21).Value = 0 Else .Cells(lNum, 21).Value = (thismoveline - thisAldiPrice) / thisAldiPrice
                    Else
                        .Cells(lNum, 20).Value = "N/A"
                        .Cells(lNum, 21).Value = "N/A"
                    End If
                    thismoveline = MoveRetailintoLine(CBA_COM_Match(a).MatchType, CBA_COM_Match(a).Pricedata(thisDic(CBA_COM_Match(a).compet), "nonpromoprorata", tR.State))
                    If thismoveline <> 0 Then
                        .Cells(lNum, 22).Value = thismoveline
                        If thisAldiPrice = 0 Then .Cells(lNum, 23).Value = 0 Else .Cells(lNum, 23).Value = (thismoveline - thisAldiPrice) / thisAldiPrice
                    Else
                        .Cells(lNum, 22).Value = "N/A"
                        .Cells(lNum, 23).Value = "N/A"
                    End If
                End If
            Next
                        
            .PageSetup.PrintArea = Range(.Cells(1, 1), .Cells(lNum, 15)).Address
            .PageSetup.Zoom = False
            .PageSetup.FitToPagesWide = 1
            .PageSetup.FitToPagesTall = False
            .PageSetup.LeftFooter = "CORP BUYING, Admin, per: " & Format(Date, "DD/MM/YYYY") & Chr(10) & Application.ActiveWorkbook.FullName
            .PageSetup.Orientation = xlLandscape
            .PageSetup.PrintGridlines = True
            .PageSetup.PrintTitleRows = Range(.Cells(1, 1), .Cells(5, 15)).Address
            .PageSetup.RightFooter = "&P of &N"
            Range(.Cells(5, 1), .Cells(5, 23)).AutoFilter
            .Visible = xlSheetVisible
            .Activate
            .Cells(6, 1).Select
        End With
        
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    
    End If
    CBAR_Runtime.createActiveReport ("Active Match Report")
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AMR", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Function MoveRetailintoLine(ByVal MatchType As String, ByVal CompRetail As Single) As Single

    If InStr(1, MatchType, "CB") > 0 Or InStr(1, MatchType, "PB") > 0 Then
            MoveRetailintoLine = CompRetail * 0.9
    ElseIf InStr(1, MatchType, "ML") > 0 Or MatchType = "FCQ" Or MatchType = "DMQ" Then
            MoveRetailintoLine = CompRetail * 0.7
    ElseIf InStr(1, LCase(MatchType), "pl") > 0 Or InStr(1, LCase(MatchType), "colescoles") > 0 _
        Or InStr(1, LCase(MatchType), "wwww") > 0 Or InStr(1, LCase(MatchType), "select") > 0 Then
            MoveRetailintoLine = CompRetail * 0.9
    ElseIf InStr(1, LCase(MatchType), "sb") > 0 Or InStr(1, LCase(MatchType), "hb") > 0 Or _
        InStr(1, LCase(MatchType), "dm") > 0 Or InStr(1, LCase(MatchType), "fc") > 0 Or InStr(1, LCase(MatchType), "val") > 0 Then
            MoveRetailintoLine = CompRetail * 0.9
    ElseIf InStr(1, LCase(MatchType), "web") > 0 Or InStr(1, LCase(MatchType), "www") > 0 Then
            MoveRetailintoLine = 0
    End If

End Function

