Attribute VB_Name = "CBAR_StartPromo"
Option Explicit
Option Private Module
Private CM As Collection, WM As Collection, DM As Collection, FM As Collection, AM As Collection
Private CS As String, WS As String, DS As String, FS As String, AMS As String

Sub CBAR_OnPromoReport(Optional ByVal BuyerEmailers As Boolean)
    Dim DFrom As Date, Dto As Date
    Dim a As Long, b As Long
    Dim datearr As Variant
    Dim dirMatch As Scripting.Dictionary
    Dim no As Long, d As Long
    Dim m As Variant
    Dim wks_OP As Worksheet
    Dim st As String, strProds As String
    Dim scraped() As Date
    Dim dates As Long, Scnt As Long, weeks As Long
    Dim BDBADic As Scripting.Dictionary, prod As Variant
    Dim tR As CBAR_Report
    Dim isThereData As Boolean, bOutput As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    Set CM = New Collection: Set WM = New Collection: Set DM = New Collection: Set FM = New Collection: Set AM = New Collection
    CS = "": WS = "": DS = "": FS = "": AMS = ""

    tR = CBAR_Runtime.getActiveReport
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Preparing to run 'On Promo Report'"
        
        
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
            Exit Sub
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
            Exit Sub
        End If
    End If
    If tR.AldiProds Is Nothing Then
    Else
        For Each prod In tR.AldiProds
            If strProds = "" Then strProds = prod Else strProds = strProds & ", " & prod
        Next
    End If
    If tR.BD = "Produce" Or tR.CG = 58 Then Else Dto = CBA_COM_Runtime.CBA_getWedDate
    
    no = 5
    If BuyerEmailers = False Then
        If tR.BD = "Produce" Or tR.CG = 58 Then Else Dto = CBA_COM_Runtime.CBA_getWedDate
        
        Dto = CBA_COM_Runtime.CBA_getWedDate
        weeks = 4
        DFrom = DateAdd("WW", -weeks, Dto)
        
        dates = 0
        For d = 0 To DateDiff("D", DFrom, Dto)
            If WeekDay(DFrom + d, vbWednesday) = 1 Then
                dates = dates + 1
                ReDim Preserve scraped(1 To dates)
                scraped(dates) = DFrom + d
            End If
        Next
    Else
        Dto = CBA_COM_Runtime.CBA_getWedDate
        scraped = CBAR_ReportParamaters.getEmailerScrapedDatesArray
        dates = UBound(scraped)
        weeks = UBound(scraped) - 1
    End If
    
    
    Set dirMatch = New Scripting.Dictionary
    'petcare = 39
    
    If BuyerEmailers = False Then isThereData = CBA_COM_SetupMatchArray.CBA_SetupMatchArray(True, DFrom, Dto, tR.CG, tR.scg, strProds, True, tR.Competitor) Else isThereData = True
    If isThereData = True Then
        Application.ScreenUpdating = False
        
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
        
        Set BDBADic = getBDBADic
        
        For a = LBound(CBA_COM_Match) To UBound(CBA_COM_Match)
            If CBA_COM_Match(a).Pricedata(Dto, "Special", tR.State, False) = True Then
                If dirMatch.Exists(CBA_COM_Match(a).CompCode) = False Then
                    dirMatch.Add CBA_COM_Match(a).CompCode, CBA_COM_Match(a)
                    addMatchSM CBA_COM_Match(a)
                End If
            End If
        Next
        
        If dirMatch.Count > 0 Then
            CBAR_PA.Copy
            Set wks_OP = ActiveSheet
            wks_OP.Name = "StartPromo"
            CBAR_ReportParamaters.setBuyerEmailerWorksheet "CBAR_StartPromo", wks_OP
            With wks_OP
                .Rows(5).ClearContents
                .Cells(5, 1).Value = "Competitor"
                .Cells(5, 2).Value = "CompCode"
                .Cells(5, 3).Value = "Comp Description"
                .Cells(5, 4).Value = "Comp Packsize"
                .Cells(5, 5).Value = "State of Promotion"
                .Cells(5, 6).Value = "Non-Promo Retail"
                .Cells(5, 7).Value = "Promo Shelf Retail"
                .Cells(5, 8).Value = "Promo Discount"
                .Cells(5, 9).Value = "Promo ProRata Retail"
                .Cells(5, 10).Value = "Aldi Retail"
                .Cells(5, 11).Value = "Aldi Cheaper by %"
                .Cells(5, 12).Value = "Weeks on Promo"
                .Cells(5, 13).Value = "Aldi Product Code"
                .Cells(5, 14).Value = "Aldi Product Description"
                .Cells(5, 15).Value = "MatchType"
                .Cells(5, 16).Value = "CG"
                .Cells(5, 17).Value = "SCG"
                .Cells(5, 18).Value = "GBD"
                .Cells(5, 19).Value = "BD"
                .Cells(5, 20).Value = "BAs"
                For a = 1 To 20
                    Select Case a
                        Case 6, 7, 8, 9, 10
                            .Columns(a).NumberFormat = "$#,0.00"
                        Case 11
                            .Columns(a).NumberFormat = "#,0.0%"
                        
                        Case Else
                            .Columns(a).NumberFormat = "General"
                    End Select
                Next
                For a = 1 To 20
                    Select Case a
                        Case 1, 2, 4
                            .Columns(a).ColumnWidth = 13
                        Case 3
                            .Columns(a).ColumnWidth = 80
                        Case 5
                            .Columns(a).ColumnWidth = 20
                        Case 14
                            .Columns(a).ColumnWidth = 26
                        Case 6, 7, 8, 9, 10, 11, 18, 19, 20
                            .Columns(a).ColumnWidth = 6
                        Case 12, 16, 17
                            .Columns(a).ColumnWidth = 4
                        Case 13, 15
                            .Columns(a).ColumnWidth = 8
                    End Select
                Next
                .Cells(3, 3).Value = "Comp Products On Promotion"
                For Each m In dirMatch
                    no = no + 1
                    .Cells(no, 1).Value = dirMatch(m).Competitor
                    .Cells(no, 2).Value = dirMatch(m).CompCode
                    .Cells(no, 3).Value = dirMatch(m).CompProdName
                    .Cells(no, 4).Value = dirMatch(m).CompPacksize
                    .Cells(no, 6).Value = dirMatch(m).Pricedata(Dto, "Shelf", tR.State) + dirMatch(m).Pricedata(Dto, "discount", tR.State)
                    .Cells(no, 7).Value = dirMatch(m).Pricedata(Dto, "shelf", tR.State)
                    .Cells(no, 8).Value = dirMatch(m).Pricedata(Dto, "discount", tR.State)
                    .Cells(no, 9).Value = dirMatch(m).Pricedata(Dto, "prorata", tR.State)
                    .Cells(no, 10).Value = dirMatch(m).Pricedata(Dto, "aldiretail", tR.State)
                    .Cells(no, 11).Value = dirMatch(m).Pricedata(Dto, "delta", tR.State)
                    If .Cells(no, 11).Value <> 0 Then .Cells(no, 11).Interior.ColorIndex = CBA_COM_ReportFunctions.CBA_COM_isTrafficLight(.Cells(no, 11).Value, dirMatch(m).MatchType)
                    Scnt = 0
                    For a = UBound(scraped) - 1 To LBound(scraped) Step -1
                        If dirMatch(m).Pricedata(scraped(a), "special", tR.State) = True Then
                            Scnt = Scnt + 1
                        Else
                            Exit For
                        End If
                    Next
                    If .Cells(no, 5).Value = "" Then
                        For b = 1 To 5
                            If b = 1 Then st = "NSW" Else If b = 2 Then st = "VIC" Else If b = 3 Then st = "QLD" Else If b = 4 Then st = "SA" Else If b = 5 Then st = "WA"
                            If dirMatch(m).Pricedata(Dto, "special", st) = True Then
                                If .Cells(no, 5).Value = "" Then
                                    .Cells(no, 5).Value = st
                                Else
                                    .Cells(no, 5).Value = .Cells(no, 5).Value & ", " & st
                                End If
                            End If
                        Next
                    End If
                    If Scnt = weeks Then .Cells(no, 12).Value = weeks & "+" Else .Cells(no, 12).Value = "NEW"
                    .Cells(no, 13).Value = dirMatch(m).AldiPCode
                    .Cells(no, 14).Value = dirMatch(m).AldiPName
                    .Cells(no, 15).Value = dirMatch(m).MatchType
                    .Cells(no, 16).Value = dirMatch(m).AldiPCG
                    .Cells(no, 17).Value = dirMatch(m).AldiPSCG
                    If BDBADic.Exists(dirMatch(m).AldiPCode) = True Then
                    .Cells(no, 18).Value = BDBADic(dirMatch(m).AldiPCode)("GBD")
                    .Cells(no, 19).Value = BDBADic(dirMatch(m).AldiPCode)("BD")
                    .Cells(no, 20).Value = BDBADic(dirMatch(m).AldiPCode)("BAs")
                    End If
                Next
                If no > 5 Then Range(.Cells(5, 1), .Cells(5, 20)).AutoFilter
    
            End With
        Else
            MsgBox "No Activity to report"
        End If
        Application.ScreenUpdating = True
    Else
        MsgBox "No Activity to report"
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBAR_OnPromoReport", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub
Function addMatchSM(ByRef MatchItem As CBA_COM_COMMatch)

    If MatchItem.compet = "C" Then
        CM.Add MatchItem.CompCode
        If CS = "" Then CS = MatchItem.CompCode Else CS = CS & ", " & MatchItem.CompCode
    ElseIf MatchItem.compet = "WW" Then
        WM.Add MatchItem.CompCode
        If WS = "" Then WS = MatchItem.CompCode Else WS = WS & ", " & MatchItem.CompCode
    ElseIf MatchItem.compet = "DM" Then
        DM.Add MatchItem.CompCode
        If DS = "" Then DS = MatchItem.CompCode Else DS = DS & ", " & MatchItem.CompCode
    ElseIf MatchItem.compet = "FC" Then
        FM.Add MatchItem.CompCode
        If FS = "" Then FS = MatchItem.CompCode Else FS = FS & ", " & MatchItem.CompCode
    ElseIf MatchItem.compet = "AMZ" Then
        AM.Add MatchItem.CompCode
        If AMS = "" Then AMS = MatchItem.CompCode Else AMS = AMS & ", " & MatchItem.CompCode
    End If

End Function

