Attribute VB_Name = "CBAR_ALdiNoMatch"
'Option Explicit
Option Private Module
'Private CM As Scripting.Dictionary, WM As Scripting.Dictionary, DM As Scripting.Dictionary, FM As Scripting.Dictionary, AM As Scripting.Dictionary
Private CS As String, WS As String, DS As String, FS As String, AMS As String
Sub CBAR_AldiNoMatch(Optional ByVal BuyerEmailers As Boolean)
    Dim DFrom As Date, Dto As Date
    Dim a As Long, b As Long
    Dim datearr As Variant
    Dim dirMatch As Scripting.Dictionary
    Dim no As Long
    Dim m As Variant
    Dim wks_OP As Worksheet
    Dim st As String, strProds As String
    Dim scraped() As Date
    Dim dates As Long, Scnt As Long, weeks As Long
    Dim BDBADic As Scripting.Dictionary
    Dim tR As CBAR_Report
    Dim arrMatches As Variant, COMMatcharr As Variant, Marr As Variant
    Dim tarrC() As CBA_COM_COMCompSKU, tarrW() As CBA_COM_COMCompSKU, tarrD() As CBA_COM_COMCompSKU, tarrF() As CBA_COM_COMCompSKU, tarrA() As CBA_COM_COMCompSKU
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    'Set CM = New Scripting.Dictionary: Set WM = New Scripting.Dictionary: Set DM = New Scripting.Dictionary: Set FM = New Scripting.Dictionary: Set AM = New Scripting.Dictionary
    CS = "": WS = "": DS = "": FS = "": AMS = ""

    tR = CBAR_Runtime.getActiveReport
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Preparing to run 'Active ALDI Product with no COMRADE match'"
    
        
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
    no = 5
    If BuyerEmailers = False Then
    If tR.BD = "Produce" Or tR.CG = 58 Then Dto = DateAdd(d, -1, Date) Else Dto = CBA_COM_Runtime.CBA_getWedDate
    'dto = #2/27/2019#
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
    End If
    
    Set BDBADic = getBDBADic
    
    ReDim Marr(1 To 9, 1 To 1)
    CCM_SQLQueries.CBA_COM_MATCHGenPullSQL "getActiveProdsCGSCGBuyerDetail", , , tR.CG, CStr(tR.scg), , strProds
    COMMatcharr = CBA_CBISarr
    If CBAR_SQLQueries.CBAR_GenPullSQL("CBAR_MatchedwDBName", , , tR.CG, CStr(tR.scg), , strProds) = True Then arrMatches = CBA_COMarr: Erase CBA_COMarr
    num = 0
    For a = LBound(COMMatcharr, 2) To UBound(COMMatcharr, 2)
        bfound = False
        For b = LBound(arrMatches, 2) To UBound(arrMatches, 2)
            If arrMatches(0, b) = CStr(COMMatcharr(0, a)) Then
                bfound = True
                Exit For
            End If
        Next
        If bfound = False Then
            num = num + 1
            ReDim Preserve Marr(1 To 9, 1 To num)
            For b = 1 To 7
                Marr(b, num) = COMMatcharr(b - 1, a)
            Next
            If BDBADic.Exists(Marr(1, num)) = True Then
                Marr(7, num) = BDBADic(Marr(1, num))("GBD")
                Marr(8, num) = BDBADic(Marr(1, num))("BD")
                Marr(9, num) = BDBADic(Marr(1, num))("BAs")
            End If
        End If
    Next
    
    
    If num > 0 Then
        Application.ScreenUpdating = False
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
            CBAR_PA.Copy
            Set wks_OP = ActiveSheet
            CBAR_ReportParamaters.setBuyerEmailerWorksheet "CBAR_AldiNoMatch", wks_OP
            wks_OP.Name = "ALDINoMatch"
            With wks_OP
                .Rows(5).ClearContents
                .Cells(5, 1).Value = "Aldi Product Code"
                .Cells(5, 2).Value = "Description"
                .Cells(5, 3).Value = "CG"
                .Cells(5, 4).Value = "CG Description"
                .Cells(5, 5).Value = "SCG"
                .Cells(5, 6).Value = "SCG Description"
                .Cells(5, 7).Value = "GBD"
                .Cells(5, 8).Value = "BD"
                .Cells(5, 9).Value = "BAs"
                For a = 1 To 10
                    .Columns(a).NumberFormat = "General"
                Next
                .Cells.EntireColumn.AutoFit
                .Cells(5, 3).EntireColumn.ColumnWidth = 4
                .Cells(3, 3).Value = "Active ALDI Product with no COMRADE Matches"
                For a = 1 To num
                    For b = 1 To 9
                        .Cells(a + 5, b).Value = Marr(b, a)
                    Next
                Next
                .Cells.EntireColumn.AutoFit
                If num > 0 Then Range(.Cells(5, 1), .Cells(5, 9)).AutoFilter
                .Columns(3).ColumnWidth = 5
            End With
        Application.ScreenUpdating = True
    Else
        MsgBox "No Activity to report"
    End If
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBAR_AldiNoMatch", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
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



