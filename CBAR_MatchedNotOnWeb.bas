Attribute VB_Name = "CBAR_MatchedNotOnWeb"
Option Explicit
Option Private Module
Private CS As String, WS As String, DS As String, FS As String, AMS As String

Sub CBAR_MatchedNotOnWeb(Optional ByVal BuyerEmailers As Boolean)
    Dim DFrom As Date, Dto As Date
    Dim a As Long, b As Long, c As Long, d As Long, num As Long
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
    Dim bOutput As Boolean
    Dim prod As Variant, this As Variant
    Dim SKUDic As Scripting.Dictionary, CSKUDic As Scripting.Dictionary
    Dim WSKUDic As Scripting.Dictionary, DSKUDic As Scripting.Dictionary
    Dim FSKUDic As Scripting.Dictionary, ASKUDic As Scripting.Dictionary
    Dim tarrC() As CBA_COM_COMCompSKU, tarrW() As CBA_COM_COMCompSKU, tarrD() As CBA_COM_COMCompSKU, tarrF() As CBA_COM_COMCompSKU, tarrA() As CBA_COM_COMCompSKU
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    CS = "": WS = "": DS = "": FS = "": AMS = ""

    tR = CBAR_Runtime.getActiveReport
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Preparing to run 'Matched Not On Web Report'"
    
    
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
    
    Set dirMatch = New Scripting.Dictionary
    'petcare = 39
    
    
    ReDim Marr(1 To 8, 1 To 1)
    
    CBA_COM_SQLQueries.CBA_COM_GenPullSQL "CBA_COM_SKU_Prods", DFrom, Dto, , , , WeekDay(Date, vbThursday)
    For a = LBound(CBA_COMarr, 2) To UBound(CBA_COMarr, 2)
        If CBA_COMarr(0, a) = "Coles" Then
            If CSKUDic Is Nothing Then Set CSKUDic = New Scripting.Dictionary
            If CSKUDic.Exists(CBA_COMarr(1, a)) = False Then CSKUDic.Add CBA_COMarr(1, a), CBA_COMarr(2, a)
        ElseIf CBA_COMarr(0, a) = "Woolworths" Then
            If WSKUDic Is Nothing Then Set WSKUDic = New Scripting.Dictionary
            If WSKUDic.Exists(CBA_COMarr(1, a)) = False Then WSKUDic.Add CBA_COMarr(1, a), CBA_COMarr(2, a)
        ElseIf CBA_COMarr(0, a) = "Dan Murphys" Then
            If DSKUDic Is Nothing Then Set DSKUDic = New Scripting.Dictionary
            If DSKUDic.Exists(CBA_COMarr(1, a)) = False Then DSKUDic.Add CBA_COMarr(1, a), CBA_COMarr(2, a)
        ElseIf CBA_COMarr(0, a) = "First Choice" Then
            If FSKUDic Is Nothing Then Set FSKUDic = New Scripting.Dictionary
            If FSKUDic.Exists(CBA_COMarr(1, a)) = False Then FSKUDic.Add CBA_COMarr(1, a), CBA_COMarr(2, a)
        ElseIf CBA_COMarr(0, a) = "Amazon" Then
            If ASKUDic Is Nothing Then Set ASKUDic = New Scripting.Dictionary
            If ASKUDic.Exists(CBA_COMarr(1, a)) = False Then ASKUDic.Add CBA_COMarr(1, a), CBA_COMarr(2, a)
        End If
    Next
    Set SKUDic = New Scripting.Dictionary
    If Not CSKUDic Is Nothing Then SKUDic.Add "co", CSKUDic
    If Not WSKUDic Is Nothing Then SKUDic.Add "ww", WSKUDic
    If Not DSKUDic Is Nothing Then SKUDic.Add "dm", DSKUDic
    If Not FSKUDic Is Nothing Then SKUDic.Add "fc", FSKUDic
    If Not ASKUDic Is Nothing Then SKUDic.Add "am", ASKUDic
    
    'COMMatcharr = CBA_COMarr
    Erase CBA_COMarr
    If CBAR_SQLQueries.CBAR_GenPullSQL("CBAR_MatchedwDBName", , , tR.CG, CStr(tR.scg), , strProds) = True Then arrMatches = CBA_COMarr: Erase CBA_COMarr
    num = 0
    For a = LBound(arrMatches, 2) To UBound(arrMatches, 2)
        If a / 300 = Round(a / 300, 0) Then CBA_BasicFunctions.RunningSheetAddComment 6, 5, "Compared Match: " & a & " Of " & UBound(arrMatches, 2)
        this = CCM_Mapping.CMM_getComp2Find(arrMatches(3, a), arrMatches(1, a))
        If IsEmpty(this) = False Then
            If SKUDic(LCase(Left(this, 2))).Exists(arrMatches(2, a)) = False Then
                num = num + 1
                ReDim Preserve Marr(1 To 8, 1 To num)
                Marr(4, num) = this
                If LCase(Left(Marr(4, num), 5)) = "coles" Then
                    Marr(1, num) = "Coles"
                ElseIf LCase(Left(Marr(4, num), 2)) = "ww" Then
                    Marr(1, num) = "Woolworths"
                ElseIf LCase(Left(Marr(4, num), 2)) = "dm" Then
                    Marr(1, num) = "Dan Murphys"
                ElseIf LCase(Left(Marr(4, num), 2)) = "fc" Then
                    Marr(1, num) = "First Choice"
                ElseIf LCase(Left(Marr(4, num), 3)) = "amz" Then
                    Marr(1, num) = "Amazon"
                End If
                Marr(2, num) = arrMatches(2, a)
                If CBAR_SQLQueries.CBAR_GenPullSQL("COM_PDesc", , , , , UCase(Left(Marr(4, num), 2)), CStr(arrMatches(2, a))) = True Then Marr(3, num) = CBA_COMarr(0, 0)
                Marr(4, num) = CCM_Mapping.CMM_getComp2Find(arrMatches(3, a), arrMatches(1, a))
                Marr(5, num) = arrMatches(0, a)
                Marr(6, num) = arrMatches(1, a)
            End If
        End If
    Next
    
    If num > 0 Then
        Application.ScreenUpdating = False
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
        Set BDBADic = getBDBADic
        
            CBAR_PA.Copy
            Set wks_OP = ActiveSheet
            CBAR_ReportParamaters.setBuyerEmailerWorksheet "CBA_MatchedNotOnWeb", wks_OP
            wks_OP.Name = "MatchedNotOnWeb"
            With wks_OP
                .Rows(5).ClearContents
                .Cells(5, 1).Value = "Competitor"
                .Cells(5, 2).Value = "CompCode"
                .Cells(5, 3).Value = "Comp Description"
                .Cells(5, 4).Value = "MatchType"
                .Cells(5, 5).Value = "Aldi Product Code"
                .Cells(5, 6).Value = "CG"
                .Cells(5, 7).Value = "GBD"
                .Cells(5, 8).Value = "BD"
                .Cells(5, 9).Value = "BAs"
                For a = 1 To 10
                    .Columns(a).NumberFormat = "General"
                Next
                .Cells.EntireColumn.AutoFit
                .Cells(3, 3).Value = "Matched Products not on Website"
                For a = 1 To num
                    For b = 1 To 7
                        .Cells(a + 5, b).Value = Marr(b, a)
                    Next
                    If BDBADic.Exists(CLng(Marr(5, a))) = True Then
                        .Cells(a + 5, 7).Value = BDBADic(CLng(Marr(5, a)))("GBD")
                        .Cells(a + 5, 8).Value = BDBADic(CLng(Marr(5, a)))("BD")
                        .Cells(a + 5, 9).Value = BDBADic(CLng(Marr(5, a)))("BAs")
                    End If
                Next
                .Cells.EntireColumn.AutoFit
                If num > 0 Then Range(.Cells(5, 1), .Cells(5, 9)).AutoFilter
            End With
        Application.ScreenUpdating = True
    Else
        MsgBox "No Activity to report"
    End If

Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBAR_MatchedNotOnWeb", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub

