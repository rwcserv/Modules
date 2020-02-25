Attribute VB_Name = "CBAR_Runtime"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Type CBAR_Report
    ReportName As String
    ReportNo As Long
    DateFrom As Date
    DateTo As Date
    AldiProds As Collection
    CG As Long
    CGDesc As String
    scg As Long
    SCGDesc As String
    State As String
    Stateno As Long
    Competitor As String
    CompetitorNo As Long
    BD As String
    GBD As String
    Matchhistory As Boolean
End Type
Private CBAR_GBDs As Collection
Private CBAR_ActiveReport As CBAR_Report
Private CBAR_noRepsAdmin As Long
Private CBAR_noReps As Long
Private CBAR_RepsAdmin() As String
Private CBAR_Reps() As String
Private CBAR_ProdGBDBDEmails() As Variant
Private CBAR_BDBAEmails() As Variant
Private CBAR_BDBADic As Scripting.Dictionary

Function CBAR_CheckTop150List(ByRef colProds As Collection, ByRef colPProds As Collection) As Boolean
    Dim strAllProds As String
    Dim dicAlts As Scripting.Dictionary
    Dim a As Long
    Dim ArrProdError As Variant
    Dim pro As Variant
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    CBA_DBtoQuery = 599
    
    strAllProds = ""
    For Each pro In colProds
        If strAllProds = "" Then strAllProds = pro Else strAllProds = strAllProds & ", " & pro
    Next
    For Each pro In colPProds
        If strAllProds = "" Then strAllProds = pro Else strAllProds = strAllProds & ", " & pro
    Next
    CBAR_SQLQueries.CBAR_GenPullSQL "COM_CheckT150List", , , , , , strAllProds
    
    
    If CBA_CBISarr(0, 0) = 0 Then
        CBAR_CheckTop150List = True
        Exit Function
    Else
        ArrProdError = CBA_CBISarr
        Set dicAlts = T150FindAlternative(ArrProdError)
        If dicAlts.Exists("Quit") Then
            CBAR_CheckTop150List = False
            Exit Function
        End If
            
        
        For a = 1 To colProds.Count
            If dicAlts.Exists(colProds.Item(a)) Then
                colProds.Add dicAlts(colProds.Item(a))
                colProds.Remove a
            End If
        Next
        For a = 1 To colPProds.Count
            If dicAlts.Exists(colPProds.Item(a)) Then
                colPProds.Add dicAlts(colPProds.Item(a))
                colPProds.Remove a
            End If
        Next
    End If
    
    CBAR_CheckTop150List = True
    CBA_DBtoQuery = 3
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBAR_CheckTop150List", 3)
    CBAR_CheckTop150List = False
    CBA_DBtoQuery = 3
    
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Function T150FindAlternative(ByRef ArrProdError As Variant) As Scripting.Dictionary
    Dim arrParts() As String, NPCode As String, strPrompt As String
    Dim arrThings() As String, colWordsToQuery As Collection
    Dim a As Long, b As Long, P As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
        
    Set T150FindAlternative = New Scripting.Dictionary
    
    ReDim arrThings(0 To 13)
    arrThings(0) = "g": arrThings(1) = "ml": arrThings(2) = "pk": arrThings(3) = "kg"
    arrThings(4) = "pc": arrThings(5) = "sht": arrThings(6) = "bunch": arrThings(7) = "gram"
    arrThings(8) = "l": arrThings(9) = "millilitre": arrThings(10) = "pack": arrThings(11) = "kilogram"
    arrThings(12) = "each": arrThings(13) = "per"
    
    For P = LBound(ArrProdError, 2) To UBound(ArrProdError, 2)
        Set colWordsToQuery = New Collection
        arrParts = Split(ArrProdError(1, P), " ")
        For a = LBound(arrParts) To UBound(arrParts)
            If Right(arrParts(a), 1) = "s" Then arrParts(a) = Left(arrParts(a), Len(arrParts(a)) - 1)
            For b = LBound(arrThings) To UBound(arrThings)
                If Len(arrThings(b)) < Len(arrParts(a)) Then
                    If IsNumeric(Left(arrParts(a), Len(arrParts(a)) - Len(arrThings(b)))) Then GoTo canMiss
                End If
                If arrParts(a) = arrThings(b) Or IsNumeric(arrParts(a)) Then GoTo canMiss
            Next
            colWordsToQuery.Add arrParts(a)
canMiss:
        Next
    
        CBAR_SQLQueries.CBAR_GenPullSQL "CBAR_T150NewPotentials", , , , , colWordsToQuery
        strPrompt = "Please Select a replacement productcode for: " & Chr(10) & ArrProdError(0, P) & " - " & ArrProdError(1, P) & Chr(10) & Chr(10)
        strPrompt = strPrompt & "The apparent options are:" & Chr(10) & Chr(10)
        For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
            If a > 0 Then
                If CBA_CBISarr(0, a - 1) <> CBA_CBISarr(0, a) Then
                    If CBA_CBISarr(2, a) = ArrProdError(2, P) Then strPrompt = strPrompt & CBA_CBISarr(0, a) & " - " & CBA_CBISarr(1, a) & Chr(10)
                End If
            Else
                If CBA_CBISarr(2, a) = ArrProdError(2, P) Then strPrompt = strPrompt & CBA_CBISarr(0, a) & " - " & CBA_CBISarr(1, a) & Chr(10)
            End If
        Next
        strPrompt = strPrompt & Chr(10) & "To Quit Running Top150 type 'Quit'"
TryAgain:
        NPCode = InputBox(strPrompt, "Choose New Top 150 ProductCodes")
        
        If NPCode = "" Then GoTo TryAgain
        If LCase(NPCode) = "quit" Then
            T150FindAlternative.Add "Quit", "Quit"
            Exit Function
        End If
        For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
            If CBA_CBISarr(0, a) = NPCode Then
                T150FindAlternative.Add ArrProdError(0, P), CBA_CBISarr(0, a)
                Exit For
            End If
        Next

    Next
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-T150FindAlternative", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function


Sub CBAR_SetRep()
    'IF YOU BUILD A NEW REPORT YOU NEED TO UPDATE THE BELOW VALUES AND ADD THE NAME TO THE ARRAYS
    'WHEN COMRADE LINK IS CREATED THIS SUB IS RUN
    CBAR_noRepsAdmin = 13
    CBAR_noReps = 11
    
    ReDim CBAR_RepsAdmin(1 To CBAR_noRepsAdmin)
    ReDim CBAR_Reps(1 To CBAR_noReps)
    
    CBAR_RepsAdmin(1) = "Active Match Report"
    CBAR_Reps(1) = "Active Match Report"
    CBAR_RepsAdmin(2) = "Price and Promotion History Report"
    CBAR_Reps(2) = "Price and Promotion History Report"
    CBAR_RepsAdmin(3) = "State Retail Variation Report"
    CBAR_Reps(3) = "State Retail Variation Report"
    CBAR_RepsAdmin(4) = "Missing Match Report"
    CBAR_Reps(4) = "Missing Match Report"
    CBAR_RepsAdmin(5) = "Promotional Analysis Report"
    CBAR_Reps(5) = "Promotional Analysis Report"
    CBAR_RepsAdmin(6) = "Off Promo Report"
    CBAR_Reps(6) = "Off Promo Report"
    CBAR_RepsAdmin(7) = "On Promo Report"
    CBAR_Reps(7) = "On Promo Report"
    CBAR_RepsAdmin(8) = "Permanent Price Change Report"
    CBAR_Reps(8) = "Permanent Price Change Report"
    CBAR_RepsAdmin(9) = "Matched Not On Comp Site"
    CBAR_Reps(9) = "Matched Not on Comp Site"
    CBAR_RepsAdmin(10) = "Active ALDI Prod with no Match"
    CBAR_Reps(10) = "Active ALDI Prod with no Match"
    CBAR_RepsAdmin(11) = "Weighted Basket Analysis (by Share)"
    CBAR_Reps(11) = "Weighted Basket Analysis (by Share)"
''    CBAR_RepsAdmin(12) = "Weighted Basket Analysis (by Share) - ADMIN"
''    CBAR_RepsAdmin(13) = "Price Relativity Report"
''    CBAR_RepsAdmin(14) = "Top 150"
''    CBAR_RepsAdmin(15) = "COMRADE Weekly Emails"
    '''CBAR_RepsAdmin(12) = "Weighted Basket Analysis (by Share) - ADMIN"
'    CBAR_RepsAdmin(12) = "Price Relativity Report"
    CBAR_RepsAdmin(12) = "Top 150"
    CBAR_RepsAdmin(13) = "COMRADE Weekly Emails"
End Sub
Function CBAR_getNamesOfReports(ByVal AdminUser As Boolean) As String()
    If AdminUser = True Then
       CBAR_getNamesOfReports = CBAR_RepsAdmin
    Else
       CBAR_getNamesOfReports = CBAR_Reps
    End If
End Function
Function CBAR_getNoOfReports(ByVal AdminUser As Boolean) As Long
    If AdminUser = True Then
       CBAR_getNoOfReports = CBAR_noRepsAdmin
    Else
       CBAR_getNoOfReports = CBAR_noReps
    End If
End Function
Sub RunMatchingToolfromAMR(ByVal Competitor As String, ByVal collecter As Collection)
    If Competitor = "WW" Then CCM_UDWWSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("WW", collecter)
    If Competitor = "Coles" Then CCM_UDColesSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("C", collecter)
    If Competitor = "DM" Then CCM_UDDMSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("DM", collecter)
    If Competitor = "FC" Then CCM_UDFCSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("FC", collecter)
    CBA_COM_frm_MatchingTool.setCCM_UserDefinedState (True)
    CCM_Runtime.CCM_MatchingSelectorActivate
End Sub
Sub AddProdToMatchingTool(ByVal ValToAdd As String)
    CBA_COM_frm_MatchingTool.addProdtoList ValToAdd
End Sub
Sub SetMatchTypeOnMatchingTool(ByVal Mtype As String)
    CBA_COM_frm_MatchingTool.setMatchType CCM_Mapping.MatchType(Mtype).OptionButtonName
End Sub
Sub CBAR_setGBSs()
    Dim TempArr() As Variant, a As Long
    
    Set CBAR_GBDs = New Collection
    TempArr = CCM_Runtime.CBA_COM_GetGBDNames
    For a = LBound(TempArr, 2) To UBound(TempArr, 2)
        CBAR_GBDs.Add TempArr(0, a)
    Next
End Sub
Function CBAR_setEmails()
    CBAR_SQLQueries.CBAR_GenPullSQL "CBAR_ProdGBDBDEmail"
    CBAR_ProdGBDBDEmails = CBA_CBISarr
    Erase CBA_CBISarr
    CBAR_SQLQueries.CBAR_GenPullSQL "CBAR_BDBAEmail"
    CBAR_BDBAEmails = CBA_CBISarr
    Erase CBA_CBISarr
End Function
Function getProdGBDBDEmails() As Variant
    getProdGBDBDEmails = CBAR_ProdGBDBDEmails
End Function
Function getBDBAEmails() As Variant
    getBDBAEmails = CBAR_BDBAEmails
End Function
Function CBAR_getGBSs() As Collection
    Set CBAR_getGBSs = CBAR_GBDs
End Function
Function getActiveReport() As CBAR_Report
    getActiveReport = CBAR_ActiveReport
End Function
Function createActiveReport(ByVal ReportName As String)
    clearActiveReport
    CBAR_ActiveReport.ReportName = ReportName
    Select Case ReportName
        Case "Active Match Report"
            CBAR_ActiveReport.ReportNo = 1
        Case "Price and Promotion History Report"
            CBAR_ActiveReport.ReportNo = 2
        Case "State Retail Variation Report"
            CBAR_ActiveReport.ReportNo = 3
        Case "Missing Match Report"
            CBAR_ActiveReport.ReportNo = 4
        Case "Promotional Analysis Report"
            CBAR_ActiveReport.ReportNo = 5
        Case "Off Promo Report"
            CBAR_ActiveReport.ReportNo = 6
        Case "On Promo Report"
            CBAR_ActiveReport.ReportNo = 7
        Case "Permanent Price Change Report"
            CBAR_ActiveReport.ReportNo = 8
        Case "Matched Not On Comp Site"
            CBAR_ActiveReport.ReportNo = 9
        Case "Active ALDI Prod with no Match"
            CBAR_ActiveReport.ReportNo = 10
        Case "Weighted Basket Analysis (by Share)"
            CBAR_ActiveReport.ReportNo = 11
''        Case "Weighted Basket Analysis (by Share) - ADMIN"
''            CBAR_ActiveReport.ReportNo = 12
''        Case "Price Relativity Report"                    ' @RW Report doesn't call anything
''            CBAR_ActiveReport.ReportNo = 12
        Case "Top 150"
            CBAR_ActiveReport.ReportNo = 12
        Case "COMRADE Weekly Emails"
            CBAR_ActiveReport.ReportNo = 13
    End Select
End Function
Function setActiveReportParamater(ByVal Paramater As String, Optional ByVal StringVal As String, Optional ByVal LongVal As Long, Optional ByVal dateval As Date) As Boolean
    Dim bOutput As Boolean
    Dim a As Long, stpt As Long, b As Long, prod, strProducts As String

    Select Case LCase(Paramater)
        Case "datefrom"
            If isDate(dateval) And dateval <> 0 Then
                CBAR_ActiveReport.DateFrom = Format(dateval, "YYYY-MM-DD")
                setActiveReportParamater = True
            End If
        Case "dateto"
            If isDate(dateval) And dateval <> 0 Then
                CBAR_ActiveReport.DateTo = Format(dateval, "YYYY-MM-DD")
                setActiveReportParamater = True
            End If
        Case "aldiprods"
            If StringVal <> "" Then
                Set CBAR_ActiveReport.AldiProds = New Collection
                stpt = 1
                For a = 1 To Len(StringVal)
                    If Mid(StringVal, a, 1) = " " Or Mid(StringVal, a, 1) = Chr(10) Or Mid(StringVal, a, 1) = "/" Or Mid(StringVal, a, 1) = "," Or Mid(StringVal, a, 1) = ":" Or Mid(StringVal, a, 1) = ";" Or Mid(StringVal, a, 1) = "." Or a = Len(StringVal) Then
                        If a = Len(StringVal) And IsNumeric(Mid(StringVal, a, 1)) = True Then b = 1 Else b = 0
                        If IsNumeric(Mid(StringVal, stpt, a - stpt + b)) = True Then
                            CBAR_ActiveReport.AldiProds.Add Mid(StringVal, stpt, a - stpt + b)
                            stpt = a + 1
                            setActiveReportParamater = True
                        End If
                    End If
                Next
                For Each prod In CBAR_ActiveReport.AldiProds
                    If strProducts = "" Then strProducts = "The query will run for the following products:" & Chr(10)
                    strProducts = strProducts & prod & Chr(10)
                Next
                MsgBox strProducts, vbOKOnly
            Else
                Set CBAR_ActiveReport.AldiProds = Nothing
            End If
        Case "cg"
            CBAR_ActiveReport.CG = LongVal
            setActiveReportParamater = True
            bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBIS_CGDesc", , , LongVal)
            If bOutput = True Then CBAR_ActiveReport.CGDesc = CBA_CBISarr(0, 0)
        Case "scg"
            CBAR_ActiveReport.scg = LongVal
            setActiveReportParamater = True
            bOutput = CBA_SQL_Queries.CBA_GenPullSQL("getSCGDesc", , , CBAR_ActiveReport.CG, CStr(LongVal))
            If bOutput = True Then CBAR_ActiveReport.SCGDesc = CBA_CBISarr(0, 0)
        Case "state"
            If StringVal = "NSW" Then
                CBAR_ActiveReport.State = StringVal
                CBAR_ActiveReport.Stateno = 1
                setActiveReportParamater = True
            ElseIf StringVal = "VIC" Then
                CBAR_ActiveReport.State = StringVal
                CBAR_ActiveReport.Stateno = 2
                setActiveReportParamater = True
            ElseIf StringVal = "QLD" Then
                CBAR_ActiveReport.State = StringVal
                CBAR_ActiveReport.Stateno = 3
                setActiveReportParamater = True
            ElseIf StringVal = "SA" Then
                CBAR_ActiveReport.State = StringVal
                CBAR_ActiveReport.Stateno = 4
                setActiveReportParamater = True
            ElseIf StringVal = "WA" Then
                CBAR_ActiveReport.State = StringVal
                CBAR_ActiveReport.Stateno = 5
                setActiveReportParamater = True
            ElseIf StringVal = "National" Then
                CBAR_ActiveReport.State = StringVal
                CBAR_ActiveReport.Stateno = 6
                setActiveReportParamater = True
            End If
        Case "competitor"
            If StringVal <> "" Then
                If StringVal = "Woolworths" Then
                    CBAR_ActiveReport.Competitor = StringVal
                    CBAR_ActiveReport.CompetitorNo = 1
                    setActiveReportParamater = True
                ElseIf StringVal = "Coles" Then
                    CBAR_ActiveReport.Competitor = StringVal
                    CBAR_ActiveReport.CompetitorNo = 2
                    setActiveReportParamater = True
                ElseIf StringVal = "Dan Murphys" Then
                    CBAR_ActiveReport.Competitor = StringVal
                    CBAR_ActiveReport.CompetitorNo = 3
                    setActiveReportParamater = True
                ElseIf StringVal = "First Choice" Then
                    CBAR_ActiveReport.Competitor = StringVal
                    CBAR_ActiveReport.CompetitorNo = 4
                    setActiveReportParamater = True
                ElseIf StringVal = "All Competitors" Then
                    CBAR_ActiveReport.Competitor = StringVal
                    CBAR_ActiveReport.CompetitorNo = 5
                    setActiveReportParamater = True
                End If
            End If
        Case "bd"
            'If stringval <> "" Then
                CBAR_ActiveReport.BD = StringVal
                setActiveReportParamater = True
            'End If
        Case "gbd"
            'If stringval <> "" Then
                CBAR_ActiveReport.GBD = StringVal
                setActiveReportParamater = True
            'End If
        Case "matchhistory"
            'If stringval <> "" Then
                If StringVal = "Yes" Then CBAR_ActiveReport.Matchhistory = True Else CBAR_ActiveReport.Matchhistory = False
                setActiveReportParamater = True
            'End If
    End Select
    
End Function


Function clearActiveReport()
    Set CBAR_ActiveReport.AldiProds = Nothing
    CBAR_ActiveReport.BD = ""
    CBAR_ActiveReport.CGDesc = ""
    CBAR_ActiveReport.CG = 0
    CBAR_ActiveReport.Competitor = ""
    CBAR_ActiveReport.CompetitorNo = 0
    CBAR_ActiveReport.DateFrom = 0
    CBAR_ActiveReport.DateTo = 0
    CBAR_ActiveReport.GBD = ""
    CBAR_ActiveReport.BD = ""
    CBAR_ActiveReport.ReportName = ""
    CBAR_ActiveReport.ReportNo = 0
    CBAR_ActiveReport.SCGDesc = ""
    CBAR_ActiveReport.scg = 0
    CBAR_ActiveReport.State = ""
    CBAR_ActiveReport.Stateno = 0
End Function

Function getBDBADic() As Scripting.Dictionary
    Dim BDEmail As Variant, BAEmail As Variant
    Dim thisDic As Scripting.Dictionary
    Dim a As Long, b As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    If CBAR_BDBADic Is Nothing Then
        BDEmail = CBAR_Runtime.getProdGBDBDEmails
        BAEmail = CBAR_Runtime.getBDBAEmails
        Set CBAR_BDBADic = New Scripting.Dictionary
        For a = LBound(BDEmail, 2) To UBound(BDEmail, 2)
            Set thisDic = New Scripting.Dictionary
            thisDic.Add "GBD", BDEmail(2, a)
            thisDic.Add "BD", BDEmail(4, a)
            For b = LBound(BAEmail, 2) To UBound(BAEmail, 2)
                If BAEmail(0, b) = BDEmail(3, a) Then
                    If thisDic.Exists("BAs") = True Then thisDic("BAs") = thisDic("BAs") & "; " & BAEmail(3, b) Else thisDic.Add "BAs", BAEmail(3, b)
                End If
                If BAEmail(0, b) > BDEmail(3, a) Then Exit For
            Next
            CBAR_BDBADic.Add BDEmail(0, a), thisDic
        Next
    End If
    Set getBDBADic = CBAR_BDBADic
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-getBDBADic", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function


