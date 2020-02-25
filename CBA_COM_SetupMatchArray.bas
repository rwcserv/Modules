Attribute VB_Name = "CBA_COM_SetupMatchArray"
Option Explicit
Option Private Module          ' Excel users cannot access procedures


Function CBA_SetupMatchArray(ByVal CBA_UseHistoricalChanges As Boolean, ByVal CBA_datefrom As Date, ByVal CBA_Dateto As Date, Optional ByVal CBA_CG As Long, Optional ByVal CBA_SCG As Long, Optional ByVal CBA_Prodlist As String, Optional ByVal CBA_ModeOnly As Boolean, Optional ByVal CompetitorOnly As String) As Boolean
    Dim CBA_arrprodmap()
    Dim CBA_StartTime As Single, CBA_Timer As Single, b As Long, j As Long, bSameState As Boolean
    Dim CBA_SecondsElapsed As Single, Competitortoquery
    Dim a As Long, bOutput As Boolean, strPProds As String, strSQL As String
    Dim CBA_COM_MMSRS(501 To 509) As ADODB.Recordset
    Dim CCM_MM() As Variant
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    Erase CBA_COM_Match
    CBA_StartTime = CBA_Timer
    '''''****CONNECTION SETUPS****
    Set CBA_COM_CBISCN = New ADODB.Connection
    Set CBA_COM_COMCN = New ADODB.Connection
    With CBA_COM_COMCN
        .ConnectionTimeout = 300
        .CommandTimeout = 300
        .Open "Provider= SQLNCLI10; DATA SOURCE=" & CBA_BasicFunctions.TranslateServerName("599DBL12", Date) & "; ;INTEGRATED SECURITY=sspi;"
    End With
    With CBA_COM_CBISCN
        .ConnectionTimeout = 300
        .CommandTimeout = 300
        .Open "Provider= SQLNCLI10; DATA SOURCE= 599DBL01; ;INTEGRATED SECURITY=sspi;"
    End With
    For a = 501 To 509
        If a <> 508 Then
            Set CBA_COM_MMSCN(a) = New ADODB.Connection
            With CBA_COM_MMSCN(a)
                .ConnectionTimeout = 300
                .CommandTimeout = 300
                .Open "Provider= SQLNCLI10; DATA SOURCE= " & CBA_BasicFunctions.TranslateServerName(a, Date) & "; ;INTEGRATED SECURITY=sspi;"
            End With
        End If
    Next
    CBA_DBtoQuery = 3
    
    If LCase(CompetitorOnly) = "all competitors" Or CompetitorOnly = "" Then
        Competitortoquery = ""
    Else
        If LCase(CompetitorOnly) = "c" Or LCase(CompetitorOnly) = "coles" Then
            Competitortoquery = "coles"
        ElseIf LCase(CompetitorOnly) = "ww" Or LCase(CompetitorOnly) = "woolworths" Then
            Competitortoquery = "ww"
        ElseIf LCase(CompetitorOnly) = "dm" Or LCase(CompetitorOnly) = "dan murphys" Then
            Competitortoquery = "dm"
        ElseIf LCase(CompetitorOnly) = "fc" Or LCase(CompetitorOnly) = "first choice" Then
            Competitortoquery = "fc"
        End If
    End If
    
    
    If CBA_UseHistoricalChanges = False Then
        CCM_MM = CCM_Runtime.CCM_MatchArrayData(CBA_UseHistoricalChanges, CBA_Prodlist, CBA_datefrom, CBA_Dateto, CBA_CG, CStr(CBA_SCG), Competitortoquery)
    '    bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("COM_ClassProdMap", CBA_datefrom, CBA_Dateto, CBA_CG, CStr(CBA_SCG), , CBA_Prodlist)
        On Error Resume Next
        If UBound(CCM_MM, 2) < 0 Then
            bOutput = False
        Else
            bOutput = True
            CBA_arrprodmap = CCM_MM
        End If
        Err.Clear
        On Error GoTo Err_Routine
    End If
    
    If CBA_UseHistoricalChanges = True Then
        bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("COM_ClassProdMapHis", CBA_datefrom, CBA_Dateto, CBA_CG, CStr(CBA_SCG), Competitortoquery, CBA_Prodlist)
        If bOutput = True Then CBA_arrprodmap = CBA_COMarr
    End If
    
    
    
    If bOutput = True Then
        strPProds = ""
        If CBA_UseHistoricalChanges = False Then
            For a = 0 To UBound(CBA_arrprodmap, 2)
                If CBA_arrprodmap(1, a) = "58" Then If strPProds = "" Then strPProds = CBA_arrprodmap(0, a) Else strPProds = strPProds & ", " & CBA_arrprodmap(0, a)
            Next
        Else
            For a = 0 To UBound(CBA_arrprodmap, 2)
                If CBA_arrprodmap(5, a) = "58" Then If strPProds = "" Then strPProds = CBA_arrprodmap(0, a) Else strPProds = strPProds & ", " & CBA_arrprodmap(0, a)
            Next
        End If
        
        If DateDiff("M", CBA_datefrom, CBA_Dateto) > 5 And UBound(CBA_arrprodmap, 2) > 2000 Or _
            DateDiff("M", CBA_datefrom, CBA_Dateto) > 4 And UBound(CBA_arrprodmap, 2) > 3000 Or _
                DateDiff("M", CBA_datefrom, CBA_Dateto) > 3 And UBound(CBA_arrprodmap, 2) > 5000 Or _
                    DateDiff("M", CBA_datefrom, CBA_Dateto) > 2 And UBound(CBA_arrprodmap, 2) > 7000 Or _
                        DateDiff("M", CBA_datefrom, CBA_Dateto) > 1 And UBound(CBA_arrprodmap, 2) > 10000 Or _
                            DateDiff("M", CBA_datefrom, CBA_Dateto) > 0 And UBound(CBA_arrprodmap, 2) > 20000 Then
              MsgBox "Too much data is being queried, please contract Tom on 9218"
              CBA_SetupMatchArray = False
              GoTo Exit_Routine
        End If
        
        If strPProds <> "" Then
            ReDim CBA_COM_APParr(0, 0)
            CBA_COM_APParr(0, 0) = 0
            For a = 501 To 509
                If a <> 508 Then
                    Set CBA_COM_MMSRS(a) = New ADODB.Recordset
                    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10)
                    strSQL = strSQL & "declare @Div Int if substring(@@SERVERNAME,1,1) = 0 set @Div = substring(@@SERVERNAME,2,3) else set @Div = substring(@@SERVERNAME,1,3)" & Chr(10)
                    strSQL = strSQL & "select productcode, State, validfrom, validto, Retail from (" & Chr(10)
                    strSQL = strSQL & "select productcode, case when @Div in (501,504) then 'NSW' else case when @Div in (502,505) then 'VIC' else case when @Div in (503,506) " & Chr(10)
                    strSQL = strSQL & "then 'QLD' else case when @Div = 507 then 'SA' else 'WA' end end end end as State  " & Chr(10)
                    strSQL = strSQL & ", row_number() over (Partition by productcode, case when @Div in (501,504) then 'NSW' else case when @Div in (502,505) then 'VIC' else case when @Div in (503,506) " & Chr(10)
                    strSQL = strSQL & "then 'QLD' else case when @Div = 507 then 'SA' else 'WA' end end end end , validfrom order by retail) as row" & Chr(10)
                    strSQL = strSQL & ", convert(date,validfrom) as validfrom , isnull(convert(date,validto),convert(date,getdate())) as validto,  Retail from purchase.dbo.retail" & Chr(10)
                    strSQL = strSQL & "where isnull(convert(date,validto),convert(date,getdate())) >= '" & Format(CBA_datefrom, "YYYY-MM-DD") & "'" & Chr(10)
                    strSQL = strSQL & "and isnull(convert(date,validfrom),convert(date,getdate())) <= '" & Format(CBA_Dateto, "YYYY-MM-DD") & "'" & Chr(10)
                    strSQL = strSQL & "and productcode in (" & strPProds & ")" & Chr(10)
                    strSQL = strSQL & ") a where a.row = 1 and retail <> 0" & Chr(10)
                    strSQL = strSQL & "order by productcode, state,validfrom" & Chr(10)
                    Debug.Print strSQL
                    CBA_COM_MMSRS(a).Open strSQL, CBA_COM_MMSCN(a) ', , , "adAsyncExecute"
                    If CBA_COM_MMSRS(a).EOF Then
                        ReDim CBA_MMSarr(0, 0)
                        CBA_MMSarr(0, 0) = 0
                    Else
                        CBA_MMSarr = CBA_COM_MMSRS(a).GetRows()
                    End If
                    CBA_COM_MMSRS(a).Close
                    Set CBA_COM_MMSRS(a) = Nothing

                    If CBA_MMSarr(0, 0) <> 0 Then
                        If CBA_COM_APParr(0, 0) = 0 Then
                            Erase CBA_COM_APParr
                            CBA_COM_APParr = CBA_MMSarr
                            Erase CBA_MMSarr
                        Else
                            For j = 0 To UBound(CBA_MMSarr, 2)
                            bSameState = False
                            For b = 0 To UBound(CBA_COM_APParr, 2)
                                If CBA_COM_APParr(0, b) > CBA_MMSarr(0, j) Then Exit For
                                If CBA_COM_APParr(1, b) = CBA_MMSarr(1, j) And CBA_COM_APParr(0, b) = CBA_MMSarr(0, j) Then
                                    bSameState = True
                                    If CBA_MMSarr(4, j) > CBA_COM_APParr(4, b) Then CBA_COM_APParr(4, b) = CBA_MMSarr(4, j)
                                    Exit For
                                End If
                            Next
                            If bSameState = False Then
                            ReDim Preserve CBA_COM_APParr(0 To UBound(CBA_COM_APParr, 1), 0 To UBound(CBA_COM_APParr, 2) + 1)
                            For b = 0 To UBound(CBA_COM_APParr, 1)
    '                            If MMSarr(b, 0) = 76016 Then
    '                                a = a
    '                            End If
                                CBA_COM_APParr(b, UBound(CBA_COM_APParr, 2)) = CBA_MMSarr(b, j)
                            Next
                            End If
                            Next
                            Erase CBA_MMSarr
                        End If
                    End If
                End If
            Next
        End If

        ReDim timearr(0 To 3, 0 To UBound(CBA_arrprodmap, 2))
        For a = 0 To UBound(CBA_arrprodmap, 2)
            If ((a + 1) / 100) = Round(((a + 1) / 100), 0) Then
                If CBA_BasicFunctions.isRunningSheetDisplayed = True Then
                    CBA_BasicFunctions.RunningSheetAddComment 7, 5, "Building Match Objects : " & Format((a + 1) / UBound(CBA_arrprodmap, 2), "0.00%") & " Complete"
                Else
                    CBA_BasicFunctions.CBA_Running "Generating DataCube..."
                    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 7, 5, "Building Match Objects : " & Format((a + 1) / UBound(CBA_arrprodmap, 2), "0.00%") & " Complete"
                    'CBA_BasicFunctions.CBA_Running
                End If
        '      '  Debug.Print Format((a + 1) / UBound(CBA_arrprodmap, 2), "0.00%") '& Chr(10) & "Est to complete: " & Round(((Round(CBA_Timer - CBA_StartTime, 2)) / ((a + 1) / UBound(CBA_arrprodmap, 2))) - (Round(CBA_Timer - CBA_StartTime, 2)), 0) & "s"
            End If
            DoEvents
            ReDim Preserve CBA_COM_Match(1 To a + 1)
            timearr(0, a) = CBA_arrprodmap(0, a)
            timearr(1, a) = CBA_arrprodmap(3, a)
            timearr(2, a) = CBA_arrprodmap(4, a)
            timearr(3, a) = Timer
        
            If CBA_UseHistoricalChanges = False Then
                CBA_COM_Match(a + 1).CBA_COM_Formulate CBA_arrprodmap(4, a), CBA_arrprodmap(0, a), CBA_arrprodmap(3, a), CBA_datefrom, CBA_Dateto, CBA_ModeOnly
            Else
                CBA_COM_Match(a + 1).CBA_COM_Formulate CBA_arrprodmap(4, a), CBA_arrprodmap(0, a), CBA_arrprodmap(1, a), CBA_arrprodmap(2, a), CBA_arrprodmap(3, a), CBA_ModeOnly
            End If
            timearr(3, a) = Timer - timearr(3, a)
        Next
        
        CBA_SetupMatchArray = True
    Else
        CBA_SetupMatchArray = False
    End If
   
Exit_Routine:
    On Error Resume Next
    Erase CBA_COM_APParr
    Erase CBA_COM_AParr
    CBA_COM_COMCN.Close
    Set CBA_COM_COMCN = Nothing
    CBA_COM_CBISCN.Close
    Set CBA_COM_CBISCN = Nothing
    For a = 501 To 509
        If a <> 508 Then
            CBA_COM_MMSCN(a).Close
            Set CBA_COM_MMSCN(a) = Nothing
        End If
    Next
    CBA_SecondsElapsed = Round(CBA_Timer - CBA_StartTime, 2)

    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_SetupMatchArray", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    ''If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next


End Function

