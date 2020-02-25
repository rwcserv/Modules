Attribute VB_Name = "mCBA_Nielsen"
Option Explicit                                     ' mCBA_Nielsen

Private Sub TESTNIELSENDATA()
    Dim dic As Scripting.Dictionary ''', CGSCG As Scripting.Dictionary
    Dim v As Variant
    Dim ND As cCBA_NielsenData
    
    ''FOR HomeScan
    'Set CGSCG = New Scripting.Dictionary
    'CGSCG.Add CStr(5), CStr(0)
    'Set Dic = getNielsenData(2018, eSep, False, CGSCG, False)
    'For Each v In Dic("5")
    '    Set ND = Dic("5")(v)
    '    Debug.Print ND.CategoryALDIShare
    'Next
    
    ''FOR ScanData
    'Set dic = getNielsenData(2019, eSep, True, , , "FOOD - AMBIENT")
    Set dic = GetNielsenData(2019, eSep, True)
    If dic Is Nothing Then Exit Sub
    For Each v In dic
        Set ND = dic(v)
        Debug.Print ND.Sales
    Next
End Sub
Function GetNielsenSegmentationData() As Scripting.Dictionary
    Dim CN As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim ND As cCBA_NielsenData
    Dim dic As Scripting.Dictionary
    Dim strSQL As String
    On Error GoTo Err_Routine
    CBA_Error = ""
    Set dic = New Scripting.Dictionary
    Set CN = New ADODB.Connection
    With CN
        .ConnectionTimeout = 50
        .CommandTimeout = 50
        .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Cam")  ' & CBA_BSA & "LIVE DATABASES\NielsenData.accdb"
    End With
    Set RS = New ADODB.Recordset
    strSQL = "SELECT h.HS_ACG, h.HS_CGNo, h.HS_SCGNo, Format(h.HS_CGNo,'000')+Format(h.HS_SCGno,'000') AS MSegName, ca.CA_CAT_ID, cn.CN_CategoryName" & Chr(10)
    strSQL = strSQL & "FROM (N0_HomeScan AS h LEFT JOIN L0_CategoryAllocations AS ca ON (h.HS_SCGno = ca.CA_LCGSCGNo) AND (h.HS_CGno = ca.CA_LCGCGNo)) LEFT JOIN L0_CategoryName AS cn ON ca.CA_CAT_ID = cn.CN_ID" & Chr(10)
    strSQL = strSQL & "WHERE (((h.HS_ACG) = False))" & Chr(10)
    strSQL = strSQL & "GROUP BY h.HS_ACG, h.HS_CGNo, h.HS_SCGNo, ca.CA_CAT_ID, cn.CN_CategoryName" & Chr(10)
    strSQL = strSQL & "Union" & Chr(10)
    strSQL = strSQL & "SELECT h.HS_ACG, h.HS_CGNo, h.HS_SCGNo, Format(h.HS_CGNo,'000')+Format(h.HS_SCGno,'000') AS MSegName, ca.CA_CAT_ID, cn.CN_CategoryName" & Chr(10)
    strSQL = strSQL & "FROM (N0_HomeScan AS h LEFT JOIN L0_CategoryAllocations AS ca ON (h.HS_SCGno = ca.CA_ACGSCGNo) AND (h.HS_CGno = ca.CA_ACGCGNo)) LEFT JOIN L0_CategoryName AS cn ON ca.CA_CAT_ID = cn.CN_ID" & Chr(10)
    strSQL = strSQL & "WHERE (((h.HS_ACG) = True))" & Chr(10)
    strSQL = strSQL & "GROUP BY h.HS_ACG, h.HS_CGNo, h.HS_SCGNo, ca.CA_CAT_ID, cn.CN_CategoryName;" & Chr(10)
    RS.Open strSQL, CN
    Do Until RS.EOF
        Set ND = New cCBA_NielsenData
        ND.ACG = CBool(RS(0))
        ND.CGno = CLng(RS(1))
        ND.SCGno = CLng(RS(2))
        ND.IsHomescan = True
        ND.MSegDescription = CStr(RS(3))
        ND.Category_ID = CLng(NZ(RS(4), 0)) 'IF Cat_ID = 0 then the segment is not allocated to a CAMERA Category
        ND.SelectedForCategory = CStr(NZ(RS(5), ""))
''        dic.Add CStr(RS(3)), ND
        dic.Add CStr(RS(3)) & Format(NZ(RS(4), 0), "000"), ND
        RS.MoveNext
    Loop
    Set RS = Nothing: Set RS = New ADODB.Recordset
    strSQL = "SELECT SDH.SH_ID, SDH.SH_Desc, SDA.SA_CN_ID, cn.CN_CategoryName" & Chr(10)
    strSQL = strSQL & "FROM (N0_ScanDataHeaders AS SDH LEFT JOIN L1_ScanDataAllocation AS SDA ON SDH.SH_ID = SDA.SA_SH_ID) LEFT JOIN L0_CategoryName AS cn ON SDA.SA_CN_ID = cn.CN_ID;" & Chr(10)
    RS.Open strSQL, CN
    Do Until RS.EOF
        Set ND = New cCBA_NielsenData
        ND.isScanData = True
        ND.H_ID = CLng(RS(0))
        ND.MSegDescription = CStr(RS(1))
        ND.Category_ID = CLng(NZ(RS(2), 0))
        ND.SelectedForCategory = CStr(NZ(RS(3), ""))
        ''dic.Add CStr(RS(1)), ND
        dic.Add CStr(RS(1)) & Format(NZ(RS(2), 0), "000"), ND
        RS.MoveNext
    Loop
    Set RS = Nothing: Set RS = New ADODB.Recordset
    strSQL = "SELECT L1_ManualSegments.MS_CN_ID, L1_ManualSegments.MS_ManualMSegName, cn.CN_CategoryName" & Chr(10)
    strSQL = strSQL & "FROM L1_ManualSegments LEFT JOIN L0_CategoryName AS cn ON L1_ManualSegments.MS_CN_ID = cn.CN_ID" & Chr(10)
    strSQL = strSQL & "WHERE (((L1_ManualSegments.MS_Isactive)=True));" & Chr(10)
    RS.Open strSQL, CN
    Do Until RS.EOF
        Set ND = New cCBA_NielsenData
        ND.IsManual = True
        ND.MSegDescription = CStr(RS(1))
        ND.Category_ID = CLng(RS(0))
        ND.SelectedForCategory = CStr(NZ(RS(2), ""))
'        dic.Add CStr(RS(1)), ND
        dic.Add CStr(RS(1)) & Format(NZ(RS(2), 0), "000"), ND
        RS.MoveNext
    Loop
    Set RS = Nothing
    CN.Close
    Set CN = Nothing
    Set GetNielsenSegmentationData = dic
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-mCBA_Nielsen.GetNielsenSegmentationData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error & vbCrLf & strSQL
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Function GetNielsenData(ByVal iYearNo As Integer, ByVal lMonthNo As Long, ByVal isScanData As Boolean, Optional ByVal CGSCG As Collection, _
                        Optional ACG As Boolean, Optional ScanDataSegmentDescription As String) As Scripting.Dictionary
    Dim dic As Scripting.Dictionary, subdic As Scripting.Dictionary
    Dim CN As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim ND As cCBA_NielsenData
    Dim sCurSCG As String, strSQL As String, strCGSQL As String, sCurCG As String
    Dim v As Variant
    On Error GoTo Err_Routine
    CBA_Error = ""
    If isScanData = True And ScanDataSegmentDescription = "" Then
        Debug.Print "ScanDataSegmentDescription required when pulling ScanData"
        Set GetNielsenData = Nothing
        Exit Function
    End If
    If (isScanData = False And CGSCG Is Nothing) Then
        Debug.Print "CGno  required when pulling HomeScan"
        Set GetNielsenData = Nothing
        Exit Function
    End If
    
    Set CN = New ADODB.Connection
    With CN
        .ConnectionTimeout = 50
        .CommandTimeout = 50
        .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Cam")  ' & CBA_BSA & "LIVE DATABASES\NielsenData.accdb"
    End With

    If isScanData = True Then
        strSQL = "Select N0_ScanDataHeaders.*, N0_ScanData.* from N0_ScanData " & Chr(10)
        strSQL = strSQL & "left join N0_ScanDataHeaders on N0_ScanDataHeaders.SH_ID = N0_ScanData.SD_SH_ID " & Chr(10)
        strSQL = strSQL & "where N0_ScanDataHeaders.SH_Desc = '" & ScanDataSegmentDescription & "'" & Chr(10)
        strSQL = strSQL & "and N0_ScanData.SD_MonthNo = " & lMonthNo & Chr(10)
        strSQL = strSQL & "and N0_ScanData.SD_YearNo = " & iYearNo & Chr(10)
    
'        strSQL = "Select ScanDataHeaders.*, ScanData.* from ScanData " & Chr(10)
'        strSQL = strSQL & "left join ScanDataHeaders on ScanDataHeaders.SH_ID = ScanData.SD_SH_ID " & Chr(10)
'        strSQL = strSQL & "where ScanDataHeaders.SH_Desc = '" & ScanDataSegmentDescription & "'" & Chr(10)
'        strSQL = strSQL & "and ScanData.SD_MonthNo = " & lMonthNo & Chr(10)
'        strSQL = strSQL & "and ScanData.SD_YearNo = " & iYearNo & Chr(10)
    Else
        strSQL = "Select * from N0_HomeScan" & Chr(10)
'        strSQL = "Select * from HomeScan" & Chr(10)
        strSQL = strSQL & "where HS_MonthNo = " & lMonthNo & Chr(10)
        strSQL = strSQL & "and HS_YearNo = " & iYearNo & Chr(10)
        strSQL = strSQL & "and HS_ACG = " & ACG & Chr(10)
        strCGSQL = ""
        For Each v In CGSCG
            If Right(v, 2) = "00" Then
''                strCGSQL = strCGSQL & IIf(strCGSQL = "", " and (", " or ") & "(CGNo = " & v & ")"
                strCGSQL = strCGSQL & IIf(strCGSQL = "", " and ((", ") or (") & "(HS_CGNo = " & Left(v, 3) & ")"
            Else
''                strCGSQL = strCGSQL & IIf(strCGSQL = "", " and (", " or ") & " and CGNo = " & v & " and SCGNo = " & CGSCG(v)
                strCGSQL = strCGSQL & IIf(strCGSQL = "", " and ((", ") or (") & " HS_CGNo = " & Left(v, 3) & " and HS_SCGNo = " & Right(v, 2)
            End If
        Next
        strCGSQL = strCGSQL & "))" & Chr(10)
        strSQL = strSQL & strCGSQL
    End If
    Set RS = New ADODB.Recordset
    RS.Open strSQL, CN
    
    Set dic = New Scripting.Dictionary
    If isScanData = False Then                          ' IS HOMESCAN
        If Not RS.EOF Then sCurCG = CStr(RS(1)): Set subdic = New Scripting.Dictionary
        Do Until RS.EOF
            If sCurCG <> CStr(RS(1)) Then dic.Add CStr(RS(1)), subdic: Set subdic = New Scripting.Dictionary: sCurCG = CStr(RS(1))
            Set ND = New cCBA_NielsenData
            ND.MonthNo = lMonthNo
            ND.YearNo = iYearNo
            ND.Retail = CSng(RS(3))
            ND.YOYRetail = CSng(RS(4))
            ND.QTY = CSng(RS(5))
            ND.YOYQTY = CSng(RS(6))
            ND.MeasureSales = CSng(RS(7))
            ND.YOYMeasureSales = CSng(RS(8))
            ND.CategoryMarketShare = CSng(RS(9))
            ND.CategoryALDIShare = CSng(RS(10))
            ND.MarketPLShare = CSng(RS(11))
            ND.ALDIPLShare = CSng(RS(12))
            ND.SOTRetail = CSng(RS(13))
            ND.SOTQTY = CSng(RS(14))
            ND.SOTMeasureSales = CSng(RS(15))
            ND.ACG = CSng(RS(16))
            ND.CGno = CStr(RS(1))
            ND.SCGno = CLng(RS(2))
            subdic.Add CStr(RS(2)), ND
            RS.MoveNext
        Loop
        If sCurCG <> "" Then
            dic.Add CStr(sCurCG), subdic
            Set GetNielsenData = dic
        Else
            Set GetNielsenData = Nothing
        End If
    Else                                                ' IS SCANDATA
        Do Until RS.EOF
            Set ND = New cCBA_NielsenData
            ND.MonthNo = lMonthNo
            ND.YearNo = iYearNo
            ND.TotGComporALDI = CStr(RS(6))
            ND.Sales = CSng(RS(7))
            ND.SalesYOY = CSng(RS(8))
            ND.SalesKG = CSng(RS(9))
            ND.SalesKGYOY = CSng(RS(10))
            ND.SalesQTY = CSng(RS(11))
            ND.SalesQTYYOY = CSng(RS(12))
            ND.ShareSales = CSng(RS(13))
            ND.ShareSalesYOY = CSng(RS(14))
            ND.ShareSalesKG = CSng(RS(15))
            ND.ShareSalesKGYOY = CSng(RS(16))
            ND.ShareSalesQTY = CSng(RS(17))
            ND.ShareSalesQTYYOY = CSng(RS(18))
            ND.SD_ID = CSng(RS(2))
            ND.H_ID = CSng(RS(3))
            ND.MSegDescription = CStr(RS(1))
            dic.Add CStr(RS(6)), ND
            RS.MoveNext
        Loop
        If ND Is Nothing Then Else Set GetNielsenData = dic
    End If
Exit_Routine:

    On Error Resume Next
    CN.Close
    Set CN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-mCBA_Nielsen.GetNielsenData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function

Sub ImportNielsenHomescanData()
    Dim StrFile As String, strSQL As String, ImportedFiles As String
    Dim WB As Workbook
    Dim sht As Worksheet, ImpSht As Worksheet
    Dim CN As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim lstRow As Long, a As Long, b As Long, c As Long
    Dim fso As Scripting.FileSystemObject
    On Error GoTo Err_Routine
    CBA_Error = ""
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
    Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False
    StrFile = Dir(CBA_BSA & "LIVE DATABASES\Nielsen Data Import\Homescan\")
    Do While Len(StrFile) > 0
                
        Set WB = Workbooks.Open(CBA_BSA & "LIVE DATABASES\Nielsen Data Import\Homescan\" & StrFile, , True)
        For Each sht In WB.Worksheets
            With sht
                If .Cells(2, 3).Value = "ALDI CATEGORY REPORT" Then
                    Set ImpSht = sht
                    Exit For
                End If
            End With
        Next
        
        If ImpSht Is Nothing Then
        Else
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Importing Data from " & StrFile
            Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False
            Set CN = New ADODB.Connection
            With CN
                .ConnectionTimeout = 50
                .CommandTimeout = 50
                .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Cam")  ' & CBA_BSA & "LIVE DATABASES\NielsenData.accdb"
            End With
            With ImpSht
                lstRow = .Cells(4, 2).End(xlDown).Row
                For a = 5 To lstRow
                    If .Cells(a, 2).Value = "" Then GoTo TryNextOne
                    Set RS = New ADODB.Recordset
                    strSQL = "INSERT INTO N0_HomeScan (HS_Category,HS_CGno,HS_SCGNo,HS_Retail,HS_YOYRetail,HS_QTY,HS_YOYQTY,HS_MeasureSales,HS_YOYMeasureSales,HS_CategoryMarketShare," & _
                            "HS_CategoryALDIShare,HS_MarketPLShare,HS_ALDIPLShare,HS_SOTRetail,HS_SOTQTY,HS_SOTMeasureSales,HS_MonthNo,HS_YearNo,HS_ACG)" & Chr(10)
                    strSQL = strSQL & "Values("
                    For b = 2 To 19
                        Select Case b
                         Case 2
                            strSQL = strSQL & "'" & .Cells(a, b).Value & "'"
                         Case 3
                            If Len(.Cells(a, b).Value) = 7 Then
                                strSQL = strSQL & "," & Mid(.Cells(a, b).Value, 1, 1)
                                strSQL = strSQL & "," & Mid(.Cells(a, b).Value, 2, 2)
                            ElseIf Len(.Cells(a, b).Value) = 8 Then
                                strSQL = strSQL & "," & Mid(.Cells(a, b).Value, 1, 2)
                                strSQL = strSQL & "," & Mid(.Cells(a, b).Value, 3, 2)
                            ElseIf Len(.Cells(a, b).Value) = 9 Then
                                strSQL = strSQL & "," & Mid(.Cells(a, b).Value, 1, 3)
                                strSQL = strSQL & "," & Mid(.Cells(a, b).Value, 4, 2)
                            End If
                         Case 18
                            strSQL = strSQL & "," & Month(.Cells(1, 4).Value)
                         Case 19
                            strSQL = strSQL & "," & Year(.Cells(1, 4).Value)
                         Case 4
                        
                         Case Else
                            If .Cells(a, b).Value = "NA" Then
                                .Cells(a, b).Value = 0
                            End If
                            strSQL = strSQL & "," & .Cells(a, b).Value
                        End Select
                    Next
                    strSQL = strSQL & ",False)" & Chr(10)
                    RS.Open strSQL, CN
TryNextOne:
                Next
            End With
        End If
        Set fso = New Scripting.FileSystemObject
        fso.MoveFile Source:=CBA_BSA & "LIVE DATABASES\Nielsen Data Import\Homescan\" & StrFile, _
            Destination:=CBA_BSA & "LIVE DATABASES\Nielsen Data Import\Homescan\Imported\" & StrFile
        WB.Close
        ImportedFiles = ImportedFiles & ImportedFiles & StrFile & Chr(10)
        Set ImpSht = Nothing
        StrFile = Dir
            
    Loop
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    Application.ScreenUpdating = True: DoEvents
    If ImportedFiles = "" Then
        MsgBox "Import NOT Completed ", vbOKOnly
    Else
        MsgBox "Import Complete " & Chr(10) & ImportedFiles, vbOKOnly
    End If
Exit_Routine:

    On Error Resume Next
    CN.Close
    Set CN = Nothing
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-mCBA_Nielsen.ImportNielsenHomescanData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error & vbCrLf & strSQL
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Sub ImportNielsenScanDataData()
    Dim StrFile As String, strSQL As String, ImportedFiles As String
    Dim WB As Workbook
    Dim sht As Worksheet, ImpSht As Worksheet
    Dim CN As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim lstRow As Long, a As Long, b As Long, c As Long, stNo As Long
    Dim fso As Scripting.FileSystemObject
    Dim ImportFields() As Scripting.Dictionary, HeadNameDic As Scripting.Dictionary
    Dim ImpDate As Date
    
    On Error GoTo Err_Routine
    CBA_Error = ""
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
    Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False
    StrFile = Dir(CBA_BSA & "LIVE DATABASES\Nielsen Data Import\Scandata\")
    Do While Len(StrFile) > 0
        Set WB = Workbooks.Open(CBA_BSA & "LIVE DATABASES\Nielsen Data Import\Scandata\" & StrFile)
        For Each sht In WB.Worksheets
            With sht
                If .Cells(16, 7).Value = "Val Sales (in MM AUD)" Then
                    ImpDate = CDate(Right(.Cells(6, 7).Value, 8))
                    Set ImpSht = sht
                    Exit For
                End If
            End With
        Next
        
        If ImpSht Is Nothing Then
        Else
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 2, 2, "Importing Data from " & StrFile
            Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False
            Set CN = New ADODB.Connection
            With CN
                .ConnectionTimeout = 50
                .CommandTimeout = 50
                .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Cam")  ' & CBA_BSA & "LIVE DATABASES\NielsenData.accdb"
            End With
            With ImpSht
                lstRow = .Cells(17, 5).End(xlDown).Row
                Set HeadNameDic = New Scripting.Dictionary
                Set RS = New ADODB.Recordset
                strSQL = "select SH_Desc, SH_ID from N0_ScanDataHeaders"
                RS.Open strSQL, CN
                Do Until RS.EOF
                    HeadNameDic.Add CStr(RS(0)), CStr(RS(1))
                    RS.MoveNext
                Loop
                For a = 17 To lstRow
                    If HeadNameDic.Exists(.Cells(a, 5).Value) = False Then
                        Set RS = New ADODB.Recordset
                        strSQL = "insert into N0_ScanDataHeaders(SH_Desc) VALUES('" & .Cells(a, 5).Value & "')"
                        RS.Open strSQL, CN
                        Set HeadNameDic = New Scripting.Dictionary
                        Set RS = New ADODB.Recordset
                        strSQL = "select SH_Desc, SH_ID from N0_ScanDataHeaders"
                        RS.Open strSQL, CN
                        Do Until RS.EOF
                            HeadNameDic.Add CStr(RS(0)), str(RS(1))
                        Loop
                    End If
                    For c = 1 To 3
                        Set RS = New ADODB.Recordset
                        strSQL = "INSERT INTO N0_ScanData (SD_SH_ID,MonthNo,SD_YearNo,SD_TotGCompORALDI,[SD_Sales$],[SD_Sales$YOY],SD_SalesKG,SD_SalesKGYOY," & _
                                 "SD_SalesQTY,SD_SalesQTYYOY,[SD_ShareSales$],[SD_ShareSales$YOY], SD_ShareSalesKG, SD_ShareSalesKGYOY,SD_ShareSalesQTY,SD_ShareSalesQTYYOY)" & Chr(10)
                        strSQL = strSQL & "Values(" & HeadNameDic(.Cells(a, 5).Value) & "," & Month(ImpDate) & "," & Year(ImpDate)
                        If c = 1 Then stNo = 6: strSQL = strSQL & ", 'TotG'"
                        If c = 2 Then stNo = 18: strSQL = strSQL & ", 'Comp'"
                        If c = 3 Then stNo = 30: strSQL = strSQL & ", 'ALDI'"
                        For b = 1 To 12
                            strSQL = strSQL & "," & IIf(.Cells(a, stNo + b).Value = "", 0, .Cells(a, stNo + b).Value)
                        Next
                        strSQL = strSQL & ")" & Chr(10)
                        RS.Open strSQL, CN
                    Next
                Next
            End With
        End If
        WB.Close
        
        
        Set fso = New Scripting.FileSystemObject
        fso.MoveFile Source:=CBA_BSA & "LIVE DATABASES\Nielsen Data Import\Scandata\" & StrFile, _
            Destination:=CBA_BSA & "LIVE DATABASES\Nielsen Data Import\Scandata\Imported\" & StrFile
        ImportedFiles = ImportedFiles & ImportedFiles & StrFile & Chr(10)
        Set ImpSht = Nothing
        StrFile = Dir
    Loop
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    Application.ScreenUpdating = True: DoEvents
    MsgBox "Import Complete " & Chr(10) & ImportedFiles, vbOKOnly
Exit_Routine:

    On Error Resume Next
    CN.Close
    Set CN = Nothing
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-mCBA_Nielsen.ImportNielsenScanDataData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub

