Attribute VB_Name = "CBA_POSQuery"
Option Explicit
Option Private Module          ' Excel users cannot access procedures
Private CBA_POSProductcode As Long
Private CBA_POSDateFrom As Date
Private CBA_POSDateTo As Date
Private CBA_POSRunEnable As Boolean
Private CBA_POSRunUseLinkedProducts As Boolean
Private CBA_POS_RegionCol As Collection
Private CBA_OSD() As Date
Private CBA_OSDSelected As Date
'Callback for CBA_OSDSelector onAction
Sub CBA_OSDonAction(Control As IRibbonControl, id As String, iIndex As Integer)
    CBA_OSDSelected = CBA_OSD(iIndex + 1)
    CBA_Rib.InvalidateControl "CBA_BT_UnrealisedRevenue"
End Sub
'Callback for CBA_OSDSelector getItemCount
Sub CBA_OSDgetItemCount(Control As IRibbonControl, ByRef returnedVal)
Dim wedDate As Date, satDate As Date, selDate As Date
    On Error Resume Next
    If CBA_OSD(1) = "12:00:00 AM" Then
        wedDate = DateAdd("D", -WeekDay(Date, vbThursday), Date)
        satDate = DateAdd("D", -WeekDay(Date, vbSunday), Date)
        If DateDiff("D", wedDate, satDate) > 0 Then selDate = satDate Else selDate = wedDate
        CBA_OSD = getOSDs(DateAdd("WW", -105, selDate), DateAdd("WW", 53, selDate))
    End If
    Err.Clear
    On Error GoTo 0
    returnedVal = 211
End Sub
'Callback for CBA_OSDSelector getItemLabel
Sub CBA_OSDgetItemLabel(Control As IRibbonControl, iIndex As Integer, ByRef returnedVal)
    returnedVal = CBA_OSD(iIndex + 1)
End Sub
Function getOSDs(ByVal StartDate As Date, EndDate As Date) As Date()
    Dim TempArr(1 To 211) As Date, a As Long, DT As Date
    a = 0
    For DT = StartDate To EndDate
        If WeekDay(DT, vbWednesday) = 1 Or WeekDay(DT, vbSaturday) = 1 Then
            a = a + 1
            TempArr(a) = DT
            If a = 211 Then
                Exit For
            End If
        End If
    Next
    getOSDs = TempArr
End Function
Function getCBA_OSDSelected() As Date
    getCBA_OSDSelected = CBA_OSDSelected
End Function
Function CBA_getPOSProductcode() As Date
    CBA_getPOSProductcode = CBA_POSProductcode
End Function
'Callback for CBA_RunInventory onAction
Sub CBA_InventoryQry(Control As IRibbonControl)
    Dim CBA_curDate As Date, bOutput As Boolean, CBA_Reg, bNewSht, OPwb, OPSht, a As Long, CBA_Row, CBA_Num, CBA_Col
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    If CBA_POSProductcode = 0 Or IsEmpty(CBA_POSDateFrom) Or IsEmpty(CBA_POSDateTo) Then
        MsgBox "Please enter values into all inputboxes", vbOKOnly
        Exit Sub
    End If
    CBA_DBtoQuery = 599
    'CBA_col = 0
    bNewSht = False
    
    bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBIS_INVDATA", CBA_POSDateFrom, CBA_POSDateTo, CBA_POSProductcode, CBA_Reg, , , , , , , CBA_POSRunUseLinkedProducts)
    If bOutput = True Then
        If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
        Application.ScreenUpdating = False
        bNewSht = True
        Set OPwb = Application.Workbooks.Add
        Set OPSht = OPwb.Sheets(1)
        Application.Calculation = xlCalculationManual
        With OPSht
            .Cells(2, 1).Value = "Date"
            .Cells(1, 2).Value = "Minchinbury"
            Range(.Cells(1, 2), .Cells(1, 3)).Merge
            .Cells(1, 4).Value = "Derrimut"
            Range(.Cells(1, 4), .Cells(1, 5)).Merge
            .Cells(1, 6).Value = "Stapylton"
            Range(.Cells(1, 6), .Cells(1, 7)).Merge
            .Cells(1, 8).Value = "Prestons"
            Range(.Cells(1, 8), .Cells(1, 9)).Merge
            .Cells(1, 10).Value = "Dandenong"
            Range(.Cells(1, 10), .Cells(1, 11)).Merge
            .Cells(1, 12).Value = "Brendale"
            Range(.Cells(1, 12), .Cells(1, 13)).Merge
            .Cells(1, 14).Value = "Regency Park"
            Range(.Cells(1, 14), .Cells(1, 15)).Merge
            .Cells(1, 16).Value = "Jandakot"
            Range(.Cells(1, 16), .Cells(1, 17)).Merge
            .Cells(1, 18).Value = "National"
            Range(.Cells(1, 18), .Cells(1, 19)).Merge
            For a = 1 To 9
            .Cells(2, a + (1 * a)).Value = "StoreQTY"
            .Cells(2, a + 1 + (1 * a)).Value = "WH QTY"
            Next
            Cells.Rows(1).HorizontalAlignment = xlCenter
            CBA_Row = 2
            For CBA_Num = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                If CBA_curDate <> CBA_CBISarr(1, CBA_Num) Then
                    CBA_Row = CBA_Row + 1
                    CBA_curDate = CBA_CBISarr(1, CBA_Num)
                    .Cells(CBA_Row, 1).Value = CBA_curDate
                    .Cells(CBA_Row, 18).Value = "=SUM(RC[-16],RC[-14],RC[-12],RC[-10],RC[-8],RC[-6],RC[-4],RC[-2])"
                    .Cells(CBA_Row, 19).Value = "=SUM(RC[-16],RC[-14],RC[-12],RC[-10],RC[-8],RC[-6],RC[-4],RC[-2])"
                End If
                If CBA_CBISarr(0, CBA_Num) = 509 Then CBA_Col = 16 Else CBA_Col = (CBA_CBISarr(0, CBA_Num) - 500) * 2
                .Cells(CBA_Row, CBA_Col).Value = CBA_CBISarr(2, CBA_Num)
                .Cells(CBA_Row, CBA_Col + 1).Value = CBA_CBISarr(3, CBA_Num)
            Next
            For a = 2 To 19
            OPSht.Columns(a).NumberFormat = "#,0"
            Next
            OPSht.Columns(1).NumberFormat = "DD/MM/YYYY"
        End With
    End If
    
    If bNewSht = True Then
        OPSht.Rows(1).EntireRow.Font.Bold = True
        OPSht.Rows(2).EntireRow.Font.Bold = True
        Range(OPSht.Cells(2, 1), OPSht.Cells(2, 19)).AutoFilter
        OPSht.Cells(2, 1).CurrentRegion.EntireColumn.AutoFit
    End If
    CBA_BasicFunctions.CBA_Close_Running
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_InventoryQry", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub
'Callback for CBA_RunStoreInventory onAction
Sub CBA_StoreInventoryQry(Control As IRibbonControl)
    Dim thisarr() As Variant, lRow As Long, bNewSht As Boolean, OPwb, OPSht, vDiv, bOutput As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    If CBA_POSProductcode = 0 Or IsEmpty(CBA_POSDateFrom) Or IsEmpty(CBA_POSDateTo) Then
        MsgBox "Please enter/re-enter values into all inputboxes", vbOKOnly
        Exit Sub
    End If
    
    If Abs(DateDiff("M", CBA_POSDateFrom, CBA_POSDateTo)) > 3 Then
        MsgBox "Due to the amount of data involved, you can only query a maximum of 3 months of store pos data"
        Exit Sub
    End If
    
    Set CBA_POS_RegionCol = New Collection
    CBA_POS_StoreSalesForm.Show vbModal
    If CBA_POS_RegionCol.Count > 0 Then
        lRow = 2
        If bNewSht = False Then
'            If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
'            Application.ScreenUpdating = False
            Set OPwb = Application.Workbooks.Add
            Set OPSht = OPwb.Sheets(1)
'            Application.Calculation = xlCalculationManual
            bNewSht = True
        End If
        OPSht.Cells(1, 1).Value = "Region"
        OPSht.Cells(1, 2).Value = "Sales Date"
        OPSht.Cells(1, 3).Value = "Store Name"
        OPSht.Cells(1, 4).Value = "Quantity"
        'OPSht.Cells(1, 5).Value = "Retail"
        For Each vDiv In CBA_POS_RegionCol
            bOutput = CBA_SQL_Queries.CBA_GenPullSQL("MMS_StoreSalesDATA", CBA_POSDateFrom, CBA_POSDateTo, CBA_POSProductcode, vDiv, , , , , , , CBA_POSRunUseLinkedProducts)
            thisarr = CBA_MMSarr
            Erase CBA_MMSarr
            thisarr = CBA_BasicFunctions.CBA_TransposeArray(thisarr)
            Range(OPSht.Cells(lRow, 1), OPSht.Cells(lRow + UBound(thisarr, 1), UBound(thisarr, 2) + 1)).Value2 = thisarr
            lRow = lRow + UBound(thisarr, 1) + 1
        Next
'        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
        OPSht.Cells(1, 2).EntireColumn.NumberFormat = "DD/MM/YYYY"
        OPSht.Cells.EntireColumn.AutoFit
'        Application.Calculation = xlCalculationAutomatic
'        Application.ScreenUpdating = True
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_StoreInventoryQry", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Sub CBA_POSQry(Control As IRibbonControl)
    Dim bNewSht As Boolean
    Dim OPwb As Workbook
    Dim OPSht As Worksheet
    Dim lNum As Long, lcol As Long, lRow As Long, lReg As Long, bOutput As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    If CBA_POSProductcode = 0 Or IsEmpty(CBA_POSDateFrom) Or IsEmpty(CBA_POSDateTo) Then
        MsgBox "Please enter values into all inputboxes", vbOKOnly
        Exit Sub
    End If
    CBA_DBtoQuery = 599
    lcol = 0
    bNewSht = False
    For lReg = 501 To 599
        Select Case lReg
            Case 501 To 507, 509, 599
                bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBIS_POSDATA", CBA_POSDateFrom, CBA_POSDateTo, CBA_POSProductcode, lReg, , , , , , , CBA_POSRunUseLinkedProducts)
                If bOutput = True Then
                    Application.ScreenUpdating = False
                    lcol = lcol + 2
                    If bNewSht = False Then
                        If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
                        Application.ScreenUpdating = False
                        Set OPwb = Application.Workbooks.Add
                        Set OPSht = OPwb.Sheets(1)
                        Application.Calculation = xlCalculationManual
                        bNewSht = True
                    End If
                    With OPSht
                        If lcol = 2 Then OPSht.Cells(1, 1).Value = "Date"
                        OPSht.Cells(1, lcol).Value = CBA_BasicFunctions.CBA_DivtoReg(lReg)
                        OPSht.Cells(1, lcol + 1).Value = OPSht.Cells(1, lcol).Value
                        lRow = 1
                        For lNum = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                            lRow = lRow + 1
                            If lcol = 2 Then OPSht.Cells(lRow, 1).Value = CBA_CBISarr(0, lNum)
                            OPSht.Cells(lRow, lcol).Value = CBA_CBISarr(1, lNum)
                            OPSht.Cells(lRow, lcol + 1).Value = CBA_CBISarr(2, lNum)
                        Next
                        OPSht.Columns(lcol).NumberFormat = "#,0"
                        OPSht.Columns(lcol + 1).NumberFormat = "$#,0.00"
                    End With
                End If
        End Select
    Next
    If lcol > 0 Then
        OPSht.Rows(1).EntireRow.Font.Bold = True
        OPSht.Cells(1, 1).CurrentRegion.AutoFilter
        OPSht.Cells(1, 1).CurrentRegion.EntireColumn.AutoFit
    End If
    CBA_BasicFunctions.CBA_Close_Running
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_POSQry", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
'Callback for CBA_RunStoreSales onAction
Sub CBA_StoreSalesQry(Control As IRibbonControl)
    Dim thisarr() As Variant, lRow As Long, bNewSht As Boolean, OPwb, OPSht, vDiv, bOutput As Boolean, wks_Piv
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    If CBA_POSProductcode = 0 Or IsEmpty(CBA_POSDateFrom) Or IsEmpty(CBA_POSDateTo) Then
        MsgBox "Please enter/re-enter values into all inputboxes", vbOKOnly
        Exit Sub
    End If
    
    If Abs(DateDiff("M", CBA_POSDateFrom, CBA_POSDateTo)) > 3 Then
        MsgBox "Due to the amount of data involved, you can only query a maximum of 3 months of store pos data"
        Exit Sub
    End If
    Set CBA_POS_RegionCol = New Collection
    CBA_POS_StoreSalesForm.Show vbModal
    If CBA_POS_RegionCol.Count > 0 Then
        lRow = 2
        If bNewSht = False Then
            If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
            Application.ScreenUpdating = False
            Set OPwb = Application.Workbooks.Add
            Set OPSht = OPwb.Sheets(1)
            Application.Calculation = xlCalculationManual
            bNewSht = True
        End If
        OPSht.Cells(1, 1).Value = "Region"
        OPSht.Cells(1, 2).Value = "Sales Date"
        OPSht.Cells(1, 2).EntireColumn.NumberFormat = "DD/MM/YYYY"
        OPSht.Cells(1, 3).Value = "Store Name"
        OPSht.Cells(1, 4).Value = "Quantity"
        OPSht.Cells(1, 5).Value = "Retail"
        For Each vDiv In CBA_POS_RegionCol
            bOutput = CBA_SQL_Queries.CBA_GenPullSQL("MMS_StoreSalesDATA", CBA_POSDateFrom, CBA_POSDateTo, CBA_POSProductcode, vDiv, , , , , , , CBA_POSRunUseLinkedProducts)
            thisarr = CBA_MMSarr
            Erase CBA_MMSarr
            thisarr = CBA_BasicFunctions.CBA_TransposeArray(thisarr)
            Range(OPSht.Cells(lRow, 1), OPSht.Cells(lRow + UBound(thisarr, 1), UBound(thisarr, 2) + 1)).Value2 = thisarr
            lRow = lRow + UBound(thisarr, 1) + 1
        Next
        Set wks_Piv = Sheets.Add
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range(OPSht.Cells(1, 1), OPSht.Cells(lRow, 5)), Version:=xlPivotTableVersion14).CreatePivotTable _
            TableDestination:=wks_Piv.Cells(3, 1), TableName:="StorePOS" & Format(Date, "YYYY-MM-DD"), DefaultVersion:=xlPivotTableVersion14
        With wks_Piv.PivotTables(1).PivotFields("Sales Date")
            .Orientation = xlPageField
            .Position = 1
        End With
        'wks_Piv.PivotTables(1).PivotFields("Sales Date").NumberFormat = "YYYY-MM-DD"
        wks_Piv.PivotTables(1).AddDataField wks_Piv.PivotTables(1).PivotFields("Quantity"), "Sum of Quantity", xlSum
        wks_Piv.PivotTables(1).AddDataField wks_Piv.PivotTables(1).PivotFields("Retail"), "Sum of Retail", xlSum
        With wks_Piv.PivotTables(1).PivotFields("Region")
            .Orientation = xlRowField
            .Position = 1
        End With
        With wks_Piv.PivotTables(1).PivotFields("Store Name")
            .Orientation = xlRowField
            .Position = 2
        End With

        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running

        OPSht.Cells.EntireColumn.AutoFit
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_StoreSalesQry", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Sub CBA_POS_addtoRegionCol(ByVal vDiv As Long)
    CBA_POS_RegionCol.Add vDiv
End Sub
'Callback for CBA_POSPCodeEnter onChange
Sub CBA_POSPCodeEnter_onChange(Control As IRibbonControl, ByRef text)
    If IsNumeric(text) = True And Len(text) > 3 And Len(text) < 11 Then
        CBA_POSProductcode = text
        Call CBA_isPOSRunButtonEnabled
        'CBA_Rib.InvalidateControl "CBA_BT_UnrealisedRevenue"
    Else
        CBA_POSRunEnable = False
        CBA_POSProductcode = 0
        text = ""
        'CBA_Rib.InvalidateControl "CBA_BT_UnrealisedRevenue"
        'MsgBox "Invalid productcode entered" & Chr(10) & Chr(10) & "Please try again"
        'CBA_ribbon.CBA_RefreshRibbon
        'CBA_Rib.Invalidate
    End If
    Call CBA_isPOSRunButtonEnabled
End Sub
'Callback for CBA_POSDatefrom onChange
Sub CBA_POSPDatefrom_onChange(Control As IRibbonControl, ByRef text)
    If isDate(text) = True And text <> "12:00:00 AM" Then
        CBA_POSDateFrom = text
        'Call CBA_isPOSRunButtonEnabled
    Else
        CBA_POSRunEnable = False
        'MsgBox "Invalid productcode entered" & Chr(10) & Chr(10) & "Please try again"
    'CBA_ribbon.CBA_RefreshRibbon
    'CBA_Rib.Invalidate
    End If
    Call CBA_isPOSRunButtonEnabled
End Sub
'Callback for CBA_POSDateto onChange
Sub CBA_POSDateto_onChange(Control As IRibbonControl, ByRef text)
    If isDate(text) = True And text <> "12:00:00 AM" Then
        CBA_POSDateTo = text
        'Call CBA_isPOSRunButtonEnabled
    Else
        CBA_POSRunEnable = False
        'MsgBox "Invalid productcode entered" & Chr(10) & Chr(10) & "Please try again"
        'CBA_ribbon.CBA_RefreshRibbon
        'CBA_Rib.Invalidate
    End If
    Call CBA_isPOSRunButtonEnabled
End Sub
Sub CBA_isPOSRunButtonEnabled()
    If CBA_POSProductcode = 0 Or CBA_POSDateFrom = "12:00:00 AM" Or CBA_POSDateTo = "12:00:00 AM" Then
        CBA_POSRunEnable = False
    Else
        CBA_POSRunEnable = True
    End If
        On Error GoTo GTRefresh
        CBA_Rib.InvalidateControl "CBA_BT_UnrealisedRevenue"
        CBA_Rib.InvalidateControl "CBA_RunPOS"
        CBA_Rib.InvalidateControl "CBA_RunStoreSales"
        CBA_Rib.InvalidateControl "CBA_RunInventory"
        On Error GoTo 0


GTRefresh:
        CBA_Ribbon.CBA_RefreshRibbon
        CBA_Rib.InvalidateControl "CBA_BT_UnrealisedRevenue"
        CBA_Rib.InvalidateControl "CBA_RunPOS"
        CBA_Rib.InvalidateControl "CBA_RunStoreSales"
        CBA_Rib.InvalidateControl "CBA_RunInventory"
        On Error GoTo 0
        'CBA_Rib.Invalidate
End Sub
'Callback for CBA_BT_UnrealisedRevenue getEnabled
Sub CBA_BT_OSDandProductSelected(Control As IRibbonControl, ByRef returnedVal)
    If CBA_OSDSelected <> "12:00:00 AM" And CBA_POSProductcode <> 0 Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub
Sub Get_CBA_RunPOS_Enable(Control As IRibbonControl, ByRef CBA_enabled)
    CBA_enabled = CBA_POSRunEnable
End Sub
'Callback for CBA_RunPOSConsolidated onAction
Sub CBA_FlagConsolidated(Control As IRibbonControl, pressed As Boolean)
    CBA_POSRunUseLinkedProducts = pressed
End Sub
'Callback for CBA_RunPOSConsolidated getPressed
Sub CBA_RunPOSConsolidated_DefaultSetting(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = CBA_POSRunUseLinkedProducts
End Sub


