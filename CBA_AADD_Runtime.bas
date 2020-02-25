Attribute VB_Name = "CBA_AADD_Runtime"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Private CBA_AADD_Enabled As Boolean
Private CBA_AADD_CGSCGList() As String
Private activePromotions() As CBA_AADD_frm_Promotion
Private activeCampaigns() As CBA_AADD_frm_Campaign

'Callback for id_CBA_AADD_ImportData onAction
Sub CBA_AADD_ImportData(Control As IRibbonControl)
    Dim wbk As Workbook, importWbk As Workbook
    Dim sht As Worksheet, importSht As Worksheet
    Dim lRet As Byte, colnum As Long, a As Long, firstRow As Long, lRowNo As Long
    Dim RCell As Range, cel As Range, thisRng As Range, mediaHead, endofit, strData As String, medcol
    Dim IPWB, wks_Prep, prow As Long, Lastrow As Long, campcol, thisPromoType
    Dim ImportForm As CBA_AADD_frm_Import
    On Error GoTo Err_Routine

    For Each wbk In Application.Workbooks
        lRet = MsgBox("Import File:" & Chr(10) & Chr(10) & wbk.Name, vbYesNo)
        If lRet = 6 Then
            Set importWbk = wbk
            Exit For
        End If
    Next
    If importWbk Is Nothing Then Exit Sub
    
    For Each sht In importWbk.Worksheets
        lRet = MsgBox("Import from tab named: " & Chr(10) & Chr(10) & sht.Name, vbYesNo)
        If lRet = 6 Then
            Set importSht = sht
            Exit For
        End If
    Next
    If importSht Is Nothing Then Exit Sub

    With importSht
        
        'MsgBox "Please enter the column containing the first month of campaign data", vbOKOnly, vbModeless
        Set thisRng = Selection
        Set ImportForm = New CBA_AADD_frm_Import
        ImportForm.Show vbModeless
        Do While thisRng.Address = Selection.Address
            DoEvents
        Loop
        Unload ImportForm
        Set ImportForm = Nothing
        If thisRng.Address = Selection.Address Then Exit Sub
        colnum = Selection.Column
        If colnum = 0 Then Exit Sub
        
        If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Importing Data to AADD"
        Application.ScreenUpdating = False
        Set IPWB = Workbooks.Add
        Set wks_Prep = ActiveSheet
        '.Activate
        '.Cells.EntireColumn.Hidden = False
        prow = 1
        For Each RCell In .Columns(colnum).Cells
            'RCell.MergeArea(1, 1).Value
            If RCell.MergeArea(1, 1).Value <> "" Then
                If lRowNo = 0 Then
                    firstRow = RCell.Row
                    lRowNo = RCell.Row
                    Exit For
                End If
            End If
        Next
        Lastrow = .Cells(999999, colnum).End(xlUp).Row
        For Each RCell In .Rows(firstRow + 1).Cells
            'RCell.Select
            If RCell.Value = "" And RCell.Offset(0, 3).Value = "" And RCell.Offset(0, 7).Value = "" And RCell.Offset(0, 11).Value = "" _
                And RCell.Offset(0, 15).Value = "" And RCell.Offset(0, 20).Value = "" And RCell.Offset(0, 25).Value = "" Then Exit For
            If InStr(1, LCase(RCell.Value), "commenc") > 0 Then campcol = RCell.Column
            If RCell.Column >= colnum Then
                If prow = 1 Then
                    wks_Prep.Cells(1, 1).Value = "Week Commencing"
                    wks_Prep.Cells(1, 2).Value = "Column Reference"
                    wks_Prep.Cells(1, 3).Value = "Week in Year"
                    wks_Prep.Cells(1, 4).Value = "Price"
                    wks_Prep.Cells(1, 5).Value = "PriceCI"
                    wks_Prep.Cells(1, 6).Value = "Masterbrand"
                    wks_Prep.Cells(1, 7).Value = "MasterbrandCI"
                    wks_Prep.Cells(1, 8).Value = "Special Buys"
                    wks_Prep.Cells(1, 9).Value = "Special BuysCI"
                    wks_Prep.Cells(1, 10).Value = "Campaigns"
                    wks_Prep.Cells(1, 11).Value = "CampaignsCI"
                    wks_Prep.Cells(1, 12).Value = "Mobile"
                    wks_Prep.Cells(1, 13).Value = "MobileCI"
                    wks_Prep.Cells(1, 14).Value = "Always on Search"
                    wks_Prep.Cells(1, 15).Value = "Always on SearchCI"
                    wks_Prep.Cells(1, 16).Value = "Always on Social"
                    wks_Prep.Cells(1, 17).Value = "Always on SocialCI"
                    wks_Prep.Cells(1, 18).Value = "Holidays"
                    wks_Prep.Cells(1, 19).Value = "HolidaysCI"
                    wks_Prep.Cells(1, 20).Value = "Comment"
                    wks_Prep.Cells(1, 21).Value = "CommentCI"
                End If
                prow = prow + 1
                wks_Prep.Cells(prow, 1).Value = RCell.Value
                wks_Prep.Cells(prow, 2).Value = RCell.Column
                
            End If
        Next
        For Each cel In wks_Prep.Columns(2).Cells
            mediaHead = 0
            If cel.Row > 1 Then
                If cel.Value = "" Then
                    endofit = cel.Row - 1
                    Exit For
                End If
                'If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 7, 3, "Pulling data for Week Commencing " & wks_Prep.Cells(cel.Row, 1).Value
                Application.ScreenUpdating = True
                For Each RCell In Range(.Cells(firstRow, cel.Value), .Cells(Lastrow, cel.Value))
                    If mediaHead = 0 Then
                        If InStr(1, LCase(.Cells(RCell.Row, campcol).Value), "week in year") > 0 Then wks_Prep.Cells(cel.Row, 3).Value = RCell.MergeArea(1, 1).Value
                        If InStr(1, LCase(.Cells(RCell.Row, campcol).Value), "price") > 0 Then
                            wks_Prep.Cells(cel.Row, 4).Value = RCell.MergeArea(1, 1).Value
                            wks_Prep.Cells(cel.Row, 5).Value = RCell.Interior.ColorIndex
                        End If
                        If InStr(1, LCase(.Cells(RCell.Row, campcol).Value), "masterbrand") > 0 Then wks_Prep.Cells(cel.Row, 6).Value = RCell.MergeArea(1, 1).Value: wks_Prep.Cells(cel.Row, 7).Value = RCell.Interior.ColorIndex
                        If InStr(1, LCase(.Cells(RCell.Row, campcol).Value), "special") > 0 Then wks_Prep.Cells(cel.Row, 8).Value = RCell.MergeArea(1, 1).Value: wks_Prep.Cells(cel.Row, 9).Value = RCell.Interior.ColorIndex
                        If InStr(1, LCase(.Cells(RCell.Row, campcol).Value), "campaign") > 0 Then wks_Prep.Cells(cel.Row, 10).Value = RCell.MergeArea(1, 1).Value: wks_Prep.Cells(cel.Row, 11).Value = RCell.Interior.ColorIndex
                        If InStr(1, LCase(.Cells(RCell.Row, campcol).Value), "mobile") > 0 Then wks_Prep.Cells(cel.Row, 12).Value = RCell.MergeArea(1, 1).Value: wks_Prep.Cells(cel.Row, 13).Value = RCell.Interior.ColorIndex
                        If InStr(1, LCase(.Cells(RCell.Row, campcol).Value), "search") > 0 Then wks_Prep.Cells(cel.Row, 14).Value = RCell.MergeArea(1, 1).Value: wks_Prep.Cells(cel.Row, 15).Value = RCell.Interior.ColorIndex
                        If InStr(1, LCase(.Cells(RCell.Row, campcol).Value), "social") > 0 Then wks_Prep.Cells(cel.Row, 16).Value = RCell.MergeArea(1, 1).Value: wks_Prep.Cells(cel.Row, 17).Value = RCell.Interior.ColorIndex
                        If InStr(1, LCase(.Cells(RCell.Row, campcol).Value), "holiday") > 0 Then wks_Prep.Cells(cel.Row, 18).Value = RCell.MergeArea(1, 1).Value: wks_Prep.Cells(cel.Row, 19).Value = RCell.Interior.ColorIndex
                        If InStr(1, LCase(.Cells(RCell.Row, campcol).Value), "comment") > 0 Then wks_Prep.Cells(cel.Row, 20).Value = RCell.MergeArea(1, 1).Value: wks_Prep.Cells(cel.Row, 21).Value = RCell.Interior.ColorIndex
                        If mediaHead = 0 And InStr(1, LCase(.Cells(RCell.Row, 1).Value), "media") > 0 Then
                                mediaHead = RCell.Row
                                strData = ""
                                For a = 1 To campcol
                                    If strData = "" Then strData = .Cells(RCell.Row, a).Value Else strData = strData & "|" & .Cells(RCell.Row, a).Value
                                Next
                                wks_Prep.Cells(1, 22).Value = strData
                        End If
                    End If
                    If mediaHead > 0 Then
                        If IsNumeric(RCell.Value) And RCell.Value <> "" And RCell.Row > mediaHead Then
                            If wks_Prep.Cells(cel.Row, 22).Value = "" Then medcol = 21
                            medcol = medcol + 1
                            strData = ""
                            For a = 1 To campcol
                                If strData = "" Then strData = .Cells(RCell.Row, a).Value Else strData = strData & "|" & .Cells(RCell.Row, a).Value
                            Next
                            For a = 5 To 21
                                thisPromoType = ""
                                If RCell.Interior.ColorIndex = wks_Prep.Cells(cel.Row, a).Value Then
                                    If a = 5 Then thisPromoType = "price"
                                    If a = 7 Then thisPromoType = "masterbrand"
                                    If a = 9 Then thisPromoType = "specialbuy"
                                    If a = 11 Then thisPromoType = "campaign"
                                    If a = 13 Then thisPromoType = "mobile"
                                    If a = 15 Then thisPromoType = "alwaysonsearch"
                                    If a = 17 Then thisPromoType = "alwaysonsocial"
                                    If a = 19 Then thisPromoType = "holidays"
                                    If a = 21 Then thisPromoType = "Comments"
                                    Exit For
                                End If
                            Next
                            wks_Prep.Cells(cel.Row, medcol).Value = thisPromoType & ":TARPP:" & RCell.Value & ";" & strData
                        End If
                    End If
                    
                Next
            
            End If
        Next
        For a = endofit To 1 Step -1
            If wks_Prep.Cells(a, 1).Value = "" Then
                wks_Prep.Cells(a, 1).EntireRow.Delete
            End If
        Next
        
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
        .Activate
        Application.ScreenUpdating = True
        MsgBox "Data Imported", vbOKOnly
   End With
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("CBA_AADD_ImportData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub
'Callback for CBA_AADD_Activation onAction
Sub CBA_AADD_ActivateData(Control As IRibbonControl, pressed As Boolean)
    Dim a As Long, b As Long
    On Error GoTo Err_Routine
    
    If CBA_AADD_Enabled = True Then
        CBA_AADD_Enabled = False
        dropPrivateVariables
    Else
        CBA_AADD_Enabled = True
        CBA_SQL_Queries.CBA_GenPullSQL "CBIS_CGSCGList"
        ReDim activePromotions(1 To 1)
        ReDim activeCampaigns(1 To 1)
        ReDim CBA_AADD_CGSCGList(LBound(CBA_CBISarr, 1) To UBound(CBA_CBISarr, 1), LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2))
        For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
            If a > UBound(CBA_CBISarr, 2) Then Exit For
            For b = LBound(CBA_CBISarr, 1) To UBound(CBA_CBISarr, 1)
                CBA_AADD_CGSCGList(b, a) = CStr(CBA_CBISarr(b, a))
            Next
        Next
        Set CBA_AADD_CBISCN = New ADODB.Connection
        With CBA_AADD_CBISCN
            .ConnectionTimeout = 50
            .CommandTimeout = 50
            .Open "Provider= SQLNCLI10; DATA SOURCE= 599DBL01; ;INTEGRATED SECURITY=sspi;"
        End With
        For a = 501 To 509
            If a <> 508 Then
                Set CBA_AADD_MMSCN(a) = New ADODB.Connection
                With CBA_AADD_MMSCN(a)
                    .ConnectionTimeout = 50
                    .CommandTimeout = 50
                    .Open "Provider= SQLNCLI10; DATA SOURCE= " & CBA_BasicFunctions.TranslateServerName(a, Date) & "; ;INTEGRATED SECURITY=sspi;"
                End With
            End If
        Next
        Erase CBA_CBISarr
    End If
    CBA_Ribbon.CBA_RefreshRibbon
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("CBA_AADD_ActivateData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
'Callback for CBA_AADD_Activation getPressed
Sub CBA_AADD_GetTogglePressed(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = CBA_AADD_Enabled
End Sub
'Callback for id_CBA_AADD_ManageActivityPlanner onAction
Sub CBA_AADD_ManageActivityPlanner(Control As IRibbonControl)
End Sub
'Callback for id_CBA_AADD_ManageActivityPlanner getEnabled
Sub Get_AADD_RibbonEnabled(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = CBA_AADD_Enabled
End Sub
'Callback for id_CBA_AADD_ManageCampaign onAction
Sub CBA_AADD_ManageCampaign(Control As IRibbonControl)
    Dim upd As Long, a As Long
    On Error GoTo Err_Routine
    If activeCampaigns(1) Is Nothing Then
        upd = 1
    Else
        For a = LBound(activeCampaigns) To UBound(activeCampaigns)
            If activeCampaigns(a) Is Nothing Then
                upd = a
                Exit For
            End If
        Next
    End If
    If upd = 0 Then
        ReDim Preserve activeCampaigns(1 To UBound(activeCampaigns) + 1)
        upd = UBound(activeCampaigns)
    End If
    Set activeCampaigns(upd) = New CBA_AADD_frm_Campaign
    activeCampaigns(upd).Show vbModeless
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_AADD_ManageCampaign", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
'Callback for id_CBA_AADD_ManagePromotion onAction
Sub CBA_AADD_ManagePromotion(Control As IRibbonControl)
    Dim upd As Long, a As Long
    If activePromotions(1) Is Nothing Then
        upd = 1
    Else
        For a = LBound(activePromotions) To UBound(activePromotions)
            If activePromotions(a) Is Nothing Then
                upd = a
                Exit For
            End If
        Next
    End If
    If upd = 0 Then
        ReDim Preserve activePromotions(1 To UBound(activePromotions) + 1)
        upd = UBound(activePromotions)
    End If
    Set activePromotions(upd) = New CBA_AADD_frm_Promotion
    activePromotions(upd).Show vbModeless
End Sub
Private Sub dropPrivateVariables()
Dim a As Integer
    Erase CBA_AADD_CGSCGList
    Erase activePromotions
    Erase activeCampaigns
    If CBA_AADD_CBISCN.State = 1 Then CBA_AADD_CBISCN.Close
    Set CBA_AADD_CBISCN = Nothing
    For a = 501 To 509
        If a <> 508 Then
            If CBA_AADD_MMSCN(a).State = 1 Then CBA_AADD_MMSCN(a).Close
            Set CBA_AADD_MMSCN(a) = Nothing
        End If
    Next
End Sub
Function CBA_AADD_getCGSCGList() As String()
    CBA_AADD_getCGSCGList = CBA_AADD_CGSCGList
End Function
