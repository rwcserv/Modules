Attribute VB_Name = "CBA_BT_ALDT"
Option Explicit
Option Private Module          ' Excel users cannot access procedures

Sub pullData()
    Dim a, b, c, d As Long
    Dim USW() As Single
    Dim RCell As Range
    Dim output As Boolean, TaxID As Boolean
    Dim Data As Variant
    Dim vDiv As Long, DivNo As Long
    Dim TotRet As Single, allretail As Single, precont As Single, allretailwoLoss As Single, allcost As Single, finalusw As Single, finalcsw As Single
    Dim finalcswloss As Single, activediv As Single, lossu As Single, losses As Single
    On Error GoTo Err_Routine
        
    CBA_DBtoQuery = 599
    CBA_BasicFunctions.CBA_Running "Preparing LDT Data..."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    CBA_T_ALDT.Copy
    
    With ActiveSheet
    
    If CBA_BT_frm_ALDT.lbx_ProdsToRun.ListCount > 1 Then
        Range(.Cells(10, 2), .Cells(10, 15)).Copy
        Range(.Cells(10, 2), .Cells(10 + CBA_BT_frm_ALDT.lbx_ProdsToRun.ListCount - 1, 15)).PasteSpecial xlPasteAll
        
    End If
    
    For a = 0 To CBA_BT_frm_ALDT.lbx_ProdsToRun.ListCount - 1
        .Cells(a + 10, 2).Value = CBA_BT_frm_ALDT.lbx_ProdsToRun.List(a, 0)
        .Cells(a + 10, 13).Value = DateSerial(Year(CBA_BT_frm_ALDT.lbx_ProdsToRun.List(a, 2)), Month(CBA_BT_frm_ALDT.lbx_ProdsToRun.List(a, 2)), Day(CBA_BT_frm_ALDT.lbx_ProdsToRun.List(a, 2)))
        .Cells(a + 10, 14).Value = DateSerial(Year(CBA_BT_frm_ALDT.lbx_ProdsToRun.List(a, 3)), Month(CBA_BT_frm_ALDT.lbx_ProdsToRun.List(a, 3)), Day(CBA_BT_frm_ALDT.lbx_ProdsToRun.List(a, 3)))
    Next
    
    Unload CBA_BT_frm_ALDT
    
    
    For Each RCell In .Columns(2).Cells
        If RCell.Row > 9 Then
            If RCell.Value = "" And RCell.Offset(1, 0).Value = "" And RCell.Offset(1, 0).Value = "" And RCell.Offset(2, 0).Value = "" And RCell.Offset(3, 0).Value = "" And RCell.Offset(4, 0).Value = "" Then Exit For
            'If RCell.Offset(0, -1).Value = "Yes" And RCell.Value <> "" And .Cells(RCell.Row, 13).Value <> "" And .Cells(RCell.Row, 14).Value <> "" Then
                
                If IsNumeric(RCell.Value) = True And isDate(.Cells(RCell.Row, 13).Value) And isDate(.Cells(RCell.Row, 14).Value) Then
                    'If CBA_CBISarr = Empty Then Else Erase CBA_CBISarr
                    CBA_DBtoQuery = 599
                    CBA_SQL_Queries.CBA_GenPullSQL "CBIS_Prodinfo", , , RCell.Value
                    .Cells(RCell.Row, 3).Value = CBA_CBISarr(0, 0)
                    Application.ScreenUpdating = True
                    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Preparing LDT Data..."
                    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 11, 4, "Processing data for product: " & RCell.Value & " - " & CBA_CBISarr(0, 0)
                    DoEvents
                    Application.ScreenUpdating = False
                    'If CBA_CBISarr = Empty Then Else Erase CBA_CBISarr
                    CBA_DBtoQuery = 599
                    output = CBA_SQL_Queries.CBA_GenPullSQL("CBIS_LDT", .Cells(RCell.Row, 13).Value, .Cells(RCell.Row, 14).Value, RCell.Value)
                    If CBA_CBISarr(0, 0) <> 0 And output = True Then
                        
                        Data = DataBuilder(.Cells(RCell.Row, 13).Value, .Cells(RCell.Row, 14).Value, RCell.Value)
                        
                        CBA_DBtoQuery = 599
                        CBA_SQL_Queries.CBA_GenPullSQL "CBIS_IsAlcohol", , , , , , RCell.Value
                        If CBA_CBISarr(0, 0) = "Yes" Then
                            CBA_SQL_Queries.CBA_GenPullSQL "CBIS_WeekYear", .Cells(RCell.Row, 13).Value, .Cells(RCell.Row, 14).Value
                            CBA_DBtoQuery = 1
                            output = CBA_SQL_Queries.CBA_GenPullSQL("ABI_AlcoholStores", .Cells(RCell.Row, 13).Value, .Cells(RCell.Row, 14).Value)
                            DivNo = 0
                            For vDiv = LBound(Data, 3) To UBound(Data, 3)
                                If vDiv <> 508 Then
                                    DivNo = DivNo + 1
                                    For a = LBound(Data, 2) To UBound(Data, 2)
                                        For c = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                                            If CBA_CBISarr(0, c) = Data(1, a, vDiv) Then
                                                For b = LBound(CBA_ABIarr, 2) To UBound(CBA_ABIarr, 2)
                                                    If Year(CBA_ABIarr(0, b)) >= CBA_CBISarr(1, 0) And Month(CBA_ABIarr(0, b)) >= Month(CBA_BasicFunctions.GetDayFromWeekNumber(CLng(CBA_CBISarr(1, 0)), CLng(CBA_CBISarr(0, 0)), 3)) Then
                                                        Data(5, a, vDiv) = CBA_ABIarr(DivNo, b)
                                                    End If
                                                Next
                                            End If
                                        Next
                                    Next
                                End If
                            Next
    
                        End If
                        
                        
    '''''''''*****************************************************************************''''''''''''''''''''''''''''''''''''''''''
    ''''''''''              THIS CODE IS A DEBUGGER THAT OUTPUTS THE SQL TO THE OUTPUT SHEET WHICH IS MANUALLY HIDDEN OR UNHIDDEN
    '''''''''*****************************************************************************''''''''''''''''''''''''''''''''''''''''''
    '                    Row = 2
    '                    Range(wks_Output.Cells(3, 1), wks_Output.Cells(999, 50)).ClearContents
    '                    For vDiv = LBound(data, 3) To UBound(data, 3)
    '                        If data(1, 1, vDiv) <> 0 And data(1, 2, vDiv) <> 0 Then
    '                        For a = LBound(data, 2) To UBound(data, 2)
    '                            Row = Row + 1
    '                            wks_Output.Cells(Row, 1).Value = vDiv
    '                            For b = LBound(data, 1) To UBound(data, 1)
    '                                wks_Output.Cells(Row, b + 1) = data(b, a, vDiv)
    '                            Next
    '                        Next
    '                        End If
    '                    Next
    '''''''''*****************************************************************************************************''''''''''''''''''
    '''''''''*****************************************************************************************************''''''''''''''''''
                        
    
                        CBA_DBtoQuery = 599
                        CBA_SQL_Queries.CBA_GenPullSQL "CBIS_IsTax", , , , , , RCell.Value
                        If CBA_CBISarr(0, 0) = 0 Then
                            TaxID = True
                        ElseIf CBA_CBISarr(0, 0) = 1 Then
                            TaxID = False
                        Else
                            TaxID = True
                        End If
                        
                        ReDim USW(1 To 53)
                        ReDim storeno(1 To 53)
                        ReDim cont(1 To 53)
                        ReDim contincLoss(1 To 53)
                        TotRet = 0: allretail = 0: precont = 0: allretailwoLoss = 0: allcost = 0: finalusw = 0: finalcsw = 0: finalcswloss = 0: activediv = 0
                        For vDiv = 501 To 509
                            If Data(1, UBound(Data, 2), vDiv) <> 0 Then
                                If activediv = 0 Then activediv = vDiv
                                
                                For a = LBound(Data, 2) To UBound(Data, 2)
                                
    '                                If data(18, a, vDiv) > 0 Or data(19, a, vDiv) > 0 Then
    '                                a = a
    '                                End If
                                
                                
                                    If vDiv = activediv Then TotRet = TotRet + Data(7, a, vDiv)
                                    storeno(Data(1, a, vDiv)) = storeno(Data(1, a, vDiv)) + Data(5, a, vDiv)
                                    USW(Data(1, a, vDiv)) = USW(Data(1, a, vDiv)) + Data(14, a, vDiv)
                                    cont(Data(1, a, vDiv)) = cont(Data(1, a, vDiv)) + Data(6, a, vDiv)
                                    If Data(2, a, vDiv) = 0 Then
                                        contincLoss(Data(1, a, vDiv)) = contincLoss(Data(1, a, vDiv)) + (Data(6, a, vDiv) + Data(18, a, vDiv))
                                    Else
                                        contincLoss(Data(1, a, vDiv)) = contincLoss(Data(1, a, vDiv)) + (Data(6, a, vDiv) + Data(18, a, vDiv) + (Data(19, a, vDiv) * (Data(3, a, vDiv) / Data(2, a, vDiv))))
                                    End If
    'exnc Inventory Diff
                                    'contincLoss(data(1, a, vDiv)) = contincLoss(data(1, a, vDiv)) + (data(6, a, vDiv) + data(18, a, vDiv))
    'not adjusting for Markdowns    allretail = allretail + data(3, a, vDiv)
                                    allretail = allretail + Data(3, a, vDiv) - Data(16, a, vDiv)
    
                                        
    
                                    allcost = allcost + Data(4, a, vDiv)
                                    
    'not adjusting for Markdowns    precont = precont + data(6, a, vDiv)
                                    If TaxID = True Then
                                        precont = precont + Data(6, a, vDiv) - (Data(16, a, vDiv) / 1.1)
                                    Else
                                        precont = precont + Data(6, a, vDiv) - Data(16, a, vDiv)
                                    End If
                                    If Data(2, a, vDiv) = 0 Then
                                        allretailwoLoss = allretailwoLoss + (Data(3, a, vDiv) + Data(18, a, vDiv))
                                    Else
                                        allretailwoLoss = allretailwoLoss + (Data(3, a, vDiv) + Data(18, a, vDiv) + (Data(19, a, vDiv) * (Data(3, a, vDiv) / Data(2, a, vDiv))))
                                    End If
    'exnc Inventory Diff            allretailwoLoss = allretailwoLoss + (data(3, a, vDiv) + data(18, a, vDiv))
                                    lossu = lossu + Data(19, a, vDiv)
                                Next
                            End If
                        Next
                        finalusw = 0
                        For a = LBound(Data, 2) To UBound(Data, 2)
                            finalusw = finalusw + (USW(Data(1, a, activediv)) / storeno(Data(1, a, activediv)))
                            finalcsw = finalcsw + (cont(Data(1, a, activediv)) / storeno(Data(1, a, activediv)))
                            If lossu < 0 Then
                                finalcswloss = finalcswloss + (contincLoss(Data(1, a, activediv)) / storeno(Data(1, a, activediv)))
                            Else
                                finalcswloss = finalcsw
                            End If
                        Next
                        If (allretailwoLoss - allretail) < 0 Then losses = (allretailwoLoss - allretail) Else losses = 0
                        
                        .Cells(RCell.Row, 5).Value = Format((allretail / TotRet) * 100, "#,#.00")
                        .Cells(RCell.Row, 6).Value = Round(finalusw / Data(12, 0, activediv), 0)
                        '.Cells(RCell.Row, 7).Value = Format(precont, "$#,#.00")
                        .Cells(RCell.Row, 7).Value = Format(finalcsw / Data(12, 0, activediv), "$#,#.00")
                        .Cells(RCell.Row, 8).Value = Format(precont / allretail, "#.00%")
                        .Cells(RCell.Row, 10).Value = Format((precont + losses) / allretail, "#.00%")
                        .Cells(RCell.Row, 9).Value = Format((losses / Data(12, 0, activediv)) / (allretail / Data(12, 0, activediv)), "#,#.00%")
                        .Cells(RCell.Row, 11).Value = Format(finalcswloss / Data(12, 0, activediv), "$#,#.00")
    
                    End If
                End If
            'End If
        End If
    Next
    
    
    
    
    
    
    Application.Calculation = xlCalculationAutomatic
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    End With
    Application.ScreenUpdating = True
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-PullData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub
Function DataBuilder(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal PCode As String) As Single()
    Dim aData() As Single
    Dim coldivs As Collection
    Dim a As Long, thisdiv As Long, curnum As Long, MaxNum As Long, d As Long, c As Long, b As Long
    Dim bfound As Boolean, output As Boolean
    Dim vDiv
    On Error GoTo Err_Routine
        
    Set coldivs = New Collection
    For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
        bfound = False
        If a = 0 Then
            coldivs.Add CBA_CBISarr(0, a)
        Else
            For Each vDiv In coldivs
                If vDiv = CBA_CBISarr(0, a) Then
                    bfound = True
                    Exit For
                End If
            Next
            If bfound = False Then
                coldivs.Add CBA_CBISarr(0, a)
            End If
        End If
    Next
    
    thisdiv = 0
    For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
        If thisdiv <> CBA_CBISarr(0, a) Then
            curnum = 0
            thisdiv = CBA_CBISarr(0, a)
        End If
        curnum = curnum + 1
        If MaxNum < curnum Then MaxNum = curnum
        
    Next
    
    
    
    
    'asds = coldivs.Count
    ReDim aData(LBound(CBA_CBISarr, 1) + 1 To UBound(CBA_CBISarr, 1) + 5, LBound(CBA_CBISarr, 2) To MaxNum - 1, 501 To 509)
    
    For Each vDiv In coldivs
        d = -1
        For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
            If CBA_CBISarr(0, a) = vDiv Then
                d = d + 1
                c = 0
                For b = LBound(CBA_CBISarr, 1) + 1 To UBound(CBA_CBISarr, 1)
                    c = c + 1
                    If IsNull(CBA_CBISarr(b, a)) = True Then aData(c, d, CLng(vDiv)) = 0 Else aData(c, d, CLng(vDiv)) = CBA_CBISarr(b, a)
                Next
            End If
        Next
    Next

    


    For Each vDiv In coldivs
        CBA_DBtoQuery = vDiv
        If CBA_DBtoQuery <> 508 Then
            output = CBA_SQL_Queries.CBA_GenPullSQL("MMS_MDLoss", DateFrom, DateTo, , , , PCode)
            For a = LBound(aData, 2) To UBound(aData, 2)
                For b = LBound(CBA_MMSarr, 2) To UBound(CBA_MMSarr, 2)
                    If CBA_MMSarr(0, b) = aData(1, a, vDiv) Then
                        For c = 1 To 5
                            aData(UBound(aData, 1) - (c - 1), a, vDiv) = CBA_MMSarr(6 - c, b)
                        Next
                    End If
                Next
            Next
        End If
    Next
    
    DataBuilder = aData
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("DataBuilder", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
