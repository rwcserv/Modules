Attribute VB_Name = "CBAR_REPORTStateVar"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Type stateprice
    NSW As Single
    QLD As Single
    vic As Single
    sa As Single
    wa As Single
End Type

Function StateVariationReport()
    Dim outputarr()
    Dim stp As stateprice, bfound As Boolean
    Dim cntstp As stateprice
    Dim PLmatches As Collection
    Dim tR As CBAR_Report, ouputarrcreated, usematch, check
    Dim bOutput As Boolean, prod, wbk, wks_SVR, regval
    Dim lNum As Long, strProds, a As Long, b As Long, DFrom, Dto, Reg
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    tR = CBAR_Runtime.getActiveReport
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Preparing to run 'State Variation Report'"
    
    strProds = ""
    bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("COM_2ScrapeDates")
    If bOutput = False Then
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
        MsgBox "There has been an error in querying COMRADE" & Chr(10) & Chr(10) & "Please try again later or contact " & g_Get_Dev_Sts("DevUsers")
        Exit Function
    End If
    For a = LBound(CBA_COMarr, 2) To UBound(CBA_COMarr, 2)
        If a = LBound(CBA_COMarr, 2) Then
            DFrom = CBA_COMarr(1, a)
            Dto = CBA_COMarr(1, a)
        Else
            If CBA_COMarr(1, a) < DFrom Then DFrom = CBA_COMarr(1, a)
            If CBA_COMarr(1, a) > Dto Then Dto = CBA_COMarr(1, a)
        End If
    Next
    
    If WeekDay(Dto, vbWednesday) <> 0 Then
        Dto = DateAdd("d", -WeekDay(Dto, vbWednesday) + 1, Dto)
    End If
    If WeekDay(DFrom, vbWednesday) <> 0 Then
        DFrom = DateAdd("d", -WeekDay(DFrom, vbWednesday) + 1, DFrom)
    End If
    
    If tR.BD <> "" Then
        bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("CBIS_ProdbyEmp")
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
        bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("CBIS_ProdbyEmp")
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
    DFrom = Dto
    
    If CBA_COM_SetupMatchArray.CBA_SetupMatchArray(False, DFrom, Dto, tR.CG, tR.scg, strProds) = True Then
        CBAR_SVR.Copy
        Set wbk = ActiveWorkbook
        wbk.VBProject.VBComponents.Import CBA_BSA & "VBA Development Tools\FORMS\COMRADE USERFORMS\CBA_COM_frm_Chart.frm"
        wbk.VBProject.VBComponents.Import CBA_BSA & "VBA Development Tools\FORMS\COMRADE USERFORMS\CBAR_PPHChartCreate.bas"
        wbk.VBProject.References.AddFromFile "C:\Program Files (x86)\Common Files\System\ado\msado15.dll"
        Set wks_SVR = ActiveSheet
        
        lNum = 1
        With wks_SVR
            Range(.Cells(5, 1), .Cells(99999, 29)).ClearContents
            Range(.Cells(5, 1), .Cells(99999, 29)).Interior.ColorIndex = 0
            Set PLmatches = New Collection
            PLmatches.Add "colesval"
            PLmatches.Add "colescoles"
            PLmatches.Add "colessb"
            PLmatches.Add "wwselect"
            PLmatches.Add "wwhb"
            PLmatches.Add "dm1"
            PLmatches.Add "dm2"
            ouputarrcreated = False
            For a = LBound(CBA_COM_Match) To UBound(CBA_COM_Match)
                usematch = True
                If usematch = True Then
                    cntstp.NSW = 0: cntstp.QLD = 0: cntstp.vic = 0: cntstp.sa = 0: cntstp.wa = 0: stp.NSW = 0: stp.QLD = 0: stp.vic = 0: stp.sa = 0: stp.wa = 0: regval = 0
                    
                    stp.NSW = CBA_COM_Match(a).Pricedata(Dto, "shelf", "nsw")
                    stp.QLD = CBA_COM_Match(a).Pricedata(Dto, "shelf", "qld")
                    stp.vic = CBA_COM_Match(a).Pricedata(Dto, "shelf", "vic")
                    stp.sa = CBA_COM_Match(a).Pricedata(Dto, "shelf", "sa")
                    stp.wa = CBA_COM_Match(a).Pricedata(Dto, "shelf", "wa")
                    If stp.NSW > 0 Then
                        check = stp.NSW
                    ElseIf stp.QLD > 0 Then check = stp.QLD
                    ElseIf stp.vic > 0 Then check = stp.vic
                    ElseIf stp.sa > 0 Then check = stp.sa
                    ElseIf stp.wa > 0 Then check = stp.wa
                    End If
                    
                    If stp.NSW = stp.QLD Then cntstp.NSW = cntstp.NSW + 1
                    If stp.NSW = stp.vic Then cntstp.NSW = cntstp.NSW + 1
                    If stp.NSW = stp.sa Then cntstp.NSW = cntstp.NSW + 1
                    If stp.NSW = stp.wa Then cntstp.NSW = cntstp.NSW + 1
                    
                    If stp.QLD = stp.NSW Then cntstp.QLD = cntstp.QLD + 1
                    If stp.QLD = stp.sa Then cntstp.QLD = cntstp.QLD + 1
                    If stp.QLD = stp.wa Then cntstp.QLD = cntstp.QLD + 1
                    If stp.QLD = stp.vic Then cntstp.QLD = cntstp.QLD + 1
            
                    If stp.vic = stp.NSW Then cntstp.vic = cntstp.vic + 1
                    If stp.vic = stp.sa Then cntstp.vic = cntstp.vic + 1
                    If stp.vic = stp.wa Then cntstp.vic = cntstp.vic + 1
                    If stp.vic = stp.QLD Then cntstp.vic = cntstp.vic + 1
            
                    If stp.sa = stp.NSW Then cntstp.sa = cntstp.sa + 1
                    If stp.sa = stp.QLD Then cntstp.sa = cntstp.sa + 1
                    If stp.sa = stp.wa Then cntstp.sa = cntstp.sa + 1
                    If stp.sa = stp.vic Then cntstp.sa = cntstp.sa + 1
            
                    If stp.wa = stp.NSW Then cntstp.wa = cntstp.wa + 1
                    If stp.wa = stp.sa Then cntstp.wa = cntstp.wa + 1
                    If stp.wa = stp.QLD Then cntstp.wa = cntstp.wa + 1
                    If stp.wa = stp.vic Then cntstp.wa = cntstp.wa + 1
            
                    Reg = cntstp.NSW
                    If cntstp.QLD > Reg Then Reg = cntstp.QLD
                    If cntstp.vic > Reg Then Reg = cntstp.vic
                    If cntstp.sa > Reg Then Reg = cntstp.sa
                    If cntstp.wa > Reg Then Reg = cntstp.wa
            
                    If Reg = cntstp.NSW Then
                        regval = stp.NSW
                    ElseIf Reg = cntstp.QLD Then regval = stp.QLD
                    ElseIf Reg = cntstp.vic Then regval = stp.vic
                    ElseIf Reg = cntstp.sa Then regval = stp.sa
                    ElseIf Reg = cntstp.wa Then regval = stp.wa
                    End If
                    
                    If (check = stp.QLD Or stp.QLD = 0) And (check = stp.vic Or stp.vic = 0) And (check = stp.sa Or stp.sa = 0) And (check = stp.wa Or stp.wa = 0) Then
                    Else
                        lNum = lNum + 1
                        ReDim Preserve outputarr(1 To 12, 2 To lNum)
                        ouputarrcreated = True
                        outputarr(1, lNum) = CBA_COM_Match(a).AldiPCode
                        outputarr(2, lNum) = CBA_COM_Match(a).AldiPName
                        outputarr(3, lNum) = CBA_COM_Match(a).Pricedata(Dto, "AldiRetail", "national")
                        outputarr(4, lNum) = CBA_COM_Match(a).Competitor
                        outputarr(5, lNum) = CBA_COM_Match(a).CompCode
                        outputarr(6, lNum) = CBA_COM_Match(a).CompProdName
                        outputarr(7, lNum) = stp.NSW
                        outputarr(8, lNum) = stp.QLD
                        outputarr(9, lNum) = stp.vic
                        outputarr(10, lNum) = stp.sa
                        outputarr(11, lNum) = stp.wa
                        outputarr(12, lNum) = regval
                    End If
                End If
            Next
            If ouputarrcreated = True Then
                lNum = 4
                For a = LBound(outputarr, 2) To UBound(outputarr, 2)
                    If a = LBound(outputarr, 2) Then
                        lNum = lNum + 1
                        For b = 1 To 11
                            .Cells(lNum, b).Value = outputarr(b, a)
                            If outputarr(b, a) > 0 And outputarr(b, a) <> outputarr(9, a) And b > 6 Then .Cells(lNum, b).Interior.ColorIndex = 22
                        Next
                    Else
                        bfound = False
                        For b = LBound(outputarr, 2) To a - 1
                            If outputarr(5, b) = outputarr(5, a) Then
                                bfound = True
                                Exit For
                            End If
                        Next
                        If bfound = False Then
                            lNum = lNum + 1
                            For b = 1 To 11
                                .Cells(lNum, b).Value = outputarr(b, a)
                                If outputarr(b, a) > 0 And outputarr(b, a) <> outputarr(9, a) And b > 6 Then .Cells(lNum, b).Interior.ColorIndex = 22
                            Next
                        End If
                    End If
                Next
            Else
                .Cells(2, 1).Value = "None of the match types applicable to this report returned a State Variation"
            End If
            
            
            .PageSetup.PrintArea = Range(.Cells(1, 1), .Cells(lNum, 11)).Address
            .PageSetup.Zoom = False
            .PageSetup.FitToPagesWide = 1
            .PageSetup.FitToPagesTall = False
            .PageSetup.LeftFooter = "&9CORP BUYING, Admin, per: " & Format(Date, "DD/MM/YYYY") & Chr(10) & Application.ActiveWorkbook.FullName
            .PageSetup.Orientation = xlLandscape
            .PageSetup.PrintGridlines = True
            .PageSetup.PrintTitleRows = Range(.Cells(1, 1), .Cells(5, 15)).Address
            .PageSetup.RightFooter = "&P of &N"
            .Columns(5).NumberFormat = "General"
            .Visible = xlSheetVisible
            .Activate
            .Cells(1, 1).Select
        End With
    
    End If
    
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-StateVariationReport", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

