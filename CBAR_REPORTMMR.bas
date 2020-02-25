Attribute VB_Name = "CBAR_REPORTMMR"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Sub CBAR_MMRRuntime()
    Dim AProduceP() As Variant, AldiCoreProds() As Variant, APPOut As Variant, ACPOut As Variant, botrow As Long
    Dim strProds As String, colPProds, lowestmatch, thisaldiret, thisarr, MT, ricol, pwide
    Dim tR As CBAR_Report, stdate, endate, thiscompret, avgcompret, totcompwks, avgspeccompret, totspeccompwks, comparer
    Dim bOutput As Boolean, bfound As Boolean, wks_MMR, wbk, MHis, a As Long, b As Long, v As Long, DFrom, Dto, prod
    Dim this As MatchTypeData
    Dim RowNum As Long, m As Long, col As Long, c As Long
    Dim incProduce As Boolean, incCore As Boolean, incAlcohol As Boolean, del As Boolean
''    Dim bFound As Boolean
    Dim yn As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    yn = MsgBox("Would you like to delete products that have adequate private label matches?", vbYesNo)
    
    
    Application.ScreenUpdating = False
    CBAR_MMR.Copy
    Set wks_MMR = ActiveSheet
    Set wbk = ActiveWorkbook
    MHis = False
    tR = CBAR_Runtime.getActiveReport
    
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Preparing to run 'Missing Match Report'"
    
    
    APPOut = CBA_COM_Runtime.getCCMProds
    strProds = ""
    
    
    If tR.AldiProds Is Nothing Then
    Else
        For Each prod In tR.AldiProds
            If strProds = "" Then strProds = prod Else strProds = strProds & ", " & prod
        Next
    End If
    
    
    If tR.BD <> "" And strProds = "" Then
        If tR.AldiProds Is Nothing Then Set tR.AldiProds = New Collection
        For a = LBound(APPOut, 2) To UBound(APPOut, 2)
            If APPOut(6, a) = tR.BD Then
                tR.AldiProds.Add APPOut(0, a)
                If strProds = "" Then strProds = APPOut(0, a) Else strProds = strProds & ", " & APPOut(0, a)
            End If
        Next
    End If
    
    
    If tR.GBD <> "" And strProds = "" Then
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
    
    
    
    bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("COM_2ScrapeDates")
    If bOutput = False Then
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
        MsgBox "There has been an error in querying COMRADE" & Chr(10) & Chr(10) & "Please try again later or contact " & g_Get_Dev_Sts("DevUsers")
        Exit Sub
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
    If tR.BD = "Produce" Or tR.CG = 58 Then Else Dto = CBA_COM_Runtime.CBA_getWedDate: DFrom = DateAdd("D", -8, Dto)
    
    If tR.AldiProds Is Nothing Then
        Set tR.AldiProds = New Collection
        For a = LBound(APPOut, 2) To UBound(APPOut, 2)
            tR.AldiProds.Add APPOut(0, a)
        Next
    End If
    
    
    With wks_MMR
    
    
    
    If CBA_COM_SetupMatchArray.CBA_SetupMatchArray(tR.Matchhistory, DFrom, Dto, tR.CG, tR.scg, strProds) = True Then
    RowNum = 6
    For Each prod In tR.AldiProds
        For a = LBound(APPOut, 2) To UBound(APPOut, 2)
            If prod = APPOut(0, a) Then
                RowNum = RowNum + 1
                .Cells(RowNum, 1).Value = APPOut(0, a)
                .Cells(RowNum, 2).Value = APPOut(1, a)
                .Cells(RowNum, 3).Value = APPOut(6, a)
                For b = 4 To 27
                    .Cells(RowNum, b) = "No"
                Next
                bfound = False
                For m = LBound(CBA_COM_Match) To UBound(CBA_COM_Match)
                    If CBA_COM_Match(m).AldiPCode = APPOut(0, a) Then
                        bfound = True
                        this = CCM_Mapping.MatchType(CBA_COM_Match(m).MatchType)
                        If this.CoreAlcProd = "Core" And CBA_COM_Match(m).AldiPCG <> 58 Then
                            incCore = True
                            If this.Competitor = "C" Then
                                If InStr(1, LCase(this.Description), "leader") > 0 Then
                                    col = 6
                                ElseIf InStr(1, LCase(this.Description), "smartbuy") > 0 Or InStr(1, LCase(this.Description), "private") > 0 Or InStr(1, LCase(this.Description), "value") > 0 Then
                                    col = 4
                                ElseIf InStr(1, LCase(this.Description), "phantom") > 0 Then
                                    col = 10
                                ElseIf InStr(1, LCase(this.Description), "control") > 0 Then
                                    col = 8
                                End If
                            ElseIf this.Competitor = "WW" Then
                                If InStr(1, LCase(this.Description), "leader") > 0 Then
                                    col = 7
                                ElseIf InStr(1, LCase(this.Description), "homebrand") > 0 Or InStr(1, LCase(this.Description), "private") > 0 Then
                                    col = 5
                                ElseIf InStr(1, LCase(this.Description), "phantom") > 0 Then
                                    col = 11
                                ElseIf InStr(1, LCase(this.Description), "control") > 0 Then
                                    col = 9
                                End If
                            End If
                        ElseIf this.CoreAlcProd = "Alcohol" Then
                            incAlcohol = True
                            If this.Competitor = "DM" Then
                                If InStr(1, LCase(this.Description), "price") > 0 Then
                                    col = 12
                                ElseIf InStr(1, LCase(this.Description), "quality") > 0 Then
                                    col = 14
                                End If
                            ElseIf this.Competitor = "FC" Then
                                If InStr(1, LCase(this.Description), "price") > 0 Then
                                    col = 13
                                ElseIf InStr(1, LCase(this.Description), "quality") > 0 Then
                                    col = 15
                                End If
                            End If
                        ElseIf this.CoreAlcProd = "Produce" Or CBA_COM_Match(m).AldiPCG = 58 Then
                            incProduce = True
                            If this.Competitor = "C" Then
                                If InStr(1, LCase(this.Description), "national") > 0 Then
                                    col = 16
                                ElseIf InStr(1, LCase(this.Description), "nsw") > 0 Then
                                    col = 17
                                ElseIf InStr(1, LCase(this.Description), "qld") > 0 Then
                                    col = 18
                                ElseIf InStr(1, LCase(this.Description), "vic") > 0 Then
                                    col = 19
                                ElseIf InStr(1, LCase(this.Description), "wa") > 0 Then
                                    col = 20
                                ElseIf InStr(1, LCase(this.Description), "sa") > 0 Then
                                    col = 21
                                End If
                            ElseIf this.Competitor = "WW" Then
                                If InStr(1, LCase(this.Description), "national") > 0 Then
                                    col = 22
                                ElseIf InStr(1, LCase(this.Description), "nsw") > 0 Then
                                    col = 23
                                ElseIf InStr(1, LCase(this.Description), "qld") > 0 Then
                                    col = 24
                                ElseIf InStr(1, LCase(this.Description), "vic") > 0 Then
                                    col = 25
                                ElseIf InStr(1, LCase(this.Description), "wa") > 0 Then
                                    col = 26
                                ElseIf InStr(1, LCase(this.Description), "sa") > 0 Then
                                    col = 27
                                End If
                            End If
                        End If
                        If col > 0 Then .Cells(RowNum, col).Value = "Yes"
                    End If
                Next
                If bfound = False Then
                    .Rows(RowNum).ClearContents
                    RowNum = RowNum - 1
                End If
                Exit For
            End If
        Next
    Next

    If yn = 6 Then
    For b = RowNum To 7 Step -1
            '.Cells(b, 4).Select
            del = False
            If .Cells(b, 3).Value = "Produce" Then
                If .Cells(b, 16).Value = "Yes" And .Cells(b, 22).Value = "Yes" Then
                    del = True
                Else
                    If .Cells(b, 16).Value = "Yes" And .Cells(b, 22).Value = "No" Then
                        For c = 23 To 27
                            If .Cells(b, c).Value = "No" Then Exit For
                        Next
                        del = True
                    ElseIf .Cells(b, 16).Value = "No" And .Cells(b, 22).Value = "Yes" Then
                        For c = 17 To 21
                            If .Cells(b, c).Value = "No" Then Exit For
                        Next
                        del = True
                    Else
                        For c = 17 To 21
                            If .Cells(b, c).Value = "No" Then Exit For
                        Next
                        For c = 23 To 27
                            If .Cells(b, c).Value = "No" Then Exit For
                        Next
                        del = True
                    End If
                End If
            ElseIf (.Cells(b, 4).Value = "Yes" And .Cells(b, 5).Value = "Yes") Or (.Cells(b, 12).Value = "Yes" And .Cells(b, 13).Value = "Yes") Then
                del = True
            End If
            If del = True Then
                Application.DisplayAlerts = False
                Rows(b).Delete
                Application.DisplayAlerts = True
            End If
    Next
    End If


    If incCore = False Then Range(.Cells(1, 4), .Cells(1, 11)).EntireColumn.Hidden = True
    If incAlcohol = False Then Range(.Cells(1, 12), .Cells(1, 15)).EntireColumn.Hidden = True
    If incProduce = False Then Range(.Cells(1, 16), .Cells(1, 27)).EntireColumn.Hidden = True
    
    End If
    
    
    
    
    .Activate
    .Cells(7, 1).Select
    
    If APPOut(0, 0) <> 0 Then ricol = 28 Else ricol = 11
    If botrow = 0 Then botrow = 20
    .PageSetup.PrintArea = Range(.Cells(1, 1), .Cells(botrow, ricol)).Address
    .PageSetup.Zoom = False
    If ricol = 28 Then pwide = 2 Else pwide = 1
    .PageSetup.FitToPagesWide = pwide
    .PageSetup.FitToPagesTall = False
    .PageSetup.LeftFooter = "CORP BUYING, Admin, per: " & Format(Date, "DD/MM/YYYY") & Chr(10) & Application.ActiveWorkbook.FullName
    .PageSetup.Orientation = xlLandscape
    .PageSetup.PrintGridlines = True
    .PageSetup.RightFooter = "&P of &N"
    End With
    
    CBA_BasicFunctions.CBA_Close_Running
    
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBAR_MMRRuntime", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Function formatpriceprinciple(ByVal WS As Worksheet, ByVal crow As Long, ByVal ccol As Long)
    With WS.Cells(crow, ccol)
        .Font.ColorIndex = 1
        Select Case .Value
            Case Is <= 0.3
                .Interior.ColorIndex = 3
            Case Is <= 0.4
                .Interior.ColorIndex = 45
            Case Is <= 0.5
                .Interior.ColorIndex = 43
            Case Is > 0.5
                .Interior.ColorIndex = 4
        End Select
    End With
End Function
