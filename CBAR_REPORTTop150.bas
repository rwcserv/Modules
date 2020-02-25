Attribute VB_Name = "CBAR_REPORTTop150"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Type SubBasket
    alcohol As Single
    ambientnonfood As Single
    ambientfood As Single
    chilled As Single
    frozen As Single
    produce As Single
    meat As Single
End Type

Function Top150Run(ByVal DFrom As Date, ByVal qdate As Date, ByVal strProds As String, ByVal MHis As Boolean)
    Dim wks_T150 As Worksheet
    Dim bOutput As Boolean, curProd, needtocheck, CGno, testcmpcode, compDesc, strNatRetail
    Dim lNum As Long, CGnames, a As Long, j As Long, z As Long, addcol As Long, AldiRetailProduce
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    If CBA_BasicFunctions.isRunningSheetDisplayed Then
        CBA_BasicFunctions.CBA_Close_Running
        CBA_BasicFunctions.CBA_Running "Creating Top 150 Report"
    Else
        CBA_BasicFunctions.CBA_Running "Creating Top 150 Report"
    End If
    
    If CBA_COM_SetupMatchArray.CBA_SetupMatchArray(MHis, DFrom, qdate, , , strProds, True) = True Then
    
        If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 7, 5, "Your report is being finalized...."
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        CGnames = Range(CBAR_Data.Cells(1, 3), CBAR_Data.Cells(1, 3).End(xlDown)).Value
        bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("CBIS_ProdbyEmp")
        
        
        
        CBAR_T150.Copy
        Set wks_T150 = ActiveSheet
        
        With wks_T150
            'Range(.Cells(23, 1), .Cells(999, 70)).ClearContents
            lNum = 22
            For a = LBound(CBA_COM_Match) To UBound(CBA_COM_Match)
            
            'If CBA_COM_Match(a).AldiPCode = 61125 Then
            'a = a
            'End If
            
                If CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "National") = 0 Then GoTo NextMatch
                If CBA_COM_Match(a).compet = "C" Then
                    addcol = 0
                ElseIf CBA_COM_Match(a).compet = "WW" Then
                    addcol = 24
                ElseIf CBA_COM_Match(a).compet = "DM" Then
                    GoTo NextMatch
                End If
                If a = 0 Or CBA_COM_Match(a).AldiPCode <> curProd Then
                    needtocheck = False
                    curProd = CBA_COM_Match(a).AldiPCode
            '        If curprod = 3245 Then
            '        a = a
            '        End If
                Else
                    If addcol = 0 Then
                        If .Cells(lNum, 7).Value = "" Then
                            lNum = lNum - 1
                            needtocheck = False
                        Else
                            needtocheck = True
                        End If
                    End If
                    If addcol = 24 Then
                        If .Cells(lNum, 31).Value = "" Then
                            lNum = lNum - 1
                            needtocheck = False
                        Else
                            needtocheck = True
                        End If
                    End If
                End If
            
            '''''DEBUG CODE
            '    If CBA_COM_Match(a).AldiPCode = 76083 Then
            '    a = a
            '    End If
            
                If needtocheck = False Then
                    lNum = lNum + 1
                    .Cells(lNum, 1).Value = CBA_COM_Match(a).AldiPCode
                    .Cells(lNum, 2).Value = CBA_COM_Match(a).AldiPName
                    CGno = CBA_COM_Match(a).AldiPCG
                    If CGno = 58 Then
                        .Cells(lNum, 3).Value = "Produce"
                    Else
                        For j = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                            If CBA_CBISarr(0, j) = .Cells(lNum, 1).Value Then
                                .Cells(lNum, 3).Value = CBA_CBISarr(11, j)
                                Exit For
                            End If
                        Next
                    End If
                    
                    For j = LBound(CGnames, 1) To UBound(CGnames, 1)
                        If Mid(CGnames(j, 1), 1, 2) = CStr(Format(CGno, "00")) Then
                            .Cells(lNum, 4).Value = CGnames(j, 1)
                            Exit For
                        End If
                    Next
                    
                    
                    '.Cells(lNum, 4).Value = CGno
                    Select Case CGno
                        Case 1 To 4
                            .Cells(lNum, 5).Value = "Alcohol"
                        Case 5, 40 To 50, 52 To 57
                            .Cells(lNum, 5).Value = "Ambient Food"
                        Case 6 To 37, 39, 61, 65
                            .Cells(lNum, 5).Value = "Ambient Non-Food"
                        Case 51
                            .Cells(lNum, 5).Value = "Chilled"
                        Case 62, 64
                            .Cells(lNum, 5).Value = "Meat"
                        Case 38
                            .Cells(lNum, 5).Value = "Frozen"
                        Case 58
                            .Cells(lNum, 5).Value = "Produce"
                    End Select
                    If CGno <> 58 Then
                        If .Cells(lNum, 6).Value = "" Or .Cells(lNum, 6).Value = 0 Then .Cells(lNum, 6).Value = CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "National")
                    End If
                    
                    
                    'NATIONAL DATA
                    testcmpcode = CBA_COM_Match(a).CompCode
                    If Mid(Trim(.Cells(lNum, 4).Value), 1, 2) <> "58" Then
                        .Cells(lNum, 7 + addcol).Value = CBA_COM_Match(a).Pricedata(qdate, "ProRata", "National")
                        If .Cells(lNum, 6).Value = 0 Then .Cells(lNum, 8 + addcol).Value = 0 Else .Cells(lNum, 8 + addcol).Value = (.Cells(lNum, 7 + addcol).Value - .Cells(lNum, 6).Value) / .Cells(lNum, 6).Value
                        compDesc = CBA_COM_Match(a).CompProdName
                        If InStr(1, LCase(compDesc), "smart buy") > 0 Then
                            .Cells(lNum, 9 + addcol).Value = "Smartbuy"
                        ElseIf InStr(1, LCase(compDesc), "homebrand") > 0 Then
                            .Cells(lNum, 9 + addcol).Value = "Homebrand"
                        ElseIf InStr(1, LCase(compDesc), "select") > 0 Then
                            .Cells(lNum, 9 + addcol).Value = "Select"
                        ElseIf InStr(1, LCase(compDesc), "woolworths") > 0 Then
                            .Cells(lNum, 9 + addcol).Value = "Woolworths"
                        ElseIf InStr(1, LCase(compDesc), "coles") > 0 Then
                            .Cells(lNum, 9 + addcol).Value = "Coles"
                        Else
                            .Cells(lNum, 9 + addcol).Value = compDesc
                        End If
                        If CBA_COM_Match(a).Pricedata(qdate, "isspecial", "National") = True Then .Cells(lNum, 10 + addcol).Value = "Yes" Else .Cells(lNum, 10 + addcol).Value = "No"
                    End If
                    'NSW DATA
                    If Not IsEmpty(CBA_COM_Match(a).Pricedata(qdate, "ProRata", "NSW")) Then
                        .Cells(lNum, 11 + addcol).Value = CBA_COM_Match(a).Pricedata(qdate, "ProRata", "NSW")
                        If CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "NSW") = 0 Then .Cells(lNum, 12 + addcol).Value = 0 Else .Cells(lNum, 12 + addcol).Value = (.Cells(lNum, 11 + addcol).Value - CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "NSW")) / CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "NSW")
                        compDesc = CBA_COM_Match(a).CompProdName
                        If InStr(1, LCase(compDesc), "smart buy") > 0 Then
                            .Cells(lNum, 13 + addcol).Value = "Smartbuy"
                        ElseIf InStr(1, LCase(compDesc), "homebrand") > 0 Then
                            .Cells(lNum, 13 + addcol).Value = "Homebrand"
                        ElseIf InStr(1, LCase(compDesc), "select") > 0 Then
                            .Cells(lNum, 13 + addcol).Value = "Select"
                        ElseIf InStr(1, LCase(compDesc), "woolworths") > 0 Then
                            .Cells(lNum, 13 + addcol).Value = "Woolworths"
                        ElseIf InStr(1, LCase(compDesc), "coles") > 0 Then
                            .Cells(lNum, 13 + addcol).Value = "Coles"
                        Else
                            .Cells(lNum, 13 + addcol).Value = compDesc
                        End If
                        If CBA_COM_Match(a).Pricedata(qdate, "isspecial", "VIC") = True Then .Cells(lNum, 14 + addcol).Value = "Yes" Else .Cells(lNum, 14 + addcol).Value = "No"
                    End If
                    
                    'VIC DATA
                    If Not IsEmpty(CBA_COM_Match(a).Pricedata(qdate, "ProRata", "VIC")) Then
                        .Cells(lNum, 15 + addcol).Value = CBA_COM_Match(a).Pricedata(qdate, "ProRata", "VIC")
                        If CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "VIC") = 0 Then .Cells(lNum, 16 + addcol).Value = 0 Else .Cells(lNum, 16 + addcol).Value = (.Cells(lNum, 15 + addcol).Value - CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "VIC")) / CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "VIC")
                        compDesc = CBA_COM_Match(a).CompProdName
                        If InStr(1, LCase(compDesc), "smart buy") > 0 Then
                            .Cells(lNum, 17 + addcol).Value = "Smartbuy"
                        ElseIf InStr(1, LCase(compDesc), "homebrand") > 0 Then
                            .Cells(lNum, 17 + addcol).Value = "Homebrand"
                        ElseIf InStr(1, LCase(compDesc), "select") > 0 Then
                            .Cells(lNum, 17 + addcol).Value = "Select"
                        ElseIf InStr(1, LCase(compDesc), "woolworths") > 0 Then
                            .Cells(lNum, 17 + addcol).Value = "Woolworths"
                        ElseIf InStr(1, LCase(compDesc), "coles") > 0 Then
                            .Cells(lNum, 17 + addcol).Value = "Coles"
                        Else
                            .Cells(lNum, 17 + addcol).Value = compDesc
                        End If
                        If CBA_COM_Match(a).Pricedata(qdate, "isspecial", "VIC") = True Then .Cells(lNum, 18 + addcol).Value = "Yes" Else .Cells(lNum, 18 + addcol).Value = "No"
                    End If
                    
                    'QLD DATA
                    If Not IsEmpty(CBA_COM_Match(a).Pricedata(qdate, "ProRata", "QLD")) Then
                        .Cells(lNum, 19 + addcol).Value = CBA_COM_Match(a).Pricedata(qdate, "ProRata", "QLD")
                        If CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "QLD") = 0 Then .Cells(lNum, 20 + addcol).Value = 0 Else .Cells(lNum, 20 + addcol).Value = (.Cells(lNum, 19 + addcol).Value - CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "QLD")) / CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "QLD")
                        compDesc = CBA_COM_Match(a).CompProdName
                        If InStr(1, LCase(compDesc), "smart buy") > 0 Then
                            .Cells(lNum, 21 + addcol).Value = "Smartbuy"
                        ElseIf InStr(1, LCase(compDesc), "homebrand") > 0 Then
                            .Cells(lNum, 21 + addcol).Value = "Homebrand"
                        ElseIf InStr(1, LCase(compDesc), "select") > 0 Then
                            .Cells(lNum, 21 + addcol).Value = "Select"
                        ElseIf InStr(1, LCase(compDesc), "woolworths") > 0 Then
                            .Cells(lNum, 21 + addcol).Value = "Woolworths"
                        ElseIf InStr(1, LCase(compDesc), "coles") > 0 Then
                            .Cells(lNum, 21 + addcol).Value = "Coles"
                        Else
                            .Cells(lNum, 21 + addcol).Value = compDesc
                        End If
                        If CBA_COM_Match(a).Pricedata(qdate, "isspecial", "QLD") = True Then .Cells(lNum, 22 + addcol).Value = "Yes" Else .Cells(lNum, 22 + addcol).Value = "No"
                    End If
                    'SA DATA
                    If Not IsEmpty(CBA_COM_Match(a).Pricedata(qdate, "ProRata", "SA")) Then
                        .Cells(lNum, 23 + addcol).Value = CBA_COM_Match(a).Pricedata(qdate, "ProRata", "SA")
                        If CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "SA") = 0 Then .Cells(lNum, 24 + addcol).Value = 0 Else .Cells(lNum, 24 + addcol).Value = (.Cells(lNum, 23 + addcol).Value - CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "SA")) / CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "SA")
                        compDesc = CBA_COM_Match(a).CompProdName
                        If InStr(1, LCase(compDesc), "smart buy") > 0 Then
                            .Cells(lNum, 25 + addcol).Value = "Smartbuy"
                        ElseIf InStr(1, LCase(compDesc), "homebrand") > 0 Then
                            .Cells(lNum, 25 + addcol).Value = "Homebrand"
                        ElseIf InStr(1, LCase(compDesc), "select") > 0 Then
                            .Cells(lNum, 25 + addcol).Value = "Select"
                        ElseIf InStr(1, LCase(compDesc), "woolworths") > 0 Then
                            .Cells(lNum, 25 + addcol).Value = "Woolworths"
                        ElseIf InStr(1, LCase(compDesc), "coles") > 0 Then
                            .Cells(lNum, 25 + addcol).Value = "Coles"
                        Else
                            .Cells(lNum, 25 + addcol).Value = compDesc
                        End If
                        If CBA_COM_Match(a).Pricedata(qdate, "isspecial", "SA") = True Then .Cells(lNum, 26 + addcol).Value = "Yes" Else .Cells(lNum, 26 + addcol).Value = "No"
                    End If
                    'WA DATA
                    .Cells(lNum, 27 + addcol).Value = CBA_COM_Match(a).Pricedata(qdate, "ProRata", "WA")
                    If CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "WA") = 0 Then .Cells(lNum, 28 + addcol).Value = 0 Else .Cells(lNum, 28 + addcol).Value = (.Cells(lNum, 27 + addcol).Value - CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "WA")) / CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "WA")
                    compDesc = CBA_COM_Match(a).CompProdName
                    If InStr(1, LCase(compDesc), "smart buy") > 0 Then
                        .Cells(lNum, 29 + addcol).Value = "Smartbuy"
                    ElseIf InStr(1, LCase(compDesc), "homebrand") > 0 Then
                        .Cells(lNum, 29 + addcol).Value = "Homebrand"
                    ElseIf InStr(1, LCase(compDesc), "select") > 0 Then
                        .Cells(lNum, 29 + addcol).Value = "Select"
                    ElseIf InStr(1, LCase(compDesc), "woolworths") > 0 Then
                        .Cells(lNum, 29 + addcol).Value = "Woolworths"
                    ElseIf InStr(1, LCase(compDesc), "coles") > 0 Then
                        .Cells(lNum, 29 + addcol).Value = "Coles"
                    Else
                        .Cells(lNum, 29 + addcol).Value = compDesc
                    End If
                    If CBA_COM_Match(a).Pricedata(qdate, "isspecial", "WA") = True Then .Cells(lNum, 30 + addcol).Value = "Yes" Else .Cells(lNum, 30 + addcol).Value = "No"
                    'PRODUCE NATIONAL
                    If Mid(Trim(.Cells(lNum, 4).Value), 1, 2) = "58" Then
                        .Cells(lNum, 7 + addcol).Value = 9999999
                        For j = 11 + addcol To 27 + addcol
                            If j = 11 + addcol Or j = 15 + addcol Or j = 19 + addcol Or j = 23 + addcol Or j = 27 + addcol Then
                                If .Cells(lNum, 7 + addcol).Value > .Cells(lNum, j).Value And .Cells(lNum, j).Value > 0 Then
                                    For z = 0 To 3
                                    .Cells(lNum, 7 + z + addcol).Value = .Cells(lNum, j + z).Value
                                    Next
                                    Select Case j
                                        Case 11, 35
                                            strNatRetail = "NSW"
                                        Case 15, 39
                                            strNatRetail = "VIC"
                                        Case 19, 43
                                            strNatRetail = "QLD"
                                        Case 23, 47
                                            strNatRetail = "SA"
                                        Case 27, 51
                                            strNatRetail = "WA"
                                    End Select
                                End If
                                AldiRetailProduce = CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", strNatRetail)
                                If (AldiRetailProduce <> 0 And AldiRetailProduce < .Cells(lNum, 6).Value) Or .Cells(lNum, 6).Value = 0 Then .Cells(lNum, 6).Value = AldiRetailProduce
                            End If
                        Next
                    End If
                Else
                    'NATIONAL DATA
                    If Mid(Trim(.Cells(lNum, 4).Value), 1, 2) <> "58" Then
                        If (.Cells(lNum, 7 + addcol).Value > CBA_COM_Match(a).Pricedata(qdate, "ProRata", "National") And CBA_COM_Match(a).Pricedata(qdate, "ProRata", "National") <> 0) Or .Cells(lNum, 7 + addcol).Value = 0 Then
                            .Cells(lNum, 7 + addcol).Value = CBA_COM_Match(a).Pricedata(qdate, "ProRata", "National")
                            If .Cells(lNum, 6).Value = 0 Then .Cells(lNum, 8 + addcol).Value = 0 Else .Cells(lNum, 8 + addcol).Value = (.Cells(lNum, 7 + addcol).Value - .Cells(lNum, 6).Value) / .Cells(lNum, 6).Value
                            compDesc = CBA_COM_Match(a).CompProdName
                            If InStr(1, LCase(compDesc), "smart buy") > 0 Then
                                .Cells(lNum, 9 + addcol).Value = "Smartbuy"
                            ElseIf InStr(1, LCase(compDesc), "homebrand") > 0 Then
                                .Cells(lNum, 9 + addcol).Value = "Homebrand"
                            ElseIf InStr(1, LCase(compDesc), "select") > 0 Then
                                .Cells(lNum, 9 + addcol).Value = "Select"
                            ElseIf InStr(1, LCase(compDesc), "woolworths") > 0 Then
                                .Cells(lNum, 9 + addcol).Value = "Woolworths"
                            ElseIf InStr(1, LCase(compDesc), "coles") > 0 Then
                                .Cells(lNum, 9 + addcol).Value = "Coles"
                            Else
                                .Cells(lNum, 9 + addcol).Value = compDesc
                            End If
                            If CBA_COM_Match(a).Pricedata(qdate, "isspecial", "National") = True Then .Cells(lNum, 10 + addcol).Value = "Yes" Else .Cells(lNum, 10 + addcol).Value = "No"
                        End If
                    End If
                    'NSW DATA
                    If Not IsEmpty(CBA_COM_Match(a).Pricedata(qdate, "ProRata", "NSW")) Then
                        If (.Cells(lNum, 11 + addcol).Value > CBA_COM_Match(a).Pricedata(qdate, "ProRata", "NSW") And CBA_COM_Match(a).Pricedata(qdate, "ProRata", "NSW") <> 0) Or .Cells(lNum, 11 + addcol).Value = 0 Then
                            .Cells(lNum, 11 + addcol).Value = CBA_COM_Match(a).Pricedata(qdate, "ProRata", "NSW")
                            If CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "NSW") = 0 Then .Cells(lNum, 12 + addcol).Value = 0 Else .Cells(lNum, 12 + addcol).Value = (.Cells(lNum, 7 + addcol).Value - CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "NSW")) / CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "NSW")
                            compDesc = CBA_COM_Match(a).CompProdName
                            If InStr(1, LCase(compDesc), "smart buy") > 0 Then
                                .Cells(lNum, 13 + addcol).Value = "Smartbuy"
                            ElseIf InStr(1, LCase(compDesc), "homebrand") > 0 Then
                                .Cells(lNum, 13 + addcol).Value = "Homebrand"
                            ElseIf InStr(1, LCase(compDesc), "select") > 0 Then
                                .Cells(lNum, 13 + addcol).Value = "Select"
                            ElseIf InStr(1, LCase(compDesc), "woolworths") > 0 Then
                                .Cells(lNum, 13 + addcol).Value = "Woolworths"
                            ElseIf InStr(1, LCase(compDesc), "coles") > 0 Then
                                .Cells(lNum, 13 + addcol).Value = "Coles"
                            Else
                                .Cells(lNum, 13 + addcol).Value = compDesc
                            End If
                            If CBA_COM_Match(a).Pricedata(qdate, "isspecial", "NSW") = True Then .Cells(lNum, 14 + addcol).Value = "Yes" Else .Cells(lNum, 14 + addcol).Value = "No"
                        End If
                    End If
                    'VIC DATA
                    If Not IsEmpty(CBA_COM_Match(a).Pricedata(qdate, "ProRata", "VIC")) Then
                        If (.Cells(lNum, 15 + addcol).Value > CBA_COM_Match(a).Pricedata(qdate, "ProRata", "VIC") And CBA_COM_Match(a).Pricedata(qdate, "ProRata", "VIC") <> 0) Or .Cells(lNum, 15 + addcol).Value = 0 Then
                            .Cells(lNum, 15 + addcol).Value = CBA_COM_Match(a).Pricedata(qdate, "ProRata", "VIC")
                            If CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "VIC") = 0 Then .Cells(lNum, 16 + addcol).Value = 0 Else .Cells(lNum, 16 + addcol).Value = (.Cells(lNum, 7 + addcol).Value - CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "VIC")) / CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "VIC")
                            compDesc = CBA_COM_Match(a).CompProdName
                            If InStr(1, LCase(compDesc), "smart buy") > 0 Then
                                .Cells(lNum, 17 + addcol).Value = "Smartbuy"
                            ElseIf InStr(1, LCase(compDesc), "homebrand") > 0 Then
                                .Cells(lNum, 17 + addcol).Value = "Homebrand"
                            ElseIf InStr(1, LCase(compDesc), "select") > 0 Then
                                .Cells(lNum, 17 + addcol).Value = "Select"
                            ElseIf InStr(1, LCase(compDesc), "woolworths") > 0 Then
                                .Cells(lNum, 17 + addcol).Value = "Woolworths"
                            ElseIf InStr(1, LCase(compDesc), "coles") > 0 Then
                                .Cells(lNum, 17 + addcol).Value = "Coles"
                            Else
                                .Cells(lNum, 17 + addcol).Value = compDesc
                            End If
                            If CBA_COM_Match(a).Pricedata(qdate, "isspecial", "VIC") = True Then .Cells(lNum, 18 + addcol).Value = "Yes" Else .Cells(lNum, 18 + addcol).Value = "No"
                        End If
                    End If
                    'QLD DATA
                    If Not IsEmpty(CBA_COM_Match(a).Pricedata(qdate, "ProRata", "QLD")) Then
                        If (.Cells(lNum, 19 + addcol).Value > CBA_COM_Match(a).Pricedata(qdate, "ProRata", "QLD") And CBA_COM_Match(a).Pricedata(qdate, "ProRata", "QLD") <> 0) Or .Cells(lNum, 19 + addcol).Value = 0 Then
                            .Cells(lNum, 19 + addcol).Value = CBA_COM_Match(a).Pricedata(qdate, "ProRata", "QLD")
                            If CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "QLD") = 0 Then .Cells(lNum, 20 + addcol).Value = 0 Else .Cells(lNum, 20 + addcol).Value = (.Cells(lNum, 7 + addcol).Value - CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "QLD")) / CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "QLD")
                            compDesc = CBA_COM_Match(a).CompProdName
                            If InStr(1, LCase(compDesc), "smart buy") > 0 Then
                                .Cells(lNum, 21 + addcol).Value = "Smartbuy"
                            ElseIf InStr(1, LCase(compDesc), "homebrand") > 0 Then
                                .Cells(lNum, 21 + addcol).Value = "Homebrand"
                            ElseIf InStr(1, LCase(compDesc), "select") > 0 Then
                                .Cells(lNum, 21 + addcol).Value = "Select"
                            ElseIf InStr(1, LCase(compDesc), "woolworths") > 0 Then
                                .Cells(lNum, 21 + addcol).Value = "Woolworths"
                            ElseIf InStr(1, LCase(compDesc), "coles") > 0 Then
                                .Cells(lNum, 21 + addcol).Value = "Coles"
                            Else
                                .Cells(lNum, 21 + addcol).Value = compDesc
                            End If
                            If CBA_COM_Match(a).Pricedata(qdate, "isspecial", "QLD") = True Then .Cells(lNum, 22 + addcol).Value = "Yes" Else .Cells(lNum, 22 + addcol).Value = "No"
                        End If
                    End If
                    'SA DATA
                    If Not IsEmpty(CBA_COM_Match(a).Pricedata(qdate, "ProRata", "SA")) Then
                        If (.Cells(lNum, 23 + addcol).Value > CBA_COM_Match(a).Pricedata(qdate, "ProRata", "SA") And CBA_COM_Match(a).Pricedata(qdate, "ProRata", "SA") <> 0) Or .Cells(lNum, 23 + addcol).Value = 0 Then
                            .Cells(lNum, 23 + addcol).Value = CBA_COM_Match(a).Pricedata(qdate, "ProRata", "SA")
                            If CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "SA") = 0 Then .Cells(lNum, 24 + addcol).Value = 0 Else .Cells(lNum, 24 + addcol).Value = (.Cells(lNum, 7 + addcol).Value - CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "SA")) / CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "SA")
                            compDesc = CBA_COM_Match(a).CompProdName
                            If InStr(1, LCase(compDesc), "smart buy") > 0 Then
                                .Cells(lNum, 25 + addcol).Value = "Smartbuy"
                            ElseIf InStr(1, LCase(compDesc), "homebrand") > 0 Then
                                .Cells(lNum, 25 + addcol).Value = "Homebrand"
                            ElseIf InStr(1, LCase(compDesc), "select") > 0 Then
                                .Cells(lNum, 25 + addcol).Value = "Select"
                            ElseIf InStr(1, LCase(compDesc), "woolworths") > 0 Then
                                .Cells(lNum, 25 + addcol).Value = "Woolworths"
                            ElseIf InStr(1, LCase(compDesc), "coles") > 0 Then
                                .Cells(lNum, 25 + addcol).Value = "Coles"
                            Else
                                .Cells(lNum, 25 + addcol).Value = compDesc
                            End If
                            If CBA_COM_Match(a).Pricedata(qdate, "isspecial", "SA") = True Then .Cells(lNum, 26 + addcol).Value = "Yes" Else .Cells(lNum, 26 + addcol).Value = "No"
                        End If
                    End If
                    'WA DATA
                    If Not IsEmpty(CBA_COM_Match(a).Pricedata(qdate, "ProRata", "WA")) Then
                        If (.Cells(lNum, 27 + addcol).Value > CBA_COM_Match(a).Pricedata(qdate, "ProRata", "WA") And CBA_COM_Match(a).Pricedata(qdate, "ProRata", "WA") <> 0) Or .Cells(lNum, 27 + addcol).Value = 0 Then
                            .Cells(lNum, 27 + addcol).Value = CBA_COM_Match(a).Pricedata(qdate, "ProRata", "WA")
                            If CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "WA") = 0 Then .Cells(lNum, 28 + addcol).Value = 0 Else .Cells(lNum, 28 + addcol).Value = (.Cells(lNum, 7 + addcol).Value - CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "WA")) / CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", "WA")
                            compDesc = CBA_COM_Match(a).CompProdName
                            If InStr(1, LCase(compDesc), "smart buy") > 0 Then
                                .Cells(lNum, 29 + addcol).Value = "Smartbuy"
                            ElseIf InStr(1, LCase(compDesc), "homebrand") > 0 Then
                                .Cells(lNum, 29 + addcol).Value = "Homebrand"
                            ElseIf InStr(1, LCase(compDesc), "select") > 0 Then
                                .Cells(lNum, 29 + addcol).Value = "Select"
                            ElseIf InStr(1, LCase(compDesc), "woolworths") > 0 Then
                                .Cells(lNum, 29 + addcol).Value = "Woolworths"
                            ElseIf InStr(1, LCase(compDesc), "coles") > 0 Then
                                .Cells(lNum, 29 + addcol).Value = "Coles"
                            Else
                                .Cells(lNum, 29 + addcol).Value = compDesc
                            End If
                            If CBA_COM_Match(a).Pricedata(qdate, "isspecial", "WA") = True Then .Cells(lNum, 30 + addcol).Value = "Yes" Else .Cells(lNum, 30 + addcol).Value = "No"
                        End If
                    End If
                    'PRODUCE NATIONAL
                    If Mid(Trim(.Cells(lNum, 4).Value), 1, 2) = "58" Then
                        .Cells(lNum, 7 + addcol).Value = 9999999
                        For j = 11 + addcol To 27 + addcol
                            If j = 11 + addcol Or j = 15 + addcol Or j = 19 + addcol Or j = 23 + addcol Or j = 27 + addcol Then
                                If .Cells(lNum, 7 + addcol).Value > .Cells(lNum, j).Value And .Cells(lNum, j).Value > 0 Then
                                    For z = 0 To 3
                                    .Cells(lNum, 7 + z + addcol).Value = .Cells(lNum, j + z).Value
                                    Next
                                    Select Case j
                                        Case 11, 35
                                            strNatRetail = "NSW"
                                        Case 15, 39
                                            strNatRetail = "VIC"
                                        Case 19, 43
                                            strNatRetail = "QLD"
                                        Case 23, 47
                                            strNatRetail = "SA"
                                        Case 27, 51
                                            strNatRetail = "WA"
                                    End Select
                                End If
                                AldiRetailProduce = CBA_COM_Match(a).Pricedata(qdate, "AldiRetail", strNatRetail)
                                If (AldiRetailProduce <> 0 And AldiRetailProduce < .Cells(lNum, 6).Value) Or .Cells(lNum, 6).Value = 0 Then .Cells(lNum, 6).Value = AldiRetailProduce
                            End If
                        Next
                    End If
                End If
            
            
            
NextMatch:
            Next
            
            CBAR_SQLQueries.CBAR_GenPullSQL "CBAR_CompSKUCount", DFrom
            For a = 0 To UBound(CBA_COMarr, 2)
                If CBA_COMarr(0, a) = "Coles" Then
                    .Cells(13 + a, 4).Value = CBA_COMarr(2, a)
                Else
                    .Cells(14 + a, 4).Value = CBA_COMarr(2, a)
                End If
            Next
            
            .Activate
            Range(.Cells(23, 55), .Cells(23, 70)).Copy
            Range(.Cells(24, 55), .Cells(172, 70)).PasteSpecial (xlPasteAll)
            Application.CutCopyMode = False
            
            
            .Cells(23, 1).Select
            If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
        End With
    End If
    
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-Top150Run", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function
