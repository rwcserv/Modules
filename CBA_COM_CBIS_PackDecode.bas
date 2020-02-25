Attribute VB_Name = "CBA_COM_CBIS_PackDecode"
'*******************CBA_COM_CBIS_PackDecode v0.01*********************************
Option Explicit
Option Private Module          ' Excel users cannot access procedures

Function CBA_COM_genCBISoutput(ByVal useInputArray As Boolean, ByRef Inputarr() As Variant, Optional ByVal PCode As Long, Optional ByVal strDescription As String) As Boolean
    Dim arrfordecode(), tempOutput() As Variant
    Dim random As Variant, bOutput As Boolean, a As Long, b As Long, c As Long, d As Long, e As Long
    Dim InputRows, adjustrow, adjustcol As Long
    Dim arrfromContents(), arrfromDescripton(), arrfromPCDesc(), arrfromPCDouble() As Variant
    Dim Pieceval, gramval, Mmval, pageval, litreval, sheetval, CompPieceval, Compgramval, CompMmval, Comppageval, Complitreval, Compsheetval As Single
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    adjustrow = 0
    adjustcol = 0
    On Error GoTo alldone
    random = Inputarr(0, 0)
    adjustrow = 1
    adjustcol = 1
alldone:
    Err.Clear
    On Error GoTo Err_Routine
    
    If useInputArray = True Then
        ReDim arrfordecode(1 To 2, 1 To 1)
        ReDim tempOutput(1 To 29, 1 To 1)
                For a = 1 To UBound(Inputarr, 2) + adjustrow
                    If IsEmpty(Inputarr(1 - adjustcol, a - adjustrow)) Then Exit For
                    ReDim Preserve arrfordecode(1 To 2, 1 To a)
                    ReDim Preserve tempOutput(1 To 29, 1 To a)
                    InputRows = a
                    arrfordecode(1, a) = Inputarr(1 - adjustcol, a - adjustrow)
                    arrfordecode(2, a) = Inputarr(16 - adjustcol, a - adjustrow)
                    tempOutput(1, a) = Inputarr(1 - adjustcol, a - adjustrow)
                Next
                    
                    'bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_decodeCBISdescription(True, Inputarr(), 94451, "AO: Soft Grip Clothes Pegs 20p")
                    bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_decodeCBISdescription(True, arrfordecode())
                    If bOutput = True Then
                        arrfromContents = CBA_COM_CBISarrOutput
                        Erase CBA_COM_CBISarrOutput
                    Else
                        Exit Function
                    End If
                ReDim arrfordecode(1 To 2, 1 To InputRows)
                For a = 1 To InputRows
                    arrfordecode(1, a) = Inputarr(1 - adjustcol, a - adjustrow)
                    arrfordecode(2, a) = Inputarr(20 - adjustcol, a - adjustrow)
                Next
                    bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_decodeCBISdescription(True, arrfordecode())
                        If bOutput = True Then
                        arrfromPCDesc = CBA_COM_CBISarrOutput
                        Erase CBA_COM_CBISarrOutput
                    Else
                        Exit Function
                    End If
                ReDim arrfordecode(1 To 2, 1 To InputRows)
                For a = 1 To InputRows
                    arrfordecode(1, a) = Inputarr(1 - adjustcol, a - adjustrow)
                    arrfordecode(2, a) = Inputarr(21 - adjustcol, a - adjustrow)
                Next
                    bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_decodeCBISdescription(True, arrfordecode())
                        If bOutput = True Then
                        arrfromPCDouble = CBA_COM_CBISarrOutput
                        Erase CBA_COM_CBISarrOutput
                    Else
                        Exit Function
                    End If
                'added if line to stop function from finding unwanted packsizes in CBIS
                ReDim arrfordecode(1 To 2, 1 To InputRows)
                For a = 1 To InputRows
                    If Inputarr(1 - adjustcol, 1 - adjustrow) <> 61936 Then
                        arrfordecode(1, a) = Inputarr(1 - adjustcol, a - adjustrow)
                        arrfordecode(2, a) = Inputarr(5 - adjustcol, a - adjustrow)
                    End If
                Next
                    bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_decodeCBISdescription(True, arrfordecode())
                    If bOutput = True Then
                        arrfromDescripton = CBA_COM_CBISarrOutput
                        Erase CBA_COM_CBISarrOutput
                    Else
                        Exit Function
                    End If
                Erase arrfordecode
                For a = 1 To InputRows
                    b = 2
                    If UCase(Inputarr(9 - adjustcol, a - adjustrow)) = "G" Then
                        If Inputarr(7 - adjustcol, a - adjustrow) <> 0 Then
                            tempOutput(b, a) = Inputarr(7 - adjustcol, a - adjustrow)
                            b = b + 1
                            tempOutput(b, a) = "g"
                            b = b + 1
                        End If
                    End If
                    If UCase(Inputarr(7 - adjustcol, a - adjustrow)) = "G" Then
                        If Inputarr(17 - adjustcol, a - adjustrow) <> 0 Then
                            tempOutput(b, a) = Inputarr(17 - adjustcol, a - adjustrow)
                            b = b + 1
                            tempOutput(b, a) = "g"
                            b = b + 1
                        End If
                    End If
                    If UCase(Inputarr(9 - adjustcol, a - adjustrow)) = "KG" Then
                        If Inputarr(7 - adjustcol, a - adjustrow) <> 0 And IsNumeric(Inputarr(7 - adjustcol, a - adjustrow)) Then
                            tempOutput(b, a) = Inputarr(7 - adjustcol, a - adjustrow) * 1000
                            b = b + 1
                            tempOutput(b, a) = "g"
                            b = b + 1
                        End If
                    End If
                    If UCase(Inputarr(7 - adjustcol, a - adjustrow)) = "KG" Then
                        If Inputarr(17 - adjustcol, a - adjustrow) <> 0 Or Inputarr(17 - adjustcol, a - adjustrow) = Null Then
                            tempOutput(b, a) = Inputarr(17 - adjustcol, a - adjustrow) * 1000
                            b = b + 1
                            tempOutput(b, a) = "g"
                            b = b + 1
                        End If
                    End If
    '                If UCase(Inputarr(8 - adjustcol, a - adjustrow)) <> 0 And Not UCase(Inputarr(8 - adjustcol, a - adjustrow)) = "NULL" Then
    '                    If Inputarr(8 - adjustcol, a - adjustrow) <> 0 And IsNumeric(Inputarr(8 - adjustcol, a - adjustrow)) Then
    '                        tempOutput(b, a) = Inputarr(8 - adjustcol, a - adjustrow) * 1000
    '                        b = b + 1
    '                        tempOutput(b, a) = "g"
    '                        b = b + 1
    '                    End If
    '                End If
                    If UCase(Inputarr(8 - adjustcol, a - adjustrow)) = "ML" Then
                        If Inputarr(19 - adjustcol, a - adjustrow) = 0 Or Inputarr(19 - adjustcol, a - adjustrow) = "NULL" Then
                        Else
                            tempOutput(b, a) = Inputarr(19 - adjustcol, a - adjustrow)
                            b = b + 1
                            tempOutput(b, a) = "ml"
                            b = b + 1
                        End If
                    End If
                    If UCase(Inputarr(8 - adjustcol, a - adjustrow)) = "LTR" Then
                        If Inputarr(19 - adjustcol, a - adjustrow) = 0 Or Inputarr(19 - adjustcol, a - adjustrow) = "NULL" Then
                        Else
                            tempOutput(b, a) = Inputarr(19 - adjustcol, a - adjustrow) * 1000
                            b = b + 1
                            tempOutput(b, a) = "ml"
                            b = b + 1
                        End If
                    End If
                Next
    
    
    'arrfromContents
    'arrfromPCDesc
    'arrfromPCDouble
    'arrfromDescripton
    'tempOutput
    
    
        'ReDim CBA_COM_CBISarrOutput(1 To 999, 1 To InputRows)
        ReDim CBA_COM_CBISarrOutput(1 To 199, 1 To InputRows)
        
        For a = 1 To InputRows
            CBA_COM_CBISarrOutput(1, a) = Inputarr(1 - adjustcol, a - adjustrow)
            CBA_COM_CBISarrOutput(2, a) = Inputarr(5 - adjustcol, a - adjustrow)
        Next
    
        For a = 1 To InputRows
    '    If CBA_COM_CBISarrOutput(1, a) = 76016 Then
    '    a = a
    '    End If
            d = 2
            For b = 1 To UBound(arrfromContents, 2)
                If arrfromContents(1, b) = CBA_COM_CBISarrOutput(1, a) Then
                    
                    For c = 3 To 29
                        If IsEmpty(arrfromContents(c, b)) Then
                            Exit For
                        Else
                            d = d + 1
                            CBA_COM_CBISarrOutput(d, a) = arrfromContents(c, b)
                        End If
                    Next
                Exit For
                End If
            Next
    
            For b = 1 To UBound(arrfromPCDesc, 2)
                If arrfromPCDesc(1, b) = CBA_COM_CBISarrOutput(1, a) Then
                    For c = 3 To 29
                        If IsEmpty(arrfromPCDesc(c, b)) Then
                            Exit For
                        Else
                            d = d + 1
                            CBA_COM_CBISarrOutput(d, a) = arrfromPCDesc(c, b)
                        End If
                    Next
                Exit For
                End If
            Next
    
            For b = 1 To UBound(arrfromPCDouble, 2)
                If arrfromPCDouble(1, b) = CBA_COM_CBISarrOutput(1, a) Then
                    For c = 3 To 29
                        If IsEmpty(arrfromPCDouble(c, b)) Then
                            Exit For
                        Else
                            d = d + 1
                            CBA_COM_CBISarrOutput(d, a) = arrfromPCDouble(c, b)
                        End If
                    Next
                Exit For
                End If
            Next
            
            
            For b = 1 - adjustcol To UBound(Inputarr, 1)
                If Inputarr(b, LBound(Inputarr, 1)) = CBA_COM_CBISarrOutput(1, a) Then
                    If Inputarr(10 - adjustcol, b) = 1 And Inputarr(22 - adjustcol, b) = 4 Then
                        d = d + 1
                        CBA_COM_CBISarrOutput(d, a) = 1000
                        d = d + 1
                        CBA_COM_CBISarrOutput(d, a) = "g"
                    End If
                    Exit For
                End If
            Next
            
            
            
            
        Next
        
        Erase arrfromContents
        Erase arrfromPCDesc
        Erase arrfromPCDouble
        
        For a = 1 To InputRows
        
    
            Pieceval = 0
            gramval = 0
            Mmval = 0
            pageval = 0
            litreval = 0
            sheetval = 0
            For b = 3 To 199
                If IsEmpty(CBA_COM_CBISarrOutput(b, a)) Then
                    Exit For
                Else
                    
                    If IsNumeric(CBA_COM_CBISarrOutput(b, a)) And CBA_COM_CBISarrOutput(b + 1, a) = "Pieces" Then
                        If Pieceval = 0 Then
                            Pieceval = CSng(CBA_COM_CBISarrOutput(b, a))
                        Else
                            If CSng(CBA_COM_CBISarrOutput(b, a)) > Pieceval Then Pieceval = CSng(CBA_COM_CBISarrOutput(b, a))
                        End If
                    ElseIf IsNumeric(CBA_COM_CBISarrOutput(b, a)) And CBA_COM_CBISarrOutput(b + 1, a) = "g" Then
                        If gramval = 0 Then
                            gramval = CSng(CBA_COM_CBISarrOutput(b, a))
                        Else
                            If CSng(CBA_COM_CBISarrOutput(b, a)) > gramval Then gramval = CSng(CBA_COM_CBISarrOutput(b, a))
                        End If
                    ElseIf IsNumeric(CBA_COM_CBISarrOutput(b, a)) And CBA_COM_CBISarrOutput(b + 1, a) = "mm" Then
                        If Mmval = 0 Then
                            Mmval = CSng(CBA_COM_CBISarrOutput(b, a))
                        Else
                            If CSng(CBA_COM_CBISarrOutput(b, a)) > Mmval Then Mmval = CSng(CBA_COM_CBISarrOutput(b, a))
                        End If
                    ElseIf IsNumeric(CBA_COM_CBISarrOutput(b, a)) And CBA_COM_CBISarrOutput(b + 1, a) = "Page" Then
                        If pageval = 0 Then
                            pageval = CSng(CBA_COM_CBISarrOutput(b, a))
                        Else
                            If CSng(CBA_COM_CBISarrOutput(b, a)) > pageval Then pageval = CSng(CBA_COM_CBISarrOutput(b, a))
                        End If
                    ElseIf IsNumeric(CBA_COM_CBISarrOutput(b, a)) And CBA_COM_CBISarrOutput(b + 1, a) = "ml" Then
                        If litreval = 0 Then
                            litreval = CSng(CBA_COM_CBISarrOutput(b, a))
                        Else
                            If CSng(CBA_COM_CBISarrOutput(b, a)) > litreval Then litreval = CSng(CBA_COM_CBISarrOutput(b, a))
                        End If
                    ElseIf IsNumeric(CBA_COM_CBISarrOutput(b, a)) And CBA_COM_CBISarrOutput(b + 1, a) = "Sheet" Then
                        If sheetval = 0 Then
                            sheetval = CSng(CBA_COM_CBISarrOutput(b, a))
                        Else
                            If CSng(CBA_COM_CBISarrOutput(b, a)) > sheetval Then sheetval = CSng(CBA_COM_CBISarrOutput(b, a))
                        End If
                    End If
                   CBA_COM_CBISarrOutput(b, a) = Empty
                
                End If
            Next
            c = 4
            If Pieceval > 0 Then
                CBA_COM_CBISarrOutput(3, a) = Pieceval
                CBA_COM_CBISarrOutput(4, a) = "Pieces"
            End If
            If gramval > 0 Then
                c = c + 1
                CBA_COM_CBISarrOutput(c, a) = gramval
                c = c + 1
                CBA_COM_CBISarrOutput(c, a) = "g"
            End If
            If Mmval > 0 Then
                c = c + 1
                CBA_COM_CBISarrOutput(c, a) = Mmval
                c = c + 1
                CBA_COM_CBISarrOutput(c, a) = "mm"
            End If
            If pageval > 0 Then
                c = c + 1
                CBA_COM_CBISarrOutput(c, a) = pageval
                c = c + 1
                CBA_COM_CBISarrOutput(c, a) = "Pages"
            End If
            If litreval > 0 Then
                c = c + 1
                CBA_COM_CBISarrOutput(c, a) = litreval
                c = c + 1
                CBA_COM_CBISarrOutput(c, a) = "ml"
            End If
            If sheetval > 0 Then
                c = c + 1
                CBA_COM_CBISarrOutput(c, a) = sheetval
                c = c + 1
                CBA_COM_CBISarrOutput(c, a) = "Sheets"
            End If
        Next
        
    
    
    
    
        For a = 1 To InputRows
    
            d = 1
            For b = 1 To UBound(arrfromDescripton, 2)
                If arrfromDescripton(1, b) = CBA_COM_CBISarrOutput(1, a) Then
                    Pieceval = 0
                    gramval = 0
                    Mmval = 0
                    pageval = 0
                    litreval = 0
                    sheetval = 0
                    CompPieceval = 0
                    Compgramval = 0
                    CompMmval = 0
                    Comppageval = 0
                    Complitreval = 0
                    Compsheetval = 0
                    For e = 3 To 999
                        If IsEmpty(CBA_COM_CBISarrOutput(e, a)) Then Exit For
                        If IsNumeric(CBA_COM_CBISarrOutput(e, a)) And CBA_COM_CBISarrOutput(e + 1, a) = "Pieces" Then
                            If Pieceval = 0 Then
                                Pieceval = CSng(CBA_COM_CBISarrOutput(e, a))
                            Else
                                If CSng(CBA_COM_CBISarrOutput(e, a)) > Pieceval Then Pieceval = CSng(CBA_COM_CBISarrOutput(e, a))
                            End If
                        ElseIf IsNumeric(CBA_COM_CBISarrOutput(e, a)) And CBA_COM_CBISarrOutput(e + 1, a) = "g" Then
                            If gramval = 0 Then
                                gramval = CSng(CBA_COM_CBISarrOutput(e, a))
                            Else
                                If CSng(CBA_COM_CBISarrOutput(e, a)) > gramval Then gramval = CSng(CBA_COM_CBISarrOutput(e, a))
                            End If
                        ElseIf IsNumeric(CBA_COM_CBISarrOutput(e, a)) And CBA_COM_CBISarrOutput(e + 1, a) = "mm" Then
                            If Mmval = 0 Then
                                Mmval = CSng(CBA_COM_CBISarrOutput(e, a))
                            Else
                                If CSng(CBA_COM_CBISarrOutput(e, a)) > Mmval Then Mmval = CSng(CBA_COM_CBISarrOutput(e, a))
                            End If
                        ElseIf IsNumeric(CBA_COM_CBISarrOutput(e, a)) And CBA_COM_CBISarrOutput(e + 1, a) = "Page" Then
                            If pageval = 0 Then
                                pageval = CSng(CBA_COM_CBISarrOutput(e, a))
                            Else
                                If CSng(CBA_COM_CBISarrOutput(e, a)) > pageval Then pageval = CSng(CBA_COM_CBISarrOutput(e, a))
                            End If
                        ElseIf IsNumeric(CBA_COM_CBISarrOutput(e, a)) And CBA_COM_CBISarrOutput(e + 1, a) = "ml" Then
                            If litreval = 0 Then
                                litreval = CSng(CBA_COM_CBISarrOutput(e, a))
                            Else
                                If CSng(CBA_COM_CBISarrOutput(e, a)) > litreval Then litreval = CSng(CBA_COM_CBISarrOutput(e, a))
                            End If
                        ElseIf IsNumeric(CBA_COM_CBISarrOutput(e, a)) And CBA_COM_CBISarrOutput(e + 1, a) = "Sheet" Then
                            If sheetval = 0 Then
                                sheetval = CSng(CBA_COM_CBISarrOutput(e, a))
                            Else
                                If CSng(CBA_COM_CBISarrOutput(e, a)) > sheetval Then sheetval = CSng(CBA_COM_CBISarrOutput(e, a))
                            End If
                        End If
                        CBA_COM_CBISarrOutput(e, a) = Empty
                    Next
                    For c = 3 To 29
                        If IsEmpty(arrfromDescripton(c, b)) Then Exit For
                        If IsNumeric(arrfromDescripton(c, b)) And arrfromDescripton(c + 1, b) = "Pieces" Then
                            If CompPieceval = 0 Then
                                CompPieceval = CSng(arrfromDescripton(c, b))
                            Else
                                If CSng(arrfromDescripton(c, b)) > CompPieceval Then CompPieceval = CSng(arrfromDescripton(c, b))
                            End If
                        ElseIf IsNumeric(arrfromDescripton(c, b)) And arrfromDescripton(c + 1, b) = "g" Then
                            If Compgramval = 0 Then
                                Compgramval = CSng(arrfromDescripton(c, b))
                            Else
                                If CSng(arrfromDescripton(c, b)) > Compgramval Then Compgramval = CSng(arrfromDescripton(c, b))
                            End If
                        ElseIf IsNumeric(arrfromDescripton(c, b)) And arrfromDescripton(c + 1, a) = "mm" Then
                            If CompMmval = 0 Then
                                CompMmval = CSng(arrfromDescripton(c, b))
                            Else
                                If CSng(arrfromDescripton(c, b)) > CompMmval Then CompMmval = CSng(arrfromDescripton(c, b))
                            End If
                        ElseIf IsNumeric(arrfromDescripton(c, b)) And arrfromDescripton(c + 1, a) = "Page" Then
                            If Comppageval = 0 Then
                                Comppageval = CSng(arrfromDescripton(c, b))
                            Else
                                If CSng(arrfromDescripton(c, b)) > Comppageval Then Comppageval = CSng(arrfromDescripton(c, b))
                            End If
                        ElseIf IsNumeric(arrfromDescripton(c, b)) And arrfromDescripton(c + 1, a) = "ml" Then
                            If Complitreval = 0 Then
                                Complitreval = CSng(arrfromDescripton(c, b))
                            Else
                                If CSng(arrfromDescripton(c, b)) > Complitreval Then Complitreval = CSng(arrfromDescripton(c, b))
                            End If
                        ElseIf IsNumeric(arrfromDescripton(c, b)) And arrfromDescripton(c + 1, a) = "Sheet" Then
                            If Compsheetval = 0 Then
                                Compsheetval = CSng(arrfromDescripton(c, b))
                            Else
                                If CSng(arrfromDescripton(c, b)) > Compsheetval Then Compsheetval = CSng(arrfromDescripton(c, b))
                            End If
                        End If
                    Next
    
                    c = 4
                    If CSng(CompPieceval) > CSng(Pieceval) Then
                        CBA_COM_CBISarrOutput(3, a) = CompPieceval
                        CBA_COM_CBISarrOutput(4, a) = "Pieces"
                    Else
                        CBA_COM_CBISarrOutput(3, a) = Pieceval
                        CBA_COM_CBISarrOutput(4, a) = "Pieces"
                    End If
                    If gramval > 0 Or Compgramval > 0 Then
                        If Compgramval > 0 Then
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = Compgramval
                        Else
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = gramval
                        End If
                        c = c + 1
                        CBA_COM_CBISarrOutput(c, a) = "g"
                    End If
                    If Mmval > 0 Or CompMmval > 0 Then
                        If CompMmval > 0 Then
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = CompMmval
                        Else
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = Mmval
                        End If
                        c = c + 1
                        CBA_COM_CBISarrOutput(c, a) = "mm"
                    End If
                    If pageval > 0 Or Comppageval > 0 Then
                        If Comppageval > 0 Then
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = Comppageval
                        Else
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = pageval
                        End If
                        c = c + 1
                        CBA_COM_CBISarrOutput(c, a) = "Page"
                    End If
                    If litreval > 0 Or Complitreval > 0 Then
                        If Complitreval > 0 Then
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = Complitreval
                        Else
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = litreval
                        End If
                        c = c + 1
                        CBA_COM_CBISarrOutput(c, a) = "ml"
                    End If
                    If sheetval > 0 Or Compsheetval > 0 Then
                        If Compsheetval > 0 Then
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = Compsheetval
                        Else
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = sheetval
                        End If
                        c = c + 1
                        CBA_COM_CBISarrOutput(c, a) = "Sheet"
                    End If
                Exit For
                End If
                
            Next
        Next
        Erase arrfromDescripton
        For a = 1 To InputRows
    
            d = 1
            For b = 1 To UBound(tempOutput, 2)
                If tempOutput(1, b) = CBA_COM_CBISarrOutput(1, a) Then
                    Pieceval = 0
                    gramval = 0
                    Mmval = 0
                    pageval = 0
                    litreval = 0
                    sheetval = 0
                    CompPieceval = 0
                    Compgramval = 0
                    CompMmval = 0
                    Comppageval = 0
                    Complitreval = 0
                    Compsheetval = 0
                    For e = 3 To 999
                        If IsEmpty(CBA_COM_CBISarrOutput(e, a)) Then Exit For
                        If IsNumeric(CBA_COM_CBISarrOutput(e, a)) And CBA_COM_CBISarrOutput(e + 1, a) = "Pieces" Then
                            If Pieceval = 0 Then
                                Pieceval = CSng(CBA_COM_CBISarrOutput(e, a))
                            Else
                                If CSng(CBA_COM_CBISarrOutput(e, a)) > Pieceval Then Pieceval = CSng(CBA_COM_CBISarrOutput(e, a))
                            End If
                        ElseIf IsNumeric(CBA_COM_CBISarrOutput(e, a)) And CBA_COM_CBISarrOutput(e + 1, a) = "g" Then
                            If gramval = 0 Then
                                gramval = CSng(CBA_COM_CBISarrOutput(e, a))
                            Else
                                If CSng(CBA_COM_CBISarrOutput(e, a)) > gramval Then gramval = CSng(CBA_COM_CBISarrOutput(e, a))
                            End If
                        ElseIf IsNumeric(CBA_COM_CBISarrOutput(e, a)) And CBA_COM_CBISarrOutput(e + 1, a) = "mm" Then
                            If Mmval = 0 Then
                                Mmval = CSng(CBA_COM_CBISarrOutput(e, a))
                            Else
                                If CSng(CBA_COM_CBISarrOutput(e, a)) > Mmval Then Mmval = CSng(CBA_COM_CBISarrOutput(e, a))
                            End If
                        ElseIf IsNumeric(CBA_COM_CBISarrOutput(e, a)) And CBA_COM_CBISarrOutput(e + 1, a) = "Page" Then
                            If pageval = 0 Then
                                pageval = CSng(CBA_COM_CBISarrOutput(e, a))
                            Else
                                If CSng(CBA_COM_CBISarrOutput(e, a)) > pageval Then pageval = CSng(CBA_COM_CBISarrOutput(e, a))
                            End If
                        ElseIf IsNumeric(CBA_COM_CBISarrOutput(e, a)) And CBA_COM_CBISarrOutput(e + 1, a) = "ml" Then
                            If litreval = 0 Then
                                litreval = CSng(CBA_COM_CBISarrOutput(e, a))
                            Else
                                If CSng(CBA_COM_CBISarrOutput(e, a)) > litreval Then litreval = CSng(CBA_COM_CBISarrOutput(e, a))
                            End If
                        ElseIf IsNumeric(CBA_COM_CBISarrOutput(e, a)) And CBA_COM_CBISarrOutput(e + 1, a) = "Sheet" Then
                            If sheetval = 0 Then
                                sheetval = CSng(CBA_COM_CBISarrOutput(e, a))
                            Else
                                If CSng(CBA_COM_CBISarrOutput(e, a)) > sheetval Then sheetval = CSng(CBA_COM_CBISarrOutput(e, a))
                            End If
                        End If
                        CBA_COM_CBISarrOutput(e, a) = Empty
                    Next
                    For c = 2 To 29
                        If IsEmpty(tempOutput(c, b)) Then Exit For
                        If IsNumeric(tempOutput(c, b)) And tempOutput(c + 1, b) = "Pieces" Then
                            If CompPieceval = 0 Then
                                CompPieceval = CSng(tempOutput(c, b))
                            Else
                                If CSng(tempOutput(c, b)) > CompPieceval Then CompPieceval = CSng(tempOutput(c, b))
                            End If
                        ElseIf IsNumeric(tempOutput(c, b)) And tempOutput(c + 1, b) = "g" Then
                            If Compgramval = 0 Then
                                Compgramval = CSng(tempOutput(c, b))
                            Else
                                If CSng(tempOutput(c, b)) > Compgramval Then Compgramval = CSng(tempOutput(c, b))
                            End If
                        ElseIf IsNumeric(tempOutput(c, b)) And tempOutput(c + 1, b) = "mm" Then
                            If CompMmval = 0 Then
                                CompMmval = CSng(tempOutput(c, b))
                            Else
                                If CSng(tempOutput(c, b)) > CompMmval Then CompMmval = CSng(tempOutput(c, b))
                            End If
                        ElseIf IsNumeric(tempOutput(c, b)) And tempOutput(c + 1, b) = "Page" Then
                            If Comppageval = 0 Then
                                Comppageval = CSng(tempOutput(c, b))
                            Else
                                If CSng(tempOutput(c, b)) > Comppageval Then Comppageval = CSng(tempOutput(c, b))
                            End If
                        ElseIf IsNumeric(tempOutput(c, b)) And tempOutput(c + 1, b) = "ml" Then
                            If Complitreval = 0 Then
                                Complitreval = CSng(tempOutput(c, b))
                            Else
                                If CSng(tempOutput(c, b)) > Complitreval Then Complitreval = CSng(tempOutput(c, b))
                            End If
                        ElseIf IsNumeric(tempOutput(c, b)) And tempOutput(c + 1, b) = "Sheet" Then
                            If Compsheetval = 0 Then
                                Compsheetval = CSng(tempOutput(c, b))
                            Else
                                If CSng(tempOutput(c, b)) > Compsheetval Then Compsheetval = CSng(tempOutput(c, b))
                            End If
                        End If
                    Next
                    c = 4
                    If CompPieceval > Pieceval Then
                        CBA_COM_CBISarrOutput(3, a) = CompPieceval
                        CBA_COM_CBISarrOutput(4, a) = "Pieces"
                    Else
                        CBA_COM_CBISarrOutput(3, a) = Pieceval
                        CBA_COM_CBISarrOutput(4, a) = "Pieces"
                    End If
                    If gramval > 0 Or Compgramval > 0 Then
                        If Compgramval > 0 Then
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = Compgramval
                        Else
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = gramval
                        End If
                        c = c + 1
                        CBA_COM_CBISarrOutput(c, a) = "g"
                    End If
                    If Mmval > 0 Or CompMmval > 0 Then
                        If CompMmval > 0 Then
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = CompMmval
                        Else
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = Mmval
                        End If
                        c = c + 1
                        CBA_COM_CBISarrOutput(c, a) = "mm"
                    End If
                    If pageval > 0 Or Comppageval > 0 Then
                        If Comppageval > 0 Then
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = Comppageval
                        Else
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = pageval
                        End If
                        c = c + 1
                        CBA_COM_CBISarrOutput(c, a) = "Page"
                    End If
                    If litreval > 0 Or Complitreval > 0 Then
                        If Complitreval > 0 Then
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = Complitreval
                        Else
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = litreval
                        End If
                        c = c + 1
                        CBA_COM_CBISarrOutput(c, a) = "ml"
                    End If
                    If sheetval > 0 Or Compsheetval > 0 Then
                        If Compsheetval > 0 Then
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = Compsheetval
                        Else
                            c = c + 1
                            CBA_COM_CBISarrOutput(c, a) = sheetval
                        End If
                        c = c + 1
                        CBA_COM_CBISarrOutput(c, a) = "Sheet"
                    End If
                Exit For
                End If
            Next
            
            
            'HARD CODED PACKSIZES
            If CBA_COM_CBISarrOutput(1, a) = 59663 Then
                CBA_COM_CBISarrOutput(3, a) = 30
                CBA_COM_CBISarrOutput(4, a) = "Pieces"
            End If
        Next
        Erase tempOutput
    Else
    
    InputRows = 1
                ReDim arrfordecode(1 To 2, 1)
                    arrfordecode(1, a) = PCode
                    arrfordecode(2, a) = strDescription
                    bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_decodeCBISdescription(False, arrfordecode(), PCode, strDescription)
                    If bOutput = True Then
                        arrfromDescripton = CBA_COM_CBISarrOutput
                    Else
                        Exit Function
                    End If
    
    
    End If
    
    
    
    ''outputcode
    '    wks_Data.Cells.ClearContents
    '    For a = 1 To UBound(CBA_COM_CBISarrOutput, 2)
    '        For b = 1 To 29
    '            If IsEmpty(CBA_COM_CBISarrOutput(b, a)) Then Else wks_Data.Cells(a, b) = CBA_COM_CBISarrOutput(b, a)
    '        Next
    '    Next
    
    If IsEmpty(CBA_COM_CBISarrOutput(1, 1)) Then CBA_COM_genCBISoutput = False Else CBA_COM_genCBISoutput = True
    
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_COM_genCBISoutput", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function
Function CBA_COM_decodeCBISdescription(ByVal useInputArray As Boolean, ByRef Inputarr() As Variant, Optional ByVal PCode As Long, Optional ByVal strDescription As String) As Boolean
    'This function is used to decode the CBIS Product Description to create public array 'CBA_COM_CBISarrOutput' with any package data stripped out
    Dim bOutput As Boolean, a As Long, b As Long, c As Long, lCRow As Long, lLimit As Long, oItem
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    'Dim lLimit As Long
    
    Set CBA_COM_colInput = New Collection
    Set CBA_COM_potGram = New Collection
    Set CBA_COM_potLitres = New Collection
    Set CBA_COM_potPieces = New Collection
    Set CBA_COM_potSheet = New Collection
    Set CBA_COM_potMetres = New Collection
    Set CBA_COM_potOther = New Collection
    Set CBA_COM_colMulti = New Collection
    Set CBA_COM_colAdddetail = New Collection
    Erase CBA_COM_CBISarrOutput
    'lLimit = 12805
    CBA_COM_numOutput = 0
    'duplicates = 0
    'didntDecode = 0
    
    
    
    If useInputArray = False Then
        If PCode = 0 Or strDescription = "" Then
            'No Comp can be made as string and code not both provided
            Exit Function
        Else
            ReDim Preserve CBA_COM_CBISarrOutput(1 To 59, 1)
            CBA_COM_colInput.Add "((" & PCode & ")) " & strDescription
            CBA_COM_CBISarrOutput(1, 1) = PCode
            CBA_COM_CBISarrOutput(2, 1) = strDescription
        End If
    Else
        If IsArray(Inputarr) = False Then
            'No Comp can be made as array did not exist
            Exit Function
        ElseIf CBA_COM_NumberOfArrayDimensions(arr:=Inputarr()) <> 2 Then
            'No Comp can be made as Array is not 2 dimentional
            Exit Function
        Else
            ReDim Preserve CBA_COM_CBISarrOutput(1 To 59, 1 To UBound(Inputarr, 2))
            For a = 1 To UBound(Inputarr, 2)
                If IsNumeric(Inputarr(1, a)) Then
                    CBA_COM_colInput.Add "((" & Inputarr(1, a) & ")) " & Inputarr(2, a)
                    CBA_COM_CBISarrOutput(1, a) = Inputarr(1, a)
                    CBA_COM_CBISarrOutput(2, a) = Inputarr(2, a)
                Else
                'One of the product codes provided contains letters of charachters
                Exit Function
                End If
            Next
        End If
    End If

    CBA_COM_numOutput = UBound(CBA_COM_CBISarrOutput, 2)
    lLimit = UBound(CBA_COM_CBISarrOutput, 2) + 1
    
    b = 0
    For Each oItem In CBA_COM_colInput
        
    '    If Mid(oItem, 3, 5) = "56059" Then
    '        b = b
    '    End If
        
        
        b = b + 1
        For a = 1 To Len(oItem)
            If (Mid(UCase(oItem), a, 2) = "SC" And IsNumeric(Mid(oItem, a + 2, 1))) Or (Mid(UCase(oItem), a, 2) = "SC" And IsNumeric(Mid(oItem, a + 3, 1))) Then
                CBA_COM_potGram.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
            If (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 1, 2) = "KG") Or (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 2, 2) = "KG") Then
                CBA_COM_potGram.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
            If InStr(1, UCase(oItem), "PER KG") > 0 Then
                CBA_COM_potGram.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
            If (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 1, 2) = "ML") Or (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 2, 2) = "ML") Then
                CBA_COM_potLitres.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
            If (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 1, 3) = "LNG") Or (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 2, 3) = "LNG") Then
                CBA_COM_potOther.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
    
            If (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 1, 2) = "PC") Or (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 2, 2) = "PC") Then
                CBA_COM_potPieces.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
            If (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 1, 2) = "PK") Or (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 2, 2) = "PK") Then
                CBA_COM_potPieces.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
            If (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 1, 2) = "CM") Or (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 2, 2) = "CM") Then
                CBA_COM_potMetres.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
            If (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 1, 2) = "MM") Or (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 2, 2) = "MM") Then
                CBA_COM_potMetres.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
            If (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 1, 1) = "L") Or (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 2, 1) = "L" And Not Mid(UCase(oItem), a + 1, 3) = "PLY") Then
                CBA_COM_potLitres.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
            If (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 1, 1) = "M") Or (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 2, 1) = "M") Then
                CBA_COM_potMetres.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
            If (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 1, 1) = "G") Or (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 2, 1) = "G") Then
                CBA_COM_potGram.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
            If (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 1, 1) = "P" And Not Mid(UCase(oItem), a + 1, 3) = "PLY") Or (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 2, 1) = "P" And Not Mid(UCase(oItem), a + 2, 3) = "PLY") Then
                CBA_COM_potPieces.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
            If (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 1, 1) = "X") Or (IsNumeric(Mid(oItem, a + 1, 1)) And Mid(UCase(oItem), a, 1) = "X") _
                Or (IsNumeric(Mid(oItem, a + 2, 1)) And Mid(UCase(oItem), a, 1) = "X") Or (IsNumeric(Mid(oItem, a, 1)) And Mid(UCase(oItem), a + 2, 1) = "X") Then
                CBA_COM_colMulti.Add oItem
                CBA_COM_colInput.Remove (b)
                b = b - 1
                Exit For
            End If
        Next
    Next
    
    
        c = 1
        For Each oItem In CBA_COM_colInput
            lCRow = c
            Do Until lCRow = CBA_COM_numOutput + 1
                If "((" & CBA_COM_CBISarrOutput(1, lCRow) & ")) " & CBA_COM_CBISarrOutput(2, lCRow) = oItem Then
                    CBA_COM_CBISarrOutput(3, lCRow) = "1"
                    CBA_COM_CBISarrOutput(4, lCRow) = "Pieces"
                    c = lCRow
                    'Application.StatusBar = "Row:" & c & " of " & CBA_COM_numOutput
                    Exit Do
                End If
            lCRow = lCRow + 1
            If lCRow > CBA_COM_numOutput Then Exit Do
            Loop
        Next
        'Application.StatusBar = False
    
    
    bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_SortColl(CBA_COM_potGram)
    bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_SortColl(CBA_COM_potLitres)
    bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_SortColl(CBA_COM_potMetres)
    bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_SortColl(CBA_COM_potPieces)
    bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_SortColl(CBA_COM_potSheet)
    bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_SortOther(CBA_COM_potOther)
    'bOutput = CBA_COM_CBIS_PackDecodeDecodeAddData
    bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_DecodeXData
    bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_FinaliseOutput
                 
    If UBound(CBA_COM_CBISarrOutput, 2) > 0 Then CBA_COM_decodeCBISdescription = True Else CBA_COM_decodeCBISdescription = False
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_COM_decodeCBISdescription", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Function CBA_COM_SortColl(ByVal coll As Collection)
    Dim bfound As Boolean, bOutput As Boolean, b As Long, l As Long, g As Long, oItem, askhg, lkhlkjh, testerlkj
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    b = 0
    For Each oItem In coll
        b = b + 1
        bfound = False
        
    '    If InStr(1, UCase(oItem), "56059") Then
    '       a = a
    '    End If
        
        For l = 1 To Len(oItem)
            If Mid(UCase(oItem), l, 3) = "SC1" Or Mid(UCase(oItem), l, 3) = "SC2" Then
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, "SC1")
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 2) = "CM") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = " " And Mid(UCase(oItem), l + 2, 2) = "CM") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 2) = "NM") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = " " And Mid(UCase(oItem), l + 2, 2) = "NM") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 2) = "MM") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = " " And Mid(UCase(oItem), l + 2, 2) = "MM") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 2) = "PK" And Not Mid(UCase(oItem), l + 1, 3) = "PLY") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = " " And Mid(UCase(oItem), l + 2, 2) = "PK" And Not Mid(UCase(oItem), l + 2, 3) = "PLY") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 2) = "PC") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = " " And Mid(UCase(oItem), l + 2, 2) = "PC") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 2) = "PG") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    If g = 0 Then Exit Do
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = " " And Mid(UCase(oItem), l + 2, 2) = "PG") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 2) = "KG") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = " " And Mid(UCase(oItem), l + 2, 2) = "KG") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 2) = "ML") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = " " And Mid(UCase(oItem), l + 2, 2) = "ML") Then
                g = l
                Do
                    askhg = Mid(oItem, g, 1)
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = "L") And InStr(1, UCase(oItem), "LASH") = 0 Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 1))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = " " And Mid(UCase(oItem), l + 2, 1) = "L") And InStr(1, UCase(oItem), "LASH") = 0 Then
                g = l
                Do
                    askhg = Mid(oItem, g, 1)
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 1))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = "G") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 1))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = " " And Mid(UCase(oItem), l + 2, 1) = "G") Then
                g = l
                Do
                    lkhlkjh = Mid(oItem, g, 1)
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                testerlkj = Mid(oItem, g + 1, l - g + 1)
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 2))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = "M") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 1))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = " " And Mid(UCase(oItem), l + 2, 1) = "M") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 1))
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = "P" And Not Mid(UCase(oItem), l + 1, 3) = "PLY" And Not Mid(UCase(oItem), l + 2, 1) = "/") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 1))
                bfound = True
            ElseIf (IsNumeric(Mid(oItem, l, 1)) And Mid(UCase(oItem), l + 1, 1) = " " And Mid(UCase(oItem), l + 2, 1) = "P" And Not Mid(UCase(oItem), l + 2, 3) = "PLY") Then
                g = l
                Do
                    If IsNumeric(Mid(oItem, g, 1)) Or Mid(oItem, g, 1) = "." Then
                    g = g - 1
                    Else
                    Exit Do
                    End If
                Loop
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, g + 1, l - g + 1))
                bfound = True
                bfound = True
            ElseIf InStr(1, UCase(oItem), "PER KG") > 0 Then
                bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, Mid(oItem, InStr(1, oItem, ")) ") + 3, 999))
                bfound = True
                Exit For
            End If
        Next
        
        If bfound = True Then
            coll.Remove (b)
            b = b - 1
        Else
            CBA_COM_potOther.Add oItem
            coll.Remove (b)
            b = b - 1
        End If
    Next
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_COM_SortColl", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Function CBA_COM_CreateOutputArray(ByVal ValLng As String, ByVal vVal As String)
    Dim lOutputRow As Long, lStartCol As Long, a As Long, lDidntDecode As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
        
    lOutputRow = 0
    For a = 1 To CBA_COM_numOutput
        If "((" & CBA_COM_CBISarrOutput(1, a) & ")) " & CBA_COM_CBISarrOutput(2, a) = ValLng Then
            lOutputRow = a
            Exit For
        End If
    Next
    
    For a = 3 To 29
    On Error GoTo GTResult
    If IsEmpty(CBA_COM_CBISarrOutput(a, lOutputRow)) Then
GTResult:
        Err.Clear
        On Error GoTo Err_Routine
        lStartCol = a
        Exit For
    End If
    Next
    
    If lOutputRow = 0 Then
        lDidntDecode = lDidntDecode + 1
    Else
    
        If InStr(1, UCase(vVal), "X") Then
            CBA_COM_colMulti.Add vVal
        ElseIf UCase(vVal) = "SC1" Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = 1
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
            CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1000
            CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "g"
        ElseIf UCase(vVal) = "NULL" Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = 1
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "SHEET") Then
            If Mid(vVal, 1, InStr(1, UCase(vVal), "SHEET")) = 1 Then CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = 1 Else CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal)
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Sheets"
        ElseIf InStr(1, UCase(vVal), "PACK") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal)
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "PIECE") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal)
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "PER KG") > 0 Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = 1000
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "g"
        ElseIf InStr(1, UCase(vVal), "PCE") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal)
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "PG") Then
            CBA_COM_potPieces.Add vVal
        ElseIf InStr(1, UCase(vVal), "KG") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal) * 1000
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "g"
            CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
            CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "MM") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal)
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "mm"
            CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
            CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "CM") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal) * 10
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "mm"
            CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
            CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "NM") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = 1
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "SS") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal)
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Sheets"
            CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
            CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "ML") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal)
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "ml"
            CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
            CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "LT") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal) * 1000
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "ml"
            CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
            CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "EA") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal)
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "PK") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal)
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "PC") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal)
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "G") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal)
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "g"
            CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
            CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "M") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal) * 1000
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "mm"
            CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
            CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "L") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal) * 1000
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "ml"
            CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
            CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
        ElseIf InStr(1, UCase(vVal), "P") Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(vVal)
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
        ElseIf IsNumeric(vVal) Then
            CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = vVal
            CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
        End If
    
    End If
    'bOutput = CBA_COM_PackDecode.DecodeAddData
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_COM_CreateOutputArray", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Function CBA_COM_SortOther(ByVal coll As Collection)
    Dim bfound As Boolean, bOutput As Boolean, b As Long, oItem
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    b = 0
    For Each oItem In coll
        b = b + 1
        bfound = False
        bOutput = CBA_COM_CBIS_PackDecode.CBA_COM_CreateOutputArray(oItem, "1PK")
        coll.Remove (b)
        b = b - 1
    Next
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_COM_SortOther", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function CBA_COM_DecodeXData()
    Dim numSpace, place, numAddDetail, lOutputRow, lStartCol As Long, oItem, a As Long, b As Long, F As Long, Y As Long, Loca, valtowork
    Dim Linked
    Dim LinkedVal As Single
    Dim lDidntDecode As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    numAddDetail = 0
    For Each oItem In CBA_COM_colMulti
        numAddDetail = numAddDetail + 1
        Set CBA_COM_colWhere = New Collection
        ReDim Preserve CBA_COM_arrSortDetail(1 To 29, 1 To numAddDetail)
        CBA_COM_arrSortDetail(1, numAddDetail) = oItem
        numSpace = 0
    
    '                If Mid(oItem, 3, 5) = "1078" Then
    '                    a = a
    '                End If
        
        If InStr(1, oItem, "+") Or InStr(1, oItem, "-") Then
            For a = 1 To Len(oItem)
                If Mid(oItem, a, 1) = "+" Or Mid(oItem, a, 1) = "-" Then
                    oItem = Mid(oItem, 1, a - 1) & Mid(oItem, a + 1, Len(oItem) - a)
                End If
            Next
        End If
        
        
        
        For a = 1 To Len(oItem)
                'kajsgd = Mid(oItem, a, 1)
            If Mid(oItem, a, 1) = " " Then
                numSpace = numSpace + 1
                CBA_COM_colWhere.Add a
            ElseIf Mid(UCase(oItem), a, 1) = "X" Then
                If a > 1 Then
                    If Mid(UCase(oItem), a - 1, 1) = " " Then
                    ElseIf IsNumeric(Mid(UCase(oItem), a - 1, 1)) Then
                        numSpace = numSpace + 1
                        CBA_COM_colWhere.Add a
                    End If
                Else
                    numSpace = numSpace + 1
                    CBA_COM_colWhere.Add a
                End If
            End If
            
            
        Next
        If numSpace > 0 Then
            a = 1
            place = 1
            For Each Loca In CBA_COM_colWhere
                If place = 1 Then
                    a = a + 1
                    CBA_COM_arrSortDetail(a, numAddDetail) = Trim(Mid(oItem, place, Loca - 1))
                    place = Loca
                Else
                    a = a + 1
                    CBA_COM_arrSortDetail(a, numAddDetail) = Trim(Mid(oItem, place, Loca - place))
                    place = Loca
                End If
            Next
            a = a + 1
            CBA_COM_arrSortDetail(a, numAddDetail) = Trim(Mid(oItem, place, Len(oItem) - place + 1))
        End If
    
    Next
    
    'ADD LOGIC THAT SORTS ALL PARTS
    For a = 1 To numAddDetail
        Linked = False
       
        lOutputRow = 0
        For Y = 1 To CBA_COM_numOutput
            If "((" & CBA_COM_CBISarrOutput(1, Y) & ")) " & CBA_COM_CBISarrOutput(2, Y) = CBA_COM_arrSortDetail(1, a) Then
                lOutputRow = Y
                Exit For
            End If
        Next
    
        If lOutputRow = 0 Then
            lDidntDecode = lDidntDecode + 1
        Else
            For Y = 3 To 29
                On Error GoTo GTResult
                If IsEmpty(CBA_COM_CBISarrOutput(Y, lOutputRow)) Then
GTResult:
                    On Error GoTo Err_Routine
                    lStartCol = Y
                    Exit For
                End If
            Next
            On Error GoTo Err_Routine
            Linked = False
        For b = 2 To 29
            
            If IsEmpty(CBA_COM_arrSortDetail(b, a)) Or CBA_COM_arrSortDetail(b, a) = "" Then
            '*******CONDITION HANDLE RQD*******
            Else
                
                If Linked = True Then

                    valtowork = Mid(CBA_COM_arrSortDetail(b, a), 2, Len(CBA_COM_arrSortDetail(b, a)) - 1)
                                If InStr(1, UCase(valtowork), "SHEET") Then
                                    If Mid(valtowork, 1, InStr(1, UCase(valtowork), "SHEET")) = 1 Then CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = 1 Else CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(valtowork)
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Sheets"
                                ElseIf InStr(1, UCase(valtowork), "PACK") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork)
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "PIECE") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork)
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "PCE") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork)
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "PG") Then
                                    CBA_COM_potPieces.Add valtowork
                                ElseIf InStr(1, UCase(valtowork), "KG") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork) * 1000
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "g"
                                    CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = LinkedVal
                                    CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "MM") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork)
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "mm"
                                    CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = LinkedVal
                                    CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "CM") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork) * 10
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "mm"
                                    CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = LinkedVal
                                    CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "NM") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "SS") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork)
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Sheets"
                                    CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = LinkedVal
                                    CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "ML") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork)
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "ml"
                                    CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = LinkedVal
                                    CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "LT") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork) * 1000
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "ml"
                                    CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = LinkedVal
                                    CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "EA") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork)
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "PK") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork)
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "G") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork)
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "g"
                                    CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = LinkedVal
                                    CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "M") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork) * 1000
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "mm"
                                    CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = LinkedVal
                                    CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                ElseIf InStr(1, UCase(valtowork), "L") Then
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = LinkedVal * CBA_COM_Nois(valtowork) * 1000
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "ml"
                                    CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = LinkedVal
                                    CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                ElseIf IsNumeric(valtowork) And Mid(UCase(CBA_COM_arrSortDetail(b + 1, a)), 1, 1) = "X" Then
                                    Linked = True
                                    LinkedVal = valtowork
                '                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_arrSortDetail(b + 1, a)
                '                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                                End If
                     Linked = False
                
                ElseIf CBA_COM_IsChar(CBA_COM_arrSortDetail(b, a)) Then
                'Ignore if only charachters
                ElseIf Mid(CBA_COM_arrSortDetail(b, a), 1, 2) = "((" Then
                'Ignore the product code
                ElseIf CBA_COM_IsLetter(CBA_COM_arrSortDetail(b, a)) Then
                'Ignore if just a word
                ElseIf UCase(CBA_COM_arrSortDetail(b, a)) = "SC1" Then
                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = 1
                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                    CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1000
                    CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "g"
                ElseIf UCase(CBA_COM_arrSortDetail(b, a)) = "NULL" Then
                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = 1
                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "SHEET") Then
                    If Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "SHEET")) = 1 Then CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = 1 Else CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a))
                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Sheets"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PACK") Then
                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a))
                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PIECE") Then
                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a))
                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PCE") Then
                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a))
                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PG") Then
                    CBA_COM_potPieces.Add CBA_COM_arrSortDetail(b, a)
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "KG") Then
                        For F = 1 To Len(CBA_COM_arrSortDetail(b, a))
                            If IsNumeric(Mid(CBA_COM_arrSortDetail(b, a), F, 1)) And (Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 1, 2) = "KG" Or Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 2, 2) = "KG") Then
                                On Error GoTo GTNext
                                CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a)) * 1000
                                CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "g"
                                CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
                                CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                On Error GoTo Err_Routine
                                Exit For
                            End If
                        Next
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "MM") Then
                        For F = 1 To Len(CBA_COM_arrSortDetail(b, a))
                            If IsNumeric(Mid(CBA_COM_arrSortDetail(b, a), F, 1)) And (Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 1, 2) = "MM" Or Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 2, 2) = "MM") Then
                                On Error GoTo GTNext
                                CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a))
                                CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "mm"
                                CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
                                CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                On Error GoTo Err_Routine
                                Exit For
                            End If
                        Next
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CM") Then
                        For F = 1 To Len(CBA_COM_arrSortDetail(b, a))
                            If IsNumeric(Mid(CBA_COM_arrSortDetail(b, a), F, 1)) And (Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 1, 2) = "CM" Or Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 2, 2) = "CM") Then
                                On Error GoTo GTNext
                                CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a)) * 10
                                CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "mm"
                                CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
                                CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                On Error GoTo Err_Routine
                                Exit For
                            End If
                        Next
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "NM") Then
                        For F = 1 To Len(CBA_COM_arrSortDetail(b, a))
                            If IsNumeric(Mid(CBA_COM_arrSortDetail(b, a), F, 1)) And (Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 1, 2) = "NM" Or Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 2, 2) = "NM") Then
                                On Error GoTo GTNext
                                CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = 1
                                CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                                On Error GoTo Err_Routine
                                Exit For
                            End If
                        Next
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "SS") Then
                        For F = 1 To Len(CBA_COM_arrSortDetail(b, a))
                            If IsNumeric(Mid(CBA_COM_arrSortDetail(b, a), F, 1)) And (Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 1, 2) = "SS" Or Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 2, 2) = "SS") Then
                                On Error GoTo GTNext
                                CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a))
                                CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Sheets"
                                CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
                                CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                On Error GoTo Err_Routine
                                Exit For
                            End If
                        Next
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "ML") Then
                        For F = 1 To Len(CBA_COM_arrSortDetail(b, a))
                            If IsNumeric(Mid(CBA_COM_arrSortDetail(b, a), F, 1)) And (Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 1, 2) = "ML" Or Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 2, 2) = "ML") Then
                                On Error GoTo GTNext
                                CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a))
                                CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "ml"
                                CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
                                CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                On Error GoTo Err_Routine
                                Exit For
                            End If
                        Next
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "LT") Then
                        For F = 1 To Len(CBA_COM_arrSortDetail(b, a))
                            If IsNumeric(Mid(CBA_COM_arrSortDetail(b, a), F, 1)) And (Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 1, 2) = "LT" Or Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 2, 2) = "LT") Then
                                On Error GoTo GTNext
                                CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a)) * 1000
                                CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "ml"
                                CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
                                CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                On Error GoTo Err_Routine
                                Exit For
                            End If
                        Next
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "EA") Then
                        For F = 1 To Len(CBA_COM_arrSortDetail(b, a))
                            If IsNumeric(Mid(CBA_COM_arrSortDetail(b, a), F, 1)) And (Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 1, 2) = "EA" Or Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 2, 2) = "EA") Then
                                On Error GoTo GTNext
                                CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a))
                                CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                                On Error GoTo Err_Routine
                                Exit For
                            End If
                        Next
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PK") Then
                        For F = 1 To Len(CBA_COM_arrSortDetail(b, a))
                            If IsNumeric(Mid(CBA_COM_arrSortDetail(b, a), F, 1)) And (Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 1, 2) = "PK" Or Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 2, 2) = "PK") Then
                                On Error GoTo GTNext
                                CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a))
                                CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                                On Error GoTo Err_Routine
                                Exit For
                            End If
                        Next
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "G") Then
                        For F = 1 To Len(CBA_COM_arrSortDetail(b, a))
                            If IsNumeric(Mid(CBA_COM_arrSortDetail(b, a), F, 1)) And (Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 1, 1) = "G" Or Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 2, 1) = "G") Then
                                On Error GoTo GTNext
                                CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a))
                                CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "g"
                                CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
                                CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                On Error GoTo Err_Routine
                                Exit For
                            End If
                        Next
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "M") Then
                        For F = 1 To Len(CBA_COM_arrSortDetail(b, a))
                            If IsNumeric(Mid(CBA_COM_arrSortDetail(b, a), F, 1)) And (Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 1, 1) = "M" Or Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 2, 1) = "M") Then
                                On Error GoTo GTNext
                                If InStr(1, CBA_COM_arrSortDetail(b, a), "CD") Then
                                CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = 1
                                CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                                Else
                                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a)) * 1000
                                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "mm"
                                    CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
                                    CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                End If
                                On Error GoTo Err_Routine
                                Exit For
                            End If
                        Next
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "L") Then
                        For F = 1 To Len(CBA_COM_arrSortDetail(b, a))
                            If IsNumeric(Mid(CBA_COM_arrSortDetail(b, a), F, 1)) And (Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 1, 1) = "L" Or Mid(UCase(CBA_COM_arrSortDetail(b, a)), F + 2, 1) = "L") Then
                                On Error GoTo GTNext
                                CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_Nois(CBA_COM_arrSortDetail(b, a)) * 1000
                                CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "ml"
                                CBA_COM_CBISarrOutput(lStartCol + 2, lOutputRow) = 1
                                CBA_COM_CBISarrOutput(lStartCol + 3, lOutputRow) = "Pieces"
                                On Error GoTo Err_Routine
                                Exit For
                            End If
                        Next
                ElseIf IsNumeric(CBA_COM_arrSortDetail(b, a)) And Mid(UCase(CBA_COM_arrSortDetail(b + 1, a)), 1, 1) = "X" Then
                    Linked = True
                    LinkedVal = CBA_COM_arrSortDetail(b, a)
'                    CBA_COM_CBISarrOutput(lStartCol, lOutputRow) = CBA_COM_arrSortDetail(b + 1, a)
'                    CBA_COM_CBISarrOutput(lStartCol + 1, lOutputRow) = "Pieces"
                End If
            End If
GTNext:
    Err.Clear
    On Error GoTo Err_Routine
        Next
    
        End If
    Next
    
    Erase CBA_COM_arrSortDetail
    Set CBA_COM_colMulti = New Collection
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_COM_DecodeXData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Function CBA_COM_FinaliseOutput()
    Dim transval1, transval2 As Variant, a As Long, b As Long, num, d As Long, bfound As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    For a = 1 To UBound(CBA_COM_CBISarrOutput, 2)

        
        bfound = False
        For d = 3 To 29
            If CBA_COM_CBISarrOutput(d, a) = "Pieces" Then
                bfound = True
                Exit For
            End If
                If IsEmpty(CBA_COM_CBISarrOutput(d, a)) Then Exit For
        Next
        If bfound = False Then
            For d = 3 To 29
                If IsEmpty(CBA_COM_CBISarrOutput(d, a)) Then
                    CBA_COM_CBISarrOutput(d, a) = 1
                    CBA_COM_CBISarrOutput(d + 1, a) = "Pieces"
                    Exit For
                End If
            Next
        End If
    Next
    
    For a = 1 To UBound(CBA_COM_CBISarrOutput, 2)

        b = 0
        bfound = False
        For d = 3 To 29
            If CBA_COM_CBISarrOutput(d, a) = "Pieces" Then
                num = CBA_COM_CBISarrOutput(d - 1, a)
                b = d
                bfound = True
                Exit For
            End If
            If IsEmpty(CBA_COM_CBISarrOutput(d, a)) Then Exit For
        Next
        If bfound = True Then
            For d = b + 1 To 29
                If CBA_COM_CBISarrOutput(d, a) = "Pieces" Then
                    If CBA_COM_CBISarrOutput(d - 1, a) > num Then
                        CBA_COM_CBISarrOutput(b - 1, a) = CBA_COM_CBISarrOutput(d - 1, a)
                        num = CBA_COM_CBISarrOutput(d - 1, a)
                        CBA_COM_CBISarrOutput(d - 1, a) = Empty
                        CBA_COM_CBISarrOutput(d, a) = Empty
                    Else
                        CBA_COM_CBISarrOutput(d - 1, a) = Empty
                        CBA_COM_CBISarrOutput(d, a) = Empty
                    End If
                End If
                If IsEmpty(CBA_COM_CBISarrOutput(d, a)) Then Exit For
            Next
        End If
    Next
    For a = 1 To UBound(CBA_COM_CBISarrOutput, 2)
'    If CBA_COM_CBISarrOutput(1, a) = 76016 Then
'    a = a
'    End If
        'asdas = CBA_COM_CBISarrOutput(1, a)
        For b = 3 To 29
            If b = 3 Then
                If CBA_COM_CBISarrOutput(b + 1, a) = "Pieces" Then
                    Exit For
                Else
                    transval1 = CBA_COM_CBISarrOutput(b, a)
                    transval2 = CBA_COM_CBISarrOutput(b + 1, a)
                End If
            ElseIf b > 4 And CBA_COM_CBISarrOutput(b + 1, a) = "Pieces" Then
                CBA_COM_CBISarrOutput(3, a) = CBA_COM_CBISarrOutput(b, a)
                CBA_COM_CBISarrOutput(4, a) = CBA_COM_CBISarrOutput(b + 1, a)
                CBA_COM_CBISarrOutput(b, a) = transval1
                CBA_COM_CBISarrOutput(b + 1, a) = transval2
                Exit For
            End If
        Next
    Next
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_COM_FinaliseOutput", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function
Function CBA_COM_IsLetter(ByVal strValue As String) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 65 To 90, 97 To 122
                CBA_COM_IsLetter = True
            Case Else
                CBA_COM_IsLetter = False
                Exit For
        End Select
    Next
End Function
Function CBA_COM_IsChar(ByVal strValue As String) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 0 To 47, 58 To 64, 91 To 96, 123 To 127
                CBA_COM_IsChar = True
            Case Else
                CBA_COM_IsChar = False
                Exit For
        End Select
    Next
End Function

Public Function CBA_COM_NumberOfArrayDimensions(arr As Variant) As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumberOfArrayDimensions
' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
Dim Res As Long
On Error Resume Next
' Loop, increasing the dimension index Ndx, until an error occurs.
' An error will occur when Ndx exceeds the number of dimension
' in the array. Return Ndx - 1.
Do
    Ndx = Ndx + 1
    Res = UBound(arr, Ndx)
Loop Until Err.Number <> 0

CBA_COM_NumberOfArrayDimensions = Ndx - 1

End Function
Function CBA_COM_Nois(ByVal vVal As String) As Single
    Dim startpoint, endpoint As Long, a As Long
    startpoint = 0
    
    For a = 1 To Len(Trim(vVal))
        If IsNumeric(Mid(Trim(vVal), a, 1)) Then
            startpoint = a
            Exit For
        End If
    Next
    
    If startpoint > 0 Then
        For a = startpoint To Len(Trim(vVal))
            If IsNumeric(Mid(Trim(vVal), a, 1)) Or Mid(Trim(vVal), a, 1) = "." Then
                endpoint = a
            Else
            Exit For
            End If
        Next
    
    CBA_COM_Nois = CSng(Mid(Trim(vVal), startpoint, endpoint - (startpoint - 1)))
    Else
        CBA_COM_Nois = 0
    End If

End Function
