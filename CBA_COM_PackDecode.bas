Attribute VB_Name = "CBA_COM_PackDecode"
'''''*********CBA_COM_PackDecode V0.01*******************'''''''
'''''''*********DATE ADJUSTED 17-01-18*******************'''''''
Option Explicit
Option Private Module          ' Excel users cannot access procedures

Function DecodePack(ByVal useArrayorCollection As Boolean, ByRef Inputarr() As Variant, Optional ByVal CBA_COM_colInput As Collection) As Boolean
    Dim bOutput As Boolean, bfound As Boolean, oItem, strAldiMess
    Dim lNum As Long, a As Long, v As Long, d As Long, c As Long, sStateLook, assPRet, j As Long, k As Long, Cancel
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    Set CBA_COM_potGram = New Collection
    Set CBA_COM_potLitres = New Collection
    Set CBA_COM_potMetres = New Collection
    Set CBA_COM_leftovers = New Collection
    Set CBA_COM_potPieces = New Collection
    Set CBA_COM_potPair = New Collection
    Set CBA_COM_potSheet = New Collection
    Set CBA_COM_colAdddetail = New Collection
    Set CBA_COM_colMulti = New Collection
    Set CBA_COM_colNotDecoded = New Collection
    Erase CBA_COM_arrOutput
    Erase CBA_COM_PackarrOutput
    Erase CBA_COM_arrSortDetail
    CBA_COM_numOutput = 0

    If useArrayorCollection = True Then
        Set CBA_COM_colInput = New Collection
        For a = 0 To UBound(Inputarr, 2)
                bfound = False
                For Each oItem In CBA_COM_colInput
                    If oItem = Inputarr(1, a) Then
                        bfound = True
                        Exit For
                    End If
                Next
                If bfound = False Then
                    CBA_COM_colInput.Add Inputarr(1, a)
                End If
        Next
    End If
    
        For Each oItem In CBA_COM_colInput
            DecodePack = True
            If InStr(1, LCase(oItem), "this product is not available") > 0 Or InStr(1, LCase(oItem), "due to its short shelf life") > 0 Then
            CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, LCase(oItem), "per can") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, LCase(oItem), "per cask") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, LCase(oItem), "per jar") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, LCase(oItem), "per case") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, LCase(oItem), "in any") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, UCase(oItem), "G") Then CBA_COM_potGram.Add oItem
            ElseIf InStr(1, UCase(oItem), "PER KILO") Then CBA_COM_potGram.Add oItem
            ElseIf InStr(1, UCase(oItem), "BLADES") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, UCase(oItem), "CARTON") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, UCase(oItem), "BUNCH") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, UCase(oItem), "PIECE") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, UCase(oItem), "SHEET") Then CBA_COM_potSheet.Add oItem
            ElseIf InStr(1, UCase(oItem), "PAIR") Then CBA_COM_potPair.Add oItem
            ElseIf InStr(1, UCase(oItem), "PACK") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, UCase(oItem), "PCE") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, UCase(oItem), "PK") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, UCase(oItem), "PC") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, UCase(oItem), "EA") Then CBA_COM_potPieces.Add oItem
            ElseIf InStr(1, UCase(oItem), "SS") Then CBA_COM_potSheet.Add oItem
            ElseIf InStr(1, UCase(oItem), "LL") Then CBA_COM_potLitres.Add oItem
            ElseIf InStr(1, UCase(oItem), "L") Then CBA_COM_potLitres.Add oItem
            ElseIf InStr(1, UCase(oItem), "M") Then CBA_COM_potMetres.Add oItem
            ElseIf IsNumeric(oItem) Then CBA_COM_potPieces.Add oItem
            Else
            CBA_COM_leftovers.Add oItem
            End If
        Next
    
    
    bOutput = CBA_COM_PackDecode.SortCollection(CBA_COM_potGram)
    bOutput = CBA_COM_PackDecode.SortCollection(CBA_COM_potLitres)
    bOutput = CBA_COM_PackDecode.SortCollection(CBA_COM_potMetres)
    bOutput = CBA_COM_PackDecode.SortCollection(CBA_COM_potPair)
    bOutput = CBA_COM_PackDecode.SortCollection(CBA_COM_potPieces)
    bOutput = CBA_COM_PackDecode.SortCollection(CBA_COM_potSheet)
    bOutput = CBA_COM_PackDecode.SortCollection(CBA_COM_leftovers)
    bOutput = CBA_COM_PackDecode.DecodeXData
    
    If useArrayorCollection = True Then
        ReDim CBA_COM_arrOutput(1 To 6, 1 To UBound(Inputarr, 2) + 1)
        lNum = 0
        For a = 0 To UBound(Inputarr, 2)
            'add line for CW packsize
            For v = 1 To UBound(CBA_COM_PackarrOutput, 2)
                If Inputarr(1, a) = CBA_COM_PackarrOutput(1, v) Then
                    d = 2
                    For c = 2 To UBound(CBA_COM_PackarrOutput, 1)
                        If IsEmpty(CBA_COM_PackarrOutput(c, v)) Then Exit For
                        If (sStateLook = "National" And Inputarr(3, a) = CSng(assPRet)) Or sStateLook = Inputarr(12, a) Or sStateLook = "FindingCheep" Then
                        If c >= d Then
                            lNum = lNum + 1
                            ReDim Preserve CBA_COM_arrOutput(1 To 6, 1 To lNum)
                            CBA_COM_arrOutput(1, lNum) = Inputarr(0, a)
                            If sStateLook = "National" Then CBA_COM_arrOutput(2, lNum) = sStateLook Else CBA_COM_arrOutput(2, lNum) = Inputarr(12, a)
                            CBA_COM_arrOutput(3, lNum) = Inputarr(3, a)
                            CBA_COM_arrOutput(4, lNum) = CBA_COM_PackarrOutput(c, v)
                            CBA_COM_arrOutput(5, lNum) = CBA_COM_PackarrOutput(c + 1, v)
                            d = c + 2
                        End If
                        End If
                    Next
                Exit For
                End If
            Next
            'add lines for cupinfo CW
            If Inputarr(6, a) = "" Or Inputarr(6, a) = Null Or UCase(Inputarr(6, a)) = "KGM" Then
            Else
                lNum = lNum + 1
                ReDim Preserve CBA_COM_arrOutput(1 To 6, 1 To lNum)
                CBA_COM_arrOutput(1, lNum) = Inputarr(0, a)
                If sStateLook = "National" Then CBA_COM_arrOutput(2, lNum) = sStateLook Else CBA_COM_arrOutput(2, lNum) = Inputarr(12, a)
                CBA_COM_arrOutput(3, lNum) = Inputarr(8, a)
                If UCase(Inputarr(6, a)) = "KG" Then
                    CBA_COM_arrOutput(4, lNum) = Inputarr(7, a) * 1000
                    CBA_COM_arrOutput(5, lNum) = "g"
                ElseIf UCase(Inputarr(6, a)) = "L" Then
                    CBA_COM_arrOutput(4, lNum) = Inputarr(7, a) * 1000
                    CBA_COM_arrOutput(5, lNum) = "ml"
                ElseIf UCase(Inputarr(6, a)) = "M" Then
                    CBA_COM_arrOutput(4, lNum) = Inputarr(7, a) * 1000
                    CBA_COM_arrOutput(5, lNum) = "mm"
                ElseIf UCase(Inputarr(6, a)) = "SS" Then
                    CBA_COM_arrOutput(4, lNum) = Inputarr(7, a)
                    CBA_COM_arrOutput(5, lNum) = "Sheets"
                ElseIf UCase(Inputarr(6, a)) = "EA" Then
                'CODE TO ACCOMODATE IF EACH NEEDS TO BE SHEETS OR
                    If InStr(1, UCase(Inputarr(13, a)), "1 PLY") > 0 Or InStr(1, UCase(Inputarr(13, a)), "2 PLY") > 0 Or InStr(1, UCase(Inputarr(13, a)), "3 PLY") > 0 Or InStr(1, UCase(Inputarr(13, a)), "4 PLY") Or InStr(1, UCase(Inputarr(13, a)), "1PLY") > 0 Or InStr(1, UCase(Inputarr(13, a)), "2PLY") > 0 Or InStr(1, UCase(Inputarr(13, a)), "3PLY") > 0 Or InStr(1, UCase(Inputarr(13, a)), "4PLY") > 0 Then
                        CBA_COM_arrOutput(4, lNum) = Inputarr(7, a)
                        CBA_COM_arrOutput(5, lNum) = "Sheets"
                    ElseIf InStr(1, LCase(Inputarr(13, a)), "paper towel") > 0 Or InStr(1, LCase(Inputarr(13, a)), "print serviettes") > 0 Then
                        CBA_COM_arrOutput(4, lNum) = Inputarr(7, a)
                        CBA_COM_arrOutput(5, lNum) = "Sheets"
                    ElseIf InStr(1, LCase(Inputarr(13, a)), "blades") > 0 Then
                        For j = 1 To Len(LCase(Inputarr(13, a)))
                            If Mid(LCase(Inputarr(13, a)), j, 6) = "blades" Then
                                For k = j - 1 To 1 Step -1
                                    If Mid(Inputarr(13, a), k, 1) = " " Or IsNumeric(Mid(Inputarr(13, a), k, 1)) Then
                                    Else
                                        If Trim(Mid(Inputarr(13, a), k + 1, j - k - 1)) <> "" Then
                                            CBA_COM_arrOutput(4, lNum) = Trim(Mid(Inputarr(13, a), k + 1, j - k - 1))
                                            CBA_COM_arrOutput(5, lNum) = "Pieces"
                                        End If
                                        Exit For
                                    End If
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        CBA_COM_arrOutput(4, lNum) = Inputarr(7, a)
                        CBA_COM_arrOutput(5, lNum) = "Pieces"
                    End If
                ElseIf UCase(Inputarr(6, a)) = "G" Then
                    CBA_COM_arrOutput(4, lNum) = Inputarr(7, a)
                    CBA_COM_arrOutput(5, lNum) = "g"
                Else
                    If IsNull(Inputarr(7, a)) And IsNull(Inputarr(8, a)) Then
                        If lNum > 1 Then lNum = lNum - 1 Else lNum = 1
                        ReDim Preserve CBA_COM_arrOutput(1 To 6, 1 To lNum)
                    Else
                        CBA_COM_arrOutput(4, lNum) = Inputarr(7, a)
                        CBA_COM_arrOutput(5, lNum) = LCase(Inputarr(9, a))
                    End If
                End If
            End If
        Next
        
    End If
    
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-DecodePack", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
''decodeerror:
''
''    strAldiMess = "There has been a Depack Error" & Chr(10) & Chr(10) & "Please contact " & g_Get_Dev_Sts("DevUsers") & Chr(10) & Chr(10) & "Error:Decodepack"
''    'frm_Aldimess.Show
''    MsgBox strAldiMess
''    Cancel = True

End Function
Private Function SortCollection(ByVal coll As Collection)
    Dim a As Long, b As Long, oItem, percasetext, redlenby, bOutput As Boolean, strAldiMess, Cancel
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    'On Error GoTo decodeerror
    
    Set CBA_COM_colAdddetail = New Collection
    
    b = 0
    For Each oItem In coll
        b = b + 1
        If InStr(1, UCase(oItem), "+") Then
            CBA_COM_colAdddetail.Add oItem
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "PER PACK OF") > 0 Or InStr(1, UCase(oItem), "PAR CASE OF") > 0 Or InStr(1, UCase(oItem), "PER CASE OF") > 0 _
            Or InStr(1, UCase(oItem), "IN ANY SIX") > 0 Or InStr(1, UCase(oItem), "PER CASE IN") > 0 Or InStr(1, UCase(oItem), "PER PACK (") > 0 _
            Or InStr(1, UCase(oItem), "PER CAN") > 0 Or UCase(oItem) = "PER CASK" Or UCase(oItem) = "PER JAR (IN STORE)" Then
    
        ElseIf InStr(1, UCase(oItem), "THIS PRODUCT IS NOT AVAILABLE") > 0 Or InStr(1, LCase(oItem), "due to its short shelf life") > 0 Then
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "X") And InStr(1, UCase(oItem), "SIX") = 0 Then
            CBA_COM_colMulti.Add oItem
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "-") Then
            CBA_COM_colAdddetail.Add oItem
            coll.Remove (b)
            b = b - 1
        Else
            
            For a = 1 To Len(oItem)
                If Mid(oItem, a, 1) = " " Then
                    CBA_COM_colAdddetail.Add oItem
                    coll.Remove (b)
                    b = b - 1
                    Exit For
                End If
            Next
        End If
    Next
    b = 0
    For Each oItem In coll
    '    If oItem = "ll" Then
    '    a = a
    '    End If
        b = b + 1
        If InStr(1, UCase(oItem), "X") And InStr(1, UCase(oItem), "SIX") = 0 Then
            CBA_COM_colMulti.Add oItem
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "PER CASE OF") > 0 Or InStr(1, UCase(oItem), "PER CASE IN") > 0 Or InStr(1, UCase(oItem), "PAR CASE OF") > 0 Or InStr(1, UCase(oItem), "PER PACK OF") > 0 Or InStr(1, UCase(oItem), "PER PACK (") > 0 Then
            If InStr(1, UCase(oItem), "PER CASE OF") > 0 Then percasetext = "PER CASE OF"
            If InStr(1, UCase(oItem), "PER CASE IN") > 0 Then percasetext = "PER CASE IN"
            If InStr(1, UCase(oItem), "PER CASE OF") > 0 Then percasetext = "PAR CASE OF"
            If InStr(1, UCase(oItem), "PER PACK OF") > 0 Then percasetext = "PAR PACK OF"
            If InStr(1, UCase(oItem), "PER PACK (") > 0 Then percasetext = "PER PACK ("
            If InStr(1, UCase(oItem), "(IN STORE)") > 0 Then redlenby = Len("(IN STORE)") Else If InStr(1, UCase(oItem), "PER PACK (") > 0 Then redlenby = 1 Else If Mid(oItem, Len(oItem), 1) = "." Then redlenby = 1 Else redlenby = 0
                CBA_COM_numOutput = CBA_COM_numOutput + 1
                ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
                CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
    ''            ashjgdas = Trim(Mid(oItem, Len(percasetext) + 1, Len(oItem) - Len(percasetext) - redlenby))
                'ashjgdas = Len(oItem) - Len(percasetext) - redlenby
                If IsNumeric(Trim(Mid(oItem, Len(percasetext) + 1, Len(oItem) - Len(percasetext) - redlenby))) Then
                    CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = CLng(Trim(Mid(oItem, Len(percasetext) + 1, Len(oItem) - Len(percasetext) - redlenby)))
                Else
                    CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1
                End If
                    CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces"
                    coll.Remove (b)
                    b = b - 1
        ElseIf InStr(1, UCase(oItem), "IN ANY SIX") > 0 Or UCase(oItem) = "PER CAN" Or UCase(oItem) = "PER CAN (IN STORE)" Or UCase(oItem) = "PER CAN (IN-STORE)" Or UCase(oItem) = "PER CASK" Or UCase(oItem) = "PER JAR (IN STORE)" Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf UCase(oItem) = "NULL" Or UCase(oItem) = "SMALL" Or UCase(oItem) = "MEDIUM" Or UCase(oItem) = "LARGE" Or UCase(oItem) = "JUMBO" Or UCase(oItem) = "BUNCH" Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "SHEET") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            If Mid(oItem, 1, InStr(1, UCase(oItem), "SHEET")) = 1 Then CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1 Else CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "SHEET") - 1))
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Sheets"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "PACK") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "PACK") - 1))
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf UCase(Trim(oItem)) = "EACH" Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf UCase(Trim(oItem)) = "SMALL" Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "PIECE") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "PIECE") - 1))
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "PCE") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "PCE") - 1))
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "PG") Then
            CBA_COM_potPieces.Add oItem
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "KG") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            If UCase(oItem) = "PERKG" Or UCase(oItem) = "KG" Then CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1000 Else CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "KG") - 1)) * 1000
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "g"
            CBA_COM_PackarrOutput(4, CBA_COM_numOutput) = 1
            CBA_COM_PackarrOutput(5, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "MM") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "MM") - 1))
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "mm"
            CBA_COM_PackarrOutput(4, CBA_COM_numOutput) = 1
            CBA_COM_PackarrOutput(5, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "CM") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "CM") - 1)) * 10
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "mm"
            CBA_COM_PackarrOutput(4, CBA_COM_numOutput) = 1
            CBA_COM_PackarrOutput(5, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "SS") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "SS") - 1))
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Sheets"
            CBA_COM_PackarrOutput(4, CBA_COM_numOutput) = 1
            CBA_COM_PackarrOutput(5, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "ML") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "ML") - 1))
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "ml"
            CBA_COM_PackarrOutput(4, CBA_COM_numOutput) = 1
            CBA_COM_PackarrOutput(5, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "LT") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "LT") - 1)) * 1000
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "ml"
            CBA_COM_PackarrOutput(4, CBA_COM_numOutput) = 1
            CBA_COM_PackarrOutput(5, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "EA") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            If Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "EA") - 1)) = "" Then CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1 Else CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "EA") - 1))
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "PK") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            If Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "PK") - 1)) = "" Then CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1 Else CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "PK") - 1))
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "PC") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            If Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "PC") - 1)) = "" Then CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1 Else CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "PC") - 1))
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "G") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces"
            CBA_COM_PackarrOutput(4, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "G") - 1))
            CBA_COM_PackarrOutput(5, CBA_COM_numOutput) = "g"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "M") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            If UCase(oItem) = "M/L" Then Else CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "M") - 1)) * 1000
            If UCase(oItem) = "M/L" Then Else CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "mm"
            If UCase(oItem) = "M/L" Then CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1 Else CBA_COM_PackarrOutput(4, CBA_COM_numOutput) = 1
            If UCase(oItem) = "M/L" Then CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces" Else CBA_COM_PackarrOutput(5, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "LL") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            If UCase(oItem) = "LL" Then CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1000
            If UCase(oItem) = "LL" Then CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "ml"
            If UCase(oItem) = "LL" Then CBA_COM_PackarrOutput(4, CBA_COM_numOutput) = 1
            If UCase(oItem) = "LL" Then CBA_COM_PackarrOutput(5, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf InStr(1, UCase(oItem), "L") Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            If UCase(oItem) = "L" Then Else If IsNumeric(Mid(oItem, 1, Len(oItem) - 1)) Then CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = Trim(Mid(oItem, 1, InStr(1, UCase(oItem), "L") - 1)) * 1000
            If UCase(oItem) = "L" Then Else CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "ml"
            If UCase(oItem) = "L" Then CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = 1 Else CBA_COM_PackarrOutput(4, CBA_COM_numOutput) = 1
            If UCase(oItem) = "L" Then CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces" Else CBA_COM_PackarrOutput(5, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        ElseIf IsNumeric(oItem) Then
            CBA_COM_numOutput = CBA_COM_numOutput + 1
            ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
            CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(2, CBA_COM_numOutput) = oItem
            CBA_COM_PackarrOutput(3, CBA_COM_numOutput) = "Pieces"
            coll.Remove (b)
            b = b - 1
        End If
    Next
    
    bOutput = CBA_COM_PackDecode.DecodeAddData
    
    Exit Function
    
decodeerror:
    
    strAldiMess = "There has been a Depack Error" & Chr(10) & Chr(10) & "Please contact " & g_Get_Dev_Sts("DevUsers") & Chr(10) & Chr(10) & "Error:SortCollection"
    ''frm_Aldimess.Show
    MsgBox strAldiMess
    
    Cancel = True
Exit_Routine:
    On Error Resume Next
    Exit Function
'
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-SortCollection", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function DecodeAddData()
    Dim numSpace, place, numAddDetail As Long, a As Long, b As Long, c As Long, d As Long, Loca, strAldiMess, Cancel
    Dim Linked, LinkedWord, bfound As Boolean
    Dim LinkedVal As Single
    Dim strLinkedWord As String, oItem
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    'On Error GoTo decodeerror
    
    
    numAddDetail = 0
    For Each oItem In CBA_COM_colAdddetail
        '****************DEBUG TOOL***************************
        'If oItem = "4x140gpk" Then
        'a = a
        'End If
        '****************DEBUG TOOL***************************
    
        numAddDetail = numAddDetail + 1
        Set CBA_COM_colWhere = New Collection
        ReDim Preserve CBA_COM_arrSortDetail(1 To 19, 1 To numAddDetail)
        CBA_COM_arrSortDetail(1, numAddDetail) = oItem
        numSpace = 0
        For a = 1 To Len(oItem)
            If Mid(oItem, a, 1) = " " Or Mid(oItem, a, 1) = "+" Or Mid(oItem, a, 1) = "-" Then
                numSpace = numSpace + 1
                CBA_COM_colWhere.Add a
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
    'If a = 84 Then
    'a = a
    'arrSortDetail(1, a) = arrSortDetail(1, a)
    'End If
    
    
    
    
    Linked = False
        c = 1
        CBA_COM_numOutput = CBA_COM_numOutput + 1
        ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
        CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = CBA_COM_arrSortDetail(1, a)
        
        
    ''****************DEBUG TOOL***************************
    'If CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = "4x140gpk" Then
    'a = a
    'End If
    ''****************DEBUG TOOL***************************
    
        
        
        For b = 2 To 19
            
            If IsEmpty(CBA_COM_arrSortDetail(b, a)) Or CBA_COM_arrSortDetail(b, a) = "" Then
            '*******CONDITION HANDLE RQD*******
            Else
                If InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "+") Then
                'removing anthing after a "+" sign.. THIS MAY NEED TO BE AMENDED
    
                ElseIf UCase(CBA_COM_arrSortDetail(b, a)) = "GIFT" And UCase(CBA_COM_arrSortDetail(b + 1, a)) = "BOX" Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = 1
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                    b = b + 1
                ElseIf LinkedWord = True Then
                    LinkedWord = False
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = strLinkedWord
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = CBA_COM_arrSortDetail(b, a)
                ElseIf Linked = True Then
                    Linked = False
                    If UCase(CBA_COM_arrSortDetail(b, a)) = "PAIR" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PAIRS" Then
                        c = c + 1
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = LinkedVal * 2
                        c = c + 1
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                        GoTo Nexttogo
                    End If
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = LinkedVal
                    c = c + 1
                    If UCase(CBA_COM_arrSortDetail(b, a)) = "SATCHET" Or UCase(CBA_COM_arrSortDetail(b, a)) = "SATCHETS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PK" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "BAG" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PUNNET" Or UCase(CBA_COM_arrSortDetail(b, a)) = "POUCH" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "TUBE" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BLOCK" Or UCase(CBA_COM_arrSortDetail(b, a)) = "POT" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "BOX" Or UCase(CBA_COM_arrSortDetail(b, a)) = "SINGLE" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BAR" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "BAGS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "MULTIPACK" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PACK" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "PKS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PACKS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "TUBES" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "BLOCKS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "POUCHES" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BARS" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "TUB" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BOXES" Or UCase(CBA_COM_arrSortDetail(b, a)) = "CONVENIENCE" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "TUBS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "EACH" Or UCase(CBA_COM_arrSortDetail(b, a)) = "CAPSULES" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "CAPSULE" Or UCase(CBA_COM_arrSortDetail(b, a)) = "JUMBO" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BOTTLE" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "BOTTLES" Or UCase(CBA_COM_arrSortDetail(b, a)) = "CASE" Or UCase(CBA_COM_arrSortDetail(b, a)) = "CASES" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "STICK" Or UCase(CBA_COM_arrSortDetail(b, a)) = "STICKS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PIECE" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "PIECES" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BUNCH" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BUNCHES" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "AVAILABLE" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PIECE" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BUNDLE" Then
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                    ElseIf UCase(CBA_COM_arrSortDetail(b, a)) = "SHEET" Or UCase(CBA_COM_arrSortDetail(b, a)) = "SHEETS" Then
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Sheets"
                    Else
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = CBA_COM_arrSortDetail(b, a)
                    End If
                ElseIf IsNumeric(CBA_COM_arrSortDetail(b, a)) Then
                'CBA_COM_arrSortDetail(1, a) = CBA_COM_arrSortDetail(b, a)
                    Linked = True
                    LinkedVal = CBA_COM_arrSortDetail(b, a)
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "SERVE") Then
                    LinkedWord = True
                    strLinkedWord = CBA_COM_arrSortDetail(b, a)
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "BAG") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PUNNET") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "POUCH") _
                        Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "TUB") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "BLOCK") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "POT") _
                        Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PACK") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "SINGLE") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "BAR") _
                        Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CONVNIENCE") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CONVENIENCE") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "BOTTLE") _
                        Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "EACH") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CAPSULE") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CASE") _
                        Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "STICK") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "JUMBO") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "ROLL") _
                        Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "LOAF") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "BULK") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "WHOLE") _
                        Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "HALF") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "TALL") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "ENVELOPE") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "BUNDLE") Then
                    bfound = False
                    For d = 2 To c
                        If CBA_COM_PackarrOutput(d, CBA_COM_numOutput) = "Pieces" Then
                            bfound = True
                            Exit For
                        End If
                    Next
                    If bfound = False Then
                        c = c + 1
                        If Linked = True Then
                            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = LinkedVal
                            Linked = False
                        Else
                            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = 1
                        End If
                            c = c + 1
                            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                    End If
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "1PLY") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "2PLY") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "3PLY") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "4PLY") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = 1
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "BO") Then
                    bfound = False
                    For d = 2 To c
                        If CBA_COM_PackarrOutput(d, CBA_COM_numOutput) = "Pieces" Then
                            bfound = True
                            Exit For
                        End If
                    Next
                    If bfound = False Then
                        c = c + 1
                        If Linked = True Then
                            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = LinkedVal
                            Linked = False
                        Else
                            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = 1
                        End If
                            c = c + 1
                            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                    End If
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PCE") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PCE") - 1)
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PER") And (InStr(1, UCase(CBA_COM_arrSortDetail(b + 1, a)), "KG") Or InStr(1, UCase(CBA_COM_arrSortDetail(b + 1, a)), "KILO")) Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = 1
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = 1000
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "g"
                    b = b + 1
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "KG") Then
                    c = c + 1
                    If UCase(CBA_COM_arrSortDetail(b, a)) = "KG" Then CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = 1000 Else CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "KG") - 1) * 1000
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "g"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "ML") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "ML") - 1)
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "ml"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "LT") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "LT") - 1) * 1000
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "ml"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PG") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PG") - 1)
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Page"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "SS") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "SS") - 1)
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Sheets"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PK") And (IsNumeric(CBA_COM_arrSortDetail(b + 1, a)) And IsEmpty(CBA_COM_arrSortDetail(b + 1, a)) = False) Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = CBA_COM_arrSortDetail(b + 1, a)
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                    b = b + 1
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PK") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PK") - 1)
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CE") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CE") - 1)
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CM") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CM") - 1) * 10
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "mm"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "MM") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "MM") - 1)
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "mm"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "L") And UCase(CBA_COM_arrSortDetail(b, a)) <> "COLLECTION" And UCase(CBA_COM_arrSortDetail(b, a)) <> "TRAVEL" And UCase(CBA_COM_arrSortDetail(b, a)) <> "ENVELOPE" And UCase(CBA_COM_arrSortDetail(b, a)) <> "SHELF" And UCase(CBA_COM_arrSortDetail(b, a)) <> "LIFE" And UCase(CBA_COM_arrSortDetail(b, a)) <> "VALUE" And UCase(CBA_COM_arrSortDetail(b, a)) <> "LARGE" Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "L") - 1) * 1000
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "ml"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "G") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "G") - 1)
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "g"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "M") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "M") - 1) * 1000
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "mm"
                Else
                    If CBA_COM_arrSortDetail(b, a) <> "" Then CBA_COM_colNotDecoded.Add CBA_COM_arrSortDetail(b, a)
                End If
            End If
Nexttogo:
        Next
        bfound = False
        For d = 2 To c
            If CBA_COM_PackarrOutput(d, CBA_COM_numOutput) = "Pieces" Then
                bfound = True
                Exit For
            End If
        Next
        If bfound = False Then
            c = c + 1
            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = 1
            c = c + 1
            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
        End If
    Next
    
    Erase CBA_COM_arrSortDetail
    Set CBA_COM_colAdddetail = New Collection
    Exit Function
    
decodeerror:
    
    strAldiMess = "There has been a Depack Error" & Chr(10) & Chr(10) & "Please contact " & g_Get_Dev_Sts("DevUsers") & Chr(10) & Chr(10) & "Error:DecodeAddData"
    ''frm_Aldimess.Show
    MsgBox strAldiMess
    Cancel = True
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-DecodeAddData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function DecodeXData()
    Dim numSpace As Long, place As Long, numAddDetail As Long, strAldiMess, Cancel
    Dim Linked As Boolean, LinkedWord As Boolean, bfound As Boolean, oItem, Loca, a As Long, b As Long, c As Long, d As Long
    Dim LinkedVal As Single, multi
    Dim strLinkedWord As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    'On Error GoTo depackerror
    
    
    numAddDetail = 0
    For Each oItem In CBA_COM_colMulti
    
        ''****************DEBUG TOOL***************************
        'If oItem = "4x140gpk" Then
        'a = a
        'End If
        ''****************DEBUG TOOL***************************
    
    
        numAddDetail = numAddDetail + 1
        Set CBA_COM_colWhere = New Collection
        ReDim Preserve CBA_COM_arrSortDetail(1 To 19, 1 To numAddDetail)
        CBA_COM_arrSortDetail(1, numAddDetail) = oItem
        numSpace = 0
        For a = 1 To Len(oItem)
            If Mid(oItem, a, 1) = " " Or Mid(oItem, a, 1) = "+" Or Mid(oItem, a, 1) = "-" Or Mid(UCase(oItem), a, 1) = "X" Or (IsNumeric(Mid(UCase(oItem), a, 1)) And Mid(UCase(oItem), a + 1, 3) = "GPK") Then
                If (IsNumeric(Mid(UCase(oItem), a, 1)) And Mid(UCase(oItem), a + 1, 3) = "GPK") Then
                    numSpace = numSpace + 1
                    CBA_COM_colWhere.Add a + 2
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

''****************DEBUG TOOL***************************
'If CBA_COM_arrSortDetail(1, a) = "4x140gpk" Then
'a = a
'End If
''****************DEBUG TOOL***************************


    Linked = False
    c = 1
    CBA_COM_numOutput = CBA_COM_numOutput + 1
    ReDim Preserve CBA_COM_PackarrOutput(1 To 19, 1 To CBA_COM_numOutput)
    CBA_COM_PackarrOutput(1, CBA_COM_numOutput) = CBA_COM_arrSortDetail(1, a)
    multi = 1
'    If UCase(CBA_COM_PackarrOutput(1, CBA_COM_numOutput)) = "1.2L TUB" Then
'    a = a
'    End If
    
        
        For b = 2 To 19
            
            If IsEmpty(CBA_COM_arrSortDetail(b, a)) Or CBA_COM_arrSortDetail(b, a) = "" Then
            '*******CONDITION HANDLE RQD*******
            Else
                If UCase(Mid(CBA_COM_arrSortDetail(b + 1, a), 1, 1)) = "-" Then GoTo Nexttogo
                If UCase(Mid(CBA_COM_arrSortDetail(b, a), 1, 1)) = "-" Then CBA_COM_arrSortDetail(b, a) = Mid(CBA_COM_arrSortDetail(b, a), 2, Len(CBA_COM_arrSortDetail(b, a)) - 1)
                If Linked = False And UCase(Mid(CBA_COM_arrSortDetail(b, a), 1, 1)) = "X" Then CBA_COM_arrSortDetail(b, a) = Mid(CBA_COM_arrSortDetail(b, a), 2, Len(CBA_COM_arrSortDetail(b, a)) - 1)
                If InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "+") Then
                'removing anthing after a "+" sign.. THIS MAY NEED TO BE AMENDED
                ElseIf UCase(CBA_COM_arrSortDetail(b, a)) = "GIFT" And UCase(CBA_COM_arrSortDetail(b + 1, a)) = "BO" Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = 1
                    multi = CBA_COM_PackarrOutput(c, CBA_COM_numOutput)
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                    b = b + 1
                ElseIf LinkedWord = True Then
                    LinkedWord = False
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = strLinkedWord
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = CBA_COM_arrSortDetail(b, a)
                ElseIf Linked = True Then
                    Linked = False
                    If UCase(CBA_COM_arrSortDetail(b, a)) = "PAIR" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PAIRS" Then
                        c = c + 1
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = LinkedVal * 2
                        multi = CBA_COM_PackarrOutput(c, CBA_COM_numOutput)
                        c = c + 1
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                        GoTo Nexttogo
                    End If
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = LinkedVal
                    multi = CBA_COM_PackarrOutput(c, CBA_COM_numOutput)
                    c = c + 1
                    If UCase(CBA_COM_arrSortDetail(b, a)) = "SATCHET" Or UCase(CBA_COM_arrSortDetail(b, a)) = "SATCHETS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PK" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "BAG" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PUNNET" Or UCase(CBA_COM_arrSortDetail(b, a)) = "POUCH" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "TUBE" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BLOCK" Or UCase(CBA_COM_arrSortDetail(b, a)) = "POT" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "BOX" Or UCase(CBA_COM_arrSortDetail(b, a)) = "SINGLE" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BAR" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "BAGS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "MULTIPACK" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PACK" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "PKS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PACKS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "TUBES" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "BLOCKS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "POUCHES" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BARS" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "TUB" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BOXES" Or UCase(CBA_COM_arrSortDetail(b, a)) = "CONVENIENCE" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "TUBS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "EACH" Or UCase(CBA_COM_arrSortDetail(b, a)) = "CAPSULES" _
                       Or UCase(CBA_COM_arrSortDetail(b, a)) = "CAPSULE" Or UCase(CBA_COM_arrSortDetail(b, a)) = "JUMBO" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BOTTLE" Or UCase(CBA_COM_arrSortDetail(b, a)) = "BOTTLES" Or UCase(CBA_COM_arrSortDetail(b, a)) = "CASE" Or UCase(CBA_COM_arrSortDetail(b, a)) = "CASES" Or UCase(CBA_COM_arrSortDetail(b, a)) = "STICK" Or UCase(CBA_COM_arrSortDetail(b, a)) = "STICKS" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PIECE" Or UCase(CBA_COM_arrSortDetail(b, a)) = "PIECES" Then
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                    ElseIf UCase(CBA_COM_arrSortDetail(b, a)) = "SHEET" Or UCase(CBA_COM_arrSortDetail(b, a)) = "SHEETS" Then
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Sheets"
                    Else
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = CBA_COM_arrSortDetail(b, a)
                    End If
                ElseIf IsNumeric(CBA_COM_arrSortDetail(b, a)) Then
                    If UCase(Mid(CBA_COM_arrSortDetail(b + 1, a), 1, 1)) = "X" Then
                        c = c + 1
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = CBA_COM_arrSortDetail(b, a)
                        multi = CBA_COM_PackarrOutput(c, CBA_COM_numOutput)
                        c = c + 1
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                    Else
                        Linked = True
                        LinkedVal = CBA_COM_arrSortDetail(b, a)
                    End If
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "SERVE") Then
                    LinkedWord = True
                    strLinkedWord = CBA_COM_arrSortDetail(b, a)
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "BAG") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PUNNET") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "POUCH") _
                        Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "TUB") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "BLOCK") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "POT") _
                        Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PACK") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "SINGLE") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "BAR") _
                        Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CONVNIENCE") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CONVENIENCE") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "BOTTLE") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "EACH") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CAPSULE") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CASE") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "STICK") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "JUMBO") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "SACHET") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "BIG") Or InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "TALL") Then
                    bfound = False
                    For d = 2 To c
                        If CBA_COM_PackarrOutput(d, CBA_COM_numOutput) = "Pieces" Then
                            bfound = True
                            Exit For
                        End If
                    Next
                    If bfound = False Then
                        c = c + 1
                        If Linked = True Then
                            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = LinkedVal
                            multi = CBA_COM_PackarrOutput(c, CBA_COM_numOutput)
                            Linked = False
                        Else
                            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = 1
                            multi = CBA_COM_PackarrOutput(c, CBA_COM_numOutput)
                        End If
                            c = c + 1
                            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                    End If
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "BO") Then
                    bfound = False
                    For d = 2 To c
                        If CBA_COM_PackarrOutput(d, CBA_COM_numOutput) = "Pieces" Then
                            bfound = True
                            Exit For
                        End If
                    Next
                    If bfound = False Then
                        c = c + 1
                        If Linked = True Then
                            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = LinkedVal
                            multi = CBA_COM_PackarrOutput(c, CBA_COM_numOutput)
                            Linked = False
                        Else
                            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = 1
                            multi = CBA_COM_PackarrOutput(c, CBA_COM_numOutput)
                        End If
                            c = c + 1
                            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                    End If
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PCE") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PCE") - 1)
                    multi = CBA_COM_PackarrOutput(c, CBA_COM_numOutput)
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "KG") Then
                    If CBA_COM_arrSortDetail(b, a) = ".1.8kg" Then CBA_COM_arrSortDetail(b, a) = "1.8kg"
                    If CBA_COM_arrSortDetail(b, a) = ".1.25kg" Then CBA_COM_arrSortDetail(b, a) = "1.25kg"
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "KG") - 1) * 1000 * multi
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "g"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "ML") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "ML") - 1) * multi
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "ml"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "LT") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "LT") - 1) * 1000
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "ml"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PG") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PG") - 1) * multi
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pages"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "SS") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "SS") - 1) * multi
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Sheets"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PK") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "PK") - 1)
                    multi = CBA_COM_PackarrOutput(c, CBA_COM_numOutput)
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CE") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CE") - 1)
                    multi = CBA_COM_PackarrOutput(c, CBA_COM_numOutput)
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CM") Then
                    If InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CM") > 0 And InStr(1, UCase(CBA_COM_arrSortDetail(b + 1, a)), "M") > 0 And InStr(1, UCase(CBA_COM_arrSortDetail(b + 1, a)), "CM") = 0 And InStr(1, UCase(CBA_COM_arrSortDetail(b + 1, a)), "MM") = 0 Then
                    Else
                        c = c + 1
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "CM") - 1) * 10 * multi
                        c = c + 1
                        CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "mm"
                    End If
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "MM") Then
    
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "MM") - 1) * multi
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "mm"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "L") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "L") - 1) * 1000 * multi
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "ml"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "G") Then
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "G") - 1) * multi
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "g"
                ElseIf InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "M") Then
                   'If b > 0 Then If UCase(CBA_COM_arrSortDetail(b - 1, a)) = "33CM" And UCase(CBA_COM_arrSortDetail(b, a)) = "60M" Then GoTo Nexttogo
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = Mid(CBA_COM_arrSortDetail(b, a), 1, InStr(1, UCase(CBA_COM_arrSortDetail(b, a)), "M") - 1) * 1000 * multi
                    c = c + 1
                    CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "mm"
                Else
                    If CBA_COM_arrSortDetail(b, a) <> "" Then CBA_COM_colNotDecoded.Add CBA_COM_arrSortDetail(b, a)
                End If
            End If
Nexttogo:
        Next
        bfound = False
        For d = 2 To c
            If CBA_COM_PackarrOutput(d, CBA_COM_numOutput) = "Pieces" Then
                bfound = True
                Exit For
            End If
        Next
        If bfound = False Then
            c = c + 1
            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = 1
            c = c + 1
            CBA_COM_PackarrOutput(c, CBA_COM_numOutput) = "Pieces"
        End If
        
    Next
    
    Erase CBA_COM_arrSortDetail
    Set CBA_COM_colMulti = New Collection
    
    Exit Function
    
depackerror:
    
    strAldiMess = "There has been a Depack Error" & Chr(10) & Chr(10) & "Please contact " & g_Get_Dev_Sts("DevUsers") & Chr(10) & Chr(10) & "Error:DecodeXData"
    ''frm_Aldimess.Show
    MsgBox strAldiMess
Cancel = True

Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-DecodeXData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

