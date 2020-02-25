Attribute VB_Name = "CBAR_ShelfComparisons"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Sub createShelfComp()
    Dim strProds As String, stq As String
    Dim bOutput As Boolean
    Dim a As Long, b As Long, lRowNo As Long
    Dim DFrom As Date, Dto As Date
    Dim CG As Long, scg As Long
    Dim wbk_ShelfComp As Workbook
    Dim wks_ShelfComp As Worksheet
    Dim AldiRet As Single, CompRet As Single
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    bOutput = CBAR_SQLQueries.CBAR_GenPullSQL("COM_2ScrapeDates")
    If bOutput = False Then Exit Sub
    
    For a = LBound(CBA_COMarr, 2) To UBound(CBA_COMarr, 2)
        If a = LBound(CBA_COMarr, 2) Then
            DFrom = CBA_COMarr(1, a)
            Dto = CBA_COMarr(1, a)
        Else
            If CBA_COMarr(1, a) < DFrom Then DFrom = CBA_COMarr(1, a)
            If CBA_COMarr(1, a) > Dto Then Dto = CBA_COMarr(1, a)
        End If
    Next


    If CBA_COM_SetupMatchArray.CBA_SetupMatchArray(False, DFrom, Dto, CG, scg, strProds, True) = True Then
        lRowNo = 4
        Set wbk_ShelfComp = Application.Workbooks.Add
        Set wks_ShelfComp = wbk_ShelfComp.Sheets(1)
        
        With wks_ShelfComp
            Range(.Cells(1, 1), .Cells(3, 79)).Interior.ColorIndex = 49
            .Cells(1, 1).Select
            .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
            .Cells.Font.Name = "ALDI SUED Office"
            .Cells(2, 3).Font.Size = 24
            .Cells(2, 3).Font.ColorIndex = 2
            .Cells(4, 1).Value = "Aldi Pcode"
            .Cells(4, 2).Value = "Aldi Description"
            .Cells(4, 3).Value = "Aldi Packsize"
            .Cells(4, 4).Value = "Comp Pcode"
            .Cells(4, 5).Value = "Comp Description"
            .Cells(4, 6).Value = "Comp Packsize"
            .Cells(4, 7).Value = "Aldi Shelf Price"
            .Cells(4, 8).Value = "Comp Shelf Price"
            .Cells(4, 9).Value = "State"
            Range(.Cells(4, 1), .Cells(4, 9)).EntireColumn.AutoFit
            Range(.Cells(4, 1), .Cells(4, 9)).Font.Bold = True
            Range(.Cells(4, 1), .Cells(4, 9)).Font.Underline = True
            .Cells(2, 3).Value = "Shelf Comparison Report"
    
            For a = LBound(CBA_COM_Match, 1) To UBound(CBA_COM_Match, 1)
            
                For b = 1 To 6
                    
                    If b = 1 Then
                        stq = "NSW"
                    ElseIf b = 2 Then stq = "VIC"
                    ElseIf b = 3 Then stq = "QLD"
                    ElseIf b = 4 Then stq = "SA"
                    ElseIf b = 5 Then stq = "WA"
                    ElseIf b = 6 Then stq = "National"
                    End If
                
                    AldiRet = CBA_COM_Match(a).Pricedata(Dto, "aldiretail", stq)
                    CompRet = CBA_COM_Match(a).Pricedata(Dto, "shelf", stq)
                
                    If (AldiRet = CompRet Or AldiRet + 0.01 = CompRet) And AldiRet > 0 And CBA_COM_Match(a).CompMultby <> CBA_COM_Match(a).CompDivideby Then
                        lRowNo = lRowNo + 1
                        .Cells(lRowNo, 1).Value = CBA_COM_Match(a).AldiPCode
                        .Cells(lRowNo, 2).Value = CBA_COM_Match(a).AldiPName
                        If CBA_COM_Match(a).HowComp = "L" Then
                            .Cells(lRowNo, 3).Value = CBA_COM_Match(a).CompMultby & "ml"
                        Else
                            .Cells(lRowNo, 3).Value = CBA_COM_Match(a).CompMultby & CBA_COM_Match(a).HowComp
                        End If
                        .Cells(lRowNo, 4).Value = CBA_COM_Match(a).CompCode
                        .Cells(lRowNo, 5).Value = CBA_COM_Match(a).CompProdName
                        If CBA_COM_Match(a).HowComp = "L" Then
                            .Cells(lRowNo, 6).Value = CBA_COM_Match(a).CompDivideby & "ml"
                        Else
                            .Cells(lRowNo, 6).Value = CBA_COM_Match(a).CompDivideby & CBA_COM_Match(a).HowComp
                        End If
                        .Cells(lRowNo, 7).Value = AldiRet
                        .Cells(lRowNo, 8).Value = CompRet
                        .Cells(lRowNo, 9).Value = stq
                    End If
                
                
                Next
            
            Next
            Range(.Cells(4, 1), .Cells(lRowNo, 9)).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9), Header:=xlYes
    
    
        End With
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-createShelfComp", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
