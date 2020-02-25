Attribute VB_Name = "CBA_COM_Emailer"

Function EmailBDs(ByRef WKSDic As Scripting.Dictionary) As Boolean
    Dim IsCreated As Boolean
    Dim appOutlook As Object
    Dim BDemails As Collection
    Dim thecell, noofBDs As Long
    Dim rng, endcell, startcell As Range
    Dim strReport As String, RCell
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    Set appOutlook = GetObject(, "Outlook.Application")
    If appOutlook Is Nothing Then Set appOutlook = CreateObject("Outlook.Application")
    
    Set BDBADic = getBDBADic
    Set BDDic = New Scripting.Dictionary
    Set GBDDic = New Scripting.Dictionary
    
    BDEmail = CBAR_Runtime.getProdGBDBDEmails
    
    'CBAR_SQLQueries.CBAR_GenPullSQL "CBAR_ProdGBDBDEmail"
    
    For a = LBound(BDEmail, 2) To UBound(BDEmail, 2)
        If GBDDic.Exists(BDEmail(2, a)) = False Then GBDDic.Add BDEmail(2, a), BDEmail(1, a)
        If BDDic.Exists(BDEmail(4, a)) = False Then BDDic.Add BDEmail(4, a), BDEmail(3, a)
    Next a
    
    For Each BD In BDDic
            With appOutlook.CreateItem(0)
            .To = BD
           
            .CC = ""
            .bcc = ""
            .Subject = "COMRADE WEEKLY UPDATE"
            strHTML = "<HTML><BODY @media print{@page {size: landscape}}><p style='font-family:ALDI SUED Office;font-size:20pt'><b>COMRADE Weekly Update</b></p><br>"
            
            For Each wks In WKSDic.Items
                For b = 1 To 99
                    If wks.Cells(5, b).Value = "BD" Then BDCol = b
                    If wks.Cells(5, b).Value = "" Then
                        endCol = b
                        Exit For
                    End If
                Next
                
                Set rng = Nothing
                bfound = False
                X = BDCol
                For Each RCell In wks.Columns(X).Cells
                   If RCell.Value = "" And RCell.Offset(1, 0).Value = "" And RCell.Offset(2, 0).Value = "" And RCell.Offset(3, 0).Value = "" And RCell.Offset(5, 0).Value = "" And RCell.Offset(10, 0).Value = "" Then Exit For
                   If RCell.Row = 5 Then Set rng = Range(wks.Cells(RCell.Row, 1), wks.Cells(RCell.Row, X - 2))
                   'Rng.Select
                   If RCell.Value = BD Then
                        If .CC = "" Then .CC = RCell.Offset(0, 1).Value
                        bfound = True
                        Set rng = Application.Union(rng, Range(wks.Cells(RCell.Row, 1), wks.Cells(RCell.Row, X - 2)))
                   End If
                Next
            
            
            strReport = wks.Cells(3, 3).Value
            
            If bfound = True Then strHTML = strHTML & "<p style='font-family:ALDI SUED Office;font-size:11.5pt'><u><b> " & strReport & "</b></u></p>" & RangetoHTML(rng)
            
            Next
            CBA_ErrTag = "HTML"
            
            strHTML = strHTML & "<p style='font-family:ALDI SUED Office;font-size:11.5pt'><i><b> Please See attached document for further details</b></i></p>"

            strHTML = strHTML & "</BODY></HTML>"
            .htmlBody = strHTML
    
            .Display
            End With
    Next
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-EmailBDs", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    If CBA_ErrTag = "HTML" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Function RangetoHTML(ByVal rng As Range)
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    Dim rnge As Range
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    TempFile = "C:\TempCOMRADE\" & "TEST.htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    Set TempWS = ActiveSheet
    With TempWS
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).CurrentRegion.Borders.LineStyle = xlContinuous
        .Cells(1).CurrentRegion.Borders.Color = vbBlack
        .Cells(1).CurrentRegion.Borders.Weight = xlThin
        .Cells(1).CurrentRegion.Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo Err_Routine
    End With
'    Set rnge = TempWB.Sheets(1).Range(Rng.Address)
'    For Each wb In Application.Workbooks
'        If wb.Name <> TempWB.Name Then
'            wb.Activate
'            Exit For
'        End If
'    Next
    Application.ReferenceStyle = xlA1
    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add(SourceType:=xlSourceRange, Filename:=TempFile, Sheet:=TempWS.Name, Source:=TempWS.UsedRange.Address, HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With


    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-RangetoHTML", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

