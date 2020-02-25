Attribute VB_Name = "CBA_FX"
Option Explicit
Private Sub InterfaceFXRatestoDB()
Dim ABI_CN As ADODB.Connection
Dim ABI_RS As ADODB.Recordset
Dim a As Byte, headrow As Byte
Dim RCell As Range
Dim strSQL As String

    With ActiveSheet
        headrow = 0
        For Each RCell In .Columns(1).Cells
            If RCell.Value = "3 Months" Then
                headrow = RCell.Row
                Exit For
            End If
            If RCell.Row > 40 Then Exit For
        Next
        If headrow = 0 Then
            MsgBox "The Template formatting has changed therefore the rates has NOT been uploaded"
            Exit Sub
        End If
        If .Cells(headrow - 1, 3).Value = "USD" And .Cells(headrow - 1, 4).Value = "EUR" And .Cells(headrow - 1, 5).Value = "GBP" Then
        Else
            MsgBox "The Template formatting has changed therefore the rates has NOT been uploaded"
            Exit Sub
        End If
        Set ABI_CN = New ADODB.Connection
        With ABI_CN
            .CommandTimeout = 50
            .ConnectionTimeout = 50
            .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & CBA_BSA & "LIVE DATABASES\ABI.accdb"
        End With
        For Each RCell In .Columns(1).Cells
            If RCell.Row >= headrow And RCell.Value <> "" Then
                For a = 1 To 4
                    If a = 1 Then
                        Set ABI_RS = New ADODB.Recordset
                        strSQL = "INSERT INTO FXData(YearNo,MonthNo,CurrencyFrom, CurrencyTo,Rate,DateUploaded)"
                        strSQL = strSQL & "VALUES(" & Year(RCell.Offset(0, 1).Value) & "," & Month(RCell.Offset(0, 1).Value) & _
                            ",'AUD','AUD',1,#" & Format(Date, "MM/DD/YYYY") & "#)"
                        ABI_RS.Open strSQL, ABI_CN
                    Else
                        Set ABI_RS = New ADODB.Recordset
                        strSQL = "INSERT INTO FXData(YearNo,MonthNo,CurrencyFrom, CurrencyTo,Rate,DateUploaded)"
                        strSQL = strSQL & "VALUES(" & Year(RCell.Offset(0, 1).Value) & "," & Month(RCell.Offset(0, 1).Value) & _
                            ",'AUD','" & .Cells(headrow - 1, a + 1).Value & "'," & RCell.Offset(0, a).Value & ",#" & Format(Date, "MM/DD/YYYY") & "#)"
                        ABI_RS.Open strSQL, ABI_CN
                    End If
                Next
            ElseIf RCell.Row > 19 And RCell.Value = "" Then
                Exit For
            End If
        Next
        ABI_CN.Close
        Set ABI_CN = Nothing
    End With
    MsgBox "Data Interfaced to FX Database"
End Sub

