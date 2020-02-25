Attribute VB_Name = "CBA_BasicFunctionsExcel"
Option Explicit
' This module will house all the User Functions available to Excel Users only i.e. not Option Private Module
'''''Option Private Module       ' Excel users CAN access



Public Function MergedCellData(ByVal rng As Range) As String
    MergedCellData = rng.MergeArea.Cells(1, 1).Value
End Function
Public Function getProduceWeek(ByVal isDate As Date) As Long
    getProduceWeek = DatePart("WW", isDate, vbWednesday)
End Function
