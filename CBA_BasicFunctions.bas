Attribute VB_Name = "CBA_BasicFunctions"
Option Explicit                 ' Basic Function Last Upd 29/05/2019
Option Private Module          ' Excel users cannot access procedures

Private CBA_RunningSht As Boolean
Private CBA_wks_Runsht As Worksheet
Private Declare Function GetTickCount Lib "kernel32" () As Long

Function CBA_Current_UserID() As String
    CBA_Current_UserID = Application.UserName
End Function
Function CBA_UserName(ByRef FirstName As String, ByRef LastName As String, Optional ByVal UserName As String) As Boolean
Dim op As String
Dim rgx As RegExp
Dim usrName As Variant

    If UserName = "" Then UserName = CBA_BasicFunctions.CBA_Current_UserID
    Set rgx = New RegExp
    With rgx
        .IgnoreCase = True
        .Global = True
'        .Pattern = "[^0-9a-zA-Z]+"
'        .Pattern = "^[^\(]+"
        .Pattern = "\([^(\r\n]*?\)|\(|\)"
        usrName = Split(.Replace(UserName, ""), ",")
        
    End With
    If UBound(usrName) > -1 Then CBA_UserName = True: LastName = Trim(usrName(0))
    If UBound(usrName) > 0 Then FirstName = Trim(usrName(1))
    
    
End Function
Function CBA_getRSCount(ByRef RS As ADODB.Recordset) As Long
Dim r As ADODB.Recordset
Dim cnt As Long
    
    cnt = -1
    If RS Is Nothing Then
    Else
        RS.MoveFirst
        Do Until RS.EOF
            cnt = cnt + 1
            RS.MoveNext
        Loop
        RS.MoveFirst
    End If
    CBA_getRSCount = cnt
End Function
Public Function CBA_IsoYearStart(WhichYear As Integer) As Date
' First published by John Green, Excel MVP, Sydney, Australia
    Dim WeekDay As Integer
    Dim NewYear As Date
    NewYear = DateSerial(WhichYear, 1, 1)
    WeekDay = (NewYear - 2) Mod 7
    If WeekDay < 4 Then
        CBA_IsoYearStart = NewYear - WeekDay
    Else
        CBA_IsoYearStart = NewYear - WeekDay + 7
    End If
End Function
Public Function CBA_GetActiveAddin() As String
Dim a As Long
    For a = 1 To Application.AddIns2.Count
        If Application.AddIns2.Item(a).IsOpen = True And Application.AddIns2.Item(a).Installed = True And InStr(1, Application.AddIns2.Item(a).Name, "CBStdAddinW10") > 0 Then
                CBA_GetActiveAddin = Application.AddIns2.Item(a).Name
                Exit For
        End If
    Next
End Function


Public Function CopyDictionary(ByVal SD As Scripting.Dictionary) As Scripting.Dictionary 'if the item is an object tis doesnt work. All objects must be built with a copy function, at which point this function becomes irrelevent
  Dim newDict As Scripting.Dictionary
  Dim k As Variant
  Set newDict = CreateObject("Scripting.Dictionary")
  For Each k In SD.Keys
    newDict.Add k, SD(k)
  Next
  newDict.CompareMode = SD.CompareMode
  Set CopyDictionary = newDict
End Function
Public Function CBA_IsoWeekNumber(D1 As Date) As Integer
' Attributed to Daniel Maher
    Dim D2 As Long
    D2 = DateSerial(Year(D1 - WeekDay(D1 - 1) + 4), 1, 3)
    CBA_IsoWeekNumber = Int((D1 - D2 + WeekDay(D2) + 5) / 7)
End Function
Function TranslateServerName(ByVal ServerName As String, ByVal theDate As Date) As String
    Dim CBA_Proc As String
    On Error GoTo Err_Routine
    If InStr(1, ServerName, "509") > 0 Then
        If theDate >= #3/19/2019# Then TranslateServerName = "0509Z0IDBSRVL02" Else TranslateServerName = "509DBL01\SR"
    ElseIf InStr(1, ServerName, "507") > 0 Then
        If theDate >= #3/27/2019# Then TranslateServerName = "0507Z0IDBSRVL02" Else TranslateServerName = "507DBL01\SR"
    ElseIf InStr(1, ServerName, "506") > 0 Then
        If theDate >= #3/27/2019# Then TranslateServerName = "0506Z0IDBSRVL02" Else TranslateServerName = "506DBL01\SR"
    ElseIf InStr(1, ServerName, "503") > 0 Then
        If theDate >= #3/28/2019# Then TranslateServerName = "0503Z0IDBSRVL02" Else TranslateServerName = "503DBL01\SR"
    ElseIf InStr(1, ServerName, "504") > 0 Then
        If theDate >= #3/28/2019# Then TranslateServerName = "0504Z0IDBSRVL02" Else TranslateServerName = "504DBL01\SR"
    ElseIf InStr(1, ServerName, "505") > 0 Then
        If theDate >= #4/9/2019# Then TranslateServerName = "0505Z0IDBSRVL02" Else TranslateServerName = "505DBL01\SR"
    ElseIf InStr(1, ServerName, "502") > 0 Then
        If theDate >= #4/10/2019# Then TranslateServerName = "0502Z0IDBSRVL02" Else TranslateServerName = "502DBL01\SR"
    ElseIf InStr(1, ServerName, "501") > 0 Then
        If theDate >= #4/2/2019# Then TranslateServerName = "0501Z0IDBSRVL02" Else TranslateServerName = "501DBL01\SR"
    ElseIf InStr(1, ServerName, "599DBL12") > 0 Then        ' #TP Added 191122 & 191203
        If theDate >= #12/4/2019# Then TranslateServerName = "0599Z0NDBREPL01" Else TranslateServerName = "599DBL12"
    ElseIf InStr(1, ServerName, "0599Z0NDBREPL01") > 0 Then ' #TP Added 191203
        TranslateServerName = "0599Z0NDBREPL01"
    ElseIf InStr(1, ServerName, "599DBL11") > 0 Then        ' @TP 18022020
        If theDate >= #1/18/2020# Then TranslateServerName = "0599Z0NDBSRVL01" Else TranslateServerName = "599DBL11"
    ElseIf InStr(1, ServerName, "599") > 0 Then
        TranslateServerName = "599DBL01"
    Else
        TranslateServerName = ServerName
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("TranslateServerName", 3)

    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Public Function g_GetTickCount() As Long
    ' Will get the length of time in seconds since the computer started
    g_GetTickCount = GetTickCount() / 1000
End Function
Function g_DivZero(ByRef Numerator, ByRef DivByValue) As Single
    If IsNumeric(DivByValue) And IsNumeric(Numerator) Then
        If DivByValue = 0 Then
            g_DivZero = 0
        Else
            g_DivZero = Numerator / DivByValue
        End If
    Else
        g_DivZero = 0
    End If
End Function

Function g_WorkSheetExists(ByVal wkb As Workbook, sSheetToTest As String) As Boolean
    ' This routine will test for the existance of a worksheet in a workbook, returning a true if it exists
    On Error GoTo Err_Routine
    Dim wks As Worksheet, vCell As cell
    g_WorkSheetExists = True
    Set wks = wkb.Worksheets(sSheetToTest)
    Exit Function
Err_Routine:
g_WorkSheetExists = False
End Function

Sub g_WorkSheetDelete(ByVal wkb As Workbook, sSheetToTest As String)
    ' This routine will test for the existance of a worksheet in a workbook, and delete it if it does
    On Error GoTo Err_Routine
    Dim wks As Worksheet, bDisplay As Boolean
    Set wks = wkb.Worksheets(sSheetToTest)
    Application.ScreenUpdating = False
    bDisplay = Application.DisplayAlerts
    Application.DisplayAlerts = False
    DoEvents
    wks.Delete
    DoEvents
''    Exit Sub
Err_Routine:
    On Error Resume Next
    Application.DisplayAlerts = bDisplay
    Application.ScreenUpdating = True

End Sub

Function g_WorkBookExists(ByVal sName As String) As Boolean
    ' This routine will test for the existance of a workbook
    Dim wbWB As Workbook
    On Error GoTo Err_Routine
    g_WorkBookExists = False
    ' Find out which workbook to test
    For Each wbWB In Workbooks
        If wbWB.Name = sName Then
            g_WorkBookExists = True
            Exit For
        End If
    Next wbWB
    Exit Function
Err_Routine:

End Function

Sub g_WorkBookDelete(ByVal sName As String, Optional bActivateOther As Boolean = True)
    ' This routine will test for the existance of a workbook, and delete it if it does
    Dim wbWB As Workbook, wbOld As Workbook, bDisplay As Boolean, sOldName As String
    On Error GoTo Err_Routine
    bDisplay = Application.DisplayAlerts
    ' Find out which workbook to get
    If bActivateOther Then
        For Each wbWB In Workbooks
            If wbWB.Name <> sName Then
                sOldName = wbWB.Name
                wbWB.Activate
                Exit For
            End If
        Next wbWB
    End If
    Workbooks(sName).Close False
    Set Workbooks(sName) = Nothing

    Exit Sub
Err_Routine:

End Sub

Public Function GetDayFromWeekNumber(InYear As Integer, _
                WeekNumber As Integer, _
                Optional DayInWeek1Monday7Sunday As Integer = 1) As Date
    Dim i As Integer: i = 1
    If DayInWeek1Monday7Sunday < 1 Or DayInWeek1Monday7Sunday > 7 Then
        MsgBox "Please input between 1 and 7 for the argument :" & vbCrLf & _
                "DayInWeek1Monday7Sunday!", vbOKOnly + vbCritical
        'Function will return 30/12/1899 if you don't use a good DayInWeek1Monday7Sunday
        Exit Function
    Else
    End If

    Do While WeekDay(DateSerial(InYear, 1, i), vbMonday) <> DayInWeek1Monday7Sunday
        i = i + 1
    Loop

    GetDayFromWeekNumber = DateAdd("ww", WeekNumber - 1, DateSerial(InYear, 1, i))
End Function

Function ExitQuery(ByVal Message As String)
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    Application.EnableEvents = False
    Application.ScreenUpdating = True
    DoEvents
    MsgBox Message, vbOKOnly
    DoEvents
    Application.Calculation = xlCalculationAutomatic

End Function
Function CBA_TransposeArray(ByRef arr() As Variant) As Variant()
Dim a As Long, b As Long
Dim thisarr() As Variant
    On Error Resume Next
    ReDim thisarr(LBound(arr, 2) To UBound(arr, 2), LBound(arr, 1) To UBound(arr, 1))
    For a = LBound(arr, 2) To UBound(arr, 2)
        For b = LBound(arr, 1) To UBound(arr, 1)
            thisarr(a, b) = arr(b, a)
        Next
    Next
CBA_TransposeArray = thisarr
Err.Clear
On Error GoTo 0
End Function
Function CBA_DivtoReg(ByVal CBA_Div As Variant) As String

    Select Case CBA_Div
        Case 501
            CBA_DivtoReg = "Minchinbury"
        Case 502
            CBA_DivtoReg = "Derrimut"
        Case 503
            CBA_DivtoReg = "Staplyton"
        Case 504
            CBA_DivtoReg = "Prestons"
        Case 505
            CBA_DivtoReg = "Dandenong"
        Case 506
            CBA_DivtoReg = "Brendale"
        Case 507
            CBA_DivtoReg = "RegencyPark"
        Case 509
            CBA_DivtoReg = "Jandakot"
        Case 599
            CBA_DivtoReg = "National"
        Case "Minchinbury"
            CBA_DivtoReg = "501"
        Case "Derrimut"
            CBA_DivtoReg = "502"
        Case "Staplyton"
            CBA_DivtoReg = "503"
        Case "Prestons"
            CBA_DivtoReg = "504"
        Case "Dandenong"
            CBA_DivtoReg = "505"
        Case "Brendale"
            CBA_DivtoReg = "506"
        Case "RegencyPark"
            CBA_DivtoReg = "507"
        Case "Jandakot"
            CBA_DivtoReg = "509"
        Case "National"
            CBA_DivtoReg = "599"
    End Select
End Function

Function CBA_Running(Optional ByVal strMsg As String)
Dim wbk As Workbook
    If strMsg = "" Then strMsg = "Please wait..."

    CBA_strAldiMsg = strMsg
    If CBA_RunningSht = True Then
        Call frm_Splash.SplashForm
    Else
        frm_Splash.Show vbModeless
        Call frm_Splash.SplashForm("Init")
    End If
    CBA_RunningSht = True

End Function

Function CBA_Close_Running()
    On Error Resume Next
    CBA_strAldiMsg = ""
    Unload frm_Splash
    CBA_RunningSht = False
    On Error GoTo 0
End Function

Function CBA_CollectionSort(ByRef Coltosort)
    'Declare the variables
    Dim coll As New Collection
    Dim arr() As Variant
    Dim temp As Variant
    Dim i As Long
    Dim j As Long
    On Error GoTo Err_Routine
        
    Set coll = Coltosort
    'Allocate storage space for the dynamic array
    ReDim arr(1 To coll.Count)
    
    'Fill the array with items from the Collection
    For i = 1 To coll.Count
        arr(i) = coll.Item(i)
    Next i
    
    'Sort the array using the bubble sort method
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(j)
                arr(j) = arr(i)
                arr(i) = temp
            End If
        Next j
    Next i
    
    'Remove all items from the Collection
    Set coll = Nothing
    
    'Add the sorted items from the array back to the Collection
    For i = LBound(arr) To UBound(arr)
        coll.Add Item:=arr(i)
    Next i
    
    Set Coltosort = coll
'    'Build a list of items from the Collection
'    For Each Itm In Coll
'        Txt = Txt & Itm & vbCrLf
'    Next Itm
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_CollectionSort", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function

Function CBA_sortCollection(ByRef col As Collection)
    Dim vTemp As Variant
    Dim i As Long, j As Long
    On Error GoTo Err_Routine
    'Two loops to bubble sort
    If col.Count > 2 Then
        For i = 1 To col.Count - 1
            For j = i + 1 To col.Count
'              '  Debug.Print col(i)
'              '  Debug.Print col(j)
                If col(i) > col(j) Then
                    'store the lesser item
                    vTemp = col(j)
                    'remove the lesser item
                    col.Remove j
                    're-add the lesser item before the greater Item
                    If i > 1 Then
                        If col(i - 1) <> vTemp Then
                            col.Add vTemp, vTemp, i
                        End If
                    Else
                        col.Add vTemp, vTemp, i
                    End If
                End If
            Next j
        Next i
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_sortCollection", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Function isRunningSheetDisplayed() As Boolean
'CBA_RunningSht = False
    isRunningSheetDisplayed = CBA_RunningSht
End Function
Function RunningSheetAddComment(ByVal r As Long, ByVal c As Long, ByVal AddComment As String)
    If CBA_RunningSht = True Then
    CBA_strAldiMsg = AddComment
    Call frm_Splash.SplashForm

    DoEvents
    End If
End Function
Public Sub CBA_UniqueValuesFor1DArray(ByRef vArray As Variant)
Dim dic As Scripting.Dictionary
Dim a As Long
Dim v As Variant
    Set dic = New Scripting.Dictionary
    For a = LBound(vArray) To UBound(vArray)
        If dic.Exists(CStr(vArray(a))) = False Then dic.Add CStr(vArray(a)), CStr(vArray(a))
    Next
    ReDim vArray(0 To dic.Count - 1)
    a = -1
    For Each v In dic.Keys
        a = a + 1: vArray(a) = v
    Next
End Sub
Sub CBA_Sort1DArray(vArray As Variant, arrLbound As Long, arrUbound As Long)
'Sorts a one-dimensional VBA array from smallest to largest using a very fast quicksort algorithm variant.
Dim pivotVal As Variant
Dim vSwap    As Variant
Dim tmpLow   As Long
Dim tmpHi    As Long
     On Error GoTo Err_Routine
tmpLow = arrLbound
tmpHi = arrUbound
pivotVal = vArray((arrLbound + arrUbound) \ 2)
 
While (tmpLow <= tmpHi) 'divide
   While (vArray(tmpLow) < pivotVal And tmpLow < arrUbound)
      tmpLow = tmpLow + 1
   Wend
  
   While (pivotVal < vArray(tmpHi) And tmpHi > arrLbound)
      tmpHi = tmpHi - 1
   Wend
 
   If (tmpLow <= tmpHi) Then
      vSwap = vArray(tmpLow)
      vArray(tmpLow) = vArray(tmpHi)
      vArray(tmpHi) = vSwap
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
   End If
Wend
 
  If (arrLbound < tmpHi) Then CBA_Sort1DArray vArray, arrLbound, tmpHi
  If (tmpLow < arrUbound) Then CBA_Sort1DArray vArray, tmpLow, arrUbound
  
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_Sort2DArray", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Public Sub CBA_Sort2DArray(ByRef pvarArray As Variant, plngDim As Long, plngCol As Long, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As Variant
    Dim varSwap As Variant
    Dim c As Long
    Dim cMin As Long
    Dim cMax As Long
    On Error GoTo Err_Routine
        
    cMin = LBound(pvarArray, plngDim)
    cMax = UBound(pvarArray, plngDim)
    Select Case plngDim
        Case 1
            If plngRight = 0 Then
                plngLeft = LBound(pvarArray, 2)
                plngRight = UBound(pvarArray, 2)
            End If
            lngFirst = plngLeft
            lngLast = plngRight
            varMid = pvarArray(plngCol, (plngLeft + plngRight) \ 2)
            Do
                Do While pvarArray(plngCol, lngFirst) < varMid And lngFirst < plngRight
                    lngFirst = lngFirst + 1
                Loop
                Do While varMid < pvarArray(plngCol, lngLast) And lngLast > plngLeft
                    lngLast = lngLast - 1
                Loop
                If lngFirst <= lngLast Then
                    For c = cMin To cMax
                        varSwap = pvarArray(c, lngFirst)
                        pvarArray(c, lngFirst) = pvarArray(c, lngLast)
                        pvarArray(c, lngLast) = varSwap
                    Next
                    lngFirst = lngFirst + 1
                    lngLast = lngLast - 1
                End If
            Loop Until lngFirst > lngLast
            If plngLeft < lngLast Then CBA_Sort2DArray pvarArray, plngDim, plngCol, plngLeft, lngLast
            If lngFirst < plngRight Then CBA_Sort2DArray pvarArray, plngDim, plngCol, lngFirst, plngRight
        Case 2
            If plngRight = 0 Then
                plngLeft = LBound(pvarArray, 1)
                plngRight = UBound(pvarArray, 1)
            End If
            lngFirst = plngLeft
            lngLast = plngRight
            varMid = pvarArray((plngLeft + plngRight) \ 2, plngCol)
            Do
                Do While pvarArray(lngFirst, plngCol) < varMid And lngFirst < plngRight
                    lngFirst = lngFirst + 1
                Loop
                Do While varMid < pvarArray(lngLast, plngCol) And lngLast > plngLeft
                    lngLast = lngLast - 1
                Loop
                If lngFirst <= lngLast Then
                    For c = cMin To cMax
                        varSwap = pvarArray(lngFirst, c)
                        pvarArray(lngFirst, c) = pvarArray(lngLast, c)
                        pvarArray(lngLast, c) = varSwap
                    Next
                    lngFirst = lngFirst + 1
                    lngLast = lngLast - 1
                End If
            Loop Until lngFirst > lngLast
            If plngLeft < lngLast Then CBA_Sort2DArray pvarArray, plngDim, plngCol, plngLeft, lngLast
            If lngFirst < plngRight Then CBA_Sort2DArray pvarArray, plngDim, plngCol, lngFirst, plngRight
    End Select
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_Sort2DArray", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Function CBA_HeatMap(ByVal rng As Range) As Boolean
    
    On Error GoTo Err_Routine
    
    rng.FormatConditions.AddColorScale ColorScaleType:=3
    rng.FormatConditions(rng.FormatConditions.Count).SetFirstPriority
    rng.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With rng.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
    rng.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
    rng.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With rng.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
    rng.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
    With rng.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
    CBA_HeatMap = True
    Err.Clear
    On Error GoTo Err_Routine
Exit Function
Err_Routine:
    CBA_HeatMap = False
    Err.Clear
    On Error GoTo 0
End Function

Public Function CBA_MouseCancelled(sForm As String, sStatus As String, ByVal bStatus As Boolean) As Boolean
    Static bHasBeenRun As Boolean, astcDict As Scripting.Dictionary
    ''Dim bStatus As Boolean
    
    If bHasBeenRun = False Then
        Set astcDict = New Scripting.Dictionary
        bHasBeenRun = True
    End If
    ' Test that the form exists
    If Not astcDict.Exists(sForm) Then
        bStatus = False
        astcDict.Add sForm, bStatus
        CBA_MouseCancelled = bStatus
    ElseIf sStatus = "Get" Then
        CBA_MouseCancelled = astcDict(sForm)
    ElseIf sStatus = "Set" Then
        astcDict(sForm) = bStatus
        CBA_MouseCancelled = astcDict(sForm)
    Else
        MsgBox "status " & sStatus & " not found"
    End If
    
End Function

Public Function CBA_ProcI(Optional ByVal sProc As String = "", Optional ByVal lNumProcToKeep As Long = -1, Optional ByVal bReturnLastOnly As Boolean = False) As String
    ' Will accumulate the procedure names, up to the lNumProcToKeep
    ' I.e. s-Proc1~f-Proc2 etc
    Dim aP() As String, lIdx As Long, lIdxTo As Long
    Static stcCurrProcs As String, stcLastProc As String
    Const sSep As String = "~"
    If lNumProcToKeep < 0 Then stcCurrProcs = "": If sProc = "" Then CBA_ProcI = "": Exit Function
    ' Append the new procedure, if it differs from the last
    If stcLastProc <> sProc Then
        If stcCurrProcs > "" Then
            If sProc > "" Then stcCurrProcs = stcCurrProcs & sSep & sProc
        Else
            'If lNumProcToKeep > -1 Then
            stcCurrProcs = sProc
        End If
    End If
    stcLastProc = sProc
    ' See if the number of procs needs to be shortened
    If lNumProcToKeep > 0 Then
        aP = Split(stcCurrProcs, sSep)
        If lNumProcToKeep <= UBound(aP, 1) Then
            stcCurrProcs = "": lIdxTo = UBound(aP, 1) - lNumProcToKeep + 1
            For lIdx = UBound(aP, 1) To lIdxTo Step -1
                stcCurrProcs = IIf(lIdx > lIdxTo, sSep, "") & aP(lIdx) & stcCurrProcs
            Next
        End If
    End If
    ' If returning all or just the last
    If Not bReturnLastOnly Or InStr(1, stcCurrProcs, "~") = 0 Then
        CBA_ProcI = stcCurrProcs
    Else
        aP = Split(stcCurrProcs, sSep)
        CBA_ProcI = aP(UBound(aP, 1))
    End If
End Function

Public Sub g_DelNameRngs(Optional sOnly As String = "RR")
    'PURPOSE: Delete all Named Ranges in the ActiveWorkbook (Print Areas optional)
    'SOURCE: www.TheSpreadsheetGuru.com
    Dim nm As Name
    Dim DeleteCount As Long
    'Error Handler in case Delete Function Errors out
    On Error GoTo Skip
    'Loop through each name and delete
    For Each nm In ActiveWorkbook.Names
        If Right(nm.Name, 10) = "Print_Area" Then GoTo Skip
        If Left(nm.Name, Len(sOnly)) <> sOnly Then GoTo Skip
        On Error GoTo Skip
        'Delete Named Range
        nm.Delete
        DeleteCount = DeleteCount + 1
Skip:
        'Reset Error Handler
        On Error GoTo 0
    Next
End Sub

Public Function g_Fmt_2_IDs(ByVal ID1, ByVal ID2 As Long, Optional ByVal Num_Of_Zeros2 As Long = 2, Optional ByVal Num_Of_Zeros1 As Long = 0) As String
    ' Will bring back a formatted string - Note if the string is longer than the number of zeros to format against, it will put the number in whole...
        ' i.e. if the number of zeros is 1 but the ID2 number is 22, then the back number will be 22
        ' i.e. ?g_Fmt_2_IDs(1, 22, 1) will return "122" where ?g_Fmt_2_IDs(1, 2, 1) will return "12"
    Dim sVar As String, lNum As Long, sZeros1 As String, sZeros2 As String, lNew As Long
    If Num_Of_Zeros2 < 1 Then Num_Of_Zeros2 = 1
    For lNum = 1 To Num_Of_Zeros2
        sZeros2 = sZeros2 & "0"
    Next
    If Num_Of_Zeros1 < 1 Then Num_Of_Zeros1 = 1
    For lNum = 1 To Num_Of_Zeros1
        sZeros1 = sZeros1 & "0"
    Next
    lNew = Val(ID1)
    sVar = CStr(Format(lNew, sZeros1) & Format(ID2, sZeros2))
    g_Fmt_2_IDs = sVar

End Function

Public Function g_Get_Mid_Fmt(ByVal lID As Long, ByVal Str_Pos_From_Back As Long, ByVal Num_Of_Pos As Long) As Long
    ' Will bring back a formatted string - Note if the number of zeros to format against is longer than the string(lID), it will put in the whole number ...
        ' i.e. if the Num_Of_Pos is 1 but the lID number is 22, then the back number will be 22
        ' i.e. ?g_Get_Mid_Fmt(1, 22, 1) will return "122" where ?g_Get_Mid_Fmt(1, 2, 1) will return "12"
    Dim sVar As String, lNum As Long, sZeros As String
    sZeros = "0000000000"
    sVar = Format(lID, sZeros)
    sVar = Mid(sVar, 11 - Str_Pos_From_Back, Num_Of_Pos)
    
    g_Get_Mid_Fmt = Val(sVar)

End Function

Public Sub g_EraseAry(ArrayName As Variant)
    On Error Resume Next
    Erase ArrayName
End Sub

Public Function g_SplitSQLString(ByVal strInput As String) As String
    ' Will split an SQL string that is too long - it goes over the length of a SQLS line and produces an error (it is split up by + chars
    ' NOTE : INCLUDE NO "'" CHARS AS THEY WILL BE ADDED
    Const STRLEN As Long = 999
    
    Dim sSQL As String, sNewStr As String, sSep As String, a As Long, b As Long
    sNewStr = "": a = 0: sSep = ""
    Do While Len(sNewStr) < Len(strInput)
        a = a + 1
        b = IIf(STRLEN > (Len(strInput) - Len(sNewStr)), Len(strInput), STRLEN)
        sNewStr = sNewStr & Mid(strInput, a, b)
        sSQL = sSQL & sSep & Mid(strInput, a, b)
        sSep = "'" & Chr(10) & "+ '"
        a = Len(sNewStr)
    Loop
    g_SplitSQLString = sSQL
End Function

Public Function g_SaveFileTo(sFileName) As String
    ' Displays the save file dialog
    g_SaveFileTo = Application.GetSaveAsFilename(FileFilter:= _
             "Excel Files (*.xlsx), *.xlsx", Title:=sFileName, _
            InitialFileName:="")
End Function


Public Function g_SecsElapsed(Optional sInit As String = "") As Long
    ' Call once with "Init" and again with "" to return the seconds between the calls
    Static dtDuration As Date
    If sInit = "Init" Then
        dtDuration = Now()
        g_SecsElapsed = 0
    Else
        g_SecsElapsed = DateDiff("s", dtDuration, Now())
    End If
End Function


Public Sub g_WasteTime(Finish As Long)
    Dim NowTick As Long
    Dim EndTick As Long
    ' This routine will hold the focus of the CPU until x seconds have elapsed-not sure if it will be much use
    
'    Call the GetTickCount API to decide how many mili-seconds have elapsed since the computer started
    EndTick = GetTickCount() + (Finish * 1000)
     
    Do
        NowTick = GetTickCount
        DoEvents
    Loop Until NowTick >= EndTick
End Sub
Public Function g_isArrayWData(ByRef arr As Variant) As Boolean
    On Error Resume Next
    If UBound(arr, 1) < 0 Then
        g_isArrayWData = False
    Else
        g_isArrayWData = True
    End If
End Function


Public Function g_IsNull(varInput As Variant, bln_str_lng_int As String, _
                         Optional varDefault As Variant) As Variant
    ' Is a more comprehensive version of the NZ() function
    On Error GoTo Err_Routine
Dim ErrNo As Integer
    Select Case LCase(bln_str_lng_int)
        Case "bln"
            g_IsNull = CBool(varInput)
        Case "str"
            g_IsNull = CStr(varInput)
        Case "lng"
            g_IsNull = CLng(varInput)
        Case "sng"
            g_IsNull = CSng(varInput)
        Case "int"
            g_IsNull = CInt(varInput)
        Case "ymd"
            g_IsNull = CLng(Year(varInput) & Right("0" & Month(varInput), 2) & Right("0" & Day(varInput), 2))
        Case "dmy"
            g_IsNull = CLng(Day(varInput) & Right("0" & Month(varInput), 2) & Year(varInput))
        Case "dte"
            g_IsNull = CDate(varInput)
        Case "cur"
            g_IsNull = CCur(varInput)
    End Select
            
            
Exit_IsNull:
            
Exit Function
Err_Routine:
ErrNo = ErrNo + 1
If ErrNo = 1 Then
    Select Case LCase(bln_str_lng_int)
        Case "bln"
            If IsMissing(varDefault) Then varDefault = False
            g_IsNull = CBool(varDefault)
        Case "str"
            If IsMissing(varDefault) Then varDefault = ""
            g_IsNull = CStr(varDefault)
        Case "lng"
            If IsMissing(varDefault) Then varDefault = 0
            g_IsNull = CLng(varDefault)
        Case "sng"
            If IsMissing(varDefault) Then varDefault = 0
            g_IsNull = CSng(varDefault)
        Case "int"
            If IsMissing(varDefault) Then varDefault = 0
            g_IsNull = CInt(varDefault)
        Case "ymd"
            If IsMissing(varDefault) Then varDefault = 19400101
            g_IsNull = varDefault
        Case "dmy"
            If IsMissing(varDefault) Then varDefault = 1011940
            g_IsNull = varDefault
        Case "dte"
            If IsMissing(varDefault) Then varDefault = CDate("01/01/1940")
            g_IsNull = varDefault
        Case "cur"
            If IsMissing(varDefault) Then varDefault = 0
            g_IsNull = CLng(varDefault)
    End Select
Else
    Select Case LCase(bln_str_lng_int)
        Case "bln"
            g_IsNull = False
        Case "str"
            g_IsNull = ""
        Case "lng", "sng"
            g_IsNull = 0
        Case "int"
            g_IsNull = 0
    End Select
End If
    Resume Exit_IsNull
End Function

Public Function NZ(ByVal varInput, Optional ByVal varDflt = 0) As Variant   ' #RW made the varDeflt into an optional variable that defaults to a 0
    On Error GoTo Err_Routine
    If IsNull(varInput) = True Then
        NZ = varDflt
    ElseIf g_Empty(CStr(varInput)) > 0 Then
        NZ = varInput
    Else
        NZ = varDflt
    End If
    Exit Function
Err_Routine:
    NZ = varDflt
End Function

Public Function g_PosForm(ByVal lTop As Long, ByVal lWidth As Long, ByVal lLeft As Long, Optional Top_Left_Width As String = "", Optional bReset As Boolean = False) As Long
        ' This routine will return the best lefterly position for an overlaying form
        ' After the initial Reset of the values, the lLeft value will be an offset to the centre position
        Static stcTop As Long, stcWidth As Long, stcLeft As Long
        Dim lTTop As Long, lTWidth As Long, lTLeft As Long
        If stcTop = 0 Or bReset Then stcTop = lTop
        If stcWidth = 0 Or bReset Then stcWidth = lWidth
        If stcLeft = 0 Or bReset Then stcLeft = lLeft
        
        Select Case Top_Left_Width
        Case "Left"
            lTLeft = stcLeft + ((stcWidth - lWidth) / 2) + lLeft
            If lTLeft < 1 Then lTLeft = 1
            g_PosForm = lTLeft
        Case "Width"
            g_PosForm = stcWidth
        Case "Top"
            g_PosForm = stcTop
        Case Else
            g_PosForm = 0
        End Select
        
End Function

Function g_getConfig(SCode As String, sDatabase As String, _
                        Optional Disp_sMsg_If_Not_Found As Boolean = True, _
                        Optional Return_If_Not_Found As String = "") As String
    ' Will get the required parameter from the Config table

''x = g_DLookup("CFG_Value", "A0_Config", "CFG_Prefix='" & sCode & "'", "", sDatabase, Return_If_Not_Found)

    Dim sReturn As String
    
    sReturn = g_DLookup("CFG_Value", "A0_Config", "CFG_Prefix='" & SCode & "'", "", sDatabase, Return_If_Not_Found)
    g_getConfig = Return_If_Not_Found
    If sReturn = "" Then
        If Disp_sMsg_If_Not_Found Then
            CBA_Error = SCode & " wasn't found in the Configuration table - Please tell your admin about this error."
            MsgBox CBA_Error, vbExclamation, "Config Error"
            Call g_FileWrite(g_GetDB("Gen", True), CBA_Error & vbCrLf & sDatabase, , , , True)
        End If
        Exit Function
    End If
    g_getConfig = sReturn
    
End Function

Public Function g_UpdConfig(SCode As String, sDatabase As String, Optional Update_Value As String = "") As String
    ' Will update or add the parameter to the Config table - (Desc not added)
    Dim sReturn As String, sSQL As String, RS As ADODB.Recordset, CN As ADODB.Connection, bAdd_Upd As Boolean
    On Error GoTo Err_Routine
        
    bAdd_Upd = False
    sReturn = g_DLookup("CFG_Value", "A0_Config", "CFG_Prefix='" & SCode & "'", "", sDatabase, "@X@X@")
    If sReturn = "@X@X@" Then
        bAdd_Upd = True
        g_UpdConfig = Update_Value
    Else
        g_UpdConfig = Update_Value
    End If
    If bAdd_Upd = True Then
        sSQL = "INSERT INTO A0_Config (CFG_Prefix, CFG_Value, CFG_Desc) " & _
               "VALUES ( '" & SCode & "','" & Update_Value & "','Needs Desc' )"
    Else
        sSQL = "UPDATE A0_Config SET CFG_Value='" & Update_Value & "' WHERE CFG_Prefix='" & SCode & "' ;"
    End If
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & sDatabase & ";"
    RS.Open sSQL, CN

Exit_Routine:
    Set RS = Nothing
    Set CN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-g_UpdConfig", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function

'Public Function g_Get_OADocType(ByVal Doc_ID As Long, Optional DocType As String = "OADocType") As String
'    ' Will get the 'Overall Document Type' (at the moment C ot T) or the 'Document Type Desc' from the appropriate UT_DB table
'    Dim sReturn As String, sSQL As String, RS As ADODB.Recordset, CN As ADODB.Connection, lIdx As Long
'    Static arrDB(), bActive As Boolean
'
'    On Error GoTo Err_Routine
'    If bActive = False Then
'        Set CN = New ADODB.Connection
'        Set RS = New ADODB.Recordset
'        CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("UDT", , , , , "Test") & ";"
'        sSQL = "SELECT SH_ID,SH_Cat_or_Ten, SH_Desc FROM C1_Seg_Template_Hdrs"
'        RS.Open sSQL, CN
'        arrDB = RS.GetRows
'        bActive = True
'    End If
'    ' Run through the Doc Types to find the type required
'    For lIdx = 0 To UBound(arrDB, 2)
'        If Doc_ID = arrDB(0, lIdx) Then
'            If DocType = "OADocType" Then
'                g_Get_OADocType = arrDB(1, lIdx)
'            Else
'                g_Get_OADocType = arrDB(2, lIdx)
'            End If
'            Exit For
'        End If
'    Next
'
'Exit_Routine:
'    On Error Resume Next
'    Set RS = Nothing
'    Set CN = Nothing
'    Exit Function
'
'Err_Routine:
'    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-g_Get_OADocType", 3)
'    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
'    Debug.Print CBA_Error
'    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
'    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
'    GoTo Exit_Routine
'    Resume Next
'
'End Function
'
Public Function g_IsLoaded(FormName As String) As Boolean
Dim frm As Object
For Each frm In VBA.UserForms
    If frm.Name = FormName Then
        g_IsLoaded = True
        Exit Function
    End If
Next frm
g_IsLoaded = False
End Function

Public Function g_IsDate(ByVal strDate As Variant, Optional bTest_For_Days As Boolean = False) As Boolean
On Error GoTo Err_Routine
    ' As there is a flaw in Access, IsDate (i.e. a time is a date), this routine will take the most obvious flaws, and deliver back a false.
    ' Remember to set bTest_For_Days = True if you are entering a date from a field that has formatting - g_FixDate may be better if the date is fuzzy
    g_IsDate = True
    If bTest_For_Days Then strDate = g_StripDateDDsMMs(strDate)
    If Len(strDate) < 6 Or (Len(strDate) < 13 And InStr(7, strDate & Space(12), ":") > 0) Or InStrRev(strDate, ":", 8) Then strDate = " "
    g_IsDate = isDate(strDate)
Exit_Routine:
    Exit Function
    
Err_Routine:
    g_IsDate = False
    Resume Exit_Routine
    
End Function

Public Function g_FixDate(ByVal DateInput, Optional strReturnFormat As String = CBA_DMY) As Variant
    ' This procedure will fix a date or return an input formatted date (note can be partial date or must be complete date)
    ' It will use fuzzy logic i.e. enter 13 and it will retun a date 13/this mm/this year
    '                              enter 13/5 and it will retun a date 13/05/this year
    '                              it will remove most prior formatting i.e. Wed or We and won't count it in the IsDate
    ' This will only recognise dates in Aus Format - i.e. yyyy/mm/dd or dd/mm/yyyy - it will accept diff separators of '.' or '-'
    
    Dim strDate1 As String, strDate2 As String, strADate As String
    Const cstDAY1 As String = "MON,TUE,WED,THU,FRI,SAT,SUN"
    Const cstDAY2 As String = "MO,TU,WE,TH,FR,SA,SU"
    ' If the date goes down to mili-seconds, take them out as MS doesn't recognise a date like this (IsDate will be False)
    If Len(DateInput) > 18 Then
        If Mid(DateInput, Len(DateInput) - 3, 1) = "." Then DateInput = g_Left(DateInput, 4)
    End If
    ' If the user is using '.' or '-' instead of '/'
    If InStr(1, DateInput, "/") = 0 Then
        If InStr(1, DateInput, "-") > 0 Then
            DateInput = Replace(DateInput, "-", "/")
        ElseIf InStr(1, DateInput, ".") > 0 Then
            DateInput = Replace(DateInput, ".", "/")
        End If
    End If
    strDate2 = CStr(DateInput & "")
    ' If there are named days in there, remove them
    If Len(strDate2) > 4 Then
        If InStr(cstDAY1, UCase(Left(strDate2, 3))) > 0 Then strDate2 = g_Right(strDate2, 3)
        If InStr(cstDAY2, UCase(Left(strDate2, 2))) > 0 Then strDate2 = g_Right(strDate2, 2)
    End If
    strADate = g_KeepReplace(strDate2, "Number", "*", "/")
    If Right(strADate, 1) = "/" Then strADate = g_Left(strADate, 1)
    strADate = "*" & strADate & "*"
    Do While InStr(1, strADate, "***/") = 1
        strADate = g_Right(strADate, 1)
    Loop
    Do While InStr(3, strADate, "/***") = 3
        strADate = g_Left(strADate, 1)
    Loop
    Do While InStr(1, strADate, "***") = 1
        strADate = g_Right(strADate, 1)
    Loop
    If strADate = "**" Then
        If Right(strDate2, 1) <> "/" Then strDate2 = strDate2 & "/"
        strDate1 = strDate2 & Month(Date) & "/" & Year(Date)
        If g_IsDate(strDate1) Then
            GoSub gsFormat
        End If
    ElseIf strADate = "**/**" Then
        If Right(strDate2, 1) <> "/" Then strDate2 = strDate2 & "/"
        strDate1 = strDate2 & Year(Date)
        If g_IsDate(strDate1) Then
            GoSub gsFormat
        End If
    Else
        strDate1 = strDate2
        GoSub gsFormat
    End If
    Exit Function
    
gsFormat: ' This part of the routine will do the final reformatting...
    If Left(strReturnFormat, Len("dd dd/mm/yyyy")) = "dd dd/mm/yyyy" Then
        g_FixDate = Left(Format(strDate2, "ddd"), 2) & " " & Format(strDate2, g_Right(strReturnFormat, 3))
    Else
        g_FixDate = Format(strDate1, strReturnFormat)
    End If
    If InStr(1, strReturnFormat, "/") > 0 Or InStr(1, strReturnFormat, "-") > 0 Then
        If g_IsDate(g_FixDate, True) = False Then g_FixDate = ""
    End If
    Return
End Function

Public Function g_MkDir(sPath As String) As String
    ' Will make a directory and produce no error if it already exists - note doesn't cater for the drive not being on-line
    On Error GoTo Err_Routine
    MkDir (sPath)
    g_MkDir = sPath
Exit_Routine:
    Exit Function
Err_Routine:
    GoTo Exit_Routine
End Function

Public Function g_RevDate(ByVal DateInput, Optional strInputFormat As String = "yyyymmyy", Optional strReturnFormat As String = CBA_DMY) As Variant
    ' This procedure will reverse a date, either (yyymmdd[hhnn] to dd/mm/yyyy hh:nn) or vise-versa
    Dim strDate1 As String, lYY As Long, lMM As Long, lDD As Long
    If g_IsDate(DateInput, True) = True Then
        DateInput = g_FixDate(DateInput)
        lYY = Year(DateInput)
        lMM = Month(DateInput)
        lDD = Day(DateInput)
    ElseIf Len(DateInput) = 8 Then
        lYY = Left(DateInput, 4)
        lMM = Mid(DateInput, 5, 2)
        lDD = Right(DateInput, 2)
    ElseIf Len(DateInput) = 6 Then
        lYY = Val("20" & Left(DateInput, 2))
        lMM = Mid(DateInput, 3, 2)
        lDD = Right(DateInput, 2)
    End If
    strDate1 = lDD & "/" & lMM & "/" & lYY
    g_RevDate = g_FixDate(strDate1, strReturnFormat)
    Exit Function

End Function

Public Function g_SetupIP(Optional ByVal sForm As String, Optional ByVal lLevel As Long = 0, Optional ByVal bSetas As Boolean = False, Optional ByVal bReset As Boolean = False) As Boolean
    ' This routine will hold setup at various levels, so they don't affect each other i.e. 1st level init, 2nd save or reset, 3rd add etc
    ' I.e. the main form (Level 1 setup) could be in setup mode
    ' I.e. another proc on the form (Level 2 setup) could be in setup mode
    ' Any being true will mean that g_Setup is true (some sort of setup is being applied and forms don't do their usual events)

    Static SU() As Variant, lMaxIdx As Long, bfound As Boolean, sSavedForm As String
    Dim lIdx As Long
    If sForm = "" Then sForm = sSavedForm
    If sForm = "" Then Exit Function
    bfound = False: lIdx = 0
    sSavedForm = sForm
    ' If the first...
    If lMaxIdx = 0 Then
        lMaxIdx = lMaxIdx + 1
        ReDim SU(0 To 1, 0 To lIdx)
        SU(0, lIdx) = sForm
        SU(1, lIdx) = 0
    End If
    ' Find the item to update
    For lIdx = 0 To lMaxIdx - 1
        If SU(0, lIdx) = sForm Then
            bfound = True
            Exit For
        End If
    Next
    ' If not found ...
    If bfound = False Then
        lIdx = lMaxIdx
        lMaxIdx = lMaxIdx + 1
        ReDim Preserve SU(0 To 1, 0 To lIdx)
        SU(0, lIdx) = sForm
        SU(1, lIdx) = 1
    End If
    ' If reset is specified on the setup parms
    If bReset = True Then
        SU(1, lIdx) = 0
    End If
ReStart:
    If lLevel = 0 Then
        If SU(1, lIdx) > 0 Then
            g_SetupIP = True
        Else
            g_SetupIP = False
        End If
    ElseIf lLevel > 0 Then
        SU(1, lIdx) = SU(1, lIdx) + IIf(bSetas, 1, -1)
        If SU(1, lIdx) < 0 Then SU(1, lIdx) = 0
    End If
    ''DoEvents
''    If lLevel > 0 Then
''        Debug.Print lLevel & IIf(bSetas, "+", "-") & ";";
''    End If
    Exit Function
End Function

Public Function g_StripDateDDsMMs(ByVal DateInput) As Variant
    ' This procedure will strip any named days from the date (they make a date invalid regardless whether the actual date is invalid or not)
    ' Eventually will strip named Months from it too...
    Dim strDate2 As String, aDt() As String, lIdx As Long
    Const cstDAYS1 As String = "MON,TUE,WED,THU,FRI,SAT,SUN,MO,TU,WE,TH,FR,SA,SU"
    Const cstDAYS2 As String = "MONDAY,TUESDAY,WEDNESDAY,THURSDAY,FRIDAY,SATURDAY,SUNDAY"
    strDate2 = UCase(Trim(CStr(DateInput & "")))
    If Len(strDate2) > 7 Then
        aDt() = Split(cstDAYS2, ",")
        For lIdx = 0 To UBound(aDt, 1)
            If InStr(1, strDate2, aDt(lIdx)) > 0 Then
                strDate2 = Replace(strDate2, aDt(lIdx), "")
                GoTo GTSkipDDs
            End If
        Next
    End If
    If Len(strDate2) > 4 Then
        aDt() = Split(cstDAYS1, ",")
        For lIdx = 0 To UBound(aDt, 1)
            If InStr(1, strDate2, aDt(lIdx)) > 0 Then
                strDate2 = Replace(strDate2, aDt(lIdx), "")
                GoTo GTSkipDDs
            End If
        Next
    End If
GTSkipDDs:
    ' Transfer the date
    g_StripDateDDsMMs = strDate2


End Function

Public Function g_KeepReplace(ByVal varInput As Variant, _
                          AlphaN_Alpha_Number As String, _
                          ReplaceWith As String, _
                          Optional ByVal Except_AllPunct As String = " ") _
                          As String
    ' This procedure will return the parts of a string required
    '                   i.e g_KeepReplace("Bob 0980%/)", "AlphaN", "*", " %") will return "Bob 0980%**"
    '                   i.e g_KeepReplace("Bob 0980%/)", "Alpha" , "*", " %") will return "Bob ****%**"
    '                   i.e g_KeepReplace("Bob 0980%/)", "Number", "*", " %") will return "*** 0980%**"
    ' If ReplaceWith=""     g_KeepReplace("Bob 0980%/)", "Number", "" , " %") will return " 0980%"
    ' If Except_AllPunct="" g_KeepReplace("Bob 0980%/)", "Number", "" , "") will return "0980" ' Note both ' ' and '%' were excepted characters
    
    Dim Idx1 As Long
    Dim strReplace As String
    Dim strTemp As String
    Dim strOutput As String
    Dim ExceptAll As Boolean
    Dim ExceptAllT As Boolean
    
    If InStr(1, Except_AllPunct, "AllPunct") > 0 Then
        'ExceptAll = True
        Except_AllPunct = Replace(Except_AllPunct, "<>,.?/':;|\}]{[+=_-*&^%$#@!~`", "", 1, 1)
    End If
    
    If g_Empty(varInput) = 0 Then g_KeepReplace = ReplaceWith: Exit Function
      
    strReplace = CStr(varInput)
    strOutput = ""
      
    If Except_AllPunct = " " Then Except_AllPunct = ReplaceWith
    
    For Idx1 = 1 To Len(strReplace)
        strTemp = Mid(strReplace, Idx1, 1)
        If (UCase(strTemp) >= "A" And UCase(strTemp) <= "Z") Then
            If (Mid(AlphaN_Alpha_Number, 6) = "N" Or AlphaN_Alpha_Number = "Number") Or InStr(1, Except_AllPunct, strTemp) > 0 Then
                strOutput = strOutput & strTemp
            Else
                ExceptAllT = False
                GoSub ReplaceIt
            End If
        ElseIf (strTemp >= "0" And strTemp <= "9") Then
            If Left(AlphaN_Alpha_Number, 5) = "Alpha" Or InStr(1, Except_AllPunct, strTemp) > 0 Then
                strOutput = strOutput & strTemp
            Else
                ExceptAllT = False
                GoSub ReplaceIt
            End If
        Else
            ExceptAllT = ExceptAll
            GoSub ReplaceIt
        End If
    Next
    
    g_KeepReplace = strOutput
    
    Exit Function
    
    ' This Gosub will test to see if the exceptions disqualify the char from being omitted
ReplaceIt:
    
    If ExceptAllT = True Then
        strOutput = strOutput & strTemp
    ElseIf InStr(1, Except_AllPunct, strTemp, vbBinaryCompare) = 0 Then
        strOutput = strOutput & ReplaceWith
    Else
        strOutput = strOutput & strTemp
    End If
    
    Return

End Function

Public Function g_GetKeyType(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer, Optional sDefault As String = "Num") As String
    ' Will pick the numeric type actions - Used in KeyUp or Down events, if there is no AfterUpdate key etc...
    g_GetKeyType = ""
    If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then                         ' Is a number
        g_GetKeyType = "Num"
    ElseIf KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Then         ' Is a number
        g_GetKeyType = "Num"
    ElseIf KeyCode = vbKeyDecimal Or KeyCode = 190 Then                     ' ='.'-Count as a number
        g_GetKeyType = "Num"
    ElseIf KeyCode = vbKeyLButton Or KeyCode = vbKeyBack Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyNumlock Or _
           KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyDelete Then
        g_GetKeyType = sDefault                                             ' Other keys that shouldn't be zeroed
    ElseIf KeyCode = 9 Or KeyCode = 13 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        g_GetKeyType = "Exit"                                               ' Tab or Enter...
    End If
''    If g_GetKeyType = "Num" Then
''       'Debug.Print "n";
''    ElseIf g_GetKeyType = "Exit" Then
''       'Debug.Print "e";
''    Else
''       'Debug.Print "-"
''    End If
        
End Function


Public Function g_Left(ByVal InPutString As String, Chars_2_Take_Off_Back As Integer) As String
    Dim LenInPutString
    ' Will return only the the `Left Len(InputString) - Chars_2_Take_Off_Back` chars (moded for errors)
    ' eg ?g_left("kkjdshk",9) returns "", g_left("kkjdshk",2) returns "kkjds"
    g_Left = ""
    InPutString = InPutString & ""
    LenInPutString = Len(InPutString & "")
    LenInPutString = LenInPutString - Chars_2_Take_Off_Back
    If LenInPutString < 0 Then Exit Function
    g_Left = Left(InPutString, LenInPutString)
End Function

Public Function g_Right(ByVal InPutString As String, Chars_2_Take_Off_Front As Integer) As String
    Dim LenInPutString
    ' Will return only the the `Right Len(InputString) - Chars_2_Take_Off_Front` chars (moded for errors)
    ' eg ?g_Right("kkjdshk",9) returns "", g_Right("kkjdshk",2) returns "jdshk"
    g_Right = ""
    InPutString = InPutString & ""
    LenInPutString = Len(InPutString & "")
    LenInPutString = LenInPutString - Chars_2_Take_Off_Front
    If LenInPutString < 0 Then Exit Function
    g_Right = Right(InPutString, LenInPutString)
End Function

Public Function g_Empty(ByVal strTest As Variant, _
                        Optional IncBlanks As Boolean = False) As Long
    ' This function will give the length of the string minus any
    ' leading/trailing blanks if IncBlanks = False, else the full length of the string
    ' ie - "" or " " = 0
    '      "23de   " = 4,
    
    strTest = strTest & " "
    If IncBlanks = False Then
       If Len(Trim(strTest)) = 0 Then
        g_Empty = 0
       Else
        g_Empty = Len(Trim(strTest))
       End If
    Else
       g_Empty = Len(strTest) - 1
    End If
End Function

Public Function g_Get_Dev_Sts(DevUsers_DevIP As String, Optional sDevExcept As String = "xxxx") As String
    Static sReturn As String, User_Ext As String, sDevIP As String
    Dim aBup() As String, aBup2() As String, lIdx As Long, sSep As String, sName1 As String, sName2 As String, sExt As String
    ' Will return the list of maintainers of the system, or whether the user is dev staff
                                ' DevUsers                                  DevIP                   Init will redo/reset the tests
    If sDevIP = "" Or DevUsers_DevIP = "Init" Then
        sReturn = g_getConfig("DevUsers", CBA_BSA & CBA_GEN_DB, False, "N")
        If sReturn = "N" Then sReturn = g_UpdConfig("DevUsers", CBA_BSA & CBA_GEN_DB, "Pearce, Tom:9218;White, Robert:4289;Baines, Stuart:9033;")
        aBup = Split(sReturn, ";"): sDevIP = "N": User_Ext = "": sSep = ""
        For lIdx = 0 To UBound(aBup, 1)
            If aBup(lIdx) > "" Then
            If InStr(1, aBup(lIdx), CBA_User) > 0 Then
                sDevIP = "Y"
                GoTo SkipDev
            End If
SkipDev:
            aBup2 = Split(g_KeepReplace(aBup(lIdx), "Number", "", ","), ",")
            sName1 = aBup2(1) & " "
            sName2 = aBup2(0) & " "
            sExt = g_KeepReplace(aBup(lIdx), "Alpha", "", "")
            If sExt > "" Then sExt = "Ext:" & sExt
                User_Ext = User_Ext & sSep & sName1 & sName2 & sExt
                sSep = " or "
            End If
        Next
    End If
    ' Set whether vthe user is Development personel or not
    If DevUsers_DevIP = "DevUsers" Or DevUsers_DevIP = "Init" Then
        g_Get_Dev_Sts = User_Ext
    Else
        g_Get_Dev_Sts = sDevIP
    End If
    ' Do the (testing) excepts
    If g_Get_Dev_Sts = "Y" Then
        If InStr(1, sDevExcept, CBA_User) > 0 Then g_Get_Dev_Sts = "N"
    End If
End Function

Public Function g_WordCount(ByVal sText As String) As Long
    ' Will cvount the number of words in a string
    Dim arr() As String, sTxt As String, bKRDone As Boolean
    g_WordCount = 0
    Do While sTxt <> sText
        sTxt = sText
        If bKRDone = False Then
            sText = Replace(sText, vbCrLf, ",")
            sText = Replace(sText, vbCr, ",")
            sText = Replace(sText, vbLf, ",")
            sText = g_KeepReplace(sText, "AlphaN", "", ", ")
            sText = g_KeepReplace(sText, "Number", "l", ", ")
            sText = Replace(sText, " ", ",")
            bKRDone = True
        End If
        sText = Replace(sText, ",,", ",")
    Loop
    If Left(sText, 1) = "," Then sText = g_Right(sText, 1)
    If Right(sText, 1) = "," Then sText = g_Left(sText, 1)
    If sText > "" Then
        arr = Split(sText, ",")
        g_WordCount = UBound(arr, 1) + 1
    End If
    
End Function

Public Function g_Extract_Path_File(ByVal InString As String, Optional Path_File_Ext_FileExt As String = "File.Ext") As String
    On Error GoTo Err_Routine
    ' This Proc will return either the Path, FileExt the File Ext
    
    Const StdExtLen = 4
    Dim Idx1 As Integer
    Dim iExtLen As Integer
    Dim sPath As String, sFile As String, sExt As String
    Dim iPathLen As Integer
    g_Extract_Path_File = ""
    iPathLen = Len(InString)
    If iPathLen < 2 Then Exit Function
    sFile = Mid(InString, InStrRev(InString, "\") + 1, iPathLen)
    sPath = Replace(InString, sFile, "")
    sExt = Mid(InString, InStrRev(InString, ".") + 1, iPathLen)
    ' Path
    If Path_File_Ext_FileExt = "Path" Then       ' Returns 'Path' only
        g_Extract_Path_File = sPath
    ElseIf Path_File_Ext_FileExt = "File.Ext" Then ' Returns 'File.Ext'
        g_Extract_Path_File = sFile
    ElseIf Path_File_Ext_FileExt = "Ext" Then     ' Returns '.Ext' only
        g_Extract_Path_File = sExt
    End If
Exit_Extract:
Exit Function
Err_Routine:
CBA_Error = Err.Number & " " & Err.Description
MsgBox CBA_Error
Resume Exit_Extract
End Function

Public Sub g_FileWrite(ByVal strFileName As String, ByVal Write_Msg As Variant, _
                   Optional KillFileFirst As Boolean = False, _
                   Optional AppendRevDate As Boolean = False, _
                   Optional AddDate2Msg As Boolean = False, _
                   Optional RevMsg As Boolean = False, _
                   Optional RevMsgLn2 As Boolean = False)
    ' Will write any msg logs - RevMsg will place the new error first and copy the rest of the file after it
    Dim strPathFileName As String, strString2Write As String, strSpaces As String, lFreeFile As Long, lFreeFile2 As Long, strSuffix As String
    Dim sFile As String, sF() As String, strRevFileName As String, strPathRevFileName As String, sText As String, lIdx As Long
    
    lFreeFile = FreeFile
    strSuffix = ".txt": lIdx = 0
    sF = Split("\" & strFileName, "\")
    sFile = sF(UBound(sF, 1))
    strFileName = Replace(strFileName, "\" & sFile, "\")
    strFileName = Replace(strFileName, "\\", "\")
    If Mid(sFile, Len(sFile) - 3, 1) = "." Then
        strSuffix = Right(sFile, 4)
        sFile = g_Left(sFile, 4)
    ElseIf Mid(sFile, Len(sFile) - 4, 1) = "." Then
        strSuffix = Right(sFile, 5)
        sFile = g_Left(strFileName, 5)
    End If
    If AppendRevDate = True Then
        sFile = Format(Date, "yyMMdd") & sFile
        ''strFileName = ""
    End If
    strFileName = sFile & strSuffix
    strRevFileName = sFile & "_Tmp" & strSuffix
    ''strFileName = strFileName & sFile
    If InStr(strFileName, "\") > 0 Then
        strPathFileName = strFileName
        strPathRevFileName = strRevFileName
    Else
        strPathFileName = CBA_BSA & CBA_GEN_LOGS & strFileName
        strPathRevFileName = CBA_BSA & CBA_GEN_LOGS & strRevFileName
    End If
    If AddDate2Msg Then Write_Msg = Format(Now, CBA_D3DMYHN) & " " & Write_Msg
    strString2Write = Write_Msg & strSpaces
    
    If RevMsg And Not KillFileFirst Then
        Open strPathFileName For Append As #lFreeFile
        Close #lFreeFile
        Call g_Rename(strPathFileName, strPathRevFileName)
    End If
    
   
    If KillFileFirst = True Then
        Open strPathFileName For Output As #lFreeFile
    Else
        Open strPathFileName For Append As #lFreeFile
    End If
        
    If RevMsg And Not KillFileFirst Then
        lFreeFile2 = FreeFile
        Open strPathRevFileName For Input As #lFreeFile2
        Do While Not EOF(lFreeFile2)
            lIdx = lIdx + 1
            If RevMsgLn2 = True And lIdx = 2 Then
                Print #lFreeFile, strString2Write
            ElseIf RevMsgLn2 = False And lIdx = 1 Then
                Print #lFreeFile, strString2Write
            End If
            Line Input #lFreeFile2, sText
            Print #lFreeFile, sText
        Loop
        Close #lFreeFile2
        Call g_KillFile(strPathRevFileName)
        If lIdx < 2 Then Print #lFreeFile, strString2Write
    Else
        Print #lFreeFile, strString2Write
    End If
    On Error Resume Next
    Close #lFreeFile

End Sub

Public Function g_Rename(ByVal strPath_File_To_Rename As String, Optional ByVal File_To As String = "") As Variant
    On Error GoTo Err_Rename
    Dim Idx1 As Integer
    ' Will rename the text file
    Idx1 = 1
    ' Kill any file that exists
    Call g_KillFile(File_To)
    g_Rename = ""
    If File_To = "" Then
        Name strPath_File_To_Rename As g_Left(strPath_File_To_Rename, 4) & CStr(Idx1) & _
                                            Right(strPath_File_To_Rename, 4)
        g_Rename = g_Left(strPath_File_To_Rename, 4) & CStr(Idx1) & Right(strPath_File_To_Rename, 4)
    Else
        Name strPath_File_To_Rename As File_To
        g_Rename = File_To
    End If
                                        
Exit_Rename:
    Exit Function
    
Err_Rename:
    CBA_Error = Err.Number & "," & Err.Description
    If Err.Number = 58 And File_To = "" Then
        If Idx1 > 200 Then Resume Exit_Rename
        Idx1 = Idx1 + 1
        Resume
    End If
    MsgBox CBA_Error
    Resume Exit_Rename
End Function

Public Function g_RtnStrBetween(ByVal sInput As String, ByVal s1stChar As String, ByVal s2ndChar As String, _
                Optional sPrefix As String = "", Optional sSuffix As String = "", Optional lStartPos As Long = 1) As String
    ' Return a BETWEEN! s1stChar and s2ndChar string i.e. not the search chars themselves
    ' e.g. ?g_RtnStrBetween(":/XYZ/","/","/") returns "XYZ"
    '?g_RtnStrBetween(":/XYZ?","/","/") returns ""
    '?g_RtnStrBetween(":/XYZ?","/","/",,"/") returns "XYZ?" - the suffix provides the last char if it doesn't exist
    Dim lPos1 As Long, lPos2 As Long, sRtnStr As String
    g_RtnStrBetween = ""
    sInput = g_Right(sInput, lStartPos - 1)
    If g_Empty(sInput, True) = 0 Then Exit Function
    If g_Empty(s1stChar, True) = 0 Then s1stChar = Mid(sInput, lStartPos, 1)
    If g_Empty(s2ndChar, True) = 0 Then s2ndChar = Right(sInput, 1)
    lPos1 = InStr(lStartPos, sInput, s1stChar)
    If lPos1 < 1 Then
        sInput = sPrefix & sInput
        lPos1 = InStr(lStartPos, sInput, s1stChar)
    End If
    If lPos1 < 1 Then Exit Function
    lPos2 = InStr(lPos1 + Len(s1stChar), sInput, s2ndChar)
    If lPos2 < lPos1 + Len(s1stChar) Then
        sInput = sInput & sSuffix
        lPos2 = InStr(lPos1 + Len(s1stChar), sInput, s2ndChar)
    End If
    If lPos2 < lPos1 + Len(s1stChar) Then Exit Function
    g_RtnStrBetween = Mid(sInput, lPos1 + Len(s1stChar), lPos2 - lPos1 - Len(s1stChar))
End Function

Function g_GetDB(ByVal sAppInput As String, Optional ByVal bGetErrFile As Boolean = False, Optional ByVal bGetAuth As Boolean = False, _
                 Optional ByVal bReTest As Boolean = False, Optional ByVal bNoTest As Boolean = False, Optional ByVal sAltDB As String = "") As String
    ' Will get the required parameter from the Config table for a TestDB or Live database
    Dim sReturn As String, sUser As String, sReturnShortTitle As String, lIdx As Long
    Static sdApps As Scripting.Dictionary, aryApp() As String, lAppIdx As Long, bAdd As Boolean, sTorL As String
    Const clAPP = 0, clTorL = 1, clAppDB = 2, clErrFile = 3, clAuth = 4, clMax = 4

    ' If sd is not existing
    If sdApps Is Nothing Then
        Set sdApps = New Scripting.Dictionary
        lAppIdx = 0 '': lAltIdx = -1: lTestIdx = -1: lLiveIdx = -1
        ReDim aryApp(0 To clMax, 0 To 0)
        aryApp(clTorL, lAppIdx) = "L"
        aryApp(clAppDB, lAppIdx) = CBA_BSA & CBA_GEN_DB
        aryApp(clErrFile, lAppIdx) = CBA_BSA & CBA_GEN_ERR
        aryApp(clAuth, lAppIdx) = "1"
        sdApps.Add "Gen", lAppIdx
        sTorL = "L"
    End If
    ' Test...
    bAdd = False
    If sdApps.Exists(sAppInput) Then
        lIdx = Val(sdApps.Item(sAppInput))
    Else
        lAppIdx = lAppIdx + 1
        sdApps.Add sAppInput, lAppIdx
        ReDim Preserve aryApp(0 To clMax, 0 To lAppIdx)
        aryApp(clAPP, lAppIdx) = sAppInput
        bAdd = True
        aryApp(clTorL, lAppIdx) = sTorL
        aryApp(clErrFile, lAppIdx) = CBA_BSA & CBA_GEN_ERR
        aryApp(clAuth, lAppIdx) = "1"
        lIdx = lAppIdx
    End If
    sReturn = "?"
    ' If this db has to follow the Test / Live value of the calling routine and database (i.e.
    If sAltDB = "Test" And sTorL = "T" Then
            sReturn = "Y"
    ElseIf sAltDB = "Test" And sTorL = "L" Then
            sReturn = "N"
    End If
    ' If a new database is being accessed
    If (bAdd = True And sAppInput <> "Gen") Then
        If (bReTest = True And sAppInput <> "Gen" And sReturn = "?") Then
            ' Check for Testing (first 'Y/N' = 'Y' will default to testing db)
            ' Want to prompt (second 'Y/N' - 'Y' will prompt admin for Testing db)
            ' Authority (third '#/N') - If put a 1 to 3 in here, the Authority will Override to that number regardless of who is entering data
            sReturn = g_getConfig(sAppInput & "Test", CBA_BSA & CBA_GEN_DB, , "NNN") & "NNN"
            aryApp(clAuth, lIdx) = Mid(sReturn, 3, 1)
            ''sReturn = "N"
            Call CBA_getUserShortTitle("", sReturnShortTitle, CBA_SetUser, sUser)
            ' If Testing (first = 'Y') and want to prompt (second = 'Y')...
            If Mid(sReturn, 2, 1) = "Y" Then
                sReturn = "N"           ' Default N
                If AST_Dev_Auth("DevIP") = "Y" And bNoTest = False Then
                    ''Debug.Print sUser
                    If MsgBox("Do you wish to use the " & sAppInput & " test database?", vbYesNo, "Testing?") = vbYes Then sReturn = "Y"
                End If
            ElseIf Mid(sReturn, 1, 1) = "Y" Then
                sReturn = "N"           ' Default N
                If g_Get_Dev_Sts("DevIP") = "Y" And bNoTest = False Then sReturn = "Y"
            Else
                sReturn = "N"
            End If
        End If
        ' Decide which db to use
        If sReturn = "Y" Then               ' If is Testing
            sTorL = "T"
            aryApp(clTorL, lIdx) = sTorL
            aryApp(clAppDB, lIdx) = g_getConfig(sAppInput & "TestDB", CBA_BSA & CBA_GEN_DB, , "N")
            aryApp(clErrFile, lIdx) = g_getConfig(sAppInput & "TestErrs", CBA_BSA & CBA_GEN_DB, , "N")
        Else                                ' If Live
            sTorL = "L"
            aryApp(clTorL, lIdx) = sTorL
            aryApp(clAppDB, lIdx) = g_getConfig(sAppInput & "LiveDB", CBA_BSA & CBA_GEN_DB, , "N")
            aryApp(clErrFile, lIdx) = g_getConfig("GenLiveErrs", CBA_BSA & CBA_GEN_DB, , "N")
        End If
    End If
    ' If a new ASYST App is being accessed
    If bReTest = True And sAppInput = "ASYST" Then
        ' Test for Authority
        If AST_Dev_Auth("DevIP") = "Y" And bNoTest = False Then
            varFldVars.lFldWidth = 300
            varFldVars.lFldHeight = 0
            varFldVars.lFrmLeft = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
            varFldVars.lFrmTop = g_PosForm(0, 0, 0, "Top")
            varFldVars.sHdg = "Enter Authority required"
            varFldVars.sSQL = "SELECT * FROM [A0_Authorities] ORDER BY AU_ID"
            varFldVars.sDB = "ASYST"
            varFldVars.bAllowNullOfField = False
            varFldVars.lCols = 2
            varFldVars.sType = "ComboBox"
            CBA_frmEntryField.Show vbModal
            If CBA_bDataChg Then
                aryApp(clAuth, lIdx) = varFldVars.sField1
            Else
                aryApp(clAuth, lIdx) = "0"
            End If
        End If
    End If
    ' Deliver back the requirement
    If bGetErrFile = True Then
        g_GetDB = aryApp(clErrFile, lIdx)
    ElseIf bGetAuth = True Then
        g_GetDB = aryApp(clAuth, lIdx)
    Else
        g_GetDB = aryApp(clAppDB, lIdx)
    End If
    Exit Function

End Function

Public Function g_GetFile(Optional ByRef sPath As String, Optional ByRef sFile As String) As String
    ' Will provide a dialog box for selection of a file
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Title = "Select the excel file to import from"
        .Filters.Add "Excel Files", "*.xls?", 1
        .AllowMultiSelect = False
        If .Show = True Then
            sPath = .SelectedItems(1)           ' Get the complete file path.
            sFile = Dir(.SelectedItems(1))      ' Get the file name.
            sPath = Replace(.SelectedItems(1), sFile, "")
        End If
    End With
    ' Return the file name
    If sFile <> "" Then
        g_GetFile = sPath & sFile
    Else
        g_GetFile = ""
    End If
End Function


Public Function g_Write_Err_Table(ByVal oErr As ErrObject, ByVal AdditMsg As String, ByVal sApp As String, ByVal sProc As String, ByVal lErrLine As Long, ByVal sTestErr As String) As Boolean
    ' This routine will write the error message to the Central Database
    ' Note sometimes may need to do a Err.Raise 513+ to pull an error that doesn't normally pull
    
    Dim CN As ADODB.Connection, RS As ADODB.Recordset, lIdx As Long, sSQL As String
    Static vErrNos(), bRanB4 As Boolean, lErrIdx As Long
    ' Flag asd a test error if dev personelle
    If g_Get_Dev_Sts("DevIP") = "Y" Then sTestErr = "Y"
    ' If the same error msg...
    If bRanB4 = False Then
        lErrIdx = 0
        GoSub GSAddIT
    Else
        For lIdx = 0 To UBound(vErrNos, 2)
            If vErrNos(0, lErrIdx) = Format(Now(), "dd/mm/yyyy hh") And vErrNos(1, lErrIdx) = oErr.Number Then
                g_Write_Err_Table = True
                Exit Function
            End If
        Next
        lErrIdx = lErrIdx + 1
        GoSub GSAddIT
    End If
    
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Gen") & ";"
    AdditMsg = Replace(oErr.Description & "-" & AdditMsg, "'", "`")
''    If CBA_All_Ver = 0 Then CBA_All_Ver = Val(g_KeepReplace(CBA_COM_Ver, "Alpha", ""))
    
    sSQL = "INSERT INTO L1_Error (Err_No, Err_Err, Err_Test, Err_App, Err_Version, Err_Proc, Err_Line, Err_User, Err_EmailSts, Err_ErrorSts)" & Chr(10)
    sSQL = sSQL & "VALUES (" & oErr.Number & ",'" & AdditMsg & "','" & sTestErr & "','" & Replace(sApp, "'", "`") & "'," & CBA_All_Ver & ",'" & Replace(sProc, "'", "`") & "'," & _
                   lErrLine & ",'" & CBA_User & "',1,1);"
    RS.Open sSQL, CN

Exit_Routine:
    On Error Resume Next
    CN.Close
    Set CN = Nothing
    Set RS = Nothing
    If g_Get_Dev_Sts("DevIP") = "Y" Then
        
        Stop
    Else
        MsgBox g_getConfig("MajorError", CBA_BSA & CBA_GEN_DB, False, "The application has experienced an error - please continue if possible - Details have been sent to Admin")
        End
    End If
    Exit Function
    
GSAddIT:
    ReDim Preserve vErrNos(0 To 1, 0 To lErrIdx)
    vErrNos(0, lErrIdx) = Format(Now(), "dd/mm/yyyy hh")
    vErrNos(1, lErrIdx) = oErr.Number
    Return
    
End Function

Public Function g_Val(vVal) As Double
    Dim sVal As String
    On Error Resume Next
    CBA_Error = ""
    g_Val = 0
    sVal = CStr(vVal)
    CBA_Error = "A"
    If Len(sVal) > 0 Then sVal = g_KeepReplace(sVal, "AlphaN", "", ".")
    g_Val = Val(sVal)

End Function

Public Function g_GetExcelCell(lRow As Long, lcol As Long, Optional WildCardApply As String = "") As String
    Const a_COLS As String = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AP,AQ,AR,AS,AT,AU,AW,AX,AY,AZ,BA,BB,BC,BD,BE,BF,BG,BH,BI,BJ,BK,BL,BM,BN,BO,BQ,BR,BS,BT,BU,BW,BX,BY,BZ"
    Dim Ary() As String
    On Error Resume Next
    Ary = Split(a_COLS, ",")
    g_GetExcelCell = WildCardApply & Ary(lcol - 1) & WildCardApply & CStr(lRow)
End Function


Public Function g_UnFmt(ByVal sInput, Optional str_lng_dbl_cur_sng_num As String = "str") As Variant
    ' Will un-format the input - i.e. if it is '$23,450' will format to '23450'
    On Error GoTo Err_Routine
    sInput = CStr(sInput & "")
    sInput = Trim(Replace(sInput, ",", ""))
    sInput = Replace(sInput, "%", "")
    sInput = Trim(Replace(sInput, "$", ""))
''    sInput = Replace(sInput, "#", "")????????
    If (str_lng_dbl_cur_sng_num <> "str" And Val(sInput) = 0) Then sInput = 0
    Select Case str_lng_dbl_cur_sng_num
    Case "num"
        g_UnFmt = Val(sInput)
    Case "lng"
        g_UnFmt = CLng(sInput)
    Case "byt"
        g_UnFmt = CByte(sInput)
    Case "dbl"
        g_UnFmt = CSng(sInput)
    Case "cur"
        g_UnFmt = CCur(sInput)
    Case "sng"
        g_UnFmt = CSng(sInput)
    Case Else
        g_UnFmt = CStr(sInput)
    End Select
    Exit Function
Err_Routine:
    Select Case str_lng_dbl_cur_sng_num
    Case "num"
        g_UnFmt = 0
    Case "lng"
        g_UnFmt = CLng(0)
    Case "byt"
        g_UnFmt = CByte(0)
    Case "dbl"
        g_UnFmt = CSng(0)
    Case "cur"
        g_UnFmt = CCur(0)
    Case "sng"
        g_UnFmt = CSng(0)
    Case Else
        g_UnFmt = CStr(sInput)
    End Select

End Function

Public Function g_IsNumeric(ByVal sInput, Optional bAllowblanks As Boolean = True) As Boolean
    ' Will check all the input to ensure it is numeric- i.e. Even if it is '$23,450%' will return true which IsNumeric may not
    ' if bAllowblanks then will take nulls or empty to be numeric too
    sInput = CStr(sInput & "")
    sInput = Trim(Replace(sInput, ",", ""))
    sInput = Replace(sInput, "%", "")
    sInput = Replace(Trim(sInput), "$", "")
    sInput = sInput & IIf(bAllowblanks And Len(sInput) = 0, "0", "")
    g_IsNumeric = IsNumeric(sInput)

End Function


Public Function g_DLookup(sField As String, sTable0rQuery As String, sWhere As String, sOrderby As String, _
                            sDatabase As String, VarIfNoReturn) As Variant
    ' Will do a DLookup on an Access table as Access does
    Dim sSQL As String, RS As ADODB.Recordset, CN As ADODB.Connection, lErrors As Long
    On Error GoTo Err_Routine
   
    g_DLookup = VarIfNoReturn ' Set to default
ReStart:
    
    sSQL = "SELECT " & sField & " AS NewFld FROM " & sTable0rQuery & " WHERE " & sWhere & IIf(sOrderby > "", " ORDER BY " & sOrderby, "") & ";"
    Set CN = New ADODB.Connection
    With CN
        .ConnectionTimeout = 100
        .CommandTimeout = 100
    End With
    
    Set RS = New ADODB.Recordset
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & sDatabase & ";"
    RS.Open sSQL, CN
    If Not RS.EOF Then
        g_DLookup = RS("NewFld")
    End If

Exit_Routine:
    On Error Resume Next
    Set RS = Nothing
    Set CN = Nothing
    Exit Function

Err_Routine:
    lErrors = lErrors + 1
    If lErrors < 5 Then Resume ReStart   '' 5 gives around 17 seconds
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-g_DLookup", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & vbCrLf & sSQL
    On Error Resume Next
    Set RS = Nothing
    Set CN = Nothing
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
''    GoTo Exit_Routine
    Resume Next

End Function

Public Function g_Regions(vRegion As Variant, Optional ByVal bCapitalise As Boolean = True) As Variant
    ' Get the opposite code i.e. Min = 501
    Dim lReg As Long, sReg As String
    Static sarr_Let() As String, sarr_Nos() As String, bGot As Boolean
    Const cREGLET As String = "Min,Der,Stp,Pre,Dan,Bre,Rgy,Jkt", _
          cREGNOS As String = "501,502,503,504,505,506,507,509"
    If bGot = False Then
        sarr_Let = Split(cREGLET, ",")
        sarr_Nos = Split(cREGNOS, ",")
        bGot = True
    End If
    ' If numeric then return the prefix
    If IsNumeric(vRegion) = True Then
        For lReg = 0 To UBound(sarr_Nos, 1)
            If vRegion = sarr_Nos(lReg) Then
                g_Regions = IIf(bCapitalise = True, UCase(sarr_Let(lReg)), sarr_Let(lReg))
                Exit Function
            End If
        Next
    Else    ' Else return the number
        For lReg = 0 To UBound(sarr_Let, 1)
            If UCase(vRegion) = UCase(sarr_Let(lReg)) Then
                g_Regions = sarr_Nos(lReg)
                Exit Function
            End If
        Next
    End If
    MsgBox vRegion & " not found in g_Regions"
    Stop
    Exit Function
    
End Function


Public Function g_KillFile(SPathFilExt As String)
    ' Will kill any files that exist with the name and path specified
    On Error Resume Next
    Kill SPathFilExt
End Function

Public Function g_GetNo(ByVal sInput As String) As Long
    ' Will deliver back the first number in a string i.e. M32_Label will return 32
    Dim lIdx As Long, sOutput As String
    sOutput = sInput
    For lIdx = 1 To Len(sInput)
        If IsNumeric(Left(sOutput, 1)) = False Then
            sOutput = g_Right(sOutput, 1)
        Else
            g_GetNo = Val(sOutput)
            Exit For
        End If
    Next
End Function

Public Function sProcNameSSTagAuth(ByVal sTag As String, ByVal lAuthority As Long) As String
    ' Take out the other authorities so that there is have a clear check format
    ' I.e. if it comes in with e.g. '1Lock' and the authority is 0, '0Lock' will be returned
    sProcNameSSTagAuth = g_KeepReplace(sTag, "Number", "", CStr(lAuthority))

End Function


Public Function CBA_ExitDate(ctlTmp As Control, sLabel As String, Optional strDateFmt As String = CBA_DMY) As Boolean
    ' Process the results of the entry into the Date Textbox
    ' If an invalid date, will return cancel = true (Cancel will cancel the exit or update, depending on where you put it)
    On Error Resume Next
    Dim sTempDate As String
    CBA_ExitDate = False
    sTempDate = g_FixDate(ctlTmp.Value & "")
    If Len(Trim(ctlTmp.Value & "")) = 0 Then
        ctlTmp.Value = ""
    ElseIf g_IsDate(sTempDate, True) = True Then
        ctlTmp.Value = Format(sTempDate, strDateFmt)
    Else
        MsgBox "'" & sLabel & "' Invalid date entered", vbOKOnly, "Invalid Date"
        ctlTmp.SetFocus
        CBA_ExitDate = True
        Exit Function
    End If
    Exit Function
    
End Function

Public Function CBA_getLabelCaption(meCtls As Object, ctlTmp As Control, Optional bOnErrBringBackName As Boolean = False) As String
    ' Get the Label Caption for the entry field if it exists, or the field name
    '               This is so that we can incude the caption in any warning for the field
    
    Dim frmf As Control, sField As String, sField1 As String, sPrefix As String, lIdx As Long
    On Error GoTo Err_Routine
    
    sPrefix = Left(ctlTmp.Name, 3)
    If InStr(1, UCase(",cmd,opt,chk,lbl,"), UCase(sPrefix)) > 0 Then
        sPrefix = ctlTmp.Name
    Else
        sField = Replace(ctlTmp.Name, sPrefix, "lbl")
    End If
    ' Will pull an error if not found
    Set frmf = meCtls(sField)
    CBA_getLabelCaption = frmf.Caption
Exit_Routine:
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_getLabelCaption", 3)
    On Error GoTo Exit_Routine
    If bOnErrBringBackName Then
        CBA_getLabelCaption = ctlTmp.Name
        GoTo Exit_Routine
    End If
    sField = g_Right(ctlTmp.Name, 3)
    sField1 = ""
    For lIdx = 1 To Len(sField)
        If Mid(sField, lIdx, 1) = UCase(Mid(sField, lIdx, 1)) Then
            sField1 = sField1 & Mid(sField, lIdx, 1) & " "
        Else
            sField1 = sField1 & Mid(sField, lIdx, 1)
        End If
    Next
    CBA_getLabelCaption = sField1
    GoTo Exit_Routine
End Function

Public Function CBA_getAdminUsers(Optional ByVal sUserName As String = "") As Boolean
    Dim aAry() As String, cCell As Range
    On Error GoTo Err_Routine

    ' Return the authority level of the user    - 0 = No Authority (read only), 1 = Admin Authority, 2 = GBDM Authority, 3 = BD Authority
    ' See if the user name is an Admin
    CBA_getAdminUsers = False
    If Trim(sUserName) = "" Then sUserName = Application.UserName
    aAry() = Split(sUserName & "(vvv)", "(")
    If aAry(0) = "" Then GoTo Err_Routine
    For Each cCell In CBAR_AdminUsers.Columns(1).Cells
        If cCell.Value = "" Then Exit For
        If cCell.Value Like aAry(0) & "*" Then
            CBA_getAdminUsers = True
            Exit For
        End If
    Next
    If InStr(1, Application.UserName, "Baines, Stuart") > 0 Then CBA_getAdminUsers = True                ' @RW Take out when have put Stuart into the worksheet

Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_getAdminUsers", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function

Public Function CBA_getUserShortTitle(ByRef sReturnTitle As String, ByRef sReturnShortTitle As String, Optional ByVal sUserName As String = "", Optional ByRef ReturnUserName As String = "", Optional ByRef sInits As String = "") As Long
    ' Return the UserName, Short Title and/or Title of the user
    
    Dim aUN() As String, aTtl() As String, aInits() As String, lIdx As Long
    On Error GoTo Err_Routine
    Dim sString As String
    sString = "CB AUS/BM|CB AUS/BD|CB AUS/BA|CB AUS/CRM|CB AUS/MktgM|CB AUS/MktgA|CB AUS/Mktg|CB AUS/GBD|CB AUS/GBDA|CB AUS/MerchA|CB AUS/Merch|CB AUS/MerchM|CB AUS/BSA|CB AUS/BAA|CB AUS/BOF|"
    sString = sString & "CB AUS/QAOA|CB AUS/QAA|CB AUS/QAS|CB AUS/QAN|CB AUS/QM|CB AUS/MD|CB AUS/MDPA"
    CBA_getUserShortTitle = -1
    If Trim(sUserName) = "" Then sUserName = Application.UserName
    aInits = Split(sUserName, ",")
    If UBound(aInits, 1) > 0 Then
       sInits = Left(Trim(aInits(1)), 1) & Left(Trim(aInits(0)), 1)
    Else
       sInits = "__"
    End If
    aUN() = Split(Replace(Trim(sUserName & "(vvv)"), ")", ""), "(")
    ReturnUserName = Trim(aUN(0))
    If Trim(aUN(0)) = "" Then GoTo Err_Routine
    aTtl() = Split(sString, "|")
    
    For lIdx = 0 To UBound(aTtl, 1)
        If aUN(1) = aTtl(lIdx) Then
            CBA_getUserShortTitle = lIdx
            sReturnTitle = aUN(1)
            sReturnShortTitle = Trim(Replace(aUN(1), "CB AUS/", ""))
            Exit For
        End If
    Next

Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_getUserShortTitle", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function

Public Function CBA_User() As String
    ' Hold the current user instead of getting it all the time from Application.UserName. (Also it can be changed for testing individual users)
    Static sUser As String
    If sUser = "" Then Call CBA_getUserShortTitle("", "", CBA_SetUser, sUser)
    CBA_User = sUser
End Function

Public Function CBA_getVersionStatus(sDB As String, sDspVersion As String, sApp As String, sAppN As String, Optional bPrompt As Boolean = False, Optional sTablename As String = "A0_AppVersion") As String
    ' Will check to see if the user has the latest version (not of Excel but of the particular app that is being used)
    ' Version will now come from the Central DB, in the same format
    
    Dim lDBVer As Long, lAddInVer As Long, sAddInVer As String, sTest As String
    Static bHasBeenRun As Boolean, sLastApp As String, sLastTest As String, sGenDB As String
    sTest = g_KeepReplace(Right(sDB, 15), "Alpha", "")
    If Len(sTest) > 5 Then
        sLastTest = Replace(sDspVersion & ".test", "Version:", "Ver:")
        CBA_TestIP = "Y"
    Else
        sLastTest = sDspVersion
        CBA_TestIP = "N"
    End If
    bHasBeenRun = True: sLastApp = sApp
    CBA_getVersionStatus = sLastTest
    sGenDB = g_GetDB("Gen")
    lDBVer = CLng(g_DLookup("Vn_ID", sTablename, "Vn_App= '" & sAppN & "' AND Vn_ID>0", "Vn_ID DESC", sGenDB, 0))
    ' Remove all the letters and periods
    sAddInVer = g_KeepReplace(sDspVersion, "Alpha", "")
    lAddInVer = Val(sAddInVer)
    ' If the version in the Addin is greater / equal to the version in the DB
    If lDBVer <= lAddInVer Then
        ' Prompt for info
    ElseIf lAddInVer < lDBVer Then
        If bPrompt Then
            If MsgBox("Your version of " & sApp & " is not the latest..." & vbCrLf & "Press Yes to exit the chosen option or No to continue", vbYesNo) = vbYes Then
                CBA_getVersionStatus = "Exit"
            End If
        Else
            MsgBox "Your version of " & sApp & " is not the latest..." & vbCrLf & " Please restart Excel to enjoy the features of the latest version", vbOKOnly
        End If
    End If
    
End Function

Public Function g_GetSQLDate(ByVal vInputDate, Optional ByVal sInputFormat As String = "dd/mm/yyyy") As String
    '    This routine will bring back a structured SQL Date from an input (ready to be written to the table)
    ' DATE HAS TO BE BROUGHT IN AS THE "dd/mm" TYPE FORMAT (ALTHOUGH FIXDATE WILL HANDLE SOME DIFFERENCES)
    
    Dim sDate As String, sOutputFormat As String
    
    sOutputFormat = Replace(sInputFormat, "dd/mm", "mm/dd")   ' Fix format if wrong for SQL dates
    sDate = g_FixDate(vInputDate, "dd/mm/yyyy hh:nn")
    If g_IsDate(sDate, True) = True Then
        g_GetSQLDate = "#" & g_FixDate(sDate, sOutputFormat) & "#"
    Else
        g_GetSQLDate = "NULL"
    End If
End Function
Sub Set_CBA_DBtoQuery(ByVal DBno As Long)
    CBA_DBtoQuery = DBno
End Sub

