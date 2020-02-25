Attribute VB_Name = "CBA_COM_T_ProduceDSRuntime"
Option Explicit
Option Private Module       ' Excel users cannot access procedures
Private Type ProdDSet
    ReferenceDate As Date
    Market As String
    Region As String
    Competitor As String
    Location As String
    productcode As Long
    ProductDescription As String
    ALDIPackSize As String
    Casesize As Long
    LastRetail As Single
    Retail As Single
    PromoRetail As Single
    ALDIWeightCount As Single
    CompetitorWeightCount As Single
    WeightCountUnit As String
    Comment As String
    Instore As Boolean
    InCatalogue As Boolean
    Discounted As Boolean
    CompPCode As String
    perKGident As Boolean
End Type
'Callback for CBA_COM_T_ProduceDS getEnabled
Sub CBA_COM_getEnabledCOMRADETools(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = True
End Sub

Sub CBA_COM_T_ProduceDSRun(Control As IRibbonControl)
    Dim PDS_CN As ADODB.Connection
    Dim PDS_RS As ADODB.Recordset
    Dim PDS_CM As Variant
    Dim a As Long, b As Long, c As Long, d As Long, lRet As Long
    Dim strSQL As String, curPack As String, dtToday As String
    Dim arrStates(1 To 5, 1 To 2) As String
    Dim arrDS() As ProdDSet
    Dim bfound As Boolean
    Dim ScrapedDate As Date, ColessDate As Date, WWsDate As Date, wedDate, curProd, wks_PDS, lRowNo
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
        
    If CBA_getVersionStatus(g_GetDB("Gen"), CBA_COM_Ver, "Comrade", "COM", True) = "Exit" Then Exit Sub

    lRet = MsgBox("This is going to generate a report for Produce PTP Interface. Do you wish to proceed?", vbYesNo)
    If lRet <> 6 Then Exit Sub
    
    CBA_BasicFunctions.CBA_Running "Preparing produce COMRADE datasheet"
    Application.ScreenUpdating = False
    ReDim arrDS(1 To 2, 1 To 5, 1 To 1)
    'wedDate = DateValue("15/08/2018")
    
    wedDate = DateAdd("D", -WeekDay(Date, vbWednesday) + 1, Date)
    dtToday = Format(Date, "yyyy-mm-dd")
    
    Set PDS_CN = New ADODB.Connection: Set PDS_RS = New ADODB.Recordset
    With PDS_CN
        .ConnectionTimeout = 300
        .CommandTimeout = 300
        .Open "Provider= SQLNCLI10; DATA SOURCE=" & CBA_BasicFunctions.TranslateServerName("599DBL12", Date) & "; ;INTEGRATED SECURITY=sspi;"
    End With
    strSQL = "select top 1 datescraped from tools.dbo.com_c_prod order by datescraped desc"
    PDS_RS.Open strSQL, PDS_CN
    ColessDate = PDS_RS.Fields(0)
    PDS_RS.Close: Set PDS_RS = Nothing: PDS_CN.Close: Set PDS_CN = Nothing
    
    Set PDS_CN = New ADODB.Connection: Set PDS_RS = New ADODB.Recordset
    With PDS_CN
        .ConnectionTimeout = 300
        .CommandTimeout = 300
        .Open "Provider= SQLNCLI10; DATA SOURCE=" & CBA_BasicFunctions.TranslateServerName("599DBL12", Date) & "; ;INTEGRATED SECURITY=sspi;"
    End With
    strSQL = "select top 1 datescraped from tools.dbo.com_w_prod order by datescraped desc"
    PDS_RS.Open strSQL, PDS_CN
    WWsDate = PDS_RS.Fields(0)
    PDS_RS.Close: Set PDS_RS = Nothing: PDS_CN.Close: Set PDS_CN = Nothing
    
    
    If ColessDate = WWsDate Then
        ScrapedDate = ColessDate
        ScrapedDate = ScrapedDate - 0 ' Use this to print reports for prior days
    Else
        MsgBox "Error: Coles and WW Data dates do not match. Please contact CBAnalytics@aldi.com.au", vbYesNo
        Exit Sub
    End If
    
    Set PDS_CN = New ADODB.Connection
    Set PDS_RS = New ADODB.Recordset
    With PDS_CN
        .ConnectionTimeout = 300
        .CommandTimeout = 300
        .Open "Provider= SQLNCLI10; DATA SOURCE= 599DBL01; ;INTEGRATED SECURITY=sspi;"
    End With
        
    strSQL = "Select productcode, packsize from cbis599p.dbo.product where cgno = 58"
    PDS_RS.Open strSQL, PDS_CN
    PDS_CM = PDS_RS.GetRows()
    
    arrStates(1, 1) = "NSW": arrStates(1, 2) = 1
    arrStates(2, 1) = "VIC": arrStates(2, 2) = 2
    arrStates(3, 1) = "QLD": arrStates(3, 2) = 3
    arrStates(4, 1) = "SA": arrStates(4, 2) = 4
    arrStates(5, 1) = "WA": arrStates(5, 2) = 5
    
    If CBA_COM_SetupMatchArray.CBA_SetupMatchArray(False, DateAdd("D", -1, ScrapedDate), ScrapedDate, 58, , , True) = True Then
    
    b = 0
    For a = LBound(CBA_COM_Match, 1) To UBound(CBA_COM_Match, 1)
        'If CBA_COM_Match(a).AldiPCode = PDS_CM(0, pr) Then
        If CBA_COM_Match(a).AldiPCode <> curProd And CBA_COM_Match(a).CompDivideby <> "0" Then
            curProd = CBA_COM_Match(a).AldiPCode
'                If curprod = 77426 Then
'                a = a
'                End If

            bfound = False
            For c = LBound(PDS_CM, 2) To UBound(PDS_CM, 2)
                If PDS_CM(0, c) = curProd Then
                    bfound = True
                    curPack = PDS_CM(1, c)
                    Exit For
                End If
            Next
            If bfound = False Then curPack = ""
            'Else
            b = b + 1
            ReDim Preserve arrDS(1 To 2, 1 To 5, 1 To b)
            'arrDS(Competitor, state, noofrecords)
            'End If
        End If
        If CBA_COM_Match(a).CompDivideby <> "0" Then
            For c = 1 To 5
                If CBA_COM_Match(a).compet = "WW" Then
                    d = 1
                ElseIf CBA_COM_Match(a).compet = "C" Then
                    d = 2
                Else
                    Exit For
                End If
            
                If (CBA_COM_Match(a).Pricedata(ScrapedDate, CBA_COM_Match(a).HowCompvalue, arrStates(c, 1)) <> 0 And arrDS(d, c, b).Retail = 0) _
                    Or arrDS(d, c, b).Retail > CBA_COM_Match(a).Pricedata(ScrapedDate, CBA_COM_Match(a).HowCompvalue, arrStates(c, 1)) Then
                    
                    If InStr(1, LCase(CBA_COM_Match(a).MatchType), "nat") Or (c = 1 And InStr(1, CBA_COM_Match(a).MatchType, "NSW")) Or (c = 2 And InStr(1, CBA_COM_Match(a).MatchType, "VIC")) Or (c = 3 And InStr(3, CBA_COM_Match(a).MatchType, "QLD")) _
                        Or (c = 4 And InStr(1, CBA_COM_Match(a).MatchType, "SA")) Or (c = 5 And InStr(1, CBA_COM_Match(a).MatchType, "WA")) Then
    
                        If curProd = "4703" Then
                            a = a
                        End If
    
                        arrDS(d, c, b).CompPCode = CBA_COM_Match(a).CompCode
                        arrDS(d, c, b).ReferenceDate = ScrapedDate
                        arrDS(d, c, b).Market = arrStates(c, 1)
                        arrDS(d, c, b).ProductDescription = CBA_COM_Match(a).AldiPName
                        arrDS(d, c, b).InCatalogue = CBA_COM_Match(a).Pricedata(ScrapedDate, "isspecial", arrStates(c, 1))
                        arrDS(d, c, b).Competitor = CBA_COM_Match(a).Competitor
                        arrDS(d, c, b).productcode = curProd
                        arrDS(d, c, b).WeightCountUnit = CBA_COM_Match(a).HowComp
                        If CBA_COM_Match(a).HowCompvalue = "permeasure" And CBA_COM_Match(a).HowComp = "G" And InStr(1, LCase(CBA_COM_Match(a).AldiPName), "per kg") > 0 Then arrDS(d, c, b).perKGident = True
                        arrDS(d, c, b).Casesize = curPack
                        
                        Select Case CBA_COM_Match(a).HowComp
                            Case "G"
                                If CBA_COM_Match(a).CompDivideby >= 1000 And CBA_COM_Match(a).CompMultby >= 1000 And CBA_COM_Match(a).HowComp = "G" Then
                                    If arrDS(d, c, b).perKGident = True Then
                                        arrDS(d, c, b).WeightCountUnit = "per kilogram"
                                    Else
                                        arrDS(d, c, b).WeightCountUnit = "kilograms"
                                    End If
                                Else
                                    arrDS(d, c, b).WeightCountUnit = "grams"
                                End If
                            Case "PK"
                                arrDS(d, c, b).WeightCountUnit = "pack"
                            Case Else
                                arrDS(d, c, b).WeightCountUnit = CBA_COM_Match(a).HowComp
                        End Select
                        
                        If CBA_COM_Match(a).HowComp = "G" And CBA_COM_Match(a).CompOriginalPack <> 1000 And CBA_COM_Match(a).CompOriginalPack <> 0 Then
                            If arrDS(d, c, b).perKGident = True Then
                                arrDS(d, c, b).LastRetail = (CBA_COM_Match(a).Pricedata(DateAdd("D", -1, ScrapedDate), CBA_COM_Match(a).HowCompvalue, arrStates(c, 1)) / CBA_COM_Match(a).CompDivideby) * CBA_COM_Match(a).CompOriginalPack
                                arrDS(d, c, b).Retail = (CBA_COM_Match(a).Pricedata(ScrapedDate, CBA_COM_Match(a).HowCompvalue, arrStates(c, 1)) / CBA_COM_Match(a).CompDivideby) * CBA_COM_Match(a).CompOriginalPack
                                arrDS(d, c, b).PromoRetail = (CBA_COM_Match(a).Pricedata(ScrapedDate, "promotional", arrStates(c, 1)) / CBA_COM_Match(a).CompDivideby) * CBA_COM_Match(a).CompOriginalPack
                            Else
                                arrDS(d, c, b).LastRetail = CBA_COM_Match(a).Pricedata(DateAdd("D", -1, ScrapedDate), "shelf", arrStates(c, 1))
                                arrDS(d, c, b).Retail = CBA_COM_Match(a).Pricedata(ScrapedDate, "shelf", arrStates(c, 1))
                                arrDS(d, c, b).PromoRetail = CBA_COM_Match(a).Pricedata(ScrapedDate, "promotional", arrStates(c, 1))
                            End If
                        Else
                            arrDS(d, c, b).LastRetail = CBA_COM_Match(a).Pricedata(DateAdd("D", -1, ScrapedDate), CBA_COM_Match(a).HowCompvalue, arrStates(c, 1))
                            arrDS(d, c, b).Retail = CBA_COM_Match(a).Pricedata(ScrapedDate, CBA_COM_Match(a).HowCompvalue, arrStates(c, 1))
                            arrDS(d, c, b).PromoRetail = CBA_COM_Match(a).Pricedata(ScrapedDate, "promotional", arrStates(c, 1))
                        End If
                        
                        
                        If arrDS(d, c, b).WeightCountUnit = "per kilogram" Or arrDS(d, c, b).WeightCountUnit = "kilograms" Then
                            If arrDS(d, c, b).perKGident = True Then
                                arrDS(d, c, b).ALDIPackSize = Format(CBA_COM_Match(a).CompMultby / 1000, "0.00") & " " & arrDS(d, c, b).WeightCountUnit
                                arrDS(d, c, b).CompetitorWeightCount = CBA_COM_Match(a).CompDivideby / 1000
                                arrDS(d, c, b).ALDIWeightCount = CBA_COM_Match(a).CompMultby / 1000
                            Else
                                arrDS(d, c, b).ALDIPackSize = Format(CBA_COM_Match(a).CompMultby / 1000, "0.00") & " " & arrDS(d, c, b).WeightCountUnit
                                If CBA_COM_Match(a).CompOriginalPack = 0 Then
                                    If CBA_COM_Match(a).Pricedata(ScrapedDate, "permeasure", arrStates(c, 1)) <> 0 Then
                                        arrDS(d, c, b).CompetitorWeightCount = (arrDS(d, c, b).Retail / CBA_COM_Match(a).Pricedata(ScrapedDate, "permeasure", arrStates(c, 1)))
                                    Else
                                        arrDS(d, c, b).CompetitorWeightCount = CBA_COM_Match(a).CompDivideby / 1000
                                    End If
                                Else
                                    arrDS(d, c, b).CompetitorWeightCount = CBA_COM_Match(a).CompOriginalPack / 1000
                                End If
                                arrDS(d, c, b).ALDIWeightCount = CBA_COM_Match(a).CompMultby / 1000
                            End If
                        Else
                            If CBA_COM_Match(a).CompOriginalPack = 0 Then
                                arrDS(d, c, b).ALDIPackSize = Format(CBA_COM_Match(a).CompMultby, "0.00") & " " & arrDS(d, c, b).WeightCountUnit
                                arrDS(d, c, b).CompetitorWeightCount = CBA_COM_Match(a).CompDivideby
                                arrDS(d, c, b).ALDIWeightCount = CBA_COM_Match(a).CompMultby
                            Else
                                arrDS(d, c, b).ALDIPackSize = Format(CBA_COM_Match(a).CompMultby, "0.00") & " " & arrDS(d, c, b).WeightCountUnit
                                arrDS(d, c, b).CompetitorWeightCount = CBA_COM_Match(a).CompOriginalPack
                                arrDS(d, c, b).ALDIWeightCount = CBA_COM_Match(a).CompMultby
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Next
    End If
    
    
    CBAR_ProduceDS.Copy
    Set wks_PDS = ActiveSheet
    
    With wks_PDS
        Range(.Cells(2, 1), .Cells(999, 25)).ClearContents
        lRowNo = 1
        For b = LBound(arrDS, 3) To UBound(arrDS, 3)
            For c = 1 To 5
                For d = 1 To 2
                    If arrDS(d, c, b).Retail <> 0 Then
                        lRowNo = lRowNo + 1
                        .Cells(lRowNo, 1) = arrDS(d, c, b).ReferenceDate
                        .Cells(lRowNo, 2) = arrDS(d, c, b).Market
                        .Cells(lRowNo, 3) = arrDS(d, c, b).Region
                        .Cells(lRowNo, 4) = arrDS(d, c, b).Competitor
                        .Cells(lRowNo, 5) = arrDS(d, c, b).Location
                        .Cells(lRowNo, 6) = arrDS(d, c, b).productcode
                        .Cells(lRowNo, 7) = arrDS(d, c, b).ProductDescription
                        .Cells(lRowNo, 8) = arrDS(d, c, b).ALDIPackSize
                        .Cells(lRowNo, 9) = arrDS(d, c, b).Casesize
                        .Cells(lRowNo, 10) = arrDS(d, c, b).LastRetail
                        .Cells(lRowNo, 11) = arrDS(d, c, b).Retail
                        .Cells(lRowNo, 12) = arrDS(d, c, b).PromoRetail
                        .Cells(lRowNo, 13) = arrDS(d, c, b).ALDIWeightCount
                        .Cells(lRowNo, 14) = arrDS(d, c, b).CompetitorWeightCount
                        .Cells(lRowNo, 15) = arrDS(d, c, b).WeightCountUnit
                        .Cells(lRowNo, 16) = arrDS(d, c, b).Comment
                        If arrDS(d, c, b).Instore = True Then .Cells(lRowNo, 17) = "x"
                        If arrDS(d, c, b).InCatalogue = True Then .Cells(lRowNo, 18) = "x"
                        If arrDS(d, c, b).Discounted = True Then .Cells(lRowNo, 19) = "x"
                    End If
                Next
            Next
        Next
        
        
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=Range("F2:F" & lRowNo), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SortFields.Add Key:=Range("D2:D" & lRowNo), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SortFields.Add Key:=Range("C2:C" & lRowNo), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange Range("A1:S" & lRowNo)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        With .Rows("2:" & lRowNo).Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Range(.Cells(2, 1), .Cells(lRowNo, 19)).Borders.Weight = xlThin
        .Cells(1, 1).Select
    End With
    
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    Application.ScreenUpdating = True
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_COM_T_ProduceDSRun", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

