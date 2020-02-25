Attribute VB_Name = "CBA_CDS_Runtime"
Option Explicit
Type IncoData
    id As Long
    ContractNo As Long
    DivNo As Integer
    ValidFrom As Date
    ValidTo As Date
    IncoTerm As String
End Type
Type AdjCostType
    DateFrom As Date
    DateTo As Date
    SchemeID As String
    TotalCost As Single
    Rebate As Single
End Type

Private CDS_DateFrom As Date
Private CDS_DateTo As Date
Private CDS_NSWSelected As Boolean
Private CDS_QLDSelected As Boolean
Private CDS_ACTSelected As Boolean
Private CDS_EXP_QLDSelected As Boolean
Private CDS_EXP_ACTSelected As Boolean
Private GoodToRun As Boolean

Private RCVDivProdContDic As Scripting.Dictionary
Private CDSBase As Variant
Private sR(501 To 509) As Scripting.Dictionary
Private CBISDic As Scripting.Dictionary

Sub testfuncCDSMargin()
Dim thisval As Variant

    thisval = getCDSMargin(61489, DateSerial(2019, 6, 1), DateSerial(2019, 6, 30))
    
    
End Sub

Public Function getCDSMargin(ByVal PCode As Long, ByVal DateFrom As Date, ByVal DateTo As Date) As Single
Dim CBIS_RS As ADODB.Recordset, CDS_RS As ADODB.Recordset
Dim CDS_CN As ADODB.Connection, CBIS_CN  As ADODB.Connection
Dim strSQL As String, strConts As String
Dim CBISDATA As Variant, CDSDATA As Variant, CDSCost As Variant


Set CDS_CN = New ADODB.Connection
With CDS_CN
    .CommandTimeout = 50
    .ConnectionTimeout = 50
    .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & CBA_BSA & "LIVE DATABASES\CDS_ADM.accdb"
End With
Set CBIS_CN = New ADODB.Connection
With CBIS_CN
    .ConnectionTimeout = 50
    .CommandTimeout = 50
    .Open "Provider= SQLNCLI10; DATA SOURCE= 599DBL01; ;INTEGRATED SECURITY=sspi;"
End With

    Set CBIS_RS = New ADODB.Recordset
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) '& "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10) & "DECLARE @CWDATE as Date" & Chr(10) & "DECLARE @LWDATE as Date" & Chr(10)
    strSQL = strSQL & "declare @DateFrom date = '" & Format(DateFrom, "YYYY-MM-DD") & "'" & Chr(10)
    strSQL = strSQL & "declare @DateTo date = '" & Format(DateTo, "YYYY-MM-DD") & "'" & Chr(10)
    strSQL = strSQL & "declare @Prod int = " & PCode & Chr(10)
    strSQL = strSQL & "select dayenddate, productcode, contractno, divno, sum(retail) AS RETAIL, sum(retailnet) AS RETAILNET, sum(cost) AS COST, case when recordID = '001' then sum(Pieces) else 0  end as QTY" & Chr(10)
    strSQL = strSQL & "into #RCV from cbis599p.dbo.receiving rcv" & Chr(10)
    strSQL = strSQL & "where dayenddate <= @DateTo and dayenddate >= @Datefrom and productcode = @Prod" & Chr(10)
    strSQL = strSQL & "group by dayenddate, productcode, contractno, divno, RecordID" & Chr(10)
    strSQL = strSQL & "select contractno into #C from cbis599p.dbo.contract c where c.productcode = @Prod" & Chr(10)
    strSQL = strSQL & "and c.deliveryfrom <= @DateFrom and isnull(c.deliveryto,convert(date,getdate())) >= @DateTo" & Chr(10)
    strSQL = strSQL & "select bv.Contractno, divno, case when validfrom < @DateFrom then @DateFrom else Validfrom end as ValidFrom" & Chr(10)
    strSQL = strSQL & ", case when isnull(validto, convert(date,getdate())) > @DateTo then @DateTo else isnull(validto, convert(date,getdate())) end as validto" & Chr(10)
    strSQL = strSQL & ",case when FOB is null and ExWorks is null and DDP is not null then 'DDP' when FOB is null and ExWorks is not null then 'ExWorks'" & Chr(10)
    strSQL = strSQL & "when FOB is not null and ExWorks is null then 'FOB' else 'Unknown' end as incoterms" & Chr(10)
    strSQL = strSQL & "into #INCO from cbis599p.dbo.BRAKETVALUE bv inner join #C c on c.ContractNo = bv.ContractNo" & Chr(10)
    strSQL = strSQL & "where isnull(validto, convert(date,getdate())) >= @DateFrom and validfrom <= @DateTo order by ContractNo, divno" & Chr(10)
    strSQL = strSQL & "select distinct contractno from #RCV"
    CBIS_RS.Open strSQL, CBIS_CN
    If CBIS_RS.EOF Then
        getCDSMargin = 0
        Exit Function
    Else
        strConts = ""
        Do Until CBIS_RS.EOF
            If strConts = "" Then strConts = CBIS_RS(0) Else strConts = strConts & ", " & CBIS_RS(0)
            CBIS_RS.MoveNext
        Loop
    End If
    
    Set CBIS_RS = New ADODB.Recordset
    strSQL = "select rcv.*, i.incoterms  from #RCV rcv left join  #INCO i on i.ContractNo = rcv.ContractNo and i.DivNo = rcv.DivNo and i.ValidFrom <= rcv.DayEndDate and i.validto >= rcv.DayEndDate" & Chr(10)
    strSQL = strSQL & "order by rcv.ContractNo, rcv.DivNo, i.incoterms, rcv.DayEndDate"
    CBIS_RS.Open strSQL, CBIS_CN
    If CBIS_RS.EOF Then
        getCDSMargin = 0
        Exit Function
    Else
        CBISDATA = CBIS_RS.GetRows()
    End If
    
    Set CBIS_RS = New ADODB.Recordset
    strSQL = "drop table #RCV,#C, #INCO" & Chr(10)
    CBIS_RS.Open strSQL, CBIS_CN
    CBIS_CN.Close
    Set CBIS_CN = Nothing
    
    
    Set CDS_RS = New ADODB.Recordset
    strSQL = "Select *  from Cont_Map where contractno in (" & strConts & ")"
    CDS_RS.Open strSQL, CDS_CN
    If CDS_RS.EOF Then
        getCDSMargin = 0
        Exit Function
    Else
        CDSDATA = CDS_RS.GetRows()
    End If
    
    Set CDS_RS = New ADODB.Recordset
    strSQL = "Select *  from Scheme_Cost where DateTo >= #" & Format(DateFrom, "MM/DD/YYYY") & "# and DateFrom <= #" & Format(DateTo, "MM/DD/YYYY") & "#"
    CDS_RS.Open strSQL, CDS_CN
    If CDS_RS.EOF Then
        getCDSMargin = 0
        Exit Function
    Else
        CDSCost = CDS_RS.GetRows()
    End If
    CDS_CN.Close
    Set CDS_CN = Nothing
    
    
'    For b = LBound(CBISDATA, 2) To UBound(CBISDATA, 2)
'        For d = LBound(CDSDATA, 2) To UBound(CDSDATA, 2)
'            If CBISDATA(2, b) = CDSDATA(1, d) And CBISDATA(3, b) = CDSDATA(2, d) And CBISDATA(2, b) = CDSDATA(1, d) Then
'
'            End If
'        Next d
'    Next b
    
    
        
    
    
    
    
    
    
    
    



End Function
Function CBA_CDS_SetParamaters(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal NSW As Boolean, ByVal QLD As Boolean, ByVal ACT As Boolean, ByVal EXP_ACT As Boolean, ByVal EXP_QLD As Boolean) As Boolean
    CDS_DateFrom = DateFrom
    CDS_DateTo = DateTo
    CDS_NSWSelected = NSW
    CDS_QLDSelected = QLD
    CDS_ACTSelected = ACT
    CDS_EXP_QLDSelected = EXP_QLD
    CDS_EXP_ACTSelected = EXP_ACT
    If NSW = True Or QLD = True Or ACT = True Or EXP_QLD = True Or EXP_ACT = True Then
        CBA_CDS_SetParamaters = True
        GoodToRun = True
    Else
        CBA_CDS_SetParamaters = False
        GoodToRun = False
    End If
End Function
Sub ProceedToCDSReport()
Dim CN As ADODB.Connection
Dim RS As ADODB.Recordset
Dim ProdCol As Scripting.Dictionary, Storecol As Scripting.Dictionary, Datecol As Scripting.Dictionary
Dim CBD As CBA_CBIS_CDS
Dim Inco() As IncoData
Dim DFrom As Date, Dto As Date, curDate As Date
Dim CDS_NSW()
Dim wks As Worksheet, sht As Worksheet, wks_sum As Worksheet
Dim wks_NSW As Worksheet, wks_QLD As Worksheet, wks_ACT As Worksheet, wks_EXP_ACT As Worksheet, wks_EXP_QLD As Worksheet
Dim WB, CBAddIn
Dim e As Boolean
Dim wbk As Workbook
Dim RCell As Range
Dim headrow As Long, col As Long, PivNo As Long, cnt As Long, DateQTY As Long, ProdQTY As Long, i As Long
Dim strConts As String, strSQL As String, strProds As String, curProd As String, curDiv As String
Dim div As Integer, curStore As Integer
Dim col501 As Collection, col502 As Collection, col503 As Collection, col504 As Collection, NSWCol As Collection, ACTCol As Collection
Dim EXP_ACTCol As Collection, col506 As Collection, QLDCol As Collection, EXP_QLDCol As Collection
Dim rw As Long, PCode As Long, calcTot As Single
Dim a As Variant, st As Variant
''Dim bFound As Boolean
Dim RCV As Single, b As Long
Dim DT As Date
Dim but As Object
Dim inArray() As Variant
Dim TempDic As Scripting.Dictionary, dicContainers As Scripting.Dictionary
Dim wbkOrig As Workbook
Dim ProdDic As Scripting.Dictionary, ContDic As Scripting.Dictionary
Dim NSWDic As New Scripting.Dictionary
'Dim CData As CBA_CDS_Cont, TCData As CBA_CDS_Cont
    
    Set wbkOrig = ActiveWorkbook
    CBA_CDS_frm.Show
    If GoodToRun = False Then Exit Sub
    

    For Each wks In ActiveWorkbook.Worksheets
        If wks.Cells(4, 3).Value = "CDS 'Not yet eligible flagged' Report" Then
            Set sht = wks
            Exit For
        End If
    Next
    
    
'    For Each WB In Application.AddIns
'        If InStr(1, WB.Name, "CBStdAddin") > 0 Then
'            bFound = True
'            Set CBAddIn = WB
'            Exit For
'        End If
'    Next
'    If bFound = False Then
'        MsgBox "The Buying Excel Ribbon cannot be found. Please contact CBAnalytics to have it installed 9218", vbOKOnly
'        wbk.Close
'        Exit Sub
'    End If
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then
        CBA_BasicFunctions.CBA_Running "CDS Report"
    End If
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Preparing Base Data..."
    Application.ScreenUpdating = False

    For Each RCell In sht.Columns(1).Cells
        If RCell.Row > 5 And RCell.Value = "" And RCell.Offset(1, 0).Value = "" And RCell.Offset(2, 0).Value = "" And RCell.Offset(3, 0).Value = "" And RCell.Offset(4, 0).Value = "" And RCell.Offset(5, 0).Value = "" Then Exit For
        If LCase(RCell.Value) = "contractno" Then
            headrow = RCell.Row
            Exit For
        End If
    Next
    For Each RCell In sht.Rows(headrow).Cells
        If RCell.Value = "" And RCell.Offset(0, 1).Value = "" And RCell.Offset(0, 2).Value = "" And RCell.Offset(0, 3).Value = "" And RCell.Offset(0, 4).Value = "" And RCell.Offset(0, 5).Value = "" Then Exit For
        If LCase(RCell.Value) = "eligible?" Then
            col = RCell.Column
            Exit For
        End If
    Next
    e = False
    For Each RCell In sht.Columns(col).Cells
        If RCell.Row > headrow And sht.Cells(RCell.Row, 1).Value = "" Then Exit For
        If RCell.Row > headrow And sht.Cells(RCell.Row, 1).Value <> "" And RCell.Value = "" Then
            e = True
            Exit For
        End If
    Next
'    If e = True Then
'        MsgBox "Not all contracts have been correctly flagged as eligible or not."
'        Exit Sub
'    End If
    Set CN = New ADODB.Connection
    With CN
        CN.CommandTimeout = 50
        CN.ConnectionTimeout = 50
        CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & CBA_BSA & "LIVE DATABASES\CDS_ADM.accdb"
    End With
    For Each RCell In sht.Columns(col).Cells
        Set RS = New ADODB.Recordset
        If RCell.Row > headrow And sht.Cells(RCell.Row, 1).Value = "" Then Exit For
        If RCell.Row > headrow And sht.Cells(RCell.Row, 1).Value <> "" And RCell.Value <> "" Then
            strSQL = "INSERT into  Cont_Map (ContractNo, Region, SchemeID, ContainerType, IncoTerm, Eligable,Checked,ALDIResponsible)" & Chr(10)
            strSQL = strSQL & "VALUES (" & sht.Cells(RCell.Row, 1).Value & "," & sht.Cells(RCell.Row, 7).Value & ",'" _
                & sht.Cells(RCell.Row, 10).Value & "','" & sht.Cells(RCell.Row, 11).Value & "','" & sht.Cells(RCell.Row, 12).Value & "'," & IIf(sht.Cells(RCell.Row, 18).Value = "Yes", True, False) _
                & "," & True & "," & IIf(sht.Cells(RCell.Row, 19).Value = "Yes", True, False) & ")"
            RS.Open strSQL, CN
        End If
    Next
    CN.Close
    Set CN = Nothing
    Set RS = Nothing
    Application.DisplayAlerts = False
    wbkOrig.Close
    Application.DisplayAlerts = True
    
    
    CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_Cont_Map"
    CDSBase = CBA_CDSarr
    Erase CBA_CDSarr
    strConts = ""
    Set TempDic = New Scripting.Dictionary
    For a = LBound(CDSBase, 2) To UBound(CDSBase, 2)
        If TempDic.Exists(CDSBase(1, a)) = False Then
            TempDic.Add CDSBase(1, a), CDSBase(1, a)
            If strConts = "" Then strConts = CDSBase(1, a) Else strConts = strConts & ", " & CDSBase(1, a)
        End If
    Next
    CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_CBIS_BaseData", , , , , , , strConts
    Set CBISDic = New Scripting.Dictionary
    For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
        Set CBD = New CBA_CBIS_CDS
        CBD.Branded = False: CBD.Description = "": CBD.productcode = 0: CBD.SupplierName = "": CBD.SupplierNo = 0
        CBD.productcode = CBA_CBISarr(1, a)
        CBD.Description = CBA_CBISarr(2, a)
        CBD.SupplierNo = CBA_CBISarr(3, a)
        CBD.SupplierName = CBA_CBISarr(4, a)
        CBD.Branded = CBA_CBISarr(5, a)
        CBD.Packsize = CBA_CBISarr(6, a)
        CBISDic.Add CBA_CBISarr(0, a), CBD
    Next
    
    CBA_SQL_Queries.CBA_GenPullSQL "CBA_FilterActiveContracts", CDS_DateFrom, CDS_DateTo, , , , , strConts
    strConts = ""
    Set TempDic = New Scripting.Dictionary
    For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
        If TempDic.Exists(CBA_CBISarr(0, a)) = False Then
            TempDic.Add CBA_CBISarr(0, a), CBA_CBISarr(0, a)
            If strConts = "" Then strConts = CBA_CBISarr(0, a) Else strConts = strConts & ", " & CBA_CBISarr(0, a)
        End If
    Next
    
    
    
    CBA_SQL_Queries.CBA_GenPullSQL "CBA_RCV15EitherSide", CDS_DateFrom, , , , , , strConts
    Set RCVDivProdContDic = New Scripting.Dictionary
    Set ProdDic = New Scripting.Dictionary
    Set ContDic = New Scripting.Dictionary
    Set TempDic = New Scripting.Dictionary
    calcTot = 0
    curDiv = CBA_CBISarr(2, 0)
    curProd = CBA_CBISarr(1, 0)
    For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
'        If curProd = 11009 Then
'        a = a
'        End If
        If CBA_CBISarr(2, a) <> curDiv Then
'            If curProd = 4044 Then
'            a = a
'            End If
            ContDic.Add "Total", CStr(calcTot)
            ProdDic.Add CStr(curProd), ContDic
            RCVDivProdContDic.Add CStr(curDiv), ProdDic
            Set ProdDic = New Scripting.Dictionary
            Set ContDic = New Scripting.Dictionary
            curDiv = CBA_CBISarr(2, a)
            calcTot = 0
            curProd = CBA_CBISarr(1, a)
        End If
        If CBA_CBISarr(1, a) <> curProd Then
            ContDic.Add "Total", CStr(calcTot)
            ProdDic.Add CStr(curProd), ContDic
            Set ContDic = New Scripting.Dictionary
            calcTot = 0
            curProd = CBA_CBISarr(1, a)
        End If
'        If CBA_CBISarr(0, a) = 113298 Then
'        a = a
'        End If
        ContDic.Add CStr(CBA_CBISarr(0, a)), CStr(CBA_CBISarr(3, a))
        calcTot = calcTot + CBA_CBISarr(3, a)
        If a = UBound(CBA_CBISarr, 2) Then
            ContDic.Add "Total", CStr(calcTot)
            ProdDic.Add CStr(curProd), ContDic
            RCVDivProdContDic.Add CStr(curDiv), ProdDic
        End If
    Next
    CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_Cont_Map_Where", , , , , , , strConts
    CDSBase = CBA_CDSarr
    Erase CBA_CDSarr
    strConts = ""
    Set TempDic = New Scripting.Dictionary
    For a = LBound(CDSBase, 2) To UBound(CDSBase, 2)
        If TempDic.Exists(CDSBase(1, a)) = False Then
            TempDic.Add CDSBase(1, a), CDSBase(1, a)
            If strConts = "" Then strConts = CDSBase(1, a) Else strConts = strConts & ", " & CDSBase(1, a)
        End If
    Next
    CBA_SQL_Queries.CBA_GenPullSQL "CBA_getUniquePCodesFromContractNo", , , , , , , strConts
    strProds = ""
    If CBA_CBISarr(0, 0) <> 0 Then
        For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
            If strProds = "" Then strProds = CBA_CBISarr(0, a) Else strProds = strProds & ", " & CBA_CBISarr(0, a)
        Next
    End If
    CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_QuantityCount(COMRADE)", , , , , , , strProds
    If CBA_CBISarr(0, 0) <> 0 Then
        inArray = CBA_CBISarr
        CBA_COM_genCBISoutput True, inArray
    End If
    Set dicContainers = New Scripting.Dictionary
    For a = LBound(CBA_COM_CBISarrOutput, 2) To UBound(CBA_COM_CBISarrOutput, 2)
        dicContainers.Add CBA_COM_CBISarrOutput(1, a), CBA_COM_CBISarrOutput(3, a)
    Next
    
    
    
    PivNo = -2
    Set wbk = Workbooks.Add
    Set sht = ActiveSheet
    For a = wbk.Worksheets.Count + 1 To 6
        wbk.Worksheets.Add
    Next
    cnt = 0
    For Each wks In wbk.Worksheets
        With wks
            cnt = cnt + 1
            Select Case cnt
                Case 1
                    wks.Name = "Summary"
                    wbk.VBProject.VBComponents(wks.CodeName).Name = "wks_Sum"
                    .Cells(4, 3).Value = "CDS Report (Summary)"
                    Set wks_sum = wks
                Case 2
                    wks.Name = "NSW"
                    wbk.VBProject.VBComponents(wks.CodeName).Name = "wks_NSW"
                    .Cells(4, 3).Value = "CDS NSW Report"
                    Set wks_NSW = wks
                Case 3
                    wks.Name = "ACT"
                    wbk.VBProject.VBComponents(wks.CodeName).Name = "wks_ACT"
                    .Cells(4, 3).Value = "CDS ACT Report"
                    Set wks_ACT = wks
                Case 4
                    wks.Name = "QLD"
                    wbk.VBProject.VBComponents(wks.CodeName).Name = "wks_QLD"
                    .Cells(4, 3).Value = "CDS QLD Report"
                    Set wks_QLD = wks
                Case 5
                    wks.Name = "QLD-Exp"
                    wbk.VBProject.VBComponents(wks.CodeName).Name = "wks_EXP_QLD"
                    .Cells(4, 3).Value = "CDS QLD-Export Report"
                    Set wks_EXP_QLD = wks
                Case 6
                    wks.Name = "ACT-Exp"
                    wbk.VBProject.VBComponents(wks.CodeName).Name = "wks_EXP_ACT"
                    .Cells(4, 3).Value = "CDS ACT-Export Report"
                    Set wks_EXP_ACT = wks
            End Select
            Range(.Cells(1, 1), .Cells(5, 79)).Interior.ColorIndex = 49
            .Activate
            .Cells(1, 1).Select
            .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
            .Cells.Font.Name = "ALDI SUED Office"
            .Cells(4, 3).Font.Size = 24
            .Cells(4, 3).Font.ColorIndex = 2
            .Cells(4, 8).Value = CDS_DateFrom
            .Cells(4, 9).Value = CDS_DateTo
            Range(.Cells(4, 8), .Cells(4, 9)).Font.ColorIndex = 2
        End With
    Next
    
If CDS_NSWSelected = True Then
    Application.ScreenUpdating = True
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Building NSW Data..."
    Application.ScreenUpdating = False
    
    With wks_NSW
        strConts = ""
        For a = LBound(CDSBase, 2) To UBound(CDSBase, 2)
            If InStr(1, CDSBase(3, a), "NSW") > 0 And CDSBase(7, a) = True Then
                If strConts = "" Then strConts = CDSBase(1, a) Else strConts = strConts & ", " & CDSBase(1, a)
            End If
        Next
        For div = 501 To 509
            If div = 508 Then div = 509
            'CDS_DateFrom = #6/1/2019#: CDS_DateTo = #6/30/2019#
            CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_StoreReceiving", CDS_DateFrom, CDS_DateTo, div, , , , strConts
            Set sR(div) = New Scripting.Dictionary: Set ProdCol = New Scripting.Dictionary
            Set Storecol = New Scripting.Dictionary: Set Datecol = New Scripting.Dictionary
            curProd = CBA_MMSarr(0, 0): curDate = CBA_MMSarr(1, 0): curStore = CBA_MMSarr(2, 0)
            DateQTY = 0: ProdQTY = 0
            For a = LBound(CBA_MMSarr, 2) To UBound(CBA_MMSarr, 2)
                If CBA_MMSarr(0, a) <> curProd Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    DateQTY = 0
                    ProdCol.Add curDate, Datecol
                    ProdCol.Add "totQTY", CStr(ProdQTY)
                    ProdQTY = 0
                    sR(div).Add CStr(curProd), ProdCol
                    Set ProdCol = New Scripting.Dictionary
                    curProd = CBA_MMSarr(0, a)
                    Set Datecol = New Scripting.Dictionary
                    curDate = CBA_MMSarr(1, a)
                End If
                If CBA_MMSarr(1, a) <> curDate Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    ProdCol.Add curDate, Datecol
                    curDate = CBA_MMSarr(1, a)
                    Set Datecol = New Scripting.Dictionary
                    DateQTY = 0
                End If
                Datecol.Add CStr(CBA_MMSarr(2, a)), CStr(CBA_MMSarr(3, a))
                DateQTY = DateQTY + CBA_MMSarr(3, a)
                ProdQTY = ProdQTY + CBA_MMSarr(3, a)
                If a = UBound(CBA_MMSarr, 2) Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    ProdCol.Add curDate, Datecol
                    ProdCol.Add "totQTY", CStr(ProdQTY)
                    sR(div).Add CStr(curProd), ProdCol
                End If
            Next
        Next

        CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_CBIS_IncoTerms", CDS_DateFrom, CDS_DateTo, div, , , , strConts
        ReDim Inco(0 To UBound(CBA_CBISarr, 2))
        For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
            Inco(a).id = a
            Inco(a).ContractNo = CBA_CBISarr(0, a)
            Inco(a).DivNo = CBA_CBISarr(1, a)
            Inco(a).ValidFrom = CBA_CBISarr(2, a)
            Inco(a).ValidTo = CBA_CBISarr(3, a)
            Inco(a).IncoTerm = CBA_CBISarr(4, a)
        Next
        Set col502 = New Collection
        col502.Add 32: col502.Add 75
        Set col503 = New Collection
        col503.Add 10: col503.Add 21: col503.Add 22: col503.Add 23: col503.Add 26: col503.Add 45: col503.Add 54
        col503.Add 72: col503.Add 81: col503.Add 102: col503.Add 110
        Set col504 = New Collection
        col504.Add 1:   col504.Add 2:   col504.Add 3:   col504.Add 4:   col504.Add 5:   col504.Add 6:   col504.Add 7
        col504.Add 8:   col504.Add 9:   col504.Add 11:  col504.Add 12:  col504.Add 14:  col504.Add 15:  col504.Add 16
        col504.Add 18:  col504.Add 21:  col504.Add 22:  col504.Add 23:  col504.Add 24:  col504.Add 25:  col504.Add 26
        col504.Add 27:  col504.Add 28:  col504.Add 31:  col504.Add 32:  col504.Add 33:  col504.Add 34:  col504.Add 35
        col504.Add 37:  col504.Add 38:  col504.Add 39:  col504.Add 40:  col504.Add 41:  col504.Add 42:  col504.Add 43
        col504.Add 44:  col504.Add 45:  col504.Add 46:  col504.Add 47:  col504.Add 48:  col504.Add 49:  col504.Add 51
        col504.Add 52:  col504.Add 53:  col504.Add 54:  col504.Add 55:  col504.Add 57:  col504.Add 58:  col504.Add 59
        col504.Add 60:  col504.Add 61:  col504.Add 62:  col504.Add 63:  col504.Add 64:  col504.Add 65:  col504.Add 66
        col504.Add 67:  col504.Add 68:  col504.Add 69:  col504.Add 70:  col504.Add 71:  col504.Add 72:  col504.Add 73
        col504.Add 75:  col504.Add 77:  col504.Add 78:  col504.Add 79:  col504.Add 80:  col504.Add 81:  col504.Add 82
        col504.Add 83:  col504.Add 84:  col504.Add 85:  col504.Add 86:  col504.Add 87:  col504.Add 91:  col504.Add 92
        col504.Add 93:  col504.Add 95:  col504.Add 96:  col504.Add 97:  col504.Add 98:  col504.Add 102: col504.Add 103
        col504.Add 104: col504.Add 105: col504.Add 106: col504.Add 107: col504.Add 108: col504.Add 109: col504.Add 110
        col504.Add 111: col504.Add 113: col504.Add 114: col504.Add 115: col504.Add 116: col504.Add 117: col504.Add 118
        col504.Add 119: col504.Add 144: col504.Add 145: col504.Add 146: col504.Add 147: col504.Add 148: col504.Add 180
        col504.Add 181: col504.Add 182: col504.Add 183
        cnt = 0
        
        
        
        
        Set NSWCol = New Collection
        For a = LBound(CDSBase, 2) To UBound(CDSBase, 2)
'            If CBISDic(CStr(CDSBase(1, a))).ProductCode = 10357 Then
'            a = a
'            End If
            If InStr(1, CDSBase(3, a), "NSW") > 0 And CDSBase(6, a) = True Then
                NSWCol.Add a
            End If
        Next
        
        rw = -1
        ReDim CDS_NSW(0 To 17, 0 To NSWCol.Count)
        For Each a In NSWCol
                rw = rw + 1
                CDS_NSW(0, rw) = CDSBase(2, a)    'Region
                CDS_NSW(1, rw) = CDSBase(1, a)    'ContractNo
                CDS_NSW(2, rw) = CBISDic(CStr(CDSBase(1, a))).productcode
                PCode = CDS_NSW(2, rw)
                CDS_NSW(3, rw) = CBISDic(CStr(CDSBase(1, a))).Description
                If CBISDic(CStr(CDSBase(1, a))).Branded = 1 Then CDS_NSW(4, rw) = "Yes" Else CDS_NSW(4, rw) = "No"
                CDS_NSW(5, rw) = "NSW"
                CDS_NSW(6, rw) = CBISDic(CStr(CDSBase(1, a))).SupplierNo
                CDS_NSW(7, rw) = CBISDic(CStr(CDSBase(1, a))).SupplierName
                CDS_NSW(8, rw) = CDSBase(4, a)     'ContainerType
                DFrom = #1/1/1901#: Dto = DFrom
'                If PCode = 10357 Then
'                    a = a
'                End If
                For i = LBound(Inco) To UBound(Inco)
                    If Inco(i).ContractNo = CDSBase(1, a) And Inco(i).DivNo = CDSBase(2, a) And Inco(i).IncoTerm = CDSBase(5, a) Then
                        If Inco(i).ValidFrom < CDS_NSW(9, rw) Or IsEmpty(CDS_NSW(9, rw)) Then
                            CDS_NSW(9, rw) = Inco(i).ValidFrom
                            DFrom = Inco(i).ValidFrom
                        End If
                        If Inco(i).ValidTo > CDS_NSW(10, rw) Or IsEmpty(CDS_NSW(10, rw)) Then
                            CDS_NSW(10, rw) = Inco(i).ValidTo
                            Dto = Inco(i).ValidTo
                        End If
                        If DFrom = CDS_DateFrom And Dto = CDS_DateTo Then Exit For
                    End If
                Next
                If IsEmpty(CDS_NSW(9, rw)) And IsEmpty(CDS_NSW(10, rw)) Then
                    rw = rw - 1
                    GoTo NSWSkip
                End If
                CDS_NSW(11, rw) = CDSBase(5, a)    'IncoTerms
                'CDS_NSW(12, Rw) = CBISDic(CStr(CDSBase(1, a))).Packsize
                CDS_NSW(12, rw) = dicContainers(CBISDic(CStr(CDSBase(1, a))).productcode)
                RCV = 0
                div = CDSBase(2, a)
                'On Error Resume Next
'                If PCode = 51441 And CDSBase(2, a) = 504 Then
'                    a = a
'                End If
                For DT = DFrom To Dto
                    If sR(CStr(div)).Exists(CStr(PCode)) = True Then
                        If sR(CStr(div))(CStr(PCode)).Exists(DT) Then
                            If div = 501 Then
                                If RCVDivProdContDic(CStr(div)).Exists(CStr(PCode)) Then
                                    If RCVDivProdContDic(CStr(div))(CStr(PCode))("Total") = 0 Then
                                        cnt = 0
                                        For b = LBound(CDSBase, 2) To UBound(CDSBase, 2)
                                            If CBISDic(CStr(CDSBase(1, b))).productcode = PCode And CDSBase(2, a) = CDSBase(2, b) And CDSBase(6, b) = True And CDSBase(7, b) = True And CDSBase(3, a) = CDSBase(3, b) Then
                                                cnt = cnt + 1
                                            End If
                                        Next
                                        RCV = RCV + sR(CStr(div))(CStr(PCode))(DT)("totQTY") / cnt
                                    ElseIf RCVDivProdContDic(CStr(div))(CStr(PCode))("Total") <= RCVDivProdContDic(CStr(div))(CStr(PCode))(CStr(CDSBase(1, a))) Then
                                        RCV = RCV + sR(CStr(div))(CStr(PCode))(DT)("totQTY")
                                    Else
                                        RCV = RCV + (sR(CStr(div))(CStr(PCode))(DT)("totQTY") * (RCVDivProdContDic(CStr(div))(CStr(PCode))(CStr(CDSBase(1, a))) / RCVDivProdContDic(CStr(div))(CStr(PCode))("Total")))
                                    End If
                                Else
                                    cnt = 0
                                    For b = LBound(CDSBase, 2) To UBound(CDSBase, 2)
                                        If CBISDic(CStr(CDSBase(1, b))).productcode = PCode And CDSBase(2, a) = CDSBase(2, b) And CDSBase(6, b) = True And CDSBase(7, b) = True And CDSBase(3, a) = CDSBase(3, b) Then
                                            cnt = cnt + 1
                                        End If
                                    Next
                                            If cnt = 0 Then cnt = 1
                                    
                                    RCV = RCV + sR(CStr(div))(CStr(PCode))(DT)("totQTY") / cnt
                                End If
                            ElseIf div = 502 Then
                                For Each st In col502
                                    CalcRCV RCV, PCode, div, st, DT, a
                                Next
                            ElseIf div = 503 Then
                                For Each st In col503
                                    CalcRCV RCV, PCode, div, st, DT, a
                                Next
                            ElseIf div = 504 Then
                                For Each st In col504
                                    CalcRCV RCV, PCode, div, st, DT, a
                                Next
                            End If
                        End If
                    End If
                Next
                'Err.Clear
                'On Error GoTo 0
                If RCVDivProdContDic(CStr(div)).Exists(CStr(PCode)) Then
                    CDS_NSW(13, rw) = RCVDivProdContDic(CStr(div))(CStr(PCode))("Total")
                    If RCVDivProdContDic(CStr(div))(CStr(PCode)).Exists(CStr(CDSBase(1, a))) Then
                        CDS_NSW(14, rw) = RCVDivProdContDic(CStr(div))(CStr(PCode))(CStr(CDSBase(1, a)))
                    Else
                        CDS_NSW(14, rw) = 0
                    End If
                Else
                    CDS_NSW(13, rw) = 0
                    CDS_NSW(14, rw) = 0
                End If
                CDS_NSW(15, rw) = RCV
                CDS_NSW(16, rw) = CDSBase(7, a)
                If CDSBase(7, a) = True Then
                    CDS_NSW(17, rw) = RCV * CDS_NSW(12, rw)
                Else
                    CDS_NSW(17, rw) = 0
                End If
                
NSWSkip:
        Next
        .Activate
        .Cells(6, 1).Value = "Region"
        .Cells(6, 2).Value = "Contractno"
        .Cells(6, 3).Value = "ProductCode"
        .Cells(6, 4).Value = "Product Description"
        .Cells(6, 5).Value = "Branded"
        .Cells(6, 6).Value = "SchemeID"
        .Cells(6, 7).Value = "SupplierNo"
        .Cells(6, 8).Value = "SupplierName"
        .Cells(6, 9).Value = "ContainerType"
        .Cells(6, 10).Value = "IncoFrom"
        .Cells(6, 11).Value = "IncoTo"
        .Cells(6, 12).Value = "IncoTerm"
        .Cells(6, 13).Value = "PackSize"
        .Cells(6, 14).Value = "Product RCV"
        .Cells(6, 15).Value = "Contract RCV"
        .Cells(6, 16).Value = "QTY Store Rcvd"
        .Cells(6, 17).Value = "ALDI Resp."
        .Cells(6, 18).Value = "Units To Report"
        
        Range(.Cells(6, 1), .Cells(6, 18)).Font.Bold = True
        Range(.Cells(7, 1), .Cells(7 + UBound(CDS_NSW, 2), 1 + UBound(CDS_NSW, 1))).Value2 = CBA_BasicFunctions.CBA_TransposeArray(CDS_NSW)
        
        PivNo = PivNo + 3
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            .Cells(6, 1).CurrentRegion, Version:=xlPivotTableVersion14).CreatePivotTable _
            TableDestination:=wks_sum.Cells(7, PivNo), TableName:="NSWPivot", DefaultVersion _
            :=xlPivotTableVersion14
        With wks_sum.PivotTables("NSWPivot").PivotFields("ContainerType")
            .Orientation = xlRowField
            .Position = 1
        End With
        wks_sum.PivotTables("NSWPivot").AddDataField wks_sum.PivotTables("NSWPivot").PivotFields("Units To Report"), "Total Quantity", xlSum
        Range(wks_sum.Cells(6, PivNo), wks_sum.Cells(6, PivNo + 1)).Merge
        wks_sum.Cells(6, PivNo).Value = "NSW"
    End With
End If
If CDS_ACTSelected = True Then
    Application.ScreenUpdating = True
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Building ACT Data..."
    Application.ScreenUpdating = False

    With wks_ACT
        strConts = ""
        For a = LBound(CDSBase, 2) To UBound(CDSBase, 2)
            If InStr(1, CDSBase(3, a), "ACT") > 0 And CDSBase(6, a) = True Then
                If strConts = "" Then strConts = CDSBase(1, a) Else strConts = strConts & ", " & CDSBase(1, a)
            End If
        Next
        For div = 501 To 509
            If div = 508 Then div = 509
            'CDS_DateFrom = #6/1/2019#: CDS_DateTo = #6/30/2019#
            CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_StoreReceiving", CDS_DateFrom, CDS_DateTo, div, , , , strConts
            Set sR(div) = New Scripting.Dictionary: Set ProdCol = New Scripting.Dictionary
            Set Storecol = New Scripting.Dictionary: Set Datecol = New Scripting.Dictionary
            curProd = CBA_MMSarr(0, 0): curDate = CBA_MMSarr(1, 0): curStore = CBA_MMSarr(2, 0)
            DateQTY = 0: ProdQTY = 0
            For a = LBound(CBA_MMSarr, 2) To UBound(CBA_MMSarr, 2)
                If CBA_MMSarr(0, a) <> curProd Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    DateQTY = 0
                    ProdCol.Add curDate, Datecol
                    ProdCol.Add "totQTY", CStr(ProdQTY)
                    ProdQTY = 0
                    sR(div).Add CStr(curProd), ProdCol
                    Set ProdCol = New Scripting.Dictionary
                    curProd = CBA_MMSarr(0, a)
                    Set Datecol = New Scripting.Dictionary
                    curDate = CBA_MMSarr(1, a)
                End If
                If CBA_MMSarr(1, a) <> curDate Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    ProdCol.Add curDate, Datecol
                    curDate = CBA_MMSarr(1, a)
                    Set Datecol = New Scripting.Dictionary
                    DateQTY = 0
                End If
                Datecol.Add CStr(CBA_MMSarr(2, a)), CStr(CBA_MMSarr(3, a))
                DateQTY = DateQTY + CBA_MMSarr(3, a)
                ProdQTY = ProdQTY + CBA_MMSarr(3, a)
                If a = UBound(CBA_MMSarr, 2) Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    ProdCol.Add curDate, Datecol
                    ProdCol.Add "totQTY", CStr(ProdQTY)
                    sR(div).Add CStr(curProd), ProdCol
                End If
            Next
        Next
        CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_CBIS_IncoTerms", CDS_DateFrom, CDS_DateTo, div, , , , strConts
        ReDim Inco(0 To UBound(CBA_CBISarr, 2))
        For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
            Inco(a).id = a
            Inco(a).ContractNo = CBA_CBISarr(0, a)
            Inco(a).DivNo = CBA_CBISarr(1, a)
            Inco(a).ValidFrom = CBA_CBISarr(2, a)
            Inco(a).ValidTo = CBA_CBISarr(3, a)
            Inco(a).IncoTerm = CBA_CBISarr(4, a)
        Next
        Set col504 = New Collection
        col504.Add 13:   col504.Add 17:   col504.Add 19:   col504.Add 20:   col504.Add 29:   col504.Add 30:   col504.Add 36
        col504.Add 50:   col504.Add 56:   col504.Add 74:  col504.Add 76
        cnt = 0
        Set ACTCol = New Collection
        For a = LBound(CDSBase, 2) To UBound(CDSBase, 2)
            If InStr(1, CDSBase(3, a), "ACT") > 0 And CDSBase(7, a) = True Then ACTCol.Add a
        Next
        rw = -1
        ReDim CDS_ACT(0 To 17, 0 To ACTCol.Count)
        For Each a In ACTCol
            rw = rw + 1
            CDS_ACT(0, rw) = CDSBase(2, a)    'Region
            CDS_ACT(1, rw) = CDSBase(1, a)    'ContractNo
            CDS_ACT(2, rw) = CBISDic(CStr(CDSBase(1, a))).productcode
            PCode = CDS_ACT(2, rw)
            CDS_ACT(3, rw) = CBISDic(CStr(CDSBase(1, a))).Description
            If CBISDic(CStr(CDSBase(1, a))).Branded = 1 Then CDS_ACT(4, rw) = "Yes" Else CDS_ACT(4, rw) = "No"
            CDS_ACT(5, rw) = "ACT"
            CDS_ACT(6, rw) = CBISDic(CStr(CDSBase(1, a))).SupplierNo
            CDS_ACT(7, rw) = CBISDic(CStr(CDSBase(1, a))).SupplierName
            CDS_ACT(8, rw) = CDSBase(4, a)     'ContainerType
            DFrom = #1/1/1901#: Dto = DFrom
            For i = LBound(Inco) To UBound(Inco)
                If Inco(i).ContractNo = CDSBase(1, a) And Inco(i).DivNo = CDSBase(2, a) And Inco(i).IncoTerm = CDSBase(5, a) Then
                    If Inco(i).ValidFrom < CDS_ACT(9, rw) Or IsEmpty(CDS_ACT(9, rw)) Then
                        CDS_ACT(9, rw) = Inco(i).ValidFrom
                        DFrom = Inco(i).ValidFrom
                    End If
                    If Inco(i).ValidTo > CDS_ACT(10, rw) Or IsEmpty(CDS_ACT(10, rw)) Then
                        CDS_ACT(10, rw) = Inco(i).ValidTo
                        Dto = Inco(i).ValidTo
                    End If
                    If DFrom = CDS_DateFrom And Dto = CDS_DateTo Then Exit For
                End If
            Next
            If IsEmpty(CDS_ACT(9, rw)) And IsEmpty(CDS_ACT(10, rw)) Then
                rw = rw - 1
                GoTo ACTSkip
            End If
            CDS_ACT(11, rw) = CDSBase(5, a)    'IncoTerms
            CDS_ACT(12, rw) = dicContainers(CBISDic(CStr(CDSBase(1, a))).productcode)
            RCV = 0
            div = CDSBase(2, a)
            For DT = DFrom To Dto
                If sR(CStr(div)).Exists(CStr(PCode)) = True Then
                    If sR(CStr(div))(CStr(PCode)).Exists(DT) Then
                        If div = 504 Then
'                                If PCode = 51441 Then
'                                a = a
'                                End If
                                For Each st In col504
                                    CalcRCV RCV, PCode, div, st, DT, a
                                Next
                        End If
                    End If
                End If
            Next
            If RCVDivProdContDic(CStr(div)).Exists(CStr(PCode)) Then
                CDS_ACT(13, rw) = RCVDivProdContDic(CStr(div))(CStr(PCode))("Total")
                If RCVDivProdContDic(CStr(div))(CStr(PCode)).Exists(CStr(CDSBase(1, a))) Then
                    CDS_ACT(14, rw) = RCVDivProdContDic(CStr(div))(CStr(PCode))(CStr(CDSBase(1, a)))
                Else
                    CDS_ACT(14, rw) = 0
                End If
            Else
                CDS_ACT(13, rw) = 0
                CDS_ACT(14, rw) = 0
            End If
            CDS_ACT(15, rw) = RCV
            CDS_ACT(16, rw) = CDSBase(7, a)
            If CDSBase(7, a) = True Then
                CDS_ACT(17, rw) = RCV * CDS_ACT(12, rw)
            Else
                CDS_ACT(17, rw) = 0
            End If
ACTSkip:
        Next
        .Activate
        .Cells(6, 1).Value = "Region"
        .Cells(6, 2).Value = "Contractno"
        .Cells(6, 3).Value = "ProductCode"
        .Cells(6, 4).Value = "Product Description"
        .Cells(6, 5).Value = "Branded"
        .Cells(6, 6).Value = "SchemeID"
        .Cells(6, 7).Value = "SupplierNo"
        .Cells(6, 8).Value = "SupplierName"
        .Cells(6, 9).Value = "ContainerType"
        .Cells(6, 10).Value = "IncoFrom"
        .Cells(6, 11).Value = "IncoTo"
        .Cells(6, 12).Value = "IncoTerm"
        .Cells(6, 13).Value = "PackSize"
        .Cells(6, 14).Value = "Product RCV"
        .Cells(6, 15).Value = "Contract RCV"
        .Cells(6, 16).Value = "QTY Store Rcvd"
        .Cells(6, 17).Value = "ALDI Resp."
        .Cells(6, 18).Value = "Units To Report"
        Range(.Cells(6, 1), .Cells(6, 18)).Font.Bold = True
        Range(.Cells(7, 1), .Cells(7 + UBound(CDS_ACT, 2), 1 + UBound(CDS_ACT, 1))).Value2 = CBA_BasicFunctions.CBA_TransposeArray(CDS_ACT)
    
        PivNo = PivNo + 3
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            .Cells(6, 1).CurrentRegion, Version:=xlPivotTableVersion14).CreatePivotTable _
            TableDestination:=wks_sum.Cells(7, PivNo), TableName:="ACTPivot", DefaultVersion _
            :=xlPivotTableVersion14
        With wks_sum.PivotTables("ACTPivot").PivotFields("ContainerType")
            .Orientation = xlRowField
            .Position = 1
        End With
        wks_sum.PivotTables("ACTPivot").AddDataField wks_sum.PivotTables("ACTPivot").PivotFields("Units To Report"), "Total Quantity", xlSum
        Range(wks_sum.Cells(6, PivNo), wks_sum.Cells(6, PivNo + 1)).Merge
        wks_sum.Cells(6, PivNo).Value = "ACT"
    
    End With

End If
If CDS_EXP_ACTSelected = True Then
    Application.ScreenUpdating = True
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Building ACT Export Data..."
    Application.ScreenUpdating = False
    
    With wks_EXP_ACT
        strConts = ""
        For a = LBound(CDSBase, 2) To UBound(CDSBase, 2)
            If InStr(1, CDSBase(3, a), "NSW") > 0 And CDSBase(7, a) = False And CDSBase(2, a) = 504 Then
                If strConts = "" Then strConts = CDSBase(1, a) Else strConts = strConts & ", " & CDSBase(1, a)
            End If
        Next
        For div = 501 To 509
            If div = 508 Then div = 509
            'CDS_DateFrom = #6/1/2019#: CDS_DateTo = #6/30/2019#
            CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_StoreReceiving", CDS_DateFrom, CDS_DateTo, div, , , , strConts
            Set sR(div) = New Scripting.Dictionary: Set ProdCol = New Scripting.Dictionary
            Set Storecol = New Scripting.Dictionary: Set Datecol = New Scripting.Dictionary
            curProd = CBA_MMSarr(0, 0): curDate = CBA_MMSarr(1, 0): curStore = CBA_MMSarr(2, 0)
            DateQTY = 0: ProdQTY = 0
            For a = LBound(CBA_MMSarr, 2) To UBound(CBA_MMSarr, 2)
                If CBA_MMSarr(0, a) <> curProd Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    DateQTY = 0
                    ProdCol.Add curDate, Datecol
                    ProdCol.Add "totQTY", CStr(ProdQTY)
                    ProdQTY = 0
                    sR(div).Add CStr(curProd), ProdCol
                    Set ProdCol = New Scripting.Dictionary
                    curProd = CBA_MMSarr(0, a)
                    Set Datecol = New Scripting.Dictionary
                    curDate = CBA_MMSarr(1, a)
                End If
                If CBA_MMSarr(1, a) <> curDate Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    ProdCol.Add curDate, Datecol
                    curDate = CBA_MMSarr(1, a)
                    Set Datecol = New Scripting.Dictionary
                    DateQTY = 0
                End If
                Datecol.Add CStr(CBA_MMSarr(2, a)), CStr(CBA_MMSarr(3, a))
                DateQTY = DateQTY + CBA_MMSarr(3, a)
                ProdQTY = ProdQTY + CBA_MMSarr(3, a)
                If a = UBound(CBA_MMSarr, 2) Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    ProdCol.Add curDate, Datecol
                    ProdCol.Add "totQTY", CStr(ProdQTY)
                    sR(div).Add CStr(curProd), ProdCol
                End If
            Next
        Next
        CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_CBIS_IncoTerms", CDS_DateFrom, CDS_DateTo, div, , , , strConts
        ReDim Inco(0 To UBound(CBA_CBISarr, 2))
        For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
            Inco(a).id = a
            Inco(a).ContractNo = CBA_CBISarr(0, a)
            Inco(a).DivNo = CBA_CBISarr(1, a)
            Inco(a).ValidFrom = CBA_CBISarr(2, a)
            Inco(a).ValidTo = CBA_CBISarr(3, a)
            Inco(a).IncoTerm = CBA_CBISarr(4, a)
        Next
        
        Set col504 = New Collection
        col504.Add 13:   col504.Add 17:   col504.Add 19:   col504.Add 20:   col504.Add 29:   col504.Add 30:   col504.Add 36
        col504.Add 50:   col504.Add 56:   col504.Add 74:  col504.Add 76
        cnt = 0
        Set EXP_ACTCol = New Collection
        For a = LBound(CDSBase, 2) To UBound(CDSBase, 2)
            If InStr(1, CDSBase(3, a), "NSW") > 0 And CDSBase(6, a) = True And CDSBase(2, a) = 504 Then EXP_ACTCol.Add a
        Next
        rw = -1
        ReDim CDS_EXP_ACT(0 To 17, 0 To EXP_ACTCol.Count)
        For Each a In EXP_ACTCol
                rw = rw + 1
                CDS_EXP_ACT(0, rw) = CDSBase(2, a)    'Region
                CDS_EXP_ACT(1, rw) = CDSBase(1, a)    'Contractno
                CDS_EXP_ACT(2, rw) = CBISDic(CStr(CDSBase(1, a))).productcode
                PCode = CDS_EXP_ACT(2, rw)
                CDS_EXP_ACT(3, rw) = CBISDic(CStr(CDSBase(1, a))).Description
                If CBISDic(CStr(CDSBase(1, a))).Branded = 1 Then CDS_EXP_ACT(4, rw) = "Yes" Else CDS_EXP_ACT(4, rw) = "No"
                CDS_EXP_ACT(5, rw) = "EXP_ACT"
                CDS_EXP_ACT(6, rw) = CBISDic(CStr(CDSBase(1, a))).SupplierNo
                CDS_EXP_ACT(7, rw) = CBISDic(CStr(CDSBase(1, a))).SupplierName
                CDS_EXP_ACT(8, rw) = CDSBase(4, a)     'ContainerType
                DFrom = #1/1/1901#: Dto = DFrom
                For i = LBound(Inco) To UBound(Inco)
                    If Inco(i).ContractNo = CDSBase(1, a) And Inco(i).DivNo = CDSBase(2, a) And Inco(i).IncoTerm = CDSBase(5, a) Then
                        If Inco(i).ValidFrom < CDS_EXP_ACT(9, rw) Or IsEmpty(CDS_EXP_ACT(9, rw)) Then
                            CDS_EXP_ACT(9, rw) = Inco(i).ValidFrom
                            DFrom = Inco(i).ValidFrom
                        End If
                        If Inco(i).ValidTo > CDS_EXP_ACT(10, rw) Or IsEmpty(CDS_EXP_ACT(10, rw)) Then
                            CDS_EXP_ACT(10, rw) = Inco(i).ValidTo
                            Dto = Inco(i).ValidTo
                        End If
                        If DFrom = CDS_DateFrom And Dto = CDS_DateTo Then Exit For
                    End If
                Next
                If IsEmpty(CDS_EXP_ACT(9, rw)) And IsEmpty(CDS_EXP_ACT(10, rw)) Then
                    rw = rw - 1
                    GoTo ACTEXPSkip
                End If
                CDS_EXP_ACT(11, rw) = CDSBase(5, a)    'IncoTerms
                CDS_EXP_ACT(12, rw) = dicContainers(CBISDic(CStr(CDSBase(1, a))).productcode)
                RCV = 0
                div = CDSBase(2, a)
                For DT = DFrom To Dto
                    If sR(CStr(div)).Exists(CStr(PCode)) = True Then
                        If sR(CStr(div))(CStr(PCode)).Exists(DT) Then
                            If div = 504 Then
                                For Each st In col504
                                    CalcRCV RCV, PCode, div, st, DT, a
                                Next
                            End If
                        End If
                    End If
                Next
                If RCVDivProdContDic(CStr(div)).Exists(CStr(PCode)) Then
                    CDS_EXP_ACT(13, rw) = RCVDivProdContDic(CStr(div))(CStr(PCode))("Total")
                    If RCVDivProdContDic(CStr(div))(CStr(PCode)).Exists(CStr(CDSBase(1, a))) Then
                        CDS_EXP_ACT(14, rw) = RCVDivProdContDic(CStr(div))(CStr(PCode))(CStr(CDSBase(1, a)))
                    Else
                        CDS_EXP_ACT(14, rw) = 0
                    End If
                Else
                    CDS_EXP_ACT(13, rw) = 0
                    CDS_EXP_ACT(14, rw) = 0
                End If
                CDS_EXP_ACT(15, rw) = RCV
                CDS_EXP_ACT(16, rw) = CDSBase(7, a)
                If CDSBase(7, a) = False Then
                    CDS_EXP_ACT(17, rw) = RCV * CDS_EXP_ACT(12, rw)
                Else
                    CDS_EXP_ACT(17, rw) = 0
                End If
ACTEXPSkip:
        Next
        .Activate
        .Cells(6, 1).Value = "Region"
        .Cells(6, 2).Value = "Contractno"
        .Cells(6, 3).Value = "ProductCode"
        .Cells(6, 4).Value = "Product Description"
        .Cells(6, 5).Value = "Branded"
        .Cells(6, 6).Value = "SchemeID"
        .Cells(6, 7).Value = "SupplierNo"
        .Cells(6, 8).Value = "SupplierName"
        .Cells(6, 9).Value = "ContainerType"
        .Cells(6, 10).Value = "IncoFrom"
        .Cells(6, 11).Value = "IncoTo"
        .Cells(6, 12).Value = "IncoTerm"
        .Cells(6, 13).Value = "PackSize"
        .Cells(6, 14).Value = "Product RCV"
        .Cells(6, 15).Value = "Contract RCV"
        .Cells(6, 16).Value = "QTY Store Rcvd"
        .Cells(6, 17).Value = "ALDI Resp."
        .Cells(6, 18).Value = "Units To Report"
        Range(.Cells(6, 1), .Cells(6, 15)).Font.Bold = True
        Range(.Cells(7, 1), .Cells(7 + UBound(CDS_EXP_ACT, 2), 1 + UBound(CDS_EXP_ACT, 1))).Value2 = CBA_BasicFunctions.CBA_TransposeArray(CDS_EXP_ACT)
    
    PivNo = PivNo + 3
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        .Cells(6, 1).CurrentRegion, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:=wks_sum.Cells(7, PivNo), TableName:="EXT_ACTPivot", DefaultVersion _
        :=xlPivotTableVersion14
    With wks_sum.PivotTables("EXT_ACTPivot").PivotFields("ContainerType")
        .Orientation = xlRowField
        .Position = 1
    End With
    wks_sum.PivotTables("EXT_ACTPivot").AddDataField wks_sum.PivotTables("EXT_ACTPivot").PivotFields("Units To Report"), "Total Quantity", xlSum
    Range(wks_sum.Cells(6, PivNo), wks_sum.Cells(6, PivNo + 1)).Merge
    wks_sum.Cells(6, PivNo).Value = "EXT_ACT"
        
    End With
End If


If CDS_QLDSelected = True Then
    Application.ScreenUpdating = True
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Building QLD Data..."
    Application.ScreenUpdating = False
    With wks_QLD
        strConts = ""
        For a = LBound(CDSBase, 2) To UBound(CDSBase, 2)
            If InStr(1, CDSBase(3, a), "QLD") > 0 And CDSBase(6, a) = True Then
                If strConts = "" Then strConts = CDSBase(1, a) Else strConts = strConts & ", " & CDSBase(1, a)
            End If
        Next
        For div = 501 To 509
            If div = 508 Then div = 509
            'CDS_DateFrom = #6/1/2019#: CDS_DateTo = #6/30/2019#
            CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_StoreReceiving", CDS_DateFrom, CDS_DateTo, div, , , , strConts
            Set sR(div) = New Scripting.Dictionary: Set ProdCol = New Scripting.Dictionary
            Set Storecol = New Scripting.Dictionary: Set Datecol = New Scripting.Dictionary
            curProd = CBA_MMSarr(0, 0): curDate = CBA_MMSarr(1, 0): curStore = CBA_MMSarr(2, 0)
            DateQTY = 0: ProdQTY = 0
            For a = LBound(CBA_MMSarr, 2) To UBound(CBA_MMSarr, 2)
                If CBA_MMSarr(0, a) <> curProd Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    DateQTY = 0
                    ProdCol.Add curDate, Datecol
                    ProdCol.Add "totQTY", CStr(ProdQTY)
                    ProdQTY = 0
                    sR(div).Add CStr(curProd), ProdCol
                    Set ProdCol = New Scripting.Dictionary
                    curProd = CBA_MMSarr(0, a)
                    Set Datecol = New Scripting.Dictionary
                    curDate = CBA_MMSarr(1, a)
                End If
                If CBA_MMSarr(1, a) <> curDate Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    ProdCol.Add curDate, Datecol
                    curDate = CBA_MMSarr(1, a)
                    Set Datecol = New Scripting.Dictionary
                    DateQTY = 0
                End If
                Datecol.Add CStr(CBA_MMSarr(2, a)), CStr(CBA_MMSarr(3, a))
                DateQTY = DateQTY + CBA_MMSarr(3, a)
                ProdQTY = ProdQTY + CBA_MMSarr(3, a)
                If a = UBound(CBA_MMSarr, 2) Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    ProdCol.Add curDate, Datecol
                    ProdCol.Add "totQTY", CStr(ProdQTY)
                    sR(div).Add CStr(curProd), ProdCol
                End If
            Next
        Next
        CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_CBIS_IncoTerms", CDS_DateFrom, CDS_DateTo, div, , , , strConts
        ReDim Inco(0 To UBound(CBA_CBISarr, 2))
        For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
            Inco(a).id = a
            Inco(a).ContractNo = CBA_CBISarr(0, a)
            Inco(a).DivNo = CBA_CBISarr(1, a)
            Inco(a).ValidFrom = CBA_CBISarr(2, a)
            Inco(a).ValidTo = CBA_CBISarr(3, a)
            Inco(a).IncoTerm = CBA_CBISarr(4, a)
        Next
        Set col503 = New Collection
        col503.Add 1:   col503.Add 2:   col503.Add 3:   col503.Add 4:   col503.Add 5:   col503.Add 6
        col503.Add 7:   col503.Add 8:   col503.Add 9:   col503.Add 11:  col503.Add 12:  col503.Add 13
        col503.Add 14:  col503.Add 15:  col503.Add 16:  col503.Add 17:  col503.Add 18:  col503.Add 19
        col503.Add 20:  col503.Add 24:  col503.Add 25:  col503.Add 27:  col503.Add 28:  col503.Add 30
        col503.Add 32:  col503.Add 33:  col503.Add 34:  col503.Add 38:  col503.Add 42:  col503.Add 43
        col503.Add 44:  col503.Add 46:  col503.Add 51:  col503.Add 52:  col503.Add 53:  col503.Add 57
        col503.Add 60:  col503.Add 63:  col503.Add 65:  col503.Add 68:  col503.Add 71:  col503.Add 73
        col503.Add 74:  col503.Add 75:  col503.Add 78:  col503.Add 84:  col503.Add 85:  col503.Add 92
        col503.Add 94:  col503.Add 95:  col503.Add 98:  col503.Add 101: col503.Add 105: col503.Add 106
        col503.Add 107: col503.Add 108: col503.Add 109: col503.Add 111: col503.Add 112
        
        Set col506 = New Collection
        col506.Add 1:   col506.Add 3:   col506.Add 4:   col506.Add 5:   col506.Add 8:   col506.Add 9:   col506.Add 10
        col506.Add 11:  col506.Add 12:  col506.Add 14:  col506.Add 15:  col506.Add 16:  col506.Add 17:  col506.Add 18
        col506.Add 19:  col506.Add 20:  col506.Add 21:  col506.Add 22:  col506.Add 23:  col506.Add 27:  col506.Add 28
        col506.Add 31:  col506.Add 35:  col506.Add 36:  col506.Add 37:  col506.Add 39:  col506.Add 40:  col506.Add 41
        col506.Add 47:  col506.Add 48:  col506.Add 50:  col506.Add 55:  col506.Add 56:  col506.Add 58:  col506.Add 59
        col506.Add 61:  col506.Add 62:  col506.Add 64:  col506.Add 66:  col506.Add 67:  col506.Add 69:  col506.Add 70
        col506.Add 76:  col506.Add 77:  col506.Add 79:  col506.Add 80:  col506.Add 82:  col506.Add 83:  col506.Add 86
        col506.Add 87:  col506.Add 88:  col506.Add 89:  col506.Add 91:  col506.Add 93:  col506.Add 96:  col506.Add 97
        col506.Add 103: col506.Add 104

        cnt = 0
        Set QLDCol = New Collection
        For a = LBound(CDSBase, 2) To UBound(CDSBase, 2)
            If InStr(1, CDSBase(3, a), "QLD") > 0 And CDSBase(7, a) = True Then QLDCol.Add a
        Next
        rw = -1
        ReDim CDS_QLD(0 To 17, 0 To QLDCol.Count)
        For Each a In QLDCol
            rw = rw + 1
            CDS_QLD(0, rw) = CDSBase(2, a)    'Region
            CDS_QLD(1, rw) = CDSBase(1, a)    'ContractNo
            CDS_QLD(2, rw) = CBISDic(CStr(CDSBase(1, a))).productcode
            PCode = CDS_QLD(2, rw)
            CDS_QLD(3, rw) = CBISDic(CStr(CDSBase(1, a))).Description
            If CBISDic(CStr(CDSBase(1, a))).Branded = 1 Then CDS_QLD(4, rw) = "Yes" Else CDS_QLD(4, rw) = "No"
            CDS_QLD(5, rw) = "QLD"
            CDS_QLD(6, rw) = CBISDic(CStr(CDSBase(1, a))).SupplierNo
            CDS_QLD(7, rw) = CBISDic(CStr(CDSBase(1, a))).SupplierName
            CDS_QLD(8, rw) = CDSBase(4, a)     'ContainerType
            DFrom = #1/1/1901#: Dto = DFrom
            For i = LBound(Inco) To UBound(Inco)
                If Inco(i).ContractNo = CDSBase(1, a) And Inco(i).DivNo = CDSBase(2, a) And Inco(i).IncoTerm = CDSBase(5, a) Then
                    If Inco(i).ValidFrom < CDS_QLD(9, rw) Or IsEmpty(CDS_QLD(9, rw)) Then
                        CDS_QLD(9, rw) = Inco(i).ValidFrom
                        DFrom = Inco(i).ValidFrom
                    End If
                    If Inco(i).ValidTo > CDS_QLD(10, rw) Or IsEmpty(CDS_QLD(10, rw)) Then
                        CDS_QLD(10, rw) = Inco(i).ValidTo
                        Dto = Inco(i).ValidTo
                    End If
                    If DFrom = CDS_DateFrom And Dto = CDS_DateTo Then Exit For
                End If
            Next
            If IsEmpty(CDS_QLD(9, rw)) And IsEmpty(CDS_QLD(10, rw)) Then
                rw = rw - 1
                GoTo QLDSkip
            End If
            CDS_QLD(11, rw) = CDSBase(5, a)    'IncoTerms
            CDS_QLD(12, rw) = dicContainers(CBISDic(CStr(CDSBase(1, a))).productcode)
            RCV = 0
            div = CDSBase(2, a)
            For DT = DFrom To Dto
                If sR(CStr(div)).Exists(CStr(PCode)) = True Then
                    If sR(CStr(div))(CStr(PCode)).Exists(DT) Then
                        If div = 503 Then
                                For Each st In col503
                                    CalcRCV RCV, PCode, div, st, DT, a
                                Next
                        End If
                        If div = 506 Then
                                For Each st In col506
                                    CalcRCV RCV, PCode, div, st, DT, a
                                Next
                        End If
                    End If
                End If
            Next
            If RCVDivProdContDic(CStr(div)).Exists(CStr(PCode)) Then
                CDS_QLD(13, rw) = RCVDivProdContDic(CStr(div))(CStr(PCode))("Total")
                If RCVDivProdContDic(CStr(div))(CStr(PCode)).Exists(CStr(CDSBase(1, a))) Then
                    CDS_QLD(14, rw) = RCVDivProdContDic(CStr(div))(CStr(PCode))(CStr(CDSBase(1, a)))
                Else
                    CDS_QLD(14, rw) = 0
                End If
            Else
                CDS_QLD(13, rw) = 0
                CDS_QLD(14, rw) = 0
            End If
            CDS_QLD(15, rw) = RCV
            CDS_QLD(16, rw) = CDSBase(7, a)
            If CDSBase(7, a) = True Then
                CDS_QLD(17, rw) = RCV * CDS_QLD(12, rw)
            Else
                CDS_QLD(17, rw) = 0
            End If
QLDSkip:
        Next
        .Activate
        .Cells(6, 1).Value = "Region"
        .Cells(6, 2).Value = "Contractno"
        .Cells(6, 3).Value = "ProductCode"
        .Cells(6, 4).Value = "Product Description"
        .Cells(6, 5).Value = "Branded"
        .Cells(6, 6).Value = "SchemeID"
        .Cells(6, 7).Value = "SupplierNo"
        .Cells(6, 8).Value = "SupplierName"
        .Cells(6, 9).Value = "ContainerType"
        .Cells(6, 10).Value = "IncoFrom"
        .Cells(6, 11).Value = "IncoTo"
        .Cells(6, 12).Value = "IncoTerm"
        .Cells(6, 13).Value = "PackSize"
        .Cells(6, 14).Value = "Product RCV"
        .Cells(6, 15).Value = "Contract RCV"
        .Cells(6, 16).Value = "QTY Store Rcvd"
        .Cells(6, 17).Value = "ALDI Resp."
        .Cells(6, 18).Value = "Units To Report"
        Range(.Cells(6, 1), .Cells(6, 17)).Font.Bold = True
        Range(.Cells(7, 1), .Cells(7 + UBound(CDS_QLD, 2), 1 + UBound(CDS_QLD, 1))).Value2 = CBA_BasicFunctions.CBA_TransposeArray(CDS_QLD)
    
    PivNo = PivNo + 3
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        .Cells(6, 1).CurrentRegion, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:=wks_sum.Cells(7, PivNo), TableName:="QLDPivot", DefaultVersion _
        :=xlPivotTableVersion14
    With wks_sum.PivotTables("QLDPivot").PivotFields("ContainerType")
        .Orientation = xlRowField
        .Position = 1
    End With
    wks_sum.PivotTables("QLDPivot").AddDataField wks_sum.PivotTables("QLDPivot").PivotFields("Units To Report"), "Total Quantity", xlSum
    Range(wks_sum.Cells(6, PivNo), wks_sum.Cells(6, PivNo + 1)).Merge

        
    End With
End If

If CDS_EXP_QLDSelected = True Then
    Application.ScreenUpdating = True
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Building QLD Export Data..."
    Application.ScreenUpdating = False
    With wks_EXP_QLD
        strConts = ""
        For a = LBound(CDSBase, 2) To UBound(CDSBase, 2)
            If InStr(1, CDSBase(3, a), "QLD") > 0 And CDSBase(6, a) = True And CDSBase(2, a) = 503 Then
                If strConts = "" Then strConts = CDSBase(1, a) Else strConts = strConts & ", " & CDSBase(1, a)
            End If
        Next
        For div = 501 To 509
            If div = 508 Then div = 509
            'CDS_DateFrom = #6/1/2019#: CDS_DateTo = #6/30/2019#
            CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_StoreReceiving", CDS_DateFrom, CDS_DateTo, div, , , , strConts
            Set sR(div) = New Scripting.Dictionary: Set ProdCol = New Scripting.Dictionary
            Set Storecol = New Scripting.Dictionary: Set Datecol = New Scripting.Dictionary
            curProd = CBA_MMSarr(0, 0): curDate = CBA_MMSarr(1, 0): curStore = CBA_MMSarr(2, 0)
            DateQTY = 0: ProdQTY = 0
            For a = LBound(CBA_MMSarr, 2) To UBound(CBA_MMSarr, 2)
                If CBA_MMSarr(0, a) <> curProd Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    DateQTY = 0
                    ProdCol.Add curDate, Datecol
                    ProdCol.Add "totQTY", CStr(ProdQTY)
                    ProdQTY = 0
                    sR(div).Add CStr(curProd), ProdCol
                    Set ProdCol = New Scripting.Dictionary
                    curProd = CBA_MMSarr(0, a)
                    Set Datecol = New Scripting.Dictionary
                    curDate = CBA_MMSarr(1, a)
                End If
                If CBA_MMSarr(1, a) <> curDate Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    ProdCol.Add curDate, Datecol
                    curDate = CBA_MMSarr(1, a)
                    Set Datecol = New Scripting.Dictionary
                    DateQTY = 0
                End If
                Datecol.Add CStr(CBA_MMSarr(2, a)), CStr(CBA_MMSarr(3, a))
                DateQTY = DateQTY + CBA_MMSarr(3, a)
                ProdQTY = ProdQTY + CBA_MMSarr(3, a)
                If a = UBound(CBA_MMSarr, 2) Then
                    Datecol.Add "totQTY", CStr(DateQTY)
                    ProdCol.Add curDate, Datecol
                    ProdCol.Add "totQTY", CStr(ProdQTY)
                    sR(div).Add CStr(curProd), ProdCol
                End If
            Next
        Next
        CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_CBIS_IncoTerms", CDS_DateFrom, CDS_DateTo, div, , , , strConts
        ReDim Inco(0 To UBound(CBA_CBISarr, 2))
        For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
            Inco(a).id = a
            Inco(a).ContractNo = CBA_CBISarr(0, a)
            Inco(a).DivNo = CBA_CBISarr(1, a)
            Inco(a).ValidFrom = CBA_CBISarr(2, a)
            Inco(a).ValidTo = CBA_CBISarr(3, a)
            Inco(a).IncoTerm = CBA_CBISarr(4, a)
        Next
        Set col503 = New Collection
        col503.Add 10: col503.Add 21: col503.Add 22: col503.Add 23: col503.Add 26: col503.Add 45: col503.Add 54
        col503.Add 72: col503.Add 81: col503.Add 102: col503.Add 110
        cnt = 0
        Set EXP_QLDCol = New Collection
        For a = LBound(CDSBase, 2) To UBound(CDSBase, 2)
            If InStr(1, CDSBase(3, a), "QLD") > 0 And CDSBase(7, a) = False And CDSBase(2, a) = 503 Then EXP_QLDCol.Add a
        Next
        rw = -1
        ReDim CDS_EXP_QLD(0 To 17, 0 To EXP_QLDCol.Count)
        For Each a In EXP_QLDCol
                rw = rw + 1
                CDS_EXP_QLD(0, rw) = CDSBase(2, a)    'Region
                CDS_EXP_QLD(1, rw) = CDSBase(1, a)    'ContractNo
                CDS_EXP_QLD(2, rw) = CBISDic(CStr(CDSBase(1, a))).productcode
                PCode = CDS_EXP_QLD(2, rw)
                CDS_EXP_QLD(3, rw) = CBISDic(CStr(CDSBase(1, a))).Description
                If CBISDic(CStr(CDSBase(1, a))).Branded = 1 Then CDS_EXP_QLD(4, rw) = "Yes" Else CDS_EXP_QLD(4, rw) = "No"
                CDS_EXP_QLD(5, rw) = "EXP_QLD"
                CDS_EXP_QLD(6, rw) = CBISDic(CStr(CDSBase(1, a))).SupplierNo
                CDS_EXP_QLD(7, rw) = CBISDic(CStr(CDSBase(1, a))).SupplierName
                CDS_EXP_QLD(8, rw) = CDSBase(4, a)     'ContainerType
                DFrom = #1/1/1901#: Dto = DFrom
                For i = LBound(Inco) To UBound(Inco)
                    If Inco(i).ContractNo = CDSBase(1, a) And Inco(i).DivNo = CDSBase(2, a) And Inco(i).IncoTerm = CDSBase(5, a) Then
                        If Inco(i).ValidFrom < CDS_EXP_QLD(9, rw) Or IsEmpty(CDS_EXP_QLD(9, rw)) Then
                            CDS_EXP_QLD(9, rw) = Inco(i).ValidFrom
                            DFrom = Inco(i).ValidFrom
                        End If
                        If Inco(i).ValidTo > CDS_EXP_QLD(10, rw) Or IsEmpty(CDS_EXP_QLD(10, rw)) Then
                            CDS_EXP_QLD(10, rw) = Inco(i).ValidTo
                            Dto = Inco(i).ValidTo
                        End If
                        If DFrom = CDS_DateFrom And Dto = CDS_DateTo Then Exit For
                    End If
                Next
                If IsEmpty(CDS_EXP_QLD(9, rw)) And IsEmpty(CDS_EXP_QLD(10, rw)) Then
                    rw = rw - 1
                    GoTo QLDEXPSkip
                End If
                CDS_EXP_QLD(11, rw) = CDSBase(5, a)    'IncoTerms
                CDS_EXP_QLD(12, rw) = dicContainers(CBISDic(CStr(CDSBase(1, a))).productcode)
                RCV = 0
                div = CDSBase(2, a)
                For DT = DFrom To Dto
                    If sR(CStr(div)).Exists(CStr(PCode)) = True Then
                        If sR(CStr(div))(CStr(PCode)).Exists(DT) Then
                            If div = 503 Then
                                For Each st In col503
                                    CalcRCV RCV, PCode, div, st, DT, a
                                Next
                            End If
                        End If
                    End If
                Next
                If RCVDivProdContDic(CStr(div)).Exists(CStr(PCode)) Then
                    CDS_EXP_QLD(13, rw) = RCVDivProdContDic(CStr(div))(CStr(PCode))("Total")
                    If RCVDivProdContDic(CStr(div))(CStr(PCode)).Exists(CStr(CDSBase(1, a))) Then
                        CDS_EXP_QLD(14, rw) = RCVDivProdContDic(CStr(div))(CStr(PCode))(CStr(CDSBase(1, a)))
                    Else
                        CDS_EXP_QLD(14, rw) = 0
                    End If
                Else
                    CDS_EXP_QLD(13, rw) = 0
                    CDS_EXP_QLD(14, rw) = 0
                End If
                CDS_EXP_QLD(15, rw) = RCV
                CDS_EXP_QLD(16, rw) = CDSBase(7, a)
                If CDSBase(7, a) = False Then
                    CDS_EXP_QLD(17, rw) = RCV * CDS_EXP_QLD(12, rw)
                Else
                    CDS_EXP_QLD(17, rw) = 0
                End If
QLDEXPSkip:
        Next
        .Activate
        .Cells(6, 1).Value = "Region"
        .Cells(6, 2).Value = "Contractno"
        .Cells(6, 3).Value = "ProductCode"
        .Cells(6, 4).Value = "Product Description"
        .Cells(6, 5).Value = "Branded"
        .Cells(6, 6).Value = "SchemeID"
        .Cells(6, 7).Value = "SupplierNo"
        .Cells(6, 8).Value = "SupplierName"
        .Cells(6, 9).Value = "ContainerType"
        .Cells(6, 10).Value = "IncoFrom"
        .Cells(6, 11).Value = "IncoTo"
        .Cells(6, 12).Value = "IncoTerm"
        .Cells(6, 13).Value = "PackSize"
        .Cells(6, 14).Value = "Product RCV"
        .Cells(6, 15).Value = "Contract RCV"
        .Cells(6, 16).Value = "QTY Store Rcvd"
        .Cells(6, 17).Value = "ALDI Resp."
        .Cells(6, 18).Value = "Units To Report"
        Range(.Cells(6, 1), .Cells(6, 17)).Font.Bold = True
        Range(.Cells(7, 1), .Cells(7 + UBound(CDS_EXP_QLD, 2), 1 + UBound(CDS_EXP_QLD, 1))).Value2 = CBA_BasicFunctions.CBA_TransposeArray(CDS_EXP_QLD)
    
    PivNo = PivNo + 3
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        .Cells(6, 1).CurrentRegion, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:=wks_sum.Cells(7, PivNo), TableName:="EXP_QLDPivot", DefaultVersion _
        :=xlPivotTableVersion14
    With wks_sum.PivotTables("EXP_QLDPivot").PivotFields("ContainerType")
        .Orientation = xlRowField
        .Position = 1
    End With
    wks_sum.PivotTables("EXP_QLDPivot").AddDataField wks_sum.PivotTables("EXP_QLDPivot").PivotFields("Units To Report"), "Total Quantity", xlSum
    Range(wks_sum.Cells(6, PivNo), wks_sum.Cells(6, PivNo + 1)).Merge
    End With
End If
wks_sum.Activate
    Set but = wks_sum.Shapes.AddShape(msoShapeRoundedRectangle, 900, 20, 120, 50)
    With but.OLEFormat.Object
        .Caption = "Approve for Margin"
        .Font.Size = 16
        .Font.Name = "ALDI SUED Office"
        .HorizontalAlignment = xlCenter
        .OnAction = "CBA_CDS_Runtime.ApproveForMargin"
    End With
If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
Application.ScreenUpdating = True
End Sub
Private Function CalcRCV(ByRef RCV As Single, ByVal PCode As Long, ByVal div As Long, ByVal st As Long, ByVal DT As Date, ByVal a As Long)
Dim cnt As Long
Dim b As Long
    If sR(CStr(div))(CStr(PCode))(DT).Exists(CStr(st)) Then
        If RCVDivProdContDic(CStr(div)).Exists(CStr(PCode)) Then
            If RCVDivProdContDic(CStr(div))(CStr(PCode))("Total") <= CStr(NZ(RCVDivProdContDic(CStr(div))(CStr(PCode))(CStr(CDSBase(1, a))))) Then
                If RCVDivProdContDic(CStr(div))(CStr(PCode))("Total") = 0 Then
                    cnt = 0
                    For b = LBound(CDSBase, 2) To UBound(CDSBase, 2)
                        If CBISDic(CStr(CDSBase(1, b))).productcode = PCode And CDSBase(2, a) = CDSBase(2, b) And CDSBase(7, b) = True And CDSBase(6, b) = True And CDSBase(3, a) = CDSBase(3, b) Then
                            cnt = cnt + 1
                        End If
                    Next
                    RCV = RCV + sR(CStr(div))(CStr(PCode))(DT)(CStr(st)) / cnt
                Else
                    RCV = RCV + sR(CStr(div))(CStr(PCode))(DT)(CStr(st))
                End If
            Else
                RCV = RCV + (sR(CStr(div))(CStr(PCode))(DT)(CStr(st)) * (IIf(NZ(RCVDivProdContDic(CStr(div))(CStr(PCode))(CStr(CDSBase(1, a)))) < 0, 0, NZ(RCVDivProdContDic(CStr(div))(CStr(PCode))(CStr(CDSBase(1, a))))) / NZ(RCVDivProdContDic(CStr(div))(CStr(PCode))("Total"))))
            End If
        Else
            cnt = 0
            For b = LBound(CDSBase, 2) To UBound(CDSBase, 2)
                If CBISDic(CStr(CDSBase(1, b))).productcode = PCode And CDSBase(2, a) = CDSBase(2, b) And CDSBase(7, b) = True And CDSBase(6, b) = True And CDSBase(3, a) = CDSBase(3, b) Then
                    cnt = cnt + 1
                End If
            Next
            If cnt = 0 Then cnt = 1
            RCV = RCV + sR(CStr(div))(CStr(PCode))(DT)(CStr(st)) / cnt
        End If
    End If
End Function

Sub ApproveForMargin()
'Dim Piv As PivotTable
'Dim Data(1 To 5) As Variant
'Dim CN As ADODB.Connection
'Dim RS As ADODB.Recordset
'Dim d As Long
'
'
'    Set CTcol = New Collection: CTcol.Add "Aluminium": CTcol.Add "Glass": CTcol.Add "Glass -Amber / Brown"
'    CTcol.Add "Glass -Flint / Clear": CTcol.Add "Glass -Green": CTcol.Add "HDPE": CTcol.Add "Liquid Paper Board"
'    CTcol.Add "Other": CTcol.Add "Other Plastics": CTcol.Add "PET": CTcol.Add "Steel"
'    Set wkb = ActiveWorkbook
'    Set TopDic = New Scripting.Dictionary
'    With ActiveSheet
'        For Each Piv In .PivotTables
'            Set Dic = New Scripting.Dictionary
'            For it = 1 To Piv.RowRange.Count - 2
'                Dic.Add Piv.RowFields(1).PivotItems(it), Piv.GetPivotData("QTY Rcvd", "ContainerType", Piv.RowFields(1).PivotItems(it))
'            Next
'            TopDic.Add Mid(Piv.Name, 1, InStr(1, Piv.Name, "Pivot") - 1), Dic
'        Next
'    End With
'
'    For Each wks In wkb.Worksheets
'        'wks.Activate
'        headrow = 0
'        Debug.Print wks.CodeName
'        For Each cell In wks.Columns(1).Cells
'            If cell.Value = "Region" Then headrow = cell.Row
'            If headrow = 0 And cell.Row > 20 Then
'                If wks.CodeName <> "wks_Sum" Then MsgBox "No Header Row Identified (" & wks.Name & " Sheet)"
'                Exit For
'            ElseIf headrow <> 0 Then
'                Select Case wks.CodeName
'                    Case "wks_NSW"
'                        wks.Activate
'                        'wks.Cells(headrow, 1).CurrentRegion.Select
'                        Data(1) = wks.Cells(headrow, 1).CurrentRegion.Value2
'                        Exit For
'                    Case "wks_ACT"
'                        Data(2) = wks.Cells(headrow, 1).CurrentRegion.Value2
'                        Exit For
'                    Case "wks_QLD"
'                        Data(3) = wks.Cells(headrow, 1).CurrentRegion.Value2
'                        Exit For
'                    Case "wks_EXP_QLD"
'                        Data(4) = wks.Cells(headrow, 1).CurrentRegion.Value2
'                        Exit For
'                    Case "wks_EXP_ACT"
'                        Data(5) = wks.Cells(headrow, 1).CurrentRegion.Value2
'                        Exit For
'                End Select
'            End If
'        Next
'    Next
'
'    Set CN = New ADODB.Connection
'    With CN
'        CN.CommandTimeout = 50
'        CN.ConnectionTimeout = 50
'        CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & CBA_BSA & "LIVE DATABASES\CDS_ADM.accdb"
'    End With
'
'    For d = 1 To 5
'        For a = 2 To UBound(Data(d), 1)
'            If Data(d)(a, 15) > 0 Then
'                strSQL = "INSERT INTO Margin_Data (Region,Contractno,Productcode,Branded,SchemeID,Supplierno, Containertype,IncoFrom,IncoTo,IncoTerm,Packsize,QtyRCVD,UnitRCVD)" & Chr(10)
'                strSQL = strSQL & "VALUES (" & Data(d)(a, 1) & "," & Data(d)(a, 2) & "," & Data(d)(a, 3) & "," & IIf(Data(d)(a, 5) = "Yes", True, False) & ",'" & Data(d)(a, 6) & "'," & Data(d)(a, 7) & ",'" & Data(d)(a, 9) & "',#" & Format(Data(d)(a, 10), "MM/DD/YYYY") & "#,#" & Format(Data(d)(a, 11), "MM/DD/YYYY") & "#,'" & Data(d)(a, 12) & "'," & Data(d)(a, 13) & "," & Data(d)(a, 14) & "," & Data(d)(a, 15) & ")"
'                Set RS = New ADODB.Recordset
'                RS.Open strSQL, CN
'                Set RS = Nothing
'            End If
'        Next
'    Next
'
''    If TopDic.Exists("NSW") Then
''        For Each ct In CTcol
''            If TopDic("NSW").Exists(ct) Then
''                strSQL = "INSERT INTO CT_Date (DateFrom,DateTo,SchemeID, Containertype,QTY)"
''                strSQL = strSQL & "VALUES (#" &
''            End If
''        Next
''    End If
''    If TopDic.Exists("ACT") Then
''
''    End If
''    If TopDic.Exists("QLD") Then
''
''    End If
''    If TopDic.Exists("EXT_ACT") Then
''
''    End If
''    If TopDic.Exists("EXT_QLD") Then
''
''    End If
End Sub
Sub GeneratePreCDSCheckReport(Control As IRibbonControl)
Dim wbk As Workbook
Dim sht As Worksheet
Dim rw As Long, b As Long, a As Long, c As Long
Dim bfound As Boolean, include As Boolean
Dim but As Object
    
    Application.ScreenUpdating = False
    CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_CBISQuery"
    CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_Cont_Map"
    Set wbk = Workbooks.Add
    Set sht = ActiveSheet
    With sht
        Range(.Cells(1, 1), .Cells(5, 79)).Interior.ColorIndex = 49
        .Cells(1, 1).Select
        .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
        .Cells(4, 3).Value = "CDS 'Not yet eligible flagged' Report"
        .Cells.Font.Name = "ALDI SUED Office"
        .Cells(4, 3).Font.Size = 24
        .Cells(4, 3).Font.ColorIndex = 2
        rw = 6
        .Cells(rw, 1).Value = "ContractNo"
        .Cells(rw, 2).Value = "ProductCode"
        .Cells(rw, 3).Value = "Product Description"
        .Cells(rw, 4).Value = "CG No."
        .Cells(rw, 5).Value = "SCG No."
        .Cells(rw, 6).Value = "ProductClass"
        .Cells(rw, 7).Value = "Region"
        .Cells(rw, 8).Value = "Supplier No."
        .Cells(rw, 9).Value = "Supplier Name"
        .Cells(rw, 10).Value = "SchemeID"
        .Cells(rw, 11).Value = "Container Type"
        .Cells(rw, 12).Value = "Inco Terms"
        .Cells(rw, 13).Value = "Pickup Street"
        .Cells(rw, 14).Value = "Pickup City"
        .Cells(rw, 15).Value = "Pickup State"
        .Cells(rw, 16).Value = "Pickup Country"
        .Cells(rw, 17).Value = "Branded"
        .Cells(rw, 18).Value = "Eligible?"
        .Cells(rw, 19).Value = "ALDI Responsible?"
        For b = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
            bfound = False
            include = True
            If CBA_CDSarr(0, 0) <> 0 Then
                For a = LBound(CBA_CDSarr, 2) To UBound(CBA_CDSarr, 2)
                    If CLng(CBA_CBISarr(0, b)) = CLng(CBA_CDSarr(1, a)) Then
                        If CLng(CBA_CBISarr(6, b)) = CLng(CBA_CDSarr(2, a)) And _
                            CStr(CBA_CBISarr(9, b)) = CStr(CBA_CDSarr(3, a)) And _
                                CStr(CBA_CBISarr(10, b)) = CStr(CBA_CDSarr(4, a)) And _
                                    CStr(NZ(CBA_CBISarr(11, b), "")) = CStr(NZ(CBA_CDSarr(5, a), "")) And _
                                        CBA_CDSarr(8, a) = True Then
                            include = False
                            Exit For
                        End If
                    End If
                Next a
            End If
            If bfound = False And include = True Then
                rw = rw + 1
                For c = LBound(CBA_CBISarr, 1) To UBound(CBA_CBISarr, 1)
                    .Cells(rw, c + 1).Value = CBA_CBISarr(c, b)
                Next
                .Cells(rw, 18).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes,No"
            End If
        Next b
    End With
    
        Set but = sht.Shapes.AddShape(msoShapeRoundedRectangle, 900, 20, 120, 50)
        With but.OLEFormat.Object
            .Caption = "Complete Check"
            .Font.Size = 16
            .Font.Name = "ALDI SUED Office"
            .HorizontalAlignment = xlCenter
            .OnAction = "CBA_CDS_Runtime.ProceedToCDSReport"
        End With
    
    Application.ScreenUpdating = True
End Sub
Sub enterCostData(Control As IRibbonControl)
Dim sht As Worksheet
Dim rw As Long
Dim but As Object

    Workbooks.Add
    Set sht = ActiveSheet
    With sht
        Range(.Cells(1, 1), .Cells(5, 79)).Interior.ColorIndex = 49
        .Cells(1, 1).Select
        .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
        .Cells(4, 3).Value = "CDS Cost Entry Form"
        .Cells.Font.Name = "ALDI SUED Office"
        .Cells(4, 3).Font.Size = 24
        .Cells(4, 3).Font.ColorIndex = 2
        rw = 6
        .Cells(rw, 1).Value = "DateFrom"
        .Cells(rw, 2).Value = "DateTo"
        .Cells(rw, 3).Value = "SchemeID"
        .Cells(rw, 4).Value = "ContainerType"
        .Cells(rw, 5).Value = "Cost"
        .Cells(rw, 6).Value = "Forecasted Containers"
        Range(.Cells(rw, 1), .Cells(rw, 6)).Font.Bold = True
        Range(.Cells(rw, 1), .Cells(rw, 6)).EntireColumn.ColumnWidth = 25.5
        Set but = .Shapes.AddShape(msoShapeRoundedRectangle, 900, 20, 120, 50)
        With but.OLEFormat.Object
            .Caption = "Submit Costs"
            .Font.Size = 16
            .Font.Name = "ALDI SUED Office"
            .HorizontalAlignment = xlCenter
            .OnAction = "CBA_CDS_Runtime.SubmitCostsToDatabase"
        End With
    End With

End Sub
Sub SubmitCostsToDatabase()
Dim CN As ADODB.Connection
Dim RS As ADODB.Recordset
Dim DT() As AdjCostType
Dim noOf As Long, cnt As Long
Dim sht As Worksheet
Dim RCell As Range
Dim curDFrom As Date, curDto As Date
Dim curSID, thisval
Dim yn As Long
Dim strSQL As String
Dim a


Set CN = New ADODB.Connection



With CN
    CN.CommandTimeout = 50
    CN.ConnectionTimeout = 50
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & CBA_BSA & "LIVE DATABASES\CDS_ADM.accdb"
End With



    Set sht = ActiveSheet
    With sht
        cnt = 0: noOf = 0
        For Each RCell In .Columns(1).Cells
            If RCell.Row > 5 And RCell.Value = "" And RCell.Offset(1, 0).Value = "" And RCell.Offset(2, 0).Value = "" Then Exit For
            If RCell.Row > 6 And RCell.Value <> "" Then
                If cnt = 0 Then noOf = 1: curDFrom = RCell.Value: curDto = RCell.Offset(0, 1).Value: curSID = RCell.Offset(0, 2).Value
                cnt = cnt + 1
                If curDFrom <> RCell.Value Or curDto <> RCell.Offset(0, 1).Value Or curSID <> RCell.Offset(0, 2).Value Then
                    curDFrom = RCell.Value: curDto = RCell.Offset(0, 1).Value:  curSID = RCell.Offset(0, 2).Value: noOf = noOf + 1
                End If
            End If
        Next
        ReDim DT(1 To noOf)
        cnt = 0
        For Each RCell In .Columns(1).Cells
            If RCell.Row > 5 And RCell.Value = "" And RCell.Offset(1, 0).Value = "" And RCell.Offset(2, 0).Value = "" Then
                If noOf > 0 Then
enterRebate2:
                    thisval = InputBox("Please enter the total rebate (excl. GST) for:" & Chr(10) & "DateFrom:" & DT(noOf).DateFrom & Chr(10) & "DateTo:" & DT(noOf).DateTo & Chr(10) & "SchemeID:" & DT(noOf).SchemeID, "Rebate")
                    If thisval <> "" Then
                        DT(noOf).Rebate = thisval
                        yn = MsgBox("You entered a total rebate of: " & DT(noOf).Rebate & Chr(10) & Chr(10) & "Is this correct?", vbYesNo)
                        If yn = 7 Then GoTo enterRebate2
                    Else
                        CN.Close
                        Set CN = Nothing
                        Exit Sub
                    End If
                End If
                Exit For
            End If
            If RCell.Row > 6 And RCell.Value <> "" Then
                If cnt = 0 Then noOf = 1: DT(noOf).DateFrom = RCell.Value: DT(noOf).DateTo = RCell.Offset(0, 1).Value: DT(noOf).SchemeID = RCell.Offset(0, 2).Value
                cnt = cnt + 1
                If DT(noOf).DateFrom <> RCell.Value Or DT(noOf).DateTo <> RCell.Offset(0, 1).Value Or DT(noOf).SchemeID <> RCell.Offset(0, 2).Value Then
                    DT(noOf).DateFrom = RCell.Value: DT(noOf).DateTo = RCell.Offset(0, 1).Value:  DT(noOf).SchemeID = RCell.Offset(0, 2).Value: noOf = noOf + 1
enterRebate1:
                    thisval = InputBox("Please enter the total rebate (excl. GST) for:" & Chr(10) & "DateFrom:" & DT(noOf).DateFrom & Chr(10) & "DateTo:" & DT(noOf).DateTo & Chr(10) & "SchemeID:" & DT(noOf).SchemeID, "Rebate")
                    If thisval <> "" Then
                        DT(noOf).Rebate = thisval
                        yn = MsgBox("You entered a total rebate of: " & DT(noOf).Rebate & Chr(10) & Chr(10) & "Is this correct?", vbYesNo)
                        If yn = 7 Then GoTo enterRebate1
                    Else
                        CN.Close
                        Set CN = Nothing
                        Exit Sub
                    End If
                    
                End If
                DT(noOf).TotalCost = DT(noOf).TotalCost + (RCell.Offset(0, 4).Value * RCell.Offset(0, 5).Value)
            End If
        Next
        For Each RCell In .Columns(1).Cells
            If RCell.Row > 5 And RCell.Value = "" And RCell.Offset(1, 0).Value = "" And RCell.Offset(2, 0).Value = "" Then Exit For
            If RCell.Row > 6 And RCell.Value <> "" Then
                Set RS = New ADODB.Recordset
                strSQL = "SELECT ID FROM Scheme_Cost" & Chr(10)
                strSQL = strSQL & "where DateFrom = #" & Format(RCell.Value, "MM/DD/YYYY") & "#" & Chr(10)
                strSQL = strSQL & "and DateTo = #" & Format(RCell.Offset(0, 1).Value, "MM/DD/YYYY") & "#" & Chr(10)
                strSQL = strSQL & "and SchemeID = '" & RCell.Offset(0, 2).Value & "'" & Chr(10)
                strSQL = strSQL & "and ContainerType = '" & RCell.Offset(0, 3).Value & "'" & Chr(10)
                RS.Open strSQL, CN
                If RS.EOF Then
                    strSQL = "INSERT INTO Scheme_Cost (DateFrom,DateTo,SchemeID,ContainerType,Cost,ForecastedContainers,AdjustedCostRate) " & Chr(10)
                    strSQL = strSQL & "VALUES (#" & Format(RCell.Value, "MM/DD/YYYY") & "#"
                    strSQL = strSQL & ",#" & Format(RCell.Offset(0, 1).Value, "MM/DD/YYYY") & "#"
                    strSQL = strSQL & ",'" & RCell.Offset(0, 2).Value & "'"
                    strSQL = strSQL & ",'" & RCell.Offset(0, 3).Value & "'"
                    strSQL = strSQL & "," & RCell.Offset(0, 4).Value
                    strSQL = strSQL & "," & RCell.Offset(0, 5).Value
                    For a = 1 To noOf
                        If DT(noOf).DateFrom = RCell.Value And DT(noOf).DateTo = RCell.Offset(0, 1).Value And DT(noOf).SchemeID = RCell.Offset(0, 2).Value Then
                            thisval = ((RCell.Offset(0, 4).Value * RCell.Offset(0, 5).Value) / DT(noOf).TotalCost) * ((DT(noOf).TotalCost - DT(noOf).Rebate) / RCell.Offset(0, 5).Value)
                            strSQL = strSQL & "," & thisval & ")"
                            Exit For
                        End If
                    Next
                    Set RS = New ADODB.Recordset
                    RS.Open strSQL, CN
                Else
                    strSQL = "UPDATE Scheme_Cost" & Chr(10)
                    strSQL = strSQL & "Set Cost = " & RCell.Offset(0, 4).Value & Chr(10)
                    strSQL = strSQL & ", ForecastedContainers = " & RCell.Offset(0, 5).Value & Chr(10)
                    For a = 1 To noOf
                        If DT(noOf).DateFrom = RCell.Value And DT(noOf).DateTo = RCell.Offset(0, 1).Value And DT(noOf).SchemeID = RCell.Offset(0, 2).Value Then
                            thisval = ((RCell.Offset(0, 4).Value * RCell.Offset(0, 5).Value) / DT(noOf).TotalCost) * ((DT(noOf).TotalCost - DT(noOf).Rebate) / RCell.Offset(0, 5).Value)
                            strSQL = strSQL & ", AdjustedCostRate = " & thisval & Chr(10)
                            Exit For
                        End If
                    Next
                    strSQL = strSQL & "where ID = " & RS(0) & Chr(10)
                    Set RS = New ADODB.Recordset
                    RS.Open strSQL, CN
                End If
            End If
        Next
    End With
'    Application.DisplayAlerts = False
'    ActiveWorkbook.Close
'    Application.DisplayAlerts = True
    
CN.Close
Set CN = Nothing

End Sub
Sub CDS_PreAuditReport(Control As IRibbonControl)
Dim wbk As Workbook
Dim but As Object
Dim rw As Long, b As Long, d As Long, T As Long, m As Long
Dim isSame As Boolean

    Application.ScreenUpdating = False
    
    CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_AuditbyCG"
    CBA_SQL_Queries.CBA_GenPullSQL "CBA_CDS_AuditCheckData"
    If CBA_CBISarr(0, 0) = 0 Then
    Else
        CBA_CDS_PreAudit.Copy
        Set wbk = ActiveWorkbook
        With ActiveSheet
            Range(.Cells(1, 1), .Cells(5, 79)).Interior.ColorIndex = 49
            .Cells(1, 1).Select
            .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
            .Cells(4, 3).Value = "CDS Pre-Audit Report"
            .Cells.Font.Name = "ALDI SUED Office"
            .Cells(4, 3).Font.Size = 24
            .Cells(4, 3).Font.ColorIndex = 2
            Set but = .Shapes.AddShape(msoShapeRoundedRectangle, 900, 20, 120, 50)
            With but.OLEFormat.Object
                .Caption = "Send to DataBase"
                .Font.Size = 16
                .Font.Name = "ALDI SUED Office"
                .HorizontalAlignment = xlCenter
                .OnAction = "CBA_CDS_PreAudit.InterfaceToDB"
            End With
            rw = 6
            .Cells(rw, 1).Value = "ContractNo"
            .Cells(rw, 2).Value = "ProductCode"
            .Cells(rw, 3).Value = "Product Description"
            .Cells(rw, 4).Value = "CG No."
            .Cells(rw, 5).Value = "CG Description"
            .Cells(rw, 6).Value = "SCG No."
            .Cells(rw, 7).Value = "SCG Description"
            .Cells(rw, 8).Value = "Supplier No."
            .Cells(rw, 9).Value = "Supplier Name"
            .Cells(rw, 10).Value = "ProductClass"
            .Cells(rw, 11).Value = "On Sale Date"
            .Cells(rw, 12).Value = "DeliveryFrom"
            .Cells(rw, 13).Value = "DeliveryTo"
            .Cells(rw, 14).Value = "SA"
            .Cells(rw, 15).Value = "NSW"
            .Cells(rw, 16).Value = "QLD"
            .Cells(rw, 17).Value = "ACT"
            .Cells(rw, 18).Value = "CDS Applied?"
            Range(.Cells(6, 1), .Cells(6, 18)).Font.Bold = True
            For b = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                For d = LBound(CBA_CDSarr, 2) To UBound(CBA_CDSarr, 2)
                    If CLng(CBA_CBISarr(0, b)) < CLng(CBA_CDSarr(1, d)) Then Exit For
                    If CLng(CBA_CBISarr(0, b)) = CLng(CBA_CDSarr(1, d)) Then
                        isSame = True
                        For T = LBound(CBA_CBISarr, 1) To UBound(CBA_CBISarr, 1)
                            Select Case T
                                Case 0: m = 1
                                Case 1: m = 2
                                Case 3: m = 3
                                Case 5: m = 4
                                Case 7: m = 5
                                Case 9: m = 6
                                Case 10: m = 7
                                Case 11: m = 8
                                Case 12: m = 9
                                Case 13: m = 10
                                Case 14: m = 11
                                Case 15: m = 12
                                Case 16: m = 13
                                Case 17: m = 14
                                Case Else: m = 0
                            End Select
                            If m > 0 Then
                                If m < 7 Then
                                    If CLng(CBA_CBISarr(T, b)) <> CLng(CBA_CDSarr(m, d)) Then
                                        isSame = False
                                        Exit For
                                    End If
                                ElseIf m < 10 Then
                                    If CStr(NZ(CBA_CBISarr(T, b), "")) <> CStr(NZ(CBA_CDSarr(m, d), "")) Then
                                        isSame = False
                                        Exit For
                                    End If
                                Else
                                    If (CBA_CBISarr(T, b) = "No" And CBA_CDSarr(m, d) = True) Or (CBA_CBISarr(T, b) = "Yes" And CBA_CDSarr(m, d) = False) Then
                                        isSame = False
                                        Exit For
                                    End If
                                End If
                            End If
                        Next T
                        If isSame = True Then
                            If CBA_CDSarr(14, d) = False Then
                                rw = rw + 1
                                For T = LBound(CBA_CBISarr, 1) To UBound(CBA_CBISarr, 1)
                                    .Cells(rw, T + 1).Value = CBA_CBISarr(T, b)
                                Next
                                Exit For
                            End If
                        End If
                    End If
                Next d
            Next b
            Range(.Cells(6, 1), .Cells(6, 18)).EntireColumn.AutoFit
            Range(.Cells(7, 18), .Cells(rw, 18)).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes,No"
        End With
    End If
    Application.ScreenUpdating = True

End Sub
Sub ImportPreAuditDataToDatabase(ByRef Data As Variant)
Dim strSQL As String
Dim CN As ADODB.Connection
Dim RS As ADODB.Recordset
Dim b As Long


Set CN = New ADODB.Connection
With CN
    CN.CommandTimeout = 50
    CN.ConnectionTimeout = 50
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & CBA_BSA & "LIVE DATABASES\CDS_ADM.accdb"
End With

    For b = LBound(Data, 1) + 1 To UBound(Data, 1)
        If Data(b, 18) = "Yes" Or Data(b, 18) = "No" Then
            Set RS = New ADODB.Recordset
            strSQL = "Update CDS_CONT_CHECK " & Chr(10)
            If Data(b, 18) = "No" Then strSQL = strSQL & "Set CDS_No_Applies = True,  Checked = True" & Chr(10)
            If Data(b, 18) = "Yes" Then strSQL = strSQL & "Set CDS_No_Applies = False,  Checked = True" & Chr(10)
            strSQL = strSQL & "Where ContractNo = " & Data(b, 1) & Chr(10)
            strSQL = strSQL & "and productcode = " & Data(b, 2) & Chr(10)
            strSQL = strSQL & "and CGno     = " & Data(b, 4) & Chr(10)
            strSQL = strSQL & "and SCGno    = " & Data(b, 6) & Chr(10)
            strSQL = strSQL & "and SupplierNo = " & Data(b, 8) & Chr(10)
            strSQL = strSQL & "and ProductClass     = " & Data(b, 10) & Chr(10)
            If IsEmpty(Data(b, 11)) Or IsNull(Data(b, 11)) Then
                strSQL = strSQL & "and OSD is Null " & Chr(10)
            Else
                strSQL = strSQL & "and OSD = #" & Format(Data(b, 11), "MM/DD/YYYY") & "#" & Chr(10)
            End If
            If IsEmpty(Data(b, 12)) Or IsNull(Data(b, 12)) Then
                strSQL = strSQL & "and Deliveryfrom is Null " & Chr(10)
            Else
                strSQL = strSQL & "and Deliveryfrom = #" & Format(Data(b, 12), "MM/DD/YYYY") & "#" & Chr(10)
            End If
            If IsEmpty(Data(b, 13)) Or IsNull(Data(b, 13)) Then
                strSQL = strSQL & "and Deliveryto is Null " & Chr(10)
            Else
                strSQL = strSQL & "and Deliveryto = #" & Format(Data(b, 13), "MM/DD/YYYY") & "#" & Chr(10)
            End If
            If Data(b, 14) = "No" Then strSQL = strSQL & "and SA = False" & Chr(10) Else strSQL = strSQL & "and SA = True" & Chr(10)
            If Data(b, 15) = "No" Then strSQL = strSQL & "and NSW = False" & Chr(10) Else strSQL = strSQL & "and NSW = True" & Chr(10)
            If Data(b, 16) = "No" Then strSQL = strSQL & "and QLD = False" & Chr(10) Else strSQL = strSQL & "and QLD = True" & Chr(10)
            If Data(b, 17) = "No" Then strSQL = strSQL & "and ACT = False" & Chr(10) Else strSQL = strSQL & "and ACT = True" & Chr(10)
            'Debug.Print strSQL
            RS.Open strSQL, CN
'        ElseIf Data(b, 18) = "" Then
'            a = a
        End If
    Next b
CN.Close
Set CN = Nothing
Set RS = Nothing
End Sub
