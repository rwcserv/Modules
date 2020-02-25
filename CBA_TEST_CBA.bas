Attribute VB_Name = "CBA_TEST_CBA"
Option Explicit
Option Private Module       ' Excel users cannot access procedures
Sub TestProdObjectData()
Dim iMonth As Long, iYear As Long
Dim colCG As Collection
Dim PG As cCBA_ProdGroup
Dim P   As cCBA_Prod
Dim a As Long
Dim arr As Variant, barr As Variant
Dim pdteFrom As Date, pdteTo As Date
Dim COM As CBA_COM_COMMatch
Dim v As Variant
Dim col As Collection
Dim va As Variant

iMonth = 12
iYear = 2019
pdteFrom = #1/1/2019#
pdteTo = #12/31/2019#
Set colCG = New Collection
'colCG.Add "044001"
'colCG.Add "044002"
'colCG.Add "044003"
'colCG.Add "044005"
colCG.Add "004001"


Set PG = New cCBA_ProdGroup
CBA_BasicFunctions.CBA_Running "TEST"
'If PG.RunDataGeneration(DateSerial(iYear - 1, iMonth + 1, 1), DateSerial(iYear, iMonth + 1, 0), False, colCG, "BW") = True Then

'If PG.RunDataGeneration(DateSerial(iYear - 1, iMonth + 1, 1), DateSerial(iYear, iMonth + 1, 0), False, colCG, "BWPRTMsCm") = True Then
If PG.RunDataGeneration(DateSerial(iYear - 1, iMonth + 1, 1), DateSerial(iYear, iMonth + 1, 0), False, colCG, "BPR") = True Then
CBA_BasicFunctions.CBA_Close_Running



'Debug.Print PG.getWastageData(eRetail, eBoth, pdteFrom, pdteTo)
'Debug.Print Format(PG.getUSW(eUSWp, pdteFrom, pdteTo, eQty, 1, 501), "0.0")
'Debug.Print Format(PG.getUSW(eUSWp, pdteFrom, pdteTo, eQty, 1), "0.0")
'Debug.Print Format(PG.getPOS(pdteFrom, pdteTo, eRetail, 1, 505), "$#,#0.0")
'Debug.Print Format(PG.getRCV(pdteFrom, pdteTo, eRCVCost), "$#,#0.0")
'Debug.Print Format(PG.getRCVMargin(eRCVContributionDollar, pdteFrom, pdteTo), "$#,#0.0")
'Debug.Print Format(PG.getSellThrough(pdteFrom, pdteTo), "0.0%")







Set P = PG.getProdObject(40301, True)

P.CalculateTrialPeriod
Debug.Print Format(P.getPOSdata(pdteFrom, pdteTo, True), "#")
Debug.Print Format(P.getPOSdata(pdteFrom, pdteTo, False), "#.00")
Debug.Print Format(P.getPOSdata(pdteFrom, pdteTo, True, eProductLevel), "#")
Debug.Print Format(P.getPOSdata(pdteFrom, pdteTo, False, eProductLevel), "#.00")





Debug.Print Format(P.getRCVShare(pdteFrom, pdteTo, True), "#0.00%")
Debug.Print Format(P.getRCVShare(pdteFrom, pdteTo, False), "#0.00%")

Debug.Print Format(P.getRCVdata(pdteFrom, pdteTo, "Retail", 501), "$#,#0.00")
Debug.Print Format(P.getPOSdata(pdteFrom, pdteTo, False, eNotUSW, , , 501), "$#,#0.00")


iYear = iYear
'Debug.Print P.getPOSdata(pdteFrom, pdteTo, True, e_POSUSWTypes.eProductLevel, , , 501)
'arr = P.getIncoTermData("CVfVtS", , , , 501)
'arr = P.getContractData("CRDfSnTx", , , , 501)
'barr = P.getBraketData("VfDEFTFr", , , , 501)

'For a = LBound(arr, 2) To UBound(arr, 2)
'    Debug.Print arr(0, a), arr(1, a), arr(2, a), arr(3, a) ', arr(4, a) ', arr(5, 0) ', arr(6, 0) ', arr(7, 0)
'Next
'For a = LBound(barr, 2) To UBound(barr, 2)
'    Debug.Print barr(0, 0), barr(1, 0), barr(2, 0), barr(3, 0), barr(4, 0) ', barr(5, 0) ', barr(6, 0) ', barr(7, 0)
'Next

'Debug.Print P.getPOSdata(pdteFrom, pdteTo, True, , , , 505)

'Debug.Print PG.getPOS(pdteFrom, pdteTo, eQty, 505, 1)

'Debug.Print PG.getUSW(eUSWp, pdteFrom, pdteTo, eQty, , 503)
'Set col = PG.sdCOMdic("11525")
''    Set col = PG.sdCOMdic(v)
'    For Each va In col
'        Set COM = va
'        Debug.Print COM.CompCode, COM.CompProdName
'
'
'    Next

'Debug.Print PG.getWastageData(eRetail, eInventoryDifference, pdteFrom, pdteTo, , 1, 501)
'Debug.Print PG.getWastageData(eRetail, eMarkdowns, pdteFrom, pdteTo, , 1, 501)
'Debug.Print PG.getWastageData(eRetail, eBoth, pdteFrom, pdteTo, , 1, 501)
'Debug.Print PG.getWastageData(eQty, eInventoryDifference, pdteFrom, pdteTo, , 1, 501)
'Debug.Print PG.getWastageData(eQty, eMarkdowns, pdteFrom, pdteTo, , 1, 501)
'Debug.Print PG.getWastageData(eQty, eBoth, pdteFrom, pdteTo, , 1, 501)



End If

End Sub
Sub testbuildingSalesdata()
Dim PG As cCBA_ProdGroup
If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
Set PG = New cCBA_ProdGroup
PG.buildSalesDataRS 2017, "", True
If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
End Sub
Sub testforecast()
Dim v1 As Variant, v2 As Variant, v3 As Variant

    CBA_BTF_SetupForecastArray.CBA_BTF_SetupForecastArray 2020, 2020, 1, 12
    
    Debug.Print FCbM(2020, 1, 1).CGData(1, 0).Sales().ReForecast
    Debug.Print FCbM(2020, 1, 1).CGData(1, 0).Sales().ReForecast

End Sub



Sub testCOMRADEMatch()
Dim COM As CBA_COM_COMMatch
Dim v As Variant
Dim col As Collection
Dim pdteFrom As Date, pdteTo As Date
Dim a As Long

pdteFrom = #2/1/2019#
pdteTo = #1/31/2020#

    If CBA_COM_SetupMatchArray.CBA_SetupMatchArray(False, pdteFrom, pdteTo, , , 11525, True) = True Then
        For a = LBound(CBA_COM_Match) To UBound(CBA_COM_Match)
            Set COM = CBA_COM_Match(a)
        Next
    End If



End Sub

Sub testnielsen()
Dim dic As Scripting.Dictionary
Dim v As Variant
Dim ND As cCBA_NielsenData

Set dic = New Scripting.Dictionary
Set dic = mCBA_Nielsen.GetNielsenSegmentationData '(2019, 12, False)

Debug.Print dic.Keys(5)

For Each v In dic
    Set ND = dic(v)
    If ND.isScanData = True Then
        Debug.Print ND.MSegDescription
    End If
    
Next




End Sub

Sub testmsegsel()
Dim mss As fCAM_MSegSelect

Set mss = New fCAM_MSegSelect
''If mss.bIsActive = True Then mss.Show vbModeless Else MsgBox "No Active Data Object"

End Sub


Sub TestLCForm()

Dim LCF As fCAM_LineCount

Set LCF = New fCAM_LineCount
If LCF.bIsActive Then LCF.Show

End Sub

Sub TestDocStuff()

'mCAM_Runtime.cDocHolder.ControlObject(

End Sub


''
''Sub TestLCFormWthListView()
''
''Dim LCF As fCAM_LineCountListView
''
''Set LCF = New fCAM_LineCountListView
''If LCF.bIsActive Then LCF.Show
''
''End Sub


'Sub TestGetAllMsegNames()
'Dim dic As Scripting.Dictionary
'Dim arr As Variant
'arr = mCAM_Runtime.cRibbonData.GetAllMSegNames(eScanData)
'
'End Sub



Sub TESTCoreRangeCategoryReview()
    Dim cDocObject As cCAM_Category, DocID As Long
    Set cDocObject = New cCAM_Category
    CBA_BasicFunctions.CBA_Running
    DocID = -1
    If cDocObject.Construct("Tea", #6/1/2019#, #12/31/2019#, False, e_DocuType.eCoreRangeCategoryReview) = True Then
        If mCAM_Runtime.CreateDocument(cDocObject.eDocumentType, "TEST", DocID, cDocObject) = True Then
            If mCAM_Runtime.clsDocHolder.Render_Document(DocID) = False Then
                MsgBox "Camera didn't Render Document correctly"
            End If
        ''MsgBox "Success"
        Else
            MsgBox "Oh No"
        End If
    End If
    CBA_BasicFunctions.CBA_Close_Running

''    Dim cDocH As cCBA_DocumentHolder
''    Set cDocH = New cCBA_DocumentHolder
''    Call g_GetDB("UDT", , , , , "Cam")
''    Call cDocH.Add_Document(eCoreRangeCategoryReview, "TESTCATREV")
End Sub


Sub TESTSpecialsCategoryReview()
    Dim cDocObject As cCAM_Category, DocID As Long
    Set cDocObject = New cCAM_Category
    CBA_BasicFunctions.CBA_Running
    DocID = -1
    If cDocObject.Construct("Tea", #6/1/2019#, #12/31/2019#, False, e_DocuType.eSpecialsCategoryReview) = True Then
        If mCAM_Runtime.CreateDocument(cDocObject.eDocumentType, "TEST", DocID, cDocObject) = True Then
            If mCAM_Runtime.clsDocHolder.Render_Document(DocID) = False Then
                MsgBox "Camera didn't Render Document correctly"
            End If
        ''MsgBox "Success"
        Else
            MsgBox "Oh No"
        End If
    End If
    CBA_BasicFunctions.CBA_Close_Running

''    Dim cDocH As cCBA_DocumentHolder
''    Set cDocH = New cCBA_DocumentHolder
''    Call g_GetDB("UDT", , , , , "Cam")
''    Call cDocH.Add_Document(eCoreRangeCategoryReview, "TESTCATREV")
End Sub

Sub TESTSpecSeasPerformance(Optional ByRef DocID As Long = -1)
    Dim cDocObject As cCAM_Category
    Set cDocObject = New cCAM_Category
    CBA_BasicFunctions.CBA_Running
    If cDocObject.Construct("Tea", #2/1/2019#, #1/31/2020#, False, e_DocuType.eSpecSeasPerformance) = True Then
        If mCAM_Runtime.CreateDocument(cDocObject.eDocumentType, "TEST", DocID, cDocObject) = True Then
            If mCAM_Runtime.clsDocHolder.Render_Document(DocID) = False Then
                MsgBox "Camera didn't Render (DocID=" & DocID & ") Document correctly"
                Stop
            End If
            If mCAM_Runtime.clsDocHolder.ControlObject(DocID, , "Save") = False Then
                MsgBox "Camera didn't Save (DocID=" & DocID & ") Document correctly"
                Stop
            End If
        Else
            MsgBox "Oh No"
        End If
    End If
    CBA_BasicFunctions.CBA_Close_Running

''    Dim cDocH As cCBA_DocumentHolder, DocName As String
''    Set cDocH = New cCBA_DocumentHolder
''    DocName = "TESTCATREV8"
''    Call g_GetDB("UDT", , , , , "Cam")
''    Call cDocH.Add_Document(eSpecSeasPerformance, DocName)
''    Call cDocH.Render_Document(1)
    
End Sub

Sub TESTLineCountOverviewReport(Optional ByRef DocID As Long = -1)
    Dim cDocObject As cCAM_Category
    Set cDocObject = New cCAM_Category
    CBA_BasicFunctions.CBA_Running
    If cDocObject.Construct("Tea", #2/1/2019#, #1/31/2020#, False, e_DocuType.eLineCountOverviewReport) = True Then
        If mCAM_Runtime.CreateDocument(cDocObject.eDocumentType, "TEST", DocID, cDocObject) = True Then
            If mCAM_Runtime.clsDocHolder.Render_Document(DocID) = False Then
                MsgBox "Camera didn't Render (DocID=" & DocID & ") Document correctly"
                Stop
            End If
''            If mCAM_Runtime.clsDocHolder.ControlObject(DocID, , "Save") = False Then
''                MsgBox "Camera didn't Save (DocID=" & DocID & ") Document correctly"
''                Stop
''            End If
        Else
            MsgBox "Oh No"
        End If
    End If
    CBA_BasicFunctions.CBA_Close_Running

''    Dim cDocH As cCBA_DocumentHolder, DocName As String
''    Set cDocH = New cCBA_DocumentHolder
''    DocName = "TESTCATREV8"
''    Call g_GetDB("UDT", , , , , "Cam")
''    Call cDocH.Add_Document(eSpecSeasPerformance, DocName)
''    Call cDocH.Render_Document(1)
    
End Sub
''Sub TESTCoreRangeTEN()
''    Dim cDocH As cCBA_DocumentHolder, DocName As String
''    Set cDocH = New cCBA_DocumentHolder
''    DocName = "TESTCATREV8"
''    Call g_GetDB("UDT", , , , , "Ten")
''    Call cDocH.Add_Document(eCoreRangeTEN, DocName)
''    Call cDocH.Render_Document(1)
''
''End Sub
Sub TESTRenderingTheDoc(Optional DocName As String = "TESTCATREV3", Optional DocID As Long = 0)
''    Dim cDocH As cCBA_DocumentHolder
''    Set cDocH = New cCBA_DocumentHolder
''    Call g_GetDB("UDT", , , , , "Cam")
''    Call cDocH.Add_Document(eCoreRangeCategoryReview, DocName, DocID)
''    Call cDocH.Render_Document(DocName, DocID)
''    Call cDocH.Save_Document(DocName, DocID)
End Sub

Sub TESTFCasting(lCG As Long)
    Dim cFCast As cCBA_Forecasting, lRowCount As Long, lIdx As Long
    Set cFCast = New cCBA_Forecasting
    Call cFCast.GenForecasting(2020, 2020, "", "", lCG, 0, 0)
    
    lRowCount = cFCast.GetForecastValue(lIdx, "Count")                                     ' Record count
    For lIdx = 1 To lRowCount
           Debug.Print cFCast.GetForecastValue(lIdx, "Desc");                              ' Return this one into Sales Forecast
           Debug.Print "," & CStr(Round(cFCast.GetForecastValue(lIdx, "SalesRF"), 0));     ' Return this one into Sales Forecast
           Debug.Print "," & CStr(Round(cFCast.GetForecastValue(lIdx, "MarginRF"), 0));    ' Return this one into Margin Forecast
           Debug.Print "," & Round(cFCast.GetForecastValue(lIdx, "SalesYoY"), 3);          ' Return this one into Sales YoY %
           Debug.Print "," & Round(cFCast.GetForecastValue(lIdx, "MarginYoY"), 3)          ' Return this one into Margin YoY %
    Next
    Debug.Print "Grand Total for " & lCG;                                               ' Total Desc
    Debug.Print "," & Round(cFCast.GetForecastValue(lRowCount, "TSalesRF"), 0);            ' Return this one into Sales Total Forecast
    Debug.Print "," & Round(cFCast.GetForecastValue(lRowCount, "TMarginRF"), 0);           ' Return this one into Margin Total Forecast
    Debug.Print "," & Round(cFCast.GetForecastValue(lRowCount, "TSalesYoY"), 3);           ' Return this one into Sales Total YoY %
    Debug.Print "," & Round(cFCast.GetForecastValue(lRowCount, "TMarginYoY"), 3)           ' Return this one into Margin Total YoY %

End Sub

'''Callback for CBA_WhatRibbon onAction
''Sub WhatRibbon(control As IRibbonControl)
''
''    Debug.Print control.id
''    'Debug.Print ObjPtr(CBA_lRibbonPointer)
''    Debug.Print ObjPtr(CBA_Rib)
''
''End Sub
''
''Sub identifyribb()
''    Application.EnableEvents = True
''
''    'Debug.Print ObjPtr(IRibbonUI)
''
''
''
''
''End Sub
''Sub testimplement()
''Dim ob(0) As iCBA_Doc_Data
''Dim co As CBA_CAMERA_Data
''Set t = New CBA_CAMERA_Data
''Set ob(0) = t
''    If ob(0).DataForDocumentType = CoreRangeCategoryReview Then
''        Set co = ob(0)
''        Debug.Print co.getThis(True)
''    End If
''End Sub
''Sub testgenoneCBA_ProdObject()
''Dim PrdGrp As cCBA_ProdGroup
''Dim Prd As cCBA_Prod
''        Set PrdGrp = New cCBA_ProdGroup
''        PrdGrp.RunDataGeneration Date, Date, False, Nothing, "B", 1035
''        Set Prd = PrdGrp.getProdObject(Pcode)
''        Set PrdGrp = Nothing
''        Debug.Print Prd.PDesc
''End Sub
''Sub testCAMERAObject()
''Dim CAMD As CBA_CAMERA_Data
''    Set CAMD = New CBA_CAMERA_Data
''    Debug.Print CAMD.cPGrp.buildcode
''
''
''End Sub
''Sub testActiveXObjectCreationandControl()
''Dim objCmdBtn As OLEObject
''
''Set rnge = Application.Workbooks("Book4").Worksheets("Sheet1").Cells(1, 1)
''
''Set objCmdBtn = Application.Workbooks("Book4").Worksheets("Sheet1").OLEObjects.add(ClassType:="Forms.CommandButton.1", _
''                                     Link:=False, _
''                                     DisplayAsIcon:=False, _
''                                     Left:=rnge.Left, _
''                                     Top:=rnge.Top, _
''                                     Width:=rnge.Width, _
''                                     Height:=rnge.Height) '_
''                                     'Name:= "cmd_ButtonAdded")
''
''                                     With objCmdBtn
''                                        'set a String value as object's name
''                                        .Name = "cmd_AddedBut"
''                                        With .Object
''                                             .Caption = "Press Me please!"
''                                             With .Font
''                                                  .Name = "Arial"
''                                                  .Bold = True
''                                                  .Size = 7
''                                                  .Italic = False
''                                                  .Underline = False
''                                             End With
''                                        End With
''                                    End With
''
''
''
''
''End Sub
Sub testuserformstuff()
    ' NOTE: HAVE TO HAVE AN ACTIVEDATAOBJECT TO RUN... I.E. HAVE TO HAVE SET 'CATEGORY SELECTION' ETC
    Dim UF As fCam_FrontForm
    Call g_GetDB("Cam", , , True)
    If LoadFrontForm = False Then
        MsgBox "Error found in calling Front Form"
    End If
''    Set UF = New fCam_FrontForm
''    If UF.IsActive Then UF.Show

End Sub
''
''
''
''
''Sub testwhatgoesinalistbox()
''Dim UF As fCam_ManageCategory
''    Set UF = New fCam_ManageCategory
''    UF.Show
''
''
''End Sub
''Function generaltestfunction(ByVal sd As Collection) As Boolean
''    If sd(2) = "Rocks" Then generaltestfunction = True
''
''End Function
''
''Sub testcreatedocholder(ByVal DocName As String)
''
''    'Set CBA_DocHolder = New cCBA_DocumentHolder
''
''    CBA_DocHolder.ControlObject DocName
''
''End Sub
''Sub thisismetestingthis()
''Dim sd As Collection
''    Set sd = New Collection
''    sd.add "This"
''    sd.add "Rocks"
''    Debug.Print generaltestfunction(sd)
''End Sub
''
''Function TestThis(ByVal DocName As String, ByRef target As Range) As Variant
''
''    TestThis = "Yo MAMA"
''
''End Function
''
''
''Sub TESTRECOMMENDUFORM()
''    Dim Rcmd As CBA_TEN_Recommend
''    Dim col As Collection
''
''    Set Rcmd = New CBA_TEN_Recommend
''    Rcmd.TotalBusiness = 9000000000#
''    Rcmd.StoreCount = 530
''    Rcmd.setEstDelCost 1.89, 500, DDP
''    Rcmd.setEstDelCost 1.93, 501, ExWorks_FG
''    Rcmd.setEstDelCost 1.95, 502, ExWorks_FG
''    Rcmd.setEstDelCost 1.96, 503, ExWorks_FG
''    Rcmd.setEstDelCost 1.97, 504, ExWorks_FG
''    Set col = New Collection
''    col.add "Good Supplier"
''    col.add "Bad Supplier"
''    col.add "Not Bad Supplier"
''    Set Rcmd.SupplierNameCollection = col
''
''    Rcmd.Show vbModeless
''
''End Sub
''Sub TESTRETURNFROMRECOMMEND(ByRef Rcmd As CBA_TEN_Recommend)
''
''    MsgBox "I have the value of " & Rcmd.EstCaseDelCost & " as the Estimated Case Cost"
''
''End Sub
''''Option Private Module       ' Excel users cannot access procedures
''''
''''
''''Sub getCBISDATATEST()
''''Dim inArray() As Variant
''''
''''Set CBA_COM_CBISCN = New ADODB.Connection
''''Set CBA_COM_CBISRS = New ADODB.Recordset
''''
''''CBA_COM_CBISCN.Open
''''
''''strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10) & "SET DATEFIRST 1" & Chr(10) & "SET DATEFORMAT dmy" & Chr(10)
''''strSQL = strSQL & "select distinct p.productcode, p.cgno, p.scgno, p.productclass, p.description, p.packsize, p.unitcode_wght, p.unitcode_Vol, p.unitcode_Pack," & Chr(10)
''''strSQL = strSQL & "p.israndomweight, p.iscoldstorage, p.isfrozen, p.isfood, p.isdirect, p.weightvolume, p.contents, p.weight, p.weightKG," & Chr(10)
''''strSQL = strSQL & "p.volume , p.pricecrddesc, p.pricecrddoubleinf, unitpricebase  from cbis599p.dbo.product as P where p.productcode in (" & 1035 & ")" & Chr(10)
'''''Debug.Print strSQL
''''CBA_COM_CBISRS.Open strSQL, CBA_COM_CBISCN ', , , "adAsyncExecute"
''''If CBA_COM_CBISRS.EOF Then
''''    MsgBox "Error in CBIS Generation"
''''Else
''''    inArray = CBA_COM_CBISRS.GetRows()
''''    CBA_COM_genCBISoutput True, inArray
''''End If
''''CBA_COM_CBISRS.Close
''''Set CBA_COM_CBISRS = Nothing
''''
''''End Sub
''''
''''
''''
''''
''''''Sub testerror()
''''''Dim a  As Long
''''''    Dim lngval As Long, Appl
''''''    Dim colModNames As Scripting.Dictionary
''''''    Dim k
'''''''    On Error GoTo ErrHandler
'''''''   Debug.Print "Error handler enabled"
'''''''   lngval = 2 / 0
'''''''    Set Appl = Application.VBE.CodePane(1) '''("CBA_TEST_CBA")
''''''2 ErrHandler:
''''''3   On Error Resume Next
''''''4    lngval = 2 / 0
'''''''   'Debug.Print Err.Description
''''''   'Debug.Print Erl()
'''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.Name
'''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.CountOfLines
'''''''   'Debug.Print Application.VBE.ActiveCodePane.VBE.SelectedVBComponent.Type
'''''''
''''''    Set colModNames = New Scripting.Dictionary
''''''    For a = 1 To Application.VBE.ActiveCodePane.CodeModule.CountOfLines
''''''
''''''        If colModNames.Exists(Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(a, vbext_pk_Proc)) = False _
''''''            And Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(a, vbext_pk_Proc) <> "" Then
''''''           'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(a, vbext_pk_Proc) & "; " & a
''''''            colModNames.Add Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(a, vbext_pk_Proc), a
''''''        End If
''''''    Next
'''''''    For Each k In colModNames
'''''''       'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcBodyLine(k, vbext_pk_Proc) & "---";
'''''''       'Debug.Print k.Name  '''colModNames(k)
'''''''
'''''''    Next
''''''''    For k = 0 To colModNames.Count
''''''''       'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcBodyLine(colModNames(k), vbext_pk_Proc) & "---";
''''''''''       'Debug.Print k.Name  '''colModNames(k)
''''''''
''''''''    Next
''''''
''''''
''''''   'Debug.Print
''''''
''''''''   'Debug.Print "BodyLin" & Application.VBE.ActiveCodePane("CBA_TEST_CBA").ProcBodyLine("testerror", vbext_pk_Proc)
''''''   'Debug.Print "BodyLin" & Application.VBE.ActiveCodePane.CodeModule.ProcBodyLine("testerror", vbext_pk_Proc)
'''''''   'Debug.Print "BodyGet" & Application.VBE.ActiveCodePane.CodeModule.ProcBodyLine("testerror", vbext_pk_Get)
''''''   'Debug.Print "StrtLin" & Application.VBE.ActiveCodePane.CodeModule.ProcStartLine("testerror", vbext_pk_Proc)
'''''''   'Debug.Print "StrtGet" & Application.VBE.ActiveCodePane.CodeModule.ProcStartLine("testerror", vbext_pk_Get)
''''''   'Debug.Print "Cnt Lin" & Application.VBE.ActiveCodePane.CodeModule.ProcCountLines("testerror", vbext_pk_Proc)
'''''''   'Debug.Print "Cnt Get" & Application.VBE.ActiveCodePane.CodeModule.ProcCountLines("testerror", vbext_pk_Get)
''''''   'Debug.Print Erl()
''''''
'''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcCountLines
'''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcOfLine
'''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcStartLine
''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.Lines(Application.VBE.ActiveCodePane.CodeModule.ProcStartLine("testerror", vbext_pk_Proc), 1)
''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.Name
'''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.
''''''
''''''    MsgBox "An error occured in line: " & Erl()
''''''    Err.Clear
''''''    On Error GoTo 0
''''''    On Error GoTo here
''''''    lngval = 2 / 0
''''''
''''''
''''''    a = a
''''''here:
''''''    a = a
''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(42, vbext_pk_Proc)
''''''       'Debug.Print Err.Description
''''''   'Debug.Print Err.Source
''''''   'Debug.Print Err.HelpFile
''''''   'Debug.Print Erl()
''''''End Sub
''''''
''''
''''''Sub testerror2()
''''''Dim a  As Long
''''''    Dim lngval As Long, Appl
''''''    Dim colModNames As scripting.Dictionary
''''''    Dim k
'''''''    On Error GoTo ErrHandler
'''''''   Debug.Print "Error handler enabled"
'''''''   lngval = 2 / 0
''''''
''''''2 ErrHandler:
''''''3   On Error Resume Next
''''''4    lngval = 2 / 0
'''''''   'Debug.Print Err.Description
''''''   'Debug.Print Erl()
'''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.Name
'''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.CountOfLines
'''''''   'Debug.Print Application.VBE.ActiveCodePane.VBE.SelectedVBComponent.Type
'''''''
''''''    Set colModNames = New scripting.Dictionary
''''''    For a = 1 To Application.VBE.ActiveCodePane.CodeModule.CountOfLines
''''''
''''''        If colModNames.Exists(Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(a, vbext_pk_Proc)) = False _
''''''            And Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(a, vbext_pk_Proc) <> "" Then
''''''           'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(a, vbext_pk_Proc) & "; " & a
''''''            colModNames.Add Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(a, vbext_pk_Proc), a
''''''        End If
''''''    Next
'''''''    For Each k In colModNames
'''''''       'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcBodyLine(k, vbext_pk_Proc) & "---";
'''''''       'Debug.Print k.Name  '''colModNames(k)
'''''''
'''''''    Next
''''''''    For k = 0 To colModNames.Count
''''''''       'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcBodyLine(colModNames(k), vbext_pk_Proc) & "---";
''''''''''       'Debug.Print k.Name  '''colModNames(k)
''''''''
''''''''    Next
''''''
''''''
''''''   'Debug.Print
''''''
''''''   'Debug.Print "BodyLin" & Application.VBE.ActiveCodePane("CBA_TEST_CBA").ProcBodyLine("testerror", vbext_pk_Proc)
''''''   'Debug.Print "BodyLin" & Application.VBE.ActiveCodePane.CodeModule.ProcBodyLine("testerror", vbext_pk_Proc)
'''''''   'Debug.Print "BodyGet" & Application.VBE.ActiveCodePane.CodeModule.ProcBodyLine("testerror", vbext_pk_Get)
''''''   'Debug.Print "StrtLin" & Application.VBE.ActiveCodePane.CodeModule.ProcStartLine("testerror", vbext_pk_Proc)
'''''''   'Debug.Print "StrtGet" & Application.VBE.ActiveCodePane.CodeModule.ProcStartLine("testerror", vbext_pk_Get)
''''''   'Debug.Print "Cnt Lin" & Application.VBE.ActiveCodePane.CodeModule.ProcCountLines("testerror", vbext_pk_Proc)
'''''''   'Debug.Print "Cnt Get" & Application.VBE.ActiveCodePane.CodeModule.ProcCountLines("testerror", vbext_pk_Get)
''''''   'Debug.Print Erl()
''''''
'''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcCountLines
'''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcOfLine
'''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcStartLine
''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.Lines(Application.VBE.ActiveCodePane.CodeModule.ProcStartLine("testerror", vbext_pk_Proc), 1)
''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.Name
'''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.
''''''
''''''    MsgBox "An error occured in line: " & Erl()
''''''    Err.Clear
''''''    On Error GoTo 0
''''''    On Error GoTo here
''''''    lngval = 2 / 0
''''''
''''''
''''''    a = a
''''''here:
''''''    a = a
''''''   'Debug.Print Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(42, vbext_pk_Proc)
''''''       'Debug.Print Err.Description
''''''   'Debug.Print Err.Source
''''''   'Debug.Print Err.HelpFile
''''''   'Debug.Print Erl()
''''''End Sub
''''''
''''
''''Public Sub AST_Test()
''''    ' Send the prepared Promotions to the regions
''''    Dim lRegion As Long, dtRegionsDate As Date
''''    Const CFLDS As String = "PD_Product_Code,PD_Future_Prod_Code,PD_CGSCG,PD_CGDesc,PD_GBD,PD_BD,Freq_Items,PD_Product_Desc,PD_On_Sale_Date,PD_End_Date," & _
''''                            "Actual_Marketing_Support,PD_TV,PD_Radio,PD_Out_Of_Home,PD_Press_Dates,PD_Cover_Date,PD_Cannabalised,PD_Complementary," & _
''''                            "PD_Table_Merch,PD_EndCap_Merch,PV_Fill_Qty,PV_Curr_Retail_Price,PV_Retail_Price,PV_Unit_Cost,PV_Supplier_Cost_Support," & _
''''                            "PV_UPSPW,Sales_Multiplier,Expected_Sales,PD_ID,PV_ID,PG_Promo_Desc,PG_Theme,Region,Sts_Seq,PG_ID"
''''    ' Create the xls file to go to the regions
''''    For lRegion = 501 To 509
''''        If lRegion <> 508 Then
''''            dtRegionsDate = CDate(AST_CrtProductRegionSS(CFLDS, "PD_,PV_,PG_", "qry_L3_ProductRegions", "PD_ID,PG_ID,PV_ID", lRegion, ",PD_GBDM_Approved_Date,PDLastUpd,PVLastUpd,PDStatus,PVStatus"))
''''        End If
''''    Next lRegion
''''    MsgBox "Excel files have been created successfully - they will be sent to the regions in approx 30 minutes", vbOKOnly
''''End Sub
''''
''''
''''Public Function t_FixUpliftData()
''''    ' Fix the  Uplift data to 0 if it is not null - Run this on the live......
''''    Dim sSQL As String, lRecs As Long
''''    Dim RS As ADODB.Recordset, CN As ADODB.Connection
''''    Const THISTABLE1 = "CGData"
''''    Const THISTABLE2 = "SCGData"
''''    Set CN = New ADODB.Connection
''''    Set RS = New ADODB.Recordset
''''    With RS
''''        .CursorLocation = adUseClient
''''        .CursorType = adOpenDynamic
''''        .LockType = adLockOptimistic
''''    End With
''''    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_getDB("ForeCast") & ";"
''''    sSQL = "SELECT * FROM " & THISTABLE1
''''    GoSub GSProcess
''''    RS.Close
''''    sSQL = "SELECT * FROM " & THISTABLE2
''''    GoSub GSProcess
''''    MsgBox lRecs & " records processed with Uplift Data"
''''Exit_Routine:
''''    Set RS = Nothing
''''    Set CN = Nothing
''''    Exit Function
''''GSProcess:
''''    RS.Open sSQL, CN
''''    Do While Not RS.EOF
''''        lRecs = lRecs + 1
''''        If IsNull(RS!FUplift) = False Or RS!FRetail > 0 Or RS!FReRetail = 0 Then RS!FUplift = 0
''''        If IsNull(RS!FReUplift) = False Or RS!FReRetail > 0 Then RS!FReUplift = 0
''''        RS.Update
''''        RS.MoveNext
''''    Loop
''''    Return
''''End Function
''''
''''Public Function t_FixAuth()
''''    ' Fix the Authority table
''''    Dim sSQL As String, lRecs As Long, lNum As Long
''''    Dim RS As ADODB.Recordset, CN As ADODB.Connection
''''    Const THISTABLE1 = "A2_FormFieldAuth"
''''    Set CN = New ADODB.Connection
''''    Set RS = New ADODB.Recordset
''''    With RS
''''        .CursorLocation = adUseClient
''''        .CursorType = adOpenDynamic
''''        .LockType = adLockOptimistic
''''    End With
''''    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_getDB("ASYST") & ";"
''''    sSQL = "SELECT * FROM " & THISTABLE1
''''    RS.Open sSQL, CN
''''    Do While Not RS.EOF
''''        lRecs = lRecs + 1
''''        For lNum = 4 To 6
''''            RS("FA_Lock" & lNum) = Replace(RS!FA_Lock0, "0", lNum)
''''            RS("FA_Vis" & lNum) = Replace(RS!FA_Vis0, "0", lNum)
''''        Next
''''        RS.Update
''''        RS.MoveNext
''''    Loop
''''
''''
''''    MsgBox lRecs & " records processed with Auth Data"
''''Exit_Routine:
''''    Set RS = Nothing
''''    Set CN = Nothing
''''    Exit Function
''''GSProcess:
''''
''''    Return
''''End Function
''''
''''''Public Sub t_pullError()
''''''    Dim sProc As String
''''''    sProc = AST_SaveDict(Application.VBE.ActiveCodePane.CodeModule, Application.VBE.ActiveCodePane.CodeModule.ProcStartLine()
''''''    sProc = sProc
''''''End Sub
''''
''''Public Sub t_TransferData(Optional bbOld As Boolean = True)
''''    ' This routine will copy the old SCGData to the new table
''''    ' COULDN'T GET THIS WORKING HERE SO DID IT IN THE DATABASE
''''
''''    Dim CN As ADODB.Connection, RS As ADODB.Recordset
''''    Dim sSQL As String, lCG As Long, lPCls As Long, lSCG As Long, lMn As Long, lYr As Long, lDels As Long, dtDate As Date
''''
''''    Set CN = New ADODB.Connection
''''    Set RS = New ADODB.Recordset
''''    With RS
''''        .CursorLocation = adUseClient
''''        .CursorType = adOpenDynamic
''''        .LockType = adLockOptimistic
''''    End With
''''    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_getDB("ForeCast") & ";"
''''
''''    sSQL = "SELECT * FROM Copy_Of_SCGData"
''''    sSQL = sSQL & " ORDER BY CG,SCG,ProductClass,YearNo,MonthNo,ID DESC,DateTimeSubmitted "
''''    RS.Open sSQL
''''    Do While Not RS.EOF
''''        If lCG <> RS!CG Or lSCG <> RS!scg Or lPCls <> RS!Productclass Then dtDate = CDate("01/01/2000")
''''        If lCG <> RS!CG Or lSCG <> RS!scg Or lPCls <> RS!Productclass Or lMn <> RS!MonthNo Or lYr <> RS!YearNo Then
''''            If dtDate = CDate("01/01/2000") Then
''''                dtDate = Format(RS!DateTimeSubmitted, "dd/mm/yyyy hh:nn")
''''            Else
''''                RS!DateTimeSubmitted = dtDate
''''                RS.Update
''''            End If
''''            lCG = RS!CG: lSCG = RS!scg: lPCls = RS!Productclass: lMn = RS!MonthNo: lYr = RS!YearNo
''''        Else
''''            lDels = lDels + 1
''''            RS.Delete
''''        End If
''''        RS.MoveNext
''''    Loop
''''    RS.Close
''''
''''
''''    sSQL = "INSERT INTO SCGData ( DateTimeSubmitted, CG, SCG, ProductClass, ForecastDate, MonthNo, YearNo, FRetail, FReRetail, FMarginP, FReMarginP, UserName ) " _
''''      & "SELECT s.DateTimeSubmitted, s.CG, s.SCG, s.ProductClass, CDate('01/' & s.monthno & '/' & s.yearno AS FC, s.MonthNo, " _
''''        & "s.YearNo, s.FRetail, s.FReRetail, s.FMarginP, s.FReMarginP, s.UserName " _
''''        & "FROM Copy_Of_SCGData s " _
''''        & " ORDER BY s.DateTimeSubmitted, s.CG, s.SCG, CDate('01/' & s.monthno & '/' & s.yearno );"
''''
''''        RS.Open sSQL, CN
''''
''''Exit_Routine:
''''    MsgBox "Apply has completed", vbOKOnly
''''    On Error Resume Next
''''
''''    CN.Close
''''    Set CN = Nothing
''''    Set RS = Nothing
''''
''''End Sub
''''
''''
''''Public Function t_FixForecastDate()
''''    ' Fix the Forecast Date, do this after fixing the duplicates
''''    Dim sSQL As String, lRecs As Long
''''    Dim RS As ADODB.Recordset, CN As ADODB.Connection
''''    Const THISTABLE2 = "SCGData"
''''    Const THISTABLE3 = "ProductData"
''''    Set CN = New ADODB.Connection
''''    Set RS = New ADODB.Recordset
''''    With RS
''''        .CursorLocation = adUseClient
''''        .CursorType = adOpenDynamic
''''        .LockType = adLockOptimistic
''''    End With
''''    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_getDB("ForeCast") & ";"
''''
''''    sSQL = "SELECT * FROM " & THISTABLE2
''''    GoSub GSProcess
''''    RS.Close
''''    sSQL = "SELECT * FROM " & THISTABLE3
''''    GoSub GSProcess
''''    RS.Close
''''    MsgBox lRecs & " records processed with ForecastDate"
''''Exit_Routine:
''''    Set RS = Nothing
''''    Set CN = Nothing
''''    Exit Function
''''GSProcess:
''''    RS.Open sSQL, CN
''''    Do While Not RS.EOF
''''        lRecs = lRecs + 1
''''        RS!ForecastDate = CDate(g_FixDate("01/" & RS("MonthNo") & "/" & RS("YearNo")))
''''        RS.Update
''''        RS.MoveNext
''''    Loop
''''    Return
''''End Function
''''
''''
''''''    ' Delete any duplicate values in the ProductData table - HAS BEEN RUN
''''''    Dim sSQL As String, lCG As Long, lPCls As Long, lSCG As Long, lPd As Long, lMth As Long, lYr As Long, lDels As Long
''''''    Dim RS As ADODB.Recordset, CN As ADODB.Connection
''''''    Const THISTABLE = "ProductData"
''''''    Set CN = New ADODB.Connection
''''''    Set RS = New ADODB.Recordset
''''''    With RS
''''''        .CursorLocation = adUseClient
''''''        .CursorType = adOpenDynamic
''''''        .LockType = adLockOptimistic
''''''    End With
''''''    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_getDB("ForeCast") & ";"
''''''    sSQL = "SELECT * FROM " & THISTABLE
''''''    sSQL = sSQL & " ORDER BY YearNo,MonthNo,CG,SCG,ProductClass,ProductCode,DateTimeSubmitted Desc"
''''''    RS.Open sSQL, CN
''''''    Do While Not RS.EOF
''''''        If lCG <> RS!CG Or lSCG <> RS!SCG Or lPCls <> RS!ProductClass Or lPd <> RS!ProductCode Or lYr <> RS!YearNo Or lMth <> RS!MonthNo Then
''''''            lCG = RS!CG: lSCG = RS!SCG: lPCls = RS!ProductClass: lPd = RS!ProductCode: lYr = RS!YearNo: lMth = RS!MonthNo
''''''        Else
''''''            lDels = lDels + 1
''''''            RS.Delete
''''''        End If
''''''        RS.MoveNext
''''''
''''''    Loop
''''''    MsgBox lDels & " deleted from " & THISTABLE
''''''Exit_Routine:
''''''    Set RS = Nothing
''''''    Set CN = Nothing
''''''    Exit Function
''''''End Function
''''
''''Public Function t_FixDupValsSCG()
''''    ' Delete any duplicate values in the SCGData table - HAS BEEN RUN
''''    Dim sSQL As String, lCG As Long, lPCls As Long, lSCG As Long, lMth As Long, lYr As Long, lDels As Long
''''    Dim RS As ADODB.Recordset, CN As ADODB.Connection
''''    Const THISTABLE = "SCGData"
''''    Set CN = New ADODB.Connection
''''    Set RS = New ADODB.Recordset
''''    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_getDB("ForeCast") & ";"
''''    With RS
''''        .CursorLocation = adUseClient
''''        .CursorType = adOpenDynamic
''''        .LockType = adLockOptimistic
''''    End With
''''    sSQL = "SELECT * FROM " & THISTABLE
''''    sSQL = sSQL & " ORDER BY YearNo,MonthNo,CG,SCG,ProductClass,DateTimeSubmitted Desc"
''''    RS.Open sSQL, CN
''''    Do While Not RS.EOF
''''        If lCG <> RS!CG Or lSCG <> RS!scg Or lPCls <> RS!Productclass Or lYr <> RS!YearNo Or lMth <> RS!MonthNo Then
''''            lCG = RS!CG: lSCG = RS!scg: lPCls = RS!Productclass: lYr = RS!YearNo: lMth = RS!MonthNo
''''        Else
''''            lDels = lDels + 1
''''            RS.Delete
''''        End If
''''        RS.MoveNext
''''
''''    Loop
''''    MsgBox lDels & " deleted from " & THISTABLE
''''Exit_Routine:
''''    Set RS = Nothing
''''    Set CN = Nothing
''''    Exit Function
''''End Function
''''
''''Public Function t_FixDupValsCG()
''''    ' Delete any duplicate values in the CGData table  - HAS BEEN RUN
''''    Dim sSQL As String, lCG As Long, lPCls As Long, lMth As Long, lYr As Long, lDels As Long
''''    Dim RS As ADODB.Recordset, CN As ADODB.Connection
''''    Const THISTABLE = "CGData"
''''    Set CN = New ADODB.Connection
''''    Set RS = New ADODB.Recordset
''''    With RS
''''        .CursorLocation = adUseClient
''''        .CursorType = adOpenDynamic
''''        .LockType = adLockOptimistic
''''    End With
''''    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_getDB("ForeCast") & ";"
''''    sSQL = "SELECT * FROM " & THISTABLE
''''    sSQL = sSQL & " ORDER BY YearNo,MonthNo,CG,ProductClass,DateTimeSubmitted Desc"
''''    RS.Open sSQL, CN
''''    Do While Not RS.EOF
''''        If lCG <> RS!CG Or lPCls <> RS!Productclass Or lYr <> RS!YearNo Or lMth <> RS!MonthNo Then
''''            lCG = RS!CG: lPCls = RS!Productclass: lYr = RS!YearNo: lMth = RS!MonthNo
''''        Else
''''            lDels = lDels + 1
''''            RS.Delete
''''        End If
''''        RS.MoveNext
''''
''''    Loop
''''    MsgBox lDels & " deleted from " & THISTABLE
''''Exit_Routine:
''''    Set RS = Nothing
''''    Set CN = Nothing
''''    Exit Function
''''
''''End Function
''''
''''Public Function t_FixSubmittedDate()
''''    ' Fix the Submitted Date, do this after fixing the duplicates
''''    Dim sSQL As String, lRecs As Long
''''    Dim RS As ADODB.Recordset, CN As ADODB.Connection
''''    Const THISTABLE1 = "CGData"
''''    Const THISTABLE2 = "SCGData"
''''    Const THISTABLE3 = "ProductData"
''''    Set CN = New ADODB.Connection
''''    Set RS = New ADODB.Recordset
''''    With RS
''''        .CursorLocation = adUseClient
''''        .CursorType = adOpenDynamic
''''        .LockType = adLockOptimistic
''''    End With
''''    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_getDB("ForeCast") & ";"
''''    sSQL = "SELECT * FROM " & THISTABLE1
''''    GoSub GSProcess
''''    RS.Close
''''    sSQL = "SELECT * FROM " & THISTABLE2
''''    GoSub GSProcess
''''    RS.Close
''''    sSQL = "SELECT * FROM " & THISTABLE3
''''    GoSub GSProcess
''''    RS.Close
''''    MsgBox lRecs & " records processed with SubmittedDate"
''''Exit_Routine:
''''    Set RS = Nothing
''''    Set CN = Nothing
''''    Exit Function
''''GSProcess:
''''    RS.Open sSQL, CN
''''    Do While Not RS.EOF
''''        lRecs = lRecs + 1
''''        RS!DateTimeSubmitted = g_FixDate(RS!DateTimeSubmitted, "dd/mm/yyyy hh:nn")
''''        RS.Update
''''        RS.MoveNext
''''    Loop
''''    Return
''''End Function
''''
''''''Public Function t_FixUpliftData()
''''''    ' Fix the  Uplift data to 0 if it is not null
''''''    Dim sSQL As String, lRecs As Long
''''''    Dim RS As ADODB.Recordset, CN As ADODB.Connection
''''''    Const THISTABLE1 = "CGData"
''''''    Const THISTABLE2 = "SCGData"
''''''    Set CN = New ADODB.Connection
''''''    Set RS = New ADODB.Recordset
''''''    With RS
''''''        .CursorLocation = adUseClient
''''''        .CursorType = adOpenDynamic
''''''        .LockType = adLockOptimistic
''''''    End With
''''''    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_getDB("ForeCast") & ";"
''''''    sSQL = "SELECT * FROM " & THISTABLE1
''''''    GoSub GSProcess
''''''    RS.Close
''''''    sSQL = "SELECT * FROM " & THISTABLE2
''''''    GoSub GSProcess
''''''    MsgBox lRecs & " records processed with ForecastDate"
''''''Exit_Routine:
''''''    Set RS = Nothing
''''''    Set CN = Nothing
''''''    Exit Function
''''''GSProcess:
''''''    RS.Open sSQL, CN
''''''    Do While Not RS.EOF
''''''        lRecs = lRecs + 1
''''''        If IsNull(RS!FUplift) = False Then RS!FUplift = 0
''''''        If IsNull(RS!FReUplift) = False Then RS!FReUplift = 0
''''''        RS.Update
''''''        RS.MoveNext
''''''    Loop
''''''    Return
''''''End Function
''''
''''
'''''Private SKUarr() As CBA_COM_COMCompSKU
'''''Sub thisisatest()
'''''testtime = Now
''''''On Error Resume Next
''''''If UBound(SKUarr, 1) < 1 Then
''''''On Error GoTo 0
''''''SKUarr = New CBA_COM_COMCompSKU
'''''    SKUarr = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("WW")
''''''End If
''''''On Error GoTo 0
'''''
'''''Debug.Print "------------"
'''''Debug.Print testtime
'''''Debug.Print Now
'''''Debug.Print "------------"
'''''Debug.Print Now - testtime
'''''Debug.Print "------------"
'''''
'''''
'''''
'''''
'''''Set thiswb = Application.Workbooks.Add
'''''
'''''With thiswb
'''''    .Activate
'''''    For a = LBound(SKUarr) To UBound(SKUarr) ' UBound(SKUarr)
'''''        b = a + 1
'''''        .Sheets(1).Cells(b, 1).Value = SKUarr(a).CBA_COM_SKU_compcode
'''''        .Sheets(1).Cells(b, 2).Value = SKUarr(a).CBA_COM_SKU_CompProdName
'''''        .Sheets(1).Cells(b, 3).Value = SKUarr(a).CBA_COM_SKU_Competitor
'''''        .Sheets(1).Cells(b, 4).Value = SKUarr(a).CBA_COM_SKU_CompPacksize
'''''        .Sheets(1).Cells(b, 5).Value = SKUarr(a).getPriceData("National", "Mean", False)
'''''        .Sheets(1).Cells(b, 6).Value = SKUarr(a).getPriceData("National", "Mean", True)
'''''        .Sheets(1).Cells(b, 7).Value = SKUarr(a).getPriceData("NSW", "Median", False)
'''''        .Sheets(1).Cells(b, 8).Value = SKUarr(a).getPriceData("NSW", "Median", True)
'''''        .Sheets(1).Cells(b, 9).Value = SKUarr(a).getPriceData("VIC", "Mode", False)
'''''        .Sheets(1).Cells(b, 10).Value = SKUarr(a).getPriceData("VIC", "Mode", True)
'''''        .Sheets(1).Cells(b, 11).Value = SKUarr(a).getPriceData("QLD", "Highest", False)
'''''        .Sheets(1).Cells(b, 12).Value = SKUarr(a).getPriceData("QLD", "Highest", True)
'''''        .Sheets(1).Cells(b, 13).Value = SKUarr(a).getPriceData("WA", "Lowest", False)
'''''        .Sheets(1).Cells(b, 14).Value = SKUarr(a).getPriceData("WA", "Lowest", True)
'''''
'''''    Next
'''''End With
'''''
'''''
'''''
'''''
'''''End Sub
'''''Sub testeringing()
'''''Dim j As CBA_BTF_SCG
'''''Set j = New CBA_BTF_SCG
''''
''''
'''''j.createSCG
''''
''''
''''
'''''End Sub
''''
''''
''''''Public Function AST_SaveDict(ByVal Appl, StrLin As Long) As String   Fix up!!!!!!!! works to a large degree but brings back line number of start
''''''
''''''   ''t_Pullerror in Test - is close to working but need the
''''''    Dim lIdx As Long, a As Long, lIdx2 As Long '', sMdlName As String, sProcN As String
''''''    AST_SaveDict = "Unknown"
''''''    For a = 0 To Appl.CountOfLines
''''''        If StrLin < a Then ''Appl.ProcStartLine(a, vbext_pk_Proc) Then                                               ' Start Line in module
''''''            AST_SaveDict = Appl.ProcOfLine(a, vbext_pk_Proc)                                                ' Procedure name
''''''            Exit Function
''''''        End If
''''''    Next
''''''
''''''''    Static lib1(), bHasBeenRun As Boolean, lMaxIdx As Long, bFound As Boolean
''''''''    Static sMdlName As String, sProcN As String
''''''''
''''''''    If bHasBeenRun = False Then
''''''''        lMaxIdx = -1: lIdx = 0
''''''''        GoSub GSApply
''''''''        bHasBeenRun = True
''''''''        GoTo GetItem
''''''''    End If
''''''''
''''''''    For lIdx = 0 To lMaxIdx
''''''''        If lib1(0, lIdx) = Appl.Name Then
''''''''            If lib1(2, lIdx) = Appl.ProcStartLine(lIdx, vbext_pk_Proc) Then
''''''''                sModule = lib1(0, lIdx)
''''''''                sProc = lib1(1, lIdx)
''''''''            End If
''''''''        End If
''''''''    Next
''''''''    sModule = Appl.Name
''''''''    sProc = Appl.ProcCountLines(StrLin, vbext_pk_Proc)
''''''''    End If
''''''
''''''''    Select Case Module_Proc
''''''''GetForecastValue(lidx, "Module"
''''''''        AST_SaveDict = Appl.Name
''''''
''''''GetItem:
''''''
''''''
''''''
''''''    Exit Function
''''''''GSApply:
''''''''    ReDim Preserve lib1(0 To 2, 0 To lIdx)
''''''''
''''''''    lib1(0, lIdx) = Appl.Name                                                                               ' Module name
''''''''    For a = 0 To Appl.Line
''''''''        lib1(1, lIdx) = Appl.ProcOfLine(a, vbext_pk_Proc)                                                   ' Procedure name
''''''''        lib1(2, lIdx) = Appl.ProcStartLine(a, vbext_pk_Proc)                                                ' Start Line in module
''''''''    Next
''''''''
''''''''''        sName = Appl.Name
''''''''''        sProc = Appl.ProcCountLines(StrLin, vbext_pk_Proc)
''''''''''    End If
''''''
''''''End Function
''''
''''
''''''Public Function AST_SaveDict(Appl, StrLin As Long, Module_Proc As String)
''''''
''''''    Static lib1 As Scripting.Dictionary, lib2 As Scripting.Dictionary, lib3 As Scripting.Dictionary, bHasBeenRun As Boolean
''''''    Static sName As String, sProc As String
''''''
''''''    If bHasBeenRun = False Then
''''''        Set lib1 = New Dictionary
''''''        Set lib2 = New Dictionary
''''''        Set lib3 = New Dictionary
''''''        bHasBeenRun = True
''''''    End If
''''''
''''''   If lib1.Exists(Appl.Name) = False Then
''''''        lib1.Add "Module", Appl.Name
''''''        For a = 0 To Appl.CountOfLines
''''''            If lib2.Exists(Appl.ProcOfLine(a, vbext_pk_Proc)) = False And Appl.ProcOfLine(a, vbext_pk_Proc) <> "" Then
''''''                ''Debug.Print Appl.ProcOfLine(a, vbext_pk_Proc) & "; " & a
''''''                lib2.Add Appl.ProcOfLine(a, vbext_pk_Proc), a                                                   ' Start Line in module
''''''                lib3.Add Appl.ProcOfLine(a, vbext_pk_Proc), a + Appl.ProcCountLines(a, vbext_pk_Proc)           ' Lines in Procedure
''''''            End If
''''''        Next
''''''        sName = Appl.Name
''''''        sProc = Appl.ProcCountLines(StrLin, vbext_pk_Proc)
''''''    End If
''''''
''''''    Select Case Module_Proc
''''''GetForecastValue(lidx, "Module"
''''''        AST_SaveDict = Appl.Name
''''''
''''''
''''''
''''''
''''''    Exit Function
''''''End Function
''''''
''''
''''
''''
