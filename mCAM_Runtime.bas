Attribute VB_Name = "mCAM_Runtime"
Option Explicit         ' mCAM_Runtime
Private Const NATIONAL_DB As Long = 599
Private Const RDIVFROM  As Long = 501
Private Const RDIVTO As Long = 509
Private pclsRibbonData As cCAM_Global               ' The DataObject that contains all the data for the Ribbon.
Private pclsActiveDataObject As cCAM_Category       ' The data object that contains all the data for the ribbon's selected category
Private pclsDocHolder As cCBA_DocumentHolder ' Contains all the CBA_documents for the Camera System
Private pcolALLGeneratedDataObjects As Collection   ' A backup of every cCAMERA_Category Object that is created (to save having to generate the objects more than once)
Private pbIsActive As Boolean
''Private pvBaseUsers() As Variant
''Private psdUserList As Scripting.Dictionary
''Private psdEToNUList As Scripting.Dictionary
Private plCurB_ID As Long
Private psOwnerUser As String

Private Enum DataCont
    [_First]
    ContractData = 0
    BraketData = 1
    IncoData = 2
    Pricedata = 3
    [_Last]
End Enum
Private Enum DataWaste
    [_First]
    Markdown = 0
    Loss = 1
    [_Last]
End Enum

    ' The holder of the MAIN code around CAMERA.
    ' All ribbon buttons should route but this should not contact a database.
    ' All datbase connections shold be handled via the data objects (FULL STOP!!!!)
    ' The Runtime is the only thing that can createa a cCAMERA_Category Object

'Callback for cCAMCategoryCGSelection onAction
Sub RibCbo_cCAMCategoryCGSelection(Control As IRibbonControl, id As String, index As Integer)
    Dim sCurACGSel As Byte
    Dim CurCategoryName As String
    If Not bIsActive Then bIsActive = True: Call g_GetDB("Cam", , , True)
    sCurACGSel = cRibbonData.byACGSelected
    cRibbonData.byACGSelected = index
    If sCurACGSel <> cRibbonData.byACGSelected And cRibbonData.lCurCatIndexSelected > -1 Then
        If Not cActiveDataObject Is Nothing Then If cActiveDataObject.bACG = IIf(cRibbonData.byACGSelected = 1, True, False) Then Exit Sub
        Set cActiveDataObject = CreateCAMERADataObject(cRibbonData.GetAllCategoryNames(cRibbonData.lCurCatIndexSelected), , , IIf(cRibbonData.byACGSelected = 1, True, False))
        If cActiveDataObject Is Nothing Then MsgBox "cActiveDataObject Not Created": If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    End If
End Sub

'Callback for cCAMCategoryCGSelection getItemCount
Sub RibCbo_cCAMCategoryCGSelection_ItemCount(Control As IRibbonControl, ByRef returnedVal)
    If Not bIsActive Then bIsActive = True: Call g_GetDB("Cam", , , True)
    returnedVal = 2
End Sub

'Callback for cCAMCategoryCGSelection getItemLabel
Sub RibCbo_cCAMCategoryCGSelection_ItemLabel(Control As IRibbonControl, index As Integer, ByRef returnedVal)
    Dim arr() As String
    If Not bIsActive Then bIsActive = True: Call g_GetDB("Cam", , , True)
    If index = 0 Then returnedVal = "Legacy" Else returnedVal = "ACG"
End Sub
'Callback for CBA_CAMERAManageCategory onAction
Sub RibBtn_ManageCategory(Control As IRibbonControl)
    Dim UF As fCAM_ManageCategory
    If Not bIsActive Then bIsActive = True: Call g_GetDB("Cam", , , True)
    Set UF = New fCAM_ManageCategory
    UF.Show vbModeless
End Sub

'Callback for CBA_CAMERACategorySelection onAction
Sub RibCbo_CategorySelection(Control As IRibbonControl, id As String, index As Integer)
    If Not bIsActive Then bIsActive = True: Call g_GetDB("Cam", , , True)
''    If CBA_User() = "White, Robert" Then
''        If MsgBox("Do you wish to select?", vbYesNo) = vbNo Then Exit Sub
''    End If
    cRibbonData.lCurCatIndexSelected = index
    'CBA_COM_RefreshRibbon "cCAMCategoryCGSelection"
    
    If cRibbonData.byACGSelected < 99 Then
        Set cActiveDataObject = CreateCAMERADataObject(cRibbonData.GetAllCategoryNames(index), , , IIf(cRibbonData.byACGSelected = 1, True, False))
        If cActiveDataObject Is Nothing Then MsgBox "cActiveDataObject Not Created": If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    End If
End Sub

'Private Function checkRibbon(Optional ByVal ForceRebuild As Boolean = False) As Boolean
'    If cRibbonData Is Nothing Or ForceRebuild = True Then Set cRibbonData = New cCAM_Global: RefreshCategorySelectionDropDownOnRibbon
'    checkRibbon = cRibbonData.IsActive
'End Function
'Callback for CBA_CAMERACategorySelection getItemCount
Sub RibCbo_CategorySelection_ItemCount(Control As IRibbonControl, ByRef returnedVal)
    If Not bIsActive Then bIsActive = True: Call g_GetDB("Cam", , , True)
    returnedVal = UBound(cRibbonData.GetAllCategoryNames) + 1
End Sub
'Callback for CBA_CAMERACategorySelection getItemLabel
Sub RibCbo_CategorySelection_ItemLabel(Control As IRibbonControl, index As Integer, ByRef returnedVal)
    If Not bIsActive Then bIsActive = True: Call g_GetDB("Cam", , , True)
    returnedVal = cRibbonData.GetAllCategoryNames(index)
End Sub

Sub RibBtn_LineCount(Control As IRibbonControl)
    Dim UF As fCAM_LineCount
    If ActiveDO Is Nothing Then
        MsgBox "There is no Active Data Object"
        Exit Sub
    End If
    Set UF = GetUserForm("fCAM_LineCount", mCAM_Runtime.ActiveDO.lDoc_ID)
    If UF Is Nothing Then
        Set UF = New fCAM_LineCount
        UF.lDoc_ID = mCAM_Runtime.ActiveDO.lDoc_ID
        If UF.bIsActive = True Then UF.Show vbModeless
    End If
End Sub
'''Callback for CBA_CAMERALineCount onAction
''Sub RibBtn_LineCount(control As IRibbonControl)
''    Dim UF As fCAM_LineCount
''    If cActiveDataObject Is Nothing Then
''        MsgBox "There is no Active Data Object"
''        Exit Sub
''    End If
''    Set UF = New fCAM_LineCount
''    If UF.bIsActive = True Then UF.Show vbModeless
''
''End Sub
'Callback for CBA_CAMERAMarketSegment onAction
Sub RibBtn_MarketSegment(Control As IRibbonControl)
    Dim UF As fCAM_MSegSelect
    If ActiveDO Is Nothing Then
        MsgBox "There is no Active Data Object"
        Exit Sub
    End If
    Set UF = GetUserForm("fCAM_MSegSelect", mCAM_Runtime.ActiveDO.lDoc_ID)
    If UF Is Nothing Then
        Set UF = New fCAM_MSegSelect
        UF.lDoc_ID = mCAM_Runtime.ActiveDO.lDoc_ID
        If UF.bIsActive = True Then UF.Show vbModeless
    Else
        UF.Repaint
    End If
End Sub
'Callback for CBA_CAMERAMarketSegment onAction
''Sub RibBtn_MarketSegment(control As IRibbonControl)
''    Dim UF As fCAM_MSegSelect
''''    If cActiveDataObject Is Nothing Then MsgBox "There is no Active Data Object": Exit Sub          ' @RWCam Uncomment when ready
''    If cActiveDataObject Is Nothing Then
''        MsgBox "There is no Active Data Object"
''        Exit Sub
''    End If
''    Set UF = GetUserForm(mCAM_Runtime.ActiveDO.lDoc_ID, "fCAM_MSegSelect")
''    If UF Is Nothing Then
''        Set UF = New fCAM_MSegSelect
''        If UF.bIsActive = True Then UF.Show vbModeless
''    Else
''        UF.Repaint
''    End If
''End Sub
'Callback for CBA_CAMERAProductAllocation onAction
Sub RibBtn_ProductAllocation(Control As IRibbonControl)
    Dim UF As fCAM_ProdToMSeg
    If ActiveDO Is Nothing Then
        MsgBox "There is no Active Data Object"
        Exit Sub
    End If
    Set UF = GetUserForm("fCam_ProdToMSeg", mCAM_Runtime.ActiveDO.lDoc_ID)
    If UF Is Nothing Then
        Set UF = New fCAM_ProdToMSeg
        UF.lDoc_ID = mCAM_Runtime.ActiveDO.lDoc_ID
        If UF.bIsActive Then UF.Show vbModeless
    End If
End Sub
'''Callback for CBA_CAMERAProductAllocation onAction
''Sub RibBtn_ProductAllocation(control As IRibbonControl)
''    Dim UF As fCAM_ProdToMSeg
''    If cActiveDataObject Is Nothing Then
''        MsgBox "There is no Active Data Object"
''        Exit Sub
''    End If
''''    If cActiveDataObject Is Nothing Then MsgBox "There is no Active Data Object": Exit Sub          ' @RWCam Uncomment when ready
''    Set UF = GetUserForm(mCAM_Runtime.ActiveDO.lDoc_ID, "fCam_ProdToMSeg")
''    If UF Is Nothing Then
''        Set UF = New fCAM_ProdToMSeg
''        If UF.bIsActive Then UF.Show vbModeless
''    End If
''End Sub

'Callback for CBA_CAMERAReporting onAction
Sub RibBtn_Reporting(Control As IRibbonControl)
    Dim UF As fCAM_Parameters
    If Not bIsActive Then bIsActive = True: Call g_GetDB("Cam", , , True)
''    If cActiveDataObject Is Nothing Then MsgBox "There is no Active Data Object": Exit Sub          ' @RWCam Uncomment when ready
    Set UF = New fCAM_Parameters
    UF.Show vbModeless
End Sub

Public Function CategoryObject(Optional ByVal Doc_ID As Long = 0) As cCAM_Category
    Dim v As Variant
    Dim Doc As cCBA_Document
    If Doc_ID = 0 Then
        Set CategoryObject = cActiveDataObject
        Exit Function
    Else
        If clsDocHolder Is Nothing Then
            Set CategoryObject = Nothing
            Exit Function
        Else
            For Each v In clsDocHolder.sdDocumentList
                Set Doc = clsDocHolder.sdDocumentList(v)
                If Doc.lDocumentID = Doc_ID Then
                    Set CategoryObject = Doc.iDocDataObject
                    Exit Function
                End If
            Next
        End If
    End If
    Set CategoryObject = Nothing
End Function
Public Function ActiveDO() As cCAM_Category
    Dim Doc As cCBA_Document
    If clsDocHolder Is Nothing Then
        Set ActiveDO = cActiveDataObject
    ElseIf clsDocHolder.sdLinkDoc.Exists(ActiveWorkbook.Name & "-" & ActiveSheet.CodeName) Then
        Set ActiveDO = clsDocHolder.sdDocumentList(clsDocHolder.sdLinkDoc(ActiveWorkbook.Name & "-" & ActiveSheet.CodeName)).iDocDataObject
    Else
        Set ActiveDO = cActiveDataObject
    End If
End Function

Public Function CreateDocument(ByVal DocType As e_DocuType, Optional DocName As String, Optional ByRef DocID As Long = -1, Optional ByRef cDocObject As iCBA_Doc_Data) As Boolean
    If clsDocHolder Is Nothing Then Set clsDocHolder = New cCBA_DocumentHolder
    Select Case DocType
        Case e_DocuType.eCoreRangeCategoryReview
            If cDocObject Is Nothing Then MsgBox "No Document Data Object Provided": CreateDocument = False: Exit Function
            If DocName = "" Then MsgBox "No Document Name Provided": CreateDocument = False: Exit Function
    End Select
    CreateDocument = clsDocHolder.Add_Document(DocType, DocName, DocID, cDocObject)
End Function

''Public Function RenderDocument(Optional ByRef DocID As Long = -1) As Boolean
''    RenderDocument = clsDocHolder.Render_Document(DocID)
''End Function
''
Public Function ControlObject(ByVal linkID As String, Optional ByRef Target As Range, Optional ByVal sCmd As String, Optional sVal1 As String, Optional lVal1 As Long, Optional lVal2 As Long) As Long
    If sCmd = "Close" Then On Error Resume Next
    ControlObject = clsDocHolder.ControlObject(linkID, Target, sCmd, sVal1, lVal1, lVal2)
End Function

Public Sub RefreshCategorySelectionDropDownOnRibbon()
    CBA_COM_RefreshRibbon "cCAMCategorySelection"
    CBA_COM_RefreshRibbon "cCAMCategoryCGSelection"
End Sub
'Callback for CBA_CAMERALineCount onAction

Public Function GetActiveCAMERADataObject() As cCAM_Category
    If Not cActiveDataObject Is Nothing Then GetActiveCAMERADataObject = cActiveDataObject
End Function

Public Function AddToALLGeneratedDataObjects(ByRef cObj As cCAM_Category) As Boolean
    If ALLGeneratedDataObjects Is Nothing Then Set ALLGeneratedDataObjects = New Collection
    ALLGeneratedDataObjects.Add cObj
End Function

''Public Function getCameraDocuTypes(Optional ByVal DocName As String) As Variant
''Dim arr() As Variant
''Dim dic As Scripting.Dictionary
''Dim v As Variant
''Dim a As Long
''
''    Set dic = New Scripting.Dictionary
''    dic.Add "Core Range Category Review", e_DocuType.eCoreRangeCategoryReview
'''    dic.Add "Produce Category Review", e_DocuType.eProduceCategoryReview
'''    dic.Add "Alcohol Category Review", e_DocuType.eAlcoholCategoryReview
'''    dic.Add "Specials Category Review", e_DocuType.eSpecialsCategoryReview
''    dic.Add "Specials & Seasonal Performance", e_DocuType.eSpecSeasPerformance
''    dic.Add "Topline Category Performance", e_DocuType.eToplineCategoryPerformance
''    dic.Add "Line Count Overview Report", e_DocuType.eLineCountOverviewReport
''    dic.Add "Core Range Performance", e_DocuType.eCoreRangePerformance
''    dic.Add "Market Overview Report", e_DocuType.eMarketOverview
''    dic.Add "Core Range Product Listing", e_DocuType.eCoreRangeProductListing
''
''    If DocName = "" Then
''        ReDim arr(0 To dic.Count - 1)
''        a = -1
''        For Each v In dic
''            a = a + 1: arr(a) = CStr(v)
''        Next
''       getCameraDocuTypes = arr
''    Else
''        getCameraDocuTypes = dic(DocName)
''    End If
''
''End Function
Public Sub CheckDocHolder()
    If clsDocHolder Is Nothing Then Set clsDocHolder = New cCBA_DocumentHolder
End Sub

Public Function GetUserForm(ByVal TagName As String, Optional ByVal lDocID As Long = -1) As Object
Dim temp As Object
'Dim doc As cCBA_Document
    For Each temp In VBA.UserForms
        If lDocID > -1 Then
            If LCase(temp.Tag) = LCase(TagName) And temp.lDoc_ID = lDocID Then
                Set GetUserForm = temp
                Exit Function
            End If
        Else
            If LCase(temp.Tag) = LCase(TagName) Then
                Set GetUserForm = temp
                Exit Function
            End If
        End If
    Next
    Set GetUserForm = Nothing
End Function

Public Function CreateCAMERADataObject(ByVal CategoryName As String, Optional ByVal DateFrom As Date, Optional ByVal DateTo As Date, Optional ByVal ACG As Boolean = False) As cCAM_Category
Dim CDO As cCAM_Category
    On Error GoTo Err_Routine
    CBA_Error = ""
        If DateFrom = 0 Then DateFrom = DateSerial(Year(Date) - 2, 1, 1)
        If DateTo = 0 Then DateTo = IIf(Month(Date) > 1, DateSerial(Year(Date), Month(Date), 0), DateSerial(Year(Date) - 1, 12, 31))
        If ALLGeneratedDataObjects Is Nothing Then Set ALLGeneratedDataObjects = New Collection
        For Each CDO In ALLGeneratedDataObjects
            If CDO.sCategoryName = CategoryName And CDO.dtDateFrom = DateFrom And CDO.dtDateTo = DateTo And CDO.bACG = ACG Then Set CreateCAMERADataObject = CDO: Exit Function
        Next
        If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "CAMERA Category Data Generation"
        ' the object will need to pull both ACG and LCG and hold two PrdGrp Items??
        If cRibbonData.GetCategoryList(CategoryName, ACG) Is Nothing Then
            MsgBox "There is no allocation of CG/SCG's against the " & IIf(ACG = False, "Legacy", "ACG") & " Section of the Category. Please select a different CG Structure or add ACG selections to the category via the Manage Category button. "
            Set CreateCAMERADataObject = Nothing
            Exit Function
        End If
        Set CDO = New cCAM_Category
        If CDO.Construct(CategoryName, DateFrom, DateTo, ACG) = False Then Set CreateCAMERADataObject = Nothing: Exit Function
        ALLGeneratedDataObjects.Add CDO
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
        Set CreateCAMERADataObject = CDO
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-mCAM_Runtime.createCAMERADataObject", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Function InsertCatRevVariables(ByVal i_ACG As Boolean, ByVal i_ExecSum As Boolean, ByVal i_CatPerfOver As Boolean, _
    ByVal i_TotCatPerf As Boolean, ByVal i_MarkOver As Boolean, ByVal i_CRTrialPerf As Boolean, ByVal i_CRPerf As Boolean, _
        ByVal i_SeasPerf As Boolean, ByVal i_SpecPerf As Boolean, ByVal i_Wastage As Boolean, ByVal i_LineCount As Boolean, _
            ByVal i_Forecast As Boolean, ByVal i_InterComp As Boolean, ByVal i_CompAnal As Boolean, ByVal i_CatRevActions As Boolean, _
                ByVal i_CatStrat As Boolean, ByVal i_SuppBase As Boolean, ByVal i_CRPerfProds As Boolean, ByVal i_QA As Boolean, _
                    ByVal i_CRTrialProdList As Boolean, ByVal i_CorpResp As Boolean, ByVal i_CatPlano As Boolean, ByVal i_CommUpd As Boolean, _
                        ByVal i_KPIs As Boolean, ByVal i_Mnthfrom As Date, ByVal i_MnthTo As Date, ByRef i_CGSelection As Collection) As Boolean
''    ' @RWCam  Took out - Not set up to work as vars not defed Uncommented
''
''    ACG = i_ACG: ExecSum = i_ExecSum: CatPerfOver = i_CatPerfOver: TotCatPerf = i_TotCatPerf: MarkOver = i_MarkOver
''    CRTrialPerf = i_CRTrialPerf: CRPerf = i_CRPerf: SeasPerf = i_SeasPerf: SpecPerf = i_SpecPerf: Wastage = i_Wastage
''    LineCount = i_LineCount: Forecast = i_Forecast: InterComp = i_Forecast: CompAnal = i_CompAnal: CatRevActions = i_CatRevActions
''    CatStrat = i_CatStrat: SuppBase = i_SuppBase: CRPerfProds = i_CRPerfProds: QA = i_QA: CRTrialProdList = i_CRTrialProdList
''    CorpResp = i_CorpResp: CatPlano = i_CatPlano: CommUpd = i_CommUpd: KPIs = i_KPIs: Mnthfrom = i_Mnthfrom: MnthTo = i_MnthTo
''    Set CGSelection = i_CGSelection
''
''    B_RdyToRun = True
''    InsertCatRevVariables = B_RdyToRun

End Function

Function getCBA_ProdEntity(ByVal PCode As String) As cCBA_Prod
    If cActiveDataObject Is Nothing Then getCBA_ProdEntity = Nothing: Exit Function
    If cActiveDataObject.cPGrp.getProductState(PCode) = False Then getCBA_ProdEntity = Nothing: Exit Function
    Set getCBA_ProdEntity = cActiveDataObject.cPGrp.getProdObject(PCode)
End Function
Function GetCGListing(ByVal ACG As Boolean) As Variant
    If ACG = True Then GetCGListing = pclsRibbonData.sdACGList Else Set GetCGListing = pclsRibbonData.sdLCGList
End Function
Function LoadFrontForm() As Boolean
    LoadFrontForm = True
    If pclsActiveDataObject.CreateUserList() = True And ChangePermissions(, True) = True Then
                fCam_FrontForm.GenerateAndDisplay pclsActiveDataObject.vBaseUsers, pclsActiveDataObject.sdUserList, pclsActiveDataObject.sdEToNUList, CBA_User
    Else
        MsgBox "Error in LoadFrontForm"
    End If
End Function

Function ChangePermissions(Optional sdDic As Scripting.Dictionary, Optional bUseZero As Boolean = False) As Boolean
    Dim lTmpB_ID As Long
    If bUseZero = True Then lTmpB_ID = -1 Else lTmpB_ID = lCurB_ID
    ChangePermissions = pclsActiveDataObject.ChangePermissions(sdDic, lTmpB_ID)
End Function

Function LoadPermissionsForm() As Boolean
    Dim a As Long, b As Long
    Dim bfound As Boolean
    Dim AddUserCol As Scripting.Dictionary
    Dim vBaseUsers As Variant                  '', sdUserList As Scripting.Dictionary, EToNUList As Scripting.Dictionary
    vBaseUsers = pclsActiveDataObject.vBaseUsers

    If pclsActiveDataObject.CreateUserList = False Then MsgBox "UserList Not Created", vbOKOnly: Exit Function
    For a = LBound(vBaseUsers, 2) To UBound(vBaseUsers, 2)
        If vBaseUsers(0, a) = lCurB_ID Then
            bfound = True
            Set AddUserCol = New Scripting.Dictionary
            For b = 8 To UBound(vBaseUsers, 1)
                If IsNull(vBaseUsers(b, a)) = False Then
                    AddUserCol.Add pclsActiveDataObject.sdEToNUList(CStr(vBaseUsers(b, a))), vBaseUsers(b, a)
                End If
            Next
            Exit For
        End If
    Next
    If bfound = False Then MsgBox "CrossReference of B_ID not found", vbOKOnly: Exit Function
    fCam_Permissions.Loadup CBA_User, pclsActiveDataObject.sdUserList, AddUserCol, sOwnerUser
    LoadPermissionsForm = True
End Function
Function setCurID(ByVal id As Long, ByVal OUsr As String)
    lCurB_ID = id
    sOwnerUser = OUsr
End Function
Function setOwnerUser(ByVal OUsr As String)
    sOwnerUser = OUsr
End Function

Public Property Get cRibbonData() As cCAM_Global
    If pclsRibbonData Is Nothing Then Set pclsRibbonData = New cCAM_Global: RefreshCategorySelectionDropDownOnRibbon
    Set cRibbonData = pclsRibbonData
End Property
Private Property Set cRibbonData(ByVal objNewValue As cCAM_Global): Set pclsRibbonData = objNewValue: End Property

Public Property Get cActiveDataObject() As cCAM_Category: Set cActiveDataObject = pclsActiveDataObject: End Property
Private Property Set cActiveDataObject(ByVal objNewValue As cCAM_Category): Set pclsActiveDataObject = objNewValue: End Property

Public Property Get cDocHolder() As cCBA_DocumentHolder: Set cDocHolder = pclsDocHolder: End Property
Private Property Set cDocHolder(ByVal objNewValue As cCBA_DocumentHolder): Set pclsDocHolder = objNewValue: End Property

Public Property Get ALLGeneratedDataObjects() As Collection: Set ALLGeneratedDataObjects = pcolALLGeneratedDataObjects: End Property
Private Property Set ALLGeneratedDataObjects(ByVal objNewValue As Collection): Set pcolALLGeneratedDataObjects = objNewValue: End Property

Public Property Get clsDocHolder() As cCBA_DocumentHolder: Set clsDocHolder = pclsDocHolder: End Property
Private Property Set clsDocHolder(ByVal objNewValue As cCBA_DocumentHolder): Set pclsDocHolder = objNewValue: End Property

Private Property Get lCurB_ID() As Long: lCurB_ID = plCurB_ID: End Property
Private Property Let lCurB_ID(ByVal lNewValue As Long): plCurB_ID = lNewValue: End Property

Private Property Get sOwnerUser() As String: sOwnerUser = psOwnerUser: End Property
Private Property Let sOwnerUser(ByVal sNewValue As String): psOwnerUser = sNewValue: End Property

'Private Function RunConnectionSetups(ByVal CBIS As Boolean, ByVal MMS As Boolean, ByVal Spaceman As Boolean, _
'    ByVal COMRADE_STAR As Boolean, ByVal CAT_REV As Boolean) As Boolean
'Dim div As Long
''''''''-----------TEST VARIABLES---------------------
''Variables for Development
'Set RegDivMiss = New Scripting.Dictionary
'RegDivMiss.add 508, 508
''End of test variables
''''''''-----------END TEST VARIABLES---------------------
'RunConnectionSetups = True
'If COMRADE_STAR = True Then
'    If SCN Is Nothing Then
'        Set SCN = New ADODB.Connection
'        With SCN
'            .ConnectionTimeout = 50
'            .CommandTimeout = 50
'            .Open "Provider=SQLNCLI10;DATA SOURCE=" & TranslateServerName("599DBL12",Date) & ";;INTEGRATED SECURITY=sspi;"
'        End With
'    End If
'    If SCN.State <> 1 Then RunConnectionSetups = False
'End If
'If Spaceman = True Then
'    If SMCN Is Nothing Then
'        Set SMCN = New ADODB.Connection
'        With SMCN
'            .ConnectionTimeout = 50
'            .CommandTimeout = 50
'            .Open "Provider=SQLNCLI10;DATA SOURCE=599DBL11;;INTEGRATED SECURITY=sspi;"
'        End With
'    End If
'    If SMCN.State <> 1 Then RunConnectionSetups = False
'End If
'If CBIS = True Then
'    If CCN Is Nothing Then
'        Set CCN = New ADODB.Connection
'        With CCN
'            .ConnectionTimeout = 50
'            .CommandTimeout = 50
''            .Open "Provider=SQLOLEDB;DATA SOURCE=" & NATIONAL_DB & "DBL01;;INTEGRATED SECURITY=sspi;DataTypeCompatibility=80"
'            .Open "Provider=SQLNCLI10;DATA SOURCE=" & NATIONAL_DB & "DBL01;;INTEGRATED SECURITY=sspi;"
'        End With
'    End If
'    If CCN.State <> 1 Then RunConnectionSetups = False
'End If
'If MMS = True Then
'    For div = RDIVFROM To RDIVTO
'        If RCN(div) Is Nothing Then
'            If RegDivMiss.Exists(div) = False Then
'                Set RCN(div) = New ADODB.Connection
'                With RCN(div)
'                    .ConnectionTimeout = 50
'                    .CommandTimeout = 50
'                    .Open "Provider=SQLNCLI10;DATA SOURCE=0" & div & "Z0IDBSRVL02;;INTEGRATED SECURITY=sspi;"
'                End With
'            End If
'        End If
'        If RegDivMiss.Exists(div) = False Then If RCN(div).State <> 1 Then RunConnectionSetups = False
'    Next
'End If
'If CAT_REV = True Then
'    If CRCN Is Nothing Then
'        Set CRCN = New ADODB.Connection
'        With CRCN
'            .ConnectionTimeout = 50
'            .CommandTimeout = 50
'            .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & CBA_BSA & "LIVE DATABASES\CBCAT_REV_BE.accdb"
'        End With
'    End If
'    If CRCN.State <> 1 Then RunConnectionSetups = False
'End If
'
'
'End Function
''''********************************************************************''''''''''''
''''********************************************************************''''''''''''
''''********************************************************************''''''''''''
''MULTITHREADING IDEAS''
'Function splitProcessCompilePlist(ByRef PList As Variant) As Scripting.Dictionary
'Dim lnt As Long, a As Long
'Dim col(0 To 3) As Collection
'Dim xlApp(0 To 3) As Excel.Application
'Dim op As Scripting.Dictionary
'For a = 0 To 3
'    Set xlApp(a) = New Excel.Application
'    Set col(a) = New Collection
'Next
'
'    lnt = UBound(PList, 2)
'    For a = LBound(PList, 2) To UBound(PList, 2)
'        If a < lnt / 4 Then
'            col(0).add PList(0, a)
'        ElseIf a < lnt / 2 Then
'            col(1).add PList(0, a)
'        ElseIf a < lnt - (lnt / 4) Then
'            col(2).add PList(0, a)
'        Else
'            col(3).add PList(0, a)
'        End If
'    Next
'a = a
'
'Set ad = xlApp(0).AddIns("CBStdAddinW10V.xlam")
'For a = 0 To 3
'    For Each ad In xlApp(a).AddIns
'        If ad.Name = "CBStdAddinW10V.xlam" Then Set Adi(a) = ad: Exit For
'    Next
'    Set op(a) = xlApp(0).Run("!mCAM_Runtime.buildCBA_PcodeDic(col(0))")
'Next
'
'End Function
'
'Function buildCBA_PcodeDic(ByRef Pcol As Collection) As Scripting.Dictionary
'Dim ProductDic As Scripting.Dictionary
'Dim curRnd As Long, ter As Single
'    Set ProductDic = New Scripting.Dictionary
'    curRnd = 0: ter = Timer
'    For a = LBound(PList, 2) To UBound(PList, 2)
'        If Round((a / UBound(PList, 2)), 2) > curRnd Then Rate = (Timer - ter): curRnd = Round((a / UBound(PList, 2)), 2): ter = Timer
'        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 1, 1, "Generating ALDI Product DataCube: " & Round((a / UBound(PList, 2)), 2) * 100 & "%  : Est. Time to Complete: " & CStr(Round((Rate * (1 - Round((a / UBound(PList, 2)), 2))) * 100 / 60, 2)) & " minutes"
'        Set C_P = New cCBA_Prod
'        If IsEmpty(COMdic(CStr(PList(0, a)))) Then Set col = Nothing Else Set col = COMdic(CStr(PList(0, a)))
'        C_P.Build PList(0, a), buildcode, TotBus, pullBaseData(PList(0, a), buildcode), _
'            pullSalesData(PList(0, a), buildcode), RCost, RegDivMiss, NATIONAL_DB, RDIVFROM, RDIVTO, col
'        ProductDic.add CStr(PList(0, a)), C_P
'    Next
'
'End Function




' TESTING FUNCTIONS FROM HERE DOWN
'Sub TESTLineCountForm()
    'Dim Dic As Scripting.Dictionary
    'Set Dic = New Scripting.Dictionary
    'Set LineCountFormDic = Nothing
    '
    'Dic.add CStr(1111), "Core"
    'Dic.add CStr(1112), "Core"
    'Dic.add CStr(1113), "Core"
    'Dic.add CStr(1114), "Core"
    'Dic.add CStr(1115), "Core"
    'Dic.add CStr(1116), "Brand"
    'Dic.add CStr(1117), "Brand"
    'Dic.add CStr(1118), "Spec"
    'Dic.add CStr(1119), "Spec"
    'Dic.add CStr(1120), "TrialCur"
    '
    'OpenLineCountForm "Test", Dic
    '
'End Sub
'Sub TESTDOCHOLDER()
'    CBA_DocHolder.add CoreRangeCategoryReview, "TESTCATREV"
'End Sub
'Sub TEstrenderingTheDoc()
'    CBA_DocHolder.add CoreRangeCategoryReview, "TESTCATREV"
'    CBA_DocHolder.Render_Document "TESTCATREV"
'End Sub
'
'Sub testUTArrayTick()
    'Dim a As Long
'        mTEN_Runtime.Generate_TD
'        mTEN_Runtime.Get_db_TD_TD True
'        'Set CellObjectsIndex = mTEN_Runtime.pullpsdUT
'
    'CellObjects = mTEN_Runtime.pullpUT_Ar1
    '
    'For a = LBound(CellObjects) To UBound(CellObjects)
    '    Debug.Print CellObjects(a).Get_Class_UT("UT_ID")
    'Next
    '
'End Sub
'
'Sub TESTLineCountAndProdDetailForm()
    'Dim a As Long
    'Dim colCG As Collection
    'Dim Lcnt As fCam_LineCount
    'Dim addi As AddIn
    'Dim Doc As cCBA_Document
    'Mnthfrom = CDate(#8/1/2018#)
    'MnthTo = CDate(#7/31/2019#)
    'Set frm_LCnt = New Scripting.Dictionary
    'Set PG = New cCBA_ProdGroup
    'Set Lcnt = New fCam_LineCount: frm_LCnt.add "MAT", Lcnt
    'Set Lcnt = New fCam_LineCount: frm_LCnt.add "MATPY", Lcnt
    'Set Lcnt = New fCam_LineCount: frm_LCnt.add CStr(Year(Mnthfrom) - 1), Lcnt
    'Set Lcnt = New fCam_LineCount: frm_LCnt.add CStr(Year(Mnthfrom) - 2), Lcnt
    'Set frm_MSeg = New fCam_MSegSelect
    'Set frm_MSegAllocate = New fCam_ProdToMSeg
    ''ReDim CGDist(0 To 1, 0 To 0)
    ''CGDist(0, 0) = 53
    ''CGDist(1, 0) = 4
    'Set colCG = New Collection
    'colCG.add "05304"
    '
    '
    '
    'If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "TEST RUN"
    'If PG.RunDataGeneration(Mnthfrom, MnthTo, False, colCG, "BCRPTLW") = False Then
    '    MsgBox "Error in generation of the ProductData"
    '    Exit Sub
    'End If
    'If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.CBA_Close_Running
    '
    ''CreateInitalLineCount CDate(#8/1/2018#), CDate(#7/31/2019#)
    '
    'frm_LCnt("MAT").populateListBoxes "MAT", CreateInitalLineCount(CDate(#8/1/2018#), CDate(#7/31/2019#))
    '
    'frm_LCnt("MAT").Show 'this will be modeless eventually
    ''OpenLineCountForm "Test", CreateInitalLineCount(CDate(#8/1/2018#), CDate(#7/31/2019#))
    '
    '
    'CBA_DocHolder.add CoreRangeCategoryReview, "TESTCATREV"
    '
    '
    'fillCellObjectsWithCorrectValues
    '
    '
    ''Set Doc.UTs = CellObjects
    '
    'CBA_DocHolder.Render_Document "TESTCATREV"
    ''CloseConnections True, True, True, True, True
    'RenderTable
'End Sub

'Function RenderTable()
'Dim Sht As Worksheet
'mTEN_Runtime.Generate_TD
'mTEN_Runtime.Get_db_TD_TD True
'
'Set CellObjectsIndex = mTEN_Runtime.pullpsdUT
'CellObjects = mTEN_Runtime.pullpUT_Ar1
'fillCellObjectsWithCorrectValues
'
'mTEN_Runtime.CrtTenderDoc 4, "Core Range CATREV"
'Set Sht = ActiveSheet
'Sht.Name = "Cat Rev"
'
'
'
'End Function
'

'Function CreateNewCatRev() As Boolean
'Dim a As Long
'Dim scg As Variant
'Dim strPlist As String
'
'
'    fCam_NewCatRev.Show
'    If B_RdyToRun = True Then
'        Debug.Print "Start: " & Now
'
'
'
'        ReDim CGDist(0 To 1, 0 To CGSelection.Count)
'        a = -1
'        For Each scg In CGSelection
'            a = a + 1
'            CGDist(0, a) = Mid(scg, 1, InStr(1, scg, "-") - 1)
'            CGDist(1, a) = Mid(scg, InStr(1, scg, "-") + 1, 99)
'        Next
'
'
'        If TotCatPerf = True Then AddToBuildCode "R": AddToBuildCode "P": AddToBuildCode "B": AddToBuildCode "R": AddToBuildCode "T": AddToBuildCode "c": AddToBuildCode "F": AddToBuildCode "L"
'
'
'
'        If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
'        Set CG_PDic = GenProductDic
'        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
'
'
'
'
'
'
'
'        CreateNewCatRev = True
'    End If
'End Function

'Function setCGDist(ByRef colCGs As Collection) As Boolean
'Dim v As Variant
'Dim a As Long
'    a = -1
'    ReDim CGDist(0 To 1, 0 To colCGs.Count)
'    For Each v In colCGs
'        a = a + 1
'        CGDist(0, a) = Mid(v, 1, InStr(1, v, "-") - 1)
'        CGDist(1, a) = Mid(v, InStr(1, v, "-") + 1, 9)
'    Next
'
'End Function
'
'
'
'Function fillCellObjectsWithCorrectValues() As Boolean
'Dim UT As cCBA_UDT
'Dim var As Variant, v As Variant
'Dim grpNo As Long, a As Long
'Dim Pd As cCBA_Prod
'Dim OutPutNo As Single
'Dim DF(0 To 3) As Date, DT(0 To 3) As Date
'
'
'''filling the TotalCategoryPerformance UT Values
'''    77901 - 78803 = Group Line in Total Category Performance
'
'    DF(0) = CDate(Year(MnthTo) - 3 & "-01-01"): DF(1) = CDate(Year(MnthTo) - 2 & "-01-01"): DF(2) = CDate(Year(MnthTo) - 1 & "-01-01"): DF(3) = DateAdd("D", 1, DateAdd("YYYY", -1, MnthTo))
'    DT(0) = CDate(Year(MnthTo) - 3 & "-12-31"): DT(1) = CDate(Year(MnthTo) - 2 & "-12-31"): DT(2) = CDate(Year(MnthTo) - 1 & "-12-31"): DT(3) = MnthTo
'
'    CellObjects(CellObjectsIndex(77901)).Upd_Class_UT "UT_Default_Value", Year(MnthTo) - 2
'    CellObjects(CellObjectsIndex(77902)).Upd_Class_UT "UT_Default_Value", Year(MnthTo) - 1
'    CellObjects(CellObjectsIndex(77903)).Upd_Class_UT "UT_Default_Value", "MAT " & Format(MnthTo, "MM/YYYY")
'    For a = 78001 To 78003
'        CellObjects(CellObjectsIndex(a)).Upd_Class_UT "UT_Default_Value", PG.getPOS(DF(a - 78000), DT(a - 78000), eQty)
'        var = getProdArrayData("getPOSdata", , , , , , , , , , , , , True, , , , , , DF(a - 78001), DT(a - 78001))
'        If var = 0 Then CellObjects(CellObjectsIndex(a + 100)).Upd_Class_UT "UT_Default_Value", "100"
'        If var <> 0 Then CellObjects(CellObjectsIndex(a + 100)).Upd_Class_UT "UT_Default_Value", ((CellObjects(CellObjectsIndex(a)).Get_Class_UT("UT_Default_Value") - var) / var) * 100
'        CellObjects(CellObjectsIndex(a + 200)).Upd_Class_UT "UT_Default_Value", getProdArrayData("getPOSdata", , , , , , , , , , , , , False, , , , , , DF(a - 78000), DT(a - 78000))
'        var = getProdArrayData("getPOSdata", , , , , , , , , , , , , False, , , , , , DF(a - 78001), DT(a - 78001))
'        If var = 0 Then CellObjects(CellObjectsIndex(a + 300)).Upd_Class_UT "UT_Default_Value", "100"
'        If var <> 0 Then CellObjects(CellObjectsIndex(a + 300)).Upd_Class_UT "UT_Default_Value", ((CellObjects(CellObjectsIndex(a + 200)).Get_Class_UT("UT_Default_Value") - var) / var) * 100
'        var = getProdArrayData("getRCVdata", , , , , , , "Retail", , , , , , , , , , , , DF(a - 78000), DT(a - 78000))
'        CellObjects(CellObjectsIndex(a + 700)).Upd_Class_UT "UT_Default_Value", getProdArrayData("getRCVContribution", , , , , , , , , , , , , , , , , , , DF(a - 78000), DT(a - 78000))
'        If var = 0 Then CellObjects(CellObjectsIndex(a + 400)).Upd_Class_UT "UT_Default_Value", "-"
'        If var <> 0 Then CellObjects(CellObjectsIndex(a + 400)).Upd_Class_UT "UT_Default_Value", getProdArrayData("getRCVContribution", , , , , , , , , , , , , , , , , , , DF(a - 78000), DT(a - 78000)) / var
'        var = getProdArrayData("getRCVdata", , , , , , , "Retail", , , , , , , , , , , , DF(a - 78001), DT(a - 78001))
'        If var = 0 Then CellObjects(CellObjectsIndex(a + 500)).Upd_Class_UT "UT_Default_Value", "-"
'        If var <> 0 Then CellObjects(CellObjectsIndex(a + 500)).Upd_Class_UT "UT_Default_Value", CellObjects(CellObjectsIndex(a + 400)).Get_Class_UT("UT_Default_Value") - (getProdArrayData("getRCVContribution", , , , , , , , , , , , , , , , , , , DF(a - 78000), DT(a - 78000)) / var)
'        var = getProdArrayData("getTotBusData", , , , , , , "POSRetail", , , , , , , , , , , , DF(a - 78000), DT(a - 78000))
'        If var = 0 Then CellObjects(CellObjectsIndex(a + 600)).Upd_Class_UT "UT_Default_Value", 0
'        If var <> 0 Then CellObjects(CellObjectsIndex(a + 600)).Upd_Class_UT "UT_Default_Value", CellObjects(CellObjectsIndex(a + 200)).Get_Class_UT("UT_Default_Value") / var
'    Next
'
'
'    mTEN_Runtime.putpUT_Ar1 CellObjects
'
'
'
'
'
'
'
'
'End Function

'Public Function Route_Document(sAction As String, DocName As String, Optional Target)
'    'Dim bSuccess As Boolean
'
''    Select Case sAction
''        Case "DeleteObjectLink"
''            If cDocHolder.DeleteObjectWorkbookLink(-1) = False Then
''                MsgBox "Workbook didn't delete links correctly", vbOKOnly
''            End If
''        Case "ControlObject"
''
''    End Select
'    ''Debug.Print DocName & " " & sAction
'End Function
'Function getCRDateFrom() As Date
'    getCRDateFrom = Mnthfrom
'End Function
'Function getCRDateTo() As Date
'    getCRDateTo = MnthTo
'End Function
'

'
''Sub OpenLineCountForm(ByVal Period As String, Optional ByVal LineCountDic As scripting.Dictionary)
'''Dim LCF As fCam_LineCount
''    If LineCountFormDic Is Nothing Then Set LineCountFormDic = New scripting.Dictionary
''    If LineCountFormDic.Exists(Period) = False Then
''        If LineCountDic Is Nothing Then
''            MsgBox "Error as no LineCountDic passed to OpenLineCountForm sub", vbOKOnly
''            Exit Sub
''        End If
''        Set LCF = New fCam_LineCount
''        LineCountFormDic.Add Period, LCF
''        LineCountFormDic(Period).populateListBoxes Period, LineCountDic
''    End If
''    LineCountFormDic(Period).Show
''    Set LCF = New fCam_LineCount
''    LCF.populateListBoxes Period, LineCountDic
''    LCF.Show
''End Sub
'Function SetLineCountDicValues(ByVal Period As String, ByVal LineCountForm As fCam_LineCount) As Boolean
'    If LineCountFormDic Is Nothing Then
'        Set LineCountFormDic = New Scripting.Dictionary
'    End If
'    LineCountFormDic.add Period, LineCountForm
'    SetLineCountDicValues = True
'End Function
'Function getProdArrayData(ByVal QueryType As String, Optional ByVal LongVal1 As Long, Optional ByVal LongVal2 As Long _
'    , Optional ByVal LongVal3 As Long, Optional ByVal LongVal4 As Long, Optional ByVal LongVal5 As Long, Optional ByVal LongVal6 As Long _
'    , Optional ByVal StringVal1 As String, Optional ByVal StringVal2 As String, Optional ByVal StringVal3 As String, Optional ByVal StringVal4 As String _
'    , Optional ByVal StringVal5 As String, Optional ByVal StringVal6 As String, Optional ByVal BooleanVal1 As Boolean, Optional ByVal BooleanVal2 As Boolean _
'    , Optional ByVal BooleanVal3 As Boolean, Optional ByVal BooleanVal4 As Boolean, Optional ByVal BooleanVal5 As Boolean, Optional ByVal BooleanVal6 As Boolean _
'    , Optional ByVal DateVal1 As Date, Optional ByVal DateVal2 As Date, Optional ByVal DateVal3 As Date, Optional ByRef ObjectVal1 As Object, Optional ByRef ObjectVal2 As Object) As Variant
    'Dim v As Variant
    'Dim OutPutVal As Variant
    'Dim Pd As cCBA_Prod
'
''For Each v In CG_PDic
''    Set Pd = CG_PDic(v)
''    Select Case QueryType
''        Case "getPOSdata"
''            OutPutVal = OutPutVal + Pd.getPOSdata(DateVal1, DateVal2, BooleanVal1)
''        Case "getRCVContribution"
''            OutPutVal = OutPutVal + Pd.getRCVContribution(DateVal1, DateVal2)
''        Case "getRCVdata"
''            OutPutVal = OutPutVal + Pd.getRCVdata(DateVal1, DateVal2, StringVal1)
''        Case "AveCatUSW"
''            'NEED TO DO A LOT OF WORK TO GET THIS RIGHT
''            'OutPutVal = OutPutVal + Pd.getRCVdata(DateVal1, DateVal2, "USW")
''    End Select
''Next
''    getProdArrayData = OutPutVal
'End Function
'Function AddToBuildCode(ByVal RefVal As String) As Boolean
'    If InStr(1, buildcode, RefVal) = 0 Then buildcode = buildcode & RefVal
'End Function
'

Private Property Get bIsActive() As Boolean:
    bIsActive = pbIsActive
End Property
Private Property Let bIsActive(ByVal bNewValue As Boolean): pbIsActive = bNewValue: End Property
