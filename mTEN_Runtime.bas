Attribute VB_Name = "mTEN_Runtime"
Option Explicit       ' mTEN_Runtime Changed 191101
Option Private Module
' NOTE : plID_Idx and plLV_Idx (being indexes of Addable Access table arrays (pID_Ary and pLV_Ary)), will be 1 based arrays - all the rest are 0 based arrays
Private plSH_Idx As Long, plSL_Idx As Long, plUT_Idx1 As Long, plUT_Idx2 As Long, plID_Idx As Long, plLV_Idx As Long, plIDLV_Idx As Long, plUTL_Idx As Long, plATP_Idx As Long, plATP_Cols As Long
Private pbPFCreated As Boolean, psPF_No As String, psPF_Ver_No As String, plATPLV_ID As Long, plCARows As Long
Private plSH_ID As Long, plLV_ID As Long, plID_ID As Long, plSU_ID As Long, plTH_ID As Long, plAH_ID As Long
Private psNPD_R_MSO As String, psSH_Desc As String, pwbWSht As Worksheet, psWShtName As String ', psWBkName As String, pwbWBk As Workbook
Private psLCG As String, psLSCG As String, psACG As String, psASCG As String, pdteFrom As Date, pdteTo As Date, psActWBkName As String, plWShtNo As Long, psWShtLast As Worksheet

Private psInits As String, plTSStart As Long, plTSEnd As Long, plTSDiff As Long, plOVRow As Long, plOVCol As Long
Private pclsSimProd As cCBA_Prod, psTotBus As Single

Private psdCOM_Idx As Scripting.Dictionary, psdCont_Idx As Scripting.Dictionary

Private pID_Ary() As cTEN1_Ideas                               ' Contains all the Idea data for the prior year?
Private pSH_Ary() As cTEN3_Templates                           ' Template Hdrs from C1_Seg_Template Hdrs - e.g NDP, Retender and MSO Core Range
Private pSL_Ary() As cTEN4_Segments                            ' Template Lines from C2_Seg_Template_Lines - e.g. all the relavent segment for the above
Private pUTL_Ary() As cTEN8_Lines                              ' UT Lines - Holds a record for each Grouped Tender Document line (normal lines and headings lines are omitted)
Private pUT_Ary1() As cCBA_UDT                                 ' The Master UDT Array
Private pUT_Ary2() As cCBA_UDT                                 ' The Copied UDT Array
Private pATP_Ary() As cTEN7_ATP                                ' Element 0 will be the master - then up to nine number of supplier copies (also applies to any manually added copies)
Private pLV_Ary() As cTEN2_Levels                              ' The L2_Levels class - contains the links between Idea, Portfolio, Version, Product, Contract, Tender Docs and Tender Template + Any Link IDS
Private pIDLV_Ary() As Long                                    ' Provides an array link between the Idea ID and the Levels ID?
Private clsND As cCBA_NielsenData

' TD Vars
Private psdSH As Scripting.Dictionary                           ' Template Hdrs from C1_Seg_Template Hdrs - e.g NDP, Retender and MSO Core Range
Private psdSL As Scripting.Dictionary                           ' Template Lines from C2_Seg_Template_Lines
Private psdUTL As Scripting.Dictionary                          ' UT Lines - Holds a record for each Grouped Tender Document line (normal lines and headings lines are omitted)
Private psdUT_Idx2 As Scripting.Dictionary                      ' Template UDTs subset copied from psdUT_Idx1
Private psdRC_Idx As Scripting.Dictionary                       ' Template UDTs subset Index by Row / Column (format("rrrrcc")) - Points to the pUT_Ary2 Index
Private psdATPD_Idx As Scripting.Dictionary                     ' ATP to ATP Idx
Private psdATPN_Idx As Scripting.Dictionary                     ' ATP Name to ATP Idx
Private psdND As Scripting.Dictionary                           ' Neilsen Data scripting dictionary
Private psdCG As Scripting.Dictionary                           ' Commodity Groups...
Private pclsPFV As CBA_PfVersion                                ' Porfolio Version scripting dictionary
Private pclsProdGrp As cCBA_ProdGroup
Private pary_BDBA()                                             ' BD / BA / GBD Array

' ID Vars
Private psdID As Scripting.Dictionary                           ' Ideas scripting dictionary
Private psdLV As Scripting.Dictionary                           ' Levels scripting dictionary
Private psdBA As Scripting.Dictionary                           ' BA scripting dictionary
Private psdBD As Scripting.Dictionary                           ' BD scripting dictionary
Private psdGBD As Scripting.Dictionary                          ' GBD scripting dictionary
Private psdPFV As Scripting.Dictionary                          ' Portfolio Version scripting dictionary
Private psdIdeaSts As Scripting.Dictionary                      ' Idea Statuses
Private Const COL_WIDTH = 3, NO_OF_COLS = 33, PAGE_MARGIN = 18, ADDIT_DFT = "000000,000000,000000,000000,000000,000000,000000,000000"
Private Const NUM_REGIONS = 8                                   ' If change number of regions, will have to change the a_Regs field too

Private Enum e_EmpRoles     ' Refers to the downloaded employee array of BAs, BDs and GBD and the positions in the array of the following elements
    [_First]
    eEmpNo = 0              ' Employee Number
    eGroupID = 1            ' Employee Number of Boss (except GBD, which will = 0)
    eLast = 2               ' Last Name
    eFirst = 3              ' First Name
    eRole = 4               ' Role - BA, BD or GBD
    [_Last]
End Enum
''Private Enum e_UTFldFmt       ' Refers to number of zeros in xxxxx_Ext fields i.e. UT_ID_Ext is formatted from UT_ID with eUT_ID # of zeros on the end  (UT_ID=2 >>> UT_ID_Ext=20000 (eUT_ID=4))
''    [_First]                ' These values are generated on the fly and are not kept in the database anywhere-As they are liable to exceed Long size, they are kept as strings (Mostly Keys for Dictionaries)
''    eRowCol = 2             ' RowCol_Ext has eRowCol zeros
''    eUT_ID = 4              ' UT_ID_Ext has eUT_ID zeros                (will be further split in Multiple(1st two #) and Group(2nd two #))
''    eUT_Link = 4            ' UT_Link_ID_Ext                            If Link_ID > 0 then will format it as Link_ID and formatted zeros
''    [_Last]
''End Enum
Private Enum e_Addit        ' Refers to 'Number Range' positions in UT_Addit in cCBA_UDT - UT_Addit is generated on the fly in Fill_UT_Data in the format Const ADDIT_DFT
    [_First]                ' These values are captured in the 'Header' or 'Header Button' UDT, for the relating lines, or in the line itself
    eMHgt = 1               ' Kept in the Header Button UT and the grouped lines, the Merge Height of the grouped lines - 'BTN_Add'/'BTN_UnHide'/'BTN_Transfer'
    eRows = 2               ' Kept in the Header Button UT, the Number of valid (have a Dft_Value in UT_Pos_Left=1) Grouped Lines - 'BTN_UnHide'
                            '                               the Number of valid Grouped Lines - 'BTN_Add'
                            '                               the Number of Grouped Lines - 'BTN_Transfer'
    eDDP = 2                ' CostDDP in 'Tender Submission', UT_Pos_Left (Supplier UDT) Idx to 'Tender Submission'=CostDDP
    eDiff = 3               ' Kept in the Header Button UT, the Diff between the UT_Pos_Top of 'BTN_UnHide' & the topmost related grouped line (MIN #0101)
                            '                               the Diff between the UT_Pos_Top of 'BTN_ADD' & the topmost related line, where any of that Mult have a valid Supplier (MIN #xx01)
                            '                               the Diff between the UT_Pos_Top of 'BTN_Transfer' & the topmost related grouped line (MIN #0101)
    eExW = 3                ' CostEx-Works in 'Tender Submission', UT_Pos_Left (Supplier UDT) Idx to 'Tender Submission'=CostEx-Works
    eLastGrpLin = 4         ' Kept in the Header Button UT, this is the line (UT_Pos_Top) number of the last Group (MIN-##01) line - 'BTN_ADD'
                            '                               in 'BTN_Transfer', this is the 'BTN_ADD' Idx number
    eFOB = 4                ' CostFOB in 'Tender Submission', UT_Pos_Left (Supplier UDT) Idx to 'Tender Submission'=CostFOB
''    eTrial = 5              ' CostTrial in 'Tender Submission', UT_Pos_Left (Supplier UDT) Idx to 'Tender Submission'=CostTrial
    eLinIdx = 5             ' In 'Tender Submission', UT_Pos_Left (Supplier UDT) Idx to 'Recommendation' Cost UDT
    eHdrIdx = 6             ' In all corresponding grouped lines, the array index number of the 'Header Button' UDT.
    eFlgIdx = 7             ' In 'Tender Submission', the input Addit array.El(#) of the selected link that will be taken-2=DDP(eDDP); 3=ExW/FG(eExW); 4=FOB(eFOB)....
    [_Last]
End Enum
Private Enum e_Src          ' Refers to character positions in UT_Source in cCBA_UDT
    [_First]
    eInit = 1               ' N=Blank Default_Value upon Copy; D=Allow Default_Value upon Copy
    eLink = 2               ' L=Use Link_ID to link to another UDT (created on the Fly) to get a Default Value, E=Established Link in UT Master
                            ' E=Established link used to link to segments from different Templates
                            ' C=Use Link_ID to link to another UDT - is established but not complete as has Group_No but no Mult_No
    eDsp = 3                ' 0=Do not Display Line(If a Grouped Line,Pos_Left=1,Dft_Val="")
                            ' 0=Do not Display [Hdr] Button(If no valid Grouped Lines (as above))
                            ' A=Always Display (i.e. Save Button)
                            ' H=Hidden - for rows that are always hidden...
    eNumbr = 4              ' '#' Will insert the group number (region) into the UT_Name upon formatting (e.g. CostDDP1 to CostDDP8)
                            ' Later (upon user selection) 'Y' will flag the UT to be transfered from another UT field '0' will null it
    [_Last]
End Enum
Private Const TEN_ERR_EXCEPTS As String = "" '''"White, Robert"   '' "White, Robert;Pearce, Tom" ''' = "White, Robert;Pearce, Tom" ' @RWTen TAKE OUT WHEN COMPLETE
Private pbRW As Boolean

' PLEASE NOTE: THE ABOVE DESCRIPTIONS ARE PROOF POSITIVE, THAT I (ROBERT WHITE) CAN DESCRIBE THINGS WITHOUT THE LIBERAL USE OF TERMS SUCH AS DOUBRIES, WHATSITS, ITS OR THINGIES

Public Sub Generate_TD()
    Dim lIdx As Long
    ' Generate the TD (Tender Doc) Class
    plSH_Idx = -1: plSL_Idx = -1: plUT_Idx1 = -1
    Set psdSH = New Scripting.Dictionary
    Set psdSL = New Scripting.Dictionary
    ReDim pSH_Ary(0 To 0)
    ReDim pSL_Ary(0 To 0)
    ReDim pUT_Ary1(0 To 0)
    If CBA_User = "White, Robert" Then pbRW = True
    Call Get_db_TD_TD
End Sub

Public Sub Generate_ID()
    Dim lIdx As Long
    ' Generate the ID (Ideas) Class
    Set psdIdeaSts = New Scripting.Dictionary
    Set psdBA = New Scripting.Dictionary
    Set psdBD = New Scripting.Dictionary
    Set psdGBD = New Scripting.Dictionary
''    Call Get_db_TD_ID
End Sub

Public Sub TEN_Set_Values(Optional SH_ID As Long = 0, Optional LV_ID As Long = 0, Optional SU_ID As Long = 0, Optional TH_ID As Long = 0, _
                        Optional NPD_R_MSO As String = "", Optional SH_Desc As String = "", Optional AH_ID As Long = 0, Optional ATPLV_ID As Long = 0, Optional ID_ID As Long = 0)
    ' Capture some of the values as they are set into the runtime..
    If SH_ID < 0 Then
        plSH_ID = 0
        plLV_ID = 0
        plSU_ID = 0
        plTH_ID = 0
        plAH_ID = 0
        plID_ID = 0
        psNPD_R_MSO = ""
        psSH_Desc = ""
        Exit Sub
    End If
    If SH_ID > 0 Then plSH_ID = SH_ID
    If LV_ID > 0 Then plLV_ID = LV_ID
    If SU_ID > 0 Then plSU_ID = SU_ID
    If TH_ID > 0 Then plTH_ID = TH_ID
    If AH_ID > 0 Then plAH_ID = AH_ID
    If ID_ID > 0 Then plID_ID = ID_ID
    If ATPLV_ID > 0 Then plATPLV_ID = ATPLV_ID
    If NPD_R_MSO > "" Then psNPD_R_MSO = NPD_R_MSO
    If SH_Desc > "" Then psSH_Desc = SH_Desc
End Sub

Public Function SaveIP_ID(Optional SetGet As String = "Get", Optional bSetAt As Boolean = False) As Boolean
    ' Will deliver back whether the ID is ready to be saved
    Static bValue As Boolean
    If LCase(SetGet) <> "get" Then bValue = bSetAt
    SaveIP_ID = bValue
End Function

Public Function LoadIP_ID(Optional SetGet As String = "Get", Optional bSetAt As Boolean = False) As Boolean
    ' Will deliver back whether there is a load of data IP
    Static bValue As Boolean
    If LCase(SetGet) <> "get" Then bValue = bSetAt
    LoadIP_ID = bValue
End Function

Public Function SaveIP_LV(Optional SetGet As String = "Get", Optional bSetAt As Boolean = False) As Boolean
    ' Will deliver back whether the Levels / Portfolio/s) is/are ready to be saved
    Static bValue As Boolean
    If LCase(SetGet) <> "get" Then bValue = bSetAt
    SaveIP_LV = bValue
End Function

Public Function SaveIP_UT(Optional SetGet As String = "Get", Optional bSetAt As Boolean = False) As Boolean
    ' Will deliver back whether the Tender Document is ready to be saved to the database
    Static bValue As Boolean
    If LCase(SetGet) <> "get" Then bValue = bSetAt
    SaveIP_UT = bValue
End Function

Public Function ShowIP_ID(Optional SetShow_SetNotShow_GetShow As String = "NotShown", Optional bInShowOrNot As Boolean = False) As Boolean
    ' Will deliver back whether the Tender Document is being formatted, Loaded, shown etc
    Static bShown As Boolean, bShowOrNot As Boolean, bLoading As Boolean, sLastVal As String
    Dim sThisVal As String
    If SetShow_SetNotShow_GetShow = "SetNotShow" Then
        bShown = False: bShowOrNot = False '': bLoading = False
        ShowIP_ID = False
    ElseIf SetShow_SetNotShow_GetShow = "SetShow" Then
        bShown = True: bShowOrNot = False ''': bLoading = False
        ShowIP_ID = True
    ElseIf SetShow_SetNotShow_GetShow = "SetShowOrNot" Then
        bShown = True: bShowOrNot = True ''': bLoading = False
        ShowIP_ID = True
    ElseIf SetShow_SetNotShow_GetShow = "SetLoading" Then
        ''bShown = False:
        bLoading = True
    ElseIf SetShow_SetNotShow_GetShow = "SetLoaded" Then
        bLoading = False
        ShowIP_ID = False
        ''ShowIP_ID = True
    ElseIf SetShow_SetNotShow_GetShow = "IsShowing" Then
        If bShowOrNot = True Then
            ShowIP_ID = bInShowOrNot
        Else
            ShowIP_ID = bShown
        End If
    ElseIf SetShow_SetNotShow_GetShow = "IsLoading" Then
        ShowIP_ID = bLoading
    Else
        MsgBox SetShow_SetNotShow_GetShow & " not found in ShowIP_ID"
    End If
    sThisVal = SetShow_SetNotShow_GetShow & " = " & ShowIP_ID
''    If sLastVal <> sThisVal Then Debug.Print sThisVal
    sLastVal = sThisVal
End Function

Public Function Get_db_TD_ID(Optional lBD_Emp_No As Long = 0, Optional lBA_Emp_No As Long = 0, Optional sSearch As String = "", Optional sFromDate As String = "", Optional lSts As Long = 1, Optional lAddedID As Long = 0) As Long
    ' This routine will get the Class Lines for the Idea and Portfolios
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim sSQL As String, lRecs As Long, bAdditSQL As Boolean
    Dim ID_ID As Long, LastID_ID As Long, ID_Desc As String, ID_Sts_ID As Long, Sts_Desc As String, ID_BA As String, ID_BD As String, ID_BA_Emp_No As Long, ID_BD_Emp_No As Long, ID_GD_Emp_No As Long
    Dim LV_ID As Long, LV_Portfolio_No As String, LV_Portfolio_Desc As String, LV_Sts_ID As Long
    Dim LV_Version_No As String, LV_Version_Desc As String
    Dim LV_Product_Code As String, LV_Prior_Product_Code As String, LV_Contract_No As String, LV_TH_Docs As String, LV_ATPFile As String, LV_TH_ID As Long, LV_AH_ID As Long
    Dim LV_UpdDate As String, LV_UpdUser As String, lEmp_No As Long, sLastBA As String, sLastBD As String
    Dim vVar, lID As Long, bOutput As Boolean, lIdx As Long

    On Error GoTo Err_Routine
    Get_db_TD_ID = 0
    CBA_ErrTag = "SQL": lRecs = 0: LastID_ID = -1
    plID_Idx = 0: plLV_Idx = 0: plIDLV_Idx = -1
    Set psdID = New Scripting.Dictionary
    Set psdLV = New Scripting.Dictionary
    ReDim pID_Ary(1 To 1)
    ReDim pLV_Ary(1 To 1)
    ReDim pIDLV_Ary(0 To 2, 0 To 0)
    
    Set CN = New ADODB.Connection
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Ten") & ";"
    
    ' Get the Statuses
    If psdIdeaSts.Exists("1") = False Then
        Set RS = New ADODB.Recordset
        sSQL = "SELECT * FROM C0_Statuses WHERE Sts_IdeaValid='Y'"
        Set RS = New ADODB.Recordset
        RS.Open sSQL, CN
        Do While Not RS.EOF
            ID_Sts_ID = NZ(RS!Sts_ID, 0): Sts_Desc = NZ(RS!Sts_Desc, "unknown")
            psdIdeaSts.Add CStr(ID_Sts_ID), Sts_Desc
            RS.MoveNext
        Loop
    End If
    ' Get the BDs & BAs
    If psdBA.Count = 0 Then
        bOutput = CBA_COM_SQLQueries.CBA_COM_GenPullSQL("CBIS_BDBA_List")
        pary_BDBA = CBA_CBISarr
        Call g_EraseAry(CBA_CBISarr)
        For lIdx = 0 To UBound(pary_BDBA, 2)
            lEmp_No = pary_BDBA(e_EmpRoles.eEmpNo, lIdx)
            Select Case pary_BDBA(e_EmpRoles.eRole, lIdx)
            Case "BD"
                psdBD.Add CStr(CStr(lEmp_No)), lIdx
            Case "BA"
                psdBA.Add CStr(CStr(lEmp_No)), lIdx
            Case Else
                psdGBD.Add CStr(CStr(lEmp_No)), lIdx
            End Select
        Next
    End If
    
    ' Get the Idea records and insert into the class module arrays / dictionaries
    Set RS = New ADODB.Recordset
    sSQL = "SELECT * FROM qry_L2_Ideas WHERE ( " & IIf(lSts = 1, "ID_Sts_ID=1 ", "ID_Sts_ID>0 ")
    ' If there is a Search, BA or BD...
    If lBD_Emp_No > 0 Then sSQL = sSQL & " AND ID_BD_Emp_No=" & lBD_Emp_No: bAdditSQL = True
    If lBA_Emp_No > 0 Then sSQL = sSQL & " AND ID_BA_Emp_No=" & lBA_Emp_No: bAdditSQL = True
    If sSearch > "" Then sSQL = sSQL & " AND ID_DESC LIKE '%" & sSearch & "%'": bAdditSQL = True
    sSQL = sSQL & " AND ID_UpdDate >= " & g_GetSQLDate(sFromDate)
    
    If bAdditSQL = False Then
        sSQL = sSQL & " AND ( ID_UpdUser='" & CBA_User & "' OR ID_CrtUser='" & CBA_User & "' )"
    End If
    sSQL = sSQL & " ) "
    If lAddedID > 0 Then sSQL = sSQL & " OR ( ID_ID=" & lAddedID & " ) "
    ' And order By ID
    sSQL = sSQL & " ORDER BY ID_UpdDate, ID_ID, LV_ID ;"
''    Debug.Print sSQL
    RS.Open sSQL, CN
    Do While Not RS.EOF
        CBA_Error = ""
        If LastID_ID <> NZ(RS!ID_ID, 0) Then
            ID_ID = NZ(RS!ID_ID, 0): ID_Desc = NZ(RS!ID_Desc, ""): ID_Sts_ID = NZ(RS!ID_Sts_ID, 0)
            ID_BA = NZ(RS!ID_BA, ""): ID_BD = NZ(RS!ID_BD, ""): ID_BA_Emp_No = NZ(RS!ID_BA_Emp_No, 0): ID_BD_Emp_No = NZ(RS!ID_BD_Emp_No, 0): ID_GD_Emp_No = NZ(RS!ID_GD_Emp_No, 0)
            Sts_Desc = NZ(RS!Sts_Desc, "")
            Call Add_TD_ID(ID_ID, ID_Desc, ID_Sts_ID, Sts_Desc, ID_BA_Emp_No, ID_BD_Emp_No, ID_GD_Emp_No, RS!ID_UpdDate, RS!ID_UpdUser, RS!ID_CrtUser, "N")
            LastID_ID = ID_ID
        End If
        lRecs = lRecs + 1
        LV_ID = NZ(RS!LV_ID, 0): LV_Sts_ID = NZ(RS!LV_Sts_ID, 0)
        LV_Portfolio_No = NZ(RS!LV_Portfolio_No, ""): LV_Portfolio_Desc = NZ(RS!LV_Portfolio_Desc, ""):
        LV_Version_No = NZ(RS!LV_Version_No, ""): LV_Version_Desc = NZ(RS!LV_Version_Desc, ""):
        LV_Product_Code = NZ(RS!LV_Product_Code, ""): LV_Prior_Product_Code = NZ(RS!LV_Prior_Product_Code, ""): LV_Contract_No = NZ(RS!LV_Contract_No, ""): LV_TH_Docs = NZ(RS!LV_TH_Docs, "")
        LV_ATPFile = NZ(RS!LV_ATPFile, ""): LV_TH_ID = NZ(RS!LV_TH_ID, 0): LV_AH_ID = NZ(RS!LV_AH_ID, 0): LV_UpdUser = NZ(RS!LV_UpdUser, ""): LV_UpdDate = NZ(RS!LV_UpdDate, "")
        ' Add the class to the LV Array
        Call Add_TD_LV(ID_ID, LV_ID, LV_Portfolio_No, LV_Portfolio_Desc, LV_Sts_ID, LV_Version_No, LV_Version_Desc, LV_Product_Code, LV_Prior_Product_Code, _
                        LV_Contract_No, LV_TH_Docs, LV_ATPFile, LV_TH_ID, LV_AH_ID, LV_UpdUser, LV_UpdDate, "N")
''        ' Associate the LV array with the ID
''        Call Set_ID_LV_Idx(ID_ID, LV_ID)
        RS.MoveNext
    Loop
    
Exit_Routine:
    On Error Resume Next
    Call SaveIP_ID("Set", False)
    Call SaveIP_LV("Set", False)
    CN.Close
    Set RS = Nothing
    Set CN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-mTEN_Runtime.Get_db_TD_ID", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function Get_db_TD_TD(Optional bIsCatTool As Boolean = False) As String
    ' This routine will get the Segment Line, UDT Lines and Default ATP Document Templates
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim sSQL As String, lRecs As Long, lsSH_ID As Long, lGrp_No As Long, lSavedGrp_No As Long, lGrpIdx As Long, lIdxIdx As Long
    Dim SH_Desc As String, laryL() As Long, UT_Default_Value As String, sUT_ID As String, sUT_ID_Ext As String '', SH_Sts_ID As Long, Sts_Desc As String
    Dim SL_ID As Long, UT_Link_ID As Long, SL_Seq As Long, SL_Desc As String, SL_SH_ID As String, SL_Mandatory As String, SL_Valid As String, SL_UpdDate As String, SL_UpdUser As String
    Dim UT_Pos_Top As Long, UT_Pos_Left As Long, UT_Merge_Width As Long, lTotal As Long, bValidLine As Boolean, sSL_SysType As String
    Dim vVar, lID As Long, sFld As String
    On Error GoTo Err_Routine
    
    CBA_ErrTag = "SQL": lRecs = 0
    Set CN = New ADODB.Connection
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Ten") & ";"
    ' Set the Load Type
    If bIsCatTool Then
        sSL_SysType = "C"                ' Set up for the Category Tool
    Else
        sSL_SysType = "T"                ' Set up for the TEN Tool
    End If
    ' Get the Template Hdrs
    Set RS = New ADODB.Recordset
    sSQL = "SELECT * FROM C1_Seg_Template_Hdrs WHERE SH_SysType='" & sSL_SysType & "' ORDER BY SH_ID;"
    Set RS = New ADODB.Recordset
    RS.Open sSQL, CN
    Do While Not RS.EOF
        CBA_ErrTag = ""
        lsSH_ID = RS!SH_ID
        SH_Desc = NZ(RS!SH_Desc, 0) '': SH_Sts_ID = NZ(RS!SH_Sts_ID, 0)
        ' Add an Header
        plSH_Idx = plSH_Idx + 1
        If psdSH.Exists(CStr(lsSH_ID)) Then
            MsgBox lsSH_ID & " Tender Hdr already exists"
            Exit Function
        End If
        psdSH.Add CStr(lsSH_ID), plSH_Idx
        ReDim Preserve pSH_Ary(0 To plSH_Idx)
        Set pSH_Ary(plSH_Idx) = New cTEN3_Templates
        Call pSH_Ary(plSH_Idx).Add_Class_SH(plSH_Idx, lsSH_ID, SH_Desc)
        RS.MoveNext
    Loop
   
    ' Get the Segment records and insert into the class module arrays / dictionaries (note: brings in SL_SH_ID=,#, (Template ID) records too. i.e. records that are for all templates)
    Set RS = New ADODB.Recordset
    sSQL = "SELECT * FROM qry_C1_Tenders WHERE SL_Sts_ID=1 AND UT_Sts_ID=1 AND SL_SysType='" & sSL_SysType & "' ORDER BY SL_Seq, UT_Pos_Top, UT_Pos_Left"
    RS.Open sSQL, CN
    Do While Not RS.EOF
        CBA_ErrTag = ""
        If SL_ID <> NZ(RS!SL_ID, 0) Then
            UT_Pos_Top = 0
            SL_ID = NZ(RS!SL_ID, 0): SL_SH_ID = NZ(RS!SL_SH_ID, "")
            SL_Seq = NZ(RS!SL_Seq, 0): SL_Desc = NZ(RS!SL_Desc, ""):
            SL_Mandatory = NZ(RS!SL_Mandatory, "N"): SL_Valid = NZ(RS!SL_Valid, "Y")
            ' Add a Class Line
            plSL_Idx = plSL_Idx + 1
            If SL_ID < 0 Then SL_ID = plSL_Idx
            If psdSL.Exists(CStr(SL_ID)) Then
                MsgBox SL_ID & " Tender Line already exists"
                Exit Function
            End If
            psdSL.Add CStr(SL_ID), plSL_Idx
            ReDim Preserve pSL_Ary(0 To plSL_Idx)
            Set pSL_Ary(plSL_Idx) = New cTEN4_Segments
            Call pSL_Ary(plSL_Idx).Add_Class_SL(plSL_Idx, SL_ID, SL_Seq, SL_Desc, SL_SH_ID, SL_Mandatory, 0, SL_Valid)
        End If
        
''        ' Add the row to the udt array
        sUT_ID = CStr(RS!UT_ID)
        sUT_ID_Ext = g_Fmt_2_IDs(sUT_ID, 0, e_UTFldFmt.eUT_ID)
        If RS!UT_Link_ID > 0 Then UT_Link_ID = g_Fmt_2_IDs(RS!UT_Link_ID, lGrp_No, e_UTFldFmt.eUT_Link) Else UT_Link_ID = "0"
        ' Set up the ID and it's array index to the dictionary
        lRecs = lRecs + 1
        plUT_Idx1 = plUT_Idx1 + 1
''        If psdUT_Idx1.Exists(sUT_ID) Then
''            MsgBox sUT_ID & " UDT already exists"
''            Exit Function
''        End If
''        ' Add the ID and it's array index to the master dictionary
''        psdUT_Idx1.Add sUT_ID, plUT_Idx1
        ReDim Preserve pUT_Ary1(0 To plUT_Idx1)
        Set pUT_Ary1(plUT_Idx1) = New cCBA_UDT
        UT_Default_Value = NZ(RS!UT_Default_Value, "")
        ' Add the UDT to the class
        Call pUT_Ary1(plUT_Idx1).Add_Class_UT(plUT_Idx1, sUT_ID, sUT_ID_Ext, RS!UT_SL_ID, RS!UT_Pos_Left, RS!UT_Pos_Top, RS!UT_Grp_No, RS!UT_Merge_Width, RS!UT_Merge_Height, UT_Link_ID, NZ(RS!UT_Name, ""), _
                            RS!UT_Hdg_Type, UT_Default_Value, RS!UT_Wrap_Text, RS!UT_Font_Size, RS!UT_Locked, RS!UT_BG_Color, RS!UT_FG_Color, RS!UT_Bold, RS!UT_Italic, RS!UT_Underline, RS!UT_StrikeThrough, RS!UT_TextAlign, _
                            RS!UT_Field_Type, NZ(RS!UT_Formula, ""), NZ(RS!UT_Procedure, ""), NZ(RS!UT_Line_Type, ""), NZ(RS!UT_Border, ""), NZ(RS!UT_Cond_Format, ""), NZ(RS!UT_Hyperlink, ""), _
                            NZ(RS!UT_Image, ""), NZ(RS!UT_Source, ""), ADDIT_DFT, RS!UT_Grp_No)
        ' Next record
        RS.MoveNext
    Loop
    
    ' Now get the ATP Document Defaults
    plATP_Idx = -1: lRecs = -1: Set psdATPD_Idx = New Scripting.Dictionary: Set psdATPN_Idx = New Scripting.Dictionary
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    sSQL = "SELECT * FROM C0_Template_Header WHERE TH_Valid='Y' ORDER BY TH_Name DESC;"
    RS.Open sSQL, CN
    ' How many records???
    If Not RS.EOF Then
        RS.MoveLast
        plATP_Idx = RS.RecordCount - 1
        RS.MoveFirst
        ReDim pATP_Ary(0 To plATP_Idx, 0 To 0)
    End If
    ' Get all the ADT Template records
    Do While Not RS.EOF
        lRecs = lRecs + 1
        Set pATP_Ary(lRecs, 0) = New cTEN7_ATP
        ' Add the ATP Template record
        Call pATP_Ary(lRecs, 0).Add_Class_ATP(RS!TH_ID, RS!TH_Alt, RS!TH_Name, "N")
        ' Add the unique psd by Alt
        sFld = RS!TH_Alt
        If psdATPD_Idx.Exists(sFld) Then
            MsgBox sFld & " Template Rec already exists"
            Stop
        End If
        psdATPD_Idx.Add sFld, lRecs
        ' Add the unique psd by Name
        sFld = RS!TH_Name
        If psdATPN_Idx.Exists(sFld) Then
            MsgBox sFld & " Template Rec already exists"
            Stop
        End If
        psdATPN_Idx.Add sFld, lRecs
        ' Next record
        RS.MoveNext
    Loop
    
Exit_Routine:
    On Error Resume Next
    Call SaveIP_UT("Set", False)
    CN.Close
    Set RS = Nothing
    Set CN = Nothing
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-Get_db_TD_TD", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Sub Get_db_TD_TH()
    ' This routine will get the Class L9_TDoc_UDTs - i.e. the saved data for the tender documents
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim sSQL As String, lRecs As Long, lUTIdx As Long, lUTLIdx As Long, sUT_ID_Ext As String, sUT_Pos_Top As String, lTDNo As Long, lLastTDNo As Long, sLastUT_ID_Ext As Long, lLastUT_SL_ID As Long
    Dim UT_Default_Value As String, sUT_UpdA As String, sUT_UpdB As String, lUT_Link_ID As Long, lNoOfCalcedRows As Long, lHdgIdx As Long, lHdgTop As Long, sHdg_Proc As String
    On Error GoTo Err_Routine
    ' Establish the template number
    Select Case psNPD_R_MSO
        Case "NPD"
            lTDNo = 1
        Case "R"
            lTDNo = 2
        Case "MSO"
            lTDNo = 3
        Case Else
            MsgBox psNPD_R_MSO & " Not found in mTEN_Runtime.Get_db_TD_TH"
    End Select
    CBA_ErrTag = "SQL": lRecs = 0
    Set CN = New ADODB.Connection
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Ten") & ";"
    
    ' Get the Tender Document Lines
    Set RS = New ADODB.Recordset
    sSQL = "SELECT * FROM qry_L9_TDoc_UT_UDTs WHERE TU_TH_ID=" & plTH_ID & " AND TU_Template_No <=" & lTDNo & " ORDER BY UT_SL_ID, TU_UT_ID, TU_UT_No,TU_Template_No DESC"
    Set RS = New ADODB.Recordset
    RS.Open sSQL, CN
    CBA_ErrTag = ""
    Do While Not RS.EOF
        sUT_ID_Ext = g_Fmt_2_IDs(RS!TU_UT_ID, RS!TU_UT_No, e_UTFldFmt.eUT_ID)
        ' Flag the segment as being input from the database
        If lLastUT_SL_ID <> RS!UT_SL_ID Then
            lLastUT_SL_ID = RS!UT_SL_ID
            Call Upd_TD_SL(lLastUT_SL_ID, "SL_Template_No", lTDNo, "N")
        End If
''        If RS!TU_UT_ID = 549 Then
''            sUT_ID_Ext = sUT_ID_Ext
''        End If
''        If Val(sUT_ID_Ext) = 549001 Or Val(sUT_ID_Ext) = 549002 Or Val(sUT_ID_Ext) = 549003 Then
''            sUT_ID_Ext = sUT_ID_Ext
''        End If
        ' When the UT_ID changes, only take the first record found (i.e. The order will be in the order (left to right and top to bottom) of the excel worksheet to be written)
        If sLastUT_ID_Ext <> sUT_ID_Ext Then
            lLastTDNo = RS!TU_Template_No
            sLastUT_ID_Ext = sUT_ID_Ext
            ' Update the UDT Class with the L9 Line values
            If psdUT_Idx2.Exists(sUT_ID_Ext) Then
                lUTIdx = psdUT_Idx2.Item(sUT_ID_Ext)
                ' If the record is flagged as hidden or other flag other than "N" then write it to the UTLine Array
                If RS!TU_Sts_Flg <> "N" Then
                    sUT_Pos_Top = Get_TD_UT(lUTIdx, "UT_Pos_Top")
                    If psdUTL.Exists(sUT_Pos_Top) Then
                        lUTLIdx = psdUTL.Item(sUT_Pos_Top)
                        Let pUTL_Ary(lUTLIdx).sTL_Flag = RS!TU_Sts_Flg
                    End If
                End If
                ' Depending on the template requested...
                If lTDNo = RS!TU_Template_No Then
                    Call Upd_TD_UT(lUTIdx, "UT_Default_Value", NZ(RS!TU_Value, ""), "TBL")                         ' If the template is the same as requested, flag the record as available for change
                Else
                    Call Upd_TD_UT(lUTIdx, "UT_Default_Value", NZ(RS!TU_Value, ""), CStr(RS!TU_Template_No))        ' Else Put in the Template No in the update field
                End If
            End If
        End If
        RS.MoveNext
    Loop
    ' Fill any new UDTs (i.e. Ones the are still Adds with any prior data) - e.g there may be a 'NDP Tender Doc' but this is a new 'Retender Doc'. Certain fields are not the same segment # but are copied
    If lTDNo < 2 Then GoTo Exit_Routine                                         ' If it is the first one there is no point in testing prior template lines
    ' Get the 'Linked' Tender Document Lines
    Set RS = New ADODB.Recordset
    sSQL = "SELECT * FROM qry_L9_TDoc_UT_UDTs WHERE TU_TH_ID=" & plTH_ID & " AND TU_Template_No <" & lTDNo & " AND UT_Link_ID>0 ORDER BY UT_SL_ID, TU_UT_ID, TU_UT_No,TU_Template_No DESC"
    Set RS = New ADODB.Recordset
    RS.Open sSQL, CN
    CBA_ErrTag = ""
    Do While Not RS.EOF
        If Mid(NZ(RS!UT_Source, ""), e_Src.eLink, 1) = "E" Then                              ' If the 2nd char is 'E', then this is an established link and is used to link to other Template segments
            sUT_ID_Ext = g_Fmt_2_IDs(RS!UT_Link_ID, RS!TU_UT_No, e_UTFldFmt.eUT_ID)
''            If Val(sUT_ID_Ext) = 53500 Or Val(sUT_ID_Ext) = 73100 Then
''                sUT_ID_Ext = sUT_ID_Ext
''            End If
            ' Update the UDT Class with the L9 Line values
            If psdUT_Idx2.Exists(sUT_ID_Ext) Then
                lUTIdx = psdUT_Idx2.Item(sUT_ID_Ext)
                sUT_UpdA = Get_TD_UT(lUTIdx, "UT_Upd")
                If sUT_UpdA <> "N" And Val(sUT_UpdA) < RS!TU_Template_No Then
                    ' Depending on the template requested...
                    Call Upd_TD_UT(lUTIdx, "UT_Default_Value", NZ(RS!TU_Value, ""), CStr(RS!TU_Template_No))        ' Else Put in the Template No in the Upd field
''                    Call Upd_TD_UT(lUTIdx, "UT_Default_Value", NZ(RS!TU_Value, ""), "NOUPD")       ' Else put in the value but no
                End If
            End If
        End If
        RS.MoveNext
    Loop
        
Exit_Routine:
    On Error Resume Next
    CN.Close
    Set RS = Nothing
    Set CN = Nothing
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-Get_db_TD_TH", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Public Function Get_XLS_File(sPath As String, sFile As String) As Long
    Dim app As New Excel.Application
    Dim wbk As Excel.Workbook, wst As Excel.Worksheet
''    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    ''Dim  bFound As Boolean, bSuppFound As Boolean, sSavedSupp As String, sThisSupp As String, lSavedCol As Long '', bSuppFound As Long
''    Dim db As DAO.Database
    Dim lRow As Long, lcol As Long, lLRec As Long, lRRec As Long, lLCol As Long, lMaxCol As Long, lIdx As Long, sCellVal As String, sda As Variant
    Dim bSuppFound As Boolean, lSuppRow As Long, sField As String, sAltFld As String, lColIdx As Long '', bApplyVal As Boolean            ', sPriorFld As String, lColIdx As Long
    Dim sdCols As Scripting.Dictionary
    Const MaxColsToCheck = 20, MaxSuppRowsToCheck = 20, ALTQ_LEN = 100
    On Error GoTo Err_Routine
    ' Redim the ATP Array - keep the first element as this holds the template values
    ReDim Preserve pATP_Ary(0 To plATP_Idx, 0 To 0): plATP_Cols = 0

    Get_XLS_File = 0
    Application.ScreenUpdating = False
   
''    Application.Cursor = xlWait
    Set sdCols = New Scripting.Dictionary
    Set wbk = app.Workbooks.Open(sPath & sFile)
    Set wst = wbk.Sheets(1)
    ' First of all, go through the first x lines of the xls file and see if the file is valid - also capture the supplier id and it's row position
    bSuppFound = False: lSuppRow = 0: lMaxCol = 2
    lRow = 2: lcol = 1: lColIdx = 0
    For lRow = lRow To MaxSuppRowsToCheck
        If NZ(wst.Cells(lRow, 1), "") = "" Then GoTo SkipToNext1
        If bSuppFound = False And g_KeepReplace(NZ(wst.Cells(lRow, 1), ""), "AlphaN", "") = "Supplier" Then
            lSuppRow = lRow: bSuppFound = True
            For lcol = 2 To MaxColsToCheck
                ' Capture the valid (have supplier value) column numbers
                If NZ(wst.Cells(lRow, lcol), "") > "" Then     ' If a valid supplier column...
                    lMaxCol = lcol: lColIdx = lColIdx + 1
                    ReDim Preserve pATP_Ary(0 To plATP_Idx, 0 To lColIdx)
                    sdCols.Add CStr(lcol), lColIdx
                    ' Add a class instance for each valid clause in the array of valid clauses
                    For lIdx = 0 To plATP_Idx
                        Set pATP_Ary(lIdx, lColIdx) = New cTEN7_ATP
                    Next
                    lIdx = psdATPN_Idx.Item("Supplier")
''                    If InStr(1, pATP_Ary(lIdx, 0).TH_Name, "Trial") > 0 Then
''                        lIdx = lIdx
''                    End If
                    sCellVal = Replace(Replace(wst.Cells(lRow, lcol), "'", "`"), ",", ".")          ' Replace ' and , values in field (',' is Euro decimal point)
                    '                                 Capture 'act col','cell value','original field name' ,'Upd flag'
                    Call pATP_Ary(lIdx, lColIdx).Add_Class_ATP(lcol, sCellVal, pATP_Ary(lIdx, 0).TH_Name, "A")
                End If
            Next
            GoTo SkipToNext1
        ElseIf bSuppFound = False And lRow > 12 Then
            MsgBox "Supplier not found before row " & lRow & " - please check that this is a valid file to load"
            Get_XLS_File = -1
            GoTo Exit_Routine
        End If
''        If lRow <> lSuppRow Then GoTo SkipToNext1
    Next
SkipToNext1:
    ' Go through the rest of the Xls file and pick up valid other than supplier clause/column values
    lRRec = wst.Cells(wst.Rows.Count, "A").End(xlUp).Row
    For lRow = lSuppRow + 1 To lRRec            ' WorksheetFunction.CountA(Columns("A:A"))
        If NZ(wst.Cells(lRow, 1), "") > "" Then
        For lcol = 2 To lMaxCol
            If Not sdCols.Exists(CStr(lcol)) Then GoTo SkiptoNextColLine
                ' Set up the Alt (Key field) (spaces, punc, etc will be removed and made lower case, to attempt to make it short and unique)
                sAltFld = LCase(g_KeepReplace(wst.Cells(lRow, 1), "AlphaN", "", ""))
                If Len(sAltFld) > ALTQ_LEN Then
                    sAltFld = Left(sAltFld, ALTQ_LEN)
                    lRow = lRow
                End If
                ' If the Clause doesn't exist then see if a shorter version does..
                If Not psdATPD_Idx.Exists(sAltFld) And sAltFld Like "deliveredddpunitcostaud*" Then
''                    For Each sda In psdATPD_Idx.Keys
''                        If sda = Left(sAltFld, Len(sda)) Then
                            sAltFld = "deliveredddpunitcostaud"
''                            Exit For
''                        End If
''                    Next
                End If
''                If sAltFld Like "*delivereddd*" Then
''                    sAltFld = sAltFld
''                End If
                ' Check if the ATP Template exists for this clause
                If psdATPD_Idx.Exists(sAltFld) Then
                    lLRec = psdATPD_Idx.Item(sAltFld)
                    lLCol = sdCols.Item(CStr(lcol))
                    Call pATP_Ary(lLRec, lLCol).Add_Class_ATP(lcol, wst.Cells(lRow, lcol), pATP_Ary(lLRec, 0).TH_Name, "A")
                End If
               
SkiptoNextColLine:
            Next
        End If
    Next
SkipRest:
    Get_XLS_File = lColIdx
    plATP_Cols = lColIdx

Exit_Routine:
    On Error Resume Next
    app.Quit
    Set app = Nothing
    Application.ScreenUpdating = True
''    Application.Cursor = xlDefault
    If Get_XLS_File > 0 Then
        Call Upd_TD_LV(plLV_ID, "LV_ATPFile", sFile, "")
    End If
    Exit Function
Err_Routine:
    ' Put a breakpoint here to check any errors
    CBA_Error = Err.Number & Err.Description & "- in Ten_FE-Get_XLS_File"
    Debug.Print CBA_Error
    Stop
''    Call g_FileWrite(g_GetValueFromConfig("GenLiveErrs"), psError)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function Add_TD_ID(ID_ID As Long, ID_Desc As String, ID_Sts_ID As Long, Sts_Desc As String, ID_BA_Emp_No As Long, ID_BD_Emp_No As Long, ID_GD_Emp_No As Long, ID_UpdDate As String, ID_UpdUser As String, ID_CrtUser As String, ByVal ID_Upd As String) As Long
    ' Add an Header
    Dim ID_BA As String, ID_BD As String
    Add_TD_ID = 0
    If ID_ID <= 0 Then ID_ID = plID_Idx + 1
''    Add_TD_ID = True
    If psdID.Exists(CStr(ID_ID)) Then
        Add_TD_ID = 0
        MsgBox ID_ID & " already exists"
        Exit Function
    End If
    Add_TD_ID = ID_ID
    ID_BA = Get_BDBA_Value(ID_BA_Emp_No, "Inits")
    ID_BD = Get_BDBA_Value(ID_BD_Emp_No, "Inits")
    If g_IsDate(ID_UpdDate, True) = False Then ID_UpdDate = g_FixDate(Now, CBA_DMYHN)
    If ID_UpdUser = "" Then ID_UpdUser = CBA_User
    If ID_CrtUser = "" Then ID_CrtUser = CBA_User
    If Sts_Desc = "" Then Sts_Desc = psdIdeaSts.Item(CStr(ID_Sts_ID))
    ' Increment the ID_IDs
    plID_Idx = plID_Idx + 1
    psdID.Add CStr(ID_ID), plID_Idx
    ReDim Preserve pID_Ary(1 To plID_Idx)
    Set pID_Ary(plID_Idx) = New cTEN1_Ideas
    Call pID_Ary(plID_Idx).Add_Class_ID(ID_ID, ID_Desc, ID_Sts_ID, Sts_Desc, ID_BA, ID_BD, ID_BA_Emp_No, ID_BD_Emp_No, ID_GD_Emp_No, ID_UpdDate, ID_UpdUser, ID_CrtUser, ID_Upd)
    
End Function

Public Function Add_TD_LV(ID_ID As Long, LV_ID As Long, LV_Portfolio_No As String, LV_Portfolio_Desc As String, LV_Sts_ID As Long, _
                   LV_Version_No As String, LV_Version_Desc As String, _
                   LV_Product_Code As String, LV_Prior_Product_Code As String, LV_Contract_No As String, LV_TH_Docs As String, _
                   LV_ATPFile As String, LV_TH_ID As Long, LV_AH_ID As Long, LV_UpdUser As String, LV_UpdDate As String, ByVal LV_Upd As String) As Long
    Dim lIdx As Long, lSIdx As Long, bfound As Boolean
    Static lSupp_ID As Long
    ' Add a Class Line
    Add_TD_LV = 0
    plLV_Idx = plLV_Idx + 1
    If LV_ID <= 0 Then LV_ID = plLV_Idx
    If psdLV.Exists(CStr(LV_ID)) Then
        MsgBox LV_ID & " already exists"
        Exit Function
    End If
    psdLV.Add CStr(LV_ID), plLV_Idx
    Add_TD_LV = LV_ID
    If g_IsDate(LV_UpdDate, True) = False Then LV_UpdDate = g_FixDate(Now, CBA_DMYHN)
    If LV_UpdUser = "" Then LV_UpdUser = CBA_User
    ReDim Preserve pLV_Ary(1 To plLV_Idx)
    ' Add the parts of the line to the array
    Set pLV_Ary(plLV_Idx) = New cTEN2_Levels
    Call pLV_Ary(plLV_Idx).Add_Class_LV(ID_ID, LV_ID, LV_Portfolio_No, LV_Portfolio_Desc, LV_Sts_ID, LV_Version_No, LV_Version_Desc, _
                                        LV_Product_Code, LV_Prior_Product_Code, LV_Contract_No, LV_TH_Docs, LV_ATPFile, LV_TH_ID, LV_AH_ID, LV_UpdUser, LV_UpdDate, ByVal LV_Upd)
    ' Associate the LV array with the ID
    Call Set_ID_LV_Idx(ID_ID, LV_ID)
                                        
End Function

Public Sub Upd_TD_SL(ByVal SL_ID As Long, sField As String, ByVal NewValue, ByRef SL_Upd As String)
    Dim lIDPosx As Long
    ' Get the Index to use
    lIDPosx = Get_SL_Idx(SL_ID)
    ' Update the values
    If sField = "SL_Seq" Then
        pSL_Ary(lIDPosx).lSL_Seq = NZ(NewValue, 0)                                        ' Line Foreign key to the TLLine
    ElseIf sField = "SL_Desc" Then
        pSL_Ary(lIDPosx).sSL_Desc = NZ(NewValue, "")
    ElseIf sField = "SL_Mandatory" Then
        pSL_Ary(lIDPosx).sSL_Mandatory = NZ(NewValue, "")
    ElseIf sField = "SL_Template_No" Then
        pSL_Ary(lIDPosx).lSL_Template_No = NZ(NewValue, "")
    ElseIf sField = "SL_Valid" Then
        pSL_Ary(lIDPosx).sSL_Valid = NZ(NewValue, "")
    Else
        MsgBox "Field " & sField & " not found"
        Exit Sub
    End If
''    Call pSL_Ary(lIDPosx).Upd_Class_SL(sField, NewValue)
End Sub

Private Sub Upd_TD_UTRC(lRow As Long, lcol As Long, vValue)
    ''Dim lIDPosx As Long, sFlag As String, lIDPos As Long
    Dim sRowCol As String, sUT_ID_Ext As String, lUT_Idx As Long
    ' Get the Index to use
    sRowCol = g_Fmt_2_IDs(lRow, lcol, e_UTFldFmt.eRowCol)
    If Not psdRC_Idx.Exists(sRowCol) Then                       ' if UDT has an index error
        MsgBox sRowCol & "  Row/Col should exist and doesn't"
        Stop
        GoTo Exit_Routine
    End If
    lUT_Idx = psdRC_Idx.Item(sRowCol)
''    if pbRW = True then Debug.Print sRowCol & " RC upd to " & vValue & " for UT_ID of " & pUT_Ary2(lUT_Idx).Get_Class_UT("UT_ID") & "; ";
    If g_IsNumeric(vValue, False) Then vValue = g_UnFmt(vValue, "dbl")
    Call Upd_TD_UT(lUT_Idx, "UT_Default_Value", vValue, "U")

Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-Upd_TD_UTRC", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Public Sub Upd_TD_UT(ByVal UT_Idx As Long, sField As String, ByVal NewValue, Optional ByVal INIT_TBL_UPD As String = "NOUPD")
    ' Upd Class Header
    Dim lIdx As Long, sUT_Upd As String, bUpd As Boolean        ''', ExtValue, ExtField As String, sNew_Upd As String, sUT_Upd As String '', lIdx1 As Long, sUpd As String, Sts_Desc As String
    lIdx = UT_Idx
    ' Update the value
    If sField = "UT_SL_ID" Then
        Let pUT_Ary2(lIdx).lUT_SL_ID = NZ(NewValue, 0)
    ElseIf sField = "UT_Pos_Left" Then
        Let pUT_Ary2(lIdx).lUT_Pos_Left = NZ(NewValue, 0)
    ElseIf sField = "UT_Pos_Top" Then
        Let pUT_Ary2(lIdx).lUT_Pos_Top = NZ(NewValue, 0)
    ElseIf sField = "UT_Grp_No" Then
        Let pUT_Ary2(lIdx).lUT_Grp_No = NZ(NewValue, 0)
    ElseIf sField = "UT_Merge_Width" Then
        Let pUT_Ary2(lIdx).lUT_Merge_Width = NZ(NewValue, 0)
    ElseIf sField = "UT_Merge_Height" Then
        Let pUT_Ary2(lIdx).lUT_Merge_Height = NZ(NewValue, 0)
    ElseIf sField = "UT_Link_ID" Then
        Let pUT_Ary2(lIdx).lUT_Link_ID = NZ(NewValue, 0)
    ElseIf sField = "UT_Default_Value" Then
        Let pUT_Ary2(lIdx).sUT_Default_Value = NZ(NewValue, "")
    ElseIf sField = "UT_Wrap_Text" Then
        Let pUT_Ary2(lIdx).sUT_Wrap_Text = NZ(NewValue, "")
    ElseIf sField = "UT_Font_Size" Then
        Let pUT_Ary2(lIdx).lUT_Font_Size = NZ(NewValue, 8)
    ElseIf sField = "UT_Locked" Then
        Let pUT_Ary2(lIdx).sUT_Locked = NZ(NewValue, "Y")
    ElseIf sField = "UT_BG_Color" Then
        Let pUT_Ary2(lIdx).lUT_BG_Color = NZ(NewValue, 2)
    ElseIf sField = "UT_FG_Color" Then
        Let pUT_Ary2(lIdx).lUT_FG_Color = NZ(NewValue, 1)
    ElseIf sField = "UT_Bold" Then
        Let pUT_Ary2(lIdx).sUT_Bold = NZ(NewValue, "Y")
    ElseIf sField = "UT_Italic" Then
        Let pUT_Ary2(lIdx).sUT_Italic = NZ(NewValue, "N")
    ElseIf sField = "UT_Underline" Then
        Let pUT_Ary2(lIdx).sUT_Underline = NZ(NewValue, "N")
    ElseIf sField = "UT_StrikeThrough" Then
        Let pUT_Ary2(lIdx).sUT_StrikeThrough = NZ(NewValue, "N")
    ElseIf sField = "UT_TextAlign" Then
        Let pUT_Ary2(lIdx).sUT_TextAlign = NZ(NewValue, "L")
    ElseIf sField = "UT_Field_Type" Then
        Let pUT_Ary2(lIdx).sUT_Field_Type = NZ(NewValue, "txt")
    ElseIf sField = "UT_Formula" Then
        Let pUT_Ary2(lIdx).sUT_Formula = NZ(NewValue, "")
    ElseIf sField = "UT_Procedure" Then
        Let pUT_Ary2(lIdx).sUT_Procedure = NZ(NewValue, "")
    ElseIf sField = "UT_Line_Type" Then
        Let pUT_Ary2(lIdx).sUT_Line_Type = NZ(NewValue, "N")
    ElseIf sField = "UT_Border" Then
        Let pUT_Ary2(lIdx).sUT_Border = NZ(NewValue, "N")
    ElseIf sField = "UT_Name" Then
        Let pUT_Ary2(lIdx).sUT_Name = NZ(NewValue, "")
    ElseIf sField = "UT_Cond_Format" Then
        Let pUT_Ary2(lIdx).sUT_Cond_Format = NZ(NewValue, "")
    ElseIf sField = "UT_Hyperlink" Then
        Let pUT_Ary2(lIdx).sUT_Hyperlink = NZ(NewValue, "")
    ElseIf sField = "UT_Image" Then
        Let pUT_Ary2(lIdx).sUT_Image = NZ(NewValue, "")
        
''    ElseIf sField = "UT_Source" Then
''        Call pUT_Ary2(lIdx).LetsUT_Source(NZ(NewValue, ""))
''    ElseIf sField = "UT_Source" & e_Src.eInit Then
''        Let pUT_Ary2(lIdx).sUT_Source1 = NZ(NewValue, "")
''    ElseIf sField = "UT_Source" & e_Src.eLink Then
''        Let pUT_Ary2(lIdx).sUT_Source2 = NZ(NewValue, "")
''    ElseIf sField = "UT_Source" & e_Src.eDsp Then
''        Let pUT_Ary2(lIdx).sUT_Source3 = NZ(NewValue, "")
''    ElseIf sField = "UT_Source" & e_Src.exxx Then
''        Let pUT_Ary2(lIdx).sUT_Source4 = NZ(NewValue, "")
        
    ElseIf InStr(1, sField, "UT_Addit") > 0 Then
        Call pUT_Ary2(lIdx).LetsUT_Addit(NZ(NewValue, ""), CLng(Val(Right(sField, 1))))
''    ElseIf sField = "UT_Addit" & e_Addit.eMHgt Then
''        Let pUT_Ary2(lIdx).lUT_Addit1 = NZ(NewValue, 0)
''    ElseIf sField = "UT_Addit" & e_Addit.eRows Then
''        Let pUT_Ary2(lIdx).lUT_Addit2 = NZ(NewValue, 0)
''    ElseIf sField = "UT_Addit" & e_Addit.eDiff Then
''        Let pUT_Ary2(lIdx).lUT_Addit3 = NZ(NewValue, 0)
''    ElseIf sField = "UT_Addit" & e_Addit.eLastGrpLin Then
''        Let pUT_Ary2(lIdx).lUT_Addit4 = NZ(NewValue, 0)
''    ElseIf sField = "UT_Addit" & e_Addit.eHdrIdx Then
''        Let pUT_Ary2(lIdx).lUT_Addit5 = NZ(NewValue, 0)
''    ElseIf sField = "UT_Upd" Then                         ' This is not needed as the field will be set by the sUpd Flag as a paramenter
''        Let pUT_Ary2(lIdx).sUT_Upd = NZ(NewValue, "")
''    ElseIf sField = "UT_Sts_ID" Then
''        Let pUT_Ary2(lIdx).lUT_Sts_ID = NZ(NewValue, 1)
    Else
        MsgBox "Field " & sField & " not found"
        Exit Sub
    End If
    bUpd = False
    sUT_Upd = pUT_Ary2(lIdx).sUT_Upd                        ' The Old Upd Flag...
    INIT_TBL_UPD = UCase(INIT_TBL_UPD)
    If INIT_TBL_UPD = "U" Then INIT_TBL_UPD = "UPD"
    If INIT_TBL_UPD = "INIT" Then
        sUT_Upd = "A"
        bUpd = True
    ElseIf INIT_TBL_UPD = "TBL" Then
        sUT_Upd = "N"
        bUpd = True
    ElseIf Val(INIT_TBL_UPD) > 0 Then                       ' Put in the Template Number
        sUT_Upd = INIT_TBL_UPD
        bUpd = True
    ElseIf INIT_TBL_UPD = "UPD" Then
        If sUT_Upd <> "A" And Val(sUT_Upd) = 0 Then sUT_Upd = "U"
        bUpd = True
    ElseIf INIT_TBL_UPD = "NOUPD" Then
    Else
        MsgBox INIT_TBL_UPD & " not found"
    End If

    ' Apply the update Flag
    If bUpd = True Then Let pUT_Ary2(lIdx).sUT_Upd = sUT_Upd
End Sub

Public Sub Upd_TD_ID(ByVal ID_ID As Long, sField As String, ByVal NewValue, ByRef ID_Upd As String)
    ' Upd Class Header
    Dim lIdx As Long, ExtValue, ExtField As String, sNew_Upd As String '', lIdx1 As Long, sUpd As String, Sts_Desc As String
    lIdx = Get_ID_Idx(ID_ID)
    ' Update the value
''    let pID_Ary(lIdx).Upd_Class_ID(sField, NewValue)
    If sField = "ID_ID" Then
        Let pID_Ary(lIdx).lID_ID = NewValue
    ElseIf sField = "ID_Desc" Then
        Let pID_Ary(lIdx).sID_Desc = NewValue
    ElseIf sField = "ID_Sts_ID" Then
        Let pID_Ary(lIdx).lID_Sts_ID = NewValue                                             ' Status for the Idea
    ElseIf sField = "Sts_Desc" Then
        Let pID_Ary(lIdx).sSts_Desc = NewValue                                     ' Status for the Idea
    ElseIf sField = "ID_BA" Then
        Let pID_Ary(lIdx).sID_BA = NewValue
    ElseIf sField = "ID_BD" Then
        Let pID_Ary(lIdx).sID_BD = NewValue
''    ElseIf sField = "ID_UpdDate" Then
''        Let pID_Ary(lIdx).sID_UpdDate = NewValue
''    ElseIf sField = "ID_UpdUser" Then
''        Let pID_Ary(lIdx).sID_UpdUser = NewValue
''    ElseIf sField = "ID_CrtUser" Then
''        Let pID_Ary(lIdx).sID_CrtUser = NewValue
''    ElseIf sField = "ID_Upd" Then
''        Let pID_Ary(lIdx).sID_Upd = NewValue
    ElseIf sField = "ID_BA_Emp_No" Then
        Let pID_Ary(lIdx).lID_BA_Emp_No = NewValue
        ExtValue = Get_BDBA_Value(CLng(NewValue), "Inits")
        Let pID_Ary(lIdx).sID_BA = ExtValue
    ElseIf sField = "ID_BD_Emp_No" Then
        Let pID_Ary(lIdx).lID_BD_Emp_No = NewValue
        ExtValue = Get_BDBA_Value(CLng(NewValue), "Inits")
        Let pID_Ary(lIdx).sID_BD = ExtValue
    ElseIf sField = "ID_GD_Emp_No" Then
        Let pID_Ary(lIdx).lID_GD_Emp_No = NewValue
'        ExtField = "ID_GBD"
'        ExtValue = Get_BDBA_Value(CLng(NewValue), "Inits")
'        let pID_Ary(lIdx).( ExtValue)
    End If
    ' Get the update Flag
    sNew_Upd = p_UpdFlag(pID_Ary(lIdx).sID_Upd, ID_Upd)
    Let pID_Ary(lIdx).sID_Upd = sNew_Upd
    Let pID_Ary(lIdx).sID_UpdUser = CBA_User
    Let pID_Ary(lIdx).sID_UpdDate = Now()
End Sub

Public Sub Upd_TD_LV(ByVal LV_LV As Long, sField As String, ByVal NewValue, ByRef LV_Upd As String)
    ' Upd Levels Class Header
    Dim lIdx As Long, sNew_Upd As String '', lIdx1 As Long, sUpd As String, Sts_Desc As String
    lIdx = Get_LV_Idx(LV_LV)
    ' Update the value
    If sField = "ID_ID" Then
        Let pLV_Ary(lIdx).lID_ID = NZ(NewValue, 0)
    ElseIf sField = "LV_ID" Then
        Let pLV_Ary(lIdx).lLV_ID = NZ(NewValue, 0)
    ElseIf sField = "LV_Portfolio_No" Then
        Let pLV_Ary(lIdx).sLV_Portfolio_No = NZ(NewValue, "")
    ElseIf sField = "LV_Portfolio_Desc" Then
        Let pLV_Ary(lIdx).sLV_Portfolio_Desc = NZ(NewValue, "")
    ElseIf sField = "LV_Version_No" Then
        Let pLV_Ary(lIdx).sLV_Version_No = NZ(NewValue, "")
    ElseIf sField = "LV_Version_Desc" Then
        Let pLV_Ary(lIdx).sLV_Version_Desc = NZ(NewValue, "")
    ElseIf sField = "LV_Product_Code" Then
        Let pLV_Ary(lIdx).sLV_Product_Code = NZ(NewValue, "")
    ElseIf sField = "LV_Prior_Product_Code" Then
        Let pLV_Ary(lIdx).sLV_Prior_Product_Code = NZ(NewValue, "")
    ElseIf sField = "LV_Contract_No" Then
        Let pLV_Ary(lIdx).sLV_Contract_No = NZ(NewValue, "")
    ElseIf sField = "LV_TH_Docs" Then
        Let pLV_Ary(lIdx).sLV_TH_Docs = NZ(NewValue, "")
    ElseIf sField = "LV_Sts_ID" Then
        Let pLV_Ary(lIdx).lLV_Sts_ID = NZ(NewValue, "")
    ElseIf sField = "LV_ATPFile" Then
        Let pLV_Ary(lIdx).sLV_ATPFile = NZ(NewValue, "")
    ElseIf sField = "LV_TH_ID" Then
        Let pLV_Ary(lIdx).lLV_TH_ID = NZ(NewValue, 0)
    ElseIf sField = "LV_AH_ID" Then
        Let pLV_Ary(lIdx).lLV_AH_ID = NZ(NewValue, 0)
''    ElseIf sField = "LV_UpdDate" Then
''        Let pLV_Ary(lIdx).sLV_UpdDate = NZ(NewValue, "")
''    ElseIf sField = "LV_Upd" Then
''        Let pLV_Ary(lIdx).sLV_Upd = NZ(NewValue, "")
''    ElseIf sField = "LV_UpdUser" Then
''        Let pLV_Ary(lIdx).sLV_UpdUser = NZ(NewValue, "")
    Else
        MsgBox "Field " & sField & " not found"
        Exit Sub
    End If

''    Call pLV_Ary(lIdx).Upd_Class_LV(sField, NewValue)
    sNew_Upd = p_UpdFlag(pLV_Ary(lIdx).sLV_Upd, LV_Upd)
    Let pLV_Ary(lIdx).sLV_Upd = sNew_Upd
    Let pLV_Ary(lIdx).sLV_UpdUser = CBA_User
    Let pLV_Ary(lIdx).sLV_UpdDate = Now()
End Sub

Public Function Get_TD_SH(ByVal SH_ID As Long, sField As String)
    ' Get a Header value
    Dim lIdx As Long '', lIdx1 As Long
    If Not psdSH.Exists(CStr(SH_ID)) Then
        MsgBox "Hdr Line " & SH_ID & " doesn't exist"
        Exit Function
    Else
        lIdx = psdSH.Item(CStr(SH_ID))
    End If
    ' Get the value
    If sField = "SH_Desc" Then
        Get_TD_SH = pSH_Ary(lIdx).sSH_Desc
    ElseIf sField = "SH_No" Then
        Get_TD_SH = pSH_Ary(lIdx).lSH_No
    ElseIf sField = "SH_ID" Then
        Get_TD_SH = pSH_Ary(lIdx).lSH_ID
    Else
        MsgBox "Field " & sField & " not found"
    End If
    ''Get_TD_SH = pSH_Ary(lIdx).Get_Class_SH(sField)
End Function

Public Function Get_TD_SL(ByVal SL_ID As Long, sField As String)
    Dim lIdx As Long
    ' Get the Index to use
    lIdx = Get_SL_Idx(SL_ID)
    ' Get the value
    If sField = "SL_ID" Then
        Get_TD_SL = pSL_Ary(lIdx).lSL_ID                                         ' Line ID
    ElseIf sField = "SL_No" Then
        Get_TD_SL = pSL_Ary(lIdx).lSL_No                                         ' Line Array Number
    ElseIf sField = "SL_Seq" Then
        Get_TD_SL = pSL_Ary(lIdx).lSL_Seq                                        ' Line Foreign key to the TLLine
    ElseIf sField = "SL_SH_ID" Then
        Get_TD_SL = pSL_Ary(lIdx).sSL_SH_ID                                      ' Line Foreign key to the PF Status
    ElseIf sField = "SL_Desc" Then
        Get_TD_SL = pSL_Ary(lIdx).sSL_Desc
    ElseIf sField = "SL_Mandatory" Then
        Get_TD_SL = pSL_Ary(lIdx).sSL_Mandatory
    ElseIf sField = "SL_Template_No" Then
        Get_TD_SL = pSL_Ary(lIdx).lSL_Template_No                               ' Usually the same as the DocType Number (1 to 3 for Ten)
    ElseIf sField = "SL_Valid" Then
        Get_TD_SL = pSL_Ary(lIdx).sSL_Valid
    Else
        MsgBox "Field " & sField & " not found"
        Exit Function
    End If
''    Get_TD_SL = pSL_Ary(lIdx).Get_td_SL(sField)
End Function

Public Function Get_TD_UT(ByVal UT_Idx As Long, sField As String) '', Optional bAdd As Boolean = False, Optional bSQL As Boolean = False, Optional sTablePrefix As String = "")
    ' Get the UDT Index
    Dim lIdx As Long, lEndVal As Long
    ' Get the Index to use
    lIdx = UT_Idx
    lEndVal = CLng(Val(Right(sField, 1)))
    ' Get the value
    If sField = "UT_ID" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_ID
    ElseIf sField = "UT_ID_Ext" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_ID_Ext
    ElseIf sField = "UT_No" Then
        Get_TD_UT = pUT_Ary2(lIdx).lUT_No
    ElseIf sField = "UT_SL_ID" Then
        Get_TD_UT = pUT_Ary2(lIdx).lUT_SL_ID
    ElseIf sField = "UT_RC" Then
        Get_TD_UT = g_Fmt_2_IDs(pUT_Ary2(lIdx).lUT_Pos_Top, pUT_Ary2(lIdx).lUT_Pos_Left)
    ElseIf sField = "UT_Pos_Left" Then
        Get_TD_UT = pUT_Ary2(lIdx).lUT_Pos_Left
    ElseIf sField = "UT_Pos_Top" Then
        Get_TD_UT = pUT_Ary2(lIdx).lUT_Pos_Top
    ElseIf sField = "UT_Grp_No" Then
        Get_TD_UT = pUT_Ary2(lIdx).lUT_Grp_No
    ElseIf sField = "UT_Merge_Width" Then
        Get_TD_UT = pUT_Ary2(lIdx).lUT_Merge_Width
    ElseIf sField = "UT_Merge_Height" Then
        Get_TD_UT = pUT_Ary2(lIdx).lUT_Merge_Height
    ElseIf sField = "UT_Link_ID" Then
        Get_TD_UT = pUT_Ary2(lIdx).lUT_Link_ID
    ElseIf sField = "UT_Name" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Name
    ElseIf sField = "UT_Hdg_Type" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Hdg_Type
    ElseIf sField = "UT_Line_Type" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Line_Type
    ElseIf sField = "UT_Name" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Name
    ElseIf sField = "UT_Default_Value" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Default_Value
    ElseIf sField = "UT_Wrap_Text" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Wrap_Text
    ElseIf sField = "UT_Font_Size" Then
        Get_TD_UT = pUT_Ary2(lIdx).lUT_Font_Size
    ElseIf sField = "UT_Locked" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Locked
    ElseIf sField = "UT_BG_Color" Then
        Get_TD_UT = pUT_Ary2(lIdx).lUT_BG_Color
    ElseIf sField = "UT_FG_Color" Then
        Get_TD_UT = pUT_Ary2(lIdx).lUT_FG_Color
    ElseIf sField = "UT_Bold" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Bold
    ElseIf sField = "UT_Italic" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Italic
    ElseIf sField = "UT_Underline" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Underline
    ElseIf sField = "UT_StrikeThrough" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_StrikeThrough
    ElseIf sField = "UT_TextAlign" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_TextAlign
    ElseIf sField = "UT_Field_Type" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Field_Type
    ElseIf sField = "UT_Formula" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Formula
    ElseIf sField = "UT_Procedure" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Procedure
    ElseIf sField = "UT_Line_Type" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Line_Type
    ElseIf sField = "UT_Border" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Border
    ElseIf sField = "UT_Cond_Format" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Cond_Format
    ElseIf sField = "UT_Hyperlink" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Hyperlink
    ElseIf sField = "UT_Image" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Image
        
    ElseIf InStr(1, sField, "UT_Source") > 0 Then
        Get_TD_UT = pUT_Ary2(lIdx).GetsUT_Source(lEndVal)
''    ElseIf sField = "UT_Source" & e_Src.eInit Then
''        Get_TD_UT = pUT_Ary2(lIdx).sUT_Source1
''    ElseIf sField = "UT_Source" & e_Src.eLink Then
''        Get_TD_UT = pUT_Ary2(lIdx).sUT_Source2
''    ElseIf sField = "UT_Source" & e_Src.eDsp Then
''        Get_TD_UT = pUT_Ary2(lIdx).sUT_Source3
''    ElseIf sField = "UT_Source" & e_Src.eNumbr Then
''        Get_TD_UT = pUT_Ary2(lIdx).sUT_Source4
        
    ElseIf InStr(1, sField, "UT_Addit") > 0 Then
        Get_TD_UT = pUT_Ary2(lIdx).GetsUT_Addit(lEndVal)
''    ElseIf sField = "UT_Addit" & e_Addit.eMHgt Then
''        Get_TD_UT = pUT_Ary2(lIdx).lUT_Addit1
''    ElseIf sField = "UT_Addit" & e_Addit.eRows Then
''        Get_TD_UT = pUT_Ary2(lIdx).lUT_Addit2
''    ElseIf sField = "UT_Addit" & e_Addit.eDiff Then
''        Get_TD_UT = pUT_Ary2(lIdx).lUT_Addit3
''    ElseIf sField = "UT_Addit" & e_Addit.eHdrIdx Then
''        Get_TD_UT = pUT_Ary2(lIdx).lUT_Addit5
    ElseIf sField = "UT_Upd" Then
        Get_TD_UT = pUT_Ary2(lIdx).sUT_Upd
    Else
        MsgBox "Field " & sField & " not found"
        Exit Function
    End If
    
End Function

Public Function Get_TD_UTRC(lRow As Long, lcol As Long, sField As String)
    ''Dim lIDPosx As Long, sFlag As String, lIDPos As Long
    Dim sRowCol As String, sUT_ID_Ext As String, lUT_Idx As Long
    ' Get the Index to use
    sRowCol = g_Fmt_2_IDs(lRow, lcol, e_UTFldFmt.eRowCol)
    If Not psdRC_Idx.Exists(sRowCol) Then                       ' if UDT has an index error
        MsgBox sRowCol & "  Row/Col should exist and doesn't in Get_TD_UTRC"
        Stop
        GoTo Exit_Routine
    End If
    lUT_Idx = psdRC_Idx.Item(sRowCol)

    Get_TD_UTRC = Get_TD_UT(lUT_Idx, sField)

Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-Get_TD_UTRC", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Public Function Get_TD_ID(ByVal ID_ID As Long, sField As String, Optional bAdd As Boolean = False, Optional bSQL As Boolean = False, Optional sTablePrefix As String = "")
    ' Get a Ideas Header value
    Dim lIdx As Long
    ' Get the Index to use
    lIdx = Get_ID_Idx(ID_ID)
    ' Get the value
    If sField = "ID_ID" Then
        Get_TD_ID = pID_Ary(lIdx).lID_ID
    ElseIf sField = "ID_Desc" Then
        Get_TD_ID = pID_Ary(lIdx).sID_Desc
    ElseIf sField = "ID_Sts_ID" Then
        Get_TD_ID = pID_Ary(lIdx).lID_Sts_ID                                                    ' Pointer to the Idea Status table
    ElseIf sField = "Sts_Desc" Then
        Get_TD_ID = pID_Ary(lIdx).sSts_Desc
    ElseIf sField = "ID_BA" Then
        Get_TD_ID = pID_Ary(lIdx).sID_BA
    ElseIf sField = "ID_BD" Then
        Get_TD_ID = pID_Ary(lIdx).sID_BD
    ElseIf sField = "ID_BA_Emp_No" Then
        Get_TD_ID = pID_Ary(lIdx).lID_BA_Emp_No
    ElseIf sField = "ID_BD_Emp_No" Then
        Get_TD_ID = pID_Ary(lIdx).lID_BD_Emp_No
    ElseIf sField = "ID_GD_Emp_No" Then
        Get_TD_ID = pID_Ary(lIdx).lID_GD_Emp_No
    ElseIf sField = "ID_UpdDate" Then
        Get_TD_ID = pID_Ary(lIdx).sID_UpdDate
    ElseIf sField = "ID_UpdUser" Then
        Get_TD_ID = pID_Ary(lIdx).sID_UpdUser
    ElseIf sField = "ID_CrtUser" Then
        Get_TD_ID = pID_Ary(lIdx).sID_CrtUser
    ElseIf sField = "ID_Upd" Then
        Get_TD_ID = pID_Ary(lIdx).sID_Upd
    Else
        MsgBox "Field " & sField & " not found in Get_TD_ID"
        Exit Function
    End If
    ' If required for SQL...
    If bSQL Then
        Get_TD_ID = p_SQL_Field(sField, Get_TD_ID, bAdd, bSQL, sTablePrefix)
    End If

End Function

Public Function Get_TD_LV(ByVal LV_ID As Long, sField As String, Optional bAdd As Boolean = False, Optional bSQL As Boolean = False, Optional sTablePrefix As String = "")
    Dim lIdx As Long
    ' Get the Index to use
    lIdx = Get_LV_Idx(LV_ID)
    ' Get the value
    If sField = "LV_ID" Then
        Get_TD_LV = pLV_Ary(lIdx).lLV_ID                                         ' Line Level ID
    ElseIf sField = "ID_ID" Or sField = "LV_ID_ID" Then
        Get_TD_LV = pLV_Ary(lIdx).lID_ID
    ElseIf sField = "LV_Sts_ID" Then
        Get_TD_LV = pLV_Ary(lIdx).lLV_Sts_ID                                     ' Line Foreign key to the PF Status
    ElseIf sField = "LV_Portfolio_No" Then
        Get_TD_LV = pLV_Ary(lIdx).sLV_Portfolio_No
    ElseIf sField = "LV_Portfolio_Desc" Then
        Get_TD_LV = pLV_Ary(lIdx).sLV_Portfolio_Desc
    ElseIf sField = "LV_Version_No" Then
        Get_TD_LV = pLV_Ary(lIdx).sLV_Version_No
    ElseIf sField = "LV_Version_Desc" Then
        Get_TD_LV = pLV_Ary(lIdx).sLV_Version_Desc
    ElseIf sField = "LV_Product_Code" Then
        Get_TD_LV = pLV_Ary(lIdx).sLV_Product_Code
    ElseIf sField = "LV_Prior_Product_Code" Then
        Get_TD_LV = pLV_Ary(lIdx).sLV_Prior_Product_Code
    ElseIf sField = "LV_Contract_No" Then
        Get_TD_LV = pLV_Ary(lIdx).sLV_Contract_No
    ElseIf sField = "LV_TH_Docs" Then
        Get_TD_LV = pLV_Ary(lIdx).sLV_TH_Docs
    ElseIf sField = "LV_ATPFile" Then
        Get_TD_LV = pLV_Ary(lIdx).sLV_ATPFile
    ElseIf sField = "LV_TH_ID" Then
        Get_TD_LV = pLV_Ary(lIdx).lLV_TH_ID
    ElseIf sField = "LV_AH_ID" Then
        Get_TD_LV = pLV_Ary(lIdx).lLV_AH_ID
    ElseIf sField = "LV_UpdDate" Then
        Get_TD_LV = pLV_Ary(lIdx).sLV_UpdDate
    ElseIf sField = "LV_Upd" Then
        Get_TD_LV = pLV_Ary(lIdx).sLV_Upd
    ElseIf sField = "LV_UpdUser" Then
        Get_TD_LV = pLV_Ary(lIdx).sLV_UpdUser
    Else
        MsgBox "Field " & sField & " not found in Get_TD_LV"
        Exit Function
    End If
'''''    Get_TD_LV = pLV_Ary(lIdx).Get_td_LV(sField)
  ' If required for SQL...
    If bSQL Then
        Get_TD_LV = p_SQL_Field(sField, Get_TD_LV, bAdd, bSQL, sTablePrefix)
    End If
End Function

Private Function Get_SL_Idx(ByVal SL_ID As Long) As Long
    ' Will get the line array index number
    If Not psdSL.Exists(CStr(SL_ID)) Then
        Get_SL_Idx = SL_ID                                   ' Will be this one before the records are written
    Else
        Get_SL_Idx = psdSL.Item(CStr(SL_ID))                 ' Will be the SL_ID key after the record is written
    End If
End Function

''Private Function Get_UT_Idx(ByVal UT_ID As Long) As Long
''    ' Will get the UDT array index number
''    If Not psdUT_Idx2.Exists(CStr(UT_ID)) Then
''        Get_UT_Idx = UT_ID                                   ' Will be this one before the records are written
''    Else
''        Get_UT_Idx = psdUT_Idx2.Item(CStr(UT_ID))            ' Will be the UT_ID key after the record is written
''    End If
''End Function

Private Function Get_LV_Idx(ByVal LV_ID As Long) As Long
    ' Will get the line array index number
    If Not psdLV.Exists(CStr(LV_ID)) Then
        Get_LV_Idx = LV_ID                                   ' Will be this one before the records are written
    Else
        Get_LV_Idx = psdLV.Item(CStr(LV_ID))                 ' Will be the LV_ID key after the record is written
    End If
End Function

Private Function Get_ID_Idx(ByVal ID_ID As Long) As Long
    ' Will get the Ideas array index number
    If Not psdID.Exists(CStr(ID_ID)) Then
        Get_ID_Idx = ID_ID                                   ' Will be this one before the records are written
    Else
        Get_ID_Idx = psdID.Item(CStr(ID_ID))                 ' Will be the ID_ID key after the record is written
    End If
End Function

Private Function Get_ID_LV_Idx(ByVal ID_ID As Long, ByVal sDir As String, ByVal sReturn As String) As Long
    ' Will get the Ideas / Levels array index number
    Dim lIdx As Long, lNo As Long
    Static lLastID_ID As Long, lLastNo As Long
    If sDir = "Add" Then
        lNo = -1
        For lIdx = 0 To plIDLV_Idx
            If pIDLV_Ary(0, lIdx) = ID_ID Then lNo = lNo + 1
        Next
        lNo = lNo + 1
        Get_ID_LV_Idx = lNo
    ElseIf lLastID_ID = ID_ID And sDir = "+" Then
        GoTo Get_LV
    ElseIf lLastID_ID <> ID_ID And sDir = "+" Then
        lLastID_ID = ID_ID: lLastNo = -1
        GoTo Get_LV
    End If
    Exit Function
Get_LV:
    For lIdx = 0 To plIDLV_Idx
        If pIDLV_Ary(0, lIdx) = lLastID_ID And pIDLV_Ary(1, lIdx) > lLastNo Then    ' If the next number
            lLastNo = lLastNo + 1
            Get_ID_LV_Idx = pIDLV_Ary(2, lIdx)                                      ' Deliver back the LV_ID
            Exit Function
        End If
    Next
    Get_ID_LV_Idx = -1: lLastID_ID = -1                                             ' If no more LD_IDs, Deliver back -1 to say 'No more in group'
    Exit Function
End Function

Public Function Get_DB_UTData(Optional TORC As String = "C") As Boolean
    Dim lUT_Idx As Long, lSegIdx As Long, iMonth As Integer, iYear As Integer, colCG As Collection
    Dim lProdCode As Long
    Dim lYr As Long, lMth As Long
    Dim v As Variant, va As Variant
    
    On Error GoTo Err_Routine
    CBA_Error = ""
    ' Reset
    If TORC = "T" Then
        ' Set up the Portfolio data
        psPF_No = Get_TD_LV(plLV_ID, "LV_Portfolio_No")
        psPF_Ver_No = Get_TD_LV(plLV_ID, "LV_Version_No")
    End If
    ' Copy the base UT data into another array
    Call Copy_UTData(TORC)
    
    If plTH_ID > 0 Then                         ' If there is previous data then go and get it
        Call Get_db_TD_TH
    Else
        'THIS IS ALL WHERE THE DATA IS COLLECTED IF IT NOT ALREADY IN THE DATABASE FROM A PREVIOUS TENDER DOC ITERATION
        Call CBA_Portfolio_Create.set_PortfolioID(psPF_No)
        Call CBA_Portfolio_Create.set_PfVersionID(NZ(psPF_Ver_No, "0"))
        Set psdPFV = New Scripting.Dictionary
        Set psdPFV = CBA_Portfolio_Create.get_CBA_PortfolioDictionary
        Set pclsPFV = psdPFV(psPF_No)(NZ(psPF_Ver_No, "0"))
        ' Create the Nielsen Data
        iYear = CInt(Year(Date))
        iMonth = CInt(Month(Date)) - 1
        If iMonth < 1 Then iMonth = 12: iYear = CInt(Year(Date)) - 1
    
    ''''''''''''''''''''''''''''''''    iYear = 2019: iMonth = eSep
        Set colCG = New Collection
        colCG.Add Format(pclsPFV.LCGCGno, "000") & Format(pclsPFV.LCGSCGno, "00")
        Set psdND = mCBA_Nielsen.GetNielsenData(iYear, iMonth, False, colCG, False)
        pbPFCreated = True
        On Error Resume Next
        If Not pclsProdGrp Is Nothing Then
            If pclsProdGrp.CompareBuildCode("BcLTRPCm") = True Or pclsProdGrp.CompareBuildCode("BcLTRPCmAs") = True Then
                If CDate(pclsProdGrp.dtMnthFrom) <= CDate(DateSerial(iYear - 1, iMonth + 1, 1)) And CDate(pclsProdGrp.dtMnthTo) >= CDate(DateSerial(iYear, iMonth + 1, 0)) Then
                    For Each v In pclsProdGrp.colCGs
                        If v = Format(pclsPFV.LCGCGno, "000") & Format(pclsPFV.LCGSCGno, "00") Then
                            Set pclsProdGrp = pclsProdGrp
                        Else
                            Set pclsProdGrp = Nothing
                        End If
                        Exit For
                    Next
                End If
            End If
        End If
        
        If pclsProdGrp Is Nothing Then
            Set pclsProdGrp = New cCBA_ProdGroup
            If pclsProdGrp.RunDataGeneration(DateSerial(iYear - 1, iMonth + 1, 1), DateSerial(iYear, iMonth + 1, 0), False, colCG, "BcLTRPCm") = True Then
                psTotBus = 0
                For lYr = Year(pdteFrom) To Year(pdteTo) + 1
                    For lMth = 1 To 12
                        If (lYr = Year(pdteFrom) And lMth >= Month(pdteFrom)) Or (lYr = Year(pdteTo) And lMth <= Month(pdteTo)) Then
                            psTotBus = psTotBus + pclsProdGrp.sdTotBus(CStr(lYr))(CStr(lMth))("POSRetail")
                        End If
                    Next
                Next
                If lProdCode = 0 Then lProdCode = pclsPFV.CompPCode
                If lProdCode = 0 Then lProdCode = Get_TD_LV(plLV_ID, "LV_Prior_Product_Code")
                If lProdCode = 0 Then lProdCode = Get_TD_LV(plLV_ID, "LV_Product_Code")
                If lProdCode <> 0 Then generateComparableProdSpecificObject lProdCode
            Else
    '            lMth = lMth
                'handle Prod Data Not Generating
            End If
        End If
    End If
    ' Do any extra fill of the UT data
    Call Fill_UTData(TORC)

Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-mTEN_Runtime.Get_DB_UTData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function

Private Sub generateComparableProdSpecificObject(ByVal ProdCodeToGenerate As Long)
    Dim tempPG As cCBA_ProdGroup
    Dim iMonth As Integer, iYear As Integer
    Dim colCG As Collection
    
    If pclsProdGrp Is Nothing Or pclsPFV Is Nothing Then Exit Sub
    If pclsProdGrp.getProductState(ProdCodeToGenerate) = True Then
        Set pclsSimProd = pclsProdGrp.getProdObject(ProdCodeToGenerate, True)
        Set pclsSimProd.sdTotalBus = pclsProdGrp.sdTotBus
    Else
        If CBA_BasicFunctions.isRunningSheetDisplayed Then
            CBA_BasicFunctions.CBA_Close_Running
            CBA_BasicFunctions.CBA_Running "Loading Non CG/SCG Similar Comparable Product Data"
        End If
        Set colCG = New Collection
        colCG.Add Format(pclsPFV.LCGCGno, "000") & Format(pclsPFV.LCGSCGno, "00")
        iYear = CInt(Year(Date))
        iMonth = CInt(Month(Date)) - 1
        If iMonth < 1 Then iMonth = 12: iYear = CInt(Year(Date)) - 1
        Set tempPG = New cCBA_ProdGroup
        If tempPG.RunDataGeneration(DateSerial(iYear - 2, iMonth + 1, 1), DateSerial(iYear, iMonth + 1, 0), False, colCG, "BRPC", ProdCodeToGenerate) = True Then
            Set pclsSimProd = tempPG.getProdObject(ProdCodeToGenerate, True)
            Set pclsSimProd.sdTotalBus = pclsProdGrp.sdTotBus
        End If
    End If
End Sub

Public Function Get_BDBA_Value(ByVal lEmp_No As Long, ByVal sField As String)
    ' Will get the BA & BD Values
    Dim lIdx As Long
    On Error GoTo Err_Routine
    CBA_Error = ""
    ' Get the row in the BD Array...
    If psdBA.Exists(CStr(lEmp_No)) Then
        lIdx = psdBA.Item(CStr(lEmp_No))
    ElseIf psdBD.Exists(CStr(lEmp_No)) Then
        lIdx = psdBD.Item(CStr(lEmp_No))
    ElseIf psdGBD.Exists(CStr(lEmp_No)) Then
        lIdx = psdGBD.Item(CStr(lEmp_No))
    Else
        MsgBox "Emp_No " & lEmp_No & " doesn't exist "
        'Stop
        If sField = "Emp_No" Then
            Get_BDBA_Value = 0
        Else
            Get_BDBA_Value = "Unknown"
        End If
        Exit Function
    End If
    ' Get the field value
    If sField = "Emp_No" Then
        Get_BDBA_Value = CLng(pary_BDBA(0, lIdx))
    ElseIf sField = "Grp_Emp_No" Then
        Get_BDBA_Value = CLng(pary_BDBA(1, lIdx))
    ElseIf sField = "Last,First" Then
        Get_BDBA_Value = CStr(pary_BDBA(2, lIdx) & ", " & pary_BDBA(3, lIdx))
    ElseIf sField = "First Last" Then
        Get_BDBA_Value = CStr(pary_BDBA(3, lIdx) & " " & pary_BDBA(2, lIdx))
    ElseIf sField = "Inits" Then
        Get_BDBA_Value = CStr(Left(pary_BDBA(3, lIdx), 1) & Left(pary_BDBA(2, lIdx), 1))
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-mTEN_Runtime.Get_BDBA_Value", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Sub Set_ID_LV_Idx(ByVal ID_ID As Long, ByVal LV_ID As Long)
    ' Will set the Ideas / Levels array index number
    Dim lNo As Long
    lNo = Get_ID_LV_Idx(ID_ID, "Add", "No")
    plIDLV_Idx = plIDLV_Idx + 1
    ReDim Preserve pIDLV_Ary(0 To 2, 0 To plIDLV_Idx)
    pIDLV_Ary(0, plIDLV_Idx) = ID_ID
    pIDLV_Ary(1, plIDLV_Idx) = lNo
    pIDLV_Ary(2, plIDLV_Idx) = LV_ID
End Sub

Public Function p_Flag_Listbox_TD(SL_ID As Long, To_From As String) As Boolean
    ' Will flag the SL_Mandatory of the SL_Desc to appear in whatever listbox
    Dim SL_Mandatory As String
    On Error GoTo Err_Routine
    p_Flag_Listbox_TD = False
    SL_Mandatory = Get_TD_SL(SL_ID, "SL_Mandatory")
    If (SL_Mandatory = "A" Or SL_Mandatory = "N") And To_From = "<" Then
        Call Upd_TD_SL(SL_ID, "SL_Mandatory", "U", "N")
        p_Flag_Listbox_TD = True
    ElseIf (SL_Mandatory = "U" Or SL_Mandatory = "N") And To_From = ">" Then
        Call Upd_TD_SL(SL_ID, "SL_Mandatory", "A", "N")
        p_Flag_Listbox_TD = True
    ElseIf SL_Mandatory = "Y" And To_From = ">" Then
        MsgBox "The Segment you have chosen, can't be moved"
    End If
    On Error Resume Next
''    if pbRW = True then Debug.Print sAccVal & sSep & "0"
    DoEvents
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-mTEN_Runtime.p_Flag_Listbox_TD", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Sub p_FillListBoxes_TD(lstM As Object, lstA As Object, lSH_ID As Long)
    ' Will fill the specified List Box with data from the Template Lines
    Dim lRowM As Long, lRowA As Long, SL_ID As Long, lRec As Long, SL_Mandatory As String, SL_SH_ID As String
    On Error GoTo Err_Routine
    lstM.Clear: lRowM = -1
    lstA.Clear: lRowA = -1
    ' For each row in the array...
    For lRec = 0 To plSL_Idx
        ' Firstly, see if the segment is valid for this Template
        SL_SH_ID = Get_TD_SL(lRec, "SL_SH_ID")
        If InStr(1, SL_SH_ID, "," & lSH_ID & ",") > 0 Then
            ' On the first column for each row, add the item to the list row, and init the fields
            SL_ID = Get_TD_SL(lRec, "SL_ID")
            SL_Mandatory = Get_TD_SL(SL_ID, "SL_Mandatory")
            If SL_Mandatory = "Y" Or SL_Mandatory = "U" Then
                lstM.AddItem SL_ID
                lRowM = lRowM + 1
                lstM.List(lRowM, 1) = SL_Mandatory
                lstM.List(lRowM, 2) = CStr(Get_TD_SL(SL_ID, "SL_Desc"))
            Else
                lstA.AddItem SL_ID
                lRowA = lRowA + 1
                lstA.List(lRowA, 1) = SL_Mandatory
                lstA.List(lRowA, 2) = CStr(Get_TD_SL(SL_ID, "SL_Desc"))
            End If
        End If
    Next
    On Error Resume Next
''    Debug.Print sAccVal & sSep & "0"
    DoEvents
Exit_Routine:

    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-mTEN_Runtime.p_FillListBoxes_TD", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Public Sub p_FillListBox_ID(lst As Object, lCols As Long, bActiveOnly As Boolean, Optional lAdded_ID_ID As Long = 0)
    ' Will fill the Ideas List Box with data
    Dim lRow As Long, ID_ID As Long, savID_ID As Long, lRec As Long, Sts_ID As Long
    On Error GoTo Err_Routine
''    ' Save any existing position
''    If lst.ListIndex > -1 Then
''        savID_ID = lst.Value
''    End If
    lst.Clear: lRow = -1: ID_ID = 1
    ' For each row in the array...
    For lRec = 1 To plID_Idx
        ' On the first column for each row, add the item to the list row, and init the fields
        ID_ID = Get_TD_ID(lRec, "ID_ID")
        Sts_ID = Get_TD_ID(ID_ID, "ID_Sts_ID")
        If (bActiveOnly And Sts_ID = 1) Or bActiveOnly = False Then
            lst.AddItem ID_ID
            lRow = lRow + 1
            lst.List(lRow, 1) = CStr(Get_TD_ID(ID_ID, "ID_Desc"))
            lst.List(lRow, 2) = CStr(Get_TD_ID(ID_ID, "ID_BD"))
            lst.List(lRow, 3) = CStr(Get_TD_ID(ID_ID, "ID_BA"))
            lst.List(lRow, 4) = CStr(Get_TD_ID(ID_ID, "Sts_Desc"))
            lst.List(lRow, 5) = Sts_ID          'CStr(Get_TD_ID(ID_ID, "ID_Sts_ID"))
            lst.List(lRow, 6) = g_FixDate(CStr(Get_TD_ID(ID_ID, "ID_UpdDate")), CBA_D3DMY)
            lst.List(lRow, 7) = CStr(Get_TD_ID(ID_ID, "ID_UpdUser"))
            lst.List(lRow, 8) = CStr(Get_TD_ID(ID_ID, "ID_CrtUser"))
        End If
    Next
    On Error Resume Next
''    Debug.Print sAccVal & sSep & "0"
    DoEvents
Exit_Routine:

    On Error Resume Next
    If lAdded_ID_ID > 0 Then savID_ID = lAdded_ID_ID
    If savID_ID > 0 Then
        lst.Value = Null
        lst.Value = savID_ID
        lst.Value = Null
        lst.Value = savID_ID
    End If
    
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-mTEN_Runtime.p_FillListBox_ID", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Public Sub p_FillListBox_LV(lst As Object, lCols As Long, bActiveOnly As Boolean, lID_ID As Long, Optional ByVal lAdded_LV_ID As Long = 0)
    ' Will fill the PF (Levels) List Box with data from the PF, PV, PD, PC, and LV_ATPFile
    Dim lRow As Long, LV_ID As Long, savLV_ID As Long, lRec As Long, Sts_ID As Long
    On Error GoTo Err_Routine
''    ' Save any existing position
''    If lst.ListIndex > -1 Then
''        savLV_ID = lst.Value
''    End If
    lst.Clear: lRow = -1: LV_ID = 1
    ' For each row in the array...
    Do While LV_ID > 0
        LV_ID = Get_ID_LV_Idx(lID_ID, "+", "LV_ID")
        If LV_ID > 0 Then
            ' On the first column for each row, add the item to the list row, and init the fields
            ''LV_ID = Get_TD_LV(lRec, "LV_ID")
            Sts_ID = Get_TD_LV(LV_ID, "LV_Sts_ID")
            If (bActiveOnly And Sts_ID = 1) Or bActiveOnly = False Then
                lst.AddItem Get_TD_LV(LV_ID, "LV_ID")
                lRow = lRow + 1
                lst.List(lRow, 1) = Sts_ID
                lst.List(lRow, 2) = CStr(Get_TD_LV(LV_ID, "LV_Portfolio_Desc"))
                lst.List(lRow, 3) = CStr(Get_TD_LV(LV_ID, "LV_Version_No") & " " & Get_TD_LV(LV_ID, "LV_Version_Desc"))
                lst.List(lRow, 4) = CStr(Get_TD_LV(LV_ID, "LV_Prior_Product_Code"))
                lst.List(lRow, 5) = CStr(Get_TD_LV(LV_ID, "LV_Product_Code"))
                lst.List(lRow, 6) = CStr(Get_TD_LV(LV_ID, "LV_Contract_No"))
                lst.List(lRow, 7) = CStr(Get_TD_LV(LV_ID, "LV_ATPFile"))
                lst.List(lRow, 8) = CStr(Get_TD_LV(LV_ID, "LV_TH_Docs"))
                lst.List(lRow, 9) = g_FixDate(CStr(Get_TD_LV(LV_ID, "LV_UpdDate")), CBA_DMY) & "    " & CStr(Get_TD_LV(LV_ID, "LV_UpdUser"))
            End If
        End If
    Loop
    On Error Resume Next
    If lAdded_LV_ID > 0 Then savLV_ID = lAdded_LV_ID
    If savLV_ID > 0 Then
        lst.Value = Null
        lst.Value = savLV_ID
        lst.Value = Null
        lst.Value = savLV_ID
    End If
''    Debug.Print sAccVal & sSep & "0"
    DoEvents
Exit_Routine:

    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-mTEN_Runtime.p_FillListBox_LV", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Public Sub p_FillListBox_BDBA(lst1 As Object, lst2 As Object, ByVal sBDEmpNo As String, Optional ByVal sBAEmpNo As String = "", Optional ByVal bAllBAs As Boolean = False)
    ' Will fill the BA & BD List Boxes
    On Error GoTo Err_Routine
    Dim BDBA, BDSaved As String, BASaved As String, lIdx1 As Long, lIdx2 As Long, sLastEmpNo As String
    CBA_Error = ""
    If lst1.ListIndex > -1 Then
        BDSaved = lst1
        If sBDEmpNo = "" Then sBDEmpNo = BDSaved
    ElseIf lst1.ListCount = 0 Then
        BDSaved = "xxxx"
    End If
    If lst2.ListIndex > -1 Then
        BASaved = lst2
        If sBAEmpNo = "" Then sBAEmpNo = BASaved
        ElseIf lst2.ListCount = 0 Then
        BASaved = "xxxx"
    End If
    If BDSaved = "xxxx" Then
        lst1.Clear: lIdx2 = -1
        ' For each row in the BD part of the Array...
        For lIdx1 = 0 To UBound(pary_BDBA, 2)
            If pary_BDBA(e_EmpRoles.eRole, lIdx1) = "BD" Then
                sLastEmpNo = pary_BDBA(e_EmpRoles.eEmpNo, lIdx1)
''                If sLastEmpNo = sBDEmpNo Or sLastEmpNo = sBAEmpNo Then
''                    sLastEmpNo = sLastEmpNo
''                End If
                lIdx2 = lIdx2 + 1
                lst1.AddItem sLastEmpNo
                lst1.Column(1, lIdx2) = pary_BDBA(e_EmpRoles.eLast, lIdx1) & ", " & pary_BDBA(e_EmpRoles.eFirst, lIdx1)
                lst1.Column(2, lIdx2) = Left(pary_BDBA(e_EmpRoles.eLast, lIdx1), 1) & Left(pary_BDBA(e_EmpRoles.eFirst, lIdx1), 1)
            End If
        Next
    End If

    ' Now do the second listbox
    lst2.Clear: lIdx2 = -1
    If sBDEmpNo > "" Or bAllBAs = True Then
        ' For each row in the BA of the Array...
        For lIdx1 = 0 To UBound(pary_BDBA, 2)
            ' BAs according to the BD or bAllBAs=All BAs or BA matches the input one
            If pary_BDBA(e_EmpRoles.eRole, lIdx1) = "BA" And CStr((pary_BDBA(e_EmpRoles.eGroupID, lIdx1)) = sBDEmpNo Or bAllBAs = True Or CStr(pary_BDBA(e_EmpRoles.eEmpNo, lIdx1)) = sBAEmpNo) Then
                sLastEmpNo = pary_BDBA(e_EmpRoles.eEmpNo, lIdx1)
                lIdx2 = lIdx2 + 1
                lst2.AddItem sLastEmpNo
                lst2.Column(1, lIdx2) = pary_BDBA(e_EmpRoles.eLast, lIdx1) & ", " & pary_BDBA(e_EmpRoles.eFirst, lIdx1)
                lst2.Column(2, lIdx2) = Left(pary_BDBA(e_EmpRoles.eLast, lIdx1), 1) & Left(pary_BDBA(e_EmpRoles.eFirst, lIdx1), 1)
            End If
        Next
    End If
    ' For each row in the BD Scripting Array...
    On Error Resume Next
    ' Fill the first listbox - (sometimes the null and reset doesn't work so do it twicw)
    If sBDEmpNo > "" Then
        lst1.Value = sBDEmpNo
        lst1.Value = Null
        lst1.Value = sBDEmpNo
    End If
    If sBAEmpNo > "" Then
        lst2.Value = sBAEmpNo
        lst2.Value = Null
        lst2.Value = sBAEmpNo
    End If
''    Debug.Print sAccVal & sSep & "0"
''    DoEvents
Exit_Routine:

    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-mTEN_Runtime.p_FillListBox_BDBA", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Private Function p_UpdFlag(sExisting_Upd As String, sNew_Upd As String) As String
    ' Set the Update Flag
    p_UpdFlag = sExisting_Upd
    If sNew_Upd = "" Then
    ElseIf sNew_Upd = "A" And sExisting_Upd <> "A" Then
        p_UpdFlag = sNew_Upd
    ElseIf sNew_Upd = "U" And sExisting_Upd = "N" Then
        p_UpdFlag = sNew_Upd
    End If
End Function

Private Function p_SQL_Field(sField As String, vVal, bAdd As Boolean, bSQL As Boolean, Optional sTablePrefix As String = "")
    ' Procees the SQL Values as an Update field ("Field1=[formatted field]) or an Add field ("[formatted field])
    If bSQL Then
        If NZ(vVal, "") = "" Or vVal = "NULL" Then
            vVal = "NULL"
        ElseIf IsNumeric(vVal) Then
        ElseIf g_IsDate(vVal, True) Then
            If g_IsDate(vVal) = False Then vVal = g_FixDate(vVal, CBA_DMYHN)
            vVal = g_GetSQLDate(vVal, CBA_DMYHN)
        Else
            vVal = "'" & Replace(vVal, "'", "`") & "'"
        End If
    End If
    If Not bAdd And bSQL Then
        p_SQL_Field = sTablePrefix & sField & "=" & vVal
    Else
        p_SQL_Field = vVal
    End If
    
End Function


Function Copy_UTData(TORC As String) As Boolean
    ' This routine will just copy all the releventy lines from the master array and sort out the initial positions on the document
    Dim UT_No As Long, sUT_ID As String, UT_SL_ID As Long, UT_Pos_Left As Long, UT_Pos_Top As Long, UT_Grp_No As Long, UT_Merge_Width As Long, UT_Merge_Height As Long, UT_Link_ID As Long
    Dim UT_Name As String, UT_Hdg_Type As String, UT_Default_Value As String, Orig_Default_Value As String, UT_Wrap_Text As String, UT_Upd As String, UT_Addit As String
    Dim UT_Font_Size As Long, UT_Locked As String, UT_BG_Color As Long, UT_FG_Color As Long, UT_Bold As String, UT_Italic As String, UT_Underline As String, UT_StrikeThrough As String, UT_TextAlign As String
    Dim UT_Field_Type As String, UT_Formula As String, UT_Procedure As String, UT_Line_Type As String, UT_Border As String, UT_Cond_Format As String, UT_Hyperlink As String, UT_Image As String, UT_Source As String
    
    Dim lRow As Long, lcol As Long, lSeg As Long, lSegIdx As Long, lUT_Idx As Long, sSL_Mandatory As String, lSL_Template_No As Long '', RCell '', bNewRow As Boolean ''As cell
    Dim lLastSeg As Long, lLastTop As Long, lUdtTop As Long, lUdtBot As Long, lRowTop As Long, lRowBot As Long, lLineTop As Long '', lLeftLeft As Long '', lLastNoI As Long '', lLen As Long
    Dim UT_Mult_No As Long, lMult_NoI As Long, bIncr As Boolean, UT_Orig_Grp_No As Long
''    Dim lChkNo As Long, lDDBNo As Long, lBtnNo As Long
''    Dim lNewIdx As Long, lNoOfCalcedRows As Long, lHdgIdx As Long, lHdgTop As Long, sHdg_Proc As String, bHdg_Upd As Boolean, bCalc0 As Boolean
    Dim bIsGrp As Boolean, bIsMult As Boolean, bGrpD As Boolean, lMult_NoO As Long, lElNo As Long
    Dim lGrpTopIdx As Long, lGrp_NoD As Long, lGrp_NoO As Long, lGrp_NoI As Long, sRowCol As String, sUT_ID_Ext As String, SL_SH_ID As String
    On Error GoTo Err_Routine
    CBA_Error = ""
    ' The following dictionary will hold the links between the incoming UDT table (C2_Seg_Template_Lines), the xls Row and Column and the data table L9_TDoc_UDTs
    ' I.e. The links between the UDT (each individual range defined below) and it's corresponding record in the data table
    ' I.e. xls Row and Column ->>> sUT_ID & Format(UT_Grp_No,"000") (TU_UT_ID & TU_UT_No in the L9_TDoc_UDTs table)
    ' Set up a new scripting Dictionary that will hold the 'UDT array Index' to the 'Row / Column', link
    Set psdRC_Idx = Nothing
    Set psdRC_Idx = New Scripting.Dictionary
    plUT_Idx2 = -1
    Set psdUT_Idx2 = Nothing
    Set psdUT_Idx2 = New Scripting.Dictionary
    Call g_EraseAry(pUT_Ary2)
    ReDim pUT_Ary2(0 To 0)
    ' Restart a new scripting dictionary that will hold all the UT Line data
    plUTL_Idx = -1
    Set psdUTL = Nothing
    Set psdUTL = New Scripting.Dictionary
    Call g_EraseAry(pUTL_Ary)
    ReDim pUTL_Ary(0 To 0)

    bIsGrp = False: lGrpTopIdx = -1: lGrp_NoI = 0: lGrp_NoD = 0: lMult_NoI = 0: lMult_NoO = 0
    ' Cycle through all the UDT Array to copy the relevent lines
    For lUT_Idx = 0 To plUT_Idx1
        lSeg = pUT_Ary1(lUT_Idx).lUT_SL_ID
''        UT_Default_Value = pUT_Ary1(lUT_Idx).sUT_Default_Value
''        sUT_ID = pUT_Ary1(lUT_Idx).sUT_ID
''        If UT_Default_Value = "Similar Core Range Product" Then
''            UT_Default_Value = UT_Default_Value
''        End If
''        If Val(sUT_ID) = 738 Or Val(sUT_ID) = 570 Or Val(sUT_ID) = 561 Then
''            sUT_ID = sUT_ID
''        End If
        
        
        If lLastSeg <> lSeg Then
''        If lLastSeg <> lSeg And lGrp_NoI = 0 Then
            ' On change of segment
            lSegIdx = Get_TD_SL(lSeg, "SL_No")
            sSL_Mandatory = pSL_Ary(lSegIdx).sSL_Mandatory
            lSL_Template_No = pSL_Ary(lSegIdx).lSL_Template_No
            If sSL_Mandatory = "Y" Or sSL_Mandatory = "U" Then
                ' See if the segment is valid for this Template
                SL_SH_ID = Get_TD_SL(lSeg, "SL_SH_ID")
                If InStr(1, SL_SH_ID, "," & plSH_ID & ",") = 0 Then sSL_Mandatory = "N"
                If sSL_Mandatory = "Y" Or sSL_Mandatory = "U" Then                  ' @RWTen Will have to be worked upon for seasonal etc
                    lElNo = IIf(psNPD_R_MSO = "NPD", 1, IIf(psNPD_R_MSO = "R", 2, 0))
                Else
                    lElNo = 0
                End If
            End If
            lLastSeg = lSeg: lLastTop = 0: lLineTop = 0: bGrpD = False: bIncr = True
        ElseIf lLastSeg <> lSeg Then
            bIncr = True
''            lIterIdx = lIterIdx
        End If
        If sSL_Mandatory = "Y" Or sSL_Mandatory = "U" Then   ' sSL_Mandatory = "Y" is a Mandatory Segment --- sSL_Mandatory = "U"  is a User Included Segment
            ' Get all the UDT values
ReDoGrp:
            Call pUT_Ary1(lUT_Idx).Get_Class_UT(UT_No, sUT_ID, sUT_ID_Ext, UT_SL_ID, UT_Pos_Left, UT_Pos_Top, UT_Grp_No, UT_Merge_Width, UT_Merge_Height, UT_Link_ID, UT_Name, UT_Hdg_Type, UT_Default_Value, UT_Wrap_Text, _
                                           UT_Font_Size, UT_Locked, UT_BG_Color, UT_FG_Color, UT_Bold, UT_Italic, UT_Underline, UT_StrikeThrough, UT_TextAlign, UT_Field_Type, UT_Formula, _
                                           UT_Procedure, UT_Line_Type, UT_Border, UT_Cond_Format, UT_Hyperlink, UT_Image, UT_Source, UT_Addit, UT_Orig_Grp_No, UT_Upd)
            If Len(UT_Source) > 4 And lElNo > 0 Then
                If Mid(UT_Source, 4 + lElNo, 1) = "H" Then Mid(UT_Source, e_Src.eDsp, 1) = "H"
            End If
            UT_Mult_No = g_Get_Mid_Fmt(UT_Grp_No, 4, 2) * 100
            UT_Grp_No = g_Get_Mid_Fmt(UT_Grp_No, 2, 2)
            UT_Orig_Grp_No = UT_Grp_No
            ' Sort out the Left-Top cell position
            If lLastTop <> UT_Pos_Top Then
                If UT_Pos_Left = 1 And bGrpD = False Then
                    If bIncr Then
                        lRow = lRowBot + 1
                        If lRow >= 45 And lRow <= 52 Then
                            lRow = lRow
                        End If
                    End If
                    lLineTop = lRow
                    lRowBot = lRow + UT_Merge_Height - 1
                    If bIsGrp = False And bIsMult = False And UT_Grp_No > 0 Then
                        ' Capture the number of CA Rows
                        If UT_Procedure = "CompetitionAnalysis" Then
                            If UT_Grp_No > 0 Then plCARows = UT_Grp_No
                        End If
                        ' If this is a new set of group rows, capture the ID of it, set the upper and lower boundaries and the fact that it is a group
                        lGrp_NoO = UT_Grp_No: bGrpD = False: lGrp_NoD = UT_Grp_No + 1: lGrp_NoI = 0: bIsGrp = True: lGrpTopIdx = lUT_Idx: bIncr = False
                        If UT_Mult_No > 0 Then
                            If UT_Procedure = "TenderSubmissions" Then
                                UT_Mult_No = CBA_MAX_SUPPS ' Make the Mult the same as what is hard coded
                                plTSStart = plUT_Idx2 + 1: plTSEnd = plUT_Idx2 + 1: plTSDiff = -1
                            End If
                            lMult_NoI = 100: bIsMult = True
                        Else
                            lMult_NoI = 0: bIsMult = False
                        End If
                        lMult_NoO = UT_Mult_No * 100
                        GoTo ReDoGrp
                    ElseIf bIsGrp = True And lGrp_NoD > 1 Then
                        lGrp_NoD = lGrp_NoD - 1
                        lLastTop = 0: bGrpD = True: lGrp_NoI = lGrp_NoI + 1: lUT_Idx = lGrpTopIdx: bIncr = False
                        GoTo ReDoGrp
                    ElseIf bIsMult = True And lMult_NoI < lMult_NoO And lGrp_NoD = 1 Then
                        lGrp_NoD = lGrp_NoO + 1: lGrp_NoI = 0: bIsGrp = True: bGrpD = False:
                        lLastTop = 0: lMult_NoI = lMult_NoI + 100: lUT_Idx = lGrpTopIdx: bIncr = False
                        GoTo ReDoGrp
                    Else                'If bGrpD = False Then
                        lGrp_NoD = 0: lGrp_NoI = 0: lMult_NoO = 0: lMult_NoI = 0: bIsGrp = False: bIsMult = False
                    End If
                Else
                    If lGrp_NoI > 0 And bIncr = False Then
                        lRowBot = lRow + UT_Merge_Height - 1
                        If UT_Procedure = "TenderSubmissions" Then
                            If plTSStart < plUT_Idx2 + 1 And plTSDiff = -1 Then plTSDiff = plTSEnd - plTSStart + 1
                        End If
                    End If
                    lLineTop = lRowBot - UT_Merge_Height + 1
                End If
''                If lRow = 20 Then
''                    lUdtTop = lUdtTop
''                End If
                bGrpD = False: lLastTop = UT_Pos_Top: bIncr = True
            End If
            ' Capture the last postion for Tender Submissions
            If UT_Procedure = "TenderSubmissions" Then
                plTSEnd = plUT_Idx2 + 1
            End If

            ' Final Line Top position...
            lUdtTop = lLineTop
           
            sUT_ID_Ext = g_Fmt_2_IDs(Val(sUT_ID), lMult_NoI + lGrp_NoI, e_UTFldFmt.eUT_ID)
            ' Create a UDT Line for every Left_Pos = 1 and Group_No > 0 line
''            If lMult_NoI + lGrp_NoI > 0 And UT_Pos_Left = 1 Then                  ''  @RWTEN THIS IS WHERE THE EXT VERSION ADDS UTL RECORDS*********************
            If ((lMult_NoI + lGrp_NoI) > 0 Or lElNo > 0) And UT_Pos_Left = 1 Then
                ' Add the UT Line ID and it's array index to the UT Lines dictionary
                If Mid(UT_Source, e_Src.eDsp, 1) = "H" Then
                    UT_Default_Value = UT_Default_Value
                End If
                plUTL_Idx = plUTL_Idx + 1
                psdUTL.Add CStr(lUdtTop), plUTL_Idx
                ReDim Preserve pUTL_Ary(0 To plUTL_Idx)
                Set pUTL_Ary(plUTL_Idx) = New cTEN8_Lines
                Call pUTL_Ary(plUTL_Idx).Add_Class_TL(lUdtTop, UT_Merge_Height, "N", "N")
              '  Debug.Print lUdtTop & ";";
            End If
            ' Set the Initial value for the field - i.e. if the Default value has a value but is needed to be initialised to "" upon entry
            If Mid(UT_Source, e_Src.eInit, 1) = "N" Then
                UT_Default_Value = ""
            End If
            ' Will insert the group number (region) into the UT_Name
            If Mid(UT_Source, e_Src.eNumbr, 1) = "#" Then
                UT_Name = UT_Name & CStr(lGrp_NoI)
            End If
            ' If a Group Linked field then incorporate the group no (region number) into the Ext ID
            If Mid(UT_Source, e_Src.eLink, 1) = "L" And UT_Grp_No > 0 Then
                UT_Link_ID = UT_Link_ID + lMult_NoI + lGrp_NoI
            ' ElseIf a Group Linked field but the Mult_No is as yet unknown then incorporate the group no (region number) into the Ext ID
            ElseIf Mid(UT_Source, e_Src.eLink, 1) = "C" And UT_Grp_No > 0 Then
                UT_Link_ID = UT_Link_ID + lGrp_NoI
            End If
            ' Add the ID and it's array index to the dictionary
            plUT_Idx2 = plUT_Idx2 + 1
            psdUT_Idx2.Add CStr(sUT_ID_Ext), plUT_Idx2
            ReDim Preserve pUT_Ary2(0 To plUT_Idx2)
            Set pUT_Ary2(plUT_Idx2) = New cCBA_UDT
''            If UT_Default_Value = "Similar Core Range Product" Then
''                UT_Default_Value = UT_Default_Value
''            End If
            ' Add the UDT to the class
            Call pUT_Ary2(plUT_Idx2).Add_Class_UT(plUT_Idx2, sUT_ID, sUT_ID_Ext, UT_SL_ID, UT_Pos_Left, lUdtTop, lMult_NoI + lGrp_NoI, UT_Merge_Width, UT_Merge_Height, UT_Link_ID, UT_Name, UT_Hdg_Type, UT_Default_Value, _
                               UT_Wrap_Text, UT_Font_Size, UT_Locked, UT_BG_Color, UT_FG_Color, UT_Bold, UT_Italic, UT_Underline, UT_StrikeThrough, UT_TextAlign, _
                               UT_Field_Type, UT_Formula, UT_Procedure, UT_Line_Type, UT_Border, UT_Cond_Format, UT_Hyperlink, UT_Image, UT_Source, UT_Addit, UT_Orig_Grp_No)
            ' Add the new key to the scripting dicts, if it's not a heading
            sRowCol = g_Fmt_2_IDs(lUdtTop, UT_Pos_Left, e_UTFldFmt.eRowCol)
''            UT_Default_Value = sRowCol & "~" & sUT_ID_Ext '''& "~" & UT_Default_Value   '''' #RW ..
            If UT_Hdg_Type <> "Y" Then
                If lGrp_NoD > 999 Then
                    MsgBox " Group Number (number of duplicated lines) should be less than 1000"
                    Stop
                End If
                If Not psdRC_Idx.Exists(sRowCol) Then                      ' UDT may exist as an existing UDT record but not yet as a row/column
                    psdRC_Idx.Add sRowCol, plUT_Idx2
                Else
                    MsgBox sRowCol & " already exists in RowCol"
                    Stop
                End If
            End If
        End If
ReIter:
    Next
  '  Debug.Print
Exit_Routine:

    On Error Resume Next
    Exit Function
    
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-mTEN_Runtime.Copy_UTData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function

Function Fill_UTData(TORC As String) As Boolean
    ' This routine will do the actual filling of the TEN data, formatting number of group lines etc
    On Error GoTo Err_Routine
    CBA_Error = ""

    Dim UT_No As Long, sUT_ID As String, UT_SL_ID As Long, UT_Pos_Left As Long, UT_Pos_Top As Long, UT_Grp_No As Long, UT_Merge_Width As Long, UT_Merge_Height As Long, UT_Link_ID As Long
    Dim UT_Name As String, UT_Hdg_Type As String, UT_Default_Value As String, Orig_Default_Value As String, UT_Wrap_Text As String, UT_Upd As String, UT_Addit As String, UT_Orig_Grp_No As Long
    Dim UT_Font_Size As Long, UT_Locked As String, UT_BG_Color As Long, UT_FG_Color As Long, UT_Bold As String, UT_Italic As String, UT_Underline As String, UT_StrikeThrough As String, UT_TextAlign As String
    Dim UT_Field_Type As String, UT_Formula As String, UT_Procedure As String, UT_Line_Type As String, UT_Border As String, UT_Cond_Format As String, UT_Hyperlink As String, UT_Image As String, UT_Source As String
    
    Dim lNoOfAddedRows As Long, lHdgIdx As Long, lHdgTop As Long, lRowTop As Long, lRowTop1 As Long, sHdg_Proc As String, bHdg_Upd As Boolean, bWriteFrstRow As Boolean, lTS_Idx As Long '', aryCHdgs(1 To 4, 1 To 2) As Long
    Dim bCalc0 As Boolean, bCalc1 As Boolean, sUT_ID_Ext As String, Get_Left_Pos1_Dft_Val As String, lUT_Idx As Long, l1stInVldGrpRow As Long, lLastTSIdx As Long
    Dim sName As String, lGrp_No As Long, sLink_ID_Ext As String, lLinkIdx As Long, aFla() As String, l1stDiff As Long, bSelUpd As Boolean '', bLin_Upd As Boolean

    ' Cycle through all the UDT Array to copy the relevent ones
    For lUT_Idx = 0 To plUT_Idx2
            ' Get all the UDT values
ReDoGrp:
        Call pUT_Ary2(lUT_Idx).Get_Class_UT(UT_No, sUT_ID, sUT_ID_Ext, UT_SL_ID, UT_Pos_Left, UT_Pos_Top, UT_Grp_No, UT_Merge_Width, UT_Merge_Height, UT_Link_ID, UT_Name, UT_Hdg_Type, UT_Default_Value, UT_Wrap_Text, _
                                       UT_Font_Size, UT_Locked, UT_BG_Color, UT_FG_Color, UT_Bold, UT_Italic, UT_Underline, UT_StrikeThrough, UT_TextAlign, UT_Field_Type, UT_Formula, _
                                       UT_Procedure, UT_Line_Type, UT_Border, UT_Cond_Format, UT_Hyperlink, UT_Image, UT_Source, UT_Addit, UT_Orig_Grp_No, UT_Upd)
        ' Save the UT_Default_Value
        Orig_Default_Value = UT_Default_Value
        If Mid(UT_Source, e_Src.eDsp, 1) = "H" And UT_Pos_Left = 1 Then
            UT_Procedure = UT_Procedure
        End If
        If Len(UT_Field_Type) >= 3 Then
            Select Case Left(UT_Field_Type, 3)
                Case "txa"
                    UT_Default_Value = "'" & UT_Default_Value
                Case "now"
                    If plTH_ID = 0 Then UT_Default_Value = Format(Date, CBA_DMY)
            End Select
        End If

        ' Reset the Default Value if it is affected by the Procedure ...
        bSelUpd = False
        If UT_Upd = "A" Then bSelUpd = True
        If UT_Procedure = "Recommendation" And UT_Name <> "Supplier" And g_Val(UT_Default_Value) = 0 Then
            bSelUpd = True
        End If
        If UT_Procedure = "TenderFor" And (UT_Name = "ProdCode" Or UT_Name = "IdeaDesc") Then
            bSelUpd = True
        End If
        If UT_Procedure > "" And bSelUpd = True And UT_Hdg_Type <> "Y" Then
            UT_Default_Value = Set_TD_TD_Value(UT_Procedure, UT_Name, UT_Source, UT_Default_Value, UT_Pos_Top, UT_Pos_Left, UT_Grp_No)
        End If
        ' Capture the col and row for the list of regions selected in recommendation
        If UT_Procedure = "Prod&MktOV" And UT_Name = "RegionCap" Then
            plOVRow = UT_Pos_Top: plOVCol = UT_Pos_Left
        End If
        ' Only the UTs that are buttons
        If UT_Pos_Left = 1 Or UT_Pos_Left > NO_OF_COLS Then
            ' If it is a Hdg Row and there is data in the heading that is applicable then save it...
            If UT_Pos_Left > NO_OF_COLS And UT_Procedure > "" And UT_Hdg_Type = "Y" Then
                sHdg_Proc = UT_Procedure: lHdgIdx = lUT_Idx: lHdgTop = UT_Pos_Top: bCalc0 = True: bCalc1 = True: l1stDiff = 0
                Call Upd_TD_UT(lHdgIdx, "UT_Addit" & e_Addit.eHdrIdx, lHdgIdx)          ' Write the Header Index to the Button
                lNoOfAddedRows = NUM_REGIONS
                If UT_Procedure = "CompetitionAnalysis" Then lNoOfAddedRows = plCARows
            ' Else if the same Proc and the left =1 and not a heading
            ElseIf UT_Pos_Left = 1 And UT_Procedure = sHdg_Proc And UT_Procedure > "" And UT_Hdg_Type <> "Y" Then
                Get_Left_Pos1_Dft_Val = UT_Default_Value
                ' Write to the Line UDT
                Call Upd_TD_UT(lUT_Idx, "UT_Addit" & e_Addit.eMHgt, UT_Merge_Height)            ' Write the Merge Height to every 1st Left Pos line
                Call Upd_TD_UT(lUT_Idx, "UT_Addit" & e_Addit.eHdrIdx, lHdgIdx)                  ' Write the Header Index to every 1st Left Pos line
                ' On the first line of the group, capture the line Height and the number of lines in the group
                If g_Get_Mid_Fmt(UT_Grp_No, 2, 2) = 1 And bCalc1 = True Then
                    Call Upd_TD_UT(lHdgIdx, "UT_Addit" & e_Addit.eMHgt, UT_Merge_Height)        ' Write the Merge Height to the Header Line
                    Call Upd_TD_UT(lHdgIdx, "UT_Addit" & e_Addit.eRows, lNoOfAddedRows)
                    bCalc1 = False
                End If
                If UT_Procedure = "TenderSubmissions" Then              ' For Tender Submissions only  ***************Add Button/TenderSubmissions"*****************
                    ' Calc the New Top Diff - (Should go from 1 to Num_REGIONS)
                    If g_Get_Mid_Fmt(UT_Grp_No, 2, 2) = 1 Then
                        lRowTop = UT_Pos_Top - lHdgTop
                        bWriteFrstRow = True
                        If l1stDiff = 0 Then l1stDiff = lRowTop
                        If g_Get_Mid_Fmt(UT_Grp_No, 4, 2) = 1 Then
                            lRowTop1 = NUM_REGIONS * (CBA_MAX_SUPPS - 1) * UT_Merge_Height
                            Call Upd_TD_UT(lHdgIdx, "UT_Addit" & e_Addit.eLastGrpLin, lRowTop1 + l1stDiff) ' Write the Row Diff between the last Group UDT Line line to the Header Row
                        End If
                    End If
                    ' On the left = 1, regardless of region #, capture the total difference number between the Add button Row and the top grouped line - Put it in the header button UT
                    If bWriteFrstRow = True Then
                        If l1stInVldGrpRow = 0 And (UT_Default_Value = "" Or g_Get_Mid_Fmt(UT_Grp_No, 4, 2) = CBA_MAX_SUPPS) Then
                            l1stInVldGrpRow = lRowTop                                               ' lRowTop captures the 1st line of the group
                            lRowTop1 = NUM_REGIONS * (Val(g_Get_Mid_Fmt(UT_Grp_No, 4, 2)) - 1) * UT_Merge_Height
                            Call Upd_TD_UT(lHdgIdx, "UT_Addit" & e_Addit.eDiff, lRowTop1 + l1stDiff)    ' Write the Row Diff between the first hidden line to the Top Header Row
                            bWriteFrstRow = False                                                   ' Flag as being taken
                        End If
                    End If
                    lLastTSIdx = lHdgIdx                                                            ' Capture the TS Heading index for use in the Recommendation section
                    lTS_Idx = lUT_Idx                                                               ' Capture the TS Line index for use in the Linking of lines to Costs
                ElseIf UT_Procedure = "Recommendation" Then              ' For Recommendations only  ***************btnTransfer / Recommendation"*****************
                    ' On the first line of the group, capture the line Height and the difference between the Add button Row and the first UDT line - Put in in the header button UT
                    If g_Get_Mid_Fmt(UT_Grp_No, 2, 2) = 1 Then
                        lRowTop = UT_Pos_Top - lHdgTop
                        Call Upd_TD_UT(lHdgIdx, "UT_Addit" & e_Addit.eMHgt, UT_Merge_Height)        ' Write the Merge Height to the Header Line
                        Call Upd_TD_UT(lHdgIdx, "UT_Addit" & e_Addit.eDiff, lRowTop)                ' Write the Row Diff between the 1st line to the to the Header Line
                        Call Upd_TD_UT(lHdgIdx, "UT_Addit" & e_Addit.eLastGrpLin, lLastTSIdx)       ' Write the Add Button HdrIndex so that it can be accessed
                    End If
              End If
            ' ElseIf it is the corresponding Group Row and the same procedure name, then recover it and write it back to the heading UDT ...
            ElseIf UT_Pos_Left > NO_OF_COLS And UT_Procedure = sHdg_Proc And UT_Procedure > "" And UT_Hdg_Type <> "Y" Then
                bHdg_Upd = True                                     ' Assume its going to update the hdg dets   ***************UnHide Button/CompetitionAnalysis"*****************
''                If bCalc0 = True Then    ' @RWTen This is where you may alter the display of lines
''                    If Get_Left_Pos1_Dft_Val > "" Then
''''                        lNoOfAddedRows = UT_Grp_No                                             ' Calculate the number of lines concerned
''                    Else
''                        bHdg_Upd = False
''                    End If
''                End If
                ' On the first line of the group, capture the line Height and the difference between the Header button Row and the first line
                If g_Get_Mid_Fmt(UT_Grp_No, 2, 2) = 1 Then
                    lHdgTop = UT_Pos_Top - lHdgTop
                    Call Upd_TD_UT(lHdgIdx, "UT_Addit" & e_Addit.eDiff, lHdgTop)            ' Write the Row Diff between the 1st line to the to the Header Line
                End If
                ' Update the hdg details for use when the Unhide button is pressed
                If bHdg_Upd Then
                    Call Upd_TD_UT(lUT_Idx, "UT_Addit" & e_Addit.eMHgt, UT_Merge_Height)    ' Write the Merge Height to the 1st Left Pos line
                End If
            End If
        ElseIf UT_Pos_Left > 1 And UT_Pos_Left < NO_OF_COLS And UT_Procedure = "TenderSubmissions" And UT_Hdg_Type <> "Y" Then
            sName = g_Left(UT_Name, 1): lGrp_No = Val(Right(UT_Name, 1)): sLink_ID_Ext = g_Fmt_2_IDs(g_Get_Mid_Fmt(UT_Link_ID, 9, 5), lGrp_No, e_UTFldFmt.eUT_Link)
            If sName = "CostDDP" Then
                Call Upd_TD_UT(lTS_Idx, "UT_Addit" & e_Addit.eDDP, lUT_Idx)                         ' Write the CostDDP Linked ID to the 1st Left Pos line
                If psdUT_Idx2.Exists(sLink_ID_Ext) Then
                    lLinkIdx = psdUT_Idx2.Item(sLink_ID_Ext)
                    Call Upd_TD_UT(lTS_Idx, "UT_Addit" & e_Addit.eLinIdx, lLinkIdx)                 ' Write the Line index of the UT_ID that will be updated
                Else
                    MsgBox sLink_ID_Ext & " doesn't exist in Fill_UT_Data"
                    GoTo ReIter
                End If
            ElseIf sName = "CostEx-Works" Then
                Call Upd_TD_UT(lTS_Idx, "UT_Addit" & e_Addit.eExW, lUT_Idx)                         ' Write the CostEx-W Linked ID to the 1st Left Pos line
            ElseIf sName = "CostFOB" Then
                Call Upd_TD_UT(lTS_Idx, "UT_Addit" & e_Addit.eFOB, lUT_Idx)                         ' Write the CostFOB Linked ID to the 1st Left Pos line
            End If
        End If
        ' If the value has changed then update it in the respective UDT class
        If Orig_Default_Value <> UT_Default_Value Then
            Call Upd_TD_UT(lUT_Idx, "UT_Default_Value", UT_Default_Value, "Upd")
        End If
ReIter:
    Next
    
Exit_Routine:
    On Error Resume Next
    Exit Function
    
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-mTEN_Runtime.Fill_UTData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function

Public Sub CrtTenderDoc(Optional ByVal SH_ID As Long = 0, Optional SH_Desc As String = "", Optional bGetDB As Boolean = True, Optional TORC As String = "C")
''    Dim Data() As Variant
''    Dim RP As CBA_BTF_ReportParamaters
    Dim rng As Range, lSavedIdx As Long, lFreezePos As Long, sFla As String '', wks(1 To 5) As Worksheet
    
    Dim UT_No As Long, sUT_ID As String, UT_SL_ID As Long, UT_Pos_Left As Long, UT_Pos_Top As Long, UT_Grp_No As Long, UT_Merge_Width As Long, UT_Merge_Height As Long, UT_Link_ID As Long
    Dim UT_Name As String, UT_Hdg_Type As String, UT_Default_Value As String, UT_Wrap_Text As String, UT_Upd As String, UT_Addit As String, UT_Orig_Grp_No As Long
    Dim UT_Font_Size As Long, UT_Locked As String, UT_BG_Color As Long, UT_FG_Color As Long, UT_Bold As String, UT_Italic As String, UT_Underline As String, UT_StrikeThrough As String, UT_TextAlign As String
    Dim UT_Field_Type As String, UT_Formula As String, UT_Procedure As String, UT_Line_Type As String, UT_Border As String, UT_Cond_Format As String, UT_Hyperlink As String, UT_Image As String, UT_Source As String
    Dim bfound As Boolean, WB As Workbook
''    Dim lRow As Long, lCol As Long, lSeg As Long, lSegIdx As Long, lUT_Idx As Long, sSL_Mandatory As String ''', lSL_Template_No As Long ''As cell
''    Dim lUdtBot As Long, lLen As Long '', ''lLastLeft As Long, lLeftLeft As Long, lLen As Long
    
    Dim sUT_ID_Ext As String, lChkNo As Long, lDDBNo As Long, lBtnNo As Long, lUT_Idx As Long, lUTL_Idx As Long, lUdtBot As Long, lLen As Long, aFla() As String, lGrp_No As Long, lOrigGrp_No As Long, lIdx As Long
    Dim aGrp() As Variant, sdGrp As Scripting.Dictionary, lGrp As Long, sStsFlag As String, lUTLIdx As Long, bDspBtn As Boolean, lRow As Long, lMergeHeight As Long, sRngName As String, rRng As Range, aHdg() As String
''    Static sLastPF_No As String, sLastPF_Ver_No As String
    
    On Error GoTo Err_Routine
    Call ShowIP_ID("SetNotShow")
    Call ShowIP_ID("SetLoading")
    If SH_ID > 0 Then plSH_ID = SH_ID
    If SH_Desc > "" Then psSH_Desc = SH_Desc
    psLCG = "": psLSCG = "": psACG = "": psASCG = ""
    CBA_ErrTag = ""
    Call g_SetupIP("TenderDocs", 1, True)
    ' Capture the original WorkBook name
    If ActiveWorkbook Is Nothing Then Application.Workbooks.Add
    If psActWBkName <> "" Then
        bfound = False
        For Each WB In Application.Workbooks
            If WB.Name = psActWBkName Then
                bfound = True
                Exit For
            End If
        Next
        If bfound = False Then
            psActWBkName = ActiveWorkbook.Name
        End If
    Else
        psActWBkName = ActiveWorkbook.Name
    End If
    Set psWShtLast = ActiveSheet
    ' Copy across new worksheet
    CBA_TEN_WBK.Copy After:=Workbooks(psActWBkName).Sheets(Workbooks(psActWBkName).Sheets.Count)                '''Before:=Workbooks(pwb WBk.Name).Sheets(1)
    Set pwbWSht = ActiveSheet
    plWShtNo = plWShtNo + 1
    If TORC = "T" Then
        psWShtName = "TenderDoc" & plWShtNo
    Else
        psWShtName = "CATREV" & plWShtNo
    End If
    pwbWSht.Name = psWShtName
    pbPFCreated = False
    pdteFrom = DateSerial(Year(Date) - 1, Month(Date), 1)
    pdteTo = DateSerial(Year(Date), Month(Date), 0)
    ReDim aGrp(0 To 4, 0 To 0): Set sdGrp = New Scripting.Dictionary: lGrp = -1
        
    CBA_BasicFunctions.CBA_Running "Loading Data " & SH_Desc
''    If cba_user <> "White, Robert" Then
    Application.ScreenUpdating = False
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Preparing Report for " & IIf(psNPD_R_MSO = "R", "Retender", psNPD_R_MSO)
    ' The following dictionary will hold the links between the incoming UDT table (C2_Seg_Template_Lines), the xls Row and Column and the data table L9_TDoc_UDTs
    ' I.e. The links between the UDT (each individual range defined below) and it's corresponding record in the data table
    ' I.e. xls Row and Column ->>> sUT_ID & Format(UT_Grp_No,"000") (TU_UT_ID & TU_UT_No in the L9_TDoc_UDTs table)
   
    ' Get any data if needed
    If bGetDB Then Call Get_DB_UTData(TORC)
    
    ' And in the workbook / worksheet
    With pwbWSht
         .Activate
''         ' Delete any named ranges created...
''         Call g_DelNameRngs("T10")
         ' Set the base column widths
         Range(.Cells(, 1), .Cells(, 33)).ColumnWidth = COL_WIDTH
         For lUT_Idx = 0 To plUT_Idx2
            ' Get all the UDT values
''            If Val(sUT_ID) >= 480 And Val(sUT_ID) <= 488 Then
''                sUT_ID = sUT_ID
''            End If
            Call pUT_Ary2(lUT_Idx).Get_Class_UT(UT_No, sUT_ID, sUT_ID_Ext, UT_SL_ID, UT_Pos_Left, UT_Pos_Top, UT_Grp_No, UT_Merge_Width, UT_Merge_Height, UT_Link_ID, UT_Name, UT_Hdg_Type, UT_Default_Value, _
                                            UT_Wrap_Text, UT_Font_Size, UT_Locked, UT_BG_Color, UT_FG_Color, UT_Bold, UT_Italic, UT_Underline, UT_StrikeThrough, UT_TextAlign, UT_Field_Type, UT_Formula, _
                                            UT_Procedure, UT_Line_Type, UT_Border, UT_Cond_Format, UT_Hyperlink, UT_Image, UT_Source, UT_Addit, UT_Orig_Grp_No, UT_Upd)
            ' Save the UT_Default_Value
''            Orig_Default_Value = UT_Default_Value
            lUdtBot = UT_Pos_Top + UT_Merge_Height - 1
            ' Set the Freeze Position
            If lFreezePos = 0 Then lFreezePos = lUdtBot
            lLen = UT_Pos_Left + UT_Merge_Width - 1
''            If Mid(UT_Source, e_Src.eDsp, 1) = "H" And UT_Pos_Left = 1 Then
''                UT_Procedure = UT_Procedure
''            End If

''            If UT_Default_Value = "Similar Core Range Product" Then
''                UT_Default_Value = UT_Default_Value
''            End If
            
''            If Val(sUT_ID_Ext) = 1889001 Then
''                RetGrp_Left = UT_Pos_Left
''                RetGrp_Top = UT_Pos_Top
''            ElseIf Val(sUT_ID_Ext) = 645000 Then
''                ActiveWorkbook.Names.Add Name:="RETGRP", RefersTo:=Range(ActiveSheet.Cells(RetGrp_Top, RetGrp_Left), ActiveSheet.Cells(RetGrp_Top + 15, RetGrp_Left))
''            End If
            
            ' Set the named range
            CBA_ErrTag = "Rng"
            sRngName = "RR" & UT_Pos_Top & "CC" & UT_Pos_Left
            Set rRng = pwbWSht.Range(.Cells(UT_Pos_Top, UT_Pos_Left), .Cells(lUdtBot, lLen))
''            ActiveWorkbook.Names.Add Name:=sRngName, RefersTo:=rRng                   ''''', Visible:=False
            ' Apply the formatting required
            rRng.Merge
            CBA_ErrTag = ""
            If UT_Image > "" Then
                .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\" & UT_Image
            End If
            
            If UT_Wrap_Text = "Y" Then rRng.WrapText = True
            If UT_BG_Color > 0 Then rRng.Interior.ColorIndex = UT_BG_Color
            If UT_FG_Color > 0 Then rRng.Font.ColorIndex = UT_FG_Color
            If UT_Font_Size > 0 Then rRng.Font.Size = UT_Font_Size
            If UT_Bold = "Y" Then rRng.Font.Bold = True
            If UT_Italic = "Y" Then rRng.Font.Italic = True
            If UT_Underline = "Y" Then rRng.Font.Underline = True
            If UT_StrikeThrough = "Y" Then rRng.Font.Strikethrough = True
            If UT_Merge_Height = 0 Then rRng.Height = 3
            
            ' Horisontal Alignment
            Select Case Left(UT_TextAlign, 1)
                Case "L"
                    rRng.HorizontalAlignment = xlHAlignLeft
                Case "C"
                    rRng.HorizontalAlignment = xlHAlignCenter
                Case "R"
                    rRng.HorizontalAlignment = xlHAlignRight
            End Select
            ' Vertical Alignment
            Select Case Right(UT_TextAlign, 1)
                Case "T"
                    rRng.VerticalAlignment = xlVAlignTop
                Case "M", "C"
                    rRng.VerticalAlignment = xlVAlignCenter
                Case "B"
                    rRng.VerticalAlignment = xlVAlignBottom
            End Select
            ' Borders
            If Len(UT_Border) = 2 And Left(UT_Border, 1) = "F" Then
                If Val(Right(UT_Border, 1)) > 0 Then rRng.Borders.Weight = Val(Right(UT_Border, 1))
            ElseIf Len(UT_Border) > 1 Then
                UT_Border = UT_Border & "          "
                If Mid(UT_Border, 1, 1) = "T" And Val(Mid(UT_Border, 2, 1)) > 0 Then rRng.Borders(xlEdgeTop).Weight = Val(Mid(UT_Border, 2, 1)) + 1
                If Mid(UT_Border, 3, 1) = "R" And Val(Mid(UT_Border, 4, 1)) > 0 Then rRng.Borders(xlEdgeRight).Weight = Val(Mid(UT_Border, 4, 1)) + 1
                If Mid(UT_Border, 5, 1) = "B" And Val(Mid(UT_Border, 6, 1)) > 0 Then rRng.Borders(xlEdgeBottom).Weight = Val(Mid(UT_Border, 6, 1)) + 1
                If Mid(UT_Border, 7, 1) = "L" And Val(Mid(UT_Border, 8, 1)) > 0 Then rRng.Borders(xlEdgeLeft).Weight = Val(Mid(UT_Border, 8, 1)) + 1
            End If
            ' Other parts set-up
            If UT_Formula > "" Then
                UT_Formula = UT_Formula
            End If
            
            If UT_Formula = "COMRADEDiff" Then
                UT_Formula = "=IFERROR((" & ActiveSheet.Cells(rRng.Row - 2, rRng.Column).Address & "-MAX(T10RetGrp))/" & ActiveSheet.Cells(rRng.Row - 2, rRng.Column).Address & ",0)"
                rRng.Formula = UT_Formula
            ElseIf InStr(1, UT_Formula, "~") > 0 Then               ' Handle any formulas
                CBA_ErrTag = "Rng"
                aFla = Split(UT_Formula & "~~~", "~")
                If aFla(1) > "" Then
                ''ReDim aGrp(0 To 4): Set sdGrp = New Scripting.Dictionary: lGrp = -1
                    lGrp_No = g_Get_Mid_Fmt(UT_Grp_No, 2, 2)
                    lOrigGrp_No = g_Get_Mid_Fmt(UT_Orig_Grp_No, 2, 2)
                    sFla = Replace(aFla(1), "#", CStr(lGrp_No))
                    If Right(aFla(1), 1) = "#" Or lGrp_No = 0 Then
''                        ActiveWorkbook.Names.Add Name:=sFla, RefersTo:=rRng         'Range(ActiveSheet.Cells(UT_Pos_Top, UT_Pos_Left), ActiveSheet.Cells(UT_Pos_Top, UT_Pos_Left + ut_Width))
                        ActiveWorkbook.Names.Add Name:=sFla, RefersTo:=Range(ActiveSheet.Cells(UT_Pos_Top, UT_Pos_Left), ActiveSheet.Cells(UT_Pos_Top, UT_Pos_Left))
                    ElseIf lGrp_No = 1 Then
                        lGrp = lGrp + 1
                        sdGrp.Add sFla, lGrp
                        ReDim Preserve aGrp(0 To 4, 0 To lGrp)
                        aGrp(0, lGrp) = sFla
                        aGrp(1, lGrp) = g_GetExcelCell(UT_Pos_Top, UT_Pos_Left)
                    ElseIf lGrp_No = lOrigGrp_No Then
                        If sdGrp.Exists(sFla) Then
                            lIdx = sdGrp.Item(sFla)
                            aGrp(1, lIdx) = aGrp(1, lIdx) & ", " & g_GetExcelCell(UT_Pos_Top, UT_Pos_Left)
 ''                           ActiveWorkbook.Names.Add Name:=sFla, RefersTo:=Range(ActiveSheet.Cells(aGrp(1, lIdx), aGrp(2, lIdx)), ActiveSheet.Cells(UT_Pos_Top + UT_Merge_Height - 1, UT_Pos_Left + UT_Merge_Width - 1))
                            ActiveWorkbook.Names.Add Name:=sFla, RefersTo:=Range(aGrp(1, lIdx))
                        End If
                    ElseIf lGrp_No > 1 Then
                        If sdGrp.Exists(sFla) Then
                            lIdx = sdGrp.Item(sFla)
                            aGrp(1, lIdx) = aGrp(1, lIdx) & ", " & g_GetExcelCell(UT_Pos_Top, UT_Pos_Left)
                        End If
                    End If
                End If
                CBA_ErrTag = "Fmt"
                If aFla(2) > "" Then
                    sFla = Replace(aFla(2), "#", CStr(lGrp_No))
                    Do While InStr(1, sFla, "^") > 0
                        sFla = p_RtnStr(sFla, UT_Pos_Top)
                    Loop
                    ''sFla = Replace(aFla(2), """", "'x")
                    rRng.Formula = sFla
                End If
                If aFla(0) > "" Then UT_Formula = aFla(0)
            End If
            If UT_Hyperlink > "" Then
                rRng.Hyperlink = UT_Hyperlink
            End If
            If UT_Cond_Format > "" Then
                CBA_ErrTag = "Fmt"
                lGrp_No = g_Get_Mid_Fmt(UT_Grp_No, 2, 2)
                Call f_Cond_Formatting(Val(UT_Cond_Format), rRng, lGrp_No)
            End If
            ' Reset the Default Value if it is a split heading
            If InStr(1, UT_Default_Value, "~") > 0 And Left(UT_Source, 1) = "2" And UT_Hdg_Type = "Y" Then
                aHdg = Split(UT_Default_Value, "~")
                If psNPD_R_MSO = "NPD" Then
                    UT_Default_Value = aHdg(0)
                Else
                    UT_Default_Value = aHdg(1)
                End If
            End If
            ' Set the sts of the line (mostly group lines
            If UT_Pos_Left = 1 Then
''                If Mid(UT_Source, e_Src.eDsp, 1) = "H" Then
''                    UT_Procedure = UT_Procedure
''                End If
''                If UT_Pos_Top = 52 Or UT_Pos_Top = 56 Or UT_Pos_Top = 60 Then
''                    UT_Pos_Top = UT_Pos_Top
''                End If
                If Mid(UT_Source, e_Src.eDsp, 1) = "H" Then
                        sStsFlag = "H"
                ElseIf psdUTL.Exists(CStr(UT_Pos_Top)) Then
                  '  Debug.Print UT_Pos_Top & ";";
                    lUTLIdx = psdUTL.Item(CStr(UT_Pos_Top))
                    sStsFlag = pUTL_Ary(lUTLIdx).sTL_Flag
                  '  Debug.Print UT_Pos_Top & "=" & sStsFlag;
                    If UT_Default_Value = "" And Mid(UT_Source, e_Src.eDsp, 1) = "0" Then  'Get_TD_UT(lUT_Idx, "UT_Source" & e_Src.eDsp) = "0" Then
                        sStsFlag = "L"
                      '  Debug.Print "=(" & sStsFlag & ")";
                        Let pUTL_Ary(lUTLIdx).sTL_Flag = sStsFlag
                    End If
                  '  Debug.Print ";";
                Else
                    sStsFlag = "N"
                End If
            End If
            ' Field Type
            If Len(UT_Field_Type) >= 3 Then
                Select Case Left(UT_Field_Type, 3)
                    Case "dte"
                        If Len(UT_Field_Type) > 3 Then rRng.NumberFormat = g_Right(UT_Field_Type, 3)
                    Case "txt"
                        rRng.NumberFormat = xlGeneral
                    Case "txa"
                        rRng.NumberFormat = xlGeneral
                    Case "cur", "per", "num", "now"
                        If Len(UT_Field_Type) > 3 Then
                            rRng.NumberFormat = g_Right(UT_Field_Type, 3)
                        End If
                    Case "chk"              ' Create a CheckBox where specified
                        lChkNo = lChkNo + 1
                        Call ActiveSheet.AddCheckbox(lChkNo, IIf(UT_Default_Value = "Y", True, False), UT_Pos_Top, UT_Pos_Left, g_GetExcelCell(UT_Pos_Top, UT_Pos_Left))
                    Case "ddb"              ' Create a DropDownBox where specified
                        rRng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=UT_Formula
                        DoEvents
                    Case "btn"              ' Create a button at the end of the line where specified
                        bDspBtn = True
                        If sStsFlag = "L" Then
                            bDspBtn = False
                        End If
                        If bDspBtn = True Then
                            lBtnNo = lBtnNo + 1
                            Call ActiveSheet.AddButton(lBtnNo, g_Right(UT_Field_Type, 3), UT_Pos_Top, UT_Pos_Left, UT_Addit, g_GetExcelCell(UT_Pos_Top, UT_Pos_Left), g_GetExcelCell(lUdtBot, UT_Pos_Left))
                        End If
                End Select
            End If

            If UT_Default_Value > "" Then
                If Left(UT_Field_Type, 3) = "dte" And UT_Default_Value <> "" Then
                    .Cells(UT_Pos_Top, UT_Pos_Left).Value = "'" & g_FixDate(UT_Default_Value, g_Right(UT_Field_Type, 3))
                Else
                    .Cells(UT_Pos_Top, UT_Pos_Left).Value = UT_Default_Value
                End If
            End If

            If UT_Locked = "Y" Then
                rRng.Locked = True
            Else
                rRng.Locked = False
            End If
            ' Set the Height for the line
            If UT_Pos_Left = 1 Then
                If sStsFlag = "L" Or sStsFlag = "H" Then
                    rRng.RowHeight = 0
                Else
                    rRng.RowHeight = CBA_ROW_HEIGHT
                End If
''            ElseIf UT_Wrap_Text = "Y" Then                    ' @RWTen Couldn't get working
''                rRng.Rows.AutoFit
            End If
            If UT_Wrap_Text = "Y" Then rRng.WrapText = True                    ' @RWTen Couldn't get working
''                rRng.Rows.AutoFit

ReIter:
        Next
        ' Now go and hide the 'Hidden' Fields - can't do this earlier as it doesn't produce the buttons properly if they are 0 high (i.e. the height is set at Left=1)
        For lUTL_Idx = 0 To plUTL_Idx
            lRow = pUTL_Ary(lUTL_Idx).lTL_Row
''            If lRow = 52 Or lRow = 56 Or lRow = 60 Then
''                UT_Pos_Top = UT_Pos_Top
''            End If

            If pUTL_Ary(lUTL_Idx).sTL_Flag = "H" Then
                lRow = pUTL_Ary(lUTL_Idx).lTL_Row: lMergeHeight = pUTL_Ary(lUTL_Idx).lTL_Height
               .Rows(lRow & ":" & lRow + lMergeHeight - 1).RowHeight = 0
            End If
        Next
        
        lLen = NO_OF_COLS + 2
        .Activate
        Application.PrintCommunication = False
        Range(.Cells(1, 1), .Cells(lUdtBot, lLen - 1)).Select
        .PageSetup.FitToPagesTall = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.RightFooter = "Sheet: &A"
        .PageSetup.LeftMargin = PAGE_MARGIN
        .PageSetup.RightMargin = PAGE_MARGIN
        .PageSetup.TopMargin = PAGE_MARGIN
        .PageSetup.BottomMargin = PAGE_MARGIN
        .PageSetup.HeaderMargin = 1
        .PageSetup.FooterMargin = PAGE_MARGIN
        .PageSetup.PrintArea = "$A$1:" & g_GetExcelCell(lUdtBot, lLen, "$")

        Application.PrintCommunication = True
        .Cells(lFreezePos + 1, lLen - 1).Select
        ActiveWindow.FreezePanes = True
    End With
    
    ActiveSheet.Protect "Passw"
    Call g_SetupIP("TenderDocs", 1, False)
    Call ShowIP_ID("SetShowOrNot")
    Call ShowIP_ID("SetLoaded")

Exit_Routine:
    On Error Resume Next
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    Application.ScreenUpdating = True
    Exit Sub

Err_Routine:
    If Err.Number = 1004 And CBA_ErrTag <> "Rng" Then
        ''Debug.Print "Lock err " & UT_Default_Value & "; ";
        Resume Next
    End If
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("-s-mTEN_Runtime.CrtTenderDoc", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub

Private Function Set_TD_TD_Value(sProcedure As String, sName As String, sSource As String, sDefaultValue As String, lRow As Long, lcol As Long, Optional ByVal lGrp_Idx As Long = -1) As String
    ' Process where the data is going to come from...
    Static RegionCap As String
    Dim aRegs() As String, lGrp_No As Long
    Const a_Regs As String = "MIN,DER,STP,PRE,DAN,BRE,RGY,JKT"
    Set_TD_TD_Value = sDefaultValue
    ' Ensure that the region is only the last two digits of the Group_No
    lGrp_No = g_Get_Mid_Fmt(lGrp_Idx, 2, 2)
    If sName = "Regions(All)" Then
        aRegs = Split(a_Regs, ",")                              ' Deliver back the appropriate region name
        GoTo Regions
    ElseIf sName = "Regions" Then                               ' Deliver back the appropriate region name
        aRegs = Split(RegionCap, ",")
        GoTo Regions
    ElseIf sName = "RegionCap" Then                             ' Capture the regions when they come through
        RegionCap = sDefaultValue
        Set_TD_TD_Value = sDefaultValue
        'Exit Function
    ElseIf sName = "RegionNo" And lGrp_Idx = 0 Then             ' Deliver back the selected regions
        aRegs = Split(RegionCap, ",")
        Set_TD_TD_Value = CStr(UBound(aRegs, 1)) + 1
        'Exit Function
    ElseIf sProcedure = "TenderFor" Then
        Set_TD_TD_Value = f_TenderFor(sSource, sName, sDefaultValue, lRow, lcol, lGrp_Idx)
        'Exit Function
    ElseIf sProcedure = "Prod&MktOV" Then
        Set_TD_TD_Value = f_ProdMktOV(sSource, sName, sDefaultValue, lRow, lcol, lGrp_Idx)
        'Exit Function
    ElseIf sProcedure = "KeyMetrics" Then
        Set_TD_TD_Value = f_KeyMetrics(sSource, sName, sDefaultValue, lRow, lcol, lGrp_Idx)
        'Exit Function
    ElseIf sProcedure = "CompetitionAnalysis" Then
        Set_TD_TD_Value = f_CompetitionAnalysis(sSource, sName, sDefaultValue, lRow, lcol, lGrp_Idx)
        'Exit Function
    ElseIf sProcedure = "SimilarCoreRangeProd1" Then
        Set_TD_TD_Value = f_SimilarCoreRangeProd1(sSource, sName, sDefaultValue, lRow, lcol, lGrp_Idx)
        'Exit Function
    ElseIf sProcedure = "SimilarCoreRangeProd3" Then
        Set_TD_TD_Value = f_SimilarCoreRangeProd3(sSource, sName, sDefaultValue, lRow, lcol, lGrp_Idx)
        'Exit Function
    ElseIf sProcedure = "TenderSubmissions" Then
        Set_TD_TD_Value = f_TenderSubmissions(sSource, sName, sDefaultValue, lRow, lcol, lGrp_Idx)
        'Exit Function
    ElseIf sProcedure = "Recommendation" Then
        Set_TD_TD_Value = f_Recommendation(sSource, sName, sDefaultValue, lRow, lcol, lGrp_Idx)
    ElseIf sProcedure = "Save" Then
        Set_TD_TD_Value = f_Save(sSource, sName, sDefaultValue, lRow, lcol, lGrp_Idx)
''        Exit Function
    ElseIf sProcedure = "SignOffChecklist" Then
        Set_TD_TD_Value = f_SignOffChecklist(sSource, sName, sDefaultValue, lRow, lcol, lGrp_Idx)
        'Exit Function
    Else
        If g_Get_Dev_Sts("DevIP") = "Y" Then
            MsgBox sProcedure & " not found in Set_TD_TD_Value"
            Stop
        End If
    End If
    Exit Function
    
Regions:
    If LBound(aRegs, 1) > lGrp_No - 1 Or UBound(aRegs, 1) < lGrp_No - 1 Then
''        MsgBox lGrp_INo & " Group No is lessor or greater than number of regions..." & IIf(sName = "Regions(All)", a_Regs, RegionCap)
''        Stop
        Exit Function
    End If
    Set_TD_TD_Value = aRegs(lGrp_No - 1)
    Exit Function
End Function

Public Function Write_db_Class_ID(Optional ByRef lID_ID As Long = 0, Optional ByRef lLV_ID As Long = 0) As Long
    ' This routine will write the Idea Class & Lines to the database
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim sSQL As String, lHIdx As Long, lLIdx As Long, sUpd As String
    Dim ID_ID As Long, lATemp As Long
    Dim LV_ID As Long, PF_ID As Long, PV_ID As Long, PD_ID As Long, PC_ID As Long
    On Error GoTo Err_Routine

    CBA_ErrTag = "SQL"
    Write_db_Class_ID = 0
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Ten") & ";"
    ' Add or update the Idea Header/s
    For lHIdx = 1 To plID_Idx
        ID_ID = Get_TD_ID(lHIdx, "ID_ID")
        sUpd = Get_TD_ID(ID_ID, "ID_Upd")
        If sUpd = "A" Then
            sSQL = "INSERT INTO L1_Ideas ( ID_Desc,ID_Sts_ID,ID_BA,ID_BD,ID_BA_Emp_No,ID_BD_Emp_No,ID_GD_Emp_No,ID_UpdUser,ID_CrtUser )" & Chr(10) & " VALUES ( "
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_Desc", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_Sts_ID", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_BA", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_BD", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_BA_Emp_No", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_BD_Emp_No", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_GD_Emp_No", True, True) & "," & Chr(10)
            sSQL = sSQL & "'" & CBA_User & "','" & CBA_User & "' )"
            ' Write the SQL
            RS.Open sSQL, CN
            ' Get the new ID
            lATemp = g_DLookup("ID_ID", "L1_Ideas", "ID_ID>0", "ID_ID DESC", g_GetDB("Ten"), 0)
            Call Upd_TD_ID(ID_ID, "ID_ID", lATemp, "N")
            lID_ID = lATemp
            ' Put the new header ID in the class lines
            LV_ID = 1
            ' For each row in the array...
            Do While LV_ID > 0
                LV_ID = Get_ID_LV_Idx(ID_ID, "+", "LV_ID")
                If LV_ID > 0 Then
                    Call Upd_TD_LV(LV_ID, "ID_ID", lATemp, "N")
                End If
            Loop
        ElseIf sUpd = "U" Then
            sSQL = "UPDATE L1_Ideas SET "
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_Desc", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_Sts_ID", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_BA", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_BD", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_BA_Emp_No", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_BD_Emp_No", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_ID(ID_ID, "ID_GD_Emp_No", , True) & "," & Chr(10)
            sSQL = sSQL & "ID_UpdUser='" & CBA_User & "', ID_UpdDate=" & g_GetSQLDate(Now(), CBA_DMYHN)
            sSQL = sSQL & " WHERE ID_ID=" & ID_ID
            RS.Open sSQL, CN
        End If
    Next

    ' Add or update the Lines
    For lLIdx = 1 To plLV_Idx
        ' Get the LevelID   --- Have takjen ouyt the ATP updates from here as they will be done in the Levels Update after the Tnder Doc is saved
        LV_ID = Get_TD_LV(lLIdx, "LV_ID")
        If Get_TD_LV(LV_ID, "LV_Upd") = "A" Then
            sSQL = "INSERT INTO L2_Levels ( LV_ID_ID,LV_Portfolio_No,LV_Portfolio_Desc,LV_Version_No,LV_Version_Desc,LV_Product_Code,LV_Prior_Product_Code,LV_Contract_No,LV_TH_Docs," & _
                    "LV_Sts_ID,LV_TH_ID,LV_CrtUser,LV_UpdUser )" & Chr(10)
''            sSQL = "INSERT INTO L2_Levels ( LV_ID_ID,LV_Portfolio_No,LV_Portfolio_Desc,LV_Version_No,LV_Version_Desc,LV_Product_Code,LV_Prior_Product_Code,LV_Contract_No,LV_TH_Docs," & _
''                    "LV_Sts_ID,LV_ATPFile,LV_TH_ID,LV_AH_ID,LV_CrtUser,LV_UpdUser )" & Chr(10)
            sSQL = sSQL & "VALUES ( " & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_ID_ID", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Portfolio_No", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Portfolio_Desc", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Version_No", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Version_Desc", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Product_Code", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Prior_Product_Code", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Contract_No", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_TH_Docs", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Sts_ID", True, True) & "," & Chr(10)
''            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_ATPFile", True, True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_TH_ID", True, True) & "," & Chr(10)
''            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_AH_ID", True, True) & "," & Chr(10)
            sSQL = sSQL & "'" & CBA_User & "','" & CBA_User & "' )"
            CBA_ErrTag = sSQL
            RS.Open sSQL, CN
            ' Get the new ID
            lATemp = g_DLookup("LV_ID", "L2_Levels", "LV_ID>0", "LV_ID DESC", g_GetDB("Ten"), 0)
            Call Upd_TD_LV(LV_ID, "LV_ID", lATemp, "N")
            lLV_ID = lATemp
        ElseIf Get_TD_LV(LV_ID, "LV_Upd") <> "N" Then
            sSQL = "UPDATE L2_Levels SET LV_UpdDate=" & g_GetSQLDate(Now, CBA_DMYHN) & ", LV_UpdUser='" & CBA_User & "',"
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Portfolio_No", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Portfolio_Desc", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Version_No", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Version_Desc", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Product_Code", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Prior_Product_Code", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Contract_No", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_TH_Docs", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_Sts_ID", , True) & "," & Chr(10)
''            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_ATPFile", , True) & "," & Chr(10)
            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_TH_ID", , True) & Chr(10)
''            sSQL = sSQL & Get_TD_LV(LV_ID, "LV_AH_ID", , True) & Chr(10)
            sSQL = sSQL & " WHERE LV_ID=" & LV_ID
            CBA_ErrTag = sSQL
            RS.Open sSQL, CN
        End If
    Next

Exit_Routine:

    On Error Resume Next
    CN.Close
    Set RS = Nothing
    Set CN = Nothing
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Exit Function
    
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-Write_db_Class_ID", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function Write_db_Class_TD() As Long
    ' This routine will write the Class Tender Documemt & Lines to the database
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim sSQL As String, lUT_Idx As Long, lLIdx As Long, lUTLIdx As Long, lTDNo As Long, wbWB As Workbook, Last_SL_ID As Long
    Dim lATemp As Long, bHdrDone As Boolean, sTemplates_Used As String, bUTLExists As Boolean
    Dim sUT_ID As String, sUT_Pos_Top As String, sUT_Prior_Top As String, SL_SH_ID As String, sUTL_Flag As String, sUTL_Upd As String, UT_Orig_Grp_No As Long  '', lSegIdx As Long '', PD_ID As Long, PC_ID As Long
    Dim UT_Pos_Left As Long, UT_Pos_Top As Long, UT_SL_ID As Long, UT_Grp_No As Long, UT_Hdg_Type As String, UT_Default_Value As String, UT_Field_Type As String, UT_Source As String, UT_Addit As String, UT_Upd As String

    On Error GoTo Err_Routine
    
    CBA_BasicFunctions.CBA_Running "Saving Data for " & psSH_Desc
    Application.ScreenUpdating = False
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "Saving " & IIf(psNPD_R_MSO = "R", "Retender", psNPD_R_MSO) & " Data..."
    Write_db_Class_TD = 0
    CBA_ErrTag = "SQL"
    Select Case psNPD_R_MSO
        Case "NPD"
            lTDNo = 1
        Case "R"
            lTDNo = 2
        Case "MSO"
            lTDNo = 3
        Case Else
            MsgBox psNPD_R_MSO & " (NPD_R_MSO) Not found in mTEN_Runtime.Write_db_Class_TD"
    End Select
''    Call FilePrint(vbCrLf & vbCrLf & vbCrLf & "Write", 0, 0)
    bHdrDone = False
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Ten") & ";"
    ' Add or update the TD Header/s
    For lUT_Idx = 1 To plUT_Idx2
        ' Get the required variables
        Call pUT_Ary2(lUT_Idx).Get_Class_Ltd_UT(UT_Pos_Left, UT_Pos_Top, UT_SL_ID, UT_Grp_No, UT_Hdg_Type, UT_Default_Value, UT_Field_Type, UT_Source, UT_Addit, UT_Orig_Grp_No, UT_Upd)
        ' Don't Add or update headings
        If UT_Hdg_Type = "Y" Then GoTo GoNext
        ' Ensure only the selected segments are being written
        If Last_SL_ID <> UT_SL_ID Then
            SL_SH_ID = Get_TD_SL(UT_SL_ID, "SL_SH_ID")
            Last_SL_ID = UT_SL_ID
        End If
        If InStr(1, SL_SH_ID, "," & lTDNo & ",") = 0 Then GoTo GoNext
        ' Test for each Grouped UDT (If in the future psdUTL is every line, this will have to be looked at)
        If UT_Pos_Left = 1 Then
            If Mid(UT_Source, e_Src.eDsp, 1) = "H" Then
                UT_Default_Value = UT_Default_Value
            End If
            bUTLExists = False
            If UT_Grp_No > 0 Then
                sUT_Pos_Top = CStr(UT_Pos_Top)
                If psdUTL.Exists(sUT_Pos_Top) Then
                    bUTLExists = True
                    sUT_Prior_Top = sUT_Pos_Top
                ElseIf psdUTL.Exists(sUT_Prior_Top) Then
                    bUTLExists = True
                    sUT_Pos_Top = sUT_Prior_Top
                End If
            Else
                sUT_Prior_Top = sUT_Pos_Top
            End If
            If bUTLExists Then
                lUTLIdx = psdUTL.Item(sUT_Pos_Top)
                sUTL_Flag = pUTL_Ary(lUTLIdx).sTL_Flag
            Else
                sUTL_Flag = "N"
            End If
        End If
        ' If taken from a prior Template Type Linked field, then it is really an add
        If Val(UT_Upd) > 0 Then UT_Upd = "A"
        
        ' If a valid UTL Line then flag the whole line as per the UT_Pos_Left and UTL calculations
        If bUTLExists Then
            If sUTL_Flag = "L" Then                     ' If empty non shown line then no add/update
                UT_Upd = "N"
            Else
                If sUTL_Upd = "U" Then
                    If UT_Upd <> "A" Then UT_Upd = "U"
                End If
            End If
        End If
        ' No point in writing blank records
        If UT_Upd = "A" And UT_Default_Value = "" Then
            UT_Upd = "N"
        End If
        ' Only Lines that need to be Updated or Added
        If UT_Upd <> "N" Then
            sUT_ID = Get_TD_UT(lUT_Idx, "UT_ID")
''            Call FilePrint("", CLng(sUT_ID_Ext), lUT_Idx)
            If bHdrDone = False Then
                bHdrDone = True
                If plTH_ID = 0 Then
                    sSQL = "INSERT INTO L7_TDoc_Hdrs ( TH_SH_ID,TH_LV_ID,TH_SU_ID,TH_UpdUser,TH_CrtUser )" & Chr(10) & " VALUES ( "
                    sSQL = sSQL & plSH_ID & ","
                    sSQL = sSQL & plLV_ID & ","
                    sSQL = sSQL & plSU_ID & ","
                    sSQL = sSQL & "'" & CBA_User & "','" & CBA_User & "' )"
                    ' Write the SQL
                    RS.Open sSQL, CN
                    ' Get the new ID
                    lATemp = g_DLookup("TH_ID", "L7_TDoc_Hdrs", "TH_ID>0", "TH_ID DESC", g_GetDB("Ten"), 0)
''                    Call Upd_TD_ID(plTH_ID, "TH_ID", lATemp, "N")
                    plTH_ID = lATemp
                    Call TEN_Set_Values(, , , plTH_ID)
                Else
                    sSQL = "UPDATE L7_TDoc_Hdrs SET "
                    sSQL = sSQL & "TH_UpdUser='" & CBA_User & "', TH_UpdDate=" & g_GetSQLDate(Now(), CBA_DMYHN)
                    If plSU_ID > 0 Then sSQL = sSQL & ", TH_SU_ID = " & plSU_ID
                    sSQL = sSQL & " WHERE TH_ID=" & plTH_ID
                    RS.Open sSQL, CN
                End If
            End If
''            If Val(sUT_ID_Ext) = 53500 Or Val(sUT_ID_Ext) = 73100 Then
''                sUT_ID_Ext = sUT_ID_Ext
''            End If
''            UT_ID = Val(g_Left(CStr(sUT_ID_Ext), 3))
''            UT_Grp_No = Get_TD_UT(lUT_Idx, "UT_Grp_No") '''Val(Right(sUT_ID_Ext, 3))
            ' Take out the first single apostrophy if it has been inserted - all others will be changed to "`"
            If UT_Field_Type = "txa" And Left(UT_Default_Value, 1) = "'" Then UT_Default_Value = g_Right(UT_Default_Value, 1)
            If UT_Upd = "A" Then
                sSQL = "INSERT INTO L9_TDoc_UDTs ( TU_TH_ID,TU_UT_ID,TU_UT_No,TU_Template_No,TU_Value,TU_Sts_Flg)" & Chr(10)
                sSQL = sSQL & "VALUES ( "
                sSQL = sSQL & plTH_ID & ","
                sSQL = sSQL & sUT_ID & ","
                sSQL = sSQL & UT_Grp_No & ","
                sSQL = sSQL & lTDNo & ","
                sSQL = sSQL & "'" & Replace(UT_Default_Value, "'", "`") & "',"
                sSQL = sSQL & "'" & sUTL_Flag & "'"
                sSQL = sSQL & " )"
                CBA_ErrTag = sSQL
                RS.Open sSQL, CN
            Else
                sSQL = "UPDATE L9_TDoc_UDTs SET "
                sSQL = sSQL & "TU_Value='" & Replace(UT_Default_Value, "'", "`") & "'," & Chr(10)
                sSQL = sSQL & "TU_Sts_Flg='" & sUTL_Flag & "'" & Chr(10)
                sSQL = sSQL & " WHERE TU_UT_ID=" & sUT_ID & " AND TU_UT_No=" & UT_Grp_No & " AND TU_Template_No=" & lTDNo
                CBA_ErrTag = sSQL
                RS.Open sSQL, CN
            End If
        End If
GoNext:
    Next
    ' If records have been added or updated then update the Levels header
    If bHdrDone = True Then
        sTemplates_Used = Get_TD_LV(plLV_ID, "LV_TH_Docs")
        If sTemplates_Used = "" Then
            sTemplates_Used = psNPD_R_MSO
        Else
            If InStr(1, sTemplates_Used, psNPD_R_MSO) = 0 Then sTemplates_Used = sTemplates_Used & "-" & psNPD_R_MSO
        End If
        ' Upd the LevelID
        sSQL = "UPDATE L2_Levels SET LV_UpdDate=" & g_GetSQLDate(Now, CBA_DMYHN) & ", LV_UpdUser='" & CBA_User & "',"
        If plAH_ID = 0 And Get_TD_LV(plLV_ID, "LV_ATPFile") > "" Then                   ' Won't update unless there is something in the ATPFile
            sSQL = sSQL & " LV_AH_ID=1,"                                                ' Make LV_AH_ID=1 for now as Tender Comparison Documents aren't yet written to their respective tables
            sSQL = sSQL & " LV_ATPFile= '" & Get_TD_LV(plLV_ID, "LV_ATPFile") & "',"
        End If
        sSQL = sSQL & " LV_TH_ID=" & plTH_ID & ","
        sSQL = sSQL & " LV_TH_Docs='" & sTemplates_Used & "'"
        sSQL = sSQL & " WHERE LV_ID=" & plLV_ID
        CBA_ErrTag = sSQL
        RS.Open sSQL, CN
    End If
    Write_db_Class_TD = 1
    On Error Resume Next
    ' If the plATPLV_ID > 0 but not matching plPLV_ID, the doc has been loaded against a different Leval than it will be saved against, so null it out
    If plATPLV_ID > 0 And plATPLV_ID <> plLV_ID And Get_TD_LV(plATPLV_ID, "LV_ATPFile") > "" And Get_TD_LV(plATPLV_ID, "LV_AH_ID") = 0 Then
        Call Upd_TD_LV(plATPLV_ID, "LV_ATPFile", "", "N")           ' Null out the non matching Doc
    End If
    ' Also update the Templates Used field in the Levels Class module, and the Tender Comparison Document if entered
    If plAH_ID = 0 And Get_TD_LV(plLV_ID, "LV_ATPFile") > "" Then
        Call Upd_TD_LV(plLV_ID, "LV_AH_ID", 1, "N")                 ' Make LV_AH_ID=1 for now as Tender Comparison Documents aren't yet written to their respective tables
        plAH_ID = 1
    End If
    plATPLV_ID = 0: plATP_Cols = 0
    Call Upd_TD_LV(plLV_ID, "LV_TH_Docs", sTemplates_Used, "N")
    Call Upd_TD_LV(plLV_ID, "LV_TH_ID", plTH_ID, "N")
Exit_Routine:

    On Error Resume Next
    CN.Close
    Set RS = Nothing
    Set CN = Nothing
    ' Close the worksheet
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    ' Delete the old file...
    Call TenWSDelete
    Exit Function
    
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-Write_db_Class_TD", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function Btn_Add(UHStr As String)
    ' Will process the Add... changes to the UT Line
    ' UHStr is an array of UT Rows that have been calculated by the CBA_TEN_WBK worksheet. These are the rows that have been affected by the button press
    Dim lUTLIdx As Long, lIdx As Long, lRow As Long, UHAry() As String, UTFlag As String, UTUpd As String, bUpd As Boolean
    Dim lHdgIdx As Long, lUT_Idx As Long, lNextTop As Long, sRowCol As String, lCurrTop As Long, lMergeH As Long
    UHAry = Split(UHStr, ",")
    For lIdx = 0 To UBound(UHAry, 1)
        lRow = UHAry(lIdx): bUpd = False
        If psdUTL.Exists(CStr(lRow)) Then
            lUTLIdx = psdUTL.Item(CStr(lRow)): UTFlag = pUTL_Ary(lUTLIdx).sTL_Flag: UTUpd = pUTL_Ary(lUTLIdx).sTL_Upd: bUpd = False
            If UTFlag <> "N" Then
                Let pUTL_Ary(lUTLIdx).sTL_Flag = "N"
                bUpd = True
            End If
            ' Update if changed - Note: The U here doesn't interfere with the A/U of the UT Arrays - is taken into consideration (if Supp = blank=N, else Add/Upd) at save time
            ''Debug.Print "r=" & lRow;
            If bUpd = True And UTUpd <> "U" Then
                Let pUTL_Ary(lUTLIdx).sTL_Upd = "U"
                ''Debug.Print ";U---";
''            Else
                ''Debug.Print ";N---";
            End If
        Else
            MsgBox lRow & " can't be found in Btn_Add"
        End If
    Next
    ' Update the header button, with the new difference
    sRowCol = g_Fmt_2_IDs(lRow, 1, e_UTFldFmt.eRowCol)
    lUT_Idx = psdRC_Idx.Item(sRowCol)
    lHdgIdx = Val(Get_TD_UT(lUT_Idx, "UT_Addit" & e_Addit.eHdrIdx))
    lCurrTop = Val(Get_TD_UT(lHdgIdx, "UT_Addit" & e_Addit.eDiff))                          ' Get the current difference
    lMergeH = Val(Get_TD_UT(lHdgIdx, "UT_Addit" & e_Addit.eMHgt))                           ' Get the current merge height
    lNextTop = lCurrTop + (lMergeH * NUM_REGIONS)                                           ' Times it out
    Call Upd_TD_UT(lHdgIdx, "UT_Addit" & e_Addit.eDiff, lNextTop)                           ' Write the Row Diff between the 1st hidden line to the to the Header Line
    'Debug.Print "====" & Get_TD_UT(lHdgIdx, "UT_Addit")

''    MsgBox "The Add Button has been pressed"
End Function

Private Function Btn_Hide(lRow As Long, lcol As Long)
    Dim lUTLIdx As Long
    If psdUTL.Exists(CStr(lRow)) Then
        lUTLIdx = psdUTL.Item(CStr(lRow))
        Let pUTL_Ary(lUTLIdx).sTL_Flag = "H"
        Let pUTL_Ary(lUTLIdx).sTL_Upd = "U"
    Else
        MsgBox lRow & " can't be found in Btn_Hide"
    End If
End Function

Private Function Btn_Unhide(UHStr As String)
    ' Will process the UnHide changes to the UT Line
    ' UHStr is an array of UT Rows that have been calculated by the CBA_TEN_WBK worksheet. These are the rows that have been affected by the button press
    Dim lUTLIdx As Long, lIdx As Long, lRow As Long, UHAry() As String, UTFlag As String, UTUpd As String, bUpd As Boolean
    Debug.Print UHStr
    UHAry = Split(UHStr, ",")
    For lIdx = 0 To UBound(UHAry, 1)
        lRow = UHAry(lIdx): bUpd = False
        If psdUTL.Exists(CStr(lRow)) Then
            lUTLIdx = psdUTL.Item(CStr(lRow)): UTFlag = pUTL_Ary(lUTLIdx).sTL_Flag: UTUpd = pUTL_Ary(lUTLIdx).sTL_Upd: bUpd = False
''            If UTFlag = "N" Then
''''                Let pUTL_Ary(lUTLIdx).sTL_Flag = "H"
''''                bUpd = True
''            Else
            If UTFlag = "H" Then
                Let pUTL_Ary(lUTLIdx).sTL_Flag = "N"
                bUpd = True
            End If
            ' Update if changed - Note: The U here doesn't interfere with the A/U of the UT Arrays - is taken into consideration at save time
            If bUpd = True And UTUpd <> "U" Then Let pUTL_Ary(lUTLIdx).sTL_Upd = "U"
        Else
            MsgBox lRow & " can't be found in Btn_UnHide"
        End If
    Next
End Function

''Private Function Btn_Select(lRow As Long, lCol As Long)
''    Call f_Recommendation("Select", "Supplier", "", lRow, lCol)
''    Call f_Recommendation("Select", "Regions", "", lRow, lCol)
''End Function
''
''Private Function Btn_Remove(lRow As Long, lCol As Long)
''    Call f_Recommendation("Remove", "Supplier", "", lRow, lCol)
''    Call f_Recommendation("Remove", "Regions", "", lRow, lCol)
''End Function

Private Function Btn_Save()
    Dim sR As String
    ' Save the Tender Document
    sR = IIf(psNPD_R_MSO = "R", "Retender", psNPD_R_MSO)
    If Write_db_Class_TD = 1 Then
        MsgBox sR & " Tender Doc has been saved successfully"
    Else
        MsgBox sR & " Tender Doc had errors in saving"
    End If
    Call fTEN_1_Ideas.p_FillListBox_LV
End Function

Public Function Btn_Transfer(sAddit As String) As String
    ' Will process the Transfer of the Tender Submissions to the Recommendation section
    On Error GoTo Err_Routine
    Dim lLinIdx As Long, lEleIdx As Long, lRecIdx As Long, lLinkId As Long, sUT_Default_Value As String, sSupp As String, aAd() As String, saFlg(1 To NUM_REGIONS) As String, lFlgIdx As Long
    Dim lRow As Long, lcol As Long, lLinkIdx As Long, lTheIdx As Long, sAccStr As String, sSep As String
    Btn_Transfer = ""
'    plTSStart = 269: plTSEnd = 844: plTSDiff = 8
    ' Null out the value
    For lLinIdx = plTSStart To plTSEnd Step plTSDiff
        Call Upd_TD_UT(lLinIdx, "UT_Addit" & e_Addit.eFlgIdx, 0)                         ' Zero the flags for each Tender Submission UDT
    Next
    varFldVars.lFrmLeft = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
    varFldVars.lFrmTop = g_PosForm(0, 0, 0, "Top")
    varFldVars.sHdg = "Recommendation Selection"
    varFldVars.sSQL = ""
    varFldVars.sDB = "TEN"
    varFldVars.sField1 = sAddit
    varFldVars.sField2 = Format(plTSStart, "0000") & "," & Format(plTSEnd, "0000") & "," & Format(plTSDiff, "00") & ","
    varFldVars.sField2 = varFldVars.sField2 & Format(e_Addit.eDDP, "00") & "," & Format(e_Addit.eExW, "00") & "," & Format(e_Addit.eFOB, "00") & "," & Format(e_Addit.eFlgIdx, "00")
'    If pbRW = True Then Debug.Print "Btn_Tfr-Act;";
    fTEN_0_Recommend.Show vbModal
''    If pbRW = True Then Debug.Print "Btn_Tfr-Exit;";
    Call fTEN_1_Ideas.p_Dummy
    If CBA_bDataChg = False Then GoTo Exit_Routine
    Btn_Transfer = "Returned"
''    If pbRW = True Then Debug.Print "-" & Btn_Transfer
    lFlgIdx = 0: sAccStr = "": sSep = ""
    ' Distribute the values selected
    For lLinIdx = plTSStart To plTSEnd Step plTSDiff
        lFlgIdx = lFlgIdx + 1
        If lFlgIdx > NUM_REGIONS Then lFlgIdx = 1
        sAddit = Get_TD_UT(lLinIdx, "UT_Addit")                ' & e_Addit.eFlgIdx)                        ' Get the Tender Submission UDT Addit No
        aAd = Split(sAddit, ",")
        lEleIdx = Val(aAd(e_Addit.eFlgIdx - 1))
        If lEleIdx = 0 Then GoTo NextIdx
        ''lEleIdx = 1
        lRecIdx = Val(aAd(lEleIdx - 1))
        sSupp = "": sUT_Default_Value = ""
        If lEleIdx > 0 And lEleIdx - 1 < 5 And lRecIdx > 0 Then
            sSupp = Get_TD_UT(lLinIdx, "UT_Default_Value")
            sUT_Default_Value = Get_TD_UT(lRecIdx, "UT_Default_Value")
        End If
        If saFlg(lFlgIdx) <> "U" And sSupp > "" Then
            saFlg(lFlgIdx) = "U"
            With pwbWSht
                lTheIdx = lLinIdx
                GoSub GSGetData
                .Cells(lRow, lcol).Value = sSupp
                sAccStr = sAccStr & sSep & .Cells(lRow, 7).Value
                sSep = ","
                lRecIdx = lLinIdx + 2
                lTheIdx = lRecIdx
                GoSub GSGetData
                .Cells(lRow, lcol).Value = sUT_Default_Value
                lLinkIdx = lLinkIdx + 2
                lRow = Get_TD_UT(lLinkIdx, "UT_Pos_Top"): lcol = Get_TD_UT(lLinkIdx, "UT_Pos_Left")
                .Cells(lRow, lcol).Value = IIf(lEleIdx = 2, "DDP", IIf(lEleIdx = 3, "ExW/FG", IIf(lEleIdx = 4, "FOB", "???")))
            End With
        End If
NextIdx:
    Next
    ' Update the Pro & Mkting OV Regions with the new values
    pwbWSht.Cells(plOVRow, plOVCol).Value = sAccStr
    ' Null any now unupdated UTs if they have a value
    lFlgIdx = 0
    For lLinIdx = plTSStart To plTSEnd Step plTSDiff
        lFlgIdx = lFlgIdx + 1
        If lFlgIdx > NUM_REGIONS Then Exit For
        If saFlg(lFlgIdx) <> "U" Then
            sAddit = Get_TD_UT(lLinIdx, "UT_Addit")
            lEleIdx = Val(aAd(e_Addit.eFlgIdx - 1))
            If lEleIdx > 0 Then
                GoTo NextIdx2
            End If
            lRecIdx = Val(aAd(lEleIdx))
            sSupp = "": sUT_Default_Value = ""
''            If lEleIdx > 0 And lEleIdx < 5 And lRecIdx > 0 Then
''                sSupp = Get_TD_UT(lLinIdx, "UT_Default_Value")
''                sUT_Default_Value = Get_TD_UT(lRecIdx, "UT_Default_Value")
''            End If
            saFlg(lFlgIdx) = "U"
            With pwbWSht
                lTheIdx = lLinIdx
                GoSub GSGetData
                If .Cells(lRow, lcol).Value > "" Then
                    .Cells(lRow, lcol).Value = ""
                    lRecIdx = lLinIdx + 2
                    lTheIdx = lRecIdx
                    GoSub GSGetData
                    .Cells(lRow, lcol).Value = ""
                    lLinkIdx = lLinkIdx + 2
                    lRow = Get_TD_UT(lLinkIdx, "UT_Pos_Top"): lcol = Get_TD_UT(lLinkIdx, "UT_Pos_Left")
                    .Cells(lRow, lcol).Value = ""
                End If
            End With

        End If
NextIdx2:
    Next
    
Exit_Routine:
    On Error Resume Next
    Exit Function
GSGetData:
    lLinkId = Get_TD_UT(lTheIdx, "UT_Link_ID")
    lLinkId = g_Fmt_2_IDs(g_Get_Mid_Fmt(lLinkId, 9, 5), g_Get_Mid_Fmt(lLinkId, 2, 2), 4)
    lLinkIdx = psdUT_Idx2.Item(CStr(lLinkId))
    lRow = Get_TD_UT(lLinkIdx, "UT_Pos_Top"): lcol = Get_TD_UT(lLinkIdx, "UT_Pos_Left")

    Return

Err_Routine:
''    If Err.Number = 91 Then f_TenderFor = "": GoTo Exit_Routine
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-Btn_Transfer", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function

Private Function f_TenderFor(sSource As String, sName As String, sDefaultValue As String, lRow As Long, lcol As Long, Optional ByVal lGrp_Idx As Long = -1) As String
    ' Will return the 'Tender For' Type values
    On Error GoTo Err_Routine
    Dim iMonth As Integer, iYear As Integer
    Dim v As Variant, sngValue1 As Single, sngValue2 As Single, sngTotal As Single
    Dim PCode As String
    
    f_TenderFor = sDefaultValue
    Select Case sName
        Case "ProdCode"
            PCode = NZ(Get_TD_LV(plLV_ID, "LV_Product_Code"), "0")
            If PCode = "0" Then PCode = NZ(Get_TD_LV(plLV_ID, "LV_Prior_Product_Code"), "0")
            f_TenderFor = IIf(psNPD_R_MSO <> "NPD", PCode, "NPD")
        Case "IdeaDesc"
            If psNPD_R_MSO <> "NPD" Then
                f_TenderFor = Get_TD_LV(plLV_ID, "LV_Portfolio_Desc")
            Else
                f_TenderFor = Get_TD_ID(plID_ID, "ID_Desc")
            End If
        Case "TenderFor"
''            If Not pclsPFV Is Nothing Then f_TenderFor = pclsPFV.Description
            f_TenderFor = Get_TD_LV(plLV_ID, "LV_Portfolio_Desc")
        Case "Date"
            f_TenderFor = DateSerial(Year(Date), Month(Date), Day(Date))
        Case Else
            MsgBox "f_TenderFor " & sName & " not found"
    End Select
    
    
Exit_Routine:
    On Error Resume Next
    Exit Function
Err_Routine:
''    If Err.Number = 91 Then f_TenderFor = "": GoTo Exit_Routine
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-f_TenderFor", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function f_ProdMktOV(sSource As String, sName As String, sDefaultValue As String, lRow As Long, lcol As Long, Optional ByVal lGrp_Idx As Long = -1) As String
    Dim Prd As cCBA_Prod, bfound As Boolean
    Dim Val1 As Variant, Val2 As Variant, Val3 As Variant
    Dim PrdGrp As cCBA_ProdGroup
    Dim PCode As String
    
    On Error GoTo Err_Routine
    ' If there is already data for this item then skip it
    If plTH_ID > 0 Then
        f_ProdMktOV = sDefaultValue
        Exit Function
    End If
    
    If PrdGrp Is Nothing Or lGrp_Idx = 0 Then Set PrdGrp = pclsProdGrp
    PCode = NZ(Get_TD_LV(plLV_ID, "LV_Product_Code"), "0")
    If PCode = "0" Then PCode = NZ(Get_TD_LV(plLV_ID, "LV_Prior_Product_Code"), "0")
    If PCode = "0" Then
    Else
        If PrdGrp.getProductState(CLng(PCode)) Then
            Set Prd = PrdGrp.getProdObject(PCode, True)
        Else
            Set Prd = Nothing
        End If
    End If
    If Not psdND Is Nothing Then If pclsPFV.LCGCGno > 5 Then Set clsND = psdND(CStr(pclsPFV.LCGCGno))(CStr(pclsPFV.LCGSCGno))
    Select Case sName
        Case "ProdClass"
            If Prd Is Nothing Then Else f_ProdMktOV = Prd.lPClass
            'NEED TO PULL CLASS FROM PORTFOLIO
        Case "TenderType"
            f_ProdMktOV = psNPD_R_MSO
            If psNPD_R_MSO = "NDP" Then f_ProdMktOV = "Trial"
            If psNPD_R_MSO = "R" Then f_ProdMktOV = "Retender"
        Case "ContractType"
            f_ProdMktOV = "Please Select"
        Case "Amount"
            'f_ProdMktOV = pclsPFV.getQuantity(1035, False)
        Case "Units"
            'If pclsPFV.getQuantity() > 0 Then f_ProdMktOV = "Units"
        Case "RegionCap"
            'pclsPFV.
            'CAN YOU PULL REGIONS FROM PORTFOLIO
        Case "ProdDesc"
            'If Prd Is Nothing Then f_ProdMktOV = pclsPFV.Description Else f_ProdMktOV = Prd.sProdDesc
        Case "Rationale"
            f_ProdMktOV = "Enter concise summary of why this product should be ranged. For example: Category Review action, Gap in range, in response to market activity"
        Case "CmtforRetender"
            f_ProdMktOV = "Enter concise summary of why this product should be ranged. For example: Category Review action, Gap in range, in response to market activity"
        Case "NielsenCat"
            f_ProdMktOV = pclsPFV.LCGCGno & " - " & pclsPFV.LCGSCGno
        Case "MATDollars"
            If clsND Is Nothing Then f_ProdMktOV = 0 Else f_ProdMktOV = CStr(Round(clsND.Retail, 0))
        Case "MAT$Growth"
            If clsND Is Nothing Then f_ProdMktOV = 0 Else f_ProdMktOV = clsND.Retail - clsND.YOYRetail
        Case "PvtLbl$Share"
            If clsND Is Nothing Then f_ProdMktOV = 0 Else f_ProdMktOV = clsND.MarketPLShare * clsND.Retail
        Case "TmplType_NotShown"
''            If clsND Is Nothing Then f_ProdMktOV = 0 Else f_ProdMktOV = IIf(psNPD_R_MSO = "R", "Retender", psNPD_R_MSO)
        Case Else
            MsgBox "f_ProdMktOV " & sName & " not found"
    End Select
Exit_Routine:
        On Error Resume Next
        Exit Function
Err_Routine:
''     f_ProdMktOV = "": GoTo Exit_Routine
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-f_ProdMktOV", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function f_KeyMetrics(sSource As String, sName As String, sDefaultValue As String, lRow As Long, lcol As Long, Optional ByVal lGrp_Idx As Long = -1, Optional sPriorProd As String = "", Optional sProd As String = "") As String
    Dim Prd As cCBA_Prod, bfound As Boolean
    Dim Val1 As Variant, Val2 As Variant, Val3 As Variant
    Dim PrdGrp As cCBA_ProdGroup
    Dim PCode As Long
    On Error GoTo Err_Routine
    
    ' If there is already data for this item then skip it
    If plTH_ID > 0 Then
        f_KeyMetrics = sDefaultValue
        Exit Function
    End If
    'f_KeyMetrics = sDefaultValue
    If PrdGrp Is Nothing Or lGrp_Idx = 0 Then Set PrdGrp = pclsProdGrp
    PCode = NZ(Get_TD_LV(plLV_ID, "LV_Product_Code"), 0)
    If PCode = 0 Then PCode = NZ(Get_TD_LV(plLV_ID, "LV_Prior_Product_Code"), 0)
    If PrdGrp.getProductState(PCode) = True Then
        Set Prd = PrdGrp.getProdObject(PCode, True)
'        If Prd.getContractData("UC")(0, 0) = 0 Then
'            Debug.Print "Product in ProdGroup but no contract data exists"
'            Set PrdGrp = New cCBA_ProdGroup
'            PrdGrp.RunDataGeneration DateSerial(2001, 1, 1), pdteTo, False, Nothing, "BPC", Pcode
'            Set Prd = PrdGrp.getProdObject(PCode,True)
'            val1 = val1
'        End If
    Else
        'Debug.Print "Product Not Found in ProdGroup"
        If PCode > 0 Then
            Set PrdGrp = New cCBA_ProdGroup
            PrdGrp.RunDataGeneration pdteFrom, pdteTo, False, Nothing, "BP", PCode
            Set Prd = PrdGrp.getProdObject(PCode, True)
        End If
    End If
    
    
    Select Case sName
        Case "CG"
            f_KeyMetrics = pclsPFV.LCGCGno & "-" & pclsPFV.LCGCommodityGroup & vbCrLf & pclsPFV.LCGSCGno & "-" & pclsPFV.LCGSubCommodityGroup
        Case "AveCatUSW"
            f_KeyMetrics = PrdGrp.getUSW(e_USWorUSWpp.eUSWp, pdteFrom, pdteTo, e_RetailorQTY.eQTY, 1)
            'f_KeyMetrics = PrdGrp.QueryData(ePOSUSWpProd, , , , , , , , , , , , , , , , , , , pdteFrom, pdteTo)
            'f_KeyMetrics = CBA_CATREV_Runtime.getProdArrayData("AveCatUSW", , , , , , , , , , , , , , , , , , , DateSerial(Year(Date) - 1, Month(Date), 1), DateSerial(Year(Date), Month(Date), 0))
        Case "AveCatMargin%"
            f_KeyMetrics = PrdGrp.getRCVMargin(e_RCVMarginType.eRCVMarginPercent, pdteFrom, pdteTo, 1)
            'f_KeyMetrics = PrdGrp.QueryData(eRCVMargin, , , , , , , , , , , , , , , , , , , pdteFrom, pdteTo)
        Case "CatContri$-sw"
            Val1 = PrdGrp.getRCVMargin(e_RCVMarginType.eRCVMarginPercent, pdteFrom, pdteTo, 1)
            If Val1 <> 0 Then Val2 = PrdGrp.getUSW(e_USWorUSWpp.eUSWp, pdteFrom, pdteTo, e_RetailorQTY.eRetail, 1)
            If Val2 <> 0 Then f_KeyMetrics = Val1 * Val2 Else f_KeyMetrics = 0
        Case "ProdEst.USW"
            f_KeyMetrics = "=IF(T10USW=""""," & IIf(PCode = 0, 0, Prd.getPOSdata(pdteFrom, pdteTo, True, e_POSUSWTypes.eProductLevel)) & ",T10USW)"
        Case "ProdMargin%"
            f_KeyMetrics = "=IF(IFERROR(MODE(IF(T10Margin1="""",0,T10Margin1),IF(T10Margin2="""",1,T10Margin2),IF(T10Margin3="""",2,T10Margin3),IF(T10Margin4="""",3,T10Margin4),IF(T10Margin5="""",4,T10Margin5),IF(T10Margin6="""",5,T10Margin6),IF(T10Margin7="""",6,T10Margin7),IF(T10Margin8="""",7,T10Margin8)),0)=0," & IIf(PCode = 0, 0, Prd.getRCVMargin(pdteFrom, pdteTo)) & ",IFERROR(MODE(IF(T10Margin1="""",0,T10Margin1),IF(T10Margin2="""",1,T10Margin2),IF(T10Margin3="""",2,T10Margin3),IF(T10Margin4="""",3,T10Margin4),IF(T10Margin5="""",4,T10Margin5),IF(T10Margin6="""",5,T10Margin6),IF(T10Margin7="""",6,T10Margin7),IF(T10Margin8="""",7,T10Margin8)),0))"
        Case "ProdContri$-S/W"
            f_KeyMetrics = "=IF(SUM(IF(T10ContSW1="""",0,T10ContSW1),IF(T10ContSW2="""",0,T10ContSW2),IF(T10ContSW3="""",0,T10ContSW3),IF(T10ContSW4="""",0,T10ContSW4),IF(T10ContSW5="""",0,T10ContSW5),IF(T10ContSW6="""",0,T10ContSW6),IF(T10ContSW7="""",0,T10ContSW7),IF(T10ContSW8="""",0,T10ContSW8)>0),SUM(IF(T10ContSW1="""",0,T10ContSW1),IF(T10ContSW2="""",0,T10ContSW2),IF(T10ContSW3="""",0,T10ContSW3),IF(T10ContSW4="""",0,T10ContSW4),IF(T10ContSW5="""",0,T10ContSW5),IF(T10ContSW6="""",0,T10ContSW6),IF(T10ContSW7="""",0,T10ContSW7),IF(T10ContSW8="""",0,T10ContSW8))," & IIf(PCode > 0, Prd.getRCVMargin(pdteFrom, pdteTo) * Prd.getPOSdata(pdteFrom, pdteTo, False, e_POSUSWTypes.eProductLevel), 0) & ")"
        Case Else
            MsgBox "f_KeyMetrics " & sName & " not found"
    End Select
Exit_Routine:
    On Error Resume Next
    Exit Function
Err_Routine:
''    If Err.Number = 91 Then f_KeyMetrics = "": GoTo Exit_Routine
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-f_KeyMetrics", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_CompetitionAnalysis(sSource As String, sName As String, sDefaultValue As String, lRow As Long, lcol As Long, Optional ByVal lGrp_Idx As Long = -1) As String
    Dim Prd As cCBA_Prod, bfound As Boolean
    Dim Val1 As Variant, Val2 As Variant, Val3 As Variant
    Dim PrdGrp As cCBA_ProdGroup
    Dim PCode As Long
    Dim a As Long, b As Long, c As Long
    Dim COM As CBA_COM_COMMatch
    Dim a_COM As Variant, m As Variant
''    Dim bFound As Boolean
    
    On Error GoTo Err_Routine
    'f_CompetitionAnalysis = sDefaultValue
    ' If there is already data for this item then skip it
    If plTH_ID > 0 Then
        f_CompetitionAnalysis = sDefaultValue
        Exit Function
    End If
    If PrdGrp Is Nothing Or lGrp_Idx = 0 Then Set PrdGrp = pclsProdGrp
    PCode = NZ(Get_TD_LV(plLV_ID, "LV_Product_Code"), 0)
    If PCode = 0 Then PCode = NZ(Get_TD_LV(plLV_ID, "LV_Prior_Product_Code"), 0)
    If PrdGrp.getProductState(PCode) = True Then
        Set Prd = PrdGrp.getProdObject(PCode, True)
    Else
        'Debug.Print "Product Not Found in ProdGroup"
        If PCode > 0 And sName = "NoOfCOMRADERows" Then
            Set PrdGrp = New cCBA_ProdGroup
            PrdGrp.RunDataGeneration pdteFrom, pdteTo, False, Nothing, "BP", PCode
            Set Prd = PrdGrp.getProdObject(PCode, True)
        ElseIf PCode = 0 Then
            If pclsSimProd Is Nothing Then
                Set Prd = pclsSimProd
            Else
                Set PrdGrp = New cCBA_ProdGroup
                PrdGrp.RunDataGeneration pdteFrom, pdteTo, False, Nothing, "BP", PCode
                Set Prd = PrdGrp.getProdObject(PCode, True)
            End If
        End If
    End If
    If lGrp_Idx = 1 And lcol = 1 Then Set psdCOM_Idx = New Scripting.Dictionary
    If psdCOM_Idx.Exists(CStr(lGrp_Idx)) = False Then
        If Not Prd Is Nothing Then
            If Not Prd.colComrade Is Nothing Then
                If Prd.colComrade.Count > 0 Then
        '''''''''''''''''DeBug Code''''''''''''''''''''''''
        '                For Each COM In Prd.Comrade
        '                    Debug.Print COM.MatchType
        '                    Debug.Print COM.CompCode
        '                Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''
                    For Each COM In Prd.colComrade
                        bfound = False
                        For Each m In psdCOM_Idx
                            If psdCOM_Idx(m).AldiPCode = COM.AldiPCode And psdCOM_Idx(m).MatchType = COM.MatchType Then bfound = True: Exit For
                        Next
                        If bfound = False Then psdCOM_Idx.Add CStr(lGrp_Idx), COM: Exit For
                    Next
                End If
            End If
        End If
    End If
    
    If psdCOM_Idx.Exists(CStr(lGrp_Idx)) = True Then
        Set COM = psdCOM_Idx(CStr(lGrp_Idx))
        Select Case sName
            Case "Retailer"
                f_CompetitionAnalysis = COM.Competitor
            Case "Prod"
                f_CompetitionAnalysis = COM.CompProdName
            Case "COMMatchType"
                f_CompetitionAnalysis = CCM_Mapping.MatchType(COM.MatchType).CompetitorLng & " " & CCM_Mapping.MatchType(COM.MatchType).Description
            Case "ShelfRetail"
                f_CompetitionAnalysis = COM.Pricedata(pdteTo, "shelf", "national", True)
            Case "Pro-RataRetail"
                f_CompetitionAnalysis = COM.Pricedata(pdteTo, "nonpromoprorata", "national", True)
            Case "Pro-RataPromoRetail"
                f_CompetitionAnalysis = COM.Pricedata(pdteTo, "prorata", "national", True)
            Case "FreqofPromo"
                a_COM = COM.RetailsArray()
                f_CompetitionAnalysis = 0
                If a_COM(1, 1) = 0 Then
''                    f_CompetitionAnalysis = 0
                Else
                    For a = LBound(a_COM, 2) To UBound(a_COM, 2)
                        If LCase(a_COM(2, a)) = "national" And a_COM(1, a) >= DateAdd("M", -3, pdteTo) Then
                            Val2 = Val2 + 1
                            If a_COM(9, a) = True Then Val1 = Val1 + 1
                        End If
                    Next
                    If NZ(Val2, 0) > 0 Then f_CompetitionAnalysis = Val1 / Val2
                End If
            Case "AveShelfRetail"
                a_COM = COM.RetailsArray()
                f_CompetitionAnalysis = 0
                If a_COM(1, 1) = 0 Then
''                    f_CompetitionAnalysis = 0
                Else
                    For a = LBound(a_COM, 2) To UBound(a_COM, 2)
                        If LCase(a_COM(2, a)) = "national" And a_COM(1, a) >= DateAdd("M", -3, pdteTo) Then
                            Val2 = Val2 + 1: Val1 = Val1 + a_COM(3, a)
                        End If
                    Next
                    If NZ(Val2, 0) > 0 Then f_CompetitionAnalysis = Val1 / Val2
                End If
            Case "NoOfCOMRADERows"
                On Error Resume Next
                f_CompetitionAnalysis = 0
                If Not Prd Is Nothing Then f_CompetitionAnalysis = Prd.colComrade.Count
                
''
''            Case "FreqofPromo"
''                a_COM = COM.RetailsArray()
''                For a = LBound(a_COM, 2) To UBound(a_COM, 2)
''                    If LCase(a_COM(2, a)) = "national" And a_COM(1, a) >= DateAdd("M", -3, pdteTo) Then
''                        val2 = val2 + 1
''                        If a_COM(9, a) = True Then val1 = val1 + 1
''                    End If
''                Next
''                f_CompetitionAnalysis = val1 / val2
''            Case "AveShelfRetail"
''                a_COM = COM.RetailsArray()
''                For a = LBound(a_COM, 2) To UBound(a_COM, 2)
''                    If LCase(a_COM(2, a)) = "national" And a_COM(1, a) >= DateAdd("M", -3, pdteTo) Then
''                        val2 = val2 + 1: val1 = val1 + a_COM(3, a)
''                    End If
''                Next
''                f_CompetitionAnalysis = val1 / val2
            Case "PriceBenchmarkProd"
                If InStr(1, COM.MatchType, "PL") > 0 Then f_CompetitionAnalysis = "Yes" Else f_CompetitionAnalysis = "No"
            Case "ShelfRetail%", "Pro-RataRetail%", "Pro-RataPromoRetail%", "AveShelfRetail%", "FreqofPromo%"
                f_CompetitionAnalysis = "=IFERROR(-(R[-1]C-MAX(RetGrp))/R[-1]C,"")"
            Case "X"
                f_CompetitionAnalysis = sDefaultValue
            Case Else
                MsgBox "f_CompetitionAnalysis " & sName & " not found"
        End Select
    Else
        f_CompetitionAnalysis = ""
    End If
Exit_Routine:
        On Error Resume Next
        Exit Function
Err_Routine:
'''    If Err.Number = 91 Or Err.Number = 49 Then f_CompetitionAnalysis = "": GoTo Exit_Routine
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-f_CompetitionAnalysis", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_SimilarCoreRangeProd1(sSource As String, sName As String, sDefaultValue As String, lRow As Long, lcol As Long, Optional ByVal lGrp_Idx As Long = -1) As String
Dim PCode As Long, DivToGen As Long
Dim Prd As cCBA_Prod
Dim arr As Variant, iarr As Variant, barr As Variant, parr As Variant
Dim PrdGrp As cCBA_ProdGroup
Dim a As Long
Dim bfound As Boolean
Dim m As Variant
Dim PY As Single, CY As Single
Dim ActiveContract(501 To 509) As Long
Dim maxDelDate As Date
Dim Ret As Single, Mar As Single
    On Error GoTo Err_Routine
    ' If there is already data for this item then skip it
    If plTH_ID > 0 Then
        f_SimilarCoreRangeProd1 = sDefaultValue
        Exit Function
    End If
    
    'f_SimilarCoreRangeProd1 = sDefaultValue
    If Not pclsSimProd Is Nothing Then PCode = pclsSimProd.lPcode
    If PCode = 0 Then PCode = NZ(Get_TD_LV(plLV_ID, "LV_Product_Code"), 0)
    If PCode = 0 Then PCode = NZ(Get_TD_LV(plLV_ID, "LV_Prior_Product_Code"), 0)
    'If Not pclsPFV Is Nothing Then If pclsPFV.CompPCode <> "" Then PCode = g_Val(pclsPFV.CompPCode)
    If PrdGrp Is Nothing Or lGrp_Idx = 0 Then Set PrdGrp = pclsProdGrp
    If Not pclsSimProd Is Nothing Then
        Set Prd = pclsSimProd
    ElseIf PrdGrp.getProductState(PCode) = True Then
        Set Prd = PrdGrp.getProdObject(PCode, True)
    Else
        'Debug.Print "Product Not Found in ProdGroup"
        If PCode > 0 Then
            Set PrdGrp = New cCBA_ProdGroup
            PrdGrp.RunDataGeneration pdteFrom, pdteTo, False, Nothing, "BP", PCode
            Set Prd = PrdGrp.getProdObject(PCode, True)
        End If
    End If
    If psdCont_Idx Is Nothing Or lGrp_Idx = 1 Then
        Set psdCont_Idx = New Scripting.Dictionary
    End If
    If lGrp_Idx > 0 Then
        If lGrp_Idx < NUM_REGIONS Then DivToGen = CLng("50" & lGrp_Idx) Else DivToGen = 509
        If psdCont_Idx.Exists(CStr(lGrp_Idx)) Then
            arr = psdCont_Idx(CStr(lGrp_Idx))
        Else
            If Not Prd Is Nothing Then
                arr = Prd.getIncoTermData("CVtS", , , , DivToGen)
                If arr(0, 0) <> 0 Then
                    For a = LBound(arr, 2) To UBound(arr, 2)
                        If arr(2, a) = True Then
                            If arr(1, a) > maxDelDate Then maxDelDate = arr(1, a)
                        End If
                    Next
                    For a = LBound(arr, 2) To UBound(arr, 2)
                        If arr(1, a) = maxDelDate Then
                            ActiveContract(DivToGen) = arr(0, a)
                            Exit For
                        End If
                    Next
                End If
                arr = Prd.getContractData("CRDfSnTx", , , ActiveContract(DivToGen), DivToGen)
                psdCont_Idx.Add CStr(lGrp_Idx), arr
            Else
                ReDim arr(0 To 0, 0 To 0)
                arr(0, 0) = 0
            End If
        End If
        If arr(0, 0) <> 0 Then
            If psdCont_Idx.Exists("b" & CStr(lGrp_Idx)) Then
                barr = psdCont_Idx("b" & CStr(lGrp_Idx))
            Else
                barr = Prd.getBraketData("VfDEFTFr", , , ActiveContract(DivToGen), DivToGen)
                psdCont_Idx.Add "b" & CStr(lGrp_Idx), barr
            End If
            If psdCont_Idx.Exists("i" & CStr(lGrp_Idx)) Then
                iarr = psdCont_Idx("i" & CStr(lGrp_Idx))
            Else
                iarr = Prd.getIncoTermData("VfI", , , ActiveContract(DivToGen), DivToGen)
                psdCont_Idx.Add "i" & CStr(lGrp_Idx), iarr
            End If
            If psdCont_Idx.Exists("p" & CStr(lGrp_Idx)) Then
                parr = psdCont_Idx("p" & CStr(lGrp_Idx))
            Else
                parr = Prd.getPriceData("VfP", pdteTo, , DivToGen)
                psdCont_Idx.Add "p" & CStr(lGrp_Idx), parr
            End If
        End If
    End If
    Select Case sName
        Case "SimilarCoreRangeProd"
            f_SimilarCoreRangeProd1 = CStr(PCode)
        Case "ProdDesc"
            If Not Prd Is Nothing Then f_SimilarCoreRangeProd1 = Prd.sProdDesc
        Case "GST"
            If Not Prd Is Nothing Then If Prd.lTaxID = 2 Then f_SimilarCoreRangeProd1 = "Yes" Else f_SimilarCoreRangeProd1 = "No"
        Case "Share"
            If Not Prd Is Nothing Then f_SimilarCoreRangeProd1 = (Prd.getRCVShare(pdteFrom, pdteTo, True) * 100)
        Case "Supplier"
            If arr(0, 0) = 0 Then f_SimilarCoreRangeProd1 = "" Else f_SimilarCoreRangeProd1 = arr(3, UBound(arr, 2) - 1)
        Case "Regions(All)"
            If arr(0, 0) = 0 Then f_SimilarCoreRangeProd1 = "" Else f_SimilarCoreRangeProd1 = CBA_BasicFunctions.CBA_DivtoReg(arr(1, UBound(arr, 2)))
        Case "Incoterm"
            If arr(0, 0) = 0 Then
                f_SimilarCoreRangeProd1 = ""
            Else
                If iarr(0, 0) = 0 Then
                    f_SimilarCoreRangeProd1 = ""
                Else
                    f_SimilarCoreRangeProd1 = iarr(1, UBound(iarr, 2) - 1)
                End If
            End If
        Case "DelivCost"
            If arr(0, 0) = 0 Then
                f_SimilarCoreRangeProd1 = ""
            Else
                If barr(0, 0) = 0 Then
                    f_SimilarCoreRangeProd1 = ""
                ElseIf barr(2, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd1 = barr(2, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                ElseIf barr(3, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd1 = barr(3, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                ElseIf barr(4, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd1 = barr(4, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                Else
                    f_SimilarCoreRangeProd1 = barr(1, UBound(barr, 2) - 1)
                End If
            End If
        Case "Retail"
            If arr(0, 0) = 0 Then
                f_SimilarCoreRangeProd1 = ""
            Else
                If parr(0, 0) = 0 Then f_SimilarCoreRangeProd1 = "" Else f_SimilarCoreRangeProd1 = parr(1, UBound(parr, 2) - 1)
            End If
        Case "CaseSize"
            If Not Prd Is Nothing Then f_SimilarCoreRangeProd1 = Prd.lPackSize
        Case "Margin"
            If IsEmpty(arr) Then
                f_SimilarCoreRangeProd1 = ""
            ElseIf arr(0, 0) = 0 Then
                f_SimilarCoreRangeProd1 = ""
            Else
                If barr(0, 0) = 0 Then
                    f_SimilarCoreRangeProd1 = ""
                ElseIf barr(2, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd1 = barr(2, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                ElseIf barr(3, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd1 = barr(3, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                ElseIf barr(4, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd1 = barr(4, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                Else
                    f_SimilarCoreRangeProd1 = barr(1, UBound(barr, 2) - 1)
                End If
            End If
            If IsEmpty(arr) Then
                Ret = 0
            ElseIf arr(0, 0) = 0 Then
                Ret = 0
            Else
                If parr(0, 0) = 0 Then Ret = 0 Else Ret = parr(1, UBound(parr, 2) - 1)
            End If
            If Prd Is Nothing Then
                f_SimilarCoreRangeProd1 = 0
            ElseIf Prd.lTaxID = 2 Then
                If Ret = 0 Or f_SimilarCoreRangeProd1 = "" Then f_SimilarCoreRangeProd1 = 0 Else f_SimilarCoreRangeProd1 = ((Ret / 1.1) - (f_SimilarCoreRangeProd1 / Prd.lPackSize)) / Ret
            Else
                If Ret = 0 Or f_SimilarCoreRangeProd1 = "" Then f_SimilarCoreRangeProd1 = 0 Else f_SimilarCoreRangeProd1 = (Ret - (f_SimilarCoreRangeProd1 / Prd.lPackSize)) / Ret
            End If
'            If Not Prd Is Nothing Then f_SimilarCoreRangeProd1 = Prd.getRCVMargin(pdteFrom, pdteTo)
        Case "USW"
            If Not Prd Is Nothing Then f_SimilarCoreRangeProd1 = Prd.getPOSdata(pdteFrom, pdteTo, True, e_POSUSWTypes.eProductLevel, , , DivToGen)
        Case "Contr.$/S/W"
            If IsEmpty(arr) Then
                f_SimilarCoreRangeProd1 = ""
            ElseIf arr(0, 0) = 0 Then
                f_SimilarCoreRangeProd1 = ""
            Else
                If barr(0, 0) = 0 Then
                    f_SimilarCoreRangeProd1 = ""
                ElseIf barr(2, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd1 = barr(2, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                ElseIf barr(3, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd1 = barr(3, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                ElseIf barr(4, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd1 = barr(4, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                Else
                    f_SimilarCoreRangeProd1 = barr(1, UBound(barr, 2) - 1)
                End If
            End If
            If IsEmpty(arr) Then
                Ret = 0
            ElseIf arr(0, 0) = 0 Then
                Ret = 0
            Else
                If parr(0, 0) = 0 Then Ret = 0 Else Ret = parr(1, UBound(parr, 2) - 1)
            End If
            If Prd Is Nothing Then
                Mar = 0
            ElseIf Prd.lTaxID = 2 Then
                If Ret = 0 Or f_SimilarCoreRangeProd1 = "" Then f_SimilarCoreRangeProd1 = 0 Else Mar = ((Ret / 1.1) - (f_SimilarCoreRangeProd1 / Prd.lPackSize)) / Ret
            Else
                If Ret = 0 Or f_SimilarCoreRangeProd1 = "" Then f_SimilarCoreRangeProd1 = 0 Else Mar = (Ret - (f_SimilarCoreRangeProd1 / Prd.lPackSize)) / Ret
            End If
            If Not Prd Is Nothing Then f_SimilarCoreRangeProd1 = Mar * Ret * Prd.getPOSdata(pdteFrom, pdteTo, True, e_POSUSWTypes.eProductLevel, , , DivToGen)
'            If Not Prd Is Nothing Then f_SimilarCoreRangeProd1 = Prd.getPOSdata(pdteFrom, pdteTo, False, e_POSUSWTypes.eProductLevel) * Prd.getRCVMargin(pdteFrom, pdteTo)        Case "Vol%+/-"
        Case "Vol%+/-"
            If Not Prd Is Nothing Then
                PY = Prd.getPOSdata(DateAdd("YYYY", -1, pdteFrom), DateAdd("YYYY", -1, pdteTo), True, e_POSUSWTypes.eNotUSW, , , DivToGen)
                CY = Prd.getPOSdata(pdteFrom, pdteTo, True, e_POSUSWTypes.eNotUSW, , , DivToGen)
                If PY = 0 Then f_SimilarCoreRangeProd1 = "N/A" Else f_SimilarCoreRangeProd1 = (CY - PY) / PY
            End If
        Case "$Val%+/-"
            If Not Prd Is Nothing Then
                PY = Prd.getPOSdata(DateAdd("YYYY", -1, pdteFrom), DateAdd("YYYY", -1, pdteTo), False, e_POSUSWTypes.eNotUSW, , , DivToGen)
                CY = Prd.getPOSdata(pdteFrom, pdteTo, False, e_POSUSWTypes.eNotUSW, , , DivToGen)
                If PY = 0 Then f_SimilarCoreRangeProd1 = "N/A" Else f_SimilarCoreRangeProd1 = (CY - PY) / PY
            End If

''        Case "Vol%+/-"
''            If Not PrdGrp Is Nothing Then
''                If Not Prd Is Nothing Then
''                    PY = PrdGrp.getUSW(eUSWP, pdteFrom, pdteTo, eQty)
''                    CY = Prd.getPOSdata(pdteFrom, pdteTo, True, ProductLevel)
''                    If PY = 0 Then f_SimilarCoreRangeProd1 = "N/A" Else f_SimilarCoreRangeProd1 = (CY - PY) / PY
''                Else
''                    f_SimilarCoreRangeProd1 = "N/A"
''                End If
''            Else
''                f_SimilarCoreRangeProd1 = "N/A"
''            End If
''        Case "$Val%+/-"
''            If Not PrdGrp Is Nothing Then
''                If Not Prd Is Nothing Then
''                    PY = PrdGrp.getRCVMargin(eRCVMarginPercent, pdteFrom, pdteTo) * PrdGrp.getUSW(eUSWP, pdteFrom, pdteTo, eRetail)
''                    CY = Prd.getPOSdata(pdteFrom, pdteTo, False, ProductLevel) * Prd.getRCVMargin(pdteFrom, pdteTo)
''                    If PY = 0 Then f_SimilarCoreRangeProd1 = "N/A" Else f_SimilarCoreRangeProd1 = (CY - PY) / PY
''                Else
''                    f_SimilarCoreRangeProd1 = "N/A"
''                End If
''            Else
''                f_SimilarCoreRangeProd1 = "N/A"
''            End If

        Case Else
            MsgBox "f_SimilarCoreRangeProd1 " & sName & " not found"
    End Select
Exit_Routine:
    On Error Resume Next
    Exit Function
Err_Routine:
''    If Err.Number = 91 Then f_SimilarCoreRangeProd1 = "": GoTo Exit_Routine
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-f_SimilarCoreRangeProd1", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_SimilarCoreRangeProd3(sSource As String, sName As String, sDefaultValue As String, lRow As Long, lcol As Long, Optional ByVal lGrp_Idx As Long = -1) As String
Dim PCode As Long, DivToGen As Long
Dim Prd As cCBA_Prod
Dim arr As Variant, iarr As Variant, barr As Variant, parr As Variant
Dim PrdGrp As cCBA_ProdGroup
Dim a As Long
Dim bfound As Boolean
Dim m As Variant, v As Variant
Dim PY As Single, CY As Single
Dim ActiveContract(501 To 509) As Long
Dim maxDelDate As Date
Dim SD As Scripting.Dictionary
Dim Ret As Single, Mar As Single
    On Error GoTo Err_Routine
    ' If there is already data for this item then skip it
    If plTH_ID > 0 Then
        f_SimilarCoreRangeProd3 = sDefaultValue
        Exit Function
    End If

    If Not pclsPFV Is Nothing Then If pclsPFV.CompPCode <> "" Then PCode = g_Val(pclsPFV.CompPCode)
    If PCode = 0 Then PCode = NZ(Get_TD_LV(plLV_ID, "LV_Product_Code"), 0)
    If PCode = 0 Then PCode = NZ(Get_TD_LV(plLV_ID, "LV_Prior_Product_Code"), 0)
    If PrdGrp Is Nothing Or lGrp_Idx = 0 Then Set PrdGrp = pclsProdGrp
    If Not pclsSimProd Is Nothing Then
        Set Prd = pclsSimProd
    ElseIf PrdGrp.getProductState(PCode) = True Then
        Set Prd = PrdGrp.getProdObject(PCode, True)
    Else
        'Debug.Print "Product Not Found in ProdGroup"
        If PCode > 0 Then
            Set PrdGrp = New cCBA_ProdGroup
            PrdGrp.RunDataGeneration pdteFrom, pdteTo, False, Nothing, "BP", PCode
            Set Prd = PrdGrp.getProdObject(PCode, True)
        End If
    End If
    If psdCont_Idx Is Nothing Or lGrp_Idx = 1 Then
       Set psdCont_Idx = New Scripting.Dictionary
    End If
'    If lGrp_Idx > 0 Then
        If lGrp_Idx < NUM_REGIONS Then DivToGen = CLng("50" & lGrp_Idx + 1) Else DivToGen = 509
        If psdCont_Idx.Exists(CStr(lGrp_Idx)) Then
            arr = psdCont_Idx(CStr(lGrp_Idx))
        Else
            If Not Prd Is Nothing Then
                arr = Prd.getIncoTermData("RCVtS")
                If arr(0, 0) <> 0 Then
                    Set SD = New Scripting.Dictionary
                    For a = LBound(arr, 2) To UBound(arr, 2)
                        If SD.Exists(arr(1, a)) Then SD(arr(1, a)) = SD(arr(1, a)) + 1 Else SD.Add arr(1, a), 1
                    Next
                    For Each v In SD
                        If ActiveContract(DivToGen) = 0 Then
                            ActiveContract(DivToGen) = CLng(v)
                        Else
                            If SD(v) > SD(ActiveContract(DivToGen)) Then ActiveContract(DivToGen) = CStr(v)
                        End If
                    Next
                End If

'                arr = Prd.getIncoTermData("CVtS", , , , DivToGen)
'                If arr(0, 0) <> 0 Then
'                    For a = LBound(arr, 2) To UBound(arr, 2)
'                        If arr(2, a) = 1 Then
'                            If arr(1, a) > maxDelDate Then maxDelDate = arr(1, a)
'                        End If
'                    Next
'                    For a = LBound(arr, 2) To UBound(arr, 2)
'                        If arr(1, a) = maxDelDate Then
'                            ActiveContract(DivToGen) = arr(0, a)
'                            Exit For
'                        End If
'                    Next
'                End If
                arr = Prd.getContractData("CRDfSnTx", , , ActiveContract(DivToGen), DivToGen)
                psdCont_Idx.Add CStr(lGrp_Idx), arr
            End If
        End If
        If IsEmpty(arr) = False Then
            If arr(0, 0) <> 0 Then
                If psdCont_Idx.Exists("b" & CStr(lGrp_Idx)) Then
                    barr = psdCont_Idx("b" & CStr(lGrp_Idx))
                Else
                    If Not Prd Is Nothing Then
                        barr = Prd.getBraketData("VfDEFTFr", , , ActiveContract(DivToGen), DivToGen)
                        psdCont_Idx.Add "b" & CStr(lGrp_Idx), barr
                    End If
                End If
                If psdCont_Idx.Exists("i" & CStr(lGrp_Idx)) Then
                    iarr = psdCont_Idx("i" & CStr(lGrp_Idx))
                Else
                    If Not Prd Is Nothing Then
                        iarr = Prd.getIncoTermData("VfI", , , ActiveContract(DivToGen), DivToGen)
                        psdCont_Idx.Add "i" & CStr(lGrp_Idx), iarr
                    End If
                End If
                If psdCont_Idx.Exists("p" & CStr(lGrp_Idx)) Then
                    parr = psdCont_Idx("p" & CStr(lGrp_Idx))
                Else
                    If Not Prd Is Nothing Then
                        parr = Prd.getPriceData("VfP", pdteTo, , DivToGen)
                        psdCont_Idx.Add "p" & CStr(lGrp_Idx), parr
                    End If
                End If
            End If
        End If
'    End If
    Select Case sName
        Case "SimilarCoreRangeProd"
            f_SimilarCoreRangeProd3 = CStr(PCode)
        Case "Suppliers"
            If IsEmpty(arr) Then
                f_SimilarCoreRangeProd3 = ""
            Else
                If arr(0, 0) = 0 Then f_SimilarCoreRangeProd3 = "" Else f_SimilarCoreRangeProd3 = arr(3, UBound(arr, 2) - 1)
            End If
        Case "NatDelivCost"
            If IsEmpty(arr) Then
                f_SimilarCoreRangeProd3 = ""
            ElseIf arr(0, 0) = 0 Then
                f_SimilarCoreRangeProd3 = ""
            Else
                If barr(0, 0) = 0 Then
                    f_SimilarCoreRangeProd3 = ""
                ElseIf barr(2, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd3 = barr(2, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                ElseIf barr(3, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd3 = barr(3, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                ElseIf barr(4, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd3 = barr(4, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                Else
                    f_SimilarCoreRangeProd3 = barr(1, UBound(barr, 2) - 1)
                End If
            End If
        Case "Retail"
            If IsEmpty(arr) Then
                f_SimilarCoreRangeProd3 = ""
            ElseIf arr(0, 0) = 0 Then
                f_SimilarCoreRangeProd3 = ""
            Else
                If parr(0, 0) = 0 Then f_SimilarCoreRangeProd3 = "" Else f_SimilarCoreRangeProd3 = parr(1, UBound(parr, 2) - 1)
            End If
        Case "GST"
            If Prd Is Nothing Then
                f_SimilarCoreRangeProd3 = "Yes"
            ElseIf Prd.lTaxID = 2 Then
                f_SimilarCoreRangeProd3 = "Yes"
            Else
                f_SimilarCoreRangeProd3 = "No"
            End If
        Case "USW"
            If Not Prd Is Nothing Then f_SimilarCoreRangeProd3 = Prd.getPOSdata(pdteFrom, pdteTo, True, e_POSUSWTypes.eProductLevel, , , DivToGen)
        Case "Contr.$/S/W"
            If IsEmpty(arr) Then
                f_SimilarCoreRangeProd3 = ""
            ElseIf arr(0, 0) = 0 Then
                f_SimilarCoreRangeProd3 = ""
            Else
                If barr(0, 0) = 0 Then
                    f_SimilarCoreRangeProd3 = ""
                ElseIf barr(2, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd3 = barr(2, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                ElseIf barr(3, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd3 = barr(3, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                ElseIf barr(4, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd3 = barr(4, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                Else
                    f_SimilarCoreRangeProd3 = barr(1, UBound(barr, 2) - 1)
                End If
            End If
            If IsEmpty(arr) Then
                Ret = 0
            ElseIf arr(0, 0) = 0 Then
                Ret = 0
            Else
                If parr(0, 0) = 0 Then Ret = 0 Else Ret = parr(1, UBound(parr, 2) - 1)
            End If
            If Prd Is Nothing Then
                Mar = 0
            ElseIf Prd.lTaxID = 2 Then
                If Ret = 0 Then f_SimilarCoreRangeProd3 = 0 Else Mar = ((Ret / 1.1) - (f_SimilarCoreRangeProd3 / Prd.lPackSize)) / Ret
            Else
                If Ret = 0 Then f_SimilarCoreRangeProd3 = 0 Else Mar = (Ret - (f_SimilarCoreRangeProd3 / Prd.lPackSize)) / Ret
            End If
            If Not Prd Is Nothing Then f_SimilarCoreRangeProd3 = Mar * Ret * Prd.getPOSdata(pdteFrom, pdteTo, True, e_POSUSWTypes.eProductLevel, , , DivToGen)
        
'            If Not Prd Is Nothing Then f_SimilarCoreRangeProd3 = Prd.getPOSdata(pdteFrom, pdteTo, False, e_POSUSWTypes.eProductLevel) * Prd.getRCVMargin(pdteFrom, pdteTo)
    Case "Margin"
            If IsEmpty(arr) Then
                f_SimilarCoreRangeProd3 = ""
            ElseIf arr(0, 0) = 0 Then
                f_SimilarCoreRangeProd3 = ""
            Else
                If barr(0, 0) = 0 Then
                    f_SimilarCoreRangeProd3 = ""
                ElseIf barr(2, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd3 = barr(2, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                ElseIf barr(3, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd3 = barr(3, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                ElseIf barr(4, UBound(barr, 2)) > 0 Then
                    f_SimilarCoreRangeProd3 = barr(4, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
                Else
                    f_SimilarCoreRangeProd3 = barr(1, UBound(barr, 2) - 1)
                End If
            End If
            If IsEmpty(arr) Then
                Ret = 0
            ElseIf arr(0, 0) = 0 Then
                Ret = 0
            Else
                If parr(0, 0) = 0 Then Ret = 0 Else Ret = parr(1, UBound(parr, 2) - 1)
            End If
            If Prd Is Nothing Then
            f_SimilarCoreRangeProd3 = 0
            ElseIf Prd.lTaxID = 2 Then
            If Ret = 0 Then f_SimilarCoreRangeProd3 = 0 Else f_SimilarCoreRangeProd3 = ((Ret / 1.1) - (f_SimilarCoreRangeProd3 / Prd.lPackSize)) / Ret
            Else
            If Ret = 0 Then f_SimilarCoreRangeProd3 = 0 Else f_SimilarCoreRangeProd3 = (Ret - (f_SimilarCoreRangeProd3 / Prd.lPackSize)) / Ret
            End If
            
            
'            If Not Prd Is Nothing Then f_SimilarCoreRangeProd3 = Prd.getRCVMargin(pdteFrom, pdteTo)
        
        Case "Share"
            If Not Prd Is Nothing Then f_SimilarCoreRangeProd3 = Prd.getRCVShare(pdteFrom, pdteTo, True)
        Case "Vol%+/-"
            If Not Prd Is Nothing Then
                PY = Prd.getPOSdata(DateAdd("YYYY", -1, pdteFrom), DateAdd("YYYY", -1, pdteTo), True, e_POSUSWTypes.eNotUSW, , , DivToGen)
                CY = Prd.getPOSdata(pdteFrom, pdteTo, True, e_POSUSWTypes.eNotUSW, , , DivToGen)
                If PY = 0 Then f_SimilarCoreRangeProd3 = "N/A" Else f_SimilarCoreRangeProd3 = (CY - PY) / PY
            End If
        Case "$Val%+/-"
            If Not Prd Is Nothing Then
                PY = Prd.getPOSdata(DateAdd("YYYY", -1, pdteFrom), DateAdd("YYYY", -1, pdteTo), False, e_POSUSWTypes.eNotUSW, , , DivToGen)
                CY = Prd.getPOSdata(pdteFrom, pdteTo, False, e_POSUSWTypes.eNotUSW, , , DivToGen)
                If PY = 0 Then f_SimilarCoreRangeProd3 = "N/A" Else f_SimilarCoreRangeProd3 = (CY - PY) / PY
            End If
        Case "Cmts"
            f_SimilarCoreRangeProd3 = ""
        Case Else
            MsgBox "f_SimilarCoreRangeProd3 " & sName & " not found"
            Stop
    End Select
Exit_Routine:
    On Error Resume Next
    Exit Function
Err_Routine:
''    If Err.Number = 91 Then f_SimilarCoreRangeProd3 = "": GoTo Exit_Routine
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-f_SimilarCoreRangeProd3", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function


''Private Function f_SimilarCoreRangeProd1(sSource As String, sName As String, sDefaultValue As String, lRow As Long, lCol As Long, Optional ByVal lGrp_Idx As Long = -1) As String
''Dim PCode As Long, DivToGen As Long
''Dim Prd As cCBA_Prod
''Dim arr As Variant, iarr As Variant, barr As Variant, parr As Variant
''Dim PrdGrp As cCBA_ProdGroup
''Dim a As Long
''Dim bFound As Boolean
''Dim m As Variant
''Dim PY As Single, CY As Single
''
''    On Error GoTo Err_Routine
''    ' If there is already data for this item then skip it
''    If plTH_ID > 0 Then
''        f_SimilarCoreRangeProd1 = sDefaultValue
''        Exit Function
''    End If
''
''    'f_SimilarCoreRangeProd1 = sDefaultValue
''    If PCode = 0 Then PCode = NZ(Get_TD_LV(plLV_ID, "LV_Product_Code"), 0)
''    If PCode = 0 Then PCode = NZ(Get_TD_LV(plLV_ID, "LV_Prior_Product_Code"), 0)
''    If Not pclsPFV Is Nothing Then If pclsPFV.CompPCode <> "" Then PCode = g_Val(pclsPFV.CompPCode)
''    If PrdGrp Is Nothing Or lGrp_Idx = 0 Then Set PrdGrp = pclsProdGrp
''    If Not pclsSimProd Is Nothing Then
''        Set Prd = pclsSimProd
''    ElseIf PrdGrp.getProductState(PCode) = True Then
''        Set Prd = PrdGrp.getProdObject(PCode,True)
''    Else
''        'Debug.Print "Product Not Found in ProdGroup"
''        If PCode > 0 Then
''            Set PrdGrp = New cCBA_ProdGroup
''            PrdGrp.RunDataGeneration pdteFrom, pdteTo, False, Nothing, "BP", PCode
''            Set Prd = PrdGrp.getProdObject(PCode,True)
''        End If
''    End If
''    If psdCont_Idx Is Nothing Or lGrp_Idx = 1 Then
''        Set psdCont_Idx = New Scripting.Dictionary
''    End If
''    If lGrp_Idx > 0 Then
''        If lGrp_Idx < NUM_REGIONS Then DivToGen = CLng("50" & lGrp_Idx) Else DivToGen = 509
''        If psdCont_Idx.Exists(CStr(lGrp_Idx)) Then
''            arr = psdCont_Idx(CStr(lGrp_Idx))
''        Else
''            If Not Prd Is Nothing Then
''                arr = Prd.getContractData("CRDfSnTx", , , , DivToGen)
''                psdCont_Idx.Add CStr(lGrp_Idx), arr
''            Else
''                ReDim arr(0 To 0, 0 To 0)
''                arr(0, 0) = 0
''            End If
''        End If
''        If arr(0, 0) <> 0 Then
''            If psdCont_Idx.Exists("b" & CStr(lGrp_Idx)) Then
''                barr = psdCont_Idx("b" & CStr(lGrp_Idx))
''            Else
''                barr = Prd.getBraketData("VfDEFTFr", , , CLng(arr(0, UBound(arr, 2) - 1)), DivToGen)
''                psdCont_Idx.Add "b" & CStr(lGrp_Idx), barr
''            End If
''            If psdCont_Idx.Exists("i" & CStr(lGrp_Idx)) Then
''                iarr = psdCont_Idx("i" & CStr(lGrp_Idx))
''            Else
''                iarr = Prd.getIncoTermData("VfI", , , CLng(arr(0, UBound(arr, 2) - 1)), DivToGen)
''                psdCont_Idx.Add "i" & CStr(lGrp_Idx), iarr
''            End If
''            If psdCont_Idx.Exists("p" & CStr(lGrp_Idx)) Then
''                parr = psdCont_Idx("p" & CStr(lGrp_Idx))
''            Else
''                parr = Prd.getPriceData("VfP", pdteTo, , DivToGen)
''                psdCont_Idx.Add "p" & CStr(lGrp_Idx), parr
''            End If
''        End If
''    End If
''    Select Case sName
''        Case "SimilarCoreRangeProd"
''            f_SimilarCoreRangeProd1 = CStr(PCode)
''        Case "GST"
''            If Not Prd Is Nothing Then If Prd.TaxID = 2 Then f_SimilarCoreRangeProd1 = "Yes" Else f_SimilarCoreRangeProd1 = "No"
''        Case "Share"
''            If Not Prd Is Nothing Then f_SimilarCoreRangeProd1 = (Prd.getRCVShare(pdteFrom, pdteTo, True) * 100)
''        Case "Supplier"
''            If arr(0, 0) = 0 Then f_SimilarCoreRangeProd1 = "" Else f_SimilarCoreRangeProd1 = arr(3, UBound(arr, 2) - 1)
''        Case "Regions(All)"
''            If arr(0, 0) = 0 Then f_SimilarCoreRangeProd1 = "" Else f_SimilarCoreRangeProd1 = CBA_BasicFunctions.CBA_DivtoReg(arr(1, UBound(arr, 2)))
''        Case "Incoterm"
''            If arr(0, 0) = 0 Then
''                f_SimilarCoreRangeProd1 = ""
''            Else
''                If iarr(0, 0) = 0 Then
''                    f_SimilarCoreRangeProd1 = ""
''                Else
''                    f_SimilarCoreRangeProd1 = iarr(1, UBound(iarr, 2) - 1)
''                End If
''            End If
''        Case "DelivCost"
''            If arr(0, 0) = 0 Then
''                f_SimilarCoreRangeProd1 = ""
''            Else
''                If barr(0, 0) = 0 Then
''                    f_SimilarCoreRangeProd1 = ""
''                ElseIf barr(2, UBound(barr, 2)) > 0 Then
''                    f_SimilarCoreRangeProd1 = barr(2, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
''                ElseIf barr(3, UBound(barr, 2)) > 0 Then
''                    f_SimilarCoreRangeProd1 = barr(3, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
''                ElseIf barr(4, UBound(barr, 2)) > 0 Then
''                    f_SimilarCoreRangeProd1 = barr(4, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
''                Else
''                    f_SimilarCoreRangeProd1 = barr(1, UBound(barr, 2) - 1)
''                End If
''            End If
''        Case "Retail"
''            If arr(0, 0) = 0 Then
''                f_SimilarCoreRangeProd1 = ""
''            Else
''                If parr(0, 0) = 0 Then f_SimilarCoreRangeProd1 = "" Else f_SimilarCoreRangeProd1 = parr(1, UBound(parr, 2) - 1)
''            End If
''        Case "CaseSize"
''            If Not Prd Is Nothing Then f_SimilarCoreRangeProd1 = Prd.Packsize
''        Case "Margin"
''            If Not Prd Is Nothing Then f_SimilarCoreRangeProd1 = Prd.getRCVMargin(pdteFrom, pdteTo)
''        Case "USW"
''            If Not Prd Is Nothing Then f_SimilarCoreRangeProd1 = Prd.getPOSdata(pdteFrom, pdteTo, True, e_POSUSWTypes.eProductLevel)
''        Case "Contr.$/S/W"
''            If Not Prd Is Nothing Then f_SimilarCoreRangeProd1 = Prd.getPOSdata(pdteFrom, pdteTo, False, e_POSUSWTypes.eProductLevel) * Prd.getRCVMargin(pdteFrom, pdteTo)
''        Case "Vol%+/-"
''            If Not Prd Is Nothing Then
''                PY = Prd.getPOSdata(DateAdd("YYYY", -1, pdteFrom), DateAdd("YYYY", -1, pdteTo), True, e_POSUSWTypes.eNotUSW)
''                CY = Prd.getPOSdata(pdteFrom, pdteTo, True, e_POSUSWTypes.eNotUSW)
''                If PY = 0 Then f_SimilarCoreRangeProd1 = "N/A" Else f_SimilarCoreRangeProd1 = (CY - PY) / PY
''            End If
''        Case "$Val%+/-"
''            If Not Prd Is Nothing Then
''                PY = Prd.getPOSdata(DateAdd("YYYY", -1, pdteFrom), DateAdd("YYYY", -1, pdteTo), False, e_POSUSWTypes.eNotUSW)
''                CY = Prd.getPOSdata(pdteFrom, pdteTo, False, e_POSUSWTypes.eNotUSW)
''                If PY = 0 Then f_SimilarCoreRangeProd1 = "N/A" Else f_SimilarCoreRangeProd1 = (CY - PY) / PY
''            End If
''
''''        Case "Vol%+/-"
''''            If Not PrdGrp Is Nothing Then
''''                If Not Prd Is Nothing Then
''''                    PY = PrdGrp.getUSW(eUSWP, pdteFrom, pdteTo, eQty)
''''                    CY = Prd.getPOSdata(pdteFrom, pdteTo, True, ProductLevel)
''''                    If PY = 0 Then f_SimilarCoreRangeProd1 = "N/A" Else f_SimilarCoreRangeProd1 = (CY - PY) / PY
''''                Else
''''                    f_SimilarCoreRangeProd1 = "N/A"
''''                End If
''''            Else
''''                f_SimilarCoreRangeProd1 = "N/A"
''''            End If
''''        Case "$Val%+/-"
''''            If Not PrdGrp Is Nothing Then
''''                If Not Prd Is Nothing Then
''''                    PY = PrdGrp.getRCVMargin(eRCVMarginPercent, pdteFrom, pdteTo) * PrdGrp.getUSW(eUSWP, pdteFrom, pdteTo, eRetail)
''''                    CY = Prd.getPOSdata(pdteFrom, pdteTo, False, ProductLevel) * Prd.getRCVMargin(pdteFrom, pdteTo)
''''                    If PY = 0 Then f_SimilarCoreRangeProd1 = "N/A" Else f_SimilarCoreRangeProd1 = (CY - PY) / PY
''''                Else
''''                    f_SimilarCoreRangeProd1 = "N/A"
''''                End If
''''            Else
''''                f_SimilarCoreRangeProd1 = "N/A"
''''            End If
''
''        Case Else
''            MsgBox "f_SimilarCoreRangeProd1 " & sName & " not found"
''    End Select
''Exit_Routine:
''    On Error Resume Next
''    Exit Function
''Err_Routine:
''''    If Err.Number = 91 Then f_SimilarCoreRangeProd1 = "": GoTo Exit_Routine
''    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-f_SimilarCoreRangeProd1", 3)
''    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    Debug.Print CBA_Error
''    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
''    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
''    GoTo Exit_Routine
''    Resume Next
''
''End Function
''
''Private Function f_SimilarCoreRangeProd3(sSource As String, sName As String, sDefaultValue As String, lRow As Long, lCol As Long, Optional ByVal lGrp_Idx As Long = -1) As String
''Dim PCode As Long, DivToGen As Long
''Dim Prd As cCBA_Prod
''Dim arr As Variant, iarr As Variant, barr As Variant, parr As Variant
''Dim PrdGrp As cCBA_ProdGroup
''Dim a As Long
''Dim bFound As Boolean
''Dim m As Variant
''Dim PY As Single, CY As Single
''
''    On Error GoTo Err_Routine
''    ' If there is already data for this item then skip it
''    If plTH_ID > 0 Then
''        f_SimilarCoreRangeProd3 = sDefaultValue
''        Exit Function
''    End If
''
''    If Not pclsPFV Is Nothing Then If pclsPFV.CompPCode <> "" Then PCode = g_Val(pclsPFV.CompPCode)
''    If PCode = 0 Then PCode = NZ(Get_TD_LV(plLV_ID, "LV_Product_Code"), 0)
''    If PCode = 0 Then PCode = NZ(Get_TD_LV(plLV_ID, "LV_Prior_Product_Code"), 0)
''    If PrdGrp Is Nothing Or lGrp_Idx = 0 Then Set PrdGrp = pclsProdGrp
''    If Not pclsSimProd Is Nothing Then
''        Set Prd = pclsSimProd
''    ElseIf PrdGrp.getProductState(PCode) = True Then
''        Set Prd = PrdGrp.getProdObject(PCode,True)
''    Else
''        'Debug.Print "Product Not Found in ProdGroup"
''        If PCode > 0 Then
''            Set PrdGrp = New cCBA_ProdGroup
''            PrdGrp.RunDataGeneration pdteFrom, pdteTo, False, Nothing, "BP", PCode
''            Set Prd = PrdGrp.getProdObject(PCode,True)
''        End If
''    End If
''    If psdCont_Idx Is Nothing Or lGrp_Idx = 1 Then
''        Set psdCont_Idx = New Scripting.Dictionary
''    End If
'''    If lGrp_Idx > 0 Then
''        If lGrp_Idx < NUM_REGIONS Then DivToGen = CLng("50" & lGrp_Idx + 1) Else DivToGen = 509
''        If psdCont_Idx.Exists(CStr(lGrp_Idx)) Then
''            arr = psdCont_Idx(CStr(lGrp_Idx))
''        Else
''            If Not Prd Is Nothing Then
''                arr = Prd.getContractData("CRDfSnTx", , , , DivToGen)
''                psdCont_Idx.Add CStr(lGrp_Idx), arr
''            End If
''        End If
''        If IsEmpty(arr) = False Then
''            If arr(0, 0) <> 0 Then
''                If psdCont_Idx.Exists("b" & CStr(lGrp_Idx)) Then
''                    barr = psdCont_Idx("b" & CStr(lGrp_Idx))
''                Else
''                    If Not Prd Is Nothing Then
''                        barr = Prd.getBraketData("VfDEFTFr", , , CLng(arr(0, UBound(arr, 2) - 1)), DivToGen)
''                        psdCont_Idx.Add "b" & CStr(lGrp_Idx), barr
''                    End If
''                End If
''                If psdCont_Idx.Exists("i" & CStr(lGrp_Idx)) Then
''                    iarr = psdCont_Idx("i" & CStr(lGrp_Idx))
''                Else
''                    If Not Prd Is Nothing Then
''                        iarr = Prd.getIncoTermData("VfI", , , CLng(arr(0, UBound(arr, 2) - 1)), DivToGen)
''                        psdCont_Idx.Add "i" & CStr(lGrp_Idx), iarr
''                    End If
''                End If
''                If psdCont_Idx.Exists("p" & CStr(lGrp_Idx)) Then
''                    parr = psdCont_Idx("p" & CStr(lGrp_Idx))
''                Else
''                    If Not Prd Is Nothing Then
''                        parr = Prd.getPriceData("VfP", pdteTo, , DivToGen)
''                        psdCont_Idx.Add "p" & CStr(lGrp_Idx), parr
''                    End If
''                End If
''            End If
''        End If
'''    End If
''    Select Case sName
''        Case "SimilarCoreRangeProd"
''            f_SimilarCoreRangeProd3 = CStr(PCode)
''        Case "Suppliers"
''            If IsEmpty(arr) Then
''                f_SimilarCoreRangeProd3 = ""
''            Else
''                If arr(0, 0) = 0 Then f_SimilarCoreRangeProd3 = "" Else f_SimilarCoreRangeProd3 = arr(3, UBound(arr, 2) - 1)
''            End If
''        Case "NatDelivCost"
''            If IsEmpty(arr) Then
''                f_SimilarCoreRangeProd3 = ""
''            ElseIf arr(0, 0) = 0 Then
''                f_SimilarCoreRangeProd3 = ""
''            Else
''                If barr(0, 0) = 0 Then
''                    f_SimilarCoreRangeProd3 = ""
''                ElseIf barr(2, UBound(barr, 2)) > 0 Then
''                    f_SimilarCoreRangeProd3 = barr(2, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
''                ElseIf barr(3, UBound(barr, 2)) > 0 Then
''                    f_SimilarCoreRangeProd3 = barr(3, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
''                ElseIf barr(4, UBound(barr, 2)) > 0 Then
''                    f_SimilarCoreRangeProd3 = barr(4, UBound(barr, 2) - 1) + barr(5, UBound(barr, 2) - 1)
''                Else
''                    f_SimilarCoreRangeProd3 = barr(1, UBound(barr, 2) - 1)
''                End If
''            End If
''        Case "Retail"
''            If IsEmpty(arr) Then
''                f_SimilarCoreRangeProd3 = ""
''            ElseIf arr(0, 0) = 0 Then
''                f_SimilarCoreRangeProd3 = ""
''            Else
''                If parr(0, 0) = 0 Then f_SimilarCoreRangeProd3 = "" Else f_SimilarCoreRangeProd3 = parr(1, UBound(parr, 2) - 1)
''            End If
''        Case "GST"
''            If Prd Is Nothing Then
''                f_SimilarCoreRangeProd3 = "Yes"
''            ElseIf Prd.TaxID = 2 Then
''                f_SimilarCoreRangeProd3 = "Yes"
''            Else
''                f_SimilarCoreRangeProd3 = "No"
''            End If
''        Case "USW"
''            If Not Prd Is Nothing Then f_SimilarCoreRangeProd3 = Prd.getPOSdata(pdteFrom, pdteTo, True, e_POSUSWTypes.eProductLevel)
''        Case "Contr.$/S/W"
''            If Not Prd Is Nothing Then f_SimilarCoreRangeProd3 = Prd.getPOSdata(pdteFrom, pdteTo, False, e_POSUSWTypes.eProductLevel) * Prd.getRCVMargin(pdteFrom, pdteTo)
''        Case "Margin"
''            If Not Prd Is Nothing Then f_SimilarCoreRangeProd3 = Prd.getRCVMargin(pdteFrom, pdteTo)
''        Case "Share"
''            If Not Prd Is Nothing Then f_SimilarCoreRangeProd3 = Prd.getRCVShare(pdteFrom, pdteTo, True)
''        Case "Vol%+/-"
''            If Not Prd Is Nothing Then
''                PY = Prd.getPOSdata(DateAdd("YYYY", -1, pdteFrom), DateAdd("YYYY", -1, pdteTo), True, e_POSUSWTypes.eNotUSW)
''                CY = Prd.getPOSdata(pdteFrom, pdteTo, True, e_POSUSWTypes.eNotUSW)
''                If PY = 0 Then f_SimilarCoreRangeProd3 = "N/A" Else f_SimilarCoreRangeProd3 = (CY - PY) / PY
''            End If
''        Case "$Val%+/-"
''            If Not Prd Is Nothing Then
''                PY = Prd.getPOSdata(DateAdd("YYYY", -1, pdteFrom), DateAdd("YYYY", -1, pdteTo), False, e_POSUSWTypes.eNotUSW)
''                CY = Prd.getPOSdata(pdteFrom, pdteTo, False, e_POSUSWTypes.eNotUSW)
''                If PY = 0 Then f_SimilarCoreRangeProd3 = "N/A" Else f_SimilarCoreRangeProd3 = (CY - PY) / PY
''            End If
''        Case "Cmts"
''            f_SimilarCoreRangeProd3 = ""
''        Case Else
''            MsgBox "f_SimilarCoreRangeProd3 " & sName & " not found"
''            Stop
''    End Select
''Exit_Routine:
''    On Error Resume Next
''    Exit Function
''Err_Routine:
''''    If Err.Number = 91 Then f_SimilarCoreRangeProd3 = "": GoTo Exit_Routine
''    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-f_SimilarCoreRangeProd3", 3)
''    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    Debug.Print CBA_Error
''    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
''    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
''    GoTo Exit_Routine
''    Resume Next
''
''End Function

Private Function f_TenderSubmissions(ByVal sSource As String, ByVal sName As String, ByVal sDefaultValue As String, ByVal lRow As Long, ByVal lcol As Long, Optional ByVal lGrp_Idx As Long = -1) As String
    Dim sRowCol As String, lUT_Idx2 As Long, lMidVal As Long, lATPIdx As Long, lSupp As Long, bIsRegn As Boolean, bIsCost As Boolean, bAppliedVal As Boolean, sBaseName As String, sApndName As String

    On Error GoTo Err_Routine
    f_TenderSubmissions = sDefaultValue
    If plATP_Cols < 1 Then Exit Function
    If plATPLV_ID <> plLV_ID And plATPLV_ID > 0 Then Exit Function                  ' If the ATP document was selected for a different LV, then don't load it
    sRowCol = g_Fmt_2_IDs(lRow, lcol, e_UTFldFmt.eRowCol)                             ' Get the RowCol
    lUT_Idx2 = psdRC_Idx.Item(sRowCol)                                              ' Find the index for the pUT_Ary2 (tender UT array)
    lMidVal = Val(pUT_Ary2(lUT_Idx2).sUT_ID_Ext)                                    ' What is the mid val (Supp column) of the Ext value
    lSupp = g_Get_Mid_Fmt(lMidVal, 4, 2)                                            ' Supplier Column
    sBaseName = sName: sApndName = "": bIsRegn = False: bIsCost = False
    If sBaseName Like "*Trial*" Then
        sBaseName = sBaseName
    End If
    If lSupp <= plATP_Cols Then                                                     ' If the Supplier column is within the bondaries of the ATP array...
        If Val(Right(sName, 1)) > 0 Then bIsRegn = True                             ' Is it a Region Value???
        If InStr(1, sName, "Cost") > 0 Then bIsCost = True                          ' Is it a Cost Value???
        GoSub GSATPTest                                                             ' Test for a valid value
        ' If bAppliedVal is false and is Region Cost
        If bAppliedVal = False And bIsCost And bIsRegn Then                         ' If no applied cost at the regions
            sBaseName = g_Left(sBaseName, 1): sApndName = "HC"                      ' Take off the region and test for Home Currency...
            GoSub GSATPTest                                                         ' Test for a valid value
        End If
        ' If bAppliedVal still is false and is Region Cost
        If bAppliedVal = False And bIsCost And bIsRegn Then                         ' If no applied cost at the regions
            sApndName = "AUS"                                                       ' Take off the region and test for AUS...
            GoSub GSATPTest                                                         ' Test for a valid value
        End If
        ' If bAppliedVal still is false and is Region Cost
        If bAppliedVal = False And bIsCost And bIsRegn Then                         ' If no applied cost at the regions
            sApndName = ""                                                          ' Take off the region and test for ''...
            GoSub GSATPTest                                                         ' Test for a valid value
        End If
        
    End If
''    if
    
Exit_Routine:
    On Error Resume Next
    Exit Function
GSATPTest:
    If psdATPN_Idx.Exists(sBaseName & sApndName) Then
        lATPIdx = psdATPN_Idx.Item(sBaseName & sApndName)                           ' ATP Row number...
        f_TenderSubmissions = pATP_Ary(lATPIdx, lSupp).TH_Alt                       ' Applied value
        If Val(f_TenderSubmissions) > 0 Then bAppliedVal = True                     ' Cost Applied???
    End If
    Return
Err_Routine:
''    If Err.Number = 91 Then f_TenderSubmissions = "": GoTo Exit_Routine
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-f_TenderSubmissions", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_Recommendation(sSource As String, sName As String, sDefaultValue As String, ByVal lRow As Long, ByVal lcol As Long, Optional ByVal lGrp_Idx As Long = -1) As String
    On Error GoTo Err_Routine
    Dim sRowCol As String, lUT_Idx1 As Long, lUT_Idx2 As Long, lLink_ID As Long, sValue As String, lLRow As Long, lLCol As Long, lIdx As Long
    Static sdStores As New Scripting.Dictionary, bStores As Boolean, lTotStores As Long
    ' This is called from the Select and Remove buttons and basically copies appropriate data over from 'Tender Submissions'
    Dim bIsAlchohol As Boolean, v As Variant
 
    f_Recommendation = sDefaultValue
    If sName = "Supplier" Then
        GoSub GSGetLink
    ElseIf sName = "Regions" Then                           ' This and the following are for group lines
    ElseIf sName = "CostofGoods" Then
    ElseIf sName = "CaseCostofGoods" Then
    ElseIf sName = "AgreedIncoterm" Then
    ElseIf sName = "Est.Freight" Then
    ElseIf sName = "Est.DelivCost" Then
    ElseIf sName = "Retail" Then
    ElseIf sName = "Margin%" Then
    ElseIf sName = "Est.Contri-SSW" Then
    ElseIf sName = "LaunchMonth" Then                       ' This and the following are for Launch Month lines
    ElseIf sName = "CaseSize" Then
    ElseIf sName = "Trademark" Then
    ElseIf sName = "GST" Then
    ElseIf sName = "CaseCost" Then
    ElseIf sName = "Est.USW" Then
    ElseIf sName = "Est.Share" Then
    ElseIf sName = "FXRate" Then
    ElseIf sName = "RegStoreCount" Then
        If bStores = False Then
            bStores = True: lTotStores = 0
            Set sdStores = New Scripting.Dictionary
            
            If Not pclsPFV Is Nothing Then
                If pclsPFV.LCGCGno < 5 Then bIsAlchohol = True
            ElseIf Not pclsProdGrp Is Nothing Then
                If InStr(1, pclsProdGrp.sBuildCode, "As") > 0 Then bIsAlchohol = True
            ElseIf Not psdPFV Is Nothing Then
                For Each v In psdPFV
                   If psdPFV(v).LCGCGno < 5 Then bIsAlchohol = True: Exit For
                Next
            End If
            If bIsAlchohol = True Then
                If pclsProdGrp.bHasAlcoholStores = True Then
                    For lIdx = 501 To 509
                        If lIdx = 508 Then lIdx = 509
                        sdStores.Add CStr(lIdx - 500), pclsProdGrp.sdAlcoholStoreNos(Year(Date))(Month(Date))(lIdx)
                    Next
                    lTotStores = pclsProdGrp.sdAlcoholStoreNos(Year(Date))(Month(Date))(599)
                End If
            Else
                Call CBA_SQL_Queries.CBA_GenPullSQL("CBA_AST_StoreNos", g_FixDate(Date), g_FixDate(Date))
                For lIdx = 501 To 508
                    sdStores.Add CStr(lIdx - 500), CLng(CBA_CBISarr(1, lIdx - 501))
                    lTotStores = lTotStores + CLng(CBA_CBISarr(1, lIdx - 501))
                    If lIdx = 508 Then lIdx = 509
                Next
            End If
        
        End If
         f_Recommendation = sdStores.Item(CStr(lGrp_Idx))
    ElseIf sName = "TotStoreCount" Then
        ' Get from "CBA_AST_ALLStoreNos"
        f_Recommendation = lTotStores
    Else
        MsgBox "f_Recommendation " & sName & " not found"
        Exit Function
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function
GSGetLink:
    sRowCol = g_Fmt_2_IDs(lRow, lcol, e_UTFldFmt.eRowCol)
    lUT_Idx1 = psdRC_Idx.Item(sRowCol)
    lLink_ID = Val(Get_TD_UT(lUT_Idx1, "UT_Link_ID"))
    lUT_Idx2 = psdUT_Idx2.Item(CStr(lLink_ID))
    lLRow = Val(Get_TD_UT(lUT_Idx2, "UT_Pos_Top"))
    lLCol = Val(Get_TD_UT(lUT_Idx2, "UT_Pos_Left"))
    If sSource = "Select" Then
        sValue = Get_TD_UT(lUT_Idx1, "UT_Default_Value")
    Else
        sValue = ""
    End If
    Call Upd_TD_UT(lUT_Idx2, "UT_Default_Value", sValue, "UPD")
    Call g_SetupIP("TenderDocs", 1, True)
    ActiveSheet.Cells(lLRow, lLCol) = sValue
    Call g_SetupIP("TenderDocs", 1, False)
    Return
Err_Routine:
''    If Err.Number = 91 Then f_Recommendation = "": GoTo Exit_Routine
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-f_Recommendation", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_Save(sSource As String, sName As String, sDefaultValue As String, ByVal lRow As Long, ByVal lcol As Long, Optional ByVal lGrp_Idx As Long = -1) As String
    On Error GoTo Err_Routine
''    Dim sRowCol As String, lUT_Idx1 As Long, lUT_Idx2 As Long, lLink_ID As Long, sValue As String, lLRow As Long, lLCol As Long, lIdx As Long
    Dim PCode As Long, v
    If plTH_ID > 0 Then
        f_Save = sDefaultValue
        Exit Function
    End If
    ' This is called from the Save
Dim P As cCBA_Prod
    f_Save = sDefaultValue
    If sName = "TotBusNo" Then
        f_Save = psTotBus
    ElseIf sName = "HurdleRate" Then
        PCode = NZ(Get_TD_LV(plLV_ID, "LV_Product_Code"), 0)
        If PCode = 0 Then PCode = NZ(Get_TD_LV(plLV_ID, "LV_Prior_Product_Code"), 0)
        If pclsProdGrp.getProductState(PCode) Then
            If Not pclsProdGrp.getProdObject(PCode).sdPLabels Is Nothing Then
                For Each v In pclsProdGrp.getProdObject(PCode).sdPLabels
                Set P = pclsProdGrp.getProdObject(PCode)
                    If InStr(1, LCase(P.sPCrdBnd), "just organic") > 0 Or _
                        InStr(1, LCase(P.sPCrdBnd), "has no") > 0 Or _
                            InStr(1, LCase(P.sPCrdBnd), "gluten free") > 0 Then
                            f_Save = "H"
                    Else
                        If v = "Branded" Then f_Save = "B"
                    End If
                Next
            Else
                f_Save = "S"
            End If
        End If
''        ' Get from "CBA_AST_ALLStoreNos"
''        f_Save = lTotStores
    Else
        MsgBox "f_Save " & sName & " not found"
        Exit Function
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
''    If Err.Number = 91 Then f_Save = "": GoTo Exit_Routine
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-f_Save", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function f_SignOffChecklist(sSource As String, sName As String, sDefaultValue As String, lRow As Long, lcol As Long, Optional ByVal lGrp_Idx As Long = -1) As String
    On Error GoTo Err_Routine
    Dim lEmp_No As Long
    f_SignOffChecklist = sDefaultValue
    If sName = "ProdFacility" Then
    ElseIf sName = "GFSIStd" Then
    ElseIf sName = "ActionPlan" Then
    ElseIf sName = "BD" Then
        lEmp_No = Val(Get_TD_ID(plID_Idx, "ID_BD_Emp_No"))
        f_SignOffChecklist = Get_BDBA_Value(lEmp_No, "First Last")
    ElseIf sName = "GBD" Then
        lEmp_No = Val(Get_TD_ID(plID_Idx, "ID_GD_Emp_No"))
        f_SignOffChecklist = Get_BDBA_Value(lEmp_No, "First Last")
    Else
        MsgBox "f_SignOffChecklist " & sName & " not found"
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function
Err_Routine:
''    If Err.Number = 91 Then f_SignOffChecklist = "": GoTo Exit_Routine
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-f_SignOffChecklist", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Sub f_Cond_Formatting(lValue As Long, ByRef rRng As Range, ByVal lGrp_No As Long)
Dim ValForCF As Single
    On Error GoTo Err_Routine
''        ' If there is already data for this item then skip it
''    If plTH_ID > 0 Then
''''        f_Cond_Formatting = sDefaultValue
''        Exit Sub
''    End If
''
'''total busines value = ActiveSheet.Cells(1, 37).MergeArea(1, 1).Value


    rRng.Select
    Select Case lValue
         Case 0
             ValForCF = NZ(CBA_COM_ReportFunctions.CBA_COM_DiffToMaintain(CCM_Mapping.CompToFindLongDesc(ActiveSheet.Cells(rRng.Row - 2, 13).MergeArea(1, 1).Value)), 0)
         Case 1
            If rRng.Column = 12 Then
                ValForCF = TraslateHurdleToValue(ActiveSheet.Cells(1, 38).MergeArea(1, 1).Value, "USW")
            Else
                ValForCF = TraslateHurdleToValue(ActiveSheet.Cells(1, 38).MergeArea(1, 1).Value, "CSW")
            End If
             
         Case 2
            If rRng.Column = 31 Then
                ValForCF = TraslateHurdleToValue(ActiveSheet.Cells(1, 38).MergeArea(1, 1).Value, "Share")
            ElseIf rRng.Column = 23 Then
                ValForCF = TraslateHurdleToValue(ActiveSheet.Cells(1, 38).MergeArea(1, 1).Value, "USW")
            Else
                ValForCF = TraslateHurdleToValue(ActiveSheet.Cells(1, 38).MergeArea(1, 1).Value, "CSW")
            End If
         Case 3
            rRng.Select
             'ValForCF = TraslateHurdleToValue(rRng.MergeArea(1, 1).Value)
         Case 4
            If rRng.Column = 28 Then
                ValForCF = TraslateHurdleToValue(ActiveSheet.Cells(1, 38).MergeArea(1, 1).Value, "CSW")
            ElseIf rRng.Column = 22 Then
                ValForCF = TraslateHurdleToValue(ActiveSheet.Cells(1, 38).MergeArea(1, 1).Value, "USW")
            Else
                ValForCF = TraslateHurdleToValue(ActiveSheet.Cells(1, 38).MergeArea(1, 1).Value, "Share")
            End If
             'ValForCF = TraslateHurdleToValue(rRng.MergeArea(1, 1).Value)
         Case Else
             ValForCF = 0
     End Select
    
    
    rRng.FormatConditions.AddIconSetCondition
    rRng.FormatConditions(rRng.FormatConditions.Count).SetFirstPriority
    With rRng.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3Symbols)
    End With
    With rRng.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = ValForCF
        .Operator = xlGreaterEqual
    End With
    With rRng.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = ValForCF
        .Operator = xlGreater
    End With
Exit_Routine:
    On Error Resume Next
    Exit Sub
Err_Routine:
    ''If Err.Number = 91 Then f_Cond_Formatting = "": GoTo Exit_Routine
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("mTEN_Runtime-f-f_Cond_Formatting", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    'Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    Stop ' @RWTen Take out when tested and enable Call g_Write_Err_Table
    Resume Next
    GoTo Exit_Routine
    Resume Next

End Sub

Private Function TraslateHurdleToValue(ByVal Hurdle As String, ByVal HurdleType As String) As Single
    If HurdleType = "USW" Then
        If Hurdle = "H" Or Hurdle = "B" Then TraslateHurdleToValue = 15
        If Hurdle = "S" Then TraslateHurdleToValue = 15
    ElseIf HurdleType = "CSW" Then
        TraslateHurdleToValue = 15
    ElseIf HurdleType = "Share" Then
        If Hurdle = "S" Or Hurdle = "H" Then TraslateHurdleToValue = 0.02
        If Hurdle = "B" Then TraslateHurdleToValue = 0.03
    End If
End Function

Private Function p_RtnStr(sFla As String, lUT_Pos_Top As Long)
    ' Will change the RC number in a formula into a letter / number formt - i.e. ^R00C22) into "V260"
    Dim sSRC As String, sNewSRC As String, lPos As Long, lRNo As Long, lCNo As Long
    p_RtnStr = sFla
    lPos = InStr(1, sFla, "^")
    If lPos > 0 Then
        sSRC = Mid(sFla, lPos, 7)
        If Val(Mid(sSRC, 3, 2)) > 0 Then
            lRNo = Val(Mid(sSRC, 3, 2))
        Else
            lRNo = lUT_Pos_Top
        End If
        sNewSRC = g_GetExcelCell(lRNo, CLng(Val(Mid(sSRC, 6, 2))))
        p_RtnStr = Replace(sFla, sSRC, sNewSRC)
    End If
    
End Function

Public Function TenTestWS() As Boolean
    ' This routine will test for the existence of the worksheet
    On Error Resume Next
    TenTestWS = g_WorkSheetExists(ActiveWorkbook, psWShtName)
    Exit Function
End Function

Public Sub TenWSDelete()
    ' This routine will delete a Ten worksheet it it exists
    On Error Resume Next
    ActiveSheet.Unprotect "Passw"
    Call g_WorkSheetDelete(ActiveWorkbook, psWShtName)
    Set pwbWSht = Nothing
    Set psWShtLast = ActiveSheet
    Call ShowIP_ID("SetShow")

    Exit Sub
Err_Routine:
    
    psWShtName = ""

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Call g_EraseAry(pSH_Ary)
    Call g_EraseAry(pSL_Ary)
    Call g_EraseAry(pUT_Ary1)
    Call g_EraseAry(pUT_Ary2)
    Call g_EraseAry(pary_BDBA)
''    Set psdUT_Idx1 = Nothing
    Set psdUT_Idx2 = Nothing
    Set psdSL = Nothing
    Set psdRC_Idx = Nothing
    Set psdSH = Nothing
    Set psdND = Nothing
    Set psdCG = Nothing
    Set pclsPFV = Nothing
    Set psdPFV = Nothing
End Sub

