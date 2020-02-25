Attribute VB_Name = "CBA_COM_Runtime"
Option Explicit
Option Private Module          ' Excel users cannot access procedures

Private CBA_COM_ToggleButtonState As Boolean
Private CBA_COM_RibbonState As Boolean
Private CBA_COM_SCGState As Boolean
Private CBA_COM_MatchingToolState As Boolean
Private CBA_COM_COLECTBuyers As Collection
Private CBA_COM_COLECTCGS As Collection
Private CBA_COM_COLECTSCGS As Collection
Private CBA_COM_ChosenBuyer As String
Private CBA_COM_ChosenCG As Long
Private CBA_COM_ChosenSCG As Long
Private SKUarr() As CBA_COM_COMCompSKU
Private CBA_COM_ACTIVEPRODS() As Variant
Private CBA_COM_SKU_AllProds() As Variant
Private CCS_Results() As Variant
Private CCS_WWData As Boolean
Private CCS_ColesData As Boolean
Private CCS_DMData As Boolean
Private CCS_FCData As Boolean
Private CBA_CSS_FORM As CBA_frm_CCS
Private CBA_COM_WeekstoUse As Long
Private CBA_COM_AdminUser As Boolean
Private CBA_COM_OnlyMatchedSKUs As Boolean
Private CBA_CopyDelete_Results() As Variant
Private CBA_COM_ToggleOn As Boolean
Private CBA_WedDate As Date
Private WordApp

Sub CBA_COM_GetTogglePressed(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = CBA_COM_ToggleOn
End Sub
Function CBA_COM_isTogglePressed()
    CBA_COM_isTogglePressed = CBA_COM_ToggleOn
End Function
Sub CBA_COM_OpenUserManual(Control As IRibbonControl)
    Dim IEapp As Object
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    If CBA_getVersionStatus(g_GetDB("Gen"), CBA_COM_Ver, "Comrade", "COM", True) = "Exit" Then Exit Sub
    Set IEapp = CreateObject("InternetExplorer.Application")
    IEapp.navigate "https://collab.aldi-599.loc/teams/aus-cb-dtm/BDTM/_layouts/15/WopiFrame2.aspx?sourcedoc=/teams/aus-cb-dtm/BDTM/Buying%20Process%20General/021005b%20-%20COMRADE%20User%20Manual.doc&action=default"
    IEapp.Visible = True
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_COM_OpenUserManual", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Sub CBA_COM_RefreshRibbon(Optional ByVal ControlID As String)
    Dim CBA_Proc As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    If CBA_Rib Is Nothing Then
        Set CBA_Rib = CBA_Ribbon.CBA_GetRibbon(CBA_DataSheet.Range("A1").Value)
        If ControlID = "" Then
            CBA_Rib.Invalidate
        Else
            CBA_Rib.InvalidateControl ControlID
        End If
    Else
        If ControlID = "" Then
            CBA_Rib.Invalidate
        Else
            CBA_Rib.InvalidateControl ControlID
        End If
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_COM_RefreshRibbon", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Sub CBA_COM_ActivateData(Control As IRibbonControl, pressed As Boolean)
    Dim CBA_Proc As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    If pressed = True Then
        If CBA_getVersionStatus(g_GetDB("Gen"), CBA_COM_Ver, "Comrade", "COM", True) = "Exit" Then Exit Sub
        If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "A connection to COMRADE is being established..." '& Chr(10) & Chr(10) & "This sould only take a few seconds.."
        'Application.ScreenUpdating = False
'        CBA_COM_MATCHRuntime.CBA_COM_PullArraysMulti
        'SKUarr = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("WW")
        
        
        CBA_WedDate = DateAdd("D", 1 - WeekDay(Date, 4), Date)
        CBA_COM_MatchingToolState = False: CBA_COM_RibbonState = True: CBA_COM_SCGState = False
        CBA_COM_Runtime.CBA_COM_BuildRibbonArrays
        CCM_Runtime.CCM_updateMatches
        
        
        'CBA_COM_BUYERS = CBA_COM_MATCHRuntime.CBA_COM_GetBuyersNames
        If UBound(CBA_COM_SKU_AllProds, 2) > 0 Then CBA_COM_ToggleOn = True Else CBA_COM_ToggleOn = False
        If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
        CBA_COM_RefreshRibbon
        
        'Application.ScreenUpdating = False
    Else
        CBA_COM_Runtime.CBA_COM_ErasePublicVariables
        CBA_COM_SCGState = False
        CBA_COM_SetRibbonState False
        CBA_COM_SetMatchingToolState False
        CBA_COM_SetToggleButtonState True
        CBA_COM_ToggleOn = False
        CBA_COM_RefreshRibbon
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_COM_ActivateData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Sub CBA_COM_BuildRibbonArrays()
Dim a As Long
Dim bfound As Boolean, bOutput As Boolean, Buyer
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

Set CBA_COM_COLECTBuyers = New Collection
    CBA_COM_ACTIVEPRODS = CCM_Runtime.CBA_COM_getActiveProdsCGSCGBuyer
        For a = LBound(CBA_COM_ACTIVEPRODS, 2) To UBound(CBA_COM_ACTIVEPRODS, 2)
            bfound = False
            For Each Buyer In CBA_COM_COLECTBuyers
                If Buyer = CBA_COM_ACTIVEPRODS(6, a) Then
                    bfound = True
                    Exit For
                End If
            Next
            If bfound = False Then CBA_COM_COLECTBuyers.Add CBA_COM_ACTIVEPRODS(6, a)
        Next
    bOutput = CBA_COM_SQLQueries.CBA_COM_GenPullSQL("CBA_COM_SKU_Prods", , , 2)
    If bOutput = True Then
        CBA_COM_SKU_AllProds = CBA_COMarr
        Erase CBA_COMarr
    End If
    CBAR_Runtime.CBAR_setGBSs
    CBAR_Runtime.CBAR_SetRep
    CBAR_Runtime.CBAR_setEmails
    CBA_BasicFunctions.CBA_sortCollection CBA_COM_COLECTBuyers
    CBA_COM_UpdateChosenValues
    
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_COM_BuildRibbonArrays", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Sub CBA_COM_getEnabledCOMRADETools(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = True
End Sub
'Callback for CBA_COM_DeleteMatches onAction
Sub CBA_COM_OpenDeleteMatches(Control As IRibbonControl)
    CBA_frm_DeleteMatching.Show vbModeless
End Sub
Sub CBA_COM_OpenCopyMatches(Control As IRibbonControl)
    CBA_frm_CopyMatching.Show vbModeless
End Sub
'Callback for CBA_COM_ONLYMATCHED onAction
Sub ToggleOnlyMATCHDATA(Control As IRibbonControl, pressed As Boolean)
    If CBA_COM_getonlymatchedSKU <> pressed Then
        Erase CCM_WWSKU: Erase CCM_ColesSKU: Erase CCM_DMSKU: Erase CCM_FCSKU: Erase CCM_UDWWSKU: Erase CCM_UDColesSKU: Erase CCM_UDDMSKU: Erase CCM_UDFCSKU
        If pressed = True Then
            MsgBox "COMRADE set to build OMD Datasets" & Chr(10) & Chr(10) & " Any presviously build datasets have now been dropped"
        Else
            MsgBox "COMRADE set to build CAD Datasets" & Chr(10) & Chr(10) & " Any presviously build datasets have now been dropped"
        End If
    End If
    CBA_COM_OnlyMatchedSKUs = pressed
End Sub
'Callback for CBA_COM_ONLYMATCHED getPressed
Sub isOnlyMAtchDATAActive(Control As IRibbonControl, ByRef returnedVal)
    CBA_COM_OnlyMatchedSKUs = returnedVal
End Sub
Function CBA_COM_getonlymatchedSKU() As Boolean
    CBA_COM_getonlymatchedSKU = CBA_COM_OnlyMatchedSKUs
End Function
'Callback for CBA_COM_WeeksUsed onAction
Sub CBA_COM_WeeksUsed_onAction(Control As IRibbonControl, id As String, iIndex As Integer)
Dim oldW2U As Long, lRet As Long
    If CBA_COM_WeekstoUse = 0 Then
        CBA_COM_WeekstoUse = iIndex + 1
    Else
        oldW2U = CBA_COM_WeekstoUse
        If oldW2U <> iIndex + 1 Then
            lRet = MsgBox("Any Datasets already created will be dropped. Do you still want to change the week paramater?", vbYesNo)
            If lRet = 6 Then
                CBA_COM_WeekstoUse = iIndex + 1
                Erase CCM_WWSKU: Erase CCM_ColesSKU: Erase CCM_DMSKU: Erase CCM_FCSKU: Erase CCM_UDWWSKU: Erase CCM_UDColesSKU: Erase CCM_UDDMSKU: Erase CCM_UDFCSKU
            Else
                CBA_COM_RefreshRibbon "CBA_COM_WeeksUsed"
            End If
        End If
    End If
End Sub

'Callback for CBA_COM_WeeksUsed getItemCount
Sub CBA_COM_WeeksUsed_getItemCount(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = 52
End Sub
'Callback for CBA_COM_WeeksUsed getItemLabel
Sub CBA_COM_WeeksUsed_getItemLabel(Control As IRibbonControl, iIndex As Integer, ByRef returnedVal)
        On Error Resume Next
        'Debug.Print control.id
        returnedVal = iIndex + 1
        On Error GoTo 0
End Sub
'Sub CBA_COM_WeeksUsedSelectedItemID(control As IRibbonControl, ByRef itemID As Variant)
'    If IsEmpty(CBA_COM_WeeksUsed) = True Then CBA_COM_WeeksUsed = 2
'    itemID = CBA_COM_WeeksUsed
'End Sub
Sub CBA_COM_WeeksUsedSelectedItemIndex(Control As IRibbonControl, ByRef returnedVal)
    If IsEmpty(CBA_COM_WeekstoUse) = True Then CBA_COM_WeekstoUse = 2
    returnedVal = CBA_COM_WeekstoUse - 1
End Sub
Function getWeekstoUse() As Long
    If CBA_COM_WeekstoUse = 0 Then CBA_COM_WeekstoUse = 2
    getWeekstoUse = CBA_COM_WeekstoUse
End Function
Function getBuyers() As Collection
    Set getBuyers = CBA_COM_COLECTBuyers
End Function
Function getCGs() As Collection
    Set getCGs = CBA_COM_COLECTCGS
End Function
Function getSCGs() As Collection
    Set getSCGs = CBA_COM_COLECTSCGS
End Function

Function updateCopyDelete_Results()
    Dim a As Long, b As Long
    Dim thisarr() As Variant
    Dim tmm() As Variant
    Dim bOutput As Boolean, bfound As Boolean
    Dim strProds As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    
    Erase CBA_CopyDelete_Results
    ReDim thisarr(1 To 5, 1 To UBound(CBA_COM_ACTIVEPRODS, 2) + 1)
    
    
    For a = LBound(CBA_COM_ACTIVEPRODS, 2) To UBound(CBA_COM_ACTIVEPRODS, 2)
        For b = LBound(CBA_COM_ACTIVEPRODS, 1) To UBound(CBA_COM_ACTIVEPRODS, 1)
            If b > 1 Then
                If b = 2 Then thisarr(3, a + 1) = CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(3, a)
                If b = 4 Then thisarr(4, a + 1) = CBA_COM_ACTIVEPRODS(4, a) & "-" & CBA_COM_ACTIVEPRODS(5, a)
                If b = 6 Then thisarr(5, a + 1) = CBA_COM_ACTIVEPRODS(6, a)
            Else
                thisarr(b + 1, a + 1) = CBA_COM_ACTIVEPRODS(b, a)
            End If
        Next
    Next
    
    
    bOutput = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("getAldiProdsforAllMatched")
    tmm = CBA_COMarr
    Erase CBA_COMarr
    If bOutput = True Then
        For a = LBound(tmm, 2) To UBound(tmm, 2)
            If strProds = "" Then strProds = tmm(0, a) Else strProds = strProds & ", " & tmm(0, a)
        Next
        bOutput = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("getAldiProdInfo", , , , , , strProds)
        If bOutput = True Then
            tmm = CBA_CBISarr
            Erase CBA_CBISarr
            For a = LBound(tmm, 1) To UBound(tmm, 2)
                bfound = False
                For b = LBound(CBA_COM_ACTIVEPRODS, 2) To UBound(CBA_COM_ACTIVEPRODS, 2)
                    If CBA_COM_ACTIVEPRODS(0, b) = CLng(tmm(0, a)) Then
                        bfound = True
                        Exit For
                    ElseIf CBA_COM_ACTIVEPRODS(0, b) > CLng(tmm(0, a)) Then
                        Exit For
                    End If
                Next
                If bfound = False Then
                    ReDim Preserve thisarr(1 To 5, 1 To UBound(thisarr, 2) + 1)
                    thisarr(1, UBound(thisarr, 2)) = tmm(0, a)
                    thisarr(2, UBound(thisarr, 2)) = tmm(1, a)
                    thisarr(3, UBound(thisarr, 2)) = tmm(2, a) & "-" & tmm(3, a)
                    thisarr(4, UBound(thisarr, 2)) = tmm(4, a) & "-" & tmm(5, a)
                    thisarr(5, UBound(thisarr, 2)) = tmm(6, a)
                End If
            Next
        End If
    End If
    
    
    CBA_BasicFunctions.CBA_Sort2DArray thisarr, 1, 1
    thisarr = CBA_BasicFunctions.CBA_TransposeArray(thisarr)
    
    
    '    For a = LBound(thisarr, 2) To UBound(thisarr, 2)
    '        For b = LBound(thisarr, 1) To UBound(thisarr, 1)
    '          '  Debug.Print thisarr(b, a)
    '        Next
    '    Next
    
    
    CBA_CopyDelete_Results = thisarr
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-updateCopyDelete_Results", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next


End Function
Function getCopyDelete_Results() As Variant()

    On Error Resume Next
    If UBound(CBA_CopyDelete_Results, 2) = 1 Then
    On Error GoTo 0
        updateCopyDelete_Results
    End If
    Err.Clear
    On Error GoTo 0
    getCopyDelete_Results = CBA_CopyDelete_Results

End Function
Sub CBA_COM_UpdateChosenValues(Optional BuyerVal As String, Optional CGVal As Long, Optional SCGVal As Long)
    Dim a As Long, CG, scg
    Dim bfound As Boolean, BuyerChange As Boolean, CGChange As Boolean, SCGChange As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    If CBA_COM_ChosenBuyer <> BuyerVal Then
        CGVal = 0: BuyerChange = True: CGChange = True: SCGChange = True: SCGVal = 0
    End If
    If CBA_COM_ChosenCG <> CGVal Then CGChange = True
    If CBA_COM_ChosenSCG <> SCGVal Then SCGChange = True
    If CGVal = 0 Then Set CBA_COM_COLECTCGS = New Collection
    Set CBA_COM_COLECTSCGS = New Collection
    
    
        If BuyerVal = "" Then
            CBA_COM_ChosenBuyer = ""
            If CGVal = 0 Then
                CBA_COM_ChosenCG = 0
                If SCGVal = 0 Then
                    CBA_COM_ChosenSCG = 0
                Else
                    CBA_COM_ChosenSCG = SCGVal
                End If
            ElseIf SCGVal = 0 Then
                  CBA_COM_ChosenCG = CGVal
            Else
                CBA_COM_ChosenCG = CGVal
                CBA_COM_ChosenSCG = SCGVal
            End If
        ElseIf BuyerVal <> "" Then
            CBA_COM_ChosenBuyer = BuyerVal
            If CGVal <> 0 Then
                CBA_COM_ChosenCG = CGVal
                If SCGVal <> 0 Then
                    CBA_COM_ChosenSCG = SCGVal
                Else
                    CBA_COM_ChosenSCG = 0
                End If
            ElseIf SCGVal <> 0 Then
                CBA_COM_ChosenCG = 0
                CBA_COM_ChosenSCG = SCGVal
            ElseIf CGVal = 0 And SCGVal = 0 Then
                CBA_COM_ChosenCG = 0
                CBA_COM_ChosenSCG = 0
            End If
        End If
    
    
        If BuyerVal = "" And CGVal = 0 And SCGVal = 0 Then
            For a = LBound(CBA_COM_ACTIVEPRODS, 2) To UBound(CBA_COM_ACTIVEPRODS, 2)
                bfound = False
                For Each CG In CBA_COM_COLECTCGS
                    If CG = CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(3, a) Then
                        bfound = True
                        Exit For
                    End If
                Next
                If bfound = False Then CBA_COM_COLECTCGS.Add CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(3, a)
                bfound = False
                For Each scg In CBA_COM_COLECTSCGS
                    If scg = CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(4, a) & "-" & CBA_COM_ACTIVEPRODS(5, a) Then
                        bfound = True
                        Exit For
                    End If
                Next
                If bfound = False Then CBA_COM_COLECTSCGS.Add CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(4, a) & "-" & CBA_COM_ACTIVEPRODS(5, a)
            Next
        Else
            If BuyerVal <> "" Then
                If CGVal = 0 And SCGVal = 0 Then
                    For a = LBound(CBA_COM_ACTIVEPRODS, 2) To UBound(CBA_COM_ACTIVEPRODS, 2)
                        If CBA_COM_ACTIVEPRODS(6, a) = BuyerVal Then
                            bfound = False
                            For Each CG In CBA_COM_COLECTCGS
                                If CG = CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(3, a) Then
                                    bfound = True
                                    Exit For
                                End If
                            Next
                            If bfound = False Then CBA_COM_COLECTCGS.Add CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(3, a)
                            bfound = False
                            For Each scg In CBA_COM_COLECTSCGS
                                If scg = CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(4, a) & "-" & CBA_COM_ACTIVEPRODS(5, a) Then
                                    bfound = True
                                    Exit For
                                End If
                            Next
                            If bfound = False Then CBA_COM_COLECTSCGS.Add CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(4, a) & "-" & CBA_COM_ACTIVEPRODS(5, a)
                        End If
                    Next
                Else
                    If CGVal <> 0 Then
                        If SCGVal <> 0 Then
                            For a = LBound(CBA_COM_ACTIVEPRODS, 2) To UBound(CBA_COM_ACTIVEPRODS, 2)
                                If CLng(CBA_COM_ACTIVEPRODS(2, a)) = CGVal Then
                                    CBA_COM_COLECTCGS.Add CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(3, a)
                                    Exit For
                                End If
                            Next
                            For a = LBound(CBA_COM_ACTIVEPRODS, 2) To UBound(CBA_COM_ACTIVEPRODS, 2)
                                If CLng(CBA_COM_ACTIVEPRODS(2, a)) = CGVal And CLng(CBA_COM_ACTIVEPRODS(4, a)) = SCGVal Then
                                    CBA_COM_COLECTSCGS.Add CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(4, a) & "-" & CBA_COM_ACTIVEPRODS(5, a)
                                    Exit For
                                End If
                            Next
                        Else
                            For a = LBound(CBA_COM_ACTIVEPRODS, 2) To UBound(CBA_COM_ACTIVEPRODS, 2)
                                If CBA_COM_ACTIVEPRODS(6, a) = BuyerVal And CLng(CBA_COM_ACTIVEPRODS(2, a)) = CGVal Then
                                    bfound = False
                                    For Each scg In CBA_COM_COLECTSCGS
                                        If scg = CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(4, a) & "-" & CBA_COM_ACTIVEPRODS(5, a) Then
                                            bfound = True
                                            Exit For
                                        End If
                                    Next
                                    If bfound = False Then CBA_COM_COLECTSCGS.Add CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(4, a) & "-" & CBA_COM_ACTIVEPRODS(5, a)
                                End If
                            Next
                        End If
                    Else
                        'this should not occur where you have a SCGVal without a CG or Buyer Val
                        CBA_COM_SCGState = False
                    End If
                End If
            Else
                If CGVal <> 0 Then
                    For a = LBound(CBA_COM_ACTIVEPRODS, 2) To UBound(CBA_COM_ACTIVEPRODS, 2)
                        If CLng(CBA_COM_ACTIVEPRODS(2, a)) = CGVal Then
                            bfound = False
                            For Each scg In CBA_COM_COLECTSCGS
                                If scg = CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(4, a) & "-" & CBA_COM_ACTIVEPRODS(5, a) Then
                                    bfound = True
                                    Exit For
                                End If
                            Next
                            If bfound = False Then CBA_COM_COLECTSCGS.Add CBA_COM_ACTIVEPRODS(2, a) & "-" & CBA_COM_ACTIVEPRODS(4, a) & "-" & CBA_COM_ACTIVEPRODS(5, a)
                        End If
                    Next
                End If
            End If
        End If
            
            CBA_BasicFunctions.CBA_sortCollection CBA_COM_COLECTBuyers
            CBA_BasicFunctions.CBA_sortCollection CBA_COM_COLECTCGS
            CBA_BasicFunctions.CBA_sortCollection CBA_COM_COLECTSCGS
            
            If BuyerVal <> "" Or CGVal <> 0 Then
                CBA_COM_MatchingToolState = True
            Else
                CBA_COM_MatchingToolState = False
            End If
            If CGVal <> 0 Then CBA_COM_SCGState = True Else CBA_COM_SCGState = False
            If BuyerChange = True Then CBA_COM_RefreshRibbon "CBA_COM_CGSelector"
            If CGChange = True Then CBA_COM_RefreshRibbon "CBA_COM_SCGSelector"
            CBA_COM_RefreshRibbon "CBA_COM_OpenMatchingSelector"
    '    CBA_COM_RefreshRibbon "CBA_COM_CopyMatches"
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CBA_COM_UpdateChosenValues", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub


'Callback for CBA_COM_SKU_Search onChange
Sub CBA_COM_SKU_Search_onChange(Control As IRibbonControl, text As String)
    Dim aVal() As String
    Dim cnt As Long, POS As Long, numfnd As Long, a As Long, b As Long, c As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    Erase CCS_Results
    cnt = 0: POS = 1: numfnd = 0
    ReDim aVal(1 To 1)
    If InStr(1, LCase(text), "*+*") > 0 Then
        For a = 1 To Len(LCase(text))
            If Mid(LCase(text), a, 3) = "*+*" Then
                cnt = cnt + 1
                ReDim Preserve aVal(1 To cnt)
                aVal(cnt) = LCase(Trim(Mid(LCase(text), POS, a - POS)))
                POS = a + 3
            End If
            If a = Len(LCase(text)) Then
                If LCase(Trim(Mid(LCase(text), POS, a - POS))) <> "" Then
                cnt = cnt + 1
                ReDim Preserve aVal(1 To cnt)
                aVal(cnt) = LCase(Trim(Mid(LCase(text), POS, a + 1 - POS)))
                End If
            End If
        Next
    Else
        cnt = cnt + 1
        aVal(cnt) = Trim(LCase(text))
    End If
    
    On Error Resume Next
    If LBound(CBA_COM_SKU_AllProds, 2) = 100 Then
        On Error GoTo Err_Routine
        CBA_COM_BuildRibbonArrays
    End If
    On Error GoTo Err_Routine
    
    
    numfnd = 0
    For a = LBound(CBA_COM_SKU_AllProds, 2) To UBound(CBA_COM_SKU_AllProds, 2)
        For b = 1 To cnt
            If InStr(1, LCase(CBA_COM_SKU_AllProds(2, a)), aVal(b)) > 0 And _
                (((CCS_WWData = True And CBA_COM_SKU_AllProds(0, a) = "Woolworths") Or _
                (CCS_ColesData = True And CBA_COM_SKU_AllProds(0, a) = "Coles") Or _
                (CCS_DMData = True And CBA_COM_SKU_AllProds(0, a) = "Dan Murphys") Or _
                (CCS_FCData = True And CBA_COM_SKU_AllProds(0, a) = "First Choice")) Or _
                CCS_WWData = False And CCS_ColesData = False And CCS_DMData = False And CCS_FCData = False) Then
                numfnd = numfnd + 1
                ReDim Preserve CCS_Results(1 To 3, 1 To numfnd)
                For c = 1 To 3
                    CCS_Results(c, numfnd) = CBA_COM_SKU_AllProds(c - 1, a)
                Next
            End If
        Next
    Next
        
    If numfnd = 0 Then
        MsgBox "No Products match those descriptions", vbOKOnly
    Else
        On Error Resume Next
        Unload CBA_CSS_FORM
        Err.Clear
        On Error GoTo 0
        Set CBA_CSS_FORM = New CBA_frm_CCS
        CBA_CSS_FORM.Show vbModeless
    End If
    
    CBA_COM_RefreshRibbon Control.id
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("CBA_COM_SKU_Search_onChange", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    ''If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Function getCCS_Results() As Variant()
    'Debug.Print CCS_Results(1, 1)
    getCCS_Results = CCS_Results
End Function
'Callback for CBA_COM_SKU_Cbox_WW onAction
Sub EnableWWData(Control As IRibbonControl, pressed As Boolean)
    CCS_WWData = pressed
End Sub

'Callback for CBA_COM_SKU_Cbox_WW getPressed
Sub isWWDataActive(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = CCS_WWData
End Sub
Sub EnableColesData(Control As IRibbonControl, pressed As Boolean)
    CCS_ColesData = pressed
End Sub

'Callback for CBA_COM_SKU_Cbox_Coles getPressed
Sub isColesDataActive(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = CCS_ColesData
End Sub
Sub EnableDMData(Control As IRibbonControl, pressed As Boolean)
    CCS_DMData = pressed
End Sub
'Callback for CBA_COM_SKU_Cbox_DM getPressed
Sub isDMDataActive(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = CCS_DMData
End Sub
Sub EnableFCData(Control As IRibbonControl, pressed As Boolean)
    CCS_FCData = pressed
End Sub
'Callback for CBA_COM_SKU_Cbox_FC getPressed
Sub isFCDataActive(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = CCS_FCData
End Sub
Sub CBA_COM_getEnabledCOMRADE(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = CBA_COM_RibbonState
End Sub
Sub CBA_COM_GetEnabledOpenMatchingSelector(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = CBA_COM_MatchingToolState
End Sub
Sub CBA_COM_SetRibbonState(ByVal COM_state As Boolean)
    CBA_COM_RibbonState = COM_state
End Sub
Sub CBA_COM_SetMatchingToolState(ByVal COM_state As Boolean)
    CBA_COM_MatchingToolState = COM_state
End Sub
Sub CBA_COM_BuyerSelector_getItemCount(Control As IRibbonControl, ByRef returnedVal)
    If CBA_COM_RibbonState = True Then
        returnedVal = CBA_COM_COLECTBuyers.Count
    End If
End Sub
Sub CBA_COM_BuyerSelector_getItemLabel(Control As IRibbonControl, iIndex As Integer, ByRef returnedVal)
        On Error Resume Next
        returnedVal = CBA_COM_COLECTBuyers(iIndex + 1)
        On Error GoTo 0
End Sub
Sub CBA_COM_CGSelector_getItemCount(Control As IRibbonControl, ByRef returnedVal)
    If CBA_COM_RibbonState = True Then
        returnedVal = CBA_COM_COLECTCGS.Count
    End If
End Sub
Sub CBA_COM_CGSelector_getItemLabel(Control As IRibbonControl, iIndex As Integer, ByRef returnedVal)
        On Error Resume Next
        returnedVal = CBA_COM_COLECTCGS(iIndex + 1)
        On Error GoTo 0
End Sub
Sub CBA_COM_SCGSelector_getItemCount(Control As IRibbonControl, ByRef returnedVal)
    If CBA_COM_SCGState = True Then
        returnedVal = CBA_COM_COLECTSCGS.Count
    End If

End Sub
Sub CBA_COM_SCGSelector_getItemLabel(Control As IRibbonControl, iIndex As Integer, ByRef returnedVal)
        If CBA_COM_SCGState = True Then
            On Error Resume Next
            returnedVal = CBA_COM_COLECTSCGS(iIndex + 1)
            On Error GoTo 0
        End If
End Sub
Sub CBA_COM_BuyerSelector_onAction(Control As IRibbonControl, id As String, iIndex As Integer)
    CBA_COM_UpdateChosenValues CBA_COM_COLECTBuyers(iIndex + 1), CBA_COM_ChosenCG, CBA_COM_ChosenSCG
End Sub
Sub CBA_COM_CGSelector_onAction(Control As IRibbonControl, id As String, iIndex As Integer)
    CBA_COM_UpdateChosenValues CBA_COM_ChosenBuyer, Mid(CBA_COM_COLECTCGS(iIndex + 1), 1, 2) ', CBA_COM_ChosenSCG
End Sub
Sub CBA_COM_SCGSelector_onAction(Control As IRibbonControl, id As String, iIndex As Integer)

'Debug.Print Mid(CBA_COM_COLECTSCGS(iIndex + 1), InStr(1, CBA_COM_COLECTSCGS(iIndex + 1), "-") + 1, 2)
    
    
    CBA_COM_UpdateChosenValues CBA_COM_ChosenBuyer, CBA_COM_ChosenCG, Mid(CBA_COM_COLECTSCGS(iIndex + 1), InStr(1, CBA_COM_COLECTSCGS(iIndex + 1), "-") + 1, 2)
End Sub
Sub CBA_COM_OpenMatchingselector(Control As IRibbonControl)
    CBA_COM_frm_MatchingTool.setCCM_UserDefinedState (False)
    CCM_Runtime.CCM_MatchingSelectorActivate
End Sub
Function getCCMProds() As Variant()
    getCCMProds = CBA_COM_ACTIVEPRODS
End Function
Function getCCMBuyer() As String
    getCCMBuyer = CBA_COM_ChosenBuyer
End Function
Function getCCMCG() As Long
    getCCMCG = CBA_COM_ChosenCG
End Function
Function getCCMSCG() As Long
    getCCMSCG = CBA_COM_ChosenSCG
End Function
Sub CBA_COM_SKU_Search_getText(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = ""
End Sub
'Sub CBA_COM_getEnabledActivateData(control As IRibbonControl, ByRef returnedVal)
'    returnedVal = CBA_COM_ToggleButtonState
'End Sub
Sub CBA_COM_SetToggleButtonState(ByVal COM_state As Boolean)
    CBA_COM_ToggleButtonState = COM_state
End Sub
Sub CBA_COM_getEnabledSCGData(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = CBA_COM_SCGState
End Sub
Sub CBA_COM_SetRibbonSCGState(ByVal COM_state As Boolean)
    CBA_COM_SCGState = COM_state
End Sub
'Callback for CBAR_ReportSelector onAction
Sub CBAR_ReportSelector_onAction(Control As IRibbonControl, id As String, iIndex As Integer)
    Dim TempArr() As String
    If CBA_getVersionStatus(g_GetDB("Gen"), CBA_COM_Ver, "Comrade", "COM", True) = "Exit" Then Exit Sub
    TempArr = CBAR_Runtime.CBAR_getNamesOfReports(True)
    CBAR_Runtime.createActiveReport (TempArr(iIndex + 1))
End Sub
'Callback for CBAR_ReportSelector getItemCount
Sub CBAR_ReportSelector_getItemCount(Control As IRibbonControl, ByRef returnedVal)
    Dim user As String
    user = Application.UserName
    If InStr(1, LCase(user), "lentini, david") > 0 Or InStr(1, LCase(user), "pearce, tom") > 0 Or InStr(1, LCase(user), "hollier, moira") > 0 Or InStr(1, LCase(user), "white, robert") > 0 _
        Or InStr(1, LCase(user), "collett, sarah") > 0 Or InStr(1, LCase(user), "woods, mark") > 0 _
            Or InStr(1, LCase(user), "sanders, justine") > 0 Or InStr(1, LCase(user), "hillsmith, caroline") > 0 Or InStr(1, LCase(user), "baines, stuart") > 0 Then
        CBA_COM_AdminUser = True
    Else
        CBA_COM_AdminUser = False
    End If
    returnedVal = CBAR_Runtime.CBAR_getNoOfReports(CBA_COM_AdminUser)
End Sub
'Callback for CBAR_ReportSelector getItemLabel
Sub CBAR_ReportSelector_getItemLabel(Control As IRibbonControl, iIndex As Integer, ByRef returnedVal)
Dim TempArr() As String
    TempArr = CBAR_Runtime.CBAR_getNamesOfReports(CBA_COM_AdminUser)
    returnedVal = TempArr(iIndex + 1)
End Sub
'Callback for CBA_RunCREP onAction
Sub CBARQry(Control As IRibbonControl)
    CBAR_ReportParamaters.Show
End Sub
Function CBAR_getAdminUsers()
    Dim RCell
    If CBA_COM_AdminUser = False Then
        For Each RCell In CBAR_AdminUsers.Columns(1).Cells
            If RCell.Value = "" Then Exit For
            If RCell.Value = Application.UserName Then
                CBA_COM_AdminUser = True
                Exit For
            End If
        Next
    End If
    If InStr(1, Application.UserName, "Baines, Stuart") > 0 Then CBA_COM_AdminUser = True                ' @RW Take out when have put Stuart into the worksheet
    If InStr(1, Application.UserName, "Lentini, David") > 0 Then CBA_COM_AdminUser = True                ' @RW Take out when have tested
    CBAR_getAdminUsers = CBA_COM_AdminUser
End Function

Sub CBA_COM_ErasePublicVariables()
    CBA_DBtoQuery = 0: intRefreshSec = 0: CBA_strAldiMsg = ""
    Set CBA_CBISarr = Nothing: Set CBA_ABIarr = Nothing: Set CBA_COMarr = Nothing: Set CBA_MMSarr = Nothing: Set CBA_COM_colInput = Nothing
    Set CBA_COM_potGram = Nothing: Set CBA_COM_potLitres = Nothing: Set CBA_COM_leftovers = Nothing: Set CBA_COM_potMetres = Nothing: Set CBA_COM_potPieces = Nothing: Set CBA_COM_potPair = Nothing
    Set CBA_COM_potOther = Nothing: Set CBA_COM_potSheet = Nothing: Set CBA_COM_colAdddetail = Nothing: Set CBA_COM_colMulti = Nothing: Set CBA_COM_colWhere = Nothing: Set CBA_COM_colNotDecoded = Nothing
    Erase CBA_COM_PackarrOutput: Erase CBA_COM_arrOutput: Erase CBA_COM_arrSortDetail: Erase CBA_COM_CBISarrOutput: CBA_COM_numOutput = 0
    Erase CBA_COM_arrWW: Erase CBA_COM_arrDM: Erase CBA_COM_arrCBISPack
    CBA_COM_ACGno = 0: CBA_COM_ASCGno = 0: CBA_COM_APcode = 0: CBA_COM_entryrow = 0: Set CBA_colProds = Nothing: Set CBA_COM_colmm = Nothing: Set CBA_COM_owsht = Nothing: CBA_COM_owret = 0
    CBA_COM_owpr = 0: CBA_COM_Aret = 0: CBA_COM_owrow = 0: CBA_COM_owcol = 0: CBA_COM_statelookup = "": Erase CBA_COM_matchedinfo
    Set CBA_COM_COLECTSCGS = Nothing: Set CBA_COM_COLECTCGS = Nothing: Set CBA_COM_COLECTBuyers = Nothing
    CBA_COM_ChosenBuyer = "": CBA_COM_ChosenCG = 0: CBA_COM_ChosenSCG = 0: CBA_COM_WeekstoUse = 2
    Erase SKUarr: Erase CBA_COM_SKU_AllProds: Erase CCS_Results: Erase CBA_CopyDelete_Results
    CCS_WWData = False: CCS_ColesData = False: CCS_DMData = False: CCS_FCData = False
    Erase CCM_WWSKU: Erase CCM_ColesSKU: Erase CCM_DMSKU: Erase CCM_FCSKU: Erase CCM_UDWWSKU: Erase CCM_UDColesSKU: Erase CCM_UDDMSKU: Erase CCM_UDFCSKU
    On Error Resume Next
    Unload CBA_COM_frm_MatchingTool
    Unload CBA_CSS_FORM
    Err.Clear
    On Error GoTo 0
    CCM_Runtime.CCM_setDefaultDataset 0
End Sub
Function CBA_getWedDate(Optional ByVal dtWedDate As Date) As Date
    Dim CN As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim livedata As Boolean
    Dim strSQL As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    'If g_IsDate(dtWedDate) Then CBA_WedDate = CDate(dtWedDate)
    'if g_IsDate(CBA_WedDate) Then CBA_WedDate = Date
''    If CBA_WedDate = 0 Or IsMissing(dtWedDate) Then CBA_WedDate = DateAdd("D", 1 - WeekDay(Date, 4), Date)          ' @RW Added 'IsMissing' as is now producing an error
''    If CBA_WedDate = 0 Or IsMissing(dtWedDate) Then dtWedDate = Date                                                ' @RW OR THIS
    If CBA_WedDate = 0 Then CBA_WedDate = DateAdd("D", 1 - WeekDay(Date, 4), Date)
    If dtWedDate = CStr(Date) Then
        livedata = True
        Set CN = New ADODB.Connection
        With CN
            .ConnectionTimeout = 50
            .CommandTimeout = 50
            .Open "Provider= SQLNCLI10; DATA SOURCE= " & CBA_BasicFunctions.TranslateServerName("599DBL12", Date) & "; ;INTEGRATED SECURITY=sspi;"
        End With
        Set RS = New ADODB.Recordset
        If livedata = True Then strSQL = "select count(colesproductid) from tools.dbo.com_c_prod where datescraped = '" & Format(Date, "YYYY-MM-DD") & "'": RS.Open strSQL, CN
        If RS.EOF Then livedata = False Else RS.Close: Set RS = Nothing: Set RS = New ADODB.Recordset
        If livedata = True Then strSQL = "select count(stockcode) from tools.dbo.com_w_prod where datescraped = '" & Format(Date, "YYYY-MM-DD") & "'": RS.Open strSQL, CN
        If RS.EOF Then livedata = False Else RS.Close: Set RS = Nothing: Set RS = New ADODB.Recordset
        If livedata = True Then strSQL = "select count(Productid) from tools.dbo.com_fc_prod where datescraped = '" & Format(Date, "YYYY-MM-DD") & "'": RS.Open strSQL, CN
        If RS.EOF Then livedata = False Else RS.Close: Set RS = Nothing: Set RS = New ADODB.Recordset
        If livedata = True Then strSQL = "select count(productid) from tools.dbo.com_dm_prod where datescraped = '" & Format(Date, "YYYY-MM-DD") & "'": RS.Open strSQL, CN
        If RS.EOF Then livedata = False Else RS.Close: Set RS = Nothing: Set RS = New ADODB.Recordset
        If livedata = False Then CBA_getWedDate = DateAdd("D", -7, CBA_WedDate) Else CBA_getWedDate = CBA_WedDate
        CN.Close
        Set CN = Nothing
    Else
        CBA_getWedDate = CBA_WedDate
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_getWedDate", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Function CBA_COM_getCompProds() As Variant
    CBA_COM_getCompProds = CBA_COM_SKU_AllProds
End Function
















