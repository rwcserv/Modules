Attribute VB_Name = "CCM_Runtime"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Private CCM_Prods As Variant
Private CCM_Buyer As String
Private CCM_CG As Long
Private CCM_SCG As Long
Private CCM_MatchMapping() As MatchMap


Type MatchMap
    AldiPCode As String
    ColesWNAT1 As String
    ColesWNAT2 As String
    ColesWNAT3 As String
    ColesWNAT4 As String
    ColesWNSW As String
    ColesWVIC As String
    ColesWQLD As String
    ColesWSA As String
    ColesWWA As String
    WWWeb As String
    WWWNAT1 As String
    WWWNAT2 As String
    WWWNAT3 As String
    WWWNAT4 As String
    WWWNSW As String
    WWWVIC As String
    WWWQLD As String
    WWWSA As String
    WWWWA As String
    ColesWeb As String
    ColesSB1 As String
    ColesSB2 As String
    ColesSB3 As String
    ColesPL1 As String
    ColesPL2 As String
    ColesPL3 As String
    ColesPL4 As String
    ColesVal1 As String
    ColesVal2 As String
    ColesVal3 As String
    ColesCB1 As String
    ColesPB1 As String
    ColesPB2 As String
    ColesML1 As String
    ColesML2 As String
    ColesML3 As String
    WWWW1 As String
    WWWW2 As String
    WWWW3 As String
    WWWW4 As String
    WWWW5 As String
    WWHB1 As String
    WWHB2 As String
    WWHB3 As String
    WWCB1 As String
    WWPB1 As String
    WWPB2 As String
    WWML1 As String
    WWML2 As String
    WWML3 As String
    DM1 As String
    DM1Pack As String
    DM2 As String
    DM2Pack As String
    DMQ As String
    DMQPack As String
    FC1 As String
    FC1Pack As String
    FC2 As String
    FC2Pack As String
    FCQ As String
    FCQPack As String
End Type
Private CMM_DefaultDataset As Long

Sub CCM_setDefaultDataset(ByVal DatasetNo As Long)
    CMM_DefaultDataset = DatasetNo
End Sub
Function CMM_getDefaultDataset() As Long
    CMM_getDefaultDataset = CMM_DefaultDataset
End Function
Function CBA_COM_getActiveProdsCGSCGBuyer() As Variant
    CCM_SQLQueries.CBA_COM_MATCHGenPullSQL ("getActiveProdsCGSCGBuyer")
    CBA_COM_getActiveProdsCGSCGBuyer = CBA_CBISarr
    Erase CBA_CBISarr
End Function
Function CBA_COM_GetBuyersNames() As Variant()
    CCM_SQLQueries.CBA_COM_MATCHGenPullSQL ("getBuyerlisting")
    CBA_COM_GetBuyersNames = CBA_CBISarr
    Erase CBA_CBISarr
End Function
Function CBA_COM_GetGBDNames() As Variant()
    CCM_SQLQueries.CBA_COM_MATCHGenPullSQL ("getGBDlisting")
    CBA_COM_GetGBDNames = CBA_CBISarr
    Erase CBA_CBISarr
End Function

Sub CCM_MatchingSelectorActivate()
    Dim a As Long
    Dim BuyerDA As Boolean, CGDA As Boolean, SCGDA As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    CCM_Prods = CBA_COM_Runtime.getCCMProds
    CCM_Buyer = CBA_COM_Runtime.getCCMBuyer
    CCM_CG = CBA_COM_Runtime.getCCMCG
    CCM_SCG = CBA_COM_Runtime.getCCMSCG
    
    If CCM_Buyer <> "" Then BuyerDA = True
    If CCM_CG > 0 Then CGDA = True
    If CCM_SCG > 0 Then SCGDA = True
    If CCM_SCG > 0 And CCM_CG = 0 And CCM_Buyer = "" Then GoTo errorSCGonlyWTF
    
    If BuyerDA = True Then
        'USE THIS TO DETERMINE IF JASON OR PAUL THEN ALCOHOL
    
    End If
    
    If CBA_COM_frm_MatchingTool.getCCM_UserDefinedState = False Then
        If CBA_COM_Runtime.getWeekstoUse > 4 And CBA_COM_Runtime.CBA_COM_getonlymatchedSKU = False Then
            MsgBox "CAD Datasets cannot contain more than 4 weeks data"
            Exit Sub
        ElseIf CBA_COM_Runtime.getWeekstoUse > 16 And CBA_COM_Runtime.CBA_COM_getonlymatchedSKU = True Then
            MsgBox "OMD Datasets cannot contain more than 16 weeks data"
            Exit Sub
        End If
    End If
    
    
   CBA_COM_frm_MatchingTool.LoadupProcess
    
    
    If CMM_getDefaultDataset = 0 Then Exit Sub
    'Looking to add ALDI PRODUCT CODE DATA
    On Error Resume Next
    If UBound(CCM_Prods, 2) < 0 Then
    Err.Clear
    On Error GoTo Err_Routine
    GoTo noCCM_ProdsArray
    End If
    On Error GoTo Err_Routine
    
     For a = LBound(CCM_Prods, 2) To UBound(CCM_Prods, 2)
    
        If BuyerDA = True Then
            If SCGDA = True Then
                If CGDA = True Then
                    If CCM_Buyer = CCM_Prods(6, a) And Format(CStr(CCM_CG), "00") = CCM_Prods(2, a) And Format(CStr(CCM_SCG), "00") = CCM_Prods(4, a) Then
                        CBA_COM_frm_MatchingTool.Box_ProdList.AddItem CCM_Prods(0, a) & "-" & CCM_Prods(1, a)
                    End If
                Else
                    'THIS SHOULD NOT HAPPEN!!!!!!!!!!!!!!
                End If
            ElseIf CGDA = True Then
                    If CCM_Buyer = CCM_Prods(6, a) And Format(CStr(CCM_CG), "00") = CCM_Prods(2, a) Then
                        CBA_COM_frm_MatchingTool.Box_ProdList.AddItem CCM_Prods(0, a) & "-" & CCM_Prods(1, a)
                    End If
            Else
                If CCM_Buyer = CCM_Prods(6, a) Then
                    CBA_COM_frm_MatchingTool.Box_ProdList.AddItem CCM_Prods(0, a) & "-" & CCM_Prods(1, a)
                    'CBA_COM_frm_MatchingTool.Box_ProdList(1).Value = CCM_Prods(0, a) '& "-" & CCM_Prods(1, a)
                End If
            End If
        ElseIf CGDA = True Then
            If SCGDA = True Then
                If Format(CStr(CCM_CG), "00") = CCM_Prods(2, a) And Format(CStr(CCM_SCG), "00") = CCM_Prods(4, a) Then
                    CBA_COM_frm_MatchingTool.Box_ProdList.AddItem CCM_Prods(0, a) & "-" & CCM_Prods(1, a)
                End If
            Else
            
            End If
        
        End If
    Next
    If CBA_COM_frm_MatchingTool.Box_ProdList.ListCount > 0 Then CBA_COM_frm_MatchingTool.updateMatchCollections
noCCM_ProdsArray:
    CBA_COM_frm_MatchingTool.Show vbModeless

Exit Sub
errorSCGonlyWTF:
    MsgBox "Something very wierd has happened. I suggest you just disable COMRADE and start again... Or call Tom on 9218", vbOKOnly
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CCM_MatchingSelectorActivate", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Function CCM_dropMouseClassModule()
''On Error Resume Next

End Function


Function CCM_QueryMatches(Optional Aldiprodno As Collection) As Variant()
    Dim strProd As String, bOutput As Boolean, prod
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    On Error Resume Next
    If Aldiprodno.Count < 0 Then
        Err.Clear
        On Error GoTo Err_Routine
        bOutput = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("CCM_MatchMap")
    Else
        Err.Clear
        On Error GoTo Err_Routine
        CBA_BasicFunctions.CBA_sortCollection Aldiprodno
        strProd = ""
        For Each prod In Aldiprodno
            If strProd = "" Then strProd = prod Else strProd = strProd & ", " & prod
        Next
        bOutput = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("CCM_MatchMapbyprod", , , , , , strProd)
    End If
    
    If bOutput = True Then
        CCM_QueryMatches = CBA_COMarr
        Erase CBA_COMarr
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CCM_QueryMatches", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Function CCM_updateMatches() As MatchMap()
    Dim tmm() As Variant
    Dim newmm() As MatchMap
    Dim lNum As Long, a As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
  
    tmm = CCM_QueryMatches
    lNum = 0
    ReDim newmm(1 To 1)
    For a = LBound(tmm, 2) To UBound(tmm, 2)
        lNum = lNum + 1
        ReDim Preserve newmm(1 To lNum)
        newmm(lNum).AldiPCode = tmm(0, a)
        'asdasd = IsNA(tmm(MatchType("ColesWNAT1").MappingTableNumber, a), 0)
    
        If IsNull(tmm(MatchType("ColesWNAT1").MappingTableNumber, a)) = False Then newmm(lNum).ColesWNAT1 = tmm(MatchType("ColesWNAT1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesWNAT2").MappingTableNumber, a)) = False Then newmm(lNum).ColesWNAT2 = tmm(MatchType("ColesWNAT2").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesWNAT3").MappingTableNumber, a)) = False Then newmm(lNum).ColesWNAT3 = tmm(MatchType("ColesWNAT3").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesWNAT4").MappingTableNumber, a)) = False Then newmm(lNum).ColesWNAT4 = tmm(MatchType("ColesWNAT4").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesWNSW").MappingTableNumber, a)) = False Then newmm(lNum).ColesWNSW = tmm(MatchType("ColesWNSW").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesWVIC").MappingTableNumber, a)) = False Then newmm(lNum).ColesWVIC = tmm(MatchType("ColesWVIC").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesWQLD").MappingTableNumber, a)) = False Then newmm(lNum).ColesWQLD = tmm(MatchType("ColesWQLD").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesWSA").MappingTableNumber, a)) = False Then newmm(lNum).ColesWSA = tmm(MatchType("ColesWSA").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesWWA").MappingTableNumber, a)) = False Then newmm(lNum).ColesWWA = tmm(MatchType("ColesWWA").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWeb").MappingTableNumber, a)) = False Then newmm(lNum).WWWeb = tmm(MatchType("WWWeb").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWNAT1").MappingTableNumber, a)) = False Then newmm(lNum).WWWNAT1 = tmm(MatchType("WWWNAT1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWNAT2").MappingTableNumber, a)) = False Then newmm(lNum).WWWNAT2 = tmm(MatchType("WWWNAT2").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWNAT3").MappingTableNumber, a)) = False Then newmm(lNum).WWWNAT3 = tmm(MatchType("WWWNAT3").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWNAT4").MappingTableNumber, a)) = False Then newmm(lNum).WWWNAT4 = tmm(MatchType("WWWNAT4").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWNSW").MappingTableNumber, a)) = False Then newmm(lNum).WWWNSW = tmm(MatchType("WWWNSW").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWVIC").MappingTableNumber, a)) = False Then newmm(lNum).WWWVIC = tmm(MatchType("WWWVIC").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWQLD").MappingTableNumber, a)) = False Then newmm(lNum).WWWQLD = tmm(MatchType("WWWQLD").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWSA").MappingTableNumber, a)) = False Then newmm(lNum).WWWSA = tmm(MatchType("WWWSA").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWWA").MappingTableNumber, a)) = False Then newmm(lNum).WWWWA = tmm(MatchType("WWWWA").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesWeb").MappingTableNumber, a)) = False Then newmm(lNum).ColesWeb = tmm(MatchType("ColesWeb").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesSB1").MappingTableNumber, a)) = False Then newmm(lNum).ColesSB1 = tmm(MatchType("ColesSB1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesSB2").MappingTableNumber, a)) = False Then newmm(lNum).ColesSB2 = tmm(MatchType("ColesSB2").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesSB3").MappingTableNumber, a)) = False Then newmm(lNum).ColesSB3 = tmm(MatchType("ColesSB3").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesPL1").MappingTableNumber, a)) = False Then newmm(lNum).ColesPL1 = tmm(MatchType("ColesPL1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesPL2").MappingTableNumber, a)) = False Then newmm(lNum).ColesPL2 = tmm(MatchType("ColesPL2").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesPL3").MappingTableNumber, a)) = False Then newmm(lNum).ColesPL3 = tmm(MatchType("ColesPL3").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesPL4").MappingTableNumber, a)) = False Then newmm(lNum).ColesPL4 = tmm(MatchType("ColesPL4").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesVal1").MappingTableNumber, a)) = False Then newmm(lNum).ColesVal1 = tmm(MatchType("ColesVal1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesVal2").MappingTableNumber, a)) = False Then newmm(lNum).ColesVal2 = tmm(MatchType("ColesVal2").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesVal3").MappingTableNumber, a)) = False Then newmm(lNum).ColesVal3 = tmm(MatchType("ColesVal3").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesCB1").MappingTableNumber, a)) = False Then newmm(lNum).ColesCB1 = tmm(MatchType("ColesCB1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesPB1").MappingTableNumber, a)) = False Then newmm(lNum).ColesPB1 = tmm(MatchType("ColesPB1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesPB2").MappingTableNumber, a)) = False Then newmm(lNum).ColesPB2 = tmm(MatchType("ColesPB2").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesML1").MappingTableNumber, a)) = False Then newmm(lNum).ColesML1 = tmm(MatchType("ColesML1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesML2").MappingTableNumber, a)) = False Then newmm(lNum).ColesML2 = tmm(MatchType("ColesML2").MappingTableNumber, a)
        If IsNull(tmm(MatchType("ColesML3").MappingTableNumber, a)) = False Then newmm(lNum).ColesML3 = tmm(MatchType("ColesML3").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWeb").MappingTableNumber, a)) = False Then newmm(lNum).WWWeb = tmm(MatchType("WWWeb").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWW1").MappingTableNumber, a)) = False Then newmm(lNum).WWWW1 = tmm(MatchType("WWWW1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWW2").MappingTableNumber, a)) = False Then newmm(lNum).WWWW2 = tmm(MatchType("WWWW2").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWW3").MappingTableNumber, a)) = False Then newmm(lNum).WWWW3 = tmm(MatchType("WWWW3").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWW4").MappingTableNumber, a)) = False Then newmm(lNum).WWWW4 = tmm(MatchType("WWWW4").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWWW5").MappingTableNumber, a)) = False Then newmm(lNum).WWWW5 = tmm(MatchType("WWWW5").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWHB1").MappingTableNumber, a)) = False Then newmm(lNum).WWHB1 = tmm(MatchType("WWHB1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWHB2").MappingTableNumber, a)) = False Then newmm(lNum).WWHB2 = tmm(MatchType("WWHB2").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWHB3").MappingTableNumber, a)) = False Then newmm(lNum).WWHB3 = tmm(MatchType("WWHB3").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWCB1").MappingTableNumber, a)) = False Then newmm(lNum).WWCB1 = tmm(MatchType("WWCB1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWPB1").MappingTableNumber, a)) = False Then newmm(lNum).WWPB1 = tmm(MatchType("WWPB1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWPB2").MappingTableNumber, a)) = False Then newmm(lNum).WWPB2 = tmm(MatchType("WWPB2").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWML1").MappingTableNumber, a)) = False Then newmm(lNum).WWML1 = tmm(MatchType("WWML1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWML2").MappingTableNumber, a)) = False Then newmm(lNum).WWML2 = tmm(MatchType("WWML2").MappingTableNumber, a)
        If IsNull(tmm(MatchType("WWML3").MappingTableNumber, a)) = False Then newmm(lNum).WWML3 = tmm(MatchType("WWML3").MappingTableNumber, a)
        If IsNull(tmm(MatchType("DM1").MappingTableNumber, a)) = False Then newmm(lNum).DM1 = tmm(MatchType("DM1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("DM1").MappingTableNumberPack, a)) = False Then newmm(lNum).DM1Pack = tmm(MatchType("DM1").MappingTableNumberPack, a)
        If IsNull(tmm(MatchType("DM2").MappingTableNumber, a)) = False Then newmm(lNum).DM2 = tmm(MatchType("DM2").MappingTableNumber, a)
        If IsNull(tmm(MatchType("DM2").MappingTableNumberPack, a)) = False Then newmm(lNum).DM2Pack = tmm(MatchType("DM2").MappingTableNumberPack, a)
        If IsNull(tmm(MatchType("DMQ").MappingTableNumber, a)) = False Then newmm(lNum).DMQ = tmm(MatchType("DMQ").MappingTableNumber, a)
        If IsNull(tmm(MatchType("DMQ").MappingTableNumberPack, a)) = False Then newmm(lNum).DMQPack = tmm(MatchType("DMQ").MappingTableNumberPack, a)
        If IsNull(tmm(MatchType("FC1").MappingTableNumber, a)) = False Then newmm(lNum).FC1 = tmm(MatchType("FC1").MappingTableNumber, a)
        If IsNull(tmm(MatchType("FC1").MappingTableNumberPack, a)) = False Then newmm(lNum).FC1Pack = tmm(MatchType("FC1").MappingTableNumberPack, a)
        If IsNull(tmm(MatchType("FC2").MappingTableNumber, a)) = False Then newmm(lNum).FC2 = tmm(MatchType("FC2").MappingTableNumber, a)
        If IsNull(tmm(MatchType("FC2").MappingTableNumberPack, a)) = False Then newmm(lNum).FC2Pack = tmm(MatchType("FC2").MappingTableNumberPack, a)
        If IsNull(tmm(MatchType("FCQ").MappingTableNumber, a)) = False Then newmm(lNum).FCQ = tmm(MatchType("FCQ").MappingTableNumber, a)
        If IsNull(tmm(MatchType("FCQ").MappingTableNumberPack, a)) = False Then newmm(lNum).FCQPack = tmm(MatchType("FCQ").MappingTableNumberPack, a)

    Next
    CCM_MatchMapping = newmm
    CCM_updateMatches = CCM_MatchMapping
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CCM_updateMatches", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Function CCM_getMatches() As MatchMap()
    CCM_getMatches = CCM_MatchMapping
End Function

Function CCM_MatchArrayData(Optional ByVal useMatchHistory As Boolean = False, Optional strProd As String, Optional DateFrom As Date, Optional DateTo As Date, Optional CGno As Long, Optional SCGno As Long, Optional ByVal Competitor As String) As Variant()
    Dim stpt As Long, lNum As Long, a As Long, b As Long, prod, cnt As Long
    Dim arrProdSplit() As String
    Dim colProds As Collection
    Dim bfound As Boolean, bOutput As Boolean
    Dim TempArr() As Variant, altarr() As Variant
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    If strProd <> "" Then
        Set colProds = New Collection
        arrProdSplit = Split(strProd, ",")
        For a = LBound(arrProdSplit) To UBound(arrProdSplit)
            colProds.Add Trim(arrProdSplit(a))
        Next
    End If
    
    If useMatchHistory = False Then
    
            bOutput = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("getMMData", DateFrom, DateTo, CGno, CStr(SCGno), useMatchHistory, strProd)
            If strProd <> "" And bOutput = True Then
                lNum = -1 ': st = 0
                For Each prod In colProds
                    bfound = False
                    For a = 0 To UBound(CBA_COMarr, 2)
                        prod = Replace(prod, Chr(13), "")
                        If Trim(prod) = CBA_COMarr(0, a) Then
                            bfound = True
                            lNum = lNum + 1
                            ReDim Preserve TempArr(0 To UBound(CBA_COMarr, 1), 0 To lNum)
                            For b = 0 To UBound(CBA_COMarr, 1)
                                TempArr(b, lNum) = CBA_COMarr(b, a)
                            Next
                            'st = a
                        ElseIf bfound = True Then
                            Exit For
                        End If
                    Next
                Next
                Erase CBA_COMarr
                
                On Error Resume Next
                If UBound(TempArr, 2) < 0 Then
                    bOutput = False
                    Err.Clear
                End If
                On Error GoTo Err_Routine
                
            ElseIf bOutput = True Then
                TempArr = CBA_COMarr
                Erase CBA_COMarr
            End If
            If bOutput = True Then
                If LCase(Competitor) = "all competitors" Or Competitor = "" Then
                    For a = LBound(TempArr, 2) To UBound(TempArr, 2)
                        TempArr(3, a) = CCM_Mapping.CMM_getComp2Find(TempArr(3, a), TempArr(1, a))
                    Next
                    CCM_MatchArrayData = TempArr
                    Erase TempArr
                Else
                    cnt = -1
                    ReDim altarr(0 To 5, 0 To 0)
                    For a = LBound(TempArr, 2) To UBound(TempArr, 2)
                        If InStr(1, LCase(CCM_Mapping.CMM_getComp2Find(TempArr(3, a), TempArr(1, a))), Competitor) > 0 Then
                            cnt = cnt + 1
                            If UBound(altarr, 2) < cnt Then ReDim Preserve altarr(0 To 5, 0 To cnt)
                            For b = 0 To 5
                                altarr(b, cnt) = TempArr(b, a)
                            Next
                        End If
                    Next
                    CCM_MatchArrayData = altarr
                    Erase TempArr
                    Erase altarr
                End If
    
            End If
 
    Else
      
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CCM_MatchArrayData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    ''If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function















