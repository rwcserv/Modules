Attribute VB_Name = "CBA_ASyst"
Option Explicit             ' Module for the ASyst Super Saver Project @CBA_ASyst Changed 190623
Option Private Module       ' Excel users cannot access procedures

'               FldNam, FldTyp  , Lck Sts , Vis Sts ,Audit Sts, New Values
Private Const A_FN = 0, A_NV = 1 ' A_Ft = 1, A_Clr = 2, A_LK = 3, A_Vs = 4, A_Ad = 5,
Private cls_Prod As CBA_AST_Product

Public Function CBA_AST_getProdClassMod(ByRef clsASTProdCMod As CBA_AST_Product) As Boolean
    
    If cls_Prod Is Nothing Then
        CBA_AST_getProdClassMod = False
    Else
        CBA_AST_getProdClassMod = True
        Set clsASTProdCMod = cls_Prod
    End If
  
End Function
Public Function CBA_AST_DestroyClassModule() As Boolean
    If cls_Prod Is Nothing Then
        CBA_AST_DestroyClassModule = True
    Else
        CBA_AST_DestroyClassModule = True
        Set cls_Prod = Nothing
    End If
End Function
Public Sub CBA_AST_Promo(Control As IRibbonControl)
    Call AST_Setup
    CBA_AST_frm_Promotions.Show vbModeless
End Sub

Public Sub CBA_AST_Products(Control As IRibbonControl)
    Call AST_Setup
    CBA_AST_frm_Products.Show vbModeless
End Sub

Public Sub CBA_AST_ProdRows(Control As IRibbonControl)
    Call AST_Setup
    CBA_AST_frm_ProductRows.Show vbModeless
End Sub

Public Function AST_Setup()
    ''Call CBA_getVersionStatus(g_getDB("ASYST", , , True), CBA_AST_Ver, "Super Saver Tool", "AST")    ' Test to see if it is the latest version
    CBA_lPromotion_ID = 0: CBA_lProduct_ID = 0: CBA_lAuthority = 0
    Call g_GetDB("ASYST", , , True)
    ' Get the Authority level
    CBA_lAuthority = AST_getUserASystAuthority(CBA_SetUser, True, True)
    ' Get the AST Fields
    Call AST_FillTagArrays("", -1, 0)
    ' Test for the latest Product data if it hasn't lately been done
    Call AST_Redo_Buyers
End Function

Public Function AST_Dev_Auth(DevUsers_DevIP As String) As String
    Static sReturn As String, sDevIP As String
    Dim aBup() As String, aBup2() As String, lIdx As Long, sSep As String, sName1 As String, sName2 As String, sExt As String
    ' Will return the list of maintainers of the system, or whether the user is dev staff
                                ' DevUsers                                  DevIP                   Init will redo the tests
    If sDevIP = "" Or DevUsers_DevIP = "Init" Then
        sReturn = g_getConfig("ASTDevUsers", CBA_BSA & CBA_GEN_DB, False, "N")
        aBup = Split(sReturn, ";"): sDevIP = "N": sSep = ""
        For lIdx = 0 To UBound(aBup, 1)
            If aBup(lIdx) > "" Then
                If InStr(1, aBup(lIdx), CBA_User) > 0 Then
                    sDevIP = "Y"
                    GoTo SkipDev
                End If
            End If
        Next
    End If
SkipDev:
    AST_Dev_Auth = sDevIP
End Function


Public Function AST_Redo_Buyers(Optional sInit As String = "NoInit", Optional sReset As String = "N") As String
    ' This routine tests to see if the L2_Product Records have the latest data - the deciding date is contained within the config table
    ' Note this routine will also enter associated data that may only have to be done once (when records have been added from a spreadsheet)
    Static sCheckBuyersOK As String, sCheckBuyerDate As String, sCheckDataOK As String, sCheckDataDate As String, sApplyRegions As String
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim strSQL As String, sSaveBuyerParm As String, sSaveDataParm As String, aryProds() As String, sGST As String
    Dim lProdAryIdx As Long, lProds As Long, lProdIdx As Long, sSep As String, bValidPC As Boolean, bDataPC As Boolean
    Dim sEF As String, sECF As String, scg As String, sSCg As String, sASTDoRedoBuyers As String
    Dim kyCG As Variant, sCGSCG As String, sCGSCGDesc As String, sACatCGSCG As String, sACatDesc As String, dctCG As Scripting.Dictionary, dctCG_SCG As Scripting.Dictionary
    Const MAX_PCS_IN_PASS = 1000
    On Error GoTo Err_Routine
    
    ' Reset if specified ' sApplyRegions = "Y":sCheckDataOK = "Y":sASTDoRedoBuyers="Y":sReset = "Y"
    If sInit = "Init" Then
        sCheckBuyersOK = ""
        sCheckDataOK = ""
    End If
    sApplyRegions = g_getConfig("ASTApplyRegions", g_GetDB("ASYST"), , "N")
    If sApplyRegions = "Y" Then
        Call g_UpdConfig("ASTApplyRegions", g_GetDB("ASYST"), "N")
        sCheckDataOK = "Y"
    End If
    sASTDoRedoBuyers = g_getConfig("ASTDoRedoBuyers", g_GetDB("ASYST"), , "N")
    If sASTDoRedoBuyers = "N" Then GoTo Exit_Routine

    AST_Redo_Buyers = "Done"
    CBA_ErrTag = ""
    ' Check whether the Buyer Data needs to be updated
    If sCheckBuyersOK = "" Then
        CBA_ErrTag = "Buyer"
        sCheckBuyerDate = g_getConfig("RedoBuyerParms", g_GetDB("ASYST"), , "N")
        ' Split the parms and see what they are...
        sCheckBuyersOK = Left(sCheckBuyerDate, 1)
        sCheckBuyerDate = Mid(sCheckBuyerDate, 2, 10)
        sSaveBuyerParm = sCheckBuyersOK
        sCheckBuyersOK = "N"
        If g_IsDate(sCheckBuyerDate) = False Then
            sCheckBuyersOK = "N"                               ' N means the parms have to be reset and the buyers have to be redone in the Products table
        ElseIf CDate(g_FixDate(sCheckBuyerDate)) >= Date - Val(sCheckBuyersOK) Then
            sCheckBuyersOK = "Y"
        End If
        ' If a reset is required
        If sCheckBuyersOK = "N" Then
            If Val(sSaveBuyerParm) = 0 And g_IsNumeric(sSaveBuyerParm, bAllowblanks:=False) = False Then sSaveBuyerParm = 3 ' Fix up value if incorrect
            sCheckBuyerDate = g_FixDate(Date)                       ' Parm reset
        Else
            sCheckBuyersOK = "Y"                                    ' Y means the buyers don't have to be redone in the Products table
        End If
        ' If an update is required
        If sCheckBuyersOK = "N" Then
            Call g_UpdConfig("RedoBuyerParms", g_GetDB("ASYST"), sSaveBuyerParm & g_FixDate(Date) & CBA_User)
            sCheckBuyersOK = "Y"
        End If
    Else
        GoTo Exit_Routine
    End If
    ' Check whether the Product Data needs to be updated
    If sCheckDataOK = "" Then
        CBA_ErrTag = "Data"
        sCheckDataDate = g_getConfig("RedoDataParms", g_GetDB("ASYST"), , "N")
        ' Split the parms and see what they are...
        sCheckDataOK = Left(sCheckDataDate, 1)
        sCheckDataDate = Mid(sCheckDataDate, 2, 10)
        sSaveDataParm = sCheckDataOK
        sCheckDataOK = "N"
        If g_IsDate(sCheckDataDate) = False Then
            sCheckDataOK = "N"                               ' N means the parms have to be reset and the Datas have to be redone in the Products table
        ElseIf CDate(g_FixDate(sCheckDataDate)) >= Date - Val(sCheckDataOK) Then
            sCheckDataOK = "Y"
        End If
        ' If a reset is required
        If sCheckDataOK = "N" Then
            If Val(sSaveDataParm) = 0 And g_IsNumeric(sSaveDataParm, bAllowblanks:=False) = False Then sSaveDataParm = 3 ' Fix up value if incorrect
            sCheckDataDate = g_FixDate(Date)                        ' Parm reset
        Else
            sCheckDataOK = "Y"                               ' Y means the Datas don't have to be redone in the Products table
        End If
        ' If an update is required
        If sCheckDataOK = "N" Then
            ' Get the user concerned
            Call g_UpdConfig("RedoDataParms", g_GetDB("ASYST"), sSaveDataParm & g_FixDate(Date) & CBA_User)
            sCheckDataOK = "Y"
        End If
    End If
    ' Get all the products that are affected - there may be a 'WHERE Status < x' later on when we are only after the latest
    strSQL = "SELECT PD_Product_Code FROM L2_Products GROUP BY PD_Product_Code;"
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    CBA_ErrTag = "SQL"
    lProdAryIdx = -1: lProds = MAX_PCS_IN_PASS
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";"
    ' Accumulate the ProductCodes into an array
    RS.Open strSQL, CN
    Do While Not RS.EOF
        If lProds >= MAX_PCS_IN_PASS Then
            sSep = ""
            lProds = 0
            lProdAryIdx = lProdAryIdx + 1
            ReDim Preserve aryProds(0 To lProdAryIdx)
        End If
        ' Accumulate the ProductCodes
        aryProds(lProdAryIdx) = aryProds(lProdAryIdx) & sSep & RS!PD_Product_Code
        lProds = lProds + 1
        sSep = ","
        RS.MoveNext
    Loop
    RS.Close
    ' Set up the scripting Dictionaries
    Set dctCG = New Scripting.Dictionary
    Set dctCG_SCG = New Scripting.Dictionary
    ' Process the array
    For lProds = 0 To lProdAryIdx
        ' Get the Product details - PC,CG,SCG,BD,GBD
                    '           0       ,1      ,2                      ,3                      ,4                      ,5                          ,6      ,7      ,8      ,9                  ,10
        strSQL = "SELECT p.ProductCode,p.CGNo,p.SCGNo,emp.EmpSign AS BDInits,em.EmpSign AS GBDInits,cg.Description AS CGDesc,scg.Description AS SCGDesc, p.TaxID, en.CatNo,en.CGNo AS ACG,en.SCGNo AS ASCG,en.Category AS ACat,en.CommodityGroup AS ACG,en.SubCommodityGroup AS ASCG  " & _
                "FROM cbis599p.dbo.Product p " & _
                "LEFT JOIN cbis599p.dbo.tf_acgmap() en ON en.ACGEntityID = p.ACGEntityID " & _
                "LEFT JOIN cbis599p.dbo.CommodityGroup cg on cg.CGNo = p.CGNo " & _
                "LEFT JOIN cbis599p.dbo.SubCommodityGroup scg on scg.CGNo = p.cgno and scg.SCGNo = p.SCGNo " & _
                "LEFT JOIN cbis599p.dbo.EMPLOYEE emp on emp.EmpNo = p.Empno " & _
                "LEFT JOIN cbis599p.dbo.EMPLOYEE em on emp.EmpNo_Grp = em.EmpNo " & _
                "       WHERE p.ProductCode IN (" & aryProds(lProds) & "); "
        CBA_DBtoQuery = 599
        bValidPC = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CBIS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", strSQL, 120, , , False)
        If bValidPC = True Then
            For lProdIdx = 0 To UBound(CBA_CBISarr, 2)
                bValidPC = False
                sSep = ""
                CBA_ErrTag = "SQL"
                strSQL = "UPDATE L2_Products SET "
                ' Set the CG...
                If NZ(CBA_CBISarr(1, lProdIdx), "") > "" Then
                    ' Put the CG into the dictionary
                    If dctCG.Exists(CBA_CBISarr(1, lProdIdx)) = False Then
                        dctCG.Add CBA_CBISarr(1, lProdIdx), 1
                        dctCG_SCG.Add CBA_CBISarr(1, lProdIdx) & "~" & CBA_CBISarr(2, lProdIdx), 1
                    Else
                        If dctCG_SCG.Exists(CBA_CBISarr(1, lProdIdx) & "~" & CBA_CBISarr(2, lProdIdx)) = False Then
                             dctCG_SCG.Add CBA_CBISarr(1, lProdIdx) & "~" & CBA_CBISarr(2, lProdIdx), 1
                        End If
                    End If
                    ' Accumulate the SQL
                    scg = CBA_CBISarr(1, lProdIdx) & IIf(g_Empty(CBA_CBISarr(5, lProdIdx)) > 0, "-" & CBA_CBISarr(5, lProdIdx), "")
                    sSCg = CBA_CBISarr(2, lProdIdx) & IIf(g_Empty(CBA_CBISarr(6, lProdIdx)) > 0, "-" & CBA_CBISarr(6, lProdIdx), "")
                    sCGSCG = CBA_CBISarr(1, lProdIdx) & IIf(g_Empty(CBA_CBISarr(2, lProdIdx)) > 0, "-" & CBA_CBISarr(2, lProdIdx), "")
                    sCGSCGDesc = scg & IIf(g_Empty(sSCg) > 0, " / " & sSCg, "")
                    sGST = IIf(Val(CBA_CBISarr(7, lProdIdx) & "") = 2, "Y", "N")
                    ' ACat data
                    sACatCGSCG = Format(CBA_CBISarr(8, lProdIdx), "00") & "-" & Format(CBA_CBISarr(9, lProdIdx), "00") & "-" & Format(CBA_CBISarr(10, lProdIdx), "00")
                    sACatDesc = Format(CBA_CBISarr(8, lProdIdx), "00") & "-" & CBA_CBISarr(11, lProdIdx) & "/" & Format(CBA_CBISarr(9, lProdIdx), "00") & "-" & CBA_CBISarr(12, lProdIdx) & "/" & Format(CBA_CBISarr(10, lProdIdx), "00") & "-" & CBA_CBISarr(13, lProdIdx)
                    bValidPC = True
                    strSQL = strSQL & sSep & "PD_CGSCG='" & sCGSCG & "'"
                    sSep = ","
                    strSQL = strSQL & sSep & "PD_CGDesc='" & sCGSCGDesc & "'"
                    strSQL = strSQL & sSep & "PD_ACatCGSCG='" & sACatCGSCG & "'"
                    strSQL = strSQL & sSep & "PD_ACatDesc='" & Replace(sACatDesc, "'", "`") & "'"
                    strSQL = strSQL & sSep & "PD_GST='" & sGST & "'"
                End If
                ' Set the BD...
                If NZ(CBA_CBISarr(3, lProdIdx), "") > "" Then
                    bValidPC = True
                    strSQL = strSQL & sSep & "PD_BD='" & CBA_CBISarr(3, lProdIdx) & "'"
                    sSep = ","
                End If
                ' Set the GBD...
                If NZ(CBA_CBISarr(4, lProdIdx), "") > "" Then
                    bValidPC = True
                    strSQL = strSQL & sSep & "PD_GBD='" & CBA_CBISarr(4, lProdIdx) & "'"
                    sSep = ","
                End If
                ' If there is something to write
                If bValidPC = True And g_IsNumeric(CBA_CBISarr(0, lProdIdx), bAllowblanks:=False) Then
                    strSQL = strSQL & " WHERE PD_Product_Code=" & CBA_CBISarr(0, lProdIdx)
                    RS.Open strSQL, CN
                End If
            Next lProdIdx
        End If
    Next lProds
    ' Does an Update of the data need to be done
    If sCheckDataOK = "N" Then GoTo Exit_Routine
    ' Get the SubBasket and SuperSaverType
    CBA_DBtoQuery = 1
    CBA_ErrTag = "SQL"
    strSQL = "SELECT SB_ID,SST_ID,SST_LeadWeeks,CG_CGNo FROM qry_L2_ProductDets WHERE CG_CGNo IN ( "
    sSep = "": bValidPC = False
    ' Set each CG into the SQL
    For Each kyCG In dctCG.Keys
        If NZ(kyCG, "") > "" Then     '' dctCG(kyCG) or kyCG
            bValidPC = True
            strSQL = strSQL & sSep & kyCG  '' dctCG(kyCG) or kyCG
            sSep = ","
        End If
    Next
    strSQL = strSQL & " );"
    bDataPC = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "qry_L2_ProductDets", g_GetDB("ASYST"), CBA_MSAccess, strSQL, 120, , , False)
    If bDataPC = True Then
        Set RS = New ADODB.Recordset
        For lProdIdx = 0 To UBound(CBA_ABIarr, 2)
            strSQL = "UPDATE L2_Products SET PD_Sub_Basket=" & CBA_ABIarr(0, lProdIdx) & ",PD_Super_Saver_Type=" & CBA_ABIarr(1, lProdIdx) & _
                     ",PD_Approval_Req_Date=([L2_Products].[PD_On_Sale_Date]) - " & (NZ(CBA_ABIarr(2, lProdIdx), lProdIdx) * 7) & _
                     " WHERE PD_CGSCG LIKE '" & Format(Val(CBA_ABIarr(3, lProdIdx)), "00") & "%'"
            RS.Open strSQL, CN
''            Debug.Print strSQL
        Next
    
    End If
    ' Get Price Elasticity details
    CBA_DBtoQuery = 1
    CBA_ErrTag = "SQL"
    strSQL = "SELECT CGNo, SCGNo, Elastic, Avg_Price, Elasticity_State, Confidence_Level FROM L0_Price_Elasticity_Sheet WHERE ([CGNo] & '~' & [SCGNo]) IN ("
    sSep = "": bValidPC = False
    ' Set each CG into the SQL
    For Each kyCG In dctCG_SCG.Keys
        If NZ(kyCG, "") > "" Then           '' dctCG_SCG(kyCG) or kyCG
            bValidPC = True
            strSQL = strSQL & sSep & "'" & Val(Left(kyCG, 2)) & "~" & Val(Right(kyCG, 2)) & "'"  '' dctCG_SCG(kyCG) or kyCG
            sSep = ","
        End If
    Next
    strSQL = strSQL & " );"
    bDataPC = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "L0_Price_Elasticity_Sheet", g_GetDB("ASYST"), CBA_MSAccess, strSQL, 120, , , False)
    If bDataPC = True Then
        Set RS = New ADODB.Recordset
        For lProdIdx = 0 To UBound(CBA_ABIarr, 2)
            sSep = ""
            sEF = Left(NZ(CBA_ABIarr(4, lProdIdx), "Low"), 1)
            sECF = Left(NZ(CBA_ABIarr(5, lProdIdx), "Low"), 1)
            strSQL = "UPDATE L2_Products SET PD_Elasticity_Flag='" & sEF & "',PD_Elasticity_Confidence_Flag='" & sECF & "'" & _
                     " WHERE PD_CGSCG = '" & Format(Val(CBA_ABIarr(0, lProdIdx)), "00") & "-" & Format(Val(CBA_ABIarr(1, lProdIdx)), "00") & "'"
            RS.Open strSQL, CN
        Next
    End If
    ' Set the flag to say it is done
    sCheckDataOK = "N"
    ' If the flag states, then check to ensure all the products have region records
    If sApplyRegions = "Y" Then
        Call AST_UpdSalesData
        strSQL = "SELECT PD_ID FROM qry_L2_MissingRegions ORDER BY PD_ID;"
        Set RS = New ADODB.Recordset
        CBA_ErrTag = "SQL"
        ' Apply any missing regions
        RS.Open strSQL, CN
        Do While Not RS.EOF
            If RS!PD_ID >= 57 Or RS!PD_ID <= 60 Then
                CBA_ErrTag = CBA_ErrTag
            End If
            Call AST_CrtProdRegionRecs(RS!PD_ID)
            RS.MoveNext
        Loop
        RS.Close
    End If

Exit_Routine:

    On Error Resume Next
    CN.Close
    Set RS = Nothing
    Set CN = Nothing
    Set dctCG = Nothing
    Set dctCG_SCG = Nothing
    
    Exit Function
        
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_Redo_Buyers", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & strSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function AST_Lock_Table(sDB As String, sPrefix As String, sTable As String, sWhere As String, lAuth As Long, Get_SetX As String) As String
    ' This routine tests to see if a record is locked and flags the record/s as locked
    ' Note this routine will flag many records as locked but I envision that really it will be only one - when an update happens
    Static sSameUserOk As String, sCheckLock As String
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim strSQL As String, dtLockUpd As Date, lRecsUpd As Long
    On Error GoTo Err_Routine
    
    AST_Lock_Table = "Locked": CBA_ErrTag = ""
    dtLockUpd = Format(Now(), "dd/mm/yyyy hh:nn")
    If sSameUserOk = "" Then
        sSameUserOk = g_getConfig("SameUserOK", g_GetDB("ASYST"), , "NY")
        sCheckLock = Left(sSameUserOk, 1)
        sSameUserOk = Right(sSameUserOk, 1)
    End If
    ' Get the user concerned
    CBA_ErrTag = "SQL"
    strSQL = "SELECT " & sPrefix & "LockUpd AS Lock_Upd," & sPrefix & "LockUser AS Lock_User," & sPrefix & "LockFlag AS Lock_Flag FROM " & sTable & " WHERE " & sWhere

    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB(sDB) & ";"
    ' Check that the record/s are not already locked
    If Left(Get_SetX, 3) = "Get" Then
        lRecsUpd = 0
        RS.Open strSQL, CN
        Do While Not RS.EOF
            ' If no validation is being done then skip it
            If sCheckLock = "N" Then GoTo GTMoveNext
            ' If the same user is trying to maintain the same record again, one assumes something went wrong with the last, so prompt
            If RS!Lock_Flag = "Y" And sSameUserOk = "Y" And CBA_User = RS!Lock_User And lAuth <> 1 Then GoTo GTMoveNext
            If RS!Lock_Flag = "Y" Then
                lRecsUpd = lRecsUpd + 1
                If lAuth <> 1 Then
                    MsgBox "Record/s you are attempting to update are already being updated by " & RS!Lock_User
                    AST_Lock_Table = "AlreadyLocked"
                    GoTo Exit_Routine
                Else
                    If MsgBox("Record/s you are attempting to update are already being updated by " & RS!Lock_User & vbCrLf & _
                                    "Press 'Yes' if the record is showing as being maintained when it isn't", vbYesNo + vbDefaultButton2, "Update warning") = vbNo Then
                    AST_Lock_Table = "AlreadyLocked"
                    GoTo Exit_Routine
                End If
                
                End If
            End If
GTMoveNext:
            RS.MoveNext
        Loop
        RS.Close
    End If

    ' Set the record/s as locked
    If Right(Get_SetX, 4) = "SetY" Then
        strSQL = "UPDATE " & sTable & " SET " & sPrefix & "LockFlag = 'Y'," & sPrefix & "LockUser = '" & CBA_User & "'," & sPrefix & "LockUpd = Now()" & " WHERE " & sWhere & ";"
        RS.Open strSQL, CN
    End If
    ' Set the record/s as unlocked
    If Right(Get_SetX, 4) = "SetN" Then
        strSQL = "UPDATE " & sTable & " SET " & sPrefix & "LockFlag = 'N'" & " WHERE " & sWhere & ";"
        RS.Open strSQL, CN
    End If
        
Exit_Routine:

    On Error Resume Next
    CN.Close
    Set RS = Nothing
    Set CN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_Lock_Table", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & strSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function AST_FillTagArrays(ByVal sFullTag As String, lFrmID As Long, lAuth As Long, Optional Return_Tag As String, Optional NoTagRetVal As String = "") As String
    ' This routine will put all the fields into an array, and return their values fully formed when required
    
    Dim p_LNo As Long, P_VNo As Long, sTagName As String, sSQL As String, bPassOK As Boolean, sLock As String, sVis As String, lAltIdx As Long, lAltFID As Long
    ''Dim lIdx2 As Long '', sTagErr As String
    Static aTag(), aaTag() As String, stcsTagName As String, stclFrmNo As Long, stclAuth As Long, stclIdx As Long, stcLoadDone As Boolean
    Const p_Frm = 0, p_FN = 1, p_FD = 2, p_FT = 3, p_Clr = 4, p_Hdg = 5, p_Wth = 6, p_Adt = 7
    On Error GoTo Err_Routine
    
    CBA_ErrTag = ""
    If Not stcLoadDone Or lFrmID < 0 Then
        stcLoadDone = True
        CBA_DBtoQuery = 1
        sSQL = "SELECT FA_TF_ID,FF_FieldName,FF_FieldDesc,FA_FormType,FA_Clr,FF_Hdg,FF_Hdg_Width,FA_Audit,FA_Lock0,FA_Vis0,FA_Lock1,FA_Vis1,FA_Lock2,FA_Vis2,FA_Lock3,FA_Vis3,FA_Lock4,FA_Vis4,FA_Lock5,FA_Vis5,FA_Lock6,FA_Vis6 " & _
               "FROM qry_A2_FormFieldAuth WHERE (FA_Type=0);"
        bPassOK = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "qry_A2_FormFieldAuth", g_GetDB("ASYST"), CBA_MSAccess, sSQL, 120, , , False)
        If bPassOK = True Then
            aTag = CBA_ABIarr
            Erase CBA_ABIarr
        Else
            MsgBox "No field parameters found - Major error"
            CBA_Error = " Error - No fields found in " & "-" & CBA_ProcI(, 0, True) & vbCrLf & sSQL
            CBA_ErrTag = "No Fields"
            GoTo Skip_Err_Routine
        End If
        ' Hop out if an init
        If lFrmID < 0 Then Exit Function
        CBA_DBtoQuery = 3
    End If
    ' Set the default
    AST_FillTagArrays = ""
    ' Get the 1st element of the tag
    sFullTag = Trim(sFullTag)
    aaTag = Split(sFullTag & CBA_S & CBA_S, CBA_S)
    sTagName = aaTag(0)
    If UCase(sTagName) = "REGION" Then
        sTagName = sTagName
    End If
    ' Set up the Authority Level Lock and Vis Indexes
    p_LNo = p_Adt + 1 + (2 * lAuth)
    P_VNo = p_LNo + 1
    If stcsTagName = sTagName And stclFrmNo = lFrmID And stclAuth = lAuth Then
        GoTo GTReturnTag
    Else
        For stclIdx = 0 To UBound(aTag, 2)
            ' Capture any alternate index for the field that may exist - it maybe used later if an index for the actual item is not found
            If aTag(p_FN, stclIdx) = sTagName And aTag(p_Frm, stclIdx) <> lFrmID Then
''                lAltIdx = stclIdx
                If lAltFID = 0 Then
                    lAltIdx = stclIdx    'aTag(p_FN, stclIdx)
                    lAltFID = aTag(p_Frm, stclIdx)
                ElseIf (lAltFID > lFrmID And aTag(p_Frm, stclIdx) < lFrmID) Or (lAltFID < lFrmID And aTag(p_Frm, stclIdx) < lFrmID) Then
                    lAltIdx = stclIdx ''aTag(p_FN, stclIdx)
                    lAltFID = aTag(p_Frm, stclIdx)
                End If
            End If
            ' Capture the index of the fieldname/form found
            If aTag(p_FN, stclIdx) = sTagName And aTag(p_Frm, stclIdx) = lFrmID Then
                stcsTagName = sTagName
                stclFrmNo = lFrmID
                stclAuth = lAuth
                GoTo GTReturnTag
            End If
        Next
    End If
    ' Apply the Alt no if the actual is not found
    If lAltIdx > 0 Then
        stclIdx = lAltIdx
        GoTo GTReturnTag
    End If
    ' Return a default value if the tag isn't found
    AST_FillTagArrays = NoTagRetVal
    Exit Function
GTReturnTag:

    ' Test for what needs to be returned
    sLock = LCase(aTag(p_LNo, stclIdx))
    sVis = LCase(aTag(P_VNo, stclIdx))
    Select Case Return_Tag
    Case Is = "Name"
        AST_FillTagArrays = aTag(p_FN, stclIdx)
    Case Is = "BaseFormat"                                  ' The first three letters of the format - e.g. txt,num,chk etc
        AST_FillTagArrays = LCase(Left(aTag(p_FT, stclIdx), 3))
    Case Is = "Format"                                      ' Format after the first three letter - i.e. the format of the field
        AST_FillTagArrays = LCase(g_Right(aTag(p_FT, stclIdx), 3))
    Case Is = "FullFormat"                                  ' Full Format
        AST_FillTagArrays = LCase(aTag(p_FT, stclIdx))
    Case Is = "Clr"
        AST_FillTagArrays = LCase(aTag(p_Clr, stclIdx))
    Case Is = "Hdg", "Heading"
        AST_FillTagArrays = aTag(p_Hdg, stclIdx)
    Case Is = "Width"
        AST_FillTagArrays = Val(NZ(aTag(p_Wth, stclIdx), "-1"))
        If AST_FillTagArrays = -1 Then
            CBA_Error = " Error - " & aTag(p_FN, stclIdx) & " has no width, set to 20 -" & CBA_ProcI(, 0, True)
            If CBA_TestIP = "Y" Then MsgBox CBA_Error, vbOKOnly
            AST_FillTagArrays = 20
            GoTo Skip_Err_Routine
        End If
    Case Is = "Lock"
        AST_FillTagArrays = sLock
    Case Is = "Vis"
        AST_FillTagArrays = sVis
    Case Is = "Audit"
        AST_FillTagArrays = LCase(aTag(p_Adt, stclIdx))
    Case Is = "ToolTip"                                     ' What to put in to the tool tip
        If LCase(Left(aTag(p_FT, stclIdx), 3)) = "dte" And sLock = lAuth & "alock" Then
            AST_FillTagArrays = aTag(p_FD, stclIdx) & " - Click to change"
        Else
            AST_FillTagArrays = aTag(p_FD, stclIdx)
        End If
    Case Else
        CBA_Error = " Error - Tag Name " & Return_Tag & " not found -" & CBA_ProcI(, 0, True)
        If CBA_TestIP Then MsgBox CBA_Error, vbOKOnly
        GoTo Skip_Err_Routine
    End Select

Exit_Routine:
    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_FillTagArrays", 3)
    CBA_Error = " Error - Field=" & sTagName & " - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
Skip_Err_Routine:
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
'    Resume Next
    GoTo Exit_Routine
    Resume Next
End Function

Public Function AST_getFieldTag(ByRef frmf As Control, ByVal lFrmID As Long, ByVal lAuth As Long, Optional vVal = "", Optional ByVal Return_Tag As String)

    ' This routine will return the relavent Tag details from the Label or the Field itself
    '(note: lock sts etc comes now from the label)
    ' lAuth = (0=Readonly); (1=Admin); (2=BA);(3=BD?)
    
    
    Static aTag() As String, sTag As String, sName As String
    Dim lIdx2 As Long, sAccumTag As String
    On Error GoTo Err_Routine
    
    ' Tag positions ...
    ' FldName (a_fn), Form Field Type eg txt (a_ft) ,  Lock/Lock (a_lk)  ,   Vis/Vis (a_vs) ,    Audit True / False (a_ad)
    vVal = NZ(vVal, "")
    CBA_ErrTag = "Tag"
    sTag = NZ(frmf.Tag, "")
    aTag = Split((sTag & CBA_Sn), CBA_S)
RestartFld:   ' Restart point
    
    Select Case Return_Tag
    Case Is = "ApdUpd"                                  ' Will append '~' to the field's tag to say that it has been updated
        If Len(NZ(aTag(A_NV), "|")) > 1 And Right(NZ(aTag(A_NV), "|"), 1) <> "~" Then
            vVal = "~"
        Else
            vVal = ""
        End If
        frmf.Tag = NZ(frmf.Tag, "") & vVal
        Return_Tag = "Format"            ' Return the format
        GoTo RestartFld
    Case Is = "RtrnUpdTag"
        vVal = aTag(A_NV)              '           Trim(Replace(aTag(A_NV), "~", ""))
        GoSub GSAccum
        AST_getFieldTag = sAccumTag & vVal
    Case Else
        AST_getFieldTag = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, Return_Tag)
''    Case Is = "GetInitTag"
''        AST_getFieldTag = sTag
''    Case Is = "ChkVal"
''        AST_getFieldTag = "|"
''        If aTaga(A_Ad) = "true" And vVal <> Trim(Replace(aTag(A_NV), "~", "")) Then AST_getFieldTag = "~Chg~"
''    Case Is = "Val"
''        If aTaga(A_Ad) = "true" Then
''            AST_getFieldTag = Trim(Replace(aTag(A_NV), "~", ""))
''        Else
''            AST_getFieldTag = "|"
''        End If
    End Select

Exit_Routine:
    On Error Resume Next
    Exit Function
    
GSAccum:
    sAccumTag = ""
    For lIdx2 = A_FN To A_NV - 1
        sAccumTag = sAccumTag & aTag(lIdx2) & CBA_S
    Next
    Return
       
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_getFieldTag", 3)
    CBA_Error = CBA_ErrTag & " Error - Field=" & sName & " - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "Tag" Then
        AST_getFieldTag = "~~"
        GoTo Exit_Routine
    Else
        Debug.Print CBA_Error
        Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
        Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
        AST_getFieldTag = "~~"
        GoTo Exit_Routine
    End If

End Function

Public Sub AST_FillForm(meCtls As Object, sTbl_Flds As String, sFld_Prefix As String, ByVal lFrmID As Long, ByVal lAuth As Long, _
                        Optional bNullFields As Boolean = False, Optional sExceptFields As String = "")
    ' Fill the form fields from the array, obtained from the table concerned
    ' bNullFields will null the entire form
    ' sExceptFields will omit the named field as leave them as they are
    
    Dim frmf As Control, lIdx As Long, bfound As Boolean, lIdx2 As Long, sAccumTag As String
    Dim aTag() As String, aFrmFlds() As String, sTag As String       ', vVal
    Dim sFullFmt As String, sBaseFmt As String
    On Error GoTo Err_Routine
    ' Process the name of the Form fields into a private array
    aFrmFlds = Split(Replace(Replace(sTbl_Flds, sFld_Prefix, ""), "_", ""), ",")
    CBA_ErrTag = ""
    ' Cycle throught the fields on the form to fill their values and format them correctly as per values in the tab
    For Each frmf In meCtls
        CBA_ErrTag = "No"
        If (frmf.Tag & "") > "" And Left(frmf.Name, 3) <> "lbl" Then
            CBA_ErrTag = "Yes"
            bfound = False
            For lIdx = 0 To UBound(aFrmFlds(), 1)
                sTag = frmf.Tag
                If frmf.Name = "txtRetailDiscount" Then
                    sTag = sTag
                End If
                aTag = Split((sTag & CBA_Sn), CBA_S)
                ' FldName (a_fn), Form Field Type eg txt (a_ft) ,  Lock/Lock (a_lk)  ,   Vis/Vis (a_vs) ,    Audit True / False (a_ad)    ,  (Tag makeup)
                If sExceptFields > "" And InStr(1, sExceptFields, aTag(A_FN)) > 0 Then GoTo NoTag
                sBaseFmt = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "BaseFormat")
                If sBaseFmt = "" Then sBaseFmt = Left(frmf.Name, 3)
                If bNullFields Then GoTo NullFields      ' Don't process array if the fields need to be nulled
                If aFrmFlds(lIdx) = aTag(A_FN) Then
                    sFullFmt = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "FullFormat")
                    If g_IsDate(CBA_ABIarr(lIdx, 0)) Then
                        sFullFmt = sFullFmt
                    End If
                    ''aTag = Split((sTag & CBA_Sn), CBA_S)
                    bfound = True
                    If IIf(IsNull(CBA_ABIarr(lIdx, 0)) = True, "", CBA_ABIarr(lIdx, 0)) = "" Then
                        frmf.Value = ""
                    ElseIf sFullFmt = "dtedmyy" Then
                        frmf.Value = g_FixDate(CDate(CBA_ABIarr(lIdx, 0)), CBA_DMY)
                    ElseIf sFullFmt = "dtedmyyhn" Then
                        frmf.Value = g_FixDate(CDate(CBA_ABIarr(lIdx, 0)), CBA_DMYHN)
                    ElseIf sFullFmt = "dted3dmyy" Then
                        frmf.Value = g_FixDate(CDate(CBA_ABIarr(lIdx, 0)), CBA_D3DMY)
                    ElseIf sFullFmt = "dted3dmyyhn" Then
                        frmf.Value = g_FixDate(CDate(CBA_ABIarr(lIdx, 0)), CBA_D3DMYHN)
                    ElseIf sFullFmt = "dted2dmyy" Then
                        frmf.Value = g_FixDate(CDate(CBA_ABIarr(lIdx, 0)), CBA_D2DMY)
                    ElseIf sFullFmt = "dted2dmyyhn" Then
                        frmf.Value = g_FixDate(CDate(CBA_ABIarr(lIdx, 0)), CBA_D2DMYHN)
                    ElseIf sFullFmt = "num#,0" Then
                        frmf.Value = Format(CBA_ABIarr(lIdx, 0), "#,0")
                    ElseIf Left(sFullFmt, 5) = "num0." Then
                        frmf.Value = Format(CBA_ABIarr(lIdx, 0), g_Right(sFullFmt, 3))
                    ElseIf (CBA_ABIarr(lIdx, 0) = "N") And (sFullFmt = "opt" Or sFullFmt = "chk") Then
                        bfound = True
                        frmf.Value = False
                    ElseIf (CBA_ABIarr(lIdx, 0) = "Y") And (sFullFmt = "opt" Or sFullFmt = "chk") Then
                        bfound = True
                        frmf.Value = True

''                    ElseIf sFullFmt = "DDDDateTime" Then
''                        frmf.Value = Format(CDate(CBA_ABIarr(lIdx, 0)), "ddd dd/mm/yyyy hh:nn AMPM")
                    Else
                        frmf.Value = CBA_ABIarr(lIdx, 0)
                    End If
''                ElseIf (aFrmFlds(lIdx) & CBA_ABIarr(lIdx, 0) = aTag(A_FN)) And sFullFmt = "optyn" Then
''                    bFound = True
''                    frmf.Value = True
                End If
''                ' Set the value to add to the tag...
''                If Left(sFullFmt, 3) = "dte" Then
''                     vVal = CDate(CBA_ABIarr(lIdx, 0))
''                Else
''                    vVal = frmf.Visible
''                End If
                If bfound = True Then
'                    If aTag(A_Ad) = "true" Then
'                        'Debug.Print aTag(A_FN) & ";";
                    GoSub GSAccum
                    frmf.Tag = sAccumTag & "~" & frmf.Value
'                    End If
                    Exit For    ' If has been found then hop out
                End If
            Next
NullFields:
            If bfound = False And bNullFields = True Then
                If sBaseFmt <> "cmd" Then
                    CBA_ErrTag = "Set"
                    frmf.Locked = True
                    frmf.BackColor = CBA_Grey
                    frmf.Value = ""
                End If
            End If
        End If
NoTag:
    Next
        
    Exit Sub
GSAccum:
    sAccumTag = ""
    For lIdx2 = A_FN To A_NV - 1
        sAccumTag = sAccumTag & aTag(lIdx2) & CBA_S
    Next
    Return
    
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-f-AST_FillForm", 3)
    If CBA_ErrTag = "No" Then
        Resume NoTag
    ElseIf CBA_ErrTag = "Set" Then
        Resume Next
    Else
        CBA_Error = frmf.Name & "-" & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
        Debug.Print CBA_Error
        Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
        Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
        Resume Next
    End If
End Sub

Public Function AST_FillYellow(cActCtl As Control, sClr As String) As Boolean
    ' This routine will test whether to change a field to yellow or not
    On Error Resume Next
    Static stcDoYellow As String
    If sClr <> "y" Then Exit Function
    If stcDoYellow = "" Then
        stcDoYellow = g_getConfig("FillYellow", g_GetDB("ASYST"), , "N")
    End If
    If stcDoYellow = "Y" Then
        If NZ(cActCtl.Value, "") = "" Then
            cActCtl.BackColor = CBA_EntryYellow
        ElseIf g_IsNumeric(cActCtl.Value) = True And g_UnFmt(cActCtl.Value, "sng") = 0 Then
            cActCtl.BackColor = CBA_EntryYellow
        Else
            cActCtl.BackColor = CBA_White
        End If
    Else
        cActCtl.BackColor = CBA_EntryYellow
    End If
    
End Function


Public Function AST_getUserASystAuthority(Optional ByVal sUserName As String = "", Optional ByVal bReTest As Boolean = False, Optional ByVal bNoTest As Boolean = False) As Long
    ' ASYST System Only
    ' Return the authority level of the user    - 0 = No Authority (read only), 1 = Admin Authority, 2 = BD Authority, 3 = Marketing, 4 = Merchandising, 5 = GBDM Authority, 6 = GBDM Authority
    ' Now will set a default authority for testing purposes
    Dim sReturnTitle As String, sReturnShortTitle As String, aTag() As String
    Static stcAuth As String, stcAuthStr As String, stcAuth2 As String, stcAuth3 As String, stcAuth4 As String, stcAuth5 As String, stcAuth6 As String
    On Error GoTo Err_Routine
    
    If stcAuthStr = "" Then       ' Get the list of authorisations from the config table
        stcAuthStr = g_getConfig("AuthorityTitles", g_GetDB("ASYST", , , , bNoTest), True, "")
        aTag = Split(stcAuthStr & ";;;;;;", ";")
        stcAuth2 = "|" & aTag(0) & "|"
        stcAuth3 = "|" & aTag(1) & "|"
        stcAuth4 = "|" & aTag(2) & "|"
        stcAuth5 = "|" & aTag(3) & "|"
        stcAuth6 = "|" & aTag(4) & "|"
    End If
    If bReTest Then stcAuth = ""
    ' If a default auth has been specified, it will have number between 1 and 6...
    If stcAuth = "" Then
        stcAuth = g_GetDB("ASYST", , True, , bNoTest)
''        If Val(stcAuth) = 0 Then stcAuth = "N"
    End If
    If Val(stcAuth) = 0 And stcAuth <> "0" Then
        AST_getUserASystAuthority = 0
        If CBA_getAdminUsers(sUserName) = True Then
            AST_getUserASystAuthority = 1
        Else
            If CBA_getUserShortTitle(sReturnTitle, sReturnShortTitle, sUserName) > -1 Then
                If sReturnShortTitle = "" Then sReturnShortTitle = "x"
                If InStr(1, stcAuth2, "|" & sReturnShortTitle & "|") > 0 Then
                    AST_getUserASystAuthority = 2
                ElseIf InStr(1, stcAuth3, "|" & sReturnShortTitle & "|") > 0 Then
                    AST_getUserASystAuthority = 3
                ElseIf InStr(1, stcAuth4, "|" & sReturnShortTitle & "|") > 0 Then
                    AST_getUserASystAuthority = 4
                ElseIf InStr(1, stcAuth5, "|" & sReturnShortTitle & "|") > 0 Then
                    AST_getUserASystAuthority = 5
                ElseIf InStr(1, stcAuth6, "|" & sReturnShortTitle & "|") > 0 Then
                    AST_getUserASystAuthority = 6
                End If
            End If
        End If
    Else
        AST_getUserASystAuthority = Val(stcAuth)
    End If
    
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_getUserASystAuthority", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
   
End Function

Public Function AST_WriteTable(meCtls As Object, sTbl_Flds As String, sFld_Prefix As String, sTbl_Nam As String, _
                               lFrmID As Long, lAuth As Long, Optional lID As Long = 0) As Boolean
                        
    ' Write to the table from the ASYST form concerned
    Dim frmf As Control, sField As String, lIdx As Long, bfound As Boolean
    Dim aTag() As String, aFrmFlds() As String, aTblFlds() As String, sTag As String
    Dim sSQLFlds As String, sSQLVals As String, vVal, sPref As String, lPGID As Long
    Dim sAuditOld As String, sAuditNew As String, sAuditSQL As String, bAudit As Boolean, lTblID As Long
    Dim CN As ADODB.Connection, RS As ADODB.Recordset, bChgdFld As Boolean, sFmt As String, sAdt As String
    'Static bNot1st As Boolean
    On Error GoTo Err_Routine
    
    AST_WriteTable = False: CBA_ErrTag = ""
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    CBA_ErrTag = ""
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";" ';INTEGRATED SECURITY=sspi;"
    ' Split the name of the Form fields into a private lcase array
    aFrmFlds = Split(Replace(Replace(sTbl_Flds, sFld_Prefix, ""), "_", ""), ",")
    lTblID = NZ(meCtls("txt" & aFrmFlds(0)).Value, 0)                                           ' Get the ID of the table involved
    ' Split the name of the Table fields into a private array
    aTblFlds = Split(sTbl_Flds, ",")
    ' Decide if an add or an update
    If lID = 0 Then
        sSQLFlds = "INSERT INTO " & sTbl_Nam & " ( "
        sSQLVals = vbCrLf & " VALUES("
    Else
        sSQLFlds = "UPDATE " & sTbl_Nam & " SET "
        sSQLVals = vbCrLf & " WHERE " & sFld_Prefix & "ID = " & lID
    End If
    ' Cycle throught the fields on the form to fill the field names and the two Sqls
    For Each frmf In meCtls
        CBA_ErrTag = "No"
        sTag = frmf.Tag
        If (frmf.Tag & "") > "" And InStr(1, ",lbl,cmd,", "," & Left(frmf.Name, 3) & ",") = 0 Then
            sFmt = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "FullFormat")
            sAdt = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "Audit")
            CBA_ErrTag = "Yes"
            bfound = False
            aTag = Split(sTag & CBA_Sn, CBA_S)
            sFmt = Replace(sFmt, "d2", "")          ' remove any numbers in the date to form the base date
            sFmt = Replace(sFmt, "d3", "")
            If aTag(A_FN) = "pgid" Then
                aTag(A_FN) = aTag(A_FN)
                sField = frmf.Name
            End If
            'Debug.Print aTag(a_fn) & ",";
            For lIdx = 0 To UBound(aFrmFlds(), 1)
''               'Debug.Print aFrmFlds(lIdx) & ",";
''                aTag = Split((frmf.Tag & CBA_Sn), CBA_S)
                If aFrmFlds(lIdx) = aTag(A_FN) Then
                    bfound = True
                    
                    If IIf(IsNull(frmf.Value) = True, "", frmf.Value) = "" Then
                        vVal = "NULL"
                    ElseIf InStr(1, ",dtedmyy,", "," & sFmt & ",") > 0 Then
                        vVal = g_GetSQLDate(frmf.Value, CBA_DMY)
                    ElseIf InStr(1, ",dtedmyyhn,", "," & sFmt & ",") > 0 Then
                        vVal = g_GetSQLDate(frmf.Value, CBA_DMYHN)
                    ElseIf Left(sFmt, 3) = "num" Then
                        vVal = g_UnFmt(frmf.Value, "num")
                    ElseIf (frmf.Value = True) And (sFmt = "opt" Or sFmt = "chk") Then
                        vVal = "'Y'"
                    ElseIf (frmf.Value = False) And (sFmt = "opt" Or sFmt = "chk") Then
                        vVal = "'N'"
                    Else
                        vVal = frmf.Value
                    End If
''                ElseIf (aFrmFlds(lIdx) & frmf.Value = aTag(A_FN)) And sFmt = "optyn" Then
''                    bFound = True
''                    vVal = True
                End If
                If bfound = True Then
                    sAuditOld = Replace(NZ(aTag(A_NV), ""), "'", "`")
                    sAuditNew = CStr(Replace(NZ(frmf.Value, ""), "'", "`"))
''                    bChgdFld = ((Right(sAuditOld, 1) = "~") And Len(sAuditOld) > 1)
                    bChgdFld = (Replace(sAuditOld, "~", "") <> Replace(sAuditNew, "~", ""))               ' Has the field changed?????
                    bAudit = ((Right(sAuditOld, 1) = "~") And Len(sAuditOld) > 1 And sAdt = "true")
                    ' Process any differences...
                    If aTag(A_FN) = "LastUpd" Then
                        vVal = g_GetSQLDate(Now(), CBA_DMYHN)
                        bChgdFld = True
                    ElseIf aTag(A_FN) = "UpdUser" Or aTag(A_FN) = "CrtUser" Then
                        vVal = "'" & CBA_User & "'"
                        bChgdFld = True
                    ElseIf sFmt = "txt1" Then
                        vVal = "'" & Left(vVal, 1) & "'"                     ' If the value is 1 char long
                    ElseIf vVal = "NULL" Or vVal = "'Y'" Or vVal = "'N'" Then
                        ' Do nothing on a null
                    ElseIf sFmt = "txt" Then
                        vVal = "'" & Replace(vVal, "'", "`") & "'"
                    End If
                    GoSub GSAdd_Fld_2_SQL
                    GoTo SkipFields
                End If
            Next
SkipFields:
''            If bFound = False Then
''                CBA_ErrTag = "Set"
''                vVal = ""
''            End If
        End If
NoTag:
    Next
    
Exit_Routine:
    If sPref = ", " Then
        CBA_ErrTag = "Write"
        GoSub GSEnd_Of_SQL
        RS.Open sSQLFlds & sSQLVals, CN
        If lID > 0 Then
            lPGID = g_DLookup("PD_PG_ID", "L2_Products", "PD_ID=" & lID, "", g_GetDB("ASYST"), 0)
            ' As the Promotion has been changed, it has to be resent to the regions, so blank the field that will eanble it
            sSQLFlds = "UPDATE L1_Promotions SET PG_Region_Date = NULL WHERE PG_ID = " & lPGID & ";"
            RS.Open sSQLFlds, CN
        End If
    End If
    If AST_WriteTable = True Then
        MsgBox "Unexplained error occurred on save - please inform the ASYST maintanance staff", vbOKOnly
    Else
        MsgBox "Save was successful", vbOKOnly
    End If
    On Error Resume Next
    Set CN = Nothing
    Set RS = Nothing
        
    Exit Function
    
    
GSAdd_Fld_2_SQL:    ' Fill the SQL fields as per the Tag direction
    If lIdx > 0 Then         ' The first in the array is the Table_ID
        If lID = 0 Then                                                             ' (INIT SQL=) "sSQLFlds = "INSERT INTO)   &   sSQLVals = " VALUES("
            sSQLFlds = sSQLFlds & sPref & aTblFlds(lIdx)
            sSQLVals = sSQLVals & sPref & vVal
            sPref = ", "      ' After the first field has been completed, fill in the comma
        Else                                                                        ' (INIT SQL=) "sSQLFlds = "UPDATE " & sTbl_Nam & " SET "
            If bChgdFld Then        ' If field has changed...
                sSQLFlds = sSQLFlds & sPref & aTblFlds(lIdx) & "=" & vVal
                If bAudit Then      ' If field has changed, and is an Audit field...
                    CBA_ErrTag = "Audit"
                    sAuditSQL = "INSERT INTO L1_Audit (PA_Tbl_ID, PA_TF_ID, PA_Field, PA_OldValue, PA_NewValue, PA_CrtUser) " & _
                                " VALUES (" & lTblID & "," & lFrmID & ",'" & Replace(aTblFlds(lIdx), sFld_Prefix, "") & "','" & _
                                    Replace(sAuditOld, "~", "") & "','" & sAuditNew & "','" & CBA_User & "');"
                    RS.Open sAuditSQL, CN
                End If
                sPref = ", "      ' After the first field has been completed, fill in the comma
            End If
        End If
    End If
    Return

GSEnd_Of_SQL:    ' Fill the last ')' and ';' when needed
    If lID = 0 Then
        sSQLFlds = sSQLFlds & ") "
        sSQLVals = sSQLVals & ");"
    Else
        sSQLFlds = sSQLFlds & ", " & sFld_Prefix & "LockFlag='N'"
        sSQLVals = sSQLVals & ";"
    End If
    Return

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_WriteTable", 3)
    CBA_Error = CBA_ErrTag & " Error -" & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "No" Then
        Resume NoTag
    ElseIf CBA_ErrTag = "Set" Then
        Resume Next
    ElseIf CBA_ErrTag = "Write" Then
        AST_WriteTable = True
        CBA_Error = CBA_Error & vbCrLf & sSQLFlds & sSQLVals
        Debug.Print CBA_Error
        Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
        Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
        Resume Next
    ElseIf CBA_ErrTag = "Audit" Then
        CBA_Error = CBA_Error & vbCrLf & sAuditSQL
        Debug.Print CBA_Error
        Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
        Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
        Resume Next
    Else
        Debug.Print CBA_Error
        Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
        Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
        AST_WriteTable = True
        sPref = ""
        GoTo Exit_Routine
    End If
End Function

Public Sub AST_SetLockVis(meCtls As Object, bLockOverRide As Boolean, ByVal lFrmID As Long, ByVal lAuth As Long)        'Optional lVersion As Long = 1
    ' Set each field to Visible and or locked as per the value in the Tag
    
    Dim frmf As Control, sLock As String, sVis As String, bHasClr As Boolean, sClr As String
    Dim bVis As Boolean, bLocked As Boolean, lColour As Long, aTag() As String, sFmt As String, sBaseFmt As String
    Dim sToolTip As String, bIsChgDate As Boolean
    On Error GoTo Err_Routine

    CBA_ErrTag = ""
    ' For each field in the userform, find out it's locked and visable sts
    For Each frmf In meCtls
        CBA_ErrTag = "No"
        If Right(frmf.Name, Len("ApprovalDate")) = "ApprovalDate" And bLockOverRide = False Then
             lColour = lColour
        End If
''        If frmf.Visible = False Then GoTo NextTag          ' If the field is invisible from the off, then ignore it
        If (frmf.Tag & "") > "" Then
            CBA_ErrTag = "Yes"
            sFmt = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "FullFormat")
            sBaseFmt = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "BaseFormat")
            sClr = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "Clr")
            aTag = Split((frmf.Tag & CBA_Sn), CBA_S)
            
            bVis = False: bLocked = False: bHasClr = True: bIsChgDate = True
            sToolTip = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "ToolTip")
''            If sBaseFmt <> "dte" Then
            ' Set the control's tooltip
            frmf.ControlTipText = sToolTip
            If sBaseFmt = "" Or sBaseFmt = "cmd" Or Left(frmf.Name, 3) = "lbl" Then
                GoTo NextTag               ' Don't process CMD keys here as they are done later by themselves
            End If
            ' Set the default colour
            lColour = CBA_White
            ' Process Lock / Visible
            sLock = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "Lock")
            sVis = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "Vis")
            
            If InStr(1, aTag(A_FN), "lasticity") > 0 Then
                aTag(A_FN) = aTag(A_FN)
            End If
            ' Handle the visibility
            If Right(sVis, 5) = "unvis" Then                                ' If Unvis then don't handle the visibility aspect of the field
                GoTo SkipVis
            ElseIf sVis = "invis" Then
                bVis = False
            ElseIf sVis = "vis" Or sVis = "" Then
                bVis = True
            ElseIf sProcNameSSTagAuth(sVis, lAuth) = (lAuth & "vis") Then
                bVis = True
            ElseIf sProcNameSSTagAuth(sVis, lAuth) = (lAuth & "invis") Then
                bVis = False
            ElseIf bLockOverRide = True Then
                bVis = False
            Else
               bVis = False
            End If
            ' Set the visable as per we have worked out
            frmf.Visible = bVis
            ' Don't do anything else if now invisible
            If bVis = False Then GoTo NextTag
SkipVis:
            ' Process locked / unlocked / clrs
            If bLockOverRide = True Then
                bLocked = True
                lColour = CBA_Grey
            ElseIf sLock = "unlock" Then
                bLocked = False
                lColour = CBA_White
            ElseIf sLock = "lock" Then
                bLocked = True
                lColour = CBA_Grey
''            ElseIf bLockOverRide = True Then              ' Why was this here?????????????????????????????????
''                bLocked = True
''                lColour = CBA_Grey
            ElseIf sProcNameSSTagAuth(sLock, lAuth) = (lAuth & "lock") Then
                bLocked = True
                lColour = CBA_Grey
            ElseIf sProcNameSSTagAuth(sLock, lAuth) = (lAuth & "unlock") Then
                bLocked = False
            ElseIf sProcNameSSTagAuth(sLock, lAuth) = (lAuth & "alock") Then        ' Add it as an alock if a date is to be changed when clicked
                bLocked = True
                lColour = CBA_White
            Else
                bLocked = False
                lColour = CBA_White
            End If
            ' Set the locked sts as per worked out
            frmf.Locked = bLocked
            Select Case sBaseFmt
                Case "opt"
                    ''frmf.ForeColor = lColour
                Case "chk"
                    ''frmf.ForeColor = lColour
                    If bLocked = True Then
                        frmf.Enabled = False
                    Else
                        frmf.Enabled = True
                    End If
                Case Else
                    If lColour = CBA_Grey And sClr <> "d" And sClr <> "n" Then
                        frmf.BackColor = lColour
                    Else
                        Select Case sClr
                        Case Is = "y"
                            Call AST_FillYellow(frmf, "y")
                        Case Is = "n", "d"
                        Case Is = "o"
                            frmf.BackColor = CBA_OffYellow
                        Case Is = "g"
                            frmf.BackColor = CBA_Grey
                        Case Is = "p"
                            frmf.BackColor = CBA_Pink
                        Case Else
                            frmf.BackColor = lColour
                        End Select
                    End If
            End Select
        'frmf.Top = 10
        End If
NextTag:
''    If Left(frmf.Name, 3) <> "lbl" And bLocked = False And bVis = True Then
''       'Debug.Print frmf.Name & "," & lAuth & "," & IIf(bLocked, "lock", "unL") & "," & IIf(bVis, "vis", "inV")
''    End If
    If frmf.Name = "txtFutureProdCode" Then
        lAuth = lAuth
    End If
    Next
Exit_Routine:
    Exit Sub
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-AST_SetLockVis", 3)
    CBA_Error = frmf.Name & "-" & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "No" Then
        Resume NextTag
    Else
        Debug.Print CBA_Error
        Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
        Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
        GoTo Exit_Routine
        Resume Next
    End If
End Sub


Public Sub AST_FillDDBox(cbo As ComboBox, ByVal lCols As Long, Optional vVal As Variant = "", Optional bClear As Boolean = True)
    ' Will fill the specified DD Box with data from the array
    Dim lRow As Long, lcol As Long, lNewVal As Long, bHiVal As Boolean
    On Error GoTo Err_Routine
    If bClear Then cbo.Clear
    ' Is it a hival, if so then capture the highest value as it goes though and position on it
    bHiVal = (Val(vVal) = CBA_LongHiVal)
    ' For each row in the array...
    For lRow = 0 To UBound(CBA_ABIarr, 2)
        'Debug.Print CBA_ABIarr(0, lRow) & " " & CBA_ABIarr(1, lRow) & " ";
        cbo.AddItem
        For lcol = 0 To lCols - 1
            cbo.List(lRow, lcol) = NZ(CBA_ABIarr(lcol, lRow), "")
            If lcol = 0 And bHiVal Then
                If Val(CBA_ABIarr(0, lRow)) > lNewVal Then lNewVal = Val(CBA_ABIarr(0, lRow))
            End If
        Next
    Next
    If CStr(vVal) <> "" Then
        cbo.Value = Null
        If bHiVal Then vVal = lNewVal
        cbo.Value = vVal
    End If
Exit_Routine:

    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-AST_FillDDBox", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Public Sub AST_FillListBox(lst As Object, lCols As Long, Optional vVal As Variant = "", Optional bClear As Boolean = True)
    ' Will fill the specified List Box with data from the array
    ' NOTE - IF THE LISTBOX DOESN'T POSITION PROPERLY AFTER BEING CALLED,
    '        IT MAY BE THAT YOU NEED TO ENSURE THAT IT'S CHANGE ROUTINE ISN'T PULLED WHEN IT IS REPOSITIONED!!!!!!!
    Dim lRow As Long, lcol As Long, lNewVal As Variant, bHiVal As Boolean, lIdx As Long
    On Error GoTo Err_Routine
    
    If bClear Then lst.Clear
    ' Is it a hival?, if so then capture the highest value as it goes though and position on it
    bHiVal = (Val(vVal) = CBA_LongHiVal)
    lIdx = -1
    ' For each row in the array...
    For lRow = 0 To UBound(CBA_ABIarr, 2)
        'Debug.Print CBA_ABIarr(0, lRow) & " " & CBA_ABIarr(1, lRow) & " ";
        For lcol = 0 To lCols - 1
            If lcol = 0 Then lst.AddItem NZ(CBA_ABIarr(0, lRow), "")
            If lcol = 0 Then
                If bHiVal Then
                    If Val(CBA_ABIarr(0, lRow)) > lNewVal Then
                        lNewVal = Val(CBA_ABIarr(0, lRow))
''                        lIdx = lst.ListIndex
                    End If
                ElseIf CStr(vVal) <> "" Then
                    If CBA_ABIarr(0, lRow) = vVal Then
                        lNewVal = Val(CBA_ABIarr(0, lRow))
                    End If
                End If
            End If
            lst.List(lRow, lcol) = NZ(CBA_ABIarr(lcol, lRow), "")
        Next
    Next
    On Error Resume Next
    If CStr(vVal) <> "" Then
''        lst.SetFocus
        lst.Value = Null                                    ' Set it to null in case you are setting it to the same value as it doesn't take otherwise
        vVal = lNewVal
        lst.Value = lNewVal
    End If
    DoEvents
Exit_Routine:

    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-AST_FillListBox", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Public Sub AST_setEndDate(sFormName As String, meCtls As Object, cObj As Control, sFormat As String)
    ' Will set the EndDate for whatever ASyst form
    Dim sOnSaleDate As String, sNamePrefix As String, sOnSaleDateName As String, sWeeksOfSaleName As String
    Dim sEndDateName As String, lWeeksOfSale As Long
    
'    If g_SetupIP(sFormName) = False Then
        If g_GetNo(cObj.Name) > 0 Then                      ' If there is a number in the name....
            sNamePrefix = Left(cObj.Name, 5)                ' Use 5 chars of name - e.g. 'txt01'
        Else
            sNamePrefix = Left(cObj.Name, 3)                ' Use 3 chars of name - e.g. 'txt'
        End If
        If sNamePrefix = "cbo" Then sNamePrefix = "txt"
        sOnSaleDateName = sNamePrefix & "OnSaleDate"
        sWeeksOfSaleName = "cboWeeksOfSale"
        sEndDateName = sNamePrefix & "EndDate"
               
        sOnSaleDate = g_FixDate(meCtls(sOnSaleDateName).Value)
        lWeeksOfSale = Val(meCtls(sWeeksOfSaleName).Value)
        ' EndDate is set by OnSaleDate + (WeeksOfSale * 7) or is null
        If Not g_IsDate(sOnSaleDate) Then sOnSaleDate = ""
        If g_IsDate(sOnSaleDate) Then
            meCtls(sEndDateName).Value = g_FixDate(CDate(sOnSaleDate) + (lWeeksOfSale * 7) - 1, sFormat)
        Else
            meCtls(sEndDateName).Value = ""
        End If
'    End If
End Sub

Public Sub AST_setEndDate1(sFormName As String, OSDt As MSForms.TextBox, WoS As MSForms.TextBox, EDt As MSForms.TextBox, sFormat As String)
    ' Will set the EndDate for whatever ASyst form  ' @RWAST New Proc - change everything over to and delete original
    Dim sOnSaleDate As String, lWeeksOfSale As Long
    If g_SetupIP(sFormName) = False Then
        sOnSaleDate = g_FixDate(OSDt.Value)
        lWeeksOfSale = Val(WoS.Value)
        ' EndDate is set by OnSaleDate + (WeeksOfSale * 7) or is null
        If Not g_IsDate(sOnSaleDate) Then sOnSaleDate = ""
        If g_IsDate(sOnSaleDate) Then
            EDt.Value = g_FixDate(CDate(sOnSaleDate) + (lWeeksOfSale * 7) - 1, sFormat)
        Else
            EDt.Value = ""
        End If
    End If
End Sub

Public Function AST_TF(FldNam_No, Optional sTable_Flds As String, Optional sPrefix As String) As Variant
     ' Fill this in the ProductRows forms to capture the Field positions in the array
     ' To bring back the Field Name at the position x, enter the number i.e. 0 = the first element (Field Name)
     ' To bring back the number of the element, enter the Field/Tag Name i.e. ProductCode = 0
     
    Static paFF() As String, bHasRun As Boolean
    Dim lIdx As Long
    On Error GoTo Err_Routine
    If bHasRun = False Or sTable_Flds > "" Then
        paFF = Split(Replace(Replace(sTable_Flds, sPrefix, ""), "_", ""), ",")
        bHasRun = True
    End If
    AST_TF = -1
    If IsNumeric(FldNam_No) Then
        AST_TF = paFF(FldNam_No)
        Exit Function
    Else
        For lIdx = 0 To UBound(paFF, 1)
            If FldNam_No = paFF(lIdx) Then
                AST_TF = lIdx
                Exit Function
            End If
        Next
    End If
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("AST_TF", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function AST_NumberOfStoresCompared(ByVal dtPriorDate As Date, ByVal dtToDate As Date) As Single
    ' Get the number of stores that existed at a Prior time and compare it to now
    ' This can then be used to estimate sales - e.g if the compare is 1.3 then X the sales then, by 1.3 to get est sales now
    Dim bOutput As Boolean, bfound As Boolean, lIdx As Long
    Static aDate() As Variant, dtDate2 As Date, lStores2 As Long, lMaxIdx As Long
    On Error GoTo Err_Routine
    
    AST_NumberOfStoresCompared = 1
    
    ' Set the date back to the beginning of the week...
    dtPriorDate = dtPriorDate - WeekDay(dtPriorDate, vbUseSystem) + 1: bfound = False
    ' See if the date and no of stores has already been captured
    If lMaxIdx > 0 Then
        For lIdx = 0 To lMaxIdx - 1
            If aDate(0, lIdx) = dtPriorDate Then
                bfound = True
                Exit For
            End If
        Next
    End If
    ' If not found then add it to the array
    If Not bfound Then
        lMaxIdx = lMaxIdx + 1
        lIdx = lMaxIdx - 1
        ReDim Preserve aDate(0 To 1, 0 To (lMaxIdx - 1))
        aDate(0, lIdx) = dtPriorDate
        bOutput = CBA_COM_MATCHGenPullSQL("PORT_NOSTORE", , dtPriorDate)
        If bOutput = True Then
            aDate(1, lIdx) = NZ(CBA_CBISarr(0, 0), 0)
        Else
            aDate(1, lIdx) = 0
        End If
    End If
    
    ' If the now date has not been done yet
    If Not g_IsDate(dtDate2) Then
        If Year(dtToDate) < Year(Date) Then dtToDate = Date
        dtDate2 = dtToDate - WeekDay(dtToDate, vbUseSystem) + 1
        bOutput = CBA_COM_MATCHGenPullSQL("PORT_NOSTORE", , dtDate2)
        If bOutput = True Then
            lStores2 = NZ(CBA_CBISarr(0, 0), 0)
        Else
            lStores2 = 0
        End If
    End If
    If lStores2 > 0 And aDate(1, lIdx) > 0 Then
        AST_NumberOfStoresCompared = (lStores2 / aDate(1, lIdx))
    Else
        AST_NumberOfStoresCompared = 1
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_NumberOfStoresCompared", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function AST_getCompareSalesCBis(ByVal sdtFromDt As String, ByVal sdtToDt As String, ByVal lPcode As Long, _
                            ByVal Prior_Now As String, Optional sFmt As String = "#,0", Optional sngMult As Single = 1) As String
    ' Get the POS between two dates a year since - Use this to Estimate the Comparative POS between then and now - CBIS...
    Dim bOutput As Boolean, lPosQty As Long, lIdx As Long, dtFromDt As Date, dtToDt As Date
    Static lPriorQty As Long, sngUplift As Single
    On Error GoTo Err_Routine
    dtFromDt = CDate(g_FixDate(sdtFromDt))
    dtToDt = CDate(g_FixDate(sdtToDt))

    If Prior_Now = "Prior" Then
        CBA_DBtoQuery = 599
        bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBIS_POSDATA", dtFromDt, dtToDt, lPcode)
        If bOutput Then
            For lIdx = 0 To UBound(CBA_CBISarr, 2)
                lPosQty = lPosQty + CBA_CBISarr(1, lIdx)
                CBA_CBISarr(2, lIdx) = CBA_CBISarr(2, lIdx)
            Next
        Else
            lPosQty = 0
        End If
        AST_getCompareSalesCBis = Format(lPosQty, sFmt)
        sngUplift = AST_NumberOfStoresCompared(dtFromDt, Date)
        lPriorQty = lPosQty
    Else
        AST_getCompareSalesCBis = Format((lPriorQty * sngUplift * sngMult), sFmt)
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_getCompareSalesCBis", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function AST_CrtProdRegionRecs(ByVal lProductID As Long, Optional bDeleteFirst As Boolean = False) As String
    ' Get the POS for each region and create a record in the database for each region / product
    Dim lIdx As Long, sSQL As String, aRegs() As String, sReg As String
    Dim dtFromDt As Date, dtToDt As Date, lPcode As Long, lPDID As Long, lPVID As Long, bPassOK As Boolean ', sngMult As Single
    Dim CN As ADODB.Connection, RS As ADODB.Recordset, RSI As ADODB.Recordset
    Dim curPDUnitCost As Currency, curPDSuppCostSupp As Currency, lPDPrior_Sales As Long, lPDExp_Sales As Long, lPDCalc_Sales As Long
    Dim lPD_UPSPW As Long, lPDFill_Qty As Long, curPDCurr_Retail_Price As Currency, curPDRetail_Price As Currency, sngPDOrig_Multiplier As Single
    Dim sngRetail As Single, lRecs As Long, sngQty As Single
    Dim sngUPSPW As Single, sngRPSPW As Single, sngPCDiv As Single, lPV_Prior_Sales As Long, lPV_Est_Sales As Long, lPV_Calc_Sales As Long ', lValidDivs As Long '', sngRPSPW As Single
    
    Const a_Regs As String = "MIN,DER,STP,PRE,DAN,BRE,RGY,xxx,JKT"
    On Error GoTo Err_Routine
    
    CBA_ErrTag = ""
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    Set RSI = New ADODB.Recordset
    ' Split the Regions
    aRegs = Split(a_Regs, ",")
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";"
    CBA_ErrTag = "SQL"
  
    ' Delete the records first and then recreate
    If bDeleteFirst Then
        sSQL = "DELETE FROM L3_ProductRegions WHERE PV_PD_ID = " & lProductID
        RS.Open sSQL, CN
        ''RS.Close
        Set RS = New ADODB.Recordset
    End If
' Get the product concerned...
    sSQL = "SELECT * FROM qry_L3_Products_v_Regions WHERE PD_ID=" & lProductID
    RSI.Open sSQL, CN
    Do While Not RSI.EOF
        lPVID = NZ(RSI!PV_ID, 0)
        If lPVID = 0 Then
            ' Take the default values from the Product table
            dtFromDt = DateAdd("m", -12, RSI!PD_On_Sale_Date)
            dtToDt = DateAdd("m", -12, RSI!PD_End_Date)
            lPcode = RSI!PD_Product_Code
            lPDID = RSI!PD_ID
            curPDUnitCost = NZ(RSI!PD_Unit_Cost, 0)
            curPDSuppCostSupp = NZ(RSI!PD_Supplier_Cost_Support, 0)
            lPDPrior_Sales = NZ(RSI!PD_Prior_Sales, 0)
            lPDExp_Sales = NZ(RSI!PD_Expected_Sales, 0)
            lPDCalc_Sales = NZ(RSI!PD_Calculated_Sales, 0)
            lPD_UPSPW = NZ(RSI!PD_UPSPW, 0)
            lPDFill_Qty = NZ(RSI!PD_Fill_Qty, 0)
            curPDCurr_Retail_Price = NZ(RSI!PD_Curr_Retail_Price, 0)
            curPDRetail_Price = NZ(RSI!PD_Retail_Price, 0)
            sngPDOrig_Multiplier = NZ(RSI!PD_Sales_Multiplier, 0)
            ' Accumulate and position as per the regions
            CBA_ErrTag = "Data"
            For lIdx = 501 To 509
                ' Get the region short name
                sReg = aRegs(lIdx - 501)
                If lIdx = 501 Then
                    CBA_DBtoQuery = lIdx
                    Call AST_GetEstTierData(lPcode, 0, 0)
                    sngRetail = 0
                    For lRecs = 0 To UBound(CBA_CBISarr, 2)
                        If UBound(CBA_CBISarr, 1) > 2 Then
                            sngRetail = sngRetail + NZ(CBA_CBISarr(3, lRecs), 0)
                        End If
                    Next
                End If
                If lIdx <> 508 Then
                    ' Get the comparative number of stores
                    sngUPSPW = 0: sngRPSPW = 0: sngQty = 0
                    If bPassOK = True Then
                        For lRecs = 0 To UBound(CBA_CBISarr, 2)
                            If lIdx = NZ(CBA_CBISarr(0, lRecs), 0) Then
                                If UBound(CBA_CBISarr, 1) > 2 Then
                                    sngUPSPW = sngUPSPW + NZ(CBA_CBISarr(2, lRecs), 0)
                                    sngRPSPW = sngRPSPW + NZ(CBA_CBISarr(3, lRecs), 0)
                                    sngQty = sngQty + NZ(CBA_CBISarr(4, lRecs), 0)
                                    ''lValidDivs = lValidDivs + 1
                                End If
                            End If
                        Next
                        sngPCDiv = sngRPSPW / sngRetail
                        lPV_Prior_Sales = lPDPrior_Sales * sngPCDiv
                        lPV_Est_Sales = Round(lPDExp_Sales * sngPCDiv, 0)
                        lPV_Calc_Sales = Round(lPDCalc_Sales * sngPCDiv, 0)
                    End If
                    ' Add the record to the table if it doesn't exist
                    If lPVID = 0 Then
                        CBA_ErrTag = "SQL"
                        sSQL = ""
                        sSQL = sSQL & "INSERT INTO L3_ProductRegions ( PV_PD_ID,PV_Region,PV_Unit_Cost,PV_Supplier_Cost_Support,PV_Prior_Sales,PV_OrigEstSales," & _
                                      "PV_OrigCalcSales,PV_UPSPW,PV_Fill_Qty,PV_Curr_Retail_Price,PV_Retail_Price,PV_OrigEstMultiplier )"
                        sSQL = sSQL & " VALUES ( " & lPDID & "," & lIdx & "," & curPDUnitCost & "," & curPDSuppCostSupp & "," & lPV_Prior_Sales & "," & lPV_Est_Sales & _
                                       "," & lPV_Calc_Sales & "," & lPD_UPSPW & "," & lPDFill_Qty & "," & curPDCurr_Retail_Price & "," & curPDRetail_Price & "," & sngPDOrig_Multiplier & " )"
                    Else
                    End If
                    RS.Open sSQL, CN
                End If
    
            Next
        End If
        RSI.MoveNext
    Loop
Exit_Routine:
    On Error Resume Next
    Set RSI = Nothing
    Set RS = Nothing
    Set CN = Nothing
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_CrtProdRegionRecs", 3)
    CBA_Error = " Error -" & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function AST_CrtProdRegionClass(ByVal lProductID As Long, Optional bRecreate As Boolean = False) As String
''    ' Get the POS for each region and create a record in the database for each region / product
''    Dim lIdx As Long, sSQL As String, aRegs() As String, sReg As String
''    Dim dtFromDt As Date, dtToDt As Date, lPCode As Long, lPDID As Long, lPVID As Long, bPassOK As Boolean ', sngMult As Single
''    Dim CN As ADODB.Connection, RS As ADODB.Recordset, RSI As ADODB.Recordset
''    Dim Unit_Cost As Currency, Supplier_Cost_Support As Currency, Prior_Sales As Long, OrigEstSales As Long, OrigCalcSales As Long
''    Dim UPSPW As Long, Fill_Qty As Long, Curr_Retail_Price As Currency, Retail_Price, EstMultiplier
''    Dim sngRetail As Single, lRecs As Long, sngQty As Single
''    Dim sngUPSPW As Single, sngRPSPW As Single, sngPCDiv As Single, Prior_SalesDiv As Long, EstSalesDiv As Long, CalcSalesDiv As Long ', lValidDivs As Long '', sngRPSPW As Single
    
    
    
    
    Const a_Regs As String = "Min,Der,Stp,Pre,Dan,Bre,Rgy,xxx,Jkt"
    
    
    
    
    
    On Error GoTo Err_Routine
    
    
    If cls_Prod Is Nothing Then
        Set cls_Prod = New CBA_AST_Product
        cls_Prod.formulate lProductID
    Else
        If cls_Prod.plPDID <> lProductID Then
            Set cls_Prod = Nothing
            Set cls_Prod = New CBA_AST_Product
            cls_Prod.formulate lProductID
        End If
    
    End If
    
    
    
    
    
    
    
    
    
    
'    CBA_ErrTag = ""
'    Set CN = New ADODB.Connection
'    Set RS = New ADODB.Recordset
'    Set RSI = New ADODB.Recordset
'    ' Split the Regions
'    aRegs = Split(a_Regs, ",")
'    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_getDB("ASYST") & ";"
'    CBA_ErrTag = "SQL"
'    ' If the data is being recreated...
''    If bRecreate Then
''        sSQL = "DELETE FROM L3_ProductRegions WHERE PV_PD_ID = " & lProductID
''        RS.Open sSQL, CN
''        If RS.State = 1 Then RS.Close
''        Set RS = New ADODB.Recordset
''    End If
'    ' Get the product concerned...
'    sSQL = "SELECT * FROM qry_L3_Products_v_Regions WHERE PD_ID=" & lProductID
'    RSI.Open sSQL, CN
'    Do While Not RSI.EOF
'        lPVID = NZ(RSI!PV_ID, 0)
'        If lPVID = 0 Then
'            ' Take the default values from the Product table
'            dtFromDt = DateAdd("m", -12, RSI!PD_On_Sale_Date)
'            dtToDt = DateAdd("m", -12, RSI!PD_End_Date)
'            lPCode = RSI!PD_Product_Code
'            lPDID = RSI!PD_ID
'            Unit_Cost = NZ(RSI!PD_Unit_Cost, 0)
'            Supplier_Cost_Support = NZ(RSI!PD_Supplier_Cost_Support, 0)
'            Prior_Sales = NZ(RSI!PD_Prior_Sales, 0)
'            OrigEstSales = NZ(RSI!PD_Expected_Sales, 0)
'            OrigCalcSales = NZ(RSI!PD_Calculated_Sales, 0)
'            UPSPW = NZ(RSI!PD_UPSPW, 0)
'            Fill_Qty = NZ(RSI!PD_Fill_Qty, 0)
'            Curr_Retail_Price = NZ(RSI!PD_Curr_Retail_Price, 0)
'            Retail_Price = NZ(RSI!PD_Retail_Price, 0)
'            EstMultiplier = NZ(RSI!PD_Sales_Multiplier, 0)
'            ' Accumulate and position as per the regions
'            CBA_ErrTag = "Data"
'            For lIdx = 501 To 509
'                ' Get the region short name
'                sReg = aRegs(lIdx - 501)
'                If lIdx = 501 Then
'                    CBA_DBtoQuery = lIdx
'                    Call AST_GetEstTierData(lPCode, 0, bPassOK)
'                    sngRetail = 0
'                    For lRecs = 0 To UBound(CBA_CBISarr, 2)
'                        If UBound(CBA_CBISarr, 1) > 2 Then
'                            sngRetail = sngRetail + NZ(CBA_CBISarr(3, lRecs), 0)
'                        End If
'                    Next
'                End If
'                If lIdx <> 508 Then
'                    ' Get the comparative number of stores
'                    sngUPSPW = 0: sngRPSPW = 0: sngQty = 0
'                    If bPassOK = True Then
'                        For lRecs = 0 To UBound(CBA_CBISarr, 2)
'                            If lIdx = NZ(CBA_CBISarr(0, lRecs), 0) Then
'                                If UBound(CBA_CBISarr, 1) > 2 Then
'                                    sngUPSPW = sngUPSPW + NZ(CBA_CBISarr(2, lRecs), 0)
'                                    sngRPSPW = sngRPSPW + NZ(CBA_CBISarr(3, lRecs), 0)
'                                    sngQty = sngQty + NZ(CBA_CBISarr(4, lRecs), 0)
'                                    ''lValidDivs = lValidDivs + 1
'                                End If
'                            End If
'                        Next
'                        sngPCDiv = sngRPSPW / sngRetail
'                        Prior_SalesDiv = Prior_Sales * sngPCDiv
'                        EstSalesDiv = Round(OrigEstSales * sngPCDiv, 0)
'                        CalcSalesDiv = Round(OrigCalcSales * sngPCDiv, 0)
'                    End If
'                    ' Add the record to the table if it doesn't exist
'                    If lPVID = 0 Then
'                        CBA_ErrTag = "SQL"
'                        sSQL = ""
'                        sSQL = sSQL & "INSERT INTO L3_ProductRegions ( PV_PD_ID,PV_Region,PV_Unit_Cost,PV_Supplier_Cost_Support,PV_Prior_Sales,PV_OrigEstSales," & _
'                                      "PV_OrigCalcSales,PV_UPSPW,PV_Fill_Qty,PV_Curr_Retail_Price,PV_Retail_Price,PV_OrigEstMultiplier )"
'                        sSQL = sSQL & " VALUES ( " & lPDID & "," & lIdx & "," & Unit_Cost & "," & Supplier_Cost_Support & "," & Prior_SalesDiv & "," & EstSalesDiv & _
'                                       "," & CalcSalesDiv & "," & UPSPW & "," & Fill_Qty & "," & Curr_Retail_Price & "," & Retail_Price & "," & EstMultiplier & " )"
'                    Else
'                    End If
'                    RS.Open sSQL, CN
'                End If
'
'            Next
'        End If
'        RSI.MoveNext
'    Loop
Exit_Routine:
    On Error Resume Next
''    Set RSI = Nothing
''    Set RS = Nothing
''    Set CN = Nothing
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_CrtProdRegionClass", 3)
    CBA_Error = " Error -" & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function AST_AveProdRegionRecs(ByVal lProductID As Long, sRegions As String) As String
    ' Average the regions and refill the Product table with the values
    Dim sSQL As String     ',lIdx As Long, aRegs() As String, sReg As String
''    Dim dtFromDt As Date, dtToDt As Date, lPCode As Long, lPDID As Long, lPVID As Long, bPassOK As Boolean ', sngMult As Single
    Dim CN As ADODB.Connection, RS As ADODB.Recordset, RSI As ADODB.Recordset
    Dim Unit_Cost As Currency, Supplier_Cost_Support As Currency, Prior_Sales As Long, OrigEstSales As Long, OrigCalcSales As Long, Retail_Discount As Single
    Dim UPSPW As Long, Fill_Qty As Long, Curr_Retail_Price As Currency, Retail_Price, EstMultiplier, lPGID As Long, lNoRegions As Long, sRegion As String
    On Error GoTo Err_Routine
    
    CBA_ErrTag = ""
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    Set RSI = New ADODB.Recordset

    ' Split the Regions
''    aRegs = Split(a_Regs, ",")
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";"
    CBA_ErrTag = "SQL"
    Unit_Cost = 0: Supplier_Cost_Support = 0: Prior_Sales = 0: OrigEstSales = 0: OrigCalcSales = 0: UPSPW = 0
    Fill_Qty = 0: Curr_Retail_Price = 0: Retail_Price = 0: EstMultiplier = 0
    ' Get the product records concerned...
    sSQL = "SELECT * FROM L3_ProductRegions WHERE PV_PD_ID=" & lProductID
    RSI.Open sSQL, CN
    Do While Not RSI.EOF
        CBA_ErrTag = "Data"
        If lPGID = 0 Then
            lPGID = g_DLookup("PD_PG_ID", "L2_Products", "PD_ID=" & lProductID, "", g_GetDB("ASYST"), 0) ''NZ(RSI!PV_PG_ID, 0)
        End If
        ' Accumulate the values
        Unit_Cost = Unit_Cost + NZ(RSI!PV_Unit_Cost, 0)
        Supplier_Cost_Support = Supplier_Cost_Support + NZ(RSI!PV_Supplier_Cost_Support, 0)
        Prior_Sales = Prior_Sales + NZ(RSI!PV_Prior_Sales, 0)
        OrigEstSales = OrigEstSales + NZ(RSI!PV_OrigEstSales, 0)
        OrigCalcSales = OrigCalcSales + NZ(RSI!PV_OrigCalcSales, 0)
        UPSPW = UPSPW + NZ(RSI!PV_UPSPW, 0)
        Fill_Qty = Fill_Qty + NZ(RSI!PV_Fill_Qty, 0)
        Curr_Retail_Price = Curr_Retail_Price + NZ(RSI!PV_Curr_Retail_Price, 0)
        Retail_Price = Retail_Price + NZ(RSI!PV_Retail_Price, 0)
        EstMultiplier = EstMultiplier + NZ(RSI!PV_OrigEstMultiplier, 0)
        RSI.MoveNext
    Loop
    ' Average the values
    Unit_Cost = Round(Unit_Cost / 8, 2)
    Supplier_Cost_Support = Round(Supplier_Cost_Support / 8, 2)
    Prior_Sales = Round(Prior_Sales, 0)
    OrigEstSales = Round(OrigEstSales, 0)
    OrigCalcSales = Round(OrigCalcSales, 0)
    UPSPW = Round(UPSPW / 8, 0)
    Fill_Qty = Round(Fill_Qty / 8, 0)
    Curr_Retail_Price = Round(Curr_Retail_Price / 8, 2)
    Retail_Price = Round(Retail_Price / 8, 2)
    EstMultiplier = Round(EstMultiplier / 8, 3)
    If Curr_Retail_Price > 0 Then
        Retail_Discount = 100 * (Round(((Curr_Retail_Price - Retail_Price) / Curr_Retail_Price), 3))
    Else
        Retail_Discount = 0
    End If
    ' Upd the product record
    CBA_ErrTag = "SQL"
    sSQL = ""
    sSQL = sSQL & "UPDATE L2_Products SET PD_Regions='" & sRegions & "',PD_Unit_Cost=" & Unit_Cost & ",PD_Supplier_Cost_Support=" & Supplier_Cost_Support & ",PD_Prior_Sales=" & Prior_Sales & _
                    ",PD_Expected_Sales=" & OrigEstSales & ",PD_Calculated_Sales=" & OrigCalcSales & ",PD_UPSPW=" & UPSPW & ",PD_Fill_Qty=" & Fill_Qty & _
                    ",PD_Curr_Retail_Price=" & Curr_Retail_Price & ",PD_Retail_Price=" & Retail_Price & ",PD_Sales_Multiplier=" & EstMultiplier & ",PD_Retail_Discount=" & Retail_Discount & _
                    " WHERE PD_ID=" & lProductID
    RS.Open sSQL, CN
    ' As the Promotion has been changed, it has to be resent to the regions, so blank the field that will eanble it
    sSQL = "UPDATE L1_Promotions SET PG_Region_Date = NULL WHERE PG_ID = " & lPGID & ";"
    RS.Open sSQL, CN
    
    
Exit_Routine:
    On Error Resume Next
    ''RS.Close
    Set RSI = Nothing
    Set RS = Nothing
    Set CN = Nothing
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_AveProdRegionRecs", 3)
    CBA_Error = " Error -" & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function AST_GetEstTierData(ByVal lPcode As Long, lCG As Long, Optional bPassOK As Boolean = False) As Long
    ' Get the Estimated Tier Data for a new product - also get the Units per Store Per Week
    Dim dtFromDt As Date, dtToDt As Date, lMinQty As Long, sSQL As String, sFreshFood As String
    Dim lRecs  As Long, lValidDivs As Long, sngUPSPW As Single, sngRPSPW As Single
    On Error GoTo Err_Routine

    ' Get the list of 'Tier1 Products by Default'
    sFreshFood = g_getConfig("Tier1Prods", g_GetDB("ASYST"), True, "")
    If InStr(1, sFreshFood, "," & CStr(lCG) & ",") > 0 Then
        AST_GetEstTierData = 1
    Else
        AST_GetEstTierData = 2
    End If
    
    
    
    
    
    
    
    
'        dtToDt = Date - 3
'        dtFromDt = dtToDt - 7
'        lUPSPW = 0: lValidDivs = 0
'        CBA_ErrTag = "SQL"
'        ' SQL for Units per store per week
'        sSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
'        sSQL = sSQL & "declare @PROD int = " & lPCode & Chr(10)
'        sSQL = sSQL & "declare @SDATE date = '" & Format(dtFromDt, "YYYY-MM-DD") & "'" & Chr(10)
'        sSQL = sSQL & "declare @EDATE date = '" & Format(dtToDt, "YYYY-MM-DD") & "'" & Chr(10)
'
'        sSQL = sSQL & "select productcode INTO #P from cbis599p.dbo.product " & Chr(10)
'        sSQL = sSQL & "where productcode in (@PROD) " & Chr(10)
'        sSQL = sSQL & "select ss.posdate, ss.divno, ss.productcode, isnull(ss.retail,0) as retail, isnull(ss.quantity,0) as quantity " & Chr(10)
'        sSQL = sSQL & "into #SALES " & Chr(10)
'        sSQL = sSQL & "from cbis599p.dbo.pos ss " & Chr(10)
'        sSQL = sSQL & "inner join #P as p on p.ProductCode = ss.ProductCode " & Chr(10)
'        sSQL = sSQL & "where posDate >= @SDATE " & Chr(10)
'        sSQL = sSQL & "and posdate <= @EDATE " & Chr(10)
'        sSQL = sSQL & "select * into #PDATA from ( " & Chr(10)
'        sSQL = sSQL & "select ss.divno, datepart(iso_week,posdate) as IsoWk, YEAR(DATEADD(day, 26 - DATEPART(ISOWW, posdate), posdate)) as IsoYr " & Chr(10)
'        sSQL = sSQL & ",productcode, s.NoOfStores as storecount, sum(quantity) as QTY, sum(Retail) as Ret, count(distinct posdate) as daycount " & Chr(10)
'        sSQL = sSQL & "from #SALES ss " & Chr(10)
'        sSQL = sSQL & "left join cbis599p.Portfolio.Stores s on datepart(iso_week,s.Validfrom) = datepart(iso_week,posdate) " & Chr(10)
'        sSQL = sSQL & "and YEAR(DATEADD(day, 26 - DATEPART(ISOWW, posdate), posdate)) =YEAR(DATEADD(day, 26 - DATEPART(ISOWW, s.Validfrom), s.Validfrom)) and s.divno = ss.divno " & Chr(10)
'        sSQL = sSQL & "group by ss.divno, datepart(iso_week,posdate) , YEAR(DATEADD(day, 26 - DATEPART(ISOWW, posdate), posdate)), productcode, s.NoOfStores " & Chr(10)
'        sSQL = sSQL & ") a " & Chr(10)
'        sSQL = sSQL & "where a.daycount > 2 " & Chr(10)
'        sSQL = sSQL & "order by productcode,IsoYr, IsoWk " & Chr(10)
'        sSQL = sSQL & "select divno, productcode,  avg(QTY / nullif(storecount,0)) as USW, avg(Ret / nullif(storecount,0)) as RSW, count(productcode) as wkCnt " & Chr(10)
'        sSQL = sSQL & "from #PDATA " & Chr(10)
'        sSQL = sSQL & "group by divno,productcode " & Chr(10)
'        sSQL = sSQL & "drop table  #P, #SALES, #PDATA " & Chr(10)
'        ' Get the CBis data
'        CBA_DBtoQuery = 599
'        bPassOK = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CBIS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", sSQL, 120, , , False)
'        If bPassOK = True Then
'            For lRecs = 0 To UBound(CBA_CBISarr, 2)
'                CBA_ErrTag = ""
'                sngUPSPW = sngUPSPW + NZ(CBA_CBISarr(2, lRecs), 0)
'                sngRPSPW = sngRPSPW + NZ(CBA_CBISarr(3, lRecs), 0)
'                lValidDivs = lValidDivs + 1
'            Next
'        Else
'            sngUPSPW = 0
'            lValidDivs = 0
'        End If
'        ' Calc the average retail Price
'        If sngUPSPW > 0 Then
'            sngRetail = Round(sngRPSPW / sngUPSPW, 2)
'        Else
'            sngRetail = 0
'        End If
'        If lValidDivs > 0 Then lUPSPW = Round((sngUPSPW / lValidDivs), 0)
'        ' Calc the average UPSPW for all the Aldi Stores
'        If lValidDivs > 0 Then lUPSPW = Round((sngUPSPW / lValidDivs), 0)
'        ' Get the Min Qty required to make it a Tier 2
'        lMinQty = Val(g_getConfig("MinUPSPW", g_getDB("ASYST"), , 10))
'        If lUPSPW >= lMinQty Then
'            AST_GetEstTierData = 2
'        Else
'            AST_GetEstTierData = 3
'        End If
'    End If



Exit_Routine:
    On Error Resume Next
    Exit Function
    GoTo Exit_Routine
    
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_GetEstTierData", 3)
    CBA_Error = Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & vbCrLf & g_GetDB("ASYST")
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function AST_UpdGBDMDate(ByVal stApplyDt As String, lPromoID As Long, Optional bTest As Boolean = False, _
                Optional sInProds As String = "", Optional sDspmsg As String = "All", Optional DtTyp As String = "GBDM", Optional bIncDates As Boolean = True) As String
    ' Copy the GBDM Approval Date into all the lines of the Product Table ('GBDM Approval Date' and 'Product Approval Date')
    Dim sSQL As String, lRecs As Long, bUpd As Boolean, lRtn As Long, bError As Boolean, lVar As Long
    Dim sOSD As String, sComp As String, sNoDate As String, sMsg As String, sTmpMsg As String, sProd As String, sProds As String, sSep As String
    Dim RS As ADODB.Recordset, CN As ADODB.Connection
    Const THISTABLE1 = "L2_Products"
    On Error GoTo Err_Routine
    AST_UpdGBDMDate = "True"
    CBA_ErrTag = "SQL"
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    With RS
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
    End With
    sOSD = "": sComp = "": sNoDate = "": bError = False: sProds = "": sProd = "": sSep = ""
    stApplyDt = g_FixDate(stApplyDt)
    sMsg = "Some products have a 'x' Status " & vbCrLf & "   Do you want to change these to " & IIf(g_IsDate(stApplyDt), " the GBDM Date ", " a blank GBDM Date ") & " too?"
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";"
        sSQL = "SELECT * FROM " & THISTABLE1 & " WHERE PD_PG_ID = " & lPromoID & IIf(sInProds > "", " AND PD_ID IN (" & sInProds & ")", "")
    GoSub GSProcess
    RS.Close
    If bTest = False And sProds = "" Then
        bUpd = True
        GoSub GSProcess
    Else
        If sProds > "" Then
            sTmpMsg = "The following Products ( " & sProds & " ) are not complete. These will need to be fixed before the " & DtTyp & " Date can be applied."
            If sDspmsg = "All" Or sDspmsg = "Fatal" Then
                MsgBox sTmpMsg, vbOKOnly
                AST_UpdGBDMDate = "False"
            Else
                AST_UpdGBDMDate = sTmpMsg
            End If
        End If
    End If
''    MsgBox lRecs & " records processed with Uplift Data"
Exit_Routine:
    Set RS = Nothing
    Set CN = Nothing
    Exit Function
GSProcess:
    ' Cycle through all the lines to upd and test the lines for GBDM Date
    RS.Open sSQL, CN
    Do While Not RS.EOF
        CBA_ErrTag = ""
        lRecs = lRecs + 1
        sProd = RS!PD_Product_Code
        If bUpd Then
            If RS!PD_Status < 4 Then
                ' If no Approival date - will transfer it from the GBDM Date
                If g_IsDate(RS!PD_GBDM_Approval_Date, True) <> g_IsDate(stApplyDt, True) Then
                    If bIncDates Then
                        RS!PD_GBDM_Approval_Date = IIf(g_IsDate(stApplyDt, True) = True, stApplyDt, Null)
                        RS!PD_GBDM_Approved_Date = Now
                    End If
                End If
                ' Fill in any Tiers
RedoTiers:
                If NZ(RS!PD_Actual_Marketing_Support, "") = "" Then
                    lVar = InputBox(sProd & " has no 'Actual Marketing Support' entered. Enter the number of the Tier Req.", "Actual Marketing Support", 0)
                    If lVar < 1 Or lVar > 3 Then
                        MsgBox "Invalid entry - please retry"
                        GoTo RedoTiers
                    Else
                        RS!PD_Actual_Marketing_Support = lVar
                    End If
                End If
RedoFreqs:
                ' Fill in any FreqsRedoFreqs:
                If NZ(RS!PD_Freq_Items, "") = "" Then
                    lVar = InputBox(sProd & " has no 'Frequency / Item Basket' entered. Enter 1 for 'Frequency' and anything else for 'Item Basket'.", "Frequency / Item Basket", 0)
                    If lVar < 1 Then
                        MsgBox "Invalid entry - please retry"
                        GoTo RedoFreqs
                    Else
                        RS!PD_Freq_Items = IIf(lVar = 1, 1, 2)
                    End If
                End If
                ' Set the status
                RS!PD_Status = 3
''                    End If
            ElseIf RS!PD_Status = 4 And sOSD = "Y" Or RS!PD_Status = 5 And sComp = "Y" Then
                If bIncDates Then RS!PD_GBDM_Approval_Date = IIf(g_IsDate(stApplyDt, True) = True, stApplyDt, Null)
            End If
            If g_IsDate(RS!PD_Product_Approval_Date) = False And sNoDate = "Y" Then
                If bIncDates Then RS!PD_Product_Approval_Date = IIf(g_IsDate(stApplyDt, True) = True, stApplyDt, Null)
            End If
            RS!PD_LastUpd = Now()
            RS!PD_UpdUser = CBA_User
            RS.Update
        Else
            ' If Completed, Suspended or Deleted then don't change
            If RS!PD_Status > 3 Or g_IsDate(stApplyDt, True) = False Then GoTo NextDate
            ' Get any errors that will stop the GBDM date from being applied
            bError = False
''            If RS!PD_Actual_Marketing_Support = "" Or RS!PD_Freq_Items = "" Then
''                bError = True
''            Else
            If NZ(RS!PD_Unit_Cost, 0) = 0 Or NZ(RS!PD_Retail_Price, 0) = 0 Or NZ(RS!PD_Sales_Multiplier, 0) = 0 Or NZ(RS!PD_Calculated_Sales, 0) = 0 Then
                bError = True
            ElseIf g_DLookup("PV_ID", "L3_ProductRegions", "PV_PD_ID=" & RS!PD_ID, "PV_PD_ID", g_GetDB("ASYST"), 0) = 0 Then
                bError = True           ' No regions have been applied
            End If
            If bError Then
                sProds = sProds & IIf(sSep > "", sSep, "") & sProd
                sSep = ", "
            End If
            ' If there are errors then hop out
            If sProds > "" Then GoTo NextDate
            ' Other tests that decide if the GBDM can be applied...
            If RS!PD_Status = 4 And sOSD = "" Then                              ' If On Sale....
                If sDspmsg = "All" Or sDspmsg = "Test" Then
                    lRtn = MsgBox(Replace(sMsg, "a 'x'", "an 'On Sale'"), vbYesNo, "Update Warning")
                Else
                    lRtn = vbYes
                End If
                If lRtn = vbYes Then
                    sOSD = "Y"
                Else
                    sOSD = "N"
                End If
            ElseIf RS!PD_Status = 5 And sComp = "" Then                         ' If Completed
                If sDspmsg = "All" Or sDspmsg = "Test" Then
                    lRtn = MsgBox(Replace(sMsg, "'x'", "'Completed'"), vbYesNo, "Update Warning")
                Else
                    lRtn = vbYes
                End If
                If lRtn = vbYes Then
                    sComp = "Y"
                Else
                    sComp = "N"
                End If
            End If
            If g_IsDate(stApplyDt, True) = True And DtTyp = "GBDM" Then                                         ' If a date...
                If g_IsDate(RS!PD_Product_Approval_Date) = False And sNoDate = "" Then      ' If No Product Approval....
                    If sDspmsg = "All" Or sDspmsg = "Test" Then
                        lRtn = MsgBox("Some products have no 'Product Approval Date' - This date will be added?", vbOKOnly, "Update Warning")
                    Else
                        lRtn = vbYes
                    End If
                    If lRtn = lRtn Then
                        sNoDate = "Y"
                    Else
                        sNoDate = "N"
                    End If
                End If
''            Else                                                                       ' If Blanked out...
''                If g_IsDate(RS!PD_Product_Approval_Date) = True And sNoDate = "" Then      ' If No Product Approval....
''                    lRtn = MsgBox("Some products have 'GBD Product Approval' - Do you want to blank this date?", vbYesNo, "Update Warning")
''                    If lRtn = vbYes Then
''                        sNoDate = "Y"
''                    Else
''                        sNoDate = "N"
''                    End If
''                End If
            End If
        End If
NextDate:
        RS.MoveNext
    Loop
    
    Return

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_UpdGBDMDate", 3)
    CBA_Error = Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & vbCrLf & g_GetDB("ASYST")
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function AST_CrtProductRegionSS(sTbl_Flds As String, sFld_Prefix As String, sTbl_Nam As String, sExcepts As String, lRegion As Long, sAddits As String) As String
    ' Will create the Excel File that is sent to the regions
    Dim wbk_Tmp As Workbook, wks_Tmp As Worksheet, lIdx As Long, sRevDate As String, aSp() As String
    Static dtSentDate As Date
    Const a_HdgRow As Long = 7
    Const a_COLS As String = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AB,AC,AD,AE,AF,AG,AH"
    Const a_Regs As String = "Min,Der,Stp,Pre,Dan,Bre,Rgy,xxx,Jkt"
    
    
    Dim lRow As Long, sPromo As String, sSavePath As String, rngR As Range, sAllFields As String, sSep As String
    Dim aCols() As String, aRegs() As String, sReg As String, aHdgFlds() As String, aTblFlds() As String, aTabFlds() As String
    Dim sSQL As String, lProds As Long, sFmt As String, sBaseFmt As String, sHdg As String, lWidth As Long, dtDate As Date
    Dim lPG_ID As Long, lcol As Long, bDeleted As Boolean, bProdChg As Boolean, bRegChg As Boolean, lFldIdx As Long
    Static bDateWritten As Boolean
    
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Create ASYST Super Saver Report for Region " & lRegion
    Application.ScreenUpdating = False
    
    If lRegion = 501 Then
        bDateWritten = False
        dtSentDate = CDate("00:00")
    End If
    If g_IsDate(dtSentDate) = False Then dtSentDate = Now()
    AST_CrtProductRegionSS = dtSentDate
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    'Static bNot1st As Boolean
    On Error GoTo Err_Routine
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";"
        
    ' Get more managable names for headings and control
    sAllFields = sTbl_Flds
    ' Sort out the fields that you don't want to change
    aSp = Split(sExcepts, ",")
    For lcol = 0 To UBound(aSp, 1)
        sAllFields = Replace(sAllFields, aSp(lcol), Replace(aSp(lcol), "_", ""))
    Next
    ' Take out the field prefixes so you have the base field name
    aSp = Split(sFld_Prefix, ",")
    For lcol = 0 To UBound(aSp, 1)
        sAllFields = Replace(sAllFields, aSp(lcol), "")
    Next
    ' Split the name of the Form fields into a private array with spaces etc
    aHdgFlds = Split(Replace(sAllFields, "_", " "), ",")
    ' Split the name of the Form fields into a private array without spaces etc
    aTabFlds = Split(Replace(sAllFields, "_", ""), ",")
    ' Call a routine that will give back the column number of the field
    Call AST_TF(0, Replace(sAllFields, "_", ""))
    ' Split the name of the Table fields into a private array
    aTblFlds = Split(sTbl_Flds, ",")
    ' Split the Columns
    aCols = Split(a_COLS, ",")
    ' Split the Regions
    aRegs = Split(a_Regs, ",")
    ' Get the region short name
    sReg = aRegs(lRegion - 501)
''    ' Get the promo name and desc etc
''    sSQL = "PG_ID=" & sRevDate
''    sPromo = g_DLookup("PG_Promo_Desc & ""-"" & PG_Theme", "L1_Promotions", sSQL, "", g_getDB("ASYST"), "Unknown")
    ' Open a new workbook...
    Set wbk_Tmp = Application.Workbooks.Add

    ' Add a Summary Worksheet
    Set wks_Tmp = ActiveSheet
    'wks_Tmp.Visible = xlSheetVeryHidden
    wks_Tmp.Name = "Summary"
    wks_Tmp.Cells.Locked = True
    ' Set up the initials for the Summary sheet
    Range(wks_Tmp.Cells(1, 1), wks_Tmp.Cells(5, AST_TF("PVID"))).Interior.ColorIndex = 49
    wks_Tmp.Cells(1, 1).Select
    wks_Tmp.Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
    Cells.Font.Name = "ALDI SUED Office"
    wks_Tmp.Cells(3, 4).Value = "Region: " & lRegion & " - Promotions as of : " & g_FixDate(Now(), CBA_D3DMYHN)
    wks_Tmp.Cells(3, 4).Font.Size = 22
    wks_Tmp.Cells(3, 4).Font.ColorIndex = 2
    ' Reset the accumulator...
    Call PL(-1, 0, True)
    ' For each field in the headings array, set up the headings
    Set rngR = Range(wks_Tmp.Cells(a_HdgRow, 1), wks_Tmp.Cells(a_HdgRow, AST_TF("PVID")))
    rngR.Interior.ColorIndex = 6
    rngR.BorderAround xlContinuous, xlThick, xlColorIndexNone
    rngR.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    rngR.Borders(xlInsideVertical).LineStyle = xlContinuous
    rngR.HorizontalAlignment = xlHAlignCenter
    rngR.VerticalAlignment = xlVAlignCenter
    wks_Tmp.Activate
    Rows(a_HdgRow).Select
    
    Selection.Rows.WrapText = True
    Selection.Rows.AutoFit
    
    For lIdx = 0 To AST_TF("PVID")
        Call PL(1)      ' Increment by one
        sHdg = AST_FillTagArrays(aTabFlds(PL()), 2, 0, "Hdg")
        lWidth = Val(AST_FillTagArrays(aTabFlds(PL()), 2, 0, "Width"))
        wks_Tmp.Cells(a_HdgRow, PL(0, 1)).Value = sHdg
        If sHdg <> "Press Dates" Or sReg <> "Jkt" Then
            Range(wks_Tmp.Cells(a_HdgRow, PL(0, 1)), wks_Tmp.Cells(a_HdgRow, PL(0, 1))).EntireColumn.ColumnWidth = lWidth
        Else
            Range(wks_Tmp.Cells(a_HdgRow, PL(0, 1)), wks_Tmp.Cells(a_HdgRow, PL(0, 1))).EntireColumn.ColumnWidth = 0
        End If
    Next lIdx
    ' Cycle through all the lines - Note the Non-GBDM Approved, New or 'Not Changed but have been sent' and wrongly statused promotions have already been ignored
    lRow = a_HdgRow: lProds = 0: lPG_ID = -1: sSep = "": sPromo = ""
    sSQL = "SELECT " & sTbl_Flds & sAddits & " FROM [" & sTbl_Nam & "] WHERE Region = " & lRegion & " ORDER BY PD_PG_ID, PD_ID ;"     ''' PVStatus < 3 AND PDStatus = 3 AND PV_ReturnDate IS NULL AND
    RS.Open sSQL, CN
    Do While Not RS.EOF
        dtDate = IIf(g_IsDate(RS!PD_GBDM_Approved_Date), RS!PD_GBDM_Approved_Date, Now)
        bDeleted = (RS!PVStatus = 7) Or (RS!PDStatus = 7)
        bProdChg = (NZ(RS!PDLastUpd, CDate("01/01/2000")) > dtDate)
        bRegChg = (NZ(RS!PVLastUpd, CDate("01/01/2000")) > dtDate)
        If lPG_ID <> RS!PD_PG_ID Then
            If lPG_ID > 0 Then
                lRow = lRow + 2
            Else
                lRow = lRow + 1
            End If
            wks_Tmp.Cells(lRow, 1).Font.Size = 16
            wks_Tmp.Cells(lRow, 1).Value = RS!PG_Promo_Desc & "-" & RS!PG_Theme
            lPG_ID = RS!PD_PG_ID
            sPromo = sPromo & sSep & lPG_ID
            sSep = ","
        End If
        
'        If RS!PD_Regions = "All" Then
'            bValidPC = True     ' Product included by All region
'        ElseIf InStr(1, RS!PD_Regions, sReg) > 0 Then
'            bValidPC = True     ' Product included by named region
'        Else
'            bValidPC = False    ' This Product is not included for the region
'        End If
'        If bValidPC = False Then GoTo GTNextPC
        lRow = lRow + 1
        lProds = lProds + 1
        Call PL(-1, 0, True)
        ' For each field in the array, add the field to the xls file
        For lIdx = 0 To AST_TF("PVID")
            Call PL(1)          ' Increment by 1
            sFmt = AST_FillTagArrays(aTabFlds(PL()), 2, 0, "Format")
            sBaseFmt = AST_FillTagArrays(aTabFlds(PL()), 2, 0, "BaseFormat")
            If sBaseFmt = "num" And InStr(1, aTabFlds(PL()), "Price") > 0 Then sBaseFmt = "cur"
            lFldIdx = PL()
            If Right(UCase(aTabFlds(PL())), 2) = "ID" Then
                wks_Tmp.Cells(lRow, PL(0, 1)).Value = RS.Fields(PL())
            ElseIf sBaseFmt = "dte" Then
                wks_Tmp.Cells(lRow, PL(0, 1)).NumberFormat = "ddd dd/mm/yy"
                wks_Tmp.Cells(lRow, PL(0, 1)).Value = RS.Fields(PL())
            ElseIf sBaseFmt = "cur" And sFmt = "0.00" Then
                wks_Tmp.Cells(lRow, PL(0, 1)).Value = Format(RS.Fields(PL()), "$" & sFmt)
            ElseIf sFmt > "" Then
                wks_Tmp.Cells(lRow, PL(0, 1)).Value = Format(RS.Fields(PL()), sFmt)
            Else
                wks_Tmp.Cells(lRow, PL(0, 1)).Value = RS.Fields(PL())
            End If
            ' If it is one of the fields to be updated
            If lIdx = AST_TF("SalesMultiplier") Or lIdx = AST_TF("ExpectedSales") Then
                wks_Tmp.Cells(lRow, PL(0, 1)).Locked = False
                wks_Tmp.Cells(lRow, PL(0, 1)).Interior.ColorIndex = 6
            ' If this is a resend because the product details have been changed, colour the product oriented fields
            ElseIf lIdx < AST_TF("FillQty") And bProdChg And Not bDeleted Then
                wks_Tmp.Cells(lRow, PL(0, 1)).Interior.ColorIndex = 45
            ' If this is a resend because the region details have been changed, colour the region oriented fields
            ElseIf lIdx >= AST_TF("FillQty") And bRegChg And Not bDeleted Then
                wks_Tmp.Cells(lRow, PL(0, 1)).Interior.ColorIndex = 45
            End If
        Next lIdx
        Rows(lRow).Select
        Selection.Rows.WrapText = True
        Selection.Rows.AutoFit
        Selection.BorderAround xlContinuous, xlThin, xlColorIndexNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
        Selection.HorizontalAlignment = xlHAlignCenter
        Selection.VerticalAlignment = xlVAlignCenter
        If bDeleted Then
            Selection.Font.Strikethrough = True
        End If
GTNextPC:
        RS.MoveNext
    Loop
    RS.Close
    ' Set the Reverse Date
    sRevDate = Left(g_FixDate(dtSentDate, "yymmddhhnn"), 8)
    ' Save the xls file if rows have been added
    If lRow > a_HdgRow Then
        wks_Tmp.Protect "SuperSavers", True, True, True
''        wbs_PLF.
        sSavePath = g_getConfig("PromoRegionsPath", g_GetDB("ASYST")) & sRevDate
        Call g_MkDir(sSavePath)
        sSavePath = sSavePath & "\" & "Promo" & sRevDate & "-Region" & lRegion & ".xlsx"
        wbk_Tmp.SaveAs sSavePath
    End If
    wbk_Tmp.Close savechanges:=False
    
    ' If it is the first region, update the promotion
    If bDateWritten = False And sPromo > "" Then
        sSQL = "UPDATE L1_Promotions SET PG_Region_Date = " & g_GetSQLDate(dtSentDate, CBA_DMYHN) & " WHERE PG_ID IN (" & sPromo & ") ;"
        bDateWritten = True
        RS.Open sSQL, CN
    End If
    
    ' Tell the Central DB that the file is ready to be sent
    CN.Close
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Gen") & ";"
    sSQL = "INSERT INTO LA1_EMails (Em_PromoID,Em_Promo,Em_Region,Em_FileToSend1,Em_SendTime,Em_TestDB) " & _
            "VALUES (" & sRevDate & ",'" & sPromo & "'," & lRegion & ",'" & sSavePath & "'," & g_GetSQLDate(dtSentDate, CBA_DMYHN) & ",'" & IIf(CBA_TestIP = "Y", "Y", "N") & "' )"
    RS.Open sSQL, CN
    
Exit_Routine:
    On Error Resume Next
    Application.ScreenUpdating = True
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    Set RS = Nothing
    Set CN = Nothing
    Exit Function
    
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_CrtProductRegionSS", 3)
    CBA_Error = Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    Resume Next
    
End Function

Public Function AST_ResetProdRegionSS(dtDate As String, Optional bDelete As Boolean = False) As Boolean
    ' Will delete the Excel files, created by AST_CrtProductRegionSS,  provided that they haven't already been sent to the regions
        
    Dim sPromo As String, sSavePath As String, sSep As String, sSQL As String, sRevDate As String
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    On Error GoTo Err_Routine
    AST_ResetProdRegionSS = False
    Set CN = New ADODB.Connection
    With CN
        .ConnectionTimeout = 100
        .CommandTimeout = 100
    End With
    Set RS = New ADODB.Recordset
    With RS
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
    End With
    ' Set the Reverse Date
    sRevDate = Left(g_FixDate(dtDate, "yymmddhhnn"), 8)
    sSavePath = g_getConfig("PromoRegionsPath", g_GetDB("ASYST")) & sRevDate
    sSavePath = sSavePath & "\" & "Promo" & sRevDate & "*.*"
    ' Flag all the existing LA1_EMails records as deleted
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Gen") & ";"
    sSQL = "SELECT * FROM LA1_EMails WHERE Em_EmailSts = 1 AND Em_SendTime=" & g_GetSQLDate(dtDate, CBA_DMYHN) & " AND Em_TestDB='" & IIf(CBA_TestIP = "Y", "Y", "N") & "' ;"
    RS.Open sSQL, CN
''    ' If there are records, flag the routine as such
''    If Not RS.EOF Then AST_ResetProdRegionSS = True
    ' If just a test then hop out
    If bDelete = False Then GoTo Exit_Routine
    ' Go through and flag as deleted
    Do While Not RS.EOF
        If NZ(RS!Em_Promo, "") > "" Then
            If sPromo = "" Then sPromo = RS!Em_Promo
            ' If there are records, flag the routine as such
            AST_ResetProdRegionSS = True
        End If
        RS!Em_EmailSts = 4
        RS.Update
        RS.MoveNext
    Loop
    RS.Close
    CN.Close
    ' Reset the Promotions so that they haven't been scheduled for sending
    If AST_ResetProdRegionSS = True Then
        CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";"
        Set RS = New ADODB.Recordset
        sSQL = "UPDATE L1_Promotions SET PG_Region_Date = NULL WHERE PG_ID IN (" & sPromo & ") ;"
        RS.Open sSQL, CN
        ' Delete any old Excel files
        Call g_KillFile(sSavePath)
    End If
Exit_Routine:
    Set RS = Nothing
    Set CN = Nothing
    Exit Function
    
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_ResetProdRegionSS", 3)
    CBA_Error = Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    Resume Next
    
End Function


Public Function AST_CrtProductGBDMSS(sTbl_Flds As String, sFld_Prefix As String, sTbl_Nam As String, sExcepts As String) As String
    ' Will create the Excel File that is created for the GBDM
    Dim wbk_Tmp As Workbook, wks_Tmp As Worksheet, lIdx As Long, sRevDate As String, aSp() As String
    Static dtSentDate As Date
    Const a_HdgRow As Long = 7
    Const a_COLS As String = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AB,AC,AD,AE,AF,AG,AH"
    
    
    Dim lRow As Long, sPromo As String, rngR As Range, sAllFields As String, sSep As String
    Dim aCols() As String, aHdgFlds() As String, aTblFlds() As String, aTabFlds() As String
    Dim sSQL As String, lProds As Long, sFmt As String, sBaseFmt As String, sHdg As String, lWidth As Long
    Dim lPG_ID As Long, lcol As Long
    
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Create ASYST Super Saver Report for GBDM"
    Application.ScreenUpdating = False
    
    If g_IsDate(dtSentDate) = False Then dtSentDate = Now()
    AST_CrtProductGBDMSS = dtSentDate
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    'Static bNot1st As Boolean
    On Error GoTo Err_Routine
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";"
        
    ' Get more managable names for headings and control
    sAllFields = sTbl_Flds
    ' Sort out the fields that you don't want to change
    aSp = Split(sExcepts, ",")
    For lcol = 0 To UBound(aSp, 1)
        sAllFields = Replace(sAllFields, aSp(lcol), Replace(aSp(lcol), "_", ""))
    Next
    ' Take out the field prefixes so you have the base field name
    aSp = Split(sFld_Prefix, ",")
    For lcol = 0 To UBound(aSp, 1)
        sAllFields = Replace(sAllFields, aSp(lcol), "")
    Next
    ' Split the name of the Form fields into a private array with spaces etc
    aHdgFlds = Split(Replace(sAllFields, "_", " "), ",")
    ' Split the name of the Form fields into a private array without spaces etc
    aTabFlds = Split(Replace(sAllFields, "_", ""), ",")
    ' Call a routine that will give back the column number of the field
    Call AST_TF(0, Replace(sAllFields, "_", ""))
    ' Split the name of the Table fields into a private array
    aTblFlds = Split(sTbl_Flds, ",")
    ' Split the Columns
    aCols = Split(a_COLS, ",")

    Set wbk_Tmp = Application.Workbooks.Add

    ' Add a Summary Worksheet
    Set wks_Tmp = ActiveSheet
    wks_Tmp.Name = "Summary"
    wks_Tmp.Cells.Locked = True
    ' Set up the initials for the Summary sheet
    Range(wks_Tmp.Cells(1, 1), wks_Tmp.Cells(5, AST_TF("PDID"))).Interior.ColorIndex = 49
    wks_Tmp.Cells(1, 1).Select
    wks_Tmp.Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
    Cells.Font.Name = "ALDI SUED Office"
    wks_Tmp.Cells(3, 4).Value = "Promotions as of : " & g_FixDate(Now(), CBA_D3DMYHN)
    wks_Tmp.Cells(3, 4).Font.Size = 22
    wks_Tmp.Cells(3, 4).Font.ColorIndex = 2
    ' Reset the accumulator...
    Call PL(-1, 0, True)
    ' For each field in the headings array, set up the headings
    Set rngR = Range(wks_Tmp.Cells(a_HdgRow, 1), wks_Tmp.Cells(a_HdgRow, AST_TF("PDID")))
    rngR.Interior.ColorIndex = 6
    rngR.BorderAround xlContinuous, xlThick, xlColorIndexNone
    rngR.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    rngR.Borders(xlInsideVertical).LineStyle = xlContinuous
    rngR.HorizontalAlignment = xlHAlignCenter
    rngR.VerticalAlignment = xlVAlignCenter
    wks_Tmp.Activate
    Rows(a_HdgRow).Select
    
    Selection.Rows.WrapText = True
    Selection.Rows.AutoFit
    
    For lIdx = 0 To AST_TF("PDID")
        Call PL(1)      ' Increment by one
        sHdg = AST_FillTagArrays(aTabFlds(PL()), 2, 0, "Hdg")
        lWidth = Val(AST_FillTagArrays(aTabFlds(PL()), 2, 0, "Width"))
        wks_Tmp.Cells(a_HdgRow, PL(0, 1)).Value = sHdg
        Range(wks_Tmp.Cells(a_HdgRow, PL(0, 1)), wks_Tmp.Cells(a_HdgRow, PL(0, 1))).EntireColumn.ColumnWidth = lWidth
    Next lIdx
    ' Cycle through all the lines
    lRow = a_HdgRow: lProds = 0: lPG_ID = -1: sSep = "": sPromo = ""
    sSQL = "SELECT " & sTbl_Flds & " FROM [" & sTbl_Nam & "] WHERE PD_Status < 3 and PD_Product_Approval_Date IS NOT NULL;"
    RS.Open sSQL, CN
    Do While Not RS.EOF
        If lPG_ID <> RS!PD_PG_ID Then
            If lPG_ID > 0 Then
                lRow = lRow + 2
            Else
                lRow = lRow + 1
            End If
            wks_Tmp.Cells(lRow, 1).Font.Size = 16
            wks_Tmp.Cells(lRow, 1).Value = RS!PG_Promo_Desc & "-" & RS!PG_Theme
            lPG_ID = RS!PD_PG_ID
            sPromo = sPromo & sSep & lPG_ID
            sSep = ","
        End If
        
        lRow = lRow + 1
        lProds = lProds + 1
        Call PL(-1, 0, True)
        ' For each field in the array, add the field to the xls file
        For lIdx = 0 To AST_TF("PDID")
            Call PL(1)          ' Increment by 1
            sFmt = AST_FillTagArrays(aTabFlds(PL()), 2, 0, "Format")
            sBaseFmt = AST_FillTagArrays(aTabFlds(PL()), 2, 0, "BaseFormat")
            If sBaseFmt = "num" And InStr(1, aTabFlds(PL()), "Price") > 0 Then sBaseFmt = "cur"
            
            If Right(UCase(aTabFlds(PL())), 2) = "ID" Then
                wks_Tmp.Cells(lRow, PL(0, 1)).Value = RS.Fields(PL())
            ElseIf sBaseFmt = "dte" Then
                wks_Tmp.Cells(lRow, PL(0, 1)).NumberFormat = "ddd dd/mm/yy"
                wks_Tmp.Cells(lRow, PL(0, 1)).Value = RS.Fields(PL())
            ElseIf sBaseFmt = "cur" And sFmt = "0.00" Then
                wks_Tmp.Cells(lRow, PL(0, 1)).Value = Format(RS.Fields(PL()), "$" & sFmt)
            ElseIf sFmt > "" Then
                wks_Tmp.Cells(lRow, PL(0, 1)).Value = Format(RS.Fields(PL()), sFmt)
            Else
                wks_Tmp.Cells(lRow, PL(0, 1)).Value = RS.Fields(PL())
            End If
''            ' If it is one of the fields to be updated
''            If lIdx = AST_TF("SalesMultiplier") Or lIdx = AST_TF("ExpectedSales") Then
''                wks_Tmp.Cells(lRow, PL(0, 1)).Locked = False
''                wks_Tmp.Cells(lRow, PL(0, 1)).Interior.ColorIndex = 6
''            End If
        Next lIdx
        Rows(lRow).Select
        Selection.Rows.WrapText = True
        Selection.Rows.AutoFit
        Selection.BorderAround xlContinuous, xlThin, xlColorIndexNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
        Selection.HorizontalAlignment = xlHAlignCenter
        Selection.VerticalAlignment = xlVAlignCenter
GTNextPC:
        RS.MoveNext
    Loop
    RS.Close
    ' Set the Reverse Date
    sRevDate = Left(g_FixDate(dtSentDate, "yymmddhhnn"), 9)
    ' Save the xls file if rows have been added
    If lRow > a_HdgRow Then
        ''wks_Tmp.Protect Password:="SuperSavers"
        wks_Tmp.Protect "SuperSavers", True, True, True
''        wbs_PLF.
        'sSavePath = g_getConfig("PromoGBDMPath", g_getDB("ASYST")) & "Promo" & sRevDate & ".xlsx"
''        sSavePath = g_SaveFileTo("GBDM Report " & g_RevDate(Date))
''
''        wbk_Tmp.SaveAs sSavePath
    End If
''    wbk_Tmp.Close SaveChanges:=False
    
''    ' If it is the last region, update the promotion
''    If lRegion = 509 Then
''        sSQL = "UPDATE L1_Promotions SET PG_Region_Date = " & g_GetSQLDate(dtSentDate, CBA_DMYHN) & " WHERE PG_ID IN (" & sPromo & ") ;"
''        RS.Open sSQL, CN
''    End If
    
    ' Tell the Central DB that the file is ready to be sent
''    CN.Close
''    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_getDB("Gen") & ";"
''    sSQL = "INSERT INTO LA1_EMails (Em_PromoID,Em_Promo,Em_Region,Em_FileToSend1,Em_SendTime) " & _
''            "VALUES (" & sRevDate & ",'" & sPromo & "'," & lRegion & ",'" & sSavePath & "'," & g_GetSQLDate(dtSentDate, CBA_DMYHN) & ")"
''    RS.Open sSQL, CN
    
 
    
Exit_Routine:
    On Error Resume Next
    Application.ScreenUpdating = False
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    wbk_Tmp.Application.Visible = True
    Set RS = Nothing
    Set CN = Nothing
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_CrtProductGBDMSS", 3)
    CBA_Error = Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function

Public Function AST_UpdSalesData(Optional sReset As String = "N") As String
    ' As for various reasons, some of the original sales data is 0, refill it in a new unit of time
    ' THIS IS ONLY A FIX FOR TESTING - IT SHOULDN'T BE RUN ON LIVE DATA UNLESS YOU ARE SURE YOU KNOW WHAT YOU ARE DOING
    
    Dim sSQL As String, lRecs As Long, bUpd As Boolean, lRtn As Long, bError As Boolean, lVar As Long, bPassOK As Boolean
    Dim sOSD As String, sComp As String, sNoDate As String, sMsg As String, sTmpMsg As String, sProd As String, sProds As String, sSep As String
    Dim sOnSaleDate12 As String, sEndDate12 As String, sngMult As Single, lIdx As Long, lUPSPW As Long, lExpMSupp As Long, sngRetail As Single
    Dim RS As ADODB.Recordset, CN As ADODB.Connection
    Const THISTABLE1 = "L2_Products"
    On Error GoTo Err_Routine
    AST_UpdSalesData = "True"
    CBA_ErrTag = "SQL"
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    With RS
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
    End With
    bPassOK = False: sProds = ""
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";"
    sSQL = "SELECT * FROM L2_Products;"
    ' Cycle through all the lines to upd and test the lines for values
    RS.Open sSQL, CN
    Do While Not RS.EOF
        CBA_ErrTag = "Upd"
        lRecs = lRecs + 1
        If RS!PD_ID >= 57 Or RS!PD_ID <= 60 Then
            CBA_ErrTag = CBA_ErrTag
        End If
        sProd = RS!PD_Product_Code
        ' If they are zero then Fill the Current Retail Value and the Costs for the item
        If NZ(RS!PD_Orig_Unit_Cost, 0) = 0 Or sReset = "Y" Then
            CBA_DBtoQuery = 599
            bPassOK = CBA_SQL_Queries.CBA_GenPullSQL("CBIS_Retail and Cost", Date, , sProd)
            If bPassOK = False Then
                bPassOK = CBA_SQL_Queries.CBA_GenPullSQL("CBIS_Retail and Cost2", Date, , sProd)
            End If
            If bPassOK Then
                RS!PD_Orig_Unit_Cost = NZ(CBA_CBISarr(2, 0), 0)
                If NZ(RS!PD_Unit_Cost, 0) = 0 Or sReset = "Y" Then RS!PD_Unit_Cost = NZ(CBA_CBISarr(2, 0), 0)
                If NZ(RS!PD_Curr_Retail_Price, 0) = 0 Or sReset = "Y" Then RS!PD_Curr_Retail_Price = NZ(CBA_CBISarr(1, 0), 0)
                If NZ(RS!PD_Retail_Price, 0) = 0 Or sReset = "Y" Then RS!PD_Retail_Price = NZ(CBA_CBISarr(1, 0), 0)
            End If
            CBA_ErrTag = "Upd"
        End If
        ' If they are zero then Fill the sales figures item
        If NZ(RS!PD_Prior_Sales, 0) = 0 Or NZ(RS!PD_Prior_UPSPW, 0) = 0 Or sReset = "Y" Then
            sOnSaleDate12 = DateAdd("m", -12, g_FixDate(RS!PD_On_Sale_Date))
            RS!PD_End_Date = (RS!PD_On_Sale_Date + (RS!PD_Weeks_Of_Sale * 7) - 1)
            sEndDate12 = DateAdd("m", -12, g_FixDate(RS!PD_End_Date))
            RS!PD_Prior_Sales = AST_getCompareSalesCBis(sOnSaleDate12, sEndDate12, sProd, "Prior")
            If NZ(RS!PD_Sales_Multiplier, 0) = 0 Then RS!PD_Sales_Multiplier = 1
            sngMult = RS!PD_Sales_Multiplier
            ' Handle the expected sales, if 0
            If NZ(RS!PD_Expected_Sales, 0) = 0 Or sReset = "Y" Then RS!PD_Expected_Sales = AST_getCompareSalesCBis(sOnSaleDate12, sEndDate12, sProd, "Now", , sngMult)
            ' Handle the calculated sales, if 0
            If NZ(RS!PD_Calculated_Sales, 0) = 0 Or sReset = "Y" Then RS!PD_Calculated_Sales = RS!PD_Expected_Sales * sngMult
            lUPSPW = 0: sngRetail = 0: lExpMSupp = 0
            ' Various other fields if 0
            lExpMSupp = AST_GetEstTierData(sProd, Left(NZ(RS!PD_CGSCG, "  "), 2))
            If NZ(RS!PD_Expected_Marketing_Support, "") = "" Or sReset = "Y" Then RS!PD_Expected_Marketing_Support = lExpMSupp
            If NZ(RS!PD_UPSPW, 0) = 0 Or sReset = "Y" Then RS!PD_UPSPW = lUPSPW
            If NZ(RS!PD_Prior_UPSPW, 0) = 0 Or sReset = "Y" Then RS!PD_Prior_UPSPW = lUPSPW
            If NZ(RS!PD_Retail_Price, 0) = 0 Or sReset = "Y" Then RS!PD_Retail_Price = sngRetail
            If NZ(RS!PD_Actual_Marketing_Support, "") = "" Or sReset = "Y" Then RS!PD_Actual_Marketing_Support = lExpMSupp
            ' Work out the discount if 0
            If RS!PD_Retail_Discount = 0 Then
                If RS!PD_Curr_Retail_Price > 0 Then
                    RS!PD_Retail_Discount = 100 * (Round(((RS!PD_Curr_Retail_Price - RS!PD_Retail_Price) / RS!PD_Curr_Retail_Price), 3))
                Else
                    RS!PD_Retail_Discount = 0
                End If
            End If
            CBA_ErrTag = "Upd"
        End If
        ' Finally fill in the user and updated dates???
        RS!PD_LastUpd = Now()
        RS!PD_UpdUser = CBA_User
        RS.Update
        RS.MoveNext
    Loop
    
Exit_Routine:
    Set RS = Nothing
    Set CN = Nothing
    Exit Function


Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-AST_UpdSalesData", 3)
    CBA_Error = Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & vbCrLf & g_GetDB("ASYST")
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Public Function AST_GBDMDate(sGBDMDate As String) As String
    ' Will return and Update the GBDM Date - i.e. the last time that the creation of the GBDM Report was done
    Static sExistingDate As String

    AST_GBDMDate = ""
    If sExistingDate = "" Then
        sExistingDate = g_getConfig("GBDMDateReport", g_GetDB("ASYST"), False, "x")
    End If
    If g_IsDate(sGBDMDate, True) = False Then
        AST_GBDMDate = IIf(g_IsDate(sExistingDate, True) = True, g_FixDate(sExistingDate, CBA_DMYHN), "")
        Exit Function
    Else
        sExistingDate = g_UpdConfig("GBDMDateReport", g_GetDB("ASYST"), g_FixDate(sGBDMDate, CBA_DMYHN))
        AST_GBDMDate = sExistingDate
    End If
    
End Function

Public Function AST_RegionDate(sRegionDate As String) As String
    ' Will return and Update the Region Date - i.e. the last time that the creation of the Region Report was done
    Static sExistingDate As String

    AST_RegionDate = ""
    If sExistingDate = "" Then
        sExistingDate = g_getConfig("RegionDateReport", g_GetDB("ASYST"), False, "x")
    End If
    If g_IsDate(sRegionDate, True) = False Then
        AST_RegionDate = IIf(g_IsDate(sExistingDate, True) = True, g_FixDate(sExistingDate, CBA_DMYHN), "")
        Exit Function
    Else
        sExistingDate = g_UpdConfig("RegionDateReport", g_GetDB("ASYST"), g_FixDate(sRegionDate, CBA_DMYHN))
        AST_RegionDate = sExistingDate
    End If
    
End Function

Public Sub AST_WeeksOfSale(cActCtl As ComboBox)
    ' Will fill the Weeks Of Sale DDBox
    Dim lRow As Long
    Static sWeeksOfSale As String, aryWks As Variant
    If sWeeksOfSale = "" Then
        sWeeksOfSale = g_getConfig("WeeksOfSale", g_GetDB("ASYST"), False, "4,6,8")
        aryWks = Split(sWeeksOfSale, ",")
    End If
   
    ' For each row in the array...
    For lRow = 0 To UBound(aryWks, 1)
        cActCtl.AddItem aryWks(lRow)
    Next
    
End Sub

Public Sub AST_SetCmdVis(cmdK As CommandButton, ByVal bVis As Boolean, ByVal lFrmID As Long, ByVal lAuthority As Long)
    ' Set the visibility of the CmdKey according to what is in the tab
    Dim bNew As Boolean, sTag As String
    sTag = AST_FillTagArrays(cmdK.Tag, lFrmID, lAuthority, "Vis")
    If bVis = True Then
        If sTag = "vis" Then bNew = bVis
    Else
        bNew = False
    End If
    
    If Not cmdK.Visible = bNew Then cmdK.Visible = bNew
    ''Debug.Print cmdK.Tag & " (" & IIf(cmdK.Visible = True, "Vis", "Invis") & ") - ";
End Sub


Public Function PL(Optional ByVal lInc As Long = 0, Optional ByVal lAdditInc As Long = 0, Optional ByVal bReset As Boolean = False, Optional ByVal lIdxVal As Long = 0) As Long
    ' Will form an accumulated increment routine easily
    ' Start with = PL(0,0,true) to start from 0 .... if starting from 1 use PL(1,0,true)
    ' Then PL(1)  (if Incrementing) and PL() (if the same no required) - PL(,,1) will add 1 to the number without affecting the incremented value
    ' lIdxVal allows to to do several (Up to 9) separate index accumulations independantly
    Static lNo(0 To 9) As Long
    If bReset Then lNo(lIdxVal) = 0
    lNo(lIdxVal) = lNo(lIdxVal) + lInc
    PL = lNo(lIdxVal) + lAdditInc
End Function

