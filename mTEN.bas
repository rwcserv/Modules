Attribute VB_Name = "mTEN"
Option Explicit             ' Module for the Ten Project @mTen Changed 190930
Option Private Module       ' Excel users cannot access procedures

' This is the general Overall TEN Module - it mostly starts TEN or has routines that are public in TEN

Public Sub CBA_TEN_Ideas(Control As IRibbonControl)
''    If AST_Dev_Auth("DevIP") = "Y" Then
    Call g_GetDB("Ten", , , True)
    Call mTEN_Runtime.Generate_ID
    Call mTEN_Runtime.Generate_TD
    fTEN_1_Ideas.Show vbModeless
''    End If
End Sub

Public Sub Ten_UpdLevelSts(ByVal lSts As Long)
    ' This routine will update the status of the L1_Levels table
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim sSQL As String, sSep As String, lPFID As Long, lLV_ID As Long '',lLblNo As Long, bReForecast As Boolean, lCG As Long, lSCG As Long, lPCls As Long, sDate As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Ten") & ";"
    ' On an Update of status
    If lPFID = 0 Then
        CBA_ErrTag = ""
        sSQL = "UPDATE L1_Levels SET LV_Sts_ID=" & lSts & ",LV_UpdUser='" & CBA_User & "', LV_UpdDate=" & g_GetSQLDate(Now, CBA_DMYHN) & " " & Chr(10)
        sSQL = sSQL & "WHERE LV_ID= " & lLV_ID & ";"
        CBA_ErrTag = sSQL
        RS.Open sSQL, CN
    End If
    CBA_ErrTag = ""
Exit_Routine:
    On Error Resume Next
    CN.Close
    Set CN = Nothing
    Set RS = Nothing
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-Ten_UpdLevelSts", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Public Function Ten_GetListbox(ByRef lstList As Object, sType As String, sSearch As String, lLV_ID As Long)
    ' This routine will get the Ten data and put it into the Select TEN Product listbox
    Dim sSQL As String, bOutput As Boolean, lIdx As Long, lCols As Long, sColWidths As String, sID As String, sPf As String, sPfVer As String, sPfProd As String, sPfPProd As String, sPfCont As String
    Dim lBoundCol As Long
    ' Get any saved values      mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_Sts_ID")
    sID = mTEN_Runtime.Get_TD_LV(lLV_ID, "ID_ID")
    sPf = mTEN_Runtime.Get_TD_LV(lLV_ID, "LV_Portfolio_No")
    sPfVer = mTEN_Runtime.Get_TD_LV(lLV_ID, "LV_Version_No")
    sPfProd = mTEN_Runtime.Get_TD_LV(lLV_ID, "LV_Product_Code")
    sPfPProd = mTEN_Runtime.Get_TD_LV(lLV_ID, "LV_Prior_Product_Code")
    sPfCont = mTEN_Runtime.Get_TD_LV(lLV_ID, "LV_Contract_No")
    lBoundCol = 1: sColWidths = "0 pt;0 pt;400 pt": lCols = 3
    ' Format the SQL
    If sType = "PF_ID" Then
        If sSearch = "" Then
            MsgBox "Have to enter a search criteria here"
            Exit Function 'GoTo Exit_Routine
        End If
        sSQL = "SELECT pfl.PortfolioID, pfl.Description FROM cbis599p.Portfolio.PortfolioLng pfl" & _
            " WHERE pfl.LanguageID = 0 AND pfl.Description LIKE '%" & sSearch & "%' ORDER BY pfl.Description;"
        lCols = 2: sColWidths = "0 pt;400 pt"
    ElseIf sType = "PV_ID" Then
        If sSearch > "" Then
            sSQL = "SELECT pfvl.pfVersionID, pfvl.Description, pfvl.Supplier FROM cbis599p.Portfolio.PfVersionLng pfvl " & _
                " WHERE pfvl.LanguageID = 0 AND pfvl.PortfolioID='" & sPf & "' AND (pfvl.Description LIKE '%" & sSearch & "%' OR pfvl.Supplier LIKE '%" & sSearch & "%') ORDER BY pfvl.pfVersionID;"
        Else
            sSQL = "SELECT pfvl.pfVersionID, pfvl.Description, pfvl.Supplier FROM cbis599p.Portfolio.PfVersionLng pfvl " & _
                " WHERE pfvl.LanguageID = 0 AND pfvl.PortfolioID='" & sPf & "' ORDER BY pfvl.pfVersionID;"
        End If
    ElseIf sType = "PD_ID" Then
        If sSearch > "" Then
            sSQL = "SELECT pfvm.ProductCode, p.Description, s.Name1 AS Supplier from cbis599p.Portfolio.PfVersionMapping pfvm " & _
                "left join cbis599p.Portfolio.PortfolioLng pfl on pfl.portfolioID = pfvm.portfolioID " & _
                "left join cbis599p.Portfolio.PfVersionLng pfvl on pfvl.portfolioID = pfvm.portfolioID and pfvl.PfVersionID = pfvm.PfVersionID " & _
                "left join cbis599p.dbo.PRODUCT p on p.productcode = pfvm.productcode " & _
                "left join cbis599p.dbo.contract c on c.contractno = pfvm.ContractNo " & _
                "left join cbis599p.dbo.supplier s on s.supplierno = c.SupplierNo " & _
                "WHERE pfl.LanguageID = 0 AND pfvl.PortfolioID='" & sPf & "' AND pfvm.PfVersionID IN ( " & sPfVer & " ) AND " & _
                    "(pfvm.ProductCode LIKE '%" & sSearch & "%' OR p.Description LIKE '%" & sSearch & "%' OR s.Name1 LIKE '%" & sSearch & "%') " & _
                "ORDER BY pfvl.pfVersionID,pfvm.ProductCode;"
        Else
            sSQL = "SELECT pfvm.ProductCode, p.Description, s.Name1 AS Supplier from cbis599p.Portfolio.PfVersionMapping pfvm " & _
                "left join cbis599p.Portfolio.PortfolioLng pfl on pfl.portfolioID = pfvm.portfolioID " & _
                "left join cbis599p.Portfolio.PfVersionLng pfvl on pfvl.portfolioID = pfvm.portfolioID and pfvl.PfVersionID = pfvm.PfVersionID " & _
                "left join cbis599p.dbo.PRODUCT p on p.productcode = pfvm.productcode " & _
                "left join cbis599p.dbo.contract c on c.contractno = pfvm.ContractNo " & _
                "left join cbis599p.dbo.supplier s on s.supplierno = c.SupplierNo " & _
                "WHERE pfl.LanguageID = 0 AND pfvl.PortfolioID='" & sPf & "' AND pfvm.PfVersionID IN ( " & sPfVer & " ) " & _
                "ORDER BY pfvl.pfVersionID,pfvm.ProductCode;"
        End If
    ElseIf sType = "PPD_ID" Then
        If sSearch > "" Then
            sSQL = "SELECT p.ProductCode, p.Description, ' ' AS Supplier from cbis599p.dbo.PRODUCT p " & _
                "WHERE p.Con_ProductCode IS NULL AND ((CHARINDEX('" & sSearch & "', p.ProductCode)>0) OR (CHARINDEX('" & sSearch & "', LOWER(p.Description))>0)) " & _
                "ORDER BY p.ProductCode;"
        Else
            sSQL = "SELECT pfvm.ProductCode, p.Description, s.Name1 AS Supplier from cbis599p.Portfolio.PfVersionMapping pfvm " & _
                "left join cbis599p.Portfolio.PortfolioLng pfl on pfl.portfolioID = pfvm.portfolioID " & _
                "left join cbis599p.Portfolio.PfVersionLng pfvl on pfvl.portfolioID = pfvm.portfolioID and pfvl.PfVersionID = pfvm.PfVersionID " & _
                "left join cbis599p.dbo.PRODUCT p on p.productcode = pfvm.productcode " & _
                "left join cbis599p.dbo.contract c on c.contractno = pfvm.ContractNo " & _
                "left join cbis599p.dbo.supplier s on s.supplierno = c.SupplierNo " & _
                "WHERE pfl.LanguageID = 0 AND pfvl.PortfolioID='" & sPf & "' AND pfvm.PfVersionID IN ( " & sPfVer & " ) " & _
                "ORDER BY pfvl.pfVersionID,pfvm.ProductCode;"
        End If
    ElseIf sType = "PC_ID" Then
            sSQL = "SELECT pfvm.ProductCode, pfvm.ContractNo, s.Name1 AS Supplier from cbis599p.Portfolio.PfVersionMapping pfvm " & _
                "left join cbis599p.Portfolio.PortfolioLng pfl on pfl.portfolioID = pfvm.portfolioID " & _
                "left join cbis599p.Portfolio.PfVersionLng pfvl on pfvl.portfolioID = pfvm.portfolioID and pfvl.PfVersionID = pfvm.PfVersionID " & _
                "left join cbis599p.dbo.contract c on c.contractno = pfvm.ContractNo " & _
                "left join cbis599p.dbo.supplier s on s.supplierno = c.SupplierNo " & _
                "WHERE pfl.LanguageID = 0 AND pfvl.PortfolioID='" & sPf & "' AND pfvm.PfVersionID IN ( " & sPfVer & " ) AND pfvm.ContractNo IS NOT NULL " & _
                "ORDER BY pfvm.ProductCode, pfvm.ContractNo;"
''        End If
        lBoundCol = 2
    End If
    ' Fill the List Box
    CBA_DBtoQuery = 599
    lstList.ColumnWidths = sColWidths: lstList.ColumnCount = lCols: lstList.BoundColumn = lBoundCol
    bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CBIS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", sSQL, 120, , , False)
    lstList.Clear
    If bOutput = True Then
        CBA_ABIarr = CBA_CBISarr
        Call Ten_FillListBoxAry(lstList, lCols, lBoundCol)
    Else
        If sSearch > "" Then MsgBox "No match found for " & sSearch
    End If
    CBA_DBtoQuery = 3
    

End Function

Public Sub Ten_FillListBoxAry(lst As Object, ByVal lCols As Long, ByVal lBoundCol As Long, Optional bClear As Boolean = True)
    ' Will fill the specified List Box with data from the array
    Dim lRow As Long, lcol As Long, sVal As String, sSep As String
    On Error GoTo Err_Routine
    
    If bClear Then lst.Clear
    If lBoundCol > 0 Then lBoundCol = lBoundCol - 1
    ' For each row in the array...
    For lRow = 0 To UBound(CBA_ABIarr, 2)
        lst.AddItem NZ(CBA_ABIarr(lBoundCol, lRow), "")
        lst.List(lRow, 1) = NZ(CBA_ABIarr(lBoundCol + 1, lRow), "")
        sVal = "": sSep = ""
        'Debug.Print CBA_ABIarr(0, lRow) & " " & CBA_ABIarr(1, lRow) & " ";
        For lcol = 0 To lCols - 1
            sVal = sVal & sSep & NZ(CBA_ABIarr(lcol, lRow), "")
            sSep = "; "
        Next
        lst.List(lRow, 2) = sVal
    Next
    On Error Resume Next
''    If CStr(vVal) <> "" Then
''''        lst.SetFocus
''        lst.Value = Null                                    ' Set it to null in case you are setting it to the same value as it doesn't take otherwise
''        vVal = lNewVal
''        lst.Value = lNewVal
''    End If
    DoEvents
Exit_Routine:

    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-Ten_FillListBoxAry", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
