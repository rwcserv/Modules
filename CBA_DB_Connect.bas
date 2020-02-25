Attribute VB_Name = "CBA_DB_Connect"
Option Explicit
Option Private Module          ' Excel users cannot access procedures

'#######################################
'###  DATABASE_CONNECT MODULE v2 02  ### USED FOR NON CLASS RELATED DB CONNECTIONS - OUTPUTS TO ARRAY
'#######################################
'Changing Connections to allow Access DB Connect, as well as Access DB Update
'2.01 - Changed to function and added error trapping
'2.02 - Added seperate XL_Connect Function to allow connection to Excel .xlsx file - Not Tied into the Class Module. Will need further updates for integration
'2.02 - Added Output to Array from CBIS SQL Query
Function CBA_DB_CC_NonC(ByVal CBA_strConnectionType, _
                        ByVal CBA_strDBConnectionName, _
                        ByVal CBA_strDBServerName, _
                        ByVal CBA_strDBProvider, _
                        ByVal CBA_strSQL_TBLNAME, _
                        Optional ByVal CBA_intTimeOuts = 300, _
                        Optional ByVal CBA_strInitialCatalog = "", _
                        Optional ByVal CBA_arrUPDATE, _
                        Optional CBA_tfADAsyncExecute As Boolean) As Boolean

                        'SCRIPT VARIABLES
                        'strConnectionType                      Type of Connection, update/retrieve
                        'strDBConnectionName                    Description for connection as displayed in frm_SQLWAIT                     501PCH or CBIS
                        'strDBServerName                        Server to connect to                                                       501DBL01\SR or 599DBL01
                        'strDBProvider                          SQLNCLI10 (CBIS database) or SQLOLEDB (Purchasing Databases) or " & CBA_MSAccess & ";" for Access Databases
                        'strSQL_TBLNAME                         SQL Query to execute (RETRIEVE) or TABLE TO UPDATE (UPDATE)
                        'intTimeOuts                            (Optional) Timeout in seconds for SQL disconnection of query               Default of 300
                        'strInitialCatalog                      Initial Catalog to use in connection string                                Purchase
                        'arrUPDATE                              Array storing update fields in 0 in update values in fields 1
                        'tfADAsyncExecute                       Adds adAsyncExecute to end of connection string - use only for Access DB connections, set to false for CBIS

                        'CLASS MODULE VARIABLES
                        'Private strREG As String               'Stores Region name
                        'Private intNUM As Integer              'Stores number added in order - used for showing hiding labels in userform
                        'Private rsDBC As ADODB.Recordset       'Stores all records from set database (Recordset Database Connection)





    '-------------------------------------------------------------------------------------------------------
    'CREATING CONNECTIONS TO DATABASES + STORE IN CLASS MODULE cREGDBC
    '-------------------------------------------------------------------------------------------------------
    Dim strSync As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    CBA_DB_CC_NonC = False: CBA_ErrTag = "Conn"

    CBA_strDBServerName = CBA_BasicFunctions.TranslateServerName(CBA_strDBServerName, Date)


    If CBA_strInitialCatalog <> "" Then CBA_strInitialCatalog = ";INITIAL CATALOG=" & CBA_strInitialCatalog
    If CBA_tfADAsyncExecute = True Then strSync = "adAsyncExecute"

    Set CBA_DBCN = New ADODB.Connection
    Set CBA_DBRS = New ADODB.Recordset

    'Set curDBConnection = New cDBCONNS ' Creates instance of class module cREGDBC, used to store Open Connections and Recordsets

    'curDBConnection.strREG = strDBConnectionName
    'curDBConnection.strSTT = "Executing SQL Query, please wait ..."

    'curDBConnection.intNUM = intCountDBCN

    If UCase(CBA_strConnectionType) = "RETRIEVE" Or (UCase(CBA_strConnectionType) = "UPDATE" And (CBA_DBtoQuery = 2 Or CBA_DBtoQuery = 4 Or CBA_DBtoQuery = 6 Or CBA_DBtoQuery = 8 Or CBA_DBtoQuery = 10)) Then
        With CBA_DBCN
            .ConnectionTimeout = CBA_intTimeOuts
            .CommandTimeout = CBA_intTimeOuts
            If CBA_DBtoQuery < 3 Or (CBA_DBtoQuery > 4 And CBA_DBtoQuery < 11) Then
            .Open "Provider=" & CBA_strDBProvider & ";DATA SOURCE=" & CBA_strDBServerName ''' '& ";" & strInitialCatalog & ";INTEGRATED SECURITY=sspi;"
            Else
            .Open "Provider=" & CBA_strDBProvider & ";DATA SOURCE=" & CBA_strDBServerName & ";" & CBA_strInitialCatalog & ";INTEGRATED SECURITY=sspi;"
            End If

        End With
        'Debug.Print CBA_strSQL_TBLNAME
        CBA_ErrTag = "SQL"
        If strSync = "" Then
            CBA_DBRS.Open CBA_strSQL_TBLNAME, CBA_DBCN
        Else
            CBA_DBRS.Open CBA_strSQL_TBLNAME, CBA_DBCN, , , strSync
        End If

    ElseIf UCase(CBA_strConnectionType) = "UPDATE" Then

        With CBA_DBCN
            .ConnectionTimeout = CBA_intTimeOuts
            .CommandTimeout = CBA_intTimeOuts
            .Provider = CBA_strDBProvider
            .ConnectionString = "Provider=" & CBA_strDBProvider & ";Integrated Security=SSPI;Server=" & CBA_strDBServerName ''& ";Database=" & CBA_strServerPartition & ";"
            .Open
        End With
        CBA_ErrTag = "SQL"
'      '  Debug.Print CBA_strSQL_TBLNAME
        CBA_DBRS.Open CBA_strSQL_TBLNAME, CBA_DBCN

        'sT = DBRS.RecordCount

    Else

        MsgBox "DB_Create_Connection called with incorrect switches, set UPDATE or RETRIEVE as first variable", vbCritical, "FAILED"
        Exit Function

    End If

    'sT = DBRS.RecordCount
    'curDBConnection.rsDBC = DBRS

    'allDBConnections.Add curDBConnection, curDBConnection.strREG 'Stores recordset and region details in Class Module cREGDBC
    'intCountDBCN = intCountDBCN + 1
    CBA_ErrTag = "Array Error"
    If UCase(CBA_strConnectionType) = "RETRIEVE" Then
        If CBA_DBRS.State = 0 Then GoTo closedDBRS
        If CBA_DBRS.EOF Then
closedDBRS:
            CBA_DB_CC_NonC = False
            If CBA_DBtoQuery = 1 Then
                ReDim CBA_ABIarr(0, 0)
                CBA_ABIarr(0, 0) = 0
            ElseIf CBA_DBtoQuery = 3 Then
                ReDim CBA_COMarr(0, 0)
                CBA_COMarr = 0
            ElseIf CBA_DBtoQuery = 5 Then
                ReDim CBA_SSarr(0, 0)
                CBA_SSarr = 0
            ElseIf CBA_DBtoQuery = 7 Then
                ReDim CBA_CBFCarr(0, 0)
                CBA_CBFCarr = 0
            ElseIf CBA_DBtoQuery = 9 Then
                ReDim CBA_CDSarr(0, 0)
                CBA_CDSarr(0, 0) = 0
            ElseIf CBA_DBtoQuery > 500 And CBA_DBtoQuery < 510 Then
                ReDim CBA_MMSarr(0, 0)
                CBA_MMSarr = 0
            Else
                ReDim CBA_CBISarr(0, 0)
                CBA_CBISarr(0, 0) = 0
            End If
        Else
            CBA_DB_CC_NonC = True
            If CBA_DBtoQuery = 1 Then
                CBA_ABIarr = CBA_DBRS.GetRows()
            ElseIf CBA_DBtoQuery = 3 Then
                CBA_COMarr = CBA_DBRS.GetRows()
            ElseIf CBA_DBtoQuery = 5 Then
                CBA_SSarr = CBA_DBRS.GetRows()
            ElseIf CBA_DBtoQuery = 7 Then
                CBA_CBFCarr = CBA_DBRS.GetRows()
            ElseIf CBA_DBtoQuery = 9 Then
                CBA_CDSarr = CBA_DBRS.GetRows()
            ElseIf CBA_DBtoQuery > 500 And CBA_DBtoQuery < 510 Then
                CBA_MMSarr = CBA_DBRS.GetRows
            Else
                CBA_CBISarr = CBA_DBRS.GetRows()
            End If

        End If
        If CBA_DBRS.State <> 0 Then CBA_DBRS.Close
    End If

    If Not UCase(CBA_strConnectionType) = "RETRIEVE" Then CBA_DB_CC_NonC = True
Exit_Routine:
    On Error Resume Next
    Set CBA_DBRS = Nothing
    CBA_DBCN.Close
    Set CBA_DBCN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("CBA_DB_Connect-f-CBA_DB_CC_NonC", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

'#######################################
'###  DATABASE_CONNECT MODULE v2 02  ### USED FOR NON CLASS RELATED DB CONNECTIONS - Update SQL
'#######################################
'Changing Connections to allow Access DB Connect, as well as Access DB Update
'2.01 - Changed to function and added error trapping
'2.02 - Added seperate XL_Connect Function to allow connection to Excel .xlsx file - Not Tied into the Class Module. Will need further updates for integration
'2.02 - Added Output to Array from CBIS SQL Query
Function CBA_DB_CBADBUpdate(ByVal CBA_strConnectionType, _
                        ByVal CBA_strDBConnectionName, _
                        ByVal CBA_strDBServerName, _
                        ByVal CBA_strDBProvider, _
                        ByVal CBA_strSQL_TBLNAME)
    Dim CBA_Proc As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    Set CBA_DBCN = New ADODB.Connection
    Set CBA_DBRS = New ADODB.Recordset
        With CBA_DBCN
            .ConnectionTimeout = 100
            .CommandTimeout = 300
            .Open "Provider=" & CBA_strDBProvider & ";DATA SOURCE=" & CBA_strDBServerName '& ";" & CBA_strInitialCatalog & ";INTEGRATED SECURITY=sspi;"
        End With

    CBA_DBRS.Open CBA_strSQL_TBLNAME, CBA_DBCN ',  , , False

Exit_Routine:
    On Error Resume Next
    CBA_DBCN.Close
    Set CBA_DBCN = Nothing
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_DB_CBADBUpdate", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

