Attribute VB_Name = "CBA_PublicVariables"
Option Explicit
Option Private Module          ' Excel users cannot access procedures
'---------Specific Add-in Variables
Public CBA_wks_Run As Workbook
Public CBA_Rib As IRibbonUI
''--------CBA_COM Variables
Public CBA_DBCN As ADODB.Connection, CBA_COM_SKU_CBISCN As ADODB.Connection
Public CBA_DBRS As ADODB.Recordset, CBA_COM_SKU_CBISRS As ADODB.Recordset
Public CBA_DBtoQuery As Long

Public intRefreshSec As Integer
Public CBA_strAldiMsg As String
Public CBA_CBISarr As Variant, CBA_ABIarr As Variant, CBA_COMarr As Variant, CBA_MMSarr As Variant, CBA_CBFCarr As Variant, CBA_SSarr As Variant
Public CBA_COM_colInput As Collection, CBA_COM_potGram As Collection, CBA_COM_potLitres As Collection, CBA_COM_leftovers As Collection
Public CBA_COM_potMetres As Collection, CBA_COM_potPieces As Collection, CBA_COM_potPair As Collection, CBA_COM_potOther As Collection
Public CBA_COM_potSheet As Collection, CBA_COM_colAdddetail As Collection, CBA_COM_colMulti As Collection, CBA_COM_colWhere As Collection, CBA_COM_colNotDecoded As Collection
Public CBA_COM_PackarrOutput() As Variant, CBA_COM_arrOutput() As Variant, CBA_COM_arrSortDetail() As Variant, CBA_COM_CBISarrOutput() As Variant
Public CBA_COM_numOutput As Long

Public CBA_COM_arrWW(), CBA_COM_arrDM(), CBA_COM_arrCBISPack()
'Public CBA_COM_Statelook '', CBA_COM_ADesc
Public CBA_COM_ACGno As Long, CBA_COM_ASCGno As Long, CBA_COM_APcode As Long, CBA_COM_entryrow As Long
Public CBA_colProds As Collection, CBA_COM_colmm As Collection
Public CBA_COM_owsht As Worksheet
Public CBA_COM_owret As Single, CBA_COM_owpr As Single, CBA_COM_Aret As Single
Public CBA_COM_owrow As Long, CBA_COM_owcol As Long
Public CBA_COM_statelookup As String
Public CBA_COM_matchedinfo() As Variant
'----For columns1-7 duplicating matches functionality-----'
Public CBA_PCodeforduplicate As Long
'-----Values for COM_Match Class Module
Public CBA_COM_Match() As New CBA_COM_COMMatch
Public CBA_COM_strNotCreated As String
Public CBA_COM_CBISCN As ADODB.Connection
Public CBA_COM_COMCN As ADODB.Connection
Public CBA_COM_CBIS2CN As ADODB.Connection
Public CBA_COM_MMSCN(501 To 509) As ADODB.Connection
Public CBA_COM_AParr() As Variant, CBA_COM_ADarr() As Variant, CBA_COM_APParr() As Variant
'''-----Values for CCBA_COM_SKU Class
Public CBA_COM_SKU_COMCN As ADODB.Connection
Public CBA_COM_SKU_COMRS As ADODB.Recordset
'-------Values for CCM (COMRADE MATGHING TOOL)
Public CCM_WWSKU() As CBA_COM_COMCompSKU
Public CCM_ColesSKU() As CBA_COM_COMCompSKU
Public CCM_DMSKU() As CBA_COM_COMCompSKU
Public CCM_FCSKU() As CBA_COM_COMCompSKU
Public CCM_UDWWSKU() As CBA_COM_COMCompSKU
Public CCM_UDColesSKU() As CBA_COM_COMCompSKU
Public CCM_UDDMSKU() As CBA_COM_COMCompSKU
Public CCM_UDFCSKU() As CBA_COM_COMCompSKU
Public CCM_FormState(0 To 4) As Boolean

'---------Admin USer Variables
Public CBA_AdminUser As Boolean
''-------------SQLMultiQuery Variables
Public CBA_arrOP1 As Variant, CBA_arrOP2 As Variant, CBA_arrOP3 As Variant, CBA_arrOP4 As Variant, CBA_arrOP5 As Variant
Public CBA_arrOP6 As Variant, CBA_arrOP7 As Variant, CBA_arrOP8 As Variant, CBA_arrOP9 As Variant, CBA_arrOP10 As Variant
Public CBA_arrOP11 As Variant, CBA_arrOP12 As Variant, CBA_arrOP13 As Variant, CBA_arrOP14 As Variant, CBA_arrOP15 As Variant
Public CBA_arrOP16 As Variant, CBA_arrOP17 As Variant, CBA_arrOP18 As Variant, CBA_arrOP19 As Variant, CBA_arrOP20 As Variant
Public CBA_SQLnum As Long
Public CBA_strMSQL() As String
Public CBA_newmatchtype As String

'''----- Public Types for Matching Tool
Public thisDataViewer As CBA_COM_frm_MatchingTool

'''----- Public Types for Forecasting Tool
Public FCbM() As CBA_BTF_MonthData

''' - General date formats
Public Const CBA_DMY As String = "dd/mm/yyyy", CBA_DMYHN As String = "dd/mm/yyyy hh:nn", CBA_DMYH As String = "dd/mm/yyyy hh", CBA_DM2Y = "dd/mm/yy"
Public Const CBA_D2DMY As String = "dd dd/mm/yyyy", CBA_D2DMYHN As String = "dd dd/mm/yyyy hh:nn"    ' Will have to use g_FixDate to format
Public Const CBA_D3DMY As String = "ddd dd/mm/yyyy", CBA_D3DMYHN As String = "ddd dd/mm/yyyy hh:nn"

''' - General path for CBStdAddin stuff
Public Const CBA_BSA = "G:\Central_Buying\General\4_Administration\Z - Buying Systems Analyst\"
''' - Other General stuff
Public Const CBA_TESTUSER As String = "" '"Pearce, Tom"   ''"Lentini, David"   ''"Baines, Stuart" '"Collett, Sarah"  '' "Whiteford, Michael" '' CB AUS/BDM or GBDM etc  Hollier, Moira (CB AUS/PABAD)
Public Const CBA_White As Long = -2147483643, CBA_Grey As Long = 12632256, CBA_Pink As Long = 12632319
Public Const CBA_LtYellow As Long = 12648384, CBA_LtGreen As Long = 12648447     ' These provide the alternate colours for the datasheet type forms
Public Const CBA_MSAccess As String = "Microsoft.ACE.OLEDB.12.0"
Public Const CBA_LongHiVal As Long = 999999999, CBA_CalHeight As Long = 170, CBA_CalWidth As Long = 150
Public Const CBA_EntryYellow As Long = 8454143, CBA_OffYellow As Long = 12640511
Public Const CBA_Red As Long = 16744703, CBA_Green As Long = 8454016, CBA_AldiBlue = 8388608
Public Type typCal                              ' Calendar type
    lCalTop As Long                             ' Top pos of calendar
    lCalLeft As Long                            ' Left pos of calendar
    sDate As String                             ' Return Date or blank
    bCalValReturned As Boolean                  ' If date or blank returned
    bAllowNullOfDate As Boolean                 ' Allow a blank to be returned (make the Null cmd key Visible)
End Type
Public varCal As typCal
Public Type typFld                              ' Field type                    Will define the field that is being accessed in CBA_AST_frm_AssocFld and maybe others
    sHdg As String                              ' Heading of Form
    lFrmTop As Long                             ' Top pos of Form
    lFrmLeft As Long                            ' Left pos of Form
    lFldWidth As Long                           ' Width of Field
    lFldHeight As Long                          ' Height of Field
    sType As String                             ' "Textbox" or "ComboBox"
    sSQL As String                              ' SQL to use to get "ComboBox"
    sDB As String                               ' DB Prfix (e.g "ASYST") to use to get "ComboBox"
    lCols As Long                               ' Number of columns
    lID1 As Long                                 ' ID of record
    lID2 As Long                                 ' ID of record
    sField1 As Variant                          ' Input / Return Field 1
    sField2 As Variant                          '         Return Field 2
    sField3 As Variant                          '         Return Field 3
    sField4 As Variant                          '         Return Field 4
    bFieldReturned As Boolean                   ' If Field or blank returned
    bAllowNullOfField As Boolean                ' Allow a blank to be returned
End Type
Public CBA_Test1 As String, CBA_Test2 As String, CBA_Test3 As String, CBA_Test4 As String
Public varFldVars As typFld, CBA_Erl As Long
Public CBA_TestIP As String, CBA_DevIP As String  ' Testing IP and or Dev IP
Public Const CBA_GEN_ERR = "LIVE DATABASES\Logs\General_Errors.txt"
Public Const CBA_GEN_LOGS = "LIVE DATABASES\Logs\"
Public Const CBA_GEN_DB = "LIVE DATABASES\Central_Ext.accdb"

' ---- Public Variables for the ASyst Super Saver Project
Public CBA_lAuthority As Long, CBA_lPromotion_ID As Long, CBA_lProduct_ID As Long
Public Const CBA_S = "|", CBA_Sn = "||||||"
Public CBA_Error As String, CBA_ErrTag As String
Public CBA_PRa()                                ' ASYST ProductRows Variable for array
Public CBA_bDataChg As Boolean                  ' Has data changed in sub-form
Public CBA_lFrmID As Long                       ' Temp FormID
Public CBA_lAuth As Long                        ' Temp Auth

' ------ Public variables for the Forcasting system
Public Const CBA_FC_MAX_YRS = 2   ' The maximum number of yearly increments allowed - i.e. '1'= 1x'year'; '2'=2x'year); (0=none) etc
Public CBA_bFCast_NoDataReturned As Boolean
Public Const CBA_FCAST_1Min As String = "You have tried to apply a forecast within a minute of your application of the same forecast." & vbCrLf & vbCrLf & _
                                        "You may only apply one forecast per minute.. please wait."
Public Const CBA_FCAST_Apply  As String = "Apply has completed suceesfully." & vbCrLf & vbCrLf & _
                                          "Forecast has been applied for all values for the currently selected product class."
Public Const CBA_FCAST_CHG As String = " Forecasts have been entered or changed but not Applied." & vbCrLf & _
                                       "Press 'Yes' to go back and Apply, or 'No' to lose those changes"

'''----- Public Types for AADD
Public CBA_AADD_CBISCN As ADODB.Connection
Public CBA_AADD_MMSCN(501 To 509) As ADODB.Connection
'''----- Public Types for AADD
'''----- Public Types for CDS_ADM
Public CBA_CDSarr As Variant
'''----- Public Types for TEN
Public Const CBA_MAX_SUPPS = 9             ' If changed in the UDT, will have to be changed here too
Global Const CBA_ROW_HEIGHT = 6.3          ' Worksheet ROW_HEIGHT


Global Const CBA_200_ID As Long = 200      ' If the ID is less than this figure, then it is treated as a NEW ID i.e. it is added; else it is updated
' VERSIONS -- UPDATE AS NEEDED --------------------------------------------------------------
Public Const CBA_AST_Ver = "Version: 19.12.07"
Public Const CBA_FCAST_Ver = "Version: 19.12.07"
Public Const CBA_COM_Ver = "Version: 20.02.12"
Public Const CBA_Cam_Ver = "Version: 20.02.12"
Public Const CBA_Ten_Ver = "Version: 20.02.13"                              ' Set at the date for the (FORMAT=YY.MM.DD)
Public Const CBA_All_Ver As Long = 200213                                   ' Set at the latest date of the above versions (FORMAT=YYMMDD)

' General Change format @RW=Code has questions ?RWAS =Question/s remain with regard to #RW Was changed and date
' Other Questions on a project that may have to wait a while to fix are prefixed with Project acronym - i.e. ?RWFC (Search @RW & ?RW to chg)
' Set up PromoRows for bringing in from database
' Set up ProductRows for bringing in from database


Public Function pfCBA_ROW_HEIGHT() As Long
    pfCBA_ROW_HEIGHT = CBA_ROW_HEIGHT
End Function

Public Function CBA_SetUser(Optional Name_or_NA As String = "") As String
    ' The purpose of this routine is to extend the testing ability of the AddIn
    ' Will change the user name if
    ' 1. CBA_TestIP = "Y"               : NOTE: If testing live with a specific name, you can change this manually to "Y"
    ' 2. If there is a name in CBA_TESTUSER or sInput is a name
    Static sSAvedUser As String
    If CBA_TestIP = "Y" Then
        Select Case UCase(Name_or_NA)
            Case ""
                If sSAvedUser = "" Then
                    sSAvedUser = CBA_TESTUSER
                End If
            Case "NA"
                sSAvedUser = CBA_TESTUSER
            Case Else
                sSAvedUser = Name_or_NA
        End Select
    End If
    CBA_SetUser = sSAvedUser
End Function

