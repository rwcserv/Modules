Attribute VB_Name = "CBA_COM_SetupSKUArray"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Private SKUBUILD_ALLMATCHES As CompCols
Private SKUBUILD_MatchesSetup As Boolean
Private Type CompDics
    WW As Scripting.Dictionary
    Coles As Scripting.Dictionary
    DM As Scripting.Dictionary
    FC As Scripting.Dictionary
End Type
Private Type CompCols
    WW As Collection
    Coles As Collection
    DM As Collection
    FC As Collection
End Type
Private Function SKUBUILD_updateMatchCollections()
    Dim tmm() As MatchMap
    Dim WWcnt As Long, Ccnt As Long, DMcnt As Long, FCcnt As Long, a As Long, T
    Dim tc As CompDics
    Set tc.WW = New Scripting.Dictionary
    Set tc.Coles = New Scripting.Dictionary
    Set tc.DM = New Scripting.Dictionary
    Set tc.FC = New Scripting.Dictionary
    Dim thisCD As CompDics
    Dim thisCC As CompCols
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    Ccnt = 0: WWcnt = 0: DMcnt = 0: FCcnt = 0
    tmm = CCM_Runtime.CCM_getMatches
    For a = LBound(tmm) To UBound(tmm)
        If tmm(a).ColesCB1 <> "" Then If tc.Coles.Exists(tmm(a).ColesCB1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesCB1, Ccnt
        If tmm(a).ColesWNAT1 <> "" Then If tc.Coles.Exists(tmm(a).ColesWNAT1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWNAT1, Ccnt
        If tmm(a).ColesWNAT2 <> "" Then If tc.Coles.Exists(tmm(a).ColesWNAT2) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWNAT2, Ccnt
        If tmm(a).ColesWNAT3 <> "" Then If tc.Coles.Exists(tmm(a).ColesWNAT3) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWNAT3, Ccnt
        If tmm(a).ColesWNAT4 <> "" Then If tc.Coles.Exists(tmm(a).ColesWNAT4) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWNAT4, Ccnt
        If tmm(a).ColesWNSW <> "" Then If tc.Coles.Exists(tmm(a).ColesWNSW) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWNSW, Ccnt
        If tmm(a).ColesWVIC <> "" Then If tc.Coles.Exists(tmm(a).ColesWVIC) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWVIC, Ccnt
        If tmm(a).ColesWQLD <> "" Then If tc.Coles.Exists(tmm(a).ColesWQLD) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWQLD, Ccnt
        If tmm(a).ColesWSA <> "" Then If tc.Coles.Exists(tmm(a).ColesWSA) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWSA, Ccnt
        If tmm(a).ColesWWA <> "" Then If tc.Coles.Exists(tmm(a).ColesWWA) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWWA, Ccnt
        If tmm(a).WWWeb <> "" Then If tc.WW.Exists(tmm(a).WWWeb) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWeb, WWcnt
        If tmm(a).WWWNAT1 <> "" Then If tc.WW.Exists(tmm(a).WWWNAT1) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWNAT1, WWcnt
        If tmm(a).WWWNAT2 <> "" Then If tc.WW.Exists(tmm(a).WWWNAT2) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWNAT2, WWcnt
        If tmm(a).WWWNAT3 <> "" Then If tc.WW.Exists(tmm(a).WWWNAT3) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWNAT3, WWcnt
        If tmm(a).WWWNAT4 <> "" Then If tc.WW.Exists(tmm(a).WWWNAT4) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWNAT4, WWcnt
        If tmm(a).WWWNSW <> "" Then If tc.WW.Exists(tmm(a).WWWNSW) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWNSW, WWcnt
        If tmm(a).WWWVIC <> "" Then If tc.WW.Exists(tmm(a).WWWVIC) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWVIC, WWcnt
        If tmm(a).WWWQLD <> "" Then If tc.WW.Exists(tmm(a).WWWQLD) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWQLD, WWcnt
        If tmm(a).WWWSA <> "" Then If tc.WW.Exists(tmm(a).WWWSA) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWSA, WWcnt
        If tmm(a).WWWWA <> "" Then If tc.WW.Exists(tmm(a).WWWWA) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWWA, WWcnt
        If tmm(a).ColesWeb <> "" Then If tc.Coles.Exists(tmm(a).ColesWeb) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWeb, Ccnt
        If tmm(a).ColesSB1 <> "" Then If tc.Coles.Exists(tmm(a).ColesSB1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesSB1, Ccnt
        If tmm(a).ColesSB2 <> "" Then If tc.Coles.Exists(tmm(a).ColesSB2) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesSB2, Ccnt
        If tmm(a).ColesSB3 <> "" Then If tc.Coles.Exists(tmm(a).ColesSB3) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesSB3, Ccnt
        If tmm(a).ColesPL1 <> "" Then If tc.Coles.Exists(tmm(a).ColesPL1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesPL1, Ccnt
        If tmm(a).ColesPL2 <> "" Then If tc.Coles.Exists(tmm(a).ColesPL2) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesPL2, Ccnt
        If tmm(a).ColesPL3 <> "" Then If tc.Coles.Exists(tmm(a).ColesPL3) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesPL3, Ccnt
        If tmm(a).ColesPL4 <> "" Then If tc.Coles.Exists(tmm(a).ColesPL4) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesPL4, Ccnt
        If tmm(a).ColesVal1 <> "" Then If tc.Coles.Exists(tmm(a).ColesVal1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesVal1, Ccnt
        If tmm(a).ColesVal2 <> "" Then If tc.Coles.Exists(tmm(a).ColesVal2) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesVal2, Ccnt
        If tmm(a).ColesVal3 <> "" Then If tc.Coles.Exists(tmm(a).ColesVal3) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesVal3, Ccnt
        If tmm(a).ColesCB1 <> "" Then If tc.Coles.Exists(tmm(a).ColesCB1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesCB1, Ccnt
        If tmm(a).ColesPB1 <> "" Then If tc.Coles.Exists(tmm(a).ColesPB1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesPB1, Ccnt
        If tmm(a).ColesPB2 <> "" Then If tc.Coles.Exists(tmm(a).ColesPB2) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesPB2, Ccnt
        If tmm(a).ColesML1 <> "" Then If tc.Coles.Exists(tmm(a).ColesML1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesML1, Ccnt
        If tmm(a).ColesML2 <> "" Then If tc.Coles.Exists(tmm(a).ColesML2) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesML2, Ccnt
        If tmm(a).ColesML3 <> "" Then If tc.Coles.Exists(tmm(a).ColesML3) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesML3, Ccnt
        If tmm(a).WWWW1 <> "" Then If tc.WW.Exists(tmm(a).WWWW1) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWW1, WWcnt
        If tmm(a).WWWW2 <> "" Then If tc.WW.Exists(tmm(a).WWWW2) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWW2, WWcnt
        If tmm(a).WWWW3 <> "" Then If tc.WW.Exists(tmm(a).WWWW3) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWW3, WWcnt
        If tmm(a).WWWW4 <> "" Then If tc.WW.Exists(tmm(a).WWWW4) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWW4, WWcnt
        If tmm(a).WWWW5 <> "" Then If tc.WW.Exists(tmm(a).WWWW5) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWW5, WWcnt
        If tmm(a).WWHB1 <> "" Then If tc.WW.Exists(tmm(a).WWHB1) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWHB1, WWcnt
        If tmm(a).WWHB2 <> "" Then If tc.WW.Exists(tmm(a).WWHB2) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWHB2, WWcnt
        If tmm(a).WWHB3 <> "" Then If tc.WW.Exists(tmm(a).WWHB3) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWHB3, WWcnt
        If tmm(a).WWCB1 <> "" Then If tc.WW.Exists(tmm(a).WWCB1) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWCB1, WWcnt
        If tmm(a).WWPB1 <> "" Then If tc.WW.Exists(tmm(a).WWPB1) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWPB1, WWcnt
        If tmm(a).WWPB2 <> "" Then If tc.WW.Exists(tmm(a).WWPB2) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWPB2, WWcnt
        If tmm(a).WWML1 <> "" Then If tc.WW.Exists(tmm(a).WWML1) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWML1, WWcnt
        If tmm(a).WWML2 <> "" Then If tc.WW.Exists(tmm(a).WWML2) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWML2, WWcnt
        If tmm(a).WWML3 <> "" Then If tc.WW.Exists(tmm(a).WWML3) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWML3, WWcnt
        If tmm(a).DM1 <> "" Then If tc.DM.Exists(tmm(a).DM1) = False Then DMcnt = DMcnt + 1: tc.DM.Add tmm(a).DM1, DMcnt
        If tmm(a).DM1Pack <> "" Then If tc.DM.Exists(tmm(a).DM1Pack) = False Then DMcnt = DMcnt + 1:  tc.DM.Add tmm(a).DM1Pack, DMcnt
        If tmm(a).DM2 <> "" Then If tc.DM.Exists(tmm(a).DM2) = False Then DMcnt = DMcnt + 1:  tc.DM.Add tmm(a).DM2, DMcnt
        If tmm(a).DM2Pack <> "" Then If tc.DM.Exists(tmm(a).DM2Pack) = False Then DMcnt = DMcnt + 1:  tc.DM.Add tmm(a).DM2Pack, DMcnt
        If tmm(a).DMQ <> "" Then If tc.DM.Exists(tmm(a).DMQ) = False Then DMcnt = DMcnt + 1:  tc.DM.Add tmm(a).DMQ, DMcnt
        If tmm(a).DMQPack <> "" Then If tc.DM.Exists(tmm(a).DMQPack) = False Then DMcnt = DMcnt + 1:  tc.DM.Add tmm(a).DMQPack, DMcnt
        If tmm(a).FC1 <> "" Then If tc.FC.Exists(tmm(a).FC1) = False Then FCcnt = FCcnt + 1: tc.FC.Add tmm(a).FC1, FCcnt
        If tmm(a).FC1Pack <> "" Then If tc.FC.Exists(tmm(a).FC1Pack) = False Then FCcnt = FCcnt + 1: tc.FC.Add tmm(a).FC1Pack, FCcnt
        If tmm(a).FC2 <> "" Then If tc.FC.Exists(tmm(a).FC2) = False Then FCcnt = FCcnt + 1: tc.FC.Add tmm(a).FC2, FCcnt
        If tmm(a).FC2Pack <> "" Then If tc.FC.Exists(tmm(a).FC2Pack) = False Then FCcnt = FCcnt + 1: tc.FC.Add tmm(a).FC2Pack, FCcnt
        If tmm(a).FCQ <> "" Then If tc.FC.Exists(tmm(a).FCQ) = False Then FCcnt = FCcnt + 1: tc.FC.Add tmm(a).FCQ, FCcnt
        If tmm(a).FCQPack <> "" Then If tc.FC.Exists(tmm(a).FCQPack) = False Then FCcnt = FCcnt + 1: tc.FC.Add tmm(a).FCQPack, FCcnt
    Next
    
    thisCD = tc
    Set thisCC.WW = New Collection
    For Each T In thisCD.WW.Keys
        thisCC.WW.Add T
    Next
    If thisCC.WW.Count > 0 Then CBA_BasicFunctions.CBA_CollectionSort thisCC.WW
    Set thisCC.Coles = New Collection
    For Each T In thisCD.Coles.Keys
        thisCC.Coles.Add T
    Next
    If thisCC.Coles.Count > 0 Then CBA_BasicFunctions.CBA_CollectionSort thisCC.Coles
    Set thisCC.DM = New Collection
    For Each T In thisCD.DM.Keys
        thisCC.DM.Add T
    Next
    If thisCC.DM.Count > 0 Then CBA_BasicFunctions.CBA_CollectionSort thisCC.DM
    Set thisCC.FC = New Collection
    For Each T In thisCD.FC.Keys
        thisCC.FC.Add T
    Next
    If thisCC.FC.Count > 0 Then CBA_BasicFunctions.CBA_CollectionSort thisCC.FC
    
    Set SKUBUILD_ALLMATCHES.WW = thisCC.WW
    Set SKUBUILD_ALLMATCHES.Coles = thisCC.Coles
    Set SKUBUILD_ALLMATCHES.DM = thisCC.DM
    Set SKUBUILD_ALLMATCHES.FC = thisCC.FC
    SKUBUILD_MatchesSetup = True
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-SKUBUILD_updateMatchCollections", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Function CBA_SetupSKUArray(ByVal Competitor As String, Optional cProds As Collection) As CBA_COM_COMCompSKU()
    Dim CBA_COM_SKU_arrprodmap()
    Dim CBA_COM_SKU_LastSeen()
    Dim CBA_COM_SKU_Run As Boolean, CCS_strProds_Exists As Boolean
    Dim SKUarr() As New CBA_COM_COMCompSKU, buildrate, tot, SKU_NUM As Long
    Dim CCS_strProds As String, compnamefull As String, strSQL As String, strProds As String
    Dim PMcur As Long, LScur As Long, LSsp As Long, PMsp As Long, SKU_CNT As Long, a As Long
    Dim DaystoQuery As Long, prod, Match
    Dim CBA_StartTime As Single, CBA_Timer As Single
    Dim WedDates As Scripting.Dictionary
    Dim WedCount As Long
    Dim DT As Date
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    CBA_StartTime = CBA_Timer
    If CBA_COM_getonlymatchedSKU = True And SKUBUILD_MatchesSetup = False Then SKUBUILD_updateMatchCollections
    If IsEmpty(CBA_COM_Match) = False Then Erase CBA_COM_Match
    
    DaystoQuery = (CBA_COM_Runtime.getWeekstoUse * 7) + 1
    CCS_strProds = ""
    If Competitor = "WW" Then
        compnamefull = "Woolworths"
    ElseIf Competitor = "C" Then compnamefull = "Coles"
    ElseIf Competitor = "DM" Then compnamefull = "Dan Murphys"
    ElseIf Competitor = "FC" Then compnamefull = "First Choice"
    End If
    
    
    On Error Resume Next
    For Each prod In cProds
        If CBA_COM_getonlymatchedSKU = False Then
            If prod <> "" Then
                If CCS_strProds = "" Then CCS_strProds = "'" & prod & "'" Else CCS_strProds = CCS_strProds & ", '" & prod & "'"
            End If
        Else
            If IsEmpty(prod) Then
            Else
            If Competitor = "WW" Then
                For Each Match In SKUBUILD_ALLMATCHES.WW
                    If prod <> "" And Match = prod Then
                        If CCS_strProds = "" Then CCS_strProds = "'" & prod & "'" Else CCS_strProds = CCS_strProds & ", '" & prod & "'"
                        Exit For
                    End If
                Next
            ElseIf Competitor = "C" Then
                For Each Match In SKUBUILD_ALLMATCHES.Coles
                    If prod <> "" And Match = prod Then
                        If CCS_strProds = "" Then CCS_strProds = "'" & prod & "'" Else CCS_strProds = CCS_strProds & ", '" & prod & "'"
                        Exit For
                    End If
                Next
            ElseIf Competitor = "DM" Then
                For Each Match In SKUBUILD_ALLMATCHES.DM
                    If prod <> "" And Match = prod Then
                        If CCS_strProds = "" Then CCS_strProds = "'" & prod & "'" Else CCS_strProds = CCS_strProds & ", '" & prod & "'"
                        Exit For
                    End If
                Next
            ElseIf Competitor = "FC" Then
                For Each Match In SKUBUILD_ALLMATCHES.FC
                    If prod <> "" And Match = prod Then
                        If CCS_strProds = "" Then CCS_strProds = "'" & prod & "'" Else CCS_strProds = CCS_strProds & ", '" & prod & "'"
                        Exit For
                    End If
                Next
            End If
            End If
        End If
    Next
    Err.Clear
    On Error GoTo Err_Routine
    If CCS_strProds = "" And CBA_COM_getonlymatchedSKU = True Then
        If Competitor = "WW" Then
            For Each Match In SKUBUILD_ALLMATCHES.WW
                If CCS_strProds = "" Then CCS_strProds = "'" & Match & "'" Else CCS_strProds = CCS_strProds & ", '" & Match & "'"
            Next
        ElseIf Competitor = "C" Then
            For Each Match In SKUBUILD_ALLMATCHES.Coles
                If CCS_strProds = "" Then CCS_strProds = "'" & Match & "'" Else CCS_strProds = CCS_strProds & ", '" & Match & "'"
            Next
        ElseIf Competitor = "DM" Then
            For Each Match In SKUBUILD_ALLMATCHES.DM
                If CCS_strProds = "" Then CCS_strProds = "'" & Match & "'" Else CCS_strProds = CCS_strProds & ", '" & Match & "'"
            Next
        ElseIf Competitor = "FC" Then
            For Each Match In SKUBUILD_ALLMATCHES.FC
                If CCS_strProds = "" Then CCS_strProds = "'" & Match & "'" Else CCS_strProds = CCS_strProds & ", '" & Match & "'"
            Next
        End If
    End If
    
    If CCS_strProds = "" Then
        If CBA_COM_getonlymatchedSKU = True Then
            ReDim SKUarr(1 To 1)
            SKUarr(1).CBA_COM_SKUFormulate Competitor, "", "No Data Found", "", Date
            GoTo Finish
        Else
            CCS_strProds_Exists = False
        End If
    Else
        CCS_strProds_Exists = True
    End If

    
    Set CBA_COM_SKU_COMCN = New ADODB.Connection
    With CBA_COM_SKU_COMCN
        .ConnectionTimeout = 100
        .CommandTimeout = 300
        .Open "Provider= SQLNCLI10; DATA SOURCE=" & CBA_BasicFunctions.TranslateServerName("599DBL12", Date) & "; ;INTEGRATED SECURITY=sspi;"
    End With

    Set CBA_COM_SKU_COMRS = New ADODB.Recordset
    
    WedCount = 0
    Set WedDates = New Scripting.Dictionary
    For DT = DateAdd("D", -15, Date) To Date
        If WeekDay(DT, vbWednesday) = 1 Then
            WedCount = WedCount + 1
            WedDates.Add WedCount, DT
        End If
    Next
    
    
    
    If Competitor = "WW" Then
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_NULLS , ANSI_WARNINGS ON " & Chr(10)
        strSQL = strSQL & "select distinct p.stockcode as ccode,p.name, p.Packagesize as Packsize into #PRODS from tools.dbo.com_w_prod p " & Chr(10)
        For a = 1 To WedCount
            If a = 1 Then
                strSQL = strSQL & "where (datescraped = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
            ElseIf a = WedCount Then
                strSQL = strSQL & "or datescraped = '" & Format(WedDates(a), "YYYY-MM-DD") & "')" & Chr(10)
            Else
                strSQL = strSQL & "or datescraped = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
            End If
        Next
        
        If CCS_strProds_Exists = True Then strSQL = strSQL & "and p.stockcode in (" & CCS_strProds & ")" & Chr(10)
        strSQL = strSQL & "group by  p.stockcode, p.name,p.Packagesize order by p.stockcode" & Chr(10)
        strSQL = strSQL & "select p.stockcode as ccode, p.DateScraped, p.Price, case  when (p.IsEdrSpecial = 1 or p.SavingsAmount <> 0) then 'Promo'" & Chr(10)
        strSQL = strSQL & "when Charindex('LP|',MultiItemPrice,1) > 0 then 'EDLP'" & Chr(10)
        strSQL = strSQL & "when Charindex('PD|',MultiItemPrice,1) > 0 then 'EDLP'" & Chr(10)
        strSQL = strSQL & "Else '' end as MetaData, p.AddressID, MultiItemPrice" & Chr(10)
        strSQL = strSQL & "into #BASE from tools.dbo.com_w_prod p  inner join #PRODS pr on pr.ccode = p.Stockcode " & Chr(10)
        For a = 1 To WedCount
            If a = 1 Then
                strSQL = strSQL & "where (datescraped = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
            ElseIf a = WedCount Then
                strSQL = strSQL & "or datescraped = '" & Format(WedDates(a), "YYYY-MM-DD") & "')" & Chr(10)
            Else
                strSQL = strSQL & "or datescraped = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
            End If
        Next
        strSQL = strSQL & "select addressID into #USEDADD from #BASE group by AddressID" & Chr(10)
                        
        strSQL = strSQL & "select s.addressid ,case when substring(convert(nvarchar(max), s.Postcode),1,1) = '1' then 'NSW'" & Chr(10)
        strSQL = strSQL & "when substring(convert(nvarchar(max), s.Postcode),1,1) = '2' then 'NSW'" & Chr(10)
        strSQL = strSQL & "when substring(convert(nvarchar(max), s.Postcode),1,1) = '3' then 'VIC'" & Chr(10)
        strSQL = strSQL & "when substring(convert(nvarchar(max), s.Postcode),1,1) = '4' then 'QLD'" & Chr(10)
        strSQL = strSQL & "when substring(convert(nvarchar(max), s.Postcode),1,1) = '5' then 'SA'" & Chr(10)
        strSQL = strSQL & "when substring(convert(nvarchar(max), s.Postcode),1,1) = '6' then 'WA'" & Chr(10)
        strSQL = strSQL & "when substring(convert(nvarchar(max), s.Postcode),1,1) = '8' then 'VIC'" & Chr(10)
        strSQL = strSQL & "when substring(convert(nvarchar(max), s.Postcode),1,1) = '9' then 'QLD'" & Chr(10)
        strSQL = strSQL & "Else 'NAT' end as 'State'into #State from tools.dbo.com_w_stores s" & Chr(10)
        strSQL = strSQL & "inner join #USEDADD u on u.AddressID = s.AddressID" & Chr(10)
                        
        strSQL = strSQL & "select p.ccode, p.DateScraped, p.Price, p.MetaData, s.state , p.MultiItemPrice " & Chr(10)
        strSQL = strSQL & "into #OP from #BASE p" & Chr(10)
        strSQL = strSQL & "left join #State s on p.AddressID = s.AddressID" & Chr(10)
        'Debug.Print strSQL
        CBA_COM_SKU_COMRS.Open strSQL, CBA_COM_SKU_COMCN
    ElseIf Competitor = "C" Then
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_NULLS , ANSI_WARNINGS ON " & Chr(10)
        strSQL = strSQL & "select distinct p.colesproductid as ccode, p.brand + ' ' + p.Name as Name, p.Packsize into #PRODS from tools.dbo.com_c_prod p " & Chr(10)
        For a = 1 To WedCount
            If a = 1 Then
                strSQL = strSQL & "where (datescraped = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
            ElseIf a = WedCount Then
                strSQL = strSQL & "or datescraped = '" & Format(WedDates(a), "YYYY-MM-DD") & "')" & Chr(10)
            Else
                strSQL = strSQL & "or datescraped = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
            End If
        Next
        If CCS_strProds_Exists = True Then strSQL = strSQL & "and p.colesproductid in (" & CCS_strProds & ")" & Chr(10)
        strSQL = strSQL & "group by  p.colesproductid, p.brand , p.Name, p.Packsize" & Chr(10)
        strSQL = strSQL & "order by  p.colesproductid" & Chr(10)
        strSQL = strSQL & "select p.colesproductid as ccode, p.DateScraped, p.Price, case  when charindex('Special',p.MultiItemPrice,1) > 0 then 'Promo' " & Chr(10)
        strSQL = strSQL & "when charindex('EveryDay',p.MultiItemPrice,1) > 0 then 'EDLP'" & Chr(10)
        strSQL = strSQL & "when charindex('X',p.MultiItemPrice,1) > 0 then 'EDLP'" & Chr(10)
        strSQL = strSQL & "when charindex('N',p.MultiItemPrice,1) > 0 then 'New' else '' end as MetaData" & Chr(10)
        strSQL = strSQL & ",convert(int,substring(p.colesproductid,1,len(p.colesproductid)-1)) as ref" & Chr(10)
        strSQL = strSQL & ",case when substring(convert(nvarchar(max), p.UrlScan_StoreSeoToken),3,3) = 'NSW' then 'NSW'" & Chr(10)
        strSQL = strSQL & "when substring(convert(nvarchar(max), p.UrlScan_StoreSeoToken),3,3) = 'VIC' then 'VIC'" & Chr(10)
        strSQL = strSQL & "when substring(convert(nvarchar(max), p.UrlScan_StoreSeoToken),3,3) = 'QLD' then 'QLD'" & Chr(10)
        strSQL = strSQL & "when substring(convert(nvarchar(max), p.UrlScan_StoreSeoToken),3,2) = 'SA' then 'SA'" & Chr(10)
        strSQL = strSQL & "when substring(convert(nvarchar(max), p.UrlScan_StoreSeoToken),3,2) = 'WA' then 'WA' else 'Unknown'end as 'State', '' as MBopt" & Chr(10)
        strSQL = strSQL & "into #OP from tools.dbo.Com_C_Prod p  inner join #PRODS pr on pr.ccode = p.ColesProductId " & Chr(10)
        For a = 1 To WedCount
            If a = 1 Then
                strSQL = strSQL & "where (datescraped = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
            ElseIf a = WedCount Then
                strSQL = strSQL & "or datescraped = '" & Format(WedDates(a), "YYYY-MM-DD") & "')" & Chr(10)
            Else
                strSQL = strSQL & "or datescraped = '" & Format(WedDates(a), "YYYY-MM-DD") & "'" & Chr(10)
            End If
        Next
        'Debug.Print strSQL
        CBA_COM_SKU_COMRS.Open strSQL, CBA_COM_SKU_COMCN
    
    ElseIf Competitor = "DM" Then
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_NULLS , ANSI_WARNINGS ON " & Chr(10)
        strSQL = strSQL & "select productid as CCode, brand + ' ' + name as name, 'each' as packsize" & Chr(10)
        strSQL = strSQL & "into #PRODS from tools.dbo.com_dm_prod" & Chr(10)
        strSQL = strSQL & "where datescraped >= DateAdd(Day, -15, getdate())" & Chr(10)
        If CCS_strProds_Exists = True Then strSQL = strSQL & "and productid in (" & CCS_strProds & ")" & Chr(10)
        strSQL = strSQL & "group by productid, brand, name order by productid" & Chr(10)
        strSQL = strSQL & "select ProductId as ccode, DateScraped" & Chr(10)
        strSQL = strSQL & ", case  when CHARINDEX('each',SpecialSizeDescription1,1) > 0 then SpecialPrice1" & Chr(10)
        strSQL = strSQL & "        when CHARINDEX('in bottle',SpecialSizeDescription1,1) > 0 then SpecialPrice1" & Chr(10)
        strSQL = strSQL & "        when CHARINDEX('per bottle',SpecialSizeDescription1,1) > 0 then SpecialPrice1" & Chr(10)
        strSQL = strSQL & "        when CHARINDEX('per can',SpecialSizeDescription1,1) > 0 then SpecialPrice1" & Chr(10)
        strSQL = strSQL & "        when CHARINDEX('per case of 1',SpecialSizeDescription1,1) > 0 then SpecialPrice1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per cask',SpecialSizeDescription1,1) > 0 then SpecialPrice1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per collection',SpecialSizeDescription1,1) > 0 then SpecialPrice1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per decanter',SpecialSizeDescription1,1) > 0 then SpecialPrice1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per giftpack',SpecialSizeDescription1,1) > 0 then SpecialPrice1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per jar ',SpecialSizeDescription1,1) > 0 then SpecialPrice1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per keg',SpecialSizeDescription1,1) > 0 then SpecialPrice1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('each',SizeDescription1,1) > 0 then Price1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('in bottle',SizeDescription1,1) > 0 then Price1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per bottle',SizeDescription1,1) > 0 then Price1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per can',SizeDescription1,1) > 0 then Price1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per case of 1',SizeDescription1,1) > 0 then Price1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per cask',SizeDescription1,1) > 0 then Price1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per collection',SizeDescription1,1) > 0 then Price1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per decanter',SizeDescription1,1) > 0 then Price1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per giftpack',SizeDescription1,1) > 0 then Price1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per jar ',SizeDescription1,1) > 0 then Price1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per keg',SizeDescription1,1) > 0 then Price1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('each',SizeDescription2,1) > 0 then Price2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('in bottle',SizeDescription2,1) > 0 then Price2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per bottle',SizeDescription2,1) > 0 then Price2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per can',SizeDescription2,1) > 0 then Price2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per case of 1',SizeDescription2,1) > 0 then Price2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per cask',SizeDescription2,1) > 0 then Price2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per collection',SizeDescription2,1) > 0 then Price2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per decanter',SizeDescription2,1) > 0 then Price2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per giftpack',SizeDescription2,1) > 0 then Price2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per jar ',SizeDescription2,1) > 0 then Price2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per keg',SizeDescription2,1) > 0 then Price2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('each',SizeDescription3,1) > 0 then Price3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('in bottle',SizeDescription3,1) > 0 then Price3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per bottle',SizeDescription3,1) > 0 then Price3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per can',SizeDescription3,1) > 0 then Price3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per case of 1',SizeDescription3,1) > 0 then Price3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per cask',SizeDescription3,1) > 0 then Price3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per collection',SizeDescription3,1) > 0 then Price3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per decanter',SizeDescription3,1) > 0 then Price3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per giftpack',SizeDescription3,1) > 0 then Price3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per jar ',SizeDescription3,1) > 0 then Price3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per keg',SizeDescription3,1) > 0 then Price3" & Chr(10)
        strSQL = strSQL & "Else '' end as Price" & Chr(10)
        strSQL = strSQL & ", case  when (CHARINDEX('OOS|',SpecialRequirements1,1) + CHARINDEX('LA|',SpecialRequirements1,1)) > 0 and (CHARINDEX('O|',SpecialRequirements1,1) + CHARINDEX('MBS|',SpecialRequirements1,1))  > 0" & Chr(10)
        strSQL = strSQL & "and CHARINDEX('ORG|',SpecialRequirements1,1) > 0 then 'Low/No Stock, Promo, Organic'" & Chr(10)
        strSQL = strSQL & "when (CHARINDEX('OOS|',SpecialRequirements1,1) + CHARINDEX('LA|',SpecialRequirements1,1)) > 0 and (CHARINDEX('O|',SpecialRequirements1,1) + CHARINDEX('MBS|',SpecialRequirements1,1))  > 0" & Chr(10)
        strSQL = strSQL & "then 'Low/No Stock, Promo'" & Chr(10)
        strSQL = strSQL & "when (CHARINDEX('O|',SpecialRequirements1,1) + CHARINDEX('MBS|',SpecialRequirements1,1))  > 0 and CHARINDEX('ORG|',SpecialRequirements1,1) > 0 then 'Promo, Organic'" & Chr(10)
        strSQL = strSQL & "when (CHARINDEX('OOS|',SpecialRequirements1,1) + CHARINDEX('LA|',SpecialRequirements1,1)) > 0 and CHARINDEX('ORG|',SpecialRequirements1,1) > 0 then 'Low/No Stock, Organic'" & Chr(10)
        strSQL = strSQL & "when (CHARINDEX('OOS|',SpecialRequirements1,1) + CHARINDEX('LA|',SpecialRequirements1,1)) > 0 then 'Low/No Stock'" & Chr(10)
        strSQL = strSQL & "when (CHARINDEX('O|',SpecialRequirements1,1) + CHARINDEX('MBS|',SpecialRequirements1,1))  > 0 then 'Promo'" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('ORG|',SpecialRequirements1,1) > 0 then 'Organic'" & Chr(10)
        strSQL = strSQL & "Else '' end as MetaData" & Chr(10)
        strSQL = strSQL & ", case when s.StoreScan_StoreState = 'ACT' then 'NSW' else s.StoreScan_StoreState end as state" & Chr(10)
        strSQL = strSQL & ", case  when CHARINDEX('for',lower(SpecialRequirements1),1) > 0 then SpecialRequirements1 + ' ' +  convert(nvarchar(20),SpecialPrice1)" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per case',SpecialSizeDescription1,1) > 0 and not CHARINDEX('per case of 1',SpecialSizeDescription1,1) > 0 then convert(nvarchar(20),SpecialPrice1) + ' ' + SpecialSizeDescription1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('in any six',SpecialSizeDescription1,1) > 0 then convert(nvarchar(20),SpecialPrice1) + ' ' + SpecialSizeDescription1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per pack',SpecialSizeDescription1,1) > 0 then convert(nvarchar(20),SpecialPrice1) + ' ' + SpecialSizeDescription1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per bundle',SpecialSizeDescription1,1) > 0 then convert(nvarchar(20),SpecialPrice1)  + ' ' + SpecialSizeDescription1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('FOR',SpecialSizeDescription1,1) > 0 then convert(nvarchar(20),SpecialPrice1)  + ' ' + SpecialSizeDescription1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per case',SizeDescription1,1) > 0 and not CHARINDEX('per case of 1',SizeDescription1,1) > 0 then convert(nvarchar(20),Price1) + ' ' + SizeDescription1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('in any six',SizeDescription1,1) > 0 then convert(nvarchar(20),Price1) + ' ' + SizeDescription1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per pack',SizeDescription1,1) > 0 then convert(nvarchar(20),Price1) + ' ' + SizeDescription1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per bundle',SizeDescription1,1) > 0 then convert(nvarchar(20),Price1)  + ' ' + SizeDescription1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('FOR',SizeDescription1,1) > 0 then convert(nvarchar(20),Price1)  + ' ' + SizeDescription1" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per case',SizeDescription2,1) > 0 and not CHARINDEX('per case of 1',SizeDescription2,1) > 0 then convert(nvarchar(20),Price2) + ' ' + SizeDescription2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('in any six',SizeDescription2,1) > 0 then convert(nvarchar(20),Price2) + ' ' + SizeDescription2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per pack',SizeDescription2,1) > 0 then convert(nvarchar(20),Price2) + ' ' + SizeDescription2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per bundle',SizeDescription2,1) > 0 then convert(nvarchar(20),Price2)  + ' ' + SizeDescription2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('FOR',SizeDescription2,1) > 0 then convert(nvarchar(20),Price2)  + ' ' + SizeDescription2" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per case',SizeDescription3,1) > 0 and not CHARINDEX('per case of 1',SizeDescription3,1) > 0 then convert(nvarchar(20),Price3) + ' ' + SizeDescription3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('in any six',SizeDescription3,1) > 0 then convert(nvarchar(20),Price3) + ' ' + SizeDescription3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per pack',SizeDescription3,1) > 0 then convert(nvarchar(20),Price3) + ' ' + SizeDescription3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('per bundle',SizeDescription3,1) > 0 then convert(nvarchar(20),Price3)  + ' ' + SizeDescription3" & Chr(10)
        strSQL = strSQL & "when CHARINDEX('FOR',SizeDescription3,1) > 0 then convert(nvarchar(20),Price3)  + ' ' + SizeDescription3" & Chr(10)
        strSQL = strSQL & "Else '' end as MBopt" & Chr(10)
        strSQL = strSQL & "into #OP from tools.dbo.com_dm_prod p" & Chr(10)
        strSQL = strSQL & "left join tools.dbo.Com_DM_Stores s on s.StoreScan_StoreNo = p.t0_Stores_StoreScan_StoreNo" & Chr(10)
        strSQL = strSQL & "inner join #PRODS pr on pr.ccode = p.ProductId" & Chr(10)
        strSQL = strSQL & "where datescraped >= DateAdd(Day, -" & DaystoQuery & ", getdate())" & Chr(10)
        'strSQL = "drop table #PRODS, #OP"
        CBA_COM_SKU_COMRS.Open strSQL, CBA_COM_SKU_COMCN
    ElseIf Competitor = "FC" Then
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_NULLS , ANSI_WARNINGS ON " & Chr(10)
        strSQL = strSQL & "select p.productid as ccode, p.Brand + ' ' + p.Description as name, 'each' as Packsize into #PRODS from tools.dbo.com_FC_prod p " & Chr(10)
        strSQL = strSQL & "where datescraped >= dateadd(DAY,-15,getdate()) " & Chr(10)
        If CCS_strProds_Exists = True Then strSQL = strSQL & "and p.productid in (" & CCS_strProds & ")" & Chr(10)
        strSQL = strSQL & "group by  p.productid, p.Brand + ' ' + p.Description order by p.productid" & Chr(10)
        strSQL = strSQL & "select p.productid as ccode, p.DateScraped, p.PriceEach" & Chr(10)
        strSQL = strSQL & ", case  when charindex('SPECIAL',p.Special,1) > 0 then 'Promo'" & Chr(10)
        strSQL = strSQL & "        when charindex('% OFF',p.Special,1) > 0 then 'Promo'" & Chr(10)
        strSQL = strSQL & "        when charindex('SAVE',p.Special,1) > 0 then 'Promo'" & Chr(10)
        strSQL = strSQL & "        when charindex('POINTS',p.Special,1) > 0 then 'Points'" & Chr(10)
        strSQL = strSQL & "        when charindex('INTRO PRICE',p.Special,1) > 0 then 'New'" & Chr(10)
        strSQL = strSQL & "        when charindex('NEW',p.Special,1) > 0 then 'New'" & Chr(10)
        strSQL = strSQL & "        when charindex('Stock',p.Special,1) > 0 then 'Low/No Stock'" & Chr(10)
        strSQL = strSQL & "        else ''" & Chr(10)
        strSQL = strSQL & "end as Metadata ,s.State, case  when charindex('For',p.Special,1) > 0 then p.Special else '' end as MBopt " & Chr(10)
        strSQL = strSQL & "into #OP from tools.dbo.Com_FC_Prod p" & Chr(10)
        strSQL = strSQL & "left join tools.dbo.Com_FC_Stores s on s.StoreNo = p.StoreNo" & Chr(10)
        strSQL = strSQL & "inner join #PRODS pr on pr.ccode = p.ProductId where p.DateScraped >= dateadd(D,-" & DaystoQuery & ",getdate())" & Chr(10)
        CBA_COM_SKU_COMRS.Open strSQL, CBA_COM_SKU_COMCN
    End If
    
    Set CBA_COM_SKU_COMRS = New ADODB.Recordset
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_NULLS , ANSI_WARNINGS ON " & Chr(10)
    strSQL = strSQL & "select ccode, name, Packsize from #PRODS order by ccode" & Chr(10)
    CBA_COM_SKU_COMRS.Open strSQL, CBA_COM_SKU_COMCN
    CBA_COM_SKU_arrprodmap = CBA_COM_SKU_COMRS.GetRows
    CBA_COM_SKU_COMRS.Close
    
    Set CBA_COM_SKU_COMRS = New ADODB.Recordset
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_NULLS , ANSI_WARNINGS ON " & Chr(10)
    strSQL = strSQL & "select ccode,max(datescraped) from #OP group by ccode order by ccode" & Chr(10)
    
    CBA_COM_SKU_COMRS.Open strSQL, CBA_COM_SKU_COMCN
    CBA_COM_SKU_LastSeen = CBA_COM_SKU_COMRS.GetRows
    CBA_COM_SKU_COMRS.Close
        
    If isRunningSheetDisplayed = True Then CBA_COM_SKU_Run = True Else CBA_COM_SKU_Run = False
    If CCS_strProds_Exists = False And CBA_COM_SKU_Run = False Then
        If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
        CBA_COM_SKU_Run = True
    End If
    strProds = "": PMsp = -1: LSsp = -1: SKU_CNT = 0
    'If competitor = "C" Then buildrate = 1300 Else
    buildrate = 650
    tot = UBound(CBA_COM_SKU_arrprodmap, 2) + 1
    ReDim SKUarr(1 To tot)
    For a = 1 To tot
        If a <= tot Then
            If strProds = "" Then
                If Competitor = "C" Then
                    strProds = Mid(CBA_COM_SKU_arrprodmap(0, a - 1), 1, Len(CBA_COM_SKU_arrprodmap(0, a - 1)) - 1)
                Else
                    'competitor = "WW" Or competitor = "FC" Or competitor = "DM" Then
                    strProds = "'" & CBA_COM_SKU_arrprodmap(0, a - 1) & "'"
                End If
            Else
                If Competitor = "C" Then
                    strProds = strProds & ", " & Mid(CBA_COM_SKU_arrprodmap(0, a - 1), 1, Len(CBA_COM_SKU_arrprodmap(0, a - 1)) - 1)
                Else 'If competitor = "DM" Then
                    strProds = strProds & ", '" & CBA_COM_SKU_arrprodmap(0, a - 1) & "'"
                End If
                
            End If
        End If
        If Round(a / buildrate, 0) = a / buildrate Or a = tot Then
            If tot = 1 Then
                If Competitor = "C" Then
                    strProds = Mid(CBA_COM_SKU_arrprodmap(0, a - 1), 1, Len(CBA_COM_SKU_arrprodmap(0, a - 1)) - 1)
                Else
                    'competitor = "WW" Or competitor = "FC" Or competitor = "DM" Then
                    strProds = "'" & CBA_COM_SKU_arrprodmap(0, a - 1) & "'"
                End If
            End If
            
            If CBA_COM_SKU_COMRS.State = 1 Then CBA_COM_SKU_COMRS.Close
            
            If Competitor = "WW" Or Competitor = "FC" Or Competitor = "DM" Then
                strSQL = "select * from #OP where ccode in (" & strProds & ") order by ccode, DateScraped, State"
            ElseIf Competitor = "C" Then
                
                strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_NULLS , ANSI_WARNINGS ON " & Chr(10)
                strSQL = strSQL & "select distinct ref into #P from #OP where Ref in (" & strProds & ")" & Chr(10)
                strSQL = strSQL & "select ccode, DateScraped, Price,MetaData, State, MBopt from #OP op inner join #P p on p.ref = op.ref" & Chr(10)
                strSQL = strSQL & "order by ccode, DateScraped, State" & Chr(10)
                strSQL = strSQL & "drop table #P" & Chr(10)
                    'OLD CODE: strSQL = "select ccode, DateScraped, Price,MetaData, State, MBopt from #OP where ref in (" & strProds & ") order by ccode, DateScraped, State"
            End If
            CBA_COM_SKU_COMRS.Open strSQL, CBA_COM_SKU_COMCN, adOpenForwardOnly
            For SKU_NUM = (a - (buildrate - 1)) To a
                If SKU_NUM < 0 Then SKU_NUM = 1
                For PMcur = PMsp + 1 To UBound(CBA_COM_SKU_arrprodmap, 2)
                    If CBA_COM_SKU_COMRS.Fields(0) = CBA_COM_SKU_arrprodmap(0, PMcur) Then
                        PMsp = PMcur
                        Exit For
                    End If
                Next
                For LScur = LSsp + 1 To UBound(CBA_COM_SKU_LastSeen, 2)
                    If CBA_COM_SKU_COMRS.Fields(0) = CBA_COM_SKU_LastSeen(0, LScur) Then
                        LSsp = LScur
                        Exit For
                    End If
                Next
                SKU_CNT = SKU_CNT + 1
                
                SKUarr(SKU_CNT).CBA_COM_SKUFormulate Competitor, CBA_COM_SKU_arrprodmap(0, PMsp), CBA_COM_SKU_arrprodmap(1, PMsp), CBA_COM_SKU_arrprodmap(2, PMsp), CBA_COM_SKU_LastSeen(1, LSsp)
                If CBA_COM_SKU_COMRS.EOF = True Then Exit For
                'debug.Print Timer - thistime
            Next
            
            If CBA_COM_SKU_Run = True Then
                If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 9, 3, "Building " & compnamefull & " SKU Objects " & SKU_NUM & " of " & tot
            End If
            strProds = ""
        End If
    Next
    
    CBA_COM_SKU_COMCN.Close
    
    Set CBA_COM_SKU_COMCN = Nothing
    ReDim Preserve SKUarr(1 To SKU_CNT)
Finish:
    Application.StatusBar = False
    CBA_SetupSKUArray = SKUarr
    If CCS_strProds_Exists = False And isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("CBA_COM_SetupSKUArray-f-CBA_SetupSKUArray", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
''    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function


