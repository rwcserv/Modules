Attribute VB_Name = "CBA_COM_ReportFunctions"
Option Explicit
Option Private Module          ' Excel users cannot access procedures
Function CBA_COM_DiffToMaintain(ByVal MatchType As String) As Single
    Select Case MatchType
        Case "ColesWeb", "WWWeb"
            CBA_COM_DiffToMaintain = -1
        Case "ColesSB3", "ColesSB2", "ColesSB1", "ColesVal3", "ColesVal2", "ColesVal1", "ColesWNAT1", "ColesWNAT2", "ColesWNAT3", "ColesWNAT4", "ColesWNSW", "ColesWQLD", "ColesWSA", "ColesWVIC", "ColesWWA", "WWWNAT1", "WWWNAT2", "WWWNAT3", "WWWNAT4", "WWWNSW", "WWWQLD", "WWWSA", "WWWVIC", "WWWWA", "WWHB1", "WWHB2", "WWHB3"
            CBA_COM_DiffToMaintain = 0
        Case "ColesPL4", "ColesPL3", "ColesPL2", "ColesPL1", "ColesPB2", "ColesPB1", "ColesCB1", "FC1", "FC2", "DM1", "DM2", "WWCB1", "WWPB1", "WWPB2", "WWWW1", "WWWW2", "WWWW3", "WWWW4", "WWWW5"
            CBA_COM_DiffToMaintain = 0.1
        Case "ColesML3", "ColesML2", "ColesML1", "FCQ", "DMQ", "WWML1", "WWML2", "WWML3"
            CBA_COM_DiffToMaintain = 0.3
    End Select
End Function
Function CBA_COM_isTrafficLight(ByVal Diff As Single, ByVal Mtype As String) As Long
    Dim CBA_Proc As String
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
        If InStr(1, LCase(Mtype), "ml") > 0 Then
            'std red if less that 30% / yellow if 30-40% / green if 40% or more / blue if 50% or more
            If Diff < 0.3 Then
                CBA_COM_isTrafficLight = 3
            ElseIf Diff < 0.4 Then
                CBA_COM_isTrafficLight = 36
            ElseIf Diff < 0.5 Then
                CBA_COM_isTrafficLight = 43
            Else
                CBA_COM_isTrafficLight = 37
            End If

        ElseIf InStr(1, LCase(Mtype), "pb") > 0 Or InStr(1, LCase(Mtype), "cb") > 0 Or InStr(1, LCase(Mtype), "colespl") > 0 Or InStr(1, LCase(Mtype), "colescoles") > 0 Or InStr(1, LCase(Mtype), "wwselect") > 0 Or InStr(1, LCase(Mtype), "wwww") > 0 Then
            'std red if less than 10% / yellow if 10-15% / green if 15% or more
            If Diff < 0.1 Then
                CBA_COM_isTrafficLight = 3
            ElseIf Diff < 0.15 Then
                CBA_COM_isTrafficLight = 36
            Else
                CBA_COM_isTrafficLight = 43
            End If

        ElseIf InStr(1, LCase(Mtype), "colessb") > 0 Or InStr(1, LCase(Mtype), "colesval") > 0 Or InStr(1, LCase(Mtype), "wwhb") > 0 Or Mtype = "DM1" Or Mtype = "DM2" Or Mtype = "FC1" Or Mtype = "FC2" Then
            'std red if less than 0% / yellow if 0-5% / green if 5% or more
            If Diff < 0 Then
                CBA_COM_isTrafficLight = 3
            ElseIf Diff < 0.05 Then
                CBA_COM_isTrafficLight = 36
            Else
                CBA_COM_isTrafficLight = 43
            End If
        ElseIf Mtype = "DMQ" Or Mtype = "FCQ" Then
            CBA_COM_isTrafficLight = 35
        ElseIf InStr(1, LCase(Mtype), "colesw") > 0 Or InStr(1, LCase(Mtype), "www") > 0 Then
            'std red if less that 0%
            If Diff < 0 Then CBA_COM_isTrafficLight = 3
        End If
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-CBA_COM_isTrafficLight", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function
