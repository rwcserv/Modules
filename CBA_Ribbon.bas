Attribute VB_Name = "CBA_Ribbon"
Option Explicit
#If VBA7 Then
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#Else
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#End If
'Callback for customUI.onLoad
Sub CBA_OnLoad(CBA_Ribbon As IRibbonUI)
    Set CBA_Rib = CBA_Ribbon
    CBA_DataSheet.Range("A1").Value = ObjPtr(CBA_Ribbon)
    CBA_COM_Runtime.CBA_COM_SetToggleButtonState True
    CBA_COM_Runtime.CBA_COM_SetRibbonState False
    CBA_COM_Runtime.CBA_COM_SetRibbonSCGState False
    CBA_COM_Runtime.CBA_COM_SetMatchingToolState False
    
    'If you want to run a macro below when you open the workbook
    'you can call the macro like this :
    'Call EnableControlsWithCertainTag3
''    Call CBA_AST_Start '@RWAST 191205 Took out for now as doesn't work and may compromise Excel integrity
End Sub

''Public Sub CBA_AST_Start()
''    ' Is run at the start to enable the aspects that are relevent
''    CBA_lAuthority = AST_getUserASystAuthority(CBA_SetUser, True, True)
''    If CBA_lAuthority <> 1 And CBA_lAuthority <> 5 Then
''        CBA_Rib.InvalidateControl "CBA_ASYST_Promo"
''        CBA_Ribbon.CBA_RefreshRibbon
''    End If
''
''End Sub

#If VBA7 Then
Function CBA_GetRibbon(ByVal CBA_lRibbonPointer As LongPtr) As Object
#Else
Function CBA_GetRibbon(ByVal CBA_lRibbonPointer As Long) As Object
#End If
        Dim CBA_objRibbon As Object
        CopyMemory CBA_objRibbon, CBA_lRibbonPointer, LenB(CBA_lRibbonPointer)
        Set CBA_GetRibbon = CBA_objRibbon
        Set CBA_objRibbon = Nothing
        CBA_COM_Runtime.CBA_COM_SetToggleButtonState True
        CBA_COM_Runtime.CBA_COM_SetRibbonState False
        CBA_COM_Runtime.CBA_COM_SetMatchingToolState False
End Function
Sub CBA_RefreshRibbon()
    If CBA_Rib Is Nothing Then
        Set CBA_Rib = CBA_GetRibbon(CBA_DataSheet.Range("A1").Value)
        CBA_Rib.Invalidate
    Else
        CBA_Rib.Invalidate
    End If
End Sub
Sub GetAdminUser(Control As IRibbonControl, ByRef bVisible)
    bVisible = CBA_BasicFunctions.CBA_getAdminUsers
End Sub
Sub GetMerchUser(Control As IRibbonControl, ByRef bVisible)
Dim a As String, b As String
    CBA_BasicFunctions.CBA_getUserShortTitle a, b
    If InStr(1, LCase(b), "merch") > 0 Or CBA_BasicFunctions.CBA_getAdminUsers Then
        bVisible = True
    Else
        bVisible = False
    End If
End Sub
Sub CBA_StartSTAR(Control As IRibbonControl)
    CBA_STAR.Show
End Sub
