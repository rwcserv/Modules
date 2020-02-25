Attribute VB_Name = "COM_Mousewheel"
Option Explicit     ' Com_MouseWheel - last changed 190627
' This module will work on both ListBoxes and ComboBoxes
' The jury is out as to whether you have to call the mw_RemoveBoxHook at the close of a form as the module does not hold individual copies of Boxes - only one

' The following is put on every Box that needs it - Don't put it on lists or combos that have few items - it isn't worth it
''''Private Sub ListBox2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
''''    Call mw_SetBoxHook(ListBox2)
''''End Sub

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MSLLHOOKSTRUCT
    pt As POINTAPI
    MouseData As Long
    Flags As Long
    Time As Long
    DWExtraInfo As Long
End Type

#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal Point As LongLong) As LongPtr
    #Else
        Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
    #End If
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As LongPtr
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As LongPtr
    Private hwnd As LongPtr, lMouseHook As LongPtr
#Else
    Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
    Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private hwnd As Long, lMouseHook As Long
#End If

Private Const WH_MOUSE_LL = 14
Private Const WM_MOUSEWHEEL = &H20A
Private Const HC_ACTION = 0
Private bOnOff As Boolean
Private oBox As Object

' Set up for the Mousewheel
Public Sub mw_SetBoxHook(ByVal Control As Object)
    Dim tPt As POINTAPI, sBuffer As String, lRet As Long
    Static stcRun As String
    If stcRun = "" Then
        stcRun = g_getConfig("RunMousewheel", g_GetDB("Gen"), , "N")
''        stcRun = "N"                                                        '@RW ..
    End If
    If bOnOff = False Then
''        Debug.Print "on";
        bOnOff = True
    End If
    ' If the config flagged to not to run then exit sub
    If stcRun = "N" Then Exit Sub
''    Debug.Print "m";
    Set oBox = Control
    Call mw_RemoveBoxHook(False)
    GetCursorPos tPt
    #If Win64 Then
        Dim lPt As LongPtr
        CopyMemory lPt, tPt, LenB(tPt)
        hwnd = WindowFromPoint(lPt)
    #Else
        hwnd = WindowFromPoint(tPt.X, tPt.Y)
    #End If
    sBuffer = Space(256)
    lRet = GetClassName(GetParent(hwnd), sBuffer, 256)
    If InStr(Left(sBuffer, lRet), "MdcPopup") Or InStr(Left(sBuffer, lRet), "F3 Server") Then
        ''If InStr(Left(sBuffer, lRet), "MdcPopup") Then SetFocus hwnd
        #If Win64 Then
            lMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, Application.HinstancePtr, 0)
        #Else
            lMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, Application.hInstance, 0)
        #End If
    End If
End Sub

Public Sub mw_RemoveBoxHook(Optional bCalled As Boolean = True)
    If bOnOff = True And bCalled = True Then
''        Debug.Print ",off";
        bOnOff = Not bOnOff
    End If
    UnhookWindowsHookEx lMouseHook
End Sub

#If VBA7 Then
    Private Function MouseProc(ByVal nCode As Long, ByVal wParam As LongPtr, lParam As MSLLHOOKSTRUCT) As LongPtr
#Else
    Private Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, lParam As MSLLHOOKSTRUCT) As Long
#End If

    Dim sBuffer As String
    Dim lRet As Long
        
    sBuffer = Space(256)
    lRet = GetClassName(GetActiveWindow, sBuffer, 256)
    If Left(sBuffer, lRet) = "wndclass_desked_gsk" Then Call mw_RemoveBoxHook
    If IsWindow(hwnd) = 0 Then Call mw_RemoveBoxHook
    
    If (nCode = HC_ACTION) Then
        If wParam = WM_MOUSEWHEEL Then
        #If Win64 Then
            Dim lPt As LongPtr
            CopyMemory lPt, lParam.pt, LenB(lParam.pt)
            If WindowFromPoint(lPt) = hwnd Then
        #Else
            If WindowFromPoint(lParam.pt.X, lParam.pt.Y) = hwnd Then
        #End If
                On Error Resume Next
                    If lParam.MouseData > 0 Then
                        oBox.TopIndex = oBox.TopIndex - 1
                        ''oBox.ListIndex = oBox.ListIndex - 1
                    Else
                        oBox.TopIndex = oBox.TopIndex + 1
                        ''oBox.ListIndex = oBox.ListIndex + 1
                    End If
                On Error GoTo 0
            End If
        End If
    End If
    
''    MouseProc = CallNextHookEx(lMouseHook, nCode, wParam, ByVal lParam)
End Function

