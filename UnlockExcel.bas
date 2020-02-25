Attribute VB_Name = "UnlockExcel"

' modUnlockRoutines
'
' Module provides Excel workbook and sheet unlock routines. The algorithm
' relies on a backdoor password that can be 1 to 9 characters long where each
' character is either an "A" or "B" except the last which can be any character
' from ASCII code 32 to 255.
'
' Implemented as a regular module for use with any Excel VBA project.

Option Explicit
 
Private Sub DisplayStatus( _
      ByVal PasswordsTried As Long _
   )
 
' Display the status in the Excel status bar.
'
' Syntax
'
' DisplayStatus(PasswordsTried)
'
' PasswordsTried - The number of passwords tried thus far.
 
   Static LastStatus As String
 
   LastStatus = Format(PasswordsTried / 57120, "0%") & " of possible passwords tried."
   If Application.StatusBar <> LastStatus Then
      Application.StatusBar = LastStatus
      DoEvents
   End If
 
End Sub
 
Private Function TrySheetPasswordSize( _
      ByVal Size As Long, _
      ByRef PasswordsTried As Long, _
      ByRef Password As String, _
      Optional ByVal Base As String _
   ) As Boolean
 
' Try unlocking the sheet with all passwords of the specified size.
'
' TrySheetPasswordSize(Size, PasswordsTried, Password, [Base])
'
' Size - The size of the password to try.
'
' PasswordsTried - The cummulative number of passwords tried thus far.
'
' Password - The current password.
'
' Base - The base password from the calling routine.
  
   Dim index As Long
  
   On Error Resume Next
   If IsMissing(Base) Then Base = vbNullString
   If Len(Base) < Size - 1 Then
      For index = 65 To 66
         If TrySheetPasswordSize(Size, PasswordsTried, Password, Base & Chr(index)) Then
            TrySheetPasswordSize = True
            Exit Function
         End If
      Next index
   ElseIf Len(Base) < Size Then
      For index = 32 To 255
         ActiveSheet.Unprotect Base & Chr(index)
         If Not ActiveSheet.ProtectContents Then
            TrySheetPasswordSize = True
            Password = Base & Chr(index)
            Exit Function
         End If
         PasswordsTried = PasswordsTried + 1
      Next index
   End If
   On Error GoTo 0
  
   DisplayStatus PasswordsTried
 
End Function
 
Private Function TryWorkbookPasswordSize( _
      ByVal Size As Long, _
      ByRef PasswordsTried As Long, _
      ByRef Password As String, _
      Optional ByVal Base As String _
   ) As Boolean
  
' Try unlocking the workbook with all passwords of the specified size.
'
' TryWorkbookPasswordSize(Size, PasswordsTried, Password, [Base])
'
' Size - The size of the password to try.
'
' PasswordsTried - The cummulative number of passwords tried thus far.
'
' Password - The current password.
'
' Base - The base password from the calling routine.
  
   Dim index As Long
  
   On Error Resume Next
   If IsMissing(Base) Then Base = vbNullString
   If Len(Base) < Size - 1 Then
      For index = 65 To 66
         If TryWorkbookPasswordSize(Size, PasswordsTried, Password, Base & Chr(index)) Then
            TryWorkbookPasswordSize = True
            Exit Function
         End If
      Next index
   ElseIf Len(Base) < Size Then
      For index = 32 To 255
         ActiveWorkbook.Unprotect Base & Chr(index)
         If Not ActiveWorkbook.ProtectStructure And Not ActiveWorkbook.ProtectWindows Then
            TryWorkbookPasswordSize = True
            Password = Base & Chr(index)
            Exit Function
         End If
         PasswordsTried = PasswordsTried + 1
      Next index
   End If
   On Error GoTo 0
  
   DisplayStatus PasswordsTried
 
End Function
 
Public Sub UnlockSheet()
 
' Unlock the active sheet using a backdoor Excel provides where an alternate password is created that is more limited.
 
   Dim PasswordSize As Variant
   Dim PasswordsTried As Long
   Dim Password As String
 
   PasswordsTried = 0
   If Not ActiveSheet.ProtectContents Then
      MsgBox "The sheet is already unprotected."
      Exit Sub
   End If
   On Error Resume Next
   ActiveSheet.Protect ""
   ActiveSheet.Unprotect ""
   On Error GoTo 0
   If ActiveSheet.ProtectContents Then
      For Each PasswordSize In Array(5, 4, 6, 7, 8, 3, 2, 1)
         If TrySheetPasswordSize(PasswordSize, PasswordsTried, Password) Then Exit For
      Next PasswordSize
   End If
   If Not ActiveSheet.ProtectContents Then
      MsgBox "The sheet " & ActiveSheet.Name & " has been unprotected with password '" & Password & "'."
   End If
   Application.StatusBar = False
  
End Sub
 
Public Sub UnlockWorkbook()
 
' Unlock the active workbook using a backdoor Excel provides where an alternate password is created that is more limited.
 
   Dim PasswordSize As Variant
   Dim PasswordsTried As Long
   Dim Password As String
  
   PasswordsTried = 0
   If Not ActiveWorkbook.ProtectStructure And Not ActiveWorkbook.ProtectWindows Then
      MsgBox "The workbook is already unprotected."
      Exit Sub
   End If
   On Error Resume Next
   ActiveWorkbook.Unprotect vbNullString
   On Error GoTo 0
   If ActiveWorkbook.ProtectStructure Or ActiveWorkbook.ProtectWindows Then
      For Each PasswordSize In Array(5, 4, 6, 7, 8, 3, 2, 1)
         If TryWorkbookPasswordSize(PasswordSize, PasswordsTried, Password) Then Exit For
      Next PasswordSize
   End If
   If Not ActiveWorkbook.ProtectStructure And Not ActiveWorkbook.ProtectWindows Then
      MsgBox "The workbook " & ActiveWorkbook.Name & " has been unprotected with password '" & Password & "'."
   End If
   Application.StatusBar = False
  
End Sub
Sub HackVBProject()
    If Hook Then
        MsgBox "VBA Project is unprotected!", vbInformation, "*****"
    End If
End Sub
