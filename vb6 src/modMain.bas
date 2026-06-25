Attribute VB_Name = "modMain"
Option Explicit

Private Const CMD_SHUTDOWN = "uyhgbrf"

Sub Main()
  InitCommonControls
  Dim r As Boolean
  If LCase$(Command$) = CMD_SHUTDOWN Then r = False Else r = True
  If r = True Then RegisterRunOnce
  TriggerShutdown r
End Sub

Private Sub RegisterRunOnce()
  Dim hKey As Long, subKey As String
OpenKey:
  subKey = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
  If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", 0, KEY_SET_VALUE, hKey) <> 0 Then
    Select Case MsgBox("Error: Could not write to Windows registry!", vbRetryCancel Or vbCritical, "Reboot and Shutdown - ERROR!")
      Case vbRetry
        GoTo OpenKey
      Case Else
        Exit Sub
    End Select
  End If
  Dim p As String
  p = """" & GetAppPath() & """ " & CMD_SHUTDOWN
  Call RegSetValueEx(hKey, "Reboot and Shutdown", 0, REG_SZ, p & vbNullChar, Len(p))
  Call RegCloseKey(hKey)
End Sub
