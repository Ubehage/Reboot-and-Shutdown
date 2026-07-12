Attribute VB_Name = "modMain"
Option Explicit

Private Const CMD_SHUTDOWN = "uyhgbrf"

Global Const COUNTDOWN_SECONDS As Integer = 5

Global Const COLOR_BACKGROUND As Long = 2105376
Global Const COLOR_CONTROLS As Long = 2763306
Global Const COLOR_BUTTON_HOVER As Long = 3684408
Global Const COLOR_BUTTON_PRESSED As Long = 3289650
Global Const COLOR_BACKGROUND_DISABLED As Long = 5263440

Global Const COLOR_TEXT As Long = 14737632
Global Const COLOR_TEXT_HOVER As Long = 15790320
Global Const COLOR_TEXT_DISABLED As Long = 7895160
Global Const COLOR_TEXT_ONGREEN As Long = 15463654
Global Const COLOR_TEXT_ONRED As Long = 15395579
Global Const COLOR_TEXT_DISABLED_ONGREEN As Long = 13355947
Global Const COLOR_TEXT_DISABLED_ONRED As Long = 10592542

Global Const COLOR_GREEN As Long = 5023791
Global Const COLOR_GREEN_HOVER As Long = 6339651
Global Const COLOR_GREEN_PRESSED As Long = 4033061
Global Const COLOR_GREEN_DISABLED As Long = 6455130
Global Const COLOR_YELLOW As Long = 4965861
Global Const COLOR_YELLOW_HOVER As Long = 6673645
Global Const COLOR_YELLOW_PRESSED As Long = 3710156
Global Const COLOR_YELLOW_DISABLED As Long = 7112080
Global Const COLOR_RED As Long = 4539862
Global Const COLOR_RED_HOVER As Long = 6513642
Global Const COLOR_RED_PRESSED As Long = 3223992
Global Const COLOR_RED_DISABLED As Long = 6776730
Global Const COLOR_OUTLINE As Long = 3815994
Global Const COLOR_OUTLINE_LIGHT As Long = 7368816

Sub Main()
  InitCommonControls
  Dim r As Boolean
  If LCase$(Command$) = CMD_SHUTDOWN Then r = False Else r = True
  If r = True Then
    RegisterRunOnce
    TriggerShutdown r
  Else
    Load frmCountdown
    frmCountdown.SetForm
    frmCountdown.SetCountdownTimer COUNTDOWN_SECONDS
  End If
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
