Attribute VB_Name = "modShutdown"
Option Explicit

Private Const EWX_SHUTDOWN As Long = &H1
Private Const EWX_REBOOT As Long = &H2
Private Const EWX_FORCE As Long = &H4
Private Const EWX_POWEROFF As Long = &H8
Private Const EWX_LOGOFF As Long = &H0

Private Const SHUTDOWN_INSTALL_UPDATES As Long = &H40
Private Const SHUTDOWN_RESTART As Long = &H4
Private Const SHUTDOWN_POWEROFF As Long = &H8

Private Const SHTDN_REASON_MAJOR_OPERATINGSYSTEM As Long = &H20000
Private Const SHTDN_REASON_MINOR_UPGRADE As Long = &H3

Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Private Const TOKEN_QUERY As Long = &H8
Private Const SE_PRIVILEGE_ENABLED As Long = &H2

Private Type LUID
  dwLowPart As Long
  dwHighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
  udtLUID As LUID
  dwAttributes As Long
End Type

Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  laa As LUID_AND_ATTRIBUTES
End Type

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Long) As Long

Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function InitiateShutdown Lib "advapi32.dll" Alias "InitiateShutdownW" (ByVal lpMachineName As Long, ByVal lpMessage As Long, ByVal dwGracePeriod As Long, ByVal dwShutdownFlags As Long, ByVal dwReason As Long) As Long

Public Sub TriggerShutdown(Optional DoReboot As Boolean = False)
  Call EnableShutdownPrivileges
  If IsWindowsVistaOrHigher() = True Then
    Call InitiateShutdown(0&, 0&, 0&, SHUTDOWN_INSTALL_UPDATES Or IIf(DoReboot, SHUTDOWN_RESTART, SHUTDOWN_POWEROFF), SHTDN_REASON_MAJOR_OPERATINGSYSTEM Or SHTDN_REASON_MINOR_UPGRADE)
  Else
    Call ExitWindowsEx(IIf(DoReboot, EWX_REBOOT, EWX_SHUTDOWN), 0&)
  End If
End Sub

Private Function EnableShutdownPrivileges() As Boolean
  Dim hProcessHandle As Long
  Dim hTokenHandle As Long
  Dim lpv_la As LUID
  Dim Token As TOKEN_PRIVILEGES
  hProcessHandle = GetCurrentProcess
  If Not hProcessHandle = 0 Then
    If Not OpenProcessToken(hProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hTokenHandle) = 0 Then
      If Not LookupPrivilegeValue(vbNullString, "SeShutdownPrivilege", lpv_la) = 0 Then
        With Token
          .PrivilegeCount = 1
          With .laa
            .udtLUID = lpv_la
            .dwAttributes = SE_PRIVILEGE_ENABLED
          End With
        End With
        If Not AdjustTokenPrivileges(hTokenHandle, False, Token, ByVal 0&, ByVal 0&, ByVal 0&) = 0 Then
          EnableShutdownPrivileges = True
        End If
      End If
    End If
  End If
End Function

