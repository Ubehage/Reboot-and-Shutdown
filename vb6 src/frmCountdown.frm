VERSION 5.00
Begin VB.Form frmCountdown 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3450
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7905
   ControlBox      =   0   'False
   Icon            =   "frmCountdown.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Reboot_Shutdown.Button cmdCancel 
      Height          =   315
      Left            =   5460
      TabIndex        =   1
      Top             =   2490
      Width           =   1365
      _ExtentX        =   1429
      _ExtentY        =   556
      Caption         =   "Abort countdown"
   End
   Begin VB.Label lblCountdown 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shutting down in %s seconds."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   450
      Left            =   735
      TabIndex        =   0
      Top             =   675
      Width           =   4440
   End
End
Attribute VB_Name = "frmCountdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COUNTDOWN_TEXT = "Shutting down in %s seconds."

Dim WithEvents CountdownTimer As RebootTimer
Attribute CountdownTimer.VB_VarHelpID = -1

Dim EndTime As Date
Dim CancelCountdown As Boolean

Friend Sub SetForm()
  WindowOnTop Me.hWnd, True
  lblCountdown.Visible = False
  lblCountdown.Caption = GetCountdownText(999)
  lblCountdown.Move (Screen.TwipsPerPixelX * 15), (Screen.TwipsPerPixelY * 15)
  Me.Width = ((Me.Width - Me.ScaleWidth) + (lblCountdown.Width + (lblCountdown.Left * 2)))
  Me.Height = ((Me.Height - Me.ScaleHeight) + ((lblCountdown.Height + (lblCountdown.Top * 2)) + cmdCancel.Height))
  lblCountdown.Caption = ""
  lblCountdown.Visible = True
  cmdCancel.Move (Me.ScaleWidth - (cmdCancel.Width + (Screen.TwipsPerPixelX * 5))), (Me.ScaleHeight - (cmdCancel.Height + (Screen.TwipsPerPixelY * 5)))
  Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
  Me.Show
End Sub

Friend Sub SetCountdownTimer(Optional CountdownSeconds As Integer = COUNTDOWN_SECONDS)
  KillCountdownTimer
  Set CountdownTimer = New RebootTimer
  CountdownTimer.Interval = 500
  EndTime = DateAdd("s", CLng(CountdownSeconds), Now)
  UpdateCountdownText
  CountdownTimer.Enabled = True
End Sub

Private Sub KillCountdownTimer()
  If CountdownTimer Is Nothing Then Exit Sub
  CountdownTimer.Enabled = False
  Set CountdownTimer = Nothing
End Sub

Private Function GetCountdownText(RemainingSeconds As Integer) As String
  GetCountdownText = Replace$(COUNTDOWN_TEXT, "%s", CStr(RemainingSeconds))
End Function

Private Sub UpdateCountdownText()
  lblCountdown.Caption = GetCountdownText(DateDiff("s", Now, EndTime))
End Sub

Private Sub cmdCancel_Click()
  CancelCountdown = True
End Sub

Private Sub CountdownTimer_Timer()
  CountdownTimer.Enabled = False
  If (Now >= EndTime Or CancelCountdown = True) Then
    Unload Me
    If CancelCountdown = False Then TriggerShutdown False
  Else
    UpdateCountdownText
  End If
  If Not CountdownTimer Is Nothing Then CountdownTimer.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  KillCountdownTimer
  WindowOnTop Me.hWnd, False
End Sub
