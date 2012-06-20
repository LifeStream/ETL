VERSION 5.00
Begin VB.Form frmSysTray 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Hard Drive"
   ClientHeight    =   675
   ClientLeft      =   1425
   ClientTop       =   2295
   ClientWidth     =   1680
   Icon            =   "frmSysTray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   45
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic 
      Height          =   555
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   60
      Width           =   495
   End
   Begin VB.Timer tmrMain 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "Balloon Tip"
         Index           =   0
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Exit"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents SysTray As clsSysTray
Attribute SysTray.VB_VarHelpID = -1

Private Sub Form_Load()
    Set SysTray = New clsSysTray
    Me.WindowState = vbMinimized
    DoEvents
    Me.Hide
    SysTray.Init Me, appName & " - " & appVersion
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SysTray = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    SysTray.MouseMove Button, X, Me
    Debug.Print Button, X
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    SysTray.MouseMove Button, X, Me
End Sub

Private Sub mnuPopup_Click(Index As Integer)
    Select Case Me.mnuPopup(Index).Caption
        Case "Balloon Tip"
          SysTray.ShowBalloonTip "Sample message of some length.", beInformation, "Balloon Tip"
          
        Case "Exit"
          subTerminator
          Unload Me
          'End   % Cannot Execute End, just 'Unload Me'. Otherwise the sysTray Icon will not remove asap from the taskbar.
          
        Case Else
          MsgBox Me.mnuPopup(Index).Caption
          
    End Select
End Sub

Private Sub SysTray_BalloonClicked()
    MsgBox "Balloon tip was clicked", vbInformation, "Notice"
End Sub

Private Sub SysTray_DoubleClick()
    SysTray.ShowBalloonTip "Double click tray icon", beInformation, "Balloon Tip"
End Sub

Private Sub SysTray_RightClick()
    PopupMenu Me.mnuSysTray
End Sub

Private Sub tmrMain_Timer()

  'Initialize time marker
  If stopWatch = 0 Then resetStopWatch
  
  'Reset time marker is after 12:00am
  If (Timer - stopWatch) < 0 Then resetStopWatch
  
  If (Timer - stopWatch) >= colSettings("queue") Then
    executeETL
    resetStopWatch
  End If
  
End Sub

Private Sub resetStopWatch()
  stopWatch = Timer
End Sub
