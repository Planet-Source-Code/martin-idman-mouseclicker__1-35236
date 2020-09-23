VERSION 5.00
Begin VB.Form frmClicker 
   Caption         =   "Mouse Clicker"
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2475
   LinkTopic       =   "Form1"
   ScaleHeight     =   885
   ScaleWidth      =   2475
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   480
   End
   Begin VB.CommandButton btnTest 
      Caption         =   "Testbutton"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtDelay 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton btnRun 
      Caption         =   "&Start"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblClicker 
      Caption         =   "s"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblClicker 
      Caption         =   "Delay between clicks:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmClicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sub to handle run button click
Private Sub btnRun_Click()
    If btnRun.Caption = "&Start" Then
        ' If it's stopped, start it
        tmrTimer.Enabled = True
        btnRun.Caption = "&Stop"
        txtDelay.Enabled = False
    Else
        ' If it's started, stop it
        tmrTimer.Enabled = False
        btnRun.Caption = "&Start"
        txtDelay.Enabled = True
    End If
End Sub

' Sub to handle timer event
Private Sub tmrTimer_Timer()
    ' Emulate mouse click
    LeftClick
End Sub

' Sub to check the input in the delay textbox
Private Sub txtDelay_LostFocus()
    ' If it's less than 0.1 seconds, change it
    If Val(txtDelay.Text) < 0.1 Then
        MsgBox "Smallest value allowed is 0.1s", vbInformation
        txtDelay.Text = "0.1"
    End If
    ' Set timers interval to new value
    tmrTimer.Interval = Val(txtDelay.Text) * 1000
End Sub

' Sub that handles the testbuton click event
Private Sub btnTest_Click()
    ' Set buttons caption to random number
    ' just to show it's clicking the button
    btnTest.Caption = Int(Rnd() * 10000)
End Sub
