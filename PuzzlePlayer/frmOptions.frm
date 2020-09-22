VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox chkBG 
      Caption         =   "Background Music"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   780
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   315
      Left            =   2700
      TabIndex        =   6
      Top             =   1200
      Width           =   915
   End
   Begin VB.CheckBox chkEffects 
      Caption         =   "Sound Effects"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   780
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin VB.HScrollBar scrSpeed 
      Height          =   195
      LargeChange     =   5
      Left            =   1500
      Max             =   10
      Min             =   1
      TabIndex        =   1
      Top             =   240
      Value           =   6
      Width           =   1875
   End
   Begin VB.Label lblSpeed 
      Caption         =   "6"
      Height          =   195
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fast"
      Height          =   195
      Index           =   1
      Left            =   3060
      TabIndex        =   3
      Top             =   480
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Slow"
      Height          =   195
      Index           =   0
      Left            =   1500
      TabIndex        =   2
      Top             =   480
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Animation Speed:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
scrSpeed.Value = IntNum
chkEffects.Value = SoundE
chkBG.Value = SoundB
If Dir(GetDir(frmMain.CD1.FileName) & "bg.mid") = "" Then chkBG.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.tmrAnim.Interval = (11 - scrSpeed.Value) * 10
IntNum = scrSpeed.Value
If scrSpeed.Value = 10 Then frmMain.tmrAnim.Interval = 1
If scrSpeed.Value = 1 Then frmMain.tmrAnim.Interval = 200
If chkBG Then
    If SoundB = 0 Then frmMid.MM1.Command = "Play"
Else
    frmMid.MM1.Command = "Stop"
End If
SoundE = chkEffects.Value
SoundB = chkBG.Value
End Sub

Private Sub scrSpeed_Change()
lblSpeed.Caption = scrSpeed.Value
End Sub

Private Sub scrSpeed_Scroll()
lblSpeed.Caption = scrSpeed.Value
End Sub
