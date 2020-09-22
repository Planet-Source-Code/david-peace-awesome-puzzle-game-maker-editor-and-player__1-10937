VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmMid 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   330
   LinkTopic       =   "Form1"
   ScaleHeight     =   390
   ScaleWidth      =   330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MCI.MMControl MM1 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      PauseEnabled    =   -1  'True
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "frmMid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move -5, -5, 2, 2
If Dir(GetDir(frmMain.CD1.FileName) & "bg.mid") <> "" Then
    MM1.Command = "Close"
    MM1.Notify = False
    MM1.Wait = True
    MM1.Shareable = False
    MM1.FileName = GetDir(frmMain.CD1.FileName) & "bg.mid"
    MM1.Command = "Open"
End If
If SoundB = 1 Then MM1.Command = "Play"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MM1.Command = "Close"
End Sub
