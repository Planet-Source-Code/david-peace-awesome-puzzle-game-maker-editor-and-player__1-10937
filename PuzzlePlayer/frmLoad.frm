VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading . . ."
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   ControlBox      =   0   'False
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PB 
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   525
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
      Max             =   266
   End
   Begin VB.Label LBL 
      AutoSize        =   -1  'True
      Caption         =   "Loading images into memory . . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   1
      Top             =   150
      Width           =   2730
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Show
DoEvents
MakeMS
End Sub

Function MakeMS()
If Dir(GetDir(frmMain.CD1.FileName) & "anim_mask.bmp") <> "" Then
    frmMain.picAMask.Picture = LoadPicture(GetDir(frmMain.CD1.FileName) & "anim_mask.bmp")
    GoTo DoSprite
End If
frmMain.picAnim.Picture = LoadPicture(GetDir(frmMain.CD1.FileName) & "anim.bmp")
LBL.Caption = "Generating necessary mask images . . ."
For Looper = 1 To 133
    For Looper2 = 1 To 133
        If GetPixel(frmMain.picAnim.hdc, Looper, Looper2) = RGB(255, 0, 255) Then
            SetPixel frmMain.picAMask.hdc, Looper, Looper2, vbWhite
        Else
            SetPixel frmMain.picAMask.hdc, Looper, Looper2, vbBlack
        End If
        PB.Value = Looper
    Next Looper2
Next Looper
frmMain.picAMask.Refresh
DoEvents
SavePicture frmMain.picAMask.Image, GetDir(frmMain.CD1.FileName) & "anim_mask.bmp"
'BitBlt frmMain.TMP.hdc, (Looper2 - 1) * 32, (Looper - 1) * 32, 32, 32, frmMain.picAnim.hdc, 1, 100, SRCCOPY
DoSprite:
PB.Value = 133
If Dir(GetDir(frmMain.CD1.FileName) & "anim_sprite.bmp") <> "" Then
    frmMain.picASprite.Picture = LoadPicture(GetDir(frmMain.CD1.FileName) & "anim_sprite.bmp")
    GoTo SkipSprite
End If
LBL.Caption = "Generating necessary sprite images . . ."
For Looper = 1 To 133
    For Looper2 = 1 To 133
        If GetPixel(frmMain.picAnim.hdc, Looper, Looper2) = RGB(255, 0, 255) Then
            SetPixel frmMain.picASprite.hdc, Looper, Looper2, vbBlack
        Else
            SetPixel frmMain.picASprite.hdc, Looper, Looper2, GetPixel(frmMain.picAnim.hdc, Looper, Looper2)
        End If
        PB.Value = Looper + 133
    Next Looper2
Next Looper
frmMain.picASprite.Refresh
DoEvents
SavePicture frmMain.picASprite.Image, GetDir(frmMain.CD1.FileName) & "anim_sprite.bmp"
SkipSprite:
PB.Value = 266
frmMain.picASprite.Refresh
LBL.Caption = "Loading bg.mid . . . please wait."
frmMid.Show
DrawLevel 1, frmMain.CD1.FileName
frmMain.Show
Unload Me
End Function
