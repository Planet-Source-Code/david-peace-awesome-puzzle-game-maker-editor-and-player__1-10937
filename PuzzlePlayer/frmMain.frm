VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "untitled game.wcf - Level 1 of 1  [PuzzleMaker - Player]"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9615
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   501
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picASprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   10125
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   6
      Top             =   8175
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.PictureBox picAMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   10050
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   5
      Top             =   8250
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.PictureBox picAnim 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   9975
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   4
      Top             =   8325
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.PictureBox picPieces 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   9975
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   356
      TabIndex        =   3
      Top             =   7275
      Visible         =   0   'False
      Width           =   5340
   End
   Begin VB.PictureBox TMP 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   9975
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   9600
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   7200
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   16933
            MinWidth        =   16933
            Text            =   "Backspace = Repeat;   F3 = Previous Level;   F4 = Next Level;   O = Options;   A = About;   Esc = Quit;"
            TextSave        =   "Backspace = Repeat;   F3 = Previous Level;   F4 = Next Level;   O = Options;   A = About;   Esc = Quit;"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPlay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      Begin VB.Timer tmrSHOWBMP 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   75
         Top             =   1125
      End
      Begin VB.Timer tmrAnim 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   75
         Top             =   675
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   75
         Top             =   75
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Counter As Integer, MOVING As Boolean
Public Working As Boolean, iDir As Integer

Private Sub Form_Load()
If Dir(App.Path & "\options.ini") = "" Then
    SoundE = 1
    SoundB = 1
    IntNum = 6
Else
    SoundE = Val(GetKeyVal(App.Path & "\options.ini", "OPTIONS", "SoundE"))
    SoundB = Val(GetKeyVal(App.Path & "\options.ini", "OPTIONS", "SoundB"))
    IntNum = Val(GetKeyVal(App.Path & "\options.ini", "OPTIONS", "Anim"))
End If
Me.Show
DoEvents
On Error GoTo errhandler
If Right(Command$, 4) = ".WCF" And Dir(Command$) <> "" Then
    frmMain.CD1.FileName = Command$
    LevName = GetTitle(Command$)
    frmLoad.Show vbModal
    Exit Sub
Else
    frmMain.CD1.DialogTitle = "Open WCF Puzzle File"
    frmMain.CD1.Filter = "Puzzle File (*.wcf)|*.wcf|All Files (*.*)|*.*"
    frmMain.CD1.ShowOpen
    LevName = frmMain.CD1.FileTitle
    frmLoad.Show vbModal
    Exit Sub
errhandler:
    If Err.Number = 32755 Then End
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Ans As String
Ans = MsgBox("Are you sure you wish to quit?", vbYesNo + vbQuestion, "Quit?")
If Ans = vbNo Then
    Cancel = True
Else
    frmMid.MM1.Command = "Stop"
    frmMid.MM1.Command = "Close"
    WriteOpt
    End
End If
End Sub

Function WriteOpt()
Dim TXT As String
If Dir(App.Path & "\options.ini") <> "" Then Kill App.Path & "\options.ini"
Open App.Path & "\options.ini" For Append As #1
TXT = "[OPTIONS]" & vbCrLf & "Anim=" & IntNum & vbCrLf & "SoundE=" & SoundE & vbCrLf & "SoundB=" & SoundB
Print #1, TXT
Close #1
End Function

Private Sub picPlay_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Working And Not MOVING Then
Working = True
Select Case KeyCode
    Case vbKeyUp:
        If Mid(CurMap(CurPosY - 1), CurPosX, 1) = "1" Or Mid(CurMap(CurPosY - 1), CurPosX, 1) = "6" Or Mid(CurMap(CurPosY - 1), CurPosX, 1) = "7" Or Mid(CurMap(CurPosY - 1), CurPosX, 1) = "8" Then
            BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
            BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picAMask.hdc, 1, 67, SRCAND
            BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picASprite.hdc, 1, 67, SRCINVERT
            picPlay.Refresh
        End If
        If Mid(CurMap(CurPosY - 1), CurPosX, 1) = "3" Then
            If Mid(CurMap(CurPosY - 2), CurPosX, 1) = "2" Or Mid(CurMap(CurPosY - 2), CurPosX, 1) = "S" Then
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 3) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY - 3) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 2) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY - 2) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                CurMap(CurPosY - 1) = Left(CurMap(CurPosY - 1), CurPosX - 1) & "2" & Right(CurMap(CurPosY - 1), 20 - CurPosX)
                CurMap(CurPosY - 2) = Left(CurMap(CurPosY - 2), CurPosX - 1) & "3" & Right(CurMap(CurPosY - 2), 20 - CurPosX)
                picPlay.Refresh
                If Dir(sDir & "push.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "push.wav", 0, 0
            ElseIf Mid(CurMap(CurPosY - 2), CurPosX, 1) = "5" Then
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 3) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY - 3) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 2) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY - 2) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                CurMap(CurPosY - 1) = Left(CurMap(CurPosY - 1), CurPosX - 1) & "2" & Right(CurMap(CurPosY - 1), 20 - CurPosX)
                CurMap(CurPosY - 2) = Left(CurMap(CurPosY - 2), CurPosX - 1) & "4" & Right(CurMap(CurPosY - 2), 20 - CurPosX)
                TotCovered = TotCovered + 1
                picPlay.Refresh
                If Dir(sDir & "target.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "target.wav", 0, 0
            Else
                BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picAMask.hdc, 1, 67, SRCAND
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picASprite.hdc, 1, 67, SRCINVERT
                picPlay.Refresh
                If Dir(sDir & "error.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "error.wav", 0, 0
            End If
        ElseIf Mid(CurMap(CurPosY - 1), CurPosX, 1) = "4" Then
            If Mid(CurMap(CurPosY - 2), CurPosX, 1) = "2" Or Mid(CurMap(CurPosY - 2), CurPosX, 1) = "S" Then
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 3) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY - 3) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 2) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY - 2) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                CurMap(CurPosY - 1) = Left(CurMap(CurPosY - 1), CurPosX - 1) & "5" & Right(CurMap(CurPosY - 1), 20 - CurPosX)
                CurMap(CurPosY - 2) = Left(CurMap(CurPosY - 2), CurPosX - 1) & "3" & Right(CurMap(CurPosY - 2), 20 - CurPosX)
                TotCovered = TotCovered - 1
                picPlay.Refresh
                If Dir(sDir & "push.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "push.wav", 0, 0
            ElseIf Mid(CurMap(CurPosY - 2), CurPosX, 1) = "5" Then
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 3) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY - 3) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 2) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY - 2) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                CurMap(CurPosY - 1) = Left(CurMap(CurPosY - 1), CurPosX - 1) & "5" & Right(CurMap(CurPosY - 1), 20 - CurPosX)
                CurMap(CurPosY - 2) = Left(CurMap(CurPosY - 2), CurPosX - 1) & "4" & Right(CurMap(CurPosY - 2), 20 - CurPosX)
                picPlay.Refresh
                If Dir(sDir & "target.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "target.wav", 0, 0
            Else
                BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picAMask.hdc, 1, 67, SRCAND
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picASprite.hdc, 1, 67, SRCINVERT
                picPlay.Refresh
                If Dir(sDir & "error.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "error.wav", 0, 0
            End If
        End If
        picPlay.Refresh
        If Mid(CurMap(CurPosY - 1), CurPosX, 1) = "2" Or Mid(CurMap(CurPosY - 1), CurPosX, 1) = "S" Or Mid(CurMap(CurPosY - 1), CurPosX, 1) = "5" Or Mid(CurMap(CurPosY - 1), CurPosX, 1) = "9" Then Animate 1
    Case vbKeyRight:
        If Mid(CurMap(CurPosY), CurPosX + 1, 1) = "1" Or Mid(CurMap(CurPosY), CurPosX + 1, 1) = "6" Or Mid(CurMap(CurPosY), CurPosX + 1, 1) = "7" Or Mid(CurMap(CurPosY), CurPosX + 1, 1) = "8" Then
            BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
            BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picAMask.hdc, 1, 34, SRCAND
            BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picASprite.hdc, 1, 34, SRCINVERT
            picPlay.Refresh
        End If
    
        If Mid(CurMap(CurPosY), CurPosX + 1, 1) = "3" Then
            If Mid(CurMap(CurPosY), CurPosX + 2, 1) = "2" Or Mid(CurMap(CurPosY), CurPosX + 2, 1) = "S" Then
                BitBlt picPlay.hdc, (CurPosX + 1) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX + 1) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                CurMap(CurPosY) = Left(CurMap(CurPosY), CurPosX) & "23" & Right(CurMap(CurPosY), 18 - CurPosX)
                picPlay.Refresh
                If Dir(sDir & "push.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "push.wav", 0, 0
            ElseIf Mid(CurMap(CurPosY), CurPosX + 2, 1) = "5" Then
                BitBlt picPlay.hdc, (CurPosX + 1) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX + 1) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                CurMap(CurPosY) = Left(CurMap(CurPosY), CurPosX) & "24" & Right(CurMap(CurPosY), 18 - CurPosX)
                TotCovered = TotCovered + 1
                picPlay.Refresh
                If Dir(sDir & "target.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "target.wav", 0, 0
            Else
                BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picAMask.hdc, 1, 34, SRCAND
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picASprite.hdc, 1, 34, SRCINVERT
                picPlay.Refresh
                If Dir(sDir & "error.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "error.wav", 0, 0
            End If
        ElseIf Mid(CurMap(CurPosY), CurPosX + 1, 1) = "4" Then
            If Mid(CurMap(CurPosY), CurPosX + 2, 1) = "2" Or Mid(CurMap(CurPosY), CurPosX + 2, 1) = "S" Then
                BitBlt picPlay.hdc, (CurPosX + 1) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX + 1) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                CurMap(CurPosY) = Left(CurMap(CurPosY), CurPosX) & "53" & Right(CurMap(CurPosY), 18 - CurPosX)
                TotCovered = TotCovered - 1
                picPlay.Refresh
                If Dir(sDir & "push.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "push.wav", 0, 0
            ElseIf Mid(CurMap(CurPosY), CurPosX + 2, 1) = "5" Then
                BitBlt picPlay.hdc, (CurPosX + 1) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX + 1) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                CurMap(CurPosY) = Left(CurMap(CurPosY), CurPosX) & "54" & Right(CurMap(CurPosY), 18 - CurPosX)
                picPlay.Refresh
                If Dir(sDir & "target.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "target.wav", 0, 0
            Else
                BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picAMask.hdc, 1, 34, SRCAND
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picASprite.hdc, 1, 34, SRCINVERT
                picPlay.Refresh
                If Dir(sDir & "error.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "error.wav", 0, 0
            End If
        End If
        picPlay.Refresh
        If Mid(CurMap(CurPosY), CurPosX + 1, 1) = "2" Or Mid(CurMap(CurPosY), CurPosX + 1, 1) = "S" Or Mid(CurMap(CurPosY), CurPosX + 1, 1) = "5" Or Mid(CurMap(CurPosY), CurPosX + 1, 1) = "9" Then Animate 2
    Case vbKeyDown:
        If Mid(CurMap(CurPosY + 1), CurPosX, 1) = "1" Or Mid(CurMap(CurPosY + 1), CurPosX, 1) = "6" Or Mid(CurMap(CurPosY + 1), CurPosX, 1) = "7" Or Mid(CurMap(CurPosY + 1), CurPosX, 1) = "8" Then
            BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
            BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picAMask.hdc, 1, 100, SRCAND
            BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picASprite.hdc, 1, 100, SRCINVERT
            picPlay.Refresh
        End If
    
        If Mid(CurMap(CurPosY + 1), CurPosX, 1) = "3" Then
            If Mid(CurMap(CurPosY + 2), CurPosX, 1) = "2" Or Mid(CurMap(CurPosY + 2), CurPosX, 1) = "S" Then
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY + 1) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY + 1) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                CurMap(CurPosY + 1) = Left(CurMap(CurPosY + 1), CurPosX - 1) & "2" & Right(CurMap(CurPosY + 1), 20 - CurPosX)
                CurMap(CurPosY + 2) = Left(CurMap(CurPosY + 2), CurPosX - 1) & "3" & Right(CurMap(CurPosY + 2), 20 - CurPosX)
                picPlay.Refresh
                If Dir(sDir & "push.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "push.wav", 0, 0
            ElseIf Mid(CurMap(CurPosY + 2), CurPosX, 1) = "5" Then
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY + 1) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY + 1) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                CurMap(CurPosY + 1) = Left(CurMap(CurPosY + 1), CurPosX - 1) & "2" & Right(CurMap(CurPosY + 1), 20 - CurPosX)
                CurMap(CurPosY + 2) = Left(CurMap(CurPosY + 2), CurPosX - 1) & "4" & Right(CurMap(CurPosY + 2), 20 - CurPosX)
                TotCovered = TotCovered + 1
                picPlay.Refresh
                If Dir(sDir & "target.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "target.wav", 0, 0
            Else
                BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picAMask.hdc, 1, 100, SRCAND
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picASprite.hdc, 1, 100, SRCINVERT
                picPlay.Refresh
                If Dir(sDir & "error.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "error.wav", 0, 0
            End If
        ElseIf Mid(CurMap(CurPosY + 1), CurPosX, 1) = "4" Then
            If Mid(CurMap(CurPosY + 2), CurPosX, 1) = "2" Or Mid(CurMap(CurPosY + 2), CurPosX, 1) = "S" Then
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY + 1) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY + 1) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                CurMap(CurPosY + 1) = Left(CurMap(CurPosY + 1), CurPosX - 1) & "5" & Right(CurMap(CurPosY + 1), 20 - CurPosX)
                CurMap(CurPosY + 2) = Left(CurMap(CurPosY + 2), CurPosX - 1) & "3" & Right(CurMap(CurPosY + 2), 20 - CurPosX)
                TotCovered = TotCovered - 1
                picPlay.Refresh
                If Dir(sDir & "push.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "push.wav", 0, 0
            ElseIf Mid(CurMap(CurPosY + 2), CurPosX, 1) = "5" Then
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY + 1) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY + 1) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 1) * 32, (CurPosY) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                CurMap(CurPosY + 1) = Left(CurMap(CurPosY + 1), CurPosX - 1) & "5" & Right(CurMap(CurPosY + 1), 20 - CurPosX)
                CurMap(CurPosY + 2) = Left(CurMap(CurPosY + 2), CurPosX - 1) & "4" & Right(CurMap(CurPosY + 2), 20 - CurPosX)
                picPlay.Refresh
                If Dir(sDir & "target.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "target.wav", 0, 0
            Else
                BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picAMask.hdc, 1, 100, SRCAND
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picASprite.hdc, 1, 100, SRCINVERT
                picPlay.Refresh
                If Dir(sDir & "error.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "error.wav", 0, 0
            End If
        End If
        picPlay.Refresh
        If Mid(CurMap(CurPosY + 1), CurPosX, 1) = "2" Or Mid(CurMap(CurPosY + 1), CurPosX, 1) = "S" Or Mid(CurMap(CurPosY + 1), CurPosX, 1) = "5" Or Mid(CurMap(CurPosY + 1), CurPosX, 1) = "9" Then Animate 3
    Case vbKeyLeft:
        If Mid(CurMap(CurPosY), CurPosX - 1, 1) = "1" Or Mid(CurMap(CurPosY), CurPosX - 1, 1) = "6" Or Mid(CurMap(CurPosY), CurPosX - 1, 1) = "7" Or Mid(CurMap(CurPosY), CurPosX - 1, 1) = "8" Then
            BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
            BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picAMask.hdc, 1, 1, SRCAND
            BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picASprite.hdc, 1, 1, SRCINVERT
            picPlay.Refresh
        End If
    
        If Mid(CurMap(CurPosY), CurPosX - 1, 1) = "3" Then
            If Mid(CurMap(CurPosY), CurPosX - 2, 1) = "2" Or Mid(CurMap(CurPosY), CurPosX - 2, 1) = "S" Then
                BitBlt picPlay.hdc, (CurPosX - 3) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 3) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 2) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 2) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                CurMap(CurPosY) = Left(CurMap(CurPosY), CurPosX - 3) & "32" & Right(CurMap(CurPosY), 21 - CurPosX)
                picPlay.Refresh
                If Dir(sDir & "push.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "push.wav", 0, 0
            ElseIf Mid(CurMap(CurPosY), CurPosX - 2, 1) = "5" Then
                BitBlt picPlay.hdc, (CurPosX - 3) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 3) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 2) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 2) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 32, 0, SRCCOPY
                CurMap(CurPosY) = Left(CurMap(CurPosY), CurPosX - 3) & "42" & Right(CurMap(CurPosY), 21 - CurPosX)
                TotCovered = TotCovered + 1
                picPlay.Refresh
                If Dir(sDir & "target.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "target.wav", 0, 0
            Else
                BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picAMask.hdc, 1, 1, SRCAND
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picASprite.hdc, 1, 1, SRCINVERT
                picPlay.Refresh
                If Dir(sDir & "error.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "error.wav", 0, 0
            End If
        ElseIf Mid(CurMap(CurPosY), CurPosX - 1, 1) = "4" Then
            If Mid(CurMap(CurPosY), CurPosX - 2, 1) = "2" Or Mid(CurMap(CurPosY), CurPosX - 2, 1) = "S" Then
                BitBlt picPlay.hdc, (CurPosX - 3) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 3) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 64, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 2) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 2) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                CurMap(CurPosY) = Left(CurMap(CurPosY), CurPosX - 3) & "35" & Right(CurMap(CurPosY), 21 - CurPosX)
                TotCovered = TotCovered - 1
                picPlay.Refresh
                If Dir(sDir & "push.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "push.wav", 0, 0
            ElseIf Mid(CurMap(CurPosY), CurPosX - 2, 1) = "5" Then
                BitBlt picPlay.hdc, (CurPosX - 3) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 3) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 96, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 2) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                BitBlt TMP.hdc, (CurPosX - 2) * 32, (CurPosY - 1) * 32, 32, 32, picPieces.hdc, 128, 0, SRCCOPY
                CurMap(CurPosY) = Left(CurMap(CurPosY), CurPosX - 3) & "45" & Right(CurMap(CurPosY), 21 - CurPosX)
                picPlay.Refresh
                If Dir(sDir & "target.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "target.wav", 0, 0
            Else
                BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picAMask.hdc, 1, 1, SRCAND
                BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picASprite.hdc, 1, 1, SRCINVERT
                picPlay.Refresh
                If Dir(sDir & "error.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "error.wav", 0, 0
            End If
        End If
        picPlay.Refresh
        If Mid(CurMap(CurPosY), CurPosX - 1, 1) = "2" Or Mid(CurMap(CurPosY), CurPosX - 1, 1) = "S" Or Mid(CurMap(CurPosY), CurPosX - 1, 1) = "5" Or Mid(CurMap(CurPosY), CurPosX - 1, 1) = "9" Then Animate 4
    Case vbKeyF4:
        If CurLevel = TotLevels Then
            DrawLevel 1, CD1.FileName
        Else
            DrawLevel CurLevel + 1, CD1.FileName
        End If
    Case vbKeyF3:
        If CurLevel = 1 Then
            DrawLevel TotLevels, CD1.FileName
        Else
            DrawLevel CurLevel - 1, CD1.FileName
        End If
    Case vbKeyBack:
        DrawLevel CurLevel, CD1.FileName
    Case vbKeyEscape:
        Unload Me
    Case vbKeyO:
        frmOptions.Show vbModal
    Case vbKeyA:
        frmAbout.Show vbModal
End Select
DoEvents
Working = False
End If
End Sub

Function EndLevel()
Dim sDir As String
sDir = GetDir(CD1.FileName)
BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picAMask.hdc, 1, 100, SRCAND
BitBlt picPlay.hdc, (CurPosX - 1) * 32, (CurPosY - 1) * 32, 32, 32, picASprite.hdc, 1, 100, SRCINVERT
picPlay.Refresh
If Dir(sDir & "finish.wav") <> "" And frmOptions.chkEffects Then PlaySound GetDir(CD1.FileName) & "finish.wav", 0, 0
DoEvents
If CurLevel = TotLevels Then
    DrawLevel 1, CD1.FileName
Else
    DrawLevel CurLevel + 1, CD1.FileName
End If
End Function

Function Animate(iDirection As Integer)
If Not MOVING Then
    iDir = iDirection
    tmrAnim.Enabled = True
End If
End Function

Private Sub tmrAnim_Timer()
If Not MOVING Then
    MOVING = True
    Counter = Counter + 1
    If iDir = 1 Then
        BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
        BitBlt picPlay.hdc, (CurPosX - 1) * 32, ((CurPosY - 1) * 32) - (10 * (Counter - 1)), 32, 32, picAMask.hdc, ((Counter - 1) * 32) + (1 * Counter), 67, SRCAND
        BitBlt picPlay.hdc, (CurPosX - 1) * 32, ((CurPosY - 1) * 32) - (10 * (Counter - 1)), 32, 32, picASprite.hdc, ((Counter - 1) * 32) + (1 * Counter), 67, SRCINVERT
        picPlay.Refresh
        If Counter = 4 Then
            BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
            BitBlt picPlay.hdc, (CurPosX - 1) * 32, ((CurPosY - 2) * 32), 32, 32, picAMask.hdc, 1, 67, SRCAND
            BitBlt picPlay.hdc, (CurPosX - 1) * 32, ((CurPosY - 2) * 32), 32, 32, picASprite.hdc, 1, 67, SRCINVERT
            picPlay.Refresh
            Counter = 0
            tmrAnim.Enabled = False
            CurPosY = CurPosY - 1
        End If
    ElseIf iDir = 2 Then
        BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
        BitBlt picPlay.hdc, (CurPosX - 1) * 32 + (10 * (Counter - 1)), ((CurPosY - 1) * 32), 32, 32, picAMask.hdc, ((Counter - 1) * 32) + (1 * Counter), 34, SRCAND
        BitBlt picPlay.hdc, (CurPosX - 1) * 32 + (10 * (Counter - 1)), ((CurPosY - 1) * 32), 32, 32, picASprite.hdc, ((Counter - 1) * 32) + (1 * Counter), 34, SRCINVERT
        picPlay.Refresh
        If Counter = 4 Then
            BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
            BitBlt picPlay.hdc, (CurPosX) * 32, ((CurPosY - 1) * 32), 32, 32, picAMask.hdc, 1, 34, SRCAND
            BitBlt picPlay.hdc, (CurPosX) * 32, ((CurPosY - 1) * 32), 32, 32, picASprite.hdc, 1, 34, SRCINVERT
            picPlay.Refresh
            Counter = 0
            tmrAnim.Enabled = False
            CurPosX = CurPosX + 1
        End If
    ElseIf iDir = 3 Then
        BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
        BitBlt picPlay.hdc, (CurPosX - 1) * 32, ((CurPosY - 1) * 32) + (10 * (Counter - 1)), 32, 32, picAMask.hdc, ((Counter - 1) * 32) + (1 * Counter), 100, SRCAND
        BitBlt picPlay.hdc, (CurPosX - 1) * 32, ((CurPosY - 1) * 32) + (10 * (Counter - 1)), 32, 32, picASprite.hdc, ((Counter - 1) * 32) + (1 * Counter), 100, SRCINVERT
        picPlay.Refresh
        If Counter = 4 Then
            BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
            BitBlt picPlay.hdc, (CurPosX - 1) * 32, CurPosY * 32, 32, 32, picAMask.hdc, 1, 100, SRCAND
            BitBlt picPlay.hdc, (CurPosX - 1) * 32, CurPosY * 32, 32, 32, picASprite.hdc, 1, 100, SRCINVERT
            picPlay.Refresh
            Counter = 0
            tmrAnim.Enabled = False
            CurPosY = CurPosY + 1
        End If
    ElseIf iDir = 4 Then
        BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
        BitBlt picPlay.hdc, (CurPosX - 1) * 32 - (10 * (Counter - 1)), ((CurPosY - 1) * 32), 32, 32, picAMask.hdc, ((Counter - 1) * 32) + (1 * Counter), 1, SRCAND
        BitBlt picPlay.hdc, (CurPosX - 1) * 32 - (10 * (Counter - 1)), ((CurPosY - 1) * 32), 32, 32, picASprite.hdc, ((Counter - 1) * 32) + (1 * Counter), 1, SRCINVERT
        picPlay.Refresh
        If Counter = 4 Then
            BitBlt picPlay.hdc, 0, 0, 640, 480, TMP.hdc, 0, 0, SRCCOPY
            BitBlt picPlay.hdc, (CurPosX - 2) * 32, ((CurPosY - 1) * 32), 32, 32, picAMask.hdc, 1, 1, SRCAND
            BitBlt picPlay.hdc, (CurPosX - 2) * 32, ((CurPosY - 1) * 32), 32, 32, picASprite.hdc, 1, 1, SRCINVERT
            picPlay.Refresh
            Counter = 0
            tmrAnim.Enabled = False
            CurPosX = CurPosX - 1
        End If
    End If
    DoEvents
    If TotCovered = TotTargets Then EndLevel
    MOVING = False
End If
End Sub
