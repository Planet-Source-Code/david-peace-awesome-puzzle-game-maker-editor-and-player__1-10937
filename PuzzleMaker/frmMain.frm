VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "untitled game.wcf - Level 1 of 1  [PuzzleMaker - Editor]"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9600
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   2340
      TabIndex        =   7
      Top             =   7200
      Width           =   5655
      Begin VB.PictureBox picP 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   60
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   288
         TabIndex        =   10
         Top             =   140
         Width           =   4320
         Begin VB.Shape SHP2 
            BorderWidth     =   2
            DrawMode        =   6  'Mask Pen Not
            Height          =   465
            Left            =   15
            Top             =   15
            Visible         =   0   'False
            Width           =   465
         End
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   5100
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   140
         Width           =   480
      End
      Begin VB.PictureBox picS 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   4500
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   8
         Top             =   140
         Width           =   480
      End
   End
   Begin VB.PictureBox picTMP 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   9720
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "&Previous Level"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   7380
      Width           =   1395
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next Level"
      Height          =   375
      Left            =   8100
      TabIndex        =   4
      Top             =   7380
      Width           =   1395
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1740
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   7320
      Width           =   480
   End
   Begin VB.PictureBox picAnim 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   4500
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   2
      Top             =   7920
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.PictureBox picPieces 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   60
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   288
      TabIndex        =   1
      Top             =   7920
      Visible         =   0   'False
      Width           =   4320
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
      Begin MSComDlg.CommonDialog CD2 
         Left            =   480
         Top             =   1980
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   60
         Top             =   60
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Shape shpFollow 
         DrawMode        =   6  'Mask Pen Not
         Height          =   480
         Left            =   480
         Top             =   540
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu fileNew 
         Caption         =   "&New Puzzle"
         Shortcut        =   ^N
      End
      Begin VB.Menu fileSave 
         Caption         =   "&Save "
         Shortcut        =   ^S
      End
      Begin VB.Menu fileSaveAs 
         Caption         =   "Save &As...         "
      End
      Begin VB.Menu fileOpen 
         Caption         =   "&Open File"
         Shortcut        =   ^O
      End
      Begin VB.Menu fsp1 
         Caption         =   "-"
      End
      Begin VB.Menu fileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu editRemove 
         Caption         =   "&Remove Level"
         Shortcut        =   ^R
      End
      Begin VB.Menu editAdd 
         Caption         =   "&Add Level (s)"
         Shortcut        =   ^A
      End
      Begin VB.Menu editsep 
         Caption         =   "-"
      End
      Begin VB.Menu editIPieces 
         Caption         =   "Import &pieces.bmp"
         Shortcut        =   ^P
      End
      Begin VB.Menu editIAnim 
         Caption         =   "&Import anim.bmp"
         Shortcut        =   ^I
      End
      Begin VB.Menu editBG 
         Caption         =   "&Change BG Color"
         Shortcut        =   ^B
      End
      Begin VB.Menu editsep2 
         Caption         =   "-"
      End
      Begin VB.Menu ediNext 
         Caption         =   "&Next Level"
         Shortcut        =   {F4}
      End
      Begin VB.Menu editPrev 
         Caption         =   "&Previous Level"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu viewAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Down As Boolean

Private Sub cmdNext_Click()
If CurLev = TotLevs Then
    DrawLevel 1
Else
    DrawLevel CurLev + 1
End If
End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHP2.Visible = False
shpFollow.Visible = False
End Sub

Private Sub cmdPrev_Click()
If CurLev = 1 Then
    DrawLevel TotLevs
Else
    DrawLevel CurLev - 1
End If
End Sub

Private Sub cmdPrev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHP2.Visible = False
shpFollow.Visible = False
End Sub

Private Sub ediNext_Click()
cmdNext_Click
End Sub

Private Sub editAdd_Click()
Dim Ans As String
Ans = InputBox("Please enter the number of levels to add.  Max: " & 50 - TotLevs, "Add Levels", "1")
If Not IsNumeric(Ans) Or Val(Ans) > 50 - TotLevs Or Val(Ans) <= 0 Then
    Beep
    Exit Sub
Else
    Changed = True
    frmMain.Caption = LCase(LevName) & "* - Level " & LevelNum & " of " & TotLevs & "  [PuzzleMaker - Editor]"
    For Looper = TotLevs + 1 To Val(Ans)
        For Looper2 = 1 To 15
            CurGame.Lev(Looper).Map(Looper2) = "00000000000000000000"
        Next Looper2
    Next Looper
    TotLevs = TotLevs + Val(Ans)
    DrawLevel CurLev
End If
End Sub

Private Sub editBG_Click()
picColor_DblClick
End Sub

Private Sub editIAnim_Click()
picS_DblClick
End Sub

Private Sub editIPieces_Click()
picP_DblClick
End Sub

Private Sub editPrev_Click()
cmdPrev_Click
End Sub

Private Sub editRemove_Click()
Dim Ans As String, STR As String
Ans = InputBox("Please enter the level number you wish remove (1-" & TotLevs & ")", "Remove Level", CurLev)
If Not IsNumeric(Ans) Or Val(Ans) > TotLevs Or Val(Ans) < 1 Then
    Beep
    Exit Sub
Else
    Changed = True
    frmMain.Caption = LCase(LevName) & "* - Level " & LevelNum & " of " & TotLevs & "  [PuzzleMaker - Editor]"
    STR = MsgBox("Are you sure you wish to delete Level " & Ans & "?", vbYesNo + vbQuestion, "Delete Level?")
    If STR = vbYes Then
        For Looper = Val(Ans) To TotLevs - 1
            For Looper2 = 1 To 15
                CurGame.Lev(Looper).Map(Looper2) = CurGame.Lev(Looper + 1).Map(Looper2)
            Next Looper2
        Next Looper
        TotLevs = TotLevs - 1
        DrawLevel CurLev
    End If
End If
End Sub

Private Sub fileExit_Click()
Unload Me
End Sub

Private Sub fileNew_Click()
Dim Ans As String
If Changed Then
    Ans = MsgBox("Save before creating new WCF?", vbYesNoCancel + vbQuestion, "Save?")
    If Ans = vbCancel Then Exit Sub
    If Ans = vbNo Then GoTo NextPart
    If Ans = vbYes Then fileSave_Click
End If
NextPart:
KnowLoc = False
Changed = False
TotLevs = 1
LevName = "Untitled Game.wcf"
MakeDefault
DoEvents
LoadGame App.Path & "\default"
Me.Caption = LCase(LevName) & "* - Level " & CurLev & " of " & TotLevs & "  [PuzzleMaker - Editor]"
End Sub

Function MakeDefault()
Dim STR As String
If Dir(App.Path & "\default") <> "" Then Kill App.Path & "\default"
Open App.Path & "\default" For Output As #1
STR = "[GAME]" & vbCrLf
STR = STR & "Levels=1" & vbCrLf
STR = STR & "BG.R=0" & vbCrLf
STR = STR & "BG.G=0" & vbCrLf
STR = STR & "BG.B=0" & vbCrLf & vbCrLf
STR = STR & "[1]" & vbCrLf
STR = STR & "1=00000000000000000000" & vbCrLf
STR = STR & "2=00000000000000000000" & vbCrLf
STR = STR & "3=00000000000000000000" & vbCrLf
STR = STR & "4=00000000000000000000" & vbCrLf
STR = STR & "5=00000000000000000000" & vbCrLf
STR = STR & "6=00000000000000000000" & vbCrLf
STR = STR & "7=00000000000000000000" & vbCrLf
STR = STR & "8=00000000000000000000" & vbCrLf
STR = STR & "9=00000000000000000000" & vbCrLf
STR = STR & "10=00000000000000000000" & vbCrLf
STR = STR & "11=00000000000000000000" & vbCrLf
STR = STR & "12=00000000000000000000" & vbCrLf
STR = STR & "13=00000000000000000000" & vbCrLf
STR = STR & "14=00000000000000000000" & vbCrLf
STR = STR & "15=00000000000000000000" & vbCrLf
Print #1, STR
Close #1
End Function

Private Sub fileOpen_Click()
On Error GoTo errhandler
CD1.DialogTitle = "Open Puzzle File (*.wcf)"
CD1.Filter = "Puzzle File (*.wcf)|*.wcf|All Files (*.*)|*.*"
CD1.ShowOpen
LevName = CD1.FileTitle
LoadGame CD1.FileName
errhandler:
End Sub

Private Sub fileSave_Click()
Dim Ans As String
On Error GoTo errhandler
If KnowLoc = False Or Dir(CurGame.sFileName) = "" Then
    CD1.DialogTitle = "Save WCF Puzzle File..."
    CD1.Filter = "Puzzle File (*.wcf)|*.wcf|All Files (*.*)|*.*"
    CD1.ShowSave
    If Dir(CD1.FileName) = "" Then
        LevName = CD1.FileTitle
        SaveLevel CD1.FileName
        KnowLoc = True
    Else
        Ans = MsgBox("  Save over existing file?              ", vbYesNo + vbQuestion, "Replace?")
        If Ans = vbYes Then LevName = CD1.FileTitle: SaveLevel CD1.FileName: KnowLoc = True
    End If
errhandler:
Else
    SaveLevel CurGame.sFileName
End If
End Sub

Function SaveLevel(LevStr As String)
If Dir(LevStr) <> "" Then Kill LevStr
frmMain.Caption = LCase(LevName) & " - Level " & CurLev & " of " & TotLevs & "  [PuzzleMaker - Editor]"
Changed = False
AddToINI LevStr, "GAME", "Levels", TotLevs
AddToINI LevStr, "GAME", "BG.R", TakeRGB(1, picColor.BackColor)
AddToINI LevStr, "GAME", "BG.G", TakeRGB(2, picColor.BackColor)
AddToINI LevStr, "GAME", "BG.B", TakeRGB(3, picColor.BackColor)
For Looper = 1 To TotLevs
    For Looper2 = 1 To 15
        AddToINI LevStr, Looper, Looper2, CurGame.Lev(Looper).Map(Looper2)
    Next Looper2
Next Looper
DrawLevel CurLev
End Function

Function TakeRGB(IColor As Integer, Color As Long) As Integer
If IColor = 1 Then TakeRGB = Color Mod 256
If IColor = 2 Then TakeRGB = (Color \ 256) Mod 256
If IColor = 3 Then TakeRGB = Color \ 65536
End Function

Private Sub fileSaveAs_Click()
CD1.DialogTitle = "Save WCF Puzzle File As..."
CD1.Filter = "Puzzle File (*.wcf)|*.wcf|All Files (*.*)|*.*"
CD1.ShowSave
If Dir(CD1.FileName) = "" Then
    LevName = CD1.FileTitle
    SaveLevel CD1.FileName
    KnowLoc = True
Else
    Ans = MsgBox("  Save over existing file?              ", vbYesNo + vbQuestion, "Replace?")
    If Ans = vbYes Then LevName = CD1.FileTitle: SaveLevel CD1.FileName: KnowLoc = True
End If
errhandler:
End Sub

Private Sub Form_Load()
Me.Show
DoEvents
If UCase(Right(Command$, 4)) = ".WCF" And Dir(Command$) <> "" Then
    LevName = GetTitle(Command$)
    CD1.FileName = Command$
    LoadGame CD1.FileName
    KnowLoc = True
Else
    On Error GoTo errhandler
    LevName = "Untitled Game.wcf"
    CD1.DialogTitle = "Open Puzzle File (*.wcf)"
    CD1.Filter = "Puzzle File (*.wcf)|*.wcf|All Files (*.*)|*.*"
    CD1.ShowOpen
    LevName = CD1.FileTitle
    LoadGame CD1.FileName
    KnowLoc = True
    Exit Sub
errhandler:
    fileNew_Click
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHP2.Visible = False
shpFollow.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Ans As String
If Changed Then
    Ans = MsgBox("Do you wish to save before quitting?", vbYesNoCancel + vbQuestion, "Save?")
    If Ans = vbCancel Then Cancel = True
    If Ans = vbYes Then
        fileSave_Click
        If Dir(App.Path & "\default") <> "" Then Kill App.Path & "\default"
        End
    End If
    If Ans = vbNo Then
        If Dir(App.Path & "\default") <> "" Then Kill App.Path & "\default"
        End
    End If
Else
    Ans = MsgBox("Are you sure you wish to quit?", vbYesNo + vbQuestion, "Quit?")
    If Ans = vbNo Then
        Cancel = True
    Else
        If Dir(App.Path & "\default") <> "" Then Kill App.Path & "\default"
        End
    End If
End If
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHP2.Visible = False
shpFollow.Visible = False
End Sub

Private Sub picColor_Click()
picPiece.BackColor = picColor.BackColor
picPiece.Cls
CurChar = "0"
frmMain.Caption = LCase(LevName) & "* - Level " & CurLev & " of " & TotLevs & "  [PuzzleMaker - Editor]"
End Sub

Private Sub picColor_DblClick()
On Error GoTo errhandler
CD1.ShowColor
picColor.BackColor = CD1.Color
Changed = True
DrawLevel CurLev
errhandler:
End Sub

Private Sub picColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHP2.Visible = False
shpFollow.Visible = False
End Sub

Private Sub picP_Click()
BitBlt picPiece.hDC, 0, 0, 32, 32, picPieces.hDC, SHP2.Left - 1, 0, SRCCOPY
CurChar = SHP2.Left \ 32 + 1
picPiece.Refresh
End Sub

Private Sub picP_DblClick()
On Error GoTo errhandler
CD1.DialogTitle = "Import pieces.bmp"
CD1.Filter = "pieces.bmp|pieces.bmp"
CD1.ShowOpen
picP.Picture = LoadPicture(CD1.FileName)
picPieces.Picture = picP.Picture
BitBlt picPiece.hDC, 0, 0, 32, 32, picPieces.hDC, (Val(CurChar) - 1) * 32, 0, SRCCOPY
picPiece.Refresh
DrawLevel CurLev
errhandler:
End Sub

Private Sub picP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim CurX As Integer, CurY As Integer
CurX = X \ 32
CurY = Y \ 32
SHP2.Visible = True
SHP2.Move CurX * 32 + 1, CurY * 32 + 1
shpFollow.Visible = False
End Sub

Private Sub picPiece_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHP2.Visible = False
shpFollow.Visible = False
End Sub

Private Sub picPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim CurL As String
CurL = CurGame.Lev(CurLev).Map(shpFollow.Top / 32 + 1)
Down = True
Select Case Button
    Case 1:
        Changed = True
        frmMain.Caption = LCase(LevName) & "* - Level " & CurLev & " of " & TotLevs & "  [PuzzleMaker - Editor]"
        BitBlt picTMP.hDC, shpFollow.Left, shpFollow.Top, 32, 32, picPiece.hDC, 0, 0, SRCCOPY
        BitBlt picPlay.hDC, 0, 0, 640, 480, picTMP.hDC, 0, 0, SRCCOPY
        CurGame.Lev(CurLev).Map(shpFollow.Top \ 32 + 1) = Left(CurL, shpFollow.Left \ 32) & CurChar & Right(CurL, Len(CurL) - shpFollow.Left \ 32 - 1)
        picPlay.Refresh
    Case Else:
        BitBlt picPiece.hDC, 0, 0, 32, 32, picTMP.hDC, shpFollow.Left, shpFollow.Top, SRCCOPY
        CurChar = Mid(CurL, shpFollow.Left / 32 + 1, 1)
        picPiece.Refresh
End Select
End Sub

Private Sub picPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim CurX As Integer, CurY As Integer
Dim CurL As String
CurX = X \ 32
CurY = Y \ 32
SHP2.Visible = False
shpFollow.Visible = True
shpFollow.Move CurX * 32, CurY * 32
DoEvents
If CurGame.sFileName <> "" Then CurL = CurGame.Lev(CurLev).Map(shpFollow.Top / 32 + 1)
If Down Then
    Select Case Button
        Case 1:
            BitBlt picTMP.hDC, shpFollow.Left, shpFollow.Top, 32, 32, picPiece.hDC, 0, 0, SRCCOPY
            BitBlt picPlay.hDC, 0, 0, 640, 480, picTMP.hDC, 0, 0, SRCCOPY
            CurGame.Lev(CurLev).Map(shpFollow.Top \ 32 + 1) = Left(CurL, shpFollow.Left \ 32) & CurChar & Right(CurL, Len(CurL) - shpFollow.Left \ 32 - 1)
            picPlay.Refresh
        Case Else:
            BitBlt picPiece.hDC, 0, 0, 32, 32, picTMP.hDC, shpFollow.Left, shpFollow.Top, SRCCOPY
            CurChar = Mid(CurGame.Lev(CurLev).Map(shpFollow.Top / 32 + 1), shpFollow.Left / 32 + 1, 1)
            picPiece.Refresh
    End Select
End If
End Sub

Private Sub picPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Down = False
End Sub

Private Sub picS_Click()
BitBlt picPiece.hDC, 0, 0, 32, 32, picS.hDC, 0, 0, SRCCOPY
picPiece.Refresh
CurChar = "S"
End Sub

Private Sub picS_DblClick()
On Error GoTo errhandler
CD1.DialogTitle = "Import anim.bmp"
CD1.Filter = "anim.bmp|anim.bmp"
CD1.ShowOpen
picAnim.Picture = LoadPicture(CD1.FileName)
BitBlt picS.hDC, 0, 0, 32, 32, picAnim.hDC, 1, 100, SRCCOPY
picS.Refresh
BitBlt picPiece.hDC, 0, 0, 32, 32, picS.hDC, 0, 0, SRCCOPY
picPiece.Refresh
CurChar = "S"
DrawLevel CurLev
errhandler:
End Sub

Private Sub picS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHP2.Visible = False
shpFollow.Visible = False
End Sub

Private Sub viewAbout_Click()
frmAbout.Show vbModal
End Sub
