Attribute VB_Name = "MOD"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCINVERT = &H660046

Type cLevel
    Map(0 To 16) As String
End Type

Type cGame
    Lev(1 To 50) As cLevel
    sFileName As String
End Type

Global CurGame As cGame
Global TotLevs As Integer
Global CurLev As Integer
Global Looper As Integer, Looper2 As Integer
Global CurChar As String
Global LevName As String
Global Changed As Boolean
Global KnowLoc As Boolean

Sub Main()
frmMain.Show
End Sub

Function LoadGame(LevelFile As String)
CurLev = 0
CurChar = "0"
CurGame.sFileName = LevelFile
For Looper = 1 To 50
    For Looper2 = 1 To 15
        CurGame.Lev(Looper).Map(Looper2) = "00000000000000000000"
    Next Looper2
Next Looper
TotLevs = GetKeyVal(LevelFile, "GAME", "Levels")
For Looper = 1 To TotLevs
    For Looper2 = 1 To 15
        CurGame.Lev(Looper).Map(Looper2) = GetKeyVal(LevelFile, Looper, Looper2)
    Next Looper2
Next Looper
If Dir(GetDir(LevelFile) & "pieces.bmp") <> "" Then
    frmMain.picPieces.Picture = LoadPicture(GetDir(LevelFile) & "pieces.bmp")
Else
    On Error GoTo ERR1
    frmMain.CD2.DialogTitle = "Please locate ""pieces.bmp"""
    frmMain.CD2.Filter = "pieces.bmp|pieces.bmp"
    frmMain.CD2.ShowOpen
    frmMain.picPieces.Picture = LoadPicture(frmMain.CD2.FileName)
End If
frmMain.picP.Picture = frmMain.picPieces.Picture
ERR1:
If Dir(GetDir(LevelFile) & "anim.bmp") <> "" Then
    frmMain.picAnim.Picture = LoadPicture(GetDir(LevelFile) & "anim.bmp")
Else
    On Error GoTo errhandler
    frmMain.CD2.DialogTitle = "Please locate ""anim.bmp"""
    frmMain.CD2.Filter = "anim.bmp|anim.bmp"
    frmMain.CD2.ShowOpen
    frmMain.picAnim.Picture = LoadPicture(frmMain.CD2.FileName)
End If
errhandler:
BitBlt frmMain.picS.hDC, 0, 0, 32, 32, frmMain.picAnim.hDC, 1, 100, SRCCOPY
frmMain.picTMP.BackColor = RGB(Val(GetKeyVal(CurGame.sFileName, "GAME", "BG.R")), Val(GetKeyVal(CurGame.sFileName, "GAME", "BG.G")), Val(GetKeyVal(CurGame.sFileName, "GAME", "BG.B")))
frmMain.picColor.BackColor = RGB(Val(GetKeyVal(CurGame.sFileName, "GAME", "BG.R")), Val(GetKeyVal(CurGame.sFileName, "GAME", "BG.G")), Val(GetKeyVal(CurGame.sFileName, "GAME", "BG.B")))
BitBlt frmMain.picPiece.hDC, 0, 0, 32, 32, frmMain.picColor.hDC, 0, 0, SRCCOPY
frmMain.picPiece.Refresh
frmMain.picS.Refresh
DrawLevel 1
End Function

Public Function DrawLevel(LevelNum As Integer)
Dim CurC As String * 1
CurLev = LevelNum
frmMain.Caption = LCase(LevName) & " - Level " & LevelNum & " of " & TotLevs & "  [PuzzleMaker - Editor]"
frmMain.picTMP.Cls
frmMain.picTMP.BackColor = frmMain.picColor.BackColor
For Looper = 1 To 15
    For Looper2 = 1 To 20
        CurC = Mid(CurGame.Lev(LevelNum).Map(Looper), Looper2, 1)
        BitBlt frmMain.picTMP.hDC, (Looper2 - 1) * 32, (Looper - 1) * 32, 32, 32, frmMain.picPieces.hDC, (Val(CurC) - 1) * 32, 0, SRCCOPY
        If CurC = "S" Then BitBlt frmMain.picTMP.hDC, (Looper2 - 1) * 32, (Looper - 1) * 32, 32, 32, frmMain.picS.hDC, 0, 0, SRCCOPY
    Next Looper2
Next Looper
frmMain.picTMP.Refresh
BitBlt frmMain.picPlay.hDC, 0, 0, 640, 480, frmMain.picTMP.hDC, 0, 0, SRCCOPY
frmMain.picPlay.Refresh
End Function

Function GetTitle(sFileName As String) As String
For Looper = Len(sFileName) To 1 Step -1
    If Mid(sFileName, Looper, 1) = "\" Then Exit For
Next Looper
GetTitle = Right(sFileName, Len(sFileName) - Looper)
End Function

Function GetDir(sFileName As String) As String
For Looper = Len(sFileName) To 1 Step -1
    If Mid(sFileName, Looper, 1) = "\" Then Exit For
Next Looper
GetDir = Left(sFileName, Looper)
End Function
