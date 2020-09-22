Attribute VB_Name = "MOD"
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCINVERT = &H660046

Global Map(0 To 16) As String
Global CurMap(0 To 16) As String
Global CurPosX As Integer
Global CurPosY As Integer
Global Looper As Integer, Looper2 As Integer
Global CurLevel As Integer
Global TotLevels As Integer
Global TotTargets As Integer
Global TotCovered As Integer
Global LevName As String

Global SoundE As Integer
Global SoundB As Integer
Global IntNum As Integer

Sub Main()
frmMain.Show
End Sub

Public Function DrawLevel(LevelNum As Integer, LevelFile As String)
Working = True
For Looper = 1 To 15
    Map(Looper) = GetKeyVal(LevelFile, LevelNum, Looper)
    CurMap(Looper) = Map(Looper)
Next Looper
TotTargets = 0
TotCovered = 0
For Looper = 1 To 15
    For Looper2 = 1 To 21
        If Mid(Map(Looper), Looper2, 1) = "5" Then TotTargets = TotTargets + 1
    Next Looper2
Next Looper
TotLevels = Val(GetKeyVal(LevelFile, "GAME", "Levels"))
CurLevel = LevelNum
frmMain.Caption = LCase(LevName) & " - Level " & LevelNum & " of " & TotLevels & "  [PuzzleMaker - Player]"
frmMain.picPlay.Cls
frmMain.TMP.Cls
frmMain.picPlay.BackColor = RGB(Val(GetKeyVal(LevelFile, "GAME", "BG.R")), Val(GetKeyVal(LevelFile, "GAME", "BG.G")), Val(GetKeyVal(LevelFile, "GAME", "BG.B")))
frmMain.TMP.BackColor = frmMain.picPlay.BackColor
frmMain.picPieces.Picture = LoadPicture(GetDir(LevelFile) & "pieces.bmp")
For Looper = 1 To 15
    For Looper2 = 1 To 20
        Select Case Mid(Map(Looper), Looper2, 1)
            Case 1 To 9:
                BitBlt frmMain.TMP.hdc, (Looper2 - 1) * 32, (Looper - 1) * 32, 32, 32, frmMain.picPieces.hdc, (Val(Mid(Map(Looper), Looper2, 1)) - 1) * 32, 0, SRCCOPY
        End Select
    Next Looper2
Next Looper
frmMain.TMP.Refresh
BitBlt frmMain.picPlay.hdc, 0, 0, 640, 480, frmMain.TMP.hdc, 0, 0, SRCCOPY
frmMain.picPlay.Refresh
For Looper = 1 To 15
    For Looper2 = 1 To 20
        Select Case Mid(Map(Looper), Looper2, 1)
            Case "S":
                BitBlt frmMain.picPlay.hdc, (Looper2 - 1) * 32, (Looper - 1) * 32, 32, 32, frmMain.picPieces.hdc, 32, 0, SRCCOPY
                BitBlt frmMain.TMP.hdc, (Looper2 - 1) * 32, (Looper - 1) * 32, 32, 32, frmMain.picPieces.hdc, 32, 0, SRCCOPY
                BitBlt frmMain.picPlay.hdc, (Looper2 - 1) * 32, (Looper - 1) * 32, 32, 32, frmMain.picAMask.hdc, 1, 100, SRCAND
                BitBlt frmMain.picPlay.hdc, (Looper2 - 1) * 32, (Looper - 1) * 32, 32, 32, frmMain.picASprite.hdc, 1, 100, SRCINVERT
                CurPosX = Looper2: CurPosY = Looper
        End Select
    Next Looper2
Next Looper
frmMain.picPlay.Refresh
frmMain.TMP.Refresh
frmMain.tmrAnim.Enabled = False
Working = False
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
