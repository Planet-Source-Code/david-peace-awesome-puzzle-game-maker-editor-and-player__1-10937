Attribute VB_Name = "INIs"
'Not thouroughly commented, comments desribe what each function does.
'Please see Form1 code to see how to call each function

Option Explicit
'APIs to access INI files and retrieve data
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

Function GetKeyVal(ByVal FileName As String, ByVal Section As String, ByVal Key As String)
'Returns info from an INI file
Dim RetVal As String, Worked As Integer
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
RetVal = String$(255, 0)
Worked = GetPrivateProfileString(Section, Key, "", RetVal, Len(RetVal), FileName)
If Worked = 0 Then
    GetKeyVal = ""
Else
    GetKeyVal = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
End If
End Function

Function AddToINI(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal KeyValue As String) As Integer
'Add info to an INI file
'Function returns 1 if successful and 0 if unsuccessful
WritePrivateProfileString Section, Key, KeyValue, FileName
AddToINI = 1
End Function

Function DeleteSection(ByVal FileName As String, ByVal Section As String) As Integer
'Delete an entire section and all it's keys from a given INI file
'Function returns 1 if successful and 0 if unsuccessful
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
If Not SectionExists(FileName, Section) Then MsgBox "Section, " & Section & ", Not Found. ~(DeleteSection)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Section Not Found.": Exit Function
WritePrivateProfileString Section, vbNullString, vbNullString, FileName
DeleteSection = 1
End Function

Function DeleteKey(ByVal FileName As String, ByVal Section As String, ByVal Key As String) As Integer
'Delete a key from an INI file
'Function returns 1 if successful and 0 if unsuccessful
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
If Not SectionExists(FileName, Section) Then MsgBox "Section, " & Section & ", Not Found. ~(DeleteKey)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Section Not Found.": Exit Function
If Not KeyExists(FileName, Section, Key) Then MsgBox "Key, " & Key & ", Not Found. ~(DeleteKey)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Key Not Found.": Exit Function
WritePrivateProfileString Section, Key, vbNullString, FileName
DeleteKey = 1
End Function

Function DeleteKeyValue(ByVal FileName As String, ByVal Section As String, ByVal Key As String) As Integer
'Delete a key's value from an INI file
'Function returns 1 if successful and 0 if unsuccessful
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
If Not SectionExists(FileName, Section) Then MsgBox "Section, " & Section & ", Not Found. ~(DeleteKeyValue)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Section Not Found.": Exit Function
If Not KeyExists(FileName, Section, Key) Then MsgBox "Key, " & Key & ", Not Found. ~(DeleteKeyValue)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Key Not Found.": Exit Function
WritePrivateProfileString Section, Key, "", FileName
DeleteKeyValue = 1
End Function

Public Function TotalSections(ByVal FileName As String) As Long
'Returns the total number of sections in a given INI file
Dim Counter As Integer
Dim InputData As String
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
Open FileName For Input As #1

Do While Not EOF(1)
    Line Input #1, InputData
    If IsSection(InputData) Then Counter = Counter + 1
Loop
Close #1
TotalSections = Counter
End Function

Public Function TotalKeys(ByVal FileName As String) As Long
'Returns the total number of keys in a given INI file
Dim Counter As Integer
Dim InputData As String
Dim Looper As Integer
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
Open FileName For Input As #2

Do While Not EOF(2)
    Line Input #2, InputData
    If IsKey(InputData) Then Counter = Counter + 1
Loop
Close #2
TotalKeys = Counter
End Function

Public Function NumKeys(ByVal FileName As String, ByVal Section As String) As Integer
'Returns the total number of keys in 1 given section.
Dim Counter As Integer
Dim InputData As String
Dim Looper As Integer
Dim InZone As Boolean
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
If Not SectionExists(FileName, Section) Then MsgBox "Section, " & Section & ", Not Found. ~(NumKeys)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Section Not Found.": Exit Function
InZone = False
Open FileName For Input As #3

Do While Not EOF(3)
    Line Input #3, InputData
    If InZone Then
        If IsSection(InputData) Or EOF(3) Then
            If EOF(3) Then
                NumKeys = Counter + 1
                Exit Do
            Else
                NumKeys = Counter
                Exit Do
            End If
        Else
            If IsKey(InputData) Then Counter = Counter + 1
        End If
    Else
        If InputData = "[" & Section & "]" Then
            InZone = True
        End If
    End If
Loop
Close #3
End Function

Public Function RenameSection(ByVal FileName As String, ByVal SectionName As String, ByVal NewSectionName As String) As Integer
'Renames a section in a given INI file.
'Function returns 1 if successful and 0 if unsuccessful
Dim TopKeys As String
Dim BotKeys As String
Dim Looper As Integer
Dim InputData As String
Dim InZone As Boolean
Dim Key1 As String, Key2 As String
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
If Not SectionExists(FileName, SectionName) Then MsgBox "Section, " & SectionName & ", Not Found. ~(RenameSection)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Section Not Found.": RenameSection = 0: Exit Function
If SectionExists(FileName, NewSectionName) Then MsgBox NewSectionName & " allready exists.  ~(RenameSection)", vbInformation, "Duplicate Section": RenameSection = 0: Exit Function
Open FileName For Input As #4

Do While Not EOF(4)
    Line Input #4, InputData
    If InZone Then
        If BotKeys = "" Then BotKeys = InputData Else BotKeys = BotKeys & vbCrLf & InputData
        If EOF(4) Then
            Close #4
            Kill FileName
            Open FileName For Append As #5
            If TopKeys <> "" Then Print #5, TopKeys
            Print #5, "[" & NewSectionName & "]" & vbCrLf & BotKeys
            Close #5
            RenameSection = 1
            Exit Function
        End If
    Else
        If InputData = "[" & SectionName & "]" Then
            InZone = True
        Else
            If TopKeys = "" Then TopKeys = InputData Else TopKeys = TopKeys & vbCrLf & InputData
        End If
    End If
Loop
Close #4
End Function

Public Function RenameKey(ByVal FileName As String, ByVal Section As String, ByVal KeyName As String, ByVal NewKeyName As String) As Integer
'Renames a key in a given INI file
'Function returns 1 if successful and 0 if unsuccessful
Dim KeyVal As String
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": RenameKey = 0: Exit Function
If Not SectionExists(FileName, Section) Then MsgBox "Section, " & Section & ", Not Found. ~(RenameKey)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Section Not Found.": RenameKey = 0: Exit Function
If Not KeyExists(FileName, Section, KeyName) Then MsgBox "Key, " & KeyName & ", Not Found. ~(RenameKey)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Key Not Found.": RenameKey = 0: Exit Function
If KeyExists(FileName, Section, NewKeyName) Then MsgBox NewKeyName & " allready exists in the section, " & Section & ".  ~(RenameKey)", vbInformation, "Duplicate Key.": RenameKey = 0: Exit Function
KeyVal = GetKeyVal(FileName, Section, KeyName)
DeleteKey FileName, Section, KeyName
AddToINI FileName, Section, NewKeyName, KeyVal
RenameKey = 1
End Function

Public Function GetKey(ByVal FileName As String, ByVal Section As String, ByVal KeyIndexNum As Integer) As String
'This function returns the name of a key which is identified by it's IndexNumber.
'The Section is identified as Text - GetKey2 identifies Section by it's IndexNumber
'IndexNumbers begin at 0 and increment up
Dim Counter As Integer
Dim InputData As String
Dim InZone As Boolean
Dim Looper As Integer
Dim KeyName As String
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
If Not SectionExists(FileName, Section) Then MsgBox "Section, " & Section & ", Not Found. ~(GetKey)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Section Not Found.": Exit Function
If NumKeys(FileName, Section) - 1 < KeyIndexNum Then MsgBox KeyIndexNum & ", not a valid Key Index Number. ~(GetKey)", vbInformation, "Invalid Index Number.": Exit Function

Counter = -1
Open FileName For Input As #6
Do While Not EOF(6)
    Line Input #6, InputData
    If InZone Then
        If IsKey(InputData) Then
            Counter = Counter + 1
            If Counter = KeyIndexNum Then
                For Looper = 1 To Len(InputData)
                    If Mid(InputData, Looper, 1) = "=" Then
                        GetKey = KeyName
                        Exit Do
                    Else
                        KeyName = KeyName & Mid(InputData, Looper, 1)
                    End If
                Next Looper
            End If
        End If
    Else
        If InputData = "[" & Section & "]" Then InZone = True
    End If
Loop
Close #6
End Function

Public Function GetKey2(ByVal FileName As String, ByVal SectionIndexNum As Integer, ByVal KeyIndexNum As Integer) As String
'This function returns the name of a key which is identified by it's IndexNumber.
'The Section is identified by it's IndexNumber
'IndexNumbers begin at 0 and increment up
Dim Counter As Integer
Dim Counter2 As Integer
Dim InputData As String
Dim InZone As Boolean
Dim Looper As Integer
Dim KeyName As String
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
If TotalSections(FileName) - 1 < SectionIndexNum Then MsgBox SectionIndexNum & ", not a valid Section Index Number. ~(GetKey2)", vbInformation, "Invalid Index Number.": Exit Function
If NumKeys(FileName, GetSection(FileName, SectionIndexNum)) - 1 < KeyIndexNum Then MsgBox KeyIndexNum & ", not a valid Key Index Number. ~(GetKey2)", vbInformation, "Invalid Index Number.": Exit Function
Counter = -1
Counter2 = -1
Open FileName For Input As #7
Do While Not EOF(7)
    Line Input #7, InputData
    If InZone Then
        If IsKey(InputData) Then
            Counter = Counter + 1
            If Counter = KeyIndexNum Then
                For Looper = 1 To Len(InputData)
                    If Mid(InputData, Looper, 1) = "=" Then
                        GetKey2 = KeyName
                        Exit Do
                    Else
                        KeyName = KeyName & Mid(InputData, Looper, 1)
                    End If
                Next Looper
            End If
        End If
    Else
        If IsSection(InputData) Then Counter2 = Counter2 + 1
        If Counter2 = SectionIndexNum Then InZone = True
    End If
Loop
Close #7
End Function

Public Function GetSection(ByVal FileName As String, ByVal SectionIndexNum As Integer) As String
'Returns a section name which is identified by it's indexnumber
'IndexNumbers begin at 0 and increment up
Dim InputData As String
Dim Counter As Integer
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
If TotalSections(FileName) - 1 < SectionIndexNum Then MsgBox SectionIndexNum & ", not a valid Section Index Number. ~(GetSection)", vbInformation, "Invalid Index Number.": Exit Function
Counter = -1
Open FileName For Input As #8
Do While Not EOF(8)
    Line Input #8, InputData
    If IsSection(InputData) Then
        Counter = Counter + 1
        InputData = Right(InputData, Len(InputData) - 1)
        InputData = Left(InputData, Len(InputData) - 1)
        If Counter = SectionIndexNum Then GetSection = InputData: Exit Do
    End If
Loop
Close #8
End Function

Public Function IsKey(ByVal TextLine As String) As Boolean
'This function determines whether or not a line of text is a valid Key (ex. "This=key")
'This returns True or False
Dim Looper As Integer
For Looper = 1 To Len(TextLine)
    If Mid(TextLine, Looper, 1) = "=" Then IsKey = True: Exit Function
Next Looper
IsKey = False
End Function

Public Function IsSection(ByVal TextLine As String) As Boolean
'This function determines whether or not a line of text is a
'valid section (ex. "[section]")
'This return's True or False
Dim FirstChar As String, LastChar As String
If TextLine = "" Then Exit Function
FirstChar = Mid(TextLine, 1, 1)
LastChar = Mid(TextLine, Len(TextLine), 1)
If FirstChar = "[" And LastChar = "]" Then IsSection = True Else IsSection = False
End Function

Public Function KeyExists(ByVal FileName As String, ByVal Section As String, ByVal Key As String) As Boolean
'This function determines if a key exists in a given section
'The Section is identified as Text - KeyExists2 identifies Section by its IndexNumber
'This returns True or False
Dim InZone As Boolean
Dim InputData As String
Dim Looper As Integer
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
If Not SectionExists(FileName, Section) Then MsgBox "Section, " & Section & ", Not Found. ~(KeyExists)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Section Not Found.": Exit Function
Open FileName For Input As #9
Do While Not EOF(9)
    Line Input #9, InputData
    If InZone Then
        If IsKey(InputData) Then
            If Left(InputData, Len(Key)) = Key Then
                KeyExists = True
                Exit Do
            End If
        ElseIf IsSection(InputData) Then
            KeyExists = False
            Exit Do
        End If
    Else
        If InputData = "[" & Section & "]" Then InZone = True
    End If
Loop
Close #9
End Function

Public Function KeyExists2(ByVal FileName As String, ByVal SectionIndexNum As Integer, ByVal Key As String) As Boolean
'This function determines if a key exists in a given section
'The Section is identified by its IndexNumber
'IndexNumbers begin at 0 and increment up
'This returns True or False
Dim InZone As Boolean
Dim InputData As String
Dim Looper As Integer
Dim Counter As Integer
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
If TotalSections(FileName) - 1 < SectionIndexNum Then MsgBox SectionIndexNum & ", not a valid Section Index Number. ~(KeyExists2)", vbInformation, "Invalid Index Number.": Exit Function
Counter = -1
Open FileName For Input As #10
Do While Not EOF(10)
    Line Input #10, InputData
    If InZone Then
        If IsKey(InputData) Then
            If Left(InputData, Len(Key)) = Key Then
                KeyExists2 = True
                Exit Do
            End If
        ElseIf IsSection(InputData) Then
            KeyExists2 = False
            Exit Do
        End If
    Else
        If IsSection(InputData) Then Counter = Counter + 1
        If Counter = SectionIndexNum Then InZone = True
    End If
Loop
Close #10
End Function

Public Function SectionExists(ByVal FileName As String, ByVal Section As String)
'This determines if a section exists in a given INI file
'This returns True or False
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
Dim InputData As String
Open FileName For Input As #11
Do While Not EOF(11)
    Line Input #11, InputData
    If "[" & Section & "]" = InputData Then SectionExists = True: Exit Do
    SectionExists = False
Loop
Close #11
End Function

Public Function GetSectionIndex(ByVal FileName As String, ByVal Section As String) As Integer
'This function is used to get the IndexNumber for a given Section
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
If Not SectionExists(FileName, Section) Then MsgBox "Section, " & Section & ", Not Found. ~(GetSectionIndex)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Section Not Found.": Exit Function
Dim InputData As String
Dim Counter As Integer
Counter = -1
Open FileName For Input As #12
Do While Not EOF(12)
    Line Input #12, InputData
    If IsSection(InputData) Then Counter = Counter + 1
    If "[" & Section & "]" = InputData Then GetSectionIndex = Counter
Loop
Close #12
End Function

Public Function GetKeyIndex(ByVal FileName As String, ByVal Section As String, ByVal Key As String) As Integer
'This function returns the IndexNumber of a key in a given Section
'The Section is identified as Text - GetKeyIndex2, Section is
'identified by it's IndexNumber
'IndexNumbers start at 0 and increment up
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
If Not SectionExists(FileName, Section) Then MsgBox "Section, " & Section & ", Not Found. ~(GetKeyIndex)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Section Not Found.": Exit Function
If Not KeyExists(FileName, Section, Key) Then MsgBox "Key, " & Key & ", Not Found. ~(GetKetIndex)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Key Not Found.": Exit Function
Dim InputData As String
Dim InZone As Boolean
Dim Counter As Integer
Counter = -1
Open FileName For Input As #13
Do While Not EOF(13)
    Line Input #13, InputData
    If InZone Then
        If IsKey(InputData) Then
            Counter = Counter + 1
            If Left(InputData, Len(Key)) = Key Then
                GetKeyIndex = Counter
                Exit Do
            End If
        ElseIf IsSection(InputData) Then
            Exit Do
        End If
    Else
        If "[" & Section & "]" = InputData Then InZone = True
    End If
Loop
Close #13
End Function

Public Function GetKeyIndex2(ByVal FileName As String, ByVal SectionIndexNum As Integer, ByVal Key As String) As Integer
'This function returns the IndexNumber of a key in a given Section
'The Section is identified by it's IndexNumber
'IndexNumbers start at 0 and increment up
If Dir(FileName) = "" Then MsgBox FileName & " not found.", vbCritical, "File Not Found": Exit Function
If TotalSections(FileName) - 1 < SectionIndexNum Then MsgBox SectionIndexNum & ", not a valid Section Index Number. ~(GetKeyIndex2)", vbInformation, "Invalid Index Number.": Exit Function
If Not KeyExists(FileName, GetSection(FileName, SectionIndexNum), Key) Then MsgBox "Key, " & Key & ", Not Found. ~(GetKetIndex2)" & vbCrLf & "Verify spelling and capitilization is correct.  Case-sensative.", vbInformation, "Key Not Found.": Exit Function
Dim InputData As String
Dim Counter As Integer
Dim Counter2 As Integer
Dim InZone As Boolean
Counter = -1
Counter2 = -1
Open FileName For Input As #14
Do While Not EOF(14)
    Line Input #14, InputData
    If InZone Then
        If IsKey(InputData) Then
            Counter = Counter + 1
            If Left(InputData, Len(Key)) = Key Then
                GetKeyIndex2 = Counter
                Exit Do
            End If
        ElseIf IsSection(InputData) Then
            Exit Do
        End If
    Else
        If IsSection(InputData) Then Counter2 = Counter2 + 1
        If Counter2 = SectionIndexNum Then InZone = True
    End If
Loop
Close #14
End Function
