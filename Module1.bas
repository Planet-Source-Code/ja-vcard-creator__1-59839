Attribute VB_Name = "Module1"
Private Const CP_UTF8 = 65001

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Function RASCII59(f As String) As String
For i% = 1 To Len(f)
    l% = Asc(Mid$(f, i%, 1))
    If l% = 59 Or l% = 58 Then
        Mid$(f, i%, 1) = Chr$(32)
    End If
Next
RASCII59 = f

        
End Function

Public Function GetTmpName(prefix) As String
Dim TempFileName As String * 256
    Dim X As Long
    Dim DriveName As String
    t$ = GetTmpPath
    X = GetTempFileName(t$, prefix, 0, TempFileName)
    GetTmpName = Left$(TempFileName, InStr(TempFileName, Chr(0)) - 1)


End Function

Public Function GetTmpPath() As String
Dim xd As String
xd = String$(255, 0)
Ret& = GetTempPath(Len(xd), xd)
GetTmpPath = Left$(xd, Ret&)

End Function
Public Function UTF8_Encode(ByVal Text As String) As String

Dim sBuffer As String
Dim lLength As Long

If Text <> "" Then
lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), -1, 0, 0, 0, 0)
sBuffer = Space$(lLength)
lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), -1, StrPtr(sBuffer), Len(sBuffer), 0, 0)
sBuffer = StrConv(sBuffer, vbUnicode)
UTF8_Encode = Left$(sBuffer, lLength - 1)
Else
UTF8_Encode = ""
End If

End Function


