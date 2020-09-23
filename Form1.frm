VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "vCard Creator"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   600
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save As (ANSI)"
      Height          =   495
      Left            =   4080
      TabIndex        =   19
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Telephone Numbers"
      Height          =   1815
      Left            =   0
      TabIndex        =   22
      Top             =   1080
      Width           =   4935
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Work"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label Label7 
         Caption         =   "Home"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label Label8 
         Caption         =   "FAX Work"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   150
         TabIndex        =   40
         Top             =   1080
         Width           =   1605
      End
      Begin VB.Label Label9 
         Caption         =   "FAX Home"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label Label13 
         Caption         =   "Cell"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adresses"
      Height          =   2775
      Left            =   5160
      TabIndex        =   23
      Top             =   120
      Width           =   3495
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2760
         TabIndex        =   16
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Street"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label Label10 
         Caption         =   "eMail"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   2280
         Width           =   525
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "ZIP"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2400
         TabIndex        =   36
         Top             =   1800
         Width           =   285
      End
      Begin VB.Label Label15 
         Caption         =   "City"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label16 
         Caption         =   "Country"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2040
         Width           =   405
      End
      Begin VB.Label Label20 
         Caption         =   "Work"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   1320
         Width           =   2205
      End
      Begin VB.Label Label17 
         Caption         =   "Country"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "ZIP"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   31
         Top             =   720
         Width           =   285
      End
      Begin VB.Label Label14 
         Caption         =   "City"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label18 
         Caption         =   "Street"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "Home"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   240
         Width           =   2205
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   7800
      TabIndex        =   27
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send Using BlueTooth"
      Height          =   495
      Left            =   6960
      TabIndex        =   21
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save As  (UTF-8)"
      Height          =   495
      Left            =   5520
      TabIndex        =   20
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label12 
      Caption         =   "Note"
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   840
      Width           =   1605
   End
   Begin VB.Label Label11 
      Caption         =   "Job Title"
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   600
      Width           =   1605
   End
   Begin VB.Label Label3 
      Caption         =   "Last Name"
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   360
      Width           =   1605
   End
   Begin VB.Label Label2 
      Caption         =   "First Name"
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Type STARTUPINFO
         cb As Long
         lpReserved As String
         lpDesktop As String
         lpTitle As String
         dwX As Long
         dwY As Long
         dwXSize As Long
         dwYSize As Long
         dwXCountChars As Long
         dwYCountChars As Long
         dwFillAttribute As Long
         dwFlags As Long
         wShowWindow As Integer
         cbReserved2 As Integer
         lpReserved2 As Long
         hStdInput As Long
         hStdOutput As Long
         hStdError As Long
      End Type

      Private Type PROCESS_INFORMATION
         hProcess As Long
         hThread As Long
         dwProcessID As Long
         dwThreadID As Long
      End Type

      Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
         hHandle As Long, ByVal dwMilliseconds As Long) As Long

      Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
         lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
         lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
         ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
         ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
         lpStartupInfo As STARTUPINFO, lpProcessInformation As _
         PROCESS_INFORMATION) As Long

      Private Declare Function CloseHandle Lib "kernel32" (ByVal _
         hObject As Long) As Long

      Private Const NORMAL_PRIORITY_CLASS = &H20&
      Private Const INFINITE = -1&

Private Declare Function CreateProcess Lib "kernel32" Alias _
      "CreateProcessA" (ByVal lpApplicationName As String, ByVal _
      lpCommandLine As String, lpProcessAttributes As Any, _
      lpThreadAttributes As Any, ByVal bInheritHandles As Long, _
      ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal _
      lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, _
      lpProcessInformation As PROCESS_INFORMATION) As Long

Sub Base64Init()
 Dim l As Long
   
   ReDim Base64Reverse(255)
   
   Base64Lookup = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
   
   For l = 0 To 63
      Base64Reverse(Base64Lookup(l)) = l
   Next
   
End Sub


Function CheckEmpty() As Boolean
If Trim(Text1.Text) <> "" Or Trim(Text2.Text) <> "" Or Trim(Text3.Text) <> "" Or Trim(Text4.Text) <> "" Or Trim(Text5.Text) <> "" Or Trim(Text6.Text) <> "" Or Trim(Text7.Text) <> "" Or Trim(Text8.Text) <> "" Or Trim(Text9.Text) <> "" Or Trim(Text10.Text) <> "" Or Trim(Text11.Text) <> "" Or Trim(Text12.Text) <> "" Or Trim(Text14.Text) <> "" Or Trim(Text15.Text) <> "" Or Trim(Text16.Text) <> "" Or Trim(Text17.Text) <> "" Or Trim(Text18.Text) <> "" Or Trim(Text19.Text) <> "" Then
CheckEmpty = False
Else
CheckEmpty = True
End If


End Function

Function cProc(app As String, cmdLine As String)
Dim pInfo As PROCESS_INFORMATION
Dim sInfo As STARTUPINFO
Dim sNull As String

sInfo.cb = Len(sInfo)
success& = CreateProcess(sNull, app & " " & cmdLine, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, sNull, sInfo, pInfo)


   ' ProcessID& = Shell("Calc.exe", vbNormalFocus)
   ' ProcessHandle& = OpenProcess(SYNCHRONIZE, True, ProcessID&)

End Function

      Public Sub ExecCmd(cmdLine$)
         Dim proc As PROCESS_INFORMATION
         Dim start As STARTUPINFO

         ' Initialize the STARTUPINFO structure:
         start.cb = Len(start)
        
         ' Start the shelled application:
         Ret& = CreateProcessA(0&, cmdLine$, 0&, 0&, 1&, _
            NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

         ' Wait for the shelled application to finish:
         Ret& = WaitForSingleObject(proc.hProcess, INFINITE)
         Ret& = CloseHandle(proc.hProcess)
      End Sub

Private Sub Command1_Click()
Dim vCard  As String
If CheckEmpty = True Then
MsgBox "Please insert some data"
Exit Sub
End If
cd1.FileName = Text1.Text & " " & Text2.Text & ".vcf"
cd1.DefaultExt = "vcf"
cd1.DialogTitle = "Save file as..."
cd1.Filter = "vCard file(*vcf)|*.vcf|All Files(*.*|*.*"
cd1.FilterIndex = 0
On Error GoTo cderror1
cd1.CancelError = True
cd1.Flags = cdlOFNExtensionDifferent Or cdlOFNHideReadOnly + cdlOFNNoReadOnlyReturn Or cdlOFNPathMustExist
cd1.ShowSave
vCard = "BEGIN:VCARD"
vCard = vCard & vbCrLf & "VERSION:2.1"
vCard = vCard + vbCrLf & "N;CHARSET=UTF-8:" & RASCII59(Text2.Text) & ";" & RASCII59(Text1.Text)
vCard = vCard + vbCrLf & "FN;CHARSET=UTF-8:" & RASCII59(Text2.Text) & " " & RASCII59(Text1.Text)
vCard = vCard & vbCrLf & "TEL;VOICE;HOME:" & RASCII59(Text6)
vCard = vCard & vbCrLf & "TEL;VOICE;WORK:" & RASCII59(Text5)
vCard = vCard & vbCrLf & "TEL;CELL:" & RASCII59(Text12)
vCard = vCard & vbCrLf & "TEL;FAX;HOME:" & RASCII59(Text8)
vCard = vCard & vbCrLf & "TEL;FAX;WORK:" & RASCII59(Text7)
vCard = vCard & vbCrLf & "EMAIL:" & Text9
vCard = vCard & vbCrLf & "NOTE;CHARSET=UTF-8:" & RASCII59(Text11.Text)
vCard = vCard & vbCrLf & "ORG;CHARSET=UTF-8:" & RASCII59(Text10.Text)
vCard = vCard & vbCrLf & "ADR;HOME;CHARSET=UTF-8:;;" & RASCII59(Text3.Text) & ";" & RASCII59(Text14.Text) & ";;" & RASCII59(Text16.Text) & ";" & RASCII59(Text18.Text)
vCard = vCard & vbCrLf & "ADR;WORK;CHARSET=UTF-8:;;" & RASCII59(Text4.Text) & ";" & RASCII59(Text15.Text) & ";;" & RASCII59(Text17.Text) & ";" & RASCII59(Text19.Text)
vCard = vCard & vbCrLf & "END:VCARD"


Dim vCardUTF8 As String
vCardUTF8 = UTF8_Encode(vCard)
Open cd1.FileName For Output As #1
Print #1, vCardUTF8
Close #1
Exit Sub
cderror1:
If Err = cdlCancel Then
Exit Sub
End If
Resume Next




End Sub
Private Sub Command2_Click()
Dim vCard  As String
If CheckEmpty = True Then
MsgBox "Please insert some data"
Exit Sub
End If
vCard = "BEGIN:VCARD"
vCard = vCard & vbCrLf & "VERSION:2.1"
vCard = vCard + vbCrLf & "N;CHARSET=UTF-8:" & RASCII59(Text2.Text) & ";" & RASCII59(Text1.Text)
vCard = vCard + vbCrLf & "FN;CHARSET=UTF-8:" & RASCII59(Text2.Text) & " " & RASCII59(Text1.Text)
vCard = vCard & vbCrLf & "TEL;VOICE;HOME:" & RASCII59(Text6)
vCard = vCard & vbCrLf & "TEL;VOICE;WORK:" & RASCII59(Text5)
vCard = vCard & vbCrLf & "TEL;CELL:" & RASCII59(Text12)
vCard = vCard & vbCrLf & "TEL;FAX;HOME:" & RASCII59(Text8)
vCard = vCard & vbCrLf & "TEL;FAX;WORK:" & RASCII59(Text7)
vCard = vCard & vbCrLf & "EMAIL:" & Text9
vCard = vCard & vbCrLf & "NOTE;CHARSET=UTF-8:" & RASCII59(Text11.Text)
vCard = vCard & vbCrLf & "ORG;CHARSET=UTF-8:" & RASCII59(Text10.Text)
vCard = vCard & vbCrLf & "ADR;HOME;CHARSET=UTF-8:;;" & RASCII59(Text3.Text) & ";" & RASCII59(Text14.Text) & ";;" & RASCII59(Text16.Text) & ";" & RASCII59(Text18.Text)
vCard = vCard & vbCrLf & "ADR;WORK;CHARSET=UTF-8:;;" & RASCII59(Text4.Text) & ";" & RASCII59(Text15.Text) & ";;" & RASCII59(Text17.Text) & ";" & RASCII59(Text19.Text)
vCard = vCard & vbCrLf & "END:VCARD"


Dim vCardUTF8 As String
vCardUTF8 = UTF8_Encode(vCard)
fg$ = GetTmpName("vcf")
fg$ = Left$(fg$, Len(fg$) - 3)
fg$ = fg$ & "vcf"

Open fg$ For Output As #1
Print #1, vCardUTF8
Close #1


'Shell "C:\Program Files\LevelOne MDU-0005USB\btsendto_explorer.exe  c:\1.vcf", vbNormalFocus
 ExecCmd "C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe " & fg$
         
         Kill fg$
  'cProc "C:\Program Files\LevelOne MDU-0005USB\btsendto_explorer.exe", "c:\1.vcf"
  
End Sub


Private Sub Command3_Click()
Dim X() As Byte
Dim k As Long
Dim v As String
Open "c:\me.jpg" For Binary As #1
k = LOF(1)
ReDim X(k)
Get #1, , X
v = EncodeByteArray(X)
Open "c:\me2.jpg" For Binary As #2
Put #2, , v
Close #1, #2
Dim j As String
Dim f As String
Dim y() As Byte
Open "c:\me2.jpg" For Input Access Read As #1
'Do While Not EOF(1)
Input #1, j
f = f & j
'Loop

y = DecodeToByteArray(f)
Open "c:\me3.jpg" For Binary As #2
Put #2, , y
Close #1, #2

End Sub


Private Sub Command4_Click()
Dim vCard  As String
If CheckEmpty = True Then
MsgBox "Please insert some data"
Exit Sub
End If
cd1.FileName = Text1.Text & " " & Text2.Text & ".vcf"
cd1.DefaultExt = "vcf"
cd1.DialogTitle = "Save file as..."
cd1.Filter = "vCard file(*vcf)|*.vcf|All Files(*.*|*.*"
cd1.FilterIndex = 0
On Error GoTo cderror
cd1.CancelError = True
cd1.Flags = cdlOFNExtensionDifferent Or cdlOFNHideReadOnly + cdlOFNNoReadOnlyReturn Or cdlOFNPathMustExist
cd1.ShowSave
vCard = "BEGIN:VCARD"
vCard = vCard & vbCrLf & "VERSION:2.1"
vCard = vCard + vbCrLf & "N:" & RASCII59(Text2.Text) & ";" & RASCII59(Text1.Text)
vCard = vCard + vbCrLf & "FN:" & RASCII59(Text2.Text) & " " & RASCII59(Text1.Text)
vCard = vCard & vbCrLf & "TEL;VOICE;HOME:" & RASCII59(Text6)
vCard = vCard & vbCrLf & "TEL;VOICE;WORK:" & RASCII59(Text5)
vCard = vCard & vbCrLf & "TEL;CELL:" & RASCII59(Text12)
vCard = vCard & vbCrLf & "TEL;FAX;HOME:" & RASCII59(Text8)
vCard = vCard & vbCrLf & "TEL;FAX;WORK:" & RASCII59(Text7)
vCard = vCard & vbCrLf & "EMAIL:" & Text9
vCard = vCard & vbCrLf & "NOTE:" & RASCII59(Text11.Text)
vCard = vCard & vbCrLf & "ORG:" & RASCII59(Text10.Text)
vCard = vCard & vbCrLf & "ADR;HOME:;;" & RASCII59(Text3.Text) & ";" & RASCII59(Text14.Text) & ";;" & RASCII59(Text16.Text) & ";" & RASCII59(Text18.Text)
vCard = vCard & vbCrLf & "ADR;WORK:;;" & RASCII59(Text4.Text) & ";" & RASCII59(Text15.Text) & ";;" & RASCII59(Text17.Text) & ";" & RASCII59(Text19.Text)
vCard = vCard & vbCrLf & "END:VCARD"

Open cd1.FileName For Output As #1
Print #1, vCard
Close #1


Exit Sub
cderror:
If Err = cdlCancel Then
Exit Sub
End If
Resume Next
End Sub

Private Sub Form_Load()
Base64Init
'Open "c:\logo_blue.jpg" For Binary As #1
Dim v As Byte
'Do While Not EOF(1)
'Get #1, , v
'Text13.Text = Text13.Text & Chr$(v)
'Loop
Close
End Sub


