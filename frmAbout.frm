VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Neophyte Word"
   ClientHeight    =   5295
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3654.703
   ScaleMode       =   0  'User
   ScaleWidth      =   6437.199
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrEgg 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6360
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   1560
      ScaleHeight     =   795
      ScaleWidth      =   5115
      TabIndex        =   6
      Top             =   3000
      Width           =   5175
      Begin VB.Label Label3 
         Caption         =   "Product ID:  52364-666-0000025-00034"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label lblUsername 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdTechSupport 
      Caption         =   "&Tech Support..."
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5400
      TabIndex        =   0
      Top             =   4080
      Width           =   1305
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   5400
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblEasterEggs 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label lblEgg 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
      Height          =   615
      Left            =   1560
      TabIndex        =   12
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Height          =   145
      Left            =   2066
      MouseIcon       =   "frmAbout.frx":00AE
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4920
      Width           =   692
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   """Failure is not an option, it comes bundled with the software."""
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "This product is licensed to:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright© Neophyte 2003. All rights reserved."
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Neophyte™ Word 2003"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   3660
      Left            =   240
      Picture         =   "frmAbout.frx":0200
      Top             =   120
      Width           =   915
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   6310.427
      Y1              =   2733.262
      Y2              =   2733.262
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   6324.513
      Y1              =   2743.615
      Y2              =   2743.615
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1FE7
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1185
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   4950
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Sub cmdSysInfo_Click()
On Error GoTo SysInfoErr
Dim rc As Long
Dim SysInfoPath As String
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
            If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            Else
                GoTo SysInfoErr
            End If
    Else
        GoTo SysInfoErr
    End If
Shell SysInfoPath, vbNormalFocus
Exit Sub
SysInfoErr:
MsgBox "System Information Is Unavailable At This Time", vbOKOnly + vbInformation, "Error"
End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub cmdTechSupport_Click()
MsgBox "For technical support email me (click the ""click here"" part of the copyright warning).", vbOKOnly + vbInformation, "Support"
End Sub

Private Sub Form_Load()
    If (EEgg1 + EEgg2 + EEgg3) > 0 Then
        lblEasterEggs.Visible = True
        lblEasterEggs.Caption = "Easter Eggs Found: " & EEgg1 + EEgg2 + EEgg3
    End If
    If EEgg1 = 1 Then lblUsername.Caption = "Bobby The Rubber Chicken" Else lblUsername.Caption = GetUsersName
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
Dim i           As Long
Dim rc          As Long
Dim hKey        As Long
Dim hDepth      As Long
Dim KeyValType  As Long
Dim tmpVal      As String
Dim KeyValSize  As Long
rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
tmpVal = String$(1024, 0)
KeyValSize = 1024
rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
        tmpVal = Left(tmpVal, KeyValSize - 1)
    Else
        tmpVal = Left(tmpVal, KeyValSize)
    End If
    Select Case KeyValType
        Case REG_SZ
            KeyVal = tmpVal
        Case REG_DWORD
                For i = Len(tmpVal) To 1 Step -1
                    KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
                Next
            KeyVal = Format$("&h" + KeyVal)
    End Select
GetKeyValue = True
rc = RegCloseKey(hKey)
Exit Function
GetKeyError:
KeyVal = ""
GetKeyValue = False
rc = RegCloseKey(hKey)
End Function

Private Sub lblEgg_DblClick()
tmrEgg.Enabled = True
End Sub

Private Sub lblEmail_Click()
ShellExecute GetDesktopWindow, "open", "iexplore", "mailto:some_phuker@hotmail.com?subject=Neophyte " _
             & "Word&body=Dear Neophyte,%0A%0D[Type your message here...]", "C:\", vbNormalFocus
End Sub

Private Sub tmrEgg_Timer()
Dim Position As POINTAPI
GetCursorPos Position
    If Position.x = 0 And Position.y = 0 Then
        EEgg2 = 1
        MsgBox "Rachael Binnie is the prettiest girl in the world." & vbCrLf & "I luv you baby.", vbOKOnly + vbInformation, "Easter Egg!"
    End If
tmrEgg.Enabled = False
lblEasterEggs.Visible = True
lblEasterEggs.Caption = "Easter Eggs Found: " & EEgg1 + EEgg2 + EEgg3
End Sub
