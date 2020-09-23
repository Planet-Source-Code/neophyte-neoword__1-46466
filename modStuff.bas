Attribute VB_Name = "modStuff"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Changed          As Boolean
Public EEgg1            As Integer
Public EEgg2            As Integer
Public EEgg3            As Integer

Sub Main()
On Error Resume Next
InitCommonControls
frmSplash.Show
End Sub

Function GetUsersName()
Dim Buffer      As String
Dim Size        As Long
Buffer = String(255, Chr(0))
Size = 255
GetUserName Buffer, Size
GetUsersName = Left(Buffer, Size)
End Function

Function MakeTop(Form As Form)
SetWindowPos Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Function

Function MakeNormal(Form As Form)
SetWindowPos Form.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Function
