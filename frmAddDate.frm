VERSION 5.00
Begin VB.Form frmAddDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Date/Time"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdDone 
      Cancel          =   -1  'True
      Caption         =   "&Done"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      ItemData        =   "frmAddDate.frx":0000
      Left            =   240
      List            =   "frmAddDate.frx":001F
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   1320
      X2              =   3840
      Y1              =   220
      Y2              =   220
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1340
      X2              =   3860
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Select a format:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
    With frmMain.ActiveForm.rtfText
        Select Case List1.Text
            Case "1/2/03"
                .SelRTF = Format(Date, "d/m/yy")
            Case "01/02/03"
                .SelRTF = Format(Date, "dd/mm/yy")
            Case "1st Febuary 2003"
                .SelRTF = Format(Date, "d") & GetSuffix & " " & Format(Date, "mmmm yyyy")
            Case "Monday"
                .SelRTF = Format(Date, "dddd")
            Case "Monday 1st"
                .SelRTF = Format(Date, "dddd ") & Format(Date, "d") & GetSuffix
            Case "Monday 1st Febuary"
                .SelRTF = Format(Date, "dddd ") & Format(Date, "d") & GetSuffix & " " & Format(Date, "mmmm")
            Case "Monday 1st Febuary 2003"
                .SelRTF = Format(Date, "dddd ") & Format(Date, "d") & GetSuffix & " " & Format(Date, "mmmm yyyy")
            Case "1:30"
                .SelRTF = IIf(Hour(Time) > 12, Hour(Time) - 12, Hour(Time)) & ":" & Format(Minute(Time), "00")
            Case "13:30"
                .SelRTF = Format(Time, "hh:mm")
        End Select
    End With
End Sub

Private Sub Form_Load()
MakeTop Me
End Sub

Private Function GetSuffix() As String
Dim Suffix As String
    Select Case Day(Date)
        Case "11", "12", "13"
            Suffix = "th"
        Case Else
                Select Case Right$(Day(Date), 1)
                    Case "1"
                        Suffix = "st"
                    Case "2"
                        Suffix = "nd"
                    Case "3"
                        Suffix = "rd"
                    Case Else
                        Suffix = "th"
                End Select
    End Select
GetSuffix = Suffix
End Function

Private Sub List1_DblClick()
cmdOk_Click
End Sub
