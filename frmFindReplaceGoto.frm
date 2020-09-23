VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmFindReplaceGoto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Tabsa 
      BorderStyle     =   0  'None
      Height          =   2100
      Index           =   0
      Left            =   200
      ScaleHeight     =   2100
      ScaleWidth      =   4530
      TabIndex        =   12
      Top             =   480
      Width           =   4530
      Begin VB.CheckBox chkCase 
         Caption         =   "Match Case"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtFind 
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   3120
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Find:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Tabsa 
      BorderStyle     =   0  'None
      Height          =   2100
      Index           =   1
      Left            =   200
      ScaleHeight     =   2100
      ScaleWidth      =   4530
      TabIndex        =   6
      Top             =   480
      Width           =   4530
      Begin VB.CheckBox chkCase2 
         Caption         =   "Match Case"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtReplace 
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "&Replace"
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtReplace2 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Find:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Replace With:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.PictureBox Tabsa 
      BorderStyle     =   0  'None
      Height          =   2100
      Index           =   2
      Left            =   200
      ScaleHeight     =   2100
      ScaleWidth      =   4530
      TabIndex        =   2
      Top             =   480
      Width           =   4530
      Begin VB.CheckBox chkCase3 
         Caption         =   "Match Case"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtGoto 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton cmdGoTo 
         Caption         =   "&Go To"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Go To:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4471
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Find"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Find"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Replace"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Replace"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Go To"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Go To"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmFindReplaceGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Position As Long

Private Sub cmdFind_Click()
    With frmMain.ActiveForm.rtfText
            If chkCase.Value = False Then
                    If InStr(1, .Text, txtFind.Text, vbTextCompare) <> 0 Then
                        .SelStart = InStr(1, .Text, txtFind.Text, vbTextCompare) - 1
                        .SelLength = Len(txtFind.Text)
                        Position = InStr(1, .Text, txtFind.Text, vbTextCompare)
                    Else
                        Display txtFind.Text
                    End If
            Else
                    If InStr(1, .Text, txtFind.Text) <> 0 Then
                        .SelStart = InStr(1, .Text, txtFind.Text) - 1
                        .SelLength = Len(txtFind.Text)
                        Position = InStr(1, .Text, txtFind.Text)
                    Else
                        Display txtFind.Text
                    End If
            End If
        .SetFocus
    End With
End Sub

Private Sub cmdGoTo_Click()
    With frmMain.ActiveForm.rtfText
            If chkCase3.Value = False Then
                    If InStr(1, .Text, txtGoto.Text, vbTextCompare) <> 0 Then
                        .SelStart = InStr(1, .Text, txtGoto.Text, vbTextCompare) - 1
                    Else
                        Display txtGoto.Text
                    End If
            Else
                    If InStr(1, .Text, txtGoto.Text) <> 0 Then
                        .SelStart = InStr(1, .Text, txtGoto.Text) - 1
                    Else
                        Display txtGoto.Text
                    End If
            End If
        .SetFocus
    End With
End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub cmdReplace_Click()
    With frmMain.ActiveForm.rtfText
            If chkCase2.Value = False Then
                    If InStr(1, .Text, txtReplace.Text, vbTextCompare) <> 0 Then
                        .Text = Replace(.Text, txtReplace.Text, txtReplace2.Text, , , vbTextCompare)
                    Else
                        Display txtReplace.Text
                    End If
            Else
                    If InStr(1, .Text, txtReplace.Text) <> 0 Then
                        .Text = Replace(.Text, txtReplace.Text, txtReplace2.Text)
                    Else
                        Display txtReplace.Text
                    End If
            End If
        .SetFocus
    End With
End Sub

Private Sub Form_Load()
Position = 1
MakeTop Me
End Sub

Private Sub TabStrip1_Click()
    Select Case TabStrip1.SelectedItem
        Case "Find"
            Tabsa(0).Visible = True
            Tabsa(1).Visible = False
            Tabsa(2).Visible = False
        Case "Replace"
            Tabsa(0).Visible = False
            Tabsa(1).Visible = True
            Tabsa(2).Visible = False
        Case "Go To"
            Tabsa(0).Visible = False
            Tabsa(1).Visible = False
            Tabsa(2).Visible = True
    End Select
End Sub

Private Sub Display(Message As String)
MakeNormal Me
MsgBox Message & " was not found!", vbOKOnly + vbInformation, "Error"
MakeTop Me
End Sub
