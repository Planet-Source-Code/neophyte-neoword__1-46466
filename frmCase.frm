VERSION 5.00
Begin VB.Form frmCase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Case"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3990
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
   ScaleHeight     =   2535
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.OptionButton OptCase 
      Caption         =   "Narrow"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton OptCase 
      Caption         =   "Wide"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton OptCase 
      Caption         =   "Hiragana"
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton OptCase 
      Caption         =   "Katakana"
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton OptCase 
      Caption         =   "Title Case"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton OptCase 
      Caption         =   "lower case"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton OptCase 
      Caption         =   "UPPER CASE"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Select a case:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   1340
      X2              =   3860
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
End
Attribute VB_Name = "frmCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
On Error Resume Next
    With frmMain.ActiveForm.rtfText
            If OptCase(0).Value = True Then
                .SelText = StrConv(.SelText, vbProperCase)
            ElseIf OptCase(1).Value = True Then
                .SelText = StrConv(.SelText, vbUpperCase)
            ElseIf OptCase(2).Value = True Then
                .SelText = StrConv(.SelText, vbLowerCase)
            ElseIf OptCase(3).Value = True Then
                .SelText = StrConv(.SelText, vbWide)
            ElseIf OptCase(4).Value = True Then
                .SelText = StrConv(.SelText, vbNarrow)
            ElseIf OptCase(5).Value = True Then
                .SelText = StrConv(.SelText, vbKatakana)
            ElseIf OptCase(6).Value = True Then
                .SelText = StrConv(.SelText, vbHiragana)
            End If
        .SetFocus
    End With
End Sub

Private Sub Form_Load()
MakeTop Me
End Sub
