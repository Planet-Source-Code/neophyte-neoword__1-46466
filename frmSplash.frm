VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
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
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   315.003
   ScaleMode       =   0  'User
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Label lblEEgg 
      BackStyle       =   0  'Transparent
      Height          =   331
      Left            =   4770
      TabIndex        =   5
      Top             =   271
      Width           =   1140
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Line Line3 
      X1              =   80
      X2              =   400
      Y1              =   191.394
      Y2              =   191.394
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   400
      Y1              =   143.546
      Y2              =   143.546
   End
   Begin VB.Line Line1 
      X1              =   88
      X2              =   88
      Y1              =   0
      Y2              =   159.495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "All rights reserved. This program is protected by US and international copyright laws as described in the Help About."
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
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © Neophyte 2003."
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
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "52364-666-0000025-00034"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   3120
      Width           =   1935
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Long
Me.Show
MakeTop Me
frmMain.Visible = False
lblUsername = GetUsersName
    For i = 0 To Screen.FontCount
    DoEvents
        frmMain.cboFonts.AddItem Screen.Fonts(i)
    Next i
frmMain.cboFonts.AddItem "Easter Egg"
frmMain.Visible = True
Unload Me
End Sub

Private Sub Form_LostFocus()
Me.SetFocus
End Sub

Private Sub lblEEgg_DblClick()
EEgg1 = 1
lblUsername.Caption = "Bobby The Rubber Chicken"
MakeNormal Me
MsgBox "Quuaaaack!!", vbOKOnly + vbInformation, "Easter Egg!"
MakeTop Me
End Sub
