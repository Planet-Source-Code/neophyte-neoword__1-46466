VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDocument 
   BackColor       =   &H80000010&
   Caption         =   "Untitled"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13890
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   12615
      TabIndex        =   5
      Top             =   0
      Width           =   12615
      Begin ComctlLib.Slider sdrTop 
         Height          =   320
         Left            =   720
         TabIndex        =   6
         Top             =   0
         Width           =   11906
         _ExtentX        =   21008
         _ExtentY        =   556
         _Version        =   327682
         LargeChange     =   198
         Max             =   11906
         SelStart        =   1800
         TickFrequency   =   198
         Value           =   1800
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11415
      Left            =   0
      ScaleHeight     =   11415
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   0
      Width           =   435
      Begin ComctlLib.Slider sdrLeft 
         Height          =   16838
         Left            =   0
         TabIndex        =   4
         Top             =   500
         Width           =   320
         _ExtentX        =   556
         _ExtentY        =   29713
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   280
         Max             =   16838
         SelStart        =   1440
         TickFrequency   =   280
         Value           =   1440
      End
   End
   Begin VB.VScrollBar VScroll 
      Height          =   3855
      LargeChange     =   3367
      Left            =   12840
      Max             =   16837
      SmallChange     =   842
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   375
      Width           =   255
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   16838
      Left            =   720
      ScaleHeight     =   16815
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   500
      Width           =   11906
      Begin RichTextLib.RichTextBox rtfText 
         Height          =   13958
         Left            =   0
         TabIndex        =   2
         Top             =   1440
         Width           =   10106
         _ExtentX        =   17833
         _ExtentY        =   24633
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmDocument.frx":058A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line BR2 
         BorderColor     =   &H00C0C0C0&
         X1              =   10106
         X2              =   10106
         Y1              =   15408
         Y2              =   15898
      End
      Begin VB.Line BR1 
         BorderColor     =   &H00C0C0C0&
         X1              =   10106
         X2              =   10486
         Y1              =   15408
         Y2              =   15408
      End
      Begin VB.Line BL2 
         BorderColor     =   &H00C0C0C0&
         X1              =   1800
         X2              =   1800
         Y1              =   15408
         Y2              =   15888
      End
      Begin VB.Line BL1 
         BorderColor     =   &H00C0C0C0&
         X1              =   1800
         X2              =   1320
         Y1              =   15408
         Y2              =   15408
      End
      Begin VB.Line TL2 
         BorderColor     =   &H00C0C0C0&
         X1              =   1800
         X2              =   1800
         Y1              =   1430
         Y2              =   950
      End
      Begin VB.Line TL1 
         BorderColor     =   &H00C0C0C0&
         X1              =   1800
         X2              =   1320
         Y1              =   1430
         Y2              =   1430
      End
      Begin VB.Line TR2 
         BorderColor     =   &H00C0C0C0&
         X1              =   10106
         X2              =   10106
         Y1              =   1430
         Y2              =   950
      End
      Begin VB.Line TR1 
         BorderColor     =   &H00C0C0C0&
         X1              =   10106
         X2              =   10586
         Y1              =   1430
         Y2              =   1430
      End
   End
   Begin VB.Shape BShadow 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   50
      Left            =   840
      Top             =   17338
      Width           =   11881
   End
   Begin VB.Shape RShadow 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   16815
      Left            =   12626
      Top             =   575
      Width           =   50
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Resize()
Form_Resize
End Sub

Private Sub Form_Resize()
picPage.Left = (Me.Width / 2) - (11906 / 2)
rtfText.SelIndent = 1800
VScroll.Left = Me.Width - 375
VScroll.Height = Me.Height - 875
picTop.Width = Me.Width
picLeft.Height = Me.Height
sdrTop.Left = picPage.Left
RShadow.Left = picPage.Left + 11906
BShadow.Left = picPage.Left + 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim DoSave As VbMsgBoxResult
    If Changed = True Then
        DoSave = MsgBox("Do you want to save the changes to " & Me.Caption & "?", vbYesNoCancel + vbExclamation, "Neophyte Word")
            Select Case DoSave
                Case vbYes
                    frmMain.Save
                Case vbCancel
                    Cancel = 1
            End Select
    End If
End Sub

Private Sub rtfText_Change()
Changed = True
End Sub

Private Sub rtfText_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then rtfText.SelRTF = vbTab
End Sub
Private Sub rtfText_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu frmMain.mnuRightRTF
End Sub

Private Sub rtfText_SelChange()
On Error Resume Next
    With frmMain
        .tbToolBar.Buttons("Bold").Value = IIf(rtfText.SelBold, tbrPressed, tbrUnpressed)
        .tbToolBar.Buttons("Italic").Value = IIf(rtfText.SelItalic, tbrPressed, tbrUnpressed)
        .tbToolBar.Buttons("Underline").Value = IIf(rtfText.SelUnderline, tbrPressed, tbrUnpressed)
        .tbToolBar.Buttons("Strike Through").Value = IIf(rtfText.SelStrikeThru, tbrPressed, tbrUnpressed)
        .tbToolBar.Buttons("Align Left").Value = IIf(rtfText.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
        .tbToolBar.Buttons("Center").Value = IIf(rtfText.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
        .tbToolBar.Buttons("Align Right").Value = IIf(rtfText.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
        .cboFonts.Text = rtfText.SelFontName
        .cboSize.Text = rtfText.SelFontSize
        .sbStatusBar.Panels(1).Text = "Ln " & rtfText.GetLineFromChar(rtfText.SelStart) + 1
        .sbStatusBar.Panels(2).Text = "Col " & rtfText.SelStart + 1
    End With
End Sub

Private Sub sdrLeft_Change()
rtfText.Top = sdrLeft.Value
TL1.Y1 = rtfText.Top - 13
TL1.Y2 = rtfText.Top - 13
TL2.Y1 = rtfText.Top - 10
TL2.Y2 = rtfText.Top - 490
TR1.Y1 = rtfText.Top - 10
TR1.Y2 = rtfText.Top - 10
TR2.Y1 = rtfText.Top - 10
TR2.Y2 = rtfText.Top - 490
BL1.Y1 = rtfText.Top + 13968
BL1.Y2 = rtfText.Top + 13968
BL2.Y1 = rtfText.Top + 13968
BL2.Y2 = rtfText.Top + 14448
BR1.Y1 = rtfText.Top + 13968
BR1.Y2 = rtfText.Top + 13968
BR2.Y1 = rtfText.Top + 13968
BR2.Y2 = rtfText.Top + 14448
End Sub

Private Sub sdrTop_Change()
rtfText.SelIndent = sdrTop.Value
TL1.X1 = sdrTop.Value
TL1.X2 = sdrTop.Value - 480
TL2.X1 = sdrTop.Value
TL2.X2 = sdrTop.Value
BL1.X1 = sdrTop.Value
BL1.X2 = sdrTop.Value - 480
BL2.X1 = sdrTop.Value
BL2.X2 = sdrTop.Value
End Sub

Private Sub VScroll_Change()
picPage.Top = 500 - VScroll.Value
sdrLeft.Top = 500 - VScroll.Value
RShadow.Top = picPage.Top + 75
BShadow.Top = picPage.Top + picPage.Height
End Sub
