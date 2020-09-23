VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Neophyte Word"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9075
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   6330
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Ln 1"
            TextSave        =   "Ln 1"
            Object.ToolTipText     =   "Line"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Col 1"
            TextSave        =   "Col 1"
            Object.ToolTipText     =   "Column"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Caps Lock"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Num Lock"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "SCRL"
            Object.ToolTipText     =   "Scroll Lock"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "OVR"
            TextSave        =   "OVR"
            Object.ToolTipText     =   "Override"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlMenuIcons 
      Left            =   720
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D58
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E41
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F28
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F87
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2068
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2100
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2163
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2251
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2302
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbFonts 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   714
      ButtonWidth     =   609
      ButtonHeight    =   714
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      _Version        =   393216
      Begin VB.ComboBox cboSize 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":235B
         Left            =   2550
         List            =   "frmMain.frx":2392
         TabIndex        =   3
         Text            =   "12"
         Top             =   50
         Width           =   735
      End
      Begin VB.ComboBox cboFonts 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   50
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "Times New Roman"
         Top             =   50
         Width           =   2415
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1320
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   13
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Strike Through"
            Object.ToolTipText     =   "Strike Through"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font Colour"
            Object.ToolTipText     =   "Font Colour"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Default"
                  Text            =   "Default"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Select Colour"
                  Text            =   "Select Colour..."
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7040
         ScaleHeight     =   375
         ScaleWidth      =   345
         TabIndex        =   4
         Top             =   0
         Width           =   345
         Begin VB.Line Line4 
            BorderColor     =   &H8000000F&
            X1              =   0
            X2              =   330
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Line3 
            BorderColor     =   &H8000000F&
            X1              =   320
            X2              =   0
            Y1              =   320
            Y2              =   320
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000F&
            X1              =   330
            X2              =   330
            Y1              =   0
            Y2              =   330
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000F&
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   330
         End
         Begin VB.Label lblFontColour 
            Alignment       =   2  'Center
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   5
            ToolTipText     =   "Font Colour"
            Top             =   0
            Width           =   345
         End
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23DA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24EC
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25FE
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2710
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2822
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2934
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A46
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B58
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C6A
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D7C
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E8E
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FA0
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30B2
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31C4
            Key             =   "Strike Through"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32D6
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33E8
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34FA
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":360C
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select Al&l"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "Clear"
         Begin VB.Menu mnuEditClearFormat 
            Caption         =   "&Format"
         End
         Begin VB.Menu mnuEditClearContents 
            Caption         =   "&Contents"
         End
      End
      Begin VB.Menu mnuEditBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "R&eplace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditGoto 
         Caption         =   "&Go To..."
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbars 
         Caption         =   "&Toolbars"
         Begin VB.Menu mnuviewToolbarsMain 
            Caption         =   "&Main"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarsFont 
            Caption         =   "&Font"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewRulers 
         Caption         =   "R&ulers"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuInsertTab 
         Caption         =   "T&ab"
      End
      Begin VB.Menu mnuInsertBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertPicture 
         Caption         =   "&Picture..."
      End
      Begin VB.Menu mnuInsertFile 
         Caption         =   "&File..."
      End
      Begin VB.Menu mnuInsertBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertDate 
         Caption         =   "Date and &Time..."
      End
      Begin VB.Menu mnuInsertSymbol 
         Caption         =   "&Symbol..."
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuFormatfont 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuFormatBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatAlign 
         Caption         =   "&Align"
         Begin VB.Menu mnuFormatAlignLeft 
            Caption         =   "&Left"
         End
         Begin VB.Menu mnuFormatAlignCenter 
            Caption         =   "&Center"
         End
         Begin VB.Menu mnuFormatAlignRight 
            Caption         =   "&Right"
         End
      End
      Begin VB.Menu mnuFormatBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatDropCap 
         Caption         =   "&Drop Cap"
      End
      Begin VB.Menu mnuFormatChangeCase 
         Caption         =   "Change Cas&e..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsWordCount 
         Caption         =   "&Word Count..."
      End
      Begin VB.Menu mnuToolsBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
   Begin VB.Menu mnuRightRTF 
      Caption         =   "RightRTF"
      Visible         =   0   'False
      Begin VB.Menu mnuRightRTFCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuRightRTFCut 
         Caption         =   "C&ut"
      End
      Begin VB.Menu mnuRightRTFPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuRightRTFBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightRTFSelectAll 
         Caption         =   "Select Al&l"
      End
      Begin VB.Menu mnuRightRTFBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightRTFFont 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuRightRTFColour 
         Caption         =   "Colou&r..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Const EM_UNDO = &HC7
Const MF_BYPOSITION = &H400&
Const WM_PASTE = &H302
Dim CurrentColour As Long

Private Sub cboFonts_Click()
On Error Resume Next
    If cboFonts.Text = "Easter Egg" Then
        EEgg3 = 1
        MsgBox "Easter Egg 3", vbOKOnly + vbInformation, "Easter Egg!"
    Else
        ActiveForm.rtfText.SelFontName = cboFonts.Text
    End If
End Sub

Private Sub cboFonts_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
ActiveForm.rtfText.SelFontName = cboFonts.Text
    If cboFonts.Text = "Easter Egg" Then
        EEgg3 = 1
        MsgBox "Easter Egg 3"
    Else
        ActiveForm.rtfText.SelFontName = cboFonts.Text
    End If
    If KeyCode = vbKeyReturn Then ActiveForm.rtfText.SetFocus
End Sub

Private Sub cboSize_Click()
On Error Resume Next
ActiveForm.rtfText.SelFontSize = cboSize.Text
End Sub

Private Sub cboSize_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
ActiveForm.rtfText.SelFontSize = cboSize.Text
If KeyCode = vbKeyReturn Then ActiveForm.rtfText.SetFocus
End Sub

Private Sub lblFontColour_Click()
ActiveForm.rtfText.SelColor = CurrentColour
ButtonNormal
End Sub

Private Sub lblFontColour_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ButtonDown
End Sub

Private Sub lblFontColour_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ButtonOver
End Sub

Private Sub lblFontColour_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
ButtonUp
End Sub

Private Sub MDIForm_Load()
LoadNewDoc
CurrentColour = vbBlack
    With imlMenuIcons
        SetMenuItemBitmaps GetSubMenu(GetMenu(Me.hwnd), 1), 0, MF_BYPOSITION, .ListImages(1).Picture, .ListImages(1).Picture
        SetMenuItemBitmaps GetSubMenu(GetMenu(Me.hwnd), 1), 1, MF_BYPOSITION, .ListImages(2).Picture, .ListImages(2).Picture
        SetMenuItemBitmaps GetSubMenu(GetMenu(Me.hwnd), 1), 4, MF_BYPOSITION, .ListImages(3).Picture, .ListImages(3).Picture
        SetMenuItemBitmaps GetSubMenu(GetMenu(Me.hwnd), 1), 5, MF_BYPOSITION, .ListImages(3).Picture, .ListImages(3).Picture
        SetMenuItemBitmaps GetSubMenu(GetMenu(Me.hwnd), 1), 8, MF_BYPOSITION, .ListImages(4).Picture, .ListImages(4).Picture
        SetMenuItemBitmaps GetSubMenu(GetMenu(Me.hwnd), 2), 0, MF_BYPOSITION, .ListImages(5).Picture, .ListImages(5).Picture
        SetMenuItemBitmaps GetSubMenu(GetMenu(Me.hwnd), 2), 1, MF_BYPOSITION, .ListImages(6).Picture, .ListImages(6).Picture
        SetMenuItemBitmaps GetSubMenu(GetMenu(Me.hwnd), 2), 3, MF_BYPOSITION, .ListImages(7).Picture, .ListImages(7).Picture
        SetMenuItemBitmaps GetSubMenu(GetMenu(Me.hwnd), 2), 4, MF_BYPOSITION, .ListImages(8).Picture, .ListImages(8).Picture
        SetMenuItemBitmaps GetSubMenu(GetMenu(Me.hwnd), 2), 5, MF_BYPOSITION, .ListImages(9).Picture, .ListImages(9).Picture
        SetMenuItemBitmaps GetSubMenu(GetMenu(Me.hwnd), 2), 6, MF_BYPOSITION, .ListImages(10).Picture, .ListImages(10).Picture
        SetMenuItemBitmaps GetSubMenu(GetMenu(Me.hwnd), 2), 12, MF_BYPOSITION, .ListImages(11).Picture, .ListImages(11).Picture
        SetMenuItemBitmaps GetSubMenu(GetMenu(Me.hwnd), 5), 0, MF_BYPOSITION, .ListImages(12).Picture, .ListImages(12).Picture
        SetMenuItemBitmaps GetSubMenu(GetSubMenu(GetMenu(Me.hwnd), 5), 2), 0, MF_BYPOSITION, .ListImages(13).Picture, .ListImages(13).Picture
        SetMenuItemBitmaps GetSubMenu(GetSubMenu(GetMenu(Me.hwnd), 5), 2), 1, MF_BYPOSITION, .ListImages(14).Picture, .ListImages(14).Picture
        SetMenuItemBitmaps GetSubMenu(GetSubMenu(GetMenu(Me.hwnd), 5), 2), 2, MF_BYPOSITION, .ListImages(15).Picture, .ListImages(15).Picture
    End With
End Sub

Private Sub LoadNewDoc()
Dim NewForm As frmDocument
Set NewForm = New frmDocument
NewForm.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuEditClearContents_Click()
ActiveForm.rtfText.Text = ""
End Sub

Private Sub mnuEditClearFormat_Click()
    With ActiveForm.rtfText
            If .SelText = "" Then
                .SelStart = 0
                .SelLength = Len(.Text)
            End If
        .SelAlignment = rtfLeft
        .SelFontName = "Times New Roman"
        .SelFontSize = "12"
        .SelBold = False
        .SelItalic = False
        .SelUnderline = False
        .SelStrikeThru = False
        .SelColor = vbBlack
        .SelLength = 0
    End With
ActiveForm.sdrLeft.Value = 1440
ActiveForm.sdrTop.Value = 1800
End Sub

Private Sub mnuEditDelete_Click()
ActiveForm.rtfText.SelText = ""
End Sub

Private Sub mnuEditFind_Click()
frmFindReplaceGoto.Show
End Sub

Private Sub mnuEditGoto_Click()
frmFindReplaceGoto.Show
End Sub

Private Sub mnuEditReplace_Click()
frmFindReplaceGoto.Show
End Sub

Private Sub mnuEditSelectAll_Click()
ActiveForm.rtfText.SelStart = 0
ActiveForm.rtfText.SelLength = Len(ActiveForm.rtfText.Text)
End Sub

Private Sub mnuFormatAlignCenter_Click()
ActiveForm.rtfText.SelAlignment = rtfCenter
End Sub

Private Sub mnuFormatAlignLeft_Click()
ActiveForm.rtfText.SelAlignment = rtfLeft
End Sub

Private Sub mnuFormatAlignRight_Click()
ActiveForm.rtfText.SelAlignment = rtfRight
End Sub

Private Sub mnuFormatChangeCase_Click()
frmCase.Show
End Sub

Private Sub mnuFormatDropCap_Click()
ActiveForm.rtfText.SelText = StrConv(ActiveForm.rtfText.SelText, vbLowerCase)
End Sub

Private Sub mnuFormatFont_Click()
    With dlgCommonDialog
        .FontName = ActiveForm.rtfText.SelFontName
        .FontSize = ActiveForm.rtfText.SelFontSize
        .FontBold = ActiveForm.rtfText.SelBold
        .FontItalic = ActiveForm.rtfText.SelItalic
        .FontUnderline = ActiveForm.rtfText.SelUnderline
        .FontStrikethru = ActiveForm.rtfText.SelStrikeThru
        .Color = ActiveForm.rtfText.SelColor
        .Flags = cdlCFBoth Or cdlCFEffects Or cdlCFForceFontExist
        .ShowFont
        ActiveForm.rtfText.SelFontName = .FontName
        ActiveForm.rtfText.SelFontSize = .FontSize
        ActiveForm.rtfText.SelBold = .FontBold
        ActiveForm.rtfText.SelItalic = .FontItalic
        ActiveForm.rtfText.SelUnderline = .FontUnderline
        ActiveForm.rtfText.SelStrikeThru = .FontStrikethru
        ActiveForm.rtfText.SelColor = .Color
    End With
End Sub

Private Sub mnuInsertDate_Click()
frmAddDate.Show
End Sub

Private Sub mnuInsertFile_Click()
On Error GoTo Cancel
    With dlgCommonDialog
        .DialogTitle = "Select a File..."
        .Filter = "All Files (*.*)|*.*"
        .Flags = cdlOFNFileMustExist Or cdlOFNNoDereferenceLinks
        .CancelError = True
        .ShowOpen
        ActiveForm.rtfText.OLEObjects.Add , , .FileName
    End With
Cancel:
ClearFilename
End Sub

Private Sub mnuInsertPicture_Click()
On Error GoTo Cancel
    With dlgCommonDialog
        .DialogTitle = "Select a Picture..."
        .Filter = "Picture Files (*.bmp, *.jpg, *.gif, etc)|*.bmp;*.dib;*.jpeg;*.jpg;*.jpe;*.jfif;*.gif;*.tif;*.tiff;*.png"
        .Flags = cdlOFNFileMustExist Or cdlOFNNoDereferenceLinks
        .CancelError = True
        .ShowOpen
        Clipboard.Clear
            'If PictureWidth > ActiveForm.rtfText.Width - ActiveForm.rtfText.SelIndent Then
                'Percent = 100 / PictureWidth * (ActiveForm.rtfText.Width - ActiveForm.rtfText.SelIndent)
                'PictureHeight = PictureHeight / 100 * Percent
                'PictureWidth = PictureWidth / 100 * Percent
            'End If
        Clipboard.SetData LoadPicture(.FileName)
        SendMessage ActiveForm.rtfText.hwnd, WM_PASTE, 0&, 0&
    End With
Cancel:
ClearFilename
End Sub

Private Sub mnuInsertSymbol_Click()
frmSymbol.Show
End Sub

Private Sub mnuInsertTab_Click()
ActiveForm.rtfText.SelRTF = vbTab
End Sub

Private Sub mnuRightRTFColour_Click()
    With dlgCommonDialog
        .Color = ActiveForm.rtfText.SelColor
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .ShowColor
        CurrentColour = .Color
        ActiveForm.rtfText.SelColor = .Color
        lblFontColour.ForeColor = .Color
    End With
End Sub

Private Sub mnuRightRTFCopy_Click()
mnuEditCopy_Click
End Sub

Private Sub mnuRightRTFCut_Click()
mnuEditCut_Click
End Sub

Private Sub mnuRightRTFFont_Click()
mnuFormatFont_Click
End Sub

Private Sub mnuRightRTFPaste_Click()
mnuEditPaste_Click
End Sub

Private Sub mnuRightRTFSelectAll_Click()
ActiveForm.rtfText.SelStart = 0
ActiveForm.rtfText.SelLength = Len(ActiveForm.rtfText.Text)
End Sub

Private Sub mnuToolsWordCount_Click()
frmStats.Show
End Sub

Private Sub mnuViewRulers_Click()
    With ActiveForm
        mnuViewRulers.Checked = Not mnuViewRulers.Checked
        .picTop.Visible = mnuViewRulers.Checked
        .picLeft.Visible = mnuViewRulers.Checked
            If .picTop.Visible = False Then
                .VScroll.Top = 0
                .VScroll.Height = .Height - 500
            Else
                .VScroll.Top = 375
                .VScroll.Height = .Height - 875
            End If
    End With
End Sub

Private Sub mnuViewToolbarsFont_Click()
mnuViewToolbarsFont.Checked = Not mnuViewToolbarsFont.Checked
tbFonts.Visible = mnuViewToolbarsFont.Checked
End Sub

Private Sub mnuviewToolbarsMain_Click()
mnuviewToolbarsMain.Checked = Not mnuviewToolbarsMain.Checked
tbToolBar.Visible = mnuviewToolbarsMain.Checked
End Sub

Private Sub picFontColour_Click()
lblFontColour_Click
End Sub

Private Sub tbFonts_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ButtonNormal
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Undo"
            'mnuEditUndo_Click
        Case "Redo"
            'mnuEditRedo_Click
        Case "Find"
            mnuEditFind_Click
        Case "Bold"
            ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
            Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
            Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
            Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Strike Through"
            ActiveForm.rtfText.SelStrikeThru = Not ActiveForm.rtfText.SelStrikeThru
            Button.Value = IIf(ActiveForm.rtfText.SelStrikeThru, tbrPressed, tbrUnpressed)
        Case "Align Left"
            ActiveForm.rtfText.SelAlignment = rtfLeft
        Case "Center"
            ActiveForm.rtfText.SelAlignment = rtfCenter
        Case "Align Right"
            ActiveForm.rtfText.SelAlignment = rtfRight
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuWindowNewWindow_Click()
LoadNewDoc
End Sub

Private Sub mnuViewRefresh_Click()
ActiveForm.Refresh
End Sub

Private Sub mnuViewStatusBar_Click()
mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuEditPaste_Click()
On Error Resume Next
ActiveForm.rtfText.SelRTF = Clipboard.GetText
End Sub

Private Sub mnuEditCopy_Click()
On Error Resume Next
Clipboard.SetText ActiveForm.rtfText.SelRTF
End Sub

Private Sub mnuEditCut_Click()
On Error Resume Next
Clipboard.SetText ActiveForm.rtfText.SelRTF
ActiveForm.rtfText.SelText = vbNullString
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFilePrint_Click()
On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    With dlgCommonDialog
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
            If ActiveForm.rtfText.SelLength = 0 Then
                .Flags = .Flags + cdlPDAllPages
            Else
                .Flags = .Flags + cdlPDSelection
            End If
        .ShowPrinter
            If Err <> MSComDlg.cdlCancel Then
                ActiveForm.rtfText.SelPrint .hDC
            End If
    End With
End Sub

Private Sub mnuFilePageSetup_Click()
On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With
End Sub

Private Sub mnuFileSaveAs_Click()
Dim sFile As String
    If ActiveForm Is Nothing Then Exit Sub
    With dlgCommonDialog
        .DialogTitle = "Save As"
        .CancelError = False
        .Filter = "Rich Text Documents (*.rtf)|*.rtf|Plain Text Files (*.txt)|*.txt|Word Documents (*.doc)|*.doc|All Files (*.*)|*.*"
        .ShowSave
            If Len(.FileName) = 0 Then Exit Sub
        sFile = .FileName
        ActiveForm.rtfText.SaveFile sFile, IIf(Right$(.FileName, 3) = "txt", rtfText, rtfRTF)
    End With
ActiveForm.Caption = sFile
Changed = False
ClearFilename
End Sub

Private Sub mnuFileSave_Click()
Dim sFile As String
    If Right$(ActiveForm.Caption, 8) = "Untitled" Then
        mnuFileSaveAs_Click
    Else
        sFile = ActiveForm.Caption
        ActiveForm.rtfText.SaveFile sFile
    End If
Changed = False
ClearFilename
End Sub

Private Sub mnuFileClose_Click()
ActiveForm.Unload
End Sub

Private Sub mnuFileOpen_Click()
Dim sFile As String
    If ActiveForm Is Nothing Then LoadNewDoc
    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        .Filter = "NeoWord Files (*.rtf, *.txt)|*.rtf;*.txt|Rich Text Documents (*.rtf)|*.rtf|Plain Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .ShowOpen
            If Len(.FileName) = 0 Then Exit Sub
        sFile = .FileName
    End With
ActiveForm.rtfText.LoadFile sFile
ActiveForm.Caption = sFile
ClearFilename
End Sub

Private Sub mnuFileNew_Click()
LoadNewDoc
End Sub

Private Sub tbToolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Select Colour"
                With dlgCommonDialog
                    .Color = ActiveForm.rtfText.SelColor
                    .Flags = cdlCCFullOpen Or cdlCCRGBInit
                    .ShowColor
                    CurrentColour = .Color
                    ActiveForm.rtfText.SelColor = .Color
                    lblFontColour.ForeColor = .Color
                End With
        Case "Default"
            ActiveForm.rtfText.SelColor = vbBlack
    End Select
End Sub

Private Sub tbToolBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ButtonNormal
End Sub

Private Sub ButtonNormal()
Line1.BorderColor = vbButtonFace
Line2.BorderColor = vbButtonFace
Line3.BorderColor = vbButtonFace
Line4.BorderColor = vbButtonFace
lblFontColour.Top = 0
lblFontColour.Left = 0
End Sub

Private Sub ButtonUp()
Line1.BorderColor = vbButtonFace
Line2.BorderColor = vbButtonFace
Line3.BorderColor = vbButtonFace
Line4.BorderColor = vbButtonFace
lblFontColour.Top = 20
lblFontColour.Left = 20
End Sub

Private Sub ButtonDown()
Line1.BorderColor = vbButtonShadow
Line2.BorderColor = vbHighlightText
Line3.BorderColor = vbHighlightText
Line4.BorderColor = vbButtonShadow
lblFontColour.Top = 20
lblFontColour.Left = 20
End Sub

Private Sub ButtonOver()
Line1.BorderColor = vbHighlightText
Line2.BorderColor = vbButtonShadow
Line3.BorderColor = vbButtonShadow
Line4.BorderColor = vbHighlightText
End Sub

Private Sub ClearFilename()
dlgCommonDialog.FileName = ""
End Sub

Public Sub Save()
mnuFileSave_Click
End Sub

Public Sub Font()
mnuFormatFont_Click
End Sub
