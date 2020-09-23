VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAxiomMain 
   Caption         =   "Axiom"
   ClientHeight    =   5790
   ClientLeft      =   540
   ClientTop       =   750
   ClientWidth     =   8505
   Icon            =   "Axiom_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   567
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picRevert 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   585
      Picture         =   "Axiom_Main.frx":0CCA
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   21
      Top             =   2880
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picAppend 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      Picture         =   "Axiom_Main.frx":0DB4
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   20
      Top             =   4455
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picAbout 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   585
      Picture         =   "Axiom_Main.frx":0E32
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   1620
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picInsert 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   585
      Picture         =   "Axiom_Main.frx":0F1C
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   1350
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picUndo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   585
      Picture         =   "Axiom_Main.frx":1006
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   17
      Top             =   990
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picSwap 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      Picture         =   "Axiom_Main.frx":1084
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   4140
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picAll 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      Picture         =   "Axiom_Main.frx":116E
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   3825
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picX 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      Picture         =   "Axiom_Main.frx":11EC
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   3555
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picToRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      Picture         =   "Axiom_Main.frx":126A
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picToLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      Picture         =   "Axiom_Main.frx":12E8
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   2925
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picPaste 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      Picture         =   "Axiom_Main.frx":1366
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   2655
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picCopy 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      Picture         =   "Axiom_Main.frx":13E4
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   2340
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picCut 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      Picture         =   "Axiom_Main.frx":1462
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   1980
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picSave 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      Picture         =   "Axiom_Main.frx":14E0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   990
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picNew 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      Picture         =   "Axiom_Main.frx":155E
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   7
      Top             =   1350
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox picOpen 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      Picture         =   "Axiom_Main.frx":15DC
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   1665
      Visible         =   0   'False
      Width           =   195
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insert"
            Object.ToolTipText     =   "Insert File"
            ImageKey        =   "insert"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reveret"
            Object.ToolTipText     =   "Revert (Loose all changes)"
            ImageKey        =   "reveret"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "saveas"
            Object.ToolTipText     =   "Save As"
            ImageKey        =   "saveas"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "move"
            Object.ToolTipText     =   "move"
            ImageKey        =   "move"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   1935
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Axiom_Main.frx":1656
            Key             =   "move"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Axiom_Main.frx":197E
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Axiom_Main.frx":27D2
            Key             =   "saveas"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Axiom_Main.frx":2AEE
            Key             =   "new"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Axiom_Main.frx":2E0A
            Key             =   "options"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Axiom_Main.frx":3126
            Key             =   "reveret"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Axiom_Main.frx":3A02
            Key             =   "insert"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status 
      Height          =   375
      Left            =   1890
      TabIndex        =   1
      Top             =   5130
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   23812
            MinWidth        =   23812
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   7290
      Top             =   2745
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin RichTextLib.RichTextBox MainText 
      Height          =   3345
      Left            =   900
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   5900
      _Version        =   393217
      BackColor       =   15921906
      HideSelection   =   0   'False
      ScrollBars      =   3
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"Axiom_Main.frx":3D1E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip Tabs 
      Height          =   330
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5130
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      MultiRow        =   -1  'True
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Input"
            Key             =   "Input"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Output"
            Key             =   "Output"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pad"
            Key             =   "Pad"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox OutputText 
      Height          =   3345
      Left            =   1440
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   5900
      _Version        =   393217
      BackColor       =   16776178
      HideSelection   =   0   'False
      ScrollBars      =   3
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"Axiom_Main.frx":3E2C
   End
   Begin RichTextLib.RichTextBox PadText 
      Height          =   3255
      Left            =   1890
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1575
      Visible         =   0   'False
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   5741
      _Version        =   393217
      BackColor       =   16056319
      HideSelection   =   0   'False
      ScrollBars      =   3
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"Axiom_Main.frx":3F3A
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "mnuRightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuRightClick_Undo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuRightClick_Cut 
         Caption         =   "Cu&t"
      End
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
      Begin VB.Menu mnuFileInsert 
         Caption         =   "&Insert File..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFileRevert 
         Caption         =   "&Revert"
         Enabled         =   0   'False
      End
      Begin VB.Menu zMenuSeperator0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu zMenuSeperator13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Preferences..."
      End
      Begin VB.Menu zMenuSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMRU 
         Caption         =   "MRU"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu Hiphen 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMove 
      Caption         =   "&Move"
      Begin VB.Menu mnuIn_to_Out 
         Caption         =   "Move &Input to Output"
      End
      Begin VB.Menu mnuOut_to_In 
         Caption         =   "Move &Output to Input"
      End
      Begin VB.Menu zMenuSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIn_to_Pad 
         Caption         =   "&Append Input to Pad"
      End
      Begin VB.Menu mnuOut_to_Pad 
         Caption         =   "&Append Output to Pad"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Shortcut        =   ^Y
         Visible         =   0   'False
      End
      Begin VB.Menu zMenuSeperator3 
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
         Caption         =   "De&lete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu zMenuSeperator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLocalEdit 
         Caption         =   "&Local Edit"
         Begin VB.Menu mnuLocalCut 
            Caption         =   "Cu&t"
            Shortcut        =   +{DEL}
         End
         Begin VB.Menu mnuLocalCopy 
            Caption         =   "&Copy"
            Shortcut        =   ^{INSERT}
         End
         Begin VB.Menu mnuLocalPaste 
            Caption         =   "&Paste..."
            Shortcut        =   +{INSERT}
         End
         Begin VB.Menu mnuLocalClip 
            Caption         =   "&Show Local Clipboard..."
            Shortcut        =   {F4}
         End
      End
      Begin VB.Menu zMenuSeperator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAppend 
         Caption         =   "&Append"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEditSwap 
         Caption         =   "S&wap"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu zMenuSeperator12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find / Replace ..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuOperations 
      Caption         =   "&Lines"
      Begin VB.Menu mnuAddBA 
         Caption         =   "&Add Before/After Every Line..."
      End
      Begin VB.Menu mnuAddLineNumbers 
         Caption         =   "&Add Line Numbers..."
      End
      Begin VB.Menu mnuMaxLineWidth 
         Caption         =   "&Set Maximum Line Width"
      End
      Begin VB.Menu mnuSetMaxWord 
         Caption         =   "Set Max Line width (Word)"
      End
      Begin VB.Menu zMenuSeperator6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteChars 
         Caption         =   "&Delete Chars..."
      End
      Begin VB.Menu mnuInsertString 
         Caption         =   "&Insert..."
      End
      Begin VB.Menu zMenuSeperator10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSort 
         Caption         =   "&Sort"
         Begin VB.Menu mnuSortAscending 
            Caption         =   "Ascending    (A to Z)"
         End
         Begin VB.Menu mnuSortDescending 
            Caption         =   "Descending  (Z to A)"
         End
      End
      Begin VB.Menu mnuTrimSpaces 
         Caption         =   "&Trim Spaces"
         Begin VB.Menu mnuTrimLeft 
            Caption         =   "&Left Only"
         End
         Begin VB.Menu mnuTrimRight 
            Caption         =   "&Right Only"
         End
         Begin VB.Menu mnuTrimBoth 
            Caption         =   "&Trim Both"
         End
      End
      Begin VB.Menu zMenuSeperator11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLinesCompact 
         Caption         =   "&Compact Blank Lines"
      End
      Begin VB.Menu mnuRemoveBlankLines 
         Caption         =   "Remove Blank Lines"
      End
   End
   Begin VB.Menu mnuText 
      Caption         =   "&Text"
      Begin VB.Menu mnuRemoveSpaces 
         Caption         =   "Remove Successive Spaces"
      End
      Begin VB.Menu mnuDotBreak 
         Caption         =   "&Breack only after Dot"
      End
      Begin VB.Menu mnuTab2Spc 
         Caption         =   "&Convert TABs to SPACEs"
      End
      Begin VB.Menu mnuRemoveChars 
         Caption         =   "&Remove Non-Alphanumeric Chars"
      End
      Begin VB.Menu zMenuSeperator7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeCase 
         Caption         =   "&Change Case"
         Begin VB.Menu mnuUpperCase 
            Caption         =   "&UPPER CASE"
         End
         Begin VB.Menu mnuLowerCase 
            Caption         =   "&lower case"
         End
         Begin VB.Menu mnuTitleCase 
            Caption         =   "&Title Case"
         End
      End
      Begin VB.Menu mnuTextFixNewLine 
         Caption         =   "&Fix NewLine Char"
      End
      Begin VB.Menu mnuFromUnicode 
         Caption         =   "Convert Unicode to ANSI"
      End
      Begin VB.Menu mnuTextReverse 
         Caption         =   "&Reverse"
         Begin VB.Menu mnuReverseText 
            Caption         =   "&Entire Text"
         End
         Begin VB.Menu mnuReverseLines 
            Caption         =   "&Each Line"
         End
      End
      Begin VB.Menu zhyph1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuText_RXFind 
         Caption         =   "RegExp Extract/Replace..."
      End
      Begin VB.Menu mnuTextTrigger 
         Caption         =   "Prompted Replace..."
      End
   End
   Begin VB.Menu mnuHTML 
      Caption         =   "&HTML"
      Begin VB.Menu mnuTags 
         Caption         =   "Remove/Extract Tags..."
      End
      Begin VB.Menu mnuRemoveAllTags 
         Caption         =   "Remove All Tags"
      End
      Begin VB.Menu mnuRemoveHtmlComments 
         Caption         =   "Remove Comments <!--  -->"
      End
      Begin VB.Menu mnuHTML_RemovePath 
         Caption         =   "Remove Path..."
      End
      Begin VB.Menu zMenuSeperator8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExtractHREFs 
         Caption         =   "Extract HREFs"
      End
      Begin VB.Menu mnuHTMLAddBR 
         Caption         =   "Add <&BR>"
      End
      Begin VB.Menu mnuHTMLize 
         Caption         =   "&HTMLize..."
      End
      Begin VB.Menu n7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCSS 
         Caption         =   "Put CSS file content inside HTML"
      End
      Begin VB.Menu mnuHTML_ValidateImg 
         Caption         =   "&Validate IMG Tags"
      End
   End
   Begin VB.Menu mnuPlugIns 
      Caption         =   "Plug-ins"
      Begin VB.Menu mnuPlugInX 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmAxiomMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' Code for the 3 RichTextBoxes resides in mo_TheTextBox event procs
Public WithEvents mo_TheTextBox As RichTextBox
Attribute mo_TheTextBox.VB_VarHelpID = -1

Public mo_CStrList As CStrList

'For Removing:
Dim gv_T As Long
#Const TIMING = True

Private Sub KindaFixForRTFBoxes()

' for some reason, the RichTextBox automatically switches language to
' "Ar" , the following "kinda" prevents this behavour.
' For some fonts, even this does not work!!!!
MainText.Text = "X": MainText.Text = ""
OutputText.Text = "X": OutputText.Text = ""
PadText.Text = "X": PadText.Text = ""
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub


Function SelectInput() As String
'//Form-Specific Procedure // Do NOT Move.

If MainText.SelLength <> 0 Then

    SelectInput = MainText.SelText

Else

    SelectInput = MainText.Text
End If

#If TIMING = True Then
    gv_T = GetTickCount()
#End If

Status.Panels.Item(4).Text = "Working..."

End Function

Private Sub RedirectOutput(OutText As String)
'//Form-Specific Procedure // Do NOT Move.

If AxiomSettings.MultipleTextBoxes Then
        'Redirect to "OUTPUT" Textbox
        OutputText.Text = OutText
        Tabs.Tabs.Item("Output").Selected = True
        Set mo_TheTextBox = OutputText
        
    Else
        'Redirect to "INPUT" Textbox
        If Len(MainText.SelText) <> 0 Then
            MainText.SelText = OutText
        Else
            'There is no selection // Replace All
            MainText.Text = OutText
        End If
        Tabs.Tabs.Item("Input").Selected = True
        Set mo_TheTextBox = MainText
    End If

Status.Panels.Item(4).Text = "Ready"

#If TIMING = True Then
    Status.Panels.Item(4).Text = Format((GetTickCount - gv_T) / 1000, "##0.00") & " Seconds"
#End If

End Sub
Private Sub SetMenuPix()

SetMenuIcon Me.hWnd, 0, 0, picNew
SetMenuIcon Me.hWnd, 0, 1, picOpen
SetMenuIcon Me.hWnd, 0, 2, picInsert
SetMenuIcon Me.hWnd, 0, 3, picRevert
'-
SetMenuIcon Me.hWnd, 0, 5, picSave

SetMenuIcon Me.hWnd, 1, 0, picToRight
SetMenuIcon Me.hWnd, 1, 1, picToLeft

SetMenuIcon Me.hWnd, 2, 0, picUndo


SetMenuIcon Me.hWnd, 2, 2, picCut
SetMenuIcon Me.hWnd, 2, 3, picCopy
SetMenuIcon Me.hWnd, 2, 4, picPaste
SetMenuIcon Me.hWnd, 2, 5, picX
'....
SetMenuIcon Me.hWnd, 2, 9, picAppend
SetMenuIcon Me.hWnd, 2, 10, picSwap
SetMenuIcon Me.hWnd, 2, 11, picAll

'....
SetMenuIcon Me.hWnd, 7, 0, picAbout

End Sub

Public Sub UpdateMRU()
Dim idx As Long
    'MRU List
    If go_MRU.Count > 0 Then
            frmAxiomMain.Hiphen.Visible = True
            For idx = 0 To go_MRU.Count - 1
                frmAxiomMain.mnuMRU(idx).Caption = "&" & Trim$(CStr(idx + 1)) _
                & " " & go_MRU.Item(idx)
                frmAxiomMain.mnuMRU(idx).Visible = True
            Next idx
    End If

End Sub

Private Function WhichBox() As WhichTextbox
'Decide which Text Box is currently active.

    If mo_TheTextBox Is MainText Then
        WhichBox = Input_TextBox
    ElseIf mo_TheTextBox Is OutputText Then
        WhichBox = Output_TextBox
    ElseIf mo_TheTextBox Is PadText Then
        WhichBox = Pad_TextBox
    End If
    
End Function




Private Sub foo_Click()

Dim sTemp As String
sTemp = SelectInput()

sTemp = (sTemp)

RedirectOutput sTemp


End Sub

Private Sub Form_Activate()

On Error Resume Next
    mo_TheTextBox.SetFocus

End Sub

Private Sub Form_Load()
Dim idx As Long

For idx = 1 To 4
    Load mnuMRU(idx)
Next idx

Set mo_TheTextBox = MainText
Set mo_CStrList = New CStrList

'Applay settings (from settings class)
AxiomSettings.ApplySettings


'Add images to menus
SetMenuPix

'Fix Keyboard Auto-Switching to "Ar"
KindaFixForRTFBoxes

'Public Properties:
IsDirty(All_TextBoxes) = False
CurrentDir = ""
CurrentFile = ""

'Load Plug-Ins:

If go_PlugIns.Count > 0 Then
    For idx = 0 To go_PlugIns.Count - 1
        Load mnuPlugInX(idx + 1)
        mnuPlugInX(idx + 1).Visible = True
        mnuPlugInX(idx + 1).Caption = go_PlugIns.FunctionDescription(idx) '& idx
    Next idx
    mnuPlugInX(0).Visible = False
Else
    mnuPlugInX(0).Caption = "No Plug-Ins Found"
    mnuPlugInX(0).Enabled = False
End If

End Sub
Private Sub Form_Resize()
On Error Resume Next 'for Minimize

Tabs.Left = 0
Tabs.Width = Me.ScaleWidth

MainText.Top = Toolbar.Height
MainText.Left = 0
MainText.Width = Me.ScaleWidth
MainText.Height = Me.ScaleHeight - Status.Height - Toolbar.Height

OutputText.Top = Toolbar.Height
OutputText.Left = 0
OutputText.Width = Me.ScaleWidth
OutputText.Height = Me.ScaleHeight - Status.Height - Toolbar.Height

PadText.Top = Toolbar.Height
PadText.Left = 0
PadText.Width = Me.ScaleWidth
PadText.Height = Me.ScaleHeight - Status.Height - Toolbar.Height


Tabs.Top = MainText.Height + Toolbar.Height
Tabs.Left = 0

Status.Top = MainText.Height + Toolbar.Height

MainText.Visible = True
OutputText.Visible = True
PadText.Visible = True

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim Result As VbMsgBoxResult
Dim frmX As Form

'Check for IsDirty():
If IsDirty(Input_TextBox) Then 'We check here ONLY for InputBox
    Result = MsgBox("Current 'Input File' has changed, Save?" _
                  , vbYesNoCancel + vbQuestion, "New")
    If Result = vbCancel Then
        Cancel = True 'Stop Unloading
        Exit Sub
    ElseIf Result = vbNo Then
        'nothing here
    ElseIf Result = vbYes Then 'ElseIf is kinda overkill here!
        Call mnuFileSave_Click
    End If
End If

' Save Settings:
AxiomSettings.SaveSettings
go_HTMLTags.SaveToFile RemoveSlash(App.Path) & "\html_tags.ini"

' Unload Classes
Set mo_CStrList = Nothing
Set AxiomSettings = Nothing
Set go_MRU = Nothing
Set go_HTMLTags = Nothing
Set go_PlugIns = Nothing

' Unload Forms, Except "Me" since we're in the Me_Unload event anyway
For Each frmX In Forms
    If frmX.Name <> "frmAxiomMain" Then
        Unload frmX
        Set frmX = Nothing
    End If
Next frmX

End

End Sub
Private Sub MainText_Change()
    
    IsDirty(Input_TextBox) = True

End Sub

Private Sub mnuAbout_Click()
' kinda overkill!
    
    Load frmAboutAxiom
    frmAboutAxiom.Show vbModal
    Unload frmAboutAxiom
    Set frmAboutAxiom = Nothing
    
End Sub

Private Sub mnuAddBA_Click()
Dim sTemp As String

With frmAddBA
    
    .Show vbModal

    If .Canceled = False Then
         sTemp = SelectInput()  '// Text OR SelText
         sTemp = AddToLines(sTemp, .AddBefore, .AddAfter, .IgnoreEmptyLines)
         RedirectOutput sTemp  '// To MainText OR OutputText
    End If

End With

Unload frmAddBA

End Sub

Private Sub mnuAddLineNumbers_Click()
Dim sTemp As String

frmLineNumbers.Show vbModal
If frmLineNumbers.Canceled = False Then
    
    sTemp = SelectInput()  '// Text OR SelText
    If sTemp <> "" Then
        With frmLineNumbers
             sTemp = AddLineNumbers(sTemp, .NumStart, .NumStep, .Delimiter, .NumDigits, .IgnoreEmptyLines)
        End With
        RedirectOutput sTemp  '// To MainText OR OutputText
    End If
End If

Unload frmLineNumbers

End Sub

Private Sub mnuCSS_Click()

Dim sTemp As String
sTemp = SelectInput()

'sTemp = DoCSS(sTemp)
sTemp = RX_ProcessLink(sTemp)

RedirectOutput sTemp

End Sub


Private Sub mnuDeleteChars_Click()

Dim DelType As DeletionType
Dim DelFirst As Long, DelLast As Long
Dim Inclusive As Boolean, MatchCase As Boolean
Dim DelToWhat As String
Dim sTemp  As String
ReDim TempArray(1 To 1) As String
Dim idx As Long

frmDelChars.Show vbModal
    DelType = frmDelChars.WhichDeletionType
    DelFirst = frmDelChars.DelFirst
    DelLast = frmDelChars.DelLast
    Inclusive = frmDelChars.Inclusive
    MatchCase = frmDelChars.MatchCase
    DelToWhat = frmDelChars.DelToWhat
Unload frmDelChars

sTemp = SelectInput()  '// Text OR SelText

If Len(sTemp) = 0 Then Exit Sub

Select Case DelType
Case None 'Cancel was pressed
    Exit Sub

Case DelFirstChars
    If DelFirst = 0 Then MsgBox "Oops": Exit Sub
    Text2Array sTemp, TempArray
    For idx = LBound(TempArray) To UBound(TempArray)
        TempArray(idx) = DelLeft(TempArray(idx), DelFirst)
    Next idx
    sTemp = Array2Text(TempArray)
    
Case DelLastChars
    If DelLast = 0 Then MsgBox "Oops": Exit Sub
    Text2Array sTemp, TempArray
    For idx = LBound(TempArray) To UBound(TempArray)
        TempArray(idx) = DelRight(TempArray(idx), DelLast)
    Next idx
    sTemp = Array2Text(TempArray)

Case DelFromStart
    Text2Array sTemp, TempArray
    For idx = LBound(TempArray) To UBound(TempArray)
        TempArray(idx) = DelLeftTo(TempArray(idx), DelToWhat, MatchCase, Inclusive)
    Next idx
    sTemp = Array2Text(TempArray)


Case DelFromEnd
    Text2Array sTemp, TempArray
    For idx = LBound(TempArray) To UBound(TempArray)
        TempArray(idx) = DelRightTo(TempArray(idx), DelToWhat, MatchCase, Inclusive)
    Next idx
    sTemp = Array2Text(TempArray)

End Select


RedirectOutput sTemp  '// To MainText OR OutputText


End Sub

Private Sub mnuDotBreak_Click()
Dim sTemp As String

sTemp = SelectInput()  '// Text OR SelText

sTemp = BreackOnlyAfter(sTemp, "")

RedirectOutput sTemp  '// To MainText OR OutputText

End Sub

Private Sub mnuEdit_Click()

If mo_TheTextBox.SelLength > 0 Then
    mnuEditCut.Enabled = True
    mnuEditCopy.Enabled = True
    mnuEditDelete.Enabled = True
Else
    mnuEditCut.Enabled = False
    mnuEditCopy.Enabled = False
    mnuEditDelete.Enabled = False
End If

If mo_TheTextBox.Text <> "" Then
    mnuEditSelectAll.Enabled = True
Else
    mnuEditSelectAll.Enabled = False
End If

If Clipboard.GetFormat(vbCFText) Then  'there is text in the clipboard
    mnuEditPaste.Enabled = True
Else
    mnuEditPaste.Enabled = False
End If

' can SWAP only if there is a selection and there is text in Clipboard
' can APPEND only if there is a selection and there is text in Clipboard
If (mo_TheTextBox.SelLength > 0) And Clipboard.GetFormat(vbCFText) Then
    mnuEditSwap.Enabled = True
    mnuEditAppend.Enabled = True
Else
    mnuEditSwap.Enabled = False
    mnuEditAppend.Enabled = False
End If

' Can Undo?
If CanUndo(mo_TheTextBox) Then
    mnuEditUndo.Enabled = True
Else
    mnuEditUndo.Enabled = False
End If

End Sub

Private Sub mnuEditAppend_Click()
Dim sTemp As String

sTemp = Clipboard.GetText(vbCFText)
sTemp = sTemp & mo_TheTextBox.SelText
Clipboard.SetText sTemp, vbCFText

End Sub

Private Sub mnuEditCopy_Click()

If mnuEditCopy.Enabled Then
    Clipboard.SetText mo_TheTextBox.SelText
End If

End Sub

Private Sub mnuEditCut_Click()

If mnuEditCut.Enabled Then
    Clipboard.SetText mo_TheTextBox.SelText
    mo_TheTextBox.SelText = ""
End If


End Sub


Private Sub mnuEditDelete_Click()

If mnuEditDelete.Enabled Then
    mo_TheTextBox.SelText = ""
End If

End Sub

Private Sub mnuEditPaste_Click()

If mnuEditPaste.Enabled Then
    mo_TheTextBox.SelText = Clipboard.GetText(vbCFText)
End If

End Sub

Private Sub mnuEditSelectAll_Click()

mo_TheTextBox.SelStart = 0
mo_TheTextBox.SelLength = Len(mo_TheTextBox.Text)

End Sub

Private Sub mnuEditSwap_Click()
Dim sSwap As String

sSwap = Clipboard.GetText()
Clipboard.SetText mo_TheTextBox.SelText
mo_TheTextBox.SelText = sSwap


End Sub

Private Sub mnuEditUndo_Click()

If mnuEditUndo.Enabled Then
    PerformUndo mo_TheTextBox
End If

End Sub

Private Sub mnuExtractHREFs_Click()
Dim sTemp As String
sTemp = SelectInput()

sTemp = RX_ExtractHREFs(sTemp)

RedirectOutput sTemp

End Sub

Private Sub mnuFile_Click()

'Only if we have a filename and the file has changed
If ((CurrentFile() <> "") And IsDirty(Input_TextBox)) Then
    mnuFileRevert.Enabled = True
Else
    mnuFileRevert.Enabled = False
End If


End Sub

Private Sub mnuFileExit_Click()

Unload Me

End Sub

Private Sub mnuFileInsert_Click()

Dim sFilter As String
Dim TextFile As New CTextFile

sFilter = "All Files|*.*|Text Files|*.txt|HTML Files|*.htm;*.html"
cdlg.Filter = sFilter
cdlg.DialogTitle = "Insert"
cdlg.FileName = ""
cdlg.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly

On Error Resume Next
    cdlg.ShowOpen
If Err Then Err = 0: Exit Sub
On Error GoTo 0

TextFile.FileOpen cdlg.FileName, OpenForInput
mo_TheTextBox.SelText = TextFile.ReadAll
TextFile.FileClose

Set TextFile = Nothing 'overkill

End Sub
Private Sub mnuFileNew_Click()
Dim Result As VbMsgBoxResult

If IsDirty(Input_TextBox) Then 'We check here ONLY for InputBox
            Result = MsgBox("Current 'Input File' has changed, Save?" _
                          , vbYesNoCancel + vbQuestion, "New")
            If Result = vbCancel Then
                Exit Sub
            ElseIf Result = vbNo Then
                'nothing here
            ElseIf Result = vbYes Then 'ElseIf is kinda overkill here!
                Call mnuFileSave_Click
            End If
End If



' Set Public Properties:
CurrentDir() = ""
CurrentFile() = ""
MainText.FileName = ""
MainText.Text = "" 'here IsDirty() will be set to True, so...
IsDirty(Input_TextBox) = False
Set mo_TheTextBox = MainText
Tabs.Tabs.Item("Input").Selected = True
Me.Caption = "Axiom"

' SHOULD WE???
'OutputText.Text = ""

End Sub

Private Sub mnuFileOpen_Click()
Dim sFilter As String
Dim Result As VbMsgBoxResult

If IsDirty(Input_TextBox) Then 'We check here ONLY for InputBox
            Result = MsgBox("Current 'Input File' has changed, Save?" _
                          , vbYesNoCancel + vbQuestion, "New")
            If Result = vbCancel Then
                Exit Sub
            ElseIf Result = vbNo Then
                'nothing here
            ElseIf Result = vbYes Then 'ElseIf is kinda overkill here!
                Call mnuFileSave_Click
            End If
End If


sFilter = "All Files|*.*|Text Files|*.txt|HTML Files|*.htm;*.html"
cdlg.Filter = sFilter
cdlg.DialogTitle = "Open"
cdlg.FileName = ""
cdlg.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly

On Error Resume Next
    cdlg.ShowOpen
If Err Then Err = 0: Exit Sub
On Error GoTo 0

MainText.LoadFile cdlg.FileName, rtfText
Tabs.Tabs.Item(1).Selected = True
Me.Caption = cdlg.FileTitle & " - Axiom"

' Set Public Properties:
go_MRU.Add cdlg.FileName
UpdateMRU
CurrentDir() = cdlg.FileName   ' will automatically extract dir-name
CurrentFile() = cdlg.FileName
IsDirty(Input_TextBox) = False

End Sub

Private Sub mnuFileRevert_Click()
Dim Result As VbMsgBoxResult

Result = MsgBox("Loose All Changes to 'Input File' Since Last Save?" _
              , vbOKCancel + vbQuestion + vbDefaultButton2, "Revert")

If Result = vbCancel Then
    'do nothing
ElseIf Result = vbOK Then
    MainText.LoadFile CurrentFile, rtfText
    Tabs.Tabs.Item(1).Selected = True
    ' Set Public Properties:
    IsDirty(Input_TextBox) = False
End If

End Sub


Private Sub mnuFileSave_Click()
On Error GoTo Err_SaveFile

If WhichBox() = Input_TextBox Then
        If CurrentFile <> "" Then
            MainText.SaveFile CurrentFile, rtfText
            IsDirty(Input_TextBox) = False
        Else
                cdlg.FileName = ""
                cdlg.Filter = "All Files (*.*)|*.*"
                cdlg.DialogTitle = "Save As"
                cdlg.flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
                
                On Error Resume Next
                    cdlg.ShowSave
                If Err Then Exit Sub
                
                If Len(MainText.Text) > 0 Then
                        MainText.SaveFile cdlg.FileName, rtfText
                        CurrentDir() = cdlg.FileName
                        CurrentFile() = cdlg.FileName
                        IsDirty(Input_TextBox) = False
                Else
                        MsgBox "Nothing to Save", vbExclamation, "Oops"
                End If
          End If
Else
    Call mnuFileSaveAs_Click 'OUTPUT and PAD can only "Save As"
End If

Exit Sub
Err_SaveFile:
    MsgBox "ERROR: " & Err.Description, vbCritical, "Error"
    Err = 0
End Sub

Private Sub mnuFileSaveAs_Click()

cdlg.FileName = ""
cdlg.Filter = "All Files (*.*)|*.*"
cdlg.DialogTitle = "Save As"
cdlg.flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly

On Error Resume Next
    cdlg.ShowSave
If Err Then Exit Sub

If Len(mo_TheTextBox.Text) > 0 Then
    mo_TheTextBox.SaveFile cdlg.FileName, rtfText
    IsDirty(WhichBox) = False
    If WhichBox = Input_TextBox Then
        'change only when saving "Input" Textbox
        CurrentFile = cdlg.FileName
        CurrentDir = cdlg.FileName
    End If
Else
    MsgBox "Nothing to Save", vbExclamation, "Oops"
End If

End Sub

Private Sub mnuFind_Click()

With frmFindReplace
    Set .OwnerForm = Me
    .Show

End With

End Sub

Private Sub mnuFindNext_Click()

If gs_FindWhat = "" Then
    mnuFind_Click
Else
    If gl_Pos = 0 Then gl_Pos = mo_TheTextBox.SelStart + mo_TheTextBox.SelLength
    gl_Pos = mo_TheTextBox.Find(gs_FindWhat, gl_Pos, Len(mo_TheTextBox.Text), gl_Options) + 1
    If gl_Pos <= 0 Then gl_Pos = 1
End If

End Sub
Private Sub mnuFromUnicode_Click()
Dim sTemp As String

sTemp = SelectInput()  '// Text OR SelText

sTemp = StrConv(sTemp, vbFromUnicode)

RedirectOutput sTemp  '// To MainText OR OutputText

End Sub

Private Sub mnuHTML_RemovePath_Click()
Dim sTemp As String

With frmRemovePath
    .Show vbModal
    If .Canceled = False Then
            sTemp = SelectInput()
            
            If .RemoveBackground Then
                sTemp = RX_RemoveTagAttrPath(sTemp, "BACKGROUND", .RemoveLocalOnly)
            End If
            If .RemoveHref Then
            sTemp = RX_RemoveTagAttrPath(sTemp, "HREF", .RemoveLocalOnly)
            End If
            If .RemoveSrc Then
                sTemp = RX_RemoveTagAttrPath(sTemp, "SRC", .RemoveLocalOnly)
            End If
            RedirectOutput sTemp
    End If

End With


End Sub

Private Sub mnuHTML_ValidateImg_Click()

Dim sTemp As String
sTemp = SelectInput()

sTemp = RX_ValidateImageTags(sTemp)

RedirectOutput sTemp


End Sub

Private Sub mnuHTMLAddBR_Click()

Dim sTemp As String

sTemp = SelectInput()  '// Text OR SelText

sTemp = RX_AddBR(sTemp)

RedirectOutput sTemp  '// To MainText OR OutputText
        
       

End Sub

Private Sub mnuHTMLize_Click()

Dim sTemp As String

QHTMLizer.Show vbModal
If QHTMLizer.Canceled = False Then
    
    sTemp = SelectInput()  '// Text OR SelText
    
    If sTemp <> "" Then
        With QHTMLizer
        sTemp = HTMLize(sTemp, .PageTitle, .PicturePath, .PageBackColor, _
                        .TextFontName, .TextColor, .TextSize, _
                        .CopyPicture, .BackScroll, .TextBold, .PreserveSpaces, .KeepHTTP, .Target)
        sTemp = DoEMails(sTemp)
        End With
        
        RedirectOutput sTemp  '// To MainText OR OutputText
    End If
    
End If
Unload QHTMLizer

End Sub



Private Sub mnuIn_to_Out_Click()

    'Move input to output
    OutputText.Text = MainText.Text
    MainText.Text = ""
    Tabs.Tabs.Item(2).Selected = True

End Sub

Private Sub mnuIn_to_Pad_Click()

    'Append Output to Pad
    PadText.SelText = MainText.Text
    Tabs.Tabs.Item("Pad").Selected = True


End Sub

Private Sub mnuInsertString_Click()
Dim InsertWhat As String
Dim InsertPos As Long
Dim TempText As String
Dim IgnoreEmptyLines  As Boolean
Dim idx As Long
ReDim TempArray(1 To 1) As String

frmInsertString.Show vbModal
    
    If frmInsertString.Canceled Then Exit Sub
    
    InsertWhat = frmInsertString.InsertWhat
    InsertPos = frmInsertString.InsertPos
    IgnoreEmptyLines = frmInsertString.IgnoreEmptyLines
Unload frmInsertString


TempText = SelectInput()

If TempText = "" Then
    Beep
    Exit Sub
End If

InsertWhat = stringf(InsertWhat)     ' handle special chars

Text2Array TempText, TempArray
For idx = LBound(TempArray) To UBound(TempArray)
    If TempArray(idx) = "" And IgnoreEmptyLines Then
        'do nothing
    Else
        'Insert...
        TempArray(idx) = InsertString(TempArray(idx), InsertWhat, InsertPos)
    End If
Next idx
TempText = Array2Text(TempArray)

RedirectOutput TempText


End Sub

Private Sub mnuLinesCompact_Click()

Dim sTemp As String

sTemp = SelectInput()

'sTemp = CompactBlankLines(sTemp)

sTemp = RX_CompactBlankLines(sTemp)

RedirectOutput sTemp


End Sub

Private Sub mnuLocalClip_Click()

mnuLocalPaste_Click

End Sub

Private Sub mnuLocalCopy_Click()

Dim Sel As String

    
    Sel = mo_TheTextBox.SelText
    If Sel <> "" Then
        frmLocalClipboard.lstLocal.AddItem DoEllipses(Sel, 60)
        mo_CStrList.AddItem Sel
    End If
    

End Sub

Private Sub mnuLocalCut_Click()
Dim Sel As String

    
    Sel = mo_TheTextBox.SelText
    If Sel <> "" Then
        frmLocalClipboard.lstLocal.AddItem DoEllipses(Sel, 60)
        mo_TheTextBox.SelText = ""
        mo_CStrList.AddItem Sel
    End If
    

End Sub


Private Sub mnuLocalPaste_Click()
Dim Sel As String
ReDim sArray(0 To 1)
Dim idx As Long


With frmLocalClipboard
    .Show vbModal
    
    Select Case .Operation
        
        Case PasteItem
            If .lstLocal.ListCount >= 1 Then
                Sel = mo_CStrList.Item(.lstLocal.ListIndex)
                mo_TheTextBox.SelText = Sel
            End If
        Case PasteAll
            If .lstLocal.ListCount >= 1 Then
                Sel = mo_CStrList.Text
                mo_TheTextBox.SelText = Sel
            End If
    Case Else
            ' do nothing '
    End Select

End With

End Sub

Private Sub mnuLowerCase_Click()

Dim sTemp As String

sTemp = SelectInput()  '// Text OR SelText

sTemp = LCase(sTemp)

RedirectOutput sTemp  '// To MainText OR OutputText
        


End Sub

Private Sub mnuMaxLineWidth_Click()

Dim iMax As Integer
Dim sTemp As String


iMax = Val(InputBox("Max Line Width", "Axiom"))
If iMax <> 0 Then
    
    sTemp = SelectInput()  '// Text OR SelText
    sTemp = SetTextMaxWidth(sTemp, iMax)
    RedirectOutput sTemp  '// To MainText OR OutputText

End If


End Sub


Private Sub mnuMRU_Click(Index As Integer)
Dim sFileName As String
Dim Result As VbMsgBoxResult

If IsDirty(Input_TextBox) Then 'We check here ONLY for InputBox
            Result = MsgBox("Current 'Input File' has changed, Save?" _
                          , vbYesNoCancel + vbQuestion, "New")
            If Result = vbCancel Then
                Exit Sub
            ElseIf Result = vbNo Then
                'nothing here
            ElseIf Result = vbYes Then 'ElseIf is kinda overkill here!
                Call mnuFileSave_Click
            End If
End If


sFileName = Right$(mnuMRU(Index).Caption, Len(mnuMRU(Index).Caption) - 3)
On Error Resume Next
    MainText.LoadFile sFileName, rtfText
    If Err Then
        Err = 0
        MsgBox "Cannot Find  " & sFileName, vbCritical, "Oops"
        Exit Sub
    End If
On Error GoTo 0

Tabs.Tabs.Item(1).Selected = True

' Set Public Properties:
go_MRU.Add sFileName
UpdateMRU
CurrentDir() = sFileName   ' will automatically extract dir-name
CurrentFile() = sFileName
IsDirty(Input_TextBox) = False

End Sub

Private Sub mnuOptions_Click()

frmAxiomOptions.Show vbModal

AxiomSettings.ApplySettings



End Sub

Private Sub mnuOut_to_In_Click()
'Move Output to Input
        MainText.Text = OutputText.Text
        OutputText.Text = ""
        Tabs.Tabs.Item(1).Selected = True
End Sub

Private Sub mnuOut_to_Pad_Click()

    'Append Output to Pad
    PadText.SelText = OutputText.Text
    Tabs.Tabs.Item(3).Selected = True


End Sub

Private Sub mnuPlugInDLL_Click()
Dim sFilter As String
Dim Result As VbMsgBoxResult

sFilter = "ActiveX DLL|*.dll|All Files|*.*"
cdlg.Filter = sFilter
cdlg.DialogTitle = "Select ActiveX DLL"
cdlg.FileName = ""
cdlg.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly

On Error Resume Next
    cdlg.ShowOpen
If Err Then Err = 0: Exit Sub

DoEvents
ExecWait RemoveSlash(App.Path) & "\RegSvr32.exe """ & cdlg.FileName & """"
'MsgBox "If the DLL is

End Sub
Private Sub mnuPlugInX_Click(Index As Integer)
On Error GoTo PlugIn_Error
Dim sTemp As String, x As Object
Dim idx As Long, Argv As Variant, NumArgs As Long, lb As Long
Dim FunctionName As String, Description As String, Args As String

Dim sLabels() As String, sTexts() As String

sTemp = SelectInput()
Set x = CreateObject(go_PlugIns.FunctionClass(Index - 1))
FunctionName = go_PlugIns.FunctionName(Index - 1)
Description = go_PlugIns.FunctionDescription(Index - 1)
Args = go_PlugIns.FunctionArgs(Index - 1)
Argv = Split(Args, ",")
NumArgs = UBound(Argv) - LBound(Argv) + 1
Argv(0) = sTemp
For idx = LBound(Argv) + 1 To UBound(Argv)
    'Argv(idx) = InputBox(Argv(idx), Description)
Next idx

If NumArgs > 1 Then
    ReDim sLabels(0 To NumArgs - 2)
    ReDim sTexts(0 To NumArgs - 2)
    
    For idx = 0 To NumArgs - 2
        sLabels(idx) = Argv(idx + 1)
    Next idx
    frmArgs.NumArgs = NumArgs - 1
    frmArgs.Labels sLabels
    frmArgs.Caption = Description
    frmArgs.Show vbModal
    If frmArgs.Canceled Then
        Set frmArgs = Nothing
        Exit Sub
    End If
    frmArgs.GetVals sTexts
    Set frmArgs = Nothing
    For idx = LBound(sTexts) To UBound(sTexts)
        Argv(idx + 1) = sTexts(idx)
    Next
End If

lb = LBound(Argv)

Select Case NumArgs
    Case 1
       sTemp = CallByName(x, FunctionName, VbMethod, CStr(Argv(lb)))
    Case 2
       sTemp = CallByName(x, FunctionName, VbMethod, CStr(Argv(lb)), CStr(Argv(lb + 1)))
    Case 3
       sTemp = CallByName(x, FunctionName, VbMethod, CStr(Argv(lb)), CStr(Argv(lb + 1)), CStr(Argv(lb + 2)))
    Case 4
       sTemp = CallByName(x, FunctionName, VbMethod, CStr(Argv(lb)), CStr(Argv(lb + 1)), CStr(Argv(lb + 2)), CStr(Argv(lb + 3)))
    Case 5
       sTemp = CallByName(x, FunctionName, VbMethod, CStr(Argv(lb)), CStr(Argv(lb + 1)), CStr(Argv(lb + 2)), CStr(Argv(lb + 3)), CStr(Argv(lb + 4)))
End Select

RedirectOutput sTemp
Set x = Nothing

'ERROR HANDLER
Exit Sub
PlugIn_Error:
Set x = Nothing
MsgBox "Error:  " & Err.Description & vbCrLf & "While Executing:  " & FunctionName & vbCrLf & "Member of:  " & go_PlugIns.FunctionClass(Index - 1), vbCritical, "Oops"
Exit Sub

End Sub

Private Sub mnuRemoveAllTags_Click()
Dim sTemp As String

sTemp = SelectInput()

If Len(sTemp) > 0 Then
    sTemp = RX_RemoveAllTags(sTemp)
    sTemp = Replace(sTemp, "&nbsp;", " ")
    sTemp = Replace(sTemp, "&gt;", ">")
    sTemp = Replace(sTemp, "&lt;", "<")
    sTemp = Replace(sTemp, "&quot;", """")
    'sTemp = TrimSpaces(sTemp, False, True, True)
    RedirectOutput sTemp
End If

End Sub

Private Sub mnuRemoveBlankLines_Click()

Dim sTemp As String

sTemp = SelectInput()

sTemp = RX_RemoveBlankLines(sTemp)

RedirectOutput sTemp


End Sub
Private Sub mnuRemoveChars_Click()

Dim sTemp As String

sTemp = SelectInput()  '// Text OR SelText

' RemoveNonAlphaNum2 turned out faster than RemoveNonAlphaNum1
' RemoveNonAlphaNum3 turned out MUCH faster than RemoveNonAlphaNum2

' RemoveNonAlphaNum4 turned out about 30% faster than RemoveNonAlphaNum3
' on a 1.77 MB file   32 sec V/S  45 sec
sTemp = RemoveNonAlphaNum4(sTemp, AxiomSettings.CharsToKeep)

RedirectOutput sTemp  '// To MainText OR OutputText

End Sub

Private Sub mnuRemoveHtmlComments_Click()
Dim sTemp As String
sTemp = SelectInput()

sTemp = RX_RemoveCommentTagAndContent(sTemp)

RedirectOutput sTemp

End Sub

Private Sub mnuRemoveSpaces_Click()
Dim sTemp As String

sTemp = SelectInput()  '// Text OR SelText

sTemp = CompactSpaces(sTemp)

RedirectOutput sTemp  '// To MainText OR OutputText

End Sub


Private Sub mnuReverseLines_Click()
Dim sTemp As String

sTemp = SelectInput()  '// Text OR SelText

sTemp = ReverseStr(sTemp, ByLine:=True)

RedirectOutput sTemp  '// To MainText OR OutputText

End Sub

Private Sub mnuReverseText_Click()

Dim sTemp As String

sTemp = SelectInput()  '// Text OR SelText

sTemp = ReverseStr(sTemp, ByLine:=False)

RedirectOutput sTemp  '// To MainText OR OutputText

End Sub

Private Sub mnuSetMaxWord_Click()
Dim iMax As Integer
Dim sTemp As String


iMax = Val(InputBox("Max Line Width", "Axiom"))
If iMax <> 0 Then
    
    sTemp = SelectInput()  '// Text OR SelText
    sTemp = SetTextMaxWidthWords(sTemp, iMax)
    RedirectOutput sTemp  '// To MainText OR OutputText

End If

End Sub

Private Sub mnuSortAscending_Click()
ReDim sArray(1 To 1) As String
Dim sTemp As String

sTemp = SelectInput()

Text2Array sTemp, sArray

StrSort sArray, Ascending:=True, AllLowerCase:=True

sTemp = Array2Text(sArray)

RedirectOutput sTemp


End Sub


Private Sub mnuSortDescending_Click()

ReDim sArray(1 To 1) As String
Dim sTemp As String

sTemp = SelectInput()

Text2Array sTemp, sArray

StrSort sArray, Ascending:=False, AllLowerCase:=True

sTemp = Array2Text(sArray)

RedirectOutput sTemp


End Sub


Private Sub mnuTab2Spc_Click()
    
Dim sTemp As String

sTemp = SelectInput()  '// Text OR SelText

sTemp = Tab2Spaces(sTemp, AxiomSettings.SpacesPerTab)

RedirectOutput sTemp  '// To MainText OR OutputText


End Sub

Private Sub mnuTags_Click()
Dim sTemp As String
Dim tmpSArray() As String
Dim tmpBArray() As Boolean

Dim idx As Long

frmHTMLOps.Show vbModal
If frmHTMLOps.Operation <> Cancel Then
    sTemp = SelectInput()
    tmpSArray = frmHTMLOps.HTMLTags
    tmpBArray = frmHTMLOps.TagIsSingle
    
    Select Case frmHTMLOps.Operation
        
        Case DeleteTagAndContent
             For idx = LBound(tmpSArray) To UBound(tmpSArray)
                 sTemp = RX_RemoveTagWithContents(sTemp, tmpSArray(idx), tmpBArray(idx))
             Next idx
             
        Case DeleteTagKeepContent
             For idx = LBound(tmpSArray) To UBound(tmpSArray)
                sTemp = RX_RemoveOpenCloseTagKeepContent(sTemp, tmpSArray(idx))
             Next idx
             
        Case ExtractTagAndContent
             For idx = LBound(tmpSArray) To UBound(tmpSArray)
                sTemp = RX_ExtractTagWithContents(sTemp, tmpSArray(idx))
             Next idx
    End Select
    
    RedirectOutput sTemp
End If

End Sub

Private Sub mnuText_RXFind_Click()
Dim sTemp As String
Dim sPattern As String, sReplaceWith As String

frmRegExp.Show vbModal
sPattern = frmRegExp.Pattern
sReplaceWith = frmRegExp.ReplaceWith
Unload frmRegExp

If sPattern <> "" Then
    sTemp = SelectInput()  '// Text OR SelText
        
        If sReplaceWith = vbNullChar Then
            sTemp = RX_GenericExtract(sTemp, sPattern)
        Else
            sTemp = RX_GenericReplace(sTemp, sPattern, sReplaceWith)
        End If
    
    RedirectOutput sTemp  '// To MainText OR OutputText
End If

End Sub

Private Sub mnuTextFixNewLine_Click()

Dim sTemp As String

MainText.SelLength = 0 ' to force SelectInput to take the WHOLE text

sTemp = SelectInput()

sTemp = FixNewLineChars(sTemp)

RedirectOutput sTemp

End Sub


Private Sub mnuTextTrigger_Click()
Dim sTemp As String

sTemp = SelectInput()  '// Text OR SelText

sTemp = HandleTextTrigger(sTemp, TriggerChars:="~")

RedirectOutput sTemp  '// To MainText OR OutputText

End Sub

Private Sub mnuTitleCase_Click()

Dim sTemp As String

sTemp = SelectInput()  '// Text OR SelText

sTemp = StrConv(sTemp, vbProperCase)

RedirectOutput sTemp  '// To MainText OR OutputText

End Sub

Private Sub mnuTrimBoth_Click()
Dim sTemp As String

sTemp = SelectInput()

sTemp = TrimSpaces(sTemp, True, True, AxiomSettings.TrimTabsAlso)

RedirectOutput sTemp

End Sub

Private Sub mnuTrimLeft_Click()

Dim sTemp As String

sTemp = SelectInput()

sTemp = TrimSpaces(sTemp, True, False, AxiomSettings.TrimTabsAlso)

RedirectOutput sTemp


End Sub

Private Sub mnuTrimRight_Click()

Dim sTemp As String

sTemp = SelectInput()

sTemp = TrimSpaces(sTemp, False, True, AxiomSettings.TrimTabsAlso)

RedirectOutput sTemp

End Sub



Private Sub mnuUpperCase_Click()
        
Dim sTemp As String

sTemp = SelectInput()  '// Text OR SelText

sTemp = UCase(sTemp)

RedirectOutput sTemp  '// To MainText OR OutputText
        
       

End Sub

Private Sub mo_TheTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
' Overriding the RTFBox built-in shortcuts,
' Required because the RitchTextBox does not pass Key Presses
' to the form!
Static CurrentTab As Integer
Const CONTEXT_BUTTON = 93 ' &H5D

If Shift <> 0 Then     ' Shift or Ctrl is pressed
    Call mnuEdit_Click ' to Enable/Disable Commands
End If

'Trap [CTRL]+[A]   SELECT ALL
If (Shift = vbCtrlMask) And (KeyCode = vbKeyA) Then
    KeyCode = 0
    mnuEditSelectAll_Click
End If

If (Shift = vbCtrlMask) And (KeyCode = vbKeyC) Then
    'Trap [CTRL]+[C]   COPY
    KeyCode = 0
    mnuEditCopy_Click

ElseIf (Shift = vbCtrlMask) And (KeyCode = vbKeyV) Then
    'Trap [CTRL]+[V]   PASTE
    KeyCode = 0
    mnuEditPaste_Click

ElseIf (Shift = vbCtrlMask) And (KeyCode = vbKeyX) Then
    'Trap [CTRL]+[X]   CUT
    KeyCode = 0
    mnuEditCut_Click

ElseIf (Shift = vbCtrlMask) And (KeyCode = vbKeyZ) Then
    'Trap [CTRL]+[Z]   UNDO
    KeyCode = 0
    If CanUndo(mo_TheTextBox) Then PerformUndo mo_TheTextBox

ElseIf (Shift = vbShiftMask) And (KeyCode = vbKeyInsert) Then
    'Trap [SHIFT]+[INS]  LOCAL PASTE
    KeyCode = 0
    mnuLocalPaste_Click

ElseIf (Shift = vbCtrlMask) And (KeyCode = vbKeyInsert) Then
    'Trap [CTRL]+[INS]    LOCAL COPY
    KeyCode = 0
    mnuLocalCopy_Click

ElseIf (Shift = vbShiftMask) And (KeyCode = vbKeyDelete) Then
    'Trap [SHIFT]+[DEL]   LOCAL CUT
    KeyCode = 0
    mnuLocalCut_Click

'   CONTEXT MENU BUTTON  ' value obtained thru test!
ElseIf KeyCode = CONTEXT_BUTTON Then
    PopupMenu mnuEdit, vbRightButton
'   [SHIFT]+[F10]
ElseIf (Shift = vbShiftMask) And (KeyCode = vbKeyF10) Then
    PopupMenu mnuEdit, vbRightButton
    
'   TRAP [CTRL]+[TAB]
ElseIf (Shift = vbCtrlMask) And (KeyCode = vbKeyTab) Then
    KeyCode = 0
    For CurrentTab = 1 To Tabs.Tabs.Count
        If Tabs.Tabs.Item(CurrentTab).Selected = True Then Exit For
    Next CurrentTab
    CurrentTab = CurrentTab + 1
    If CurrentTab > Tabs.Tabs.Count Then CurrentTab = 1
    Tabs.Tabs.Item(CurrentTab).Selected = True
End If

End Sub

Private Sub mo_TheTextBox_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbRightButton Then
    PopupMenu mnuEdit, vbRightButton
End If

End Sub


Private Sub mo_TheTextBox_SelChange()
Dim rtfRow As Long, rtfCol As Long
Dim sTemp As String

Rtf_SelChange mo_TheTextBox, rtfRow, rtfCol
    
'don't know why result is sometimes 0 ???
If (rtfRow >= 0 And rtfCol >= 0) Then
    sTemp = "Line: " & Format(rtfRow) & "  Col: " & _
            Format(rtfCol) & "    Char: " & _
            Format(mo_TheTextBox.SelStart + 1)
Else
    sTemp = ""
End If

Status.Panels.Item(1).Text = sTemp
frmFindReplace.CurPos = mo_TheTextBox.SelStart

gl_Pos = mo_TheTextBox.SelStart + mo_TheTextBox.SelLength

End Sub

Private Sub OutputText_Change()
    
   IsDirty(Output_TextBox) = True
    
End Sub

Private Sub PadText_Change()
    
    IsDirty(Pad_TextBox) = True
    
End Sub




Private Sub Tabs_Click()

Status.Panels.Item(1).Text = "" 'Clear Line,Col info

Select Case Tabs.SelectedItem.Key

Case "Input"
    Set mo_TheTextBox = MainText
    MainText.ZOrder 0
    OutputText.ZOrder 1
    PadText.ZOrder 1
    MainText.SetFocus
Case "Output"
    Set mo_TheTextBox = OutputText
    OutputText.ZOrder 0
    MainText.ZOrder 1
    PadText.ZOrder 1
    OutputText.SetFocus
Case "Pad"
    Set mo_TheTextBox = PadText
    PadText.ZOrder 0
    OutputText.ZOrder 1
    MainText.ZOrder 1
    PadText.SetFocus
Case Else
    'do nothing
End Select

gl_Pos = mo_TheTextBox.SelStart

End Sub


Private Sub Tabs_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbRightButton Then
    PopupMenu Me.mnuMove, 2 Or 4
End If

End Sub



Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case LCase$(Button.Key)
    Case "open"
        mnuFileOpen_Click
    Case "new"
        mnuFileNew_Click
    Case "reveret"
        mnuFileRevert_Click
    
    Case "saveas"
        mnuFileSaveAs_Click
    Case "new"
        mnuFileNew_Click
    Case "insert"
        mnuFileInsert_Click
    Case "move"
         If (AxiomSettings.MultipleTextBoxes = True) And (OutputText.Text <> "") Then
            mnuOut_to_In_Click
         Else
            Beep
         End If
    Case Else

End Select

End Sub



