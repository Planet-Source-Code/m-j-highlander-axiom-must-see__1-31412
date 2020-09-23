VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form QHTMLizer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Convert Text to HTML"
   ClientHeight    =   3705
   ClientLeft      =   900
   ClientTop       =   1305
   ClientWidth     =   6555
   Icon            =   "qhtml.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   247
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   ShowInTaskbar   =   0   'False
   Begin Axiom.FakeButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   4920
      TabIndex        =   24
      Top             =   3120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      Caption         =   "Cancel"
      ForeColor       =   255
   End
   Begin Axiom.FakeButton cmdOk 
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   23
      Top             =   3120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      Caption         =   "Ok"
      ForeColor       =   16711680
   End
   Begin VB.ComboBox cboFontSize 
      Height          =   315
      Left            =   4125
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1800
      Width           =   645
   End
   Begin Axiom.ScalablePic RichTextBox1 
      Height          =   915
      Index           =   0
      Left            =   480
      TabIndex        =   21
      Top             =   4380
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1614
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin Axiom.CoolButton cmdBrowse 
      Height          =   330
      Left            =   6060
      TabIndex        =   18
      Top             =   480
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      Caption         =   ""
      ShowFocusRect   =   0   'False
      BackPicture     =   "qhtml.frx":000C
   End
   Begin Axiom.ColorButton clrTextColor 
      Height          =   285
      Left            =   2055
      TabIndex        =   17
      Top             =   1755
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   503
   End
   Begin Axiom.ColorButton clrBGColor 
      Height          =   285
      Left            =   2055
      TabIndex        =   16
      Top             =   1395
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   503
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   660
      Top             =   3900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   " Web Links "
      Height          =   1440
      Left            =   120
      TabIndex        =   13
      Top             =   2175
      Width           =   3045
      Begin VB.TextBox txtTarget 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   19
         Text            =   "_blank"
         Top             =   1035
         Width           =   1740
      End
      Begin VB.OptionButton optNoHttp 
         Caption         =   "www.something.com"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   15
         Top             =   630
         Width           =   2535
      End
      Begin VB.OptionButton optHasHttp 
         Caption         =   "http://www.something.com"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   14
         Top             =   315
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Target"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   20
         Top             =   1080
         Width           =   555
      End
   End
   Begin VB.CheckBox chkBGScrollable 
      Caption         =   "Background Scrollable"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3585
      TabIndex        =   12
      Top             =   870
      Width           =   2775
   End
   Begin VB.CheckBox chkKeepSpaces 
      Caption         =   "Preserve spaces"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3540
      TabIndex        =   11
      Top             =   2295
      Width           =   2775
   End
   Begin VB.TextBox txtFontName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4095
      TabIndex        =   8
      Text            =   "Verdana,Tahoma,Arial"
      Top             =   1425
      Width           =   2370
   End
   Begin VB.CheckBox chkFontBold 
      Caption         =   "Bold"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5250
      TabIndex        =   10
      Top             =   1830
      Width           =   750
   End
   Begin VB.CheckBox chkFullPath 
      Caption         =   "Use full path"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   870
      Width           =   2340
   End
   Begin VB.TextBox txtBGPic 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1965
      TabIndex        =   3
      Top             =   510
      Width           =   4050
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   645
      TabIndex        =   1
      Top             =   120
      Width           =   5745
   End
   Begin VB.Label lblFontName 
      AutoSize        =   -1  'True
      Caption         =   "Font Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3165
      TabIndex        =   7
      Top             =   1455
      Width           =   915
   End
   Begin VB.Label lblFontSize 
      AutoSize        =   -1  'True
      Caption         =   "Font Size"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3180
      TabIndex        =   9
      Top             =   1845
      Width           =   780
   End
   Begin VB.Label lblTextColor 
      AutoSize        =   -1  'True
      Caption         =   "Text Color"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   6
      Top             =   1815
      Width           =   900
   End
   Begin VB.Label lblBGPic 
      AutoSize        =   -1  'True
      Caption         =   "Background Picture"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   525
      Width           =   1665
   End
   Begin VB.Label lblBGColor 
      AutoSize        =   -1  'True
      Caption         =   "Background Color"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   420
      TabIndex        =   5
      Top             =   1455
      Width           =   1545
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   360
   End
End
Attribute VB_Name = "QHTMLizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ms_PageTitle As String
Private ms_PicturePath As String
Private mb_CopyPicture As Boolean
Private mb_BackScroll As Boolean
Private ms_PageBackColor As String
Private ms_TextColor As String
Private ms_TextFontName As String
Private ms_TextSize As String
Private mb_TextBold As Boolean
Private mb_PreserveSpaces As Boolean
Private mb_Canceled As Boolean
Private mb_KeepHTTP As Boolean
Private ms_Target As String

Public Property Get Target() As String
       Target = ms_Target
End Property

Public Property Let Target(ByVal sNewValue As String)
       ms_Target = sNewValue
End Property

Public Property Get KeepHTTP() As Boolean
       KeepHTTP = mb_KeepHTTP
End Property

Public Property Let KeepHTTP(ByVal bNewValue As Boolean)
       mb_KeepHTTP = bNewValue
End Property

Public Property Get Canceled() As Boolean
       Canceled = mb_Canceled
End Property

Public Property Let Canceled(ByVal bNewValue As Boolean)
       mb_Canceled = bNewValue
End Property

Public Property Get PreserveSpaces() As Boolean
       PreserveSpaces = mb_PreserveSpaces
End Property

Public Property Let PreserveSpaces(ByVal bNewValue As Boolean)
       mb_PreserveSpaces = bNewValue
End Property

Public Property Get TextBold() As Boolean
       TextBold = mb_TextBold
End Property

Public Property Let TextBold(ByVal bNewValue As Boolean)
       mb_TextBold = bNewValue
End Property

Public Property Get TextSize() As String
       TextSize = ms_TextSize
End Property

Public Property Let TextSize(ByVal sNewValue As String)
       ms_TextSize = sNewValue
End Property

Public Property Get TextFontName() As String
       TextFontName = ms_TextFontName
End Property

Public Property Let TextFontName(ByVal sNewValue As String)
       ms_TextFontName = sNewValue
End Property

Public Property Get TextColor() As String
       TextColor = ms_TextColor
End Property

Public Property Let TextColor(ByVal sNewValue As String)
       ms_TextColor = sNewValue
End Property

Public Property Get PageBackColor() As String
       PageBackColor = ms_PageBackColor
End Property

Public Property Let PageBackColor(ByVal sNewValue As String)
       ms_PageBackColor = sNewValue
End Property

Public Property Get BackScroll() As Boolean
       BackScroll = mb_BackScroll
End Property

Public Property Let BackScroll(ByVal bNewValue As Boolean)
       mb_BackScroll = bNewValue
End Property

Public Property Get CopyPicture() As Boolean
       CopyPicture = mb_CopyPicture
End Property

Public Property Let CopyPicture(ByVal bNewValue As Boolean)
       mb_CopyPicture = bNewValue
End Property

Public Property Get PicturePath() As String
       PicturePath = ms_PicturePath
End Property

Public Property Let PicturePath(ByVal sNewValue As String)
       ms_PicturePath = sNewValue
End Property

Public Property Get PageTitle() As String
       PageTitle = ms_PageTitle
End Property

Public Property Let PageTitle(ByVal sNewValue As String)
       ms_PageTitle = sNewValue
End Property


Private Sub cmdBrowse_Click()
'/////// OLD CODE 1
'On Error Resume Next
'cdlg.FileName = ""
'cdlg.Filter = "Pictures (*.jpg;*.gif;*.png;*.bmp)|*.jpg;*.gif;*.png;*.bmp|All Files (*.*)|*.*"
'
'cdlg.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
'cdlg.ShowOpen
'If Err Then Exit Sub  'Cancel was selected
'txtBGPic.Text = cdlg.FileName
'''''''''''''''''''''''''''''''''''''''''''''''
'/////// OLD CODE 2
'Load frmSelectPic
'frmSelectPic.Treefolder1.SetPath ExtractDirName(txtBGPic.Text)
'frmSelectPic.Show vbModal
'If frmSelectPic.ImageFilePath <> "" Then
'        txtBGPic.Text = frmSelectPic.ImageFilePath
'End If
'Unload frmSelectPic
'''''''''''''''''''''''''''''''''''''''''''''''
'/////// NEW CODE
Dim cdlg As New CdlgEx
  cdlg.InitDir = ExtractDirName(txtBGPic.Text)
  cdlg.hOwner = Me.hWnd
  cdlg.Left = 100
  cdlg.Top = 100
  cdlg.OKText = "Open"
  cdlg.CancelText = "Cancel"
  cdlg.HelpText = "No Help"
  cdlg.DialogTitle = "Select Image File"
  ' CDlg.CancelError = True
  cdlg.Filter = "Picture Files|*.bmp;*.gif;*.jpg;*.ico;*.wmf|All files|*.*"
  cdlg.flags = &H4 Or &H1000
  Load RichTextBox1(1)

Set rtb = RichTextBox1(1)
cdlg.ShowOpen
If RichTextBox1.Count > 1 Then Unload RichTextBox1(1)

If cdlg.Filename <> "" Then txtBGPic.Text = cdlg.Filename
Set cdlg = Nothing

End Sub

Private Sub cmdCancel_Click()

Canceled = True
Me.Hide
End Sub

Private Sub cmdOk_Click()

'// Set Properties //

If chkFullPath.Value = vbChecked Then
    PicturePath = CStr(txtBGPic.Text)
Else
    PicturePath = ExtractFileName(CStr(txtBGPic.Text))
End If
Canceled = False
PageTitle = CStr(txtTitle.Text)
CopyPicture = CBool(chkFullPath.Value)
BackScroll = CBool(chkBGScrollable.Value)
TextFontName = CStr(txtFontName.Text)
TextSize = CStr(cboFontSize.Text)
TextBold = CBool(chkFontBold)
PreserveSpaces = CBool(chkKeepSpaces.Value)
TextColor = ColorToHex(clrTextColor.Color)
PageBackColor = ColorToHex(clrBGColor.Color)
KeepHTTP = CBool(optHasHttp.Value)
Target = CStr(txtTarget.Text)

Me.Hide

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()



clrBGColor.Color = vbWhite
clrBGColor.hwndOwner = Me.hWnd
clrTextColor.Color = vbBlack
clrBGColor.hwndOwner = Me.hWnd

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
        cmdCancel_Click
End If

If KeyAscii = 13 Then
        cmdOk_Click
End If

End Sub


Private Sub Form_Load()
Dim idx As Long

AddBorderToAllTextBoxes Me

For idx = 1 To 7
    cboFontSize.AddItem CStr(idx)
Next idx

cboFontSize.ListIndex = 1

End Sub


Private Sub Form_Unload(Cancel As Integer)
Canceled = True
End Sub

