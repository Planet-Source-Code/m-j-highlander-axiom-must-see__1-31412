VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInsertString 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert String"
   ClientHeight    =   2220
   ClientLeft      =   1545
   ClientTop       =   2040
   ClientWidth     =   5925
   Icon            =   "InsertString.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   148
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   395
   ShowInTaskbar   =   0   'False
   Begin Axiom.FakeButton cmdOk 
      Default         =   -1  'True
      Height          =   450
      Left            =   795
      TabIndex        =   7
      Top             =   1680
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   794
      Caption         =   "Ok"
      ForeColor       =   16711680
   End
   Begin VB.CheckBox chkIgnoreEmptyLines 
      Caption         =   "&Ignore empty lines"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1395
      TabIndex        =   6
      Top             =   1125
      Value           =   1  'Checked
      Width           =   3705
   End
   Begin Axiom.CoolButton cmdSpecial 
      Height          =   315
      Left            =   4860
      TabIndex        =   5
      Top             =   180
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      Caption         =   ">"
      Font.Charset    =   0
      Font.Name       =   "Verdana"
      Font.size       =   6.75
   End
   Begin VB.TextBox txtInsertPos 
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
      Height          =   320
      Left            =   1590
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      Top             =   630
      Width           =   735
   End
   Begin VB.TextBox txtInsertWhat 
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
      Height          =   320
      Left            =   810
      TabIndex        =   0
      Top             =   180
      Width           =   3975
   End
   Begin MSComCtl2.UpDown upInsertPos 
      Height          =   330
      Left            =   2415
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   630
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      OrigLeft        =   3270
      OrigTop         =   2250
      OrigRight       =   3510
      OrigBottom      =   2580
      Max             =   999
      Enabled         =   -1  'True
   End
   Begin Axiom.FakeButton cmdCancel 
      Height          =   450
      Left            =   3615
      TabIndex        =   8
      Top             =   1680
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   794
      Caption         =   "Cancel"
      ForeColor       =   255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "at char position"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   675
      Width           =   1320
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   225
      Width           =   510
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "mnuHidden"
      Visible         =   0   'False
      Begin VB.Menu mnuPara 
         Caption         =   "Paragrap"
      End
      Begin VB.Menu mnuTab 
         Caption         =   "Tab"
      End
   End
End
Attribute VB_Name = "frmInsertString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mb_Canceled As Boolean
Private mb_IgnoreEmptyLines As Boolean
Public Property Get IgnoreEmptyLines() As Boolean
       IgnoreEmptyLines = mb_IgnoreEmptyLines
End Property
Public Property Let IgnoreEmptyLines(ByVal bNewValue As Boolean)
       mb_IgnoreEmptyLines = bNewValue
End Property




Private Sub cmdCancel_Click()

Me.Canceled = True
Me.Hide

End Sub

Private Sub cmdOk_Click()

IgnoreEmptyLines = CBool(chkIgnoreEmptyLines.Value)
Me.Canceled = False
Me.Hide

End Sub


Private Sub cmdSpecial_MouseDown()
PopupMenu mnuHidden, 2, cmdSpecial.Left, cmdSpecial.Top + cmdSpecial.Height
txtInsertWhat.SetFocus

End Sub


Private Sub Form_Load()

AddBorderToAllTextBoxes Me
FlatAllBtns Me

End Sub





Public Property Get InsertPos() As Integer
    InsertPos = Val(txtInsertPos.Text)
End Property


Public Property Get InsertWhat() As String
    InsertWhat = CStr(txtInsertWhat.Text)
End Property





Private Sub mnuPara_Click()
    
    txtInsertWhat.SelText = "\n"

End Sub

Private Sub mnuTab_Click()
txtInsertWhat.SelText = "\t"
End Sub


Private Sub txtInsertPos_Change()
Dim v As Integer

v = Val(txtInsertPos.Text)
If v = 0 Then
    v = 1
    txtInsertPos = "1"
End If

upInsertPos.Value = v

End Sub


Private Sub upInsertPos_Change()
txtInsertPos.Text = upInsertPos.Value
End Sub



Public Property Get Canceled() As Boolean
    Canceled = mb_Canceled
End Property

Public Property Let Canceled(ByVal bNewValue As Boolean)
    mb_Canceled = bNewValue
End Property
