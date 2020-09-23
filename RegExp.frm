VERSION 5.00
Begin VB.Form frmRegExp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Regular Expression  Extract / Replace"
   ClientHeight    =   1620
   ClientLeft      =   1020
   ClientTop       =   2295
   ClientWidth     =   6945
   Icon            =   "RegExp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtReplace 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   600
      Width           =   4935
   End
   Begin VB.CheckBox chkReplace 
      Caption         =   "&Replace With"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin Axiom.FakeButton cmdOk 
      Default         =   -1  'True
      Height          =   345
      Left            =   2940
      TabIndex        =   2
      Top             =   1200
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
      Caption         =   "Ok"
      ForeColor       =   16711680
   End
   Begin VB.TextBox txtPattern 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin Axiom.FakeButton FakeButton1 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4380
      TabIndex        =   3
      Top             =   1200
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
      Caption         =   "Cancel"
      ForeColor       =   255
   End
   Begin Axiom.FakeButton FakeButton2 
      Height          =   345
      Left            =   5700
      TabIndex        =   4
      Top             =   1200
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
      Caption         =   "Help"
      ForeColor       =   8421376
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Search Pattern"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1065
   End
End
Attribute VB_Name = "frmRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ms_Pattern As String
Private bPressed_Help As Boolean
Private ms_ReplaceWith As String

Public Property Get ReplaceWith() As String
       ReplaceWith = ms_ReplaceWith
End Property

Public Property Let ReplaceWith(ByVal sNewValue As String)
       ms_ReplaceWith = sNewValue
End Property

Public Property Get Pattern() As String
       Pattern = ms_Pattern
End Property

Public Property Let Pattern(ByVal sNewValue As String)
       ms_Pattern = sNewValue
End Property

Private Sub chkReplace_Click()

If chkReplace.Value = vbChecked Then
    txtReplace.SetFocus
End If

End Sub

Private Sub cmdOk_Click()

Pattern = txtPattern.Text
If chkReplace.Value = vbChecked Then
    ReplaceWith = txtReplace.Text
Else
    ReplaceWith = vbNullChar
End If

Me.Hide

End Sub

Private Sub FakeButton1_Click()

Pattern = ""
ReplaceWith = ""

Me.Hide

End Sub

Private Sub FakeButton2_Click()

HHelp_Show RemoveSlash(App.Path) & "\regexp.chm", "RegExp Pattern.html"
bPressed_Help = True

End Sub

Private Sub Form_Load()

bPressed_Help = False

AddBorderToAllTextBoxes Me

End Sub


Private Sub Form_Unload(Cancel As Integer)

If bPressed_Help Then
    HHelp_Close
End If

End Sub


Private Sub txtReplace_Change()

'If txtReplace.Text <> "" Then
'    chkReplace.Value = vbChecked
'End If

End Sub


