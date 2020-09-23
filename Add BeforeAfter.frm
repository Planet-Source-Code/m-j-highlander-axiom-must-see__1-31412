VERSION 5.00
Begin VB.Form frmAddBA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Before and/or After Every Line"
   ClientHeight    =   1935
   ClientLeft      =   1800
   ClientTop       =   1935
   ClientWidth     =   6240
   Icon            =   "Add BeforeAfter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin Axiom.FakeButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   3773
      TabIndex        =   6
      Top             =   1380
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      Caption         =   "Cancel"
      ForeColor       =   255
   End
   Begin Axiom.FakeButton cmdOk 
      Default         =   -1  'True
      Height          =   405
      Left            =   893
      TabIndex        =   5
      Top             =   1380
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      Caption         =   "Ok"
      ForeColor       =   12582912
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
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   960
      Value           =   1  'Checked
      Width           =   2760
   End
   Begin VB.TextBox txtAddAfter 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   525
      Width           =   5010
   End
   Begin VB.TextBox txtAddBefore 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   5010
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Add After"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   570
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Add Before"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   165
      Width           =   795
   End
End
Attribute VB_Name = "frmAddBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mb_Canceled As Boolean
Private ms_AddBefore As String
Private ms_AddAfter As String
Private mb_IgnoreEmptyLines As Boolean

Public Property Get IgnoreEmptyLines() As Boolean
       IgnoreEmptyLines = mb_IgnoreEmptyLines
End Property

Public Property Let IgnoreEmptyLines(ByVal bNewValue As Boolean)
       mb_IgnoreEmptyLines = bNewValue
End Property




Public Property Get AddBefore() As String
       AddBefore = ms_AddBefore
End Property

Public Property Let AddBefore(ByVal sNewValue As String)
       ms_AddBefore = sNewValue
End Property


Public Property Get AddAfter() As String
       AddAfter = ms_AddAfter
End Property

Public Property Let AddAfter(ByVal sNewValue As String)
       ms_AddAfter = sNewValue
End Property

Public Property Get Canceled() As Boolean
       Canceled = mb_Canceled
End Property

Public Property Let Canceled(ByVal bNewValue As Boolean)
       mb_Canceled = bNewValue
End Property

Private Sub cmdCancel_Click()

Canceled = True
Me.Hide

End Sub

Private Sub cmdOk_Click()

If txtAddBefore.Text = "" And txtAddAfter.Text = "" Then
    Beep
    Exit Sub
End If

'[ Set Properties ]
Canceled = False
AddBefore = CStr(txtAddBefore.Text)
AddAfter = CStr(txtAddAfter.Text)
IgnoreEmptyLines = CBool(chkIgnoreEmptyLines.Value)

Me.Hide

End Sub

Private Sub Form_Load()

AddBorderToAllTextBoxes Me

End Sub


Private Sub Form_Unload(Cancel As Integer)

Canceled = True

End Sub


