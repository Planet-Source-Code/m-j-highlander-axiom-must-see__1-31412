VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLineNumbers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Line Numering Options"
   ClientHeight    =   3210
   ClientLeft      =   2415
   ClientTop       =   1935
   ClientWidth     =   3855
   Icon            =   "LineNumbers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin Axiom.FakeButton cmdOk 
      Default         =   -1  'True
      Height          =   420
      Left            =   420
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
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
      Height          =   240
      Left            =   720
      TabIndex        =   9
      Top             =   1935
      Value           =   1  'Checked
      Width           =   2760
   End
   Begin VB.TextBox txtDelimiter 
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
      Height          =   330
      Left            =   2070
      TabIndex        =   8
      Text            =   "- "
      Top             =   1395
      Width           =   1185
   End
   Begin VB.TextBox txtDigits 
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
      Height          =   320
      Left            =   2070
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "1"
      Top             =   990
      Width           =   915
   End
   Begin VB.TextBox txtNumStep 
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
      Height          =   330
      Left            =   2070
      TabIndex        =   3
      Text            =   "1"
      Top             =   585
      Width           =   915
   End
   Begin VB.TextBox txtNumStart 
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
      Height          =   330
      Left            =   2070
      TabIndex        =   2
      Text            =   "1"
      Top             =   180
      Width           =   915
   End
   Begin MSComCtl2.UpDown upNumDigits 
      Height          =   330
      Left            =   3015
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   990
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Value           =   1
      OrigLeft        =   3270
      OrigTop         =   2250
      OrigRight       =   3510
      OrigBottom      =   2580
      Max             =   999999999
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown upStart 
      Height          =   330
      Left            =   3015
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   180
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Value           =   1
      OrigLeft        =   3270
      OrigTop         =   2250
      OrigRight       =   3510
      OrigBottom      =   2580
      Max             =   999999999
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown upStep 
      Height          =   330
      Left            =   3015
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   585
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Value           =   1
      OrigLeft        =   3270
      OrigTop         =   2250
      OrigRight       =   3510
      OrigBottom      =   2580
      Max             =   999999999
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin Axiom.FakeButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   1980
      TabIndex        =   13
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Caption         =   "Cancel"
      ForeColor       =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Delimiter"
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
      Left            =   1110
      TabIndex        =   7
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Number of Digits"
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
      Left            =   480
      TabIndex        =   6
      Top             =   1035
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Step"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   630
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Start Nubering From"
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
      Left            =   255
      TabIndex        =   0
      Top             =   270
      Width           =   1740
   End
End
Attribute VB_Name = "frmLineNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ml_NumStart As Long
Private ml_NumStep As Long
Private ml_NumDigits As Long
Private ms_Delimiter As String
Private mb_IgnoreEmptyLines As Boolean
Private mb_Canceled As Boolean

Public Property Get Canceled() As Boolean
       Canceled = mb_Canceled
End Property

Public Property Let Canceled(ByVal bNewValue As Boolean)
       mb_Canceled = bNewValue
End Property

Public Property Get IgnoreEmptyLines() As Boolean
       IgnoreEmptyLines = mb_IgnoreEmptyLines
End Property

Public Property Let IgnoreEmptyLines(ByVal bNewValue As Boolean)
       mb_IgnoreEmptyLines = bNewValue
End Property

Public Property Get Delimiter() As String
       Delimiter = ms_Delimiter
End Property

Public Property Let Delimiter(ByVal sNewValue As String)
       ms_Delimiter = sNewValue
End Property

Public Property Get NumDigits() As Long
       NumDigits = ml_NumDigits
End Property

Public Property Let NumDigits(ByVal lNewValue As Long)
       ml_NumDigits = lNewValue
End Property

Public Property Get NumStep() As Long
       NumStep = ml_NumStep
End Property

Public Property Let NumStep(ByVal lNewValue As Long)
       ml_NumStep = lNewValue
End Property

Public Property Get NumStart() As Long
       NumStart = ml_NumStart
End Property

Public Property Let NumStart(ByVal lNewValue As Long)
       ml_NumStart = lNewValue
End Property

Private Sub Command1_Click()

End Sub

Private Sub cmdCancel_Click()

Canceled = True
Me.Hide

End Sub


Private Sub cmdOk_Click()

Canceled = False

NumStart = CLng(Val(txtNumStart.Text))
NumStep = CLng(Val(txtNumStep.Text))
NumDigits = CLng(Val(txtDigits.Text))
Delimiter = CStr(txtDelimiter.Text)
IgnoreEmptyLines = CBool(chkIgnoreEmptyLines.Value)

 Me.Hide

End Sub


Private Sub Form_Load()

AddBorderToAllTextBoxes Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

Canceled = True

End Sub


Private Sub txtDigits_Change()
Dim v As Integer

v = Val(txtDigits.Text)
If v = 0 Then
    v = 1
    txtDigits.Text = "1"
End If

upNumDigits.Value = v

'Dim s As String
's = String(Val(txtDigits.Text), "0")
'txtNumStart.Text = Format(Val(txtNumStart.Text), s)

End Sub

Private Sub txtNumStart_Change()
Dim v As Integer

v = Val(txtNumStart.Text)
If v = 0 Then
    v = 1
    txtNumStart.Text = "1"
End If

upStart.Value = v

End Sub


Private Sub txtNumStep_Change()
Dim v As Integer

v = Val(txtNumStep.Text)
If v = 0 Then
    v = 1
    txtNumStep = "1"
End If

upStep.Value = v

End Sub


Private Sub upNumDigits_Change()

txtDigits.Text = upNumDigits.Value


End Sub

Private Sub upStart_Change()
txtNumStart = upStart.Value
End Sub


Private Sub upStep_Change()
txtNumStep.Text = upStep.Value
End Sub


