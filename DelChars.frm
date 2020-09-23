VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDelChars 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Chars"
   ClientHeight    =   3285
   ClientLeft      =   2475
   ClientTop       =   1545
   ClientWidth     =   4485
   Icon            =   "DelChars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   ShowInTaskbar   =   0   'False
   Begin Axiom.FakeButton cmdOk 
      Default         =   -1  'True
      Height          =   435
      Left            =   315
      TabIndex        =   14
      Top             =   2745
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   767
      Caption         =   "Ok"
      ForeColor       =   16711680
   End
   Begin Axiom.FakeButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   2535
      TabIndex        =   13
      Top             =   2745
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   767
      Caption         =   "Cancel"
      ForeColor       =   255
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Match Case"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2565
      TabIndex        =   0
      Top             =   1650
      Width           =   1215
   End
   Begin VB.CheckBox chkInclusive 
      Caption         =   "Inclusive"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2565
      TabIndex        =   1
      Top             =   1395
      Width           =   1215
   End
   Begin VB.TextBox txtDelFirst 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   1830
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "1"
      Top             =   180
      Width           =   915
   End
   Begin VB.OptionButton optDelEnd 
      Caption         =   "Delete from End to"
      Height          =   225
      Left            =   525
      TabIndex        =   6
      Top             =   1665
      Width           =   2655
   End
   Begin VB.OptionButton optDelStart 
      Caption         =   "Delete from Start to"
      Height          =   225
      Left            =   525
      TabIndex        =   5
      Top             =   1380
      Width           =   2145
   End
   Begin VB.TextBox txtDelTo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   795
      TabIndex        =   4
      Top             =   2025
      Width           =   2805
   End
   Begin VB.TextBox txtDelLast 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   1845
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "1"
      Top             =   585
      Width           =   915
   End
   Begin MSComCtl2.UpDown upDelLast 
      Height          =   330
      Left            =   2850
      TabIndex        =   2
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
      Max             =   999
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown upDelFirst 
      Height          =   330
      Left            =   2835
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   135
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Value           =   1
      OrigLeft        =   3270
      OrigTop         =   2250
      OrigRight       =   3510
      OrigBottom      =   2580
      Max             =   999
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.OptionButton optDelFirst 
      Caption         =   "Delete First"
      Height          =   225
      Left            =   570
      TabIndex        =   11
      Top             =   225
      Value           =   -1  'True
      Width           =   2145
   End
   Begin VB.OptionButton optDelLast 
      Caption         =   "Delete Last"
      Height          =   225
      Left            =   570
      TabIndex        =   12
      Top             =   630
      Width           =   2145
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   284
      X2              =   4
      Y1              =   75
      Y2              =   75
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Chars"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3270
      TabIndex        =   10
      Top             =   225
      Width           =   420
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Chars"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3285
      TabIndex        =   9
      Top             =   645
      Width           =   420
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000003&
      X1              =   284
      X2              =   4
      Y1              =   74
      Y2              =   74
   End
End
Attribute VB_Name = "frmDelChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum DeletionType
    None = -1
    DelFirstChars = 1
    DelLastChars = 2
    DelFromStart = 3
    DelFromEnd = 4
End Enum


Private mi_WhichDeletionType As Integer

Property Get WhichDeletionType() As Integer
    
    If optDelFirst.Value = True Then
        WhichDeletionType = DelFirstChars
    
    ElseIf optDelLast.Value = True Then
        WhichDeletionType = DelLastChars
    
    ElseIf optDelStart.Value = True Then
        WhichDeletionType = DelFromStart
    
    ElseIf optDelEnd.Value = True Then
        WhichDeletionType = DelFromEnd
    
    End If
    
    If mi_WhichDeletionType = -1 Then
        WhichDeletionType = None
        mi_WhichDeletionType = 0
    End If
    
End Property

Property Let WhichDeletionType(iNewValue As Integer)
    mi_WhichDeletionType = iNewValue
End Property

Private Sub cmdCancel_Click()

WhichDeletionType = None
Me.Hide
End Sub

Private Sub cmdOk_Click()
Me.Hide
End Sub


Private Sub Form_Load()

CButtons Me
AddBorderToAllTextBoxes Me
NumbersOnly txtDelFirst
NumbersOnly txtDelLast

End Sub

Private Sub optDelFirst_Click()

txtDelFirst.SetFocus

End Sub


Private Sub optDelLast_Click()

txtDelLast.SetFocus

End Sub


Private Sub txtDelFirst_Change()
Dim v As Integer

v = Val(txtDelFirst.Text)
If v = 0 Then
    v = 1
    txtDelFirst = "1"
End If

upDelFirst.Value = v

End Sub

Private Sub txtDelFirst_GotFocus()

optDelFirst.Value = True

End Sub


Private Sub txtDelLast_Change()
Dim v As Integer


v = Val(txtDelLast.Text)
If v = 0 Then
    v = 1
    txtDelLast = "1"
End If

upDelLast.Value = v

End Sub

Private Sub txtDelLast_GotFocus()

optDelLast.Value = True

End Sub


Private Sub upDelFirst_Change()
optDelFirst.Value = True
txtDelFirst.Text = upDelFirst.Value

End Sub

Private Sub upDelLast_Change()
optDelLast.Value = True
txtDelLast.Text = upDelLast.Value

End Sub


Public Property Get DelFirst() As Integer
    DelFirst = Val(txtDelFirst.Text)
End Property

Public Property Get DelLast() As Integer
    DelLast = Val(txtDelLast.Text)
End Property

Public Property Get DelToWhat() As String
    DelToWhat = CStr(txtDelTo.Text)
End Property

Public Property Get Inclusive() As Boolean

    If chkInclusive.Value = 0 Then
    Inclusive = False
    Else
    Inclusive = True
    End If

End Property

Public Property Get MatchCase() As Boolean

    If chkMatchCase.Value = 0 Then
        MatchCase = False
    Else
        MatchCase = True
    End If

End Property


