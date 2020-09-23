VERSION 5.00
Begin VB.Form frmFindReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find / Replace"
   ClientHeight    =   1965
   ClientLeft      =   1650
   ClientTop       =   1860
   ClientWidth     =   5910
   ForeColor       =   &H8000000D&
   Icon            =   "Replace.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin Axiom.FakeButton cmdFindNext 
      Default         =   -1  'True
      Height          =   435
      Left            =   4620
      TabIndex        =   8
      Top             =   60
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Caption         =   "&Find Next"
      ForeColor       =   -2147483630
      Font.bold       =   -1  'True
      Font.Weight     =   700
   End
   Begin Axiom.CoolButton cmdSpecialReplace 
      Height          =   330
      Left            =   4185
      TabIndex        =   7
      Top             =   510
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      Caption         =   ">"
      ForeColor       =   -2147483630
      ShowFocusRect   =   0   'False
   End
   Begin Axiom.CoolButton cmdSpecialFind 
      Height          =   330
      Left            =   4185
      TabIndex        =   6
      Top             =   105
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      Caption         =   ">"
      ForeColor       =   -2147483630
      ShowFocusRect   =   0   'False
   End
   Begin VB.CheckBox chkWord 
      Caption         =   "Whole word only"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1020
      Width           =   1635
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Match case"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1380
      Width           =   1455
   End
   Begin VB.TextBox txtReplaceWith 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1125
      TabIndex        =   1
      Top             =   510
      Width           =   2985
   End
   Begin VB.TextBox txtFindWhat 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1125
      TabIndex        =   0
      Top             =   105
      Width           =   2985
   End
   Begin Axiom.FakeButton cmdReplace 
      Height          =   435
      Left            =   4620
      TabIndex        =   9
      Top             =   510
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Caption         =   "Replace"
      ForeColor       =   -2147483630
   End
   Begin Axiom.FakeButton cmdReplaceAll 
      Height          =   435
      Left            =   4620
      TabIndex        =   10
      Top             =   975
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Caption         =   "Replace All"
      ForeColor       =   -2147483630
   End
   Begin Axiom.FakeButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   4620
      TabIndex        =   11
      Top             =   1500
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Caption         =   "Cancel"
      ForeColor       =   255
   End
   Begin VB.Label lblReplace 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replace with"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   555
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find what"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   195
      Width           =   690
   End
   Begin VB.Menu mnuSpecialChars 
      Caption         =   "mnuHIDDEN"
      Visible         =   0   'False
      Begin VB.Menu mnuPara 
         Caption         =   "Paragraph Mark"
      End
      Begin VB.Menu mnuTab 
         Caption         =   "Tab"
      End
      Begin VB.Menu mnuBackSlash 
         Caption         =   "\"
      End
   End
End
Attribute VB_Name = "frmFindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ml_CurPos As Long         ' with properties

Private ml_EndPos As Long
Private ml_StartPos As Long

Private mi_WhichTextBox As Integer ' not used with properties

Public Enum FindReplaceConstants
    Operation_Cancel = 0
    Operation_Find = 1
    Operation_Replace = 2
End Enum


Private me_Operation As FindReplaceConstants

Private mb_MatchCase As Boolean
Private mb_WholeWords As Boolean
Private mo_OwnerForm As Form


Public Property Get CurPos() As Long
       CurPos = ml_CurPos
End Property

Public Property Let CurPos(ByVal lNewValue As Long)
       ml_CurPos = lNewValue
End Property

Public Property Get OwnerForm() As Form
    Set OwnerForm = mo_OwnerForm
End Property

Public Property Set OwnerForm(ByVal oNewValue As Form)
       Set mo_OwnerForm = oNewValue
End Property

Public Property Get WholeWords() As Boolean
       WholeWords = mb_WholeWords
End Property

Public Property Let WholeWords(ByVal bNewValue As Boolean)
       mb_WholeWords = bNewValue
End Property

Public Property Get MatchCase() As Boolean
       MatchCase = mb_MatchCase
End Property

Public Property Let MatchCase(ByVal bNewValue As Boolean)
       mb_MatchCase = bNewValue
End Property

Public Property Get FindWhat() As String
    
    FindWhat = stringf(txtFindWhat.Text)
    
End Property

Public Property Let FindWhat(ByVal sNewValue As String)
    txtFindWhat.Text = sNewValue
End Property

Public Property Get ReplaceWith() As String

    ReplaceWith = stringf(txtReplaceWith.Text)
    
End Property

Public Property Let ReplaceWith(ByVal sNewValue As String)
    
    txtReplaceWith.Text = sNewValue
    
End Property


Private Sub cmdCancel_Click()

Me.Hide

End Sub

Private Sub cmdFindNext_Click()
Dim sFindWhat As String

sFindWhat = FindWhat()

With OwnerForm.mo_TheTextBox

    ml_StartPos = .SelStart
    ml_EndPos = Len(.Text)

    If CurPos = 0 Then CurPos = ml_StartPos
    'Find() returns Char Pos, 0 if not found
    CurPos = .Find(sFindWhat, CurPos, ml_EndPos, Options) + 1
    If CurPos = 0 Then Beep

End With

End Sub

Private Sub cmdReplace_Click()

Dim sFindWhat As String, sReplaceWith As String

sFindWhat = FindWhat()
sReplaceWith = ReplaceWith()

With OwnerForm.mo_TheTextBox
    If .SelLength = 0 Or LCase(.SelText) <> LCase(sFindWhat) Then
        'no selection or selection is not the "FindWhat" string
        'so do search:
        CurPos = .Find(sFindWhat, CurPos, Len(.Text), Options)
        If CurPos > 0 Then
              .SelText = sReplaceWith
              CurPos = .SelStart 'CurPos + 1 'Len(sReplaceWith)
        End If
    
    Else 'selection matches "FindWhat", so replace
        .SelText = sReplaceWith
        CurPos = .SelStart 'CurPos + 1 'Len(sReplaceWith)
    
    End If
    
End With

End Sub

Private Sub cmdReplaceAll_Click()
Dim Count As Long
Dim sFindWhat As String, sReplaceWith As String

sFindWhat = FindWhat()
sReplaceWith = ReplaceWith()

With OwnerForm.mo_TheTextBox
    LockWindowUpdate OwnerForm.hWnd
    CurPos = 0
    Do
        CurPos = .Find(sFindWhat, CurPos, Len(.Text), Options)
        If CurPos = -1 Then
            Exit Do
        End If
        .SelText = sReplaceWith
        'CurPos = CurPos + Len(sReplaceWith)
        CurPos = .SelStart
        Count = Count + 1
    Loop
    LockWindowUpdate 0
End With

Me.Hide
MsgBox "Performed " & CStr(Count) & " Replaces", vbInformation, "Replace"

End Sub

Private Sub cmdSpecialFind_MouseDown()

mi_WhichTextBox = 1 'First TextBox
PopupMenu mnuSpecialChars, 2, cmdSpecialFind.Left, cmdSpecialFind.Top + cmdSpecialFind.Height
txtFindWhat.SetFocus

End Sub


Private Sub cmdSpecialReplace_MouseDown()
mi_WhichTextBox = 2 'Second TextBox
PopupMenu mnuSpecialChars, 2, cmdSpecialReplace.Left, cmdSpecialReplace.Top + cmdSpecialReplace.Height
txtReplaceWith.SetFocus

End Sub


Private Sub Form_Load()

CButtons Me
AddBorderToAllTextBoxes Me

SetTopMost Me.hWnd, True
Me.Show
Me.txtFindWhat.SetFocus

txtFindWhat.Text = FindWhat()
txtReplaceWith.Text = ReplaceWith()
CurPos = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

gs_FindWhat = FindWhat()
gl_Options = Options


End Sub


Private Sub mnuBackSlash_Click()

If mi_WhichTextBox = 1 Then
    txtFindWhat.SelText = "\\"
Else
    txtReplaceWith.SelText = "\\"
End If

End Sub

Private Sub mnuPara_Click()

If mi_WhichTextBox = 1 Then
    txtFindWhat.SelText = "\n"
Else
    txtReplaceWith.SelText = "\n"
End If

End Sub

Private Sub mnuTab_Click()

If mi_WhichTextBox = 1 Then
    txtFindWhat.SelText = "\t"
Else
    txtReplaceWith.SelText = "\t"
End If


End Sub



Public Property Get Options() As Long
' No Property Let  ==> READ-ONLY
Dim ml_Options As Long

If (chkCase.Value = vbChecked And chkWord.Value = vbChecked) Then
    ml_Options = rtfMatchCase Or rtfWholeWord
ElseIf chkCase.Value = vbChecked Then
    ml_Options = rtfMatchCase
ElseIf chkWord.Value = vbChecked Then
    ml_Options = rtfWholeWord
Else
    ml_Options = 0
End If
    
    Options = ml_Options

End Property

Private Sub txtFindWhat_Change()
    
    CurPos = 0


End Sub


