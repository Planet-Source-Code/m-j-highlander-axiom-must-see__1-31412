VERSION 5.00
Begin VB.Form frmHTMLOps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTML Tag Operations"
   ClientHeight    =   4410
   ClientLeft      =   1350
   ClientTop       =   1230
   ClientWidth     =   5775
   Icon            =   "HTMLOps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   " Add Tag"
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   3075
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   660
         Width           =   1995
      End
      Begin Axiom.CoolButton cmdAdd 
         Height          =   315
         Left            =   1980
         TabIndex        =   11
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         Caption         =   "Add  -->"
      End
      Begin VB.TextBox txtTag 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   300
         Width           =   1995
      End
      Begin VB.CheckBox chkTagSingle 
         Caption         =   "Tag is Single"
         Height          =   195
         Left            =   960
         TabIndex        =   8
         Top             =   1080
         Width           =   1275
      End
      Begin Axiom.CoolButton CoolButton1 
         Height          =   315
         Left            =   960
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         Caption         =   "Edit"
      End
      Begin VB.Label Description 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tag"
         Height          =   195
         Left            =   540
         TabIndex        =   10
         Top             =   360
         Width           =   285
      End
   End
   Begin Axiom.FakeButton cmdOk 
      Default         =   -1  'True
      Height          =   435
      Left            =   1125
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Caption         =   "Ok"
      ForeColor       =   16711680
   End
   Begin Axiom.FakeButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   3180
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Caption         =   "Cancel"
      ForeColor       =   255
   End
   Begin VB.ListBox lstTags 
      Height          =   3375
      Left            =   3300
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   360
      Width           =   2355
   End
   Begin VB.Frame Frame1 
      Caption         =   " Operation "
      Height          =   1590
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   3075
      Begin VB.OptionButton optExtract 
         Caption         =   "Extract Tag Contents"
         Height          =   195
         Left            =   300
         TabIndex        =   4
         Top             =   1140
         Width           =   2115
      End
      Begin VB.OptionButton optDelTag 
         Caption         =   "Delete Tag And Contents"
         Height          =   195
         Left            =   300
         TabIndex        =   3
         Top             =   405
         Value           =   -1  'True
         Width           =   2715
      End
      Begin VB.OptionButton optDelTagKeepContent 
         Caption         =   "Delete Tag but Leave Contents"
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Top             =   780
         Width           =   2595
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Select Tag(s)"
      Height          =   195
      Left            =   3300
      TabIndex        =   12
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmHTMLOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private menum_Operation As HTML_Operation
Private ms_HTMLTags() As String
Private mb_TagIsSingle() As Boolean

Private Tags() As String
Private Const MAX_NUM_TAGS = 10
Sub PresetTags(TagsArray() As String)

ReDim TagsArray(0 To MAX_NUM_TAGS)

TagsArray(0) = "<style"
TagsArray(1) = "<img>"
TagsArray(2) = "<script"
TagsArray(3) = "<iframe"
TagsArray(4) = "<iframe>"
TagsArray(5) = "<textarea"
TagsArray(6) = "<pre"
TagsArray(7) = "<meta>"
TagsArray(8) = "<marquee"
TagsArray(9) = "<input>"
TagsArray(10) = "<a"

StrSort TagsArray, True, True

End Sub

Public Property Get HTMLTags() As String()
       HTMLTags = ms_HTMLTags()
End Property

Public Property Let HTMLTags(ByRef sNewValue() As String)
       ms_HTMLTags() = sNewValue()
End Property


Public Property Get TagIsSingle() As Boolean()
       TagIsSingle = mb_TagIsSingle()
End Property

Public Property Let TagIsSingle(ByRef bNewValue() As Boolean)
       
       mb_TagIsSingle() = bNewValue()
       
End Property

Public Property Get Operation() As HTML_Operation
       Operation = menum_Operation
End Property

Public Property Let Operation(ByVal enumNewValue As HTML_Operation)
       menum_Operation = enumNewValue
End Property

Private Sub cmdAdd_Click()
Dim idx As Long

If txtTag.Text = "" Then
    Beep
    Exit Sub
End If
If txtDescription.Text = "" Then
    txtDescription.Text = StrConv(txtTag.Text, vbProperCase)
End If

go_HTMLTags.Add LCase$(txtTag.Text), txtDescription.Text, CBool(chkTagSingle.Value)

lstTags.Clear
For idx = 0 To go_HTMLTags.Count - 1
    lstTags.AddItem go_HTMLTags.Description(idx)
Next idx

End Sub

Private Sub cmdCancel_Click()

Operation = Cancel
Me.Hide

End Sub

Private Sub cmdOk_Click()
ReDim tmpArray(0 To 99) As String
ReDim tmpBArray(0 To 99) As Boolean

Dim idx As Long, cntr As Long


'[ Set Properties ]
'HTMLTag = txtOpenTag.Text
idx = 0
For cntr = 0 To go_HTMLTags.Count - 1
    If lstTags.Selected(cntr) Then
        tmpArray(idx) = go_HTMLTags.Name(cntr)
        tmpBArray(idx) = go_HTMLTags.IsSingle(cntr)
        idx = idx + 1
    End If
Next cntr
If idx = 0 Then 'nothing selected
    Beep
    Exit Sub
End If

ReDim Preserve tmpArray(0 To idx - 1)
HTMLTags = tmpArray()
TagIsSingle = tmpBArray()

If optDelTag.Value = True Then
    Operation = DeleteTagAndContent

ElseIf optDelTagKeepContent.Value = True Then
    Operation = DeleteTagKeepContent

ElseIf optExtract.Value = True Then
    Operation = ExtractTagAndContent
End If

Me.Hide

End Sub

Private Sub Form_Load()

AddBorderToAllTextBoxes Me
CButtons Me

Dim idx As Integer


For idx = 0 To go_HTMLTags.Count - 1
    lstTags.AddItem go_HTMLTags.Description(idx)
Next idx

End Sub


Private Sub lstTags_Click()
Dim idx As Long
idx = lstTags.ListIndex

txtTag.Text = go_HTMLTags.Name(idx)
txtDescription.Text = go_HTMLTags.Description(idx)
chkTagSingle.Value = Abs(go_HTMLTags.IsSingle(idx))

End Sub


