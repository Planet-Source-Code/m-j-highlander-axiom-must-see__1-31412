VERSION 5.00
Begin VB.Form frmAxiomOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Axiom Options"
   ClientHeight    =   4245
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6450
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Begin Axiom.FakeButton cmdOk 
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   3660
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Ok"
      ForeColor       =   16711680
   End
   Begin Axiom.FakeButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   19
      Top             =   3660
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Cancel"
      ForeColor       =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   " Edit Windows "
      Height          =   2220
      Left            =   4410
      TabIndex        =   11
      Top             =   1215
      Width           =   1905
      Begin Axiom.ColorButton clrInput 
         Height          =   240
         Left            =   1125
         TabIndex        =   15
         Top             =   315
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   423
      End
      Begin Axiom.ColorButton clrOutput 
         Height          =   240
         Left            =   1125
         TabIndex        =   16
         Top             =   630
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   423
      End
      Begin Axiom.ColorButton clrPad 
         Height          =   240
         Left            =   1125
         TabIndex        =   17
         Top             =   990
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   423
      End
      Begin Axiom.CoolButton cmdFont 
         Height          =   690
         Left            =   135
         TabIndex        =   18
         Top             =   1395
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1217
         Caption         =   "Font"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pad Color"
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   990
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Output Color"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   675
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Input Color"
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   315
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Remove Char Options "
      Height          =   1590
      Left            =   180
      TabIndex        =   7
      Top             =   1215
      Width           =   4020
      Begin VB.OptionButton optAllTextChars 
         Caption         =   "All Text Chars (including Tabs and Linebreaks)"
         Height          =   420
         Left            =   225
         TabIndex        =   10
         Top             =   1080
         Value           =   -1  'True
         Width           =   3750
      End
      Begin VB.OptionButton optPrintable 
         Caption         =   "All Printable Chars"
         Height          =   420
         Left            =   225
         TabIndex        =   9
         Top             =   675
         Width           =   2940
      End
      Begin VB.OptionButton optAlphaNum 
         Caption         =   "Alpha Numeric Chars only"
         Height          =   330
         Left            =   225
         TabIndex        =   8
         Top             =   315
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   420
      Left            =   2880
      ScaleHeight     =   420
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   630
      Width           =   3435
      Begin VB.TextBox txtSpacesPerTab 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1515
         MaxLength       =   2
         TabIndex        =   4
         Top             =   60
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Convert one TAB to "
         Height          =   195
         Left            =   0
         TabIndex        =   6
         Top             =   90
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SPACEs"
         Height          =   195
         Left            =   2160
         TabIndex        =   5
         Top             =   90
         Width           =   600
      End
   End
   Begin VB.CheckBox chkMultipleTextBoxes 
      Caption         =   "Multiple Textboxes"
      Height          =   375
      Left            =   270
      TabIndex        =   2
      Top             =   180
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkWordWrap 
      Caption         =   "Word Wrap"
      Height          =   375
      Left            =   270
      TabIndex        =   1
      Top             =   675
      Width           =   2175
   End
   Begin VB.CheckBox chkTrimTabs 
      Caption         =   "Trim TABs when trimming spaces"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   135
      Width           =   3795
   End
End
Attribute VB_Name = "frmAxiomOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CenterCtlInCtl(ctlParent As Control, ctlChild As Control)

ctlChild.Left = (ctlParent.ScaleWidth - ctlChild.Width) / 2
ctlChild.Top = (ctlParent.ScaleHeight - ctlChild.Height) / 2

End Sub


Private Sub cmdCancel_Click()

Me.Hide
End Sub

Private Sub cmdFont_Click()
'cdlCFFixedPitchOnly or cdlCFScreenFonts Or cdlCFEffects

With frmAxiomMain.cdlg
On Error Resume Next
    .flags = cdlCFBoth Or cdlCFANSIOnly Or cdlCFForceFontExist
    .FontBold = cmdFont.Font.Bold
    .FontName = cmdFont.Font.Name
    .FontSize = cmdFont.Font.Size
    frmAxiomMain.cdlg.ShowFont
If Err Then Exit Sub

cmdFont.Font.Bold = .FontBold
cmdFont.Font.Name = .FontName
cmdFont.Font.Size = .FontSize
cmdFont.Caption = "Font" 'to refresh
'cmdFont.ForeColor = .Color

End With

End Sub

Private Sub cmdOk_Click()

'***** SET OPTIONS (Properties of AxiomSettings CLASS) ******'

'1) MULTIPLE / SINGLE TextBox
If chkMultipleTextBoxes.Value = vbChecked Then
    AxiomSettings.MultipleTextBoxes = True
ElseIf chkMultipleTextBoxes.Value = vbUnchecked Then
    AxiomSettings.MultipleTextBoxes = False
End If

'2) WORD WRAP ?
If chkWordWrap.Value = vbChecked Then
    AxiomSettings.WordWrap = True
Else
    AxiomSettings.WordWrap = False
End If

'3) TRIM TABS ?
If chkTrimTabs.Value = vbChecked Then
    AxiomSettings.TrimTabsAlso = True
Else
    AxiomSettings.TrimTabsAlso = False
End If
'''''''''''''''''''''''''''''''''''''''

'4) SPACES PER TAB #
AxiomSettings.SpacesPerTab = Max(Val(txtSpacesPerTab.Text), 1)  'Minimum is 1'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'5) CHARS TO KEEP:
If optAlphaNum.Value = True Then
    AxiomSettings.CharsToKeep = [AlphaNumeric Only]

ElseIf optPrintable.Value = True Then
    AxiomSettings.CharsToKeep = [All Printable]

ElseIf optAllTextChars.Value = True Then
    AxiomSettings.CharsToKeep = [All Text Chars]
End If
'''''''''''''''''''''''''''''''''''''''''

'6) COLORS:
AxiomSettings.Colors = CStr(clrInput.Color) & "," & CStr(clrOutput.Color) & "," & CStr(clrPad.Color)

'7) FONT:
AxiomSettings.TextFont = CStr(cmdFont.Font.Name) & "," _
                       & CStr(cmdFont.Font.Size) & "," _
                       & CStr(cmdFont.Font.Bold)

Me.Hide

End Sub

Private Sub Command1_Click()
End Sub

Private Sub CoolButton1_Click()

End Sub

Private Sub Form_Activate()

With frmAxiomMain.MainText
    cmdFont.Font.Bold = .Font.Bold
    cmdFont.Font.Name = .Font.Name
    cmdFont.Font.Size = .Font.Size
    'cmdFont.ForeColor = .ForeColor
End With
    
clrInput.Color = frmAxiomMain.MainText.BackColor
clrInput.hwndOwner = Me.hWnd
clrOutput.Color = frmAxiomMain.OutputText.BackColor
clrOutput.hwndOwner = Me.hWnd
clrPad.Color = frmAxiomMain.PadText.BackColor
clrPad.hwndOwner = Me.hWnd

End Sub

Private Sub Form_Load()

'Set Style for Controls:
AddBorderToAllTextBoxes Me
CButtons Me
NumbersOnly txtSpacesPerTab

'Sync GUI with current settings:
''''''''''''''''''''''''''''''''
'1) MULTIPLE / SINGLE TextBox
If AxiomSettings.MultipleTextBoxes Then
      chkMultipleTextBoxes.Value = vbChecked
Else
      chkMultipleTextBoxes.Value = vbUnchecked
End If

'2) WORD WRAP
If AxiomSettings.WordWrap Then
    chkWordWrap.Value = vbChecked
Else
    chkWordWrap.Value = vbUnchecked
End If

'3) TRIM TABS
If AxiomSettings.TrimTabsAlso = True Then
    chkTrimTabs.Value = vbChecked
Else
    chkTrimTabs.Value = vbUnchecked
End If

'4) SPACES PER TAB #
txtSpacesPerTab.Text = CStr(Max(AxiomSettings.SpacesPerTab, 1)) 'minimum is 1

'5) CHARS TO KEEP
Select Case AxiomSettings.CharsToKeep
    Case [AlphaNumeric Only]
        optAlphaNum.Value = True
    Case [All Printable]
        optPrintable.Value = True
    Case [All Text Chars]
        optAllTextChars.Value = True
End Select

End Sub

