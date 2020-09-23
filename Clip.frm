VERSION 5.00
Begin VB.Form frmLocalClipboard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Local Clipboard"
   ClientHeight    =   5700
   ClientLeft      =   2145
   ClientTop       =   720
   ClientWidth     =   5460
   ControlBox      =   0   'False
   Icon            =   "Clip.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPreview 
      BackColor       =   &H8000000F&
      Height          =   2040
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   3645
      Width           =   5370
   End
   Begin Axiom.CoolButton cmdPasteAll 
      Height          =   495
      Left            =   630
      TabIndex        =   5
      ToolTipText     =   "Paste All Items"
      Top             =   45
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   ""
      ShowFocusRect   =   0   'False
      BackPicture     =   "Clip.frx":000C
   End
   Begin Axiom.CoolButton cmdClear 
      Height          =   495
      Left            =   1755
      TabIndex        =   4
      ToolTipText     =   "Clear All"
      Top             =   45
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   873
      Caption         =   ""
      ShowFocusRect   =   0   'False
      BackPicture     =   "Clip.frx":029E
   End
   Begin Axiom.CoolButton cmdRemoveItem 
      Height          =   495
      Left            =   1215
      TabIndex        =   3
      ToolTipText     =   "Remove Selected Item"
      Top             =   45
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   ""
      ShowFocusRect   =   0   'False
      BackPicture     =   "Clip.frx":0530
   End
   Begin Axiom.CoolButton cmdCopyToWindowsClip 
      Height          =   450
      Left            =   4410
      TabIndex        =   2
      ToolTipText     =   "Copy Selected Item to Windows Clipboard"
      Top             =   45
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   ""
      ShowFocusRect   =   0   'False
      BackPicture     =   "Clip.frx":07C2
   End
   Begin Axiom.CoolButton cmdPasteItem 
      Height          =   500
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Paste Selected Item"
      Top             =   45
      Width           =   500
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   ""
      ShowFocusRect   =   0   'False
      BackPicture     =   "Clip.frx":0A54
   End
   Begin VB.ListBox lstLocal 
      Height          =   2985
      Left            =   45
      TabIndex        =   0
      Top             =   585
      Width           =   5370
   End
   Begin Axiom.CoolButton cmdCancel 
      Height          =   450
      Left            =   4950
      TabIndex        =   6
      ToolTipText     =   "Cancel"
      Top             =   45
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   ""
      ShowFocusRect   =   0   'False
      BackPicture     =   "Clip.frx":0CE6
   End
End
Attribute VB_Name = "frmLocalClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum LocalClipboardOps
    
    PasteItem
    PasteAll
    DummyCancel
End Enum




Private ml_Operation As Long

Public Property Get Operation() As LocalClipboardOps
       Operation = ml_Operation
End Property

Public Property Let Operation(ByVal lNewValue As LocalClipboardOps)
       ml_Operation = lNewValue
End Property

Private Sub cmdCancel_Click()

     Operation = DummyCancel
     Me.Hide

End Sub

Private Sub cmdClear_Click()

If lstLocal.ListCount >= 1 Then
    frmAxiomMain.mo_CStrList.Clear
    lstLocal.Clear
End If

txtPreview.Text = ""
lstLocal.SetFocus
Me.Hide


End Sub

Private Sub cmdCopyToWindowsClip_Click()

If lstLocal.ListCount >= 1 Then
    Clipboard.SetText frmAxiomMain.mo_CStrList.Item(lstLocal.ListIndex)
End If

lstLocal.SetFocus

End Sub


Private Sub cmdPasteAll_Click()

Operation = PasteAll
Me.Hide

End Sub

Private Sub cmdPasteItem_Click()

Operation = PasteItem
Me.Hide


End Sub

Private Sub cmdRemoveItem_Click()

If lstLocal.ListCount >= 1 Then
    frmAxiomMain.mo_CStrList.RemoveItem lstLocal.ListIndex
    lstLocal.RemoveItem lstLocal.ListIndex
End If

lstLocal.SetFocus

End Sub








Private Sub Form_Activate()

lstLocal.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyEscape Then
     Me.Hide
End If


If KeyAscii = vbKeyReturn Then
     cmdPasteItem_Click
End If

End Sub

Private Sub Form_Load()

txtPreview.Text = ""

'Disable Context Menu
 
 'OldTextBoxProc = SetWindowLong( _
       txtPreview.hWnd, GWL_WNDPROC, _
        AddressOf NewTextBoxProc)

End Sub

Private Sub Form_Unload(Cancel As Integer)

''Return Default Handler
' OldTextBoxProc = SetWindowLong( _
'       txtPreview.hWnd, GWL_WNDPROC, _
'        AddressOf OldTextBoxProc)


End Sub


Private Sub lstLocal_Click()

If lstLocal.ListCount >= 1 Then
    txtPreview.Text = frmAxiomMain.mo_CStrList.Item(lstLocal.ListIndex)
End If

End Sub

Private Sub lstLocal_DblClick()

cmdPasteItem_Click

End Sub


Private Sub lstLocal_GotFocus()

If lstLocal.ListIndex = -1 And lstLocal.ListCount >= 1 Then
        lstLocal.ListIndex = 0
End If

End Sub


