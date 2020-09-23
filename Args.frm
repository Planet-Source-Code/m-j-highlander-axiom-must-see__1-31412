VERSION 5.00
Begin VB.Form frmArgs 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2985
   ClientLeft      =   1665
   ClientTop       =   1065
   ClientWidth     =   6450
   Icon            =   "Args.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6450
   Begin Axiom.FakeButton cmdOk 
      Default         =   -1  'True
      Height          =   420
      Left            =   3105
      TabIndex        =   2
      Top             =   1350
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   741
      Caption         =   "Ok"
      ForeColor       =   16711680
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Index           =   0
      Left            =   1530
      TabIndex        =   1
      Top             =   135
      Width           =   4785
   End
   Begin Axiom.FakeButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   4815
      TabIndex        =   3
      Top             =   1350
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   741
      Caption         =   "Cancel"
      ForeColor       =   255
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "xxxxxxxxxxxx"
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   180
      Width           =   900
   End
End
Attribute VB_Name = "frmArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ml_NumArgs As Long
Private mb_Canceled As Boolean

Public Property Get Canceled() As Boolean
       Canceled = mb_Canceled
End Property

Public Property Let Canceled(ByVal bNewValue As Boolean)
       mb_Canceled = bNewValue
End Property

Public Sub GetVals(ByRef sArray() As String)
Dim idx As Long

'ReDim sArray(0 To NumArgs - 1)

For idx = 0 To NumArgs - 1
    sArray(idx) = txtText(idx).Text
Next idx
    
End Sub

Public Sub Labels(ByRef lNewValue() As String)
Dim idx As Long

For idx = 0 To NumArgs - 1
    lblLabel(idx).Caption = lNewValue(idx)
    lblLabel(idx).Left = txtText(idx).Left - lblLabel(idx).Width - 150
Next idx
    
End Sub

Public Property Get NumArgs() As Long
       NumArgs = ml_NumArgs
End Property

Public Property Let NumArgs(ByVal lNewValue As Long)
Dim idx As Long
       ml_NumArgs = lNewValue
    
lblLabel(0).Top = txtText(0).Top + 100

For idx = 1 To lNewValue - 1
    Load lblLabel(idx)
    Load txtText(idx)
    txtText(idx).Top = txtText(idx - 1).Top + txtText(idx - 1).Height + 100
    lblLabel(idx).Top = txtText(idx).Top + 100
    lblLabel(idx).Visible = True
    txtText(idx).Visible = True
Next idx

cmdOk.Top = txtText(idx - 1).Top + txtText(idx - 1).Height + 200
cmdCancel.Top = txtText(idx - 1).Top + txtText(idx - 1).Height + 200

Me.Height = cmdOk.Top + cmdOk.Height + 500

End Property

Private Sub cmdCancel_Click()

Canceled = True
Me.Hide

End Sub

Private Sub cmdOk_Click()

Canceled = False
Me.Hide

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)

Canceled = True

End Sub


