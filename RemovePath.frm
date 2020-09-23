VERSION 5.00
Begin VB.Form frmRemovePath 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove Path"
   ClientHeight    =   2175
   ClientLeft      =   2820
   ClientTop       =   2205
   ClientWidth     =   4275
   Icon            =   "RemovePath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4275
   Begin VB.CheckBox chkLocalOnly 
      Caption         =   "Remove local paths only"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1740
      Width           =   2295
   End
   Begin Axiom.FakeButton cmdOk 
      Default         =   -1  'True
      Height          =   435
      Left            =   3000
      TabIndex        =   4
      Top             =   1020
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   767
      Caption         =   "Ok"
      ForeColor       =   16711680
   End
   Begin VB.Frame Frame1 
      Caption         =   " Remove Path from "
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.CheckBox chkHref 
         Caption         =   "HREF"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   660
         Width           =   915
      End
      Begin VB.CheckBox chkSrc 
         Caption         =   "SRC"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox chkBackground 
         Caption         =   "BACKGROUND"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1635
      End
   End
   Begin Axiom.FakeButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   3000
      TabIndex        =   5
      Top             =   1620
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   767
      Caption         =   "Cancel"
      ForeColor       =   255
   End
End
Attribute VB_Name = "frmRemovePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mb_RemoveLocalOnly As Boolean
Private mb_RemoveBackground As Boolean
Private mb_RemoveHref As Boolean
Private mb_RemoveSrc As Boolean
Private mb_Canceled As Boolean

Public Property Get Canceled() As Boolean
       Canceled = mb_Canceled
End Property

Public Property Let Canceled(ByVal bNewValue As Boolean)
       mb_Canceled = bNewValue
End Property

Public Property Get RemoveSrc() As Boolean
       RemoveSrc = mb_RemoveSrc
End Property

Public Property Let RemoveSrc(ByVal bNewValue As Boolean)
       mb_RemoveSrc = bNewValue
End Property

Public Property Get RemoveHref() As Boolean
       RemoveHref = mb_RemoveHref
End Property

Public Property Let RemoveHref(ByVal bNewValue As Boolean)
       mb_RemoveHref = bNewValue
End Property

Public Property Get RemoveBackground() As Boolean
       RemoveBackground = mb_RemoveBackground
End Property

Public Property Let RemoveBackground(ByVal bNewValue As Boolean)
       mb_RemoveBackground = bNewValue
End Property

Public Property Get RemoveLocalOnly() As Boolean
       RemoveLocalOnly = mb_RemoveLocalOnly
End Property

Public Property Let RemoveLocalOnly(ByVal bNewValue As Boolean)
       mb_RemoveLocalOnly = bNewValue
End Property

Private Sub cmdCancel_Click()


Canceled = True
Me.Hide

End Sub

Private Sub cmdOk_Click()


RemoveBackground = CBool(chkBackground.Value)
RemoveHref = CBool(chkHref.Value)
RemoveSrc = CBool(chkSrc.Value)
RemoveLocalOnly = CBool(chkLocalOnly.Value)

Canceled = False
Me.Hide

End Sub

