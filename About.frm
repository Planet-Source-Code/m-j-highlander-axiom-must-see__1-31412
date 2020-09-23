VERSION 5.00
Begin VB.Form frmAboutAxiom 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   1860
   ClientTop       =   1740
   ClientWidth     =   5595
   ControlBox      =   0   'False
   Icon            =   "About.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5595
   Begin Axiom.CoolButton CoolButton1 
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Top             =   3180
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   503
      Caption         =   "Ok"
      ForeColor       =   16711680
      ShowFocusRect   =   0   'False
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "About.frx":000C
      Top             =   1860
      Width           =   2355
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   0
      ScaleHeight     =   3600
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   0
      Width           =   1185
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Top             =   900
         Width           =   480
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3465
         Left            =   180
         TabIndex        =   4
         Top             =   60
         Width           =   810
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "BETA VERSION"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   1620
      TabIndex        =   6
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Multifunctional Text/HTML Processor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1380
      TabIndex        =   2
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Axiom"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1755
      Left            =   1395
      TabIndex        =   1
      Top             =   0
      Width           =   3690
   End
End
Attribute VB_Name = "frmAboutAxiom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End
End Sub


Private Sub CoolButton1_Click()
Unload Me

End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me

End Sub


Private Sub Form_Load()
Image1.Picture = frmAxiomMain.Icon
lblV.Caption = "A" & vbCrLf & " " & vbCrLf & "i" & vbCrLf & "o" & vbCrLf & "m"

'SUBCLASSING:
'OldAboutTextBoxProc = SetWindowLong( _
       Text1.hWnd, GWL_WNDPROC, _
        AddressOf NewAboutTextBoxProc)

End Sub


Private Sub Label1_Click()
Unload Me

End Sub

Private Sub Label2_Click()
Unload Me

End Sub


Private Sub lblV_Click()
Unload Me

End Sub


Private Sub Picture1_Click()
Unload Me

End Sub


Private Sub Text1_Click()
Unload Me

End Sub


