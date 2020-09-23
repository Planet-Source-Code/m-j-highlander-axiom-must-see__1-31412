VERSION 5.00
Begin VB.UserControl FakeButton 
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   DefaultCancel   =   -1  'True
   ScaleHeight     =   1995
   ScaleWidth      =   3990
   ToolboxBitmap   =   "FakeBtn.ctx":0000
   Begin VB.CheckBox chk 
      Caption         =   "Caption goes here..."
      Height          =   1365
      Left            =   945
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   270
      Width           =   2310
   End
End
Attribute VB_Name = "FakeButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click()

Public Property Get About() As String
       About = "FakeBtn by (highlander@nycny.net)"
End Property

Public Property Let About(ByVal bNewValue As String)
       
End Property






Public Property Get Caption() As String
       Caption = chk.Caption
End Property

Public Property Let Caption(ByVal sNewValue As String)

       chk.Caption = sNewValue
End Property



Public Property Get BackColor() As OLE_COLOR
       BackColor = chk.BackColor
End Property

Public Property Let BackColor(ByVal lNewValue As OLE_COLOR)
      chk.BackColor = lNewValue
End Property


Public Property Get ForeColor() As OLE_COLOR
       ForeColor = chk.ForeColor
End Property

Public Property Let ForeColor(ByVal lNewValue As OLE_COLOR)
       chk.ForeColor = lNewValue
End Property


Private Sub chk_Click()

If chk.Value = 1 Then chk.Value = 2
RaiseEvent Click

End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()

'UserControl_Resize


End Sub

Private Sub UserControl_InitProperties()

Caption = Extender.Name

End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Caption = PropBag.ReadProperty("Caption", Extender.Name)
ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)

Font.Bold = PropBag.ReadProperty("Font.Bold", Ambient.Font.Bold)
Font.Charset = PropBag.ReadProperty("Font.Charset", Ambient.Font.Charset)
Font.Italic = PropBag.ReadProperty("Font.Italic", Ambient.Font.Italic)
Font.Name = PropBag.ReadProperty("Font.Name", Ambient.Font.Name)
Font.Size = PropBag.ReadProperty("Font.Size", Ambient.Font.Size)
Font.Strikethrough = PropBag.ReadProperty("Font.Strikethrough", Ambient.Font.Strikethrough)
Font.Underline = PropBag.ReadProperty("Font.Underline", Ambient.Font.Underline)
Font.Weight = PropBag.ReadProperty("Font.Weight", Ambient.Font.Weight)

Set Picture = PropBag.ReadProperty("Picture", UserControl.Picture)

End Sub

Private Sub UserControl_Resize()

chk.Left = 0
chk.Top = 0
chk.Width = ScaleWidth
chk.Height = ScaleHeight

End Sub



Public Property Get Picture() As IPictureDisp
    Set Picture = chk.Picture
End Property

Public Property Set Picture(ByVal picNewValue As IPictureDisp)
   Set chk.Picture = picNewValue
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

PropBag.WriteProperty "Caption", Caption, Extender.Name
PropBag.WriteProperty "ForeColor", ForeColor, Ambient.ForeColor
PropBag.WriteProperty "BackColor", BackColor, Ambient.BackColor

PropBag.WriteProperty "Font.bold", Font.Bold, Ambient.Font.Bold
PropBag.WriteProperty "Font.Charset", Font.Charset, Ambient.Font.Charset
PropBag.WriteProperty "Font.Italic", Font.Italic, Ambient.Font.Italic
PropBag.WriteProperty "Font.Name", Font.Name, Ambient.Font.Name
PropBag.WriteProperty "Font.size", Font.Size, Ambient.Font.Size
PropBag.WriteProperty "Font.Strikethrough", Font.Strikethrough, Ambient.Font.Strikethrough
PropBag.WriteProperty "Font.Underline", Font.Underline, Ambient.Font.Underline
PropBag.WriteProperty "Font.Weight", Font.Weight, Ambient.Font.Weight

PropBag.WriteProperty "Picture", Picture, UserControl.Picture


End Sub



Public Property Get Font() As IFontDisp
    
    Set Font = chk.Font
    
    Font.Bold = chk.Font.Bold
    Font.Charset = chk.Font.Charset
    Font.Italic = chk.Font.Italic
    Font.Name = chk.Font.Name
    Font.Size = chk.Font.Size
    Font.Strikethrough = chk.Font.Strikethrough
    Font.Underline = chk.Font.Underline
    Font.Weight = chk.Font.Weight

End Property

Public Property Set Font(ByVal fontNewValue As IFontDisp)
    'chk.Font = fontNewValue
    
    chk.Font.Bold = fontNewValue.Bold
    chk.Font.Charset = fontNewValue.Charset
    chk.Font.Italic = fontNewValue.Italic
    chk.Font.Name = fontNewValue.Name
    chk.Font.Size = fontNewValue.Size
    chk.Font.Strikethrough = fontNewValue.Strikethrough
    chk.Font.Underline = fontNewValue.Underline
    chk.Font.Weight = fontNewValue.Weight
    
End Property
