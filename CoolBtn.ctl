VERSION 5.00
Begin VB.UserControl CoolButton 
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   39
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   ToolboxBitmap   =   "CoolBtn.ctx":0000
End
Attribute VB_Name = "CoolButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click()
Public Event MouseDown()
Public Event MouseUp()


Private ms_Caption As String
Private m_Picture As Picture
Private mb_ShowFocusRect As Boolean

Public Sub About()
Attribute About.VB_UserMemId = -552
    MsgBox "CoolButton by highlander@nycny.net", vbInformation, "About"
End Sub


Public Property Get ShowFocusRect() As Boolean
       ShowFocusRect = mb_ShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal bNewValue As Boolean)
       mb_ShowFocusRect = bNewValue
End Property

Private Sub DrawPic(bDown As Boolean)
    Dim ButtonTop As Long
    Dim BkColor As Long
    Dim TLng1 As Double, TLng2 As Double
    Dim CX As Long, CY As Long
    Dim w As Long, h As Long
    
    Dim hMemDC As Long
    Dim hOldBmp As Long
    

    If (m_Picture Is Nothing) Then
        'no pic =  m_Picture.Width = 0
    Else
            'Check picture dimensions
        CX = UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels)
        CY = UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels)
        w = 0 '(UserControl.ScaleWidth - cx) / 2
        h = 0 '(UserControl.ScaleHeight - cy) / 2
        If bDown = True Then
            w = w + 1
            h = h + 1
        End If
        hMemDC = CreateCompatibleDC(UserControl.hDC)
        hOldBmp = SelectObject(hMemDC, m_Picture.Handle)
                
        BitBlt UserControl.hDC, w, h, CX, CY, hMemDC, 0&, 0&, vbSrcCopy
        'StretchBlt UserControl.hdc, 0&, ButtonTop, 12&, 24&, hMemDC, 0&, 0&, cx, cy, vbSrcCopy

        SelectObject hMemDC, hOldBmp
        DeleteDC hMemDC
    End If

End Sub


Public Property Get BackPicture() As Picture
    Set BackPicture = m_Picture
End Property

Public Property Set BackPicture(ByVal picNewValue As Picture)
   
  Set m_Picture = picNewValue
       PropertyChanged "BackPicture"
       UserControl_Paint

End Property

Public Property Get Font() As IFontDisp
    
    Set Font = UserControl.Font
    
    Font.Bold = UserControl.Font.Bold
    Font.Charset = UserControl.Font.Charset
    Font.Italic = UserControl.Font.Italic
    Font.Name = UserControl.Font.Name
    Font.Size = UserControl.Font.Size
    Font.Strikethrough = UserControl.Font.Strikethrough
    Font.Underline = UserControl.Font.Underline
    Font.Weight = UserControl.Font.Weight

End Property

Public Property Set Font(ByVal fontNewValue As IFontDisp)
    'chk.Font = fontNewValue
    
    UserControl.Font.Bold = fontNewValue.Bold
    UserControl.Font.Charset = fontNewValue.Charset
    UserControl.Font.Italic = fontNewValue.Italic
    UserControl.Font.Name = fontNewValue.Name
    UserControl.Font.Size = fontNewValue.Size
    UserControl.Font.Strikethrough = fontNewValue.Strikethrough
    UserControl.Font.Underline = fontNewValue.Underline
    UserControl.Font.Weight = fontNewValue.Weight
    PropertyChanged "Font"
    UserControl_Paint

End Property


Public Property Get ForeColor() As OLE_COLOR
       ForeColor = UserControl.ForeColor
        
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
       UserControl.ForeColor = NewValue
       PropertyChanged "ForeColor"
       UserControl_Paint

End Property

Public Property Get BackColor() As OLE_COLOR
       BackColor = UserControl.BackColor
        
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
       UserControl.BackColor = NewValue
       PropertyChanged "BackColor"
       UserControl_Paint

End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
       Caption = ms_Caption
End Property

Public Property Let Caption(ByVal sNewValue As String)
       ms_Caption = sNewValue
       PropertyChanged "Caption"
       UserControl_Paint
End Property

Private Sub UserControl_GotFocus()

Dim r As RECT

r.Left = 3
r.Top = 3
r.Right = ScaleWidth - 3
r.Bottom = ScaleHeight - 3

If ShowFocusRect Then DrawFocusRect hDC, r
'Refresh



End Sub

Private Sub UserControl_InitProperties()

Caption = Extender.Name
'Set Picture = Nothing


End Sub


Private Sub UserControl_LostFocus()
UserControl_Paint
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> vbLeftButton Then Exit Sub

Dim r As RECT
Cls

DrawPic True

r.Left = 0
r.Top = 0
r.Right = ScaleWidth
r.Bottom = ScaleHeight
DrawEdge hDC, r, BDR_SUNKENOUTER, BF_RECT

r.Left = 2
r.Top = 2
r.Right = ScaleWidth
r.Bottom = ScaleHeight
DrawText hDC, Caption, Len(Caption), r, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER

RaiseEvent MouseDown

End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> vbLeftButton Then Exit Sub

Dim r As RECT
Cls

DrawPic False

r.Left = 0
r.Top = 0
r.Right = ScaleWidth
r.Bottom = ScaleHeight
DrawEdge hDC, r, BDR_RAISEDOUTER, BF_RECT

DrawText hDC, Caption, Len(Caption), r, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER


RaiseEvent Click
RaiseEvent MouseUp

UserControl_GotFocus

End Sub


Private Sub UserControl_Paint()
Dim r As RECT

Cls
DrawPic False

r.Left = 0
r.Top = 0
r.Right = ScaleWidth
r.Bottom = ScaleHeight
DrawEdge hDC, r, BDR_RAISEDOUTER, BF_RECT

DrawText hDC, Caption, Len(Caption), r, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER



End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Caption = PropBag.ReadProperty("Caption", Extender.Name)
BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
ShowFocusRect = PropBag.ReadProperty("ShowFocusRect", True)

Font.Bold = PropBag.ReadProperty("Font.Bold", Ambient.Font.Bold)
Font.Charset = PropBag.ReadProperty("Font.Charset", Ambient.Font.Charset)
Font.Italic = PropBag.ReadProperty("Font.Italic", Ambient.Font.Italic)
Font.Name = PropBag.ReadProperty("Font.Name", Ambient.Font.Name)
Font.Size = PropBag.ReadProperty("Font.Size", Ambient.Font.Size)
Font.Strikethrough = PropBag.ReadProperty("Font.Strikethrough", Ambient.Font.Strikethrough)
Font.Underline = PropBag.ReadProperty("Font.Underline", Ambient.Font.Underline)
Font.Weight = PropBag.ReadProperty("Font.Weight", Ambient.Font.Weight)

Set BackPicture = PropBag.ReadProperty("BackPicture", Nothing)
'Set m_picture = PropBag.ReadProperty("Picture", Nothing)

UserControl_Paint

End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

PropBag.WriteProperty "Caption", Caption, Extender.Name
PropBag.WriteProperty "BackColor", BackColor, Ambient.BackColor
PropBag.WriteProperty "ForeColor", ForeColor, Ambient.ForeColor
PropBag.WriteProperty "ShowFocusRect", ShowFocusRect, True

PropBag.WriteProperty "Font.bold", Font.Bold, Ambient.Font.Bold
PropBag.WriteProperty "Font.Charset", Font.Charset, Ambient.Font.Charset
PropBag.WriteProperty "Font.Italic", Font.Italic, Ambient.Font.Italic
PropBag.WriteProperty "Font.Name", Font.Name, Ambient.Font.Name
PropBag.WriteProperty "Font.size", Font.Size, Ambient.Font.Size
PropBag.WriteProperty "Font.Strikethrough", Font.Strikethrough, Ambient.Font.Strikethrough
PropBag.WriteProperty "Font.Underline", Font.Underline, Ambient.Font.Underline
PropBag.WriteProperty "Font.Weight", Font.Weight, Ambient.Font.Weight

PropBag.WriteProperty "BackPicture", BackPicture, Nothing
'PropBag.WriteProperty("Picture", m_picture, Nothing)


End Sub


