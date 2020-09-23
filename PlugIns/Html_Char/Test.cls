VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTMLClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function HTML_AddBR(ByVal sText As String) As String
Attribute HTML_AddBR.VB_Description = "Add <BR> tags to text lines\r\nText"

HTML_AddBR = Replace(sText, vbCrLf, "<BR>" & vbCrLf)

End Function


Public Function HTML_FirstChar(ByVal sText, ByVal sColor As String, ByVal sSize) As String
Attribute HTML_FirstChar.VB_Description = "Change Font and Color of Every First Char\r\nInputs:Text,Font Color,Font Size"
Dim idx As Long
Dim a As Variant
Dim Ch As String, NewCh As String

a = Split(sText, " ")
For idx = LBound(a) To UBound(a)
        Ch = Left(a(idx), 1)
        NewCh = "<FONT COLOR=""%COLOR%"" SIZE=""%SIZE%"">" & Ch & "</FONT>"
        NewCh = Replace(NewCh, "%COLOR%", sColor)
        NewCh = Replace(NewCh, "%SIZE%", sSize)
        a(idx) = Replace(a(idx), Ch, NewCh, , 1)
Next idx
sText = Join(a, " ")
''''''''''''''''''''''''''''''
a = Split(sText, vbCrLf)
For idx = LBound(a) + 1 To UBound(a)
        Ch = Left(a(idx), 1)
        NewCh = "<FONT COLOR=""%COLOR%"" SIZE=""%SIZE%"">" & Ch & "</FONT>"
        NewCh = Replace(NewCh, "%COLOR%", sColor)
        NewCh = Replace(NewCh, "%SIZE%", sSize)
        a(idx) = Replace(a(idx), Ch, NewCh, , 1)
Next idx
sText = Join(a, vbCrLf)

HTML_FirstChar = sText
End Function

Public Function HTML_Font(ByVal Text As String, ByVal sFontName As String, ByVal sFontColor As String, ByVal sFontSize As String) As String
Attribute HTML_Font.VB_Description = "Apply Font Attributes to Text\r\nInputs:Text,Font Name,Font Color,FontSize"

HTML_Font = "<FONT FACE=""" & sFontName & """ COLOR=""" & sFontColor & """ SIZE=""" & sFontSize & """>" & _
            vbCrLf & Text & vbCrLf & "</FONT>"

End Function


