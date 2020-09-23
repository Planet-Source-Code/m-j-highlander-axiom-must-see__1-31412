Attribute VB_Name = "VBC2Html"
Option Explicit

Function DoCode(ByVal Text As String) As String

Dim sNoStrs As String, sNoComments As String
ReDim sStrings(1 To 1) As String
ReDim sComments(1 To 1) As String
Dim idx As Long, cntr As Long
Dim StartPos As Long
Dim sTempCharsX As String

'StripStrings Text, sNoStrs, sStrings()
'StripComments sNoStrs, sNoComments, sComments()
'sNoComments = DoKeyWords(sNoComments)

StripComments Text, sNoComments, sComments()
StripStrings sNoComments, sNoStrs, sStrings()
sNoComments = DoKeyWords(sNoStrs)
   
For cntr = LBound(sComments) To UBound(sComments)
    sTempCharsX = String(cntr, Chr$(5))
    sComments(cntr) = "<FONT COLOR=""GREEN"">" & sComments(cntr) & "</FONT>"
    sNoComments = Replace(sNoComments, sTempCharsX, sComments(cntr), 1, 1)
Next cntr

For cntr = LBound(sStrings) To UBound(sStrings)
    sTempCharsX = String(cntr, Chr$(7))
    sStrings(cntr) = "<FONT COLOR=""#999900"">" & Replace(sStrings(cntr), """", "&quot;") & "</FONT>"
    sNoComments = Replace(sNoComments, sTempCharsX, sStrings(cntr), 1, 1)
Next cntr

sNoComments = Replace(sNoComments, " ( ", "(")
sNoComments = Replace(sNoComments, " ) ", ")")
sNoComments = Replace(sNoComments, " , ", ",")

DoCode = sNoComments

End Function

Function DoKeyWords(Text As String) As String
Dim sTemp As String
Dim idx As Long, odx As Long
Dim v As Variant
ReDim sLines(1 To 1) As String

sTemp = Text

Text2Array sTemp, sLines

For odx = LBound(sLines) To UBound(sLines)
    sLines(odx) = Replace(sLines(odx), ")", " ) ")
    sLines(odx) = Replace(sLines(odx), "(", " ( ")
    sLines(odx) = Replace(sLines(odx), ",", " , ")
    
'    MsgBox sLines(odx)
'    v = ""
    v = Split(sLines(odx), " ")
'    MsgBox LBound(v) & UBound(v)
    If UBound(v) <> -1 Then
'
             For idx = LBound(v) To UBound(v)
                 Select Case LCase(v(idx))
                     Case "if", "then", "else", "dim", "redim", "end", "function", _
                          "sub", "select", "case", "for", "to", "next", "do", "loop", _
                          "while", "until", "long", "integer", "string", _
                          "doevents", "private", "declare", "byval", "byref", "call", _
                          "msgbox", "enum", "public", "private", "const", "type", _
                          "option", "explicit", "boolean", "single", "begin", _
                          "as", "let", "get", "property", "raisevent", "event", "double", _
                          "object", "control", "new", "preserve", "open", _
                          "binary", "input", "output", "random", "seek", "put", "close", _
                          "exit"
                          
                          
                         v(idx) = "<FONT COLOR=""BLUE"">" & v(idx) & "</FONT>"
                     Case "true", "false", "and", "or", "not", "xor"
                          v(idx) = "<FONT COLOR=""PURPLE"">" & v(idx) & "</FONT>"
                     Case Else
                 End Select
             Next idx
             sLines(odx) = Join(v, " ")

'            MsgBox sLines(odx)
    End If
Next odx

DoKeyWords = Array2Text(sLines)

End Function

Sub StripComments(ByVal InText As String, ByRef OutText As String, ByRef sComments() As String)

Dim sTemp As String
ReDim sLines(1 To 1) As String
Dim idx As Long, cntr As Long
Dim StartPos As Long
Dim sq As String
Dim sTempCharsX As String

cntr = 1
sq = "'" 'Single Qoute
sTemp = InText

Text2Array sTemp, sLines

For idx = LBound(sLines) To UBound(sLines)
    StartPos = InStr(1, sLines(idx), sq)
    If StartPos <> 0 Then
        sComments(cntr) = Right(sLines(idx), Len(sLines(idx)) - StartPos + 1)
        'MsgBox sComments(cntr)
        sTempCharsX = String(cntr, Chr$(5))
        sLines(idx) = Replace(sLines(idx), sComments(cntr), sTempCharsX, 1, 1) 'replace only once
        cntr = cntr + 1
        ReDim Preserve sComments(1 To cntr)
    End If
Next idx

If cntr = 1 Then
    'do nothing
Else
    ReDim Preserve sComments(1 To cntr - 1)  ' Kill the extra cell
End If

sTemp = Array2Text(sLines)
OutText = sTemp

End Sub

Sub StripStrings(ByVal InText As String, ByRef OutText As String, ByRef sStrings() As String)

Dim sTemp As String
Dim idx As Long, cntr As Long
Dim StartPos As Long
Dim EndPos As Long
Dim q As String
Dim sTempCharsX As String


q = """" 'Qoute
sTemp = InText
cntr = 1
EndPos = 0

Do
    'StartPos = InStr(EndPos + 1, sTemp, q)
    StartPos = InStr(1, sTemp, q)
    If StartPos = 0 Then Exit Do
    'EndPos = InStr(StartPos + 1, sTemp, q)
    EndPos = InStr(StartPos + 1, sTemp, q)
    If EndPos = 0 Then Exit Do
    
    sStrings(cntr) = Mid(sTemp, StartPos, EndPos - StartPos + 1)
    'MsgBox sStrings(cntr)
    sTempCharsX = String(cntr, Chr$(7))
    sTemp = Replace(sTemp, sStrings(cntr), sTempCharsX, 1, 1) 'replace only once
    cntr = cntr + 1
    ReDim Preserve sStrings(1 To cntr)
Loop

If cntr = 1 Then
    'do nothing
Else
    ReDim Preserve sStrings(1 To cntr - 1)  ' Kill the extra cell
End If
OutText = sTemp

End Sub

