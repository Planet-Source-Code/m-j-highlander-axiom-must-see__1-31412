Attribute VB_Name = "RegExp_Functions"
Option Explicit

Private Const Quote = """"
'Private Const ALL_SPECIAL_CHARS = "[\s" & Quote & "> !#$%&'\(\)\*\+,\-\./:;=\?@\[\]\^_`{\|}~\\]"

Public Function RX_GenericReplace(ByVal Text As String, ByVal Pattern As String, ByVal ReplaceWith As String) As String

Dim m As Match
Dim objRegExp As RegExp

Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = Pattern

'For Each m In objRegExp.Execute(Text)
'    MsgBox m.Value
'Next

RX_GenericReplace = objRegExp.Replace(Text, ReplaceWith)

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function
Public Function RX_ExtractHREFs(ByVal Html As String) As String
Dim SC As CStrCat
Dim sImgFile As String
Dim m As Match
Dim objRegExp As RegExp

Set SC = New CStrCat
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = "< ?A[^\v]*?HREF=""([^\v]*?)""[^\v]*?>"

SC.MaxLength = Len(Html)
For Each m In objRegExp.Execute(Html)
    'MsgBox m.Value
    'MsgBox m.SubMatches(0)
    SC.AddStr m.SubMatches(0) & vbCrLf
Next

RX_ExtractHREFs = SC 'default value

'Overkill, it will go out of scope anyway.
Set SC = Nothing
Set objRegExp = Nothing

End Function
Public Function RX_ValidateImageTags(ByVal Html As String) As String
Dim sImgFile As String
Dim m As Match
Dim objRegExp As RegExp

Set objRegExp = New RegExp
objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = "< ?IMG[^\v]*?SRC=""([^\v]*?)""[^\v]*?>"

For Each m In objRegExp.Execute(Html)
    'MsgBox m.Value
    sImgFile = CurrentDir() & "\" & Replace$(m.SubMatches(0), "/", "\")
    If FileExists(sImgFile) = False Then 'file not found,remove <IMG...>
        Html = Replace$(Html, m.Value, "", 1, 1, vbTextCompare)
    End If
Next

RX_ValidateImageTags = Html

End Function

Public Function RX_RemoveMultipleSpaces(ByVal Text As String) As String
Dim RegEx As RegExp

Set RegEx = New RegExp
RegEx.Pattern = " {2,}"
RegEx.MultiLine = True
RegEx.Global = True

RX_RemoveMultipleSpaces = RegEx.Replace(Text, " ")

End Function
Public Function RX_RemoveTagAttrPath(ByVal Html As String, TagAttr As String, LocalOnly As Boolean) As String
Dim objRegExp As RegExp
Dim m As Match

Set objRegExp = New RegExp
objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "(" & TagAttr & " ?= ?""[^\v]*?"")" 'optional spaces

For Each m In objRegExp.Execute(Html)
    If LocalOnly = False Then
        Html = Replace(Html, m.Value, RemovePath(m.Value), 1, 1)
    Else
        If IsURLLocal(m.Value) Then
            Html = Replace(Html, m.Value, RemovePath(m.Value), 1, 1)
        End If
    End If
Next

RX_RemoveTagAttrPath = Html

End Function



Public Function RX_ProcessLink(ByVal Html As String) As String
Dim sTemp As String, CSS_FilePath As String, CSS_Contents As String
Dim m As Match
Dim objRegExp, strOutput

Set objRegExp = New RegExp
objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "< ?link[^\v]*?href ?= ?""(.*?)""[^\v]*?>"

For Each m In objRegExp.Execute(Html)
'    MsgBox m.Value
    CSS_FilePath = m.SubMatches(0) 'the text between qoutes
    CSS_FilePath = CurrentDir & "\" & CSS_FilePath
    CSS_Contents = GetTextFileContents(CSS_FilePath)
    If CSS_Contents <> "" Then
        CSS_Contents = Make_CSS_Style(CSS_Contents)
    End If
'    MsgBox CSS_Contents
    Html = Replace$(Html, m.Value, CSS_Contents, 1, 1)
Next

RX_ProcessLink = Html

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function
Public Function RX_RemoveHREFPath(ByVal Html As String) As String
Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = "(HREF ?= ?""[^\v]*?"")" 'optional spaces

Dim m As Match
For Each m In objRegExp.Execute(Html)
'    MsgBox m.Value
    Html = Replace(Html, m.Value, RemovePath(m.Value), 1, 1)
Next

RX_RemoveHREFPath = Html

End Function

Public Function RX_GenericExtract(ByVal Text As String, ByVal Pattern As String) As String
'Dim sTemp As String
Dim SC As CStrCat
Dim m As Match
Dim objRegExp As RegExp

Set SC = New CStrCat
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = Pattern

SC.MaxLength = Len(Text)

For Each m In objRegExp.Execute(Text)
'    MsgBox m.Value
'    sTemp = sTemp & m.Value & vbCrLf
    SC.AddStr m.Value & vbCrLf
Next

RX_GenericExtract = SC 'default value

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing
Set SC = Nothing

End Function

Public Function RX_ExtractURLs(ByVal Html As String) As String
Dim sTemp As String

Dim objRegExp, strOutput
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "((ht|f)tp://w?w?w?\.?.*?\..*?)[""\s<>]"

Dim m
For Each m In objRegExp.Execute(Html)
    'MsgBox m.Value
    sTemp = sTemp & m.SubMatches(0) & vbCrLf
Next

RX_ExtractURLs = sTemp

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function

Public Function RX_ChangeFontSize(ByVal Html As String, ByVal NewFontSize As Byte) As String
Dim Size  As String
Dim objRegExp, strOutput
Set objRegExp = New RegExp

Size = Quote & CStr(NewFontSize) & Quote

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "<FONT[^\v]*?SIZE=(""?[1-7]""?)"

Dim m
For Each m In objRegExp.Execute(Html)
'    MsgBox m.Value
'    Html = Replace(Html, m.Value, RemovePath(m.Value), 1, 1)
Next


RX_ChangeFontSize = Html

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function


Public Function RX_ChangeFont(ByVal Html As String, ByVal NewFont As String) As String

Dim objRegExp, strOutput
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "FACE=""?[^ ]*""?" 'anything but spaces,quotes optional


NewFont = "FACE=" & Quote & NewFont & Quote
RX_ChangeFont = objRegExp.Replace(Html, NewFont)

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function

Public Function RX_(ByVal Html As String) As String

Dim objRegExp, strOutput
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = ""

'Dim m
'For Each m In objRegExp.Execute(Html)
'    MsgBox m.Value
'   sTemp = sTemp & vbCrLf & m.Value
'Next

RX_ = objRegExp.Replace(Html, "")

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function



Public Function RX_RemoveBACKGROUNDPath(ByVal Html As String) As String
Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = "(BACKGROUND ?= ?""[^\v]*?"")" 'optional spaces

Dim m As Match
For Each m In objRegExp.Execute(Html)
'    MsgBox m.Value
    Html = Replace(Html, m.Value, RemovePath(m.Value), 1, 1)
Next

RX_RemoveBACKGROUNDPath = Html

End Function


Public Function RX_CompactBlankLines(ByVal Html As String) As String

Dim objRegExp, strOutput
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "(\r\n){3,}" ' match 3 or more CRLF

RX_CompactBlankLines = objRegExp.Replace(Html, vbCrLf & vbCrLf)

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function


Public Function RX_RemoveBlankLines(ByVal Html As String) As String

Dim objRegExp, strOutput
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "(\r\n){2,}" ' match 2 or more CRLF

RX_RemoveBlankLines = objRegExp.Replace(Html, vbCrLf)

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function

Public Function RX_AddBR(ByVal Html As String) As String

Dim objRegExp, strOutput
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "\r\n" ' = CRLF

RX_AddBR = objRegExp.Replace(Html, "<BR>" & vbCrLf)

'//OR, without RegExp (slightly slower):
'RX_AddBR = Replace(HTML, vbCrLf, "<BR>" & vbCrLf)

'Overkill, it will go out of scope anyway.
Set objRegExp = Nothing

End Function


Public Function RX_RemoveSRCPath(ByVal Html As String) As String
Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = "(SRC ?= ?""[^\v]*?"")" 'optional spaces

Dim m As Match
For Each m In objRegExp.Execute(Html)
'    MsgBox m.Value
    Html = Replace(Html, m.Value, RemovePath(m.Value), 1, 1)
Next

RX_RemoveSRCPath = Html

End Function

Function StripHTML(strHTML)
'Strips the HTML tags from strHTML

  Dim objRegExp, strOutput
  Set objRegExp = New RegExp

  objRegExp.IgnoreCase = True
  objRegExp.Global = True
  objRegExp.Pattern = "<(.|\n)+?>"

  'Replace all HTML tag matches with the empty string
  strOutput = objRegExp.Replace(strHTML, "")
  
  StripHTML = strOutput    'Return the value of strOutput

  Set objRegExp = Nothing
End Function


Public Function RX_ExtractTagWithContents(ByVal Html As String, ByVal Tag As String) As String
Dim sTemp As String
Dim objRegExp As RegExp
Dim sOpenTag As String, sCloseTag As String

Tag = Replace$(Tag, "<", "")
Tag = Trim$(Replace$(Tag, ">", ""))
sOpenTag = "<" & Tag
sCloseTag = "</" & Tag & ">"

Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = sOpenTag & "([^\v]*?)" & sCloseTag
Dim m
For Each m In objRegExp.Execute(Html)
    sTemp = sTemp & m.Value & vbCrLf
Next


RX_ExtractTagWithContents = sTemp

End Function

Public Function RX_RemoveTagWithContents(ByVal Html As String, ByVal Tag As String, Optional ByVal TagIsSingle As Boolean = True) As String
Dim objRegExp As RegExp
Dim sOpenTag As String, sCloseTag As String

Tag = Replace$(Tag, "<", "")
Tag = Trim$(Replace$(Tag, ">", ""))
sOpenTag = "<" & Tag
If Not (TagIsSingle) Then
    sCloseTag = "</" & Tag & ">"
Else
    sCloseTag = ">"
End If

Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = sOpenTag & "([^\v]*?)" & sCloseTag
Html = objRegExp.Replace(Html, "")

RX_RemoveTagWithContents = Html

End Function

Public Function RX_RemoveCommentTagAndContent(ByVal Html As String) As String
Dim objRegExp As RegExp
Dim sOpenTag As String, sCloseTag As String

Set objRegExp = New RegExp

sOpenTag = "<!"
sCloseTag = "->"

objRegExp.IgnoreCase = True
objRegExp.Global = True
'BOTH WORK FINE!
'objRegExp.Pattern = sOpenTag & "((>[^<\n\r])|\w|[""\n\r\t\.\(\)\[\]\+:\|&;/,@ =%{<}\?#'!/\-\*])*" & sCloseTag
objRegExp.Pattern = sOpenTag & "(([^\-]>)|\w|[""\n\r\t\.\(\)\[\]\+:\|&;/,@ =%{<}\?#'!/\-\*])*" & sCloseTag
'Dim m
'For Each m In objRegExp.Execute(Html)
'    MsgBox m.Value
'Next

Html = objRegExp.Replace(Html, "")

RX_RemoveCommentTagAndContent = Html

End Function

Public Function RX_RemoveAllTags(Text As String)

Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True


objRegExp.Pattern = "<([^\v]*?)>"
RX_RemoveAllTags = objRegExp.Replace(Text, "")

'For Each ma In objRegExp.Execute(Text1.Text)
'    MsgBox ma.Value
'Next
 
End Function
Public Function RX_RemoveOpenCloseTagKeepContent(ByVal Html As String, ByVal Tag As String) As String
Dim objRegExp As RegExp
Dim sOpenTag As String, sCloseTag As String

Tag = Replace$(Tag, "<", "")
Tag = Trim$(Replace$(Tag, ">", ""))
sOpenTag = "<" & Tag
sCloseTag = "</" & Tag & ">"

Set objRegExp = New RegExp
objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = sOpenTag & "([^\v]*?)>"
Html = objRegExp.Replace(Html, "")

objRegExp.Pattern = sCloseTag
Html = objRegExp.Replace(Html, "")

RX_RemoveOpenCloseTagKeepContent = Html

End Function

