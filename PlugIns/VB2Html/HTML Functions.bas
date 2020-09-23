Attribute VB_Name = "HTML_Functions"
Option Explicit


Type sTag
    href As String
    Text As String
End Type

Public Function DoLinks(Text As String, KeepHTTP As Boolean) As String
'Supported protocols:
' ftp://
' http://
' www.

Dim sTemp As String

sTemp = Text
sTemp = Replace(sTemp, "http://www.", Chr$(7), 1, -1, vbTextCompare)
sTemp = Replace(sTemp, "www.", "http://www.", 1, -1, vbTextCompare)
sTemp = Replace(sTemp, Chr$(7), "http://www.", 1, -1, vbTextCompare)
sTemp = DoHyperLinks(sTemp, "ftp://", KeepHTTP)
sTemp = DoHyperLinks(sTemp, "http://", KeepHTTP)

DoLinks = sTemp

End Function
Function BeautifyLink(HyperLink As String, KeepHTTP As Boolean, SmallCase As Boolean) As String
Dim sTemp As String

If LCase(Left(HyperLink, 7)) <> "http://" Then
    sTemp = HyperLink  ' Not HTTP , so do nothing
Else
    If KeepHTTP Then
            sTemp = HyperLink  ' Keep HTTP , so do nothing
    Else
            sTemp = Right(HyperLink, Len(HyperLink) - 7) ' Remove HTTP://
    End If
End If

If SmallCase Then
    sTemp = LCase(sTemp)
End If

BeautifyLink = sTemp

End Function

Public Function DoStrings(Text As String)
Const Qout = """"
Dim xTemp
Dim sTemp As String
Dim StartPos As Long, EndPos As Long, AtPos As Long
Dim EndChars As String
Dim sTempChars As String
Dim sTempCharsX As String
ReDim CurrentTag(1 To 1) As String
Dim idx As Long
Dim iOpeningPos As Long
Dim iClosingPos As Long
Dim q As String

q = """"

'iOpeningPos = -1 ' Non-Zero value
StartPos = 1  ' Start of Search

sTemp = Text

''++++++++++++++++++++++++++++++++++++++++++++++++++++
EndPos = 1
idx = 1

Do
    StartPos = InStr(EndPos, sTemp, q, vbTextCompare)

    If StartPos = 0 Then Exit Do
    
    EndPos = InStr(StartPos, sTemp, q, vbTextCompare)
    
    CurrentTag(idx) = Mid(sTemp, StartPos, EndPos - StartPos)
   ' MsgBox CurrentTag(idx)
    sTempChars = String(EndPos - StartPos, "X")
    sTempCharsX = String(idx, Chr$(7))
    sTemp = Replace(sTemp, CurrentTag(idx), sTempCharsX, 1, 1) 'replace only once
    idx = idx + 1
    ReDim Preserve CurrentTag(1 To idx)
Loop

If idx = 1 Then
    'do nothing
Else
    ReDim Preserve CurrentTag(1 To idx - 1)  ' Kill the extra cell

    For idx = LBound(CurrentTag) To UBound(CurrentTag)
        sTempCharsX = String(idx, Chr$(7))
        CurrentTag(idx) = "*" & CurrentTag(idx) & "*"
        sTemp = Replace(sTemp, sTempCharsX, CurrentTag(idx), 1, 1)
    Next idx

End If


DoStrings = sTemp

End Function



Function ExtractText(sLine As String) As String
Dim idx As Integer
Dim ch As String * 1
Dim sTemp As String
Dim InTag As Boolean

For idx = 1 To Len(sLine)
    ch = Mid(sLine, idx, 1)
    
    Select Case ch
        Case "<"
          InTag = True
        Case ">"
          InTag = False
        Case Else
        'do nothing
    End Select
    
    
    If Not (InTag) Then
        Select Case ch
            Case ">"
                ch = ""
            Case Chr(13), Chr(10), Chr(9)
                ch = " "
            Case Else
            'do nothing
         End Select
    sTemp = sTemp + ch
    End If

Next idx

ExtractText = DeSpace(sTemp)

End Function


Function ExtractURL(Tag As String) As String
Dim qpos1 As Integer
Dim qpos2 As Integer
Dim hpos As Integer

hpos = InStr(LCase(Tag), "href")
qpos1 = InStr(hpos + 1, Tag, Chr(34))
qpos2 = InStr(qpos1 + 1, Tag, Chr(34))
ExtractURL = LCase(Mid(Tag, qpos1, qpos2 - qpos1 + 1))

End Function

Function FindTags(SourceText As String, LeftTag As String, RightTag As String, TagArray() As sTag)
Dim pos1 As Long
Dim pos2 As Long
Dim CurrentTag As String
Dim idx As Long
Dim lText As String

lText = LCase(SourceText)
LeftTag = "href="
RightTag = "</a>"

    
pos1 = InStr(lText, LeftTag)
idx = 0
ReDim TagArray(1 To 1)
Do While pos1 <> 0

    pos2 = InStr(pos1 + 1, lText, RightTag)
    CurrentTag = Mid(SourceText, pos1, pos2 - pos1 + 4)
    CurrentTag = "<a " + CurrentTag
    pos1 = InStr(pos2 + 1, lText, LeftTag)
    idx = idx + 1
    ReDim Preserve TagArray(1 To idx)
    'TagArray(idx) = CurrentTag
    TagArray(idx).href = ExtractURL(CurrentTag)
    TagArray(idx).Text = ExtractText(CurrentTag)
Loop

End Function

Public Function xHTML_RemoveTag(ByVal sHTML As String, ByVal sOpenTag As String, sCloseTag As String) As String
'Removes a HTML Tag with its content

Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long

'iOpeningPos = -1 ' Non-Zero value
iClosingPos = 1  ' Start of Search

Do
    iOpeningPos = InStr(iClosingPos, sHTML, sOpenTag, vbTextCompare)

If iOpeningPos = 0 Then Exit Do
    
    iClosingPos = InStr(iOpeningPos, sHTML, sCloseTag, vbTextCompare)
    sTempChars = String(iClosingPos - iOpeningPos + Len(sCloseTag), Chr$(7))
    Mid(sHTML, iOpeningPos) = sTempChars
Loop

sHTML = Replace(sHTML, Chr$(7), "")
If sCloseTag <> ">" Then
    xHTML_RemoveTag = Replace(sHTML, sCloseTag, "")
Else
    xHTML_RemoveTag = sHTML
End If

End Function








Function DoHyperLinks(Text As String, Protocol As String, KeepHTTP As Boolean) As String
Const Qout = """"
Dim xTemp
Dim sTemp As String
Dim StartPos As Long, EndPos As Long
Dim EndChars As String
Dim sTempChars As String
Dim sTempCharsX As String
ReDim CurrentTag(1 To 1) As String
Dim idx As Long

'possible endings:
EndChars = " ,)]<" & vbCrLf & vbTab & Qout

sTemp = Text
EndPos = 1
idx = 1
StartPos = 0
Do
    StartPos = InStr(StartPos + 1, sTemp, Protocol, vbTextCompare)

If StartPos = 0 Then Exit Do

    EndPos = MultiInstr(StartPos, sTemp, EndChars, vbTextCompare)
    CurrentTag(idx) = Mid(sTemp, StartPos, EndPos - StartPos)
    sTempChars = String(EndPos - StartPos, "X")
    sTempCharsX = String(idx, Chr$(7))
    'Mid(sTemp, StartPos, EndPos - StartPos) = sTempChars
    'sTemp = InsertString(sTemp, "<B>" & CurrentTag(idx - 1) & "</B>", StartPos - 1)
    sTemp = Replace(sTemp, CurrentTag(idx), sTempCharsX, 1, 1) 'replace only once
    idx = idx + 1
    ReDim Preserve CurrentTag(1 To idx)
Loop

If idx = 1 Then
    'do nothing
Else
    ReDim Preserve CurrentTag(1 To idx - 1)  ' Kill the extra cell
End If

For idx = LBound(CurrentTag) To UBound(CurrentTag)
    sTempCharsX = String(idx, Chr$(7))
    CurrentTag(idx) = "<A HREF=" & Qout & CurrentTag(idx) & Qout & ">" & BeautifyLink(CurrentTag(idx), KeepHTTP, False) & "</A>"
    sTemp = Replace(sTemp, sTempCharsX, CurrentTag(idx), 1, 1)
Next idx
 
 
DoHyperLinks = sTemp

End Function

Function AddBR(sText As String, bPre As Boolean) As String
'Adds <BR> tags, and Handles < > & "

Dim sTemp As String
Dim idx As Long
ReDim blines(1 To 1) As String
sTemp = sText
Text2LinesEx sTemp, blines()

sTemp = ""

'REPLACE <  and  >
For idx = LBound(blines) To UBound(blines)
    blines(idx) = Replace(blines(idx), "&", "&amp;")
    blines(idx) = Replace(blines(idx), Chr$(34), "&quot;")
    blines(idx) = Replace(blines(idx), "<", "&lt;")
    blines(idx) = Replace(blines(idx), ">", "&gt;")
Next idx

'ARRAY --> TEXT

If bPre Then
    sTemp = Join(blines(), vbCrLf)
Else
    sTemp = Join(blines(), "<BR>" & vbCrLf)
End If

AddBR = sTemp

End Function


Public Sub Text2LinesEx(Text As String, Lines() As String)
' check if Text is Empty BEFORE calling this sub.

Dim vTemp As Variant
Dim lLBound As Long
Dim lUBound As Long

vTemp = Split(Text, vbCrLf)
lLBound = LBound(vTemp)
lUBound = UBound(vTemp)
ReDim Lines(lLBound To lUBound)

Lines = vTemp

End Sub

Function RevRGB(ByVal hexRGB As String) As String
Dim var1 As String
Dim var2 As String
Dim Var3 As String

var1 = Left(hexRGB, 2)
var2 = Mid(hexRGB, 3, 2)
Var3 = Right(hexRGB, 2)

RevRGB = Var3 & var2 & var1

End Function


Public Function HTMLize(Text As String, _
                        PageTitle As String, _
                        PicturePath As String, _
                        PageBackColor As String, _
                        TextFontName As String, _
                        TextColor As String, _
                        TextSize As String, _
                        CopyPicture As Boolean, _
                        BackScroll As Boolean, _
                        TextBold As Boolean, _
                        PreserveSpaces As Boolean, _
                        KeepHTTP As Boolean _
                        ) As String
'// URLization NOT implemented Tet :-(

Dim sHTML As String
Dim sHead As String
Dim sBody As String
Dim sBGPic As String
Dim sBoldOpen As String, sBoldClose As String
Dim sPreOpen As String, sPreClose As String
Dim sFont As String, sBGColor As String, sTextColor As String
Dim sBGScrollable As String

sHead = "<HEAD>" & vbCrLf
sHead = sHead & "<TITLE>" & PageTitle & "</TITLE>" & vbCrLf & "</HEAD>" & vbCrLf

sFont = "<FONT FACE=" & Chr(34) & TextFontName & Chr(34) & " SIZE=" & TextSize & ">" & vbCrLf

If PicturePath = "" Then
    sBGPic = ""
Else
'    If chkCopy.Value = vbChecked Then
'        sPicFile = ExtractFileName(Trim(txtBGPic.Text))
'        sTgtDir = ExtractDirName(sTgtFile)
'        On Error Resume Next
'        FileCopy Trim(txtBGPic.Text), sTgtDir & sPicFile
'        If Err Then
'            sCopyResult = vbCrLf & "Couldn't copy " & Chr(34) & UCase(sPicFile) & Chr(34)
'        Else
'            sCopyResult = vbCrLf & Chr(34) & UCase(sPicFile) & Chr(34) & " was copied successfully."
'        End If
'        On Error GoTo 0
'    Else
'        sPicFile = Trim(txtBGPic.Text)
'    End If
    sBGPic = " BACKGROUND=" & Chr$(34) & PicturePath & Chr$(34)
End If

If TextBold Then
    sBoldOpen = "<B>" & vbCrLf
    sBoldClose = "</B>" & vbCrLf
Else
    sBoldOpen = ""
    sBoldClose = ""
End If

If PreserveSpaces Then
    sPreOpen = vbCrLf & "<PRE>" & vbCrLf
    sPreClose = vbCrLf & "</PRE>" & vbCrLf
    sHTML = AddBR(Text, True)
Else
    sPreOpen = ""
    sPreClose = ""
    sHTML = AddBR(Text, False)
End If

sHTML = DoLinks(sHTML, KeepHTTP) ' http://  ftp://  www. (will ad http:// to it | IS IT A BUG?)

If BackScroll Then
    sBGScrollable = ""
Else
    sBGScrollable = " BGPROPERTIES = FIXED "
End If

sBody = "<BODY BGCOLOR=" & PageBackColor & " TEXT=" & TextColor & sBGPic & sBGScrollable & ">" & vbCrLf
sBody = sBody & sPreOpen & sFont & sBoldOpen


sHTML = "<HTML>" & vbCrLf & sHead & sBody & sHTML & sBoldClose & "</FONT>" & sPreClose & "</BODY>" & vbCrLf & "</HTML>"

HTMLize = sHTML

End Function

Function ColorToHex(ByVal lColor As Long) As String
Dim sTemp As String

sTemp = Hex$(lColor)

If Len(sTemp) < 6 Then sTemp = String(6 - Len(sTemp), "0") + sTemp
sTemp = Chr(34) & "#" & RevRGB(sTemp) & Chr(34)

ColorToHex = sTemp

End Function


Public Function HTML_RemoveTag(ByVal sHTML As String, ByVal sOpenTag As String, sCloseTag As String) As String
'Removes a HTML Tag with its content

Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long

'iOpeningPos = -1 ' Non-Zero value
iClosingPos = 1  ' Start of Search

Do
    iOpeningPos = InStr(iClosingPos, sHTML, sOpenTag, vbTextCompare)

If iOpeningPos = 0 Then Exit Do
    
    iClosingPos = InStr(iOpeningPos, sHTML, sCloseTag, vbTextCompare)
    sTempChars = String(iClosingPos - iOpeningPos + Len(sCloseTag), Chr$(7))
    Mid(sHTML, iOpeningPos) = sTempChars
Loop

sHTML = Replace(sHTML, Chr$(7), "")
If sCloseTag <> ">" Then
    HTML_RemoveTag = Replace(sHTML, sCloseTag, "")
Else
    HTML_RemoveTag = sHTML
End If

End Function


Public Function HTML_RemoveScripts(ByVal sHTML As String) As String
Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long

'iOpeningPos = -1 ' Non-Zero value
iClosingPos = 1  ' Start of Search

Do
    iOpeningPos = InStr(iClosingPos, sHTML, "<script", vbTextCompare)
'MsgBox iOpeningPos

If iOpeningPos = 0 Then Exit Do
    
    iClosingPos = InStr(iOpeningPos, sHTML, "</script>", vbTextCompare)
        
'MsgBox iClosingPos
    sTempChars = String(iClosingPos - iOpeningPos + 9, Chr$(7))
    Mid(sHTML, iOpeningPos) = sTempChars
Loop

HTML_RemoveScripts = Replace(sHTML, Chr$(7), "")

End Function



Public Function HTML_RemoveIFrameTags(ByVal sHTML As String) As String
Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long

'iOpeningPos = -1 ' Non-Zero value
iClosingPos = 1  ' Start of Search

Do
    iOpeningPos = InStr(iClosingPos, sHTML, "<iframe", vbTextCompare)
'MsgBox iOpeningPos

If iOpeningPos = 0 Then Exit Do
    
    iClosingPos = InStr(iOpeningPos, sHTML, ">", vbTextCompare)
        
'MsgBox iClosingPos
    sTempChars = String(iClosingPos - iOpeningPos + 1, Chr$(7))
    Mid(sHTML, iOpeningPos) = sTempChars
Loop

sHTML = Replace(sHTML, Chr$(7), "")
HTML_RemoveIFrameTags = Replace(sHTML, "</iframe>", "")
End Function

Public Function HTML_RemoveImageTags(ByVal sHTML As String) As String
Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long

'iOpeningPos = -1 ' Non-Zero value
iClosingPos = 1  ' Start of Search

Do
    iOpeningPos = InStr(iClosingPos, sHTML, "<img", vbTextCompare)
'MsgBox iOpeningPos

If iOpeningPos = 0 Then Exit Do
    
    iClosingPos = InStr(iOpeningPos, sHTML, ">", vbTextCompare)
        
'MsgBox iClosingPos
    sTempChars = String(iClosingPos - iOpeningPos + 1, Chr$(7))
    Mid(sHTML, iOpeningPos) = sTempChars
Loop

HTML_RemoveImageTags = Replace(sHTML, Chr$(7), "")

End Function


