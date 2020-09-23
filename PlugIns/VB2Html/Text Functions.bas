Attribute VB_Name = "Text_Functions"
Option Explicit


Public Enum TextArrayOps
    BREACK_ONLY_AFTER_DOTS
End Enum

Function AddLineNumbers(Text As String, ByVal NumStart As Long, ByVal NumStep As Long, ByVal Delimiter As String, NumDigits As Long, IgnoreEmptyLines As Boolean) As String
Dim FormatStr As String
Dim idx As Long
Dim cntr As Long
ReDim sTempArray(1 To 1) As String

FormatStr = String(NumDigits, "0")

Text2Array Text, sTempArray

cntr = NumStart
For idx = LBound(sTempArray) To UBound(sTempArray)
    If IgnoreEmptyLines And sTempArray(idx) = "" Then
        'do nothing
    Else
        sTempArray(idx) = Format(cntr, FormatStr) & Delimiter & sTempArray(idx)
        cntr = cntr + NumStep
    End If
    
Next idx

AddLineNumbers = Join(sTempArray, vbCrLf)

End Function



Function MultiInstrRev(Start As Long, Text As String, LookFor As String, Compare As VbCompareMethod)

Dim iLen As Long
Dim chLookFor As String * 1
Dim idx As Long
Dim iPos As Long
Dim iFirstPos As Long


iLen = Len(LookFor)
iFirstPos = 0

For idx = 1 To iLen
    chLookFor = Mid(LookFor, idx, 1)
    iPos = InStrRev(Text, chLookFor, Start, Compare)
    If (iPos <> 0 And iPos > iFirstPos) Then
         iFirstPos = iPos
'         MsgBox iFirstPos
    End If

Next idx


'If iFirstPos = Len(Text) + 1 Then ' value didn't change / nothing found
'    MultiInstrRev = 0
'    MsgBox "X"
'Else
    MultiInstrRev = iFirstPos
'End If

End Function


Function MultiInstr(Start As Long, Text As String, LookFor As String, Compare As VbCompareMethod)

Dim iLen As Long
Dim chLookFor As String * 1
Dim idx As Long
Dim iPos As Long
Dim iFirstPos As Long


iLen = Len(LookFor)
iFirstPos = Len(Text) + 1 ' out of text boundry // an impossible value

For idx = 1 To iLen
    chLookFor = Mid(LookFor, idx, 1)
    iPos = InStr(Start, Text, chLookFor, Compare)
    If (iPos <> 0 And iPos < iFirstPos) Then
         iFirstPos = iPos
    End If
    
Next idx

If iFirstPos = Len(Text) + 1 Then ' value didn't change / nothing found
    MultiInstr = 0
Else
    MultiInstr = iFirstPos
End If

End Function

Function InsertString(MainStr As String, SubStr As String, Position As Long)
Dim sLeft As String, sRight As String

Select Case Position
    Case Is > Len(MainStr)
        sLeft = MainStr
        sRight = ""
    Case Is <= 0
        sLeft = ""
        sRight = MainStr
    Case Is < Len(MainStr)
        sLeft = Left(MainStr, Position)
        sRight = Right(MainStr, Len(MainStr) - Position)
End Select
    
InsertString = sLeft & SubStr & sRight

End Function

Function DeSpace(sText As String) As String
Dim idx As Integer
Dim sTemp As String
Dim sResult As String
Dim ch As String * 1
Dim InSpaces As Boolean
Dim scount As Integer

sTemp = Trim$(sText)
sResult = ""

scount = 0
For idx = 1 To Len(sTemp)
    ch = Mid(sTemp, idx, 1)
        
    Select Case ch
        Case " "
        scount = scount + 1
        Case Else
        InSpaces = False
        scount = 0
    End Select
    
    
    If scount > 1 Then
        InSpaces = True
    End If
    
    If Not (InSpaces) Then
        sResult = sResult & ch
    End If
    
Next idx
DeSpace = sResult
End Function



Function TrimChars(Text) As String

Const Pattern = "[a-zA-Z0-9" & vbCrLf & ".,;: ]"

Dim idx As Integer
Dim sTemp As String
Dim char As String * 1


sTemp = ""
    For idx = 1 To Len(Text)
        char = Mid(Text, idx, 1)
        If char Like Pattern Then sTemp = sTemp & char
    Next

TrimChars = sTemp

'0 64,91 96 , 123 255


End Function


Public Function Max(ValA, ValB)

Max = IIf(ValA > ValB, ValA, ValB)

End Function

Function DoArrayOp(sArray() As String, WhichOp As TextArrayOps) As String

ReDim sArray(1 To 1) As String
Dim idx As Integer
Dim sTemp As String


'For idx = LBound(sArray) To UBound(sArray)
'Select Case WhichOp
'Case 0
'    'sArray(idx) = fx(sArray(idx), iMax)
'Next idx
'
'sTemp = Join(sArray, vbCrLf)
'
'DoArrayOp = sTemp

End Function

Public Function DelLeftTo(sMainStr As String, sSubStr As String, boolMatchCase As Boolean, boolInclusive As Boolean) As String

Dim iPos As Integer
If boolMatchCase = True Then
        iPos = InStr(1, sMainStr, sSubStr, vbBinaryCompare)
Else
        iPos = InStr(1, sMainStr, sSubStr, vbTextCompare)
End If


If (iPos = 0) Then
            DelLeftTo = sMainStr
Else
            If boolInclusive = True Then
                        DelLeftTo = Right(sMainStr, Len(sMainStr) - iPos - Len(sSubStr) + 1)
            Else
                        DelLeftTo = Right(sMainStr, Len(sMainStr) - iPos + 1)
            End If
End If

End Function

Public Function DelRightTo(sMainStr As String, sSubStr As String, boolMatchCase As Boolean, boolInclusive As Boolean) As String

Dim iPos As Integer
If boolMatchCase = True Then
        iPos = InStrRev(sMainStr, sSubStr, -1, vbBinaryCompare)
Else
        iPos = InStrRev(sMainStr, sSubStr, -1, vbTextCompare)
End If


If (iPos = 0) Then
            DelRightTo = sMainStr
Else
            If boolInclusive = True Then
                        DelRightTo = Left(sMainStr, iPos - 1)
            Else
                        DelRightTo = Left(sMainStr, iPos + Len(sSubStr) - 1)
            End If
End If

End Function


Function Array2Text(sArray() As String) As String

Array2Text = Join(sArray, vbCrLf)

End Function

Function DelLeft(ByVal TextLine As String, n As Integer) As String

If n <= 0 Then
    DelLeft = TextLine
ElseIf n > Len(TextLine) Then
    DelLeft = ""
Else
    DelLeft = Right(TextLine, Len(TextLine) - n)
End If

End Function

Function DelRight(ByVal TextLine As String, n As Integer) As String

If n <= 0 Then
    DelRight = TextLine
ElseIf n > Len(TextLine) Then
    DelRight = ""
Else
    DelRight = Left(TextLine, Len(TextLine) - n)
End If

End Function

Function FixNewLineChars(ByVal Text As String) As String

Dim LineFeed As String * 1
Dim CarrigeReturn As String * 1
Dim BeepChar As String * 1

LineFeed = Chr(10)
CarrigeReturn = Chr(13)
BeepChar = Chr(7)
Text = Replace(Text, vbCrLf, BeepChar)
Text = Replace(Text, LineFeed, BeepChar)
Text = Replace(Text, CarrigeReturn, BeepChar)
Text = Replace(Text, LineFeed & CarrigeReturn, BeepChar)

FixNewLineChars = Replace(Text, BeepChar, vbCrLf)
End Function


Function InsertString2(ByVal TextLine As String, ByVal StrToInsert As String, ByVal Position As Integer) As String
Dim sLeft As String, sRight As String

If Position > Len(TextLine) Then
    Position = Len(TextLine)
End If

If TextLine = "" Then
    InsertString2 = ""
Else
    
    sLeft = Left(TextLine, Position)
    sRight = Right(TextLine, Len(TextLine) - Position)
    
    InsertString2 = sLeft & StrToInsert & sRight

End If

End Function

Function RemoveBlankLines(sArray() As String) As String
Dim idx As Integer
Dim sTemp As String

For idx = LBound(sArray) To UBound(sArray)
    If sArray(idx) = "" Then
        sArray(idx) = Chr$(7)
    End If
Next idx

sTemp = Join(sArray, vbCrLf)
sTemp = Replace(sTemp, Chr$(7) & vbCrLf, "")

RemoveBlankLines = Replace(sTemp, Chr$(7), "") 'in case one stayed!

End Function

Function SetLineMaxWidth(ByVal TextLine As String, ByVal MaxWidth As Integer) As String

Dim HowManyTimes As Integer
Dim idx As Integer
Dim PosShift As Long
Dim sTemp As String

HowManyTimes = Len(TextLine) \ MaxWidth
PosShift = MaxWidth
If MaxWidth < Len(TextLine) Then
    sTemp = TextLine
    For idx = 1 To HowManyTimes
        sTemp = InsertString(sTemp, vbCrLf, PosShift)
        PosShift = PosShift + MaxWidth + 2
    Next idx
    SetLineMaxWidth = sTemp
Else
    SetLineMaxWidth = TextLine
End If

End Function

Function stringf(Text As String) As String
Dim sTemp As String

sTemp = Replace(Text, "\n", vbCrLf, , , vbTextCompare)
sTemp = Replace(sTemp, "\t", vbTab, , , vbTextCompare)
sTemp = Replace(sTemp, "\\", "\", , , vbTextCompare)

stringf = sTemp

End Function

Function Tab2Spaces(Text As String, NumSpaces As Integer) As String

Dim sTemp As String

If NumSpaces < 1 Then NumSpaces = 1
sTemp = Space$(NumSpaces)

Tab2Spaces = Replace(Text, Chr(9), sTemp)


End Function

Function BreackOnlyAfter(ByVal Text As String, ByVal AfterWhat As String) As String

Dim sTemp As String


sTemp = Replace(Text, vbCrLf, "")
sTemp = Replace(sTemp, ".", "." & vbCrLf)

BreackOnlyAfter = sTemp


End Function


Function TrimSpaces(Text As String, TrimLeft As Boolean, TrimRight As Boolean, TrimTab As Boolean) As String

Dim idx As Integer
ReDim sTempArray(1 To 1) As String

Text2Array Text, sTempArray

For idx = LBound(sTempArray) To UBound(sTempArray)
    If (TrimLeft And TrimRight) Then
        sTempArray(idx) = Trim(sTempArray(idx))
        If TrimTab Then sTempArray(idx) = TrimTabs(sTempArray(idx), True, True)
     ElseIf TrimLeft Then
        sTempArray(idx) = LTrim(sTempArray(idx))
        If TrimTab Then sTempArray(idx) = TrimTabs(sTempArray(idx), True, False)
    ElseIf TrimRight Then
        sTempArray(idx) = RTrim(sTempArray(idx))
       If TrimTab Then sTempArray(idx) = TrimTabs(sTempArray(idx), False, True)
    Else
        'Do Nothing
    End If
Next idx

TrimSpaces = Join(sTempArray, vbCrLf)

End Function


Sub Text2Array(ByVal sText As String, ByRef sArray() As String)
    ' sText should not be Empty:
    ' Check for it in the calling routine.
    
    Dim vTmpArray As Variant
    Dim idx As Integer
    
    vTmpArray = Split(sText, vbCrLf)
    ReDim sArray(LBound(vTmpArray) To UBound(vTmpArray))
    
    For idx = LBound(vTmpArray) To UBound(vTmpArray)
    
        sArray(idx) = vTmpArray(idx)
'        MsgBox sArray(idx)
    Next idx

End Sub
Function TrimTabs(ByVal TextLine As String, TrimLeft As Boolean, TrimRight As Boolean) As String
Dim idx As Integer
Dim ch As String * 1

If (TrimLeft And TrimRight) Then
    'Trim both
    For idx = 1 To Len(TextLine)
        ch = Mid(TextLine, idx, 1)
        If ch = Chr(9) Then 'TAB
               Mid(TextLine, idx) = Chr(7)
        Else
                Exit For
        End If
    Next idx
    For idx = Len(TextLine) To 1 Step -1
        ch = Mid(TextLine, idx, 1)
        If ch = Chr(9) Then 'TAB
               Mid(TextLine, idx) = Chr(7)
        Else
                Exit For
        End If
    Next idx

ElseIf TrimLeft Then
    For idx = 1 To Len(TextLine)
        ch = Mid(TextLine, idx, 1)
        If ch = Chr(9) Then 'TAB
               Mid(TextLine, idx) = Chr(7)
        Else
                Exit For
        End If
    Next idx

ElseIf TrimRight Then
    For idx = Len(TextLine) To 1 Step -1
        ch = Mid(TextLine, idx, 1)
        If ch = Chr(9) Then 'TAB
               Mid(TextLine, idx) = Chr(7)
        Else
                Exit For
        End If
    Next idx

End If

TrimTabs = Replace(TextLine, Chr(7), "")

End Function


