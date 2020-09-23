VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlugIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Property Vars:
Private ms_PlugInsFolder As String
Private ml_PlugInsCount As Long

'Local Vars:
Private DLLs() As String
'Private ClassNames() As String


Private ms_ClassName(0 To 100) As String
Private ms_FunctionName(0 To 100) As String
Private ms_FunctionDescription(0 To 100) As String
Private ms_FunctionArgs(0 To 100) As String

Private CurIdx As Long

Private Sub ShellSortAsc(SortArray() As String, LinkedArray() As String, UpBound As Long, AllLowerCase As Boolean)
Dim sVal1 As String, sVal2 As String

Dim Row As Long
Dim MaxRow As Long
Dim MinRow As Long
Dim Swtch As Long
Dim Limit As Long
Dim Offset As Long

MaxRow = UpBound
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
            If AllLowerCase Then
                sVal1 = LCase(SortArray(Row))
                sVal2 = LCase(SortArray(Row + Offset))
            Else
                sVal1 = SortArray(Row)
                sVal2 = SortArray(Row + Offset)
            End If
            If sVal1 > sVal2 Then
               Swap SortArray(Row), SortArray(Row + Offset)
               Swap LinkedArray(Row), LinkedArray(Row + Offset)
               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop
End Sub

Private Sub Swap(ByRef var1 As String, ByRef var2 As String)
    Dim x As String
    x = var1
    var1 = var2
    var2 = x
End Sub

Private Function RightTo(ByVal sString As String, ByVal sChar As String) As String
' Returns a sub-string from the rightmost of the first argument
' until a charachter (the second argumennt) is found.
' The charachter is not included in the returned string.
' If the charachter is not found, the entire string will be returned.


Dim iCntr As Integer
Dim iLastPos As Integer
Dim iCurrPos As Integer

Dim ch As String * 1

ch = Left$(sChar, 1)

iLastPos = InStrRev(sString, ch) 'iCurrPos

RightTo = Right$(sString, Len(sString) - iLastPos)

End Function

Private Function LeftTo(ByVal sString As String, ByVal sChar As String) As String
' Returns a sub-string from the leftmost of the first argument
' until a charachter (the second argumennt) is found.
' The charachter is not included in the returned string.
' If the charachter is not found, an empty string will be returned.



Dim iPos As Integer
Dim ch As String * 1

ch = Left$(sChar, 1)
iPos = InStr(sString, ch)
If iPos <> 0 Then
    LeftTo = Left$(sString, iPos - 1)
Else
    LeftTo = ""
End If

End Function

Private Function RemoveSlash(ByVal sPath As String) As String

sPath = Trim(sPath)
If Right$(sPath, 1) = "\" Then
    RemoveSlash = Left(sPath, Len(sPath) - 1)
Else
    RemoveSlash = sPath
End If

End Function

Public Property Get FunctionClass(ByVal Index As Long) As String
       
       FunctionClass = ms_ClassName(Index)
       
End Property
Public Property Get FunctionArgs(ByVal Index As Long) As String
       
       FunctionArgs = ms_FunctionArgs(Index)
       
End Property

Public Property Get FunctionDescription(ByVal Index As Long) As String
       
       FunctionDescription = ms_FunctionDescription(Index)
       
End Property




Private Sub GetFuncs()
Dim idx As Long, cntr As Long
Dim ClassName As String
Dim bResult As Boolean
Dim fn() As String
Dim fh() As String
Dim fd() As String

For idx = 0 To UBound(DLLs)
    bResult = ProcessTypeLibrary(DLLs(idx), ClassName, fn(), fh(), fd())
    If bResult Then
        For cntr = LBound(fn) To UBound(fn)
            ms_ClassName(CurIdx) = ClassName
            If fd(cntr) = "" Then fd(cntr) = fn(cntr)
            ms_FunctionName(CurIdx) = fd(cntr) & vbCrLf & fn(cntr)
            CurIdx = CurIdx + 1
        Next cntr
    End If
Next idx

Count = CurIdx  'COUNT not INDEX

ShellSortAsc ms_FunctionName(), ms_ClassName(), CurIdx - 1, True
For idx = 0 To CurIdx - 1
    ms_FunctionDescription(idx) = LeftTo(ms_FunctionName(idx), vbCrLf)
    If ms_FunctionDescription(idx) = "" Then ms_FunctionDescription(idx) = RemoveCrLf(ms_FunctionName(idx))
    ms_FunctionArgs(idx) = RemoveCrLf(StrBetween(ms_FunctionName(idx), ":", vbCrLf))
    If ms_FunctionArgs(idx) = "" Then ms_FunctionArgs(idx) = "Text"
    ms_FunctionName(idx) = RemoveCrLf(RightTo(ms_FunctionName(idx), vbCrLf))
Next

End Sub

Public Property Get FunctionName(ByVal Index As Long) As String
       
       FunctionName = ms_FunctionName(Index)
       
End Property

Public Property Get Count() As Long
       Count = ml_PlugInsCount
End Property

Private Property Let Count(ByVal lNewValue As Long)
       ml_PlugInsCount = lNewValue
End Property

Public Property Get PlugInsFolder() As String
       PlugInsFolder = ms_PlugInsFolder
End Property

Public Property Let PlugInsFolder(ByVal sNewValue As String)
Dim fname As String
Dim idx As Long

ms_PlugInsFolder = sNewValue
    
ReDim DLLs(0 To 0) As String

fname = Dir$(RemoveSlash(sNewValue) & "\*.dll")
Do
    If fname = "" Then Exit Do
    DLLs(idx) = RemoveSlash(sNewValue) & "\" & fname
    idx = idx + 1
    ReDim Preserve DLLs(0 To idx)
    fname = Dir$()
Loop

If idx > 0 Then
    idx = idx - 1
    ReDim Preserve DLLs(0 To idx)
    GetFuncs
Else
    'idx=0 :
    Count = 0
End If

End Property
Private Function RemoveCrLf(ByVal Text As String) As String
    Text = Replace$(Text, vbCr, "")
    Text = Replace$(Text, vbLf, "")
    RemoveCrLf = Trim$(Text)
End Function

Private Function StrBetween(ByVal Text As String, ByVal StrStart As String, StrStop As String) As String
Dim StartPos As Long
Dim StopPos As Long

StartPos = InStr(Text, StrStart)
StopPos = InStr(StartPos + 1, Text, StrStop)

If StopPos = 0 Then StopPos = Len(Text)
If (StartPos > 0) And (StopPos > StartPos) Then
    StrBetween = Mid$(Text, StartPos + 1, StopPos - StartPos)
Else
    StrBetween = ""
End If

End Function

