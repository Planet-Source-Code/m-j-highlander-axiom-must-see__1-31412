Attribute VB_Name = "Sort_String"
Option Explicit

Sub SwapStrings(pbString1 As String, pbString2 As String)
    Dim l_Hold As Long
    CopyMemory l_Hold, ByVal VarPtr(pbString1), 4
    CopyMemory ByVal VarPtr(pbString1), ByVal VarPtr(pbString2), 4
    CopyMemory ByVal VarPtr(pbString2), l_Hold, 4
End Sub

Public Sub ShellSortAsc(SortArray() As String, AllLowerCase As Boolean)
'The fastets sort algorithm!
Dim sVal1 As String, sVal2 As String

Dim Row As Long
Dim MaxRow As Long
Dim MinRow As Long
Dim Swtch As Long
Dim Limit As Long
Dim Offset As Long

MaxRow = UBound(SortArray)
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


Public Sub ShellSortDesc(SortArray() As String, AllLowerCase As Boolean)
'The fastets sort algorithm!
Dim sVal1 As String, sVal2 As String

Dim Row As Long
Dim MaxRow As Long
Dim MinRow As Long
Dim Swtch As Long
Dim Limit As Long
Dim Offset As Long

MaxRow = UBound(SortArray)
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
            If sVal1 < sVal2 Then
               Swap SortArray(Row), SortArray(Row + Offset)
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




Public Static Sub StrSort(Lines() As String, Ascending As Boolean, AllLowerCase As Boolean)

If Ascending Then
    ShellSortAsc Lines(), AllLowerCase
Else
    ShellSortDesc Lines(), AllLowerCase
End If

''''''Old VERY VERY SLOW Code:
'Dim i As Long
'Dim J As Long
'Dim NumInArray, LowerBound As Integer
'NumInArray = UBound(words)
'LowerBound = LBound(words)
'For i = LowerBound To NumInArray
'    J = 0
'    For J = LowerBound To NumInArray
'        If AllLowerCase = True Then
'            If Ascending = True Then
'                If StrComp(LCase(words(i)), _
'                     LCase(words(J))) = -1 Then
'                    Call Swap(words(i), words(J))
'                End If
'            Else
'                If StrComp(LCase(words(i)), _
'                       LCase(words(J))) = 1 Then
'                    Call Swap(words(i), words(J))
'                End If
'            End If
'        Else
'            If Ascending = True Then
'                If StrComp(words(i), words(J)) = -1 Then
'                    Call Swap(words(i), words(J))
'                End If
'            Else
'                If StrComp(words(i), _
'                    words(J)) = 1 Then
'                    Call Swap(words(i), words(J))
'                End If
'            End If
'        End If
'    Next J
'Next i
End Sub

Private Sub Swap(ByRef var1 As String, ByRef var2 As String)
    Dim X As String
    X = var1
    var1 = var2
    var2 = X
End Sub
