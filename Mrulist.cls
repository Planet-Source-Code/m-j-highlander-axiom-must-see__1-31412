VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMRUList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private TheArray() As String

Private ml_Count As Long
Private ml_MaxCount As Long

Public Property Get MaxCount() As Long
       MaxCount = ml_MaxCount
End Property

Public Property Let MaxCount(ByVal lNewValue As Long)
       
       If lNewValue <= 0 Then lNewValue = 1
       ml_MaxCount = lNewValue
       ReDim TheArray(0 To MaxCount - 1)
       Count = 0

End Property

Public Property Get Item(Index As Long) As String
       
If Index < 0 Or Index > MaxCount Then
    Err.Raise 9 'subscript out o'range!
Else
       Item = TheArray(Index)
End If

End Property

Private Function IndexOf(ByRef StrArray() As String, ByVal Element As String) As Long
Dim idx As Long, lb As Long, ub As Long
Dim iTmp As Long

lb = LBound(StrArray)
ub = UBound(StrArray)

iTmp = -1

For idx = lb To ub
    If LCase$(StrArray(idx)) = LCase$(Element) Then
        iTmp = idx
        Exit For
    End If
Next idx

IndexOf = iTmp

End Function

Private Sub MoveToFirst(ByRef StrArray() As String, ByVal Index As Long)

Dim idx As Long, lb As Long, ub As Long
Dim sTmp As String

lb = LBound(StrArray)
ub = UBound(StrArray)

sTmp = StrArray(Index)

For idx = Index To (lb + 1) Step -1
    StrArray(idx) = StrArray(idx - 1)
Next idx

StrArray(lb) = sTmp

End Sub

Private Sub AddFirst(ByRef StrArray() As String, ByVal Element As String)
Dim idx As Long, lb As Long, ub As Long

lb = LBound(StrArray)
ub = UBound(StrArray)

For idx = ub To (lb + 1) Step -1
    StrArray(idx) = StrArray(idx - 1)
Next idx
StrArray(lb) = Element

End Sub

Public Property Get Count() As Long
       Count = ml_Count
End Property

Public Property Let Count(ByVal lNewValue As Long)
       ml_Count = lNewValue
End Property

Public Sub Add(ByVal Element As String)

If Trim$(Element) = "" Then Exit Sub

If IsUnique(TheArray, Element) Then
    AddFirst TheArray, Element
    If Count < MaxCount Then Count = Count + 1

Else
    MoveToFirst TheArray, IndexOf(TheArray, Element)
End If

End Sub
Private Function IsUnique(sArray() As String, Element As String) As Boolean
Dim idx As Long, lb As Long, ub As Long
Dim bAlreadyExists  As Boolean

lb = LBound(sArray)
ub = UBound(sArray)

bAlreadyExists = False

For idx = lb To ub
    If LCase$(sArray(idx)) = LCase$(Element) Then
        bAlreadyExists = True
        Exit For
    End If
Next idx

IsUnique = Not bAlreadyExists

End Function

Private Sub Class_Initialize()

MaxCount = 5

End Sub


