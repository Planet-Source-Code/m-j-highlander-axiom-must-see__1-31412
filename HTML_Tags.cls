VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHTMLTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private TagInfoArray(0 To 199) As TagInfo 'kinda kool!
Private ms_Temp As String

Private ml_Count As Long
Private ms_Name As String
Private mb_IsSingle As Boolean
Private ms_Description As String

Public Sub SaveToFile(ByVal Filename As String)
Dim iFF As Long, idx As Long
Dim sTemp As String

If ml_Count < 0 Then Exit Sub

ReDim tmpArray(0 To ml_Count) As String
For idx = 0 To ml_Count
    tmpArray(idx) = Description(idx) & "," & Name(idx) & "," & CStr(Abs(CLng(IsSingle(idx))))
Next idx
sTemp = Join$(tmpArray, vbCrLf)

iFF = FreeFile
Open Filename For Output As #iFF
    Print #iFF, sTemp
Close iFF


End Sub

Private Sub SwapTagInfos(ByRef var1 As TagInfo, ByRef var2 As TagInfo)
    Dim X As TagInfo
    X = var1
    var1 = var2
    var2 = X
End Sub

Private Sub SortTagInfos(ByRef SortArray() As TagInfo)
Dim sVal1 As String, sVal2 As String

Dim Row As Long
Dim MaxRow As Long
Dim MinRow As Long
Dim Swtch As Long
Dim Limit As Long
Dim Offset As Long

MaxRow = ml_Count
MinRow = 0
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False

         For Row = MinRow To Limit
                sVal1 = SortArray(Row).Description
                sVal2 = SortArray(Row + Offset).Description
            If sVal1 > sVal2 Then
               SwapTagInfos SortArray(Row), SortArray(Row + Offset)
               Swtch = Row
            End If
         Next Row

         Limit = Swtch - Offset
      Loop While Swtch

      Offset = Offset \ 2
   Loop
End Sub

Public Sub Add(ByVal TagName As String, ByVal TagDescription As String, ByVal TagIsSingle As Boolean)

ml_Count = ml_Count + 1 ' no redim, static array

TagInfoArray(ml_Count).Name = TagName
TagInfoArray(ml_Count).Description = TagDescription
TagInfoArray(ml_Count).IsSingle = TagIsSingle

SortTagInfos TagInfoArray

End Sub

Public Property Get Description(ByVal idx As Long) As String
'Read Only
    Description = TagInfoArray(idx).Description
    
End Property

Public Property Get IsSingle(ByVal idx As Long) As Boolean
'Read Only
       
       IsSingle = TagInfoArray(idx).IsSingle
       
End Property


Public Property Get Name(ByVal idx As Long) As String
'Read Only

       Name = TagInfoArray(idx).Name
       
End Property


Public Function LoadFromFile(ByVal Filename As String) As Boolean
Dim iFF As Long

On Error GoTo Err_LoadFromFile
    iFF = FreeFile
    Open Filename For Input As #iFF
        ms_Temp = Input$(LOF(iFF), iFF)
    Close iFF
    LoadFromFile = True

Call Parse

Exit Function

Err_LoadFromFile:
    ml_Count = -1
    Err = 0 'Clear the Error , kinda overkill
    MsgBox "Error HTML_TAGS.INI Not Found", vbCritical, "Oops"
    LoadFromFile = False

End Function
Private Function Parse() As String
Dim tmpArray() As String
Dim idx As Long, cntr As Long


tmpArray = Split(ms_Temp, vbCrLf) 'text --> array
'Sort by description, since description is first
'StrSort tmpArray, True, True

idx = -1  ' +1=0
For cntr = LBound(tmpArray) To UBound(tmpArray)
    
    If Len(tmpArray(cntr)) > 0 Then
        idx = idx + 1
        TagInfoArray(idx) = SplitStrStrBool(tmpArray(cntr))
    End If
    
Next cntr
'if nothing found =-1 ,very kool!
ml_Count = idx

'now sort by description
SortTagInfos TagInfoArray

End Function

Private Function SplitStrStrBool(ByVal TextLine As String) As TagInfo
Dim tmpArray() As String
    
    tmpArray = Split(TextLine, ",")
    
    SplitStrStrBool.Description = tmpArray(0)
    SplitStrStrBool.Name = tmpArray(1)
    SplitStrStrBool.IsSingle = CBool(tmpArray(2))


End Function



Public Property Get Count() As Long
'Read Only (No Put Proc)
    
    Count = ml_Count + 1 'Count NOT UBound !
    
End Property

