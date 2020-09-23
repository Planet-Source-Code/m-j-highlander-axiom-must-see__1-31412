VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VB2C_Class"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function Text2VBCode(Text As String, VarName As String) As String
Attribute Text2VBCode.VB_Description = "Convert A Text Block to a VB Function.\r\nInputs: Text,Function Name\r\n"
Dim sTemp As String
ReDim sArray(1 To 1) As String
Dim idx As Long


sTemp = Text
sTemp = Replace(sTemp, Qoute, "\q")

Text2Array sTemp, sArray
For idx = LBound(sArray) To UBound(sArray)
    sArray(idx) = VarName & "_TmpStr(" & idx & ")=" & Qoute & sArray(idx) & Qoute
    
Next idx

sTemp = "Private Function " & VarName & "() As String" & vbCrLf & _
        "Dim sTemp As String , Qoute As String" & vbCrLf & "Qoute = Chr$(34)" & vbCrLf & _
        "ReDim " & VarName & "_TmpStr(" & LBound(sArray) & " To " & UBound(sArray) & ") As String" & vbCrLf & _
         Join(sArray, vbCrLf) & vbCrLf & _
         "sTemp = Join(" & VarName & "_TmpStr,vbCrLf)" & vbCrLf & _
         VarName & " = Replace(sTemp,""\q"",Qoute,,,vbTextCompare)" & vbCrLf & _
         vbCrLf & "End Function"
         
Text2VBCode = sTemp

End Function
Private Function Qoute()
    Qoute = Chr$(34)
End Function

Private Function EscapeQoute(Text As String) As String
    Dim sTemp As String
    sTemp = Replace(Text, "\q", Qoute, , , vbTextCompare)
    EscapeQoute = sTemp
End Function

Private Function Array2Text(sArray() As String) As String
    Array2Text = Join(sArray, vbCrLf)
End Function

Private Sub Text2Array(ByVal sText As String, ByRef sArray() As String)
    ' sText should not be Empty:
    ' Check for it in the calling routine.
    
    Dim vTmpArray As Variant
    Dim idx As Integer
    
    vTmpArray = Split(sText, vbCrLf)
    ReDim sArray(LBound(vTmpArray) To UBound(vTmpArray))
    
    For idx = LBound(vTmpArray) To UBound(vTmpArray)
        sArray(idx) = vTmpArray(idx)
    Next idx

End Sub



