VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBCode2Html"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function VB2HTML(ByVal Text As String) As String
Attribute VB2HTML.VB_Description = "Convert BASIC Code to Colored HTML\r\nInputs:Text"
' "<PRE><FONT FACE=Verdana Size=2>" &

Text = Replace(Text, ">", "&gt;")
Text = Replace(Text, "<", "&lt;")

VB2HTML = "<PRE><FONT FACE=""FixedSys,Verdana,Tahoma,Arial"" Size=""2"">" & _
          vbCrLf & DoCode(Text) & _
          vbCrLf & "</FONT></PRE>" & vbCrLf

End Function

