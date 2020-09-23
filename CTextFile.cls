VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Enum FileOpenMode

        OpenForInput = 0
        OpenForOutput = 1
        OpenForAppend = 2
        
End Enum

Private m_iFileHandle As Integer
Private m_sFileName As String


Public Function CreateTempFile() As String
'Create a temp file and open for output
'Return_Value: Created file name (can be used in deleting it)


Dim sBaseName As String
Dim sTempDir As String
Dim sTempFile As String
Dim iRand As Integer

Randomize
iRand = CInt(Rnd * 9999)
sBaseName = "~ctxtmp@" & Format(iRand, "0000")
sTempDir = Environ("Temp")

If Right(sTempDir, 1) <> "\" Then
        sTempDir = sTempDir & "\"
End If
sTempFile = sTempDir & sBaseName & ".tmp"
sTempFile = LCase(Trim(sTempFile))

'Open file for output
FileOpen sTempFile, OpenForOutput

'Return()
CreateTempFile = sTempFile

End Function

Public Property Get FileName() As String
        FileName = LCase$(m_sFileName)
End Property

Private Property Let FileName(ByVal sNewValue As String)
'This is a Private method, making "FileName" Read-Only outside Class

        m_sFileName = LCase$(sNewValue)
        
End Property

Public Sub FileOpen(sFileName As String, OpenMode As FileOpenMode)
Dim iFileNum As Integer


iFileNum = FreeFile

Select Case OpenMode
        Case OpenForInput
                Open sFileName For Input As iFileNum
        Case OpenForOutput
                Open sFileName For Output As iFileNum
        Case OpenForAppend
                Open sFileName For Append As iFileNum
        Case Else
        'dunno!
End Select

'Assign to properties
FileHandle = iFileNum
FileName = sFileName

End Sub


Public Function FileClose()
        
        Close FileHandle
        ' Invalidate File Handle and File Name
        FileHandle = vbEmpty
        FileName = vbNullString
        
End Function

Public Sub Flush()
'Force writing to file without closing it

Dim lOpenMode As Long
Dim sTmpName As String

'Get current open mode
lOpenMode = FileAttr(FileHandle, 1)

'Save name
sTmpName = FileName

'Close, thus writing contents
FileClose

'ReOpen:
'NOTE that this might change FILE HANDLE
'this is no problem since it should never be saved in a var

Select Case lOpenMode
        Case 2, 8 '2=output,8=append
                ' BUT we open for_append anyway
                ' coz we want to resume at end of file
                FileOpen sTmpName, OpenForAppend
        '('''''''''Case 32=binary --> CBinFile!!!)
        Case Else
        'flush is invalid for Open_For_Input mode
End Select

        
End Sub

Public Function ReadAll() As String
Dim bArray() As Byte

ReDim bArray(0 To LOF(FileHandle) - 1) As Byte

bArray = Input$(LOF(FileHandle), FileHandle)

ReadAll = bArray  'no need for: StrConv(bArray, vbUnicode)
                  'since file is opened for input not binary (i guess?)
End Function
Public Function ReadLnArray() As Variant
'Returns a ZERO-Based array

Dim sTmp As String
Dim idx As Long
ReDim sTmpArray(0 To 999) As String

idx = 0
Do While Not VBA.EOF(FileHandle)
        Line Input #FileHandle, sTmp
        sTmpArray(idx) = sTmp
        idx = idx + 1
        If idx > UBound(sTmpArray) Then
                'redim at 1000's for max speed
                ReDim Preserve sTmpArray(0 To UBound(sTmpArray) + 1000)
        End If
Loop
' -2=(-1)+(-1) Since it's 0-based, and last idx=idx+1 is not used
ReDim Preserve sTmpArray(0 To idx - 2)
ReadLnArray = sTmpArray

End Function

Public Property Get EOF() As Boolean
' To avoide ambeguity and/or clash we use VBA.EOF()

If VBA.EOF(FileHandle) Then
        EOF = True
Else
        EOF = False
End If

End Property

Public Function ReadLn() As String
Dim sTmp As String

Line Input #FileHandle, sTmp
ReadLn = sTmp

End Function




Public Sub WriteLnStr(Optional ByVal sStrToWrite As String)

        Print #FileHandle, sStrToWrite

End Sub

Public Sub WriteStr(sStrToWrite As String)

        Print #FileHandle, sStrToWrite;

End Sub


Public Property Get FileHandle() As Integer
        FileHandle = m_iFileHandle
End Property

Private Property Let FileHandle(ByVal iNewValue As Integer)
'This is a Private method, making "FileHandle" Read-Only outside Class

        m_iFileHandle = iNewValue
        
End Property
