Attribute VB_Name = "AxiomMain"
''''''''''''''[ Program Execution Starts Here ]''''''''''''''
Option Explicit

Public Type TagInfo
    Name As String
    IsSingle As Boolean
    Description As String
End Type


' Check the "Property" Procs to see how Properties
' are superior to Public Vars
Private ms_CurrentDir As String
Private ms_CurrentFile As String

Private mb_InputIsDirty As Boolean
Private mb_OutputIsDirty As Boolean
Private mb_PadIsDirty As Boolean


Public Enum WhichTextbox
    All_TextBoxes = 0
    Input_TextBox = 1
    Output_TextBox = 2
    Pad_TextBox = 3
    
End Enum

Public gs_FindWhat As String
Public gl_Options As Long
Public gl_Pos As Long

Public AxiomSettings As CAxiomSettings
Public go_HTMLTags As CHTMLTags
Public go_MRU As CMRUList
Public go_PlugIns As CPlugIns
Public Property Get IsDirty(ByVal Index As WhichTextbox) As Boolean
        
Select Case Index
    
    Case Input_TextBox
        IsDirty = mb_InputIsDirty
    
    Case Output_TextBox
        IsDirty = mb_OutputIsDirty
    
    Case Pad_TextBox
        IsDirty = mb_PadIsDirty
    
    Case All_TextBoxes 'Not to be used really
        IsDirty = mb_InputIsDirty Or mb_OutputIsDirty Or mb_PadIsDirty
        
End Select
        
End Property

Public Property Let IsDirty(ByVal Index As WhichTextbox, ByVal bNewValue As Boolean)
       
    Select Case Index
        
        Case Input_TextBox
            mb_InputIsDirty = bNewValue
        
        Case Output_TextBox
            mb_OutputIsDirty = bNewValue
        
        Case Pad_TextBox
            mb_PadIsDirty = bNewValue
        
        Case All_TextBoxes ' Set All
            mb_InputIsDirty = bNewValue
            mb_OutputIsDirty = bNewValue
            mb_PadIsDirty = bNewValue
            
    End Select
       
End Property

Public Property Get CurrentFile() As String
       CurrentFile = ms_CurrentFile
End Property

Public Property Let CurrentFile(ByVal sNewValue As String)
       ms_CurrentFile = sNewValue
End Property

Public Sub Main() '<-----------------[Program Execution Starts Here]

Set go_HTMLTags = New CHTMLTags
Set go_MRU = New CMRUList
Set AxiomSettings = New CAxiomSettings
Set go_PlugIns = New CPlugIns

go_PlugIns.PlugInsFolder = (RemoveSlash(App.Path) & "\PlugIns")
go_HTMLTags.LoadFromFile (RemoveSlash(App.Path) & "\html_tags.ini")

AxiomSettings.LoadSettings

Load frmAxiomMain

' Handle Command Line:
If FileExists(Command$) Then
    frmAxiomMain.MainText.LoadFile Command$, rtfText
    IsDirty(Input_TextBox) = False
    frmAxiomMain.cdlg.InitDir = ExtractDirName(Command$)
    CurrentDir = Command$  ' Public Property
    CurrentFile = Command$
End If

frmAxiomMain.Show

End Sub
Public Property Get CurrentDir() As String
    
    CurrentDir = ms_CurrentDir
    
End Property

Public Property Let CurrentDir(ByVal sNewValue As String)
    
    ms_CurrentDir = RemoveSlash(ExtractDirName(sNewValue))
    
End Property
