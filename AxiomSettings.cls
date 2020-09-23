VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAxiomSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mb_MultipleTextBoxes As Boolean
Private mi_SpacesPerTab As Integer
Private mb_TrimTabsAlso As Boolean
Private mb_WordWrap As Boolean
Private me_CharsToKeep As CharRangeConstants

Private ms_Colors As String
Private ms_TextFont As String

Public Property Get TextFont() As String
       TextFont = ms_TextFont
End Property

Public Property Let TextFont(ByVal sNewValue As String)
       ms_TextFont = sNewValue
End Property

Public Property Get Colors() As String
       Colors = ms_Colors
End Property

Public Property Let Colors(ByVal sNewValue As String)
       ms_Colors = sNewValue
End Property

Public Sub ApplySettings()
Dim sTemp As String, vArray As Variant, idx As Long
With frmAxiomMain
    
    '1) MULTIPLE | SINGLE TextBox:
    If MultipleTextBoxes = True Then
        .Tabs.Tabs.Clear
        .Tabs.Tabs.Add
        .Tabs.Tabs.Item(1).Caption = "Input"
        .Tabs.Tabs.Item(1).Key = "Input"
        .Tabs.Tabs.Add
        .Tabs.Tabs.Item(2).Caption = "Output"
        .Tabs.Tabs.Item(2).Key = "Output"
        .Tabs.Tabs.Add
        .Tabs.Tabs.Item(3).Caption = "Pad"
        .Tabs.Tabs.Item(3).Key = "Pad"
    Else
        .Tabs.Tabs.Clear
        .Tabs.Tabs.Add
        .Tabs.Tabs.Item(1).Caption = "Input"
        .Tabs.Tabs.Item(1).Key = "Input"
        .Tabs.Tabs.Add
        .Tabs.Tabs.Item(2).Caption = "Pad"
        .Tabs.Tabs.Item(2).Key = "Pad"
    End If
    
    'WORD WRAP:
    If AxiomSettings.WordWrap Then
        SetWordWrap frmAxiomMain.MainText, True
        SetWordWrap frmAxiomMain.OutputText, True
        SetWordWrap frmAxiomMain.PadText, True
    Else
        SetWordWrap frmAxiomMain.MainText, False
        SetWordWrap frmAxiomMain.OutputText, False
        SetWordWrap frmAxiomMain.PadText, False
    End If
    
    'COLORS:
    On Error Resume Next
        sTemp = Colors
        vArray = Split(sTemp, ",")
        frmAxiomMain.MainText.BackColor = CLng(vArray(LBound(vArray)))
        frmAxiomMain.OutputText.BackColor = CLng(vArray(LBound(vArray) + 1))
        frmAxiomMain.PadText.BackColor = CLng(vArray(LBound(vArray) + 2))
    On Error GoTo 0

    'FONT:
    On Error Resume Next
        sTemp = TextFont
        vArray = ""
        vArray = Split(sTemp, ",")
        LockWindowUpdate frmAxiomMain.hWnd
            frmAxiomMain.MainText.Font.Name = CStr(vArray(LBound(vArray)))
            frmAxiomMain.MainText.Font.Size = CCur(vArray(LBound(vArray) + 1))
            frmAxiomMain.MainText.Font.Bold = CBool(vArray(LBound(vArray) + 2))
            
            frmAxiomMain.OutputText.Font.Name = CStr(vArray(LBound(vArray)))
            frmAxiomMain.OutputText.Font.Size = CCur(vArray(LBound(vArray) + 1))
            frmAxiomMain.OutputText.Font.Bold = CBool(vArray(LBound(vArray) + 2))
            
            frmAxiomMain.PadText.Font.Name = CStr(vArray(LBound(vArray)))
            frmAxiomMain.PadText.Font.Size = CCur(vArray(LBound(vArray)) + 1)
            frmAxiomMain.PadText.Font.Bold = CBool(vArray(LBound(vArray) + 2))
        LockWindowUpdate 0
    On Error GoTo 0

    'MRU List
    If go_MRU.Count > 0 Then
            frmAxiomMain.Hiphen.Visible = True
            For idx = 0 To go_MRU.Count - 1
                frmAxiomMain.mnuMRU(idx).Caption = "&" & Trim$(CStr(idx + 1)) _
                & " " & go_MRU.Item(idx)
                frmAxiomMain.mnuMRU(idx).Visible = True
            Next idx
    End If

End With

End Sub

Public Property Let WordWrap(ByVal bNewValue As Boolean)

    mb_WordWrap = bNewValue
    
End Property


Public Property Get WordWrap() As Boolean

    WordWrap = mb_WordWrap

End Property


Public Property Let CharsToKeep(ByVal eNewValue As CharRangeConstants)
        
        me_CharsToKeep = eNewValue
        
End Property

Public Property Let TrimTabsAlso(ByVal bNewValue As Boolean)

    mb_TrimTabsAlso = bNewValue
    
End Property


Public Property Get TrimTabsAlso() As Boolean

     TrimTabsAlso = mb_TrimTabsAlso

End Property


Public Property Let SpacesPerTab(ByVal iNewValue As Integer)

    mi_SpacesPerTab = iNewValue
    
End Property


Public Property Get SpacesPerTab() As Integer

    SpacesPerTab = mi_SpacesPerTab

End Property


Public Property Get CharsToKeep() As CharRangeConstants
        
        CharsToKeep = me_CharsToKeep
        
End Property


Public Sub UpdateSettings()

End Sub


Public Property Get MultipleTextBoxes() As Boolean
       MultipleTextBoxes = mb_MultipleTextBoxes
End Property

Public Property Let MultipleTextBoxes(ByVal bNewValue As Boolean)
       mb_MultipleTextBoxes = bNewValue
End Property

Public Sub LoadSettings()
Dim idx As Long
Dim AxiomIni As New CINIFileAccess

AxiomIni.FileName = RemoveSlash(App.Path) & "\Axiom.ini"
AxiomIni.Section = "Options"
AxiomIni.Key = "MultipleTextBoxes"
AxiomIni.DefaultValue = "1"
MultipleTextBoxes = CBool(AxiomIni.Value)

AxiomIni.Key = "TrimTabsAlso"
AxiomIni.DefaultValue = "1"
TrimTabsAlso = CBool(AxiomIni.Value)

AxiomIni.Key = "SpacesPerTab"
AxiomIni.DefaultValue = "8"
SpacesPerTab = CInt(AxiomIni.Value)

AxiomIni.Key = "WordWrap"
AxiomIni.DefaultValue = "1"
WordWrap = CBool(AxiomIni.Value)

AxiomIni.Key = "CharsToKeep"
AxiomIni.DefaultValue = CStr([All Text Chars])
CharsToKeep = CLng(AxiomIni.Value)

AxiomIni.Key = "Colors"
AxiomIni.DefaultValue = "15921906,16776178,16056319"
AxiomSettings.Colors = AxiomIni.Value

AxiomIni.Key = "Font"
AxiomIni.DefaultValue = "MS Sans Serif,8,False"
AxiomSettings.TextFont = AxiomIni.Value

go_MRU.MaxCount = 5
AxiomIni.Section = "MRU"
For idx = 5 To 1 Step -1  ' MRU adds last in first pos
    AxiomIni.Key = "File" & Trim$(CStr(idx))
    AxiomIni.DefaultValue = ""
    go_MRU.Add AxiomIni.Value
Next idx

Set AxiomIni = Nothing

End Sub

Public Sub SaveSettings()
Dim idx As Long
Dim AxiomIni As New CINIFileAccess

AxiomIni.FileName = RemoveSlash(App.Path) & "\Axiom.ini"

AxiomIni.Section = "Options"
AxiomIni.Key = "MultipleTextBoxes"
AxiomIni.Value = CStr(MultipleTextBoxes)

AxiomIni.Key = "TrimTabsAlso"
AxiomIni.Value = CStr(TrimTabsAlso)

AxiomIni.Key = "SpacesPerTab"
AxiomIni.Value = CStr(SpacesPerTab)

AxiomIni.Key = "WordWrap"
AxiomIni.Value = CStr(WordWrap)

AxiomIni.Key = "CharsToKeep"
AxiomIni.Value = CStr(CharsToKeep)

AxiomIni.Key = "Colors"
AxiomIni.Value = AxiomSettings.Colors

AxiomIni.Key = "Font"
AxiomIni.Value = AxiomSettings.TextFont

AxiomIni.Section = "MRU"
For idx = 0 To 4
    AxiomIni.Key = "File" & Trim$(CStr(idx + 1))
    AxiomIni.Value = go_MRU.Item(idx)
Next idx

Set AxiomIni = Nothing
    
End Sub

