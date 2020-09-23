Attribute VB_Name = "API_Code"
Option Explicit

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hsubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)
Public Declare Function GetTickCount Lib "Kernel32" () As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWndCaller As Long, ByVal pszFileName As String, ByVal uCommand As Long, ByVal dwData As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Const EM_CANUNDO = &HC6
Private Const EM_UNDO = &HC7
Private Const EM_LINEINDEX = &HBB
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const WM_USER = &H400
Private Const TB_SETSTYLE = WM_USER + 56
Private Const TB_GETSTYLE = WM_USER + 57
Private Const TBSTYLE_FLAT = &H800
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Const EM_SETTARGETDEVICE = (WM_USER + 72)
Private Const ES_LOWERCASE = &H10&
Private Const ES_UPPERCASE = &H8&
Private Const ES_NUMBER = &H2000&
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_CLOSE_ALL = &H12
Private Const MF_BITMAP = &H4&
Private Const MIIM_ID = &H2
Private Const MIIM_TYPE = &H10
Private Const MFT_STRING = &H0&

Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "Kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Const STILL_ACTIVE = &H103
Const PROCESS_QUERY_INFORMATION = &H400

Sub ExecWait(ByVal JobToDo As String)

         Dim hProcess As Long
         Dim RetVal As Long
         'The next line launches JobToDo as icon,

         'captures process ID
         hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(JobToDo, 1))

         Do

             'Get the status of the process
             GetExitCodeProcess hProcess, RetVal

             'Sleep command recommended as well as DoEvents
             DoEvents: Sleep 100

         'Loop while the process is active
         Loop While RetVal = STILL_ACTIVE


End Sub

Public Sub SetMenuIcon(hWnd As Long, MenuIndex As Long, SubIndex As Long, pic As Picture)
Dim hMenu As Long, hsubMenu As Long, hID As Long

'Get the menuhandle of the form
hMenu = GetMenu(hWnd)

'Get the handle of the first submenu
hsubMenu = GetSubMenu(hMenu, MenuIndex)

'Get the menuId of the first entry
hID = GetMenuItemID(hsubMenu, SubIndex)

'Add the bitmap
SetMenuItemBitmaps hMenu, hID, MF_BITMAP, pic, pic

End Sub
Public Function CanUndo(TextCtl As Control) As Boolean
    
CanUndo = CBool(SendMessage(TextCtl.hWnd, EM_CANUNDO, ByVal CLng(0), ByVal CLng(0)))

End Function
Public Function PerformUndo(TextCtl As Control) As Long

PerformUndo = SendMessage(TextCtl.hWnd, EM_UNDO, ByVal CLng(0), ByVal CLng(0))

End Function

Public Sub AddBorderToAllTextBoxes(frmX As Form)

Dim x As Control

On Error Resume Next
For Each x In frmX.Controls
        If TypeOf x Is TextBox Then
                AddOfficeBorder x
        End If
Next

End Sub


Public Sub AddOfficeBorder(ctlX As Control)
    
    Dim lngRetVal As Long
    
    'Retrieve the current border style
    lngRetVal = GetWindowLong(ctlX.hWnd, GWL_EXSTYLE)
    
    'Calculate border style to use
    lngRetVal = lngRetVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    
    'Apply the changes
    SetWindowLong ctlX.hWnd, GWL_EXSTYLE, lngRetVal
    SetWindowPos ctlX.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
End Sub

Public Sub HHelp_Show(ByVal ChmFileName As String, HtmFileName As String)

Call HtmlHelp(0, ChmFileName, HH_DISPLAY_TOPIC, HtmFileName)

End Sub


Public Sub HHelp_Close()

Call HtmlHelp(0, "", HH_CLOSE_ALL, "")

End Sub



Public Sub CButtons(frmX As Form, Optional Identifier As String)
' Button.Style must be GRAPHICAL

Dim Ctl As Control

For Each Ctl In frmX      'loop trough all the controls on the form
    
    '3 Methods of doing it
    'If LCase(Left(Control.Name, Len(Identifier))) = LCase(Identifier) Then
    'If TypeName(Control) = "CommandButton" Then
    If TypeOf Ctl Is CommandButton Then
                SendMessage Ctl.hWnd, &HF4&, &H0&, 0&
    End If

Next Ctl

End Sub

Public Sub SetTopMost(ByVal lHwnd As Long, ByVal bTopMost As Boolean)
'
' Set the hwnd of the window topmost or not topmost
'
    Dim lUseVal  As Long
    Dim lRet As Long
    
    lUseVal = IIf(bTopMost, HWND_TOPMOST, HWND_NOTOPMOST)
    
    lRet = SetWindowPos(lHwnd, lUseVal, 0, 0, 0, 0, flags)
    
    If lRet < 0 Then
'
' Couldn't do operation - handle error here
'
'        DisplayWinAPIError lRet
    End If

End Sub


' Comments  : Allow only numbers in a textbox
' Returns   : The Style of the textbox before the change.
Public Function NumbersOnly(tBox As TextBox)
Dim DefaultStyle As Long
DefaultStyle = GetWindowLong(tBox.hWnd, GWL_STYLE)
NumbersOnly = SetWindowLong(tBox.hWnd, GWL_STYLE, DefaultStyle Or ES_NUMBER)
End Function

Public Function UpperCaseOnly(tBox As TextBox)

Dim DefaultStyle As Long
DefaultStyle = GetWindowLong(tBox.hWnd, GWL_STYLE)
UpperCaseOnly = SetWindowLong(tBox.hWnd, GWL_STYLE, DefaultStyle Or ES_UPPERCASE)

End Function

' Comments  : Allow only lowercase letters in a textbox
' Returns   : The Style of the textbox before the change.
Public Function LowerCaseOnly(tBox As TextBox)
Dim DefaultStyle As Long
DefaultStyle = GetWindowLong(tBox.hWnd, GWL_STYLE)
LowerCaseOnly = SetWindowLong(tBox.hWnd, GWL_STYLE, DefaultStyle Or ES_LOWERCASE)
End Function


' Comments  : Sets the style of a textbox.
' Returns   : The new style.
Public Function SetStyle(tBox As TextBox, NewStyle As Long)
SetStyle = SetWindowLong(tBox.hWnd, GWL_STYLE, NewStyle)
End Function


' Comments  : Gets the current style of a textbox.
' Returns   : The Style of the textbox.
Public Function GetStyle(tBox As TextBox)
GetStyle = GetWindowLong(tBox.hWnd, GWL_STYLE)
End Function

Public Function StyleNumberToText(tBox As TextBox)
Dim StyleNum  As Long
Dim StyleText As String

StyleNum = GetStyle(tBox)

Select Case StyleNum
    Case 1409360064: StyleText = "Number"
    Case 1409351880: StyleText = "Uppercase"
    Case 1409351888: StyleText = "Lowercase"
    Case Else: StyleText = "Other"
End Select

StyleNumberToText = StyleText
End Function

Public Sub SetWordWrap(RichTextBox As RichTextBox, WordWrap As Boolean)

If WordWrap Then
    'Enable word wrap:
    SendMessageLong RichTextBox.hWnd, EM_SETTARGETDEVICE, 0, 0
Else
    'Disable word wrap:
    SendMessageLong RichTextBox.hWnd, EM_SETTARGETDEVICE, 0, 1
End If

End Sub



Public Sub Rtf_SelChange(rtf As RichTextBox, Row As Long, Col As Long)
    Row = rtf.GetLineFromChar(rtf.SelStart) + 1 ' Get the current line
    Col = rtf.SelStart - SendMessage(rtf.hWnd, EM_LINEINDEX, -1, 0&) + 1
End Sub

                

Sub UnFlatAllBtns(frmX As Form)

Dim btnX As Control
For Each btnX In frmX.Controls
    If Left(btnX.Name, 3) = "cmd" Then
            UnbtnFlat btnX
    End If
    
Next btnX

End Sub
Public Function btnFlat(cmdFlat As CommandButton)
    SetWindowLong cmdFlat.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
    cmdFlat.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function

Public Function UnbtnFlat(cmdFlat As CommandButton)
    SetWindowLong cmdFlat.hWnd, GWL_STYLE, WS_CHILD
    cmdFlat.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function

Sub FlatAllBtns(frmX As Form)

Dim btnX As Control
For Each btnX In frmX.Controls
    If TypeOf btnX Is CommandButton Then
            btnFlat btnX
    End If
    
Next btnX

End Sub



Sub ToolFlat(ControlName As Control, flat As Boolean)
    Dim style As Long
    Dim hToolbar As Long
    Dim r As Long
       
'Now Make it Flat
    'First get the hWnd
    hToolbar = FindWindowEx(ControlName.hWnd, 0&, "ToolbarWindow32", vbNullString)
    'get Style
    style = SendMessageLong(hToolbar, TB_GETSTYLE, 0&, 0&)
    'Change style
    If (style And TBSTYLE_FLAT) And Not flat Then
        style = style Xor TBSTYLE_FLAT
    ElseIf flat Then
        style = style Or TBSTYLE_FLAT
    End If
    'Set the Style
    r = SendMessageLong(hToolbar, TB_SETSTYLE, 0, style)
    'Now show what we've done, this isn't neccesary if used in form_load
    ControlName.Refresh
End Sub

