Attribute VB_Name = "WindowHook"
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - November 2002
'If you got it elsewhere - they stole it from PSC.

'Please visit our website at www.psst.com.au

Option Explicit
'Used to catch the "WM_SETFOCUS" message to remove
'the Focus rectangle from Checkboxes, Optionbuttons and Commandbuttons
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const GWL_WNDPROC = (-4)
Public lpPrevWndProc As Long
Private Const WM_SETFOCUS = &H7

Public Sub Hook(mHwnd As Long)
    lpPrevWndProc = SetWindowLong(mHwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook(mHwnd As Long)
    SetWindowLong mHwnd, GWL_WNDPROC, lpPrevWndProc
End Sub

Function WindowProc(ByVal mHwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Select Case uMsg
        Case WM_SETFOCUS
            Exit Function
    End Select
    WindowProc = CallWindowProc(lpPrevWndProc, mHwnd, uMsg, wParam, lParam)

End Function

