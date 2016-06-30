Attribute VB_Name = "mousewheel"
Option Explicit
Option Private Module

Private Type POINTAPI
     X As Long
     Y As Long
End Type

#If Win64 Then
    Private Type MOUSEHOOKSTRUCT
         pt As POINTAPI
         Hwnd As LongPtr
         wHitTestCode As Long
         dwExtraInfo As LongPtr
    End Type
#Else
    Private Type MOUSEHOOKSTRUCT
         pt As POINTAPI
         Hwnd As Long
         wHitTestCode As Long
         dwExtraInfo As Long
    End Type
#End If

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" _
                         Alias "FindWindowA" ( _
                                 ByVal lpClassName As String, _
                                 ByVal lpWindowName As String) As Long
    
    Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" _
                         Alias "GetWindowLongA" ( _
                                 ByVal Hwnd As Long, _
                                 ByVal nIndex As Long) As Long
    
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" _
                         Alias "SetWindowsHookExA" ( _
                                 ByVal idHook As Long, _
                                 ByVal lpfn As LongPtr, _
                                 ByVal hmod As Long, _
                                 ByVal dwThreadId As Long) As Long
    
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" ( _
                                 ByVal hHook As Long, _
                                 ByVal nCode As Long, _
                                 ByVal wParam As Long, _
                                 lParam As Any) As Long
    
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" ( _
                                 ByVal hHook As Long) As Long
    
    Private Declare PtrSafe Function PostMessage Lib "user32.dll" _
                         Alias "PostMessageA" ( _
                                 ByVal Hwnd As Long, _
                                 ByVal wMsg As Long, _
                                 ByVal wParam As Long, _
                                 ByVal lParam As Long) As Long
    
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" ( _
                                 ByVal xPoint As LongPtr, _
                                 ByVal yPoint As LongPtr) As Long
    
    Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" ( _
                                 ByRef lpPoint As POINTAPI) As Long
#Else
    Private Declare Function FindWindow Lib "user32" _
                         Alias "FindWindowA" ( _
                                 ByVal lpClassName As String, _
                                 ByVal lpWindowName As String) As Long
    
    Private Declare Function GetWindowLong Lib "user32.dll" _
                         Alias "GetWindowLongA" ( _
                                 ByVal Hwnd As Long, _
                                 ByVal nIndex As Long) As Long
    
    Private Declare Function SetWindowsHookEx Lib "user32" _
                         Alias "SetWindowsHookExA" ( _
                                 ByVal idHook As Long, _
                                 ByVal lpfn As Long, _
                                 ByVal hmod As Long, _
                                 ByVal dwThreadId As Long) As Long
    
    Private Declare Function CallNextHookEx Lib "user32" ( _
                                 ByVal hHook As Long, _
                                 ByVal nCode As Long, _
                                 ByVal wParam As Long, _
                                 lParam As Any) As Long
    
    Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
                                 ByVal hHook As Long) As Long
    
    Private Declare Function PostMessage Lib "user32.dll" _
                         Alias "PostMessageA" ( _
                                 ByVal Hwnd As Long, _
                                 ByVal wMsg As Long, _
                                 ByVal wParam As Long, _
                                 ByVal lParam As Long) As Long
    
    Private Declare Function WindowFromPoint Lib "user32" ( _
                                 ByVal xPoint As Long, _
                                 ByVal yPoint As Long) As Long
    
    Private Declare Function GetCursorPos Lib "user32.dll" ( _
                                 ByRef lpPoint As POINTAPI) As Long
#End If

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const HC_ACTION As Long = 0
Private Const GWL_HINSTANCE As Long = (-6)

'Private Const WM_KEYDOWN As Long = &H100
'Private Const WM_KEYUP As Long = &H101
'Private Const VK_UP As Long = &H26
'Private Const VK_DOWN As Long = &H28
'Private Const WM_LBUTTONDOWN As Long = &H201

Private mLngMouseHook As Long
Private mListHwnd As Long
Private mbHook As Boolean
Private mCtl As MSForms.control
Dim n As Long

Private Const Verbose As Boolean = False
Private Sub output(txt As String)
'will output to debug window
If Verbose Then
    Debug.Print txt
End If
End Sub
Sub HookListScroll(frm As Object, ctl As MSForms.control)
Dim lngAppInst As Long
Dim hwndUnderCursor As Long
Dim tPT As POINTAPI

output "Trying to hook control..."

'get cursorposition
    GetCursorPos tPT
    output "cursor position: " & CStr(tPT.X) & " x " & CStr(tPT.Y)
'get window handler of window under cursor
    hwndUnderCursor = WindowFromPoint(tPT.X, tPT.Y)
    output "Window handle under cursor: " & CStr(hwndUnderCursor)
'set focus
    If Not frm.ActiveControl Is ctl Then
        ctl.SetFocus
    End If
'check if window under cursor is stored window
    If mListHwnd <> hwndUnderCursor Then
        output "Window under cursor is not stored window."
        'unhook to create new hook
            UnhookListScroll
        'store control
            Set mCtl = ctl
            output "storing control: " & ctl.Name
        'store window handler
            mListHwnd = hwndUnderCursor
            output "storing window handler..."
        'setup hook
            output "getting window long..."
            lngAppInst = GetWindowLong(mListHwnd, GWL_HINSTANCE)
            output CStr(lngAppInst)
            If Not mbHook Then
                output "Hooking..."
                mLngMouseHook = SetWindowsHookEx( _
                    WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
                output "MouseHook Long value: " & mLngMouseHook
                mbHook = mLngMouseHook <> 0
                output "Hooking succes: " & mbHook
            End If
    End If
    output vbNullString
End Sub

Sub UnhookListScroll()
'check hook
    If mbHook Then
        'null stored control
            Set mCtl = Nothing
        'unhook
            UnhookWindowsHookEx mLngMouseHook
        'reset variables
            mLngMouseHook = 0
            mListHwnd = 0
            mbHook = False
    End If
End Sub

#If Win64 Then
Private Function MouseProc( _
             ByVal nCode As Long, ByVal wParam As Long, _
             ByRef lParam As MOUSEHOOKSTRUCT) As LongPtr
#Else
Private Function MouseProc( _
             ByVal nCode As Long, ByVal wParam As Long, _
             ByRef lParam As MOUSEHOOKSTRUCT) As Long
#End If
Dim idx As Long

On Error GoTo errH
If (nCode = HC_ACTION) Then
    If WindowFromPoint(lParam.pt.X, lParam.pt.Y) = mListHwnd Then
        If wParam = WM_MOUSEWHEEL Then
            output "Mousewheel motion in stored control detected..."
            output "control name: " & mCtl.Name
            output "control listindex: " & mCtl.ListIndex
            MouseProc = True
            output "lparam hwnd: " & lParam.Hwnd
            'value for lParam.Hwnd found by trial and error
            If lParam.Hwnd > 7864320 Then idx = -1 Else idx = 1
            output "Mouse wheel motion direction: " & idx
            idx = idx + mCtl.ListIndex
            If idx >= 0 And idx <= mCtl.ListCount - 1 Then
                output "Listindex change request from " & mCtl.ListIndex & " to " & idx
                mCtl.ListIndex = idx
            Else
                output "No change request (not possible in this direction)"
            End If
            output vbNullString
            Exit Function
        End If
    Else
        UnhookListScroll
    End If
End If
MouseProc = CallNextHookEx( _
                        mLngMouseHook, nCode, wParam, ByVal lParam)
Exit Function
errH:
     UnhookListScroll
End Function

'************
'IN USERFORM:
'************
'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'     UnhookListScroll
'End Sub
'Private Sub CONTROL_NAME_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
''start the hook
'    HookListScroll Me, Me.CONTROL_NAME
'End Sub
