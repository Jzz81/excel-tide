VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tresholds_edit_form 
   Caption         =   "Drempels"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7965
   OleObjectBlob   =   "tresholds_edit_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tresholds_edit_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 0
Option Compare Text


Private Sub save_close_btn_Click()
    Call proj.treshold_form_save_click
    Call ws_gui.display_sail_plan
    Unload Me
End Sub
Private Sub SaveBtn_Click()
    Call proj.treshold_form_save_click
End Sub
Private Function validate_keycode(k_code As Long) As Long
'only allow numbers and a point
'numbers
If k_code >= 48 And k_code <= 57 Then
    validate_keycode = k_code
'numpad numbers
ElseIf k_code >= 96 And k_code <= 105 Then
    validate_keycode = k_code
'comma
ElseIf k_code = 188 Then
    validate_keycode = 190
'tab
ElseIf k_code = 9 Then
    validate_keycode = k_code
'backspace
ElseIf k_code = 8 Then
    validate_keycode = k_code
'arrows (left and right)
ElseIf k_code = 37 Or k_code = 39 Then
    validate_keycode = k_code
'point
ElseIf k_code = 190 Then
    validate_keycode = k_code
'numpad point
ElseIf k_code = 110 Then
    validate_keycode = 190
Else
    validate_keycode = 0
End If

End Function
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = validate_keycode(CLng(KeyCode))
End Sub
Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = validate_keycode(CLng(KeyCode))
End Sub
Private Sub TextBox4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = validate_keycode(CLng(KeyCode))
End Sub
Private Sub TextBox5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = validate_keycode(CLng(KeyCode))
End Sub

Private Sub tidal_data_cmb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'start the hook
    UnhookListScroll
    HookListScroll Me, Me.tidal_data_cmb
End Sub

Private Sub tresholds_lb_Click()
    Call proj.treshold_form_listbox_click
End Sub

Private Sub tresholds_lb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'start the hook
    UnhookListScroll
    HookListScroll Me, Me.tresholds_lb
End Sub

Private Sub UserForm_Initialize()
    Me.tresholds_lb.ColumnCount = 2
    Me.tresholds_lb.ColumnWidths = "0;"
    
    Me.UKC_unit_cb.AddItem "m"
    Me.UKC_unit_cb.AddItem "%"
    Me.UKC_unit_cb.Value = "%"
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'unhook
    UnhookListScroll
End Sub

