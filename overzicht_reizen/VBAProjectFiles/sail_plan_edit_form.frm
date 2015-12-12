VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sail_plan_edit_form 
   Caption         =   "Nieuw vaarplan"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7770
   OleObjectBlob   =   "sail_plan_edit_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sail_plan_edit_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cancelflag As Boolean

Dim caller_ctr As MSForms.Control
Private WithEvents cal As cCalendar
Attribute cal.VB_VarHelpID = -1

Private Sub cal_Click()
caller_ctr.text = cal.Value

End Sub

Private Sub cal_DblClick()
caller_ctr.text = cal.Value
Call destroy_datepicker

End Sub


Private Sub cal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal shift As Integer)
If KeyCode = vbKeyEscape Then
    Call destroy_datepicker
End If

End Sub

Private Sub destroy_datepicker()
Me.datepicker_frame.Visible = False
Set cal = Nothing
'restore backcolor
caller_ctr.BackColor = -2147483643

End Sub
Private Sub create_datepicker()
'in this case, all controls that create the datepicker are on the multipage.
'multipage behaves strangely if global position is calculated. Probably the
'tabstrip that is not properly calculated.
Set cal = New cCalendar
Dim t As Double
Dim L As Double
Dim ctr As MSForms.Control
Set ctr = caller_ctr
On Error Resume Next
    Do Until ctr.Parent.Name = Me.Name
        t = t + ctr.Top
        L = L + ctr.Left
        Set ctr = ctr.Parent
    Loop
On Error GoTo 0
'add 15 to compensate for the tabstrip (see above)
t = t + caller_ctr.Height + 15

With Me.datepicker_frame
    .Visible = True
    .Top = t
    .Left = L
    .ZOrder (0)
End With
'set red color to control
caller_ctr.BackColor = vbRed

cal.Today
cal.Add_Calendar_into_Frame Me.datepicker_frame
End Sub


Private Sub CommandButton1_Click()
Call proj.sail_plan_form_ok_click
End Sub




Private Sub CommandButton2_Click()
cancelflag = True
Me.Hide
End Sub

Private Sub eta_date_tb_Enter()
Set caller_ctr = Me.eta_date_tb
Call create_datepicker

End Sub

Private Sub eta_date_tb_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Not cal Is Nothing Then Call destroy_datepicker

End Sub

Private Sub eta_date_tb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal shift As Integer)
If KeyCode = vbKeyEscape Then
    If Not cal Is Nothing Then Call destroy_datepicker
End If

End Sub


Private Sub MultiPage1_Change()

End Sub

Private Sub route_lb_MouseUp(ByVal Button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal y As Single)
Call proj.sail_plan_form_route_lb_click

End Sub

Private Sub routes_cb_Change()
Call proj.sail_plan_form_route_cb_exit
End Sub

Private Sub routes_cb_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub rta_date_tb_Enter()
Set caller_ctr = Me.rta_date_tb
Call create_datepicker

End Sub

Private Sub rta_date_tb_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Not cal Is Nothing Then Call destroy_datepicker

End Sub

Private Sub rta_date_tb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal shift As Integer)
If KeyCode = vbKeyEscape Then
    If Not cal Is Nothing Then Call destroy_datepicker
End If

End Sub
'input masks
Private Sub current_after_tb_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If Not input_mask_time(Me.current_after_tb) Then
    Cancel = True
End If
End Sub
Private Sub current_before_tb_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If Not input_mask_time(Me.current_before_tb) Then
    Cancel = True
End If
End Sub

Private Sub rta_ob_Click()
Me.rta_frame.Visible = True
Me.current_frame.Visible = False
End Sub
Private Sub current_ob_Click()
Me.current_frame.Visible = True
Me.rta_frame.Visible = False
End Sub

Private Sub rta_time_tb_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If Not input_mask_time(Me.rta_time_tb) Then
    Cancel = True
End If
End Sub
Private Sub window_after_edit_tb_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If Not input_mask_time(Me.window_after_edit_tb) Then
    Cancel = True
End If
End Sub

Private Sub window_after_edit_tb_Change()
Call proj.sail_plan_form_window_edit_after_change
End Sub

Private Sub window_after_tb_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If Not input_mask_time(Me.window_after_tb) Then
    Cancel = True
End If
End Sub

Private Sub window_pre_edit_tb_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If Not input_mask_time(Me.window_pre_edit_tb) Then
    Cancel = True
End If
End Sub

Private Sub window_pre_edit_tb_Change()
Call proj.sail_plan_form_window_edit_pre_change
End Sub

Private Sub window_pre_tb_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If Not input_mask_time(Me.window_pre_tb) Then
    Cancel = True
End If
End Sub
Private Sub eta_time_tb_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If Not input_mask_time(Me.eta_time_tb) Then
    Cancel = True
End If
End Sub

Private Sub ships_cb_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Call proj.sail_plan_form_ship_cb_exit
End Sub

Private Sub speed_cmb_Change()
Call proj.sail_plan_form_speed_change
End Sub

Private Sub UKC_unit_cb_Change()
Call proj.sail_plan_form_ukc_change
End Sub

Private Sub UKC_val_tb_Change()
Call proj.sail_plan_form_ukc_change
End Sub

Private Sub UserForm_Initialize()
Me.datepicker_frame.Visible = False

Me.ship_types_cb.ColumnCount = 2
Me.ship_types_cb.ColumnWidths = ";0"

Me.routes_cb.ColumnCount = 2
Me.routes_cb.ColumnWidths = ";0"

Me.rta_tresholds_cb.ColumnCount = 2
Me.rta_tresholds_cb.ColumnWidths = ";0"

Me.ships_cb.ColumnCount = 8
Me.ships_cb.ColumnWidths = ";0;0;0;0;0;0;0"

Me.route_lb.ColumnCount = 6
Me.route_lb.ColumnWidths = "75;31;30;30;30;30"

Me.speed_cmb.ColumnCount = 2
Me.speed_cmb.ColumnWidths = ";0"

Me.UKC_unit_cb.AddItem "%"
Me.UKC_unit_cb.AddItem "m"
Me.UKC_unit_cb.Value = "%"

Me.speed_edit_frame.Visible = False
Me.speed_edit_frame.Top = 6
Me.UKC_edit_frame.Visible = False
Me.UKC_edit_frame.Top = 6
Me.window_edit_frame.Visible = False
Me.window_edit_frame.Top = 45

Me.rta_frame.Visible = False
Me.current_frame.Visible = False
Me.rta_frame.Left = 6
Me.current_frame.Left = 6

Me.eta_ob = True

Me.current_before_cb.AddItem "voor"
Me.current_before_cb.AddItem "na"
Me.current_before_cb.Value = "voor"

Me.current_after_cb.AddItem "voor"
Me.current_after_cb.AddItem "na"
Me.current_after_cb.Value = "na"

End Sub
Private Function input_mask_time(tb As MSForms.TextBox) As Boolean
Dim ss() As String

input_mask_time = True

If tb.text Like "####" Then
    'could be a time without seperator
    ReDim ss(0 To 1) As String
    ss(0) = Left(tb.text, 2)
    ss(1) = Right(tb.text, 2)
ElseIf tb.text Like "##.##" Or tb.text Like "#.##" Then
    ss = Split(tb.text, ".")
ElseIf tb.text Like "##;##" Or tb.text Like "#;##" Then
    ss = Split(tb.text, ";")
ElseIf tb.text Like "##:##" Or tb.text Like "#:##" Then
    ss = Split(tb.text, ":")
Else
    GoTo NotValid
End If

If val(ss(0)) < 0 Or val(ss(0)) > 23 Then GoTo NotValid
If val(ss(1)) < 0 Or val(ss(1)) > 59 Then GoTo NotValid

tb.text = Format(ss(0), "00") & ":" & Format(ss(1), "00")

Exit Function

NotValid:
MsgBox "De ingevoerde tijd wordt niet herkend als valide tijdswaarde"
input_mask_time = False
tb.SetFocus
tb.SelStart = 0
tb.SelLength = Len(tb.text)
End Function
