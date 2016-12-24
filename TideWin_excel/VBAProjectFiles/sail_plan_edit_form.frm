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
Public vessel_underway As Boolean

Dim caller_ctr As MSForms.control
Private WithEvents cal As cCalendar
Attribute cal.VB_VarHelpID = -1

Private Sub cal_Click()
caller_ctr.Text = cal.Value

End Sub

Private Sub cal_DblClick()
    caller_ctr.Text = cal.Value
    Call destroy_datepicker
    Call select_next_ctr(caller_ctr.parent, caller_ctr.TabIndex)
End Sub
Private Sub select_next_ctr(parent_ctr As MSForms.control, tab_index As Long)
'will select the next control by tabindex
Dim ctr As MSForms.control
Dim dif As Long
Dim c_name As String

'large number
    dif = 100000

'loop controls
    For Each ctr In parent_ctr.Controls
        'check tabindex
            If ctr.TabIndex > tab_index And ctr.TabIndex - tab_index < dif Then
                'store dif and name
                    dif = ctr.TabIndex - tab_index
                    c_name = ctr.Name
            End If
        If dif = 1 Then Exit For
    Next ctr

'if a control is found:
    If dif < 100000 Then
        'carefull, some controls cannot be focussed
            On Error Resume Next
                parent_ctr.Controls(c_name).SetFocus
            On Error GoTo 0
    End If
    
End Sub

Private Sub cal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal shift As Integer)
If KeyCode = vbKeyEscape Then
    Call destroy_datepicker
End If

End Sub

Private Sub destroy_datepicker()
Dim ctr As MSForms.control

Me.datepicker_frame.Visible = False

Set cal = Nothing
Do While Me.datepicker_frame.Controls.Count > 0
    Me.datepicker_frame.Controls.Remove (0)
Loop

'restore backcolor
caller_ctr.BackColor = -2147483643

End Sub
Private Sub create_datepicker()
Set cal = New cCalendar
Dim t As Double
Dim L As Double
Dim ctr As MSForms.control
Set ctr = caller_ctr

'calculate global position of the datepicker
'(flush left and below the control)
    t = ctr.Height
    On Error Resume Next
        Do While True
            t = t + ctr.Top
            L = L + ctr.Left
            If ctr.parent.Name = Me.Name Then Exit Do
            Set ctr = ctr.parent
        Loop
    On Error GoTo 0

'position the container frame
    With Me.datepicker_frame
        .Visible = True
        .Top = t
        .Left = L
        .ZOrder (0)
    End With

'set red color to control
    caller_ctr.BackColor = vbRed

'build datepicker
    cal.Add_Calendar_into_Frame Me.datepicker_frame

End Sub


Private Sub CommandButton1_Click()
Call proj.sail_plan_form_ok_click
End Sub




Private Sub CommandButton2_Click()
cancelflag = True
Me.Hide
End Sub

Private Sub dr_dbl_ob_Click()
'double draught
Dim t As Long
'frame visible
t = 66
Me.draught_single_frame.Visible = False

With Me.draught_double_frame
    .Visible = True
    .Top = t
    .Left = 12
    t = t + .Height + 3
End With
Me.Label3.Top = t + 6
Me.TextBox3.Top = t

t = t + 18
Me.Label4.Top = t + 6
Me.TextBox4.Top = t

t = t + 18
Me.Label5.Top = t + 6
Me.TextBox5.Top = t

t = t + 18
Me.Label6.Top = t + 6
Me.ship_types_cb.Top = t

End Sub

Private Sub dr_single_ob_Click()
'single draught
Dim t As Long
'frame visible
t = 66
Me.draught_double_frame.Visible = False

With Me.draught_single_frame
    .Visible = True
    .Top = t
    .Left = 12
    t = t + .Height + 3
End With
Me.Label3.Top = t
Me.TextBox3.Top = t

Me.Label3.Top = t + 6
Me.TextBox3.Top = t

t = t + 18
Me.Label4.Top = t + 6
Me.TextBox4.Top = t

t = t + 18
Me.Label5.Top = t + 6
Me.TextBox5.Top = t

t = t + 18
Me.Label6.Top = t + 6
Me.ship_types_cb.Top = t
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

Private Sub eta_time_tb_Enter()
eta_time_tb.SelStart = 0
eta_time_tb.SelLength = Len(eta_time_tb.Text)
End Sub

Private Sub route_lb_MouseUp(ByVal Button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single)
Call proj.sail_plan_form_route_lb_click

End Sub

Private Sub routes_cb_Change()
Call proj.sail_plan_form_route_cb_exit
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

Private Sub eta_ob_Click()
Me.rta_frame.Visible = False
Me.current_frame.Visible = False
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

Private Sub ship_types_cb_Change()
Call proj.sail_plan_form_ship_type_cb_change
End Sub

Private Sub ships_cb_Change()
Me.ships_cb.Value = UCase(Me.ships_cb.Value)

Call proj.sail_plan_form_ship_cb_change

End Sub

Private Sub ships_cb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal shift As Integer)
If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    KeyCode = 0
End If
End Sub


Private Sub TextBox2_Change()
TextBox2.Text = UCase(TextBox2.Text)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    UnhookListScroll
End Sub

Private Sub window_after_edit_tb_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If Not input_mask_time(Me.window_after_edit_tb) Then
    Cancel = True
End If
Call check_route_list_tidal_windows
End Sub

Private Sub window_after_edit_tb_Change()
Call proj.sail_plan_form_window_edit_after_change
End Sub


Private Sub window_pre_edit_tb_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If Not input_mask_time(Me.window_pre_edit_tb) Then
    Cancel = True
End If
Call check_route_list_tidal_windows
End Sub
Private Sub window_pre_edit_tb_Change()
Call proj.sail_plan_form_window_edit_pre_change
End Sub

Public Sub check_route_list_tidal_windows()
'will loop the tidal windows in the route list and check them against
'the set tidal window in the voyage data. If one (or more) is smaller, display warning.
Dim w0_min As Date
Dim w0_max As Date
Dim w1_min As Date
Dim w1_max As Date

Dim i As Long

'set very high
    w1_min = 10000
    w0_min = 10000

'find min and max values
    For i = 0 To Me.route_lb.ListCount - 1 Step 2
        If Me.route_lb.List(i, 4) > w1_max Then w1_max = Me.route_lb.List(i, 4)
        If Me.route_lb.List(i, 5) > w0_max Then w0_max = Me.route_lb.List(i, 5)
        
        If Me.route_lb.List(i, 4) < w1_min Then w1_min = Me.route_lb.List(i, 4)
        If Me.route_lb.List(i, 5) < w0_min Then w0_min = Me.route_lb.List(i, 5)
    Next i

'set minimum value in tbs
    Me.window_pre_tb.Text = Format(w0_min, "hh:nn")
    Me.window_after_tb.Text = Format(w1_min, "hh:nn")

'set or reset warning labels
    If w0_min <> w0_max Then
        'show warning labels
            Me.warning_label.Visible = True
            Me.warning_color_label.Visible = True
        'set text color to red
            Me.window_pre_tb.ForeColor = vbRed
    End If
    If w1_min <> w1_max Then
        'show warning labels
            Me.warning_label.Visible = True
            Me.warning_color_label.Visible = True
        'set text color to red
            Me.window_after_tb.ForeColor = vbRed
    End If
    If w0_min = w0_max And w1_min = w1_max Then
        'hide warning labels
            Me.warning_label.Visible = False
            Me.warning_color_label.Visible = False
        'reset text color to black
            Me.window_after_tb.ForeColor = vbBlack
            Me.window_pre_tb.ForeColor = vbBlack
    End If

End Sub

'routines for 'global' window tbs (pre and after)
Private Sub window_pre_tb_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not input_mask_time(Me.window_pre_tb) Then
        Cancel = True
    End If
    Call insert_windows_into_route_list
End Sub
Private Sub window_after_tb_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not input_mask_time(Me.window_after_tb) Then
        Cancel = True
    End If
    Call insert_windows_into_route_list
End Sub

Private Sub eta_time_tb_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If Not input_mask_time(Me.eta_time_tb) Then
    Cancel = True
End If
End Sub
Private Sub insert_windows_into_route_list()
'will insert the tidal windows from the voyage tab into the route list
Dim w0 As Date
Dim w1 As Date
Dim i As Long

'tidal windows
    w0 = CDate(Me.window_pre_tb)
    w1 = CDate(Me.window_after_tb)

'insert these windows into route listbox
    For i = 0 To Me.route_lb.ListCount - 1 Step 2
        Me.route_lb.List(i, 4) = Format(w1, "hh:nn")
        Me.route_lb.List(i, 5) = Format(w0, "hh:nn")
    Next i

'set text color to black
    Me.window_after_tb.ForeColor = vbBlack
    Me.window_pre_tb.ForeColor = vbBlack

'hide warnings
    Me.warning_label.Visible = False
    Me.warning_color_label.Visible = False

'unset edit mode in route listbox
    Me.route_lb.ListIndex = -1
    Call proj.sail_plan_form_unset_sail_plan_edit_mode

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
Me.warning_label.Visible = False
Me.warning_color_label.Visible = False

Me.ship_types_cb.ColumnCount = 3
Me.ship_types_cb.ColumnWidths = ";0;0"

Me.routes_cb.ColumnCount = 2
Me.routes_cb.ColumnWidths = ";0"

Me.rta_tresholds_cb.ColumnCount = 2
Me.rta_tresholds_cb.ColumnWidths = ";0"

Me.current_tresholds_cb.ColumnCount = 2
Me.current_tresholds_cb.ColumnWidths = ";0"

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

Me.dr_single_ob.Value = True

Me.MultiPage1.Value = 0

End Sub
Private Function input_mask_time(tb As MSForms.TextBox) As Boolean
Dim ss() As String

input_mask_time = True

If tb.Text Like "####" Then
    'could be a time without seperator
    ReDim ss(0 To 1) As String
    ss(0) = Left(tb.Text, 2)
    ss(1) = Right(tb.Text, 2)
ElseIf tb.Text Like "##.##" Or tb.Text Like "#.##" Then
    ss = Split(tb.Text, ".")
ElseIf tb.Text Like "##;##" Or tb.Text Like "#;##" Then
    ss = Split(tb.Text, ";")
ElseIf tb.Text Like "##:##" Or tb.Text Like "#:##" Then
    ss = Split(tb.Text, ":")
Else
    GoTo NotValid
End If

If val(ss(0)) < 0 Or val(ss(0)) > 23 Then GoTo NotValid
If val(ss(1)) < 0 Or val(ss(1)) > 59 Then GoTo NotValid

tb.Text = Format(ss(0), "00") & ":" & Format(ss(1), "00")

Exit Function

NotValid:
MsgBox "De ingevoerde tijd wordt niet herkend als valide tijdswaarde"
input_mask_time = False
tb.SetFocus
tb.SelStart = 0
tb.SelLength = Len(tb.Text)
End Function
