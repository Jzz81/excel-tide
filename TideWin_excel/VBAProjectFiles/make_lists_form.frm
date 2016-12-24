VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} make_lists_form 
   Caption         =   "Lijsten maken"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7605
   OleObjectBlob   =   "make_lists_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "make_lists_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim caller_ctr As MSForms.control
Public sail_plan_id As Long

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
    Me.datepicker_frame.Visible = False
    Set cal = Nothing
    'remove all controls
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

t = ctr.Height

On Error Resume Next
    Do While True
        t = t + ctr.Top
        L = L + ctr.Left
        If ctr.parent.Name = Me.Name Then Exit Do
        Set ctr = ctr.parent
    Loop
On Error GoTo 0

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

Private Sub cancel_btn_Click()
unload Me
End Sub

'**************************
'Datepicker caller routines
'**************************
Private Sub date_0_tb_Enter()
    Set caller_ctr = Me.date_0_tb
    Call create_datepicker
End Sub
Private Sub date_0_tb_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not cal Is Nothing Then Call destroy_datepicker
End Sub
Private Sub date_1_tb_Enter()
    Set caller_ctr = Me.date_1_tb
    Call create_datepicker
End Sub
Private Sub date_1_tb_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not cal Is Nothing Then Call destroy_datepicker
End Sub


Private Sub ok_btn_Click()
Call make_lists.lists_form_ok_btn_click
End Sub

Private Sub rta_ob_Click()

With Me.rta_frame
    .Visible = True
    .Left = 6
End With
Me.tide_frame.Visible = False

End Sub

Private Sub tide_ob_Click()
With Me.tide_frame
    .Visible = True
    .Left = 6
End With
Me.rta_frame.Visible = False

End Sub

Private Sub type_maxT_ob_Click()
Call set_type("maxT")
Me.rta_ob = True
Me.tide_ob.Enabled = False
End Sub

Private Sub type_window_ob_Click()
Call set_type("windows")
Me.tide_ob.Enabled = True
End Sub

Private Sub set_type(tpe As String)
'will adapt the form to the selected type
If tpe = "windows" Then
    Me.Label3.Caption = "De ingevoerde diepgangen moeten hele getallen zijn."
    Me.Label2.Caption = "Bereken van:"
    Me.Label23.Caption = "dm t/m"
ElseIf tpe = "maxT" Then
    Me.Label3.Caption = "Bovenstaande waardes hebben veel invloed op de berekeningssnelheid. Vooral een minimale diepgang die te hoog gekozen wordt geeft een erg lange berekeningstijd. Ook een te grote bandbreedte geeft extra berekeningstijd. Foutieve berekeningen zijn niet mogelijk als gevolg van deze instelling." _
        & vbCr & "Het programma probeert deze instelling tijdens de berekeningen nog bij te sturen."
    Me.Label2.Caption = "Zoek tussen:"
    Me.Label23.Caption = "dm en"
End If

End Sub

Private Sub UserForm_Initialize()
Me.diff_before_after_cbb.AddItem "voor"
Me.diff_before_after_cbb.AddItem "na"
Me.diff_before_after_cbb.Value = "voor"

Me.hw_lw_cbb.AddItem "hoogwater"
Me.hw_lw_cbb.AddItem "laagwater"
Me.hw_lw_cbb.Value = "hoogwater"

Me.datepicker_frame.Visible = False

Me.rta_ob = True
Me.tide_ob.Enabled = False

Me.current_before_cb.AddItem "voor"
Me.current_before_cb.AddItem "na"
Me.current_before_cb.Value = "voor"

Me.current_after_cb.AddItem "voor"
Me.current_after_cb.AddItem "na"
Me.current_after_cb.Value = "na"

End Sub
