VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} search_voyage_form 
   Caption         =   "Zoek reizen"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16080
   OleObjectBlob   =   "search_voyage_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "search_voyage_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public extreme_dt_start As Date
Public extreme_dt_end As Date

'***boa***
Private Sub boa_cb_Click()
    Call enable_frame(Me.boa_frame, Me.boa_cb)
    Call proj.search_form_show_results
End Sub
Private Sub boa_end_tb_Change()
    Call proj.search_form_show_results
End Sub
Private Sub boa_start_tb_Change()
    Call proj.search_form_show_results
End Sub
Public Function start_boa() As Double
'will get the start_boa if filled in and valid, else extreme
If Me.boa_start_tb.Text = vbNullString Then
    start_boa = 0
    Exit Function
End If

On Error Resume Next
    start_boa = CDbl(Replace(Me.boa_start_tb.Text, ".", ","))
    If Err.Number <> 0 Then
        start_boa = 0
        Me.boa_start_tb.BackColor = vbRed
    Else
        Me.boa_start_tb.BackColor = vbWhite
    End If
On Error GoTo 0

End Function
Public Function end_boa() As Double
'will get the end_boa if filled in and valid, else extreme
If Me.boa_end_tb.Text = vbNullString Then
    end_boa = 1000
    Exit Function
End If

On Error Resume Next
    end_boa = CDbl(Replace(Me.boa_end_tb.Text, ".", ","))
    If Err.Number <> 0 Then
        end_boa = 0
        Me.boa_end_tb.BackColor = vbRed
    Else
        Me.boa_end_tb.BackColor = vbWhite
    End If
On Error GoTo 0

End Function

'***draught***
Private Sub draught_cb_Click()
    Call enable_frame(Me.draught_frame, Me.draught_cb)
    Call proj.search_form_show_results
End Sub
Private Sub draught_end_tb_Change()
    Call proj.search_form_show_results
End Sub
Private Sub draught_start_tb_Change()
    Call proj.search_form_show_results
End Sub
Public Function start_draught() As Double
'will get the start_draught if filled in and valid, else extreme
If Me.draught_start_tb.Text = vbNullString Then
    start_draught = 0
    Exit Function
End If

On Error Resume Next
    start_draught = CDbl(Replace(Me.draught_start_tb.Text, ".", ","))
    If Err.Number <> 0 Then
        start_draught = 0
        Me.draught_start_tb.BackColor = vbRed
    Else
        Me.draught_start_tb.BackColor = vbWhite
    End If
On Error GoTo 0

End Function
Public Function end_draught() As Double
'will get the end_draught if filled in and valid, else extreme
If Me.draught_end_tb.Text = vbNullString Then
    end_draught = 1000
    Exit Function
End If

On Error Resume Next
    end_draught = CDbl(Replace(Me.draught_end_tb.Text, ".", ","))
    If Err.Number <> 0 Then
        end_draught = 0
        Me.draught_end_tb.BackColor = vbRed
    Else
        Me.draught_end_tb.BackColor = vbWhite
    End If
On Error GoTo 0

End Function

'***loa***
Private Sub loa_cb_Click()
    Call enable_frame(Me.loa_frame, Me.loa_cb)
    Call proj.search_form_show_results
End Sub
Private Sub loa_end_tb_Change()
    Call proj.search_form_show_results
End Sub
Private Sub loa_start_tb_Change()
    Call proj.search_form_show_results
End Sub
Public Function start_loa() As Double
'will get the start_loa if filled in and valid, else extreme
If Me.loa_start_tb.Text = vbNullString Then
    start_loa = 0
    Exit Function
End If

On Error Resume Next
    start_loa = CDbl(Replace(Me.loa_start_tb.Text, ".", ","))
    If Err.Number <> 0 Then
        start_loa = 0
        Me.loa_start_tb.BackColor = vbRed
    Else
        Me.loa_start_tb.BackColor = vbWhite
    End If
On Error GoTo 0

End Function
Public Function end_loa() As Double
'will get the end_loa if filled in and valid, else extreme
If Me.loa_end_tb.Text = vbNullString Then
    end_loa = 1000
    Exit Function
End If

On Error Resume Next
    end_loa = CDbl(Replace(Me.loa_end_tb.Text, ".", ","))
    If Err.Number <> 0 Then
        end_loa = 0
        Me.loa_end_tb.BackColor = vbRed
    Else
        Me.loa_end_tb.BackColor = vbWhite
    End If
On Error GoTo 0

End Function

'***period***
Private Sub period_cb_Click()
    Call enable_frame(Me.period_frame, Me.period_cb)
    Call proj.search_form_show_results
End Sub
Private Sub dt_end_tb_Change()
    Call proj.search_form_show_results
End Sub
Private Sub dt_start_tb_Change()
    Call proj.search_form_show_results
End Sub
Public Function start_dt() As Date
'will get the start_dt if filled in and valid, else extreme
If Me.dt_start_tb.Text = vbNullString Then
    start_dt = extreme_dt_start
    Exit Function
End If

On Error Resume Next
    start_dt = CDate(Me.dt_start_tb.Text)
    If Err.Number <> 0 Then
        start_dt = extreme_dt_start
        Me.dt_start_tb.BackColor = vbRed
    Else
        Me.dt_start_tb.BackColor = vbWhite
    End If
On Error GoTo 0

End Function
Public Function end_dt() As Date
'will get the end_dt if filled in and valid, else extreme
If Me.dt_end_tb.Text = vbNullString Then
    end_dt = extreme_dt_start
    Exit Function
End If

On Error Resume Next
    end_dt = CDate(Me.dt_end_tb.Text)
    If Err.Number <> 0 Then
        end_dt = extreme_dt_end
        Me.dt_end_tb.BackColor = vbRed
    Else
        Me.dt_end_tb.BackColor = vbWhite
    End If
On Error GoTo 0

End Function

'***restrict result***
Private Sub restric_show_count_cb_Click()
    'will force-show the results if more than 250 records
    Call proj.search_form_show_results
End Sub

'***voyage succes***
Private Sub voyage_succes_cb_Click()
    Call enable_frame(Me.voyage_succes_frame, Me.voyage_succes_cb)
    Call proj.search_form_show_results
End Sub
Private Sub voyage_succes_no_ob_Click()
    Call proj.search_form_show_results
End Sub
Private Sub voyage_succes_yes_ob_Click()
    Call proj.search_form_show_results
End Sub
Public Function voyage_succes() As String
If Me.voyage_succes_yes_ob Then
    voyage_succes = "TRUE"
Else
    voyage_succes = "FALSE"
End If
End Function

'***ship type***
Private Sub ship_type_cb_Click()
    Call enable_frame(Me.ship_type_frame, Me.ship_type_cb)
    Call proj.search_form_show_results
End Sub
Private Sub ship_type_cbb_Change()
    Call proj.search_form_show_results
End Sub

'***route***
Private Sub route_cb_Click()
    Call enable_frame(Me.route_frame, Me.route_cb)
    Call proj.search_form_show_results
End Sub
Private Sub route_cbb_Change()
    Call proj.search_form_show_results
End Sub

'***ingoing***
Private Sub ingoing_cb_Click()
    Call enable_frame(Me.ingoing_frame, Me.ingoing_cb)
    Call proj.search_form_show_results
End Sub
Private Sub ingoing_cbb_Change()
    Call proj.search_form_show_results
End Sub

Private Sub enable_frame(frm As MSForms.Frame, switch As Boolean)
'will enable or disable a frame and all its controls
Dim ctr As MSForms.control
frm.Enabled = switch
For Each ctr In frm.Controls
    ctr.Enabled = switch
Next ctr

End Sub

Private Sub UserForm_Initialize()
    Call enable_frame(Me.period_frame, False)
    Call enable_frame(Me.boa_frame, False)
    Call enable_frame(Me.draught_frame, False)
    Call enable_frame(Me.loa_frame, False)
    Call enable_frame(Me.ship_type_frame, False)
    Call enable_frame(Me.voyage_succes_frame, False)
    Call enable_frame(Me.route_frame, False)
    Call enable_frame(Me.ingoing_frame, False)

Me.results_lb.ColumnCount = 9
Me.results_lb.ColumnWidths = "0;50;100;60;40;40;40;80;40"

End Sub
Private Sub UserForm_Click()

End Sub

