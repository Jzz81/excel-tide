VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} stats_form 
   Caption         =   "Statistiek"
   ClientHeight    =   8370.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11325
   OleObjectBlob   =   "stats_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "stats_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim caller_ctr As MSForms.control
Private WithEvents cal As cCalendar
Attribute cal.VB_VarHelpID = -1
Public start_date As Date
Public end_date As Date

Private Sub cal_Click()
caller_ctr.Text = cal.Value

End Sub

Private Sub cal_DblClick()
caller_ctr.Text = cal.Value
Call destroy_datepicker

End Sub


Private Sub cal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyEscape Then
    Call destroy_datepicker
End If

End Sub

Private Sub destroy_datepicker()
Me.datepicker_frame.Visible = False
Set cal = Nothing
Do While Me.datepicker_frame.Controls.Count > 0
    Me.datepicker_frame.Controls.Remove (0)
Loop
'restore backcolor
caller_ctr.BackColor = -2147483643

End Sub
Private Sub create_datepicker()
'in this case, all controls that create the datepicker are on the multipage.
'multipage behaves strangely if global position is calculated. Probably the
'tabstrip that is not properly calculated.
Set cal = New cCalendar
Dim T As Double
Dim L As Double
Dim ctr As MSForms.control
Set ctr = caller_ctr
On Error Resume Next
    Do Until ctr.parent.Name = Me.Name
        T = T + ctr.Top
        L = L + ctr.Left
        Set ctr = ctr.parent
    Loop
On Error GoTo 0
'add 15 to compensate for the tabstrip (see above)
T = T + caller_ctr.Height + 15

With Me.datepicker_frame
    .Visible = True
    .Top = T
    .Left = L
    .ZOrder (0)
End With
'set red color to control
caller_ctr.BackColor = vbRed

cal.Today
cal.Add_Calendar_into_Frame Me.datepicker_frame
End Sub




Private Sub CommandButton1_Click()

End Sub

Private Sub dt_end_tb_Enter()
Set caller_ctr = Me.dt_end_tb
Call create_datepicker

End Sub

Private Sub dt_end_tb_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Not cal Is Nothing Then Call destroy_datepicker

End Sub

Private Sub dt_start_tb_Enter()
Set caller_ctr = Me.dt_start_tb
Call create_datepicker

End Sub


Private Sub dt_start_tb_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Not cal Is Nothing Then Call destroy_datepicker

End Sub
Public Sub sort_listbox(lb As MSForms.control)
'Sorts ListBox List
Dim i As Long
Dim j As Long
Dim temp As String
    
With lb
    For j = 0 To .ListCount - 2
        For i = 0 To .ListCount - 2
            If .List(i) > .List(i + 1) Then
                temp = .List(i)
                .List(i) = .List(i + 1)
                .List(i + 1) = temp
            End If
        Next i
    Next j
End With
End Sub

Private Sub Frame1_Click()

End Sub


Private Sub ok_btn_Click()
Call proj.stats_form_ok_click
End Sub

Private Sub save_coll_btn_Click()
Call proj.stats_form_save_collection_click
End Sub

Private Sub search_sp_btn_Click()
Call proj.stats_form_search_sp_click

End Sub

Private Sub UserForm_Initialize()
Me.datepicker_frame.Visible = False
Me.save_coll_lb.ColumnCount = 2
Me.save_coll_lb.ColumnWidths = ";0"

End Sub


Private Sub voy_add_col_btn_Click()
Dim i As Long

If Me.voyage_lb.ListIndex < 0 Then Exit Sub

For i = Me.voyage_lb.ListCount - 1 To 0 Step -1
    If Me.voyage_lb.Selected(i) = True Then
        Me.collection_lb.AddItem Me.voyage_lb.List(i)
        Me.voyage_lb.RemoveItem (i)
    End If
Next i
Call Me.sort_listbox(Me.collection_lb)

End Sub

Private Sub voy_del_col_btn_Click()
Dim i As Long

If Me.collection_lb.ListIndex < 0 Then Exit Sub

For i = Me.collection_lb.ListCount - 1 To 0 Step -1
    If Me.collection_lb.Selected(i) Then
        Me.voyage_lb.AddItem Me.collection_lb(i)
        Me.collection_lb.RemoveItem (i)
    End If
Next i

Call Me.sort_listbox(Me.voyage_lb)

End Sub

Private Sub voyage_lb_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'will add the selected voyage to the collection listbox
If Me.voyage_lb.ListIndex < 0 Then Exit Sub
'add to collection lb
    Me.collection_lb.AddItem Me.voyage_lb.List(Me.voyage_lb.ListIndex)
'remove from voyage lb
    Me.voyage_lb.RemoveItem (Me.voyage_lb.ListIndex)
'sort collection lb
    Call Me.sort_listbox(Me.collection_lb)
End Sub
