VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} routes_edit_form 
   Caption         =   "Routes"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8925
   OleObjectBlob   =   "routes_edit_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "routes_edit_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim loading As Boolean

Private Sub CommandButton1_Click()
    Call proj.routes_form_insert_click
End Sub

Private Sub CommandButton2_Click()
    Call proj.routes_form_save_click
End Sub

Private Sub CommandButton3_Click()
    Call proj.routes_form_delete_treshold_click
End Sub

Private Sub CommandButton4_Click()
    Call proj.routes_form_delete_route_click
End Sub

Private Sub CommandButton5_Click()
    Call proj.routes_form_new_route_click
End Sub

Private Sub routes_lb_Click()
    Call proj.routes_form_routes_lb_click
End Sub

Private Sub speeds_cb_Click()
    If Not loading Then Call proj.routes_form_speeds_cb_change
End Sub

Private Sub treshold_cb_Click()
    If Not loading Then Call proj.routes_form_treshold_cb_change
End Sub

Private Sub tresholds_lb_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    loading = True
    Call proj.routes_form_tresholds_lb_click
    loading = False
End Sub

Private Sub UKC_unit_cb_Change()
    Call proj.routes_form_UKC_unit_cb_change
End Sub

Private Sub UKC_value_tb_Change()
    Call proj.routes_form_UKC_value_tb_change
End Sub

Private Sub UserForm_Initialize()
Me.routes_lb.ColumnCount = 2
Me.routes_lb.ColumnWidths = "0;"

Me.tresholds_lb.ColumnCount = 5
Me.tresholds_lb.ColumnWidths = "50;25;16;;0"

Me.ukc_unit_cb.AddItem "%"
Me.ukc_unit_cb.AddItem "m"
Me.ukc_unit_cb.Value = "%"

Me.speeds_cb.ColumnCount = 2
Me.speeds_cb.ColumnWidths = "0;"

Me.treshold_cb.ColumnCount = 2
Me.treshold_cb.ColumnWidths = ";0"

End Sub
