VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} connections_edit_form 
   Caption         =   "verbindingen"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8100
   OleObjectBlob   =   "connections_edit_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "connections_edit_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub CommandButton1_Click()
    Call proj.connection_form_save_click
End Sub

Private Sub CommandButton2_Click()
    Call proj.connection_form_del_click
End Sub

Private Sub conn_lb_Click()
    Call proj.connection_form_lb_click
End Sub

Private Sub UserForm_Initialize()
Me.conn_lb.ColumnCount = 3
Me.conn_lb.ColumnWidths = "0;100;20"
End Sub
