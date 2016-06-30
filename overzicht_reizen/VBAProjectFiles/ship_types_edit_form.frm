VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ship_types_edit_form 
   Caption         =   "Scheepstypes"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7770
   OleObjectBlob   =   "ship_types_edit_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ship_types_edit_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Option Base 0
Option Compare Text

Private Sub CommandButton1_Click()
    Call proj.ship_type_form_del_click
End Sub

Private Sub SaveBtn_Click()
    Call proj.ship_type_form_save_click
End Sub

Private Sub ship_types_lb_Click()
    Call proj.ship_type_form_listbox_click
End Sub

Private Sub UserForm_Initialize()
    Me.ship_types_lb.ColumnCount = 2
    Me.ship_types_lb.ColumnWidths = "0;"
End Sub
