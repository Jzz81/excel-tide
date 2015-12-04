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


Private Sub SaveBtn_Click()
    Call proj.treshold_form_save_click
End Sub
Private Sub tresholds_lb_Click()
    Call proj.treshold_form_listbox_click
End Sub
Private Sub UserForm_Initialize()
    Me.tresholds_lb.ColumnCount = 2
    Me.tresholds_lb.ColumnWidths = "0;"
    
    Me.UKC_unit_cb.AddItem "m"
    Me.UKC_unit_cb.AddItem "%"
    Me.UKC_unit_cb.Value = "%"
End Sub
