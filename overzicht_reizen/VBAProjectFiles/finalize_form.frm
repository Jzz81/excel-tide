VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} finalize_form 
   Caption         =   "reis finalizeren"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   OleObjectBlob   =   "finalize_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "finalize_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cancelflag As Boolean
Private Sub cancel_btn_Click()
cancelflag = True
Me.Hide
End Sub

Private Sub ok_btn_Click()
Call proj.finalize_form_ok_click
End Sub
