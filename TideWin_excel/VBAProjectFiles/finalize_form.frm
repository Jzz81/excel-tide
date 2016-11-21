VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} finalize_form 
   Caption         =   "reis finalizeren"
   ClientHeight    =   4215
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
cancelflag = False
Call proj.finalize_form_ok_click
End Sub

Private Sub planning_ob_no_Change()
Dim L As Long
L = 50
If Me.planning_ob_no.Value Then
    Me.reason_frame.Visible = True
    Me.ata_frame.Top = Me.ata_frame.Top + L
    Me.ok_btn.Top = Me.ok_btn.Top + L
    Me.cancel_btn.Top = Me.cancel_btn.Top + L
    Me.remarks_frame.Top = Me.remarks_frame.Top + L
    Me.Height = Me.Height + L
Else
    Me.reason_frame.Visible = False
    Me.ata_frame.Top = Me.ata_frame.Top - L
    Me.ok_btn.Top = Me.ok_btn.Top - L
    Me.cancel_btn.Top = Me.cancel_btn.Top - L
    Me.remarks_frame.Top = Me.remarks_frame.Top - L
    Me.Height = Me.Height - L
End If

End Sub

Private Sub UserForm_Initialize()
cancelflag = True
Me.reason_frame.Visible = False

End Sub
