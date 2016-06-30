VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} deviations_validation_form 
   Caption         =   "Controleer de afwijkingen"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5610
   OleObjectBlob   =   "deviations_validation_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "deviations_validation_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ok_btn_Click()
Dim ctr As MSForms.control
Dim B As Boolean

B = False

For Each ctr In Me.Controls
    If InStr(1, ctr.Name, "tb_", vbTextCompare) <> 0 Then
        If Not IsNumeric(ctr.text) Then
            ctr.BackColor = vbRed
            B = True
        Else
            ctr.BackColor = vbGreen
        End If
    End If
Next ctr

If B Then
    MsgBox "Er zijn ongeldige waarden aangetroffen!", vbExclamation
    Exit Sub
End If
    
Me.Hide
End Sub

Private Sub print_btn_Click()
Me.PrintForm
End Sub

