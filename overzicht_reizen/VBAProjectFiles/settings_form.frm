VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} settings_form 
   Caption         =   "Programma instellingen"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5595
   OleObjectBlob   =   "settings_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "settings_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub CommandButton1_Click()
    Call proj.settings_form_ok_click
End Sub

Private Sub CommandButton2_Click()
    unload Me
End Sub

Private Sub path_btn_admittance_template_Click()
    Me.path_tb_admittance_template.Value = aux_.get_single_file("Selecteer het sjabloon")
End Sub
Private Sub path_btn_hw_data_Click()
    Me.path_tb_hw_data.Value = aux_.get_single_file("Selecteer de database")
End Sub
Private Sub path_btn_Libdir_Click()
    Me.path_tb_Libdir.Value = aux_.get_single_file("Selecteer de map")
End Sub
Private Sub path_btn_sail_plan_archive_Click()
    Me.path_tb_sail_plan_archive.Value = aux_.get_single_file("Selecteer de database")
End Sub
Private Sub path_btn_sail_plan_db_Click()
    Me.path_tb_sail_plan_db.Value = aux_.get_single_file("Selecteer de database")
End Sub
Private Sub path_btn_tidal_data_Click()
    Me.path_tb_tidal_data.Value = aux_.get_single_file("Selecteer de database")
End Sub

Private Sub UserForm_Click()

End Sub
