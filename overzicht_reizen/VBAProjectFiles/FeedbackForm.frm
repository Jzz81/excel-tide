VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FeedbackForm 
   Caption         =   "UserForm1"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4680
   OleObjectBlob   =   "FeedbackForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FeedbackForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit
Public cancelflag As Boolean

Private Sub CancelBTN_Click()
cancelflag = True
End Sub
