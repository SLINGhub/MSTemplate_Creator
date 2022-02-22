VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Overwrite 
   Caption         =   "Overwrite"
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   2790
   OleObjectBlob   =   "Overwrite.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Overwrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public whatsclicked As String

Private Sub Transition_Name_Cancel_Click()
    whatsclicked = "Cancel"
    Overwrite.Hide
End Sub

Private Sub Transition_Name_Overwrite_Click()
    whatsclicked = "Overwrite"
    Overwrite.Hide
End Sub

