VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Clear_Transition_Name_Annot 
   Caption         =   "Clear Transition_Name_Annot"
   ClientHeight    =   2355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3660
   OleObjectBlob   =   "Clear_Transition_Name_Annot.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Clear_Transition_Name_Annot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Transition_Name_Annot_Clear_Click()
    If Transition_Name.Value = True Then
        Call Utilities.Clear_Columns("Transition_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
    If Transition_Name_ISTD.Value = True Then
        Call Utilities.Clear_Columns("Transition_Name_ISTD", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
End Sub

