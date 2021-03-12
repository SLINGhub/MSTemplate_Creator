VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Clear_ISTD_Table 
   Caption         =   "Clear ISTD_Table"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3540
   OleObjectBlob   =   "Clear_ISTD_Table.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Clear_ISTD_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ISTD_Table_Clear_Click()
    If Transition_Name_ISTD.Value = True Then
        Call Utilities.Clear_Columns("Transition_Name_ISTD", HeaderRowNumber:=2, DataStartRowNumber:=4)
    End If
    If ISTD_Conc_ngmL.Value = True Then
        Call Utilities.Clear_Columns("ISTD_Conc_[ng/mL]", HeaderRowNumber:=3, DataStartRowNumber:=4)
    End If
    If ISTD_MW.Value = True Then
        Call Utilities.Clear_Columns("ISTD_[MW]", HeaderRowNumber:=3, DataStartRowNumber:=4)
    End If
    If ISTD_Conc_nM.Value = True Then
        Call Utilities.Clear_Columns("ISTD_Conc_[nM]", HeaderRowNumber:=3, DataStartRowNumber:=4)
        Call Utilities.Clear_Columns("Custom_Unit", HeaderRowNumber:=2, DataStartRowNumber:=4)
    End If
End Sub

