VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Clear_Sample_Annot 
   Caption         =   "Clear Sample_Annot"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   OleObjectBlob   =   "Clear_Sample_Annot.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Clear_Sample_Annot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Sample_Annot_Clear_Click()
    If Data_File_Name.Value = True Then
        Call Utilities.Clear_Columns("Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
    If Merge_Status.Value = True Then
        Call Utilities.Clear_Columns("Merge_Status", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
    If Sample_Name.Value = True Then
        Call Utilities.Clear_Columns("Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
    If Sample_Type.Value = True Then
        Call Utilities.Clear_Columns("Sample_Type", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
    If Sample_Amount.Value = True Then
        Call Utilities.Clear_Columns("Sample_Amount", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
    If Sample_Amount_Unit.Value = True Then
        Call Utilities.Clear_Columns("Sample_Amount_Unit", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
    If ISTD_Mixture_Volume.Value = True Then
        Call Utilities.Clear_Columns("ISTD_Mixture_Volume_[ul]", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
    If Concentration_Unit.Value = True Then
        Call Utilities.Clear_Columns("Concentration_Unit", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
End Sub
