VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Clear_Dilution_Annot 
   Caption         =   "Clear Dilution_Annot"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3555
   OleObjectBlob   =   "Clear_Dilution_Annot.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Clear_Dilution_Annot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dilution_Annot_Clear_Click()
    If Data_File_Name.Value = True Then
        Call Utilities.Clear_Columns("Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
    If Sample_Name.Value = True Then
        Call Utilities.Clear_Columns("Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
    If Dilution_Batch_Name.Value = True Then
        Call Utilities.Clear_Columns("Dilution_Batch_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
    If Dilution_Factor.Value = True Then
        Call Utilities.Clear_Columns("Dilution_Factor_[%]", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
    If Injection_Volume_uL.Value = True Then
        Call Utilities.Clear_Columns("Injection_Volume_[uL]", HeaderRowNumber:=1, DataStartRowNumber:=2)
    End If
End Sub
