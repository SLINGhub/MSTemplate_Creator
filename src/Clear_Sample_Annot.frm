VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Clear_Sample_Annot 
   Caption         =   "Clear Sample_Annot"
   ClientHeight    =   4695
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   4245
   OleObjectBlob   =   "Clear_Sample_Annot.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Clear_Sample_Annot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@Folder("Sample_Annot_Buttons")

'' Function: Sample_Annot_Clear_Click
'' --- Code
''  Private Sub Sample_Annot_Clear_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the Clear Data button is
'' left clicked
''
'' (see Sample_Annot_Clear_Data_Button.png)
''
'' All data found in the columns that was checked will be cleared.
''
Private Sub Sample_Annot_Clear_Click()
    If Data_File_Name.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Data_File_Name", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
    If Merge_Status.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Merge_Status", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
    If Sample_Name.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Sample_Name", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
    If Sample_Type.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Sample_Type", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
    If Sample_Amount.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Sample_Amount", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
    If Sample_Amount_Unit.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Sample_Amount_Unit", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
    If ISTD_Mixture_Volume.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="ISTD_Mixture_Volume_[ul]", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
    If Concentration_Unit.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Concentration_Unit", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
End Sub
