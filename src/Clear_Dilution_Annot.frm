VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Clear_Dilution_Annot 
   Caption         =   "Clear Dilution_Annot"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   3825
   OleObjectBlob   =   "Clear_Dilution_Annot.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Clear_Dilution_Annot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@Folder("Dilution_Annot_Buttons")

'' Function: Dilution_Annot_Clear_Click
''
'' Description:
''
'' Function that controls what happens when the Clear Data button is
'' left clicked
''
'' (see Dilution_Annot_Clear_Data_Button.png)
''
'' All data found in the columns that was checked will be cleared.
''
Private Sub Dilution_Annot_Clear_Click()
    If Data_File_Name.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Data_File_Name", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
    If Sample_Name.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Sample_Name", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
    If Dilution_Batch_Name.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Dilution_Batch_Name", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
    If Relative_Sample_Amount.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Relative_Sample_Amount_[%]", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
    If Injection_Volume_uL.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Injection_Volume_[uL]", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
End Sub

