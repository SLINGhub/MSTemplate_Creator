VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SampleAnnotSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@IgnoreModule ProcedureNotUsed
'@Folder("Sheets")

Private Sub AutofillBySampleType_Click()
    Autofill_By_Sample_Type_Click
End Sub

Private Sub LoadSampleAnnotationRaw_Click()
    Load_Sample_Annot_Raw_Column_Name_Click
End Sub

Private Sub LoadSampleAnnotationTidy_Click()
    Load_Sample_Annot_Tidy_Column_Name_Click
End Sub

Private Sub AutofillSampleType_Click()
    Autofill_Sample_Type_Click
End Sub

Private Sub AutofillConcentrationUnit_Click()
    Autofill_Concentration_Unit_Click
End Sub

Private Sub ClearSampleTable_Click()
    Clear_Sample_Table_Click
End Sub

Private Sub LoadSampleNameToDilution_Click()
    Load_Sample_Name_To_Dilution_Annot_Click
End Sub

