VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub LoadSampleAnnotationRaw_Click()
    Call Load_Sample_Annot_Raw_Column_Name_Click
End Sub

Private Sub LoadSampleAnnotationTidy_Click()
    Call Load_Sample_Annot_Tidy_Column_Name_Click
End Sub

Private Sub AutofillSampleType_Click()
    Call Autofill_Sample_Type_Click
End Sub

Private Sub AutofillConcentrationUnit_Click()
    Call Autofill_Concentration_Unit_Click
End Sub

Private Sub ClearSampleTable_Click()
    Call Clear_Sample_Table_Click
End Sub

Private Sub LoadSampleNameToDilution_Click()
    Call Load_Sample_Name_To_Dilution_Annot_Click
End Sub

