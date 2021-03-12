Attribute VB_Name = "Dilution_Annot_Buttons"
'Sheet Dilution_Annot Functions

Sub Clear_Dilution_Annot_Click()
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    Clear_Dilution_Annot.Show
End Sub
