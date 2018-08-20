VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Call ColourTracker.ChangeToBlankWhenChanged(Target)
End Sub

Private Sub ClearTransitionSheet_Click()
    Call Clear_Transition_Name_Annot_Click
End Sub

Private Sub LoadSampleAnnotation_Click()
    'In sheet 3 functions
    Call Load_Sample_Annot_Column_Name_Click
End Sub

Private Sub LoadISTD_Click()
    Call Load_Transition_Name_ISTD_Click
End Sub

Private Sub ValidateISTD_Click()
    Call Validate_ISTD_Click
End Sub

Private Sub LoadTransition_Click()
    Call GetTransitionArray_Click
End Sub

