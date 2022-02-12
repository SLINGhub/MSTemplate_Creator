VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISTDAnnot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Call ColourTracker.ISTD_Calculation_Checker(Target)
End Sub

Private Sub ClearISTDTable_Click()
    Call Clear_ISTD_Table_Click
End Sub

Private Sub nMcalculation_Click()
    Call nM_calculation_Click
End Sub

