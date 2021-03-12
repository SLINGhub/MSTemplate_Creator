VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Call ISTDCalculationChecker(Target)
End Sub

Private Sub ClearISTDTable_Click()
    Call Clear_ISTD_Table_Click
End Sub

Private Sub nMcalculation_Click()
    Call nM_calculation_Click
End Sub

