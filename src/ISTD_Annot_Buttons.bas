Attribute VB_Name = "ISTD_Annot_Buttons"
'Sheet ISTD_Annot Functions

Sub Clear_ISTD_Table_Click()
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    Clear_ISTD_Table.Show
End Sub

Sub nM_calculation_Click()
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    Sheets("ISTD_Annot").Activate
    
    Dim ISTD_Custom_Unit_ColNumber As Integer
    ISTD_Custom_Unit_ColNumber = Utilities.Get_Header_Col_Position("Custom_Unit", 2)
    
    Dim Custom_Unit As String
    Custom_Unit = Cells(3, ISTD_Custom_Unit_ColNumber)
    
    Dim ISTD_Conc_nM() As String
    Dim ISTD_Custom_Unit() As String
    ISTD_Conc_nM = ISTD_Annot.Get_ISTD_Conc_nM_Array(ColourCellRequired:=True)
    Call Utilities.Load_To_Excel(ISTD_Conc_nM, "ISTD_Conc_[nM]", HeaderRowNumber:=3, DataStartRowNumber:=4, MessageBoxRequired:=False)
    ISTD_Custom_Unit = ISTD_Annot.Convert_Conc_nM_Array(Custom_Unit)
    Call Utilities.Load_To_Excel(ISTD_Custom_Unit, "Custom_Unit", HeaderRowNumber:=2, DataStartRowNumber:=4, MessageBoxRequired:=False)
    
    'Resume monitoring of sheet
    Application.EnableEvents = True
End Sub
