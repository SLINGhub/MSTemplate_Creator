Attribute VB_Name = "ISTD_Annot_Buttons"
Option Explicit
'@Folder("ISTD Annot Functions")
'@IgnoreModule IntegerDataType

'' Function: Clear_ISTD_Table_Click
'' --- Code
''  Public Sub Clear_ISTD_Table_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see ISTD_Annot_Clear_Columns_Button.png)
''
'' The following pop up box will appear. Asking the users
'' which column to clear.
''
'' (see ISTD_Annot_Clear_Data_Pop_Up.png)
''
Public Sub Clear_ISTD_Table_Click()
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    Clear_ISTD_Table.Show
End Sub

'' Function: Convert_To_Nanomolar_Click
'' --- Code
''  Public Sub Convert_To_Nanomolar_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see ISTD_Annot_Convert_To_Nanomolar_Button.png)
''
'' ISTD_Conc_[nM] and Custom_Units will be calculated
'' abnd output to their repective columns.
''
'' The ISTD_Conc_[nM] columns and  Custom_Units cells will
'' be coloured green when when the row has either a valid
'' Transition_Name_ISTD, ISTD_Conc_[ng/mL] and ISTD_[MW]
'' or a valid Transition_Name_ISTD and ISTD_Conc_[nM]
''
'' Columns and rows will turn green when the button
'' "Convert to nM and Verify" is pressed when either
'' both ISTD_Conc_[nM] and ISTD_[MW] is entered or
'' the ISTD_Conc_[nM] is entered. Custom units will
'' be automatically calculated.
''
'' (see ISTD_Annot_Press_Convert_Button.png)
''
Public Sub Convert_To_Nanomolar_Click()
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    
    ' Get the ISTD_Annot worksheet from the active workbook
    ' The ISTDAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim ISTD_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "ISTDAnnotSheet") = False Then
        MsgBox ("Sheet ISTD_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set ISTD_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "ISTDAnnotSheet")
      
    ISTD_Annot_Worksheet.Activate
        
    Dim ISTD_Custom_Unit_ColNumber As Integer
    ISTD_Custom_Unit_ColNumber = Utilities.Get_Header_Col_Position("Custom_Unit", 2)
    
    Dim Custom_Unit As String
    Custom_Unit = ISTD_Annot_Worksheet.Cells.Item(3, ISTD_Custom_Unit_ColNumber)
    
    Dim ISTD_Conc_nM() As String
    Dim ISTD_Custom_Unit() As String
    ISTD_Conc_nM = ISTD_Annot.Get_ISTD_Conc_nM_Array(ColourCellRequired:=True)
    
    Utilities.Load_To_Excel Data_Array:=ISTD_Conc_nM, _
                            HeaderName:="ISTD_Conc_[nM]", _
                            HeaderRowNumber:=3, _
                            DataStartRowNumber:=4, _
                            MessageBoxRequired:=False
                            
    ISTD_Custom_Unit = ISTD_Annot.Convert_Conc_nM_Array(Custom_Unit)
    
    Utilities.Load_To_Excel Data_Array:=ISTD_Custom_Unit, _
                            HeaderName:="Custom_Unit", _
                            HeaderRowNumber:=2, _
                            DataStartRowNumber:=4, _
                            MessageBoxRequired:=False
    
    'Resume monitoring of sheet
    Application.EnableEvents = True
End Sub
