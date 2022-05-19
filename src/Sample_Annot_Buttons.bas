Attribute VB_Name = "Sample_Annot_Buttons"
Attribute VB_Description = "Functions that control what happens when buttons in the Sample_Annot worksheet are clicked."
Option Explicit
'@ModuleDescription("Functions that control what happens when buttons in the Sample_Annot worksheet are clicked.")

'@Folder("Sample Annot Functions")
'@IgnoreModule IntegerDataType
'@Description("Function that controls what happens when the Autofill By Sample Type button is left clicked.")

'' Function: Autofill_By_Sample_Type_Click
'' --- Code
''  Public Sub Autofill_By_Sample_Type_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Sample_Annot_Autofill_By_Sample_Type_Button.png)
''
'' The following pop up box will appear. Asking the users
'' which column to clear.
''
'' (see Sample_Annot_Autofill_By_Sample_Type_Pop_Up.png)
''
Public Sub Autofill_By_Sample_Type_Click()
Attribute Autofill_By_Sample_Type_Click.VB_Description = "Function that controls what happens when the Autofill By Sample Type button is left clicked."

    ' Get the Sample_Annot worksheet from the active workbook
    ' The SampleAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Sample_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "SampleAnnotSheet") = False Then
        MsgBox ("Sheet Sample_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Sample_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "SampleAnnotSheet")
      
    Sample_Annot_Worksheet.Activate
       
    Autofill_By_Sample_Type.Show

    Unload Autofill_By_Sample_Type
     
End Sub

'@Description("Function that controls what happens when the Load RQC Sample Name To Dilution_Table button is left clicked.")

'' Function: Load_Sample_Name_To_Dilution_Annot_Click
'' --- Code
''  Public Sub Load_Sample_Name_To_Dilution_Annot_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Sample_Annot_Load_Sample_Name_To_Dilution_Annot_Button.png)
''
'' The system will copy all RQC samples into the sheet "Dilution_Annot"
''
'' (see Sample_Annot_Load_Sample_Name_To_Dilution_Annot_Results.png)
''
Public Sub Load_Sample_Name_To_Dilution_Annot_Click()
Attribute Load_Sample_Name_To_Dilution_Annot_Click.VB_Description = "Function that controls what happens when the Load RQC Sample Name To Dilution_Table button is left clicked."
    
    ' Get the Sample_Annot worksheet from the active workbook
    ' The SampleAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Sample_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "SampleAnnotSheet") = False Then
        MsgBox ("Sheet Sample_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Sample_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "SampleAnnotSheet")
      
    Sample_Annot_Worksheet.Activate
    
    'To ensure that Filters does not affect the assignment
    Utilities.Remove_Filter_Settings
    
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    
    Dim SampleNameArray() As String
    Dim FileNameArray() As String
        
    'Check if the column Sample_Type exists
    Dim SampleType_pos As Integer
    SampleType_pos = Utilities.Get_Header_Col_Position("Sample_Type", HeaderRowNumber:=1)
    
    'Filter Rows by "RQC"
    Sample_Annot_Worksheet.Range("A1").AutoFilter Field:=SampleType_pos, _
                                                  Criteria1:="RQC", _
                                                  VisibleDropDown:=True
                                       
                                       

    'Load the Sample_Name columns content from Sample_Annot
    SampleNameArray = Utilities.Load_Columns_From_Excel(HeaderName:="Sample_Name", HeaderRowNumber:=1, _
                                                        DataStartRowNumber:=2, MessageBoxRequired:=True, _
                                                        RemoveBlksAndReplicates:=False, _
                                                        IgnoreHiddenRows:=True, IgnoreEmptyArray:=True)
                                                    
    'Load the Data_File_Name columns content from Sample_Annot
    FileNameArray = Utilities.Load_Columns_From_Excel(HeaderName:="Data_File_Name", HeaderRowNumber:=1, _
                                                      DataStartRowNumber:=2, MessageBoxRequired:=False, _
                                                      RemoveBlksAndReplicates:=False, _
                                                      IgnoreHiddenRows:=True, IgnoreEmptyArray:=True)

    'Debug.Print FileNameArray(1)
    
    'To ensure that Filters does not affect the assignment
    Utilities.Remove_Filter_Settings
    
    'Resume monitoring of sheet
    Application.EnableEvents = True
                                                           
    'Check if SampleNameArray has any elements
    'If not no need to transfer
    If Len(Join(SampleNameArray, vbNullString)) = 0 Then
        Exit Sub
    End If
    
    ' Get the Dilution_Annot worksheet from the active workbook
    ' The DilutionAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Dilution_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "DilutionAnnotSheet") = False Then
        MsgBox ("Sheet Dilution_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Dilution_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "DilutionAnnotSheet")
      
    Dilution_Annot_Worksheet.Activate
    
    'Check if FileNameArray has any elements
    If Len(Join(FileNameArray, vbNullString)) > 0 Then
        Utilities.Overwrite_Header HeaderName:="Data_File_Name", _
                                  HeaderRowNumber:=1, _
                                  DataStartRowNumber:=2
        Utilities.Load_To_Excel Data_Array:=FileNameArray, _
                                HeaderName:="Data_File_Name", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2, _
                                MessageBoxRequired:=False
        Utilities.Overwrite_Header HeaderName:="Sample_Name", _
                                  HeaderRowNumber:=1, _
                                  DataStartRowNumber:=2
        Utilities.Load_To_Excel Data_Array:=SampleNameArray, _
                                HeaderName:="Sample_Name", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2, _
                                MessageBoxRequired:=True
    Else
        Utilities.Overwrite_Header HeaderName:="Sample_Name", _
                                  HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
        Utilities.Load_To_Excel Data_Array:=SampleNameArray, _
                                HeaderName:="Sample_Name", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2, _
                                MessageBoxRequired:=True
    End If
    
End Sub

'@Description("Function that controls what happens when the Clear Columns button is left clicked.")

'' Function: Clear_Sample_Table_Click
'' --- Code
''  Public Sub Clear_Sample_Table_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Sample_Annot_Clear_Columns_Button.png)
''
'' The following pop up box will appear. Asking the users
'' which column to clear.
''
'' (see Sample_Annot_Clear_Data_Pop_Up.png)
''
Public Sub Clear_Sample_Table_Click()
Attribute Clear_Sample_Table_Click.VB_Description = "Function that controls what happens when the Clear Columns button is left clicked."
    'To ensure that Filters does not affect the assignment
    Utilities.Remove_Filter_Settings
    
    Clear_Sample_Annot.Show
End Sub

'@Description("Function that controls what happens when the Autofill Concentration Unit button is left clicked.")

'' Function: Autofill_Concentration_Unit_Click
'' --- Code
''  Public Sub Autofill_Concentration_Unit_Click(Optional ByVal MessageBoxRequired As Boolean = True, _
''                                               Optional ByVal Testing As Boolean = False)
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Sample_Annot_Autofill_Concentration_Unit_Button.png)
''
'' If there are some entry in the Sample_Amount_Unit column,
'' the program will use the custom unit section from the ISTD_Annot sheet
'' to create the following concentration units.
''
'' (see Sample_Annot_Autofill_Concentration_Unit_Results.png)
''
'' Concentration units will be updated
'' if there are some entry in the Sample_Amount_Unit column
'' and the custom unit section from the ISTD_Annot sheet is changed.
''
Public Sub Autofill_Concentration_Unit_Click(Optional ByVal MessageBoxRequired As Boolean = True, _
                                             Optional ByVal Testing As Boolean = False)
Attribute Autofill_Concentration_Unit_Click.VB_Description = "Function that controls what happens when the Autofill Concentration Unit button is left clicked."
                                      
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    
    'Check if the column Custom_Unit exists
    Dim ISTD_Custom_Unit_ColNumber As Integer
    ISTD_Custom_Unit_ColNumber = Utilities.Get_Header_Col_Position("Custom_Unit", 2, _
                                                                   WorksheetName:="ISTD_Annot")
                                                                   
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
                                                                   
    'Get the mol concentration from custom unit value
    Dim Custom_Unit As String
    Custom_Unit = ISTD_Annot_Worksheet.Cells.Item(3, ISTD_Custom_Unit_ColNumber)
    Application.EnableEvents = True
    
    'Custom Unit Value is of the form "[?M] or [?mol/uL]"
    'Function tries to get ?mol from the above string
    Dim Right_Custom_Unit As String
    Right_Custom_Unit = Concentration_Unit.Get_Mol_From_Custom_ISTD_Concentration_Unit(Custom_Unit)
    'Debug.Print Right_Custom_Unit

    'To ensure that Filters does not affect the assignment
    Utilities.Remove_Filter_Settings
    
    Dim Sample_Amount_Unit() As String
    Sample_Amount_Unit = Utilities.Load_Columns_From_Excel(HeaderName:="Sample_Amount_Unit", HeaderRowNumber:=1, _
                                                           DataStartRowNumber:=2, MessageBoxRequired:=False, _
                                                           RemoveBlksAndReplicates:=False, _
                                                           WorksheetName:="Sample_Annot", _
                                                           IgnoreHiddenRows:=False, IgnoreEmptyArray:=True)
    'Get the length of Sample_Amount_Unit
    Dim max_length As Integer
    max_length = 0
    If Utilities.Get_String_Array_Len(Sample_Amount_Unit) > max_length Then
            max_length = Utilities.Get_String_Array_Len(Sample_Amount_Unit)
    End If
    
    'Leave the program if max_length is 0
    If max_length = 0 Then
        'Application.EnableEvents = True
        Exit Sub
    End If
    'Else we proceed to update the concetration unit
    
    'If the active sheet is ISTD_Annot,
    'inform the users that concentration unit must be updated
    If ActiveSheet.Name = "ISTD_Annot" And MessageBoxRequired = True Then
        MsgBox "Updating Concentration_Unit in Sample_Annot " & _
               "as at least one row in the Sample Amount Unit " & _
               "column is filled"
    End If
    
    Dim ConcentrationUnitArray() As String
    Dim UniqueConcentrationUnitArray() As String
    Dim UniqueArraryLength As Integer
    UniqueArraryLength = 0
    'Resize the array to max_length
    ReDim Preserve ConcentrationUnitArray(max_length)
    
    Dim lenArrayIndex As Integer
    Dim InArray As Boolean
    
    'Add the concentration unit when necessary
    For lenArrayIndex = 0 To max_length - 1
        Dim ConcentrationUnit As String
        If Len(Sample_Amount_Unit(lenArrayIndex)) <> 0 Then
            ConcentrationUnit = Right_Custom_Unit & "/" & Sample_Amount_Unit(lenArrayIndex)
            ConcentrationUnitArray(lenArrayIndex) = ConcentrationUnit
            
            'Collect Unique concentration unit
            InArray = Utilities.Is_In_Array(ConcentrationUnit, UniqueConcentrationUnitArray)
            If Not InArray Then
                ReDim Preserve UniqueConcentrationUnitArray(UniqueArraryLength)
                UniqueConcentrationUnitArray(UniqueArraryLength) = ConcentrationUnit
                'Debug.Print UniqueConcentrationUnitArray(UniqueArraryLength)
                UniqueArraryLength = UniqueArraryLength + 1
            End If
            
        End If
    Next lenArrayIndex
    
    'Load to Excel
    Utilities.Overwrite_Header HeaderName:="Concentration_Unit", _
                              HeaderRowNumber:=1, _
                              DataStartRowNumber:=2, _
                              WorksheetName:="Sample_Annot"
    Utilities.Load_To_Excel Data_Array:=ConcentrationUnitArray, _
                            HeaderName:="Concentration_Unit", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            WorksheetName:="Sample_Annot", _
                            MessageBoxRequired:=False
                                 
    'Display a summary box of unique concentration units
    If Utilities.Get_String_Array_Len(UniqueConcentrationUnitArray) <> 0 Then
        'Put the concentration units in the list box to be displayed
        For lenArrayIndex = 0 To UBound(UniqueConcentrationUnitArray) - LBound(UniqueConcentrationUnitArray)
            Concentration_Unit_MsgBox.Concentration_Unit_ListBox.AddItem UniqueConcentrationUnitArray(lenArrayIndex)
        Next lenArrayIndex
        Concentration_Unit_MsgBox.Show
        If Testing Then
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    
End Sub

'@Description("Function that controls what happens when the Autofill Sample Type button is left clicked.")

'' Function: Autofill_Sample_Type_Click
'' --- Code
''  Public Sub Autofill_Sample_Type_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Sample_Annot_Autofill_Sample_Type_Button.png)
''
'' Any blank cells in the Sample_Type columns with a Sample_Name
'' will be filled with sample type "SPL"
''
'' (see Sample_Annot_Autofill_Sample_Type_Results.png)
''
Public Sub Autofill_Sample_Type_Click()
Attribute Autofill_Sample_Type_Click.VB_Description = "Function that controls what happens when the Autofill Sample Type button is left clicked."

    ' Get the Sample_Annot worksheet from the active workbook
    ' The SampleAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Sample_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "SampleAnnotSheet") = False Then
        MsgBox ("Sheet Sample_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Sample_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "SampleAnnotSheet")
      
    Sample_Annot_Worksheet.Activate
    
    'To ensure that Filters does not affect the assignment
    Utilities.Remove_Filter_Settings
    
    Dim SampleArray() As String
    Dim TotalRows As Long
    Dim SampleArrayIndex As Long
    
    'Check if the column Sample_Name exists
    Dim SampleName_pos As Integer
    SampleName_pos = Utilities.Get_Header_Col_Position("Sample_Name", HeaderRowNumber:=1)
    
    'Check if the column Sample_Type exists
    Dim SampleType_pos As Integer
    SampleType_pos = Utilities.Get_Header_Col_Position("Sample_Type", HeaderRowNumber:=1)
   
    'Find the total number of rows and resize the array accordingly
    TotalRows = Sample_Annot_Worksheet.Cells.Item(Sample_Annot_Worksheet.Rows.Count, Utilities.Convert_To_Letter(SampleName_pos)).End(xlUp).Row
    ReDim SampleArray(0 To TotalRows - 1)
    
    'Assign "Sample" if there is no sample type
    If TotalRows > 1 Then
        For SampleArrayIndex = 2 To TotalRows
            If Sample_Annot_Worksheet.Cells.Item(SampleArrayIndex, SampleType_pos).Value = vbNullString Then
                SampleArray(SampleArrayIndex - 2) = "SPL"
            Else
                SampleArray(SampleArrayIndex - 2) = Sample_Annot_Worksheet.Cells.Item(SampleArrayIndex, SampleType_pos).Value
            End If
            'Debug.Print SampleArray(i - 2)
        Next SampleArrayIndex
    End If
    
    Utilities.Load_To_Excel Data_Array:=SampleArray, _
                            HeaderName:="Sample_Type", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
    'Range(Utilities.Convert_To_Letter(SampleType_pos) & "2").Resize(UBound(SampleArray) + 1) = Application.Transpose(SampleArray)

End Sub

'@Description("Function that controls what happens when the Load Sample Annotation from Table Data button is left clicked.")

'' Function: Load_Sample_Annot_Tidy_Column_Name_Click
'' --- Code
''  Public Sub Load_Sample_Annot_Tidy_Column_Name_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Sample_Annot_Load_Sample_Annot_Tidy_Column_Name_Button.png)
''
'' The following pop up box will appear. Asking the users
'' which column to clear.
''
'' (see Sample_Annot_Load_Sample_Annot_Tidy_Column_Name_Pop_Up.png)
''
Public Sub Load_Sample_Annot_Tidy_Column_Name_Click()
Attribute Load_Sample_Annot_Tidy_Column_Name_Click.VB_Description = "Function that controls what happens when the Load Sample Annotation from Table Data button is left clicked."
    
    ' Get the Sample_Annot worksheet from the active workbook
    ' The SampleAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Sample_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "SampleAnnotSheet") = False Then
        MsgBox ("Sheet Sample_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Sample_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "SampleAnnotSheet")
      
    Sample_Annot_Worksheet.Activate
    
    Load_Sample_Annot_Tidy.Show
    
    'If the Load Annotation button is clicked
    Select Case Load_Sample_Annot_Tidy.whatsclicked
    Case "Create_New_Sample_Annot_Tidy_Button"
 
        Sample_Annot.Create_New_Sample_Annot_Tidy _
                            TidyDataFiles:=Load_Sample_Annot_Tidy.Tidy_Data_File_Path.Text, _
                            DataFileType:=Load_Sample_Annot_Tidy.Data_File_Type_ComboBox.Text, _
                            SampleProperty:=Load_Sample_Annot_Tidy.Sample_Name_Property_ComboBox.Text, _
                            StartingRowNum:=Load_Sample_Annot_Tidy.Starting_Row_Number_TextBox.Value, _
                            StartingColumnNum:=Load_Sample_Annot_Tidy.Starting_Column_Number_TextBox.Value
    
    End Select
    
    Unload Load_Sample_Annot_Tidy
    
End Sub

'@Description("Function that controls what happens when the Load Sample Annotation from Raw Data button is left clicked.")

'' Function: Load_Sample_Annot_Raw_Column_Name_Click
'' --- Code
''  Public Sub Load_Sample_Annot_Raw_Column_Name_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Sample_Annot_Load_Sample_Annot_Raw_Column_Name_Button.png)
''
'' The following pop up box will appear. Asking the users
'' which column to clear.
''
'' (see Sample_Annot_Load_Sample_Annot_Raw_Column_Name_Pop_Up.png)
''
Public Sub Load_Sample_Annot_Raw_Column_Name_Click()
Attribute Load_Sample_Annot_Raw_Column_Name_Click.VB_Description = "Function that controls what happens when the Load Sample Annotation from Raw Data button is left clicked."
    'Assume first row are the headers
    'Assume headers are fully filled, not empty
    'Assume no duplicate headers
    
    ' Get the Sample_Annot worksheet from the active workbook
    ' The SampleAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Sample_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "SampleAnnotSheet") = False Then
        MsgBox ("Sheet Sample_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Sample_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "SampleAnnotSheet")
      
    Sample_Annot_Worksheet.Activate
       
    Load_Sample_Annot_Raw.Show
    
    'If the Load Annotation button is clicked
    Select Case Load_Sample_Annot_Raw.whatsclicked
    Case "Merge_With_Sample_Annot_Button"
        Sample_Annot.Merge_With_Sample_Annot RawDataFiles:=Load_Sample_Annot_Raw.Raw_Data_File_Path.Text, _
                                             SampleAnnotFile:=Load_Sample_Annot_Raw.Sample_Annot_File_Path.Text
    Case "Create_New_Sample_Annot_Raw_Button"
        Sample_Annot.Create_New_Sample_Annot_Raw RawDataFiles:=Load_Sample_Annot_Raw.Raw_Data_File_Path.Text
    End Select
    
    Unload Load_Sample_Annot_Raw
    
End Sub
