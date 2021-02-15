Attribute VB_Name = "Buttons"
'Sheet Dilution_Annot Functions

Sub Clear_Dilution_Annot_Click()
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    Clear_Dilution_Annot.Show
End Sub

'Sheet Sample_Annot Functions

Sub Load_Sample_Name_To_Dilution_Annot_Click()
    Sheets("Sample_Annot").Activate
    
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    
    Dim SampleNameArray() As String
    Dim FileNameArray() As String
        
    'Check if the column Sample_Type exists
    Dim SampleType_pos As Integer
    SampleType_pos = Utilities.Get_Header_Col_Position("Sample_Type", HeaderRowNumber:=1)
    
    'Filter Rows by "RQC"
    ActiveSheet.Range("A1").AutoFilter Field:=SampleType_pos, _
                                       Criteria1:="RQC", _
                                       VisibleDropDown:=True
                                       
                                       

    'Load the Sample_Name columns content from Sample_Annot
    SampleNameArray = Utilities.Load_Columns_From_Excel("Sample_Name", HeaderRowNumber:=1, _
                                                        DataStartRowNumber:=2, MessageBoxRequired:=True, _
                                                        RemoveBlksAndReplicates:=False, _
                                                        IgnoreHiddenRows:=True, IgnoreEmptyArray:=True)
                                                    
    'Load the Raw_Data_File_Name columns content from Sample_Annot
    FileNameArray = Utilities.Load_Columns_From_Excel("Raw_Data_File_Name", HeaderRowNumber:=1, _
                                                       DataStartRowNumber:=2, MessageBoxRequired:=True, _
                                                       RemoveBlksAndReplicates:=False, _
                                                       IgnoreHiddenRows:=True, IgnoreEmptyArray:=True)

    'Debug.Print FileNameArray(1)
    
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    'Resume monitoring of sheet
    Application.EnableEvents = True
                                                        
    'Check if SampleNameArray has any elements
    If Len(Join(SampleNameArray, "")) = 0 Then
        End
    End If
    
    'Go to the Dilution_Annot sheet
    Sheets("Dilution_Annot").Activate
    Call Utilities.OverwriteHeader("Raw_Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Load_To_Excel(FileNameArray, "Raw_Data_File_Name", HeaderRowNumber:=1, _
                                 DataStartRowNumber:=2, MessageBoxRequired:=False)
    Call Utilities.OverwriteHeader("Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Load_To_Excel(SampleNameArray, "Sample_Name", HeaderRowNumber:=1, _
                                 DataStartRowNumber:=2, MessageBoxRequired:=True)

End Sub

Sub Clear_Sample_Table_Click()
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    Clear_Sample_Annot.Show
End Sub

Sub Autofill_Concentration_Unit_Click()

    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    Sheets("ISTD_Annot").Activate
    
    'Check if the column Custom_Unit exists
    Dim ISTD_Custom_Unit_ColNumber As Integer
    ISTD_Custom_Unit_ColNumber = Utilities.Get_Header_Col_Position("Custom_Unit", 2)
    
    'Get the custom unit value
    Dim Custom_Unit As String
    Custom_Unit = Cells(3, ISTD_Custom_Unit_ColNumber)
    Application.EnableEvents = True

    Dim Right_Custom_Unit As String
    Dim RightConcUnitRegEx As New RegExp
    'Get the right value after "or"
    RightConcUnitRegEx.Pattern = "(.*or)"
    RightConcUnitRegEx.Global = True
    Right_Custom_Unit = Trim(RightConcUnitRegEx.Replace(Custom_Unit, " "))
    'Remove square brackets and mL
    RightConcUnitRegEx.Pattern = "[\[\]]"
    Right_Custom_Unit = Trim(RightConcUnitRegEx.Replace(Right_Custom_Unit, " "))
    RightConcUnitRegEx.Pattern = "/mL"
    'Right_Custom_Unit = RightConcUnitRegEx.Execute(Custom_Unit)(0).SubMatches(0)
    Right_Custom_Unit = Trim(RightConcUnitRegEx.Replace(Right_Custom_Unit, " "))
    'Debug.Print Right_Custom_Unit

    Sheets("Sample_Annot").Activate
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    Dim Sample_Amount_Unit() As String
    Sample_Amount_Unit = Utilities.Load_Columns_From_Excel("Sample_Amount_Unit", HeaderRowNumber:=1, _
                                                           DataStartRowNumber:=2, MessageBoxRequired:=False, _
                                                           RemoveBlksAndReplicates:=False, _
                                                           IgnoreHiddenRows:=False, IgnoreEmptyArray:=True)
    'Get the length of Sample_Amount_Unit
    Dim max_length As Integer
    max_length = 0
    If Utilities.StringArrayLen(Sample_Amount_Unit) > max_length Then
            max_length = Utilities.StringArrayLen(Sample_Amount_Unit)
    End If
    
    'Leave the program if max_length is 0
    If max_length = 0 Then
        'Application.EnableEvents = True
        End
    End If
    
    Dim ConcentrationUnitArray() As String
    Dim UniqueConcentrationUnitArray() As String
    Dim UniqueArraryLength As Integer
    UniqueArraryLength = 0
    'Resize the array to max_length
    ReDim Preserve ConcentrationUnitArray(max_length)
    
    'Add the concentration unit when necessary
    For i = 0 To max_length - 1
        Dim ConcentrationUnit As String
        If Len(Sample_Amount_Unit(i)) <> 0 Then
            ConcentrationUnit = Right_Custom_Unit & "/" & Sample_Amount_Unit(i)
            ConcentrationUnitArray(i) = ConcentrationUnit
            
            'Collect Unique concentration unit
            InArray = Utilities.IsInArray(ConcentrationUnit, UniqueConcentrationUnitArray)
            If Not InArray Then
                ReDim Preserve UniqueConcentrationUnitArray(UniqueArraryLength)
                UniqueConcentrationUnitArray(UniqueArraryLength) = ConcentrationUnit
                'Debug.Print UniqueConcentrationUnitArray(UniqueArraryLength)
                UniqueArraryLength = UniqueArraryLength + 1
            End If
            
        End If
    Next
    
    'Load to Excel
    Call Utilities.OverwriteHeader("Concentration_Unit", HeaderRowNumber:=1, _
                                   DataStartRowNumber:=2)
    Call Utilities.Load_To_Excel(ConcentrationUnitArray, "Concentration_Unit", _
                                 HeaderRowNumber:=1, DataStartRowNumber:=2, _
                                 MessageBoxRequired:=False)
                                 
    'Display a summary box of unique concentration units
    If Utilities.StringArrayLen(UniqueConcentrationUnitArray) <> 0 Then
        'Put the concentration units in the list box to be displayed
        For i = 0 To UBound(UniqueConcentrationUnitArray) - LBound(UniqueConcentrationUnitArray)
            Concentration_Unit_MsgBox.Concentration_Unit_ListBox.AddItem UniqueConcentrationUnitArray(i)
        Next i
        Concentration_Unit_MsgBox.Show
        If Testing Then
            Exit Sub
        Else
            End
        End If
    End If
    
End Sub

Sub Autofill_Sample_Type_Click()
    Sheets("Sample_Annot").Activate
    
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    Dim SampleArray() As String
    Dim TotalRows As Long
    Dim i As Long
    
    'Check if the column Sample_Name exists
    Dim SampleName_pos As Integer
    SampleName_pos = Utilities.Get_Header_Col_Position("Sample_Name", HeaderRowNumber:=1)
    
    'Check if the column Sample_Type exists
    Dim SampleType_pos As Integer
    SampleType_pos = Utilities.Get_Header_Col_Position("Sample_Type", HeaderRowNumber:=1)
   
    'Find the total number of rows and resize the array accordingly
    TotalRows = Cells(Rows.Count, ConvertToLetter(SampleName_pos)).End(xlUp).Row
    ReDim SampleArray(0 To TotalRows - 1)
    
    'Assign "Sample" if there is no sample type
    If TotalRows > 1 Then
        For i = 2 To TotalRows
            If Cells(i, SampleType_pos).Value = "" Then
                SampleArray(i - 2) = "SPL"
            Else
                SampleArray(i - 2) = Cells(i, SampleType_pos).Value
            End If
            'Debug.Print SampleArray(i - 2)
        Next i
    End If
    
    Call Utilities.Load_To_Excel(SampleArray, "Sample_Type", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
    'Range(ConvertToLetter(SampleType_pos) & "2").Resize(UBound(SampleArray) + 1) = Application.Transpose(SampleArray)

End Sub

Sub Load_Sample_Annot_Tidy_Column_Name_Click()
    Sheets("Sample_Annot").Activate
    
    Load_Sample_Annot_Tidy.Show
    
    'If the Load Annotation button is clicked
    Select Case Load_Sample_Annot_Tidy.whatsclicked
    Case "Create_New_Sample_Annot_Tidy_Button"
        'Debug.Print Load_Sample_Annot_Tidy.Tidy_Data_File_Path.Text
        'Debug.Print Load_Sample_Annot_Tidy.Data_File_Type_ComboBox.Text
        'Debug.Print Load_Sample_Annot_Tidy.Sample_Name_Property_ComboBox.Text
        'Debug.Print Load_Sample_Annot_Tidy.Starting_Row_Number_TextBox.Value
        'Debug.Print Load_Sample_Annot_Tidy.Starting_Column_Number_TextBox.Value
        
        Call Sample_Annot.Create_New_Sample_Annot_Tidy( _
                            TidyDataFiles:=Load_Sample_Annot_Tidy.Tidy_Data_File_Path.Text, _
                            DataFileType:=Load_Sample_Annot_Tidy.Data_File_Type_ComboBox.Text, _
                            SampleProperty:=Load_Sample_Annot_Tidy.Sample_Name_Property_ComboBox.Text, _
                            StartingRowNum:=Load_Sample_Annot_Tidy.Starting_Row_Number_TextBox.Value, _
                            StartingColumnNum:=Load_Sample_Annot_Tidy.Starting_Column_Number_TextBox.Value)
    
    End Select
    
    Unload Load_Sample_Annot_Tidy
    
End Sub

Sub Load_Sample_Annot_Raw_Column_Name_Click()
    'Assume first row are the headers
    'Assume headers are fully filled, not empty
    'Assume no duplicate headers
    
    Sheets("Sample_Annot").Activate
       
    Load_Sample_Annot_Raw.Show
    
    'If the Load Annotation button is clicked
    Select Case Load_Sample_Annot_Raw.whatsclicked
    Case "Merge_With_Sample_Annot_Button"
        Call Sample_Annot.Merge_With_Sample_Annot(RawDataFiles:=Load_Sample_Annot_Raw.Raw_Data_File_Path.Text, _
                                                  SampleAnnotFile:=Load_Sample_Annot_Raw.Sample_Annot_File_Path.Text)
    Case "Create_New_Sample_Annot_Raw_Button"
        Call Sample_Annot.Create_New_Sample_Annot_Raw(RawDataFiles:=Load_Sample_Annot_Raw.Raw_Data_File_Path.Text)
    End Select
    
    Unload Load_Sample_Annot_Raw
    
End Sub

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
    ISTD_Conc_nM = ISTD_Annot.Get_ISTD_Conc_nM_Array()
    Call Utilities.Load_To_Excel(ISTD_Conc_nM, "ISTD_Conc_[nM]", HeaderRowNumber:=3, DataStartRowNumber:=4, MessageBoxRequired:=False)
    ISTD_Custom_Unit = ISTD_Annot.Convert_Conc_nM_Array(Custom_Unit)
    Call Utilities.Load_To_Excel(ISTD_Custom_Unit, "Custom_Unit", HeaderRowNumber:=2, DataStartRowNumber:=4, MessageBoxRequired:=False)
    
    'Resume monitoring of sheet
    Application.EnableEvents = True
End Sub

'Sheet Transition_Name_Annot Function

Sub Clear_Transition_Name_Annot_Click()
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    Clear_Transition_Name_Annot.Show
End Sub

Sub Load_Transition_Name_ISTD_Click()
    Sheets("Transition_Name_Annot").Activate

    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    Dim ISTD_Array() As String
    ISTD_Array = Utilities.Load_Columns_From_Excel("Transition_Name_ISTD", HeaderRowNumber:=1, _
                                                   DataStartRowNumber:=2, MessageBoxRequired:=True, _
                                                   RemoveBlksAndReplicates:=True, _
                                                   IgnoreHiddenRows:=False, IgnoreEmptyArray:=True)
                                                    
    'Excel resume monitoring the sheet
    Application.EnableEvents = True
    
    'Validate the ISTD column
    Call Buttons.Validate_ISTD_Click(MessageBoxRequired:=False)
      
    'Go to the ISTD_Annot sheet
    Sheets("ISTD_Annot").Activate
    Call Utilities.OverwriteHeader("Transition_Name_ISTD", HeaderRowNumber:=2, DataStartRowNumber:=4)
    Call Utilities.Load_To_Excel(ISTD_Array, "Transition_Name_ISTD", HeaderRowNumber:=2, DataStartRowNumber:=4, MessageBoxRequired:=True)
End Sub

Sub Validate_ISTD_Click(Optional ByVal MessageBoxRequired As Boolean = True, Optional ByVal Testing As Boolean = False)
    Sheets("Transition_Name_Annot").Activate
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    Dim Transition_Array() As String
    Dim ISTD_Array() As String
    Transition_Array = Utilities.Load_Columns_From_Excel("Transition_Name", HeaderRowNumber:=1, _
                                                         DataStartRowNumber:=2, MessageBoxRequired:=False, _
                                                         RemoveBlksAndReplicates:=True, _
                                                         IgnoreHiddenRows:=False, IgnoreEmptyArray:=True)
    ISTD_Array = Utilities.Load_Columns_From_Excel("Transition_Name_ISTD", HeaderRowNumber:=1, _
                                                   DataStartRowNumber:=2, MessageBoxRequired:=False, _
                                                   RemoveBlksAndReplicates:=True, _
                                                   IgnoreHiddenRows:=False, IgnoreEmptyArray:=True)
    'Excel resume monitoring the sheet
    Application.EnableEvents = True
    
    'Both arrays should not be empty
    Call Transition_Name_Annot.VerifyISTD(Transition_Array, ISTD_Array, MessageBoxRequired:=MessageBoxRequired, Testing:=Testing)
    
End Sub

Sub GetTransitionArray_Click()
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    Sheets("Transition_Name_Annot").Activate
    Dim Transition_Array() As String
    Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array_Raw()
    
    'Excel resume monitoring the sheet
    Application.EnableEvents = True
    
    'Leave the program if we have an empty array
    If Len(Join(Transition_Array, "")) = 0 Then
        'Don't need to display message as we did that in
        'Transition_Name_Annot.Get_Sorted_Transition_Array_Raw
        'MsgBox "Could not find any Transition Names"
        Exit Sub
    End If
    
    Call Utilities.OverwriteHeader("Transition_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Load_To_Excel(Transition_Array, "Transition_Name", HeaderRowNumber:=1, _
                                 DataStartRowNumber:=2, MessageBoxRequired:=True)
End Sub

Sub GetTransitionArrayTidy_Click()
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    Sheets("Transition_Name_Annot").Activate
    Dim Transition_Array() As String
    Load_Transition_Name_Tidy.Show
     
    'If the Load Annotation button is clicked
    Select Case Load_Transition_Name_Tidy.whatsclicked
    Case "Create_New_Transition_Annot_Button"
        Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array_Tidy( _
                           TidyDataFiles:=Load_Transition_Name_Tidy.Tidy_Data_File_Path.Text, _
                           DataFileType:=Load_Transition_Name_Tidy.Data_File_Type_ComboBox.Text, _
                           TransitionProperty:=Load_Transition_Name_Tidy.Transition_Name_Property_ComboBox.Text, _
                           StartingRowNum:=Load_Transition_Name_Tidy.Starting_Row_Number_TextBox.Value, _
                           StartingColumnNum:=Load_Transition_Name_Tidy.Starting_Column_Number_TextBox.Value)
    End Select
    
    Unload Load_Transition_Name_Tidy
    
    'Excel resume monitoring the sheet
    Application.EnableEvents = True
    
    'Leave the program if we have an empty array
    If Len(Join(Transition_Array, "")) = 0 Then
        'Don't need to display message as we did that in
        'Transition_Name_Annot.Get_Sorted_Transition_Array_Tidy
        Exit Sub
    End If
    
    Call Utilities.OverwriteHeader("Transition_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Load_To_Excel(Transition_Array, "Transition_Name", HeaderRowNumber:=1, _
                                 DataStartRowNumber:=2, MessageBoxRequired:=True)
    
End Sub


