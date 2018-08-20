Attribute VB_Name = "Buttons"
'Sheet Sample_Annot Functions
Sub Clear_Sample_Table_Click()
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    Clear_Sample_Annot.Show
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

Sub Load_Sample_Annot_Column_Name_Click()
    'Assume first row are the headers
    'Assume headers are fully filled, not empty
    'Assume no duplicate headers
    
    Sheets("Sample_Annot").Activate
       
    Load_Sample_Annot.Show
    
    'If the Load Annotation button is clicked
    Select Case Load_Sample_Annot.whatsclicked
        Case "Merge_With_Sample_Annot_Button"
            Call Sample_Annot.Merge_With_Sample_Annot(RawDataFiles:=Load_Sample_Annot.Raw_Data_File_Path.Text, SampleAnnotFile:=Load_Sample_Annot.Sample_Annot_File_Path.Text)
        Case "Create_New_Sample_Annot_Button"
            Call Sample_Annot.Create_new_Sample_Annot(RawDataFiles:=Load_Sample_Annot.Raw_Data_File_Path.Text)
    End Select
    
    Unload Load_Sample_Annot
    
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
    Dim ISTD_Conc_nM() As String
    ISTD_Conc_nM = ISTD_Annot.Get_ISTD_Conc_nM_Array()
    Call Utilities.Load_To_Excel(ISTD_Conc_nM, "ISTD_Conc_[nM]", HeaderRowNumber:=3, DataStartRowNumber:=4, MessageBoxRequired:=False)
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
                                                    RemoveBlksAndReplicates:=True, IgnoreEmptyArray:=False)
                                                    
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
                                                    RemoveBlksAndReplicates:=True, IgnoreEmptyArray:=False)
    ISTD_Array = Utilities.Load_Columns_From_Excel("Transition_Name_ISTD", HeaderRowNumber:=1, _
                                                    DataStartRowNumber:=2, MessageBoxRequired:=False, _
                                                    RemoveBlksAndReplicates:=False, IgnoreEmptyArray:=False)
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
    Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array()
    
    'Leave the program if we have an empty array
    If Len(Join(Transition_Array, "")) = 0 Then
        'Don't need to display message as we did that in Transition_Name_Annot.Get_Sorted_Transition_Array
        'MsgBox "Could not find any Transition Names"
        Exit Sub
    End If
    
    'Excel resume monitoring the sheet
    Application.EnableEvents = True
    
    Call Utilities.OverwriteHeader("Transition_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Load_To_Excel(Transition_Array, "Transition_Name", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=True)
End Sub
