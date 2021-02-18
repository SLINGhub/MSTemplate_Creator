Attribute VB_Name = "Transition_Name_Annot_Buttons"
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
    Call Validate_ISTD_Click(MessageBoxRequired:=False)
      
    'Go to the ISTD_Annot sheet
    Sheets("ISTD_Annot").Activate
    Call Utilities.OverwriteHeader("Transition_Name_ISTD", HeaderRowNumber:=2, DataStartRowNumber:=4)
    Call Utilities.Load_To_Excel(ISTD_Array, "Transition_Name_ISTD", HeaderRowNumber:=2, DataStartRowNumber:=4, MessageBoxRequired:=True)
End Sub

Sub Validate_ISTD_Click(Optional ByVal MessageBoxRequired As Boolean = True, _
                        Optional ByVal Testing As Boolean = False)
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
    Dim RawDataFiles As String
    
    
    xFileNames = Application.GetOpenFilename(Title:="Load MS Raw Data", MultiSelect:=True)
    
    'When no file is selected
    If TypeName(xFileNames) = "Boolean" Then
        End
    End If
    On Error GoTo 0
    
    RawDataFiles = Join(xFileNames, ";")
    Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array_Raw(RawDataFiles:=RawDataFiles)
    
    'Excel resume monitoring the sheet
    Application.EnableEvents = True
    
    'Leave the program if we have an empty array
    If Len(Join(Transition_Array, "")) = 0 Then
        'Don't need to display message as we did that in
        'Transition_Name_Annot.Get_Sorted_Transition_Array_Raw
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




