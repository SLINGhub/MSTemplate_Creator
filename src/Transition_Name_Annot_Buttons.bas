Attribute VB_Name = "Transition_Name_Annot_Buttons"
Option Explicit
'@Folder("Transition_Name_Annot Functions")

'' Function: Clear_Transition_Name_Annot_Click
'' --- Code
''  Public Sub Clear_Transition_Name_Annot_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Transition_Annot_Clear_Columns_Button.png)
''
'' The following pop up box will appear. Asking the users
'' which column to clear.
''
'' (see Transition_Annot_Clear_Data_Pop_Up.png)
''
Public Sub Clear_Transition_Name_Annot_Click()
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    Clear_Transition_Name_Annot.Show
End Sub

'' Function: Load_Transition_Name_ISTD_Click
'' --- Code
''  Public Sub Load_Transition_Name_ISTD_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Transition_Annot_Load_ISTD_To_ISTD_Table_Button.png)
''
'' The program will first verify if all the ISTD are valid.
''
'' If not all are valid, the following message box will appear and the program will stop
''
'' (see Transition_Name_Annot_Verify_ISTD_Invalid_ISTD_Message.png)
''
'' If there are no entries in the Transition_Name_ISTD column, the following meesage will appear
''
'' (see Transition_Annot_Load_Zero_ISTD.png)
''
'' If all entries in the Transition_Name_ISTD column are valid, the program will transfer the entries
'' to the Transition_Name_ISTD column in ISTD_Annot sheet with its duplicates removed.
''
'' (see Transition_Annot_Load_Two_ISTD.png)
''
Public Sub Load_Transition_Name_ISTD_Click()

    ' Get the Transition_Name_Annot worksheet from the active workbook
    ' The TransitionNameAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Transition_Name_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "TransitionNameAnnotSheet") = False Then
        MsgBox ("Sheet Transition_Name_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Transition_Name_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "TransitionNameAnnotSheet")
      
    Transition_Name_Annot_Worksheet.Activate
    
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    Dim ISTD_Array() As String
    ISTD_Array = Utilities.Load_Columns_From_Excel(HeaderName:="Transition_Name_ISTD", HeaderRowNumber:=1, _
                                                   DataStartRowNumber:=2, MessageBoxRequired:=True, _
                                                   RemoveBlksAndReplicates:=True, _
                                                   IgnoreHiddenRows:=False, IgnoreEmptyArray:=True)
                                                    
    'Excel resume monitoring the sheet
    Application.EnableEvents = True
    
    'If we have an empty array, leave the sub
    If Len(Join(ISTD_Array, vbNullString)) = 0 Then
        Exit Sub
    End If
    
    'Validate the ISTD column
    Transition_Name_Annot_Buttons.Validate_ISTD_Click MessageBoxRequired:=False
      
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
  
    Utilities.OverwriteHeader HeaderName:="Transition_Name_ISTD", _
                              HeaderRowNumber:=2, _
                              DataStartRowNumber:=4
                              
    Utilities.Load_To_Excel Data_Array:=ISTD_Array, _
                            HeaderName:="Transition_Name_ISTD", _
                            HeaderRowNumber:=2, _
                            DataStartRowNumber:=4, _
                            MessageBoxRequired:=True
End Sub

'' Function: Validate_ISTD_Click
'' --- Code
''  Public Sub Validate_ISTD_Click(Optional ByVal MessageBoxRequired As Boolean = True, _
''                                 Optional ByVal Testing As Boolean = False)
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Transition_Annot_Validate_ISTD_Button.png)
''
'' The program will first load entries from the Transition_Annot and the
'' Transition_Annot_ISTD column.
''
'' It will first check if these columns are empty or not.
''
'' The following message box will appear if both columns are empty.
''
'' (see Transition_Annot_Validate_No_Transition_And_ISTD_Message.png)
''
'' The following message box will appear if only the Transition_Annot column is empty.
''
'' (see Transition_Annot_Validate_No_Transition_Message.png)
''
'' The following message box will appear if only the Transition_Annot_ISTD column is empty.
''
'' (see Transition_Annot_Validate_No_ISTD_Message.png)
''
'' Next, it uses the function Transition_Name_Annot.Verify_ISTD to verify
'' if the ISTD is valid.
''
'' Input is valid if the entires can also be found in the column Transition_Annot
''
'' See this function documentation for more information
''
'' Parameters:
''
''    MessageBoxRequired As Boolean - When set to True, the following pop up boxes will appear
''
'' (see Transition_Annot_Validate_No_Transition_And_ISTD_Message.png)
''
'' (see Transition_Annot_Validate_No_Transition_Message.png)
''
'' (see Transition_Annot_Validate_No_ISTD_Message.png)
''
'' (see Transition_Name_Annot_Verify_ISTD_All_Valid_ISTD_Message.png)
''
''    Testing As Boolean - When set to False, after the function is used, it will exit the program as this
''                         function is meant to be run alone or as a last/final step.
''                         When set to True, after the function is used, it will exit the function and
''                         other functions can be called
''
''
Public Sub Validate_ISTD_Click(Optional ByVal MessageBoxRequired As Boolean = True, _
                               Optional ByVal Testing As Boolean = False)
                        
    ' Get the Transition_Name_Annot worksheet from the active workbook
    ' The TransitionNameAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Transition_Name_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "TransitionNameAnnotSheet") = False Then
        MsgBox ("Sheet Transition_Name_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Transition_Name_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "TransitionNameAnnotSheet")
      
    Transition_Name_Annot_Worksheet.Activate

    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    Dim Transition_Array() As String
    Dim ISTD_Array() As String
    Transition_Array = Utilities.Load_Columns_From_Excel(HeaderName:="Transition_Name", HeaderRowNumber:=1, _
                                                         DataStartRowNumber:=2, MessageBoxRequired:=False, _
                                                         RemoveBlksAndReplicates:=True, _
                                                         IgnoreHiddenRows:=False, IgnoreEmptyArray:=True)
    ISTD_Array = Utilities.Load_Columns_From_Excel(HeaderName:="Transition_Name_ISTD", HeaderRowNumber:=1, _
                                                   DataStartRowNumber:=2, MessageBoxRequired:=False, _
                                                   RemoveBlksAndReplicates:=True, _
                                                   IgnoreHiddenRows:=False, IgnoreEmptyArray:=True)
    'Excel resume monitoring the sheet
    Application.EnableEvents = True
    
    'If we have an empty array, leave the sub
    If Len(Join(Transition_Array, vbNullString)) = 0 And Len(Join(ISTD_Array, vbNullString)) = 0 Then
        If MessageBoxRequired Then
            MsgBox "No entries in both Transition_Name and Transition_Name_ISTD to validate."
        End If
        Exit Sub
    ElseIf Len(Join(Transition_Array, vbNullString)) <> 0 And Len(Join(ISTD_Array, vbNullString)) = 0 Then
        If MessageBoxRequired Then
            MsgBox "No entries in Transition_Name_ISTD to validate."
        End If
        Exit Sub
    ElseIf Len(Join(Transition_Array, vbNullString)) = 0 And Len(Join(ISTD_Array, vbNullString)) <> 0 Then
        If MessageBoxRequired Then
            MsgBox "No entries in Transition_Name to validate."
        End If
        Exit Sub
    End If
    
    'Both arrays should not be empty
    Transition_Name_Annot.Verify_ISTD Transition_Array:=Transition_Array, _
                                      ISTD_Array:=ISTD_Array, _
                                      MessageBoxRequired:=MessageBoxRequired, _
                                      Testing:=Testing
    
End Sub

'' Function: Get_Transition_Array_Click
'' --- Code
''  Public Sub Get_Transition_Array_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Transition_Annot_Load_Transition_Name_Raw_Button.png)
''
'' The following pop up box will appear. Asking the users
'' which raw file to load.
''
'' (see Transition_Annot_Load_Transition_Name_Raw_Choose_Files.png)
''
'' Click on "Open" and the transition names from the raw data will automatically be loaded
''
'' (see Transition_Annot_Load_Transition_Name_Raw_Loaded.png)
''
Public Sub Get_Transition_Array_Click()
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    
    ' Get the Transition_Name_Annot worksheet from the active workbook
    ' The TransitionNameAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Transition_Name_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "TransitionNameAnnotSheet") = False Then
        MsgBox ("Sheet Transition_Name_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Transition_Name_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "TransitionNameAnnotSheet")
      
    Transition_Name_Annot_Worksheet.Activate
    
    Dim Transition_Array() As String
    Dim RawDataFiles As String
    Dim xFileNames As Variant
    
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
    If Len(Join(Transition_Array, vbNullString)) = 0 Then
        'Don't need to display message as we did that in
        'Transition_Name_Annot.Get_Sorted_Transition_Array_Raw
        Exit Sub
    End If
    
    Utilities.OverwriteHeader HeaderName:="Transition_Name", _
                              HeaderRowNumber:=1, _
                              DataStartRowNumber:=2
                              
    Utilities.Load_To_Excel Data_Array:=Transition_Array, _
                            HeaderName:="Transition_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=True
End Sub

'' Function: Get_Transition_Array_Tidy_Click
'' --- Code
''  Public Sub Get_Transition_Array_Tidy_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Transition_Annot_Load_Transition_Name_Tidy_Button.png)
''
'' The following pop up box will appear. Asking the users
'' which tabular data file to load.
''
'' (see Transition_Annot_Load_Transition_Name_Tidy_Pop_Up.png)
''
'' Correct usage are summarised as follows
''
'' (see Transition_Annot_Load_Transition_Name_Tidy_Column_Loaded.png)
''
'' (see Transition_Annot_Load_Transition_Name_Tidy_Row_Loaded.png)
''
Public Sub Get_Transition_Array_Tidy_Click()
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    
    ' Get the Transition_Name_Annot worksheet from the active workbook
    ' The TransitionNameAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Transition_Name_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "TransitionNameAnnotSheet") = False Then
        MsgBox ("Sheet Transition_Name_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Transition_Name_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "TransitionNameAnnotSheet")
      
    Transition_Name_Annot_Worksheet.Activate
    
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
    If Len(Join(Transition_Array, vbNullString)) = 0 Then
        'Don't need to display message as we did that in
        'Transition_Name_Annot.Get_Sorted_Transition_Array_Tidy
        Exit Sub
    End If
    
    Utilities.OverwriteHeader HeaderName:="Transition_Name", _
                              HeaderRowNumber:=1, _
                              DataStartRowNumber:=2
                              
    Utilities.Load_To_Excel Data_Array:=Transition_Array, _
                            HeaderName:="Transition_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=True
    
End Sub

