Attribute VB_Name = "Integration_Test"
Public Sub Run_Integration_Test()
    Nothing_To_Transfer_Test
    Transition_Name_and_ISTD_Annot_Integration_Test
    Sample_Annot_Integration_Test
    Sample_Annot_and_Dilution_Annot_Integration_Test
End Sub

Public Sub Nothing_To_Transfer_Test()
    On Error GoTo TestFail
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    'Test that the button works when there are no
    'Transition_Name_ISTD to validate and transfer
    'from Transition_Name_Annot to ISTD_Annot
    Sheets("Transition_Name_Annot").Activate
    
    'When both columns are empty
    Transition_Name_Annot_Buttons.Load_Transition_Name_ISTD_Click
    Transition_Name_Annot_Buttons.Validate_ISTD_Click
    
    'When only Transition_Name_ISTD column is empty
    Dim Transition_Array(2) As String
    Transition_Array(0) = "Transition_Array1"
    Transition_Array(1) = "Transition_Array2"
    Call Utilities.Load_To_Excel(Transition_Array, "Transition_Name", HeaderRowNumber:=1, _
                                 DataStartRowNumber:=2, MessageBoxRequired:=False)
    Transition_Name_Annot_Buttons.Validate_ISTD_Click
    Call Utilities.Clear_Columns("Transition_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    
    'When only Transition_Name column is empty
    Transition_Array(0) = "Transition_ISTD1"
    Transition_Array(1) = "Transition_ISTD2"
    Call Utilities.Load_To_Excel(Transition_Array, "Transition_Name_ISTD", HeaderRowNumber:=1, _
                                 DataStartRowNumber:=2, MessageBoxRequired:=False)
    Transition_Name_Annot_Buttons.Validate_ISTD_Click
    Call Utilities.Clear_Columns("Transition_Name_ISTD", HeaderRowNumber:=1, DataStartRowNumber:=2)
    
    MsgBox "Nothing to transfer in Transition_Name_Annot test complete"
    
    'Test that the button works when there are no
    'samples with sample type RQC to transfer from Sample_Annot
    'to Dilution_Annot
    Sheets("Sample_Annot").Activate
    Sample_Annot_Buttons.Load_Sample_Name_To_Dilution_Annot_Click
    Call Utilities.Clear_Columns("Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Merge_Status", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Type", HeaderRowNumber:=1, DataStartRowNumber:=2)
    MsgBox "Nothing to transfer in Sample_Annot test complete"

    'Excel resume monitoring the sheet
    Application.EnableEvents = True
    
TestFail:
    Application.EnableEvents = True
    Exit Sub
End Sub

Public Sub Transition_Name_and_ISTD_Annot_Integration_Test()
    On Error GoTo TestFail
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    Sheets("Transition_Name_Annot").Activate
    
    Dim TestFolder As String
    Dim xFileNames As Variant
    Dim FileThere As Boolean
    Dim TidyDataRowFiles As String
    Dim TidyDataColumnFiles As String
    Dim InvalidRawDataFiles As String
    Dim AgilentRawDataFiles As String
    Dim JoinedFiles As String
    Dim Transition_Array() As String
    Dim Transition_Name_ISTD_ColLetter As String
    Dim Transition_Name_ISTD_ColNumber As Integer
    Dim ISTD_Array() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    TidyDataRowFiles = TestFolder & "TidyTransitionRow.csv"
    TidyDataColumnFiles = TestFolder & "TidyTransitionColumn.csv"
    InvalidRawDataFiles = TestFolder & "InvalidDataTest1.csv"
    AgilentRawDataFiles = TestFolder & "AgilentRawDataTest1.csv"
    
    'Check if the data file exists
    JoinedFiles = Join(Array(TidyDataRowFiles, TidyDataColumnFiles, InvalidRawDataFiles, AgilentRawDataFiles), ";")
    xFileNames = Split(JoinedFiles, ";")
    
    'xFileNames = Array(TestFolder & "AgilentRawDataTest1.csv")

    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > "")
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Test creating a new transition annotation from tidy data file with transitons as column variables
    Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array_Tidy(TidyDataFiles:=TidyDataColumnFiles, _
                                                                              DataFileType:="csv", _
                                                                              TransitionProperty:="Read as column variables", _
                                                                              StartingRowNum:=1, _
                                                                              StartingColumnNum:=2)
                                                                              
    Call Utilities.OverwriteHeader("Transition_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Load_To_Excel(Transition_Array, "Transition_Name", HeaderRowNumber:=1, _
                                 DataStartRowNumber:=2, MessageBoxRequired:=True)
                                                   
    MsgBox "Create new transition annotation from tidy column data test complete"
    
    'Test creating a new transition annotation from tidy data file with transitons as row observations
    Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array_Tidy(TidyDataFiles:=TidyDataRowFiles, _
                                                                              DataFileType:="csv", _
                                                                              TransitionProperty:="Read as row observations", _
                                                                              StartingRowNum:=2, _
                                                                              StartingColumnNum:=1)
                                                   
    Call Utilities.OverwriteHeader("Transition_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Load_To_Excel(Transition_Array, "Transition_Name", HeaderRowNumber:=1, _
                                 DataStartRowNumber:=2, MessageBoxRequired:=True)
                                 
    MsgBox "Create new transition annotation from tidy row data test complete"
    
    Call Utilities.Clear_Columns("Transition_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    
    'Testing with an invalid file
    Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array_Raw(RawDataFiles:=InvalidRawDataFiles)
    MsgBox "Invalid file input test complete"
    
    'Load the transition names and load it to excel
    Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array_Raw(RawDataFiles:=AgilentRawDataFiles)
    Call Utilities.OverwriteHeader("Transition_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Load_To_Excel(Transition_Array, "Transition_Name", HeaderRowNumber:=1, _
                                 DataStartRowNumber:=2, MessageBoxRequired:=True)
    
    'Excel resume monitoring the sheet
    Application.EnableEvents = True
    
    'Fill in the ISTD automatically
    Transition_Name_ISTD_ColNumber = Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 1)
    Transition_Name_ISTD_ColLetter = Utilities.ConvertToLetter(Transition_Name_ISTD_ColNumber)
    Range(Transition_Name_ISTD_ColLetter & 2 & ":" & Transition_Name_ISTD_ColLetter & 23) = "LPC 17:0"
    Range(Transition_Name_ISTD_ColLetter & 24 & ":" & Transition_Name_ISTD_ColLetter & 31) = "MHC d18:1/16:0d3 (IS)"
    
    'Validate ISTD, ensure that wrong ISTD is detected
    Call Validate_ISTD_Click(Testing:=True)
    
    'Correct the wrong ISTD, validate ISTD and transfer them to ISTD_Annot sheet
    Range(Transition_Name_ISTD_ColLetter & 2 & ":" & Transition_Name_ISTD_ColLetter & 23) = "LPC 17:0 (IS)"
    Call Validate_ISTD_Click(Testing:=True)
    
    Call Load_Transition_Name_ISTD_Click
    
    'Clear the columns as the test for this worksheet is complete
    Sheets("Transition_Name_Annot").Activate
    Call Utilities.Clear_Columns("Transition_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Transition_Name_ISTD", HeaderRowNumber:=1, DataStartRowNumber:=2)
    MsgBox "Transition_Name_Annot test complete"
    
    'Proceed with the ISTD_Annot test
    Sheets("ISTD_Annot").Activate
    
    Dim ISTD_Conc_ng_ColNumber As Integer
    Dim ISTD_MW_ColNumber As Integer
    Dim ISTD_Conc_ng_ColLetter As String
    Dim ISTD_MW_ColLetter As String
    ISTD_Conc_ng_ColNumber = Utilities.Get_Header_Col_Position("ISTD_Conc_[ng/mL]", 3)
    ISTD_MW_ColNumber = Utilities.Get_Header_Col_Position("ISTD_[MW]", 3)
    ISTD_Conc_ng_ColLetter = Utilities.ConvertToLetter(ISTD_Conc_ng_ColNumber)
    ISTD_MW_ColLetter = Utilities.ConvertToLetter(ISTD_MW_ColNumber)
    
    'Fill in the concentration values automatically
    Range(ISTD_Conc_ng_ColLetter & 4 & ":" & ISTD_MW_ColLetter & 5) = [{100,2,30,10}]
    
    'Perform the calculation
    ISTD_Annot_Buttons.Convert_To_Nanomolar_Click
    MsgBox "Calculation is complete"
    
    'Clear the columns as the test for this worksheet is complete
    Call Utilities.Clear_Columns("Transition_Name_ISTD", HeaderRowNumber:=2, DataStartRowNumber:=4)
    Call Utilities.Clear_Columns("ISTD_Conc_[ng/mL]", HeaderRowNumber:=3, DataStartRowNumber:=4)
    Call Utilities.Clear_Columns("ISTD_[MW]", HeaderRowNumber:=3, DataStartRowNumber:=4)
    Call Utilities.Clear_Columns("ISTD_Conc_[nM]", HeaderRowNumber:=3, DataStartRowNumber:=4)
    Call Utilities.Clear_Columns("Custom_Unit", HeaderRowNumber:=2, DataStartRowNumber:=4)
    MsgBox "ISTD_Annot test complete"
    
TestFail:
    Application.EnableEvents = True
    Exit Sub
End Sub

Public Sub Sample_Annot_Integration_Test()
    On Error GoTo TestFail
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = True
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    Sheets("Sample_Annot").Activate
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim TidyDataRowFiles As String
    Dim TidyDataColumnFiles As String
    Dim JoinedFiles As String
    Dim SampleAnnotFile As String
    Dim xFileNames() As String
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "AgilentRawDataTest1.csv"
    SampleAnnotFile = TestFolder & "Sample_Annotation_Example.csv"
    TidyDataRowFiles = TestFolder & "TidySampleRow.csv"
    TidyDataColumnFiles = TestFolder & "TidySampleColumn.csv"
   
    JoinedFiles = Join(Array(RawDataFiles, TidyDataRowFiles, TidyDataColumnFiles), ";")
    xFileNames = Split(JoinedFiles, ";")
    
    'Check if the data file exists
    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > "")
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Check if the sample annotation file exists
    FileThere = (Dir(SampleAnnotFile) > "")
    If FileThere = False Then
        MsgBox "File name " & SampleAnnotFile & " cannot be found."
        End
    End If
    
    'Test creating a new sample annotation from raw data file
    Call Sample_Annot.Create_New_Sample_Annot_Raw(RawDataFiles:=RawDataFiles)
    Call Autofill_Sample_Type_Click
    
    'Call Autofill_Concentration_Unit_Click
    MsgBox "Create new sample annotation test from raw data complete"
    
    'Clear the sample annotation
    Call Utilities.Clear_Columns("Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Merge_Status", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Type", HeaderRowNumber:=1, DataStartRowNumber:=2)
    
    'Fill in the Sample_Amount_Unit
    Dim Sample_Amount_Unit_ColNumber As Integer
    Dim Sample_Amount_Unit_ColLetter As String
    Sample_Amount_Unit_ColNumber = Utilities.Get_Header_Col_Position("Sample_Amount_Unit", 1)
    Sample_Amount_Unit_ColLetter = Utilities.ConvertToLetter(Sample_Amount_Unit_ColNumber)
    
    'Fill in the concentration values automatically
    Range(Sample_Amount_Unit_ColLetter & 2 & ":" & _
          Sample_Amount_Unit_ColLetter & 7) = Application.Transpose(Array("cell_number", "mg_dry_weight", _
                                                                          "mg_fresh_weight", "ng_DNA", _
                                                                          "ug_total_protein", "uL"))
    Call Autofill_Concentration_Unit_Click(Testing:=True)
    MsgBox "Autofill concentration unit test complete"
    
    'Clear the sample annotation
    Call Utilities.Clear_Columns("Sample_Amount_Unit", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Concentration_Unit", HeaderRowNumber:=1, DataStartRowNumber:=2)
    
    'Prepare for merging
    Load_Sample_Annot_Raw.Sample_Name_Text.Text = "Sample"
    Load_Sample_Annot_Raw.Sample_Amount_Text.Text = "Cell Number"
    Load_Sample_Annot_Raw.ISTD_Mixture_Volume_Text.Text = "ISTD Volume"
    
    'Test merging with existing sample annotation with raw data
    Sample_Annot.Merge_With_Sample_Annot RawDataFiles:=RawDataFiles, _
                                         SampleAnnotFile:=SampleAnnotFile
    Sample_Annot_Buttons.Autofill_Sample_Type_Click
    MsgBox "Merging raw data with sample annotation test complete"
    
    'Clear the sample annotation
    Call Utilities.Clear_Columns("Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Merge_Status", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Type", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Amount", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Amount_Unit", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("ISTD_Mixture_Volume_[ul]", HeaderRowNumber:=1, DataStartRowNumber:=2)
    
    'Test creating a new sample annotation from tidy data file
    Call Sample_Annot.Create_New_Sample_Annot_Tidy(TidyDataFiles:=TidyDataColumnFiles, _
                                                   DataFileType:="csv", _
                                                   SampleProperty:="Read as column variables", _
                                                   StartingRowNum:=1, _
                                                   StartingColumnNum:=2)
                                                   
    MsgBox "Create new sample annotation from tidy column data test complete"
    
    Call Sample_Annot.Create_New_Sample_Annot_Tidy(TidyDataFiles:=TidyDataRowFiles, _
                                                   DataFileType:="csv", _
                                                   SampleProperty:="Read as row observations", _
                                                   StartingRowNum:=2, _
                                                   StartingColumnNum:=1)
                                                   
    MsgBox "Create new sample annotation from tidy row data test complete"
    Call Utilities.Clear_Columns("Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Merge_Status", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Type", HeaderRowNumber:=1, DataStartRowNumber:=2)
    
    MsgBox "Sample_Annot test complete"
       
TestFail:
    Application.EnableEvents = True
    Exit Sub
End Sub

Public Sub Sample_Annot_and_Dilution_Annot_Integration_Test()
    On Error GoTo TestFail
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = True
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    Sheets("Sample_Annot").Activate
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim xFileNames() As String
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "DilutionTest.csv"
    
    xFileNames = Split(RawDataFiles, ";")
    
    'Check if the data file exists
    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > "")
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Test creating a new sample annotation
    Call Sample_Annot.Create_New_Sample_Annot_Raw(RawDataFiles:=RawDataFiles)
    MsgBox "Load RQC samples to copy"
    
    'Load to Dilution_Annot
    Call Load_Sample_Name_To_Dilution_Annot_Click
    MsgBox "Copy RQC Samples to Dilution Annot complete"
    
    'Clear the dilution annotation
    Sheets("Dilution_Annot").Activate
    Call Utilities.Clear_Columns("Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Dilution_Batch_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Relative_Sample_Amount_[%]", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Injection_Volume_[uL]", HeaderRowNumber:=1, DataStartRowNumber:=2)

    'Clear the sample annotation
    Sheets("Sample_Annot").Activate
    Call Utilities.Clear_Columns("Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Merge_Status", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Type", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Amount", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Amount_Unit", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("ISTD_Mixture_Volume_[uL]", HeaderRowNumber:=1, DataStartRowNumber:=2)
    MsgBox "Dilution_Annot test complete"

    
TestFail:
    Application.EnableEvents = True
    Exit Sub
End Sub
