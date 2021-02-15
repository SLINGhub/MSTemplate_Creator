Attribute VB_Name = "Integration_Test"
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
    Dim Transition_Array() As String
    Dim ISTD_Array() As String
    
    Dim Transition_Name_ISTD_ColLetter As String
    Dim Transition_Name_ISTD_ColNumber As Integer
    Transition_Name_ISTD_ColNumber = Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 1)
    Transition_Name_ISTD_ColLetter = Utilities.ConvertToLetter(Transition_Name_ISTD_ColNumber)
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    
    'Check if the data file exists
    
    xFileNames = Array(TestFolder & "AgilentRawDataTest1.csv")

    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > "")
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Testing with an invalid file
    Transition_Array = Load_Raw_Data.Get_Transition_Array(xFileNames:=Array(TestFolder & "Autophagy_Samples_List.csv"))
    MsgBox "Invalid file input test complete"
    
    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array(xFileNames:=xFileNames)
    Call Utilities.OverwriteHeader("Transition_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Load_To_Excel(Transition_Array, "Transition_Name", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=True)
    
    'Excel resume monitoring the sheet
    Application.EnableEvents = True
    
    'Fill in the ISTD automatically
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
    Call nM_calculation_Click
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
    
    xFileNames = Split(RawDataFiles, ";")
    
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
    
    'Test creating a new sample annotation
    Call Sample_Annot.Create_New_Sample_Annot_Raw(RawDataFiles:=RawDataFiles)
    Call Autofill_Sample_Type_Click
    MsgBox "Create new sample annotation test complete"
    
    'Clear the sample annotation
    Call Utilities.Clear_Columns("Raw_Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Merge_Status", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Type", HeaderRowNumber:=1, DataStartRowNumber:=2)
    
    'Prepare for merging
    Load_Sample_Annot_Raw.Sample_Name_Text.Text = "Sample"
    Load_Sample_Annot_Raw.Sample_Amount_Text.Text = "Cell Number"
    Load_Sample_Annot_Raw.ISTD_Mixture_Volume_Text.Text = "ISTD Volume"
    
    'Test merging with existing sample annotation with raw data
    Call Sample_Annot.Merge_With_Sample_Annot(RawDataFiles:=RawDataFiles, SampleAnnotFile:=SampleAnnotFile)
    Call Autofill_Sample_Type_Click
    MsgBox "Merging raw data with sample annotation test complete"
    
    'Clear the sample annotation
    Call Utilities.Clear_Columns("Raw_Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Merge_Status", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Type", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Amount", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Amount_Unit", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("ISTD_Mixture_Volume_[ul]", HeaderRowNumber:=1, DataStartRowNumber:=2)
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
    RawDataFiles = TestFolder & "DogCat.csv"
    
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
    Call Utilities.Clear_Columns("Raw_Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Dilution_Batch_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Dilution_Factor_[%]", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Injection_Volume_[uL]", HeaderRowNumber:=1, DataStartRowNumber:=2)

    'Clear the sample annotation
    Sheets("Sample_Annot").Activate
    Call Utilities.Clear_Columns("Raw_Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Merge_Status", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Type", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Amount", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("Sample_Amount_Unit", HeaderRowNumber:=1, DataStartRowNumber:=2)
    Call Utilities.Clear_Columns("ISTD_Mixture_Volume_[ul]", HeaderRowNumber:=1, DataStartRowNumber:=2)
    MsgBox "Dilution_Annot test complete"

    
TestFail:
    Application.EnableEvents = True
    Exit Sub
End Sub
