Attribute VB_Name = "Integration_Test"
Option Explicit
'@Folder("Tests")
'@IgnoreModule ProcedureNotUsed
'@IgnoreModule IntegerDataType

'' Function: Run_Integration_Test
'' --- Code
''  Public Sub Run_Integration_Test()
'' ---
''
'' Description:
''
'' A simple wrapper function to run the following integration test
''
'' - Integration_Test.Nothing_To_Transfer_Test
'' - Integration_Test.Transition_Name_and_ISTD_Annot_Integration_Test
'' - Integration_Test.Sample_Annot_Integration_Test
'' - Integration_Test.Sample_Annot_and_Dilution_Annot_Integration_Test
''
Public Sub Run_Integration_Test()
    Integration_Test.Nothing_To_Transfer_Test
    Integration_Test.Transition_Name_and_ISTD_Annot_Integration_Test
    Integration_Test.Sample_Annot_Integration_Test
    Integration_Test.Sample_Annot_and_Dilution_Annot_Integration_Test
End Sub

'' Function: Nothing_To_Transfer_Test
'' --- Code
''  Public Sub Nothing_To_Transfer_Test()
'' ---
''
'' Description:
''
'' This function test that
''
'' - The button works when there are no Transition_Name_ISTD to validate and transfer
''   from Transition_Name_Annot to ISTD_Annot.
'' - The button works when there are no samples with sample type RQC to transfer from
''   Sample_Annot to Dilution_Annot.
''
Public Sub Nothing_To_Transfer_Test()
    On Error GoTo TestFail
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    'To ensure that Filters does not affect the assignment
    Utilities.Remove_Filter_Settings
    
    'Test that the button works when there are no
    'Transition_Name_ISTD to validate and transfer
    'from Transition_Name_Annot to ISTD_Annot
    
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
    
    'When both columns are empty
    Transition_Name_Annot_Buttons.Load_Transition_Name_ISTD_Click
    Transition_Name_Annot_Buttons.Validate_ISTD_Click
    
    'When only Transition_Name_ISTD column is empty
    Dim Transition_Array(2) As String
    Transition_Array(0) = "Transition_Array1"
    Transition_Array(1) = "Transition_Array2"
    Utilities.Load_To_Excel Data_Array:=Transition_Array, _
                            HeaderName:="Transition_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
    Transition_Name_Annot_Buttons.Validate_ISTD_Click
    Utilities.Clear_Columns HeaderToClear:="Transition_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    
    'When only Transition_Name column is empty
    Transition_Array(0) = "Transition_ISTD1"
    Transition_Array(1) = "Transition_ISTD2"
    Utilities.Load_To_Excel Data_Array:=Transition_Array, _
                            HeaderName:="Transition_Name_ISTD", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
    Transition_Name_Annot_Buttons.Validate_ISTD_Click
    Utilities.Clear_Columns HeaderToClear:="Transition_Name_ISTD", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    
    MsgBox "Nothing to transfer in Transition_Name_Annot test complete"
    
    'Test that the button works when there are no
    'samples with sample type RQC to transfer from Sample_Annot
    'to Dilution_Annot
    
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

    Sample_Annot_Buttons.Load_Sample_Name_To_Dilution_Annot_Click
    Utilities.Clear_Columns HeaderToClear:="Data_File_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Merge_Status", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Type", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    MsgBox "Nothing to transfer in Sample_Annot test complete"

    'Excel resume monitoring the sheet
    Application.EnableEvents = True
    
TestFail:
    Application.EnableEvents = True
    Exit Sub
End Sub

'' Function: Transition_Name_and_ISTD_Annot_Integration_Test
'' --- Code
''  Public Sub Transition_Name_and_ISTD_Annot_Integration_Test()
'' ---
''
'' Description:
''
'' This function test that
''
'' - Is able to create a new transition annotation from tidy/tabular data file with transitons as column variables.
'' - Is able to create a new transition annotation from tidy/tabular data file with transitons as row observations.
'' - Is able to indicate if users give an invalid file.
'' - Is able to create a new transition annotation from Agilent raw data file (Wide Table Form).
'' - Is able to verify and indicate invalid ISTD.
'' - Is able to transfer unique entries in Transition_Name_ISTD to ISTD_Annot sheet.
'' - Is able to convert the concentration inputs to nM.
''
Public Sub Transition_Name_and_ISTD_Annot_Integration_Test()
    On Error GoTo TestFail
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    'To ensure that Filters does not affect the assignment
    Utilities.Remove_Filter_Settings
    
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
    
    Dim TestFolder As String
    Dim FileThere As Boolean
    Dim TidyDataRowFiles As String
    Dim TidyDataColumnFiles As String
    Dim InvalidRawDataFiles As String
    Dim AgilentRawDataFiles As String
    Dim JoinedFiles As String
    Dim Transition_Array() As String
    Dim Transition_Name_ISTD_ColLetter As String
    Dim Transition_Name_ISTD_ColNumber As Integer
    Dim xFileNames() As String
    Dim xFileName As Variant
    'Dim ISTD_Array() As String
    
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
        FileThere = (Dir(xFileName) > vbNullString)
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
                                                                              
    Utilities.Overwrite_Header HeaderName:="Transition_Name", _
                               HeaderRowNumber:=1, _
                               DataStartRowNumber:=2
    Utilities.Load_To_Excel Data_Array:=Transition_Array, _
                            HeaderName:="Transition_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=True
                                                   
    MsgBox "Create new transition annotation from tidy column data test complete"
    
    'Test creating a new transition annotation from tidy data file with transitons as row observations
    Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array_Tidy(TidyDataFiles:=TidyDataRowFiles, _
                                                                              DataFileType:="csv", _
                                                                              TransitionProperty:="Read as row observations", _
                                                                              StartingRowNum:=2, _
                                                                              StartingColumnNum:=1)
                                                   
    Utilities.Overwrite_Header HeaderName:="Transition_Name", _
                               HeaderRowNumber:=1, _
                               DataStartRowNumber:=2
    Utilities.Load_To_Excel Data_Array:=Transition_Array, _
                            HeaderName:="Transition_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=True
                                 
    MsgBox "Create new transition annotation from tidy row data test complete"
    
    Utilities.Clear_Columns HeaderToClear:="Transition_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    
    'Testing with an invalid file
    Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array_Raw(RawDataFiles:=InvalidRawDataFiles)
    MsgBox "Invalid file input test complete"
    
    'Load the transition names and load it to excel
    Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array_Raw(RawDataFiles:=AgilentRawDataFiles)
    Utilities.Overwrite_Header HeaderName:="Transition_Name", _
                               HeaderRowNumber:=1, _
                               DataStartRowNumber:=2
    Utilities.Load_To_Excel Data_Array:=Transition_Array, _
                            HeaderName:="Transition_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=True
    
    'Excel resume monitoring the sheet
    Application.EnableEvents = True
    
    'Fill in the ISTD automatically
    Transition_Name_ISTD_ColNumber = Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 1)
    Transition_Name_ISTD_ColLetter = Utilities.Convert_To_Letter(Transition_Name_ISTD_ColNumber)
    Transition_Name_Annot_Worksheet.Range(Transition_Name_ISTD_ColLetter & "2:" & Transition_Name_ISTD_ColLetter & "23").Value = "LPC 17:0"
    Transition_Name_Annot_Worksheet.Range(Transition_Name_ISTD_ColLetter & "24:" & Transition_Name_ISTD_ColLetter & "31").Value = "MHC d18:1/16:0d3 (IS)"
    
    'Validate ISTD, ensure that wrong ISTD is detected
    Transition_Name_Annot_Buttons.Validate_ISTD_Click Testing:=True
    
    'Correct the wrong ISTD, validate ISTD and transfer them to ISTD_Annot sheet
    Transition_Name_Annot_Worksheet.Range(Transition_Name_ISTD_ColLetter & "2:" & Transition_Name_ISTD_ColLetter & "23").Value = "LPC 17:0 (IS)"
    Transition_Name_Annot_Buttons.Validate_ISTD_Click Testing:=True
    
    Transition_Name_Annot_Buttons.Load_Transition_Name_ISTD_Click
    
    'Clear the columns as the test for this worksheet is complete
    Transition_Name_Annot_Worksheet.Activate
    Utilities.Clear_Columns HeaderToClear:="Transition_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Transition_Name_ISTD", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    MsgBox "Transition_Name_Annot test complete"
    
    'Proceed with the ISTD_Annot test
    
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
    
    Dim ISTD_Conc_ng_ColNumber As Integer
    Dim ISTD_MW_ColNumber As Integer
    Dim ISTD_Conc_ng_ColLetter As String
    Dim ISTD_MW_ColLetter As String
    ISTD_Conc_ng_ColNumber = Utilities.Get_Header_Col_Position("ISTD_Conc_[ng/mL]", 3)
    ISTD_MW_ColNumber = Utilities.Get_Header_Col_Position("ISTD_[MW]", 3)
    ISTD_Conc_ng_ColLetter = Utilities.Convert_To_Letter(ISTD_Conc_ng_ColNumber)
    ISTD_MW_ColLetter = Utilities.Convert_To_Letter(ISTD_MW_ColNumber)
    
    'Fill in the concentration values automatically
    'ISTD_Annot_Worksheet.Range.Item(ISTD_Conc_ng_ColLetter & 4 & ":" & ISTD_MW_ColLetter & 5) = [{100,2,30,10}]
    
    ISTD_Annot_Worksheet.Range(ISTD_Conc_ng_ColLetter & 4 & ":" & _
                               ISTD_MW_ColLetter & 5).Value = Application.Transpose(Array("100", "2", _
                                                                                          "30", "10"))
    
    'Perform the calculation
    ISTD_Annot_Buttons.Convert_To_Nanomolar_Click
    MsgBox "Calculation is complete"
    
    'Clear the columns as the test for this worksheet is complete
    Utilities.Clear_Columns HeaderToClear:="Transition_Name_ISTD", _
                            HeaderRowNumber:=2, _
                            DataStartRowNumber:=4
    Utilities.Clear_Columns HeaderToClear:="ISTD_Conc_[ng/mL]", _
                            HeaderRowNumber:=3, _
                            DataStartRowNumber:=4
    Utilities.Clear_Columns HeaderToClear:="ISTD_[MW]", _
                            HeaderRowNumber:=3, _
                            DataStartRowNumber:=4
    Utilities.Clear_Columns HeaderToClear:="ISTD_Conc_[nM]", _
                            HeaderRowNumber:=3, _
                            DataStartRowNumber:=4
    Utilities.Clear_Columns HeaderToClear:="Custom_Unit", _
                            HeaderRowNumber:=2, _
                            DataStartRowNumber:=4
    MsgBox "ISTD_Annot test complete"
    
TestFail:
    Application.EnableEvents = True
    Exit Sub
End Sub

'' Function: Sample_Annot_Integration_Test
'' --- Code
''  Public Sub Sample_Annot_Integration_Test()
'' ---
''
'' Description:
''
'' This function test that
''
'' - Is able to create a new sample annotation from Agilent raw data file (Wide Table Form).
'' - Is able to autofill blank entries in the Sample_Type column with "SPL".
'' - Is able to autofill concentration units in the Concentration_Unit column.
'' - Is able to merge Agilent raw data file with input sample annotation file in csv.
'' - Is able to create a new transition annotation from tidy/tabular data file with sample names as column variables.
'' - Is able to create a new transition annotation from tidy/tabular data file with sample names as row observations.
''
Public Sub Sample_Annot_Integration_Test()
    On Error GoTo TestFail
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = True
    'To ensure that Filters does not affect the assignment
    Utilities.Remove_Filter_Settings
    
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
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim TidyDataRowFiles As String
    Dim TidyDataColumnFiles As String
    Dim JoinedFiles As String
    Dim SampleAnnotFile As String
    Dim xFileNames() As String
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
    'Dim MS_File_Array() As String
    'Dim Sample_Name_Array_from_Raw_Data() As String
    
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
        FileThere = (Dir(xFileName) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Check if the sample annotation file exists
    FileThere = (Dir(SampleAnnotFile) > vbNullString)
    If FileThere = False Then
        MsgBox "File name " & SampleAnnotFile & " cannot be found."
        End
    End If
    
    'Test creating a new sample annotation from raw data file
    Sample_Annot.Create_New_Sample_Annot_Raw RawDataFiles:=RawDataFiles
    Sample_Annot_Buttons.Autofill_Sample_Type_Click
    
    'Call Autofill_Concentration_Unit_Click
    MsgBox "Create new sample annotation test from raw data complete"
    
    'Clear the sample annotation
    Utilities.Clear_Columns HeaderToClear:="Data_File_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Merge_Status", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Type", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    
    'Fill in the Sample_Amount_Unit
    Dim Sample_Amount_Unit_ColNumber As Integer
    Dim Sample_Amount_Unit_ColLetter As String
    Sample_Amount_Unit_ColNumber = Utilities.Get_Header_Col_Position("Sample_Amount_Unit", 1)
    Sample_Amount_Unit_ColLetter = Utilities.Convert_To_Letter(Sample_Amount_Unit_ColNumber)
    
    'Fill in the concentration values automatically
    Sample_Annot_Worksheet.Range(Sample_Amount_Unit_ColLetter & 2 & ":" & _
                                 Sample_Amount_Unit_ColLetter & 7).Value = Application.Transpose(Array("cell_number", "mg_dry_weight", _
                                                                                                       "mg_fresh_weight", "ng_DNA", _
                                                                                                       "ug_total_protein", "uL"))
    Sample_Annot_Buttons.Autofill_Concentration_Unit_Click Testing:=True
    MsgBox "Autofill concentration unit test complete"
    
    'Clear the sample annotation
    Utilities.Clear_Columns HeaderToClear:="Sample_Amount_Unit", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Concentration_Unit", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    
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
    Utilities.Clear_Columns HeaderToClear:="Data_File_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Merge_Status", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Type", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Amount", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Amount_Unit", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="ISTD_Mixture_Volume_[ul]", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    
    'Test creating a new sample annotation from tidy data file
    Sample_Annot.Create_New_Sample_Annot_Tidy TidyDataFiles:=TidyDataColumnFiles, _
                                              DataFileType:="csv", _
                                              SampleProperty:="Read as column variables", _
                                              StartingRowNum:=1, _
                                              StartingColumnNum:=2
                                                   
    MsgBox "Create new sample annotation from tidy column data test complete"
    
    Sample_Annot.Create_New_Sample_Annot_Tidy TidyDataFiles:=TidyDataRowFiles, _
                                              DataFileType:="csv", _
                                              SampleProperty:="Read as row observations", _
                                              StartingRowNum:=2, _
                                              StartingColumnNum:=1
                                                   
    MsgBox "Create new sample annotation from tidy row data test complete"
    Utilities.Clear_Columns HeaderToClear:="Data_File_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Merge_Status", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Type", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    
    MsgBox "Sample_Annot test complete"
       
TestFail:
    Application.EnableEvents = True
    Exit Sub
End Sub

'' Function: Sample_Annot_and_Dilution_Annot_Integration_Test
'' --- Code
''  Public Sub Sample_Annot_and_Dilution_Annot_Integration_Test()
'' ---
''
'' Description:
''
'' This function test that
''
'' - Is able to create a new sample annotation from Agilent raw data file (Wide Table Form).
'' - Is able to transfer samples whose "Sample_Type" value is "RQC" to the Dilution_Annot sheet.
''
Public Sub Sample_Annot_and_Dilution_Annot_Integration_Test()
    On Error GoTo TestFail
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = True
    'To ensure that Filters does not affect the assignment
    Utilities.Remove_Filter_Settings
    
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
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim xFileNames() As String
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
    'Dim MS_File_Array() As String
    'Dim Sample_Name_Array_from_Raw_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "DilutionTest.csv"
    
    xFileNames = Split(RawDataFiles, ";")
    
    'Check if the data file exists
    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Test creating a new sample annotation
    Sample_Annot.Create_New_Sample_Annot_Raw RawDataFiles:=RawDataFiles
    MsgBox "Load RQC samples to copy"
    
    'Load to Dilution_Annot
    Sample_Annot_Buttons.Load_Sample_Name_To_Dilution_Annot_Click
    MsgBox "Copy RQC Samples to Dilution Annot complete"
    
    'Clear the dilution annotation
    
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
    
    Utilities.Clear_Columns HeaderToClear:="Data_File_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Dilution_Batch_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Relative_Sample_Amount_[%]", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Injection_Volume_[uL]", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2

    'Clear the sample annotation
    Sample_Annot_Worksheet.Activate
    
    Utilities.Clear_Columns HeaderToClear:="Data_File_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Merge_Status", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Type", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Amount", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Amount_Unit", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="ISTD_Mixture_Volume_[uL]", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    MsgBox "Dilution_Annot test complete"

    
TestFail:
    Application.EnableEvents = True
    Exit Sub
End Sub
