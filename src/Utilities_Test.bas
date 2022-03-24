Attribute VB_Name = "Utilities_Test"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
'Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    'Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub

'@TestInitialize
'Public Sub TestInitialize()
'    'this method runs before every test in the module.
'End Sub

'@TestCleanup
'Public Sub TestCleanup()
'    'this method runs after every test in the module.
'End Sub

'@TestMethod("Sheet Name Integrity Test")

'' Function: Check_Sheet_Code_Name_Exists_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Check_Sheet_Code_Name_Exists is working
''
'' Function will assert if the sheet code name exists
Public Sub Check_Sheet_Code_Name_Exists_Test()
    On Error GoTo TestFail
    
    Assert.AreEqual Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "Lists"), True
    Assert.AreEqual Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "Does not Exists"), False

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Sheet Name Integrity Test")

'' Function: Get_Sheet_By_Code_Name_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Get_Sheet_By_Code_Name is working
''
'' Function will assert if the sheet name is correct given
'' a provided code name.
Public Sub Get_Sheet_By_Code_Name_Test()
    On Error GoTo TestFail
    
    Dim ISTD_Annot_Worksheet As Worksheet
    Set ISTD_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "ISTDAnnotSheet")
    
    Assert.AreEqual ISTD_Annot_Worksheet.CodeName, "ISTDAnnotSheet"

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Annotation Properties")

'' Function: Get_Header_Col_Position_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Get_Header_Col_Position is working
''
'' Function will assert if the correct header/column position
'' is provided when user input a header/column name from the
'' sheet.
Public Sub Get_Header_Col_Position_Test()
    On Error GoTo TestFail
    
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
    
    Assert.AreEqual Utilities.Get_Header_Col_Position("Transition_Name", 1), CLng(1)
    Assert.AreEqual Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 1), CLng(2)
    
    'Check if it works in a sheet that is not active
    Assert.AreEqual Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 2, "ISTD_Annot"), CLng(1)
    
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

    Assert.AreEqual Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 2), CLng(1)
    Assert.AreEqual Utilities.Get_Header_Col_Position("ISTD_Conc_[ng/mL]", 3), CLng(2)
    Assert.AreEqual Utilities.Get_Header_Col_Position("ISTD_[MW]", 3), CLng(3)
    Assert.AreEqual Utilities.Get_Header_Col_Position("ISTD_Conc_[nM]", 3), CLng(5)
    
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
    
    Assert.AreEqual Utilities.Get_Header_Col_Position("Sample_Name", 1), CLng(3)
    Assert.AreEqual Utilities.Get_Header_Col_Position("Sample_Type", 1), CLng(4)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Annotation Properties")

'' Function: Last_Used_Row_Number_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Last_Used_Row_Number_Test is working
''
'' Function will assert if the number of used rows
'' is correct in the Lists sheet.
Public Sub Last_Used_Row_Number_Test()
    On Error GoTo TestFail
    
    ' Get the Lists worksheet from the active workbook
    ' The Lists is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Lists_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "Lists") = False Then
        MsgBox ("Sheet Lists is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Lists_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "Lists")
      
    Lists_Worksheet.Activate
    
    'Debug.Print Utilities.Last_Used_Row_Number
    Assert.AreEqual Utilities.Last_Used_Row_Number, CLng(22)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Letter Conversion")

'' Function: Convert_To_Letter_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Convert_To_Letter_Test is working
''
'' Function will assert if the correct alphabet is converted
'' when it receives an integer.
Public Sub Convert_To_Letter_Test()
    On Error GoTo TestFail
    
    Assert.AreEqual Utilities.Convert_To_Letter(1), "A"
    Assert.AreEqual Utilities.Convert_To_Letter(26), "Z"
    Assert.AreEqual Utilities.Convert_To_Letter(27), "AA"
    Assert.AreEqual Utilities.Convert_To_Letter(52), "AZ"
    Assert.AreEqual Utilities.Convert_To_Letter(53), "BA"
    Assert.AreEqual Utilities.Convert_To_Letter(520), "SZ"

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: Concantenate_String_Arrays_Test
'' --- Code
''  Public Sub Concantenate_String_Arrays_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Concantenate_String_Arrays_Test is working
''
'' Function will assert if two string array are correctly
'' concatenated to a singel string array.
'@TestMethod("String Array Test")
Public Sub Concantenate_String_Arrays_Test()
    On Error GoTo TestFail
    
    Dim TopArray(2) As String
    Dim BottomArray(2) As String
    Dim Concatenated_Array() As String
    Dim Concatenated_Array_Correct(5) As String
    
    TopArray(0) = "SM 36:0"
    TopArray(1) = "SM 36:1"
    TopArray(2) = "SM 36:2"
    BottomArray(0) = "SM 38:0"
    BottomArray(1) = "SM 38:1"
    BottomArray(2) = "SM 38:2"
    
    Concatenated_Array_Correct(0) = "SM 36:0"
    Concatenated_Array_Correct(1) = "SM 36:1"
    Concatenated_Array_Correct(2) = "SM 36:2"
    Concatenated_Array_Correct(3) = "SM 38:0"
    Concatenated_Array_Correct(4) = "SM 38:1"
    Concatenated_Array_Correct(5) = "SM 38:2"
    
    Concatenated_Array = Utilities.Concantenate_String_Arrays(TopArray:=TopArray, _
                                                              BottomArray:=BottomArray)
    
    Assert.SequenceEquals Concatenated_Array, Concatenated_Array_Correct

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("String Array Test")

'' Function: Get_String_Array_Len_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Get_String_Array_Len_Test is working
''
'' Function will assert if the total number of elements
'' in the input string array is correct.
Public Sub Get_String_Array_Len_Test()
    On Error GoTo TestFail
    
    Dim TestArray As Variant
    Dim EmptyArray As Variant
    
    TestArray = Array("11_PQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_PQC_2.d", _
                      "PQC_PC-PE-SM_01.d", "PQC1_LPC-LPE-PG-PI-PS_01.d", "PQC1_02.d", _
                      "11_BQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_BQC_2.d", "018_BQC_PQC01")
    EmptyArray = Array()
    
    Assert.AreEqual Utilities.Get_String_Array_Len(TestArray), CLng(8)
    Assert.AreEqual Utilities.Get_String_Array_Len(EmptyArray), CLng(0)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("String Array Test")

'' Function: Where_In_Array_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Where_In_Array_Test is working
''
'' Function will assert if the correct array position
'' is provided given the input array and an element to
'' search in the array.
Public Sub Where_In_Array_Test()
    On Error GoTo TestFail
    
    Dim TestArray As Variant
    Dim CorrectPositions(0 To 3) As String
    CorrectPositions(0) = "0"
    CorrectPositions(1) = "2"
    CorrectPositions(2) = "4"
    CorrectPositions(3) = "5"
    Dim Positions() As String
    
    'Ensure that it works and gives the right position
    TestArray = Array("Here", "11_PQC-2.d", "Here", "No", "Here", "Here")
    Positions = Utilities.Where_In_Array("Here", TestArray)
    Assert.SequenceEquals Positions, CorrectPositions
    
    'Ensure it gives an empty string array when there is no match
    Positions = Utilities.Where_In_Array("Her", TestArray)
    Assert.AreEqual Utilities.Get_String_Array_Len(Positions), CLng(0)
    
    'Ensure it gives an empty string array when test array is empty
    TestArray = Array()
    Positions = Utilities.Where_In_Array("Her", TestArray)
    Assert.AreEqual Utilities.Get_String_Array_Len(Positions), CLng(0)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("String Array Test")

'' Function: Is_In_Array_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Is_In_Array_Test is working
''
'' Function will assert if the function correctly
'' identified if an input string is in a given
'' string array.
Public Sub Is_In_Array_Test()
    On Error GoTo TestFail
    
    Dim TestArray As Variant
    TestArray = Array("Here", "11_PQC-2.d", "Here", "No", "Here", "Here")
    
    Assert.IsTrue Utilities.Is_In_Array("Here", TestArray)
    Assert.IsTrue Utilities.Is_In_Array("11_PQC-2.d", TestArray)
    Assert.IsFalse Utilities.Is_In_Array("NotHere", TestArray)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("String Array Test")

'' Function: Clear_DotD_In_Agilent_Data_File_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Clear_DotD_In_Agilent_Data_File_Test is working
''
'' Function will assert if the function removes
'' the ".d" at the end of an input string.
Public Sub Clear_DotD_In_Agilent_Data_File_Test()
    On Error GoTo TestFail
    
    Dim Sample_Name_Array(1) As String
    Dim Cleared_Sample_Name_Array() As String

    Sample_Name_Array(0) = "Sample_Name_1.d"
    Sample_Name_Array(1) = "Sample_Name_2.d"

    Cleared_Sample_Name_Array = Utilities.Clear_DotD_In_Agilent_Data_File(Sample_Name_Array)
    
    Assert.AreEqual Cleared_Sample_Name_Array(0), "Sample_Name_1"
    Assert.AreEqual Cleared_Sample_Name_Array(1), "Sample_Name_2"
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Sorting Test")

'' Function: Quick_Sort_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Quick_Sort_Test is working
''
'' Function will assert if an input string array
'' is sorted in alphabetical order correctly.
Public Sub Quick_Sort_Test()
    On Error GoTo TestFail
    
    Dim TestArray As Variant
    Dim SortedArray As Variant
    TestArray = Array("SM C36:2", "lipid", "Cer d18:1/C16:0")
    SortedArray = Array("Cer d18:1/C16:0", "SM C36:2", "lipid")
    Utilities.Quick_Sort ThisArray:=TestArray
    
    Assert.SequenceEquals TestArray, SortedArray

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Array From One Excel Column")

'' Function: Load_Columns_From_Excel_NoFilter_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Load_Columns_From_Excel is working
''
'' Function will assert if the Concentration_Unit in
'' the Lists Sheet is loaded correctly.
Public Sub Load_Columns_From_Excel_NoFilter_Test()
    On Error GoTo TestFail
    
    ' Get the Lists worksheet from the active workbook
    ' The Lists is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Lists_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "Lists") = False Then
        MsgBox ("Sheet Lists is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Lists_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "Lists")
      
    Lists_Worksheet.Activate
    
    Dim CorrectArray(0 To 4) As String
    CorrectArray(0) = "[M] or [umol/uL]"
    CorrectArray(1) = "[mM] or [nmol/uL]"
    CorrectArray(2) = "[uM] or [pmol/uL]"
    CorrectArray(3) = "[nM] or [fmol/uL]"
    CorrectArray(4) = "[pM] or [amol/uL]"
    
    Dim Concentration_Unit_Array() As String
    Concentration_Unit_Array = Utilities.Load_Columns_From_Excel("Concentration_Unit", HeaderRowNumber:=1, _
                                                                 DataStartRowNumber:=2, MessageBoxRequired:=True, _
                                                                 RemoveBlksAndReplicates:=True, _
                                                                 IgnoreHiddenRows:=False, IgnoreEmptyArray:=False)
    Assert.SequenceEquals Concentration_Unit_Array, CorrectArray

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Array From One Excel Column")

'' Function: Load_Columns_From_Excel_Filter_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Load_Columns_From_Excel is working
''
'' Function will assert if the filtered Sample_Type
'' (only load those that contains "QC") in
'' the Lists Sheet is loaded correctly.
Public Sub Load_Columns_From_Excel_Filter_Test()
    On Error GoTo TestFail
    
    ' Get the Lists worksheet from the active workbook
    ' The Lists is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Lists_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "Lists") = False Then
        MsgBox ("Sheet Lists is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Lists_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "Lists")
      
    Lists_Worksheet.Activate
    
    ActiveSheet.Range("SampleType").AutoFilter Field:=1
    
    'Check if the column Sample_Type exists
    'Dim Factor_pos As Integer
    'Factor_pos = Utilities.Get_Header_Col_Position("Factor", HeaderRowNumber:=1)
    
    'Filter Rows that contains the word "QC"
    ActiveSheet.Range("SampleType").AutoFilter Field:=1, _
                                               Criteria1:="*QC*", _
                                               VisibleDropDown:=True
                                          
    Dim CorrectArray(0 To 3) As String
    CorrectArray(0) = "EQC"
    CorrectArray(1) = "BQC"
    CorrectArray(2) = "TQC"
    CorrectArray(3) = "RQC"
    
    Dim Concentration_Unit_Array() As String
    Concentration_Unit_Array = Utilities.Load_Columns_From_Excel("Sample_Type", HeaderRowNumber:=1, _
                                                                 DataStartRowNumber:=2, MessageBoxRequired:=True, _
                                                                 RemoveBlksAndReplicates:=True, _
                                                                 IgnoreHiddenRows:=True, IgnoreEmptyArray:=False)
                                                                
    Assert.SequenceEquals Concentration_Unit_Array, CorrectArray

    GoTo TestExit
TestExit:
    ActiveSheet.Range("SampleType").AutoFilter Field:=1
    Exit Sub
TestFail:
    ActiveSheet.Range("SampleType").AutoFilter Field:=1
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Read Files")

'' Function: Read_File_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Read_File is working
''
'' Function will assert if it reads the file
'' Sample_Annotation_Example.csv correctly.
Public Sub Read_File_Test()
    On Error GoTo TestFail
    
    Dim SampleAnnotFile As String
    Dim TestFolder As String

    Dim Lines() As String

    TestFolder = ThisWorkbook.Path & "\Testdata\"
    SampleAnnotFile = TestFolder & "Sample_Annotation_Example.csv"

    Lines = Utilities.Read_File(SampleAnnotFile)
    
    Assert.AreEqual Lines(0), "Sample,ID,TimePoint,Cell Number,ISTD Volume"
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Read Files")

'' Function: Get_Delimiter_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Get_Delimiter is working
''
'' Function will assert if it correctly
'' get the delimiter from the file
'' Sample_Annotation_Example.csv.
Public Sub Get_Delimiter_Test()
    On Error GoTo TestFail

    Dim SampleAnnotFile As String
    Dim TestFolder As String

    Dim Delimiter As String

    TestFolder = ThisWorkbook.Path & "\Testdata\"
    SampleAnnotFile = TestFolder & "Sample_Annotation_Example.csv"

    Delimiter = Utilities.Get_Delimiter(SampleAnnotFile)
    
    Assert.AreEqual Delimiter, ","
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Read Files")

'' Function: Get_File_Base_Name_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Get_File_Base_Name is working
''
'' Function will assert if it correctly
'' get the file base name from a given
'' test file path.
Public Sub Get_File_Base_Name_Test()
    On Error GoTo TestFail
    
    Dim SampleAnnotFile As String
    Dim TestFolder As String
    Dim FileName As String

    TestFolder = ThisWorkbook.Path & "\Testdata\"
    SampleAnnotFile = TestFolder & "Sample_Annotation_Example.csv"
    
    FileName = Utilities.Get_File_Base_Name(SampleAnnotFile)
    
    Assert.AreEqual FileName, "Sample_Annotation_Example.csv"
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Read Files")

'' Function: Get_Raw_Data_File_Type_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Get_Raw_Data_File_Type is working
''
'' Function will assert if the correct raw data
'' file type is returned for an Agilent Wide Table
'' Form file, Agilent Compound Table file and
'' Sciex text file.
Public Sub Get_Raw_Data_File_Type_Test()
    
    Dim Lines() As String
    Dim Delimiter As String
    Dim FileName As String
    Dim RawDataFileType As String
    Dim TestFolder As String
    Dim RawDataFile As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFile = TestFolder & "AgilentRawDataTest1.csv"
    
    Lines = Utilities.Read_File(RawDataFile)
    Delimiter = Utilities.Get_Delimiter(RawDataFile)
    FileName = Utilities.Get_File_Base_Name(RawDataFile)
    RawDataFileType = Utilities.Get_Raw_Data_File_Type(Lines, Delimiter, FileName)
    
    'Debug.Print RawDataFileType
    Assert.AreEqual RawDataFileType, "AgilentWideForm"
    
    RawDataFile = TestFolder & "CompoundTableForm.csv"
    
    Lines = Utilities.Read_File(RawDataFile)
    Delimiter = Utilities.Get_Delimiter(RawDataFile)
    FileName = Utilities.Get_File_Base_Name(RawDataFile)
    RawDataFileType = Utilities.Get_Raw_Data_File_Type(Lines, Delimiter, FileName)
    
    'Debug.Print RawDataFileType
    Assert.AreEqual RawDataFileType, "AgilentCompoundForm"
    
    RawDataFile = TestFolder & "SciExTestData.txt"
    
    Lines = Utilities.Read_File(RawDataFile)
    Delimiter = Utilities.Get_Delimiter(RawDataFile)
    FileName = Utilities.Get_File_Base_Name(RawDataFile)
    RawDataFileType = Utilities.Get_Raw_Data_File_Type(Lines, Delimiter, FileName)
    
    'Debug.Print RawDataFileType
    Assert.AreEqual RawDataFileType, "Sciex"

End Sub

'@TestMethod("Load Data From 2Darray")

'' Function: Get_Header_Col_Position_From_2Darray_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Get_Header_Col_Position_From_2Darray is working
''
'' Function will assert if the correct column position
'' is provided.
Public Sub Get_Header_Col_Position_From_2Darray_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim TidyDataColumnFile As String

    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    TidyDataColumnFile = TestFolder & "TidyTransitionColumn.csv"
    
    'Read the csv files
    Dim Lines() As String
    Lines = Utilities.Read_File(TidyDataColumnFile)
    
    'Get column position of a given header name
    Dim HeaderColNumber As Variant
    
    ' Find which column is "Sample_Name" found in the "HeaderName" row.
    ' "HeaderName" row is the first row hence HeaderRowNumber is set as 0
    ' HeaderColNumber should return 0 as "Sample_Name" is the first column
    HeaderColNumber = Utilities.Get_Header_Col_Position_From_2Darray(Lines:=Lines, _
                                                                     HeaderName:="Sample_Name", _
                                                                     HeaderRowNumber:=0, _
                                                                     Delimiter:=",")

    'Debug.Print HeaderColNumber
    Assert.AreEqual CLng(HeaderColNumber), CLng(0)
                                                                     
    ' HeaderColNumber should return 1 as "LPC 14:0" is at the second column
    HeaderColNumber = Utilities.Get_Header_Col_Position_From_2Darray(Lines:=Lines, _
                                                                     HeaderName:="LPC 14:0", _
                                                                     HeaderRowNumber:=0, _
                                                                     Delimiter:=",")
    Assert.AreEqual CLng(HeaderColNumber), CLng(1)
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Load Data From 2Darray")

'' Function: Get_RowName_Position_From_2Darray_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Get_RowName_Position_From_2Darray is working
''
'' Function will assert if the correct row position
'' is provided.
Public Sub Get_RowName_Position_From_2Darray_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim TidyDataColumnFile As String

    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    TidyDataColumnFile = TestFolder & "TidyTransitionColumn.csv"
    
    'Read the csv files
    Dim Lines() As String
    Lines = Utilities.Read_File(TidyDataColumnFile)
    
    Dim RowNamePosition As Variant
    
    ' Find which row is "Sample_Name" found in the "RowName" column.
    ' "RowName" column is the first column hence RowNameNumber is set as 0
    ' RowNamePosition should return 0 as "Sample_Name" is the first row
    RowNamePosition = Utilities.Get_RowName_Position_From_2Darray(Lines:=Lines, _
                                                                  RowName:="Sample_Name", _
                                                                  RowNameNumber:=0, _
                                                                  Delimiter:=",")
    'Debug.Print RowNamePosition
    Assert.AreEqual CLng(RowNamePosition), CLng(0)
    
    ' RowNamePosition should return 1 as "Sample1" is the second row
    ' in the first column
    RowNamePosition = Utilities.Get_RowName_Position_From_2Darray(Lines:=Lines, _
                                                                  RowName:="Sample1", _
                                                                  RowNameNumber:=0, _
                                                                  Delimiter:=",")
    Assert.AreEqual CLng(RowNamePosition), CLng(1)
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Load Data From 2Darray")

'' Function: Load_Rows_From_2Darray_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Load_Rows_From_2Darray is working
''
'' Function will assert if a row from an input data in tabular
'' form is read correctly into a string array.
Public Sub Load_Rows_From_2Darray_Test()
    On Error GoTo TestFail

    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim Transition_Array_Correct(7) As String
    Dim TidyDataColumnFile As String
    
    'Create the correct array sequence
    Transition_Array_Correct(0) = "LPC 14:0"
    Transition_Array_Correct(1) = "LPC 15:0/LPC-O 16:0"
    Transition_Array_Correct(2) = "LPC 16:0"
    Transition_Array_Correct(3) = "SM 43:1"
    Transition_Array_Correct(4) = "SM 43:2"
    Transition_Array_Correct(5) = "SM 44:1"
    Transition_Array_Correct(6) = "SM 44:2"
    Transition_Array_Correct(7) = "SM 46:0"
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    TidyDataColumnFile = TestFolder & "TidyTransitionColumn.csv"
    
    'Read the csv files
    Dim Lines() As String
    Lines = Utilities.Read_File(TidyDataColumnFile)
    
    Transition_Array = Utilities.Load_Rows_From_2Darray(InputStringArray:=Transition_Array, _
                                                        Lines:=Lines, _
                                                        DataStartColumnNumber:=1, _
                                                        Delimiter:=",", _
                                                        RemoveBlksAndReplicates:=True, _
                                                        DataStartRowNumber:=0)
                                                        
    'Dim header_line_index As Long
    'For header_line_index = LBound(Transition_Array) To UBound(Transition_Array)
    '    Debug.Print Transition_Array(header_line_index)
    'Next header_line_index
    
    Assert.SequenceEquals Transition_Array, Transition_Array_Correct
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Load Data From 2Darray")

'' Function: Load_Columns_From_2Darray_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Load_Rows_From_2Darray is working
''
'' Function will assert if a column from an input data in tabular
'' form is read correctly into a string array.
Public Sub Load_Columns_From_2Darray_Test()
    On Error GoTo TestFail

    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim Transition_Array_Correct(6) As String
    Dim TidyDataRowFile As String
    
    'Create the correct array sequence
    Transition_Array_Correct(0) = "LPC 14:0"
    Transition_Array_Correct(1) = "LPC 15:0/LPC-O 16:0"
    Transition_Array_Correct(2) = "LPC 16:0"
    Transition_Array_Correct(3) = "SM 43:2"
    Transition_Array_Correct(4) = "SM 44:1"
    Transition_Array_Correct(5) = "SM 44:2"
    Transition_Array_Correct(6) = "SM 46:0"
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    TidyDataRowFile = TestFolder & "TidyTransitionRow.csv"
    
    'Read the csv files
    Dim Lines() As String
    Lines = Utilities.Read_File(TidyDataRowFile)
    
    Transition_Array = Utilities.Load_Columns_From_2Darray(InputStringArray:=Transition_Array, _
                                                           Lines:=Lines, _
                                                           DataStartColumnNumber:=0, _
                                                           DataStartRowNumber:=1, _
                                                           Delimiter:=",", _
                                                           RemoveBlksAndReplicates:=True)
                                                        
    'Dim header_line_index As Long
    'For header_line_index = LBound(Transition_Array) To UBound(Transition_Array)
    '    Debug.Print Transition_Array(header_line_index)
    'Next header_line_index
    
    Assert.SequenceEquals Transition_Array, Transition_Array_Correct
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Load and Clear Data in Excel")

'' Function: Load_To_Excel_And_Clear_Columns_Test
''
'' Description:
''
'' Function used to test if the function
'' Utilities.Load_To_Excel and
'' Utilities.Clear_Columns are working
''
'' Function will assert if a given string array is
'' loaded unto the excel sheet correctly. Once loaded,
'' it will check if the given array can be cleared.
Public Sub Load_To_Excel_And_Clear_Columns_Test()
    On Error GoTo TestFail
    
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

    Dim Transition_Array(3) As String
    Transition_Array(0) = "SM 44:0"
    Transition_Array(1) = "SM 44:1"
    Transition_Array(2) = "SM 46:2"
    Transition_Array(3) = "SM 46:3"
    
    Utilities.Load_To_Excel Data_Array:=Transition_Array, _
                            HeaderName:="Transition_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
                            
    Assert.AreEqual Transition_Name_Annot_Worksheet.Cells.Item(2, 1).Value, "SM 44:0"
    Assert.AreEqual Transition_Name_Annot_Worksheet.Cells.Item(3, 1).Value, "SM 44:1"
    Assert.AreEqual Transition_Name_Annot_Worksheet.Cells.Item(4, 1).Value, "SM 46:2"
    Assert.AreEqual Transition_Name_Annot_Worksheet.Cells.Item(5, 1).Value, "SM 46:3"
    
    Utilities.Clear_Columns HeaderToClear:="Transition_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
                            
    Assert.AreEqual Transition_Name_Annot_Worksheet.Cells.Item(2, 1).Value, vbNullString
    Assert.AreEqual Transition_Name_Annot_Worksheet.Cells.Item(3, 1).Value, vbNullString
    Assert.AreEqual Transition_Name_Annot_Worksheet.Cells.Item(4, 1).Value, vbNullString
    Assert.AreEqual Transition_Name_Annot_Worksheet.Cells.Item(5, 1).Value, vbNullString
    
    GoTo TestExit
TestExit:
    Transition_Name_Annot_Worksheet.Cells.Item(2, 1).Value = vbNullString
    Transition_Name_Annot_Worksheet.Cells.Item(3, 1).Value = vbNullString
    Transition_Name_Annot_Worksheet.Cells.Item(4, 1).Value = vbNullString
    Transition_Name_Annot_Worksheet.Cells.Item(5, 1).Value = vbNullString
    Exit Sub
TestFail:
    Transition_Name_Annot_Worksheet.Cells.Item(2, 1).Value = vbNullString
    Transition_Name_Annot_Worksheet.Cells.Item(3, 1).Value = vbNullString
    Transition_Name_Annot_Worksheet.Cells.Item(4, 1).Value = vbNullString
    Transition_Name_Annot_Worksheet.Cells.Item(5, 1).Value = vbNullString
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
