Attribute VB_Name = "Load_Raw_Data_Test"
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

'@TestMethod("Get Transition_Annot From Raw MS Data")

'' Function: Get_Transition_Array_Wide_Table_Test
''
'' Description:
''
'' Function used to test if the function
'' Load_Raw_Data.Get_Transition_Array_Raw is working
''
'' Test files are
''
''  - AgilentRawDataTest1.csv
''
'' Function will assert if Transition_Array has 30 elements
''
Public Sub Get_Transition_Array_Wide_Table_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim RawDataFiles As String
    Dim FileThere As Boolean
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "AgilentRawDataTest1.csv"
    
    'Check if the data file exists
    FileThere = (Dir(RawDataFiles) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & RawDataFiles & " cannot be found."
            End
        End If
    
    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=RawDataFiles)
    
    Assert.AreEqual Utilities.Get_String_Array_Len(Transition_Array), CLng(30)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Transition_Annot From Raw MS Data")

'' Function: Get_Transition_Array_Wide_Table_With_Qualifier_Test
''
'' Description:
''
'' Function used to test if the function
'' Load_Raw_Data.Get_Transition_Array_Raw is working
''
'' Test files are
''
''  - AgilentRawDataTest3_Qualifier.csv
''
'' Function will assert if Transition_Array has 15 elements
''
Public Sub Get_Transition_Array_Wide_Table_With_Qualifier_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim RawDataFiles As String
    Dim FileThere As Boolean
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "AgilentRawDataTest3_Qualifier.csv"
    
    'Check if the data file exists
    FileThere = (Dir(RawDataFiles) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & RawDataFiles & " cannot be found."
            End
        End If
    
    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=RawDataFiles)
    
    Assert.AreEqual Utilities.Get_String_Array_Len(Transition_Array), CLng(15)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Transition_Annot From Raw MS Data")

'' Function: Get_Transition_Array_Compound_Table_Test
''
'' Description:
''
'' Function used to test if the function
'' Load_Raw_Data.Get_Transition_Array_Raw is working
''
'' Test files are
''
''  - CompoundTableForm.csv
''
'' Function will assert if Transition_Array has 122 elements
''
Public Sub Get_Transition_Array_Compound_Table_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim RawDataFiles As String
    Dim FileThere As Boolean
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "CompoundTableForm.csv"
    
    'Check if the data file exists
    FileThere = (Dir(RawDataFiles) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & RawDataFiles & " cannot be found."
            End
        End If

    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=RawDataFiles)
    
    Assert.AreEqual Utilities.Get_String_Array_Len(Transition_Array), CLng(122)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Transition_Annot From Raw MS Data")

'' Function: Get_Transition_Array_Agilent_Compound_Test
''
'' Description:
''
'' Function used to test if the function
'' Load_Raw_Data.Get_Transition_Array_Agilent_Compound is working
''
'' Test files are
''
''  - CompoundTableForm_Qualifier.csv
''
'' Function will assert if Transition_Name_And_Qualifier_Transition_Column_Indexes is
'' {"1", "10", "14", "18"}.
''
'' Function will assert if Transition_Array has 15 elements
''
'' Function will assert if Transition_Array is able to get two
'' qualifiers for Sph d16:1 and three qualifiers for Sph d18:0
''
'' First three elements are
'' {"Sph d16:1", "Qualifier (272.2 -> 236.1)", "Qualifier (272.2 -> 224.1)"}
''
'' The sixth to ninth elements are
'' {"Sph d18:0", "Qualifier (302.3 -> 266.2)", "Qualifier (302.3 -> 254.2)", "Qualifier (302.3 -> 60.2)"}
''
Public Sub Get_Transition_Array_Agilent_Compound_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim RawDataFile As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFile = TestFolder & "CompoundTableForm_Qualifier.csv"
    
    Dim Lines() As String
    Dim Delimiter As String
    Lines = Utilities.Read_File(RawDataFile)
    Delimiter = Utilities.Get_Delimiter(RawDataFile)
    
    Dim first_header_line() As String
    Dim second_header_line() As String
    Dim first_header_line_index As Long
    Dim second_header_line_index As Long
    first_header_line = Split(Lines(0), Delimiter)
    second_header_line = Split(Lines(1), Delimiter)
    
    Dim Transition_Name_And_Qualifier_Transition_Column_Indexes() As Long
    Dim Transition_Name_And_Qualifier_Transition_Column_Indexes_Length As Long
    Transition_Name_And_Qualifier_Transition_Column_Indexes_Length = 0
    
    'Get the index of compound method name
    'It should appear before the qualifier
    For second_header_line_index = LBound(second_header_line) To UBound(second_header_line)
        If second_header_line(second_header_line_index) = "Name" Then
            ReDim Preserve Transition_Name_And_Qualifier_Transition_Column_Indexes(Transition_Name_And_Qualifier_Transition_Column_Indexes_Length)
            Transition_Name_And_Qualifier_Transition_Column_Indexes(Transition_Name_And_Qualifier_Transition_Column_Indexes_Length) = second_header_line_index
            Transition_Name_And_Qualifier_Transition_Column_Indexes_Length = Transition_Name_And_Qualifier_Transition_Column_Indexes_Length + 1
            Exit For
        End If
    Next second_header_line_index
    
    'Do a forward fill on the first header line
    Dim fill_header_word As String
    fill_header_word = first_header_line(0)
    
    For first_header_line_index = LBound(first_header_line) To UBound(first_header_line)
        If first_header_line(first_header_line_index) = vbNullString Then
            first_header_line(first_header_line_index) = fill_header_word
        Else
            fill_header_word = first_header_line(first_header_line_index)
        End If
    Next first_header_line_index
    
    'Get the index of qualifier method transition
    'Get the index of data file
    'Get the max number of qualifier a transition can have
    Dim Qualifier_Method_Col As RegExp
    Set Qualifier_Method_Col = New RegExp
    Dim Transition_Col As RegExp
    Set Transition_Col = New RegExp
    Dim DataFileName_Col As RegExp
    Set DataFileName_Col = New RegExp
    
    Qualifier_Method_Col.Pattern = "Qualifier \d Method"
    Transition_Col.Pattern = "Transition"
    DataFileName_Col.Pattern = "Data File"
    
    Dim isQualifier_Method_Col As Boolean
    Dim isTransition_Col As Boolean
    Dim isDataFileName_Col As Boolean
    Dim Qualifier_Method_Col_BoolArrray() As Boolean
    Dim DataFileName_Col_BoolArrray() As Boolean
    
    Dim No_of_Qualifier_Method_Transition As Long
    Dim No_of_DataFileName_Col As Long
    Dim No_of_Qual_per_Transition As Long
    No_of_Qualifier_Method_Transition = 0
    No_of_DataFileName_Col = 0
    
    Dim ArrayLength As Long
    ArrayLength = 0
    
    For first_header_line_index = LBound(first_header_line) To UBound(first_header_line)
    
        isQualifier_Method_Col = Qualifier_Method_Col.Test(first_header_line(first_header_line_index))
        isTransition_Col = Transition_Col.Test(second_header_line(first_header_line_index))
        isDataFileName_Col = DataFileName_Col.Test(second_header_line(first_header_line_index))
    
        ReDim Preserve Qualifier_Method_Col_BoolArrray(ArrayLength)
        ReDim Preserve DataFileName_Col_BoolArrray(ArrayLength)
        Qualifier_Method_Col_BoolArrray(ArrayLength) = isQualifier_Method_Col And isTransition_Col
        DataFileName_Col_BoolArrray(ArrayLength) = isDataFileName_Col
    
        If isQualifier_Method_Col And isTransition_Col Then
            No_of_Qualifier_Method_Transition = No_of_Qualifier_Method_Transition + 1
            ReDim Preserve Transition_Name_And_Qualifier_Transition_Column_Indexes(Transition_Name_And_Qualifier_Transition_Column_Indexes_Length)
            Transition_Name_And_Qualifier_Transition_Column_Indexes(Transition_Name_And_Qualifier_Transition_Column_Indexes_Length) = first_header_line_index
            Transition_Name_And_Qualifier_Transition_Column_Indexes_Length = Transition_Name_And_Qualifier_Transition_Column_Indexes_Length + 1
        ElseIf isDataFileName_Col Then
            No_of_DataFileName_Col = No_of_DataFileName_Col + 1
        End If
    
        ArrayLength = ArrayLength + 1
    
    Next first_header_line_index
    
    No_of_Qual_per_Transition = No_of_Qualifier_Method_Transition \ No_of_DataFileName_Col
    ReDim Preserve Transition_Name_And_Qualifier_Transition_Column_Indexes(No_of_Qual_per_Transition)
        
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Agilent_Compound(Transition_Array:=Transition_Array, _
                                                                           Lines:=Lines, _
                                                                           Transition_Name_And_Qualifier_Transition_Column_Indexes:=Transition_Name_And_Qualifier_Transition_Column_Indexes, _
                                                                           DataStartRowNumber:=2, _
                                                                           Delimiter:=Delimiter, _
                                                                           RemoveBlksAndReplicates:=True, _
                                                                           IgnoreEmptyArray:=True)
    
    Assert.AreEqual Transition_Name_And_Qualifier_Transition_Column_Indexes(0), CLng(1)
    Assert.AreEqual Transition_Name_And_Qualifier_Transition_Column_Indexes(1), CLng(10)
    Assert.AreEqual Transition_Name_And_Qualifier_Transition_Column_Indexes(2), CLng(14)
    Assert.AreEqual Transition_Name_And_Qualifier_Transition_Column_Indexes(3), CLng(18)
    
    Assert.AreEqual Utilities.Get_String_Array_Len(Transition_Array), CLng(15)
    
    Assert.AreEqual Transition_Array(0), "Sph d16:1"
    Assert.AreEqual Transition_Array(1), "Qualifier (272.2 -> 236.1)"
    Assert.AreEqual Transition_Array(2), "Qualifier (272.2 -> 224.1)"
    
    Assert.AreEqual Transition_Array(5), "Sph d18:0"
    Assert.AreEqual Transition_Array(6), "Qualifier (302.3 -> 266.2)"
    Assert.AreEqual Transition_Array(7), "Qualifier (302.3 -> 254.2)"
    Assert.AreEqual Transition_Array(8), "Qualifier (302.3 -> 60.2)"

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Transition_Annot From Raw MS Data")

'' Function: Get_Transition_Array_Compound_Table_With_Qualifier_Test
''
'' Description:
''
'' Function used to test if the function
'' Load_Raw_Data.Get_Transition_Array_Raw is working
''
'' Test files are
''
''  - CompoundTableForm_Qualifier.csv
''
'' Function will assert if Transition_Array has 15 elements
''
Public Sub Get_Transition_Array_Compound_Table_With_Qualifier_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim RawDataFiles As String
    Dim FileThere As Boolean
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "CompoundTableForm_Qualifier.csv"
    
    'Check if the data file exists
    FileThere = (Dir(RawDataFiles) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & RawDataFiles & " cannot be found."
            End
        End If

    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=RawDataFiles)
    
    Assert.AreEqual Utilities.Get_String_Array_Len(Transition_Array), CLng(15)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Transition_Annot From Raw MS Data")

'' Function: Get_Transition_Array_SciEx_Data_Test
''
'' Description:
''
'' Function used to test if the function
'' Load_Raw_Data.Get_Transition_Array_Raw is working
''
'' Test files are
''
''  - SciExTestData.txt
''
'' Function will assert if Transition_Array has 224 elements
''
Public Sub Get_Transition_Array_SciEx_Data_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim RawDataFiles As String
    Dim FileThere As Boolean
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "SciExTestData.txt"

    'Check if the data file exists
    FileThere = (Dir(RawDataFiles) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & RawDataFiles & " cannot be found."
            End
        End If
  
    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=RawDataFiles)
    
    Assert.AreEqual Utilities.Get_String_Array_Len(Transition_Array), CLng(224)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Transition_Annot From Raw MS Data")

'' Function: Get_Transition_Array_Multiple_Data_Test
''
'' Description:
''
'' Function used to test if the function
'' Load_Raw_Data.Get_Transition_Array_Raw is working
''
'' Test files are
''
''  - MultipleDataTest1.csv
''  - MultipleDataTest2.csv
''  - SciExTestData.txt
''
'' Function will assert if Transition_Array has 653 elements
''
Public Sub Get_Transition_Array_Multiple_Data_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim RawDataFiles As String
    Dim RawDataFilesArray() As String
    Dim RawDataFile As Variant
    Dim JoinedFiles As String
    Dim FileThere As Boolean
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "MultipleDataTest1.csv" & ";" & _
                   TestFolder & "MultipleDataTest2.csv" & ";" & _
                   TestFolder & "SciExTestData.txt"
    RawDataFilesArray = Split(RawDataFiles, ";")
                       
    'Check if the data file exists
    For Each RawDataFile In RawDataFilesArray
        FileThere = (Dir(RawDataFile) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & RawDataFile & " cannot be found."
            End
        End If
    Next RawDataFile
    
    JoinedFiles = Join(RawDataFilesArray, ";")
    
    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=JoinedFiles)
    
    Assert.AreEqual Utilities.Get_String_Array_Len(Transition_Array), CLng(653)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Transition_Annot From Raw MS Data")

'' Function: Get_Transition_Array_Invalid_Data_Test
''
'' Description:
''
'' Function used to test if the function
'' Load_Raw_Data.Get_Transition_Array_Raw is working
''
'' Test files are
''
''  - InvalidDataTest1.csv
''
'' Function will assert if Transition_Array has 0 elements
''
Public Sub Get_Transition_Array_Invalid_Data_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim RawDataFiles As String
    Dim FileThere As Boolean
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "InvalidDataTest1.csv"
    
    'Check if the data file exists
    FileThere = (Dir(RawDataFiles) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & RawDataFiles & " cannot be found."
            End
        End If
    
    'Load the transition names into an array
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=RawDataFiles)
    
    Assert.AreEqual Utilities.Get_String_Array_Len(Transition_Array), CLng(0)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Sample_Annot From Raw MS Data")

'' Function: Get_Sample_Name_Array_Wide_Table_Test
''
'' Description:
''
'' Function used to test if the function
'' Load_Raw_Data.Get_Sample_Name_Array is working
''
'' Test files are
''
''  - MultipleDataTest2.csv
''
'' Ensure that an empty string array is declared
'' MS_File_Array()
''
'' Function will assert if Sample_Name_Array has 533 elements
'' Function will assert if MS_File_Array has 533 elements
''
Public Sub Get_Sample_Name_Array_Wide_Table_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim RawDataFilesArray() As String
    Dim RawDataFile As Variant
    Dim FileThere As Boolean

    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "MultipleDataTest2.csv"
    RawDataFilesArray = Split(RawDataFiles, ";")
    
    'Check if the data file exists
    For Each RawDataFile In RawDataFilesArray
        FileThere = (Dir(RawDataFile) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & RawDataFile & " cannot be found."
            End
        End If
    Next RawDataFile

    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray, _
                                                                          MS_File_Array)

    Assert.AreEqual Utilities.Get_String_Array_Len(Sample_Name_Array_from_Raw_Data), CLng(533)
    Assert.AreEqual Utilities.Get_String_Array_Len(MS_File_Array), CLng(533)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Sample_Annot From Raw MS Data")

'' Function: Get_Sample_Name_Array_Compound_Table_Test
''
'' Description:
''
'' Function used to test if the function
'' Load_Raw_Data.Get_Sample_Name_Array is working
''
'' Test files are
''
''  - CompoundTableForm.csv
''
'' Ensure that an empty string array is declared
'' MS_File_Array()
''
'' Function will assert if Sample_Name_Array has 50 elements
'' Function will assert if MS_File_Array has 50 elements
''
Public Sub Get_Sample_Name_Array_Compound_Table_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim RawDataFilesArray() As String
    Dim RawDataFile As Variant
    Dim FileThere As Boolean
    
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "CompoundTableForm.csv"
    
    RawDataFilesArray = Split(RawDataFiles, ";")
    
    'Check if the data file exists
    For Each RawDataFile In RawDataFilesArray
        FileThere = (Dir(RawDataFile) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & RawDataFile & " cannot be found."
            End
        End If
    Next RawDataFile
    
    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray, MS_File_Array)
    
    Assert.AreEqual Utilities.Get_String_Array_Len(Sample_Name_Array_from_Raw_Data), CLng(50)
    Assert.AreEqual Utilities.Get_String_Array_Len(MS_File_Array), CLng(50)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Sample_Annot From Raw MS Data")

'' Function: Get_Sample_Name_Array_SciEx_Data_Test
''
'' Description:
''
'' Function used to test if the function
'' Load_Raw_Data.Get_Sample_Name_Array is working
''
'' Test files are
''
''  - SciExTestData.txt
''
'' Ensure that an empty string array is declared
'' MS_File_Array()
''
'' Function will assert if Sample_Name_Array has 61 elements
'' Function will assert if MS_File_Array has 61 elements
''
Public Sub Get_Sample_Name_Array_SciEx_Data_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim RawDataFilesArray() As String
    Dim RawDataFile As Variant
    Dim FileThere As Boolean
    
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "SciExTestData.txt"
    
    RawDataFilesArray = Split(RawDataFiles, ";")
    
    'Check if the data file exists
    For Each RawDataFile In RawDataFilesArray
        FileThere = (Dir(RawDataFile) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & RawDataFile & " cannot be found."
            End
        End If
    Next RawDataFile
    
    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray, MS_File_Array)
    
    Assert.AreEqual Utilities.Get_String_Array_Len(Sample_Name_Array_from_Raw_Data), CLng(61)
    Assert.AreEqual Utilities.Get_String_Array_Len(MS_File_Array), CLng(61)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Sample_Annot From Raw MS Data")

'' Function: Get_Sample_Name_Array_Multiple_Data_Test
''
'' Description:
''
'' Function used to test if the function
'' Load_Raw_Data.Get_Sample_Name_Array is working
''
'' Test files are
''
''  - MultipleDataTest1.csv
''  - MultipleDataTest2.csv
''  - SciExTestData.txt
''
'' Ensure that an empty string array is declared
'' MS_File_Array()
''
'' Function will assert if Sample_Name_Array has 664 elements
'' Function will assert if MS_File_Array has 664 elements
''
Public Sub Get_Sample_Name_Array_Multiple_Data_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim RawDataFilesArray() As String
    Dim RawDataFile As Variant
    Dim FileThere As Boolean
    
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "MultipleDataTest1.csv" & ";" & _
                   TestFolder & "MultipleDataTest2.csv" & ";" & _
                   TestFolder & "SciExTestData.txt"
    RawDataFilesArray = Split(RawDataFiles, ";")
    
    'Check if the data file exists
    For Each RawDataFile In RawDataFilesArray
        FileThere = (Dir(RawDataFile) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & RawDataFile & " cannot be found."
            End
        End If
    Next RawDataFile
    
    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray, MS_File_Array)
    
    Assert.AreEqual Utilities.Get_String_Array_Len(Sample_Name_Array_from_Raw_Data), CLng(664)
    Assert.AreEqual Utilities.Get_String_Array_Len(MS_File_Array), CLng(664)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Sample_Annot From Raw MS Data")

'' Function: Get_Sample_Name_Array_Invalid_Data_Test
''
'' Description:
''
'' Function used to test if the function
'' Load_Raw_Data.Get_Sample_Name_Array is working
''
'' Test files are
''
''  - InvalidDataTest1.csv
''
'' Ensure that an empty string array is declared
'' MS_File_Array()
''
'' Function will assert if Sample_Name_Array has 0 elements
'' Function will assert if MS_File_Array has 0 elements
''
Public Sub Get_Sample_Name_Array_Invalid_Data_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim RawDataFilesArray() As String
    Dim RawDataFile As Variant
    Dim FileThere As Boolean
    
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "InvalidDataTest1.csv"
    
    RawDataFilesArray = Split(RawDataFiles, ";")
    
    'Check if the data file exists
    For Each RawDataFile In RawDataFilesArray
        FileThere = (Dir(RawDataFile) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & RawDataFile & " cannot be found."
            End
        End If
    Next RawDataFile
    
    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray, MS_File_Array)
    
    Assert.AreEqual Utilities.Get_String_Array_Len(Sample_Name_Array_from_Raw_Data), CLng(0)
    Assert.AreEqual Utilities.Get_String_Array_Len(MS_File_Array), CLng(0)

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


