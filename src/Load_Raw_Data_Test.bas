Attribute VB_Name = "Load_Raw_Data_Test"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Get Transition_Annot From Raw MS Data")
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
    FileThere = (Dir(RawDataFiles) > "")
        If FileThere = False Then
            MsgBox "File name " & RawDataFiles & " cannot be found."
            End
        End If
    
    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=RawDataFiles)
    
    Assert.AreEqual Utilities.StringArrayLen(Transition_Array), 30

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Transition_Annot From Raw MS Data")
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
    FileThere = (Dir(RawDataFiles) > "")
        If FileThere = False Then
            MsgBox "File name " & RawDataFiles & " cannot be found."
            End
        End If

    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=RawDataFiles)
    
    Assert.AreEqual Utilities.StringArrayLen(Transition_Array), 122

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Transition_Annot From Raw MS Data")
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
    FileThere = (Dir(RawDataFiles) > "")
        If FileThere = False Then
            MsgBox "File name " & RawDataFiles & " cannot be found."
            End
        End If
  
    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=RawDataFiles)
    
    Assert.AreEqual Utilities.StringArrayLen(Transition_Array), 224

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Transition_Annot From Raw MS Data")
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
        FileThere = (Dir(RawDataFile) > "")
        If FileThere = False Then
            MsgBox "File name " & RawDataFile & " cannot be found."
            End
        End If
    Next RawDataFile
    
    JoinedFiles = Join(RawDataFilesArray, ";")
    
    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=JoinedFiles)
    
    Assert.AreEqual Utilities.StringArrayLen(Transition_Array), 653

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Transition_Annot From Raw MS Data")
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
    FileThere = (Dir(RawDataFiles) > "")
        If FileThere = False Then
            MsgBox "File name " & RawDataFiles & " cannot be found."
            End
        End If
    
    'Load the transition names into an array
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=RawDataFiles)
    
    Assert.AreEqual Utilities.StringArrayLen(Transition_Array), 0

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Sample_Annot From Raw MS Data")
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
        FileThere = (Dir(RawDataFile) > "")
        If FileThere = False Then
            MsgBox "File name " & RawDataFile & " cannot be found."
            End
        End If
    Next RawDataFile

    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray, _
                                                                          MS_File_Array)

    Assert.AreEqual Utilities.StringArrayLen(Sample_Name_Array_from_Raw_Data), 533
    Assert.AreEqual Utilities.StringArrayLen(MS_File_Array), 533

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Get Sample_Annot From Raw MS Data")
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
        FileThere = (Dir(RawDataFile) > "")
        If FileThere = False Then
            MsgBox "File name " & RawDataFile & " cannot be found."
            End
        End If
    Next RawDataFile
    
    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray, MS_File_Array)
    
    Assert.AreEqual Utilities.StringArrayLen(Sample_Name_Array_from_Raw_Data), 50
    Assert.AreEqual Utilities.StringArrayLen(MS_File_Array), 50

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Sample_Annot From Raw MS Data")
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
        FileThere = (Dir(RawDataFile) > "")
        If FileThere = False Then
            MsgBox "File name " & RawDataFile & " cannot be found."
            End
        End If
    Next RawDataFile
    
    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray, MS_File_Array)
    
    Assert.AreEqual Utilities.StringArrayLen(Sample_Name_Array_from_Raw_Data), 61
    Assert.AreEqual Utilities.StringArrayLen(MS_File_Array), 61

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Sample_Annot From Raw MS Data")
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
        FileThere = (Dir(RawDataFile) > "")
        If FileThere = False Then
            MsgBox "File name " & RawDataFile & " cannot be found."
            End
        End If
    Next RawDataFile
    
    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray, MS_File_Array)
    
    Assert.AreEqual Utilities.StringArrayLen(Sample_Name_Array_from_Raw_Data), 664
    Assert.AreEqual Utilities.StringArrayLen(MS_File_Array), 664

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Sample_Annot From Raw MS Data")
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
        FileThere = (Dir(RawDataFile) > "")
        If FileThere = False Then
            MsgBox "File name " & RawDataFile & " cannot be found."
            End
        End If
    Next RawDataFile
    
    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray, MS_File_Array)
    
    Assert.AreEqual Utilities.StringArrayLen(Sample_Name_Array_from_Raw_Data), 0
    Assert.AreEqual Utilities.StringArrayLen(MS_File_Array), 0

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


