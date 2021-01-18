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

'@TestMethod
Public Sub Get_Transition_Array_Wide_Table_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim xFileNames As Variant
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
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
    
    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array(xFileNames:=xFileNames)
    
    Assert.AreEqual Utilities.StringArrayLen(Transition_Array), 30

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Get_Transition_Array_Compound_Table_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim xFileNames As Variant
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    
    'Check if the data file exists
    xFileNames = Array(TestFolder & "CompoundTableForm.csv")
    
    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > "")
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array(xFileNames:=xFileNames)
    
    Assert.AreEqual Utilities.StringArrayLen(Transition_Array), 122

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Get_Transition_Array_SciEx_Data_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim xFileNames As Variant
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    
    'Check if the data file exists
    xFileNames = Array(TestFolder & "SciExTestData.txt")
    
    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > "")
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array(xFileNames:=xFileNames)
    
    Assert.AreEqual Utilities.StringArrayLen(Transition_Array), 224

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Get_Transition_Array_Multiple_Data_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim xFileNames As Variant
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    
    'Check if the data file exists
    xFileNames = Array(TestFolder & "sPerfect_Index_AllLipids_raw.csv", _
                       TestFolder & "Autophagy_Data_Nov 2017.csv", _
                       TestFolder & "SciExTestData.txt")
    
    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > "")
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Load the transition names and load it to excel
    Transition_Array = Load_Raw_Data.Get_Transition_Array(xFileNames:=xFileNames)
    
    Assert.AreEqual Utilities.StringArrayLen(Transition_Array), 653

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Get_Transition_Array_Invalid_Data_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim xFileNames As Variant
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    
    'Check if the data file exists
    xFileNames = Array(TestFolder & "Autophagy_Samples_List.csv")
    
    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > "")
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Load the transition names into an array
    Transition_Array = Load_Raw_Data.Get_Transition_Array(xFileNames:=xFileNames)
    
    Assert.AreEqual Utilities.StringArrayLen(Transition_Array), 0

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Get_Sample_Name_Array_Wide_Table_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim xFileNames() As String
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "sPerfect_Index_AllLipids_raw.csv"
    
    xFileNames = Split(RawDataFiles, ";")
    
    'Check if the data file exists
    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > "")
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(xFileNames, MS_File_Array)
    
    Assert.AreEqual Utilities.StringArrayLen(Sample_Name_Array_from_Raw_Data), 533
    Assert.AreEqual Utilities.StringArrayLen(MS_File_Array), 533

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Get_Sample_Name_Array_Compound_Table_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim xFileNames() As String
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "CompoundTableForm.csv"
    
    xFileNames = Split(RawDataFiles, ";")
    
    'Check if the data file exists
    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > "")
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(xFileNames, MS_File_Array)
    
    Assert.AreEqual Utilities.StringArrayLen(Sample_Name_Array_from_Raw_Data), 50
    Assert.AreEqual Utilities.StringArrayLen(MS_File_Array), 50

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Get_Sample_Name_Array_SciEx_Data_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim xFileNames() As String
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "SciExTestData.txt"
    
    xFileNames = Split(RawDataFiles, ";")
    
    'Check if the data file exists
    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > "")
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(xFileNames, MS_File_Array)
    
    Assert.AreEqual Utilities.StringArrayLen(Sample_Name_Array_from_Raw_Data), 61
    Assert.AreEqual Utilities.StringArrayLen(MS_File_Array), 61

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Get_Sample_Name_Array_Multiple_Data_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim xFileNames() As String
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "Autophagy_Data_Nov 2017.csv" & ";" & _
                   TestFolder & "sPerfect_Index_AllLipids_raw.csv" & ";" & _
                   TestFolder & "SciExTestData.txt"
    xFileNames = Split(RawDataFiles, ";")
    
    'Check if the data file exists
    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > "")
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(xFileNames, MS_File_Array)
    
    Assert.AreEqual Utilities.StringArrayLen(Sample_Name_Array_from_Raw_Data), 664
    Assert.AreEqual Utilities.StringArrayLen(MS_File_Array), 664

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Get_Sample_Name_Array_Invalid_Data_Test()
    On Error GoTo TestFail
    
    Dim TestFolder As String
    Dim RawDataFiles As String
    Dim xFileNames() As String
    Dim xFileName As Variant
    Dim FileThere As Boolean
    
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    RawDataFiles = TestFolder & "Autophagy_Samples_List.csv"
    
    xFileNames = Split(RawDataFiles, ";")
    
    'Check if the data file exists
    For Each xFileName In xFileNames
        FileThere = (Dir(xFileName) > "")
        If FileThere = False Then
            MsgBox "File name " & xFileName & " cannot be found."
            End
        End If
    Next xFileName
    
    'Load the sample name and datafile name into the two arrays
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(xFileNames, MS_File_Array)
    
    Assert.AreEqual Utilities.StringArrayLen(Sample_Name_Array_from_Raw_Data), 0
    Assert.AreEqual Utilities.StringArrayLen(MS_File_Array), 0

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


