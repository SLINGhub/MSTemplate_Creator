Attribute VB_Name = "Load_Tidy_Data_Test"
Attribute VB_Description = "Test units for the functions in Load Tidy Data Module."
Option Explicit
Option Private Module
'@ModuleDescription("Test units for the functions in Load Tidy Data Module.")

'@TestModule
'@Folder("Tests")

Private Assert As Object
'Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    'Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub

''@TestInitialize
'Private Sub TestInitialize()
'    'This method runs before every test in the module..
'End Sub

''@TestCleanup
'Private Sub TestCleanup()
'    'this method runs after every test in the module.
'End Sub

'@TestMethod("Get Transition_Annot From Tidy Data")
'@Description("Function used to test if the function Load_Tidy_Data.Get_Transition_Array_Tidy is working.")

'' Function: Get_Transition_Array_Tidy_Data_Row_Test
'' --- Code
''  Public Sub Get_Transition_Array_Tidy_Data_Row_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Load_Tidy_Data.Get_Transition_Array_Tidy is working
''
'' Test files are
''
''  - TidyTransitionRow.csv
''
'' Function will assert if Transition_Array has 7 elements
''
Public Sub Get_Transition_Array_Tidy_Data_Row_Test()
Attribute Get_Transition_Array_Tidy_Data_Row_Test.VB_Description = "Function used to test if the function Load_Tidy_Data.Get_Transition_Array_Tidy is working."
    On Error GoTo TestFail
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim TidyDataRowFiles As String
    Dim FileThere As Boolean
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    TidyDataRowFiles = TestFolder & "TidyTransitionRow.csv"
    
    'Check if the data file exists
    FileThere = (Dir(TidyDataRowFiles) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & TidyDataRowFiles & " cannot be found."
            End
        End If
    
    'Test creating a new transition annotation from tidy data file with transitons as row observations
    Transition_Array = Load_Tidy_Data.Get_Transition_Array_Tidy(TidyDataFiles:=TidyDataRowFiles, _
                                                                DataFileType:="csv", _
                                                                TransitionProperty:="Read as row observations", _
                                                                StartingRowNum:=2, _
                                                                StartingColumnNum:=1)
    Assert.AreEqual Utilities.Get_String_Array_Len(Transition_Array), CLng(7)
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Transition_Annot From Tidy Data")
'@Description("Function used to test if the function Load_Tidy_Data.Get_Transition_Array_Tidy is working.")

'' Function: Get_Transition_Array_Tidy_Data_Column_Test
'' --- Code
''  Public Sub Get_Transition_Array_Tidy_Data_Column_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Load_Tidy_Data.Get_Transition_Array_Tidy is working
''
'' Test files are
''
''  - TidyTransitionColumn.csv
''
'' Function will assert if Transition_Array has 8 elements
''
Public Sub Get_Transition_Array_Tidy_Data_Column_Test()
Attribute Get_Transition_Array_Tidy_Data_Column_Test.VB_Description = "Function used to test if the function Load_Tidy_Data.Get_Transition_Array_Tidy is working."
    On Error GoTo TestFail
    Dim TestFolder As String
    Dim Transition_Array() As String
    Dim TidyDataColumnFiles As String
    Dim FileThere As Boolean
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    TidyDataColumnFiles = TestFolder & "TidyTransitionColumn.csv"
    
    'Check if the data file exists
    FileThere = (Dir(TidyDataColumnFiles) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & TidyDataColumnFiles & " cannot be found."
            End
        End If
        
    'Test creating a new transition annotation from tidy data file with transitons as column variables
    Transition_Array = Load_Tidy_Data.Get_Transition_Array_Tidy(TidyDataFiles:=TidyDataColumnFiles, _
                                                                DataFileType:="csv", _
                                                                TransitionProperty:="Read as column variables", _
                                                                StartingRowNum:=1, _
                                                                StartingColumnNum:=2)
                                                                              
    Assert.AreEqual Utilities.Get_String_Array_Len(Transition_Array), CLng(8)
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Sample_Annot From Tidy Data")
'@Description("Function used to test if the function Load_Tidy_Data.Get_Sample_Name_Array_Tidy is working.")

'' Function: Get_Sample_Array_Tidy_Data_Row_Test
'' --- Code
''  Public Sub Get_Sample_Array_Tidy_Data_Row_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Load_Tidy_Data.Get_Sample_Name_Array_Tidy is working
''
'' Test files are
''
''  - TidySampleRow.csv
''
'' Function will assert if Sample_Name_Array_from_Tidy_Data has 7 elements
'' Function will assert if MS_File_Array has 7 elements
''
Public Sub Get_Sample_Array_Tidy_Data_Row_Test()
Attribute Get_Sample_Array_Tidy_Data_Row_Test.VB_Description = "Function used to test if the function Load_Tidy_Data.Get_Sample_Name_Array_Tidy is working."
    On Error GoTo TestFail
    Dim TestFolder As String
    'Dim Transition_Array() As String
    Dim TidyDataRowFiles As String
    Dim JoinedFiles As String
    Dim FileThere As Boolean
    Dim TidyDataFilesArray() As String
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Tidy_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    TidyDataRowFiles = TestFolder & "TidySampleRow.csv"
    
    'Check if the data file exists
    FileThere = (Dir(TidyDataRowFiles) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & TidyDataRowFiles & " cannot be found."
            End
        End If
    
    JoinedFiles = Join(Array(TidyDataRowFiles), ";")
    TidyDataFilesArray = Split(JoinedFiles, ";")
        
    Sample_Name_Array_from_Tidy_Data = Load_Tidy_Data.Get_Sample_Name_Array_Tidy(TidyDataFilesArray(), _
                                                                                 MS_File_Array, _
                                                                                 DataFileType:="csv", _
                                                                                 SampleProperty:="Read as row observations", _
                                                                                 StartingRowNum:=2, _
                                                                                 StartingColumnNum:=1)
                                                                                 
    Assert.AreEqual Utilities.Get_String_Array_Len(Sample_Name_Array_from_Tidy_Data), CLng(7)
    Assert.AreEqual Utilities.Get_String_Array_Len(MS_File_Array), CLng(7)
        
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Sample_Annot From Tidy Data")
'@Description("Function used to test if the function Load_Tidy_Data.Get_Sample_Name_Array_Tidy is working.")

'' Function: Get_Sample_Array_Tidy_Data_Column_Test
'' --- Code
''  Public Sub Get_Sample_Array_Tidy_Data_Column_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Load_Tidy_Data.Get_Sample_Name_Array_Tidy is working
''
'' Test files are
''
''  - TidySampleColumn.csv
''
'' Function will assert if Sample_Name_Array_from_Tidy_Data has 7 elements
'' Function will assert if MS_File_Array has 7 elements
''
Public Sub Get_Sample_Array_Tidy_Data_Column_Test()
Attribute Get_Sample_Array_Tidy_Data_Column_Test.VB_Description = "Function used to test if the function Load_Tidy_Data.Get_Sample_Name_Array_Tidy is working."
    On Error GoTo TestFail
    Dim TestFolder As String
    'Dim Transition_Array() As String
    Dim TidyDataColumnFiles As String
    Dim JoinedFiles As String
    Dim FileThere As Boolean
    Dim TidyDataFilesArray() As String
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Tidy_Data() As String
    
    'Indicate path to the test data folder
    TestFolder = ThisWorkbook.Path & "\Testdata\"
    TidyDataColumnFiles = TestFolder & "TidySampleColumn.csv"
    
    'Check if the data file exists
    FileThere = (Dir(TidyDataColumnFiles) > vbNullString)
        If FileThere = False Then
            MsgBox "File name " & TidyDataColumnFiles & " cannot be found."
            End
        End If
    
    JoinedFiles = Join(Array(TidyDataColumnFiles), ";")
    TidyDataFilesArray = Split(JoinedFiles, ";")
        
    Sample_Name_Array_from_Tidy_Data = Load_Tidy_Data.Get_Sample_Name_Array_Tidy(TidyDataFilesArray:=TidyDataFilesArray, _
                                                                                 MS_File_Array:=MS_File_Array, _
                                                                                 DataFileType:="csv", _
                                                                                 SampleProperty:="Read as column variables", _
                                                                                 StartingRowNum:=1, _
                                                                                 StartingColumnNum:=2)
                                                                                 
    Assert.AreEqual Utilities.Get_String_Array_Len(Sample_Name_Array_from_Tidy_Data), CLng(7)
    Assert.AreEqual Utilities.Get_String_Array_Len(MS_File_Array), CLng(7)
       
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
