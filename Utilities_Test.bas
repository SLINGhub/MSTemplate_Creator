Attribute VB_Name = "Utilities_Test"
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
Public Sub Get_Header_Col_Position_Test()
    On Error GoTo TestFail
    
    Sheets("Transition_Name_Annot").Activate
    Assert.AreEqual Utilities.Get_Header_Col_Position("Transition_Name", 1), 1
    Assert.AreEqual Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 1), 2
    
    Sheets("ISTD_Annot").Activate
    Assert.AreEqual Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 2), 1
    Assert.AreEqual Utilities.Get_Header_Col_Position("ISTD_Conc_[ng/mL]", 3), 2
    Assert.AreEqual Utilities.Get_Header_Col_Position("ISTD_[MW]", 3), 3
    Assert.AreEqual Utilities.Get_Header_Col_Position("ISTD_Conc_[nM]", 3), 5
    
    Sheets("Sample_Annot").Activate
    Assert.AreEqual Utilities.Get_Header_Col_Position("Sample_Name", 1), 3
    Assert.AreEqual Utilities.Get_Header_Col_Position("Sample_Type", 1), 4

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub LastUsedRowNumber_Test()
    On Error GoTo TestFail
    
    Sheets("Lists").Activate
    'Debug.Print Utilities.LastUsedRowNumber
    Assert.AreEqual Utilities.LastUsedRowNumber, CLng(16)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub ConvertToLetterTest() 'TODO Rename test
    On Error GoTo TestFail
    
    Assert.AreEqual Utilities.ConvertToLetter(1), "A"
    Assert.AreEqual Utilities.ConvertToLetter(26), "Z"
    Assert.AreEqual Utilities.ConvertToLetter(27), "AA"
    Assert.AreEqual Utilities.ConvertToLetter(52), "AZ"
    Assert.AreEqual Utilities.ConvertToLetter(53), "BA"
    Assert.AreEqual Utilities.ConvertToLetter(520), "SZ"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub StringArrayLenTest()
    On Error GoTo TestFail
    
    Dim TestArray
    Dim EmptyArray
    
    TestArray = Array("11_PQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_PQC_2.d", _
                    "PQC_PC-PE-SM_01.d", "PQC1_LPC-LPE-PG-PI-PS_01.d", "PQC1_02.d", _
                    "11_BQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_BQC_2.d", "018_BQC_PQC01")
    EmptyArray = Array()
    
    Assert.AreEqual Utilities.StringArrayLen(TestArray), 8
    Assert.AreEqual Utilities.StringArrayLen(EmptyArray), 0

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub WhereInArrayTest()
    On Error GoTo TestFail
    
    Dim TestArray
    Dim CorrectPositions(0 To 3) As String
    CorrectPositions(0) = "0"
    CorrectPositions(1) = "2"
    CorrectPositions(2) = "4"
    CorrectPositions(3) = "5"
    Dim Positions() As String
    
    'Ensure that it works and gives the right position
    TestArray = Array("Here", "11_PQC-2.d", "Here", "No", "Here", "Here")
    Positions = Utilities.WhereInArray("Here", TestArray)
    Assert.SequenceEquals Positions, CorrectPositions
    
    'Ensure it gives an empty string array when there is no match
    Positions = Utilities.WhereInArray("Her", TestArray)
    Assert.AreEqual Utilities.StringArrayLen(Positions), 0
    
    'Ensure it gives an empty string array when test array is empty
    TestArray = Array()
    Positions = Utilities.WhereInArray("Her", TestArray)
    Assert.AreEqual Utilities.StringArrayLen(Positions), 0

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub IsInArrayTest()
    On Error GoTo TestFail
    
    Dim TestArray
    TestArray = Array("Here", "11_PQC-2.d", "Here", "No", "Here", "Here")
    
    Assert.IsTrue Utilities.IsInArray("Here", TestArray)
    Assert.IsTrue Utilities.IsInArray("11_PQC-2.d", TestArray)
    Assert.IsFalse Utilities.IsInArray("NotHere", TestArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub QuickSortTest()
    On Error GoTo TestFail
    
    Dim TestArray
    Dim SortedArray
    TestArray = Array("SM C36:2", "lipid", "Cer d18:1/C16:0")
    SortedArray = Array("Cer d18:1/C16:0", "SM C36:2", "lipid")
    Utilities.QuickSort ThisArray:=TestArray
    
    Assert.SequenceEquals TestArray, SortedArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Load_Columns_From_Excel_Test()
    On Error GoTo TestFail
    
    Sheets("Lists").Activate
    
    Dim CorrectArray(0 To 6) As String
    CorrectArray(0) = "[mg]_dry_weight"
    CorrectArray(1) = "[mg]_fresh_weight"
    CorrectArray(2) = "[ng]_DNA"
    CorrectArray(3) = "[ug]_protein/[ml]"
    CorrectArray(4) = "[ug]_total_protein"
    CorrectArray(5) = "[uL]"
    CorrectArray(6) = "[uM]"
    
    Dim Sample_Amount_Unit_Array() As String
    Sample_Amount_Unit_Array = Utilities.Load_Columns_From_Excel("Sample_Amount_Unit", HeaderRowNumber:=1, _
                                                    DataStartRowNumber:=2, MessageBoxRequired:=True, _
                                                    RemoveBlksAndReplicates:=True, IgnoreEmptyArray:=False)
    Assert.SequenceEquals Sample_Amount_Unit_Array, CorrectArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

