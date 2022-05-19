Attribute VB_Name = "Transition_Name_Annot_Test"
Attribute VB_Description = "Test units for the functions in Transition_Name_Annot Module."
Option Explicit
Option Private Module
'@ModuleDescription("Test units for the functions in Transition_Name_Annot Module.")

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

''@TestInitialize
'Public Sub TestInitialize()
'    'this method runs before every test in the module.
'End Sub
'
''@TestCleanup
'Public Sub TestCleanup()
'    'this method runs after every test in the module.
'End Sub

'@TestMethod("Verify Transition Name ISTD")
'@Description("Function used to test if the function Transition_Annot.Verify_ISTD is working.")

'' Function: Verify_ISTD_Test
'' --- Code
''  Public Sub Verify_ISTD_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Transition_Annot.Verify_ISTD is working.
''
'' Test data are
''
''  - A string array Transition_Array = {"SM 44:1", "SM 44:2", "SM 46:2", "SM 46:3"}
''  - A string array ISTD_Array = {"SM 44:1", "SM 44:2", "", "SM 44:1"}
''
'' Function will assert if cell colours are correct
'' on the Transition_Name and Transition_Name_ISTD column
''
''  - Cells with "SM 44:0" as Transition_Name and "SM 44:1" as Transition_Name_ISTD should both be green
''  - Cells with "SM 46:3" as Transition_Name and "SM 44:1" as Transition_Name_ISTD should both be green
''  - Cells with "SM 44:1" as Transition_Name and "SM 44:2" as Transition_Name_ISTD should both be white
''  - Cells with "SM 46:2" as Transition_Name should be green
''  - Cells with "" as Transition_Name_ISTD should be yellow
''
Public Sub Verify_ISTD_Test()
Attribute Verify_ISTD_Test.VB_Description = "Function used to test if the function Transition_Annot.Verify_ISTD is working."
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
    
    Dim ISTD_Array(3) As String
    ISTD_Array(0) = "SM 44:1"
    ISTD_Array(1) = "SM 44:2"
    ISTD_Array(2) = vbNullString
    ISTD_Array(3) = "SM 44:1"
    
    Utilities.Load_To_Excel Data_Array:=Transition_Array, _
                            HeaderName:="Transition_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
                            
    Utilities.Load_To_Excel Data_Array:=ISTD_Array, _
                            HeaderName:="Transition_Name_ISTD", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
                 
    Transition_Name_Annot.Verify_ISTD Transition_Array:=Transition_Array, _
                                      ISTD_Array:=ISTD_Array, _
                                      MessageBoxRequired:=False, _
                                      Testing:=True
    
    ' Cells with "SM 44:0" as Transition_Name and "SM 44:1" as Transition_Name_ISTD should both be green
    Assert.AreEqual CLng(Transition_Name_Annot_Worksheet.Cells.Item(2, 1).Interior.Color), RGB(204, 255, 204)
    Assert.AreEqual CLng(Transition_Name_Annot_Worksheet.Cells.Item(2, 2).Interior.Color), RGB(204, 255, 204)
    
    ' Cells with "SM 46:3" as Transition_Name and "SM 44:1" as Transition_Name_ISTD should both be green
    Assert.AreEqual CLng(Transition_Name_Annot_Worksheet.Cells.Item(5, 1).Interior.Color), RGB(204, 255, 204)
    Assert.AreEqual CLng(Transition_Name_Annot_Worksheet.Cells.Item(5, 2).Interior.Color), RGB(204, 255, 204)
                                      
    ' Cells with "SM 44:1" as Transition_Name and "SM 44:2" as Transition_Name_ISTD should both be white
    Assert.AreEqual CLng(Transition_Name_Annot_Worksheet.Cells.Item(3, 1).Interior.Color), RGB(255, 255, 255)
    Assert.AreEqual CLng(Transition_Name_Annot_Worksheet.Cells.Item(3, 2).Interior.Color), RGB(255, 255, 255)
                                      
    ' Cells with "SM 46:2" as Transition_Name should be green
    ' Cells with "" as Transition_Name_ISTD should be yellow
    Assert.AreEqual CLng(Transition_Name_Annot_Worksheet.Cells.Item(4, 1).Interior.Color), RGB(204, 255, 204)
    Assert.AreEqual CLng(Transition_Name_Annot_Worksheet.Cells.Item(4, 2).Interior.Color), RGB(255, 255, 153)

    GoTo TestExit
TestExit:
    Utilities.Clear_Columns HeaderToClear:="Transition_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Transition_Name_ISTD", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Application.EnableEvents = True
    Exit Sub
TestFail:
    Application.EnableEvents = True
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
