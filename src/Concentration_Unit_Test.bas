Attribute VB_Name = "Concentration_Unit_Test"
Option Explicit
Option Private Module

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
'Public Sub TestInitialize()
'    'this method runs before every test in the module.
'End Sub
'
''@TestCleanup
'Public Sub TestCleanup()
'    'this method runs after every test in the module.
'End Sub

'@TestMethod("Get Concentration Unit")

'' Function: Get_Mol_From_Custom_ISTD_Concentration_Unit_Test
''
'' Description:
''
'' Function used to test if the function
'' Concentration_Unit.Get_Mol_From_Custom_ISTD_Concentration_Unit is working
''
'' Test string is "[uM] or [pmol/uL]". The function should return
'' "pmol"
''
Public Sub Get_Mol_From_Custom_ISTD_Concentration_Unit_Test()
    On Error GoTo TestFail
    
    Dim Custom_ISTD_Concentration_Unit As String
    Dim Output_Custom_Unit As String
    
    Custom_ISTD_Concentration_Unit = "[uM] or [pmol/uL]"
    Output_Custom_Unit = Concentration_Unit.Get_Mol_From_Custom_ISTD_Concentration_Unit(Custom_ISTD_Concentration_Unit)
    
    Assert.AreEqual Output_Custom_Unit, "pmol"
    
    GoTo TestExit

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
