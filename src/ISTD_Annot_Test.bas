Attribute VB_Name = "ISTD_Annot_Test"
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

'@TestMethod("Convert ISTD Concentration")
Public Sub Convert_Conc_nM_Array_Test()

    ' Get the ISTD_Annot worksheet from the active workbook
    ' The ISTDAnnot is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim ISTD_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "ISTDAnnot") = False Then
        MsgBox ("Sheet ISTD_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set ISTD_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "ISTDAnnot")
      
    ISTD_Annot_Worksheet.Activate

   Dim ISTD_Conc_nM(2) As String
   Dim ISTD_Custom_Unit() As String
   Dim Custom_Unit As String

   ISTD_Conc_nM(0) = "1000"
   ISTD_Conc_nM(1) = "2000"
   ISTD_Conc_nM(2) = "3000"
   
   Custom_Unit = "[uM] or [pmol/uL]"
   
   Utilities.Load_To_Excel Data_Array:=ISTD_Conc_nM, _
                           HeaderName:="ISTD_Conc_[nM]", _
                           HeaderRowNumber:=3, _
                           DataStartRowNumber:=4, _
                           MessageBoxRequired:=False

   ISTD_Custom_Unit = ISTD_Annot.Convert_Conc_nM_Array(Custom_Unit)

   Assert.AreEqual ISTD_Custom_Unit(0), "1"
   Assert.AreEqual ISTD_Custom_Unit(1), "2"
   Assert.AreEqual ISTD_Custom_Unit(2), "3"

TestExit:
    Utilities.Clear_Columns HeaderToClear:="ISTD_Conc_[nM]", _
                            HeaderRowNumber:=3, _
                            DataStartRowNumber:=4
    Exit Sub
TestFail:
    Utilities.Clear_Columns HeaderToClear:="ISTD_Conc_[nM]", _
                            HeaderRowNumber:=3, _
                            DataStartRowNumber:=4
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
