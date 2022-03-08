Attribute VB_Name = "ISTD_Annot_Test"
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

''@TestInitialize
'Public Sub TestInitialize()
'    'this method runs before every test in the module.
'End Sub
'
''@TestCleanup
'Public Sub TestCleanup()
'    'this method runs after every test in the module.
'End Sub

'' Function: Convert_Conc_nM_Array_Test
'' --- Code
''  Public Sub Convert_Conc_nM_Array_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' ISTD_Annot.Convert_Conc_nM_Array is working
''
'' Test data is a string array of {"1000", "2000", "3000"}.
''
'' Function will assert if the output string array is
'' {"1", "2", "3"}.
''
'@TestMethod("Convert ISTD Concentration")
Public Sub Convert_Conc_nM_Array_Test()
    On Error GoTo TestFail

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
   
    GoTo TestExit

TestExit:
    Utilities.Clear_Columns HeaderToClear:="ISTD_Conc_[nM]", _
                            HeaderRowNumber:=3, _
                            DataStartRowNumber:=4
    Application.EnableEvents = True
    Exit Sub
TestFail:
    Application.EnableEvents = True
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: Get_ISTD_Conc_nM_Array_Test
'' --- Code
''  Public Sub Get_ISTD_Conc_nM_Array_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' ISTD_Annot.Get_ISTD_Conc_nM_Array_Test is working
''
'' Test data are
''
''  - A string array Transition_Name_ISTD = {"ISTD1", "ISTD2"}
''  - A string array ISTD_Conc_ng_per_mL = {"1"}
''  - A string array ISTD_Conc_MW = {"2"}
''  - A string array ISTD_Conc_nM = {"", "100"}
''
''
'' Function will assert if
''
''  - The Custom_Unit is "[uM] or [pmol/uL]"
''  - The output ISTD_Conc_nM_result string array = {"500", "100"}
''  - The output ISTD_Custom_Unit string array = {"0.5", "0.1"}
''  - Cells occupied with numbers should be green
''
'@TestMethod("Get ISTD Conc nM Array")
Public Sub Get_ISTD_Conc_nM_Array_Test()
    On Error GoTo TestFail
    
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False

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
    
    Dim Transition_Name_ISTD(1) As String
    Dim ISTD_Conc_ng_per_mL(0) As String
    Dim ISTD_Conc_MW(0) As String
    Dim ISTD_Conc_nM(1) As String
    Dim ISTD_Conc_nM_Results() As String
    Dim ISTD_Custom_Unit_ColNumber As Long
    Dim Custom_Unit As String
    Dim ISTD_Custom_Unit() As String
    
    Transition_Name_ISTD(0) = "ISTD1"
    Transition_Name_ISTD(1) = "ISTD2"
    ISTD_Conc_ng_per_mL(0) = "1"
    ISTD_Conc_MW(0) = "2"
    ISTD_Conc_nM(0) = vbNullString
    ISTD_Conc_nM(1) = "100"
    ISTD_Custom_Unit_ColNumber = Utilities.Get_Header_Col_Position("Custom_Unit", 2)
    Custom_Unit = ISTD_Annot_Worksheet.Cells.Item(3, ISTD_Custom_Unit_ColNumber)
    
    Utilities.Load_To_Excel Data_Array:=Transition_Name_ISTD, _
                            HeaderName:="Transition_Name_ISTD", _
                            HeaderRowNumber:=2, _
                            DataStartRowNumber:=4, _
                            MessageBoxRequired:=False
    
    Utilities.Load_To_Excel Data_Array:=ISTD_Conc_ng_per_mL, _
                            HeaderName:="ISTD_Conc_[ng/mL]", _
                            HeaderRowNumber:=3, _
                            DataStartRowNumber:=4, _
                            MessageBoxRequired:=False
    
    Utilities.Load_To_Excel Data_Array:=ISTD_Conc_MW, _
                            HeaderName:="ISTD_[MW]", _
                            HeaderRowNumber:=3, _
                            DataStartRowNumber:=4, _
                            MessageBoxRequired:=False
   
    Utilities.Load_To_Excel Data_Array:=ISTD_Conc_nM, _
                            HeaderName:="ISTD_Conc_[nM]", _
                            HeaderRowNumber:=3, _
                            DataStartRowNumber:=4, _
                            MessageBoxRequired:=False
   
    ISTD_Conc_nM_Results = ISTD_Annot.Get_ISTD_Conc_nM_Array(ColourCellRequired:=True)
   
    Utilities.Load_To_Excel Data_Array:=ISTD_Conc_nM_Results, _
                            HeaderName:="ISTD_Conc_[nM]", _
                            HeaderRowNumber:=3, _
                            DataStartRowNumber:=4, _
                            MessageBoxRequired:=False
                            
    ISTD_Custom_Unit = ISTD_Annot.Convert_Conc_nM_Array(Custom_Unit)
    
    Utilities.Load_To_Excel Data_Array:=ISTD_Custom_Unit, _
                            HeaderName:="Custom_Unit", _
                            HeaderRowNumber:=2, _
                            DataStartRowNumber:=4, _
                            MessageBoxRequired:=False
    
    Assert.AreEqual Custom_Unit, "[uM] or [pmol/uL]"
    Assert.AreEqual ISTD_Conc_nM_Results(0), "500"
    Assert.AreEqual ISTD_Conc_nM_Results(1), "100"
    Assert.AreEqual ISTD_Custom_Unit(0), "0.5"
    Assert.AreEqual ISTD_Custom_Unit(1), "0.1"
    
    ' Cells should both be green
    Assert.AreEqual CLng(ISTD_Annot_Worksheet.Cells.Item(4, 2).Interior.Color), RGB(204, 255, 204)
    Assert.AreEqual CLng(ISTD_Annot_Worksheet.Cells.Item(4, 3).Interior.Color), RGB(204, 255, 204)
    Assert.AreEqual CLng(ISTD_Annot_Worksheet.Cells.Item(4, 5).Interior.Color), RGB(204, 255, 204)
    Assert.AreEqual CLng(ISTD_Annot_Worksheet.Cells.Item(4, 6).Interior.Color), RGB(204, 255, 204)
    Assert.AreEqual CLng(ISTD_Annot_Worksheet.Cells.Item(5, 5).Interior.Color), RGB(204, 255, 204)
    Assert.AreEqual CLng(ISTD_Annot_Worksheet.Cells.Item(5, 6).Interior.Color), RGB(204, 255, 204)
      
    GoTo TestExit

TestExit:
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
                            
    ISTD_Annot_Worksheet.Cells.Item(4, 2).Interior.Color = xlNone
    ISTD_Annot_Worksheet.Cells.Item(4, 3).Interior.Color = xlNone
    ISTD_Annot_Worksheet.Cells.Item(4, 5).Interior.Color = xlNone
    ISTD_Annot_Worksheet.Cells.Item(4, 6).Interior.Color = xlNone
    ISTD_Annot_Worksheet.Cells.Item(5, 5).Interior.Color = xlNone
    ISTD_Annot_Worksheet.Cells.Item(5, 6).Interior.Color = xlNone
    
    Application.EnableEvents = True
    Exit Sub
    
TestFail:
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
                            
    ISTD_Annot_Worksheet.Cells.Item(4, 2).Interior.Color = xlNone
    ISTD_Annot_Worksheet.Cells.Item(4, 3).Interior.Color = xlNone
    ISTD_Annot_Worksheet.Cells.Item(4, 5).Interior.Color = xlNone
    ISTD_Annot_Worksheet.Cells.Item(4, 6).Interior.Color = xlNone
    ISTD_Annot_Worksheet.Cells.Item(5, 5).Interior.Color = xlNone
    ISTD_Annot_Worksheet.Cells.Item(5, 6).Interior.Color = xlNone
    
    Application.EnableEvents = True
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

