Attribute VB_Name = "Sample_Annot_Test"
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

'@TestMethod("Autofill Columns By QC Sample Type")
Public Sub Autofill_Column_By_QC_Sample_Type_Test()
    On Error GoTo TestFail
    
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
    
    Dim Sample_Amount_Array() As String
    Dim ISTD_Mixture_Volume_uL_Array() As String
    
    Dim QC_Sample_Type_Array(4) As String
    QC_Sample_Type_Array(0) = "SPL"
    QC_Sample_Type_Array(1) = "BQC"
    QC_Sample_Type_Array(2) = "TQC"
    QC_Sample_Type_Array(3) = "TQC"
    QC_Sample_Type_Array(4) = "BQC"

    Utilities.Load_To_Excel Data_Array:=QC_Sample_Type_Array, _
                            HeaderName:="Sample_Type", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
                            
    Autofill_Column_By_QC_Sample_Type Sample_Type:="BQC", _
                                      Header_Name:="Sample_Amount", _
                                      Autofill_Value:="10"
                                      
    Autofill_Column_By_QC_Sample_Type Sample_Type:="All Sample Types", _
                                      Header_Name:="ISTD_Mixture_Volume_[uL]", _
                                      Autofill_Value:="190"
                                      
    Sample_Amount_Array = Utilities.Load_Columns_From_Excel("Sample_Amount", HeaderRowNumber:=1, _
                                                            DataStartRowNumber:=2, MessageBoxRequired:=False, _
                                                            RemoveBlksAndReplicates:=False, _
                                                            IgnoreHiddenRows:=False, IgnoreEmptyArray:=False)
                                                            
    ISTD_Mixture_Volume_uL_Array = Utilities.Load_Columns_From_Excel("ISTD_Mixture_Volume_[uL]", HeaderRowNumber:=1, _
                                                                     DataStartRowNumber:=2, MessageBoxRequired:=False, _
                                                                     RemoveBlksAndReplicates:=False, _
                                                                     IgnoreHiddenRows:=False, IgnoreEmptyArray:=False)
                                          
    ' Rows whose Sample_Type is "BQC" must have its Sample_Amount of "10"
    Assert.AreEqual Sample_Amount_Array(0), vbNullString
    Assert.AreEqual Sample_Amount_Array(1), "10"
    Assert.AreEqual Sample_Amount_Array(2), vbNullString
    Assert.AreEqual Sample_Amount_Array(3), vbNullString
    Assert.AreEqual Sample_Amount_Array(4), "10"

    'All rows must have its ISTD_Mixture_Volume_[uL] value of "190"
    Assert.AreEqual ISTD_Mixture_Volume_uL_Array(0), "190"
    Assert.AreEqual ISTD_Mixture_Volume_uL_Array(1), "190"
    Assert.AreEqual ISTD_Mixture_Volume_uL_Array(2), "190"
    Assert.AreEqual ISTD_Mixture_Volume_uL_Array(3), "190"
    Assert.AreEqual ISTD_Mixture_Volume_uL_Array(4), "190"
    
    GoTo TestExit
                                      
TestExit:
    Utilities.Clear_Columns HeaderToClear:="Sample_Type", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="Sample_Amount", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="ISTD_Mixture_Volume_[uL]", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Application.EnableEvents = True
    Exit Sub
TestFail:
    Application.EnableEvents = True
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
