Attribute VB_Name = "Sample_Annot_Test"
Attribute VB_Description = "Test units for the functions in Sample_Annot Module."
Option Explicit
Option Private Module
'@ModuleDescription("Test units for the functions in Sample_Annot Module.")
'@IgnoreModule IntegerDataType

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
'@Description("Function used to test if the function Sample_Annot.Autofill_Column_By_QC_Sample_Type is working.")

'' Function: Autofill_Column_By_QC_Sample_Type_Test
'' --- Code
''  Public Sub Autofill_Column_By_QC_Sample_Type_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Annot.Autofill_Column_By_QC_Sample_Type is working
''
'' Ensure that
'' A string array called QC_Sample_Type_Array is created
'' {"SPL", "BQC", "TQC", "TQC", "BQC"} and is loaded to the
'' Sample_Type column of the Sample_Annot sheet.
''
'' Sample_Annot.Autofill_Column_By_QC_Sample_Type will be used to
'' to fill BQC's Sample Amount to 10 and all sample types' ISTD
'' mixture volume to 190
''
'' These two columns will be loaded to an array Sample_Amount_Array
'' and ISTD_Mixture_Volume_uL_Array
''
'' Check if the Sample_Amount_Array is
'' {"", "10", "", "", "10"}
''
'' Check if the ISTD_Mixture_Volume_uL_Array is
'' {"190", "190", "190", "190", "190"}
''
Public Sub Autofill_Column_By_QC_Sample_Type_Test()
Attribute Autofill_Column_By_QC_Sample_Type_Test.VB_Description = "Function used to test if the function Sample_Annot.Autofill_Column_By_QC_Sample_Type is working."
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
                            
    Sample_Annot.Autofill_Column_By_QC_Sample_Type Sample_Type:="BQC", _
                                                   Header_Name:="Sample_Amount", _
                                                   Autofill_Value:="10"
                                      
    Sample_Annot.Autofill_Column_By_QC_Sample_Type Sample_Type:="All Sample Types", _
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

'@TestMethod("Get Sample Annot Information")
'@Description("Function used to test if the function Sample_Annot.Get_Sample_Name_Array_From_Annot_File is working.")

'' Function: Get_Sample_Name_Array_From_Annot_File_Test
'' --- Code
''  Public Sub Get_Sample_Name_Array_From_Annot_File_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Annot.Get_Sample_Name_Array_From_Annot_File is working
''
'' Test files are
''
''  - Sample_Annotation_Example.csv
''
'' Ensure that
'' Load_Sample_Annot_Raw.Sample_Name_Text.Text = "Sample"
'' Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True
''
'' Output of Sample_Annot.Get_Sample_Name_Array_From_Annot_File(SampleAnnotFile)
'' is a string array of 55 elements
''
'' The first three elements are "1_untreated", "1_untreated" and "1_3h"
''
Public Sub Get_Sample_Name_Array_From_Annot_File_Test()
Attribute Get_Sample_Name_Array_From_Annot_File_Test.VB_Description = "Function used to test if the function Sample_Annot.Get_Sample_Name_Array_From_Annot_File is working."
    On Error GoTo TestFail
    
    Load_Sample_Annot_Raw.Sample_Name_Text.Text = "Sample"
    Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True
     
    'Load the Sample_Name from Sample Annotation

    Dim SampleAnnotFile As String
    Dim TestFolder As String
    Dim Sample_Name_Array_from_Sample_Annot() As String

    TestFolder = ThisWorkbook.Path & "\Testdata\"
    SampleAnnotFile = TestFolder & "Sample_Annotation_Example.csv"

    Sample_Name_Array_from_Sample_Annot = Sample_Annot.Get_Sample_Name_Array_From_Annot_File(SampleAnnotFile)
    
    'Debug.Print Utilities.Get_String_Array_Len(Sample_Name_Array_from_Sample_Annot)
    'Debug.Print Sample_Name_Array_from_Sample_Annot(0)
    'Debug.Print Sample_Name_Array_from_Sample_Annot(1)
    'Debug.Print Sample_Name_Array_from_Sample_Annot(2)
    
    Assert.AreEqual Utilities.Get_String_Array_Len(Sample_Name_Array_from_Sample_Annot), CLng(55)
    Assert.AreEqual Sample_Name_Array_from_Sample_Annot(0), "1_untreated"
    Assert.AreEqual Sample_Name_Array_from_Sample_Annot(1), "1_untreated"
    Assert.AreEqual Sample_Name_Array_from_Sample_Annot(2), "1_3h"
    
    GoTo TestExit
    
TestExit:
    Load_Sample_Annot_Raw.Sample_Name_Text.Text = vbNullString
    Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True
    Application.EnableEvents = True
    Exit Sub
TestFail:
    Application.EnableEvents = True
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Sample Annot Information")
'@Description("Function used to test if the function Sample_Annot.Get_Sample_Column_Name_Position_From_Annot_File is working.")

'' Function: Get_Sample_Column_Name_Position_From_Annot_File_Test
'' --- Code
''  Public Sub Get_Sample_Column_Name_Position_From_Annot_File_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Annot.Get_Sample_Column_Name_Position_From_Annot_File is working
''
'' Ensure that
'' A string array called first_line is created
'' {"Sample", "ID", "TimePoint", "Cell Number", "ISTD Volume"}
''
'' Load_Sample_Annot_Raw.Sample_Name_Text.Text = "Sample"
'' Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True
''
'' Sample_Annot.Get_Sample_Column_Name_Position_From_Annot_File(first_line)
'' should return 0 as "Sample" is is the first element of first_line
''
Public Sub Get_Sample_Column_Name_Position_From_Annot_File_Test()
Attribute Get_Sample_Column_Name_Position_From_Annot_File_Test.VB_Description = "Function used to test if the function Sample_Annot.Get_Sample_Column_Name_Position_From_Annot_File is working."
    On Error GoTo TestFail
    
    'Load the Sample_Name from Sample Annotation

    Dim first_line(4) As String

    first_line(0) = "Sample"
    first_line(1) = "ID"
    first_line(2) = "TimePoint"
    first_line(3) = "Cell Number"
    first_line(4) = "ISTD Volume"

    Load_Sample_Annot_Raw.Sample_Name_Text.Text = "Sample"
    Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True

    'Get the data starting row and the right column for the Sample Name
    Dim Sample_Column_Name_pos As Integer
    Sample_Column_Name_pos = Sample_Annot.Get_Sample_Column_Name_Position_From_Annot_File(first_line)
    
    Assert.AreEqual Sample_Column_Name_pos, 0
    
    GoTo TestExit
    
TestExit:
    Load_Sample_Annot_Raw.Sample_Name_Text.Text = vbNullString
    Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True
    Application.EnableEvents = True
    Exit Sub
TestFail:
    Application.EnableEvents = True
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get Sample Annot Information")
'@Description("Function used to test if the function Sample_Annot.Get_Sample_Annot_Starting_Line_From_Annot_File is working.")

'' Function: Get_Sample_Annot_Starting_Line_From_Annot_File_Test
'' --- Code
''  Public Sub Get_Sample_Annot_Starting_Line_From_Annot_File_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Annot.Get_Sample_Annot_Starting_Line_From_Annot_File is working
''
'' If Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = False
'' Sample_Annot.Get_Sample_Annot_Starting_Line_From_Annot_File
'' should return 0
''
'' If Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True
'' Sample_Annot.Get_Sample_Annot_Starting_Line_From_Annot_File
'' should return 1
''
Public Sub Get_Sample_Annot_Starting_Line_From_Annot_File_Test()
Attribute Get_Sample_Annot_Starting_Line_From_Annot_File_Test.VB_Description = "Function used to test if the function Sample_Annot.Get_Sample_Annot_Starting_Line_From_Annot_File is working."
    On Error GoTo TestFail
    
    'Get the data starting row

    Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True

    Dim data_starting_line As Integer
    data_starting_line = Sample_Annot.Get_Sample_Annot_Starting_Line_From_Annot_File
    
    Assert.AreEqual data_starting_line, 1
    
    Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = False
    data_starting_line = Sample_Annot.Get_Sample_Annot_Starting_Line_From_Annot_File
    
    Assert.AreEqual data_starting_line, 0
    
    GoTo TestExit
    
TestExit:
    Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True
    Application.EnableEvents = True
    Exit Sub
TestFail:
    Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True
    Application.EnableEvents = True
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Load Sample Annot Information")
'@Description("Function used to test if the function Sample_Annot.Load_Sample_Info_To_Excel is working.")

'' Function: Load_Sample_Info_To_Excel_Test
'' --- Code
''  Public Sub Load_Sample_Info_To_Excel_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Annot.Load_Sample_Info_To_Excel is working
''
'' Test files are
''
''  - Sample_Annotation_Example.csv
''
'' Ensure that
'' Load_Sample_Annot_Raw.Sample_Amount_Text.Text = "Cell Number"
'' Load_Sample_Annot_Raw.ISTD_Mixture_Volume_Text.Text = "ISTD Volume"
''
'' A string array called MatchingIndexArray
'' {"0", "1", "2", "", "3"}
''
'' Sample_Annot.Load_Sample_Info_To_Excel will load values on the Sample_Amount and
'' ISTD_Mixture_Volume_[uL] column of the Sample_Annot sheet
''
'' Function will assert if the Sample_Amount column array is
'' {"10", "10", "10", "", "10"}.
''
'' Function will assert if the ISTD_Mixture_Volume_[uL] column array is
'' {"1", "2", "4", "", "5"}.
''
Public Sub Load_Sample_Info_To_Excel_Test()
Attribute Load_Sample_Info_To_Excel_Test.VB_Description = "Function used to test if the function Sample_Annot.Load_Sample_Info_To_Excel is working."
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
    Dim SampleAnnotFile As String
    Dim TestFolder As String

    TestFolder = ThisWorkbook.Path & "\Testdata\"
    SampleAnnotFile = TestFolder & "Sample_Annotation_Example.csv"

    Load_Sample_Annot_Raw.Sample_Amount_Text.Text = "Cell Number"
    Load_Sample_Annot_Raw.ISTD_Mixture_Volume_Text.Text = "ISTD Volume"

    Dim MatchingIndexArray(4) As String
    MatchingIndexArray(0) = "0"
    MatchingIndexArray(1) = "1"
    MatchingIndexArray(2) = "2"
    MatchingIndexArray(3) = vbNullString
    MatchingIndexArray(4) = "3"
    
    Sample_Annot.Load_Sample_Info_To_Excel xFileName:=SampleAnnotFile, _
                                           MatchingIndexArray:=MatchingIndexArray
                                             
    Sample_Amount_Array = Utilities.Load_Columns_From_Excel("Sample_Amount", HeaderRowNumber:=1, _
                                                            DataStartRowNumber:=2, MessageBoxRequired:=False, _
                                                            RemoveBlksAndReplicates:=False, _
                                                            IgnoreHiddenRows:=False, IgnoreEmptyArray:=False)
                                                            
    ISTD_Mixture_Volume_uL_Array = Utilities.Load_Columns_From_Excel("ISTD_Mixture_Volume_[uL]", HeaderRowNumber:=1, _
                                                                     DataStartRowNumber:=2, MessageBoxRequired:=False, _
                                                                     RemoveBlksAndReplicates:=False, _
                                                                     IgnoreHiddenRows:=False, IgnoreEmptyArray:=False)

    ' Rows except fourth must have its Sample_Amount of "10"
    Assert.AreEqual Sample_Amount_Array(0), "10"
    Assert.AreEqual Sample_Amount_Array(1), "10"
    Assert.AreEqual Sample_Amount_Array(2), "10"
    Assert.AreEqual Sample_Amount_Array(3), vbNullString
    Assert.AreEqual Sample_Amount_Array(4), "10"

    ' Rows except fourth must have its corresponding values
    Assert.AreEqual ISTD_Mixture_Volume_uL_Array(0), "1"
    Assert.AreEqual ISTD_Mixture_Volume_uL_Array(1), "2"
    Assert.AreEqual ISTD_Mixture_Volume_uL_Array(2), "4"
    Assert.AreEqual ISTD_Mixture_Volume_uL_Array(3), vbNullString
    Assert.AreEqual ISTD_Mixture_Volume_uL_Array(4), "5"
    
    GoTo TestExit

TestExit:
    Utilities.Clear_Columns HeaderToClear:="Sample_Amount", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Utilities.Clear_Columns HeaderToClear:="ISTD_Mixture_Volume_[uL]", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2
    Load_Sample_Annot_Raw.Sample_Amount_Text.Text = vbNullString
    Load_Sample_Annot_Raw.ISTD_Mixture_Volume_Text.Text = vbNullString
    Application.EnableEvents = True
    Exit Sub
TestFail:
    Load_Sample_Annot_Raw.Sample_Amount_Text.Text = vbNullString
    Load_Sample_Annot_Raw.ISTD_Mixture_Volume_Text.Text = vbNullString
    Application.EnableEvents = True
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
