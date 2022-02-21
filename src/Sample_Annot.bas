Attribute VB_Name = "Sample_Annot"
Option Explicit
'@Folder("Sample_Annot Functions")
'@IgnoreModule IntegerDataType


'' Function: Autofill_Column_By_QC_Sample_Type
'' --- Code
''  Public Sub Autofill_Column_By_QC_Sample_Type(ByRef Sample_Type As String, _
''                                               ByRef Header_Name As String, _
''                                               ByVal Autofill_Value As String)
'' ---
''
'' Description:
''
'' Fill in the column indicated by Header_Name with the value
'' indicated by Autofill_Value on rows whose sample type matches
'' Sample_Type
''
'' Parameters:
''
''    Sample_Type As String - Name of the QC Sample Type to match like
''                            SPL or BQC. To choose all sample types,
''                            the input is All Sample Types
''
''    Header_Name As String - Name of the column to fill in. Currently,
''                            we only use "Sample Amount", "Sample_Amount_Unit"
''                            and "ISTD_Mixture_Volume_[uL]"
''
''    Autofill_Value As String - Value to fill in at column "Header_Name on
''                               rows whose sample type matches Sample_Type
''
'' Examples:
''
'' --- Code
''    ' Get the Sample_Annot worksheet from the active workbook
''    ' The SampleAnnotSheet is a code name
''    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
''    Dim Sample_Annot_Worksheet As Worksheet
''
''    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "SampleAnnotSheet") = False Then
''        MsgBox ("Sheet Sample_Annot is missing")
''        Application.EnableEvents = True
''        Exit Sub
''    End If
''
''    Set Sample_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "SampleAnnotSheet")
''
''    Sample_Annot_Worksheet.Activate
''
''    Dim Sample_Amount_Array() As String
''    Dim ISTD_Mixture_Volume_uL_Array() As String
''
''    Dim QC_Sample_Type_Array(4) As String
''    QC_Sample_Type_Array(0) = "SPL"
''    QC_Sample_Type_Array(1) = "BQC"
''    QC_Sample_Type_Array(2) = "TQC"
''    QC_Sample_Type_Array(3) = "TQC"
''    QC_Sample_Type_Array(4) = "BQC"
''
''    Utilities.Load_To_Excel Data_Array:=QC_Sample_Type_Array, _
''                            HeaderName:="Sample_Type", _
''                            HeaderRowNumber:=1, _
''                            DataStartRowNumber:=2, _
''                            MessageBoxRequired:=False
''
''    Autofill_Column_By_QC_Sample_Type Sample_Type:="BQC", _
''                                      Header_Name:="Sample_Amount", _
''                                      Autofill_Value:="10"
''
''    Autofill_Column_By_QC_Sample_Type Sample_Type:="All Sample Types", _
''                                      Header_Name:="ISTD_Mixture_Volume_[uL]", _
''                                      Autofill_Value:="190
'' ---
Public Sub Autofill_Column_By_QC_Sample_Type(ByRef Sample_Type As String, _
                                             ByRef Header_Name As String, _
                                             ByVal Autofill_Value As String)

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
   
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    'We don't want excel to monitor the sheet when runnning this code
    Application.EnableEvents = False
    
    Dim SampleTypeArray() As String
    Dim AutofillValueArray() As String
    Dim TotalRows As Long
    Dim SampleTypeArrayIndex As Long
    
    'Check if the column Sample_Type exists
    Dim SampleType_pos As Integer
    SampleType_pos = Utilities.Get_Header_Col_Position("Sample_Type", HeaderRowNumber:=1)
    
    'Check if the column given in Header_Name exists
    Dim AutofillValue_pos As Integer
    AutofillValue_pos = Utilities.Get_Header_Col_Position(Header_Name, HeaderRowNumber:=1)
    
    'Filter Rows by input Sample_Type
    If Sample_Type <> "All Sample Types" Then
        ActiveSheet.Range("A1").AutoFilter Field:=SampleType_pos, _
                                           Criteria1:=Sample_Type, _
                                           VisibleDropDown:=True
    End If
    
    'Load the Sample_Type column content from Sample_Annot
    SampleTypeArray = Utilities.Load_Columns_From_Excel("Sample_Type", _
                                                        HeaderRowNumber:=1, _
                                                        DataStartRowNumber:=2, _
                                                        MessageBoxRequired:=False, _
                                                        RemoveBlksAndReplicates:=False, _
                                                        IgnoreHiddenRows:=True, _
                                                        IgnoreEmptyArray:=True)
                                                              
    'Check if SampleTypeArray has any elements
    'If not there is nothing to fill
    If Len(Join(SampleTypeArray, vbNullString)) = 0 Then
        'To ensure that Filters does not affect the assignment
        Utilities.RemoveFilterSettings
    
        'Resume monitoring of sheet
        Application.EnableEvents = True
        
        'Give a message that the Sample Type cannot be found
        MsgBox "No samples have Sample Type : " & Sample_Type
        
        Exit Sub
    End If
    
    'Load the Sample_Amount_Unit column content from Sample_Annot
    AutofillValueArray = Utilities.Load_Columns_From_Excel(Header_Name, _
                                                           HeaderRowNumber:=1, _
                                                           DataStartRowNumber:=2, _
                                                           MessageBoxRequired:=False, _
                                                           RemoveBlksAndReplicates:=False, _
                                                           IgnoreHiddenRows:=True, _
                                                           IgnoreEmptyArray:=True)
                                                              
                                                              
    'Check if SampleAmountUnitArray has any elements
    'If yes, give an overwrite warning
    If Len(Join(AutofillValueArray, vbNullString)) > 0 Then
        Utilities.OverwriteHeader HeaderName:=Header_Name, _
                                  HeaderRowNumber:=1, _
                                  DataStartRowNumber:=2, _
                                  ClearContent:=False
        'Filter Rows by input Sample_Type
        If Sample_Type <> "All Sample Types" Then
            Sample_Annot_Worksheet.Range("A1").AutoFilter Field:=SampleType_pos, _
                                                          Criteria1:=Sample_Type, _
                                                          VisibleDropDown:=True
        End If
    End If
     
    'Find the total number of rows and resize the array accordingly
    TotalRows = Sample_Annot_Worksheet.Cells.Item(Sample_Annot_Worksheet.Rows.Count, ConvertToLetter(SampleType_pos)).End(xlUp).Row
    
    'Assign the relevant sample amount unit for that sample type
    If TotalRows > 1 Then
        If Sample_Type = "All Sample Types" Then
            For SampleTypeArrayIndex = 2 To TotalRows
                Sample_Annot_Worksheet.Cells.Item(SampleTypeArrayIndex, AutofillValue_pos).Value = Autofill_Value
            Next SampleTypeArrayIndex
        Else
            For SampleTypeArrayIndex = 2 To TotalRows
                If Sample_Annot_Worksheet.Cells.Item(SampleTypeArrayIndex, SampleType_pos).Value = Sample_Type Then
                    Sample_Annot_Worksheet.Cells.Item(SampleTypeArrayIndex, AutofillValue_pos).Value = Autofill_Value
                End If
            Next SampleTypeArrayIndex
            
            MsgBox "Autofill" & vbNewLine & "Sample_Type : " & Sample_Type & vbNewLine & Header_Name & " : " & Autofill_Value & "."
            
            'To ensure that Filters does not affect the assignment
            Utilities.RemoveFilterSettings
        End If
    End If
    
   
    'Resume monitoring of sheet
    Application.EnableEvents = True
    
                                          
End Sub

'' Function: Create_New_Sample_Annot_Tidy
'' --- Code
''  Public Sub Create_New_Sample_Annot_Tidy(ByVal TidyDataFiles As String, _
''                                          ByRef DataFileType As String, _
''                                          ByRef SampleProperty As String, _
''                                          ByRef StartingRowNum As Integer, _
''                                          ByRef StartingColumnNum As Integer)
'' ---
''
'' Description:
''
'' Create Sample Annotation from an input data file in tabular form and output
'' them into the Sample_Annot sheet. The columns filled will be
''
'' - Data_File_Name
'' - Merge_Status
'' - Sample_Name
'' - Sample_Type
''
'' Parameters:
''
''    TidyDataFiles As String - File path to a tabular/tidy data file.
''                              If multiple files are required, the different
''                              file path must be separated by ";"
''                              Eg. {FilePath 1};{FilePath 2}
''
''    DataFileType As String - File type of the input tabular/tidy data file
''
''    TransitionProperty As String - Choose "Read as column variables" if transition names are the column name.
''                                   Choose "Read as row observations" if transition names are row entries.
''
''    StartingRowNum As Integer - Starting row number to read the tabular data
''
''    StartingColumnNum As Integer - Starting column number to read the tabular data
''
'' Examples:
''
'' --- Code
''    Dim TestFolder As String
''    Dim TidyDataColumnFiles As String
''
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    TidyDataColumnFiles = TestFolder & "TidySampleColumn.csv"
''
''    Sample_Annot.Create_New_Sample_Annot_Tidy TidyDataFiles:=TidyDataColumnFiles, _
''                                              DataFileType:="csv", _
''                                              SampleProperty:="Read as column variables", _
''                                              StartingRowNum:=1, _
''                                              StartingColumnNum:=2
'' ---
Public Sub Create_New_Sample_Annot_Tidy(ByVal TidyDataFiles As String, _
                                        ByRef DataFileType As String, _
                                        ByRef SampleProperty As String, _
                                        ByRef StartingRowNum As Integer, _
                                        ByRef StartingColumnNum As Integer)
    
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
    
    'File are taken from userfrom Load_Sample_Annot_Tidy
    'Hence they must exists and joined together by ;
    Dim TidyDataFilesArray() As String
    TidyDataFilesArray = Split(TidyDataFiles, ";")

    'Load the Sample_Name from Raw Data
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Tidy_Data() As String
    'Dim TotalRows As Long
    Sample_Name_Array_from_Tidy_Data = Load_Tidy_Data.Get_Sample_Name_Array_Tidy(TidyDataFilesArray, _
                                                                                 MS_File_Array, _
                                                                                 DataFileType, _
                                                                                 SampleProperty, _
                                                                                 StartingRowNum, _
                                                                                 StartingColumnNum)
                                                                                 
    Dim MergeStatus() As String
    Dim SampleType() As String
    Dim ArrayLength As Integer
    Dim Sample_Name_Array_Index As Integer
    ArrayLength = 0
    
    For Sample_Name_Array_Index = 0 To UBound(Sample_Name_Array_from_Tidy_Data) - LBound(Sample_Name_Array_from_Tidy_Data)
        ReDim Preserve MergeStatus(ArrayLength)
        ReDim Preserve SampleType(ArrayLength)
        
        MergeStatus(ArrayLength) = "Valid"
        SampleType(ArrayLength) = Sample_Type_Identifier.Get_QC_Sample_Type(Sample_Name_Array_from_Tidy_Data(Sample_Name_Array_Index))
        
        ArrayLength = ArrayLength + 1
    Next Sample_Name_Array_Index
    
    Dim HeaderNameArray(0 To 3) As String
    HeaderNameArray(0) = "Data_File_Name"
    HeaderNameArray(1) = "Merge_Status"
    HeaderNameArray(2) = "Sample_Name"
    HeaderNameArray(3) = "Sample_Type"
    
    Utilities.OverwriteSeveralHeaders HeaderNameArray:=HeaderNameArray, _
                                      HeaderRowNumber:=1, _
                                      DataStartRowNumber:=2
    
    Utilities.Load_To_Excel Data_Array:=MS_File_Array, _
                            HeaderName:="Data_File_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
                            
    Utilities.Load_To_Excel Data_Array:=MergeStatus, _
                            HeaderName:="Merge_Status", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
                            
    Utilities.Load_To_Excel Data_Array:=Sample_Name_Array_from_Tidy_Data, _
                            HeaderName:="Sample_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
                            
    Utilities.Load_To_Excel Data_Array:=SampleType, _
                            HeaderName:="Sample_Type", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
    
End Sub

'' Function: Create_New_Sample_Annot_Raw
'' --- Code
''  Public Sub Create_New_Sample_Annot_Raw(ByVal RawDataFiles As String)
'' ---
''
'' Description:
''
'' Create Sample Annotation from an input raw data file and output
'' them into the Sample_Annot sheet. The columns filled will be
''
'' - Data_File_Name
'' - Merge_Status
'' - Sample_Name
'' - Sample_Type
''
'' Parameters:
''
''    RawDataFiles As String - File path to a Raw Data (Agilent) File in csv.
''                             If multiple files are required, the different
''                             file path must be separated by ";"
''                             Eg. {FilePath 1};{FilePath 2}
''
'' Examples:
''
'' --- Code
''    Dim TestFolder As String
''    Dim RawDataFiles As String
''
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    RawDataFiles = TestFolder & "AgilentRawDataTest1.csv"
''
''    Sample_Annot.Create_New_Sample_Annot_Raw RawDataFiles:=RawDataFiles
'' ---
Public Sub Create_New_Sample_Annot_Raw(ByVal RawDataFiles As String)

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

    'File are taken from userfrom Load_Sample_Annot_Raw
    'Hence they must exists and joined together by ;
    Dim RawDataFilesArray() As String
    RawDataFilesArray = Split(RawDataFiles, ";")
    
    'Load the Sample_Name from Raw Data
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    'Dim TotalRows As Long
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray, MS_File_Array)
    
    'If there is no data loaded, stop the process
    If Utilities.StringArrayLen(Sample_Name_Array_from_Raw_Data) = CLng(0) Then
        End
    End If
    
    Dim MergeStatus() As String
    Dim SampleType() As String
    Dim ArrayLength As Long
    Dim Sample_Name_Array_Index As Integer
    ArrayLength = 0
    
    For Sample_Name_Array_Index = 0 To UBound(Sample_Name_Array_from_Raw_Data) - LBound(Sample_Name_Array_from_Raw_Data)
        ReDim Preserve MergeStatus(ArrayLength)
        ReDim Preserve SampleType(ArrayLength)
        
        MergeStatus(ArrayLength) = "Valid"
        SampleType(ArrayLength) = Sample_Type_Identifier.Get_QC_Sample_Type(Sample_Name_Array_from_Raw_Data(Sample_Name_Array_Index))
        
        ArrayLength = ArrayLength + 1
    Next Sample_Name_Array_Index
    
    Dim HeaderNameArray(0 To 3) As String
    HeaderNameArray(0) = "Data_File_Name"
    HeaderNameArray(1) = "Merge_Status"
    HeaderNameArray(2) = "Sample_Name"
    HeaderNameArray(3) = "Sample_Type"
    
    Utilities.OverwriteSeveralHeaders HeaderNameArray:=HeaderNameArray, _
                                      HeaderRowNumber:=1, _
                                      DataStartRowNumber:=2
    
    Utilities.Load_To_Excel Data_Array:=MS_File_Array, _
                            HeaderName:="Data_File_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
                            
    Utilities.Load_To_Excel Data_Array:=MergeStatus, _
                            HeaderName:="Merge_Status", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
                            
    Utilities.Load_To_Excel Data_Array:=Sample_Name_Array_from_Raw_Data, _
                            HeaderName:="Sample_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
                            
    Utilities.Load_To_Excel Data_Array:=SampleType, _
                            HeaderName:="Sample_Type", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
    
    
End Sub

'' Function: Merge_With_Sample_Annot
'' --- Code
''  Public Sub Merge_With_Sample_Annot(ByVal RawDataFiles As String, _
''                                     ByRef SampleAnnotFile As String)
'' ---
''
'' Description:
''
'' Merge the an input raw data file with a user input
'' sample annotation file. The merged data is then outputted
'' into the Sample_Annot sheet. The columns filled will be
''
'' - Data_File_Name
'' - Merge_Status
'' - Sample_Name
'' - Sample_Type
''
'' It is also possible to fill in the these columns if the user's
'' input sample annotation file provides these information.
''
'' - Sample_Amount
'' - ISTD_Mixture_Volume_[uL]
''
'' Should there be any issues with the merge, a message box will appear.
''
'' (see Sample_Annot_Merge_Issues.png)
''
'' Users are encouraged to see the Merge_Status column to see what is wrong.
''
'' (see Sample_Annot_Merge_Issue_Details.png)
''
'' Parameters:
''
''    RawDataFiles As String - File path to a Raw Data (Agilent) File in csv.
''                             If multiple files are required, the different
''                             file path must be separated by ";"
''                             Eg. {FilePath 1};{FilePath 2}
''
''    SampleAnnotFile As String - File path to a user input sample annotation
''                                file in csv.
''
''
'' --- Code
''    Dim TestFolder As String
''    Dim RawDataFiles As String
''    Dim SampleAnnotFile As String
''
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    RawDataFiles = TestFolder & "AgilentRawDataTest1.csv"
''    SampleAnnotFile = TestFolder & "Sample_Annotation_Example.csv"
''
''    Sample_Annot.Merge_With_Sample_Annot RawDataFiles:=RawDataFiles, _
''                                         SampleAnnotFile:=SampleAnnotFile
'' ---
Public Sub Merge_With_Sample_Annot(ByVal RawDataFiles As String, _
                                   ByRef SampleAnnotFile As String)

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
    
    'File are taken from userfrom Load_Sample_Annot_Raw
    'Hence they must exists and joined together by ;
    Dim RawDataFilesArray() As String
    RawDataFilesArray = Split(RawDataFiles, ";")
    
    'Load the Sample_Name from Raw Data
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    'Dim TotalRows As Long
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray, MS_File_Array)
    
    'Load the Sample_Name from Sample Annotation
    Dim Sample_Name_Array_from_Sample_Annot() As String
    Sample_Name_Array_from_Sample_Annot = Sample_Annot.Get_Sample_Name_Array_From_Annot_File(SampleAnnotFile)
    
    'If there is no data loaded, stop the process
    If Utilities.StringArrayLen(Sample_Name_Array_from_Raw_Data) = CLng(0) Then
        End
    End If
    
    'Match the Sample_Name from Raw Data to the one in Sample Annotation
    'Store merge status and matching index to the array
    Dim MatchingIndexArray() As String
    Dim MergeStatus() As String
    Dim SampleType() As String
    Dim MergeFailure As Boolean
    Dim ArrayLength As Integer
    Dim Sample_Name_Array_Index As Integer
    
    'MergeFailure = False
    ArrayLength = 0
    
    'For debugging
    'For i = 0 To UBound(Sample_Name_Array_from_Sample_Annot) - LBound(Sample_Name_Array_from_Sample_Annot)
    '    Debug.Print Sample_Name_Array_from_Sample_Annot(i)
    'Next i

    For Sample_Name_Array_Index = 0 To UBound(Sample_Name_Array_from_Raw_Data) - LBound(Sample_Name_Array_from_Raw_Data)
        'Get the positions of where the sample name of the raw data can be found in the sample annotation
        Dim Positions() As String
        Positions = WhereInArray(Sample_Name_Array_from_Raw_Data(Sample_Name_Array_Index), Sample_Name_Array_from_Sample_Annot)
        
        ReDim Preserve MergeStatus(ArrayLength)
        ReDim Preserve MatchingIndexArray(ArrayLength)
        ReDim Preserve SampleType(ArrayLength)
        
        'Display results if there is no match, unique match or duplicates
        If StringArrayLen(Positions) = 0 Then
            'Debug.Print "Empty"
            MergeStatus(ArrayLength) = "Missing in Annot File"
            MatchingIndexArray(ArrayLength) = vbNullString
            MergeFailure = True
        ElseIf StringArrayLen(Positions) > 1 Then
            'Debug.Print "Duplicate"
            Dim DuplicatePositionIndex As Integer
            For DuplicatePositionIndex = 0 To UBound(Positions) - LBound(Positions)
                If Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True Then
                    Positions(DuplicatePositionIndex) = CStr(CInt(Positions(DuplicatePositionIndex)) + 2)
                Else
                    Positions(DuplicatePositionIndex) = CStr(CInt(Positions(DuplicatePositionIndex)) + 1)
                End If
            Next DuplicatePositionIndex
            
            MergeStatus(ArrayLength) = "Duplicate at line " & Join(Positions, ", ")
            MatchingIndexArray(ArrayLength) = vbNullString
            MergeFailure = True
        Else
            'Debug.Print "Ok"
            MergeStatus(ArrayLength) = "Valid"
            MatchingIndexArray(ArrayLength) = Positions(0)
        End If
        
        SampleType(ArrayLength) = Sample_Type_Identifier.Get_QC_Sample_Type(Sample_Name_Array_from_Raw_Data(Sample_Name_Array_Index))
        
        ArrayLength = ArrayLength + 1
    Next Sample_Name_Array_Index
    
    Dim HeaderNameArray(0 To 3) As String
    HeaderNameArray(0) = "Data_File_Name"
    HeaderNameArray(1) = "Merge_Status"
    HeaderNameArray(2) = "Sample_Name"
    HeaderNameArray(3) = "Sample_Type"
    
    Utilities.OverwriteSeveralHeaders HeaderNameArray:=HeaderNameArray, _
                                      HeaderRowNumber:=1, _
                                      DataStartRowNumber:=2
      
    'Load Data into the excel sheet
    Sample_Annot.Load_Sample_Info_To_Excel xFileName:=SampleAnnotFile, _
                                           MatchingIndexArray:=MatchingIndexArray
                                           
    Utilities.Load_To_Excel Data_Array:=MS_File_Array, _
                            HeaderName:="Data_File_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
                            
    Utilities.Load_To_Excel Data_Array:=MergeStatus, _
                            HeaderName:="Merge_Status", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
                            
    Utilities.Load_To_Excel Data_Array:=Sample_Name_Array_from_Raw_Data, _
                            HeaderName:="Sample_Name", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
                            
    Utilities.Load_To_Excel Data_Array:=SampleType, _
                            HeaderName:="Sample_Type", _
                            HeaderRowNumber:=1, _
                            DataStartRowNumber:=2, _
                            MessageBoxRequired:=False
    
    'Notify the user if some rows in the raw data cannot merge with the sample annotation
    If MergeFailure = True Then
        MsgBox ("Some rows in the raw data are unable to merge with the annotation file.")
    End If
    
End Sub

'' Function: Get_Sample_Name_Array_From_Annot_File
'' --- Code
''  Public Function Get_Sample_Name_Array_From_Annot_File(ByRef xFileName As String) As String()
'' ---
''
'' Description:
''
'' Get an array of Sample Names from a given
'' sample annotation file in csv and tabular form.
''
'' If the sample names ends with .d, we will remove the .d.
''
'' Parameters:
''
''    xFileName As String - File path to a Sample Annotation File in csv.
''
'' Returns:
''    A string array of Sample Names.
''
'' Examples:
''
'' --- Code
''
''    Load_Sample_Annot_Raw.Sample_Name_Text.Text = "Sample"
''    Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True
''
''    'Load the Sample_Name from Sample Annotation
''
''    Dim SampleAnnotFile As String
''    Dim TestFolder As String
''    Dim Sample_Name_Array_from_Sample_Annot() As String
''
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    SampleAnnotFile = TestFolder & "Sample_Annotation_Example.csv"
''
''    Sample_Name_Array_from_Sample_Annot = Sample_Annot.Get_Sample_Name_Array_From_Annot_File(SampleAnnotFile)
'' ---
Public Function Get_Sample_Name_Array_From_Annot_File(ByRef xFileName As String) As String()

    'When no file is selected
    If TypeName(xFileName) = "Boolean" Then
        End
    End If
    On Error GoTo 0
    
    Dim Lines() As String
    Dim Delimiter As String
    Dim first_line() As String
    Dim LinesIndex As Integer
    Lines = Utilities.Read_File(xFileName)
    Delimiter = Utilities.Get_Delimiter(xFileName)
    
    'Get the first line from sample annot file
    first_line = Split(Lines(0), Delimiter)
    
    'Get the data starting row and the right column for the Sample Name
    Dim data_starting_line As Integer
    Dim Sample_Column_Name_pos As Integer
    data_starting_line = Sample_Annot.Get_Sample_Annot_Starting_Line_From_Annot_File
    Sample_Column_Name_pos = Sample_Annot.Get_Sample_Column_Name_Position_From_Annot_File(first_line)
    
    'Get the column name to extract the sample name from sample annotation file
    Dim Sample_Column_Name As String
    Sample_Column_Name = Load_Sample_Annot_Raw.Sample_Name_Text.Text

    'For the function output
    Dim Sample_Name_Array() As String
    Dim Sample_Data As String
    Dim ArrayLength As Integer
    ArrayLength = 0

    'Check that it is not empty, it should not be empty based on how we code the userform Load_Sample_Annot_Raw
    If Sample_Column_Name <> vbNullString Then
                        
        'Extract the sample name into the array
        For LinesIndex = data_starting_line To UBound(Lines) - 1
            'Get the data at the right pos and remove whitespaces
            Sample_Data = Trim$(Split(Lines(LinesIndex), Delimiter)(Sample_Column_Name_pos))
            ReDim Preserve Sample_Name_Array(ArrayLength)
            Sample_Name_Array(ArrayLength) = Sample_Data
            'Debug.Print Sample_Name_Array(ArrayLength)
            ArrayLength = ArrayLength + 1
        Next LinesIndex
            
        Sample_Name_Array = Utilities.Clear_DotD_In_Agilent_Data_File(Sample_Name_Array)
    End If
    Get_Sample_Name_Array_From_Annot_File = Sample_Name_Array
    
End Function

'' Function: Get_Sample_Column_Name_Position_From_Annot_File
'' --- Code
''  Public Function Get_Sample_Column_Name_Position_From_Annot_File(ByRef first_line() As String) As Integer
'' ---
''
'' Description:
''
'' Get the column position where the "Sample Name" column is located
'' as indicated in the Sample_Name text box.
''
'' In the case when the input sample annotation file has headers, it will
'' look like this.
''
'' (see Sample_Annot_Merge_Sample_Name_With_Headers.png)
''
'' In the case when the input sample annotation file has headers, it will
'' look like this.
''
'' (see Sample_Annot_Merge_Sample_Name_No_Headers.png)
''
'' Parameters:
''
''    first_line() As String - A string array of column names or the first row
''                             of the input sample annotation file.
''
'' Returns:
''    An integer indicating where "Sample Name" column is located
''    as indicated in the Sample_Name text box.
''
'' Examples:
''
'' --- Code
''    'Load the Sample_Name from Sample Annotation
''
''    Dim first_line(4) As String
''
''    first_line(0) = "Sample"
''    first_line(1) = "ID"
''    first_line(2) = "TimePoint"
''    first_line(3) = "Cell Number"
''    first_line(4) = "ISTD Volume"
''
''    Load_Sample_Annot_Raw.Sample_Name_Text.Text = "Sample"
''    Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True
''
''    'Get the data starting row and the right column for the Sample Name
''    Dim Sample_Column_Name_pos As Integer
''    Sample_Column_Name_pos = Sample_Annot.Get_Sample_Column_Name_Position_From_Annot_File(first_line)
'' ---
Public Function Get_Sample_Column_Name_Position_From_Annot_File(ByRef first_line() As String) As Integer

    'Get the column name to extract the sample name from sample annotation file
    Dim Sample_Column_Name As String
    Sample_Column_Name = Load_Sample_Annot_Raw.Sample_Name_Text.Text
    
    Dim Sample_Column_Name_pos As Integer
    
    'If not empty, get sample name position from sample annot file
    'Should not have an error as we have check that the columns present in the userform
    If Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True Then
        'If the sample annotation has headers, we take the first line
        Sample_Column_Name_pos = Application.Match(Sample_Column_Name, first_line, False) - 1
    Else
        'If the sample annotation has no headers,
        'Name will be "Column {Some Number}" use regular expression to get the number
        Dim column_number As String
        Dim regEx As RegExp
        Set regEx = New RegExp
        regEx.Pattern = "\d+"
        column_number = regEx.Execute(Sample_Column_Name).Item(0)
        Sample_Column_Name_pos = CInt(column_number) - 1
    End If
    
    Get_Sample_Column_Name_Position_From_Annot_File = Sample_Column_Name_pos

End Function

'' Function: Get_Sample_Annot_Starting_Line_From_Annot_File
'' --- Code
''  Public Function Get_Sample_Annot_Starting_Line_From_Annot_File() As Integer
'' ---
''
'' Description:
''
'' Get the starting line where the data is from the annotation file.
'' It should be 0 if the data has no headers and 1 if there is.
'' We assume that the column names is on the first line.
'' Basically, it just check if this check box is checked or not
''
'' (see Sample_Annot_Merge_Sample_Name_Headers_Checkbox.png)
''
'' Returns:
''    An integer indicating 0 if the data has no headers and 1 if there is.
''
'' Examples:
''
'' --- Code
''    'Get the data starting row
''
''    Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True
''
''    Dim data_starting_line As Integer
''    data_starting_line = Sample_Annot.Get_Sample_Annot_Starting_Line_From_Annot_File
'' ---
Public Function Get_Sample_Annot_Starting_Line_From_Annot_File() As Integer

    Dim data_starting_line As Integer
    
    'If not empty, get sample name position from sample annot file
    'Should not have an error as we have check that the columns present in the userform
    If Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True Then
        'If the sample annotation has headers, we set starting position as 1
        data_starting_line = 1
    Else
        'If the sample annotation has no headers, we set starting position as 0
        data_starting_line = 0
    End If
    
    Get_Sample_Annot_Starting_Line_From_Annot_File = data_starting_line

End Function

'' Function: Load_Sample_Info_To_Excel
'' --- Code
''  Public Sub Load_Sample_Info_To_Excel(ByRef xFileName As String, ByRef MatchingIndexArray() As String)
'' ---
''
'' Description:
''
'' Output the sample information (Sample_Amount and ISTD_Mixture_Volume_[uL])
'' found in the sample annotation file to the Sample_Annot sheet
''
'' Parameters:
''
''    xFileName As String - File path to a Sample Annotation File in csv.
''
''    MatchingIndexArray() As String - A string array of row positions indicating
''                                     which unique row in the sample annotation file
''                                     matches the sample name of the input raw data.
''                                     Sample information from these matching rows will
''                                     be loaded into the Sample_Annot sheet.
''                                     Rows that cannot be matched because it is not
''                                     indicated or is repeated more than once in the
''                                     sample annotation file are left as blank.
''
'' Examples:
''
'' --- Code
''    Dim SampleAnnotFile As String
''    Dim TestFolder As String
''
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    SampleAnnotFile = TestFolder & "Sample_Annotation_Example.csv"
''
''    Load_Sample_Annot_Raw.Sample_Amount_Text.Text = "Cell Number"
''    Load_Sample_Annot_Raw.ISTD_Mixture_Volume_Text.Text = "ISTD Volume"
''
''    Dim MatchingIndexArray(4) As String
''    MatchingIndexArray(0) = "0"
''    MatchingIndexArray(1) = "1"
''    MatchingIndexArray(2) = "2"
''    MatchingIndexArray(3) = vbNullString
''    MatchingIndexArray(4) = "3"
''
''    Sample_Annot.Load_Sample_Info_To_Excel xFileName:=SampleAnnotFile, _
''                                           MatchingIndexArray:=MatchingIndexArray
'' ---
Public Sub Load_Sample_Info_To_Excel(ByRef xFileName As String, _
                                     ByRef MatchingIndexArray() As String)

    'Assign the textbox values from UserFrom Load_Sample_Annot_Raw to array MapHeaders
    Dim MapHeaders(0 To 1) As String
    'MapHeaders(0) = Load_Sample_Annot_Raw.Sample_Name_Text.Text
    MapHeaders(0) = Load_Sample_Annot_Raw.Sample_Amount_Text.Text
    MapHeaders(1) = Load_Sample_Annot_Raw.ISTD_Mixture_Volume_Text.Text
    
    Dim DestHeaders(0 To 1) As String
    'DestHeaders(0) = "Sample_Name"
    DestHeaders(0) = "Sample_Amount"
    DestHeaders(1) = "ISTD_Mixture_Volume_[uL]"
    
    Dim pos As Integer
    
    'When no file is selected
    If TypeName(xFileName) = "Boolean" Then
        End
    End If
    On Error GoTo 0
    
    Dim Lines() As String
    Dim Delimiter As String
    Dim one_line() As String
    Lines = Utilities.Read_File(xFileName)
    Delimiter = Utilities.Get_Delimiter(xFileName)
    
    'Get the first line from sample annot file
    one_line = Split(Lines(0), Delimiter)
    
    Dim MapHeadersIndex As Integer
    
    For MapHeadersIndex = LBound(MapHeaders) To UBound(MapHeaders)
        'Check that it is not empty
        If MapHeaders(MapHeadersIndex) <> vbNullString Then
            'If not empty, get header position from sample annot file
            'Should not have an error as we have check that the columns are in oneline
            If Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True Then
                pos = Application.Match(MapHeaders(MapHeadersIndex), one_line, False) - 1
            Else
                'If the sample annotation has no headers
                'Name will be "Column {Some Number}" use regular expression to get the number
                Dim column_number As String
                Dim regEx As RegExp
                Set regEx = New RegExp
                regEx.Pattern = "\d+"
                column_number = regEx.Execute(MapHeaders(MapHeadersIndex)).Item(0)
                pos = CInt(column_number) - 1
            End If
            
            'Get that position data from sample annot file and assign it to an MapHeaders_Array
            Dim MapHeaders_Array() As String
            Dim MatchingIndex As Integer
            Dim Sample_Data As String
            Dim ArrayLength As Integer
            ArrayLength = 0
            
            For MatchingIndex = 0 To UBound(MatchingIndexArray)
                ReDim Preserve MapHeaders_Array(ArrayLength)
                If Len(MatchingIndexArray(MatchingIndex)) <> 0 Then
                    If Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True Then
                        'We need to add one as the sample annot file has an additional header which we do not want to include
                        Sample_Data = Trim$(Split(Lines(CInt(MatchingIndexArray(MatchingIndex)) + 1), Delimiter)(pos))
                    Else
                        Sample_Data = Trim$(Split(Lines(CInt(MatchingIndexArray(MatchingIndex))), Delimiter)(pos))
                    End If
                    MapHeaders_Array(ArrayLength) = Sample_Data
                    'Debug.Print MapHeaders_Array(ArrayLength)
                Else
                    MapHeaders_Array(ArrayLength) = vbNullString
                End If
                ArrayLength = ArrayLength + 1
            Next MatchingIndex
            
            'We clear any existing entries when necessary, by then the user should have indicated that they want
            'to overwrite the data in the sub Merge_With_Sample_Annot
            Utilities.Clear_Columns HeaderToClear:=DestHeaders(MapHeadersIndex), _
                                    HeaderRowNumber:=1, _
                                    DataStartRowNumber:=2
            Utilities.Load_To_Excel Data_Array:=MapHeaders_Array, _
                                    HeaderName:=DestHeaders(MapHeadersIndex), _
                                    HeaderRowNumber:=1, _
                                    DataStartRowNumber:=2, _
                                    MessageBoxRequired:=False
            
        End If
    Next MapHeadersIndex
    
End Sub

