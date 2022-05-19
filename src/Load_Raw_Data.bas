Attribute VB_Name = "Load_Raw_Data"
Attribute VB_Description = "Functions that are commonly called to read raw data."
Option Explicit
'@ModuleDescription("Functions that are commonly called to read raw data.")
'@Folder("Load Data Functions")

'@Description("Get Transition Names from an input Agilent raw data file in compound table form and put them into a string array.")

'' Function: Get_Transition_Array_Agilent_Compound
'' --- Code
''  Public Function Get_Transition_Array_Agilent_Compound(ByRef Transition_Array() As String, _
''                                                        ByRef Lines() As String, _
''                                                        ByRef Transition_Name_And_Qualifier_Transition_Column_Indexes() As Long, _
''                                                        ByVal DataStartRowNumber As Long, _
''                                                        ByVal Delimiter As String, _
''                                                        ByVal RemoveBlksAndReplicates As Boolean, _
''                                                        Optional ByVal IgnoreEmptyArray As Boolean = True) As String()
'' ---
''
'' Description:
''
'' Get Transition Names from an input Agilent raw data file
'' in compound table form and put them into a string array.
''
'' Parameters:
''
'' Returns:
''    A string array of Transition Names
''
'' Examples:
''
'' --- Code
''    Dim Transition_Array() As String
''    Dim TestFolder As String
''    Dim RawDataFile As String
''
''    ' Indicate path to the test data folder
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    RawDataFile = TestFolder & "CompoundTableForm_Qualifier.csv"
''
''    Dim Lines() As String
''    Lines = Utilities.Read_File(RawDataFile)
''
''    Dim Transition_Name_And_Qualifier_Transition_Column_Indexes(3) As Long
''    Transition_Name_And_Qualifier_Transition_Column_Indexes(0) = 1
''    Transition_Name_And_Qualifier_Transition_Column_Indexes(1) = 10
''    Transition_Name_And_Qualifier_Transition_Column_Indexes(2) = 14
''    Transition_Name_And_Qualifier_Transition_Column_Indexes(3) = 18
''
''    'Load the sample name and datafile name into the two arrays
''    Transition_Array = Load_Raw_Data.Get_Transition_Array_Agilent_Compound(Transition_Array:=Transition_Array, _
''                                                                           Lines:=Lines, _
''                                                                           Transition_Name_And_Qualifier_Transition_Column_Indexes:=Transition_Name_And_Qualifier_Transition_Column_Indexes, _
''                                                                           DataStartRowNumber:=2, _
''                                                                           Delimiter:=",", _
''                                                                           RemoveBlksAndReplicates:=True, _
''                                                                           IgnoreEmptyArray:=True)
'' ---
Public Function Get_Transition_Array_Agilent_Compound(ByRef Transition_Array() As String, _
                                                      ByRef Lines() As String, _
                                                      ByRef Transition_Name_And_Qualifier_Transition_Column_Indexes() As Long, _
                                                      ByVal DataStartRowNumber As Long, _
                                                      ByVal Delimiter As String, _
                                                      ByVal RemoveBlksAndReplicates As Boolean, _
                                                      Optional ByVal IgnoreEmptyArray As Boolean = True) As String()
Attribute Get_Transition_Array_Agilent_Compound.VB_Description = "Get Transition Names from an input Agilent raw data file in compound table form and put them into a string array."
                                                       
    'We are updating the InputStringArray
    'Dim TotalRows As Long
    Dim Transition_Name As String
    Dim InArray As Boolean
    Dim LinesIndex As Long
    Dim Transition_Name_And_Qualifier_Transition_Column_Index As Long
    Dim ArrayLength As Long
    ArrayLength = Utilities.Get_String_Array_Len(Transition_Array)
    
    For LinesIndex = DataStartRowNumber To UBound(Lines) - 1
        For Transition_Name_And_Qualifier_Transition_Column_Index = LBound(Transition_Name_And_Qualifier_Transition_Column_Indexes) To UBound(Transition_Name_And_Qualifier_Transition_Column_Indexes)
            'Get the Transition_Name or Qualifier Transitions and remove the whitespaces
            Transition_Name = Trim$(Split(Lines(LinesIndex), Delimiter)(Transition_Name_And_Qualifier_Transition_Column_Indexes(Transition_Name_And_Qualifier_Transition_Column_Index)))
            
            If RemoveBlksAndReplicates Then
                'Check if the Transition name is not empty and duplicate
                InArray = Utilities.Is_In_Array(Transition_Name, Transition_Array)
                If Len(Transition_Name) <> 0 And Not InArray Then
                    'If we found a qualifier, rename it as it is provided in transitions (###->###)
                    If Not Transition_Name_And_Qualifier_Transition_Column_Index = 0 Then
                        Transition_Name = "Qualifier (" & Transition_Name & ")"
                    End If
                    ReDim Preserve Transition_Array(ArrayLength)
                    Transition_Array(ArrayLength) = Transition_Name
                    'Debug.Print Transition_Array(ArrayLength)
                    ArrayLength = ArrayLength + 1
                End If
            Else
                'If we found a qualifier, rename it as it is provided in transitions (###->###)
                If Not Transition_Name_And_Qualifier_Transition_Column_Index = 0 Then
                    Transition_Name = "Qualifier (" & Transition_Name & ")"
                End If
                ReDim Preserve Transition_Array(ArrayLength)
                Transition_Array(ArrayLength) = Transition_Name
                'Debug.Print Transition_Array(ArrayLength)
                ArrayLength = ArrayLength + 1
            End If
            
        Next Transition_Name_And_Qualifier_Transition_Column_Index
    Next LinesIndex
    
    Get_Transition_Array_Agilent_Compound = Transition_Array

End Function

'@Description("Get Sample Names from an input raw data file, put them into a string array.")

'' Function: Get_Sample_Name_Array
'' --- Code
''  Public Function Get_Sample_Name_Array(ByRef RawDataFilesArray() As String, _
''                                        ByRef MS_File_Array() As String) As String()
'' ---
''
'' Description:
''
'' Get Sample Names from an input raw data file, put them into
'' a string array.
''
'' Parameters:
''
''    RawDataFilesArray() As String - A string array of File path to a Raw Data (Agilent) File in csv.
''                                    Eg. {FilePath 1, FilePath 2}
''
''    MS_File_Array() As String - A string array of Data File Names to be loaded to the Data_File_Name
''                                column of the Sample_Annot sheet. If this function is used more than
''                                once, this array will be appended.
''                                Eg. {Input_File_Name_1, Input_File_Name_1, Input_File_Name_2}
''
''
'' Returns:
''    A string array of Sample Names and a string array of Data File Names
''
'' Examples:
''
'' --- Code
''   Dim TestFolder As String
''   Dim RawDataFiles As String
''   Dim RawDataFilesArray() As String
''
''   Dim MS_File_Array() As String
''   Dim Sample_Name_Array_from_Raw_Data() As String
''
''   ' Indicate path to the test data folder
''   TestFolder = ThisWorkbook.Path & "\Testdata\"
''   RawDataFiles = TestFolder & "MultipleDataTest2.csv"
''   RawDataFilesArray = Split(RawDataFiles, ";")
''
''   'Load the sample name and datafile name into the two arrays
''   Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray:=RawDataFilesArray, _
''                                                                         MS_File_Array:=MS_File_Array)
'' ---
Public Function Get_Sample_Name_Array(ByRef RawDataFilesArray() As String, _
                                      ByRef MS_File_Array() As String) As String()
Attribute Get_Sample_Name_Array.VB_Description = "Get Sample Names from an input raw data file, put them into a string array."
                                      
    'Get the Sample_Name and where it comes from
    
    'Initialise the Sample Name Array
    Dim Sample_Name_Array() As String
    
    'When no file is selected
    If TypeName(RawDataFilesArray) = "Boolean" Then
        End
    End If
    On Error GoTo 0
    
    Dim xFileName As Variant
    
    For Each xFileName In RawDataFilesArray
    
        Dim Lines() As String
        Dim Delimiter As String
        Dim FileName As String
        Dim RawDataFileType As String
        Lines = Utilities.Read_File(xFileName)
        Delimiter = Utilities.Get_Delimiter(xFileName)
        FileName = Utilities.Get_File_Base_Name(xFileName)
        RawDataFileType = Utilities.Get_Raw_Data_File_Type(Lines, Delimiter, FileName)
        
        Dim Sample_Name_SubArray() As String
        Dim Sample_Name_SubArray_Index As Long
        Dim MS_File_SubArray() As String
        Dim SubarrayLength As Long
        SubarrayLength = 0
        
        'When the Raw file is from Agilent WideTableForm
        If RawDataFileType = "AgilentWideForm" Then
        
            Sample_Name_SubArray = Utilities.Load_Columns_From_2Darray(InputStringArray:=Sample_Name_SubArray, _
                                                                       Lines:=Lines, _
                                                                       HeaderName:="Data File", _
                                                                       HeaderRowNumber:=1, _
                                                                       DataStartRowNumber:=2, _
                                                                       Delimiter:=Delimiter, _
                                                                       RemoveBlksAndReplicates:=True)
            Sample_Name_SubArray = Utilities.Clear_DotD_In_Agilent_Data_File(Sample_Name_SubArray)
            
            'When the Raw file is from Agilent CompoundTableForm
        ElseIf RawDataFileType = "AgilentCompoundForm" Then
            
            Dim Sample_Name As String
            Dim header_line_index As Long
            Dim header_line() As String
            Dim one_line() As String
            
            'Get the header line (row 2) and first line of the data (should be row 3)
            header_line = Split(Lines(1), Delimiter)
            one_line = Split(Lines(2), Delimiter)
            
            For header_line_index = LBound(header_line) To UBound(header_line)
                If header_line(header_line_index) = "Data File" Then
                
                    'Get the Sample_Name and remove the whitespaces
                    Sample_Name = Trim$(one_line(header_line_index))
                    
                    'Check if the Sample_Name is not empty
                    If Len(Sample_Name) <> 0 Then
                        ReDim Preserve Sample_Name_SubArray(SubarrayLength)
                        Sample_Name_SubArray(SubarrayLength) = Trim$(Sample_Name)
                        'Debug.Print Sample_Name_Array(SubarrayLength)
                        SubarrayLength = SubarrayLength + 1
                        
                    End If
                End If
            Next header_line_index
            
            Sample_Name_SubArray = Utilities.Clear_DotD_In_Agilent_Data_File(Sample_Name_SubArray)
            
            'When the Raw File is from Sciex
        ElseIf RawDataFileType = "Sciex" Then
            Sample_Name_SubArray = Utilities.Load_Columns_From_2Darray(InputStringArray:=Sample_Name_SubArray, _
                                                                       Lines:=Lines, _
                                                                       HeaderName:="Sample Name", _
                                                                       HeaderRowNumber:=0, _
                                                                       DataStartRowNumber:=1, _
                                                                       Delimiter:=Delimiter, _
                                                                       RemoveBlksAndReplicates:=True)
        End If
        
        ' When the file type is valid and we have read some data
        If RawDataFileType <> vbNullString Then
            'Update the subarray to the original arrays
            Sample_Name_Array = Utilities.Concantenate_String_Arrays(Sample_Name_Array, Sample_Name_SubArray)
            SubarrayLength = 0
            
            For Sample_Name_SubArray_Index = 0 To Utilities.Get_String_Array_Len(Sample_Name_SubArray) - 1
                ReDim Preserve MS_File_SubArray(SubarrayLength)
                MS_File_SubArray(Sample_Name_SubArray_Index) = FileName
                SubarrayLength = SubarrayLength + 1
            Next Sample_Name_SubArray_Index
            MS_File_Array = Utilities.Concantenate_String_Arrays(MS_File_Array, MS_File_SubArray)
        
        End If
        
        Erase Sample_Name_SubArray
        Erase MS_File_SubArray
        
    Next xFileName
    Get_Sample_Name_Array = Sample_Name_Array
End Function

'@Description("Get Transition Names from an input raw data file, put them into a string array.")

'' Function: Get_Transition_Array_Raw
'' --- Code
''  Public Function Get_Transition_Array_Raw(ByVal RawDataFiles As String) As String()
'' ---
''
'' Description:
''
'' Get Transition Names from an input raw data file, put them into
'' a string array.
''
'' Parameters:
''
''    RawDataFiles As String - File path to a Raw Data (Agilent) File in csv.
''                             If multiple files are required, the different
''                             file path must be separated by ";"
''                             Eg. {FilePath 1};{FilePath 2}
''
'' Returns:
''    A string array of Transition Names.
''
'' Examples:
''
'' --- Code
''   Dim TestFolder As String
''   Dim RawDataFiles As String
''   Dim Transition_Array() As String
''
''   ' Indicate path to the test data folder
''   TestFolder = ThisWorkbook.Path & "\Testdata\"
''   RawDataFiles = TestFolder & "AgilentRawDataTest1.csv"
''
''   ' Get the transition names
''   Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=RawDataFiles)
'' ---
Public Function Get_Transition_Array_Raw(ByVal RawDataFiles As String) As String()
Attribute Get_Transition_Array_Raw.VB_Description = "Get Transition Names from an input raw data file, put them into a string array."
    'If TypeName(xFileNames) = "Boolean" Then
    '    xFileNames = Application.GetOpenFilename(Title:="Load MS Raw Data", MultiSelect:=True)
    '    'When no file is selected
    '    If TypeName(xFileNames) = "Boolean" Then
    '        End
    '    End If
    '    On Error GoTo 0
    'End If
    
    'File are taken from userfrom Load_Transition_Name_Tidy
    'Hence they must exists and joined together by ;
    Dim RawDataFile As Variant
    Dim RawDataFilesArray() As String
    RawDataFilesArray = Split(RawDataFiles, ";")
    
    'Initialise the Transition Array
    Dim Transition_Array() As String
    Dim ArrayLength As Long
    'ArrayLength = 0
      
    For Each RawDataFile In RawDataFilesArray
        
        Dim Lines() As String
        Dim Delimiter As String
        Dim FileName As String
        Dim RawDataFileType As String
        Lines = Utilities.Read_File(RawDataFile)
        Delimiter = Utilities.Get_Delimiter(RawDataFile)
        FileName = Utilities.Get_File_Base_Name(RawDataFile)
        RawDataFileType = Utilities.Get_Raw_Data_File_Type(Lines, Delimiter, FileName)
    
        'When the Raw file is from Agilent WideTableForm
        If RawDataFileType = "AgilentWideForm" Then
        
            'We just look at the first row
            Dim Transition_Name As String
            Dim InArray As Boolean
            Dim first_line() As String
            Dim first_line_index As Long
            first_line = Split(Lines(0), Delimiter)
            
            'We update the array length of Transition_Array
            ArrayLength = Utilities.Get_String_Array_Len(Transition_Array)
            
            For first_line_index = 1 To UBound(first_line)
                'Remove the whitespace and results and method'
                Transition_Name = Trim$(Replace(first_line(first_line_index), "Results", vbNullString))
                Transition_Name = Trim$(Replace(Transition_Name, "Method", vbNullString))
                'Check if the Transition name is not empty and duplicate
                InArray = Utilities.Is_In_Array(Transition_Name, Transition_Array)
                If Transition_Name <> vbNullString And Not InArray Then
                    ReDim Preserve Transition_Array(ArrayLength)
                    Transition_Array(ArrayLength) = Transition_Name
                    ArrayLength = ArrayLength + 1
                End If
            Next first_line_index
            'When the Raw file is from Agilent CompoundTableForm
        ElseIf RawDataFileType = "AgilentCompoundForm" Then
        
            Dim first_header_line() As String
            Dim second_header_line() As String
            Dim first_header_line_index As Long
            Dim second_header_line_index As Long
            first_header_line = Split(Lines(0), Delimiter)
            second_header_line = Split(Lines(1), Delimiter)
            
            Dim Transition_Name_And_Qualifier_Transition_Column_Indexes() As Long
            Dim Transition_Name_And_Qualifier_Transition_Column_Indexes_Length As Long
            Transition_Name_And_Qualifier_Transition_Column_Indexes_Length = 0
            
            'Get the index of compound method name
            'It should appear before the qualifier
            For second_header_line_index = LBound(second_header_line) To UBound(second_header_line)
                If second_header_line(second_header_line_index) = "Name" Then
                    ReDim Preserve Transition_Name_And_Qualifier_Transition_Column_Indexes(Transition_Name_And_Qualifier_Transition_Column_Indexes_Length)
                    Transition_Name_And_Qualifier_Transition_Column_Indexes(Transition_Name_And_Qualifier_Transition_Column_Indexes_Length) = second_header_line_index
                    Transition_Name_And_Qualifier_Transition_Column_Indexes_Length = Transition_Name_And_Qualifier_Transition_Column_Indexes_Length + 1
                    Exit For
                End If
            Next second_header_line_index
            
            'Do a forward fill on the first header line
            Dim fill_header_word As String
            fill_header_word = first_header_line(0)
            
            For first_header_line_index = LBound(first_header_line) To UBound(first_header_line)
                If first_header_line(first_header_line_index) = vbNullString Then
                    first_header_line(first_header_line_index) = fill_header_word
                Else
                    fill_header_word = first_header_line(first_header_line_index)
                End If
            Next first_header_line_index
            
            'Get the index of qualifier method transition
            'Get the index of data file
            'Get the max number of qualifier a transition can have
            Dim Qualifier_Method_Col As RegExp
            Set Qualifier_Method_Col = New RegExp
            Dim Transition_Col As RegExp
            Set Transition_Col = New RegExp
            Dim DataFileName_Col As RegExp
            Set DataFileName_Col = New RegExp
            
            Qualifier_Method_Col.Pattern = "Qualifier \d Method"
            Transition_Col.Pattern = "Transition"
            DataFileName_Col.Pattern = "Data File"
            
            Dim isQualifier_Method_Col As Boolean
            Dim isTransition_Col As Boolean
            Dim isDataFileName_Col As Boolean
            Dim Qualifier_Method_Col_BoolArrray() As Boolean
            Dim DataFileName_Col_BoolArrray() As Boolean
            
            Dim No_of_Qualifier_Method_Transition As Long
            Dim No_of_DataFileName_Col As Long
            Dim No_of_Qual_per_Transition As Long
            No_of_Qualifier_Method_Transition = 0
            No_of_DataFileName_Col = 0
            
            ArrayLength = 0
            
            For first_header_line_index = LBound(first_header_line) To UBound(first_header_line)
            
                isQualifier_Method_Col = Qualifier_Method_Col.Test(first_header_line(first_header_line_index))
                isTransition_Col = Transition_Col.Test(second_header_line(first_header_line_index))
                isDataFileName_Col = DataFileName_Col.Test(second_header_line(first_header_line_index))
                
                ReDim Preserve Qualifier_Method_Col_BoolArrray(ArrayLength)
                ReDim Preserve DataFileName_Col_BoolArrray(ArrayLength)
                Qualifier_Method_Col_BoolArrray(ArrayLength) = isQualifier_Method_Col And isTransition_Col
                DataFileName_Col_BoolArrray(ArrayLength) = isDataFileName_Col
                
                If isQualifier_Method_Col And isTransition_Col Then
                    No_of_Qualifier_Method_Transition = No_of_Qualifier_Method_Transition + 1
                    ReDim Preserve Transition_Name_And_Qualifier_Transition_Column_Indexes(Transition_Name_And_Qualifier_Transition_Column_Indexes_Length)
                    Transition_Name_And_Qualifier_Transition_Column_Indexes(Transition_Name_And_Qualifier_Transition_Column_Indexes_Length) = first_header_line_index
                    Transition_Name_And_Qualifier_Transition_Column_Indexes_Length = Transition_Name_And_Qualifier_Transition_Column_Indexes_Length + 1
                ElseIf isDataFileName_Col Then
                    No_of_DataFileName_Col = No_of_DataFileName_Col + 1
                End If
                
                ArrayLength = ArrayLength + 1
                
            Next first_header_line_index
            
            No_of_Qual_per_Transition = No_of_Qualifier_Method_Transition \ No_of_DataFileName_Col
            ReDim Preserve Transition_Name_And_Qualifier_Transition_Column_Indexes(No_of_Qual_per_Transition)
                      
            Transition_Array = Load_Raw_Data.Get_Transition_Array_Agilent_Compound(Transition_Array:=Transition_Array, _
                                                                                   Lines:=Lines, _
                                                                                   Transition_Name_And_Qualifier_Transition_Column_Indexes:=Transition_Name_And_Qualifier_Transition_Column_Indexes, _
                                                                                   DataStartRowNumber:=2, _
                                                                                   Delimiter:=Delimiter, _
                                                                                   RemoveBlksAndReplicates:=True, _
                                                                                   IgnoreEmptyArray:=True)
            
            'When the Raw File is from Sciex
        ElseIf RawDataFileType = "Sciex" Then
            Transition_Array = Utilities.Load_Columns_From_2Darray(InputStringArray:=Transition_Array, _
                                                                   Lines:=Lines, _
                                                                   HeaderName:="Component Name", _
                                                                   HeaderRowNumber:=0, _
                                                                   DataStartRowNumber:=1, _
                                                                   Delimiter:=Delimiter, _
                                                                   RemoveBlksAndReplicates:=True)
        End If
    Next RawDataFile
    Get_Transition_Array_Raw = Transition_Array
End Function
