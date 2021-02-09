Attribute VB_Name = "Load_Raw_Data"
Private Sub ClearDotD_inAgilentDataFile(ByRef AgilentDataFile() As String)
    For i = 0 To Utilities.StringArrayLen(AgilentDataFile) - 1
        AgilentDataFile(i) = Trim(Replace(AgilentDataFile(i), ".d", ""))
    Next i
End Sub

Private Function ConcantenateStringArrays(Array1() As String, Array2() As String) As String()
    'Update the Sample Name Array
    Dim Array1Length As Long
    Dim Array2Length As Long
    Array1Length = Len(Join(Array1, ""))
    Array2Length = Len(Join(Array2, ""))
    
    If Array1Length > 0 And Array2Length > 0 Then
        ConcantenateStringArrays = Split(Join(Array1, ",") & "," & Join(Array2, ","), ",")
    ElseIf Array1Length > 0 Then
        ConcantenateStringArrays = Array1
    ElseIf Array2Length > 0 Then
        ConcantenateStringArrays = Array2
    Else
        MsgBox "Two arrays cannot be empty"
        Exit Function
    End If
End Function

Private Function GetDelimiter(xFileName As Variant) As String
    FileExtent = Right(xFileName, Len(xFileName) - InStrRev(xFileName, "."))
    'Get the first line
    If FileExtent = "csv" Then
        GetDelimiter = ","
    ElseIf FileExtent = "txt" Then
        GetDelimiter = vbTab
    Else
        MsgBox "Cannot identify delimiter due to unusual file type"
        End
    End If
End Function

Private Function Get_Transition_Array_Agilent_Compound(ByRef Transition_Array() As String, ByRef Lines() As String, ByRef line_subset_index() As Integer, _
                                                       DataStartRowNumber As Integer, Delimiter As String, _
                                                       MessageBoxRequired As Boolean, RemoveBlksAndReplicates As Boolean, _
                                                       Optional ByVal IgnoreEmptyArray As Boolean = True) As String()
                                                       
    'We are updating the strArray
    'Dim TotalRows As Long
    Dim i As Long
    Dim ArrayLength As Long
    ArrayLength = Utilities.StringArrayLen(Transition_Array)
    
    For i = DataStartRowNumber To UBound(Lines) - 1
        For j = LBound(line_subset_index) To UBound(line_subset_index)
            'Get the Transition_Name and remove the whitespaces
            Transition_Name = Trim(Split(Lines(i), Delimiter)(line_subset_index(j)))
            
            If RemoveBlksAndReplicates Then
                'Check if the Transition name is not empty and duplicate
                InArray = Utilities.IsInArray(Transition_Name, strArray)
                If Len(Transition_Name) <> 0 And Not InArray Then
                    If Not j = 0 Then
                        Transition_Name = "Qualifier (" & Transition_Name & ")"
                    End If
                    ReDim Preserve Transition_Array(ArrayLength)
                    Transition_Array(ArrayLength) = Transition_Name
                    'Debug.Print Transition_Array(ArrayLength)
                    ArrayLength = ArrayLength + 1
                End If
            Else
                If Not j = 0 Then
                    Transition_Name = "Qualifier (" & Transition_Name & ")"
                End If
                ReDim Preserve Transition_Array(ArrayLength)
                Transition_Array(ArrayLength) = Transition_Name
                'Debug.Print Transition_Array(ArrayLength)
                ArrayLength = ArrayLength + 1
            End If
            
        Next j
    Next i
    
    Get_Transition_Array_Agilent_Compound = Transition_Array

End Function

'Get the Sample_Name and where it comes from
Public Function Get_Sample_Name_Array(ByRef xFileNames() As String, ByRef MS_File_Array() As String) As String()
    
    'Initialise the Sample Name Array
    Dim Sample_Name_Array() As String
    
    'When no file is selected
    If TypeName(xFileNames) = "Boolean" Then
        End
    End If
    On Error GoTo 0
    
    For Each xFileName In xFileNames
    
        Dim Lines() As String
        Dim Delimiter As String
        Dim FileName As String
        Dim RawDataFileType As String
        Lines = Utilities.Read_File(xFileName)
        Delimiter = GetDelimiter(xFileName)
        FileName = Utilities.Get_File_Base_Name(xFileName)
        RawDataFileType = Utilities.Get_Raw_Data_File_Type(Lines, Delimiter, FileName)
        
        Dim Sample_Name_SubArray() As String
        Dim MS_File_SubArray() As String
        Dim SubarrayLength As Long
        SubarrayLength = 0
        
        'When the Raw file is from Agilent WideTableForm
        If RawDataFileType = "AgilentWideForm" Then
        
            Sample_Name_SubArray = Utilities.Load_Columns_From_2Darray(Sample_Name_SubArray, Lines, _
                                                                       HeaderName:="Data File", _
                                                                       HeaderRowNumber:=1, _
                                                                       DataStartRowNumber:=2, _
                                                                       Delimiter:=Delimiter, _
                                                                       MessageBoxRequired:=True, _
                                                                       RemoveBlksAndReplicates:=True)
            Call ClearDotD_inAgilentDataFile(Sample_Name_SubArray)
            
            'When the Raw file is from Agilent CompoundTableForm
        ElseIf RawDataFileType = "AgilentCompoundForm" Then
            
            'Get the header line (row 2) and first line of the data (should be row 3)
            header_line = Split(Lines(1), Delimiter)
            one_line = Split(Lines(2), Delimiter)
            
            For i = LBound(header_line) To UBound(header_line)
                If header_line(i) = "Data File" Then
                
                    'Get the Sample_Name and remove the whitespaces
                    Sample_Name = Trim(one_line(i))
                    
                    'Check if the Sample_Name is not empty
                    If Len(Sample_Name) <> 0 Then
                        ReDim Preserve Sample_Name_SubArray(SubarrayLength)
                        Sample_Name_SubArray(SubarrayLength) = Trim(Sample_Name)
                        'Debug.Print Sample_Name_Array(SubarrayLength)
                        SubarrayLength = SubarrayLength + 1
                        
                    End If
                End If
            Next i
            
            Call ClearDotD_inAgilentDataFile(Sample_Name_SubArray)
            
            'When the Raw File is from Sciex
        ElseIf RawDataFileType = "Sciex" Then
            Sample_Name_SubArray = Utilities.Load_Columns_From_2Darray(Sample_Name_SubArray, Lines, _
                                                                       HeaderName:="Sample Name", _
                                                                       HeaderRowNumber:=0, _
                                                                       DataStartRowNumber:=1, _
                                                                       Delimiter:=Delimiter, _
                                                                       MessageBoxRequired:=True, _
                                                                       RemoveBlksAndReplicates:=True)
        End If
        
        'Update the subarray to the original arrays
        Sample_Name_Array = ConcantenateStringArrays(Sample_Name_Array, Sample_Name_SubArray)
        SubarrayLength = 0
            
        For i = 0 To Utilities.StringArrayLen(Sample_Name_SubArray) - 1
            ReDim Preserve MS_File_SubArray(SubarrayLength)
            MS_File_SubArray(i) = FileName
            SubarrayLength = SubarrayLength + 1
        Next i
        MS_File_Array = ConcantenateStringArrays(MS_File_Array, MS_File_SubArray)
        
        Erase Sample_Name_SubArray
        Erase MS_File_SubArray
        
    Next xFileName
    Get_Sample_Name_Array = Sample_Name_Array
End Function

Public Function Get_Transition_Array(Optional ByVal xFileNames As Variant = False) As String()
    If TypeName(xFileNames) = "Boolean" Then
        xFileNames = Application.GetOpenFilename(Title:="Load MS Raw Data", MultiSelect:=True)
        'When no file is selected
        If TypeName(xFileNames) = "Boolean" Then
            End
        End If
        On Error GoTo 0
    End If
    
    'Initialise the Transition Array
    Dim Transition_Array() As String
    Dim ArrayLength As Long
    ArrayLength = 0
      
    For Each xFileName In xFileNames
        
        Dim Lines() As String
        Dim Delimiter As String
        Dim FileName As String
        Dim RawDataFileType As String
        Lines = Utilities.Read_File(xFileName)
        Delimiter = GetDelimiter(xFileName)
        FileName = Utilities.Get_File_Base_Name(xFileName)
        RawDataFileType = Utilities.Get_Raw_Data_File_Type(Lines, Delimiter, FileName)
    
        'When the Raw file is from Agilent WideTableForm
        If RawDataFileType = "AgilentWideForm" Then
            'We just look at the first row
            Dim first_line() As String
            first_line = Split(Lines(0), Delimiter)
            
            'We update the array length of Transition_Array
            ArrayLength = Utilities.StringArrayLen(Transition_Array)
            
            For i = 1 To UBound(first_line)
                'Remove the whitespace and results and method'
                Transition_Name = Trim(Replace(first_line(i), "Results", ""))
                Transition_Name = Trim(Replace(Transition_Name, "Method", ""))
                'Check if the Transition name is not empty and duplicate
                InArray = Utilities.IsInArray(Transition_Name, Transition_Array)
                If Len(Transition_Name) <> 0 And Not InArray Then
                    ReDim Preserve Transition_Array(ArrayLength)
                    Transition_Array(ArrayLength) = Transition_Name
                    ArrayLength = ArrayLength + 1
                End If
            Next i
            'When the Raw file is from Agilent CompoundTableForm
        ElseIf RawDataFileType = "AgilentCompoundForm" Then
            Dim first_header_line() As String
            Dim second_header_line() As String
            first_header_line = Split(Lines(0), Delimiter)
            second_header_line = Split(Lines(1), Delimiter)
            
            Dim line_subset_index() As Integer
            Dim line_subset_index_length As Integer
            line_subset_index_length = 0
            
            'Get the index of compound method name
            'It should appear before the qualifier
            For i = LBound(second_header_line) To UBound(second_header_line)
                If second_header_line(i) = "Name" Then
                    ReDim Preserve line_subset_index(line_subset_index_length)
                    line_subset_index(line_subset_index_length) = i
                    line_subset_index_length = line_subset_index_length + 1
                    Exit For
                End If
            Next i
            
            'Get the index of qualifier method transition
            'Get the index of data file
            'Get the max number of qualifier a transition can have
            Dim Qualifier_Method_Col As New RegExp
            Dim Transition_Col As New RegExp
            Dim DataFileName_Col As New RegExp
            Qualifier_Method_Col.Pattern = "Qualifier \d Method"
            Transition_Col.Pattern = "Transition"
            DataFileName_Col.Pattern = "Data File"
            
            Dim isQualifier_Method_Col As Boolean
            Dim isTransition_Col As Boolean
            Dim isDataFileName_Col As Boolean
            Dim Qualifier_Method_Col_BoolArrray() As Boolean
            Dim DataFileName_Col_BoolArrray() As Boolean
            
            Dim No_of_Qualifier_Method_Transition As Integer
            Dim No_of_DataFileName_Col As Integer
            Dim No_of_Qual_per_Transition As Integer
            No_of_Qualifier_Method_Transition = 0
            No_of_DataFileName_Col = 0
            
            ArrayLength = 0
            
            For i = LBound(first_header_line) To UBound(first_header_line)
                isQualifier_Method_Col = Qualifier_Method_Col.Test(first_header_line(i))
                isTransition_Col = Transition_Col.Test(second_header_line(i))
                isDataFileName_Col = DataFileName_Col.Test(second_header_line(i))
                
                ReDim Preserve Qualifier_Method_Col_BoolArrray(ArrayLength)
                ReDim Preserve DataFileName_Col_BoolArrray(ArrayLength)
                Qualifier_Method_Col_BoolArrray(ArrayLength) = isQualifier_Method_Col And isTransition_Col
                DataFileName_Col_BoolArrray(ArrayLength) = isDataFileName_Col
                
                If isQualifier_Method_Col And isTransition_Col Then
                    No_of_Qualifier_Method_Transition = No_of_Qualifier_Method_Transition + 1
                    ReDim Preserve line_subset_index(line_subset_index_length)
                    line_subset_index(line_subset_index_length) = i
                    line_subset_index_length = line_subset_index_length + 1
                ElseIf isDataFileName_Col Then
                    No_of_DataFileName_Col = No_of_DataFileName_Col + 1
                End If
                
                ArrayLength = ArrayLength + 1
                
            Next i
            
            No_of_Qual_per_Transition = No_of_Qualifier_Method_Transition \ No_of_DataFileName_Col
            ReDim Preserve line_subset_index(No_of_Qual_per_Transition)
            
            'For i = LBound(line_subset_index) To UBound(line_subset_index)
            '    Debug.Print line_subset_index(i)
            'Next i
            
            Transition_Array = Get_Transition_Array_Agilent_Compound(Transition_Array, Lines, line_subset_index, 2, Delimiter, True, True)
            
            'When the Raw File is from Sciex
        ElseIf RawDataFileType = "Sciex" Then
            Transition_Array = Utilities.Load_Columns_From_2Darray(Transition_Array, Lines, _
                                                                   HeaderName:="Component Name", _
                                                                   HeaderRowNumber:=0, _
                                                                   DataStartRowNumber:=1, _
                                                                   Delimiter:=Delimiter, _
                                                                   MessageBoxRequired:=True, _
                                                                   RemoveBlksAndReplicates:=True)
        End If
    Next xFileName
    Get_Transition_Array = Transition_Array
End Function


