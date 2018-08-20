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

Private Function GetRawDataFileType(ByRef Lines() As String, Delimiter As String, xFileName As String) As String
    Dim first_line() As String
    Dim second_line() As String
    'Get the first line
    first_line = Split(Lines(0), Delimiter)
    
    'If sample is in first line, check the second line
    If first_line(0) = "Sample" Then
        If Utilities.StringArrayLen(Lines) > 1 Then
            second_line = Split(Lines(1), Delimiter)
            If Utilities.IsInArray("Data File", second_line) Then
                GetRawDataFileType = "AgilentWideForm"
            End If
        End If
    ElseIf first_line(0) = "Compound Method" Then
        GetRawDataFileType = "AgilentCompoundForm"
    ElseIf first_line(0) = "Sample Name" Then
        GetRawDataFileType = "Sciex"
    End If
    
    'Give an error if we are unable to find up where the raw data is coming from
    If GetRawDataFileType = "" Then
        MsgBox "Cannot identify the raw data file type (Agilent or SciEx) for " & xFileName
        Exit Function
    End If
    
End Function

Private Function ReadFile(xFileName As Variant) As String()
    ' Load the file into a string.
    fnum = FreeFile
    Open xFileName For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum
    
    ' Break the file into lines.
    ReadFile = Split(whole_file, vbCrLf)
    
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

Private Function GetFileBaseName(xFileName As Variant) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileBaseName = fso.GetFileName(xFileName)
End Function

Private Function Get_Header_Col_Position_From_2Darray(ByRef Lines() As String, HeaderName As String, _
                                                      HeaderRowNumber As Integer, Delimiter As String) As Variant
    
    'Go to the next line
    Dim header_line() As String
    header_line = Split(Lines(HeaderRowNumber), Delimiter)

    Get_Header_Col_Position_From_2Darray = Null
    'Find the index where the header name first occurred
    For i = LBound(header_line) To UBound(header_line)
        If header_line(i) = HeaderName Then
            Get_Header_Col_Position_From_2Darray = i
            Exit For
        End If
    Next i
    
    If IsNull(Get_Header_Col_Position_From_2Darray) Then
        MsgBox HeaderName & " is missing in the input file "
        End
    End If
    
End Function


Private Function Load_Columns_From_2Darray(ByRef strArray() As String, ByRef Lines() As String, HeaderName As String, _
                                           HeaderRowNumber As Integer, DataStartRowNumber As Integer, Delimiter As String, _
                                           MessageBoxRequired As Boolean, RemoveBlksAndReplicates As Boolean, _
                                           Optional ByVal IgnoreEmptyArray As Boolean = True) As String()
    'We are updating the strArray
    'Dim TotalRows As Long
    Dim i As Long
    Dim ArrayLength As Long
    ArrayLength = Utilities.StringArrayLen(strArray)
    
    'Get column position of Transition_Name_ISTD
    Dim HeaderColNumber As Variant
    HeaderColNumber = Get_Header_Col_Position_From_2Darray(Lines(), HeaderName, HeaderRowNumber, Delimiter)
    
    For i = DataStartRowNumber To UBound(Lines) - 1
        'Get the Transition_Name and remove the whitespaces
        Transition_Name = Trim(Split(Lines(i), Delimiter)(HeaderColNumber))
        If RemoveBlksAndReplicates Then
            'Check if the Transition name is not empty and duplicate
            InArray = Utilities.IsInArray(Transition_Name, strArray)
            If Len(Transition_Name) <> 0 And Not InArray Then
                ReDim Preserve strArray(ArrayLength)
                strArray(ArrayLength) = Transition_Name
                'Debug.Print Transition_Array(ArrayLength)
                ArrayLength = ArrayLength + 1
            End If
        Else
            ReDim Preserve strArray(ArrayLength)
            strArray(ArrayLength) = Transition_Name
            'Debug.Print Transition_Array(ArrayLength)
            ArrayLength = ArrayLength + 1
        End If
    Next i
    
    Load_Columns_From_2Darray = strArray
    
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
        Lines = ReadFile(xFileName)
        Delimiter = GetDelimiter(xFileName)
        FileName = GetFileBaseName(xFileName)
        RawDataFileType = GetRawDataFileType(Lines, Delimiter, FileName)
        
        Dim Sample_Name_SubArray() As String
        Dim MS_File_SubArray() As String
        Dim SubarrayLength As Long
        SubarrayLength = 0
        
        'When the Raw file is from Agilent WideTableForm
        If RawDataFileType = "AgilentWideForm" Then
        
            Sample_Name_SubArray = Load_Columns_From_2Darray(Sample_Name_SubArray, Lines, "Data File", 1, 2, Delimiter, True, True)
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
            Sample_Name_SubArray = Load_Columns_From_2Darray(Sample_Name_SubArray, Lines, "Sample Name", 0, 1, Delimiter, True, True)
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
        Lines = ReadFile(xFileName)
        Delimiter = GetDelimiter(xFileName)
        FileName = GetFileBaseName(xFileName)
        RawDataFileType = GetRawDataFileType(Lines, Delimiter, FileName)
    
        'When the Raw file is from Agilent WideTableForm
        If RawDataFileType = "AgilentWideForm" Then
            'We just look at the first row
            Dim first_line() As String
            first_line = Split(Lines(0), Delimiter)
            
            'We update the array length if Transition_Array
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
            Transition_Array = Load_Columns_From_2Darray(Transition_Array, Lines, "Name", 1, 2, Delimiter, True, True)
        'When the Raw File is from Sciex
        ElseIf RawDataFileType = "Sciex" Then
            Transition_Array = Load_Columns_From_2Darray(Transition_Array, Lines, "Component Name", 0, 1, Delimiter, True, True)
        End If
    Next xFileName
    Get_Transition_Array = Transition_Array
End Function
