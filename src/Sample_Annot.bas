Attribute VB_Name = "Sample_Annot"
Public Sub Create_New_Sample_Annot_Tidy(TidyDataFiles As String, _
                                        DataFileType As String, _
                                        SampleProperty As String, _
                                        StartingRowNum As Integer, _
                                        StartingColumnNum As Integer)
    Sheets("Sample_Annot").Activate
    
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
    ArrayLength = 0
    
    For i = 0 To UBound(Sample_Name_Array_from_Tidy_Data) - LBound(Sample_Name_Array_from_Tidy_Data)
        ReDim Preserve MergeStatus(ArrayLength)
        ReDim Preserve SampleType(ArrayLength)
        
        MergeStatus(ArrayLength) = "Valid"
        SampleType(ArrayLength) = Sample_Type_Identifier.Get_Sample_Type(Sample_Name_Array_from_Tidy_Data(i))
        
        ArrayLength = ArrayLength + 1
    Next i
    
    Dim HeaderNameArray(0 To 3) As String
    HeaderNameArray(0) = "Data_File_Name"
    HeaderNameArray(1) = "Merge_Status"
    HeaderNameArray(2) = "Sample_Name"
    HeaderNameArray(3) = "Sample_Type"
    
    Call Utilities.OverwriteSeveralHeaders(HeaderNameArray, HeaderRowNumber:=1, DataStartRowNumber:=2)
    
    Call Utilities.Load_To_Excel(MS_File_Array, "Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
    Call Utilities.Load_To_Excel(MergeStatus, "Merge_Status", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
    Call Utilities.Load_To_Excel(Sample_Name_Array_from_Tidy_Data, "Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
    Call Utilities.Load_To_Excel(SampleType, "Sample_Type", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
    
                                                                                    

End Sub

Public Sub Create_New_Sample_Annot_Raw(RawDataFiles As String)
    Sheets("Sample_Annot").Activate
    'File are taken from userfrom Load_Sample_Annot_Raw
    'Hence they must exists and joined together by ;
    Dim RawDataFilesArray() As String
    RawDataFilesArray = Split(RawDataFiles, ";")
    
    'Load the Sample_Name from Raw Data
    Dim MS_File_Array() As String
    Dim Sample_Name_Array_from_Raw_Data() As String
    'Dim TotalRows As Long
    Sample_Name_Array_from_Raw_Data = Load_Raw_Data.Get_Sample_Name_Array(RawDataFilesArray, MS_File_Array)
    
    Dim MergeStatus() As String
    Dim SampleType() As String
    Dim ArrayLength As Integer
    ArrayLength = 0
    
    For i = 0 To UBound(Sample_Name_Array_from_Raw_Data) - LBound(Sample_Name_Array_from_Raw_Data)
        ReDim Preserve MergeStatus(ArrayLength)
        ReDim Preserve SampleType(ArrayLength)
        
        MergeStatus(ArrayLength) = "Valid"
        SampleType(ArrayLength) = Sample_Type_Identifier.Get_Sample_Type(Sample_Name_Array_from_Raw_Data(i))
        
        ArrayLength = ArrayLength + 1
    Next i
    
    Dim HeaderNameArray(0 To 3) As String
    HeaderNameArray(0) = "Data_File_Name"
    HeaderNameArray(1) = "Merge_Status"
    HeaderNameArray(2) = "Sample_Name"
    HeaderNameArray(3) = "Sample_Type"
    
    Call Utilities.OverwriteSeveralHeaders(HeaderNameArray, HeaderRowNumber:=1, DataStartRowNumber:=2)
    
    Call Utilities.Load_To_Excel(MS_File_Array, "Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
    Call Utilities.Load_To_Excel(MergeStatus, "Merge_Status", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
    Call Utilities.Load_To_Excel(Sample_Name_Array_from_Raw_Data, "Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
    Call Utilities.Load_To_Excel(SampleType, "Sample_Type", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
    
    
End Sub

Public Sub Merge_With_Sample_Annot(RawDataFiles As String, SampleAnnotFile As String)
    Sheets("Sample_Annot").Activate
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
    Sample_Name_Array_from_Sample_Annot = Sample_Annot.Get_Sample_Name_Array(SampleAnnotFile)
    
    'Match the Sample_Name from Raw Data to the one in Sample Annotation
    'Store merge status and matching index to the array
    Dim MatchingIndex() As String
    Dim MergeStatus() As String
    Dim SampleType() As String
    Dim MergeFailure As Boolean
    Dim ArrayLength As Integer
    MergeFailure = False
    ArrayLength = 0
    
    'For debugging
    'For i = 0 To UBound(Sample_Name_Array_from_Sample_Annot) - LBound(Sample_Name_Array_from_Sample_Annot)
    '    Debug.Print Sample_Name_Array_from_Sample_Annot(i)
    'Next i
    
    For i = 0 To UBound(Sample_Name_Array_from_Raw_Data) - LBound(Sample_Name_Array_from_Raw_Data)
        'Get the positions of where the sample name of the raw data can be found in the sample annotation
        Dim Positions() As String
        Positions = WhereInArray(Sample_Name_Array_from_Raw_Data(i), Sample_Name_Array_from_Sample_Annot)
        
        ReDim Preserve MergeStatus(ArrayLength)
        ReDim Preserve MatchingIndex(ArrayLength)
        ReDim Preserve SampleType(ArrayLength)
        
        'Display results if there is no match, unique match or duplicates
        If StringArrayLen(Positions) = 0 Then
            'Debug.Print "Empty"
            MergeStatus(ArrayLength) = "Missing in Annot File"
            MatchingIndex(ArrayLength) = ""
            MergeFailure = True
        ElseIf StringArrayLen(Positions) > 1 Then
            'Debug.Print "Duplicate"
            For j = 0 To UBound(Positions) - LBound(Positions)
                If Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True Then
                    Positions(j) = CStr(CInt(Positions(j)) + 2)
                Else
                    Positions(j) = CStr(CInt(Positions(j)) + 1)
                End If
            Next j
            
            MergeStatus(ArrayLength) = "Duplicate at line " & Join(Positions, ", ")
            MatchingIndex(ArrayLength) = ""
            MergeFailure = True
        Else
            'Debug.Print "Ok"
            MergeStatus(ArrayLength) = "Valid"
            MatchingIndex(ArrayLength) = Positions(0)
        End If
        
        SampleType(ArrayLength) = Sample_Type_Identifier.Get_Sample_Type(Sample_Name_Array_from_Raw_Data(i))
        
        ArrayLength = ArrayLength + 1
    Next i
    
    Dim HeaderNameArray(0 To 3) As String
    HeaderNameArray(0) = "Data_File_Name"
    HeaderNameArray(1) = "Merge_Status"
    HeaderNameArray(2) = "Sample_Name"
    HeaderNameArray(3) = "Sample_Type"
    Call Utilities.OverwriteSeveralHeaders(HeaderNameArray, HeaderRowNumber:=1, DataStartRowNumber:=2)
      
    'Load Data into the excel sheet
    Call Sample_Annot.Load_Sample_Info_To_Excel(SampleAnnotFile, MatchingIndex)
    Call Utilities.Load_To_Excel(MS_File_Array, "Data_File_Name", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
    Call Utilities.Load_To_Excel(MergeStatus, "Merge_Status", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
    Call Utilities.Load_To_Excel(Sample_Name_Array_from_Raw_Data, "Sample_Name", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
    Call Utilities.Load_To_Excel(SampleType, "Sample_Type", HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
    
    'Notify the user if some rows in the raw data cannot merge with the sample annotation
    If MergeFailure Then
        MsgBox ("Some rows in the raw data are unable to merge with the annotation file.")
    End If
    
End Sub

Private Function Get_Sample_Name_Array(ByRef xFileName As String) As String()

    'When no file is selected
    If TypeName(xFileName) = "Boolean" Then
        End
    End If
    On Error GoTo 0
    
    Dim Lines() As String
    Dim Delimiter As String
    Dim first_line() As String
    Lines = ReadFile(xFileName)
    Delimiter = GetDelimiter(xFileName)
    
    'Get the first line from sample annot file
    first_line = Split(Lines(0), Delimiter)
    
    'Get the data starting row and the right column for the Sample Name
    Dim data_starting_line As Integer
    Dim Sample_Column_Name_pos As Integer
    data_starting_line = GetSampleAnnotStartingLine
    Sample_Column_Name_pos = GetSampleColumnNamePosition(first_line)
    
    'Get the column name to extract the sample name from sample annotation file
    Dim Sample_Column_Name As String
    Sample_Column_Name = Load_Sample_Annot_Raw.Sample_Name_Text.Text

    'For the function output
    Dim Sample_Name_Array() As String
    Dim Sample_Data As String
    Dim ArrayLength As Integer
    ArrayLength = 0

    'Check that it is not empty, it should not be empty based on how we code the userform Load_Sample_Annot_Raw
    If Sample_Column_Name <> "" Then
                        
        'Extract the sample name into the array
        For j = data_starting_line To UBound(Lines) - 1
            'Get the data at the right pos and remove whitespaces
            Sample_Data = Trim(Split(Lines(j), Delimiter)(Sample_Column_Name_pos))
            ReDim Preserve Sample_Name_Array(ArrayLength)
            Sample_Name_Array(ArrayLength) = Sample_Data
            'Debug.Print Sample_Name_Array(ArrayLength)
            ArrayLength = ArrayLength + 1
        Next j
            
        Call ClearDotD_inAgilentDataFile(Sample_Name_Array)
    End If
    Get_Sample_Name_Array = Sample_Name_Array
    
End Function

Private Function GetSampleColumnNamePosition(first_line() As String) As Integer

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
        Dim regEx As New RegExp
        regEx.Pattern = "\d+"
        Sample_Column_Name_pos = CInt(regEx.Execute(Sample_Column_Name)(0)) - 1
    End If
    
    GetSampleColumnNamePosition = Sample_Column_Name_pos

End Function

Private Function GetSampleAnnotStartingLine() As Integer

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
    
    GetSampleAnnotStartingLine = data_starting_line

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

Private Function ReadFile(xFileName As Variant) As String()
    ' Load the file into a string.
    fnum = FreeFile
    Open xFileName For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum
    
    ' Break the file into lines.
    ReadFile = Split(whole_file, vbCrLf)
    
End Function

Private Sub ClearDotD_inAgilentDataFile(ByRef AgilentDataFile() As String)
    For i = 0 To Utilities.StringArrayLen(AgilentDataFile) - 1
        AgilentDataFile(i) = Trim(Replace(AgilentDataFile(i), ".d", ""))
    Next i
End Sub

Private Sub Load_Sample_Info_To_Excel(ByRef xFileName As String, ByRef MatchingIndex() As String)

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
    Lines = ReadFile(xFileName)
    Delimiter = GetDelimiter(xFileName)
    
    'Get the first line from sample annot file
    one_line = Split(Lines(0), Delimiter)
    
    For i = LBound(MapHeaders) To UBound(MapHeaders)
        'Check that it is not empty
        If MapHeaders(i) <> "" Then
            'If not empty, get header position from sample annot file
            'Should not have an error as we have check that the columns are in oneline
            If Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True Then
                pos = Application.Match(MapHeaders(i), one_line, False) - 1
            Else
                'If the sample annotation has no headers
                'Name will be "Column {Some Number}" use regular expression to get the number
                Dim regEx As New RegExp
                regEx.Pattern = "\d+"
                pos = CInt(regEx.Execute(MapHeaders(i))(0)) - 1
            End If
            
            'Get that position data from sample annot file and assign it to an MapHeaders_Array
            Dim MapHeaders_Array() As String
            ArrayLength = 0
            
            For j = 0 To UBound(MatchingIndex)
                ReDim Preserve MapHeaders_Array(ArrayLength)
                If Len(MatchingIndex(j)) <> 0 Then
                    If Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True Then
                        'We need to add one as the sample annot file has an additional header which we do not want to include
                        Sample_Data = Trim(Split(Lines(CInt(MatchingIndex(j)) + 1), Delimiter)(pos))
                    Else
                        Sample_Data = Trim(Split(Lines(CInt(MatchingIndex(j))), Delimiter)(pos))
                    End If
                    MapHeaders_Array(ArrayLength) = Sample_Data
                    'Debug.Print MapHeaders_Array(ArrayLength)
                Else
                    MapHeaders_Array(ArrayLength) = ""
                End If
                ArrayLength = ArrayLength + 1
            Next j
            
            'We clear any existing entries when necessary, by then the user should have indicated that they want
            'to overwrite the data in the sub Merge_With_Sample_Annot
            Call Utilities.Clear_Columns(DestHeaders(i), HeaderRowNumber:=1, DataStartRowNumber:=2)
            Call Utilities.Load_To_Excel(MapHeaders_Array, DestHeaders(i), HeaderRowNumber:=1, DataStartRowNumber:=2, MessageBoxRequired:=False)
            
        End If
    Next i
    
End Sub

