Attribute VB_Name = "Utilities"
Option Explicit
'@Folder("Utilities Functions")

'' Function: Concantenate_String_Arrays
'' --- Code
''  Public Function Concantenate_String_Arrays(ByRef TopArray() As String, _
''                                             ByRef BottomArray() As String) As String()
'' ---
''
'' Description:
''
'' Concantenate two string arrays into one string array
''
'' Parameters:
''
''    TopArray() As String - First string array to be concantenated
''
''    BottomArray() As String - Second string array to be concantenated
''
'' Returns:
''    A string array of that contains {Contents of TopArray, Contents of BottomArray}
''
'' Examples:
''
'' --- Code
''    Dim TopArray(2) As String
''    Dim BottomArray(2) As String
''    Dim Concatenated_Array() As String
''
''    TopArray(0) = "SM 36:0"
''    TopArray(1) = "SM 36:1"
''    TopArray(2) = "SM 36:2"
''    BottomArray(0) = "SM 38:0"
''    BottomArray(1) = "SM 38:1"
''    BottomArray(2) = "SM 38:2"
''
''    Concatenated_Array = Utilities.Concantenate_String_Arrays(TopArray:=TopArray, _
''                                                              BottomArray:=BottomArray)
'' ---
Public Function Concantenate_String_Arrays(ByRef TopArray() As String, _
                                           ByRef BottomArray() As String) As String()
    'Update the Sample Name Array
    Dim TopArrayLength As Long
    Dim BottomArrayLength As Long
    TopArrayLength = Len(Join(TopArray, vbNullString))
    BottomArrayLength = Len(Join(BottomArray, vbNullString))
    
    If TopArrayLength > 0 And BottomArrayLength > 0 Then
        Concantenate_String_Arrays = Split(Join(TopArray, ",") & "," & Join(BottomArray, ","), ",")
    ElseIf TopArrayLength > 0 Then
        Concantenate_String_Arrays = TopArray
    ElseIf BottomArrayLength > 0 Then
        Concantenate_String_Arrays = BottomArray
    Else
        MsgBox "Two arrays cannot be empty"
        Exit Function
    End If
End Function

'' Function: Get_RowName_Position_From_2Darray
'' --- Code
''  Public Function Get_RowName_Position_From_2Darray(ByRef Lines() As String, _
''                                                    ByVal RowName As String, _
''                                                    ByVal RowNameNumber As Variant, _
''                                                    ByVal Delimiter As String) As Variant
'' ---
''
'' Description:
''
'' Get the row position of "RowName" in the column highlighted
'' by "RowNameNumber" in a 2D string array Lines().
''
'' Parameters:
''
''    Lines() As String - A 2D string array. We let the number of elements (rows)
''                        of Lines be n. Each element of Lines is
''                        in the form "EntryRow1 Column1,EntryRow1 Column2,
''                        ...,EntryRow1 Column p" where
''                        p is the number of columns.
''
''    RowName As String - Name of the entry to look for in Lines() to get the
''                        starting row position
''
''    RowNameNumber As Variant - Column number to look at in Lines(), to look for RowName to get the
''                               starting row position. 0 is the first column. p-1 is the last column.
''
''    Delimiter As String - Delimiter used to separate each entries in
''                          one element of Lines(). For example, the delimiter for
''                          "EntryRow1 Column1,EntryRow1 Column2,...,EntryRow1 Column p"
''                          is ","
''
'' Returns:
''    An integer indicating the row position of "RowName" .
''
'' Examples:
''
'' --- Code
''    Dim TestFolder As String
''    Dim TidyDataColumnFile As String
''
''    'Indicate path to the test data folder
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    TidyDataColumnFile = TestFolder & "TidyTransitionColumn.csv"
''
''    'Read the csv files
''    Dim Lines() As String
''    Lines = Utilities.Read_File(TidyDataColumnFile)
''
''    Dim RowNamePosition As Variant
''
''    ' Find which row is "Sample_Name" found in the "RowName" column.
''    ' "RowName" column is the first column hence RowNameNumber is set as 0
''    ' RowNamePosition should return 0 as "Sample_Name" is the first row
''    RowNamePosition = Utilities.Get_RowName_Position_From_2Darray(Lines:=Lines, _
''                                                                  RowName:="Sample_Name", _
''                                                                  RowNameNumber:=0, _
''                                                                  Delimiter:=",")
''
''    ' Must be 0
''    Debug.Print RowNamePosition
''
''    ' RowNamePosition should return 1 as "Sample1" is the second row
''    ' in the first column
''    RowNamePosition = Utilities.Get_RowName_Position_From_2Darray(Lines:=Lines, _
''                                                                  RowName:="Sample1", _
''                                                                  RowNameNumber:=0, _
''                                                                  Delimiter:=",")
''    ' Must be 1
''    Debug.Print RowNamePosition
'' ---
Public Function Get_RowName_Position_From_2Darray(ByRef Lines() As String, _
                                                  ByVal RowName As String, _
                                                  ByVal RowNameNumber As Variant, _
                                                  ByVal Delimiter As String) As Variant
    
    Dim row_name() As String
    Dim lines_index As Long
    Get_RowName_Position_From_2Darray = Null
    
    For lines_index = LBound(Lines) To UBound(Lines) - 1
        'Get the Row_Name and remove the whitespaces
        row_name = Split(Lines(lines_index), Delimiter)
        'Debug.Print Trim(row_name(RowNameNumber))
        If Trim$(row_name(RowNameNumber)) = RowName Then
            Get_RowName_Position_From_2Darray = lines_index
            Exit For
        End If
    Next lines_index
    
    If IsNull(Get_RowName_Position_From_2Darray) Then
        MsgBox RowName & " is missing in the input file "
        End
    End If
    
End Function

'' Function: Load_Rows_From_2Darray
'' --- Code
''  Public Function Load_Rows_From_2Darray(ByRef InputStringArray() As String, _
''                                         ByRef Lines() As String, _
''                                         ByVal DataStartColumnNumber As Long, _
''                                         ByVal Delimiter As String, _
''                                         ByVal RemoveBlksAndReplicates As Boolean, _
''                                         Optional ByVal RowName As String, _
''                                         Optional ByVal RowNameNumber As Variant, _
''                                         Optional ByVal DataStartRowNumber As Variant) As String()
'' ---
''
'' Description:
''
'' Load row entries from a 2D string array given a
'' starting column and row number. If starting row number
'' cannot be identified, user can key in the starting "RowName" to look
'' out for at column number "RowNameNumber" and the system will
'' help identify the correct starting row number.
''
'' Parameters:
''
''    InputStringArray() As String - A string Array to load/append the row entries
''
''    Lines() As String - A 2D string array. We let the number of elements (rows)
''                        of Lines be n. Each element of Lines is
''                        in the form "EntryRow1 Column1,EntryRow1 Column2,
''                        ...,EntryRow1 Column p" where
''                        p is the number of columns.
''
''    DataStartColumnNumber As Long - An integer to indicate the starting column to
''                                    read data and put it to InputStringArray
''
''    Delimiter As String - Delimiter used to separate each entries in
''                          one element of Lines(). For example, the delimiter for
''                          "EntryRow1 Column1,EntryRow1 Column2,...,EntryRow1 Column p"
''                          is ","
''
''    RemoveBlksAndReplicates As Boolean - If set to True, the system will remove duplicates
''                                         and blank values in InputStringArray
''
''    RowName As String - Name of the entry to look for in Lines() to get the
''                        starting row position
''
''    RowNameNumber As Variant - Column number to look at in Lines(), to look for RowName to get the
''                               starting row position. 0 is the first column. p-1 is the last column.
''
''    DataStartRowNumber As Variant - An integer to indicate the starting row to read data
''                                    and put it to InputStringArray.
''
'' Returns:
''    A string array of row entries from a 2D string array.
''
'' Examples:
''
'' --- Code
''    Dim TestFolder As String
''    Dim Transition_Array() As String
''    Dim TidyDataRowFile As String
''
''    'Indicate path to the test data folder
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    TidyDataRowFile = TestFolder & "TidyTransitionRow.csv"
''
''    'Read the csv files
''    Dim Lines() As String
''    Lines = Utilities.Read_File(TidyDataRowFile)
''
''    Transition_Array = Utilities.Load_Columns_From_2Darray(InputStringArray:=Transition_Array, _
''                                                           Lines:=Lines, _
''                                                           DataStartColumnNumber:=0, _
''                                                           DataStartRowNumber:=1, _
''                                                           Delimiter:=",", _
''                                                           RemoveBlksAndReplicates:=True)
''
''    Dim header_line_index As Long
''    For header_line_index = LBound(Transition_Array) To UBound(Transition_Array)
''        Debug.Print Transition_Array(header_line_index)
''    Next header_line_index
'' ---
Public Function Load_Rows_From_2Darray(ByRef InputStringArray() As String, _
                                       ByRef Lines() As String, _
                                       ByVal DataStartColumnNumber As Long, _
                                       ByVal Delimiter As String, _
                                       ByVal RemoveBlksAndReplicates As Boolean, _
                                       Optional ByVal RowName As String, _
                                       Optional ByVal RowNameNumber As Variant, _
                                       Optional ByVal DataStartRowNumber As Variant) As String()
    
                                     
    'Get column position of a given header name
    Dim RowNamePosition As Variant
    If Not Trim$(RowName) = vbNullString And Not IsMissing(RowNameNumber) Then
        RowNamePosition = Utilities.Get_RowName_Position_From_2Darray(Lines:=Lines, _
                                                                      RowName:=RowName, _
                                                                      RowNameNumber:=RowNameNumber, _
                                                                      Delimiter:=Delimiter)
    ElseIf Not IsMissing(DataStartRowNumber) Then
        RowNamePosition = DataStartRowNumber
    End If
    
    'We just look at the one row the user indicates
    Dim Transition_Name As String
    Dim ArrayLength As Long
    Dim InArray As Boolean
    Dim row_line_index As Long
    Dim row_line() As String
    row_line = Split(Lines(RowNamePosition), Delimiter)
    
    'We update the array length of Transition_Array
    ArrayLength = Utilities.Get_String_Array_Len(InputStringArray)
    
    For row_line_index = DataStartColumnNumber To UBound(row_line)
        'Get the Transition_Name and remove the whitespaces
        Transition_Name = Trim$(row_line(row_line_index))
        
        If RemoveBlksAndReplicates Then
            'Check if the Transition name is not empty and duplicate
            InArray = Utilities.Is_In_Array(Transition_Name, InputStringArray)
            If Len(Transition_Name) <> 0 And Not InArray Then
                ReDim Preserve InputStringArray(ArrayLength)
                InputStringArray(ArrayLength) = Transition_Name
                'Debug.Print InputStringArray(ArrayLength)
                ArrayLength = ArrayLength + 1
            End If
        Else
            ReDim Preserve InputStringArray(ArrayLength)
            InputStringArray(ArrayLength) = Transition_Name
            'Debug.Print InputStringArray(ArrayLength)
            ArrayLength = ArrayLength + 1
        End If
        
    Next row_line_index
    
    Load_Rows_From_2Darray = InputStringArray
    
End Function

'' Function: Get_Header_Col_Position_From_2Darray
'' --- Code
''  Public Function Get_Header_Col_Position_From_2Darray(ByRef Lines() As String, _
''                                                       ByVal HeaderName As String, _
''                                                       ByVal HeaderRowNumber As Variant, _
''                                                       ByVal Delimiter As String) As Variant
'' ---
''
'' Description:
''
'' Get the column position of "HeaderName" in the row highlighted
'' by "HeaderRowNumber" in a 2D string array Lines().
''
'' Parameters:
''
''    Lines() As String - A 2D string array. We let the number of elements (rows)
''                        of Lines be n. Each element of Lines is
''                        in the form "EntryRow1 Column1,EntryRow1 Column2,
''                        ...,EntryRow1 Column p" where
''                        p is the number of columns.
''
''    HeaderName As String - Name of the entry to look for in Lines() to get the
''                           starting column position
''
''    HeaderRowNumber As Variant - Row number to look at in Lines(), to look for HeaderName to get the
''                                 starting column position. 0 is the first row. n-1 is the last row.
''
''    Delimiter As String - Delimiter used to separate each entries in
''                          one element of Lines(). For example, the delimiter for
''                          "EntryRow1 Column1,EntryRow1 Column2,...,EntryRow1 Column p"
''                          is ","
''
'' Returns:
''    An integer indicating the column position of "HeaderName" .
''
'' Examples:
''
'' --- Code
''    Dim TestFolder As String
''    Dim TidyDataColumnFile As String
''
''    'Indicate path to the test data folder
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    TidyDataColumnFile = TestFolder & "TidyTransitionColumn.csv"
''
''    'Read the csv files
''    Dim Lines() As String
''    Lines = Utilities.Read_File(TidyDataColumnFile)
''
''    'Get column position of a given header name
''    Dim HeaderColNumber As Variant
''
''    ' Find which column is "Sample_Name" found in the "HeaderName" row.
''    ' "HeaderName" row is the first row hence HeaderRowNumber is set as 0
''    ' HeaderColNumber should return 0 as "Sample_Name" is the first column
''    HeaderColNumber = Utilities.Get_Header_Col_Position_From_2Darray(Lines:=Lines, _
''                                                                     HeaderName:="Sample_Name", _
''                                                                     HeaderRowNumber:=0, _
''                                                                     Delimiter:=",")
''
''    ' Must be 0
''    Debug.Print HeaderColNumber
''
''    ' HeaderColNumber should return 1 as "LPC 14:0" is at the second column
''    HeaderColNumber = Utilities.Get_Header_Col_Position_From_2Darray(Lines:=Lines, _
''                                                                     HeaderName:="LPC 14:0", _
''                                                                     HeaderRowNumber:=0, _
''                                                                     Delimiter:=",")
''
''    ' Must be 1
''    Debug.Print HeaderColNumber
'' ---
Public Function Get_Header_Col_Position_From_2Darray(ByRef Lines() As String, _
                                                     ByVal HeaderName As String, _
                                                     ByVal HeaderRowNumber As Variant, _
                                                     ByVal Delimiter As String) As Variant
    
    'Go to the next line
    Dim header_line_index As Long
    Dim header_line() As String
    header_line = Split(Lines(HeaderRowNumber), Delimiter)

    Get_Header_Col_Position_From_2Darray = Null
    'Find the index where the header name first occurred
    For header_line_index = LBound(header_line) To UBound(header_line)
        If header_line(header_line_index) = HeaderName Then
            Get_Header_Col_Position_From_2Darray = header_line_index
            Exit For
        End If
    Next header_line_index

    If IsNull(Get_Header_Col_Position_From_2Darray) Then
        MsgBox HeaderName & " is missing in the input file "
        End
    End If
    
End Function

'' Function: Load_Columns_From_2Darray
'' --- Code
''  Public Function Load_Columns_From_2Darray(ByRef InputStringArray() As String, _
''                                            ByRef Lines() As String, _
''                                            ByVal DataStartRowNumber As Long, _
''                                            ByVal Delimiter As String, _
''                                            ByVal RemoveBlksAndReplicates As Boolean, _
''                                            Optional ByVal HeaderName As String, _
''                                            Optional ByVal HeaderRowNumber As Variant, _
''                                            Optional ByVal DataStartColumnNumber As Variant) As String()
'' ---
''
'' Description:
''
'' Load column entries from a 2D string array given a
'' starting row and column number. If starting column number
'' cannot be identified, user can key in the starting "HeaderName" to look
'' out for at row number "HeaderRowNumber" and the system will
'' help identify the correct starting column number.
''
'' Parameters:
''
''    InputStringArray() As String - A string Array to load/append the column entries
''
''    Lines() As String - A 2D string array. We let the number of elements (rows)
''                        of Lines be n. Each element of Lines is
''                        in the form "EntryRow1 Column1,EntryRow1 Column2,
''                        ...,EntryRow1 Column p" where
''                        p is the number of columns.
''
''    DataStartRowNumber As Long - An integer to indicate the starting row to
''                                 read data and put it to InputStringArray
''
''    Delimiter As String - Delimiter used to separate each entries in
''                          one element of Lines(). For example, the delimiter for
''                          "EntryRow1 Column1,EntryRow1 Column2,...,EntryRow1 Column p"
''                          is ","
''
''    RemoveBlksAndReplicates As Boolean - If set to True, the system will remove duplicates
''                                         and blank values in InputStringArray
''
''    HeaderName As String - Name of the entry to look for in Lines() to get the
''                           starting column position
''
''    HeaderRowNumber As Variant - Row number to look at in Lines(), to look for HeaderName to get the
''                                 starting column position. 0 is the first row. n-1 is the last row.
''
''    DataStartColumnNumber As Variant - An integer to indicate the starting column to read data
''                                       and put it to InputStringArray.
''
'' Returns:
''    A string array of column entries from a 2D string array.
''
'' Examples:
''
'' --- Code
''    Dim TestFolder As String
''    Dim Transition_Array() As String
''    Dim TidyDataRowFile As String
''
''
''    'Indicate path to the test data folder
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    TidyDataRowFile = TestFolder & "TidyTransitionRow.csv"
''
''    'Read the csv files
''    Dim Lines() As String
''    Lines = Utilities.Read_File(TidyDataRowFile)
''
''    Transition_Array = Utilities.Load_Columns_From_2Darray(InputStringArray:=Transition_Array, _
''                                                           Lines:=Lines, _
''                                                           DataStartColumnNumber:=0, _
''                                                           DataStartRowNumber:=1, _
''                                                           Delimiter:=",", _
''                                                           RemoveBlksAndReplicates:=True)
''
''    Dim header_line_index As Long
''    For header_line_index = LBound(Transition_Array) To UBound(Transition_Array)
''        Debug.Print Transition_Array(header_line_index)
''    Next header_line_index
'' ---
Public Function Load_Columns_From_2Darray(ByRef InputStringArray() As String, _
                                          ByRef Lines() As String, _
                                          ByVal DataStartRowNumber As Long, _
                                          ByVal Delimiter As String, _
                                          ByVal RemoveBlksAndReplicates As Boolean, _
                                          Optional ByVal HeaderName As String, _
                                          Optional ByVal HeaderRowNumber As Variant, _
                                          Optional ByVal DataStartColumnNumber As Variant) As String()
    'We are updating the InputStringArray
    'Dim TotalRows As Long
    Dim lines_index As Long
    Dim ArrayLength As Long
    ArrayLength = Utilities.Get_String_Array_Len(InputStringArray)

    'Get column position of a given header name
    Dim HeaderColNumber As Variant
    If Not Trim$(HeaderName) = vbNullString And Not IsMissing(HeaderRowNumber) Then
        HeaderColNumber = Utilities.Get_Header_Col_Position_From_2Darray(Lines:=Lines, _
                                                                         HeaderName:=HeaderName, _
                                                                         HeaderRowNumber:=HeaderRowNumber, _
                                                                         Delimiter:=Delimiter)
    ElseIf Not IsMissing(DataStartColumnNumber) Then
        HeaderColNumber = DataStartColumnNumber
    End If
    
    Dim Transition_Name As String
    Dim InArray As Boolean
    
    For lines_index = DataStartRowNumber To UBound(Lines) - 1
        'Get the Transition_Name and remove the whitespaces
        Transition_Name = Trim$(Split(Lines(lines_index), Delimiter)(HeaderColNumber))
        If RemoveBlksAndReplicates Then
            'Check if the Transition name is not empty and duplicate
            InArray = Utilities.Is_In_Array(Transition_Name, InputStringArray)
            If Len(Transition_Name) <> 0 And Not InArray Then
                ReDim Preserve InputStringArray(ArrayLength)
                InputStringArray(ArrayLength) = Transition_Name
                'Debug.Print InputStringArray(ArrayLength)
                ArrayLength = ArrayLength + 1
            End If
        Else
            ReDim Preserve InputStringArray(ArrayLength)
            InputStringArray(ArrayLength) = Transition_Name
            'Debug.Print InputStringArray(ArrayLength)
            ArrayLength = ArrayLength + 1
        End If
    Next lines_index
    
    Load_Columns_From_2Darray = InputStringArray
    
End Function

'' Function: Read_File
'' --- Code
''  Public Function Read_File(ByVal xFileName As Variant) As String()
'' ---
''
'' Description:
''
'' Read the input file line by line.
''
'' Parameters:
''
''    xFileName As Variant - File path to an input file in csv or tab separated.
''
'' Returns:
''    A string array in which each entry is one row of data separated
''    by a delimiter. For example, in a .csv file, the delimiter is
''    usually ",". If the file has a header and one line of data, the
''    array will look like this.
''
''    - Read_File(0) = "Column1,Column2"
''    - Read_File(1) = "Data1,Data2"
''
'' Examples:
''
'' --- Code
''    Dim SampleAnnotFile As String
''    Dim TestFolder As String
''
''    Dim Lines() As String
''    Dim Delimiter As String
''
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    SampleAnnotFile = TestFolder & "Sample_Annotation_Example.csv"
''
''    Lines = Utilities.Read_File(SampleAnnotFile)
''    Delimiter = Utilities.Get_Delimiter(SampleAnnotFile)
'' ---
Public Function Read_File(ByVal xFileName As Variant) As String()
    ' Load the file into a string.
    'Dim fn As String, whole_file As String
    'Set fso = CreateObject("Scripting.FileSystemObject")
    'whole_file = fso.OpenTextFile(xFileName).ReadAll
    Dim fnum As Variant
    Dim whole_file As Variant
    
    fnum = FreeFile()
    Open xFileName For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum
    
    ' Break the file into lines.
    Read_File = Split(whole_file, vbCrLf)
    
End Function

'' Function: Get_Delimiter
'' --- Code
''  Public Function Get_Delimiter(ByRef xFileName As Variant) As String
'' ---
''
'' Description:
''
'' Get the delimiter of the input file. Currently, we can only
'' identify "," if a .csv file is provided and tab if a .txt file
'' is provided.
''
'' Parameters:
''
''    xFileName As String - File path to the input file in csv or tab separated.
''
'' Returns:
''    A string "," if a .csv file is provided or vbtab if a
''    .txt file is provided.
''
'' Examples:
''
'' --- Code
''    Dim SampleAnnotFile As String
''    Dim TestFolder As String
''
''    Dim Lines() As String
''    Dim Delimiter As String
''
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    SampleAnnotFile = TestFolder & "Sample_Annotation_Example.csv"
''
''    Lines = Utilities.Read_File(SampleAnnotFile)
''    Delimiter = Utilities.Get_Delimiter(SampleAnnotFile)
'' ---
Public Function Get_Delimiter(ByVal xFileName As Variant) As String

    Dim FileExtent As Variant
    FileExtent = Right$(xFileName, Len(xFileName) - InStrRev(xFileName, "."))
    'Get the first line
    If FileExtent = "csv" Then
        Get_Delimiter = ","
    ElseIf FileExtent = "txt" Then
        Get_Delimiter = vbTab
    Else
        MsgBox "Cannot identify delimiter due to unusual file type"
        End
    End If
    
End Function

'' Function: Get_File_Base_Name
'' --- Code
''  Public Function Get_File_Base_Name(ByVal xFileName As Variant) As String
'' ---
''
'' Description:
''
'' Get the base name of a given file path.
''
'' Parameters:
''
''    xFileName As String - File path to the input file in csv or tab separated.
''
'' Returns:
''    A string indicating the base name of a given file path
''
'' Examples:
''
'' --- Code
''    Dim SampleAnnotFile As String
''    Dim TestFolder As String
''    Dim FileName As String
''
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    SampleAnnotFile = TestFolder & "Sample_Annotation_Example.csv"
''
''    FileName = Utilities.Get_File_Base_Name(SampleAnnotFile)
''
''    Debug.Print FileName
'' ---
Public Function Get_File_Base_Name(ByVal xFileName As Variant) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Get_File_Base_Name = fso.GetFileName(xFileName)
End Function

'' Function: Get_Raw_Data_File_Type
'' --- Code
''  Public Function Get_Raw_Data_File_Type(ByRef Lines() As String, _
''                                         ByVal Delimiter As String, _
''                                         ByVal xFileName As String) As String
'' ---
''
'' Description:
''
'' Get the type of an input raw data file. It can be in
'' "AgilentWideForm", "AgilentCompoundForm" or "Sciex"
''
'' Parameters:
''
''    Lines() As String - A 2D string array. We let the number of elements (rows)
''                        of Lines be n. Each element of Lines is
''                        in the form "EntryRow1 Column1,EntryRow1 Column2,
''                        ...,EntryRow1 Column p" where
''                        p is the number of columns.
''
''    Delimiter As String - Delimiter used to separate each entries in
''                          one element of Lines(). For example, the delimiter for
''                          "EntryRow1 Column1,EntryRow1 Column2,...,EntryRow1 Column p"
''                          is ","
''
''    xFileName As String - File path to the input file in csv or tab separated.
''
'' Returns:
''    A string indicating the base name of a given file path
''
'' Examples:
''
'' --- Code
''    Dim Lines() As String
''    Dim Delimiter As String
''    Dim FileName As String
''    Dim RawDataFileType As String
''    Dim TestFolder As String
''    Dim RawDataFile As String
''
''    'Indicate path to the test data folder
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    RawDataFile = TestFolder & "AgilentRawDataTest1.csv"
''
''    Lines = Utilities.Read_File(RawDataFile)
''    Delimiter = Utilities.Get_Delimiter(RawDataFile)
''    FileName = Utilities.Get_File_Base_Name(RawDataFile)
''    RawDataFileType = Utilities.Get_Raw_Data_File_Type(Lines, Delimiter, FileName)
''
''    Debug.Print RawDataFileType
''
''    RawDataFile = TestFolder & "CompoundTableForm.csv"
''
''    Lines = Utilities.Read_File(RawDataFile)
''    Delimiter = Utilities.Get_Delimiter(RawDataFile)
''    FileName = Utilities.Get_File_Base_Name(RawDataFile)
''    RawDataFileType = Utilities.Get_Raw_Data_File_Type(Lines, Delimiter, FileName)
''
''    Debug.Print RawDataFileType
''
''    RawDataFile = TestFolder & "SciExTestData.txt"
''
''    Lines = Utilities.Read_File(RawDataFile)
''    Delimiter = Utilities.Get_Delimiter(RawDataFile)
''    FileName = Utilities.Get_File_Base_Name(RawDataFile)
''    RawDataFileType = Utilities.Get_Raw_Data_File_Type(Lines, Delimiter, FileName)
''
''    Debug.Print RawDataFileType
'' ---
Public Function Get_Raw_Data_File_Type(ByRef Lines() As String, _
                                       ByVal Delimiter As String, _
                                       ByVal xFileName As String) As String
    Dim first_line() As String
    Dim second_line() As String
    'Get the first line
    first_line = Split(Lines(0), Delimiter)
    
    'If sample is in first line, check the second line
    If first_line(0) = "Sample" Then
        If Utilities.Get_String_Array_Len(Lines) > 1 Then
            second_line = Split(Lines(1), Delimiter)
            If Utilities.Is_In_Array("Data File", second_line) Then
                Get_Raw_Data_File_Type = "AgilentWideForm"
            End If
        End If
    ElseIf first_line(0) = "Compound Method" Then
        Get_Raw_Data_File_Type = "AgilentCompoundForm"
    ElseIf first_line(0) = "Sample Name" Then
        Get_Raw_Data_File_Type = "Sciex"
    End If
    
    'Give an error if we are unable to find up where the raw data is coming from
    If Get_Raw_Data_File_Type = vbNullString Then
        MsgBox "Cannot identify the raw data file type (Agilent or SciEx) for " & xFileName
        Exit Function
        'End
    End If
    
End Function

'' Function: Remove_Filter_Settings
'' --- Code
''  Public Sub Remove_Filter_Settings()
'' ---
''
'' Description:
''
'' Remove any filter settings in the active sheet.
''
Public Sub Remove_Filter_Settings()
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
        If ActiveSheet.FilterMode Then
            ActiveSheet.ShowAllData
        End If
    ElseIf ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
End Sub

'' Function: Get_Header_Col_Position
'' --- Code
''  Public Function Get_Header_Col_Position(ByVal HeaderName As String, _
''                                          ByVal HeaderRowNumber As Long, _
''                                          Optional ByVal WorksheetName As String = vbNullString) As Long
'' ---
''
'' Description:
''
'' Get the column position for a given header name and
'' a header row number to look at.
''
'' Parameters:
''
''    HeaderName As String - Name of the header/column name to look for
''
''    HeaderRowNumber As Long - Row position in which the header/column name
''                              could possibly be.
''
''    WorksheetName As String - Name of the worksheet to search for the
''                              header/column name. If it is blank, the
''                              active sheet will be used.
''
'' Returns:
''    An integer indicating the column position for a given header name
''
'' Examples:
''
'' --- Code
''    'Should be 1 as it is the first column
''    Debug.Print Utilities.Get_Header_Col_Position("Transition_Name", 1, "Transition_Name_Annot")
''
''    'Should be 2 as it is the second column
''    Debug.Print Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 1, "Transition_Name_Annot")
'' ---
Public Function Get_Header_Col_Position(ByVal HeaderName As String, _
                                        ByVal HeaderRowNumber As Long, _
                                        Optional ByVal WorksheetName As String = vbNullString) As Long
                                        
    'Get column position of Header Name
    Dim pos As Long
    If WorksheetName = vbNullString Then
        pos = Application.Match(HeaderName, ActiveWorkbook.Worksheets.Item(ActiveSheet.Name).Rows(HeaderRowNumber).Value, False)
        If IsError(pos) Then
            MsgBox HeaderName & " is missing in the headers of" & ActiveSheet.Name & " sheet"
            'Excel resume monitoring the sheet
            Application.EnableEvents = True
        End
    End If
    Else
        pos = Application.Match(HeaderName, ActiveWorkbook.Worksheets.Item(WorksheetName).Rows(HeaderRowNumber).Value, False)
        If IsError(pos) Then
            MsgBox HeaderName & " is missing in the headers of" & WorksheetName & " sheet"
            'Excel resume monitoring the sheet
            Application.EnableEvents = True
            End
        End If
    End If

    Get_Header_Col_Position = pos
End Function

'' Function: Last_Used_Row_Number
'' --- Code
''  Public Function Last_Used_Row_Number() As Long
'' ---
''
'' Description:
''
'' Get the last used row number of the active sheet
''
'' Returns:
''    An integer indicating the row position where the last entry is
''
'' Examples:
''
'' --- Code
''    ' Get the Lists worksheet from the active workbook
''    ' The Lists is a code name
''    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
''    Dim Lists_Worksheet As Worksheet
''
''    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "Lists") = False Then
''        MsgBox ("Sheet Lists is missing")
''        Application.EnableEvents = True
''        Exit Sub
''    End If
''
''    Set Lists_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "Lists")
''
''    Lists_Worksheet.Activate
''
''    Debug.Print Utilities.Last_Used_Row_Number
'' ---
Public Function Last_Used_Row_Number() As Long
    Dim maxRowNumber As Long
    Dim Column As Long
    Dim TotalColumns As Long
    maxRowNumber = 0
    
    'Find the last non-blank cell in row 1
    'Debug.Print ActiveSheet.UsedRange.Address(0, 0)
    'Debug.Print ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    TotalColumns = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    'For each column, find the last used rows and then take the max value
    For Column = 1 To TotalColumns
        If ActiveSheet.Cells(ActiveSheet.Rows.Count, Column).End(xlUp).Row > maxRowNumber Then
            maxRowNumber = ActiveSheet.Cells(ActiveSheet.Rows.Count, Column).End(xlUp).Row
        End If
    Next Column
    
    Last_Used_Row_Number = maxRowNumber
    
End Function

'' Function: Convert_To_Letter
'' --- Code
''  Public Function Convert_To_Letter(ByVal Column_Index As Long) As String
'' ---
''
'' Description:
''
'' Convert column position in integers to alphabets
''
'' Parameters:
''
''    Column_Index As Long - An integer indicating the column position
''
'' Returns:
''    An string representing a column position in alphabets
''
'' Examples:
''
'' --- Code
''    ' Should be A
''    Debug.Print Utilities.Convert_To_Letter(1)
''    ' Should be Z
''    Debug.Print Utilities.Convert_To_Letter(26)
''    ' Should be AA
''    Debug.Print Utilities.Convert_To_Letter(27)
''    ' Should be AZ
''    Debug.Print Utilities.Convert_To_Letter(52)
''    ' Should be BA
''    Debug.Print Utilities.Convert_To_Letter(53)
''    ' Should be SZ
''    Debug.Print Utilities.Convert_To_Letter(520)
'' ---
Public Function Convert_To_Letter(ByVal Column_Index As Long) As String
    'Convert column number values into their equivalent alphabetical characters:
    Dim Alphabet_Index As Long
    Dim Remainder_Index As Long
    Alphabet_Index = Int((Column_Index - 1) / 26)
    Remainder_Index = Column_Index - (Alphabet_Index * 26)
    If Alphabet_Index > 0 Then
        Convert_To_Letter = Chr$(Alphabet_Index + 64)
    End If
    If Remainder_Index > 0 Then
        Convert_To_Letter = Convert_To_Letter & Chr$(Remainder_Index + 64)
    End If
End Function

'' Function: Get_String_Array_Len
'' --- Code
''  Public Function Get_String_Array_Len(ByRef Some_Array As Variant) As Long
'' ---
''
'' Description:
''
'' Get the number of elements in an input string array
''
'' Parameters:
''
''    Some_Array As Variant - An input string array
''
'' Returns:
''    An integer indicating how many elements are in the input string array
''
'' Examples:
''
'' --- Code
''    Dim TestArray As Variant
''    Dim EmptyArray As Variant
''
''    TestArray = Array("11_PQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_PQC_2.d", _
''                      "PQC_PC-PE-SM_01.d", "PQC1_LPC-LPE-PG-PI-PS_01.d", "PQC1_02.d", _
''                      "11_BQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_BQC_2.d", "018_BQC_PQC01")
''    EmptyArray = Array()
''
''    ' Should be 8
''    Debug.Print Utilities.Get_String_Array_Len(TestArray)
''
''    ' Should be 0
''    Debug.Print Utilities.Get_String_Array_Len(EmptyArray)
'' ---
Public Function Get_String_Array_Len(ByRef Some_Array As Variant) As Long
    'Get the length of the array
    If Len(Join(Some_Array, vbNullString)) = 0 Then
        Get_String_Array_Len = 0
    Else
        Get_String_Array_Len = UBound(Some_Array) - LBound(Some_Array) + 1
    End If
End Function

'' Function: Where_In_Array
'' --- Code
''  Public Function Where_In_Array(ByVal valToBeFound As Variant, _
''                                 ByVal arr As Variant) As String()
'' ---
''
'' Description:
''
'' Get the position in the input string array an input element is located
''
'' Parameters:
''
''    valToBeFound As Variant - Input string to be searched in the input array.
''
''    arr As Variant - Input array used to search for the input string
''
'' Returns:
''    A string array indicating the position the input string is
''    located in the input array. If there is no match, the string
''    array will be of length zero.
''
'' Examples:
''
'' --- Code
''    Dim TestArray As Variant
''    Dim Positions() As String
''
''    'Ensure that it works and gives the right position
''    TestArray = Array("Here", "11_PQC-2.d", "Here", "No", "Here", "Here")
''
''    ' A string array of {"0", "2", "4", "5"}
''    Positions = Utilities.Where_In_Array("Here", TestArray)
'' ---
Public Function Where_In_Array(ByVal valToBeFound As Variant, _
                               ByVal arr As Variant) As String()
    'Return the position of where valToBeFound in the arr
    Dim Positions() As String
    Dim ArrayLength As Long
    Dim Index As Long
    ArrayLength = 0
    Index = 0
    Dim element As Variant

    On Error GoTo IsInArrayError:                'array is empty
    For Each element In arr
        'If we have a match, we store the position
        If element = valToBeFound Then
            ReDim Preserve Positions(ArrayLength)
            Positions(ArrayLength) = CStr(Index)
            ArrayLength = ArrayLength + 1
        End If
        Index = Index + 1
    Next element
    'Return the array that stores the occurences
    Where_In_Array = Positions
IsInArrayError:
    On Error GoTo 0

End Function

'' Function: Is_In_Array
'' --- Code
''  Public Function Is_In_Array(ByVal valToBeFound As Variant, _
''                              ByVal arr As Variant) As Boolean
'' ---
''
'' Description:
''
'' Check if an input element is located in the input string array
''
'' Parameters:
''
''    valToBeFound As Variant - Input string to be searched in the input array.
''
''    arr As Variant - Input array used to search for the input string
''
'' Returns:
''    True if an input element is located in the input string array.
''    False otherwise.
''
'' Examples:
''
'' --- Code
''    Dim TestArray As Variant
''    TestArray = Array("Here", "11_PQC-2.d", "Here", "No", "Here", "Here")
''
''    ' Should be True
''    Debug.Print Utilities.Is_In_Array("Here", TestArray)
''    ' Should be True
''    Debug.Print Utilities.Is_In_Array("11_PQC-2.d", TestArray)
''    ' Should be False
''    Debug.Print Utilities.Is_In_Array("NotHere", TestArray)
'' ---
Public Function Is_In_Array(ByVal valToBeFound As Variant, _
                            ByVal arr As Variant) As Boolean
    'DEVELOPER: Ryan Wells (wellsr.com)
    'DESCRIPTION: Function to check if a value is in an array of values
    'INPUT: Pass the function a value to search for and an array of values of any data type.
    'OUTPUT: True if is in array, false otherwise
    Dim element As Variant
    On Error GoTo IsInArrayError:                'array is empty
    For Each element In arr
        If element = valToBeFound Then
            Is_In_Array = True
            Exit Function
        End If
    Next element
    Exit Function
IsInArrayError:
    On Error GoTo 0
    Is_In_Array = False
End Function

'' Function: Quick_Sort
'' --- Code
''  Public Sub Quick_Sort(ByRef ThisArray As Variant)
'' ---
''
'' Description:
''
'' Sort an the input string array in alphabetical order.
''
'' Parameters:
''
''    ThisArray As Variant - Input string array to be sorted in alphabetical order.
''
'' Returns:
''    A string array sorted in alphabetical order.
''
'' Examples:
''
'' --- Code
''    Dim TestArray As Variant
''    TestArray = Array("SM C36:2", "lipid", "Cer d18:1/C16:0")
''    ' Should be {"Cer d18:1/C16:0", "SM C36:2", "lipid"}
''    Utilities.Quick_Sort ThisArray:=TestArray
'' ---
Public Sub Quick_Sort(ByRef ThisArray As Variant)

    'Sort an array alphabetically
    Dim LowerBound As Variant
    Dim UpperBound As Variant
    LowerBound = LBound(ThisArray)
    UpperBound = UBound(ThisArray)

    Utilities.Quick_Sort_Recursive ThisArray, LowerBound, UpperBound

End Sub

Private Sub Quick_Sort_Recursive(ByRef ThisArray As Variant, _
                                 ByVal LowerBound As Variant, _
                                 ByVal UpperBound As Variant)

    'Approximate implementation of https://en.wikipedia.org/wiki/Quicksort
    Dim PivotValue As Variant
    Dim LowerSwap As Variant
    Dim UpperSwap As Variant
    Dim TempItem As Variant
    
    'Zero or 1 item to sort
    If UpperBound - LowerBound < 1 Then Exit Sub

    'Only 2 items to sort
    If UpperBound - LowerBound = 1 Then
        If ThisArray(LowerBound) > ThisArray(UpperBound) Then
            TempItem = ThisArray(LowerBound)
            ThisArray(LowerBound) = ThisArray(UpperBound)
            ThisArray(UpperBound) = TempItem
        End If
        Exit Sub
    End If

    '3 or more items to sort
    PivotValue = ThisArray(Int((LowerBound + UpperBound) / 2))
    ThisArray(Int((LowerBound + UpperBound) / 2)) = ThisArray(LowerBound)
    LowerSwap = LowerBound + 1
    UpperSwap = UpperBound

    Do
        'Find the right LowerSwap
        Do While LowerSwap < UpperSwap And ThisArray(LowerSwap) <= PivotValue
            LowerSwap = LowerSwap + 1
        Loop

        'Find the right UpperSwap
        Do While LowerBound < UpperSwap And ThisArray(UpperSwap) > PivotValue
            UpperSwap = UpperSwap - 1
        Loop
        
        'Swap values if LowerSwap is less than UpperSwap
        If LowerSwap < UpperSwap Then
            TempItem = ThisArray(LowerSwap)
            ThisArray(LowerSwap) = ThisArray(UpperSwap)
            ThisArray(UpperSwap) = TempItem
        End If
    Loop While LowerSwap < UpperSwap
    
    ThisArray(LowerBound) = ThisArray(UpperSwap)
    ThisArray(UpperSwap) = PivotValue

    'Recursively call function
    
    '2 or more items in first section
    If LowerBound < (UpperSwap - 1) Then Quick_Sort_Recursive ThisArray, LowerBound, UpperSwap - 1

    '2 or more items in second section
    If UpperSwap + 1 < UpperBound Then Quick_Sort_Recursive ThisArray, UpperSwap + 1, UpperBound

End Sub

'' Function: Overwrite_Several_Headers
'' --- Code
''  Public Sub Overwrite_Several_Headers(ByRef HeaderNameArray() As String, _
''                                       ByVal HeaderRowNumber As Long, _
''                                       ByVal DataStartRowNumber As Long)
'' ---
''
'' Description:
''
'' Function used to open this overwrite confirmation box
''
'' (see Overwrite_Box_Several_Headers_Example.png)
''
'' If the overwrite button is pressed, data in the columns indicated
'' in HeaderNameArray will be replaced with the updated data.
''
'' Parameters:
''
''    HeaderNameArray() As String - Input string array of headers or column names.
''
''    HeaderRowNumber As Long - An integer indicating the row where the
''                              headers or column names are in the sheet.
''
''    DataStartRowNumber As Long - An integer indicating the starting row
''                                 where the data are.
''
Public Sub Overwrite_Several_Headers(ByRef HeaderNameArray() As String, _
                                     ByVal HeaderRowNumber As Long, _
                                     ByVal DataStartRowNumber As Long)
    'Check with user if data should be overwritten
    Dim TotalRows As Long
    TotalRows = Utilities.Last_Used_Row_Number()
    
    'If there are no entries, overwrite is not needed, leave the sub
    If TotalRows < DataStartRowNumber Then
        Exit Sub
    End If
    
    Overwrite.Show
    Select Case Overwrite.whatsclicked
    Case "Cancel"
        'Excel resume monitoring the sheet
        Application.EnableEvents = True
        End
    Case "Overwrite"
        'To ensure that Filters does not affect the assignment
        Utilities.Remove_Filter_Settings
        'Clear the contents. We do not want to clean the headers
        Dim HeaderNameArrayIndex As Long
        For HeaderNameArrayIndex = 0 To UBound(HeaderNameArray) - LBound(HeaderNameArray)
            Utilities.Clear_Columns HeaderToClear:=HeaderNameArray(HeaderNameArrayIndex), _
                                    HeaderRowNumber:=HeaderRowNumber, _
                                    DataStartRowNumber:=DataStartRowNumber
        Next HeaderNameArrayIndex
    End Select
    Unload Overwrite

End Sub

'' Function: Overwrite_Header
'' --- Code
''  Public Sub Overwrite_Header(ByVal HeaderName As String, _
''                              ByVal HeaderRowNumber As Long, _
''                              ByVal DataStartRowNumber As Long, _
''                              Optional ByVal WorksheetName As String = vbNullString, _
''                              Optional ByVal ClearContent As Boolean = True, _
''                              Optional ByVal Testing As Boolean = False)
'' ---
''
'' Description:
''
'' Function used to open this overwrite confirmation box
''
'' (see Overwrite_Box_One_Header_Example.png)
''
'' If the overwrite button is pressed, data in the columns indicated
'' in HeaderName will be replaced with the updated data.
''
'' Parameters:
''
''    HeaderName As String - Input string indicating the header or column name whose
''                           data to overwrite.
''
''    HeaderRowNumber As Long - An integer indicating the row where the
''                              headers or column names are in the sheet.
''
''    DataStartRowNumber As Long - An integer indicating the starting row
''                                 where the data are.
''
''    WorksheetName As String - Name of the worksheet to search for the
''                              header/column name. If it is blank, the
''                              active sheet will be used.
''
''    ClearContent As Boolean - If True, contents will be cleared before
''                              updating/adding new data.
''
''    Testing As Boolean - When set to False, after the function is used, it will exit the program as this
''                         function is meant to be run alone or as a last/final step.
''                         When set to True, after the function is used, it will exit the function and
''                         other functions can be called
''
Public Sub Overwrite_Header(ByVal HeaderName As String, _
                            ByVal HeaderRowNumber As Long, _
                            ByVal DataStartRowNumber As Long, _
                            Optional ByVal WorksheetName As String = vbNullString, _
                            Optional ByVal ClearContent As Boolean = True, _
                            Optional ByVal Testing As Boolean = False)
                           
                           
    'Activate the correct sheet
    If WorksheetName <> vbNullString Then
        ActiveWorkbook.Worksheets.Item(WorksheetName).Activate
    End If
    
    Dim HeaderColNumber As Long
    HeaderColNumber = Utilities.Get_Header_Col_Position(HeaderName, HeaderRowNumber, _
                                                        WorksheetName:=WorksheetName)
    
    'Check if the header has entries
    Dim TotalRows As Long
    TotalRows = ActiveSheet.Cells(ActiveSheet.Rows.Count, Utilities.Convert_To_Letter(HeaderColNumber)).End(xlUp).Row
    
    'If there are no entries, overwrite is not needed, leave the sub
    If TotalRows < DataStartRowNumber Then
        Exit Sub
    End If
    
    'Show the Overwrite choice box
    If HeaderName <> vbNullString Then
        Overwrite.Are_You_Sure_Message.Caption = "There exists " & HeaderName & " in the sheet. Do you want to overwrite them ?"
    End If
    Overwrite.Show
    
    Select Case Overwrite.whatsclicked
    Case "Cancel"
        'Excel resume monitoring the sheet
        Application.EnableEvents = True
        If Testing = True Then
            Exit Sub
        End If
        End
    Case "Overwrite"
        'To ensure that Filters does not affect the assignment
        Utilities.Remove_Filter_Settings
        
        If ClearContent Then
            'Clear the contents. We do not want to clean the headers
            ActiveSheet.Range(Utilities.Convert_To_Letter(HeaderColNumber) & CStr(DataStartRowNumber) & ":" & Utilities.Convert_To_Letter(HeaderColNumber) & TotalRows).ClearContents
        End If

    End Select
    Unload Overwrite
    
End Sub

'' Function: Load_To_Excel
'' --- Code
''  Public Sub Load_To_Excel(ByRef Data_Array() As String, _
''                           ByVal HeaderName As String, _
''                           ByVal HeaderRowNumber As Long, _
''                           ByVal DataStartRowNumber As Long, _
''                           ByVal MessageBoxRequired As Boolean, _
''                           Optional ByVal WorksheetName As String = vbNullString, _
''                           Optional ByVal NumberFormat As String = "General")
'' ---
''
'' Description:
''
'' Function used to load a string array to an active excel sheet
'' given the header name, row position where the header name
'' is located and the starting row of the data.
''
'' Parameters:
''
''    Data_Array() As String - Input string array to be loaded to excel
''
''    HeaderName As String - Name of the header/column to tell the system
''                           which column to output the data.
''
''    HeaderRowNumber As Long - Input integer indicating the row number
''                              where the header/column name is located.
''
''    DataStartRowNumber As Long - Input integer indicating which row
''                                 to start outputing the data.
''
''    MessageBoxRequired As Boolean - If set to True, a message box
''                                    will appear after outputing the data
''                                    saying how many entries have been loaded.
''
''    WorksheetName As String - Name of the worksheet to search for the
''                              header/column name. If it is blank, the
''                              active sheet will be used.
''
''    NumberFormat As String - Indicate how number loaded to excel will be
''                             expressed as. By default, it is general.
''                             Other options can be found in this
''                             <link: https://support.microsoft.com/en-us/office/available-number-formats-in-excel-0afe8f52-97db-41f1-b972-4b46e9f1e8d2>.
''
'' Examples:
''
'' --- Code
''    ' Get the Transition_Name_Annot worksheet from the active workbook
''    ' The TransitionNameAnnotSheet is a code name
''    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
''    Dim Transition_Name_Annot_Worksheet As Worksheet
''
''    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "TransitionNameAnnotSheet") = False Then
''        MsgBox ("Sheet Transition_Name_Annot is missing")
''        Application.EnableEvents = True
''        Exit Sub
''    End If
''
''    Set Transition_Name_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "TransitionNameAnnotSheet")
''
''    Transition_Name_Annot_Worksheet.Activate
''
''    Dim Transition_Array(3) As String
''    Transition_Array(0) = "SM 44:0"
''    Transition_Array(1) = "SM 44:1"
''    Transition_Array(2) = "SM 46:2"
''    Transition_Array(3) = "SM 46:3"
''
''    Utilities.Load_To_Excel Data_Array:=Transition_Array, _
''                            HeaderName:="Transition_Name", _
''                            HeaderRowNumber:=1, _
''                            DataStartRowNumber:=2, _
''                            MessageBoxRequired:=False
'' ---
Public Sub Load_To_Excel(ByRef Data_Array() As String, _
                         ByVal HeaderName As String, _
                         ByVal HeaderRowNumber As Long, _
                         ByVal DataStartRowNumber As Long, _
                         ByVal MessageBoxRequired As Boolean, _
                         Optional ByVal WorksheetName As String = vbNullString, _
                         Optional ByVal NumberFormat As String = "General")
    
    Dim HeaderColNumber As Long
    HeaderColNumber = Utilities.Get_Header_Col_Position(HeaderName, HeaderRowNumber, _
                                                        WorksheetName:=WorksheetName)
    
    'Assume ISTD_Array is checked to be non-empty by an earlier function
    If UBound(Data_Array) - LBound(Data_Array) + 1 <> 0 Then
        ActiveSheet.Range(Utilities.Convert_To_Letter(HeaderColNumber) & CStr(DataStartRowNumber)).Resize(UBound(Data_Array) + 1) = Application.Transpose(Data_Array)
        'Ensure that the number format is always kept at "General"
        ActiveSheet.Range(Utilities.Convert_To_Letter(HeaderColNumber) & CStr(DataStartRowNumber)).Resize(UBound(Data_Array) + 1).NumberFormat = NumberFormat
        If MessageBoxRequired = True Then
            MsgBox "Loaded " & UBound(Data_Array) + 1 & " " & HeaderName & "."
        End If
    End If
End Sub

'' Function: Clear_Columns
'' --- Code
''  Public Sub Clear_Columns(ByVal HeaderToClear As String, _
''                           ByVal HeaderRowNumber As Long, _
''                           ByVal DataStartRowNumber As Long, _
''                           Optional ByVal ClearFormat As Boolean = False)
'' ---
''
'' Description:
''
'' Function used to clear a column in an active excel sheet
'' given the header name to clear, row position where the header name
'' is located and the starting row of the data.
''
'' Parameters:
''
''    HeaderName As String - Name of the header/column to tell the system
''                           which column to output the data.
''
''    HeaderRowNumber As Long - Input integer indicating the row number
''                              where the header/column name is located.
''
''    DataStartRowNumber As Long - Input integer indicating which row
''                                 to start outputing the data.
''
'' Examples:
''
'' --- Code
''    ' Get the Transition_Name_Annot worksheet from the active workbook
''    ' The TransitionNameAnnotSheet is a code name
''    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
''    Dim Transition_Name_Annot_Worksheet As Worksheet
''
''    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "TransitionNameAnnotSheet") = False Then
''        MsgBox ("Sheet Transition_Name_Annot is missing")
''        Application.EnableEvents = True
''        Exit Sub
''    End If
''
''    Set Transition_Name_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "TransitionNameAnnotSheet")
''
''    Transition_Name_Annot_Worksheet.Activate
''
''    Dim Transition_Array(3) As String
''    Transition_Array(0) = "SM 44:0"
''    Transition_Array(1) = "SM 44:1"
''    Transition_Array(2) = "SM 46:2"
''    Transition_Array(3) = "SM 46:3"
''
''    Utilities.Load_To_Excel Data_Array:=Transition_Array, _
''                            HeaderName:="Transition_Name", _
''                            HeaderRowNumber:=1, _
''                            DataStartRowNumber:=2, _
''                            MessageBoxRequired:=False
''
''    Utilities.Clear_Columns HeaderToClear:="Transition_Name", _
''                            HeaderRowNumber:=1, _
''                            DataStartRowNumber:=2
'' ---
Public Sub Clear_Columns(ByVal HeaderToClear As String, _
                         ByVal HeaderRowNumber As Long, _
                         ByVal DataStartRowNumber As Long, _
                         Optional ByVal ClearFormat As Boolean = False)

    Dim HeaderColNumber As Long
    HeaderColNumber = Utilities.Get_Header_Col_Position(HeaderToClear, HeaderRowNumber)

    'We do not want to clean the headers
    Dim TotalRows As Long
    TotalRows = ActiveSheet.Cells(ActiveSheet.Rows.Count, Utilities.Convert_To_Letter(HeaderColNumber)).End(xlUp).Row
    If TotalRows < DataStartRowNumber Then
        Exit Sub
    End If
    
    If ClearFormat Then
        ActiveSheet.Range(Utilities.Convert_To_Letter(HeaderColNumber) & CStr(DataStartRowNumber) & ":" & Utilities.Convert_To_Letter(HeaderColNumber) & TotalRows).Clear
    Else
        ActiveSheet.Range(Utilities.Convert_To_Letter(HeaderColNumber) & CStr(DataStartRowNumber) & ":" & Utilities.Convert_To_Letter(HeaderColNumber) & TotalRows).ClearContents
    End If
End Sub

'' Function: Clear_DotD_In_Agilent_Data_File
'' --- Code
''  Public Function Clear_DotD_In_Agilent_Data_File(ByRef AgilentDataFile() As String) As String()
'' ---
''
'' Description:
''
'' Remove the ".d" in the AgilentDataFile String Array.
'' The Agilent files from Agilent MassHunter Quant has a
'' "Data File" column which is unique in every row and is usually
'' used as the Sample Name. The "Data File" column usually ends
'' with ".d". This function helps to remove this ".d"
''
'' Parameters:
''
''    AgilentDataFile() As String - A string array in which each
''                                  entry is a data file name that
''                                  ends with ".d"
''
'' Returns:
''    A string array in which each entry has ".d" remove
''
'' Examples:
''
'' --- Code
''    Dim Sample_Name_Array(1) As String
''
''    Sample_Name_Array(0) = "Sample_Name_1.d"
''    Sample_Name_Array(1) = "Sample_Name_2.d"
''
''    Sample_Name_Array = Utilities.Clear_DotD_In_Agilent_Data_File(Sample_Name_Array)
'' ---
Public Function Clear_DotD_In_Agilent_Data_File(ByRef AgilentDataFile() As String) As String()
    Dim AgilentDataFileIndex As Long
    For AgilentDataFileIndex = 0 To Utilities.Get_String_Array_Len(AgilentDataFile) - 1
        AgilentDataFile(AgilentDataFileIndex) = Trim$(Replace(AgilentDataFile(AgilentDataFileIndex), ".d", vbNullString))
    Next AgilentDataFileIndex
    Clear_DotD_In_Agilent_Data_File = AgilentDataFile
End Function

'' Function: Load_Columns_From_Excel
'' --- Code
''  Public Function Load_Columns_From_Excel(ByVal HeaderName As String, _
''                                          ByVal HeaderRowNumber As Long, _
''                                          ByVal DataStartRowNumber As Long, _
''                                          ByVal MessageBoxRequired As Boolean, _
''                                          ByVal RemoveBlksAndReplicates As Boolean, _
''                                          Optional ByVal WorksheetName As String = vbNullString, _
''                                          Optional ByVal IgnoreHiddenRows As Boolean = True, _
''                                          Optional ByVal IgnoreEmptyArray As Boolean = True) As String()
'' ---
''
'' Description:
''
'' Load data from a column in Excel and put it to a string array
''
'' Parameters:
''
''    HeaderName As String - Name of the header/column to tell the system
''                           which column to output the data.
''
''    HeaderRowNumber As Long - Input integer indicating the row number
''                              where the header/column name is located.
''
''    DataStartRowNumber As Long - Input integer indicating which row
''                                 to start outputing the data.
''
''    MessageBoxRequired As Boolean - If set to True, a message box
''                                    will appear after loading the data
''                                    saying how many entries have been loaded.
''
''    RemoveBlksAndReplicates As Boolean - If set to True, the system will remove duplicates
''                                         and blank values in output string array
''
''    WorksheetName As String - Name of the worksheet to search for the
''                              header/column name. If it is blank, the
''                              active sheet will be used.
''
''    IgnoreHiddenRows As Boolean - If Cell is hidden and IgnoreHiddenRows is True,
''                                  skip to the next row.
''
''    IgnoreEmptyArray As Boolean - If IgnoreEmptyArray is True, the system will end
''                                  if it received an empty array as an output and will
''                                  not proceed to any other process that may come after
''                                  this function is utilise.
''
'' Returns:
''    A string array containing the loaded data
''
'' Examples:
''
'' --- Code
''    ' Get the Lists worksheet from the active workbook
''    ' The Lists is a code name
''    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
''    Dim Lists_Worksheet As Worksheet
''
''    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "Lists") = False Then
''        MsgBox ("Sheet Lists is missing")
''        Application.EnableEvents = True
''        Exit Sub
''    End If
''
''    Set Lists_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "Lists")
''
''    Lists_Worksheet.Activate
''
''    Dim Concentration_Unit_Array() As String
''    Concentration_Unit_Array = Utilities.Load_Columns_From_Excel("Concentration_Unit", HeaderRowNumber:=1, _
''                                                                 DataStartRowNumber:=2, MessageBoxRequired:=True, _
''                                                                 RemoveBlksAndReplicates:=True, _
''                                                                 IgnoreHiddenRows:=False, IgnoreEmptyArray:=False)
'' ---
Public Function Load_Columns_From_Excel(ByVal HeaderName As String, _
                                        ByVal HeaderRowNumber As Long, _
                                        ByVal DataStartRowNumber As Long, _
                                        ByVal MessageBoxRequired As Boolean, _
                                        ByVal RemoveBlksAndReplicates As Boolean, _
                                        Optional ByVal WorksheetName As String = vbNullString, _
                                        Optional ByVal IgnoreHiddenRows As Boolean = True, _
                                        Optional ByVal IgnoreEmptyArray As Boolean = True) As String()
    Dim InputStringArray() As String
    Dim TotalRows As Long
    Dim Row_Index As Long
    Dim ArrayLength As Long
    Dim InputWorksheetName As String
    
    InputWorksheetName = WorksheetName
    If InputWorksheetName = vbNullString Then
        InputWorksheetName = ActiveSheet.Name
    End If
    
    'Get column position of Transition_Name_ISTD
    Dim HeaderColNumber As Long
    HeaderColNumber = Utilities.Get_Header_Col_Position(HeaderName, HeaderRowNumber, _
                                                        WorksheetName:=InputWorksheetName)
    
    'Get the total number of rows
    TotalRows = ActiveWorkbook.Worksheets.Item(InputWorksheetName).Cells(ActiveSheet.Rows.Count, Utilities.Convert_To_Letter(HeaderColNumber)).End(xlUp).Row
    ArrayLength = 0
    
    Dim InArray As Boolean
    Dim Entries As String
    'Get the entries
    For Row_Index = DataStartRowNumber To TotalRows
    
        'If Cell is hidden and IgnoreHiddenRows is True, we skip to the next row
        If ActiveWorkbook.Worksheets.Item(InputWorksheetName).Cells(Row_Index, HeaderColNumber).RowHeight <> 0 Or Not IgnoreHiddenRows Then
        
            If RemoveBlksAndReplicates Then
                'Check that it is not empty or has only spaces
                If Not IsEmpty(ActiveWorkbook.Worksheets.Item(InputWorksheetName).Cells(Row_Index, HeaderColNumber)) Then
                    Entries = Trim$(ActiveWorkbook.Worksheets.Item(InputWorksheetName).Cells(Row_Index, HeaderColNumber).Value)
                    InArray = Utilities.Is_In_Array(Entries, InputStringArray)
                    If Len(Entries) <> 0 And Not InArray Then
                        ReDim Preserve InputStringArray(ArrayLength)
                        InputStringArray(ArrayLength) = Entries
                        'Debug.Print InputStringArray(ArrayLength)
                        ArrayLength = ArrayLength + 1
                    End If
                End If
            Else
                ReDim Preserve InputStringArray(ArrayLength)
                InputStringArray(ArrayLength) = CStr(ActiveWorkbook.Worksheets.Item(InputWorksheetName).Cells(Row_Index, HeaderColNumber))
                'Debug.Print InputStringArray(ArrayLength)
                ArrayLength = ArrayLength + 1
            End If
        End If
    
    Next Row_Index
        
    'If we have an empty array
    If Len(Join(InputStringArray, vbNullString)) = 0 Then
        If MessageBoxRequired Then
            MsgBox "Loaded " & 0 & " " & HeaderName & "."
        End If
        If Not IgnoreEmptyArray Then
            'Excel resume monitoring the sheet
            Application.EnableEvents = True
            End
        End If
    End If
    
    Load_Columns_From_Excel = InputStringArray
    'Debug.Print ArrayLength
    
End Function

'' Function: Check_Sheet_Code_Name_Exists
'' --- Code
''  Public Function Check_Sheet_Code_Name_Exists(ByVal InputWorkbook As Workbook, _
''                                               ByVal InputCodeName As String) As Boolean
'' ---
''
'' Description:
''
'' Check if the sheet code name exists in this workbook
''
'' Parameters:
''
''    InputWorkbook As Workbook, - Name of the workbook to check for the code name
''
''    InputCodeName As String - Name of the code name to search in the workbook to
''                              see if it exists.
''
'' Returns:
''    Returns True if the code name exists. False otherwise.
''
'' Examples:
''
'' --- Code
''    ' Should return True
''    Debug.Print Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "Lists")
''    ' Should return False
''    Debug.Print Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "Does not Exists")
'' ---
Public Function Check_Sheet_Code_Name_Exists(ByVal InputWorkbook As Workbook, _
                                             ByVal InputCodeName As String) As Boolean
     
    Dim Sheet As Worksheet
     
    For Each Sheet In InputWorkbook.Sheets
        'Debug.Print oSht.CodeName
        If Sheet.CodeName = InputCodeName Then
            Check_Sheet_Code_Name_Exists = True
            Exit For
        End If
    Next
     
End Function

'' Function: Get_Sheet_By_Code_Name
'' --- Code
''  Public Function Check_Sheet_Code_Name_Exists(ByVal InputWorkbook As Workbook, _
''                                               ByVal InputCodeName As String) As Boolean
'' ---
''
'' Description:
''
'' Get the name of the sheet from the sheet's code name
''
'' Parameters:
''
''    InputWorkbook As Workbook, - Name of the workbook to check for the code name
''
''    InputCodeName As String - Name of the code name to search in the workbook to
''                              see if it exists.
''
'' Returns:
''    Returns the name of the sheet from the sheet's code name
''
'' Examples:
''
'' --- Code
''    Dim ISTD_Annot_Worksheet As Worksheet
''    Set ISTD_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "ISTDAnnotSheet")
''
''    ' Should be "ISTDAnnotSheet"
''    Debug.Print ISTD_Annot_Worksheet.CodeName
'' ---
Public Function Get_Sheet_By_Code_Name(ByVal InputWorkbook As Workbook, _
                                       ByVal InputCodeName As String) As Worksheet

    ' Check if the sheet whose code name exists
    Dim Sheet As Worksheet
    For Each Sheet In InputWorkbook.Worksheets
        If Sheet.CodeName = InputCodeName Then
            Set Get_Sheet_By_Code_Name = Sheet
            Exit Function
        End If
    Next
    
End Function

'' Function: Get_Folder
'' --- Code
''  Public Function Get_Folder() As String
'' ---
''
'' Description:
''
'' Get the file path of a folder selected by the user.
''
'' The function will first open a pop up box to allow users
'' to select a folder.
''
'' (see Utilities_Get_Folder.png)
''
'' After the folder is selected, the file path of the folder is returned
''
'' Returns:
''    Returns the file path of a folder selected by the user
Public Function Get_Folder() As String
    'https://stackoverflow.com/questions/26392482/vba-excel-to-prompt-user-response-to-select-folder-and-return-the-path-as-string
    Dim Folder As FileDialog
    Dim Selected_Item As String
    Set Folder = Application.FileDialog(msoFileDialogFolderPicker)
    With Folder
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        Selected_Item = .SelectedItems.Item(1)
    End With
NextCode:
    Get_Folder = Selected_Item
    Set Folder = Nothing
End Function
