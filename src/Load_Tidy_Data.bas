Attribute VB_Name = "Load_Tidy_Data"
Option Explicit
'@Folder("Load Data Functions")

'' Function: Get_Sample_Name_Array_Tidy
''
'' Description:
''
'' Get Sample Names from an input data file in tabular form, put them into
'' a string array.
''
'' Parameters:
''
''    TidyDataFilesArray() As String - A string array of File path to a tabular/tidy data file.
''                                     Eg. {FilePath 1, FilePath 2}
''
''    MS_File_Array() As String - A string array of Data File Names to be loaded to the Data_File_Name
''                                column of the Sample_Annot sheet. If this function is used more than
''                                once, this array will be appended.
''                                Eg. {Input_File_Name_1, Input_File_Name_1, Input_File_Name_2}
''
''    DataFileType As String - File type of the input tabular/tidy data file
''
''    SampleProperty As String - Choose "Read as column variables" if transition names are the column name.
''                               Choose "Read as row observations" if transition names are row entries.
''
''    StartingRowNum As Integer - Starting row number to read the tabular data
''
''    StartingColumnNum As Integer - Starting column number to read the tabular data
''
'' Returns:
''    A string array of Sample Names and a string array of Data File Names
''
'' Examples:
''
'' --- Code
''    Dim TestFolder As String
''    Dim TidyDataColumnFiles As String
''    Dim TidyDataFilesArray() As String
''    Dim MS_File_Array() As String
''    Dim Sample_Name_Array_from_Tidy_Data() As String
''
''    'Indicate path to the test data folder
''    TestFolder = ThisWorkbook.Path & "\Testdata\"
''    TidyDataColumnFiles = TestFolder & "TidySampleColumn.csv"
''    TidyDataFilesArray = Split(TidyDataColumnFiles, ";")
''
''    Sample_Name_Array_from_Tidy_Data = Load_Tidy_Data.Get_Sample_Name_Array_Tidy(TidyDataFilesArray:=TidyDataFilesArray, _
''                                                                                 MS_File_Array:=MS_File_Array, _
''                                                                                 DataFileType:="csv", _
''                                                                                 SampleProperty:="Read as column variables", _
''                                                                                 StartingRowNum:=1, _
''                                                                                 StartingColumnNum:=2)
'' ---
Public Function Get_Sample_Name_Array_Tidy(ByRef TidyDataFilesArray() As String, _
                                           ByRef MS_File_Array() As String, _
                                           ByVal DataFileType As String, _
                                           ByVal SampleProperty As String, _
                                           ByVal StartingRowNum As Long, _
                                           ByVal StartingColumnNum As Long) As String()
 
    'Initialise the Sample Name Array
    Dim Sample_Name_Array() As String
    
    'When no file is selected
    If TypeName(TidyDataFilesArray) = "Boolean" Then
        End
    End If
    On Error GoTo 0
    
    Dim TidyDataFile As Variant
      
    For Each TidyDataFile In TidyDataFilesArray
    
        Dim Sample_Name_SubArray() As String
        Dim MS_File_SubArray() As String
        Dim SubarrayLength As Long
        SubarrayLength = 0
        
        Select Case DataFileType
        Case "csv"
            'Read the csv files
            Dim Lines() As String
            Dim FileName As String
            Lines = Utilities.Read_File(TidyDataFile)
            FileName = Utilities.Get_File_Base_Name(TidyDataFile)
            
            If SampleProperty = "Read as column variables" Then
            Sample_Name_SubArray = Utilities.Load_Rows_From_2Darray(InputStringArray:=Sample_Name_SubArray, _
                                                                    Lines:=Lines, _
                                                                    DataStartColumnNumber:=StartingColumnNum - 1, _
                                                                    Delimiter:=",", _
                                                                    RemoveBlksAndReplicates:=True, _
                                                                    DataStartRowNumber:=StartingRowNum - 1)
            ElseIf SampleProperty = "Read as row observations" Then
            Sample_Name_SubArray = Utilities.Load_Columns_From_2Darray(InputStringArray:=Sample_Name_SubArray, _
                                                                       Lines:=Lines, _
                                                                       DataStartColumnNumber:=StartingColumnNum - 1, _
                                                                       DataStartRowNumber:=StartingRowNum - 1, _
                                                                       Delimiter:=",", _
                                                                       RemoveBlksAndReplicates:=True)
            End If
        End Select
        
        'Update the subarray to the original arrays
        Dim Sample_Name_SubArray_Index As Long
        Sample_Name_Array = Utilities.Concantenate_String_Arrays(Sample_Name_Array, Sample_Name_SubArray)
        SubarrayLength = 0
            
        For Sample_Name_SubArray_Index = 0 To Utilities.Get_String_Array_Len(Sample_Name_SubArray) - 1
            ReDim Preserve MS_File_SubArray(SubarrayLength)
            MS_File_SubArray(Sample_Name_SubArray_Index) = FileName
            SubarrayLength = SubarrayLength + 1
        Next Sample_Name_SubArray_Index
        MS_File_Array = Utilities.Concantenate_String_Arrays(MS_File_Array, MS_File_SubArray)
        
        Erase Sample_Name_SubArray
        Erase MS_File_SubArray
    
    Next TidyDataFile
    Get_Sample_Name_Array_Tidy = Sample_Name_Array
End Function

'Private Function Get_Transition_Array_Tidy_Excel(ByVal TidyDataFiles As String, _
'                                                 ByVal TransitionProperty As String, _
'                                                 ByVal StartingRowNum As Long, _
'                                                 ByVal StartingColumnNum As Long) As String()
'
'    'File are taken from userfrom Load_Transition_Name_Tidy
'    'Hence they must exists and joined together by ;
'    Dim TidyDataFilesArray() As String
'    TidyDataFilesArray = Split(TidyDataFiles, ";")
'
'    'When no file is selected
'    If TypeName(TidyDataFilesArray) = "Boolean" Then
'        End
'    End If
'    On Error GoTo 0
'
'    'Check if file is truly excel
'    Dim fso As Object
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Dim TidyDataFile As Variant
'
'    For Each TidyDataFile In TidyDataFilesArray
'        If Not fso.GetExtensionName(TidyDataFile) Like "*xls*" Then
'            MsgBox TidyDataFile & " is not an excel file"
'            Exit Function
'        End If
'    Next TidyDataFile
'
'    'Initialise the Transition Array
'    Dim Transition_Array() As String
'    Dim ArrayLength As Long
'    'ArrayLength = 0
'
'    For Each TidyDataFile In TidyDataFilesArray
'
'        'We update the array length for Transition_Array
'        ArrayLength = Utilities.Get_String_Array_Len(Transition_Array)
'
'        'Ensure that the excel file do not pop up when
'        'the code to reading the excel file
'        Application.ScreenUpdating = False
'
'        'Open the excel file
'        Dim Source_Workbook As Workbook
'        Set Source_Workbook = Workbooks.Open(TidyDataFile, UpdateLinks:=True, ReadOnly:=True)
'
'        Dim TotalRows As Long
'        TotalRows = Source_Workbook.Worksheets.Item("sheet1").Cells.Item(Source_Workbook.Worksheets.Item("sheet1").Rows.Count, 1).End(xlUp).Row
'
'        'Debug.Print TotalRows
'        Dim Row_Index As Long
'        Dim Transition_Name As String
'        Dim InArray As Boolean
'
'        For Row_Index = 2 To TotalRows
'            Transition_Name = Source_Workbook.Worksheets.Item("sheet1").Cells.Item(Row_Index, 1).Value
'            'Debug.Print Transition_Name
'
'            'Check if the Transition name is not empty and duplicate
'            InArray = Utilities.Is_In_Array(Transition_Name, Transition_Array)
'            If Len(Transition_Name) <> 0 And Not InArray Then
'                    ReDim Preserve Transition_Array(ArrayLength)
'                    Transition_Array(ArrayLength) = Transition_Name
'                    ArrayLength = ArrayLength + 1
'            End If
'
'        Next Row_Index
'
'        ' Close the source file but do not save the source file
'        Source_Workbook.Close SaveChanges:=False
'        Set Source_Workbook = Nothing
'
'        Application.ScreenUpdating = True
'
'    Next TidyDataFile
'    Get_Transition_Array_Tidy_Excel = Transition_Array
'
'End Function

'' Function: Get_Transition_Array_Tidy_CSV
''
'' Description:
''
'' Get Transition Names from an input data file in tabular form, put them into
'' a string array.
''
'' Parameters:
''
''    TidyDataFiles As String - File path to a tabular/tidy data file in csv.
''                              If multiple files are required, the different
''                              file path must be separated by ";"
''                              Eg. {FilePath 1};{FilePath 2}
''
''    TransitionProperty As String - Choose "Read as column variables" if transition names are the column name.
''                                   Choose "Read as row observations" if transition names are row entries.
''
''    StartingRowNum As Integer - Starting row number to read the tabular data
''
''    StartingColumnNum As Integer - Starting column number to read the tabular data
''
'' Returns:
''    A string array of Transition Names.
''
'' Examples:
''
'' --- Code
''   Dim TestFolder As String
''   Dim TidyDataColumnFiles As String
''   Dim Transition_Array() As String
''
''   ' Indicate path to the test data folder
''   TestFolder = ThisWorkbook.Path & "\Testdata\"
''   TidyDataColumnFiles = TestFolder & "TidyTransitionColumn.csv"
''
''   ' Get the transition names
''   Transition_Array = Load_Tidy_Data.Get_Transition_Array_Tidy_CSV(TidyDataFiles:=TidyDataColumnFiles, _
''                                                                   TransitionProperty:="Read as column variables", _
''                                                                   StartingRowNum:=1, _
''                                                                   StartingColumnNum:=2)
'' ---
Public Function Get_Transition_Array_Tidy_CSV(ByVal TidyDataFiles As String, _
                                              ByVal TransitionProperty As String, _
                                              ByVal StartingRowNum As Long, _
                                              ByVal StartingColumnNum As Long) As String()
                                               
    'File are taken from userfrom Load_Transition_Name_Tidy
    'Hence they must exists and joined together by ;
    Dim TidyDataFilesArray() As String
    TidyDataFilesArray = Split(TidyDataFiles, ";")
    
    'When no file is selected
    If TypeName(TidyDataFilesArray) = "Boolean" Then
        End
    End If
    On Error GoTo 0
    
    
    'Check if file is truly excel
    Dim TidyDataFile As Variant
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each TidyDataFile In TidyDataFilesArray
        If Not fso.GetExtensionName(TidyDataFile) Like "*csv*" Then
            MsgBox TidyDataFile & " is not a csv file"
            Exit Function
        End If
    Next TidyDataFile

    'Initialise the Transition Array
    Dim Transition_Array() As String
    'Dim ArrayLength As Long
    'ArrayLength = 0
      
    For Each TidyDataFile In TidyDataFilesArray
    
        'Read the csv files
        Dim Lines() As String
        'Dim FileName As String
        Lines = Utilities.Read_File(TidyDataFile)
        'FileName = Utilities.Get_File_Base_Name(TidyDataFile)
        
        If TransitionProperty = "Read as column variables" Then
            Transition_Array = Utilities.Load_Rows_From_2Darray(InputStringArray:=Transition_Array, _
                                                                Lines:=Lines, _
                                                                DataStartColumnNumber:=StartingColumnNum - 1, _
                                                                Delimiter:=",", _
                                                                RemoveBlksAndReplicates:=True, _
                                                                DataStartRowNumber:=StartingRowNum - 1)
            'Transition_Array = Utilities.Load_Rows_From_2Darray(Transition_Array, Lines(), _
            '                                                    DataStartColumnNumber:=StartingColumnNum - 1, _
            '                                                    Delimiter:=",", _
            '                                                    RemoveBlksAndReplicates:=True, _
            '                                                    RowName:="Sample_Name", _
            '                                                    RowNameNumber:=0)
        ElseIf TransitionProperty = "Read as row observations" Then
            Transition_Array = Utilities.Load_Columns_From_2Darray(InputStringArray:=Transition_Array, _
                                                                   Lines:=Lines, _
                                                                   DataStartColumnNumber:=StartingColumnNum - 1, _
                                                                   DataStartRowNumber:=StartingRowNum - 1, _
                                                                   Delimiter:=",", _
                                                                   RemoveBlksAndReplicates:=True)
        End If
    Next TidyDataFile
    
    
    Get_Transition_Array_Tidy_CSV = Transition_Array
                                                 
End Function

'' Function: Get_Transition_Array_Tidy
''
'' Description:
''
'' Get Transition Names from an input data file in tabular form, put them into
'' a string array.
''
'' Parameters:
''
''    TidyDataFiles As String - File path to a tabular/tidy data file in csv.
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
'' Returns:
''    A string array of Transition Names.
''
'' Examples:
''
'' --- Code
''   Dim TestFolder As String
''   Dim TidyDataColumnFiles As String
''   Dim Transition_Array() As String
''
''   ' Indicate path to the test data folder
''   TestFolder = ThisWorkbook.Path & "\Testdata\"
''   TidyDataColumnFiles = TestFolder & "TidyTransitionColumn.csv"
''
''   ' Get the transition names
''   Transition_Array = Load_Tidy_Data.Get_Transition_Array_Tidy(TidyDataFiles:=TidyDataColumnFiles, _
''                                                               DataFileType:="csv", _
''                                                               TransitionProperty:="Read as column variables", _
''                                                               StartingRowNum:=1, _
''                                                               StartingColumnNum:=2)
'' ---
Public Function Get_Transition_Array_Tidy(ByVal TidyDataFiles As String, _
                                          ByVal DataFileType As String, _
                                          ByVal TransitionProperty As String, _
                                          ByVal StartingRowNum As Long, _
                                          ByVal StartingColumnNum As Long) As String()
                                          
    ' Get the Transition_Name_Annot worksheet from the active workbook
    ' The TransitionNameAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Transition_Name_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "TransitionNameAnnotSheet") = False Then
        MsgBox ("Sheet Transition_Name_Annot is missing")
        Application.EnableEvents = True
        Exit Function
    End If
    
    Set Transition_Name_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "TransitionNameAnnotSheet")
      
    Transition_Name_Annot_Worksheet.Activate
    
    Dim Transition_Array() As String
    
    'If the Load Annotation button is clicked
    Select Case DataFileType
'    Case "Excel"
'        Transition_Array = Get_Transition_Array_Tidy_Excel(TidyDataFiles:=TidyDataFiles, _
'                                                           TransitionProperty:=TransitionProperty, _
'                                                           StartingRowNum:=StartingRowNum, _
'                                                           StartingColumnNum:=StartingColumnNum)
    Case "csv"
        Transition_Array = Get_Transition_Array_Tidy_CSV(TidyDataFiles:=TidyDataFiles, _
                                                         TransitionProperty:=TransitionProperty, _
                                                         StartingRowNum:=StartingRowNum, _
                                                         StartingColumnNum:=StartingColumnNum)
    End Select
    
    Get_Transition_Array_Tidy = Transition_Array
End Function
