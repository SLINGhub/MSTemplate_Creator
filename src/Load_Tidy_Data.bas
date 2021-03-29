Attribute VB_Name = "Load_Tidy_Data"
Public Function Get_Sample_Name_Array_Tidy(ByRef TidyDataFilesArray() As String, _
                                           ByRef MS_File_Array() As String, _
                                           DataFileType As String, _
                                           SampleProperty As String, _
                                           StartingRowNum As Integer, _
                                           StartingColumnNum As Integer) As String()
 
    'Initialise the Sample Name Array
    Dim Sample_Name_Array() As String
    
    'When no file is selected
    If TypeName(TidyDataFilesArray) = "Boolean" Then
        End
    End If
    On Error GoTo 0
      
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
            Sample_Name_SubArray = Utilities.Load_Rows_From_2Darray(Sample_Name_SubArray, Lines(), _
                                                                    DataStartColumnNumber:=StartingColumnNum - 1, _
                                                                    Delimiter:=",", _
                                                                    RemoveBlksAndReplicates:=True, _
                                                                    DataStartRowNumber:=StartingRowNum - 1)
            ElseIf SampleProperty = "Read as row observations" Then
            Sample_Name_SubArray = Utilities.Load_Columns_From_2Darray(Sample_Name_SubArray, Lines, _
                                                                       DataStartColumnNumber:=StartingColumnNum - 1, _
                                                                       DataStartRowNumber:=StartingRowNum - 1, _
                                                                       Delimiter:=",", _
                                                                       RemoveBlksAndReplicates:=True)
            End If
        End Select
        
        'Update the subarray to the original arrays
        Sample_Name_Array = Utilities.Concantenate_String_Arrays(Sample_Name_Array, Sample_Name_SubArray)
        SubarrayLength = 0
            
        For i = 0 To Utilities.StringArrayLen(Sample_Name_SubArray) - 1
            ReDim Preserve MS_File_SubArray(SubarrayLength)
            MS_File_SubArray(i) = FileName
            SubarrayLength = SubarrayLength + 1
        Next i
        MS_File_Array = Utilities.Concantenate_String_Arrays(MS_File_Array, MS_File_SubArray)
        
        Erase Sample_Name_SubArray
        Erase MS_File_SubArray
    
    Next TidyDataFile
    Get_Sample_Name_Array_Tidy = Sample_Name_Array
End Function

'TODO include sheet name
Private Function Get_Transition_Array_Tidy_Excel(TidyDataFiles As String, _
                                                 TransitionProperty As String, _
                                                 StartingRowNum As Integer, _
                                                 StartingColumnNum As Integer) As String()
                                                  
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
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each TidyDataFile In TidyDataFilesArray
        If Not fso.GetExtensionName(TidyDataFile) Like "*xls*" Then
            MsgBox TidyDataFile & " is not an excel file"
            Exit Function
        End If
    Next TidyDataFile
    
    'Initialise the Transition Array
    Dim Transition_Array() As String
    Dim ArrayLength As Long
    ArrayLength = 0
      
    For Each TidyDataFile In TidyDataFilesArray
    
        'We update the array length for Transition_Array
        ArrayLength = Utilities.StringArrayLen(Transition_Array)
    
        'Ensure that the excel file do not pop up when
        'the code to reading the excel file
        Application.ScreenUpdating = False
        
        'Open the excel file
        Dim src As Workbook
        Set src = Workbooks.Open(TidyDataFile, UpdateLinks = True, ReadOnly = True)

        Dim TotalRows As Long
        TotalRows = src.Worksheets("sheet1").Cells(Rows.Count, 1).End(xlUp).Row
        
        'Debug.Print TotalRows
        
        For i = 2 To TotalRows
            Transition_Name = Cells(i, 1).Value
            'Debug.Print Transition_Name
            
            'Check if the Transition name is not empty and duplicate
            InArray = Utilities.IsInArray(Transition_Name, Transition_Array)
            If Len(Transition_Name) <> 0 And Not InArray Then
                    ReDim Preserve Transition_Array(ArrayLength)
                    Transition_Array(ArrayLength) = Transition_Name
                    ArrayLength = ArrayLength + 1
            End If

        Next i
            
        ' Close the source file but do not save the source file
        src.Close (SaveChanges = False)
        Set src = Nothing
        
        Application.ScreenUpdating = True
        
    Next TidyDataFile
    Get_Transition_Array_Tidy_Excel = Transition_Array
                                           
End Function

Private Function Get_Transition_Array_Tidy_CSV(TidyDataFiles As String, _
                                               TransitionProperty As String, _
                                               StartingRowNum As Integer, _
                                               StartingColumnNum As Integer) As String()
                                               
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
    Dim ArrayLength As Long
    ArrayLength = 0
      
    For Each TidyDataFile In TidyDataFilesArray
    
        'Read the csv files
        Dim Lines() As String
        Dim FileName As String
        Lines = Utilities.Read_File(TidyDataFile)
        FileName = Utilities.Get_File_Base_Name(TidyDataFile)
        
        If TransitionProperty = "Read as column variables" Then
            Transition_Array = Utilities.Load_Rows_From_2Darray(Transition_Array, Lines(), _
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
            Transition_Array = Utilities.Load_Columns_From_2Darray(Transition_Array, Lines, _
                                                                   DataStartColumnNumber:=StartingColumnNum - 1, _
                                                                   DataStartRowNumber:=StartingRowNum - 1, _
                                                                   Delimiter:=",", _
                                                                   RemoveBlksAndReplicates:=True)
        End If
    Next TidyDataFile
    
    
    Get_Transition_Array_Tidy_CSV = Transition_Array
                                                 
End Function

Public Function Get_Transition_Array_Tidy(TidyDataFiles As String, _
                                          DataFileType As String, _
                                          TransitionProperty As String, _
                                          StartingRowNum As Integer, _
                                          StartingColumnNum As Integer) As String()

    Sheets("Transition_Name_Annot").Activate
    
    Dim Transition_Array() As String
    
    'If the Load Annotation button is clicked
    Select Case DataFileType
    Case "Excel"
        Transition_Array = Get_Transition_Array_Tidy_Excel(TidyDataFiles:=TidyDataFiles, _
                                                           TransitionProperty:=TransitionProperty, _
                                                           StartingRowNum:=StartingRowNum, _
                                                           StartingColumnNum:=StartingColumnNum)
    Case "csv"
        Transition_Array = Get_Transition_Array_Tidy_CSV(TidyDataFiles:=TidyDataFiles, _
                                                         TransitionProperty:=TransitionProperty, _
                                                         StartingRowNum:=StartingRowNum, _
                                                         StartingColumnNum:=StartingColumnNum)
    End Select
    
    Get_Transition_Array_Tidy = Transition_Array
End Function
