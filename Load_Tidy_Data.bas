Attribute VB_Name = "Load_Tidy_Data"
'TODO include sheet name
Private Function Get_Transition_Array_Tidy_Excel(TidyDataFiles As String, _
                                                 TransitionProperty As String, _
                                                 StartingRowNum As Integer, _
                                                 StartingColumnNum As Integer) As String()
                                                 
    Debug.Print TransitionProperty
    Debug.Print StartingRowNum
    Debug.Print StartingColumnNum
    
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
            'We just look at the one row the user indicates
            Dim transition_line() As String
            transition_line = Split(Lines(StartingRowNum - 1), ",")
            
            'We update the array length of Transition_Array
            ArrayLength = Utilities.StringArrayLen(Transition_Array)
            
            For i = StartingColumnNum - 1 To UBound(transition_line)
                'Remove the whitespace
                Transition_Name = Trim(transition_line(i))
                'Check if the Transition name is not empty and duplicate
                InArray = Utilities.IsInArray(Transition_Name, Transition_Array)
                If Len(Transition_Name) <> 0 And Not InArray Then
                    ReDim Preserve Transition_Array(ArrayLength)
                    Transition_Array(ArrayLength) = Transition_Name
                    ArrayLength = ArrayLength + 1
                End If
            Next i
        ElseIf TransitionProperty = "Read as row observations" Then
            Transition_Array = Utilities.Load_Columns_From_2Darray(Transition_Array, Lines, _
                                                                   DataStartColumnNumber:=StartingColumnNum - 1, _
                                                                   DataStartRowNumber:=StartingRowNum - 1, _
                                                                   Delimiter:=",", _
                                                                   MessageBoxRequired:=True, _
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
