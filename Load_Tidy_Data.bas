Attribute VB_Name = "Load_Tidy_Data"
Private Function GetFileBaseName(xFileName As Variant) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileBaseName = fso.GetFileName(xFileName)
End Function

Public Function Get_Transition_Array_Tidy(TidyDataFiles As String) As String()

    Sheets("Transition_Name_Annot").Activate
    'File are taken from userfrom Load_Transition_Name_Tidy
    'Hence they must exists and joined together by ;
    Dim TidyDataFilesArray() As String
    TidyDataFilesArray = Split(TidyDataFiles, ";")

    'When no file is selected
    If TypeName(TidyDataFilesArray) = "Boolean" Then
        End
    End If
    On Error GoTo 0
    
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
        Debug.Print xFileName
        Set src = Workbooks.Open(TidyDataFile, UpdateLinks = True, ReadOnly = True)

        Dim TotalRows As Long
        TotalRows = src.Worksheets("sheet1").Cells(Rows.Count, 1).End(xlUp).Row
        
        Debug.Print TotalRows
        
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
        
        
        'FileExtent = Right(xFileName, Len(xFileName) - InStrRev(xFileName, "."))
        'Debug.Print FileExtent
        
        ' Close the source file but do not save the source file
        src.Close (SaveChanges = False)
        Set src = Nothing
        
        Application.ScreenUpdating = True
        
    Next TidyDataFile
    Get_Transition_Array_Tidy = Transition_Array
End Function
