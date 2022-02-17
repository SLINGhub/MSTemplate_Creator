Attribute VB_Name = "Transition_Name_Annot"
Option Explicit
'@Folder("Transition_Name_Annot Functions")
'@IgnoreModule IntegerDataType

'' Function: Get_Sorted_Transition_Array_Raw
'' --- Code
''  Public Function Get_Sorted_Transition_Array_Raw(ByRef RawDataFiles As String) As String()
'' ---
''
'' Description:
''
'' Get Transition Names from an input raw data file, put them into
'' a string array and sort them in alphabetical order.
''
'' Parameters:
''
''    RawDataFiles As String - File path to a Raw Data (Agilent) File in csv.
''                             If multiple files are required, the different
''                             file path must be separated by ";"
''                             Eg. {FilePath 1};{FilePath 2}
''
'' Returns:
''    A string array of Transition Names sorted in alphabetical order.
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
''   Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array_Raw(RawDataFiles:=RawDataFiles)
'' ---
Public Function Get_Sorted_Transition_Array_Raw(ByRef RawDataFiles As String) As String()
    Dim Transition_Array() As String
    Transition_Array = Load_Raw_Data.Get_Transition_Array_Raw(RawDataFiles:=RawDataFiles)

    'Leave the program if we have an empty array
    'If Len(Join(Transition_Array, "")) = 0 Then
    '    MsgBox "Could not find any Transition Names"
    '    Exit Function
    'End If
    
    'If there is no data loaded, stop the process
    If Utilities.StringArrayLen(Transition_Array) = CLng(0) Then
        Exit Function
    End If
    
    'Sort the array
    QuickSort ThisArray:=Transition_Array
    Get_Sorted_Transition_Array_Raw = Transition_Array
End Function

'' Function: Get_Sorted_Transition_Array_Tidy
'' --- Code
''  Public Function Get_Sorted_Transition_Array_Tidy(ByRef TidyDataFiles As String, _
''                                                   ByRef DataFileType As String, _
''                                                   ByRef TransitionProperty As String, _
''                                                   ByRef StartingRowNum As Integer, _
''                                                   ByRef StartingColumnNum As Integer) As String()
'' ---
''
'' Description:
''
'' Get Transition Names from an input data file in tabular form, put them into
'' a string array and sort them in alphabetical order.
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
'' Returns:
''    A string array of Transition Names sorted in alphabetical order.
''
'' Examples:
''
'' --- Code
''   Dim TestFolder As String
''   Dim TidyDataRowFiles As String
''   Dim TidyDataColumnFiles As String
''   Dim Transition_Array() As String
''
''   ' Indicate path to the test data folder
''   TestFolder = ThisWorkbook.Path & "\Testdata\"
''   TidyDataRowFiles = TestFolder & "TidyTransitionRow.csv"
''   TidyDataColumnFiles = TestFolder & "TidyTransitionColumn.csv"
''
''   ' Get the transition names from tidy data file with transitons as column variables
''   Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array_Tidy(TidyDataFiles:=TidyDataColumnFiles, _
''                                                                             DataFileType:="csv", _
''                                                                             TransitionProperty:="Read as column variables", _
''                                                                             StartingRowNum:=1, _
''                                                                             StartingColumnNum:=2)
''
''   'Get the transition names from tidy data file with transitons as row observations
''   Transition_Array = Transition_Name_Annot.Get_Sorted_Transition_Array_Tidy(TidyDataFiles:=TidyDataRowFiles, _
''                                                                             DataFileType:="csv", _
''                                                                             TransitionProperty:="Read as row observations", _
''                                                                             StartingRowNum:=2, _
''                                                                             StartingColumnNum:=1)
''
'' ---
Public Function Get_Sorted_Transition_Array_Tidy(ByRef TidyDataFiles As String, _
                                                 ByRef DataFileType As String, _
                                                 ByRef TransitionProperty As String, _
                                                 ByRef StartingRowNum As Integer, _
                                                 ByRef StartingColumnNum As Integer) As String()
                                                 
    Dim Transition_Array() As String
    Transition_Array = Load_Tidy_Data.Get_Transition_Array_Tidy(TidyDataFiles:=TidyDataFiles, _
                                                                DataFileType:=DataFileType, _
                                                                TransitionProperty:=TransitionProperty, _
                                                                StartingRowNum:=StartingRowNum, _
                                                                StartingColumnNum:=StartingColumnNum)
                                                                
    'Leave the program if we have an empty array
    'If Len(Join(Transition_Array, "")) = 0 Then
    '    MsgBox "Could not find any Transition Names"
    '    Exit Function
    'End If
    
    'If there is no data loaded, stop the process
    If Utilities.StringArrayLen(Transition_Array) = CLng(0) Then
        Exit Function
    End If
    
    'Sort the array
    QuickSort ThisArray:=Transition_Array
    Get_Sorted_Transition_Array_Tidy = Transition_Array
End Function

'' Function: Verify_ISTD
'' --- Code
''  Public Sub Verify_ISTD(ByRef Transition_Array() As String, ByRef ISTD_Array() As String, _
''                         Optional ByVal MessageBoxRequired As Boolean = True, _
''                         Optional ByVal Testing As Boolean = False)
'' ---
''
'' Description:
''
'' Verify if the entries in column Transition_Annot_ISTD is valid.
'' Input is valid if the entires can also be found in the column Transition_Annot
''
'' If input is valid, both entries in the Transition_Annot and Transition_Annot_ISTD
'' columns will turn green
''
'' If input is empty, the entry in the Transition_Annot column will turn green while
'' the entry in the Transition_Annot_ISTD will turn yellow
''
'' If input is in valid, both entries in the Transition_Annot and Transition_Annot_ISTD
'' columns will turn white
''
'' (see Transition_Name_Annot_Verify_ISTD_Cell_Colour_Change.png)
''
'' In addition, the following message box will appear
''
'' (see Transition_Name_Annot_Verify_ISTD_Invalid_ISTD_Message.png)
''
'' If all entries are valid, the following message box will appear if MessageBoxRequired
'' is set to True.
''
'' (see Transition_Name_Annot_Verify_ISTD_All_Valid_ISTD_Message.png)
''
'' Parameters:
''
''    Transition_Array() As String - String array containing the transition names
''
''    ISTD_Array() As String - String array containing the transition name internal standard (ISTD)
''
''    MessageBoxRequired As Boolean - When set to True, the following pop up box will appear
''
'' (see Transition_Name_Annot_Verify_ISTD_All_Valid_ISTD_Message.png)
''
''    Testing As Boolean - When set to False, after the function is used, it will exit the program as this
''                         function is meant to be run alone or as a last/final step.
''                         When set to True, after the function is used, it will exit the function and
''                         other functions can be called
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
''    Dim ISTD_Array(3) As String
''    ISTD_Array(0) = "SM 44:1"
''    ISTD_Array(1) = "SM 44:2"
''    ISTD_Array(2) = vbNullString
''    ISTD_Array(3) = "SM 44:1"
''
''    Utilities.Load_To_Excel Data_Array:=Transition_Array, _
''                            HeaderName:="Transition_Name", _
''                            HeaderRowNumber:=1, _
''                            DataStartRowNumber:=2, _
''                            MessageBoxRequired:=False
''
''    Utilities.Load_To_Excel Data_Array:=ISTD_Array, _
''                            HeaderName:="Transition_Name_ISTD", _
''                            HeaderRowNumber:=1, _
''                            DataStartRowNumber:=2, _
''                            MessageBoxRequired:=False
''
''    Transition_Name_Annot.Verify_ISTD Transition_Array:=Transition_Array, _
''                                      ISTD_Array:=ISTD_Array, _
''                                      MessageBoxRequired:=False, _
''                                      Testing:=True
''
'' ---
Public Sub Verify_ISTD(ByRef Transition_Array() As String, ByRef ISTD_Array() As String, _
                       Optional ByVal MessageBoxRequired As Boolean = True, _
                       Optional ByVal Testing As Boolean = False)
                      
    ' Get the Transition_Name_Annot worksheet from the active workbook
    ' The TransitionNameAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Transition_Name_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "TransitionNameAnnotSheet") = False Then
        MsgBox ("Sheet Transition_Name_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set Transition_Name_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "TransitionNameAnnotSheet")
      
    Transition_Name_Annot_Worksheet.Activate
                      
    'Check if ISTD are valid (found in the column transition name)
    'ISTD Array is deprecated and not used in this function
    'If not, tell users which one is the problem
    'Assume that ISTD column exists in the same sheet and are on the same column level
    Dim InvalidISTD() As String
    Dim ArrayLength As Long
    ArrayLength = 0
    
    'Get column position of the headers
    Dim Transition_Name_ColNumber As Integer
    Dim Transition_Name_ISTD_ColNumber As Integer
    Transition_Name_ColNumber = Utilities.Get_Header_Col_Position("Transition_Name", 1)
    Transition_Name_ISTD_ColNumber = Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 1)
    
    'Get the number of entries in the "Transition_Name" column (including the title)
    Dim TotalRows As Long
    TotalRows = Transition_Name_Annot_Worksheet.Cells.Item(Transition_Name_Annot_Worksheet.Rows.Count, ConvertToLetter(Transition_Name_ColNumber)).End(xlUp).Row
    
    Dim rowIndex As Integer
    Dim InArray As Boolean
    'For rowIndex = 0 To UBound(Transition_Array) - LBound(Transition_Array)
    For rowIndex = 0 To TotalRows - 2
        ' If there is no Transition Name ISTD on that row, colour the Transition Name green and the Transition Name ISTD yellow
        If Transition_Name_Annot_Worksheet.Cells.Item(rowIndex + 2, Transition_Name_ISTD_ColNumber).Value = vbNullString Then
            Transition_Name_Annot_Worksheet.Cells.Item(rowIndex + 2, Transition_Name_ColNumber).Interior.Color = RGB(204, 255, 204)
            Transition_Name_Annot_Worksheet.Cells.Item(rowIndex + 2, Transition_Name_ISTD_ColNumber).Interior.Color = RGB(255, 255, 153)
        Else
            ' Check if the ISTD is valid
            InArray = Utilities.IsInArray(Transition_Name_Annot_Worksheet.Cells.Item(rowIndex + 2, Transition_Name_ISTD_ColNumber).Value, Transition_Array)
            If Not InArray Then
                ReDim Preserve InvalidISTD(ArrayLength)
                InvalidISTD(ArrayLength) = Transition_Name_Annot_Worksheet.Cells.Item(rowIndex + 2, Transition_Name_ISTD_ColNumber).Value
                ArrayLength = ArrayLength + 1
            Else
                'If the value is a valid ISTD, colour both the Transition Name green and the Transition Name ISTD green
                Transition_Name_Annot_Worksheet.Cells.Item(rowIndex + 2, Transition_Name_ColNumber).Interior.Color = RGB(204, 255, 204)
                Transition_Name_Annot_Worksheet.Cells.Item(rowIndex + 2, Transition_Name_ISTD_ColNumber).Interior.Color = RGB(204, 255, 204)
            End If
        End If
    Next rowIndex
        
    If Utilities.StringArrayLen(InvalidISTD) <> 0 Then
        'Put the invalid ISTD in the list box to be displayed
        For rowIndex = 0 To UBound(InvalidISTD) - LBound(InvalidISTD)
            Invalid_ISTD_MsgBox.Invalid_ISTD_ListBox.AddItem InvalidISTD(rowIndex)
        Next rowIndex
        Invalid_ISTD_MsgBox.Show
        If Testing Then
            Exit Sub
        Else
            'Excel resume monitoring the sheet
            'Invalid_ISTD_MsgBox.Show
            Application.EnableEvents = True
            End
        End If
    Else
        If MessageBoxRequired Then
            MsgBox ("All ISTD entries can be found in the column Transition_Name")
        End If
    End If
    
    
End Sub

