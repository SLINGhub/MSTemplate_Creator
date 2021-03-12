Attribute VB_Name = "Utilities"
Public Function Concantenate_String_Arrays(Array1() As String, Array2() As String) As String()
    'Update the Sample Name Array
    Dim Array1Length As Long
    Dim Array2Length As Long
    Array1Length = Len(Join(Array1, ""))
    Array2Length = Len(Join(Array2, ""))
    
    If Array1Length > 0 And Array2Length > 0 Then
        Concantenate_String_Arrays = Split(Join(Array1, ",") & "," & Join(Array2, ","), ",")
    ElseIf Array1Length > 0 Then
        Concantenate_String_Arrays = Array1
    ElseIf Array2Length > 0 Then
        Concantenate_String_Arrays = Array2
    Else
        MsgBox "Two arrays cannot be empty"
        Exit Function
    End If
End Function

Public Function Get_RowName_Position_From_2Darray(ByRef Lines() As String, RowName As String, _
                                                  RowNameNumber As Variant, Delimiter As String) As Variant
    
    Dim row_name() As String
    Get_RowName_Position_From_2Darray = Null
    
    For i = LBound(Lines) To UBound(Lines) - 1
        'Get the Row_Name and remove the whitespaces
        row_name = Split(Lines(i), Delimiter)
        'Debug.Print Trim(row_name(RowNameNumber))
        If Trim(row_name(RowNameNumber)) = RowName Then
            Get_RowName_Position_From_2Darray = i
            Exit For
        End If
    Next i
    
    If IsNull(Get_RowName_Position_From_2Darray) Then
        MsgBox RowName & " is missing in the input file "
        End
    End If
    
End Function

Public Function Load_Rows_From_2Darray(ByRef strArray() As String, ByRef Lines() As String, _
                                       DataStartColumnNumber As Integer, Delimiter As String, _
                                       RemoveBlksAndReplicates As Boolean, _
                                       Optional ByVal RowName As String, _
                                       Optional ByVal RowNameNumber As Variant, _
                                       Optional ByVal DataStartRowNumber As Variant) As String()
    
                                     
    'Get column position of a given header name
    Dim RowNamePosition As Variant
    If Not Trim(RowName) = vbNullString And Not IsMissing(RowNameNumber) Then
        RowNamePosition = Utilities.Get_RowName_Position_From_2Darray(Lines(), _
                                                                      RowName:=RowName, _
                                                                      RowNameNumber:=RowNameNumber, _
                                                                      Delimiter:=Delimiter)
    ElseIf Not IsMissing(DataStartRowNumber) Then
        RowNamePosition = DataStartRowNumber
    End If
    
    'We just look at the one row the user indicates
    Dim row_line() As String
    row_line = Split(Lines(RowNamePosition), Delimiter)
    'We update the array length of Transition_Array
    ArrayLength = Utilities.StringArrayLen(strArray)
    
    For i = DataStartColumnNumber To UBound(row_line)
        'Get the Transition_Name and remove the whitespaces
        Transition_Name = Trim(row_line(i))
        
        If RemoveBlksAndReplicates Then
            'Check if the Transition name is not empty and duplicate
            InArray = Utilities.IsInArray(Transition_Name, strArray)
            If Len(Transition_Name) <> 0 And Not InArray Then
                ReDim Preserve strArray(ArrayLength)
                strArray(ArrayLength) = Transition_Name
                'Debug.Print strArray(ArrayLength)
                ArrayLength = ArrayLength + 1
            End If
        Else
            ReDim Preserve strArray(ArrayLength)
            strArray(ArrayLength) = Transition_Name
            'Debug.Print strArray(ArrayLength)
            ArrayLength = ArrayLength + 1
        End If
        
    Next i
    
    Load_Rows_From_2Darray = strArray
    
End Function

Public Function Get_Header_Col_Position_From_2Darray(ByRef Lines() As String, HeaderName As String, _
                                                     HeaderRowNumber As Variant, Delimiter As String) As Variant
    
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

Public Function Load_Columns_From_2Darray(ByRef strArray() As String, ByRef Lines() As String, _
                                          DataStartRowNumber As Integer, Delimiter As String, _
                                          RemoveBlksAndReplicates As Boolean, _
                                          Optional ByVal HeaderName As String, _
                                          Optional ByVal HeaderRowNumber As Variant, _
                                          Optional ByVal DataStartColumnNumber As Variant) As String()
    'We are updating the strArray
    'Dim TotalRows As Long
    Dim i As Long
    Dim ArrayLength As Long
    ArrayLength = Utilities.StringArrayLen(strArray)

    'Get column position of a given header name
    Dim HeaderColNumber As Variant
    If Not Trim(HeaderName) = vbNullString And Not IsMissing(HeaderRowNumber) Then
        HeaderColNumber = Utilities.Get_Header_Col_Position_From_2Darray(Lines(), HeaderName, HeaderRowNumber, Delimiter)
    ElseIf Not IsMissing(DataStartColumnNumber) Then
        HeaderColNumber = DataStartColumnNumber
    End If
    
    For i = DataStartRowNumber To UBound(Lines) - 1
        'Get the Transition_Name and remove the whitespaces
        Transition_Name = Trim(Split(Lines(i), Delimiter)(HeaderColNumber))
        If RemoveBlksAndReplicates Then
            'Check if the Transition name is not empty and duplicate
            InArray = Utilities.IsInArray(Transition_Name, strArray)
            If Len(Transition_Name) <> 0 And Not InArray Then
                ReDim Preserve strArray(ArrayLength)
                strArray(ArrayLength) = Transition_Name
                'Debug.Print strArray(ArrayLength)
                ArrayLength = ArrayLength + 1
            End If
        Else
            ReDim Preserve strArray(ArrayLength)
            strArray(ArrayLength) = Transition_Name
            'Debug.Print strArray(ArrayLength)
            ArrayLength = ArrayLength + 1
        End If
    Next i
    
    Load_Columns_From_2Darray = strArray
    
End Function

Public Function Read_File(xFileName As Variant) As String()
    ' Load the file into a string.
    fnum = FreeFile
    Open xFileName For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum
    
    ' Break the file into lines.
    Read_File = Split(whole_file, vbCrLf)
    
End Function

Public Function Get_File_Base_Name(xFileName As Variant) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Get_File_Base_Name = fso.GetFileName(xFileName)
End Function

Public Function Get_Raw_Data_File_Type(ByRef Lines() As String, Delimiter As String, xFileName As String) As String
    Dim first_line() As String
    Dim second_line() As String
    'Get the first line
    first_line = Split(Lines(0), Delimiter)
    
    'If sample is in first line, check the second line
    If first_line(0) = "Sample" Then
        If Utilities.StringArrayLen(Lines) > 1 Then
            second_line = Split(Lines(1), Delimiter)
            If Utilities.IsInArray("Data File", second_line) Then
                Get_Raw_Data_File_Type = "AgilentWideForm"
            End If
        End If
    ElseIf first_line(0) = "Compound Method" Then
        Get_Raw_Data_File_Type = "AgilentCompoundForm"
    ElseIf first_line(0) = "Sample Name" Then
        Get_Raw_Data_File_Type = "Sciex"
    End If
    
    'Give an error if we are unable to find up where the raw data is coming from
    If Get_Raw_Data_File_Type = "" Then
        MsgBox "Cannot identify the raw data file type (Agilent or SciEx) for " & xFileName
        Exit Function
    End If
    
End Function

Public Sub RemoveFilterSettings()
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
        If ActiveSheet.FilterMode Then
            ActiveSheet.ShowAllData
        End If
    ElseIf ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
End Sub

Public Function Get_Header_Col_Position(HeaderName As String, HeaderRowNumber As Integer) As Integer
    'Get column position of Header Name
    Dim pos As Integer
    pos = Application.Match(HeaderName, Rows(HeaderRowNumber).Value, False)
    If IsError(pos) Then
        MsgBox HeaderName & " is missing in the headers of" & ActiveSheet.Name & " sheet"
        'Excel resume monitoring the sheet
        Application.EnableEvents = True
        End
    End If
    Get_Header_Col_Position = pos
End Function

Public Function LastUsedRowNumber() As Long
    Dim maxRowNumber As Long
    Dim TotalColumns As Long
    maxRowNumber = 0
    
    'Find the last non-blank cell in row 1
    'Debug.Print ActiveSheet.UsedRange.Address(0, 0)
    'Debug.Print ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    TotalColumns = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    'For each column, find the last used rows and then take the max value
    For i = 1 To TotalColumns
        If Cells(Rows.Count, i).End(xlUp).Row > maxRowNumber Then
            maxRowNumber = Cells(Rows.Count, i).End(xlUp).Row
        End If
    Next i
    
    LastUsedRowNumber = maxRowNumber
    
End Function

Public Function ConvertToLetter(iCol As Integer) As String
    'Convert column number values into their equivalent alphabetical characters:
    Dim iAlpha As Integer
    Dim iRemainder As Integer
    iAlpha = Int((iCol - 1) / 26)
    iRemainder = iCol - (iAlpha * 26)
    If iAlpha > 0 Then
        ConvertToLetter = Chr(iAlpha + 64)
    End If
    If iRemainder > 0 Then
        ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
    End If
End Function

Public Function StringArrayLen(Some_Array As Variant) As Integer
    'Get the length of the array
    If Len(Join(Some_Array, "")) = 0 Then
        StringArrayLen = 0
    Else
        StringArrayLen = UBound(Some_Array) - LBound(Some_Array) + 1
    End If
End Function

Public Function WhereInArray(valToBeFound As Variant, arr As Variant) As String()
    'Return the position of where valToBeFound in the arr
    Dim Positions() As String
    Dim ArrayLength As Integer
    Dim Index As Integer
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
    WhereInArray = Positions
IsInArrayError:
    On Error GoTo 0

End Function

Public Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function to check if a value is in an array of values
    'INPUT: Pass the function a value to search for and an array of values of any data type.
    'OUTPUT: True if is in array, false otherwise
    Dim element As Variant
    On Error GoTo IsInArrayError:                'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
    Exit Function
IsInArrayError:
    On Error GoTo 0
    IsInArray = False
End Function

Public Sub QuickSort(ByRef ThisArray As Variant)

    'Sort an array alphabetically
    Dim LowerBound, UpperBound
    LowerBound = LBound(ThisArray)
    UpperBound = UBound(ThisArray)

    QuickSortRecursive ThisArray, LowerBound, UpperBound

End Sub

Private Sub QuickSortRecursive(ByRef ThisArray, ByVal LowerBound, ByVal UpperBound)

    'Approximate implementation of https://en.wikipedia.org/wiki/Quicksort
    Dim PivotValue, LowerSwap, UpperSwap, TempItem

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
        While LowerSwap < UpperSwap And ThisArray(LowerSwap) <= PivotValue
            LowerSwap = LowerSwap + 1
        Wend

        'Find the right UpperSwap
        While LowerBound < UpperSwap And ThisArray(UpperSwap) > PivotValue
            UpperSwap = UpperSwap - 1
        Wend
        
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
    If LowerBound < (UpperSwap - 1) Then QuickSortRecursive ThisArray, LowerBound, UpperSwap - 1

    '2 or more items in second section
    If UpperSwap + 1 < UpperBound Then QuickSortRecursive ThisArray, UpperSwap + 1, UpperBound

End Sub

Public Sub OverwriteSeveralHeaders(HeaderNameArray() As String, HeaderRowNumber As Integer, DataStartRowNumber As Integer)
    'Check with user if data should be overwritten
    Dim TotalRows As Long
    TotalRows = Utilities.LastUsedRowNumber()
    
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
        Utilities.RemoveFilterSettings
        'Clear the contents. We do not want to clean the headers
        Dim HeaderName As String
        For i = 0 To UBound(HeaderNameArray) - LBound(HeaderNameArray)
            Call Utilities.Clear_Columns(HeaderNameArray(i), HeaderRowNumber, DataStartRowNumber)
        Next i
    End Select
    Unload Overwrite

End Sub

Public Sub OverwriteHeader(HeaderName As String, HeaderRowNumber As Integer, _
                           DataStartRowNumber As Integer, Optional ByVal Testing As Boolean = False)
    
    Dim HeaderColNumber As Integer
    HeaderColNumber = Utilities.Get_Header_Col_Position(HeaderName, HeaderRowNumber)
    
    'Check if the header has entries
    Dim TotalRows As Long
    TotalRows = Cells(Rows.Count, ConvertToLetter(HeaderColNumber)).End(xlUp).Row
    
    'If there are no entries, overwrite is not needed, leave the sub
    If TotalRows < DataStartRowNumber Then
        Exit Sub
    End If
    
    'Show the Overwrite choice box
    If HeaderName <> "" Then
        Overwrite.Label1.Caption = "There exists " & HeaderName & " in the sheet. Do you want to overwrite them ?"
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
        Utilities.RemoveFilterSettings
        'Clear the contents. We do not want to clean the headers
        Range(ConvertToLetter(HeaderColNumber) & CStr(DataStartRowNumber) & ":" & ConvertToLetter(HeaderColNumber) & TotalRows).ClearContents
    End Select
    Unload Overwrite
    
End Sub

Public Sub Load_To_Excel(ByRef Data_Array() As String, HeaderName As String, HeaderRowNumber As Integer, DataStartRowNumber As Integer, MessageBoxRequired As Boolean, Optional ByVal NumberFormat As String = "General")
    
    Dim HeaderColNumber As Integer
    HeaderColNumber = Utilities.Get_Header_Col_Position(HeaderName, HeaderRowNumber)
    
    'Assume ISTD_Array is checked to be non-empty by an earlier function
    If UBound(Data_Array) - LBound(Data_Array) + 1 <> 0 Then
        Range(ConvertToLetter(HeaderColNumber) & CStr(DataStartRowNumber)).Resize(UBound(Data_Array) + 1) = Application.Transpose(Data_Array)
        'Ensure that the number format is always kept at "General"
        Range(ConvertToLetter(HeaderColNumber) & CStr(DataStartRowNumber)).Resize(UBound(Data_Array) + 1).NumberFormat = NumberFormat
        If MessageBoxRequired Then
            MsgBox "Loaded " & UBound(Data_Array) + 1 & " " & HeaderName & "."
        End If
    End If
End Sub

Public Sub Clear_Columns(HeaderToClear As String, HeaderRowNumber As Integer, DataStartRowNumber As Integer, Optional ByVal ClearFormat As Boolean = False)

    Dim HeaderColNumber As Integer
    HeaderColNumber = Utilities.Get_Header_Col_Position(HeaderToClear, HeaderRowNumber)

    'We do not want to clean the headers
    Dim TotalRows As Long
    TotalRows = Cells(Rows.Count, ConvertToLetter(HeaderColNumber)).End(xlUp).Row
    If TotalRows < DataStartRowNumber Then
        Exit Sub
    End If
    
    If ClearFormat Then
        Range(ConvertToLetter(HeaderColNumber) & CStr(DataStartRowNumber) & ":" & ConvertToLetter(HeaderColNumber) & TotalRows).Clear
    Else
        Range(ConvertToLetter(HeaderColNumber) & CStr(DataStartRowNumber) & ":" & ConvertToLetter(HeaderColNumber) & TotalRows).ClearContents
    End If
End Sub

Public Function Load_Columns_From_Excel(HeaderName As String, HeaderRowNumber As Integer, _
                                        DataStartRowNumber As Integer, _
                                        MessageBoxRequired As Boolean, _
                                        RemoveBlksAndReplicates As Boolean, _
                                        Optional ByVal IgnoreHiddenRows As Boolean = True, _
                                        Optional ByVal IgnoreEmptyArray As Boolean = True) As String()
    Dim strArray() As String
    Dim TotalRows As Long
    Dim i As Long
    Dim ArrayLength As Long
    
    'Get column position of Transition_Name_ISTD
    Dim HeaderColNumber As Integer
    HeaderColNumber = Utilities.Get_Header_Col_Position(HeaderName, HeaderRowNumber)
    
    'Get the total number of rows
    TotalRows = Cells(Rows.Count, ConvertToLetter(HeaderColNumber)).End(xlUp).Row
    ArrayLength = 0
    
    'Get the entries
    For i = DataStartRowNumber To TotalRows
    
        'If Cell is hidden and IgnoreHiddenRows is True, we skip to the next row
        If Cells(i, HeaderColNumber).RowHeight <> 0 Or Not IgnoreHiddenRows Then
        
            If RemoveBlksAndReplicates Then
                'Check that it is not empty or has only spaces
                If Not IsEmpty(Cells(i, HeaderColNumber)) Then
                    entries = Trim(Cells(i, HeaderColNumber).Value)
                    InArray = Utilities.IsInArray(entries, strArray)
                    If Len(entries) <> 0 And Not InArray Then
                        ReDim Preserve strArray(ArrayLength)
                        strArray(ArrayLength) = entries
                        'Debug.Print strArray(ArrayLength)
                        ArrayLength = ArrayLength + 1
                    End If
                End If
            Else
                ReDim Preserve strArray(ArrayLength)
                strArray(ArrayLength) = CStr(Cells(i, HeaderColNumber))
                'Debug.Print strArray(ArrayLength)
                ArrayLength = ArrayLength + 1
            End If
        End If
    
    Next
        
    'If we have an empty array
    If Len(Join(strArray, "")) = 0 Then
        If MessageBoxRequired Then
            MsgBox "Loaded " & 0 & " " & HeaderName & "."
        End If
        If Not IgnoreEmptyArray Then
            'Excel resume monitoring the sheet
            Application.EnableEvents = True
            End
        End If
    End If
    
    Load_Columns_From_Excel = strArray
    'Debug.Print ArrayLength
    
End Function


