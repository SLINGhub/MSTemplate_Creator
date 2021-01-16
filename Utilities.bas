Attribute VB_Name = "Utilities"
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

On Error GoTo IsInArrayError: 'array is empty
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
On Error GoTo IsInArrayError: 'array is empty
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

Public Sub OverwriteHeader(HeaderName As String, HeaderRowNumber As Integer, DataStartRowNumber As Integer, Optional ByVal Testing As Boolean = False)
    
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
