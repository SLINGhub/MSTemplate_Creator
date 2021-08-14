Attribute VB_Name = "Transition_Name_Annot"
Public Function Get_Sorted_Transition_Array_Raw(RawDataFiles As String) As String()
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

Public Function Get_Sorted_Transition_Array_Tidy(TidyDataFiles As String, _
                                                 DataFileType As String, _
                                                 TransitionProperty As String, _
                                                 StartingRowNum As Integer, _
                                                 StartingColumnNum As Integer) As String()
                                                 
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

Public Sub VerifyISTD(ByRef Transition_Array() As String, ByRef ISTD_Array() As String, _
                      Optional ByVal MessageBoxRequired As Boolean = True, _
                      Optional ByVal Testing As Boolean = False)
                      
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
    TotalRows = Cells(Rows.Count, ConvertToLetter(Transition_Name_ColNumber)).End(xlUp).Row
    
    'For i = 0 To UBound(Transition_Array) - LBound(Transition_Array)
    For i = 0 To TotalRows - 2
        If Cells(i + 2, Transition_Name_ISTD_ColNumber).Value = "" Then
            Cells(i + 2, Transition_Name_ColNumber).Interior.Color = RGB(204, 255, 204)
            Cells(i + 2, Transition_Name_ISTD_ColNumber).Interior.Color = RGB(255, 255, 153)
        Else
            InArray = Utilities.IsInArray(Cells(i + 2, Transition_Name_ISTD_ColNumber).Value, Transition_Array)
            If Not InArray Then
                ReDim Preserve InvalidISTD(ArrayLength)
                InvalidISTD(ArrayLength) = Cells(i + 2, Transition_Name_ISTD_ColNumber).Value
                ArrayLength = ArrayLength + 1
            Else
                'If the value is a valid ISTD, remove the yellow background
                Cells(i + 2, Transition_Name_ColNumber).Interior.Color = RGB(204, 255, 204)
                Cells(i + 2, Transition_Name_ISTD_ColNumber).Interior.Color = RGB(204, 255, 204)
            End If
        End If
    Next i
    
    'For i = 0 To UBound(ISTD_Array) - LBound(ISTD_Array)
    '    'If the value is blank, remove the yellow background
    '    If ISTD_Array(i) = "" Then
    '        Cells(i + 2, Transition_Name_ColNumber).Interior.Color = xlNone
    '        Cells(i + 2, Transition_Name_ISTD_ColNumber).Interior.Color = xlNone
    '    Else
    '        InArray = Utilities.IsInArray(ISTD_Array(i), Transition_Array)
    '        If Not InArray Then
    '            ReDim Preserve InvalidISTD(ArrayLength)
    '            InvalidISTD(ArrayLength) = ISTD_Array(i)
    '            ArrayLength = ArrayLength + 1
    '        Else
    '            'If the value is a valid ISTD, remove the yellow background
    '            Cells(i + 2, Transition_Name_ColNumber).Interior.Color = xlNone
    '            Cells(i + 2, Transition_Name_ISTD_ColNumber).Interior.Color = xlNone
    '        End If
    '    End If
    'Next i
    
    If Utilities.StringArrayLen(InvalidISTD) <> 0 Then
        'Put the invalid ISTD in the list box to be displayed
        For i = 0 To UBound(InvalidISTD) - LBound(InvalidISTD)
            Invalid_ISTD_MsgBox.Invalid_ISTD_ListBox.AddItem InvalidISTD(i)
        Next i
        Invalid_ISTD_MsgBox.Show
        If Testing Then
            Exit Sub
        Else
            'Excel resume monitoring the sheet
            Application.EnableEvents = True
            End
        End If
    Else
        If MessageBoxRequired Then
            MsgBox ("All ISTD entries can be found in the column Transition_Name")
        End If
    End If
    
    
End Sub

