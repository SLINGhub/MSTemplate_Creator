Attribute VB_Name = "ColourTracker"
'Sheet ISTD_Annot Functions
Public Function ISTDCalculationChecker(ByVal Target As Range)
    Sheets("ISTD_Annot").Activate
    
    'Application.EnableEvents = True
    
    'Get column position of Transition_Name_ISTD
    Dim ISTDHeaderColNumber As Integer
    Dim ISTD_Conc_ng_ColNumber As Integer
    Dim ISTD_MW_ColNumber As Integer
    Dim ISTD_Conc_nM_ColNumber As Integer
    Dim ISTD_Custom_Unit_ColNumber As Integer
    ISTDHeaderColNumber = Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 2)
    ISTD_Conc_ng_ColNumber = Utilities.Get_Header_Col_Position("ISTD_Conc_[ng/mL]", 3)
    ISTD_MW_ColNumber = Utilities.Get_Header_Col_Position("ISTD_[MW]", 3)
    ISTD_Conc_nM_ColNumber = Utilities.Get_Header_Col_Position("ISTD_Conc_[nM]", 3)
    ISTD_Custom_Unit_ColNumber = Utilities.Get_Header_Col_Position("Custom_Unit", 2)
    
    Dim ISTDHeaderColLetter As String
    Dim ISTD_Conc_ng_ColLetter As String
    Dim ISTD_MW_ColLetter As String
    Dim ISTD_Conc_nM_ColLetter As String
    Dim ISTD_Custom_Unit_ColLetter As String
    ISTDHeaderColLetter = Utilities.ConvertToLetter(ISTDHeaderColNumber)
    ISTD_Conc_ng_ColLetter = Utilities.ConvertToLetter(ISTD_Conc_ng_ColNumber)
    ISTD_MW_ColLetter = Utilities.ConvertToLetter(ISTD_MW_ColNumber)
    ISTD_Conc_nM_ColLetter = Utilities.ConvertToLetter(ISTD_Conc_nM_ColNumber)
    ISTD_Custom_Unit_ColLetter = Utilities.ConvertToLetter(ISTD_Custom_Unit_ColNumber)
    
    Dim RelatedRange As String
    RelatedRange = ISTDHeaderColLetter & ":" & ISTD_Custom_Unit_ColLetter
        
    If Not Intersect(Range(RelatedRange), Target) Is Nothing Then
        'Debug.Print Intersect(Range(RelatedRange), Target).Address
        
        'If rows are deleted, leave the function
        'Debug.Print Target.Address
        'Debug.Print Target.EntireRow.Address
        'If Target.Address = Target.EntireRow.Address Then
        '    Exit Function
        'End If
        
        For Each Cell In Intersect(Range(RelatedRange), Target)
            Select Case Cell.Column
            Case ISTDHeaderColNumber
                'ISTD has been modified or just added
                Cells(Cell.Row, ISTDHeaderColNumber).Interior.Color = xlNone
                Cells(Cell.Row, ISTD_Conc_ng_ColNumber).Interior.Color = xlNone
                Cells(Cell.Row, ISTD_MW_ColNumber).Interior.Color = xlNone
                If Cell.Value = "" Then
                    Cells(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = xlNone
                    Cells(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = xlNone
                Else
                    'Warn users they need to be changed or filled in
                    Cells(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = RGB(255, 200, 200)
                    Cells(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = RGB(255, 200, 200)
                End If
            Case ISTD_Conc_ng_ColNumber, ISTD_MW_ColNumber
                'Remove the green format if it exists
                Cells(Cell.Row, ISTD_Conc_ng_ColNumber).Interior.Color = xlNone
                Cells(Cell.Row, ISTD_MW_ColNumber).Interior.Color = xlNone
                'If it has been modified under the presence of an ISTD
                If Not Cells(Cell.Row, ISTDHeaderColNumber) = "" Then
                    'Warn users they need to be changed or filled in
                    Cells(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = RGB(255, 200, 200)
                    Cells(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = RGB(255, 200, 200)
                End If
            Case ISTD_Conc_nM_ColNumber
                'Warns user that they must fill up the cell as there is an ISTD
                If Cell.Value = "" And Not Cells(Cell.Row, ISTDHeaderColNumber) = "" Then
                    Cells(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = RGB(255, 200, 200)
                    Cells(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = RGB(255, 200, 200)
                    'Blank value is valid as there is no ISTD
                ElseIf Cell.Value = "" And Cells(Cell.Row, ISTDHeaderColNumber) = "" Then
                    Cells(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = xlNone
                    Cells(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = xlNone
                    'Warn users values has been modified
                ElseIf Not Cells(Cell.Row, ISTDHeaderColNumber) = "" Then
                    Cells(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = RGB(255, 200, 200)
                    Cells(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = RGB(255, 200, 200)
                End If
            Case ISTD_Custom_Unit_ColNumber
                'Warns user that they must fill up the cell as there is an ISTD
                If Cell.Row > 3 And Cell.Value = "" And Not Cells(Cell.Row, ISTDHeaderColNumber) = "" Then
                    Cells(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = RGB(255, 200, 200)
                    Cells(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = RGB(255, 200, 200)
                    'Blank value is valid as there is no ISTD
                ElseIf Cell.Row > 3 And Cell.Value = "" And Cells(Cell.Row, ISTDHeaderColNumber) = "" Then
                    Cells(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = xlNone
                    Cells(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = xlNone
                    'Warn users values has been modified
                ElseIf Cell.Row > 3 And Not Cells(Cell.Row, ISTDHeaderColNumber) = "" Then
                    Cells(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = RGB(255, 200, 200)
                    Cells(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = RGB(255, 200, 200)
                ElseIf Cell.Row = 3 Then
                    Application.EnableEvents = False
                    Dim ISTD_Custom_Unit() As String
                    ISTD_Custom_Unit = ISTD_Annot.Convert_Conc_nM_Array(Cell.Value)
                    Call Utilities.Load_To_Excel(ISTD_Custom_Unit, "Custom_Unit", HeaderRowNumber:=2, _
                                                 DataStartRowNumber:=4, MessageBoxRequired:=False)
                    'Update the Concentration Unit in Sample_Annot sheet if there are entries.
                    Call Autofill_Concentration_Unit_Click
                    Sheets("ISTD_Annot").Activate
                    Application.EnableEvents = True
                End If
            End Select
        Next Cell
    End If

End Function

'Sheet Transition_Name_Annot Function
Public Function ChangeToBlankWhenChanged(ByVal Target As Range)
    'Application.ScreenUpdating = False
    'EventState = Application.EnableEvents
    Application.EnableEvents = False
    'CalcState = Application.Calculation
    'Application.Calculation = xlCalculationManual
    'PageBreakState = ActiveSheet.DisplayPageBreaks
    'ActiveSheet.DisplayPageBreaks = False
    
    Sheets("Transition_Name_Annot").Activate
    

    'Get column position of the headers
    Dim Transition_Name_ColNumber As Integer
    Dim Transition_Name_ISTD_ColNumber As Integer
    Transition_Name_ColNumber = Utilities.Get_Header_Col_Position("Transition_Name", 1)
    Transition_Name_ISTD_ColNumber = Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 1)
    
    Dim Transition_Name_ColLetter As String
    Dim Transition_Name_ISTD_ColLetter As String
    Transition_Name_ColLetter = Utilities.ConvertToLetter(Transition_Name_ColNumber)
    Transition_Name_ISTD_ColLetter = Utilities.ConvertToLetter(Transition_Name_ISTD_ColNumber)

    Dim RelatedRange As String
    RelatedRange = Transition_Name_ColLetter & ":" & Transition_Name_ISTD_ColLetter
    
    'When there is a change
    If Not Intersect(Range(RelatedRange), Target) Is Nothing Then
        'Debug.Print Intersect(Range(RelatedRange), Target).Address
        
        'If rows are deleted, leave the function
        'If Target.Address = Target.EntireRow.Address Then
        '    Exit Function
        'End If
        
        
        For Each Cell In Intersect(Range(RelatedRange), Target)
            Select Case Cell.Column
                'If changes are made in the Transition_Name column
            Case Transition_Name_ColNumber
                Dim TotalRows As Long
                TotalRows = Cells(Rows.Count, ConvertToLetter(Transition_Name_ColNumber)).End(xlUp).Row
                If TotalRows > 1 Then
                    Range(Transition_Name_ColLetter & Cell.Row & ":" & Transition_Name_ISTD_ColLetter & Cell.Row).Interior.Color = xlNone
                    'Whole Transition_Name_ISTD column must be white
                    Range(Transition_Name_ISTD_ColLetter & "2:" & Transition_Name_ISTD_ColLetter & TotalRows).Interior.Color = xlNone
                Else
                    'Whole Transition_Name_ISTD column must be white
                    Range(Transition_Name_ColLetter & Cell.Row & ":" & Transition_Name_ISTD_ColLetter & Cell.Row).Interior.Color = xlNone
                End If
            Case Transition_Name_ISTD_ColNumber
                Range(Transition_Name_ISTD_ColLetter & Cell.Row).Interior.Color = xlNone
            End Select

        Next Cell
        
    End If

    
    'ActiveSheet.DisplayPageBreaks = PageBreakState
    'Application.Calculation = CalcState
    'Application.EnableEvents = EventState
    Application.ScreenUpdating = True

End Function

