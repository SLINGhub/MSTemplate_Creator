Attribute VB_Name = "ColourTracker"
'@IgnoreModule IntegerDataType
Option Explicit
'@Folder("Colour Tracker")

'' Group: Colour Tracker
''
'' Function: ISTD_Calculation_Checker
''
'' Description:
''
'' Function that controls what happens when certain cells in
'' the ISTD_Annot sheet is changed.
''
'' Here is the list of expected behaviours
''
'' When users first fill in the column Transition_Name_ISTD,
'' The correspnding rows in the ISTD_Conc_[nM] and Custom Unit
'' column will turn red
''
'' (see ISTD_Annot_First_Fill_In.png)
''
'' Columns and rows will turn green when the button
'' "Convert to nM and Verify" is pressed when either
'' both ISTD_Conc_[nM] and ISTD_[MW] is entered or
'' the ISTD_Conc_[nM] is entered. Custom units will
'' be automatically calculated. Do note that this
'' colouring of cells is done by the function
'' ISTD_Annot.Get_ISTD_Conc_nM_Array and not this
'' function.
''
'' (see ISTD_Annot_Press_Convert_Button.png)
''
'' From there if any entries is modified but the row has an
'' Transition_Name_ISTD entry, both
'' ISTD_Conc_[nM] and ISTD_[MW] cells will turn white
'' but ISTD_Conc_[nM] and Custom Unit will turn red.
''
'' (see ISTD_Annot_Modify_Entries.png)
''
'' If the Transition_Name_ISTD entry is removed, all the
'' columns will turn white
''
'' (see ISTD_Annot_Remove_ISTD_Entries.png)
''
Public Sub ISTD_Calculation_Checker(ByVal Target As Range)

    ' Get the ISTD_Annot worksheet from the active workbook
    ' The ISTDAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim ISTD_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "ISTDAnnotSheet") = False Then
        MsgBox ("Sheet ISTD_Annot is missing")
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Set ISTD_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "ISTDAnnotSheet")
      
    ISTD_Annot_Worksheet.Activate
    
    Application.EnableEvents = True
    
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
    'Dim ISTD_Conc_ng_ColLetter As String
    'Dim ISTD_MW_ColLetter As String
    'Dim ISTD_Conc_nM_ColLetter As String
    Dim ISTD_Custom_Unit_ColLetter As String
    ISTDHeaderColLetter = Utilities.Convert_To_Letter(ISTDHeaderColNumber)
    'ISTD_Conc_ng_ColLetter = Utilities.Convert_To_Letter(ISTD_Conc_ng_ColNumber)
    'ISTD_MW_ColLetter = Utilities.Convert_To_Letter(ISTD_MW_ColNumber)
    'ISTD_Conc_nM_ColLetter = Utilities.Convert_To_Letter(ISTD_Conc_nM_ColNumber)
    ISTD_Custom_Unit_ColLetter = Utilities.Convert_To_Letter(ISTD_Custom_Unit_ColNumber)
    
    Dim RelatedRange As String
    Dim Cell As Range

    RelatedRange = ISTDHeaderColLetter & ":" & ISTD_Custom_Unit_ColLetter
        
    If Not Intersect(ISTD_Annot_Worksheet.Range(RelatedRange), Target) Is Nothing Then
        'Debug.Print Intersect(Range(RelatedRange), Target).Address
        
        'If rows are deleted, leave the function
        'Debug.Print Target.Address
        'Debug.Print Target.EntireRow.Address
        'If Target.Address = Target.EntireRow.Address Then
        '    Exit Function
        'End If
           
        For Each Cell In Intersect(ISTD_Annot_Worksheet.Range(RelatedRange), Target)
            Select Case Cell.Column
            Case ISTDHeaderColNumber
                ' When ISTD has been modified or just added in that row,
                ' remove all colours in the Transition_Name_ISTD, ISTD_Conc_[ng/mL]
                ' and ISTD_[MW] column for that corresponding row
                ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTDHeaderColNumber).Interior.Color = xlNone
                ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Conc_ng_ColNumber).Interior.Color = xlNone
                ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_MW_ColNumber).Interior.Color = xlNone
                ' For the ISTD_Conc_[nM] column and the Custom Unit column,
                ' If the row has an ISTD added/modified, colour them red to warn users they need to be changed or filled in
                ' If the row has an ISTD remove, colur them white
                If Cell.Value = vbNullString Then
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = xlNone
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = xlNone
                Else
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = RGB(255, 200, 200)
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = RGB(255, 200, 200)
                End If
            Case ISTD_Conc_ng_ColNumber, ISTD_MW_ColNumber
                'Remove the green format if it exists
                ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Conc_ng_ColNumber).Interior.Color = xlNone
                ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_MW_ColNumber).Interior.Color = xlNone
                'If it has been modified under the presence of an ISTD
                If Not ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTDHeaderColNumber) = vbNullString Then
                    'Warn users they need to be changed or filled in
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = RGB(255, 200, 200)
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = RGB(255, 200, 200)
                End If
            Case ISTD_Conc_nM_ColNumber
                'Warns user that they must fill up the cell as there is an ISTD
                If Cell.Value = vbNullString And Not ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTDHeaderColNumber) = vbNullString Then
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = RGB(255, 200, 200)
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = RGB(255, 200, 200)
                    'Blank value is valid as there is no ISTD
                ElseIf Cell.Value = vbNullString And ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTDHeaderColNumber) = vbNullString Then
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = xlNone
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = xlNone
                    'Warn users values has been modified
                ElseIf Not ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTDHeaderColNumber) = vbNullString Then
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = RGB(255, 200, 200)
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = RGB(255, 200, 200)
                End If
            Case ISTD_Custom_Unit_ColNumber
                'Warns user that they must fill up the cell as there is an ISTD
                If Cell.Row > 3 And Cell.Value = vbNullString And Not ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTDHeaderColNumber) = vbNullString Then
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = RGB(255, 200, 200)
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = RGB(255, 200, 200)
                    'Blank value is valid as there is no ISTD
                ElseIf Cell.Row > 3 And Cell.Value = vbNullString And ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTDHeaderColNumber) = vbNullString Then
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = xlNone
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = xlNone
                    'Warn users values has been modified
                ElseIf Cell.Row > 3 And Not ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTDHeaderColNumber) = vbNullString Then
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Conc_nM_ColNumber).Interior.Color = RGB(255, 200, 200)
                    ISTD_Annot_Worksheet.Cells.Item(Cell.Row, ISTD_Custom_Unit_ColLetter).Interior.Color = RGB(255, 200, 200)
                ElseIf Cell.Row = 3 Then
                    'Update the Concentration Unit in Sample_Annot sheet if the ISTD_Custom_Unit
                    'is changed
                    Sample_Annot_Buttons.Autofill_Concentration_Unit_Click
                    'Convert the units
                    Application.EnableEvents = False
                    Dim ISTD_Custom_Unit() As String
                    ISTD_Custom_Unit = ISTD_Annot.Convert_Conc_nM_Array(Cell.Value)
                    Utilities.Load_To_Excel Data_Array:=ISTD_Custom_Unit, _
                                            HeaderName:="Custom_Unit", _
                                            HeaderRowNumber:=2, _
                                            DataStartRowNumber:=4, _
                                            MessageBoxRequired:=False
                    ISTD_Annot_Worksheet.Activate
                    Application.EnableEvents = True
                End If
            End Select
        Next Cell
    End If

End Sub


'' Function: Transition_Name_Annot_Checker
''
'' Description:
''
'' Function that controls what happens when certain cells in
'' the Transition_Name_Annot sheet is changed.
''
'' Here is the list of expected behaviours
''
'' When a cell in Transition_Name_ISTD has been modified,
'' the cell will turn white.
''
'' (see Transition_Name_Annot_Modify_ISTD_Entries.png)
''
'' When a cell in Transition_Name has been modified,
'' the cell will turn white. All cells in the Transition_Name_ISTD
'' column will also turn white
''
'' (see Transition_Name_Annot_Modify_Transition_Entries.png)
''
Public Sub Transition_Name_Annot_Checker(ByVal Target As Range)
    'Application.ScreenUpdating = False
    'EventState = Application.EnableEvents
    Application.EnableEvents = False
    'CalcState = Application.Calculation
    'Application.Calculation = xlCalculationManual
    'PageBreakState = ActiveSheet.DisplayPageBreaks
    'ActiveSheet.DisplayPageBreaks = False
    
    ' Get the Transition_Annot worksheet from the active workbook
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
    
    'Get column position of the headers
    Dim Transition_Name_ColNumber As Integer
    Dim Transition_Name_ISTD_ColNumber As Integer
    Transition_Name_ColNumber = Utilities.Get_Header_Col_Position("Transition_Name", 1)
    Transition_Name_ISTD_ColNumber = Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 1)
    
    Dim Transition_Name_ColLetter As String
    Dim Transition_Name_ISTD_ColLetter As String
    Transition_Name_ColLetter = Utilities.Convert_To_Letter(Transition_Name_ColNumber)
    Transition_Name_ISTD_ColLetter = Utilities.Convert_To_Letter(Transition_Name_ISTD_ColNumber)

    Dim RelatedRange As String
    Dim Cell As Range
    RelatedRange = Transition_Name_ColLetter & ":" & Transition_Name_ISTD_ColLetter
    
    'When there is a change
    If Not Intersect(Transition_Name_Annot_Worksheet.Range(RelatedRange), Target) Is Nothing Then
        'Debug.Print Intersect(Range(RelatedRange), Target).Address
        
        'If rows are deleted, leave the function
        'If Target.Address = Target.EntireRow.Address Then
        '    Exit Function
        'End If

        For Each Cell In Intersect(Transition_Name_Annot_Worksheet.Range(RelatedRange), Target)
            Select Case Cell.Column
                'If changes are made in the Transition_Name column
            Case Transition_Name_ColNumber
                Dim TotalRows As Long
                TotalRows = Transition_Name_Annot_Worksheet.Cells.Item(Transition_Name_Annot_Worksheet.Rows.Count, Utilities.Convert_To_Letter(Transition_Name_ColNumber)).End(xlUp).Row
                If TotalRows > 1 Then
                    Transition_Name_Annot_Worksheet.Range(Transition_Name_ColLetter & Cell.Row & ":" & Transition_Name_ISTD_ColLetter & Cell.Row).Interior.Color = xlNone
                    'Whole Transition_Name_ISTD column must be white
                    Transition_Name_Annot_Worksheet.Range(Transition_Name_ISTD_ColLetter & "2:" & Transition_Name_ISTD_ColLetter & TotalRows).Interior.Color = xlNone
                Else
                    'Whole Transition_Name_ISTD column must be white
                    Transition_Name_Annot_Worksheet.Range(Transition_Name_ColLetter & Cell.Row & ":" & Transition_Name_ISTD_ColLetter & Cell.Row).Interior.Color = xlNone
                End If
            Case Transition_Name_ISTD_ColNumber
                Transition_Name_Annot_Worksheet.Range(Transition_Name_ISTD_ColLetter & Cell.Row).Interior.Color = xlNone
            End Select

        Next Cell
        
    End If

    'ActiveSheet.DisplayPageBreaks = PageBreakState
    'Application.Calculation = CalcState
    'Application.EnableEvents = EventState
    'Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub


