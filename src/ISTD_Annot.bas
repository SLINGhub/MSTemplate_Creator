Attribute VB_Name = "ISTD_Annot"
Attribute VB_Description = "Functions that are commonly called in the ISTD_Annot worksheet."
Option Explicit
'@ModuleDescription("Functions that are commonly called in the ISTD_Annot worksheet.")
'@Folder("ISTD Annot Functions")
'@IgnoreModule IntegerDataType

'@Description("Get the concentration values from the ISTD_Conc_[nM] column, convert them based the units provided in the input Custom_Unit and output as a string array.")

'' Function: Convert_Conc_nM_Array
'' --- Code
''  Public Function Convert_Conc_nM_Array(ByVal Custom_Unit As String) As String()
'' ---
''
'' Description:
''
'' Get the concentration values from the ISTD_Conc_[nM] column,
'' convert them based the units provided in the input Custom_Unit
'' and output as a string array.
''
'' (see ISTD_Annot_ISTD_Conc_nM_Column.png)
''
'' Currently, the accepted input for Custom_Unit are
''
''  - [M] or [umol/uL]
''  - [mM] or [nmol/uL]
''  - [uM] or [pmol/uL]
''  - [nM] or [fmol/uL]
''  - [pM] or [amol/uL]
''
'' Any other input will still return a string array but no conversion
'' of value.
''
'' Parameters:
''
''    Custom_Unit As String - Concentration units to be converted into.
''                            Accepted input are provided in the dropdown
''                            button provided by the Custom_Unit column.
''
'' Returns:
''    A string array of concentration converted to the provided input units.
''
'' Examples:
''
'' --- Code
''    ' Get the ISTD_Annot worksheet from the active workbook
''    ' The ISTDAnnotSheet is a code name
''    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
''    Dim ISTD_Annot_Worksheet As Worksheet
''
''    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "ISTDAnnotSheet") = False Then
''        MsgBox ("Sheet ISTD_Annot is missing")
''        Application.EnableEvents = True
''        Exit Sub
''    End If
''
''    Set ISTD_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "ISTDAnnotSheet")
''
''    Dim ISTD_Conc_nM(2) As String
''    Dim ISTD_Custom_Unit() As String
''    Dim Custom_Unit As String
''
''    ISTD_Conc_nM(0) = "1000"
''    ISTD_Conc_nM(1) = "2000"
''    ISTD_Conc_nM(2) = "3000"
''
''    Custom_Unit = "[uM] or [pmol/uL]"
''
''    Utilities.Load_To_Excel Data_Array:=ISTD_Conc_nM, _
''                            HeaderName:="ISTD_Conc_[nM]", _
''                            HeaderRowNumber:=3, _
''                            DataStartRowNumber:=4, _
''                            MessageBoxRequired:=False
''
''    ISTD_Custom_Unit = ISTD_Annot.Convert_Conc_nM_Array(Custom_Unit:=Custom_Unit)
'' ---
Public Function Convert_Conc_nM_Array(ByVal Custom_Unit As String) As String()
Attribute Convert_Conc_nM_Array.VB_Description = "Get the concentration values from the ISTD_Conc_[nM] column, convert them based the units provided in the input Custom_Unit and output as a string array."
    Dim ISTD_Conc() As String
    Dim lenArrayIndex As Integer
    Dim lenArray As Integer
    Dim FactorValue As Double
    
    ' Get the ISTD_Annot worksheet from the active workbook
    ' The ISTDAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim ISTD_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "ISTDAnnotSheet") = False Then
        MsgBox ("Sheet ISTD_Annot is missing")
        Application.EnableEvents = True
        Exit Function
    End If
    
    Set ISTD_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "ISTDAnnotSheet")
      
    ISTD_Annot_Worksheet.Activate
      
    ISTD_Conc = Utilities.Load_Columns_From_Excel("ISTD_Conc_[nM]", HeaderRowNumber:=3, _
                                                  DataStartRowNumber:=4, MessageBoxRequired:=False, _
                                                  RemoveBlksAndReplicates:=False, _
                                                  IgnoreHiddenRows:=False, IgnoreEmptyArray:=True)
    
    lenArray = Utilities.Get_String_Array_Len(ISTD_Conc)
    
    'Leave the program if lenArray is 0
    If lenArray = 0 Then
        Application.EnableEvents = True
        End
    End If
    
    'Get Factor Value based on Custom Unit
    Select Case Custom_Unit
    Case "[M] or [umol/uL]"
        FactorValue = 10 ^ (-9)
    Case "[mM] or [nmol/uL]"
        FactorValue = 10 ^ (-6)
    Case "[uM] or [pmol/uL]"
        FactorValue = 10 ^ (-3)
    Case "[nM] or [fmol/uL]"
        FactorValue = 1
    Case "[pM] or [amol/uL]"
        FactorValue = 10 ^ 3
    Case Else
        FactorValue = 1
    End Select
    
    'Perform the convertion from nM to NewUnit
    For lenArrayIndex = 0 To lenArray - 1
        'Perform convertion only when there is text
        If Len(Trim$(ISTD_Conc(lenArrayIndex))) > 0 Then
            ISTD_Conc(lenArrayIndex) = CDec(CDbl(ISTD_Conc(lenArrayIndex))) * FactorValue
        End If
    Next
    
    Convert_Conc_nM_Array = ISTD_Conc
    
End Function

'@Description("Get the concentration values for the ISTD_Conc_[nM] column using the values from the ISTD_Conc_[ng/mL] and ISTD_[MW] column or from manual input from the users on the ISTD_Conc_[nM] column itself.")

'' Function: Get_ISTD_Conc_nM_Array
'' --- Code
''  Public Function Get_ISTD_Conc_nM_Array(ByVal ColourCellRequired As Boolean) As String()
'' ---
''
'' Description:
''
'' Get the concentration values for the ISTD_Conc_[nM] column using
'' the values from the ISTD_Conc_[ng/mL] and ISTD_[MW] column or from
'' manual input from the users on the ISTD_Conc_[nM] column itself.
''
'' (see ISTD_Annot_ISTD_Conc_nM_Calculation_1.png)
''
'' Output string array will be {"500", "100"}
''
'' Calculation of ISTD_Conc_[nM] from ISTD_Conc_[ng/mL] and ISTD_[MW] is
''
'' ISTD_Conc_[nM] = ISTD_Conc_[ng/mL] / ISTD_[MW] * 1000
''
'' If all ISTD_Conc_[ng/mL], ISTD_[MW] and ISTD_Conc_[nM] column are filled
'' the calculated concentration values from ISTD_Conc_[ng/mL] and ISTD_[MW]
'' will take priority to avoid confusion.
''
'' (see ISTD_Annot_ISTD_Conc_nM_Calculation_2.png)
''
'' Output string array will be {"500", "100"}
''
'' Parameters:
''
''    ColourCellRequired As Boolean - If set to True, the program will also colour the
''                                    cells in the ISTD_Conc_[nM] column and Custom_Unit
''                                    column to green when the row has either a valid
''                                    Transition_Name_ISTD, ISTD_Conc_[ng/mL] and ISTD_[MW]
''                                    or a valid Transition_Name_ISTD and ISTD_Conc_[nM]
''
'' Returns:
''    A string array of concentration from the ISTD_Conc_[nM] column.
''
'' Examples:
''
'' --- Code
''    'We don't want excel to monitor the sheet when runnning this code
''    Application.EnableEvents = False
''
''    ' Get the ISTD_Annot worksheet from the active workbook
''    ' The ISTDAnnotSheet is a code name
''    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
''    Dim ISTD_Annot_Worksheet As Worksheet
''
''    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "ISTDAnnotSheet") = False Then
''        MsgBox ("Sheet ISTD_Annot is missing")
''        Application.EnableEvents = True
''        Exit Sub
''    End If
''
''    Set ISTD_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "ISTDAnnotSheet")
''
''    ISTD_Annot_Worksheet.Activate
''
''    Dim Transition_Name_ISTD(1) As String
''    Dim ISTD_Conc_ng_per_mL(0) As String
''    Dim ISTD_Conc_MW(0) As String
''    Dim ISTD_Conc_nM(1) As String
''    Dim ISTD_Conc_nM_Results() As String
''    Dim ISTD_Custom_Unit_ColNumber As Integer
''    Dim Custom_Unit As String
''    Dim ISTD_Custom_Unit() As String
''
''    Transition_Name_ISTD(0) = "ISTD1"
''    Transition_Name_ISTD(1) = "ISTD2"
''    ISTD_Conc_ng_per_mL(0) = "1"
''    ISTD_Conc_MW(0) = "2"
''    ISTD_Conc_nM(0) = vbNullString
''    ISTD_Conc_nM(1) = "100"
''    ISTD_Custom_Unit_ColNumber = Utilities.Get_Header_Col_Position("Custom_Unit", 2)
''    Custom_Unit = ISTD_Annot_Worksheet.Cells.Item(3, ISTD_Custom_Unit_ColNumber)
''
''    Utilities.Load_To_Excel Data_Array:=Transition_Name_ISTD, _
''                            HeaderName:="Transition_Name_ISTD", _
''                            HeaderRowNumber:=2, _
''                            DataStartRowNumber:=4, _
''                            MessageBoxRequired:=False
''
''    Utilities.Load_To_Excel Data_Array:=ISTD_Conc_ng_per_mL, _
''                            HeaderName:="ISTD_Conc_[ng/mL]", _
''                            HeaderRowNumber:=3, _
''                            DataStartRowNumber:=4, _
''                            MessageBoxRequired:=False
''
''    Utilities.Load_To_Excel Data_Array:=ISTD_Conc_MW, _
''                            HeaderName:="ISTD_[MW]", _
''                            HeaderRowNumber:=3, _
''                            DataStartRowNumber:=4, _
''                            MessageBoxRequired:=False
''
''    Utilities.Load_To_Excel Data_Array:=ISTD_Conc_nM, _
''                            HeaderName:="ISTD_Conc_[nM]", _
''                            HeaderRowNumber:=3, _
''                            DataStartRowNumber:=4, _
''                            MessageBoxRequired:=False
''
''    ISTD_Conc_nM_Results = ISTD_Annot.Get_ISTD_Conc_nM_Array(ColourCellRequired:=True)
''
''    Utilities.Load_To_Excel Data_Array:=ISTD_Conc_nM_Results, _
''                            HeaderName:="ISTD_Conc_[nM]", _
''                            HeaderRowNumber:=3, _
''                            DataStartRowNumber:=4, _
''                            MessageBoxRequired:=False
''
''    ISTD_Custom_Unit = ISTD_Annot.Convert_Conc_nM_Array(Custom_Unit)
''
''    Utilities.Load_To_Excel Data_Array:=ISTD_Custom_Unit, _
''                            HeaderName:="Custom_Unit", _
''                            HeaderRowNumber:=2, _
''                            DataStartRowNumber:=4, _
''                            MessageBoxRequired:=False
''
''    Application.EnableEvents = True
'' ---
Public Function Get_ISTD_Conc_nM_Array(ByVal ColourCellRequired As Boolean) As String()
Attribute Get_ISTD_Conc_nM_Array.VB_Description = "Get the concentration values for the ISTD_Conc_[nM] column using the values from the ISTD_Conc_[ng/mL] and ISTD_[MW] column or from manual input from the users on the ISTD_Conc_[nM] column itself."

    ' Get the ISTD_Annot worksheet from the active workbook
    ' The ISTDAnnotSheet is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim ISTD_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "ISTDAnnotSheet") = False Then
        MsgBox ("Sheet ISTD_Annot is missing")
        Application.EnableEvents = True
        Exit Function
    End If
    
    Set ISTD_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "ISTDAnnotSheet")
      
    ISTD_Annot_Worksheet.Activate

    'Declare the column letter position
    Dim Transition_Name_ISTD_ColLetter As String
    Dim ISTD_Conc_ng_ColLetter As String
    Dim ISTD_MW_ColLetter As String
    Dim ISTD_Conc_nM_ColLetter As String
    Dim ISTD_Custom_Unit_ColLetter As String
    Transition_Name_ISTD_ColLetter = Utilities.Convert_To_Letter(Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 2))
    ISTD_Conc_ng_ColLetter = Utilities.Convert_To_Letter(Utilities.Get_Header_Col_Position("ISTD_Conc_[ng/mL]", 3))
    ISTD_MW_ColLetter = Utilities.Convert_To_Letter(Utilities.Get_Header_Col_Position("ISTD_[MW]", 3))
    ISTD_Conc_nM_ColLetter = Utilities.Convert_To_Letter(Utilities.Get_Header_Col_Position("ISTD_Conc_[nM]", 3))
    ISTD_Custom_Unit_ColLetter = Utilities.Convert_To_Letter(Utilities.Get_Header_Col_Position("Custom_Unit", 2))
    
    'Declare three dynamic arrays
    Dim ISTD_Conc_ngmL() As String
    Dim ISTD_MW() As String
    Dim ISTD_Conc_nM() As String
    
    'Get the columns as array string
    ISTD_Conc_ngmL = Utilities.Load_Columns_From_Excel("ISTD_Conc_[ng/mL]", HeaderRowNumber:=3, _
                                                       DataStartRowNumber:=4, MessageBoxRequired:=False, _
                                                       RemoveBlksAndReplicates:=False, _
                                                       IgnoreHiddenRows:=False, IgnoreEmptyArray:=True)
    ISTD_MW = Utilities.Load_Columns_From_Excel("ISTD_[MW]", HeaderRowNumber:=3, _
                                                DataStartRowNumber:=4, MessageBoxRequired:=False, _
                                                RemoveBlksAndReplicates:=False, _
                                                IgnoreHiddenRows:=False, IgnoreEmptyArray:=True)
    ISTD_Conc_nM = Utilities.Load_Columns_From_Excel("ISTD_Conc_[nM]", HeaderRowNumber:=3, _
                                                     DataStartRowNumber:=4, MessageBoxRequired:=False, _
                                                     RemoveBlksAndReplicates:=False, _
                                                     IgnoreHiddenRows:=False, IgnoreEmptyArray:=True)
    
    'Declare an array of the three array length
    Dim lenArrayIndex As Integer
    Dim lenArray(0 To 2) As Integer
    lenArray(0) = Utilities.Get_String_Array_Len(ISTD_Conc_ngmL)
    lenArray(1) = Utilities.Get_String_Array_Len(ISTD_MW)
    lenArray(2) = Utilities.Get_String_Array_Len(ISTD_Conc_nM)
    
    'Get the length of the longest array
    Dim max_length As Integer
    max_length = 0
    For lenArrayIndex = 0 To UBound(lenArray) - LBound(lenArray)
        If lenArray(lenArrayIndex) > max_length Then
            max_length = lenArray(lenArrayIndex)
        End If
    Next
    
    'Leave the program if max_length is 0
    If max_length = 0 Then
        'Application.EnableEvents = True
        End
    End If
    
    'Resize the three array to the largest length but keep the values
    ReDim Preserve ISTD_Conc_ngmL(max_length)
    ReDim Preserve ISTD_MW(max_length)
    ReDim Preserve ISTD_Conc_nM(max_length)
    
    'Perform the calculation of ISTD_Conc_ngmL divide by ISTD_MW when necessary
    For lenArrayIndex = 0 To max_length - 1
        If Len(ISTD_Conc_ngmL(lenArrayIndex)) <> 0 And Len(ISTD_MW(lenArrayIndex)) <> 0 Then
            
            'Ensure that concentration are not negative
            If ISTD_Conc_ngmL(lenArrayIndex) <= 0 Then
                MsgBox ("There are non-positive values in the ISTD_Conc_ngmL column")
                Application.EnableEvents = True
                End
            End If
            
            'Ensure that molecular weight is not zero or less than zero
            If ISTD_MW(lenArrayIndex) <= 0 Then
                MsgBox ("There are non-positive values in the ISTD_MW column")
                Application.EnableEvents = True
                End
            End If
            
            'Perform the calculation since the two values are valid
            ISTD_Conc_nM(lenArrayIndex) = CDec(CDbl(ISTD_Conc_ngmL(lenArrayIndex)) / CDbl(ISTD_MW(lenArrayIndex))) * 1000
            
            If ColourCellRequired Then
                'If the value is a valid ISTD, change the colour to green
                If Not ISTD_Annot_Worksheet.Range(Transition_Name_ISTD_ColLetter & CStr(lenArrayIndex + 4)).Value = vbNullString Then
                    ISTD_Annot_Worksheet.Range(ISTD_Conc_ng_ColLetter & CStr(lenArrayIndex + 4)).Interior.Color = RGB(204, 255, 204)
                    ISTD_Annot_Worksheet.Range(ISTD_MW_ColLetter & CStr(lenArrayIndex + 4)).Interior.Color = RGB(204, 255, 204)
                    ISTD_Annot_Worksheet.Range(ISTD_Conc_nM_ColLetter & CStr(lenArrayIndex + 4)).Interior.Color = RGB(204, 255, 204)
                    ISTD_Annot_Worksheet.Range(ISTD_Custom_Unit_ColLetter & CStr(lenArrayIndex + 4)).Interior.Color = RGB(204, 255, 204)
                End If
            End If
            
        End If
        
        'If the ISTD_Conc_[nM] column entry is not empty
        If Not ISTD_Annot_Worksheet.Range(ISTD_Conc_nM_ColLetter & CStr(lenArrayIndex + 4)).Value = vbNullString Then
            If ISTD_Annot_Worksheet.Range(ISTD_Conc_nM_ColLetter & CStr(lenArrayIndex + 4)).Value <= 0 Then
                MsgBox ("There non-positive values in the ISTD_Conc_nM column")
                Application.EnableEvents = True
                End
            End If
            
            If ColourCellRequired Then
                'If the value is a valid ISTD and ISTD_Conc_nM is filled correctly, change the colour to green
                If Not ISTD_Annot_Worksheet.Range(Transition_Name_ISTD_ColLetter & CStr(lenArrayIndex + 4)).Value = vbNullString Then
                    ISTD_Annot_Worksheet.Range(ISTD_Conc_nM_ColLetter & CStr(lenArrayIndex + 4)).Interior.Color = RGB(204, 255, 204)
                    ISTD_Annot_Worksheet.Range(ISTD_Custom_Unit_ColLetter & CStr(lenArrayIndex + 4)).Interior.Color = RGB(204, 255, 204)
                End If
            End If
        End If

    Next
    
    Get_ISTD_Conc_nM_Array = ISTD_Conc_nM
    
End Function
