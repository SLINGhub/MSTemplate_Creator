Attribute VB_Name = "ISTD_Annot"
Option Explicit
'@Folder("ISTD Annot Functions")
'@IgnoreModule IntegerDataType

'' Function: Convert_Conc_nM_Array
'' --- Code
''  Public Function Convert_Conc_nM_Array(ByVal Custom_Unit As String) As String()
'' ---
''
'' Description:
''
'' Get the concentration values from the ISTD_Conc_[nM] column
'' and convert them based the units provided in the input Custom_Unit
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
'' --- Code
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
    Dim ISTD_Conc() As String
    Dim lenArrayIndex As Integer
    Dim lenArray As Integer
    Dim FactorValue As Double
    
    ' Get the ISTD_Annot worksheet from the active workbook
    ' The ISTDAnnot is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim ISTD_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "ISTDAnnot") = False Then
        MsgBox ("Sheet ISTD_Annot is missing")
        Application.EnableEvents = True
        Exit Function
    End If
    
    Set ISTD_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "ISTDAnnot")
      
    ISTD_Annot_Worksheet.Activate
      
    ISTD_Conc = Utilities.Load_Columns_From_Excel("ISTD_Conc_[nM]", HeaderRowNumber:=3, _
                                                  DataStartRowNumber:=4, MessageBoxRequired:=False, _
                                                  RemoveBlksAndReplicates:=False, _
                                                  IgnoreHiddenRows:=False, IgnoreEmptyArray:=True)
    
    lenArray = Utilities.StringArrayLen(ISTD_Conc)
    
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


'' Function: Get_ISTD_Conc_nM_Array
'' --- Code
''  Public Function Get_ISTD_Conc_nM_Array(ColourCellRequired As Boolean) As String()
'' ---
''
'' Description:
''
'' Get the concentration values for the ISTD_Conc_[nM] column using
'' the values from the ISTD_Conc_[ng/mL] and ISTD_[MW] column or from
'' manual input from the users on the ISTD_Conc_[nM] column itself.
''
'' If all ISTD_Conc_[ng/mL], ISTD_[MW] and ISTD_Conc_[nM] column are filled
'' the calculated concentration values from ISTD_Conc_[ng/mL] and ISTD_[MW]
'' will take priority to avoid confusion.
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
Public Function Get_ISTD_Conc_nM_Array(ByVal ColourCellRequired As Boolean) As String()

    ' Get the ISTD_Annot worksheet from the active workbook
    ' The ISTDAnnot is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim ISTD_Annot_Worksheet As Worksheet
    
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "ISTDAnnot") = False Then
        MsgBox ("Sheet ISTD_Annot is missing")
        Application.EnableEvents = True
        Exit Function
    End If
    
    Set ISTD_Annot_Worksheet = Utilities.Get_Sheet_By_Code_Name(ActiveWorkbook, "ISTDAnnot")
      
    ISTD_Annot_Worksheet.Activate

    'Declare the column letter position
    Dim Transition_Name_ISTD_ColLetter As String
    Dim ISTD_Conc_ng_ColLetter As String
    Dim ISTD_MW_ColLetter As String
    Dim ISTD_Conc_nM_ColLetter As String
    Dim ISTD_Custom_Unit_ColLetter As String
    Transition_Name_ISTD_ColLetter = Utilities.ConvertToLetter(Utilities.Get_Header_Col_Position("Transition_Name_ISTD", 2))
    ISTD_Conc_ng_ColLetter = Utilities.ConvertToLetter(Utilities.Get_Header_Col_Position("ISTD_Conc_[ng/mL]", 3))
    ISTD_MW_ColLetter = Utilities.ConvertToLetter(Utilities.Get_Header_Col_Position("ISTD_[MW]", 3))
    ISTD_Conc_nM_ColLetter = Utilities.ConvertToLetter(Utilities.Get_Header_Col_Position("ISTD_Conc_[nM]", 3))
    ISTD_Custom_Unit_ColLetter = Utilities.ConvertToLetter(Utilities.Get_Header_Col_Position("Custom_Unit", 2))
    
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
    lenArray(0) = Utilities.StringArrayLen(ISTD_Conc_ngmL)
    lenArray(1) = Utilities.StringArrayLen(ISTD_MW)
    lenArray(2) = Utilities.StringArrayLen(ISTD_Conc_nM)
    
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


