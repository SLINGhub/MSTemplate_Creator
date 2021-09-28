Attribute VB_Name = "ISTD_Annot"
Public Function Convert_Conc_nM_Array(Custom_Unit As String) As String()
    Dim ISTD_Conc() As String
    Dim lenArray As Integer
    Dim FactorValue As Double
    
    Sheets("ISTD_Annot").Activate
    
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
    
    'Get Factor Value based on NewUnit
    FactorValue = 1
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
    End Select
    
    'Perform the convertion from nM to NewUnit
    For i = 0 To lenArray - 1
        'Perform convertion only when there is text
        If Len(Trim(ISTD_Conc(i))) > 0 Then
            ISTD_Conc(i) = CDec(CDbl(ISTD_Conc(i))) * FactorValue
        End If
    Next
    
    Convert_Conc_nM_Array = ISTD_Conc
    
End Function

Public Function Get_ISTD_Conc_nM_Array() As String()
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
    Dim lenArray(0 To 2) As Integer
    lenArray(0) = Utilities.StringArrayLen(ISTD_Conc_ngmL)
    lenArray(1) = Utilities.StringArrayLen(ISTD_MW)
    lenArray(2) = Utilities.StringArrayLen(ISTD_Conc_nM)
    
    'Get the length of the longest array
    max_length = 0
    For i = 0 To UBound(lenArray) - LBound(lenArray)
        If lenArray(i) > max_length Then
            max_length = lenArray(i)
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
    For i = 0 To max_length - 1
        If Len(ISTD_Conc_ngmL(i)) <> 0 And Len(ISTD_MW(i)) <> 0 Then
            
            'Ensure that concentration are not negative
            If ISTD_Conc_ngmL(i) <= 0 Then
                MsgBox ("There are non-positive values in the ISTD_Conc_ngmL column")
                Application.EnableEvents = True
                End
            End If
            
            'Ensure that molecular weight is not zero or less than zero
            If ISTD_MW(i) <= 0 Then
                MsgBox ("There are non-positive values in the ISTD_MW column")
                Application.EnableEvents = True
                End
            End If
            
            'Perform the claculation since the two values are valid
            ISTD_Conc_nM(i) = CDec(CDbl(ISTD_Conc_ngmL(i)) / CDbl(ISTD_MW(i))) * 1000
            
            'If the value is a valid ISTD, change the colour to green
            If Not Range(Transition_Name_ISTD_ColLetter & CStr(i + 4)).Value = "" Then
                Range(ISTD_Conc_ng_ColLetter & CStr(i + 4)).Interior.Color = RGB(204, 255, 204)
                Range(ISTD_MW_ColLetter & CStr(i + 4)).Interior.Color = RGB(204, 255, 204)
                Range(ISTD_Conc_nM_ColLetter & CStr(i + 4)).Interior.Color = RGB(204, 255, 204)
                Range(ISTD_Custom_Unit_ColLetter & CStr(i + 4)).Interior.Color = RGB(204, 255, 204)
            End If
            
        End If
        
        'If the ISTD_Conc_[nM] column entry is not empty
        If Not Range(ISTD_Conc_nM_ColLetter & CStr(i + 4)).Value = "" Then
            If Range(ISTD_Conc_nM_ColLetter & CStr(i + 4)).Value <= 0 Then
                MsgBox ("There non-positive values in the ISTD_Conc_nM column")
                Application.EnableEvents = True
                End
            End If
            'If the value is a valid ISTD and ISTD_Conc_nM is filled correctly
            If Not Range(Transition_Name_ISTD_ColLetter & CStr(i + 4)).Value = "" Then
                Range(ISTD_Conc_nM_ColLetter & CStr(i + 4)).Interior.Color = RGB(204, 255, 204)
                Range(ISTD_Custom_Unit_ColLetter & CStr(i + 4)).Interior.Color = RGB(204, 255, 204)
            End If
        End If

    Next
    
    Get_ISTD_Conc_nM_Array = ISTD_Conc_nM
    
End Function


