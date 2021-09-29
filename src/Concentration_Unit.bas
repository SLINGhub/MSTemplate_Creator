Attribute VB_Name = "Concentration_Unit"
Public Function Get_Mol_From_Custom_ISTD_Concentration_Unit(Custom_Unit As String) As String
    'Input is of the form "[?M] or [?mol/uL]"
    'Function tries to get ?mol from the above string

    Dim Right_Custom_Unit As String
    Dim RightConcUnitRegEx As New RegExp
    'Get the right custom unit value after "or"
    RightConcUnitRegEx.Pattern = "(.*or)"
    RightConcUnitRegEx.Global = True
    Right_Custom_Unit = Trim(RightConcUnitRegEx.Replace(Custom_Unit, " "))
    'Remove square brackets and mL
    RightConcUnitRegEx.Pattern = "[\[\]]"
    Right_Custom_Unit = Trim(RightConcUnitRegEx.Replace(Right_Custom_Unit, " "))
    RightConcUnitRegEx.Pattern = "/uL"
    'Right_Custom_Unit = RightConcUnitRegEx.Execute(Custom_Unit)(0).SubMatches(0)
    Right_Custom_Unit = Trim(RightConcUnitRegEx.Replace(Right_Custom_Unit, " "))
    
    Get_Mol_From_Custom_ISTD_Concentration_Unit = Right_Custom_Unit

End Function
