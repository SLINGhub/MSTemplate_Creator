Attribute VB_Name = "Concentration_Unit"
Option Explicit
'@Folder("Sample Annot Functions")

'' Function: Get_Mol_From_Custom_ISTD_Concentration_Unit
''
'' Description:
''
'' Function used to extract *?mol* from the string input
'' "[?M] or [?mol/uL]". Currently *?mol* can be umol,
'' nmol, pmol, fmol, amol
''
'' Parameters:
''
''    Custom_Unit As String - String input of the form "[?M] or [?mol/uL]"
''                            where *?mol* can be umol, nmol, pmol, fmol, amol.
''
'' Returns:
''    A string in which *?mol* is extracted from "[?M] or [?mol/uL]".
''
'' Examples:
''
'' --- Code
''   Dim Custom_ISTD_Concentration_Unit As String
''   Dim Output_Custom_Unit As String
''
''   Custom_ISTD_Concentration_Unit = "[uM] or [pmol/uL]"
''   Output_Custom_Unit = Concentration_Unit.Get_Mol_From_Custom_ISTD_Concentration_Unit(Custom_ISTD_Concentration_Unit)
''
''   ' Output should be "pmol"
''   Debug.Print Output_Custom_Unit
'' ---
Public Function Get_Mol_From_Custom_ISTD_Concentration_Unit(ByVal Custom_Unit As String) As String
    'Input is of the form "[?M] or [?mol/uL]"
    'Function tries to get ?mol from the above string

    Dim Right_Custom_Unit As String
    Dim RightConcUnitRegEx As RegExp
    Set RightConcUnitRegEx = New RegExp
    
    'Get the right custom unit value after "or"
    RightConcUnitRegEx.Pattern = "(.*or)"
    RightConcUnitRegEx.Global = True
    Right_Custom_Unit = Trim$(RightConcUnitRegEx.Replace(Custom_Unit, " "))
    
    'Remove square brackets and mL
    RightConcUnitRegEx.Pattern = "[\[\]]"
    Right_Custom_Unit = Trim$(RightConcUnitRegEx.Replace(Right_Custom_Unit, " "))
    RightConcUnitRegEx.Pattern = "/uL"
    
    'Right_Custom_Unit = RightConcUnitRegEx.Execute(Custom_Unit)(0).SubMatches(0)
    Right_Custom_Unit = Trim$(RightConcUnitRegEx.Replace(Right_Custom_Unit, " "))
    
    Get_Mol_From_Custom_ISTD_Concentration_Unit = Right_Custom_Unit

End Function
