VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISTDAnnotSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@IgnoreModule ProcedureNotUsed
'@Folder("Sheets")

'' Function: Worksheet_Change
'' --- Code
''  Private Sub Worksheet_Change
'' ---
''
'' Description:
''
'' Function that runs when a cell of a group of cells changed.
'' Currently, it calls the function ColourTracker.ISTD_Calculation_Checker Target
''
'' See documentation of this function for more information.
''
'' Parameters:
''
''    Target - The changed range in the sheet. Can be more than one cell.
''
Private Sub Worksheet_Change(ByVal Target As Range)
    ColourTracker.ISTD_Calculation_Checker Target
End Sub

'' Function: ClearISTDTable_Click
'' --- Code
''  Private Sub ClearISTDTable_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' Currently, it calls the function ISTD_Annot_Buttons.Clear_ISTD_Table_Click
''
'' See documentation of this function for more information.
''
Private Sub ClearISTDTable_Click()
    ISTD_Annot_Buttons.Clear_ISTD_Table_Click
End Sub

'' Function: ConvertToNanomolar_Click
'' --- Code
''  Private Sub ConvertToNanomolar_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' Currently, it calls the function ISTD_Annot_Buttons.Convert_To_Nanomolar_Click
''
'' See documentation of this function for more information.
''
Private Sub ConvertToNanomolar_Click()
    ISTD_Annot_Buttons.Convert_To_Nanomolar_Click
End Sub

