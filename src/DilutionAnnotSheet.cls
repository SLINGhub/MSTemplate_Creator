VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DilutionAnnotSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@IgnoreModule ProcedureNotUsed
'@Folder("Dilution Annot Functions")

'' Function: ClearDilutionAnnot_Click
'' --- Code
''  Private Sub ClearDilutionAnnot_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Dilution_Annot_Clear_Columns_Button.png)
''
'' Currently, it calls the function Dilution_Annot_Buttons.Clear_Dilution_Annot_Click
''
'' See documentation of this function for more information.
''
Private Sub ClearDilutionAnnot_Click()
    Dilution_Annot_Buttons.Clear_Dilution_Annot_Click
End Sub
