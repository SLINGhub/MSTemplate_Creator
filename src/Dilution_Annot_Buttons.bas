Attribute VB_Name = "Dilution_Annot_Buttons"
Attribute VB_Description = "Functions that control what happens when buttons in the Dilution_Annot worksheet are clicked."
Option Explicit
'@ModuleDescription("Functions that control what happens when buttons in the Dilution_Annot worksheet are clicked.")

'@Folder("Dilution Annot Functions")
'@Description("Function that controls what happens when the Clear Columns button is left clicked.")

'' Function: Clear_Dilution_Annot_Click
'' --- Code
''  Public Sub Clear_Dilution_Annot_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Dilution_Annot_Clear_Columns_Button.png)
''
'' The following pop up box will appear. Asking the users
'' which column to clear.
''
'' (see Dilution_Annot_Clear_Data_Pop_Up.png)
''
Public Sub Clear_Dilution_Annot_Click()
Attribute Clear_Dilution_Annot_Click.VB_Description = "Function that controls what happens when the Clear Columns button is left clicked."
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    Clear_Dilution_Annot.Show
End Sub
