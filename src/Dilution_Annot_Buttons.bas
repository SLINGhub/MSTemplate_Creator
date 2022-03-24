Attribute VB_Name = "Dilution_Annot_Buttons"
Option Explicit
'@Folder("Dilution Annot Functions")

'' Function: Clear_Dilution_Annot_Click
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
    'To ensure that Filters does not affect the assignment
    Utilities.RemoveFilterSettings
    
    Clear_Dilution_Annot.Show
End Sub
