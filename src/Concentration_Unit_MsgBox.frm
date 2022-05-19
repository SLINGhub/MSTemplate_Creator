VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Concentration_Unit_MsgBox 
   Caption         =   "Concentration Unit"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   165
   ClientWidth     =   4395
   OleObjectBlob   =   "Concentration_Unit_MsgBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Concentration_Unit_MsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Functions that control what happens when buttons in the Concentration Unit MsgBox are clicked."

Option Explicit
'@ModuleDescription("Functions that control what happens when buttons in the Concentration Unit MsgBox are clicked.")
'@Folder("Sample Annot Buttons")

'@Description("Function that controls what happens when the Ok button is left clicked.")

'' Function: Concentration_Unit_Ok_Button_Click
'' --- Code
''  Private Sub Concentration_Unit_Ok_Button_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the Ok button is
'' left clicked on this message box.
''
'' (see Concentration_Unit_Ok_Button.png)
''
'' Message box will be unloaded.
''
Private Sub Concentration_Unit_Ok_Button_Click()
Attribute Concentration_Unit_Ok_Button_Click.VB_Description = "Function that controls what happens when the Ok button is left clicked."
    Unload Concentration_Unit_MsgBox
End Sub
