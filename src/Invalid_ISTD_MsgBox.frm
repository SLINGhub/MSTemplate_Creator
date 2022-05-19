VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Invalid_ISTD_MsgBox 
   Caption         =   "Invalid ISTD"
   ClientHeight    =   3495
   ClientLeft      =   135
   ClientTop       =   495
   ClientWidth     =   4365
   OleObjectBlob   =   "Invalid_ISTD_MsgBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Invalid_ISTD_MsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Functions that control what happens when buttons in the Invalid ISTD MsgBox are clicked."

Option Explicit
'@ModuleDescription("Functions that control what happens when buttons in the Invalid ISTD MsgBox are clicked.")
'@Folder("Transition Annot Buttons")

'@Description("Function that controls what happens when the Ok button is left clicked.")

'' Function: Invalid_ISTD_Ok_Button_Click
'' --- Code
''  Private Sub Invalid_ISTD_Ok_Button_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Transition_Name_Annot_Invalid_ISTD_Message_OK_Button.png)
''
'' Message box will disappear after clicking the button.
'' Users must correct the invalid ISTD input.
''
Private Sub Invalid_ISTD_Ok_Button_Click()
Attribute Invalid_ISTD_Ok_Button_Click.VB_Description = "Function that controls what happens when the Ok button is left clicked."
    Unload Invalid_ISTD_MsgBox
End Sub
