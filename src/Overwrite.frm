VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Overwrite 
   Caption         =   "Overwrite"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   5730
   OleObjectBlob   =   "Overwrite.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Overwrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Functions that control what happens when buttons in the Overwrite Box are clicked."
Option Explicit
'@ModuleDescription("Functions that control what happens when buttons in the Overwrite Box are clicked.")
'@Folder("Overwrite Functions")

'Public whatsclicked As String
Private master_whatsclicked As String

Public Property Get whatsclicked() As String
    whatsclicked = master_whatsclicked
End Property

Public Property Let whatsclicked(ByVal let_whatsclicked As String)
    master_whatsclicked = let_whatsclicked
End Property

'@Description("Function that controls what happens when the Overwrite Box Cancel button is left clicked.")

'' Function: Cancel_Button_Click
'' --- Code
''  Private Sub Cancel_Button_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the
'' when the following button is
'' left clicked.
''
'' (see Overwrite_Box_Cancel_Button.png)
''
'' Overwrite Box will be hidden.
''
Private Sub Cancel_Button_Click()
Attribute Cancel_Button_Click.VB_Description = "Function that controls what happens when the Overwrite Box Cancel button is left clicked."
    whatsclicked = "Cancel"
    Overwrite.Hide
End Sub

'@Description("Function that controls what happens when the Overwrite Box Overwrite button is left clicked.")

'' Function: Overwrite_Button_Click
'' --- Code
''  Private Sub Overwrite_Button_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the
'' when the following button is
'' left clicked.
''
'' (see Overwrite_Box_Overwrite_Button.png)
''
'' Overwrite Box will be hidden.
''
Private Sub Overwrite_Button_Click()
Attribute Overwrite_Button_Click.VB_Description = "Function that controls what happens when the Overwrite Box Overwrite button is left clicked."
    whatsclicked = "Overwrite"
    Overwrite.Hide
End Sub
