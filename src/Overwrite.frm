VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Overwrite 
   Caption         =   "Overwrite"
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   2790
   OleObjectBlob   =   "Overwrite.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Overwrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@Folder("Overwrite Functions")

'Public whatsclicked As String
Private master_whatsclicked As String

Public Property Get whatsclicked() As String
    whatsclicked = master_whatsclicked
End Property

Public Property Let whatsclicked(ByVal let_whatsclicked As String)
    master_whatsclicked = let_whatsclicked
End Property

'' Function: Cancel_Button_Click
'' --- Code
''  Private Sub Cancel_Button_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the
'' when the following button is
'' left clicked
''
'' (see Overwrite_Box_Cancel_Button.png)
''
'' Public Property whatsclicked = "Cancel"
'' Overwrite Box will be hidden
''
Private Sub Cancel_Button_Click()
    whatsclicked = "Cancel"
    Overwrite.Hide
End Sub

'' Function: Overwrite_Button_Click
'' --- Code
''  Private Sub Overwrite_Button_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the
'' when the following button is
'' left clicked
''
'' (see Overwrite_Box_Overwrite_Button.png)
''
'' Public Property whatsclicked = "Overwrite"
'' Overwrite Box will be hidden
''
Private Sub Overwrite_Button_Click()
    whatsclicked = "Overwrite"
    Overwrite.Hide
End Sub

