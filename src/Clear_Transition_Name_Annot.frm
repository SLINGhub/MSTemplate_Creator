VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Clear_Transition_Name_Annot 
   Caption         =   "Clear Transition_Name_Annot"
   ClientHeight    =   2715
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   3810
   OleObjectBlob   =   "Clear_Transition_Name_Annot.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Clear_Transition_Name_Annot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@Folder("Transition_Annot_Buttons")

'' Function: Transition_Name_Annot_Clear_Click
''
'' Description:
''
'' Function that controls what happens when the Clear Data button is
'' left clicked
''
'' (see Transition_Name_Annot_Clear_Data_Button.png)
''
'' All data found in the columns that was checked will be cleared.
''
Private Sub Transition_Name_Annot_Clear_Click()
    If Transition_Name.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Transition_Name", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
    If Transition_Name_ISTD.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Transition_Name_ISTD", _
                                HeaderRowNumber:=1, _
                                DataStartRowNumber:=2
    End If
End Sub

