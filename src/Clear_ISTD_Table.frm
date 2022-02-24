VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Clear_ISTD_Table 
   Caption         =   "Clear ISTD_Table"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   408
   ClientWidth     =   2820
   OleObjectBlob   =   "Clear_ISTD_Table.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Clear_ISTD_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@Folder("ISTD_Annot_Buttons")

'' Function: ISTD_Table_Clear_Click
'' --- Code
''  Private Sub ISTD_Table_Clear_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the Clear Data button is
'' left clicked
''
'' (see ISTD_Annot_Clear_Data_Button.png)
''
'' All data found in the columns that was checked will be cleared.
''
Private Sub ISTD_Table_Clear_Click()
    If Transition_Name_ISTD.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="Transition_Name_ISTD", _
                                HeaderRowNumber:=2, _
                                DataStartRowNumber:=4
    End If
    If ISTD_Conc_ngmL.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="ISTD_Conc_[ng/mL]", _
                                HeaderRowNumber:=3, _
                                DataStartRowNumber:=4
    End If
    If ISTD_MW.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="ISTD_[MW]", _
                                HeaderRowNumber:=3, _
                                DataStartRowNumber:=4
    End If
    If ISTD_Conc_nM.Value = True Then
        Utilities.Clear_Columns HeaderToClear:="ISTD_Conc_[nM]", _
                                HeaderRowNumber:=3, _
                                DataStartRowNumber:=4
        Utilities.Clear_Columns HeaderToClear:="Custom_Unit", _
                                HeaderRowNumber:=2, _
                                DataStartRowNumber:=4
    End If
End Sub

