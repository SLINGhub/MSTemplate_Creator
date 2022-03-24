VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransitionNameAnnotSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@IgnoreModule ProcedureNotUsed
'@Folder("Sheets")

'' Function: Worksheet_Change
''
'' Description:
''
'' Function that runs when a cell of a group of cells changed.
'' Currently, it calls the function ColourTracker.Transition_Name_Annot_Checker Target
''
'' See documentation of this function for more information.
''
'' Parameters:
''
''    Target - The changed range in the sheet. Can be more than one cell.
''
Private Sub Worksheet_Change(ByVal Target As Range)
    ColourTracker.Transition_Name_Annot_Checker Target
End Sub

'' Function: ClearTransitionSheet_Click
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Transition_Annot_Clear_Columns_Button.png)
''
'' Currently, it calls the function Transition_Name_Annot_Buttons.Clear_Transition_Name_Annot_Click
''
'' See documentation of this function for more information.
''
Private Sub ClearTransitionSheet_Click()
    Transition_Name_Annot_Buttons.Clear_Transition_Name_Annot_Click
End Sub

'' Function: LoadSampleAnnotationRaw_Click
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' Currently, it calls the function Sample_Annot_Buttons.Load_Sample_Annot_Raw_Column_Name_Click
''
'' (see Transition_Annot_Load_Sample_Annot_Raw.png)
''
'' See documentation of this function for more information.
''
Private Sub LoadSampleAnnotationRaw_Click()
    'In sheet 3 functions
    Sample_Annot_Buttons.Load_Sample_Annot_Raw_Column_Name_Click
End Sub

'' Function: LoadSampleAnnotationTidy_Click
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Transition_Annot_Load_Sample_Annot_Tidy.png)
''
'' Currently, it calls the function Sample_Annot_Buttons.Load_Sample_Annot_Tidy_Column_Name_Click
''
'' See documentation of this function for more information.
''
Private Sub LoadSampleAnnotationTidy_Click()
    'In sheet 3 functions
    Sample_Annot_Buttons.Load_Sample_Annot_Tidy_Column_Name_Click
End Sub

'' Function: LoadISTD_Click
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Transition_Annot_Load_ISTD_To_ISTD_Table_Button.png)
''
'' Currently, it calls the function Transition_Name_Annot_Buttons.Load_Transition_Name_ISTD_Click
''
'' See documentation of this function for more information.
''
Private Sub LoadISTD_Click()
    Transition_Name_Annot_Buttons.Load_Transition_Name_ISTD_Click
End Sub

'' Function: ValidateISTD_Click
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Transition_Annot_Validate_ISTD_Button.png)
''
'' Currently, it calls the function Transition_Name_Annot_Buttons.Validate_ISTD_Click
''
'' See documentation of this function for more information.
''
Private Sub ValidateISTD_Click()
    Transition_Name_Annot_Buttons.Validate_ISTD_Click
End Sub

'' Function: LoadTransition_Click
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Transition_Annot_Load_Transition_Name_Raw_Button.png)
''
'' Currently, it calls the function Transition_Name_Annot_Buttons.Get_Transition_Array_Click
''
'' See documentation of this function for more information.
''
Private Sub LoadTransition_Click()
    Transition_Name_Annot_Buttons.Get_Transition_Array_Click
End Sub

'' Function: LoadTransitionTidy_Click
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Transition_Annot_Load_Transition_Name_Tidy_Button.png)
''
'' Currently, it calls the function Transition_Name_Annot_Buttons.Get_Transition_Array_Tidy_Click
''
'' See documentation of this function for more information.
''
Private Sub LoadTransitionTidy_Click()
    Transition_Name_Annot_Buttons.Get_Transition_Array_Tidy_Click
End Sub

