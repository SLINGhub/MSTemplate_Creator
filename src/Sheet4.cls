VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Sheet Dilution_Annot Functions
'Taken from https://superuser.com/questions/870926/restrict-paste-into-dropdown-cells-in-excel
'https://www.extendoffice.com/documents/excel/2686-excel-drop-down-list-prevent-paste.html
Private Function HasValidation(r As Range) As Boolean
    On Error Resume Next
'    'On Error GoTo ErrHandler
    DVtype = r.Validation.Type
    If DVtype = 2 Then
        HasValidation = True
    Else
        HasValidation = False
    End If
''ErrHandler:
''    HasValidation = True
End Function

Private Sub ClearDilutionAnnot_Click()
    Call Clear_Dilution_Annot_Click
End Sub
