VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Load_Sample_Annot_Tidy 
   Caption         =   "Load_Sample_Annot_Table"
   ClientHeight    =   2820
   ClientLeft      =   90
   ClientTop       =   210
   ClientWidth     =   4845
   OleObjectBlob   =   "Load_Sample_Annot_Tidy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Load_Sample_Annot_Tidy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public whatsclicked As String

Private Sub Create_New_Sample_Annot_Tidy_Button_Click()
    whatsclicked = "Create_New_Sample_Annot_Tidy_Button"
    Load_Sample_Annot_Tidy.Hide
End Sub

' Load the file path of the tidy data
Private Sub Browse_Tidy_Data_Click()
    Dim xFileNames As Variant
    xFileNames = Application.GetOpenFilename(Title:="Load Raw Data File", MultiSelect:=True)
    
    'When no file is selected
    If TypeName(xFileNames) = "Boolean" Then
        Exit Sub
    End If
    On Error GoTo 0
    
    'Fill in the Tidy_Data_File_Path textbox value
    Tidy_Data_File_Path.Text = Join(xFileNames, ";")
    
    ' If there is an input, the button will be enabled.
    If Tidy_Data_File_Path.Text <> "" Then
        Load_Sample_Annot_Tidy.Create_New_Sample_Annot_Tidy_Button.Enabled = True
    End If

End Sub

'Clear all text when people try to edit the file path
Private Sub Tidy_Data_File_Path_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Load_Sample_Annot_Tidy.Create_New_Sample_Annot_Tidy_Button.Enabled = False
    Tidy_Data_File_Path.Text = ""
End Sub

' Check if input is a positive number, must be integer
Private Sub Starting_Row_Number_TextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Starting_Row_Number_TextBox.Value = "" Then
        MsgBox "Please enter a positive integer"
        Starting_Row_Number_TextBox.SetFocus
        Cancel = True
    ElseIf Starting_Row_Number_TextBox.Value Like "*[!0-9]*" Then
        MsgBox "Please enter a positive integer"
        Starting_Row_Number_TextBox.SetFocus
        Cancel = True
    ElseIf Starting_Row_Number_TextBox.Value <= 0 Or Not IsNumeric(Starting_Row_Number_TextBox.Value) Then
        MsgBox "Please enter a positive integer"
        Starting_Row_Number_TextBox.SetFocus
        Cancel = True
    End If
End Sub

' Check if input is a positive number, must be integer
Private Sub Starting_Column_Number_TextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Starting_Column_Number_TextBox.Value = "" Then
        MsgBox "Please enter a positive integer"
        Starting_Column_Number_TextBox.SetFocus
        Cancel = True
    ElseIf Starting_Column_Number_TextBox.Value Like "*[!0-9]*" Then
        MsgBox "Please enter a positive integer"
        Starting_Column_Number_TextBox.SetFocus
        Cancel = True
    ElseIf Starting_Column_Number_TextBox.Value <= 0 Or Not IsNumeric(Starting_Column_Number_TextBox.Value) Then
        MsgBox "Please enter a positive integer"
        Starting_Column_Number_TextBox.SetFocus
        Cancel = True
    End If
End Sub

' Change the default values of the strating rows and columns based on which property is choosen.
Private Sub Sample_Name_Property_ComboBox_Change()
    Select Case Sample_Name_Property_ComboBox.SelText
    Case "Read as column variables"
        Starting_Row_Number_TextBox.Value = 1
        Starting_Column_Number_TextBox.Value = 2
    Case "Read as row observations"
        Starting_Row_Number_TextBox.Value = 2
        Starting_Column_Number_TextBox.Value = 1
    End Select

End Sub

' Give default values
Private Sub UserForm_Initialize()
    Data_File_Type_ComboBox.AddItem "csv"
    'Data_File_Type_ComboBox.AddItem "Excel"
    'Take the first option as the default value
    Data_File_Type_ComboBox.ListIndex = 0
    
    Sample_Name_Property_ComboBox.AddItem "Read as column variables"
    Sample_Name_Property_ComboBox.AddItem "Read as row observations"
    'Take the first option as the default value
    Sample_Name_Property_ComboBox.ListIndex = 0
    
    Starting_Row_Number_TextBox.Value = 1
    Starting_Column_Number_TextBox.Value = 2
    
End Sub
