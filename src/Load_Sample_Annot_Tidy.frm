VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Load_Sample_Annot_Tidy 
   Caption         =   "Load_Sample_Annot_Table"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   330
   ClientWidth     =   9300.001
   OleObjectBlob   =   "Load_Sample_Annot_Tidy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Load_Sample_Annot_Tidy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Functions that control what happens when buttons in the Load Sample Annot Tidy Box are clicked."
Option Explicit
'@ModuleDescription("Functions that control what happens when buttons in the Load Sample Annot Tidy Box are clicked.")
'@Folder("Load Sample Annotation Functions")

'Public whatsclicked As String
Private master_whatsclicked As String

Public Property Get whatsclicked() As String
    whatsclicked = master_whatsclicked
End Property

Public Property Let whatsclicked(ByVal let_whatsclicked As String)
    master_whatsclicked = let_whatsclicked
End Property

'@Description("Function that controls what happens when the Create New Sample Annot button is left clicked.")

'' Group: Functions
''
'' Function: Create_New_Sample_Annot_Tidy_Button_Click
'' --- Code
''  Private Sub Create_New_Sample_Annot_Tidy_Button_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the
'' when the following button is
'' left clicked.
''
'' (see Sample_Annot_Create_New_Sample_Annot_Button.png)
''
'' Public Property whatsclicked = "Create_New_Sample_Annot_Tidy_Button"
'' Load_Sample_Annot_Tidy Box will be hidden
''
Private Sub Create_New_Sample_Annot_Tidy_Button_Click()
Attribute Create_New_Sample_Annot_Tidy_Button_Click.VB_Description = "Function that controls what happens when the Create New Sample Annot button is left clicked."
    whatsclicked = "Create_New_Sample_Annot_Tidy_Button"
    Load_Sample_Annot_Tidy.Hide
End Sub

'@Description("Function that controls what happens when the Browse button is left clicked.")

'' Function: Browse_Tidy_Data_Click
'' --- Code
''  Private Sub Browse_Tidy_Data_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked.
''
'' (see Sample_Annot_Browse_Tidy_Data_Button.png)
''
'' Users will be asked to choose the input file in tabular form.
'' Once done, the Create new Sample Annotation button will
'' be enabled.
''
Private Sub Browse_Tidy_Data_Click()
Attribute Browse_Tidy_Data_Click.VB_Description = "Function that controls what happens when the Browse button is left clicked."

    ' Load the file path of the tidy data
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
    If Tidy_Data_File_Path.Text <> vbNullString Then
        Load_Sample_Annot_Tidy.Create_New_Sample_Annot_Tidy_Button.Enabled = True
    End If

End Sub

'@Description("Function that controls what happens when the data file path text box is edited.")

'' Function: Tidy_Data_File_Path_KeyUp
'' --- Code
''  Private Sub Tidy_Data_File_Path_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'' ---
''
'' Description:
''
'' Function that controls what happens when the following text box is
'' edited.
''
'' (see Sample_Annot_Tidy_Data_File_Path_KeyUp_Text_Box.png)
''
'' The text box will be cleared to prevent an invalid file path.
'' The Create new Sample Annotation button will be disabled.
''
Private Sub Tidy_Data_File_Path_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Attribute Tidy_Data_File_Path_KeyUp.VB_Description = "Function that controls what happens when the data file path text box is edited."
    'Clear all text when people try to edit the file path
    Load_Sample_Annot_Tidy.Create_New_Sample_Annot_Tidy_Button.Enabled = False
    Tidy_Data_File_Path.Text = vbNullString
End Sub

'@Description("Function that controls what happens when the starting row number text box is exited.")

'' Function: Starting_Row_Number_TextBox_Exit
'' --- Code
''  Private Sub Starting_Row_Number_TextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'' ---
''
'' Description:
''
'' Function that controls what happens when the following text box is
'' exited.
''
'' (see Sample_Annot_Starting_Row_Number_Text_Box.png)
''
'' The system will check if the input is valid. Invalid inputs will
'' be given this message box error.
''
'' (see Transition_Annot_Starting_Row_Number_Text_Box_Invalid_Input.png)
''
Private Sub Starting_Row_Number_TextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Attribute Starting_Row_Number_TextBox_Exit.VB_Description = "Function that controls what happens when the starting row number text box is exited."
    ' Check if input is a positive number, must be integer
    If Starting_Row_Number_TextBox.Value = vbNullString Then
        MsgBox "Please enter a positive integer"
        Starting_Row_Number_TextBox.SetFocus
        Cancel.Value = True
    ElseIf Starting_Row_Number_TextBox.Value Like "*[!0-9]*" Then
        MsgBox "Please enter a positive integer"
        Starting_Row_Number_TextBox.SetFocus
        Cancel.Value = True
    ElseIf Starting_Row_Number_TextBox.Value <= 0 Or Not IsNumeric(Starting_Row_Number_TextBox.Value) Then
        MsgBox "Please enter a positive integer"
        Starting_Row_Number_TextBox.SetFocus
        Cancel.Value = True
    End If
End Sub

'@Description("Function that controls what happens when the starting column number text box is exited.")

'' Function: Starting_Column_Number_TextBox_Exit
'' --- Code
''  Private Sub Starting_Column_Number_TextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'' ---
''
'' Description:
''
'' Function that controls what happens when the following text box is
'' exited.
''
'' (see Sample_Annot_Starting_Column_Number_Text_Box.png)
''
'' The system will check if the input is valid. Invalid inputs will
'' be given this message box error.
''
'' (see Transition_Annot_Starting_Row_Number_Text_Box_Invalid_Input.png)
''
Private Sub Starting_Column_Number_TextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Attribute Starting_Column_Number_TextBox_Exit.VB_Description = "Function that controls what happens when the starting column number text box is exited."
    ' Check if input is a positive number, must be integer
    If Starting_Column_Number_TextBox.Value = vbNullString Then
        MsgBox "Please enter a positive integer"
        Starting_Column_Number_TextBox.SetFocus
        Cancel.Value = True
    ElseIf Starting_Column_Number_TextBox.Value Like "*[!0-9]*" Then
        MsgBox "Please enter a positive integer"
        Starting_Column_Number_TextBox.SetFocus
        Cancel.Value = True
    ElseIf Starting_Column_Number_TextBox.Value <= 0 Or Not IsNumeric(Starting_Column_Number_TextBox.Value) Then
        MsgBox "Please enter a positive integer"
        Starting_Column_Number_TextBox.SetFocus
        Cancel.Value = True
    End If
End Sub

'@Description("Function that controls what happens when the sample annot property combo box is changed.")

'' Function: Sample_Name_Property_ComboBox_Change
'' --- Code
''  Private Sub Sample_Name_Property_ComboBox_Change()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following combo box is
'' changed.
''
'' (see Sample_Annot_Property_Combo_Box.png)
''
'' If Read as column variables is chosen, the starting row number
'' will be set to 1 while the starting column number will be set to 2
''
'' (see Sample_Annot_Load_Transition_Name_Tidy_Pop_Up.png)
''
'' If Read as row observations is chosen, the starting row number
'' will be set to 2 while the starting column number will be set to 1
''
'' (see Sample_Annot_Load_Sample_Annot_Tidy_Column_Name_Pop_Up2.png)
''
Private Sub Sample_Name_Property_ComboBox_Change()
Attribute Sample_Name_Property_ComboBox_Change.VB_Description = "Function that controls what happens when the sample annot property combo box is changed."
    ' Change the default values of the strating rows and columns based on which property is chosen.
    Select Case Sample_Name_Property_ComboBox.SelText
    Case "Read as column variables"
        Starting_Row_Number_TextBox.Value = 1
        Starting_Column_Number_TextBox.Value = 2
    Case "Read as row observations"
        Starting_Row_Number_TextBox.Value = 2
        Starting_Column_Number_TextBox.Value = 1
    End Select

End Sub

'@Description("Function that controls what happens when Load_Sample_Annot_Tidy form is initialized.")

'' Group: Form Initialisation
''
'' Function: Load_Sample_Annot_Tidy form initialisation
'' --- Code
''  Private Sub UserForm_Initialize()
'' ---
''
'' Description:
''
'' Function that controls what happens when the
'' Load_Sample_Annot_Tidy form is initialize when
'' user click on the button "Load Sample Annotation from Table Data"
'' in the Transition_Name_Annot sheet.
''
'' The function will create the Data_File_Type combo box.
'' Currently, only "csv" is added in the dropdown and hence it
'' is also the default value.
''
'' It also create the Sample_Name_Property combo box by
'' adding entries "Read as column variables" and
'' "Read as row observations". "Read as column variables"
'' is the default.
''
'' Next, it will set the Starting_Row_Number text box value as 1
'' and Starting_Row_Number text box as 2.
''
Private Sub UserForm_Initialize()
Attribute UserForm_Initialize.VB_Description = "Function that controls what happens when Load_Sample_Annot_Tidy form is initialized."
    ' Give default values
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
