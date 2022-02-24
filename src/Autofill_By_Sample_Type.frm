VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Autofill_By_Sample_Type 
   Caption         =   "Autofill_By_Sample_Type"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   408
   ClientWidth     =   6180
   OleObjectBlob   =   "Autofill_By_Sample_Type.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Autofill_By_Sample_Type"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@Folder("Sample_Annot_Buttons")

'' Group: Autofill Button Function
''
'' Function: Autofill_ISTD_Mixture_Volume_Button_Click
'' --- Code
''  Private Sub Autofill_ISTD_Mixture_Volume_Button_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked
''
'' (see Autofill_ISTD_Mixture_Volume_Button.png)
''
''
'' Currently, clicking on this button will autofill the column
'' ISTD_Mixture_Volume_[uL] in the active sheet only rows whose
'' column Sample_Type matches the input Sample_Type based on the
'' autofill value in the input ISTD_Mixture_Volume text box value
''
Private Sub Autofill_ISTD_Mixture_Volume_Button_Click()
    Sample_Annot.Autofill_Column_By_QC_Sample_Type Sample_Type:=Sample_Type_ComboBox.Value, _
                                                   Header_Name:="ISTD_Mixture_Volume_[uL]", _
                                                   Autofill_Value:=ISTD_Mixture_Volume_TextBox.Value
End Sub

'' Function: Autofill_Sample_Amount_Button_Click
'' --- Code
''  Private Sub Autofill_Sample_Amount_Button_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked
''
'' (see Autofill_Sample_Amount_Button_Click.png)
''
''
'' Currently, clicking on this button will autofill the column
'' Sample_Amount in the active sheet only rows whose
'' column Sample_Type matches the input Sample_Type based on the
'' autofill value in the input Sample_Amount text box value
''
Private Sub Autofill_Sample_Amount_Button_Click()
    Sample_Annot.Autofill_Column_By_QC_Sample_Type Sample_Type:=Sample_Type_ComboBox.Value, _
                                                   Header_Name:="Sample_Amount", _
                                                   Autofill_Value:=Sample_Amount_TextBox.Value
End Sub

'' Function: Autofill_Sample_Amount_Unit_Button_Click
'' --- Code
''  Private Sub Autofill_Sample_Amount_Unit_Button_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked
''
'' (see Autofill_Sample_Amount_Unit_Button.png)
''
''
'' Currently, clicking on this button will autofill the column
'' Sample_Amount Unit in the active sheet only rows whose
'' column Sample_Type matches the input Sample_Type based on the
'' autofill value in the input Sample_Amount_Unit text box value
''
Private Sub Autofill_Sample_Amount_Unit_Button_Click()
    Sample_Annot.Autofill_Column_By_QC_Sample_Type Sample_Type:=Sample_Type_ComboBox.Value, _
                                                   Header_Name:="Sample_Amount_Unit", _
                                                   Autofill_Value:=Sample_Amount_Unit_ComboBox.Value
End Sub

'' Group: Autofill Text Box Change Function
''
'' Function: ISTD_Mixture_Volume_TextBox_Change
'' --- Code
''  Private Sub ISTD_Mixture_Volume_TextBox_Change()
'' ---
''
'' Description:
''
'' Function that controls what happens when the users
'' change the ISTD_Mixture_Volume text box value
''
'' By default, the button will be masked. However, if there is
'' an input in *both* the Sample_Type and the ISTD_Mixture_Volume
'' boxes, the ISTD_Mixture_Volume Autofill button will be enabled.
''
'' (see Autofill_ISTD_Mixture_Volume_Text_Change.png)
''
Private Sub ISTD_Mixture_Volume_TextBox_Change()
    ' If there is an input for both combo boxes, the Autofill button will be enabled.
    Autofill_ISTD_Mixture_Volume_Button.Enabled = False
    If Sample_Type_ComboBox.Value <> vbNullString And ISTD_Mixture_Volume_TextBox.Value <> vbNullString Then
        Autofill_ISTD_Mixture_Volume_Button.Enabled = True
    End If
End Sub

'' Function: Sample_Amount_TextBox_Change
'' --- Code
''  Private Sub Sample_Amount_TextBox_Change()
'' ---
''
'' Description:
''
'' Function that controls what happens when the users
'' change the Sample_Amount text box value
''
'' By default, the button will be masked. However, if there is
'' an input in *both* the Sample_Type and the Sample_Amount
'' boxes, the Sample_Amount Autofill button will be enabled.
''
'' (see Autofill_Sample_Amount_Text_Change.png)
''
Private Sub Sample_Amount_TextBox_Change()
    ' If there is an input for both combo boxes, the Autofill button will be enabled.
    Autofill_Sample_Amount_Button.Enabled = False
    If Sample_Type_ComboBox.Value <> vbNullString And Sample_Amount_TextBox.Value <> vbNullString Then
        Autofill_Sample_Amount_Button.Enabled = True
    End If
End Sub

'' Group: Autofill Combo Box Change Function
''
'' Function: Sample_Amount_Unit_ComboBox_Change
'' --- Code
''  Private Sub Sample_Amount_Unit_ComboBox_Change()
'' ---
''
'' Description:
''
'' Function that controls what happens when the users
'' change the Sample_Amount_Unit combo box value
''
'' By default, the button will be masked. However, if there is
'' an input in *both* the Sample_Type and the Sample_Amount_Unit
'' boxes, the Sample_Amount_Unit Autofill button will be enabled.
''
'' (see Autofill_Sample_Amount_Unit_ComboBox_Change.png)
''
Private Sub Sample_Amount_Unit_ComboBox_Change()
    Autofill_Sample_Amount_Unit_Button.Enabled = False
    ' If there is an input for both combo boxes, the Autofill button will be enabled.
    If Sample_Type_ComboBox.Value <> vbNullString And Sample_Amount_Unit_ComboBox.Value <> vbNullString Then
        Autofill_Sample_Amount_Unit_Button.Enabled = True
    End If
End Sub

'' Function: Sample_Type_ComboBox_Change
'' --- Code
''  Private Sub Sample_Type_ComboBox_Change()
'' ---
''
'' Description:
''
'' Function that controls what happens when the users
'' change the Sample_Type combo box value
''
'' By default, the button will be masked. However, if there is
'' an input in *both* the Sample_Type and the corresponding
'' (Sample_Amount, Sample_Amount_Unit or ISTD_Mixture_Volume)
'' boxes, the corresponding autofill buttons
'' (Sample_Amount, Sample_Amount_Unit or ISTD_Mixture_Volume)
'' will be enabled.
''
'' (see Autofill_Sample_Type_ComboBox_Change.png)
''
Private Sub Sample_Type_ComboBox_Change()
    ' If there is an input for both combo boxes, the Autofill button will be enabled.
    If Sample_Type_ComboBox.Value <> vbNullString And ISTD_Mixture_Volume_TextBox.Value <> vbNullString Then
        Autofill_ISTD_Mixture_Volume_Button.Enabled = True
    End If
    If Sample_Type_ComboBox.Value <> vbNullString And Sample_Amount_TextBox.Value <> vbNullString Then
        Autofill_Sample_Amount_Button.Enabled = True
    End If
    If Sample_Type_ComboBox.Value <> vbNullString And Sample_Amount_Unit_ComboBox.Value <> vbNullString Then
        Autofill_Sample_Amount_Unit_Button.Enabled = True
    End If
End Sub

'' Group: Autofill Text Box Exit Function
''
'' Function: ISTD_Mixture_Volume_TextBox_Exit
'' --- Code
''  Private Sub ISTD_Mixture_Volume_TextBox_Exit()
'' ---
''
'' Description:
''
'' Function that controls what happens when the users
'' leaves the ISTD_Mixture_Volume text box
''
'' The function will check if a valid positive number
'' is keyed in. If an invalid value is given, the following
'' message box will pop out to inform the user to enter
'' a valid input.
''
'' (see Autofill_ISTD_Mixture_Volume_Text_Exit.png)
''
Private Sub ISTD_Mixture_Volume_TextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Check if input is a positive real number
    If ISTD_Mixture_Volume_TextBox.Value = vbNullString Then
        ' If no input is given, leave the function
        Exit Sub
    ElseIf ISTD_Mixture_Volume_TextBox.Value <= 0 Or Not IsNumeric(ISTD_Mixture_Volume_TextBox.Value) Then
        ' Give an error if input is not a positive real number
        ' and force the user to enter again
        MsgBox "Please enter a positive number"
        ISTD_Mixture_Volume_TextBox.SetFocus
        Cancel.Value = True
    End If
End Sub

'' Function: Sample_Amount_TextBox_Exit
'' --- Code
''  Private Sub Sample_Amount_TextBox_Exit()
'' ---
''
'' Description:
''
'' Function that controls what happens when the users
'' leaves the Sample_Amount text box
''
'' The function will check if a valid positive number
'' is keyed in. If an invalid value is given, the following
'' message box will pop out to inform the user to enter
'' a valid input.
''
'' (see Autofill_Sample_Amount_Text_Exit.png)
''
Private Sub Sample_Amount_TextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Check if input is a positive real number
    If Sample_Amount_TextBox.Value = vbNullString Then
        ' If no input is given, leave the function
        Exit Sub
    ElseIf Sample_Amount_TextBox.Value <= 0 Or Not IsNumeric(Sample_Amount_TextBox.Value) Then
        ' Give an error if input is not a positive real number
        ' and force the user to enter again
        MsgBox "Please enter a positive number"
        Sample_Amount_TextBox.SetFocus
        Cancel.Value = True
    End If
End Sub

'' Group: Form Initialisation
''
'' Function: Autofill_By_Sample_Type form initialisation
'' --- Code
''  Private Sub UserForm_Initialize()
'' ---
''
'' Description:
''
'' Function that controls what happens when the
'' Autofill_By_Sample_Type form is initialize when
'' user click on the button "Autofill by Sample_Type"
'' in the Sample_Annot sheet
''
'' The function will first check of if the sheet whose
'' code name is Lists exists before proceeding. If it does,
'' It then fills in the combox boxes (Sample_Type and Sample_Amount_Unit)
'' with the relevant entries provided in the respective tables
'' (SampleType and SampleAmountUnit)
'' found in the "List" worksheet.
''
Private Sub UserForm_Initialize()
    ' UserForm_Initialize must have no parameters
    Dim Cell_Location As Range
    
    ' Get the Lists worksheet from the active workbook
    ' Lists is a code name
    ' Refer to https://riptutorial.com/excel-vba/example/11272/worksheet--name---index-or--codename
    Dim Lists_Worksheet As Worksheet
    
    ' Check if the sheet whose code name is Lists exists
    ' If not, give an error message
    If Utilities.Check_Sheet_Code_Name_Exists(ActiveWorkbook, "Lists") = False Then
        MsgBox ("Sheet containing list of Sample Types is missing")
        Exit Sub
    End If
    
    Set Lists_Worksheet = Lists
    
    ' In the Sample_Type_ComboBox, add "All Sample Types"
    ' as the first item in the drop down list
    Sample_Type_ComboBox.AddItem "All Sample Types"
    
    ' Add in the Sample_Type_ComboBox, the list given
    ' in the Table called SampleType found in the "Lists" sheet
    For Each Cell_Location In Lists_Worksheet.Range("SampleType")
         With Me.Sample_Type_ComboBox
              .AddItem Cell_Location.Value
         End With
    Next Cell_Location
    
    ' Add in the Sample_Amount_Unit_ComboBox, the list given
    ' in the Table called SampleAmountUnit found in the "Lists" sheet
    For Each Cell_Location In Lists_Worksheet.Range("SampleAmountUnit")
         With Me.Sample_Amount_Unit_ComboBox
              .AddItem Cell_Location.Value
         End With
    Next Cell_Location
    
End Sub
