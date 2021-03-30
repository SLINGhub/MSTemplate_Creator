VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Autofill_By_Sample_Type 
   Caption         =   "Autofill_By_Sample_Type"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7740
   OleObjectBlob   =   "Autofill_By_Sample_Type.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Autofill_By_Sample_Type"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Autofill_ISTD_Mixture_Volume_Button_Click()
    Call Autofill_Column_By_Sample_Type(Sample_Type:=Sample_Type_ComboBox.Value, _
                                        Header_Name:="ISTD_Mixture_Volume_[uL]", _
                                        Autofill_Value:=ISTD_Mixture_Volume_TextBox.Value)
End Sub

Private Sub Autofill_Sample_Amount_Button_Click()
    Call Autofill_Column_By_Sample_Type(Sample_Type:=Sample_Type_ComboBox.Value, _
                                        Header_Name:="Sample_Amount", _
                                        Autofill_Value:=Sample_Amount_TextBox.Value)
End Sub

Private Sub Autofill_Sample_Amount_Unit_Button_Click()
    Call Autofill_Column_By_Sample_Type(Sample_Type:=Sample_Type_ComboBox.Value, _
                                        Header_Name:="Sample_Amount_Unit", _
                                        Autofill_Value:=Sample_Amount_Unit_ComboBox.Value)
End Sub

Private Sub ISTD_Mixture_Volume_TextBox_Change()
    ' If there is an input for both combo boxes, the Autofill button will be enabled.
    If Sample_Type_ComboBox.Value <> "" And ISTD_Mixture_Volume_TextBox.Value <> "" Then
        Autofill_ISTD_Mixture_Volume_Button.Enabled = True
    Else
        Autofill_ISTD_Mixture_Volume_Button.Enabled = False
    End If
End Sub

Private Sub Sample_Amount_TextBox_Change()
    ' If there is an input for both combo boxes, the Autofill button will be enabled.
    If Sample_Type_ComboBox.Value <> "" And Sample_Amount_TextBox.Value <> "" Then
        Autofill_Sample_Amount_Button.Enabled = True
    Else
        Autofill_Sample_Amount_Button.Enabled = False
    End If
End Sub

' Check if input is a positive number can be decimal
Private Sub ISTD_Mixture_Volume_TextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ISTD_Mixture_Volume_TextBox.Value = "" Then
        Exit Sub
    ElseIf ISTD_Mixture_Volume_TextBox.Value <= 0 Or Not IsNumeric(ISTD_Mixture_Volume_TextBox.Value) Then
        MsgBox "Please enter a positive number"
        ISTD_Mixture_Volume_TextBox.SetFocus
        Cancel = True
    End If
End Sub

' Check if input is a positive number can be decimal
Private Sub Sample_Amount_TextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Sample_Amount_TextBox.Value = "" Then
        Exit Sub
    ElseIf Sample_Amount_TextBox.Value <= 0 Or Not IsNumeric(Sample_Amount_TextBox.Value) Then
        MsgBox "Please enter a positive number"
        Sample_Amount_TextBox.SetFocus
        Cancel = True
    End If
End Sub

Private Sub Sample_Type_ComboBox_Change()
    ' If there is an input for both combo boxes, the Autofill button will be enabled.
    If Sample_Type_ComboBox.Value <> "" And Sample_Amount_Unit_ComboBox.Value <> "" Then
        Autofill_Sample_Amount_Unit_Button.Enabled = True
    End If
End Sub

Private Sub Sample_Amount_Unit_ComboBox_Change()
    ' If there is an input for both combo boxes, the Autofill button will be enabled.
    If Sample_Type_ComboBox.Value <> "" And Sample_Amount_Unit_ComboBox.Value <> "" Then
        Autofill_Sample_Amount_Unit_Button.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim cLoc As Range
    Dim ws As Worksheet
    Set ws = Worksheets("Lists")
    
    Sample_Type_ComboBox.AddItem "All Sample Types"
    
    For Each cLoc In ws.Range("SampleType")
         With Me.Sample_Type_ComboBox
              .AddItem cLoc.Value
         End With
    Next cLoc
    
    For Each cLoc In ws.Range("SampleAmountUnit")
         With Me.Sample_Amount_Unit_ComboBox
              .AddItem cLoc.Value
         End With
    Next cLoc
    
End Sub
