VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Autofill_Sample_Amount_Unit 
   Caption         =   "Autofill_Sample_Amount_Unit"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4035
   OleObjectBlob   =   "AutoFill_Sample_Amount_Unit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AutoFill_Sample_Amount_Unit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public whatsclicked As String

Private Sub Autofill_Sample_Amount_Unit_Button_Click()
    Call Autofill_Sample_Unit(Sample_Type:=Sample_Type_ComboBox.Value, _
                              Sample_Amount_Unit:=Sample_Amount_Unit_ComboBox.Value)
End Sub

Private Sub Sample_Type_ComboBox_Change()
    ' If there is an input for both combo boxes, the Autofill button will be enabled.
    If Sample_Type_ComboBox.Value <> "" And Sample_Amount_Unit_ComboBox.Value <> "" Then
        AutoFill_Sample_Amount_Unit.Autofill_Sample_Amount_Unit_Button.Enabled = True
    End If
End Sub

Private Sub Sample_Amount_Unit_ComboBox_Change()
    ' If there is an input for both combo boxes, the Autofill button will be enabled.
    If Sample_Type_ComboBox.Value <> "" And Sample_Amount_Unit_ComboBox.Value <> "" Then
        AutoFill_Sample_Amount_Unit.Autofill_Sample_Amount_Unit_Button.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim cLoc As Range
    Dim ws As Worksheet
    Set ws = Worksheets("Lists")
    
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
