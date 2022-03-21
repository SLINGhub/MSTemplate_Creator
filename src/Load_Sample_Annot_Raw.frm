VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Load_Sample_Annot_Raw 
   Caption         =   "Load_Sample_Annot_Raw"
   ClientHeight    =   7725
   ClientLeft      =   90
   ClientTop       =   285
   ClientWidth     =   12690
   OleObjectBlob   =   "Load_Sample_Annot_Raw.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Load_Sample_Annot_Raw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@Folder("Load Sample Annotation Functions")

'Public whatsclicked As String
Private master_whatsclicked As String

Public Property Get whatsclicked() As String
    whatsclicked = master_whatsclicked
End Property

Public Property Let whatsclicked(ByVal let_whatsclicked As String)
    master_whatsclicked = let_whatsclicked
End Property

'' Function: Browse_Raw_Data_Click
'' --- Code
''  Private Sub Browse_Raw_Data_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked
''
'' (see Sample_Annot_Browse_Raw_Data_Button.png)
''
'' Users will be asked to choose the input file in tabular form.
'' Once done, the Create new Sample Annotation button will
'' be enabled.
''
Private Sub Browse_Raw_Data_Click()

    Dim xFileNames As Variant
    xFileNames = Application.GetOpenFilename(Title:="Load Raw Data File", MultiSelect:=True)
    
    'When no file is selected
    If TypeName(xFileNames) = "Boolean" Then
        Exit Sub
    End If
    On Error GoTo 0
    
    'Fill in the Raw_Data_File_Path textbox value
    Raw_Data_File_Path.Text = Join(xFileNames, ";")
    
    If Raw_Data_File_Path.Text <> vbNullString And Sample_Name_Text.Text = vbNullString Then
        Load_Sample_Annot_Raw.Create_New_Sample_Annot_Raw_Button.Enabled = True
        Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = False
    End If
    
    If Raw_Data_File_Path.Text <> vbNullString And Sample_Name_Text.Text <> vbNullString Then
        Load_Sample_Annot_Raw.Create_New_Sample_Annot_Raw_Button.Enabled = False
        Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = True
    End If

End Sub

'' Function: Browse_Sample_Annot_Click
'' --- Code
''  Private Sub Browse_Sample_Annot_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked
''
'' (see Sample_Annot_Browse_Sample_Annot_Button.png)
''
'' Users will be asked to choose the input file in tabular form.
'' Once done, the Detected Column section will be filled with
'' the column names if the check box "My annotation file has header at line 1"
'' is checked. Else, it will be filled with "Column 1, Column 2, ..., Column X"
'' where X is the total number of columns the system can find.
''
Private Sub Browse_Sample_Annot_Click()

    Dim xFileName As Variant
    xFileName = Application.GetOpenFilename(Title:="Load Sample Annotation File", MultiSelect:=False)
    
    'When no file is selected
    If TypeName(xFileName) = "Boolean" Then
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim FileExtent As String
    FileExtent = Right$(xFileName, Len(xFileName) - InStrRev(xFileName, "."))
    
    If FileExtent = "xlsx" Or FileExtent = "xls" Then
        MsgBox ("We do not support Excel files")
        Exit Sub
        '    Application.ScreenUpdating = False
        '    Application.EnableEvents = False
        '    Dim src As Workbook
        '    Set src = Workbooks.Open(xFileName, True, True)
        '    If Cells(1, Columns.Count).End(xlToLeft).Column = 1 Then
        '        ReDim first_line(1 To 1, 1 To 1)
        '        first_line(1, 1) = src.Worksheets("sheet1").Range("A1").Value
        '        first_line = Application.Index(first_line, 1, 0)
        '    Else
        '        last_cell = Cells(1, Columns.Count).End(xlToLeft).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        '        first_line = src.Worksheets("sheet1").Range("A1:" & last_cell)
        '        first_line = Application.Index(first_line, 1, 0)
        '    End If
        '    src.Close
    End If
    
    ' Load the file into a string.
    Dim fnum As Variant
    Dim whole_file As Variant
    fnum = FreeFile()
    Open xFileName For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum
    
    ' Break the file into lines.
    Dim Lines() As String
    Lines = Split(whole_file, vbCrLf)
    
    'Fill in the Sample_Annot_File_Path textbox value
    Sample_Annot_File_Path.Text = xFileName
    If Load_Sample_Annot_Raw.Is_Column_Name_Present.Value = True Then
     
        'Get the first line
        Dim first_line() As String
        Dim first_line_index As Long
        
        If FileExtent = "csv" Then
            first_line = Split(Lines(0), ",")
        ElseIf FileExtent = "txt" Then
            first_line = Split(Lines(0), vbTab)
        End If
        
        'Update the Listbox to show the column names detected
        Load_Sample_Annot_Raw.Column_Name_List.Clear
        For first_line_index = LBound(first_line) To UBound(first_line)
            Load_Sample_Annot_Raw.Column_Name_List.AddItem first_line(first_line_index)
        Next first_line_index
    Else
        'We find how many columns does the data has
        Dim MaxColumn As Long
        Dim NumOfColumns As Long
        Dim lines_index As Long
        
        MaxColumn = 0
        For lines_index = 0 To UBound(Lines) - 1
            NumOfColumns = Utilities.Get_String_Array_Len(Split(Lines(lines_index), ","))
            If NumOfColumns > MaxColumn Then
                MaxColumn = NumOfColumns
            End If
        Next lines_index
        
        'Update the Listbox to show the column names detected
        Load_Sample_Annot_Raw.Column_Name_List.Clear
        For lines_index = 1 To MaxColumn
            Load_Sample_Annot_Raw.Column_Name_List.AddItem "Column " & CInt(lines_index)
        Next lines_index
        
    End If
    
End Sub

'' Function: Sample_Amount_Map_Click
'' --- Code
''  Private Sub Sample_Amount_Map_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked
''
'' (see Sample_Annot_Sample_Amount_Map_Button.png)
''
'' This button determines which column in the sample annotation file
'' goes to the "Sample_Amount" column in the "Sample_Annot" sheet
'' given the merge is successful.
''
'' In the example below, as "Cell Number" is highlighted, clicking on this
'' button will fill the Sample_Amount_Text Text Box with the highlighted
'' option "Cell Number"
''
'' (see Sample_Annot_Sample_Amount_Map_Example.png)
''
Private Sub Sample_Amount_Map_Click()
    If Column_Name_List.ListIndex = -1 Then
        Exit Sub
    End If
    Sample_Amount_Text.Text = Column_Name_List.List(Column_Name_List.ListIndex)
End Sub

'' Function: Sample_Name_Map_Click
'' --- Code
''  Private Sub Sample_Name_Map_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked
''
'' (see Sample_Annot_Sample_Name_Map_Button.png)
''
'' This button determines which column in the sample annotation file will be
'' used to merge/join with the Sample Name column from the input raw data.
''
'' In the example below, as "Sample" is highlighted, clicking on this
'' button will fill the Sample_Name_Text Text Box with the highlighted
'' option "Sample"
''
'' (see Sample_Annot_Sample_Name_Map_Example.png)
''
'' In addition, the button "Merge with Sample Annotation" will be enabled
'' but the button "Create new Sample Annotation" will be disabled.
''
Private Sub Sample_Name_Map_Click()
    If Column_Name_List.ListIndex = -1 Then
        Exit Sub
    End If
    Sample_Name_Text.Text = Column_Name_List.List(Column_Name_List.ListIndex)
    
    If Raw_Data_File_Path.Text <> vbNullString Then
        Load_Sample_Annot_Raw.Create_New_Sample_Annot_Raw_Button.Enabled = False
        Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = True
    End If
End Sub

'' Function: ISTD_Mixture_Volume_Map_Click
'' --- Code
''  Private Sub ISTD_Mixture_Volume_Map_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked
''
'' (see Sample_Annot_ISTD_Mixture_Volume_Map_Button.png)
''
'' This button determines which column in the sample annotation file
'' goes to the "ISTD_Mixture_Volume_[uL]" column in the "Sample_Annot" sheet
'' given the merge is successful.
''
'' In the example below, as "ISTD Volume" is highlighted, clicking on this
'' button will fill the ISTD_Mixture_Volume_Text Text Box with the highlighted
'' option "ISTD Volume"
''
'' (see Sample_Annot_ISTD_Mixture_Volume_Map_Example.png)
''
Private Sub ISTD_Mixture_Volume_Map_Click()
    If Column_Name_List.ListIndex = -1 Then
        Exit Sub
    End If
    ISTD_Mixture_Volume_Text.Text = Column_Name_List.List(Column_Name_List.ListIndex)
End Sub

'' Function: Merge_With_Sample_Annot_Button_Click
'' --- Code
''  Private Sub Merge_With_Sample_Annot_Button_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked
''
'' (see Sample_Annot_Merge_With_Sample_Annot_Button.png)
''
'' Public Property whatsclicked = "Merge_With_Sample_Annot_Button"
'' Load_Sample_Annot_Raw Box will be hidden
''
Private Sub Merge_With_Sample_Annot_Button_Click()
    whatsclicked = "Merge_With_Sample_Annot_Button"
    Load_Sample_Annot_Raw.Hide
End Sub

'' Function: Create_New_Sample_Annot_Raw_Button_Click
'' --- Code
''  Private Sub Create_New_Sample_Annot_Raw_Button_Click()
'' ---
''
'' Description:
''
'' Function that controls what happens when the following button is
'' left clicked
''
'' (see Sample_Annot_Create_New_Sample_Annot_Raw_Button.png)
''
'' Public Property whatsclicked = "Create_New_Sample_Annot_Raw_Button"
'' Load_Sample_Annot_Raw Box will be hidden
''
Private Sub Create_New_Sample_Annot_Raw_Button_Click()
    whatsclicked = "Create_New_Sample_Annot_Raw_Button"
    Load_Sample_Annot_Raw.Hide
End Sub

'Private Sub Raw_Data_File_Path_Change()
'    Load_Sample_Annot_Raw.Create_New_Sample_Annot_Raw_Button.Enabled = False
'    Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = False
'End Sub

'Private Sub Sample_Annot_File_Path_Change()
'    Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = False
'    'Update the Listbox will be cleared
'    Load_Sample_Annot_Raw.Column_Name_List.Clear
'    'All other sample annot related textboxes will be cleared
'    Sample_Name_Text.Text = ""
'    Sample_Amount_Text.Text = ""
'    ISTD_Mixture_Volume_Text.Text = ""
'End Sub

'' Function: Raw_Data_File_Path_KeyUp
'' --- Code
''  Private Sub Raw_Data_File_Path_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'' ---
''
'' Description:
''
'' Function that controls what happens when the following text box is
'' edited
''
'' (see Sample_Annot_Raw_Data_File_Path_KeyUp_Text_Box.png)
''
'' The text box will be cleared to prevent an invalid file path.
'' In addition, both buttons "Create new Sample Annotation" and "Merge with Sample Annotation"
'' will be disabled. Users must browse the file path again.
''
Private Sub Raw_Data_File_Path_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Clear all text when people try to edit them
    Load_Sample_Annot_Raw.Create_New_Sample_Annot_Raw_Button.Enabled = False
    Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = False
    Raw_Data_File_Path.Text = vbNullString
End Sub

'' Function: Sample_Annot_File_Path_KeyUp
'' --- Code
''  Private Sub Sample_Annot_File_Path_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'' ---
''
'' Description:
''
'' Function that controls what happens when the following text box is
'' edited
''
'' (see Sample_Annot_Sample_Annot_File_Path_KeyUp_Text_Box.png)
''
'' The text box will be cleared to prevent an invalid file path.
'' In addition, the button "Merge with Sample Annotation"
'' will be disabled. Users must browse the file path again.
'' If there is a raw data file path input,
'' the button "Create new Sample Annotation"
'' will be enabled.
''
'' The Load_Sample_Annot_Raw.Column_Name_List will be cleared as
'' well as the text boxes Sample_Name_Text.Text, Sample_Amount_Text.Text,
'' and ISTD_Mixture_Volume_Text.Text
''
Private Sub Sample_Annot_File_Path_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = False
    If Raw_Data_File_Path.Text <> vbNullString Then
        Load_Sample_Annot_Raw.Create_New_Sample_Annot_Raw_Button.Enabled = True
    End If
    Sample_Annot_File_Path.Text = vbNullString
    'Update the Listbox will be cleared
    Load_Sample_Annot_Raw.Column_Name_List.Clear
    'All other sample annot related textboxes will be cleared
    Sample_Name_Text.Text = vbNullString
    Sample_Amount_Text.Text = vbNullString
    ISTD_Mixture_Volume_Text.Text = vbNullString
End Sub

'' Function: Sample_Name_Text_KeyUp
'' --- Code
''  Private Sub Sample_Name_Text_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'' ---
''
'' Description:
''
'' Function that controls what happens when the following text box is
'' edited
''
'' (see Sample_Annot_Sample_Name_Text_KeyUp_Text_Box.png)
''
'' The text box will be cleared to prevent an invalid input.
'' In addition, the button "Merge with Sample Annotation"
'' will be disabled. If there is a raw data file path input,
'' the button "Create new Sample Annotation"
'' will be enabled.
''
Private Sub Sample_Name_Text_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Clear all text when people try to edit them
    Sample_Name_Text.Text = vbNullString
    Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = False
    If Raw_Data_File_Path.Text <> vbNullString Then
        Load_Sample_Annot_Raw.Create_New_Sample_Annot_Raw_Button.Enabled = True
    End If
End Sub

'' Function: Sample_Amount_Text_KeyUp
'' --- Code
''  Private Sub Sample_Amount_Text_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'' ---
''
'' Description:
''
'' Function that controls what happens when the following text box is
'' edited
''
'' (see Sample_Annot_Sample_Amount_Text_KeyUp_Text_Box.png)
''
'' The text box will be cleared to prevent an invalid input.
''
Private Sub Sample_Amount_Text_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Sample_Amount_Text.Text = vbNullString
End Sub

'' Function: ISTD_Mixture_Volume_Text_KeyUp
'' --- Code
''  Private Sub ISTD_Mixture_Volume_Text_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'' ---
''
'' Description:
''
'' Function that controls what happens when the following text box is
'' edited
''
'' (see Sample_Annot_ISTD_Mixture_Volume_Text_KeyUp_Text_Box.png)
''
'' The text box will be cleared to prevent an invalid input.
''
Private Sub ISTD_Mixture_Volume_Text_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ISTD_Mixture_Volume_Text.Text = vbNullString
End Sub

