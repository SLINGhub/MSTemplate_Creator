VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Load_Sample_Annot_Raw 
   Caption         =   "Load_Sample_Annot_Raw"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12525
   OleObjectBlob   =   "Load_Sample_Annot_Raw.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Load_Sample_Annot_Raw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public whatsclicked As String

Private Sub Browse_Raw_Data_Click()
    xFileNames = Application.GetOpenFilename(Title:="Load Raw Data File", MultiSelect:=True)
    
    'When no file is selected
    If TypeName(xFileNames) = "Boolean" Then
        Exit Sub
    End If
    On Error GoTo 0
    
    'Fill in the Raw_Data_File_Path textbox value
    Raw_Data_File_Path.Text = Join(xFileNames, ";")
    
    If Raw_Data_File_Path.Text <> "" And Sample_Name_Text.Text = "" Then
        Load_Sample_Annot_Raw.Create_New_Sample_Annot_Raw_Button.Enabled = True
        Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = False
    End If
    
    If Raw_Data_File_Path.Text <> "" And Sample_Name_Text.Text <> "" Then
        Load_Sample_Annot_Raw.Create_New_Sample_Annot_Raw_Button.Enabled = False
        Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = True
    End If

End Sub

Private Sub Browse_Sample_Annot_Click()
    xFileName = Application.GetOpenFilename(Title:="Load Sample Annotation File", MultiSelect:=False)
    
    'When no file is selected
    If TypeName(xFileName) = "Boolean" Then
        Exit Sub
    End If
    On Error GoTo 0
    
    'Dim FileExtent As String
    FileExtent = Right(xFileName, Len(xFileName) - InStrRev(xFileName, "."))
    
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
    fnum = FreeFile
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
        If FileExtent = "csv" Then
            first_line = Split(Lines(0), ",")
        ElseIf FileExtent = "txt" Then
            first_line = Split(Lines(0), vbTab)
        End If
        
        'Update the Listbox to show the column names detected
        Load_Sample_Annot_Raw.Column_Name_List.Clear
        For i = LBound(first_line) To UBound(first_line)
            Load_Sample_Annot_Raw.Column_Name_List.AddItem first_line(i)
        Next i
    Else
        'We find how many columns does the data has
        Dim MaxColumn As Integer
        Dim NumOfColumns As Integer
        MaxColumn = 0
        For j = 0 To UBound(Lines) - 1
            NumOfColumns = Utilities.StringArrayLen(Split(Lines(i), ","))
            If NumOfColumns > MaxColumn Then
                MaxColumn = NumOfColumns
            End If
        Next j
        
        'Update the Listbox to show the column names detected
        Load_Sample_Annot_Raw.Column_Name_List.Clear
        For j = 1 To MaxColumn
            Load_Sample_Annot_Raw.Column_Name_List.AddItem "Column " & CInt(j)
        Next j
        
    End If
    
End Sub

Private Sub Sample_Amount_Map_Click()
    If Column_Name_List.ListIndex = -1 Then
        Exit Sub
    End If
    Sample_Amount_Text.Text = Column_Name_List.List(Column_Name_List.ListIndex)
End Sub

Private Sub Sample_Name_Map_Click()
    If Column_Name_List.ListIndex = -1 Then
        Exit Sub
    End If
    Sample_Name_Text.Text = Column_Name_List.List(Column_Name_List.ListIndex)
    
    If Raw_Data_File_Path.Text <> "" Then
        Load_Sample_Annot_Raw.Create_New_Sample_Annot_Raw_Button.Enabled = False
        Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = True
    End If
End Sub

Private Sub ISTD_Mixture_Volume_Map_Click()
    If Column_Name_List.ListIndex = -1 Then
        Exit Sub
    End If
    ISTD_Mixture_Volume_Text.Text = Column_Name_List.List(Column_Name_List.ListIndex)
End Sub

Private Sub Merge_With_Sample_Annot_Button_Click()
    whatsclicked = "Merge_With_Sample_Annot_Button"
    Load_Sample_Annot_Raw.Hide
End Sub

Private Sub Create_New_Sample_Annot_Raw_Button_Click()
    whatsclicked = "Create_New_Sample_Annot_Raw_Button"
    Load_Sample_Annot_Raw.Hide
End Sub

Private Sub Raw_Data_File_Path_Change()
    Load_Sample_Annot_Raw.Create_New_Sample_Annot_Raw_Button.Enabled = False
    Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = False
End Sub

'Private Sub Sample_Annot_File_Path_Change()
'    Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = False
'    'Update the Listbox will be cleared
'    Load_Sample_Annot_Raw.Column_Name_List.Clear
'    'All other sample annot related textboxes will be cleared
'    Sample_Name_Text.Text = ""
'    Sample_Amount_Text.Text = ""
'    ISTD_Mixture_Volume_Text.Text = ""
'End Sub

'Clear all text when people try to edit them
Private Sub Raw_Data_File_Path_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Load_Sample_Annot_Raw.Create_New_Sample_Annot_Raw_Button.Enabled = False
    Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = False
    Raw_Data_File_Path.Text = ""
End Sub

Private Sub Sample_Annot_File_Path_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = False
    Sample_Annot_File_Path.Text = ""
    'Update the Listbox will be cleared
    Load_Sample_Annot_Raw.Column_Name_List.Clear
    'All other sample annot related textboxes will be cleared
    Sample_Name_Text.Text = ""
    Sample_Amount_Text.Text = ""
    ISTD_Mixture_Volume_Text.Text = ""
End Sub

'Clear all text when people try to edit them
Private Sub Sample_Name_Text_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Sample_Name_Text.Text = ""
    Load_Sample_Annot_Raw.Merge_With_Sample_Annot_Button.Enabled = False
    If Raw_Data_File_Path.Text <> "" Then
        Load_Sample_Annot_Raw.Create_New_Sample_Annot_Raw_Button.Enabled = True
    End If
End Sub

Private Sub Sample_Amount_Text_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Sample_Amount_Text.Text = ""
End Sub

Private Sub ISTD_Mixture_Volume_Text_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ISTD_Mixture_Volume_Text.Text = ""
End Sub

