Attribute VB_Name = "Module1"
Option Explicit
'objects
Dim f_dialog As Object
Dim target, source As Workbook
Dim new_sheet As Worksheet
'others
Dim file_name As String
'iteration variables
Dim i As Integer
Dim file As Variant


Sub import_csv()
Application.ScreenUpdating = False
Set target = ThisWorkbook

'Create file dialog object
Set f_dialog = Application.FileDialog(msoFileDialogFilePicker)
With f_dialog
    .Title = "Please select the files you want to import"
    .AllowMultiSelect = True
    .Filters.Clear
    'Add filter for csv files
    .Filters.Add "Excel-Files", "*.csv; *.xlsx; *.xls", 1

    'Show the file dialog
    If f_dialog.Show = -1 Then
        'Exit sub if no files have been chosen
        If .SelectedItems.Count < 1 Then
            MsgBox "No files were selected"
            Exit Sub
        End If
    End If
        
        'Loop through the selected files and import them
        For Each file In .SelectedItems
            file_name = Dir(file)
            Set source = Workbooks.Open(file)
            Set new_sheet = target.Sheets.Add(After:=target.Worksheets(target.Worksheets.Count))
            new_sheet.Name = file_name
            source.Worksheets(1).Cells.Copy Destination:=new_sheet.Range("A1")

            'Close the current file and open the next one
            source.Close savechanges:=False
        Next
        
End With

Application.ScreenUpdating = True
End Sub
