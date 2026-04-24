# Excel VBA: Merge Multiple Excel Files (All Sheets)

This VBA macro allows you to select multiple Excel files and merge all worksheets from those files into a single new workbook.

Each worksheet is copied as a separate sheet (no data appending), and duplicate sheet names are handled automatically.

---

## 🚀 Features

- Select multiple Excel files via file picker
- Merge **all worksheets** from each file
- Keeps sheets separate (no data consolidation)
- Handles duplicate sheet names automatically
- Removes default blank sheets
- Saves merged file in the same folder as source files
- Auto-generates file name with timestamp

---

## 🧾 VBA Code

```vba
Sub Merge_Multiple_Excel_Files_All_Sheets_SaveSameFolder()

    Dim fd As FileDialog
    Dim SelectedFile As Variant
    Dim SourceWB As Workbook
    Dim TargetWB As Workbook
    Dim ws As Worksheet
    Dim SheetName As String
    Dim SavePath As String
    Dim FileName As String
    Dim i As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Create merged workbook
    Set TargetWB = Workbooks.Add

    'File picker
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Excel Files to Merge"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"

        If .Show <> -1 Then Exit Sub
    End With

    'Get folder path from first selected file
    SavePath = Left(fd.SelectedItems(1), InStrRev(fd.SelectedItems(1), "\"))

    'Loop through selected files
    For Each SelectedFile In fd.SelectedItems

        Set SourceWB = Workbooks.Open(SelectedFile, ReadOnly:=True)

        For Each ws In SourceWB.Worksheets

            ws.Copy After:=TargetWB.Sheets(TargetWB.Sheets.Count)

            'Handle duplicate sheet names
            SheetName = ws.Name
            i = 1
            Do While SheetExists(SheetName, TargetWB)
                SheetName = ws.Name & "_" & i
                i = i + 1
            Loop
            TargetWB.Sheets(TargetWB.Sheets.Count).Name = SheetName

        Next ws

        SourceWB.Close False

    Next SelectedFile

    'Remove default blank sheets
    Do While TargetWB.Sheets.Count > 1 And _
             TargetWB.Sheets(1).UsedRange.Count = 1
        TargetWB.Sheets(1).Delete
    Loop

    'Generate file name with timestamp
    FileName = "Merged_File_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"

    'Save merged workbook in same folder
    TargetWB.SaveAs FileName:=SavePath & FileName, _
                    FileFormat:=xlOpenXMLWorkbook

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Merged file saved at:" & vbCrLf & SavePath & FileName, vbInformation

End Sub

Function SheetExists(SheetName As String, wb As Workbook) As Boolean
    On Error Resume Next
    SheetExists = Not wb.Sheets(SheetName) Is Nothing
    On Error GoTo 0
End Function
