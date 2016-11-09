Attribute VB_Name = "Lib"

Function selectFile() As String
    nameOfFile = ""
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .InitialFileName = "*.*"
        .Title = "Select a file"
        .Show
        If .SelectedItems.Count = 1 Then nameOfFile = .SelectedItems(1)
    End With
    selectFile = nameOfFile
End Function

Function getCol(n As Integer, text As String) As Integer
    result = 0
    For i = 1 To lastColumn()
        If Cells(n, i) = text Then
            result = i
            Exit For
        End If
    Next i
    getCol = result
End Function

Function getRow(n As Integer, text As String) As Integer
    result = 0
    For i = 1 To lastRow()
        If Cells(i, n) = text Then
            result = i
            Exit For
        End If
    Next i
    getRow = result
End Function

Function lastColumn() As Integer
    lastColumn = ActiveSheet.UsedRange.Column + ActiveSheet.UsedRange.Columns.Count - 1
End Function

Function lastRow() As Integer
    lastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
End Function


