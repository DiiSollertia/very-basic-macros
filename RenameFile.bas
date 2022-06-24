Attribute VB_Name = "Module1"
Sub AddRFShortcut()
    Application.OnKey "+{R}", "RenameFile"
End Sub
Sub RmRFShortcut()
    Application.OnKey "+{R}"
End Sub

Sub RenameFile()
Dim Source As Range
Dim OldFile As String
Dim NewFile As String
Dim Dir As String
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
If .Show = -1 Then
    Dir = .SelectedItems(1)
End If
End With
Set Source = Cells(1, 1).CurrentRegion

Range("A:B").Select

For Row = 2 To Source.Rows.Count
    OldFile = Dir & Application.PathSeparator & Cells(Row, 1).Value & Cells(2, 3).Value
    NewFile = Dir & Application.PathSeparator & Cells(Row, 2).Value & Cells(2, 3).Value
    On Error Resume Next
    ' rename files
    Name OldFile As NewFile

Next
MsgBox ("Completed")
End Sub

