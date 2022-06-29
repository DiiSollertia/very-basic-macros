Attribute VB_Name = "Module1"
Sub Name2Name()
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

' automatic importing of filenames, not recommended
If IsEmpty(Cells(2, 1)) Then
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(Dir)
 
    For Each oFile In oFolder.Files
        Cells(i + 2, 1) = oFSO.GetBaseName(oFile)
        i = i + 1
        Next oFile
End If

For Row = 2 To Source.Rows.Count
    OldFile = Dir & Application.PathSeparator & Cells(Row, 1).Value & Cells(2, 3).Value
    NewFile = Dir & Application.PathSeparator & Cells(Row, 2).Value & Cells(2, 3).Value
    On Error Resume Next
    ' rename files
    Name OldFile As NewFile
    Next

MsgBox ("Completed")
End Sub
Sub CommonString()
Dim Source As Range
Dim OldFile As String
Dim NewFile As String
Dim Extension As String
Dim ComStr As String
Dim Dir As String
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
If .Show = -1 Then
    Dir = .SelectedItems(1)
End If
End With
Set Source = Cells(1, 1).CurrentRegion

Range("A:B").Select

' automatic importing of filenames, not recommended
If IsEmpty(Cells(2, 1)) Then
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(Dir)
 
    For Each oFile In oFolder.Files
        Cells(i + 2, 1) = oFSO.GetBaseName(oFile)
        i = i + 1
        Next oFile
End If

Extension = Cells(6, 2).Value
ComStr = Cells(2, 2).Value

Select Case Cells(4, 2).Value
Case "APPEND"
    For Row = 2 To Source.Rows.Count
        OldFile = Dir & Application.PathSeparator & Cells(Row, 1).Value
        NewFile = OldFile & ComStr
        On Error Resume Next
        ' rename files
        Name OldFile & Extension As NewFile & Extension
        Next
Case "ADD BEFORE"
    For Row = 2 To Source.Rows.Count
        OldFile = Dir & Application.PathSeparator & Cells(Row, 1).Value & Extension
        NewFile = Dir & Application.PathSeparator & ComStr & Cells(Row, 1).Value & Extension
        On Error Resume Next
        ' rename files
        Name OldFile As NewFile
        Next
Case "REMOVE"
    For Row = 2 To Source.Rows.Count
        OldFile = Dir & Application.PathSeparator & Cells(Row, 1).Value & Extension
        NewFile = Replace(OldFile, ComStr, "")
        On Error Resume Next
        ' rename files
        Name OldFile As NewFile
        On Error Resume Next
        Next
End Select

MsgBox ("Completed")
End Sub
Sub ClrNames()
Dim Source As Range
Dim LastRow As Long


Select Case ActiveSheet.Name
Case "Name2Name"
    With ActiveSheet
        LastRow = WorksheetFunction.Max(.Cells(.Rows.Count, "A").End(xlUp).Row, .Cells(.Rows.Count, "B").End(xlUp).Row)
        If LastRow > 1 Then
            Set Source = .Range("A2:B" & LastRow)
            Source.Clear
        End If
    End With
    
Case "CommonString"
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If LastRow > 1 Then
            Set Source = .Range("A2:A" & LastRow)
            Source.Clear
        End If
        Cells(2, 2).Clear
    End With
End Select

MsgBox ("File Names Cleared")
End Sub

