Attribute VB_Name = "FileTools"
Option Explicit
'@Folder("Library")
'@IgnoreModule ProcedureNotUsed, UseMeaningfulName

Public Function GetDirectory() As String
    Dim file As FileDialog
    Dim dir_name As String
    Set file = Application.FileDialog(msoFileDialogFolderPicker)
    With file
        .Title = "Select a directory"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        '@Ignore IndexedDefaultMemberAccess
        dir_name = .SelectedItems(1)
    End With

NextCode:
    GetDirectory = dir_name
    Set file = Nothing
End Function

Public Function GetFiles() As String()
    Dim file As FileDialog
    Dim fileNames() As String
    Dim Index As Long

    Set file = Application.FileDialog(msoFileDialogFilePicker)
    With file
        .Title = "Select one or more files"
        .AllowMultiSelect = True
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        Dim numSelected As Long
        numSelected = .SelectedItems.Count
        ReDim fileNames(numSelected)

        Index = 0
        Dim FileName As Variant
        For Each FileName In .SelectedItems
            fileNames(Index) = FileName
            Index = Index + 1
        Next FileName
    End With

NextCode:
    GetFiles = fileNames
    Set file = Nothing
End Function

Public Function GetFilesInDirectory(ByVal Directory As String, Optional ByVal ext As String) As Variant
    Dim file_path As String
    Dim files(100) As String

    '    If ext Is Nothing Then ext = "*"

    file_path = Dir(Directory & "/*." & ext)
    Dim i As Long
    i = 0
    Do Until file_path = vbNullString
        files(i) = file_path
        file_path = Dir
        i = i + 1
    Loop

    GetFilesInDirectory = files
End Function

'@Description "Add list of filenames to a ListBox (lb)."
Public Sub ListFiles(ByVal lb As ListBox, ByVal fileNames As Object)
    Dim FileName As String
    For Each FileName In fileNames
        lb.AddItem (FileName)
    Next FileName
End Sub

