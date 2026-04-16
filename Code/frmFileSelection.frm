VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFileSelection 
   Caption         =   "Benefits Eligibility Calculator"
   ClientHeight    =   8265.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11520
   OleObjectBlob   =   "frmFileSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFileSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder(FTECalc)
'@IgnoreModule HungarianNotation
Option Explicit

Private fileNames As Variant

Public Function GetFileNames() As Variant
    If Main.DEBUG_ON Then Debug.Print "UserForm_FileSelection.GetFileNames()"
    
    GetFileNames = fileNames
End Function

Public Sub SetFileNames(ByVal FileNameList As Variant)
    If Main.DEBUG_ON Then Debug.Print "UserForm_FileSelection.SetFileNames()"
    fileNames = FileNameList
End Sub

'@Description(Clear text from List Boxes.)
Public Sub ResetListBoxesInForm()
    If Main.DEBUG_ON Then Debug.Print "UserForm_FileSelection.ResetListBoxesInForm()"
    'Reset Files Lists'
    Me.ListBox_FilesList.Clear
    Me.ListBox_FilesAdded.Clear
End Sub

Private Sub CommandButton_Add_Click()
    If Main.DEBUG_ON Then Debug.Print "UserForm_FileSelection.CommandButton_Add_Click()"
    Dim Index As Long
    Dim listCount As Long
    Index = 0
    listCount = Me.ListBox_FilesList.listCount
    
    Do While Index < listCount
        If Me.ListBox_FilesList.Selected(Index) Then
            Me.ListBox_FilesAdded.AddItem (Me.ListBox_FilesList.Column(0, Index))
        End If
        Index = Index + 1
    Loop
        
    Index = listCount - 1
    Do While Index >= 0
        If Me.ListBox_FilesList.Selected(Index) Then
            Me.ListBox_FilesList.RemoveItem (Index)
        End If
        Index = Index - 1
    Loop
End Sub

Private Sub CommandButton_Next_Click()
    If Main.DEBUG_ON Then Debug.Print "UserForm_FileSelection.CommandButton_Next_Click()"
    Me.SetFileNames Me.ListBox_FilesAdded.list

    Main.Main
End Sub

Private Sub CommandButton_Remove_Click()
    If Main.DEBUG_ON Then Debug.Print "UserForm_FileSelection.CommandButton_Remove_Click()"
    Dim Index As Long
    Dim listCount As Long

    Index = 0
    listCount = Me.ListBox_FilesAdded.listCount
    
    Do While Index < Me.ListBox_FilesAdded.listCount
        If Me.ListBox_FilesAdded.Selected(Index) Then
            Me.ListBox_FilesList.AddItem (Me.ListBox_FilesAdded.Column(0, Index))
        End If
        Index = Index + 1
    Loop
        
    Index = listCount - 1
    Do While Index >= 0
        If Me.ListBox_FilesAdded.Selected(Index) Then
            Me.ListBox_FilesAdded.RemoveItem (Index)
        End If
        Index = Index - 1
    Loop
    
End Sub

Private Sub CommandButton_SelectDir_Click()
    If Main.DEBUG_ON Then Debug.Print "UserForm_FileSelection.CommandButton_SelectDir_Click()"
    ResetListBoxesInForm
    'Get and Set Directory'
    Main.Directory = (FileTools.GetDirectory)
    Me.labelSelectedDirectory.Caption = Main.Directory
    
    'Add all files of type xlsx to FilesList'
    Dim fileName As Variant
    For Each fileName In FileTools.GetFilesInDirectory(Main.Directory, "xlsx")
        If fileName <> vbNullString Then
            Me.ListBox_FilesList.AddItem fileName
        End If
    Next fileName
End Sub

Private Sub ControlButton_AddAll_Click()
    If Main.DEBUG_ON Then Debug.Print "UserForm_FileSelection.ControlButton_AddAll_Click()"
    Dim fileName As Variant
    For Each fileName In Me.ListBox_FilesList.list
        If fileName <> vbNullString Then
            Me.ListBox_FilesAdded.AddItem fileName
        End If
    Next fileName
        
    Me.ListBox_FilesList.Clear
End Sub

Private Sub ControlButton_RemoveAll_Click()
    If Main.DEBUG_ON Then Debug.Print "UserForm_FileSelection.ControlButton_RemoveAll_Click()"
    Dim fileName As Variant
    For Each fileName In Me.ListBox_FilesAdded.list
        If fileName <> vbNullString Then
            Me.ListBox_FilesList.AddItem fileName
        End If
    Next fileName
        
    Me.ListBox_FilesAdded.Clear
End Sub

Private Sub UserForm_Initialize()
    If Main.DEBUG_ON Then Debug.Print "UserForm_FileSelection.UserForm_Initialize()"
    Dim fileName As Variant
    If Main.DEBUG_ON Then
        Main.Directory = ".\Data"
        For Each fileName In FileTools.GetFilesInDirectory(Main.Directory, "xlsx")
            If fileName <> vbNullString Then
                Me.ListBox_FilesList.AddItem fileName
            End If
        Next fileName
    End If
End Sub

