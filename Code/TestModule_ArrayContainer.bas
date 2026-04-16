Attribute VB_Name = "TestModule_ArrayContainer"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestMethod("Initialize")
Private Sub TestMethod_NewArrayContainer()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed, UseMeaningfulName
    Dim AC As ArrayContainer
    
    'Act:
    '@Ignore AssignmentNotUsed
    Set AC = New ArrayContainer
    
    'Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Setter")
Private Sub TestMethod_SetData()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim AC As ArrayContainer
    '@Ignore UseMeaningfulName
    Dim DA(2, 3) As Variant
    Dim ColumnIndex As Long
    Dim RowIndex As Long
    
    ColumnIndex = 3
    RowIndex = 2
    
    Set AC = New ArrayContainer
    
    DA(0, 0) = "A"
    DA(0, 1) = "B"
    DA(0, 2) = "C"
    DA(0, 3) = "D"
    DA(1, 0) = "E"
    DA(1, 1) = "F"
    DA(1, 2) = "G"
    DA(1, 3) = "H"
    DA(2, 0) = "I"
    DA(2, 1) = "J"
    DA(2, 2) = "K"
    DA(2, 3) = "L"
    
    'Act:
    AC.SetData RowIndex, ColumnIndex, DA
    
    'Assert:
    Assert.IsTrue AC.Rows = 2
    Assert.IsTrue AC.columns = 3
    Assert.SequenceEquals AC.Data(), DA
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

