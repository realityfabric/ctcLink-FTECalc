Attribute VB_Name = "TestModule_ArrayContainer"
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

Private Assert As Object
'@Ignore VariableNotUsed
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
'@Ignore EmptyMethod
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
'@Ignore EmptyMethod
Private Sub TestCleanup()
    'this method runs after every test in the module.
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
    Dim AC As ArrayContainer
    Dim DA(2, 3) As Variant
    Dim c As Long
    Dim r As Long
    
    c = 3
    r = 2
    
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
    AC.SetData r, c, DA
    
    
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

