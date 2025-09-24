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

'@TestMethod("TODO")
Private Sub TestMethod_SetData()
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    
    'Assert:
    Assert.Inconclusive

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

