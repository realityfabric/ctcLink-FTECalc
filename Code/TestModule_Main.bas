Attribute VB_Name = "TestModule_Main"
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

'@TestMethod("TODO")
Private Sub TestMethod_GetTimestamp()
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

'@TestMethod("TODO")
Private Sub TestMethod_GetTimestampStr()
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

'@TestMethod("TODO")
Private Sub TestMethod_UnixTime()
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

