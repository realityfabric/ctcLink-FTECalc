Attribute VB_Name = "TestModule_Employee"
'@IgnoreModule VariableNotUsed, EmptyMethod
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

Private Assert As Object
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
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Initialize")
Private Sub TestMethod_InitializeClass()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim E As Employee
    
    'Act:
    Set E = New Employee
    
    'Assert:
    Assert.IsTrue E.HoursWorked = 0
    Assert.IsTrue E.DeptID = vbNullString
    Assert.IsTrue E.EmplID = vbNullString
    Assert.IsTrue E.JobCode = vbNullString
    Assert.IsTrue E.Name = vbNullString

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Getter")
Private Sub TestMethod_GetHoursWorked()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim E As Employee
    Dim Hours(25) As Long
    Set E = New Employee
    
    'Act:
    Hours(0) = E.HoursWorked("01A")
    Hours(1) = E.HoursWorked("01B")
    Hours(2) = E.HoursWorked("02A")
    Hours(3) = E.HoursWorked("02B")
    ' TODO: Test 03A-12B and OTH
    
    'Assert:
    Assert.IsTrue E.HoursWorked = 0
    Assert.IsTrue Hours(0) = 0
    Assert.IsTrue Hours(1) = 0
    Assert.IsTrue Hours(2) = 0
    Assert.IsTrue Hours(3) = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Letter")
Private Sub TestMethod_LetHoursWorked()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim E As Employee
    Dim Hours(25) As Long
    Set E = New Employee
    
    'Act:
    E.HoursWorked("01A") = 10
    E.HoursWorked("01B") = 11
    E.HoursWorked("02A") = 20
    E.HoursWorked("02B") = 22
    
    'Assert:
    Assert.IsTrue E.HoursWorked = 63

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub TestMethod_Merge_Defaults_NoHours()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim E1 As Employee
    '@Ignore UseMeaningfulName
    Dim E2 As Employee
    '@Ignore UseMeaningfulName
    Dim E_Merged As Employee
    
    Set E1 = New Employee
    Set E2 = New Employee
    
    E1.EmplID = "111"
    E2.EmplID = "111"
    E1.Name = "John Doe"
    E2.Name = "John Doe"
    E1.DeptID = "DEPT"
    E2.DeptID = "DEPT"
    E1.JobCode = "JOB"
    E2.JobCode = "JOB"
    
    Set E_Merged = E1.Merge(E2)
    
    'Assert:
    Assert.IsTrue E_Merged.EmplID = "111"
    Assert.IsTrue E_Merged.Name = "John Doe"
    Assert.IsTrue E_Merged.DeptID = "DEPT"
    Assert.IsTrue E_Merged.JobCode = "JOB"
    Assert.IsTrue E_Merged.HoursWorked = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub TestMethod_Merge_Defaults_E1Hours_E2NoHours()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim E1 As Employee
    '@Ignore UseMeaningfulName
    Dim E2 As Employee
    '@Ignore UseMeaningfulName
    Dim E_Merged As Employee
    
    Set E1 = New Employee
    Set E2 = New Employee
    
    E1.EmplID = "111"
    E2.EmplID = "111"
    E1.Name = "John Doe"
    E2.Name = "John Doe"
    E1.DeptID = "DEPT"
    E2.DeptID = "DEPT"
    E1.JobCode = "JOB"
    E2.JobCode = "JOB"
    
    E1.HoursWorked("01A") = 10
    
    Set E_Merged = E1.Merge(E2)
    
    'Assert:
    Assert.IsTrue E_Merged.EmplID = "111"
    Assert.IsTrue E_Merged.Name = "John Doe"
    Assert.IsTrue E_Merged.DeptID = "DEPT"
    Assert.IsTrue E_Merged.JobCode = "JOB"
    Assert.IsTrue E_Merged.HoursWorked = 10
    Assert.IsTrue E_Merged.HoursWorked("01A") = 10

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub TestMethod_Merge_Defaults_E1Hours_E2Hours()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim E1 As Employee
    '@Ignore UseMeaningfulName
    Dim E2 As Employee
    '@Ignore UseMeaningfulName
    Dim E_Merged As Employee
    
    Set E1 = New Employee
    Set E2 = New Employee
    
    E1.EmplID = "111"
    E2.EmplID = "111"
    E1.Name = "John Doe"
    E2.Name = "John Doe"
    E1.DeptID = "DEPT"
    E2.DeptID = "DEPT"
    E1.JobCode = "JOB"
    E2.JobCode = "JOB"
    
    E1.HoursWorked("01A") = 10
    E2.HoursWorked("12B") = 11
    
    Set E_Merged = E1.Merge(E2)
    
    'Assert:
    Assert.IsTrue E_Merged.EmplID = "111"
    Assert.IsTrue E_Merged.Name = "John Doe"
    Assert.IsTrue E_Merged.DeptID = "DEPT"
    Assert.IsTrue E_Merged.JobCode = "JOB"
    Assert.IsTrue E_Merged.HoursWorked = 21
    Assert.IsTrue E_Merged.HoursWorked("01A") = 10
    Assert.IsTrue E_Merged.HoursWorked("12B") = 11

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub TestMethod_Merge_Defaults_EmplIDMismatch()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim E1 As Employee
    '@Ignore UseMeaningfulName
    Dim E2 As Employee
    '@Ignore UseMeaningfulName
    Dim E_Merged As Employee
    
    Set E1 = New Employee
    Set E2 = New Employee
    
    E1.EmplID = "111"
    E2.EmplID = "222"
    E1.Name = "John Doe"
    E2.Name = "John Doe"
    E1.DeptID = "DEPT"
    E2.DeptID = "DEPT"
    E1.JobCode = "JOB"
    E2.JobCode = "JOB"
    
    Set E_Merged = E1.Merge(E2)
    
    'Assert:
    Assert.IsNothing E_Merged

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub TestMethod_Merge_Defaults_DeptIDMismatch()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim E1 As Employee
    '@Ignore UseMeaningfulName
    Dim E2 As Employee
    '@Ignore UseMeaningfulName
    Dim E_Merged As Employee
    
    Set E1 = New Employee
    Set E2 = New Employee
    
    E1.EmplID = "111"
    E2.EmplID = "111"
    E1.Name = "John Doe"
    E2.Name = "John Doe"
    E1.DeptID = "DEPT1"
    E2.DeptID = "DEPT2"
    E1.JobCode = "JOB"
    E2.JobCode = "JOB"
    
    Set E_Merged = E1.Merge(E2)
    
    'Assert:
    Assert.IsNothing E_Merged

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub TestMethod_Merge_Defaults_JobCodeMismatch()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim E1 As Employee
    '@Ignore UseMeaningfulName
    Dim E2 As Employee
    '@Ignore UseMeaningfulName
    Dim E_Merged As Employee
    
    Set E1 = New Employee
    Set E2 = New Employee
    
    E1.EmplID = "111"
    E2.EmplID = "111"
    E1.Name = "John Doe"
    E2.Name = "John Doe"
    E1.DeptID = "DEPT"
    E2.DeptID = "DEPT"
    E1.JobCode = "JOB1"
    E2.JobCode = "JOB2"
    
    Set E_Merged = E1.Merge(E2)
    
    'Assert:
    Assert.IsNothing E_Merged

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub TestMethod_Merge_NoPreserveDeptID_DeptIDMatch()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim E1 As Employee
    '@Ignore UseMeaningfulName
    Dim E2 As Employee
    '@Ignore UseMeaningfulName
    Dim E_Merged As Employee
    
    Set E1 = New Employee
    Set E2 = New Employee
    
    E1.EmplID = "111"
    E2.EmplID = "111"
    E1.Name = "John Doe"
    E2.Name = "John Doe"
    E1.DeptID = "DEPT"
    E2.DeptID = "DEPT"
    E1.JobCode = "JOB"
    E2.JobCode = "JOB"
    
    Set E_Merged = E1.Merge(E2, PreserveDeptID:=False)
    
    'Assert:
    Assert.IsTrue E_Merged.EmplID = "111"
    Assert.IsTrue E_Merged.Name = "John Doe"
    Assert.IsTrue E_Merged.DeptID = "*"
    Assert.IsTrue E_Merged.JobCode = "JOB"
    Assert.IsTrue E_Merged.HoursWorked = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub TestMethod_Merge_NoPreserveDeptID_DeptIDMismatch()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim E1 As Employee
    '@Ignore UseMeaningfulName
    Dim E2 As Employee
    '@Ignore UseMeaningfulName
    Dim E_Merged As Employee
    
    Set E1 = New Employee
    Set E2 = New Employee
    
    E1.EmplID = "111"
    E2.EmplID = "111"
    E1.Name = "John Doe"
    E2.Name = "John Doe"
    E1.DeptID = "DEPT1"
    E2.DeptID = "DEPT2"
    E1.JobCode = "JOB"
    E2.JobCode = "JOB"
    
    Set E_Merged = E1.Merge(E2, PreserveDeptID:=False)
    
    'Assert:
    Assert.IsTrue E_Merged.EmplID = "111"
    Assert.IsTrue E_Merged.Name = "John Doe"
    Assert.IsTrue E_Merged.DeptID = "*"
    Assert.IsTrue E_Merged.JobCode = "JOB"
    Assert.IsTrue E_Merged.HoursWorked = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub TestMethod_Merge_NoPreserveJobCode_JobCodeMatch()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim E1 As Employee
    '@Ignore UseMeaningfulName
    Dim E2 As Employee
    '@Ignore UseMeaningfulName
    Dim E_Merged As Employee
    
    Set E1 = New Employee
    Set E2 = New Employee
    
    E1.EmplID = "111"
    E2.EmplID = "111"
    E1.Name = "John Doe"
    E2.Name = "John Doe"
    E1.DeptID = "DEPT"
    E2.DeptID = "DEPT"
    E1.JobCode = "JOB"
    E2.JobCode = "JOB"
    
    Set E_Merged = E1.Merge(E2, PreserveJobCode:=False)
    
    'Assert:
    Assert.IsTrue E_Merged.EmplID = "111"
    Assert.IsTrue E_Merged.Name = "John Doe"
    Assert.IsTrue E_Merged.DeptID = "DEPT"
    Assert.IsTrue E_Merged.JobCode = "*"
    Assert.IsTrue E_Merged.HoursWorked = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub TestMethod_Merge_NoPreserveJobCode_JobCodeMismatch()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim E1 As Employee
    '@Ignore UseMeaningfulName
    Dim E2 As Employee
    '@Ignore UseMeaningfulName
    Dim E_Merged As Employee
    
    Set E1 = New Employee
    Set E2 = New Employee
    
    E1.EmplID = "111"
    E2.EmplID = "111"
    E1.Name = "John Doe"
    E2.Name = "John Doe"
    E1.DeptID = "DEPT"
    E2.DeptID = "DEPT"
    E1.JobCode = "JOB1"
    E2.JobCode = "JOB2"
    
    Set E_Merged = E1.Merge(E2, PreserveJobCode:=False)
    
    'Assert:
    Assert.IsTrue E_Merged.EmplID = "111"
    Assert.IsTrue E_Merged.Name = "John Doe"
    Assert.IsTrue E_Merged.DeptID = "DEPT"
    Assert.IsTrue E_Merged.JobCode = "*"
    Assert.IsTrue E_Merged.HoursWorked = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Getter")
Private Sub TestMethod_GetDeptID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim E As Employee
    Dim id As String
    Set E = New Employee
    E.DeptID = "12345"
    
    'Act:
    id = E.DeptID
    
    'Assert:
    Assert.IsTrue id = "12345"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Getter")
Private Sub TestMethod_GetEmplID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim E As Employee
    Dim id As String
    Set E = New Employee
    E.EmplID = "12345"
    
    'Act:
    
    id = E.EmplID
    
    'Assert:
    Assert.IsTrue id = "12345"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Getter")
Private Sub TestMethod_GetJobCode()
    On Error GoTo TestFail
    
    'Arrange:
    Dim E As Employee
    Dim jc As String
    Set E = New Employee
    E.JobCode = "ABC"
    
    'Act:
    
    jc = E.JobCode
    
    'Assert:
    Assert.jc = "ABC"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Getter")
Private Sub TestMethod_GetName()
    On Error GoTo TestFail
    
    'Arrange:
    Dim E As Employee
    Dim n As String
    Set E = New Employee
    E.Name = "Harry Haywood"
    
    'Act:
    
    n = E.Name
    
    'Assert:
    Assert.IsTrue n = "Harry Haywood"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Letter")
Private Sub TestMethod_LetDeptID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim E As Employee
    Set E = New Employee
    
    'Act:
    E.DeptID = "12345"
    
    'Assert:
    Assert.IsTrue E.DeptID = "12345"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TODO")
Private Sub TestMethod_LetEmplID()
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
Private Sub TestMethod_LetJobCode()
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
Private Sub TestMethod_LetName()
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

'@TestMethod("Calculation")
Private Sub TestMethod_FTE_OTHeq10()
    On Error GoTo TestFail
    
    'Arrange:
    Dim E As Employee
    Set E = New Employee
    Dim fte As Single
    
    E.HoursWorked("OTH") = 10
    
    'Act:
    fte = E.fte()
    
    'Assert:
    Assert.IsTrue fte = 5.05

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Calculation")
Private Sub TestMethod_FTE_NoHours()
    On Error GoTo TestFail
    
    'Arrange:
    Dim E As Employee
    Set E = New Employee
    Dim fte As Single
    
    'Act:
    fte = E.fte()
    
    'Assert:
    Assert.IsTrue fte = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TODO")
Private Sub TestMethod_InitializeHoursWorked()
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

