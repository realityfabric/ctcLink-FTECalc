Attribute VB_Name = "TestModule_Employee"
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

'@TestMethod("Initialization")
Private Sub Test_EmployeeNew()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim Emp As Employee
    
    'Act:
    '@Ignore AssignmentNotUsed
    Set Emp = New Employee
    
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

'@TestMethod("Letters")
Private Sub Test_EmployeeLetName()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Set Emp = New Employee
    
    'Act:
    Emp.Name = "John Doe"
    
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

'@TestMethod("Getters")
Private Sub Test_EmployeeGetName()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Dim Name As String
    Set Emp = New Employee
    Emp.Name = "John Doe"
    
    'Act:
    Name = Emp.Name
    
    'Assert:
    Assert.IsTrue Name = "John Doe"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Letters")
Private Sub Test_EmployeeLetDepartment()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Set Emp = New Employee
    
    'Act:
    Emp.Department = "00000"
    
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

'@TestMethod("Letters")
Private Sub Test_EmployeeLetDepartmentNonNumeric()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Set Emp = New Employee
    
    'Act:
    Emp.Department = "ThisDepartmentIsNonNumeric"
    
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

'@TestMethod("Getters")
Private Sub Test_EmployeeGetDepartment()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Dim Department As String
    Set Emp = New Employee
    Emp.Department = "00000"
    
    'Act:
    Department = Emp.Department
    
    'Assert:
    Assert.IsTrue Department = "00000"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Getters")
Private Sub Test_EmployeeGetDepartmentNonNumeric()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Dim Department As String
    Set Emp = New Employee
    Emp.Department = "ThisDepartmentIsNonNumeric"
    
    'Act:
    Department = Emp.Department
    
    'Assert:
    Assert.IsTrue Department = "ThisDepartmentIsNonNumeric"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Getters")
Private Sub Test_EmployeeGetDepartmentNonNumericWithSpaces()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Set Emp = New Employee
    Emp.Department = "This Department Is Non Numeric"
    
    'Act:
    
    ' No actions to take
    
    'Assert:
    Assert.IsTrue Emp.Department = "This Department Is Non Numeric"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Letters")
Private Sub Test_EmployeeLetEmplID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Set Emp = New Employee
    
    'Act:
    Emp.EmplID = "000000000"
    
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

'@TestMethod("Getters")
Private Sub Test_EmployeeGetEmplID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Dim EmplID As String
    Set Emp = New Employee
    Emp.EmplID = "000000000"
    
    'Act:
    EmplID = Emp.EmplID
    
    'Assert:
    Assert.IsTrue EmplID = "000000000"
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Initialization")
Private Sub Test_EmployeeInitialize()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    
    'Act:
    Set Emp = New Employee
    
    'Assert:
    Assert.IsTrue Emp.hoursWorked("01A") = 0
    Assert.IsTrue Emp.hoursWorked("01B") = 0
    
    Assert.IsTrue Emp.hoursWorked("02A") = 0
    Assert.IsTrue Emp.hoursWorked("02B") = 0
    
    Assert.IsTrue Emp.hoursWorked("03A") = 0
    Assert.IsTrue Emp.hoursWorked("03B") = 0
    
    Assert.IsTrue Emp.hoursWorked("04A") = 0
    Assert.IsTrue Emp.hoursWorked("04B") = 0
    
    Assert.IsTrue Emp.hoursWorked("05A") = 0
    Assert.IsTrue Emp.hoursWorked("05B") = 0
    
    Assert.IsTrue Emp.hoursWorked("06A") = 0
    Assert.IsTrue Emp.hoursWorked("06B") = 0
    
    Assert.IsTrue Emp.hoursWorked("07A") = 0
    Assert.IsTrue Emp.hoursWorked("07B") = 0
    
    Assert.IsTrue Emp.hoursWorked("08A") = 0
    Assert.IsTrue Emp.hoursWorked("08B") = 0
    
    Assert.IsTrue Emp.hoursWorked("09A") = 0
    Assert.IsTrue Emp.hoursWorked("09B") = 0
    
    Assert.IsTrue Emp.hoursWorked("10A") = 0
    Assert.IsTrue Emp.hoursWorked("10B") = 0
    
    Assert.IsTrue Emp.hoursWorked("11A") = 0
    Assert.IsTrue Emp.hoursWorked("11B") = 0
    
    Assert.IsTrue Emp.hoursWorked("12A") = 0
    Assert.IsTrue Emp.hoursWorked("12B") = 0
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Getters")
Private Sub Test_EmployeeGetIsHourly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim HourlyEmployee As Employee
    Dim AppointedEmployee As Employee
    
    Dim HourlyEmployeeIsHourly As Boolean
    Dim AppointedEmployeeIsHourly As Boolean
    
    Set HourlyEmployee = New Employee
    Set AppointedEmployee = New Employee
    
    HourlyEmployee.IsHourly = True
    AppointedEmployee.IsHourly = False
    
    'Act:
    HourlyEmployeeIsHourly = HourlyEmployee.IsHourly
    AppointedEmployeeIsHourly = AppointedEmployee.IsHourly
    
    'Assert:
    Assert.IsTrue HourlyEmployeeIsHourly
    Assert.IsFalse AppointedEmployeeIsHourly
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Getters")
Private Sub Test_EmployeeLetIsHourly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Set Emp = New Employee
    
    'Act:
    ' Check to see if IsHourly can be set to true without error.
    Emp.IsHourly = True
    
    ' Check to see if IsHourly can be set to false without error.
    Emp.IsHourly = False
    
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

'TODO Test Let JobCode
'TODO Test Get JobCode
'TODO Test Let HoursWorked
'TODO Test Get HoursWorked
'TODO Test Key
'TODO zvKeys
'TODO zvMergeEmployeeHours
'TODO Test Print_Employee
'TODO Test Get Source
'TODO Test Let Source

'@TestMethod("Initialization")
Private Sub Test_Copy_EmplID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim OriginalEmp As Employee
    Dim CopyEmp As Employee
    
    Set OriginalEmp = New Employee
    OriginalEmp.EmplID = "111111111"
    OriginalEmp.Name = "Protagonist,Hero"
    OriginalEmp.Department = "PIZZA"
    OriginalEmp.JobCode = "SCT"
    OriginalEmp.Source = "https://en.wikipedia.org/"
    OriginalEmp.hoursWorked("01A") = 1
    OriginalEmp.hoursWorked("01B") = 2
    OriginalEmp.hoursWorked("02A") = 3
    OriginalEmp.hoursWorked("02B") = 4
    OriginalEmp.hoursWorked("03A") = 5
    OriginalEmp.hoursWorked("03B") = 6
    OriginalEmp.hoursWorked("04A") = 7
    OriginalEmp.hoursWorked("04B") = 8
    OriginalEmp.hoursWorked("05A") = 9
    OriginalEmp.hoursWorked("05B") = 10
    OriginalEmp.hoursWorked("06A") = 11
    OriginalEmp.hoursWorked("06B") = 12
    OriginalEmp.hoursWorked("07A") = 13
    OriginalEmp.hoursWorked("07B") = 14
    OriginalEmp.hoursWorked("08A") = 15
    OriginalEmp.hoursWorked("08B") = 16
    OriginalEmp.hoursWorked("09A") = 17
    OriginalEmp.hoursWorked("09B") = 18
    OriginalEmp.hoursWorked("10A") = 19
    OriginalEmp.hoursWorked("10B") = 20
    OriginalEmp.hoursWorked("11A") = 21
    OriginalEmp.hoursWorked("11B") = 22
    OriginalEmp.hoursWorked("12A") = 23
    OriginalEmp.hoursWorked("12B") = 24
    
    OriginalEmp.IsHourly = True
    
    'Act:
    
    Set CopyEmp = OriginalEmp.Copy
    
    'Assert:
    Assert.IsTrue CopyEmp.EmplID = "111111111"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Initialization")
Private Sub Test_Copy_Name()
    On Error GoTo TestFail
    
    'Arrange:
    Dim OriginalEmp As Employee
    Dim CopyEmp As Employee
    
    Set OriginalEmp = New Employee
    OriginalEmp.EmplID = "111111111"
    OriginalEmp.Name = "Protagonist,Hero"
    OriginalEmp.Department = "PIZZA"
    OriginalEmp.JobCode = "SCT"
    OriginalEmp.Source = "https://en.wikipedia.org/"
    OriginalEmp.hoursWorked("01A") = 1
    OriginalEmp.hoursWorked("01B") = 2
    OriginalEmp.hoursWorked("02A") = 3
    OriginalEmp.hoursWorked("02B") = 4
    OriginalEmp.hoursWorked("03A") = 5
    OriginalEmp.hoursWorked("03B") = 6
    OriginalEmp.hoursWorked("04A") = 7
    OriginalEmp.hoursWorked("04B") = 8
    OriginalEmp.hoursWorked("05A") = 9
    OriginalEmp.hoursWorked("05B") = 10
    OriginalEmp.hoursWorked("06A") = 11
    OriginalEmp.hoursWorked("06B") = 12
    OriginalEmp.hoursWorked("07A") = 13
    OriginalEmp.hoursWorked("07B") = 14
    OriginalEmp.hoursWorked("08A") = 15
    OriginalEmp.hoursWorked("08B") = 16
    OriginalEmp.hoursWorked("09A") = 17
    OriginalEmp.hoursWorked("09B") = 18
    OriginalEmp.hoursWorked("10A") = 19
    OriginalEmp.hoursWorked("10B") = 20
    OriginalEmp.hoursWorked("11A") = 21
    OriginalEmp.hoursWorked("11B") = 22
    OriginalEmp.hoursWorked("12A") = 23
    OriginalEmp.hoursWorked("12B") = 24
    
    OriginalEmp.IsHourly = True
    
    'Act:
    
    Set CopyEmp = OriginalEmp.Copy
    
    'Assert:
    Assert.IsTrue CopyEmp.Name = "Protagonist,Hero"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Initialization")
Private Sub Test_Copy_Department()
    On Error GoTo TestFail
    
    'Arrange:
    Dim OriginalEmp As Employee
    Dim CopyEmp As Employee
    
    Set OriginalEmp = New Employee
    OriginalEmp.EmplID = "111111111"
    OriginalEmp.Name = "Protagonist,Hero"
    OriginalEmp.Department = "PIZZA"
    OriginalEmp.JobCode = "SCT"
    OriginalEmp.Source = "https://en.wikipedia.org/"
    OriginalEmp.hoursWorked("01A") = 1
    OriginalEmp.hoursWorked("01B") = 2
    OriginalEmp.hoursWorked("02A") = 3
    OriginalEmp.hoursWorked("02B") = 4
    OriginalEmp.hoursWorked("03A") = 5
    OriginalEmp.hoursWorked("03B") = 6
    OriginalEmp.hoursWorked("04A") = 7
    OriginalEmp.hoursWorked("04B") = 8
    OriginalEmp.hoursWorked("05A") = 9
    OriginalEmp.hoursWorked("05B") = 10
    OriginalEmp.hoursWorked("06A") = 11
    OriginalEmp.hoursWorked("06B") = 12
    OriginalEmp.hoursWorked("07A") = 13
    OriginalEmp.hoursWorked("07B") = 14
    OriginalEmp.hoursWorked("08A") = 15
    OriginalEmp.hoursWorked("08B") = 16
    OriginalEmp.hoursWorked("09A") = 17
    OriginalEmp.hoursWorked("09B") = 18
    OriginalEmp.hoursWorked("10A") = 19
    OriginalEmp.hoursWorked("10B") = 20
    OriginalEmp.hoursWorked("11A") = 21
    OriginalEmp.hoursWorked("11B") = 22
    OriginalEmp.hoursWorked("12A") = 23
    OriginalEmp.hoursWorked("12B") = 24
    
    OriginalEmp.IsHourly = True
    
    'Act:
    
    Set CopyEmp = OriginalEmp.Copy
    
    'Assert:
    Assert.IsTrue CopyEmp.Department = "PIZZA"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Initialization")
Private Sub Test_Copy_JobCode()
    On Error GoTo TestFail
    
    'Arrange:
    Dim OriginalEmp As Employee
    Dim CopyEmp As Employee
    
    Set OriginalEmp = New Employee
    OriginalEmp.EmplID = "111111111"
    OriginalEmp.Name = "Protagonist,Hero"
    OriginalEmp.Department = "PIZZA"
    OriginalEmp.JobCode = "SCT"
    OriginalEmp.Source = "https://en.wikipedia.org/"
    OriginalEmp.hoursWorked("01A") = 1
    OriginalEmp.hoursWorked("01B") = 2
    OriginalEmp.hoursWorked("02A") = 3
    OriginalEmp.hoursWorked("02B") = 4
    OriginalEmp.hoursWorked("03A") = 5
    OriginalEmp.hoursWorked("03B") = 6
    OriginalEmp.hoursWorked("04A") = 7
    OriginalEmp.hoursWorked("04B") = 8
    OriginalEmp.hoursWorked("05A") = 9
    OriginalEmp.hoursWorked("05B") = 10
    OriginalEmp.hoursWorked("06A") = 11
    OriginalEmp.hoursWorked("06B") = 12
    OriginalEmp.hoursWorked("07A") = 13
    OriginalEmp.hoursWorked("07B") = 14
    OriginalEmp.hoursWorked("08A") = 15
    OriginalEmp.hoursWorked("08B") = 16
    OriginalEmp.hoursWorked("09A") = 17
    OriginalEmp.hoursWorked("09B") = 18
    OriginalEmp.hoursWorked("10A") = 19
    OriginalEmp.hoursWorked("10B") = 20
    OriginalEmp.hoursWorked("11A") = 21
    OriginalEmp.hoursWorked("11B") = 22
    OriginalEmp.hoursWorked("12A") = 23
    OriginalEmp.hoursWorked("12B") = 24
    
    OriginalEmp.IsHourly = True
    
    'Act:
    
    Set CopyEmp = OriginalEmp.Copy
    
    'Assert:
    Assert.IsTrue CopyEmp.JobCode = "SCT"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Initialization")
Private Sub Test_Copy_IsHourly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim OriginalEmp As Employee
    Dim CopyEmp As Employee
    
    Set OriginalEmp = New Employee
    OriginalEmp.EmplID = "111111111"
    OriginalEmp.Name = "Protagonist,Hero"
    OriginalEmp.Department = "PIZZA"
    OriginalEmp.JobCode = "SCT"
    OriginalEmp.Source = "https://en.wikipedia.org/"
    OriginalEmp.hoursWorked("01A") = 1
    OriginalEmp.hoursWorked("01B") = 2
    OriginalEmp.hoursWorked("02A") = 3
    OriginalEmp.hoursWorked("02B") = 4
    OriginalEmp.hoursWorked("03A") = 5
    OriginalEmp.hoursWorked("03B") = 6
    OriginalEmp.hoursWorked("04A") = 7
    OriginalEmp.hoursWorked("04B") = 8
    OriginalEmp.hoursWorked("05A") = 9
    OriginalEmp.hoursWorked("05B") = 10
    OriginalEmp.hoursWorked("06A") = 11
    OriginalEmp.hoursWorked("06B") = 12
    OriginalEmp.hoursWorked("07A") = 13
    OriginalEmp.hoursWorked("07B") = 14
    OriginalEmp.hoursWorked("08A") = 15
    OriginalEmp.hoursWorked("08B") = 16
    OriginalEmp.hoursWorked("09A") = 17
    OriginalEmp.hoursWorked("09B") = 18
    OriginalEmp.hoursWorked("10A") = 19
    OriginalEmp.hoursWorked("10B") = 20
    OriginalEmp.hoursWorked("11A") = 21
    OriginalEmp.hoursWorked("11B") = 22
    OriginalEmp.hoursWorked("12A") = 23
    OriginalEmp.hoursWorked("12B") = 24
    
    OriginalEmp.IsHourly = True
    
    'Act:
    
    Set CopyEmp = OriginalEmp.Copy
    
    'Assert:
    Assert.IsTrue CopyEmp.IsHourly = True

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Initialization")
Private Sub Test_Copy_Source()
    On Error GoTo TestFail
    
    'Arrange:
    Dim OriginalEmp As Employee
    Dim CopyEmp As Employee
    
    Set OriginalEmp = New Employee
    OriginalEmp.EmplID = "111111111"
    OriginalEmp.Name = "Protagonist,Hero"
    OriginalEmp.Department = "PIZZA"
    OriginalEmp.JobCode = "SCT"
    OriginalEmp.Source = "https://en.wikipedia.org/"
    OriginalEmp.hoursWorked("01A") = 1
    OriginalEmp.hoursWorked("01B") = 2
    OriginalEmp.hoursWorked("02A") = 3
    OriginalEmp.hoursWorked("02B") = 4
    OriginalEmp.hoursWorked("03A") = 5
    OriginalEmp.hoursWorked("03B") = 6
    OriginalEmp.hoursWorked("04A") = 7
    OriginalEmp.hoursWorked("04B") = 8
    OriginalEmp.hoursWorked("05A") = 9
    OriginalEmp.hoursWorked("05B") = 10
    OriginalEmp.hoursWorked("06A") = 11
    OriginalEmp.hoursWorked("06B") = 12
    OriginalEmp.hoursWorked("07A") = 13
    OriginalEmp.hoursWorked("07B") = 14
    OriginalEmp.hoursWorked("08A") = 15
    OriginalEmp.hoursWorked("08B") = 16
    OriginalEmp.hoursWorked("09A") = 17
    OriginalEmp.hoursWorked("09B") = 18
    OriginalEmp.hoursWorked("10A") = 19
    OriginalEmp.hoursWorked("10B") = 20
    OriginalEmp.hoursWorked("11A") = 21
    OriginalEmp.hoursWorked("11B") = 22
    OriginalEmp.hoursWorked("12A") = 23
    OriginalEmp.hoursWorked("12B") = 24
    
    OriginalEmp.IsHourly = True
    
    'Act:
    
    Set CopyEmp = OriginalEmp.Copy
    
    'Assert:
    Assert.IsTrue CopyEmp.Source = "https://en.wikipedia.org/"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Initialization")
Private Sub Test_Copy_HoursWorked()
    On Error GoTo TestFail
    
    'Arrange:
    Dim OriginalEmp As Employee
    Dim CopyEmp As Employee
    
    Set OriginalEmp = New Employee
    OriginalEmp.EmplID = "111111111"
    OriginalEmp.Name = "Protagonist,Hero"
    OriginalEmp.Department = "PIZZA"
    OriginalEmp.JobCode = "SCT"
    OriginalEmp.Source = "https://en.wikipedia.com/"
    OriginalEmp.hoursWorked("01A") = 1
    OriginalEmp.hoursWorked("01B") = 2
    OriginalEmp.hoursWorked("02A") = 3
    OriginalEmp.hoursWorked("02B") = 4
    OriginalEmp.hoursWorked("03A") = 5
    OriginalEmp.hoursWorked("03B") = 6
    OriginalEmp.hoursWorked("04A") = 7
    OriginalEmp.hoursWorked("04B") = 8
    OriginalEmp.hoursWorked("05A") = 9
    OriginalEmp.hoursWorked("05B") = 10
    OriginalEmp.hoursWorked("06A") = 11
    OriginalEmp.hoursWorked("06B") = 12
    OriginalEmp.hoursWorked("07A") = 13
    OriginalEmp.hoursWorked("07B") = 14
    OriginalEmp.hoursWorked("08A") = 15
    OriginalEmp.hoursWorked("08B") = 16
    OriginalEmp.hoursWorked("09A") = 17
    OriginalEmp.hoursWorked("09B") = 18
    OriginalEmp.hoursWorked("10A") = 19
    OriginalEmp.hoursWorked("10B") = 20
    OriginalEmp.hoursWorked("11A") = 21
    OriginalEmp.hoursWorked("11B") = 22
    OriginalEmp.hoursWorked("12A") = 23
    OriginalEmp.hoursWorked("12B") = 24
    
    OriginalEmp.IsHourly = True
    
    'Act:
    
    Set CopyEmp = OriginalEmp.Copy
    
    'Assert:
    Assert.IsTrue CopyEmp.hoursWorked("01A") = 1
    Assert.IsTrue CopyEmp.hoursWorked("01B") = 2
    Assert.IsTrue CopyEmp.hoursWorked("02A") = 3
    Assert.IsTrue CopyEmp.hoursWorked("02B") = 4
    Assert.IsTrue CopyEmp.hoursWorked("03A") = 5
    Assert.IsTrue CopyEmp.hoursWorked("03B") = 6
    Assert.IsTrue CopyEmp.hoursWorked("04A") = 7
    Assert.IsTrue CopyEmp.hoursWorked("04B") = 8
    Assert.IsTrue CopyEmp.hoursWorked("05A") = 9
    Assert.IsTrue CopyEmp.hoursWorked("05B") = 10
    Assert.IsTrue CopyEmp.hoursWorked("06A") = 11
    Assert.IsTrue CopyEmp.hoursWorked("06B") = 12
    Assert.IsTrue CopyEmp.hoursWorked("07A") = 13
    Assert.IsTrue CopyEmp.hoursWorked("07B") = 14
    Assert.IsTrue CopyEmp.hoursWorked("08A") = 15
    Assert.IsTrue CopyEmp.hoursWorked("08B") = 16
    Assert.IsTrue CopyEmp.hoursWorked("09A") = 17
    Assert.IsTrue CopyEmp.hoursWorked("09B") = 18
    Assert.IsTrue CopyEmp.hoursWorked("10A") = 19
    Assert.IsTrue CopyEmp.hoursWorked("10B") = 20
    Assert.IsTrue CopyEmp.hoursWorked("11A") = 21
    Assert.IsTrue CopyEmp.hoursWorked("11B") = 22
    Assert.IsTrue CopyEmp.hoursWorked("12A") = 23
    Assert.IsTrue CopyEmp.hoursWorked("12B") = 24

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

