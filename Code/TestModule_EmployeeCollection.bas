Attribute VB_Name = "TestModule_EmployeeCollection"
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

'@TestMethod("Uncategorized")
Private Sub Test_Add()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim Emp As Employee
    Set Employees = New EmployeeCollection
    Set Emp = New Employee
    
    'Act:
    Employees.Add Emp
    
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

'@TestMethod("Uncategorized")
Private Sub Test_Remove()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim Emp As Employee
    Set Employees = New EmployeeCollection
    Set Emp = New Employee
    Employees.Add Emp
    
    'Act:
    Employees.Remove (1)
    
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

'@TestMethod("Uncategorized")
Private Sub Test_Count_Empty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Set Employees = New EmployeeCollection
    
    'Act:
    ' No Actions to take
    
    'Assert:
    Assert.IsTrue Employees.Count() = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_Count_One()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim Emp As Employee
    Set Employees = New EmployeeCollection
    Set Emp = New Employee
    
    'Act:
    Employees.Add Emp
    
    'Assert:
    Assert.IsTrue Employees.Count() = 1

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_Add_Remove_Count_Empty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim Emp As Employee
    Set Employees = New EmployeeCollection
    Set Emp = New Employee
    
    'Act:
    Employees.Add Emp
    Employees.Remove 1
    
    'Assert:
    Assert.IsTrue Employees.Count() = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_Concat()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees_A As EmployeeCollection
    Dim Employees_B As EmployeeCollection
    Dim Emp_A As Employee
    Dim Emp_B As Employee
    Set Employees_A = New EmployeeCollection
    Set Employees_B = New EmployeeCollection
    Set Emp_A = New Employee
    Set Emp_B = New Employee
    Employees_A.Add Emp_A
    Employees_B.Add Emp_B
    
    'Act:
    Employees_A.Concat Employees_B
    
    'Assert:
    Assert.IsTrue Employees_A.Count() = 2

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_Filter_EmplID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim Filtered_EC As EmployeeCollection
    Dim Emps(5) As Employee
    Dim Index As Long
    
    Set Employees = New EmployeeCollection
    For Index = 0 To 4
        Set Emps(Index) = New Employee
    Next Index
    
    Emps(0).EmplID = "FILTER"
    Emps(0).Name = "Emp 0"
    Emps(0).DeptID = "0"
    Emps(0).JobCode = "0"
    
    Emps(1).EmplID = "FILTER"
    Emps(1).Name = "Emp 1"
    Emps(1).DeptID = "1"
    Emps(1).JobCode = "1"
    
    Emps(2).EmplID = "2"
    Emps(2).Name = "Emp 2"
    Emps(2).DeptID = "2"
    Emps(2).JobCode = "2"
    
    Emps(3).EmplID = "3"
    Emps(3).Name = "Emp 3"
    Emps(3).DeptID = "3"
    Emps(3).JobCode = "3"
    
    Emps(4).EmplID = "4"
    Emps(4).Name = "Emp 4"
    Emps(4).DeptID = "4"
    Emps(4).JobCode = "4"
    
    For Index = 0 To 4
        Employees.Add Emps(Index)
    Next Index
    
    'Act:
    Set Filtered_EC = Employees.Filter(EmplIDFilter:="FILTER")
    
    'Assert:
    Assert.IsTrue Filtered_EC.Item(1).EmplID = "FILTER"
    Assert.IsTrue Filtered_EC.Item(1).Name = "Emp 0"
    Assert.IsTrue Filtered_EC.Item(2).EmplID = "FILTER"
    Assert.IsTrue Filtered_EC.Item(2).Name = "Emp 1"
    Assert.IsTrue Filtered_EC.Count = 2

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_Filter_DeptID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim Filtered_EC As EmployeeCollection
    Dim Emps(5) As Employee
    Dim Index As Long
    
    Set Employees = New EmployeeCollection
    For Index = 0 To 4
        Set Emps(Index) = New Employee
    Next Index
    
    Emps(0).EmplID = "0"
    Emps(0).Name = "Emp 0"
    Emps(0).DeptID = "FILTER"
    Emps(0).JobCode = "0"
    
    Emps(1).EmplID = "1"
    Emps(1).Name = "Emp 1"
    Emps(1).DeptID = "FILTER"
    Emps(1).JobCode = "1"
    
    Emps(2).EmplID = "2"
    Emps(2).Name = "Emp 2"
    Emps(2).DeptID = "2"
    Emps(2).JobCode = "2"
    
    Emps(3).EmplID = "3"
    Emps(3).Name = "Emp 3"
    Emps(3).DeptID = "3"
    Emps(3).JobCode = "3"
    
    Emps(4).EmplID = "4"
    Emps(4).Name = "Emp 4"
    Emps(4).DeptID = "4"
    Emps(4).JobCode = "4"
    
    For Index = 0 To 4
        Employees.Add Emps(Index)
    Next Index
    
    'Act:
    Set Filtered_EC = Employees.Filter(DeptIDFilter:="FILTER")
    
    'Assert:
    Assert.IsTrue Filtered_EC.Item(1).DeptID = "FILTER"
    Assert.IsTrue Filtered_EC.Item(1).Name = "Emp 0"
    Assert.IsTrue Filtered_EC.Item(2).DeptID = "FILTER"
    Assert.IsTrue Filtered_EC.Item(2).Name = "Emp 1"
    Assert.IsTrue Filtered_EC.Count = 2

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_Filter_JobCode()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim Filtered_EC As EmployeeCollection
    Dim Emps(5) As Employee
    Dim Index As Long
    
    Set Employees = New EmployeeCollection
    For Index = 0 To 4
        Set Emps(Index) = New Employee
    Next Index
    
    Emps(0).EmplID = "0"
    Emps(0).Name = "Emp 0"
    Emps(0).DeptID = "0"
    Emps(0).JobCode = "FILTER"
    
    Emps(1).EmplID = "1"
    Emps(1).Name = "Emp 1"
    Emps(1).DeptID = "1"
    Emps(1).JobCode = "FILTER"
    
    Emps(2).EmplID = "2"
    Emps(2).Name = "Emp 2"
    Emps(2).DeptID = "2"
    Emps(2).JobCode = "2"
    
    Emps(3).EmplID = "3"
    Emps(3).Name = "Emp 3"
    Emps(3).DeptID = "3"
    Emps(3).JobCode = "3"
    
    Emps(4).EmplID = "4"
    Emps(4).Name = "Emp 4"
    Emps(4).DeptID = "4"
    Emps(4).JobCode = "4"
    
    For Index = 0 To 4
        Employees.Add Emps(Index)
    Next Index
    
    'Act:
    Set Filtered_EC = Employees.Filter(JobCodeFilter:="FILTER")
    
    'Assert:
    Assert.IsTrue Filtered_EC.Item(1).JobCode = "FILTER"
    Assert.IsTrue Filtered_EC.Item(1).Name = "Emp 0"
    Assert.IsTrue Filtered_EC.Item(2).JobCode = "FILTER"
    Assert.IsTrue Filtered_EC.Item(2).Name = "Emp 1"
    Assert.IsTrue Filtered_EC.Count = 2

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_ToArrayContainer()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmployeeArray As ArrayContainer
    Dim Employees As EmployeeCollection
    Dim Emp(1 To 5) As Employee
    Dim Names(1 To 5) As String
    Dim Index As Long
    Dim Data As Variant
    
    Set Employees = New EmployeeCollection
    
    Names(1) = "Abby"
    Names(2) = "Brenda"
    Names(3) = "Chelsea"
    Names(4) = "Dorothy"
    Names(5) = "Ericka"
    
    For Index = 1 To 5
        Set Emp(Index) = New Employee
        Emp(Index).EmplID = Trim$(Str$(Index * 1000))
        Emp(Index).Name = Names(Index)
        Emp(Index).DeptID = Trim$(Str$(111 * Index))
        Emp(Index).JobCode = Trim$(Str$(Index))
        Emp(Index).HoursWorked("OTH") = Index * 5
        
        Employees.Add Emp(Index)
    Next Index
    
    'Act:
    Set EmployeeArray = Employees.ToArrayContainer()
    Data = EmployeeArray.Data
    
    'Assert:
    Assert.IsTrue EmployeeArray.Rows = Employees.Count + 1         ' Count + Headers
    Assert.IsTrue Data(0, 0) = "EmplID"
    Assert.IsTrue Data(0, 1) = "Name"
    Assert.IsTrue Data(0, 2) = "DeptID"
    Assert.IsTrue Data(0, 3) = "JobCode"
    Assert.IsTrue Data(0, 4) = "Hours"
    Assert.IsTrue Data(0, 5) = "FTE%"
    
    Assert.IsTrue Data(1, 0) = "1000"
    Assert.IsTrue Data(1, 1) = Names(1)
    Assert.IsTrue Data(1, 2) = "111"
    Assert.IsTrue Data(1, 3) = "1"
    Assert.IsTrue Data(1, 4) = 5
    Assert.IsTrue Data(1, 5) = Round(5 * 100 / 198, 2)
    
    Assert.IsTrue Data(2, 0) = "2000"
    Assert.IsTrue Data(2, 1) = Names(2)
    Assert.IsTrue Data(2, 2) = "222"
    Assert.IsTrue Data(2, 3) = "2"
    Assert.IsTrue Data(2, 4) = 10
    Assert.IsTrue Data(2, 5) = Round(10 * 100 / 198, 2)
    
    Assert.IsTrue Data(3, 0) = "3000"
    Assert.IsTrue Data(3, 1) = Names(3)
    Assert.IsTrue Data(3, 2) = "333"
    Assert.IsTrue Data(3, 3) = "3"
    Assert.IsTrue Data(3, 4) = 15
    Assert.IsTrue Data(3, 5) = Round(15 * 100 / 198, 2)
    
    Assert.IsTrue Data(4, 0) = "4000"
    Assert.IsTrue Data(4, 1) = Names(4)
    Assert.IsTrue Data(4, 2) = "444"
    Assert.IsTrue Data(4, 3) = "4"
    Assert.IsTrue Data(4, 4) = 20
    Assert.IsTrue Data(4, 5) = Round(20 * 100 / 198, 2)

    Assert.IsTrue Data(5, 0) = "5000"
    Assert.IsTrue Data(5, 1) = Names(5)
    Assert.IsTrue Data(5, 2) = "555"
    Assert.IsTrue Data(5, 3) = "5"
    Assert.IsTrue Data(5, 4) = 25
    Assert.IsTrue Data(5, 5) = Round(25 * 100 / 198, 2)
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_ToArrayContainer_EmplID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmployeeArray As ArrayContainer
    Dim Employees As EmployeeCollection
    Dim Emp(1 To 5) As Employee
    Dim Names(1 To 5) As String
    Dim Index As Long
    Dim Data As Variant
    
    Set Employees = New EmployeeCollection
    
    Names(1) = "Abby"
    Names(2) = "Brenda"
    Names(3) = "Chelsea"
    Names(4) = "Dorothy"
    Names(5) = "Ericka"
    
    For Index = 1 To 5
        Set Emp(Index) = New Employee
        Emp(Index).EmplID = Trim$(Str$(Index * 1000))
        Emp(Index).Name = Names(Index)
        Emp(Index).DeptID = Trim$(Str$(111 * Index))
        Emp(Index).JobCode = Trim$(Str$(Index))
        Emp(Index).HoursWorked("OTH") = Index * 5
        
        Employees.Add Emp(Index)
    Next Index
    
    'Act:
    Set EmployeeArray = Employees.ToArrayContainer()
    
    Data = EmployeeArray.Data
    
    'Assert:
    Assert.IsTrue Data(0, 0) = "EmplID"
    Assert.IsTrue Data(1, 0) = "1000"
    Assert.IsTrue Data(2, 0) = "2000"
    Assert.IsTrue Data(3, 0) = "3000"
    Assert.IsTrue Data(4, 0) = "4000"
    Assert.IsTrue Data(5, 0) = "5000"
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_CountUniqueEmplIDs_NoEmployees()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Set Employees = New EmployeeCollection
    
    'Act:
    ' No Actions to Take
    
    'Assert:
    Assert.IsTrue Employees.CountUniqueEmplIDs = 0
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_CountUniqueEmplIDs_OneEmployees()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim Emp As Employee
    Set Employees = New EmployeeCollection
    Set Emp = New Employee
    
    Emp.EmplID = "111"
    Employees.Add Emp
    
    'Act:
    ' No Actions to Take
    
    'Assert:
    Assert.IsTrue Employees.CountUniqueEmplIDs = 1
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_CountUniqueEmplIDs_TwoEmployees_TwoUniqueEmployees()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim EmpArray(1 To 2) As Employee
    Set Employees = New EmployeeCollection
    Set EmpArray(1) = New Employee
    Set EmpArray(2) = New Employee
    
    EmpArray(1).EmplID = "111"
    EmpArray(2).EmplID = "222"
    Employees.Add EmpArray(1)
    Employees.Add EmpArray(2)
    
    'Act:
    ' No Actions to Take
    
    'Assert:
    Assert.IsTrue Employees.CountUniqueEmplIDs = 2
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_CountUniqueEmplIDs_TwoEmployees_OneUniqueEmployees()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmpCollection As EmployeeCollection
    Dim EmpArray(1 To 2) As Employee
    Set EmpCollection = New EmployeeCollection
    Set EmpArray(1) = New Employee
    Set EmpArray(2) = New Employee
    
    EmpArray(1).EmplID = "111"
    EmpArray(2).EmplID = "111"
    EmpCollection.Add EmpArray(1)
    EmpCollection.Add EmpArray(2)
    
    'Act:
    ' No Actions to Take
    
    'Assert:
    Assert.IsTrue EmpCollection.CountUniqueEmplIDs = 1
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub Test_MergeAllEmployeesOnEmplID_UniqueEmployees()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmpCollection As EmployeeCollection
    Dim MergedEmpCollection As EmployeeCollection
    Dim EmpArray(1 To 10) As Employee
    Dim Index As Long
    
    Set EmpCollection = New EmployeeCollection
    For Index = 1 To 10
        Set EmpArray(Index) = New Employee
        EmpArray(Index).EmplID = Trim$(Str$(Index))
        EmpArray(Index).DeptID = "1"
        EmpArray(Index).JobCode = "1"
        EmpArray(Index).HoursWorked("01A") = 1
        EmpCollection.Add EmpArray(Index)
    Next Index
    
    'Act:
    
    Set MergedEmpCollection = EmpCollection.MergeAllEmployeesOnEmplID()
    
    'Assert:
    Assert.IsTrue MergedEmpCollection.Count = 10

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub Test_MergeAllEmployeesOnEmplID_OneEmployee()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EmpCollection As EmployeeCollection
    Dim MergedEmpCollection As EmployeeCollection
    Dim Emp As Employee
    
    Set EmpCollection = New EmployeeCollection
    Set Emp = New Employee
    Emp.EmplID = "1"
    Emp.DeptID = "1"
    Emp.JobCode = "1"
    Emp.HoursWorked("01A") = 1
    EmpCollection.Add Emp
    
    'Act:
    Set MergedEmpCollection = EmpCollection.MergeAllEmployeesOnEmplID()
    
    'Assert:
    Assert.IsTrue MergedEmpCollection.Count = 1
    Assert.IsTrue MergedEmpCollection.Item(1).EmplID = "1"
    Assert.IsTrue MergedEmpCollection.Item(1).DeptID = "1"
    Assert.IsTrue MergedEmpCollection.Item(1).JobCode = "1"
    Assert.IsTrue MergedEmpCollection.Item(1).HoursWorked = 1
    Assert.IsTrue MergedEmpCollection.Item(1).HoursWorked("01A") = 1

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_CreateFromWorksheet_Hourly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim TestWorkbook As Workbook
    Dim HourlyWorksheet As Worksheet
    Dim FilePath As String
    Dim FileName As String
    FilePath = ThisWorkbook.Path & "/TestData/"
    FileName = "Test Workbook - Appointed and Hourly.xlsx"
    
    Workbooks.Open FileName:=FilePath & FileName, ReadOnly:=True
    
    Set TestWorkbook = Workbooks.Item(FileName)
    
    Set HourlyWorksheet = WBTools.GetSheetLike("*Hourly*", TestWorkbook)
    Set Employees = New EmployeeCollection
    
    'Act:
    Set Employees = Employees.CreateEmployeeCollectionFromWorksheet(HourlyWorksheet, True)
    
    'Assert:
    Dim Index As Long
    Dim PayPeriodString As String
    For Index = 0 To Employees.Count - 1
        If Index Mod 2 = 0 Then
            PayPeriodString = Format$(WorksheetFunction.RoundUp((Index + 1) / 2, 0), "00") & "A"
        Else
            PayPeriodString = Format$(WorksheetFunction.RoundUp((Index + 1) / 2, 0), "00") & "B"
        End If
                
        ' Emp only has 10 hours, and those hours are in a specific period (different for each Emp)
        Assert.IsTrue Employees.Item(Index + 1).HoursWorked() = 10
        Assert.IsTrue Employees.Item(Index + 1).HoursWorked(PayPeriodString) = 10
        Assert.IsTrue Employees.Item(Index + 1).DeptID = Format$(Index + 1, "00000")
    Next Index

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    TestWorkbook.Close
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_CreateFromWorksheet_Appointed_WithIndependentStudy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim TestWorkbook As Workbook
    Dim AppointedWorksheet As Worksheet
    Dim FilePath As String
    Dim FileName As String
    FilePath = ThisWorkbook.Path & "/TestData/"
    FileName = "Test Workbook - Appointed and Hourly - With Independent Study.xlsx"
    
    Workbooks.Open FileName:=FilePath & FileName, ReadOnly:=True
    
    Set TestWorkbook = Workbooks.Item(FileName)
    
    Set AppointedWorksheet = WBTools.GetSheetLike("*Appointed*", TestWorkbook)
    Set Employees = New EmployeeCollection
    
    'Act:
    Set Employees = Employees.CreateEmployeeCollectionFromWorksheet(AppointedWorksheet, False)
    
    'Assert:
    Assert.IsTrue Employees.Item(1).HoursWorked = 11
    Assert.IsTrue Employees.Item(2).HoursWorked = 22
    ' Skip the 3rd row (independent study)
    Assert.IsTrue Employees.Item(3).HoursWorked = 44
    Assert.IsTrue Employees.Item(4).HoursWorked = 55
    Assert.IsTrue Employees.Count = 23
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    TestWorkbook.Close
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub Test_MergeAllEmployees_10Employees_1UniqueEmployee_Count()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmpCollection As EmployeeCollection
    Dim MergedEmpCollection As EmployeeCollection
    Dim EmpArray(1 To 10) As Employee
    Dim Index As Long
    
    Set EmpCollection = New EmployeeCollection
    
    For Index = 1 To 10
        Set EmpArray(Index) = New Employee
        EmpArray(Index).EmplID = "1"
        EmpArray(Index).Name = "John Doe"
        EmpArray(Index).DeptID = "DEPT"
        EmpArray(Index).JobCode = "JOB"
        EmpArray(Index).HoursWorked("01A") = 1
        
        EmpCollection.Add EmpArray(Index)
    Next Index
    
    'Act:
    Set MergedEmpCollection = EmpCollection.MergeAllEmployees()
    
    'Assert:
    Assert.IsTrue MergedEmpCollection.Count = 1

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub Test_MergeAllEmployees_10Employees_1UniqueEmployee_EmplID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmpCollection As EmployeeCollection
    Dim MergedEmpCollection As EmployeeCollection
    Dim EmpArray(1 To 10) As Employee
    Dim Index As Long
    
    Set EmpCollection = New EmployeeCollection
    
    For Index = 1 To 10
        Set EmpArray(Index) = New Employee
        EmpArray(Index).EmplID = "1"
        EmpArray(Index).Name = "John Doe"
        EmpArray(Index).DeptID = "DEPT"
        EmpArray(Index).JobCode = "JOB"
        EmpArray(Index).HoursWorked("01A") = 1
        
        EmpCollection.Add EmpArray(Index)
    Next Index
    
    'Act:
    Set MergedEmpCollection = EmpCollection.MergeAllEmployees()
    
    'Assert:
    Assert.IsTrue MergedEmpCollection.Item(1).EmplID = "1"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub Test_MergeAllEmployees_10Employees_1UniqueEmployee_Name()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmpCollection As EmployeeCollection
    Dim MergedEmpCollection As EmployeeCollection
    Dim EmpArray(1 To 10) As Employee
    Dim Index As Long
    
    Set EmpCollection = New EmployeeCollection
    
    For Index = 1 To 10
        Set EmpArray(Index) = New Employee
        EmpArray(Index).EmplID = "1"
        EmpArray(Index).Name = "John Doe"
        EmpArray(Index).DeptID = "DEPT"
        EmpArray(Index).JobCode = "JOB"
        EmpArray(Index).HoursWorked("01A") = 1
        
        EmpCollection.Add EmpArray(Index)
    Next Index
    
    'Act:
    Set MergedEmpCollection = EmpCollection.MergeAllEmployees()
    
    'Assert:
    Assert.IsTrue MergedEmpCollection.Item(1).Name = "John Doe"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub Test_MergeAllEmployees_10Employees_1UniqueEmployee_DeptID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmpCollection As EmployeeCollection
    Dim MergedEmpCollection As EmployeeCollection
    Dim EmpArray(1 To 10) As Employee
    Dim Index As Long
    
    Set EmpCollection = New EmployeeCollection
    
    For Index = 1 To 10
        Set EmpArray(Index) = New Employee
        EmpArray(Index).EmplID = "1"
        EmpArray(Index).Name = "John Doe"
        EmpArray(Index).DeptID = "DEPT"
        EmpArray(Index).JobCode = "JOB"
        EmpArray(Index).HoursWorked("01A") = 1
        
        EmpCollection.Add EmpArray(Index)
    Next Index
    
    'Act:
    Set MergedEmpCollection = EmpCollection.MergeAllEmployees()
    
    'Assert:
    Assert.IsTrue MergedEmpCollection.Item(1).DeptID = "DEPT"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub Test_MergeAllEmployees_10Employees_1UniqueEmployee_JobCode()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmpCollection As EmployeeCollection
    Dim MergedEmpCollection As EmployeeCollection
    Dim EmpArray(1 To 10) As Employee
    Dim Index As Long
    
    Set EmpCollection = New EmployeeCollection
    
    For Index = 1 To 10
        Set EmpArray(Index) = New Employee
        EmpArray(Index).EmplID = "1"
        EmpArray(Index).Name = "John Doe"
        EmpArray(Index).DeptID = "DEPT"
        EmpArray(Index).JobCode = "JOB"
        EmpArray(Index).HoursWorked("01A") = 1
        
        EmpCollection.Add EmpArray(Index)
    Next Index
    
    'Act:
    Set MergedEmpCollection = EmpCollection.MergeAllEmployees()
    
    'Assert:
    Assert.IsTrue MergedEmpCollection.Item(1).JobCode = "JOB"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub Test_MergeAllEmployees_10Employees_1UniqueEmployee_HoursWorked()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EmpCollection As EmployeeCollection
    Dim MergedEmpCollection As EmployeeCollection
    Dim EmpArray(1 To 10) As Employee
    Dim Index As Long
    
    Set EmpCollection = New EmployeeCollection
    
    For Index = 1 To 10
        Set EmpArray(Index) = New Employee
        EmpArray(Index).EmplID = "1"
        EmpArray(Index).Name = "John Doe"
        EmpArray(Index).DeptID = "DEPT"
        EmpArray(Index).JobCode = "JOB"
        EmpArray(Index).HoursWorked("01A") = 1
        
        EmpCollection.Add EmpArray(Index)
    Next Index
    
    'Act:
    Set MergedEmpCollection = EmpCollection.MergeAllEmployees()
    
    'Assert:
    Assert.IsTrue MergedEmpCollection.Item(1).HoursWorked = 10
    Assert.IsTrue MergedEmpCollection.Item(1).HoursWorked("01A") = 10

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Multi-Method")
Private Sub Test_MergeAllEmployees_ToArrayContainer_10Employees_1UniqueEmployee_Rows()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmpCollection As EmployeeCollection
    Dim MergedEmpCollection As EmployeeCollection
    Dim MergedEmpArray As ArrayContainer
    Dim EmpArray(1 To 10) As Employee
    Dim Index As Long
    
    Set EmpCollection = New EmployeeCollection
    
    For Index = 1 To 10
        Set EmpArray(Index) = New Employee
        EmpArray(Index).EmplID = "1"
        EmpArray(Index).Name = "John Doe"
        EmpArray(Index).DeptID = "DEPT"
        EmpArray(Index).JobCode = "JOB"
        EmpArray(Index).HoursWorked("01A") = 1
        
        EmpCollection.Add EmpArray(Index)
    Next Index
    
    'Act:
    Set MergedEmpCollection = EmpCollection.MergeAllEmployees()
    Set MergedEmpArray = MergedEmpCollection.ToArrayContainer()
    
    'Assert:
    Assert.IsTrue MergedEmpArray.Rows = 2

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Multi-Method")
Private Sub Test_MergeAllEmployees_ToArrayContainer_10Employees_1UniqueEmployee_Rows_NoHeader()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmpCollection As EmployeeCollection
    Dim MergedEmpCollection As EmployeeCollection
    Dim MergedEmpArray As ArrayContainer
    Dim EmpArray(1 To 10) As Employee
    Dim Index As Long
    
    Set EmpCollection = New EmployeeCollection
    
    For Index = 1 To 10
        Set EmpArray(Index) = New Employee
        EmpArray(Index).EmplID = "1"
        EmpArray(Index).Name = "John Doe"
        EmpArray(Index).DeptID = "DEPT"
        EmpArray(Index).JobCode = "JOB"
        EmpArray(Index).HoursWorked("01A") = 1
        
        EmpCollection.Add EmpArray(Index)
    Next Index
    
    'Act:
    Set MergedEmpCollection = EmpCollection.MergeAllEmployees()
    Set MergedEmpArray = MergedEmpCollection.ToArrayContainer(Headers:=False)
    
    'Assert:
    Assert.IsTrue MergedEmpArray.Rows = 1

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Multi-Method")
Private Sub Test_MergeAllEmployeesOnDeptID_ToArrayContainer_10Employees_1UniqueEmployee_Rows_NoHeader()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmpCollection As EmployeeCollection
    Dim MergedEmpCollection As EmployeeCollection
    Dim MergedEmpArray As ArrayContainer
    Dim EmpArray(1 To 10) As Employee
    Dim Index As Long
    
    Set EmpCollection = New EmployeeCollection
    
    For Index = 1 To 10
        Set EmpArray(Index) = New Employee
        EmpArray(Index).EmplID = "1"
        EmpArray(Index).Name = "John Doe"
        EmpArray(Index).DeptID = "DEPT"
        EmpArray(Index).JobCode = "JOB"
        EmpArray(Index).HoursWorked("01A") = 1
        
        EmpCollection.Add EmpArray(Index)
    Next Index
    
    'Act:
    Set MergedEmpCollection = EmpCollection.MergeAllEmployees()
    Set MergedEmpArray = MergedEmpCollection.ToArrayContainer(Headers:=False, IncludeJobCode:=False)
    
    'Assert:
    Assert.IsTrue MergedEmpArray.Rows = 1

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Multi-Method")
Private Sub Test_MergeAllEmployeesOnDeptID_ToArrayContainer_10Employees_2UniqueEmployee_Rows()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmpCollection As EmployeeCollection
    Dim MergedEmpCollection As EmployeeCollection
    Dim MergedEmpArray As ArrayContainer
    Dim EmpArray(1 To 10) As Employee
    Dim Index As Long
    
    Set EmpCollection = New EmployeeCollection
    
    For Index = 1 To 5
        Set EmpArray(Index) = New Employee
        EmpArray(Index).EmplID = "1"
        EmpArray(Index).Name = "John Doe"
        EmpArray(Index).DeptID = "DEPT"
        EmpArray(Index).JobCode = "JOB"
        EmpArray(Index).HoursWorked("01A") = 1
        
        EmpCollection.Add EmpArray(Index)
    Next Index
    
    For Index = 6 To 10
        Set EmpArray(Index) = New Employee
        EmpArray(Index).EmplID = "2"
        EmpArray(Index).Name = "Jane Doe"
        EmpArray(Index).DeptID = "DEPT"
        EmpArray(Index).JobCode = "JOB"
        EmpArray(Index).HoursWorked("01A") = 1
        
        EmpCollection.Add EmpArray(Index)
    Next Index
    
    'Act:
    Set MergedEmpCollection = EmpCollection.MergeAllEmployees()
    Set MergedEmpArray = MergedEmpCollection.ToArrayContainer(IncludeJobCode:=False)
    
    'Assert:
    Assert.IsTrue MergedEmpArray.Rows = 3

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

