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
    Dim EC As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E As Employee
    Set EC = New EmployeeCollection
    Set E = New Employee
    
    'Act:
    EC.Add E
    
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
    Dim EC As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E As Employee
    Set EC = New EmployeeCollection
    Set E = New Employee
    EC.Add E
    
    'Act:
    EC.Remove (1)
    
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
    Dim EC As EmployeeCollection
    Set EC = New EmployeeCollection
    
    'Act:
    ' No Actions to take
    
    'Assert:
    Assert.IsTrue EC.Count() = 0

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
    Dim EC As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E As Employee
    Set EC = New EmployeeCollection
    Set E = New Employee
    
    'Act:
    EC.Add E
    
    'Assert:
    Assert.IsTrue EC.Count() = 1

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
    Dim EC As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E As Employee
    Set EC = New EmployeeCollection
    Set E = New Employee
    
    'Act:
    EC.Add E
    EC.Remove 1
    
    'Assert:
    Assert.IsTrue EC.Count() = 0

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
    '@Ignore UseMeaningfulName
    Dim EC1 As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim EC2 As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E1 As Employee
    '@Ignore UseMeaningfulName
    Dim E2 As Employee
    Set EC1 = New EmployeeCollection
    Set EC2 = New EmployeeCollection
    Set E1 = New Employee
    Set E2 = New Employee
    EC1.Add E1
    EC2.Add E2
    
    'Act:
    EC1.Concat EC2
    
    'Assert:
    Assert.IsTrue EC1.Count() = 2

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
    Dim EC As EmployeeCollection
    Dim Filtered_EC As EmployeeCollection
    Dim Emps(5) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
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
        EC.Add Emps(Index)
    Next Index
    
    'Act:
    Set Filtered_EC = EC.Filter(EmplIDFilter:="FILTER")
    
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
    Dim EC As EmployeeCollection
    Dim Filtered_EC As EmployeeCollection
    Dim Emps(5) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
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
        EC.Add Emps(Index)
    Next Index
    
    'Act:
    Set Filtered_EC = EC.Filter(DeptIDFilter:="FILTER")
    
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
    Dim EC As EmployeeCollection
    Dim Filtered_EC As EmployeeCollection
    Dim Emps(5) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
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
        EC.Add Emps(Index)
    Next Index
    
    'Act:
    Set Filtered_EC = EC.Filter(JobCodeFilter:="FILTER")
    
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
    '@Ignore UseMeaningfulName
    Dim AC As ArrayContainer
    Dim EC As EmployeeCollection
    Dim Emp(1 To 5) As Employee
    Dim Names(1 To 5) As String
    Dim Index As Long
    Dim Data As Variant
    
    Set EC = New EmployeeCollection
    
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
        
        EC.Add Emp(Index)
    Next Index
    
    'Act:
    Set AC = EC.ToArrayContainer()
    Data = AC.Data
    
    'Assert:
    Assert.IsTrue AC.Rows = EC.Count + 1         ' Count + Headers
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
    '@Ignore UseMeaningfulName
    Dim AC As ArrayContainer
    Dim EC As EmployeeCollection
    Dim Emp(1 To 5) As Employee
    Dim Names(1 To 5) As String
    Dim Index As Long
    Dim Data As Variant
    
    Set EC = New EmployeeCollection
    
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
        
        EC.Add Emp(Index)
    Next Index
    
    'Act:
    Set AC = EC.ToArrayContainer()
    
    Data = AC.Data
    
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
    '@Ignore UseMeaningfulName
    Dim EC As EmployeeCollection
    Set EC = New EmployeeCollection
    
    'Act:
    ' No Actions to Take
    
    'Assert:
    Assert.IsTrue EC.CountUniqueEmplIDs = 0
    
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
    '@Ignore UseMeaningfulName
    Dim EC As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E As Employee
    Set EC = New EmployeeCollection
    Set E = New Employee
    
    E.EmplID = "111"
    EC.Add E
    
    'Act:
    ' No Actions to Take
    
    'Assert:
    Assert.IsTrue EC.CountUniqueEmplIDs = 1
    
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
    '@Ignore UseMeaningfulName
    Dim EC As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E(1 To 2) As Employee
    Set EC = New EmployeeCollection
    Set E(1) = New Employee
    Set E(2) = New Employee
    
    E(1).EmplID = "111"
    E(2).EmplID = "222"
    EC.Add E(1)
    EC.Add E(2)
    
    'Act:
    ' No Actions to Take
    
    'Assert:
    Assert.IsTrue EC.CountUniqueEmplIDs = 2
    
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
    '@Ignore UseMeaningfulName
    Dim EC As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E(1 To 2) As Employee
    Set EC = New EmployeeCollection
    Set E(1) = New Employee
    Set E(2) = New Employee
    
    E(1).EmplID = "111"
    E(2).EmplID = "111"
    EC.Add E(1)
    EC.Add E(2)
    
    'Act:
    ' No Actions to Take
    
    'Assert:
    Assert.IsTrue EC.CountUniqueEmplIDs = 1
    
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
    '@Ignore UseMeaningfulName
    Dim EC As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E(1 To 10) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    For Index = 1 To 10
        Set E(Index) = New Employee
        E(Index).EmplID = Trim$(Str$(Index))
        E(Index).DeptID = "1"
        E(Index).JobCode = "1"
        E(Index).HoursWorked("01A") = 1
        EC.Add E(Index)
    Next Index
    
    'Act:
    
    Set EC_Merged = EC.MergeAllEmployeesOnEmplID()
    
    'Assert:
    Assert.IsTrue EC_Merged.Count = 10

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
    Dim EC As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E As Employee
    
    Set EC = New EmployeeCollection
    Set E = New Employee
    E.EmplID = "1"
    E.DeptID = "1"
    E.JobCode = "1"
    E.HoursWorked("01A") = 1
    EC.Add E
    
    'Act:
    Set EC_Merged = EC.MergeAllEmployeesOnEmplID()
    
    'Assert:
    Assert.IsTrue EC_Merged.Count = 1
    Assert.IsTrue EC_Merged.Item(1).EmplID = "1"
    Assert.IsTrue EC_Merged.Item(1).DeptID = "1"
    Assert.IsTrue EC_Merged.Item(1).JobCode = "1"
    Assert.IsTrue EC_Merged.Item(1).HoursWorked = 1
    Assert.IsTrue EC_Merged.Item(1).HoursWorked("01A") = 1

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
    Dim EC As EmployeeCollection
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim FilePath As String
    Dim FileName As String
    FilePath = ThisWorkbook.Path & "/TestData/"
    FileName = "Test Workbook - Appointed and Hourly.xlsx"
    
    Workbooks.Open FileName:=FilePath & FileName, ReadOnly:=True
    
    Set wb = Workbooks.Item(FileName)
    
    Set ws = WBTools.GetSheetLike("*Hourly*", wb)
    Set EC = New EmployeeCollection
    
    'Act:
    Set EC = EC.CreateEmployeeCollectionFromWorksheet(ws, True)
    
    'Assert:
    Dim Index As Long
    Dim PayPeriodString As String
    For Index = 0 To EC.Count - 1
        If Index Mod 2 = 0 Then
            PayPeriodString = Format$(WorksheetFunction.RoundUp((Index + 1) / 2, 0), "00") & "A"
        Else
            PayPeriodString = Format$(WorksheetFunction.RoundUp((Index + 1) / 2, 0), "00") & "B"
        End If
                
        ' Emp only has 10 hours, and those hours are in a specific period (different for each Emp)
        Assert.IsTrue EC.Item(Index + 1).HoursWorked() = 10
        Assert.IsTrue EC.Item(Index + 1).HoursWorked(PayPeriodString) = 10
        Assert.IsTrue EC.Item(Index + 1).DeptID = Format$(Index + 1, "00000")
    Next Index

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    wb.Close
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_CreateFromWorksheet_Appointed_WithIndependentStudy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC As EmployeeCollection
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim FilePath As String
    Dim FileName As String
    FilePath = ThisWorkbook.Path & "/TestData/"
    FileName = "Test Workbook - Appointed and Hourly - With Independent Study.xlsx"
    
    Workbooks.Open FileName:=FilePath & FileName, ReadOnly:=True
    
    Set wb = Workbooks.Item(FileName)
    
    Set ws = WBTools.GetSheetLike("*Appointed*", wb)
    Set EC = New EmployeeCollection
    
    'Act:
    Set EC = EC.CreateEmployeeCollectionFromWorksheet(ws, False)
    
    'Assert:
    Assert.IsTrue EC.Item(1).HoursWorked = 11
    Assert.IsTrue EC.Item(2).HoursWorked = 22
    ' Skip the 3rd row (independent study)
    Assert.IsTrue EC.Item(3).HoursWorked = 44
    Assert.IsTrue EC.Item(4).HoursWorked = 55
    Assert.IsTrue EC.Count = 23
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    wb.Close
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Merge")
Private Sub Test_MergeAllEmployees_10Employees_1UniqueEmployee_Count()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E(1 To 10) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 10
        Set E(Index) = New Employee
        E(Index).EmplID = "1"
        E(Index).Name = "John Doe"
        E(Index).DeptID = "DEPT"
        E(Index).JobCode = "JOB"
        E(Index).HoursWorked("01A") = 1
        
        EC.Add E(Index)
    Next Index
    
    'Act:
    Set EC_Merged = EC.MergeAllEmployees()
    
    'Assert:
    Assert.IsTrue EC_Merged.Count = 1

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
    Dim EC As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E(1 To 10) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 10
        Set E(Index) = New Employee
        E(Index).EmplID = "1"
        E(Index).Name = "John Doe"
        E(Index).DeptID = "DEPT"
        E(Index).JobCode = "JOB"
        E(Index).HoursWorked("01A") = 1
        
        EC.Add E(Index)
    Next Index
    
    'Act:
    Set EC_Merged = EC.MergeAllEmployees()
    
    'Assert:
    Assert.IsTrue EC_Merged.Item(1).EmplID = "1"

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
    Dim EC As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E(1 To 10) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 10
        Set E(Index) = New Employee
        E(Index).EmplID = "1"
        E(Index).Name = "John Doe"
        E(Index).DeptID = "DEPT"
        E(Index).JobCode = "JOB"
        E(Index).HoursWorked("01A") = 1
        
        EC.Add E(Index)
    Next Index
    
    'Act:
    Set EC_Merged = EC.MergeAllEmployees()
    
    'Assert:
    Assert.IsTrue EC_Merged.Item(1).Name = "John Doe"

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
    Dim EC As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E(1 To 10) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 10
        Set E(Index) = New Employee
        E(Index).EmplID = "1"
        E(Index).Name = "John Doe"
        E(Index).DeptID = "DEPT"
        E(Index).JobCode = "JOB"
        E(Index).HoursWorked("01A") = 1
        
        EC.Add E(Index)
    Next Index
    
    'Act:
    Set EC_Merged = EC.MergeAllEmployees()
    
    'Assert:
    Assert.IsTrue EC_Merged.Item(1).DeptID = "DEPT"

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
    Dim EC As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E(1 To 10) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 10
        Set E(Index) = New Employee
        E(Index).EmplID = "1"
        E(Index).Name = "John Doe"
        E(Index).DeptID = "DEPT"
        E(Index).JobCode = "JOB"
        E(Index).HoursWorked("01A") = 1
        
        EC.Add E(Index)
    Next Index
    
    'Act:
    Set EC_Merged = EC.MergeAllEmployees()
    
    'Assert:
    Assert.IsTrue EC_Merged.Item(1).JobCode = "JOB"

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
    Dim EC As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim E(1 To 10) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 10
        Set E(Index) = New Employee
        E(Index).EmplID = "1"
        E(Index).Name = "John Doe"
        E(Index).DeptID = "DEPT"
        E(Index).JobCode = "JOB"
        E(Index).HoursWorked("01A") = 1
        
        EC.Add E(Index)
    Next Index
    
    'Act:
    Set EC_Merged = EC.MergeAllEmployees()
    
    'Assert:
    Assert.IsTrue EC_Merged.Item(1).HoursWorked = 10
    Assert.IsTrue EC_Merged.Item(1).HoursWorked("01A") = 10

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
    Dim EC As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
    Dim AC_Merged As ArrayContainer
    '@Ignore UseMeaningfulName
    Dim E(1 To 10) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 10
        Set E(Index) = New Employee
        E(Index).EmplID = "1"
        E(Index).Name = "John Doe"
        E(Index).DeptID = "DEPT"
        E(Index).JobCode = "JOB"
        E(Index).HoursWorked("01A") = 1
        
        EC.Add E(Index)
    Next Index
    
    'Act:
    Set EC_Merged = EC.MergeAllEmployees()
    Set AC_Merged = EC_Merged.ToArrayContainer()
    
    'Assert:
    Assert.IsTrue AC_Merged.Rows = 2

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
    Dim EC As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
    Dim AC_Merged As ArrayContainer
    '@Ignore UseMeaningfulName
    Dim E(1 To 10) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 10
        Set E(Index) = New Employee
        E(Index).EmplID = "1"
        E(Index).Name = "John Doe"
        E(Index).DeptID = "DEPT"
        E(Index).JobCode = "JOB"
        E(Index).HoursWorked("01A") = 1
        
        EC.Add E(Index)
    Next Index
    
    'Act:
    Set EC_Merged = EC.MergeAllEmployees()
    Set AC_Merged = EC_Merged.ToArrayContainer(Headers:=False)
    
    'Assert:
    Assert.IsTrue AC_Merged.Rows = 1

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
    Dim EC As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
    Dim AC_Merged As ArrayContainer
    '@Ignore UseMeaningfulName
    Dim E(1 To 10) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 10
        Set E(Index) = New Employee
        E(Index).EmplID = "1"
        E(Index).Name = "John Doe"
        E(Index).DeptID = "DEPT"
        E(Index).JobCode = "JOB"
        E(Index).HoursWorked("01A") = 1
        
        EC.Add E(Index)
    Next Index
    
    'Act:
    Set EC_Merged = EC.MergeAllEmployees()
    Set AC_Merged = EC_Merged.ToArrayContainer(Headers:=False, IncludeJobCode:=False)
    
    'Assert:
    Assert.IsTrue AC_Merged.Rows = 1

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
    Dim EC As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
    Dim AC_Merged As ArrayContainer
    '@Ignore UseMeaningfulName
    Dim E(1 To 10) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 5
        Set E(Index) = New Employee
        E(Index).EmplID = "1"
        E(Index).Name = "John Doe"
        E(Index).DeptID = "DEPT"
        E(Index).JobCode = "JOB"
        E(Index).HoursWorked("01A") = 1
        
        EC.Add E(Index)
    Next Index
    
    For Index = 6 To 10
        Set E(Index) = New Employee
        E(Index).EmplID = "2"
        E(Index).Name = "Jane Doe"
        E(Index).DeptID = "DEPT"
        E(Index).JobCode = "JOB"
        E(Index).HoursWorked("01A") = 1
        
        EC.Add E(Index)
    Next Index
    
    'Act:
    Set EC_Merged = EC.MergeAllEmployees()
    Set AC_Merged = EC_Merged.ToArrayContainer(IncludeJobCode:=False)
    
    'Assert:
    Assert.IsTrue AC_Merged.Rows = 3

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

