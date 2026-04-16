Attribute VB_Name = "TestModule_Main"
'@IgnoreModule UseMeaningfulName
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

'@TestMethod("Calculations")
Private Sub Test_CalculateFTE()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EmpHours0 As Long
    Dim EmpHours1 As Long
    Dim EmpHours2 As Long
    Dim EmpHours5 As Long
    
    EmpHours0 = 0
    EmpHours1 = 198
    EmpHours2 = 396
    EmpHours5 = 990
    
    Dim FTE0 As Long
    Dim FTE1 As Long
    Dim FTE2 As Long
    Dim FTE5 As Long
    
    'Act:
    FTE0 = Main.CalculateFTE(EmpHours0)
    FTE1 = Main.CalculateFTE(EmpHours1)
    FTE2 = Main.CalculateFTE(EmpHours2)
    FTE5 = Main.CalculateFTE(EmpHours5)
    
    'Assert:
    Assert.IsTrue FTE0 = 0
    Assert.IsTrue FTE1 = 1
    Assert.IsTrue FTE2 = 2
    Assert.IsTrue FTE5 = 5
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_BuildFTECombined_Headers_OneEmployeePerPeriod_HourlyOnly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsFTECombined As Worksheet
    Dim EC As EmployeeCollection
    Dim Emp(1 To 24) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 24
        Set Emp(Index) = New Employee
    Next
    
    Emp(1).Name = "Jan A"
    Emp(2).Name = "Jan B"
    Emp(3).Name = "Feb A"
    Emp(4).Name = "Feb B"
    Emp(5).Name = "Mar A"
    Emp(6).Name = "Mar B"
    Emp(7).Name = "Apr A"
    Emp(8).Name = "Apr B"
    Emp(9).Name = "May A"
    Emp(10).Name = "May B"
    Emp(11).Name = "Jun A"
    Emp(12).Name = "Jun B"
    Emp(13).Name = "Jul A"
    Emp(14).Name = "Jul B"
    Emp(15).Name = "Aug A"
    Emp(16).Name = "Aug B"
    Emp(17).Name = "Sep A"
    Emp(18).Name = "Sep B"
    Emp(19).Name = "Oct A"
    Emp(20).Name = "Oct B"
    Emp(21).Name = "Nov A"
    Emp(22).Name = "Nov B"
    Emp(23).Name = "Dec A"
    Emp(24).Name = "Dec B"
    
    For Index = 1 To 24
        Emp(Index).EmplID = Index
        Emp(Index).JobCode = "JOB"
        Emp(Index).Department = "DEPRT"
        Emp(Index).IsHourly = True
    Next
    
    Emp(1).hoursWorked("01A") = 10
    Emp(2).hoursWorked("01B") = 10
    Emp(3).hoursWorked("02A") = 10
    Emp(4).hoursWorked("02B") = 10
    Emp(5).hoursWorked("03A") = 10
    Emp(6).hoursWorked("03B") = 10
    Emp(7).hoursWorked("04A") = 10
    Emp(8).hoursWorked("04B") = 10
    Emp(9).hoursWorked("05A") = 10
    Emp(10).hoursWorked("05B") = 10
    Emp(11).hoursWorked("06A") = 10
    Emp(12).hoursWorked("06B") = 10
    Emp(13).hoursWorked("07A") = 10
    Emp(14).hoursWorked("07B") = 10
    Emp(15).hoursWorked("08A") = 10
    Emp(16).hoursWorked("08B") = 10
    Emp(17).hoursWorked("09A") = 10
    Emp(18).hoursWorked("09B") = 10
    Emp(19).hoursWorked("10A") = 10
    Emp(20).hoursWorked("10B") = 10
    Emp(21).hoursWorked("11A") = 10
    Emp(22).hoursWorked("11B") = 10
    Emp(23).hoursWorked("12A") = 10
    Emp(24).hoursWorked("12B") = 10
    
    For Index = 1 To 24
        EC.Add Emp(Index)
    Next
    
    'Act:
    
    Set wsFTECombined = Main.BuildFTECombined(EC)
    
    'Assert:
    Assert.IsTrue wsFTECombined.Range("A1").Value = "Empl ID"
    Assert.IsTrue wsFTECombined.Range("B1").Value = "Name (LN, FN)"
    Assert.IsTrue wsFTECombined.Range("C1").Value = "Department"
    Assert.IsTrue wsFTECombined.Range("D1").Value = "Job Code"
    Assert.IsTrue wsFTECombined.Range("E1").Value = "Hours"
    Assert.IsTrue wsFTECombined.Range("F1").Value = "FTE%"
    Assert.IsTrue wsFTECombined.Range("G1").Value = "Source"
 
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Dim AlertsEnabled As Boolean
    
    ' Ensure that DisplayAlerts is disabled, delete a sheet,
    '   then ensure that DisplayAlerts is set to whatever value it held before this action.
    AlertsEnabled = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wsFTECombined.Delete
    Application.DisplayAlerts = AlertsEnabled
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_BuildFTECombined_Data_OneEmployeePerPeriod_HourlyOnly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsFTECombined As Worksheet
    Dim EC As EmployeeCollection
    Dim Emp(1 To 24) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 24
        Set Emp(Index) = New Employee
    Next
    
    Emp(1).Name = "Jan A"
    Emp(2).Name = "Jan B"
    Emp(3).Name = "Feb A"
    Emp(4).Name = "Feb B"
    Emp(5).Name = "Mar A"
    Emp(6).Name = "Mar B"
    Emp(7).Name = "Apr A"
    Emp(8).Name = "Apr B"
    Emp(9).Name = "May A"
    Emp(10).Name = "May B"
    Emp(11).Name = "Jun A"
    Emp(12).Name = "Jun B"
    Emp(13).Name = "Jul A"
    Emp(14).Name = "Jul B"
    Emp(15).Name = "Aug A"
    Emp(16).Name = "Aug B"
    Emp(17).Name = "Sep A"
    Emp(18).Name = "Sep B"
    Emp(19).Name = "Oct A"
    Emp(20).Name = "Oct B"
    Emp(21).Name = "Nov A"
    Emp(22).Name = "Nov B"
    Emp(23).Name = "Dec A"
    Emp(24).Name = "Dec B"
    
    For Index = 1 To 24
        Emp(Index).EmplID = Index
        Emp(Index).JobCode = "JOB"
        Emp(Index).Department = "DEPRT"
        Emp(Index).IsHourly = True
    Next
    
    Emp(1).hoursWorked("01A") = 10
    Emp(2).hoursWorked("01B") = 10
    Emp(3).hoursWorked("02A") = 10
    Emp(4).hoursWorked("02B") = 10
    Emp(5).hoursWorked("03A") = 10
    Emp(6).hoursWorked("03B") = 10
    Emp(7).hoursWorked("04A") = 10
    Emp(8).hoursWorked("04B") = 10
    Emp(9).hoursWorked("05A") = 10
    Emp(10).hoursWorked("05B") = 10
    Emp(11).hoursWorked("06A") = 10
    Emp(12).hoursWorked("06B") = 10
    Emp(13).hoursWorked("07A") = 10
    Emp(14).hoursWorked("07B") = 10
    Emp(15).hoursWorked("08A") = 10
    Emp(16).hoursWorked("08B") = 10
    Emp(17).hoursWorked("09A") = 10
    Emp(18).hoursWorked("09B") = 10
    Emp(19).hoursWorked("10A") = 10
    Emp(20).hoursWorked("10B") = 10
    Emp(21).hoursWorked("11A") = 10
    Emp(22).hoursWorked("11B") = 10
    Emp(23).hoursWorked("12A") = 10
    Emp(24).hoursWorked("12B") = 10
    
    For Index = 1 To 24
        EC.Add Emp(Index)
    Next
    
    Dim DataRange() As Variant
    
    'Act:
    
    Set wsFTECombined = Main.BuildFTECombined(EC)
    
    'Assert:
    
    DataRange = wsFTECombined.Range("A2:G25").Value2
    
    For Index = 1 To 24:
        Assert.IsTrue DataRange(Index, 1) = Emp(Index).EmplID
        Assert.IsTrue DataRange(Index, 2) = Emp(Index).Name
        Assert.IsTrue DataRange(Index, 3) = Emp(Index).Department
        Assert.IsTrue DataRange(Index, 4) = Emp(Index).JobCode
        Assert.IsTrue DataRange(Index, 5) = Emp(Index).hoursWorked
        Assert.IsTrue DataRange(Index, 6) = Emp(Index).hoursWorked / 198 * 100
        Assert.IsTrue DataRange(Index, 7) = Emp(Index).Source
    Next Index
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Dim AlertsEnabled As Boolean
    
    ' Ensure that DisplayAlerts is disabled, delete a sheet,
    '   then ensure that DisplayAlerts is set to whatever value it held before this action.
    AlertsEnabled = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wsFTECombined.Delete
    Application.DisplayAlerts = AlertsEnabled
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_BuildFTESummaryByDepartment_Data_OneEmployeePerPeriod_OneDepartment_HourlyOnly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsFTEComboDept As Worksheet
    Dim EC As EmployeeCollection
    Dim Emp(1 To 24) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 24
        Set Emp(Index) = New Employee
    Next
    
    Emp(1).Name = "Jan A"
    Emp(2).Name = "Jan B"
    Emp(3).Name = "Feb A"
    Emp(4).Name = "Feb B"
    Emp(5).Name = "Mar A"
    Emp(6).Name = "Mar B"
    Emp(7).Name = "Apr A"
    Emp(8).Name = "Apr B"
    Emp(9).Name = "May A"
    Emp(10).Name = "May B"
    Emp(11).Name = "Jun A"
    Emp(12).Name = "Jun B"
    Emp(13).Name = "Jul A"
    Emp(14).Name = "Jul B"
    Emp(15).Name = "Aug A"
    Emp(16).Name = "Aug B"
    Emp(17).Name = "Sep A"
    Emp(18).Name = "Sep B"
    Emp(19).Name = "Oct A"
    Emp(20).Name = "Oct B"
    Emp(21).Name = "Nov A"
    Emp(22).Name = "Nov B"
    Emp(23).Name = "Dec A"
    Emp(24).Name = "Dec B"
    
    For Index = 1 To 24
        Emp(Index).EmplID = Index
        Emp(Index).JobCode = "JOB"
        Emp(Index).Department = "DEPRT"
        Emp(Index).IsHourly = True
    Next
    
    Emp(1).hoursWorked("01A") = 10
    Emp(2).hoursWorked("01B") = 10
    Emp(3).hoursWorked("02A") = 10
    Emp(4).hoursWorked("02B") = 10
    Emp(5).hoursWorked("03A") = 10
    Emp(6).hoursWorked("03B") = 10
    Emp(7).hoursWorked("04A") = 10
    Emp(8).hoursWorked("04B") = 10
    Emp(9).hoursWorked("05A") = 10
    Emp(10).hoursWorked("05B") = 10
    Emp(11).hoursWorked("06A") = 10
    Emp(12).hoursWorked("06B") = 10
    Emp(13).hoursWorked("07A") = 10
    Emp(14).hoursWorked("07B") = 10
    Emp(15).hoursWorked("08A") = 10
    Emp(16).hoursWorked("08B") = 10
    Emp(17).hoursWorked("09A") = 10
    Emp(18).hoursWorked("09B") = 10
    Emp(19).hoursWorked("10A") = 10
    Emp(20).hoursWorked("10B") = 10
    Emp(21).hoursWorked("11A") = 10
    Emp(22).hoursWorked("11B") = 10
    Emp(23).hoursWorked("12A") = 10
    Emp(24).hoursWorked("12B") = 10
    
    For Index = 1 To 24
        EC.Add Emp(Index)
    Next
    
    Dim DataRange() As Variant
    
    'Act:
    
    Set wsFTEComboDept = Main.BuildFTESummaryByDepartment(EC)
    
    'Assert:
    
    DataRange = wsFTEComboDept.Range("A2:E25").Value2
    
    For Index = 1 To 24:
        Assert.IsTrue DataRange(Index, 1) = Emp(Index).EmplID
        Assert.IsTrue DataRange(Index, 2) = Emp(Index).Name
        Assert.IsTrue DataRange(Index, 3) = Emp(Index).Department
        Assert.IsTrue DataRange(Index, 4) = Emp(Index).hoursWorked
        Assert.IsTrue DataRange(Index, 5) = Emp(Index).hoursWorked / 198 * 100
    Next Index
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Dim AlertsEnabled As Boolean
    
    ' Ensure that DisplayAlerts is disabled, delete a sheet,
    '   then ensure that DisplayAlerts is set to whatever value it held before this action.
    AlertsEnabled = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wsFTEComboDept.Delete
    Application.DisplayAlerts = AlertsEnabled
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_BuildFTESummaryByDepartment_Data_OneEmployeePerPeriod_MultipleDepartments_HourlyOnly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsFTEComboDept As Worksheet
    Dim EC As EmployeeCollection
    Dim Emp(1 To 24) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 24
        Set Emp(Index) = New Employee
    Next
    
    Emp(1).Name = "Jan A"
    Emp(2).Name = "Jan B"
    Emp(3).Name = "Feb A"
    Emp(4).Name = "Feb B"
    Emp(5).Name = "Mar A"
    Emp(6).Name = "Mar B"
    Emp(7).Name = "Apr A"
    Emp(8).Name = "Apr B"
    Emp(9).Name = "May A"
    Emp(10).Name = "May B"
    Emp(11).Name = "Jun A"
    Emp(12).Name = "Jun B"
    Emp(13).Name = "Jul A"
    Emp(14).Name = "Jul B"
    Emp(15).Name = "Aug A"
    Emp(16).Name = "Aug B"
    Emp(17).Name = "Sep A"
    Emp(18).Name = "Sep B"
    Emp(19).Name = "Oct A"
    Emp(20).Name = "Oct B"
    Emp(21).Name = "Nov A"
    Emp(22).Name = "Nov B"
    Emp(23).Name = "Dec A"
    Emp(24).Name = "Dec B"
    
    For Index = 1 To 12
        Emp(Index).EmplID = Index
        Emp(Index).JobCode = "JOB"
        Emp(Index).Department = "DEPT1"
        Emp(Index).IsHourly = True
    Next
    
    For Index = 13 To 24
        Emp(Index).EmplID = Index
        Emp(Index).JobCode = "JOB"
        Emp(Index).Department = "DEPT2"
        Emp(Index).IsHourly = True
    Next
    
    Emp(1).hoursWorked("01A") = 10
    Emp(2).hoursWorked("01B") = 10
    Emp(3).hoursWorked("02A") = 10
    Emp(4).hoursWorked("02B") = 10
    Emp(5).hoursWorked("03A") = 10
    Emp(6).hoursWorked("03B") = 10
    Emp(7).hoursWorked("04A") = 10
    Emp(8).hoursWorked("04B") = 10
    Emp(9).hoursWorked("05A") = 10
    Emp(10).hoursWorked("05B") = 10
    Emp(11).hoursWorked("06A") = 10
    Emp(12).hoursWorked("06B") = 10
    Emp(13).hoursWorked("07A") = 10
    Emp(14).hoursWorked("07B") = 10
    Emp(15).hoursWorked("08A") = 10
    Emp(16).hoursWorked("08B") = 10
    Emp(17).hoursWorked("09A") = 10
    Emp(18).hoursWorked("09B") = 10
    Emp(19).hoursWorked("10A") = 10
    Emp(20).hoursWorked("10B") = 10
    Emp(21).hoursWorked("11A") = 10
    Emp(22).hoursWorked("11B") = 10
    Emp(23).hoursWorked("12A") = 10
    Emp(24).hoursWorked("12B") = 10
    
    For Index = 1 To 24
        EC.Add Emp(Index)
    Next
    
    Dim DataRange() As Variant
    
    'Act:
    
    Set wsFTEComboDept = Main.BuildFTESummaryByDepartment(EC)
    
    'Assert:
    
    DataRange = wsFTEComboDept.Range("A2:E25").Value2
    
    For Index = 1 To 24:
        Assert.IsTrue DataRange(Index, 1) = Emp(Index).EmplID
        Assert.IsTrue DataRange(Index, 2) = Emp(Index).Name
        Assert.IsTrue DataRange(Index, 3) = Emp(Index).Department
        Assert.IsTrue DataRange(Index, 4) = Emp(Index).hoursWorked
        Assert.IsTrue DataRange(Index, 5) = Emp(Index).hoursWorked / 198 * 100
    Next Index
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Dim AlertsEnabled As Boolean
    
    ' Ensure that DisplayAlerts is disabled, delete a sheet,
    '   then ensure that DisplayAlerts is set to whatever value it held before this action.
    AlertsEnabled = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wsFTEComboDept.Delete
    Application.DisplayAlerts = AlertsEnabled
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_BuildFTESummaryByDepartment_Data_OneEmployeePerPeriod_MultipleDepartmentsWithSpaces_HourlyOnly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsFTEComboDept As Worksheet
    Dim EC As EmployeeCollection
    Dim Emp(1 To 24) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 24
        Set Emp(Index) = New Employee
    Next
    
    Emp(1).Name = "Jan A"
    Emp(2).Name = "Jan B"
    Emp(3).Name = "Feb A"
    Emp(4).Name = "Feb B"
    Emp(5).Name = "Mar A"
    Emp(6).Name = "Mar B"
    Emp(7).Name = "Apr A"
    Emp(8).Name = "Apr B"
    Emp(9).Name = "May A"
    Emp(10).Name = "May B"
    Emp(11).Name = "Jun A"
    Emp(12).Name = "Jun B"
    Emp(13).Name = "Jul A"
    Emp(14).Name = "Jul B"
    Emp(15).Name = "Aug A"
    Emp(16).Name = "Aug B"
    Emp(17).Name = "Sep A"
    Emp(18).Name = "Sep B"
    Emp(19).Name = "Oct A"
    Emp(20).Name = "Oct B"
    Emp(21).Name = "Nov A"
    Emp(22).Name = "Nov B"
    Emp(23).Name = "Dec A"
    Emp(24).Name = "Dec B"
    
    For Index = 1 To 8
        Emp(Index).EmplID = Index
        Emp(Index).JobCode = "JOB"
        Emp(Index).Department = "Accounting Department"
        Emp(Index).IsHourly = True
    Next
    
    For Index = 9 To 16
        Emp(Index).EmplID = Index
        Emp(Index).JobCode = "JOB"
        Emp(Index).Department = "Nursing Department"
        Emp(Index).IsHourly = True
    Next
    
    For Index = 17 To 24
        Emp(Index).EmplID = Index
        Emp(Index).JobCode = "JOB"
        Emp(Index).Department = "DEPT1"
        Emp(Index).IsHourly = True
    Next
    
    Emp(1).hoursWorked("01A") = 10
    Emp(2).hoursWorked("01B") = 10
    Emp(3).hoursWorked("02A") = 10
    Emp(4).hoursWorked("02B") = 10
    Emp(5).hoursWorked("03A") = 10
    Emp(6).hoursWorked("03B") = 10
    Emp(7).hoursWorked("04A") = 10
    Emp(8).hoursWorked("04B") = 10
    Emp(9).hoursWorked("05A") = 10
    Emp(10).hoursWorked("05B") = 10
    Emp(11).hoursWorked("06A") = 10
    Emp(12).hoursWorked("06B") = 10
    Emp(13).hoursWorked("07A") = 10
    Emp(14).hoursWorked("07B") = 10
    Emp(15).hoursWorked("08A") = 10
    Emp(16).hoursWorked("08B") = 10
    Emp(17).hoursWorked("09A") = 10
    Emp(18).hoursWorked("09B") = 10
    Emp(19).hoursWorked("10A") = 10
    Emp(20).hoursWorked("10B") = 10
    Emp(21).hoursWorked("11A") = 10
    Emp(22).hoursWorked("11B") = 10
    Emp(23).hoursWorked("12A") = 10
    Emp(24).hoursWorked("12B") = 10
    
    For Index = 1 To 24
        EC.Add Emp(Index)
    Next
    
    Dim DataRange() As Variant
    
    'Act:
    
    Set wsFTEComboDept = Main.BuildFTESummaryByDepartment(EC)
    
    'Assert:
    
    DataRange = wsFTEComboDept.Range("A2:E25").Value2
    
    For Index = 1 To 8:
        Assert.IsTrue DataRange(Index, 1) = Emp(Index).EmplID
        Assert.IsTrue DataRange(Index, 2) = Emp(Index).Name
        Assert.IsTrue DataRange(Index, 3) = "Accounting Department"
        Assert.IsTrue DataRange(Index, 4) = Emp(Index).hoursWorked
        Assert.IsTrue DataRange(Index, 5) = Emp(Index).hoursWorked / 198 * 100
    Next Index
    
    For Index = 9 To 16:
        Assert.IsTrue DataRange(Index, 1) = Emp(Index).EmplID
        Assert.IsTrue DataRange(Index, 2) = Emp(Index).Name
        Assert.IsTrue DataRange(Index, 3) = "Nursing Department"
        Assert.IsTrue DataRange(Index, 4) = Emp(Index).hoursWorked
        Assert.IsTrue DataRange(Index, 5) = Emp(Index).hoursWorked / 198 * 100
    Next Index
    
    For Index = 17 To 24:
        Assert.IsTrue DataRange(Index, 1) = Emp(Index).EmplID
        Assert.IsTrue DataRange(Index, 2) = Emp(Index).Name
        Assert.IsTrue DataRange(Index, 3) = "DEPT1"
        Assert.IsTrue DataRange(Index, 4) = Emp(Index).hoursWorked
        Assert.IsTrue DataRange(Index, 5) = Emp(Index).hoursWorked / 198 * 100
    Next Index
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Dim AlertsEnabled As Boolean
    
    ' Ensure that DisplayAlerts is disabled, delete a sheet,
    '   then ensure that DisplayAlerts is set to whatever value it held before this action.
    AlertsEnabled = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wsFTEComboDept.Delete
    Application.DisplayAlerts = AlertsEnabled
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_BuildFTESummaryByDepartment_Headers_OneEmployeePerPeriod_OneDepartment_HourlyOnly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsFTEComboDept As Worksheet
    Dim EC As EmployeeCollection
    Dim Emp(1 To 24) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 24
        Set Emp(Index) = New Employee
    Next
    
    Emp(1).Name = "Jan A"
    Emp(2).Name = "Jan B"
    Emp(3).Name = "Feb A"
    Emp(4).Name = "Feb B"
    Emp(5).Name = "Mar A"
    Emp(6).Name = "Mar B"
    Emp(7).Name = "Apr A"
    Emp(8).Name = "Apr B"
    Emp(9).Name = "May A"
    Emp(10).Name = "May B"
    Emp(11).Name = "Jun A"
    Emp(12).Name = "Jun B"
    Emp(13).Name = "Jul A"
    Emp(14).Name = "Jul B"
    Emp(15).Name = "Aug A"
    Emp(16).Name = "Aug B"
    Emp(17).Name = "Sep A"
    Emp(18).Name = "Sep B"
    Emp(19).Name = "Oct A"
    Emp(20).Name = "Oct B"
    Emp(21).Name = "Nov A"
    Emp(22).Name = "Nov B"
    Emp(23).Name = "Dec A"
    Emp(24).Name = "Dec B"
    
    For Index = 1 To 24
        Emp(Index).EmplID = Index
        Emp(Index).JobCode = "JOB"
        Emp(Index).Department = "DEPRT"
        Emp(Index).IsHourly = True
    Next
    
    Emp(1).hoursWorked("01A") = 10
    Emp(2).hoursWorked("01B") = 10
    Emp(3).hoursWorked("02A") = 10
    Emp(4).hoursWorked("02B") = 10
    Emp(5).hoursWorked("03A") = 10
    Emp(6).hoursWorked("03B") = 10
    Emp(7).hoursWorked("04A") = 10
    Emp(8).hoursWorked("04B") = 10
    Emp(9).hoursWorked("05A") = 10
    Emp(10).hoursWorked("05B") = 10
    Emp(11).hoursWorked("06A") = 10
    Emp(12).hoursWorked("06B") = 10
    Emp(13).hoursWorked("07A") = 10
    Emp(14).hoursWorked("07B") = 10
    Emp(15).hoursWorked("08A") = 10
    Emp(16).hoursWorked("08B") = 10
    Emp(17).hoursWorked("09A") = 10
    Emp(18).hoursWorked("09B") = 10
    Emp(19).hoursWorked("10A") = 10
    Emp(20).hoursWorked("10B") = 10
    Emp(21).hoursWorked("11A") = 10
    Emp(22).hoursWorked("11B") = 10
    Emp(23).hoursWorked("12A") = 10
    Emp(24).hoursWorked("12B") = 10
    
    For Index = 1 To 24
        EC.Add Emp(Index)
    Next
    
    Dim HeadersRange() As Variant
    
    'Act:
    Set wsFTEComboDept = Main.BuildFTESummaryByDepartment(EC)
    
    'Assert:
    HeadersRange = wsFTEComboDept.Range("A1:E1").Value2
    
    Assert.IsTrue HeadersRange(1, 1) = "Empl ID"
    Assert.IsTrue HeadersRange(1, 2) = "Name (LN, FN)"
    Assert.IsTrue HeadersRange(1, 3) = "Department"
    Assert.IsTrue HeadersRange(1, 4) = "Hours"
    Assert.IsTrue HeadersRange(1, 5) = "FTE%"
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Dim AlertsEnabled As Boolean
    
    ' Ensure that DisplayAlerts is disabled, delete a sheet,
    '   then ensure that DisplayAlerts is set to whatever value it held before this action.
    AlertsEnabled = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wsFTEComboDept.Delete
    Application.DisplayAlerts = AlertsEnabled
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_BuildFTESummaryByJobCode_Data_OneEmployeePerPeriod_OneDepartment_HourlyOnly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsFTEComboJC As Worksheet
    Dim EC As EmployeeCollection
    Dim Emp(1 To 24) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 24
        Set Emp(Index) = New Employee
    Next
    
    Emp(1).Name = "Jan A"
    Emp(2).Name = "Jan B"
    Emp(3).Name = "Feb A"
    Emp(4).Name = "Feb B"
    Emp(5).Name = "Mar A"
    Emp(6).Name = "Mar B"
    Emp(7).Name = "Apr A"
    Emp(8).Name = "Apr B"
    Emp(9).Name = "May A"
    Emp(10).Name = "May B"
    Emp(11).Name = "Jun A"
    Emp(12).Name = "Jun B"
    Emp(13).Name = "Jul A"
    Emp(14).Name = "Jul B"
    Emp(15).Name = "Aug A"
    Emp(16).Name = "Aug B"
    Emp(17).Name = "Sep A"
    Emp(18).Name = "Sep B"
    Emp(19).Name = "Oct A"
    Emp(20).Name = "Oct B"
    Emp(21).Name = "Nov A"
    Emp(22).Name = "Nov B"
    Emp(23).Name = "Dec A"
    Emp(24).Name = "Dec B"
    
    For Index = 1 To 24
        Emp(Index).EmplID = Index
        Emp(Index).JobCode = "JOB"
        Emp(Index).Department = "DEPRT"
        Emp(Index).IsHourly = True
    Next
    
    Emp(1).hoursWorked("01A") = 10
    Emp(2).hoursWorked("01B") = 10
    Emp(3).hoursWorked("02A") = 10
    Emp(4).hoursWorked("02B") = 10
    Emp(5).hoursWorked("03A") = 10
    Emp(6).hoursWorked("03B") = 10
    Emp(7).hoursWorked("04A") = 10
    Emp(8).hoursWorked("04B") = 10
    Emp(9).hoursWorked("05A") = 10
    Emp(10).hoursWorked("05B") = 10
    Emp(11).hoursWorked("06A") = 10
    Emp(12).hoursWorked("06B") = 10
    Emp(13).hoursWorked("07A") = 10
    Emp(14).hoursWorked("07B") = 10
    Emp(15).hoursWorked("08A") = 10
    Emp(16).hoursWorked("08B") = 10
    Emp(17).hoursWorked("09A") = 10
    Emp(18).hoursWorked("09B") = 10
    Emp(19).hoursWorked("10A") = 10
    Emp(20).hoursWorked("10B") = 10
    Emp(21).hoursWorked("11A") = 10
    Emp(22).hoursWorked("11B") = 10
    Emp(23).hoursWorked("12A") = 10
    Emp(24).hoursWorked("12B") = 10
    
    For Index = 1 To 24
        EC.Add Emp(Index)
    Next
    
    Dim DataRange() As Variant
    
    'Act:
    
    Set wsFTEComboJC = Main.BuildFTESummaryByJobCode(EC)
    
    'Assert:
    
    DataRange = wsFTEComboJC.Range("A2:E25").Value2
    
    For Index = 1 To 24:
        Assert.IsTrue DataRange(Index, 1) = Emp(Index).EmplID
        Assert.IsTrue DataRange(Index, 2) = Emp(Index).Name
        Assert.IsTrue DataRange(Index, 3) = Emp(Index).JobCode
        Assert.IsTrue DataRange(Index, 4) = Emp(Index).hoursWorked
        Assert.IsTrue DataRange(Index, 5) = Emp(Index).hoursWorked / 198 * 100
    Next Index
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Dim AlertsEnabled As Boolean
    
    ' Ensure that DisplayAlerts is disabled, delete a sheet,
    '   then ensure that DisplayAlerts is set to whatever value it held before this action.
    AlertsEnabled = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wsFTEComboJC.Delete
    Application.DisplayAlerts = AlertsEnabled
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_BuildFTESummaryByJobCode_Headers_OneEmployeePerPeriod_OneDepartment_HourlyOnly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsFTEComboJC As Worksheet
    Dim EC As EmployeeCollection
    Dim Emp(1 To 24) As Employee
    Dim Index As Long
    
    Set EC = New EmployeeCollection
    
    For Index = 1 To 24
        Set Emp(Index) = New Employee
    Next
    
    Emp(1).Name = "Jan A"
    Emp(2).Name = "Jan B"
    Emp(3).Name = "Feb A"
    Emp(4).Name = "Feb B"
    Emp(5).Name = "Mar A"
    Emp(6).Name = "Mar B"
    Emp(7).Name = "Apr A"
    Emp(8).Name = "Apr B"
    Emp(9).Name = "May A"
    Emp(10).Name = "May B"
    Emp(11).Name = "Jun A"
    Emp(12).Name = "Jun B"
    Emp(13).Name = "Jul A"
    Emp(14).Name = "Jul B"
    Emp(15).Name = "Aug A"
    Emp(16).Name = "Aug B"
    Emp(17).Name = "Sep A"
    Emp(18).Name = "Sep B"
    Emp(19).Name = "Oct A"
    Emp(20).Name = "Oct B"
    Emp(21).Name = "Nov A"
    Emp(22).Name = "Nov B"
    Emp(23).Name = "Dec A"
    Emp(24).Name = "Dec B"
    
    For Index = 1 To 24
        Emp(Index).EmplID = Index
        Emp(Index).JobCode = "JOB"
        Emp(Index).Department = "DEPRT"
        Emp(Index).IsHourly = True
    Next
    
    Emp(1).hoursWorked("01A") = 10
    Emp(2).hoursWorked("01B") = 10
    Emp(3).hoursWorked("02A") = 10
    Emp(4).hoursWorked("02B") = 10
    Emp(5).hoursWorked("03A") = 10
    Emp(6).hoursWorked("03B") = 10
    Emp(7).hoursWorked("04A") = 10
    Emp(8).hoursWorked("04B") = 10
    Emp(9).hoursWorked("05A") = 10
    Emp(10).hoursWorked("05B") = 10
    Emp(11).hoursWorked("06A") = 10
    Emp(12).hoursWorked("06B") = 10
    Emp(13).hoursWorked("07A") = 10
    Emp(14).hoursWorked("07B") = 10
    Emp(15).hoursWorked("08A") = 10
    Emp(16).hoursWorked("08B") = 10
    Emp(17).hoursWorked("09A") = 10
    Emp(18).hoursWorked("09B") = 10
    Emp(19).hoursWorked("10A") = 10
    Emp(20).hoursWorked("10B") = 10
    Emp(21).hoursWorked("11A") = 10
    Emp(22).hoursWorked("11B") = 10
    Emp(23).hoursWorked("12A") = 10
    Emp(24).hoursWorked("12B") = 10
    
    For Index = 1 To 24
        EC.Add Emp(Index)
    Next
    
    Dim HeadersRange() As Variant
    
    'Act:
    Set wsFTEComboJC = Main.BuildFTESummaryByJobCode(EC)
    
    'Assert:
    HeadersRange = wsFTEComboJC.Range("A1:E1").Value2
    
    Assert.IsTrue HeadersRange(1, 1) = "Empl ID"
    Assert.IsTrue HeadersRange(1, 2) = "Name (LN, FN)"
    Assert.IsTrue HeadersRange(1, 3) = "Job Code"
    Assert.IsTrue HeadersRange(1, 4) = "Hours"
    Assert.IsTrue HeadersRange(1, 5) = "FTE%"
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Dim AlertsEnabled As Boolean
    
    ' Ensure that DisplayAlerts is disabled, delete a sheet,
    '   then ensure that DisplayAlerts is set to whatever value it held before this action.
    AlertsEnabled = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wsFTEComboJC.Delete
    Application.DisplayAlerts = AlertsEnabled
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Uncategorized")
Private Sub Test_InJobCodeList()
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    
    'Assert:
    Assert.IsTrue Main.inJobCodeList("PTF")
    Assert.IsTrue Main.inJobCodeList("PTH")
    Assert.IsTrue Main.inJobCodeList("SUB")
    Assert.IsFalse Main.inJobCodeList("NOT")
    Assert.IsFalse Main.inJobCodeList("TWO")
    Assert.IsFalse Main.inJobCodeList("THISISWRONG")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_BuildFTESummaryByJobCode_Data_PeriodOTH()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsFTEComboJC As Worksheet
    Dim EC As EmployeeCollection
    Dim Emp(1 To 9) As Employee
    Dim Index As Long
    Dim DataRange() As Variant
    Dim E1_Name As String
    Dim E2_Name As String
    Dim E3_Name As String
    Dim E1_EmplID As String
    Dim E2_EmplID As String
    Dim E3_EmplID As String
    
    Set EC = New EmployeeCollection
    E1_Name = "John Smith"
    E2_Name = "Jane Doe"
    E3_Name = "Jim Bob"
    
    E1_EmplID = "1"
    E2_EmplID = "2"
    E3_EmplID = "3"
    
    For Index = 1 To 9
        Set Emp(Index) = New Employee
        If Index <= 3 Then
            Emp(Index).Name = E1_Name
            Emp(Index).EmplID = E1_EmplID
        ElseIf Index <= 6 Then
            Emp(Index).Name = E2_Name
            Emp(Index).EmplID = E2_EmplID
        Else
            Emp(Index).Name = E3_Name
            Emp(Index).EmplID = E3_EmplID
        End If
        
        Emp(Index).hoursWorked("OTH") = 10
    Next Index
    
    Emp(1).JobCode = "JC1"
    Emp(2).JobCode = "JC1"
    Emp(3).JobCode = "JC1"
    
    Emp(4).JobCode = "JC1"
    Emp(5).JobCode = "JC1"
    Emp(6).JobCode = "JC2"
    
    Emp(7).JobCode = "JC1"
    Emp(8).JobCode = "JC2"
    Emp(9).JobCode = "JC3"
    
    For Index = 1 To 9:
        EC.Add Emp(Index)
    Next Index
    
    'Act:
    Set wsFTEComboJC = Main.BuildFTESummaryByJobCode(EC)
    
    'Assert:
    DataRange = wsFTEComboJC.Range("A2:E25").Value2
    
    Assert.IsTrue DataRange(1, 1) = E1_EmplID
    Assert.IsTrue DataRange(1, 2) = E1_Name
    Assert.IsTrue DataRange(1, 3) = "JC1"
    Assert.IsTrue DataRange(1, 4) = 30
    
    Assert.IsTrue DataRange(2, 1) = E2_EmplID
    Assert.IsTrue DataRange(2, 2) = E2_Name
    Assert.IsTrue DataRange(2, 3) = "JC1"
    Assert.IsTrue DataRange(2, 4) = 20
    
    Assert.IsTrue DataRange(3, 1) = E2_EmplID
    Assert.IsTrue DataRange(3, 2) = E2_Name
    Assert.IsTrue DataRange(3, 3) = "JC2"
    Assert.IsTrue DataRange(3, 4) = 10
    
    Assert.IsTrue DataRange(4, 1) = E3_EmplID
    Assert.IsTrue DataRange(4, 2) = E3_Name
    Assert.IsTrue DataRange(4, 3) = "JC1"
    Assert.IsTrue DataRange(4, 4) = 10
    
    Assert.IsTrue DataRange(5, 1) = E3_EmplID
    Assert.IsTrue DataRange(5, 2) = E3_Name
    Assert.IsTrue DataRange(5, 3) = "JC2"
    Assert.IsTrue DataRange(5, 4) = 10
    
    Assert.IsTrue DataRange(6, 1) = E3_EmplID
    Assert.IsTrue DataRange(6, 2) = E3_Name
    Assert.IsTrue DataRange(6, 3) = "JC3"
    Assert.IsTrue DataRange(6, 4) = 10
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Dim AlertsEnabled As Boolean
    
    ' Ensure that DisplayAlerts is disabled, delete a sheet,
    '   then ensure that DisplayAlerts is set to whatever value it held before this action.
    AlertsEnabled = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wsFTEComboJC.Delete
    Application.DisplayAlerts = AlertsEnabled
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
