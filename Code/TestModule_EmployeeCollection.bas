Attribute VB_Name = "TestModule_EmployeeCollection"
'@IgnoreModule EmptyMethod
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
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'TODO Test EmployeeCollection Initialization
'TODO Test EmployeeCollection Add
'TODO Test EmployeeCollection Remove
'TODO Test EmployeeCollection Item
'TODO Test EmployeeCollection Count
'TODO Test EmployeeCollection Employees (Get)

'@TestMethod("Getters")
Private Sub Test_EmployeeEmployeesWithJobCode()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC As EmployeeCollection
    Set EC = New EmployeeCollection
    
    Dim JohnDoe As Employee
    Dim JaneDeer As Employee
    Dim DonaldDuck As Employee
    Dim DaisyDuck As Employee
    Set JohnDoe = New Employee
    Set JaneDeer = New Employee
    Set DonaldDuck = New Employee
    Set DaisyDuck = New Employee
    
    JohnDoe.Name = "John Doe"
    JohnDoe.EmplID = "000000000"
    JohnDoe.Department = "00000"
    JohnDoe.JobCode = "TARGET"

    JaneDeer.Name = "Jane Deer"
    JaneDeer.EmplID = "111111111"
    JaneDeer.Department = "00000"
    JaneDeer.JobCode = "NOT_TARGET"
    
    DonaldDuck.Name = "Donald Duck"
    DonaldDuck.EmplID = "222222222"
    DonaldDuck.Department = "00000"
    DonaldDuck.JobCode = "TARGET"
    
    DaisyDuck.Name = "DaisyDuck"
    DaisyDuck.EmplID = "333333333"
    DaisyDuck.Department = "00000"
    DaisyDuck.JobCode = "NOT_TARGET"
    
    EC.Add JohnDoe, "TestModule_EmployeeCollection.Test_EmployeeEmployeesWithJobCode.Employees"
    EC.Add JaneDeer, "TestModule_EmployeeCollection.Test_EmployeeEmployeesWithJobCode.Employees"
    EC.Add DonaldDuck, "TestModule_EmployeeCollection.Test_EmployeeEmployeesWithJobCode.Employees"
    EC.Add DaisyDuck, "TestModule_EmployeeCollection.Test_EmployeeEmployeesWithJobCode.Employees"
    
    'Act:
    Dim EC_Subset As EmployeeCollection
    Set EC_Subset = EC.EmployeesWithJobCode("TARGET")
    
    Dim HasKeyJohnDoe As Boolean
    Dim HasKeyJaneDeer As Boolean
    Dim HasKeyDonaldDuck As Boolean
    Dim HasKeyDaisyDuck As Boolean
    
    HasKeyJohnDoe = EC_Subset.HasKey(JohnDoe.eKey)
    HasKeyJaneDeer = EC_Subset.HasKey(JaneDeer.eKey)
    HasKeyDonaldDuck = EC_Subset.HasKey(DonaldDuck.eKey)
    HasKeyDaisyDuck = EC_Subset.HasKey(DaisyDuck.eKey)
    
    'Assert:
    Assert.IsTrue HasKeyJohnDoe
    Assert.IsFalse HasKeyJaneDeer
    Assert.IsTrue HasKeyDonaldDuck
    Assert.IsFalse HasKeyDaisyDuck

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Getters")
Private Sub Test_EmployeeEmployeesWithDepartment()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC As EmployeeCollection
    Set EC = New EmployeeCollection
    
    Dim JohnDoe As Employee
    Dim JaneDeer As Employee
    Dim DonaldDuck As Employee
    Dim DaisyDuck As Employee
    Set JohnDoe = New Employee
    Set JaneDeer = New Employee
    Set DonaldDuck = New Employee
    Set DaisyDuck = New Employee
    
    JohnDoe.Name = "John Doe"
    JohnDoe.EmplID = "000000000"
    JohnDoe.Department = "TARGET"
    JohnDoe.JobCode = "AAA"

    JaneDeer.Name = "Jane Deer"
    JaneDeer.EmplID = "111111111"
    JaneDeer.Department = "NOT_TARGET"
    JaneDeer.JobCode = "AAA"
    
    DonaldDuck.Name = "Donald Duck"
    DonaldDuck.EmplID = "222222222"
    DonaldDuck.Department = "TARGET"
    DonaldDuck.JobCode = "AAA"
    
    DaisyDuck.Name = "DaisyDuck"
    DaisyDuck.EmplID = "333333333"
    DaisyDuck.Department = "NOT_TARGET"
    DaisyDuck.JobCode = "AAA"
    
    EC.Add JohnDoe, "TestModule_EmployeeCollection.Test_EmployeeEmployeesWithDepartment.Employees"
    EC.Add JaneDeer, "TestModule_EmployeeCollection.Test_EmployeeEmployeesWithDepartment.Employees"
    EC.Add DonaldDuck, "TestModule_EmployeeCollection.Test_EmployeeEmployeesWithDepartment.Employees"
    EC.Add DaisyDuck, "TestModule_EmployeeCollection.Test_EmployeeEmployeesWithDepartment.Employees"
    
    'Act:
    Dim EC_Subset As EmployeeCollection
    Set EC_Subset = EC.EmployeesWithDepartment("TARGET")
    
    Dim HasKeyJohnDoe As Boolean
    Dim HasKeyJaneDeer As Boolean
    Dim HasKeyDonaldDuck As Boolean
    Dim HasKeyDaisyDuck As Boolean
    
    HasKeyJohnDoe = EC_Subset.HasKey(JohnDoe.eKey)
    HasKeyJaneDeer = EC_Subset.HasKey(JaneDeer.eKey)
    HasKeyDonaldDuck = EC_Subset.HasKey(DonaldDuck.eKey)
    HasKeyDaisyDuck = EC_Subset.HasKey(DaisyDuck.eKey)
    
    'Assert:
    Assert.IsTrue HasKeyJohnDoe
    Assert.IsFalse HasKeyJaneDeer
    Assert.IsTrue HasKeyDonaldDuck
    Assert.IsFalse HasKeyDaisyDuck

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'TODO Test EmployeeCollection GetItemsWithDepartment
'TODO Test EmployeeCollection GetItemsWithEmplID

'@TestMethod("Booleans")
Private Sub Test_EmployeeCollectionHasKeyTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC As EmployeeCollection
    Set EC = New EmployeeCollection
    
    Dim Emp As Employee
    Set Emp = New Employee
    EC.Add Emp
    
    Dim HasKey As Boolean
    
    'Act:
    HasKey = EC.HasKey(Emp.eKey)
    
    'Assert:
    Assert.IsTrue HasKey

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Booleans")
Private Sub Test_EmployeeCollectionHasKeyFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC As EmployeeCollection
    Set EC = New EmployeeCollection
    
    Dim Emp As Employee
    Set Emp = New Employee
    
    Dim HasKey As Boolean
    
    'Act:
    HasKey = EC.HasKey(Emp.eKey)
    
    'Assert:
    Assert.IsFalse HasKey

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_EmployeeCollectionCombine()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC_A As EmployeeCollection
    Dim EC_B As EmployeeCollection
    Dim EC_Combined As EmployeeCollection
    
    Set EC_A = New EmployeeCollection
    Set EC_B = New EmployeeCollection
    
    Dim John As Employee
    Dim Jane As Employee
    Dim Donald As Employee
    Dim Daisy As Employee
    
    Set John = New Employee
    Set Jane = New Employee
    Set Donald = New Employee
    Set Daisy = New Employee
    
    John.EmplID = "000000000"
    Jane.EmplID = "111111111"
    Donald.EmplID = "222222222"
    Daisy.EmplID = "333333333"
    
    John.Name = "John"
    Jane.Name = "Jane"
    Donald.Name = "Donald"
    Daisy.Name = "Daisy"
    
    John.Department = "AAAAA"
    Jane.Department = "BBBBB"
    Donald.Department = "CCCCC"
    Daisy.Department = "DDDDD"
    
    John.JobCode = "!!!"
    Jane.JobCode = "@@@"
    Donald.JobCode = "###"
    Daisy.JobCode = "$$$"
    
    EC_A.Add John, "EmployeesA"
    EC_A.Add Jane, "EmployeesA"
    EC_B.Add Donald, "EmployeesB"
    EC_B.Add Daisy, "EmployeesB"
    
    'Act:
    Set EC_Combined = EC_A.Combine(EC_B)
    
    'Assert:
    Assert.IsTrue EC_Combined.HasKey("000000000*AAAAA*!!!")
    Assert.IsTrue EC_Combined.HasKey("111111111*BBBBB*@@@")
    Assert.IsTrue EC_Combined.HasKey("222222222*CCCCC*###")
    Assert.IsTrue EC_Combined.HasKey("333333333*DDDDD*$$$")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'TODO Test EmployeeCollection MergeDuplicateEmployeesOnDepartment
'TODO Test EmployeeCollection MergeDuplicateEmployeesOnJobCode
'TODO Test EmployeeCollection PruneEmployeesToJobCodeList

'@TestMethod("Uncategorized")
Private Sub Test_EmployeesWithJobCodeSameJobCodeSameDepartment()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
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
    
    Set EC_Merged = EC.EmployeesWithJobCode("JOB")
    
    'Assert:
    
    For Index = 1 To 24
        Assert.IsTrue EC_Merged.Item(Index).hoursWorked = 10
    Next

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_MergeEmployeesWithJobCodeSameJobCodeSameDepartment()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC As EmployeeCollection
    Dim EC_withJobCode As EmployeeCollection
    Dim EC_Merged As EmployeeCollection
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
    
    Set EC_withJobCode = EC.EmployeesWithJobCode("JOB")
    Set EC_Merged = EC_withJobCode.MergeDuplicateEmployeesOnJobCode()
    
    'Assert:
    
    For Index = 1 To 24
        Assert.IsTrue EC_Merged.Item(Index).hoursWorked = 10
        Assert.IsTrue EC_Merged.Item(Index).Name = Emp(Index).Name
    Next

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_Copy_RunWithoutErrors()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC As EmployeeCollection
    '@Ignore VariableNotUsed
    Dim ECCopy As EmployeeCollection
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
    '@Ignore AssignmentNotUsed
    Set ECCopy = EC.Copy()
    
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
Private Sub Test_Copy_AssertAreNotSame()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC As EmployeeCollection
    Dim ECCopy As EmployeeCollection
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
    Set ECCopy = EC.Copy()
    
    'Assert:
    Assert.AreNotSame EC, ECCopy

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_Combine_UniqueEmps()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Index As Long
    
    '@Ignore UseMeaningfulName
    Dim EC1 As EmployeeCollection
    '@Ignore UseMeaningfulName
    Dim EC2 As EmployeeCollection
    Dim EC_Combo As EmployeeCollection
    
    Set EC1 = New EmployeeCollection
    Set EC2 = New EmployeeCollection
    
    Dim Emp As Employee
    
    For Index = 0 To 9
        Set Emp = New Employee
        
        Emp.EmplID = Str$(Index)
        Emp.Name = ("Employee Number " & Str$(Index))
        Emp.JobCode = "JOB"
        Emp.Department = "DEPT1"
        Emp.IsHourly = True
        Emp.Source = "Trust Me, Bro"
        
        EC1.Add Emp
    Next Index
    
    For Index = 10 To 19
        Set Emp = New Employee
        
        Emp.EmplID = Str$(Index)
        Emp.Name = ("Employee Number " & Str$(Index))
        Emp.JobCode = "JOB"
        Emp.Department = "DEPT2"
        Emp.IsHourly = True
        Emp.Source = "Trust Me, Bro"
        
        EC2.Add Emp
    Next Index
    
    'Act:
    Set EC_Combo = EC1.Combine(EC2)
    
    'Assert:
    For Index = 0 To 19
        Assert.IsTrue EC_Combo.Item(Index + 1).EmplID = Index
        Assert.IsTrue EC_Combo.Item(Index + 1).Name = ("Employee Number " & Str$(Index))
        Assert.IsTrue EC_Combo.Item(Index + 1).JobCode = "JOB"
        Assert.IsTrue EC_Combo.Item(Index + 1).IsHourly = True
        Assert.IsTrue EC_Combo.Item(Index + 1).Source = "Trust Me, Bro"
    Next Index
    
    For Index = 0 To 9
        Assert.IsTrue EC_Combo.Item(Index + 1).Department = "DEPT1"
    Next Index
    
    For Index = 10 To 19
        Assert.IsTrue EC_Combo.Item(Index + 1).Department = "DEPT2"
    Next Index

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_GetEmployeeCollectionFromWorksheet_Hourly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim EC As EmployeeCollection
    
    Dim WorkbookFileDir As String
    Dim WorkbookFileName As String
    WorkbookFileDir = ThisWorkbook.Path & "/TestData/"
    WorkbookFileName = "Test Workbook 1.xlsx"
    
    Workbooks.Open fileName:=WorkbookFileDir & WorkbookFileName, ReadOnly:=True
    Set wb = Workbooks.Item(WorkbookFileName)
    Set ws = wb.Sheets.Item("Hourly")
    
    Set EC = New EmployeeCollection
    
    'Act:
    Set EC = EC.CreateEmployeeCollectionFromWorksheet(ws, True, Source:=wb.Path)
    
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
        Assert.IsTrue EC.Item(Index + 1).hoursWorked() = 10
        Assert.IsTrue EC.Item(Index + 1).hoursWorked(PayPeriodString) = 10
        Assert.IsTrue EC.Item(Index + 1).Department = Format$(Index + 1, "00000")
    Next Index

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_EmployeeCollection_ReadFromWorkbook_Hourly_CorrectCount()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC As EmployeeCollection
    Set EC = New EmployeeCollection
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Dim WorkbookFileDir As String
    Dim WorkbookFileName As String
    
    WorkbookFileDir = ThisWorkbook.Path & "/TestData/"
    WorkbookFileName = "Test Workbook 1.xlsx"
    
    'Act:
    Workbooks.Open fileName:=WorkbookFileDir & WorkbookFileName, ReadOnly:=True
    Set wb = Workbooks.Item(WorkbookFileName)
    Set ws = wb.Sheets.Item("Hourly")
    Set EC = EC.CreateEmployeeCollectionFromWorksheet(ws, True, Source:=WorkbookFileName)
    wb.Close
    
    'Assert:
    Assert.IsTrue EC.Count = 24

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_EmployeeCollection_ReadFromWorkbook_Appointed_CorrectCount()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC As EmployeeCollection
    Set EC = New EmployeeCollection
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Dim WorkbookFileDir As String
    Dim WorkbookFileName As String
    
    WorkbookFileDir = ThisWorkbook.Path & "/TestData/"
    WorkbookFileName = "Test Workbook 2.xlsx"
    
    'Act:
    Workbooks.Open fileName:=WorkbookFileDir & WorkbookFileName, ReadOnly:=True
    Set wb = Workbooks.Item(WorkbookFileName)
    Set ws = wb.Sheets.Item("Appointed")
    Set EC = EC.CreateEmployeeCollectionFromWorksheet(ws, True, Source:=WorkbookFileName)
    wb.Close
    
    'Assert:
    Assert.IsTrue EC.Count = 24

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_EmployeeCollectionCombine_MultipleWorkbooks_AppointedAndHourly_CorrectCount()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC_Appointed As EmployeeCollection
    Dim EC_Hourly As EmployeeCollection
    Dim EC_Combined As EmployeeCollection
    
    Set EC_Appointed = New EmployeeCollection
    Set EC_Hourly = New EmployeeCollection
    
    Dim wb_Appointed As Workbook
    Dim wb_Hourly As Workbook
    Dim ws_Appointed As Worksheet
    Dim ws_Hourly As Worksheet
    
    Dim WorkbookFileDir As String
    Dim WorkbookFileName_Appointed As String
    Dim WorkbookFileName_Hourly As String
    
    WorkbookFileDir = ThisWorkbook.Path & "/TestData/"
    WorkbookFileName_Appointed = "Test Workbook - Multiple Workbooks - Appointed and Hourly - File 2.xlsx"
    WorkbookFileName_Hourly = "Test Workbook - Multiple Workbooks - Appointed and Hourly - File 1.xlsx"
    
    Workbooks.Open fileName:=WorkbookFileDir & WorkbookFileName_Appointed, ReadOnly:=True
    Set wb_Appointed = Workbooks.Item(WorkbookFileName_Appointed)
    Set ws_Appointed = wb_Appointed.Sheets.Item("Appointed")
    Set EC_Appointed = EC_Appointed.CreateEmployeeCollectionFromWorksheet(ws_Appointed, False, Source:=WorkbookFileName_Appointed)
    wb_Appointed.Close
    
    Workbooks.Open fileName:=WorkbookFileDir & WorkbookFileName_Hourly, ReadOnly:=True
    Set wb_Hourly = Workbooks.Item(WorkbookFileName_Hourly)
    Set ws_Hourly = wb_Hourly.Sheets.Item("Hourly")
    Set EC_Hourly = EC_Hourly.CreateEmployeeCollectionFromWorksheet(ws_Hourly, True, Source:=WorkbookFileName_Hourly)
    wb_Hourly.Close
    
    'Act:
    Set EC_Combined = EC_Appointed.Combine(EC_Hourly)
    
    'Assert:
    Debug.Print EC_Appointed.Count
    Debug.Print EC_Hourly.Count
    Debug.Print EC_Combined.Count
    Assert.IsTrue EC_Combined.Count = 48

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_EmployeeCollectionCombine_SingleWorkbook_AppointedAndHourly_CorrectCount()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC_Appointed As EmployeeCollection
    Dim EC_Hourly As EmployeeCollection
    Dim EC_Combined As EmployeeCollection
    
    Set EC_Appointed = New EmployeeCollection
    Set EC_Hourly = New EmployeeCollection
    
    Dim wb As Workbook
    Dim ws_Appointed As Worksheet
    Dim ws_Hourly As Worksheet
    
    Dim WorkbookFileDir As String
    Dim WorkbookFileName As String
    
    WorkbookFileDir = ThisWorkbook.Path & "/TestData/"
    WorkbookFileName = "Test Workbook - Appointed and Hourly.xlsx"
    
    Workbooks.Open fileName:=WorkbookFileDir & WorkbookFileName, ReadOnly:=True
    Set wb = Workbooks.Item(WorkbookFileName)
    Set ws_Appointed = wb.Sheets.Item("Appointed")
    Set ws_Hourly = wb.Sheets.Item("Hourly")
    Set EC_Appointed = EC_Appointed.CreateEmployeeCollectionFromWorksheet(ws_Appointed, False, Source:=WorkbookFileName)
    Set EC_Hourly = EC_Hourly.CreateEmployeeCollectionFromWorksheet(ws_Hourly, True, Source:=WorkbookFileName)
    wb.Close
    
    'Act:
    Set EC_Combined = EC_Appointed.Combine(EC_Hourly)
    
    'Assert:
    Assert.IsTrue EC_Combined.Count = 48

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_EmployeeCollectionCombine_EmptyCollections_CorrectCount()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EC1 As EmployeeCollection
    Dim EC2 As EmployeeCollection
    Dim EC_Combined As EmployeeCollection
    
    Set EC1 = New EmployeeCollection
    Set EC2 = New EmployeeCollection
    
    'Act:
    Set EC_Combined = EC1.Combine(EC2)
    
    'Assert:
    Assert.IsTrue EC_Combined.Count = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_EmployeeCollectionCombine_BaseCollectionEmpty_CorrectCount()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Index As Long
    Dim EC1 As EmployeeCollection
    Dim EC2 As EmployeeCollection
    Dim EC_Combined As EmployeeCollection
    Dim Emp(5) As Employee
    
    Set EC1 = New EmployeeCollection
    Set EC2 = New EmployeeCollection
    
    For Index = 0 To 4
        Set Emp(Index) = New Employee
        Emp(Index).EmplID = Str(Index)
        Emp(Index).Name = "Employee Number " & Str(Index)
        EC2.Add Emp(Index)
    Next Index
    
    'Act:
    Set EC_Combined = EC1.Combine(EC2)
    
    'Assert:
    Assert.IsTrue EC_Combined.Count = 5

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_EmployeeCollectionCombine_BaseCollectionEmpty_CorrectEmplIDs()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Index As Long
    Dim EC1 As EmployeeCollection
    Dim EC2 As EmployeeCollection
    Dim EC_Combined As EmployeeCollection
    Dim Emp(5) As Employee
    
    Set EC1 = New EmployeeCollection
    Set EC2 = New EmployeeCollection
    
    For Index = 0 To 4
        Set Emp(Index) = New Employee
        Emp(Index).EmplID = Index
        Emp(Index).Name = "Employee Number " & Index
        EC2.Add Emp(Index)
    Next Index
    
    'Act:
    Set EC_Combined = EC1.Combine(EC2)
    
    'Assert:
    Assert.IsTrue EC_Combined.Item(1).EmplID = "0"
    Assert.IsTrue EC_Combined.Item(2).EmplID = "1"
    Assert.IsTrue EC_Combined.Item(3).EmplID = "2"
    Assert.IsTrue EC_Combined.Item(4).EmplID = "3"
    Assert.IsTrue EC_Combined.Item(5).EmplID = "4"
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_EmployeeCollectionCombine_SecondCollectionEmpty_CorrectCount()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Index As Long
    Dim EC1 As EmployeeCollection
    Dim EC2 As EmployeeCollection
    Dim EC_Combined As EmployeeCollection
    Dim Emp(5) As Employee
    
    Set EC1 = New EmployeeCollection
    Set EC2 = New EmployeeCollection
    
    For Index = 0 To 4
        Set Emp(Index) = New Employee
        Emp(Index).EmplID = Str(Index)
        Emp(Index).Name = "Employee Number " & Str(Index)
        EC1.Add Emp(Index)
    Next Index
    
    'Act:
    Set EC_Combined = EC1.Combine(EC2)
    
    'Assert:
    Assert.IsTrue EC_Combined.Count = 5

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_EmployeeCollectionCombine_SecondCollectionEmpty_CorrectEmplIDs()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Index As Long
    Dim EC1 As EmployeeCollection
    Dim EC2 As EmployeeCollection
    Dim EC_Combined As EmployeeCollection
    Dim Emp(5) As Employee
    
    Set EC1 = New EmployeeCollection
    Set EC2 = New EmployeeCollection
    
    For Index = 0 To 4
        Set Emp(Index) = New Employee
        Emp(Index).EmplID = Index
        Emp(Index).Name = "Employee Number " & Index
        EC1.Add Emp(Index)
    Next Index
    
    'Act:
    Set EC_Combined = EC1.Combine(EC2)
    
    'Assert:
    Assert.IsTrue EC_Combined.Item(1).EmplID = "0"
    Assert.IsTrue EC_Combined.Item(2).EmplID = "1"
    Assert.IsTrue EC_Combined.Item(3).EmplID = "2"
    Assert.IsTrue EC_Combined.Item(4).EmplID = "3"
    Assert.IsTrue EC_Combined.Item(5).EmplID = "4"
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_EmployeeCollectionCombine_FiveAndFive_CorrectCount()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Index As Long
    Dim EC1 As EmployeeCollection
    Dim EC2 As EmployeeCollection
    Dim EC_Combined As EmployeeCollection
    Dim Emp(10) As Employee
    
    Set EC1 = New EmployeeCollection
    Set EC2 = New EmployeeCollection
    
    For Index = 0 To 4
        Set Emp(Index) = New Employee
        Emp(Index).EmplID = Str(Index)
        Emp(Index).Name = "Employee Number " & Index
        EC1.Add Emp(Index)
    Next Index
    
    For Index = 5 To 9
        Set Emp(Index) = New Employee
        Emp(Index).EmplID = Str(Index)
        Emp(Index).Name = "Employee Number " & Index
        EC2.Add Emp(Index)
    Next Index
    
    'Act:
    Set EC_Combined = EC1.Combine(EC2)
    
    'Assert:
    Assert.IsTrue EC_Combined.Count = 10

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_EmployeeCollectionCombine_FiveAndFive_CorrectEmplIDs()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Index As Long
    Dim EC1 As EmployeeCollection
    Dim EC2 As EmployeeCollection
    Dim EC_Combined As EmployeeCollection
    Dim Emp(10) As Employee
    
    Set EC1 = New EmployeeCollection
    Set EC2 = New EmployeeCollection
    
    For Index = 0 To 9
        Set Emp(Index) = New Employee
    Next Index
    
    Emp(0).EmplID = "123"
    Emp(1).EmplID = "456"
    Emp(2).EmplID = "789"
    Emp(3).EmplID = "012"
    Emp(4).EmplID = "345"
    Emp(5).EmplID = "678"
    Emp(6).EmplID = "901"
    Emp(7).EmplID = "234"
    Emp(8).EmplID = "567"
    Emp(9).EmplID = "890"
    
    Emp(0).Name = "Ash"
    Emp(1).Name = "Misty"
    Emp(2).Name = "Brock"
    Emp(3).Name = "Gary"
    Emp(4).Name = "Oak"
    Emp(5).Name = "Elm"
    Emp(6).Name = "Bill"
    Emp(7).Name = "Ericka"
    Emp(8).Name = "Lance"
    Emp(9).Name = "James"
    
    For Index = 0 To 4
        EC1.Add Emp(Index)
    Next Index
    
    For Index = 5 To 9
        EC2.Add Emp(Index)
    Next Index
    
    'Act:
    Set EC_Combined = EC1.Combine(EC2)
    
    'Assert:
    Assert.IsTrue EC_Combined.Item(1).EmplID = "123"
    Assert.IsTrue EC_Combined.Item(2).EmplID = "456"
    Assert.IsTrue EC_Combined.Item(3).EmplID = "789"
    Assert.IsTrue EC_Combined.Item(4).EmplID = "012"
    Assert.IsTrue EC_Combined.Item(5).EmplID = "345"
    Assert.IsTrue EC_Combined.Item(6).EmplID = "678"
    Assert.IsTrue EC_Combined.Item(7).EmplID = "901"
    Assert.IsTrue EC_Combined.Item(8).EmplID = "234"
    Assert.IsTrue EC_Combined.Item(9).EmplID = "567"
    Assert.IsTrue EC_Combined.Item(10).EmplID = "890"
    

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test_EmployeeCollectionCombine_FiveAndFive_CorrectNames()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Index As Long
    Dim EC1 As EmployeeCollection
    Dim EC2 As EmployeeCollection
    Dim EC_Combined As EmployeeCollection
    Dim Emp(10) As Employee
    
    Set EC1 = New EmployeeCollection
    Set EC2 = New EmployeeCollection
    
    For Index = 0 To 9
        Set Emp(Index) = New Employee
    Next Index
    
    Emp(0).EmplID = "123"
    Emp(1).EmplID = "456"
    Emp(2).EmplID = "789"
    Emp(3).EmplID = "012"
    Emp(4).EmplID = "345"
    Emp(5).EmplID = "678"
    Emp(6).EmplID = "901"
    Emp(7).EmplID = "234"
    Emp(8).EmplID = "567"
    Emp(9).EmplID = "890"
    
    Emp(0).Name = "Ash"
    Emp(1).Name = "Misty"
    Emp(2).Name = "Brock"
    Emp(3).Name = "Gary"
    Emp(4).Name = "Oak"
    Emp(5).Name = "Elm"
    Emp(6).Name = "Bill"
    Emp(7).Name = "Ericka"
    Emp(8).Name = "Lance"
    Emp(9).Name = "James"
    
    For Index = 0 To 4
        EC1.Add Emp(Index)
    Next Index
    
    For Index = 5 To 9
        EC2.Add Emp(Index)
    Next Index
    
    'Act:
    Set EC_Combined = EC1.Combine(EC2)
    
    'Assert:
    Assert.IsTrue EC_Combined.Item(1).Name = "Ash"
    Assert.IsTrue EC_Combined.Item(2).Name = "Misty"
    Assert.IsTrue EC_Combined.Item(3).Name = "Brock"
    Assert.IsTrue EC_Combined.Item(4).Name = "Gary"
    Assert.IsTrue EC_Combined.Item(5).Name = "Oak"
    Assert.IsTrue EC_Combined.Item(6).Name = "Elm"
    Assert.IsTrue EC_Combined.Item(7).Name = "Bill"
    Assert.IsTrue EC_Combined.Item(8).Name = "Ericka"
    Assert.IsTrue EC_Combined.Item(9).Name = "Lance"
    Assert.IsTrue EC_Combined.Item(10).Name = "James"
    

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
