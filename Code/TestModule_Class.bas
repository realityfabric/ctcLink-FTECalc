Attribute VB_Name = "TestModule_Class"
'@IgnoreModule EmptyMethod, VariableNotUsed, UseMeaningfulName
'\@TestModule
'@Folder("donotuse.Tests")

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
'@Ignore EmptyMethod
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Setters/Getters")
'@Ignore UseMeaningfulName
Private Sub Test_ClassBeginEqual19010101()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim C As Class
    Set C = New Class
    
    'Act:
    C.ClassBegin = #1/1/1901#
    
    'Assert:
    Assert.IsTrue C.ClassBegin = #1/1/1901#

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Setters/Getters")
'@Ignore UseMeaningfulName
Private Sub Test_ClassEndEqual19010101()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim C As Class
    Set C = New Class
    
    'Act:
    C.ClassEnd = #1/1/1901#
    
    'Assert:
    Assert.IsTrue C.ClassEnd = #1/1/1901#

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Setters/Getters")
Private Sub Test_ClassNumberEqualsABCD()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Class
    Set C = New Class
    
    'Act:
    C.ClassNumber = "ABCD"
    
    'Assert:
    Assert.IsTrue C.ClassNumber = "ABCD"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Setters/Getters")
Private Sub Test_ClassNumberEquals1234()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Class
    Set C = New Class
    
    'Act:
    C.ClassNumber = "1234"
    
    'Assert:
    Assert.IsTrue C.ClassNumber = "1234"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Setters/Getters")
Private Sub Test_ComboCodeEqualsABCD()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Class
    Set C = New Class
    
    'Act:
    C.ComboCode = "ABCD"
    
    'Assert:
    Assert.IsTrue C.ComboCode = "ABCD"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Setters/Getters")
Private Sub Test_ContactHoursAreZeroByDefault()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Class
    
    'Act:
    Set C = New Class
    
    'Assert:
    Assert.IsTrue C.ContactHours("01A") = 0
    Assert.IsTrue C.ContactHours("01B") = 0
    Assert.IsTrue C.ContactHours("02A") = 0
    Assert.IsTrue C.ContactHours("02B") = 0
    Assert.IsTrue C.ContactHours("03A") = 0
    Assert.IsTrue C.ContactHours("03B") = 0
    Assert.IsTrue C.ContactHours("04A") = 0
    Assert.IsTrue C.ContactHours("04B") = 0
    Assert.IsTrue C.ContactHours("05A") = 0
    Assert.IsTrue C.ContactHours("05B") = 0
    Assert.IsTrue C.ContactHours("06A") = 0
    Assert.IsTrue C.ContactHours("06B") = 0
    Assert.IsTrue C.ContactHours("07A") = 0
    Assert.IsTrue C.ContactHours("07B") = 0
    Assert.IsTrue C.ContactHours("08A") = 0
    Assert.IsTrue C.ContactHours("08B") = 0
    Assert.IsTrue C.ContactHours("09A") = 0
    Assert.IsTrue C.ContactHours("09B") = 0
    Assert.IsTrue C.ContactHours("10A") = 0
    Assert.IsTrue C.ContactHours("10B") = 0
    Assert.IsTrue C.ContactHours("11A") = 0
    Assert.IsTrue C.ContactHours("11B") = 0
    Assert.IsTrue C.ContactHours("12A") = 0
    Assert.IsTrue C.ContactHours("12B") = 0
    Assert.IsTrue C.ContactHours("OTH") = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Setters/Getters")
Private Sub Test_ContactHoursWithoutArgZeroByDefault()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Class
    
    'Act:
    Set C = New Class
    
    'Assert:
    Assert.IsTrue C.ContactHours() = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Setters/Getters")
Private Sub Test_ContactHoursSetTo10ForEachPeriod()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Class
    Set C = New Class
    
    'Act:
    C.ContactHours("01A") = 10
    C.ContactHours("01B") = 10
    C.ContactHours("02A") = 10
    C.ContactHours("02B") = 10
    C.ContactHours("03A") = 10
    C.ContactHours("03B") = 10
    C.ContactHours("04A") = 10
    C.ContactHours("04B") = 10
    C.ContactHours("05A") = 10
    C.ContactHours("05B") = 10
    C.ContactHours("06A") = 10
    C.ContactHours("06B") = 10
    C.ContactHours("07A") = 10
    C.ContactHours("07B") = 10
    C.ContactHours("08A") = 10
    C.ContactHours("08B") = 10
    C.ContactHours("09A") = 10
    C.ContactHours("09B") = 10
    C.ContactHours("10A") = 10
    C.ContactHours("10B") = 10
    C.ContactHours("11A") = 10
    C.ContactHours("11B") = 10
    C.ContactHours("12A") = 10
    C.ContactHours("12B") = 10
    C.ContactHours("OTH") = 10
    
    
    'Assert:
    Assert.IsTrue C.ContactHours("01A") = 10
    Assert.IsTrue C.ContactHours("01B") = 10
    Assert.IsTrue C.ContactHours("02A") = 10
    Assert.IsTrue C.ContactHours("02B") = 10
    Assert.IsTrue C.ContactHours("03A") = 10
    Assert.IsTrue C.ContactHours("03B") = 10
    Assert.IsTrue C.ContactHours("04A") = 10
    Assert.IsTrue C.ContactHours("04B") = 10
    Assert.IsTrue C.ContactHours("05A") = 10
    Assert.IsTrue C.ContactHours("05B") = 10
    Assert.IsTrue C.ContactHours("06A") = 10
    Assert.IsTrue C.ContactHours("06B") = 10
    Assert.IsTrue C.ContactHours("07A") = 10
    Assert.IsTrue C.ContactHours("07B") = 10
    Assert.IsTrue C.ContactHours("08A") = 10
    Assert.IsTrue C.ContactHours("08B") = 10
    Assert.IsTrue C.ContactHours("09A") = 10
    Assert.IsTrue C.ContactHours("09B") = 10
    Assert.IsTrue C.ContactHours("10A") = 10
    Assert.IsTrue C.ContactHours("10B") = 10
    Assert.IsTrue C.ContactHours("11A") = 10
    Assert.IsTrue C.ContactHours("11B") = 10
    Assert.IsTrue C.ContactHours("12A") = 10
    Assert.IsTrue C.ContactHours("12B") = 10
    Assert.IsTrue C.ContactHours("OTH") = 10

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Setters/Getters")
Private Sub Test_ContactHoursSetTo10ForEachPeriodWithoutArgsEquals250()
    On Error GoTo TestFail
    
    'Arrange:
    Dim C As Class
    Set C = New Class
    
    'Act:
    C.ContactHours("01A") = 10
    C.ContactHours("01B") = 10
    C.ContactHours("02A") = 10
    C.ContactHours("02B") = 10
    C.ContactHours("03A") = 10
    C.ContactHours("03B") = 10
    C.ContactHours("04A") = 10
    C.ContactHours("04B") = 10
    C.ContactHours("05A") = 10
    C.ContactHours("05B") = 10
    C.ContactHours("06A") = 10
    C.ContactHours("06B") = 10
    C.ContactHours("07A") = 10
    C.ContactHours("07B") = 10
    C.ContactHours("08A") = 10
    C.ContactHours("08B") = 10
    C.ContactHours("09A") = 10
    C.ContactHours("09B") = 10
    C.ContactHours("10A") = 10
    C.ContactHours("10B") = 10
    C.ContactHours("11A") = 10
    C.ContactHours("11B") = 10
    C.ContactHours("12A") = 10
    C.ContactHours("12B") = 10
    C.ContactHours("OTH") = 10
    
    'Assert:
    Assert.IsTrue C.ContactHours() = 250
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

