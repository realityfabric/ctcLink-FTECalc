Attribute VB_Name = "Main"
'@Folder(FTECalc)
Option Explicit

Public Const DEBUG_ON As Boolean = True

Private Type TMain
    Directory As String
End Type

Private this As TMain

Public Property Let Directory(ByVal Dir As String)
    this.Directory = Dir
End Property

Public Property Get Directory() As String
    Directory = this.Directory
End Property

Public Function CalculateFTE(ByVal hoursWorked As Single) As Single
    Dim FTE As Single
    FTE = hoursWorked / 198
    
    CalculateFTE = FTE
End Function

Public Function inJobCodeList(ByVal Value As Variant) As Boolean
    ' Default to False
    inJobCodeList = False
    
    If Trim$(Value) = "PTF" Then
        inJobCodeList = True
    ElseIf Trim$(Value) = "PTH" Then
        inJobCodeList = True
    ElseIf Trim$(Value) = "SUB" Then
        inJobCodeList = True
    Else
        inJobCodeList = False
    End If
End Function

'@Description (The primary (main) Sub. The application should start and end here.)
Public Sub Main()
    ' TODO Reorganize code so that the application enters via Main, activating the necessary forms, instead of the other way around.
    Dim wb As Workbook
    
    Dim HourlyEmployees As EmployeeCollection
    Dim HourlyEmployees_Temp As EmployeeCollection
    Dim AppointedEmployees As EmployeeCollection
    Dim AppointedEmployees_Temp As EmployeeCollection
    
    Set HourlyEmployees = New EmployeeCollection
    Set HourlyEmployees_Temp = New EmployeeCollection
    Set AppointedEmployees = New EmployeeCollection
    Set AppointedEmployees_Temp = New EmployeeCollection
    
    Dim fileName As Variant
    For Each fileName In frmFileSelection.GetFileNames
        If fileName <> vbNullString Then
            Workbooks.Open fileName:=this.Directory & "\" & fileName, ReadOnly:=True
            Set wb = Workbooks.Item(fileName)
            
            Dim ws As Worksheet
            For Each ws In wb.Sheets
                If ws.Name Like "*Hourly*" Then
                    Set HourlyEmployees_Temp = HourlyEmployees_Temp.CreateEmployeeCollectionFromWorksheet(ws, True, Source:=wb.Name)
                    Set HourlyEmployees = HourlyEmployees.Combine(HourlyEmployees_Temp)
                ElseIf ws.Name Like "*Appointed*" Then
                    Set AppointedEmployees_Temp = AppointedEmployees_Temp.CreateEmployeeCollectionFromWorksheet(ws, False, Source:=wb.Name)
                    Set AppointedEmployees = AppointedEmployees.Combine(AppointedEmployees_Temp)
                End If
            Next ws
            
            Workbooks.Item(fileName).Close SaveChanges:=False
        End If
    Next fileName

    Dim HourlyAndAppointedEmployees As EmployeeCollection
    Set HourlyAndAppointedEmployees = AppointedEmployees.Combine(HourlyEmployees)
    
    Dim wsFTECombined As Worksheet
    Dim wsFTEByJobCode As Worksheet
    Dim wsFTEByDepartment As Worksheet
    Dim AlertsEnabled As Boolean
    
    ' Ensure that DisplayAlerts is disabled, delete sheets,
    '   then ensure that DisplayAlerts is set to whatever value it held before this action.
    AlertsEnabled = Application.DisplayAlerts
    Application.DisplayAlerts = True
    
    Set wsFTECombined = BuildFTECombined(HourlyAndAppointedEmployees)
    wsFTECombined.Range("A1:ZZ9999").Copy _
        Destination:=Sheet1.Range("A1:ZZ9999")
    wsFTECombined.Delete
    
    Set wsFTEByJobCode = BuildFTESummaryByJobCode(HourlyAndAppointedEmployees)
    wsFTEByJobCode.Range("A1:ZZ9999").Copy _
        Destination:=Sheet2.Range("A1:ZZ9999")
    wsFTEByJobCode.Delete
    
    Set wsFTEByDepartment = BuildFTESummaryByDepartment(HourlyAndAppointedEmployees)
    wsFTEByDepartment.Range("A1:ZZ9999").Copy _
        Destination:=Sheet3.Range("A1:ZZ9999")
    wsFTEByDepartment.Delete
    
    Application.DisplayAlerts = AlertsEnabled
        
    End
    
    Unload frmFileSelection
    
End Sub

Public Function BuildFTECombined(ByVal Employees As EmployeeCollection) As Worksheet
    Dim ws As Worksheet
    Dim EC As EmployeeCollection
    Dim Emp As Employee
    Dim Index As Long
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets.Item(ThisWorkbook.Sheets.Count))
    Set EC = Employees.Copy
    
    ws.Range("A1").Value = "Empl ID"
    ws.Range("B1").Value = "Name (LN, FN)"
    ws.Range("C1").Value = "Department"
    ws.Range("D1").Value = "Job Code"
    ws.Range("E1").Value = "Hours"
    ws.Range("F1").Value = "FTE%"
    ws.Range("G1").Value = "Source"
    
    For Index = 1 To EC.Count
        Set Emp = EC.Item(Index)
        WBTools.setCellValueAt ws, 1, Index + 1, Emp.EmplID
        WBTools.setCellValueAt ws, 2, Index + 1, Emp.Name
        WBTools.setCellValueAt ws, 3, Index + 1, Emp.Department
        WBTools.setCellValueAt ws, 4, Index + 1, Emp.JobCode
        WBTools.setCellValueAt ws, 5, Index + 1, Emp.hoursWorked
        WBTools.setCellValueAt ws, 6, Index + 1, Emp.hoursWorked / 198 * 100
        WBTools.setCellValueAt ws, 7, Index + 1, Emp.Source
    Next Index
    
    Set BuildFTECombined = ws
    Set EC = Nothing
End Function

Public Function BuildFTESummaryByJobCode(ByVal Employees As EmployeeCollection) As Worksheet
    Dim ws As Worksheet
    Dim EC As EmployeeCollection
    Dim Emp As Employee
    Dim Index As Long
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets.Item(ThisWorkbook.Sheets.Count))
    Set EC = Employees.Copy
    
    ws.Range("A1").Value = "Empl ID"
    ws.Range("B1").Value = "Name (LN, FN)"
    ws.Range("C1").Value = "Job Code"
    ws.Range("D1").Value = "Hours"
    ws.Range("E1").Value = "FTE%"
    
    Dim employeesByJobCode As EmployeeCollection
    Set employeesByJobCode = EC.MergeDuplicateEmployeesOnJobCode()
    
    For Index = 1 To employeesByJobCode.Count
        Set Emp = employeesByJobCode.Item(Index)
        WBTools.setCellValueAt ws, 1, Index + 1, Emp.EmplID
        WBTools.setCellValueAt ws, 2, Index + 1, Emp.Name
        WBTools.setCellValueAt ws, 3, Index + 1, Emp.JobCode
        WBTools.setCellValueAt ws, 4, Index + 1, Emp.hoursWorked
        WBTools.setCellValueAt ws, 5, Index + 1, Emp.hoursWorked / 198 * 100
    Next Index
    
    Set BuildFTESummaryByJobCode = ws
    Set EC = Nothing
End Function

Public Function BuildFTESummaryByDepartment(ByVal Employees As EmployeeCollection) As Worksheet
    Dim ws As Worksheet
    Dim EC As EmployeeCollection
    
    Dim DataArray() As Variant

    Dim Index As Long

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets.Item(ThisWorkbook.Sheets.Count))
    Set EC = Employees.Copy
    
    ws.Range("A1").Value = "Empl ID"
    ws.Range("B1").Value = "Name (LN, FN)"
    ws.Range("C1").Value = "Department"
    ws.Range("D1").Value = "Hours"
    ws.Range("E1").Value = "FTE%"
    
    Dim employeesByDepartment As EmployeeCollection
    Set employeesByDepartment = EC.MergeDuplicateEmployeesOnDepartment()
    
    ReDim DataArray(employeesByDepartment.Count, 5)
    
    For Index = 0 To employeesByDepartment.Count - 1:
        DataArray(Index, 0) = employeesByDepartment.Item(Index + 1).EmplID
        DataArray(Index, 1) = employeesByDepartment.Item(Index + 1).Name
        DataArray(Index, 2) = employeesByDepartment.Item(Index + 1).Department
        DataArray(Index, 3) = employeesByDepartment.Item(Index + 1).hoursWorked
        DataArray(Index, 4) = employeesByDepartment.Item(Index + 1).hoursWorked / 198 * 100
    Next Index
    
    ws.Range("A2:E25").Value = DataArray
    
    Set BuildFTESummaryByDepartment = ws
    Set EC = Nothing
End Function


