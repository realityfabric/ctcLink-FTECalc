VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmployeeCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder(HCM)
'@ModuleDescription(A custom implementation of Collection used to store, sort, etc instances of the Employee class.)

Private Type TEmployeeCollection
    Employees As Collection
    ' Note: The Collection class has the following:
    ' Collection.Add(Item, [Key], [Before], [After])
    ' Collection.Remove(Index)
    ' Collection.Item(Index)
    ' Collection.Count()
End Type

Private this As TEmployeeCollection

'@Description(Adds an instance of the Employee class to the Collection.)
Public Sub Add(ByVal employeeToBeAdded As Employee, Optional ByVal CollectionName As String = vbNullString)
    If Main.DEBUG_ON Then
        Dim DebugString As String
        DebugString = "EmployeeCollection.Add " & employeeToBeAdded.Name & " (" & employeeToBeAdded.eKey
        If CollectionName <> vbNullString Then DebugString = DebugString & ", " & CollectionName
        DebugString = DebugString & ")"
        Debug.Print DebugString
    End If
    
    If Not employeeToBeAdded Is Nothing Then
        If Not HasKey(employeeToBeAdded.eKey) Then
            this.Employees.Add employeeToBeAdded, employeeToBeAdded.eKey
            If Main.DEBUG_ON Then Debug.Print "Added: " & employeeToBeAdded.Name & " (" & employeeToBeAdded.eKey & ")"
        Else
            If Main.DEBUG_ON Then Debug.Print "Already Exists: " & employeeToBeAdded.Name & " (" & employeeToBeAdded.eKey & ")"
            Dim employeeExisting As Employee
            Dim employeeMerged As Employee
            Set employeeExisting = Item(employeeToBeAdded.eKey)
            Set employeeMerged = employeeToBeAdded.MergeEmployeeHours(employeeExisting, employeeExisting.Department, employeeExisting.JobCode, employeeExisting.Source)
            If Main.DEBUG_ON Then Debug.Print employeeToBeAdded.eKey & " - " & Item(employeeToBeAdded.eKey).Name
            Remove Item(employeeToBeAdded.eKey)
            this.Employees.Add employeeMerged, employeeMerged.eKey
        End If
    Else
        If Main.DEBUG_ON Then Debug.Print "Empty Object"
    End If
End Sub

'@Obsolete
'@Description(Checks to see if an Employee key exists in the Collection, returning True or False.)
Public Function HasKey(ByVal Key As Variant) As Boolean
    '@Ignore VariableNotUsed
    Dim var As Employee
    On Error GoTo ERROR_HANDLING
    '@Ignore AssignmentNotUsed
    Set var = this.Employees.Item(Key)
    HasKey = (Err.Number = 0)
    
    GoTo FINALLY
    
ERROR_HANDLING:
    HasKey = (Err.Number = 0)
    
FINALLY:
    If Main.DEBUG_ON Then Debug.Print "EmployeeCollection.HasKey(" & Key & ") = " & (Err.Number = 0)
    Err.Clear
End Function

'@Description(Removes an instance of the Employee class from the Collection.)
Public Sub Remove(ByVal Emp As Employee)
    If Main.DEBUG_ON Then Debug.Print "EmployeeCollection.Remove: " & Emp.eKey
    this.Employees.Remove (Emp.eKey)
End Sub

Public Property Get Item(ByVal Index As Variant) As Employee
    Set Item = this.Employees.Item(Index)
End Property

Public Property Get Employees() As Collection
    Set Employees = this.Employees
End Property

Private Sub Class_Initialize()
    Set this.Employees = New Collection
End Sub

Public Function Count() As Long
    Count = this.Employees.Count
End Function

Public Function GetItemsWithEmplID(ByVal EmplID As String) As Collection
    Dim Index As Long
    Dim subCollection As Collection
    Dim Emp As Employee
    
    Set subCollection = New Collection
    
    For Index = 1 To Count()
        Set Emp = Item(Index)
        If Emp.EmplID = EmplID Then subCollection.Add Emp
    Next Index
    
    Set GetItemsWithEmplID = subCollection
End Function

Public Function GetItemsWithDept(ByVal dept As String) As Collection
    Dim Index As Long
    Dim subCollection As Collection
    Dim Emp As Employee

    Set subCollection = New Collection

    For Index = 1 To Count()
        Set Emp = Item(Index)
        If Emp.Department = dept Then subCollection.Add Emp
    Next Index

    Set GetItemsWithDept = subCollection
End Function

'@Description(Returns an EmployeeCollection including all Employee instances with the specified JobCode.)
Public Function EmployeesWithJobCode(ByVal JobCode As String) As EmployeeCollection
    If Main.DEBUG_ON Then Debug.Print "EmployeeCollection.EmployeesWithJobCode(" & JobCode & ")"
    Dim EmployeesWithJobCode_Result As EmployeeCollection
    Set EmployeesWithJobCode_Result = New EmployeeCollection
    
    Dim Index As Long
    Dim Emp As Employee
    Set Emp = Nothing
    
    For Index = 1 To this.Employees.Count
        Set Emp = this.Employees.Item(Index)
        
        If Emp.JobCode = JobCode Then
            EmployeesWithJobCode_Result.Add Emp
        End If
        Set Emp = Nothing
    Next Index

    Set EmployeesWithJobCode = EmployeesWithJobCode_Result
End Function

'@Description(Returns an EmployeeCollection including all Employee instances with the specified Department.)
Public Function EmployeesWithDepartment(ByVal Department As String) As EmployeeCollection
    If Main.DEBUG_ON Then Debug.Print "EmployeeCollection.EmployeesWithDepartment(" & Department & ")"
    Dim EmployeesWithDepartment_Result As EmployeeCollection
    Set EmployeesWithDepartment_Result = New EmployeeCollection
    
    Dim Index As Long
    Dim Emp As Employee
    Set Emp = Nothing
    
    For Index = 1 To this.Employees.Count
        Set Emp = this.Employees.Item(Index)
        
        If Emp.Department = Department Then
            EmployeesWithDepartment_Result.Add Emp
        End If
        Set Emp = Nothing
        Index = Index + 1
    Next Index

    Set EmployeesWithDepartment = EmployeesWithDepartment_Result
End Function

Public Function MergeDuplicateEmployeesOnJobCode() As EmployeeCollection
    Dim mergedColletion As EmployeeCollection
    Set mergedColletion = New EmployeeCollection
    Dim Emp As Employee
    Dim Index As Long
    
    For Index = 1 To this.Employees.Count
        Set Emp = this.Employees.Item(Index).Copy()
        Emp.Department = vbNullString

        mergedColletion.Add Emp
    Next Index
    
    Set MergeDuplicateEmployeesOnJobCode = mergedColletion
End Function

Public Function MergeDuplicateEmployeesOnDepartment() As EmployeeCollection
    If Main.DEBUG_ON Then Debug.Print "START MergeDuplicateEmployeesOnDepartment()"

    Dim EC As EmployeeCollection
    Dim Emp As Employee
    Dim Index As Long

    Set EC = New EmployeeCollection

    For Index = 1 To this.Employees.Count
        Set Emp = this.Employees.Item(Index).Copy

        If Emp.Name <> vbNullString Then
            Emp.JobCode = vbNullString
            EC.Add Emp
            
            If Main.DEBUG_ON Then Emp.Print_Employee
        End If
        
        Set Emp = Nothing
    Next
    
    Set MergeDuplicateEmployeesOnDepartment = EC
    If Main.DEBUG_ON Then Debug.Print "END MergeDuplicateEmployeesOnDepartment()"
End Function

Public Function PruneEmployeesToJobCodeList() As EmployeeCollection
    Dim NewEmployeeCollection As EmployeeCollection
    Set NewEmployeeCollection = New EmployeeCollection
    Dim Emp As Employee
    
    For Each Emp In this.Employees
        If Main.inJobCodeList(Emp.JobCode) Then
            NewEmployeeCollection.Add Emp.Copy
        End If
    Next Emp
    
    Set PruneEmployeesToJobCodeList = NewEmployeeCollection
End Function

Public Function Combine(ByVal EmployeeCollectionToBeCombined As EmployeeCollection) As EmployeeCollection
    Dim CombinedEmployees As EmployeeCollection
    Set CombinedEmployees = New EmployeeCollection
    
    Dim Emp As Employee
    For Each Emp In this.Employees
        CombinedEmployees.Add Emp
    Next Emp
    For Each Emp In EmployeeCollectionToBeCombined.Employees
        CombinedEmployees.Add Emp
    Next Emp
    
    Set Combine = CombinedEmployees
End Function

Public Function Copy() As EmployeeCollection
    Dim EmpsCopy As EmployeeCollection
    Dim Emp As Employee
    
    Set EmpsCopy = New EmployeeCollection
    
    For Each Emp In this.Employees
        EmpsCopy.Add Emp.Copy
    Next Emp
    
    Set Copy = EmpsCopy
End Function

Public Function CreateEmployeeCollectionFromWorksheet(ByVal ws As Worksheet, ByVal IsHourly As Boolean, Optional ByVal HeaderRow As Long = 1, Optional ByVal Source As String = vbNullString) As EmployeeCollection
    If IsHourly Then
        Set CreateEmployeeCollectionFromWorksheet = CreateEmployeeCollectionFromWorksheet_Hourly(ws, HeaderRow, Source)
    Else
        Set CreateEmployeeCollectionFromWorksheet = CreateEmployeeCollectionFromWorksheet_Appointed(ws, HeaderRow, Source)
    End If
End Function

Private Function CreateEmployeeCollectionFromWorksheet_Hourly(ByVal ws As Worksheet, ByVal HeaderRow As Long, ByVal Source As String) As EmployeeCollection
    Dim EC As EmployeeCollection
    Dim Emp As Employee
    
    Set EC = New EmployeeCollection
    ' Emp will be set in loop below
    
    Dim LastRow As Long
    Dim EmplIDCol As Long
    Dim NameCol As Long
    Dim DeptCol As Long
    Dim JobCodeCol As Long
    Dim CanceledClassCol As Long
    
    ' 01A, 01B, 02A, ..., 12A, 12B
    Dim PayPeriodCol(24) As Long
    Dim PayPeriodString As String

    Dim Index As Long
    Dim Row As Long
    
    LastRow = WBTools.FindLastRowInSheet(ws)
    EmplIDCol = WBTools.FindColumnLikeName(ws, "*Empl*ID*")
    NameCol = WBTools.FindColumnLikeName(ws, "Name*")
    DeptCol = WBTools.FindColumnLikeName(ws, "DEPT*")
    JobCodeCol = WBTools.FindColumnLikeName(ws, "*Job Code*")
    CanceledClassCol = WBTools.FindColumnByName(ws, "Canceled Class")
    
    For Index = 0 To 23
        If Index Mod 2 = 0 Then
            PayPeriodString = Format$(WorksheetFunction.RoundUp((Index + 1) / 2, 0), "00") & "A"
        Else
            PayPeriodString = Format$(WorksheetFunction.RoundUp((Index + 1) / 2, 0), "00") & "B"
        End If
        
        PayPeriodCol(Index) = WBTools.FindColumnLikeName(ws, PayPeriodString & "*")
    Next Index
    
    For Row = HeaderRow + 1 To LastRow
        Set Emp = New Employee
        
        '@Ignore EmptyIfBlock
        If WBTools.getCellValueAt(ws, EmplIDCol, Row) = Empty Then
            ' No EmplID. Skip.
        Else
            Emp.EmplID = WBTools.getCellValueAt(ws, EmplIDCol, Row)
            Emp.Name = WBTools.getCellValueAt(ws, NameCol, Row)
            Emp.Department = WBTools.getCellValueAt(ws, DeptCol, Row)
            Emp.IsCanceled = Not WBTools.getCellValueAt(ws, CanceledClassCol, Row) = Empty
            Emp.IsHourly = True
            Emp.JobCode = WBTools.getCellValueAt(ws, JobCodeCol, Row)
            Emp.Source = Source
            
            For Index = 0 To 23
                If Index Mod 2 = 0 Then
                    PayPeriodString = Format$(WorksheetFunction.RoundUp((Index + 1) / 2, 0), "00") & "A"
                Else
                    PayPeriodString = Format$(WorksheetFunction.RoundUp((Index + 1) / 2, 0), "00") & "B"
                End If
                
                If PayPeriodCol(Index) > -1 Then Emp.hoursWorked(PayPeriodString) = WBTools.getCellValueAt(ws, PayPeriodCol(Index), Row)
            Next Index
            
            If Not Emp Is Nothing Then EC.Add Emp
        End If
     Next Row
     
     Set CreateEmployeeCollectionFromWorksheet_Hourly = EC
End Function

Private Function CreateEmployeeCollectionFromWorksheet_Appointed(ByVal ws As Worksheet, ByVal HeaderRow As Long, ByVal Source As String) As EmployeeCollection
    Dim EC As EmployeeCollection
    Dim Emp As Employee
    
    Set EC = New EmployeeCollection
    ' Emp will be set in loop below

    Dim LastRow As Long
    Dim EmplIDCol As Long
    Dim NameCol As Long
    Dim DeptCol As Long
    Dim JobCodeCol As Long
    Dim CanceledClassCol As Long
    Dim PayPeriodCol As Long

    Dim Row As Long

    LastRow = WBTools.FindLastRowInSheet(ws)
    EmplIDCol = WBTools.FindColumnLikeName(ws, "*Empl*ID*")
    NameCol = WBTools.FindColumnLikeName(ws, "Name*")
    DeptCol = WBTools.FindColumnLikeName(ws, "DEPT*")
    JobCodeCol = WBTools.FindColumnLikeName(ws, "*Job Code*")
    CanceledClassCol = WBTools.FindColumnByName(ws, "Canceled Class")
    PayPeriodCol = WBTools.FindColumnByName(ws, "Cntct hrs")
    
    For Row = HeaderRow + 1 To LastRow
        Set Emp = New Employee
        
        '@Ignore EmptyIfBlock
        If WBTools.getCellValueAt(ws, EmplIDCol, Row) = Empty Then
            ' No EmplID. Skip.
        Else
            Emp.EmplID = WBTools.getCellValueAt(ws, EmplIDCol, Row)
            Emp.Name = WBTools.getCellValueAt(ws, NameCol, Row)
            Emp.Department = WBTools.getCellValueAt(ws, DeptCol, Row)
            Emp.IsCanceled = Not WBTools.getCellValueAt(ws, CanceledClassCol, Row) = Empty
            Emp.IsHourly = False
            Emp.JobCode = WBTools.getCellValueAt(ws, JobCodeCol, Row)
            Emp.Source = Source
            Emp.hoursWorked("OTH") = WBTools.getCellValueAt(ws, PayPeriodCol, Row)
            
            If Not Emp Is Nothing Then EC.Add Emp
        End If
     Next Row
     
     Set CreateEmployeeCollectionFromWorksheet_Appointed = EC
End Function
