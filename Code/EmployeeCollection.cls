VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmployeeCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FTECalc")
Option Explicit

Private Type TEmployeeCollection
    Employees As Collection
End Type

Private this As TEmployeeCollection

Private Sub Class_Initialize()
    Set this.Employees = New Collection
End Sub

'@Ignore ParameterNotUsed
Public Sub Add(ByVal Item As Employee _
               , Optional ByVal Key As String = vbNullString _
                , Optional ByVal Before As Variant = Nothing _
                 , Optional ByVal After As Variant = Nothing)
    ' TODO: Implement Key, Before, After
    this.Employees.Add Item
End Sub

Public Function Item(ByVal Index As Variant) As Employee
    Set Item = this.Employees.Item(Index)
End Function

Public Sub Remove(ByVal Index As Variant)
    this.Employees.Remove Index
End Sub

Public Function Count() As Long
    Count = this.Employees.Count()
End Function

'@Ignore UseMeaningfulName
Public Sub Concat(ByVal EmployeeCollection2 As EmployeeCollection)
    Dim Index As Long
    For Index = 1 To EmployeeCollection2.Count
        Add EmployeeCollection2.Item(Index)
    Next Index
End Sub

Public Function Filter(Optional ByVal EmplIDFilter As String = vbNullString _
                       , Optional ByVal DeptIDFilter As String = vbNullString _
                        , Optional ByVal JobCodeFilter As String = vbNullString _
                         ) As EmployeeCollection
    Dim FilteredCollection As EmployeeCollection
    Dim Index As Long
    Set FilteredCollection = New EmployeeCollection

    For Index = 1 To Count
        If (EmplIDFilter = vbNullString Or EmplIDFilter = Item(Index).EmplID) _
           And (DeptIDFilter = vbNullString Or DeptIDFilter = Item(Index).DeptID) _
           And (JobCodeFilter = vbNullString Or JobCodeFilter = Item(Index).JobCode) Then
            FilteredCollection.Add Item(Index)
        End If
    Next Index
    
    Set Filter = FilteredCollection
End Function

' Headers: Will include headers at (0,) if True, otherwise data will begin at (0,)
' IncludeDeptID: Include DeptID field if True
' IncludeJobCode: Include JobCode field if True
Public Function ToArrayContainer(Optional ByVal Headers As Boolean = True _
                                 , Optional ByVal IncludeDeptID As Boolean = True _
                                  , Optional ByVal IncludeJobCode As Boolean = True _
                                   ) As ArrayContainer
        
    ' Data Columns:
    ' EmplID
    ' Name
    ' [DeptID]
    ' [JobCode]
    ' Hours
    ' FTE
        
    Dim DataContainer As ArrayContainer
    Dim DataArray() As Variant
    Dim RowCount As Long                         ' (r,)
    Dim ColumnCount As Long                      ' (,c)
    Dim CollectionIndex As Long
    Dim RowIndex As Long
    Dim ColumnIndex As Long
    Dim Emp As Employee
    
    Set DataContainer = New ArrayContainer
    
    RowCount = Count()
    If Headers Then RowCount = RowCount + 1
    
    ColumnCount = 4                              ' EmplID, Name, Hours, FTE
    If IncludeDeptID Then ColumnCount = ColumnCount + 1
    If IncludeJobCode Then ColumnCount = ColumnCount + 1
    
    ReDim DataArray(RowCount, ColumnCount)
    
    RowIndex = 0
    If Headers Then
        ColumnIndex = 0
        DataArray(RowIndex, ColumnIndex) = "EmplID"
        
        ColumnIndex = ColumnIndex + 1
        DataArray(RowIndex, ColumnIndex) = "Name"
        
        If IncludeDeptID Then
            ColumnIndex = ColumnIndex + 1
            DataArray(RowIndex, ColumnIndex) = "DeptID"
        End If
        
        If IncludeJobCode Then
            ColumnIndex = ColumnIndex + 1
            DataArray(RowIndex, ColumnIndex) = "JobCode"
        End If
        
        ColumnIndex = ColumnIndex + 1
        DataArray(RowIndex, ColumnIndex) = "Hours"
        
        ColumnIndex = ColumnIndex + 1
        DataArray(RowIndex, ColumnIndex) = "FTE%"
    End If
    
    For CollectionIndex = 1 To Count
        RowIndex = RowIndex + 1
        ColumnIndex = 0
        Set Emp = Item(CollectionIndex)
        DataArray(RowIndex, ColumnIndex) = Emp.EmplID
        
        ColumnIndex = ColumnIndex + 1
        DataArray(RowIndex, ColumnIndex) = Emp.Name
        
        If IncludeDeptID Then
            ColumnIndex = ColumnIndex + 1
            DataArray(RowIndex, ColumnIndex) = Emp.DeptID
        End If
        
        If IncludeJobCode Then
            ColumnIndex = ColumnIndex + 1
            DataArray(RowIndex, ColumnIndex) = Emp.JobCode
        End If
        
        ColumnIndex = ColumnIndex + 1
        DataArray(RowIndex, ColumnIndex) = Emp.HoursWorked()
        
        ColumnIndex = ColumnIndex + 1
        DataArray(RowIndex, ColumnIndex) = Emp.fte()
    Next CollectionIndex
    
    DataContainer.SetData RowCount, ColumnCount, DataArray
    
    Set ToArrayContainer = DataContainer
End Function

Public Function CountUniqueEmplIDs() As Long
    Dim EmplIDCounts As Variant
    Dim EmplID As String
    Dim Index As Long
    Dim UniqueEmplIDs As Long
    Set EmplIDCounts = CreateObject("Scripting.Dictionary")
    UniqueEmplIDs = 0
    
    For Index = 1 To Count
        EmplID = Item(Index).EmplID
        If Not EmplIDCounts.Exists(EmplID) Then
            UniqueEmplIDs = UniqueEmplIDs + 1
            EmplIDCounts.Add EmplID, 1
        End If
    Next Index
    
    CountUniqueEmplIDs = UniqueEmplIDs
End Function

' Returns a dictionary with key=EmplID and val=count
Public Function UniqueEmplIDs() As Variant
    Dim EmplIDs As Variant
    Dim EmplID As String
    Dim Index As Long
    
    Set EmplIDs = CreateObject("Scripting.Dictionary")
    
    For Index = 1 To Count
        EmplID = Item(Index).EmplID
        If Not EmplIDs.Exists(EmplID) Then
            EmplIDs.Add EmplID, 1
        Else
            EmplIDs(EmplID) = EmplIDs(EmplID) + 1
        End If
    Next Index
    
    Set UniqueEmplIDs = EmplIDs
End Function

' Returns a dictionary with key=DeptID and val=count
Public Function UniqueDeptIDs() As Variant
    Dim DeptIDs As Variant
    Dim DeptID As String
    Dim Index As Long
    
    Set DeptIDs = CreateObject("Scripting.Dictionary")
    
    For Index = 1 To Count
        DeptID = Item(Index).DeptID
        If Not DeptIDs.Exists(DeptID) Then
            DeptIDs.Add DeptID, 1
        Else
            DeptIDs(DeptID) = DeptIDs(DeptID) + 1
        End If
    Next Index
    
    Set UniqueDeptIDs = DeptIDs
End Function

' Returns a dictionary with key=EmplID and val=count
Public Function UniqueJobCodes() As Variant
    Dim JobCodes As Variant
    Dim JobCode As String
    Dim Index As Long
    
    Set JobCodes = CreateObject("Scripting.Dictionary")
    
    For Index = 1 To Count
        JobCode = Item(Index).JobCode
        If Not JobCodes.Exists(JobCode) Then
            JobCodes.Add JobCode, 1
        Else
            JobCodes(JobCode) = JobCodes(JobCode) + 1
        End If
    Next Index
    
    Set UniqueJobCodes = JobCodes
End Function

Public Function MergeAllEmployeesOnEmplID() As EmployeeCollection
    Dim MergedCollection As EmployeeCollection
    Dim FilteredCollection As EmployeeCollection
    Dim MergedEmployee As Employee
    Dim Index As Long
    Dim EmplIDs As Variant
    Dim Key As Variant
    
    Set MergedCollection = New EmployeeCollection
    Set FilteredCollection = New EmployeeCollection
    Set EmplIDs = UniqueEmplIDs()
    
    For Each Key In EmplIDs.Keys
        Set FilteredCollection = Filter(EmplIDFilter:=Key)
        Set MergedEmployee = FilteredCollection.Item(1)
        For Index = 2 To FilteredCollection.Count
            Set MergedEmployee = MergedEmployee.Merge( _
                                 FilteredCollection.Item(Index), _
                                 PreserveDeptID:=False, _
                                 PreserveJobCode:=False)
        Next Index
        MergedCollection.Add MergedEmployee
    Next Key

    Set MergeAllEmployeesOnEmplID = MergedCollection
End Function

Public Function MergeAllEmployees() As EmployeeCollection
    Dim MergedCollection As EmployeeCollection
    Dim FilteredCollection As EmployeeCollection
    Dim MergedEmployee As Employee
    Dim Index As Long
    Dim EmplIDs As Variant
    Dim DeptIDs As Variant
    Dim JobCodes As Variant
    Dim EmplIDKey As Variant
    Dim DeptIDKey As Variant
    Dim JobCodeKey As Variant
    
    Set MergedCollection = New EmployeeCollection
    Set FilteredCollection = New EmployeeCollection
    Set EmplIDs = UniqueEmplIDs()
    Set DeptIDs = UniqueDeptIDs()
    Set JobCodes = UniqueJobCodes()
    
    For Each EmplIDKey In EmplIDs.Keys
        For Each DeptIDKey In DeptIDs.Keys
            For Each JobCodeKey In JobCodes.Keys
                Debug.Print "E: " & EmplIDKey & "; D: " & DeptIDKey & "; J: " & JobCodeKey
                Set FilteredCollection = Filter( _
                                         EmplIDFilter:=EmplIDKey, _
                                         DeptIDFilter:=DeptIDKey, _
                                         JobCodeFilter:=JobCodeKey)
                If FilteredCollection.Count > 0 Then
                    Set MergedEmployee = FilteredCollection.Item(1)
                End If
                If FilteredCollection.Count > 1 Then
                    For Index = 2 To FilteredCollection.Count
                        Set MergedEmployee = MergedEmployee.Merge(FilteredCollection.Item(Index))
                    Next Index
                End If

                If FilteredCollection.Count > 0 Then MergedCollection.Add MergedEmployee
            Next JobCodeKey
        Next DeptIDKey
    Next EmplIDKey

    Set MergeAllEmployees = MergedCollection
End Function

Public Function MergeAllEmployeesOnDeptID() As EmployeeCollection
    Dim MergedCollection As EmployeeCollection
    Dim FilteredCollection As EmployeeCollection
    Dim MergedEmployee As Employee
    Dim Index As Long
    Dim EmplIDs As Variant
    Dim DeptIDs As Variant
    Dim EmplIDKey As Variant
    Dim DeptIDKey As Variant
    
    Set MergedCollection = New EmployeeCollection
    Set FilteredCollection = New EmployeeCollection
    Set EmplIDs = UniqueEmplIDs()
    Set DeptIDs = UniqueDeptIDs()
    
    For Each EmplIDKey In EmplIDs.Keys
        For Each DeptIDKey In DeptIDs.Keys
            Set FilteredCollection = Filter(EmplIDFilter:=EmplIDKey, DeptIDFilter:=DeptIDKey)
            If FilteredCollection.Count > 0 Then Set MergedEmployee = FilteredCollection.Item(1)
            If FilteredCollection.Count > 1 Then
                For Index = 2 To FilteredCollection.Count
                    Set MergedEmployee = MergedEmployee.Merge( _
                                         FilteredCollection.Item(Index), _
                                         PreserveJobCode:=False)
                Next Index
            End If

            If FilteredCollection.Count > 0 Then MergedCollection.Add MergedEmployee
        Next DeptIDKey
    Next EmplIDKey
    
    Set MergeAllEmployeesOnDeptID = MergedCollection
End Function

Public Function MergeAllEmployeesOnJobCode() As EmployeeCollection
    Dim MergedCollection As EmployeeCollection
    Dim FilteredCollection As EmployeeCollection
    Dim MergedEmployee As Employee
    Dim Index As Long
    Dim EmplIDs As Variant
    Dim JobCodes As Variant
    Dim EmplIDKey As Variant
    Dim JobCodeKey As Variant
    
    Set MergedCollection = New EmployeeCollection
    Set FilteredCollection = New EmployeeCollection
    Set EmplIDs = UniqueEmplIDs()
    Set JobCodes = UniqueJobCodes()
    
    For Each EmplIDKey In EmplIDs.Keys
        For Each JobCodeKey In JobCodes.Keys
            Set FilteredCollection = Filter(EmplIDFilter:=EmplIDKey, JobCodeFilter:=JobCodeKey)
            If FilteredCollection.Count > 0 Then Set MergedEmployee = FilteredCollection.Item(1)
            For Index = 2 To FilteredCollection.Count
                Set MergedEmployee = MergedEmployee.Merge( _
                                     FilteredCollection.Item(Index), _
                                     PreserveDeptID:=False)
            Next Index
            If FilteredCollection.Count > 0 Then MergedCollection.Add MergedEmployee
        Next JobCodeKey
    Next EmplIDKey
    
    Set MergeAllEmployeesOnJobCode = MergedCollection
End Function

Public Function CreateEmployeeCollectionFromWorksheet( _
       ByVal ws As Worksheet, _
       ByVal IsHourly As Boolean, _
       Optional ByVal HeaderRow As Long = 1 _
       ) As EmployeeCollection
    If IsHourly Then
        Set CreateEmployeeCollectionFromWorksheet = _
                                                  CreateEmployeeCollectionFromWorksheet_Hourly(ws, HeaderRow)
    Else
        Set CreateEmployeeCollectionFromWorksheet = _
                                                  CreateEmployeeCollectionFromWorksheet_Appointed(ws, HeaderRow)
    End If
End Function

Private Function CreateEmployeeCollectionFromWorksheet_Hourly( _
        ByVal ws As Worksheet, _
        ByVal HeaderRow As Long _
        ) As EmployeeCollection
    Dim EC As EmployeeCollection
    Dim Emp As Employee
    
    Dim LastRow As Long
    Dim EmplIDCol As Long
    Dim NameCol As Long
    Dim DeptIDCol As Long
    Dim JobCodeCol As Long
    Dim CanceledClassCol As Long
    
    Set EC = New EmployeeCollection
    
    ' 01A, 01B, 02A, ..., 12A, 12B
    Dim PayPeriodCol(24) As Long
    Dim PayPeriodString As String

    Dim Index As Long
    Dim Row As Long
    
    LastRow = WBTools.FindLastRowInSheet(ws)
    EmplIDCol = WBTools.FindColumnLikeName(ws, "*Empl*ID*")
    NameCol = WBTools.FindColumnLikeName(ws, "Name*")
    DeptIDCol = 4
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
        
        ' If there is an EmplID and the class is not canceled, then...
        If WBTools.getCellValueAt(ws, EmplIDCol, Row) <> Empty _
           And WBTools.getCellValueAt(ws, CanceledClassCol, Row) = Empty Then
            
            Emp.EmplID = WBTools.getCellValueAt(ws, EmplIDCol, Row)
            Emp.Name = WBTools.getCellValueAt(ws, NameCol, Row)
            Emp.DeptID = WBTools.getCellValueAt(ws, DeptIDCol, Row)
            Emp.JobCode = WBTools.getCellValueAt(ws, JobCodeCol, Row)
            
            For Index = 0 To 23
                If Index Mod 2 = 0 Then
                    PayPeriodString = Format$(WorksheetFunction.RoundUp((Index + 1) / 2, 0), "00") & "A"
                Else
                    PayPeriodString = Format$(WorksheetFunction.RoundUp((Index + 1) / 2, 0), "00") & "B"
                End If
                
                If PayPeriodCol(Index) > -1 Then
                    Emp.HoursWorked(PayPeriodString) = _
                                                     WBTools.getCellValueAt(ws, PayPeriodCol(Index), Row)
                End If
            Next Index
            
            If Not Emp Is Nothing Then EC.Add Emp
        End If
    Next Row
     
    Set CreateEmployeeCollectionFromWorksheet_Hourly = EC
End Function

Private Function CreateEmployeeCollectionFromWorksheet_Appointed( _
        ByVal ws As Worksheet, _
        ByVal HeaderRow As Long _
        ) As EmployeeCollection
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
    DeptCol = 6
    JobCodeCol = WBTools.FindColumnLikeName(ws, "*Job Code*")
    CanceledClassCol = WBTools.FindColumnByName(ws, "Canceled Class")
    PayPeriodCol = WBTools.FindColumnByName(ws, "Cntct hrs")
    
    For Row = HeaderRow + 1 To LastRow
        Set Emp = New Employee
        
        ' Not empty, not canceled, and a number value (ignore Independent Study)
        If WBTools.getCellValueAt(ws, EmplIDCol, Row) <> Empty _
           And WBTools.getCellValueAt(ws, CanceledClassCol, Row) = Empty _
           And ( _
           VarType(WBTools.getCellValueAt(ws, PayPeriodCol, Row)) = 2 Or _
           VarType(WBTools.getCellValueAt(ws, PayPeriodCol, Row)) = 3 Or _
           VarType(WBTools.getCellValueAt(ws, PayPeriodCol, Row)) = 4 Or _
           VarType(WBTools.getCellValueAt(ws, PayPeriodCol, Row)) = 5 Or _
           VarType(WBTools.getCellValueAt(ws, PayPeriodCol, Row)) = 14 Or _
           VarType(WBTools.getCellValueAt(ws, PayPeriodCol, Row)) = 20 _
           ) Then
            Emp.EmplID = WBTools.getCellValueAt(ws, EmplIDCol, Row)
            Emp.Name = WBTools.getCellValueAt(ws, NameCol, Row)
            Emp.DeptID = WBTools.getCellValueAt(ws, DeptCol, Row)
            Emp.JobCode = WBTools.getCellValueAt(ws, JobCodeCol, Row)
            Emp.HoursWorked("OTH") = WBTools.getCellValueAt(ws, PayPeriodCol, Row)
            
            If Not Emp Is Nothing Then EC.Add Emp
        End If
    Next Row
     
    Set CreateEmployeeCollectionFromWorksheet_Appointed = EC
End Function


