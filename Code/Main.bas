Attribute VB_Name = "Main"
'@Folder(FTECalc)
Option Explicit

'@VariableDescription("Stores the Unix Timestamp at runtime, set in the Main method.")
Private UnixTimestamp As LongLong

'@Description("Returns the Unix Timestamp recorded at runtime in the Main method.")
Public Function GetTimestamp() As LongLong
    GetTimestamp = UnixTimestamp
End Function

'@Description("Returns the Unix Timestamp recorded at runtime in the Main method as a string.")
Public Function GetTimestampStr() As String
    GetTimestampStr = Trim$(Str$(GetTimestamp()))
End Function

'@EntryPoint
Public Sub Main()
    UnixTimestamp = UnixTime()
    Dim JobCodes(1 To 3) As String
    JobCodes(1) = "PTF"
    JobCodes(2) = "PTH"
    JobCodes(3) = "SUB"
    
    ' History Vars
    Dim ApplicationDisplayAlerts As Boolean
    
    ' Loop Vars
    Dim Elem As Variant
    
    ' IO Vars
    Dim Output As Workbook
    Dim OutputFileName As Variant
    Dim InputFileNames As Variant
    
    ' EmployeeCollection Vars
    Dim EC_Temp As EmployeeCollection
    Dim EC_All As EmployeeCollection
    Dim EC_Filtered As EmployeeCollection
    
    ' EmployeeCollection Array Vars
    Dim AC_Filtered_Merged As ArrayContainer
    Dim AC_GroupByDeptID As ArrayContainer
    Dim AC_GroupByJobCode As ArrayContainer
    
    ' Workbook/Worksheet Vars
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set EC_Temp = New EmployeeCollection
    Set EC_All = New EmployeeCollection
    Set EC_Filtered = New EmployeeCollection
    
    ' Get list of workbooks from user
    InputFileNames = Application.GetOpenFilename( _
                     MultiSelect:=True, _
                     FileFilter:="Excel Documents, *.xls;*.xlsx;*.xlsm", _
                     Title:="Select Workbooks for FTE Calculation")

    ' If InputFileNames is not an array of Variants then exit
    If VarType(InputFileNames) <> 8204 Then Exit Sub

    For Each Elem In InputFileNames
        Set wb = Workbooks.Open(Elem, ReadOnly:=True)
        
        ' Create Appointed EmployeeCollection
        Set ws = WBTools.GetSheetLike("*Appointed*", wb)
        If Not ws Is Nothing Then
            Set EC_Temp = New EmployeeCollection
            Set EC_Temp = EC_Temp.CreateEmployeeCollectionFromWorksheet(ws, False)
            EC_All.Concat EC_Temp
        End If
        Set EC_Temp = Nothing
        
        ' Create Hourly EmployeeCollection
        Set ws = WBTools.GetSheetLike("*Hourly*", wb)
        If Not ws Is Nothing Then
            Set EC_Temp = New EmployeeCollection
            Set EC_Temp = EC_Temp.CreateEmployeeCollectionFromWorksheet(ws, True)
            EC_All.Concat EC_Temp
        End If
        Set EC_Temp = Nothing
        
        wb.Close
    Next Elem
    
    ' filter for only the desired JobCodes
    For Each Elem In JobCodes
        EC_Filtered.Concat EC_All.Filter(JobCodeFilter:=Elem)
    Next Elem
    Set EC_All = Nothing
    
    Set Output = Workbooks.Add
    With Output
        .Title = "FTECalc Output " & GetTimestampStr()
        .Subject = "FTE"
        
        '@Ignore AssignmentNotUsed
        Set ws = .Worksheets.Item(1)
        
        .Worksheets.Add After:=ws, Count:=3
        
        ApplicationDisplayAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = ApplicationDisplayAlerts
    End With
    
    Set AC_Filtered_Merged = EC_Filtered.MergeAllEmployees().ToArrayContainer()
    With Output.Worksheets.Item(1)
        .Name = "FTE Summary"
        .Range( _
        "A1:" _
      & WBTools.GetColumnLetterByNumber(AC_Filtered_Merged.Columns) _
      & Trim$(Str$(AC_Filtered_Merged.Rows)) _
        ) = AC_Filtered_Merged.Data
    End With
    Set AC_Filtered_Merged = Nothing
    
    Set AC_GroupByDeptID = _
                         EC_Filtered.MergeAllEmployeesOnDeptID() _
                         .ToArrayContainer(IncludeJobCode:=False)
    With Output.Worksheets.Item(2)
        .Name = "GrpBy DeptID"
        .Range( _
        "A1:" _
      & WBTools.GetColumnLetterByNumber(AC_GroupByDeptID.Columns) _
      & Trim$(Str$(AC_GroupByDeptID.Rows)) _
        ) = AC_GroupByDeptID.Data
    End With
    Set AC_GroupByDeptID = Nothing
    
    Set AC_GroupByJobCode = _
                          EC_Filtered.MergeAllEmployeesOnJobCode() _
                          .ToArrayContainer(IncludeDeptID:=False)
    Set EC_Filtered = Nothing
    With Output.Worksheets.Item(3)
        .Name = "GrpBy JobCode"
        .Range( _
        "A1:" _
      & WBTools.GetColumnLetterByNumber(AC_GroupByJobCode.Columns) _
      & Trim$(Str$(AC_GroupByJobCode.Rows)) _
        ) = AC_GroupByJobCode.Data
    End With
    Set AC_GroupByJobCode = Nothing

    ' save the workbook
    With Output
        Do
            OutputFileName = Application.GetSaveAsFilename( _
                             InitialFileName:="FTECalc_Output_" & GetTimestampStr() & ".xlsx", _
                             FileFilter:="Excel Files (*.xlsx),*.xlsx")
        Loop Until OutputFileName <> False
        
        .SaveAs FileName:=OutputFileName
        .Close
    End With

End Sub

Public Function UnixTime() As LongLong
    UnixTime = DateDiff("s", "1/1/1970 00:00:00", Now)
End Function


