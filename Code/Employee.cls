VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder(HCM)
'@ModuleDescription(An object representing an employee record.)

Private Type TEmployee
    EmplID As String
    Name As String
    Department As String
    JobCode As String
    IsCanceled As Boolean
    IsHourly As Boolean
    ' this.hoursWorked will be used as a Dictionary
    hoursWorked As Object
    Source As String
End Type

Private this As TEmployee

Private Sub Class_Initialize()
    Set this.hoursWorked = CreateObject("Scripting.Dictionary")
    this.hoursWorked.Add "01A", 0
    this.hoursWorked.Add "01B", 0
    this.hoursWorked.Add "02A", 0
    this.hoursWorked.Add "02B", 0
    this.hoursWorked.Add "03A", 0
    this.hoursWorked.Add "03B", 0
    this.hoursWorked.Add "04A", 0
    this.hoursWorked.Add "04B", 0
    this.hoursWorked.Add "05A", 0
    this.hoursWorked.Add "05B", 0
    this.hoursWorked.Add "06A", 0
    this.hoursWorked.Add "06B", 0
    this.hoursWorked.Add "07A", 0
    this.hoursWorked.Add "07B", 0
    this.hoursWorked.Add "08A", 0
    this.hoursWorked.Add "08B", 0
    this.hoursWorked.Add "09A", 0
    this.hoursWorked.Add "09B", 0
    this.hoursWorked.Add "10A", 0
    this.hoursWorked.Add "10B", 0
    this.hoursWorked.Add "11A", 0
    this.hoursWorked.Add "11B", 0
    this.hoursWorked.Add "12A", 0
    this.hoursWorked.Add "12B", 0
    this.hoursWorked.Add "OTH", 0 ' For non-Pay-Period-Specific cases
End Sub

Public Property Get EmplID() As String
    EmplID = this.EmplID
End Property

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Get Department() As String
    Department = this.Department
End Property

Public Property Get IsCanceled() As Boolean
    IsCanceled = this.IsCanceled
End Property

Public Property Get IsHourly() As Boolean
    IsHourly = this.IsHourly
End Property

Public Property Get JobCode() As String
    JobCode = this.JobCode
End Property

Public Property Let EmplID(ByVal eid As String)
    this.EmplID = Trim$(eid)
End Property

Public Property Let Name(ByVal Name As String)
    ' TODO: remove all trailing whitespace and whitespace after commas
    this.Name = Trim$(Name)
End Property

Public Property Let Department(ByVal dept As String)
    this.Department = Trim$(dept)
End Property

Public Property Let IsCanceled(ByVal IsCanceled As Boolean)
    this.IsCanceled = IsCanceled
End Property

Public Property Let IsHourly(ByVal IsHourly As Boolean)
    this.IsHourly = IsHourly
End Property

Public Property Let JobCode(ByVal JobCode As String)
    this.JobCode = Trim$(JobCode)
End Property

Public Property Let Source(ByVal SourceString As String)
    this.Source = Trim$(SourceString)
End Property

Public Property Get Source() As String
    Source = this.Source
End Property

'@Description(Returns the sum of all hours worked, or the hours worked for a specific period if the optional key is used.)
Public Property Get hoursWorked(Optional ByVal Key As Variant = vbNullString) As Single
    '    If Main.DEBUG_ON Then Debug.Print "Employee Property Get hoursWorked(" & key & ")"
    Dim hours As Single
    hours = 0
    
    If Key = vbNullString Then
        Dim PayPeriod As Variant
        
        For Each PayPeriod In this.hoursWorked.Items
            hours = hours + PayPeriod
            '            If Main.DEBUG_ON Then Debug.Print "hours = " & hours & " + " & payPeriod & " (" & this.name & " - " & e_key & ")"
        Next PayPeriod
    Else
        hours = this.hoursWorked.Item(Key)
        '        If Main.DEBUG_ON Then Debug.Print "hours for key " & key & ": " & hours & " (" & this.name & " - " & e_key & ")"
    End If
    
    hoursWorked = hours
End Property

'@Description(Assigns the number of hours worked to the specified period.)
Public Property Let hoursWorked(ByVal Key As Variant, ByVal Value As Single)
    this.hoursWorked(Key) = Value
End Property

'@Obsolete(Use hoursWorked (Let) instead.)
'@Description(Assigns the number of hours worked to the specified period.)
Public Sub Set_hoursWorked(ByVal Key As Variant, ByVal Value As Single)
    '@Ignore IndexedUnboundDefaultMemberAccess
    this.hoursWorked(Key) = Value
End Sub

Public Function eKey() As String
    eKey = this.EmplID & "*" & this.Department & "*" & this.JobCode
End Function

Public Sub Print_Employee()
    Debug.Print this.EmplID & vbTab & this.Name & vbTab & this.Department & vbTab & this.JobCode
End Sub

Public Function MergeEmployeeHours(ByVal employeeToBeMerged As Employee, Optional ByVal Department As String = vbNullString, Optional ByVal JobCode As String = vbNullString, Optional ByVal SourceString As String = vbNullString) As Employee
    '    If Main.DEBUG_ON Then Debug.Print "MergeEmployeeHours: " & name & eKey & " + " & employeeToBeMerged.name & employeeToBeMerged.eKey
    
    Dim employeeMerged As Employee
    Set employeeMerged = New Employee
    
    If SourceString = vbNullString Then SourceString = this.Source & "; " & employeeToBeMerged.Source
    
    If employeeToBeMerged.EmplID <> this.EmplID Then
        Debug.Print "Cannot merge records of different employees. Check EmplIDs."
        Set MergeEmployeeHours = Nothing
    Else
        employeeMerged.Name = this.Name
        employeeMerged.EmplID = this.EmplID
        employeeMerged.Department = Department
        employeeMerged.JobCode = JobCode
        employeeMerged.Source = SourceString
        Dim Key As Variant
        For Each Key In this.hoursWorked.Keys
            '            If Main.DEBUG_ON Then Debug.Print "Merging " & key & ": " & hoursWorked(key) & " + " & employeeToBeMerged.hoursWorked(key)
            employeeMerged.hoursWorked(Key) = (this.hoursWorked(Key) + employeeToBeMerged.hoursWorked(Key))
        Next Key
    End If
    
    Set MergeEmployeeHours = employeeMerged
End Function

Public Function Keys() As Variant
    Keys = this.hoursWorked.Keys
End Function

Public Function Copy() As Employee
    Dim employeeCopy As Employee
    Set employeeCopy = New Employee

    employeeCopy.EmplID = this.EmplID
    employeeCopy.Name = this.Name
    employeeCopy.Department = this.Department
    employeeCopy.JobCode = this.JobCode
    employeeCopy.Source = this.Source
    employeeCopy.IsHourly = this.IsHourly
    
    Dim Key As Variant
    For Each Key In Keys
        employeeCopy.hoursWorked(Key) = this.hoursWorked(Key)
    Next Key
    
    Set Copy = employeeCopy
End Function

