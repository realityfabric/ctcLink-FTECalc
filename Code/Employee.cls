VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FTECalc")
Option Explicit

Private Type TEmployee
    EmplID As String
    Name As String
    DeptID As String
    JobCode As String
    HoursWorked As Object                        ' hoursWorked will be used as a Dictionary
End Type

Private this As TEmployee

Private Sub Class_Initialize()
    HoursWorked_Initialize
End Sub

Private Sub HoursWorked_Initialize()
    Set this.HoursWorked = CreateObject("Scripting.Dictionary")
    this.HoursWorked.Add "01A", 0
    this.HoursWorked.Add "01B", 0
    this.HoursWorked.Add "02A", 0
    this.HoursWorked.Add "02B", 0
    this.HoursWorked.Add "03A", 0
    this.HoursWorked.Add "03B", 0
    this.HoursWorked.Add "04A", 0
    this.HoursWorked.Add "04B", 0
    this.HoursWorked.Add "05A", 0
    this.HoursWorked.Add "05B", 0
    this.HoursWorked.Add "06A", 0
    this.HoursWorked.Add "06B", 0
    this.HoursWorked.Add "07A", 0
    this.HoursWorked.Add "07B", 0
    this.HoursWorked.Add "08A", 0
    this.HoursWorked.Add "08B", 0
    this.HoursWorked.Add "09A", 0
    this.HoursWorked.Add "09B", 0
    this.HoursWorked.Add "10A", 0
    this.HoursWorked.Add "10B", 0
    this.HoursWorked.Add "11A", 0
    this.HoursWorked.Add "11B", 0
    this.HoursWorked.Add "12A", 0
    this.HoursWorked.Add "12B", 0
    this.HoursWorked.Add "OTH", 0                ' For non-Pay-Period-Specific cases
End Sub

Public Property Get EmplID() As String
    EmplID = this.EmplID
End Property

Public Property Let EmplID(ByVal Value As String)
    this.EmplID = Trim$(Value)
End Property

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal Value As String)
    this.Name = Trim$(Value)
End Property

Public Property Get DeptID() As String
    DeptID = this.DeptID
End Property

Public Property Let DeptID(ByVal Value As String)
    this.DeptID = Trim$(Value)
End Property

Public Property Get JobCode() As String
    JobCode = this.JobCode
End Property

Public Property Let JobCode(ByVal Value As String)
    this.JobCode = Trim$(Value)
End Property

Public Property Get HoursWorked(Optional ByVal PayPeriod As String = vbNullString) As Long
    Dim Hours As Long
    Hours = 0
    If PayPeriod = vbNullString Then
        Dim Period As Variant
        For Each Period In this.HoursWorked
            Hours = Hours + this.HoursWorked.Item(Period)
        Next Period
    Else
        Hours = this.HoursWorked.Item(PayPeriod)
    End If
    
    HoursWorked = Hours
End Property

Public Property Let HoursWorked( _
       ByVal PayPeriod As String _
       , ByVal Value As Long)
    this.HoursWorked.Item(PayPeriod) = Value
End Property

Public Function fte() As Single
    fte = Round(HoursWorked() * 100 / 198, 2)
End Function

'@Description "Returns Employee if success, Nothing if fail"
'@Ignore UseMeaningfulName
Public Function Merge(ByVal Employee2 As Employee _
                      , Optional ByVal PreserveDeptID As Boolean = True _
                       , Optional ByVal PreserveJobCode As Boolean = True _
                        ) As Employee
Attribute Merge.VB_Description = "Returns Employee if success, Nothing if fail"
    Dim MergedEmployee As Employee
    
    Set MergedEmployee = New Employee
    
    If this.EmplID <> Employee2.EmplID Then
        Set MergedEmployee = Nothing
    ElseIf PreserveDeptID And this.DeptID <> Employee2.DeptID Then
        Set MergedEmployee = Nothing
    ElseIf PreserveJobCode And this.JobCode <> Employee2.JobCode Then
        Set MergedEmployee = Nothing
    Else
        MergedEmployee.EmplID = this.EmplID
        MergedEmployee.Name = this.Name
        
        If PreserveDeptID Then
            MergedEmployee.DeptID = this.DeptID
        Else
            MergedEmployee.DeptID = "*"
        End If
        
        If PreserveJobCode Then
            MergedEmployee.JobCode = this.JobCode
        Else
            MergedEmployee.JobCode = "*"
        End If
        
        Dim Key As Variant
        For Each Key In this.HoursWorked.Keys
            MergedEmployee.HoursWorked(Key) = this.HoursWorked.Item(Key) + Employee2.HoursWorked(Key)
        Next Key
    End If
    
    Set Merge = MergedEmployee
End Function
