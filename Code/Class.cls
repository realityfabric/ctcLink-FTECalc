VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed, IndexedUnboundDefaultMemberAccess, UDTMemberNotUsed
'@Folder("Archive")
Option Explicit

Private Type TClass
    ClassBegin As Date
    ClassEnd As Date
    ClassNumber As String
    ComboCode As String
    ' this.ContactHours will be used as a Dictionary
    ContactHours As Object
    CourseDescription As String
    CourseID As String                           ' Subject + Catalog
    DepartmentDescription As String
    DeptID As String
    EmplID As String
    EmplName As String
    '@Ignore IntegerDataType
    EmplRcd As Integer
    FTE As Single
    IsCanceled As Boolean
    IsLab As Boolean                             ' True: Lab, False: Lecture
    JobCode As String
    PayRate As Single
    '@Ignore IntegerDataType
    Units As Integer
End Type

Private this As TClass

Private Sub Class_Initialize()
    Set this.ContactHours = CreateObject("Scripting.Dictionary")
    this.ContactHours.Add "01A", 0
    this.ContactHours.Add "01B", 0
    this.ContactHours.Add "02A", 0
    this.ContactHours.Add "02B", 0
    this.ContactHours.Add "03A", 0
    this.ContactHours.Add "03B", 0
    this.ContactHours.Add "04A", 0
    this.ContactHours.Add "04B", 0
    this.ContactHours.Add "05A", 0
    this.ContactHours.Add "05B", 0
    this.ContactHours.Add "06A", 0
    this.ContactHours.Add "06B", 0
    this.ContactHours.Add "07A", 0
    this.ContactHours.Add "07B", 0
    this.ContactHours.Add "08A", 0
    this.ContactHours.Add "08B", 0
    this.ContactHours.Add "09A", 0
    this.ContactHours.Add "09B", 0
    this.ContactHours.Add "10A", 0
    this.ContactHours.Add "10B", 0
    this.ContactHours.Add "11A", 0
    this.ContactHours.Add "11B", 0
    this.ContactHours.Add "12A", 0
    this.ContactHours.Add "12B", 0
    this.ContactHours.Add "OTH", 0               ' For non-Pay-Period-Specific cases
End Sub

Public Property Get ClassBegin() As Date
    ClassBegin = this.ClassBegin
End Property

Public Property Let ClassBegin(ByVal StartDate As Date)
    this.ClassBegin = StartDate
End Property

Public Property Get ClassEnd() As Date
    ClassEnd = this.ClassEnd
End Property

Public Property Let ClassEnd(ByVal EndDate As Date)
    this.ClassEnd = EndDate
End Property

Public Property Get ClassNumber() As String
    ClassNumber = this.ClassNumber
End Property

Public Property Let ClassNumber(ByVal Number As String)
    this.ClassNumber = Number
End Property

Public Property Get ComboCode() As String
    ComboCode = this.ComboCode
End Property

Public Property Let ComboCode(ByVal CombinationCode As String)
    this.ComboCode = CombinationCode
End Property

Public Property Get ContactHours(Optional ByVal PayPeriod As String = vbNullString) As Single
    If PayPeriod <> vbNullString Then
        ContactHours = this.ContactHours.Item(PayPeriod)
    Else
        Dim Period As Variant
        ContactHours = 0
        For Each Period In this.ContactHours
            ContactHours = ContactHours + this.ContactHours(Period)
        Next Period
    End If
End Property

Public Property Let ContactHours(ByVal PayPeriod As String, ByVal Hours As Single)
    this.ContactHours.Item(PayPeriod) = Hours
End Property

'  ContactHours As Integer
'  CourseDescription As String
'  CourseID As String ' Subject + Catalog
'  DepartmentDescription As String
'\  DeptID As String
' EmplID As String
' EmplName As String
' EmplRcd As Integer
' FTE As Single
' IsCanceled As Boolean
' IsLab As Boolean ' True: Lab, False: Lecture
' JobCode As String
' PayRate As Single
' Units As Integer
