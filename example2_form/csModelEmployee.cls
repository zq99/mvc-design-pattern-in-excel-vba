VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csModelEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author: Zaid Qureshi
'Website: datapluscode.com

Option Explicit

Private mstrFirstName   As String
Private mstrLastName    As String
Private mstrGender      As String
Private mstrJobTitle    As String
Private mwsData         As Worksheet
Private mrngIds         As Range
Private mintEmployeeNo  As Integer

Public Property Get employeeNo() As Integer
    employeeNo = mintEmployeeNo
End Property

Public Property Let employeeNo(ByVal value As Integer)
    mintEmployeeNo = value
End Property

Public Property Get firstName() As String
    firstName = mstrFirstName
End Property

Public Property Let firstName(ByVal value As String)
    mstrFirstName = value
End Property

Public Property Get lastName() As String
    lastName = mstrLastName
End Property

Public Property Let lastName(ByVal value As String)
    mstrLastName = value
End Property

Public Property Get gender() As String
    gender = mstrGender
End Property

Public Property Let gender(ByVal value As String)
    mstrGender = value
End Property

Public Property Get jobTitle() As String
    jobTitle = mstrJobTitle
End Property

Public Property Let jobTitle(ByVal value As String)
    mstrJobTitle = value
End Property

Public Function UpdateEmployeeRecord(ByVal employee As csModelEmployee) As Boolean
    Dim rngFind As Range
    
    UpdateEmployeeRecord = False
    Set rngFind = mrngIds.Find(What:=employee.employeeNo, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngFind Is Nothing Then
        With mwsData
            .Cells(rngFind.Row, 1).value = employee.employeeNo
            .Cells(rngFind.Row, 2).value = employee.firstName
            .Cells(rngFind.Row, 3).value = employee.lastName
            .Cells(rngFind.Row, 4).value = employee.gender
            .Cells(rngFind.Row, 5).value = employee.jobTitle
        End With
        UpdateEmployeeRecord = True
    End If
    Set rngFind = Nothing
End Function


Public Function GetFirstEmployeeRecord() As Boolean
    If GetEmployeeRecord(1) Then
        GetFirstEmployeeRecord = True
    Else
        GetFirstEmployeeRecord = False
    End If
End Function

Public Function GetEmployeeRecord(ByVal intId As Integer) As Boolean
    Dim rngFind As Range
    
    GetEmployeeRecord = False
    Set rngFind = mrngIds.Find(What:=intId, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngFind Is Nothing Then
        With mwsData
            Me.employeeNo = .Cells(rngFind.Row, 1).value
            Me.firstName = .Cells(rngFind.Row, 2).value
            Me.lastName = .Cells(rngFind.Row, 3).value
            Me.gender = .Cells(rngFind.Row, 4).value
            Me.jobTitle = .Cells(rngFind.Row, 5).value
        End With
        GetEmployeeRecord = True
    End If
    Set rngFind = Nothing
End Function

Private Sub Class_Initialize()
    Dim llastRow As Long
    Set mwsData = ThisWorkbook.Worksheets("Data")
    llastRow = mwsData.Cells(mwsData.Rows.Count, "A").End(xlUp).Row
    Set mrngIds = mwsData.Range("A1:A" & llastRow)
End Sub

Private Sub Class_Terminate()
    Set mwsData = Nothing
    Set mrngIds = Nothing
End Sub
