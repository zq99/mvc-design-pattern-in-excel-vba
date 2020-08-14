VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csViewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author: Zaid Qureshi
'Website: datapluscode.com

Option Explicit

Public Sub ClearEmployeeDisplay()
    Application.ScreenUpdating = False
    With frmFrontend
        .txtId.value = Empty
        .txtFirstName.value = Empty
        .txtLastName.value = Empty
        .txtGender.value = Empty
        .txtJobTitle.value = Empty
    End With
    Application.ScreenUpdating = True
End Sub

Public Sub DisplayEmployeeRecord(employee As csModelEmployee)
    Application.ScreenUpdating = False
    With frmFrontend
        .txtId.value = employee.employeeNo
        .txtFirstName.value = employee.firstName
        .txtLastName.value = employee.lastName
        .txtGender.value = employee.gender
        .txtJobTitle.value = employee.jobTitle
    End With
    Application.ScreenUpdating = True
End Sub

Public Function GetDisplayedEmployeeRecord() As csModelEmployee
    Set GetDisplayedEmployeeRecord = New csModelEmployee
    With frmFrontend
        GetDisplayedEmployeeRecord.employeeNo = .txtId.value
        GetDisplayedEmployeeRecord.firstName = .txtFirstName.value
        GetDisplayedEmployeeRecord.lastName = .txtLastName.value
        GetDisplayedEmployeeRecord.gender = .txtGender.value
        GetDisplayedEmployeeRecord.jobTitle = .txtJobTitle.value
    End With
End Function

