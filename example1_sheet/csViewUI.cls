VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csViewUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author: Zaid Qureshi
'Website: datapluscode.com

Option Explicit

Private mwsUserInterface As Worksheet

Public Sub ClearEmployeeDisplay()
    Application.ScreenUpdating = False
    With mwsUserInterface
        .Range("B2").value = Empty
        .Range("B4").value = Empty
        .Range("B5").value = Empty
        .Range("B6").value = Empty
        .Range("B7").value = Empty
    End With
    Application.ScreenUpdating = True
End Sub

Public Sub DisplayEmployeeRecord(employee As csModelEmployee)
    Application.ScreenUpdating = False
    With mwsUserInterface
        .Range("B2").value = employee.employeeNo
        .Range("B4").value = employee.firstName
        .Range("B5").value = employee.lastName
        .Range("B6").value = employee.gender
        .Range("B7").value = employee.jobTitle
    End With
    Application.ScreenUpdating = True
End Sub

Public Function GetDisplayedEmployeeRecord() As csModelEmployee
    Set GetDisplayedEmployeeRecord = New csModelEmployee
    With mwsUserInterface
        GetDisplayedEmployeeRecord.employeeNo = .Range("B2").value
        GetDisplayedEmployeeRecord.firstName = .Range("B4").value
        GetDisplayedEmployeeRecord.lastName = .Range("B5").value
        GetDisplayedEmployeeRecord.gender = .Range("B6").value
        GetDisplayedEmployeeRecord.jobTitle = .Range("B7").value
    End With
End Function

Private Sub Class_Initialize()
    Set mwsUserInterface = ThisWorkbook.Sheets("FrontEnd")
End Sub

Private Sub Class_Terminate()
    Set mwsUserInterface = Nothing
End Sub
