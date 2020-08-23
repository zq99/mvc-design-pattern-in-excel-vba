VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author: Zaid Qureshi
'Website: datapluscode.com

Option Explicit

Private mobjEmployee        As csModelEmployee
Private mobjUserInterface   As csViewForm

Private Sub Class_Initialize()
    Set mobjEmployee = New csModelEmployee
    Set mobjUserInterface = New csViewForm
End Sub

Private Sub Class_Terminate()
    Set mobjEmployee = Nothing
    Set mobjUserInterface = Nothing
End Sub

Public Sub UpdateEmployee()
    mobjEmployee.UpdateEmployeeRecord mobjUserInterface.GetDisplayedEmployeeRecord
End Sub

Public Sub FindEmployee(ByVal employeeNo As Integer)
    If mobjEmployee.GetEmployeeRecord(employeeNo) Then
        mobjUserInterface.ClearEmployeeDisplay
        mobjUserInterface.DisplayEmployeeRecord mobjEmployee
    Else
        mobjUserInterface.ClearEmployeeDisplay
    End If
End Sub

Public Sub GetFirstEmployee()
    If mobjEmployee.GetFirstEmployeeRecord Then
        mobjUserInterface.ClearEmployeeDisplay
        mobjUserInterface.DisplayEmployeeRecord mobjEmployee
    Else
        mobjUserInterface.ClearEmployeeDisplay
    End If
End Sub
