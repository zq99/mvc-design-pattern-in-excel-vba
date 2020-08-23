VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtFrontEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private controller As csController

Public Sub btnFind()
    If controller Is Nothing Then
        Set controller = New csController
    End If
    controller.FindEmployee Me.Range("B2").value
End Sub

Public Sub btnUpdate()
    If controller Is Nothing Then
        Set controller = New csController
    End If
    controller.UpdateEmployee
End Sub

Private Sub Worksheet_Activate()
    Set controller = New csController
End Sub

Private Sub Worksheet_Deactivate()
    Set controller = Nothing
End Sub


