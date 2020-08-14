VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFrontend 
   Caption         =   "HR Application"
   ClientHeight    =   4280
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   6315
   OleObjectBlob   =   "frmFrontend.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFrontend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private controller As csController

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdFind_Click()
    If controller Is Nothing Then
        Set controller = New csController
    End If
    controller.FindEmployee Me.txtId.value
End Sub

Private Sub cmdSave_Click()
    If controller Is Nothing Then
        Set controller = New csController
    End If
    controller.UpdateEmployee
End Sub

Private Sub UserForm_Initialize()
    If controller Is Nothing Then
        Set controller = New csController
    End If
    controller.GetFirstEmployee
End Sub


Private Sub UserForm_Terminate()
    Set controller = Nothing
End Sub
