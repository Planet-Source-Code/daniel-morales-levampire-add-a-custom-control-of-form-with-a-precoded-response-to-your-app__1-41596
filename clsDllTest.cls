VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDllTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Declare the Usercontrol Variable for Creation of the Control
Dim WithEvents ctlDynamic As VBControlExtender
Attribute ctlDynamic.VB_VarHelpID = -1

Public Sub fLoad(fMDIParent As Object)
Dim frm As New fDllTestChild
    'loads the form on memory
    Load frm
    'Sets the parent to be the MDI Form calling from the exe
    SetParent frm.hwnd, fMDIParent.hwnd
    'sets the scalemode to pixels since the movewindow function is in pixels
    'This can be ommited if you use the left, top, width, & height properties
    frm.ScaleMode = vbPixels
    'places the form inside its new parent
    MoveWindow frm.hwnd, 0, 0, frm.ScaleWidth, frm.ScaleHeight, True
    'shows the form
    frm.Show
End Sub

Public Sub cLoad(fParent As Object)
On Error Resume Next
'sets the control to be the user control created in the Ocx
Set ctlDynamic = fParent.Controls.Add("proOcxTest.butClose", "cClose", fParent) '1
'stablishes the control's properties
With ctlDynamic '2
    .Left = fParent.ScaleWidth - .Width - 120 '3
    .Top = fParent.ScaleHeight - .Height - 120 '4
    .Visible = True '7
End With '8
Exit Sub
'handles posible errors
err:
    MsgBox "Err: " & _
           "Description: " & err.Description & _
           " Number: " & err.Number & _
           " Source: " & err.Source & "."
    err.Clear
End Sub
