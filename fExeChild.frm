VERSION 5.00
Begin VB.Form fExeChild 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form From Exe"
   ClientHeight    =   2070
   ClientLeft      =   2265
   ClientTop       =   3810
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2070
   ScaleWidth      =   4860
End
Attribute VB_Name = "fExeChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'loads the control from the ocx with the class from the dll
    obs.cLoad Me
End Sub
