VERSION 5.00
Begin VB.Form fDllTestMDIChild 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MDIChild Form from Dll"
   ClientHeight    =   2190
   ClientLeft      =   6600
   ClientTop       =   4020
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2190
   ScaleWidth      =   5235
   Begin VB.CommandButton cMsgBox 
      Caption         =   "MsgBox"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "fDllTestMDIChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cMsgBox_Click()
    MsgBox "It just seems like it loaded correctly!!"
End Sub
