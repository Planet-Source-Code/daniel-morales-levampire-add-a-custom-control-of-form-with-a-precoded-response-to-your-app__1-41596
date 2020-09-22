VERSION 5.00
Begin VB.MDIForm mfExeParent 
   BackColor       =   &H8000000C&
   Caption         =   "Parent form From Exe"
   ClientHeight    =   5685
   ClientLeft      =   2085
   ClientTop       =   1995
   ClientWidth     =   6585
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnuFormExe 
      Caption         =   "Load form From Exe"
   End
   Begin VB.Menu mnuFormDll 
      Caption         =   "Load Form from Dll"
   End
   Begin VB.Menu mnuClose 
      Caption         =   "Close"
   End
End
Attribute VB_Name = "mfExeParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuClose_Click()
    'Terminates program execution
    'For some odd reason it also terminates the vb application
    'if you find a reason for this please let me know
    End
End Sub

Private Sub mnuFormDll_Click()
    'opens the form from the dll
    obs.fLoad Me
End Sub

Private Sub mnuFormExe_Click()
    'opens local form
    fExeChild.Show
End Sub

