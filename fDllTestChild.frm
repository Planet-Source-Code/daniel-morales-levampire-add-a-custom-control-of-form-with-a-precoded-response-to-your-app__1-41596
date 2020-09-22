VERSION 5.00
Begin VB.Form fDllTestChild 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form From Dll"
   ClientHeight    =   2145
   ClientLeft      =   2610
   ClientTop       =   3345
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   5220
   Begin VB.CommandButton cMsgBox 
      Caption         =   "MsgBox"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "fDllTestChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cMsgBox_Click()
    'this message is just so you see it worked, lol
    MsgBox "It just seems like it loaded correctly!!"
End Sub
