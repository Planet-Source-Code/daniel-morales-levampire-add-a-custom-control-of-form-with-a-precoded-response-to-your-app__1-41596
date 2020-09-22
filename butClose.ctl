VERSION 5.00
Begin VB.UserControl butClose 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1320
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleWidth      =   1320
   Begin VB.CommandButton cClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "butClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub cClose_Click()
    'The Whole purpose of this post....
    Unload UserControl.Parent
End Sub

Private Sub UserControl_Initialize()
    'Set the usercontrol's Dimentions
    Height = 375
    Width = 1095
End Sub

Private Sub UserControl_Resize()
    'Resice the CommandButton inside the UC when the UC resizes...
    cClose.Height = Height
    cClose.Width = Width
End Sub
