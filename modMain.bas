Attribute VB_Name = "modMain"
'I took most of this codes from PSC so credit to whomever wrote them
'The purpose of this aplication is to explain how you can custumize controls
'with their own code inside an OCX and the call them from a Dll into your application
'See if you can find a use for it, I already got mine...
'HINT: To be able to "Configure" personalized applications from a mother application, so you
'don't have to code things all over again every tyme you want to write a new program...
'if I didn't Explain myself, I believe you can figure it out....

Option Explicit
Public obs As New clsDllTest

Public Sub Main()
    'Show the MDI Form
    mfExeParent.Show
End Sub

