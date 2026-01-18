VERSION 5.00
Begin VB.Form FORVELOCK 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FORVELOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   '
   ANAME$ = Command$
   MsgBox ANAME$
   N1% = FreeFile
   On Error Resume Next
   Open ANAME$ For Random Access Read Write Lock Write As #N1% Len = 128
   MsgBox Err
   On Error GoTo 0
   Close: End
End Sub
