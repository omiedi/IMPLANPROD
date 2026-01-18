VERSION 5.00
Begin VB.Form BLACKSCREEN 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8700
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "BLACKSCREEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
   If FORDEMO.Visible = True Then
     FORDEMO.Show
   End If
End Sub

Private Sub Form_Load()
   If Width < Screen.Width Then
      Width = Screen.Width
   End If
   If Height < Screen.Height - 300 Then
      Height = Screen.Height - 300
   End If
End Sub
