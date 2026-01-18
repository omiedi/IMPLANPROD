VERSION 5.00
Begin VB.Form FORM5 
   Caption         =   "PRUEBA5"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   750
      Left            =   945
      TabIndex        =   0
      Top             =   1890
      Width           =   2325
   End
End
Attribute VB_Name = "FORM5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   On Error Resume Next
   AppActivate "PRUEBA6"
   SendKeys Chr$(31)
   On Error GoTo 0
End Sub

Private Sub Form_Activate()
   'If WindowState <> 0 Then
   '   WindowState = 0
   'End If
   'MsgBox "ME ACTIVE 5"
   'DoEvents
End Sub

Private Sub Form_Click()
   'WindowState = 0
   'MsgBox "CLICK"
End Sub

Private Sub Form_GotFocus()
   'WindowState = 0
   'MsgBox "GOTFOCUS"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 31 Then
     While WindowState <> 0
       WindowState = 0
       DoEvents
     Wend
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   'WindowState = 0
   'MsgBox "KEYPRESS"
End Sub
