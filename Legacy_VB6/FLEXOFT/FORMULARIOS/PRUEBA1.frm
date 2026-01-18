VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Leer MOVISTO"
      Height          =   435
      Left            =   1155
      TabIndex        =   0
      Top             =   630
      Width           =   2430
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   
   TINI = Timer
   Call SELECHO("LIPRE001")
   'Call BLOQARCH("MOVISTO")
   'URE& = ULTREG&("MOVISTO")
   'For REG& = 1 To URE&
   '  XXX$ = REGLEIDO$("MOVISTO", REG&)
   'Next REG&
   'Call CIERRARCH("MOVISTO")
   MsgBox Timer - TINI
End Sub

Private Sub Form_Load()
   Call VERJOBID
   Call BLOQARCH("MOVISTO")
End Sub
