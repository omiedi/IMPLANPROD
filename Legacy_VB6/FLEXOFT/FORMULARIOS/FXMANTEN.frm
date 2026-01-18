VERSION 5.00
Begin VB.Form FXMANTEN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aviso de Mantenimiento"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar Mensaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1785
      TabIndex        =   0
      Top             =   1575
      Width           =   2325
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "FXMANTEN.frx":0000
      Top             =   525
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FLEXOFT entrará en Modo de Mantenimiento. Debe cerrar todas las Aplicaciones y salir de  FLEXOFT.Muchas Gracias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   945
      TabIndex        =   1
      Top             =   210
      Width           =   4215
   End
End
Attribute VB_Name = "FXMANTEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TXTX$

Private Sub Command1_Click()
   Timer1.Enabled = False
   Close
   End
End Sub

Private Sub Form_Load()
    RST$ = ""
    Randomize Timer
    For KL% = 1 To 6
10    PPX% = Int(Rnd * 255)
      If PPX% < 48 Or PPX% > 57 Then GoTo 10
      RST$ = RST$ + Chr$(PPX%)
    Next KL%
    '
    TXTX$ = Caption + " (" + RST$ + ")"
    Caption = TXTX$
    '
End Sub

Private Sub Timer1_Timer()
   Static VECES
   '
   VECES = VECES + 1
   If VECES > 10 Then
     Call Command1_Click
     Exit Sub
   End If
   '
   On Error Resume Next
   AppActivate Caption
   SendKeys Chr$(1)
   DoEvents
   On Error GoTo 0
End Sub

