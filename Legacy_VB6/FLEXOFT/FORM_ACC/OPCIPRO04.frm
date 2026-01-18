VERSION 5.00
Begin VB.Form OPCIPRO04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analisis de Productividad - Opciones"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Equipos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1155
      Picture         =   "OPCIPRO04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Equipos de Producción"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rutas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2100
      Picture         =   "OPCIPRO04.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Hojas de Ruta - Secuencia Operaciones Standard"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Operar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   210
      Picture         =   "OPCIPRO04.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Padron de Operarios"
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   2700
      Picture         =   "OPCIPRO04.frx":091E
      Top             =   780
      Width           =   1515
   End
End
Attribute VB_Name = "OPCIPRO04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Command1.Enabled = False
   If EXISTE%("AMASTO04.EXE") > 0 Then
        Call CONECRUN("AMASTO04", "Maestro de Artículos de Stock")
      Else
        Call CONECRUN("HORUTAS", "Hojas de Ruta - Secuencia de Operaciones")
   End If
   Command1.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   Call FINAL("*PADRON")
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call FINAL("*PADMAQ")
   Command2.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call INFOPER
   Command3.Enabled = True
End Sub
