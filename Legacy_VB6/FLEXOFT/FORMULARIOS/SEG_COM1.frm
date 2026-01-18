VERSION 5.00
Begin VB.Form SEG_COM1 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimientos Comerciales"
   ClientHeight    =   7650
   ClientLeft      =   195
   ClientTop       =   390
   ClientWidth     =   10905
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7650
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "SALIR"
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "CONSULTAS"
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "NOVEDADES"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "CONTACTOS"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label0 
      BackStyle       =   0  'Transparent
      Caption         =   "    Sistema Modular Integrado de Procesamiento de Datos                              Linea FLEXOFT (r) - Version 9.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3465
      TabIndex        =   4
      Top             =   6930
      Width           =   7365
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   0
      Picture         =   "SEG_COM1.frx":0000
      Stretch         =   -1  'True
      Top             =   6825
      Width           =   10920
   End
End
Attribute VB_Name = "SEG_COM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  SEG_COM2.Show
End Sub

Private Sub Command2_Click()
  SEG_COM3.Show
End Sub

Private Sub Command4_Click()
  Close: End
End Sub

