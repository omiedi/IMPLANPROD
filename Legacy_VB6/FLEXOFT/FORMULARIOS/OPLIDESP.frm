VERSION 5.00
Begin VB.Form OPLIDESP 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado Pedidos Pendientes de Despacho"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Controles No Visibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   105
      TabIndex        =   16
      Top             =   3885
      Visible         =   0   'False
      Width           =   3585
      Begin VB.Label AGRUPA 
         Height          =   225
         Left            =   1890
         TabIndex        =   18
         Top             =   315
         Width           =   1485
      End
      Begin VB.Label OPIMPRE 
         Height          =   225
         Left            =   210
         TabIndex        =   17
         Top             =   315
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   5160
      Left            =   3780
      ScaleHeight     =   5100
      ScaleWidth      =   1215
      TabIndex        =   10
      Top             =   -40
      Width           =   1275
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "OPLIDESP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Aprueba - Emite Listado"
         Top             =   2520
         Width           =   600
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   13
         Top             =   2205
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "OPLIDESP.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   945
         Width           =   600
      End
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "OPLIDESP.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   315
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPLIDESP.frx":075E
         Top             =   3270
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Pedidos Preparados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   105
      TabIndex        =   8
      Top             =   2960
      Width           =   3585
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Netear Pedidos ya Preparados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   315
         TabIndex        =   9
         Top             =   315
         Value           =   1  'Checked
         Width           =   2955
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Agrupamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   105
      TabIndex        =   1
      Top             =   2100
      Width           =   3585
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Por Artículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1890
         TabIndex        =   3
         Top             =   315
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Por Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   315
         TabIndex        =   2
         Top             =   315
         Value           =   -1  'True
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Opciones de Impresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   3585
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Para Fabricar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   735
         TabIndex        =   7
         Top             =   1365
         Width           =   2500
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Pedidos en Preparación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   735
         TabIndex        =   6
         Top             =   1050
         Width           =   2500
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listos para Despachar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   735
         TabIndex        =   5
         Top             =   735
         Width           =   2500
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Todos los Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   735
         TabIndex        =   4
         Top             =   420
         Value           =   -1  'True
         Width           =   2500
      End
   End
End
Attribute VB_Name = "OPLIDESP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
   OPIMPRE = ""
   Hide
End Sub

Private Sub Command2_Click()
   '
   If Option3.Value = True Then
       OPIMPRE = "0"
     ElseIf Option4.Value = True Then
       OPIMPRE = "1"
     ElseIf Option5.Value = True Then
       OPIMPRE = "2"
     ElseIf Option6.Value = True Then
       OPIMPRE = "3"
   End If
   '
   If Option1.Value = True Then
       AGRUPA = "1"
     ElseIf Option2.Value = True Then
       AGRUPA = "2"
   End If
   '
   Hide
   '
End Sub

Private Sub Form_Load()
   Height = 4200
End Sub
