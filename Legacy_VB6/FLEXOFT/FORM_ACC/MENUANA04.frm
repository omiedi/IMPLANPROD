VERSION 5.00
Begin VB.Form MENUANA04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis Estadístico de Ventas"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   8520
      Left            =   3730
      ScaleHeight     =   8460
      ScaleWidth      =   915
      TabIndex        =   14
      Top             =   0
      Width           =   975
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   130
         Picture         =   "MENUANA04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   785
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   130
         Picture         =   "MENUANA04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   105
         Width           =   650
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   130
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3045
         Width           =   640
         Begin VB.Frame Frame14 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -420
            TabIndex        =   16
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "MENUANA04.frx":0454
         Top             =   3465
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Por Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   210
      TabIndex        =   7
      Top             =   2100
      Width           =   3270
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   8
         Top             =   945
         Width           =   2325
         Begin VB.CommandButton Command3 
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
            Left            =   0
            Picture         =   "MENUANA04.frx":2376
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Análisis de Entregas en Unidades por Artículo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Entregas en Unidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   10
            Top             =   105
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   11
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command8 
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
            Left            =   0
            Picture         =   "MENUANA04.frx":24C0
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Análisis de Ventas en Pesos por Artículo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Ventas Totales en Pesos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   13
            Top             =   50
            Width           =   1380
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Por Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3270
      Begin VB.PictureBox Picture14 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   1
         Top             =   945
         Width           =   2325
         Begin VB.CommandButton Command22 
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
            Left            =   0
            Picture         =   "MENUANA04.frx":27CA
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Análisis de Entregas por Cliente"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Entregas en Unidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   3
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture16 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   4
         Top             =   420
         Width           =   2325
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
            Left            =   0
            Picture         =   "MENUANA04.frx":2AD4
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Análisis de Ventas por Cliente"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Ventas Totales en Pesos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   6
            Top             =   105
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "MENUANA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   ANAENTCLI.Show 1
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   ANAENTCOD.Show 1
   Command3.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   ANAVENCLI.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   ANAVENCOD.Show 1
   Command8.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

