VERSION 5.00
Begin VB.Form CONCARMA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C.R.P. - Programación de Carga de Máquinas"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5400
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Operaciones Pendientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   210
      TabIndex        =   2
      Top             =   105
      Width           =   4005
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2730
         TabIndex        =   6
         Top             =   1155
         Width           =   2790
         Begin VB.CommandButton Command15 
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
            Picture         =   "CONCARMA.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Operaciones Pendientes por Máquina o Equipo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Máquina o Equi- po de Producción"
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
            Left            =   735
            TabIndex        =   7
            Top             =   105
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2685
         TabIndex        =   3
         Top             =   630
         Width           =   2745
         Begin VB.CommandButton Command26 
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
            Picture         =   "CONCARMA.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Operaciones Pendientes por Orden de Fabricación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Pendientes por Or- den de Fabricación"
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
            Left            =   735
            TabIndex        =   5
            Top             =   60
            Width           =   1800
         End
      End
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Direct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4410
      Picture         =   "CONCARMA.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
      Top             =   945
      Width           =   750
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
      Height          =   645
      Left            =   4410
      Picture         =   "CONCARMA.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3820
      Picture         =   "CONCARMA.frx":0A68
      Top             =   1695
      Width           =   2010
   End
End
Attribute VB_Name = "CONCARMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   Screen.MousePointer = 11
   FIN& = ULTREG("OPERFAP")
   For U& = 1 To FIN&
     XX$ = REGLEIDO$("OPERFAP", U&)
     TITO = 60 * CVS(Mid$(XX$, 131, 4))
     Call REPLA(XX$, MKS$(TITO), 131, 4)
     Call GRAREG("OPERFAH", XX$, U&)
   Next U&
   Call SETULTREG("OPERFAH", FIN&)
   Call CIERRARCH("OPERFAP")
   Screen.MousePointer = 1
   Call FINAL("*SECAREQ")
   Command15.Enabled = True
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   Call FINAL("*SECAROF")
   Command26.Enabled = True
End Sub

Private Sub Command34_Click()
    Call ACCDIR("CARGMAQ")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub

