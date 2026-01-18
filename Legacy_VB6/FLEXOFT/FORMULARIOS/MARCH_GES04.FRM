VERSION 5.00
Begin VB.Form MARCH_GES04 
   BackColor       =   &H00D0F0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archivos Maestros - Actualización y Consulta"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404080&
      Height          =   4335
      Left            =   4800
      ScaleHeight     =   4275
      ScaleWidth      =   915
      TabIndex        =   19
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   22
         Top             =   3000
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   23
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command29 
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
         Picture         =   "MARCH_GES04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   820
         Width           =   600
      End
      Begin VB.CommandButton Command24 
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
         Height          =   600
         Left            =   120
         Picture         =   "MARCH_GES04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   210
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -320
         Picture         =   "MARCH_GES04.frx":0454
         Top             =   3320
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0FF&
      Caption         =   "Actualización Bases de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   4425
      Begin VB.Frame Frame4 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Materiales, Fórmulas y Rutas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   4005
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   10
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command1 
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
               Picture         =   "MARCH_GES04.frx":2376
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Maestro de Artí- culos de Stock"
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
               TabIndex        =   12
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   2
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command4 
            Caption         =   "Inventarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   8
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   7
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Críticos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Estadísticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   4
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Análisis A-B-C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   3
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Clientes / Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   210
         TabIndex        =   1
         Top             =   1800
         Width           =   4005
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   16
            Top             =   840
            Width           =   2325
            Begin VB.CommandButton Command18 
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
               Left            =   -20
               Picture         =   "MARCH_GES04.frx":2680
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Clientes en Cuenta Corriente"
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
               TabIndex        =   18
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   13
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command19 
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
               Left            =   -20
               Picture         =   "MARCH_GES04.frx":2AC2
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Proveedores y Subcontratistas"
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
               TabIndex        =   15
               Top             =   0
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Archi 
      Caption         =   "Archivos Maestros"
      Begin VB.Menu Mastok 
         Caption         =   "Materiales, Fórmulas, Rutas"
      End
      Begin VB.Menu PLH1 
         Caption         =   "-"
      End
      Begin VB.Menu Amacli 
         Caption         =   "Clientes - Cuentas Corrientes"
      End
      Begin VB.Menu Proveed 
         Caption         =   "Proveedores / Contratistas"
      End
      Begin VB.Menu PLH2 
         Caption         =   "-"
      End
      Begin VB.Menu List 
         Caption         =   "Listados"
      End
      Begin VB.Menu plh3 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Tables 
      Caption         =   "Tablas Auxiliares"
   End
   Begin VB.Menu Help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "MARCH_GES04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Amacli_Click()
   Call Command18_Click
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   If EXISTE%("AMASTO02.EXE") > 0 Then
       Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
     Else
       AMAS_GES04.Show 1
   End If
   Command1.Enabled = True
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    If EXISTE%("AMACLI02.EXE") > 0 Then
       Call CONECRUN("AMACLI02", "Maestro de Clientes y Deudores")
     Else
       ACLI_GES04.Show 1
    End If
    Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    ACRE_GES04.Show 1
    'Call CONECRUN("AMAPRO", "Actualización de Base de Datos de Proveedores")
    Command19.Enabled = True
End Sub

Private Sub Command24_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("ARCHIMA")
End Sub

Private Sub Exit_Click()
   Call Command24_Click
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EJER% = 1
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   If Left$(APLINVO$, 6) = "AMACLI" Then
       ACLI_GES.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 6) = "AMASTO" Then
       AMASTOBAS.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 6) = "AMAPRO" Then
       Call Command19_Click  ' Actualización de Base de Datos de Proveedores
       Call FINAL("")
   End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Private Sub Help_Click()
   Call Command29_Click
End Sub

Private Sub Mastok_Click()
   Call Command1_Click
End Sub

Private Sub Proveed_Click()
   Call Command19_Click
End Sub
