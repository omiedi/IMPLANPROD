VERSION 5.00
Begin VB.Form OPREACH04 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Configuración"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404000&
      Height          =   8415
      Left            =   4725
      ScaleHeight     =   8355
      ScaleWidth      =   1005
      TabIndex        =   18
      Top             =   -105
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   22
         Top             =   1995
         Width           =   640
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
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "OPREACH04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   420
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Check"
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
         Left            =   105
         Picture         =   "OPREACH04.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8730
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "OPREACH04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ayuda en Línea"
         Top             =   1050
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPREACH04.frx":075E
         Top             =   2310
         Width           =   1515
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   4530
      Begin VB.CommandButton Command20 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3255
         TabIndex        =   5
         Top             =   285
         Width           =   175
      End
      Begin VB.Frame Frame11 
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
         TabIndex        =   10
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command26 
            Caption         =   "Inventarios"
            Height          =   435
            Left            =   210
            TabIndex        =   16
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Negativos"
            Height          =   435
            Left            =   2205
            TabIndex        =   15
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command24 
            Caption         =   "Consignación"
            Height          =   435
            Left            =   210
            TabIndex        =   14
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command23 
            Caption         =   "Críticos"
            Height          =   435
            Left            =   2205
            TabIndex        =   13
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Estadísticas"
            Height          =   435
            Left            =   210
            TabIndex        =   12
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Análisis A-B-C"
            Height          =   435
            Left            =   2205
            TabIndex        =   11
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Text            =   " "
         Top             =   285
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   705
         Width           =   3225
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3570
         Picture         =   "OPREACH04.frx":2680
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   285
         Width           =   750
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Opciones de Trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   6
         Top             =   1080
         Width           =   4110
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Vendedores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   2310
            TabIndex        =   2
            Top             =   315
            Width           =   1590
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Condiciones de Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   3
            Left            =   315
            TabIndex        =   3
            Top             =   630
            Width           =   1590
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   315
            TabIndex        =   1
            Top             =   315
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Materia Prima"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   4
            Left            =   2310
            TabIndex        =   4
            Top             =   630
            Width           =   1590
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla:"
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
         Left            =   -735
         TabIndex        =   17
         Top             =   345
         Width           =   2325
      End
   End
End
Attribute VB_Name = "OPREACH04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$
Private Sub Command16_Click()
'
    Command16.Enabled = False
    If TRIM$(MODUCONE$) <> "" Then
        If TRIM$(MODUCONE$) = "GRUCOCLI" Then
            LDT$ = LUDAT$
            Call CIERRARCH("GRUCOCLI")
            Call CIERRARCH("!GRUCOCLI")
            'Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
            Call TRALOCAL("GRUCOCLI", "")
            VTB% = VENTABU%("GRUCOCLI")
            If VTB% < 0 Then GoTo 99
            J& = 0
            For I& = 1 To ULTREG&("!GRUCOCLI")
                X$ = REGLEIDO$("!GRUCOCLI", I&)
                If CVI(Left$(X$, 2)) > 0 Then
                    J& = J& + 1
                    Call GRAREG("GRUCOCLI", X$, J&)
                End If
            Next I&
            Call SETULTREG("GRUCOCLI", J&)
            Call CIERRARCH("GRUCOCLI")
          ElseIf MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TATIMAT" Then
            Call CARTATIMA
          ElseIf MODUCONE = "CONPAGO" Then
            Call CARCONPAGOS
          ElseIf MODUCONE = "TAPROVIN" Then
            Call CARTAPROVIN
          ElseIf MODUCONE = "CUEINGAS" Then
            Call CARCUEINGAS
          Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
        End If
    End If
99  Command16.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Form_Load()
  '
  Screen.MousePointer = 11
  '
  Call CENTRAFORM(Me)
  '
  TABLA$(1) = "TAMONEDA"
  TABLA$(2) = "*VENDEDOR"
  TABLA$(3) = "CONPAGO"
  TABLA$(4) = "TAMAPRIM"
  '
  Screen.MousePointer = 1
  '
End Sub
Private Sub Option1_Click(Index As Integer)
    Option1(Index).Value = True
    Text2.TEXT = Option1(Index).Caption
    MODUCONE$ = TABLA$(Index)
    TETABLA$ = MODUCONE$
    If Left$(TETABLA$, 1) = "#" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(#)"
      ElseIf Left$(TETABLA$, 1) = "*" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(*)"
    End If
    Text1.TEXT = TETABLA$
    '
    For K% = 1 To 4
      If K% <> Index Then
          Option1(K%).Value = False
      End If
    Next K%
    'Command16.SetFocus
End Sub

Private Sub Option1_DblClick(Index As Integer)
    If TRIM$(MODUCONE$) <> "" Then
        If TRIM$(MODUCONE$) = "CONPAGO" Then
          Call CARCONPAGOS
          ElseIf MODUCONE = "TAMONEDA" Then
            'Call CARTAPROVIN
          ElseIf MODUCONE = "CUEINGAS" Then
            'Call CARCUEINGAS
          Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
        End If
    End If
End Sub
