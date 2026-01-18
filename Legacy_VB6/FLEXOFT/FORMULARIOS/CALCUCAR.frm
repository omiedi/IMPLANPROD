VERSION 5.00
Begin VB.Form CALCUCAR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-Cálculo de Carga de Máquinas"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10845
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00EFFFE0&
      Caption         =   "Avance de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   6000
      Begin VB.Frame Frame6 
         BackColor       =   &H00EFFFE0&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   1600
         Left            =   350
         TabIndex        =   40
         Top             =   2380
         Width           =   1065
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.O/F"
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
            Index           =   0
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   1000
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle"
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
            Left            =   0
            TabIndex        =   44
            Top             =   315
            Width           =   1000
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Operación"
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
            Index           =   3
            Left            =   0
            TabIndex        =   43
            Top             =   630
            Width           =   1000
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
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
            Index           =   4
            Left            =   0
            TabIndex        =   42
            Top             =   945
            Width           =   1000
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
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
            Index           =   5
            Left            =   0
            TabIndex        =   41
            Top             =   1240
            Width           =   1000
         End
      End
      Begin VB.PictureBox MARCOBARRA 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         Height          =   1530
         Left            =   5565
         ScaleHeight     =   1470
         ScaleWidth      =   60
         TabIndex        =   37
         Top             =   420
         Width           =   120
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   5000
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Calculando Fechas Proyectadas de Entrega"
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
         Left            =   525
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1745
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   4305
         TabIndex        =   17
         Text            =   " "
         Top             =   3570
         Width           =   645
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   1470
         TabIndex        =   16
         Text            =   " "
         Top             =   3570
         Width           =   2220
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   4305
         TabIndex        =   15
         Text            =   " "
         Top             =   3255
         Width           =   645
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   1470
         TabIndex        =   14
         Text            =   " "
         Top             =   3255
         Width           =   2220
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1995
         TabIndex        =   13
         Text            =   " "
         Top             =   2940
         Width           =   3690
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1470
         TabIndex        =   12
         Text            =   " "
         Top             =   2940
         Width           =   435
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   4515
         TabIndex        =   11
         Text            =   " "
         Top             =   2310
         Width           =   1170
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1470
         TabIndex        =   10
         Text            =   " "
         Top             =   2625
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1470
         TabIndex        =   9
         Text            =   " "
         Top             =   2310
         Width           =   1275
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Verificando Tiempos Standard"
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
         Left            =   525
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1215
         Width           =   5160
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Ordenamiento Ordenes de Fabricación por Prioridad"
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
         Left            =   525
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   950
         Width           =   5055
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Selección de Operaciones  Pendientes"
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
         Left            =   525
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   685
         Width           =   4110
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Selección de Ordenes de Fabricación Lanzadas"
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
         Left            =   525
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   420
         Width           =   4950
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Secuenciando Carga de Maquinas"
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
         Left            =   525
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1480
         Width           =   5055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "a las"
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
         Index           =   9
         Left            =   3465
         TabIndex        =   19
         Top             =   3630
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "a las"
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
         Index           =   8
         Left            =   3465
         TabIndex        =   18
         Top             =   3300
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "horas"
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
         Index           =   7
         Left            =   4725
         TabIndex        =   8
         Top             =   3630
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "horas"
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
         Index           =   6
         Left            =   4830
         TabIndex        =   7
         Top             =   3300
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Index           =   2
         Left            =   3020
         TabIndex        =   6
         Top             =   2385
         Width           =   1380
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EFFFE0&
      Caption         =   "Parámetros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   6300
      TabIndex        =   22
      Top             =   210
      Width           =   4425
      Begin VB.PictureBox Picture4 
         Height          =   520
         Left            =   2980
         ScaleHeight     =   465
         ScaleWidth      =   1185
         TabIndex        =   39
         Top             =   3380
         Width           =   1250
         Begin VB.Image Image2 
            Height          =   510
            Left            =   -380
            Picture         =   "CALCUCAR.frx":0000
            Top             =   0
            Width           =   2010
         End
      End
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
         Height          =   620
         Left            =   2980
         Picture         =   "CALCUCAR.frx":35EA
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Cancelar Proceso de Cálculo"
         Top             =   2730
         Width           =   600
      End
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
         Height          =   620
         Left            =   3640
         Picture         =   "CALCUCAR.frx":38F4
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Inicia Proceso de Cálculo"
         Top             =   2730
         Width           =   600
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EFFFE0&
         Caption         =   "O/F's Suspendidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   210
         TabIndex        =   33
         Top             =   2625
         Width           =   2640
         Begin VB.CheckBox Check7 
            BackColor       =   &H00EFFFE0&
            Caption         =   "Incluir O/F's Suspendidas"
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
            Left            =   525
            TabIndex        =   34
            Top             =   525
            Width           =   1485
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Ordenamiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   210
         TabIndex        =   29
         Top             =   1470
         Width           =   4025
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EFFFE0&
            Caption         =   "Por Prioridades"
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
            Left            =   525
            TabIndex        =   32
            Top             =   420
            Value           =   -1  'True
            Width           =   1905
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00EFFFE0&
            Caption         =   "Por Secuencia"
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
            Left            =   525
            TabIndex        =   31
            Top             =   630
            Width           =   1905
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
            Height          =   540
            Left            =   3255
            MaskColor       =   &H8000000F&
            Picture         =   "CALCUCAR.frx":41BE
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Re-Ordenar Secuenciaa"
            Top             =   315
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Inicio de Carga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         TabIndex        =   23
         Top             =   420
         Width           =   4025
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   9
            Left            =   1050
            TabIndex        =   26
            Top             =   420
            Width           =   1065
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   10
            Left            =   3045
            TabIndex        =   25
            Top             =   420
            Width           =   750
         End
         Begin VB.CommandButton Command3 
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
            Height          =   360
            Left            =   2100
            TabIndex        =   24
            Top             =   420
            Width           =   175
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   -1260
            TabIndex        =   28
            Top             =   525
            Width           =   2220
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hora:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   11
            Left            =   1365
            TabIndex        =   27
            Top             =   525
            Width           =   1590
         End
      End
      Begin VB.Line Line2 
         X1              =   4230
         X2              =   4230
         Y1              =   3360
         Y2              =   3900
      End
      Begin VB.Line Line1 
         X1              =   2980
         X2              =   4230
         Y1              =   3895
         Y2              =   3895
      End
   End
   Begin VB.ListBox LIORFALA 
      Height          =   450
      Left            =   5985
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   3675
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "CALCUCAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call CALCUMAQ
    Command2.Enabled = True
    Hide
End Sub

Private Sub Command3_Click()
    Call SELECHO("SELFECHA")
    FECHANUE$ = TRIM$(VALACT1("SELFECHA"))
    If FECHANUE$ <> "" Then
        Text1(9).TEXT = FECHANUE$
    End If
End Sub

Private Sub Command4_Click()
    Option2.Value = True
    Call SECUORF
End Sub

Private Sub Form_Load()
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    HOX% = HOY(HO$)
    Text1(9).TEXT = HO$
    Text1(10).TEXT = Left$(Time$, 5)
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub Text1_Change(Index As Integer)
    Static HOX%
    If HOX% = 0 Then
        HOX% = HOY(HO$)
    End If
    If Index = 9 Then
      If FECHANUM(Text1(9).TEXT) > HOX% Then
        Text1(10).TEXT = "0:00"
      End If
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 9 Then
            Text1(10).SetFocus
          ElseIf Index = 10 Then
            Command2.SetFocus
        End If
    End If
End Sub
