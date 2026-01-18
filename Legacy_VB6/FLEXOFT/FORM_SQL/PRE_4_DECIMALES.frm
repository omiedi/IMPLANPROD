VERSION 5.00
Begin VB.Form ANACOSI09 
   BackColor       =   &H00D0F0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costos Standard - Análisis y Presupuestacion de Productos"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11910
   Begin VB.CommandButton Command18 
      Caption         =   "C.REP"
      Height          =   255
      Left            =   11040
      TabIndex        =   104
      Top             =   7440
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   4440
      Top             =   3240
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00D0F0C0&
      Caption         =   "Última Modificación de Estructura"
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
      Left            =   210
      TabIndex        =   94
      Top             =   4240
      Width           =   4740
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   3570
         TabIndex        =   98
         Top             =   255
         Width           =   1005
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Index           =   12
         Left            =   2920
         TabIndex        =   97
         Top             =   250
         Width           =   750
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   840
         TabIndex        =   96
         Top             =   255
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
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
         Index           =   11
         Left            =   105
         TabIndex        =   95
         Top             =   250
         Width           =   750
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00D0F0C0&
      Caption         =   "Precios Calculados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   105
      TabIndex        =   73
      Top             =   6480
      Width           =   4005
      Begin VB.CommandButton Command10 
         Caption         =   "Ok"
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
         Left            =   3400
         TabIndex        =   83
         Top             =   300
         Width           =   480
      End
      Begin VB.CommandButton Command11 
         Caption         =   "X"
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
         Left            =   3400
         TabIndex        =   82
         Top             =   700
         Width           =   480
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2415
         TabIndex        =   81
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2415
         TabIndex        =   80
         Top             =   300
         Width           =   750
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dif:              %"
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
         Index           =   10
         Left            =   1840
         TabIndex        =   79
         Top             =   760
         Width           =   1485
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dif:              %"
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
         Left            =   1840
         TabIndex        =   78
         Top             =   350
         Width           =   1485
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   525
         TabIndex        =   77
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   525
         TabIndex        =   76
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " u$s:"
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
         Left            =   -300
         TabIndex        =   75
         Top             =   760
         Width           =   750
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " $: "
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
         Left            =   -350
         TabIndex        =   74
         Top             =   350
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   400
      Left            =   210
      ScaleHeight     =   345
      ScaleWidth      =   10545
      TabIndex        =   66
      ToolTipText     =   "Doble Clic Para Ver cotizacion de Dolar/M.Obra Real"
      Top             =   3840
      Width           =   10605
      Begin VB.CommandButton Command17 
         Caption         =   "P.O.E.'S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5660
         TabIndex        =   102
         ToolTipText     =   "Acceso Directo P.O.E.'S"
         Top             =   30
         Width           =   950
      End
      Begin VB.CommandButton Command16 
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
         Height          =   300
         Left            =   4650
         TabIndex        =   101
         ToolTipText     =   "Carga de Operaciones"
         Top             =   30
         Width           =   1000
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   7290
         TabIndex        =   100
         Top             =   30
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3700
         TabIndex        =   72
         Top             =   30
         Width           =   950
      End
      Begin VB.CommandButton Command14 
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
         Height          =   300
         Left            =   2760
         TabIndex        =   71
         Top             =   30
         Width           =   950
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1820
         TabIndex        =   70
         Top             =   30
         Width           =   950
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Nuevo Componente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   69
         Top             =   30
         Width           =   1820
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "M.O.R."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   6690
         TabIndex        =   99
         Top             =   75
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Cot.Dolar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   8280
         TabIndex        =   68
         Top             =   75
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   9150
         TabIndex        =   67
         Top             =   30
         Visible         =   0   'False
         Width           =   1380
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D0F0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   3405
      TabIndex        =   54
      Top             =   6510
      Width           =   7470
      Begin VB.CommandButton Command8 
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
         Left            =   2565
         TabIndex        =   61
         Top             =   700
         Width           =   175
      End
      Begin VB.TextBox Text7 
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
         Left            =   1635
         TabIndex        =   60
         Top             =   700
         Width           =   960
      End
      Begin VB.CommandButton Command5 
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
         Left            =   2565
         TabIndex        =   58
         Top             =   300
         Width           =   175
      End
      Begin VB.TextBox Text6 
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
         Left            =   1635
         TabIndex        =   57
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2835
         TabIndex        =   93
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2835
         TabIndex        =   92
         Top             =   705
         Width           =   1290
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   91
         ToolTipText     =   "Doble Click para Editar"
         Top             =   705
         Width           =   450
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   90
         ToolTipText     =   "Doble Click para Editar"
         Top             =   300
         Width           =   450
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6090
         TabIndex        =   65
         Top             =   705
         Width           =   1290
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6090
         TabIndex        =   64
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4725
         TabIndex        =   63
         ToolTipText     =   "Doble Click para Editar"
         Top             =   705
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4725
         TabIndex        =   62
         ToolTipText     =   "Doble Click para Editar"
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "P.Va (u$s) "
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
         Left            =   -210
         TabIndex        =   59
         Top             =   750
         Width           =   1905
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "   Ultima modif.  Usuario           L-Pr.  Neto                Final"
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
         Left            =   1500
         TabIndex        =   56
         Top             =   45
         Width           =   5895
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "P.Vta ($) "
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
         Left            =   -210
         TabIndex        =   55
         Top             =   315
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0F0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   -420
      TabIndex        =   44
      Top             =   4940
      Width           =   6210
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2040
         TabIndex        =   53
         Top             =   315
         Width           =   840
      End
      Begin VB.Label Ltext36 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   3645
         TabIndex        =   85
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Ltext5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   3540
         TabIndex        =   84
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   4725
         TabIndex        =   52
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "   Total Materia Prima: "
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
         Left            =   2805
         TabIndex        =   51
         Top             =   45
         Width           =   1905
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "   Mano de Obra:                min   x                  $:          "
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
         Left            =   525
         TabIndex        =   50
         Top             =   345
         Width           =   4320
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   4725
         TabIndex        =   49
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label68 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   4725
         TabIndex        =   48
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Gastos Generales:                %:"
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
         Index           =   2
         Left            =   1995
         TabIndex        =   47
         Top             =   630
         Width           =   3585
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   4725
         TabIndex        =   46
         Top             =   1050
         Width           =   1395
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Total:"
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
         Index           =   1
         Left            =   3800
         TabIndex        =   45
         Top             =   1060
         Width           =   4000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   6510
      TabIndex        =   32
      Top             =   4500
      Width           =   4320
      Begin VB.TextBox Text42 
         Alignment       =   1  'Right Justify
         Height          =   270
         Left            =   1760
         TabIndex        =   41
         Top             =   1470
         Width           =   750
      End
      Begin VB.Label Ltext43 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1680
         TabIndex        =   88
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Ltext39 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1680
         TabIndex        =   87
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Ltext41 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1680
         TabIndex        =   86
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Comision Vendedor:             %:"
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
         Index           =   31
         Left            =   0
         TabIndex        =   43
         Top             =   1550
         Width           =   2745
      End
      Begin VB.Label Label110 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   2730
         TabIndex        =   42
         Top             =   1500
         Width           =   1400
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Costo Calculado:"
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
         Index           =   28
         Left            =   -315
         TabIndex        =   40
         Top             =   1100
         Width           =   2955
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   2730
         TabIndex        =   39
         Top             =   1050
         Width           =   1400
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Impuestos:               %:"
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
         Index           =   30
         Left            =   630
         TabIndex        =   38
         Top             =   645
         Width           =   2115
      End
      Begin VB.Label Label104 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   2730
         TabIndex        =   37
         Top             =   600
         Width           =   1400
      End
      Begin VB.Label Label107 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   2730
         TabIndex        =   36
         Top             =   300
         Width           =   1400
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Utilidad:               %:"
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
         Index           =   29
         Left            =   840
         TabIndex        =   35
         Top             =   345
         Width           =   1905
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "            Scrap:               %:"
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
         Index           =   22
         Left            =   315
         TabIndex        =   34
         Top             =   45
         Width           =   2430
      End
      Begin VB.Label Label93 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   2730
         TabIndex        =   33
         Top             =   0
         Width           =   1400
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   210
      TabIndex        =   31
      ToolTipText     =   "Doble-Click para Agregar o Modificar"
      Top             =   1800
      Width           =   10620
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   10545
      TabIndex        =   29
      Top             =   1400
      Width           =   10605
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         Caption         =   $"ANACOSI09.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   30
         ToolTipText     =   "Doble-Click para Agregar o Modificar"
         Top             =   50
         Width           =   10515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0F0C0&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1120
      Index           =   0
      Left            =   9300
      TabIndex        =   20
      Top             =   105
      Width           =   1540
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   140
         Picture         =   "ANACOSI09.frx":0087
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   315
         Width           =   630
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Rest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   775
         Picture         =   "ANACOSI09.frx":0391
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3255
      Picture         =   "ANACOSI09.frx":069B
      TabIndex        =   19
      Top             =   210
      Width           =   175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1260
      MaxLength       =   15
      TabIndex        =   18
      Top             =   210
      Width           =   2040
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   675
      Width           =   4965
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1040
      Left            =   9030
      ScaleHeight     =   975
      ScaleWidth      =   60
      TabIndex        =   15
      Top             =   200
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5000
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8310
      Left            =   11025
      ScaleHeight     =   8250
      ScaleWidth      =   1005
      TabIndex        =   4
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "ANACOSI09.frx":09A5
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   4080
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
         Picture         =   "ANACOSI09.frx":0CAF
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ayuda en Línea"
         Top             =   745
         Width           =   650
      End
      Begin VB.CommandButton Command37 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "ANACOSI09.frx":0FB9
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   5010
         Width           =   650
      End
      Begin VB.CommandButton Command38 
         Caption         =   "List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "ANACOSI09.frx":13FB
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Complementos, Consultas, Listados"
         Top             =   5660
         Width           =   650
      End
      Begin VB.CommandButton Command39 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "ANACOSI09.frx":1705
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprime Ficha del Articulo"
         Top             =   6315
         Width           =   650
      End
      Begin VB.CommandButton Command25 
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
         Picture         =   "ANACOSI09.frx":1D6F
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8730
         Width           =   650
      End
      Begin VB.CommandButton Command24 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "ANACOSI09.frx":2079
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Baja Lógica Individual"
         Top             =   3375
         Width           =   650
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Subst"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "ANACOSI09.frx":2383
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Sustitución de Código de Artículo"
         Top             =   2735
         Width           =   650
      End
      Begin VB.CommandButton Command21 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "ANACOSI09.frx":24CD
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   1440
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "ANACOSI09.frx":2617
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Alta de Nuevo Item por Copia Completa de Datos de Artículo Existente"
         Top             =   2080
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "ANACOSI09.frx":2761
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   105
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "ANACOSI09.frx":28AB
         Top             =   7100
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0F0C0&
      Caption         =   "Acc.Directos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1120
      Index           =   1
      Left            =   7350
      TabIndex        =   1
      Top             =   105
      Width           =   1540
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
         Height          =   680
         Left            =   760
         Picture         =   "ANACOSI09.frx":47CD
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Otros Accesos Directos"
         Top             =   315
         Width           =   630
      End
      Begin VB.CommandButton Command7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   140
         Picture         =   "ANACOSI09.frx":4AD7
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Planera Electrónica"
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.TextBox Text45 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4215
      MaxLength       =   25
      TabIndex        =   0
      Top             =   210
      Width           =   2985
   End
   Begin VB.ListBox List2 
      Height          =   1185
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   89
      Top             =   1995
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11025
      Y1              =   6420
      Y2              =   6420
   End
   Begin VB.Line Line1 
      X1              =   10815
      X2              =   10815
      Y1              =   1365
      Y2              =   4200
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   11865
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   39
      Left            =   405
      TabIndex        =   28
      Top             =   285
      Width           =   795
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   38
      Left            =   -210
      TabIndex        =   27
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   315
      TabIndex        =   26
      Top             =   930
      Width           =   750
   End
   Begin VB.Label Label37 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6615
      TabIndex        =   25
      Top             =   930
      Width           =   585
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "U-M::"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   40
      Left            =   5985
      TabIndex        =   24
      Top             =   675
      Width           =   795
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Pl.Nro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   41
      Left            =   3360
      TabIndex        =   23
      Top             =   285
      Width           =   795
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu Abri 
         Caption         =   "Abrir"
      End
      Begin VB.Menu phl5 
         Caption         =   "-"
      End
      Begin VB.Menu Anuevo 
         Caption         =   "Nuevo Item de Stock"
      End
      Begin VB.Menu Acopia 
         Caption         =   "Alta por Copia"
      End
      Begin VB.Menu plh0 
         Caption         =   "-"
      End
      Begin VB.Menu Guar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu Imp 
         Caption         =   "Imprimir Ficha"
      End
      Begin VB.Menu plh2 
         Caption         =   "-"
      End
      Begin VB.Menu Sal 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Edi 
      Caption         =   "Edición"
      Begin VB.Menu DesCam 
         Caption         =   "Deshacer Cambios"
      End
      Begin VB.Menu phl4 
         Caption         =   "-"
      End
      Begin VB.Menu SustCod 
         Caption         =   "Sustitución de Código de Articulo"
      End
      Begin VB.Menu BaLog 
         Caption         =   "Baja Lógica Individual"
      End
      Begin VB.Menu mnu 
         Caption         =   "Procesos"
         Begin VB.Menu mnuasigfech 
            Caption         =   "Fecha de P.Vta en $ (Solo sin Fecha)"
         End
      End
   End
   Begin VB.Menu ListCon 
      Caption         =   "Listados y Consultas"
      Begin VB.Menu GralMa 
         Caption         =   "Generales del Maestro"
         Begin VB.Menu AtGral 
            Caption         =   "Atributos Generales"
         End
         Begin VB.Menu StoLot 
            Caption         =   "Stock Minino y Lote de Reposición"
         End
         Begin VB.Menu CapInv 
            Caption         =   "Captura para Inventarios"
         End
      End
      Begin VB.Menu CosRep 
         Caption         =   "Costos de Reposición"
         Begin VB.Menu NumPro 
            Caption         =   "Por Numero de Proveedor"
         End
         Begin VB.Menu CodArt 
            Caption         =   "Por Código de Articulo"
         End
      End
      Begin VB.Menu ItInac 
         Caption         =   "Items Inactivos"
         Begin VB.Menu ConRes 
            Caption         =   "Consulta y Restauración"
         End
         Begin VB.Menu Listad 
            Caption         =   "Listado"
         End
      End
      Begin VB.Menu phl6 
         Caption         =   "-"
      End
      Begin VB.Menu CosRan 
         Caption         =   "Costos Calculados por Rango de Código"
      End
   End
   Begin VB.Menu AccDire 
      Caption         =   "Accesos Directos"
      Begin VB.Menu MaePro 
         Caption         =   "Maestro de Proveedores"
      End
      Begin VB.Menu LisPre 
         Caption         =   "Lista de Precios de Venta"
      End
      Begin VB.Menu php6a 
         Caption         =   "-"
      End
      Begin VB.Menu Otro 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu Config 
      Caption         =   "Configuración"
   End
   Begin VB.Menu ayu 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "ANACOSI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim APREPES$, APREDOL$
Dim YACAR%(15), MODIFIC%(15)
Dim RESECO$(128), DESOPE$(128)
Dim PROCESOACTUCOSTO%




Sub ACTUALIZADATLISTA(NULIS%)
   'ACTUALIZO EL ARCHIVO ARCHIPRE
   '******************************
   On Error Resume Next
   ACLISPRE07.Visible = False
   ACLISPRE07.List1.ListIndex = NULIS% - 1
   ACLISPRE07.GENERAR_DAT (NULIS%)
   On Error GoTo 0

End Sub

Sub ACTUALIZALISTA(MON%, Optional OK%)
    '
    OK% = 0
    'ASIGNO VALORES SEGUN LA LISTA DE PRECIOS
    '****************************************
    If MON% = 1 Then            'MONEDA PESOS
      PREPES# = XVALO(Label1)
      FEXI% = FECHANUM(Text6)
    ElseIf MON% = 2 Then        'MONEDA DOLARES
      PREPES# = XVALO(Label6)
      FEXI% = FECHANUM(Text7)
    End If
    '
    'INICIALIZO BANDERA PARA MARCAR SI YA EL CODIGO ESTA EXISTENTE EN ALGUNA LISTA
    SINLIST% = 0
    '
    'ACTUALIZA LISTA EN SQL,RECORRE TODAS LAS LISTAS Y DONDE EXISTE
    'MISMO CODIGO EN LISTA ACTUALIZA EL PRECIO
    '*************************************************************
    SQLX$ = "SELECT * FROM LISTAPRE with(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE moneda = " & MON%
    SQLX$ = SQLX$ + " ORDER BY COD_LISTA"
    Dim LISTtmp As New ADODB.Recordset
    Set LISTtmp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With LISTtmp
      Do While Not .EOF
          If CI% > 0 Then
            NUMLIS% = XVALO(!Cod_Lista)
            '
            CONDI$ = "COD_PIEZA = " & CI%
            CONDI$ = CONDI$ + " AND COD_LISTA= " & NUMLIS%
            Call ACTUREGISTRO("PREVENTA", "Precio_Unitario", CONDI$, FORMATNUM$(PREPES#, "F16.5"), REG&, "NOMESS")
            Call ACTUREGISTRO("PREVENTA", "Cant_Limite", CONDI$, FEXI%, REG&, "NOMESS")
            '
            If REG& > 0 Then
              SINLIST% = 1 ' EXISTE EN LISTA
              Call ACTUALIZADATLISTA(NUMLIS%)
              OK% = 1
            End If
          End If
          .MoveNext
      Loop
    End With
    '
    'SI SINLIST% = 0 ES POR QUE NO EXISTE EL CODIGO EN NINGUNA LISTA DE PRECIOS
    'EN ESTE CASO PRESENTA SELECHO DE LISTAS PARA SELECCIONAR A CUAL SE VA A AGREGAR
    '*******************************************************************************
    If SINLIST% = 0 Then
         Call SELECHO("LISTAS")
         APREPES1$ = TRIM$(VALACT1$("LISTAS"))
         NUMLIS% = XVALO(VALACT1$("LISTAS"))        'NUMERO DE LISTA DE SELECCIONADA
         
         'SI NO SELECCIONA LISTA DE PRECIOS PRESENTA MENSAJE LIMPIA LOS LABEL Y SALE
         If NUMLIS% = 0 Then
           Call MENSERR(24, "Imposible Asignar Precio, Debe Seleccionar una Lista de Precios")
           If MON% = 1 Then            'MONEDA PESOS
              Label1 = ""
              Text6 = ""
           ElseIf MON% = 2 Then        'MONEDA DOLARES
              Label6 = ""
              Text7 = ""
           End If
           Exit Sub
         End If
         '
         If MON% = 1 Then                           'PESOS
           Label15 = APREPES1$                      'NUMERO DE LISTA
           Label15.ToolTipText = VALACT2$("LISTAS") 'DESCRIPCION DE LISTA
         ElseIf MON% = 2 Then                       'DOLAR
           Label16 = APREPES1$                      'NUMERO DE LISTA
           Label16.ToolTipText = VALACT2$("LISTAS") 'DESCRIPCION DE LISTA
         End If
         '
         'GENERA NUEVO REGISTRO
         If CantRegistros("PREVENTA", CONDI$) < 1 Then
100         ABUS$ = RANDSTRING$(-10)
            CONDI$ = "Num_Original = '" & ABUS$ & "'"
            If CantRegistros("PREVENTA", CONDI$) > 0 Then GoTo 100
            Call CreaRegistro("PREVENTA", "Num_Original", ABUS$, "NOMESS")
         End If
         '
         'GUARDA LOS DATOS DEL REGISTRO CREADO
         '************************************
         Call ACTUREGISTRO("PREVENTA", "COD_LISTA", CONDI$, APREPES1$, REG&, "NOMESS")
         '
         Numorden% = XVALO(VALOBADA("PREVENTA", "Max(Num_Orden)", "Cod_Lista =" & APREPES1, "NOMESS")) + 1
         Call ACTUREGISTRO("PREVENTA", "Precio_Unitario", CONDI$, FORMATNUM$(PREPES#, "F16.5"), REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Num_Orden", CONDI$, Numorden%, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Desc_Pieza", CONDI$, Text3, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Denom_Pieza", CONDI$, Text3, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Carac_Control", CONDI$, "", REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Cant_Limite", CONDI$, FEXI%, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Porc_Limite", CONDI$, 0, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Marborra", CONDI$, "", REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Cod_Pieza", CONDI$, CI%, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Num_Original", CONDI$, "", REG&, "NOMESS")
         'ACTUALIZA LISTA DE SELECCION
         Call ACTUALIZADATLISTA(NUMLIS%)
         OK% = 1
         '
    End If
    '\\\OT-08-0136-FL-FIN///


End Sub



Sub GRAUSUARIOYFECHA()

   'GRABO LA FECHA Y EL USUARIO QUE MODIFICO LA ESTRUCTURA
   HOII% = HOY(HOS$)
   Label19 = USERNAME$
   Label20 = FECHATEX$(HOII%)

   CONDI$ = "Codint=" & CODINT(Text1)
   Call ACTUREGISTRO("Master", "UsuApFor", CONDI$, Label19, REG&)
   Call ACTUREGISTRO("Master", "FechApFor", CONDI$, CVDAT(HOII%), REG&)
   '
End Sub

Sub MAILCOTI()
10  Call SELECHO("PROVED1")
    NCX1% = XVALO(VALACT1$("PROVED1"))
    If NCX1% > 0 Then
       '
       Call APERBASDAT("MASTER")
       '
       'Lectura desde SQL
       SQLX$ = "SELECT * FROM MASTER with(NOLOCK) "
       SQLX$ = SQLX$ + " WHERE PROVHABI = " & NCX1% & " AND Status > 0"
       SQLX$ = SQLX$ + " ORDER BY CODIGO ASC"
       '
       Dim MasterBaja As New ADODB.Recordset
       Set MasterBaja = FLEXCONN.Execute(FILTSQL$(SQLX$))
       '
       With MasterBaja
          If Err Then
            On Error GoTo 0
            Call MENSERR(24, "No hay Registros que Coincidan con el Proveedor Elegido .")
          End If
          '
          I& = 0
          Do While Not .EOF
             COD$ = !CODIGO
             MON% = !Monecos
             PODES1 = ROUND#(!PorDescCos, 1): If PODES1 > 99 Then PODES1 = 0
             COS1# = !COSUNI / (1 - PODES1 / 100)
             COS2# = !COSUNI
             FECH = CVINT(0)
             '
             Y$ = REGBLAN$("COTIPROV")
             CODI% = CODINT%(COD$)
             Call REPLA(Y$, MKI$(CODI%), 1, 2)
             Call REPLA(Y$, MKI$(MON%), 3, 2)
             Call REPLA(Y$, MKD$(COS1#), 5, 8)
             Call REPLA(Y$, MKS$(PODES1), 13, 4)
             Call REPLA(Y$, MKD$(COS2#), 17, 8)
             FCOS% = CVINT%(FECH)
             Call REPLA(Y$, MKI$(FCOS%), 25, 2)
             I& = I& + 1
             Call GRAREG("!COTIPROV", Y$, I&)
            .MoveNext
            '
          Loop
          Call SETULTREG("!COTIPROV", I&)
       End With
       '
       
40     JUN% = VENTABU%("COTIPROV/TITUPAN=Costo de Reposición - " + RASOCLI$((-1) * NCX1%))
       '
       If JUN% < 0 Then GoTo 10
       
       
       HOII% = HOY(HOS$)
       '
       ENCA$ = "Favor de Cotizar Los Siguientes Artículos" + vbCrLf
       ENCA$ = ENCA$ + "--------------------------------------------------------"
       TXCUERPO$ = ENCA$ & vbCrLf & vbCrLf
       
        For j& = 1 To ULTREG&("!COTIPROV")
            Z$ = REGLEIDO$("!COTIPROV", j&)
            Canti1 = CVD(Mid$(Z$, 27, 8))
            If Canti1 > 0.05 Then
                COD3% = CVI(Mid$(Z$, 1, 2))
                PODES1 = CVS(Mid$(Z$, 13, 4)): If PODES1 < 0 Or PODES1 > 99 Then PODES1 = 0
                COU = (CVD(Mid$(Z$, 5, 8))) * (1 - PODES1 / 100)
                MONEI% = CVI(Mid$(Z$, 3, 2))
                FCOS% = CVI(Mid$(Z$, 25, 2))
                CUERPO$ = ""
                Call REPLA(CUERPO$, "Código: ", 1, 10)
                Call REPLA(CUERPO$, CODEXT$(COD3%), 12, 16)
                Call REPLA(CUERPO$, DESCRIT0$(COD3%), 30, 32)
                TXCUERPO$ = TXCUERPO$ + CUERPO$ + vbCrLf
                  CUERPO$ = "Cantidad: " & FORMATNUM$(Canti1, "F12.1")
                TXCUERPO$ = TXCUERPO$ + CUERPO$ + vbCrLf
                  CUERPO$ = "Costo: " & SIMBOLPES$(MONEI%) & " " & FORMATNUM$(COU, "F12.2")
                TXCUERPO$ = TXCUERPO$ + CUERPO$ + vbCrLf
                  CUERPO$ = "Fecha Entrega: " & FECHATEX$(FCOS%)
                TXCUERPO$ = TXCUERPO$ + CUERPO$ + vbCrLf
                  CUERPO$ = "---------------------------------------------"
                TXCUERPO$ = TXCUERPO$ + CUERPO$ + vbCrLf + vbCrLf
            End If
            '
        Next j&
        '
        Debug.Print TXCUERPO$
        STDMAIL.Text1 = RASOCLI$(-1 * NCX1%)
        STDMAIL.Text6 = "Solicitud de Cotización"
        STDMAIL.Text2 = EMAILCLI$(-1 * NCX1%)
        STDMAIL.Text5 = TXCUERPO$
        STDMAIL.Show 1
        GoTo 40
        '
       End If
    

End Sub

Sub PRECIOENLISTA(CI%)

'    For U& = 1 To ULTREG&("LISTAS")
'       XX$ = REGLEIDO$("LISTAS", U&)
'       If Asc(Mid$(XX$, 50, 1)) = 1 Then           'VALIDA TIPO DE MONEDA = $
'          APREPES1$ = TRIM$(Str$(U&))
'          While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
'          APREPES1$ = "LIPRE" + APREPES1$
''          PREPES# = PRELISTA#(APREPES1$, CI%, FEX1%)
'          PREPES# = PRELISTA#(APREPES1$, CI%)
'          If PREPES# >= 0.001 Then
'             Label1 = FORMATNUM$(PREPES#, "F10.4")
'             Label7 = FORMATNUM$(PREPES# * (1 + ALIVA / 100), "F10.4")
'             Text6 = FECHATEX$(FEX1)
'             Label15 = U&
'             UI% = U&
'             Label15.ToolTipText = DENOLISTA$(UI%)
'             GoTo 20
'          End If
'          Call CIERRARCH(APREPES1$)
'       End If
'    Next U&
'    '
'20  For U& = 1 To ULTREG&("LISTAS")
'        XX$ = REGLEIDO$("LISTAS", U&)
'        If Asc(Mid$(XX$, 50, 1)) = 2 Then           'VALIDA TIPO DE MONEDA = U$S
'            APREPES1$ = TRIM$(Str$(U&))
'            While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
'            APREPES1$ = "LIPRE" + APREPES1$
''            PREDOL# = PRELISTA#(APREPES1$, CI%, FEX2%)
'            PREDOL# = PRELISTA#(APREPES1$, CI%)
'            If PREDOL# >= 0.005 Then
'                Label6 = FORMATNUM$(PREDOL#, "F10.4")
'                Label8 = FORMATNUM$(PREDOL# * (1 + ALIVA / 100), "F10.4")
'                Text7 = FECHATEX$(FEX2)
'                Label16 = U&
'                UI% = U&
'                Label16.ToolTipText = DENOLISTA$(UI%)
'                Exit For
'            End If
'        End If
'    Next U&
   '
   SQLX$ = "SELECT * FROM LISTAPRE with(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE moneda = 1 "

   SQLX$ = SQLX$ + " ORDER BY COD_LISTA"
   Dim LISTtmp As New ADODB.Recordset
   Set LISTtmp = FLEXCONN.Execute(FILTSQL$(SQLX$))
  
   With LISTtmp

     Do While Not .EOF
          NUM% = XVALO(!Cod_Lista)
          APREPES1$ = TRIM$(Str$(NUM%))
          While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
          APREPES1$ = "LIPRE" + APREPES1$
          PREPES# = PRELISTA#(APREPES1$, CI%)
          If PREPES# >= 0.001 Then
             Label1 = FORMATNUM$(PREPES#, "F10.2")
             
             Label7 = FORMATNUM$(PREPES# * (1 + ALIVA / 100), "F10.2")
             CONDI$ = "COD_LISTA = " & NUM% & " And COD_PIEZA = " & CI%
             FEX1 = XVALO(VALOBADA("PREVENTA", "CANT_LIMITE", CONDI$))
             Text6 = FECHATEX$(FEX1)
             Label15 = NUM%
             Label15.ToolTipText = DENOLISTA$(NUM%)
             GoTo 20
          End If
         .MoveNext
     Loop
   End With
   '
20 SQLX$ = "SELECT * FROM LISTAPRE with(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE moneda = 2 "
   SQLX$ = SQLX$ + " ORDER BY COD_LISTA"
   Dim LISTtmp1 As New ADODB.Recordset
   Set LISTtmp1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
  
   With LISTtmp1
  
     Do While Not .EOF
         NUM% = XVALO(!Cod_Lista)
         APREPES1$ = TRIM$(Str$(NUM%))
         While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
         APREPES1$ = "LIPRE" + APREPES1$
         PREDOL# = PRELISTA#(APREPES1$, CI%)
         If PREDOL# >= 0.005 Then
             Label6 = FORMATNUM$(PREDOL#, "F10.2")
             Label8 = FORMATNUM$(PREDOL# * (1 + ALIVA / 100), "F10.2")
             CONDI$ = "COD_LISTA = " & NUM% & " And COD_PIEZA = " & CI%
             FEX2 = XVALO(VALOBADA("PREVENTA", "CANT_LIMITE", CONDI$))
             Text7 = FECHATEX$(FEX2)
             Label16 = NUM%
             Label16.ToolTipText = DENOLISTA$(NUM%)
         End If
         .MoveNext
     Loop
   End With
End Sub
Sub PRINTINFOGRILLAJM(NPROV%, DES%, HAS%)
  
    If INFO_GRILLAJM.MSF.Rows < 1 Then
        Call COMUNI("Imposible Imprimir  - Sin Información")
        Exit Sub
    End If
    '
    Call BORRAREGISTROS("PRINTNOVEDADES", "CODINT = '' OR CODINT IS NULL OR CODINT > 0 ", REGAF&, "NOMESS")
    
    SQLX$ = "SELECT * FROM PRINTNOVEDADES "
    SQLX$ = SQLX$ + " WHERE CODINT = -1"
    '
    Set RST = New ADODB.Recordset
25  On Error Resume Next
    RST.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0

  With INFO_GRILLAJM
          For I& = 1 To .MSF.Rows - 1
             CODI$ = TRIM$(.MSF.TextMatrix(I&, 1))
             FE1% = FECHANUM(.MSF.TextMatrix(I&, 3))
             P_ANT = XVALO(.MSF.TextMatrix(I&, 4))
             FE2% = FECHANUM(.MSF.TextMatrix(I&, 5))
             P_NEW = XVALO(.MSF.TextMatrix(I&, 6))
             PJE = XVALO(.MSF.TextMatrix(I&, 7))
             MONEX% = TRIM$(.MSF.TextMatrix(I&, 8))
             '
             With RST
               .AddNew
               !CODINT = CODINT%(CODI$)
               !PORCENTAJE = PJE
                !Fech_Mod1 = CVDAT(FE1%)
                !Fech_Mod2 = CVDAT(FE2%)
                !Prec_Ant = P_ANT
                !Prec_New = P_NEW
                !MONEDA = MONEX%
               .Update
             End With
   
          Next I&
  End With
  '
  
  FILTX$ = TRIM$(NPROV%) & ";" & TRIM$(Str$(DES%) & ";" & TRIM$(Str$(HAS%)))
    
  Call STDFORM("NOV-JYM", FILTX$)

  
        
End Sub

Sub PRINTINFOGRILLA07(CI1%, DES%, HAS%)
  
    If INFO_GRILLA07.MSF.Rows < 1 Then
        Call COMUNI("Imposible Imprimir  - Sin Información")
        Exit Sub
    End If
    '
    Call BORRAREGISTROS("PRINTNOVEDADES", "CODINT = '' OR CODINT IS NULL OR CODINT > 0 OR CODINT <= 0", REGAF&, "NOMESS")
    '
    SQLX$ = "SELECT * FROM PRINTNOVEDADES "
    SQLX$ = SQLX$ + " WHERE CODINT = -1"
    '
    Set RST = New ADODB.Recordset
25  On Error Resume Next
    RST.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    With INFO_GRILLA07
          For I& = 1 To .MSF.Rows - 1
             FE1% = FECHANUM(.MSF.TextMatrix(I&, 1))
             P_ANT = XVALO(.MSF.TextMatrix(I&, 2))
             P_NEW = XVALO(.MSF.TextMatrix(I&, 3))
             PJE = XVALO(.MSF.TextMatrix(I&, 4))
             MONEX% = TRIM$(.MSF.TextMatrix(I&, 5))
             NOMBUSER$ = TRIM$(Left$(.MSF.TextMatrix(I&, 6), 24))
             OBSERV$ = Left$(.MSF.TextMatrix(I&, 7), 128)
             '
             With RST
                .AddNew
                !PORCENTAJE = PJE
                !Fech_Mod1 = CVDAT(FE1%)
                !Prec_Ant = P_ANT
                !Prec_New = P_NEW
                !Observacion = OBSERV$
                !MONEDA = MONEX%
                !NOM_USER = NOMBUSER$
               .Update
             End With
   
          Next I&
    End With
    '
    FILTX$ = TRIM$(CI1%) & ";" & TRIM$(Str$(DES%) & ";" & TRIM$(Str$(HAS%)))
    '
    Call STDFORM("NOV2-JYM", FILTX$)

  
        
End Sub


Private Sub Abri_Click()
   Call Command9_Click
End Sub

Private Sub Acopia_Click()
   Call Command22_Click
End Sub

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Anuevo_Click()
   Call Command21_Click
End Sub

Private Sub AtGral_Click()
   Call OPLIMAS04.Command9_Click
End Sub
Private Sub BaLog_Click()
   Call Command24_Click
End Sub

Private Sub CapInv_Click()
   Call OPLIMAS04.Command15_Click
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   Command9.Enabled = False
   Text1.Enabled = False
   Call GRADATITEM
   Call GRAFORMULA
   '
   Call GRATIEMOD
   '
   MODIFIC%(0) = 0
   Text1.Enabled = True
   Command9.Enabled = True
   Command1.Enabled = True
End Sub

Private Sub Command10_Click()
   '
   ' aqui es para mandar el precio a las listas de precios
   '
    CI% = CODINT%(Text1.TEXT)
    If CI% < 1 Then Exit Sub
    '
    If DERACCE%("GRAPREVE", "Grabar Precios de Venta s/Costos") < 2 Then
       Exit Sub
    End If
    '
    HOII% = HOY(HOS$)
    If COMALTER%("Desea Registrar los Nuevos Precios de Venta\con Fecha '" + HOS$ + "' ?\\Si, Registrarlos\No, Conservar Anteriores") <> 1 Then
       Exit Sub
    End If
    '
    Label1.Caption = FORMATNUM$(XVALO(Label103.Caption) + XVALO(Label110.Caption), "F10.2")
    If XVALO(Label11) > 0 Then
       Label6.Caption = FORMATNUM$((XVALO(Label103.Caption) + XVALO(Label110.Caption)) / XVALO(Label11.Caption), "F10.2")
    End If
    '
    Label7.Caption = FORMATNUM$(((XVALO(Label1.Caption) * ALIVA) / 100) + XVALO(Label1.Caption), "F10.2")
    Label8.Caption = FORMATNUM$(((XVALO(Label6.Caption) * ALIVA) / 100) + XVALO(Label6.Caption), "F10.2")
    Text6 = HOS$: Text7 = HOS$
    '
    '\\\OT-08-0095-FL-14/03/08///
    'ACTUALIZA LISTAS EN PESOS
'    For U& = 1 To ULTREG&("LISTAS")
'        XX$ = REGLEIDO$("LISTAS", U&)
'        If Asc(Mid$(XX$, 50, 1)) = 1 Then       'VALIDA TIPO DE MONEDA = $
'            APREPES1$ = TRIM$(Str$(U&))
'            While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
'            APREPES1$ = "LIPRE" + APREPES1$
'            PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
'            CODIN% = CVI(Mid$(PRE1$, 1, 2))        'VALIDA QUE EL CODIGO PERTENEZCA
'            If CODIN% > 0 Then                     'A LA LISTA DE PRECIOS
'                PREPES# = XVALO(Label1)
'                FEXI% = FECHANUM(Text6)
'                Call REPLA(PRE1$, MKI$(CI%), 1, 2)
'                Call REPLA(PRE1$, MKD$(PREPES#), 62, 8)
'                Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
'                Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
'                SINLIST% = 1
'            End If
'         End If
'    Next U&
    'ACTUALIZA LISTAS EN DOLARES
'    For U& = 1 To ULTREG&("LISTAS")
'        XX$ = REGLEIDO$("LISTAS", U&)
'        If Asc(Mid$(XX$, 50, 1)) = 2 Then       'VALIDA TIPO DE MONEDA = U$S
'            APREPES1$ = TRIM$(Str$(U&))
'            While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
'            APREPES1$ = "LIPRE" + APREPES1$
'            PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
'            CODIN% = CVI(Mid$(PRE1$, 1, 2))        'VALIDA QUE EL CODIGO PERTENEZCA
'            If CODIN% > 0 Then                     'A LA LISTA DE PRECIOS
'                PREDOL# = XVALO(Label6)
'                FEXI% = FECHANUM(Text6)
'                Call REPLA(PRE1$, MKI$(CI%), 1, 2)
'                Call REPLA(PRE1$, MKD$(PREDOL#), 62, 8)
'                Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
'                Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
'                SINLIST% = 1
'            End If
'         End If
'    Next U&
    '\\\OT-08-0095-FL-FIN///
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ACTUALIZA LISTAS EN PESOS SQL
    SQLX$ = "SELECT * FROM LISTAPRE with(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE moneda = 1 "
    SQLX$ = SQLX$ + " ORDER BY COD_LISTA"
    '
    Dim LISTtmp As New ADODB.Recordset
    Set LISTtmp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With LISTtmp
      
      Do While Not .EOF
        If CI% > 0 Then
            NUM% = XVALO(!Cod_Lista)
            PREPES# = XVALO(Label1)
            FEXI% = FECHANUM(Text6)
            '
            CONDI$ = "COD_PIEZA = " & CI% & " AND Cod_Lista=" & NUM%
            Call ACTUREGISTRO("PREVENTA", "Precio_Unitario", CONDI$, FORMATNUM$(PREPES#, "F16.5"), REG&, "NOMESS")
            If REG& > 0 Then SINLIST% = 1
        End If
       .MoveNext
      Loop
    End With
      
    'ACTUALIZA LISTAS EN DOLARES SQL
    SQLX$ = "SELECT * FROM LISTAPRE with(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE moneda = 2 "
    SQLX$ = SQLX$ + " ORDER BY COD_LISTA"
    '
    Dim LISTtmp1 As New ADODB.Recordset
    Set LISTtmp1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With LISTtmp1
      Do While Not .EOF
        If CI% > 0 Then
           NUM% = XVALO(!Cod_Lista)
           PREDOL# = XVALO(Label6)
           FEXI% = FECHANUM(Text6)
           '
           CONDI$ = "COD_PIEZA = " & CI% & " AND Cod_Lista=" & NUM%
           Call ACTUREGISTRO("PREVENTA", "Precio_Unitario", CONDI$, FORMATNUM$(PREDOL#, "F16.5"), REG&, "NOMESS")
           If REG& > 0 Then SINLIST% = 1
        End If
       .MoveNext
      Loop
    End With
    '
        '\\\OT-08-0136-FL-04/04/08///
    If SINLIST% = 0 Then
         Call SELECHO("LISTAS")
         APREPES1$ = TRIM$(VALACT1$("LISTAS"))
'        While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
'        APREPES1$ = "LIPRE" + APREPES1$
'        PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
'        If Asc(Mid$(XX$, 50, 1)) = 2 Then
'            PREPES# = XVALO(Label6)
'        Else
'            PREPES# = XVALO(Label1)
'        End If
         '
         If APREPES1$ = "LIPRE002" Then
            PREPES# = XVALO(Label6)
         Else
            PREPES# = XVALO(Label1)
         End If
         '
         FEXI% = FECHANUM(Text6)
'        Call REPLA(PRE1$, MKI$(CI%), 1, 2)
'        Call REPLA(PRE1$, MKD$(PREPES#), 62, 8)
'        Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
'        Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
         '
         If CantRegistros("PREVENTA", CONDI$) < 1 Then
100         ABUS$ = RANDSTRING$(-10)
            CONDI$ = "Num_Original = '" & ABUS$ & "'"
            If CantRegistros("PREVENTA", CONDI$) > 0 Then GoTo 100
            Call CreaRegistro("PREVENTA", "Num_Original", ABUS$, "NOMESS")
         End If
         '
         Call ACTUREGISTRO("PREVENTA", "COD_LISTA", CONDI$, APREPES1, REG&, "NOMESS")
         '
         Numorden% = XVALO(VALOBADA("PREVENTA", "Max(Num_Orden)", "Cod_Lista =" & APREPES1, "NOMESS")) + 1
         '
         Call ACTUREGISTRO("PREVENTA", "Precio_Unitario", CONDI$, FORMATNUM$(PREPES#, "F16.5"), REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Num_Orden", CONDI$, Numorden%, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Desc_Pieza", CONDI$, Text3, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Denom_Pieza", CONDI$, Text3, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Carac_Control", CONDI$, "", REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Cant_Limite", CONDI$, 0, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Porc_Limite", CONDI$, 0, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Marborra", CONDI$, "", REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Cod_Pieza", CONDI$, CI%, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Num_Original", CONDI$, "", REG&, "NOMESS")
         '
    End If
    '\\\OT-08-0136-FL-FIN///
    '
    Call PRECIOENLISTA(CI%)
    '
    Call CALCUCOS

End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call Command21_Click
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   '
   If List1.ListCount > 0 Then
     If List1.ListIndex >= 0 Then
       If List1.ListIndex < List1.ListCount Then
         Call GRAFORMULA
         CICOMPO% = CODINT%(Left$(List1.TEXT, 15))
         Call APERBASDAT("MASTER")
'         Master.FindFirst ("CODINT = " & CICOMPO%)
         CONDI$ = "CODINT = " & CICOMPO%
'             If Master.NoMatch = False Then
'                PORDE1 = ROUND#(Master!PorDescCos, 1)
                PORDE1 = VALOBADA("MASTER", "PorDescCos", CONDI$, "NOMESS")
                If PORDE1 > 99 Then PORDE1 = 0
'             End If
             COSTO1 = COSUNI(CICOMPO%)
             PRECIO1 = (COSTO1 / (1 - PORDE1 / 100))
             VDEF$ = MKI$(CICOMPO%) + MKS$(COSTO1) + MKI$(MONECOSTO%(CICOMPO%)) + MKI$(PROVHABI%(CICOMPO%)) + MKI$(FECOSTO%(CICOMPO%)) + MKS$(PRECIO1) + MKS$(PORDE1)
             OBX$ = OBJPLA$("AJUCOSTO", VDEF$)
             If OBX$ <> "" Then
               PRECIO1 = CVS(Mid$(OBX$, 13, 4))
               MONEI% = CVI(Mid$(OBX$, 7, 2))
               PROV1% = CVI(Mid$(OBX$, 9, 2))
               FECH1% = CVI(Mid$(OBX$, 11, 2))
               PORDE1 = CVS(Mid$(OBX$, 17, 4))
                 If PORDE1 < 0 Or PORDE1 > 99 Then PORDE1 = 0
               COSTO1 = PRECIO1 * (1 - PORDE1 / 100)
               Call ABREMASTER
               If CICOMPO% > 0 Then
'                 Master.FindFirst ("CODINT = " & CICOMPO%)
'                 If Master.NoMatch = False Then
'                   Master.Edit
'                   Master!COSUNI = COSTO1
'                   Master!Monecos = MONEI%
'                   Master!FECOSTO = CVDAT(FECH1%)
'                   Master!PorDescCos = PORDE1
'                   Master!PROVHABI = PROV1%
'                   Master.Update
                   '
                   'SQL
                   SQLX$ = " SELECT * FROM Master"
                   SQLX$ = SQLX$ + " WHERE CODINT = " & CICOMPO% & ""
                   Dim MODMASTER As ADODB.Recordset
                   Set MODMASTER = New ADODB.Recordset
                   MODMASTER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
                   On Error Resume Next
                   With MODMASTER
                     .MoveFirst
                     If Err < 1 Then
                        On Error GoTo 0
                        !COSUNI = COSTO1
                        !Monecos = MONEI%
                        !FECOSTO = CVDAT(FECH1%)
                        !PorDescCos = PORDE1
                        !PROVHABI = PROV1%
                        .Update
                     End If
                   End With
                   Set MODMASTER = Nothing
                   Call COMUNI("Costo Actualizado")
                   Call CARFORMULA
                   '
                   Call CALCUCOS
                   '
'                 End If
               End If
             End If
           End If
         End If
       End If
       '
    End Sub
    
Private Sub Command14_Click()
   Command14.Enabled = False
   Call List1_DblClick
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Call Command39_Click
End Sub

Private Sub Command16_Click()
'    Form1.Show
'    Exit Sub
    Command16.Enabled = False
    CODI$ = TRIM$(Text1)
    CIXI% = CODINT%(CODI$)
    If CIXI% < 1 Then
       Call COMUNI("Código Inexistente o No Válido")
       GoTo 99
    End If
    '
    TIEMMOB09.Text15.TEXT = CODI$
    TIEMMOB09.Show 1
    GoTo 99
    
    Call SETULTREG("!SECOPEJM", 0)
    
'    CONDI$ = " CODICONJ = " & CIXI%
'    ACUMUTOT = 0
'    SQLX$ = "SELECT * FROM  SECOPER WITH(NOLOCK) "
'    SQLX$ = SQLX$ + " WHERE CODICONJ = " & CIXI%
'    SQLX$ = SQLX$ + " AND NUORD > 0"
'    SQLX$ = SQLX$ + " ORDER BY NUORD"
'    Set RST = READSET(SQLX$)
'    With RST
'      Do While Not .EOF
'         NUOR% = XVALO(!NUORD)
'         DESCRI$ = SAFETEXT$(!OBSEOPER)
'         TIEMP% = XVALO(!MinuStan) / 60
'         SECCI = XVALO(!SecciPro)
'         CANTI = XVALO(!CantiOps)
'         ACUMUTOT = ACUMUTOT + CANTI
'         FECHREV = CVDAT(CVINT(!FECH_REVIS))
'         '
'         Call REPLA(TX$, MKI$(NUOR%), 1, 2)
'         Call REPLA(TX$, MKS$(TIEMP%), 3, 4)
'         Call REPLA(TX$, MKS$(SECCI), 7, 4)
'         Call REPLA(TX$, MKS(CANTI), 11, 2)
'         Call REPLA(TX$, MKI$(FECHREV), 13, 2)
'         U& = U& + 1
'         Call GRAREG("!SECOPEJM", TX, U&)
'         .MoveNext
'      Loop
'    End With
'    RST.Close
'    Set RST = Nothing
'
'    VTB% = VENTABU%("SECOPEJM")
'    If VTB% < 0 Then GoTo 99
'    '
'    'AQUI PRESENTAR MENSAJE DE CONFIRMACION
'
'     Call ACTUREGISTRO("SECOPER", "MARBORRA", CONDI$, 1, REG&, "NOMESS")
     
     SQLX$ = "SELECT * FROM SECOPER "
     SQLX$ = SQLX$ + " WHERE " & CONDI$
     Set RST1 = New ADODB.Recordset
     RST1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       With RST1
         For I& = 1 To ULTREG&("!SECOPEJM")
           XX$ = REGLEIDO$("!SECOPEJM", I&)
           NUOR% = CVI(Mid$(XX$, 1, 2))
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
           DESCRI$ = 1 'VER
           TIEM = CVS(Mid$(XX$, 3, 4))
           SECCI = CVS(Mid$(XX$, 7, 4))
           CANTI = CVS(Mid$(XX$, 13, 2))
           ACUMUTOT = ACUMUTOT + CANTI
           FECHREV = CVI(Mid$(XX$, 13, 2))
           '
           !NUORD = NUOR%
           !OBSEOPER = DESCRI$ 'GUARDAR LA DESCRIPCION CORTA POR DEFECTO
           !MinuStan = TIEMP%
           !SecciPro = SECCI
           !CantiOps = CANTI
            'VALIDA SI SE PRODUJO ALGUNA MODIFICACION EN EL ITEM Y EN ESTE CASO SE LE CAMBIA LA FECHA DE ULTIMA ACTUALIZACION
             CONDI1$ = " CODICONJ = " & CIXI% & " AND NUORD = " & NUOR% & " AND MARBORRA = 1"
             CANTI_ANT = XVALO(VALOBADA("SECOPER", "CANTIOPS", CONDI$, "NOMESS"))
             TIEM_ANT = XVALO(VALOBADA("SECOPER", "MINUSTAN", CONDI$, "NOMESS"))
             SECC_ANT = XVALO(VALOBADA("SECOPER", "SECCI", CONDI$, "NOMESS"))
             If CANTI_ENT <> XVALO(CANTI) Or TIEM_ANT <> XVALO(TIEMP%) Or SECC_ANT <> XVALO(SECCI) Then
                FECHREV = Date
             End If
           !FECH_REVIS = FECHREV
            
           !MARBORRA = 0
           .Update
         Next I&
       End With
       RST1.Close
       Set RST1 = Nothing
       
       CONDI$ = CONDI$ & " AND MARBORRA = 1"
       Call BORRAREGISTROS("PROVECOD", CONDI$, REG&)
      


   
   
   
99 Command16.Enabled = True

End Sub

Private Sub Command17_Click()
    CODX$ = REPLACAR(TRIM$(Text1), "/", "|")
    Call CONECRUN("AMASTO07/" + "P.O.E.'S" + CODX$, "Archivos Maestros")
End Sub




Private Sub Command18_Click()
   Command18.Enabled = False
   
   
   If CLACONTRA% > 0 Then
     
'         Call COMUNI("Comienzo de Proceso de Eliminar Items Repetidos en Formulas")
'
'         Call ENUMERA("FORMULAS", "ORDENNUMER", "CODICONJ > 0 ")
'        '
'         j& = 0
'         FIN& = CantRegistros&("FORMULAS", "CODICONJ > 0", "NOMESS")
'         SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK)"
'         SQLX$ = SQLX$ + " ORDER BY CodIConj, CODIELEM"
'         Set RTABLA = READSET(SQLX$)
'         With RTABLA
'           Do While Not .EOF
'              j& = j& + 1
'              Call TRACE(24, j&, FIN&)
'              CODICONJUN1% = XVALO(!CODICONJ)
'              If CODICONJUN1% = CODCONJ_ANT% Then
'                 CODIELEM1% = XVALO(!CODIELEM)
'                 If CODIELEM1% = CODIELEM1_ANT% Then
'                    NUORD& = XVALO(!ORDENNUMER)
'                    Call BORRAREGISTROS("FORMULAS", "ORDENNUMER = " & NUORD&, REGAF&, "NOMESS")
'                 End If
'                 CODIELEM1_ANT% = CODIELEM1%
'              Else
'                 CODIELEM1_ANT% = 0
'              End If
'              CODCONJ_ANT% = CODICONJUN1%
'           .MoveNext
'           Loop
'         End With
'         '
'         RTABLA.Close
'         Set RTABLA = Nothing
'         '
'        Call COMUNI("Proceso 1 Finalizado")

        Call COMUNI("Inicio de Proceso Actualización de Costos")
    
        PROCESOACTUCOSTO% = 1
        Command18.Enabled = False
        Call TRALOCAL("GRUCOVEN", "")
        Screen.MousePointer = 11
        For I& = 1 To ULTREG&("!GRUCOVEN")
            X$ = REGLEIDO$("!GRUCOVEN", I&)
            GRUCONTA% = CVI((Left$(X$, 2)))
            If GRUCONTA% > 0 Then
                CONDI$ = "GRUCOVEN = " & GRUCONTA%
                CONDI$ = CONDI$ + " AND STATUS >=0"
                FIN& = CantRegistros&("MASTER", CONDI$, "NOMESS")
                j& = j& + 1
                Call TRACE(24, j&, FIN&)
                SQLX$ = "SELECT CODIGO FROM MASTER WITH(NOLOCK)"
                SQLX$ = SQLX$ + " WHERE " & CONDI$
                Set RS = READSET(SQLX$)
                With RS
                Do While Not .EOF
                  Text1 = SAFETEXT$(!CODIGO)
                  .MoveNext
                Loop
                End With
            End If
         Next I&
         '
         Call CIERRARCH("GRUCOVEN")
         On Error Resume Next
         RS.Close
         Set RS = Nothing
         On Error GoTo 0
         PROCESOACTUCOSTO% = 0
         Screen.MousePointer = 1
         Call COMUNI("Proceso Actualización de Costos Finalizado")

    End If
    '
    Command18.Enabled = True
    '
End Sub


Private Sub Command21_Click()
   '
   Command21.Enabled = False
   Call GRADATITEM
   '
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
   Call REPLA(VDEF$, "U.", 109, 2)
   Call REPLA(VDEF$, Chr$(1), 93, 1)

   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   If OKX% = 1 Then
     Call GENECOMAS
   End If
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command22_Click()
   '
   '10/04/07 (OT-07-0137)
   Dim FIELDVAL(256)
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command22.Enabled = False
   Call GRADATITEM
   If VALCODART% < 1 Then GoTo 99
   '
   CODIVIE$ = AJUSTI$(Text1.TEXT, 16)
   CII% = CODINT%(Text1.TEXT)
   If CII% < 1 Or CII% > NUMAS% Then
      Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      Command25.SetFocus
      GoTo 99
   End If
   '
   'If REMAS& < 2 Or REMAS& > ULTREG&("MASTER") Then
   '   Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
   '   Command25.SetFocus
   '   GoTo 99
   'End If
   '
   VDEF$ = REGBLAN$("MASTER")
   CINUE% = 1 + NUMAS%
   Call REPLA(VDEF$, MKI$(CINUE%), 1, 2)
   Call REPLA(VDEF$, MKI$(GRUCOVEI%(CII%)), 3, 2)
   Call REPLA(VDEF$, DESCRIT0$(CII%), 5, 64)
   Call REPLA(VDEF$, UNIMED$(CII%), 109, 2)
   Call REPLA(VDEF$, Chr$(TIMATE%(CII%)), 93, 1)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   '
   Call ABREMASTER
   If OKX% = 1 Then
'     CACAM% = ARTICULOS.TableDefs("Master").Fields.Count
'     Master.FindFirst ("CODINT = " & CII%)
     
'     If Master.NoMatch = False Then
'       Master.Edit
       
'       For KU% = 1 To CACAM%
'         FIELDVAL(KU%) = Master.Fields(KU% - 1).Value
'       Next KU%
'       '
'       Master.FindFirst ("CODINT = " & CINUE%)
'       If Master.NoMatch = False Then
'         Master.Edit
'         For KU% = 1 To CACAM%
'           NOCAMPO$ = TRIM$(UCase$(Master.Fields(KU% - 1).Name))
'           If NOCAMPO$ <> "CODIGO" Then
'             If NOCAMPO$ <> "DESCRIPCION" Then
'               If NOCAMPO$ <> "CODINT" Then
'                 If NOCAMPO$ <> "ALIAS" Then
'                   If NOCAMPO$ <> "UMEDIDA" Then
'                     If NOCAMPO$ <> "GRUCOVEN" Then
'                      Master.Fields(KU% - 1).Value = FIELDVAL(KU%)
'                     End If
'                   End If
'                 End If
'               End If
'             End If
'           End If
'         Next KU%
'         Master.Update
'       End If
       '
       ' Lectura codigo anterior
       CONDI$ = "CODINT = " & CII%
       DESCRI$ = VALOBADA("MASTER", "DESCRIPCION", CONDI$, "NOMESS")
       CODINTE% = XVALO(VALOBADA("MASTER", "CODINT", CONDI$, "NOMESS"))
       Alias$ = VALOBADA("MASTER", "ALIAS", CONDI$, "NOMESS")
       UMED$ = VALOBADA("MASTER", "UMEDIDA", CONDI$, "NOMESS")
       'COMENTADO POR QUE EN EL ALTA YA GUARDA EL QUE SE ESTA EN EL OBJETO QUE SE PUEDE EDITAR.
       'GRUVEN% = XVALO(VALOBADA("MASTER", "GRUCOVEN", CONDI$, "NOMESS"))
       ' Grabacion de codigo nuevo en Tabla Master
       condi2$ = "CODINT = " & CINUE%
       Call ACTUREGISTRO("MASTER", "ALIAS", condi2$, Alias$, REG&, "NOMESS")
       Call ACTUREGISTRO("MASTER", "UMEDIDA", condi2$, UMED$, REG&, "NOMESS")
       
       'COMENTADO POR QUE CUANDO DA EL ALTA SE PUEDE CAMBIAR EL GRUPO CONTABLE Y ACA ESTA TOMANDO EL DEL CODIGO ORIGINAL
       'Call ACTUREGISTRO("MASTER", "GRUCOVEN", condi2$, GRUVEN%, REG&, "NOMESS")
       
       PORCOM = XVALO(Text42) ' se modifico por que estaba mal el tipo de dato es tipo 6
       Call ACTUREGISTRO("MASTER", "PorComVen", condi2$, PORCOM, REG&, "NOMESS")
       ' Grabacion en Tabla Secoper
       If Text2 <> "" And XVALO(Text2) > 0.005 Then
            NOPYI% = 10
            CONDI$ = "CODICONJ =" & CINUE% & " and NumeOper = " & NOPYI%
            If CantRegistros("SECOPER", CONDI$) < 1 Then
100           ABUS$ = RANDSTRING$(-10)
              CONDI$ = "DESCOPER = '" & ABUS$ & "'"
              If CantRegistros("SECOPER", CONDI$) > 0 Then GoTo 100
              Call CreaRegistro("SECOPER", "DESCOPER", ABUS$, "NOMESS")
            End If
            Call ACTUREGISTRO("SECOPER", "CODICONJ", CONDI$, CINUE%, REG&, "NOMESS")
            MINUST# = XVALO(Text2)
            '
            Call ACTUREGISTRO("SECOPER", "NUMEOPER", CONDI$, NOPYI%, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "statoper", CONDI$, 1, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "CodMaq0", CONDI$, "", REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "CodMaq2", CONDI$, "", REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "CodMaq3", CONDI$, "", REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "HorsPrep", CONDI$, 0, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "HorsFijo", CONDI$, 0, REG&, "NOMESS")
            'SE AGREGO EL FORMATNUM POR QUE NO ESTABA GUARDANDO EN LOS CASOS CON DECIMALES POR QUE EL MISMO TENIA UNA COMA
            Call ACTUREGISTRO("SECOPER", "MinuStan", CONDI$, FORMATNUM$(MINUST#, "F16.6"), REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "HorsLimp", CONDI$, 0, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "HorsMHer", CONDI$, 0, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "PieCiclo", CONDI$, 0, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "PorDemor", CONDI$, 0, REG&, "NOMESS")
            If MINUST# > 0.005 Then
                 Caden% = 60 / MINUST#
                 Call ACTUREGISTRO("SECOPER", "Cadencia", CONDI$, Caden%, REG&, "NOMESS")
            End If
            Call ACTUREGISTRO("SECOPER", "CantiOps", CONDI$, 1, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "CategOps", CONDI$, 0, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "CoefSolp", CONDI$, 0, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "LOTESTAN", CONDI$, 1, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "ProgCNum", CONDI$, "", REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "HoProFab", CONDI$, "", REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "ProTerOp", CONDI$, 0, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "PreUnid", CONDI$, 0, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "MonePrec", CONDI$, 0, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "FechPrec", CONDI$, 0, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "OBSEOPER", CONDI$, "", REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "NUOPER_PREP", CONDI$, 0, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "CATOPER_PREP ", CONDI$, 0, REG&, "NOMESS")
            Call ACTUREGISTRO("SECOPER", "DESCOPER", CONDI$, "Mano de Obra Directa", REG&, "NOMESS")
       End If
       Call COPYFORMU(CII%, CINUE%)
       ' COMENTADA POR QUE YA SE ESTA CREANDO AQUI LA SECUENCIA DE OPERACIONES
       'Call COPYRUTA(CII%, CINUE%)
'     End If
     '
     CODINUE$ = Mid$(VDEF$, 77, 16)
     Text1.TEXT = CODINUE$
     '
   End If
   '
99 Call GENECOMAS
   Command22.Enabled = True
   '10/04/07 (FIN)
   '
End Sub

Private Sub Command23_Click()
   '
   '10/04/07 (OT-07-0137)
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   Command23.Enabled = False
   If VALCODART% < 1 Then GoTo 99
   '
   VDEF$ = ""
10 OBX$ = OBJPLA$("ENTRACODNUE", VDEF$)
   If OBX$ = "" Then GoTo 99
   '
   VDEF$ = OBX$
   CODINUE$ = TRIM$(OBX$)
   CODIVIE$ = AJUSTI$(Text1.TEXT, 16)
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > 15 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
      GoTo 10
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   CI% = CODINT%(Text1.TEXT)
   If CI% > 0 Then
'    Master.FindFirst ("CODINT = " & CI%)
     SQLX$ = " SELECT * FROM Master"
     SQLX$ = SQLX$ + " WHERE CODINT = " & CI% & ""
     Dim MASTTTER As ADODB.Recordset
     Set MASTTTER = New ADODB.Recordset
     MASTTTER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     On Error Resume Next
     With MASTTTER
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         !CODIGO = CODINUE$
         .Update
'        Master.Edit
'        Master!CODIGO = CODINUE$
'        Master.Update
         UPDATMASTER% = 1
       End If
     End With
     Set MASTTTER = Nothing
   End If
   '
   Text1.TEXT = CODINUE$
   '
99 Call GENECOMAS
   Screen.MousePointer = 1
   Command23.Enabled = True
   '10/04/07 (FIN)
   '
End Sub

Private Sub Command24_Click()
   '
   '10/04/07 (OT-07-0137)
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command24.Enabled = False
   '
   If VALCODART% < 1 Then
      Text1.SetFocus
      GoTo 99
   End If
   '
   CII% = CODINT%(Text1.TEXT)
   COD$ = AJUSTI$(CODEXT$(CII%), 15)
   '
   If COMALTER%(" Baja Lógica del Codigo '" + TRIM$(COD$) + "' \" + TRIM$(DESCRIT$(CI%)) + "\\Cancelar\Confirmar") <> 2 Then GoTo 99
   '
   Screen.MousePointer = 11
   CII% = CODINT%(Text1.TEXT)
   If CII% > 0 Then
'     Master.FindFirst ("CODINT = " & CI%)
'     Master.Edit
'     Master!Status = (-1)
'     Master.Update
'     UPDATMASTER% = 1
      SQLX$ = "UPDATE Master SET Status = (-1)"
      SQLX$ = SQLX$ + " WHERE CODINT = " & CI% & ""
      FLEXCONN.Execute (FILTSQL$(SQLX$))
      UPDATMASTER% = 1
      Call BLANFORMU
      Text1 = ""
   End If
   '
99 Call GENECOMAS
   '
   TTXX$ = "Baja Lógica del Código '" + TRIM$(COD$) + "' Completada."
   TTXX$ = TTXX$ + "\  \¿ Desea Eliminar el Artículo\de Todas las Listas de Precios ?\  \(Esta Operación no Podrá Revertirse)"
   If COMALTER%(TTXX$ + "\\No, Continuar\Si, Eliminarlo") = 2 Then
      '
      CONDI$ = "COD_PIEZA = " & CI%
      Call BORRAREGISTROS("PREVENTA", CONDI$, REGAF&, "NOMESS")
      '
'      NULIS& = ULTREG&("LISTAS")
'      For U& = 1 To NULIS&
'         Screen.MousePointer = 11
'         LISNO% = U&
'         ALISTA$ = TRIM$(Str$(LISNO%)): While Len(ALISTA$) < 3
'                                          ALISTA$ = "0" + ALISTA$
'                                        Wend
'         APRECIO$ = "LIPRE" + ALISTA$
'         If ULTREG&(APRECIO$) > 0 Then
'            Call FILTERDELETE(APRECIO$, 1, MKI$(CII%))
'         End If
'      Next U&
   End If
   '
   Screen.MousePointer = 1
   Command24.Enabled = True
   '10/04/07 (FIN)
   '
End Sub
Private Sub Command3_Click()
   '
   CI% = CODINT%(Text1)
   If CI% > 0 Then
     If MODIFIC%(0) > 0 Then
       If COMALTER%("Se ha Modificado Datos del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") = 1 Then
           Call Command1_Click
       End If
       MODIFIC%(0) = 0
     End If
   End If
   '
   Call CIERRARCH("*.*")
   Call FINAL("")     '30/03/07 (OT-07-0137)
   '
End Sub

Private Sub Command37_Click()
   '
   '09/04/07 (OT-07-0137)
   Command37.Enabled = False
   If DERACCE%("AJUPARCO", "Ajustar Parámetros de Cálculo de Costos") >= 2 Then
      STPCOS07.Show 1
      Call RECONECTA
   End If
   Command37.Enabled = True
   '09/04/07 (FIN)
   '
End Sub

Private Sub Command38_Click()
   '
   Command38.Enabled = False
   OPLIMAS07.Show 1
   Command38.Enabled = True
   '
End Sub

Private Sub Command39_Click()
   '
   '10/04/07 (OT-07-0137)
   Command39.Enabled = False
    OPT% = COMALTER%("Seleccione una Opción:\\Imprimir\Enviar por Maíl")
    If OPT% = 1 Then
        SENDBYMAIL$ = ""
      ElseIf OPT% = 2 Then
        SENDBYMAIL$ = ""
10      Call SELECHO("DEUDOR1")
        NCX1% = XVALO(VALACT1$("DEUDOR1"))
        If NCX1% >= 0 Then
           SBM$ = EMAILCLI$(NCX1%)
           If NCX1% = 0 Then SBM$ = InputBox("Ingrese Dirección de E-mail", "Dirección de E-mail")
           If SBM$ = "" Or InStr(SBM$, "@") < 2 Then
               SENDBYMAIL$ = ""
               Call MENSERR(24, "Ese Cliente no Tiene Dirección de Mail.")
             Else
               SENDBYMAIL$ = SBM$
           End If
        End If
      ElseIf OPT% = 0 Then GoTo 99
    End If
   Call PRIANACOS
   '
99 Command39.Enabled = True
   '10/04/07 (OT-07-0137)
   '
End Sub

Private Sub Command5_Click()
   '
   ''30/03/07 (OT-07-0137)
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text6.TEXT = VDEVV$
        FF% = FECHANUM(Text6.TEXT)
    End If
   ''30/03/07 (FIN)
   '
End Sub

Private Sub Command6_Click()
    '
    OPT% = 0
10  OPT% = ALTERNA%("Actualización de Costos por Proveedor\Recálculo de Items Recaudados\ - \Acceso a Archivos Maestros\Grupos Contables de Venta\Listas de Precios de Venta\Acceso a Presupuestos\Estadísticas de Cambios de Precios\Generar Mail de Cotización\Recálculo de Compras y Pedidos \-\Otros Accesos Directos")
    If OPT% = 1 Then
         Call ACTUCOSTO_PROVEEDOR
       ElseIf OPT% = 2 Then
         Call SEGUITRE
         GoTo 10
       ElseIf OPT% = 4 Then
         Call ARCHMAES
       ElseIf OPT% = 5 Then
         Call GRUPCONT
       ElseIf OPT% = 6 Then
         Call LISTAPRE
       ElseIf OPT% = 7 Then
         Call CONECRUN("PRESUP07", "Modulo de Presupuestos")
       ElseIf OPT% = 8 Then
         Call REPOCAMBIOS
            GoTo 10
       ElseIf OPT% = 9 Then
         Call MAILCOTI
            GoTo 10
       ElseIf OPT% = 10 Then
         Call RECALCOMPED
            GoTo 10
       ElseIf OPT% = 11 Then
         Call ACCDIR("ANACOS")
    End If
   '
End Sub
Sub RECALCOMPED()
    'ATENCION ++++++++++++++++++++++++++
    'LA CONSULTA CONSIDERA LA FECHA DE ENTREGA DEL DETALLE
    'DONDE PUEDE SER QUE DE N MISMO PEDIDO HAYA FECHAS DIFRENTES
    'POR LO CUAL SOLO VAN A ENTRAR LAS QUE CORRESPONDAN AL RANGO SELECCIONADO
    'EJEMPLO: UN PEDIDO TIENE DEL CODIGO 1 10 UNIDADES SOLICITADAS POR UN IMPORTE DE $150
    'CON FECHA SOLICITADA PARA EL 30 DE ENERO
    'PARA EL CODIGO 2 20 UNIDADES POR UN IMPORTE DE $400 PERO CON FECHA SOLIC. PARA EL 5 DE FEBRERO
    'SI EL RANGO ES PARA EL MES DE ENERO VA A TRAER QUE PARA DICHO PEDIDO UN IMPORTE DE $150
5   R1% = COMALTER%("Elija Opción\\Pedidos\Compras")
    If R1% = 0 Then Exit Sub
    PEDIDO% = 0: compra% = 0
    If R1% = 1 Then PEDIDO% = 1
    If R1% = 2 Then compra% = 1
10  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 5
    End If
    '
    VDEF$ = OBX$
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    RECALPEND09.Label2.Caption = PERIO$
    '
    Screen.MousePointer = 11
    Campos$ = "N.Ped./F6.0;N.Cte./F6.0;R.Social/A45;F.Ent.Sol/D10;Importe/F12.2;M/A3"
    Call CARGA_ENCABEZADO(RECALPEND09.MSF1, Campos$, RECALPEND09, 1)
    Call CARGA_ENCABEZADO(RECALPEND09.MSF_2, Campos$, RECALPEND09)
    '
    RECALPEND09.MSF1.MergeCells = flexMergeRestrictColumns
    RECALPEND09.MSF1.MergeCol(1) = True
    RECALPEND09.MSF2.MergeCells = flexMergeRestrictColumns
    RECALPEND09.MSF2.MergeCol(1) = True

    Call CARGA_ENCABEZADO(RECALPEND09.MSF2, Campos$, RECALPEND09)
    '
    DESDE = FETEXCOR1$(DES%)
    HASTA = FETEXCOR1$(HAS%)
    '
    If PEDIDO% = 1 Then
        SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDDETA.FeSolEnt, PEDDETA.CodiInt,PEDDETA.TIPOCAM"
        SQLX$ = SQLX$ + " , PEDDETA.CantEnt As CantE, PEDDETA.CanPed as Cantp ,PEDDETA.PreUnid,PEDDETA.MONEEMIS "
        SQLX$ = SQLX$ + " FROM PEDENCA "
        SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
        SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
        SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
        SQLX$ = SQLX$ + " AND (PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) OR PEDDETA.CanPed < 0.05) "   ' para incluir pedidos abiertos
        SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
        SQLX$ = SQLX$ + " AND PEDDETA.FeSolEnt >=  '" & DESDE & "'"
        SQLX$ = SQLX$ + " AND PEDDETA.FeSolEnt <=  '" & HASTA & "'"
        
'        SQLX$ = SQLX$ + " GROUP BY  PEDENCA.NroPed,PEDENCA.NroClie,PEDDETA.FeSolEnt,PEDDETA.CodiInt,PEDDETA.PreUnid"
        SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
        Dim RS As New ADODB.Recordset
        Set RS = FLEXCONN.Execute(FILTSQL$(SQLX$))
        Label15 = ""
        '
'        RECALPEND09.MSF1.Rows = 1
'        RECALPEND09.MSF2.Rows = 1
        j& = 0
        With RS
          Do While Not .EOF
            NPED& = XVALO(!NROPED)
            '
            '
           If NPED_ANT& <> NPED& Then
               j& = j& + 1
               RECALPEND09.MSF1.Rows = j& + 1
               ACUMUIMPO = 0
            End If

            NC% = XVALO(!NroClie)
            RACLI$ = Left$(RASOCLI$(NC%), 44)
            FEENT = FECHATEX$(CVINT(!FeSolEnt))
            CANTIPEND = XVALO(!Cantp) - XVALO(!CantE): If CANTIPEND < 1 Then CANTIPEND = 0
            TIPCAMB = XVALO(!TIPOCAM)
            If TIPCAMB < 0.05 Then TIPCAMB = 1
            PREUNITPESOS = XVALO(!PreUnid) * TIPCAMB
            IMPO = CANTIPEND * PREUNITPESOS
            ACUMUIMPO = ACUMUIMPO + IMPO
            NPED_ANT& = NPED&
            '
            RECALPEND09.MSF1.TextMatrix(j&, 1) = TRIM(Str$(NPED_ANT&))
            RECALPEND09.MSF1.TextMatrix(j&, 2) = TRIM$(Str(NC%))
            RECALPEND09.MSF1.TextMatrix(j&, 3) = RACLI$
            RECALPEND09.MSF1.TextMatrix(j&, 4) = FEENT
            RECALPEND09.MSF1.TextMatrix(j&, 5) = FORMATNUM$(ACUMUIMPO, "F12.2")
            RECALPEND09.MSF1.TextMatrix(j&, 6) = "$"
            RECALPEND09.Label8 = FORMATNUM$((XVALO(RECALPEND09.Label8) + IMPO), "F12.2")
'            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(RECALPEND09.MSF1, COLOR1, j&)
            .MoveNext
          Loop
        End With
        RS.Close
        Set RS = Nothing
        j& = j& + 1
    ElseIf compra% = 1 Then
        SQLX$ = "SELECT OCOMDETA.Nume_Ocom,OCOMDETA.NPRO_OCOM,OCOMDETA.Fentre_Sol,OCOMDETA.Cod_Inte,"
        SQLX$ = SQLX$ + " OCOMDETA.VALO_UNID,OCOMDETA.Cant_Rec,OCOMDETA.Cant_Ocom, prun_neto,ocomdeta.mone_cos "
        SQLX$ = SQLX$ + " FROM OCOMENCA INNER JOIN OCOMDETA "
        SQLX$ = SQLX$ + " ON OCOMENCA.NUME_OCOM = OCOMDETA.NUME_OCOM "
        SQLX$ = SQLX$ + " WHERE OCOMDETA.Stat_Ocom < 8 "  ' Excluyo las que estan anuladas O COMPLIDAS
        SQLX$ = SQLX$ + " AND OCOMDETA.Cant_Rec < OCOMDETA.Cant_Ocom - 0.05"
        SQLX$ = SQLX$ + " AND OCOMENCA.Stat_Ocom = 0 "
        SQLX$ = SQLX$ & " AND OCOMENCA.Nume_Ocom > 0 "
        SQLX$ = SQLX$ + " AND OCOMDETA.Fentre_Sol >=  '" & DESDE & "'"
        SQLX$ = SQLX$ + " AND OCOMDETA.Fentre_Sol <=  '" & HASTA & "'"
        SQLX$ = SQLX$ + " ORDER BY OCOMENCA.Nume_Ocom  ASC"
        '
'        SQLX$ = SQLX$ + " GROUP BY  OCOMDETA.Nume_Ocom,OCOMDETA.NPRO_OCOM,OCOMDETA.Fentre_Sol,OCOMDETA.Cod_Inte,OCOMDETA.Pre_Unit"
        Dim RS1 As New ADODB.Recordset
        Set RS1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
        Label15 = ""
        j& = 0
'        RECALPEND09.MSF1.Rows = 1
'        RECALPEND09.MSF2.Rows = 1
        VUELTA% = 0
        With RS1
          Do While Not .EOF
            NPED& = XVALO(!Nume_Ocom)
            '

            If NPED_ANT& <> NPED& Then
               j& = j& + 1
               RECALPEND09.MSF1.Rows = j& + 1
               ACUMUIMPO = 0
            End If
            '

            NC% = XVALO(!Npro_Ocom)
            RACLI$ = Left$(RASOCLI$(-1 * NC%), 44)
            FEENT = FECHATEX$(CVINT(!Fentre_Sol))
            CANTIPEND = XVALO(!cant_ocom) - XVALO(!Cant_Rec): If CANTIPEND < 1 Then CANTIPEND = 0
            PREUNIT1 = XVALO(!prun_neto)
            IMPO = CANTIPEND * PREUNIT1
            ACUMUIMPO = ACUMUIMPO + IMPO
            NPED_ANT& = NPED&
            '
            RECALPEND09.MSF1.TextMatrix(j&, 1) = TRIM(Str$(NPED_ANT&))
            RECALPEND09.MSF1.TextMatrix(j&, 2) = TRIM$(Str(NC%))
            RECALPEND09.MSF1.TextMatrix(j&, 3) = RACLI$
            RECALPEND09.MSF1.TextMatrix(j&, 4) = FEENT
            RECALPEND09.MSF1.TextMatrix(j&, 5) = FORMATNUM$(ACUMUIMPO, "F12.2")
            
             MONX% = XVALO(!mone_cos)
             SIMBOX$ = TRIM$(SIMBOLPES$(MONX%))
             If SIMBOX$ = "" Then SIMBOX$ = "$"
             
            RECALPEND09.MSF1.TextMatrix(j&, 6) = SIMBOX$
            If XVALO(!mone_cos) <= 1 Then
               RECALPEND09.Label8 = FORMATNUM$(XVALO(RECALPEND09.Label8) + IMPO, "F12.2")
               RECALPEND09.Label8.Refresh
            Else
               RECALPEND09.Label9 = FORMATNUM$(XVALO(RECALPEND09.Label9) + IMPO, "F12.2")
               RECALPEND09.Label9.Refresh
            End If
'            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(RECALPEND09.MSF1, COLOR1, j&)
            '
            .MoveNext
          Loop
        End With
        RS1.Close
        Set RS1 = Nothing
    End If
    Screen.MousePointer = 1
    If PEDIDO% = 1 Then RECALPEND09.Caption = "Pedidos"
    If compra% = 1 Then
       RECALPEND09.Caption = "Compras"
       RECALPEND09.Label9.Visible = True
       RECALPEND09.Label10.Visible = True
       RECALPEND09.Label12.Visible = True
       RECALPEND09.Label13.Visible = True
       RECALPEND09.Label6.Left = 5400
       '
       RECALPEND09.Label14.Visible = True
       RECALPEND09.Label15.Visible = True
       RECALPEND09.Label16.Visible = True
       RECALPEND09.Label17.Visible = True
       RECALPEND09.Label18.Visible = True
       RECALPEND09.Label19.Visible = True
       '
       COTIHOY# = TICAMONE(2)
       Label17 = (XVALO(lable13) / COTIHOY#) + XVALO(Label5)
       Label14 = (XVALO(lable9) / COTIHOY#) - XVALO(Label8)

    End If
    RECALPEND09.Refresh
    RECALPEND09.Show 1
    GoTo 10
End Sub



Sub REPOCAMBIOS()
10  OPT% = ALTERNA%("Estadísticas de Cambios en Lista por Artículos\Estadísticas de Cambios Por Proveedor (Costos)")
    If OPT% = 1 Then
         Call ESTAXART
         GoTo 10
       ElseIf OPT% = 2 Then
         Call ESTAXPRV
         GoTo 10
     End If
End Sub
Sub ESTAXPRV()
   
10  Call SELECHO("PROVED1")
    NCX1% = XVALO(VALACT1$("PROVED1"))
    If NCX1% < 1 Then Exit Sub
   '
20 Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
   If PERIO$ = "" Then GoTo 10

   DESFE = FETEXCOR1$(DESDEX%)
   HASFE = FETEXCOR1$(HASTAX%, "HASTA")
   '
   
   'VALIDO SI HAY NOVEDADES CON MAS DE UNA MONEDA
   CONDI$ = " Fech_Mod  >=  '" & DESFE & " '"
   CONDI$ = CONDI$ + " and Fech_Mod  <=  '" & HASFE & " '"
   CONDI$ = CONDI$ + "  AND Lista_Repos = 'COSTO'"
   CONDI$ = CONDI$ + "  AND NumProv  = " & NCX1%
   condi2$ = "  AND MONEDA <= 1"
   condi3$ = "  AND MONEDA > 1"
   If CantRegistros("NOVEDADES", CONDI$ + condi2$, "NOMESS") > 0 Then PESO% = 1
   If CantRegistros("NOVEDADES", CONDI$ + condi3$, "NOMESS") > 0 Then DOL% = 1
   'SI HAY MAS DE UNA MONEDA PREGUNTO CUAL QUIERE VER
   If PESO% + DOL% > 1 Then
60    OPX% = COMALTER%("Se Han Detectado Cambios de Precios en 2 Tipos de Monedas\\Ver Cambios en Pesos\Ver Cambios en Dolares")
      If OPX% < 1 Or OPX% > 2 Then GoTo 20
      TIPOMONE% = OPX%
   ' SI HAY SOLO UNA MONEDA MUESTRA DIRECTAMENTE ESTA.
   Else
      TIPOMONE% = 1
      If DOL% > 0 Then TIPOMONE% = 2
   End If
   '
   REG& = 0
   Call CARGA_ENCA_X_PROV
   '
   SQLX$ = "SELECT CODINT  FROM NOVEDADES WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE   " & CONDI$
   If TIPOMONE% = 1 Then SQLX$ = SQLX$ + " AND MONEDA <= " & TIPOMONE%
   If TIPOMONE% = 2 Then SQLX$ = SQLX$ + " AND MONEDA >= " & TIPOMONE%
   SQLX$ = SQLX$ + " GROUP BY CODINT "
   '
   INFO_GRILLAJM.MSF.Redraw = False
   Screen.MousePointer = 11
   Set RS = READSET(SQLX$)
   With RS
    Do While Not .EOF
      REG& = REG& + 1
      '
      INFO_GRILLAJM.MSF.Rows = REG& + 1
      '
      CI1% = XVALO(!CODINT)
      'TOMO LA FECHA PRIMERA DENTRO DEL RANGO
      CONDI1$ = "CODINT = " & CI1%
      CONDI1$ = CONDI1$ + " and Fech_Mod  >=  '" & DESFE & " '"
      CONDI1$ = CONDI1$ + " and Fech_Mod  <=  '" & HASFE & " '"
      CONDI1$ = CONDI1$ + "  AND NumProv  = " & NCX1%
      CONDI1$ = CONDI1$ + " ORDER BY Fech_Mod ASC"
      FE1% = CVINT(VALOBADA("NOVEDADES", "FECH_MOD", CONDI1$, "NOMESS"))
      If FE1% < 1 Then
        Call COMUNI("No Se Detectaron Cambios en el Rango Seleccionado")
        GoTo 20
      End If
      '
      'LEO EL  PRIMER PRECIO
      condi2$ = "CODINT = " & CI1%
      condi2$ = condi2$ + " and Fech_Mod  >=  '" & FETEXCOR1$(FE1%) & " '"
      condi2$ = condi2$ + " and Fech_Mod  <=  '" & FETEXCOR1$(FE1%, "HASTA") & " '"
      condi2$ = condi2$ + "  AND NumProv  = " & NCX1%
      condi2$ = condi2$ + " ORDER BY Fech_Mod ASC"
      PRECIO1# = XVALO(VALOBADA("NOVEDADES", "Prec_Ant", condi2$, "NOMESS"))
      COTIZA1 = XVALO(VALOBADA("NOVEDADES", "COTIZA_DOLAR", condi2$, "NOMESS"))
      If COTIZA1 < 0.005 Then COTIZA1 = 1
      '
      'LEO LA ULTIAMA FECHA DENTRO DEL RANGO
      condi3$ = "CODINT = " & CI1%
      condi3$ = condi3$ + " and Fech_Mod  >=  '" & DESFE & " '"
      condi3$ = condi3$ + " and Fech_Mod  <=  '" & HASFE & " '"
      condi3$ = condi3$ + "  AND NumProv  = " & NCX1%
      condi3$ = condi3$ + " ORDER BY Fech_Mod DESC"
      FE2% = CVINT(VALOBADA("NOVEDADES", "FECH_MOD", condi3$, "NOMESS"))
      '
      'LEO EL ULTIMO  PRECIO
      CONDI4$ = "CODINT = " & CI1%
      CONDI4$ = CONDI4$ + " and Fech_Mod  >=  '" & FETEXCOR1$(FE2%) & " '"
      CONDI4$ = CONDI4$ + " and Fech_Mod  <=  '" & FETEXCOR1$(FE2%, "HASTA") & " '"
      CONDI4$ = CONDI4$ + "  AND NumProv  = " & NCX1%
      CONDI4$ = CONDI4$ + " ORDER BY Fech_Mod DESC"

      PRECIO2# = XVALO(VALOBADA("NOVEDADES", "Prec_New", CONDI4$, "NOMESS"))
      COTIZA2 = XVALO(VALOBADA("NOVEDADES", "COTIZA_DOLAR", CONDI4$, "NOMESS"))
      If COTIZA2 < 0.005 Then COTIZA2 = 1
      '
      'CALCULO EL PORCENTAJE DE MODIFICACION ENTRE EL PRIMER VALOR Y EL ULTIMO SEGUN RANGO
      If PRECIO1# > 0.005 Then
         PJE = ((PRECIO2# / PRECIO1#) - 1) * 100
      ElseIf PRECIO2# < 0.005 Then
         PJE = -100
      'CASO QUE SEA INGRESO DE UN ITEM QUE SE CARGA EL PRECIO POR PRIMERA VEZ  X Q' CUANDO SE GRABE EN LA TABLA NOVEDADES NO VA A TRAER NINGUN COSTO ANTERIOR
      ElseIf PRECIO1# < 0.005 Then
         PJE = 0
      End If
      '
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 1) = CODEXT$(CI1%)
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 2) = DESCRIT0$(CI1%)
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 3) = FECHATEX$(FE1%)
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 4) = FORMATNUM$(PRECIO1# / COTIZA1, "F12.2")
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 5) = FECHATEX$(FE2%)
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 6) = FORMATNUM$(PRECIO2# / COTIZA2, "F12.2")
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 7) = FORMATNUM$(PJE, "F7.2")
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 8) = FORMATNUM$(TIPOMONE%, "F2.0")
      '
      .MoveNext
    Loop
   End With
   
   RS.Close
   Set RS = Nothing
   Screen.MousePointer = 1

   INFO_GRILLAJM.MSF.Redraw = True
   
   If REG& < 1 Then
      Call COMUNI("No Se Detectaron Cambios en el Rango Seleccionado")
      GoTo 20
   End If

   If TIPOMONE% < 1 Then TIPOMONE% = 1
   DEMONE$ = DEMONECO$(TIPOMONE%)

   INFO_GRILLAJM.Caption = "Cambios de Precios en " & DEMONE$ & " - Prov.: " & RASOCLI(NCX1% * -1) & " - Desde: " & FECHATEX$(DESDEX%) & " - Hasta: " & FECHATEX$(HASTAX%)
   Call COLOREAR_GRILLASQL(INFO_GRILLAJM.MSF, &HC0E0FF)
   INFO_GRILLAJM.NPROV% = NCX1%
   INFO_GRILLAJM.Command7.Visible = True
   INFO_GRILLAJM.Show 1
   
   If PESO% + DOL% > 1 Then GoTo 60
   
   GoTo 20
End Sub



Sub CARGA_ENCA_X_PROV()
     With INFO_GRILLAJM
        .MSF.Row = 1
        .MSF.Cols = 9
        NUCOLU% = 1: .TEPRUEBA = "90": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = ""
        NUCOLU% = 2: .TEPRUEBA = "12345678901234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Còdigo"
        NUCOLU% = 3: .TEPRUEBA = "123456789012345678901234567890123": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Descripciòn"
        
        NUCOLU% = 4: .TEPRUEBA = "1234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha 1"
        
        NUCOLU% = 5: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "$.Ant"
        NUCOLU% = 6: .TEPRUEBA = "1234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha 2"
        
        NUCOLU% = 7: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "$.Nuevo"
        NUCOLU% = 8: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "%"
        '
        NUCOLU% = 9: .TEPRUEBA = "12345": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Mon"
        '
    End With
                                                            
End Sub
Sub ESTAXART(Optional CI1%)
   If CI1% < 1 Then
     Call SELECHO("MASTER")
     COD$ = TRIM$(VALACT1$("MASTER"))
     CI1% = CODINT%(COD$)
   End If
   
   If CI1% < 1 Then Exit Sub
   '
   
20 Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
   If PERIO$ = "" Then Exit Sub

   DESFE = FETEXCOR1$(DESDEX%)
   HASFE = FETEXCOR1$(HASTAX%, "HASTA")
   '
   'VALIDO SI HAY NOVEDADES CON MAS DE UNA MONEDA
   CONDI$ = " Fech_Mod  >=  '" & DESFE & " '"
   CONDI$ = CONDI$ + " and Fech_Mod  <=  '" & HASFE & " '"
   CONDI$ = CONDI$ + "  AND Lista_Repos = 'LISTA'"
   CONDI$ = CONDI$ + "  AND CODINT = " & CI1%
   condi2$ = "  AND MONEDA <= 1"
   condi3$ = "  AND MONEDA > 1"
   PESO% = 0: DOL% = 0
   If CantRegistros("NOVEDADES", CONDI$ + condi2$, "NOMESS") > 0 Then PESO% = 1
   If CantRegistros("NOVEDADES", CONDI$ + condi3$, "NOMESS") > 0 Then DOL% = 1
   'SI HAY MAS DE UNA MONEDA PREGUNTO CUAL QUIERE VER
   If PESO% + DOL% > 1 Then
60    OPX% = COMALTER%("Se Han Detectado Cambios de Precios en 2 Tipos de Monedas\\Ver Cambios en Pesos\Ver Cambios en Dolares")
      If OPX% < 1 Or OPX% > 2 Then GoTo 20
      TIPOMONE% = OPX%
   ' SI HAY SOLO UNA MONEDA MUESTRA DIRECTAMENTE ESTA.
   Else
      TIPOMONE% = 1
      If DOL% > 0 Then TIPOMONE% = 2
   End If
   SQLX$ = "SELECT  * FROM NOVEDADES WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE   " & CONDI$
   If TIPOMONE% = 1 Then SQLX$ = SQLX$ + " AND MONEDA <= " & TIPOMONE%
   If TIPOMONE% = 2 Then SQLX$ = SQLX$ + " AND MONEDA >= " & TIPOMONE%
   SQLX$ = SQLX$ + " ORDER BY FECH_MOD"
   '
   INFO_GRILLA07.MSF.Redraw = False
   Screen.MousePointer = 11
   Call CARGA_ENCABEZADO_X_ART
   REG& = 0
   Set RS = READSET(SQLX$)
   With RS
    Do While Not .EOF
      REG& = REG& + 1
      '
      INFO_GRILLA07.MSF.Rows = REG& + 1
      Call TRACE(24, REG&, FIN&)
      '
      INFO_GRILLA07.MSF.TextMatrix(REG&, 1) = FECHATEX$(CVINT(!Fech_Mod))
      P_ANT# = XVALO(!Prec_Ant)
      P_NEW# = XVALO(!Prec_New)
      INFO_GRILLA07.MSF.TextMatrix(REG&, 2) = FORMATNUM$(P_ANT#, "F12.2")
      INFO_GRILLA07.MSF.TextMatrix(REG&, 3) = FORMATNUM$(P_NEW#, "F12.2")
       PJE = ((P_NEW# / P_ANT#) - 1) * 100
      INFO_GRILLA07.MSF.TextMatrix(REG&, 4) = FORMATNUM$(PJE, "F6.2")
      INFO_GRILLA07.MSF.TextMatrix(REG&, 5) = FORMATNUM$(TIPOMONE%, "F2.0")
       
      INFO_GRILLA07.MSF.TextMatrix(REG&, 6) = TRIM$(ECOARCH$("FLEXUSR", MKI$(XVALO(!Nume_Usuar))))
      INFO_GRILLA07.MSF.TextMatrix(REG&, 7) = SAFETEXT$(!Observacion)

      .MoveNext
    Loop
   End With
   
   RS.Close
   Set RS = Nothing
   Screen.MousePointer = 1
   If REG& < 1 Then
      Call COMUNI("No Se Detectaron Cambios en el Rango Seleccionado")
      GoTo 20
   End If
   
   INFO_GRILLA07.MSF.Redraw = True
   
   If TIPOMONE% < 1 Then TIPOMONE% = 1
   DEMONE$ = DEMONECO$(TIPOMONE%)
   INFO_GRILLA07.Caption = "Cambios de Precios en " & DEMONE$ & "  - Art.: " & CODEXT$(CI1%) & " - Desde: " & FECHATEX$(DESDEX%) & " - Hasta: " & FECHATEX$(HASTAX%)
   Call COLOREAR_GRILLASQL(INFO_GRILLA07.MSF, &HEFF1F1)
   '
   INFO_GRILLA07.Command7.Visible = True
   INFO_GRILLA07.CI1X = CI1%
   '
   INFO_GRILLA07.Show 1
   '
   If PESO% + DOL% > 1 Then GoTo 60
   '
   GoTo 20
   
End Sub

   Sub CARGA_ENCABEZADO_X_ART()
   
     With INFO_GRILLA07
        .MSF.Row = 1
        .MSF.Cols = 8
        NUCOLU% = 1: .TEPRUEBA = "90": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = ""
        '
        NUCOLU% = 2: .TEPRUEBA = "1234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha"
        NUCOLU% = 3: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "$.Ant"
        NUCOLU% = 4: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "$.Nuevo"
        NUCOLU% = 5: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "%"
        '
        NUCOLU% = 6: .TEPRUEBA = "123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Mon"
        NUCOLU% = 7: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Usuario"
        NUCOLU% = 8: .TEPRUEBA = "1234567890123456789+123456789+123456789+123456789+123456789+": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Observación"
        '
     End With
End Sub
Private Sub Command7_Click()
   Command7.Enabled = False
   Call CONECRUN("PLANEL04", "Planera Electronica")
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   '
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text7.TEXT = VDEVV$
        FF% = FECHANUM(Text7.TEXT)
    End If
    '
End Sub

Private Sub Command9_Click()
   '
   Command9.Enabled = False
   Text1.Enabled = False
   '
   If MODIFIC%(0) > 0 Then 'VALIDO SI HUBO CAMBIOS
     Call GRADATITEM
     Call GRAFORMULA
     Call GRATIEMOD
   End If
   '
   Text1.Enabled = True
   
'  Call GRADATITEM
   '
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
   Call CALCUCOS '09/04/07 (OT-07-0137)
   
   Command9.Enabled = True
   '
End Sub

Sub GRADATITEM()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   '
   If DERACCE%("MODITMAS", "Modificación Datos Maestro de Stock") >= 2 Then
      Screen.MousePointer = 11
      MODESCRI% = 0
      DE$ = DESCRIT0$(CI%)
      If TRIM$(Text3) <> "" Then
        If TRIM$(Text3) <> DE$ Or TRIM$(Text45) <> NUPLANO$(CI%) Then
          DE$ = TRIM$(Text3)
          MODESCRI% = 1
          UPDATMASTER% = 1
        End If
        Call ABREMASTER
'        Master.FindFirst ("CODINT = " & CI%)
'        Master.Edit
'        Master!Descripcion = DE$
'        Master!NPLANO = TRIM$(Left$(Text45, 24))
'        Master!PorComVen = XVALO(Text42)
'        Master.Update
        ' SE CAMBIA DE LUGAR PARA QUE LA VALIDACION QUEDE AL PRINCIPIO ANTES DE EMPEZAR A GABAR
        CONDI$ = "CODINT = " & CI%
        'GRABACION DE COSTO DE REPOSICION SEGUN EL PRECIO CALCULADO
        If MODIFIC%(0) > 0 Then 'VALIDO SI HUBO CAMBIOS
          'MENSAJE DE CONFIRMACION DE ACTUALIZAR ABM DE ARTICULOS
          If COMALTER%("¿Desea Actualizar El Costo de Reposición? \Para Guardar este Cambio Deberá Contar\ Con el Derecho Correspondiente\\Sí, Actualizar \No, Continuar") = 1 Then
            'VALIDACION DE DERECHOS DE MODIFICAR COSTO DE REPOSICION DESDE COSTOS
            If DERACCE%("MODCOREP", "Modificar Costo de Reposicion Desde Costos") > 1 Then
              'VALOR EN $
              COSTOREPO = XVALO(Label12)
              'SI EL TIPO DE MONEDA ES <> PESOS TOMA EL DEL LABEL CORRESPONDIENTE A DOLARES
              If XVALO(VALOBADA("MASTER", "MONECOS", CONDI, "NOMESS")) > 1 Then COSTOREPO = XVALO(Label3)
              Call ACTUREGISTRO("MASTER", "COSUNI", CONDI$, FORMATNUM$(COSTOREPO, "F16.6"), REG&, "NOMESS")
            End If
          End If
        End If
        '
        Call ACTUREGISTRO("MASTER", "Descripcion", CONDI$, DE$, REG&, "NOMESS")
        Call ACTUREGISTRO("MASTER", "NPLANO", CONDI$, TRIM$(Left$(Text45, 24)), REG&, "NOMESS")
        Call ACTUREGISTRO("MASTER", "PorComVen", CONDI$, XVALO(Text42), REG&, "NOMESS")
        Call ACTUREGISTRO("MASTER", "MANO_OBRA_REAL", CONDI$, FORMATNUM$(XVALO(Text4), "F16.6"), REG&, "NOMESS")
        '
        'VARIABLE BANDERA PARA MODIFICACION DE COSTO DE REPOSICION
        MODIFIC%(0) = 0
        Call GENECOMAS
      End If
   End If
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
End Sub


Private Sub Config_Click()
   '
   Call Command37_Click
   '
End Sub


Private Sub ConRes_Click()
   Call OPLIMAS04.Command3_Click
End Sub


Private Sub Form_Load()
   '
   PROCESOACTUCOSTO% = 0
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
   End If
   '
   Call CENTRAFORM(Me)
   '
   '09/04/07 (OT-07-0137)
   Label11.Caption = FORMATNUM$(TICAMONE(2), "F10.4")
   '
   DCC1$ = TRIM$(Control$("ANACOSTO", "GAGENMAT"))
   Ltext41 = FORMATNUM$(XVALO(DCC1$), "F6.2")
   '
   DCC2$ = TRIM$(Control$("ANACOSTO", "GAGENFAB"))
   Ltext36 = FORMATNUM$(XVALO(DCC2$), "F6.2")
   '
   DCC3$ = TRIM$(Control$("ANACOSTO", "PORUTILI"))
   Ltext39 = FORMATNUM$(XVALO(DCC3$), "F6.2")
   '
   DCC4$ = TRIM$(Control$("ANACOSTO", "PORIMPUE"))
   Ltext43 = FORMATNUM$(XVALO(DCC4$), "F6.2")
   '
   '//OT-07-0477//
   '
   Ltext5 = FORMATNUM$((VALHORA(0) / 60), "F7.4")
   '
   '//OT-07-0477//
   '
   'MODIIICADO PARA QUE SE VALIDE DERECHOS Y SI NO TIENE DESHABILITAR BOTON
   If DERACCE%("MODIMINU/NOMESS", "Modificar Tiempo de Mano de Obra") < 2 Then
     Text2.Enabled = False
     Call COMUNI("Usuario No Habilitado, Para Modificar Tiempo de Mano de Obra")
   End If
   '
   'MODIIICADO PARA QUE SE VALIDE DERECHOS Y SI NO TIENE DESHABILITAR BOTON
   If DERACCE%("MODPORVE/NOMESS", "Modificar Porcentaje de Comisión de Vendedor") < 2 Then
      Text42.Enabled = False
      Call COMUNI("Usuario No Habilitado, Para Modificar Comisión de Vendedor")
   End If
   '
   
   APREPES$ = TRIM$(Control$("ANACOSTO", "LIPREPES"))
'   MsgBox "ATENCION SALTO DE CODIGO "
'   Exit Sub
   If APREPES$ = "" Or XVALO(APREPES$) < 1 Or XVALO(APREPES$) > ULTREG&("LISTAS") Then
      Call MENSERR(24, "Error de Configuración de Listas de Precios")
      Call Command37_Click
      Exit Sub
   End If
   While Len(APREPES$) < 3: APREPES$ = "0" + APREPES$: Wend
   APREPES$ = "LIPRE" + APREPES$
   '
   APREDOL$ = Control$("ANACOSTO", "LIPREDOL")
   If APREDOL$ = "" Or XVALO(APREDOL$) < 1 Or XVALO(APREDOL$) > ULTREG&("LISTAS") Then
      Call MENSERR(24, "Error de Configuración de Listas de Precios")
      Call Command37_Click
      Exit Sub
   End If
   While Len(APREDOL$) < 3: APREDOL$ = "0" + APREDOL$: Wend
   APREDOL$ = "LIPRE" + APREDOL$
   '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Cancel = 0 Then
      Call Command3_Click
   End If
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub

Private Sub Guar_Click()
   Call Command1_Click
End Sub

Private Sub Imp_Click()
   Call Command39_Click
End Sub

Private Sub Label1_DblClick_VIE()
    '
    ' aqui es para mandar el precio a las listas de precios
    '
    CI% = CODINT%(Text1.TEXT)
    If CI% < 1 Then Exit Sub
    '
    If DERACCE%("GRAPREVE", "Grabar Precios de Venta s/Costos") < 2 Then
       Exit Sub
    End If
    '
    HOII% = HOY(HOS$)
    NUPRECIO$ = InputBox$("Nuevo Precio", "Registrar Nuevo Precio en Pesos al " + HOS$, Label1)
    If TRIM$(NUPRECIO$) = "" Then Exit Sub
    If XVALO(NUPRECIO$) < 0.00005 Then Exit Sub
    '
    Label1.Caption = FORMATNUM$(XVALO(NUPRECIO$), "F10.2")
    Label7.Caption = FORMATNUM$(((XVALO(Label1.Caption) * ALIVA) / 100) + XVALO(Label1.Caption), "F10.2")
    Text6 = HOS$ ': Text7 = HOS$
    '
    '\\\OT-08-0095-FL-14/03/08///
    'ACTUALIZA lISTA EN PESOS
    SINLIST% = 0
'    For U& = 1 To ULTREG&("LISTAS")
'        XX$ = REGLEIDO$("LISTAS", U&)
'        If Asc(Mid$(XX$, 50, 1)) = 1 Then           'VALIDA TIPO DE MONEDA = $
'            APREPES1$ = TRIM$(Str$(U&))
'            While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
'            APREPES1$ = "LIPRE" + APREPES1$
'            PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
'            CODIN% = CVI(Mid$(PRE1$, 1, 2))            'VALIDA QUE EL CODIGO PERTENEZCA
'            If CODIN% > 0 Then                         'A LA LISTA DE PRECIOS
'                PREPES# = XVALO(Label1)
'                FEXI% = FECHANUM(Text6)
'                Call REPLA(PRE1$, MKI$(CI%), 1, 2)
'                Call REPLA(PRE1$, MKD$(PREPES#), 62, 8)
'                Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
'                Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
'                SINLIST% = 1
'            End If
'        End If
'    Next U&
    '\\\OT-08-0095-FL-FIN///
    '
    'ACTUALIZA LISTA EN PESOS EN SQL
    SQLX$ = "SELECT * FROM LISTAPRE with(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE moneda = 1 "
    SQLX$ = SQLX$ + " ORDER BY COD_LISTA"
    '
    Dim LISTtmp As New ADODB.Recordset
    Set LISTtmp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With LISTtmp
      '
      Do While Not .EOF
          If CI% > 0 Then
            NUM% = XVALO(!Cod_Lista)
            PREPES# = XVALO(Label1)
            FEXI% = FECHANUM(Text6)
            '
            CONDI$ = "COD_PIEZA = " & CI%
            CONDI$ = CONDI$ + " AND COD_LISTA= " & NUM%
            Call ACTUREGISTRO("PREVENTA", "Precio_Unitario", CONDI$, PREPES#, REG&, "NOMESS")
            Call ACTUREGISTRO("PREVENTA", "Cant_Limite", CONDI$, FEXI%, REG&, "NOMESS")
            
            If REG& > 0 Then
              SINLIST% = 1
              Call ACTUALIZADATLISTA(NUM%)
            End If
          End If
          .MoveNext
      Loop
    End With
    '
    '\\\OT-08-0136-FL-04/04/08///
    If SINLIST% = 0 Then
         Call SELECHO("LISTAS")
         APREPES1$ = TRIM$(VALACT1$("LISTAS"))
         NUMLIS% = XVALO(VALACT1$("LISTAS")) ' NUMERO DE LISTA DE SELECCIONADA
         Label15 = APREPES1$
         Label15.ToolTipText = VALACT2$("LISTAS")
'        While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
'        APREPES1$ = "LIPRE" + APREPES1$
'        PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
         PREPES# = XVALO(Label1)
         FEXI% = FECHANUM(Text6)
'        Call REPLA(PRE1$, MKI$(CI%), 1, 2)
'        Call REPLA(PRE1$, MKD$(PREPES#), 62, 8)
'        Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
'        Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
         '
         If CantRegistros("PREVENTA", CONDI$) < 1 Then
100         ABUS$ = RANDSTRING$(-10)
            CONDI$ = "Num_Original = '" & ABUS$ & "'"
            If CantRegistros("PREVENTA", CONDI$) > 0 Then GoTo 100
            Call CreaRegistro("PREVENTA", "Num_Original", ABUS$, "NOMESS")
         End If
         '
         Call ACTUREGISTRO("PREVENTA", "COD_LISTA", CONDI$, APREPES1$, REG&, "NOMESS")
         '
         Numorden% = XVALO(VALOBADA("PREVENTA", "Max(Num_Orden)", "Cod_Lista =" & APREPES1, "NOMESS")) + 1
         '
         Call ACTUREGISTRO("PREVENTA", "Precio_Unitario", CONDI$, PREPES#, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Num_Orden", CONDI$, Numorden%, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Desc_Pieza", CONDI$, Text3, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Denom_Pieza", CONDI$, Text3, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Carac_Control", CONDI$, "", REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Cant_Limite", CONDI$, FEXI%, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Porc_Limite", CONDI$, 0, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Marborra", CONDI$, "", REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Cod_Pieza", CONDI$, CI%, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Num_Original", CONDI$, "", REG&, "NOMESS")
         'ACTUALIZA LISTA DE SELECCION
         Call ACTUALIZADATLISTA(NUMLIS%)
         '
    End If
    '\\\OT-08-0136-FL-FIN///
    '
    Call CALCUCOS
   
    '
End Sub
Sub Label1_DblClick()
    '
    Command9.Enabled = False
    Text1.Enabled = False

    'VALIDACION DE CODIGO
    CI% = CODINT%(Text1.TEXT)
    If CI% < 1 Then GoTo 99
    '
    'VALIDACION DE DERECHOS
    If DERACCE%("GRAPREVE", "Grabar Precios de Venta s/Costos") < 2 Then
        GoTo 99
    End If
    '
    HOII% = HOY(HOS$)
    '
    'INGRESO DE NUEVO PRECIO Y VALIDACION
30  NUPRECIO$ = InputBox$("Nuevo Precio", "Registrar Nuevo Precio en Pesos al " + HOS$, Label1)
    If TRIM$(NUPRECIO$) = "" Then GoTo 99
    PREC_NVO# = XVALO(NUPRECIO$)
    If PREC_NVO# < 0.00005 Then GoTo 99
    '
    PREC_ACTUAL# = XVALO(Label1.Caption)
    '
    'ACTUALIZALISTA LISTA
    If Abs(PREC_ACTUAL# - PREC_NVO#) > 0.005 Then 'SI CAMBIO EL PRECIO
        OBSERVAC.Caption = "Observación de Cambio de Precio (Mandataria)"
        OBSERVAC.Text28 = ""
        OBSERVAC.Show 1                        'LLAMA A VENTANA DE OBSERVACIONES
        If OBSERVAC.Label1.Caption <> "OK" Or TRIM$(OBSERVAC.Text28) = "" Then
           OBSERVAC.Text28 = ""
           Call COMUNI("Debe Ingresar un Texto Motivo del Cambio de Precio\No se va a Registrar el Mismo")
           GoTo 30                             'SI NO SE INGRESA OBSERVACION VUELVE A LA VENTANA AL INPUT BOX
        End If
        '
        'SI CAMBIO EL PRECIO Y TIENE OBSERVACION ACTUALIZA LA LISTA Y AGREGA LA NOVEDAD
        'ASIGNACION DE PRECIO EN PANTALLA
        Label1.Caption = FORMATNUM$(PREC_NVO#, "F10.2")
        Label7.Caption = FORMATNUM$(((XVALO(Label1.Caption) * ALIVA) / 100) + XVALO(Label1.Caption), "F10.2")
        Text6 = HOS$

        Call ACTUALIZALISTA(1, OK%)
        Observacion$ = OBSERVAC.Text28
        
        If OK% = 1 Then
            'VALIDO SI CAMBIO EL PRECIO EN ESE CASO GRABO EL USUARIO
            
               CONDI$ = "CODINT = " & CI%
               Call ACTUREGISTRO("MASTER", "UsuPrePes", CONDI$, USERNAME$, REGAF&, "NOMESS")
               Label18 = USERNAME$
               '
               Call ADD_NOVEDAD(CI%, 0, PREC_ACTUAL#, PREC_NVO#, Observacion$, 1, "LISTA")
        End If
   
    Else
        'CASO QUE NO CAMBIO EL PRECIO SE VA.
        GoTo 99
    End If

    '

    'REALIZA CALCULO DE COSTOS
    Call CALCUCOS
99  Command9.Enabled = True
    Text1.Enabled = True
   
    '
End Sub
Private Sub Label103_Change()
   '
   
   '
End Sub

Private Sub Label11_DblClick()
   '
   Call CARTAMONED
   Label11.Caption = FORMATNUM$(TICAMONE(-2), "F10.4")
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   '
   List1.Clear
   For U& = 1 To ULTREG&("!BOLREDET")
       Y$ = REGLEIDO$("!BOLREDET", U&)
       CICOMPO% = CVI(Mid$(Y$, 83, 2))
       CANTI# = 0: PUNIT# = 0: PTOTA# = 0
       CANTI# = CVS(Mid$(Y$, 159, 4))
       '
       X$ = Space(128)
       CODX$ = CODEXT$(CICOMPO%)
       Call REPLA(X$, CODX$, 1, 16)
       DESC$ = DESCRIT0$(CICOMPO%)
       Call REPLA(X$, DESC$, 17, 32)
       Call REPLA(X$, UNIMED$(CICOMPO%), 48, 2)
       Call REPLA(X$, FORMATNUM$(CANTI#, "F10.4"), 53, 10)
         MONECO% = MONECOSTO%(CICOMPO%)
       Call REPLA(X$, SIMBOLPES$(MONECO%), 65, 2)
       Call REPLA(X$, FORMATNUM$(COSUNI(CICOMPO%), "F9.2"), 67, 9)
       COSTOPAR = COSPESOS(CICOMPO%) * CANTI#
       Call REPLA(X$, FORMATNUM$(COSTOPAR, "F12.4"), 76, 12)
       List1.AddItem X$
    Next U&
    '
    Call CALCUCOS
    Screen.MousePointer = 1
    '
End Sub




Private Sub Label12_Change()
   'VARIABLE BANDERA PARA MODIFICACION DE COSTO DE REPOSICION EN ABM ARTICULOS
   MODIFIC%(0) = 1
   If PROCESOACTUCOSTO% > 0 Then
        CI1% = CODINT%(Text1)
        CONDI$ = "CODINT = " & CI1%
        'GRABACION DE COSTO DE REPOSICION SEGUN EL PRECIO CALCULADO
        COSTOREPO = XVALO(Label12)
        'SI EL TIPO DE MONEDA ES <> PESOS TOMA EL DEL LABEL CORRESPONDIENTE A DOLARES
        If XVALO(VALOBADA("MASTER", "MONECOS", CONDI, "NOMESS")) > 1 Then COSTOREPO = XVALO(Label3)
        Call ACTUREGISTRO("MASTER", "COSUNI", CONDI$, FORMATNUM$(COSTOREPO, "F16.6"), REG&, "NOMESS")
   End If

End Sub

Private Sub Label21_DblClick()
   Call Picture2_DblClick
End Sub

Private Sub Label27_DblClick()
   '04/03/07 (OT-07-0137)
   Call List1_DblClick
   '04/03/07 (FIN)
End Sub



Private Sub Label6_DblClick_VIE()
   '
   ' aqui es para mandar el precio a las listas de precios
   '
    If XVALO(Label11) < 0.05 Then Exit Sub
    CI% = CODINT%(Text1.TEXT)
    If CI% < 1 Then Exit Sub
    '
    If DERACCE%("GRAPREVE", "Grabar Precios de Venta s/Costos") < 2 Then
       Exit Sub
    End If
    '
    HOII% = HOY(HOS$)
    NUPRECIO$ = InputBox$("Nuevo Precio", "Registrar Nuevo Precio en Dolares al " + HOS$, Label6)
    If TRIM$(NUPRECIO$) = "" Then Exit Sub
    If XVALO(NUPRECIO$) < 0.00005 Then Exit Sub
    '
    Label6.Caption = FORMATNUM$(XVALO(NUPRECIO$), "F10.2")
    Label8.Caption = FORMATNUM$(((XVALO(Label6.Caption) * ALIVA) / 100) + XVALO(Label6.Caption), "F10.2")
    Text7 = HOS$
    '
    '\\\OT-08-0095-FL-14/03/08///
    'ACTUALIZA LISTAS EN DOLARES
'    For U& = 1 To ULTREG&("LISTAS")
'        XX$ = REGLEIDO$("LISTAS", U&)
'        If Asc(Mid$(XX$, 50, 1)) = 2 Then           'VALIDA TIPO DE MONEDA = U$S
'            APREPES1$ = TRIM$(Str$(U&))
'            While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
'            APREPES1$ = "LIPRE" + APREPES1$
'            PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
'            CODIN% = CVI(Mid$(PRE1$, 1, 2))            'VALIDA QUE EL CODIGO PERTENEZCA
'            If CODIN% > 0 Then                         'A LA LISTA DE PRECIOS
'                PREDOL# = XVALO(Label6)
'                FEXI% = FECHANUM(Text6)
'                Call REPLA(PRE1$, MKI$(CI%), 1, 2)
'                Call REPLA(PRE1$, MKD$(PREDOL#), 62, 8)
'                Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
'                Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
'                SINLIST% = 1
'            End If
'         End If
'    Next U&
    '\\\OT-08-0095-FL-FIN///
    '
    SQLX$ = "SELECT * FROM LISTAPRE with(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE moneda = 2 "
    SQLX$ = SQLX$ + " ORDER BY COD_LISTA"
    '
    Dim LISTtmp1 As New ADODB.Recordset
    Set LISTtmp1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With LISTtmp1
       
       Do While Not .EOF
        If CI% > 0 Then
            NUM% = XVALO(!Cod_Lista)
            PREDOL# = XVALO(Label6)
            FEXI% = FECHANUM(Text7)
            '
            CONDI$ = "COD_PIEZA = " & CI%
            CONDI$ = CONDI$ + " AND COD_LISTA= " & NUM%
            Call ACTUREGISTRO("PREVENTA", "Precio_Unitario", CONDI$, PREDOL#, REG&, "NOMESS")
            Call ACTUREGISTRO("PREVENTA", "Cant_Limite", CONDI$, FEXI%, REG&, "NOMESS")

            If REG& > 0 Then
              SINLIST% = 1
              Call ACTUALIZADATLISTA(NUM%)
            End If
        End If
        .MoveNext
       Loop
    End With
    '
    '\\\OT-08-0136-FL-04/04/08///
    If SINLIST% = 0 Then
         Call SELECHO("LISTAS")
         APREPES1$ = TRIM$(VALACT1$("LISTAS"))
         NUMLIS% = XVALO(VALACT1$("LISTAS")) ' NUMERO DE LISTA DE SELECCIONADA
         Label16 = APREPES1$
         Label16.ToolTipText = VALACT2$("LISTAS")
'        While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
'        APREPES1$ = "LIPRE" + APREPES1$
'        PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
         PREDOL# = XVALO(Label6)
         FEXI% = FECHANUM(Text6)
'        Call REPLA(PRE1$, MKI$(CI%), 1, 2)
'        Call REPLA(PRE1$, MKD$(PREPES#), 62, 8)
'        Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
'        Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
         '
         If CantRegistros("PREVENTA", CONDI$) < 1 Then
100         ABUS$ = RANDSTRING$(-10)
            CONDI$ = "Num_Original = '" & ABUS$ & "'"
            If CantRegistros("PREVENTA", CONDI$) > 0 Then GoTo 100
            Call CreaRegistro("PREVENTA", "Num_Original", ABUS$, "NOMESS")
         End If
         '
         Call ACTUREGISTRO("PREVENTA", "COD_LISTA", CONDI$, APREPES1$, REG&, "NOMESS")
         Numorden% = XVALO(VALOBADA("PREVENTA", "Max(Num_Orden)", "Cod_Lista =" & APREPES1, "NOMESS")) + 1
         
         Call ACTUREGISTRO("PREVENTA", "Precio_Unitario", CONDI$, PREDOL#, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Num_Orden", CONDI$, Numorden%, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Desc_Pieza", CONDI$, Text3, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Denom_Pieza", CONDI$, Text3, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Carac_Control", CONDI$, "", REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Cant_Limite", CONDI$, FEXI%, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Porc_Limite", CONDI$, 0, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Marborra", CONDI$, "", REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Cod_Pieza", CONDI$, CI%, REG&, "NOMESS")
         Call ACTUREGISTRO("PREVENTA", "Num_Original", CONDI$, "", REG&, "NOMESS")
         Call ACTUALIZADATLISTA(NUMLIS%)
         '
    End If
    '\\\OT-08-0136-FL-FIN///
    '
    Call CALCUCOS
    '
End Sub

Sub Label6_DblClick()
    Command9.Enabled = False
    Text1.Enabled = False
    '
    If XVALO(Label11) < 0.05 Then Exit Sub

    'VALIDACION DE CODIGO
    CI% = CODINT%(Text1.TEXT)
    If CI% < 1 Then GoTo 99
    '
    'VALIDACION DE DERECHOS
    If DERACCE%("GRAPREVE", "Grabar Precios de Venta s/Costos") < 2 Then
        GoTo 99
    End If
    '
    HOII% = HOY(HOS$)
    '
    'INGRESO DE NUEVO PRECIO Y VALIDACION
30  NUPRECIO$ = InputBox$("Nuevo Precio", "Registrar Nuevo Precio en Pesos al " + HOS$, Label6)
    If TRIM$(NUPRECIO$) = "" Then GoTo 99
    PREC_NVO# = XVALO(NUPRECIO$)
    If PREC_NVO# < 0.00005 Then GoTo 99
    '
    PREC_ACTUAL# = XVALO(Label6.Caption)
    '
    'ACTUALIZALISTA LISTA
    If Abs(PREC_ACTUAL# - PREC_NVO#) > 0.05 Then 'SI CAMBIO EL PRECIO
        OBSERVAC.Caption = "Observación de Cambio de Precio (Mandataria)"
        OBSERVAC.Text28 = ""
        OBSERVAC.Show 1                        'LLAMA A VENTANA DE OBSERVACIONES
        If OBSERVAC.Label1.Caption <> "OK" Or TRIM$(OBSERVAC.Text28) = "" Then
           Call COMUNI("Debe Ingresar un Texto Motivo del Cambio de Precio\No se va a Registrar el Mismo")
           GoTo 30                             'SI NO SE INGRESA OBSERVACION VUELVE A LA VENTANA AL INPUT BOX
        End If
        '
        'ASIGNACION DE PRECIO EN PANTALLA
         Label6.Caption = FORMATNUM$(PREC_NVO#, "F10.2")
         Label8.Caption = FORMATNUM$(((XVALO(Label6.Caption) * ALIVA) / 100) + XVALO(Label6.Caption), "F10.2")
         Text7 = HOS$

        'SI CAMBIO EL PRECIO Y TIENE OBSERVACION ACTUALIZA LA LISTA Y AGREGA LA NOVEDAD
        Call ACTUALIZALISTA(2, OK%)
        Observacion$ = OBSERVAC.Text28
        '
        If OK% = 1 Then
            'VALIDO SI CAMBIO EL PRECIO EN ESE CASO GRABO EL USUARIO
            CONDI$ = "CODINT = " & CI%
            Call ACTUREGISTRO("MASTER", "UsuPrePes", CONDI$, USERNAME$, REGAF&, "NOMESS")
            Label17 = USERNAME$
            '
            Call ADD_NOVEDAD(CI%, 0, PREC_ACTUAL#, PREC_NVO#, Observacion$, 2, "LISTA")
            '
        End If
   
    Else
        'CASO QUE NO CAMBIO EL PRECIO SE VA.
        GoTo 99
    End If

    '

    'REALIZA CALCULO DE COSTOS
    Call CALCUCOS
99  Command9.Enabled = True
    Text1.Enabled = True

End Sub


Private Sub Label9_DblClick()
     Call Picture2_DblClick
End Sub

Private Sub List1_DblClick()
    '
    '04/03/07 (OT-07-0137)
    If Text1 <> "" Then
        'Call BLANARCH("!BOLREDET")
        Call SETULTREG("!BOLREDET", 0)
        For U& = 1 To List1.ListCount
            List1.ListIndex = U& - 1
            Z$ = List1.TEXT
            COD$ = TRIM$(Mid$(Z$, 1, 16))
            CODI1% = CODINT%(COD$)
            DESC$ = TRIM$(Mid$(Z$, 17, 30))
            UNMED$ = TRIM$(Mid$(Z$, 48, 4))
            CANTI# = XVALO(Mid$(Z$, 53, 10))
            MONE$ = TRIM(Mid$(Z$, 65, 2))
            PUNIT# = XVALO(Mid$(Z$, 67, 9))
            PTOTA# = XVALO(Mid$(Z$, 76, 12))
            '
            Y$ = REGBLAN$("!BOLREDET")
            Call REPLA(Y$, MKI$(CODI1%), 83, 2)
            Call REPLA(Y$, UNMED$, 85, 4)
            Call REPLA(Y$, MKS$(CANTI#), 159, 4)
            '
            Call REGAPP("!BOLREDET", Y$)
        Next U&
        Call CIERRARCH("!BOLREDET")
        '
        VTB% = VENTABU%("DETACOS")
        'If VTB% <= 0 Then Exit Sub
        '
        If VTB% >= 0 Then
            MODIFIC%(0) = 1
            List1.Clear
            For U& = 1 To ULTREG&("!BOLREDET")
                 Y$ = REGLEIDO$("!BOLREDET", U&)
                 CICOMPO% = CVI(Mid$(Y$, 83, 2))
                 CANTI# = 0: PUNIT# = 0: PTOTA# = 0
                 CANTI# = CVS(Mid$(Y$, 159, 4))
                 '
                 X$ = Space(128)
                 CODX$ = CODEXT$(CICOMPO%)
                 Call REPLA(X$, CODX$, 1, 16)
                 DESC$ = DESCRIT0$(CICOMPO%)
                 Call REPLA(X$, DESC$, 17, 32)
                 Call REPLA(X$, UNIMED$(CICOMPO%), 48, 2)
                 Call REPLA(X$, FORMATNUM$(CANTI#, "F10.4"), 53, 10)
                   MONECO% = MONECOSTO%(CICOMPO%)
                 Call REPLA(X$, SIMBOLPES$(MONECO%), 65, 2)
                 Call REPLA(X$, FORMATNUM$(COSUNI(CICOMPO%), "F9.4"), 67, 9)
                 COSTOPAR = COSPESOS(CICOMPO%) * CANTI#
                 Call REPLA(X$, FORMATNUM$(COSTOPAR, "F10.4"), 78, 13) '76,10
                 List1.AddItem X$
             Next U&
             Call GRAUSUARIOYFECHA 'GUARDA EL USUARIO Y LA FECHA DE MODIFICACION DE ESTRUCTURA
         End If
         Call CALCUCOS '09/04/07 (OT-07-0137)
    End If
    '04/03/07 (FIN)
    '
End Sub
Private Sub Listad_Click()
   Call OPLIMAS04.Command2_Click
End Sub



Private Sub mnuasigfech_Click()
   Call COMUNI("Este Proceso Asigna Fecha Segun Selección\ Solo a los que Tengan Precio Unitario y No Registrada una Fecha de Cambio")
   '
   If CLACONTRA% > 0 Then
      
10    OBX$ = OBJPLA$("entrafecha1", VDEF$)
      If OBX$ = "" Then
          Call COMUNI("No se Realizó Ningun Proceso")
          Exit Sub
      End If
      '
      VDEF$ = OBX$
      DES% = CVI(Left$(OBX$, 2))
      If DES% < 1 Then
         Call COMUNI("No se Realizó Ningun Proceso")
         Exit Sub
      End If
      '
      
      SQLX$ = "SELECT * FROM LISTAPRE with(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE moneda = 1"
      SQLX$ = SQLX$ + " ORDER BY COD_LISTA"
      Set LISTtmp = READSET(SQLX$)
      '
      ACTUALICE% = 0
      With LISTtmp
        Do While Not .EOF
            NUMLIS% = XVALO(!Cod_Lista)
            If NUMLIS% > 0 Then
                SQLX$ = "SELECT * FROM PREVENTA with(NOLOCK) "
                SQLX$ = SQLX$ + " WHERE COD_LISTA  = " & NUMLIS%
                SQLX$ = SQLX$ + " ORDER BY COD_LISTA"
                Set RS = READSET(SQLX$)
                With RS
                    Do While Not .EOF
                      FE% = XVALO(!Cant_Limite)
                      If FE% < 1 Then
                         CI1% = XVALO(!Cod_Pieza)
                         If CI1% > 0 Then
                           IMPO# = XVALO(!Precio_Unitario)
                           If IMPO# > 0.005 Then
                                CONDI$ = "Cod_Pieza = " & CI1% & " AND COD_LISTA = " & NUMLIS%
                                Call ACTUREGISTRO("PREVENTA", "Cant_Limite", CONDI$, DES%, REG&, "NOMESS")
                                ACTUALICE% = ACTUALICE% + 1
                           End If
                         End If
                      End If
                      .MoveNext
                    Loop
                End With
                RS.Close
                Set RS = Nothing
            '
            End If
          .MoveNext
        Loop
      End With
      LISTtmp.Close
      Set LISTtmp = Nothing
      '
      Call COMUNI("Proceso Finalizado!\Se Han Actualizado: " & ACTUALICE% & " Registros")

   End If
End Sub

Private Sub Otro_Click()
   Call ACCDIR("ANACOS")
End Sub

Private Sub Picture2_DblClick()
'//OT-07-0478//
    If DERACCE%("VERCOTI", "Ver Cotización del Dolar") >= 2 Then
     If Label9.Visible = False Then
        Label9.Visible = True
        Label11.Visible = True
        Picture2.ToolTipText = ""
        Label9.Visible = True
        Label21.Visible = True
        Text4.Visible = True
      Else
        Label9.Visible = False
        Label11.Visible = False
        Picture2.ToolTipText = "Doble Clic Para Ver cotizacion de Dolar/M.Obra Real"
        Text4.Visible = False
        Label21.Visible = False
     End If
   End If
  '
  
'//OT-07-0478//
End Sub
Private Sub Sal_Click()
   Call Command3_Click
End Sub
Private Sub StoLot_Click()
   Call OPLIMAS04.Command7_Click
End Sub

Private Sub SustCod_Click()
   Call Command23_Click
End Sub

Private Sub Text1_Change()
   '
   '05/03/07 (OT-07-0137)
   CIXII% = CODINT%(Text1.TEXT)
   '
   Call BLANFORMU
   If CIXII% > 0 Then
      Call CARDATITEM
      Call CARFORMULA
      Call CARTIEMOD
      Call CALCUCOS
      MODIFIC%(0) = 0
   End If
   
   '05/03/07 (FIN)
   '
End Sub

Private Sub Text1_GotFocus()
   '
   CIXII% = CODINT%(Text1.TEXT)
   'PUSE EL CODIGO EN UN TIMER POR QUE SI NO ENTRA EN UN CICLO Y SIEMPRE PRESENTA EL MENSAJE.
   Timer1.Enabled = True
   '
End Sub

Private Sub Text1_DblClick()
   '
   Call Command9_Click '05/03/07 (OT-07-0137)
   '
End Sub

Private Sub Text2_Change()
  '
  MODIFIC%(0) = 1
  '
  CI% = CODINT%(Text1)
  If CI% > 0 Then
    Call CALCOSOPE
  End If
  '
End Sub

Private Sub Text2_LostFocus()
   '
   Text2 = FORMATNUM$(XVALO(Text2), "F6.2") '30/03/07 (OT-07-0137)
   '
End Sub

Private Sub Text3_Change()
    MODIFIC%(0) = 1
End Sub

Private Sub Text4_LostFocus()
   '
   MODIFIC%(0) = 1
   Text4 = FORMATNUM$(XVALO(Text4), "F6.2") '30/03/07 (OT-07-0137)
   '
End Sub

Private Sub Text42_Change()
   '
  MODIFIC%(0) = 1
  CI% = CODINT%(Text1)
  If CI% > 0 Then
    Call CALCUCOS
  End If
   '
  
   '
End Sub

Private Sub Text42_LostFocus()
   '
   Text42 = FORMATNUM$(XVALO(Text42), "F5.2") '30/03/07 (OT-07-0137)
   '
End Sub

Private Sub Text45_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text6_Change()
   F$ = " "
   FF = FECHANUM(F$)
   
End Sub

Private Sub Text6_DblClick()
   '
   Call Command5_Click '30/03/07 (OT-07-0137)
   '
End Sub

Private Sub Text6_LostFocus()
   '
   '30/03/07 (OT-07-0137)
    VDEVV$ = TRIM$(Text6.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FF% = FECHANUM(VDEVV$)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          Text6.TEXT = HO$
          'Text6.SetFocus
    End If
   '30/03/07 (FIN)
   '
End Sub

Private Sub Text7_DblClick()
   '
   Call Command8_Click '30/03/07 (OT-07-0137)
   '
End Sub

Private Sub Text7_LostFocus()
   '
   '30/03/07 (OT-07-0137)
    VDEVV$ = TRIM$(Text7.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FF% = FECHANUM(VDEVV$)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          Text7.TEXT = HO$
          'Text7.SetFocus
    End If
   '30/03/07 (FIN)
   '
End Sub
'
Sub GRAFORMULA()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   '
   If DERACCE%("MODISTRU", "Modificación Fórmula de Conjunto") >= 2 Then
      Screen.MousePointer = 11
      CAIT% = List1.ListCount
      '
      For I& = 1 To CAIT%
         CICOMPO% = CODINT%(Left$(List1.List(I& - 1), 16))
         Call EXPLOMUL(CICOMPO%)
         For JJI% = 1 To CONTAMUL%
            If CIK%(JJI%) = CI% Then
               Call MENSERR(24, "Imposible Registrar Esta Fórmula.\Error de Recursividad en Código " + TRIM$(CODEXT$(CI%)) + "\" + DESCRIT0$(CI%) + ".")
               GoTo 99
            End If
         Next JJI%
      Next I&
      '
      CAIT1% = CAIT%
      For JJI% = 1 To 256
         EXPLOLIN$(JJI%) = ""
         If JJI% <= List1.ListCount Then
            XX1$ = List1.List(JJI% - 1)
            ZZ1$ = Space$(128)
            Call REPLA(ZZ1$, Left$(XX1$, 16), 1, 16)
            Call REPLA(ZZ1$, Mid$(XX1$, 17, 28), 17, 28)
            Call REPLA(ZZ1$, Mid$(XX1$, 48, 2), 47, 2)
            Call REPLA(ZZ1$, Mid$(XX1$, 51, 12), 49, 12)
            EXPLOLIN$(JJI%) = ZZ1$
         End If
      Next JJI%
      '
      CODD$ = AJUSTI$(Text1.TEXT, 16)
      CABAS = 1
      CAIT% = List1.ListCount
      Call EXPLOGR(CODD$, CABAS)
      '
   End If
   '
99 Screen.MousePointer = 1
   '
End Sub
'
Sub CARDATITEM()
    '
    CODD$ = Text1.TEXT
    CI% = CODINT%(CODD$)
    If CI% < 1 Then
       If TRIM$(CODD$) <> "" Then
         If CI% < 0 Then
             Call MENSERR(24, "Codigo Anulado o de Baja")
           Else
             Call MENSERR(24, "Código Inexistente o no Válido")
         End If
       End If
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text3.TEXT = DESCRIT0$(CI%)
    Label37.Caption = UNIMED$(CI%)
    Label28.Caption = Str$(CI%)
    Text45 = NUPLANO$(CI%)
    Text42 = FORMATNUM$(PORCOMVE(CI%), "F5.2")
    Text4 = FORMATNUM$(MANO_OBRA_REAL(CI%), "F6.2") ' MANO DE OBRA REAL
    
    'USUARIOS QUE MODIFICARON LISTA DE PRECIOS
    Label18 = USUPREPES(CI%)
    Label17 = USUPREDOL(CI%)
    '
    'USUARIO Y FECHA DE MODIFICACION DE ESTRUCTURA
    CONDI$ = "Codint=" & CODINT(Text1)
    Label19 = VALOBADA("Master", "UsuApFor", CONDI$, "NOMESS")
    Label20 = ""
    FECH$ = FECHATEX$(FEREFOR%(CI%))
    
    If FECH$ <> "01/01/80" Then Label20 = FECH$
    '
    FOCOREP% = 1
    '
    Call PRECIOENLISTA(CI%)
    '
    '
'    PRE1$ = FILTERGETRECORD$(APREPES$, 1, MKI$(CI%))
'    PREPES# = CVD(Mid$(PRE1$, 62, 8))
'    FEX1 = CVI(Mid$(PRE1$, 70, 2))
'    Label1 = FORMATNUM$(PREPES#, "F10.4")
'    Label7 = FORMATNUM$(PREPES# * (1 + ALIVA / 100), "F10.4")
'    Text6 = FECHATEX$(FEX1)
'    '
'
'    '
'    PRE2$ = FILTERGETRECORD$(APREDOL$, 1, MKI$(CI%))
'    PREDOL# = CVD(Mid$(PRE2$, 62, 8))
'    FEX2 = CVI(Mid$(PRE2$, 70, 2))
'    Label6 = FORMATNUM$(PREDOL#, "F10.4")
'    Label8 = FORMATNUM$(PREDOL# * (1 + ALIVA / 100), "F10.4")
'    Text7 = FECHATEX$(FEX2)
    '
99  Screen.MousePointer = 1
    '
End Sub

'
Sub CARFORMULA()
   '
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   PORMILLAR% = 0
   Call LEEEXPLO(CODD$, 1)
   '
   List1.Clear
   Call SETULTREG("!BOLREDET", 0)
   For I% = 1 To CAIT%
     ZZ1$ = EXPLOLIN$(I%)
     CICOMPO% = CODINT%(TRIM$(Left$(ZZ1$, 15)))
     XX1$ = Space$(128)
     Call REPLA(XX1$, Left$(ZZ1$, 16), 1, 16)
     Call REPLA(XX1$, Mid$(ZZ1$, 17, 28), 17, 28)
       UNMED$ = UNIMED$(CICOMPO%)
     Call REPLA(XX1$, UNMED$, 48, 2)
       CANTI# = XVALO(Mid$(ZZ1$, 49, 12))
     Call REPLA(XX1$, FORMATNUM$(CANTI#, "F10.4"), 53, 10)

     'Call REPLA(XX1$, Mid$(ZZ1$, 51, 10), 53, 10)
       MONECO% = MONECOSTO%(CICOMPO%)
     Call REPLA(XX1$, SIMBOLPES$(MONECO%), 65, 2)
     Call REPLA(XX1$, FORMATNUM$(COSUNI(CICOMPO%), "F9.4"), 67, 9)
       COSTOPAR = COSPESOS(CICOMPO%) * XVALO(Mid$(ZZ1$, 53, 10))
     Call REPLA(XX1$, FORMATNUM$(COSTOPAR, "F12.4"), 76, 12) '76,10
     List1.AddItem XX1$
     '
     Y$ = REGBLAN$("!BOLREDET")
     Call REPLA(Y$, MKI$(CICOMPO%), 83, 2)
     Call REPLA(Y$, UNMED$, 85, 4)
     Call REPLA(Y$, MKS$(CANTI#), 159, 4)
     Call REGAPP("!BOLREDET", Y$)
   Next I%
   '
   Call CIERRARCH("!BOLREDET")
   Screen.MousePointer = 1
   '
End Sub
'
Sub BLANFORMU()
   '
   '09/04/07 (OT-07-0137)
   Text2 = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text6.TEXT = ""
   Text7.TEXT = ""
   Text42.TEXT = ""
   Text45.TEXT = ""
   '
   Label1.Caption = ""
   Label2.Caption = ""
   Label4.Caption = ""
   Label5.Caption = ""
   Label6.Caption = ""
   Label7.Caption = ""
   Label8.Caption = ""
   Label93.Caption = ""
   Label103.Caption = ""
   Label104.Caption = ""
   Label107.Caption = ""
   Label110.Caption = ""
   Label28.Caption = ""
   Label37.Caption = ""
   Label68.Caption = ""
   Label12.Caption = ""
   Label3.Caption = ""
   Label13.Caption = ""
   Label14.Caption = ""
   Label15.Caption = ""
   Label16.Caption = ""
   Label17.Caption = ""
   Label18.Caption = ""
   Label19.Caption = ""
   Label20.Caption = ""
   '
   List1.Clear
   '
   PREVENETO# = 0
   '
   MODIFIC%(0) = 0
   
   '09/04/07 (FIN)
   '
End Sub

Sub ALTACODIGO(VDEF$, OKX%)
   '
   Call ABREMASTER
   '
10 OBX$ = OBJPLA$("ALTACODIGO", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   CODINUE$ = TRIM$(Mid$(OBX$, 77, 16))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > 15 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
      GoTo 10
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If
   '
   If TRIM$(Mid$(OBX$, 5, 64)) = "" Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
      GoTo 10
   End If
   '
   'ESTO ESTABA COMENTADO  - SE DESCOMENTO POR QUE SI NO NO GUARDABA EL GRUPO CONTABLE
   GRCONTA% = CVI(Mid$(OBX$, 3, 2))
   If ECOARCH$("GRUCOVEN", MKI(GRCONTA%)) = "" Then
     GRCONTA% = 0
   End If

   
   DENUE$ = TRIM$(Mid$(OBX$, 5, 64))
   '
   TIMA% = Asc(Mid$(OBX$, 93, 1))
   If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
      Call MENSERR(24, "Tipo de Material no Válido")
      GoTo 10
   End If
   '
   UNID$ = Mid$(OBX$, 109, 2)
   If ECOARCH$("UNIMED", UNID$) = "" Then
      Call MENSERR(24, "Unidad de Medida no Válida")
      GoTo 10
   End If
   '
'   If ECOARCH$("GRUCOVEN", MKI(GRCONTA%)) = "" Then
'    GRCONTA% = 0
'   End If


   Screen.MousePointer = 11
   CODII% = NUMAS% + 1
   '
20 On Error Resume Next
   Call ABREMASTER
'   Master.MoveFirst
'   If Err < 1 Then
'     On Error GoTo 0
'     ABUS$ = "Codint = " & CODII%
'     Master.FindFirst "Codint = " & CODII%
'     If Master.NoMatch = False Then
'       CODII% = CODII% + 1
'       GoTo 20
'     End If
'   End If
   '
'   Dim RMaster As ADODB.Recordset
'   Set RMaster = New ADODB.Recordset
'   RMaster.CursorType = adOpenDynamic
'   RMaster.LockType = adLockPessimistic
'   strsql = ("Select * from Master where codint=" & CODII%)
'   RMaster.Open FILTSQL$(strsql), FLEXCONN, , , adCmdText
'   If Not (RMaster.EOF And RMaster.BOF) Then 'si no esta vacio por lo tanto existe
'      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
'      GoTo 10
'   End If
'
'   With RMaster
'      .AddNew
'      On Error Resume Next
'      !CodiEmpr = CodiEmp%
'      On Error GoTo 0
'      !CODIGO = CODINUE$
'      !Descripcion = DENUE$
'      !Status = 1
'      !CODINT = CODII%
'      !UMEDIDA = UNID$
'      !TIMATE = TIMA%
'      !Grucoven = GRCONTA%
'      On Error Resume Next
'      .Update
'      OKX% = 1
'      If Err Then
'         Call MENSERR(24, "Alta no Procesada Correctamente")
'         OKX% = 0
'      End If
'      On Error GoTo 0
''      .Bookmark = .LastModified
'      .Close
'      UPDATMASTER% = 1
'   End With
   
   SQLX$ = "SELECT * FROM MASTER "
   SQLX$ = SQLX$ + "WHERE CODINT = " & CODII% & " "
   Set RMaster = New ADODB.Recordset
   RMaster.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With RMaster
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CodiEmpr = CodiEmp%
         On Error GoTo 0
         !CODIGO = CODINUE$
         !Descripcion = DENUE$
         !Status = 1
         !CODINT = CODII%
         !UMEDIDA = UNID$
         !TIMATE = TIMA%
         !Grucoven = GRCONTA%
         On Error Resume Next
         .Update
         OKX% = 1
         If Err Then
           Call MENSERR(24, "Alta no Procesada Correctamente")
           OKX% = 0
         End If
         On Error GoTo 0
'       .Bookmark = .LastModified
        .Close
        UPDATMASTER% = 1
   End With
   '
99 Screen.MousePointer = 1
Set RMaster = Nothing
End Sub
'
Sub MENSAEDIT()
   '10/04/07 (OT-07-0137)
   Call MENSERR(24, "Operación no Válida.\  \Complete Previamente la Edición de la Fórmula\Haciendo Click sobre 'Append' o 'Insert'.\  \Luego Vuelva a Intentarlo.")
   '10/04/07 (FIN)
End Sub
'
Function VALCODART%()
   '10/04/07 (OT-07-0137)
   VALCODART% = 1
   If CODINT%(Text1.TEXT) < 1 Then
      Call MENSERR(24, "Falta Ingresar Código de Artículo")
      VALCODART% = 0
      Text1.SetFocus
   End If
   '10/04/07 (FIN)
End Function

'
Sub COPYRUTA(CODORI%, CODEST%)
   '
   '10/04/07 (OT-07-0137)
   Dim REGTX$(256)
   Dim RESECOI$(256), DESOPEI$(256)
   '
   CODOR$ = TRIM$(CODEXT$(CODORI%))
   CODES$ = TRIM$(CODEXT$(CODEST%))
   '
   Screen.MousePointer = 11
   Text1 = CODOR$
   YACAR%(3) = 0
   Call CARHORUTA
   '
   Text1 = CODES$
   For KU% = 1 To CARESE%
     RESECO$(KU% - 1) = RESECOI$(KU% - 1)
     DESOPE$(KU% - 1) = DESOPEI$(KU% - 1)
     '
     NOPY% = CVI(Mid$(RESECO$(KU% - 1), 3, 2))
     NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
     DOPE$ = Mid$(RESECO$(KU% - 1), 5, 48)
     SECO$ = NOPEX$ + "     " + DOPE$
     List3.AddItem SECO$
   Next KU%
   '
   YACAR%(3) = 1
   MODIFIC%(3) = 1
   GRABATODO% = 1
   Call GRAHORUTA
   '
   Text1 = ""
   YACAR%(3) = 0
   Text1 = CODES$
   Call CARHORUTA
   '
Exit Sub
   '
   RFFX$ = RECFILT$("PLACONTR", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("PLACONTR", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("PLACONTR")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("PLACONTR", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("PLACONTR")
     End If
     VALANT$ = REGLEIDO$("PLACONTR", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("PLACONTR", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("PLACONTR")
     Call GRAREG("PLACONTR", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("PLACONTR")
   '
   RFFX$ = RECFILT$("PUEPUNTO", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("PUEPUNTO", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("PUEPUNTO")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("PUEPUNTO", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("PUEPUNTO")
     End If
     VALANT$ = REGLEIDO$("PUEPUNTO", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("PUEPUNTO", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("PUEPUNTO")
     Call GRAREG("PUEPUNTO", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("PUEPUNTO")
   '
   Screen.MousePointer = 1
   '10/04/07 (FIN)
   '
End Sub

'
'09/04/07 (OT-07-0137)
Sub CARTAMONED()
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
End Sub
'09/04/07 (FIN)
'
Sub CALCUCOS()
   '
   TOTACUM# = 0
   For U& = 1 To List1.ListCount
            List1.ListIndex = U& - 1
            Z$ = List1.TEXT
            PTOTA# = XVALO(Mid$(Z$, 76, 12))
            TOTACUM# = TOTACUM# + PTOTA#
   Next U&
   '
   'If TOTACUM# > 0 Then ... SE SUPRIME PARA QUE HAGA EL CÁLCULO AUNQUE HAYA SOLO M.O.
      Label5.Caption = FORMATNUM$(TOTACUM#, "F10.4")
      Label68.Caption = FORMATNUM$((XVALO(Label4.Caption) * XVALO(Ltext36)) / 100, "F10.4")
      Label2.Caption = FORMATNUM$(XVALO(Label68.Caption) + XVALO(Label4.Caption) + XVALO(Label5.Caption), "F10.4")
      Label93.Caption = FORMATNUM$((XVALO(Label2.Caption) * XVALO(Ltext41)) / 100, "F10.4")
      Label107.Caption = FORMATNUM$((XVALO(Label2.Caption) * XVALO(Ltext39)) / 100, "F10.4")
      Label104.Caption = FORMATNUM$((XVALO(Label2.Caption) * XVALO(Ltext43)) / 100, "F10.4")
      Label103.Caption = FORMATNUM$(XVALO(Label2.Caption) + XVALO(Label93.Caption) + XVALO(Label107.Caption) + XVALO(Label104.Caption), "F10.4")
      Label110.Caption = FORMATNUM$((XVALO(Label103.Caption) * XVALO(Text42)) / 100, "F10.4")
      '
      Label12.Caption = FORMATNUM$(XVALO(Label103.Caption) + XVALO(Label110.Caption), "F10.4")
      If XVALO(Label11) > 0 Then
          Label3.Caption = FORMATNUM$((XVALO(Label103.Caption) + XVALO(Label110.Caption)) / XVALO(Label11.Caption), "F10.4")
      End If
      '
      On Error Resume Next
         Label14.Caption = TRIM$(FORMATNUM$((((XVALO(Label3.Caption) - XVALO(Label6.Caption)) / XVALO(Label6.Caption)) * 100), "F6.1"))
         Label13.Caption = TRIM$(FORMATNUM$((((XVALO(Label12.Caption) - XVALO(Label1.Caption)) / XVALO(Label1.Caption)) * 100), "F6.1"))

      On Error GoTo 0
       
   'End If
       '
End Sub

Sub CALCOSOPE()
   '
   LOTESTA = 1
   CALOTES = 1
   MINPREP = 0
   MINVARI = XVALO(Text2)
   MINLOTE = MINVARI
   HORAPROD = MINLOTE / 60 / LOTESTA
   '
   CAOPS = 1
   CATOP% = 0
   CUOTOPE = VALHORA(CATOP%)
   '
   CIXI% = CODINT%(Text1)
   NOPYI% = 10
   '
   COSMOD = CUOTOPE * CAOPS * HORAPROD
   Label4.Caption = TRIM$(CSTRING$(MKS$(COSMOD), 3, 12, 4, 2))
   '
   Call CALCUCOS
   '
End Sub

Sub GRATIEMOD()
   
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   '
   Screen.MousePointer = 11
   Call ABRESECOP
   '
'66 With Secoper
      NOPYI% = 10
      CANPRE = 0: CATPRE% = 0
'      .FindFirst "CODICONJ =" & CI% & " and NumeOper = " & NOPYI%
      CONDI$ = "CODICONJ =" & CI% & " and NumeOper = " & NOPYI%
'      If .NoMatch = True Then
'          .AddNew
'         Else
'          .Edit
'      End If
      '
      If CantRegistros("SECOPER", CONDI$) < 1 Then
100      ABUS$ = RANDSTRING$(-10)
         CONDI$ = "DESCOPER = '" & ABUS$ & "'"
         If CantRegistros("SECOPER", CONDI$) > 0 Then GoTo 100
         Call CreaRegistro("SECOPER", "DESCOPER", ABUS$, "NOMESS")
      End If
      MINUST# = XVALO(Text2)
      Call ACTUREGISTRO("SECOPER", "CODICONJ", CONDI$, CI%, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "CODXCONJ", CONDI$, CODEXT$(CI%), REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "NUMEOPER", CONDI$, NOPYI%, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "statoper", CONDI$, 1, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "CodMaq0", CONDI$, "", REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "CodMaq2", CONDI$, "", REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "CodMaq3", CONDI$, "", REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "HorsPrep", CONDI$, 0, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "HorsFijo", CONDI$, 0, REG&, "NOMESS")
      'SE AGREGO FORMATNUM POR QUE CUANDO ERA CON DECIMALES DABA UN ERROR Y NO SE GRABABA
      Call ACTUREGISTRO("SECOPER", "MinuStan", CONDI$, FORMATNUM$(MINUST#, "F16.6"), REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "HorsLimp", CONDI$, 0, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "HorsMHer", CONDI$, 0, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "PieCiclo", CONDI$, 0, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "PorDemor", CONDI$, 0, REG&, "NOMESS")
      If MINUST# > 0.005 Then
        Caden% = 60 / MINUST#
        Call ACTUREGISTRO("SECOPER", "Cadencia", CONDI$, Caden%, REG&, "NOMESS")
      End If
      Call ACTUREGISTRO("SECOPER", "CantiOps", CONDI$, 1, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "CategOps", CONDI$, 0, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "CoefSolp", CONDI$, 0, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "LOTESTAN", CONDI$, 1, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "ProgCNum", CONDI$, "", REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "HoProFab", CONDI$, "", REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "ProTerOp", CONDI$, 0, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "PreUnid", CONDI$, 0, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "MonePrec", CONDI$, 0, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "FechPrec", CONDI$, 0, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "OBSEOPER", CONDI$, "", REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "NUOPER_PREP", CONDI$, 0, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "CATOPER_PREP ", CONDI$, 0, REG&, "NOMESS")
      Call ACTUREGISTRO("SECOPER", "DESCOPER", CONDI$, "Mano de Obra Directa", REG&, "NOMESS")
      
'      !CODICONJ = CI%
'      !CODXCONJ = CODEXT$(CI%)
'      !NUMEOPER = NOPYI%
'      ''!DESCOPER = "Mano de Obra Directa"
'      !statoper = 1
'      !CodMaq0 = ""
'      !CodMaq1 = ""
'      !CodMaq2 = ""
'      !CodMaq3 = ""
'      !HorsPrep = 0
'      !HorsFijo = 0
'      !MinuStan = XVALO(Text2)
'      !HorsLimp = 0
'      !HorsMHer = 0
'      !PieCiclo = 0
'      !PorDemor = 0
'      If !MinuStan > 0.005 Then
'        !Cadencia = 60 / !MinuStan
'      End If
'      !CantiOps = 1
'      !CategOps = 0
'      !CoefSolp = 0
'      !LOTESTAN = 1
'      !ProgCNum = ""
'      !HoProFab = ""
'      !ProTerOp = 0
'      !PreUnid = 0
'      !MonePrec = 0
'      !FechPrec = 0
'      '
'      !OBSEOPER = ""
'      !NUOPER_PREP = 0
'      !CATOPER_PREP = 0
'      .Update
      '
'   End With   ' secoper
   '
   Screen.MousePointer = 1
   '
End Sub

Sub CARTIEMOD()
   
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   '
   Screen.MousePointer = 11
   Call ABRESECOP
   '
'66 With Secoper
      NOPYI% = 10
      CANPRE = 0: CATPRE% = 0
'      .FindFirst "CODICONJ =" & CI% & " and NumeOper = " & NOPYI%
      CONDI$ = "CODICONJ =" & CI% & " and NumeOper = " & NOPYI%
      MINU# = XVALO(VALOBADA("SECOPER", "MINUSTAN", CONDI$, "NOMESS"))
'      If .NoMatch = False Then
'         Text2 = FORMATNUM$(!MinuStan, "F5.2")
          Text2 = FORMATNUM$(MINU#, "F6.2")
'      End If
'   End With   ' secoper
   Screen.MousePointer = 1
   '
End Sub
'
Sub PRIANACOS()
   '
        '10/04/07 (OT-07-0137)
        If VALCODART% < 1 Then
          Text1.SetFocus
          Exit Sub
        End If
        '
        FORIPRE$ = "ANC2-JYM"
        If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
         Call COMUNI("Falta Formulario de Impresión: " & FORIPRE$ & ".FRM")
         Exit Sub
        End If
        '
        If DERACCE%("FORMUSUP/NOMESS", "Impresión de Formulario con Suplemento de Costos Calculados") > 1 Then
          OPX% = COMALTER%("Opciones de Impresión:\\Con Suplemento\Sin Suplemento")
          If OPX% = 1 Then
            FORIPRE$ = "ANC1-JYM"
            If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
              Call COMUNI("Falta Formulario de Impresión: " & FORIPRE$ & ".FRM")
              Exit Sub
            End If
          End If
        End If
        '
        Screen.MousePointer = 11
        FEX = HOY(HOS$)
        FECHDOC$ = HOS$
        NUMEDOC$ = ""
        TIPODOC$ = "Analisis de Costo " + TRIM$(Text1.TEXT)
        '
        CODPARAM$(1) = "COD-MAT": DOCPARAM$(1) = TRIM$(Text1)
        CODPARAM$(2) = "DES-MAT": DOCPARAM$(2) = TRIM$(Text3)
        '
        CODPARAM$(3) = "TIE-PREP": DOCPARAM$(3) = TRIM$(Text2)
        CODPARAM$(4) = "TOT-DEBE": DOCPARAM$(4) = TRIM$(Ltext5)
        CODPARAM$(5) = "TIE-PROD": DOCPARAM$(5) = TRIM$(Ltext36)
        CODPARAM$(6) = "TOT-HABE": DOCPARAM$(6) = Label4.Caption
        '
        CODPARAM$(25) = "IMP-TOTA": DOCPARAM$(25) = Label5.Caption
        CODPARAM$(7) = "CANT-RES": DOCPARAM$(7) = Label5.Caption
        CODPARAM$(8) = "TOT-PREP": DOCPARAM$(8) = Label4.Caption
        CODPARAM$(9) = "TOT-PROD": DOCPARAM$(9) = Label68.Caption
        CODPARAM$(10) = "TOT-TIEM": DOCPARAM$(10) = Label2.Caption
        '
        CODPARAM$(11) = "PIE-COMP": DOCPARAM$(11) = TRIM$(Ltext41)
        CODPARAM$(12) = "DIB-COMP": DOCPARAM$(12) = Label2.Caption
        CODPARAM$(13) = "MAQ-COMP": DOCPARAM$(13) = TRIM$(Ltext39)
        CODPARAM$(14) = "REF-CONJ": DOCPARAM$(14) = Label2.Caption
        CODPARAM$(15) = "ACH-COMP": DOCPARAM$(15) = TRIM$(Ltext43)
        CODPARAM$(16) = "LAR-COMP": DOCPARAM$(16) = Label2.Caption
        '
        CODPARAM$(17) = "CAN-COMP": DOCPARAM$(17) = Label93.Caption
        CODPARAM$(18) = "UNI-COMP": DOCPARAM$(18) = Label107.Caption
        CODPARAM$(19) = "PES-COMP": DOCPARAM$(19) = Label104.Caption
        CODPARAM$(20) = "IMPNEDES": DOCPARAM$(20) = Label103.Caption
        '
        CODPARAM$(21) = "POR-IVA": DOCPARAM$(21) = TRIM$(Text42)
        '
        CODPARAM$(22) = "IMP-DESC": DOCPARAM$(22) = Label103.Caption
        CODPARAM$(23) = "PRE-NETO": DOCPARAM$(23) = Label110.Caption
        '
        CODPARAM$(24) = "TOTA-CAN": DOCPARAM$(24) = Label1.Caption
        '
        'DATOS DE PIE DE  PAGINA. SOLIICITADO POR JORGE
        CODPARAM$(26) = "FECH-VEN": DOCPARAM$(26) = Label20.Caption  'FECHA
        CODPARAM$(27) = "REF-MAT2": DOCPARAM$(27) = Ltext41.Caption  'SCRAP
        CODPARAM$(28) = "REF-MAT5": DOCPARAM$(28) = Ltext39.Caption  'UTILIDAD
        CODPARAM$(29) = "LAR-MATE": DOCPARAM$(29) = Ltext43.Caption  'IMPUESTOS
        CODPARAM$(30) = "REF-MAT3": DOCPARAM$(30) = Label93.Caption  'IMPORTE SCRAP
        CODPARAM$(31) = "REF-MAT6": DOCPARAM$(31) = Label107.Caption 'IMPORTE UTILIDAD
        CODPARAM$(32) = "LAR-CUCO": DOCPARAM$(32) = Label104.Caption 'IMPORTE IMPUESTOS
        CODPARAM$(33) = "REF-MAT4": DOCPARAM$(33) = Label5.Caption   'IMPORTE DE TOTOAL M.PRIMA
        CODPARAM$(34) = "DET-PASE": DOCPARAM$(34) = Ltext5.Caption   'PORCENTAJE DE MANO DE OBRA
        CODPARAM$(35) = "LAR-CUTE": DOCPARAM$(35) = Label4.Caption   'IMPORTE DE MANO DE OBRA
        CODPARAM$(36) = "LAR-MACO": DOCPARAM$(36) = Ltext36.Caption  'GASTOS GENERALES
        CODPARAM$(37) = "REA-PLCO": DOCPARAM$(37) = Label68.Caption  'IMPORTE DE GASTOS GENERALES
        CODPARAM$(38) = "VALO-DEC": DOCPARAM$(38) = Label103.Caption 'COSTO CALCULADO
        CODPARAM$(39) = "ACH-CUTE": DOCPARAM$(39) = Label2.Caption   'SUBTOTAL
        CODPARAM$(40) = "ACH-MATE": DOCPARAM$(40) = Text42           'COMISION DE VENDEDOR
        CODPARAM$(41) = "ACH-CUCO": DOCPARAM$(41) = Label110.Caption 'IMPORTE DE COMISION
        CODPARAM$(42) = "ACH-MACO": DOCPARAM$(42) = Label12.Caption  'PRECIO CALCULADO PESOS
        CODPARAM$(43) = "UBI-MPRI": DOCPARAM$(43) = Label13.Caption  'DIF $
        CODPARAM$(44) = "MEDIDA-1": DOCPARAM$(44) = Text6            'FECHA PRECIO DE VENTA $
        CODPARAM$(45) = "DESTINO": DOCPARAM$(45) = Label18.Caption   'USUARIO $
        CODPARAM$(46) = "CPOS-CLI": DOCPARAM$(46) = Label15.Caption  'LPRE $
        CODPARAM$(47) = "OPE-PUEP": DOCPARAM$(47) = Label1.Caption   'NETO EN $
        CODPARAM$(48) = "PAR-PUEP": DOCPARAM$(48) = Label7.Caption   'FINAL $
        CODPARAM$(49) = "NUME-CHE": DOCPARAM$(49) = Label3.Caption   'PRECIO CALCULADO U$
        CODPARAM$(50) = "PAR-MPRI": DOCPARAM$(50) = Label14.Caption  'DIF. u$
        CODPARAM$(51) = "MEDIDA-2": DOCPARAM$(51) = Text7             'FECHA u$
        CODPARAM$(52) = "DES-CONJ": DOCPARAM$(52) = Label17.Caption  'USUARIO u$
        CODPARAM$(53) = "LOCA-CLI": DOCPARAM$(53) = Label16.Caption  'LPRE u$
        CODPARAM$(54) = "MAQ-PUEP": DOCPARAM$(54) = Label6.Caption   'NETO u$
        CODPARAM$(55) = "HER-PUEP": DOCPARAM$(55) = Label8.Caption   'FINAL u$
        CODPARAM$(56) = "REF-MAT1": DOCPARAM$(56) = Label19.Caption  'USUARIO
        CAPARDOC% = 56
        '
        CODTABU1$(1) = "COD-MPRI"
        CODTABU1$(2) = "DES-MPRI"
        CODTABU1$(3) = "UNI-MPRI"
        CODTABU1$(4) = "CAN-MPRI"
        CODTABU1$(5) = "IMP-NETO"
        CODTABU1$(6) = "SUB-TOTA"
        CODTABU1$(7) = "TIP-MPRI"
        CACOLTAB1% = 7
        '
        For KK1% = 1 To 7
          For KK2% = 1 To List1.ListCount
            TETABU1$(KK1%, KK2%) = ""
          Next KK2%
        Next KK1%
        '
        For U& = 1 To List1.ListCount
            List1.ListIndex = U& - 1
            Z$ = List1.TEXT
            COD$ = TRIM$(Mid$(Z$, 1, 16))
            CODI1% = CODINT%(COD$)
            DESC$ = TRIM$(Mid$(Z$, 17, 30))
            UNMED$ = TRIM$(Mid$(Z$, 48, 4))
            NOMED$ = TRIM$(ECOARCH$("UMEDIDA", UNMED$))
            CANTI# = XVALO(Mid$(Z$, 53, 10))
            MON$ = TRIM$(Mid$(Z$, 65, 2))
            PUNIT# = XVALO(Mid$(Z$, 67, 9))
            PTOTA# = XVALO(Mid$(Z$, 76, 12))
            '
            TETABU1$(1, U&) = COD$
            TETABU1$(2, U&) = DESCRIT0$(CODI1%)
            TETABU1$(3, U&) = NOMED$
            TETABU1$(4, U&) = FORMATNUM$(CANTI#, "F10.4")
            TETABU1$(5, U&) = FORMATNUM$(PUNIT#, "F10.4")
            TETABU1$(6, U&) = FORMATNUM$(PTOTA#, "F10.4")
            TETABU1$(7, U&) = MON$
            
        Next U&
        '
        CAITAB1% = List1.ListCount
        '
        
        Call PRINTDOC("@" + FORIPRE$)
        Screen.MousePointer = 1
        '
        '10/04/07 (OT-07-0137)
     ' End If
      '
End Sub
'
Sub CARHORUTA()
   '
   If YACAR%(3) < 1 Then
      '
      CODD$ = Text1.TEXT
      CI% = CODINT%(CODD$)
      If CI% < 1 Then Exit Sub
      '
      Screen.MousePointer = 11
      Label55 = ""
      Text7 = ""
      Text11 = ""
      Label43 = ""
      Text29 = ""
      Text30 = ""
      Text33 = ""
      Text34 = ""
      Label51 = ""
      '
      ULTOP% = 0
      SQLX$ = "SELECT * FROM Secoper"
      SQLX$ = SQLX$ & " WHERE CodIConj = " & CI%
      SQLX$ = SQLX$ & " AND StatOper >0 "
      SQLX$ = SQLX$ + "ORDER BY NumeOper ASC"
'     Set SECOPTEMP = ARTICULOS.OpenRecordset(SQLX$, dbOpenSnapshot)
      Dim SECOPTEMP As ADODB.Recordset
      Set SECOPTEMP = New ADODB.Recordset
      SECOPTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
      REBLA$ = REGBLAN$("SECOPER") + String$(64, 0)
      IOP% = 0
      On Error Resume Next
      With SECOPTEMP
        .MoveFirst
        If Err Then GoTo 25
        '
        Do Until .EOF = True
          If IsNull(!AltOper) = True Or !AltOper < 1 Then
            NOPY% = !NumeOper
            DOPE$ = !DescOper
            LOTESTA = LOTESTAN(CIXI%)
               If LOTESTA < 1 Then LOTESTA = 1
            CICLOHE = !PieCiclo
               If CICLOHE < 1 Then CICLOHE = 1
            HORPREP = !HorsPrep
            HORFIJO = !HorsFijo
            MINVARI = !MinuStan
            HORHERR = !HorsMHer
            PIECICL = !PieCiclo
            HORLIMP = !HorsLimp
            DEMODIS = !PorDemor
            SOLAPAI% = !CoefSolp
            '
            CAOPS = !CantiOps
            CATOP% = !CategOps
            EQUIPO$ = !CodMaq0
            EQUIP1$ = !CodMaq1
            EQUIP2$ = !CodMaq2
            EQUIP3$ = !CodMaq3
            '
            OBSOPX$ = !OBSEOPER
            '
            PROTERI% = !ProTerOp
            PRECTER# = !PreUnid
            MONETER% = !MonePrec
            FEPRETER% = FECHANUM(Format$(!FechPrec, "dd/mm/yy"))
            '
            X$ = REBLA$
            Call REPLA(X$, MKI$(CI%), 1, 2)
            Call REPLA(X$, MKI$(NOPY%), 3, 2)
            Call REPLA(X$, DOPE$, 5, 48)
            Call REPLA(X$, EQUIPO$, 53, 6)
            Call REPLA(X$, MKS$(HORPREP), 59, 4)
            Call REPLA(X$, MKS$(HORFIJO), 63, 4)
            Call REPLA(X$, MKS$(MINVARI), 67, 4)
            Call REPLA(X$, MKS$(DEMODIS), 75, 4)
            Call REPLA(X$, MKS$(HORLIMP), 79, 4)
            Call REPLA(X$, MKS$(CAOPS), 83, 4)
            Call REPLA(X$, Chr$(CATOP%), 87, 1)
            Call REPLA(X$, MKI$(SOLAPAI%), 89, 2)
            Call REPLA(X$, MKS$(LOTESTA), 91, 4)
            Call REPLA(X$, MKS$(HORHERR), 95, 4)
            Call REPLA(X$, MKS$(PIECICL), 99, 4)
            Call REPLA(X$, EQUIP1$, 103, 6)
            Call REPLA(X$, EQUIP2$, 109, 6)
            Call REPLA(X$, EQUIP3$, 115, 6)
            '
            Call REPLA(X$, MKI$(PROTERI%), 149, 2)
            Call REPLA(X$, MKD$(PRECTER#), 151, 8)
            Call REPLA(X$, MKI$(MONETER%), 159, 2)
            Call REPLA(X$, MKI$(FEPRETER%), 161, 2)
            
            If NOPY% > ULTOP% Then ULTOP% = NOPY%
            NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
            SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
            List3.AddItem SECO$
            RESECO$(List3.ListCount - 1) = X$  ' Left$(X$ + String$(128, 0), 128) + Left$(NUHOPROF$(CI%, NOPY%) + Space$(16), 16) + MKI$(CI%) + MKI$(NOPY%) + String$(28, 0)
            DESOPE$(List3.ListCount - 1) = OBSOPX$
            '
          End If
          .MoveNext
        Loop
      End With
      Set SECOPTEMP = Nothing
      '
'      RNC$ = RECFILT$("COSOPTER", 1, MKI$(CI%))
'      URN& = ULTREG&("COSOPTER")
      '
'      For U& = 1 To Len(RNC$) Step 4
'        RENOTA& = CVS(Mid$(RNC$, U&, 4))
'        If RENOTA& > 0 Then
'          If RENOTA& <= URN& Then
'            TTZZ$ = REGLEIDO$("COSOPTER", RENOTA&)
'            NOPY% = CVI(Mid$(TTZZ$, 3, 2))
'            For KKI% = 1 To List3.ListCount
'              If CVI(Mid$(RESECO$(KKI% - 1), 3, 2)) = NOPY% Then
'                Mid$(RESECO$(KKI% - 1), 145, 32) = TTZZ$
'              End If
'            Next KKI%
'          End If
'        End If
'      Next U&
      '
'      RNC$ = RECFILT$("OBSEOPER", 1, MKI$(CI%))
'      URN& = ULTREG&("OBSEOPER")
'      '
'      For U& = 1 To Len(RNC$) Step 4
'        RENOTA& = CVS(Mid$(RNC$, U&, 4))
'        If RENOTA& > 0 Then
'          If RENOTA& <= URN& Then
'            TTZZ$ = REGLEIDO$("OBSEOPER", RENOTA&)
'            NOPY% = CVI(Mid$(TTZZ$, 3, 2))
'            For KKI% = 1 To List3.ListCount
'              If CVI(Mid$(RESECO$(KKI% - 1), 3, 2)) = NOPY% Then
'                DESOPE$(KKI% - 1) = DESOPE$(KKI% - 1) + Mid$(TTZZ$, 5)
'              End If
'            Next KKI%
'          End If
'        End If
'      Next U&
      '
25    If List3.ListCount > 0 Then
        List3.ListIndex = 0
      End If
      '
      YACAR%(3) = 1
      MODIFIC%(3) = 0
      Screen.MousePointer = 1
      '
   End If
   '
End Sub
'
Sub GRAHORUTA()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   If YACAR%(3) < 1 Then MODIFIC%(3) = 0
   If MODIFIC%(3) < 1 Then Exit Sub
   '
   If DERACCE%("MODIRUTA", "Modificación Hojas de Ruta") < 2 Then
      Exit Sub
   End If
   '
   
'   LO COMENTO POR QUE ESTO SEGURAMENTE VIENE DE AMASTO, PERO EL CARGENITEM NO EXISTE EN ESTE PROYECTO
'   If GRABATODO% <> 1 Then
'     If COMALTER%("Se ha Modificado la Hoja de Ruta del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
'       YACAR%(3) = 0
'       MODIFIC%(3) = 0
''       Call CARGENITEM
'       Exit Sub
'     End If
'   End If
   '
   Screen.MousePointer = 11
   Call ABRESECOP
   'sql
   Dim RSECOPER As ADODB.Recordset
   Set RSECOPER = New ADODB.Recordset
   RSECOPER.CursorType = adOpenDynamic
   RSECOPER.LockType = adLockPessimistic
   strsql = ("Select * from Secoper where CodIConj=") & CI%
   RSECOPER.Open FILTSQL$(strsql), FLEXCONN, , , adCmdText
   With RSECOPER
      strsql_U = ("Update Secoper set StatOper = -1 where StatOper>0 and CodIConj=") & CI%
      FLEXCONN.Execute strsql_U, cdadreg
        
      
    For LL% = 1 To List2.ListCount
    NOPYI% = CVI(Mid$(RESECO$(LL% - 1), 3, 2))
    CANPRE = 0: CATPRE% = 0
     '.FindFirst "CODICONJ =" & CI% & " and NumeOper = " & NOPYI%
     'If .NoMatch = False Then
     Dim MIRS As ADODB.Recordset
     strsql_S = ("Select * from Secoper where CODICONJ=") & CI% & " and NumeOper=" & NOPYI%
     Set MIRS = FLEXCONN.Execute(FILTSQL$(strsql_S))
     If Not (MIRS.EOF And MIRS.BOF) Then
       CANPRE = MIRS!NUOPER_PREP
       CATPRE% = MIRS!CATOPER_PREP
     End If
     'End If
     '
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
     !CODICONJ = CI%
     !CODXCONJ = CODEXT$(CI%)
     !NumeOper = CVI(Mid$(RESECO$(LL% - 1), 3, 2))
     !DescOper = TRIM$(Mid$(RESECO$(LL% - 1), 5, 48))
     !statoper = 1
     !CodMaq0 = TRIM$(Mid$(RESECO$(LL% - 1), 53, 6))
     !CodMaq1 = TRIM$(Mid$(RESECO$(LL% - 1), 103, 6))
     !CodMaq2 = TRIM$(Mid$(RESECO$(LL% - 1), 109, 6))
     !CodMaq3 = TRIM$(Mid$(RESECO$(LL% - 1), 115, 6))
     !HorsPrep = CVS(Mid$(RESECO$(LL% - 1), 59, 4))
     !HorsFijo = CVS(Mid$(RESECO$(LL% - 1), 63, 4))
     !MinuStan = CVS(Mid$(RESECO$(LL% - 1), 67, 4))
     !HorsLimp = CVS(Mid$(RESECO$(LL% - 1), 79, 4))
     !HorsMHer = CVS(Mid$(RESECO$(LL% - 1), 95, 4))
     !PieCiclo = CVS(Mid$(RESECO$(LL% - 1), 99, 4))
     !PorDemor = CVS(Mid$(RESECO$(LL% - 1), 75, 4))
     If !MinuStan > 0.005 Then
       !Cadencia = 60 / !MinuStan
     End If
     !CantiOps = CVS(Mid$(RESECO$(LL% - 1), 83, 4))
     !CategOps = Asc(Mid$(RESECO$(LL% - 1), 87, 1))
     !CoefSolp = CVI(Mid$(RESECO$(LL% - 1), 89, 2))
     !LOTESTAN = CVS(Mid$(RESECO$(LL% - 1), 91, 4))
     !ProgCNum = TRIM$(Mid$(RESECO$(LL% - 1), 121, 8))
     !HoProFab = TRIM$(Mid$(RESECO$(LL% - 1), 129, 16))
     !ProTerOp = CVI(Mid$(RESECO$(LL% - 1), 149, 2))
     !PreUnid = CVD(Mid$(RESECO$(LL% - 1), 151, 8))
     !MonePrec = CVI(Mid$(RESECO$(LL% - 1), 159, 2))
     !FechPrec = CVDAT(CVI(Mid$(RESECO$(LL% - 1), 161, 2)))
     '
     !OBSEOPER = DESOPE(LL% - 1)
     !NUOPER_PREP = CANPRE
     !CATOPER_PREP = CATPRE%
     .Update
   
     Next LL%
     RSECOPER.Close
     Set RSECOPER = Nothing
   End With
      strsql_D = ("Delete Secoper where StatOper=-1 and CodIConj=") & CI%
      FLEXCONN.Execute strsql_D, cdadreg
      
'   With Secoper
'     On Error Resume Next
'     .MoveFirst
'     If Err Then
'       On Error GoTo 0
'       GoTo 66
'     End If
'     YAPRIMO% = 0
'     ABUS$ = "CodIConj =" & CI% & " AND StatOper > 0"
'     On Error Resume Next
'62   If YAPRIMO% = 0 Then
'         .FindFirst ABUS$
'         YAPRIMO% = 1
'       Else
'621      .FindNext ABUS$
'     End If
'     If .NoMatch = False Then
'       If IsNull(!AltOper) = True Or !AltOper < 1 Then
'         .Edit
'         !statoper = (-1)
'         .Update
'         GoTo 62
'       End If
'       GoTo 621
'     End If
'   End With
'   '
'66 With Secoper
'   For LL% = 1 To List3.ListCount
'    NOPYI% = CVI(Mid$(RESECO$(LL% - 1), 3, 2))
'    CANPRE = 0: CATPRE% = 0
'     .FindFirst "CODICONJ =" & CI% & " and NumeOper = " & NOPYI%
'     If .NoMatch = False Then
'       CANPRE = !NUOPER_PREP
'       CATPRE% = !CATOPER_PREP
'     End If
'     '
'     .AddNew
'     !CODICONJ = CI%
'     !CODXCONJ = CODEXT$(CI%)
'     !NUMEOPER = CVI(Mid$(RESECO$(LL% - 1), 3, 2))
'     !DESCOPER = TRIM$(Mid$(RESECO$(LL% - 1), 5, 48))
'     !statoper = 1
'     !CodMaq0 = TRIM$(Mid$(RESECO$(LL% - 1), 53, 6))
'     !CodMaq1 = TRIM$(Mid$(RESECO$(LL% - 1), 103, 6))
'     !CodMaq2 = TRIM$(Mid$(RESECO$(LL% - 1), 109, 6))
'     !CodMaq3 = TRIM$(Mid$(RESECO$(LL% - 1), 115, 6))
'     !HorsPrep = CVS(Mid$(RESECO$(LL% - 1), 59, 4))
'     !HorsFijo = CVS(Mid$(RESECO$(LL% - 1), 63, 4))
'     !MinuStan = CVS(Mid$(RESECO$(LL% - 1), 67, 4))
'     !HorsLimp = CVS(Mid$(RESECO$(LL% - 1), 79, 4))
'     !HorsMHer = CVS(Mid$(RESECO$(LL% - 1), 95, 4))
'     !PieCiclo = CVS(Mid$(RESECO$(LL% - 1), 99, 4))
'     !PorDemor = CVS(Mid$(RESECO$(LL% - 1), 75, 4))
'     If !MinuStan > 0.005 Then
'       !Cadencia = 60 / !MinuStan
'     End If
'     !CantiOps = CVS(Mid$(RESECO$(LL% - 1), 83, 4))
'     !CategOps = Asc(Mid$(RESECO$(LL% - 1), 87, 1))
'     !CoefSolp = CVI(Mid$(RESECO$(LL% - 1), 89, 2))
'     !LOTESTAN = CVS(Mid$(RESECO$(LL% - 1), 91, 4))
'     !ProgCNum = TRIM$(Mid$(RESECO$(LL% - 1), 121, 8))
'     !HoProFab = TRIM$(Mid$(RESECO$(LL% - 1), 129, 16))
'     !ProTerOp = CVI(Mid$(RESECO$(LL% - 1), 149, 2))
'     !PreUnid = CVD(Mid$(RESECO$(LL% - 1), 151, 8))
'     !MonePrec = CVI(Mid$(RESECO$(LL% - 1), 159, 2))
'     !FechPrec = CVDAT(CVI(Mid$(RESECO$(LL% - 1), 161, 2)))
'     '
'     !OBSEOPER = DESOPE(LL% - 1)
'     !NUOPER_PREP = CANPRE
'     !CATOPER_PREP = CATPRE%
'     .Update
'     '
'   Next LL%
'   '
'63 .FindFirst "CODICONJ = " & CI%
'   If .NoMatch = False Then
'     If !statoper = (-1) Then
'       .Delete
'       GoTo 63
'     End If
'   End If
'   '
'64 .FindNext "CODICONJ = " & CI%
'   If .NoMatch = False Then
'     If !statoper = (-1) Then
'       .Delete
'       GoTo 63
'     End If
'     GoTo 64
'   End If
'   '
'   .Close
'   End With   ' secoper
'   '
   MODIFIC(3) = 0
   Screen.MousePointer = 1
   '
End Sub


'Sub ACTULISCOSPRO()
'
'PASADA A CLIEFUN
'10  Call SELECHO("PROVED1")
'    NCX1% = XVALO(VALACT1$("PROVED1"))
'    If NCX1% > 0 Then
'       '
'       Call APERBASDAT("MASTER")
'       '
''       Set MasterBaja = ARTICULOS.OpenRecordset("SELECT * FROM Master WHERE PROVHABI = " & NCX1% & " AND Status > 0 ORDER BY CODIGO ASC;", dbOpenSnapshot)
''       With MasterBaja
''       On Error Resume Next
''       MasterBaja.MoveFirst
''       If Err Then
''          On Error GoTo 0
''          Call MENSERR(24, "No hay Registros que Coincidan con el Proveedor Elegido .")
''          GoTo 20
''       End If
''       On Error GoTo 0
''       '
''20         Call BLANARCH("!ACTUCOS")
''       '
''       I& = 0
''       Do While Not .EOF
''          COD$ = !CODIGO
''          MON% = !Monecos
''          PODES1 = ROUND#(!PorDescCos, 1): If PODES1 > 99 Then PODES1 = 0
''          COS1# = !COSUNI / (1 - PODES1 / 100)
''          COS2# = !COSUNI
''          FECH = !FECOSTO
''          '
''          Y$ = REGBLAN$("ACTUCOS")
''          CODI% = CODINT%(COD$)
''          Call REPLA(Y$, MKI$(CODI%), 1, 2)
''          Call REPLA(Y$, MKI$(MON%), 3, 2)
''          Call REPLA(Y$, MKD$(COS1#), 5, 8)
''          Call REPLA(Y$, MKS$(PODES1), 13, 4)
''          Call REPLA(Y$, MKD$(COS2#), 17, 8)
''          FCOS% = CVINT%(FECH)
''          Call REPLA(Y$, MKI$(FCOS%), 25, 2)
''          I& = I& + 1
''          Call GRAREG("!ACTUCOS", Y$, I&)
''       .MoveNext
''       '
''       Loop
'       '
'       'Lectura desde SQL
'       SQLX$ = "SELECT * FROM MASTER with(NOLOCK) "
'       SQLX$ = SQLX$ + " WHERE PROVHABI = " & NCX1% & " AND Status > 0"
'       SQLX$ = SQLX$ + " ORDER BY CODIGO ASC"
'       '
'       Dim MasterBaja As New ADODB.Recordset
'       Set MasterBaja = FLEXCONN.Execute(FILTSQL$(SQLX$))
'       '
'       With MasterBaja
'          If Err Then
'            On Error GoTo 0
'            Call MENSERR(24, "No hay Registros que Coincidan con el Proveedor Elegido .")
'          End If
'          '
'          I& = 0
'          Do While Not .EOF
'             COD$ = !CODIGO
'             MON% = !Monecos
'             PODES1 = ROUND#(!PorDescCos, 1): If PODES1 > 99 Then PODES1 = 0
'             COS1# = !COSUNI / (1 - PODES1 / 100)
'             COS2# = !COSUNI
'             FECH = !FECOSTO
'             '
'             Y$ = REGBLAN$("ACTUCOS")
'             CODI% = CODINT%(COD$)
'             Call REPLA(Y$, MKI$(CODI%), 1, 2)
'             Call REPLA(Y$, MKI$(MON%), 3, 2)
'             Call REPLA(Y$, MKD$(COS1#), 5, 8)
'             Call REPLA(Y$, MKS$(PODES1), 13, 4)
'             Call REPLA(Y$, MKD$(COS2#), 17, 8)
'             FCOS% = CVINT%(FECH)
'             Call REPLA(Y$, MKI$(FCOS%), 25, 2)
'             I& = I& + 1
'             Call GRAREG("!ACTUCOS", Y$, I&)
'            .MoveNext
'            '
'          Loop
'          Call SETULTREG("!ACTUCOS", I&)
'       End With
'       '
'
'40     JUN% = VENTABU%("ACTUCOS/TITUPAN=Costo de Reposición - " + RASOCLI$((-1) * NCX1%))
'       '
'       If JUN% < 0 Then GoTo 10
'
'       If JUN% >= 0 Then
'       OBSERVAC.Caption = "Observación General de Cambio de Precio x Proveedor (Mandataria)"
'       OBSERVAC.Text28 = ""
'       OBSERVAC.Show 1                        'LLAMA A VENTANA DE OBSERVACIONES
'       If OBSERVAC.Label1.Caption <> "OK" Or TRIM$(OBSERVAC.Text28) = "" Then
'         Call COMUNI("Debe Ingresar un Texto General para el Cambio de Precio\No se van a Registrar el Mismo")
'         GoTo 40                             'SI NO SE INGRESA OBSERVACION VUELVE A LA VENTANA AL INPUT BOX
'       End If
'       Observacion$ = OBSERVAC.Text28
'
'
'       HOII% = HOY(HOS$)
'       '
'       Dim COSTOANTERIOR#(32767)
'       For I& = 1 To 32767
'         COSTOANTERIOR#(I&) = 0
'       Next I&
'
'        Call ABREMASTER
'        For j& = 1 To ULTREG&("!ACTUCOS")
'            Z$ = REGLEIDO$("!ACTUCOS", j&)
'            COD3% = CVI(Mid$(Z$, 1, 2))
'            PODES1 = CVS(Mid$(Z$, 13, 4)): If PODES1 < 0 Or PODES1 > 99 Then PODES1 = 0
'            COU = (CVD(Mid$(Z$, 5, 8))) * (1 - PODES1 / 100)
'            MONEI% = CVI(Mid$(Z$, 3, 2))
'            FCOS% = CVI(Mid$(Z$, 25, 2))
'            PHAB% = PROVHABI%(COD3%)
'            '
'            If PHAB% > 0 And NCX1% <> PHAB% Then
'               TTXX$ = "Para el Artículo " + TRIM$(CODEXT$(COD3%)) + "\el Proveedor Habitual Registrado es\"
'               TTXX$ = TTXX$ + TRIM$(RASOCLI$((-1) * PHAB%)) + "\  \"
'               TTXX$ = TTXX$ + "Desea Cambiar el nuevo Proveedor Habitual a\"
'               TTXX$ = TTXX$ + RASOCLI$((-1) * NCX1%)
'               If COMALTER%(TTXX$ + "\\No, Conservar Anterior\Si, Cambiar") = 1 Then GoTo 10
'            End If
'            '
''            Master.FindFirst ("CODINT =" & COD3%)
''            If Master.NoMatch = False Then
''               Master.Edit
''               Master!PROVHABI = NCX1%
''               Master!COSUNI = COU
''               Master!PorDescCos = PODES1
''               Master!FECOSTO = CVDAT(FECOS%)
''               Master!Monecos = MONEI%
''               Master.Update
''           End If
'            ' Grabacion en Tabla SQL
'            SQLX$ = " SELECT * FROM Master"
'            SQLX$ = SQLX$ + " WHERE CODINT = " & COD3% & ""
'            Dim MODMASTER As ADODB.Recordset
'            Set MODMASTER = New ADODB.Recordset
'            MODMASTER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
'            On Error Resume Next
'                With MODMASTER
'                    .MoveFirst
'                    If Err < 1 Then
'                        On Error GoTo 0
'                        !PROVHABI = NCX1%
'                        !FECOSTO = CVDAT(FCOS%)
'                        'EN EL CASO QUE CAMBIE EL COSTO  PONE LA FECHA DE HOY
'                        COSTOANTERIOR#(COD3%) = XVALO(!COSUNI)
'                        If ROUND(XVALO(!COSUNI), 4) <> ROUND(COU, 4) Then
'                           !FECOSTO = CVDAT(HOII%)
'                        End If
'                        !COSUNI = COU
'                        !PorDescCos = PODES1
'                        !Monecos = MONEI%
'                       .Update
'                    End If
'                 End With
'
'        Next j&
'        '
'        'RECORRO NUEVAMENTE PARA ACTUALIZAR LAS NOVEDADES
'        For j& = 1 To ULTREG&("!ACTUCOS")
'            Z$ = REGLEIDO$("!ACTUCOS", j&)
'            COD3% = CVI(Mid$(Z$, 1, 2))
'            PODES1 = CVS(Mid$(Z$, 13, 4)): If PODES1 < 0 Or PODES1 > 99 Then PODES1 = 0
'            COU = (CVD(Mid$(Z$, 5, 8))) * (1 - PODES1 / 100)
'            MONEI% = CVI(Mid$(Z$, 3, 2))
'            PHAB% = NCX1%
'
'            If ROUND(COSTOANTERIOR#(COD3%), 4) <> ROUND(COU, 4) Then
'               PREC_NVO# = XVALO(COU)
'               Call ADD_NOVEDAD(COD3%, NCX1%, COSTOANTERIOR#(COD3%), PREC_NVO#, Observacion$, MONEI%, "COSTO")
'            End If
'        Next j&
'
'        Call CARFORMULA
'        '
'       End If
'    End If
'
'End Sub
'
Sub ARCHMAES()
'    Call CONECRUN("ARCHIG04", "Archivos Maestros")
    'TRANSPORTA EL CODIGO ACTIVO AL MAESTRO
    'SE REEMPLAZA "/" PR "|" POR QUE CUANDO LO PASA DE ANACOS09 LO REALILZA A LA INVERSA
    'POR QUE SI SE PASA ORIGINALMENTE CON EL "/" LO TRUNCA HASTA AHI.
    CODX$ = REPLACAR(TRIM$(Text1), "/", "|")
    Call CONECRUN("AMASTO07/" + "ANACOS09" + CODX$, "Archivos Maestros")
End Sub

Sub GRUPCONT()
    '
    Call BLANARCH("!GRUCOVEN")
    For I& = 1 To ULTREG&("GRUCOVEN")
       X$ = REGLEIDO$("GRUCOVEN", I&)
       If CVI(Left$(X$, 2)) > 0 Then
           Call REGAPP("!GRUCOVEN", X$)
       End If
    Next I&
    Call CIERRARCH("GRUCOVEN")
    '
    VTB% = VENTABU%("GRUCOVEN/NC/TITUPAN=Grupos Contables")
    If VTB% < 0 Then Exit Sub
    j& = 0
    For I& = 1 To ULTREG&("!GRUCOVEN")
       X$ = REGLEIDO$("!GRUCOVEN", I&)
       If CVI(Left$(X$, 2)) > 0 Then
           j& = j& + 1
           Call GRAREG("GRUCOVEN", X$, j&)
       End If
    Next I&
    Call SETULTREG("GRUCOVEN", j&)
    Call CIERRARCH("GRUCOVEN")

End Sub


Sub LISTAPRE()
'    Call CONECRUN("LISPRE04", "Listas de Precios de Venta")
    Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
End Sub

Sub SEGUITRE()
10  OPX% = COMALTER%("Opciones de Procesamiento:\\Recálculo Mensual\Listado por Fechas")
    If OPX% = 1 Then
      If DERACCE%("RECAMEN", "Re-Calculo Mensual de Costos") <> 0 Then
          Call RECAMEN
      End If
        GoTo 10
    ElseIf OPX% = 2 Then
        If DERACCE%("LIPORFE", "Listados Resumen de Costos por Fecha") <> 0 Then
            Call LIPORFE
        End If
        GoTo 10
    Else
        Exit Sub
    End If
End Sub

Sub RECAMEN()

    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    PERIOSEL$ = MESANO$(MESASEL$, DES%, HAS%)
'    DESDAT = Int(CDbl(CVDAT(DES%)))
'    HASDAT = Int(CDbl(CVDAT(HAS%)))
    DESDAT = FETEXCOR1$(DES%)
    HASDAT = FETEXCOR1$(HAS%)
    If PERIOSEL$ = "" Then Exit Sub
    '
    XXANT$ = FILTERGETRECORD$("SEGUITRE", 1, PERIOSEL$)
    If UCase$(TRIM$(Left$(XXANT$, 16))) = UCase$(TRIM$(PERIOSEL$)) Then
       fecalcx = CVI(Mid$(XXANT$, 127, 2))
       If COMALTER%("Ya Existe Información Para el Período Elegido\de " + TRIM$(PERIOSEL$) + ", Calculada el " + FECHATEX$(fecalcx) + ".\Opciones de Trabajo:\\Mantenerla\Sobre-Escribirla") <> 2 Then
          Exit Sub
       End If
    End If
    '
    Dim TOTVENTA#(32767)
    Dim TOTCANT#(32767)
    Dim TOTCOSTO#(32767)
    '
    j& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
      TOTCANT#(U&) = 0
      TOTCOSTO#(U&) = 0
    Next U&

    Call APERBASDAT("CABAN")
'    SQLX$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "WHERE INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
'    SQLX$ = SQLX$ + "AND TipoMovim = 'VENTA' "
'
'    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'    '
'    With CABATEMP
'      On Error Resume Next
'      .MoveLast
'      If Err > 0 Then
'        On Error GoTo 0
'        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
'        GoTo 99
'      End If
'      On Error GoTo 0
'      FIN& = .RecordCount
'      .MoveFirst
    '
    'Lectura desde Tabla SQL
    CONDI$ = " FechConta >= '" & DESDAT & "'"
    CONDI$ = CONDI$ + "AND FechConta <= '" & HASDAT & "' "
    CONDI$ = CONDI$ + "AND TipoMovim = 'VENTA' "
    FIN& = CantRegistros&("CAJABANC", CONDI$, "NOMESS")
    If FIN& < 1 Then
       Call MENSERR(24, "Sin Movimientos en el Período Elegido")
       GoTo 99
    End If
    '
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE " + CONDI$
    Dim CABATEMP As New ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With CABATEMP
      Screen.MousePointer = 11
      ARTMAX% = 0
      KU& = 0
      '
      Do While Not .EOF
        KU& = KU& + 1
        Call TRACE(20, KU&, FIN&)
        TOTFAC# = ROUND(!IHABECOMP - !IDEBECOMP, 2)
        NCLIE% = !NUCLIEN
        If Abs(TOTFAC#) >= 0.005 Then
          NROFAC$ = !codicom_lar
          SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK) "
          SQLX$ = SQLX$ + "WHERE DoSecLar = " & "'" & NROFAC$ & "'"
          SQLX$ = SQLX$ + "AND (Nume_Clie = " & NCLIE% & " "
             ORGD% = ORIGRUDI%(NCLIE%)
             If ORGD% > 0 Then
               If ORGD% <> NCLIE% Then
                  SQLX$ = SQLX$ + " OR Nume_Clie = " & ORGD% & ""
               End If
             End If
          SQLX$ = SQLX$ + ") "
          SQLX$ = SQLX$ + "Order by NuorItem ASC "
          '
          'Set STOTEMP1 = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
          Dim STOTEMP1 As New ADODB.Recordset
          Set STOTEMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
          '
          With STOTEMP1
            On Error Resume Next
            .MoveFirst
            If Err > 0 Then
               TOTVENTA#(0) = TOTVENTA#(0) + TOTFAC#
               On Error GoTo 0
               GoTo 30
            End If
            On Error GoTo 0 '
            '
            'COEFICIENTE DE PRECIOS
            SUMATVEN# = 0
            Do While Not .EOF
              CIXI% = !COD_INTE
              CANTU = (!CANTSALID - !CANTINGRE)
              PREUN# = !VALOUNIT
              If !MoneVal > 1 Then
                PREUN# = PREUN# * !TIPOCAM
              End If
              SUMATVEN# = SUMATVEN# + PREUN# * CANTU
            .MoveNext
            Loop
            COEFIMA# = 1
            If Abs(SUMATVEN#) > 0.005 Then
              COEFIMA# = Abs(TOTFAC# / SUMATVEN#)
            End If
            '
            .MoveFirst
            Do While Not .EOF
              nart% = !COD_INTE: If nart% < 0 Then nart% = 0
              If nart% > ARTMAX% Then ARTMAX% = nart%
              '\\\OT-06-0494-RG-20/11/06///
              CANTI# = (!CANTSALID - !CANTINGRE)
              TOTCANT#(nart%) = TOTCANT#(nart%) + CANTI#
              '\\\OT-06-0494-RG-FIN///
              NETO# = (!CANTSALID - !CANTINGRE) * !VALOUNIT * COEFIMA#
              COSTONETO# = ROUND((!CANTSALID - !CANTINGRE) * !CostoRep, 2)
              If !MoneVal > 1 Then
                NETO# = NETO# * !TIPOCAM
              End If
              NETO# = ROUND(NETO#, 2)
              TOTVENTA#(nart%) = TOTVENTA#(nart%) + NETO#
              TOTCOSTO#(nart%) = TOTCOSTO#(nart%) + COSTONETO#
              TOTFAC# = TOTFAC# - NETO#
              .MoveNext
            Loop
            '
            TOTVENTA#(0) = TOTVENTA#(0) + TOTFAC#
          End With                    ' SOBRE MOVISTO
          STOTEMP1.Close
          Set STOTEMP1 = Nothing
          '
        End If
        '
30    .MoveNext
      Loop
      '
    End With                          ' SOBLE CAJABANC
    '
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    For I% = 0 To ARTMAX%
      TOTAGEN# = TOTAGEN# + TOTVENTA#(I%)
    Next I%
    '
    List2.Clear
    For NCI% = 0 To 32766
       If Abs(TOTVENTA#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTVENTA#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCI%)), 4, 12, 2, 2)
          List2.AddItem RELIB$
       End If
    Next NCI%
    '
    j& = 0: PORACU = 0
    FIN& = List2.ListCount
    '
    For U& = 1 To FIN&
        List2.ListIndex = U& - 1
        NC% = XVALO(Mid$(List2.TEXT, 14, 6))
        TONEVEN# = XVALO(Mid$(List2.TEXT, 28, 12))
        PORPAR = XVALO(Mid$(List2.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKPIE")
        W$ = REGBLAN$("RANMABRU")
        
        j& = j& + 1: JI% = j&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call REPLA(Z$, MKD$(TOTCANT#(NC%)), 21, 8) '10/05/07 (OT-07-0188)
        Call GRAREG("RANKPIE", Z$, j&)
        '\\\OT-06-0494-RG-20/11/06///
        Call REPLA(W$, MKI$(JI%), 1, 2)
        Call REPLA(W$, MKI$(NC%), 3, 2)
        Call REPLA(W$, MKD$(TOTCANT#(NC%)), 5, 8)
        Call REPLA(W$, MKD$(TONEVEN#), 13, 8)
        '//OT-07-0458//
        'CU# = COSPESOS(NC%)
        CU# = 0
        If TOTCANT#(NC%) > 0.005 Then CU# = TOTCOSTO#(NC%) / TOTCANT#(NC%)
        Call REPLA(W$, MKD$(CU#), 21, 8)
        CT# = TOTCOSTO#(NC%)
        '//OT-07-0458//
        Call REPLA(W$, MKD$(CT#), 29, 8)
        UB# = TONEVEN# - CT#
        Call REPLA(W$, MKD$(UB#), 37, 8)
        Call REPLA(W$, MKS$(PORPAR), 45, 4)
        Call REPLA(W$, MKS$(PORACU), 49, 4)
        '\\\OT-07-0014-RG-24/1/07///
        IMPIVA# = (TONEVEN# * PORCEIVA(NC%)) / 100
        Call REPLA(W$, MKD$(IMPIVA#), 53, 8)
        ITOVE# = TONEVEN# + IMPIVA#
        Call REPLA(W$, MKD$(ITOVE#), 61, 8)
        '\\\OT-07-0014-RG-FIN///
        Call GRAREG("RANMABRU", W$, j&)
        '\\\OT-06-0494-RG-FIN///
    Next U&
    
    Call SETULTREG("RANKPIE", j&)
    Call CIERRARCH("RANKPIE")

    Call SETULTREG("RANMABRU", j&)
    Call CIERRARCH("RANMABRU")
 
    List2.Clear
  
    FIN& = ULTREG&("RANKPIE")
    For JX& = 1 To FIN&
        Call TRACE(20, JX&, FIN&)
        XX$ = REGLEIDO$("RANKPIE", JX&)
        CODIG% = CVI(Mid$(XX$, 3, 2))
        CANITEM# = CVD(Mid$(XX$, 21, 8))
        Text1.TEXT = CODEXT$(CODIG%)
        Openforms = DoEvents
        MINMANOBRA# = MINMANOBRA# + XVALO(Text2.TEXT) * CANITEM#
        MINMANOBRA_REAL# = MINMANOBRA_REAL# + XVALO(Text4.TEXT) * CANITEM#
        PESMANOBRA# = PESMANOBRA# + XVALO(Label4.Caption) * CANITEM#
        TOTMATPRIM# = TOTMATPRIM# + XVALO(Label5.Caption) * CANITEM#
        GASTOSGRAL# = GASTOSGRAL# + XVALO(Label68.Caption) * CANITEM#
        SCRAPMOBRA# = SCRAPMOBRA# + XVALO(Label93.Caption) * CANITEM#
        UTILIDAD# = UTILIDAD# + XVALO(Label107.Caption) * CANITEM#
        IMPUESTOS# = IMPUESTOS# + XVALO(Label104.Caption) * CANITEM#
        COMIVENDED# = COMIVENDED# + XVALO(Label110.Caption) * CANITEM#
        COSTOTOTAL# = COSTOTOTAL# + XVALO(Label12.Caption) * CANITEM#
        IMPNET# = IMPNET# + CVD(Mid$(XX$, 5, 8))
    Next JX&
    
    Text1.TEXT = ""
    TXT$ = REGBLAN$("SEGUITRE")
    Call REPLA(TXT$, PERIOSEL$, 1, 16)
    Call REPLA(TXT$, MKI(DES%), 17, 2)
    Call REPLA(TXT$, MKS(MINMANOBRA#), 19, 4)
    Call REPLA(TXT$, MKD(PESMANOBRA#), 23, 8)
    Call REPLA(TXT$, MKD(TOTMATPRIM#), 31, 8)
    Call REPLA(TXT$, MKD(GASTOSGRAL#), 39, 8)
    Call REPLA(TXT$, MKD(SCRAPMOBRA#), 47, 8)
    Call REPLA(TXT$, MKD(UTILIDAD#), 55, 8)
    Call REPLA(TXT$, MKD(IMPUESTOS#), 63, 8)
    Call REPLA(TXT$, MKD(COMIVENDED#), 71, 8)
    Call REPLA(TXT$, MKD(COSTOTOTAL#), 111, 8)
    Call REPLA(TXT$, MKD(IMPNET#), 119, 8)
    HOII% = HOY(HOS$)
    Call REPLA(TXT$, MKI(HOII%), 127, 2)
    Call REPLA(TXT$, MKS(MINMANOBRA_REAL#), 129, 4)
    '
    Call FILTERSETRECORD("SEGUITRE", 1, PERIOSEL$, TXT$)
    Call FILESORT("SEGUITRE", "", 2, 2)
    Call MENSERR(24, "Fin Calculo del Mes " + PERIOSEL)
    Exit Sub
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
End Sub


Sub LIPORFE()

    Call DESHASFECHA(DES%, HAS%, PERIO$)
    Call COPYSTRU("SEGUITRE", "SEGUITRF")
    FIN& = ULTREG&("SEGUITRE")
    j& = 0
    For JJ& = 1 To FIN&
        TXT$ = REGLEIDO$("SEGUITRE", JJ&)
        X% = CVI(Mid$(TXT$, 17, 2))
        If X% >= DES% And X% <= HAS% Then
            j& = j& + 1
            Call GRAREG("SEGUITRF", TXT$, j&)
        End If
    Next JJ&
    Call SETULTREG("SEGUITRF", j&)
    Call CIERRARCH("SEGUITRF")
    Call HEADERS("SEGUITRF", "")
    Call HEADERS("SEGUITRF", "Período: " + PERIO$)
    Call CONECRUN("*SEGUITRF", "Seguimiento de Items Recaudados")
    
End Sub

Private Sub Timer1_Timer()

   If MODIFIC%(0) > 0 Then
       If COMALTER%("Se ha Modificado Datos del Artículo\'" + TRIM$(CODEXT$(CIXII%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") = 1 Then
           Call Command1_Click
       End If
   End If
   MODIFIC%(0) = 0
   Timer1.Enabled = False

End Sub


