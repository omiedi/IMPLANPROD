VERSION 5.00
Begin VB.Form ANACOSIM 
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
         Top             =   700
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
         Top             =   700
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
      ToolTipText     =   "Doble Click Para Ver cotizacion de Dolar"
      Top             =   3840
      Width           =   10610
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
         Left            =   4880
         TabIndex        =   72
         Top             =   30
         Width           =   1485
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
         Left            =   3390
         TabIndex        =   71
         Top             =   30
         Width           =   1485
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
         Left            =   1910
         TabIndex        =   70
         Top             =   30
         Width           =   1485
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
         Width           =   1905
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Cotización del Dolar:"
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
         Left            =   7140
         TabIndex        =   68
         Top             =   70
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   8990
         TabIndex        =   67
         Top             =   30
         Visible         =   0   'False
         Width           =   1400
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
         Left            =   3195
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
         Left            =   2265
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
         Left            =   3195
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
         Left            =   2265
         TabIndex        =   57
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3675
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
         Left            =   3675
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
         Left            =   5985
         TabIndex        =   65
         Top             =   705
         Width           =   1395
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5985
         TabIndex        =   64
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4305
         TabIndex        =   63
         ToolTipText     =   "Doble Click para Editar"
         Top             =   705
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4305
         TabIndex        =   62
         ToolTipText     =   "Doble Click para Editar"
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pr.Venta (u$s) "
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
         Left            =   315
         TabIndex        =   59
         Top             =   750
         Width           =   1905
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ultima modif.     L-Pr.                   Neto                    Final"
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
         Left            =   2100
         TabIndex        =   56
         Top             =   45
         Width           =   5160
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pr.Venta ($) "
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
         Left            =   315
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
      Top             =   4500
      Width           =   6210
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1980
         TabIndex        =   53
         Top             =   315
         Width           =   900
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
         Caption         =   "   Mano de Obra:                 min   x                  $:          "
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
         Left            =   460
         TabIndex        =   50
         Top             =   350
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
      Height          =   2055
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
         Caption         =   $"ANACOSIM.frx":0000
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
         Picture         =   "ANACOSIM.frx":0087
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
         Picture         =   "ANACOSIM.frx":0391
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
      Picture         =   "ANACOSIM.frx":069B
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
         Picture         =   "ANACOSIM.frx":09A5
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
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSIM.frx":0CAF
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   4830
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
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSIM.frx":10F1
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Complementos, Consultas, Listados"
         Top             =   5520
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
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSIM.frx":13FB
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprime Ficha del Articulo"
         Top             =   6210
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
         Picture         =   "ANACOSIM.frx":1A65
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
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSIM.frx":1D6F
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Baja Lógica Individual"
         Top             =   3770
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
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSIM.frx":2079
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Sustitución de Código de Artículo"
         Top             =   3080
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
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSIM.frx":21C3
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   1680
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
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSIM.frx":230D
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Alta de Nuevo Item por Copia Completa de Datos de Artículo Existente"
         Top             =   2390
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
         Picture         =   "ANACOSIM.frx":2457
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   105
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "ANACOSIM.frx":25A1
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
         Picture         =   "ANACOSIM.frx":44C3
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
         Picture         =   "ANACOSIM.frx":47CD
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
      Y1              =   6405
      Y2              =   6405
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
Attribute VB_Name = "ANACOSIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim APREPES$, APREDOL$
Dim YACAR%(15), MODIFIC%(15)
Dim RESECO$(128), DESOPE$(128)


Sub PRECIOENLISTA(CI%)

    For U& = 1 To ULTREG&("LISTAS")
       XX$ = REGLEIDO$("LISTAS", U&)
       If Asc(Mid$(XX$, 50, 1)) = 1 Then           'VALIDA TIPO DE MONEDA = $
          APREPES1$ = TRIM$(Str$(U&))
          While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
          APREPES1$ = "LIPRE" + APREPES1$
          PREPES# = PRELISTA#(APREPES1$, CI%, FEX1%)
          If PREPES# >= 0.001 Then
             Label1 = FORMATNUM$(PREPES#, "F10.4")
             Label7 = FORMATNUM$(PREPES# * (1 + ALIVA / 100), "F10.4")
             Text6 = FECHATEX$(FEX1)
             Label15 = U&
             UI% = U&
             Label15.ToolTipText = DENOLISTA$(UI%)
             GoTo 20
          End If
          Call CIERRARCH(APREPES1$)
       End If
    Next U&
    '
20  For U& = 1 To ULTREG&("LISTAS")
        XX$ = REGLEIDO$("LISTAS", U&)
        If Asc(Mid$(XX$, 50, 1)) = 2 Then           'VALIDA TIPO DE MONEDA = U$S
            APREPES1$ = TRIM$(Str$(U&))
            While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
            APREPES1$ = "LIPRE" + APREPES1$
            PREDOL# = PRELISTA#(APREPES1$, CI%, FEX2%)
            If PREDOL# >= 0.005 Then
                Label6 = FORMATNUM$(PREDOL#, "F10.4")
                Label8 = FORMATNUM$(PREDOL# * (1 + ALIVA / 100), "F10.4")
                Text7 = FECHATEX$(FEX2)
                Label16 = U&
                UI% = U&
                Label16.ToolTipText = DENOLISTA$(UI%)
                Exit For
            End If
        End If
    Next U&

End Sub

Private Sub Abri_Click()
   Call Command9_Click
End Sub

Private Sub Acopia_Click()
   Call Command22_Click
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
   Call GRADATITEM
   Call GRAFORMULA
   '
   Call GRATIEMOD
   '
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
    Label1.Caption = FORMATNUM$(XVALO(Label103.Caption) + XVALO(Label110.Caption), "F10.4")
    If XVALO(Label11) > 0 Then
       Label6.Caption = FORMATNUM$((XVALO(Label103.Caption) + XVALO(Label110.Caption)) / XVALO(Label11.Caption), "F10.4")
    End If
    '
    Label7.Caption = FORMATNUM$(((XVALO(Label1.Caption) * ALIVA) / 100) + XVALO(Label1.Caption), "F10.4")
    Label8.Caption = FORMATNUM$(((XVALO(Label6.Caption) * ALIVA) / 100) + XVALO(Label6.Caption), "F10.4")
    Text6 = HOS$: Text7 = HOS$
    '
    '\\\OT-08-0095-FL-14/03/08///
    'ACTUALIZA LISTAS EN PESOS
    For U& = 1 To ULTREG&("LISTAS")
        XX$ = REGLEIDO$("LISTAS", U&)
        If Asc(Mid$(XX$, 50, 1)) = 1 Then       'VALIDA TIPO DE MONEDA = $
            APREPES1$ = TRIM$(Str$(U&))
            While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
            APREPES1$ = "LIPRE" + APREPES1$
            PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
            CODIN% = CVI(Mid$(PRE1$, 1, 2))        'VALIDA QUE EL CODIGO PERTENEZCA
            If CODIN% > 0 Then                     'A LA LISTA DE PRECIOS
                PREPES# = XVALO(Label1)
                FEXI% = FECHANUM(Text6)
                Call REPLA(PRE1$, MKI$(CI%), 1, 2)
                Call REPLA(PRE1$, MKD$(PREPES#), 62, 8)
                Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
                Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
                SINLIST% = 1
            End If
         End If
    Next U&
    '
    'ACTUALIZA LISTAS EN DOLARES
    For U& = 1 To ULTREG&("LISTAS")
        XX$ = REGLEIDO$("LISTAS", U&)
        If Asc(Mid$(XX$, 50, 1)) = 2 Then       'VALIDA TIPO DE MONEDA = U$S
            APREPES1$ = TRIM$(Str$(U&))
            While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
            APREPES1$ = "LIPRE" + APREPES1$
            PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
            CODIN% = CVI(Mid$(PRE1$, 1, 2))        'VALIDA QUE EL CODIGO PERTENEZCA
            If CODIN% > 0 Then                     'A LA LISTA DE PRECIOS
                PREDOL# = XVALO(Label6)
                FEXI% = FECHANUM(Text6)
                Call REPLA(PRE1$, MKI$(CI%), 1, 2)
                Call REPLA(PRE1$, MKD$(PREDOL#), 62, 8)
                Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
                Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
                SINLIST% = 1
            End If
         End If
    Next U&
    '\\\OT-08-0095-FL-FIN///
    
        '\\\OT-08-0136-FL-04/04/08///
    If SINLIST% = 0 Then
         Call SELECHO("LISTAS")
         APREPES1$ = TRIM$(VALACT1$("LISTAS"))
         While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
         APREPES1$ = "LIPRE" + APREPES1$
         PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
         If Asc(Mid$(XX$, 50, 1)) = 2 Then
            PREPES# = XVALO(Label6)
         Else
            PREPES# = XVALO(Label1)
         End If
         FEXI% = FECHANUM(Text6)
         Call REPLA(PRE1$, MKI$(CI%), 1, 2)
         Call REPLA(PRE1$, MKD$(PREPES#), 62, 8)
         Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
         Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
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
         Master.FindFirst ("CODINT = " & CICOMPO%)
             If Master.NoMatch = False Then
                PORDE1 = ROUND#(Master!PorDescCos, 1)
                If PORDE1 > 99 Then PORDE1 = 0
             End If
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
                 Master.FindFirst ("CODINT = " & CICOMPO%)
                 If Master.NoMatch = False Then
                   Master.Edit
                   Master!COSUNI = COSTO1
                   Master!MONECOS = MONEI%
                   Master!FECOSTO = CVDAT(FECH1%)
                   Master!PorDescCos = PORDE1
                   Master!PROVHABI = PROV1%
                   Master.Update
                   Call COMUNI("Costo Actualizado")
                   Call CARFORMULA
                   '
                   Call CALCUCOS
                   '
                 End If
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
     CACAM% = Articulos.TableDefs("Master").Fields.Count
     Master.FindFirst ("CODINT = " & CII%)
     If Master.NoMatch = False Then
       Master.Edit
       For KU% = 1 To CACAM%
         FIELDVAL(KU%) = Master.Fields(KU% - 1).Value
       Next KU%
       '
       Master.FindFirst ("CODINT = " & CINUE%)
       If Master.NoMatch = False Then
         Master.Edit
         For KU% = 1 To CACAM%
           NOCAMPO$ = TRIM$(UCase$(Master.Fields(KU% - 1).Name))
           If NOCAMPO$ <> "CODIGO" Then
             If NOCAMPO$ <> "DESCRIPCION" Then
               If NOCAMPO$ <> "CODINT" Then
                 If NOCAMPO$ <> "ALIAS" Then
                   If NOCAMPO$ <> "UMEDIDA" Then
                     If NOCAMPO$ <> "GRUCOVEN" Then
                      Master.Fields(KU% - 1).Value = FIELDVAL(KU%)
                     End If
                   End If
                 End If
               End If
             End If
           End If
         Next KU%
         Master.Update
       End If
       '
       Call COPYFORMU(CII%, CINUE%)
       Call COPYRUTA(CII%, CINUE%)
     End If
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
     Master.FindFirst ("CODINT = " & CI%)
     Master.Edit
     Master!Codigo = CODINUE$
     Master.Update
     UPDATMASTER% = 1
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
     Master.FindFirst ("CODINT = " & CI%)
     Master.Edit
     Master!Status = (-1)
     Master.Update
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
      NULIS& = ULTREG&("LISTAS")
      For U& = 1 To NULIS&
         Screen.MousePointer = 11
         LISNO% = U&
         ALISTA$ = TRIM$(Str$(LISNO%)): While Len(ALISTA$) < 3
                                          ALISTA$ = "0" + ALISTA$
                                        Wend
         APRECIO$ = "LIPRE" + ALISTA$
         If ULTREG&(APRECIO$) > 0 Then
            Call FILTERDELETE(APRECIO$, 1, MKI$(CII%))
         End If
      Next U&
   End If
   '
   Screen.MousePointer = 1
   Command24.Enabled = True
   '10/04/07 (FIN)
   '
End Sub


Private Sub Command3_Click()
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
   OPLIMAS04.Show 1
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
10          Call SELECHO("DEUDOR1")
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
10  OPT% = ALTERNA%("Actualización de Costos por Proveedor\Seguimiento de Items Recaudados\ - \Acceso a Archivos Maestros\Grupos Contables de Venta\Listas de Precios de Venta\ - \Otros Accesos Directos")
    If OPT% = 1 Then
         Call ACTULISCOSPRO
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
         Call ACCDIR("ANACOS")
    End If
   '
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
   Call GRADATITEM
   '
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
   Call CALCUCOS '09/04/07 (OT-07-0137)
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
        Master.FindFirst ("CODINT = " & CI%)
        Master.Edit
        Master!DESCRIPCION = DE$
        Master!NPLANO = TRIM$(Left$(Text45, 24))
        Master!PorComVen = XVALO(Text42)
        Master.Update
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
   DCC1$ = TRIM$(CONTROL$("ANACOSTO", "GAGENMAT"))
   Ltext41 = FORMATNUM$(XVALO(DCC1$), "F6.2")
   '
   DCC2$ = TRIM$(CONTROL$("ANACOSTO", "GAGENFAB"))
   Ltext36 = FORMATNUM$(XVALO(DCC2$), "F6.2")
   '
   DCC3$ = TRIM$(CONTROL$("ANACOSTO", "PORUTILI"))
   Ltext39 = FORMATNUM$(XVALO(DCC3$), "F6.2")
   '
   DCC4$ = TRIM$(CONTROL$("ANACOSTO", "PORIMPUE"))
   Ltext43 = FORMATNUM$(XVALO(DCC4$), "F6.2")
   '
   '//OT-07-0477//
   '
   Ltext5 = FORMATNUM$((VALHORA(0) / 60), "F7.4")
   '
   '//OT-07-0477//
   '
   APREPES$ = TRIM$(CONTROL$("ANACOSTO", "LIPREPES"))
   If APREPES$ = "" Or XVALO(APREPES$) < 1 Or XVALO(APREPES$) > ULTREG&("LISTAS") Then
      Call MENSERR(24, "Error de Configuración de Listas de Precios")
      Call Command37_Click
      Exit Sub
   End If
   While Len(APREPES$) < 3: APREPES$ = "0" + APREPES$: Wend
   APREPES$ = "LIPRE" + APREPES$
   '
   APREDOL$ = CONTROL$("ANACOSTO", "LIPREDOL")
   If APREDOL$ = "" Or XVALO(APREDOL$) < 1 Or XVALO(APREDOL$) > ULTREG&("LISTAS") Then
      Call MENSERR(24, "Error de Configuración de Listas de Precios")
      Call Command37_Click
      Exit Sub
   End If
   While Len(APREDOL$) < 3: APREDOL$ = "0" + APREDOL$: Wend
   APREDOL$ = "LIPRE" + APREDOL$
   '
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

Private Sub Label1_DblClick()
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
    Label1.Caption = FORMATNUM$(XVALO(NUPRECIO$), "F10.4")
    Label7.Caption = FORMATNUM$(((XVALO(Label1.Caption) * ALIVA) / 100) + XVALO(Label1.Caption), "F10.4")
    Text6 = HOS$ ': Text7 = HOS$
    '
    '\\\OT-08-0095-FL-14/03/08///
    'ACTUALIZALISTA EN PESOS
    SINLIST% = 0
    For U& = 1 To ULTREG&("LISTAS")
        XX$ = REGLEIDO$("LISTAS", U&)
        If Asc(Mid$(XX$, 50, 1)) = 1 Then           'VALIDA TIPO DE MONEDA = $
            APREPES1$ = TRIM$(Str$(U&))
            While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
            APREPES1$ = "LIPRE" + APREPES1$
            PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
            CODIN% = CVI(Mid$(PRE1$, 1, 2))            'VALIDA QUE EL CODIGO PERTENEZCA
            If CODIN% > 0 Then                         'A LA LISTA DE PRECIOS
                PREPES# = XVALO(Label1)
                FEXI% = FECHANUM(Text6)
                Call REPLA(PRE1$, MKI$(CI%), 1, 2)
                Call REPLA(PRE1$, MKD$(PREPES#), 62, 8)
                Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
                Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
                SINLIST% = 1
            End If
        End If
    Next U&
    '\\\OT-08-0095-FL-FIN///
    '
    '\\\OT-08-0136-FL-04/04/08///
    If SINLIST% = 0 Then
         Call SELECHO("LISTAS")
         APREPES1$ = TRIM$(VALACT1$("LISTAS"))
         Label15 = APREPES1$
         Label15.ToolTipText = VALACT2$("LISTAS")
         While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
         APREPES1$ = "LIPRE" + APREPES1$
         PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
         PREPES# = XVALO(Label1)
         FEXI% = FECHANUM(Text6)
         Call REPLA(PRE1$, MKI$(CI%), 1, 2)
         Call REPLA(PRE1$, MKD$(PREPES#), 62, 8)
         Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
         Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
    End If
    '\\\OT-08-0136-FL-FIN///
    '
    Call CALCUCOS
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



Private Sub Label27_DblClick()
   '04/03/07 (OT-07-0137)
   Call List1_DblClick
   '04/03/07 (FIN)
End Sub



Private Sub Label6_DblClick()
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
    Label6.Caption = FORMATNUM$(XVALO(NUPRECIO$), "F10.4")
    Label8.Caption = FORMATNUM$(((XVALO(Label6.Caption) * ALIVA) / 100) + XVALO(Label6.Caption), "F10.4")
    Text7 = HOS$
    '
    '\\\OT-08-0095-FL-14/03/08///
    'ACTUALIZA LISTAS EN DOLARES
    For U& = 1 To ULTREG&("LISTAS")
        XX$ = REGLEIDO$("LISTAS", U&)
        If Asc(Mid$(XX$, 50, 1)) = 2 Then           'VALIDA TIPO DE MONEDA = U$S
            APREPES1$ = TRIM$(Str$(U&))
            While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
            APREPES1$ = "LIPRE" + APREPES1$
            PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
            CODIN% = CVI(Mid$(PRE1$, 1, 2))            'VALIDA QUE EL CODIGO PERTENEZCA
            If CODIN% > 0 Then                         'A LA LISTA DE PRECIOS
                PREDOL# = XVALO(Label6)
                FEXI% = FECHANUM(Text6)
                Call REPLA(PRE1$, MKI$(CI%), 1, 2)
                Call REPLA(PRE1$, MKD$(PREDOL#), 62, 8)
                Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
                Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
                SINLIST% = 1
            End If
         End If
    Next U&
    '\\\OT-08-0095-FL-FIN///
    
    '\\\OT-08-0136-FL-04/04/08///
    If SINLIST% = 0 Then
         Call SELECHO("LISTAS")
         APREPES1$ = TRIM$(VALACT1$("LISTAS"))
         Label16 = APREPES1$
         Label16.ToolTipText = VALACT2$("LISTAS")
         While Len(APREPES1$) < 3: APREPES1$ = "0" + APREPES1$: Wend
         APREPES1$ = "LIPRE" + APREPES1$
         PRE1$ = FILTERGETRECORD$(APREPES1$, 1, MKI$(CI%))
         PREPES# = XVALO(Label6)
         FEXI% = FECHANUM(Text6)
         Call REPLA(PRE1$, MKI$(CI%), 1, 2)
         Call REPLA(PRE1$, MKD$(PREPES#), 62, 8)
         Call REPLA(PRE1$, MKI$(FEXI%), 70, 2)
         Call FILTERPUTRECORD(APREPES1$, 1, MKI$(CI%), PRE1$)
    End If
    '\\\OT-08-0136-FL-FIN///
    '
    Call CALCUCOS
    '
End Sub



Private Sub List1_DblClick()
    '
    '04/03/07 (OT-07-0137)
    If Text1 <> "" Then
        Call BLANARCH("!BOLREDET")
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
         End If
         Call CALCUCOS '09/04/07 (OT-07-0137)
    End If
    '04/03/07 (FIN)
    '
End Sub

Private Sub Listad_Click()
   Call OPLIMAS04.Command2_Click
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
      Else
        Label9.Visible = False
        Label11.Visible = False
        Picture2.ToolTipText = "Doble Clic Para Ver cotizacion de Dolar"
     End If
  End If
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
   Call BLANFORMU
   If CIXII% > 0 Then
      Call CARDATITEM
      Call CARFORMULA
      Call CARTIEMOD
      Call CALCUCOS
   End If
   '05/03/07 (FIN)
   '
End Sub

Private Sub Text1_DblClick()
   '
   Call Command9_Click '05/03/07 (OT-07-0137)
   '
End Sub

Private Sub Text2_Change()
   Call CALCOSOPE
End Sub

Private Sub Text2_LostFocus()
   '
   Text2 = FORMATNUM$(XVALO(Text2), "F6.2") '30/03/07 (OT-07-0137)
   '
End Sub

Private Sub Text42_Change()
   '
   Call CALCUCOS
   '
End Sub

Private Sub Text42_LostFocus()
   '
   Text42 = FORMATNUM$(XVALO(Text42), "F5.2") '30/03/07 (OT-07-0137)
   '
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
    FOCOREP% = 1
    '
    Call PRECIOENLISTA(CI%)
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
   '
   List1.Clear
   '
   PREVENETO# = 0
   '
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

   If ECOARCH$("GRUCOVEN", MKI(GRCONTA%)) = "" Then
    GRCONTA% = 0
   End If


   Screen.MousePointer = 11
   CODII% = NUMAS% + 1
   '
20 On Error Resume Next
   Call ABREMASTER
   Master.MoveFirst
   If Err < 1 Then
     On Error GoTo 0
     ABUS$ = "Codint = " & CODII%
     Master.FindFirst "Codint = " & CODII%
     If Master.NoMatch = False Then
       CODII% = CODII% + 1
       GoTo 20
     End If
   End If
   '
   With Master
      .AddNew
      !Codigo = CODINUE$
      !DESCRIPCION = DENUE$
      !Status = 1
      !CODINT = CODII%
      !UMEDIDA = UNID$
      !TIMATE = TIMA%
      !GRUCOVEN = GRCONTA%
      On Error Resume Next
      .Update
      OKX% = 1
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
         OKX% = 0
      End If
      On Error GoTo 0
      .Bookmark = .LastModified
      .Close
      UPDATMASTER% = 1
   End With
   '
99 Screen.MousePointer = 1

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
66 With Secoper
      NOPYI% = 10
      CANPRE = 0: CATPRE% = 0
      .FindFirst "CODICONJ =" & CI% & " and NumeOper = " & NOPYI%
      If .NoMatch = True Then
          .AddNew
         Else
          .Edit
      End If
      '
      !CODICONJ = CI%
      !CODXCONJ = CODEXT$(CI%)
      !NumeOper = NOPYI%
      !DescOper = "Mano de Obra Directa"
      !statoper = 1
      !CodMaq0 = ""
      !CodMaq1 = ""
      !CodMaq2 = ""
      !CodMaq3 = ""
      !HorsPrep = 0
      !HorsFijo = 0
      !MinuStan = XVALO(Text2)
      !HorsLimp = 0
      !HorsMHer = 0
      !PieCiclo = 0
      !PorDemor = 0
      If !MinuStan > 0.005 Then
        !Cadencia = 60 / !MinuStan
      End If
      !CantiOps = 1
      !CategOps = 0
      !CoefSolp = 0
      !LOTESTAN = 1
      !ProgCNum = ""
      !HoProFab = ""
      !ProTerOp = 0
      !PreUnid = 0
      !MonePrec = 0
      !FechPrec = 0
      '
      !OBSEOPER = ""
      !NUOPER_PREP = 0
      !CATOPER_PREP = 0
      .Update
      '
   End With   ' secoper
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
66 With Secoper
      NOPYI% = 10
      CANPRE = 0: CATPRE% = 0
      .FindFirst "CODICONJ =" & CI% & " and NumeOper = " & NOPYI%
      If .NoMatch = False Then
         Text2 = FORMATNUM$(!MinuStan, "F6.2")
      End If
   End With   ' secoper
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
   'FORIPRE$ = "PLANICOS"
   'If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
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
        CAPARDOC% = 25
        '
        CODTABU1$(1) = "COD-MPRI"
        CODTABU1$(2) = "DES-MPRI"
        CODTABU1$(3) = "UNI-MPRI"
        CODTABU1$(4) = "CAN-MPRI"
        CODTABU1$(5) = "IMP-NETO"
        CODTABU1$(6) = "SUB-TOTA"
        CACOLTAB1% = 6
        '
        For KK1% = 1 To 6
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
            PUNIT# = XVALO(Mid$(Z$, 67, 9))
            PTOTA# = XVALO(Mid$(Z$, 76, 12))
            '
            TETABU1$(1, U&) = COD$
            TETABU1$(2, U&) = DESCRIT0$(CODI1%)
            TETABU1$(3, U&) = NOMED$
            TETABU1$(4, U&) = FORMATNUM$(CANTI#, "F10.4")
            TETABU1$(5, U&) = FORMATNUM$(PUNIT#, "F10.4")
            TETABU1$(6, U&) = FORMATNUM$(PTOTA#, "F10.4")
        Next U&
        '
        CAITAB1% = List1.ListCount
        '
        Call PRINTDOC("PLANICOS")
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
      Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      REBLA$ = REGBLAN$("SECOPER") + String$(64, 0)
      IOP% = 0
      On Error Resume Next
      With SecopTemp
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
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Hoja de Ruta del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(3) = 0
       MODIFIC%(3) = 0
       'Call CARGENITEM
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   Call ABRESECOP
   '
   With Secoper
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       GoTo 66
     End If
     YAPRIMO% = 0
     ABUS$ = "CodIConj =" & CI% & " AND StatOper > 0"
     On Error Resume Next
62   If YAPRIMO% = 0 Then
         .FindFirst ABUS$
         YAPRIMO% = 1
       Else
621      .FindNext ABUS$
     End If
     If .NoMatch = False Then
       If IsNull(!AltOper) = True Or !AltOper < 1 Then
         .Edit
         !statoper = (-1)
         .Update
         GoTo 62
       End If
       GoTo 621
     End If
   End With
   '
66 With Secoper
   For LL% = 1 To List3.ListCount
    NOPYI% = CVI(Mid$(RESECO$(LL% - 1), 3, 2))
    CANPRE = 0: CATPRE% = 0
     .FindFirst "CODICONJ =" & CI% & " and NumeOper = " & NOPYI%
     If .NoMatch = False Then
       CANPRE = !NUOPER_PREP
       CATPRE% = !CATOPER_PREP
     End If
     '
     .AddNew
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
     '
   Next LL%
   '
63 .FindFirst "CODICONJ = " & CI%
   If .NoMatch = False Then
     If !statoper = (-1) Then
       .Delete
       GoTo 63
     End If
   End If
   '
64 .FindNext "CODICONJ = " & CI%
   If .NoMatch = False Then
     If !statoper = (-1) Then
       .Delete
       GoTo 63
     End If
     GoTo 64
   End If
   '
   .Close
   End With   ' secoper
   '
   MODIFIC(3) = 0
   Screen.MousePointer = 1
   '
End Sub


Sub ACTULISCOSPRO()
10  Call SELECHO("PROVED1")
    NCX1% = XVALO(VALACT1$("PROVED1"))
    If NCX1% > 0 Then
       '
       Call APERBASDAT("MASTER")
       '
       Set MasterBaja = Articulos.OpenRecordset("SELECT * FROM Master WHERE PROVHABI = " & NCX1% & " AND Status > 0 ORDER BY CODIGO ASC;", dbOpenSnapshot)
       With MasterBaja
       On Error Resume Next
       MasterBaja.MoveFirst
       If Err Then
          On Error GoTo 0
          Call MENSERR(24, "No hay Registros que Coincidan con el Proveedor Elegido .")
          GoTo 20
       End If
       On Error GoTo 0
       '
20         Call BLANARCH("!ACTUCOS")
       '
       I& = 0
       Do While Not .EOF
          COD$ = !Codigo
          MON% = !MONECOS
          PODES1 = ROUND#(!PorDescCos, 1): If PODES1 > 99 Then PODES1 = 0
          COS1# = !COSUNI / (1 - PODES1 / 100)
          COS2# = !COSUNI
          FECH = !FECOSTO
          '
          Y$ = REGBLAN$("ACTUCOS")
          CODI% = CODINT%(COD$)
          Call REPLA(Y$, MKI$(CODI%), 1, 2)
          Call REPLA(Y$, MKI$(MON%), 3, 2)
          Call REPLA(Y$, MKD$(COS1#), 5, 8)
          Call REPLA(Y$, MKS$(PODES1), 13, 4)
          Call REPLA(Y$, MKD$(COS2#), 17, 8)
          FCOS% = CVINT%(FECH)
          Call REPLA(Y$, MKI$(FCOS%), 25, 2)
          I& = I& + 1
          Call GRAREG("!ACTUCOS", Y$, I&)
       .MoveNext
       '
       Loop
       '
       Call SETULTREG("!ACTUCOS", I&)
       '
       End With
       '
       JUN% = VENTABU%("ACTUCOS/TITUPAN=Costo de Reposición - " + RASOCLI$((-1) * NCX1%))
       '
       If JUN% < 0 Then GoTo 10
       If JUN% >= 0 Then
        Call ABREMASTER
        For J& = 1 To ULTREG&("!ACTUCOS")
            Z$ = REGLEIDO$("!ACTUCOS", J&)
            COD3% = CVI(Mid$(Z$, 1, 2))
              PODES1 = CVS(Mid$(Z$, 13, 4)): If PODES1 < 0 Or PODES1 > 99 Then PODES1 = 0
            COU = (CVD(Mid$(Z$, 5, 8))) * (1 - PODES1 / 100)
            MONEI% = CVI(Mid$(Z$, 3, 2))
            FECOS% = CVI(Mid$(Z$, 25, 2))
            PHAB% = PROVHABI%(COD3%)
            '
            If PHAB% > 0 And NCX1% <> PHAB% Then
               TTXX$ = "Para el Artículo " + TRIM$(CODEXT$(COD3%)) + "\el Proveedor Habitual Registrado es\"
               TTXX$ = TTXX$ + TRIM$(RASOCLI$((-1) * PHAB%)) + "\  \"
               TTXX$ = TTXX$ + "Desea Cambiar el nuevo Proveedor Habitual a\"
               TTXX$ = TTXX$ + RASOCLI$((-1) * NCX1%)
               If COMALTER%(TTXX$ + "\\No, Conservar Anterior\Si, Cambiar") = 1 Then GoTo 10
            End If
            '
            Master.FindFirst ("CODINT =" & COD3%)
            If Master.NoMatch = False Then
               Master.Edit
               Master!PROVHABI = NCX1%
               Master!COSUNI = COU
               Master!PorDescCos = PODES1
               Master!FECOSTO = CVDAT(FECOS%)
               Master!MONECOS = MONEI%
               Master.Update
           End If
            '
        Next J&
        '
        Call CARFORMULA
        '
       End If
    End If

End Sub
'
Sub ARCHMAES()
    Call CONECRUN("ARCHIG04", "Archivos Maestros")
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
    J& = 0
    For I& = 1 To ULTREG&("!GRUCOVEN")
       X$ = REGLEIDO$("!GRUCOVEN", I&)
       If CVI(Left$(X$, 2)) > 0 Then
           J& = J& + 1
           Call GRAREG("GRUCOVEN", X$, J&)
       End If
    Next I&
    Call SETULTREG("GRUCOVEN", J&)
    Call CIERRARCH("GRUCOVEN")

End Sub


Sub LISTAPRE()
    Call CONECRUN("LISPRE04", "Listas de Precios de Venta")
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
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
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
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
      TOTCANT#(U&) = 0
      TOTCOSTO#(U&) = 0
    Next U&

    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'VENTA' "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      '
      Call APERBASDAT("STOCKS")
      Screen.MousePointer = 11
      ARTMAX% = 0
      KU& = 0
      '
      Do While Not .EOF
        KU& = KU& + 1
        Call TRACE(20, KU&, FIN&)
        TOTFAC# = ROUND(!IHABECOMP - !IDEBECOMP, 2)
        NCLIE% = !NuClien
        If Abs(TOTFAC#) >= 0.005 Then
          NROFAC$ = !CodiCom_Lar
          SQLX$ = "SELECT * FROM MOVISTO "
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
          Set STOTEMP1 = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
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
              CIXI% = !Cod_Inte
              CANTU = (!Cantsalid - !CantIngre)
              PREUN# = !ValoUnit
              If !MoneVal > 1 Then
                PREUN# = PREUN# * !TipoCam
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
              nart% = !Cod_Inte: If nart% < 0 Then nart% = 0
              If nart% > ARTMAX% Then ARTMAX% = nart%
              '\\\OT-06-0494-RG-20/11/06///
              CANTI# = (!Cantsalid - !CantIngre)
              TOTCANT#(nart%) = TOTCANT#(nart%) + CANTI#
              '\\\OT-06-0494-RG-FIN///
              NETO# = (!Cantsalid - !CantIngre) * !ValoUnit * COEFIMA#
              COSTONETO# = ROUND((!Cantsalid - !CantIngre) * !CostoRep, 2)
              If !MoneVal > 1 Then
                NETO# = NETO# * !TipoCam
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
        End If
        '
30    .MoveNext
      Loop
      '
    End With                          ' SOBLE CAJABANC
    '
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
    J& = 0: PORACU = 0
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
        
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call REPLA(Z$, MKD$(TOTCANT#(NC%)), 21, 8) '10/05/07 (OT-07-0188)
        Call GRAREG("RANKPIE", Z$, J&)
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
        Call GRAREG("RANMABRU", W$, J&)
        '\\\OT-06-0494-RG-FIN///
    Next U&
    
    Call SETULTREG("RANKPIE", J&)
    Call CIERRARCH("RANKPIE")

    Call SETULTREG("RANMABRU", J&)
    Call CIERRARCH("RANMABRU")
 
    List2.Clear
  
    FIN& = ULTREG&("RANKPIE")
    For JX& = 1 To FIN&
        Call TRACE(20, JX&, FIN&)
        XX$ = REGLEIDO$("RANKPIE", JX&)
        CODIG% = CVI(Mid$(XX$, 3, 2))
        CANITEM# = CVD(Mid$(XX$, 21, 8))
        Text1.TEXT = CODEXT$(CODIG%)
        openforms = DoEvents
        MINMANOBRA# = MINMANOBRA# + XVALO(Text2.TEXT) * CANITEM#
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
    J& = 0
    For JJ& = 1 To FIN&
        TXT$ = REGLEIDO$("SEGUITRE", JJ&)
        X% = CVI(Mid$(TXT$, 17, 2))
        If X% >= DES% And X% <= HAS% Then
            J& = J& + 1
            Call GRAREG("SEGUITRF", TXT$, J&)
        End If
    Next JJ&
    Call SETULTREG("SEGUITRF", J&)
    Call CIERRARCH("SEGUITRF")
    Call HEADERS("SEGUITRF", "")
    Call HEADERS("SEGUITRF", "Período: " + PERIO$)
    Call CONECRUN("*SEGUITRF", "Seguimiento de Items Recaudados")
    
End Sub
