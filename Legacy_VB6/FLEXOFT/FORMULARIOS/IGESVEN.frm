VERSION 5.00
Begin VB.Form IGESVEN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Gestion de Ventas"
   ClientHeight    =   8220
   ClientLeft      =   390
   ClientTop       =   300
   ClientWidth     =   11205
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11205
   Begin VB.CommandButton Command20 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10185
      TabIndex        =   84
      ToolTipText     =   "Análisis por Cliente y Artículo"
      Top             =   6510
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Clear"
      Height          =   855
      Left            =   5130
      Picture         =   "IGESVEN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Anular y Re-Frescar Costos Registrados"
      Top             =   1775
      Width           =   750
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Costos"
      Height          =   1650
      Left            =   10185
      TabIndex        =   79
      Top             =   4760
      Width           =   875
      Begin VB.CommandButton Command6 
         Height          =   645
         Left            =   105
         Picture         =   "IGESVEN.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Planilla de Carga de Costos Informados"
         Top             =   350
         Width           =   645
      End
      Begin VB.TextBox MERMAPRO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   80
         Text            =   " "
         ToolTipText     =   "Porcentaje Adicional Sobre Costos Calculados"
         Top             =   1180
         Width           =   540
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "%"
         Height          =   330
         Left            =   660
         TabIndex        =   83
         Top             =   1260
         Width           =   150
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Plan"
      Height          =   800
      Left            =   10185
      Picture         =   "IGESVEN.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   78
      ToolTipText     =   "Planillas de División por Producto y Mercado"
      Top             =   3875
      Width           =   855
   End
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   5040
      Sorted          =   -1  'True
      TabIndex        =   77
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mon"
      Height          =   1455
      Left            =   5145
      TabIndex        =   61
      Top             =   240
      Width           =   750
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "U$S"
         Height          =   645
         Left            =   150
         TabIndex        =   63
         Top             =   735
         Value           =   -1  'True
         Width           =   450
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "$"
         Height          =   330
         Left            =   150
         TabIndex        =   62
         Top             =   315
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Listados Resúmen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   5985
      TabIndex        =   64
      Top             =   210
      Width           =   4005
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1365
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   74
         Top             =   1590
         Width           =   2325
         Begin VB.CommandButton Command18 
            Height          =   600
            Left            =   0
            Picture         =   "IGESVEN.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Ranking de Ventas por Articulo por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Ranking de Ven- tas por Articulo"
            Height          =   540
            Left            =   735
            TabIndex        =   76
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   71
         Top             =   1200
         Width           =   2325
         Begin VB.CommandButton Command19 
            Height          =   600
            Left            =   0
            Picture         =   "IGESVEN.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Ranking de Ventas por Cliente por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Ranking de Ven- tas por Cliente"
            Height          =   540
            Left            =   670
            TabIndex        =   73
            Top             =   -10
            Width           =   1605
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   65
         Top             =   810
         Width           =   2325
         Begin VB.CommandButton Command15 
            Height          =   600
            Left            =   -20
            Picture         =   "IGESVEN.frx":0EAA
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Resúmenes de Entregas por Cliente y Artículo "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen de En- tregas por Cliente"
            Height          =   750
            Left            =   630
            TabIndex        =   67
            Top             =   -10
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   68
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command17 
            Height          =   600
            Left            =   0
            Picture         =   "IGESVEN.frx":11B4
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Listado Ventas Totales por Artículo "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen de Ven- tas por Producto"
            Height          =   600
            Index           =   4
            Left            =   735
            TabIndex        =   70
            Top             =   -10
            Width           =   1620
         End
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Informe"
      Height          =   800
      Left            =   10185
      Picture         =   "IGESVEN.frx":14BE
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Informe Mensual de Ventas"
      Top             =   2985
      Width           =   855
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Análisis por Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5370
      Left            =   5145
      TabIndex        =   41
      Top             =   2730
      Width           =   4845
      Begin VB.Frame Frame13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Distribución"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   210
         TabIndex        =   48
         Top             =   315
         Width           =   4425
         Begin VB.CommandButton Command14 
            Height          =   540
            Left            =   3780
            Picture         =   "IGESVEN.frx":17C8
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   210
            Width           =   540
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Costo Por Tipo de Ventas"
            Height          =   225
            Left            =   420
            TabIndex        =   53
            Top             =   420
            Width           =   2850
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Venta Vs. Costo Materias Primas"
            Height          =   225
            Left            =   420
            TabIndex        =   52
            Top             =   735
            Value           =   -1  'True
            Width           =   3480
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "% Margen Bruto por Producto"
            Height          =   225
            Left            =   420
            TabIndex        =   51
            Top             =   1050
            Width           =   3060
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "V.Abs Margen Bruto por Producto"
            Height          =   225
            Left            =   420
            TabIndex        =   50
            Top             =   1365
            Width           =   3585
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Participación de Ventas por Producto"
            Height          =   225
            Left            =   420
            TabIndex        =   49
            Top             =   1680
            Width           =   3585
         End
      End
      Begin VB.CommandButton Command13 
         Height          =   540
         Left            =   3990
         Picture         =   "IGESVEN.frx":1AD2
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   525
         Width           =   540
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Series Históricas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   210
         TabIndex        =   42
         Top             =   2520
         Width           =   4425
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ventas y Margen Total"
            Height          =   225
            Left            =   420
            TabIndex        =   46
            Top             =   525
            Value           =   -1  'True
            Width           =   2850
         End
         Begin VB.ListBox List2 
            Height          =   1620
            Left            =   420
            TabIndex        =   45
            Top             =   915
            Width           =   2850
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rango"
            Height          =   1695
            Left            =   3360
            TabIndex        =   44
            Top             =   840
            Width           =   960
            Begin VB.OptionButton Option1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "48"
               Height          =   225
               Index           =   3
               Left            =   210
               TabIndex        =   59
               Top             =   1260
               Width           =   645
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "36"
               Height          =   225
               Index           =   2
               Left            =   210
               TabIndex        =   58
               Top             =   945
               Width           =   645
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "24"
               Height          =   225
               Index           =   1
               Left            =   210
               TabIndex        =   57
               Top             =   630
               Value           =   -1  'True
               Width           =   645
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "12"
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   56
               Top             =   315
               Width           =   645
            End
         End
         Begin VB.CommandButton Command12 
            Height          =   540
            Left            =   3780
            Picture         =   "IGESVEN.frx":1DDC
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   210
            Width           =   540
         End
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10185
      TabIndex        =   29
      ToolTipText     =   "Re-Calculo General"
      Top             =   6510
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Análisis por Mercado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5370
      Left            =   210
      TabIndex        =   8
      Top             =   2730
      Width           =   4845
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Series Históricas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   210
         TabIndex        =   32
         Top             =   2520
         Width           =   4425
         Begin VB.CommandButton Command11 
            Height          =   540
            Left            =   3780
            Picture         =   "IGESVEN.frx":20E6
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   210
            Width           =   540
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rango"
            Height          =   1695
            Left            =   3360
            TabIndex        =   35
            Top             =   840
            Width           =   960
            Begin VB.OptionButton Option3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "48"
               Height          =   225
               Index           =   3
               Left            =   210
               TabIndex        =   39
               Top             =   1260
               Width           =   600
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "36"
               Height          =   225
               Index           =   2
               Left            =   210
               TabIndex        =   38
               Top             =   945
               Width           =   600
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "24"
               Height          =   225
               Index           =   1
               Left            =   210
               TabIndex        =   37
               Top             =   630
               Value           =   -1  'True
               Width           =   600
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "12"
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   36
               Top             =   315
               Width           =   600
            End
         End
         Begin VB.ListBox List1 
            Height          =   1620
            Left            =   420
            TabIndex        =   34
            Top             =   915
            Width           =   2850
         End
         Begin VB.OptionButton Option21 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ventas y Margen Total"
            Height          =   225
            Left            =   420
            TabIndex        =   33
            Top             =   525
            Value           =   -1  'True
            Width           =   2850
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Distribución"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   210
         TabIndex        =   9
         Top             =   315
         Width           =   4425
         Begin VB.CommandButton Command8 
            Height          =   540
            Left            =   3780
            Picture         =   "IGESVEN.frx":23F0
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   210
            Width           =   540
         End
         Begin VB.OptionButton Option15 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Participación de Ventas por Mercado"
            Height          =   225
            Left            =   420
            TabIndex        =   31
            Top             =   1680
            Width           =   3585
         End
         Begin VB.OptionButton Option14 
            BackColor       =   &H00C0FFFF&
            Caption         =   "V.Abs Margen Bruto por Mercado"
            Height          =   225
            Left            =   420
            TabIndex        =   30
            Top             =   1365
            Width           =   3585
         End
         Begin VB.OptionButton Option13 
            BackColor       =   &H00C0FFFF&
            Caption         =   "% Margen Bruto por Mercado"
            Height          =   225
            Left            =   420
            TabIndex        =   12
            Top             =   1050
            Width           =   3060
         End
         Begin VB.OptionButton Option12 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Venta Vs. Costo Materias Primas"
            Height          =   225
            Left            =   420
            TabIndex        =   11
            Top             =   735
            Value           =   -1  'True
            Width           =   3480
         End
         Begin VB.OptionButton Option11 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Costo Por Tipo de Ventas"
            Height          =   225
            Left            =   420
            TabIndex        =   10
            Top             =   420
            Width           =   2850
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Período Mensual Activo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   4845
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   720
         Left            =   3780
         ScaleHeight     =   660
         ScaleWidth      =   60
         TabIndex        =   27
         Top             =   420
         Width           =   120
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   5000
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tipo de Cambio"
         Height          =   645
         Left            =   2625
         TabIndex        =   24
         Top             =   1280
         Width           =   2010
         Begin VB.TextBox TIPOCAM 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   720
            TabIndex        =   26
            Top             =   260
            Width           =   1180
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "$/u$s:"
            Height          =   225
            Index           =   5
            Left            =   105
            TabIndex        =   25
            Top             =   315
            Width           =   750
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   150
         TabIndex        =   18
         Top             =   470
         Width           =   800
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Créditos:"
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   23
            Top             =   1575
            Width           =   750
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Débitos:"
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   22
            Top             =   1260
            Width           =   750
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ventas:"
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   21
            Top             =   945
            Width           =   750
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Período:"
            Height          =   225
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   750
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cálculo:"
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   19
            Top             =   420
            Width           =   750
         End
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
         Height          =   280
         Left            =   3390
         TabIndex        =   7
         Top             =   420
         Width           =   175
      End
      Begin VB.CommandButton Command2 
         Height          =   705
         Left            =   3990
         Picture         =   "IGESVEN.frx":26FA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Re-Calcula Resumen Mensual de Ventas"
         Top             =   420
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   945
         TabIndex        =   17
         Top             =   1995
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   2
         Left            =   945
         TabIndex        =   16
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   1
         Left            =   945
         TabIndex        =   15
         Top             =   1365
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   0
         Left            =   3150
         TabIndex        =   14
         Top             =   1995
         Width           =   1485
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Total:"
         Height          =   225
         Index           =   4
         Left            =   2310
         TabIndex        =   13
         Top             =   2055
         Width           =   750
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   945
         TabIndex        =   5
         Top             =   840
         Width           =   2640
      End
      Begin VB.Label MESAN 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   945
         TabIndex        =   4
         ToolTipText     =   "Listado Analisis por Facturacion"
         Top             =   420
         Width           =   2460
      End
   End
   Begin VB.CommandButton command1 
      Caption         =   "Exit"
      Height          =   800
      Left            =   10185
      Picture         =   "IGESVEN.frx":2A04
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      Height          =   800
      Left            =   10185
      Picture         =   "IGESVEN.frx":2B4E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1205
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   800
      Left            =   10185
      Picture         =   "IGESVEN.frx":2E58
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprime Ficha"
      Top             =   2095
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   9620
      Picture         =   "IGESVEN.frx":34C2
      ToolTipText     =   "Doble Click para Configurar Modalidad de Cálculo"
      Top             =   7570
      Width           =   2010
   End
End
Attribute VB_Name = "IGESVEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESDEFE%, HASTAFE%
Dim CORRESP$, ABREMES$, CORRESQ$
Dim DESFE$, HASFE$
Dim CODIMES%
'
Dim GRCC%(32767)
Dim GRCV%(32767)
Dim COREGIS#(32767)
'
Dim MONECO%      ' 1=Pesos, 2=Dolar
'
Private Sub Command10_Click()
   '
   Call CARSELMESA
   For U& = 31 To 3 Step -1   ' ULTREG&("SELMESA") To 1 Step -1
     '
     MESA$ = REGLEIDO$("SELMESA", U&)
     MESB$ = MESANO$(MESA$, DES%, HAS%)
     TICAM$ = "1.00"
     If DES% > FECHANUM("30/11/01") Then TICAM$ = "1.40"
     If DES% > FECHANUM("31/12/01") Then TICAM$ = "1.80"
     If DES% > FECHANUM("31/01/02") Then TICAM$ = "2.00"
     If DES% > FECHANUM("28/02/02") Then TICAM$ = "2.40"
     If DES% > FECHANUM("31/03/02") Then TICAM$ = "3.00"
     '
     MESAN.Caption = TRIM$(MESB$)
     MESAN.Refresh
     Call ARMAMES
     Call RECURESMEN
     If TRIM$(Label1(0)) = "" Then
       TIPOCAM = TICAM$
       TIPOCAM.Refresh
       Call CALCURESU
     End If
     Call CIERRARCH("*.*")
     '
   Next U&
   '
End Sub

Private Sub Command11_Click()
   '
   GRUCLI% = 0
   If Option21.Value <> True Then
     For KKI% = 1 To List1.ListCount
       If List1.ListIndex = KKI% - 1 Then
         GRUCLI% = Val(Left$(List1.List(KKI% - 1), 5))
       End If
     Next KKI%
   End If
   '
   CAMESA% = 24
   For KKI% = 0 To 3
     If Option3(KKI%).Value = True Then
       CAMESA% = 12 * (1 + KKI%)
     End If
   Next KKI%
   '
   Call GRAFUNCICLI(GRUCLI%, CAMESA%, 1)
   '
End Sub

Private Sub Command14_Click()
   If Option9.Value = True Then
         Call TAVENTAPRO(1)
      ElseIf Option8.Value = True Then
         Call GRABARRAPRO(1)
      ElseIf Option7.Value = True Then
         Call GRAPORMABRUPRO(1)
      ElseIf Option6.Value = True Then
         Call GRAVAMABRUPRO(1)
      ElseIf Option5.Value = True Then
         Call GRAPARVENPRO(1)
   End If
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   Call RESUENCLI1
   Command15.Enabled = True
End Sub

Private Sub Command16_Click()
   '
   Command16.Enabled = False
   Call ARMAMES
   If DESDEFE% < 1 Then GoTo 99
   '
   If Val(TIPOCAM) < 0.8 Then
      Call MENSERR(24, "Falta Indicar Tipo de Cambio")
      On Error Resume Next
      TIPOCAM.SetFocus
      On Error GoTo 0
      GoTo 99
   End If
   '
   If COMALTER%("Confirme Refrescar Costos\de " + TRIM$(MESAN) + ".\\Cancelar\Aprobar Refrescar") <> 2 Then
     GoTo 99
   End If
   '
   Call BLANCOSINFO
   Call CALCURESU
   '
99 Command16.Enabled = True
   '
End Sub

Private Sub Command17_Click()
   Command17.Enabled = False
   Call REVEPRO1
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   Call RANKPIE1
   Command18.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   Call RANKCLI1
   Command19.Enabled = True
End Sub

Private Sub Command20_Click()
   Command20.Enabled = False
   MENUANA.Show 1
   Command20.Enabled = True
End Sub

Private Sub Command3_Click()
   PRINTFORM
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   Call ARMAMES
   If DESDEFE% < 1 Then GoTo 99
   '
   Call PRITAMERC
   Call PRITARES
99 Command4.Enabled = True
   '
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
   Call ARMAMES
   If DESDEFE% < 1 Then GoTo 99
   '
   If TRIM$(Label3) = "" Then
     Call MENSERR(24, "No hay Cálculo Primario\para el Mes Elegido.")
     GoTo 99
   End If
   '
   Screen.MousePointer = 11
   For KKI% = 1 To 32766
     COREGIS#(KKI%) = 0
   Next KKI%
   '
   RFF$ = RECFILT$("PVECLART", 1, MKI$(CODIMES%))
   For U& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, U&, 4))
     RELE$ = REGLEIDO$("PVECLART", REG&)
     CII% = CVI(Mid$(RELE$, 7, 2))
     If CII% > 0 Then
       If CII% <= NUMAS% Then
         COREGIS#(CII%) = (-1)
         COLEI# = CVD(Mid$(RELE$, 185, 8))
         If Abs(COLEI#) >= 0.00005 Then
           COREGIS#(CII%) = COLEI#
         End If
       End If
     End If
   Next U&
   '
   REBLA$ = REGBLAN("COSINFOR")
   JJ& = 0
   '
   FIN& = NUINV%
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     IK0$ = REGLEIDO$("INVERT", U& + 1)
     CII% = CVI(Mid$(IK0$, 17, 2))
     If CII% > 0 Then
       If CII% <= NUMAS% Then
         If Abs(COREGIS#(CII%)) >= 0.00005 Then
           COLEI# = COREGIS#(CII%) / FACONVERT(CODIMES%)
           If COLEI# < 0 Then COLEI# = 0
           TTXX$ = REBLA$
           Call REPLA(TTXX$, MKI$(CII%), 1, 2)
           Call REPLA(TTXX$, MKD$(COLEI#), 3, 8)
           JJ& = JJ& + 1
           Call GRAREG("!COSINFOR", TTXX$, JJ&)
         End If
       End If
     End If
   Next U&
   Call SETULTREG("!COSINFOR", JJ&)
   Call CIERRARCH("!COSINFOR")
   '
   TITUMO$ = "Pesos"
   If Option10.Value = True Then
      TITUMO$ = "Dólares U$S"
   End If
   '
   Screen.MousePointer = 1
   VTB% = VENTABU%("COSINFOR/TITUPAN=Planilla de Informe de Costo Unitario Materias Primas - " + TITUMO$)
   '
   If VTB% >= 0 Then
      Screen.MousePointer = 11
      For KKI% = 1 To 32766
        COREGIS#(KKI%) = 0
      Next KKI%
      '
      For U& = 1 To ULTREG&("!COSINFOR")
        TTXX$ = REGLEIDO$("!COSINFOR", U&)
        CII% = CVI(Left$(TTXX$, 2))
        If CII% > 0 Then
          If CII% <= NUMAS% Then
            COREGIS#(CII%) = CVD(Mid$(TTXX$, 3, 8)) * FACONVERT(CODIMES%)
          End If
        End If
      Next U&
      '
      For U& = 1 To Len(RFF$) Step 4
        REG& = CVS(Mid$(RFF$, U&, 4))
        RELE$ = REGLEIDO$("PVECLART", REG&)
        CII% = CVI(Mid$(RELE$, 7, 2))
        If CII% > 0 Then
          If CII% <= NUMAS% Then
            Call REPLA(RELE$, MKD$(COREGIS#(CII%)), 185, 8)
            Call GRAREG("PVECLART", RELE$, REG&)
          End If
        End If
      Next U&
      Call CIERRARCH("PVECLART")
      Call CALCURESU
      '
   End If
   '
99 Screen.MousePointer = 1
   Command6.Enabled = True
   '
End Sub

Private Sub Command9_Click()
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   ALTUPAPEL = Printer.ScaleHeight
   ANCHPAPEL = Printer.ScaleWidth
   '
   Call CARAPRIN
   Call CARAMERC
   '
   Call TAVECOSTO(2)
   Call PRINTABLA(FINTABLA)
   '
   Call GRAVECOSTO(2)
   GRABARRA.MSChart1.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   BIZQ = 0.1 * ANCHPAPEL
   ANCHO = 0.89 * ANCHPAPEL
   BSUP = FINTABLA + 0.03 * ALTUPAPEL
   ALTUR = 0.92 * (ALTUPAPEL - BSUP) / 2
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   '
   GRABARRA.MSChart2.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   BSUP = BSUP + ALTUR / 0.92
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   Printer.NewPage
   '
   Call GRAPORMABRUVE(2)
   GRABARRA.MSChart1.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   BIZQ = 0.1 * ANCHPAPEL
   ANCHO = 0.89 * ANCHPAPEL
   BSUP = 0.03 * ALTUPAPEL
   ALTUR = 0.9 * (ALTUPAPEL) / 4
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   '
   GRABARRA.MSChart2.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   BSUP = BSUP + 0.94 * ALTUPAPEL / 4
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   '
   Call GRAVAMABRUVE(2)
   GRABARRA.MSChart1.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   BSUP = BSUP + 0.94 * ALTUPAPEL / 4
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   '
   GRABARRA.MSChart2.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   BSUP = BSUP + 0.94 * ALTUPAPEL / 4
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   '
   Printer.NewPage
   '
   ' PARTICIPACION DE VENTAS POR MERCADO
5  Call GRAPARVENMERC(2)
   BIZQ = 0.1 * ANCHPAPEL
   ANCHO = 0.89 * ANCHPAPEL
   BSUP = 0.03 * ALTUPAPEL
   ALTUR = 0.9 * (ALTUPAPEL) / 2
   RELA = ANCHO / ALTUR
   RETANCH = GRATORTA.MSChart1.Width
   RETLEFT = GRATORTA.MSChart1.Left
   RETALTU = GRATORTA.MSChart1.Height
   GRATORTA.MSChart1.Left = 0
   ANCHOCAR = GRATORTA.MSChart1.Height * RELA: RELB = 1
   If ANCHOCAR > 11500 Then RELB = 11500 / ANCHOCAR
   GRATORTA.MSChart1.Height = GRATORTA.MSChart1.Height * RELB
   GRATORTA.MSChart1.Width = GRATORTA.MSChart1.Height * RELA
   GRATORTA.MSChart1.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   AAA = Printer.CurrentY
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   GRATORTA.MSChart1.Width = RETANCH
   GRATORTA.MSChart1.Left = RETLEFT
   GRATORTA.MSChart1.Height = RETALTU
   '
   RETANCH = GRATORTA.MSChart2.Width
   RETLEFT = GRATORTA.MSChart2.Left
   RETALTU = GRATORTA.MSChart2.Height
   BIZQ = 0.1 * ANCHPAPEL
   ANCHO = 0.89 * ANCHPAPEL
   BSUP = ALTUPAPEL / 2
   ALTUR = 0.9 * (ALTUPAPEL) / 2
   RELA = ANCHO / ALTUR
   GRATORTA.MSChart2.Left = 0
   ANCHOCAR = GRATORTA.MSChart2.Height * RELA: RELB = 1
   If ANCHOCAR > 11500 Then RELB = 11500 / ANCHOCAR
   GRATORTA.MSChart2.Height = GRATORTA.MSChart2.Height * RELB
   GRATORTA.MSChart2.Width = GRATORTA.MSChart2.Height * RELA
   GRATORTA.MSChart2.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   GRATORTA.MSChart2.Width = RETANCH
   GRATORTA.MSChart2.Left = RETLEFT
   GRATORTA.MSChart1.Height = RETALTU
   Printer.NewPage
   '
   ' SERIES HISTORICAS
10 CAMESA% = 24
   For KKI% = 0 To 3
     If Option3(KKI%).Value = True Then
       CAMESA% = 12 * (1 + KKI%)
     End If
   Next KKI%
   '
   IPAGI = 0
   For KKI% = 1 To List1.ListCount + 1
     If KKI% > List1.ListCount Then
         GRUCLI% = 0
       Else
         GRUCLI% = Val(Left$(List1.List(KKI% - 1), 5))
     End If
     '
     Call GRAFUNCICLI(GRUCLI%, CAMESA%, 2)
     GRAFUNCI.MSChart1.EditCopy
     FORMALTER.Image3.Picture = Clipboard.GetData
     BIZQ = 0.1 * ANCHPAPEL
     ANCHO = 0.88 * ANCHPAPEL
     BSUP = (0.03 + IPAGI / 2) * ALTUPAPEL
     ALTUR = 0.22 * ALTUPAPEL
     Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
     '
     GRAFUNCI.MSChart2.EditCopy
     FORMALTER.Image3.Picture = Clipboard.GetData
     BIZQ = 0.1 * ANCHPAPEL
     ANCHO = 0.88 * ANCHPAPEL
     BSUP = (0.27 + IPAGI / 2) * ALTUPAPEL
     ALTUR = 0.22 * ALTUPAPEL
     Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
     '
     IPAGI = IPAGI + 1
     If (IPAGI Mod 2) = 0 Then
       IPAGI = 0
       Printer.NewPage
     End If
   '
   Next KKI%
   If IPAGI <> 0 Then
     IPAGI = 0
     Printer.NewPage
   End If
   '
25 Call CARAPROD
   '
   Call TAVENTAPRO(2)
   Call PRINTABLA(FINTABLA)
   '
   Call GRABARRAPRO(2)
   GRABARRA.MSChart1.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   BIZQ = 0.1 * ANCHPAPEL
   ANCHO = 0.89 * ANCHPAPEL
   BSUP = FINTABLA + 0.03 * ALTUPAPEL
   ALTUR = 0.92 * (ALTUPAPEL - BSUP) / 2
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   '
   GRABARRA.MSChart2.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   BSUP = BSUP + ALTUR / 0.92
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   Printer.NewPage
   '
   Call GRAPORMABRUPRO(2)
   GRABARRA.MSChart1.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   BIZQ = 0.1 * ANCHPAPEL
   ANCHO = 0.89 * ANCHPAPEL
   BSUP = 0.03 * ALTUPAPEL
   ALTUR = 0.9 * (ALTUPAPEL) / 4
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   '
   GRABARRA.MSChart2.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   BSUP = BSUP + 0.94 * ALTUPAPEL / 4
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   '
   Call GRAVAMABRUPRO(2)
   GRABARRA.MSChart1.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   BSUP = BSUP + 0.94 * ALTUPAPEL / 4
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   '
   GRABARRA.MSChart2.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   BSUP = BSUP + 0.94 * ALTUPAPEL / 4
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   '
   Printer.NewPage
   '

   ' PARTICIPACION DE VENTAS POR PRODUCTO
35 Call GRAPARVENPRO(2)
   BIZQ = 0.1 * ANCHPAPEL
   ANCHO = 0.89 * ANCHPAPEL
   BSUP = 0.03 * ALTUPAPEL
   ALTUR = 0.9 * (ALTUPAPEL) / 2
   RELA = ANCHO / ALTUR
   RETANCH = GRATORTA.MSChart1.Width
   RETLEFT = GRATORTA.MSChart1.Left
   RETALTU = GRATORTA.MSChart1.Height
   GRATORTA.MSChart1.Left = 0
   ANCHOCAR = GRATORTA.MSChart1.Height * RELA: RELB = 1
   If ANCHOCAR > 11500 Then RELB = 11500 / ANCHOCAR
   GRATORTA.MSChart1.Height = GRATORTA.MSChart1.Height * RELB
   GRATORTA.MSChart1.Width = GRATORTA.MSChart1.Height * RELA
   GRATORTA.MSChart1.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   AAA = Printer.CurrentY
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   GRATORTA.MSChart1.Width = RETANCH
   GRATORTA.MSChart1.Left = RETLEFT
   GRATORTA.MSChart1.Height = RETALTU
   '
   RETANCH = GRATORTA.MSChart2.Width
   RETLEFT = GRATORTA.MSChart2.Left
   RETALTU = GRATORTA.MSChart2.Height
   BIZQ = 0.1 * ANCHPAPEL
   ANCHO = 0.89 * ANCHPAPEL
   BSUP = ALTUPAPEL / 2
   ALTUR = 0.9 * (ALTUPAPEL) / 2
   RELA = ANCHO / ALTUR
   GRATORTA.MSChart2.Left = 0
   ANCHOCAR = GRATORTA.MSChart2.Height * RELA: RELB = 1
   If ANCHOCAR > 11500 Then RELB = 11500 / ANCHOCAR
   GRATORTA.MSChart2.Height = GRATORTA.MSChart2.Height * RELB
   GRATORTA.MSChart2.Width = GRATORTA.MSChart2.Height * RELA
   GRATORTA.MSChart2.EditCopy
   FORMALTER.Image3.Picture = Clipboard.GetData
   Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
   GRATORTA.MSChart2.Width = RETANCH
   GRATORTA.MSChart2.Left = RETLEFT
   GRATORTA.MSChart1.Height = RETALTU
   Printer.NewPage
   '
   ' SERIES HISTORICAS
40 CAMESA% = 24
   For KKI% = 0 To 3
     If Option3(KKI%).Value = True Then
       CAMESA% = 12 * (1 + KKI%)
     End If
   Next KKI%
   '
   IPAGI = 0
   For KKI% = 1 To List2.ListCount + 1
     If KKI% > List2.ListCount Then
         GRUCLI% = 0
       Else
         GRUCLI% = Val(Left$(List2.List(KKI% - 1), 5))
     End If
     '
     Call GRAFUNCIPRO(GRUCLI%, CAMESA%, 2)
     GRAFUNPR.MSChart1.Height = 4000 * 0.36 / 0.26
     GRAFUNPR.MSChart1.Refresh
     GRAFUNPR.MSChart1.EditCopy
     FORMALTER.Image3.Picture = Clipboard.GetData
     BIZQ = 0.1 * ANCHPAPEL
     ANCHO = 0.89 * ANCHPAPEL
     BSUP = 0.03 * ALTUPAPEL
     ALTUR = 0.36 * ALTUPAPEL
     Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
     GRAFUNPR.MSChart1.Height = 4000
     GRAFUNPR.MSChart1.Refresh
     '
     If GRUCLI% <> 0 Then
         GRAFUNPR.MSChart2.Top = 2200
         GRAFUNPR.MSChart2.Height = 4000
         GRAFUNPR.MSChart2.Refresh
         GRAFUNPR.MSChart2.EditCopy
         FORMALTER.Image3.Picture = Clipboard.GetData
         BIZQ = 0.1 * ANCHPAPEL
         ANCHO = 0.89 * ANCHPAPEL
         BSUP = 0.41 * ALTUPAPEL
         ALTUR = 0.26 * ALTUPAPEL
         Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
         '
         GRAFUNPR.MSChart3.Top = 2200
         GRAFUNPR.MSChart3.Height = 4000
         GRAFUNPR.MSChart3.Refresh
         GRAFUNPR.MSChart3.EditCopy
         FORMALTER.Image3.Picture = Clipboard.GetData
         BIZQ = 0.1 * ANCHPAPEL
         ANCHO = 0.89 * ANCHPAPEL
         BSUP = 0.69 * ALTUPAPEL
         ALTUR = 0.26 * ALTUPAPEL
         Printer.PaintPicture FORMALTER.Image3.Picture, BIZQ, BSUP, ANCHO, ALTUR
         '
         GRAFUNPR.MSChart1.Top = 100
         GRAFUNPR.MSChart1.Height = 4000
         GRAFUNPR.MSChart1.Refresh
         GRAFUNPR.MSChart2.Top = 4200
         GRAFUNPR.MSChart2.Height = 2000
         GRAFUNPR.MSChart2.Refresh
         GRAFUNPR.MSChart3.Top = 6300
         GRAFUNPR.MSChart3.Height = 2000
         GRAFUNPR.MSChart3.Refresh
         '
     End If
     '
     If KKI% < List2.ListCount + 1 Then
       Printer.NewPage
     End If
     '
   Next KKI%
   Printer.EndDoc
   '
   Call PRITAMERC       ' Planilla de Division por Mercado
   Call PRITARES        ' Planilla de Division por Producto
   '
End Sub

Private Sub Command12_Click()
   '
   GRUPRO% = 0
   If Option4.Value <> True Then
     For KKI% = 1 To List2.ListCount
       If List2.ListIndex = KKI% - 1 Then
         GRUPRO% = Val(Left$(List2.List(KKI% - 1), 5))
       End If
     Next KKI%
   End If
   '
   CAMESA% = 24
   For KKI% = 0 To 3
     If Option1(KKI%).Value = True Then
       CAMESA% = 12 * (1 + KKI%)
     End If
   Next KKI%
   '
   Call GRAFUNCIPRO(GRUPRO%, CAMESA%, 1)
   '
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    '
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Show
    Refresh
    '
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("GRUCOCLI")
      X$ = REGLEIDO$("GRUCOCLI", U&)
      GRUCO% = CVI(Left$(X$, 2))
      DEGRUPO$ = TRIM$(Mid$(X$, 3, 30))
      TTXX$ = Space$(48)
      Call REPLA(TTXX$, FORMATNUM$(GRUCO%, "I5"), 1, 5)
      Call REPLA(TTXX$, DEGRUPO$, 9, 30)
      List1.AddItem TTXX$
    Next U&
    List1.ListIndex = (-1)
    '
    If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then
      List2.AddItem "    1    DE/DEC"
      List2.AddItem "    2    DEPP"
      List2.AddItem "    4    MILEN/DD"
      List2.AddItem "    5    DV/DVR"
      List2.AddItem "    6    DSH"
      List2.AddItem "    7    GAST/VACIO"
      List2.AddItem "    8    CHEM/PULSA"
      List2.AddItem "    9    REPUESTOS"
      List2.AddItem "   99    VARIOS"
      GoTo 15
    End If
    '
    For U& = 1 To ULTREG&("GRUCOVEN")
      X$ = REGLEIDO$("GRUCOVEN", U&)
      GRUCO% = CVI(Left$(X$, 2))
      DEGRUPO$ = TRIM$(Mid$(X$, 3, 30))
      TTXX$ = Space$(48)
      Call REPLA(TTXX$, FORMATNUM$(GRUCO%, "I5"), 1, 5)
      Call REPLA(TTXX$, DEGRUPO$, 9, 30)
      List2.AddItem TTXX$
    Next U&
15  List2.ListIndex = (-1)
    '
    FIN& = ULTREG&("DEUDOR1")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("DEUDOR1", U&) + REGLEIDO$("DEUDOR2", U&)
      NCLIE% = CVI(Left$(X$, 2))
      If NCLIE% > 0 Then
        GRCC%(NCLIE%) = CVI(Mid$(X$, 157, 2))
      End If
    Next U&
    Call CIERRARCH("DEUDOR1")
    Call CIERRARCH("DEUDOR2")
    '
    FIN& = NUMAS%
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("MASTER", U& + 1)
      CII% = CVI(Left$(X$, 2))
      If CII% > 0 Then
        GRCV%(CII%) = CVI(Mid$(X$, 125, 2))
        If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then
          GRCV%(CII%) = Int(GRCV%(CII%) / 10)
        End If
      End If
    Next U&
    Call CIERRARCH("MASTER")
    '
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command5_Click()
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
      MESAN.Refresh
      Call ARMAMES
      Call RECURESMEN
   End If
End Sub

Private Sub Image2_DblClick()
   OPCALCOS.Show 1
End Sub

Private Sub List1_Click()
   Option21.Value = False
End Sub

Private Sub List1_DblClick()
   '
   For KKI% = 1 To List1.ListCount
     If List1.ListIndex = KKI% - 1 Then
       GRUCLI% = Val(Left$(List1.List(KKI% - 1), 5))
     End If
   Next KKI%
   '
   CAMESA% = 24
   For KKI% = 0 To 3
     If Option3(KKI%).Value = True Then
       CAMESA% = 12 * (1 + KKI%)
     End If
   Next KKI%
   '
   Call GRAFUNCICLI(GRUCLI%, CAMESA%, 1)
   '
End Sub

Private Sub List2_Click()
   Option4.Value = False
End Sub

Private Sub List2_DblClick()
   '
   For KKI% = 1 To List2.ListCount
     If List2.ListIndex = KKI% - 1 Then
       GRUPRO% = Val(Left$(List2.List(KKI% - 1), 5))
     End If
   Next KKI%
   '
   CAMESA% = 24
   For KKI% = 0 To 3
     If Option1(KKI%).Value = True Then
       CAMESA% = 12 * (1 + KKI%)
     End If
   Next KKI%
   '
   Call GRAFUNCIPRO(GRUPRO%, CAMESA%, 1)
   '
End Sub

Private Sub MESAN_DblClick()
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
      MESAN.Refresh
      Call ARMAMES
      Call RECURESMEN
   End If
End Sub

Private Sub Command2_Click()
   '
   Command2.Enabled = False
   Call ARMAMES
   If DESDEFE% < 1 Then GoTo 99
   '
   If Val(TIPOCAM) < 0.8 Then
      Call MENSERR(24, "Falta Indicar Tipo de Cambio")
      On Error Resume Next
      TIPOCAM.SetFocus
      On Error GoTo 0
      GoTo 99
   End If
   '
   Call CALCURESU
99 Command2.Enabled = True
   '
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   If Option11.Value = True Then
        Call TAVECOSTO(1)
      ElseIf Option12.Value = True Then
        Call GRAVECOSTO(1)
      ElseIf Option13.Value = True Then
        Call GRAPORMABRUVE(1)
      ElseIf Option14.Value = True Then
        Call GRAVAMABRUVE(1)
      ElseIf Option15.Value = True Then
        Call GRAPARVENMERC(1)
   End If
   Command8.Enabled = True
End Sub

Sub GRABARRAPRO(modo%)
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), COSVEN(32767)
   Dim ACVEN(32767), ACCOS(32767)
   Dim NUGRU%(256), DEGRU$(256)
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   '
   If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then
      NUGRU%(1) = 1: DEGRU$(1) = "DE/DEC": VEGRUVE$ = MKI$(0) + MKI$(1)
      NUGRU%(2) = 2: DEGRU$(2) = "DEPP": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(2)
      NUGRU%(3) = 4: DEGRU$(3) = "MILEN/DD": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(4)
      NUGRU%(4) = 5: DEGRU$(4) = "DV/DVR": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(5)
      NUGRU%(5) = 6: DEGRU$(5) = "DSH": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(6)
      NUGRU%(6) = 7: DEGRU$(6) = "GAST/VACIO": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(7)
      NUGRU%(7) = 8: DEGRU$(7) = "CHEM/PULSA": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(8)
      NUGRU%(8) = 9: DEGRU$(8) = "REPUESTOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(9)
      NUGRU%(9) = 99: DEGRU$(9) = "VARIOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(99)
      CAGRU% = 9
      ULTGRU% = 99
      GoTo 99
   End If
   '
   For U& = 1 To ULTREG&("GRUCOVEN")
     X$ = REGLEIDO$("GRUCOVEN", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     NUGRU%(CAGRU%) = GRUCO%
     ULTGRU% = GRUCO%
   Next U&
   '
99 GRABARRA.MSChart1.RowCount = CAGRU%
   GRABARRA.MSChart2.RowCount = CAGRU%
   GRABARRA.MSChart1.Title = "VENTAS VS. COSTO MATERIA PRIMA POR PRODUCTO - " + CORRESP$
   GRABARRA.MSChart2.Title = "VENTAS VS. COSTO MATERIA PRIMA POR PRODUCTO - ACUMULADO A " + CORRESP$
   '
   MONECO% = 1
   GRABARRA.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "Pesos"
   GRABARRA.MSChart2.Plot.Axis(1).AxisTitle.TEXT = "Pesos"
   If Option10.Value = True Then
      MONECO% = 2
      GRABARRA.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "U$S"
      GRABARRA.MSChart2.Plot.Axis(1).AxisTitle.TEXT = "U$S"
   End If
   '
   GRABARRA.MSChart1.ColumnCount = 2
   GRABARRA.MSChart1.Column = 1
   GRABARRA.MSChart1.ColumnLabel = "Venta"
   GRABARRA.MSChart1.Column = 2
   GRABARRA.MSChart1.ColumnLabel = "Costo"
   GRABARRA.MSChart2.ColumnCount = 2
   GRABARRA.MSChart2.Column = 1
   GRABARRA.MSChart2.ColumnLabel = "Venta"
   GRABARRA.MSChart2.Column = 2
   GRABARRA.MSChart2.ColumnLabel = "Costo"
   '
   For KKI% = 1 To CAGRU%
     GRABARRA.MSChart1.Row = KKI%
     GRABARRA.MSChart1.RowLabel = DEGRU$(KKI%)
     GRABARRA.MSChart2.Row = KKI%
     GRABARRA.MSChart2.RowLabel = DEGRU$(KKI%)
   Next KKI%
   '
   GRABARRA.MSChart1.Refresh
   '
   TOTAVEN = 0: TOTACOS = 0
   ACUVEN = 0: ACUCOS = 0
   FIN& = ULTREG&("PVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("PVECLART", U&)
      CMES% = CVI(Left$(Z$, 2))
      If CMES% >= CODIMES% - 11 Then
        If CMES% <= CODIMES% Then
          CLI% = CVI(Mid$(Z$, 5, 2))
          CI% = CVI(Mid$(Z$, 7, 2))
          GRUCO% = GRCV%(CI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(CMES%)
          ACVEN(GRUCO%) = ACVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACCOS(GRUCO%) = ACCOS(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
          ACUVEN = ACUVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACUCOS = ACUCOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          '
          If CMES% = CODIMES% Then
            TOVEN(GRUCO%) = TOVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
            COSVEN(GRUCO%) = COSVEN(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
            TOTAVEN = TOTAVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
            TOTACOS = TOTACOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          End If
        End If
      End If
   Next U&
   '
   If TOTAVEN < 1 Then
      Screen.MousePointer = 1
      Call MENSERR(24, "Sin Movimiento")
      Exit Sub
   End If
   '
   GRUVARI% = ULTGRU%
   TOVEN(GRUVARI%) = TOTAVEN
   COSVEN(GRUVARI%) = TOTACOS
   ACVEN(GRUVARI%) = ACUVEN
   ACCOS(GRUVARI%) = ACUCOS
   For U& = 1 To ULTGRU% - 1
      TOVEN(GRUVARI%) = TOVEN(GRUVARI%) - TOVEN(U&)
      COSVEN(GRUVARI%) = COSVEN(GRUVARI%) - COSVEN(U&)
      ACVEN(GRUVARI%) = ACVEN(GRUVARI%) - ACVEN(U&)
      ACCOS(GRUVARI%) = ACCOS(GRUVARI%) - ACCOS(U&)
   Next U&
   '
   KKI% = 0
   For UI% = 1 To CAGRU%
     GRABARRA.MSChart1.Row = UI%
     GRABARRA.MSChart1.Column = 1
     GRABARRA.MSChart1.Data = TOVEN(NUGRU%(UI%))
     GRABARRA.MSChart1.Column = 2
     GRABARRA.MSChart1.Data = COSVEN(NUGRU%(UI%))
     '
     GRABARRA.MSChart2.Row = UI%
     GRABARRA.MSChart2.Column = 1
     GRABARRA.MSChart2.Data = ACVEN(NUGRU%(UI%))
     GRABARRA.MSChart2.Column = 2
     GRABARRA.MSChart2.Data = ACCOS(NUGRU%(UI%))
   Next UI%
   '
90 Screen.MousePointer = 1
   If modo% = 1 Then
      GRABARRA.Show 1
   End If
   '
End Sub

Sub GRATORTAPRO(modo%)
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), COSVEN(32767), PORVEN(32767)
   Dim NUGRU%(256), DEGRU$(256)
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   '
   If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then
      NUGRU%(1) = 1: DEGRU$(1) = "DE/DEC": VEGRUVE$ = MKI$(0) + MKI$(1)
      NUGRU%(2) = 2: DEGRU$(2) = "DEPP": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(2)
      NUGRU%(3) = 4: DEGRU$(3) = "MILEN/DD": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(4)
      NUGRU%(4) = 5: DEGRU$(4) = "DV/DVR": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(5)
      NUGRU%(5) = 6: DEGRU$(5) = "DSH": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(6)
      NUGRU%(6) = 7: DEGRU$(6) = "GAST/VACIO": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(7)
      NUGRU%(7) = 8: DEGRU$(7) = "CHEM/PULSA": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(8)
      NUGRU%(8) = 9: DEGRU$(8) = "REPUESTOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(9)
      NUGRU%(9) = 99: DEGRU$(9) = "VARIOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(99)
      CAGRU% = 9
      ULTGRU% = 99
      GoTo 15
   End If
   '
   For U& = 1 To ULTREG&("GRUCOVEN")
     X$ = REGLEIDO$("GRUCOVEN", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     NUGRU%(CAGRU%) = GRUCO%
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     ULTGRU% = GRUCO%
   Next U&
   '
15 GRATORTA.MSChart1.ColumnCount = CAGRU%
   GRATORTA.MSChart1.Title = "PARTICIPACION VENTAS POR PRODUCTO - " + CORRESP$
   For KKI% = 1 To CAGRU%
     GRATORTA.MSChart1.Column = KKI%
     GRATORTA.MSChart1.ColumnLabel = DEGRU$(KKI%)
   Next KKI%
   '
   TOTAVEN = 0: TOTACOS = 0
   FIN& = ULTREG&("!DVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("!DVECLART", U&)
      CLI% = CVI(Mid$(Z$, 5, 2))
      CI% = CVI(Mid$(Z$, 7, 2))
      GRUCO% = GRUCOVEI%(CI%)
      If GRUCO% < 1 Then GRUCO% = ULTGRU%
      If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
      TOVEN(GRUCO%) = TOVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8))
      COSVEN(GRUCO%) = COSVEN(GRUCO%) + CVD(Mid$(Z$, 193, 8))
      TOTAVEN = TOTAVEN + CVD(Mid$(Z$, 169, 8))
      TOTACOS = TOTACOS + CVD(Mid$(Z$, 193, 8))
   Next U&
   '
   If TOTAVEN < 1 Then
      Screen.MousePointer = 1
      Call MENSERR(24, "Sin Movimiento")
      Exit Sub
   End If
   '
   GRUVARI% = ULTGRU%
   TOVEN(GRUVARI%) = TOTAVEN
   COSVEN(GRUVARI%) = TOTACOS
   For U& = 1 To ULTGRU% - 1
      TOVEN(GRUVARI%) = TOVEN(GRUVARI%) - TOVEN(U&)
      COSVEN(GRUVARI%) = COSVEN(GRUVARI%) - COSVEN(U&)
   Next U&
   '
   For U& = 1 To ULTGRU%
      PORVEN(U&) = 100 * TOVEN(U&) / TOTAVEN
   Next U&
   '
   KKI% = 0
   For UI% = 1 To CAGRU%
     NGR% = NUGRU%(UI%)
     GRATORTA.MSChart1.Column = UI%
     GRATORTA.MSChart1.Row = 1
     GRATORTA.MSChart1.Data = TOVEN(NGR%)
     GRATORTA.MSChart1.ColumnLabel = TRIM$(DEGRU$(UI%)) + " " + TRIM$(CSTRING$(MKS$(PORVEN(NGR%)), 4, 5, 1, 2)) + " %"
   Next UI%
   '
   Screen.MousePointer = 1
   If modo% = 1 Then
      GRATORTA.Show 1
   End If
   '
End Sub

Sub TAVECOSTO(modo%)
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), COSVEN(32767), TOCAN(32767)
   Dim ACVEN(32767), ACCOS(32767), ACCAN(32767)
   Dim NUGRU%(256), DEGRU$(256)
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   For U& = 1 To ULTREG&("GRUCOCLI")
     X$ = REGLEIDO$("GRUCOCLI", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     NUGRU%(CAGRU%) = GRUCO%
     ULTGRU% = GRUCO%
   Next U&
   '
   MONECO% = 1
   TITUMO$ = "PESOS"
   If Option10.Value = True Then
      MONECO% = 2
      TITUMO$ = "U$S"
   End If
   '
   GRATABLA.Label1(0) = "COSTO POR TIPO DE VENTAS - " + CORRESP$ + " - " + TITUMO$
   GRATABLA.Label1(1) = "COSTO POR TIPO DE VENTAS - ACUMULADO ANUAL A " + CORRESP$ + " - " + TITUMO$
   For KKI% = 0 To 1
     TTXX$ = GRATABLA.Label2(KKI%)
     Call REPLA(TTXX$, "GRUPO CLIENTE", 1, 15)
     GRATABLA.Label2(KKI%) = TTXX$
   Next KKI%
   '
   TOTAVEN = 0: TOTACOS = 0: TOTACAN = 0
   ACUVEN = 0: ACUCOS = 0: ACUCAN = 0
   FIN& = ULTREG&("CVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("CVECLART", U&)
      CMES% = CVI(Left$(Z$, 2))
      If CMES% >= CODIMES% - 11 Then
        If CMES% <= CODIMES% Then
          CLI% = CVI(Mid$(Z$, 5, 2))
          CI% = CVI(Mid$(Z$, 7, 2))
          GRUCO% = GRCC%(CLI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(CMES%)
          ACVEN(GRUCO%) = ACVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACCOS(GRUCO%) = ACCOS(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
          ACCAN(GRUCO%) = ACCAN(GRUCO%) + CVD(Mid$(Z$, 51, 8))
          ACUVEN = ACUVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACUCOS = ACUCOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          ACUCAN = ACUCAN + CVD(Mid$(Z$, 51, 8))
          '
          If CMES% = CODIMES% Then
            TOVEN(GRUCO%) = TOVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
            COSVEN(GRUCO%) = COSVEN(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
            TOCAN(GRUCO%) = TOCAN(GRUCO%) + CVD(Mid$(Z$, 51, 8))
            TOTAVEN = TOTAVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
            TOTACOS = TOTACOS + CVD(Mid$(Z$, 193, 8)) / TICAM
            TOTACAN = TOTACAN + CVD(Mid$(Z$, 51, 8))
          End If
        End If
      End If
   Next U&
   '
   If TOTAVEN < 1 Then
      Screen.MousePointer = 1
      Call MENSERR(24, "Sin Movimiento")
      Exit Sub
   End If
   '
   GRUVARI% = ULTGRU%
   TOVEN(GRUVARI%) = TOTAVEN
   COSVEN(GRUVARI%) = TOTACOS
   TOCAN(GRUVARI%) = TOTACAN
   ACVEN(GRUVARI%) = ACUVEN
   ACCOS(GRUVARI%) = ACUCOS
   ACCAN(GRUVARI%) = ACUCAN
   For KKI% = 1 To CAGRU% - 1
     TOVEN(GRUVARI%) = TOVEN(GRUVARI%) - TOVEN(NUGRU%(KKI%))
     COSVEN(GRUVARI%) = COSVEN(GRUVARI%) - COSVEN(NUGRU%(KKI%))
     TOCAN(GRUVARI%) = TOCAN(GRUVARI%) - TOCAN(NUGRU%(KKI%))
     ACVEN(GRUVARI%) = ACVEN(GRUVARI%) - ACVEN(NUGRU%(KKI%))
     ACCOS(GRUVARI%) = ACCOS(GRUVARI%) - ACCOS(NUGRU%(KKI%))
     ACCAN(GRUVARI%) = ACCAN(GRUVARI%) - ACCAN(NUGRU%(KKI%))
   Next KKI%
   '
   POSY = 800
   GRATABLA.List1(0).Clear
   GRATABLA.List1(1).Clear
   For UI% = 1 To CAGRU%
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, DEGRU$(UI%), 1, 20)
     Call REPLA(TTXX$, FORMATNUM$(TOCAN(NUGRU%(UI%)), "F6.0"), 21, 6)
     Call REPLA(TTXX$, FORMATNUM$(100 * TOVEN(NUGRU%(UI%)) / TOTAVEN, "F8.1"), 27, 8)
     Call REPLA(TTXX$, FORMATNUM$(TOVEN(NUGRU%(UI%)), "F13.2"), 36, 13)
     Call REPLA(TTXX$, FORMATNUM$(COSVEN(NUGRU%(UI%)), "F13.2"), 49, 13)
     Call REPLA(TTXX$, FORMATNUM$(TOVEN(NUGRU%(UI%)) - COSVEN(NUGRU%(UI%)), "F13.2"), 62, 13)
     If TOVEN(NUGRU%(UI%)) > 1 Then
       Call REPLA(TTXX$, FORMATNUM$(100 * COSVEN(NUGRU%(UI%)) / TOVEN(NUGRU%(UI%)), "F8.1"), 75, 8)
       Call REPLA(TTXX$, FORMATNUM$(100 * (TOVEN(NUGRU%(UI%)) - COSVEN(NUGRU%(UI%))) / TOVEN(NUGRU%(UI%)), "F8.1"), 85, 8)
     End If
     GRATABLA.List1(0).AddItem TTXX$
     '
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, DEGRU$(UI%), 1, 20)
     Call REPLA(TTXX$, FORMATNUM$(ACCAN(NUGRU%(UI%)), "F6.0"), 21, 6)
     Call REPLA(TTXX$, FORMATNUM$(100 * ACVEN(NUGRU%(UI%)) / ACUVEN, "F8.1"), 27, 8)
     Call REPLA(TTXX$, FORMATNUM$(ACVEN(NUGRU%(UI%)), "F13.2"), 36, 13)
     Call REPLA(TTXX$, FORMATNUM$(ACCOS(NUGRU%(UI%)), "F13.2"), 49, 13)
     Call REPLA(TTXX$, FORMATNUM$(ACVEN(NUGRU%(UI%)) - ACCOS(NUGRU%(UI%)), "F13.2"), 62, 13)
     If ACVEN(NUGRU%(UI%)) > 1 Then
       Call REPLA(TTXX$, FORMATNUM$(100 * ACCOS(NUGRU%(UI%)) / ACVEN(NUGRU%(UI%)), "F8.1"), 75, 8)
       Call REPLA(TTXX$, FORMATNUM$(100 * (ACVEN(NUGRU%(UI%)) - ACCOS(NUGRU%(UI%))) / ACVEN(NUGRU%(UI%)), "F8.1"), 85, 8)
     End If
     GRATABLA.List1(1).AddItem TTXX$
   Next UI%
   '
   TTXX$ = Space$(96)
   Call REPLA(TTXX$, "Total", 1, 20)
   Call REPLA(TTXX$, FORMATNUM$(TOTACAN, "F6.0"), 21, 6)
   Call REPLA(TTXX$, FORMATNUM$(100, "F8.1"), 27, 8)
   Call REPLA(TTXX$, FORMATNUM$(TOTAVEN, "F13.2"), 36, 13)
   Call REPLA(TTXX$, FORMATNUM$(TOTACOS, "F13.2"), 49, 13)
   Call REPLA(TTXX$, FORMATNUM$(TOTAVEN - TOTACOS, "F13.2"), 62, 13)
   If TOTAVEN > 1 Then
     Call REPLA(TTXX$, FORMATNUM$(100 * TOTACOS / TOTAVEN, "F8.1"), 75, 8)
     Call REPLA(TTXX$, FORMATNUM$(100 * (TOTAVEN - TOTACOS) / TOTAVEN, "F8.1"), 85, 8)
   End If
   GRATABLA.Label3(0) = TTXX$
   '
   TTXX$ = Space$(96)
   Call REPLA(TTXX$, "Total", 1, 20)
   Call REPLA(TTXX$, FORMATNUM$(ACUCAN, "F6.0"), 21, 6)
   Call REPLA(TTXX$, FORMATNUM$(100, "F8.1"), 27, 8)
   Call REPLA(TTXX$, FORMATNUM$(ACUVEN, "F13.2"), 36, 13)
   Call REPLA(TTXX$, FORMATNUM$(ACUCOS, "F13.2"), 49, 13)
   Call REPLA(TTXX$, FORMATNUM$(ACUVEN - ACUCOS, "F13.2"), 62, 15)
   If TOTAVEN > 1 Then
     Call REPLA(TTXX$, FORMATNUM$(100 * ACUCOS / ACUVEN, "F8.1"), 75, 8)
     Call REPLA(TTXX$, FORMATNUM$(100 * (ACUVEN - ACUCOS) / ACUVEN, "F8.1"), 85, 8)
   End If
   GRATABLA.Label3(1) = TTXX$
   '
   Screen.MousePointer = 1
   If modo% = 1 Then
     ATENTIA% = ATENTI%
     GRATABLA.Show 1
     If ATENTI% <> ATENTIA% Then
       ATENTI% = ATENTIA%
       Call FRESHALL
       Call PRINTABLA(FINTABLA)
       Printer.EndDoc
       On Error Resume Next
       Unload GRATABLA
       On Error GoTo 0
     End If
   End If
   '
End Sub

Sub TAVENTAPRO(modo%)
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), COSVEN(32767), TOCAN(32767)
   Dim ACVEN(32767), ACCOS(32767), ACCAN(32767)
   Dim NUGRU%(256), DEGRU$(256)
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   '
   If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then
      NUGRU%(1) = 1: DEGRU$(1) = "DE/DEC": VEGRUVE$ = MKI$(0) + MKI$(1)
      NUGRU%(2) = 2: DEGRU$(2) = "DEPP": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(2)
      NUGRU%(3) = 4: DEGRU$(3) = "MILEN/DD": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(4)
      NUGRU%(4) = 5: DEGRU$(4) = "DV/DVR": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(5)
      NUGRU%(5) = 6: DEGRU$(5) = "DSH": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(6)
      NUGRU%(6) = 7: DEGRU$(6) = "GAST/VACIO": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(7)
      NUGRU%(7) = 8: DEGRU$(7) = "CHEM/PULSA": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(8)
      NUGRU%(8) = 9: DEGRU$(8) = "REPUESTOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(9)
      NUGRU%(9) = 99: DEGRU$(9) = "VARIOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(99)
      CAGRU% = 9
      ULTGRU% = 99
      GoTo 15
   End If
   '
   For U& = 1 To ULTREG&("GRUCOVEN")
     X$ = REGLEIDO$("GRUCOVEN", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     NUGRU%(CAGRU%) = GRUCO%
     ULTGRU% = GRUCO%
   Next U&
   '
15 MONECO% = 1
   TITUMO$ = "PESOS"
   If Option10.Value = True Then
      MONECO% = 2
      TITUMO$ = "U$S"
   End If
   '
   GRATABLA.Label1(0) = "COSTO POR TIPO DE VENTAS - " + CORRESP$ + " - " + TITUMO$
   GRATABLA.Label1(1) = "COSTO POR TIPO DE VENTAS - ACUMULADO ANUAL A " + CORRESP$ + " - " + TITUMO$
   For KKI% = 0 To 1
     TTXX$ = GRATABLA.Label2(KKI%)
     Call REPLA(TTXX$, "GRUPO PRODUCTO", 1, 15)
     GRATABLA.Label2(KKI%) = TTXX$
   Next KKI%
   '
   TOTAVEN = 0: TOTACOS = 0: TOTACAN = 0
   ACUVEN = 0: ACUCOS = 0: ACUCAN = 0
   FIN& = ULTREG&("PVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("PVECLART", U&)
      CMES% = CVI(Left$(Z$, 2))
      If CMES% >= CODIMES% - 11 Then
        If CMES% <= CODIMES% Then
          CI% = CVI(Mid$(Z$, 7, 2))
          GRUCO% = GRCV%(CI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(CMES%)
          ACVEN(GRUCO%) = ACVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACCOS(GRUCO%) = ACCOS(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
          ACCAN(GRUCO%) = ACCAN(GRUCO%) + CVD(Mid$(Z$, 51, 8))
          ACUVEN = ACUVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACUCOS = ACUCOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          ACUCAN = ACUCAN + CVD(Mid$(Z$, 51, 8))
          '
          If CMES% = CODIMES% Then
            TOVEN(GRUCO%) = TOVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
            COSVEN(GRUCO%) = COSVEN(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
            TOCAN(GRUCO%) = TOCAN(GRUCO%) + CVD(Mid$(Z$, 51, 8))
            TOTAVEN = TOTAVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
            TOTACOS = TOTACOS + CVD(Mid$(Z$, 193, 8)) / TICAM
            TOTACAN = TOTACAN + CVD(Mid$(Z$, 51, 8))
          End If
        End If
      End If
   Next U&
   '
   If TOTAVEN < 1 Then
      Screen.MousePointer = 1
      Call MENSERR(24, "Sin Movimiento")
      Exit Sub
   End If
   '
   GRUVARI% = ULTGRU%
   TOVEN(GRUVARI%) = TOTAVEN
   COSVEN(GRUVARI%) = TOTACOS
   TOCAN(GRUVARI%) = TOTACAN
   ACVEN(GRUVARI%) = ACUVEN
   ACCOS(GRUVARI%) = ACUCOS
   ACCAN(GRUVARI%) = ACUCAN
   For KKI% = 1 To CAGRU% - 1
     TOVEN(GRUVARI%) = TOVEN(GRUVARI%) - TOVEN(NUGRU%(KKI%))
     COSVEN(GRUVARI%) = COSVEN(GRUVARI%) - COSVEN(NUGRU%(KKI%))
     TOCAN(GRUVARI%) = TOCAN(GRUVARI%) - TOCAN(NUGRU%(KKI%))
     ACVEN(GRUVARI%) = ACVEN(GRUVARI%) - ACVEN(NUGRU%(KKI%))
     ACCOS(GRUVARI%) = ACCOS(GRUVARI%) - ACCOS(NUGRU%(KKI%))
     ACCAN(GRUVARI%) = ACCAN(GRUVARI%) - ACCAN(NUGRU%(KKI%))
   Next KKI%
   '
   POSY = 800
   GRATABLA.List1(0).Clear
   GRATABLA.List1(1).Clear
   For UI% = 1 To CAGRU%
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, DEGRU$(UI%), 1, 20)
     Call REPLA(TTXX$, FORMATNUM$(TOCAN(NUGRU%(UI%)), "F6.0"), 21, 6)
     Call REPLA(TTXX$, FORMATNUM$(100 * TOVEN(NUGRU%(UI%)) / TOTAVEN, "F8.1"), 27, 8)
     Call REPLA(TTXX$, FORMATNUM$(TOVEN(NUGRU%(UI%)), "F13.2"), 36, 13)
     Call REPLA(TTXX$, FORMATNUM$(COSVEN(NUGRU%(UI%)), "F13.2"), 49, 13)
     Call REPLA(TTXX$, FORMATNUM$(TOVEN(NUGRU%(UI%)) - COSVEN(NUGRU%(UI%)), "F13.2"), 62, 13)
     If TOVEN(NUGRU%(UI%)) > 1 Then
       Call REPLA(TTXX$, FORMATNUM$(100 * COSVEN(NUGRU%(UI%)) / TOVEN(NUGRU%(UI%)), "F8.1"), 75, 8)
       Call REPLA(TTXX$, FORMATNUM$(100 * (TOVEN(NUGRU%(UI%)) - COSVEN(NUGRU%(UI%))) / TOVEN(NUGRU%(UI%)), "F8.1"), 85, 8)
     End If
     GRATABLA.List1(0).AddItem TTXX$
     '
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, DEGRU$(UI%), 1, 20)
     Call REPLA(TTXX$, FORMATNUM$(ACCAN(NUGRU%(UI%)), "F6.0"), 21, 6)
     Call REPLA(TTXX$, FORMATNUM$(100 * ACVEN(NUGRU%(UI%)) / ACUVEN, "F8.1"), 27, 8)
     Call REPLA(TTXX$, FORMATNUM$(ACVEN(NUGRU%(UI%)), "F13.2"), 36, 13)
     Call REPLA(TTXX$, FORMATNUM$(ACCOS(NUGRU%(UI%)), "F13.2"), 49, 13)
     Call REPLA(TTXX$, FORMATNUM$(ACVEN(NUGRU%(UI%)) - ACCOS(NUGRU%(UI%)), "F13.2"), 62, 13)
     If ACVEN(NUGRU%(UI%)) > 1 Then
       Call REPLA(TTXX$, FORMATNUM$(100 * ACCOS(NUGRU%(UI%)) / ACVEN(NUGRU%(UI%)), "F8.1"), 75, 8)
       Call REPLA(TTXX$, FORMATNUM$(100 * (ACVEN(NUGRU%(UI%)) - ACCOS(NUGRU%(UI%))) / ACVEN(NUGRU%(UI%)), "F8.1"), 85, 8)
     End If
     GRATABLA.List1(1).AddItem TTXX$
   Next UI%
   '
   TTXX$ = Space$(96)
   Call REPLA(TTXX$, "Total", 1, 20)
   Call REPLA(TTXX$, FORMATNUM$(TOTACAN, "F6.0"), 21, 6)
   Call REPLA(TTXX$, FORMATNUM$(100, "F8.1"), 27, 8)
   Call REPLA(TTXX$, FORMATNUM$(TOTAVEN, "F13.2"), 36, 13)
   Call REPLA(TTXX$, FORMATNUM$(TOTACOS, "F13.2"), 49, 13)
   Call REPLA(TTXX$, FORMATNUM$(TOTAVEN - TOTACOS, "F13.2"), 62, 13)
   If TOTAVEN > 1 Then
     Call REPLA(TTXX$, FORMATNUM$(100 * TOTACOS / TOTAVEN, "F8.1"), 75, 8)
     Call REPLA(TTXX$, FORMATNUM$(100 * (TOTAVEN - TOTACOS) / TOTAVEN, "F8.1"), 85, 8)
   End If
   GRATABLA.Label3(0) = TTXX$
   '
   TTXX$ = Space$(96)
   Call REPLA(TTXX$, "Total", 1, 20)
   Call REPLA(TTXX$, FORMATNUM$(ACUCAN, "F6.0"), 21, 6)
   Call REPLA(TTXX$, FORMATNUM$(100, "F8.1"), 27, 8)
   Call REPLA(TTXX$, FORMATNUM$(ACUVEN, "F13.2"), 36, 13)
   Call REPLA(TTXX$, FORMATNUM$(ACUCOS, "F13.2"), 49, 13)
   Call REPLA(TTXX$, FORMATNUM$(ACUVEN - ACUCOS, "F13.2"), 62, 15)
   If TOTAVEN > 1 Then
     Call REPLA(TTXX$, FORMATNUM$(100 * ACUCOS / ACUVEN, "F8.1"), 75, 8)
     Call REPLA(TTXX$, FORMATNUM$(100 * (ACUVEN - ACUCOS) / ACUVEN, "F8.1"), 85, 8)
   End If
   GRATABLA.Label3(1) = TTXX$
   '
   Screen.MousePointer = 1
   If modo% = 1 Then
     ATENTIA% = ATENTI%
     GRATABLA.Show 1
     If ATENTI% <> ATENTIA% Then
       ATENTI% = ATENTIA%
       Call FRESHALL
       Call PRINTABLA(FINTABLA)
       Printer.EndDoc
       On Error Resume Next
       Unload GRATABLA
       On Error GoTo 0
     End If
   End If
   '
End Sub

Sub GRAVECOSTO(modo%)
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), COSVEN(32767)
   Dim ACVEN(32767), ACCOS(32767)
   Dim NUGRU%(256), DEGRU$(256)
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   For U& = 1 To ULTREG&("GRUCOCLI")
     X$ = REGLEIDO$("GRUCOCLI", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     NUGRU%(CAGRU%) = GRUCO%
     ULTGRU% = GRUCO%
   Next U&
   '
   GRABARRA.MSChart1.RowCount = CAGRU%
   GRABARRA.MSChart2.RowCount = CAGRU%
   GRABARRA.MSChart1.Title = "VENTAS VS. COSTO MATERIA PRIMA POR MERCADO - " + CORRESP$
   GRABARRA.MSChart2.Title = "VENTAS VS. COSTO MATERIA PRIMA POR MERCADO - ACUMULADO A " + CORRESP$
   '
   MONECO% = 1
   GRABARRA.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "Pesos"
   GRABARRA.MSChart2.Plot.Axis(1).AxisTitle.TEXT = "Pesos"
   If Option10.Value = True Then
      MONECO% = 2
      GRABARRA.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "U$S"
      GRABARRA.MSChart2.Plot.Axis(1).AxisTitle.TEXT = "U$S"
   End If
   '
   For KKI% = 1 To CAGRU%
     GRABARRA.MSChart1.Row = KKI%
     GRABARRA.MSChart1.RowLabel = DEGRU$(KKI%)
     GRABARRA.MSChart2.Row = KKI%
     GRABARRA.MSChart2.RowLabel = DEGRU$(KKI%)
   Next KKI%
   '
   TOTAVEN = 0: TOTACOS = 0
   ACUVEN = 0: ACUCOS = 0
   FIN& = ULTREG&("CVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("CVECLART", U&)
      CMES% = CVI(Left$(Z$, 2))
      If CMES% >= CODIMES% - 11 Then
        If CMES% <= CODIMES% Then
          CLI% = CVI(Mid$(Z$, 5, 2))
          CI% = CVI(Mid$(Z$, 7, 2))
          GRUCO% = GRCC%(CLI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(CMES%)
          ACVEN(GRUCO%) = ACVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACCOS(GRUCO%) = ACCOS(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
          ACUVEN = ACUVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACUCOS = ACUCOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          '
          If CMES% = CODIMES% Then
            TOVEN(GRUCO%) = TOVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
            COSVEN(GRUCO%) = COSVEN(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
            TOTAVEN = TOTAVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
            TOTACOS = TOTACOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          End If
        End If
      End If
   Next U&
   '
   If TOTAVEN < 1 Then
      Screen.MousePointer = 1
      Call MENSERR(24, "Sin Movimiento")
      Exit Sub
   End If
   '
   GRUVARI% = ULTGRU%
   TOVEN(GRUVARI%) = TOTAVEN
   COSVEN(GRUVARI%) = TOTACOS
   ACVEN(GRUVARI%) = ACUVEN
   ACCOS(GRUVARI%) = ACUCOS
   For KKI% = 1 To CAGRU% - 1
     TOVEN(GRUVARI%) = TOVEN(GRUVARI%) - TOVEN(NUGRU%(KKI%))
     COSVEN(GRUVARI%) = COSVEN(GRUVARI%) - COSVEN(NUGRU%(KKI%))
     ACVEN(GRUVARI%) = ACVEN(GRUVARI%) - ACVEN(NUGRU%(KKI%))
     ACCOS(GRUVARI%) = ACCOS(GRUVARI%) - ACCOS(NUGRU%(KKI%))
   Next KKI%
   '
   KKI% = 0
   For UI% = 1 To CAGRU%
     GRABARRA.MSChart1.Row = UI%
     GRABARRA.MSChart1.Column = 1
     GRABARRA.MSChart1.Data = TOVEN(NUGRU%(UI%))
     GRABARRA.MSChart1.Column = 2
     GRABARRA.MSChart1.Data = COSVEN(NUGRU%(UI%))
     '
     GRABARRA.MSChart2.Row = UI%
     GRABARRA.MSChart2.Column = 1
     GRABARRA.MSChart2.Data = ACVEN(NUGRU%(UI%))
     GRABARRA.MSChart2.Column = 2
     GRABARRA.MSChart2.Data = ACCOS(NUGRU%(UI%))
   Next UI%
   '
   Screen.MousePointer = 1
   If modo% = 1 Then
     GRABARRA.Show 1
   End If
   '
End Sub

Sub GRAPORMABRUVE(modo%)
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), COSVEN(32767)
   Dim ACVEN(32767), ACCOS(32767)
   Dim NUGRU%(256), DEGRU$(256)
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   For U& = 1 To ULTREG&("GRUCOCLI")
     X$ = REGLEIDO$("GRUCOCLI", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     NUGRU%(CAGRU%) = GRUCO%
     ULTGRU% = GRUCO%
   Next U&
   '
   GRABARRA.MSChart1.RowCount = CAGRU%
   GRABARRA.MSChart2.RowCount = CAGRU%
   GRABARRA.MSChart1.ColumnCount = 1
   GRABARRA.MSChart1.Column = 1
   GRABARRA.MSChart1.ColumnLabel = "% M.Bruto"
   GRABARRA.MSChart2.ColumnCount = 1
   GRABARRA.MSChart2.Column = 1
   GRABARRA.MSChart2.ColumnLabel = "% M.Bruto"
   GRABARRA.MSChart1.Title = "PORCENTAJE DE MARGEN BRUTO / VENTAS POR MERCADO - " + CORRESP$
   GRABARRA.MSChart2.Title = "PORCENTAJE DE MARGEN BRUTO / VENTAS POR MERCADO - ACUMULADO A " + CORRESP$
   GRABARRA.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "%"
   GRABARRA.MSChart2.Plot.Axis(1).AxisTitle.TEXT = "%"
   '
   For KKI% = 1 To CAGRU%
     GRABARRA.MSChart1.Row = KKI%
     GRABARRA.MSChart1.RowLabel = DEGRU$(KKI%)
     GRABARRA.MSChart2.Row = KKI%
     GRABARRA.MSChart2.RowLabel = DEGRU$(KKI%)
   Next KKI%
   '
   TOTAVEN = 0: TOTACOS = 0
   ACUVEN = 0: ACUCOS = 0
   FIN& = ULTREG&("CVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("CVECLART", U&)
      CMES% = CVI(Left$(Z$, 2))
      If CMES% >= CODIMES% - 11 Then
        If CMES% <= CODIMES% Then
          CLI% = CVI(Mid$(Z$, 5, 2))
          CI% = CVI(Mid$(Z$, 7, 2))
          GRUCO% = GRCC%(CLI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(CMES%)
          ACVEN(GRUCO%) = ACVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACCOS(GRUCO%) = ACCOS(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
          ACUVEN = ACUVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACUCOS = ACUCOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          '
          If CMES% = CODIMES% Then
            TOVEN(GRUCO%) = TOVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
            COSVEN(GRUCO%) = COSVEN(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
            TOTAVEN = TOTAVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
            TOTACOS = TOTACOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          End If
        End If
      End If
   Next U&
   '
   If TOTAVEN < 1 Then
      Screen.MousePointer = 1
      Call MENSERR(24, "Sin Movimiento")
      Exit Sub
   End If
   '
   GRUVARI% = ULTGRU%
   TOVEN(GRUVARI%) = TOTAVEN
   COSVEN(GRUVARI%) = TOTACOS
   ACVEN(GRUVARI%) = ACUVEN
   ACCOS(GRUVARI%) = ACUCOS
   For KKI% = 1 To CAGRU% - 1
     TOVEN(GRUVARI%) = TOVEN(GRUVARI%) - TOVEN(NUGRU%(KKI%))
     COSVEN(GRUVARI%) = COSVEN(GRUVARI%) - COSVEN(NUGRU%(KKI%))
     ACVEN(GRUVARI%) = ACVEN(GRUVARI%) - ACVEN(NUGRU%(KKI%))
     ACCOS(GRUVARI%) = ACCOS(GRUVARI%) - ACCOS(NUGRU%(KKI%))
   Next KKI%
   '
   KKI% = 0
   For UI% = 1 To CAGRU%
     PORMA = 0
     If TOVEN(NUGRU%(UI%)) > 0 Then
       PORMA = 100 * (TOVEN(NUGRU%(UI%)) - COSVEN(NUGRU%(UI%))) / TOVEN(NUGRU%(UI%))
     End If
     GRABARRA.MSChart1.Row = UI%
     GRABARRA.MSChart1.Column = 1
     GRABARRA.MSChart1.Data = PORMA
     '
     PORMA = 0
     If ACVEN(NUGRU%(UI%)) > 0 Then
       PORMA = 100 * (ACVEN(NUGRU%(UI%)) - ACCOS(NUGRU%(UI%))) / ACVEN(NUGRU%(UI%))
     End If
     GRABARRA.MSChart2.Row = UI%
     GRABARRA.MSChart2.Column = 1
     GRABARRA.MSChart2.Data = PORMA
   Next UI%
   '
   Screen.MousePointer = 1
   If modo% = 1 Then
     GRABARRA.Show 1
   End If
   '
End Sub

Sub GRAPORMABRUPRO(modo%)
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), COSVEN(32767)
   Dim ACVEN(32767), ACCOS(32767)
   Dim NUGRU%(256), DEGRU$(256)
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   '
   If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then
      NUGRU%(1) = 1: DEGRU$(1) = "DE/DEC": VEGRUVE$ = MKI$(0) + MKI$(1)
      NUGRU%(2) = 2: DEGRU$(2) = "DEPP": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(2)
      NUGRU%(3) = 4: DEGRU$(3) = "MILEN/DD": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(4)
      NUGRU%(4) = 5: DEGRU$(4) = "DV/DVR": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(5)
      NUGRU%(5) = 6: DEGRU$(5) = "DSH": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(6)
      NUGRU%(6) = 7: DEGRU$(6) = "GAST/VACIO": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(7)
      NUGRU%(7) = 8: DEGRU$(7) = "CHEM/PULSA": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(8)
      NUGRU%(8) = 9: DEGRU$(8) = "REPUESTOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(9)
      NUGRU%(9) = 99: DEGRU$(9) = "VARIOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(99)
      CAGRU% = 9
      ULTGRU% = 99
      GoTo 15
   End If
   '
   For U& = 1 To ULTREG&("GRUCOVEN")
     X$ = REGLEIDO$("GRUCOVEN", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     NUGRU%(CAGRU%) = GRUCO%
     ULTGRU% = GRUCO%
   Next U&
   '
15 GRABARRA.MSChart1.RowCount = CAGRU%
   GRABARRA.MSChart2.RowCount = CAGRU%
   GRABARRA.MSChart1.ColumnCount = 1
   GRABARRA.MSChart1.Column = 1
   GRABARRA.MSChart1.ColumnLabel = "% M.Bruto"
   GRABARRA.MSChart2.ColumnCount = 1
   GRABARRA.MSChart2.Column = 1
   GRABARRA.MSChart2.ColumnLabel = "% M.Bruto"
   GRABARRA.MSChart1.Title = "PORCENTAJE DE MARGEN BRUTO / VENTAS POR PRODUCTO - " + CORRESP$
   GRABARRA.MSChart2.Title = "PORCENTAJE DE MARGEN BRUTO / VENTAS POR PRODUCTO - ACUMULADO A " + CORRESP$
   GRABARRA.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "%"
   GRABARRA.MSChart2.Plot.Axis(1).AxisTitle.TEXT = "%"
   '
   For KKI% = 1 To CAGRU%
     GRABARRA.MSChart1.Row = KKI%
     GRABARRA.MSChart1.RowLabel = DEGRU$(KKI%)
     GRABARRA.MSChart2.Row = KKI%
     GRABARRA.MSChart2.RowLabel = DEGRU$(KKI%)
   Next KKI%
   '
   TOTAVEN = 0: TOTACOS = 0
   ACUVEN = 0: ACUCOS = 0
   FIN& = ULTREG&("PVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("PVECLART", U&)
      CMES% = CVI(Left$(Z$, 2))
      If CMES% >= CODIMES% - 11 Then
        If CMES% <= CODIMES% Then
          CI% = CVI(Mid$(Z$, 7, 2))
          GRUCO% = GRCV%(CI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(CMES%)
          ACVEN(GRUCO%) = ACVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACCOS(GRUCO%) = ACCOS(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
          ACUVEN = ACUVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACUCOS = ACUCOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          '
          If CMES% = CODIMES% Then
            TOVEN(GRUCO%) = TOVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
            COSVEN(GRUCO%) = COSVEN(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
            TOTAVEN = TOTAVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
            TOTACOS = TOTACOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          End If
        End If
      End If
   Next U&
   '
   If TOTAVEN < 1 Then
      Screen.MousePointer = 1
      Call MENSERR(24, "Sin Movimiento")
      Exit Sub
   End If
   '
   GRUVARI% = ULTGRU%
   TOVEN(GRUVARI%) = TOTAVEN
   COSVEN(GRUVARI%) = TOTACOS
   ACVEN(GRUVARI%) = ACUVEN
   ACCOS(GRUVARI%) = ACUCOS
   For U& = 1 To ULTGRU% - 1
      TOVEN(GRUVARI%) = TOVEN(GRUVARI%) - TOVEN(U&)
      COSVEN(GRUVARI%) = COSVEN(GRUVARI%) - COSVEN(U&)
      ACVEN(GRUVARI%) = ACVEN(GRUVARI%) - ACVEN(U&)
      ACCOS(GRUVARI%) = ACCOS(GRUVARI%) - ACCOS(U&)
   Next U&
   '
   KKI% = 0
   For UI% = 1 To CAGRU%
     PORMA = 0
     If TOVEN(NUGRU%(UI%)) > 0 Then
       PORMA = 100 * (TOVEN(NUGRU%(UI%)) - COSVEN(NUGRU%(UI%))) / TOVEN(NUGRU%(UI%))
     End If
     GRABARRA.MSChart1.Row = UI%
     GRABARRA.MSChart1.Column = 1
     GRABARRA.MSChart1.Data = PORMA
     '
     PORMA = 0
     If ACVEN(NUGRU%(UI%)) > 0 Then
       PORMA = 100 * (ACVEN(NUGRU%(UI%)) - ACCOS(NUGRU%(UI%))) / ACVEN(NUGRU%(UI%))
     End If
     GRABARRA.MSChart2.Row = UI%
     GRABARRA.MSChart2.Column = 1
     GRABARRA.MSChart2.Data = PORMA
   Next UI%
   '
   Screen.MousePointer = 1
   If modo% = 1 Then
      GRABARRA.Show 1
   End If
   '
End Sub

Sub GRAVAMABRUVE(modo%)
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), COSVEN(32767)
   Dim ACVEN(32767), ACCOS(32767)
   Dim NUGRU%(256), DEGRU$(256)
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   For U& = 1 To ULTREG&("GRUCOCLI")
     X$ = REGLEIDO$("GRUCOCLI", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     NUGRU%(CAGRU%) = GRUCO%
     ULTGRU% = GRUCO%
   Next U&
   '
   GRABARRA.MSChart1.RowCount = CAGRU%
   GRABARRA.MSChart2.RowCount = CAGRU%
   GRABARRA.MSChart1.ColumnCount = 1
   GRABARRA.MSChart1.Column = 1
   GRABARRA.MSChart1.ColumnLabel = "M.Bruto"
   GRABARRA.MSChart2.ColumnCount = 1
   GRABARRA.MSChart2.Column = 1
   GRABARRA.MSChart2.ColumnLabel = "M.Bruto"
   GRABARRA.MSChart1.Title = "VALOR ABSOLUTO MARGEN BRUTO / VENTAS POR MERCADO - " + CORRESP$
   GRABARRA.MSChart2.Title = "VALOR ABSOLUTO MARGEN BRUTO / VENTAS POR MERCADO - ACUMULADO A " + CORRESP$
   For KKI% = 1 To CAGRU%
     GRABARRA.MSChart1.Row = KKI%
     GRABARRA.MSChart1.RowLabel = DEGRU$(KKI%)
     GRABARRA.MSChart2.Row = KKI%
     GRABARRA.MSChart2.RowLabel = DEGRU$(KKI%)
   Next KKI%
   '
   MONECO% = 1
   GRABARRA.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "Pesos"
   GRABARRA.MSChart2.Plot.Axis(1).AxisTitle.TEXT = "Pesos"
   If Option10.Value = True Then
      MONECO% = 2
      GRABARRA.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "U$S"
      GRABARRA.MSChart2.Plot.Axis(1).AxisTitle.TEXT = "U$S"
   End If
   '
   TOTAVEN = 0: TOTACOS = 0
   ACUVEN = 0: ACUCOS = 0
   FIN& = ULTREG&("CVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("CVECLART", U&)
      CMES% = CVI(Left$(Z$, 2))
      If CMES% >= CODIMES% - 11 Then
        If CMES% <= CODIMES% Then
          CLI% = CVI(Mid$(Z$, 5, 2))
          CI% = CVI(Mid$(Z$, 7, 2))
          GRUCO% = GRCC%(CLI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(CMES%)
          ACVEN(GRUCO%) = ACVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACCOS(GRUCO%) = ACCOS(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
          ACUVEN = ACUVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACUCOS = ACUCOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          '
          If CMES% = CODIMES% Then
            TOVEN(GRUCO%) = TOVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
            COSVEN(GRUCO%) = COSVEN(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
            TOTAVEN = TOTAVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
            TOTACOS = TOTACOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          End If
        End If
      End If
   Next U&
   '
   If TOTAVEN < 1 Then
      Screen.MousePointer = 1
      Call MENSERR(24, "Sin Movimiento")
      Exit Sub
   End If
   '
   GRUVARI% = ULTGRU%
   TOVEN(GRUVARI%) = TOTAVEN
   COSVEN(GRUVARI%) = TOTACOS
   ACVEN(GRUVARI%) = ACUVEN
   ACCOS(GRUVARI%) = ACUCOS
   For KKI% = 1 To CAGRU% - 1
     TOVEN(GRUVARI%) = TOVEN(GRUVARI%) - TOVEN(NUGRU%(KKI%))
     COSVEN(GRUVARI%) = COSVEN(GRUVARI%) - COSVEN(NUGRU%(KKI%))
     ACVEN(GRUVARI%) = ACVEN(GRUVARI%) - ACVEN(NUGRU%(KKI%))
     ACCOS(GRUVARI%) = ACCOS(GRUVARI%) - ACCOS(NUGRU%(KKI%))
   Next KKI%
   '
   KKI% = 0
   For UI% = 1 To CAGRU%
     VALMA = TOVEN(NUGRU%(UI%)) - COSVEN(NUGRU%(UI%))
     GRABARRA.MSChart1.Row = UI%
     GRABARRA.MSChart1.Column = 1
     GRABARRA.MSChart1.Data = VALMA
     '
     VALMA = ACVEN(NUGRU%(UI%)) - ACCOS(NUGRU%(UI%))
     GRABARRA.MSChart2.Row = UI%
     GRABARRA.MSChart2.Column = 1
     GRABARRA.MSChart2.Data = VALMA
   Next UI%
   '
   Screen.MousePointer = 1
   If modo% = 1 Then
     GRABARRA.Show 1
   End If
   '
End Sub

Sub GRAVAMABRUPRO(modo%)
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), COSVEN(32767)
   Dim ACVEN(32767), ACCOS(32767)
   Dim NUGRU%(256), DEGRU$(256)
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   '
   If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then
      NUGRU%(1) = 1: DEGRU$(1) = "DE/DEC": VEGRUVE$ = MKI$(0) + MKI$(1)
      NUGRU%(2) = 2: DEGRU$(2) = "DEPP": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(2)
      NUGRU%(3) = 4: DEGRU$(3) = "MILEN/DD": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(4)
      NUGRU%(4) = 5: DEGRU$(4) = "DV/DVR": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(5)
      NUGRU%(5) = 6: DEGRU$(5) = "DSH": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(6)
      NUGRU%(6) = 7: DEGRU$(6) = "GAST/VACIO": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(7)
      NUGRU%(7) = 8: DEGRU$(7) = "CHEM/PULSA": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(8)
      NUGRU%(8) = 9: DEGRU$(8) = "REPUESTOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(9)
      NUGRU%(9) = 99: DEGRU$(9) = "VARIOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(99)
      CAGRU% = 9
      ULTGRU% = 99
      GoTo 15
   End If
   '
   For U& = 1 To ULTREG&("GRUCOVEN")
     X$ = REGLEIDO$("GRUCOVEN", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     NUGRU%(CAGRU%) = GRUCO%
     ULTGRU% = GRUCO%
   Next U&
   '
15 GRABARRA.MSChart1.RowCount = CAGRU%
   GRABARRA.MSChart2.RowCount = CAGRU%
   GRABARRA.MSChart1.ColumnCount = 1
   GRABARRA.MSChart1.Column = 1
   GRABARRA.MSChart1.ColumnLabel = "M.Bruto"
   GRABARRA.MSChart2.ColumnCount = 1
   GRABARRA.MSChart2.Column = 1
   GRABARRA.MSChart2.ColumnLabel = "M.Bruto"
   GRABARRA.MSChart1.Title = "VALOR ABSOLUTO MARGEN BRUTO POR PRODUCTO - " + CORRESP$
   GRABARRA.MSChart2.Title = "VALOR ABSOLUTO MARGEN BRUTO POR PRODUCTO - ACUMULADO A " + CORRESP$
   For KKI% = 1 To CAGRU%
     GRABARRA.MSChart1.Row = KKI%
     GRABARRA.MSChart1.RowLabel = DEGRU$(KKI%)
     GRABARRA.MSChart2.Row = KKI%
     GRABARRA.MSChart2.RowLabel = DEGRU$(KKI%)
   Next KKI%
   '
   MONECO% = 1
   GRABARRA.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "Pesos"
   GRABARRA.MSChart2.Plot.Axis(1).AxisTitle.TEXT = "Pesos"
   If Option10.Value = True Then
      MONECO% = 2
      GRABARRA.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "U$S"
      GRABARRA.MSChart2.Plot.Axis(1).AxisTitle.TEXT = "U$S"
   End If
   '
   TOTAVEN = 0: TOTACOS = 0
   ACUVEN = 0: ACUCOS = 0
   FIN& = ULTREG&("PVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("PVECLART", U&)
      CMES% = CVI(Left$(Z$, 2))
      If CMES% >= CODIMES% - 11 Then
        If CMES% <= CODIMES% Then
          CI% = CVI(Mid$(Z$, 7, 2))
          GRUCO% = GRCV%(CI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(CMES%)
          ACVEN(GRUCO%) = ACVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACCOS(GRUCO%) = ACCOS(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
          ACUVEN = ACUVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
          ACUCOS = ACUCOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          '
          If CMES% = CODIMES% Then
            TOVEN(GRUCO%) = TOVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
            COSVEN(GRUCO%) = COSVEN(GRUCO%) + CVD(Mid$(Z$, 193, 8)) / TICAM
            TOTAVEN = TOTAVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
            TOTACOS = TOTACOS + CVD(Mid$(Z$, 193, 8)) / TICAM
          End If
        End If
      End If
   Next U&
   '
   If TOTAVEN < 1 Then
      Screen.MousePointer = 1
      Call MENSERR(24, "Sin Movimiento")
      Exit Sub
   End If
   '
   GRUVARI% = ULTGRU%
   TOVEN(GRUVARI%) = TOTAVEN
   COSVEN(GRUVARI%) = TOTACOS
   ACVEN(GRUVARI%) = ACUVEN
   ACCOS(GRUVARI%) = ACUCOS
   For U& = 1 To ULTGRU% - 1
      TOVEN(GRUVARI%) = TOVEN(GRUVARI%) - TOVEN(U&)
      COSVEN(GRUVARI%) = COSVEN(GRUVARI%) - COSVEN(U&)
      ACVEN(GRUVARI%) = ACVEN(GRUVARI%) - ACVEN(U&)
      ACCOS(GRUVARI%) = ACCOS(GRUVARI%) - ACCOS(U&)
   Next U&
   '
   KKI% = 0
   For UI% = 1 To CAGRU%
     VALMA = TOVEN(NUGRU%(UI%)) - COSVEN(NUGRU%(UI%))
     GRABARRA.MSChart1.Row = UI%
     GRABARRA.MSChart1.Column = 1
     GRABARRA.MSChart1.Data = VALMA
     '
     VALMA = ACVEN(NUGRU%(UI%)) - ACCOS(NUGRU%(UI%))
     GRABARRA.MSChart2.Row = UI%
     GRABARRA.MSChart2.Column = 1
     GRABARRA.MSChart2.Data = VALMA
   Next UI%
   '
   Screen.MousePointer = 1
   If modo% = 1 Then
      GRABARRA.Show 1
   End If
   '
End Sub

Sub GRAPARVENMERC(modo%)
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), ACUVEN(32767), PORVEN(32767), PORACU(32767)
   Dim NUGRU%(256), DEGRU$(256)
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   For U& = 1 To ULTREG&("GRUCOCLI")
     X$ = REGLEIDO$("GRUCOCLI", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     NUGRU%(CAGRU%) = GRUCO%
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     ULTGRU% = GRUCO%
   Next U&
   '
   GRATORTA.MSChart1.ColumnCount = CAGRU%
   GRATORTA.MSChart2.ColumnCount = CAGRU%
   GRATORTA.MSChart1.Title = "PARTICIP.VENTAS POR MERCADO - " + CORRESP$
   GRATORTA.MSChart2.Title = "VENTAS POR MERCADO - ACUM." + CORRESP$
   For KKI% = 1 To CAGRU%
     GRATORTA.MSChart1.Column = KKI%
     GRATORTA.MSChart1.ColumnLabel = DEGRU$(KKI%)
     GRATORTA.MSChart2.Column = KKI%
     GRATORTA.MSChart2.ColumnLabel = DEGRU$(KKI%)
   Next KKI%
   '
   TOTAVEN = 0: TOTACUM = 0
   FIN& = ULTREG&("CVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("CVECLART", U&)
      COPERIO% = CVI(Left$(Z$, 2))
      If COPERIO% > CODIMES% - 12 Then
        If COPERIO% <= CODIMES% Then
          CLI% = CVI(Mid$(Z$, 5, 2))
          CI% = CVI(Mid$(Z$, 7, 2))
          GRUCO% = GRCC%(CLI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(COPERIO%)
          ACUVEN(GRUCO%) = ACUVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          TOTACUM = TOTACUM + CVD(Mid$(Z$, 169, 8)) / TICAM
          If COPERIO% = CODIMES% Then
            TOVEN(GRUCO%) = TOVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM / TICAM
            TOTAVEN = TOTAVEN + CVD(Mid$(Z$, 169, 8)) / TICAM / TICAM
          End If
        End If
      End If
   Next U&
   '
   If TOTAVEN < 1 Then
      Screen.MousePointer = 1
      Call MENSERR(24, "Sin Movimiento")
      Exit Sub
   End If
   '
   GRUVARI% = ULTGRU%
   TOVEN(GRUVARI%) = TOTAVEN
   ACUVEN(GRUVARI%) = TOTACUM
   For U& = 1 To CAGRU% - 1
      NGR% = NUGRU%(U&)
      TOVEN(GRUVARI%) = TOVEN(GRUVARI%) - TOVEN(NGR%)
      ACUVEN(GRUVARI%) = ACUVEN(GRUVARI%) - ACUVEN(NGR%)
   Next U&
   '
   KKI% = 0
   For UI% = 1 To CAGRU%
     NGR% = NUGRU%(UI%)
     GRATORTA.MSChart1.Column = UI%
     GRATORTA.MSChart1.Row = 1
     GRATORTA.MSChart1.Data = TOVEN(NGR%)
     PORVE = 100 * TOVEN(NGR%) / TOTAVEN
     GRATORTA.MSChart1.ColumnLabel = TRIM$(DEGRU$(UI%)) + " " + TRIM$(CSTRING$(MKS$(PORVE), 4, 5, 1, 2)) + " %"
     '
     GRATORTA.MSChart2.Column = UI%
     GRATORTA.MSChart2.Row = 1
     GRATORTA.MSChart2.Data = ACUVEN(NGR%)
     PORAC = 100 * ACUVEN(NGR%) / TOTACUM
     GRATORTA.MSChart2.ColumnLabel = TRIM$(DEGRU$(UI%)) + " " + TRIM$(CSTRING$(MKS$(PORAC), 4, 5, 1, 2)) + " %"
   Next UI%
   '
   Screen.MousePointer = 1
   If modo% = 1 Then
     GRATORTA.Show 1
   End If
   '
End Sub

Sub GRAPARVENPRO(modo%)
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), ACUVEN(32767), PORVEN(32767), PORACU(32767)
   Dim NUGRU%(256), DEGRU$(256)
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   '
   If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then
      NUGRU%(1) = 1: DEGRU$(1) = "DE/DEC": VEGRUVE$ = MKI$(0) + MKI$(1)
      NUGRU%(2) = 2: DEGRU$(2) = "DEPP": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(2)
      NUGRU%(3) = 4: DEGRU$(3) = "MILEN/DD": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(4)
      NUGRU%(4) = 5: DEGRU$(4) = "DV/DVR": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(5)
      NUGRU%(5) = 6: DEGRU$(5) = "DSH": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(6)
      NUGRU%(6) = 7: DEGRU$(6) = "GAST/VACIO": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(7)
      NUGRU%(7) = 8: DEGRU$(7) = "CHEM/PULSA": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(8)
      NUGRU%(8) = 9: DEGRU$(8) = "REPUESTOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(9)
      NUGRU%(9) = 99: DEGRU$(9) = "VARIOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(99)
      CAGRU% = 9
      ULTGRU% = 99
      GoTo 15
   End If
   '
   For U& = 1 To ULTREG&("GRUCOVEN")
     X$ = REGLEIDO$("GRUCOVEN", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     NUGRU%(CAGRU%) = GRUCO%
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     ULTGRU% = GRUCO%
   Next U&
   '
15 GRATORTA.MSChart1.ColumnCount = CAGRU%
   GRATORTA.MSChart2.ColumnCount = CAGRU%
   GRATORTA.MSChart1.Title = "PARTICIP.VENTAS POR PRODUCTO - " + CORRESP$
   GRATORTA.MSChart2.Title = "VENTAS POR PRODUCTO - ACUM." + CORRESP$
   For KKI% = 1 To CAGRU%
     GRATORTA.MSChart1.Column = KKI%
     GRATORTA.MSChart1.ColumnLabel = DEGRU$(KKI%)
     GRATORTA.MSChart2.Column = KKI%
     GRATORTA.MSChart2.ColumnLabel = DEGRU$(KKI%)
   Next KKI%
   '
   TOTAVEN = 0: TOTACUM = 0
   FIN& = ULTREG&("PVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("PVECLART", U&)
      COPERIO% = CVI(Left$(Z$, 2))
      If COPERIO% > CODIMES% - 12 Then
        If COPERIO% <= CODIMES% Then
          CI% = CVI(Mid$(Z$, 7, 2))
          GRUCO% = GRCV%(CI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(COPERIO%)
          ACUVEN(GRUCO%) = ACUVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          TOTACUM = TOTACUM + CVD(Mid$(Z$, 169, 8)) / TICAM
          If COPERIO% = CODIMES% Then
            TOVEN(GRUCO%) = TOVEN(GRUCO%) + CVD(Mid$(Z$, 169, 8)) / TICAM
            TOTAVEN = TOTAVEN + CVD(Mid$(Z$, 169, 8)) / TICAM
          End If
        End If
      End If
   Next U&
   '
   If TOTAVEN < 1 Then
      Screen.MousePointer = 1
      Call MENSERR(24, "Sin Movimiento")
      Exit Sub
   End If
   '
   GRUVARI% = ULTGRU%
   TOVEN(GRUVARI%) = TOTAVEN
   ACUVEN(GRUVARI%) = TOTACUM
   For U& = 1 To CAGRU% - 1
      NGR% = NUGRU%(U&)
      TOVEN(GRUVARI%) = TOVEN(GRUVARI%) - TOVEN(NGR%)
      ACUVEN(GRUVARI%) = ACUVEN(GRUVARI%) - ACUVEN(NGR%)
   Next U&
   '
   KKI% = 0
   For UI% = 1 To CAGRU%
     NGR% = NUGRU%(UI%)
     GRATORTA.MSChart1.Column = UI%
     GRATORTA.MSChart1.Row = 1
     GRATORTA.MSChart1.Data = TOVEN(NGR%)
     PORVE = 100 * TOVEN(NGR%) / TOTAVEN
     GRATORTA.MSChart1.ColumnLabel = TRIM$(DEGRU$(UI%)) + " " + TRIM$(CSTRING$(MKS$(PORVE), 4, 5, 1, 2)) + " %"
     '
     GRATORTA.MSChart2.Column = UI%
     GRATORTA.MSChart2.Row = 1
     GRATORTA.MSChart2.Data = ACUVEN(NGR%)
     PORAC = 100 * ACUVEN(NGR%) / TOTACUM
     GRATORTA.MSChart2.ColumnLabel = TRIM$(DEGRU$(UI%)) + " " + TRIM$(CSTRING$(MKS$(PORAC), 4, 5, 1, 2)) + " %"
   Next UI%
   '
   Screen.MousePointer = 1
   If modo% = 1 Then
      GRATORTA.Show 1
   End If
   '
End Sub

Sub GRAFUNCICLI(GRUCLI%, CAMESA%, modo%)
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), COSVEN(32767)
   Dim NUGRU%(256), DEGRU$(256)
   Dim ORGRU%(32767)
   Dim IVEN(256, 64), ICOS(256, 64)
   '
   HACODI% = CODIMES%
   DECODI% = HACODI% - CAMESA% + 1
   '
   GRAFUNCI.MSChart1.Repaint = True
   GRAFUNCI.MSChart2.Repaint = True
   GRAFUNCI.MSChart1.RowCount = HACODI% - DECODI% + 1
   GRAFUNCI.MSChart2.RowCount = HACODI% - DECODI% + 1
   For KKI% = DECODI% To HACODI%
     GRAFUNCI.MSChart1.Row = 1 + KKI% - DECODI%
     GRAFUNCI.MSChart1.RowLabel = ABREMESAN$(KKI%, 0)
     GRAFUNCI.MSChart2.Row = 1 + KKI% - DECODI%
     GRAFUNCI.MSChart2.RowLabel = ABREMESAN$(KKI%, 0)
   Next KKI%
   '
   GRAFUNCI.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "Pesos"
   If Option10.Value = True Then
     GRAFUNCI.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "U$S"
   End If
   GRAFUNCI.MSChart2.Plot.Axis(1).AxisTitle.TEXT = "%"
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   For U& = 1 To ULTREG&("GRUCOCLI")
     X$ = REGLEIDO$("GRUCOCLI", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     NUGRU%(CAGRU%) = GRUCO%
     ORGRU%(GRUCO%) = CAGRU%
     ULTGRU% = GRUCO%
   Next U&
   '
   GRAFUNCI.MSChart2.ColumnCount = 1
   If GRUCLI% > 0 Then
       GRAFUNCI.MSChart1.ColumnCount = 1
       GRAFUNCI.MSChart1.Title = "MERCADO DE " + DEGRU$(ORGRU%(GRUCLI%))
       GRAFUNCI.MSChart1.Column = 1
       GRAFUNCI.MSChart1.ColumnLabel = "Ventas"
     ElseIf GRUCLI% = 0 Then
       GRAFUNCI.MSChart1.ColumnCount = CAGRU% + 1
       GRAFUNCI.MSChart1.Title = "ANALISIS POR MERCADO"
       For KKI% = 1 To CAGRU%
         GRAFUNCI.MSChart1.Column = KKI%
         GRAFUNCI.MSChart1.ColumnLabel = DEGRU$(KKI%)
       Next KKI%
   End If
   '
   If GRUCLI% = 0 Then
     GRAFUNCI.MSChart1.Column = CAGRU% + 1
     GRAFUNCI.MSChart1.ColumnLabel = "Total"
   End If
   '
   TOTAVEN = 0: TOTACOS = 0
   FIN& = ULTREG&("CVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("CVECLART", U&)
      CODIPE% = CVI(Left$(Z$, 2))
      If CODIPE% >= DECODI% Then
        If CODIPE% <= HACODI% Then
          CLI% = CVI(Mid$(Z$, 5, 2))
          CI% = CVI(Mid$(Z$, 7, 2))
          GRUCO% = GRCC%(CLI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(CODIPE%)
          '
          If GRUCLI% > 0 Then
            If GRUCO% = GRUCLI% Then
              IVEN(1, 1 + CODIPE% - DECODI%) = IVEN(1, 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 169, 8)) / TICAM
              ICOS(1, 1 + CODIPE% - DECODI%) = ICOS(1, 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 193, 8)) / TICAM
            End If
            GoTo 19
          End If
          '
          IVEN(ORGRU%(GRUCO%), 1 + CODIPE% - DECODI%) = IVEN(ORGRU%(GRUCO%), 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          IVEN(0, 1 + CODIPE% - DECODI%) = IVEN(0, 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          ICOS(ORGRU%(GRUCO%), 1 + CODIPE% - DECODI%) = ICOS(ORGRU%(GRUCO%), 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 193, 8)) / TICAM
          ICOS(0, 1 + CODIPE% - DECODI%) = ICOS(0, 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 193, 8)) / TICAM
          '
        End If
      End If
19 Next U&
   '
   If GRUCLI% > 0 Then
     GRAFUNCI.MSChart1.Column = 1
     GRAFUNCI.MSChart2.Column = 1
     For UJ% = 1 To HACODI% - DECODI% + 1
       GRAFUNCI.MSChart1.Row = UJ%
       GRAFUNCI.MSChart2.Row = UJ%
       GRAFUNCI.MSChart1.Data = IVEN(1, UJ%)
       GRAFUNCI.MSChart2.Data = 0
       If IVEN(1, UJ%) > 0 Then
         GRAFUNCI.MSChart2.Data = 100 * (IVEN(1, UJ%) - ICOS(1, UJ%)) / IVEN(1, UJ%)
       End If
     Next UJ%
     GoTo 90
   End If
   '
   For UI% = 1 To CAGRU% + 1
     GRAFUNCI.MSChart1.Column = UI%
     For UJ% = 1 To HACODI% - DECODI% + 1
       GRAFUNCI.MSChart1.Row = UJ%
       GRAFUNCI.MSChart2.Row = UJ%
       GRAFUNCI.MSChart2.Data = 0
       If UI% <= CAGRU% Then
           GRAFUNCI.MSChart1.Data = IVEN(UI%, UJ%)
         Else
           GRAFUNCI.MSChart2.Column = 1
           GRAFUNCI.MSChart1.Data = IVEN(0, UJ%)
           If IVEN(0, UJ%) > 0 Then
             GRAFUNCI.MSChart2.Data = 100 * (IVEN(0, UJ%) - ICOS(0, UJ%)) / IVEN(0, UJ%)
           End If
       End If
     Next UJ%
   Next UI%
   '
90 Screen.MousePointer = 1
   If modo% = 1 Then
     GRAFUNCI.Show 1
   End If
   '
End Sub

Sub GRAFUNCIPRO(GRUCLI%, CAMESA%, modo%)
   '
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), COSVEN(32767)
   Dim NUGRU%(256), DEGRU$(256)
   Dim ORGRU%(32767)
   Dim IVEN(256, 64), ICOS(256, 64), CANT(256, 64)
   '
   HACODI% = CODIMES%
   DECODI% = HACODI% - CAMESA% + 1
   '
   If GRUCLI% = 0 Then
      GRAFUNPR.MSChart2.Visible = False
      GRAFUNPR.MSChart3.Visible = False
      GRAFUNPR.MSChart1.Height = 8100
      GRAFUNPR.MSChart1.ShowLegend = True
   End If
   '
   GRAFUNPR.MSChart1.Repaint = True
   GRAFUNPR.MSChart1.RowCount = HACODI% - DECODI% + 1
   GRAFUNPR.MSChart2.RowCount = HACODI% - DECODI% + 1
   GRAFUNPR.MSChart3.RowCount = HACODI% - DECODI% + 1
   For KKI% = DECODI% To HACODI%
     GRAFUNPR.MSChart1.Row = 1 + KKI% - DECODI%
     GRAFUNPR.MSChart1.RowLabel = ABREMESAN$(KKI%, 0)
     GRAFUNPR.MSChart2.Row = 1 + KKI% - DECODI%
     GRAFUNPR.MSChart2.RowLabel = ABREMESAN$(KKI%, 0)
     GRAFUNPR.MSChart3.Row = 1 + KKI% - DECODI%
     GRAFUNPR.MSChart3.RowLabel = ABREMESAN$(KKI%, 0)
   Next KKI%
   '
   GRAFUNPR.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "Pesos"
   If Option10.Value = True Then
     GRAFUNPR.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "U$S"
   End If
   GRAFUNPR.MSChart2.Plot.Axis(1).AxisTitle.TEXT = "%"
   GRAFUNPR.MSChart3.Plot.Axis(1).AxisTitle.TEXT = "Unidad"
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   '
   If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then
      NUGRU%(1) = 1: ORGRU%(1) = 1: DEGRU$(1) = "DE/DEC": VEGRUVE$ = MKI$(0) + MKI$(1)
      NUGRU%(2) = 2: ORGRU%(2) = 2: DEGRU$(2) = "DEPP": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(2)
      NUGRU%(3) = 4: ORGRU%(4) = 3: DEGRU$(3) = "MILEN/DD": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(4)
      NUGRU%(4) = 5: ORGRU%(5) = 4: DEGRU$(4) = "DV/DVR": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(5)
      NUGRU%(5) = 6: ORGRU%(6) = 5: DEGRU$(5) = "DSH": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(6)
      NUGRU%(6) = 7: ORGRU%(7) = 6: DEGRU$(6) = "GAST/VACIO": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(7)
      NUGRU%(7) = 8: ORGRU%(8) = 7: DEGRU$(7) = "CHEM/PULSA": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(8)
      NUGRU%(8) = 9: ORGRU%(9) = 8: DEGRU$(8) = "REPUESTOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(9)
      NUGRU%(9) = 99: ORGRU%(99) = 9: DEGRU$(9) = "VARIOS": VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(99)
      CAGRU% = 9
      ULTGRU% = 99
      GoTo 15
   End If
   '
   For U& = 1 To ULTREG&("GRUCOVEN")
     X$ = REGLEIDO$("GRUCOVEN", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     NUGRU%(CAGRU%) = GRUCO%
     ORGRU%(GRUCO%) = CAGRU%
     ULTGRU% = GRUCO%
   Next U&
   '
15 If GRUCLI% > 0 Then
       GRAFUNPR.MSChart1.Title = "PRODUCTOS " + DEGRU$(ORGRU%(GRUCLI%))
       GRAFUNPR.MSChart1.ColumnCount = 1
       GRAFUNPR.MSChart1.Column = 1
       GRAFUNPR.MSChart1.ColumnLabel = "Ventas"
       GRAFUNPR.MSChart2.ColumnCount = 1
       GRAFUNPR.MSChart2.Column = 1
       GRAFUNPR.MSChart2.ColumnLabel = "M.Bruto"
       GRAFUNPR.MSChart3.ColumnCount = 1
       GRAFUNPR.MSChart3.Column = 1
       GRAFUNPR.MSChart3.ColumnLabel = "Cantidad"
     ElseIf GRUCLI% = 0 Then
       GRAFUNPR.MSChart1.ColumnCount = CAGRU% + 1
       GRAFUNPR.MSChart1.Title = "ANALISIS POR PRODUCTOS"
       For KKI% = 1 To CAGRU%
         GRAFUNPR.MSChart1.Column = KKI%
         GRAFUNPR.MSChart1.ColumnLabel = DEGRU$(KKI%)
       Next KKI%
   End If
   '
   If GRUCLI% = 0 Then
     GRAFUNPR.MSChart1.Column = CAGRU% + 1
     GRAFUNPR.MSChart1.ColumnLabel = "Total"
   End If
   '
   TOTAVEN = 0: TOTACOS = 0
   FIN& = ULTREG&("PVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("PVECLART", U&)
      CODIPE% = CVI(Left$(Z$, 2))
      If CODIPE% >= DECODI% Then
        If CODIPE% <= HACODI% Then
          CI% = CVI(Mid$(Z$, 7, 2))
          GRUCO% = GRCV%(CI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(CODIPE%)
          '
          If GRUCLI% > 0 Then
            If GRUCO% = GRUCLI% Then
              IVEN(1, 1 + CODIPE% - DECODI%) = IVEN(1, 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 169, 8)) / TICAM
              ICOS(1, 1 + CODIPE% - DECODI%) = ICOS(1, 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 193, 8)) / TICAM
              CANT(1, 1 + CODIPE% - DECODI%) = CANT(1, 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 51, 8))
            End If
            GoTo 19
          End If
          '
          IVEN(ORGRU%(GRUCO%), 1 + CODIPE% - DECODI%) = IVEN(ORGRU%(GRUCO%), 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          ICOS(ORGRU%(GRUCO%), 1 + CODIPE% - DECODI%) = ICOS(ORGRU%(GRUCO%), 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 193, 8)) / TICAM
          CANT(ORGRU%(GRUCO%), 1 + CODIPE% - DECODI%) = CANT(ORGRU%(GRUCO%), 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 51, 8))
          IVEN(0, 1 + CODIPE% - DECODI%) = IVEN(0, 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 169, 8)) / TICAM
          ICOS(0, 1 + CODIPE% - DECODI%) = ICOS(0, 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 193, 8)) / TICAM
          CANT(0, 1 + CODIPE% - DECODI%) = CANT(0, 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 51, 8))
          '
        End If
      End If
19 Next U&
   '
   If GRUCLI% > 0 Then
     GRAFUNPR.MSChart1.Column = 1
     For UJ% = 1 To HACODI% - DECODI% + 1
       GRAFUNPR.MSChart1.Row = UJ%
       GRAFUNPR.MSChart1.Data = IVEN(1, UJ%)
       GRAFUNPR.MSChart2.Row = UJ%
       GRAFUNPR.MSChart2.Data = 0
       If IVEN(1, UJ%) > 0 Then
         GRAFUNPR.MSChart2.Data = 100 * (IVEN(1, UJ%) - ICOS(1, UJ%)) / IVEN(1, UJ%)
       End If
       GRAFUNPR.MSChart3.Row = UJ%
       GRAFUNPR.MSChart3.Data = CANT(1, UJ%)
     Next UJ%
     GoTo 90
   End If
   '
   For UI% = 1 To CAGRU% + 1
     GRAFUNPR.MSChart1.Column = UI%
     For UJ% = 1 To HACODI% - DECODI% + 1
       GRAFUNPR.MSChart1.Row = UJ%
       If UI% <= CAGRU% Then
           GRAFUNPR.MSChart1.Data = IVEN(UI%, UJ%)
         Else
           GRAFUNPR.MSChart1.Data = IVEN(0, UJ%)
       End If
     Next UJ%
   Next UI%
   '
90 Screen.MousePointer = 1
   If modo% = 1 Then
      GRAFUNPR.Show 1
   End If
   '
End Sub

Sub SORTANAFA()
   '
   RACLIEN.Clear
   FIN& = ULTREG&("LIANAFAC")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("LIANAFAC", U&)
     TTXX$ = "* " + AJUSTD$(CSTRING$(Left$(X$, 2), 1, 4, 0, 2), 4)
     TTXX$ = TTXX$ + AJUSTD$(CSTRING$(Mid$(X$, 3, 2), 1, 4, 0, 2), 4)
     TTXX$ = TTXX$ + AJUSTI$(CODEXT$(CVI(Mid$(X$, 5, 2))), 16)
     TTXX$ = TTXX$ + CSTRING$(MKS$(U&), 3, 6, 0, 2)
     '
     For KK& = 1 To RACLIEN.ListCount
       If RACLIEN.List(KK& - 1) > TTXX$ Then
         RACLIEN.AddItem TTXX$, KK& - 1
         GoTo 19
       End If
     Next KK&
     RACLIEN.AddItem TTXX$
     '
19 Next U&
   '
End Sub
'
Sub LISTANAFA()
   '
   Dim SUMACAN#(3), SUMABRU#(3), SUMANET#(3)
   Dim SUMADES#(3), SUMACOS#(3), SUMAREN#(3)
   '
   CODARANT% = 0
   GRUVEANT% = 0
   LINEAANT% = 0
   '
   NIVEINI% = 0: GoSub 60
   '
   JJ& = 0
   URE& = ULTREG&("LIANAFAC")
   FIN& = RACLIEN.ListCount
   For U& = 1 To FIN&
      RXX$ = RACLIEN.List(U& - 1)
      REG& = Val(Mid$(RXX$, 27, 6))
      If REG& > 0 Then
        If REG& <= URE& Then
          X$ = REGLEIDO$("LIANAFAC", REG&)
          LINEA% = CVI(Left$(X$, 2))
          GRUVE% = CVI(Mid$(X$, 3, 2))
          CODAR% = CVI(Mid$(X$, 5, 2))
          '
          If U& > 1 Then
            If LINEA% <> LINEAANT% Then
              GoSub 10
              GoSub 11
              GoSub 12        ' TOTALIZAR POR LINEA
              GRUVEANT% = 0
              CODARANT% = 0
              LINEAANT% = 0
            ElseIf GRUVE% <> GRUVEANT% Then
              GoSub 10
              GoSub 11        ' TOTALIZAR POR GRUPO DE VENTA
              GRUVEANT% = 0
              CODARANT% = 0
            ElseIf CODAR% <> CODARANT% Then
              GoSub 10        ' TOTAL POR ARTICULO
              CODARANT% = 0
            End If
          End If
          '
          If LINEA% <> LINEAANT% Then
            GoSub 20           ' ENCABEZADO DE LINEA
          End If
          If GRUVE% <> GRUVEANT% Then
            GoSub 21           ' ENCABEZADO DE GRUPO DE VENTA
          End If
          '
          TTXX$ = Space$(256)
          Call REPLA(TTXX$, ".", 1, 1)
          If CODAR% <> CODARANT% Then Call REPLA(TTXX$, CODEXT$(CODAR%), 6, 16)
          Call REPLA(TTXX$, Mid$(X$, 7, 18), 23, 12) ' NRO FACTURA
          FEX = CVI(Mid$(X$, 25, 2))
          Call REPLA(TTXX$, FECHATEX$(FEX), 37, 8)
          '
          CANTID# = CVD(Mid$(X$, 27, 8))
          Call REPLA(TTXX$, CSTRING$(MKD$(CANTID#), 4, 10, 0, 2), 45, 10)
          For NIVEL% = 0 To 3: SUMACAN#(NIVEL%) = SUMACAN#(NIVEL%) + CANTID#: Next NIVEL%
          '
          BRULIS# = CVD(Mid$(X$, 35, 8))
          Call REPLA(TTXX$, CSTRING$(MKD$(BRULIS#), 4, 11, 2, 2), 55, 11)
          For NIVEL% = 0 To 3: SUMABRU#(NIVEL%) = SUMABRU#(NIVEL%) + BRULIS#: Next NIVEL%
          '
          DESCUE# = CVD(Mid$(X$, 43, 8))
          Call REPLA(TTXX$, CSTRING$(MKD$(DESCUE#), 4, 11, 2, 2), 66, 11)
          For NIVEL% = 0 To 3: SUMADES#(NIVEL%) = SUMADES#(NIVEL%) + DESCUE#: Next NIVEL%
          '
          DESPOR# = CVS(Mid$(X$, 51, 4))
          Call REPLA(TTXX$, CSTRING$(MKD$(DESPOR#), 4, 10, 1, 2), 77, 10)
          '
          NETOPE# = CVD(Mid$(X$, 55, 8))
          Call REPLA(TTXX$, CSTRING$(MKD$(NETOPE#), 4, 11, 2, 2), 87, 11)
          For NIVEL% = 0 To 3: SUMANET#(NIVEL%) = SUMANET#(NIVEL%) + NETOPE#: Next NIVEL%
          '
          COSREP# = CVD(Mid$(X$, 63, 8))
          Call REPLA(TTXX$, CSTRING$(MKD$(COSREP#), 4, 11, 2, 2), 98, 11)
          For NIVEL% = 0 To 3: SUMACOS#(NIVEL%) = SUMACOS#(NIVEL%) + COSREP#: Next NIVEL%
          '
          RENTAS# = CVD(Mid$(X$, 71, 8))
          Call REPLA(TTXX$, CSTRING$(MKD$(RENTAS#), 4, 11, 2, 2), 109, 11)
          For NIVEL% = 0 To 3: SUMAREN#(NIVEL%) = SUMAREN#(NIVEL%) + RENTAS#: Next NIVEL%
          '
          RENPOR# = CVS(Mid$(X$, 79, 4))
          Call REPLA(TTXX$, CSTRING$(MKD$(RENPOR#), 4, 9, 1, 2), 120, 9)
          '
          GoSub 50
          LINEAANT% = LINEA%
          GRUVEANT% = GRUVE%
          CODARANT% = CODAR%
          '
        End If
      End If
   Next U&
   GoSub 10
   GoSub 11
   GoSub 12
   GoSub 13         ' TOTAL GENERAL
   Call SETULTREG("LISTANA", JJ&)
   Call CIERRARCH("LISTANA")
   Exit Sub
   '
10 TTXX$ = ".": GoSub 50
   TTXX$ = Space$(256)
   Call REPLA(TTXX$, "Tot. " + TRIM$(CODEXT$(CODARANT%)) + " - " + DESCRIT0$(CODARANT%), 1, 44)
   NIVEINI% = 3: GoSub 70
   TTXX$ = String$(248, "-"): GoSub 50
   TTXX$ = ".": GoSub 50
   GoSub 60
Return
   '
11 'TTXX$ = ".": GoSub 50
   TTXX$ = Space$(256)
   Call REPLA(TTXX$, "Tot. " + TITUGRU$, 1, 44)
   NIVEINI% = 2: GoSub 70
   TTXX$ = String$(248, "="): GoSub 50
   TTXX$ = ".": GoSub 50
   GoSub 60
Return
   '
12 'TTXX$ = ".": GoSub 50
   TTXX$ = Space$(256)
   Call REPLA(TTXX$, "Tot. " + TITULIN$, 1, 44)
   NIVEINI% = 1: GoSub 70
   TTXX$ = String$(248, "*"): GoSub 50
   TTXX$ = ".": GoSub 50
   GoSub 60
Return
   '
13 'TTXX$ = ".": GoSub 50
   TTXX$ = Space$(256)
   Call REPLA(TTXX$, "T O T A L   G E N E R A L ", 1, 44)
   NIVEINI% = 0: GoSub 70
   TTXX$ = ".": GoSub 50
   TTXX$ = String$(248, "*"): GoSub 50
   GoSub 60
Return
   '
   
20 TITULIN$ = "PRODUCTOS"
   If LINEA% = 2 Then TITULIN$ = "REPARACIONES"
   If LINEA% = 3 Then TITULIN$ = "DEBITOS - CREDITOS - VARIOS"
   '
   TTXX$ = TITULIN$: GoSub 50
   TTXX$ = String$(Len(TITULIN$), "*"): GoSub 50
   TTXX$ = ".": GoSub 50
Return
   '
21 TITUGRU$ = TRIM$(Str$(GRUVE%))
   While Len(TITUGRU$) < 3: TITUGRU$ = "0" + TITUGRU$: Wend
   TITUGRU$ = TITUGRU$ + " - "
   If GRUVE% = 999 Then
        TITUGRU$ = TITUGRU$ + "VARIOS"
      Else
        TITUGRU$ = TITUGRU$ + ECOARCH$("GRUCOVEN", MKI$(GRUVE%))
   End If
   TTXX$ = TITUGRU$: GoSub 50
   TTXX$ = ".": GoSub 50
Return
   '
50 JJ& = JJ& + 1
   Call GRAREG("LISTANA", TTXX$, JJ&)
Return
   '
60 For NIVEL% = NIVEINI% To 3
     SUMACAN#(NIVEL%) = 0
     SUMABRU#(NIVEL%) = 0
     SUMANET#(NIVEL%) = 0
     SUMADES#(NIVEL%) = 0
     SUMACOS#(NIVEL%) = 0
     SUMAREN#(NIVEL%) = 0
   Next NIVEL%
Return
   '
70 Call REPLA(TTXX$, CSTRING$(MKD$(SUMACAN#(NIVEINI%)), 4, 10, 0, 2), 45, 10)
   Call REPLA(TTXX$, CSTRING$(MKD$(SUMABRU#(NIVEINI%)), 4, 11, 2, 2), 55, 11)
   Call REPLA(TTXX$, CSTRING$(MKD$(SUMADES#(NIVEINI%)), 4, 11, 2, 2), 66, 11)
   '
   DESPORX# = 0: If SUMABRU#(NIVEINI%) > 0 Then DESPORX# = 100 * SUMADES#(NIVEINI%) / SUMABRU#(NIVEINI%)
   Call REPLA(TTXX$, CSTRING$(MKD$(DESPORX#), 4, 10, 1, 2), 77, 10)
   '
   Call REPLA(TTXX$, CSTRING$(MKD$(SUMANET#(NIVEINI%)), 4, 11, 2, 2), 87, 11)
   Call REPLA(TTXX$, CSTRING$(MKD$(SUMACOS#(NIVEINI%)), 4, 11, 2, 2), 98, 11)
   Call REPLA(TTXX$, CSTRING$(MKD$(SUMAREN#(NIVEINI%)), 4, 11, 2, 2), 109, 11)
   '
   DESPORX# = 0: If SUMANET#(NIVEINI%) > 0 Then DESPORX# = 100 * SUMAREN#(NIVEINI%) / SUMANET#(NIVEINI%)
   Call REPLA(TTXX$, CSTRING$(MKD$(DESPORX#), 4, 9, 1, 2), 120, 9)
   '
   GoSub 50
   '
Return
   '
End Sub
'
Sub ARMAMES()
   '
   DESDEFE% = 0: HASTAFE% = 0
   CORRESP$ = TRIM$(MESAN)
   CORRESQ$ = MESANO$(CORRESP$, DESDEFE%, HASTAFE%)
   '
   If DESDEFE% < 1 Then
     Call MENSERR(24, "Falta Indicar Período Mensual")
   End If
   '
   FEX = DESDEFE%: DESFE$ = FECHATEX$(FEX)
   FEX = HASTAFE%: HASFE$ = FECHATEX$(FEX)
   '
   CORRESQ$ = TRIM$(CORRESP$)
   ABREMES$ = Left$(CORRESQ$, 3) + Mid$(DESFE$, 7, 2)
   ANIO% = Val(Mid$(DESFE$, 7, 2)): If ANIO% < 80 Then ANIO% = 100 + ANIO%
   MESO% = Val(Mid$(DESFE$, 4, 2))
   CODIMES% = 12 * ANIO% + MESO%
   '
End Sub

Sub CALCURESU()
    '
    ' NO TOMA EL PRECIO DE LISTA SINO EL DEL PEDIDO
    ' SI NO HAY PEDIDO, EL PRECIO FACTURADO
    '
    Dim PUNMO$(32767), TICLIE%(32767)
    '
    Dim NCI%(8192)
    Dim FFI%(8192)
    Dim TI%(8192)
    Dim VA%(8192)
    Dim NRO&(8192)
    Dim SIFAC%(8192)
    Dim IFAC#(8192)
    Dim IDET#(8192)
    Dim PUMOFAC$(8192)
    '
    'Dim RMVS As String * 128
    '
    MODOCALCOS% = 1   ' dolarizado
    PORCALCOS$ = TRIM$(CONTROL$("IGESVEN", "MOCALCOS"))
    If PORCALCOS$ = "PESIFIC" Then
      MODOCALCOS% = 2 ' PESIFICADO
    End If
    '
    For KKI% = 1 To 32766
      COREGIS#(KKI%) = 0
    Next KKI%
    '
    If TRIM$(Label3) <> "" Then
      If COMALTER%("Puede Mantener los Costos Registrados\Correspondientes al Cálculo Anterior de Este Mes,\o Re-Calcular Dichos Valores\\Mantener Anteriores\Re-Calcular Costos") <> 2 Then
        Call CARCOSINFO
      End If
    End If
    VEGRUVE$ = ""
   '
   If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then
      VEGRUVE$ = MKI$(0) + MKI$(1)
      VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(2)
      VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(4)
      VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(5)
      VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(6)
      VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(7)
      VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(8)
      VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(9)
      VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(99)
      GoTo 15
   End If
   '
    For U& = 1 To ULTREG&("GRUCOVEN")
      X$ = REGLEIDO$("GRUCOVEN", U&)
      GRUCO% = CVI(Left$(X$, 2))
      VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
    Next U&
    '
15  DES% = DESDEFE%
    HAS% = HASTAFE%
    '
    FEX = HOY(HOS$)
    ANIOX% = Val(Mid$(HOS$, 7, 2))
    If ANIOX% < 80 Then ANIOX% = 100 + ANIOX%
    MESOX% = Val(Mid$(HOS$, 4, 2))
    CODIMESACT% = 12 * ANIOX% + MESOX%
    '
    Screen.MousePointer = 11
    For KKI% = 0 To 3: Label1(KKI%) = "": Next KKI%
    '
    Call COPYSTRU("MOVISTO", "MOVISVIE")
    GoSub 111    ' carga vector de facturas emitidas
    GoSub 112    ' carga vector detalle de facturacion - MOVISTO
    GoSub 113    ' carga vector detalle de facturacion - MOVIREPA
    '
    UREMO& = ULTREG&("MOVISTO")
    UREPA& = ULTREG&("MOVIREPA")
    FIN& = CAFAC%
    J& = 0: J1& = 0: J2& = 0
    '
    For JII% = 1 To CAFAC%
      I& = JII%
      Call TRACE(20, I&, FIN&)
      '
      If TI%(JII%) < 1 Or TI%(JII%) > 6 Then GoTo 28  ' descarta ajustes
      REMIFAC$ = "FC"
      If TI%(JII%) = 3 Then REMIFAC$ = "ND"
      If TI%(JII%) = 6 Then REMIFAC$ = "NC"
      VASS$ = "A": If VA%(JII%) = 2 Then VASS$ = "B"
                   If VA%(JII%) = 3 Then VASS$ = "E"
      REMIFAC$ = REMIFAC$ + "-" + TRIM$(Str$(NRO&(JII%)))
      REMIFAC$ = REMIFAC$ + "-" + VASS$
      '
      COEF# = 1
      If Abs(IDET#(JII%)) >= 0.005 Then
        COEF# = (Abs(IFAC#(JII%))) / (Abs(IDET#(JII%)))
      End If
      '
      For KKI% = 1 To Len(PUMOFAC$(JII%)) Step 6
         ARMO% = CVI(Mid$(PUMOFAC$(JII%), KKI%, 2))
         REG& = CVS(Mid$(PUMOFAC$(JII%), KKI% + 2, 4))
         If ARMO% = 1 Then
              X$ = REGLEIDO$("MOVISTO", REG&)
            ElseIf ARMO% = 2 Then
              X$ = REGLEIDO$("MOVISVIE", REG&)
            ElseIf ARMO% = 3 Then
              X$ = REGLEIDO$("MOVIREPA", Abs(REG&))
            Else
              GoTo 27
         End If
         If Mid$(X$, 23, 2) = "PL" Then GoTo 27 ' descarta packing list
         '
         CLI% = NCI%(JII%)
         FF% = FFI%(JII%)
         NRO1& = NRO&(JII%)
         CII% = CVI(Mid$(X$, 3, 2))
         CANTU# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
         VUNI = CVD(Mid$(X$, 73, 8)) * COEF#
         '
         ' AQUI BUSCAR PRECIO DE LA NOTA DE PEDIDO
         PPCC$ = RECFILT$("PEDDETA", 3, MKI$(CLI%))
         NREMITO& = Val(Mid$(X$, 26, 7))
         For KKX& = 1 To Len(PPCC$) Step 4
            KKY& = CVS(Mid$(PPCC$, KKX&, 4))
            KKZ$ = REGLEIDO$("PEDDETA", KKY&)
            If CVI(Mid$(KKZ$, 33, 2)) = CII% Then
              If CVS(Mid$(KKZ$, 105, 4)) = NREMITO& Then
                PRELI = CVS(Mid$(KKZ$, 71, 4))
                GoTo 18
              End If
            End If
         Next KKX&
         '
         ' NO ENCONTRO PEDIDO - TOMO EL PRECIO FACTURADO
         PRELI = VUNI
         PRELISTX = PRELISTA#("LIPRE001", CII%)
18       If PRELI < 0.1 Then PRELI = CVD(Mid$(X$, 73, 8))
         If REG& < 0 Then PRELI = VUNI    ' reparaciones
         '
         'SI SE FACTURO A PRECIO MAYOR AL PEDIDO TOMA EL FACTURADO
         If PRELI < VUNI Then PRELI = VUNI
         VALOUN = VUNI
         VATOT# = CANTU# * VALOUN
         '
         CDD$ = CODEXT$(CII%)
         If CII% < 1 Then
           CDD$ = "Mat.s/Código"
           If TI%(JII%) = 3 Then
               CII% = (-1) ' DEBITOS LIBRES
               CDD$ = "Déb.Libre"
             ElseIf TI%(JII%) > 3 Then
               CII% = (-2) ' CREDITOS LIBRES
               CDD$ = "Créd.Libre"
           End If
         End If
         '
         If Abs(VALOUN) < 0.005 Then
19          TTXX$ = "Precio Unitario (en PESOS) para " + TRIM$(CDD$) + " en " + TRIM$(REMIFAC$)
            INVUNI# = XVALO(InputBox$(TTXX$, "Falta Precio Unitario en Comprobante " + REMIFAC$))
            'If INVUNI# < 0.005 Then GoTo 19
            VALOUN = INVUNI#
            VATOT# = CANTU# * VALOUN
            Call REPLA(X$, MKD$(INVUNI#), 73, 8)
            If ARMO% = 1 Then
                Call GRAREG("MOVISTO", X$, REG&)
              ElseIf ARMO% = 2 Then
                Call GRAREG("MOVISVIE", X$, REG&)
              ElseIf ARMO% = 3 Then
                Call GRAREG("MOVIREPA", X$, Abs(REG&))
            End If
         End If
         Z$ = REGBLAN$("DVECLART")
         Call REPLA(Z$, MKI$(CODIMES%), 1, 2)
         Call REPLA(Z$, MKI$(1), 3, 2)       ' productos
         If REG& < 0 Then
            Call REPLA(Z$, MKI$(2), 3, 2)    ' reparaciones
         End If
         Call REPLA(Z$, MKI$(CLI%), 5, 2)
         Call REPLA(Z$, MKI$(CII%), 7, 2)
         Call REPLA(Z$, CDD$, 9, 16)
         Call REPLA(Z$, REMIFAC$, 25, 24)
         Call REPLA(Z$, MKI$(FF%), 49, 2)
         Call REPLA(Z$, MKD$(CANTU#), 51, 8)
         Call REPLA(Z$, String$(70, 0), 59, 70)
         Call REPLA(Z$, MKD$(CDbl(PRELISTX)), 129, 8)
         Call REPLA(Z$, MKD$(CANTU# * PRELISTX), 137, 8)
         Call REPLA(Z$, MKD$(CDbl(PRELI)), 145, 8)
         Call REPLA(Z$, MKD$(CANTU# * PRELI), 153, 8)
         Call REPLA(Z$, MKD$(CDbl(VUNI)), 161, 8)
         Call REPLA(Z$, MKD$(CANTU# * VUNI), 169, 8)
         PORDE1 = 0: If PRELISTX > 0 Then PORDE1 = 100 * (PRELISTX - PRELI) / PRELISTX
         Call REPLA(Z$, MKS$(PORDE1), 177, 4)
         PORDE2 = 0: If PRELI > 0 Then PORDE2 = 100 * (PRELI - VUNI) / PRELI
         Call REPLA(Z$, MKS$(PORDE2), 181, 4)
         '
         COU = 0
         If CII% > 0 Then
            If COREGIS#(CII%) > 0.00005 Then
                 COU = COREGIS#(CII%)
               Else
                 If CODIMESACT% = CODIMES% Or MODOCALCOS% = 2 Then
                     COU = COSMATCAL(CII%, 1) * (1 + Val(MERMAPRO) / 100) ' EN PESOS
                   Else
                     CODOLA = COSMATCAL(CII%, 2)   ' LO CALCULA EN DOLARES
                     COU = CODOLA * Val(TIPOCAM) * (1 + Val(MERMAPRO) / 100)
                 End If
            End If
            If REG& < 0 Then COU = 0          ' REPARACIONES
         End If
         '
         Call REPLA(Z$, MKD$(CDbl(COU)), 185, 8)
         Call REPLA(Z$, MKD$(CANTU# * COU), 193, 8)
         RENTA# = CANTU# * (VUNI - COU)
         Call REPLA(Z$, MKD$(RENTA#), 201, 8)
         POREN = 0: If Abs(CANTU# * VUNI) > 0 Then POREN = 100 * RENTA# / (CANTU# * VUNI)
         Call REPLA(Z$, MKS$(POREN), 209, 4)
         Call REPLA(Z$, MKS$(Val(TIPOCAM)), 213, 4)  ' TIPO DE CAMBIO
         '
         J1& = J1& + 1
         Call GRAREG("!DVECLART", Z$, J1&)
         '
         VALOMO# = CANTU# * VUNI * SIFAC%(JII%)
         IFAC#(JII%) = IFAC#(JII%) - VALOMO#
         '
         If CII% = (-2) Then
             Label1(3) = FORMATNUM$(VALOMO# * SIFAC%(JII%) + Val(Label1(3)), "F12.2")
             Label1(3).Refresh
           ElseIf CII% = (-1) Then
             Label1(2) = FORMATNUM$(VALOMO# * SIFAC%(JII%) + Val(Label1(2)), "F12.2")
             Label1(2).Refresh
           Else
             Label1(1) = FORMATNUM$(VALOMO# * SIFAC%(JII%) + Val(Label1(1)), "F12.2")
             Label1(1).Refresh
         End If
         '
27    Next KKI%
      '
      If Abs(IFAC#(JII%)) >= 0.005 Then
         CLI% = NCI%(JII%)
         FF% = FFI%(JII%)
         NRO1& = NRO&(JII%)
         CII% = 0
         CDD$ = "Mat.s/Código"
         If TI%(JII%) = 3 Then
              CII% = (-1) ' DEBITOS LIBRES
              CDD$ = "Déb.Libre"
            ElseIf TI%(JII%) > 3 Then
              CII% = (-2) ' CREDITOS LIBRES
              CDD$ = "Créd.Libre"
         End If
         CANTU# = SIFAC%(JII%)
         VUNI = Abs(IFAC#(JII%))
         PRELI = VUNI
         PRELISTX = VUNI
         VALOUN = VUNI
         VATOT# = CANTU# * VALOUN
         '
         Z$ = REGBLAN$("DVECLART")
         Call REPLA(Z$, MKI$(CODIMES%), 1, 2)
         Call REPLA(Z$, MKI$(1), 3, 2)       ' productos
         If REG& < 0 Then
            Call REPLA(Z$, MKI$(2), 3, 2)    ' reparaciones
         End If
         Call REPLA(Z$, MKI$(CLI%), 5, 2)
         Call REPLA(Z$, MKI$(CII%), 7, 2)
         Call REPLA(Z$, CDD$, 9, 16)
         Call REPLA(Z$, REMIFAC$, 25, 24)
         Call REPLA(Z$, MKI$(FF%), 49, 2)
         Call REPLA(Z$, MKD$(CANTU#), 51, 8)
         Call REPLA(Z$, String$(70, 0), 59, 70)
         Call REPLA(Z$, MKD$(CDbl(PRELISTX)), 129, 8)
         Call REPLA(Z$, MKD$(CANTU# * PRELISTX), 137, 8)
         Call REPLA(Z$, MKD$(CDbl(PRELI)), 145, 8)
         Call REPLA(Z$, MKD$(CANTU# * PRELI), 153, 8)
         Call REPLA(Z$, MKD$(CDbl(VUNI)), 161, 8)
         Call REPLA(Z$, MKD$(CANTU# * VUNI), 169, 8)
         PORDE1 = 0: If PRELISTX > 0 Then PORDE1 = 100 * (PRELISTX - PRELI) / PRELISTX
         Call REPLA(Z$, MKS$(PORDE1), 177, 4)
         PORDE2 = 0: If PRELI > 0 Then PORDE2 = 100 * (PRELI - VUNI) / PRELI
         Call REPLA(Z$, MKS$(PORDE2), 181, 4)
         COU = 0
         Call REPLA(Z$, MKD$(CDbl(COU)), 185, 8)
         Call REPLA(Z$, MKD$(CANTU# * COU), 193, 8)
         RENTA# = CANTU# * (VUNI - COU)
         Call REPLA(Z$, MKD$(RENTA#), 201, 8)
         POREN = 0: If Abs(CANTU# * VUNI) > 0 Then POREN = 100 * RENTA# / (CANTU# * VUNI)
         Call REPLA(Z$, MKS$(POREN), 209, 4)
         Call REPLA(Z$, MKS$(Val(TIPOCAM)), 213, 4)  ' TIPO DE CAMBIO
         '
         J1& = J1& + 1
         Call GRAREG("!DVECLART", Z$, J1&)
         '
         VALOMO# = VATOT#
         If CII% = (-2) Then
             Label1(3) = FORMATNUM$(VALOMO# + Val(Label1(3)), "F12.2")
             Label1(3).Refresh
           ElseIf CII% = (-1) Then
             Label1(2) = FORMATNUM$(VALOMO# + Val(Label1(2)), "F12.2")
             Label1(2).Refresh
           Else
             Label1(1) = FORMATNUM$(VALOMO# + Val(Label1(1)), "F12.2")
             Label1(1).Refresh
         End If
         '
      End If
      '
28  Next JII%
    '
    Call SETULTREG("!DVECLART", J1&)
    Call CIERRARCH("!DVECLART")
    '
    Call GENERESU
    Call GUARDARESU
    '
    '
    HOII% = HOY(HOS$)
    Label3 = HOS$ + " - " + Left$(Time$, 5) + " hs"
    Label3.Refresh
    '
    TTT$ = REGBLAN$("IRESVENT")
    Call REPLA(TTT$, MKI$(CODIMES%), 1, 2)
    Call REPLA(TTT$, MESAN, 3, 14)
    Call REPLA(TTT$, TRIM$(Label3), 17, 16)
    Call REPLA(TTT$, MKD$(Val(Label1(0))), 33, 8)
    Call REPLA(TTT$, MKD$(Val(Label1(1))), 41, 8)
    Call REPLA(TTT$, MKD$(Val(Label1(2))), 49, 8)
    Call REPLA(TTT$, MKD$(Val(Label1(3))), 57, 8)
    Call REPLA(TTT$, MKD$(Val(TIPOCAM)), 65, 8)
    Call REPLA(TTT$, MKS$(Val(MERMAPRO)), 73, 4)
    '
    REIRE& = 0
    RFF$ = RECFILT$("IRESVENT", 1, MKI$(CODIMES%))
    If Len(RFF$) >= 4 Then
      REIRE& = CVS(Left$(RFF$, 4))
    End If
    If REIRE& > ULTREG&("IRESVENT") Then REIRE& = 0
    If REIRE& < 1 Then REIRE& = 1 + ULTREG&("IRESVENT")
    Call GRAREG("IRESVENT", TTT$, REIRE&)
    Call CIERRARCH("IRESVENT")
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    Screen.MousePointer = 1
    '
Exit Sub
    '
' ***********************************************************
    '
111 CAFAC% = 0
    SUFAC# = 0
    FIN& = ULTREG&("FACTUR")
    '
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XS$ = REGLEIDO$("FACTUR", I&)
      FF% = CVI(Mid$(XS$, 7, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          TIX% = Asc(Mid$(XS$, 13, 1))
          If TIX% > 0 Then        ' descarta ajustes de saldos
            If TIX% < 7 Then
              '
              CAFAC% = CAFAC% + 1
              NCI%(CAFAC%) = CVI(Left$(XS$, 2))
              FFI%(CAFAC%) = FF%
              TI%(CAFAC%) = Asc(Mid$(XS$, 13, 1))
              If TI%(CAFAC%) < 0 Or TI%(CAFAC%) > 6 Then TI%(CAFAC%) = 2 ' por default factura
              VA%(CAFAC%) = Asc(Mid$(XS$, 14, 1))
              If VA%(CAFAC%) < 1 Or VA%(CAFAC%) > 2 Then VA%(CAFAC%) = 3 ' por default factura
              NRO&(CAFAC%) = CVS(Mid$(XS$, 3, 4))
              IFAC#(CAFAC%) = Abs(CVD(Mid$(XS$, 15, 8)) + CVD(Mid$(XS$, 49, 8)))
              '
              SIG% = 1: If TI%(CAFAC%) > 3 Then SIG% = (-1)
              SIFAC%(CAFAC%) = SIG%
              SUFAC# = SUFAC# + IFAC#(CAFAC%) * SIG%
              Label1(0) = CSTRING$(MKD$(SUFAC#), 4, 12, 2, 2)
              Label1(0).Refresh
            End If
          End If
        End If
      End If
    Next I&
    '
Return
'
'
112  ARMOVI$ = "MOVISTO"
     ARMO% = 1
1121 PRIRE& = 999999: ULTRE& = 0
     FIN& = ULTREG&(ARMOVI$)
     'FINUMO% = FILENBR%("MOVISTO")
     PRIRESTO& = 0
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       'Get #FINUMO%, I&, RMVS$
       XS$ = REGLEIDO$(ARMOVI$, I&)
       'XS$ = RMVS$
       FF% = CVI(Left$(XS$, 2))
       If FF% >= DES% Then
         If FF% <= HAS% Then
           If PRIRESTO& = 0 Then PRIRESTO& = I&
           COMO$ = UCase$(Mid$(XS$, 33, 2))
           If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
             TP1% = 2: If COMO$ = "NC" Then TP1% = 6
             CLI1% = CVI(Mid$(XS$, 84, 2))
             If COMO$ = "ND" Then TP1% = 3
             NRO1& = Val(Mid$(XS$, 36, 7))
             CAN# = CVD(Mid$(XS$, 96, 8)) - CVD(Mid$(XS$, 88, 8))
             PRE# = CVD(Mid$(XS$, 73, 8))
             TOT# = PRE# * CAN#
             If I& < PRIRE& Then PRIRE& = I&
             If I& > ULTRE& Then ULTRE& = I&
             REMOSTO& = I&
             GoSub 120
           End If
         End If
       End If
     Next I&
     '
     If ARMOVI$ = "MOVISTO" Then
       ARMO% = 2
       ARMOVI$ = "MOVISVIE"
       GoTo 1121
     End If
     '
Return
'
113 FIN& = ULTREG&("MOVIREPA")
    ARMO% = 3
    'FINUMO% = FILENBR%("MOVIREPA")
    PRIREREP& = 0
    For I& = 1 To FIN&
      Call TRACE(20, I& - PRIREMO&, FIN& - PRIREMO&)
      'Get #FINUMO%, I&, RMVS$   '
      XS$ = REGLEIDO$("MOVIREPA", I&)
      'XS$ = RMVS$
      FF% = CVI(Left$(XS$, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          If PRIREREP& = 0 Then PRIREREP& = I&
          COMO$ = UCase$(Mid$(XS$, 33, 2))
          If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
            TP1% = 2: If COMO$ = "NC" Then TP1% = 6
            CLI1% = CVI(Mid$(XS$, 84, 2))
            If COMO$ = "ND" Then TP1% = 3
            NRO1& = Val(Mid$(XS$, 36, 7))
            CAN# = CVD(Mid$(XS$, 96, 8)) - CVD(Mid$(XS$, 88, 8))
            PRE# = CVD(Mid$(XS$, 73, 8))
            TOT# = PRE# * CAN#
            If I& < PRIRE& Then PRIRE& = I&
            If I& > ULTRE& Then ULTRE& = I&
            REMOSTO& = I&
            GoSub 120
          End If
        End If
      End If
    Next I&
    '
Return
'
120 VUELTA% = 0
121 For JII% = 1 To CAFAC%
      If TI%(JII%) = TP1% Then
        If NRO&(JII%) = NRO1& Then
          If NCI%(JII%) = CLI1% Or VUELTA% = 1 Then
            If FFI%(JII%) = FF% Or VUELTA% = 1 Then
              If Abs(IFAC#(JII%)) >= 0.005 Then ' descarta anuladas
                IDET#(JII%) = IDET#(JII%) + TOT#
                PUMOFAC$(JII%) = PUMOFAC$(JII%) + MKI$(ARMO%) + MKS$(REMOSTO&)
              End If
              GoTo 124    ' encontro documento
            End If
          End If
        End If
      End If
    Next JII%
    '
    If VUELTA% = 0 Then
      VUELTA% = 1: GoTo 121
    End If
    '
    If VUELTA% = 1 Then
      If TP1% = 2 Then
        TP1% = 3
        VUELTA% = 2     ' recupera debitos marcados como facturas
        GoTo 121
      End If
    End If
    '
124 Return
    '

End Sub

Sub GENERESU()
    '
    Dim ABUS$(8192, 3)
    J1& = 0: J2& = 0: J3& = 0
    '
    FIN& = ULTREG&("!DVECLART")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("!DVECLART", U&)
      '
      CATEPRO% = CVI(Mid$(Z$, 3, 2))
      NCLIE% = CVI(Mid$(Z$, 5, 2))
      CII% = CVI(Mid$(Z$, 7, 2))
      '
      CANTU# = CVD(Mid$(Z$, 51, 8))
      PRUNILIS# = CVD(Mid$(Z$, 129, 8))
      TOTALIS# = CVD(Mid$(Z$, 137, 8))
      PRUNICOT# = CVD(Mid$(Z$, 145, 8))
      TOTACOT# = CVD(Mid$(Z$, 153, 8))
      PRUNIVEN# = CVD(Mid$(Z$, 161, 8))
      TOTAVEN# = CVD(Mid$(Z$, 169, 8))
      PORDESC1 = CVS(Mid$(Z$, 177, 4))
      PORDESC2 = CVS(Mid$(Z$, 181, 4))
      COSTOUNI# = CVD(Mid$(Z$, 185, 8))
      COSTOTO# = CVD(Mid$(Z$, 193, 8))
      RENTATO# = CVD(Mid$(Z$, 201, 8))
      PORENTA = CVS(Mid$(Z$, 209, 4))
      TCDOL = CVS(Mid$(Z$, 213, 4))
      '
      ABUSQUE$ = Mid$(Z$, 3, 6)
      GoSub 100        ' Agregar a RVECLART
      '
      If CII% >= 0 Then  ' excluye debitos y creditos libres
        ABUSQUE$ = Mid$(Z$, 5, 2)
        GoSub 200      ' Agregar a CVECLART
        '
        ABUSQUE$ = Mid$(Z$, 3, 2) + Mid$(Z$, 7, 2)
        GoSub 300      ' Agregar a PVECLART
      End If
      '
    Next U&
    '
    Call SETULTREG("!RVECLART", J1&)
    Call CIERRARCH("!RVECLART")
    '
    Call SETULTREG("!CVECLART", J2&)
    Call CIERRARCH("!CVECLART")
    '
    Call SETULTREG("!PVECLART", J3&)
    Call CIERRARCH("!PVECLART")
    '
Exit Sub
'
100 For JJ& = 1 To J1&
      If ABUS$(JJ&, 1) = ABUSQUE$ Then
        Z1$ = REGLEIDO$("!RVECLART", JJ&)
        CANTUT# = CANTU# + CVD(Mid$(Z1$, 51, 8))
           Call REPLA(Z1$, MKD$(CANTUT#), 51, 8)
        TOTALIST# = TOTALIS# + CVD(Mid$(Z1$, 137, 8))
           Call REPLA(Z1$, MKD$(TOTALIST#), 137, 8)
           If Abs(CANTUT#) > 0.5 Then
             Call REPLA(Z1$, MKD$(TOTALIST# / CANTUT#), 129, 8)
           End If
        TOTACOTT# = TOTACOT# + CVD(Mid$(Z1$, 153, 8))
           Call REPLA(Z1$, MKD$(TOTACOTT#), 153, 8)
           If Abs(CANTUT#) > 0.5 Then
             Call REPLA(Z1$, MKD$(TOTACOTT# / CANTUT#), 145, 8)
           End If
        TOTAVENT# = TOTAVEN# + CVD(Mid$(Z1$, 169, 8))
           Call REPLA(Z1$, MKD$(TOTAVENT#), 169, 8)
           If Abs(CANTUT#) > 0.5 Then
             Call REPLA(Z1$, MKD$(TOTAVENT# / CANTUT#), 161, 8)
           End If
        If Abs(TOTALIST#) > 0.005 Then
           PORDESCT = 100 * (TOTALIST# - TOTACOTT#) / TOTALIST#
           Call REPLA(Z1$, MKS$(PORDESCT), 177, 4)
        End If
        If Abs(TOTACOTT#) > 0.005 Then
           PORDESCT = 100 * (TOTACOTT# - TOTAVENTT#) / TOTACOTT#
           Call REPLA(Z1$, MKS$(PORDESCT), 177, 4)
        End If
        COSTOTOT# = COSTOTO# + CVD(Mid$(Z1$, 193, 8))
           Call REPLA(Z1$, MKD$(COSTOTOT#), 193, 8)
           If Abs(CANTUT#) > 0.5 Then
             Call REPLA(Z1$, MKD$(COSTOTOT# / CANTUT#), 185, 8)
           End If
         RENTATOT# = RENTATO# + CVD(Mid$(Z1$, 201, 8))
           Call REPLA(Z1$, MKD$(RENTATOT#), 201, 8)
         If Abs(TOTAVENT#) > 0.005 Then
           POREN = 100 * RENTATOT# / TOTAVENT#
           Call REPLA(Z1$, MKS$(POREN), 209, 4)
         End If
         '
         Call GRAREG("!RVECLART", Z1$, JJ&)
         GoTo 109
      End If
    Next JJ&
    '
    Z1$ = Z$
    J1& = J1& + 1
    Call GRAREG("!RVECLART", Z1$, J1&)
    ABUS$(J1&, 1) = ABUSQUE$
    '
109 Return
    '
    '
200 For JJ& = 1 To J2&
      If ABUS$(JJ&, 2) = ABUSQUE$ Then
        Z1$ = REGLEIDO$("!CVECLART", JJ&)
        CANTUT# = CANTU# + CVD(Mid$(Z1$, 51, 8))
           Call REPLA(Z1$, MKD$(CANTUT#), 51, 8)
        TOTALIST# = TOTALIS# + CVD(Mid$(Z1$, 137, 8))
           Call REPLA(Z1$, MKD$(TOTALIST#), 137, 8)
           If Abs(CANTUT#) > 0.5 Then
             Call REPLA(Z1$, MKD$(TOTALIST# / CANTUT#), 129, 8)
           End If
        TOTACOTT# = TOTACOT# + CVD(Mid$(Z1$, 153, 8))
           Call REPLA(Z1$, MKD$(TOTACOTT#), 153, 8)
           If Abs(CANTUT#) > 0.5 Then
             Call REPLA(Z1$, MKD$(TOTACOTT# / CANTUT#), 145, 8)
           End If
        TOTAVENT# = TOTAVEN# + CVD(Mid$(Z1$, 169, 8))
           Call REPLA(Z1$, MKD$(TOTAVENT#), 169, 8)
           If Abs(CANTUT#) > 0.5 Then
             Call REPLA(Z1$, MKD$(TOTAVENT# / CANTUT#), 161, 8)
           End If
        If Abs(TOTALIST#) > 0.005 Then
           PORDESCT = 100 * (TOTALIST# - TOTACOTT#) / TOTALIST#
           Call REPLA(Z1$, MKS$(PORDESCT), 177, 4)
        End If
        If Abs(TOTACOTT#) > 0.005 Then
           PORDESCT = 100 * (TOTACOTT# - TOTAVENTT#) / TOTACOTT#
           Call REPLA(Z1$, MKS$(PORDESCT), 177, 4)
        End If
        COSTOTOT# = COSTOTO# + CVD(Mid$(Z1$, 193, 8))
           Call REPLA(Z1$, MKD$(COSTOTOT#), 193, 8)
           If Abs(CANTUT#) > 0.5 Then
             Call REPLA(Z1$, MKD$(COSTOTOT# / CANTUT#), 185, 8)
           End If
         RENTATOT# = RENTATO# + CVD(Mid$(Z1$, 201, 8))
           Call REPLA(Z1$, MKD$(RENTATOT#), 201, 8)
         If Abs(TOTAVENT#) > 0.005 Then
           POREN = 100 * RENTATOT# / TOTAVENT#
           Call REPLA(Z1$, MKS$(POREN), 209, 4)
         End If
         '
         Call GRAREG("!CVECLART", Z1$, JJ&)
         GoTo 209
      End If
    Next JJ&
    '
    Z1$ = Z$
    J2& = J2& + 1
    Call GRAREG("!CVECLART", Z1$, J2&)
    ABUS$(J2&, 2) = ABUSQUE$
    '
209 Return
    '
    '
300 For JJ& = 1 To J3&
      If ABUS$(JJ&, 3) = ABUSQUE$ Then
        Z1$ = REGLEIDO$("!PVECLART", JJ&)
        CANTUT# = CANTU# + CVD(Mid$(Z1$, 51, 8))
           Call REPLA(Z1$, MKD$(CANTUT#), 51, 8)
        TOTALIST# = TOTALIS# + CVD(Mid$(Z1$, 137, 8))
           Call REPLA(Z1$, MKD$(TOTALIST#), 137, 8)
           If Abs(CANTUT#) > 0.5 Then
             Call REPLA(Z1$, MKD$(TOTALIST# / CANTUT#), 129, 8)
           End If
        TOTACOTT# = TOTACOT# + CVD(Mid$(Z1$, 153, 8))
           Call REPLA(Z1$, MKD$(TOTACOTT#), 153, 8)
           If Abs(CANTUT#) > 0.5 Then
             Call REPLA(Z1$, MKD$(TOTACOTT# / CANTUT#), 145, 8)
           End If
        TOTAVENT# = TOTAVEN# + CVD(Mid$(Z1$, 169, 8))
           Call REPLA(Z1$, MKD$(TOTAVENT#), 169, 8)
           If Abs(CANTUT#) > 0.5 Then
             Call REPLA(Z1$, MKD$(TOTAVENT# / CANTUT#), 161, 8)
           End If
        If Abs(TOTALIST#) > 0.005 Then
           PORDESCT = 100 * (TOTALIST# - TOTACOTT#) / TOTALIST#
           Call REPLA(Z1$, MKS$(PORDESCT), 177, 4)
        End If
        If Abs(TOTACOTT#) > 0.005 Then
           PORDESCT = 100 * (TOTACOTT# - TOTAVENTT#) / TOTACOTT#
           Call REPLA(Z1$, MKS$(PORDESCT), 177, 4)
        End If
        COSTOTOT# = COSTOTO# + CVD(Mid$(Z1$, 193, 8))
           Call REPLA(Z1$, MKD$(COSTOTOT#), 193, 8)
           If Abs(CANTUT#) > 0.5 Then
             Call REPLA(Z1$, MKD$(COSTOTOT# / CANTUT#), 185, 8)
           End If
         RENTATOT# = RENTATO# + CVD(Mid$(Z1$, 201, 8))
           Call REPLA(Z1$, MKD$(RENTATOT#), 201, 8)
         If Abs(TOTAVENT#) > 0.005 Then
           POREN = 100 * RENTATOT# / TOTAVENT#
           Call REPLA(Z1$, MKS$(POREN), 209, 4)
         End If
         '
         Call GRAREG("!PVECLART", Z1$, JJ&)
         GoTo 309
      End If
    Next JJ&
    '
    Z1$ = Z$
    J3& = J3& + 1
    Call GRAREG("!PVECLART", Z1$, J3&)
    ABUS(J3&, 3) = ABUSQUE$
    '
309 Return
    '
End Sub
'
Sub GUARDARESU()
   '
   AGUARDA$ = "DVECLART": GoSub 10
   AGUARDA$ = "RVECLART": GoSub 10
   AGUARDA$ = "CVECLART": GoSub 10
   AGUARDA$ = "PVECLART": GoSub 10
   '
Exit Sub
'
'
10 URN& = ULTREG&(AGUARDA$)
   RNC$ = RECFILT$(AGUARDA$, 1, MKI$(CODIMES%))
   KK& = (-3)
   '
   FIN& = ULTREG&("!" + AGUARDA$)
   For U& = 1 To FIN&
      TTYY$ = REGLEIDO$("!" + AGUARDA$, U&)
      Call REPLA(TTYY$, MKI$(CODIMES%), 1, 2)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG(AGUARDA$, TTYY$, RENOTA&)
   Next U&
   '
   While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG(AGUARDA$, String$(256, 0), RENOTA&)
   Wend
   '
   Call CIERRARCH(AGUARDA$)
   Call CIERRARCH("!" + AGUARDA$)
   '
Return
'
'
End Sub
'
Sub RECURESMEN()
   '
   Screen.MousePointer = 11
   '
   For KKI% = 0 To 3
      Label1(KKI%) = ""
   Next KKI%
   Label3 = ""
   TIPOCAM = ""
   MERMAPRO = " 0.0"
   '
   RFF$ = RECFILT$("IRESVENT", 1, MKI$(CODIMES%))
   If Len(RFF$) >= 4 Then
      REIRE& = CVS(Left$(RFF$, 4))
      TTT$ = REGLEIDO$("IRESVENT", REIRE&)
      Label3 = TRIM$(Mid$(TTT$, 17, 16)) + " hs"
      Label1(0) = FORMATNUM$(CVD(Mid$(TTT$, 33, 8)), "F12.2")
      Label1(1) = FORMATNUM$(CVD(Mid$(TTT$, 41, 8)), "F12.2")
      Label1(2) = FORMATNUM$(CVD(Mid$(TTT$, 49, 8)), "F12.2")
      Label1(3) = FORMATNUM$(CVD(Mid$(TTT$, 57, 8)), "F12.2")
      TIPOCAM = FORMATNUM$(CVD(Mid$(TTT$, 65, 8)), "F9.3")
      MERMAPRO = FORMATNUM$(CVS(Mid$(TTT$, 73, 4)), "F4.1")
   End If
   '
   RFF$ = RECFILT$("DVECLART", 1, MKI$(CODIMES%))
   JJ& = 0
   For U& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, U&, 4))
     Z$ = REGLEIDO$("DVECLART", REG&)
     JJ& = JJ& + 1
     Call GRAREG("!DVECLART", Z$, JJ&)
   Next U&
   Call SETULTREG("!DVECLART", JJ&)
   Call CIERRARCH("!DVECLART")
   '
   RFF$ = RECFILT$("RVECLART", 1, MKI$(CODIMES%))
   JJ& = 0
   For U& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, U&, 4))
     Z$ = REGLEIDO$("RVECLART", REG&)
     JJ& = JJ& + 1
     Call GRAREG("!RVECLART", Z$, JJ&)
   Next U&
   Call SETULTREG("!RVECLART", JJ&)
   Call CIERRARCH("!RVECLART")
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub

Private Sub Option11_DblClick()
   Call TAVECOSTO(1)
End Sub

Private Sub Option12_DblClick()
   Call GRAVECOSTO(1)
End Sub

Private Sub Option13_DblClick()
   Call GRAPORMABRUVE(1)
End Sub

Private Sub Option14_DblClick()
   Call GRAVAMABRUVE(1)
End Sub

Private Sub Option15_DblClick()
   Call GRAPARVENMERC(1)
End Sub

Private Sub Option21_Click()
   Option21.Value = True
   Option21.Refresh
   List1.ListIndex = (-1)
End Sub

Private Sub Option21_DblClick()
   '
   GRUCLI% = 0
   CAMESA% = 24
   '
   For KKI% = 0 To 3
     If Option3(KKI%).Value = True Then
       CAMESA% = 12 * (1 + KKI%)
     End If
   Next KKI%
   '
   Call GRAFUNCICLI(GRUCLI%, CAMESA%, 1)
   '
End Sub

Private Sub Option21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Option21.Value = True
   Option21.Refresh
End Sub

Private Sub Option4_Click()
   Option4.Value = True
   Option4.Refresh
   List2.ListIndex = (-1)
End Sub

Private Sub Option4_DblClick()
   '
   GRUPRO% = 0
   CAMESA% = 24
   '
   For KKI% = 0 To 3
     If Option1(KKI%).Value = True Then
       CAMESA% = 12 * (1 + KKI%)
     End If
   Next KKI%
   '
   Call GRAFUNCIPRO(GRUPRO%, CAMESA%, 1)
   '
End Sub

Private Sub Option4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Option4.Value = True
   Option4.Refresh
End Sub

Private Sub Option5_DblClick()
   Call GRAPARVENPRO(1)
End Sub

Private Sub Option6_DblClick()
  Call GRAVAMABRUPRO(1)
End Sub

Private Sub Option7_DblClick()
   Call GRAPORMABRUPRO(1)
End Sub

Private Sub Option8_DblClick()
   Call GRABARRAPRO(1)
End Sub

Private Sub Option9_DblClick()
   Call TAVENTAPRO(1)
End Sub

Private Sub MERMAPRO_LostFocus()
   If Val(MERMAPRO) > 99.9 Then MERMAPRO = "99.9"
   MERMAPRO = FORMATNUM$(MERMAPRO, "F4.1")
End Sub

'
Private Sub TIPOCAM_LostFocus()
   TIPOCAM = FORMATNUM$(TIPOCAM, "F9.3")
End Sub

Sub PRINTABLA(FINTABLA)
   '
   ALTUPAPEL = Printer.ScaleHeight
   ANCHPAPEL = Printer.ScaleWidth
   '
   BIZQ = 0.1 * ANCHPAPEL
   ANCHO = 0.89 * ANCHPAPEL
   BSUP = 0.03 * ALTUPAPEL
   '
   IGRA% = 0
   '
10 Printer.DrawMode = 1
   Printer.DrawStyle = 0
   Printer.DrawWidth = 4
   Printer.ForeColor = QBColor(0)
   Printer.Line (BIZQ, BSUP)-(BIZQ + ANCHO, BSUP + 350), , B
   Printer.FontName = "Arial" 'GRATABLA.Label2(0).FontName
   Printer.FontSize = 8 'GRATABLA.Label2(0).FontSize
   Printer.FontBold = True   ' GRATABLA.Label2(0).FontBold
   '
   TTXX$ = TRIM$(GRATABLA.Label1(IGRA%))
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 60
   Printer.CurrentY = BSUP + 70
   Printer.Print TTXX$
   '
   COEF = (ANCHO - 400) / Len(GRATABLA.Label2(IGRA%).Caption)
   Printer.CurrentY = BSUP + 590
   TTXX$ = GRATABLA.Label2(IGRA%): GoSub PRIRENGLO
   '
   Printer.CurrentY = BSUP + 900
   For KKI% = 1 To GRATABLA.List1(IGRA%).ListCount
     TTXX$ = GRATABLA.List1(IGRA%).List(KKI% - 1): GoSub PRIRENGLO
     BINFE = Printer.CurrentY
   Next KKI%
   '
   Printer.CurrentY = BINFE + 180
   TTXX$ = GRATABLA.Label3(IGRA%): GoSub PRIRENGLO
   '
   Printer.DrawStyle = 0
   Printer.DrawWidth = 4
   Printer.ForeColor = QBColor(0)
   Printer.Line (BIZQ, BSUP + 870)-(BIZQ + ANCHO, BSUP + 870)
   Printer.Line (BIZQ, BINFE + 120)-(BIZQ + ANCHO, BINFE + 120)
   Printer.Line (BIZQ, BSUP + 530)-(BIZQ + ANCHO, BINFE + 470), , B
   FINTABLA = BINFE + 470
   '
   Printer.DrawWidth = 1
   PX = 21: Printer.Line (BIZQ + PX * COEF, BSUP + 530)-(BIZQ + PX * COEF, BINFE + 470)
   PX = 29: Printer.Line (BIZQ + PX * COEF, BSUP + 530)-(BIZQ + PX * COEF, BINFE + 470)
   PX = 37: Printer.Line (BIZQ + PX * COEF, BSUP + 530)-(BIZQ + PX * COEF, BINFE + 470)
   PX = 51: Printer.Line (BIZQ + PX * COEF, BSUP + 530)-(BIZQ + PX * COEF, BINFE + 470)
   PX = 64: Printer.Line (BIZQ + PX * COEF, BSUP + 530)-(BIZQ + PX * COEF, BINFE + 470)
   PX = 77: Printer.Line (BIZQ + PX * COEF, BSUP + 530)-(BIZQ + PX * COEF, BINFE + 470)
   PX = 85: Printer.Line (BIZQ + PX * COEF, BSUP + 530)-(BIZQ + PX * COEF, BINFE + 470)
   '
   If IGRA% = 0 Then
     BSUP = BINFE + 1000
     IGRA% = 1
     GoTo 10
   End If
   '
Exit Sub

PRIRENGLO:
   TTYY$ = TRIM$(Left$(TTXX$, 20))
   Printer.CurrentX = BIZQ + 60
   Printer.Print TTYY$;
   '
   TTYY$ = TRIM$(Mid$(TTXX$, 21, 6))
   Printer.CurrentX = BIZQ + 60 + 28 * COEF - Printer.TextWidth(TTYY$)
   Printer.Print TTYY$;
   '
   TTYY$ = TRIM$(Mid$(TTXX$, 27, 8))
   Printer.CurrentX = BIZQ + 60 + 36 * COEF - Printer.TextWidth(TTYY$)
   Printer.Print TTYY$;
   '
   TTYY$ = TRIM$(Mid$(TTXX$, 36, 13))
   Printer.CurrentX = BIZQ + 60 + 50 * COEF - Printer.TextWidth(TTYY$)
   Printer.Print TTYY$;
   '
   TTYY$ = TRIM$(Mid$(TTXX$, 49, 13))
   Printer.CurrentX = BIZQ + 60 + 63 * COEF - Printer.TextWidth(TTYY$)
   Printer.Print TTYY$;
   '
   TTYY$ = TRIM$(Mid$(TTXX$, 62, 13))
   Printer.CurrentX = BIZQ + 60 + 76 * COEF - Printer.TextWidth(TTYY$)
   Printer.Print TTYY$;
   '
   TTYY$ = TRIM$(Mid$(TTXX$, 75, 8))
   Printer.CurrentX = BIZQ + 60 + 84 * COEF - Printer.TextWidth(TTYY$)
   Printer.Print TTYY$;
   '
   TTYY$ = TRIM$(Mid$(TTXX$, 85, 8))
   Printer.CurrentX = BIZQ + 60 + 94 * COEF - Printer.TextWidth(TTYY$)
   Printer.Print TTYY$
   '
Return
'
End Sub

Sub CARAPRIN()  ' caratula principal
   '
   ALTUPAPEL = Printer.ScaleHeight
   ANCHPAPEL = Printer.ScaleWidth
   '
   BIZQ = 0.1 * ANCHPAPEL
   ANCHO = 0.89 * ANCHPAPEL
   BSUP = 0.06 * ALTUPAPEL
   '
10 Printer.DrawMode = 1
   Printer.DrawStyle = 0
   Printer.DrawWidth = 4
   Printer.ForeColor = QBColor(0)
   Printer.Line (BIZQ, BSUP)-(BIZQ + ANCHO, 0.97 * ALTUPAPEL), , B
   Printer.Line (BIZQ + 700, 0.8 * ALTUPAPEL)-(BIZQ + ANCHO - 700, 0.8 * ALTUPAPEL + 700), , B
   '
   Printer.FontName = "Arial" 'GRATABLA.Label2(0).FontName
   Printer.FontSize = 48 'GRATABLA.Label2(0).FontSize
   Printer.FontBold = False   ' GRATABLA.Label2(0).FontBold
   '
   TTXX$ = "Informe de"
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 100
   Printer.CurrentY = BSUP + 0.18 * ALTUPAPEL
   Printer.Print TTXX$
   '
   TTXX$ = "Gestión de Ventas"
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 100
   Printer.CurrentY = BSUP + 0.3 * ALTUPAPEL
   Printer.Print TTXX$
   '
   Printer.FontName = "Arial" 'GRATABLA.Label2(0).FontName
   Printer.FontSize = 28 'GRATABLA.Label2(0).FontSize
   Printer.FontBold = True   ' GRATABLA.Label2(0).FontBold
   '
   TTXX$ = ABREMESAN$(CODIMES% - 11, 1) + " - " + ABREMESAN$(CODIMES%, 1)
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 100
   Printer.CurrentY = BSUP + 0.52 * ALTUPAPEL
   Printer.Print TTXX$
   '
   Printer.FontName = "Arial" 'GRATABLA.Label2(0).FontName
   Printer.FontSize = 10 'GRATABLA.Label2(0).FontSize
   Printer.FontBold = True   ' GRATABLA.Label2(0).FontBold
   '
   HOII% = HOY(HOS$)
   TTXX$ = "(Emitido el " + FETEXTO$(HOII%) + ")"
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 100
   Printer.CurrentY = BSUP + 0.56 * ALTUPAPEL
   Printer.Print TTXX$
   '
   Printer.NewPage
   '
End Sub

Sub CARAMERC()  ' caratula division por mercado
   '
   ALTUPAPEL = Printer.ScaleHeight
   ANCHPAPEL = Printer.ScaleWidth
   '
   BIZQ = 0.1 * ANCHPAPEL
   ANCHO = 0.89 * ANCHPAPEL
   BSUP = 0.06 * ALTUPAPEL
   '
10 Printer.DrawMode = 1
   Printer.DrawStyle = 0
   Printer.DrawWidth = 4
   Printer.ForeColor = QBColor(0)
   Printer.Line (BIZQ, BSUP)-(BIZQ + ANCHO, 0.97 * ALTUPAPEL), , B
   Printer.Line (BIZQ + 700, 0.15 * ALTUPAPEL)-(BIZQ + ANCHO - 700, 0.49 * ALTUPAPEL), , B
   Printer.Line (BIZQ + 700, 0.8 * ALTUPAPEL)-(BIZQ + ANCHO - 700, 0.8 * ALTUPAPEL + 700), , B
   '
   Printer.FontName = "Arial" 'GRATABLA.Label2(0).FontName
   Printer.FontSize = 48 'GRATABLA.Label2(0).FontSize
   Printer.FontBold = False   ' GRATABLA.Label2(0).FontBold
   '
   TTXX$ = "División por"
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 100
   Printer.CurrentY = BSUP + 0.18 * ALTUPAPEL
   Printer.Print TTXX$
   '
   TTXX$ = "Mercado"
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 100
   Printer.CurrentY = BSUP + 0.3 * ALTUPAPEL
   Printer.Print TTXX$
   '
   Printer.FontName = "Arial" 'GRATABLA.Label2(0).FontName
   Printer.FontSize = 28 'GRATABLA.Label2(0).FontSize
   Printer.FontBold = True   ' GRATABLA.Label2(0).FontBold
   '
   TTXX$ = ABREMESAN$(CODIMES% - 11, 1) + " - " + ABREMESAN$(CODIMES%, 1)
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 100
   Printer.CurrentY = BSUP + 0.52 * ALTUPAPEL
   Printer.Print TTXX$
   '
   Printer.FontName = "Arial" 'GRATABLA.Label2(0).FontName
   Printer.FontSize = 10 'GRATABLA.Label2(0).FontSize
   Printer.FontBold = True   ' GRATABLA.Label2(0).FontBold
   '
   HOII% = HOY(HOS$)
   TTXX$ = "(Emitido el " + FETEXTO$(HOII%) + ")"
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 100
   Printer.CurrentY = BSUP + 0.56 * ALTUPAPEL
   Printer.Print TTXX$
   '
   Printer.NewPage
   '
End Sub

Sub CARAPROD()  ' caratula division por PRODUCTO
   '
   ALTUPAPEL = Printer.ScaleHeight
   ANCHPAPEL = Printer.ScaleWidth
   '
   BIZQ = 0.1 * ANCHPAPEL
   ANCHO = 0.89 * ANCHPAPEL
   BSUP = 0.06 * ALTUPAPEL
   '
10 Printer.DrawMode = 1
   Printer.DrawStyle = 0
   Printer.DrawWidth = 4
   Printer.ForeColor = QBColor(0)
   Printer.Line (BIZQ, BSUP)-(BIZQ + ANCHO, 0.97 * ALTUPAPEL), , B
   Printer.Line (BIZQ + 700, 0.15 * ALTUPAPEL)-(BIZQ + ANCHO - 700, 0.49 * ALTUPAPEL), , B
   Printer.Line (BIZQ + 700, 0.8 * ALTUPAPEL)-(BIZQ + ANCHO - 700, 0.8 * ALTUPAPEL + 700), , B
   '
   Printer.FontName = "Arial" 'GRATABLA.Label2(0).FontName
   Printer.FontSize = 48 'GRATABLA.Label2(0).FontSize
   Printer.FontBold = False   ' GRATABLA.Label2(0).FontBold
   '
   TTXX$ = "División por"
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 100
   Printer.CurrentY = BSUP + 0.18 * ALTUPAPEL
   Printer.Print TTXX$
   '
   TTXX$ = "Producto"
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 100
   Printer.CurrentY = BSUP + 0.3 * ALTUPAPEL
   Printer.Print TTXX$
   '
   Printer.FontName = "Arial" 'GRATABLA.Label2(0).FontName
   Printer.FontSize = 28 'GRATABLA.Label2(0).FontSize
   Printer.FontBold = True   ' GRATABLA.Label2(0).FontBold
   '
   TTXX$ = ABREMESAN$(CODIMES% - 11, 1) + " - " + ABREMESAN$(CODIMES%, 1)
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 100
   Printer.CurrentY = BSUP + 0.52 * ALTUPAPEL
   Printer.Print TTXX$
   '
   Printer.FontName = "Arial" 'GRATABLA.Label2(0).FontName
   Printer.FontSize = 10 'GRATABLA.Label2(0).FontSize
   Printer.FontBold = True   ' GRATABLA.Label2(0).FontBold
   '
   HOII% = HOY(HOS$)
   TTXX$ = "(Emitido el " + FETEXTO$(HOII%) + ")"
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 100
   Printer.CurrentY = BSUP + 0.56 * ALTUPAPEL
   Printer.Print TTXX$
   '
   Printer.NewPage
   '
End Sub

Function FACONVERT(COMESX%)
   Static CTX%, TICADO(2048)
   If CTX% = 0 Then
     CTX% = 1
     For KKI% = 1 To 2048
       TICADO(KKI%) = 1
     Next KKI%
     For U& = 1 To ULTREG&("IRESVENT")
       TTT$ = REGLEIDO$("IRESVENT", U&)
       COMESY% = CVI(Left$(TTT$, 2))
       TICADY = CVD(Mid$(TTT$, 65, 8))
       If COMESY% > 0 Then
         If COMESY% <= 2048 Then
           If TICADY > 0.5 Then
             TICADO(COMESY%) = TICADY
           End If
         End If
       End If
     Next U&
   End If
   '
   MONECO% = 1
   If Option10.Value = True Then
     MONECO% = 2
     FACONVERT = 1
     If COMESX% > 0 Then
       If COMESX% <= 2048 Then
         FACONVERT = TICADO(COMESX%)
         If COMESX% = CODIMES% Then
           FACONVERT = Val(TIPOCAM)
         End If
       End If
     End If
   End If
   '
   If FACONVERT < 0.5 Then
     FACONVERT = 1
   End If
   '
End Function

Sub REVEPRO1()
    '
    Dim TOTVENTA#(32767), TOTCAN(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    J& = 0
    '
    FIN& = ULTREG&("DVECLART")
    Screen.MousePointer = 11
    TOTAGEN# = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("DVECLART", U&)
      FF% = CVI(Mid$(X$, 49, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          CI% = CVI(Mid$(X$, 7, 2))
          If CI% > 0 Then
            If CI% <= 32767 Then
              CMES% = CVI(Left$(X$, 2))
              CANTI = CVD(Mid$(X$, 51, 8))
              NETO# = CVD(Mid$(X$, 169, 8)) / FACONVERT(CMES%)
              TOTCAN(CI%) = TOTCAN(CI%) + CANTI
              TOTVENTA#(CI%) = TOTVENTA#(CI%) + NETO#
            End If
          End If
        End If
      End If
    Next U&
    '
    J& = 0
    For U& = 1 To NUINV%
      XXX$ = REGLEIDO$("INVERT", U& + 1)
      NCI% = CVI(Mid$(XXX$, 17, 2))
      If NCI% > 0 Then
        If NCI% <= NUMAS% Then
          If Abs(TOTCAN(NCI%)) > 0.005 Then
            RELIB$ = REGBLAN$("REVEPRO")
            Call REPLA(RELIB$, MKI$(NCI%), 1, 2)
            Call REPLA(RELIB$, MKS$(TOTVENTA#(NCI%) / TOTCAN(NCI%)), 3, 4)
            Call REPLA(RELIB$, MKS$(TOTCAN(NCI%)), 7, 4)
            Call REPLA(RELIB$, MKD$(TOTVENTA#(NCI%)), 11, 8)
            J& = J& + 1
            Call GRAREG("REVEPRO", RELIB$, J&)
          End If
        End If
      End If
    Next U&
    '
    Call SETULTREG("REVEPRO", J&)
    Call CIERRARCH("REVEPRO")
    '
    MONEC$ = "Pesos": If Option10.Value = True Then MONEC$ = "Dolares"
    Call HEADERS("REVEPRO", "")
    Call HEADERS("REVEPRO", "Periodo: " + PERIO$ + " - Moneda: " + MONEC$)
    Screen.MousePointer = 1
    '
    Call FINAL("*REVEPRO")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
'
Sub RESUENCLI1()
    '
    Dim TOCAEN$(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    '
20  VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 20
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    JJ& = 0
    '
    FIN& = ULTREG&("DVECLART")
    Screen.MousePointer = 11
    TOTAGEN# = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("DVECLART", U&)
      FF% = CVI(Mid$(X$, 49, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          CLI% = CVI(Mid$(X$, 5, 2))
          CII% = CVI(Mid$(X$, 7, 2))
          If CII% > 0 Then
            If CII% <= NUMAS% Then
              CMES% = CVI(Left$(X$, 2))
              CANTU = CVD(Mid$(X$, 51, 8))
              NETO# = CVD(Mid$(X$, 169, 8)) / FACONVERT(CMES%)
              '
              X$ = REGBLAN$("RESENCLI")
              Call REPLA(X$, MKI$(CLI%), 1, 2)
              Call REPLA(X$, MKI$(CII%), 3, 2)
              Call REPLA(X$, MKS$(CANTU), 5, 4)
              Call REPLA(X$, MKD$(NETO#), 9, 8)
              Call REPLA(X$, PERIO$, 25, 20)
              JJ& = JJ& + 1
              Call GRAREG("RESENCLI", X$, JJ&)
              '
            End If
          End If
        End If
      End If
    Next U&
    '
    Call SETULTREG("RESENCLI", JJ&)
    Call CIERRARCH("RESENCLI")
    MONEC$ = "Pesos": If Option10.Value = True Then MONEC$ = "Dolares"
    Call HEADERS("RESENCLI", "")
    Call HEADERS("RESENCLI", "Corresponde a: " + PERIO$ + " - Moneda: " + MONEC$)
    Screen.MousePointer = 1
    '
    Call FINAL("*SORENCL")
    '
99  Call CIERRARCH("*.*")
    '
End Sub

Sub RANKCLI1()
    '
    Dim TOTVENTA#(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    J& = 0
    '
    FIN& = ULTREG&("DVECLART")
    Screen.MousePointer = 11
    TOTAGEN# = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("DVECLART", U&)
      FF% = CVI(Mid$(X$, 49, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          CLI% = CVI(Mid$(X$, 5, 2))
          If CLI% > 0 Then
            If CLI% <= 32767 Then
              CMES% = CVI(Left$(X$, 2))
              CANTU = CVD(Mid$(X$, 51, 8))
              NETO# = CVD(Mid$(X$, 169, 8)) / FACONVERT(CMES%)
              '
              TOTVENTA#(CLI%) = TOTVENTA#(CLI%) + NETO#
              TOTAGEN# = TOTAGEN# + NETO#
            End If
          End If
        End If
      End If
    Next U&
    '
    List3.Clear
    For NCI% = 1 To 32766
      If Abs(TOTVENTA#(NCI%)) > 0.005 Then
        RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCI%)), 4, 12, 2, 2)
        PORPAR = 100 * TOTVENTA#(NCI%) / TOTAGEN#
        RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
        RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
        RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCI%)), 4, 12, 2, 2)
        List3.AddItem RELIB$
      End If
    Next NCI%
    '
    J& = 0
    FIN& = List3.ListCount
    '
    For U& = 1 To FIN&
      List3.ListIndex = U& - 1
      NC% = Val(Mid$(List3.TEXT, 14, 6))
      TONEVEN# = Val(Mid$(List3.TEXT, 28, 12))
      PORPAR = Val(Mid$(List3.TEXT, 20, 8))
      PORACU = PORACU + PORPAR
      '
      Z$ = REGBLAN$("RANKCLI")
      J& = J& + 1: JI% = J&
      Call REPLA(Z$, MKI$(JI%), 1, 2)
      Call REPLA(Z$, MKI$(NC%), 3, 2)
      Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
      Call REPLA(Z$, MKS$(PORPAR), 13, 4)
      Call REPLA(Z$, MKS$(PORACU), 17, 4)
      Call GRAREG("RANKCLI", Z$, J&)
      '
    Next U&
    '
    Call SETULTREG("RANKCLI", J&)
    Call CIERRARCH("RANKCLI")
    Call HEADERS("RANKCLI", "")
    MONEC$ = "Pesos": If Option10.Value = True Then MONEC$ = "Dolares"
    Call HEADERS("RANKCLI", "Periodo: " + PERIO$ + " - Moneda: " + MONEC$)
    Screen.MousePointer = 1
    List3.Clear
    '
    Call FINAL("*LRANCLI")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
'

Sub RANKPIE1()
    '
    Dim TOTVENTA#(32767), TOTCOSTO#(32767), CANPIE(32767), COSTOUN(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("DVECLART")
    Screen.MousePointer = 11
    TOTAGEN# = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("DVECLART", U&)
      FF% = CVI(Mid$(X$, 49, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          CII% = CVI(Mid$(X$, 7, 2))
          If CII% > 0 Then
            If CII% <= NUMAS% Then
              CMES% = CVI(Left$(X$, 2))
              CANTU = CVD(Mid$(X$, 51, 8))
              NETO# = CVD(Mid$(X$, 169, 8)) / FACONVERT(CMES%)
              TOTVENTA#(CII%) = TOTVENTA#(CII%) + NETO#
              NECO# = CVD(Mid$(X$, 193, 8)) / FACONVERT(CMES%)
              TOTCOSTO#(CII%) = TOTCOSTO#(CII%) + NECO#
              CANPIE(CII%) = CANPIE(CII%) + CANTU
              COSTOUN(CII%) = CVD(Mid$(X$, 185, 8)) / FACONVERT(CMES%)
              TOTAGEN# = TOTAGEN# + NETO#
            End If
          End If
        End If
      End If
    Next U&
    '
    List3.Clear
    For NCI% = 1 To NUMAS%
       If Abs(TOTVENTA#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTVENTA#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCI%)), 4, 12, 2, 2)
          COSUNPIE = COSTOUN(NCI%)
          If CANPIE(NCI%) > 0 Then
            COSUNPIE = TOTCOSTO(NCI%) / CANPIE(NCI%)
          End If
          RELIB$ = RELIB$ + CSTRING$(MKS$(COSUNPIE), 4, 12, 4, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTCOSTO#(NCI%)), 4, 12, 2, 2)
          List3.AddItem RELIB$
       End If
    Next NCI%
    '
    J& = 0
    FIN& = List3.ListCount
    '
    For U& = 1 To FIN&
        List3.ListIndex = U& - 1
        CI% = Val(Mid$(List3.TEXT, 14, 6))
        TONEVEN# = Val(Mid$(List3.TEXT, 28, 12))
        COSUNPIE = Val(Mid$(List3.TEXT, 40, 12))
        TOTACOS# = Val(Mid$(List3.TEXT, 52, 12))
        PORPAR = Val(Mid$(List3.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKPIE")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(CI%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call REPLA(Z$, MKS$(COSUNPIE), 21, 4)
        Call REPLA(Z$, MKD$(TOTACOS#), 25, 8)
        Call GRAREG("RANKPIE", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKPIE", J&)
    Call CIERRARCH("RANKPIE")
    Call HEADERS("RANKPIE", "")
    MONEC$ = "Pesos": If Option10.Value = True Then MONEC$ = "Dolares"
    Call HEADERS("RANKPIE", "Periodo: " + PERIO$ + " - Moneda: " + MONEC$)
    Screen.MousePointer = 1
    List3.Clear
    '
    Call FINAL("*LRANPIE")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
'
Sub PRITARES()
   '
   Dim VALOTOTA#(64, 4, 13)
   Dim ORGRU%(32767)
   Dim DEGRU$(64), NUGRU%(64)
   '
   Dim DERU$(4)
   DERU$(1) = "Ventas"
   DERU$(2) = "CMV"
   DERU$(3) = "Margen B"
   DERU$(4) = "% Mrg.Br."
   '
   HACODI% = CODIMES%
   DECODI% = HACODI% - 11
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   For U& = 1 To ULTREG&("GRUCOVEN")
     X$ = REGLEIDO$("GRUCOVEN", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     NUGRU%(CAGRU%) = GRUCO%
     ORGRU%(GRUCO%) = CAGRU%
     ULTGRU% = GRUCO%
   Next U&
   '
   DEGRU$(CAGRU% + 1) = "Total"
   '
   FIN& = ULTREG&("PVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("PVECLART", U&)
      CMES% = CVI(Left$(Z$, 2))
      If CMES% >= CODIMES% - 11 Then
        If CMES% <= CODIMES% Then
          CI% = CVI(Mid$(Z$, 7, 2))
          GRUCO% = GRCV%(CI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(CMES%)
          '
          VALOR# = CVD(Mid$(Z$, 169, 8)) / TICAM
          VALOTOTA#(ORGRU%(GRUCO%), 1, 1 + CMES% - DECODI%) = VALOTOTA#(ORGRU%(GRUCO%), 1, 1 + CMES% - DECODI%) + VALOR#
          VALOTOTA#(ORGRU%(GRUCO%), 1, 13) = VALOTOTA#(ORGRU%(GRUCO%), 1, 13) + VALOR#
          VALOTOTA#(CAGRU% + 1, 1, 1 + CMES% - DECODI%) = VALOTOTA#(CAGRU% + 1, 1, 1 + CMES% - DECODI%) + VALOR#
          VALOTOTA#(CAGRU% + 1, 1, 13) = VALOTOTA#(CAGRU% + 1, 1, 13) + VALOR#
          '
          VALOR# = CVD(Mid$(Z$, 193, 8)) / TICAM
          VALOTOTA#(ORGRU%(GRUCO%), 2, 1 + CMES% - DECODI%) = VALOTOTA#(ORGRU%(GRUCO%), 2, 1 + CMES% - DECODI%) + VALOR#
          VALOTOTA#(ORGRU%(GRUCO%), 2, 13) = VALOTOTA#(ORGRU%(GRUCO%), 2, 13) + VALOR#
          VALOTOTA#(CAGRU% + 1, 2, 1 + CMES% - DECODI%) = VALOTOTA#(CAGRU% + 1, 2, 1 + CMES% - DECODI%) + VALOR#
          VALOTOTA#(CAGRU% + 1, 2, 13) = VALOTOTA#(CAGRU% + 1, 2, 13) + VALOR#
          '
          VALOR# = (CVD(Mid$(Z$, 169, 8)) - CVD(Mid$(Z$, 193, 8))) / TICAM
          VALOTOTA#(ORGRU%(GRUCO%), 3, 1 + CMES% - DECODI%) = VALOTOTA#(ORGRU%(GRUCO%), 3, 1 + CMES% - DECODI%) + VALOR#
          VALOTOTA#(ORGRU%(GRUCO%), 3, 13) = VALOTOTA#(ORGRU%(GRUCO%), 3, 13) + VALOR#
          VALOTOTA#(CAGRU% + 1, 3, 1 + CMES% - DECODI%) = VALOTOTA#(CAGRU% + 1, 3, 1 + CMES% - DECODI%) + VALOR#
          VALOTOTA#(CAGRU% + 1, 3, 13) = VALOTOTA#(CAGRU% + 1, 3, 13) + VALOR#
          '
          VALOR# = 0
          If VALOTOTA#(ORGRU%(GRUCO%), 1, 1 + CMES% - DECODI%) > 1 Then
            VALOR# = 100 * (VALOTOTA#(ORGRU%(GRUCO%), 3, 1 + CMES% - DECODI%)) / VALOTOTA#(ORGRU%(GRUCO%), 1, 1 + CMES% - DECODI%)
            VALOTOTA#(ORGRU%(GRUCO%), 4, 1 + CMES% - DECODI%) = VALOR#
            VALOTOTA#(ORGRU%(GRUCO%), 4, 13) = VALOR#
            VALOTOTA#(CAGRU% + 1, 4, 1 + CMES% - DECODI%) = VALOR#
          End If
        End If
      End If
   Next U&
   '
   PORIAN = Printer.Orientation
   '
   Printer.Orientation = 2
   ALTUPAPEL = Printer.ScaleHeight
   ANCHPAPEL = Printer.ScaleWidth
   '
   BIZQ = 0.01 * ANCHPAPEL
   ANCHO = 0.98 * ANCHPAPEL
   BSUP = 0.03 * ALTUPAPEL
   '
   IGRA% = 0
   INIHOJA% = 1
   '
10 Printer.DrawMode = 1
   Printer.DrawStyle = 0
   Printer.DrawWidth = 4
   Printer.ForeColor = QBColor(0)
   Printer.Line (BIZQ, BSUP)-(BIZQ + ANCHO + 80, BSUP + 450), , B
   Printer.FontName = "Arial" 'GRATABLA.Label2(0).FontName
   Printer.FontSize = 8 'GRATABLA.Label2(0).FontSize
   Printer.FontBold = True   ' GRATABLA.Label2(0).FontBold
   '
   TITUMO$ = "PESOS"
   If Option10.Value = True Then
      TITUMO$ = "U$S"
   End If
   '
   Printer.FontSize = 16 'GRATABLA.Label2(0).FontSize
   TTXX$ = "DIVISION POR PRODUCTO - Período "
   TTXX$ = TTXX$ + ABREMESAN$(DECODI%, 1) + " - " + ABREMESAN$(HACODI%, 1) + " - " + TITUMO$
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 60
   Printer.CurrentY = BSUP + 70
   Printer.Print TTXX$
   '
   ACAMPO = (ANCHO - 2000) / 13
   '
   Printer.FontSize = 8 'GRATABLA.Label2(0).FontSize
   Printer.CurrentY = BSUP + 900
   Printer.DrawWidth = 4
   PVERT = Printer.CurrentY
   Printer.Line (BIZQ, BSUP + 650)-(BIZQ + 2040 + 13 * ACAMPO, BSUP + 650)
   Printer.Line (BIZQ, PVERT + 15)-(BIZQ + 2040 + 13 * ACAMPO, PVERT + 15)
   Printer.CurrentY = PVERT + 30
   For I% = INIHOJA% To CAGRU% + 1
     INIRE = Printer.CurrentY
     Printer.CurrentX = BIZQ + 40
     Printer.Print Left$(DEGRU$(I%), 11);
     For J% = 1 To 4
       Printer.CurrentX = BIZQ + 1290
       Printer.Print DERU$(J%);
       For K% = 1 To 13
         If J% < 4 Then
              VALOR# = VALOTOTA#(I%, J%, K%)
              TTYY$ = TRIM$(FORMATNUM$(VALOR#, "F12.2"))
            Else
              VALOR# = 0
              If VALOTOTA#(I%, 1, K%) > 1 Then
                VALOR# = 100 * VALOTOTA#(I%, 3, K%) / VALOTOTA#(I%, 1, K%)
              End If
              TTYY$ = TRIM$(FORMATNUM$(VALOR#, "F12.1"))
              If TTYY$ <> "" Then TTYY$ = TTYY$ + " %"
         End If
         Printer.CurrentX = BIZQ + 2040 + ACAMPO * K% - Printer.TextWidth(TTYY$)
         Printer.Print TTYY$;
       Next K%
       Printer.Print ""
       If J% < 4 Then
         Printer.DrawWidth = 2
         PVERT = Printer.CurrentY
         Printer.Line (BIZQ + 1250, PVERT + 15)-(BIZQ + 2070 + 13 * ACAMPO, PVERT + 15)
         Printer.CurrentY = PVERT + 30
       End If
     Next J%
     Printer.DrawWidth = 3
     PVERT = Printer.CurrentY
     Printer.Line (BIZQ, PVERT + 15)-(BIZQ + 2070 + 13 * ACAMPO, PVERT + 15)
     Printer.CurrentY = PVERT + 30
     '
     ARETOT = Printer.CurrentY - INIRE
     If I% < CAGRU% + 1 Then
       If Printer.CurrentY + ARETOT > ALTUPAPEL - 300 Then
         INIHOJA% = I% + 1
         GoTo 29
       End If
     End If
   Next I%
   INIHOJA% = 0
   '
29 Printer.DrawWidth = 4
   Printer.Line (BIZQ, BSUP + 650)-(BIZQ, PVERT + 15)
   Printer.DrawWidth = 3
   Printer.Line (BIZQ + 1250, BSUP + 650)-(BIZQ + 1250, PVERT + 15)
   Printer.Line (BIZQ + 2050, BSUP + 650)-(BIZQ + 2050, PVERT + 15)
   '
   For K% = 1 To 13
     TTYY$ = ABREMESAN$(DECODI% + K% - 1, 2)
     If K% = 13 Then TTYY$ = "Total"
     Printer.CurrentY = BSUP + 720
     Printer.CurrentX = BIZQ + 2040 + ACAMPO * K% - Printer.TextWidth(TTYY$)
     Printer.Print TTYY$;
     If K% = 13 Then Printer.DrawWidth = 4
     Printer.Line (BIZQ + 2080 + ACAMPO * K%, BSUP + 650)-(BIZQ + 2080 + ACAMPO * K%, PVERT + 15)
   Next K%
   '
   If INIHOJA% > 0 Then
     Printer.NewPage
     GoTo 10
   End If
   '
   Printer.EndDoc
   Printer.Orientation = PORIAN
   '
End Sub

Sub PRITAMERC()
   '
   Dim VALOTOTA#(64, 4, 13)
   Dim ORGRU%(32767)
   Dim DEGRU$(64), NUGRU%(64)
   '
   Dim DERU$(4)
   DERU$(1) = "Ventas"
   DERU$(2) = "CMV"
   DERU$(3) = "Margen B"
   DERU$(4) = "% Mrg.Br."
   '
   HACODI% = CODIMES%
   DECODI% = HACODI% - 11
   '
   VEGRUVE$ = ""
   CAGRU% = 0
   For U& = 1 To ULTREG&("GRUCOCLI")
     X$ = REGLEIDO$("GRUCOCLI", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
     CAGRU% = CAGRU% + 1
     DEGRU$(CAGRU%) = TRIM$(Mid$(X$, 3, 30))
     NUGRU%(CAGRU%) = GRUCO%
     ORGRU%(GRUCO%) = CAGRU%
     ULTGRU% = GRUCO%
   Next U&
   '
   DEGRU$(CAGRU% + 1) = "Total"
   '
   FIN& = ULTREG&("CVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("CVECLART", U&)
      CMES% = CVI(Left$(Z$, 2))
      If CMES% >= CODIMES% - 11 Then
        If CMES% <= CODIMES% Then
          CLI% = CVI(Mid$(Z$, 5, 2))
          GRUCO% = GRCC%(CLI%)
          If GRUCO% < 1 Then GRUCO% = ULTGRU%
          If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = ULTGRU%
          TICAM = FACONVERT(CMES%)
          '
          VALOR# = CVD(Mid$(Z$, 169, 8)) / TICAM
          VALOTOTA#(ORGRU%(GRUCO%), 1, 1 + CMES% - DECODI%) = VALOTOTA#(ORGRU%(GRUCO%), 1, 1 + CMES% - DECODI%) + VALOR#
          VALOTOTA#(ORGRU%(GRUCO%), 1, 13) = VALOTOTA#(ORGRU%(GRUCO%), 1, 13) + VALOR#
          VALOTOTA#(CAGRU% + 1, 1, 1 + CMES% - DECODI%) = VALOTOTA#(CAGRU% + 1, 1, 1 + CMES% - DECODI%) + VALOR#
          VALOTOTA#(CAGRU% + 1, 1, 13) = VALOTOTA#(CAGRU% + 1, 1, 13) + VALOR#
          '
          VALOR# = CVD(Mid$(Z$, 193, 8)) / TICAM
          VALOTOTA#(ORGRU%(GRUCO%), 2, 1 + CMES% - DECODI%) = VALOTOTA#(ORGRU%(GRUCO%), 2, 1 + CMES% - DECODI%) + VALOR#
          VALOTOTA#(ORGRU%(GRUCO%), 2, 13) = VALOTOTA#(ORGRU%(GRUCO%), 2, 13) + VALOR#
          VALOTOTA#(CAGRU% + 1, 2, 1 + CMES% - DECODI%) = VALOTOTA#(CAGRU% + 1, 2, 1 + CMES% - DECODI%) + VALOR#
          VALOTOTA#(CAGRU% + 1, 2, 13) = VALOTOTA#(CAGRU% + 1, 2, 13) + VALOR#
          '
          VALOR# = (CVD(Mid$(Z$, 169, 8)) - CVD(Mid$(Z$, 193, 8))) / TICAM
          VALOTOTA#(ORGRU%(GRUCO%), 3, 1 + CMES% - DECODI%) = VALOTOTA#(ORGRU%(GRUCO%), 3, 1 + CMES% - DECODI%) + VALOR#
          VALOTOTA#(ORGRU%(GRUCO%), 3, 13) = VALOTOTA#(ORGRU%(GRUCO%), 3, 13) + VALOR#
          VALOTOTA#(CAGRU% + 1, 3, 1 + CMES% - DECODI%) = VALOTOTA#(CAGRU% + 1, 3, 1 + CMES% - DECODI%) + VALOR#
          VALOTOTA#(CAGRU% + 1, 3, 13) = VALOTOTA#(CAGRU% + 1, 3, 13) + VALOR#
          '
          VALOR# = 0
          If VALOTOTA#(ORGRU%(GRUCO%), 1, 1 + CMES% - DECODI%) > 1 Then
            VALOR# = 100 * (VALOTOTA#(ORGRU%(GRUCO%), 3, 1 + CMES% - DECODI%)) / VALOTOTA#(ORGRU%(GRUCO%), 1, 1 + CMES% - DECODI%)
            VALOTOTA#(ORGRU%(GRUCO%), 4, 1 + CMES% - DECODI%) = VALOR#
            VALOTOTA#(ORGRU%(GRUCO%), 4, 13) = VALOR#
            VALOTOTA#(CAGRU% + 1, 4, 1 + CMES% - DECODI%) = VALOR#
          End If
        End If
      End If
   Next U&
   '
   PORIAN = Printer.Orientation
   '
   Printer.Orientation = 2
   ALTUPAPEL = Printer.ScaleHeight
   ANCHPAPEL = Printer.ScaleWidth
   '
   BIZQ = 0.01 * ANCHPAPEL
   ANCHO = 0.98 * ANCHPAPEL
   BSUP = 0.03 * ALTUPAPEL
   '
   IGRA% = 0
   INIHOJA% = 1
   '
10 Printer.DrawMode = 1
   Printer.DrawStyle = 0
   Printer.DrawWidth = 4
   Printer.ForeColor = QBColor(0)
   Printer.Line (BIZQ, BSUP)-(BIZQ + ANCHO + 80, BSUP + 450), , B
   Printer.FontName = "Arial" 'GRATABLA.Label2(0).FontName
   Printer.FontSize = 8 'GRATABLA.Label2(0).FontSize
   Printer.FontBold = True   ' GRATABLA.Label2(0).FontBold
   '
   TITUMO$ = "PESOS"
   If Option10.Value = True Then
      TITUMO$ = "U$S"
   End If
   '
   Printer.FontSize = 16 'GRATABLA.Label2(0).FontSize
   TTXX$ = "DIVISION POR MERCADO - Período "
   TTXX$ = TTXX$ + ABREMESAN$(DECODI%, 1) + " - " + ABREMESAN$(HACODI%, 1) + " - " + TITUMO$
   While Printer.TextWidth(TTXX$) < ANCHO
     TTXX$ = " " + TTXX$ + " "
   Wend
   TTXX$ = Mid$(TTXX$, 2, Len(TTXX$) - 2)
   Printer.CurrentX = BIZQ + 60
   Printer.CurrentY = BSUP + 70
   Printer.Print TTXX$
   '
   ACAMPO = (ANCHO - 2000) / 13
   '
   Printer.FontSize = 8 'GRATABLA.Label2(0).FontSize
   Printer.CurrentY = BSUP + 900
   Printer.DrawWidth = 4
   PVERT = Printer.CurrentY
   Printer.Line (BIZQ, BSUP + 650)-(BIZQ + 2040 + 13 * ACAMPO, BSUP + 650)
   Printer.Line (BIZQ, PVERT + 15)-(BIZQ + 2040 + 13 * ACAMPO, PVERT + 15)
   Printer.CurrentY = PVERT + 30
   For I% = INIHOJA% To CAGRU% + 1
     INIRE = Printer.CurrentY
     Printer.CurrentX = BIZQ + 40
     Printer.Print Left$(DEGRU$(I%), 11);
     For J% = 1 To 4
       Printer.CurrentX = BIZQ + 1290
       Printer.Print DERU$(J%);
       For K% = 1 To 13
         If J% < 4 Then
              VALOR# = VALOTOTA#(I%, J%, K%)
              TTYY$ = TRIM$(FORMATNUM$(VALOR#, "F12.2"))
            Else
              VALOR# = 0
              If VALOTOTA#(I%, 1, K%) > 1 Then
                VALOR# = 100 * VALOTOTA#(I%, 3, K%) / VALOTOTA#(I%, 1, K%)
              End If
              TTYY$ = TRIM$(FORMATNUM$(VALOR#, "F12.1"))
              If TTYY$ <> "" Then TTYY$ = TTYY$ + " %"
         End If
         Printer.CurrentX = BIZQ + 2040 + ACAMPO * K% - Printer.TextWidth(TTYY$)
         Printer.Print TTYY$;
       Next K%
       Printer.Print ""
       If J% < 4 Then
         Printer.DrawWidth = 2
         PVERT = Printer.CurrentY
         Printer.Line (BIZQ + 1250, PVERT + 15)-(BIZQ + 2070 + 13 * ACAMPO, PVERT + 15)
         Printer.CurrentY = PVERT + 30
       End If
     Next J%
     Printer.DrawWidth = 3
     PVERT = Printer.CurrentY
     Printer.Line (BIZQ, PVERT + 15)-(BIZQ + 2070 + 13 * ACAMPO, PVERT + 15)
     Printer.CurrentY = PVERT + 30
     ARETOT = Printer.CurrentY - INIRE
     '
     If I% < CAGRU% + 1 Then
       If Printer.CurrentY + ARETOT > ALTUPAPEL - 300 Then
         INIHOJA% = I% + 1
         GoTo 29
       End If
     End If
     '
   Next I%
   INIHOJA% = 0
   '
29 Printer.DrawWidth = 4
   Printer.Line (BIZQ, BSUP + 650)-(BIZQ, PVERT + 15)
   Printer.DrawWidth = 3
   Printer.Line (BIZQ + 1250, BSUP + 650)-(BIZQ + 1250, PVERT + 15)
   Printer.Line (BIZQ + 2050, BSUP + 650)-(BIZQ + 2050, PVERT + 15)
   '
   For K% = 1 To 13
     TTYY$ = ABREMESAN$(DECODI% + K% - 1, 2)
     If K% = 13 Then TTYY$ = "Total"
     Printer.CurrentY = BSUP + 720
     Printer.CurrentX = BIZQ + 2040 + ACAMPO * K% - Printer.TextWidth(TTYY$)
     Printer.Print TTYY$;
     If K% = 13 Then Printer.DrawWidth = 4
     Printer.Line (BIZQ + 2080 + ACAMPO * K%, BSUP + 650)-(BIZQ + 2080 + ACAMPO * K%, PVERT + 15)
   Next K%
   '
   If INIHOJA% > 0 Then
     Printer.NewPage
     GoTo 10
   End If
   '
   Printer.EndDoc
   Printer.Orientation = PORIAN
   '
End Sub

Sub CARCOSINFO()
   '
   CDMLIM% = 12 * 101 + 12
   CDMX% = CODIMES%: If CDMX% < CDMLIM% Then CDMX% = CDMLIM%
   '
   For KKI% = 1 To 32766
     COREGIS#(KKI%) = 0
   Next KKI%
   '
   RFF$ = RECFILT$("PVECLART", 1, MKI$(CODIMES%))
   For U& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, U&, 4))
     RELE$ = REGLEIDO$("PVECLART", REG&)
     CII% = CVI(Mid$(RELE$, 7, 2))
     If CII% > 0 Then
       If CII% <= NUMAS% Then
         COREGIS#(CII%) = CVD(Mid$(RELE$, 185, 8))
       End If
     End If
   Next U&
   '
   If CDMX% <> CODIMES% Then
     RFF$ = RECFILT$("PVECLART", 1, MKI$(CDMX%))
     For U& = 1 To Len(RFF$) Step 4
       REG& = CVS(Mid$(RFF$, U&, 4))
       RELE$ = REGLEIDO$("PVECLART", REG&)
       CII% = CVI(Mid$(RELE$, 7, 2))
       If CII% > 0 Then
         If CII% <= NUMAS% Then
           If COREGIS#(CII%) < 0.00005 Then
             COREGIS#(CII%) = CVD(Mid$(RELE$, 185, 8))
           End If
         End If
       End If
     Next U&
   End If
   '
   Call CIERRARCH("PVECLART")
   '
End Sub

Sub BLANCOSINFO()
   '
   CDMLIM% = 12 * 101 + 12
   CDMX% = CODIMES%
   '
   For KKI% = 1 To 32766
     COREGIS#(KKI%) = 0
   Next KKI%
   '
   RFF$ = RECFILT$("PVECLART", 1, MKI$(CDMX%))
   For U& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, U&, 4))
     RELE$ = REGLEIDO$("PVECLART", REG&)
     CII% = CVI(Mid$(RELE$, 7, 2))
     Call REPLA(RELE$, MKD$(0), 185, 8)
     Call GRAREG("PVECLART", RELE$, REG&)
   Next U&
   Call CIERRARCH("PVECLART")
   '
End Sub

