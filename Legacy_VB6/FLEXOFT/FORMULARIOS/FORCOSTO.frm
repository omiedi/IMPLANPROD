VERSION 5.00
Begin VB.Form FORCOSTO 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis de Costos Standard por Producto"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10350
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Consultas y Reportes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2950
      Left            =   7455
      TabIndex        =   64
      Top             =   1260
      Width           =   2745
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   610
         Left            =   210
         ScaleHeight     =   555
         ScaleWidth      =   2265
         TabIndex        =   74
         Top             =   2205
         Width           =   2325
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
            Height          =   555
            Left            =   0
            Picture         =   "FORCOSTO.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   0
            Width           =   645
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado Resumen Costos Standard"
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
            Left            =   780
            TabIndex        =   75
            Top             =   105
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   610
         Left            =   210
         ScaleHeight     =   555
         ScaleWidth      =   2265
         TabIndex        =   71
         Top             =   1575
         Width           =   2325
         Begin VB.CommandButton Command27 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   0
            Picture         =   "FORCOSTO.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   0
            Width           =   645
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Ficha Resumen Costos Standard"
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
            Left            =   780
            TabIndex        =   72
            Top             =   105
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   610
         Left            =   210
         ScaleHeight     =   555
         ScaleWidth      =   2265
         TabIndex        =   68
         Top             =   945
         Width           =   2325
         Begin VB.CommandButton Command14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   0
            Picture         =   "FORCOSTO.frx":0974
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   0
            Width           =   645
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado Costo    Std. Materiales"
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
            Left            =   780
            TabIndex        =   69
            Top             =   105
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   610
         Left            =   210
         ScaleHeight     =   555
         ScaleWidth      =   2265
         TabIndex        =   65
         Top             =   315
         Width           =   2325
         Begin VB.CommandButton Command16 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   0
            Picture         =   "FORCOSTO.frx":0C7E
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   0
            Width           =   645
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado Análisis Costos Standard"
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
            Left            =   780
            TabIndex        =   66
            Top             =   105
            Width           =   1695
         End
      End
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Save"
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
      Height          =   900
      Left            =   8400
      Picture         =   "FORCOSTO.frx":0F88
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Guarda las Modificaciones"
      Top             =   210
      Width           =   855
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   855
      Left            =   10185
      TabIndex        =   62
      Top             =   6195
      Width           =   330
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ultimo Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2050
      Left            =   7455
      TabIndex        =   54
      Top             =   4260
      Width           =   2745
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   210
         TabIndex        =   60
         Top             =   1155
         Width           =   1380
         Begin VB.Label Label35 
            BackColor       =   &H80000016&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   140
            TabIndex        =   61
            Top             =   315
            Width           =   1065
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
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
         Height          =   750
         Left            =   210
         TabIndex        =   58
         Top             =   315
         Width           =   1380
         Begin VB.Label Label34 
            BackColor       =   &H80000016&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   140
            TabIndex        =   59
            Top             =   315
            Width           =   1065
         End
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   1785
         ScaleHeight     =   75
         ScaleWidth      =   690
         TabIndex        =   56
         Top             =   420
         Width           =   750
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   57
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Calc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   1785
         Picture         =   "FORCOSTO.frx":1292
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Re-Cálculo General de Costos"
         Top             =   735
         Width           =   750
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Costo Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   3465
      TabIndex        =   36
      Top             =   5145
      Width           =   3795
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1155
         TabIndex        =   43
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "%:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   44
         Top             =   840
         Width           =   330
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2205
         TabIndex        =   42
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Costo Total:"
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
         Left            =   -210
         TabIndex        =   41
         Top             =   1365
         Width           =   2325
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gastos Generales:"
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
         Left            =   -105
         TabIndex        =   40
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2205
         TabIndex        =   39
         Top             =   735
         Width           =   1275
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Costos Directos:"
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
         Left            =   0
         TabIndex        =   38
         Top             =   420
         Width           =   2115
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2205
         TabIndex        =   37
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Complementos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   210
      TabIndex        =   23
      Top             =   5145
      Width           =   3060
      Begin VB.CommandButton Command11 
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
         Left            =   1575
         Picture         =   "FORCOSTO.frx":159C
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Actualizar Costo de Materiales por Articulo"
         Top             =   315
         Width           =   645
      End
      Begin VB.CommandButton Command13 
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
         Left            =   210
         Picture         =   "FORCOSTO.frx":18A6
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Maestro de Items de Stock"
         Top             =   315
         Width           =   645
      End
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
         Height          =   645
         Left            =   210
         Picture         =   "FORCOSTO.frx":19F0
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Padron Maestro de Proveedores "
         Top             =   945
         Width           =   645
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
         Height          =   645
         Left            =   840
         Picture         =   "FORCOSTO.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Formulas y Estructuras de Conjuntos"
         Top             =   315
         Width           =   645
      End
      Begin VB.CommandButton COMMAND3 
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
         Left            =   840
         Picture         =   "FORCOSTO.frx":2004
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Hojas de Ruta - Secuencia de Operaciones"
         Top             =   945
         Width           =   645
      End
      Begin VB.CommandButton Command9 
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
         Left            =   1575
         Picture         =   "FORCOSTO.frx":230E
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Consulta de Costo de Materiales por Proveedor"
         Top             =   945
         Width           =   645
      End
      Begin VB.CommandButton Command10 
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
         Left            =   2205
         Picture         =   "FORCOSTO.frx":2458
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Precios de Servicios por Proveedor"
         Top             =   315
         Width           =   645
      End
      Begin VB.CommandButton Command12 
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
         Left            =   2205
         Picture         =   "FORCOSTO.frx":25A2
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Precios de Servicios por Articulo"
         Top             =   945
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Fabricacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   210
      TabIndex        =   16
      Top             =   3360
      Width           =   7050
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Energía y Fuerza Motriz:"
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
         TabIndex        =   48
         Top             =   840
         Width           =   2115
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3990
         TabIndex        =   31
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3990
         TabIndex        =   30
         Top             =   735
         Width           =   1275
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3990
         TabIndex        =   29
         Top             =   1155
         Width           =   1275
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   22
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mano de Obra Directa:"
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
         TabIndex        =   21
         Top             =   420
         Width           =   2115
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   20
         Top             =   735
         Width           =   1275
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Herramental y Dispositivos:"
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
         Left            =   105
         TabIndex        =   19
         Top             =   1260
         Width           =   2325
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   18
         Top             =   1155
         Width           =   1275
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5460
         TabIndex        =   17
         ToolTipText     =   "Click Para Ver Análisis de Costos de Mano de Obra "
         Top             =   1155
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   210
      TabIndex        =   9
      Top             =   1260
      Width           =   7050
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Totales"
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
         Left            =   5460
         TabIndex        =   53
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Recibidos"
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
         Left            =   3990
         TabIndex        =   52
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.Propios"
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
         Left            =   2520
         TabIndex        =   51
         Top             =   210
         Width           =   1170
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         X1              =   0
         X2              =   7035
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Trabajos de Terceros:"
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
         TabIndex        =   47
         Top             =   1575
         Width           =   2115
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3990
         TabIndex        =   28
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3990
         TabIndex        =   27
         Top             =   1050
         Width           =   1275
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3990
         TabIndex        =   26
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Height          =   330
         Left            =   5460
         TabIndex        =   15
         ToolTipText     =   "Click para Ver Analisis de Costos de Materiales y Servicios"
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   14
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Piezas Compradas:"
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
         TabIndex        =   13
         Top             =   1155
         Width           =   2115
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   12
         Top             =   1050
         Width           =   1275
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Materias Primas:"
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
         TabIndex        =   11
         Top             =   735
         Width           =   2115
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   10
         Top             =   630
         Width           =   1275
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   210
      Width           =   2115
   End
   Begin VB.TextBox Text2 
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
      Left            =   3990
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   210
      Width           =   645
   End
   Begin VB.TextBox Text3 
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
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   735
      Width           =   7050
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3150
      TabIndex        =   1
      Top             =   210
      Width           =   175
   End
   Begin VB.CommandButton Command5 
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
      Height          =   900
      Left            =   7455
      Picture         =   "FORCOSTO.frx":26EC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   210
      Width           =   855
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
      Height          =   900
      Left            =   9345
      Picture         =   "FORCOSTO.frx":29F6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Terminar - Salir de Consulta de Costos"
      Top             =   210
      Width           =   855
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Informado:"
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
      Left            =   4620
      TabIndex        =   50
      Top             =   210
      Width           =   1170
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5880
      TabIndex        =   49
      Top             =   210
      Width           =   1380
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   315
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3045
      TabIndex        =   7
      Top             =   315
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Soluciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8925
      TabIndex        =   4
      Top             =   6680
      Width           =   3060
   End
   Begin VB.Image Image3 
      Height          =   675
      Left            =   7455
      Picture         =   "FORCOSTO.frx":2B40
      Stretch         =   -1  'True
      Top             =   6270
      Width           =   5775
   End
End
Attribute VB_Name = "FORCOSTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   MODO% = COMALTER%("Servicios por Proveedor - Elija su Opción de Trabajo\\Actualización y Consulta\Listado por Código\Cancelar")
   If MODO% = 1 Then
        Call CONECRUN("*LPREPRO", "Servicios por Proveedor")
      ElseIf MODO% = 2 Then
        Call CONECRUN("*SPRESER", "Servicios por Proveedor")
   End If
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call ACTCOSREPMAT
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   MODO% = COMALTER%("Servicios por Articulo - Elija su Opción de Trabajo\\Actualización y Consulta\Listado por Código\Cancelar")
   If MODO% = 1 Then
        Call CONECRUN("*LSERART", "Servicios por Articulo")
      ElseIf MODO% = 2 Then
        Dim PUNSER$(32767)
        URE& = ULTREG&("LPREPRO")
        For U& = 1 To URE&
           Call TRACE(20, U&, URE&)
           X$ = REGLEIDO$("LPREPRO", U&)
           CIIX% = CVI(Mid$(X$, 3, 2))
           If CIIX% > 0 Then
              If CIIX% <= 32767 Then
                 PUNSER$(CIIX%) = PUNSER$(CIIX%) + MKS$(U&)
              End If
           End If
        Next U&
        '
        FIN& = NUINV%
        J& = 0
        For U& = 1 To NUINV%
           Call TRACE(20, U&, FIN&)
           Y$ = REGLEIDO$("INVERT", U& + 1)
           CIIX% = CVI(Mid$(Y$, 17, 2))
           If CIIX% > 0 Then
              If CIIX% <= 32767 Then
                 PUNSE$ = PUNSER$(CIIX%)
                 If Len(PUNSE$) >= 4 Then
                    For W& = 1 To Len(PUNSE$) Step 4
                       REG& = CVS(Mid$(PUNSE$, W&, 4))
                       If REG& > 0 Then
                          If REG& <= URE& Then
                             X$ = REGLEIDO$("LPREPRO", REG&)
                             Z$ = Mid$(X$, 3, 38) + Left$(X$, 2) + Mid$(X$, 41)
                             J& = J& + 1
                             Call GRAREG("SERVART", Z$, J&)
                          End If
                       End If
                    Next W&
                 End If
              End If
           End If
        Next U&
        '
        Call SETULTREG("SERVART", J&)
        Call CIERRARCH("SERVART")
        '
        Call CONECRUN("*SERVART", "Listado de Servicios por Articulo")
        '
   End If
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   Call CONECRUN("AMASTO", "Maestro de Articulos de Stock")
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   EPAC$ = TRIM$(EMPREAC$)
   If InStr(EPAC$, "NOVOESPU") > 0 Then
         MABRUNOV.Show 1
      Else
         Call LIANACOS
   End If
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   Call CONECRUN("AMAPRO", "Padron Maestro de Proveedores")
   Command15.Enabled = True
End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   CIXX% = CODINT%(Text1.Text)
    If CIXX% > 0 Then
      Call CANACOSMOD(CIXX%)
      Call CANACOSMAT(CIXX%)
      Call PRIANCOSMOD(CIXX%)
   End If
   Command16.Enabled = True
End Sub


Private Sub Command2_Click()
    '
    Call SELECHO("MASTER")
    CODD$ = VALACT1$("MASTER")
    If CODD$ = "" Then
        Exit Sub
    End If
    COD$ = CODD$
    Text1.Text = COD$
    CI% = CODINT%(COD$)
    Text3.Text = DESCRIT$(CI%)
    Text2.Text = UNIMED$(CI%)
    '
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
    PrintForm
    Command27.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call CONECRUN("HORUTAS", "Hojas de Ruta y Proceso")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call CONECRUN("*LICOCAL", "Listado Resumen de Analisis de Costos")
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Call MENSERR(24, "No se Encuentra Ayuda Específica")
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    COD$ = TRIM$(Text1.Text)
    Call CIERRARCH("*.*")
    Call CONECRUN("FORMULAS/" + COD$, "Fórmulas y Estructuras de Producto")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call RECACOS
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   COD$ = TRIM$(Text1.Text)
   Call CIERRARCH("*.*")
   Call CONECRUN("*ANACOS/" + COD$, "Consulta de Analisis de Costos por Pieza")
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CONECRUN("*COCOREP", "Consulta Costos de Reposicion por Proveedor")
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    LU$ = LUDAT$
    Text4.Text = TRIM$(CONTROL$("", "PORGASGE"))
    '
    If InStr(UCase$(EPAC$), "NOVOESPU") > 0 Then
       Label37.Caption = "Listado Margen Bruto Ventas"
    End If
    '
    Label34.Caption = " " + CONTROL$("COSTOS", "FECALCOS")
    Label35.Caption = " " + CONTROL$("COSTOS", "HOCALCOS")
    '
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
    '
    Call CIERRARCH("*.*")
    Call FINAL("")
    '
End Sub

Private Sub Label10_Click()
   Label11.Enabled = False
   CIXX% = CODINT%(Text1.Text)
   If CIXX% > 0 Then
      'COD$ = TRIM$(Text1.Text)
      'Call CIERRARCH("*.*")
      'Call CONECRUN("*ANACOS/" + COD$, "Consulta de Analisis de Costos por Pieza")
      Call CANACOSMAT(CIXX%)
      FORANAMOD.Caption = "Analisis de Costos de Materiales y Servicios - " + TRIM$(CODEXT$(CIXX%)) + " - " + TRIM$(DESCRIT$(CIXX%))
      FORANAMOD.List1.Clear
      FORANAMOD.TEPRUEBA.Font = FORANAMOD.List1.Font
      FORANAMOD.TEPRUEBA.FontSize = FORANAMOD.List1.FontSize
      LOMAX& = 0
      FISAL% = FILEOPEN%("C:\FLEXOFT\ANCOSMAT.TXT", 1, 0)
      While Not EOF(FISAL%)
        Line Input #FISAL%, Z$
        FORANAMOD.List1.AddItem Z$
        FORANAMOD.TEPRUEBA.Caption = RTrim$(Z$)
        If FORANAMOD.TEPRUEBA.Width > LOMAX& Then
           LOMAX& = FORANAMOD.TEPRUEBA.Width
        End If
      Wend
      Close #FISAL%
      FORANAMOD.List1.Left = 0
      FORANAMOD.HScroll1.Min = 0
      FORANAMOD.HScroll1.Max = LOMAX + 250 - FORANAMOD.Width
      FORANAMOD.HScroll1.Value = 0
      On Error Resume Next
      FORANAMOD.HScroll1.SmallChange = FORANAMOD.HScroll1.Max / 80
      FORANAMOD.HScroll1.LargeChange = FORANAMOD.HScroll1.Max / 8
      On Error GoTo 0
      FORANAMOD.List1.Refresh
      FORANAMOD.Show 1
   End If
   Label11.Enabled = True

End Sub

Private Sub Label11_Click()
   '
   Label11.Enabled = False
   CIXX% = CODINT%(Text1.Text)
   If CIXX% > 0 Then
      Call CANACOSMOD(CIXX%)
      FORANAMOD.Caption = "Analisis de Costos de Mano de Obra - " + TRIM$(CODEXT$(CIXX%)) + " - " + TRIM$(DESCRIT$(CIXX%))
      FORANAMOD.List1.Clear
      FORANAMOD.TEPRUEBA.Font = FORANAMOD.List1.Font
      FORANAMOD.TEPRUEBA.FontSize = FORANAMOD.List1.FontSize
      LOMAX& = 0
      If EXISTE%("C:\FLEXOFT\ANCOSMOD.TXT") Then
         FISAL% = FILEOPEN%("C:\FLEXOFT\ANCOSMOD.TXT", 1, 0)
         While Not EOF(FISAL%)
           Line Input #FISAL%, Z$
           FORANAMOD.List1.AddItem Z$
           FORANAMOD.TEPRUEBA.Caption = RTrim$(Z$)
           If FORANAMOD.TEPRUEBA.Width > LOMAX& Then
              LOMAX& = FORANAMOD.TEPRUEBA.Width
           End If
         Wend
         Close #FISAL%
         FORANAMOD.List1.Left = 0
         FORANAMOD.HScroll1.Min = 0
         If LOMAX& > FORANAMOD.Width + 250 Then
            FORANAMOD.HScroll1.Max = LOMAX + 250 - FORANAMOD.Width
          Else
            FORANAMOD.HScroll1.Max = 1
         End If
         FORANAMOD.HScroll1.Value = 0
         On Error Resume Next
         FORANAMOD.HScroll1.SmallChange = FORANAMOD.HScroll1.Max / 80
         FORANAMOD.HScroll1.LargeChange = FORANAMOD.HScroll1.Max / 8
         On Error GoTo 0
         FORANAMOD.List1.Refresh
         FORANAMOD.Show 1
      End If
   End If
   Label11.Enabled = True
End Sub

Private Sub Text1_Change()
    '
    Text2.Text = ""
    Text3.Text = ""
    Label4.Caption = ""
    Label5.Caption = ""
    Label7.Caption = ""
    Label10.Caption = ""
    'Label14.Caption = ""
    If CODINT%(Text1.Text) > 0 Then
       COD$ = Text1.Text
       CI% = CODINT%(COD$)
       Text3.Text = DESCRIT$(CI%)
       Text2.Text = UNIMED$(CI%)
       Label4.Caption = TRIM$(CSTRING$(MKS$(COSUNI(CI%)), 3, 10, 4, 2))
       Call MUECOS(CI%)
    End If
    '
End Sub
'

Private Sub Text1_DBLCLICK()
    '
    Call SELECHO("MASTER")
    CODD$ = VALACT1$("MASTER")
    If CODD$ = "" Then
        Exit Sub
    End If
    COD$ = CODD$
    Text1.Text = COD$
    CI% = CODINT%(COD$)
    Text3.Text = DESCRIT$(CI%)
    Text2.Text = UNIMED$(CI%)
    '
End Sub
'

Private Sub Text1_LostFocus()
    COD$ = TRIM$(Text1.Text)
    CI% = CODINT%(COD$)
    Text3.Text = DESCRIT$(CI%)
    Text2.Text = UNIMED$(CI%)
End Sub

Private Sub Text2_GotFocus()
    FORCOSTO.Text1.SetFocus
End Sub

Private Sub Text3_Click()
    Text1.SetFocus
End Sub

Private Sub Text4_Change()
    COD$ = Text1.Text
    CI% = CODINT%(COD$)
    Call MUECOS(CI%)
End Sub

Private Sub Text4_LostFocus()
   Call GRACONTROL("", "PORGASGE", TRIM$(Text4.Text))
End Sub
