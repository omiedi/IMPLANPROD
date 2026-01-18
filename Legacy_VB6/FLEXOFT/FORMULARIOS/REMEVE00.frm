VERSION 5.00
Begin VB.Form REMEVE00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gestión de Ventas - Resúmenes Mensuales"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
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
      Height          =   855
      Left            =   10080
      Picture         =   "REMEVE00.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Re-Calcula Resúmenes Mensuales"
      Top             =   3675
      Width           =   855
   End
   Begin VB.CommandButton Command5 
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
      Height          =   855
      Left            =   10080
      Picture         =   "REMEVE00.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Imprime Ficha Resumen Mensual"
      Top             =   6195
      Width           =   855
   End
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
      Height          =   855
      Left            =   10080
      Picture         =   "REMEVE00.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1260
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0FFC0&
      Caption         =   "Período Mensual Activo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7680
      Left            =   210
      TabIndex        =   8
      Top             =   105
      Width           =   9675
      Begin VB.Frame Frame8 
         BackColor       =   &H00F0FFC0&
         Caption         =   "Avance de Proceso de Cálculo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   5145
         TabIndex        =   49
         Top             =   1470
         Width           =   4215
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   210
            ScaleHeight     =   75
            ScaleWidth      =   3735
            TabIndex        =   66
            Top             =   2625
            Width           =   3795
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   67
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00F0FFC0&
            Caption         =   "Generación de Listadores."
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
            Left            =   210
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   2310
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00F0FFC0&
            Caption         =   "Re-Imputación Automática de Facturas"
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
            Left            =   210
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   2100
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00F0FFC0&
            Caption         =   "Revisión Movimientos de Stock."
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
            Left            =   210
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   1680
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00F0FFC0&
            Caption         =   "Lectura de Imputación Manual."
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
            Left            =   210
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   1890
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00F0FFC0&
            Caption         =   "Archivo de Conciliación de Datos"
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
            Left            =   210
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   1470
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00F0FFC0&
            Caption         =   "Revisión  Base de Datos de Facturación"
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
            Left            =   210
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1260
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00F0FFC0&
            Caption         =   "Arrastre de Saldos por Cobranza"
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
            Left            =   210
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   1050
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00F0FFC0&
            Caption         =   "Cuentas Contables por Cliente."
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
            Left            =   210
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   840
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00F0FFC0&
            Caption         =   "Grupos Contables por Artículos."
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
            Left            =   210
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   630
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00F0FFC0&
            Caption         =   "Validación Cuentas Contables de I.V.A."
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
            Left            =   210
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   420
            Width           =   3795
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00F0FFC0&
         Caption         =   "Reportes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   7140
         TabIndex        =   44
         Top             =   6090
         Width           =   2220
         Begin VB.CommandButton Command9 
            Caption         =   "I.V.A"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   210
            Picture         =   "REMEVE00.frx":0DB6
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Emite Subdiario de I.V.A. por Impresora"
            Top             =   420
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Informes"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1155
            Picture         =   "REMEVE00.frx":10C0
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Resúmenes Estadísticos de Ventas"
            Top             =   420
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00F0FFC0&
         Caption         =   "Contabilizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   7140
         TabIndex        =   43
         Top             =   4620
         Width           =   2220
         Begin VB.CommandButton Command7 
            Caption         =   "Asiento"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1155
            Picture         =   "REMEVE00.frx":13CA
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Presenta Asiento Contable de Ventas"
            Top             =   420
            Width           =   855
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Subdiario"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   210
            Picture         =   "REMEVE00.frx":16D4
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Emite Subdiario de Ventas por Impresora"
            Top             =   420
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00F0FFC0&
         Caption         =   "Posición Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2880
         Left            =   315
         TabIndex        =   29
         Top             =   4620
         Width           =   6735
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   25
            Left            =   2310
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   2415
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   24
            Left            =   2310
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1365
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   23
            Left            =   2310
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   1050
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   22
            Left            =   2310
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   630
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   21
            Left            =   5250
            TabIndex        =   64
            Text            =   " "
            Top             =   2415
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   3780
            TabIndex        =   63
            TabStop         =   0   'False
            Text            =   " "
            Top             =   630
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   840
            TabIndex        =   62
            TabStop         =   0   'False
            Text            =   " "
            Top             =   630
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   20
            Left            =   3780
            TabIndex        =   46
            Text            =   " "
            Top             =   2415
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   19
            Left            =   840
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   2415
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   18
            Left            =   5250
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1995
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   17
            Left            =   840
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1995
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   16
            Left            =   840
            TabIndex        =   33
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1680
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   840
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1365
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   840
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   1050
            Width           =   1380
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   5250
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   " "
            Top             =   630
            Width           =   1380
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tasa 19-21 %"
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
            Index           =   20
            Left            =   525
            TabIndex        =   69
            Top             =   345
            Width           =   1590
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Otros:"
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
            Index           =   19
            Left            =   -2415
            TabIndex        =   65
            Top             =   1995
            Width           =   3165
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Index           =   18
            Left            =   2730
            TabIndex        =   48
            Top             =   1785
            Width           =   1590
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total General"
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
            Index           =   17
            Left            =   5355
            TabIndex        =   47
            Top             =   345
            Width           =   1170
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Retenc:"
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
            Index           =   16
            Left            =   -2415
            TabIndex        =   42
            Top             =   1680
            Width           =   3165
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "R.N.I.:"
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
            Index           =   15
            Left            =   -840
            TabIndex        =   41
            Top             =   1365
            Width           =   1590
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IVA:"
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
            Index           =   14
            Left            =   -840
            TabIndex        =   40
            Top             =   1050
            Width           =   1590
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "T.9,5-10,5 %"
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
            Index           =   11
            Left            =   1995
            TabIndex        =   39
            Top             =   345
            Width           =   1590
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Neto:"
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
            Left            =   -1995
            TabIndex        =   38
            Top             =   630
            Width           =   2745
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bruto:"
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
            Left            =   -2415
            TabIndex        =   37
            Top             =   2415
            Width           =   3165
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Exentos"
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
            Left            =   3885
            TabIndex        =   36
            Top             =   345
            Width           =   1170
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00F0FFC0&
         Caption         =   "Resumen Mensual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   315
         TabIndex        =   13
         Top             =   1470
         Width           =   4530
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   2835
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2100
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   2835
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1470
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   2835
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   420
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1155
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   840
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1155
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1155
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1155
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1470
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   2835
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1785
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   2835
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2520
            Width           =   1485
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Final del Período"
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
            Left            =   -105
            TabIndex        =   28
            Top             =   2625
            Width           =   2745
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ajustes ........................."
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
            Left            =   -525
            TabIndex        =   26
            Top             =   2205
            Width           =   3165
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Inicial del Período"
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
            Left            =   -105
            TabIndex        =   24
            Top             =   525
            Width           =   2745
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Facturas"
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
            Left            =   -525
            TabIndex        =   23
            Top             =   840
            Width           =   1590
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Débitos"
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
            Left            =   -525
            TabIndex        =   22
            Top             =   1155
            Width           =   1590
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Créditos"
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
            Left            =   -525
            TabIndex        =   21
            Top             =   1470
            Width           =   1590
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cobranza ........................."
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
            Left            =   -525
            TabIndex        =   20
            Top             =   1890
            Width           =   3165
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F0FFC0&
         Caption         =   "Generación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   5145
         TabIndex        =   10
         Top             =   315
         Width           =   4215
         Begin VB.TextBox Text4 
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
            Index           =   1
            Left            =   3150
            TabIndex        =   61
            TabStop         =   0   'False
            Text            =   " "
            Top             =   420
            Width           =   750
         End
         Begin VB.TextBox Text4 
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
            Index           =   0
            Left            =   1050
            TabIndex        =   60
            TabStop         =   0   'False
            Text            =   " "
            Top             =   420
            Width           =   1065
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
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   13
            Left            =   -1260
            TabIndex        =   12
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
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   12
            Left            =   1470
            TabIndex        =   11
            Top             =   525
            Width           =   1590
         End
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
         Left            =   3780
         TabIndex        =   0
         Top             =   735
         Width           =   175
      End
      Begin VB.Label MESAN 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1575
         TabIndex        =   75
         Top             =   740
         Width           =   2220
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4305
         Picture         =   "REMEVE00.frx":19DE
         Top             =   650
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4305
         Picture         =   "REMEVE00.frx":1CE8
         Top             =   700
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mes y año:"
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
         Index           =   10
         Left            =   -735
         TabIndex        =   9
         Top             =   840
         Width           =   2220
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      Picture         =   "REMEVE00.frx":1FF2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Re-Calcula Resúmenes Mensuales"
      Top             =   2205
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
      Height          =   855
      Left            =   10080
      Picture         =   "REMEVE00.frx":22FC
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Finaliza la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   510
      Left            =   9530
      Picture         =   "REMEVE00.frx":2446
      Top             =   7260
      Width           =   2010
   End
End
Attribute VB_Name = "REMEVE00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CUEIVA1%, CUEADU1%, CUERET1%, CUEDEUV1%, YACUE%
Dim TIS$(10)
Dim GRCOCLI%(32768), CUEMA%(32768), EXICLI%(32768), RACLI$(32768)
Dim SUMAFAC#(32)

'Public CARGAMEN%
'
Dim CORRESP$, ABREMES$, DES%, HAS%, DESFE$, HASFE$
Dim CUEV%(8), IMXX#(8)
Dim CUEIVA$, CUEADU$, CUERET$
Dim GRUCOD%(1024), GRUCLI%(1024), CUECON%(1024), CARUCO%
Dim GRUCOVE%(256), CAGRUCOVE%
Dim CUECOD%(32768)
Dim GRUCOC%(256), CAGRUCOCLI%
Dim PRIMOSTO&(8192), ULTMOSTO&(8192)
Dim IFAC#(8192), IDET#(8192), SIFAC%(8192)
Dim NRECI&(8192), REGFA$(8192)
Dim X0 As String * 96
'
Dim PORCO(256), TOCOMI#(256), TOCOB#(256)
Dim SUPED#(256), COMPED#(256), SUCOB#(256), COMCOB#(256)
'
Dim YALIQUI%
Public MODOCLAVE%     ' MODALIDAD DE CLASIFICACION DE VENTAS
Dim MODOCALVE%     ' MODALIDAD CALCULO DE VENTAS
'


Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'

Private Sub Command10_Click()
   Command10.Enabled = False
   Call CONECRUN("FLSETUP/OPCLAVEN", "Configuración de Funcionamiento")
   Command10.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call CALREVEN
   Command2.Enabled = True
End Sub
'

Private Sub Command3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
    End If
End Sub
'

Private Sub Command5_Click()
   PRINTFORM
End Sub

Sub Command6_Click()
    '
    Call ARMAMES
    If CORRESP$ = "" Then Exit Sub
    '
    If FECHANUM(Text4(0).TEXT) < 1 Then
        OPX% = COMALTER%("La Generación de Resumen Mensual\no se ha Completado.\  \Puede Ahora Re-Calcular el Resumen\o Abandonar el Proceso.\\Re-Calcular\Abandonar")
        If OPX% = 1 Then
          Call CALREVEN
        End If
        Exit Sub
    End If
    '
    Command6.Enabled = False
    Call BACKUPMEN(3, HAYERRO%)
    If HAYERRO% > 0 Then
        Command6.Enabled = True
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    X$ = REGLEIDO$("DEUDOR1", 1)
    Call GRAREG("DEUDOR1", X$, 1)
    Call CIERRARCH("DEUDOR1")
    Screen.MousePointer = 1
    '
    Call FINAL("*SUDISOR")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    Call ARMAMES
    If CORRESP$ = "" Then Exit Sub
    '
    If FECHANUM(Text4(0).TEXT) < 1 Then
        OPX% = COMALTER%("La Generación de Resumen Mensual\no se ha Completado.\  \Puede Ahora Re-Calcular el Resumen\o Consultar en Forma Parcial.\\Re-Calcular\Consultar")
        If OPX% = 1 Then
            Call CALREVEN
            Exit Sub
        End If
    End If
    Call CONASIDIA
    '
End Sub
'

Private Sub Command8_Click()
   '
   Call ARMAMES
   If CORRESP$ = "" Then Exit Sub
   '
   If FECHANUM(Text4(0).TEXT) < 1 Then
        OPX% = COMALTER%("La Generación de Resumen Mensual\no se ha Completado.\  \Puede Ahora Re-Calcular el Resumen\o Abandonar el Proceso.\\Re-Calcular\Abandonar")
        If OPX% = 1 Then
          Call CALREVEN
        End If
        Exit Sub
    End If
    '
    Command8.Enabled = False
    Call BACKUPMEN(3, HAYERRO%)
    If HAYERRO% > 0 Then
        Command8.Enabled = True
        Exit Sub
    End If
    '
    RESUVE00.MESAN = MESAN.Caption
    RESUVE00.Show 1
    Command8.Enabled = True
End Sub

Sub Command9_Click()
    '
    Call ARMAMES
    If CORRESP$ = "" Then Exit Sub
    '
    If FECHANUM(Text4(0).TEXT) < 1 Then
        OPX% = COMALTER%("La Generación de Resumen Mensual\no se ha Completado.\  \Puede Ahora Re-Calcular el Resumen\o Abandonar el Proceso.\\Re-Calcular\Abandonar")
        If OPX% = 1 Then
          Call CALREVEN
        End If
        Exit Sub
    End If
    '
    Command9.Enabled = False
    Dim GRAVA#(4), IIVA#(4)
    '
    Call BACKUPMEN(3, HAYERRO%)
    If HAYERRO% > 0 Then
        Command9.Enabled = True
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    X$ = REGLEIDO$("DEUDOR1", 1)
    Call GRAREG("DEUDOR1", X$, 1)
    Call CIERRARCH("DEUDOR1")
    Screen.MousePointer = 1
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("SUIVAVE")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("SUIVAVE", U&)
      XX1$ = Left$(XX$, 72)
      XX3$ = Mid$(XX$, 121)
      For KKI% = 1 To 4
        GRAVA#(KKI%) = 0: IIVA#(KKI%) = 0
      Next KKI%
      For KKJ% = 1 To 3
        IGRAA# = CVD(Mid$(XX$, 57 + 16 * KKJ%, 8))
        IIVAA# = CVD(Mid$(XX$, 65 + 16 * KKJ%))
        KKI% = 1
        If Abs(IGRAA#) >= 0.005 Then
          TASARE = 100 * IIVAA# / IGRAA#
          If TASARE > 20.5 And TASARE < 21.5 Then KKI% = 2
          If TASARE > 9 And TASARE <= 10 Then KKI% = 3
          If TASARE > 10 And TASARE <= 11 Then KKI% = 4
        End If
        GRAVA#(KKI%) = GRAVA#(KKI%) + IGRAA#
        IIVA#(KKI%) = IIVA#(KKI%) + IIVAA#
      Next KKJ%
      XX2$ = ""
      For KKI% = 1 To 4
        XX2$ = XX2$ + MKD$(GRAVA#(KKI%)) + MKD$(IIVA#(KKI%))
      Next KKI%
      XX$ = XX1$ + XX2$ + XX3$
      Call GRAREG("SUIVA19", XX$, U&)
    Next U&
    Call SETULTREG("SUIVA19", FIN&)
    Call CIERRARCH("SUIVA19")
    '
    Call HEADERS("SUIVA19", "")
    Call HEADERS("SUIVA19", "Corresponde a: " + MESAN)
    '
    Call FILESORT("SUIVA19", "SUIVA19", 1, 2, "ASC")
    Call FINAL("*SUIVA19")
    '
    Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Dim TAIVA(2)
    TAIVA(0) = 21: TAIVA(1) = 10.5
    For U& = 1 To 2
      If U& <= ULTREG&("GRUPIVA") Then
        X$ = REGLEIDO$("GRUPIVA", U&)
        TAIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
      End If
    Next U&
    Call CIERRARCH("GRUPIVA")
    '
    'REMEVENT.LABEL2(20).Caption = "Tasa " + TRIM$(Str$(TAIVA(0))) + " %"
    'REMEVENT.LABEL2(11).Caption = "Tasa " + TRIM$(Str$(TAIVA(1))) + " %"
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub MESAN_Change()
   TXT$ = MESAN
   If MESANO$(TXT$, DES%, HAS%) = TXT$ Then
      If Val(Control$("", "CIEMEVEN")) < HAS% Then
          Image1.Visible = False
          Image2.Visible = True
         Else
          Image2.Visible = False
          Image1.Visible = True
      End If
   End If
   Call CARGARESMEN
   Call FRESHALL
   DoEvents
End Sub

Private Sub MESAN_DblClick()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
      MESAN.Refresh
    End If
End Sub

Private Sub Image2_Click()
   If USNBR% Mod 100 = 1 Then
      N5% = FILEOPEN%("USER.DAT", 9, 64)
      CAUS% = LOF(N5%) / 64
      Close #N5%
      If CAUS% > 1 Then
         Call MENSERR(24, "Usuario Supervisor no Autorizado\Para la Operación Solicitada.\   \Delegue el Derecho de Acceso\en Otro Usuario del Sistema.")
         Exit Sub
      End If
   End If
   '
   TXT$ = MESAN
   If MESANO$(TXT$, DES%, HAS%) = TXT$ Then
      If DERACCE%("CIEMEVEN", "Cierre y Re-Apertura Mensual Ventas") >= 2 Then
         Call GRACONTROL("", "CIEMEVEN", TRIM$(Str$(HAS%)))
         Image2.Visible = False
         Image1.Visible = True
      End If
   End If
End Sub

Private Sub Image1_Click()
   If USNBR% Mod 100 = 1 Then
      N5% = FILEOPEN%("USER.DAT", 9, 64)
      CAUS% = LOF(N5%) / 64
      Close #N5%
      If CAUS% > 1 Then
         Call MENSERR(24, "Usuario Supervisor no Autorizado\Para la Operación Solicitada.\   \Delegue el Derecho de Acceso\en Otro Usuario del Sistema.")
         Exit Sub
      End If
   End If
   '
   TXT$ = MESAN
   If MESANO$(TXT$, DES%, HAS%) = TXT$ Then
      If DERACCE%("CIEMEVEN", "Cierre y Re-Apertura Mensual Ventas") >= 2 Then
         If COMALTER%("Confirme Re-Apertura del Período Mensual\\Cancelar\Re-Abrir") = 2 Then
            Call GRACONTROL("", "CIEMEVEN", TRIM$(Str$(DES% - 1)))
            Image1.Visible = False
            Image2.Visible = True
         End If
      End If
   End If
End Sub

Sub CALREVEN()
     '
Dim SALINI#(32767), SUMFAC#(32767), SUMCOB#(32767)
'
Dim NCV%(8192), TII%(8192), NROV&(8192), FFII%(8192) ' dimensionamiento
'
     Call ARMAMES
     If CORRESP$ = "" Then Exit Sub
     '
     Dim TAIVA(2)
     TAIVA(0) = 21: TAIVA(1) = 10.5
     For U& = 1 To 2
       If U& <= ULTREG&("GRUPIVA") Then
         X$ = REGLEIDO$("GRUPIVA", U&)
         TAIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
       End If
     Next U&
     Call CIERRARCH("GRUPIVA")
     '
     '
     For XI% = 0 To 9: Check1(XI%).Value = 0: Next XI%
     '
10   Check1(0).Value = 2
     Call VALICUEIVA
     Screen.MousePointer = 11
     If YACUE% <> 1 Then
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        Screen.MousePointer = 1
        Call MENSERR(24, "Faltan Cuentas Contables\de Clasificación de Ventas.")
        Call CONECRUN("FLSETUP/OPCLAVEN", "Configuración de Funcionamiento")
        Exit Sub
     End If
     Check1(0).Value = 1
     '
20   Call BACKUPMEN(1, HAYERRO%)
     If HAYERRO% > 0 Then
        Screen.MousePointer = 1
        GoTo 999
     End If
     '
30   Check1(1).Value = 2
     If YACUE% <> 1 Then
        Screen.MousePointer = 1
        GoTo 999
     End If
     Check1(1).Value = 1
     '
40   Check1(2).Value = 2
     Call CARCUECLI
     If YACUE% <> 1 Then
        Screen.MousePointer = 1
        GoTo 999
     End If
     Check1(2).Value = 1
     '
     For XI% = 0 To 32: SUMAFAC#(XI%) = 0: Next XI%
     REGDETACC& = 0
     RECUEC00.LINDICE.Clear
     '
     For IL& = 0 To 32767
        SALINI#(IL&) = 0
        SUMFAC#(IL&) = 0
        SUMCOB#(IL&) = 0
     Next IL&
     '
     For XII% = 0 To 7
        Text1(XII%).TEXT = ""
        Text1(XII%).Refresh
     Next XII%
     For XII% = 11 To 25
        Text1(XII%).TEXT = ""
        Text1(XII%).Refresh
     Next XII%
     '
     Call SETULTREG("CONCIVEN", 0)
     GoSub 10000   ' arrastre de saldos segun cobranza
     On Error Resume Next
     For XII% = 0 To 25
        Text1(XII%).TEXT = ""
        Text1(XII%).Refresh
     Next XII%
     On Error GoTo 0
     '
     GoSub 11000   ' seleccionar facturas
     GoSub 11100   ' cargar facturas a vector busqueda
     GoSub 11200   ' verificar con facturacion
     GoSub 13000   ' imputacion automatica
     GoSub 14000   ' GENERACION DE LISTADORES
     '
998  Call CIERRARCH("*.*")
     Call BACKUPMEN(2, HAYERRO%)
     If HAYERRO% = 0 Then
        GoSub 15000   ' GRABAR RESUMEN MENSUAL
     End If
     '
     Close #N2%
     Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
999  Call CIERRARCH("*.*")
     Call LIBARCH("*.*")
     Exit Sub



        '
10000 ' Revision del Archivo de Cobranza ... **********************
        '
        Check1(3).Value = 2
        FIN& = ULTREG&("COBRAN")
        CAFAC% = 0
        REGCOBCLI& = 0
        '
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           X$ = REGLEIDO$("COBRAN", IL&)
           NC% = CVI(Left$(X$, 2))
           If NC% > 0 And NC% <= 32767 Then
              FF% = CVI(Mid$(X$, 7, 2))
              TI% = 4: SIG% = (-1)
              NRO& = CVS(Mid$(X$, 3, 4))
              TOT# = (Int(100 * (CVD(Mid$(X$, 15, 8))) + 0.5)) / 100
              '
              If TOT# >= 0.005 Then
                 If EXICLI%(NC%) < 1 Then
                    Call MENSERR(24, "Cobranza Registrada sobre Cuenta '" + TRIM$(Str$(NC%)) + "' Inexistente.")
                 End If
              End If
              '
              If FF% < DES% Then
                  SUMAFAC#(0) = SUMAFAC#(0) + TOT# * SIG%
                  SUMAFAC#(7) = SUMAFAC#(7) + TOT# * SIG%
                  SALINI#(NC%) = SALINI#(NC%) + TOT# * SIG%
                ElseIf FF% <= HAS% Then
                  SUMAFAC#(5) = SUMAFAC#(5) + TOT#
                  SUMAFAC#(7) = SUMAFAC#(7) + TOT# * SIG%
                  SUMCOB#(NC%) = SUMCOB#(NC%) + TOT#
                  '
                  Y$ = REGBLAN$("DETACCD")
                  Call REPLA(Y$, Chr$(TI%), 1, 1)
                  Call REPLA(Y$, MKI$(FF%), 3, 2)
                  Call REPLA(Y$, MKS$(NRO&), 5, 4)
                  Call REPLA(Y$, MKI$(NC%), 9, 2)
                  Call REPLA(Y$, MKD$(TOT#), 11, 8)
                  Call REPLA(Y$, MKD$(TOT#), 51, 8)
                  '
                  REGDETACC& = REGDETACC& + 1
                  Call GRAREG("DETACCD", Y$, REGDETACC&)
                  FF% = CVI(Mid$(Y$, 3, 2))
                  Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REGDETACC&), 4, 8, 0, 2)
                  RECUEC00.LINDICE.AddItem Z$
                  '
                  Y$ = REGBLAN$("RECOBCLI")
                  Call REPLA(Y$, MKI$(NC%), 1, 2)
                  Call REPLA(Y$, MKI$(FF%), 3, 2)
                  Call REPLA(Y$, MKS$(NRO&), 5, 4)
                  Call REPLA(Y$, MKD$(TOT#), 25, 8)
                  REGCOBCLI& = REGCOBCLI& + 1
                  Call GRAREG("RECOBCLI", Y$, REGCOBCLI&)
                  '
              End If
           End If
10049   Next IL&
        '
        Call SETULTREG("RECOBCLI", REGCOBCLI&)
        Call CIERRARCH("RECOBCLI")
        Check1(3).Value = 1
        Return
        '

11000 ' Revision del Archivo de Facturacion ... **********************
        '
        Check1(4).Value = 2
        FIN& = ULTREG&("FACTUR")
        CAFAC% = 0
        CARESUVE& = 0
        '
        FFANT% = 0
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           X$ = REGLEIDO$("FACTUR", IL&)
           '
           FF% = CVI(Mid$(X$, 7, 2))
           If FF% < 1 Or FF% > HAS% Then GoTo 11049
           '
           NC% = CVI(Left$(X$, 2))
           If NC% < 1 Or NC% > 32767 Then GoTo 11049
           '
           TI% = Asc(Mid$(X$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
           VA% = Asc(Mid$(X$, 14, 1)): If VA% < 1 Or VA% > 3 Then VA% = 1 ' default factura
           '
11002      NETO1# = (Int(100 * (CVD(Mid$(X$, 15, 8))) + 0.5)) / 100
           IVAD1# = (Int(100 * CVD(Mid$(X$, 23, 8)) + 0.5)) / 100
           RNI# = (Int(100 * CVD(Mid$(X$, 31, 8)) + 0.5)) / 100
           NETO2# = (Int(100 * CVD(Mid$(X$, 49, 8)) + 0.5)) / 100
           IVAD2# = (Int(100 * CVD(Mid$(X$, 57, 8)) + 0.5)) / 100
           '
           If Abs(NETO2#) > 0.1 Then
             If Abs(100 * IVAD2# / NETO2# - TAIVA(0)) < 0.2 Then
               NETO1# = NETO1# + NETO2#
               Call REPLA(X$, MKD$(NETO1#), 15, 8)
               NETO2# = 0
               Call REPLA(X$, MKD$(NETO2#), 49, 8)
               IVAD1# = IVAD1# + IVAD2#
               Call REPLA(X$, MKD$(IVAD1#), 23, 8)
               IVAD2# = 0
               Call REPLA(X$, MKD$(IVAD2#), 57, 8)
               Call GRAREG("FACTUR", X$, IL&)
             End If
           End If
           '
           NETO# = NETO1# + NETO2#
           IVAD# = IVAD1# + IVAD2#
           TOT# = NETO# + IVAD# + RNI# + IPER# + IRET#
           '
           SIG% = 1
           If TI% > 3 Then
             NETO1# = Abs(NETO1#)
             NETO2# = Abs(NETO2#)
             NETO# = Abs(NETO#)
             IVAD1# = Abs(IVAD1#)
             IVAD2# = Abs(IVAD2#)
             IVAD# = Abs(IVAD#)
             RNI# = Abs(RNI#)
             TOT# = Abs(TOT#)
             SIG% = (-1)
           End If
           '
           If FF% < DES% Then
             SUMAFAC#(0) = SUMAFAC#(0) + TOT# * SIG%
             SUMAFAC#(7) = SUMAFAC#(7) + TOT# * SIG%
             If NC% > 0 Then SALINI#(NC%) = SALINI#(NC%) + TOT# * SIG%
             GoTo 11049
           End If
           '
           If FF% >= DES% Then
             If FF% <= HAS% Then
               NCVI% = NC%
               If NCVI% <= 0 Then NCVI% = 0
               '
               FEMI% = FF%
               If NCVI% = 0 Then
                 RRCC$ = Space$(46) + Chr$(0)
                 Call REPLA(RRCC$, "Clientes de Mostrador", 1, 30)
                 Call REPLA(RRCC$, "NO POSEE", 31, 16)
                 Call REPLA(RRCC$, Chr$(0), 47, 1)
                 RACLI$(NCVI%) = RRCC$
               End If
               '
               If RACLI$(NCVI%) = "" Then
                 RRCC$ = Space$(46) + Chr$(0)
                 Call REPLA(RRCC$, RASOCLI$(NCVI%), 1, 30)
                 Call REPLA(RRCC$, CUITCLI$(NCVI%), 31, 16)
                 Call REPLA(RRCC$, Chr$(PIVACLI%(NCVI%)), 47, 1)
                 RACLI$(NCVI%) = RRCC$
               End If
               '
               If CUEMA%(NCVI%) < 1 Then
                 CUEMA%(NCVI%) = CUECOINT%(CUEMADRE$(NCVI%))
               End If
               '
11045          TIXX$ = TIS$(TI%)
               NRO& = CVS(Mid$(X$, 3, 4))
               NUSUCUVE% = CVI(Mid$(X$, 47, 2))
               If TI% = 0 Then
                 If NRO& = 0 Then
                   NRO& = IL&
                   Call REPLA(X$, MKS$(NRO&), 3, 4)
                   Call GRAREG("FACTUR", X$, IL&)
                 End If
               End If
               SUTT# = SUTT# + TOT# * SIG%
               COEF# = 1
               '
               Y$ = String$(256, 0)
               Call REPLA(Y$, TIXX$, 1, 2)
               Call REPLA(Y$, MKS$(NRO&), 3, 4)
               Call REPLA(Y$, MKI$(NC%), 7, 2)
               Call REPLA(Y$, MKI$(FEMI%), 9, 2)
               Call REPLA(Y$, MKI$(NUSUCUVE%), 11, 2)
               Call REPLA(Y$, String$(80, 0), 17, 80)
               Call REPLA(Y$, Chr$(TI%) + Chr$(VA%), 97, 2)
               Call REPLA(Y$, MKD$(NETO1# * SIG%), 99, 8)
               Call REPLA(Y$, MKD$(IVAD1# * SIG%), 107, 8)
               Call REPLA(Y$, MKD$(RNI# * SIG%), 115, 8)
               Call REPLA(Y$, MKD$(NETO2# * SIG%), 123, 8)
               Call REPLA(Y$, MKD$(IVAD2# * SIG%), 131, 8)
               Call REPLA(Y$, MKD$(TOT# * SIG%), 139, 8)
               Call REPLA(Y$, MKD$(COEF#), 147, 8)
               Call REPLA(Y$, MKI$(0), 155, 2)
               Call REGAPP("CONCIVEN", Y$)
               '
               If TI% > 0 Then
                 If TI% <= 6 Then
                   If NC% >= 0 Then
                        CLI$ = RASOCLI$(NC%)
                     ElseIf NC% < 0 Then
                        RECUI& = Abs(NC%)
                        If RECUI& <= ULTREG&("CUITS") Then
                          RCUX$ = REGLEIDO$("CUITS", RECUI&)
                          CLI$ = Left$(RCUX$, 30)
                        End If
                   End If
                   '
                   YR$ = REGBLAN$("RESUVEN")
                   Call REPLA$(YR$, MKI$(NC%), 1, 2)
                   Call REPLA$(YR$, CLI$, 3, 30)
                   Call REPLA(YR$, TIXX$ + AJUSTD$(Str$(NRO&), 6), 33, 8)
                   Call REPLA(YR$, MKI$(FEMI%), 41, 2)
                   Call REPLA(YR$, MKD$(NETO# * SIG%), 43, 8)
                   CARESUVE& = CARESUVE& + 1
                   Call GRAREG("RESUVEN", YR$, CARESUVE&)
                 End If
               End If
               '
               CAFAC% = CAFAC% + 1
               '
               Select Case TI%
                  Case 0
                    SUMAFAC#(6) = SUMAFAC#(6) + TOT# * SIG%
                  Case 1
                    SUMAFAC#(1) = SUMAFAC#(1) + TOT# * SIG%
                  Case 2
                    SUMAFAC#(1) = SUMAFAC#(1) + TOT# * SIG%
                  Case 3
                    SUMAFAC#(2) = SUMAFAC#(2) + TOT# * SIG%
                  Case 4
                    SUMAFAC#(5) = SUMAFAC#(5) + Abs(TOT#)
                  Case 5
                    SUMAFAC#(3) = SUMAFAC#(3) + Abs(TOT#)
                  Case 6
                    SUMAFAC#(3) = SUMAFAC#(3) + Abs(TOT#)
                  Case 7
                    SUMAFAC#(6) = SUMAFAC#(6) + TOT# * SIG%
               End Select
               SUMAFAC#(7) = SUMAFAC#(7) + TOT# * SIG%
               SUMAFAC#(4) = SUMAFAC#(1) + SUMAFAC#(2) - SUMAFAC#(3)
               '
               NCII% = NC%: If NCII% < 0 Then NCII% = 0
               SUMFAC#(NCII%) = SUMFAC#(NCII%) + TOT# * SIG%
               If NCII% <= 0 Then
                 SUMCOB#(NCII%) = SUMCOB#(NCII%) + TOT# * SIG%
               End If
               '
               Y$ = REGBLAN$("DETACCD")
               Call REPLA(Y$, Chr$(TI%), 1, 1)
               Call REPLA(Y$, MKI$(FEMI%), 3, 2)
               Call REPLA(Y$, MKS$(NRO&), 5, 4)
               Call REPLA(Y$, MKI$(NC%), 9, 2)
               Call REPLA(Y$, MKD$(NETO#), 11, 8)
               Call REPLA(Y$, MKD$(IVAD#), 19, 8)
               Call REPLA(Y$, MKD$(RNI#), 27, 8)
               Call REPLA(Y$, MKD$(IPER#), 35, 8)
               Call REPLA(Y$, MKD$(IRET#), 43, 8)
               Call REPLA(Y$, MKD$(TOT#), 51, 8)
               '
               REGDETACC& = REGDETACC& + 1
               Call GRAREG("DETACCD", Y$, REGDETACC&)
               FEMI% = CVI(Mid$(Y$, 3, 2))
               Z$ = CSTRING$(MKI$(FEMI%), 1, 6, 0, 2) + CSTRING$(MKS$(REGDETACC&), 4, 8, 0, 2)
               RECUEC00.LINDICE.AddItem Z$
               '
               For XII% = 0 To 7
                 Text1(XII%).TEXT = CSTRING$(MKD$(SUMAFAC#(XII%)), 4, 11, 2, 2)
                 Text1(XII%).Refresh
               Next XII%
               '
             End If
             '
           End If
           '
11049   Next IL&
        '
        For XII% = 0 To 7
           Text1(XII%).TEXT = CSTRING$(MKD$(SUMAFAC#(XII%)), 4, 11, 2, 2)
           Text1(XII%).Refresh
        Next XII%
        '
        Call SETULTREG("DETACCD", REGDETACC&)
        Call SETULTREG("RESUVEN", CARESUVE&)
        '
        REGFAC& = 0
        REGDEB& = 0
        REGCRE& = 0
        REGCOB& = 0
        REGOTR& = 0
        REGRES& = 0
        REGPRV& = 0
        REGPRX& = 0
        KPX% = 0
        '
        FIN& = REGDETACC&
        '
        FIN& = RECUEC00.LINDICE.ListCount
        For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           RECUEC00.LINDICE.ListIndex = U& - 1
           REG& = Val(Mid$(RECUEC00.LINDICE.TEXT, 7, 8))
           Y$ = REGLEIDO$("DETACCD", REG&)
           TI% = Asc(Left$(Y$, 1))
           NC% = CVI(Mid$(Y$, 9, 2))
           '
           Select Case TI%
               Case 0
                  REGOTR& = REGOTR& + 1
                  Call GRAREG("DETOTRV", Y$, REGOTR&)
               Case 1
                  REGFAC& = REGFAC& + 1
                  Call GRAREG("DETFACV", Y$, REGFAC&)
               Case 2
                  REGFAC& = REGFAC& + 1
                  Call GRAREG("DETFACV", Y$, REGFAC&)
               Case 3
                  REGDEB& = REGDEB& + 1
                  Call GRAREG("DETDEBV", Y$, REGDEB&)
               Case 4
                  REGCOB& = REGCOB& + 1
                  Call GRAREG("DETCOBV", Y$, REGCOB&)
               Case 5
                  REGCRE& = REGCRE& + 1
                  Call GRAREG("DETCREV", Y$, REGCRE&)
               Case 6
                  REGCRE& = REGCRE& + 1
                  Call GRAREG("DETCREV", Y$, REGCRE&)
               Case 7
                  REGOTR& = REGOTR& + 1
                  Call GRAREG("DETOTRV", Y$, REGOTR&)
           End Select
           '
        Next U&
        '
        Call SETULTREG("DETFACV", REGFAC&)
        Call SETULTREG("DETDEBV", REGDEB&)
        Call SETULTREG("DETCREV", REGCRE&)
        Call SETULTREG("DETCOBV", REGCOB&)
        Call SETULTREG("DETOTRV", REGOTR&)
        '
        For NC% = 1 To 32767
          If Abs(SALINI#(NC%)) >= 0.005 Or Abs(SUMFAC#(NC%)) >= 0.005 Or Abs(SUMCOB#(NC%)) >= 0.005 Then
            If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
              Call MENSERR(24, "Movimiento Sobre Cliente" + Str$(NC%) + " Inexistente.")
            End If
          End If
          If NC% >= 32767 Then GoTo 11060
       Next NC%
       '
11060  For U& = 1 To ULTREG&("DEUDOR1")
            NC% = CVI(Left$(REGLEIDO$("DEUDOR1", U&), 2))
            If NC% > 0 Then
               If Abs(SALINI#(NC%)) >= 0.005 Or Abs(SUMFAC#(NC%)) >= 0.005 Or Abs(SUMCOB#(NC%)) >= 0.005 Then
                   X$ = REGBLAN$("RESUCCD")
                   Call REPLA(X$, MKI$(NC%), 1, 2)
                   Call REPLA(X$, MKD$(SALINI#(NC%)), 3, 8)
                   Call REPLA(X$, MKD$(SUMFAC#(NC%)), 11, 8)
                   Call REPLA(X$, MKD$(SUMCOB#(NC%)), 19, 8)
                   Call REPLA(X$, MKD$(SALINI#(NC%) + SUMFAC#(NC%) - SUMCOB#(NC%)), 27, 8)
                   Call REPLA(X$, MKD$(SUMFAC#(NC%) - SUMCOB#(NC%)), 35, 8)
                   '
                   REGRES& = REGRES& + 1
                   Call GRAREG("RESUCCD", X$, REGRES&)
                   '
                End If
            End If
       Next U&
       Call SETULTREG("RESUCCD", REGRES&)
       '
       Call CIERRARCH("*.*")
       Check1(4).Value = 1
       '
       If CAFAC% < 1 Then
           Screen.MousePointer = 1
           Call COMUNI("ADVERTENCIA:\No hay Facturas Registradas\en el Periodo Elegido.")
       End If
       '
       Return
        '
'*****************************************************************************
        '
11100   FIN& = ULTREG&("CONCIVEN")
        Check1(5).Value = 2
        '
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           X$ = REGLEIDO$("CONCIVEN", IL&)
           '
           JI% = IL&
           NCV%(JI%) = CVI(Mid$(X$, 7, 2))
           FFII%(JI%) = CVI(Mid$(X$, 9, 2))
           TII%(JI%) = Asc(Mid$(X$, 97, 1))
           NROV&(JI%) = CVS(Mid$(X$, 3, 4))
           IFAC#(JI%) = CVD(Mid$(X$, 99, 8)) + CVD(Mid$(X$, 123, 8))
           '
        Next IL&
        Check1(5).Value = 1
        '
        Return
        '
'*****************************************************************************
        '
11200   LU$ = LUDAT$
        Archi$ = "ICV" + ABREMES$ + ".DAT"
        ARCHIALT$ = "ICV" + Left$(ABREMES$, 3) + ".DAT"
        If EXISTE%(LU$ + Archi$) < 1 Then
            If EXISTE%(LU$ + ARCHIALT$) > 0 Then
                RESULTA% = COPYFILE%(LU$ + ARCHIALT$, LU$ + Archi$)
            End If
        End If
        '
        N2% = FILEOPEN(LU$ + Archi$, 0, 96)
        REGMAX% = 0
        '
        FIN& = LOF(N2%) / 96
        Check1(7).Value = 2
        '
        For JI% = 1 To FIN&
           '
           U& = JI%
           Call TRACE(20, U&, FIN&)
           '
           Get #N2%, JI%, X0$
           X1$ = Left$(X0$, 2)
           X2$ = Mid$(X0$, 3, 4)
           X3$ = Mid$(X0$, 7, 2)
           X4$ = Mid$(X0$, 9, 2)
           X5$ = Mid$(X0$, 11, 6)
           X6$ = Mid$(X0$, 17, 80)
           '
           TIXX$ = X1$
           NRO& = CVS(X2$)
           NC% = CVI(X3$)
           FEMI% = CVI(X4$)
           SUICC# = 0: For KK% = 1 To 8
                       If CVI(Mid$(X6$, 10 * KK% - 9, 2)) > 0 Then
                           SUICC# = SUICC# + Abs(CVD(Mid$(X6$, 10 * KK% - 7, 8)))
                       End If
                    Next KK%
                    '
           For KI% = 1 To CAFAC%
             If NCV%(KI%) = NC% Then
               If FFII%(KI%) = FEMI% Then
                 If NROV&(KI%) = NRO& Or (NRO& = 0 And TIS$(TII%(KI%)) = "AS") Then
                   If TIS$(TII%(KI%)) = TIXX$ Then
                     If Abs(Abs(IFAC#(KI%)) - SUICC#) < 0.005 Then
                       REGCON& = KI%
                       X$ = REGLEIDO$("CONCIVEN", REGCON&)
                       Call REPLA(X$, X6$, 17, 80)
                       Call REPLA(X$, MKI$(1), 155, 2)
                       Call GRAREG("CONCIVEN", X$, REGCON&)
                       GoTo 11290
                     End If
                   End If
                 End If
               End If
             End If
           Next KI%
           '
           LSet X0$ = String$(96, 0)        ' tacha movimiento en ICCmmaa.DAT
           Put #N2%, JI%, X0$
           '
11290   Next JI%
        Check1(7).Value = 1
        '
        Return
        '
'*****************************************************************************
        '
11300   PRINOCON& = 0
        FIN& = ULTREG&("CONCIVEN")
        '
        For IL& = 1 To FIN&
            Call TRACE(20, IL&, FIN&)
            X$ = REGLEIDO$("CONCIVEN", IL&)
            If Abs(CVD(Mid$(X$, 139, 8))) >= 0.005 Then
                If CVI(Mid$(X$, 155, 2)) <> 1 Then
                    TIII$ = Left$(X$, 2)
                    If TIII$ = "VC" Or TIII$ = "FC" Or TIII$ = "ND" Or TIII$ = "NC" Then
                       If PRINOCON& = 0 Then
                          PRINOCON& = IL&
                          GoTo 11390
                       End If
                    End If
                End If
            End If
        Next IL&
        '
11390   Check4.Value = 1
        Return
        '
'*****************************************************************************
        '
13000   FIN& = ULTREG&("CONCIVEN")
        '
        Check1(8).Value = 2
        SUMATO# = 0
        INCOMPLE% = 0
        '
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           X$ = REGLEIDO$("CONCIVEN", IL&)
           '
' AQUI SE MODIFICO
           INET# = CVD(Mid$(X$, 99, 8)) + CVD(Mid$(X$, 123, 8))
           SUIPU# = 0
           SUMATO# = SUMATO# + INET#
           SUMIVA# = SUMIVA# + CVD(Mid$(X$, 107, 8)) + CVD(Mid$(X$, 131, 8))
           SUMRNI# = SUMRNI# + CVD(Mid$(X$, 115, 8))
           '
           IMPUTA$ = Mid$(X$, 17, 80)
           AMO$ = "MOVISTO": PRIMO& = PRIMOSTO&(IL&): ULTMO& = ULTMOSTO&(IL&)
           If PRIMO& < 0 Then
               AMO$ = "MATVEN": PRIMO& = Abs(PRIMO&): ULTMO& = Abs(ULTMO&)
           End If
           '
           If MODOCALVE% > 0 Then
             NCLIE% = CVI(Mid$(X$, 7, 2))
             PCLI% = PIVACLI%(NCLIE%)
             If PCLI% < 0 Or PCLI% > 7 Then PCLI% = 1
             '
             TI% = Asc(Mid$(X$, 97, 1))
             CACO% = 0
             For JI% = 1 To 8
               CUEV%(JI%) = 0: IMXX#(JI%) = 0
             Next JI%
             '
             If (TI% > 0 And TI% <= 2) Or (TI% = 6 And PRIMO& > 0) Then
               CACO% = 1
               CUEV%(1) = COCUEVEN%(PCLI%)
               IMXX#(1) = INET#
               IMPUTA$ = MKI$(CUEV%(1)) + MKD$(Abs(IMXX#(1))) + String$(70, 0)
               SUMAIMPU# = INET#
               GoTo 13550
             End If
             '
             If TI% = 3 Then
               CACO% = 1
               CUEV%(1) = COCUEVEN%(8)
               IMXX#(1) = INET#
               IMPUTA$ = MKI$(CUEV%(1)) + MKD$(Abs(IMXX#(1))) + String$(70, 0)
               SUMAIMPU# = INET#
               GoTo 13550
             End If
             '
             If TI% = 6 Then
               CACO% = 1
               CUEV%(1) = COCUEVEN%(9)
               IMXX#(1) = INET#
               IMPUTA$ = MKI$(CUEV%(1)) + MKD$(Abs(IMXX#(1))) + String$(70, 0)
               SUMAIMPU# = INET#
               GoTo 13550
             End If
             '
             CACO% = 1
             CUEV%(1) = COCUEVEN%(10)
             IMXX#(1) = INET#
             IMPUTA$ = MKI$(CUEV%(1)) + MKD$(Abs(IMXX#(1))) + String$(70, 0)
             SUMAIMPU# = INET#
             GoTo 13550
             '
           End If
           '
13550      Call REPLA(X$, IMPUTA$, 17, 80)
           Call GRAREG("CONCIVEN", X$, IL&)
        Next IL&
        '
13600   Check1(8).Value = 1
        Screen.MousePointer = 11
        '
        For IL& = 1 To ULTREG&("CONCIVEN")
           X$ = Left$(REGLEIDO$("CONCIVEN", IL&), 96)
           LSet X0$ = X$
           Put #N2%, IL&, X0$
        Next IL&
        '
        For J& = IL& + 1 To LOF(N2%) / 96
           LSet X0$ = String$(96, 0)
           Put #N2%, J&, X0$
        Next J&
        '
        Close #N2%
        Screen.MousePointer = 1
        Return
        '
'*****************************************************************************
        '
14000   Check1(9).Value = 2
        Call COPYSTRU("IFACOMP", "IFACVEN")
        Call BLANARCH("SUDIVEN")
        Call BLANARCH("COSUDIV")
        Call BLANARCH("SUIVAVE")
        Call BLANARCH("IFACVEN")
        '
        TTXXY$ = Chr$(1) + Chr$(1) + MKI$(HAS%)
        J& = 0
        For IL& = 1 To ULTREG&("ASIDIAR")
           X$ = REGLEIDO$("ASIDIAR", IL&)
           If Left$(X$, 4) <> TTXXY$ Then
              J& = J& + 1
              Call GRAREG("ASIDIAR", X$, J&)
           End If
        Next IL&
        J0& = J&
        '
        For K& = J& + 1 To ULTREG&("ASIDIAR")
           Call GRAREG("ASIDIAR", String$(48, 0), K&)
        Next K&
        '
        FIN& = ULTREG&("CONCIVEN")
        '
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           X$ = REGLEIDO$("CONCIVEN", IL&)
           Y$ = String$(128, 0)
           '
           FF% = CVI(Mid$(X$, 9, 2))
           Call REPLA(Y$, MKI$(FF%), 1, 2)
           '
           NC% = CVI(Mid$(X$, 7, 2))
           TI% = Asc(Mid$(X$, 97, 1))
           VA% = Asc(Mid$(X$, 98, 1))
           '
           If TI% = 2 Then
              If NC% > 0 Then
                 If Mid$(RACLI$(NC%), 47, 1) = Chr$(0) Then
                    VA% = 2
                 End If
              End If
           End If
           '
           If VA% < 1 Or VA% > 3 Then VA% = 1
           NRO& = CVS(Mid$(X$, 3, 4))
           NUSUCUVE% = CVI(Mid$(X$, 11, 2))
           NDX$ = FORMDOC$(NRO&, TI%, VA%, NUSUCUVE%)
           
           If PIVACLI%(NC%) = 5 Then
              Call REPLA(NDX$, PREFEXPO$, 4, 4)
              Call REPLA(NDX$, LETREXPO$, 18, 1)
           End If
           '
           Call REPLA(Y$, NDX$, 3, 18)
           '
           CLI$ = "Clientes de Contado"
           iva$ = "No Posee"
           If NC% > 0 Then
                 CLI$ = Left$(RACLI$(NC%), 30)
                 iva$ = Mid$(RACLI$(NC%), 31)
               ElseIf NC% < 0 Then
                 RECUI& = Abs(NC%)
                 If RECUI& <= ULTREG&("CUITS") Then
                    RCUX$ = REGLEIDO$("CUITS", RECUI&)
                    CLI$ = Left$(RCUX$, 30)
                    iva$ = Mid$(RCUX$, 31, 16)
                 End If
           End If
           '
           Call REPLA(Y$, CLI$, 21, 30)
           Call REPLA(Y$, iva$, 51, 16)
           '
           RELIB$ = Y$
           SV$ = Left$(Y$, 64) + String$(56, 0) + Space$(8)
           '
           SIG% = Sgn(CVD(Mid$(X$, 139, 8)))
           NEFAC1# = CVD(Mid$(X$, 99, 8))
           NEFAC2# = CVD(Mid$(X$, 123, 8))
           NEFAC# = NEFAC1# + NEFAC2#
           IVAD1# = CVD(Mid$(X$, 107, 8))
           IVAD2# = CVD(Mid$(X$, 131, 8))
           IVAD# = IVAD1# + IVAD2#
           RNI# = CVD(Mid$(X$, 115, 8))
           TOFAC1# = CVD(Mid$(X$, 139, 8))
           '
           Call REPLA(Y$, MKD$(TOFAC1#), 67, 8)
           '
           NCVI% = NC%
           If NCVI% < 0 Then NCVI% = 0
           CUE% = CUEMA%(NCVI%)
           IMPO# = TOFAC1#
           GoSub 14600
           '
           YY$ = Y$: CUES$ = Space$(12)
               RGI& = CUE%
               If RGI& <= ULTREG&("CUENTAS") Then
                    CUES$ = Left$(REGLEIDO$("CUENTAS", RGI&), 12)
               End If
               Call REPLA(YY$, CUES$, 75, 12)
               If TOFAC1# > 0 Then
                     Call REPLA(YY$, MKD$(TOFAC1#), 87, 8)
                   Else
                     Call REPLA(YY$, MKD$(Abs(TOFAC1#)), 95, 8)
               End If
               '
               Call REGAPP("COSUDIV", YY$)
               '
           XX6$ = Mid$(X$, 17, 80)
           SUMA# = 0: IRET# = 0
           '
           For KI% = 1 To 8
              CUE% = CVI(Mid$(XX6$, 10 * KI% - 9, 2))
              IMM# = Abs(CVD(Mid$(XX6$, 10 * KI% - 7, 8))) * SIG%
              IMPO# = (-1) * IMM#
              GoSub 14600
              IMX# = (-1) * IMM#
              '
              If CUE% > 0 Then
                 If Abs(IMM#) > 0 Then
                    CUES$ = Space$(12)
                    RGI& = CUE%
                    If RGI& <= ULTREG&("CUENTAS") Then
                       CUES$ = Left$(REGLEIDO$("CUENTAS", RGI&), 12)
                    End If
                    Call REPLA(Y$, CUES$, 75, 12)
                    If IMX# > 0 Then
                        Call REPLA(Y$, MKD$(IMX#), 87, 8)
                      Else
                        Call REPLA(Y$, MKD$(Abs(IMX#)), 95, 8)
                    End If
                    '
                    If UCase$(Left$(NDX$, 2)) <> "AS" Then Call REGAPP("SUDIVEN", Y$)
                    Call REGAPP("COSUDIV", Y$)
                    '
                    If CUE% = CUERET1% Then
                       IRET# = IRET# + IMM#
                    End If
                    '
                    SUMA# = SUMA# + IMM#
                    Y$ = RELIB$
                 End If
              End If
           Next KI%
           '
           NEGRA1# = NEFAC1#: NEXEN# = 0
           If Abs(IVAD1#) < 0.005 Then
              NEXEN# = NEGRA1#
              NEGRA1# = 0
           End If
           NEGRA2# = NEFAC2#
           If Abs(IVAD2#) < 0.005 Then
              NEXEN# = NEXEN# + NEGRA2#
              NEGRA2# = 0
           End If
          '
           ' AGREGAR IVA
           If Abs(IVAD#) >= 0.005 Then
              CUES$ = CUEIVA$
              Call REPLA(Y$, CUES$, 75, 12)
              '
              CUE% = CUEIVA1%
              IMPO# = (-1) * IVAD#
              GoSub 14600
              '
              IMM# = (-1) * IVAD#
              If IMM# > 0 Then
                   Call REPLA(Y$, MKD$(IMM#), 87, 8)
                 Else
                   Call REPLA(Y$, MKD$(Abs(IMM#)), 95, 8)
              End If
              '
              If UCase$(Left$(NDX$, 2)) <> "AS" Then Call REGAPP("SUDIVEN", Y$)
              Call REGAPP("COSUDIV", Y$)
              '
              Y$ = RELIB$
           End If
           '
           ' AGREGAR IVA RNI
           If Abs(RNI#) >= 0.005 Then
              CUES$ = CUEADU$
              Call REPLA(Y$, CUES$, 75, 12)
              '
              CUE% = CUEADU1%
              IMPO# = (-1) * RNI#
              GoSub 14600
              '
              IMM# = (-1) * RNI#
              If IMM# > 0 Then
                   Call REPLA(Y$, MKD$(IMM#), 87, 8)
                 Else
                   Call REPLA(Y$, MKD$(Abs(IMM#)), 95, 8)
              End If
              '
              If UCase$(Left$(NDX$, 2)) <> "AS" Then Call REGAPP("SUDIVEN", Y$)
              Call REGAPP("COSUDIV", Y$)
              '
              Y$ = RELIB$
           End If
           '
           Call REPLA(SV$, MKD$(NEXEN#), 65, 8)
           Call REPLA(SV$, MKD$(NEGRA1#), 73, 8)
           Call REPLA(SV$, MKD$(IVAD1#), 81, 8)
           Call REPLA(SV$, MKD$(NEGRA2#), 89, 8)
           Call REPLA(SV$, MKD$(IVAD2#), 97, 8)
           Call REPLA(SV$, MKD$(NEGRA3#), 105, 8)
           Call REPLA(SV$, MKD$(IVAD3#), 113, 8)
           Call REPLA(SV$, MKD$(RNI#), 121, 8)
           Call REPLA(SV$, MKD$(IRET#), 129, 8)
           '
           Call REPLA(SV$, MKD$(TOFAC1#), 137, 8)
           '
           If Abs(TOFAC1#) < 0.005 Then
              Call REPLA(SV$, "ANULADA", 145, 8)
           End If
           '
           If UCase$(Left$(NDX$, 2)) <> "AS" Then
              '
              Call REGAPP("SUIVAVE", SV$)
              SUMAFAC#(11) = SUMAFAC#(11) + NEGRA1#
              SUMAFAC#(12) = SUMAFAC#(12) + NEXEN#
              SUMAFAC#(14) = SUMAFAC#(14) + IVAD1#
              SUMAFAC#(22) = SUMAFAC#(22) + NEGRA2#
              SUMAFAC#(23) = SUMAFAC#(23) + IVAD2#
              SUMAFAC#(15) = SUMAFAC#(15) + RNI#
              SUMAFAC#(16) = SUMAFAC#(16) + IRET#
              SUMAFAC#(17) = SUMAFAC#(17) + OTRO#
              '
              SUMAFAC#(13) = SUMAFAC#(11) + SUMAFAC#(12) + SUMAFAC#(22)
              SUMAFAC#(18) = SUMAFAC#(14) + SUMAFAC#(15) + SUMAFAC#(16) + SUMAFAC#(17) + SUMAFAC#(23) + SUMAFAC#(24)
              SUMAFAC#(19) = SUMAFAC#(11) + SUMAFAC#(14) + SUMAFAC#(15) + SUMAFAC#(16) + SUMAFAC#(17)
              SUMAFAC#(25) = SUMAFAC#(22) + SUMAFAC#(23) + SUMAFAC#(24)
              SUMAFAC#(20) = SUMAFAC#(12)
              SUMAFAC#(21) = SUMAFAC#(19) + SUMAFAC#(20) + SUMAFAC#(25)
              '
              For XII% = 11 To 25
                  Text1(XII%).TEXT = CSTRING$(MKD$(SUMAFAC#(XII%)), 4, 10, 2, 2)
                  Text1(XII%).Refresh
              Next XII%
              '
           End If
           TTXYZ$ = String$(64, 0)
           Call REPLA(TTXYZ$, Mid$(SV$, 3, 18), 1, 18)
           Call REPLA(TTXYZ$, Mid$(SV$, 21, 30), 21, 30)
           Call REPLA(TTXYZ$, MKD$(NEGRA1# + NEXEN# + NEGRA2#), 51, 8)
           Call REPLA(TTXYZ$, MKI$(NC%), 59, 2)
           Call REPLA(TTXYZ$, MKS$(NRO), 61, 4)
           Call REGAPP("IFACVEN", TTXYZ$)
           '
        Next IL&
        '
        Call CIERRARCH("*.*")
        Call GRARGEN("CONCIVEN", TRIM$(CORRESP$))
        '
        Call HEADERS("SUDIVEN", "")
        Call HEADERS("SUDIVEN", "Corresponde a: " + CORRESP$)
        '
        Call HEADERS("SUIVAVE", "")
        Call HEADERS("SUIVAVE", "Corresponde a: " + CORRESP$)
        '
        Call CIERRARCH("*.*")
        '
        Return


14600   BUS$ = TTXXY$ + MKI$(CUE%)
        IMPOX# = IMPO#
        '
        For K& = J0& + 1 To J&
            Z$ = REGLEIDO$("ASIDIAR", K&)
            If Left$(Z$, 6) = BUS$ Then
               IMPO# = (CVD(Mid$(Z$, 39, 8))) + ((Int(100 * IMPO# + 0.5)) / 100)
               Z$ = REPLACE$(Z$, MKD$(IMPO#), 39, 8)
               Call GRAREG("ASIDIAR", Z$, K&)
               GoTo 14690
            End If
        Next K&
        Z$ = REGBLAN$("ASIDIAR")
        Z$ = REPLACE$(Z$, BUS$, 1, 6)
        Z$ = REPLACE$(Z$, MKD$((Int(100 * IMPO# + 0.5)) / 100), 39, 8)
        J& = J& + 1: K& = J&
        Call GRAREG("ASIDIAR", Z$, J&)
        '
14690   SUTOTA# = SUTOTA# + IMPOX#
        Return
        '


15000   CORRESP$ = AJUSTI$(MESAN, 12)
        FGENE% = HOY(HO$): If INCOMPLE% > 0 Then FGENE% = 0
        HGENE% = HORANUM(Time$): If INCOMPLE% > 0 Then HGENE% = 0
        '
        REBLA$ = String$(300, 0)
        Call REPLA(REBLA$, CORRESP$, 1, 12)
        Call REPLA(REBLA$, MKI$(FGENE%), 13, 2)
        Call REPLA(REBLA$, MKI$(HGENE%), 15, 2)
        For KI% = 0 To 31
            Call REPLA(REBLA$, MKD$(SUMAFAC#(KI%)), 45 + 8 * KI%, 8)
        Next KI%
        '
        NX% = FILEOPEN%(LU$ + "IRESMEN.DAT", 0, 300)
        Dim REBLB As String * 300
        '
        For KI% = 1 To LOF(NX%) / 300
            Get #NX%, KI%, REBLB$
            If Left$(REBLB$, 12) = CORRESP$ Then
                REBLB$ = REBLA$
                Put #NX%, KI%, REBLB$
                GoTo 15090
            End If
        Next KI%
        '
        KI% = 1 + LOF(NX%) / 300
        REBLB$ = REBLA$
        Put #NX%, KI%, REBLB$
        '
15090   Close #NX%
        '
        FEX = CVI(Mid$(REBLA$, 13, 2))
        Text4(0).TEXT = FECHATEX$(FEX)
        If FEX < 1 Then Text4(0).TEXT = "00/00/00"
        Text4(0).Refresh
        HOX% = CVI(Mid$(REBLA$, 15, 2))
        Text4(1) = HORATEX$(HOX%)
        Text4(1).Refresh
        '
        Return

End Sub
'

Sub CARGARESMEN()
        '
        REBLA$ = Space$(12) + String$(288, 0)
        CORRESP$ = AJUSTI$(MESAN, 12)
        NX% = FILEOPEN%(LUDAT$ + "IRESMEN.DAT", 0, 300)
        Dim REBLB As String * 300
        '
        For KI% = 1 To LOF(NX%) / 300
            Get #NX%, KI%, REBLB$
            If Left$(REBLB$, 12) = CORRESP$ Then
                REBLA$ = REBLB$
            End If
        Next KI%
        '
        FEX = CVI(Mid$(REBLA$, 13, 2))
        Text4(0).TEXT = FECHATEX$(FEX)
        If FEX < 1 Then Text4(0).TEXT = "00/00/00"
        Text4(0).Refresh
        HOX% = CVI(Mid$(REBLA$, 15, 2))
        Text4(1) = HORATEX$(HOX%)
        Text4(1).Refresh
        '
        For KI% = 0 To 7
            SUMAFAC#(KI%) = CVD(Mid$(REBLA$, 45 + 8 * KI%, 8))
            Text1(KI%).TEXT = CSTRING$(MKD$(SUMAFAC#(KI%)), 4, 11, 2, 2)
            Text1(KI%).Refresh
        Next KI%
        '
        For KI% = 11 To 25
            SUMAFAC#(KI%) = CVD(Mid$(REBLA$, 45 + 8 * KI%, 8))
            Text1(KI%).TEXT = CSTRING$(MKD$(SUMAFAC#(KI%)), 4, 10, 2, 2)
            Text1(KI%).Refresh
        Next KI%
        '
        Close #NX%

End Sub


Sub VALICUEIVA()
     '
     Screen.MousePointer = 11
     YACUE% = 0
     '
     LU$ = LUDAT$
     TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
     For I% = 0 To 7
        TIS$(I%) = Mid$(TDAT$, 3 * I% + 1, 2)
     Next I%
     '
     CUEIVA$ = AJUSTI$(Control$("SVD451S", "CUEIVA"), 12)
     CUEADU$ = AJUSTI$(Control$("SVD451S", "CUEADU"), 12)
     CUERET$ = AJUSTI$(Control$("SVD451S", "CUERET"), 12)
     CUEDEU$ = AJUSTI$(Control$("AMACLI", "ECUECON"), 12)
     '
     CUEIVA1% = 0: CUEADU1% = 0: CUERET1% = 0: CUEDEUV1% = 0
     For U& = 1 To ULTREG&("ECUECON")
        X$ = REGLEIDO$("ECUECON", U&)
        If Left$(X$, 12) = CUEIVA$ Then CUEIVA1% = CVI(Mid$(X$, 41, 2)): DECIVA$ = Mid$(X$, 13, 28)
        If Left$(X$, 12) = CUEADU$ Then CUEADU1% = CVI(Mid$(X$, 41, 2)): DECADU$ = Mid$(X$, 13, 28)
        If Left$(X$, 12) = CUERET$ Then CUERET1% = CVI(Mid$(X$, 41, 2)): DECRET$ = Mid$(X$, 13, 28)
        If Left$(X$, 12) = CUEDEU$ Then CUEDEUV1% = CVI(Mid$(X$, 41, 2)): DECRET$ = Mid$(X$, 13, 28)
     Next U&
     '
     If CUEIVA1% < 1 Or CUEADU1% < 1 Or CUERET1% < 1 Then
         Call MENSERR(24, "Faltan Cuentas Contables Correspondientes\a I.V.A. Responsables Inscriptos y No Inscriptos.")
         Screen.MousePointer = 1
         Exit Sub
     End If
     '
     If CUEDEUV1% < 1 Then
         Call MENSERR(24, "Faltan Cuentas Contables Correspondientes\a Deudores por Ventas (omisión).")
         Screen.MousePointer = 1
         Exit Sub
     End If
     '
     MODOCALVE% = 1
     If MODOCALVE% > 0 Then
        '
        For KKI% = 0 To 11
          CUEVENTA$(KKI%) = AJUSTI$(Control$("SVD451S", "CUEVE-" + TRIM$(Str$(KKI%))), 12)
        Next KKI%
        '
        For U& = 1 To ULTREG&("ECUECON")
          X$ = REGLEIDO$("ECUECON", U&)
          For KKI% = 0 To 11
            If Left$(X$, 12) = CUEVENTA$(KKI%) Then COCUEVEN%(KKI%) = CVI(Mid$(X$, 41, 2)): DECIVA$ = Mid$(X$, 13, 28)
          Next KKI%
        Next U&
        '
        For KKI% = 0 To 11
          If COCUEVEN%(KKI%) < 1 Then
            Call MENSERR(24, "Faltan Cuentas Contables de Venta.")
            Screen.MousePointer = 1
            Exit Sub
          End If
        Next KKI%
        '
     End If
     '
     Call CIERRARCH("*.*")
     YACUE% = 1
     '
End Sub
'

Sub ARMAMES()
     '
     CORRESP$ = TRIM$(MESAN)
     CORRESQ$ = MESANO$(CORRESP$, DES%, HAS%)
     If TRIM$(CORRESQ$) = "" Then
       Call MENSERR(24, "Proceso Imposible !!!\Seleccione Período Mensual.")
       CORRESP$ = ""
       Exit Sub
     End If
     '
     FEX = DES%: DESFE$ = FECHATEX$(FEX)
     FEX = HAS%: HASFE$ = FECHATEX$(FEX)
     CORRESQ$ = TRIM$(CORRESP$)
     ABREMES$ = Left$(CORRESQ$, 3) + Mid$(DESFE$, 7, 2)
     '
End Sub
'

Sub CARCUECLI()
    '
10  YACUE% = 0
    FIN& = ULTREG&("DEUDOR1")
    '
    CANCUEN% = ULTREG&("CUENTAS")
    YACLICON% = 0
    For JJL& = 1 To FIN&
        Call TRACE(20, JJL&, FIN&)
        CLX$ = REGLEIDO$("DEUDOR1", JJL&)
        NC% = CVI(Left$(CLX$, 2))
        If NC% = 0 Then
            Call REPLA(CLX$, "Clientes de Contado", 3, 30)
            Call GRAREG("DEUDOR1", CLX$, JJL&)
            YACLICON% = 1
        End If
        '
        If NC% >= 0 Then
            If NC% <= 32768 Then
                EXICLI%(NC%) = 1
                If MODOCLAVE% = 1 Then
                    GRUCLII% = CVI(Mid$(REGLEIDO$("DEUDOR2", JJL&), 125, 2))
                  Else
                    GRUCLII% = 1
                    PIVCLII% = Asc(Mid$(REGLEIDO$("DEUDOR2", JJL&), 104, 1))
                    If PIVCLII% = 7 Then GRUCLII% = 2
                    If PIVCLII% = 5 Then GRUCLII% = 3
                End If
If GRUCLII% <> 1 Then
AAA = BBB
End If
                '                       validar GRUCLII%
                For GCI% = 1 To CAGRUCOCLI%
                    If GRUCLII% = GRUCOC%(GCI%) Then GoTo 49
                Next GCI%
                GRUCLII% = 0
                '
49              GRCOCLI%(NC%) = GRUCLII%
                '
                CUEMA%(NC%) = CVI(Mid$(REGLEIDO$("DEUDOR2", JJL&), 127, 2))
                If CUEMA%(NC%) < 0 Or CUEMA%(NC%) > CANCUEN% Then
                    CUEMA%(NC%) = 0
                End If
                '
            End If
        End If
    Next JJL&
    YACUE% = 1
    '
End Sub
'

Sub CONASIDIA()
    '
    MESA$ = TRIM$(MESAN)
    MESB$ = MESANO$(MESA$, DES%, HAS%)
    If MESB$ <> MESA$ Then
      Exit Sub
    End If
    '
    Screen.MousePointer = 11
    J& = 0
    Call BLANARCH("!ASIDIAFI")
    TTXXY$ = Chr$(1) + Chr$(1) + MKI$(HAS%)
    For IL& = 1 To ULTREG&("ASIDIAR")
        X$ = REGLEIDO$("ASIDIAR", IL&)
        If Left$(X$, 4) = TTXXY$ Then
            CUENU& = CVI(Mid$(X$, 5, 2))
            IDEB# = CVD(Mid$(X$, 39, 8))
            If Abs(IDEB#) >= 0.005 Then
                IHAB# = 0
                If IDEB# < 0 Then
                    IHAB# = Abs(IDEB#)
                    IDEB# = 0
                End If
                CUES$ = Space$(12)
                If CUENU& > 0 Then
                    If CUENU& <= ULTREG&("CUENTAS") Then
                        CUES$ = Left$(REGLEIDO$("CUENTAS", CUENU&), 12)
                    End If
                End If
                DENOCUES$ = ECOARCH$("ECUECON", CUES$)
                '
                Z$ = REGBLAN$("ASIDIAFI")
                Call REPLA(Z$, CUES$, 1, 12)
                Call REPLA(Z$, DENOCUES$, 13, 30)
                Call REPLA(Z$, MKD$(IDEB#), 43, 8)
                Call REPLA(Z$, MKD$(IHAB#), 51, 8)
                '
                J& = J& + 1
                Call GRAREG("!ASIDIAFI", Z$, J&)
            End If
        End If
    Next IL&
    Call SETULTREG("!ASIDIAFI", J&)
    '
    For K1& = 1 To J&
        For K2& = K1& + 1 To J&
            X1$ = REGLEIDO$("!ASIDIAFI", K1&)
            X2$ = REGLEIDO$("!ASIDIAFI", K2&)
            If Left$(X1$, 12) > Left$(X2$, 12) Then
                X0$ = X1$
                X1$ = X2$
                X2$ = X0$
                Call GRAREG("!ASIDIAFI", X1$, K1&)
                Call GRAREG("!ASIDIAFI", X2$, K2&)
            End If
        Next K2&
    Next K1&
    Call CIERRARCH("!ASIDIAFI")
    '
    Screen.MousePointer = 1
    VTB% = VENTABU%("ASIDIAFI")
    '
End Sub
'

Sub BACKUPMEN(MODO%, HAYERRO%)
    '
    ' MODO% = 1 ... existencia y acceso a directorio
    '
    Screen.MousePointer = 11
    Dim X As String * 2
    Dim ARESPA$(24)
    CARESPA% = 18
    '
    HAYERRO% = 0
    Call ARMAMES
    If Len(ABREMES$) <> 5 Then
        Call COMUNI("Período '" + ABREMES$ + "' no Procesable.\  \Consulte a su Soporte de Sistemas.")
        HAYERRO% = 1
        Exit Sub
    End If
    '
    ULOG$ = Left$(CurDir, 2)
    LUBAK$ = LUDAT$ + ABREMES$
    If Mid$(LUBAK$, 2, 1) <> ":" Then
       LUBAK$ = ULOG$ + LUBAK$
    End If
    EXTE$ = Mid$(DESFE$, 8, 1) + Mid$(DESFE$, 4, 2)
    '
    ARESPA$(0) = "ICV" + ABREMES$
    ARESPA$(1) = "DETACCD"
    ARESPA$(2) = "RESUCCD"
    ARESPA$(3) = "DETFACV"
    ARESPA$(4) = "DETDEBV"
    ARESPA$(5) = "DETCREV"
    ARESPA$(6) = "DETCOBV"
    ARESPA$(7) = "DETOTRV"
    ARESPA$(8) = "VEPROVIN"
    ARESPA$(9) = "VEPROVIX"
    ARESPA$(10) = "SUDIVEN"
    ARESPA$(11) = "SUIVAVE"
    ARESPA$(12) = "CONCIVEN"
    ARESPA$(13) = "RESUVEN"
    ARESPA$(14) = "REVECLI"
    ARESPA$(15) = "REVEPIE"
    ARESPA$(16) = "RECOBCLI"
    ARESPA$(17) = "COSUDIV"
    ARESPA$(18) = "IFACVEN"
    '
40  vuelta% = 1
41  ARXX$ = LUBAK$
    BRXX$ = CurDir
    On Error GoTo 43
    ChDir ARXX$
    ChDir BRXX$
    On Error GoTo 0
    GoTo 50
    '
43  ERNU$ = TRIM$(Str$(Err.Number))
    Resume 44
44  On Error GoTo 0
    If vuelta% = 1 Then
        On Error Resume Next
        MkDir ARXX$
        On Error GoTo 0
        vuelta% = 2
        GoTo 41
    End If
    '
    CRXX$ = REPLACAR(ARXX$, "\", "/")
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Apertura\"
    MENSA$ = MENSA$ + "para RESPALDO de la Siguiente Carpeta:\  \"
    MENSA$ = MENSA$ + CRXX$ + "\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Problemas de Red - No hay Acceso a U.L. '" + Left$(ARXX$, 2) + "'.\"
    MENSA$ = MENSA$ + "   2.- Usuario de Red sin Derechos\"
    MENSA$ = MENSA$ + "       de Creación de Directorios.\"
    MENSA$ = MENSA$ + "   3.- Espacio en disco Insuficiente.\  \"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Apertura de Carpeta")
    If RESPU% = 4 Then
        GoTo 40
    End If
    HAYERRO% = 1
    Exit Sub
    '
50  For I% = 0 To CARESPA%
        If TRIM$(ARESPA$(I%)) <> "" Then
            NX% = FILEOPEN(LUBAK$ + "\" + ARESPA$(I%) + "." + EXTE$, 0, 2)
            If NX% < 1 Then
                HAYERRO% = 1
                Exit Sub
            End If
            Get #NX%, 1, X$
            Put #NX%, 1, X$
            Close #NX%
        End If
    Next I%
    '
    '
    If MODO% = 2 Then
        LUDA$ = LUDAT$
        If Mid$(LUDA$, 2, 1) <> ":" Then
            LUDA$ = Left$(CurDir, 2) + LUDA$
        End If
        '
        Call CIERRARCH("*.*")
        For I% = 0 To CARESPA%
            If TRIM$(ARESPA$(I%)) <> "" Then
                '
                ORIGI$ = LUDA$ + ARESPA$(I%) + ".DAT"
                DESTI$ = LUBAK$ + "\" + ARESPA$(I%) + "." + EXTE$
                '
                If COPYFILE%(ORIGI$, DESTI$) <> 1 Then
                  HAYERRO% = 1
                  Exit Sub
                End If
                '
            End If
        Next I%
    End If
    '
    '
    If MODO% = 3 Then
        LUDA$ = LUDAT$
        If Mid$(LUDA$, 2, 1) <> ":" Then
            LUDA$ = Left$(CurDir, 2) + LUDA$
        End If
        C1$ = TRIM$(MESAN)
        C2$ = TRIM$(Text4(0).TEXT)
        C3$ = TRIM$(Text4(1).TEXT)
        '
        Call CIERRARCH("*.*")
        If TRIM$(CORRAR$("CONCIVEN")) = TRIM$(C1$) Then
            GoTo 80
        End If
        '
        For I% = 0 To CARESPA%
            If TRIM$(ARESPA$(I%)) <> "" Then
                '
                ORIGI$ = LUBAK$ + "\" + ARESPA$(I%) + "." + EXTE$
                DESTI$ = LUDA$ + ARESPA$(I%) + ".DAT"
                '
                If COPYFILE%(ORIGI$, DESTI$) <> 1 Then
                  HAYERRO% = 1
                  Exit Sub
                End If
                '
            End If
        Next I%
        '
        Call HEADERS("DETACCX", "")
        Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
        Call HEADERS("RESUCCX", "")
        Call HEADERS("RESUCCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
        Call GRARGEN("CONCIVEN", TRIM$(C1$))
        '
80      For I% = 0 To CARESPA%
            If TRIM$(ARESPA$(I%)) <> "" Then
                '
                Call HEADERS(ARESPA$(I%), "")
                Call HEADERS(ARESPA$(I%), "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
                '
            End If
        Next I%
        '
        Call HEADERS("SUDIVEN", "")
        Call HEADERS("SUDIVEN", "Corresponde a: " + TRIM$(C1$))
        '
        Call HEADERS("SUIVAVE", "")
        Call HEADERS("SUIVAVE", "Corresponde a: " + TRIM$(C1$))
        '
    End If
    '
99  Screen.MousePointer = 1
    '
End Sub
'

