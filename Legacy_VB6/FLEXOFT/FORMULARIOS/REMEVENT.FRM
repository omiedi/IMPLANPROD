VERSION 5.00
Begin VB.Form REMEVENT 
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
      Picture         =   "REMEVENT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "Configuracion de Funcionamiento"
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
      Picture         =   "REMEVENT.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   74
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
      Picture         =   "REMEVENT.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   7
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
      TabIndex        =   9
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
         TabIndex        =   50
         Top             =   1470
         Width           =   4215
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   210
            ScaleHeight     =   75
            ScaleWidth      =   3735
            TabIndex        =   67
            Top             =   2625
            Width           =   3795
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   68
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
            TabIndex        =   60
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
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   57
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
            TabIndex        =   56
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
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   420
            Width           =   3795
         End
      End
      Begin VB.TextBox MESAN 
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
         TabIndex        =   0
         Text            =   " "
         Top             =   735
         Width           =   2220
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
         TabIndex        =   45
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
            Picture         =   "REMEVENT.frx":0DB6
            Style           =   1  'Graphical
            TabIndex        =   4
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
            Picture         =   "REMEVENT.frx":10C0
            Style           =   1  'Graphical
            TabIndex        =   5
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
         TabIndex        =   44
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
            Picture         =   "REMEVENT.frx":13CA
            Style           =   1  'Graphical
            TabIndex        =   3
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
            Picture         =   "REMEVENT.frx":16D4
            Style           =   1  'Graphical
            TabIndex        =   2
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
         TabIndex        =   30
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
            TabIndex        =   73
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
            TabIndex        =   72
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
            TabIndex        =   71
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
            TabIndex        =   69
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
            TabIndex        =   65
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
            TabIndex        =   64
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
            Index           =   20
            Left            =   3780
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   36
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
            Index           =   16
            Left            =   840
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   70
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
            TabIndex        =   66
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
         TabIndex        =   14
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
            TabIndex        =   28
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
            TabIndex        =   26
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   29
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
            TabIndex        =   27
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
         TabIndex        =   11
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
            TabIndex        =   62
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
            TabIndex        =   61
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
            TabIndex        =   13
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
            TabIndex        =   12
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
         TabIndex        =   1
         Top             =   735
         Width           =   175
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4305
         Picture         =   "REMEVENT.frx":19DE
         Top             =   650
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4305
         Picture         =   "REMEVENT.frx":1CE8
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
         TabIndex        =   10
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
      Picture         =   "REMEVENT.frx":1FF2
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Picture         =   "REMEVENT.frx":22FC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Finaliza la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   510
      Left            =   9530
      Picture         =   "REMEVENT.frx":2446
      Top             =   7260
      Width           =   2010
   End
End
Attribute VB_Name = "REMEVENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'

Private Sub Command10_Click()
   OPCLAVEN.Show 1
End Sub

Private Sub Command2_Click()
    Call CALREVEN
End Sub
'

Private Sub Command3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
        MESAN.TEXT = MESASEL$
        Call FRESHALL
        Call CARGARESMEN
        MESAN.SetFocus
    End If
End Sub
'

Private Sub Command5_Click()
   PRINTFORM
End Sub

Private Sub Command6_Click()
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
    RESUMENVE.MESAN.TEXT = MESAN.TEXT
    RESUMENVE.Text4(0).TEXT = Text4(0).TEXT
    RESUMENVE.Text4(1).TEXT = Text4(1).TEXT
    Command8.Enabled = True
    RESUMENVE.Show 1
End Sub

Private Sub Command9_Click()
    If FECHANUM(Text4(0).TEXT) < 1 Then
        OPX% = COMALTER%("La Generación de Resumen Mensual\no se ha Completado.\  \Puede Ahora Re-Calcular el Resumen\o Abandonar el Proceso.\\Re-Calcular\Abandonar")
        If OPX% = 1 Then
            Call CALREVEN
        End If
        Exit Sub
    End If
    '
    Command9.Enabled = False
    MODOSABO% = 0
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "SABO") > 0 Or InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "FERVI") > 0 Then MODOSABO% = 1
    '
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
      IPERD# = 0
      For KKJ% = 1 To 3
        IGRAA# = CVD(Mid$(XX$, 57 + 16 * KKJ%, 8))
        IIVAA# = CVD(Mid$(XX$, 65 + 16 * KKJ%))
        If KKJ% = 3 Then
          If MODOSABO% = 1 Then
            IPERD# = IIVAA#
            IIVAA# = 0
          End If
        End If
        KKI% = 2   ' por default 21 %
        If Abs(IGRAA#) >= 0.005 Then
          TASARE = 100 * IIVAA# / IGRAA#
          If TASARE > 18.5 And TASARE < 19.5 Then KKI% = 1
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
      If MODOSABO% = 1 Then
        Call REPLA(XX$, MKD$(IPERD#), 105, 8)
      End If
      Call GRAREG("SUIVA19", XX$, U&)
    Next U&
    Call SETULTREG("SUIVA19", FIN&)
    Call CIERRARCH("SUIVA19")
    '
    Call HEADERS("SUIVA19", "")
    Call HEADERS("SUIVA19", "Corresponde a: " + MESAN.TEXT)
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

Private Sub Label1_Click(Index As Integer)

End Sub

Private Sub MESAN_Change()
   TXT$ = MESAN.TEXT
   If MESANO$(TXT$, DES%, HAS%) = TXT$ Then
      If Val(CONTROL$("", "CIEMEVEN")) < HAS% Then
          Image1.Visible = False
          Image2.Visible = True
         Else
          Image2.Visible = False
          Image1.Visible = True
      End If
   End If
End Sub

Private Sub MESAN_DblClick()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
        MESAN.TEXT = MESASEL$
        Call FRESHALL
        Call CARGARESMEN
    End If
End Sub

Private Sub MESAN_GOTFOCUS()
    '
    If CARGAMEN% = 0 Then
        '
        Refresh
        Screen.MousePointer = 11
        '
        CORRESP$ = TRIM$(CORRAR$("CONCIVEN"))
        MESAN.TEXT = CORRESP$
        If TRIM$(MESAN.TEXT) = "" Then
            Call CARSELMESA
            MESAN.TEXT = Left$(REGLEIDO$("SELMESA", 1), 12)
            If TRIM$(MESAN.TEXT) = "" Then
                Call COMUNI("Imposible Inicializar Período Mensual.\Consulte a su Soporte de Sistemas.")
                Call FINAL("")
            End If
        End If
        '
        'FEX = FEGEAR%("CONCIVEN")
        'Text4(0).TEXT = FECHATEX$(FEX)
        'If FEX < 1 Then Text4(0).TEXT = "00/00/00"
        'HOX% = HOGEAR%("CONCIVEN")
        'Text4(1) = HORATEX$(HOX%)
        '
        Call CARGARESMEN
        CARGAMEN% = 1
        '
    End If
    Screen.MousePointer = 1
    '
End Sub
'

Private Sub MESAN_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
        Command2.SetFocus
    End If
    '
End Sub

Private Sub MESAN_LostFocus()
    '
    MESA$ = MESAN.TEXT
    If TRIM$(MESA$) = "" Then
        On Error Resume Next
        MESAN.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    '
    MESB$ = MESANO$(MESA$, DES%, HAS%)
    If MESB$ <> MESA$ Then
        If MESB$ <> "" Then
            MESAN.TEXT = MESB$
        End If
        MESAN.SetFocus
        Call CARGARESMEN
        If FECHANUM(Text4(0).TEXT) < 1 Then
            Command2.SetFocus
        End If
        '
    End If
    '
End Sub

Private Sub Text1_Click(Index As Integer)
    MESAN.SetFocus
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
   TXT$ = MESAN.TEXT
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
   TXT$ = MESAN.TEXT
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


Private Sub Text3_Change(Index As Integer)

End Sub
