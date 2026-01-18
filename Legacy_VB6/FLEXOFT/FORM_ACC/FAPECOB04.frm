VERSION 5.00
Begin VB.Form FAPECOB04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Pendientes de Cobranza "
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   4410
      TabIndex        =   21
      Top             =   2240
      Width           =   1905
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pantalla"
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
         Left            =   315
         TabIndex        =   23
         Top             =   315
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Impresión"
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
         Left            =   315
         TabIndex        =   22
         Top             =   590
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6405
      TabIndex        =   46
      Top             =   2240
      Width           =   2010
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pl.Financiero"
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
         Index           =   10
         Left            =   315
         TabIndex        =   48
         Top             =   525
         Width           =   1485
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tabla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   315
         TabIndex        =   47
         Top             =   270
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Cust."
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
      Left            =   1890
      Picture         =   "FAPECOB04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Acceso a Maestro de Clientes"
      Top             =   5040
      Width           =   650
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00004000&
      Height          =   5685
      Left            =   8610
      ScaleHeight     =   5625
      ScaleWidth      =   1530
      TabIndex        =   38
      Top             =   0
      Width           =   1590
      Begin VB.PictureBox Picture16 
         Height          =   4620
         Left            =   0
         ScaleHeight     =   4560
         ScaleWidth      =   1020
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   4620
            Left            =   0
            Picture         =   "FAPECOB04.frx":0442
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
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
         Picture         =   "FAPECOB04.frx":27DA
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   2900
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   52
         Top             =   2900
         Width           =   650
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
         Height          =   630
         Left            =   105
         Picture         =   "FAPECOB04.frx":2AE4
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Pendientes en Moneda Extranjera"
         Top             =   2190
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   105
         Picture         =   "FAPECOB04.frx":2DEE
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Listado de Pendientes a Cobrar por Rango de Fecha."
         Top             =   1560
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   42
         Top             =   3750
         Width           =   645
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command5 
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
         Height          =   630
         Left            =   105
         Picture         =   "FAPECOB04.frx":30F8
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Opciones por Default"
         Top             =   930
         Width           =   650
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
         Height          =   630
         Left            =   105
         Picture         =   "FAPECOB04.frx":3242
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   300
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "FAPECOB04.frx":338C
         Top             =   4050
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Deuda"
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
      Left            =   5355
      Picture         =   "FAPECOB04.frx":52AE
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Composición de Deuda a Fecha"
      Top             =   6195
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   2625
      Sorted          =   -1  'True
      TabIndex        =   33
      Top             =   6195
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Rango de Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   4410
      TabIndex        =   24
      Top             =   210
      Width           =   4005
      Begin VB.TextBox Text10 
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
         Left            =   2625
         MaxLength       =   5
         TabIndex        =   30
         Top             =   210
         Width           =   855
      End
      Begin VB.CommandButton Command9 
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
         Left            =   3465
         TabIndex        =   29
         Top             =   210
         Width           =   175
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00E0E0E0&
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
         Left            =   315
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   " "
         Top             =   630
         Width           =   3330
      End
      Begin VB.TextBox Text8 
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
         Left            =   2625
         MaxLength       =   5
         TabIndex        =   27
         Text            =   " "
         Top             =   1050
         Width           =   855
      End
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
         Height          =   360
         Left            =   3465
         TabIndex        =   26
         Top             =   1050
         Width           =   175
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00E0E0E0&
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
         Left            =   315
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1470
         Width           =   3330
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         Left            =   210
         TabIndex        =   32
         Top             =   315
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta :"
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
         Left            =   210
         TabIndex        =   31
         Top             =   1155
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formato de Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   525
      TabIndex        =   17
      Top             =   5985
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Listado Tabulado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   420
         TabIndex        =   20
         Top             =   420
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Plan Financiero de Cobranza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   420
         TabIndex        =   19
         Top             =   735
         Width           =   3270
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Gráfico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   420
         TabIndex        =   18
         Top             =   1155
         Width           =   1380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3990
         Y1              =   1050
         Y2              =   1050
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4300
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   4005
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fechas de Emisión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   730
         Left            =   140
         TabIndex        =   57
         Top             =   3430
         Width           =   3750
         Begin VB.CommandButton Command15 
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
            Left            =   1680
            TabIndex        =   59
            Top             =   315
            Width           =   175
         End
         Begin VB.CommandButton Command14 
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
            Left            =   3465
            TabIndex        =   58
            Top             =   315
            Width           =   175
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
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
            Left            =   735
            TabIndex        =   63
            Top             =   315
            Width           =   960
         End
         Begin VB.Label Label123 
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
            Left            =   15
            TabIndex        =   62
            Top             =   360
            Width           =   645
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
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
            Left            =   2520
            TabIndex        =   61
            Top             =   315
            Width           =   960
         End
         Begin VB.Label Label7 
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
            Left            =   1800
            TabIndex        =   60
            Top             =   360
            Width           =   645
         End
      End
      Begin VB.TextBox Text12 
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
         Left            =   840
         TabIndex        =   55
         Text            =   " "
         Top             =   2100
         Width           =   435
      End
      Begin VB.CommandButton Command17 
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
         Left            =   1260
         TabIndex        =   54
         Top             =   2100
         Width           =   175
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1575
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "Todos - Sin Clasificación"
         Top             =   2100
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Barrida Automática"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   315
         TabIndex        =   50
         Top             =   1155
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Barrida Automática"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   315
         TabIndex        =   49
         Top             =   315
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rango de Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   830
         Left            =   140
         TabIndex        =   13
         Top             =   2505
         Width           =   3750
         Begin VB.CommandButton Command11 
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
            Left            =   1650
            TabIndex        =   37
            Top             =   345
            Width           =   175
         End
         Begin VB.TextBox Text5 
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
            Left            =   700
            TabIndex        =   35
            Top             =   345
            Width           =   1000
         End
         Begin VB.CommandButton Command6 
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
            Left            =   3465
            TabIndex        =   15
            Top             =   345
            Width           =   175
         End
         Begin VB.TextBox Text6 
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
            Left            =   2460
            TabIndex        =   14
            Top             =   345
            Width           =   1000
         End
         Begin VB.Label Label4 
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
            Left            =   1770
            TabIndex        =   36
            Top             =   480
            Width           =   645
         End
         Begin VB.Label Label3 
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
            Left            =   0
            TabIndex        =   16
            Top             =   480
            Width           =   645
         End
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0E0E0&
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
         Left            =   315
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Todos - Sin Clasificación"
         Top             =   1680
         Width           =   3435
      End
      Begin VB.CommandButton Command4 
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
         Left            =   3570
         TabIndex        =   10
         Top             =   1260
         Width           =   175
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
         Left            =   3150
         TabIndex        =   9
         Text            =   " "
         Top             =   1260
         Width           =   435
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   315
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " Todos - Sin Clasificación"
         Top             =   840
         Width           =   3435
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
         Left            =   3570
         TabIndex        =   6
         Top             =   420
         Width           =   175
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
         Left            =   3150
         TabIndex        =   5
         Text            =   " "
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.Pago"
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
         TabIndex        =   56
         Top             =   2150
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
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
         Left            =   735
         TabIndex        =   12
         Top             =   1365
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Clientes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   735
         TabIndex        =   8
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ordenamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   4410
      TabIndex        =   0
      Top             =   3255
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Emisión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   2205
         TabIndex        =   64
         Top             =   840
         Width           =   1620
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Vencimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2205
         TabIndex        =   3
         Top             =   600
         Width           =   1500
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Alfabético"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   525
         TabIndex        =   2
         Top             =   840
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   525
         TabIndex        =   1
         Top             =   600
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Por Fecha:"
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
         Left            =   1995
         TabIndex        =   66
         Top             =   315
         Width           =   2325
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Por Cliente:"
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
         Left            =   315
         TabIndex        =   65
         Top             =   315
         Width           =   2325
      End
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu facturapendientedecobranza 
         Caption         =   "Emitir Listado de Facturas Pendientes de  Cobranza"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu consultasylistados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu pendientesenmonedaextranjera 
         Caption         =   "Pendientes en Moneda Extranjera"
      End
      Begin VB.Menu pendientesacobrarporrangodefecha 
         Caption         =   "Pendientes a Cobrar por Rango de Fecha."
      End
   End
End
Attribute VB_Name = "FAPECOB04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TICUE%(32767)

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      Check2.Value = 0
      Text1 = ""
   End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = 1 Then
      Check1.Value = 0
      Text3 = ""
   End If
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Command2.Enabled = False
   Call COMPODEU
   Command2.Enabled = True
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG04/AMACLI", "Archivos Maestros")
   Command12.Enabled = True
End Sub

Sub Command13_Click()
    '
    Dim DOCULAR$(), SALPESO#(), SALDOLA#()
    '
    Command13.Enabled = False
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ <> "BIMONETA" Then
       If MODOMONE$ <> "MULTIMON" Then
          Call MENSERR(24, "Solo Disponible en Sistemas Bi- O Multi-Monetarios")
          GoTo 99
       End If
    End If
    '
    Screen.MousePointer = 11
    J& = 0
    '
    NUDOCU& = 0: RAMAX& = 0
    FIN& = ULTREG&("CARBIMON")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("CARBIMON", U&)
      NUDOCX$ = TRIM$(Mid$(XX$, 13, 18))
      DEBEPES# = ROUND#(CVD(Mid$(XX$, 65, 8)))
      HABEPES# = ROUND#(CVD(Mid$(XX$, 73, 8)))
      DEBEDOL# = ROUND#(CVD(Mid$(XX$, 97, 8)))
      HABEDOL# = ROUND#(CVD(Mid$(XX$, 105, 8)))
      '
      If AJUSTI$(NUDOCX$, 18) = AJUSTI$(Mid$(XX$, 45, 18), 18) Then ' ES MOVIMIENTO CABEZA
        If Abs(DEBEPES# - DEBEDOL#) < 0.005 Then
          If Abs(HABEPES# - HABEDOL#) < 0.005 Then
            TTXX$ = "Se han Encontrado Registros Incorrectos\en la Tabla de Control Multimoneda.\  \¿ Desea Recuperar la Tabla Dañada ?"
            If COMALTER%(TTXX$ + "\\Si, Recuperar\No, Continuar") = 1 Then
              Call RECUBIMO
              Call COMUNI("Tabla Recuperada.\Vuelva Ahora a Pedir el Listado")
              GoTo 99
            End If
          End If
        End If
      End If
      '
      For JJ& = 1 To NUDOCU&
         If DOCULAR$(JJ&) = NUDOCX$ Then
            SALPESO#(JJ&) = SALPESO#(JJ&) + DEBEPES# - HABEPES#
            SALDOLA#(JJ&) = SALDOLA#(JJ&) + DEBEDOL# - HABEDOL#
            GoTo 19
         End If
      Next JJ&
      '
      NUDOCU& = NUDOCU& + 1
      If NUDOCU& > RAMAX& Then
         RAMAX& = RAMAX& + 128
         ReDim Preserve DOCULAR$(RAMAX&), SALPESO#(RAMAX&), SALDOLA#(RAMAX&)
      End If
      DOCULAR$(NUDOCU&) = NUDOCX$
      SALPESO#(NUDOCU&) = SALPESO#(NUDOCU&) + DEBEPES# - HABEPES#
      SALDOLA#(NUDOCU&) = SALDOLA#(NUDOCU&) + DEBEDOL# - HABEDOL#
      '
19  Next U&
    '
    JJ& = 0: TEBUX$ = ""
    For U& = 1 To NUDOCU&
       Call TRACE(20, U&, NUDOCU&)
       If Abs(SALPESO#(U&)) >= 0.005 Or Abs(SALDOLA#(U&)) >= 0.005 Then
          TEBUX$ = TEBUX$ + "**" + AJUSTI$(DOCULAR$(U&), 18)
       End If
    Next U&
    '
    FIN& = ULTREG&("CARBIMON")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("CARBIMON", U&)
       NUDOCX$ = "**" + AJUSTI$(Mid$(XX$, 13, 18), 18)
       If InStr(TEBUX$, NUDOCX$) > 0 Then
          '
          ' ENJUAGUE ASEMBLI-GRAMY POR ERROR DE MONEDA
          MONEDAPLI% = CVI(Mid$(XX$, 63, 2))
          If MONEDAPLI% > ULTREG&("TAMONED") Then
            MONEDAPLI% = MONEDAPLI% Mod 6
            Call REPLA(XX$, MKI$(MONEDAPLI%), 63, 2)
            Call GRAREG("CARBIMON", XX$, U&)
            If TRIM$(Mid$(XX$, 45, 18)) = TRIM$(Mid$(XX$, 13, 18)) Then
              Call REPLA(XX$, MKI$(MONEDAPLI%), 31, 2)
              Call GRAREG("CARBIMON", XX$, U&)
            End If
          End If
          ' FIN ENJUAGUE ASEMBLI-GRAMY
          '
          J& = J& + 1
          Call GRAREG("CARBIMOX", XX$, J&)
       End If
    Next U&
    '
    Call SETULTREG("CARBIMOX", J&)
    Call CIERRARCH("CARBIMOX")
    '
    Call FILESORT("CARBIMOX", "", 1, 6, "ASC")
    Call CONECRUN("*LISBIMON", "")
    '
99  Screen.MousePointer = 1
    Command13.Enabled = True

End Sub

Sub RECUBIMO()
    '
    FIN& = ULTREG&("CARBIMON")
    Call ABRECABAN
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("CARBIMON", U&)
      NUDOCX$ = AJUSTI$(Mid$(XX$, 13, 18), 18)
      If TRIM$(NUDOCX$) <> "" Then
         DEBEPES# = ROUND#(CVD(Mid$(XX$, 65, 8)))
         HABEPES# = ROUND#(CVD(Mid$(XX$, 73, 8)))
         DEBEDOL# = ROUND#(CVD(Mid$(XX$, 97, 8)))
         HABEDOL# = ROUND#(CVD(Mid$(XX$, 105, 8)))
         NCX% = CVI(Mid$(XX$, 3, 2))
         '
         If AJUSTI$(Mid$(XX$, 45, 18), 18) = NUDOCX$ Then ' ES MOVIMIENTO CABEZA
           If Abs(DEBEPES# - DEBEDOL#) < 0.005 Then
             If Abs(HABEPES# - HABEDOL#) < 0.005 Then
               YY$ = String$(256, 0)
               Call GRAREG("CARBIMON", YY$, U&)
               '
               For JJ& = U& + 1 To FIN&
                 ZZ$ = REGLEIDO$("CARBIMON", JJ&)
                 If AJUSTI$(Mid$(ZZ$, 13, 18), 18) = NUDOCX$ Then
                   YY$ = String$(256, 0)
                   Call GRAREG("CARBIMON", YY$, JJ&)
                 End If
               Next JJ&
               '
               DOCULA$ = TRIM$(NUDOCX$)
               With CajaBanc
                  .FindFirst "NuClien = " & NCX% & " AND CODICOM_LAR = '" & DOCULA$ & "'"
201               If .NoMatch = False Then
                    .Edit
                    !Mone_Emis = 1
                    !Tipo_Cam = 1
                    !Impo_Dola = 0
                    .Update
                    .FindNext "NuClien = " & NCX% & " AND CODICOM_LAR = '" & DOCULA$ & "'"
                    GoTo 201
                  End If
               End With
               '
               With SaDeuPen
                  .FindFirst "NuClien = " & NCX% & " AND CODICOM_LAR = '" & DOCULA$ & "'"
202               If .NoMatch = False Then
                    .Edit
                    !SaldDebe_Dola = 0
                    !SaldAcre_Dola = 0
                    !Mone_Emis = 1
                    !Ibru_Dola = 0
                    .Update
                    .FindNext "NuClien = " & NCX% & " AND CODICOM_LAR = '" & DOCULA$ & "'"
                    GoTo 202
                 End If
               End With
               '
             End If
           End If
         End If
      End If
    Next U&
    '
    JJ& = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("CARBIMON", U&)
      If XX$ <> String$(256, 0) Then
         JJ& = JJ& + 1
         Call GRAREG("CARBIMON", XX$, JJ&)
      End If
    Next U&
    Call SETULTREG("CARBIMON", JJ&)
    Call CIERRARCH("CARBIMON")
    '
End Sub

Private Sub Command14_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   Label8 = VALACT1$("SELFECHA")
End Sub

Private Sub Command15_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   Label9 = VALACT1$("SELFECHA")
End Sub

Private Sub Command16_Click()
    ' envio de deuda por mail
    Call COPYSTRU("FAPECOBA", "FAPECOX")
    For U& = 1 To ULTREG&("DEUDOR1")
       X1$ = REGLEIDO$("DEUDOR1", U&)
       NC% = CVI(Left$(X1$, 2))
       RFF$ = RECFILT$("FAPECOBA", 1, MKI$(NC%))
       If Len(RFF$) > 0 Then
         UJ& = 0
         SUMADEU# = 0
         For UL& = 1 To Len(RFF$) Step 4
           UK& = CVS(Mid$(RFF$, UL&, 4))
           X$ = REGLEIDO$("FAPECOBA", UK&)
           UJ& = UJ& + 1
           Call GRAREG("FAPECOX", X$, UJ&)
           SUMADEU# = SUMADEU# + CVD(Mid$(X$, 73, 8))
         Next UL&
         Call SETULTREG("FAPECOX", UJ&)
         '
         If SUMADEU# > 0.005 Then
            JJ& = 0
            N1% = FILEOPEN%("C:\FLEXOFT\CODEUCLI.TXT", 2, 128)
            FIN& = ULTREG&("FAPECOX")
            For U3& = 1 To FIN&
               '
               Call TRACE(20, U3&, FIN&)
               X$ = REGLEIDO$("FAPECOX", U3&)
               '
               TTXX$ = Space$(80)
               NC% = CVI(Left$(X$, 2))
               NFS$ = Mid$(X$, 49, 12)
               FF% = CVI(Mid$(X$, 61, 2))
               FVI% = CVI(Mid$(X$, 71, 2))
               SALDEU# = CVD(Mid$(X$, 73, 8))
               IMORIG# = CVD(Mid$(X$, 63, 8))
               '
               Call REPLA(TTXX$, NFS$, 1, 12)
               FEX = FF%
               Call REPLA(TTXX$, FECHATEX$(FEX), 15, 8)
               Call REPLA(TTXX$, FORMATNUM$(IMORIG#, "F12.2"), 23, 12)
               FEX = FVI%
               Call REPLA(TTXX$, FECHATEX$(FEX), 37, 8)
               Call REPLA(TTXX$, FORMATNUM$(SALDEU#, "F12.2"), 45, 12)
               '
               If JJ& < 1 Then
                 HOII% = HOY(HOS$)
                 Print #N1%, "." + Space$(20) + "Buenos Aires, " + FETEXTO(HOII%)
                 Print #N1%, "Sres:"
                 Print #N1%, RASOCLI$(NC%)
                 Print #N1%, DOMICLI$(NC%)
                 Print #N1%, LOCACLI$(NC%)
                 Print #N1%, "   "
                 Print #N1%, "At: " + CONTACTCLI$(NC%)
                 Print #N1%, "   "
                 Print #N1%, "De nuestra consideración:"
                 Print #N1%, "                          Nos dirigimos a uds con el"
                 Print #N1%, "fin de recordarles el estado de su apreciada cuenta:"
                 Print #N1%, "   "
                 Print #N1%, "Factura     Fecha      Importe   Vencim.       Saldo"
                 Print #N1%, "-------------------------------------------------------"
               End If
               '
               JJ& = JJ& + 1
               Print #N1%, TTXX$
            Next U3&
            Print #N1%, "---------------------------------------------------------"
            Print #N1%, "Su Saldo Total a la Fecha: $ " + FORMATNUM$(SUMADEU#, "F12.2")
            Print #N1%, ""
            Print #N1%, "Agradecemos cancelar dicho saldo en cualquiera de las "
            Print #N1%, "formas siguientes:"
            Print #N1%, ""
            Print #N1%, "- Transferencia a nuestra CC 120-1160/1 del Banco Francés"
            Print #N1%, ". BBVA, Sucursal 120, Titular Ernesto Pablo Bauer, Número"
            Print #N1%, ". de CBU: 01701209  200  000  001  160  17"
            Print #N1%, ""
            Print #N1%, "- Cheque a la orden de Ernesto Pablo Bauer, comunicándonos"
            Print #N1%, ". por e-mail o fax (011-4524-1284) días y horarios de pago."
            Print #N1%, ""
            Print #N1%, "En caso de haber ya efectivizado el pago solicitado, favor"
            Print #N1%, "ignorar la presente y comunicarlo por email o fax a fin de"
            Print #N1%, "emitirles el recibo correspondiente."
            Print #N1%, ""
            Print #N1%, "Muchas gracias por su atención. Con cordiales saludos,"
            Print #N1%, ""
            Print #N1%, ".                                    María Florencia Ricci"
            Print #N1%, ".                                      FLEXOFT Soluciones"
            Close #N1%
            '
            '*******************************************************
            '
            TTYY$ = LOADFILE$("C:\FLEXOFT\CODEUCLI.TXT")
            MAI_DESTI$ = RASOCLI$(NC%)
            MAI_DIREL$ = EMAILCLI$(NC%)
            MAI_ASUNT$ = "Estado y Detalle de Deuda"
            MAI_TEXTO$ = TTYY$
            '
            JKI% = 0
            MAI_ADJUN$(JKI%) = ""
            '
            Call GENERAMAIL
            '
         End If
         Screen.MousePointer = 1
         '
       End If
    Next U&
    '
    Call COMUNI("Colocados en Bandeja de Salida")
    '
End Sub

Private Sub Command17_Click()
    Call CommSel("CONPAG", Command17, Text12)

End Sub

'
Private Sub Command2_Click()
    '
    Command2.Enabled = False
    Command10.Enabled = False
    '
    '\\\OT-06-0533-RG-23/11/06///
    GRUX$ = ""
    VENDX$ = ""
    N1% = Val(Text1)
    If N1% > 0 Then GRUX$ = Str(N1%) + " - " + ECOARCH("GRUCOCLI", MKI$(N1%))
    N2% = Val(Text3)
    If N2% > 0 Then VENDX$ = Str(N2%) + " - " + ECOARCH("VENDEDOR", Chr$(N2%))
    '\\\OT-06-0533-RG-FIN///
    '
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    FINF = FECHANUM(Text5.TEXT)
    FLIM = FECHANUM(TRIM$(Text6.TEXT)) ': If FLIM < HOY(HOS$) Then FLIM = HOY(HOS$)
    FEMI1% = FECHANUM(TRIM$(Label9))
    FEMI2% = FECHANUM(TRIM$(Label8))
    '
    If NCMIN% < 1 Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text10.SetFocus
       GoTo 99
    End If
    '
    If NCMAX% < 1 Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text8.SetFocus
       GoTo 99
    End If
    '
    If ECOARCH$("DEUDOR1", MKI$(NCMIN%)) = "" Then
       Call MENSERR(24, "Cuenta Inexistente")
       Text10.SetFocus
       GoTo 99
    End If
    '
    If ECOARCH$("DEUDOR1", MKI$(NCMAX%)) = "" Then
       Call MENSERR(24, "Cuenta Inexistente")
       Text8.SetFocus
       GoTo 99
    End If
    '
    If FLIM < FINF Then
      Call MENSERR(24, "Rango de Fechas No Procesable")
      GoTo 99
    End If
    '
    ' VALIDA FECHA DE EMISION
    If FEMI2% < FEMI1% Then
      Call MENSERR(24, "Rango de Fechas de Emisión No Procesable")
      GoTo 99
    End If
    '
    Call FAPECOB2
    '
    '\\\OT-06-0533-RG-11/12/06///
    If GRUX$ <> "" Then
       Call HEADERS("FAPECOBD", "Grupo:" + GRUX$)
       Call HEADERS("FAPECOBA", "Grupo:" + GRUX$)
    End If
    '
    If VENDX$ <> "" Then
       Call HEADERS("FAPECOBD", "Vendedor:" + VENDX$)
       Call HEADERS("FAPECOBA", "Vendedor:" + VENDX$)
    End If
    '\\\OT-06-0533-RG-FIN///
    '
    If Option1(4).Value = True Then  ' pantalla
         If Option1(3).Value = False Then
            Call FINAL("*COFAPEA1")
            GoTo 99
          Else
            Call FINAL("*COFAPEA2")
            GoTo 99
         End If
       Else                          ' impresora
       '\\\OT-06-0503-RG-23/11/06///
         If Option1(10).Value = True Then
           Call PLANFINA
           GoTo 99
         End If
        '\\\OT-06-0503-RG-FIN///
         MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
         If MODOMONE$ = "BIMONETA" Or MODOMONE$ = "MULTIMON" Then
'             If Option1(3).Value = False Then
'                Call FINAL("*LIFAPE1D")
'                GoTo 99
'              Else
'                Call FINAL("*LIFAPE2D")
'                GoTo 99
'             End If
              'COMENTADO SE MODIFICO POR ESTAS 7 LINEAS
              'OPCIONES DE ORDENAMIENTO POR VENCIMIENTO Y POR EMISION
              If Option1(3).Value = True Or Option1(11).Value = True Then
                 Call FINAL("*LIFAPE2D")
                 GoTo 99
               Else
                 Call FINAL("*LIFAPE1D")
                 GoTo 99
              End If
           Else
             If Option1(3).Value = False Then
                If Check1.Value = 1 Or Check2.Value = 1 Then
                     Call FILESORT("FAPECOBA", "", 12, 12, "ASC")
                     Call TRALOCAL("FAPECOBA", "FAPECOBA")
                     Call BLANARCH("FAPECOBA")
                     JJ& = 0: GRAX$ = ""
                     For U& = 1 To ULTREG&("!FAPECOBA")
                        X$ = REGLEIDO$("!FAPECOBA", U&)
                        GRACT$ = TRIM$(Mid$(X$, 101, 28))
                        If U& > 1 Then
                           If GRACT$ <> GRAX$ Then
                              JJ& = JJ& + 1
                              Call GRAREG("FAPECOBA", String$(512, "="), JJ&)
                           End If
                        End If
                        JJ& = JJ& + 1
                        Call GRAREG("FAPECOBA", X$, JJ&)
                        GRAX$ = GRACT$
                     Next U&
                     Call SETULTREG("FAPECOBA", JJ&)
                     Call FINAL("*LIFAPEG1")
                   Else
                     Call FINAL("*LIFAPEA1")
                End If
                GoTo 99
              Else
                If Check1.Value = 1 Or Check2.Value = 1 Then
                     Call FILESORT("FAPECOBA", "", 12, 12, "ASC")
                     Call TRALOCAL("FAPECOBA", "FAPECOBA")
                     Call BLANARCH("FAPECOBA")
                     JJ& = 0: GRAX$ = ""
                     For U& = 1 To ULTREG&("!FAPECOBA")
                        X$ = REGLEIDO$("!FAPECOBA", U&)
                        GRACT$ = TRIM$(Mid$(X$, 101, 28))
                        If U& > 1 Then
                           If GRACT$ <> GRAX$ Then
                              JJ& = JJ& + 1
                              Call GRAREG("FAPECOBA", String$(512, "="), JJ&)
                           End If
                        End If
                        JJ& = JJ& + 1
                        Call GRAREG("FAPECOBA", X$, JJ&)
                        GRAX$ = GRACT$
                     Next U&
                     Call SETULTREG("FAPECOBA", JJ&)
                     Call FINAL("*LIFAPEG2")
                   Else
                     Call FINAL("*LIFAPEA2")
                End If
                GoTo 99
             End If
         End If
    End If
    '
    '
99  Command10.Enabled = True
    Label9 = ""
    Label8 = ""
    Command2.Enabled = True
    '
End Sub

Private Sub Command2_GotFocus()
    '
    If CARFAPEN% = 0 Then
        Refresh
        Screen.MousePointer = 11
        Call FAPECOB1
    End If
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Command4_Click()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Command5_Click()
    Call FAPECOB1
End Sub

Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub



Sub Command7_Click()
    '
    Command7.Enabled = False
    '
    OPT% = ALTERNA%("Pendientes por Rango Vencimiento\Mas de 30 Dias Emisión\Atraso de Cobranza")
    Select Case OPT%
      Case 1
       GoTo 10
      Case 2
       Call PENMASTREINTA
       GoTo 99
      Case 3
       Call ATRACOBRA
       GoTo 99
      Case Else
       GoTo 99
    End Select
10  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If DES% < 1 Then GoTo 99
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
    '
    CORRESP$ = PERIO$
    Call HEADERS("VEXRAFE2", "")
    Call HEADERS("VEXRAFE2", "Tipo: Pendiente por Rango de Vencimiento")
    Call HEADERS("VEXRAFE2", "Corresponde a: " + CORRESP$)
    '
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    SQLX$ = SQLX$ + "AND INT(CDBL(FeVencim))<= " & HAS1# & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FeVencim))>= " & DES1# & " "
    SQLX$ = SQLX$ + "AND (SaldDebe)>= 0.005 "
    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    Set VEXRAFETMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    With VEXRAFETMP
      '
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("No Existen Vencimientos en el Rango de Fecha Seleccionado.")
        GoTo 99
      End If
      On Error GoTo 0
      J& = 0
      Y$ = REGBLAN("VEXRAFE2")
      Do While .EOF = False
          '
          NC% = !NuClien
          NFS$ = !CODICOM_LAR
          FEM% = CVINT(!FechEmisi)
          FVI% = CVINT(!FEVENCIM)
          SALDEU# = !SaldDebe - !SaldAcre
          '
          Call REPLA(Y$, MKI$(NC%), 1, 2)
          Call REPLA(Y$, CODICLI$(NC%), 3, 24)
          Call REPLA(Y$, RASOCLI$(NC%), 27, 30)
          Call REPLA(Y$, NFS$, 57, 18)
          Call REPLA(Y$, MKI$(FEM%), 75, 2)
          Call REPLA(Y$, MKI$(FVI%), 77, 2)
          Call REPLA(Y$, MKD$(SALDEU#), 79, 8)
          '
          J& = J& + 1
          Call GRAREG("VEXRAFE2", Y$, J&)
          '
        .MoveNext
        '
      Loop
      '
      Call SETULTREG("VEXRAFE2", J&)
      Call FILESORT("VEXRAFE2", "VEXRAFE2", 5, 5, "ASC")
      Call FILESORT("VEXRAFE2", "VEXRAFE2", 1, 1, "ASC")
      '
    End With
    '
    Call FINAL("*LIVERAF1")
    '
99  Command7.Enabled = True
    '
End Sub

Sub PENMASTREINTA()
    Call SELECHO("VENDEDOR")
    VENSEL% = XVALO(VALACT1("VENDEDOR"))
    DescVend$ = TRIM$(VALACT2$("VENDEDOR"))
    '\\\OT-08-0230-OD-09/05/08///
    
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    SQLX$ = SQLX$ + "AND (SaldDebe)>= 0.005 or (SaldAcre)>= 0.005 "
    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    Set VEXRAFETMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    With VEXRAFETMP
      '
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("No Hay Facturas Pendientes de Cobranza")
        GoTo 99
      End If
      On Error GoTo 0
      J& = 0
      Call HEADERS("VEXRAFE2", "")
      Call HEADERS("VEXRAFE2", "Tipo: Mas de 30 Días de Emisión")
      Y$ = REGBLAN("VEXRAFE2")
      If VENSEL% > 0 Then
        Call HEADERS("VEXRAFE2", "Vendedor: " + DescVend$)
      End If
      Do While .EOF = False
          '
          If DIENTRE(CVINT(!FechEmisi), HOY(HO$)) < 30 Then GoTo 50
          NC% = !NuClien
          NUMVEND% = VENDCLI%(NC%)
          If VENSEL% < 1 Or VENSEL% = NUMVEND% Then
            NFS$ = !CODICOM_LAR
            FVI% = CVINT(!FEVENCIM)
            SALDEU# = !SaldDebe - !SaldAcre
            FECEMI% = CVINT(!FechEmisi)
            '
            Call REPLA(Y$, MKI$(NC%), 1, 2)
            Call REPLA(Y$, CODICLI$(NC%), 3, 24)
            Call REPLA(Y$, RASOCLI$(NC%), 27, 30)
            Call REPLA(Y$, NFS$, 57, 18)
            Call REPLA(Y$, MKI$(FECEMI%), 75, 2)
            Call REPLA(Y$, MKI$(FVI%), 77, 2)
            Call REPLA(Y$, MKD$(SALDEU#), 79, 8)
            '
            J& = J& + 1
            Call GRAREG("VEXRAFE2", Y$, J&)
          End If
          '
50        .MoveNext
        '
      Loop
      If J& < 1 Then
        Call COMUNI("No Hay Facturas Pendientes de Cobranza\con más de 30 Días")
        GoTo 99
      End If
      '
      Call SETULTREG("VEXRAFE2", J&)
      Call FILESORT("VEXRAFE2", "VEXRAFE2", 5, 5, "ASC")
      Call FILESORT("VEXRAFE2", "VEXRAFE2", 1, 1, "ASC")
      '
    End With
    '
    Call FINAL("*LIVERAF1")
    '
99  Command7.Enabled = True
    '\\\OT-08-0230-OD-FIN///
    '
End Sub
Private Sub Command8_Click()
    Call SELECHO("DEUDOR1")
    NCMAX% = XVALO(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$(NCMAX%)
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("DEUDOR1")
    NCMIN% = XVALO(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$(NCMIN%)
    End If
End Sub

Private Sub facturapendientedecobranza_Click()
     Call Command2_Click
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      
      Me.Option1(1).BackColor = &HFCD7B6
      Me.Option1(2).BackColor = &HFCD7B6
      Me.Option1(3).BackColor = &HFCD7B6
      Me.Option1(4).BackColor = &HFCD7B6
      Me.Option1(5).BackColor = &HFCD7B6
      Me.Option1(6).BackColor = &HFCD7B6
      Me.Option1(7).BackColor = &HFCD7B6
      Me.Option1(8).BackColor = &HFCD7B6
      Me.Option1(9).BackColor = &HFCD7B6
      Me.Option1(10).BackColor = &HFCD7B6
      Me.Option1(11).BackColor = &HFCD7B6
      
      Me.Check1.BackColor = &HFCD7B6
      Me.Check2.BackColor = &HFCD7B6
      
      MARCOBARRA.Top = 4220
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
      ' &H8000000F
    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If InStr(EPAC$, "FLEXOFT") > 0 Then
       Command2.Top = 3150
    End If
    '
    If RECLEN("FAPECOBA") >= 128 Then
       Check1.Visible = True
       Check2.Visible = True
    End If
    '
    Call FAPECOB1
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub Label8_Click()
    Call Command14_Click
End Sub

Private Sub Label9_Click()
    Call Command15_Click
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Option1(0) = False
        Option1(1) = True
    End If
End Sub

Private Sub pendientesacobrarporrangodefecha_Click()
   Call Command7_Click
End Sub

Private Sub pendientesenmonedaextranjera_Click()
    Call Command13_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       Text3.SetFocus
    End If
End Sub

Private Sub Text10_Change()
    If Abs(XVALO(Text10.TEXT)) < 32767 Then
        NCMIN% = XVALO(Text10.TEXT)
        Text9.TEXT = RASOCLI$(NCMIN%)
      Else
        Text10.TEXT = ""
        Text9 = ""
    End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       Text8.SetFocus
    End If
End Sub

Private Sub Text10_LostFocus()
    NCMIN% = XVALO(Text10.TEXT)
    If NCMIN% < 1 Then
       Text9.TEXT = ""
       Exit Sub
    End If
    If RASOCLI$(NCMIN%) = "" Then
       Text9.TEXT = ""
       Exit Sub
    End If
    Text9.TEXT = RASOCLI$(NCMIN%)
End Sub

Private Sub Text12_Change()
  Text11 = ECOARCH$("CONPAG", MKI$(XVALO(Text12)))
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text1_LostFocus()
    If XVALO(Text1.TEXT) < 256 Then
        GRUCLI% = XVALO(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GRUCLI%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
    Text1.TEXT = ""
    Text2.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       Text5.SetFocus
    End If
End Sub

Private Sub Text3_LostFocus()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
            Text4.TEXT = VENDX$
            Exit Sub
        End If
    End If
    Text3.TEXT = ""
    Text4.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text4_CLICK()
    Text3.SetFocus
End Sub

Private Sub Text1_Change()
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
      ElseIf XVALO(Text1.TEXT) < 256 Then
        Check1.Value = 0
        GRUCLI% = XVALO(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GRUCLI%))
        If GRUX$ <> "" Then
          Text2.TEXT = GRUX$
          Exit Sub
        End If
    End If
End Sub

Private Sub Text3_Change()
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
      ElseIf XVALO(Text3.TEXT) < 256 Then
        Check2.Value = 0
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
          Text4.TEXT = VENDX$
          Exit Sub
        End If
    End If
End Sub

Private Sub Text5_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       Text6.SetFocus
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       Text10.SetFocus
    End If
End Sub

Private Sub Text8_Change()
    If Abs(XVALO(Text8.TEXT)) < 32767 Then
        NCMAX% = XVALO(Text8.TEXT)
        Text7.TEXT = RASOCLI$(NCMAX%)
      Else
        Text8.TEXT = ""
        Text7 = ""
    End If
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("DEUDOR1")
    NCMIN% = XVALO(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$(NCMIN%)
    End If
End Sub

Private Sub Text6_DBLCLICK()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("DEUDOR1")
    NCMAX% = XVALO(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$(NCMAX%)
    End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       Command2.SetFocus
    End If
End Sub

Private Sub Text8_LostFocus()
    NCMAX% = XVALO(Text8.TEXT)
    If NCMAX% < 1 Then
       Text7.TEXT = ""
       Exit Sub
    End If
    If RASOCLI$(NCMAX%) = "" Then
       Text7.TEXT = ""
       Exit Sub
    End If
    Text7.TEXT = RASOCLI$(NCMAX%)
End Sub

Private Sub Text9_Click()
    Text10.SetFocus
End Sub

Sub COMPODEU()
    '
    Dim NROPO&(8196), CLIPO%(8196)
    Dim SARRACLI#(32767), SUDEUCLI#(32767)
    Dim TIS$(10)
    TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    For I% = 0 To 7
        TIS$(I%) = Mid$(TDAT$, 3 * I% + 1, 2)
    Next I%
    '
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) < 1 Then
       Exit Sub
    End If
    '
    FLIM% = FECHANUM(VALACT1$("SELFECHA"))
    '
    FIN& = ULTREG&("FACTUR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("FACTUR", U&)
      NCLIE% = CVI(Left$(X$, 2))
      FEDOC% = CVI(Mid$(X$, 7, 2))
      If FEDOC% <= FLIM% Then
        TI% = Asc(Mid$(X$, 13, 1))
        If TI% < 0 Or TI% > 7 Then TI% = 2
        SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
        IMPO# = CVD(Mid$(X$, 15, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 23, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 31, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 49, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 57, 8))
        IMPO# = SG% * IMPO#
        SARRACLI#(NCLIE%) = SARRACLI#(NCLIE%) + IMPO#
        SARRATO# = SARRATO# + IMPO#
      End If
    Next U&
    '
    FIN& = ULTREG&("COBRAN")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("COBRAN", U&)
      NCLIE% = CVI(Left$(X$, 2))
      FEDOC% = CVI(Mid$(X$, 7, 2))
      If FEDOC% <= FLIM% Then
      IMPO# = (-1) * CVD(Mid$(X$, 15, 8))
        SARRACLI#(NCLIE%) = SARRACLI#(NCLIE%) + IMPO#
        SARRATO# = SARRATO# + IMPO#
      End If
    Next U&
    '
    ' LEER SVD202 - DESCARTAR DOCUMENTOS POSTERIORES A FLIM%
    UJ& = 0: SUDEUTO# = 0
    FIN& = ULTREG&("SVD202")
    For UI& = 1 To FIN&
       Call TRACE(20, UI&, FIN&)
       X$ = REGLEIDO$("SVD202", UI&)
       FF% = CVI(Mid$(X$, 19, 2))
       If FF% <= FLIM% Then
         NC% = CVI(Left$(X$, 2))
         If NC% > 0 Then
           If Abs(SARRACLI#(NC%)) >= 0.005 Then
             TI% = Asc(Mid$(X$, 3, 1))
             If TI% < 0 Or TI% > 10 Then TI% = 2
             TIX$ = TIS$(TI%): SIG% = 1
             If TI% > 3 Then SIG% = (-1)
             '
             NF& = CVS(Mid$(X$, 5, 4))
             While NF& > 999999
                NF& = NF& - 100000
             Wend
             NFS$ = TRIM$(Str$(NF&))
             While Len(NFS$) < 8
                NFS$ = " " + NFS$
             Wend
             NFS$ = TIX$ + NFS$
             '
15           FF% = CVI(Mid$(X$, 19, 2))
             FVI% = CVI(Mid$(X$, 9, 2))
             SALDEU# = SIG% * Abs(CVD(Mid$(X$, 11, 8)))
             If TI% = 0 Or TI% = 7 Then
                If Sgn(CVD(Mid$(X$, 21, 8))) < 0 Then
                   SALDEU# = (-1) * Abs(CVD(Mid$(X$, 11, 8)))
                End If
             End If
             If Abs(SALDEU#) < 0.005 Then GoTo 99
             '
             If FF% = 0 Then
                Screen.MousePointer = 1
                Call MENSERR(24, "Imposible Continuar Proceso de Cálculo.\Fecha de Factura no Válida en Cuenta '" + TRIM$(Str$(NC%)) + "'.\Revise la Composición de Deuda de Esta Cuenta.")
                Hide
                Exit Sub
             End If
             '
             IMORIG# = SIG% * CVD(Mid$(X$, 21, 8))
             ES$ = "    "
             If SALDEU# < 0 Then
                   ES$ = " CR "
                 Else
                   'If FVI% <= HOO% + 8 Then ES$ = " ** "
                   'If FVI% <= HOO% Then ES$ = "!!!!"
                   If FVI% < HOO% Then
                     FVIL& = FVI%: HOOL& = HOO%
                     ES$ = FORMATNUM$(DIENTRE%(FVIL&, HOOL&), "F4.0")
                     If Left$(ES$, 1) = " " Then Mid$(ES$, 1, 1) = "."
                   End If
             End If
             '
             Z$ = REGBLAN$("FAPECOB")
             Call REPLA(Z$, MKI$(NC%), 1, 2)
             Call REPLA(Z$, NFS$, 3, 10)
             Call REPLA(Z$, MKI$(FF%), 13, 2)
             Call REPLA(Z$, MKI$(FVI%), 15, 2)
             Call REPLA(Z$, MKD$(SALDEU#), 17, 8)
             Call REPLA(Z$, MKD$(IMORIG#), 25, 8)
             '
             UJ& = UJ& + 1
             Call GRAREG("FAPECOB", Z$, UJ&)
             SUDEUCLI#(NC%) = SUDEUCLI#(NC%) + SALDEU#
             SUDEUTO# = SUDEUTO# + SALDEU#
             '
           End If
         End If
       End If
       '
99   Next UI&
     '
     Call SETULTREG("FAPECOB", UJ&)
     Call CIERRARCH("FAPECOB")
     '
     Screen.MousePointer = 1
     '
    ' LEER FACTURAS - CARGAR VECTOR CON MOVIMIENTOS POSTERIORES A FLIM%
    CAMOPO% = 0
    FIN& = ULTREG&("FACTUR")
    For UI& = 1 To FIN&
      X$ = REGLEIDO$("FACTUR", UI&)
      TIPO% = Asc(Mid$(X$, 13, 1))
      If TIPO% > 0 Then
        If TIPO% < 4 Then
          FFAC% = CVI(Mid$(X$, 7, 2))
          If FFAC% > FLIM% Then
            CAMOPO% = CAMOPO% + 1
            If CAMOPO% > 8196 Then
              Call MENSERR(24, "Imposible Re-Calcular Composición de Deuda\en la Fecha Elegida.")
              Call FINAL("")
            End If
            NROPO&(CAMOPO%) = CVS(Mid$(X$, 3, 4))
            CLIPO%(CAMOPO%) = CVI(Left$(X$, 2))
          End If
        End If
      End If
    Next UI&
    '
    ' LEER APLICOB - DESCARTAR MOVIMIENTOS ANTERIORES A FLIM%
    '              - DESCARTAR MOVIMIENTOS QUE FIGURAN EN MOVIMIENTOS POSTERIORES
    '
    FIN& = ULTREG&("APLICOB")
    For UI& = 1 To FIN&
      X$ = REGLEIDO$("APLICOB", UI&)
      FAPLI% = CVI(Left$(X$, 2))
      If FAPLI% > FLIM% Then
        CLIAPLI% = CVI(Mid$(X$, 3, 2))
        If CLIAPLI% > 0 Then
          If Abs(SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)) >= 0.005 Then
            DOCUORI$ = Mid$(X$, 55, 18)
            If Mid$(DOCUORI$, 3, 1) <> "-" Or Mid$(DOCUORI$, 8, 1) <> "-" Or Mid$(DOCUORI$, 17, 1) <> "-" Then
   AAA = bbb
            End If
            NUDOAP& = XVALO(Mid$(DOCUORI$, 9, 8))
            For KKI% = 1 To CAMOPO%
              If NROPO&(KKI%) = NUDOAP& Then
                If CLIPO%(KKI%) = CLIAPLI% Then GoTo 29
              End If
            Next KKI%
            '
            Z$ = REGBLAN$("FAPECOB")
            Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
             NF& = XVALO(Mid$(DOCUORI$, 10, 7))
             While NF& > 999999: NF& = NF& - 100000: Wend
             NFS$ = TRIM$(Str$(NF&))
             While Len(NFS$) < 8: NFS$ = " " + NFS$: Wend
             NFS$ = Left$(DOCUORI$, 2) + NFS$
            Call REPLA(Z$, NFS$, 3, 10)
            TDCC$ = UCase$(Left$(NFS$, 2))
            TIXA% = 2: If TDCC$ = "ND" Then TIXA% = 3
                       If TDCC$ = "NC" Then TIXA% = 6
                       If TDCC$ = "AJ" Then TIXA% = 0
            FF% = FLIM%
            RFFA$ = RECFILT$("FACTUR", 1, MKI$(CLIAPLI%))
            For KKLL& = 1 To Len(RFFA$) Step 4
              RELEL& = CVS(Mid$(RFFA$, KKLL&, 4))
              ZFF$ = REGLEIDO$("FACTUR", RELEL&)
              If CVS(Mid$(ZFF$, 3, 4)) = NF& Then
                If Asc(Mid$(ZFF$, 13, 1)) = TIXA% Then
                  FF% = CVI(Mid$(ZFF$, 7, 2))
                  GoTo 24
                End If
              End If
            Next KKLL&
            '
24          Call REPLA(Z$, MKI$(FF%), 13, 2)
            FVI% = FAPLI%
            Call REPLA(Z$, MKI$(FVI%), 15, 2)
            SALDEU# = CVD(Mid$(X$, 23, 8))
            Call REPLA(Z$, MKD$(SALDEU#), 17, 8)
            IMORIG# = CVD(Mid$(X$, 73, 8))
            Call REPLA(Z$, MKD$(IMORIG#), 25, 8)
            '
            UJ& = UJ& + 1
            Call GRAREG("FAPECOB", Z$, UJ&)
            SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDEU#
            SUDEUTO# = SUDEUTO# + SALDEU#
            '
          End If
        End If
      End If
29  Next UI&
    '
    FIN& = ULTREG&("APLICOB")
    For UI& = 1 To FIN&
      X$ = REGLEIDO$("APLICOB", UI&)
      FAPLI% = CVI(Left$(X$, 2))
      If FAPLI% <= FLIM% Then
        CLIAPLI% = CVI(Mid$(X$, 3, 2))
        DIFESAL# = SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)
        If Abs(DIFESAL#) >= 0.005 Then
          DOCUORI$ = Mid$(X$, 55, 18)
          If Mid$(DOCUORI$, 3, 1) <> "-" Or Mid$(DOCUORI$, 8, 1) <> "-" Or Mid$(DOCUORI$, 17, 1) <> "-" Then
AAA = bbb
          End If
          NUDOAP& = XVALO(Mid$(DOCUORI$, 9, 8))
          DOCUAP$ = Mid$(X$, 5, 3) + Mid$(X$, 16, 7)
          'For KKI% = 1 To CAMOPO%
          '  If NROPO&(KKI%) = NUDOAP& Then
          '    If CLIPO%(KKI%) = CLIAPLI% Then GoTo 39
          '  End If
          'Next KKI%
          '
          Z$ = REGBLAN$("FAPECOB")
          Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
             NF& = XVALO(Mid$(DOCUAP$, 4, 7))
             While NF& > 999999: NF& = NF& - 100000: Wend
             NFS$ = TRIM$(Str$(NF&))
             While Len(NFS$) < 8: NFS$ = " " + NFS$: Wend
             NFS$ = Left$(DOCUAP$, 2) + NFS$
          Call REPLA(Z$, NFS$, 3, 10)
          FF% = FLIM%
          Call REPLA(Z$, MKI$(FF%), 13, 2)
          FVI% = FAPLI%
          Call REPLA(Z$, MKI$(FVI%), 15, 2)
          SALDEU# = CVD(Mid$(X$, 23, 8))
          Call REPLA(Z$, MKD$((-1) * SALDEU#), 17, 8)
          IMORIG# = CVD(Mid$(X$, 73, 8))
          Call REPLA(Z$, MKD$(IMORIG#), 25, 8)
          '
          UJ& = UJ& + 1
          Call GRAREG("FAPECOB", Z$, UJ&)
          SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) - SALDEU#
          SUDEUTO# = SUDEUTO# - SALDEU#
          '
        End If
      End If
39  Next UI&
    '
    FIN& = 32767
    For UI& = 1 To FIN&
      CLIAPLI% = UI&
      DIFESAL# = SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)
      If Abs(DIFESAL#) >= 0.005 Then
          '
          Z$ = REGBLAN$("FAPECOB")
          Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
          NFS$ = "Dif.Saldo"
          Call REPLA(Z$, NFS$, 3, 10)
          FF% = FLIM%
          Call REPLA(Z$, MKI$(FF%), 13, 2)
          FVI% = FLIM%
          Call REPLA(Z$, MKI$(FVI%), 15, 2)
          SALDEU# = DIFESAL#
          Call REPLA(Z$, MKD$(SALDEU#), 17, 8)
          IMORIG# = CVD(Mid$(X$, 73, 8))
          Call REPLA(Z$, MKD$(DIFESAL#), 25, 8)
          '
          UJ& = UJ& + 1
          Call GRAREG("FAPECOB", Z$, UJ&)
          SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDEU#
          '
      End If
49  Next UI&
    '
    Call SETULTREG("FAPECOB", UJ&)
    Call CIERRARCH("FAPECOB")
    '              - AGREGAR A LISTA DE SVD202
    ' ORDENAR POR NUMERO O ALFABETICO Y POR FECHA EMISION
    '
    GRUCLI% = XVALO(Text1.TEXT)
    GRUX$ = TRIM$(Text2.TEXT)
    VENDE% = XVALO(Text3.TEXT)
    VENDX$ = TRIM$(Text4.TEXT)
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    '
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    If Option1(3) = True Then MODORDE% = 3
    '
    FIN& = ULTREG&("FAPECOB")
    UJ& = 0
    List1.Clear
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPECOB", U&)
        NC% = CVI(Left$(X$, 2))
        '
        If NC% >= NCMIN% Then
          If NC% <= NCMAX% Then
            FVI% = CVI(Mid$(X$, 15, 2))
            If GRUCLI% = 0 Or GRUCOCLI%(NC%) = GRUCLI% Then
              If VENDE% = 0 Or VENDCLI%(NC%) = VENDE% Then
                '
                TLISTA$ = Space$(36)
                If MODORDE% = 1 Then
                     Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
                  ElseIf MODORDE% = 2 Then
                     Call REPLA(TLISTA$, RASOCLI$(NC%), 1, 30)
                  ElseIf MODORDE% = 3 Then
                     Call REPLA(TLISTA$, CSTRING$(MKI$(FVI%), 1, 6, 0, 2), 1, 6)
                End If
                Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
                List1.AddItem TLISTA$
                '
              End If
            End If
          End If
        End If
        '
    Next U&
    '
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        List1.ListIndex = U& - 1
        REG& = XVALO(Mid$(List1.TEXT, 31, 8))
        X$ = REGLEIDO$("FAPECOB", REG&)
        '
        NC% = CVI(Left$(X$, 2))
        NFS$ = Mid$(X$, 3, 10)
        FF% = CVI(Mid$(X$, 13, 2))
        FVI% = CVI(Mid$(X$, 15, 2))
        SALDEU# = CVD(Mid$(X$, 17, 8))
        IMORIG# = CVD(Mid$(X$, 25, 8))
        ES$ = ""   ': If FVI% < HOY(HOS$) Then ES$ = "!!!!"
        If FVI% < HOY(HOS$) Then
           FVIL& = FVI%: HOOL& = HOY(HOS$)
           ES$ = FORMATNUM$(DIENTRE%(FVIL&, HOOL&), "F4.0")
           If Left$(ES$, 1) = " " Then Mid$(ES$, 1, 1) = "."
        End If        '
        Y$ = REGBLAN$("FAPECOBA")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        Call REPLA(Y$, RASOCLI$(NC%), 3, 30)
        Call REPLA(Y$, TEL1CLI$(NC%), 33, 16)
        Call REPLA(Y$, NFS$, 49, 12)
        Call REPLA(Y$, MKI$(FF%), 61, 2)
        Call REPLA(Y$, MKD$(IMORIG#), 63, 8)
        Call REPLA(Y$, MKI$(FVI%), 71, 2)
        Call REPLA(Y$, MKD$(SALDEU#), 73, 8)
        Call REPLA(Y$, ES$, 97, 4)
        '
        UJ& = UJ& + 1
        Call GRAREG("FAPECOBA", Y$, UJ&)
        '          '
    Next U&
    '
    Call SETULTREG("FAPECOBA", UJ&)
    Call CIERRARCH("FAPECOBA")
    '
    FEX = FLIM%
    Call HEADERS("FAPECOBA", "")
    Call HEADERS("FAPECOBA", "Composición de Deuda al: " + FECHATEX$(FEX))
    '
    Call FINAL("*LIFAPEA1")
    '
    Screen.MousePointer = 1

End Sub

'Sub FAPECOB1()
'    '
'Exit Sub
'10  Screen.MousePointer = 11
'    OKEY% = 0
'    '
'    NCMIN% = 32767: NCMAX% = 0
'    FLIM = HOY(HO$)
'    FIN& = ULTREG&("SVD202")
'    '
'    For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("SVD202", U&)
'      NCX% = CVI(Left$(X$, 2))
'      If NCX% > 0 Then
'        If NCX% > NCMAX% Then NCMAX% = NCX%
'        If NCX% < NCMIN% Then NCMIN% = NCX%
'        '
'        FEVE = CVI(Mid$(X$, 9, 2))
'        If FEVE > FLIM Then FLIM = FEVE
'      End If
'    Next U&
'    '
'    Text10.TEXT = TRIM$(Str$(NCMIN%))
'    Text9.TEXT = RASOCLI$(NCMIN%)
'    Text8.TEXT = TRIM$(Str$(NCMAX%))
'    Text7.TEXT = RASOCLI$(NCMAX%)
'    Text6.TEXT = FECHATEX$(FLIM)
'    Text5.TEXT = "01/01/80"
'    '
'    Dim TIS$(10)
'    TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
'    For I% = 0 To 7
'        TIS$(I%) = Mid$(TDAT$, 3 * I% + 1, 2)
'    Next I%
'    '
'    UJ& = 0
'    For UI& = 1 To FIN&
'       Call TRACE(20, UI&, FIN&)
'       X$ = REGLEIDO$("SVD202", UI&)
'       NC% = CVI(Left$(X$, 2))
'       If NC% > 0 Then
'           TI% = Asc(Mid$(X$, 3, 1))
'           If TI% < 0 Or TI% > 10 Then TI% = 2
'           TIX$ = TIS$(TI%): SIG% = 1
'           If TI% > 3 Then SIG% = (-1)
'           '
'           NF = CVS(Mid$(X$, 5, 4))
'           While NF > 999999
'              NF = NF - 100000
'           Wend
'           NFS$ = TRIM$(Str$(NF))
'           While Len(NFS$) < 8
'              NFS$ = " " + NFS$
'           Wend
'           NFS$ = TIX$ + NFS$
'           '
'15         FF% = CVI(Mid$(X$, 19, 2))
'           FVI% = CVI(Mid$(X$, 9, 2))
'           SALDEU# = SIG% * Abs(CVD(Mid$(X$, 11, 8)))
'           If TI% = 0 Or TI% = 7 Then
'              If Sgn(CVD(Mid$(X$, 21, 8))) < 0 Then
'                 SALDEU# = (-1) * Abs(CVD(Mid$(X$, 11, 8)))
'              End If
'           End If
'           If Abs(SALDEU#) < 0.005 Then GoTo 99
'           '
'           If FF% = 0 Then
'              Screen.MousePointer = 1
'              Call MENSERR(24, "Imposible Continuar Proceso de Cálculo.\Fecha de Factura no Válida en cuenta '" + TRIM$(Str$(NC%)) + "'.\Revise la Composición de Deuda de Esta Cuenta.")
'              Hide
'              Exit Sub
'           End If
'           '
'           IMORIG# = SIG% * CVD(Mid$(X$, 21, 8))
'           ES$ = "    "
'           MONEMX% = Asc(Mid$(X$, 31, 1))
'           '
'           If SALDEU# < 0 Then
'                 ES$ = " CR "
'               Else
'                 'If FVI% <= HOO% + 8 Then ES$ = " ** "
'                 'If FVI% <= HOO% Then ES$ = "!!!!"
'                 If FVI% < HOO% Then
'                   FVIL& = FVI%: HOOL& = HOO%
'                   ES$ = FORMATNUM$(DIENTRE%(FVIL&, HOOL&), "F4.0")
'                   If Left$(ES$, 1) = " " Then Mid$(ES$, 1, 1) = "."
'                 End If
'           End If
'           '
'           Z$ = REGBLAN$("FAPECOB")
'           Call REPLA(Z$, MKI$(NC%), 1, 2)
'           Call REPLA(Z$, NFS$, 3, 10)
'           Call REPLA(Z$, MKI$(FF%), 13, 2)
'           Call REPLA(Z$, MKI$(FVI%), 15, 2)
'           Call REPLA(Z$, MKD$(SALDEU#), 17, 8)
'           Call REPLA(Z$, MKD$(IMORIG#), 25, 8)
'           Call REPLA(Z$, Chr$(MONEMX%), 33, 1)
'           Call REPLA(Z$, Chr$(TI%), 34, 1)
'           '
'           UJ& = UJ& + 1
'           Call GRAREG("FAPECOB", Z$, UJ&)
'           '
'       End If
'       '
'99   Next UI&
'     '
'     Call SETULTREG("FAPECOB", UJ&)
'     Screen.MousePointer = 1
'     '
'End Sub
'

'Sub FAPECOB2()
'    '
'    Screen.MousePointer = 11
'    GRUCLI% = XVALO(Text1.TEXT)
'    GRUX$ = TRIM$(Text2.TEXT)
'    VENDE% = XVALO(Text3.TEXT)
'    VENDX$ = TRIM$(Text4.TEXT)
'              If GRUCLI% = 0 Or GRUCOCLI%(NC%) = GRUCLI% Then
'                If VENDE% = 0 Or VENDCLI%(NC%) = VENDE% Then
'    FINF = FECHANUM(Text5.TEXT)
'    FLIM = FECHANUM(TRIM$(Text6.TEXT)) ': If FLIM < HOY(HOS$) Then FLIM = HOY(HOS$)
'    NCMIN% = XVALO(Text10.TEXT)
'    NCMAX% = XVALO(Text8.TEXT)
'    '
'    MODORDE% = 1
'    If Option1(2) = True Then MODORDE% = 2
'    If Option1(3) = True Then MODORDE% = 3
'    '
'    FIN& = ULTREG&("FAPECOB")
'    UJ& = 0
'    List1.Clear
'    Dim SALDEPE#(32767)
'    '
'    For U& = 1 To FIN&
'        '
'        Call TRACE(20, U&, FIN&)
'        X$ = REGLEIDO$("FAPECOB", U&)
'        NC% = CVI(Left$(X$, 2))
'        SALDEU# = CVD(Mid$(X$, 17, 8))
'        If NC% >= 0 Then
'          SALDEPE#(NC%) = SALDEPE#(NC%) + SALDEU#
'        End If
'        '
'        If NC% >= NCMIN% Then
'          If NC% <= NCMAX% Then
'            FVI% = CVI(Mid$(X$, 15, 2))
'            If FVI% >= FINF And FVI% <= FLIM Then
'              If GRUCLI% = 0 Or GRUCOCLI%(NC%) = GRUCLI% Then
'                If VENDE% = 0 Or VENDCLI%(NC%) = VENDE% Then
'                  '
'                  TLISTA$ = Space$(36)
'                  If MODORDE% = 1 Then
'                       Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
'                    ElseIf MODORDE% = 2 Then
'                       Call REPLA(TLISTA$, RASOCLI$(NC%), 1, 30)
'                    ElseIf MODORDE% = 3 Then
'                       Call REPLA(TLISTA$, CSTRING$(MKI$(FVI%), 1, 6, 0, 2), 1, 6)
'                  End If
'                  Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
'                  List1.AddItem TLISTA$
'                  '
'                End If
'              End If
'            End If
'          End If
'        End If
'        '
'    Next U&
'    '
'    EPAC$ = TRIM$(UCase$(EMPREAC$))
'    If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "SABO") > 0 Then
'      ATXX% = 0
'      FIN& = ULTREG&("DEUDOR1")
'      For UI& = 1 To FIN&
'        Call TRACE(20, UI&, FIN&)
'        XD1$ = REGLEIDO$("DEUDOR1", UI&)
'        NCI% = CVI(Left$(XD1$, 2))
'        If NCI% > 0 Then
'          If Abs(SALDEPE#(NCI%) - SALDCLI#(NCI%)) >= 0.005 Then
'' AKX1$ = RASOCLI$(NCI%)
'' ABC = SALDCLI#(NCI%)
'            RESCUDEU.Text1 = TRIM$(Str$(NCI%))
'            RESCUDEU.Show 1
'            ATXX% = 1
'          End If
'        End If
'      Next UI&
'      If ATXX% > 0 Then
'        '
'        Screen.MousePointer = 11
'        On Error Resume Next
'        HOII% = HOY(HOS$)
'        DRESPA$ = "C:\FLEXOFT\RP" + Left$(HOS$, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 7, 2)
'        On Error Resume Next
'        MkDir DRESPA$
'        ABAK$ = "SVD202": FRESPA$ = LOADFILE$(LUDAT$ + ABAK$ + ".DAT"): Call SAVEFILE(DRESPA$ + "\" + ABAK$ + ".DAT", FRESPA$)
'        ABAK$ = "FAPECOB": FRESPA$ = LOADFILE$(LUDAT$ + ABAK$ + ".DAT"): Call SAVEFILE(DRESPA$ + "\" + ABAK$ + ".DAT", FRESPA$)
'        ABAK$ = "FACTUR": FRESPA$ = LOADFILE$(LUDAT$ + ABAK$ + ".DAT"): Call SAVEFILE(DRESPA$ + "\" + ABAK$ + ".DAT", FRESPA$)
'        ABAK$ = "COBRAN": FRESPA$ = LOADFILE$(LUDAT$ + ABAK$ + ".DAT"): Call SAVEFILE(DRESPA$ + "\" + ABAK$ + ".DAT", FRESPA$)
'        ABAK$ = "DEUDOR1": FRESPA$ = LOADFILE$(LUDAT$ + ABAK$ + ".DAT"): Call SAVEFILE(DRESPA$ + "\" + ABAK$ + ".DAT", FRESPA$)
'        ABAK$ = "DEUDOR2": FRESPA$ = LOADFILE$(LUDAT$ + ABAK$ + ".DAT"): Call SAVEFILE(DRESPA$ + "\" + ABAK$ + ".DAT", FRESPA$)
'        On Error GoTo 0
'        Screen.MousePointer = 1
'        '
'        Call MENSERR(24, "Imposible Emitir Listado de Facturas\Pendientes de Cobranza en Contexto Actual.\  \Repita su Petición")
'        Call FINAL("")
'      End If
'    End If
'    '
'    FIN& = List1.ListCount
'    '
'    UNIFICA% = 0
'    For U& = 1 To FIN&
'        '
'        Call TRACE(20, U&, FIN&)
'        List1.ListIndex = U& - 1
'        REG& = XVALO(Mid$(List1.TEXT, 31, 8))
'        X$ = REGLEIDO$("FAPECOB", REG&)
'        '
'        NC% = CVI(Left$(X$, 2))
'        NFS$ = Mid$(X$, 3, 10)
'        NROL& = XVALO(Mid$(NFS$, 4, 7))
'        FF% = CVI(Mid$(X$, 13, 2))
'        FVI% = CVI(Mid$(X$, 15, 2))
'        SALDEU# = CVD(Mid$(X$, 17, 8))
'        IMORIG# = CVD(Mid$(X$, 25, 8))
'        ES$ = ""   ': If FVI% < HOY(HOS$) Then ES$ = "!!!!"
'        If FVI% < HOY(HOS$) Then
'           FVIL& = FVI%: HOOL& = HOY(HOS$)
'           ES$ = FORMATNUM$(DIENTRE%(FVIL&, HOOL&), "F4.0")
'           If Left$(ES$, 1) = " " Then Mid$(ES$, 1, 1) = "."
'        End If
'        MONEMX% = Asc(Mid$(X$, 33, 1))
'        TIA% = Asc(Mid$(X$, 34, 1))
'        '
'        Y$ = REGBLAN$("FAPECOBD")
'        Z$ = REGBLAN$("FAPECOBR")
'        '
'        Call REPLA(Y$, MKI$(NC%), 1, 2)
'        Call REPLA(Y$, RASOCLI$(NC%), 3, 30)
'        Call REPLA(Y$, TEL1CLI$(NC%), 33, 16)
'        Call REPLA(Y$, NFS$, 49, 10)
'        Call REPLA(Y$, MKI$(FF%), 59, 2)
'        Call REPLA(Y$, MKD$(IMORIG#), 65, 8)
'        Call REPLA(Y$, MKI$(FVI%), 73, 2)
'        Call REPLA(Y$, MKD$(SALDEU#), 75, 8)
'        '
'        Call REPLA(Z$, Left$(Y$, 60), 1, 60)
'        Call REPLA(Z$, MKD$(IMORIG#), 61, 8)
'        Call REPLA(Z$, MKI$(FVI%), 69, 2)
'        Call REPLA(Z$, MKD$(SALDEU#), 71, 8)
'        '
'        If MONEMX% = 2 Then
'          '
'          If UNIFICA% = 0 Then
'            UNIFICA% = (-1)
'            If COMALTER%("Puede Mantener Separados\los Saldos en Pesos y Dolares\o Unificarlos en Pesos al T.C. " + TRIM$(FORMATNUM$(TICAMONE(MONEMX%), "F10.4")) + "\\Mantener Separados\Unificar en Pesos") = 2 Then
'              UNIFICA% = 1
'            End If
'          End If
'          '
'          If Abs(Abs(IMORIG#) - Abs(SALDEU#)) < 0.005 Then
'            RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
'            For KUL& = Len(RFG$) - 3 To 1 Step -4
'              RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
'              XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
'              If CVS(Mid$(XXC$, 39, 4)) = NROL& Then
'                If Asc(Mid$(XXC$, 37, 1)) = TIA% Then
'                  SALDOL# = CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8))
'                  Call REPLA(Y$, MKD$(0), 75, 8)
'                  If UNIFICA% = 1 Then Call REPLA(Y$, MKD$(SALDOL# * TICAMONE(MONEMX%)), 75, 8)
'                  Call REPLA(Y$, MKD$(SALDOL#), 91, 8)
'                  GoTo 20
'                End If
'              End If
'            Next KUL&
'          End If
'        End If
'        '
'20      If MONEMX% = 2 Then
'           Call REPLA(Y$, "U$S", 61, 4)
'        End If
'        Call REPLA(Y$, ES$, 107, 10)
'        Call REPLA(Z$, ES$, 87, 10)
'        '
'        UJ& = UJ& + 1
'        Call GRAREG("FAPECOBR", Z$, UJ&)
'        Call GRAREG("FAPECOBD", Y$, UJ&)
'        '          '
'    Next U&
'    '
'    Call SETULTREG("FAPECOBR", UJ&)
'    Call CIERRARCH("FAPECOBR")
'    Call SETULTREG("FAPECOBD", UJ&)
'    Call CIERRARCH("FAPECOBD")
'    '
'    Call HEADERS("FAPECOBR", "")
'    If GRUCLI% > 0 Then Call HEADERS("FAPECOBR", "Grupo: " + GRUX$)
'    If VENDE% > 0 Then Call HEADERS("FAPECOBR", "Vendedor: " + VENDX$)
'    Call HEADERS("FAPECOBR", "Vencimiento: " + FECHATEX$(FINF) + " - " + FECHATEX$(FLIM))
'    Call HEADERS("FAPECOBR", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
'    '
'    Call HEADERS("FAPECOBD", "")
'    If GRUCLI% > 0 Then Call HEADERS("FAPECOBD", "Grupo: " + GRUX$)
'    If VENDE% > 0 Then Call HEADERS("FAPECOBD", "Vendedor: " + VENDX$)
'    Call HEADERS("FAPECOBD", "Vencimiento: " + FECHATEX$(FINF) + " - " + FECHATEX$(FLIM) + "   -   R.Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
'    'Call HEADERS("FAPECOBD", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
'    If UNIFICA% = 1 Then Call HEADERS("FAPECOBD", "Observaciones: Saldos en Dolares Unificados a Pesos al T.C. " + TRIM$(FORMATNUM$(TICAMONE(2), "F10.4")))
'    '
'    Screen.MousePointer = 1
'    '
'End Sub
''

Sub FAPECOB1()
    '
    Screen.MousePointer = 11
    '
    NCMIN% = 32767: NCMAX% = 0
    For U& = 1 To 32767: TICUE%(U&) = 1: Next U&
    FIN& = ULTREG&("DEUDOR1")
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      NCX% = CVI(Left$(REGLEIDO$("DEUDOR1", U&), 2))
      If NCX% > 0 Then
        If NCX% > NCMAX% Then NCMAX% = NCX%
        If NCX% < NCMIN% Then NCMIN% = NCX%
        TICUE%(NCX%) = 0
      End If
    Next U&
    Text10 = TRIM$(Str$(NCMIN%))
    '\\\OT-05-0302-WAR-19/05/05///
    Text9 = RASOCLI$(NCMIN%)
    Text8 = TRIM$(Str$(NCMAX%))
    Text7 = RASOCLI$(NCMAX%)
    '\\\OT-05-0302-WAR-FIN///
    '
    Text1 = ""
    Text3 = ""
    '
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    Set SACTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    FLIM = HOY(HOS$)
    With SACTEMP
      .MoveFirst
      If Err < 1 Then
        FLIM = CVINT(!FEVENCIM)
      End If
    End With
    On Error GoTo 0
    'RECHA MAX
    Text6 = FECHATEX$(FLIM)
    Text6.Refresh
    'FECHA MIN
    Text5.TEXT = "01/01/80"
    Text5.Refresh
    '
    CARFAPEN% = 1
    Screen.MousePointer = 1
    '
End Sub
'
Sub FAPECOB2()
    '
    Screen.MousePointer = 11
    '
    HOII% = HOY(HOS$)
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    '
    FDESD = CDbl(CVDAT(Int(FECHANUM(Text5.TEXT))))
    FLIM = CDbl(CVDAT(Int(FECHANUM(Text6.TEXT))))
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    DESEMI = CDbl(CVDAT(Int(FECHANUM(Label9))))
    HASEMI = CDbl(CVDAT(Int(FECHANUM(Label8))))

    '
    GRUCLI% = XVALO(Text1.TEXT)
    GRUX$ = TRIM$(Text2.TEXT)
    VENDE% = XVALO(Text3.TEXT)
    VENDX$ = TRIM$(Text4.TEXT)
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    If Option1(3) = True Then MODORDE% = 3
    '\\\OT-08-0592-OD-07/08/08///
    If Option1(11) = True Then MODORDE% = 4
    '\\\OT-08-0592-OD-FIN///
    '
    On Error Resume Next
    Kill LUDAT$ + "FAPECOBA.DAT"
    On Error GoTo 0
    On Error Resume Next
    Kill LUDAT$ + "FAPECOBD.DAT"
    On Error GoTo 0
    On Error Resume Next
    Kill LUDAT$ + "FAPECOBR.DAT"
    On Error GoTo 0
    '
    BIMONET% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "BIMONETA" Then
         BIMONET% = 1
       ElseIf MODOMONE$ = "MULTIMON" Then
         BIMONET% = 2
    End If
    '
    REBLX$ = REGBLAN$("FAPECOBA")
    REBLY$ = REGBLAN$("FAPECOBD")
    REBLZ$ = REGBLAN$("FAPECOBR")
    '
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    SQLX$ = SQLX$ + "AND NuClien >= " & NCMIN% & " "
    SQLX$ = SQLX$ + "AND NuClien <= " & NCMAX% & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FeVencim))<= " & FLIM & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FeVencim))>= " & FDESD & " "
    If TRIM$(Label8) <> "" Then
        SQLX$ = SQLX$ + "AND CDBL(FechEmisi)>= " & DESEMI & " "
        SQLX$ = SQLX$ + "AND CDBL(FechEmisi)<= " & HASEMI & " "
    End If
    If TRIM$(Text12) <> "" Then
        CondiPa% = XVALO(Text12) 'Condicion de Pago
        SQLX$ = SQLX$ + "AND CondPag = " & CondiPa% & " "
    End If

    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    Set SACTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    On Error Resume Next
    With SACTEMP
      .MoveLast
      FIN& = .RecordCount
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
        UJ& = 0: UK& = 0: UL& = 0
        Do While Not .EOF
          UL& = UL& + 1
          Call TRACE(20, UL&, FIN&)
          NC% = !NuClien
          NFS$ = !CodiCom_Cor
          COCOLAR$ = !CODICOM_LAR
          FF% = CVINT(!FechEmisi)
          FVI% = CVINT(!FEVENCIM)
          SALDEU# = !SaldDebe - !SaldAcre
          SALDOL# = 0
          On Error Resume Next
          SALDOL# = !SaldDebe_Dola - !SaldAcre_Dola
          On Error GoTo 0
          IMORIG# = !IBruComp
          ES$ = "    "
          MONEMX% = 1
          On Error Resume Next
          MONEMX% = !Mone_Emis
          On Error GoTo 0
          '
          If MODOMONEI% = 2 Then
             Call APERBASDAT("CABAN")
             ' AGREGUE LAS CONDICIONES DE TipoMovim =
             ' PORQUE SI NO A VECES TRAE EL NETO O EL IVA - EPB 24/05/07
             CajaBanc.FindFirst "CodiCom_Lar = '" & COCOLAR$ & "' AND (TipoMovim = 'FVEN' OR TipoMovim = 'RCOB' OR TipoMovim = 'AJUD')"
             If CajaBanc.NoMatch = False Then
                On Error Resume Next
                IMORIG# = CajaBanc!Impo_Dola
                If Abs(IMORIG#) < 0.005 Or Sgn(IMORIG#) <> Sgn(SALDEU#) Then
                  IPESO# = ROUND(CajaBanc!IDEBECOMP - CajaBanc!IHABECOMP)
                  MONEMX% = CajaBanc!Mone_Emis
                  TICAM = CajaBanc!Tipo_Cam: If TICAM < 0.005 Then TICAM = 1
                  IMORIG# = ROUND#(IPESO# / TICAM)
                End If
                On Error GoTo 0
             End If
          End If
          '
          If GRUCLI% = 0 Or GRUCOCLI%(NC%) = GRUCLI% Then
            If VENDE% = 0 Or VENDCLI%(NC%) = VENDE% Then
              If Abs(SALDEU#) >= 0.005 Or Abs(SALDOL#) >= 0.005 Then
                Y$ = REBLX$
                Call REPLA(Y$, MKI$(NC%), 1, 2)
                Call REPLA(Y$, RASOCLI$(NC%), 3, 30)
                Call REPLA(Y$, TEL1CLI$(NC%), 33, 16)
                Call REPLA(Y$, NFS$, 49, 12)
                Call REPLA(Y$, MKI$(FF%), 61, 2)
                Call REPLA(Y$, MKD$(IMORIG#), 63, 8)
                Call REPLA(Y$, MKI$(FVI%), 71, 2)
                Call REPLA(Y$, MKD$(SALDEU#), 73, 8)
                Call REPLA(Y$, ES$, 97, 4)
                '
                GRUVE$ = ""
                If Check1.Value = 1 Then
                     GRUVE$ = "Grupo: " + ECOARCH$("GRUCOCLI", MKI$(GRUCOCLI%(NC%)))
                  ElseIf Check2.Value = 1 Then
                     GRUVE$ = "Vendedor: " + ECOARCH$("VENDEDOR", Chr$(VENDCLI%(NC%)))
                End If
                Call REPLA(Y$, GRUVE$, 101, 28)
                '
                UJ& = UJ& + 1
                Call GRAREG("FAPECOBA", Y$, UJ&)
                '
                If BIMONET% < 1 Then GoTo 29
                '
                Y$ = REBLY$
                Z$ = REBLZ$
                '
                Call REPLA(Y$, MKI$(NC%), 1, 2)
                Call REPLA(Y$, RASOCLI$(NC%), 3, 30)
                Call REPLA(Y$, TEL1CLI$(NC%), 33, 16)
                Call REPLA(Y$, NFS$, 49, 10)
                Call REPLA(Y$, MKI$(FF%), 59, 2)
20              If MONEMX% > 1 Then
                   Call REPLA(Y$, SIMBOLPES$(MONEMX%), 61, 4)
                End If
                Call REPLA(Y$, MKD$(IMORIG#), 65, 8)
                Call REPLA(Y$, MKI$(FVI%), 73, 2)
                Call REPLA(Y$, MKD$(SALDEU#), 75, 8)
                Call REPLA(Y$, MKD$(SALDOL#), 91, 8)
                ES$ = ""
                If FVI% < HOII% Then ES$ = "Vencida"
                If Abs(SALDEU#) >= 0.005 Then
                  If Abs(SALDOL#) < 0.005 Then
                    If MONEMX% > 1 Then
                      ES$ = "Ajustar"
                    End If
                  End If
                End If
                '
                '\\\OT-08-0592-OD-07/08/08///
                CPAG% = CPAGCLI%(NC%)
                CONDPAG$ = ECOARCH$("CONPAG", Chr$(CPAG%))
                If Option1(11).Value = True Then
                  ES$ = CONDPAG$
                End If
                '\\\OT-08-0592-OD-FIN///
                '
                Call REPLA(Y$, ES$, 107, 20)
                '
                Call REPLA(Z$, Left$(Y$, 60), 1, 60)
                Call REPLA(Z$, MKD$(IMORIG#), 61, 8)
                Call REPLA(Z$, MKI$(FVI%), 69, 2)
                Call REPLA(Z$, MKD$(SALDEU#), 71, 8)
                Call REPLA(Z$, ES$, 87, 20)
                '
                UK& = UK& + 1
                Call GRAREG("FAPECOBR", Z$, UK&)
                Call GRAREG("FAPECOBD", Y$, UK&)
                '          '
29            End If
            End If
          End If
          '
        .MoveNext
        Loop
        '
     ' End If
    End With
    '
    Call SETULTREG("FAPECOBA", UJ&)
    Call CIERRARCH("FAPECOBA")
    Call SETULTREG("FAPECOBR", UK&)
    Call CIERRARCH("FAPECOBR")
    Call SETULTREG("FAPECOBD", UK&)
    Call CIERRARCH("FAPECOBD")
    '
    Call HEADERS("FAPECOBA", "")
    Call HEADERS("FAPECOBA", "Vencimiento Hasta: " + Text6.TEXT)
    Call HEADERS("FAPECOBA", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    Call HEADERS("FAPECOBR", "")
    Call HEADERS("FAPECOBR", "Vencimiento Hasta: " + Text6.TEXT)
    Call HEADERS("FAPECOBR", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    Call HEADERS("FAPECOBD", "")
    Call HEADERS("FAPECOBD", "Vencimiento Hasta: " + Text6.TEXT)
    Call HEADERS("FAPECOBD", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    '
        Call FILESORT("FAPECOBA", "", 7, 7, "ASC")
        Call FILESORT("FAPECOBD", "", 8, 8, "ASC")
    '
    Select Case MODORDE%
      Case 1
        Call FILESORT("FAPECOBA", "", 1, 1, "ASC")
        Call FILESORT("FAPECOBD", "", 1, 1, "ASC")
      Case 2
        Call FILESORT("FAPECOBA", "", 2, 2, "ASC")
        Call FILESORT("FAPECOBD", "", 2, 2, "ASC")
      Case 4
        '\\\OT-08-0592-OD-07/08/08///
        '  ORDENAR POR FECHA DE EMISION
        Call FILESORT("FAPECOBA", "", 5, 5, "ASC")
        Call FILESORT("FAPECOBD", "", 5, 5, "ASC")
        '\\\OT-08-0592-OD-FIN//
    End Select
99  Screen.MousePointer = 1
    '
End Sub

Sub FAPECOB2_VIE()
    'NUEVA
    '/////**********//////
    'FALTA TESTEAR'
    '/////*****************//////////
    Screen.MousePointer = 11
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    '
    FDESD = CDbl(CVDAT(Int(FECHANUM(Text5.TEXT))))
    FLIM = CDbl(CVDAT(Int(FECHANUM(Text6.TEXT))))

    
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    '
    GRUCLI% = XVALO(Text1.TEXT)
    GRUX$ = TRIM$(Text2.TEXT)
    VENDE% = XVALO(Text3.TEXT)
    VENDX$ = TRIM$(Text4.TEXT)
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    If Option1(3) = True Then MODORDE% = 3
    '
    On Error Resume Next
    Kill LUDAT$ + "FAPECOBA.DAT"
    On Error GoTo 0
    On Error Resume Next
    Kill LUDAT$ + "FAPECOBD.DAT"
    On Error GoTo 0
    On Error Resume Next
    Kill LUDAT$ + "FAPECOBR.DAT"
    On Error GoTo 0
    '
    BIMONET% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "BIMONETA" Then
         BIMONET% = 1
       ElseIf MODOMONE$ = "MULTIMON" Then
         BIMONET% = 2
    End If
    '
    REBLX$ = REGBLAN$("FAPECOBA")
    REBLY$ = REGBLAN$("FAPECOBD")
    REBLZ$ = REGBLAN$("FAPECOBR")
    '
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    SQLX$ = SQLX$ + "AND NuClien >= " & NCMIN% & " "
    SQLX$ = SQLX$ + "AND NuClien <= " & NCMAX% & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FeVencim))<= " & FLIM & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FeVencim))>= " & FDESD & " "
    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    Set SACTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    On Error Resume Next
    With SACTEMP
      .MoveLast
      FIN& = .RecordCount
      .MoveFirst
      If Err < 1 Then
        UJ& = 0: UK& = 0: UL& = 0
        Do While Not .EOF
          UL& = UL& + 1
          Call TRACE(20, UL&, FIN&)
          NC% = !NuClien
          NFS$ = !CodiCom_Cor
          COCOLAR$ = !CODICOM_LAR
          FF% = CVINT(!FechEmisi)
          FVI% = CVINT(!FEVENCIM)
          SALDEU# = !SaldDebe - !SaldAcre
          IMORIG# = !IBruComp
          ES$ = "    "
          '
          If MODOMONEI% = 2 Then
             Call APERBASDAT("CABAN")
             CajaBanc.FindFirst "CodiCom_Lar = '" & COCOLAR$ & "' "
             If CajaBanc.NoMatch = False Then
                On Error Resume Next
                IMORIG# = CajaBanc!Impo_Dola
                On Error GoTo 0
             End If
          End If
          '
          If GRUCLI% = 0 Or GRUCOCLI%(NC%) = GRUCLI% Then
            If VENDE% = 0 Or VENDCLI%(NC%) = VENDE% Then
              If Abs(SALDEU#) > 0.005 Then
                Y$ = REBLX$
                Call REPLA(Y$, MKI$(NC%), 1, 2)
                Call REPLA(Y$, RASOCLI$(NC%), 3, 30)
                Call REPLA(Y$, TEL1CLI$(NC%), 33, 16)
                Call REPLA(Y$, NFS$, 49, 12)
                Call REPLA(Y$, MKI$(FF%), 61, 2)
                Call REPLA(Y$, MKD$(IMORIG#), 63, 8)
                Call REPLA(Y$, MKI$(FVI%), 71, 2)
                Call REPLA(Y$, MKD$(SALDEU#), 73, 8)
                Call REPLA(Y$, ES$, 97, 4)
                '
                GRUVE$ = ""
                If Check1.Value = 1 Then
                     GRUVE$ = "Grupo: " + ECOARCH$("GRUCOCLI", MKI$(GRUCOCLI%(NC%)))
                  ElseIf Check2.Value = 1 Then
                     GRUVE$ = "Vendedor: " + ECOARCH$("VENDEDOR", Chr$(VENDCLI%(NC%)))
                End If
                Call REPLA(Y$, GRUVE$, 101, 28)
                '
                UJ& = UJ& + 1
                Call GRAREG("FAPECOBA", Y$, UJ&)
                '
                If BIMONET% < 1 Then GoTo 29
                '
                Y$ = REBLY$
                Z$ = REBLZ$
                '
                Call REPLA(Y$, MKI$(NC%), 1, 2)
                Call REPLA(Y$, RASOCLI$(NC%), 3, 30)
                Call REPLA(Y$, TEL1CLI$(NC%), 33, 16)
                Call REPLA(Y$, NFS$, 49, 10)
                Call REPLA(Y$, MKI$(FF%), 59, 2)
                Call REPLA(Y$, MKD$(IMORIG#), 65, 8)
                Call REPLA(Y$, MKI$(FVI%), 73, 2)
                Call REPLA(Y$, MKD$(SALDEU#), 75, 8)
                '
                Call REPLA(Z$, Left$(Y$, 60), 1, 60)
                Call REPLA(Z$, MKD$(IMORIG#), 61, 8)
                Call REPLA(Z$, MKI$(FVI%), 69, 2)
                Call REPLA(Z$, MKD$(SALDEU#), 71, 8)
                '
                MONEMX% = 0
                On Error Resume Next   ' por si no está el campo
                MONEMX% = !Mone_Emis
                On Error GoTo 0
                If MONEMX% < 1 Then
                  COCOLAR$ = TRIM$(!CODICOM_LAR)
                  RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
                  For KUL& = Len(RFG$) - 3 To 1 Step -4
                     RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
                     XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
                     If TRIM$(Mid$(XXC$, 45, 18)) = COCOLAR$ Then
                        MONEMX% = CVI(Mid$(XXC$, 31, 2))
                        GoTo 15
                     End If
                  Next KUL&
15                If MONEMX% < 1 Then MONEMX% = 1
                End If
                '
                If MONEMX% > 1 Then
                  SALDOL# = 0
                  RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
                  For KUL& = Len(RFG$) - 3 To 1 Step -4
                    RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
                    XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
                    If Mid$(XXC$, 13, 18) = COCOLAR$ Then
                      SALDOL# = SALDOL# + CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8))
                    End If
                  Next KUL&
                  Call REPLA(Y$, MKD$(SALDOL# * TICAMONE(MONEMX%)), 75, 8)
                  Call REPLA(Y$, MKD$(SALDOL#), 91, 8)
                End If
                '
20              If MONEMX% > 1 Then
                   Call REPLA(Y$, SIMBOLPES$(MONEMX%), 61, 4)
                End If
                Call REPLA(Y$, ES$, 107, 10)
                Call REPLA(Z$, ES$, 87, 10)
                '
                UK& = UK& + 1
                Call GRAREG("FAPECOBR", Z$, UK&)
                Call GRAREG("FAPECOBD", Y$, UK&)
                '          '
29            End If
            End If
          End If
          '
        .MoveNext
        Loop
        '
      End If
    End With
    '
    Call SETULTREG("FAPECOBA", UJ&)
    Call CIERRARCH("FAPECOBA")
    Call SETULTREG("FAPECOBR", UK&)
    Call CIERRARCH("FAPECOBR")
    Call SETULTREG("FAPECOBD", UK&)
    Call CIERRARCH("FAPECOBD")
    '
    Call HEADERS("FAPECOBA", "")
    Call HEADERS("FAPECOBA", "Vencimiento Hasta: " + Text6.TEXT)
    Call HEADERS("FAPECOBA", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    Call HEADERS("FAPECOBR", "")
    Call HEADERS("FAPECOBR", "Vencimiento Hasta: " + Text6.TEXT)
    Call HEADERS("FAPECOBR", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    Call HEADERS("FAPECOBD", "")
    Call HEADERS("FAPECOBD", "Vencimiento Hasta: " + Text6.TEXT)
    Call HEADERS("FAPECOBD", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    '
    Call FILESORT("FAPECOBA", "", 7, 7, "ASC")
    Call FILESORT("FAPECOBD", "", 8, 8, "ASC")
    Select Case MODORDE%
      Case 1
        Call FILESORT("FAPECOBA", "", 1, 1, "ASC")
        Call FILESORT("FAPECOBD", "", 1, 1, "ASC")
      Case 2
        Call FILESORT("FAPECOBA", "", 2, 2, "ASC")
        Call FILESORT("FAPECOBD", "", 2, 2, "ASC")
    End Select
    Screen.MousePointer = 1
    '
End Sub

Private Sub PLANFINA()
  '\\\OT-06-0503-RG-23/11/06///
  Dim PLAFIN#(32767, 5) 'MATRIZ QUE CONTIENE EL PLAN FINANCIERO
  Dim FECOSE(5) As Integer 'FECHA COMIENZO DE SEMANA
  '
  FLIM1% = HOY(HO$)
  NUMSEM% = SEMANU%(FLIM1%, DELX, ALX)
  FECOSE%(1) = DELX 'FECHA COMIENZO SEMANA UNO
  FECOSE%(2) = DIPOST(ALX) 'FECHA COMIENZO SEMANA DOS
  NUSEM% = SEMANU%(FECOSE%(2), DELX, ALX)
  FECOSE%(3) = DIPOST(ALX) 'FECHA COMIENZO SEMANA TRES
  NUSEM% = SEMANU%(FECOSE%(3), DELX, ALX)
  FECOSE%(4) = DIPOST(ALX) 'FECHA COMIENZO SEMANA CUATRO
  NUSEM% = SEMANU%(FECOSE%(4), DELX, ALX)
  FECOSE%(5) = DIPOST(ALX) 'FECHA COMIENZO SEMANA CUATRO
  FECOSE%(0) = (-32767)
  '
'1.- GENERAR FACTURAS PENDIENTES DE COBRANZA (FABECOBR)
  'Call FAPECOB2
'2.- DISTRIBUIR PENDIENTES EN LA MATRIZ
  FIN& = ULTREG&("FAPECOBR")
  NCMAX% = 0
  For JX& = 1 To FIN&
    '
    Call TRACE(24, JX&, FIN&)
    '
    X$ = REGLEIDO$("FAPECOBR", JX&)
    '
    FV1% = CVI(Mid$(X$, 69, 2))
    NPRO1% = CVI(Mid$(X$, 1, 2))
    IMPEN1# = CVD(Mid$(X$, 71, 8))
    If NPRO1% > 0 Then
      If NPRO1% > NCMAX% Then NCMAX% = NPRO1%
    End If
    '
    For KKI% = 0 To 4
      If FV1% >= FECOSE%(KKI%) Then  'VENCIDAS ANTERIORES
        If FV1% < FECOSE%(KKI% + 1) Then
          PLAFIN#(NPRO1%, KKI%) = PLAFIN#(NPRO1%, KKI%) + IMPEN1#
          GoTo 19
        End If
      End If
    Next KKI%
    PLAFIN#(NPRO1%, 5) = PLAFIN#(NPRO1%, 5) + IMPEN1#
19 Next JX&
   '
'3.- GENERAR ARCHIVO  PLAN FINANCIERO A PARTIR DE LA MATRIZ
    '
  REBLAZ$ = REGBLAN$("PLANFIN")
  JJ& = 0
  FIN& = NCMAX%
  For JX1% = 1 To FIN&
    '
    REGACT1& = JX1%
    Call TRACE(24, REGACT1&, FIN&)
    '
    TOTAPRO1# = 0
    For XB% = 0 To 5
     TOTAPRO1# = TOTAPRO1# + PLAFIN#(JX1%, XB%)
    Next XB%
    '
    If Abs(TOTAPRO1#) > 0.005 Then 'SI TIENE PENDIENTES
      Y$ = REBLAZ$
      Call REPLA(Y$, MKI$(JX1%), 1, 2)
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 0)), 3, 8) 'VEN
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 1)), 11, 8) '1SEM
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 2)), 19, 8) '2SEM
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 3)), 27, 8) '3SEM
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 4)), 35, 8) '4SEM
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 5)), 43, 8) 'POST
      Call REPLA(Y$, MKD$(TOTAPRO1#), 51, 8)        'TOTAL
      JJ& = JJ& + 1
      Call GRAREG("PLANFIN", Y$, JJ&)
    End If
    '
  Next JX1%
  Call SETULTREG("PLANFIN", JJ&)
'5.- LISTAR
  Call FINAL("*PLANFIN")
  '\\\OT-06-0503-RG-FIN///
End Sub

Sub ATRACOBRA()
10 Call SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
   If NCL1% < 1 And NCL2% < 1 Then GoTo 99
   
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If DES% < 1 Then GoTo 10
   DES1# = Int(CDbl(CVDAT(DES%)))
   HAS1# = Int(CDbl(CVDAT(HAS%)))
   '
   SQLX$ = "SELECT * FROM CAJABANC"
   SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'RCOB'"
   SQLX$ = SQLX$ + " AND NuClien >= " & NCL1% & ""
   SQLX$ = SQLX$ + " AND NuClien <= " & NCL2% & ""
   SQLX$ = SQLX$ + "AND INT(CDBL(FechConta))>= " & DES1# & " "
   SQLX$ = SQLX$ + "AND INT(CDBL(FechConta))<= " & HAS1# & " "
   Set PROATRATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      With PROATRATEMP
          On Error Resume Next
          .MoveFirst
          If Err Then
             On Error GoTo 0
             Call MENSERR(24, "Ningun Registro Coincide con el Filtro SQL")
             GoTo 99
          End If
          '
          Y$ = REGBLAN("ATRAPROM")
          '
          U& = 0
          Do While Not .EOF
            NCLIE% = !NuClien
            If VENDSEL% < 1 Or VENDSEL% = VENDCLI%(NCLIE%) Then
               CLIENTE% = !NuClien
               NRECIBBO$ = SAFETEXT$(!CodiCom_Cor)
               ATRACHEQQUE% = XVALO(!AtraCheque)
               ATRACCANCE% = XVALO(!AtraCance)
               ATRASSO% = XVALO(!AtraCheque) + XVALO(!AtraCance)
               Call REPLA(Y$, MKI$(CLIENTE%), 1, 2)
               Call REPLA(Y$, NRECIBBO$, 3, 12)
               Call REPLA(Y$, MKS$(ATRACHEQQUE%), 15, 4)
               Call REPLA(Y$, MKS$(ATRACCANCE%), 19, 4)
               Call REPLA(Y$, MKS$(ATRASSO%), 23, 4)
               U& = U& + 1
               Call GRAREG("ATRAPROM", Y$, U&)
            End If
            .MoveNext
          Loop
          Call SETULTREG("ATRAPROM", U&)
          Call CIERRARCH("ATRAPROM")
          Call FILESORT("ATRAPROM", "", 3, 3, "ASC")
          Call FILESORT("ATRAPROM", "", 1, 1, "ASC")
     End With
     NomV$ = DEVENDE$(VENDSEL%)
     Call HEADERS("ATRAPROM", "")
     Call HEADERS("ATRAPROM", "Periodo: " + PERIO$)
     Call HEADERS("ATRAPROM", "Rango de Cuentas: " + TRIM$(Str$(NCL1%)) + " - " + TRIM$(Str$(NCL2%)))
     If VENDSEL% > 0 Then Call HEADERS("ATRAPROM", "Vendedor: " + NomV$)

     Call FINAL("*ATRAPROM")
99   Screen.MousePointer = 1
   
End Sub

