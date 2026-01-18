VERSION 5.00
Begin VB.Form FORFAC09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación de Proveedores (Operaciones Comerciales de Compras)"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6870
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
   ScaleHeight     =   8415
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10665
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13080
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   1575
         Left            =   8160
         ScaleHeight     =   1515
         ScaleWidth      =   6405
         TabIndex        =   168
         Top             =   1800
         Width           =   6465
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0F0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   2150
            Left            =   -315
            TabIndex        =   171
            Top             =   0
            Width           =   10800
            Begin VB.ListBox List1 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   22140
               IntegralHeight  =   0   'False
               Left            =   285
               TabIndex        =   172
               Top             =   315
               Width           =   6465
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   169
            Top             =   0
            Width           =   10515
            Begin VB.Line Line5 
               X1              =   6930
               X2              =   6930
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line6 
               X1              =   630
               X2              =   630
               Y1              =   0
               Y2              =   1155
            End
            Begin VB.Line Line20 
               X1              =   4725
               X2              =   4725
               Y1              =   0
               Y2              =   1155
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "  Cod.    Referencia - Cuenta                                       Importe"
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
               TabIndex        =   170
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line23 
               X1              =   8505
               X2              =   8505
               Y1              =   0
               Y2              =   1050
            End
         End
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1260
         TabIndex        =   167
         Text            =   " "
         ToolTipText     =   "Ingrese la fecha REAL del Comprobante Recibido"
         Top             =   1485
         Width           =   960
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   600
         TabIndex        =   155
         ToolTipText     =   "Ingrese Observaciones"
         Top             =   3000
         Width           =   5055
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8400
         TabIndex        =   152
         ToolTipText     =   "Ingrese Cotización"
         Top             =   7920
         Width           =   1485
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   9840
         TabIndex        =   150
         Top             =   8160
         Width           =   200
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00C0C0C0&
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
         Height          =   540
         Left            =   9840
         ScaleHeight     =   540
         ScaleWidth      =   960
         TabIndex        =   149
         Top             =   8580
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   145
         Top             =   7320
         Width           =   5355
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3720
            TabIndex        =   180
            Top             =   220
            Width           =   1485
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Imp.Neto:"
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
            Left            =   1920
            TabIndex        =   176
            Top             =   220
            Width           =   1680
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3720
            TabIndex        =   147
            Top             =   520
            Width           =   1485
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Imp. Total:"
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
            Left            =   1920
            TabIndex        =   146
            Top             =   500
            Width           =   1680
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Redondeo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   9345
         TabIndex        =   142
         Top             =   5145
         Width           =   1425
         Begin VB.TextBox Text22 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   210
            TabIndex        =   143
            Text            =   " "
            ToolTipText     =   "Importe de Redondeo"
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Importe:"
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
            Left            =   0
            TabIndex        =   144
            Top             =   380
            Width           =   960
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resumen Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   7080
         TabIndex        =   137
         Top             =   6480
         Width           =   6480
         Begin VB.TextBox Text19 
            Alignment       =   1  'Right Justify
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
            Left            =   1860
            TabIndex        =   138
            Top             =   580
            Width           =   1365
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   3360
            TabIndex        =   179
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tot.Neto"
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
            Left            =   4200
            TabIndex        =   175
            Top             =   360
            Width           =   1140
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tot.Bruto"
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
            Left            =   -120
            TabIndex        =   174
            Top             =   360
            Width           =   1140
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   240
            TabIndex        =   173
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "% Bonif."
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
            Left            =   1260
            TabIndex        =   141
            Top             =   315
            Width           =   1260
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   4800
            TabIndex        =   140
            Top             =   585
            Width           =   1455
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Importe"
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
            Left            =   2625
            TabIndex        =   139
            Top             =   315
            Width           =   1260
         End
      End
      Begin VB.TextBox Text17 
         Height          =   280
         Left            =   5685
         TabIndex        =   92
         Top             =   9990
         Width           =   960
      End
      Begin VB.TextBox Text16 
         Height          =   280
         Left            =   3165
         TabIndex        =   90
         Top             =   9990
         Width           =   2220
      End
      Begin VB.PictureBox Picture3 
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
         Height          =   9165
         Left            =   5880
         ScaleHeight     =   9105
         ScaleWidth      =   960
         TabIndex        =   75
         Top             =   120
         Width           =   1020
         Begin VB.CommandButton Command30 
            Height          =   690
            Left            =   160
            Picture         =   "FORFAC09.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   178
            ToolTipText     =   "Listado de Ordenes de Pago Por Rango de Fechas"
            Top             =   1560
            Width           =   650
         End
         Begin VB.CommandButton Command25 
            Height          =   690
            Left            =   160
            Picture         =   "FORFAC09.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   177
            ToolTipText     =   "Anulación de Orden de Pago Emitida"
            Top             =   3080
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            Caption         =   "A-B-M"
            Height          =   690
            Left            =   160
            Picture         =   "FORFAC09.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Archivo Maestro de Proveedores"
            Top             =   2280
            Width           =   650
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
            Height          =   135
            Left            =   160
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   3780
            Width           =   640
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
               Height          =   500
               Left            =   -420
               TabIndex        =   86
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Exit"
            Height          =   690
            Left            =   160
            Picture         =   "FORFAC09.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Terminar - Salir de la Aplicación"
            Top             =   110
            Width           =   650
         End
         Begin VB.CommandButton Command29 
            Caption         =   "Help"
            Height          =   690
            Left            =   160
            Picture         =   "FORFAC09.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Ayuda Flexoft en Línea"
            Top             =   800
            Width           =   650
         End
         Begin VB.CommandButton BOTNEXT 
            Caption         =   "Save"
            Height          =   735
            Left            =   160
            Picture         =   "FORFAC09.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Grabar - Registrar Documento en la Base de Datos"
            Top             =   7080
            Width           =   650
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Note"
            Height          =   690
            Left            =   160
            Picture         =   "FORFAC09.frx":107C
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Notas de Proveedores"
            Top             =   2940
            Visible         =   0   'False
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            Caption         =   "R-Cue"
            Height          =   690
            Left            =   0
            Picture         =   "FORFAC09.frx":1386
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Resumen de Cuenta Corriente Proveedores"
            Top             =   0
            Visible         =   0   'False
            Width           =   650
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Print"
            Height          =   640
            Left            =   160
            Picture         =   "FORFAC09.frx":1690
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Imprime Formulario de Carga - Vuelco de Pantalla"
            Top             =   6360
            Width           =   650
         End
         Begin VB.CommandButton Command6 
            Caption         =   "More"
            Height          =   640
            Left            =   160
            Picture         =   "FORFAC09.frx":1CFA
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Anular, Consultar, Imputación de Créditos, Tablas"
            Top             =   5640
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Setup"
            Height          =   690
            Left            =   160
            Picture         =   "FORFAC09.frx":1E44
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Opciones de Configuración"
            Top             =   4920
            Width           =   650
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00004000&
            Caption         =   "Libro I.V.A"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   160
            TabIndex        =   134
            ToolTipText     =   "Incluir / Excluir de Listado IVA Compras"
            Top             =   4480
            Value           =   1  'Checked
            Width           =   915
         End
         Begin VB.Image Image1 
            Height          =   390
            Left            =   -255
            Picture         =   "FORFAC09.frx":2286
            Top             =   4095
            Width           =   1515
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pago Contado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   9960
         TabIndex        =   65
         Top             =   9360
         Width           =   3000
         Begin VB.CommandButton Command21 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   210
            TabIndex        =   94
            ToolTipText     =   "Click Aqui para Cancelar Pago Contado "
            Top             =   280
            Width           =   280
         End
         Begin VB.CommandButton Command13 
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
            Left            =   2625
            TabIndex        =   74
            Top             =   1200
            Width           =   180
         End
         Begin VB.TextBox Text15 
            Height          =   285
            Left            =   1680
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   1200
            Width           =   960
         End
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   210
            TabIndex        =   72
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1200
            Width           =   1275
         End
         Begin VB.CommandButton Command12 
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
            Height          =   300
            Left            =   2625
            TabIndex        =   67
            Top             =   280
            Width           =   180
         End
         Begin VB.Label Label25 
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
            Height          =   300
            Left            =   1995
            TabIndex        =   71
            Top             =   280
            Width           =   645
         End
         Begin VB.Label Label24 
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
            Left            =   735
            TabIndex        =   70
            Top             =   940
            Width           =   1485
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Número:"
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
            Left            =   -630
            TabIndex        =   69
            Top             =   940
            Width           =   1485
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Medio de Pago:"
            Height          =   225
            Left            =   420
            TabIndex        =   68
            Top             =   345
            Width           =   1485
         End
         Begin VB.Label Label21 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   280
            Left            =   210
            TabIndex        =   66
            Top             =   620
            Width           =   2610
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vencimientos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   3480
         TabIndex        =   24
         Top             =   9600
         Width           =   3315
         Begin VB.Frame Frame14 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame14"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   80
            TabIndex        =   95
            Top             =   320
            Width           =   3165
            Begin VB.VScrollBar VScroll1 
               Height          =   720
               LargeChange     =   7
               Left            =   3020
               TabIndex        =   133
               Top             =   0
               Width           =   120
            End
            Begin VB.Frame Frame15 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Caption         =   "Frame15"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3705
               Left            =   0
               TabIndex        =   96
               Top             =   0
               Width           =   2955
               Begin VB.CommandButton BOTFEVEN 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   8
                  Left            =   1155
                  TabIndex        =   122
                  Top             =   3360
                  Width           =   180
               End
               Begin VB.CommandButton BOTFEVEN 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   7
                  Left            =   1155
                  TabIndex        =   119
                  Top             =   2940
                  Width           =   180
               End
               Begin VB.CommandButton BOTFEVEN 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   5
                  Left            =   1155
                  TabIndex        =   113
                  Top             =   2100
                  Width           =   180
               End
               Begin VB.CommandButton BOTFEVEN 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   4
                  Left            =   1155
                  TabIndex        =   110
                  Top             =   1680
                  Width           =   180
               End
               Begin VB.CommandButton BOTFEVEN 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   2
                  Left            =   1155
                  TabIndex        =   104
                  Top             =   840
                  Width           =   180
               End
               Begin VB.CommandButton BOTFEVEN 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   1
                  Left            =   1155
                  TabIndex        =   101
                  Top             =   420
                  Width           =   180
               End
               Begin VB.TextBox IVENCI 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   2
                  Left            =   1470
                  TabIndex        =   105
                  Text            =   " "
                  Top             =   840
                  Width           =   1485
               End
               Begin VB.TextBox IVENCI 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   1
                  Left            =   1470
                  TabIndex        =   102
                  Text            =   " "
                  Top             =   420
                  Width           =   1485
               End
               Begin VB.TextBox IVENCI 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   0
                  Left            =   1470
                  TabIndex        =   99
                  Top             =   0
                  Width           =   1485
               End
               Begin VB.TextBox FEVEN 
                  Height          =   300
                  Index           =   2
                  Left            =   210
                  TabIndex        =   103
                  Text            =   " "
                  Top             =   840
                  Width           =   1065
               End
               Begin VB.TextBox FEVEN 
                  Height          =   300
                  Index           =   1
                  Left            =   210
                  TabIndex        =   100
                  Text            =   " "
                  Top             =   420
                  Width           =   1065
               End
               Begin VB.TextBox FEVEN 
                  Height          =   300
                  Index           =   0
                  Left            =   210
                  TabIndex        =   97
                  Top             =   0
                  Width           =   960
               End
               Begin VB.CommandButton BOTFEVEN 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   0
                  Left            =   1155
                  TabIndex        =   98
                  Top             =   0
                  Width           =   180
               End
               Begin VB.TextBox IVENCI 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   5
                  Left            =   1470
                  TabIndex        =   114
                  Text            =   " "
                  Top             =   2100
                  Width           =   1485
               End
               Begin VB.TextBox IVENCI 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   4
                  Left            =   1470
                  TabIndex        =   111
                  Text            =   " "
                  Top             =   1680
                  Width           =   1485
               End
               Begin VB.TextBox IVENCI 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   3
                  Left            =   1470
                  TabIndex        =   108
                  Top             =   1260
                  Width           =   1485
               End
               Begin VB.TextBox FEVEN 
                  Height          =   300
                  Index           =   5
                  Left            =   210
                  TabIndex        =   112
                  Text            =   " "
                  Top             =   2100
                  Width           =   1065
               End
               Begin VB.TextBox FEVEN 
                  Height          =   300
                  Index           =   4
                  Left            =   210
                  TabIndex        =   109
                  Text            =   " "
                  Top             =   1680
                  Width           =   1065
               End
               Begin VB.CommandButton BOTFEVEN 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   3
                  Left            =   1155
                  TabIndex        =   107
                  Top             =   1260
                  Width           =   180
               End
               Begin VB.TextBox FEVEN 
                  Height          =   300
                  Index           =   3
                  Left            =   210
                  TabIndex        =   106
                  Top             =   1260
                  Width           =   960
               End
               Begin VB.TextBox IVENCI 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   8
                  Left            =   1470
                  TabIndex        =   123
                  Text            =   " "
                  Top             =   3360
                  Width           =   1485
               End
               Begin VB.TextBox IVENCI 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   7
                  Left            =   1470
                  TabIndex        =   120
                  Text            =   " "
                  Top             =   2940
                  Width           =   1485
               End
               Begin VB.TextBox IVENCI 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   6
                  Left            =   1470
                  TabIndex        =   117
                  Top             =   2520
                  Width           =   1485
               End
               Begin VB.TextBox FEVEN 
                  Height          =   300
                  Index           =   8
                  Left            =   210
                  TabIndex        =   121
                  Text            =   " "
                  Top             =   3360
                  Width           =   1065
               End
               Begin VB.TextBox FEVEN 
                  Height          =   300
                  Index           =   7
                  Left            =   210
                  TabIndex        =   118
                  Text            =   " "
                  Top             =   2940
                  Width           =   1065
               End
               Begin VB.CommandButton BOTFEVEN 
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   6
                  Left            =   1155
                  TabIndex        =   116
                  Top             =   2520
                  Width           =   180
               End
               Begin VB.TextBox FEVEN 
                  Height          =   300
                  Index           =   6
                  Left            =   210
                  TabIndex        =   115
                  Top             =   2520
                  Width           =   960
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "3)"
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
                  Index           =   2
                  Left            =   -1365
                  TabIndex        =   132
                  Top             =   885
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "2)"
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
                  Index           =   1
                  Left            =   -1365
                  TabIndex        =   131
                  Top             =   465
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Index           =   0
                  Left            =   -1365
                  TabIndex        =   130
                  Top             =   45
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "6)"
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
                  Index           =   6
                  Left            =   -1365
                  TabIndex        =   129
                  Top             =   2145
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "5)"
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
                  Index           =   7
                  Left            =   -1365
                  TabIndex        =   128
                  Top             =   1725
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "4)"
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
                  Index           =   8
                  Left            =   -1365
                  TabIndex        =   127
                  Top             =   1320
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "9)"
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
                  Index           =   9
                  Left            =   -1365
                  TabIndex        =   126
                  Top             =   3405
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "8)"
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
                  Index           =   10
                  Left            =   -1365
                  TabIndex        =   125
                  Top             =   2985
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "7)"
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
                  Index           =   11
                  Left            =   -1365
                  TabIndex        =   124
                  Top             =   2565
                  Width           =   1590
               End
            End
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Percep. I.Brutos"
         Height          =   1140
         Left            =   3120
         TabIndex        =   34
         Top             =   6120
         Width           =   2505
         Begin VB.TextBox Text12 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1000
            TabIndex        =   56
            Text            =   " "
            ToolTipText     =   "Ingrese Importe Percepción I.B."
            Top             =   720
            Width           =   1380
         End
         Begin VB.TextBox TEXT10 
            Height          =   300
            Left            =   315
            TabIndex        =   54
            Text            =   " "
            ToolTipText     =   "Ingrese Provincia o Jurisdicción"
            Top             =   315
            Width           =   330
         End
         Begin VB.CommandButton Command10 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   630
            TabIndex        =   55
            Top             =   315
            Width           =   180
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Importe:"
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
            Left            =   0
            TabIndex        =   39
            Top             =   820
            Width           =   960
         End
         Begin VB.Label Label18 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            Height          =   300
            Left            =   840
            TabIndex        =   38
            Top             =   315
            Width           =   1575
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "J."
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
            Left            =   -240
            TabIndex        =   37
            ToolTipText     =   "Haga Doble Click para Varias Jurisdicciones"
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Percepciones AFIP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   240
         TabIndex        =   33
         Top             =   6120
         Width           =   2145
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   600
            TabIndex        =   53
            ToolTipText     =   "Ingrese Importe Percepción Ganancias"
            Top             =   675
            Width           =   1275
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   600
            TabIndex        =   52
            ToolTipText     =   "Ingrese Importe percepción I.V.A."
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Gan:"
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
            Left            =   30
            TabIndex        =   36
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IVA:"
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
            Left            =   0
            TabIndex        =   35
            Top             =   420
            Width           =   540
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2205
         TabIndex        =   28
         Top             =   1485
         Width           =   180
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         TabIndex        =   23
         Top             =   3600
         Width           =   5385
         Begin VB.TextBox IMPOGRA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1770
            TabIndex        =   165
            ToolTipText     =   "  "
            Top             =   280
            Width           =   1845
         End
         Begin VB.TextBox IMPOIVA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   3825
            TabIndex        =   50
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1110
            Width           =   1275
         End
         Begin VB.TextBox IMPOIVA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   6
            Left            =   3825
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1920
            Width           =   1275
         End
         Begin VB.TextBox IMPOGRA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   6
            Left            =   1770
            TabIndex        =   46
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1920
            Width           =   1845
         End
         Begin VB.TextBox IMPOGRA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   1770
            TabIndex        =   45
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1650
            Width           =   1845
         End
         Begin VB.TextBox IMPOIVA 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   5
            Left            =   2520
            TabIndex        =   135
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1840
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.TextBox IMPOIVA 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   4
            Left            =   3840
            TabIndex        =   62
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1260
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame11"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            Left            =   600
            TabIndex        =   57
            Top             =   360
            Width           =   850
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "N/Grav"
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
               Index           =   5
               Left            =   0
               TabIndex        =   136
               Top             =   1350
               Width           =   795
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Import"
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
               Index           =   6
               Left            =   0
               TabIndex        =   64
               Top             =   1620
               Width           =   770
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Exento"
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
               Index           =   4
               Left            =   0
               TabIndex        =   63
               Top             =   1080
               Width           =   800
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "21.0%"
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
               Index           =   0
               Left            =   120
               TabIndex        =   61
               Top             =   0
               Width           =   750
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "27.0%"
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
               Index           =   1
               Left            =   80
               TabIndex        =   60
               Top             =   270
               Width           =   750
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "10.5%"
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
               Index           =   2
               Left            =   80
               TabIndex        =   59
               Top             =   570
               Width           =   750
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Index           =   3
               Left            =   80
               TabIndex        =   58
               Top             =   810
               Width           =   750
            End
         End
         Begin VB.TextBox IMPOGRA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   1770
            TabIndex        =   44
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1380
            Width           =   1845
         End
         Begin VB.TextBox IMPOGRA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1770
            TabIndex        =   43
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1110
            Width           =   1845
         End
         Begin VB.TextBox IMPOIVA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   3825
            TabIndex        =   49
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   840
            Width           =   1275
         End
         Begin VB.TextBox IMPOGRA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   1770
            TabIndex        =   42
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   860
            Width           =   1845
         End
         Begin VB.TextBox IMPOIVA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   3825
            TabIndex        =   48
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   570
            Width           =   1275
         End
         Begin VB.TextBox IMPOGRA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1770
            TabIndex        =   41
            ToolTipText     =   "  "
            Top             =   570
            Width           =   1845
         End
         Begin VB.TextBox IMPOIVA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   3825
            TabIndex        =   47
            ToolTipText     =   "  "
            Top             =   300
            Width           =   1275
         End
         Begin VB.Line Line13 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   3885
            Y1              =   2310
            Y2              =   2310
         End
      End
      Begin VB.PictureBox Picture2 
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
         Height          =   1580
         Left            =   7680
         ScaleHeight     =   1515
         ScaleWidth      =   10470
         TabIndex        =   7
         Top             =   3360
         Width           =   10530
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C000&
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
            Height          =   5630
            Left            =   0
            TabIndex        =   8
            Top             =   260
            Width           =   10515
            Begin VB.ListBox List2 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   935
               IntegralHeight  =   0   'False
               Left            =   0
               TabIndex        =   9
               Top             =   0
               Width           =   10550
            End
            Begin VB.ListBox List3 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   935
               IntegralHeight  =   0   'False
               Left            =   -30
               TabIndex        =   25
               Top             =   0
               Width           =   10455
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   8400
               TabIndex        =   16
               Top             =   980
               Width           =   1695
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Width           =   1695
            End
            Begin VB.Line Line4 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   2415
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   6615
               TabIndex        =   12
               Top             =   980
               Width           =   1695
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3570
               TabIndex        =   11
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   " Total General"
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
               Left            =   105
               TabIndex        =   10
               Top             =   980
               Width           =   1695
            End
            Begin VB.Line Line14 
               X1              =   0
               X2              =   10500
               Y1              =   935
               Y2              =   935
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   10500
               Y1              =   1250
               Y2              =   1250
            End
            Begin VB.Line Line16 
               X1              =   2520
               X2              =   2520
               Y1              =   735
               Y2              =   2415
            End
            Begin VB.Line Line17 
               X1              =   3735
               X2              =   3735
               Y1              =   735
               Y2              =   2415
            End
            Begin VB.Line Line18 
               X1              =   5460
               X2              =   5460
               Y1              =   630
               Y2              =   2415
            End
            Begin VB.Line Line19 
               X1              =   6720
               X2              =   6720
               Y1              =   735
               Y2              =   2415
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   10515
            Begin VB.Line Line1 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line2 
               X1              =   6720
               X2              =   6720
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   " Documento  Número         Emision                Importe       Vencim.                  Saldo                Acredit.       "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   10410
            End
            Begin VB.Line Line9 
               X1              =   2520
               X2              =   2520
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line10 
               X1              =   3730
               X2              =   3730
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line11 
               X1              =   5460
               X2              =   5460
               Y1              =   0
               Y2              =   1050
            End
         End
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2865
         TabIndex        =   1
         ToolTipText     =   "Ingrese Número de Comprobante"
         Top             =   2520
         Width           =   2805
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   4515
         TabIndex        =   0
         Text            =   " "
         ToolTipText     =   "Ingrese la Fecha Correspondiente al Mes de Contabilización del I.V.A."
         Top             =   1500
         Width           =   960
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5460
         TabIndex        =   2
         Top             =   1500
         Width           =   180
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   960
         TabIndex        =   19
         Top             =   1920
         Width           =   4935
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "DP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   4020
            TabIndex        =   89
            ToolTipText     =   "Débito Propio"
            Top             =   230
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "CP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   2820
            TabIndex        =   88
            ToolTipText     =   "Crédito Propio"
            Top             =   230
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "CO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   105
            TabIndex        =   40
            ToolTipText     =   "Comprobante Genérico (Extracto Bancario - Póliza de Seguro - Liquidación)"
            Top             =   30
            Width           =   705
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "RH"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   420
            TabIndex        =   27
            ToolTipText     =   "Recibo de Honorarios"
            Top             =   230
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "TF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   1305
            TabIndex        =   26
            ToolTipText     =   "Ticket - Factura"
            Top             =   30
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "NC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   3705
            TabIndex        =   22
            ToolTipText     =   "Nota de Crédito del Proveedor"
            Top             =   30
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "ND"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   2505
            TabIndex        =   21
            ToolTipText     =   "Nota de Débito del Proveedor"
            Top             =   30
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "FC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   1620
            TabIndex        =   20
            ToolTipText     =   "Factura de Compra"
            Top             =   230
            Value           =   -1  'True
            Width           =   700
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   120
         TabIndex        =   157
         Top             =   0
         Width           =   5655
         Begin VB.TextBox Text1 
            Height          =   300
            Left            =   120
            MaxLength       =   5
            TabIndex        =   164
            Text            =   "  "
            ToolTipText     =   "Ingrese Número de Proveedor"
            Top             =   255
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   960
            TabIndex        =   163
            Top             =   270
            Width           =   180
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1275
            TabIndex        =   162
            TabStop         =   0   'False
            Text            =   " "
            Top             =   240
            Width           =   4200
         End
         Begin VB.TextBox Text3 
            Height          =   280
            Left            =   1275
            TabIndex        =   161
            TabStop         =   0   'False
            Text            =   " "
            Top             =   510
            Width           =   4200
         End
         Begin VB.TextBox Text4 
            Height          =   280
            Left            =   1275
            TabIndex        =   160
            TabStop         =   0   'False
            Text            =   " "
            Top             =   780
            Width           =   750
         End
         Begin VB.TextBox Text5 
            Height          =   280
            Left            =   2115
            TabIndex        =   159
            TabStop         =   0   'False
            Text            =   " "
            Top             =   780
            Width           =   3360
         End
         Begin VB.TextBox Text13 
            Height          =   280
            Left            =   1275
            TabIndex        =   158
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1050
            Width           =   4200
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   166
            ToolTipText     =   "Click Para Cambiar Categoria IVA"
            Top             =   720
            Width           =   660
         End
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Obs:"
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
         Left            =   -600
         TabIndex        =   156
         Top             =   3030
         Width           =   1140
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pesos"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   8430
         TabIndex        =   154
         ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
         Top             =   8160
         Width           =   1275
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cotiz.:"
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
         Left            =   8400
         TabIndex        =   153
         Top             =   7680
         Width           =   645
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Carga por Detalle"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10095
         TabIndex        =   151
         Top             =   8160
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   9000
         Picture         =   "FORFAC09.frx":41A8
         Top             =   8160
         Width           =   480
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9420
         TabIndex        =   148
         Top             =   8265
         Width           =   450
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   5490
         TabIndex        =   93
         Top             =   10020
         Width           =   330
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "C.A.I.:"
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
         Left            =   2490
         TabIndex        =   91
         Top             =   10050
         Width           =   960
      End
      Begin VB.Label Label26 
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3600
         TabIndex        =   87
         Top             =   11880
         Width           =   1035
      End
      Begin VB.Line Line22 
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         X1              =   10740
         X2              =   10740
         Y1              =   6405
         Y2              =   7960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Emision:"
         Height          =   435
         Left            =   210
         TabIndex        =   29
         Top             =   1420
         Width           =   960
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   9540
         TabIndex        =   18
         Top             =   8895
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Valores:"
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
         Left            =   7230
         TabIndex        =   17
         Top             =   8895
         Width           =   2220
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Factura Nro:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   -120
         TabIndex        =   6
         Top             =   2535
         Width           =   2955
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Contable:"
         Height          =   435
         Left            =   3465
         TabIndex        =   5
         Top             =   1440
         Width           =   960
      End
   End
   Begin VB.Label ONEGRA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   840
      TabIndex        =   30
      Top             =   0
      Width           =   1515
   End
   Begin VB.Label IMPIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   2415
      TabIndex        =   32
      Top             =   0
      Width           =   1230
   End
   Begin VB.Label IMPIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   2415
      TabIndex        =   31
      Top             =   0
      Width           =   1230
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA:"
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
      TabIndex        =   4
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "FORFAC09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$
Dim TOBRUFAC#        ' total bruto factura
'
Dim MODORPA%
Dim MODOSORVAL%
Dim TIDOCOM$, LETRAFA$, NUDOCUCOM$
Dim TOFAC#(10), TOFAC1#(10), FVII%(10), IMV#(10), IMV1#(10), GRU%(10), GRU1%(10), IPAR#(10), IPAR1#(10)
'
Dim TOTIMPU#, TONEGRA#, TOEXEN#, TOTIVA#, TORETIV#, TOTFAC#, ALICUIVA
'
Dim CAMON%, DMONE$(256), TCMONE#(256)
'

Private Sub ajuste_Click()
   Call FACOCOMP07.Command4_Click
End Sub

Private Sub ajuste1_Click()
   Call FACOCOMP07.Command4_Click
End Sub

Private Sub anu_comp1_Click()
   Call FACOCOMP07.Command25_Click
End Sub

Private Sub anul_comp_Click()
   Call FACOCOMP07.Command25_Click
End Sub

'
Private Sub BOTFEVEN_Click(Index As Integer)
   Call SELECHO("SELFECHA")
   VDEVU$ = VALACT1$("SELFECHA")
   If VDEVU$ <> "" Then
     FEVEN(Index).TEXT = VDEVU$
     FEVEN(Index).Refresh
   End If
End Sub

Private Sub BOTNEXT_Click()
    '
    COTIMONEM# = TCMONE#(MONEMI%)
    NiveCla% = 0
    If Picture4.Visible = True Then NiveCla% = XVALO(Mid$(Label40, 2))
    If Control$("", "VALICAI") = "SI" Then
        If VALIDACAI = 1 Then
          Call GRAFACOM
          MODIFIX% = 0
        End If
     Else
        Call GRAFACOM
        MODIFIX% = 0
    End If
    COTIMONEM# = 0
    Screen.MousePointer = 1
    '
End Sub
'
Private Sub Command1_Click()
    '
    If MODIFIX% > 0 Then
       If COMALTER%("Abandona sin Guardar los Cambios ?\\No, Continuar\Si, Cancelar") <> 2 Then
          Exit Sub
       End If
    End If
    '
    Call SELECHO("ACREDOR")
    NCXI% = XVALO(TRIM$(VALACT1$("ACREDOR")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        Text13.TEXT = CUITCLI$((-1) * NC%)
        Label26 = CODICLI((-1) * NC%)
          If LETRAFA$ <> "M" Then LETRAFA$ = "A"
          PVCL% = PIVACLI%(-1 * NC%)
          If PVCL% = 0 Or PVCL% = 2 Or PVCL% = 6 Or PVCL% = 4 Then LETRAFA$ = "C"
          If PVCL% = 5 Then LETRAFA$ = "E"
        Label20 = LETRAFA$
        If Left$(Text2.TEXT, 6) = "VARIOS" Then
            Text13.SetFocus
          Else
            Call BUSCACAI(NC%)
            If PVCL% = 0 Or PVCL% = 5 Or PVCL% = 6 Then
                Text6.SetFocus     ' FECHA DE EMISION
              Else
                Text16.SetFocus    ' NUMERO DE CAI
            End If
        End If
        'Call SIGUEFACO1
    End If

    '
End Sub

Private Sub Command10_Click()
    Call SELECHO("PROVINC")
    DPX1$ = VALACT1$("PROVINC")
    If DPX1$ <> "" Then
        DP1$ = DPX1$
        TEXT10.TEXT = TRIM$(DP1$)
        Label18.Caption = ECOARCH$("PROVINC", DP1$)
    End If
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call SELECHO("MEPAGO")
   If VALACT1$("MEPAGO") <> "" Then
     TPV% = XVALO(VALACT1$("MEPAGO"))
     If CUEMEPAG%(TPV%) = CUEICARCH% Then
         Call MENSERR(24, "Cartera Valores no Disponible\para Compras de Contado.\   \Debe Cancelar por Orden de Pago.")
         Label25 = ""
       Else
         Label25 = VALACT1$("MEPAGO")
     End If
   End If
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Call SELECHO("SELFECHA")
   Text15 = VALACT1$("SELFECHA")
   Text15.Refresh
End Sub

Private Sub Command20_Click()
   Unload Me
End Sub

Private Sub ctacte_imp_Click()
   Call RECUEP07.Command3_Click
End Sub

Private Sub ctacte_pan_Click()
   Call Command4_Click
End Sub

Private Sub Command21_Click()
   Label25 = ""
   Text14 = ""
   Text15 = ""
End Sub

Private Sub Command25_Click()

 '
    If DERACCE%("ANUFACO", "Anulación de Facturas de Compra") < 2 Then
      Exit Sub
    End If
    '
10  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ <> "" Then
      desde = FETEXCOR1$(DES%)
      hasta = FETEXCOR1$(HAS%)
    Else
      Exit Sub
    End If
    '
    Call APERBASDAT("CABAN")
    '
50  SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FCOM' OR TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + " AND FechEmisi >= '" & desde & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & hasta & "' "
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JJ& = 0
    NIVCONXXX% = NIVCONFI%
    With CABATEMP
    Do While Not .EOF
       If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70
       NCLIE% = !NuProve
       IMPO# = !ValAbsComp
       If IMPO# >= 0.005 Then
         IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
         NPX& = XVALO(Mid$(!CodiCom_Lar, 13, 6))
         If Len(!CodiCom_Lar) < 18 Then
           NPX& = XVALO(Right$(!CodiCom_Lar, 6))
         End If
         If NPX& < 1 Then NPX& = 1
         Y$ = MKS$(NPX&) + AJUSTI$(!CodiCom_Lar, 18) + IMPOS$
         JJ& = JJ& + 1
         Call GRAREG("!SELCUECO", Y$, JJ&)
       End If
70      .MoveNext
    Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    Call SETULTREG("!SELCUECO", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUECO") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Registra\Comprobantes de Compra que Puedan Anularse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUECO")
    Call FRESHECHO("!SELCUECO")
    '
    Call SELECHO("!SELCUECO")
    If VALACT1$("!SELCUECO") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUECO"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUECO"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    
    '\\\OT-08-0220-FL-06/05/08///
    If VERICONCILIA%(DOCULA$) < 1 Then Exit Sub
    '\\\OT-08-0220-FL-FIN///
    Call APERBASDAT("CABAN")
    '
    '\\\OT-06-0563-RG-21/12/06///
    RCLI$ = RASOCLI$((-1) * NC%)
    OPPX% = COMALTER%("Confirma la Anulación del Comprobante '" + TRIM$(DOCULA$) + "'\de " + RCLI$ + " \\No, Conservar\Si, Anular")
    If OPPX% <> 2 Then GoTo 10
    '\\\OT-06-0563-RG-FIN///
    '
strsql = "Delete CajaBanc where CodiCom_Lar = '" & DOCULA$ & "' "
'strsql = strsql & " AND NuProve = " & NC%

strsql = strsql & " Delete SacrePen where CodiCom_Lar = '" & DOCULA$ & "'"
'strsql = strsql & " AND NuProve = " & NC%

'se agrego borrado del registro en aplipago
strsql = strsql & " Delete ApliPago where CoDocOr_Lar = '" & DOCULA$ & "'"
strsql = strsql & " OR CoDocAp_Lar = '" & DOCULA$ & "'"
'strsql = strsql & " AND NuProve = " & NC%

Call EJECUTACOMANDO(strsql)

    '
99  JJ& = 0
    For j& = 1 To ULTREG("DETFACOM")
       XX$ = REGLEIDO("DETFACOM", j&)
       NP1% = CVI(Mid$(XX$, 1, 2))
       NF1$ = TRIM$(Mid$(XX$, 3, 18))
       If NF1$ <> DOCULA$ Then ' <> COMPROBANTE
         JJ& = JJ& + 1
         Call GRAREG("DETFACOM", XX$, JJ&)
       ElseIf NP1% <> NC% Then ' MISMO COMPROBANTE Y <> PROVEEDOR
         JJ& = JJ& + 1
         Call GRAREG("DETFACOM", XX$, JJ&)
       End If
     Next j&
     Call SETULTREG("DETFACOM", JJ&)
     '
     JJ& = 0
     For j& = 1 To ULTREG("FACBONCO")
       XX$ = REGLEIDO("FACBONCO", j&)
       NP1% = CVI(Mid$(XX$, 1, 2))
       NF1$ = TRIM$(Mid$(XX$, 9, 18))
       If NF$ <> NF1$ Then ' <> COMPROBANTE
         JJ& = JJ& + 1
         Call GRAREG("FACBONCO", XX$, JJ&)
       ElseIf NC% <> NP1% Then ' MISMO COMPROBANTE Y <> PROVEEDOR
         JJ& = JJ& + 1
         Call GRAREG("FACBONCO", XX$, JJ&)
       End If
     Next j&
     Call SETULTREG("FACBONCO", JJ&)
     '
    Screen.MousePointer = 1
    Call COMUNI("ANULACIóN COMPLETA.")
'    RECUEP07.Text1 = TRIM$(Str$(NC%))
'    RECUEP07.Show 1
    '
    GoTo 50
    '
End Sub


Private Sub Command29_Click()
   Call HELPONLINE("ACRE_GES")
End Sub

Private Sub cta_cte_Click()
   Call Command4_Click
End Sub

Private Sub cuencon_Click()
   Call OPFACOM07.Command4_Click
End Sub

Private Sub Exit_Click()
   Call Command20_Click
End Sub

Private Sub Command30_Click()
    Command30.Enabled = False
    '
    Dim RELIVA$(32)
    Dim IMASI#(32767)
    CUEMX% = 0
    Call SETULTREG("LIVACOM", 0)
    Call SETULTREG("LIVACCAI", 0)
    '
    ' Leer las Cuentas Contables de IVA
    Dim CUEIVA%(), TASIVANOM$(), ICUEN#(), CACUEIVA%
    Dim NUMEASI$(), FECHASI%(), CANASIX%
    CACUEIVA% = 0
    CANASIX% = 0
    '
    For IXX% = 0 To 6
      TXXX$ = TRIM$(Str$(IXX%))
      CIVX$ = TRIM$(Control$("", "CUEIVA-" + TXXX$))
      If CIVX$ <> "" Then
        CIVY% = CUEINT%(CIVX$)
        If CIVY% > 0 Then
          TASAX$ = Control$("", "TASIVA-" + TXXX$)
          If XVALO(TASAX$) > 0 Or IXX% > 3 Then
            If IXX% = 4 Then TASAX$ = "Aduana"
            If IXX% = 5 Then TASAX$ = "Per-IVA"
            If IXX% = 6 Then TASAX$ = "Per-GAN"
            CACUEIVA% = CACUEIVA% + 1
            ReDim Preserve CUEIVA%(CACUEIVA%), TASIVANOM$(CACUEIVA%), ICUEN#(CACUEIVA%)
            CUEIVA%(CACUEIVA%) = CIVY%
            TASIVANOM$(CACUEIVA%) = TASAX$
          End If
        End If
      End If
    Next IXX%
    '
    ' SUPRIME ASIENTOS DE DIARIO
    'If Control$("LIVACOM", "SUPRIVA") = "SI" Then GoTo 22
    '
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
      Command30.Enabled = True
      Exit Sub
    End If
    Screen.MousePointer = 11
    Call MESANO$(MESASEL$, DES%, HAS%)
    DESDAT = FETEXCOR1$(DES%)
    HASDAT = FETEXCOR1$(HAS%)
22  Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta<= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FCOM' "
    '\\\OT-05-0264-WAR-09/05/05///
    SQLX$ = SQLX$ + "AND (IncSubIva <> 'NO' "
    SQLX$ = SQLX$ + "OR IncSubIva is null OR IncSubIva = '') "
    '\\\OT-05-0264-WAR-FIN///
    '\\\OT-07-0063-RG-15/02/07///
    SQLX$ = SQLX$ + "Order by FechConta ASC, FechAcred ASC, FeHoReal ASC "
    '\\\OT-07-0063-RG-FIN///
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    JJ& = 0
    REBLA$ = REGBLAN$("LIVACCAI")
    REBLB$ = REGBLAN$("REIVACOM")
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      FIN& = .RecordCount
      .MoveFirst
      On Error GoTo 0
      U& = 0
      NIVCONXXX% = NIVCONFI%
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        '\\\OT-08-1032-OD-11/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
        If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70 'VER QUE FACTURAS MOSTRAR EN LISTADO
        '\\\OT-08-1032-OD-FIN

        'ver ernesto
        'FECHAD = Int(CDbl(!FechConta))
          'GoSub INCLUASDIAR
        'jandy sql
        FECHAD = FETEXCOR1$(CVINT%(!FECHCONTA))
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !NuProve
        ITOTA# = !ValAbsComp
        SIG% = Sgn(!IHABECOMP - !IDEBECOMP)
        MOBASI$ = TRIM$(!TipoMovim)
        '
        ZZ$ = REBLA$
        Call REPLA(ZZ$, MKI$(CVINT(!FECHCONTA)), 3, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!fechacred)), 5, 2)
        Call REPLA(ZZ$, MKS$(!NumeComp), 7, 4)
        Call REPLA(ZZ$, !CodiCom_Lar, 11, 18)
        Call REPLA(ZZ$, MKI$(NCLIE%), 31, 2)
        Call REPLA(ZZ$, RASOCLI$((-1) * NCLIE%), 33, 30)
        Call REPLA(ZZ$, CUITCLI$((-1) * NCLIE%), 63, 16)
        Call REPLA(ZZ$, POSIVAC$((-1) * NCLIE%), 79, 12)
        '\\\OT-05-0264-WAR-09/05/05///
        If IsNull(!NROCAI) Then
           NCAI$ = ""
          Else
           NCAI$ = !NROCAI
        End If
        If IsNull(!FEVTOCAI) Then
            FVCAI = 0
          Else
            FVCAI = !FEVTOCAI
        End If
        Call REPLA(ZZ$, NCAI$, 257, 24)
        Call REPLA(ZZ$, MKI$(CVINT(FVCAI)), 281, 2)
        '\\\OT-05-0264-WAR-FIN///
          ' LO QUE SIGUE ES PARA CONSIDERAR CLIENTES VARIOS
          If UCase$(Mid$(ZZ$, 33, 6)) = "VARIOS" Then
            REFEMO$ = !RefeCom
            PPXX% = InStr(REFEMO$, "//")
            If PPXX% > 0 Then
              RSCL$ = TRIM$(Left$(REFEMO$, PPXX% - 1))
              CTCL$ = TRIM$(Mid$(REFEMO$, PPXX% + 2))
              If RSCL$ <> "" Then
                If CTCL$ <> "" Then
                  Call REPLA(ZZ$, RSCL$, 33, 30)
                  Call REPLA(ZZ$, CTCL$, 63, 16)
                End If
              End If
            End If
          End If
          '
        Call REPLA(ZZ$, MKD$(ITOTA# * SIG%), 91, 8)
        '\\\OT-05-0709-WAR-02/11/05///
        Pi% = PIVACLI%(NCLIE% * (-1))
        If Pi% = 5 Then
          Call REPLA(ZZ$, MKD$(ITOTA# * SIG%), 185, 8)
        End If
        '\\\OT-05-0709-WAR-FIN///
        RESTOBRU# = ITOTA# * SIG%
        '
        '
        RELIVA$(0) = ZZ$
        For KU% = 1 To 32
          RELIVA$(KU%) = REBLA$
        Next KU%
        '
        'VER ACA
        Dim CABAFETMP As ADODB.Recordset
        Set CABAFETMP = New ADODB.Recordset
        CABAFETMP.CursorLocation = adUseClient
        SQLX$ = "SELECT * FROM CAJABANC "
        'SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
        'SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
        SQLX$ = SQLX$ + "WHERE FechConta = '" & FECHAD & "' "
        SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX$ = SQLX$ + "AND NuProve = " & NCLIE% & " "
        SQLX$ = SQLX$ + "AND TipoMovim <> '" & MOBASI$ & "' "
24      On Error Resume Next
        CABAFETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 24
       End If
       On Error GoTo 0
         '
        PROIVA% = 0: ULIVA% = 0: PROIBRU% = 0: ULIBRU% = 0
        With CABAFETMP
        If Not (.EOF And .BOF) Then
        .MoveFirst
        Do While Not .EOF
            TIMO$ = UCase$(TRIM$(!TipoMovim))
            OPMO$ = !OpciMovim
            VAMO$ = !VariMovim
            IPAX# = Abs(!ValAbsComp)
            BASE# = 0
            On Error Resume Next
            BASE# = Abs(!BASEIMPO)
            On Error GoTo 0
            '
            If Abs(IPAX#) >= 0.005 Or BASE# >= 0.005 Then
              If TIMO$ = "IVA" And UCase$(VAMO$) = "PER-IVA" Then
                  ZZ$ = RELIVA$(0)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 161, 8)
                  RELIVA$(0) = ZZ$
                  RESTOBRU# = RESTOBRU# - IPAX# * SIG%
                ElseIf (TIMO$ = "GAN" Or TIMO$ = "IVA") And UCase$(VAMO$) = "PER-GAN" Then
                  ZZ$ = RELIVA$(0)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 169, 8)
                  RELIVA$(0) = ZZ$
                  RESTOBRU# = RESTOBRU# - IPAX# * SIG%
                ElseIf TIMO$ = "IBRU" And UCase$(VAMO$) = "PERC" Then
                  ZZ$ = RELIVA$(0)
                  ACUANTEX# = CVD(Mid$(ZZ$, 193, 8))
                  Call REPLA(ZZ$, MKD$(ACUANTEX# + IPAX# * SIG%), 193, 8)
                  RESTOBRU# = RESTOBRU# - IPAX# * SIG%
                  Call REPLA(ZZ$, MKI$(!CUECONTAI), 121, 2)
                    JURIB$ = TRIM$(Mid$(OPMO$, 4))
                  Call REPLA(ZZ$, JURIB$, 203, 1)
                  Call REPLA(ZZ$, ECOARCH$("PROVINC", JURIB$), 205, 24)
                  If Abs(ACUANTEX#) >= 0.005 Then
                    Call REPLA(ZZ$, " ", 203, 1)
                    Call REPLA(ZZ$, "JURISD. VARIAS", 205, 24)
                  End If
                  RELIVA$(0) = ZZ$
                ElseIf TIMO$ = "IVA" Then
                  ZZ$ = RELIVA$(PROIVA%)
                  Call REPLA(ZZ$, VAMO$, 129, 8)
                  Call REPLA(ZZ$, MKD$(BASE# * SIG%), 137, 8)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 145, 8)
                  RESTOBRU# = RESTOBRU# - IPAX# * SIG% - BASE# * SIG%
                  If BASE# > 0.005 Then
                    Call REPLA(ZZ$, MKD$(100 * IPAX# / BASE#), 153, 8)
                  End If
                  RELIVA$(PROIVA%) = ZZ$
                  ULIVA% = PROIVA%
                  PROIVA% = PROIVA% + 1
              End If
            End If
            '
            .MoveNext
            Loop
            '
          End If
        End With
        CABAFETMP.Close
        Set CABAFETMP = Nothing
        '
       
        If Abs(RESTOBRU#) > 0.005 Then
          For KU% = 0 To ULIVA%
            ZZ$ = RELIVA$(KU%)
            If TRIM$(Mid$(ZZ$, 129, 8)) = "" Then
              If Abs(CVD(Mid$(ZZ$, 137, 8))) < 0.005 Then
                If Abs(CVD(Mid$(ZZ$, 145, 8))) < 0.005 Then
                  Call REPLA(ZZ$, "No Grav.", 129, 8)
                  Call REPLA(ZZ$, MKD$(RESTOBRU#), 137, 8)
                  RELIVA$(KU%) = ZZ$
                  GoTo 50
                End If
              End If
            End If
          Next KU%
          ULIVA% = ULIVA% + 1
          ZZ$ = RELIVA$(ULIVA%)
          Call REPLA(ZZ$, "No Grav.", 129, 8)
          Call REPLA(ZZ$, MKD$(RESTOBRU#), 137, 8)
          RELIVA$(ULIVA%) = ZZ$
        End If
        '
50      If ULIBRU% > ULIVA% Then ULIVA% = ULIBRU%
        For KU% = 0 To ULIVA%
          JJ& = JJ& + 1
          Call GRAREG("LIVACOM", RELIVA$(KU%), JJ&)
          Call GRAREG("LIVACCAI", RELIVA$(KU%), JJ&)
        Next KU%
        '
70    .MoveNext
      Loop
    End With

    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    FECHAD = CDbl(CDate("31/12/2050"))
    GoSub INCLUASDIAR
    '
    Call SETULTREG("LIVACOM", JJ&)
    Call SETULTREG("LIVACCAI", JJ&)
    '
    Call HEADERS("LIVACOM", "")
    Call HEADERS("LIVACOM", "Corresponde a: " + TRIM$(CORRESP$))
    If ULTREG&("LIVACOM") > 0 Then
      Call FINAL("*LIVCOMEN")
    End If
    '
    '
    ' CARGA RESUMEN MENSUAL DE IVA - COMPRAS
    Dim DECATIV$(128)
    CACATIV% = 0
    Dim VECREIVA$(256)
    UVECRE% = 0
    '
    For KUL& = 1 To JJ&
      RECUE$ = REGLEIDO$("LIVACOM", KUL&)
      If TRIM$(Mid$(RECUE$, 11, 18)) <> "" Then
        NUCO$ = TRIM$(Mid$(RECUE$, 11, 18))
        SIG% = Sgn(CVD(Mid$(RECUE$, 91, 8)))
      End If
      '
      CORUBRO% = 1: DERUBRO$ = "COMPRA"
      COCONCE% = 8: DECONCE$ = "COMPROBANTES VARIOS"
      '
      If SIG% < 0 Then
          COCONCE% = 9: DECONCE$ = "NOTAS DE CREDITO"
        Else
          If Left$(NUCO$, 2) = "FC" Then COCONCE% = 1: DECONCE$ = "FACTURAS"
          If Left$(NUCO$, 2) = "ND" Then COCONCE% = 2: DECONCE$ = "NOTAS DE DEBITO"
          If Left$(NUCO$, 2) = "TF" Then COCONCE% = 3: DECONCE$ = "TICKET-FACTURA"
          If Left$(NUCO$, 2) = "RH" Then COCONCE% = 4: DECONCE$ = "RECIBO HONORARIOS"
      End If
      '
      CAIVA$ = Mid$(RECUE$, 79, 12)
      If TRIM$(CAIVA$) = "" Then CAIVA$ = CAIVANT$
      DECATE$ = TRIM$(Mid$(RECUE$, 129, 8)) + " - " + CAIVA$
      CAIVANT$ = CAIVA$
      For KLI% = 1 To CACATIV%
        If DECATE$ = DECATIV$(KLI%) Then
          COCATE% = KLI%
          GoTo 20
        End If
      Next KLI%
      CACATIV% = CACATIV% + 1
      DECATIV$(CACATIV%) = DECATE$
      COCATE% = CACATIV%
      '
20    BASEIMPO# = CVD(Mid$(RECUE$, 137, 8))
      VALOIMPU# = CVD(Mid$(RECUE$, 145, 8))
      
82    If (Abs(BASEIMPO#) + Abs(VALOIMPU#)) >= 0.005 Then
        ABUS$ = Chr$(CORUBRO%) + AJUSTI$(DERUBRO$, 15)
        ABUS$ = ABUS$ + Chr$(COCONCE%) + AJUSTI$(DECONCE$, 23)
        ABUS$ = ABUS$ + Chr$(COCATE%) + AJUSTI$(DECATE$, 23)
        '
        
        For UVE% = 1 To UVECRE%
          If Left$(VECREIVA$(UVE%), 64) = ABUS$ Then
             VCC$ = VECREIVA$(UVE%)
             GoTo 84
          End If
        Next UVE%
        UVECRE% = UVECRE% + 1
        UVE% = UVECRE%
        VCC$ = REBLB$
        Call REPLA(VCC$, ABUS$, 1, 64)
        '
84      BASEIMPO# = BASEIMPO# + CVD(Mid$(VCC$, 65, 8))
        VALOIMPU# = VALOIMPU# + CVD(Mid$(VCC$, 81, 8))
        Call REPLA(VCC$, MKD$(BASEIMPO#), 65, 8)
        Call REPLA(VCC$, MKD$(VALOIMPU#), 81, 8)
        Call REPLA(VCC$, MKD$(BASEIMPO# + VALOIMPU#), 97, 8)
        VECREIVA$(UVE%) = VCC$
      End If
      '
      CORUBRO% = 2: DERUBRO$ = "PERCEPCION"        ' PERCEPCIONES
      For COCO% = 1 To 3
        COCONCE% = 10 + COCO%
        BASEIMPO# = 0
        COCATE% = 0: DECATE$ = ""
        VALOIMPU# = CVD(Mid$(RECUE$, 153 + 8 * COCO%, 8))
        If COCO% = 1 Then
           DECONCE$ = "PERCEPCION I.V.A."
          ElseIf COCO% = 2 Then
           DECONCE$ = "PERCEPCION GANANCIAS"
          ElseIf COCO% = 3 Then
           DECONCE$ = "PERCEPC.ING.BRUTOS"
           DECATE$ = Mid$(RECUE$, 205, 24)
           VALOIMPU# = CVD(Mid$(RECUE$, 193, 8))
        End If
        '
92      If (Abs(BASEIMPO#) + Abs(VALOIMPU#)) >= 0.005 Then
          ABUS$ = Chr$(CORUBRO%) + AJUSTI$(DERUBRO$, 15)
          ABUS$ = ABUS$ + Chr$(COCONCE%) + AJUSTI$(DECONCE$, 23)
          ABUS$ = ABUS$ + Chr$(COCATE%) + AJUSTI$(DECATE$, 23)
          '
          For UVE% = 1 To UVECRE%
            If Left$(VECREIVA$(UVE%), 64) = ABUS$ Then
               VCC$ = VECREIVA$(UVE%)
               GoTo 94
            End If
          Next UVE%
          UVECRE% = UVECRE% + 1
          UVE% = UVECRE%
          VCC$ = REBLB$
          Call REPLA(VCC$, ABUS$, 1, 64)
          '
94        BASEIMPO# = BASEIMPO# + CVD(Mid$(VCC$, 65, 8))
          VALOIMPU# = VALOIMPU# + CVD(Mid$(VCC$, 81, 8))
          Call REPLA(VCC$, MKD$(BASEIMPO#), 65, 8)
          Call REPLA(VCC$, MKD$(VALOIMPU#), 81, 8)
          Call REPLA(VCC$, MKD$(BASEIMPO# + VALOIMPU#), 97, 8)
          VECREIVA$(UVE%) = VCC$
        End If
         '
      Next COCO%
      '
    Next KUL&
    '
    JJK& = 0
    For UI1% = 1 To UVECRE%
      JJK& = JJK& + 1
      Call GRAREG("REIVACOM", VECREIVA$(UI1%), JJK&)
    Next UI1%
    Call SETULTREG("REIVACOM", JJK&)
    Call CIERRARCH("REIVACOM")
    '
    Call FILESORT("REIVACOM", "", 6, 6, "ASC")
    Call FILESORT("REIVACOM", "", 3, 3, "ASC")
    Call FILESORT("REIVACOM", "", 1, 1, "ASC")
    '
    Call HEADERS("REIVACOM", "")
    Call HEADERS("REIVACOM", "Corresponde a: " + TRIM$(CORRESP$))
    '

99  Screen.MousePointer = 1
    Command30.Enabled = True
    Exit Sub
    '
    '

INCLUASDIAR:
    FEMOSUDIA% = CVINT(FECHAD)
    For KUI% = 1 To CANASIX%
      If FECHASI%(KUI%) > 0 Then
        If FECHASI%(KUI%) < FEMOSUDIA% Then
          NUMEFA$ = TRIM$(NUMEASI$(KUI%))
          NCLIE% = 0
          For IXX% = 1 To CACUEIVA%
            ICUEN#(IXX%) = 0
          Next IXX%
          '
          SQLX$ = "SELECT * FROM DetaAsien "
          SQLX$ = SQLX$ + " WHERE CDBL(Fechasien) >= " & DESDAT
          SQLX$ = SQLX$ + " AND CDBL(Fechasien) <= " & HASDAT
          SQLX$ = SQLX$ + " AND statpase >0  "
          SQLX$ = SQLX$ + " AND NuIdAsien = '" & NUMEFA$ & "' "
          SQLX$ = SQLX$ + "ORDER BY OrNuPase "
          'Set DETACONTA = Contable.OpenRecordset(SQLX$, 4)
          Dim DETACONTA As ADODB.Recordset
          Set DETACONTA = New ADODB.Recordset
28        On Error Resume Next
          DETACONTA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 28
          End If
          On Error GoTo 0
           '
          With DETACONTA
            On Error Resume Next
            .MoveFirst
            If Err < 1 Then
              On Error GoTo 0
              Do While Not .EOF
                icuexi% = !CUECOINT
                For IXX% = 1 To CACUEIVA%
                  If icuexi% = CUEIVA%(IXX%) Then
                    ICUEN#(IXX%) = ICUEN#(IXX%) + !IDebePase - !IHabePase
                    REPAS$ = !RefePase
                    GoTo 192
                  End If
                Next IXX%
                If NCLIE% < 1 Then
                  Call APERBASDAT("PROVE")
'                  Proved12.FindFirst "CuMa_Cli = " & icuexi%
'                  If Proved12.NoMatch = False Then
                   SQLX$ = "SELECT * FROM Proved12 "
                   SQLX$ = SQLX$ + " WHERE CuMa_Cli = " & icuexi% & ""
                   Dim PROVVVED12 As ADODB.Recordset
                   Set PROVVVED12 = FLEXCONN.Execute(FILTSQL$(SQLX$))
                   With PROVVVED12
                    On Error Resume Next
                    If Not (.EOF And .BOF) Then
                      On Error GoTo 0
                      NCLIE% = !Nume_Cli
                    End If
                    On Error GoTo 0
                   End With
                End If
192           .MoveNext
              Loop
            End If
          End With
          PROVVVED12.Close
          Set PROVVVED12 = Nothing
          DETACONTA.Close
          Set DETACONTA = Nothing
          '
          ZZ$ = REBLA$
          Call REPLA(ZZ$, MKI$(FECHASI%(KUI%)), 3, 2)
          Call REPLA(ZZ$, MKI$(FECHASI(KUI%)), 5, 2)
          Call REPLA(ZZ$, MKS$(0), 7, 4)
          Call REPLA(ZZ$, NUMEFA$, 11, 18)
          Call REPLA(ZZ$, MKI$(NCLIE%), 31, 2)
          Call REPLA(ZZ$, RASOCLI$((-1) * NCLIE%), 33, 30)
          Call REPLA(ZZ$, CUITCLI$((-1) * NCLIE%), 63, 16)
          Call REPLA(ZZ$, POSIVAC$((-1) * NCLIE%), 79, 12)
          '
          If Abs(NCLIE%) < 1 Then
            Call REPLA(ZZ$, REPAS$, 33, 30)
            Call REPLA(ZZ$, "N/Registrado", 63, 16)
            Call REPLA(ZZ$, "N/Def", 79, 12)
          End If
          '
          PROIVA% = 0: ULIVA% = 0
          RELIVA$(0) = ZZ$
          For KU% = 1 To 32
            RELIVA$(KU%) = REBLA$
          Next KU%
          '
          ITOTAX# = 0
          For IXX% = 1 To CACUEIVA%
            If Abs(ICUEN#(IXX%)) >= 0.005 Then
              ZZ$ = RELIVA$(PROIVA%)
              VAMO$ = TASIVANOM$(IXX%)
              If TRIM$(UCase$(VAMO$)) = "PER-IVA" Then
                   ZZ$ = RELIVA$(0)
                   ITOTAA# = ICUEN#(IXX%) + CVD(Mid$(ZZ$, 161, 8))
                   Call REPLA(ZZ$, MKD$(ITOTAA#), 161, 8)
                   ITOTAX# = ITOTAX# + ICUEN#(IXX%)
                   RELIVA$(0) = ZZ$
                   GoTo 199
                 ElseIf TRIM$(UCase$(VAMO$)) = "PER-GAN" Then
                   ZZ$ = RELIVA$(0)
                   ITOTAA# = ICUEN#(IXX%) + CVD(Mid$(ZZ$, 169, 8))
                   Call REPLA(ZZ$, MKD$(ITOTAA#), 169, 8)
                   ITOTAX# = ITOTAX# + ICUEN#(IXX%)
                   RELIVA$(0) = ZZ$
                   GoTo 199
              End If
              BASE# = 0
              If XVALO(TASIVANOM$(IXX%)) > 0.5 Then
                BASE# = 100 * ICUEN#(IXX%) / XVALO(TASIVANOM$(IXX%))
              End If
              Call REPLA(ZZ$, VAMO$, 129, 8)
              Call REPLA(ZZ$, MKD$(BASE#), 137, 8)
              Call REPLA(ZZ$, MKD$(ICUEN#(IXX%)), 145, 8)
              ITOTAX# = ITOTAX# + BASE# + ICUEN#(IXX%)
              If XVALO(TASIVANOM$(IXX%)) > 0.5 Then
                Call REPLA(ZZ$, MKD$(XVALO(TASIVANOM$(IXX%))), 153, 8)
              End If
              RELIVA$(PROIVA%) = ZZ$
              ULIVA% = PROIVA%
              PROIVA% = PROIVA% + 1
              '
            End If
199       Next IXX%
          '
          FECHASI%(KUI%) = 0
          ZZ$ = RELIVA$(0)
          Call REPLA(ZZ$, MKD$(ITOTAX#), 91, 8)
          RELIVA$(0) = ZZ$
          '
          For KU% = 0 To ULIVA%
            JJ& = JJ& + 1
            Call GRAREG("LIVACOM", RELIVA$(KU%), JJ&)
            Call GRAREG("LIVACCAI", RELIVA$(KU%), JJ&)
          Next KU%
          '
        End If
      End If
    Next KUI%
    
Return

'
   

   Command30.Enabled = True

End Sub


Private Sub FEVEN_DblClick(Index As Integer)
   Call BOTFEVEN_Click(Index)
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
End Sub


Private Sub imp_cre_Click()
   Call FACOCOMP07.Command3_Click
End Sub

Private Sub IMPOGRA_Change(Index As Integer)
   '
    TASIX$ = TRIM$(REPLACAR$(TASIVA(Index).Caption, "%", ""))
    TASI = XVALO(TASIX$)
    IGRA# = XVALO(IMPOGRA(Index).TEXT)
    If Index < 4 Then
        IMPOIVA(Index).TEXT = CSTRING$(MKD$(TASI * IGRA# / 100), 4, 9, 2, 2)
        IMPOIVA(Index).Refresh
    End If
    '
    TOGRAVA# = 0
    For KKI% = 0 To 6
        IGRA# = XVALO(IMPOGRA(KKI%).TEXT)
        TOGRAVA# = TOGRAVA# + IGRA#
    Next KKI%
    '
    Label45 = FORMATNUM$(TOGRAVA#, "F12.2")
    Call ACTUALIZAR
'    TOBRUFAC# = 0
'    For KKI% = 0 To 6
'        TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
'        TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
'    Next KKI%
'    TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text7.TEXT) + 0.5)) / 100
'    TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text8.TEXT) + 0.5)) / 100
'    TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text12.TEXT) + 0.5)) / 100
'
'    IORIG# = TOBRUFAC#
'    Label43 = IORIG#
    End Sub
   '
Sub ACTUALIZAR()
       TOBRUFAC# = 0
    For KKI% = 0 To 6
        TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
        TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
    Next KKI%
    TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text7.TEXT) + 0.5)) / 100
    TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text8.TEXT) + 0.5)) / 100
    TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text12.TEXT) + 0.5)) / 100
     
    IORIG# = TOBRUFAC#
    Label43 = IORIG#
    
End Sub

Private Sub IMPOGRA_GotFocus(Index As Integer)
    IMPOGRA(Index).SelStart = 0
    IMPOGRA(Index).SelLength = Len(IMPOGRA(Index).TEXT)
End Sub

Private Sub IMPOGRA_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii% = 13 Then
      If Index < 6 Then
         IMPOGRA(Index + 1).SetFocus
        Else
         IMPOIVA(0).SetFocus
      End If
   End If
End Sub

Private Sub IMPOGRA_LostFocus(Index As Integer)
   IGRA# = Abs(CDbl(XVALO(IMPOGRA(Index).TEXT)))
   TASIX$ = TRIM$(TASIVA(Index).Caption)
   If TASIX$ = "" Then IGRA# = 0
'   If LETRAFA$ <> "A" And LETRAFA$ <> "M" Then
'      If Index <> 4 And Index <> 5 Then
'        IGRA# = 0
'      End If
'   End If
   IMPOGRA(Index).TEXT = CSTRING$(MKD$(IGRA#), 4, 11, 2, 2)
   IMPOGRA(Index).Refresh
End Sub

Private Sub IMPOIVA_Change(Index As Integer)
'   Call CALTOFAC
End Sub

Private Sub IMPOIVA_GotFocus(Index As Integer)
    If Index = 4 Then
       IMPOIVA(5).SetFocus
       Exit Sub
    End If
    IMPOIVA(Index).SelStart = 0
    IMPOIVA(Index).SelLength = Len(IMPOIVA(Index).TEXT)
End Sub

Private Sub IMPOIVA_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii% = 13 Then
      If Index < 6 Then
         NECA% = 1: If Index = 3 Then NECA% = 3
         IMPOIVA(Index + NECA%).SetFocus
        Else
         Text7.SetFocus
      End If
   End If
End Sub

Private Sub IMPOIVA_LostFocus(Index As Integer)
'   IGRA# = Abs(CDbl(XVALO(IMPOIVA(Index).TEXT)))
'   TASIX$ = TRIM$(TASIVA(Index).Caption)
'   If TASIX$ = "" Then IGRA# = 0
'   If LETRAFA$ <> "A" And LETRAFA$ <> "M" Then
'     IGRA# = 0
'   End If
'   IMPOIVA(Index).TEXT = CSTRING$(MKD$(IGRA#), 4, 9, 2, 2)
'   IMPOIVA(Index).Refresh
End Sub

Private Sub ingbru_Click()
   Call OPFACOM07.Command5_Click
End Sub

Private Sub IVENCI_Change(Index As Integer)
    IMPTOTT = 0
    For I& = 0 To 9
        IMPTOTT = IMPTOTT + XVALO(IVENCI(I&))
    Next I&
    Label43 = TRIM$(FORMATNUM(IMPTOTT, "F12.2"))
End Sub

Private Sub Label17_DblClick()
   '
   If ULTREG("PROVINB") < 1 Then
      Call COMUNI("Falta Configurar las Percepciones de Ingresos Brutos")
      Exit Sub
   End If
   '
   Call TRALOCAL("PROVINB", "")
   '
   VTB% = VENTABU%("INGREPIB")
   If VTB% >= 0 Then
     SUPIBRU# = 0
     For U& = 1 To ULTREG&("!PROVINB")
       XX$ = REGLEIDO$("!PROVINB", U&)
       SUPIBRU# = SUPIBRU# + CVD(Mid$(XX$, 57, 8))
     Next U&
     '
     If SUPIBRU# >= 0.005 Then
        TEXT10 = "*"
        Label18 = "Jurisdicciones Varias"
        Text12 = FORMATNUM$(SUPIBRU#, "F9.2")
        TEXT10.Enabled = False
        Text12.Enabled = False
        Command10.Enabled = False
        Call TRACENTRAL("PROVINB", "")
      Else
        If TEXT10 = "*" Then
          TEXT10 = ""
          Label18 = ""
          Text12 = ""
        End If
        TEXT10.Enabled = True
        Text12.Enabled = True
        Command10.Enabled = True
     End If
   End If
   '
End Sub

Private Sub Label20_Change()
   LETRAFA$ = Label20.Caption
   If TIDOCOM$ <> "CO" Then
      LETRAFA$ = Label20.Caption
      TTXX$ = Text11.TEXT
      If TRIM$(Text11.TEXT) <> "" Then
        Call REPLA(TTXX$, LETRAFA$, 4, 1)
        Text11.TEXT = TTXX$
        Text11.Refresh
      End If
   End If
   Call SIGUEFACO3
End Sub

Private Sub Label20_Click()
   LETRAX$ = Label20
   If LETRAX$ = "A" Then Label20 = "B"
   If LETRAX$ = "B" Then Label20 = "C"
   If LETRAX$ = "C" Then Label20 = "E"
   If LETRAX$ = "E" Then Label20 = "M"
   If LETRAX$ = "M" Then Label20 = "A"
   'ILETRA% = Asc(Label20.Caption + "A") + 1
   'If ILETRA% > 67 Then ILETRA% = 65
   'Label20.Caption = Chr$(ILETRA%)
End Sub

Private Sub Label25_Change()
   Label21 = ECOARCH$("MEPAGO", MKI$(XVALO(Label25)))
   If TRIM$(Label25) = "" Then
       Text14 = ""
       Text15 = ""
       Call CALTOFAC
     Else
       HOII% = HOY(HOS$)
       Text15 = HOS$
       On Error Resume Next
       Text14.SetFocus
       On Error GoTo 0
   End If
End Sub

Private Sub Label25_DblClick()
   Call SELECHO("MEPAGO")
   If VALACT1$("MEPAGO") <> "" Then
     TPV% = XVALO(VALACT1$("MEPAGO"))
     If CUEMEPAG%(TPV%) = CUEICARCH% Then
         Call MENSERR(24, "Cartera Valores no Disponible\para Compras de Contado.\   \Debe Cancelar por Orden de Pago.")
         Label25 = ""
       Else
         Label25 = VALACT1$("MEPAGO")
     End If
   End If
End Sub

Private Sub Label29_DblClick()
   '
   If Check2.Value > 0 Then
      Call MENSERR(24, "Carga por Detalle Solo en Moneda Básica.")
      Exit Sub
   End If
   '
   MONEMIA% = MONEMI%
15 MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   If DMONE$(MONEMI%) = "" Then GoTo 15
   Label29 = DMONE$(MONEMI%)
   Label29.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      Label29.ForeColor = RGB(255, 0, 0)
   End If
   '
   Call RECALTO(TCMONE#(MONEMI%), TCMONE#(MONEMIA%))
   Text18 = FORMATNUM(TCMONE#(MONEMI%), "F11.4")
'   Label29.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TCMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
'   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBX# = TICAMONE(MONEMI%)
   '
   TOTIMPU# = 0
   For U& = 1 To List1.ListCount
     Z$ = List1.List(U& - 1)
     IPA# = XVALO(Mid$(Z$, 40, 12))
     PREDOL# = IPA# * COEFCON
     TOTIMPU# = TOTIMPU# + PREDOL#
     Call REPLA(Z$, FORMATNUM$(PREDOL#, "F12.2"), 40, 12)
     '
     List1.RemoveItem U& - 1
     List1.AddItem Z$, U& - 1
   Next U&
   '
   For KKI% = 1 To 6
     IMPOGRA(KKI%) = Space$(11)
     IMPOIVA(KKI%) = Space$(8)
   Next KKI%
   '
'   NC1% = (-1) * XVALO(Text1)
'   PVCL% = PIVACLI%(NC1%)
'   If (LETRAFA$ = "A" Or LETRAFA$ = "M") And PVCL% <> 4 Then  ' CASO DE PROVEEDORES EXENTOS
'        ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
'        ITOGRA.Refresh
'      Else
'        ITOGRA.Caption = Space$(11)
'        IMPOGRA(4).TEXT = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
'        IMPOGRA(4).Refresh
'   End If
'   List1.Refresh
   '
   IMPOGRA(0).SetFocus
'   If TOTIMPU# < 0.01 Then
'      Text1.SetFocus
'   End If
'   '
End Sub

Sub RECALTO(TCNUE, TCVIE)
   '
   Label29.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TCNUE, "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   COEFCON = TCVIE / TCNUE
   TICAMBX# = TCNUE
   '
   TOTIMPU# = 0
   For U& = 1 To List1.ListCount
     Z$ = List1.List(U& - 1)
     IPA# = XVALO(Mid$(Z$, 40, 12))
     PREDOL# = IPA# * COEFCON
     TOTIMPU# = TOTIMPU# + PREDOL#
     Call REPLA(Z$, FORMATNUM$(PREDOL#, "F12.2"), 40, 12)
     '
     List1.RemoveItem U& - 1
     List1.AddItem Z$, U& - 1
   Next U&
   '
   For KKI% = 1 To 6
     IMPOGRA(KKI%) = Space$(11)
     IMPOIVA(KKI%) = Space$(8)
   Next KKI%
   '
   Label36 = FORMATNUM$(TOTIMPU#, "F11.2")
'   NC1% = (-1) * XVALO(Text1)
'   PVCL% = PIVACLI%(NC1%)
'   If (LETRAFA$ = "A" Or LETRAFA$ = "M") And PVCL% <> 4 Then  ' CASO DE PROVEEDORES EXENTOS
'        ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
'        ITOGRA.Refresh
'      Else
'        ITOGRA.Caption = Space$(11)
'        IMPOGRA(4).TEXT = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
'        IMPOGRA(4).Refresh
'   End If
   List1.Refresh
   '
   IMPOGRA(1).SetFocus
   If TOTIMPU# < 0.01 Then
      Text1.SetFocus
   End If
   '
End Sub

Private Sub Label31_Change()
  '
  For KKII% = 1 To 6
     IMPOGRA(KKII%) = ""
  Next KKII%
  '
  IMPONET1# = XVALO(Label31)
  LETRAFA$ = Label20
'  If LETRAFA$ = "A" Then
'     ITOGRA = FORMATNUM(IMPONET1#, "F10.2")
'    Else
'     IMPOGRA(4) = FORMATNUM(IMPONET1#, "F10.2")
'  End If
  TOTIMPU# = IMPONET1#
  '
End Sub

Private Sub Label36_Change()
  '
  Call Text19_Change
  '
End Sub

Private Sub Label40_DblClick()
   If DERACCE%("BAJNICLA/NOMESS", "Reducir Nivel de Clasificación") >= 2 Then
      NICLAC% = XVALO(Mid$(Label40, 2))
      If NICLAC% > 0 Then
         Label40 = "<" + TRIM$(Str$(NICLAC% - 1))
         Label40.Refresh
      End If
   End If
End Sub

Private Sub Label6_DblClick()
   If List1.ListCount > 0 Then
      NICON% = NIVCONFI%
      If NICON% > 0 Then
         Label40 = "<" + TRIM$(Str$(NICON%))
         Picture4.Visible = True
      End If
   End If
End Sub

'
Private Sub maes_pro_Click()
   Call Command11_Click
End Sub

Private Sub notas_pro_Click()
   Call Command3_Click
End Sub

Private Sub nva_fac_Click()
   Text1.SetFocus
End Sub

Private Sub print_Click()
   Call Command5_Click
End Sub

Private Sub tasa_iva_Click()
   Call Command8_Click
End Sub

Private Sub Text1_LostFocus()
   CLITE$ = TRIM$(Text1.TEXT)
   If CLITE$ <> "" Then
     If XVALO(CLITE$) < 1 Or XVALO(CLITE$) > 32767 Then
       Text1.TEXT = Space$(5)
       Exit Sub
     End If
     If ECOARCH$("ACREDOR", MKI$(XVALO(CLITE$))) = "" Then
       Text1.TEXT = Space$(5)
       Exit Sub
     End If
   End If
End Sub

Private Sub Text10_LostFocus()
   If TRIM$(TEXT10.TEXT) <> "" Then
     If TRIM$(TEXT10.TEXT) <> "*" Then
       If ECOARCH$("PROVINC", TRIM$(TEXT10.TEXT)) = "" Then
         TEXT10.TEXT = " "
         On Error Resume Next
         TEXT10.SetFocus
         On Error GoTo 0
       End If
     End If
   End If
End Sub
'
Private Sub Text10_GotFocus()
    Text12.SelStart = 0
    Text12.SelLength = Len(TEXT10.TEXT)
End Sub

Private Sub Text12_Change()
   'Call CALTOFAC
   Call ACTUALIZAR
End Sub

'
Private Sub Text12_GotFocus()
    Text12.SelStart = 0
    Text12.SelLength = Len(Text12.TEXT)
End Sub

Private Sub TEXT12_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text20.SetFocus
   End If
End Sub

Private Sub TEXT12_LostFocus()
'   IGRA# = Abs(CDbl(XVALO(Text12.TEXT)))
'   If LETRAFA$ <> "A" And LETRAFA$ <> "M" Then
'     IGRA# = 0
'   End If
'   DPX1$ = TRIM$(TEXT10.TEXT)
'   If IGRA# >= 0.005 Then
'     If DPX1$ = "" Or ECOARCH$("PROVINC", DPX1$) = "" Then
'       Call MENSERR(24, "Falta Jurisdicción Percepción Ingresos Brutos")
'       IGRA# = 0
'     End If
'   End If
'   Text12.TEXT = CSTRING$(MKD$(IGRA#), 4, 9, 2, 2)
'   Text12.Refresh
End Sub

Private Sub Text13_GotFocus()
   If XVALO(Text1.TEXT) > 0 Then
      Text1.SetFocus
      Exit Sub
   End If

   Text13.SelStart = 0
   Text13.SelLength = Len(Text13.TEXT)
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       On Error Resume Next
       Text11.SetFocus
       On Error GoTo 0
       '
    End If
    '
End Sub

Private Sub Text15_DblClick()
    '\\\OT-06-0570-RG-20/12/06///
    Call Command13_Click
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text16_GotFocus()
    Text16.SelStart = 0
    Text16.SelLength = Len(Text16)
End Sub

Private Sub Text17_GotFocus()
    Text17.SelStart = 0
    Text17.SelLength = Len(Text17)
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Text17.SetFocus
    End If
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Text6.SetFocus
    End If
End Sub

Private Sub Text17_LostFocus()
    If TRIM$(Text17) <> "" Then
      Text17 = FECHATEX$(FECHANUM(Text17))
      Text17.Refresh
    End If
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call Text18_lostfocus
   End If
End Sub

Private Sub Text18_lostfocus()
  '
  TC1# = XVALO(Text18)
  If MONEMI% <= 1 Then TC1# = 1
  If TC1# < 0.1 Then
    TC1# = TICAMONE(MONEMI%)
  End If
  Text18.TEXT = FORMATNUM$(TC1#, "F11.4")
  '
  TCMONE#(MONEMI%) = XVALO(Text18)
  Label29.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TCMONE#(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
  '
End Sub

Private Sub Text19_Change()
  '
  TN1# = Val(Label36)
  BN1 = COEFDESCU(TRIM$(Text19))
  IMTO1# = (Int(100 * TN1# * (1 - BN1 / 100) + 0.5)) / 100
  BN2# = TN1# - IMTO1#
  '
  Label35 = FORMATNUM(BN2#, "F11.2")
  Label31 = FORMATNUM(IMTO1#, "F11.2")
  '
End Sub

Private Sub Text19_GotFocus()
   If Check2.Value > 0 Then
      Text20.SetFocus
   End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       On Error Resume Next
       Text3.SetFocus
       On Error GoTo 0
       '
    End If
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      FEVEN(0).SetFocus
   End If
End Sub

Private Sub Text22_Change()
   Call CALTOFAC
End Sub

Private Sub Text22_GotFocus()
    Text22.SelStart = 0
    Text22.SelLength = Len(Text22.TEXT)

End Sub

Private Sub Text22_LostFocus()
   REDON# = CDbl(XVALO(Text22.TEXT))
   Text22.TEXT = CSTRING$(MKD$(REDON#), 4, 7, 2, 2)
   Text22.Refresh

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       On Error Resume Next
       Text4.SetFocus
       On Error GoTo 0
       '
    End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       On Error Resume Next
       Text5.SetFocus
       On Error GoTo 0
       '
    End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       On Error Resume Next
       Text13.SetFocus
       On Error GoTo 0
       '
    End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text9.SetFocus
   End If
End Sub

Private Sub Text7_Change()
   'Call CALTOFAC
   Call ACTUALIZAR
   
End Sub

'
Private Sub Text7_GotFocus()
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7.TEXT)
End Sub

Private Sub TEXT7_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text8.SetFocus
   End If
End Sub

Private Sub TEXT7_LostFocus()
'   IGRA# = Abs(CDbl(XVALO(Text7.TEXT)))
'   If LETRAFA$ <> "A" And LETRAFA$ <> "M" Then
'     IGRA# = 0
'   End If
'   Text7.TEXT = CSTRING$(MKD$(IGRA#), 4, 9, 2, 2)
'   Text7.Refresh
End Sub

Private Sub Text8_Change()
   'Call CALTOFAC
   Call ACTUALIZAR
End Sub

'
Private Sub Text8_GotFocus()
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.TEXT)
End Sub

Private Sub TEXT8_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      TEXT10.SetFocus
   End If
End Sub

Private Sub TEXT8_LostFocus()
'   IGRA# = Abs(CDbl(XVALO(Text8.TEXT)))
'   If LETRAFA$ <> "A" And LETRAFA$ <> "M" Then
'     IGRA# = 0
'   End If
'   Text8.TEXT = CSTRING$(MKD$(IGRA#), 4, 9, 2, 2)
'   Text8.Refresh
End Sub
'
Private Sub ITOGRA_Change()
   '
'   TASIX$ = TRIM$(REPLACAR$(TASIVA(0).Caption, "%", ""))
'   TASI = XVALO(TASIX$)
'   IGRA# = XVALO(ITOGRA.TEXT)
'   IMPOIVA(0).TEXT = CSTRING$(MKD$(TASI * IGRA# / 100), 4, 9, 2, 2)
'   IMPOIVA(0).Refresh
'   Call CALTOFAC
   '
End Sub
'
Private Sub Text10_DblClick()
    Call SELECHO("PROVINC")
    DPX1$ = VALACT1$("PROVINC")
    DP1$ = DPX1$
    TEXT10.TEXT = AJUSTI$(DP1$, 1)
    Label18.Caption = ECOARCH$("PROVINC", TRIM$(DP1$))
End Sub

Private Sub Text10_Change()
    DPX1$ = TRIM$(TEXT10.TEXT)
    DP1$ = AJUSTI$(DPX1$, 1)
    Label18.Caption = ECOARCH$("PROVINC", DP1$)
End Sub
'
Private Sub TEXT10_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text12.SetFocus
   End If
End Sub


Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCHIG07/AMAPRO", "Archivos Maestros")
    Command11.Enabled = True
End Sub

Private Sub Command2_Click()
    If MODIFIX% > 0 Then
       If COMALTER%("Abandona sin Guardar los Cambios ?\\No, Continuar\Si, Cancelar") <> 2 Then
          Exit Sub
       End If
    End If
    Screen.MousePointer = 11
    If List1.ListCount > 0 Or List2.ListCount > 0 Then
          Call CIERRARCH("*.*")
          Call BLAFACOM
          Screen.MousePointer = 1
          Text1.SetFocus
        Else
          Screen.MousePointer = 1
          Hide
    End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If List1.ListCount > 0 Or List2.ListCount > 0 Then
        Command2.ToolTipText = "Limpia y Refresca Pantalla"
      Else
        Command2.ToolTipText = "Cierra y Abandona la Aplicación"
   End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    Call CONECRUN("*NOTAPRO/" + NCLI$, "Notas de Proveedores")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEP07.Text1 = Text1
    PROLLAMA$ = Name
    'Hide
    RECUEP07.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    PRINTFORM
    Text1.SetFocus
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    FACOCOMP07.Show 1
    Command6.Enabled = True
End Sub


Private Sub Command7_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text6.TEXT = VDEVV$
    End If
End Sub

Sub Command8_Click()
   Command8.Enabled = False
   If DERACCE%("STPFACP", "Configuración de Funcionamiento de Facturas de Proveedores") > 1 Then
     OPFACOM07.Show 1
   End If
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub FEVEN_GotFocus(Index As Integer)
    FEVEN(Index).SelStart = 0
    FEVEN(Index).SelLength = Len(FEVEN(Index).TEXT)
End Sub

Private Sub FEVEN_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If TRIM$(FEVEN(Index).TEXT) = "" Then
         If TRIM$(IVENCI(Index).TEXT) = "" Then
            BOTNEXT.SetFocus
            Exit Sub
         End If
      End If
      IVENCI(Index).SetFocus
   End If
End Sub

Private Sub FEVEN_LostFocus(Index As Integer)
   If TRIM(FEVEN(Index).TEXT) <> "" Then
      If FECHANUM(FEVEN(Index).TEXT) < 1 Then
         Call MENSERR(24, "Fecha no Válida")
         FEVEN(Index).SetFocus
      End If
   End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call CENTRAFORM(Me)
    Refresh
    DoEvents
    '
    VScroll1.Max = (Frame15.Height - VScroll1.Height) / 100
    VScroll1.Value = 0
    '
    For IXX% = 0 To 6
      TXXX$ = TRIM$(Str$(IXX%))
      If IXX% <= 3 Then
        TASA = XVALO(Control$("", "TASIVA-" + TXXX$))
        TASIVA(IXX%).Caption = "0.00%"
        If TASA > 0.005 Then
          TASIVA(IXX%).Caption = TRIM$(CSTRING$(MKS$(TASA), 3, 6, 2, 2) + "%")
        End If
      End If
      '
      CUEM$ = Control$("", "CUEIVA-" + TXXX$)
      If CUEM$ = "" Or CUEINT%(CUEM$) < 1 Then
        Call MENSERR(24, "Falta Definir Cuentas Contables\Correspondientes a IVA Crédito Fiscal.")
        OPFACOM07.Show 1
        Hide
      End If
    Next IXX%
    '
    TIDOCOM$ = "FC"
    Text1.TEXT = Space$(5)
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.TEXT)
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    Text6.TEXT = HO$
    Text11.TEXT = ""
    '
    Call BLAFACOM
    '
    CAMON% = ULTREG&("TAMONED")
    If CAMON% < 1 Then
        CAMON% = 1
        DMONE$(1) = "Pesos"
        TCMONE#(1) = 1
      Else
        For U& = 1 To CAMON%
          XX$ = REGLEIDO$("TAMONED", U&)
          DMX$ = TRIM$(Left$(XX$, 8)) + " "
          PPXX% = InStr(DMX$, " ")
          DMX$ = Left$(DMX$, PPXX% - 1)
          '
          VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
          DMONE$(U&) = TRIM$(DMX$)
          TCMONE#(U&) = VHX
        Next U&
    End If
    '
    MONEMI% = 1
    Label29 = DMONE$(MONEMI%)
    Text18 = FORMATNUM(TCMONE#(MONEMI%), "F11.4")
    '
    If NIVCONFI% > 0 Then
       Label6.ToolTipText = "Doble Click para Clasificar"
    End If
    '
    Refresh
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    Screen.MousePointer = 1
    '
End Sub

Private Sub IVENCI_GotFocus(Index As Integer)
    If FECHANUM(FEVEN(Index).TEXT) < 1 Then
       'FEVEN(Index).SetFocus
       'Call MENSERR(24, "Fecha de Vencimiento no Válida")
       Exit Sub
    End If
    '
    IVENCI(Index).SelStart = 0
    IVENCI(Index).SelLength = Len(IVENCI(Index).TEXT)
End Sub

Private Sub IVENCI_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index < 8 Then
         FEVEN(Index + 1).SetFocus
      End If
   End If
End Sub

Private Sub IVENCI_LOSTFOCUS(Index As Integer)
   SUMAVEN# = XVALO(Label10.Caption)
   For KKI% = 0 To Index - 1
      SUMAVEN# = SUMAVEN# + XVALO(IVENCI(KKI%).TEXT)
   Next KKI%
   IPAVEN# = XVALO(IVENCI(Index).TEXT)
   If SUMAVEN# + IPAVEN# < XVALO(TOBRUFAC#) Then
      If Index < 8 Then
         IVENCI(Index + 1) = CSTRING$(MKD$(TOBRUFAC# - SUMAVEN# - IPAVEN#), 3, 11, 2, 2)
      End If
   End If
   IVENCI(Index).TEXT = CSTRING$(MKD$(IPAVEN#), 4, 11, 2, 2)
End Sub

Private Sub Label10_Change()
   Call CALTOFAC
End Sub

Private Sub List1_DblClick()
    '
    If Check2 Then
        FOFACOBR07.Show 1
      Else
        Call SIGUEFACO2
    End If
    '
End Sub

Private Sub List1_GotFocus()
   MODIFIX% = 1
End Sub

Private Sub List2_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    Z$ = List2.TEXT
    LIST2ANT$ = Z$
    TOLISTA& = List2.TopIndex
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = " " + Mid$(Z$, 56, 13)
    '
    SUMACRE# = 0: MAXACRE# = TOBRUFAC#
    List2.Visible = False
    VACON% = 1
    For U& = 1 To List2.ListCount
       List2.ListIndex = U& - 1
       SUMACRE# = SUMACRE# + XVALO(Mid$(List2.TEXT, 70, 14))
    Next U&
    List2.TopIndex = TOLISTA&
    List2.ListIndex = ILISTA&
    List2.Visible = True
    VACON% = 0
    '
    If IACRE$ = "" Then
         ACREDIT# = XVALO(IPEN$)
         If ACREDIT# > MAXACRE# - SUMACRE# Then
            ACREDIT# = MAXACRE# - SUMACRE#
            If ACREDIT# < 0 Then ACREDIT# = 0
         End If
         ACRE$ = CSTRING$(MKD$(ACREDIT#), 3, 14, 2, 2)
         Call REPLA(Z$, ACRE$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(XVALO(ACRE$))
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
    End If
    VACON% = 1
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = ILISTA&
    VACON% = 0
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 3, 14, 2, 2)
    Label10.Refresh
    '
End Sub
'

Private Sub List2_DblClick()
    '
    VACON% = 1
    ILIST2& = List2.ListIndex
    TOLIST2& = List2.TopIndex
    '
    Z$ = List2.TEXT
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = Mid$(Z$, 56, 14)
    SUMACRE# = XVALO(Label10.Caption)
    SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(LIST2ANT$, 70, 14)))
    List2.Visible = False
    List2.AddItem LIST2ANT$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    '
    Call SIGUEFACO5
    '
    VACON% = 1
    List2.TopIndex = TOLIST2&
    List2.Visible = True
    If List2.ListCount > ILIST2& Then
          List2.ListIndex = ILIST2&
        Else
          Call CIERRARCH("*.*")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Text1.SetFocus
    End If
    VACON% = 0
    '
End Sub

Private Sub Option1_Click(Index As Integer)
   For KKI% = 0 To 7
     Option1(KKI%).FontSize = 8
     Option1(KKI%).FontBold = False
     If Option1(KKI%).Value = True Then
       Option1(KKI%).FontSize = 10
       Option1(KKI%).FontBold = True
     End If
   Next KKI%
   '
   Check2 = 0
   If Option1(0) = True Then
        Label6.Caption = "Factura Nro:"
        TIDOCOM$ = "FC"
      ElseIf Option1(1) = True Then
        Label6.Caption = "N.Débito Nro:"
        TIDOCOM$ = "ND"
      ElseIf Option1(2) = True Then
        Label6.Caption = "N.Crédito Nro:"
        TIDOCOM$ = "NC"
      ElseIf Option1(3) = True Then
        Label6.Caption = "Ticket-Factura:"
        TIDOCOM$ = "TF"
      ElseIf Option1(4) = True Then
        Label6.Caption = "R.Honorarios:"
        TIDOCOM$ = "RH"
      ElseIf Option1(5) = True Then
        Label6.Caption = "Comprobante:"
        TIDOCOM$ = "CO"
      ElseIf Option1(6) = True Then
        Label6.Caption = "Crédito Propio:"
        TIDOCOM$ = "CP"
      ElseIf Option1(7) = True Then
        Label6.Caption = "Débito Propio:"
        TIDOCOM$ = "DP"
   End If
   Label6.Refresh
   Call Check2_LOSTFOCUS
   Call SIGUEFACO1
End Sub

Private Sub Text1_Change()
    '
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Check2.Enabled = True
    If TRIM(Text1) <> "" Then Check2.Enabled = False
    Call BLAFACOM
    '
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text13 = ""
    Text16 = ""
    Text17 = ""
    '
    If XVALO(Text1.TEXT) > 32767 Then
      Text1.TEXT = Left$(Text1.TEXT, Len(Text1.TEXT) - 1)
      Text1.SelStart = Len(Text1.TEXT)
    End If
    '
End Sub

Private Sub Text1_DblClick()
    '
    If MODIFIX% > 0 Then
       If COMALTER%("Abandona sin Guardar los Cambios ?\\No, Continuar\Si, Cancelar") <> 2 Then
          Exit Sub
       End If
    End If
    '
    Call SELECHO("ACREDOR")
    NCXI% = XVALO(TRIM$(VALACT1$("ACREDOR")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        Text13.TEXT = CUITCLI$((-1) * NC%)
          If LETRAFA$ <> "M" Then LETRAFA$ = "A"
          PVCL% = PIVACLI%(-1 * NC%)
          If PVCL% = 0 Or PVCL% = 2 Or PVCL% = 6 Or PVCL% = 4 Then LETRAFA$ = "C"
          If PVCL% = 5 Then LETRAFA$ = "E"
        Label20 = LETRAFA$
        If Left$(Text2.TEXT, 6) = "VARIOS" Then
            Text13.SetFocus
          Else
            Text6.SetFocus
        End If
        'Call SIGUEFACO1
    End If
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
        Refresh
        CTXX% = 1
    End If
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.TEXT)
    Screen.MousePointer = 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If XVALO(Text1.TEXT) > 0 Then
          If XVALO(Text1.TEXT) <= 32767 Then
             NC% = XVALO(Text1.TEXT)
             If ECOARCH$("ACREDOR", MKI$(NC%)) <> "" Then
                Text2.TEXT = RASOCLI$((-1) * NC%)
                Text3.TEXT = DOMICLI$((-1) * NC%)
                Text4.TEXT = CPOSCLI$((-1) * NC%)
                Text5.TEXT = LOCACLI$((-1) * NC%)
                Text13.TEXT = CUITCLI$((-1) * NC%)
                Label26 = CODICLI((-1) * NC%)
                If LETRAFA$ <> "M" Then LETRAFA$ = "A"
                PVCL% = PIVACLI%(-1 * NC%)
                If PVCL% = 0 Or PVCL% = 2 Or PVCL% = 6 Or PVCL% = 4 Then LETRAFA$ = "C"
                If PVCL% = 5 Then LETRAFA$ = "E"
                Label20 = LETRAFA$
                If Left$(Text2.TEXT, 6) = "VARIOS" Then
                    Text13.SetFocus
                  Else
                    Call BUSCACAI(NC%)
                    If PVCL% = 0 Or PVCL% = 5 Or PVCL% = 6 Then
                        Text6.SetFocus     ' FECHA DE EMISION
                      Else
                        Text16.SetFocus    ' NUMERO DE CAI
                    End If
                End If
                Exit Sub
             End If
          End If
       End If
       '
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text1.SetFocus
       '
    End If
    '
End Sub

Private Sub Text11_GotFocus()
    If TRIM$(Text11.TEXT) = "" Then
       Text11.TEXT = Space$(18)
    End If
    Text11.SelStart = 0
    Text11.SelLength = Len(Text11.TEXT)
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    IMPOGRA(Index).SelStart = 0
    IMPOGRA(Index).SelLength = Len(IMPOGRA(Index).TEXT)
    IMPOGRA(0).SetFocus
    End If
End Sub

Private Sub Text11_LostFocus()
    '
    If TRIM$(Text11.TEXT) <> "" Then
       Call SIGUEFACO1
    End If
    '
End Sub

Private Sub Text2_GotFocus()
   If XVALO(Text1.TEXT) > 0 Then
      Text1.SetFocus
      Exit Sub
   End If
   
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2.TEXT)
End Sub

Private Sub Text3_GotFocus()
   If XVALO(Text1.TEXT) > 0 Then
      Text1.SetFocus
      Exit Sub
   End If

End Sub

Private Sub Text4_GotFocus()
   If XVALO(Text1.TEXT) > 0 Then
      Text1.SetFocus
      Exit Sub
   End If

End Sub

Private Sub Text5_GotFocus()
  '
   If XVALO(Text13.TEXT) > 0 Then
      Text1.SetFocus
      Exit Sub
   End If
   '
End Sub

Private Sub Text6_Click()
    'Text1.SetFocus
End Sub

Private Sub Text6_GotFocus()
    NC% = XVALO(Text1.TEXT)
    If NC% < 1 Or ECOARCH$("ACREDOR", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Exit Sub
    End If
    '
    If Control$("", "VALICAI") = "SI" Then
      If VALIDACAI = 0 Then Exit Sub
    End If
    '
    If Left$(RASOCLI$((-1) * NC%), 6) <> "VARIOS" Then
       Text2.TEXT = RASOCLI$((-1) * NC%)
       Text3.TEXT = DOMICLI$((-1) * NC%)
       Text4.TEXT = CPOSCLI$((-1) * NC%)
       Text5.TEXT = LOCACLI$((-1) * NC%)
       Text13.TEXT = CUITCLI$((-1) * NC%)
    End If
    '
    Text6.SelStart = 0
    Text6.SelLength = Len(Text6.TEXT)
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub
'
Private Sub Text6_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text6.TEXT = VDEVV$
    End If
End Sub

Private Sub Text9_GotFocus()
    Text9.SelStart = 0
    Text9.SelLength = Len(Text9.TEXT)
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text11.SetFocus
   End If
End Sub

Private Sub Text9_LostFocus()
    VDEVV$ = TRIM$(Text9.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FF% = FECHANUM(VDEVV$)
          FEX = FF%
          Text9.TEXT = FECHATEX(FEX)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          Text9.TEXT = HO$
          Text9.SetFocus
    End If
End Sub

Private Sub Text6_LostFocus()
    VDEVV$ = TRIM$(Text6.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FF% = FECHANUM(VDEVV$)
          FEX = FF%
          Text6.TEXT = FECHATEX(FEX)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          Text6.TEXT = HO$
          Text6.SetFocus
    End If
End Sub

Sub CALTOFAC()
'   Dim FVI%(10), IM#(10)
'   '
'   TOBRUFAC# = XVALO(ITOGRA.TEXT)
'   For KKI% = 0 To 6
'     TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
'     If KKI% > 0 Then
'        TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
'     End If
'   Next KKI%
'   TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text7.TEXT) + 0.5)) / 100
'   TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text8.TEXT) + 0.5)) / 100
'   TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text12.TEXT) + 0.5)) / 100
'   TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text22.TEXT) + 0.5)) / 100
'   '
'   IVENCI(0).TEXT = CSTRING$(MKD$(TOBRUFAC# - XVALO(Label10)), 4, 11, 2, 2)
'   '
'   If XVALO(Text1) > 32767 Then
'     Call MENSERR(24, "Número de Cliente Inválido ")
'     Exit Sub
'   End If
'   NCI% = XVALO(Text1)
'   If NCI% < 1 Then Exit Sub
'   '
'   COPAGI% = CPAGCLI%((-1) * NCI%)
'   Call BUCONPA(LU$, COPAGI%, COPAG$, FREF%, PLA$)
'
'   IORIG# = TOBRUFAC#
'   FF = FECHANUM(Text6)
'   FFS$ = FECHATEX$(FF)
'   PLA$ = Left$(PLA$ + String$(6, 0), 6)
'
'   If FREF% = 2 Then FREF% = 1   ' AUN NO FUNCIONA SEMANA DE FACTURA
'   If FREF% = 3 Then
'     If XVALO(Left$(FFS$, 2)) <= 15 Then FEBAS$ = "15" + Mid$(FFS$, 3)
'     If XVALO(Left$(FFS$, 2)) > 15 Then FEBAS$ = "31" + Mid$(FFS$, 3)
'   End If
'   If FREF% = 1 Then FEBAS$ = FFS$
'   If FREF% = 4 Then FEBAS$ = "31" + Mid$(FFS$, 3)
'   '
'   For i% = 1 To 6: FVI%(i%) = 0: Next i%
'   A11 = FECHANUM(FEBAS$): B22 = A11
'   NUPAG% = 1
'   For i% = 1 To 6
'     PLAI% = Asc(Mid$(PLA$, i%, 1))
'     If PLAI% > 0 Or i% = 1 Then
'       While DIENTRE%(A11, B22) < PLAI%
'         B22 = DIPOST(B22)
'       Wend
'       FVI%(i%) = B22
'       NUPAG% = i%
'     End If
'   Next i%
'   '
'   If NUPAG% < 1 Then NUPAG% = 1
'   IM#(1) = (Int(100 * IORIG# / NUPAG% + 0.5)) / 100
'   '
'   For H% = 2 To NUPAG% - 1
'     IM#(H%) = IM#(1)
'   Next H%
'   IM#(NUPAG%) = IORIG#
'   '
'   For H% = 1 To NUPAG% - 1
'     IM#(NUPAG%) = IM#(NUPAG%) - IM#(H%)
'   Next H%
'   '
'   ICANCE# = 0
'   For H% = 1 To NUPAG%
'     If ICANCE# > 0 Then
'       If ICANCE# >= IM#(H%) Then
'           ICANCE# = ICANCE# - IM#(H%)
'           IM#(H%) = 0
'         Else
'           IM#(H%) = IM#(H%) - ICANCE#
'           ICANCE# = 0
'       End If
'     End If
'   Next H%
'   '
'   For H% = 1 To 9
'     FEX = FVI(H%)
'     FEVEN(H% - 1) = FECHATEX$(FEX)
'     IVENCI(H% - 1) = FORMATNUM$(IM#(H%), "F11.2")
'   Next H%
   '
End Sub
'
Sub BLAFACOM()
   '
   TOTIMPU# = 0
   TONEGRA# = 0
   TOEXEN# = 0
   TOTIVA# = 0
   TORETIV# = 0
   TOTFAC# = 0
   '
   List1.Clear
   List2.Clear
   Label4.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Text11.TEXT = ""
   Check1.Value = 1
   '
   FF% = HOY(HO$)
   Text9.TEXT = HO$
   Text6.TEXT = HO$
   '
   JJ& = 0
   For U& = 1 To ULTREG("PROVINB")
     XX$ = REGLEIDO$("PROVINB", U&)
     If CVI(Mid$(XX$, 27, 2)) > 0 Then
       If ECOARCH$("CUECON", Mid$(XX$, 27, 2)) <> "" Then
         Call REPLA(XX$, String$(8, 0), 57, 8)
         JJ& = JJ& + 1
         Call GRAREG("!PROVINB", XX$, JJ&)
       End If
     End If
   Next U&
   Call SETULTREG("!PROVINB", JJ&)
   Call CIERRARCH("PROVINB")
   '
   Label25 = ""
   Label21 = ""
   Text14 = ""
   Text15 = ""
   Text16 = ""
   Text17 = ""
   '
'   ITOGRA.TEXT = Space$(11): ITOGRA.Refresh
   Text7.TEXT = Space$(7)
   Text8.TEXT = Space$(7)
   Text12.TEXT = Space$(7)
   TEXT10.TEXT = " "
   Text20 = ""
   Text22 = "" ' TEXT REDONDEO
   '
   Label18.Caption = ""
   For KKI% = 0 To 6: IMPOGRA(KKI%).TEXT = Space$(11): Next KKI%
   For KKI% = 0 To 6: IMPOIVA(KKI%).TEXT = Space$(9): Next KKI%
   For KKI% = 0 To 8: FEVEN(KKI%).TEXT = Space$(8): Next KKI%
   For KKI% = 0 To 8: IVENCI(KKI%).TEXT = Space$(11): Next KKI%
   VScroll1.Value = 0
   '
   For KKI% = 0 To 10
      TOFAC#(KKI%) = 0
      FVII%(KKI%) = 0
      IMV#(KKI%) = 0
      IMV1#(KKI%) = 0
      GRU%(KKI%) = 0
      IPAR#(KKI%) = 0
   Next KKI%
   '
   MODIFIX% = 0
   Label36 = ""
   Text19 = ""
   Text19.BackColor = vbWhite
   Label35 = ""
   Label31 = ""
   '
   Label40 = "0"
   Picture4.Visible = False
   '
End Sub
'
Sub GRAFACOM()
   '
   Dim VENCIPESOS#(16)
   '
   NROX$ = TRIM$(Text11.TEXT)
   If Len(NROX$) < 1 Then
      Call MENSERR(24, "Número de Comprobante no Válido")
      Text11.SetFocus
      Exit Sub
   End If
   '
   If Left$(NROX$, 3) = "CO-" Then
      Call MENSERR(24, "Número de Comprobante no Válido")
      Text11.SetFocus
      Exit Sub
   End If
   '
   If TIDOCOM$ = "CO" Then
      If Len(NROX$) > 15 Then
         Call MENSERR(24, "Número de Comprobante no Válido")
         Text11.SetFocus
         Exit Sub
      End If
      NROX$ = "-" + NROX$
      While Len(NROX$) < 13: NROX$ = "0" + NROX$: Wend
      NROX$ = "CO-" + TRIM$(Label20.Caption) + "-" + AJUSTI$(NROX$, 13)
   End If
   '
   SIG% = 1: Tipo% = 2
   If Left$(TRIM$(NROX$), 2) = "TF" Then Tipo% = 22
   If Left$(TRIM$(NROX$), 2) = "RH" Then Tipo% = 32
   If Left$(TRIM$(NROX$), 2) = "CO" Then Tipo% = 42
   If Left$(TRIM$(NROX$), 2) = "ND" Then Tipo% = 3
   If Left$(TRIM$(NROX$), 2) = "CP" Then Tipo% = 3
   If Left$(TRIM$(NROX$), 2) = "NC" Then Tipo% = 6: SIG% = (-1)
   If Left$(TRIM$(NROX$), 2) = "DP" Then Tipo% = 6: SIG% = (-1)
   '
   If Len(NROX$) <> 18 Then
      Call MENSERR(24, "Formato Número de Comprobante no Válido")
      VACON% = 0
      Exit Sub
   End If
   '
'   SUC% = 0: NRO& = 0
'   If TIDOCOM$ <> "CO" Then
'      VA% = Asc(Mid$(NROX$, 4, 1)) - 64
'      If VA% = 5 Then VA% = 0 ' EXPORTACION
'      If LETRAFA$ = "M" Then VA% = 1
'      If VA% < 0 Or VA% > 3 Then
'          Call MENSERR(24, "Letra de Documento Debe ser 'A', 'C', 'M' o 'E'")
'          VACON% = 0
'          Exit Sub
'      End If
'      SUC% = XVALO(Mid$(NROX$, 6, 4))
'      NRO& = XVALO(Mid$(NROX$, 11, 8))
'      '
'      If NRO& < 1 Then
'         Call MENSERR(24, "Falta Número de Documento")
'      End If
'   End If
   '
   NUDOCUCOM$ = NROX$
   '
   
   NC% = XVALO(Text1.TEXT)
    If NC% > 0 Then
    If NC% < 1 Or ECOARCH$("ACREDOR", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Cuenta Inexistente en B.D. de Proveedores")
       VACON% = 0
       Exit Sub
    End If
    
    If TICUEPRO%((-1) * NC%) <> 0 Then
       Call MENSERR(24, "Imposible Registrar.\Cuenta de Gastos.")
       VACON% = 0
       Exit Sub
    End If
   End If
   '
   CUITT$ = TRIM$(Text13.TEXT)
   If NC% > 0 Then
     CUITT$ = CUITCLI$((-1) * NC%)
   End If
   If CUITT$ = "" Or VALICUIT%(CUITT$) < 1 Then
      PVCL% = PIVACLI%((-1) * NC%)
      If PVCL% <> 0 And PVCL% <> 5 Then ' EXCLUYE EXPORTACION / IMPORTACION
         Call MENSERR(24, "Número de C.U.I.T. de Proveedor no Válido")
         VACON% = 0
         Exit Sub
      End If
   End If
   '
   If NC% > 0 Then
    CUEMA$ = TRIM$(CUEMADRE$((-1) * NC%))
    If CUEMA$ = "" Or CUECOINT%(CUEMA$) < 1 Then
      Call MENSERR(24, "Falta Cuenta Madre de " + TRIM$(Text2))
      VACON% = 0
      Exit Sub
    End If
   End If
   '
   FF% = FECHANUM(Text9.TEXT)
   HOI% = HOY(HO$)
   If FF% < 1 Or FF% > HOI% + 1 Then
       Call MENSERR(24, "Fecha no Válida")
       Text9.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
   FFRE% = FECHANUM(Text6.TEXT)
   HOI% = HOY(HO$)
   If FFRE% < 1 Or FFRE% > HOI% + 1 Then
       Call MENSERR(24, "Fecha de Emisión no Válida")
       Text6.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
   If DIENTRE%(FF%, HOI%) > 55 Then
       Call MENSERR(24, "Fecha no Válida")
       Text9.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
'   FECHALIM% = XVALO(Control$("", "CIEMECOM"))
'   If FF% <= FECHALIM% Then
'       Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
'       Text9.SetFocus
'       VACON% = 0
'       Exit Sub
'   End If
   '
   If FF% < FECHANUM(Text6.TEXT) Then
       Call MENSERR(24, "No Válido Contabilizar en Fecha Anterior\a la de Emisión del Comprobante.")
       Text9.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
'   If XVALO(ITOGRA.TEXT) < (-0.005) Then
'      Call MENSERR(24, "Importes no Válidos")
'      Call SIGUEFACO3
'      Exit Sub
'   End If
   '
   If LETRAFA$ <> "A" And LETRAFA$ <> "M" Then
     If Abs(XVALO(IMPOGRA(0).TEXT)) >= 0.005 Then
       Call MENSERR(24, "Importes no Válidos")
       Call SIGUEFACO3
       Exit Sub
     End If
   End If
   '
'   SUIMPUTA# = XVALO(Label45)
'   KKI% = 0
'   For u& = 1 To List1.ListCount
'      List1.ListIndex = u& - 1
'      Z$ = List1.TEXT
'      CUE% = XVALO(Left$(Z$, 5))
'      IPA# = XVALO(Mid$(Z$, 40, 12))
'      '
'      If CUE% > 0 Then
'         If IPA# >= 0.005 Then
'            KKI% = KKI% + 1
'            If KKI% > 8 Then GoTo 55
'            GRU%(KKI%) = CUE%
'            IPAR#(KKI%) = IPA#
'            COEBON1 = COEFDESCU(TRIM$(Text19))
'            If COEBON1 > 0 Then
'               IPA# = ROUND#(IPA# * (1 - COEBON1 / 100))
'               IPAR#(KKI%) = IPA#
'            End If
'            SUIMPUTA# = SUIMPUTA# + IPA#
'         End If
'      End If
'   Next u&
   '
'55 If KKI% > 8 Then
'      Call MENSERR(24, "Demasiados Conceptos - Máximo es 8")
'      Exit Sub
'   End If
   '
   If XVALO(Text12.TEXT) >= 0.005 Then
      If TRIM$(Label18.Caption) = "" Then
         Call MENSERR(24, "Imposible Registrar - Falta Jurisdiccion I.Brutos")
         Exit Sub
      End If
   End If
   '
   For KKI% = 0 To 6
     If KKI% = 4 Then
         IMPOX# = XVALO(IMPOIVA(6))     ' iva aduana
       ElseIf KKI% = 5 Then
         IMPOX# = XVALO(Text7)
       ElseIf KKI% = 6 Then
         IMPOX# = XVALO(Text8)
       Else
         IMPOX# = XVALO(IMPOIVA(KKI%))
     End If
     '
     
     If IMPOX# >= 0.005 Then
       CUEXTE$ = Control$("", "CUEIVA-" + TRIM$(Str$(KKI%)))
       CUEINTE% = CUECOINT%(CUEXTE$)
       If TRIM$(CUEXTE$) = "" Or CUEINTE% < 1 Then
         Call MENSERR(24, "Imposible Registrar.\Faltan Cuentas Contables IVA.")
         Exit Sub
       End If
     End If
   Next KKI%
   '
   If TRIM$(TEXT10) <> "" And TRIM$(TEXT10) <> "*" Then
     If XVALO(Text12) >= 0.005 Then
       CUEXTE$ = CUEIBRUPRO$(TEXT10)
       CUEINTE% = CUECOINT%(CUEXTE$)
       If TRIM$(CUEXTE$) = "" Or CUEINTE% < 1 Then
         Call MENSERR(24, "Imposible Registrar.\Falta Cuenta Contable Correspondiente\a Percepc.Ing.Brutos " + TRIM$(Label18) + ".")
         Exit Sub
       End If
     End If
   End If
   '
   If Abs(XVALO(Text22)) > 0 Then
     If CUECOINT%(Control$("FACOMPRA", "REDONDEO")) < 1 Then
       Call MENSERR(24, "Falta Cuenta Contable de Redondeo")
       Exit Sub
     End If
   End If
   '
   SUNETO# = XVALO(IMPOGRA(0).TEXT)
   SUBRUTO# = SUNETO#
   '
   TOFAC#(1) = SIG% * SUNETO#
   TOFAC#(2) = SIG% * (Abs(XVALO(IMPOGRA(4).TEXT)) + Abs(XVALO(IMPOGRA(5).TEXT)))
   TOFAC#(3) = 0                    ' impuestos intermos
   TOFAC#(4) = 0                    ' IVA TASAS GENERALES
   TOFAC#(5) = 0                    ' iva ADUANA
   TOFAC#(6) = SIG% * Abs(XVALO(Text7.TEXT))
   TOFAC#(7) = SIG% * Abs(XVALO(Text8.TEXT))
   TOFAC#(8) = SIG% * Abs(XVALO(Text12.TEXT))
   '
   REDON# = XVALO(Text22.TEXT)
   TOFAC#(9) = SIG% * Abs(REDON#)
   If REDON# < 0 Then TOFAC#(9) = (-1) * TOFAC#(9) ' EN EL CASO DEL QUE REDONDEO SEA NEGATIVO

   '
   For KKI% = 0 To 6
     SUBRUTO# = SUBRUTO# + (Int(100 * XVALO(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
     If KKI% < 6 Then
          TOFAC#(4) = TOFAC#(4) + SIG% * (Int(100 * XVALO(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
        Else
          TOFAC#(5) = TOFAC#(5) + SIG% * (Int(100 * XVALO(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
     End If
     If KKI% > 0 Then
        SUNETO# = SUNETO# + (Int(100 * XVALO(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
        SUBRUTO# = SUBRUTO# + (Int(100 * XVALO(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
        If KKI% <> 4 And KKI% <> 5 Then
          TOFAC#(1) = TOFAC#(1) + SIG% * (Int(100 * XVALO(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
        End If
     End If
   Next KKI%
'   If Abs(SUNETO - SUIMPUTA#) >= 0.005 Then
'      Call MENSERR(24, "Importes Incoherentes")
'      Exit Sub
'   End If
   '
   SUBRUTO# = SUBRUTO# + (Int(100 * XVALO(Text7.TEXT) + 0.5)) / 100
   SUBRUTO# = SUBRUTO# + (Int(100 * XVALO(Text8.TEXT) + 0.5)) / 100
   SUBRUTO# = SUBRUTO# + (Int(100 * XVALO(Text12.TEXT) + 0.5)) / 100
   SUBRUTO# = SUBRUTO# + (Int(100 * XVALO(Text22.TEXT) + 0.5)) / 100 'REDONDEO

   TOFAC#(0) = (-1) * SIG% * SUBRUTO#
   '
   If SUNETO# < 0 Or SUBRUTO# < 0.005 Then
      Call MENSERR(24, "Falta Ingresar Importes")
      VACON% = 0
      Exit Sub
   End If
   '
   SUMAVEN# = 0
'   For KKI% = 1 To 9
'      IMV1#(KKI%) = 0
'      FVII%(KKI%) = FECHANUM(FEVEN(KKI% - 1).TEXT)
'      If FVII%(KKI%) > 0 Then
'         IMV#(KKI%) = XVALO(IVENCI(KKI% - 1).TEXT)
'         If IMV#(KKI%) < 0 Then
'            Call MENSERR(24, "Importe a Vencer Negativo")
'            VACON% = 0
'            Exit Sub
'         End If
'         IMV1#(KKI%) = IMV#(KKI%) * SIG%
'         SUMAVEN# = SUMAVEN# + IMV#(KKI%)
'      End If
'   Next KKI%
'   If Abs(SUMAVEN# + XVALO(Label10) - SUBRUTO#) >= 0.005 Then
'      Call MENSERR(24, "Suma de Vencimientos Difiere de Total")
'      VACON% = 0
'      Exit Sub
'   End If
   '
'   SUMATOT# = 0
'   For KKI% = 0 To 9
'      SUMATOT# = SUMATOT# + TOFAC#(KKI%)
'   Next KKI%
'   If Abs(SUMATOT#) >= 0.005 Then
'      Call MENSERR(24, "Valores Incoherentes")
'      VACON% = 0
'      Exit Sub
'   End If
   '
   VACON% = 1
   SUMACRE# = 0: ATX% = 0
   Call BLOQARCH("CUECORR")
   Call BLOQARCH("CUECOSS")
   Call BLOQARCH("CAJABANC")
   Call BLOQARCH("PROVED2")
   '
'   For u& = 1 To List2.ListCount
'      List2.ListIndex = u& - 1
'      Z$ = List2.TEXT
'      If XVALO(Mid$(Z$, 70, 14)) < 0 Then
'         Call MENSERR(24, "Valores Aplicados Negativos")
'         VACON% = 0
'         Exit Sub
'      End If
'      '
'      DOCU$ = Left$(Z$, 20)
'      If DOCU$ <> "Anticipo no Aplicado" Then
'         CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
'         If CANCE# > 0 Then
'            NREG& = XVALO(TRIM$(Mid$(Z$, 119, 8)))
'            If NREG& > 0 Then
'              If NREG& <= ULTREG&("CUECORR") Then
'                ORD% = XVALO(Mid$(Z$, 127, 2))
'                If ORD% > 0 Then
'                  If ORD% <= 6 Then
'                    '
'                    SS$ = REGLEIDO$("CUECORR", NREG&)
'                    If Mid$(SS$, 7, 18) <> Left$(DOCU$, 18) Then ATX% = 1
'                    SALDO$ = Mid$(SS$, 139, 60)
'                    SALDEU# = CVD(Mid$(SALDO$, 10 * ORD% - 7, 8))
'                    If SALDEU# + 0.005 < CANCE# Then ATX% = 1
'                    SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(Z$, 70, 14)))
'                  End If
'                End If
'              End If
'            End If
'         End If
'      End If
'   Next u&
   '
   If ATX% > 0 Then
      Call MENSERR(24, "Imposible Registrar este Documento.\  \Reingrese los Datos Completos.\Si Persiste el Problema,\Consulte a su Soporte de Sistemas FLEXOFT.")
      GoTo 84
   End If
   '
   If SUMACRE# > Abs(TOFAC#(0)) + 0.005 Then
      Call MENSERR(24, "Importes Aplicados Superan Total Factura")
      VACON% = 0
      Exit Sub
   End If
   '
   '
   Call BLOQARCH("CUECORR")
   Call BLOQARCH("CUECOSS")
   Call BLOQARCH("CAJABANC")
   Call BLOQARCH("APLIPAG")
   Call BLOQARCH("PROVED2")
   Screen.MousePointer = 11
   '
   REFEDOC$ = AJUSTI$(Text2.TEXT, 24) + TRIM$(Text13.TEXT)
   RECUECO& = 1 + ULTREG&("CUECORR")
   If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then
     IMV1#(1) = IMV1#(1) + XVALO(Label10)  ' DESCONTAR IMPORTES APLICADOS
   End If
   If XVALO(Label25) > 0 Then
     If TIDOCOM$ <> "CO" Then
       For KKI% = 1 To 6: IMV1#(KKI%) = 0: FVII%(KKI%) = 0: Next KKI%
     End If
   End If
   If TIDOCOM$ = "CO" Then
     DOCUACTIVO$ = NUDOCUCOM$
   End If
'   Call CUECORR(FF%, (-1) * NC%, TIPO%, VA%, SUC%, NRO&, REFEDOC$, MO$, TOFAC#(), FVII%(), IMV1#(), GRU%(), IPAR#(), 1, FFRE%)
'   Call CUECOSS(NUDOCUCOM$)
   '
   If XVALO(Label25) > 0 Then
     If TIDOCOM$ = "CO" Then
       Call COMUNI("Imposible Cancelación Contado.\Genere Orden de Pago C.Cte.")
       NRECIB$ = TRIM$(Text11)
       Exit Sub
     End If
     Call REPLA(NUDOCUCOM$, "FF", 1, 2)
     For KKI% = 0 To 10: TOFAC1#(KKI%) = (-1) * TOFAC#(KKI%): Next KKI%
     For KKI% = 0 To 10: GRU1%(KKI%) = 0: IPAR1#(KKI%) = 0: Next KKI%
'     Call CUECORR(FF%, (-1) * NC%, 14, VA%, SUC%, NRO&, REFEDOC$, MO$, TOFAC1#(), FVII%(), IMV1#(), GRU1%(), IPAR1#(), 1, FFRE%)
'     Call CUECOSS(NUDOCUCOM$)
   End If
   '
   ' grabar factura en CAJABANC
   TOFA0# = 0
   TICA0# = TCMONE#(MONEMI%)
   If Len(TRIM$(Text11)) < 18 Then
        NRECIB$ = TRIM$(Text11)
        While Len(NRECIB$) > 9: NRECIB$ = Mid$(NRECIB$, 2): Wend
        NRECIB$ = Left$(NROX$, 3) + NRECIB$
      Else
        NRECIB$ = Left$(Text11, 5) + Mid$(Text11, 9, 1) + "-" + Mid$(Text11, 14, 5)
   End If
   '
20 NC1% = (-1) * NC%
   Call APERBASDAT("CABAN")
   '
   CoComCor$ = NRECIB$
   CoComLar$ = NROX$
'   CodiEmp% = 0
'   CodiSuc% = SUC%
   NumeClie = 0
   NumeProv% = NC%
   FechEmi = CVDAT(FF%)
   '
   Referen$ = RASOCLI$(NC1%)
     If Left$(UCase$(RASOCLI$(NC1%)), 6) = "VARIOS" Then
       Referen$ = Left$(Left$(Text2, 32) + "//" + Left$(Text13, 16), 48)
     End If
   FeContab = CVDAT(FF%)
   FeAcredi = CVDAT(FF%)
   NumeTal = 0
   NumeCom = NRO&
   Observa$ = TRIM$(Text20)
      If Observa$ = "" Then Observa$ = RASOCLI$(NC1%)
   NumCai$ = TRIM$(Mid$(Text16, 1, 14))
   FechCai = CVDAT(FECHANUM(Text17))
   InclIva$ = "": If Check1.Value = 0 Then InclIva$ = "NO"
   On Error GoTo 0
   '
   ' graba detalle por imputación de compras
'   For u& = 1 To List1.ListCount
'     Z$ = List1.List(u& - 1)
'     CuentaCo% = XVALO(Left$(Z$, 5))
'     CentroCos$ = Mid$(Z$, 65, 8)
'     ItConcep$ = Mid$(Z$, 75, 8) 'P/ conceptos
'       COEBON1 = COEFDESCU(TRIM$(Text19))
'     ImpoDeb# = REDONVAL#((1 - COEBON1 / 100) * XVALO(Mid$(Z$, 40, 12)) * TICA0#, 2)
'       If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
'     DEMOVX$ = Mid$(Z$, 8, 30)
'     '
'     If CuentaCo% > 0 Then
'       If Abs(ImpoDeb#) >= 0.005 Then
'         TiMovi$ = "COMPRA"
'         OpMovi$ = Left$(Text11, 2)
'         VaMovi$ = Label20
'         Descrip$ = TRIM$(DEMOVX$) + " - " + RASOCLI$(NC1%)
'         TOFA0# = TOFA0# + ImpoDeb#
'         Call GRACABAN
'       End If
'     End If
'     CentroCos$ = ""
'   Next u&
   '
   ' grabar importes de IVA
   SUMIVA# = 0
   If XVALO(IMPOGRA(4)) >= 0.005 Then   ' GRABAR IMPORTE EXENTO
     TiMovi$ = "IVA"
     OpMovi$ = "CrFi"
     VaMovi$ = "Exento"
     ImpoDeb# = 0
     BasImpo# = REDONVAL#(XVALO(IMPOGRA(4)) * TICA0#, 2)
     Descrip$ = "IVA B.Imp. Exento " + NRECIB$
     CuentaCo% = CUECOINT%(Control$("", "CUEIVA-0"))
     TOFA0# = TOFA0# + ImpoDeb#
     Call GRACABAN
   End If
   '
   If XVALO(IMPOGRA(5)) >= 0.005 Then   ' GRABAR IMPORTE NO GRAVADO
     TiMovi$ = "IVA"
     OpMovi$ = "CrFi"
     VaMovi$ = "No Grav"
     ImpoDeb# = 0
     BasImpo# = REDONVAL#(XVALO(IMPOGRA(5)) * TICA0#, 2)
     Descrip$ = "IVA B.Imp. No Grav." + NRECIB$
     CuentaCo% = CUECOINT%(Control$("", "CUEIVA-0"))
     TOFA0# = TOFA0# + ImpoDeb#
     Call GRACABAN
   End If
   '
   For KKI% = 0 To 6
     TiMovi$ = "IVA"
     OpMovi$ = "CrFi"
     VaMovi$ = ""
     On Error Resume Next
     VaMovi$ = TASIVA(KKI%)
     On Error GoTo 0
     Descrip$ = "IVA Credito Fiscal " + NRECIB$
     '
     BasImpo# = 0
     If KKI% = 4 Then
         ImpoDeb# = XVALO(IMPOIVA(6))
         BasImpo# = XVALO(IMPOGRA(6))
         VaMovi$ = "Aduana"
       ElseIf KKI% = 5 Then
         ImpoDeb# = XVALO(Text7)
         VaMovi$ = "Per-IVA"
         Descrip$ = "Percepcion I.V.A. " + NRECIB$
       ElseIf KKI% = 6 Then
         ImpoDeb# = XVALO(Text8)
         TiMovi$ = "GAN"
         VaMovi$ = "Per-Gan"
         Descrip$ = "Percepcion Ganancias " + NRECIB$
       Else
         ImpoDeb# = XVALO(IMPOIVA(KKI%))
         SUMIVA# = SUMIVA# + ImpoDeb#
         If KKI% > 0 Then
             BasImpo# = XVALO(IMPOGRA(KKI%))
           Else
             BasImpo# = XVALO(IMPOGRA(0))
         End If
     End If
     '
     If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
     If Abs(ImpoDeb#) >= 0.005 Then
       CuentaCo% = CUECOINT%(Control$("", "CUEIVA-" + TRIM$(Str$(KKI%))))
       ImpoDeb# = REDONVAL#(ImpoDeb# * TICA0#, 2)
       BasImpo# = REDONVAL#(BasImpo# * TICA0#, 2)
       TOFA0# = TOFA0# + ImpoDeb#
       Call GRACABAN
     End If
     BasImpo# = 0
   Next KKI%
   '
   ' grabar percepcion Ingresos Brutos - CON OPCION MULTIPLES JURISDICCIONES
   If TRIM$(TEXT10) <> "" Then
     If XVALO(Text12) >= 0.005 Then
       If TRIM$(TEXT10) = "*" Then
            For U& = 1 To ULTREG&("!PROVINB")
              XXZ$ = REGLEIDO$("!PROVINB", U&)
              IPERPA# = CVD(Mid$(XXZ$, 57, 8))
              If Abs(IPERPA#) >= 0.005 Then
                CuentaCo% = CVI(Mid$(XXZ$, 27, 2))
                ImpoDeb# = IPERPA#
                  If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
                TiMovi$ = "IBRU"
                OpMovi = "IB-" + Left$(XXZ$, 1)
                VaMovi = "PERC"
                Descrip$ = "Percep.I.Brutos " + NRECIB$
                ImpoDeb# = REDONVAL#(ImpoDeb# * TICA0#, 2)
                TOFA0# = TOFA0# + ImpoDeb#
                Call GRACABAN
              End If
            Next U&
          Else
            CuentaCo% = CUECOINT%(CUEIBRUPRO$(TEXT10))
            ImpoDeb# = XVALO(Text12)
              If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
            TiMovi$ = "IBRU"
            OpMovi = "IB-" + TEXT10
            VaMovi = "PERC"
            Descrip$ = "Percep.I.Brutos " + NRECIB$
            ImpoDeb# = REDONVAL#(ImpoDeb# * TICA0#, 2)
            TOFA0# = TOFA0# + ImpoDeb#
            Call GRACABAN
       End If
     End If
   End If
   '
   CuentaCo% = CUECOINT(Control("FACOMPRA", "REDONDEO"))
   ImpoDeb# = XVALO(Text22)
   If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
   TiMovi$ = "REDCOM"
   OpMovi = Left$(Text11, 2)
   VaMovi = Label20
   Descrip$ = "Redondeo " + NRECIB$
   ImpoDeb# = REDONVAL#(ImpoDeb# * TICA0#, 2)
   TOFA0# = TOFA0# + ImpoDeb#
   Call GRACABAN
   '
   ' graba factura en c.corriente proveedor
   CuentaCo% = CUECOINT%(CUEMADRE$(-1 * NC%))
   'ImpoDeb# = (-1) * Abs(TOFAC#(0))
   ImpoDeb# = (-1) * Abs(TOFA0#)
     If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
   TiMovi = "FCOM"
   OpMovi = Left$(Text11, 2)
   VaMovi = Label20
   Descrip$ = CoComLar$ + " - " + RASOCLI$(NC1%)
   FeAcredi = CVDAT(FFRE%)
   Call GRACABAN
   FeAcredi = CVDAT(FF%)
   '
   ' grabar pago contado
   If TRIM$(Label25) <> "" Then
     If TIDOCOM$ = "CO" Then
       GoTo 80
     End If
     '
     CuentaCo% = CUECOINT%(CUEMADRE$(-1 * NC%))
     TiMovi = "OPAG"          ' ORDEN DE PAGO
     OpMovi = "PC"            ' pago contado
     VaMovi = "CTDO"          ' CONTADO
     Descrip = TRIM$(DEMOVX$) + " - " + RASOCLI$(NC1%)
     'ImpoDeb# = Abs(TOFAC#(0))
     ImpoDeb# = Abs(TOFA0#)
       If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
     Call GRACABAN
     '
     FeContab = CVDAT(FF%)
     FeAcredi = CVDAT(FECHANUM(Text15))
       If FeAcredi < Now Then FeAcredi = Now
     CuentaCo% = CUEMEPAG%(XVALO(Label25))
     TiMovi = "CHEMI"    ' VALOR EMITIDO / ENTREGADO
     OpMovi = TRIM$(Label25)
     VaMovi = Label20
       BANCO$ = TRIM$(Label21)
       NUCHE$ = TRIM$(Text14)
       FEVECHE$ = Text15
     Descrip$ = AJUSTI$(BANCO$, 18) + AJUSTI$(NUCHE$, 14) + FECHATEX$(FEX)
     'ImpoDeb# = (-1) * Abs(TOFAC#(0))
     ImpoDeb# = (-1) * Abs(TOFA0#)
       If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
     Observa$ = TRIM$(BANCO$) + " " + TRIM$(NUCHE$) + " " + TRIM$(FEVECHE$)
     Call GRACABAN
     '
   End If
   '
   ' Grabar Vencimientos de Facturas ******************************************
   SIGNO% = 1
   If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then SIGNO% = (-1)
   If XVALO(Label25) < 1 Then ' no hay cancelacion de contado
     IBruTot# = Abs(TOFAC#(0))
     IBruTot# = Abs(TOFA0#)
     IIvaTot# = SUMIVA# * TICA0#
     '
     PRICONVEN% = -1
     SALBRUTO# = Abs(TOFA0#) * SIGNO%
     For KKI% = 0 To 8
       VENCIPESOS#(KKI%) = 0
       If FEVEN(KKI%) <> "" Then
         If XVALO(IVENCI(KKI%)) >= 0.005 Then
           If PRICONVEN% < 0 Then PRICONVEN% = KKI%
           VENCIPESOS#(KKI%) = REDONVAL#(XVALO(IVENCI(KKI%)) * SIGNO% * TICA0#, 2)
           SALBRUTO# = SALBRUTO# - VENCIPESOS#(KKI%)
         End If
       End If
     Next KKI%
     If PRICONVEN% >= 0 Then
       VENCIPESOS#(PRICONVEN%) = VENCIPESOS#(PRICONVEN%) + SALBRUTO#
     End If
     '
     For KKI% = 0 To 8
       If FEVEN(KKI%) <> "" Then
         If Abs(VENCIPESOS#(KKI%)) >= 0.005 Then
           NuorVen% = KKI% + 1
           FeVenc = CVDAT(FECHANUM(FEVEN(KKI%)))
           ImVenc# = VENCIPESOS#(KKI%)
           Call GRACREPEN
         End If
       End If
     Next KKI%
     '
   End If
   
80 ' *********************************** aqui blanquear facturas canceladas
   '
   FechEmi = CVDAT(FF%)
   RefApli$ = CoComLar$
   NumeProv% = XVALO(Text1)
   '
   XX% = 0
   For U& = 1 To List2.ListCount
     List2.ListIndex = U& - 1
     Z$ = List2.List(U& - 1)
     DOCU$ = Left$(Z$, 20)
     CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
     If CANCE# > 0 Then
       DOCUCA$ = TRIM$(Left$(Z$, 18))
       ICAN# = CANCE#
       '
       DocOrCor$ = ""
       TotOrig# = 0
       IvaOrig# = 0
       '
       Call APERBASDAT("CABAN")
       SQLX$ = "SELECT * FROM SACREPEN"
       SQLX$ = SQLX$ + " WHERE NuProve = " & NumeProv% & " "
       SQLX$ = SQLX$ + " AND CODICOM_LAR = '" & DOCUCA$ & "' "
       SQLX$ = SQLX$ + " AND SALDACRE > 0 "
       Dim SACCCREPEN As ADODB.Recordset
       Set SACCCREPENP = New ADODB.Recordset
25     On Error Resume Next
       SACCCREPENP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 25
       End If
       On Error GoTo 0

       With SACCCREPENP
         '.FindFirst "NuProve = " & NumeProv% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDACRE > 0"
         'If .NoMatch = False Then
         On Error Resume Next
         If Not (.EOF And .BOF) Then
           On Error GoTo 0
           Do While Not (.EOF)
            PARCAN# = ICAN#: If !SALDAcre < ICAN# Then PARCAN# = !SALDAcre
            '.Edit
            !SALDAcre = !SALDAcre - PARCAN#
            .Update
            '
            DocOrCor$ = !CodiCom_Cor
            TotOrig# = !IBruComp
            IIVAORIG# = 0
            On Error Resume Next
            IvaOrig# = !IIvaComp
            On Error GoTo 0
            '
            ICAN# = ICAN# - PARCAN#
            If ICAN# < 0.005 Then
              Exit Do
            End If
            .MoveNext
           Loop
         End If
         On Error GoTo 0
       End With
       SACCCREPENP.Close
       Set SACCCREPENP = Nothing
       '
       DocOrLar$ = TRIM$(Left$(Z$, 18))
       ImpApli# = CANCE#
       Call GRAPLIPAG
     End If
   Next U&
   '
   OKX% = 4       ' grabó aplicaciones
   '
   'Grabo Detalle de Facturación en Archivo-Base Central
   If Check2.Value > 0 Then
      ZZ$ = REGBLAN("DETFACOM")
      ULTREDET& = ULTREG("DETFACOM")
      CONTI1& = 0
      For j& = 1 To ULTREG("!DETFACOM")
       DETO$ = REGLEIDO("!DETFACOM", j&)
       '
       CAFA1 = CVD(Mid$(DETO$, 95, 8))
       If CAFA1 > 0.005 Then
         ZZ$ = DETO$
         '
         NUFA1$ = TRIM(Text11)
         Call REPLA(ZZ$, NUFA1$, 3, 18) 'NUMERO FACTURA
         Call REPLA(ZZ$, MKI$(FFRE%), 21, 2) 'F.EMISION
         Call REPLA(ZZ$, MKI$(FF%), 23, 2) 'FECHA CONTABLE
         '
         CONTI1& = CONTI1& + 1
         Call GRAREG("DETFACOM", ZZ$, ULTREDET& + CONTI1&)
       End If
      Next j&
      Call SETULTREG("DETFACOM", ULTREDET& + CONTI1&)
   End If
   '
84 Call CIERRARCH("APLIPAG")
   Call LIBARCH("APLIPAG")
   Call CIERRARCH("CUECOSS")
   Call LIBARCH("CUECOSS")
   Call CIERRARCH("PROVED2")
   Call LIBARCH("PROVED2")
   Call CIERRARCH("CUECORR")
   Call LIBARCH("CUECORR")
   Call CIERRARCH("CAJABANC")
   Call LIBARCH("CAJABANC")
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   List1.Clear
   List2.Clear
   Label4.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Text11.TEXT = ""
'   ITOGRA.TEXT = Space$(11)
   'TOEXEN.TEXT = ""
   'TOTIVA.TEXT = ""
   'TORETIV.TEXT = ""
   'TOTFAC.Caption = ""
   For KKI% = 0 To 8
      FEVEN(KKI%).TEXT = ""
      IVENCI(KKI%).TEXT = ""
   Next KKI%
   T16$ = Text16: T17$ = Text17
   Call BLAFACOM
   Text16 = T16$: Text17 = T17$  ' PARA QUE MANTENGA EN NUMERO DE CAI
   VACON% = 0
   '
   ' PEDIDO POR TELEBIT QUE VUELVA A MONEDA PESOS - 24/04/06
   MONEMI% = 1
   Label29 = DMONE$(MONEMI%)
   Text18 = FORMATNUM(TCMONE#(MONEMI%), "F11.4")
   Label29.ForeColor = RGB(0, 0, 255)
   '
   Screen.MousePointer = 1
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
   '
End Sub
'
Sub SIGUEFACO1()
    '
    Dim NRO$(10)
    Command12.Enabled = True
    Label25.Enabled = True
    If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then
      Command12.Enabled = False
      Label25.Enabled = False
    End If
    '
    Screen.MousePointer = 11
    '
    FE$ = TRIM$(Text9.TEXT)
    FF% = FECHANUM(FE$)
    HOI% = HOY(HO$)
    If FF% < 1 Or FF% > HOI% + 1 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
    '
    If DIENTRE%(FF%, HOI%) > 55 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
   '
    Call CIERRARCH("*.*")
    '
    CUITT$ = TRIM$(Text13)
    If VALICUIT%(CUITT$) <> 1 Then
        Call MENSERR(24, "Imposible Procesar Documento.\Número de C.U.I.T. no Válido")
        Text1.SetFocus
        Screen.MousePointer = 1
        Exit Sub
      End If
    
    '
    LETRAFA$ = Label20.Caption
    If TIDOCOM$ = "CO" Then
      If Len(TRIM$(Text11)) < 4 Then
        Screen.MousePointer = 1
        Text11.SetFocus
        Exit Sub
      End If
    End If
    '
    If TIDOCOM$ <> "CO" Then    ' evalua numero ingresado
      NROX$ = TRIM$(Text11.TEXT)
      While Len(NROX$) > 0
         If Asc(Left$(NROX$, 1)) < 48 Or Asc(Left$(NROX$, 1)) > 57 Then
              NROX$ = Mid$(NROX$, 2)
            Else
              GoTo 101
         End If
      Wend
      '
101   If NROX$ = "" Then
         Screen.MousePointer = 1
         Text11.SetFocus
         Exit Sub
      End If
      '
      SUC% = 0
      C1% = 1: For c% = 1 To 10:  NRO$(c%) = "": Next c%
      '
      For c% = 1 To Len(NROX$)
         C2% = Asc(Mid$(NROX$, c%, 1))
         If C2% = 32 Then GoTo 102
         If (C2% < 48 Or C2% > 57) Then C1% = C1% + 1: GoTo 102
         NRO$(C1%) = NRO$(C1%) + Chr$(C2%)
102   Next c%
      '
      If C1% = 1 Then
         SUC% = 0
         VD# = XVALO(NRO$(1))
         NRO$(2) = NRO$(1)
         NRO$(1) = "0"
      End If
      '
      If XVALO(NRO$(2)) < 1 Or Len(NRO$(1)) > 4 Or Len(NRO$(2)) > 8 Then
          Call MENSERR(24, "Numero de Documento no Permitido")
          Screen.MousePointer = 1
          Text11.SetFocus
          Exit Sub
      End If
      '
      While Len(NRO$(1)) < 4: NRO$(1) = "0" + NRO$(1): Wend
      While Len(NRO$(2)) < 8: NRO$(2) = "0" + NRO$(2): Wend
      '
      If LETRAFA$ <> "M" Then LETRAFA$ = "A"
      PVCL% = PIVACLI%(-1 * NC%)
      If PVCL% = 0 Or PVCL% = 2 Or PVCL% = 6 Or PVCL% = 4 Then LETRAFA$ = "C"
      If PVCL% = 5 Then LETRAFA$ = "E"
      NROX$ = TIDOCOM$ + "-" + LETRAFA$ + "-" + NRO$(1) + "-" + NRO$(2)
      Text11.TEXT = NROX$
      Text11.Refresh
      Label20.Caption = LETRAFA$
      Label20.Refresh
      '
    End If
    '
    ' verifica repeticion del numero ******************************************
    Screen.MousePointer = 11
    NROY1$ = NROX$
    If TIDOCOM$ = "CO" Then           ' evalua numero ingresado
      NROY1$ = "-" + TRIM$(Text11)
      While Len(NROY1$) < 13: NROY1$ = "0" + NROY1$: Wend
      NROY1$ = "CO-" + TRIM$(Label20.Caption) + "-" + AJUSTI$(NROY1$, 13)
    End If
    
    Call REPLA(NROY1$, "%", 4, 1) '(en sql like %) para q tome por iguales facturas q difieren en la letra
    '
    Call APERBASDAT("CABAN")
    
    Dim MIRS As ADODB.Recordset
    Set MIRS = New ADODB.Recordset
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "AND CodiCom_Lar LIKE '" & NROY1$ & "' "
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With MIRS
    If Not (.EOF And .BOF) Then
         Call MENSERR(24, "Documento ya Registrado")
         Text11 = ""
         Text1.SetFocus
         Exit Sub
      End If
    End With
    MIRS.Close
    Set MIRS = Nothing

95  On Error GoTo 0
    Screen.MousePointer = 1
    Screen.MousePointer = 1
    Call SIGUEFACO2
    '
End Sub
'

Sub SIGUEFACO2()
    '
    '!*!=NUEVO PARA GRABACION DE CONCEPTO EN FACT. PROV.
    '
'''    Call BLANARCH("!CARICOG")
'''    For u& = 1 To List1.ListCount
'''        List1.ListIndex = u& - 1
'''        Z$ = List1.TEXT
'''        CUE% = XVALO(Left$(Z$, 5))
'''        '\\\OT-08-989-29/10/08-FL
'''        'DENOX$ = Mid$(Z$, 8,30)+ CODCUE$(CUE%)
'''        DENOX$ = Mid$(Z$, 8, 29) + " " + CODCUE$(CUE%)
'''        '\\\OT-08-989-FIN-FL
'''        IPA# = XVALO(Mid$(Z$, 40, 12))
'''        CENCOS$ = TRIM$(Mid$(Z$, 65, 8))
'''        CONCEP$ = TRIM$(Mid$(Z$, 75, 8)) '!*!
'''        '
'''        Y$ = REGBLAN$("!CARICOG")
'''        Call REPLA(Y$, MKI$(CUE%), 1, 2)
'''        Call REPLA(Y$, DENOX$, 3, 48)
'''        Call REPLA(Y$, MKD$(IPA#), 51, 8)
'''        Call REPLA(Y$, CENCOS$, 59, 8)
'''        Call REPLA(Y$, CONCEP$, 67, 8) '!*!
'''        Call REGAPP("!CARICOG", Y$)
'''    Next u&
'''    Call CIERRARCH("!CARICOG")
'''
'''    '
'''200 ATRI$ = "/NC"
'''    ATRI$ = ATRI$ + "/TITUPAN=Imputacion Contable"
'''    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(250 + Top + Picture1.Top))
'''    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
'''    '
'''    FOVENTA$ = "CARIFACG"
'''    If EXISTE%(LUDAT$ + "TABITCO.DAT") > 0 Then
'''       FOVENTA$ = "CARIFCON"
'''    End If
'''    '
'''    VTB% = VENTABU%(FOVENTA$ + ATRI$) '!*!
'''    If VTB% >= 0 Then
'''       RECUEP07.LINDICE.Clear
'''       List1.Clear
'''       TOTIMPU# = 0
'''       For u& = 1 To ULTREG&("!CARICOG")
'''          Y$ = REGLEIDO$("!CARICOG", u&)
'''          CUEB$ = Left$(Y$, 2)
'''          If CVI(CUEB$) > 0 Then
'''             Z$ = Space$(128)
'''             Call REPLA(Z$, CSTRING$(CUEB$, 1, 5, 0, 2), 1, 5)
'''             DECUE$ = Mid$(Y$, 3, 48) ' ECOARCH$("CUECON", CUEB$)
'''             'DECUE$ = TRIM$(TRIM$(Mid$(DECUE$, 31, 12)) + " " + TRIM$(Left$(DECUE$, 30)))
'''             Call REPLA(Z$, DECUE$, 8, 32)
'''             IMBB$ = Mid$(Y$, 51, 8)
'''             If CVD(IMBB$) < 0 Then
'''                Call MENSERR(24, "No se Admiten Valores Negativos")
'''                GoTo 200
'''             End If
'''             '
'''             Call REPLA(Z$, CSTRING$(IMBB$, 3, 13, 2, 2), 38, 13)
'''             CENCOS$ = Mid$(Y$, 59, 8)
'''             '\\\OT-06-0244-WAR-28/06/06///
'''             FCON1% = FECHANUM(Text9.TEXT)
'''             If TRIM$(CENCOS$) <> "" Then
'''              If VALIPROY%(CENCOS$, FCON1%) = (-1) Then
'''                 Call MENSERR(24, "Fecha Contable se Encuentra fuera del Límite\ de Vigencia del Centro de Costos: " + CENCOS$)
'''                 GoTo 200
'''              End If
'''             End If
'''             '\\\OT-06-0244-WAR-FIN///
'''             Call REPLA(Z$, CENCOS$, 65, 8)
'''             '
'''             CONCEP$ = Mid$(Y$, 67, 8) '!*!
'''             Call REPLA(Z$, CONCEP$, 75, 8) '!*!
'''             '
'''             TOTIMPU# = TOTIMPU# + CVD(IMBB$)
'''             If CVD(IMBB$) >= 0.005 Then
'''               RECUEP07.LINDICE.AddItem Z$
'''             End If
'''          End If
'''       Next u&
'''       '
'''       For u& = 1 To RECUEP07.LINDICE.ListCount
'''          RECUEP07.LINDICE.ListIndex = u& - 1
'''          Z$ = RECUEP07.LINDICE.TEXT
'''          List1.AddItem Z$
'''       Next u&
'''       '
'''       For KKI% = 1 To 6
'''          IMPOGRA(KKI%) = Space$(11)
'''          IMPOIVA(KKI%) = Space$(8)
'''       Next KKI%
'''       '
'''       NC1% = (-1) * XVALO(Text1)
'''       PVCL% = PIVACLI%(NC1%)
'''       Label36 = FORMATNUM$(TOTIMPU#, "F11.2") ' TOTIMPU# 27/04/07
''''       If (LETRAFA$ = "A" Or LETRAFA$ = "M") And PVCL% <> 4 Then  ' CASO DE PROVEEDORES EXENTOS
''''           ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
''''           ITOGRA.Refresh
''''         Else
''''           ITOGRA.Caption = Space$(11)
''''           IMPOGRA(4).TEXT = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
''''           IMPOGRA(4).Refresh
''''       End If
'''    End If
'''    List1.Refresh
'''    IMPOGRA(1).SetFocus
'''    Label36 = FORMATNUM$(TOTIMPU#, "F11.2")
'''    If TOTIMPU# < 0.01 Then
'''       Text1.SetFocus
'''    End If
    '
End Sub

Sub SIGUEFACO3()
   '
   For UI% = 0 To 6
     IMPOGRA(UI%).TEXT = Space$(11)
     IMPOIVA(KKI%) = Space$(8)
   Next UI%
   If LETRAFA$ = "A" Or LETRAFA$ = "M" Then
        IMPOGRA(0).TEXT = Space$(11)
        IMPOGRA(0).TEXT = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
        IMPOGRA(0).Refresh
      Else
        IMPOGRA(0).TEXT = Space$(11)
        IMPOGRA(4).TEXT = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
        IMPOGRA(4).Refresh
   End If
   TOEXEN# = XVALO(IMPOGRA(4).TEXT) + XVALO(IMPOGRA(5).TEXT)
   IMPOGRA(0).SetFocus
   TONEGRA# = TOTIMPU# - TOEXEN#
'   ITOGRA.TEXT = CSTRING$(MKD$(TONEGRA#), 4, 11, 2, 2)
'   ITOGRA.Refresh
   '
End Sub


Sub SIGUEFACO5()
    '
    If List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!CARGAPLI")
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        NROX$ = Mid$(Z$, 1, 20)
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(XVALO(Mid$(Z$, 32, 13)))
        FEVEI% = FECHANUM(Mid$(Z$, 47, 8))
        SALD# = CDbl(XVALO(Mid$(Z$, 56, 13)))
        ACRE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 1, 18)
        REG2& = XVALO(Mid$(Z$, 119, 8))
        IRE% = XVALO(Mid$(Z$, 127, 2))
        '
        Y$ = REGBLAN$("!CARGAPLI")
        Call REPLA(Y$, NROX$, 1, 20)
        Call REPLA(Y$, MKI$(FEMII%), 21, 2)
        Call REPLA(Y$, MKD$(IPED#), 23, 8)
        Call REPLA(Y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(Y$, MKD$(SALD#), 33, 8)
        Call REPLA(Y$, MKD$(ACRE#), 41, 8)
        Call REPLA(Y$, DOCU$, 49, 18)
        Call REPLA(Y$, MKS$(REG2&), 67, 4)
        Call REPLA(Y$, MKI$(IRE%), 71, 2)
        '
        Call REGAPP("!CARGAPLI", Y$)
    Next U&
    Call CIERRARCH("!CARGAPLI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Pagos"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame8.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    '
    VTB% = VENTABU%("CARGAPLI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        List2.Clear
        For U& = 1 To ULTREG&("!CARGAPLI")
            Y$ = REGLEIDO$("!CARGAPLI", U&)
            NROX$ = Mid$(Y$, 1, 20)
            If TRIM$(NROX$) = "Anticipo no Aplicado" Then
               HOI% = HOY(HO$)
               ACRE# = CVD(Mid$(Y$, 41, 8))
               Call REPLA(Y$, MKI$(HOI%), 31, 2)
               Call REPLA(Y$, MKD$(ACRE#), 23, 8)
               Call REPLA(Y$, MKD$(ACRE#), 33, 8)
               Call GRAREG("!CARGAPLI", Y$, U&)
            End If
            FEMI = CVI(Mid$(Y$, 21, 2))
            IPED# = CVD(Mid$(Y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 13, 2, 2)
            FEVE = CVI(Mid$(Y$, 31, 2))
            SALD# = CVD(Mid$(Y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 13, 2, 2)
            ACRE# = CVD(Mid$(Y$, 41, 8))
            '
            If ACRE# < 0 Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If U& > 1 Then
              If ACRE# > SALD# Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                 GoTo 200
              End If
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(Y$, 49, 18)
            REG2& = CVS(Mid$(Y$, 67, 4))
            IRE% = CVI(Mid$(Y$, 71, 2))
            REGTEX$ = CSTRING$(MKS$(REG2&), 3, 8, 0, 2)
            IRETEX$ = CSTRING$(MKI$(IRE%), 1, 2, 0, 2)
            '
            Z$ = Space$(128)
            Call REPLA(Z$, NROX$, 1, 20)
            Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(Z$, IMPTEX$, 32, 13)
            Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
            Call REPLA(Z$, SALTEX$, 56, 13)
            Call REPLA(Z$, ACRTEX$, 70, 14)
            Call REPLA(Z$, REGTEX$, 119, 8)
            Call REPLA(Z$, IRETEX$, 127, 2)
            List2.AddItem Z$
            '
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(Z$, 70, 14)))
    Next U&
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    '
    List2.Refresh
    VACON% = 0
    '
End Sub
'
Sub SIGUEBORE(NPROX%)
    '
    FORSELBRE.LIMATSEL.Clear
    '
    'NUEVO FILTRO DE FECHA
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    '
    DESDEFE = CDbl(CVDAT(DES%))
    HASTAFE = CDbl(CVDAT(HAS%))
    NBRVIE$ = "": NOCOVIE$ = ""
    DESDEF1 = CVDAT(DES%)
    HASTAF1 = CVDAT(HAS%)
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + "WHERE NumProv = " & NPROX% & " "
    SQLX$ = SQLX$ + "AND FechRecep >= '" & DESDEF1 & "' "
    SQLX$ = SQLX$ + "AND FechRecep <= '" & HASTAF1 & "' "
    SQLX$ = SQLX$ + "ORDER BY NumeBoRe ASC "
    'Set BOLDET = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim BOLDET As ADODB.Recordset
    Set BOLDET = New ADODB.Recordset
24  On Error Resume Next
    BOLDET.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 24
    End If
    On Error GoTo 0

    With BOLDET
       On Error Resume Next
       .MoveFirst
       If Err > 1 Then
          On Error GoTo 0
          GoTo 126
       End If
       On Error GoTo 0
       Do While Not .EOF
           NFAC$ = !NumFac_Prov
           CANAP = !CanToApro
           FERECEP = CVINT(!FechRecep)
           '
           If TRIM(NFAC$) = "" Then
              If CANAP > 0.005 Then
                 '
                 NBR$ = TRIM$(Str$(!NumeBoRe))
                   While Len(NBR$) < 6: NBR$ = "0" + NBR$: Wend ' para formatear
                 NRP$ = TRIM$(!NumRem_Prov)
                 NOCOM$ = TRIM$(Str$(!NumOC))
                 FBORE = FERECEP
                 BORE$ = "BR." & NBR$ & " - " & FECHATEX(FBORE) & " - Rto." & NRP$
                 If Val(NOCOM$) > 0 Then
                   BORE$ = BORE$ + " - OC." + NOCOM$
                   NOCOVIE$ = NOCOM$
                 End If
                 '
                 With FORSELBRE
                   If NBRVIE$ <> NBR$ Then
                       .LIMATSEL.AddItem BORE$
                       NBRVIE$ = NBR$
                     ElseIf NOCOM$ <> NOCOVIE$ Then
                       If Val(NOCOM$) > 0 Then
                         BORE$ = .LIMATSEL.List(.LIMATSEL.ListCount - 1)
                         .LIMATSEL.RemoveItem (.LIMATSEL.ListCount - 1)
                         BORE$ = BORE$ + ", " + NOCOM$
                         .LIMATSEL.AddItem BORE$
                         NOCOVIE$ = NOCOM$
                       End If
                   End If
                 End With
                 '
              End If
           End If
           '
       .MoveNext
       Loop
    End With
    Set BOLDET = Nothing
    '
  ' hay que prever el caso de que no haya boletas aprobadas sin facturar !!!
126 Screen.MousePointer = 1
    If FORSELBRE.LIMATSEL.ListCount < 1 Then
      Call MENSERR(24, "No Hay Boletas de Recepción\Pendientes de Facturación\para la Cuenta Elegida.")
      GoTo 25
    End If
  '
  FORSELBRE.Show 1
  HAYSEL% = 0
  For j& = 0 To FORSELBRE.LIMATSEL.ListCount - 1
    If FORSELBRE.LIMATSEL.Selected(j&) = True Then
      HAYSEL% = 1
    End If
  Next j&
  '
  If HAYSEL% > 0 Then 'SI HAY SELECCION
25  With FOFACOBR07
      .Text1 = Text1
      .Text2 = Text2
      .Text3 = Text3
      .Text4 = Text4
      .Text5 = Text5
      .Text6 = Text6
      .Text13 = Text13
      .Text9 = Text9
      .Text11 = Text11
      '
      With FORSELBRE
          J1& = 0
          For j& = 1 To FORSELBRE.LIMATSEL.ListCount
            If FORSELBRE.LIMATSEL.Selected(j& - 1) = True Then
              li1$ = FORSELBRE.LIMATSEL.List(j& - 1)
              'SACO EL NUMERO DE BOLETA DE RECEPCION
              NBORE& = Val(Mid$(li1$, 4, 6))
              On Error GoTo 0
              '
              SQLX$ = "SELECT * FROM BOLRECEP "
              SQLX$ = SQLX$ + "WHERE NumProv = " & NPROX% & " "
              SQLX$ = SQLX$ + "AND NumeBoRe = " & NBORE& & " "
              SQLX$ = SQLX$ + "AND FechRecep >= '" & DESDEF1 & "' "
              SQLX$ = SQLX$ + "AND FechRecep <= '" & HASTAF1 & "' "
              'Set BOLDET = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
              Dim BOLDET1 As ADODB.Recordset
              Set BOLDET1 = New ADODB.Recordset
26            On Error Resume Next
              BOLDET1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
 
              If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                 On Error GoTo 0
                 Call CapturaErrorOpen
                 GoTo 26
              End If
              On Error GoTo 0
              '
              With BOLDET1
                On Error Resume Next
                .MoveFirst
                If Err Then
                   On Error GoTo 0
                   GoTo 39
                End If
                On Error GoTo 0
                Do While Not .EOF
                   NFAC$ = !NumFac_Prov
                   CANAP1# = !CanToApro
                   FERECEP = CVINT(!FechRecep)
                   '
                   Y$ = REGBLAN$("!DETFACOM")
                   '
                   If TRIM(NFAC$) = "" Then
                      If CANAP1# > 0.005 Then
                         '
                         NRP$ = TRIM$(!NumRem_Prov)
                         FBORE = FERECEP
                         CIXI% = !Cod_Inte
                         CARE# = !CantNom
                         NUOC& = !NumOC
                         LO1$ = TRIM(!IdenLote)
                         LP1$ = TRIM(!LOTECOL_PROV)
                         'SACO EL NUMERO DE ITEM
                         ACA% = 0
                         For JJJ& = 1 To Len(LO1$)
                            JX$ = Mid$(LO1$, JJJ&, 1)
                            If JX$ = "-" Then
                              ACA% = ACA% + 1
                            End If
                            If ACA% = 2 Then Exit For
                         Next JJJ&
                         NUOIT% = Val(Mid(X$, 69 + JJJ&))
                         'PRECIO UNITARIO DE COMPRA
                         If NUOC& > 0 Then
                             REGOC$ = ULTIPRE$(NPROX%, CIXI%, NUOC&)
                             COEFU = CVS(Mid$(REGOC$, 63, 4))
                             PREUN1# = CVD(Mid$(REGOC$, 67, 8))
                             If COEFU > 1 Then PREUN1# = PREUN1# / COEFU
                           Else
                             PREUN1# = 0
                         End If
                         'BUSCO SI HAY CANTIDAD YA FACTURADA
                         CANTIYAFAC# = 0
                         For JJX& = ULTREG&("DETFACOM") To 1 Step -1
                           ZZZ$ = REGLEIDO("DETFACOM", JJX&)
                           If CVS(Mid$(ZZZ$, 27, 4)) = NBORE& Then
                             If CVI(Mid$(ZZZ$, 25, 2)) = CIXI% Then
                               If NUOIT% = CVS(Mid$(ZZZ$, 31, 4)) Then
                                 CANTFAC = CVD(Mid$(ZZZ$, 95, 8))
                                 CANTIYAFAC# = CANTIYAFAC# + CANTFAC
                               End If
                             End If
                           End If
                         Next JJX&
                         QPENFAC# = CANAP1# - CANTIYAFAC#
                         If QPENFAC# <= 0 Then GoTo 35
                           'IMPORTE
                         IMPOR1# = PREUN1# * CANAP1#
                           'AGREGO AL ARCHIVO
                         Call REPLA(Y$, MKI$(NPROX%), 1, 2)
                         Call REPLA(Y$, MKI$(CIXI%), 25, 2)
                         Call REPLA(Y$, MKS$(NBORE&), 27, 4)
                         Call REPLA(Y$, MKS$(NUOIT%), 31, 4)
                         Call REPLA(Y$, LO1$, 35, 16)
                         Call REPLA(Y$, LP1$, 51, 16)
                         Call REPLA(Y$, MKS$(NUOC&), 67, 4)
                         Call REPLA(Y$, MKD$(CARE#), 71, 8)
                         Call REPLA(Y$, MKD$(CANAP1#), 79, 8) 'Q.aprobada
                         Call REPLA(Y$, MKD$(QPENFAC#), 87, 8) 'Q.PENDIENTE DE FAC
                         Call REPLA(Y$, MKD$(PREUN1#), 103, 8) 'P.U. O/C
                         Call REPLA(Y$, MKD$(PREUN1#), 111, 8) 'P.U. FACTURADO
                         Call REPLA(Y$, MKD$(IMPOR1#), 135, 8)
                         '
                         J1& = J1& + 1
                         Call GRAREG("!DETFACOM", Y$, J1&)
                      End If
                   End If
                   '
35              .MoveNext
                Loop
              End With
              Set BOLDET1 = Nothing

            End If
39        Next j&
          Call SETULTREG("!DETFACOM", J1&)
      End With

      If J1& < 1 Then
        Call COMUNI("No se Encontraron Recepciones\Pendientes de Facturación\para el Set de Recepciones Elegidas.")
      End If
      'AGREGAR AL LIST DE FOFACOBR
      FOFACOBR07.List1.Clear
      IMPOTOT# = 0
      For j& = 1 To ULTREG&("!DETFACOM")
        XY$ = REGLEIDO$("!DETFACOM", j&)
        NPRO1 = CVI(Mid$(XY$, 1, 2))
        CIXI1% = CVI(Mid$(XY$, 25, 2))
        '
        NBORE1 = CVS(Mid$(XY$, 27, 4))
        NUOC1& = CVS(Mid$(XY$, 67, 4))
        QFAC1 = CVD(Mid$(XY$, 95, 8))
        QPENFAC1 = CVD(Mid$(XY$, 87, 8))
        PUFAC1 = CVD(Mid$(XY$, 111, 8))
        IMPOR1# = PUFAC1 * QFAC1
        IMPOTOT# = IMPOTOT# + IMPOR1#
        '
        NPRO2$ = FORMATNUM(NPRO1, "F5")
        CIXI2$ = CODEXT(CIXI1%)
        DESCRI2$ = Left$(DESCRIT0(CIXI1%), 15)
        NBORE2$ = FORMATNUM(NBORE1, "F5.0")
        NUOC2$ = FORMATNUM(NUOC1&, "F5.0")
        QPENFAC2$ = FORMATNUM(QPENFAC1, "F10.1")
        QFAC2$ = FORMATNUM(QFAC1, "F10.1")
        PUFAC2$ = FORMATNUM(PUFAC1, "F10.4")
        IMPOR2$ = FORMATNUM(IMPOR1#, "F10.2")
        '
        li1$ = Space(64)
        '
        Call REPLA(li1$, CIXI2$, 1, 16)
        Call REPLA(li1$, DESCRI2$, 17, 10)
        Call REPLA(li1$, NBORE2$, 29, 5)
        Call REPLA(li1$, NUOC2$, 36, 5)
        Call REPLA(li1$, QPENFAC2$, 43, 10)
        Call REPLA(li1$, QFAC2$, 55, 10)
        Call REPLA(li1$, PUFAC2$, 67, 10)
        Call REPLA(li1$, IMPOR2$, 79, 11)
        '
        .List1.AddItem li1$
      Next j&

      '
      .Text12 = ""
      .LBLIMPO = FORMATNUM(IMPOTOT#, "F11.2")
      .Show 1
    End With
    Set BOLDET1 = Nothing

  End If
  '
End Sub
'
Function ULTIPRE$(NPROX%, CIXI%, NUOC&)
   '
   ' trae el registro de OCOMDETA correspondiente a la
   ' última compra no anulada del artículo y proveedor
   ' con precio registrado mayor que cero
   '
   Call APERBASDAT("COMPRAS")
   VUELTA% = 0
   '
20 UCO$ = REGBLAN$("OCOMDETA")
   Call REPLA(UCO$, MKI$(NPROX%), 7, 2)
   Call REPLA(UCO$, MKI$(CIXI%), 9, 2)
   '
   SQLX$ = "SELECT * FROM OCOMDETA "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
   SQLX$ = SQLX$ + "AND Npro_Ocom = " & NPROX% & " "
   SQLX$ = SQLX$ + "AND Nume_Ocom = " & NUOC& & " "
   If VUELTA% = 0 Then SQLX$ = SQLX$ + "AND Pre_Unit >= 0.00005 "
   SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
   'Set OcomdetaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim OcomdetaTemp As ADODB.Recordset
   Set OcomdetaTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))

   On Error Resume Next
   With OcomdetaTemp
     .MoveFirst
     
     If Err < 1 Then
         On Error GoTo 0
         Call REPLA(UCO$, MKS$(!Nume_Ocom), 1, 4)
         Call REPLA(UCO$, MKI$(FECHANUM(Format(!Femi_Ocom, "dd/mm/yy"))), 5, 2)
         Call REPLA(UCO$, !Uni_Mas, 11, 4)
         Call REPLA(UCO$, !Uni_Com, 59, 4)
         Call REPLA(UCO$, MKS$(!Coef_Unids), 63, 4)
         Call REPLA(UCO$, Chr$(!Mone_Emis), 122, 1)
         Call REPLA(UCO$, MKD$(!Pre_Unit), 67, 8)
         Call REPLA(UCO$, MKS$(!Desc_Item), 75, 4)
       Else
         On Error GoTo 0
         If VUELTA% < 1 Then
           VUELTA% = VUELTA% + 1
           GoTo 20
         End If
     End If
   End With
   '
   ULTIPRE$ = UCO$
   '
End Function

















' AQUI GRABAR CAJABANC PARA FACTURA
'   CALL Aperbasdat("CABAN")
'   With CajaBanc
'     .AddNew
'     !CodiCom_Cor = NFCX$
'     !CodiCom_Lar = NFCY$
'     !RefeCom = RASOCLI$(NC%)
'     !FechConta = CVDAT(FFI%)   ' fecha de emision
'     !FechAcred = CVDAT(FFI%) ' fecha de acreditacion
'     !CueContaI = CUECOCLI%(NC%)
'     !NuClien = NC%
'     !NuProve = 0
'     !TipoMovim = "FCC"
'     !OpciMovim = OPMOVIX$
'     !VariMovim = Label13
'     !DescriMov = Label7 + " " + Label13.Caption + " " + Label6.Caption + " " + Label12.Caption + "-" + Text11.TEXT
'     !FechEmisi = CVDAT(FFI%)
'     !CodiEmpr = 0
'     !CodiSucu = 0
'     !NumeTalo = XVALO(Label12)
'     !Numecomp = CDbl(Text11)
'     !STATUS = 0
'     !ValAbsComp = TOTAGEN#
'     If TIPO% <> 6 Then
'         !IdebeComp = TOTAGEN#
'         !IHabeComp = 0
'       Else
'         !IdebeComp = 0
'         !IHabeComp = TOTAGEN#
'     End If
'     !ObserGen = RASOCLI$(NC%)
'     !IdenAlta = IDACTU$           ' FECHA - HORA - USUARIO - TERMINAL
'     !IdenBaja = String$(6, 0)
'     '
'     .Update
'     '
'   'graba venta neta
'     TOTANETX# = XVALO(Total(2))
'     If TOTANETX# >= 0.005 Then
'       .AddNew
'       !CodiCom_Cor = NFCX$
'       !CodiCom_Lar = NFCY$
'       !RefeCom = RASOCLI$(NC%)
'       !FechConta = CVDAT(FFI%)   ' fecha de emision
'       !FechAcred = CVDAT(FFI%) ' fecha de acreditacion
'       !CueContaI = QQQ%
'       !NuClien = NC%
'       !NuProve = 0
'       !TipoMovim = "MERC"
'       !OpciMovim = "VEN": If TIPO% = 6 Then !OpciMovim = "DEV"
'       !VariMovim = "COD": If InStr(UCase$(Label6), "LIB") > 0 Then OpciMovim = "LIB"
'       !DescriMov = "Mercaderia Vendida " + NFCX$
'       !FechEmisi = CVDAT(FFI%)
'       !CodiEmpr = 0
'       !CodiSucu = 0
'       !NumeTalo = XVALO(Label12)
'       !Numecomp = CDbl(Text11)
'       !STATUS = 0
'       !ValAbsComp = TOTANETX#
'        If TIPO% <> 6 Then
'            !IdebeComp = 0
'            !IHabeComp = TOTANETX#
'          Else
'            !IdebeComp = TOTANETX#
'            !IHabeComp = 0
'        End If
'        !ObserGen = RASOCLI$(NC%)
'        !IdenAlta = IDACTU$           ' FECHA - HORA - USUARIO - TERMINAL
'        !IdenBaja = String$(6, 0)
'        '
'        .Update
'        '
'     End If
'     '
'   'graba iva inscripto
'     TOTANETX# = XVALO(Total(3)) + XVALO(Total(4))
'     If TOTANETX# >= 0.005 Then
'       .AddNew
'       !CodiCom_Cor = NFCX$
'       !CodiCom_Lar = NFCY$
'       !RefeCom = RASOCLI$(NC%)
'       !FechConta = CVDAT(FFI%)   ' fecha de emision
'       !FechAcred = CVDAT(FFI%) ' fecha de acreditacion
'       '
'       CUEIVA$ = CONTROL$("SVD451S", "CUEIVA")
'       QQX% = CUECOINT%(CUEIVA$)
'       !CueContaI = QQX%
'       '
'       !NuClien = NC%
'       !NuProve = 0
'       !TipoMovim = "IVA"
'       !OpciMovim = "DEFI"
'       !VariMovim = "TBAS"
'       !DescriMov = "Iva Debito Fiscal " + NFCX$
'       !FechEmisi = CVDAT(FFI%)
'       !CodiEmpr = 0
'       !CodiSucu = 0
'       !NumeTalo = XVALO(Label12)
'       !Numecomp = CDbl(Text11)
'       !STATUS = 0
'       !ValAbsComp = TOTANETX#
'        If TIPO% <> 6 Then
'            !IdebeComp = 0
'            !IHabeComp = TOTANETX#
'          Else
'            !IdebeComp = TOTANETX#
'            !IHabeComp = 0
'        End If
'        !ObserGen = RASOCLI$(NC%)
'        !IdenAlta = IDACTU$           ' FECHA - HORA - USUARIO - TERMINAL
'        !IdenBaja = String$(6, 0)
'        '
'        .Update
'        '
'     End If
'     '
'   'graba iva NO inscripto
'     TOTANETX# = XVALO(Total(5))
'     If TOTANETX# >= 0.005 Then
'       .AddNew
'       !CodiCom_Cor = NFCX$
'       !CodiCom_Lar = NFCY$
'       !RefeCom = RASOCLI$(NC%)
'       !FechConta = CVDAT(FFI%)   ' fecha de emision
'       !FechAcred = CVDAT(FFI%) ' fecha de acreditacion
'       '
'       CUEIVA$ = CONTROL$("SVD451S", "CUEADU")
'       QQX% = CUECOINT%(CUEIVA$)
'       !CueContaI = QQX%
'       '
'       !NuClien = NC%
'       !NuProve = 0
'       !TipoMovim = "IVA"
'       !OpciMovim = "DEFI"
'       !VariMovim = "TRNI"
'       !DescriMov = "IVA Resp.No Inscripto " + NFCX$
'       !FechEmisi = CVDAT(FFI%)
'       !CodiEmpr = 0
'       !CodiSucu = 0
'       !NumeTalo = XVALO(Label12)
'       !Numecomp = CDbl(Text11)
'       !STATUS = 0
'       !ValAbsComp = TOTANETX#
'        If TIPO% <> 6 Then
'            !IdebeComp = 0
'            !IHabeComp = TOTANETX#
'          Else
'            !IdebeComp = TOTANETX#
'            !IHabeComp = 0
'        End If
'        !ObserGen = RASOCLI$(NC%)
'        !IdenAlta = IDACTU$           ' FECHA - HORA - USUARIO - TERMINAL
'        !IdenBaja = String$(6, 0)
'        '
'        .Update
'        '
'     End If
'     '
'   End With
'   ' TERMINA DE GRABAR CAJABANC
'   '
'
'
Private Sub Label26_Click()
   Call SELECHO("INDIACRE")
   CPROX$ = TRIM$(VALACT1$("INDIACRE"))
   If CPROX$ <> "" Then
     NCXI% = NUMEPRO%(CPROX$)
     If NCXI% > 0 Then
       If REGICLI&((-1) * NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
         Text2.TEXT = RASOCLI$((-1) * NC%)
         Text3.TEXT = DOMICLI$((-1) * NC%)
         Text4.TEXT = CPOSCLI$((-1) * NC%)
         Text5.TEXT = LOCACLI$((-1) * NC%)
         Text13.TEXT = CUITCLI$((-1) * NC%)
           If LETRAFA$ <> "M" Then LETRAFA$ = "A"
           PVCL% = PIVACLI%(-1 * NC%)
           If PVCL% = 0 Or PVCL% = 2 Or PVCL% = 6 Or PVCL% = 4 Then LETRAFA$ = "C"
           If PVCL% = 5 Then LETRAFA$ = "E"
         Label20 = LETRAFA$
         If Left$(Text2.TEXT, 6) = "VARIOS" Then
             Text13.SetFocus
           Else
             Text6.SetFocus
         End If
       End If
     End If
   End If
End Sub

Private Sub Label26_DblClick()
   Call Label26_Click
End Sub

Private Sub BUSCACAI(NC%)
    '
    Call APERBASDAT("CABAN")
    '
    'jandy sql
    ' obtengo el ultimo nro cai (no necesito el movelast)
    'recupero un unico registro aumento la performance
    
    'SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = "SELECT Top 1 * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND TipoMovim = '" & "FCOM" & "' "
    SQLX$ = SQLX$ + "AND NroCai <> '' "
'    SQLX$ = SQLX$ + "AND ISNULL(NroCai) = False "
'    SQLX$ = SQLX$ + "AND FeVtoCai >= Now "
    SQLX$ = SQLX$ + "AND NroCai is not null "
    SQLX$ = SQLX$ + "AND FeVtoCai >= '" & FechaCentral & "' "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi desc"
    'SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC"
    'Set FacCompras = CueCorri.OpenRecordset(SQLX$)
    '
    
    Dim FacCompras As ADODB.Recordset
    Set FacCompras = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    With FacCompras
      'On Error Resume Next
      '.MoveLast
      'If Err < 1 Then
      If Not (.EOF And .BOF) Then
        NUCAI$ = !NROCAI
        FECHACAI = CVINT(!FEVTOCAI)
        If FECHACAI > HOY(HOS$) Then
          Text16 = NUCAI$
          Text17 = FECHATEX$(FECHACAI)
        End If
      End If
      'On Error GoTo 0
    End With
    '
End Sub

Private Function VALIDACAI() As Integer
    '
    VALIDACAI = 1
    NC% = XVALO(Text1)
    PVCL% = PIVACLI%(-1 * NC%)
    If PVCL% = 0 Or PVCL% = 5 Or PVCL% = 6 Then
      Exit Function
    End If
    '
    Text16 = TRIM$(Text16)
    If Text16 = "" Then ' Valida la existencia del CAI
      Call MENSERR(24, "Falta Ingresar Número CAI")
      Text16.SetFocus: GoTo 99
    Else
      If Len(Text16) > 24 Then ' Valida la longitud del CAI
        Call MENSERR(24, "Número de CAI Supera Máximo de 14 Dígitos")
        Text16.SetFocus: GoTo 99
      End If
    End If
    '
    Text17 = TRIM$(Text17)
    If Text17 = "" Then ' Valida la existencia de la fecha del CAI
        Call MENSERR(24, "Falta Fecha de Vencimiento del CAI.")
        Text17.SetFocus: GoTo 99
      Else
        FECHACAI = FECHANUM(Text17)
        If FECHACAI < FECHANUM(Text6) Then
          Call MENSERR(24, "Vencimiento del CAI Anterior\a la Fecha de Emisión de Factura")
          Text17.SetFocus: GoTo 99
        End If
    End If
    '
    Exit Function
    '
99  VALIDACAI = 0
    '
End Function

Private Sub VScroll1_Change()
   Frame15.Top = (-1) * VScroll1.Value * 100
End Sub

Private Sub VScroll1_Scroll()
   Frame15.Top = (-1) * VScroll1.Value * 100
End Sub

Private Sub Check2_LOSTFOCUS()
  For KKI% = 1 To 7
    If Option1(KKI%) Then
      Check2 = 0
    End If
  Next KKI%
End Sub
'
Private Sub Check2_CLICK()
  '
  Call Check2_LOSTFOCUS
  '
  If Check2.Value > 0 Then
     MONEMIA% = MONEMI%
15   MONEMI% = 1
     Label29 = DMONE$(MONEMI%)
     Label29.ForeColor = RGB(0, 0, 255)
     '
     Call RECALTO(TCMONE#(MONEMI%), TCMONE#(MONEMIA%))
     Text18 = FORMATNUM(TCMONE#(MONEMI%), "F11.4")
  End If
  '
End Sub

