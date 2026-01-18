VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CFGPED07 
   BackColor       =   &H00E0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Pedidos de Clientes "
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6900
      Style           =   2  'Dropdown List
      TabIndex        =   140
      Top             =   6525
      Width           =   3465
   End
   Begin VB.ComboBox Combo5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1660
      Style           =   2  'Dropdown List
      TabIndex        =   138
      Top             =   6530
      Width           =   3465
   End
   Begin VB.Frame Frame29 
      BackColor       =   &H00E0E0C0&
      Height          =   6375
      Left            =   5400
      TabIndex        =   77
      Top             =   100
      Width           =   4935
      Begin VB.Frame Frame34 
         BackColor       =   &H00E0E0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2520
         TabIndex        =   129
         Top             =   4230
         Width           =   2190
         Begin VB.CheckBox Check39 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Items S/Cargo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   132
            ToolTipText     =   "Poder Marcar Items S/Cargo"
            Top             =   200
            Width           =   1500
         End
         Begin VB.TextBox Text12 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   1390
            MaxLength       =   8
            TabIndex        =   130
            ToolTipText     =   "Ingrese el Porcentaje de Tolerancia (Funciona Combinado Con la Marca de Items S/Cargo"
            Top             =   480
            Width           =   740
         End
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0C0&
            Caption         =   "% de ToleranciaVs Lista de Precio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   120
            TabIndex        =   131
            Top             =   440
            Width           =   1500
         End
      End
      Begin VB.Frame Frame33 
         BackColor       =   &H00E0E0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   150
         TabIndex        =   122
         Top             =   5520
         Width           =   4580
         Begin VB.CheckBox Check37 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "No Valida Moneda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2655
            TabIndex        =   127
            ToolTipText     =   "Si se Marca No Avisa si se está Grabando el Pedido en una Moneda Diferente a la Lista Activa."
            Top             =   390
            Width           =   1870
         End
         Begin VB.CheckBox Check35 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Observ. x Item"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2655
            TabIndex        =   125
            ToolTipText     =   "D"
            Top             =   165
            Width           =   1700
         End
         Begin VB.CheckBox Check34 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Cargo de Flete Mandatariro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   30
            TabIndex        =   124
            ToolTipText     =   "D"
            Top             =   390
            Width           =   2745
         End
         Begin VB.CheckBox Check33 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Dom. de Entrega Mandatario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   30
            TabIndex        =   123
            ToolTipText     =   "D"
            Top             =   160
            Width           =   2625
         End
      End
      Begin VB.Frame Frame31 
         BackColor       =   &H00E0E0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   150
         TabIndex        =   117
         Top             =   5050
         Width           =   4580
         Begin VB.CheckBox Check40 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Asig. Lote a Pedidos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2640
            TabIndex        =   135
            ToolTipText     =   "Asignción de Lotes al Pedido Previo a la Facturación"
            Top             =   150
            Width           =   1800
         End
         Begin VB.CheckBox Check31 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "No Mostrar Datos Entr. Cliente "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   118
            Top             =   150
            Width           =   2505
         End
      End
      Begin VB.Frame Frame30 
         BackColor       =   &H00E0E0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   150
         TabIndex        =   109
         Top             =   4230
         Width           =   2295
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   1680
            MaxLength       =   8
            TabIndex        =   110
            ToolTipText     =   "Introduzca la cantidad de dias de tolerancia  para depurar los pedidos"
            Top             =   210
            Width           =   525
         End
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Aprobación Posterior si hace más de:    días que no Compra"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   120
            TabIndex        =   111
            Top             =   200
            Width           =   1500
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0C0&
         Caption         =   "CREDITO C/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1040
         Left            =   1605
         TabIndex        =   106
         Top             =   1140
         Width           =   1395
         Begin VB.CheckBox Check16 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "V.Lim.C.C"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   130
            TabIndex        =   108
            Top             =   650
            Width           =   1150
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "V.Límite General "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   130
            TabIndex        =   107
            Top             =   220
            Width           =   1150
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Ventana"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   3090
         TabIndex        =   104
         Top             =   360
         Width           =   1650
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   120
            MaxLength       =   8
            TabIndex        =   105
            ToolTipText     =   "Introduzca el Nombre del Controlador 'PRF'"
            Top             =   315
            Width           =   1460
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00E0E0C0&
         Caption         =   "VALORIZAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   120
         TabIndex        =   101
         Top             =   360
         Width           =   1635
         Begin VB.CheckBox Check12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Ver.Stock"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   120
            TabIndex        =   103
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Mues.Precios"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00E0E0C0&
         Caption         =   "FABRICACION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1040
         Left            =   3090
         TabIndex        =   97
         Top             =   1140
         Width           =   1650
         Begin VB.CheckBox Check38 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Anul. Ped. c/OF"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   140
            TabIndex        =   128
            Top             =   780
            Width           =   1470
         End
         Begin VB.CheckBox Check15 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "O.F.Directa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   140
            TabIndex        =   100
            Top             =   580
            Width           =   1470
         End
         Begin VB.CheckBox Check13 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "No Val. St."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   140
            TabIndex        =   99
            Top             =   390
            Width           =   1410
         End
         Begin VB.CheckBox Check9 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Gen. O/Fab."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   140
            TabIndex        =   98
            Top             =   190
            Width           =   1425
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00E0E0C0&
         Caption         =   "O/C CLIENTE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1040
         Left            =   150
         TabIndex        =   94
         Top             =   1140
         Width           =   1395
         Begin VB.CheckBox Check17 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "No Repite"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   350
            Left            =   160
            TabIndex        =   96
            Top             =   650
            Width           =   1190
         End
         Begin VB.CheckBox Check10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Obliga Informar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   160
            TabIndex        =   95
            Top             =   220
            Width           =   1190
         End
      End
      Begin VB.Frame Frame20 
         BackColor       =   &H00E0E0C0&
         Caption         =   "ENTREGA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   1815
         TabIndex        =   92
         Top             =   360
         Width           =   1205
         Begin VB.CheckBox Check23 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Actualiza"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   100
            TabIndex        =   93
            Top             =   360
            Width           =   1050
         End
      End
      Begin VB.Frame Frame26 
         BackColor       =   &H00E0E0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   150
         TabIndex        =   86
         Top             =   2160
         Width           =   2840
         Begin VB.TextBox Text13 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2100
            MaxLength       =   8
            TabIndex        =   136
            ToolTipText     =   "Cantidad de dias de Vencimiento de Presupuestos a Presentar"
            Top             =   720
            Width           =   640
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   840
            MaxLength       =   8
            TabIndex        =   89
            ToolTipText     =   "Ingrese Descuento para el  Cliente Retira de Fábrica"
            Top             =   765
            Width           =   525
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2100
            MaxLength       =   8
            TabIndex        =   88
            ToolTipText     =   "Introduzca la cantidad de dias de tolerancia  para depurar los pedidos"
            Top             =   210
            Width           =   640
         End
         Begin VB.CheckBox Check24 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Validar Dias Ped.Pend."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   370
            Left            =   105
            TabIndex        =   87
            ToolTipText     =   "Validar Dias Ped. Pendientes"
            Top             =   180
            Width           =   1350
         End
         Begin VB.Label Label15 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Venc. Presup."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1500
            TabIndex        =   137
            Top             =   650
            Width           =   780
         End
         Begin VB.Label Label10 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Desc.Ret. Clte %"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            TabIndex        =   91
            Top             =   690
            Width           =   795
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Dias"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            TabIndex        =   90
            Top             =   285
            Width           =   540
         End
      End
      Begin VB.Frame Frame27 
         BackColor       =   &H00E0E0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   3090
         TabIndex        =   82
         Top             =   2160
         Width           =   1650
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   765
            Width           =   1425
         End
         Begin VB.CheckBox Check25 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Valida Codigo Solicitado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   465
            Left            =   120
            TabIndex        =   83
            ToolTipText     =   "Validación de Codigos Pedidos Anteriores Para el Cliente Activo"
            Top             =   130
            Width           =   1500
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Mon.x Defecto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   120
            TabIndex        =   85
            Top             =   540
            Width           =   1245
         End
      End
      Begin VB.Frame Frame28 
         BackColor       =   &H00E0E0C0&
         Caption         =   "CONTROL DE CAMBIOS COND. COMERCIALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   150
         TabIndex        =   78
         ToolTipText     =   $"CFGPED07.frx":0000
         Top             =   3300
         Width           =   4580
         Begin VB.CheckBox Check30 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Exige C.Pago"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3000
            TabIndex        =   116
            ToolTipText     =   "Lista de Precios"
            Top             =   550
            Width           =   1455
         End
         Begin VB.CheckBox Check29 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "No Muestra C.Pago - Variable"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   120
            TabIndex        =   115
            ToolTipText     =   "Lista de Precios"
            Top             =   550
            Width           =   2775
         End
         Begin VB.CheckBox Check28 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Fact.Vencidas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3000
            TabIndex        =   81
            Top             =   280
            Width           =   1455
         End
         Begin VB.CheckBox Check27 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Cond.Pago"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1440
            TabIndex        =   80
            Top             =   280
            Width           =   1215
         End
         Begin VB.CheckBox Check26 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "L.Precios"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   79
            ToolTipText     =   "Lista de Precios"
            Top             =   280
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2730
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   75
      Top             =   630
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      TabIndex        =   48
      Text            =   " "
      Top             =   1500
      Width           =   1500
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
      Height          =   285
      Left            =   3990
      TabIndex        =   45
      Top             =   1500
      Width           =   175
   End
   Begin VB.CommandButton Command12 
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
      Height          =   650
      Left            =   4200
      Picture         =   "CFGPED07.frx":009A
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1560
      Width           =   750
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Importación de Pedidos"
      Height          =   1275
      Left            =   240
      TabIndex        =   37
      Top             =   8400
      Width           =   6000
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
         Height          =   285
         Left            =   5040
         TabIndex        =   40
         Top             =   840
         Width           =   175
      End
      Begin VB.CommandButton Command10 
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
         Left            =   5670
         TabIndex        =   38
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4095
         TabIndex        =   42
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Corte - Inicio de Importación:"
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
         TabIndex        =   41
         Top             =   920
         Width           =   3900
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         TabIndex        =   39
         Top             =   420
         Width           =   5580
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8280
      Left            =   10440
      ScaleHeight     =   8220
      ScaleWidth      =   1110
      TabIndex        =   28
      Top             =   0
      Width           =   1170
      Begin VB.CommandButton Command15 
         Caption         =   "Desc."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "CFGPED07.frx":03A4
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Tabla de Descuentos por Monto"
         Top             =   4900
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         Caption         =   "T-Ing."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "CFGPED07.frx":04EE
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Tabla de Tipo de Ingreso"
         Top             =   4245
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "I.V.A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "CFGPED07.frx":0638
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Tabla de Tasas de IVA"
         Top             =   3585
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Cotiz."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "CFGPED07.frx":0782
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Parametros de Cálculo - Tipo de Cambio Dolar"
         Top             =   2325
         Width           =   650
      End
      Begin VB.CommandButton Command9 
         Caption         =   "S-Cod"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "CFGPED07.frx":08CC
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Tabla de Sub-Codigos"
         Top             =   2940
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "CFGPED07.frx":0A16
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Guarda la Configuracion"
         Top             =   1600
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "CFGPED07.frx":0D20
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Ayuda en Linea"
         Top             =   950
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "CFGPED07.frx":102A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cerrar y Salir"
         Top             =   300
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0C0&
      Caption         =   "PEDIDO DE CLIENTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5060
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1755
         Style           =   2  'Dropdown List
         TabIndex        =   134
         Top             =   2160
         Width           =   3100
      End
      Begin VB.Frame Frame32 
         BackColor       =   &H00E0E0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   210
         TabIndex        =   119
         Top             =   5670
         Width           =   4650
         Begin VB.CheckBox Check32 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Fabricación por Moldes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   121
            ToolTipText     =   "Utilizar Configuración de proyección de Stocks por Moldes"
            Top             =   250
            Width           =   2265
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Configuraciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   2400
            TabIndex        =   120
            Top             =   180
            Width           =   2040
         End
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1755
         Style           =   2  'Dropdown List
         TabIndex        =   113
         Top             =   940
         Width           =   3100
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Listas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   4200
         Picture         =   "CFGPED07.frx":1174
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   5040
         Width           =   675
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00E0E0C0&
         Caption         =   "LISTADOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2880
         TabIndex        =   51
         Top             =   2580
         Width           =   1170
         Begin VB.OptionButton Option8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Mon.Em"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   100
            TabIndex        =   53
            Top             =   400
            Width           =   960
         End
         Begin VB.OptionButton Option7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   100
            TabIndex        =   52
            Top             =   210
            Value           =   -1  'True
            Width           =   960
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00E0E0C0&
         Caption         =   "COM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4050
         TabIndex        =   56
         ToolTipText     =   "Comisión"
         Top             =   2580
         Width           =   840
         Begin VB.OptionButton Option10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Let."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   105
            TabIndex        =   58
            Top             =   400
            Width           =   645
         End
         Begin VB.OptionButton Option9 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   105
            TabIndex        =   57
            Top             =   210
            Value           =   -1  'True
            Width           =   540
         End
      End
      Begin VB.Frame Frame25 
         BackColor       =   &H00E0E0C0&
         Caption         =   "DESCUENTOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   2860
         TabIndex        =   71
         Top             =   3225
         Width           =   2020
         Begin VB.OptionButton Option12 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Dto. x Monto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   110
            TabIndex        =   73
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton Option11 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Dto.Standard"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   110
            TabIndex        =   72
            Top             =   300
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0C0&
         Caption         =   "NUMERACION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   210
         TabIndex        =   15
         Top             =   4140
         Width           =   1905
         Begin VB.OptionButton Option4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   945
            TabIndex        =   17
            Top             =   250
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Man"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   100
            TabIndex        =   16
            Top             =   250
            Width           =   765
         End
      End
      Begin VB.Frame Frame23 
         BackColor       =   &H00E0E0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   210
         TabIndex        =   66
         Top             =   4665
         Width           =   1905
         Begin VB.CheckBox Check21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "No Val.Rep.Cod"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   105
            TabIndex        =   67
            ToolTipText     =   "No Validar Repetición de Codigos al Generar Pedido"
            Top             =   180
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0C0&
         Caption         =   "APROB."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   2180
         TabIndex        =   34
         Top             =   4140
         Width           =   960
         Begin VB.OptionButton Option6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Direc"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   80
            TabIndex        =   36
            Top             =   280
            Width           =   780
         End
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Post."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   80
            TabIndex        =   35
            Top             =   690
            Value           =   -1  'True
            Width           =   800
         End
      End
      Begin VB.Frame Frame24 
         BackColor       =   &H00E0E0C0&
         Height          =   465
         Left            =   1800
         TabIndex        =   69
         Top             =   5205
         Width           =   1320
         Begin VB.CheckBox Check22 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Val. T-Ing"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   100
            TabIndex        =   70
            Top             =   180
            Width           =   1180
         End
      End
      Begin VB.Frame Frame22 
         BackColor       =   &H00E0E0C0&
         Caption         =   "DESC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   3200
         TabIndex        =   62
         Top             =   4140
         Width           =   920
         Begin VB.CheckBox Check36 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Obs."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   126
            ToolTipText     =   "Observaciones del Item"
            Top             =   1185
            Width           =   795
         End
         Begin VB.CheckBox Check20 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Comp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   65
            ToolTipText     =   "Descripcion Complementaria"
            Top             =   585
            Width           =   795
         End
         Begin VB.CheckBox Check19 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Princ."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   64
            ToolTipText     =   "Descripción Principal del Item"
            Top             =   285
            Width           =   795
         End
         Begin VB.CheckBox Check18 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Espec"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   63
            ToolTipText     =   "Especificaciones del Item (Grabadas en el Maestro de Items de Stocks)"
            Top             =   885
            Width           =   795
         End
      End
      Begin VB.Frame Frame21 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   210
         TabIndex        =   60
         Top             =   5205
         Width           =   1545
         Begin VB.CheckBox Check14 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Val.L.Indivis"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   61
            Top             =   180
            Width           =   1380
         End
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   3960
         Width           =   3100
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1765
         Width           =   3750
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0C0&
         Caption         =   "IMPRESION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   210
         TabIndex        =   21
         Top             =   2595
         Width           =   2640
         Begin VB.CheckBox Check11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Precios Incluyen IVA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   195
            TabIndex        =   54
            Top             =   1200
            Width           =   2115
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Impresión Fantasma"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   200
            TabIndex        =   22
            Top             =   825
            Width           =   2115
         End
         Begin VB.CheckBox Check8 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Orden de Despacho"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   200
            TabIndex        =   49
            Top             =   525
            Width           =   2180
         End
         Begin VB.CheckBox Check7 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0C0&
            Caption         =   "Orden de Venta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   200
            TabIndex        =   50
            Top             =   275
            Value           =   1  'Checked
            Width           =   1875
         End
         Begin VB.Line Line3 
            X1              =   30
            X2              =   2590
            Y1              =   1130
            Y2              =   1130
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Cód."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   4200
         TabIndex        =   18
         Top             =   4140
         Width           =   670
         Begin VB.TextBox Text8 
            Height          =   360
            Left            =   105
            TabIndex        =   19
            Text            =   "PC"
            Top             =   210
            Width           =   450
         End
      End
      Begin VB.CommandButton Command6 
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
         Height          =   650
         Left            =   4100
         Picture         =   "CFGPED07.frx":12BE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   200
         Width           =   750
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   " "
         Top             =   585
         Width           =   3750
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
         Height          =   285
         Left            =   3780
         TabIndex        =   12
         Top             =   215
         Width           =   175
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Text            =   " "
         Top             =   220
         Width           =   1380
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto Impresion "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   133
         Top             =   2280
         Width           =   1365
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto Impresion "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   345
         TabIndex        =   114
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   240
         X2              =   2760
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Despacho"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   530
         TabIndex        =   47
         Top             =   1440
         Width           =   1725
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido de Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   20
         Top             =   280
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Orden de Despacho"
      Height          =   3690
      Left            =   210
      TabIndex        =   0
      Top             =   9030
      Width           =   5160
      Begin VB.Frame Frame14 
         BackColor       =   &H00FFFFC0&
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
         Height          =   960
         Left            =   2415
         TabIndex        =   25
         Top             =   1365
         Width           =   1545
         Begin VB.CheckBox Check6 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Captura"
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
            Left            =   210
            TabIndex        =   27
            ToolTipText     =   "Captura Cola de Impresión de Ordenes de Despacho y/o Pedidos."
            Top             =   600
            Width           =   1170
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Suprime"
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
            Left            =   210
            TabIndex        =   26
            ToolTipText     =   "Suprime Impresión - Enía Datos a Cola de Impresión."
            Top             =   210
            Width           =   1170
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Generación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   2730
         TabIndex        =   23
         Top             =   2415
         Width           =   2175
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Despachar Según Disponibilidad de Stocks"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   210
            TabIndex        =   24
            Top             =   315
            Width           =   1905
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Controles Varios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   210
         TabIndex        =   7
         Top             =   2415
         Width           =   2325
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   1680
            TabIndex        =   8
            Text            =   " "
            Top             =   550
            Width           =   420
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Anticipo de Activación de Despacho (dias):"
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
            Left            =   -105
            TabIndex        =   9
            Top             =   315
            Width           =   1665
         End
      End
      Begin VB.CommandButton Command5 
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
         Height          =   880
         Left            =   4095
         Picture         =   "CFGPED07.frx":15C8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   315
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         TabIndex        =   3
         Top             =   1365
         Width           =   2070
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Manual"
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
            Left            =   420
            TabIndex        =   5
            Top             =   315
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Automatica"
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
            Left            =   420
            TabIndex        =   4
            Top             =   580
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   4095
         TabIndex        =   1
         Top             =   1365
         Width           =   855
         Begin VB.TextBox Text5 
            Height          =   360
            Left            =   210
            TabIndex        =   2
            Text            =   "OD"
            Top             =   420
            Width           =   450
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "CFGPED07.frx":18D2
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   1620
      OleObjectBlob   =   "CFGPED07.frx":1B06
      TabIndex        =   76
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Puerto Impresion Etiqutas de Despacho"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5040
      TabIndex        =   141
      Top             =   6480
      Width           =   1725
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Puerto Impresion Etiqutas de pedido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   -240
      TabIndex        =   139
      Top             =   6480
      Width           =   1725
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Formulario:"
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
      TabIndex        =   46
      Top             =   1260
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuTablaAux 
         Caption         =   "Tablas Auxiliares"
      End
      Begin VB.Menu mnuCombiequiacc 
         Caption         =   "Configuración de Combinación de Equipos con Accesorios"
      End
   End
End
Attribute VB_Name = "CFGPED07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Sub CREARTABLA_DERELISTAPRE()

   Call CREACAMPO("", "DERELISTAPRE", "Cod_Lista", 3, 0, 1)
   Call CREACAMPO("", "DERELISTAPRE", "Usuario", 3, 0, 1)
   Call CREACAMPO("", "DERELISTAPRE", "Derecho", 3, 0)
   Call CREACAMPO("", "DERELISTAPRE", "MARBORRA", 3, 0)
   
End Sub

Private Sub Check18_Click()
'   Call TRALOCAL("MOLDES", "")
'   VTB = VENTABU%("MOLDES")
'   Call TRACENTRAL("MOLDES", "")
'' If VTB% >= 0 Then
''    For U& = 1 To ULTREG&("!MOLDES")
''       X$ = REGLEIDO$("!MOLDES", U&)
''       Call REPLA(X$, MKS(U&), 1, 2)
''       Call GRAREG("MOLDES", X$, U&)
''    Next U&
''    FIN& = ULTREG&("!MOLDES")
''    Call SETULTREG("MOLDES", FIN&)
'' End If
'   Call TRALOCAL("MOLCANTI", "")
'
'20 VTB% = VENTABU%("MOLCANTI")
'   Dim LEIDO%()
'   'Verifica si el MOLDE esta repetido
'   FIN& = ULTREG("!MOLCANTI")
'   For I& = 1 To (FIN&)
'      X$ = REGLEIDO$("!MOLCANTI", I&)
'      MOL% = CVI(Mid$(X$, 1, 2))
'      If MOL% > MAXX% Then
'        ReDim Preserve LEIDO%(MOL%)
'        MAXX = MOL%
'      End If
'      If LEIDO%(MOL%) > 0 Then
'        Call COMUNI("ID " & MOL% & " Repetido")
'        Erase LEIDO(): MAXX% = 0
'
'        GoTo 20
'      Else
'        LEIDO%(MOL%) = 1
'      End If
'
'    Next I&
'
'   Call TRACENTRAL("MOLCANTI", "")
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Bases de Datos MDB(*.MDB)|*.MDB|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   Label3 = UCase$(CommonDialog1.FileName)
   '
99 Screen.MousePointer = 1
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   '
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label6 = VALACT1$("SELFECHA")
   End If
   '
End Sub

Private Sub Command12_Click()
'\\\OT-08-1003-03/11/08///
    ACONEC$ = "FLFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
'\\\OT-08-1003-03/11/08///
End Sub


Private Sub Command14_Click()
    Command14.Enabled = False
    '
    If EXISTE%(LUDAT$ + "TATINGRE.RFS") < 1 Then
      Call COMUNI("Tabla de Motivos de Ingreso de Pedidos \No Disponible en la Presente Versión")
      GoTo 99
    End If
    
    Call BLANARCH("!TATINGRE")
    '
    Call TRALOCAL("TATINGRE", "")

30  VTB% = VENTABU%("TATINGRE")
    If VTB% < 0 Then GoTo 99
    '
    For i& = 1 To ULTREG&("!TATINGRE")
       X$ = REGLEIDO$("!TATINGRE", i&)
       If Asc(Left$(X$, 1)) < 1 Then
          Call MENSERR(24, "Falta Código Para fila: " & i&)
          GoTo 30
       End If
    Next i&
    '
    Call TRACENTRAL("TATINGRE", "")
    Call CIERRARCH("TATINGRE")
    '
99  Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    Command15.Enabled = False
    '
    If EXISTE%(LUDAT$ + "DESMONTO.RFS") < 1 Then
      Call COMUNI("Tabla de Descuentos por Monto \No Disponible en la Presente Versión")
      GoTo 99
    End If
    
    Call BLANARCH("!DESMONTO")
    '
    Call TRALOCAL("DESMONTO", "")

30  VTB% = VENTABU%("DESMONTO")
    If VTB% < 0 Then GoTo 99
    '
    For i& = 1 To ULTREG&("!DESMONTO")
       X$ = REGLEIDO$("!DESMONTO", i&)
       Desde = CVS(Mid$(X$, 1, 4))
       Hasta = CVS(Mid$(X$, 5, 4))
       DESC1 = CVS(Mid$(X$, 9, 4))
       DESC2 = CVS(Mid$(X$, 13, 4))
       If Desde < 0.05 And i& > 1 Then
          Call COMUNI("Falta Valor Desde en Fila: " & i&)
          GoTo 30
       End If
       '
       If Hasta < 1 Then
          Call COMUNI("Falta Valor Hasta en Fila: " & i&)
          GoTo 30
       End If
       '
       If DESE >= Hasta Then
          Call COMUNI("Valor Campo 'Hasta' Debe ser Mayor a Campo 'Desde' en Fila: " & i&)
          GoTo 30
       End If
       '
       If DESC1 < 0.05 And DESC2 < 0.05 Then
          Call COMUNI("Debe Completar al Menos un Descuento en Fila: " & i&)
          GoTo 30
       End If
    Next i&
    '
    Call TRACENTRAL("DESMONTO", "")
    Call CIERRARCH("DESMONTO")
    '
99  Command15.Enabled = True

End Sub



Private Sub Command16_Click()
'
    Command16.Enabled = False
    '
''''''Private Sub Command3_Click()
''''''
'''''' Call TRALOCAL("MOLDES")
'''''' VTB% = VENTABU("MOLDES")
'''''' If VTB% >= 0 Then
''''''    For U& = 1 To ULTREG&("!MOLDES")
''''''       x$ = REGLEIDO$("!MOLDES", U&)
''''''       Call REPLA(x$, CVI(U& - 1), 1, 2)
''''''       Call GRAREG("MOLDES", x$, U&)
''''''    Next U&
''''''    FIN& = ULTREG&("!MOLDES")
''''''    Call SETULTREG("MOLDES", FIN&)
'''''' End If
''''''End Sub

10  Call SELECHO("FLEXUSR")
    Screen.MousePointer = 1
    If VALACT1$("FLEXUSR") = "" Then Exit Sub
    '
    USEL% = XVALO(TRIM$(VALACT1$("FLEXUSR")))
    DEUSER$ = VALACT2$("FLEXUSR")
    '
    Call ABRECONEXION
    Call CREARTABLA_DERELISTAPRE
    '
    Call ACLISDERECHO_LIST(USEL%, DEUSER$)
    DERELPRE.Show 1
    GoTo 10


Command16.Enabled = True

'Sub CARGADER(CAPLI$, DR1$, DR2$)
'    '
'    If Left$(CAPLI$, 1) = "*" Then CAPLI$ = Mid$(CAPLI$, 2)
'    If Left$(CAPLI$, 1) = "?" Then CAPLI$ = Mid$(CAPLI$, 2)
'    '
'    VUELTA% = 0
'10  For JJI& = 1 To FORDERAC.List2.ListCount
'      Z$ = FORDERAC.List2.List(JJI& - 1)
'      If TRIM$(Mid$(Z$, 53, 32)) = CAPLI$ Then
'         Call REPLA(Z$, DR1$, 45, 2)
'         Call REPLA(Z$, DR2$, 49, 2)
'         FORDERAC.List2.List(JJI& - 1) = Z$
'         Exit Sub
'      End If
'    Next JJI&
'    '
'    ' NO ENCONTRO EL DERECHO
'    If VUELTA% = 0 Then
'       VUELTA% = 1
'       Call AGREGALIST(CAPLI$, "Aplicacion " + CAPLI$)
'       GoTo 10
'    End If
'    '
'End Sub

End Sub
Sub ACLISDERECHO_LIST(USEL%, DEUSER$)
    '
    Screen.MousePointer = 11
    DERELPRE.List1.Clear
'    DERELPRE.List2.Clear
    DERELPRE.Label2.Caption = TRIM$(Str$(USEL%))
    DERELPRE.Label4.Caption = TRIM$(DEUSER$)
    '
    ' CARGAR LISTA DE TODAS LAS APLICACIONES DEL MDB
    '
    'recorro las listas y verifico si tiene derechos
    
    User% = USNBR% Mod 100
    DERELPRE.List1.Clear
    SQLX$ = "SELECT DISTINCT (Cod_Lista) FROM LISTAPRE WITH(NOLOCK) ORDER BY COD_LISTA"
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
         CODILISTA% = XVALO(!Cod_Lista)
         CONDI$ = "Cod_Lista = " & CODILISTA%
         CONDI$ = CONDI$ + " AND USUARIO = " & USEL%
         DERECHO% = XVALO(VALOBADA("DERELISTAPRE", "DERECHO", CONDI$, "NOMESS"))
         Call AGREGADET_DERECHOS_LIST(CODILISTA%, DERECHO%)
      .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    '
    Screen.MousePointer = 1
    '
End Sub

Sub AGREGADET_DERECHOS_LIST(CODLISTA%, DERECHO%)
    '
    Z$ = Space$(128)
    'DENOLISTA$
    DAPLI$ = Left$(DENOLISTA$(CODLISTA%), 32)
    Call REPLA(Z$, DAPLI$, 1, 36)
    Call REPLA(Z$, "      ", 45, 6)
    CAPLI$ = "Lista N°: " & CODLISTA%
    If DERECHO% > 0 Then Call REPLA(Z$, "SI", 45, 2)
    If DERECHO% > 1 Then Call REPLA(Z$, "SI", 49, 2)
    Call REPLA(Z$, AJUSTI$(CAPLI$, 32), 53, 32)
    '
    DERELPRE.List1.AddItem Z$
End Sub

Private Sub Command17_Click()
    CONFABPED.Show 1
End Sub

Private Sub Command2_Click()
    '
    If XVALO(Text12.TEXT) > 99.99 Then
       Call COMUNI("Porcenta Ingresado de Tolerancia No Válido, Debe Ser Menor a 100")
       Call PINTATEXT(Text12)
       Exit Sub
    End If
    '
    Call GRACONTROL("", "FORDESPA", Text1.TEXT)
    Call GRACONTROL("", "CODDESPA", Text5.TEXT)
    Call GRACONTROL("", "FORPECLI", Text6.TEXT)
    Call GRACONTROL("ADMIPED", "TOLELPRE", Text12.TEXT) 'porcentaje de tolerancia funciona con la combinacion de marca de items sin cargo
    Call GRACONTROL("", "CODPECLI", Text8.TEXT)
    Call GRACONTROL("", "VENCARPE", Text4.TEXT)
    Call GRACONTROL("PEDIDO", "ORIDATOS", Label3)
    Call GRACONTROL("PEDIDO", "FECORTE", Label6)
    valor$ = "NO"
    If Check12.Value = 1 Then
       valor$ = "SI"
    End If
    Call GRACONTROL("PEDIDO", "VERISTO", valor$)
    'NO PRESENTAR DATOS DE ENTREGA - SOLICITUD DE MEDITEA
    valor$ = "NO"
    If Check31.Value = 1 Then
       valor$ = "SI"
    End If
    Call GRACONTROL("NOPREENT", "CONFPEDI", valor$)
    '
    'Asignacion de lotes al pedido previo a la facturacion- SOLICITUD DE MEDITEA
    valor$ = "NO"
    If Check40.Value = 1 Then
       valor$ = "SI"
    End If
    Call GRACONTROL("PEDIDO", "ASILOTPE", valor$)

    valor$ = "NO"
    If Check33.Value = 1 Then
       valor$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "DOMIMAND", valor$)
    valor$ = "NO"
    If Check34.Value = 1 Then
       valor$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "FLETMAND", valor$)
    '
    valor$ = "NO" ' SE VISUALIZA CAJA DE TEXTO PARA AGREGAR TEXTO A CADA ITEMS
    If Check35.Value = 1 Then
       valor$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "OBSXITEM", valor$)
    '
    valor$ = "NO" ' SE VISUALIZA CAJA DE TEXTO PARA AGREGAR TEXTO A CADA ITEMS
    If Check36.Value = 1 Then
       valor$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "PRINOBSE", valor$)
    '
    MONUME$ = "AUTO"
    If Option1.Value = True Then
        Call COMUNI("No Está Disponible la Opción\de Numeración Manual de O/D's.")
        Exit Sub
    End If
    Call GRACONTROL("", "NUMDESPA", MONUME$)
    '
    MONUME$ = TRIM$(Text3.TEXT)
    Call GRACONTROL$("", "ANTIDESP", MONUME$)
    '
    MONUME$ = "AUTO"
    If Option3.Value = True Then
        Call COMUNI("No Está Disponible la Opción\de Numeración Manual de Pedidos de Clientes.")
        Exit Sub
    End If
    Call GRACONTROL("", "NUMPECLI", MONUME$)
    '
    MONUME$ = "NORMAL": FORCAPED% = 1
    Call GRACONTROL("ADMIPED", "FORMACAR", MONUME$)
    '
    MONUME$ = "NO": VERLICRE% = 0
    If Check1.Value = 1 Then
        MONUME$ = "SI": VERLICRE% = 1
    End If
    Call GRACONTROL("ADMIPED", "VERLICRE", MONUME$)
    '
    MONUME$ = "NO":
    If Check16.Value = 1 Then
        MONUME$ = "SI":
    End If
    Call GRACONTROL("ADMIPED", "VRLICCTE", MONUME$)

    MONUME$ = "NO": GENORFAB% = 0
    If Check9.Value = 1 Then
        MONUME$ = "SI": GENORFAB% = 1
    End If
    Call GRACONTROL("ADMIPED", "GENORFAB", MONUME$)
    '
    'VALIDA SI CONTROLA SI SE ESTA GRABANDO EL PEDIDO EN UNA MONEDA DIFERENTE A LA LISTA ACTIVA.
    MONUME$ = "SI"
    If Check37.Value = 1 Then
        MONUME$ = "NO"
    End If
    Call GRACONTROL("ADMIPED", "VALIMONE", MONUME$)
    '
    'Validacion de disponibilidad de stock
    MONUME$ = "NO": NOVALSTK% = 0
    If Check13.Value = 1 Then
        MONUME$ = "SI": NOVALSTK% = 1
    End If
    Call GRACONTROL("ADMIPED", "NOVALSTK", MONUME$)
    '
    MONUME$ = "NO": OFDIRECT% = 0
    If Check15.Value = 1 Then
        MONUME$ = "SI": OFDIRECT% = 1
    End If
    Call GRACONTROL("ADMIPED", "OFDIRECT", MONUME$)

    MONUME$ = "NO"
    If Check38.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("ADMIPED", "ANUPEDOF", MONUME$)

    '\\\OT-06-0512-RG-27/11/06///
    MONUME$ = "NO"
    If Check10.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "OBLICOM", MONUME$)
    '\\\OT-06-0512-RG-FIN///
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("PEDCLIEN", "SUPRINT", SUPRINT$)
    '
    SUPRINT$ = "NO"
    If Check5.Value = 1 Then SUPRINT$ = "SI"
    Call GRACONTROL("ORDESPA", "SUPRINT", SUPRINT$)
    '
    CAPRINT$ = "NO"
    If Check6.Value = 1 Then CAPRINT$ = "SI"
    Call GRACONTROL("CAPRINT", IDENTER$, CAPRINT$)
    '
    CMO$ = ""
    If Check3.Value = 1 Then CMO$ = "REVISTOCK"
    Call GRACONTROL$("ADMIPED", "GENDESPA", CMO$)
    '
    CMO$ = "NO"
    If Check4.Value = 1 Then CMO$ = "SI"
    Call GRACONTROL$("ADMIPED", "MUEPREC", CMO$)
    '\\\OT-05-0236-WAR-02/05/05///
    '
    'CONTROL DE NO REPITE ORDEN DE COMPRA
    CMO$ = "NO"
    If Check17.Value = 1 Then CMO$ = "SI"
    Call GRACONTROL$("ADMIPED", "NOREPOC", CMO$)
    '
    MONUME$ = "POSTERIOR"
    If Option6.Value = True Then
        MONUME$ = "DIRECT"
    End If
    Call GRACONTROL("PEDCLI", "MODOAPRO", MONUME$)
    '\\\OT-05-0236-WAR-FIN///
    '\\\OT-06-0439-MH-19/10/06///
    CMO$ = "SI"
    If Check7.Value = 0 Then CMO$ = "NO"
    Call GRACONTROL("PEDIDO", "PRINOVEN", CMO$)
    '
    CMO$ = "NO"
    If Check8.Value = 1 Then CMO$ = "SI"
    Call GRACONTROL("PEDIDO", "PRINODES", CMO$)
    '\\\OT-06-0439-MH-FIN///
    '
    MONUME$ = "PESOS"
    If Option8.Value = True Then MONUME$ = "MONEMORI"
    Call GRACONTROL("ADMIPED", "FORMALIS", MONUME$)
    '
    MOCOMIS$ = "PORCIEN"
    If Option10.Value = True Then MOCOMIS$ = "LETRA"
    Call GRACONTROL("ADMIPED", "MOCOMIS", MOCOMIS$)
    '
    AA$ = ""
    If Check11.Value = 1 Then AA$ = "SI"
    Call GRACONTROL("PEDCLIEN", "PRECOIVA", AA$) 'GRABA OPCION DE IMPRSION CON IVA
    '
    Call GRACONTROL(IDENTER$, "PUERDESP", Combo2.TEXT)
    Call GRACONTROL("*", "PUERDESP", Combo2.TEXT)
    
    Call GRACONTROL(IDENTER$, "PUERORVE", Combo3.TEXT)
    Call GRACONTROL("*", "PUERORVE", Combo3.TEXT)

    Call GRACONTROL(IDENTER$, "PUERETIP", Combo5.TEXT)
    Call GRACONTROL("*", "PUERETIP", Combo5.TEXT)
    
    Call GRACONTROL(IDENTER$, "PUERETDE", Combo6.TEXT)
    Call GRACONTROL("*", "PUERETDE", Combo6.TEXT)

    'Call GRACONTROL("ORDESPA", "PUERTOIM", Combo2.TEXT) 'GRABA LA IMPRESORA SELECCIONADA
    '
    'VALIDA SI UTILIZA LOTE INDIVISIBLE
    MONUME$ = "NO"
    If Check14.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "VALOIND", MONUME$)
    '
    'CONFIGURACION DE IMPRESION DE ESPECIFICACION EN  O.DESPACHO Y PEDIDO
    MONUME$ = ""
    If Check19.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "IMPDESPR", MONUME$)
    '
    MONUME$ = ""
    If Check20.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "IMPDESCO", MONUME$)
    '
    MONUME$ = ""
    If Check18.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "IMPRESPE", MONUME$)
    '
    'No Validar la repeticion de codigo
    MONUME$ = "SI"
    If Check21.Value = 1 Then
        MONUME$ = "NO"
    End If
    Call GRACONTROL("ADMIPED", "VALIREPE", MONUME$)

    'VALIDA LA CARGA DEL TIPO DE INGRESO DEL PEDIDO
    MONUME$ = "NO"
    If Check22.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "VALTIPIN", MONUME$)
    
    MONUME$ = "STAND"
    If Option12.Value = True Then MONUME$ = "MONTO"
    Call GRACONTROL("PEDIDO", "MODODESC", MONUME$)

    If Check18.Value = 0 And Check19.Value = 0 And Check20.Value = 0 Then
        Call COMUNI("Debe Elegir Uno de los Valores Para Descripcion")
        Exit Sub
    End If
    '
    'POR DEFECTO VACIO O SI ACTUALIZA SI ESTA TILDADO NO ACTUALIZA LUGAR DE ENTREGA
    MONUME$ = "SI"
    If Check23.Value = 0 Then
        MONUME$ = "NO"
    End If
    Call GRACONTROL("PEDIDO", "ACTULEN", MONUME$)
    '
    MONUME$ = ""
    If Check26.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "VALLISPR", MONUME$)
    '
    MONUME$ = ""
    If Check27.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "VALCONPA", MONUME$)
    '
    MONUME$ = ""
    If Check28.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "FACVENCI", MONUME$)
    '
    ' CONDICION DE PAGO VARIABLE
    MONUME$ = ""
    If Check29.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "COPAVARI", MONUME$)
    '
    ' EXIGE CARGA DE LA CONDICION DE PAGO
    MONUME$ = ""
    If Check30.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "EXICOPA", MONUME$)
    '
    'VALIDACION DE DEPURACION DE DIAS POSTERIORES A LA FECHA DE ENTREGA
    MONUME$ = ""
    If Check24.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("ADMIPED", "DEPUPPEN", MONUME$)
    '
    'CONFIGURACION DE CANTIDAD DE DIAS PASADOS DE FECHA DE ENTREGA PARA DEPURAR PEDIDOS
    Call GRACONTROL("ADMIPED", "DEPEPEN", TRIM$(Text9))
    '
    'CONFIGURACION DE CANTIDAD DE DIAS DE TOLERANCIA DE PRESENTACION DE PRESUPUESTOS.
    Call GRACONTROL("ADMIPED", "VENCPRES", TRIM$(XVALO(Text13)))

    'VALIDACION DE CODIGOS SI YA FUERON PEDIDOS POR EL CLIENTE
    MONUME$ = ""
    If Check25.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("ADMIPED", "VACOPED", MONUME$)
    '
    MONUME$ = ""
    If XVALO(Text10) > 0 Then MONUME$ = TRIM$(Text10)
    Call GRACONTROL("ADMIPED", "DESRTCTE", MONUME$)
    '
    MONUME$ = "0"
    If XVALO(Combo1.ListIndex) > 0 Then MONUME$ = TRIM$(Combo1.ListIndex)
    Call GRACONTROL("ADMIPED", "MONEDEFE", MONUME$)
    '
    'VALIDAR DIAS DE ULTIMA COMPRA - QUEDA PENDIENTE DE APROBACION
    Call GRACONTROL("ADMIPED", "QULTICOM", SAFETEXT$(XVALO(Text11.TEXT)))
    '
    valor$ = "NO"
    If Check32.Value = 1 Then
       valor$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "FABMOLDE", valor$)
    '
    'PODER MARCAR ITEMS SIN  CARGO
    MONUME$ = ""
    If Check39.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("ADMIPED", "ITEMSCGO", MONUME$)
    '
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completa.\Re-Inicie Ahora la Aplicación.")
    Call BAJALDISCO
    Call FINAL("")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
   ACONEC$ = "FLFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") = 2 Then
     Call CARTAMONED
   End If
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    '
    Command9.Enabled = False
    '
    Call TRALOCAL("TASUBCOD", "")
10  VTB% = VENTABU("TASUBCOD")
    If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TASUBCOD")
        X$ = REGLEIDO$("!TASUBCOD", U&)
        ' Valida que el codigo no sea un valor nulo o menor a dos caracteres
        CD$ = Left$(X$, 8)
        If Len(CD$) < 1 Then
          Call MENSERR(24, "Sub-Código Menor a dos Caracteres")
          GoTo 10
        End If
        ' Valida que el CC posea descripcion
        DESC$ = TRIM$(Mid$(X$, 9, 40))
        If Len(DESC$) < 1 Then
          Call MENSERR(24, "Falta Descripción / Referencia")
          GoTo 10
        End If
        ' Valida que el CC no este repetido
        For U2& = 1 To U& - 1
          Y$ = REGLEIDO$("!TASUBCOD", U2&)
          If TRIM$(Mid$(X$, 1, 8)) = TRIM$(Mid$(Y$, 1, 8)) Then
            Call MENSERR(24, "Sub-Código Repetido")
            GoTo 10
          End If
        Next U2&
      Next U&
      Call TRACENTRAL("TASUBCOD", "")
    End If
    '
99  Command9.Enabled = True
    '
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
'    EPLI$ = TRIM$(UCase$(EMPRELI$))
'    If InStr(EPAC$, "ZETRONIC") < 1 And InStr(EPLI$, "ZETRONIC") < 1 Then
'      Frame15.Visible = False
'      Height = Picture2.Height + 500 ' SE CAMBIO EL VALOR POR  MODIFICACION DE IMPRESION DE IVA
'    End If
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    Text1.TEXT = Control$("", "FORDESPA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text6.TEXT = Control$("", "FORPECLI")
    Text7.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    Text12.TEXT = TRIM$(FORMATNUM$(XVALO(Control("ADMIPED", "TOLELPRE")), "F5.2"))
    Text13 = Control("ADMIPED", "VENCPRES")
    '
    VENCARPE$ = Control$("", "VENCARPE")
    If TRIM$(VENCARPE$) = "" Then VENCARPE$ = "CARDETP"
    Text4.TEXT = VENCARPE$
    '
    '\\\OT-06-0512-RG-27/11/06///
    If Control$("ADMIPED", "OBLICOM") = "SI" Then Check10.Value = 1
    '\\\OT-06-0512-RG-FIN///
    If Control$("ADMIPED", "VERLICRE") = "SI" Then Check1.Value = 1
    If Control$("ADMIPED", "GENORFAB") = "SI" Then Check9.Value = 1
    '
    If Control$("ADMIPED", "NOVALSTK") = "SI" Then Check13.Value = 1
    If Control$("ADMIPED", "OFDIRECT") = "SI" Then Check15.Value = 1 ' control de orden de fabricacion directa
    If Control$("ADMIPED", "ANUPEDOF") = "SI" Then
        Check38.Value = 1
    ElseIf Control$("ADMIPED", "ANUPEDOF") = "" Then
        Check38.Value = 1
    End If
    If Control$("ADMIPED", "VALIMONE") = "NO" Then Check37.Value = 1 ' VALIDA SI AVISA SI SE ESTA GRABANDO CON MONEDA DIFERENTE A LA LISTA ACTIVA
    '
    'CONTROL DE NO REPITE ORDEN DE COMPRA EN EL PEDIDO
    If Control$("ADMIPED", "NOREPOC") = "SI" Then Check17.Value = 1

    If Control$("PEDCLIEN", "SUPRINT") = "SI" Then
       Check2.Value = 1
       CANCELPRINT% = 1
    End If
    If Control$("ORDESPA", "SUPRINT") = "SI" Then
       Check5.Value = 1
    End If
    '
    If IDENTER$ <> "" Then
      If Control$("CAPRINT", IDENTER$) = "SI" Then
        Check6.Value = 1
      End If
    End If
    '\\\OT-05-0236-WAR-02/05/05///
     If Control("PEDCLI", "MODOAPRO") = "DIRECT" Then
       Option6.Value = True
     End If
     '\\\OT-05-0236-WAR-FIN///
    If Control$("", "NUMDESPA") = "MANUAL" Then Option1.Value = True
    If Control$("", "NUMPECLI") = "MANUAL" Then Option3.Value = True
    If Control$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then Option5.Value = True
    If Control$("ADMIPED", "FORMALIS") = "MONEMORI" Then Option8.Value = True
    If Control$("ADMIPED", "MOCOMIS") = "LETRA" Then Option10.Value = True
    
    Text3.TEXT = TRIM$(Control$("", "ANTIDESP"))
    If XVALO(Text3.TEXT) < 1 Then Text3.TEXT = "3"
    '
    CMO$ = TRIM$(Control$("", "CODDESPA")): If CMO$ <> "" Then Text5.TEXT = CMO$
    CMO$ = TRIM$(Control$("", "CODPECLI")): If CMO$ <> "" Then Text8.TEXT = CMO$
    '
    If Control$("ADMIPED", "GENDESPA") = "REVISTOCK" Then Check3.Value = 1
    If Control$("ADMIPED", "MUEPREC") = "SI" Then Check4.Value = 1
    '
    Label3 = Control$("PEDIDO", "ORIDATOS")
    Label6 = Control$("PEDIDO", "FECORTE")
    '\\\OT-06-0439-MH-19/10/06///
    Check7.Value = 1
    If Control$("PEDIDO", "PRINOVEN") = "NO" Then Check7.Value = 0
    Check8.Value = 0
    If Control$("PEDIDO", "PRINODES") = "SI" Then Check8.Value = 1
    '
    Check11.Value = 0
    If Control("PEDCLIEN", "PRECOIVA") = "SI" Then Check11.Value = 1 ' OPCION DE IMPRESION CON IVA
    
    'PDES$ = Control$("ORDESPA", "PUERTOIM")
    '
    If Control$("PEDIDO", "VERISTO") = "SI" Then
       Check12.Value = 1
       Else
       Check12.Value = 0
    End If
    '
    'CONTROL DE LOTE INDIVISIBLE
    If Control$("ADMIPED", "VALOIND") = "SI" Then Check14.Value = 1
    '
    'CONFIGURACION DE IMPRESION DE ESPECIFICACION DE ORDEN DE DESPACHO Y PEDIDO
    Check18.Value = 0
    If Control("ADMIPED", "IMPRESPE") = "SI" Then Check18.Value = 1
    '
    Check19.Value = 0
    If Control("ADMIPED", "IMPDESPR") = "SI" Then Check19.Value = 1
    '
    Check20.Value = 0
    If Control("ADMIPED", "IMPDESCO") = "SI" Then Check20.Value = 1
    '
    'CONTROL DE VALIDACION DE LIMITE DE CREDITO EN CUENTA CORRIENTE
    Check16.Value = 0
    If Control("ADMIPED", "VRLICCTE") = "SI" Then Check16.Value = 1
    '
    'Validacion de marca de no controlar repeticion de codigos repetidos en detalle de pedido
    If Control("ADMIPED", "VALIREPE") = "NO" Then Check21.Value = 1
    '
    'TIPO DE INGRESO
    If Control$("ADMIPED", "VALTIPIN") = "SI" Then Check22.Value = 1
    '
    If Control("PEDIDO", "MODODESC") = "MONTO" Then
       Option12.Value = True
    End If
    '
    Check23.Value = 1
    If Control("PEDIDO", "ACTULEN") = "NO" Then
       Check23.Value = 0
    End If
    '
    'VALIDACION DE DEPURACION DE DIAS POSTERIORES A LA FECHA DE ENTREGA
    Check24.Value = 0
    If Control("ADMIPED", "DEPUPPEN") = "SI" Then Check24.Value = 1
    '
    'VALIDACION DE CODIGOS SI YA FUERON PEDIDOS POR EL CLIENTE
    Check25.Value = 0
    If Control("ADMIPED", "VACOPED") = "SI" Then Check25.Value = 1
    '
    Text9 = XVALO(Control("ADMIPED", "DEPEPEN"))
    Text10 = FORMATNUM$(XVALO(Control("ADMIPED", "DESRTCTE")), "F5.2")
    '
    'CONTROL DE VALIDACION DE CONDICIONES COMERCIALES
    Check26.Value = 0
    If TRIM$(Control("ADMIPED", "VALLISPR")) = "SI" Then
       Check26.Value = 1
    End If
    '
    'CONTROL DE VALIDACION DE CONDICIONES COMERCIALES
    Check28.Value = 0
    If TRIM$(Control("ADMIPED", "FACVENCI")) = "SI" Then
       Check28.Value = 1
    End If
    '
    'CONTROL DE VALIDACION DE CONDICIONES COMERCIALES
    Check27.Value = 0
    If TRIM$(Control("ADMIPED", "VALCONPA")) = "SI" Then
       Check27.Value = 1
    End If
    '
    ' NO PRESENTA CONDICION DE PAGO - CONDICION VARIABLE
    Check29.Value = 0
    If TRIM$(Control("ADMIPED", "COPAVARI")) = "SI" Then
       Check29.Value = 1
    End If
    '
    ' EXIGE CONDICION DE PAGO
    Check30.Value = 0
    If TRIM$(Control("ADMIPED", "EXICOPA")) = "SI" Then
       Check30.Value = 1
    End If

''''    Combo2.Clear ' CARGA EL COMBO DE LAS IMPRESORAS
''''    Dim PRX As Printer
''''    On Error Resume Next
''''    Combo2.TEXT = Printer.DeviceName
''''    On Error GoTo 0
''''    For Each PRX In Printers
''''      Combo2.AddItem PRX.DeviceName
''''      If PRX.DeviceName = PDES$ Then
''''        Combo2.TEXT = PDES$
''''      End If
''''    Next
    PDES$ = Control$(IDENTER$, "PUERDESP")
    If PDES$ = "" Then PDES$ = Control$("*", "PUERDESP")
    Combo2.Clear
    Dim PRX As Printer
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo2.AddItem PRX.DeviceName
      If PRX.DeviceName = PDES$ Then
        Combo2.TEXT = PDES$
      End If
    Next

    POVEN$ = Control$(IDENTER$, "PUERORVE")
    If POVEN$ = "" Then POVEN$ = Control$("*", "PUERORVE")
    Combo3.Clear
    Dim PRX1 As Printer
    On Error Resume Next
    Combo3.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX1 In Printers
      Combo3.AddItem PRX1.DeviceName
      If PRX1.DeviceName = POVEN$ Then
        Combo3.TEXT = POVEN$
      End If
    Next
    
    '\\\OT-06-0439-MH-FIN///
    'carga combo con monedas para seleccionar moneda por defecto
    Combo1.Clear
    Combo1.AddItem ""
    CAMON% = ULTREG&("TAMONED")
    If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
           DMONE$(U&) = DMX$
           TCMONE(U) = VHX
           Combo1.AddItem DMX$
         Next U&
    End If
    Combo1.ListIndex = 0
    MOX% = XVALO(Control("ADMIPED", "MONEDEFE"))
    If MOX% > 0 Then Combo1.ListIndex = MOX%
    '
    Text11 = TRIM$(Control("ADMIPED", "QULTICOM"))
    
    'PRESENTAR DATOS DE ENTREGA SI O NO
    Check31.Value = 0
    If Control("NOPREENT", "CONFPEDI") = "SI" Then Check31.Value = 1
    '
    'Asignacion de lotes al pedido previo a la facturacion- SOLICITUD DE MEDITEA
    Check40.Value = 0
    If Control("PEDIDO", "ASILOTPE") = "SI" Then Check40.Value = 1
    '
    Check32.Value = 0
    valor$ = Control("ADMIPED", "FABMOLDE")
    If valor$ = "SI" Then Check32.Value = 1
    '
    Check33.Value = 0
    If Control("ADMIPED", "DOMIMAND") = "SI" Then
       Check33.Value = 1
    End If
    
    Check34.Value = 0
    If Control("ADMIPED", "FLETMAND") = "SI" Then
       Check34.Value = 1
    End If
    '
    Check35.Value = 0
    If Control("ADMIPED", "OBSXITEM") = "SI" Then
       Check35.Value = 1
    End If
    '
    Check36.Value = 0
    If Control("ADMIPED", "PRINOBSE") = "SI" Then
       Check36.Value = 1
    End If
    
    Check39.Value = 0
    If Control("ADMIPED", "ITEMSCGO") = "SI" Then
       Check39.Value = 1
    End If

    PETIQPEDIDO$ = Control$(IDENTER$, "PUERETIP")
    If PETIQPEDIDO$ = "" Then PETIQPEDIDO$ = Control$("*", "PUERETIP")
    Combo5.Clear
    Dim PRX2 As Printer
    On Error Resume Next
    Combo5.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX2 In Printers
      Combo5.AddItem PRX2.DeviceName
      If PRX2.DeviceName = PETIQPEDIDO$ Then
        Combo5.TEXT = PETIQPEDIDO$
      End If
    Next
    
    PETIQDESPA$ = Control$(IDENTER$, "PUERETDE")
    If PETIQDESPA$ = "" Then PETIQDESPA$ = Control$("*", "PUERETDE")
    Combo6.Clear
    Dim PRX3 As Printer
    On Error Resume Next
    Combo6.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX3 In Printers
      Combo6.AddItem PRX3.DeviceName
      If PRX3.DeviceName = PETIQDESPA$ Then
        Combo6.TEXT = PETIQDESPA$
      End If
    Next
   
End Sub

Private Sub mnuCombiequiacc_Click()
   If Control("FACTURA", "VALITIVA") = "SI" Then
     Call TABLA_COMBINA_EQUIPOYACCESORIOS
   Else
     Call COMUNI("Falta Marca de Configuración de: 'Combinación de Iva (Equipo y Accesorio)'\Esta Se Realiza Desde el Setup de Ventas\Una Vez Configurada Vuelva a Intentar\Realizar La Carga de Items Que Combinados\Pueden Variar su Porcentaje de IVA")
   End If
 
End Sub

Private Sub mnuTablaAux_Click()
  Call GENERAR_ESTRUCTURAS_SQL
  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.msf2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))
  If InStr(EPAC$, "ENVAPLAST") > 0 Then
    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
    If LIXTA$ = "MATERIALES ENVAPLAST" Then
       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
       Exit Sub
    End If
  End If

  Call ABM_TABLA(IDLIST&)

End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub


Private Sub Text10_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text10_LostFocus()
    Text10 = FORMATNUM(XVALO(Text10), "F5.2")
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text7_CLICK()
    Text6.SetFocus
End Sub


Private Sub Text6_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

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

Private Sub Command13_Click()
    Command13.Enabled = False
    If DERACCE%("TASAIVA", "Tasas de IVA") = 2 Then
      OPX% = COMALTER%("Opciones de Trabajo:\\Tasa de Iva General\Grupos de Venta por Tasa IVA")
      If OPX% = 1 Then
          Call CARDATCENCALC
        ElseIf OPX% = 2 Then
          Call CARTASIVA
      End If
    End If
    Command13.Enabled = True
End Sub


