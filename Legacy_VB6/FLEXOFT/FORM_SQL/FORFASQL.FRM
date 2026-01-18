VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FORFASQL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación en Cuenta Corriente"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   13455
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   1320
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   0
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   8595
      Left            =   12600
      ScaleHeight     =   8565
      ScaleWidth      =   1185
      TabIndex        =   133
      Top             =   0
      Width           =   1215
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   145
         Top             =   6840
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   146
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   120
         Picture         =   "FORFASQL.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   143
         ToolTipText     =   "Aprueba, Graba e Imprime Documento"
         Top             =   7320
         Width           =   615
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORFASQL.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "Acceso al Control General de Gestión Comercial"
         Top             =   3500
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORFASQL.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   141
         ToolTipText     =   "Opciones Varias - Anulación / Control Diferencia de Cambio"
         Top             =   4900
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
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
         Height          =   690
         Left            =   105
         Picture         =   "FORFASQL.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   140
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   100
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORFASQL.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   139
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   2860
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORFASQL.frx":0BB2
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   5520
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORFASQL.frx":0FF4
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Acceso a Listas de Precio, Condiciones de Pago y Tasas de IVA"
         Top             =   4265
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORFASQL.frx":12FE
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Remitos Pendientes de Facturación"
         Top             =   2220
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORFASQL.frx":1608
         Style           =   1  'Graphical
         TabIndex        =   135
         ToolTipText     =   "Acceso a Bases de Datos Maestras"
         Top             =   1560
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "FORFASQL.frx":1912
         Style           =   1  'Graphical
         TabIndex        =   134
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   790
         Visible         =   0   'False
         Width           =   650
      End
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FFFF80&
      Height          =   545
      Left            =   7080
      Picture         =   "FORFASQL.frx":1C1C
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Pedidos pendientes de Despacho, Remito o Factura"
      Top             =   1045
      Width           =   550
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      BorderStyle     =   0  'None
      Height          =   8595
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   13455
      Begin VB.CheckBox chkModoCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Modo Crédito"
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
         Left            =   9120
         TabIndex        =   173
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command31 
         Caption         =   "!"
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
         Left            =   1200
         TabIndex        =   172
         ToolTipText     =   "Solo Clientes con Pedidos Pendientes de Facturación"
         Top             =   1035
         Width           =   180
      End
      Begin VB.PictureBox MARCONOTA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   3190
         Left            =   50
         ScaleHeight     =   3165
         ScaleWidth      =   12315
         TabIndex        =   158
         Top             =   1785
         Visible         =   0   'False
         Width           =   12350
         Begin VB.CommandButton Command27 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Rep."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   11685
            Style           =   1  'Graphical
            TabIndex        =   166
            ToolTipText     =   "Facturación de Reparaciones"
            Top             =   1875
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton Command26 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   11685
            Style           =   1  'Graphical
            TabIndex        =   165
            ToolTipText     =   "Facturacion de Bobinas"
            Top             =   1275
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton Command16 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   11685
            Picture         =   "FORFASQL.frx":205E
            Style           =   1  'Graphical
            TabIndex        =   164
            ToolTipText     =   "Imprime Ficha de Anotaciones"
            Top             =   2475
            Width           =   570
         End
         Begin VB.CommandButton Command15 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   11685
            Picture         =   "FORFASQL.frx":26C8
            Style           =   1  'Graphical
            TabIndex        =   163
            ToolTipText     =   "Cancela la Factura Activa"
            Top             =   675
            Width           =   570
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   11685
            Picture         =   "FORFASQL.frx":29D2
            Style           =   1  'Graphical
            TabIndex        =   162
            ToolTipText     =   "Continúa Generación de la Factura"
            Top             =   75
            Width           =   570
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C4F9FD&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2820
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   161
            Top             =   360
            Width           =   11580
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   0
            ScaleHeight     =   405
            ScaleWidth      =   11565
            TabIndex        =   159
            Top             =   -30
            Width           =   11595
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "ANOTACIONES DE CLIENTES"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   0
               TabIndex        =   160
               Top             =   105
               Width           =   10545
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   3165
         Left            =   45
         TabIndex        =   157
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
         Top             =   1785
         Width           =   12350
         _ExtentX        =   21775
         _ExtentY        =   5583
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2D3C0&
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
         ForeColor       =   &H00000000&
         Height          =   1550
         Left            =   40
         TabIndex        =   113
         Top             =   6015
         Width           =   5925
         Begin VB.TextBox Text19 
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
            Height          =   300
            Left            =   1155
            TabIndex        =   118
            Top             =   1170
            Width           =   645
         End
         Begin VB.TextBox Text20 
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
            Height          =   300
            Left            =   1155
            TabIndex        =   121
            Top             =   860
            Width           =   645
         End
         Begin VB.CommandButton Command25 
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
            Left            =   1785
            TabIndex        =   120
            Top             =   860
            Width           =   180
         End
         Begin VB.CommandButton Command24 
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
            Left            =   1785
            TabIndex        =   119
            Top             =   1170
            Width           =   180
         End
         Begin VB.TextBox LENTREGA 
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
            Index           =   3
            Left            =   1575
            MaxLength       =   32
            TabIndex        =   117
            Top             =   1800
            Visible         =   0   'False
            Width           =   3750
         End
         Begin VB.TextBox LENTREGA 
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
            Height          =   300
            Index           =   0
            Left            =   1155
            MaxLength       =   48
            TabIndex        =   116
            Top             =   220
            Width           =   4590
         End
         Begin VB.TextBox LENTREGA 
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
            Height          =   300
            Index           =   1
            Left            =   1155
            MaxLength       =   48
            TabIndex        =   115
            Top             =   540
            Width           =   4590
         End
         Begin VB.TextBox LENTREGA 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   2050
            MaxLength       =   48
            TabIndex        =   114
            Top             =   1200
            Width           =   3690
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2055
            TabIndex        =   125
            Top             =   855
            Width           =   3690
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal"
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
            Index           =   13
            Left            =   -840
            TabIndex        =   124
            Top             =   950
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Domicilio"
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
            Index           =   12
            Left            =   -840
            TabIndex        =   123
            Top             =   320
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Transporte"
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
            Index           =   11
            Left            =   -840
            TabIndex        =   122
            Top             =   1240
            Width           =   1905
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2D3C0&
         Caption         =   "CONTABLE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   620
         Left            =   40
         TabIndex        =   104
         Top             =   7800
         Width           =   5925
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
            Left            =   2310
            TabIndex        =   109
            Top             =   240
            Width           =   180
         End
         Begin VB.TextBox Text16 
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
            Height          =   300
            Left            =   1155
            TabIndex        =   108
            Top             =   240
            Width           =   1170
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1575
            TabIndex        =   107
            Top             =   1130
            Width           =   180
         End
         Begin VB.TextBox TEXT21 
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
            Index           =   3
            Left            =   1155
            TabIndex        =   106
            Text            =   " "
            Top             =   1130
            Width           =   435
         End
         Begin VB.TextBox Eco21 
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
            Height          =   300
            Index           =   3
            Left            =   1890
            TabIndex        =   105
            Text            =   " "
            Top             =   1130
            Width           =   3270
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2595
            TabIndex        =   112
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta"
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
            Index           =   5
            Left            =   -525
            TabIndex        =   111
            Top             =   320
            Width           =   1590
         End
         Begin VB.Label Label21 
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
            Height          =   330
            Index           =   3
            Left            =   -3045
            TabIndex        =   110
            Top             =   1205
            Width           =   4110
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2D3C0&
         Caption         =   "RESUMEN FINAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3465
         Left            =   6810
         TabIndex        =   61
         Top             =   5000
         Width           =   5580
         Begin VB.TextBox Text26 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4320
            TabIndex        =   177
            Top             =   2760
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.TextBox Text25 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1390
            TabIndex        =   170
            Top             =   2400
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox Text24 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2370
            TabIndex        =   168
            Top             =   2400
            Width           =   960
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   3720
            TabIndex        =   71
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1365
            Width           =   1650
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   3720
            TabIndex        =   73
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1910
            Width           =   1650
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   3735
            TabIndex        =   62
            TabStop         =   0   'False
            Text            =   " "
            Top             =   3100
            Width           =   1650
         End
         Begin VB.TextBox PORDESCU 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2280
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   480
            Width           =   1065
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   2
            Left            =   3705
            TabIndex        =   77
            TabStop         =   0   'False
            Text            =   " "
            Top             =   795
            Width           =   1650
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   1
            Left            =   3705
            TabIndex        =   76
            TabStop         =   0   'False
            Text            =   " "
            Top             =   525
            Width           =   1650
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   3705
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   240
            Width           =   1650
         End
         Begin VB.TextBox Text23 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2280
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   480
            Width           =   465
         End
         Begin VB.TextBox TLABEL28 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2265
            TabIndex        =   70
            Top             =   1090
            Width           =   1395
         End
         Begin VB.TextBox BASIVA 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   2265
            TabIndex        =   69
            Top             =   1365
            Width           =   1395
         End
         Begin VB.TextBox BASIVA 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   2265
            TabIndex        =   68
            Top             =   1635
            Width           =   1395
         End
         Begin VB.TextBox BASIVA 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   2265
            TabIndex        =   67
            Top             =   1920
            Width           =   1395
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   3345
            TabIndex        =   65
            Top             =   2400
            Width           =   960
         End
         Begin VB.TextBox Text17 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   4320
            TabIndex        =   64
            Top             =   2400
            Width           =   1050
         End
         Begin VB.TextBox Text18 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2595
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   3100
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   3720
            TabIndex        =   72
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1635
            Width           =   1650
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Factura de Anticipo"
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
            Left            =   3720
            TabIndex        =   144
            Top             =   1080
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Modo Test"
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
            Left            =   3285
            TabIndex        =   66
            Top             =   1080
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "IB TUCUMAN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3240
            TabIndex        =   178
            Top             =   2850
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "IB MISIONES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1390
            TabIndex        =   171
            Top             =   2200
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "IB CABA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2370
            TabIndex        =   169
            Top             =   2205
            Width           =   1200
         End
         Begin VB.Label PORIVA 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Index           =   0
            Left            =   840
            TabIndex        =   103
            Top             =   1350
            Width           =   645
         End
         Begin VB.Label PORIVA 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Index           =   1
            Left            =   840
            TabIndex        =   102
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label PORIVA 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Index           =   2
            Left            =   840
            TabIndex        =   101
            Top             =   1935
            Width           =   645
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dto."
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
            Index           =   1
            Left            =   1150
            TabIndex        =   100
            Top             =   570
            Width           =   585
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto sin IVA"
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
            Index           =   0
            Left            =   1875
            TabIndex        =   99
            Top             =   285
            Width           =   1740
         End
         Begin VB.Label Ltota 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   3390
            TabIndex        =   98
            Top             =   570
            Width           =   225
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Importe Total Neto"
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
            Index           =   10
            Left            =   1980
            TabIndex        =   97
            Top             =   840
            Width           =   1635
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. Items"
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
            Left            =   190
            TabIndex        =   96
            Top             =   250
            Width           =   885
         End
         Begin VB.Label IEMBALA 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   285
            TabIndex        =   95
            Top             =   795
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   195
            TabIndex        =   94
            Top             =   450
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   615
            TabIndex        =   93
            Top             =   450
            Width           =   435
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IVA-NI"
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
            Index           =   5
            Left            =   -100
            TabIndex        =   92
            Top             =   1950
            Width           =   960
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IVA"
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
            Index           =   4
            Left            =   -210
            TabIndex        =   91
            Top             =   1665
            Width           =   1065
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IVA"
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
            Index           =   2
            Left            =   0
            TabIndex        =   90
            Top             =   1365
            Width           =   855
         End
         Begin VB.Label Ltota 
            BackStyle       =   0  'Transparent
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
            Height          =   330
            Index           =   7
            Left            =   1490
            TabIndex        =   89
            Top             =   1365
            Width           =   225
         End
         Begin VB.Label Ltota 
            BackStyle       =   0  'Transparent
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
            Height          =   330
            Index           =   8
            Left            =   1490
            TabIndex        =   88
            Top             =   1665
            Width           =   225
         End
         Begin VB.Label Ltota 
            BackStyle       =   0  'Transparent
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
            Height          =   330
            Index           =   9
            Left            =   1490
            TabIndex        =   87
            Top             =   1950
            Width           =   225
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IVA No Gravado"
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
            Left            =   60
            TabIndex        =   86
            Top             =   1110
            Width           =   1590
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "IB BSAS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3345
            TabIndex        =   85
            ToolTipText     =   "Doble Click para Configurar Cliente"
            Top             =   2205
            Width           =   1080
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Percep IVA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4320
            TabIndex        =   84
            Top             =   2205
            Width           =   1065
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cambio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   83
            ToolTipText     =   "Doble Click para Configurar Cliente"
            Top             =   2235
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label TICAMBIS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   120
            TabIndex        =   82
            Top             =   2445
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "T-C:"
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
            Left            =   105
            TabIndex        =   81
            Top             =   945
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label MONEMIS 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   120
            TabIndex        =   80
            ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
            Top             =   2730
            Width           =   1380
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL DOCUMENTO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   1920
            TabIndex        =   79
            Top             =   3160
            Width           =   1635
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2D3C0&
         Caption         =   "CONDICIONES COMERCIALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   990
         Left            =   40
         TabIndex        =   48
         Top             =   5000
         Width           =   5160
         Begin VB.TextBox txtFechaVencimiento 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   4020
            MaxLength       =   10
            TabIndex        =   175
            Top             =   630
            Width           =   885
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
            Height          =   300
            Left            =   4920
            TabIndex        =   174
            Top             =   640
            Width           =   195
         End
         Begin VB.TextBox TEXT21 
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
            Height          =   300
            Index           =   0
            Left            =   1470
            TabIndex        =   50
            Text            =   " "
            Top             =   290
            Width           =   435
         End
         Begin VB.TextBox Eco21 
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
            Height          =   300
            Index           =   2
            Left            =   1890
            TabIndex        =   57
            Text            =   " "
            Top             =   1130
            Width           =   3270
         End
         Begin VB.TextBox TEXT21 
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
            Index           =   2
            Left            =   1155
            TabIndex        =   56
            Text            =   " "
            Top             =   1130
            Width           =   435
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   1575
            TabIndex        =   55
            Top             =   1130
            Width           =   180
         End
         Begin VB.TextBox Eco21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2130
            TabIndex        =   54
            Text            =   " "
            Top             =   630
            Width           =   1500
         End
         Begin VB.TextBox TEXT21 
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
            Height          =   300
            Index           =   1
            Left            =   1470
            TabIndex        =   53
            Text            =   " "
            Top             =   630
            Width           =   435
         End
         Begin VB.CommandButton Command21 
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
            Index           =   1
            Left            =   1920
            TabIndex        =   52
            Top             =   630
            Width           =   180
         End
         Begin VB.TextBox Eco21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   2130
            TabIndex        =   51
            Text            =   " "
            Top             =   290
            Width           =   2910
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1920
            TabIndex        =   49
            Top             =   290
            Width           =   180
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Vto"
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
            Index           =   6
            Left            =   3720
            TabIndex        =   176
            Top             =   730
            Width           =   270
         End
         Begin VB.Label Label21 
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
            Height          =   330
            Index           =   2
            Left            =   -3045
            TabIndex        =   60
            Top             =   1205
            Width           =   4110
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Condición de Pago"
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
            Index           =   1
            Left            =   100
            TabIndex        =   59
            Top             =   730
            Width           =   1335
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lista de Precios"
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
            Index           =   0
            Left            =   100
            TabIndex        =   58
            Top             =   400
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture999 
         Height          =   540
         Index           =   0
         Left            =   2280
         ScaleHeight     =   480
         ScaleWidth      =   1110
         TabIndex        =   46
         Top             =   420
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7800
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   1515
         Begin VB.CommandButton Command22 
            Caption         =   "Cred"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   820
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   200
            Width           =   650
         End
         Begin VB.CommandButton Command23 
            Caption         =   "N/Deb"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   50
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   200
            Width           =   650
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   3195
         Left            =   50
         ScaleHeight     =   3165
         ScaleWidth      =   12315
         TabIndex        =   37
         Top             =   1785
         Visible         =   0   'False
         Width           =   12350
         Begin VB.ListBox List3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2730
            Left            =   0
            TabIndex        =   38
            Top             =   420
            Visible         =   0   'False
            Width           =   12350
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Label17"
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
            Left            =   45
            TabIndex        =   39
            Top             =   105
            Width           =   11430
         End
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   10290
         TabIndex        =   32
         ToolTipText     =   "Doble-Click para Facturar por Ordenes de Compra de Clientes"
         Top             =   1395
         Width           =   2115
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   9000
         TabIndex        =   31
         Text            =   "Text13"
         Top             =   8160
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.PictureBox MARSELPENDE 
         BackColor       =   &H0080C0FF&
         Height          =   3180
         Left            =   210
         ScaleHeight     =   3120
         ScaleWidth      =   10455
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   10520
         Begin VB.ListBox List2 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2760
            Left            =   105
            Style           =   1  'Checkbox
            TabIndex        =   25
            Top             =   420
            Width           =   9650
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   3165
            Left            =   0
            TabIndex        =   28
            Top             =   105
            Width           =   960
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H0080C0FF&
            Height          =   470
            Left            =   -105
            ScaleHeight     =   405
            ScaleWidth      =   9795
            TabIndex        =   23
            Top             =   -30
            Width           =   9850
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Selección de Pedidos Pendientes de Entrega"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   260
               Left            =   0
               TabIndex        =   24
               Top             =   105
               Width           =   9825
            End
         End
         Begin VB.CommandButton Command19 
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
            Height          =   570
            Left            =   9820
            Picture         =   "FORFASQL.frx":2CDC
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Continúa Generación de la Factura"
            Top             =   420
            Width           =   570
         End
         Begin VB.CommandButton Command18 
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
            Height          =   570
            Left            =   9820
            Picture         =   "FORFASQL.frx":2FE6
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Cancela la Factura Activa"
            Top             =   1050
            Width           =   570
         End
         Begin VB.Line Line5 
            X1              =   9750
            X2              =   9750
            Y1              =   0
            Y2              =   3200
         End
         Begin VB.Line Line4 
            X1              =   10575
            X2              =   10575
            Y1              =   0
            Y2              =   3150
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   -120
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1410
         Width           =   3480
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   11280
         TabIndex        =   11
         Text            =   " "
         ToolTipText     =   "Este Número Puede Sufrir Modificaciones segun el Servidor de la A.F.I.P."
         Top             =   675
         Width           =   1120
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   11280
         MaxLength       =   16
         TabIndex        =   12
         Text            =   " "
         Top             =   1035
         Width           =   1120
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   11280
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   300
         Width           =   960
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
         Height          =   330
         Left            =   12240
         TabIndex        =   10
         Top             =   300
         Width           =   195
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5145
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         ToolTipText     =   "Doble Click para Ajustar Limite de Credito con Privilegios"
         Top             =   600
         Width           =   1380
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5145
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1395
         Width           =   1380
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2310
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1035
         Width           =   2640
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1035
         Width           =   750
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   675
         Width           =   3480
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   300
         Width           =   3480
      End
      Begin VB.CommandButton Command1 
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
         Height          =   330
         Left            =   930
         TabIndex        =   1
         Top             =   1035
         Width           =   180
      End
      Begin VB.TextBox Text1 
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
         Height          =   330
         Left            =   90
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   1035
         Width           =   855
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2730
         Left            =   210
         TabIndex        =   19
         Top             =   2205
         Width           =   10515
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   210
         ScaleHeight     =   405
         ScaleWidth      =   10485
         TabIndex        =   18
         Top             =   1785
         Width           =   10515
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "R."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10230
            TabIndex        =   156
            Top             =   140
            Width           =   855
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Importe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9310
            TabIndex        =   155
            Top             =   140
            Width           =   855
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "P.Unitario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7710
            TabIndex        =   154
            Top             =   140
            Width           =   855
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Desc."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7110
            TabIndex        =   153
            Top             =   140
            Width           =   855
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Lista"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6450
            TabIndex        =   152
            Top             =   140
            Width           =   855
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2790
            TabIndex        =   151
            Top             =   135
            Width           =   1335
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   880
            TabIndex        =   150
            Top             =   140
            Width           =   855
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   50
            TabIndex        =   149
            Top             =   140
            Width           =   855
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "  Cant Código          Descripción                  Lista  Desc  P.Unit     Importe  R."
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   225
            Left            =   30
            TabIndex        =   148
            Top             =   0
            Visible         =   0   'False
            Width           =   10515
         End
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   0
         OleObjectBlob   =   "FORFASQL.frx":32F0
         Top             =   1680
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   195
         Index           =   0
         Left            =   2880
         OleObjectBlob   =   "FORFASQL.frx":3524
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Frame Frame12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Com."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   950
         Left            =   5250
         TabIndex        =   126
         Top             =   5040
         Width           =   705
         Begin VB.TextBox Text22 
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
            Height          =   400
            Left            =   80
            MaxLength       =   5
            TabIndex        =   132
            Top             =   280
            Width           =   570
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   1575
            TabIndex        =   129
            Top             =   1130
            Width           =   180
         End
         Begin VB.TextBox TEXT21 
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
            Index           =   4
            Left            =   1155
            TabIndex        =   128
            Text            =   " "
            Top             =   1130
            Width           =   435
         End
         Begin VB.TextBox Eco21 
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
            Height          =   300
            Index           =   4
            Left            =   1890
            TabIndex        =   127
            Text            =   " "
            Top             =   1130
            Width           =   3270
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   80
            TabIndex        =   131
            Top             =   280
            Width           =   540
         End
         Begin VB.Label Label21 
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
            Height          =   330
            Index           =   4
            Left            =   -3045
            TabIndex        =   130
            Top             =   1205
            Width           =   4110
         End
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "TEPRUEBA"
         Height          =   300
         Left            =   0
         TabIndex        =   167
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Facturación de Ventas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   -30
         TabIndex        =   147
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label30 
         Height          =   255
         Left            =   960
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   90
         TabIndex        =   36
         Top             =   825
         Width           =   960
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   90
         TabIndex        =   35
         Top             =   1410
         Width           =   1035
      End
      Begin VB.Label DOCUDIF 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   5040
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Compra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   8760
         TabIndex        =   33
         ToolTipText     =   "Doble-Click para Facturar por Ordenes de Compra de Clientes"
         Top             =   1515
         Width           =   1440
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   7335
         X2              =   7335
         Y1              =   945
         Y2              =   1910
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Height          =   765
         Left            =   7020
         Top             =   135
         Width           =   645
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   7020
         TabIndex        =   30
         Top             =   135
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   9000
         TabIndex        =   29
         Top             =   975
         Width           =   1500
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factura:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   9000
         TabIndex        =   27
         Top             =   615
         Width           =   1500
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0001"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   10680
         TabIndex        =   26
         ToolTipText     =   "Doble-Click Para Cambiar Punto de Venta"
         Top             =   675
         Width           =   585
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito Disponible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   5145
         TabIndex        =   17
         Top             =   1150
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   10680
         TabIndex        =   16
         Top             =   1170
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   10620
         TabIndex        =   15
         Top             =   435
         Width           =   585
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Límite de Crédito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   5145
         TabIndex        =   14
         Top             =   360
         Width           =   1380
      End
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label31"
      Height          =   225
      Left            =   4305
      TabIndex        =   45
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuFacturaElectronica 
         Caption         =   "Facturación Electrónica"
         Begin VB.Menu mnuObtenerToken 
            Caption         =   "Obtener Token"
         End
         Begin VB.Menu mnuSolicitudManualCae 
            Caption         =   "Solicitud de CAE Manual"
         End
      End
      Begin VB.Menu separa0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAnulacionComprobantes 
         Caption         =   "Anulación de Comprobantes"
      End
      Begin VB.Menu mnuEliminaCbtesSinCae 
         Caption         =   "Eliminación de Cbtes. Anulados SIN CAE"
      End
      Begin VB.Menu separa1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCorregirFechaCbte 
         Caption         =   "Corregir Fecha de Emisión de Comprobante"
      End
      Begin VB.Menu mnuCorregirNumeroCbte 
         Caption         =   "Corregir Número de Comprobante "
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOFP 
         Caption         =   "Opciones de Factura Proforma"
         Begin VB.Menu mnuFActProfMatCod 
            Caption         =   "Factura Proforma x Material Codificado"
         End
         Begin VB.Menu FactProfTlibre 
            Caption         =   "Factura Proforma Texto Libre - Consultas y Listados"
         End
      End
      Begin VB.Menu lin100 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompensacionClienteProveedor 
         Caption         =   "Compensación Cliente-Proveedor"
      End
   End
   Begin VB.Menu mnuConsultasyListados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuRemitosPendientesFacturacion 
         Caption         =   "Remitos Pendientes de Facturación"
      End
      Begin VB.Menu mnuPedidosPendientes 
         Caption         =   "Pedidos Pendientes de Entrega"
      End
      Begin VB.Menu mnuControlGeneral 
         Caption         =   "Control General de Gestión Comercial"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConfiguración 
      Caption         =   "Configuración"
      Begin VB.Menu mnuTasasIva 
         Caption         =   "Tabla de Tasas de IVA"
      End
      Begin VB.Menu mnuCondicionesDePago 
         Caption         =   "Condiciones de Pago"
      End
      Begin VB.Menu mnuParametrosVarios 
         Caption         =   "Parámetros Varios"
      End
      Begin VB.Menu mnuCotizacionesMoneda 
         Caption         =   "Cotizaciones de Moneda Extranjera"
      End
      Begin VB.Menu mnuMotivosImputacion 
         Caption         =   "Tabla de Motivos de Imputación"
      End
      Begin VB.Menu mnuCodigosEquivalentesClientes 
         Caption         =   "Códigos Equivalentes de Clientes"
      End
      Begin VB.Menu mnuSetupCompensacionClienteProveedor 
         Caption         =   "Compensación Cliente-Proveedor"
         Begin VB.Menu mnuTablaCompensacionClienteProveedor 
            Caption         =   "Tabla de Compensacion Cliente-Proveedor"
         End
         Begin VB.Menu mnuTablaCompensacionImpuestos 
            Caption         =   "Tabla de Compensación de Impuestos"
         End
      End
      Begin VB.Menu separa6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfiguracionGeneral 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu mnuAccesosDirectos 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuListasDePrecios 
         Caption         =   "Listas de Precios"
      End
      Begin VB.Menu mnuMaestroClientes 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu mnuMaestroArticulos 
         Caption         =   "Maestro de Artículos de Stock"
      End
   End
End
Attribute VB_Name = "FORFASQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TIFADOC%    ' 2=FACTURA / 3=N.DEBITO / 6=N.CREDITO
Dim MODODOC%    ' 0=MAT.CODIFICADO / 1=TEXTO LIBRE
Dim VARIDOC%
Dim PIVADOC%
Dim RNIDOC%
Dim DATEDOC%
Dim VERLICRE%
Dim MODESCU%
Dim FVENCIM
Dim DEBIMANUAL%
'
Dim CONTROFIS% ' Si 0 --> No Inicializado /// 1 --> Controlador Fiscal /// -1 --> Printer Común
Dim FACELEC%
Dim ModoExportaSimple%
Dim PuntoVentaExportaSimple$
Dim NumeroExportaSimple$
Dim CAEExportaSimple$
Dim VtoCAEExportaSimple%
Dim YAIMPRIMIO%
'
Public CONSICLI%
Public ESPROFORMA& ' CONTIENE EL NRO DE PROFORMA SI ESTA ACTIVA
Dim PREFACTU$
'
Dim DESELEC As String, HASELEC As String
'
Dim NUOCOMPRA$
Dim NROPEDIDO$
Dim NROREMICO$
'
Dim ARCHIPRE$
Dim DATCLI$(32)
'
Dim NFCAM%    ' flag (default=0) que se pone en 1 si se ha cambiado manualmente en nro de factura
'
Dim MONECO$, TICAMBX#
Dim NOCAMONE%
Dim APLIVA%, RNI%
'
Dim COPIASREM%
'
Dim TASACLICK
Dim FACTLISTAEMPAQUE& ' 0 SI NO SE ESTA FACTURANDO LISTA DE EMPAQUE O CASO CONTRARIO EL NUMERO DE LA MISMA
Dim FACTUANTICIPO%
Dim FacturaSeleccionada$
Public AnulaFacturaComprobante$
Dim AplicacionFacturaCredito$
Dim TipoAplicacionFacturaCredito$

Dim VentaNetaPorCuentaContable#()
Dim NetoMultipleImputacion%

Dim PercepcionIBMisiones_BienDeUso As Integer
Sub Aplica_Fact_Anticipo(NC&)
    '
    If TIFADOC% > 3 Then Exit Sub ' SOLO PARA FACTURA O NOTA DE DEBITO
    Campos$ = "F.Emis/D8;N.Cte./A24;Saldo/F12.2"
    Call CARGA_ENCABEZADO(ANTSINAP09.MSF, Campos$, ANTSINAP09)
    Campos$ = Campos$ & ";Imp.Ajust/F12.2;Iva 10.50 %/F12.2;Grav. 10.5%/F12.2;Iva 21.00 %/F12.2;Grav. 21.0%/F12.2;Exento/F12.2"
    Call CARGA_ENCABEZADO(ANTSINAP09.MSF2, Campos$, ANTSINAP09)
    '
    SQLX$ = "select  sum(impodebe - impohabe)  as saldo, DOCU_ORIG, FechEmisi,VARIMOVIM,MONE_EMIS,COTIZA_DOLAR "
    SQLX$ = SQLX$ + " From factuanti  WITH(NOLOCK)"
    SQLX$ = SQLX$ + " Where NuClien = " & NC&
    SQLX$ = SQLX$ + " AND (TIPOMOVIM = 'FVEN' or TIPOMOVIM = 'AJUS')"
    SQLX$ = SQLX$ + " group by  DOCU_ORIG, FechEmisi,VARIMOVIM ,MONE_EMIS,COTIZA_DOLAR"
    SQLX$ = SQLX$ + " Having sum(impodebe - impohabe) > 0.005 "
    Set rs = READSET(SQLX$)
    '
    i& = 0: ACUMULASALDO = 0
    ANTSINAP09.MSF.Rows = 1
    With rs
        Do While Not .EOF
           NUFACTU$ = SAFETEXT$(!DOCU_ORIG)
           MON% = XVALO(!Mone_Emis)
           SALDOX# = XVALO(!SALDO)
           
           ACUMULASALDO = ACUMULASALDO + SALDOX#
           TIPOIVA$ = SAFETEXT$(!VARIMOVIM)
           FE% = CVINT(!FECHEMISI)
           '
           i& = i& + 1
           ANTSINAP09.MSF.Rows = ANTSINAP09.MSF.Rows + 1
           ANTSINAP09.MSF.TextMatrix(i&, 1) = FECHATEX$(FE%)
           ANTSINAP09.MSF.TextMatrix(i&, 2) = NUFACTU$
           ANTSINAP09.MSF.TextMatrix(i&, 3) = FORMATNUM$(SALDOX#, "F12.2")
           .MoveNext
         Loop
    End With
    rs.Close
    Set rs = Nothing
    ANTSINAP09.Label8 = Total(6)
    ANTSINAP09.TLABEL28 = TLABEL28
    ANTSINAP09.BASIVA(0) = BASIVA(0)
    ANTSINAP09.BASIVA(1) = BASIVA(1)
    ANTSINAP09.PORIVA(0) = PORIVA(0)
    ANTSINAP09.PORIVA(1) = PORIVA(1)
    ANTSINAP09.PORIVA(3) = PORIVA(1)
    ANTSINAP09.PORIVA(2) = PORIVA(0)
    
    ANTSINAP09.Total(3) = Total(3)
    ANTSINAP09.Total(4) = Total(4)
    'VALIDO SI TIENE FACTURAS ANTICIPO EN ESE CASO MENSAJE SI NO ACEPTA DESCARGA EL FRM
    If i& > 0 Then
       TX$ = "Este Cliente Posee Factura/s Anticipo con Saldo "
       OPX% = COMALTER%(TX$ & " \\Aplicar\Continuar sin Aplicar")
       If OPX% < 1 Or OPX% >= 2 Then
          Unload ANTSINAP09
          Exit Sub
       Else
          ANTSINAP09.Show 1
       End If
    End If
    '
End Sub


Sub AplicaNotaDeCredito(NFACT$, NFCY$)
        NCLI = XVALO(Text1)
        For i% = 1 To 4
           IMPOAPLICAR# = XVALO(Total(6))
           TiMovi1$ = "AJUS"
           BASEIMPODE# = 0
           BASEIMPOHA# = 0
           VaMovi1$ = Right$(NFACT$, 1)
           Set RS10 = New ADODB.Recordset
           SQLX$ = "SELECT * FROM FACTUANTI WHERE DOCU_ORIG = '' "
333        On Error Resume Next
           RS10.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 333
           End If
           On Error GoTo 0
           '
           PROAJ& = PROXIMO_AJUSTE_FACTUANTI&
           With RS10
                  .AddNew
                  On Error Resume Next
                  !CODIEMPR = CodiEmp%
                  On Error GoTo 0

                  !DOCU_ORIG = NFACT$
                  !VARIMOVIM = VaMovi1$
                  !DOCU_APLI = NFCY$
                  !IMPODEBE = 0
                  !IMPOHABE = IMPOAPLICAR#
                  !NUCLIEN = Int(GRUDIVCLI(NCLI))
                  !Cli_Real = NCLI
                  !refecom = rasocli$(NCLI)
                  !NumeCom = XVALO(Mid$(NFCY$, 9, 8))
                  !FECHEMISI = FECHATEX$(HOY(HO$))
                  !Status = 0
                  !TIPOMOVIM = TiMovi1$
                  !COTIZA_DOLAR = COTI_DOLAR#
                  !BASEIMPODEBE = 0
                  !BASEIMPOHABE = BASEIMPOHA#
                  !USUARIO = USERNAME$
                  !OBSERVACIONES = "N/C. A FACTURA ANTICIPO"
                  !AJUSTE_NUMERO = PROAJ&
                  !Mone_Emis = MONEMI%
                  .Update
           End With
900      Next i%
End Sub

Sub AppProformasFacturadasDePedidoActivo(NC, PEDIDOX$)
   If NC < 1 Then Exit Sub
   If TRIM$(PEDIDOX$) = "" Then Exit Sub
   ATXX% = 0
   'RECORRO EL STRING CON LOS NROS DE PEDIDOS Y LEVANTO LAS PROFORMAS FACTURADAS QUE MACHEEN.
   Call TEXTOLOTES$(PEDIDOX$, ";")
   For i% = 1 To UBound(CADENATEX$)
      NPESEL& = XVALO(CADENATEX$(i%))
      If NPESEL& > 0 Then
         If ATXX% = 1 Then
             CondicionSql$ = CondicionSql$ + "(NroPed = " & NPESEL& & " "
             ATXX% = 1
         Else
             CondicionSql$ = CondicionSql$ + "OR NroPed = " & NPESEL & " "
         End If
      End If
   Next i%
   
   If TRIM$(CondicionSql$) = "" Then Exit Sub
   
   CondicionSql$ = CondicionSql$ + ")"
   
'   SQLX$ = "SELECT * FROM PROFORMTL WITH(NOLOCK) WHERE " & CondicionSql$ & " ORDER BY NUMECOMP "
'   Set rs = READSET(SQLX$)
'   With rs
'     Do While Not .EOF
'        NRO& = XVALO(!NumeComp)
'        REFE$ = SAFETEXT$(!REFERENCIA)
'        NROP& = XVALO(!NROPED)
'        MON% = XVALO(!MONE_EMIS)
'        DENOMI$ = DEMONECO$(MON%)
'        FE% = CVINT(!FECHIMISI)
'        FEX$ = FECHATEX$(FE%)
'        TTNETO# = XVALO(!TOTNETO)
'        TTN = FORMATNUM$(TTNETO#, "F10.2")
'        .MoveNext
'     Loop
'   rs.Close: Set rs = Nothing
'
End Sub

Sub CargaDetalleMatCodProforma(NPROF&)
   'FALTA CARGAR ENCABEZADO
   SQLX$ = "SELECT * FROM DETAPROFO WITH(NOLOCK) WHERE NUMECOMP = " & NPROF&
   Set rs = READSET(SQLX$)
   With rs
     '
     JJ& = 0: CAITCAR% = 0
     Do While (.EOF = False)
       CODIMAT% = XVALO(!CODIINT)
       If CODIMAT% > 0 And CODIMAT% <= NUMAS% Then
         CAN = XVALO(!CANPED) - XVALO(!CANTENT)
         If !TipoPed = "ABIERTO" Then CAPEN = 999999
         NPEDID& = XVALO(!NROPED)
         FEMIPED% = CVINT(!FECHEMIS)
         PRECI1# = XVALO(!PreUnid)
         'PRECI1# = ROUND(XVALO(!PreUnid), 3)'propuesta de modificacion ej ahp que pedido en u$s 1.65 x 50 importe 82.48 en vez de 82.50
         LOTEX1$ = SAFETEXT$(!idenlote)
         '
         'CALCULO DE RECARGO POR CONDICION DE PAGO
         RECXPAGO$ = VALOBADA("PEDENCA", "Rec_ConPa", "NROPED = " & NPEDID&, "NOMESS")
         CODESC = COEFDESCU(RECXPAGO$)
         If Abs(XVALO(CODESC)) < 0.05 Then CODESC = 0
         PRECI1# = PRECI1# + (PRECI1# * CODESC / 100)
         '
         MONEMA% = !MoneEmis
            If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
         PRELI = XVALO(!PreList) ' MODIFICADO PARA QUE TRAIGA EL PRECIO DE LISTA CON EL CUAL SE CREO EL PEDIDO
           If PRELI < PRECIO Then PRELI = PRECI1#
           
         PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECI1#) / PRELI
         PRECIO = ROUND(PRECI1#, 4) '
'         PRECIO = ROUND(PRECIO, 3) 'propuesta de modificacion ej ahp que pedido en u$s 1.65 x 50 importe 82.48 en vez de 82.50
         '
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = ""
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = 0
           If InStr(1, EMPREAC$, "J Y M") > 0 Then FEREMIT% = FENTRESOL%
         '
         ESTRAX% = ESTRAZABLE%(CODIMAT%)
         If ESTRAX% = 2 Then
            NUMSE$ = "" ' NUMERO DE SERIE ASIGNADO AL PEDIDO
            ORDENITEM% = XVALO(!NuOrItem)
            CONDICION$ = " NroPed = " & NPEDID& & " AND NuOrItem = " & ORDENITEM% & " AND CodiInt = " & CODIMAT% & ""
            NUMSE$ = VALOBADA("PEDDETA", "NUMEROSERIE", CONDICION$)
            Call REPLA(Z$, NUMSE$, 139, 64)
         'AGREGADO AHP
         ElseIf ESTRAX% = 1 Then
'           If UCase$(VENCARPE$) = "CARDETLO" Then ' CONFIGURACION DE VENTANA EN PEDIDOS PARA TRABAJAR CON LOTES AHP
                'CASO AHP QUE VIENE EL LOTE DEL  PEDIDO DE LOS TRAZABLES 1
                Call REPLA(Z$, LOTEX1$, 267, 16)
'                'VALIDACAION DE LOTES EN EL CASO QUE SE TRABAJE DE ESTE MODO DESDE PEDIDOS
'           End If
         End If
         'FIN AGREGADO AHP
         '
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(Z$, NUPLANO$(CODIMAT%), 315, 24)
         '
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           ' IMPO# = CDbl(PRECIO) * CAN
           IMPO# = ROUND(PRECIO * ROUND(CAN, 1), 2)
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
         '
         Call REPLA(Z$, " ", 461, 1)
         If Left$(UCase$(SAFETEXT$(!VenRep)), 1) = "R" Then
            Call REPLA(Z$, "R", 461, 1)
         End If
           '
           NORITPE% = !NuOrItem
         Call REPLA(Z$, MKI$(NORITPE%), 503, 2)
'         Call REPLA(Z$, MKS$(0), 505, 4)
         Call REPLA(Z$, MKS$(CAN), 505, 4)
         Call REPLA(Z$, MKS$(0), 509, 4)
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(Z$, NUOCOX$, 355, 16)
           If TRIM$(NUOCOX$) <> "" Then
             If InStr(Text14, TRIM$(NUOCOX$) + ",") < 1 Then
               Text14 = Text14 + NUOCOX$ + ","
             End If
           End If
           '
         JJ& = JJ& + 1
         Call GRAREG("!CARDEFAC", Z$, JJ&)
         CAITCAR% = CAITCAR% + 1
         '
       End If
     .MoveNext
     Loop
   End With
   rs.Close
   Set rs = Nothing
   Call SETULTREG("!CARDEFAC", JJ&)
   ESPROFORMA& = NPROF&
   

   
   MONEMX% = VALOBADA("PROFORMTL", "MONE_EMIS", "NUMECOMP  = " & NPROF&, "NOMESS")
   If MONEMX% < 1 Then MONEMX% = 1
   MONEMI% = MONEMX%
   MONEMIS.Caption = Left$(DEMONECO$(MONEMI%), 8)
   TICAM = XVALO(VALOBADA("PROFORMTL", "TIPO_CAM", "NUMECOMP  = " & NPROF&, "NOMESS"))
   'TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   TICAMBIS = FORMATNUM$(TICAM, "F9.4") 'MANTIENE EL DE LA PROFORMA
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAM, "F10.4")) + " - Doble-Click para Cambiar Moneda"
   MONEMIS.ForeColor = &H404080
   TICAMBIS.Visible = False
   Label24.Visible = False
   If MONEMI% > 1 Then
       MONEMIS.ForeColor = &H0&
       TICAMBIS.Visible = True
       Label24.Visible = True
   End If
   
   PORDES$ = TRIM$(VALOBADA("PROFORMTL", "TDescApli", "NUMECOMP  = " & NPROF&, "NOMESS")) 'MANTIENE EL DE LA PROFORMA
     PORDESCU = PORDES$
   Text19 = XVALO(VALOBADA("PEDENCA", "NTranspo", "NroPed=" & NPEDID&)) 'LEVANTA EL DEL PEDIDO
End Sub

Sub CARGADETALLEPROFORMA(NPROF&)
'        Me.TXTCLIENTE.Enabled = False
'        Unload FORFASQL 'SI NO DESCARGO EL FRM NO FUNCIONA LOS DOBLE CLICK HARDCODEADOS DEL LABEL7

        SQLX$ = "SELECT * FROM PROFORMTL WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE NUMECOMP = " & NPROF&
        Set rs = READSET(SQLX$)
        With rs
            CONDP% = XVALO(!CODICONPAG)
            MON% = XVALO(!Mone_Emis)
            IMPO_NETO = XVALO(!TOTNETO)
            VDDOR% = XVALO(!VEND_CLI)
            REFE$ = SAFETEXT$(!REFERENCIA)
            PORCENTDESC = XVALO(!TDescApli)
        End With
        rs.Close
        Set rs = Nothing
        CONDI$ = "NumeComp = " & NPROF&
        
        If CantRegistros&("DETAPROFO", CONDI$, "NOMESS") > 0 Then
           Call CargaDetalleMatCodProforma(NPROF&)
           
           Exit Sub
        End If
        
        VUELTA% = 0
        '
        
        'CARGA EL CARDEFAC TEXTO LIBRE
        Call SETULTREG("!CARDEFAC", 0)
        TX$ = REGBLAN$("!CARDEFAC")
        If MODODOC% = 0 Then Label6_Click
20      If Label6.Caption <> "(T.Libre)" And VUELTA% < 10 Then
             VUELTA% = VUELTA% + 1
             Label6_Click
             GoTo 20
        End If
        
        Call FRATEXTO(REFE$, 64)
        For KKII& = 1 To CARETEX%
            If KKII& = 1 Then
                Call REPLA(TX$, MKS$(1), 1, 4)
                Call REPLA(TX$, RETEX$(KKII&), 7, 64)
                Call REPLA(TX$, MKS(IMPO_NETO), 429, 4)
                IMPO_NE# = IMPO_NETO
                Call REPLA(TX$, MKD(IMPO_NE#), 433, 8)
                Call REPLA(TX$, MKS$(PORCEIVA(CI%)), 441, 4)
                Call REGAPP("!CARDEFAC", TX$)
            End If
            If KKII& > 1 Then '
               TX$ = REGBLAN$("!CARDEFAC")
               Call REPLA(TX$, MKS$(1), 1, 4)
               Call REPLA(TX$, RETEX$(KKII&), 7, 64)
               Call REGAPP("!CARDEFAC", TX$)
            End If
        Next KKII&
        '
        'CARGO CONDICION DE PAGO
        TEXT21(1) = ""
        If CONDP% > 0 Then
           TEXT21(1) = CONDP%
        End If
        Call SETULTREG("!CARDEFAC", KKII& - 1)
   MONEMI% = MON%
   MONEMIS.Caption = Left$(DEMONECO$(MONEMI%), 8)
   TICAM = XVALO(VALOBADA("PROFORMTL", "TIPO_CAM", "NUMECOMP  = " & NPROF&, "NOMESS"))
   'TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   TICAMBIS = FORMATNUM$(TICAM, "F9.4") 'MANTIENE EL DE LA PROFORMA
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAM, "F10.4")) + " - Doble-Click para Cambiar Moneda"
   MONEMIS.ForeColor = &H0&
   TICAMBIS.Visible = False
   Label24.Visible = False
   If MONEMI% > 1 Then
       MONEMIS.ForeColor = &H404080
       TICAMBIS.Visible = True
       Label24.Visible = True
   End If

End Sub


Function GRABASOLOREMITOREPARACION%()
'PONER REMITO 1 Y FACTURA 0
      'CASO DE EYM CUANDO TIENE QUE GENERAR UN REMITO DE UNA REPARACION SIN COSTO (GARANTIA)
      'DEBE SALIR CON EL REMITO CLASICO.SI LLEGA ACA ES POR QUE ES EYM, NO TIENE COSTO Y ES REPARACION
      NUREP& = XVALO(Label31) 'OJO QUE EN GRAREMIFA BORRA EL LABEL31
      GRABASOLOREMITOREPARACION% = 1 'SIEMPRE ESTA DEVOLVIENDO 1 POR QUE SI LLEGA ACA ES POR QUE CUMPLE
                                     'LAS CONDICIONES Y CUANDO VUELVE NO DEBE CONTINUAR
                                     
10    FINFACTUR07.Text7 = TRIM$(Str$(DEPOEXPE%))
      FINFACTUR07.Text2(0) = 0: FINFACTUR07.Text2(0).Enabled = True: FINFACTUR07.Text2(0).BackColor = &HE0E0E0
      FINFACTUR07.Text2(1) = 1
      Call APROCOPIAS(OKX%)
      If OKX% < 0 Then
        Exit Function
      End If
      '
      If COPIASFAC% > 0 Then
         If XVALO(Total(6).TEXT) < 0.01 Then
            Call MENSERR(24, "Imposible Emitir Factura SIN CARGO !!!")
            GoTo 10
            Exit Function
         End If
      End If
      '
      Call GRAREMIFA
      Call PRIREMIFA
      
      NPVX$ = "0001" 'RT-X-0001-00002375
      'RT.2398
      REMICORT$ = "RT." & TRIM$(Text10.TEXT)
      NroRemi$ = TRIM$(Text10.TEXT)
      While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
      NroRemi$ = "RT" & "-X-" & NPVX$ + "-" & NroRemi$
      '
      '
      NCLIE = XVALO(Text1.TEXT)
      FECHDOC$ = Text9.TEXT: FE% = FECHANUM(FECHDOC$): FECHDOC$ = FETEXCOR1$(FE%)
      
      'MARCA SIN CARGO EL REMITO
     SQLX$ = "UPDATE MOVISTO SET DOSECCOR = 'S/Cargo'"
     SQLX$ = SQLX$ + ", DoSecLar ='Sin Cargo'"
     SQLX$ = SQLX$ + ", VALOUNIT =0"
     SQLX$ = SQLX$ + ", VALOUNITIVA =0"
     SQLX$ = SQLX$ + ", VALTOTIT =0"
     SQLX$ = SQLX$ + ", ValTotItIva =0"
     'CONDICION
     SQLX$ = SQLX$ + " WHERE DoPriCor = '" & REMICORT$ & "'"
     SQLX$ = SQLX$ + " AND NUMPVTA= " & NPVX$
     SQLX$ = SQLX$ + " AND Nume_Clie = " & NCLIE & " "
     SQLX$ = SQLX$ + " AND FechMov = '" & FECHDOC$ & "'"
     FLEXCONN.Execute (FILTSQL$(SQLX$))
     '
     'ACTUALIZA EN REPARA
     CONDI$ = "NUMREPA = " & NUREP&
     FE% = HOY(HO$): FENT$ = FETEXCOR1$(FE%)
     SQLX$ = "SELECT NUMREMITO,FACTURA,STATUS,FECHENT FROM REPARA WHERE " & CONDI$
     Dim rs As ADODB.Recordset
     Set rs = New ADODB.Recordset
     rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     With rs
       !NumRemito = NroRemi$
       !FACTURA = NroRemi$ 'GRABA EL NRO DE REMITO EN CAMPO FACTURA POR QUE NO VA A TENER FACTURA
       !Status = 5
       !FECHENT = FENT$
       .Update
     End With
     rs.Close
     Set rs = Nothing
     
     'GRABAR EN TANUMSE
     'OBTENGO EL NUMERO DE SERIE DESDE LA REPARACION
     NUSERI$ = TRIM$(VALOBADA("REPARA", "NUMEROSERIE", CONDI$, "NOMESS"))
     SQLX$ = "SELECT * FROM TANUMSE "
     SQLX$ = SQLX$ + " WHERE  NUMEROSERIE = '" & NUSERI$ & "'"
     Dim RS3 As ADODB.Recordset
     Set RS3 = New ADODB.Recordset
     RS3.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     FE1% = DIASPOST(FE%, 180)
     With RS3
        !FevenGaRep = CVDAT(FE1%)
        !UnumRep = NUREP&
        .Update
     End With
     RS3.Close
     Set RS3 = Nothing
     '
     'HISTORIAL EN TANUMSE
     CUERPO$ = "Equipo Remitido" & vbCrLf
     CUERPO$ = CUERPO$ + "Número de Reparación: " & NUREP& & vbCrLf '
     CUERPO$ = CUERPO$ + "Número de Remito: " & NroRemi$ & vbCrLf
     CUERPO$ = CUERPO$ + "Cliente: " & rasocli$(NCLIE)
     Call ANOTANUSE(NUSERI$, CUERPO$, , "NOMESS")
     '
     GRABASOLOREMITOREPARACION% = 1

   '
 
End Function

Sub GuardarDetalleFacturaProforma(NPROF&, MONX%, TICAMX, CONP%, NC)
   FE% = HOY(HO$)
   SQLX$ = "SELECT * FROM DETAPROFO WHERE NUMECOMP = " & NPROF&
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With rs
     For i& = 1 To ULTREG&("!CARDEFAC")
       XX$ = REGLEIDO$("!CARDEFAC", i&)
       CODIMAT% = CVI(Mid$(XX$, 5, 2))
       If CODIMAT% > 0 And CODIMAT% <= NUMAS% Then
         .AddNew
         CAN = CVS(Mid$(XX$, 1, 4))
         
         DESX$ = Mid$(XX$, 7, 64)
         TPEDITX$ = Mid$(XX$, 93, 16)
         NUOCOX$ = Mid$(XX$, 355, 16)
         NPEDID& = XVALO(Mid$(XX$, 371, 16))
         NROREMX$ = Mid$(XX$, 387, 16)
         FEREMIT% = CVI(Mid$(XX$, 419, 2))
         PRELI = CVS(Mid$(XX$, 421, 4))
         PDESX = CVS(Mid$(XX$, 425, 4))
         PRECIO = CVS(Mid$(XX$, 429, 4))
         IMPO# = CVS(Mid$(XX$, 433, 8))
         PORXIVA = CVS(Mid$(XX$, 441, 4))
         NORIT% = CVI(Mid$(XX$, 503, 2))
         PRECIO = ROUND(PRECIO, 4)

         ESTRAX% = ESTRAZABLE%(CODIMAT%)
         UMLOT$ = ""
         If ESTRAX% = 2 Then
            NUMLOT$ = Mid$(XX$, 139, 64)
         ElseIf ESTRAX5 = 1 Then
            NUMLOT$ = Mid$(XX$, 267, 16)
         End If
         !CODIINT = CODIMAT%
         !NumeComp = NPROF&
         !CANPED = CAN
         !PreUnid = PRECIO
         !ImpoTot = PRECIO * CAN
         !TPedItem = TPEDITX$
         !NROPED = NPEDID&
         !FECHEMIS = CVDAT(FE%)
         !NROCLIE = NC
         !Status = 0
         !CondPag = CONP%
         !NuOrItem = NORIT%
         !CODIEXT = CODEXT$(CODIMAT%)
         !Descrip = DESCRIT0$(CODIMAT%)
         !MoneEmis = MONX%
         !TIPOCAM = TICAMX
         .Update
       End If
     Next i&
   End With
   rs.Close
   Set rs = Nothing

End Sub
'
Sub CARGAREPA()

'   Call CREACAMPO("OFABRIC", "REPARA", "DESCTO1", 7, 0) 'DESCUENTO 1 DEL PRESUPUESTO
'   Call CREACAMPO("OFABRIC", "REPARA", "DESCTO2", 7, 0) 'DESCUENTO 2 DEL PRESUPUESTO X MEDIO DE COBRANZA
'   Call CREACAMPO("OFABRIC", "REPARA", "MED_COBRANZA", 3, 0) 'MEDIO DE COBRANZA TABLA MEDITEA
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    
   NUREP& = XVALO(Label31)
   CONDI$ = "NUMREPA = " & NUREP&
   SQLX$ = "SELECT Cod_Inte,TOT_REPUESTOS,TOT_MANO_OBRA,DESCTO1,DESCTO2,MED_COBRANZA,MoneEmis,TIPOCAM,NumeroSerie"
   SQLX$ = SQLX$ + " FROM REPARA WITH(NOLOCK) WHERE " & CONDI$
   Set rs = READSET(SQLX$)
   With rs
     CI% = XVALO(!cod_inte)
     TOTREPU = XVALO(!TOT_REPUESTOS)
     TOTMAOBRA = XVALO(!TOT_MANO_OBRA)
     DTO1 = XVALO(!DESCTO1)
     DTO2 = XVALO(!DESCTO2)
     MECOBRZA% = XVALO(!MED_COBRANZA)
     MON% = XVALO(!MoneEmis)
     TCAMB = XVALO(!TIPOCAM)
     NSERIE$ = SAFETEXT$(!NUMEROSERIE)

   End With
   rs.Close
   Set rs = Nothing
   If MON% < 1 Then MON% = 1
   If TCAMB <= 0.005 Then TCAMB = 1
'   CI% = XVALO(VALOBADA("REPARA", "Cod_Inte", Condi$, "NOMESS"))
   
   
'   CI% = XVALO(VALOBADA("REPARA", "Cod_Inte", Condi$, "NOMESS"))
   '
   'AGREGO LA INFO AL CARDEFAC
   
    TX1$ = "Reparación Sobre Código : " + CODEXT$(CI%)
    TX2$ = Left$(DESCRIT$(CI%), 64)
    TX3$ = "Según O/Reparación : " & NUREP&
'    TOTREPU = XVALO(VALOBADA("REPARA", "TOT_REPUESTOS", Condi$, "NOMESS"))
'    TOTMAOBRA = XVALO(VALOBADA("REPARA", "TOT_MANO_OBRA", Condi$, "NOMESS"))
    
    'ACA CALCULAR EL TEMA DE LOS DESCUENTOS
    'VER COMO MODIFICAR SI DESEA MODIFICAR LA MODALIDAD DE PAGO
    'O SI PRESENTAR UNA GRILLA CON LOS VALORES OPCIONALES SEGUN MODO DE PAGO.

    TOTREPA# = TOTREPU + TOTMAOBRA
    TOTREPA_CONDI_SELEC# = CALCULA_IMPO_REPA#(TOTREPA#, DTO1, DTO2) 'CONSIDERA TODOS LOS DESCUENTOS Y/O RECARGO DE LA REPARACION
    TOTREPA_SOLO_C_DESC_REPARA# = CALCULA_IMPO_REPA#(TOTREPA#, DTO1, 0) ' ACA SOLO CONSIDERA EL DESCUENTO DE LA REAPARACION NO EL DEL MEDIO DE COBRANZA
    '
    i& = 0
    Campos$ = "Código/f6;Medio Cobranza/A32;Porcentaje/F10.2;Importe/f12.2"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "Confirme o Modfique el Medio de Cobranza"
    SQLX$ = "SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE LISTA  = 'MEDIOS COBRANZA-MDT' ORDER BY CAMPO1"
    
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.MSF2.Rows = i& + 1
        FRMZELECHO.MSF2.TextMatrix(i&, 1) = XVALO(!CAMPO1)
        FRMZELECHO.MSF2.TextMatrix(i&, 2) = SAFETEXT$(!CAMPO2)
        DESCUENTOTABLA = XVALO(!CAMPO3)
        
        If MECOBRZA% = XVALO(!CAMPO1) Then
           IMPO# = TOTREPA_CONDI_SELEC#
           FRMZELECHO.MSF2.TextMatrix(i&, 2) = FRMZELECHO.MSF2.TextMatrix(i&, 2) & " **Seleccionada**"
           FRMZELECHO.MSF2.TextMatrix(i&, 3) = SAFETEXT$(DTO2)
           Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.MSF2, vbGreen, i&)
        Else
           FRMZELECHO.MSF2.TextMatrix(i&, 3) = SAFETEXT$(DESCUENTOTABLA)
           IMPO# = CALCULA_IMPO_REPA#(TOTREPA_SOLO_C_DESC_REPARA#, 0, DESCUENTOTABLA)
        End If
        '
        FRMZELECHO.MSF2.TextMatrix(i&, 4) = TRIM$(FORMATNUM$(IMPO#, "F12.2"))
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    If i& = 0 Then GoTo 60 'SI NO TIENE DATOS EN TABLA DE MEDIOS DE COBRANZA (CASO EYM)
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1

    TOTREPA# = XVALO(RESP_ZELE_VECT(4)) ' IMPORTE SELECCIONADO SEGUN CONDICION DE MEDIO DE COBRANZA
    CONDICION_SELECCIONADA_DE_GRILLA$ = RESP_ZELE_VECT(2)
60  If TOTREPA# < 0.005 And InStr(EPAC$, "EYM") < 1 Then Exit Sub
    '
'    MONEMI% = MON%
'    MONEMIS = DEMONECO$(MONEMI%)
'    MONEMIS.ForeColor = &H0&: If MONEMI% > 1 Then MONEMIS.ForeColor = &H404080
'    TICAMBIS = FORMATNUM$(TCAMB, "F9.4")
    If InStr(EPAC$, "EYM") > 0 Then
       'CAMBIO REALIZADO EL 20/03/20 MONEDA $  POR DEFECTO CALCULADO CON LA COTIZACION DE LA REPARACION
       ' **********OJO SI CAMBIA A DOLARES SE ACTUALIZA A LA COTIZACION ACTUAL **********************
       MONEMI% = 1
       TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
       MONEMIS = DEMONECO$(MONEMI%)
       MONEMIS.ForeColor = &H0&: If MONEMI% > 1 Then MONEMIS.ForeColor = &H404080
       TICAMBIS = FORMATNUM$(1, "F9.4")
       MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
       TICAMBIS.Visible = False
       Label24.Visible = False
       CONDI$ = "NUMREPA = " & NUREP&
       COTIZACIONREPA# = XVALO(VALOBADA("COTIEYM", "COTIZA_DOLAR", CONDI$, "NOMESS"))
'       DESC$ = TRIM$(VALOBADA("COTIEYM", "PORDESCU", CONDI$, "NOMESS"))
       TTT1# = 0
'       If DESC$ <> "" Then
'         PODES# = COEFDESCU(DESC$)
'         DESCUX# = (XVALO(TOTMAOBRA) * PODES# / 100) 'SOLO SOBRE LA REPARACION
'       End If
       TTT2# = FORMATNUM$(XVALO(TOTREPA#) - DESCUX#, "F12.2")
       ''''''''''''''''''''''
       '  *** FIN SETEO DE MONEDA
    End If
    
'    MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
'    TICAMBIS.Visible = True
'    Label24.Visible = True
    ' *** FIN SETEO DE MONEDA

'    If MON% <> MONEMI% Then 'SI EL CAMBIO ES DIFERENTE AL DE PANTALLA
'80     MONEMIS_DblClick           'REALIZA EL DOBLE CLIC
'       If MON% <> MONEMI% Then GoTo 80 ' TANTAS VECES HASTA QUE COINCIDA.
'    End If
    'SI MONEMI% ES MAYOR A 1 MANTIENE TCAMB = 1 PARA QUE NO LO CALCULE, POR QUE YA VIENE EL VALOR EN OTRA MONEDA.
    
    
    If MONEMI% > 1 Then TCAMB = 1
    Call SETULTREG&("!CARDEFAC", 0)
    For i& = 1 To 3
    TX$ = REGBLAN$("!CARDEFAC")
        Call REPLA(TX$, MKS$(1), 1, 4)
        Call REPLA(TX$, TX1$, 7, 64)
        If i& = 1 Then
            Call REPLA(TX$, MKS(TOTREPA# * TCAMB), 429, 4)
            Call REPLA(TX$, MKD(TOTREPA# * TCAMB), 433, 8)
        End If
        If i& = 2 Then Call REPLA(TX$, TX2$, 7, 64)
        If InStr(EPAC$, "EYM") > 0 Then TX3$ = "Reparación : " & NUREP& & " (Serie N° : " & NSERIE$ & ")"
        If i& = 3 Then Call REPLA(TX$, TX3$, 7, 64)
'        If i& = 4 Then Call REPLA(TX$, CONDICION_SELECCIONADA_DE_GRILLA$, 7, 64)
        Call REGAPP("!CARDEFAC", TX$)
    Next i&
    
    If InStr(EPAC$, "EYM") > 0 Then
        'esto es para los accesorios
        Campos$ = "Cantidad,Cod_Inte,CODIGO,DESCRIPCION,IdenLote,PrUn_Neto"
        SQLX$ = "SELECT " & Campos$
        SQLX$ = SQLX$ + " FROM ACCESORIOSCOTIZ WITH(NOLOCK) WHERE NUMREPA = " & NUREP&
        Set rs1 = READSET(SQLX$)
        With rs1
          Do While Not .EOF
             CI1% = XVALO(!cod_inte)
             Codigo$ = SAFETEXT$(!Codigo)
             CANTI = XVALO(!Cantidad)
             LOTE$ = SAFETEXT$(!idenlote)
             CI1% = XVALO(!cod_inte)
             PRECIOX# = XVALO(!prun_neto) * TCAMB
             IMPORTE# = PRECIOX * CANTI
             Descripcion$ = SAFETEXT$(!Descripcion)
             CODDESCLOTE$ = TRIM$(Left$(Codigo$, 16)) + "-" + TRIM$(Left$(Descripcion$, 30)) + "-" + TRIM$(Left$(LOTE$, 16))
             i& = i& + 1
             
             TX$ = REGBLAN$("!CARDEFAC")
             Call REPLA(TX$, MKS$(CANTI), 1, 4)
             Call REPLA(TX$, CODDESCLOTE$, 7, 64)
             Call REPLA(TX$, MKS(PRECIOX#), 429, 4)
             Call REPLA(TX$, MKD(IMPORTE#), 433, 8)
             Call REGAPP("!CARDEFAC", TX$)
            .MoveNext
          Loop
        End With
        rs1.Close
        Set rs1 = Nothing
        PORDESCU = DTO1 'ASIGNA EL DESCUENTO QUE VIENE DEL PRESUPUEESTO DE LA REPARACION
    End If
    Call SETULTREG("!CARDEFAC", i& - 1) ' marca fin de cardefac
    
    

End Sub

Function EXCEPCION_EYM_ACCESORIOS%(SALIDAVERIFICADA%)
    EXCEPCION_EYM_ACCESORIOS = 1
    If SALIDAVERIFICADA% = 1 Then
        FIN& = ULTREG&("!CARDEFAC")
        HAYEQUIPO% = 0: ACCESORIO_NO_REGISTRADO% = 0
        For UI& = 1 To FIN&
               XX$ = REGLEIDO$("!CARDEFAC", UI&)
               CI1% = CVI(Mid$(XX$, 5, 2))
               If ESEQUIPO_EYM%(CI1%) > 0 Then
                  HAYEQUIPO% = 1
                  Exit For
               End If
'               If ESACCESORIO_EYM%(CI1%) < 1 Then
'                  ACCESORIO_NO_REGISTRADO% = 1
'                  Exit For
'               End If
               
         Next UI&
     End If
     
     If HAYEQUIPO% > 0 Then
        EXCEPCION_EYM_ACCESORIOS% = 0
        Exit Function
     End If
     '
'     If ACCESORIO_NO_REGISTRADO% < 1 Then
'        EXCEPCION_EYM_ACCESORIOS = 1
'        Exit Function
'     End If
     '
End Function

Function DATOSIMPO$(CI1%, LOTE_SERIE$)

    'TRAE LOS DATOS DE IMPORTACION SEGUN EL LOTE
    Dim obj As New Art_Importados
    Call obj.VeriImportado(CODEXT$(CI1%))
    DATOSIMPO$ = ""
    
    'SI ES IMPORTADO VA A ENTRAR SI NO SALE
    If obj.EsImportado < 1 Then GoTo 99
    
    TRZAX% = ESTRAZABLE%(CI1%)
    '
    If TRZAX% = 1 Then
       Call obj.READ_DATOS_IMPO(0, LOTE_SERIE$)
       DATOSIMPO$ = " Despacho N°: " & obj.Despacho & " Año: " & obj.AÑO & " " & obj.Aduana
    ElseIf TRZAX% = 2 Then 'PARA EQUIPOS
       If LOTE_SERIE$ <> "S/N" Then
         NORFA$ = OrdenFabricacionNumSerieTanumse$(CI1%, LOTE_SERIE$) 'TRAE LA OF SEGUN N.SERIE
         LOTE$ = OrigenSegunOF$(NORFA$)  'TRAE EL LOTE GUARDADO EN EL CAMPO ORIGEN DE LA OF
         'OJO ACA OBTIENE EL CAMPO NUMERICO DE LA RECEPCION NO DE TEXTO, O SEA QUE LOS DATOS DEBEN SER GENERALES PARA TODA LA OF.
         Call obj.READ_DATOS_IMPO(XVALO(LOTE$), "")    'TRAE LOS DATOS SEGUN LOTE
         DATOSIMPO$ = " Despacho N°: " & obj.Despacho & " Año: " & obj.AÑO & " " & obj.Aduana
       End If
    End If
    
99  Set obj = Nothing
End Function

Sub GRABAFACTURAPROFORMA(NPROF&, NUFACTU$)
    If NPROF& > 0 Then
        SQLX$ = "SELECT * FROM PROFORMTL WHERE NUMECOMP = " & NPROF&
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With rs
          !Status = 1
          !FacturaReal = NUFACTU$
          !FechFactura = Date
          .Update
        End With
    End If
    rs.Close
    Set rs = Nothing
    
End Sub


'
Sub GRABACONCAE(DOCUM$)
    '
    FACELECOMUN% = XVALO(Control("FACTURA", "PUVENELE"))
    FACELEBIENCAP% = XVALO(Control("FACTURA", "PUVENCAP"))
    FACELEEXPO% = XVALO(Control("FACTURA", "PUVEFAEX"))
    PVTMODOCREDITO% = XVALO(Control("FACTURA", "PUVECRED"))
    PTOVTA% = XVALO(Label12.Caption)
    '
    Select Case PTOVTA%
        Case FACELECOMUN%: ARCHIVOLECTURA$ = "consulta_cae_v1.txt"
        Case FACELEBIENCAP%: ARCHIVOLECTURA$ = "consulta_cae_WSBFE.txt"
        Case FACELEEXPO%: ARCHIVOLECTURA$ = "consulta_cae_WSFEX.txt"
        Case PVTMODOCREDITO%: ARCHIVOLECTURA$ = "consulta_cae_v1.txt"
        Case Else: Exit Sub
    End Select
    '
    TX$ = ""
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    NARCHI2% = FILEOPEN%(CARPETOK$ + ARCHIVOLECTURA$, 1, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    Do While Not EOF(NARCHI2%) ' RECORRE
       Line Input #NARCHI2%, XX$ ' LEE CADA REGISTRO
       TX$ = TX$ + XX$ + vbCrLf
    Loop
    Close NARCHI2%
    '
    Call ABRECABAN
TXTRAZA$ = TXTRAZA$ & "PUNTO 36 = " & Now & vbCrLf
'    Call ACTUREGISTRO("CAJABANC", "TEXTOCAE", "CODICOM_LAR = '" & DOCUM$ & "' AND TIPOMOVIM='FVEN'", TX$, REGAF&, "NOMESS")
     'GRABA LOS 3 MOVIMIENTOS DE UNA VEZ
'    Call ACTUREGISTRO("CAJABANC", "TEXTOCAE", "CODICOM_LAR = '" & DOCUM$ & "'", TX$, REGAF&, "NOMESS")
     '
     'SE CAMBIA A UN UPDATE DIRECTO POR QUE EN EL ACTUREGISTRO DE ARRIBA GUARDA UNO SOLO X Q' AL SER UN TIPO DE CAMPO 99 VER ACTUREGISTRO
     SQLX$ = "UPDATE CAJABANC SET TEXTOCAE = '" & TX$ & "'"
     SQLX$ = SQLX$ + " WHERE CODICOM_LAR = '" & DOCUM$ & "'"
     FLEXCONN.Execute (SQLX$)
     '
TXTRAZA$ = TXTRAZA$ & "PUNTO 37 = " & Now & vbCrLf
'    Call ACTUREGISTRO("CAJABANC", "TEXTOCAE", "CODICOM_LAR = '" & DOCUM$ & "' AND TIPOMOVIM='VENTA'", TX$, REGAF&, "NOMESS")
TXTRAZA$ = TXTRAZA$ & "PUNTO 38 = " & Now & vbCrLf
'    Call ACTUREGISTRO("CAJABANC", "TEXTOCAE", "CODICOM_LAR = '" & DOCUM$ & "' AND TIPOMOVIM='IVA'", TX$, REGAF&, "NOMESS")
TXTRAZA$ = TXTRAZA$ & "PUNTO 39 = " & Now & vbCrLf
    '
End Sub
'
'
Sub CARDETCRE(OKX%)

    OKX% = 0
    NCLI = XVALO(Text1)
    Call COPYSTRU("SELCUERE", "SELFCND")
    SQLX$ = " TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + " AND (OPCIMovim = 'FC' OR OPCIMovim = 'ND') "
    SQLX$ = SQLX$ + " AND NuCLIEN = " & NCLI & " "
    SQLX$ = SQLX$ + " ORDER BY FechConta DESC "
    If chkModoCredito.Value = 1 Then
        SQLX$ = "nuclien=" & Text1 & " and tipomovim='fven' and opcimovim='fc' and modoopcimovim='cr' and numecomp>0"
    End If
    Call CARGALISEL("SELFCND", "CAJABANC", "NUMECOMP/I6;CODICOM_LAR/A20;FechConta/D10", SQLX$, "DISTINCT")
    '
    Call TRALOCAL("SELFCND", "")
    '
    If ULTREG("SELFCND") < 1 Then Exit Sub ' VALIDA QUE TENGA INFORMACION
    Call SELECHO("SELFCND") ' PRESENTO PARA SELECCIONAR FACTUR
    If VALACT1$("SELFCND") = "" Then Exit Sub ' SI NO ELIGIO REGISTRO
    '
    NFACTUR$ = Left$(VALACT2$("SELFCND"), 18) ' FACTURA SELECCIONADA
    NUREP& = XVALO(VALOBADA("REPARA", "NUMREPA", "FACTURA  = '" & NFACTUR$ & " '", "NOMESS"))
    Label31 = TRIM$(Str(NUREP&))
    '
    ' SE GUARDA EL NUMERO DE FACTURA PARA LA APLICACION AUTOMATICA DE LA N/CREDITO Y AGENTE DE PERCEPCION DE IIBB CABA.
    ' AGENTE DE PERCEPCION: EL APLICATIVO DEL GOBIERNO DE LA CIUDAD PERMITA DESCONTAR LA PERCEPCION SI Y SOLO SI
    ' COINCIDE AL CENTAVO LA FACTURA Y NOTA DE CREDITO EN LO QUE RESPECTA AL IMPORTE NETO Y EL IMPORTE DE LA RETENCION.
    ' CON EL NUMERO DE COMPROBANTE SE OBTIENEN ESOS DOS VALORES DE LA FACTURA Y SE COMPARAN CON LA N/CREDITO.
    AnulaFacturaComprobante$ = TRIM$(NFACTUR$)

    'CONDI$ = "DOSECLAR = '" & NFACTUR$ & "' AND COD_INTE > 0 AND NUME_CLIE = " & NCLI
    CLIRE = XVALO(VALOBADA("CAJABANC", "CLI_REAL", "CODICOM_LAR = '" & NFACTUR$ & " ' AND NUCLIEN=" & NCLI, "NOMESS"))
    If CLIRE <> NCLI Then NCLI = CLIRE
    CONDI$ = "DOSECLAR = '" & NFACTUR$ & "' AND COD_INTE > 0 AND NUME_CLIE = " & NCLI
    CANTCI& = CantRegistros("MOVISTO", CONDI$) 'CANTIDAD DE REGISTROS CON CODIGO INTERNO
    If CANTCI& < 1 Then
       Label6.Caption = "(T.Libre)": MODODOC% = 1
    Else
        Label6.Caption = "(Mat.Codif)": MODODOC% = 0
    End If
    INCOTERMS$ = TRIM$(SAFETEXT$(VALOBADA("CAJABANC", "INCOTERMS", "CODICOM_LAR = '" & NFACTUR$ & " ' AND NUCLIEN=" & NCLI, "NOMESS")))
    Text18 = INCOTERMS$
    '
    ' VALIDACIONES INGRESOS BRUTOS C.A.B.A.
    '
    If EXISTE%(LUDAT$ + "PAIBCABA.TXT") > 0 Then
        ' VALIDA QUE LA FACTURA A ANULAR CORRESPONDA AL MISMO MES Y AÑO DE LA NOTA DE CREDITO
     TiMovi$ = "IBRU"
     OpMovi$ = "IB-C"
     VaMovi$ = "PERC"
        FechaFactura% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", "CODICOM_LAR='" & AnulaFacturaComprobante$ & "' AND NUCLIEN=" & CStr(NCLI) & " AND TIPOMOVIM='IBRU' AND OPCIMOVIM='IB-C' AND VARIMOVIM='PERC'"))
        If FechaFactura% > 0 Then
            MesFactura$ = Mid$(FECHATEX$(FechaFactura%), 4) ' CONTEMPLA MES Y AÑO
            FechaNotaCredito% = FECHANUM(Text9)
            MesNotaCredito$ = Mid$(FECHATEX$(FechaNotaCredito%), 4) ' CONTEMPLA MES Y AÑO
            
            If MesFactura$ <> MesNotaCredito$ Then
                Call COMUNI("IMPOSIBLE CONTINUAR.\ \La Nota de Crédito No Puede Cancelar una Factura\Emitida en Otro Mes, cuando el Cliente\" _
                & "se Encuentra Incluido en el Padrón de IIBB C.A.B.A.")
                Exit Sub
            End If
            
            ' ADVIERTE AL USUARIO QUE LA NOTA DE CREDITO SOBRE UN CLIENTE QUE ESTA INLCUIDO EN EL PADRON DE IIBB CABA
            ' DEBE SER POR EL TOTAL DE LA FACTURA.
            TASAX# = ALIPERBRU_CABA#(NCLI, MINIPER#, VARIDOC%)
            If TASAX# > 0.005 Then
                Call COMUNI("IMPORTANTE!!\ \Tener en Cuenta que al estar el Cliente '" & TRIM$(Text2) & "'\Incluido en el Padrón de IIBB C.A.B.A.,\Sólo Puede Emitir una Nota de Crédito\por el TOTAL de la FACTURA a Anular.")
            End If
        End If
    End If

    If MODODOC% = 1 Then ' texto libre
        '
        CONDI$ = "NUMFACTURA = '" & NFACTUR$ & "' AND NUMECLI = " & NCLI
        CANTTRAZ& = CantRegistros("TANUMSE", CONDI$) '
        If CANTTRAZ& > 0 Then 'VALIDA QUE NO HAYA ART.TRAZABLES EN TANUMSE
            Call MENSERR(24, "Hay Equipos Trazables (Nivel 2)\Que No se Han Reingresado\Deberá Primero Reingresar Los Mismos ")
            OKX% = 0
            Exit Sub
        End If
        'TRAE EL IMPORTE DE LA FACTURA Y LO PRESENTA EN CARDETFAC
        CONDI$ = "TIPOMOVIM = 'VENTA' AND CODICOM_LAR = '" & NFACTUR$ & "'AND NuCLIEN = " & NCLI & ""
        TOT# = XVALO(VALOBADA("CAJABANC", "IMPO_DOLA", CONDI$, "NOMESS"))
        IMPOR = TOT# ' PARA REPETIR EL IMPORTE
        '
        TX$ = REGBLAN$("!CARDEFAC")
        NFACT$ = "Corresponde A Factura: " + NFACTUR$
        '
        Call SETULTREG("!CARDEFAC", 0)
        Call REPLA(TX$, MKS$(1), 1, 4)
        Call REPLA(TX$, NFACT$, 7, 64)
        Call REPLA(TX$, MKS(IMPOR), 429, 4)
        Call REPLA(TX$, MKD(TOT#), 433, 8)
        Call REGAPP("!CARDEFAC", TX$)
        Call REPLA(TX$, MKS$(PORCEIVA(CI%)), 441, 4)
        If NUREP& > 0 Then
           TX$ = REGBLAN$("!CARDEFAC")
           TXTX$ = "Según Reparación: " & NUREP&
           Call REPLA(TX$, MKS$(1), 1, 4)
           Call REPLA(TX$, TXTX$, 7, 64)
           Call REGAPP("!CARDEFAC", TX$)
           Call SETULTREG("!CARDEFAC", 2)
        End If
        '
        ' *** SETEA EN PANTALLA LA MONEDA DE LA FACTURA
        CondicionSql$ = "codicom_lar='" & NFACTUR$ & "' and nuclien=" & CStr(NCLI)
        TIPOMONE% = XVALO(VALOBADA("cajabanc", "mone_emis", CondicionSql$, "NOMESS")): If TIPOMONE% < 1 Then TIPOMONE% = 1
        VALOCAMBIO# = XVALO(VALOBADA("cajabanc", "tipo_cam", CondicionSql$, "NOMESS")): If VALOCAMBIO# < 0.00005 Then VALOCAMBIO# = 1
        
        MONEMI% = TIPOMONE%
        MONEMIS = DEMONECO$(MONEMI%)
        MONEMIS.ForeColor = &H0&: If MONEMI% > 1 Then MONEMIS.ForeColor = &H404080
        TICAMBIS = FORMATNUM$(VALOCAMBIO#, "F9.4")
        MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
        TICAMBIS.Visible = True
        Label24.Visible = True
        ' *** FIN SETEO DE MONEDA
        '
        OKX% = 1
        ' Advertencia por que es posible que el total de la factura difiera segun tasa de iva
        Call COMUNI("Verifique Que la Tasa de Iva\Aplicada a Esta Factura Sea la Correcta  ")
        
    ElseIf MODODOC% = 0 Then ' material codificado
        '
        '
        'OJO CONDICION DE CONSULTA SE UTILIZA LUEGO PARA RECORRER MOVISTO -
        CONDI$ = "DOSECLAR  = '" & NFACTUR$ & "'"
        CONDI$ = CONDI$ + " AND NUME_CLIE = " & NCLI & ""
        CONDI$ = CONDI$ + " AND COD_INTE > 0 "
        
        'AGREGADO PARA LEER EL TIPO DE CAMBIO Y EL VALOR DEL TIPO DE CAMBIO
        TIPOMONE% = XVALO(VALOBADA("MOVISTO", "MoneVal", CONDI$, "NOMESS")): If TIPOMONE% < 1 Then TIPOMONE% = 1
        VALOCAMBIO# = XVALO(VALOBADA("MOVISTO", "TIPOCAM", CONDI$, "NOMESS")): If VALOCAMBIO# < 0.00005 Then VALOCAMBIO# = 1
        'SE LEE EL TIPO DE CAMBIO DE PANTALLA PARA EL CASO QUE NO SEA $ POR EJEMPLO QUE SE HAYA CAMBIADO A U$S
        'TICAMPANTALLA = TICAMONE(MONEMI%): If TICAMPANTALLA < 0.00005 Then TICAMPANTALLA = 1
        'VALOCAMBIO# = VALOCAMBIO# / TICAMPANTALLA 'SE CALCULA COEFICIENTE PARA CALCULAR A PARTIR DEL QUE ESTA EN PANTALLA
        '
        ' *** SETEA EN PANTALLA LA MONEDA DE LA FACTURA
        MONEMI% = TIPOMONE%
        MONEMIS = DEMONECO$(MONEMI%)
        MONEMIS.ForeColor = &H0&: If MONEMI% > 1 Then MONEMIS.ForeColor = &H404080
        TICAMBIS = FORMATNUM$(VALOCAMBIO#, "F9.4")
        MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
        TICAMBIS.Visible = True
        Label24.Visible = True
        ' *** FIN SETEO DE MONEDA
        '
        CONDI$ = CONDI$ + " AND CANTSALID > 0 AND CODIMOVI = 'RT'"
        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE " & CONDI$
        Set FACTTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        JX& = 0
        With FACTTMP
          Do While Not .EOF ' CARGA TODOS LOS ITEMS DEL MOVISTO RELACIONADO CON FACTURA
           CI% = XVALO(!cod_inte)
           NUMORDIT% = !NuOrItem

           NUSERIE$ = TRIM$(SAFETEXT$(!idenlote))
           ' If ESTRAZABLE%(CI%) < 2 Then NUSERIE$ = "S/N"
           JX& = JX& + 1
           CAN# = XVALO(!CANTSALID)

           'PRECIO# = XVALO(!ValoUnit) * VALOCAMBIO# 'CALCULO EL IMPORTE POR EL TIPO DE CAMBIO
           PRECIO# = XVALO(!VALOUNIT)
           
           TX$ = REGBLAN$("!CARDEFAC")
           Call REPLA(TX$, MKS$(CAN#), 1, 4)
           Call REPLA(TX$, MKI$(CI%), 5, 2)
           Call REPLA(TX$, DESCRIT$(CI%), 7, 64)
''           Call REPLA(TX$, MKI$(NUMORDIT%), 117, 2) 'MODIFICACION N.CREDITO PARA QUE EL REMITO QUEDE COMO EN CONSIGNACION
           Call REPLA(TX$, MKS$(PRECIO#), 421, 4)
           Call REPLA(TX$, MKS$(0), 425, 4)       ' DESCUENTO
           Call REPLA(TX$, MKS$(PRECIO#), 429, 4)
           IMPO# = PRECIO# * CAN#
           Call REPLA(TX$, MKD$(IMPO#), 433, 8)
           Call REPLA(TX$, MKS$(PORCEIVA(CI%)), 441, 4)
''           Call REPLA(TX$, "N", 461, 1) 'MODIFICACION N.CREDITO PARA QUE EL REMITO QUEDE COMO EN CONSIGNACION
           Call REPLA(TX$, NUSERIE$, 139, 64)
           Call GRAREG("!CARDEFAC", TX$, JX&)
           '
           .MoveNext
          Loop
        End With
        FACTTMP.Close
        Set FACTTMP = Nothing
        '
''         'MODIFICACION N.CREDITO PARA QUE EL REMITO QUEDE COMO EN CONSIGNACION
''        LETX$ = Mid$(TX$, 461, 1)
''        If LETX$ = "N" Then
''           TX$ = REGBLAN$("!CARDEFAC")
''           Call REPLA(TX$, MKS$(1), 1, 4)
''           Call REPLA(TX$, NFACTUR$, 7, 64)
''           Call REGAPP("!CARDEFAC", TX$)
''           JX& = JX& + 1
''        End If
''        'FIN MODIFICACION N.CREDITO PARA QUE EL REMITO QUEDE COMO EN CONSIGNACION
        
        ' RECORRE EL TANUMSE Y LE APLICA EL NUMERO DE SERIE
        Call CREATABLA_EQUIPOS ' SE AGREGO PARA QUE CREA LA TABLA
        SQLX$ = "SELECT * FROM TANUMSE WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE NUMFACTURA = '" & NFACTUR$ & "'"
        SQLX$ = SQLX$ + " AND NUMECLI = " & NCLI
        Dim TANUTMP As ADODB.Recordset
        Set TANUTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With TANUTMP
          Do While Not .EOF
            CI_TANUMSE% = XVALO(!cod_inte)
            NUMESER$ = SAFETEXT$(!NUMEROSERIE)
            For i& = 1 To ULTREG&("!CARDEFAC")
              XX$ = REGLEIDO$("!CARDEFAC", i&)
              CI% = CVI(Mid$(XX$, 5, 2))
              If CI% = CI_TANUMSE% Then
                 NUMSERCFAC$ = TRIM$(Mid$(XX$, 139, 64))
                 If NUMSERFAC$ = NUMESER$ Then GoTo 110    ' YA LO ENCONTRO EN EL CARDEFAC
                 If NUMSERCFAC$ = "" Then
                    Call REPLA(XX$, NUMESER$, 139, 64)
                    Call GRAREG("!CARDEFAC", XX$, i&)
                    GoTo 110
                 End If
              End If
            Next i&
110       .MoveNext
          Loop
        End With
        TANUTMP.Close
        Set TANUTMP = Nothing
        '
'        ' VALIDA QUE LOS TRAZABLES 2 QUE NO TIENEN NUMEROS DE SERIE LOS QUITA
'        J& = 0
'        For I& = 1 To ULTREG&("!CARDEFAC")
'           XX$ = REGLEIDO$("!CARDEFAC", I&)
'           CI% = CVI(Mid$(XX$, 5, 2))
'           NUMESER$ = TRIM$(Mid$(XX$, 139, 64))
'           If ESTRAZABLE%(CI%) > 1 And NUMESER$ = "" Then GoTo 120
'           J& = J& + 1
'           Call GRAREG("!CARDEFAC", XX$, J&)
'120     Next I&
        'VERIFICO SI NO ESTA VACIO EL CARDEFAC
        '
        Call SETULTREG("!CARDEFAC", JX&) ' marca fin de cardefac

        If JX& < 1 Then ' si no encontro ningun ariticulo
           TTXX$ = "No Hay Materiales Para Descontar de Esta Factura.\"
           TTXX$ = TTXX$ + "Posible Cliente con Grupo Divisional.\"
           TTXX$ = TTXX$ + "Para Anular la Factura Debe Realizar una Nota de Crédito de Texto Libre\"
           TTXX$ = TTXX$ + "y Re-ingresar el Material Mediante la Opción de Devolución de Materiales\"
           TTXX$ = TTXX$ + "del Modulo de Remitos o a partir de un Ajuste de Stock"
           Call COMUNI(TTXX$)
           OKX% = 0
           Exit Sub
        Else
           If TIPOMONE% <> MONEMI% Then 'SI EL CAMBIO ES DIFERENTE AL DE PANTALLA
50           MONEMIS_DblClick           'REALIZA EL DOBLE CLIC
             If TIPOMONE% <> MONEMI% Then GoTo 50 ' TANTAS VECES HASTA QUE COINCIDA.
           End If
           OKX% = 1
        End If
        
    End If
    

'''    CODIGO ANTERIOR COMENTADO DE RESPALDO

'''    OKX% = 0
'''    NCLI = XVALO(Text1)
'''    Call COPYSTRU("SELCUERE", "SELFCND")
'''    SQLX$ = " TipoMovim = 'FVEN' "
'''    SQLX$ = SQLX$ + " AND (OPCIMovim = 'FC' OR OPCIMovim = 'ND') "
'''    SQLX$ = SQLX$ + " AND NuCLIEN = " & NCLI & " "
'''    SQLX$ = SQLX$ + " ORDER BY FechConta DESC "
'''    Call CARGALISEL("SELFCND", "CAJABANC", "NUMECOMP/I6;CODICOM_LAR/A20;FechConta/D10", SQLX$, "DISTINCT")
'''    '
'''    Call TRALOCAL("SELFCND", "")
'''    '
'''    If ULTREG("SELFCND") < 1 Then Exit Sub ' VALIDA QUE TENGA INFORMACION
'''    Call SELECHO("SELFCND") ' PRESENTO PARA SELECCIONAR FACTUR
'''    If VALACT1$("SELFCND") = "" Then Exit Sub ' SI NO ELIGIO REGISTRO
'''    '
'''    NFACTUR$ = Left$(VALACT2$("SELFCND"), 18) ' FACTURA SELECCIONADA
'''    '
'''    CONDI$ = "DOSECLAR = '" & NFACTUR$ & "' AND COD_INTE > 0"
'''    CANTCI& = CantRegistros("MOVISTO", CONDI$) 'CANTIDAD DE REGISTROS CON CODIGO INTERNO
'''    If CANTCI& < 1 Then
'''       Label6.Caption = "(T.Libre)": MODODOC% = 1
'''    Else
'''        Label6.Caption = "(Mat.Codif)": MODODOC% = 0
'''    End If
'''
'''    If MODODOC% = 1 Then ' texto libre
'''        '
'''        CONDI$ = "NUMFACTURA = '" & NFACTUR$ & "'"
'''        CANTTRAZ& = CantRegistros("TANUMSE", CONDI$) '
'''        If CANTTRAZ& > 0 Then 'VALIDA QUE NO HAYA ART.TRAZABLES EN TANUMSE
'''            Call MENSERR(24, "Hay Equipos Trazables (Nivel 2)\Que No se Han Reingresado\Deberá Primero Reingresar Los Mismos ")
'''            OKX% = 0
'''            Exit Sub
'''        End If
'''        'TRAE EL IMPORTE DE LA FACTURA Y LO PRESENTA EN CARDETFAC
'''        CONDI$ = "TIPOMOVIM = 'VENTA' AND CODICOM_LAR = '" & NFACTUR$ & "'AND NuCLIEN = " & NCLI & ""
'''        TOT# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", CONDI$, "NOMESS"))
'''        IMPOR = TOT# ' PARA REPETIR EL IMPORTE
'''        '
'''        TX$ = REGBLAN$("!CARDEFAC")
'''        NFACT$ = "Corresponde A Factura: " + NFACTUR$
'''        '
'''        Call SETULTREG("!CARDEFAC", 0)
'''        Call REPLA(TX$, MKS$(1), 1, 4)
'''        Call REPLA(TX$, NFACT$, 7, 64)
'''        Call REPLA(TX$, MKS(IMPOR), 429, 4)
'''        Call REPLA(TX$, MKD(TOT#), 433, 8)
'''        Call GRAREG("!CARDEFAC", TX$, 1)
'''        Call REPLA(TX$, MKS$(PORCEIVA(CI%)), 441, 4)
'''        OKX% = 1
'''        ' Advertencia por que es posible que el total de la factura difiera segun tasa de iva
'''        Call COMUNI("Verifique Que la Tasa de Iva \Aplicada a Esta Factura Sea la Correcta  ")
'''
'''          '
'''    ElseIf MODODOC% = 0 Then ' material codificado
'''        '
'''        CONDI$ = "NUMFACTURA = '" & NFACTUR$ & "'"
'''        Call COPYSTRU("INDINUSE", "NCINDINU")
'''
'''        Call SETULTREG("!CARDEFAC", 0)
'''        Call SETULTREG("!NCINDINU", 0)
'''        '
'''        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
'''        SQLX$ = SQLX$ + " WHERE DOSECLAR  = '" & NFACTUR$ & "'"
'''        SQLX$ = SQLX$ + " AND NUME_CLIE = " & NCLI & ""
'''        Set FACTTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'''        JX& = 0
'''        With FACTTMP
'''          Do While Not .EOF
'''           CI% = XVALO(!COD_INTE)
'''           CAN = XVALO(!cantsalid)
'''           'PRECIO = XVALO(!ValoUnit)
'''
'''           If CI% < 1 Then GoTo 100 ' PARA QUE NO MUESTRE ARTICULOS NO CODIFICADOS
'''           '
'''           If ESTRAZABLE%(CI%) > 1 Then GoTo 100 'Valida si estrazable
'''           JX& = JX& + 1
'''
''''           NUSERIE$ = VALOBADA("TANUMSE", "NUMEROSERIE", CONDI$, "NOMESS")
'''           DESCRI$ = DESCRIT$(CI%)
'''           Call REPLA(XX$, "S/N", 1, 16)
'''           Call REPLA(XX$, CODEXT$(CI%), 18, 18)
'''           Call REPLA(XX$, "Cant.= " & CAN, 37, 10)
'''           Call GRAREG("!NCINDINU", XX$, JX&) '*****CARGAR EL FORMULARIO CON LIST DE SELECCION
'''100        .MoveNext
'''          Loop
'''        End With
'''        FACTTMP.Close
'''        Set FACTTMP = Nothing
'''        '
'''        Call CREATABLA_EQUIPOS
'''        SQLX$ = "SELECT * FROM TANUMSE "
'''        SQLX$ = SQLX$ + " WHERE NUMFACTURA = '" & NFACTUR$ & "'"
'''        SQLX$ = SQLX$ + " AND NUMECLI = " & NCLI & ""
'''        Dim TANUTMP As ADODB.Recordset
'''        Set TANUTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'''        With TANUTMP
'''            '
'''            Do While Not .EOF
'''             CI% = XVALO(!COD_INTE)
'''             NUSERIE$ = SAFETEXT$(!NUMEROSERIE)
'''             JX& = JX& + 1
'''             Call REPLA(XX$, NUSERIE$, 1, 16)
'''             Call REPLA(XX$, CODEXT$(CI%), 18, 18)
'''             Call REPLA(XX$, "Cant.= 1", 37, 10)
'''             Call GRAREG("!NCINDINU", XX$, JX&) '*****CARGAR EL FORMULARIO CON LIST DE SELECCION
'''           '
'''110          .MoveNext
'''          Loop
'''        End With
'''        TANUTMP.Close
'''        Set TANUTMP = Nothing
'''        '
'''        Call SETULTREG("!NCINDINU", JX&)
'''        Call FRESHECHO("!NCINDINU")
'''        '
'''        If ULTREG&("!NCINDINU") < 1 Then
'''           Call COMUNI("No Hay Equipos Para Descontar de Esta Factura\Probablemente Se Haya Realizado N.Credito\o Reingeso de Equipos de la Misma")
'''           OKX% = 0
'''           Exit Sub
'''        End If
'''        '
'''        '*****ACA EN VEZ DE PRESENTAR SELECHO PRESENTAR FORMULARIO CON LIST PARA SELECCIONAR ARTICULOS
'''
'''        Call SELECHO("!NCINDINU")
'''        '
'''        If VALACT1("!NCINDINU") = "" Then
'''          OKX% = 0
'''          Exit Sub
'''        End If
'''        '
'''        ' *****LUEGO DE QUE SE ACEPTA LA VENTANA CARGAR EL CARDETFAC CON LA INFORMACION
'''        ' *****DE LOS ARTICULOS SELECCIONADOS EN EL FORMULARIO DE SELECCION
'''
'''        Call SETULTREG("!CARDEFAC", 0)
'''        CAN = TRIM$(Mid$(VALACT2("!NCINDINU"), 26, 10))
'''        CEXT$ = TRIM$(Mid$(VALACT2("!NCINDINU"), 1, 18))
'''        CI% = CODINT%(CEXT$)
'''        NUMESER$ = TRIM$(VALACT1$("!NCINDINU"))
'''        CONDI$ = "DOSECLAR = '" & NFACTUR$ & "'AND COD_INTE = " & CI% & ""
'''        PRECIO = XVALO(VALOBADA("MOVISTO", "VALOUNIT", CONDI$, "NOMESS"))
'''        TX$ = REGBLAN$("!CARDEFAC")
'''        Call REPLA(TX$, MKS$(CAN), 1, 4)
'''        Call REPLA(TX$, MKI$(CI%), 5, 2)
'''        Call REPLA(TX$, DESCRIT$(CI%), 7, 64)
'''        Call REPLA(TX$, MKS$(PRECIO), 421, 4)
'''        Call REPLA(TX$, MKS$(0), 425, 4)       ' DESCUENTO
'''        Call REPLA(TX$, MKS$(PRECIO), 429, 4)
'''        IMPO# = PRECIO * CAN
'''        Call REPLA(TX$, MKD$(IMPO#), 433, 8)
'''        Call REPLA(TX$, MKS$(PORCEIVA(CI%)), 441, 4)
'''        Call REPLA(TX$, NUMESER$, 139, 64)
'''        Call GRAREG("!CARDEFAC", TX$, 1)
'''        Call SETULTREG("!CARDEFAC", 1)
'''        OKX% = 1
'''    End If
'''        '
'''

End Sub


Sub GRABAR_ASIENTO_ITEM_SIN_CARGO(FACTURA$, IMPORTE)

   If MODODOC% > 0 Then Exit Sub
   If Control("ADMIPED", "ITEMSCGO") <> "SI" Then Exit Sub
   'CONTROLES DE CONFIGURACION DE CUENTAS CONTABLES
   If TIFADOC% = 6 Then 'SI ES NOTA DE CREDITO INVIERTO LAS CUENTAS
     If TRIM$(AnulaFacturaComprobante$) = "" Then Exit Sub 'SI ES NOTA DE CREDITO PERO NO ES ANULA FACTURA NO GENERA ASIENTO
     CTACONTRA$ = TRIM$(Control("ASVENTAS", "ASITEMSC")) 'CUENTA  DE ASIENTO DE VENTAS ITEMS SIN CARGO) 'cuenta configurada en setup de reparaciones
     CTAREP$ = TRIM$(Control("ASVENTAS", "CONTASSC")) 'CUENTA DE CONTRAPARTIDA DE ASIENTO DE VENTAS ITEMS SIN CARGO
      
   Else
      CTACONTRA$ = TRIM$(Control("ASVENTAS", "CONTASSC")) 'CUENTA DE CONTRAPARTIDA DE ASIENTO DE VENTAS ITEMS SIN CARGO
      CTAREP$ = TRIM$(Control("ASVENTAS", "ASITEMSC"))  'CUENTA  DE ASIENTO DE VENTAS ITEMS SIN CARGO) 'cuenta configurada en setup de reparaciones
   End If
   '
   IMPORTE = Abs(IMPORTE)
   CTACONTRAPARTIDA% = CUECOINT%(CTACONTRA$)
   CTAREPU% = CUECOINT%(CTAREP$)
   '
   If IMPORTE <= 0 Then Exit Sub
   
   FEHOY% = HOY(HO$)
   
   Call CARDATEJ
   
   NUASI$ = NUPROA1$(HOY(HO$))
   '
   OBS$ = "ASIENTO X ITEMS S/CARGO (" & FACTURA$ & ")"
   
'   On Error GoTo ERROR_GUARDAR
'   FLEXCONN.BeginTrans

   Dim EncaSiGra As ADODB.Recordset
   Set EncaSiGra = New ADODB.Recordset
   strSQL = "Select * from EncaBeAsien where NuIDAsien = '" & NUASI$ & "'"
   EncaSiGra.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With EncaSiGra
         .AddNew   ' YA CONVERTIDO
         !StatAsien = (-1)
          On Error Resume Next
          !CODIEMPR = CodiEmp%
          On Error GoTo 0
          !nuidasien = NUASI$
          If Text4 = "" Then Text4 = " "
          !refeasien = OBS$
             FECHAU% = HOY(HO$)
          !FECHASIEN = CVDAT(FECHAU%)
             PPXX% = InStr(NUASI$, "-")
          !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
          !SucuAsien = UNINEGO%
          !NuDfAsien = 0
          !TIPOASIEN = 11
          !StatAsien = 1  ' activo
          !EditAsien = -1
          !QPasAsien = 3
          !DebeAsien = IMPORTE
          !Observa = OBS$
          !Nume_Usuar = USNBR%
          !Niv_Clasi = 0
          !FeReAsien = Now
          !Origen = FACTURA$
          .Update
   End With
   
   NICLA% = 0
   ORPASE% = 0
   '
   Set DetasiGra = New ADODB.Recordset
   strSQL = "Select * from DetaAsien where NuIDAsien = '" & NUASI$ & "'"
   DetasiGra.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With DetasiGra
     For K2% = 1 To 2
          Select Case K2%
            Case 1
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next ' ASIENTO DE TOTAL CONTRAPARTIDA DE CUENTA CONTABLE DE REPARACION
              !CODIEMPR = CodiEmp%
              On Error GoTo 0
              !nuidasien = NUASI$
              ORPASE% = ORPASE% + 1
              !OrNuPase = ORPASE%
              !RefePase = Left$(DENOCUE$(CTACONTRAPARTIDA%), 24)
              FECHAU% = FEHOY%
              !FECHASIEN = CVDAT(FECHAU%)
              AA = CVDAT(FECHAU%)
              PPXX% = InStr(NUASI$, "-")
              !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
              !SucuAsien = 0
              !NuDfAsien = 0
              !StatPase = 1
              !CUECOEXT = CTACONTRA$
              !CUECOINT = CTACONTRAPARTIDA%
              !idebepase = 0
              !ihabepase = IMPORTE
              !MARCONSI = ""
              !CenCos = ""
              !Niv_Clasi = 0
              !DATRETPER = ""
             .Update
             
            Case 2
              If IMPORTE <= 0 Then GoTo 80 ' NEXT
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next 'ASIENTO DE REPUESTOS
              !CODIEMPR = CodiEmp%
              On Error GoTo 0
              !nuidasien = NUASI$
              ORPASE% = ORPASE% + 1
              !OrNuPase = ORPASE%
              !RefePase = Left$(DENOCUE$(CTAREPU%), 24)
              FECHAU% = FEHOY%
              !FECHASIEN = CVDAT(FECHAU%)
              PPXX% = InStr(NUASI$, "-")
              !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
              !SucuAsien = 0
              !NuDfAsien = 0
              !StatPase = 1
              !CUECOEXT = CTAREP$
              !CUECOINT = CTAREPU%
              !idebepase = IMPORTE
              !ihabepase = 0
              !MARCONSI = ""
              !CenCos = ""
              !Niv_Clasi = 0
             !DATRETPER = ""
             .Update
          End Select
          '
80        Next K2%
          DetasiGra.Close
          Set DetasiGra = Nothing
   End With
   '
   '
'ERROR_GUARDAR:
'    If Err <> 0 Then
'       FLEXCONN.RollbackTrans
'       On Error Resume Next
'       DetasiGra.Close
'       Set DetasiGra = Nothing
'       On Error GoTo 0
'       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
'       NUSERI$ = VALOBADA("REPARA", "NUMEROSERIE", "NUMREPA = " & NUREP&, "NOMESS")
'       CUERPO$ = "No se Generó El Asiento Automático " & vbCrLf
'       CUERPO$ = "Reparación N°:  " & NUREP& & vbCrLf
'       Call ANOTANUSE(NUSERI$, CUERPO$, , "NOMESS")
'       Exit Sub
'    Else
'       FLEXCONN.CommitTrans
'    End If
90
End Sub


Sub GRABAR_ASIENTO_ITEM_IMPORTADO(FACTURA$, REMITO$, IMPORTE)
    If TRIM$(AnulaFacturaComprobante$) = "" And IMPORTE < 0.005 Then Exit Sub  '
   're-venta va al debe y m.prima al haber.
   'CONTROLES DE CONFIGURACION DE CUENTAS CONTABLES
   If IMPORTE < 0.005 Then  'SI ES NOTA DE CREDITO INVIERTO LAS CUENTAS
     If TRIM$(AnulaFacturaComprobante$) = "" Then Exit Sub 'SI ES NOTA DE CREDITO PERO NO ES ANULA FACTURA NO GENERA ASIENTO
     CTACONTRA$ = TRIM$(Control("ASVENTIM", "ASIMAPRI")) 'CUENTA  DE ASIENTO DE VENTAS MATERIAS PRIMAS
     CTAREP$ = TRIM$(Control("ASVENTIM", "CONTIMPO"))  'CUENTA DE CONTRAPARTIDA DE IMPORTADOS
   Else
      CTACONTRA$ = TRIM$(Control("ASVENTIM", "CONTIMPO"))  'CUENTA DE CONTRAPARTIDA DE IMPORTADOS
      CTAREP$ = TRIM$(Control("ASVENTIM", "ASIMAPRI"))  'CUENTA  DE ASIENTO DE VENTAS MATERIAS PRIMAS
   End If
   '
   
   IMPORTE = Abs(IMPORTE)
   CTACONTRAPARTIDA% = CUECOINT%(CTACONTRA$)
   CTAREPU% = CUECOINT%(CTAREP$)
   '
   If IMPORTE <= 0 Then Exit Sub
   
   FEHOY% = HOY(HO$)
   
   Call CARDATEJ
   
   NUASI$ = NUPROA1$(HOY(HO$))
   '
   OBS$ = "ASIENTO X IMPORTADOS (" & FACTURA$ & ")"
   
'   On Error GoTo ERROR_GUARDAR
'   FLEXCONN.BeginTrans

   Dim EncaSiGra As ADODB.Recordset
   Set EncaSiGra = New ADODB.Recordset
   strSQL = "Select * from EncaBeAsien where NuIDAsien = '" & NUASI$ & "'"
   EncaSiGra.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With EncaSiGra
         .AddNew   ' YA CONVERTIDO
         !StatAsien = (-1)
          On Error Resume Next
          !CODIEMPR = CodiEmp%
          On Error GoTo 0
          !nuidasien = NUASI$
          If Text4 = "" Then Text4 = " "
          !refeasien = OBS$
             FECHAU% = HOY(HO$)
          !FECHASIEN = CVDAT(FECHAU%)
             PPXX% = InStr(NUASI$, "-")
          !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
          !SucuAsien = UNINEGO%
          !NuDfAsien = 0
          !TIPOASIEN = 11
          !StatAsien = 1  ' activo
          !EditAsien = -1
          !QPasAsien = 3
          !DebeAsien = IMPORTE
          !Observa = OBS$
          !Nume_Usuar = USNBR%
          !Niv_Clasi = 0
          !FeReAsien = Now
          !Origen = Left$(FACTURA$, 64)
          .Update
   End With
   
   NICLA% = 0
   ORPASE% = 0
   '
   Set DetasiGra = New ADODB.Recordset
   strSQL = "Select * from DetaAsien where NuIDAsien = '" & NUASI$ & "'"
   DetasiGra.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With DetasiGra
     For K2% = 1 To 2
          Select Case K2%
            Case 1
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next ' ASIENTO DE TOTAL CONTRAPARTIDA DE CUENTA CONTABLE DE REPARACION
              !CODIEMPR = CodiEmp%
              On Error GoTo 0
              !nuidasien = NUASI$
              ORPASE% = ORPASE% + 1
              !OrNuPase = ORPASE%
              !RefePase = Left$(DENOCUE$(CTACONTRAPARTIDA%), 24)
              FECHAU% = FEHOY%
              !FECHASIEN = CVDAT(FECHAU%)
              AA = CVDAT(FECHAU%)
              PPXX% = InStr(NUASI$, "-")
              !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
              !SucuAsien = 0
              !NuDfAsien = 0
              !StatPase = 1
              !CUECOEXT = CTACONTRA$
              !CUECOINT = CTACONTRAPARTIDA%
              !idebepase = IMPORTE
              !ihabepase = 0
              !MARCONSI = ""
              !CenCos = ""
              !Niv_Clasi = 0
              !DATRETPER = ""
             .Update
             
            Case 2
              If IMPORTE <= 0 Then GoTo 80 ' NEXT
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next 'ASIENTO DE REPUESTOS
              !CODIEMPR = CodiEmp%
              On Error GoTo 0
              !nuidasien = NUASI$
              ORPASE% = ORPASE% + 1
              !OrNuPase = ORPASE%
              !RefePase = Left$(DENOCUE$(CTAREPU%), 24)
              FECHAU% = FEHOY%
              !FECHASIEN = CVDAT(FECHAU%)
              PPXX% = InStr(NUASI$, "-")
              !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
              !SucuAsien = 0
              !NuDfAsien = 0
              !StatPase = 1
              !CUECOEXT = CTAREP$
              !CUECOINT = CTAREPU%
              !idebepase = 0
              !ihabepase = IMPORTE
              !MARCONSI = ""
              !CenCos = ""
              !Niv_Clasi = 0
             !DATRETPER = ""
             .Update
          End Select
          '
80        Next K2%
          DetasiGra.Close
          Set DetasiGra = Nothing
   End With
   '

90
End Sub


Sub GRABAR_ASIENTO_IMPORTADO(REMILAR$, NUFACLAR$)
    AcumulaImporteAsiento = 0
    SQLX$ = "select MOV.COD_INTE,MOV.IDENLOTE, MOV.CANTSALID,MOV.CANTINGRE "
    SQLX$ = SQLX$ + " from MOVISTO MOV WITH(NOLOCK) INNER JOIN MASTER MAS ON MAS.CODINT = MOV.COD_INTE"
    SQLX$ = SQLX$ + " WHERE  DOPRILAR = '" & REMILAR$ & "' AND DOSECLAR = '" & NUFACLAR$ & "' "
    SQLX$ = SQLX$ + " AND MAS.TRAZABLE = 1 AND MOV.COD_INTE > 0"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
         CI1% = XVALO(!cod_inte)
         LOTE$ = SAFETEXT$(!idenlote)
         CANTI = XVALO(!CANTSALID) - XVALO(!CANTINGRE)
         AcumulaImporteAsiento = AcumulaImporteAsiento + ControlAsientoImpo(CI1%, CANTI, LOTE$, 1)
         .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    '
    If Abs(AcumulaImporteAsiento) > 0 Then
       Call GRABAR_ASIENTO_ITEM_IMPORTADO(NUFACLAR$, REMILAR$, AcumulaImporteAsiento)
    End If
    
End Sub




Sub IMPRIMENOTAENTREGA()

       FORIPRE$ = TRIM$(Control$("", "FORNOTEN"))
       If TRIM$(FORIPRE$) = "" Then Exit Sub
       '
       PREFIREM$ = Control$("", "PREFIREM")
       If PREFIREM$ = "" Then PREFIREM$ = "0001"
       NroRemi$ = TRIM$(Text10.TEXT)
       While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
       NroRemi$ = PREFIREM$ + "-" + NroRemi$
       '
       NCLIE = XVALO(Text1.TEXT)
       FECHDOC$ = Text9.TEXT
       NUMEDOC$ = NroRemi$
       TIPODOC$ = "Nota de Entrega"
       '
       CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(TRXP$, 1, 32))
       CODPARAM$(2) = "CUIT-TRA": DOCPARAM$(2) = TRIM$(Mid$(TRXP$, 65, 16))
       'CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(TRXP$, 129, 24))
       CODPARAM$(3) = "NRO-REMI": DOCPARAM$(3) = TRIM$(NroRemi$)
       CODPARAM$(4) = "DOMI-TRA":  DOCPARAM$(4) = LENTREGA(0).TEXT
       CODPARAM$(5) = "LOCA-TRA":  DOCPARAM$(5) = TRIM$(Mid$(TRXP$, 159, 32))
       CODPARAM$(6) = "TRANSPOR": DOCPARAM$(6) = LENTREGA(2).TEXT
         NTRA% = XVALO(Text19)
       '
       CAPARDOC% = 6
       '
       CODTABU1$(1) = "COD-MAT"
       CODTABU1$(2) = "DES-MAT"
       CODTABU1$(3) = "CANTIDAD"
       CODTABU1$(4) = "ORD-ITEM"
       CACOLTAB1% = 4
       '
       DESTIDOCNUE = NCLIE
       CAITAB1% = 0
       '
       '
       ARCHILIB$ = "!1CARDEFA"
       For U& = 1 To ULTREG&("!CARDEFAC")
         XXX$ = REGLEIDO$("!CARDEFAC", U&)
         CANS$ = Str$(CVS(Left$(XXX$, 4))) ' cantidad
         NC1% = XVALO(Text1.TEXT)
         CIPRE% = CVI(Mid$(XXX$, 5, 2))
         'If XVALO(CANS$) >= 0.05 Then
           If TRIM$(Mid$(XXX$, 387, 16)) = "" Then
             CAITAB1% = CAITAB1% + 1
             DDTXT$ = Mid$(XXX$, 7, 64)                  ' descripcion
             '
             CDXXS$ = CODEXT$(CVI(Mid$(XXX$, 5, 2)))     ' codigo
             '\\\OT-06-0339-MH-24/08/06///
             If REPLAORI% = 1 Then
               ' se pasa el codigo interno, mas el nº cliente(NC1%) y trae el codigo externo solicitado
               CDXXS$ = CODORIG$(CIPRE%, NC1%) ' + SUBCO$
             End If
             '\\\OT-06-0339-MH-FIN///
             TETABU1$(1, CAITAB1%) = CDXXS$
             TETABU1$(2, CAITAB1%) = DDTXT$
             TETABU1$(3, CAITAB1%) = CANS$
             TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
             '
             'IMPRIME DETALLE DE LOTES REMITIDOS
             Call DETLOTREM(CIPRE%, U&)
             '
             Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
             '
 '
           End If
         'End If
       Next U&
       '
       '
80     CACOPDOC% = XVALO(FINFACTUR07.Text2(1).TEXT)
       Call SETPRINPORT("RESTORE")
       Call PRINTDOC("@" + FORIPRE$) ' Nota de Remito
       '
End Sub





Sub APLICACOBRA()
    '
    '
    'GRABO AL COMPROBANTE COMO QUE YA SE REALIZÓ LA CONTRAPARTIDA
    Fila$ = REDOCEL.List1.List(REDOCEL.List1.ListIndex)
    FACTURA$ = Mid$(Fila$, 10, 18)
    '
    'NUMERO DE FACTURA CORTO DEL CAJABANC
    CONDI$ = "CODICOM_LAR = '" & FACTURA$ & "' AND TIPOMOVIM = 'VENTA'"
    FACTUCOR$ = VALOBADA("CAJABANC", "CODICOM_COR", CONDI$, "NOMESS")
    NC% = VALOBADA("CAJABANC", "NUCLIEN", CONDI$, "NOMESS")
    '
    'TOTAL ORIGINAL DEL DOCUMENTO
    TotOrig# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", CONDI$, "NOMESS"))
    If TotOrig# < 0.05 Then
       TotOrig# = VALOBADA("CAJABANC", "IHABECOMP", CONDI$, "NOMESS")
    End If
    '
    'IMPORTE DEL CAJABANC SEGUN DOCUMENTO GENERADO (NOTA DE CREDITO)
    Condi1$ = "CODICOM_LAR = '" & CoComLar$ & "' AND TIPOMOVIM = 'VENTA' AND NUCLIEN = " & NC% & " "
    ImpApli# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", Condi1$, "NOMESS"))
    If ImpApli# < 0.05 Then
       ImpApli# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", Condi1$, "NOMESS"))
    End If
    '
    DocOrLar$ = FACTURA$
    DocOrCor$ = FACTUCOR$
    RefApli$ = "Reaplicación Automática"
    IvaOrig# = 0
    '
    Call GRAPLICOB
    '
    RECUEC07.Text1 = NC%
    Call RECUEC07.Command4_Click
    '
    'ACTUALIZO EL LIST PARA QUE ESTE COMPROBANTE NO APAREZCA
    Call REDOCEL.Actualizar_Lista
    '
End Sub

Function TIENEPROFORMA%(NC)
   TIENEPROFORMA% = CantRegistros&("PROFORMTL", "NUCLIEN = " & NC, "NOMESS")
End Function

Function TieneProformaPendienteDeFacturar%(NC)
   CONDI$ = "(STATUS = 0 OR STATUS IS NULL ) AND STATUS < 9 AND (FacturaReal = '' OR FacturaReal IS NULL)AND NUCLIEN = " & NC
'   TieneProformaPendienteDeFacturar% = CantRegistros&("PROFORMTL", "NUCLIEN = " & NC, "NOMESS")
    TieneProformaPendienteDeFacturar% = CantRegistros&("PROFORMTL", CONDI$, "NOMESS")
End Function

Private Sub BASIVA_Change(Index As Integer)
   If Index < 2 Then
      Total(Index + 3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(Index)) * XVALO(PORIVA(Index).Caption) * APLIVA% / 100), 3, 12, 2, 2)
      INOGRA# = XVALO(Total(2)) - XVALO(BASIVA(0)) - XVALO(BASIVA(1))
      TLABEL28 = TRIM$(FORMATNUM$(INOGRA#, "F12.2"))
      Call CALTOFA2
   End If
End Sub

Public Sub BASIVA_DBLClick(Index As Integer)
   '
   If Index <= 1 Then
     If Label6.Caption = "(T.Libre)" Then
       NCLIE% = XVALO(Text1)
       RNI% = 0: APLIVA% = 1
       If NCLIE% > 0 Then
         PPVI% = PIVACLI%(NCLIE%)
         If PPVI% = 2 Then RNI% = 1
         If PPVI% = 5 Or PPVI% = 7 Then APLIVA% = 0
       End If
       '
       Select Case Index
       Case 0 '21%
          If XVALO(BASIVA(0)) < 0.005 Then
            TASACLICK = XVALO(PORIVA(0).Caption)
            BASIVA(0) = BASIVA(1)
            BASIVA(1) = ""
          End If
          Call CALTOFA2
       Case 1 '10.5%
          If XVALO(BASIVA(1)) < 0.005 Then
            TASACLICK = XVALO(PORIVA(1).Caption)
            BASIVA(1) = BASIVA(0)
            AA1 = XVALO(BASIVA(1))
            aa2 = PORIVA(1)
            aa3 = APLIVA%
            BASIVA(0) = ""
          End If
          Call CALTOFA2
       End Select
     End If
   End If
   '
End Sub

Private Sub BASIVA_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then GoTo 99
   If Index > 1 Then GoTo 98
   If (TIFADOC% < 3 And MODODOC% = 0) Then GoTo 98 ' DESCARTA SOLAMENTE LAS FACTURA DE MAT. CODIFICADO
   Exit Sub
   '
98 On Error Resume Next
   PORDESCU.SetFocus
   On Error GoTo 0
   Exit Sub
   '
99 On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub BASIVA_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call BASIVA_LostFocus(Index)
   End If
End Sub

Private Sub BASIVA_LostFocus(Index As Integer)
   BASIVA(Index) = TRIM$(FORMATNUM$(BASIVA(Index), "F13.2"))
End Sub

Private Sub BASIVA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     '
     If Index <= 1 Then
        If Label6.Caption = "(T.Libre)" Then
            BASIVA(Index).ToolTipText = "Doble Click para Cambiar Tasa IVA"
        Else
            BASIVA(Index).ToolTipText = ""
        End If
     End If
     '
End Sub

Private Sub Check1_Click()
   If MUECOMHAS% <= 1 Then MUECOMHAS% = Check1.Value
End Sub



Private Sub Command1_Click()

   NCLIE = ListaSeleccionClientes()
   If NCLIE > 0 Then
     Text1.TEXT = TRIM$(Str$(NCLIE))
     If XVALO(Text1.TEXT) > 0 Then
       If XVALO(Text1.TEXT) <= 999999 Then
            Call CARDACLI
            Call Label12_Change
       End If
     End If
   End If

'   Call SELECHO("DEUDOR1")
'   NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
'   If NCLIE > 0 Then
'     Text1.TEXT = TRIM$(Str$(NCLIE))
'     If XVALO(Text1.TEXT) > 0 Then
'       If XVALO(Text1.TEXT) <= 999999 Then
'         NCLIE = XVALO(Text1)
'         'If REGICLI&(NC%) > 0 Then
'           Call CARDACLI
'TXTRAZA$ = TXTRAZA$ & "ANTES DE LABEL12_CHANGE= " & Now & vbCrLf
'           Call Label12_Change
'TXTRAZA$ = TXTRAZA$ & "DESPUES DE LABEL12_CHANGE= " & Now & vbCrLf
'         'End If
'       End If
'     End If
'   End If
End Sub



Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG07", "Archivos Maestros")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   If EXISTE%(LUDAT$ + "CTASFACT.RFS") < 1 Then GoTo 10
   If ULTREG("CTASFACT") > 0 Then
      Call SELECHO("CTASFACT")
      VDEVU$ = TRIM$(VALACT1$("CTASFACT"))
      If VDEVU$ <> "" Then
        Text16 = CODCUE$(XVALO(VDEVU$))
      End If
   Else
   
10   Call SELECHO("ECUECON")
     VDEVU$ = TRIM$(VALACT1$("ECUECON"))
     If VDEVU$ <> "" Then
       Text16 = VDEVU$
     End If
   End If
   
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   COGEVE07.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   'REFRESCA EL RECORDSET QUE SE GENERA CON LOS DATOS DEL CLIENTE
   a1 = STATCLI%(999): a1 = STATCLI%(XVALO(Text1))
   If VALICLIE% < 1 Then Exit Sub
   
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EMPREAC$, "EYM") > 0 Then 'SI ES EYM
      If XVALO(Label31) > 0 Then      'SI ES REPARACION
        If Abs(XVALO(Total(6))) < 0.005 Then
           Command14.Enabled = False
           If GRABASOLOREMITOREPARACION% > 0 Then
             Text1 = ""
             Command14.Enabled = True: Exit Sub
            End If
        End If
      End If
    End If
      
   If Abs(XVALO(Total(6))) < 0.005 Then Exit Sub
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 1 = " & Now & vbCrLf
   If VALIDAPUNVEN% < 1 Then Exit Sub
   '
' MsgBox MONEMI%
   Command14.Enabled = False
   Text1.Enabled = False
   Command1.Enabled = False
   '
   'se valida si es factura E que tenga el punto de venta correspondiente.
   If VERIFSALIDA%("NOMESS") > 0 Then
      If Label13 = "E" Then
        PVTAFAEXP$ = Control("FACTURA", "PUVEFAEX")  'P.VTA. FACT.EXP.
        If Label12 <> PVTAFAEXP$ Then
         Call COMUNI("Atención!!!\\Punto de Venta Seleccionado Difiere de La Condición de Iva del Cliente")
         GoTo 99
        End If
      End If
   Else
      If Label13 = "E" And ModoExportaSimple% = 0 Then
        PVTAFAEXPRINT$ = Control("FACTURA", "PUVEXPRI")  'P.VTA. FACT.EXP. MODO PRINTER
        If Label12 <> PVTAFAEXPRINT$ Then
           Call COMUNI("Atención!!!\Punto de Venta Seleccionado Difiere de La Condición de Iva del Cliente")
           GoTo 99
        End If
      End If
      
      If ModoExportaSimple% = 1 And Label13 <> "E" Then
        Call COMUNI("Atención!!!\Punto de Venta Seleccionado Difiera de la Condición de Iva del Cliente")
        GoTo 99
      End If
   End If
   
' suprimida esta validacion el 30/12/2010 - s/flavio es valido
' hacer facturas de exportacion de texto libre.
'   'VALIDA SI LA FACTURA ES POR TEXTO LIBRE Y TIENE MODALIDAD DE FACTURA ELECTRONICA
'   ' Y EL CLIENTE ES DE EXPORTACION NO PERMITA SEGUIR
'   If FACELEC% = 1 And Label6.Caption = "(T.Libre)" Then
'      If TRIM$(Label12) = TRIM$(Control("FACTURA", "PUVEFAEX")) Then
'        Call COMUNI("Imposible Realizar Factura Electrónica por Texto Libre\ Para Cliente de Exportacion ")
'        Command14.Enabled = True: Exit Sub
'      End If
'    End If
    'se traslada esta validacion a finfactu107 - verificado 30/12/2010
    ' que es correcta esta supresion según flavio y omar
''   If TRIM$(Control$("STOCKS", "CONEGATI")) = "SI" And Label6.Caption <> "(T.Libre)" And CONSICLI% < 1 And TIFADOC% <= 3 Then
''     If VALIDASTOCK < 1 Then Command14.Enabled = True: Exit Sub
''   End If
   '
'   'VALIDACION CUANDO ES NOTA DE CREDITO PARA DEFINIR SI QUIERE ASIGNAR A UNA FACTURA ANTICIPO
'   NC = XVALO(Text1)
'   FacturaSeleccionada$ = ""
'   If TIFADOC% = 6 Then
'    If TIENE_FACTURAS_ANTICIPO%(XVALO(Text1)) > 0 Then
'       OPX% = COMALTER%("Este Cliente Opera Con Facturas Anticipo\\Generar Nota de Crédito Común\Seleccionar Factura a Relacionar\Generar con Modalidad Anticipo con Aplicación Posterior\Cancelar ")
'       If OPX% = 2 Then
'          'MUESTRA LAS FACTURAS CON SALDO PENDIENTE- VALIDO QUE EL IMPORTE DE LA NC NO SEA MAYOR QUE LA FACTURA A APLICAR
'          FacturaSeleccionada$ = TRIM$(MostrarFactAnticConSaldo$(NC))
'          SaldoAnticipo# = SaldoPendFacturaAnticipo#(NC, FacturaSeleccionada$)
'          If SaldoAnticipo# > 0 Then
'            If XVALO(Total(6)) > SaldoAnticipo# Then
'              Call COMUNI("Imposible Aplicar !!! \El Importe de la Nota de Crédito Supera el Saldo (" & TRIM$(FORMATNUM$(SaldoAnticipo#, "F14.2")) & ") de La Factura Anticipo Seleccionada ")
'              GoTo 99
'            Else
'              'CASO QUE ELIGIO FACTURA
'              FACTUANTICIPO% = 2 'lo marco como 2 para identificar que es nota de credito, y en grafactura realice el movimiento de actualizaar la tabla de factuanti.
'            End If
'          Else
'            GoTo 99 ' SI NO ELIGIO FACTURA CANCELA.
'          End If
'
'       ElseIf OPX% = 3 Then
'          FACTUANTICIPO% = 1
'       ElseIf OPX% = 4 Then
'          GoTo 99
'       End If
'    End If
'   End If
   'VALIDACION QUE SOLO SEA MONEDA $

'   If MONEMI% > 1 And PIVACLI%(NC1&) <> 5 Then
'         Call COMUNI("Imposible Generar Factura Local en Moneda Extranjera (Solo En Pesos)")
'         GoTo 99
'   End If
   NC1& = XVALO(Text1)
   ' CONTROL DE ANULACION MEDIANTE NOTA DE CREDITO PARA FACTURAS CON PERCEPCION DE IIBB CABA
   If XVALO(Text24) > 0.005 And AnulaFacturaComprobante$ <> "" Then
        ImporteTotalFactura# = XVALO(VALOBADA("CAJABANC", "IMPO_DOLA", "CODICOM_LAR='" & AnulaFacturaComprobante$ & "' AND TIPOMOVIM='FVEN' AND CLI_REAL=" & NC1&))
        ImporteTotalNCredito# = XVALO(Total(6))
        ImportePercepIBFactura# = XVALO(VALOBADA("CAJABANC", "IMPO_DOLA", "CODICOM_LAR='" & AnulaFacturaComprobante$ & "' AND TIPOMOVIM='IBRU' AND OPCIMOVIM='IB-C' AND CLI_REAL=" & NC1&))
        ImportePercepIBNCredito# = XVALO(Text24)
        '
        If (Abs(ImporteTotalFactura# - ImporteTotalNCredito#) > 0.005) Or (Abs(ImportePercepIBFactura# - ImportePercepIBNCredito#) > 0.005) Then
            Call COMUNI("IMPOSIBLE CONTINUAR.\ \El Total de la Nota de Crédito o la Percepción de IIBB CABA\No Coincide con los Mismos Importes de la Factura a Anular\'" & AnulaFacturaComprobante$ & "'.\ \Para Notas de Créditos Parciales,\Utilice la Modalidad 'N/Crédito Abierta'")
            GoTo 99
        End If
   End If
   
   'VALIDACION DE FACTURA ANTICIPO
   
   If Check2.Value > 0 And Check2.Visible = True Then
      '
      If (TIFADOC% < 2 Or TIFADOC% > 3) And MODODOC% <> 1 Then
         Call COMUNI("Factura Anticipo Solo Para Factura o N.Débito de Texto Libre")
      End If
      '
'      If MONEMI% > 1 Then
'         Call COMUNI("Imposible Generar Factura Anticipo en Dolares (Solo En Pesos)")
'         GoTo 99
'      End If
      '
      TX$ = "Se Va a Generar una " & Label7 & "  Anticipo"
      OPX% = COMALTER%(TX$ & " \\Continuar\Cancelar")
      If OPX% = o Or OPX% = 2 Then GoTo 99
      '
      FACTUANTICIPO% = 1 '
      '
   Else
      
      Call Aplica_Fact_Anticipo(NC1&)
   End If
   '
   If Text22.Visible = True And TRIM$(Text22) <> "" Then
      If IsNumeric(Text22.TEXT) <> True Then
        Call MENSERR(24, "Comisión No Válida")
'        Command14.Enabled = True
'        Exit Sub
         GoTo 99
      End If
   End If
   '
   VDDOR% = XVALO(TEXT21(2).TEXT)
   If TRIM$(TEXT21(2).TEXT) <> "" And VDDOR% <> 0 Then
      If ECOARCH$("VENDEDOR", Chr$(VDDOR%)) = "" Then
         Call COMUNI("Vendedor: " & VDDOR% & " No Válido\Debe Asignar un Vendedor a La Factura")
         fech = Text9
         TIPFACT$ = "FC-"
         If Label7 = "N.Débito:" Then TIPFACT$ = "ND-"
         If Label7 = "N.Crédito:" Then TIPFACT$ = "NC-"
         
         NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
    
         NFACT$ = TIPFACT$ & Label12 & "-" & NFCZ$ & "-" & Label13
         NCLI = XVALO(Text1)
         RazonSoc$ = TRIM$(Text2)
         Call REPLA(VDEF$, MKI$(CVINT(fech)), 1, 2) ' INFORMACION PARA EL OBJETO
         Call REPLA(VDEF$, NFACT$, 3, 18)
         Call REPLA(VDEF$, MKS$(NCLI), 21, 4)
         Call REPLA(VDEF$, RazonSoc$, 25, 30)
        '
15       OBX$ = OBJPLA$("ACTUVENDE", VDEF$) ' ABRE EL OBJETO
         If OBX$ = "" Then 'SI CANCELA
            Call COMUNI("Debe Asignar un Vendedor a La Factura")
'            Command14.Enabled = True
'            Exit Sub
             GoTo 99
         End If
         '
         On Error Resume Next ' POR SI INGRESA UN VENDEDOR > 255
         VDDOR% = CVI(Mid$(OBX$, 55, 2)) ' VALIDA EL NUMERO DE VENDEDOR
         If VDDOR% < 1 Or ECOARCH$("VENDEDOR", Chr$(VDDOR%)) = "" Then
           On Error GoTo 0
           Call COMUNI("Vendedor No Válido")
           VDEF$ = OBX$
           GoTo 15
         End If
         On Error GoTo 0
         TEXT21(2) = VDDOR%
         
      End If
   End If
   '
   'VERIFICACION SI UTILIZA OPCION DE FACTURACION Y PEDIDOS CON ITEMS SIN CARGO PARA gra
   If Control("ADMIPED", "ITEMSCGO") = "SI" Then 'CONFIGURACION DE MEDITEA ITEMS SIN CARGO
      If TIFADOC% = 2 And MODODOC% = 0 Then
        HAY_ITEM_SIN_CARGO% = 0
        For i& = 1 To MSF.Rows - 1
          If XVALO(MSF.TextMatrix(i&, 7)) <= 0.01 Then
             HAY_ITEM_SIN_CARGO% = 1
             Exit For
          End If
        Next i&
        If HAY_ITEM_SIN_CARGO% > 0 Then
           CTA1% = CUECOINT%(Control("ASVENTAS", "ASITEMSC")) 'CUENTA  DE ASIENTO DE VENTAS ITEMS SIN CARGO)
           CTA2% = CUECOINT%(Control("ASVENTAS", "CONTASSC")) 'CUENTA DE CONTRAPARTIDA DE ASIENTO DE VENTAS ITEMS SIN CARGO
           If CTA1% < 1 Or CTA2% < 1 Then
              Call COMUNI("Imposible Continuar !!!\Falta Configurar Cuenta de Items Sin Cargo\Para Generar Asiento Contable")
              GoTo 99
           End If
        End If
      End If
   End If
   '
   If MODODOC% = 0 Then
    If HAY_ITEM_SIN_CARGO% > 0 Then 'CONFIGURACION DE MEDITEA ITEMS SIN CARGO
       For J& = 1 To MSF.Rows - 1
           If MSF.TextMatrix(J&, 0) = "*" Then
              If XVALO(MSF.TextMatrix(J&, 8)) <= 0.01 Then ' IMPORTE
                 CI1% = CODINT%(MSF.TextMatrix(J&, 2))
                 COSTOX = COSCALC#(CI1%)
                 If COSTOX <= 0 Then
                   COSTOX = COSPESOS(CI1%) ' COSTO REPOSICION
                 End If
                 If COSTOX <= 0 Then
                    Call COMUNI("Imposible Continuar !!!\Falta Informar Costo\Para Item ( " & CODEXT$(CI1%) & ") Sin Cargo ")
                    GoTo 99
                 End If
              End If
           End If
       Next J&
    End If
   End If
   If ESPROFORMA& > 0 Then
      OPX% = COMALTER%("Confirma Facturar Referenciando a la Proforma N°: " & ESPROFORMA& & "\\Cancelar\Facturar Referenciando Proforma")
      If OPX% < 2 Then
        ESPROFORMA& = 0
        GoTo 99
      End If
   End If
'5  If CONTROFIS% = 2 Then ' validacion de la longitud de razon social y el domicilio para impresora fiscal
'     If Len(Text2) > 30 Then
'       RZO$ = InputBox("Razon Social Supera Límite de Caracteres (30)", "Cambiar Manualmente Razon Social", Text2.TEXT)
'       If TRIM$(RZO$) = "" Then GoTo 5
'       Text2 = TRIM$(RZO$)
'       GoTo 5
'     End If
'     '
'     If Len(Text3) > 30 Then
'       RZO$ = InputBox("Domicilio Supera Límite de Caracteres (30)", "Cambiar Manualmente Domicilio", Text3.TEXT)
'       If TRIM$(RZO$) = "" Then GoTo 5
'       Text3 = TRIM$(RZO$)
'       GoTo 5
'     End If
'   End If
   '
   '\\\OT-07-0023-RG-24/01/07///
TXTRAZA$ = TXTRAZA$ & "PUNTO 2 = " & Now & vbCrLf
   FINFACTUR07.Text7 = TRIM$(Str$(DEPOEXPE%))
   '
   '\\\OT-07-0023-RG-FIN///
   If XVALO(TLABEL28) < 0 Then
      Call MENSERR(24, "Imposible Registrar este Documento.\Importe No Gravado no Admisible.")
'      Command14.Enabled = True
'      Exit Sub
       GoTo 99
   End If
   '
   
TXTRAZA$ = TXTRAZA$ & "PUNTO 3 = " & Now & vbCrLf

   ' *** CUENTA CONTABLE PERCEPCION DE IIBB PROVINCIA DE BUENOS AIRES
   If XVALO(Text15) >= 0.005 Then
     CUEPERI$ = Control$("PERIBRU", "CUECONTA")
     CUENPER% = CUECOINT%(CUEPERI$)
     If TRIM$(CUEPERI$) = "" Or CUENPER% < 1 Then
       Call MENSERR(24, "Imposible Registrar este Documento.\Falta Cuenta Contable Correspondiente\a Percepción de Ingresos Brutos\Provincia de Buenos Aires.")
        GoTo 99
     End If
   End If
   '
   ' *** CUENTA CONTABLE PERCEPCION IIBB CABA
   If XVALO(Text24) >= 0.005 Then
     CUEPERI$ = Control$("PERIBCA", "CUECONTA")
     CUENPER% = CUECOINT%(CUEPERI$)
     If TRIM$(CUEPERI$) = "" Or CUENPER% < 1 Then
       Call MENSERR(24, "Imposible Registrar este Documento.\Falta Cuenta Contable Correspondiente\a Percepción de Ingresos Brutos CABA .")
        GoTo 99
     End If
   End If
   
   ' *** CUENTA CONTABLE PERCEPCION DE IIBB PROVINCIA DE MISIONES
   If XVALO(Text25) >= 0.005 Then
     CUEPERI$ = Control$("PERIBMI", "CUECONTA")
     CUENPER% = CUECOINT%(CUEPERI$)
     If TRIM$(CUEPERI$) = "" Or CUENPER% < 1 Then
       Call MENSERR(24, "Imposible Registrar este Documento.\Falta Cuenta Contable Correspondiente\a Percepción de Ingresos Brutos\Provincia de Misiones.")
        GoTo 99
     End If
     
     ' PERCEPCION IIBB MISIONES (NO APLICA A CONSUMIDORES FINALES)
     ' SI LOS ARTICULOS DEL DETALLE SERAN UTILIZADOS COMO PARTE DE LA ACTIVIDAD (BIEN DE USO) NO SE DEBE PERCIBIR,
     ' SI EN CAMBIO SE TRATA DE UN ARTICULO DE REVENTA (BIEN DE CAMBIO) SI DEBE PERCIBIRSE
     If PIVACLI%(NC1&) <> 0 Then
        OpcionUsoDelBien% = COMALTER%("La Presente Factura cuenta con Percepción de IIBB Misiones.\Consultar si los Artículos del Detalle\serán Utilizados como Bien de Uso o de Cambio.\\Bien de Cambio\Bien de Uso")
        If OpcionUsoDelBien% = 2 Then
           PercepcionIBMisiones_BienDeUso = 1
           Call CALPERIB
        End If
     End If
   End If
   
   ' *** CUENTA CONTABLE PERCEPCION DE IIBB PROVINCIA DE TUCUMAN
   If XVALO(Text26) >= 0.005 Then
     CUEPERI$ = Control$("PERIBTU", "CUECONTA")
     CUENPER% = CUECOINT%(CUEPERI$)
     If TRIM$(CUEPERI$) = "" Or CUENPER% < 1 Then
        Call MENSERR(24, "Imposible Registrar este Documento.\Falta Cuenta Contable Correspondiente\a Percepción de Ingresos Brutos\Provincia de Tucumán.")
        GoTo 99
     End If
     
   End If
   
   ' *** VERIFICA QUE LA FECHA DE LA N CREDITO NO SUPERE LA FECHA LIMITE DE PRESENTACION,
   ' *** QUE EL IMPORTE DE LA N. CREDITO SEA MENOR AL DE LA FACTURA CUANDO ESTE DENTRO DEL MISMO MES
   ' *** Y QUE EL MISMO SEA IGUAL AL DE LA FACTURA CUANDO ESTE EN EL MES SIGUIENTE
   If XVALO(Text26) >= 0.005 And TIFADOC% = 6 Then
        Dim FechaFactura, FechaLimite, FechaNotaCredito As Date
        CondicionSql$ = "nuclien=" & Text1 & " and codicom_lar='" & AnulaFacturaComprobante$ & "' and tipomovim='ibru' and opcimovim='ib-tu' and varimovim='perc'"
        FechaFactura = VALOBADA("cajabanc", "fechconta", CondicionSql$)
        FechaLimite = DateAdd("m", 1, FechaFactura)
        DiaVencimiento% = XVALO(Control$("PERIBTU", "DIAVENCI"))
        FechaLimite = DateSerial(Year(FechaLimite), Month(FechaLimite), DiaVencimiento%)
        FechaNotaCredito = Now

        If FechaNotaCredito > FechaLimite Then
           Call MENSERR(24, "Imposible Incluir Percepción de I. Brutos Tucumán\Cuando la Factura que se está Cancelando\Fue Emitida Antes de Ultima Fecha de Presentación.")
           Text26 = ""
           GoTo 99
        End If

        CondicionSql$ = "CODICOM_LAR='" & AnulaFacturaComprobante$ & "' AND TIPOMOVIM='FVEN' AND CLI_REAL=" & NC1&
        ImporteFactura# = XVALO(VALOBADA("CAJABANC", "IMPO_DOLA", CondicionSql$))
        ImporteNotaCredito# = XVALO(Total(6))
        If Format(FechaFactura, "m") = Format(FechaNotaCredito, "m") Then ' NOTA DE CREDITO EMITIDA EN EL MISMO MES DE LA FACTURA
            If Abs(ImporteNotaCredito# - ImporteFactura#) > 0.005 Then
                Call MENSERR(24, "La Nota de Crédito Activa Intenta Cancelar\una Factura con Percep. IB Tucumán.\ \Este Documento puede Incluir Percepción\Sólo Cuando se esta Cancelando el Total de la Factura.")
                Text26 = ""
                GoTo 99
            End If
            ' *** SE ANULO EL HECHO DE QUE PUEDA EMITIR UNA NOTA DE CREDITO POR UN PARCIAL Y POR ENDE DEVOLVER PARCIALMENTE LA PERCEPCION DE LA FACTURA
            ' *** DEVUELVE LA PERCEPCION SOLO SI ESTA ANULANDO EL TOTAL DE LA FACTURA

'        If XVALO(Total(6)) > ImporteFactura# Then
'            Call MENSERR(24, "Imposible Devolver una Percepción de I. Brutos Tucumán\Mayor a los Percibido en la Factura que intenta Cancelar.")
'            GoTo 99
'        End If
        Else                                                              '
            If Abs(XVALO(Text2) - ImporteFactura#) > 0.005 Then
                Call MENSERR(24, "La Nota de Crédito Activa Intenta Cancelar\una Factura de un Mes Distinto con Percep. IB Tucumán.\ \Este Documento puede Incluir Percepción\Sólo Cuando se este Cancelando el Total de la Factura.")
                Text26 = ""
                GoTo 99
            End If
        End If
   End If
   
   ' *** FACTURA DE CREDITO - MONTO MINIMO Y CBU
   If TIFADOC% = 2 And chkModoCredito.Value = 0 And RegimenFacturaCredito%(NC1&) = 1 Then
        ImporteTotal# = XVALO(Total(6))
        If MONEMI% > 1 Then ImporteTotal# = ImporteTotal# * XVALO(TICAMBIS)
        MontoMinimo# = XVALO(Control$("FACTURA", "MOMINCRE"))
        If ImporteTotal# >= MontoMinimo# Then
            Call MENSERR(24, "Debido a que el Cliente que Ha Elegido\Se Encuentra Dentro del Régimen de Factura de Crédito\y el Monto Total de la Factura es Mayor al Mínimo del Régimen ( $" & TRIM$(FORMATNUM$(MontoMinimo#, "F10.2")) _
            & " ),\ Deberá Cambiar y Emitir una FACTURA DE CREDITO.")
            GoTo 99
        End If
   End If
   
   If TIFADOC% = 2 And chkModoCredito.Value = 1 Then
        CBU$ = TRIM$(Control$("FACTURA", "CBUCREDI"))
        If CBU$ = "" Then
            Call MENSERR(24, "Imposible Registrar Factura de Crédito.\Falta Configurar CBU.")
            GoTo 99
        End If
        
        ImporteTotal# = XVALO(Total(6))
        If MONEMI% > 1 Then ImporteTotal# = ImporteTotal# * XVALO(TICAMBIS)
        MontoMinimo# = XVALO(Control$("FACTURA", "MOMINCRE"))
        If ImporteTotal# < MontoMinimo# Then
            Call MENSERR(24, "Imposible Registrar Factura de Crédito.\El Importe Mínimo Para Este Tipo de Comprobante\es de $ " & TRIM$(FORMATNUM$(MontoMinimo#, "F10.2")) & ".")
            GoTo 99
        End If
        
        Mensaje$ = "Esta Por Registrar una Factura de Crédito\Con Fecha de Vencimiento '" & txtFechaVencimiento & "'\ \Ësta Seguro que desea Continuar?\\No, Cancelar\Si, Continuar"
        RESPUESTA% = COMALTER%(Mensaje$)
        If RESPUESTA% = 1 Then GoTo 99
   End If
   
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 4 = " & Now & vbCrLf
   FECHALIM% = XVALO(Control$("", "CIEMEVEN"))
TXTRAZA$ = TXTRAZA$ & "PUNTO 5 = " & Now & vbCrLf
   If FECHANUM(Text9) <= FECHALIM% Then
     Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
'     Command14.Enabled = True
'     Exit Sub
      GoTo 99
   End If
   '
18 EPAC$ = TRIM$(UCase$(EMPREAC$))
   '
   FINFACEX$ = ""
   MONECO$ = MONEMIS
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 6 = " & Now & vbCrLf
20 NCLIE = XVALO(Text1.TEXT)
'
   ' EXIGE EL INGRESO DE UN DNI, CUIL O PASAPORTE PARA LOS CONSUMIDORES FINALES. LO REQUIERE LA FACTURA ELECTRONICA
   If PIVACLI%(NCLIE) = 0 Then
      CuitCliente$ = CuitCli$(NCLIE)
      CuitCliente$ = TRIM$(InputBox("DNI, Pasaporte ...", "Documento de Consumidor Final", CuitCliente$))
      CuitCliente$ = SOLONUM$(CuitCliente$)
      If CuitCliente$ = "" Or Len(TRIM$(CuitCliente$)) > 13 Then
         Call COMUNI("Imposible Continuar con este Comprobante.\  \Falta Ingresar un Documento\o Ha Superado el Límite de Dígitos\(máximo 13 dígitos)")
         GoTo 99
      Else
          Text12.TEXT = TRIM$(POSIVAC$(NCLIE)) + " - CUIT " + CuitCliente$
      End If
   End If

   If VERLICRE% = 1 Then
     If TIFADOC% < 4 Then ' EXCLUYE CREDITOS
       '
       For U& = 1 To ULTREG&("!CARDEFAC")
         X$ = REGLEIDO$("!CARDEFAC", U&)
         NROPEDXX& = XVALO(TRIM$(Mid$(X$, 371, 16)))
         If NROPEDXX& < 1 Then GoTo 22
         RFF$ = RECFILT$("AUTORENT", 1, MKS$(NROPEDXX&))
         If Len(RFF$) < 4 Then GoTo 22 ' NO AUTORIZADO
         REGAUT& = CVS(RFF$)
         XAUT$ = REGLEIDO$("AUTORENT", REGAUT&)
         If UCase$(Mid$(XAUT$, 31, 2)) <> "SI" Then GoTo 22
       Next U&
       GoTo 25
       '
22     LMCR# = XVALO(Text7)     '  LIMICRE(NCLIE)
       SDEU# = SALDCLI#(NCLIE)
       CHPA# = CHEPEACRE#(NCLIE)
       TPED# = XVALO(Total(6).TEXT) * TICAMBX#
       If SDEU# + CHPA# + TPED# > LMCR# Then
         TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2))
         If TLIMI$ = "" Then TLIMI$ = "0.00"
         Call COMUNI("Imposible Emitir Este Documento.\  \Supera el Límite de Crédito Asignado\para esta Cuenta de $ " + TLIMI$ + ".\  \Informe al Supervisor de Ventas.")
'         Command14.Enabled = True
'         Exit Sub
          GoTo 99
       End If
     End If
   End If
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 7 = " & Now & vbCrLf
25 'VERIFICACION DE LIMITE DE CREDITO EN CUENTA CORRIENTE
   TPED# = XVALO(Total(6).TEXT) * TICAMBX#
   If EXEDELIMCC%(NCLIE, TPED#, DIFER) > 0 Then
     If TIFADOC% < 4 Then ' EXCLUYE CREDITOS
       LIMCCTE$ = TRIM$(FORMATNUM$(LCred_CCte(NCLIE), "f12.2"))
       TX$ = "Esta Factura + el Saldo Supera\el Límite de Crédito Asignado en C.Cte.\para esta Cuenta de $ " + LIMCCTE$
       If DERACCE%("AUTOFACT", "Autorizar Facturas con Exceso Credito") < 2 Then
         Call COMUNI("Imposible Emitir Este Documento.\  \" + TX$)
'         Command14.Enabled = True
'         Exit Sub
          GoTo 99
       Else
         If COMALTER%(TX$ + "\   \Puede Autorizar la Facturación\con Crédito Excedido.\\Cancelar\Autorizar Factura") < 2 Then
'            Command14.Enabled = True
'            Exit Sub
             GoTo 99
         End If
       End If
     End If
   End If
TXTRAZA$ = TXTRAZA$ & "PUNTO 8 = " & Now & vbCrLf
   ' ojo aca grabar lugares de entrega
   Call APERBASDAT("CLIEN")
   CONDICION$ = "Nume_Cli = " & NCLIE & ""
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 9 = " & Now & vbCrLf
   '
   If Control("FACTUREM", "AGRUPEDI") = "NOAGRUPA" Then
       TTXXT0$ = Left$(LENTREGA(0) + " ", 48)
       Call ACTUREGISTRO("DEUDOR12", "Lentre0_Cli", CONDICION$, TTXXT0$, REGAF&, "NOMESS")
       '
TXTRAZA$ = TXTRAZA$ & "PUNTO 10 = " & Now & vbCrLf
       TTXXT1$ = Left$(LENTREGA(1) + " ", 48)
       Call ACTUREGISTRO("DEUDOR12", "Lentre1_Cli", CONDICION$, TTXXT1$, REGAF&, "NOMESS")
       '
TXTRAZA$ = TXTRAZA$ & "PUNTO 11 = " & Now & vbCrLf
       TTXXT2$ = Left$(LENTREGA(2) + " ", 48)
       Call ACTUREGISTRO("DEUDOR12", "Lentre2_Cli", CONDICION$, TTXXT2$, REGAF&, "NOMESS")
   End If
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 12 = " & Now & vbCrLf
   NUOCOMPRA$ = ""
   NROPEDIDO$ = ""
   NROREMICO$ = ""
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 13 = " & Now & vbCrLf
   For U& = 1 To ULTREG&("!CARDEFAC")
     X$ = REGLEIDO$("!CARDEFAC", U&)
     NUOCOPAR$ = TRIM$(Mid$(X$, 355, 16))
     NROPEDIPAR$ = TRIM$(Mid$(X$, 371, 16))
     NROREMIPAR$ = TRIM$(Mid$(X$, 390, 13))
     '
     If NUOCOPAR$ <> "" Then
       If NUOCOPAR$ <> NUOCOPARA$ Then
         If NUOCOMPRA$ <> "" Then NUOCOMPRA$ = NUOCOMPRA$ + ", "
         NUOCOMPRA$ = NUOCOMPRA$ + NUOCOPAR$
         NUOCOPARA$ = NUOCOPAR$
       End If
     End If
     '
     If NROPEDIPAR$ <> "" Then
       If NROPEDIPAR$ <> NROPEDIPARA$ Then
         If NROPEDIDO$ <> "" Then NROPEDIDO$ = NROPEDIDO$ + ", "
         NROPEDIDO$ = NROPEDIDO$ + NROPEDIPAR$
         NROPEDIPARA$ = NROPEDIPAR$
       End If
     End If
     '
     If CONSICLI% = 1 Then
       If NROREMIPAR$ <> "" Then
         If NROREMIPAR$ <> NROREMIPARA$ Then
           If NROREMICO$ <> "" Then NROREMICO$ = NROREMICO$ + ", "
           NROREMICO$ = NROREMICO$ + NROREMIPAR$
           NROREMIPARA$ = NROREMIPAR$
         End If
       End If
     End If
     '
   Next U&
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 14 = " & Now & vbCrLf
30 If CONSICLI% = 1 Then
     REMI = 0
     '
     If TRIM$(Text14.TEXT) = "" Then
       Text14.TEXT = NUOCOMPRA$
     End If
     '
   End If
   '
   FORIPRE$ = TRIM$(Control$("", "FORFAC-" + Label13.Caption))
   If FORIPRE$ = "" Then
      Call MENSERR(24, "Falta Definir Controlador de Impresión")
      Exit Sub
   End If
   '
   'Se agrega if para que solo valide si no es factura electronica
   If FACELEC% < 1 Then
     If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Documento.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
'       Command14.Enabled = True
'       Exit Sub
        GoTo 99
     End If
   End If
   '
 TXTRAZA$ = TXTRAZA$ & "PUNTO 15 = " & Now & vbCrLf
'   If MODODOC% = 0 Then    ' Material Codificado
'     ARCHILOC$ = "!CARDEFAC"
'     Call CIERRARCH(ARCHILOC$)
'     If INTERFI$ = "" Then
'       INTERFI$ = "TRAN" + RANDSTRING$(-4)   ' negativo para que sean numeros
'     End If
'     ALOCDA$ = Mid$(ARCHILOC$, 2)
'     Call COPYSTRU(ALOCDA$, INTERFI$)
'     XY1% = COPYFILE(LUCOM$ + ALOCDA$ + ".WKF", LUDAT$ + INTERFI$ + ".DAT")
'     '
'     Call FILESORT(INTERFI$, "", 30, 30, "ASC")   ' POR NUMERO ORDEN ITEM
'     Call FILESORT(INTERFI$, "", 18, 18, "ASC")   ' POR REMITO
'     Call FILESORT(INTERFI$, "", 20, 20, "ASC")   ' POR FECHA REMITO
'     XY1% = COPYFILE(LUDAT$ + INTERFI$ + ".DAT", LUCOM$ + ALOCDA$ + ".WKF")
'     Call BAJALDISCO
'   End If
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 16 = " & Now & vbCrLf
   CARENGLO& = ULTREG&("!CARDEFAC")
   ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
   If ITPORHOJA% < 1 Then
     ITPORHOJA% = 32767
   End If
   '
   CAHOJNEC% = Int(0.98 + CARENGLO& / ITPORHOJA%)
   '
   '
   ' *** NOTA DE CREDITO Y DEBITO MODO CREDITO - IMPUTACION
   If (TIFADOC% = 3 Or TIFADOC% = 6) And (chkModoCredito.Value = 1) Then
           
        If TIFADOC% = 6 And AnulaFacturaComprobante$ <> "" Then
            AplicacionFacturaCredito$ = AnulaFacturaComprobante$
            GoTo 55
        End If

        ' ARMA LISTA DE FACTURAS A APLICAR
        Campos$ = "Fecha/D8;Factura/A24;Importe/G14.2"
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
        FRMZELECHO.Caption = "Facturas de Crédito"

        ' CARGA GRILLA CON LOS DATOS DE LA TABLA
        Dim RstFacturaCredito As ADODB.Recordset
        ConsultaSql$ = "select fechconta, codicom_lar, valabscomp from cajabanc with(nolock) where nuclien=" & Text1 & " and tipomovim='fven' and opcimovim='fc' and modoopcimovim='cr' and numecomp>0"
        Set RstFacturaCredito = READSET(ConsultaSql$)
        If Not (RstFacturaCredito.BOF And RstFacturaCredito.EOF) Then
            FRMZELECHO.MSF2.Rows = RstFacturaCredito.RecordCount + 1
        End If
        Fila& = 0
        With RstFacturaCredito
            Do While Not .EOF
                Fila& = Fila& + 1
                    FechaEmision$ = FECHATEX$(CVINT(!fechconta))
                FRMZELECHO.MSF2.TextMatrix(Fila&, 1) = FechaEmision$
                FRMZELECHO.MSF2.TextMatrix(Fila&, 2) = SAFETEXT$(!codicom_lar)
                FRMZELECHO.MSF2.TextMatrix(Fila&, 3) = TRIM$(FORMATNUM$(!ValAbsComp, "G14.2"))

                .MoveNext
            Loop
        End With
        On Error Resume Next
        RstFacturaCredito.Close
        Set RstFacturaCredito = Nothing
        On Error GoTo 0

        ' DETERMINA ANCHO A PARTIR DEL ANCHO DE LA GRILLA
        FRMZELECHO.Width = 7000
        ' DESPLIEGA LISTA
        FRMZELECHO.Show 1

        AplicacionFacturaCredito$ = TRIM$(RESP_ZELE_VECT(2))
        If AplicacionFacturaCredito$ = "" Then
            Call COMUNI("Imposible Continuar Sin Aplicar el Presente Documento\Con Alguna de las Facturas de Crédito Emitidas.")
            GoTo 99
        End If
        
        AnulaFacturaComprobante$ = AplicacionFacturaCredito$
   
55      TipoAplicacionFacturaCredito$ = "N"
        If TIFADOC% = 6 Then
            RESPUESTA% = COMALTER%("¿Se está Anulando una Factura de Crédito Rechazada por Cliente?\\No, Continuar\Si, Rechazada")
            If RESPUESTA% = 2 Then TipoAplicacionFacturaCredito$ = "S"
        End If
        
   End If
   
TXTRAZA$ = TXTRAZA$ & "PUNTO 17 = " & Now & vbCrLf
      
   Call GRAFACTURA(OKX%)
   '
   ' RECALCULA LA COMPOSICION DE DEUDA DEBIDO A LA APLICACION AUTOMATICA
   ' QUE SE HIZO EN LA MODALIDAD DE ANULA FACTURA
   If TIFADOC% = 6 And AnulaFacturaComprobante <> "" Then
        Call RECUEC07.REAPLICRE(NC1&)
   End If
   '
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
 TXTRAZA$ = TXTRAZA$ & "PUNTO 49 = " & Now & vbCrLf
   
   If OKX% < 0 Then
       GoTo 10     ' ABORTADA O SOLO REMITO
     ElseIf OKX% = 0 Then
       GoTo 18
   End If
   '
   
   If OKX% > 0 Then
       FF% = FECHANUM(Text9.TEXT)
       REMI = XVALO(Text10.TEXT)
       If FACELEC% < 1 Then
          If ModoExportaSimple% = 1 Then GoTo 60
          Call PRIFACTURA ' NO IMPRIME (DESPUER VER SI SE VA A IIMPRMIR PROVISORIO)
       Else
 TXTRAZA$ = TXTRAZA$ & "PUNTO 50 = " & Now & vbCrLf
       
          Call PRIREMIFA ' IMPRIME SOLO REMITO ' VER POR QUE EL REMITO IMPRMIE EL NUMERO DE FACTURA
 TXTRAZA$ = TXTRAZA$ & "PUNTO 51 = " & Now & vbCrLf
          CONDI$ = "CODICOM_LAR = '" & TRIM$(Label30) & "'"
          IDCAE$ = SAFETEXT$(VALOBADA("CAJABANC", "IDCAE", CONDI$))
 TXTRAZA$ = TXTRAZA$ & "PUNTO 52 = " & Now & vbCrLf
          TESTCAE09.YAGRABOCAE = 0
          TESTCAE09.Text5 = IDCAE$
          Call TESTCAE09.PIDENUEVOCAE
          'ESTAS 5 FILAS SE COMENTARON PARA DEJAR DE LADO LA ASIGNACION DE LOTES AUTOMATICA - DENUNCIO QUE NO SE LE CERRABA LA VENTANA
          If InStr(EPAC$, "AHP") > 0 And MODODOC% < 1 And TIFADOC% = 2 Then ' QUE SEA AHP, CODIFICADO Y FACTURA
            'Timer3.Enabled = True ' AHP 11 SE DESHABILITO Y SE LLAMA DIRECTAMENTE EN LA RUTINA GRABANUEFACTU
          Else 'MANTENER PARA QUE SI NO ES CODIFICADO Y  FACTURA LIMPIE EL LIST
            ASIGLOAHP.List2.Clear 'POR SI CARGO INFORAMCION EN EL LIST
          End If
       End If
       '
   End If
60 Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
   If REPRIFAC% = 1 Then
      Call FINAL("")
   End If
   '
   If OKX% > 0 Then
     If InStr(UCase$(Eco21(1)), "CONTADO") > 0 Or EsPorPedido& > 0 Then
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       If InStr(EPAC$, "DOSIVA") < 1 Then
          RECIBO08.Text1 = Text1
          NCLI = XVALO(TRIM$(Text1))
          RECIBO08.Text2 = Text2
          RECIBO08.Text3 = Text3
          RECIBO08.Text4 = Text4
          RECIBO08.Text5 = Text5
          RECIBO08.Text9 = Text9
          RECIBO08.Label26 = CODICLI$(NCLI) 'CODIGO DEL CLIENTE
          If ULTREG&("COBRADOR") > 0 Then
            XXX$ = REGLEIDO$("COBRADOR", 1)
            RECIBO08.Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
          End If
          If FACELEC% > 0 Then
              Timer2.Enabled = True  'SE HABILITA UN TIMER PARA QUE NO SE PISE LA GENERACION DEL RECIBO CON LA SOLICITUD DEL CAE Y GRABACION DEL MISMO.
          Else
              RECIBO08.Show 1
          End If
       End If
     End If
   End If
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 200.1 = " & Now & vbCrLf
10 Call BLANFORMU
TXTRAZA$ = TXTRAZA$ & "PUNTO 200.2 = " & Now & vbCrLf
   Call NUFACTURA
TXTRAZA$ = TXTRAZA$ & "PUNTO 200.3 = " & Now & vbCrLf
   Call CIERRARCH("*.*")
TXTRAZA$ = TXTRAZA$ & "PUNTO 200.4 = " & Now & vbCrLf
   Text1.TEXT = ""
TXTRAZA$ = TXTRAZA$ & "PUNTO 200.5 = " & Now & vbCrLf
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
   '
   If EsPorPedido& > 0 Then Call FINAL("")
   '
99 Command14.Enabled = True
   Text1.Enabled = True
   Command1.Enabled = True
End Sub
Sub GRABA_ASIENTO_FACTURA_ANTICIPO(NC, TOTIMPOX#, FACT_ANTICIPO$)
   CUDEB1% = CUECOINT%(Control$("SVD451S", "CUEVE-" + TRIM$(Str$(18)))) 'CUENTA DE ANTICIPOS
   If CUDEB1% < 1 Then Exit Sub
   '
   CUHAB1% = CUECOINT%(TRIM$(Control("SVD451S", "CUEVE-15"))) 'CUENTA DE VENTAS
   '
   FEDEP% = HOY(HO$)
   FEX = FEDEP%
   '
   ITRAN# = TOTIMPOX# 'IMPORTE APLICADO
   '
   'REFE$ = "As.Factura Contra Remito (Cta. " & TRIM$(Str(NC)) + ")-" + FECHATEX$(FEX)
   FACT_ANTICIPO1$ = Left$(FACT_ANTICIPO$, 3) & TRIM$(Str$(XVALO(Mid$(FACT_ANTICIPO$, 9, 8)))) & "-" & Right(FACT_ANTICIPO$, 1)
   'FC-0001-OO123456-A
   REFE$ = "P/Entrega S/Fc.Ant." & FACT_ANTICIPO1$ & "-" + FECHATEX$(FEX)
   '
   FETX$ = FECHATEX$(FEX)
   NRECIB$ = "AA-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-" 'APLICACION ANTICIPO
   NRECIBX$ = NRECIB$ + "__"
   ORMA% = (-1)
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   'With CajaBanc
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE CodiCom_Cor LIKE '" & NRECIBX$ & "'"
   Dim CAJABTEMP As ADODB.Recordset
   Set CAJABTEMP = READSET(SQLX$)
   On Error Resume Next
   With CAJABTEMP
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
          If XVALO(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
            ORMA% = XVALO(Mid$(!CodiCom_Cor, 11))
          End If
          .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set CAJABTEMP = Nothing
   ORMX$ = TRIM$(Str$(ORMA% + 1))
   While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
   NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
   '
   'MANTENGO LOS VALORES DE LAS VARIABLES PUBLICAS POR QUE SE VAN A CAMBIAR Y AL FINAL LAS RESTAURO.
   RESERVO_ANTCOR$ = CoComCor$
   RESERVO_ANTLAR$ = CoComLar$
   RESERVO_Referen$ = Referen$
   RESERVO_NumeCom = NumeCom
   RESERVO_Observa$ = Observa$
   RESERVO_CuentaCo% = CuentaCo%
   RESERVO_TiMovi = TiMovi
   RESERVO_ImpoDeb# = ImpoDeb#
   RESERVO_StatInit = StatInit
   RESERVO_NumeClie = NumeClie
   RESERVO_NumeTal = NumeTal
   RESERVO_Descrip$ = Descrip$
   '
   CoComCor$ = NRECIB$
   CoComLar$ = NRECIB$
   NumeClie = 0
   NumeProv% = 0
   FechEmi = CVDAT(FEDEP%)
   '
   Referen$ = REFE$
   FeContab = CVDAT(FEDEP%)
   FeAcredi = CVDAT(FEDEP%)
   NumeTal = Val(Mid$(NRECIB$, 11, 2))
   NumeCom = NumeCom
   Observa$ = "Asiento Aplicación Factura Anticipo"
   Descrip$ = REFE$
   '
   CuentaCo% = CUDEB1%
   TiMovi = "A-Anticipo"            '
'   Descrip = "De " + TRIM$(ECOARCH$("ECUECON", CUEDEBE$))
   ImpoDeb# = ITRAN#
   StatInit = 0
   Call GRACABAN
   '
   CuentaCo% = CUHAB1%
'   Descrip = "A " + TRIM$(ECOARCH$("ECUECON", CUEHABE$))
   ImpoDeb# = (-1) * ITRAN#
   Call GRACABAN
   
   'REPONGO LOS VALORES ANTERIORES POR LAS DUDAS.
   CoComCor$ = RESERVO_ANTCOR$
   CoComLar$ = RESERVO_ANTLAR$
   Referen$ = RESERVO_Referen$
   NumeCom = RESERVO_NumeCom
   Observa$ = RESERVO_Observa$
   CuentaCo% = RESERVO_CuentaCo%
   TiMovi = RESERVO_TiMovi
   ImpoDeb# = RESERVO_ImpoDeb#
   StatInit = RESERVO_StatInit
   NumeClie = RESERVO_NumeClie
   NumeTal = RESERVO_NumeTal
   Descrip$ = RESERVO_Descrip$
   '
End Sub
Private Sub Command23_Click()
    NDCLIEN07.Show 1
End Sub

Private Sub Command24_Click()
   'Call CommSel("TRANSPOR", Command24, Text19)
   
   Call COPYSTRU("TRANSPOR", "TRANSPO1")
   Call CARGALISEL("TRANSPO1", "TRANSPORTES", "CODI_TRA/F5;DENO_TRA/A32", "CODI_TRA > 0 ORDER BY DENO_TRA", "NOMESS")
   Call SELECHO("TRANSPO1")
   NTRAN% = XVALO(VALACT1$("TRANSPO1"))
   If NTRAN% > 0 Then
      Text19 = TRIM$(FORMATNUM(NTRAN%, "I4"))
   End If

End Sub

Private Sub Command25_Click()
   If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
      '\\\OT-06-0570-RG-27/12/06///
      Call CommSel("SUCENTRE", Command25, Text20)
      '\\\OT-06-0570-RG-FIN///
     'Call SELECHO("!SUCENTRE")
     'VDEVU$ = VALACT1$("!SUCENTRE")
     'Text20 = VDEVU$
   End If
End Sub

Private Sub Command26_Click()
    '
    Call FACLOT07.LIMPILABEL
    Call FACLOT07.LIMPITEXT
    Call FACLOT07.List2.Clear
    FACLOT07.Label15 = "ACTIVO"
    '
    Call ASIGLOT07.List1.Clear ' LIMPIO INFORMACION DEL ASIGLOT07
    Call ASIGLOT07.List2.Clear
    ASIGLOT07.Label6 = ""
    ASIGLOT07.Label10 = ""
    ASIGLOT07.Label4 = ""
    ASIGLOT07.Label8 = ""
    ASIGLOT07.Label1 = ""
    ASIGLOT07.Label3 = ""
    '
    FACLOT07.Show 1
    '
End Sub

Private Sub Frame13_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


Private Sub Command27_Click()
   Command27.Enabled = False
   '
   Call CREATABLA_OFABRIC
   'VALIDAR SI ESTA LA CUENTA CONTABLE CONFIGURADA
   If CUECOINT%(Control("SVD451S", "CUEVE-15")) < 1 Then
      Call COMUNI("Falta Configurar Cuenta Contable de Reparaciones")
      GoTo 99
   End If
   '
   'FUERZO QUE SEA FACTURA POR TEXTO LIBRE
   TIFADOC% = 2: Label7.Caption = "Factura:"
   Label6.Caption = "(T.Libre)": MODODOC% = 1
   '
   'ARMAR CARGALISEL REPARACIONES CON ORDEN DE REPARACION GENERADA O CERRADAS
   Screen.MousePointer = 11
   CONDI$ = "NUMREPA > 0 AND STATUS > 2 AND (FACTURA IS NULL OR FACTURA  = '')ORDER BY NUMREPA"
   Call CARGALISEL("RECERSFA", "REPARA", "NUMREPA/F6.0;REFEREN/A56", CONDI$, "NOMESS")
   Screen.MousePointer = 1
   Call SELECHO("RECERSFA/ORDENES DE REPARACION GENERADAS/CERRADAS SIN FACTURAR")
   NUREP& = XVALO(VALACT1("RECERSFA"))
   If NUREP& < 1 Then GoTo 99
   '
   MARCONOTA.Visible = True
   Label31 = TRIM$(Str$(NUREP&)) 'GUARDO NUMERO DE REPARACION EN LABEL
   Text16 = TRIM$(Control("SVD451S", "CUEVE-15"))

   Call Command6_Click
99 Command27.Enabled = True

End Sub

Private Sub Command31_Click()

    Command31.Enabled = False
   
    Campos$ = "Cta./F9;R.Social/A40"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)

   Call ABREPEDCLI
   SQLX$ = "SELECT distinct NroClie FROM peddeta with(nolock)"
   SQLX$ = SQLX$ + " where PEDDETA.STATUS >= 0 "   ' APROBADOS
   SQLX$ = SQLX$ + " AND PEDDETA.Status < 8 "    ' para que no este anulada ni cumplida
   SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
   'SQLX$ = SQLX$ + " ORDER BY peddeta.NroPed ASC "
   'Set RS = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   Set rs = READSET(SQLX$)
    J& = 0
    '
    With rs
      Do While Not .EOF
        NCLIE = XVALO(!NROCLIE)
        RAZCLI$ = rasocli$(NCLIE)
        '
        J& = J& + 1
        FRMZELECHO.MSF2.Rows = J& + 1
        FRMZELECHO.MSF2.TextMatrix(J&, 1) = NCLIE
        FRMZELECHO.MSF2.TextMatrix(J&, 2) = RAZCLI$
        .MoveNext
      Loop
    End With
    rs.Close
    
    Set rs = Nothing
    If J& = 0 Then
       Call COMUNI("No se Detectaron Entregas Ingresos Con Proveedor Para el Código Activo")
       Exit Sub
    End If
    
    FRMZELECHO.Caption = "Clientes Con Pendientes a Facturar "
    FRMZELECHO.Show 1
    
    NCLIE = XVALO(RESP_ZELE_VECT(1))
    If NCLIE > o Then
        Text1 = NCLIE
    End If

    Command31.Enabled = True

End Sub

Private Sub FactProfTlibre_Click()
   If Control("FACTURA", "PROFORTL") <> "SI" Then
    If CLACONTRA% > 0 Then
      Call GRACONTROL("FACTURA", "PROFORTL", "SI")
      'SOLO SE UTILIZA PARA QUE UNA VEZ DADO EL CLACONTRA NO LO VUELVA A PEDIR
    Else
      Exit Sub
    End If
   End If
   
   FACPROFORMA.Text1 = ""
   FACPROFORMA.Text11 = FORMATOCON0$(PROXIMONUMPROFORMA&, 8)
   FACPROFORMA.Text6 = "0000"
   FACPROFORMA.Show 1
End Sub
Sub REGISTRAPROFORMA_MATCOD()
    '
    If XVALO(Text1) < 1 Then
       Call COMUNI("Falta Número de Cliente")
       Exit Sub
    End If
    If MSF.Rows <= 1 Then
       Call COMUNI("Debe Primero Cargar el Detalle")
       Exit Sub
    End If
    '
    YAAVISE% = 0
    For KKU& = 1 To ULTREG&("!CARDEFAC")
                XXU$ = REGLEIDO$("!CARDEFAC", KKU&)
                NPEDXX& = XVALO(TRIM$(Mid$(XXU$, 371, 16)))
                If TRIM$(STRINGPED$) = "" Then STRINGPED$ = SAFETEXT$(NPEDXX&)
                If NPEDXX& <> NPEDXX_ANT& And NPEDXX& > 0 And NPEDXX_ANT& > 0 Then
                  If YAAVISE% < 1 Then
                     OPX% = COMALTER%("Atención Se Han Seleccionado Más de Un Pedido\\Cancelar\Continuar")
                     If OPX% < 2 Then Exit Sub
                  End If
                  If YAAVISE% < 1 Then
                     STRINGPED$ = SAFETEXT$(NPEDXX&)
                  Else
                     STRINGPED$ = STRINGPED$ & ";" & SAFETEXT$(NPEDXX&)
                  End If
                  YAAVISE% = 1
                End If
                NPEDXX_ANT& = NPEDXX&
                
    Next KKU&
            
    TIPOCOMP$ = "FP"
    NPROF& = PROXIMONUMPROFORMA&
    NUMEROTALO$ = XVALO(Label12)
    'CORTO
    NCCMP$ = TIPOCOMP$ & "-" & FORMATOCON0$(NPROF&, 4)
    'LARGO
    NLCMP$ = TIPOCOMP$ & "-" & TRIM$(Label12) & "-" & TRIM$(NPROF&)
    '
    TOTDOCUM# = XVALO(Total(6))
    NUTALO = Label12
    NC = XVALO(Text1)
    '
    'NUCOM& = PROXIMONUMPROFORMA&
    NC = XVALO(Text1)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    SQLX$ = "SELECT * FROM PROFORMTL WHERE NumeComp < 1 "
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    With rs
        .AddNew
        !NumeComp = NPROF&
        !NUMETALO = XVALO(NUTALO)
        !NUCLIEN = NC
        !CODICONPAG = XVALO(TEXT21(1))
        '!VEND_CLI = XVALO(TEXT21(0))
        !ImpoTot = XVALO(Total(6))
        !TOTNETO = XVALO(Total(2))
        !TOTBRTOSIVA = XVALO(Total(0))
        !FechFactura = CVDAT(CVINT("01/01/80"))
        !FacturaReal = ""
        !IMPCONDESC = XVALO(Total(2))
        !IMPADESCONTAR = XVALO(Total(1))
        !FECHEMISI = CVDAT(FECHANUM(Text9))
        !FeHoReal = Now
        !TDescApli = PORDESCU
        !Tipo_Cam = TICAMONE(MONEMI%)
        !Mone_Emis = MONEMI%
        !Status = 0
        
        !COTIZA_DOLAR = TICAMONE(2)
        REFEXX$ = "MATERIAL CODIFICADO CARGA MANUAL"
        If TRIM$(STRINGPED$) <> "" Then
           REFEXX$ = "Pedido/s: " & STRINGPED$
        End If
        !REFERENCIA = REFEXX$
        !NROPED = NPEDXX&
        !CODIEMPR = CodiEmp%
        
        
        
        !BASE1 = XVALO(BASIVA(0))
        !BASE2 = XVALO(BASIVA(1))
        !TASA1 = FORMATNUM$(XVALO(PORIVA(0)), "F6.2")
        !TASA2 = FORMATNUM$(XVALO(PORIVA(1)), "F6.2")
        !IMPOIVA1 = XVALO(Total(3))
        !IMPOIVA2 = XVALO(Total(4))
        !VARIMOVIM = Label13
        .Update
        '
    End With
    rs.Close: Set rs = Nothing
    
    Call GuardarDetalleFacturaProforma(NPROF&, MONEMI%, TICAMONE(MONEMI%), XVALO(TEXT21(1)), NC)
    
    If NPROF& > 0 Then
        FORIMPRE$ = TRIM$(Control("", "FAPRMACO"))
        If FORIMPRE$ = "" Then
            Call COMUNI("Falta Configurar Formulario de Impresión de Factura Proforma")
            Exit Sub
        Else
            FILTX$ = TRIM$(Str$(NPROF&) & ";" & FORMATOCON0$(NPROF&, 8))
            Call STDFORM(FORIMPRE$, FILTX$)
        End If
     End If

   '
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '
   'PONGO MARCA EN 0 DE BANDERA DE MODO DE RECUPERACION
   REDOCEL.MODORECU = 0
   '
End Sub

Sub Label12_DblClick()

   EPAC$ = TRIM$(UCase$(EMPREAC$))
   'If InStr(EPAC$, "DOSIVA") > 0 Then Exit Sub
   '
   If DERACCE%("CAMPREFI/NOMESS", "Modificación de Prefijo") < 2 Then ' derecho de acceso para modificar prefijo sin mensaje
      Exit Sub
   End If
   '
   'VALIDO SI YA ESTA USANDO LOS DATOS DE PUNTO DE VENTA DE FACTURACION COMUN
   PVTANORMALPRINT$ = Control("FACTURA", "PUVENPRI") 'P VTA. NORMAL MODO PRINTER
   PVTAFAEXPRINT$ = Control("FACTURA", "PUVEXPRI")  'P.VTA. FACT.EXP. MODO PRINTER
   PVTMODOCREDITO$ = Control("FACTURA", "PUVECRED")
   If Control$("FACTURA", "EXSIMPLE") = "SI" Then
        IDLIST& = ObtenerIdTablaAuxiliar("PTOVTAEXSIM")
        PtosVentaExportaSimple& = CantRegistros&("DATASQL", "ID_LISTA=" & CStr(IDLIST&))
   End If
   
   If XVALO(PVTANORMALPRINT$) + XVALO(PVTAFAEXPRINT$) + PtosVentaExportaSimple& + XVALO(PVTMODOCREDITO$) > 0 Then

    PVTANORMAL$ = Control("FACTURA", "PUVENELE") 'P VTA. NORMAL
    PVTABCPTAL$ = Control("FACTURA", "PUVENCAP") 'P.VTA. B.CAPITAL
    PVTAFAEXP$ = Control("FACTURA", "PUVEFAEX")  'P.VTA. FACT.EXP.

   ' Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Pun.Vta/f10;Descripción/A40;Modo/F0", FRMZELECHO, , , ANTOT)
    Campos$ = "Pun.Vta/f10;Descripción/A60;Modo/F0"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1500
    FRMZELECHO.Caption = "Puntos de Ventas Configurados"
    FRMZELECHO.MSF2.Rows = 8
    FRMZELECHO.MSF2.TextMatrix(1, 1) = PVTANORMAL$
    FRMZELECHO.MSF2.TextMatrix(1, 2) = "Punto de Venta Normal"
    FRMZELECHO.MSF2.TextMatrix(1, 3) = "1"
    
    FRMZELECHO.MSF2.TextMatrix(2, 1) = PVTABCPTAL$
    FRMZELECHO.MSF2.TextMatrix(2, 2) = "Punto de Venta Bienes de Capital"
    FRMZELECHO.MSF2.TextMatrix(2, 3) = "1"
    
    FRMZELECHO.MSF2.TextMatrix(3, 1) = PVTAFAEXP$
    FRMZELECHO.MSF2.TextMatrix(3, 2) = "Punto de Venta Clientes de Exportación"
    FRMZELECHO.MSF2.TextMatrix(3, 3) = "1"
    
    FRMZELECHO.MSF2.TextMatrix(4, 1) = PVTMODOCREDITO$
    FRMZELECHO.MSF2.TextMatrix(4, 2) = "Punto de Venta Factura de Crédito"
    FRMZELECHO.MSF2.TextMatrix(4, 3) = "1"
    
    FRMZELECHO.MSF2.TextMatrix(6, 1) = PVTANORMALPRINT$
    FRMZELECHO.MSF2.TextMatrix(6, 2) = "Punto de Venta Normal (Modo Printer)"
    FRMZELECHO.MSF2.TextMatrix(6, 3) = "2"
    
    FRMZELECHO.MSF2.TextMatrix(7, 1) = PVTAFAEXPRINT$
    FRMZELECHO.MSF2.TextMatrix(7, 2) = "Punto de Venta Bienes de Capital (Modo Printer)"
    FRMZELECHO.MSF2.TextMatrix(7, 3) = "2"
    
    PuntosDeVentaExportaSimple$ = ""
    If PtosVentaExportaSimple& > 0 Then
        FRMZELECHO.MSF2.Rows = 9
        ORDEN& = 0
        
        Dim rstExportaSimple As ADODB.Recordset
        ConsultaSql$ = "select campo1 as PuntoVenta, campo3 as RazonSocial from datasql with(nolock) where id_lista=" & CStr(IDLIST&)
        Set rstExportaSimple = READSET(ConsultaSql$)
        With rstExportaSimple
            Do While Not .EOF
                FRMZELECHO.MSF2.Rows = FRMZELECHO.MSF2.Rows + 1
                ORDEN& = ORDEN& + 1
                FRMZELECHO.MSF2.TextMatrix(7 + ORDEN&, 1) = TRIM$(!PuntoVenta)
                FRMZELECHO.MSF2.TextMatrix(7 + ORDEN&, 2) = Mid$(TRIM$(!razonSocial), 1, 32) & "(Exporta Simple)"
                FRMZELECHO.MSF2.TextMatrix(7 + ORDEN&, 3) = "2"
                
                PuntosDeVentaExportaSimple$ = PuntosDeVentaExportaSimple$ & TRIM$(!PuntoVenta) & ";"
                .MoveNext
            Loop
        End With
        On Error Resume Next
        rstExportaSimple.Close
        Set rstExportaSimple = Nothing
        On Error GoTo 0
    End If
    
    FRMZELECHO.Show 1
    PVTA$ = TRIM$(RESP_ZELE_VECT(1))
    MODX% = XVALO(RESP_ZELE_VECT(3))
    If PVTA$ = "" Then Exit Sub
    '
    If MODX% = 1 Then
'      Call GRACONTROL("FACTURA", "SALIDA", "FELECT")
      Label12.BackColor = &HE0E0E0
    Else
'      Call GRACONTROL("FACTURA", "SALIDA", "PRINTER")
      Label12.BackColor = vbGreen
    End If
    '
    ' MODALIDAD FACTURA DE CREDITO
    chkModoCredito.Value = 0
    If PVTA$ = PVTMODOCREDITO$ Then
        NumeroCliente& = XVALO(Text1)
        If NumeroCliente& = 0 Then
            Call COMUNI("Antes de Elegir la Modalidad Factura de Crédito\Por Favor Ingrese o Elija el Cliente\Sobre el Cual Desea Trabajar.")
            GoTo 99
        Else
            If RegimenFacturaCredito%(NumeroCliente&) = 1 Then
                chkModoCredito.Value = 1
            Else
                Call MENSERR(24, "El Cliente que Ha Elegido\No Está Autorizado Para Emitir Comprobantes\en la Modalidad Factura de Crédito.\ \Requiere ser Configurado\Desde el Maestro de Clientes.")
                GoTo 99
            End If
        End If
    End If
    
    Label12.Caption = PVTA$
    
    GoTo 99
   End If
   
   ' VALIDACION PARA QUE MUESTRE LOS PUNTOS DE VENTA SOLO BIENES CAPITALES O FACT.ELECT.COMUN
   If Control$("FACTURA", "SALIDA") = "FELECT" Then
'     If TRIM$(Label12) = TRIM$(CONTROL("FACTURA", "PUVENELE")) Or TRIM$(Label12) = TRIM$(CONTROL("FACTURA", "PUVENCAP")) Or TRIM$(Label12) = TRIM$(CONTROL("FACTURA", "PUVEFAEX")) Then
        PVTANOR% = XVALO(Control("FACTURA", "PUVENELE"))
        PVTABCP% = XVALO(Control("FACTURA", "PUVENCAP"))
        PVTAFAEX% = XVALO(Control("FACTURA", "PUVEFAEX"))  'P.VTA. FACT.EXP.
        If PVTANOR% > 0 And PVTABCP% > 0 Then
            If TRIM$(Label12) = TRIM$(Control("FACTURA", "PUVENELE")) Then
              Label12 = TRIM$(Control("FACTURA", "PUVENCAP"))
              GoTo 99
            End If
            If TRIM$(Label12) = TRIM$(Control("FACTURA", "PUVENCAP")) Then
              If PVTAFAEX% > 0 Then
               Label12 = TRIM$(Control("FACTURA", "PUVEFAEX"))
               GoTo 99
              Else
               Label12 = TRIM$(Control("FACTURA", "PUVENELE"))
               GoTo 99
              End If
            End If
            
            If TRIM$(Label12) = TRIM$(Control("FACTURA", "PUVEFAEX")) Then
               Label12 = TRIM$(Control("FACTURA", "PUVENELE"))
               GoTo 99
            End If
            
        End If
'     End If
   End If
   
   If Label12 = "0001" Then
       Label12 = "0002"
     ElseIf Label12 = "0002" Then
       Label12 = "0003"
     ElseIf Label12 = "0003" Then
       Label12 = "0004"
     ElseIf Label12 = "0004" Then
       Label12 = "0005"
     ElseIf Label12 = "0005" Then
       Label12 = "0006"
     ElseIf Label12 = "0006" Then
       Label12 = "0007"
     Else
       Label12 = "0001"
   End If
99 CLAVE$ = "PREFAC-" & Label13.Caption '  MODIFICADO PARA GUARDAR PREFIJO DE FACTURA POR ESTACION
   If InStr(1, PuntosDeVentaExportaSimple$, Label12) = 0 Then
        Call GRACONTROL(CLAVE$, IDENTER$, Label12)
   End If
   '
End Sub

Private Sub Label12_Change()
   '
   MsgBox "Entro en label12_change)"
   VRFF% = VERIFSALIDA%("NOMESS")
   Label12.BackColor = &HE0E0E0
   If VRFF% = 1 Then
        Label12.ToolTipText = "Fact.Elect. - Productos / Servicios Comunes - Doble Click para Cambiar"
     ElseIf VRFF% = 2 Then
        Label12.ToolTipText = "Fact.Elect. - Régimen Bienes de Capital - Doble Click para Cambiar"
     ElseIf VRFF% = 3 Then
        Label12.ToolTipText = "Fact.Elect. - Tipo Exportación - Doble Click para Cambiar"
   End If
   '
   PVTANORMALPRINT$ = Control("FACTURA", "PUVENPRI") 'P VTA. NORMAL MODO PRINTER
   PVTAFAEXPRINT$ = Control("FACTURA", "PUVEXPRI")  'P.VTA. FACT.EXP. MODO PRINTER   ' esto agrego EPB el 31/12/10 para manejar la factura electronica con una sola
   If XVALO(PVTANORMALPRINT$) = XVALO(Label12) Then
      Label12.BackColor = vbGreen
      Label12.ToolTipText = "Fact.Común. - Productos / Servicios Comunes - Doble Click para Cambiar"
   ElseIf XVALO(PVTAFAEXPRINT$) = XVALO(Label12) Then
      Label12.BackColor = vbGreen
      Label12.ToolTipText = "Fact.Común. - Tipo Exportación - Doble Click para Cambiar"
   End If
   
   ' bandera FACELEC% = 1 --> FE normal; = 2 --> FE Bienes de Capital; = 3 --> Exportacion
   FACELEC% = VRFF%
   '
   ' *** ACTIVA LA MODALIDAD EXPORTA SIMPLE
   If FACELEC% = 0 Then
        ModoExportaSimple% = 0
        If Control$("FACTURA", "EXSIMPLE") = "SI" Then
            If CantRegistros("DATASQL", "lista='PTOVTAEXSIM' and campo1 like '%" & Label12 & "%'") > 0 Then
                ModoExportaSimple% = 1
            End If
        End If
   End If
   '
TXTRAZA$ = TXTRAZA$ & "PUNTO 203.1 = " & Now & vbCrLf
 MsgBox "antes de nufactura)"
   Call NUFACTURA
 MsgBox "despues de nufactura)"
   ' *** VERIFICA QUE SE HAYAN COMPLETADO LOS DATOS NECESARIOS PARA REGISTRAR
   ' *** LA FACTURA DE EXPORTA SIMPLE
   If ModoExportaSimple% = 1 And (XVALO(PuntoVentaExportaSimple$) = 0 Or XVALO(NumeroExportaSimple$) = 0 Or XVALO(CAEExportaSimple$) = 0) Then
        Call BLANFORMU
        Text1.SetFocus
   End If
TXTRAZA$ = TXTRAZA$ & "PUNTO 203.2 = " & Now & vbCrLf
   
   'Call Controla_NroRemito(TRIM$(Label12))
   '
End Sub

Private Sub Label14_DblClick()
   Call Text14_DBLCLICK
End Sub

Private Sub Label28_DblClick()
   '
   If Label28.Caption = "" Then
       Label28.Caption = "A"
     ElseIf Label28.Caption = "A" Then
       Label28.Caption = "B"
     ElseIf Label28.Caption = "B" Then
       Label28.Caption = "C"
     ElseIf Label28.Caption = "C" Then
       Label28.Caption = "D"
     ElseIf Label28.Caption = "D" Then
       Label28.Caption = "A"
   End If
   '
End Sub

Private Sub Label7_Change()
   BASIVA(0).BackColor = &HE0E0E0
   BASIVA(1).BackColor = &HE0E0E0
   Total(3).BackColor = &HE0E0E0
   Total(4).BackColor = &HE0E0E0
   If (TIFADOC% > 2) Or (TIFADOC% = 2 And MODODOC% = 1) Then ' DEBITOS Y CRÉDITOS Y FACTURAS DE TEXTO LIBRE
      BASIVA(0).BackColor = &H80000005
      BASIVA(1).BackColor = &H80000005
      Total(3).BackColor = &H80000005
      Total(4).BackColor = &H80000005
   End If
End Sub

Private Sub LENTREGA_GotFocus(Index As Integer)
   If Index = 2 Then
     If XVALO(Text19) > 0 Then
       Call MENSERR(24, "Campo no Editable")
       On Error Resume Next
       Text19.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub


Private Sub mnuAnulacionComprobantes_Click()
    Call OPVARFA07.Command45_Click
End Sub

Private Sub mnuCodigosEquivalentesClientes_Click()
    Call OPFACTU07.Command10_Click
End Sub

Private Sub mnuCompensacionClienteProveedor_Click()
    
    Mensaje$ = "Esta Funcionalidad Permite Realizar Compensaciones entre Cliente y Proveedor."
    Mensaje$ = Mensaje$ & "\La misma se lleva a cabo Convirtiendo los Débitos Emitidos al Cliente\en Créditos para la cuenta del Proveedor Asociado."
    Mensaje$ = Mensaje$ & "\(ver Tabla de Compensaciones Cliente-Proveedor)"
    Mensaje$ = Mensaje$ & "\ \El Usuario debe Elegir las Notas de Débito utilizadas para Compensar\y el Sistema se encargará del Proceso de Conversión."
    Call COMUNI(Mensaje$)
    
    Dim Tablas As New DatosAuxiliares
    Call Tablas.ListaSeleccionTablaAuxiliar("COMPECLIPRO", "Tabla Compensación Cliente-Proveedor")
    
    CuentaCliente& = XVALO(RESP_ZELE_VECT(2))
    CuentaProveedor& = XVALO(RESP_ZELE_VECT(4))
    If CuentaCliente& > 0 Then
        ConsultaSql$ = "select fechconta, codicom_lar, valabscomp from cajabanc with(nolock) where tipomovim='fven' and opcimovim='nd' and numecomp>0 and nuclien=" & CStr(CuentaCliente&) & " and modoopcimovim<>'COMPENSA' order by fechconta"
        Set Rst = READSET(ConsultaSql$)
        Campos$ = "Fecha/D8;Comprobante/A24;Importe/F14.2"
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
        Call Carga_MSF_Recordset(Rst, Campos$, FRMZELECHO.MSF2)
        FRMZELECHO.Caption = "Notas de Débito"
        FRMZELECHO.Show 1
        
        FechaComprobante% = FECHANUM(CStr(RESP_ZELE_VECT(1)))
        Comprobante$ = TRIM$(RESP_ZELE_VECT(2))
        If Comprobante$ <> "" Then
            On Error GoTo ControlError
            FLEXCONN.BeginTrans
            
            ' REPLICA LOS REGISTROS DE LA NOTA DE DEBITO
            ConsultaSql$ = "insert into cajabanc (codiempr,fehoreal,nuasidiar,codicom_lar,codicom_cor,refecom,fechconta,fechacred,cuecontai,cencosto,itemconc,nuprove,tipomovim,opcimovim,varimovim,descrimov,fechemisi,codisucu,numetalo,numecomp,status,mone_emis,tipo_cam,impo_dola,valabscomp,idebecomp,ihabecomp,baseimpo,incsubiva,nrocai,fevtocai,obsergen,atracheque,atracance,condpag,niv_clasi,letracomis,vend_cli,cae,idcae,codcarpimp,textocae,notas,incoterms,ocompra,pedidos,cli_real,fechvtocae,simmone,cod_barra,descu1,descu2,impodesc1,impodesc2,idcaenum,cotiza_dolar,codigoimpuesto,codigoregimen,cuit,despaimpo,pos_iva,cuit_cli,nroderetencion,monedacotiza_dolar,usernum,modoopcimovim,cbualias,codicom_lar_exsimple)" _
            & " select codiempr,fehoreal,nuasidiar,codicom_lar,codicom_cor,refecom,fechconta,fechacred,cuecontai,cencosto,itemconc,nuprove,tipomovim,opcimovim,varimovim,descrimov,fechemisi,codisucu,numetalo,numecomp,status,mone_emis,tipo_cam,impo_dola,valabscomp,idebecomp,ihabecomp,baseimpo,incsubiva,nrocai,fevtocai,obsergen,atracheque,atracance,condpag,niv_clasi,letracomis,vend_cli,cae,idcae,codcarpimp,textocae,notas,incoterms,ocompra,pedidos,cli_real,fechvtocae,simmone,cod_barra,descu1,descu2,impodesc1,impodesc2,idcaenum,cotiza_dolar,codigoimpuesto,codigoregimen,cuit,despaimpo,pos_iva,cuit_cli,nroderetencion,monedacotiza_dolar,usernum,modoopcimovim,cbualias,codicom_lar_exsimple from cajabanc where codicom_lar='" & Comprobante$ & "' and nuclien=" & CStr(CuentaCliente&)
            FLEXCONN.Execute ConsultaSql$
            
            '
            Dim rstConversion As New ADODB.Recordset
            ConsultaSql$ = "select * from cajabanc where codicom_lar='" & Comprobante$ & "' and nuclien is null and nuprove=0 and fechconta='" & FETEXCOR1$(FechaComprobante%) & "'"
            rstConversion.Open ConsultaSql$, FLEXCONN, adOpenKeyset, adLockPessimistic
            
            With rstConversion
                CbteLargoConvertido$ = "DP-" & Mid$(Comprobante$, 18, 1) & "-" & Mid$(Comprobante$, 4, 13)
                CbteCortoConvertido$ = "DP-" & Mid$(Comprobante$, 11, 6) & "-" & Mid$(Comprobante$, 18, 1)
                Do While Not .EOF
                    TipoMovimiento$ = !TIPOMOVIM
                    Select Case TipoMovimiento$
                        Case "FVEN": !TIPOMOVIM = "FCOM": !OpciMovim = "DP": !CUECONTAI = CUEINT%(CUEMADRE$((-1) * CuentaProveedor&))
                        
                        Case "VENTA": !TIPOMOVIM = "COMPRA": !OpciMovim = "DP"
                        
                        Case "IVA":
                                    !OpciMovim = "CrFi"
                                    If !VARIMOVIM = "PER-IVA" Then
                                        CuentaContable% = CompensaClienteProveedorImpuestos("PER-IVA")
                                    Else
                                        TASA$ = Mid$(!VARIMOVIM, 1, 5)
                                        CuentaContable% = 0
                                        For i = 0 To 3
                                            If TRIM$(Control$("", "TASIVA-" & CStr(i))) = TASA$ Then
                                                CuentaContable% = CUEINT%(Control$("", "CUEIVA-" & CStr(i)))
                                                Exit For
                                            End If
                                        Next i
                                    End If
                                    !CUECONTAI = CuentaContable%
                                    
                        Case "IBRU":
                                    Select Case !OpciMovim
                                        Case "IB-B":
                                                    CuentaContable% = CompensaClienteProveedorImpuestos("IB-BSAS")
                                                    !CUECONTAI = CuentaContable%
                                        Case "IB-C": CuentaContable% = CompensaClienteProveedorImpuestos("IB-CABA")
                                                    !CUECONTAI = CuentaContable%
                                        Case "IB-MI": CuentaContable% = CompensaClienteProveedorImpuestos("IB-MISIONES")
                                                    !CUECONTAI = CuentaContable%
                                        Case "IB-TU": CuentaContable% = CompensaClienteProveedorImpuestos("IB-TUCUMAN")
                                                    !CUECONTAI = CuentaContable%
                                    End Select
                    End Select
                    !NUCLIEN = 0
                    !Cli_Real = 0
                    !NUPROVE = CuentaProveedor&
                    !codicom_lar = CbteLargoConvertido$
                    !CodiCom_Cor = CbteCortoConvertido$
                    !notas = Comprobante$
                    
                    .Update
                    
                    .MoveNext
                Loop
            End With
            
            ConsultaSql$ = "update cajabanc set idebecomp=0,ihabecomp=0,notas='" & CbteLargoConvertido$ & "',incsubiva='No',modoopcimovim='COMPENSA' where codicom_lar='" & Comprobante$ & "' and nuclien=" & CStr(CuentaCliente&) & " and fechconta='" & FETEXCOR1$(FechaComprobante%) & "'"
            FLEXCONN.Execute ConsultaSql$
            
            FLEXCONN.CommitTrans
            On Error GoTo 0
            
            Call COMUNI("Compensación Procesada Correctamente")
        End If
    End If
    
    Exit Sub
    
ControlError:
    On Error Resume Next
    FLEXCONN.RollbackTrans
    On Error GoTo 0
    
End Sub

Private Function CompensaClienteProveedorImpuestos(Impuesto As String) As Integer
    
    Dim TABLA As New DatosAuxiliares
    
    CuentaContable% = -1
    IdTabla& = TABLA.ObtenerIdTablaAuxiliar("COMCLPRIMPU")
    If CantRegistros("datasql", "id_lista=" & CStr(IdTabla&) & " and campo2='" & Impuesto & "'") > 0 Then
        CuentaContable% = CUEINT%(VALOBADA("datasql", "campo3", "id_lista=" & CStr(IdTabla&) & " and campo2='" & Impuesto & "'"))
    End If
    
    If CuentaContable% < 0 Then
        Call COMUNI("No se Pudo Compensar el Impuesto '" & Impuesto & "'.\Se Mantuvo la Contabilización Inicial del Débito.")
    End If
    
    CompensaClienteProveedorImpuestos = CuentaContable%
    
End Function

Private Sub mnuCondicionesDePago_Click()
    Call CARCONPAGOS
End Sub

Private Sub mnuConfiguracionGeneral_Click()
    Call Command5_Click
End Sub

Private Sub mnuControlGeneral_Click()
    REMEVE07.Show 1
End Sub

Private Sub mnuCorregirFechaCbte_Click()
    Call OPVARFA07.Command3_Click
End Sub

Private Sub mnuCorregirNumeroCbte_Click()
    Call OPVARFA07.Command11_Click
End Sub

Private Sub mnuCotizacionesMoneda_Click()
    Call OPFACTU07.Command5_Click
End Sub

Private Sub mnuEliminaCbtesSinCae_Click()
    Call OPVARFA07.Command9_Click
End Sub

Private Sub mnuFActProfMatCod_Click()
   If Control("FACTURA", "PROFORTL") <> "SI" Then
    If CLACONTRA% > 0 Then
      Call GRACONTROL("FACTURA", "PROFORTL", "SI")
      'SOLO SE UTILIZA PARA QUE UNA VEZ DADO EL CLACONTRA NO LO VUELVA A PEDIR
    Else
      Exit Sub
    End If
   End If

   Call REGISTRAPROFORMA_MATCOD
   Call BLANFORMU
End Sub

Private Sub mnuListasDePrecios_Click()
    Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
End Sub

Private Sub mnuMaestroArticulos_Click()
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG07/AMASTO", "Archivos Maestros")
End Sub

Private Sub mnuMaestroClientes_Click()
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG07/AMACLI", "Archivos Maestros")
End Sub

Private Sub mnuMotivosImputacion_Click()
    Call OPFACTU07.Command12_Click
End Sub

Private Sub mnuObtenerToken_Click()
    Call OPVARFA07.Command5_Click
End Sub

Private Sub mnuParametrosVarios_Click()
    Call OPFACTU07.Command8_Click
End Sub

Private Sub mnuPedidosPendientes_Click()
    Call Command17_Click
End Sub

Private Sub mnuRemitosPendientesFacturacion_Click()
    Call Command8_Click
End Sub

Private Sub mnuSolicitudManualCae_Click()
    Call OPVARFA07.Command7_Click
End Sub

Private Sub mnuTablaCompensacionClienteProveedor_Click()

    Dim Tablas As New DatosAuxiliares
    Call Tablas.CrearTablasAuxiliares
    
    IdTabla& = Tablas.ObtenerIdTablaAuxiliar("COMPECLIPRO")
    Call ABM_TABLA(IdTabla&, "Comp. Cliente-Proveedor")
    
End Sub

Private Sub mnuTablaCompensacionImpuestos_Click()

    Dim Tablas As New DatosAuxiliares
    Call Tablas.CrearTablasAuxiliares
    
    IdTabla& = Tablas.ObtenerIdTablaAuxiliar("COMCLPRIMPU")
    If CantRegistros("datasql", "id_lista=" & CStr(IdTabla&)) = 0 Then
        ConsultaSql$ = "insert into datasql (campo1,campo2,id_lista) values ('1','IB-CABA', '" & CStr(IdTabla&) & "')"
        FLEXCONN.Execute ConsultaSql$
        ConsultaSql$ = "insert into datasql (campo1,campo2,id_lista) values ('2','IB-BSAS', '" & CStr(IdTabla&) & "')"
        FLEXCONN.Execute ConsultaSql$
        ConsultaSql$ = "insert into datasql (campo1,campo2,id_lista) values ('3','IB-MISIONES', '" & CStr(IdTabla&) & "')"
        FLEXCONN.Execute ConsultaSql$
        ConsultaSql$ = "insert into datasql (campo1,campo2,id_lista) values ('4','IB-TUCUMAN', '" & CStr(IdTabla&) & "')"
        FLEXCONN.Execute ConsultaSql$
        ConsultaSql$ = "insert into datasql (campo1,campo2,id_lista) values ('5','PER-IVA', '" & CStr(IdTabla&) & "')"
        FLEXCONN.Execute ConsultaSql$
    End If

    Call ABM_TABLA(IdTabla&, "Comp. de Impuestos")
    
End Sub

Private Sub mnuTasasIva_Click()
    Call CARTASIVA
    Call COMUNI("Re-Inicie Ahora la Aplicación.")
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub



Private Sub MONEMIS_Change()
   '
   '\\\OT-07-0086-RG-06/03/07///
   ENTRATICA.Label6.Caption = MONEMIS.Caption
   ENTRATICA.Text4 = TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4"))
   '\\\OT-07-0086-RG-FIN///
   '
   'PARA QUE NO QUEDE EN EL FOCO DE PERCEPCION DE INGRESOS BRUTOS SI FUE SU ULTIMA POSICION
   'YA QUE EN EL EVENTO GOSTFOCUS TIENE UNA VALIDACION - DE ESTE MODO SE VUELVE A EJECUTAR LA MISMA.
   If InStr(EMPREAC$, "MEDITEA") > 0 Then
      On Error Resume Next
      PORDESCU.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub MSF_DblClick()
   If Command14.Enabled = True Then
   
     If FACLOT07.Label15 = "ACTIVO" Then
        FACLOT07.Show 1  'CARGA POR BOBINA
     Else
        Call CARDETFAC
     End If
   End If

End Sub


Private Sub PORIVA_DblClick(Index As Integer)
   If Index <= 1 Then
     '
     SUBASIM# = XVALO(BASIVA(0)) + XVALO(BASIVA(1))
     TEBASIM$ = TRIM$(FORMATNUM$(SUBASIM#, "F12.2"))
     If Index = 0 Then
         BASIVA(0) = TEBASIM$
         BASIVA(1) = ""
         
       Else
         BASIVA(0) = ""
         BASIVA(1) = TEBASIM$
     End If
     ALIVAGEN = XVALO(PORIVA(Index))
     '
     Total(4).TEXT = CSTRING$(MKD$(XVALO(BASIVA(1)) * XVALO(PORIVA(1)) * APLIVA% / 100), 3, 12, 2, 2)
     Total(3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(0)) * XVALO(PORIVA(0)) * APLIVA% / 100), 3, 12, 2, 2)
     '
     TTTX# = (XVALO(Total(3).TEXT) + XVALO(Total(4).TEXT)) * RNI%
     BASIVA(2) = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
     Total(5).TEXT = CSTRING$(MKD$(XVALO(BASIVA(2)) * XVALO(PORIVA(2).Caption) / 100), 3, 12, 2, 2)
     '
     TTTX# = 0: For KKI% = 2 To 5
                    TTTX# = TTTX# + XVALO(Total(KKI%).TEXT)
                 Next KKI%
                 TTTX# = TTTX# + XVALO(Text15)
                 TTTX# = TTTX# + XVALO(Text17)
                 TTTX# = TTTX# + XVALO(Text24)
     Total(6).TEXT = CSTRING$(MKD$(TTTX#), 3, 12, 2, 2)
     '
     For KU& = 1 To ULTREG&("!CARDEFAC")
       XXX$ = REGLEIDO$("!CARDEFAC", KU&)
       INETO1# = CVD(Mid$(XXX$, 433, 8))
       Call REPLA(XXX$, MKS$(ALIVAGEN), 441, 4)
       Call REPLA(XXX$, MKD$(INETO1# * ALIVAGEN / 100), 445, 8)
       Call REPLA(XXX$, MKD$(INETO1# * (1 + ALIVAGEN / 100)), 453, 8)
       Call GRAREG("!CARDEFAC", XXX$, KU&)
     Next KU&
     Call CIERRARCH("!CARDEFAC")
     '
   End If
End Sub

Private Sub Text14_DBLCLICK()
    '
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
    If TIFADOC% <> 2 Then
      ' solo para facturas
      Exit Sub
    End If
    '
    If MODODOC% <> 0 Then  ' 0=MAT.CODIFICADO / 1=TEXTO LIBRE
       ' solo mat codificado
       Exit Sub
    End If
    '
    If PREFACTU$ <> "REMITO" Then
      ' solo para forma de guidi
      Exit Sub
    End If
    '
    'ACA PONGO EL INDICE
    NCLIE% = XVALO(Text1)
    Screen.MousePointer = 11
    '
    Call APERBASDAT("STOCKS")
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT DOSECLAR,  NUPEDCLI, COD_INTE FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE% & " "
    SQLX$ = SQLX$ + " AND (CodiMovi = 'RT' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND (DoSECLar = '' "
    SQLX$ = SQLX$ + " OR DoSecLar LIKE 'P/F%' ) "
    SQLX$ = SQLX$ + " ORDER BY NUPEDCLI ASC "
    On Error GoTo 0
    'Set pedianuittemp = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
    Dim pedianuittemp As ADODB.Recordset
    Set pedianuittemp = New ADODB.Recordset
25  On Error Resume Next
    pedianuittemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0

    '
    JJ& = 0
    With pedianuittemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        TTXX$ = REGBLAN("INDIFAPE")
        Do While Not .EOF
         If PEDANTI& <> !NUPEDCLI Then
            NRODES& = !NUPEDCLI
            CI% = !cod_inte
            If NRODES& <= 0 Then GoTo 15
          TTXX$ = MKS$(NRODES&)
            NOCO$ = NOCOMCLI(NRODES&)
          If NOCO$ <> "" Then
            TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
          End If
            'With PedEnca
            '.FindFirst "NroPed = " & NRODES& & ""
            SQLX$ = " SELECT * FROM PedEnca"
            SQLX$ = SQLX$ + " WHERE NroPed =" & NRODES& & ""
            Dim PEDDDENCA As ADODB.Recordset
            Set PEDDDENCA = FLEXCONN.Execute(FILTSQL$(SQLX$))
            On Error Resume Next
            With PEDDDENCA
              'If .NoMatch = False Then
              If Err < 1 Then
                On Error GoTo 0
                FEX = CVINT(PedEnca!FECHEMIS)
              End If
              On Error GoTo 0
            End With
            Set PEDDDENCA = Nothing
          TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
            NCLIE% = !NROCLIE
          TTXX$ = TTXX$ + TRIM$(rasocli$(NCLIE%)) + Space$(52)
          JJ& = JJ& + 1
          Call GRAREG("INDIFAPE", Left$(TTXX$, 52), JJ&)
          PEDANTI& = NRODES&
         End If
15      .MoveNext
      Loop
      End If
      On Error GoTo 0
    End With
    Set pedianuittemp = Nothing
    Call SETULTREG("INDIFAPE", JJ&)
    Call CIERRARCH("INDIFAPE")
    Call FRESHECHO("INDIFAPE")
    Screen.MousePointer = 1
    '
    If ULTREG&("INDIFAPE") < 1 Then
       Call COMUNI("No hay Remitos Pendientes de este Pedido\que Puedan Facturarse")
       Call CIERRARCH("*.*")
       Exit Sub
    End If
    '
22  X$ = REGSEL$("INDIFAPE/Pedidos Pendientes para Facturación")
    If Len(X$) <= 4 Then GoTo 999
    '
    NUPEDIDO1& = CVS(Left$(X$, 4))
    Text14 = NOCOMCLI$(NUPEDIDO1&)
    CONSICLI% = 0
    Call CARCONSIPED(NUPEDIDO1&)
    If CONSICLI% > 0 Then
      Call CARDETFAC
      MARCONOTA.Visible = False
      Exit Sub
    End If
    '
999  End Sub

Private Sub Text15_Change()
   Call CALTOFA2
End Sub



Private Sub Text24_Change()
   Call CALTOFA2
End Sub

Private Sub Text15_GotFocus()
   '
   If VALICLIE% < 1 Then GoTo 99
   'SOLICITUD DE MEDITEA PARA PODER INGRESAR UN IMPORTE DE PERCEPCION DE MODO MANUAL
   'SI SE CUMPLEN ESTAS CONDICIONES TOMA EL FOCO Y SE PUEDE EDITAR EL CASILLERO, CASO CONTRARIO SIGUE  Y PIERDE EL FOCO
   If InStr(EMPREAC$, "MEDITEA") > 0 Then
      If (Label13 = "A" Or Label13 = "B") And TRIM$(UCase$(MONEMIS.Caption)) = "PESOS" Then
        Exit Sub
      End If
   End If
   '
   If TIFADOC% < 3 Then GoTo 98
   'If XVALO(Total(2)) < 0.5 Then GoTo 98
   Exit Sub
   '
98 On Error Resume Next
   PORDESCU.SetFocus
   On Error GoTo 0
   Exit Sub
   '
99 On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text24_GotFocus()
   '
   If VALICLIE% < 1 Then GoTo 99
   '
   If TIFADOC% < 3 Then GoTo 98
   Exit Sub
   '
98 On Error Resume Next
   PORDESCU.SetFocus
   On Error GoTo 0
   Exit Sub
   '
99 On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Text15 = TRIM$(FORMATNUM$(XVALO(Text15), "F8.2"))
  End If
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Text24 = TRIM$(FORMATNUM$(XVALO(Text24), "F8.2"))
  End If
End Sub

Private Sub Text15_LostFocus()
   Text15 = TRIM$(FORMATNUM$(XVALO(Text15), "F8.2"))
End Sub

Private Sub Text24_LostFocus()
   Text24 = TRIM$(FORMATNUM$(XVALO(Text24), "F8.2"))
End Sub

Private Sub Command15_Click()
   MARCONOTA.Visible = False
   Text1.TEXT = ""
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command16_Click()
   Call MENSERR(24, "No Disponible por el Momento")
End Sub

Private Sub Command16_LostFocus()
   On Error Resume Next
   Command6.SetFocus
   On Error GoTo 0
End Sub

Sub Command17_Click()
   Command17.Enabled = False
       'EPAC$ = TRIM$(UCase$(EMPREAC$))
   CONPED07.Show 1
       'Call SHOWPEDIPEN
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command19_Click()
   '
   NUOCOMPRA$ = ""
   NROPEDIDO$ = ""
   '
   Call CARDETFAC
   '
End Sub

Private Sub Command2_Click()
    If EsPorPedido& > 0 Then
       Call FINAL("")
       Exit Sub
    End If
    REDOCEL.MODORECU = 0
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_LostFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Sub Command20_Click()
   Command20.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "GUIDI") > 0 Then
      OPP% = ALTERNA%("Repetir Impresion de Factura\Otros Complementos de Facturación")
      If OPP% = 1 Then
         Call SELFACREPI
         GoTo 99
      End If
   End If
   Screen.MousePointer = 1
   OPVARFA07.Show 1
99 Command20.Enabled = True
End Sub

Sub SELFACREPI()
    '
    If DERACCE%("REPRIFAC", "Repetir Impresión de Facturas") < 2 Then
      Exit Sub
    End If
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    SQLX$ = SQLX$ + "Order by FECHEMISI DESC, CODICOM_LAR DESC "
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLIE% = !NUCLIEN
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            If TRIM$(IMPOS$) = "" Then IMPOS$ = AJUSTD$("0.00", 14)
            FEMISI$ = FECHATEX$(CVINT(!FECHEMISI))
            NPX& = XVALO(Mid$(!codicom_lar, 11, 8))
            If Left$(!codicom_lar, 2) = "CC" Then
              NPX& = XVALO(Mid$(!codicom_lar, 4))
            End If
            Y$ = MKS$(NPX&) + AJUSTI$(!codicom_lar, 18) + AJUSTD$(FEMISI$, 14) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUERE", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
 
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no tiene Facturas de Venta\que Puedan Re-Imprimirse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    Call SELECHO("!SELCUERE")
    If VALACT1$("!SELCUERE") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUERE"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    REPRIFAC% = 1
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & DOCULA$ & "' "
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    Text1 = TRIM$(Str$(NC%))
    Call CARDACLI
    Call Label12_Change
    DOCUCO$ = CABATEMP!CodiCom_Cor
    Text9 = FECHATEX$(CVINT(CABATEMP!FECHEMISI))
    Label12 = TRIM$(Mid$(DOCULA$, 4, 4))
    Label13 = TRIM$(Mid$(DOCULA$, 18, 1))
    Text11 = TRIM$(Str$(XVALO(Mid$(DOCULA$, 9, 8))))
    '
    MARCONOTA.Visible = False
    Call COMUNI("ADVERTENCIA: Los datos centrales de facturación\tales como lista de precios, condicion de pago\sucursal de entrega, etc son tomados de\las BASES DE DATOS CENTRALES.\   \Recomendamos revisar visualmente si corresponden\con los de la factura y eventualmente\ajustarlos manualmente.")
    '
    Call CARDETAFAC(NC%, DOCULA$)
    Call CARDETFAC
    '
End Sub
'
Sub CARDETAFAC(NCLIE%, DOCULA$)
   '
   REGODES& = 0
   Call SETULTREG("!CARDEFAC", 0)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
   End If
   JJ& = 0: CONSICLI% = 0
   '
   Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   '
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE% & " "
   SQLX$ = SQLX$ + " AND DoSecLar = '" + DOCULA$ + "' "
   
   Dim CONSITMP As ADODB.Recordset
   Set CONSITMP = New ADODB.Recordset
   CONSITMP.CursorLocation = adUseClient
25 On Error Resume Next
   CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With CONSITMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       RECONSI$ = REGBLAN("CONSICLI")
       J& = 0
       Do While .EOF = False
         CODIMAT% = XVALO(!cod_inte)
         CAN = XVALO(!CANTSALID) - XVALO(!CANTINGRE)
         NPEDID& = XVALO(!NUPEDCLI)
         NUMEPED& = 0: PREPED = 0
           If NPEDID& < 1 Then
             FEREMIT% = CVINT%(!FechMov)
             TITUCON$ = TRIM$(SAFETEXT$(!DoPriCor)) + " - Cant: " + TRIM$(Str$(CAN)) + " - Cod: " + TRIM$(CODEXT$(CODIMAT%)) + " - " + DESCRIT$(CODIMAT%)
             Call RECARPED(NCLIE%, CODIMAT%, FEREMIT%, NUMEPED&, PREPED, TITUCON$)
             NPEDID& = NUMEPED&
           End If
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = SAFETEXT$(!DoPriCor)
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = CVINT%(!FechMov)
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(Z$, NUPLANO$(CODIMAT%), 315, 24)
           '
7          MONEMA% = XVALO(!MoneVal)
           MONEMI% = MONEMA%
           PRECIO = XVALO(!VALOUNIT)
             '
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
         Call REPLA(Z$, " ", 461, 1)    ' NUNCA ES REPARACION
           NORIT% = XVALO(!NuOrItem)
         Call REPLA(Z$, MKI$(NORIT%), 503, 2)
         Call REPLA(Z$, MKS$(0), 505, 4)
         Call REPLA(Z$, MKS$(0), 509, 4)
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(Z$, NUOCOX$, 355, 16)
         '
         Call REGAPP("!CARDEFAC", Z$)
         CAITCAR% = CAITCAR% + 1
         '
       .MoveNext
       Loop
12   End If
   End With
   CONSITMP.Close
   Set CONSITMP = Nothing
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     Z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(Z$, 1, 4))) >= 0.01 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", Z$, JJ&)
     End If
   Next U&
   Call SETULTREG("!CARDEFAC", JJ&)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   Call CIERRARCH("!CARDEFAC")
   '
End Sub
'
Private Sub Command22_Click()
    Command22.Enabled = False
     Call MENSERR(24, "No Disponible.\Debe Emitir Nota de Débito Propia.")
    Command22.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("FACTURA")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEC07.Text1.TEXT = Text1.TEXT
    RECUEC07.Show 1
    Command4.Enabled = True
End Sub

Sub Command5_Click()
    Command5.Enabled = False
    If DERACCE%("STPFACTU", "Configuración de Funcionamiento Facturacion") > 1 Then OPFACTU07.Show 1
    Call FINAL("")
    Command5.Enabled = True
End Sub

Sub Command6_Click()
   '
   Static AGRUPAENTREGA% 'ASIGNA A AGRUPAENTREGA% SEGUN LA CONFIGURACION DEL SETUP
   If AGRUPAENTREGA% = 0 Then
      AGRUPAENTREGA% = (-1)
      If Control$("FACTUREM", "AGRUPEDI") = "TRANSUCU" Then AGRUPAENTREGA% = 1
   End If
   '
   MARCONOTA.Visible = False
   DOCUDIF = ""
   '
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   Call FILEKILL("!CARDEFAC")
   '
   If EsPorPedido& > 0 Then
      TIFADOC% = 2
      MODODOC% = 0
   End If
   '
   FACLOT07.Label15 = "" ' MARCA COMO NO ACTIVO MODADLIDAD DE CARGA POR BOBINAS
   If TIFADOC% = 3 Then   ' nota de debito
     If MODODOC% = 0 Then ' mat codificado
       DEBIMANUAL% = 0
       OPX% = COMALTER%("Opciones de Débito\\Ingreso Manual\Diferencia de Precios")
       If OPX% < 1 Or OPX% > 2 Then Exit Sub
       If OPX% = 2 Then
            Call DIFPREC
            Exit Sub
       End If
       DEBIMANUAL% = 1
     End If
   End If
   '
   CONSICLI% = 0
   If TIFADOC% = 2 Then    ' 2=FACTURA / 3=N.DEBITO / 6=N.CREDITO
     If MODODOC% = 0 Then  ' 0=MAT.CODIFICADO / 1=TEXTO LIBRE
       '
       If InStr(1, EMPREAC$, "GABAL") > 0 Then
            If LISTADEMPAQ%(XVALO(Text1)) > 0 Then ' 1 SI HAY LISTA DE EMPAQUE
                  OPX% = COMALTER%("Hay Listas de Empaque Pendientes\\Facturar Ahora\Continuar")
                  If OPX% = 1 Then
                     Call SELEMPAQUE
                     Call CARDETFAC
                     Exit Sub
                  End If
            End If
       End If
       '
       ESPROFORMA& = 0
       If TieneProformaPendienteDeFacturar%(XVALO(Text1)) > 0 Then
          If COMALTER%("Hay Facturas Proformas\Pendiente para este Cliente\\Facturar Factura Proforma\Continuar Normalmente") = 1 Then
            NUMEPROFORMA& = SELECCION_FACTURAPROFORMA%(XVALO(Text1))
            If NUMEPROFORMA& > 0 Then
              Call CARGADETALLEPROFORMA(NUMEPROFORMA&)
              ESPROFORMA& = NUMEPROFORMA&
              Call CARDETFAC
              Exit Sub
            End If
          End If
       End If

       '
       Call CARCONSIPEN
       If CONSICLI% > 0 Then
          Call CARDETFAC
          NUOCOMPRA$ = TRIM$(NUOCOMPRA$)
          If InStr(NUOCOMPRA$, ",") > 0 Then NUOCOMPRA$ = TRIM$(Left$(NUOCOMPRA$, Len(NUOCOMPRA$) - 1))
          Text14 = NUOCOMPRA$
          Exit Sub
       End If
       If PREFACTU$ = "REMITO" Then ' POR REMITO Y NO HAY CONSIGNACION PENDIENTE
          Screen.MousePointer = 1
          Text1 = ""
          Call BLANFORMU
          On Error Resume Next
          Text1.SetFocus
          On Error GoTo 0
          Exit Sub
       End If
       ' SELECCIONAR PEDIDOS PENDIENTES
       Screen.MousePointer = 1
       If PREFACTU$ = "ORDESPA" Then
         'If InStr(EPAC$, "DEL CIELO") < 1 Then
         If Control("FACTURA", "SISTMANT") = "SI" Then
           Call SELPEDIP1 'sistema de mantenimiento
         Else
           Call SELPEDIP
         End If
'          'VALIDACION PARA MEDITEA (NO HAYA PEDIDOS CON DIF.MOTIVOS Y QUE ESTEN AUTORIZADOS)
'          If VERIPEDMO% < 1 Then GoTo 10
       Else
           Call CARDETPEDIPEN   ' CARGA DETALLE DE PEDIDOS PENDIENTES
       End If
       '
       MONEM1% = MONEMI%: Descu1$ = "": TRANSPO1% = 0
       MALCOMBINA% = 0: HAYPESEL% = 0
       
       If VERIPEDMO% < 1 Then MALCOMBINA% = 1   ' DISTINTO MOTIVO DE PEDIDO
       '
40     If AGRUPAENTREGA% > 0 Then  ' VALIDA SI ESTA SELECCIONADA LA OPCION DE AGRUPACION DEL SETUP Y PRESENTA WARNING
          For KKU& = 1 To ULTREG&("!CARDEFAC")
            XXU$ = REGLEIDO$("!CARDEFAC", KKU&)
            NPEDXX& = XVALO(TRIM$(Mid$(XXU$, 371, 16)))
            If NPEDXX& > 0 Then HAYPESEL% = 1
            Descu2$ = TRIM$(DESCUPEDI$(NPEDXX&)): If Descu2$ = "" Then Descu2$ = "0"
            If Descu1$ = "" Then Descu1$ = Descu2$
            If Descu2$ <> Descu1$ Then
                MALCOMBINA% = 1
            End If
            MONEM2% = MONEPEDI%(NPEDXX&)
            If MONEM1% < 1 Then MONEM1% = MONEM2%
            If MONEM2% <> MONEM1% Then
                MALCOMBINA% = 1
            End If
            ' OT-8-914 - VALIDA QUE TODOS LOS PEDIDOS TENGAN EL MISMO TRANSPORTE
            Call APERBASDAT("PEDCLI")
            TRANSPO2% = XVALO(VALOBADA("PEDENCA", "NTranspo", "NroPed=" & NPEDXX&))
            If TRANSPO1% = 0 And TRANSPO2% <> 0 Then TRANSPO1% = TRANSPO2%
            If TRANSPO1% <> 0 And TRANSPO2% <> 0 And TRANSPO2% <> TRANSPO1% Then
                MALCOMBINA% = 1
            End If
            '
            DESCVAR2$ = VALOBADA("PEDENCA", "DtoxCant", "NroPed  = " & NPEDXX&, "NOMESS")
            If DESCVAR1$ = "" Then DESCVAR1$ = DESCVAR2$
            If DESCVAR2$ <> DESCVAR1$ Then
                MALCOMBINA% = 1
            End If
            '
          Next KKU&
          '
       Else
            For KKU& = 1 To ULTREG&("!CARDEFAC")
                XXU$ = REGLEIDO$("!CARDEFAC", KKU&)
                NPEDXX& = XVALO(TRIM$(Mid$(XXU$, 371, 16)))
                If NPEDXX& > 0 Then
                    HAYPESEL% = 1
                    Descu1$ = TRIM$(DESCUPEDI$(NPEDXX&))
                    If Descu1$ <> "" Then Exit For
                End If
            Next KKU&
       End If
       '
       If MALCOMBINA% > 0 Then
           Call MENSERR(24, "Se han Detectado Pedidos con Diferencias\que Imposibilitan su Agrupamiento\   \Motivos Posibles:\Diferente Motivo, Moneda, Transportista y/o Descuento.\   \Dada la Configuración Seteada, no Podrá Aprobarse la Factura\que Incluya Ordenes de Venta no Agrupables.") ' MENSAJE PEDIDOS MAL COMBINADOS
       End If
       '
       MONEMI% = MONEM1%
       MONEMIS.Caption = Left$(DEMONECO$(MONEMI%), 8)
       TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
       MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
       MONEMIS.ForeColor = &H404080
       TICAMBIS.Visible = False
       Label24.Visible = False
       If MONEMI% > 1 Then
          MONEMIS.ForeColor = &H0&
          TICAMBIS.Visible = True
          Label24.Visible = True
       End If
       '
       If HAYPESEL% > 0 Then
            PORDESCU = Descu1$
          Else
            NCXX = XVALO(Text1)
            PORDESCU = DESCUCLI$(NCXX)      ' TOMA EL STANDARD DEL CLIENTE
       End If
       '
       ' OT-8-914 - ASIGNA EL TRANSPORTE DEL PEDIDO
       If TRANSPO1% > 0 Then Text19 = TRANSPO1%
       '
       If EsPorPedido& > 0 Then 'MODIFICADA LINEA DE ABAJO PARA QUE TRAIGA EL DESCUENTO DE RETIRA CLIENTE
         PORDES$ = TRIM$(DESCUPEDI$(EsPorPedido&)) 'VALOBADA("PEDENCA", "TDescApli", "NroPed  = " & EsPorPedido&, "NOMESS")
         PORDESCU = PORDES$
       End If
       '
       Text23 = DESCVAR1$
       '
       Call CARDETFAC       ' POR PEDIDO Y SIN CONSIGNACION PENDIENTE.
       '
       Exit Sub
     End If
   End If
   '
   If XVALO(Label31) > 0 Then
      Call CARGAREPA
      Call CARDETFAC
      Exit Sub
   End If

   If TIFADOC% = 3 Or TIFADOC% = 6 Then ' débitos o créditos
     If MODODOC% = 1 Then               ' TEXTO LIBRE
       MODOMONEI% = 0
       MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESVENTA", "MODOMONE"))), 8)
       If MODOMONE$ = "MULTIMON" Then
           MODOMONEI% = 2
       End If
       If MODOMONEI% <> 2 Then Call CARDICAMBIO
     End If
   End If
   '
    ' ASIGNACION DE NOTA DE CREDITO A FACTURA
   If TIFADOC% = 6 Then  ' SI ES NOTA DE CREDITO
    '
     OP% = COMALTER%("Nota de Crédito\\N/Crédito Abierta\Anula Factura")
     '
     If OP% <> 2 Then GoTo 110 ' SI ELIGE LIBRE O CANCELA SE VA A CARDETFAC
     If DERACCE%("CREDEVOL", "Nota de Crédito a Factura") >= 2 Then
       Call CARDETCRE(OKX%)
     Else
       OKX% = 0
     End If
     If OKX% < 1 Then Text1 = "": Exit Sub

   End If
   '
110 Call CARDETFAC
   '
End Sub


Sub SELEMPAQUE()
    'ARMO LA LISTA DE SELECCION
    NC% = XVALO(Text1)
    CONDI$ = "NUMLTAEMP > 0 AND CODI_CLI = " & NC% & " AND EmpaCerra > 0 AND  (NumFactura is null or NumFactura  = '') ORDER BY NumLtaEmp"
    Call CARGALISEL("LEMPEVTA", "ENCEMPAQ", "NUMLTAEMP/I5;REFERENCIA/A48", CONDI$, "DISTINCT")
    Call FRESHECHO("LEMPEVTA")
    '
    'PRESENTO LA LISTA Y VALIDO
    Call SELECHO("LEMPEVTA/LISTAS DE EMPAQUE PENDIENTES")
    NLTAEMPAQ& = XVALO(VALACT1("LEMPEVTA"))
    If NLTAEMPAQ& < 1 Then Exit Sub
    '
    'ARMO CONSULTA PARA CARGAR EL CARDETFAC CON EL DETALLE DE LISTA DE EMPAQUE
    SQLX$ = "SELECT CODINT,NROPED,SUM(PIEZASTOTAL)AS TOTALPIEZAS FROM LISEMPAQ WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NumLtaEmp = " & NLTAEMPAQ&
    SQLX$ = SQLX$ + " GROUP BY CODINT, NROPED"
    'SQLX$ = SQLX$ + " ORDER BY NUMPALLET, NUORDPALLET"
    Dim LISEMPTMP As ADODB.Recordset
    Set LISEMPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With LISEMPTMP
     Do While Not .EOF
        CI% = !CODINT
        If CI% > 0 And CI% <= NUMAS% Then
          CAN = XVALO(!TOTALPIEZAS)
          NPEDID& = !NROPED
          CONDI$ = "NROPED = " & NPEDID& & " AND CODIINT = " & CI%
          PRECI1# = XVALO(VALOBADA("PEDDETA", "PreUnid", CONDI$, "NOMESS"))
          '
          'CALCULO DE RECARGO POR CONDICION DE PAGO
          RECXPAGO$ = VALOBADA("PEDENCA", "Rec_ConPa", "NROPED = " & NPEDID&, "NOMESS")
          CODESC = COEFDESCU(RECXPAGO$)
          If Abs(XVALO(CODESC)) < 0.05 Then CODESC = 0
          PRECI1# = PRECI1# + (PRECI1# * CODESC / 100)
          '
          'TIPO DE MONEDA
          MONEMA% = XVALO(VALOBADA("PEDDETA", "MoneEmis", CONDI$, "NOMESS"))
          If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
          COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
          '
          'SI NO TIENE PRECIO LE APLICA EL DE LA LISTA
          PRECIO = PRECI1# * COEFCON
            If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CI%)
          PRELI = XVALO(VALOBADA("PEDDETA", "prelist", CONDI$, "NOMESS"))  ' MODIFICADO PARA QUE TRAIGA EL PRECIO DE LISTA CON EL CUAL SE CREO EL PEDIDO
            If PRELI < PRECIO Then PRELI = PRECIO
          PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
          '
          'CARGA EL CARDETFAC
5             Z$ = REGBLAN$("!CARDEFAC")
          Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
            COD$ = CODEXT$(CI%)
          Call REPLA(Z$, MKI$(CI%), 5, 2)
          Call REPLA(Z$, DESCRIT$(CI%), 7, 64)
            NREMIT$ = ""
          Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
            FEREMIT% = 0
          '
          Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
          Call REPLA(Z$, NUPLANO$(CI%), 315, 24)
          '
          Call REPLA(Z$, MKS$(PRELI), 421, 4)
          Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
          Call REPLA(Z$, MKS$(PRECIO), 429, 4)
            IMPO# = CDbl(PRECIO) * CAN
          Call REPLA(Z$, MKD$(IMPO#), 433, 8)
          Call REPLA(Z$, MKS$(PORCEIVA(CI%)), 441, 4)
          '
          Call REPLA(Z$, " ", 461, 1)
            '
          NORITPE% = XVALO(VALOBADA("PEDDETA", "NuOrItem", CONDI$, "NOMESS"))
          Call REPLA(Z$, MKI$(NORITPE%), 503, 2)
          Call REPLA(Z$, MKS$(0), 505, 4)
          Call REPLA(Z$, MKS$(0), 509, 4)
          Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
            NUOCOX$ = NUOCOCLI$(NPEDID&, CI%)
10            Call REPLA(Z$, NUOCOX$, 355, 16)
            If TRIM$(NUOCOX$) <> "" Then
              If InStr(Text14, TRIM$(NUOCOX$) + ",") < 1 Then
                Text14 = Text14 + NUOCOX$ + ","
              End If
            End If
            '
          JJ& = JJ& + 1
          Call GRAREG("!CARDEFAC", Z$, JJ&)
          '
        End If
      .MoveNext
     Loop
    End With
    '
    LISEMPTMP.Close
    Set LISEMPTMP = Nothing
    '
    'ASIGNO NUMERO DE LISTA DE EMPAQUE A VARIABLE A NIVEL DE FRM
    FACTLISTAEMPAQUE& = NLTAEMPAQ&
    '

End Sub


'
Sub CARCONSIPEN()
   '
   Static Des%, Has%
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   '
   If PREFACTU$ = "PEDIDO" Or PREFACTU$ = "ORDESPA" Then
     Des% = 33: Has% = 32767: GoTo 3
    End If
   '
   If Des% < 1 Then
2    FEX = HOY(HO$)
     Has% = FEX
     HASS$ = FECHATEX$(FEX)
     DESS$ = "01" + Mid$(HASS$, 3)
     Des% = FECHANUM(DESS$)
   End If
   '
   VDEF$ = MKI$(Des%) + MKI$(Has%)
   OBX$ = OBJPLA$("DESHASFECHA/REMITOS Emitidos DESDE - HASTA Fecha", VDEF$)
   If OBX$ = "" Then
     CONSICLI% = 0
     Exit Sub
   End If
   '
   Des% = CVI(Left$(OBX$, 2))
   Has% = CVI(Mid$(OBX$, 3, 2))
   If Des% > Has% Then
      GoTo 2
   End If
   FEX = Des%: PERIO$ = FECHATEX$(FEX)
   FEX = Has%: PERIO$ = PERIO$ + " al " + FECHATEX$(FEX)
   '
'ASI ORIGINAL EN ACCESS
'3  DES1# = CDbl(CVDAT(DES%))
3   DESELEC = FETEXCOR1$(Des%)
'   HAS1# = CDbl(CVDAT(HAS%))
    HASELEC = FETEXCOR1$(Has%)
   
   
   LPRE% = XVALO(TEXT21(0).TEXT)
   NCLI = XVALO(Text1.TEXT)
   ARCHIPRE$ = TRIM$(Str$(LPRE%)): LIPRACTI% = LPRE%
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   If LPRE% < 1 Then ARCHIPRE$ = "MASTER"
   CONSICLI% = 0
   '
   VUELTA% = 0
41 REGODES& = 0
   Call SETULTREG("!CARDEFAC", 0)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
   End If
   JJ& = 0: CONSICLI% = 0
   '
   CMO$ = TRIM$(Control$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   '
'   ORIMOVI$ = "MOVISTO": ESREPA$ = " "
   '
'4  SQLX$ = " SELECT COD_INTE, CANTSALID, CANTINGRE, NuPedCli, FECHMOV, DoPriCor, MoneVal, VALOUNIT, NUORITEM "
'   SQLX$ = SQLX$ + " , OCOMPRA_RTOABIERTO FROM " + ORIMOVI$ + " WITH(NOLOCK) "
'   SQLX$ = SQLX$ + " WHERE Nume_Clie = " & NCLI & " "
'   SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
'   SQLX$ = SQLX$ + " OR CodiMovi =  'DV' "
'   SQLX$ = SQLX$ + " OR CodiMovi =  '*R') "
'   SQLX$ = SQLX$ + " AND DoPriLar <> '' "
'   SQLX$ = SQLX$ + " AND ((DoSecLar is null) OR (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
'   SQLX$ = SQLX$ + " AND FechMov >= '" & DESELEC & "' "
'   SQLX$ = SQLX$ + " AND FechMov <= '" & HASELEC & "' "
'   '
'   Dim CONSITMP As ADODB.Recordset
'   Set CONSITMP = New ADODB.Recordset
'
'   CONSITMP.CursorLocation = adUseClient
'   CONSITMP.CursorType = adOpenForwardOnly
'   CONSITMP.LockType = adLockReadOnly
'25 On Error Resume Next
'   CONSITMP.Open SQLX$, FLEXCONN, , , adCmdText
'   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'   End If
'   On Error GoTo 0
'   '
''   Set CONSITMP = READSET(SQLX$)
   '
   ORIMOVI$ = "MOVISTO": ESREPA$ = " "
4  If ORIMOVI$ = "MOVISTO" Then SQLX$ = " CARCONSIPEN_SP "
   If ORIMOVI$ = "MOVIREPA" Then SQLX$ = SQLX$ = "CARCONSIPEN_REPA_SP "

   SQLX$ = SQLX$ + " " & NCLI
   SQLX$ = SQLX$ + ",'" & DESELEC & "'"
   SQLX$ = SQLX$ + ",'" & HASELEC & "' "
   Set CONSITMP = FLEXCONN.Execute(SQLX$)
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   MODODOSI% = 0: If InStr(EPAC$, "DOSIVA") > 0 Then MODODOSI% = 1
   '
   If MODODOSI% = 1 Then
     If VUELTA% = 0 Then
       Call RECONCLI(NCLI, HUBORECU%)
       VUELTA% = 1
       If HUBORECU% > 0 Then
          CONSITMP.Close
          Set CONSITMP = Nothing
          GoTo 41
       End If
     End If
   End If
   
   With CONSITMP
'     On Error Resume Next
'     .MoveFirst
'     If Err < 1 Then
'       On Error GoTo 0
       RECONSI$ = REGBLAN("CONSICLI")
       J& = 0
       Do While .EOF = False
         CODIMAT% = XVALO(!cod_inte)
         CAN = XVALO(!CANTSALID) - XVALO(!CANTINGRE)
         NPEDID& = XVALO(!NUPEDCLI)
         If InStr(UCase$(EPAC$), "GUIDI") > 0 Then
            NUMEPED& = 0: PREPED = 0
            If NPEDID& < 1 Then
              FEREMIT% = CVINT%(!FechMov)
              TITUCON$ = TRIM$(SAFETEXT$(!DoPriCor)) + " - Cant: " + TRIM$(Str$(CAN)) + " - Cod: " + TRIM$(CODEXT$(CODIMAT%)) + " - " + DESCRIT$(CODIMAT%)
              Call RECARPED(NCLI, CODIMAT%, FEREMIT%, NUMEPED&, PREPED, TITUCON$)
              NPEDID& = NUMEPED&
            End If
         End If
         NREMIT$ = SAFETEXT$(!DoPriCor)
         '
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = CVINT%(!FechMov)
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(Z$, NUPLANO$(CODIMAT%), 315, 24)
           '
           If InStr(EPAC$, "GUIDI") > 0 Then GoTo 7 ' para forzar que traiga el precio del remito
           '
           If NPEDID& > 0 Then
             ' busca precio y moneda del pedido
             PRECI1# = PRECIOPED#(NPEDID&, CODIMAT%, MONEMA%)
             If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
             COEFCON = TICAMONE(MONEMA%, True) / TICAMONE(MONEMI%, True)
             PRECIO = PRECI1# * COEFCON
             GoTo 8
           End If
           '
7          MONEMA% = XVALO(!MoneVal)
           If MONEMA% <= 1 Then
             MONEPE% = MONEMIPED%(NPEDID&)    ' RECUPERA MONEDA DE PEDIDO
             If MONEPE% > MONEMA% Then MONEMA% = MONEPE%
           End If
           If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
           '
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           PRECIO = XVALO(!VALOUNIT) * COEFCON
             '
8            If PRECIO < 0.00005 Then PRECIO = PREPED
             If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
           PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
           If PRELI < PRECIO Then PRELI = PRECIO
           PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
         Call REPLA(Z$, ESREPA$, 461, 1)    ' PUEDE SER REPARACION
           '
           NORIT% = XVALO(!NuOrItem)
         Call REPLA(Z$, MKI$(NORIT%), 503, 2)
         Call REPLA(Z$, MKS$(0), 505, 4)
         Call REPLA(Z$, MKS$(0), 509, 4)
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
         NUOCOX$ = ""
         If NPEDID& > 0 Then
            NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
         Else
            NUOCOX$ = SAFETEXT$(!OCOMPRA_RTOABIERTO)
         End If
         '
10       Call REPLA(Z$, NUOCOX$, 355, 16)


'         AA1 = !CodiMovi
         Call REGAPP("!CARDEFAC", Z$)
         CAITCAR% = CAITCAR% + 1
         '
11     .MoveNext
       Loop
'     End If
   End With
   CONSITMP.Close
   Set CONSITMP = Nothing
   '
   If MODODOSI% = 1 Then
      If ORIMOVI$ = "MOVISTO" Then
         ORIMOVI$ = "MOVIREPA": ESREPA$ = "R"
         GoTo 4
      End If
   End If
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     Z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(Z$, 1, 4))) >= 0.01 Then
       If MODODOSI% = 1 Then
          If CVS(Mid$(Z$, 1, 4)) < 0 Then GoTo 19
       End If
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", Z$, JJ&)
     End If
19 Next U&
   Call SETULTREG("!CARDEFAC", JJ&)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   Call CIERRARCH("!CARDEFAC")
   '
   If PREFACTU$ = "REMITO" Then
     CONSICLI% = 1
     If JJ& < 1 Then
       Call MENSERR(24, "No hay Remitos Pendientes de Facturación\del " + PERIO$ + " para la Cuenta Elegida.")
       CONSICLI% = 0
     End If
     Exit Sub
   End If
   '
   If JJ& > 0 Then
     If COMALTER%("Hay Material en Consignación\Pendiente para este Cliente\\Facturar Consignaciones\Continuar Normalmente") = 1 Then
       CONSICLI% = 1
       Exit Sub
     End If
   End If
   '
   CONSICLI% = 0
   Call SETULTREG("!CARDEFAC", 0)
   Call CIERRARCH("!CARDEFAC")
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   '
End Sub
'
Sub CARCONSIPED(NPED&)
   '
   Static Des%, Has%
   '
   Des% = 0: Has% = 32767
   '
3  DES1# = CDbl(CVDAT(Des%))
   HAS1# = CDbl(CVDAT(Has%))
   '
   LPRE% = XVALO(TEXT21(0).TEXT)
   NCLIE% = XVALO(Text1.TEXT)
   ARCHIPRE$ = TRIM$(Str$(LPRE%)): LIPRACTI% = LPRE%
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   If LPRE% < 1 Then ARCHIPRE$ = "MASTER"
   CONSICLI% = 0
   '
   REGODES& = 0
   Call SETULTREG("!CARDEFAC", 0)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
   End If
   JJ& = 0: CONSICLI% = 0
   '
   CMO$ = TRIM$(Control$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   '
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE% & " "
   SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
   SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
   SQLX$ = SQLX$ + " AND DoPriLar <> '' "
   SQLX$ = SQLX$ + " AND (DoSECLar = '' "
   SQLX$ = SQLX$ + " OR DoSecLar LIKE 'P/F%' ) "
   SQLX$ = SQLX$ + " AND NuPedCli = " & NPED& & " "
   '
   On Error GoTo 0
   'Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim CONSITMP As ADODB.Recordset
   Set CONSITMP = New ADODB.Recordset
25 On Error Resume Next
   CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With CONSITMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       
       On Error GoTo 0
       RECONSI$ = REGBLAN("CONSICLI")
       J& = 0
       Do While .EOF = False
         CODIMAT% = !cod_inte
         CAN = !CANTSALID - !CANTINGRE
         NPEDID& = 0: If IsNull(!NUPEDCLI) = False Then NPEDID& = !NUPEDCLI
         NUMEPED& = 0: PREPED = 0
           If NPEDID& < 1 Then
             FEREMIT% = CVINT%(!FechMov)
             TITUCON$ = TRIM$(!DoPriCor) + " - Cant: " + TRIM$(Str$(CAN)) + " - Cod: " + TRIM$(CODEXT$(CODIMAT%)) + " - " + DESCRIT$(CODIMAT%)
             Call RECARPED(NCLIE%, CODIMAT%, FEREMIT%, NUMEPED&, PREPED, TITUCON$)
             NPEDID& = NUMEPED&
           End If
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = !DoPriCor
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = CVINT%(!FechMov)
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(Z$, NUPLANO$(CODIMAT%), 315, 24)
           '
           MONEMA% = !MoneVal
           If MONEMA% <= 1 Then
             MONEPE% = MONEMIPED%(NPEDID&)    ' RECUPERA MONEDA DE PEDIDO
             If MONEPE% > MONEMA% Then MONEMA% = MONEPE%
           End If
           '
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           PRECIO = !VALOUNIT * COEFCON
             If PRECIO < 0.00005 Then PRECIO = PREPED
             If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
           PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
           If PRELI < PRECIO Then PRELI = PRECIO
           PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
         Call REPLA(Z$, " ", 461, 1)     ' NUNCA ES REPARACION
           '
           NORIT% = !NuOrItem
         Call REPLA(Z$, MKI$(NORIT%), 503, 2)
         Call REPLA(Z$, MKS$(0), 505, 4)
         Call REPLA(Z$, MKS$(0), 509, 4)
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(Z$, NUOCOX$, 355, 16)
         '
         Call REGAPP("!CARDEFAC", Z$)
         CAITCAR% = CAITCAR% + 1
         '
       .MoveNext
       Loop
12   End If
     On Error GoTo 0
   End With
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     Z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(Z$, 1, 4))) >= 0.01 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", Z$, JJ&)
     End If
   Next U&
   Call SETULTREG("!CARDEFAC", JJ&)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   Call CIERRARCH("!CARDEFAC")
   '
   If PREFACTU$ = "REMITO" Then
     CONSICLI% = 1
     If JJ& < 1 Then
       Call MENSERR(24, "No hay Remitos Pendientes de Facturación\del " + PERIO$ + " para la Cuenta Elegida.")
       CONSICLI% = 0
     End If
     Exit Sub
   End If
   '
   If JJ& > 0 Then
     If COMALTER%("Hay Material en Consignación\Pendiente para este Cliente\\Facturar Consignaciones\Continuar Normalmente") = 1 Then
       CONSICLI% = 1
       Exit Sub
     End If
   End If
   '
   CONSICLI% = 0
   Call SETULTREG("!CARDEFAC", 0)
   Call CIERRARCH("!CARDEFAC")
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   '
End Sub
'
Sub RECARPED(NCLIE, CIPD%, FEREMIT%, NUMEPED&, PREPED, TITUCON$)
   '
   'FEREMIX = CDbl(CVDAT(FEREMIT%))
   FEREMIX = FECHATEX$(FEREMIT%)
   
   
   Call APERBASDAT("PEDCLI")
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PEDDETA "
   SQLX$ = SQLX$ + " WHERE NroClie = " & NCLIE & " "
   SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
   SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
   SQLX$ = SQLX$ + "AND CodiInt = " & CIPD% & " "
   'SQLX$ = SQLX$ + "AND CDBL(INT(FechEmis)) <= " & FEREMIX & " "
   'SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & FEREMIX & ")) "
   SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= '" & FEREMIX & "'))"
'   SQLX$ = SQLX$ + " AND FinVigen >= '" & FEREMIX & "'"
   SQLX$ = SQLX$ + " ORDER BY NroPed ASC"
   '
   'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim PEDTEMP As ADODB.Recordset
   Set PEDTEMP = New ADODB.Recordset
   PEDTEMP.CursorLocation = adUseClient
25 On Error Resume Next
   PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   NUMEPED& = 0
   PREPED = 0
   ENCUPE% = 0
   '
   With PEDTEMP
     On Error Resume Next
     .MoveLast
     CAREGIS& = .RecordCount
     '
     If CAREGIS& = 1 Then
         NUMEPED& = !NROPED
         PREPED = !PreUnid
         ENCUPE% = 1
       ElseIf CAREGIS& > 1 Then
         .MoveFirst
         JJ& = 0
         Do While (.EOF = False)
           TTXX$ = MKS$(!NROPED)
           TTXX$ = TTXX$ + AJUSTI$(!NROOCOM, 16) + "  "
              FEX = CVINT(!FECHEMIS)
           TTXX$ = TTXX$ + FECHATEX$(FEX) + "  "
              FEX = CVINT(!FeSolEnt)
           TTXX$ = TTXX$ + FECHATEX$(FEX) + "  "
           TTXX$ = TTXX$ + FORMATNUM$(!PreUnid, "F12.4")
              CAPEN = !CANPED - !CANTENT
              If CAPEN < 0 Then CAPEN = 0
              If !TipoPed = "ABIERTO" Then CAPEN = 999999
           TTXX$ = TTXX$ + FORMATNUM$(CAPEN, "F10.1")
           TTXX$ = TTXX$ + FORMATNUM(!NuOrItem, "F6.0")
           JJ& = JJ& + 1
           Call GRAREG("!SELECPE", TTXX$, JJ&)
         .MoveNext
         Loop
         '
         Call SETULTREG("!SELECPE", JJ&)
         '
20       Call FRESHECHO("!SELECPE")
         TTXX$ = REGSEL$("!SELECPE/" + TITUCON$)
         If Len(TTXX$) <= 4 Then GoTo 20
         '
         NUMEPED& = CVS(Left$(TTXX$, 4))
         PREPED = XVALO(Mid$(TTXX$, 45, 12))
         ENCUPE% = 1
      End If
    End With
    PEDTEMP.Close
    Set PEDTEMP = Nothing
    '
End Sub
'
Sub CARDICAMBIO()
   '
   DOCUDIF = ""
   MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESVENTA", "MODOMONE"))), 8)
   If MODOMONE$ <> "BIMONETA" Then
      If MODOMONE$ <> "MULTIMON" Then
         Exit Sub
      End If
   End If
   '
   Dim DOCABX$(2048), SALPES#(2048), SALDOL#(2048), REDOCA&(2048)
   '
   NCLI = XVALO(Text1)
   Call BLANARCH("!CARDEFAC")
   '
   Screen.MousePointer = 11
   CADOCPE% = 0
   RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
   For KUL& = 1 To Len(RFG$) Step 4
     RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
     XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
     DOCABY$ = Mid$(XXC$, 13, 18)
     SALPESY# = CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8))
     SALDOLY# = CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8))
     '
     For KKI% = 1 To CADOCPE%
       If DOCABX$(KKI%) = DOCABY$ Then GoTo 20
     Next KKI%
     CADOCPE% = CADOCPE% + 1
     DOCABX$(CADOCPE%) = DOCABY$
     REDOCA&(CADOCPE%) = RECABIN&
     KKI% = CADOCPE%
     '
20   SALPES#(KKI%) = SALPES#(KKI%) + SALPESY#
     SALDOL#(KKI%) = SALDOL#(KKI%) + SALDOLY#
     '
   Next KUL&
   '
   Call BLANARCH("SELDOCAP")
   For KKI% = 1 To CADOCPE%
     If TIFADOC% = 6 Then    ' N.CREDITO
       If SALDOL#(KKI%) <= 0.005 Then
         If SALPES#(KKI%) > 0 Then
           XXC$ = REGLEIDO$("CARBIMON", REDOCA&(KKI%))
           FEMI = CVI(Mid$(XXC$, 1, 2))
           Z$ = Space$(60) + MKS$(0)
           Call REPLA(Z$, DOCABX$(KKI%), 1, 18)
           Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
           Call REPLA(Z$, "$" + FORMATNUM$(SALPES#(KKI%), "F11.2"), 29, 12)
           Call REPLA(Z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F12.2"), 42, 16)
           Call REPLA(Z$, MKS$(REDOCA&(KKI%)), 61, 4)
           Call REGAPP("SELDOCAP", Z$)
         End If
       End If
     ElseIf TIFADOC% = 3 Then  ' N.DEBITO
       If SALDOL#(KKI%) > 0 Then
         If SALPES#(KKI%) <= 0.005 Then
           XXC$ = REGLEIDO$("CARBIMON", REDOCA&(KKI%))
           FEMI = CVI(Mid$(XXC$, 1, 2))
           Z$ = Space$(60) + MKS$(0)
           Call REPLA(Z$, DOCABX$(KKI%), 1, 18)
           Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
           Call REPLA(Z$, "$" + FORMATNUM$(SALPES#(KKI%), "F11.2"), 29, 12)
           Call REPLA(Z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F12.2"), 42, 16)
           Call REPLA(Z$, MKS$(REDOCA&(KKI%)), 61, 4)
           Call REGAPP("SELDOCAP", Z$)
         End If
       End If
     End If
   Next KKI%
   Call CIERRARCH("SELDOCAP")
   '
   Screen.MousePointer = 1
   If ULTREG&("SELDOCAP") < 1 Then
     Exit Sub
   End If
   '
   Call SELECHO("SELDOCAP")
   DOCABY$ = Left$(VALACT1$("SELDOCAP"), 18)
   If TRIM$(DOCABY$) = "" Then Exit Sub
   DOCUDIF = DOCABY$
   BBXX$ = VALACT2$("SELDOCAP")
   '
   MONEMIS = "Pesos"
   IDIFE = XVALO(Mid$(BBXX$, 12, 11))
   If TIFADOC% = 3 Then
     IDIFE = XVALO(Mid$(BBXX$, 28, 12)) * TICAMONE(2)
   End If
   If PIVADOC% = 1 Then
     TASA1 = XVALO(PORIVA(0).Caption)
     IDIFE = IDIFE / (1 + TASA1 / 100)
   End If
   '
   ZZ$ = REGBLAN$("CARDEFAC")
   Call REPLA(ZZ$, MKS$(1), 1, 4)
   Call REPLA(ZZ$, "Por Diferencia de Cambio", 7, 64)
   Call REPLA(ZZ$, MKS$(IDIFE), 429, 4)
   Call REPLA(ZZ$, MKD$(CDbl(IDIFE)), 433, 8)
   Call REGAPP("!CARDEFAC", ZZ$)
   ZZ$ = REGBLAN$("CARDEFAC")
   Call REPLA(ZZ$, "s/Factura " + DOCABY$, 7, 64)
   Call REGAPP("!CARDEFAC", ZZ$)
   Call CIERRARCH("!CARDEFAC")
   '
End Sub
'
Private Sub Command7_Click()
    Command7.Enabled = False
    Call CIERRARCH("*.*")
    MODX% = ALTERNA%("Listas de Precios\Condiciones de Pago\Alícuotas de I.V.A.")
    If MODX% = 1 Then
         Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
       ElseIf MODX% = 2 Then
         Call CARCONPAGOS
       ElseIf MODX% = 3 Then
         Call CARTASIVA
         Call COMUNI("Re-Inicie Ahora la Aplicación.")
         Call CIERRARCH("*.*")
         Call FINAL("")
    End If
    Command7.Enabled = True
End Sub

Sub Command8_Click()
    '
    Dim RESP%
    Dim NUMEOCXX$
    '
    Command8.Enabled = False
    MODO% = COMALTER%("Opciones de Salida\\Pantalla\Impresora\Cancelar")
    If MODO% < 1 Or MODO% > 2 Then GoTo 99
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
10  FEX = HOY(HO$)
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
        GoTo 10
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    'jandy sql
    'DES1# = CDbl(CVDAT(DES%))
    'HAS1# = CDbl(CVDAT(HAS%))
    DES1$ = FETEXCOR1$(Des%)
    HAS1$ = FETEXCOR1$(Has%)
    
    '
    If MODO% = 1 Then GoTo 56
    '////////******NUEVO FILTRO CLIENTE*******///////
    NCMIN% = 32767: NCMAX% = 0
    FIN& = ULTREG&("DEUDOR1")
    For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("DEUDOR1", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
    Next U&
    '
    With RACLIRPF
      .Text10 = TRIM$(Str$(NCMIN%))
      .Text9 = rasocli$(NCMIN%)
      .Text8 = TRIM$(Str$(NCMAX%))
      .Text7 = rasocli$(NCMAX%)
      .Show 1
    End With
    '
    'SI CANCELA EN EL FORMULARIO DE RANGO DE CTAS. DE CLIENTES
    If TRIM$(RACLIRPF.Text10.TEXT) = "" And TRIM$(RACLIRPF.Text8.TEXT) = "" Then
        Command8.Enabled = True
        Exit Sub
    End If
    '
    CLINFE = XVALO(RACLIRPF.Text10)
    CLISUP = XVALO(RACLIRPF.Text8)
    '//////********HASTA ACA NUEVO FILTRO CLIENTE****///////
56  Screen.MousePointer = 11
    Call APERBASDAT("STOCKS")
    '
    'JANDY SQL
    
    Dim CONSITMP As ADODB.Recordset
    Set CONSITMP = New ADODB.Recordset
    CONSITMP.CursorLocation = adUseClient
    
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND ((DoSecLar IS NULL) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
    SQLX$ = SQLX$ + " AND FechMov >= '" & DES1$ & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HAS1$ & "' "
'
'    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
'    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DES1# & " "
'    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
'

    'PARA EVITAR COMPROBANTES ARRUINADOS
    'SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    SQLX$ = SQLX$ + "AND REFERCOR NOT LIKE 'ANUL%'"
    'FILTRO CLIENTE
    If MODO% <> 1 Then
      SQLX$ = SQLX$ + " AND Nume_Clie >= " & CLINFE & ""
      SQLX$ = SQLX$ + "AND Nume_Clie<= " & CLISUP & ""
    End If
    '
    'Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
25  On Error Resume Next
    CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    With CONSITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Remitos Pendientes de Facturación.")  'Generados para este Cliente")
        On Error GoTo 0
        Screen.MousePointer = 1
        Command8.Enabled = True
        Exit Sub
      End If
      On Error GoTo 0
      '
      RECONSI$ = REGBLAN("CONSICLI")
      RECONSJ$ = REGBLAN("REPENFAC")
      J& = 0
      Do While .EOF = False
         NCLIE% = !NUME_CLIE
         Call REPLA(RECONSI$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECONSI$, MKI$(CVINT(!FechMov)), 3, 2)
         Call REPLA(RECONSI$, "", 5, 18)
         Call REPLA(RECONSI$, !DoPriCor, 23, 12)
         Call REPLA(RECONSI$, MKI$(!cod_inte), 35, 2)
         Call REPLA(RECONSI$, MKD$(!CANTSALID), 37, 8)
         Call REPLA(RECONSI$, MKD$(!CANTINGRE), 45, 8)
         Call REPLA(RECONSI$, MKD$(!CantSaldo), 53, 8)
         Call REPLA(RECONSI$, !ReferCor, 61, 48)
         Call REPLA(RECONSI$, String$(12, 0), 109, 12)
         Call REPLA(RECONSI$, MKS$(XVALO(Mid(!DoPriCor, 5))), 109, 4)
         Call REPLA(RECONSI$, MKI$(CVINT(!FeReMovi)), 122, 2)
          HOR = Format(!FeReMovi, "HH:MM:SS")
         Call REPLA(RECONSI$, MKI$(CVINT(HOR)), 124, 2)
         Call REPLA(RECONSI$, MKI$(XVALO(!NumUsuar)), 127, 2)
         '
         Call REPLA(RECONSJ$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECONSJ$, MKI$(!cod_inte), 3, 2)
         Call REPLA(RECONSJ$, !DoPriCor, 5, 12)
         Call REPLA(RECONSJ$, !DOPRILAR, 17, 18)
         Call REPLA(RECONSJ$, MKI$(CVINT(!FechMov)), 35, 2)
         If IsNull(!NUPEDCLI) Then
            Call REPLA(RECONSJ$, MKS$(0), 37, 4)
         Else
            Call REPLA(RECONSJ$, MKS$(!NUPEDCLI), 37, 4)
         End If
         If IsNull(!NUPEDCLI) Then
             NUMEOCXX$ = ""
           Else
             NUMEOCXX$ = NUOCOCLI$(!NUPEDCLI, !cod_inte)
         End If
         Call REPLA(RECONSJ$, NUMEOCXX$, 41, 24)
         Call REPLA(RECONSJ$, MKD$(!CANTSALID), 65, 8)
         Call REPLA(RECONSJ$, MKD$(!CANTINGRE), 73, 8)
         Call REPLA(RECONSJ$, MKD$(!CantSaldo), 81, 8)
         Call REPLA(RECONSJ$, MKD$(!VALOUNIT), 89, 8)
           TOTI# = !VALOUNIT * (!CANTSALID - !CANTINGRE)
         Call REPLA(RECONSJ$, MKD$(TOTI#), 97, 8)
         '
         J& = J& + 1
         Call GRAREG("CONSICLI", RECONSI$, J&)
         Call GRAREG("REPENFAC", RECONSJ$, J&)
      .MoveNext
      Loop
    End With
    CONSITMP.Close
    Set CONSITMP = Nothing
    '
    Call SETULTREG("CONSICLI", J&)
    Call FILESORT("CONSICLI", "", 1, 6, "ASC")
    Call CIERRARCH("CONSICLI")
    '
    Call SETULTREG("REPENFAC", J&)
    Call FILESORT("REPENFAC", "", 1, 6, "ASC")
    Call CIERRARCH("REPENFAC")
    '
    If MODO% = 1 Then
        Call CONECRUN("*CREPENFA", "Remitos Sin Facturar (Clientes)")
      Else
        RESP% = COMALTER%("Opciones de Agrupamiento\\Comprobante\Articulo\Cancelar")
        If RESP% = 1 Then
            Call CONECRUN("*LIREPEFA", "Remitos Sin Facturar (Clientes)")
          ElseIf RESP% = 2 Then
            Call HEADERS("REPENFAC", "")
            If CLINFE = CLISUP Then
                NCXX% = CLINFE
                Call HEADERS("REPENFAC", "Cliente: " + rasocli$(NCXX%))
                Call CONECRUN("*LREPENFB", "Remitos Sin Facturar (Clientes)")
              Else
                Call CONECRUN("*LREPENFA", "Remitos Sin Facturar (Clientes)")
            End If
        End If
    End If
    '
99  Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Call SELECHO("SELFECHA")
   Text9.TEXT = VALACT1$("SELFECHA")
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub

Private Sub Form_Load()
    '
    Screen.MousePointer = 11
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If    '
    Label28 = ""
    'Frame5.Width = 4500
    'Eco21(0).Width = 2250: Eco21(0).FontSize = 7: Eco21(0).Height = 300: Eco21(0).FontSize = 9
    'Eco21(1).Width = 2250: Eco21(1).FontSize = 7: Eco21(1).Height = 300: Eco21(1).FontSize = 9
    '
    If InStr(EMPREAC$, "MEDITEA") > 0 Then Command27.Visible = True
    If InStr(EMPREAC$, "EYM") > 0 Then Command27.Visible = True
    '
    If Control$("ADMIPED", "MOCOMIS") = "LETRA" Then
        Label28.Visible = True
        Text22.Visible = False
      ElseIf Control$("ADMIPED", "MOCOMIS") = "PORCIEN" Then
        Label28.Visible = False
        Text22.Visible = True
    End If
    '
    Call CREATABLA_PROFORMTL 'TABLA DE PROFORMA
    MsgBox "1"
'    CAMPOS$ = "Cant./F6;Código/A16;Descripción/A22;Lista/F10.2;Desc/F6.2;P.Unit/F12.2;Importe/F14.2;R/A1;"
'    Call CARGA_ENCABEZADO(Me.MSF, CAMPOS$, Me)
    If TRIM$(Control("FACTURA", "PREFACTU")) <> "REMITO" Then
      Campos$ = "Cant./F6;Código/A16;Descripción/A32;Lista/F10.2;Desc/F6.2;P.Unit/F12.2;Importe/F14.2;R/A1;"
    Else
      Campos$ = "Cant./F6;Código/A16;Descripción/A24;Lista/F10.2;N°Ped/F6;Desc/F6.2;P.Unit/F12.2;Importe/F14.2;R/A1;"
    End If
    Call CARGA_ENCABEZADO(Me.MSF, Campos$, Me)

    ' *** QUITA EL ESTILO NEGRITA AL TEXTO DEL ENCABEZADO
    For i = 1 To 8: MSF.COL = i: MSF.CellFontBold = False: Next i
    
    Refresh
    Frame1.Refresh
    DoEvents
    '\\\OT-05-0419-WAR-23/06/05///
    If Control$("FACTURA", "SUPCONF") <> "SI" Then
       Call CARGAPUERTOS
    End If
    '
    If EXISTE%("PEDIDO07.EXE") > 0 Then
       Command17.Visible = True
    End If
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$: DATEDOC% = HOII%
    '
    MONEMI% = 1  ' PESOS
    TICAMBIS = "1.0000"
    TICAMBIS.Visible = False
    Label24.Visible = False
    CONVERPRECIO% = 1
    '
    Call SETULTREG("!CARDEFAC", 0)
    If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
      On Error Resume Next
      Kill LUCOM$ + "1CARDEFA.WKF"
      On Error GoTo 0
    End If
    '
    TIFADOC% = 2: Label7.Caption = "Factura:"
    Label6.Caption = "(Mat.Codif)": MODODOC% = 0
    Label13.Caption = "A": PIVADOC% = 1: VARIDOC% = 1
'    UPREFI$ = TRIM$(CONTROL$("FACTURA", "PREFIFA"))
    CLAVE$ = "PREFAC-" & Label13.Caption '  MODIFICADO PARA GUARDAR PREFIJO DE FACTURA POR ESTACION
    UPREFI$ = TRIM$(Control$(CLAVE$, IDENTER$))
    If InStr(EPAC$, "DOSIVA") > 0 Then UPREFI$ = "0001"
    '
    If UPREFI$ = "" Then UPREFI$ = "0001"
    Label12.Caption = UPREFI$
    '
    Printer.TrackDefault = False
    '
    VERLICRE% = 0
    If Control$("FACTURA", "VERLICRE") = "SI" Then VERLICRE% = 1
    MsgBox "2"
    Call LEETASIVA
    MsgBox "3.1"
    'Call NUFACTURA ' TRASLADADO MAS ABAJO PARA QUE ESTE INICIALIZADO LA I.FISCAL
    '
    'If InStr(EPAC$, "DEL CIELO") > 0 Then
    If Control("FACTURA", "SISTMANT") = "SI" Then
      Command31.Visible = True
    End If
    
    MODESCU% = 0
    If Control$("FACTUREM", "MODESCUE") = "PORCIENTO" Then
       MODESCU% = 1
    End If
    '
    ' TIPO DE CAMBIO POR OMISION
    TICAMBX# = 1
    MONEMI% = 1   ' pesos
    TICAMBIS = "1.0000"
    TICAMBIS.Visible = False
    Label24.Visible = False
    '
    If EsPorPedido& > 0 Then PREFACTU$ = "ORDESPA"
    '
    If PREFACTU$ = "" Then
      PREFACTU$ = TRIM$(Control("FACTURA", "PREFACTU"))
      If PREFACTU$ = "" Then PREFACTU$ = "REMITO"   ' si no ORDESPA / PEDIDO
    End If
    '
    FACELEC% = 0
    If CONTROFIS% = 0 Then
       CONTROFIS% = (-1)
       Set FORFACPRINT = FORFASQL
       If Control$("FACTURA", "SALIDA") = "CFISCAL" Then
          CONTROFIS% = 1
          FORCARGA$ = "SINREMITO"
          Check1.Visible = True
          ' INICIALIZAR CONTROLADOR FISCAL
          'Call FORHASAR.INICIA
          Call COMUNI("No Disponible en la Presente Version")
          Call FINAL("")
          '
       ElseIf Control$("FACTURA", "SALIDA") = "IFISCAL" Then
          CONTROFIS% = 2
          Check1.Visible = True
          ' INICIALIZAR CONTROLADOR FISCAL
          'Call FORHAS320.INICIA
          Call COMUNI("No Disponible en la Presente Version")
          Call FINAL("")
          On Error Resume Next
          Call FORHAS320.HASAR1.ObtenerDatosDeInicializacion
          YY$ = FORHAS320.HASAR1.RESPUESTA(0)
          On Error GoTo 0
          PUNTODEVENTA$ = TRIM$(Mid$(YY$, 92, 4))
          Call COMUNI("El Punto de Venta de la Impresora Fiscal es '" + PUNTODEVENTA$ + "'")
          '
       ElseIf Control$("FACTURA", "SALIDA") = "FELECT" Then ' FACTURA ELECTRONICA
          '
          'Call Controla_Idcaenum
         ' Call Controla_NroRemito(Label12)
          MsgBox "2.1"
          CONTROFIS% = 0
          FACELEC% = 1
          OPVARFA07.Command7.Visible = True
          OPVARFA07.Command5.Visible = True
          '
          ' si no tiene configurado punto de venta normal
          If XVALO(Control("FACTURA", "PUVENELE")) < 1 Then GoTo 40
          '
          VUELTA% = 0
          CARPETOK$ = "C:\FLEXOFT\FACELEC\"
          ARCHIVO_TOKEN$ = "TOKEN.TXT"
30        If EXISTE%(CARPETOK$ + ARCHIVO_TOKEN$) = 1 Then
             FechArch = FileSystem.FileDateTime(CARPETOK$ + ARCHIVO_TOKEN$)
             FECHARCHIVO% = CVINT(FechArch)
             FECHHOY% = HOY(HO$)
            'EMPRESAS CON 2 FACTURAS ELECTRONICAS
             '************************************
             'OBTENGO CUIT DE EMPRESA ACTIVA
             n1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
             Dim X As String * 41
             Get #n1%, 6, X
             
             'VALIDACION
             CUITT$ = TRIM$(DESENCRI$(X))
             Valid% = VALICUIT%(CUITT$)
             Close #n1%

             If Valid% < 1 Then
               Call COMUNI("C.U.I.T. de Empresa No Válido")
               Call FINAL("")
             End If
             '
             CUIT_EMPRESA_ACTIVA$ = TRIM$(REPLACAR$(CUITT$, "-", ""))
             

             'OBTENGO EL CUIT QUE ESTA EN C:\FLEXOFT\FACELEC
             If EXISTE%(CARPETOK$ + "CUITEMP.TXT") > 0 Then
               NARCHI2% = FILEOPEN%(CARPETOK$ + "CUITEMP.TXT", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
               Line Input #NARCHI2%, CUIT_FACELEC$
             End If
             'COMPARO APARTE DE LA FECHA POR EL CUIT DE LA EMPRESA PARA LAS QUE UTILIZAN DOS EMPRESAS
             '
             Close #NARCHI2%
             '
             TIENETOKEN% = 1 ' VALIDA SI SE CREO EL ARCHIVO POR QUE SI NO TIENE ACTUALIZADO EL OBTIENE TOKEN NUNCA VA A PASAR
             If FECHARCHIVO% < FECHHOY% Or (TRIM$(CUIT_FACELEC$) <> TRIM$(CUIT_EMPRESA_ACTIVA$)) Then TIENETOKEN% = 0
             '
             If TIENETOKEN% < 1 Then
               Screen.MousePointer = 1
               Call MENSERR(24, "Token Vencido. Vuelva a Generarlo")
               OPVARFA07.Show 1
               Call FINAL("")
               'Timer2.Enabled = true
             End If
          Else
             Screen.MousePointer = 1
             Call MENSERR(24, "Debe Generarse el Archivo 'TOKEN.TXT'\al Ingresar a Facturacion.")
             OPVARFA07.Show 1
             Call FINAL("")
          End If
       End If
    End If
    '
40  'Call NUFACTURA    ' 'TEMA REMITO SE TRASLADÓ AQUÍ PARA QUE YA ESTÉ INICIALIZADO LA IMPRESORA FISCAL
    '
    
    ' *** VERIFICA SI LA EMPRESA ES AGENTE DE PERCEPCION DE INGRESOS BRUTOS PARA LA PROVINCIA DE MISIONES
    ' *** EN CASO AFIRMATIVO PONE VISIBLE LAS CAJAS DE TEXTO
    CUEPERI$ = Control$("PERIBMI", "CUECONTA")
    CUENPER% = CUECOINT%(CUEPERI$)
    If TRIM$(CUEPERI$) <> "" And CUENPER% > 1 Then
        Text25.Visible = True: Label40.Visible = True
    End If
    '
    ' *** VERIFICA SI LA EMPRESA ES AGENTE DE PERCEPCION DE INGRESOS BRUTOS PARA LA PROVINCIA DE TUCUMAN
    ' *** EN CASO AFIRMATIVO PONE VISIBLE LAS CAJAS DE TEXTO
    CUEPERI$ = Control$("PERIBTU", "CUECONTA")
    CUENPER% = CUECOINT%(CUEPERI$)
    If TRIM$(CUEPERI$) <> "" And CUENPER% > 1 Then
        Text26.Visible = True: Label41.Visible = True
    End If
    '
    If EsPorPedido& > 0 Then
      NCLIE = XVALO(VALOBADA("PEDENCA", "NroClie", "NroPed = " & EsPorPedido& & "", "NOMESS"))
      Text1 = NCLIE
    End If
    '
    MsgBox "3"
    'DERECHO DE MODIFICACION DE DESCUENTO EN PANTALLA
    If DERACCE%("MODIDESC/NOMESS", "MODIFICAR DESCUENTO EN PANTALLA") < 2 Then
       PORDESCU.Enabled = False
    End If
    '
    If InStr(1, EMPREAC$, "GABAL") > 0 Then
        PORDESCU.Width = 585
    End If
    '
    ' *** DIMENSIONA VECTOR PARA GUARDAR EL IMPORTE NETO DE LA VENTA POR CUENTA CONTABLE.
    ' *** EL INDICE ES EL CODIGO INTERNO DE LA CUENTA CONTABLE
    NetoMultipleImputacion% = 0
    If InStr(1, EMPREAC$, "CEYEQ") > 0 Then
        CodiIntMaximoCuentaContable% = XVALO(VALOBADA("PLANCUEN", "Max(CodInCue)", "CodInCue > 0 AND STATCUE > 0"))
        If CodiIntMaximoCuentaContable% > 0 Then
            ReDim VentaNetaPorCuentaContable#(CodiIntMaximoCuentaContable% + 1)
            NetoMultipleImputacion% = 1
        End If
    End If
    '
    MsgBox "4"
'    Call INICIALIZACLIREAL
    '
    Call APLICACIONSKINS(Me, Picture2)
    '
    Call InstalaFuenteAnchoFijo
    Call AplicaFuenteAnchoFijo(MSF, , 8)
    Call AplicaFuenteAnchoFijo(TLABEL28, , 8)
    Call AplicaFuenteAnchoFijo(Text17, [3. Negrita], 8)
    Call AplicaFuenteAnchoFijo(Text15, [3. Negrita], 8)
    Call AplicaFuenteAnchoFijo(Text25, [3. Negrita], 8)
    Call AplicaFuenteAnchoFijo(Text26, [3. Negrita], 8)
    Call AplicaFuenteAnchoFijo(Text24, [3. Negrita], 8)
    Call AplicaFuenteAnchoFijo(Text18, [3. Negrita], 8)
    For i = 0 To 6
        Call AplicaFuenteAnchoFijo(Total(i), [3. Negrita], 8)
        If i < 3 Then Call AplicaFuenteAnchoFijo(BASIVA(i), [3. Negrita], 8)
    Next i
    '
    MsgBox "5"
    Screen.MousePointer = 1
    '
End Sub


'
Sub INICIALIZACLIREAL()
    '
    Call ABRECONEXION
    Call CREATABLA_CUECORRI
    '
    'CONDI$ = "TIPOMOVIM = 'FVEN' AND (CLI_REAL = 0 OR CLI_REAL IS NULL)"
    CONDI$ = "TIPOMOVIM = 'FVEN' AND NUCLIEN > 0 AND (CLI_REAL = 0 OR CLI_REAL IS NULL)"
    FIN& = CantRegistros("CAJABANC", CONDI$, "NOMESS")
    If FIN& > 0 Then
        Call COMUNI("Deberan Realizarse Cambios Sobre la Base de Datos\Los Cuales Podran Demorar Algunos Minutos.\Favor de Esperar Mensaje: 'Proceso Terminado'")
        '
        Call PoneEnCeroNull("CAJABANC", "CLI_REAL", REG&, "NOMESS")
        '
        Screen.MousePointer = 11
        SQLX$ = "SELECT DISTINCT(CODICOM_LAR), NUCLIEN FROM CAJABANC WITH(NOLOCK) WHERE " & CONDI$
        Set RSCABAN = READSET(SQLX$)
        i& = 0
        With RSCABAN
            Do While Not .EOF
                i& = i& + 1
                Call TRACE(24, i&, FIN&)
                NUFACTU$ = SAFETEXT$(!codicom_lar)
                NCLI% = XVALO(!NUCLIEN)
                Condi1$ = "DOSECLAR = '" & NUFACTU$ & "'"
                NCLIE% = XVALO(VALOBADA("MOVISTO", "NUME_CLIE", Condi1$, "NOMESS"))
                If NCLIE% = 0 Then NCLIE% = NCLI%
                '
                CONDI2$ = "CODICOM_LAR='" & NUFACTU$ & "'"
                Call ACTUREGISTRO("CAJABANC", "CLI_REAL", CONDI2$, NCLIE%, REG&, "NOMESS")
            .MoveNext
            Loop
        End With
        RSCABAN.Close
        Set RSCABAN = Nothing
        '
        Screen.MousePointer = 1
        Call COMUNI("Proceso Terminado")

    End If
    
End Sub

Sub CARGAPUERTOS()
    '
    Printer.TrackDefault = False
    PORTUFAC$ = TRIM$(Control$(IDENTER$, "PORTFAC"))
      If PORTUFAC$ = "" Then PORTUFAC$ = Control$("*", "PORTFAC")
      If PORTUFAC$ = "" Then GoTo 20
      NEWPORT$ = PORTUFAC$    ' + "/NOMESS"
      Call SETPRINPORT(NEWPORT$)
      If NEWPORT$ = "" Then GoTo 20
      Call SETPRINPORT("RESTORE")
    PORTUREM$ = TRIM$(Control$(IDENTER$, "PORTREM"))
      If PORTUREM$ = "" Then PORTUREM$ = Control$("*", "PORTREM")
      If PORTUREM$ = "" Then GoTo 20
      NEWPORT$ = PORTUREM$   ' + "/NOMESS"
      Call SETPRINPORT(NEWPORT$)
      If NEWPORT$ = "" Then GoTo 20
      Call SETPRINPORT("RESTORE")
    Exit Sub    ' PUDO ACCEDER A LOS PUERTOS DE IMPRESIóN
    '
20  Screen.MousePointer = 1
    DIREPRIN.Show 1
    Call FRESHALL
    DoEvents
    If DIREPRIN.Label3 <> "OK" Then
      Call FINAL("")
    End If
    '
    PORTUFAC$ = DIREPRIN.Combo1.TEXT
    PORTUREM$ = DIREPRIN.Combo2.TEXT
    '
    If PORTUFAC$ <> "" Then
      Call SETPRINPORT(PORTUFAC$)
      Call SETPRINPORT("RESTORE")
    End If
    
    If PORTUREM$ <> "" Then
      Call SETPRINPORT(PORTUREM$)
      Call SETPRINPORT("RESTORE")
    End If
    
    Call GRACONTROL("*", "PORTFAC", PORTUFAC$)
    'If TRIM$(CONTROL$(IDENTER$, "PORTFAC")) <> "" Then
      Call GRACONTROL(IDENTER$, "PORTFAC", PORTUFAC$)
    'End If
    Call GRACONTROL("*", "PORTREM", PORTUREM$)
    'If TRIM$(CONTROL$(IDENTER$, "PORTREM")) <> "" Then
      Call GRACONTROL(IDENTER$, "PORTREM", PORTUREM$)
    'End If
    '
End Sub
'
Private Sub Form_Resize()
   MsgBox "6"
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
   MsgBox "7"
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Label19_DblClick()
   Static CTXX%
   '
   If ULTREG&("DAPIBCLI") < 1 Then
     Call MENSERR(24, "Imposible desde esta Aplicación")
     Exit Sub
   End If
   '
   NCLIE% = XVALO(TRIM$(Text1.TEXT))
   If NCLIE% > 0 Then
     If rasocli$(NCLIE%) <> "" Then
       '
       If CTXX% < 1 Then
         Call GRAREG("TACATIBC", Chr$(0) + "Exento Percepcion IB", 1)
         Call GRAREG("TACATIBC", Chr$(1) + "Sujeto Percepcion IB", 2)
         Call SETULTREG("TACATIBC", 2)
         CTXX% = 1
       End If
       '
       RECORD$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE%))
       Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
       CATIB1% = Asc(Mid$(RECORD$, 3, 1))
       If CATIB1% < 0 Or CATIB1% > 1 Then
         Call REPLA(RECORD$, Chr$(1), 3, 1)
         Call REPLA(RECORD$, Space$(32), 9, 32)
       End If
       '
10     OBX$ = OBJPLA$("CARDAPERIB", RECORD$)
       If OBX$ <> "" Then
         RECORD$ = OBX$
         CATIB1% = Asc(Mid$(RECORD$, 3, 1))
         If CATIB1% < 0 Or CATIB1% > 1 Then
           Call MENSERR(24, "Mal Categoría I.Brutos")
           GoTo 10
         End If
         If CATIB1% = 0 Then
           If TRIM$(Mid$(RECORD$, 9, 32)) = "" Then
             Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
             GoTo 10
           End If
         End If
         If CATIB1% = 1 Then
           Call REPLA(RECORD$, Space$(32), 9, 32)
         End If
         PORCONVE = CVS(Mid$(RECORD$, 5, 4))
         If PORCONVE < 0 Or PORCONVE > 100 Then
           Call MENSERR(24, "% Convenio Multilateral\Fuera de Rango")
           GoTo 10
         End If
         '
         Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECORD$, USERNAME$, 41, 20)
         Call REPLA(RECORD$, MKI$(HOY(HOS$)), 61, 2)
         Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 63, 2)
         Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIE%), RECORD$)
         Call CIERRARCH("DAPIBCLI")
         '
         Call CALPERIB
         '
       End If
     End If
   End If

End Sub

Private Sub Label26_DblClick()
   Call SELECHO("INDICLIE")
   CPROX$ = TRIM$(VALACT1$("INDICLIE"))
   If CPROX$ <> "" Then
     NCXI% = NUMECLI(CPROX$)
     If NCXI% > 0 Then
       If REGICLI&(NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
         Call CARDACLI
         Call Label12_Change
       End If
     End If
   End If
End Sub

Sub Label6_Click()
   If ULTREG&("!CARDEFAC") > 0 Then Exit Sub
   '
   'If TIFADOC% = 2 Or TIFADOC% = 6 Then
   If MODODOC% = 0 Then
        Label6.Caption = "(T.Libre)": MODODOC% = 1
' validacion suprimida porque s/flavio sí se puede hacer factura
' de texto libre para exportacion
'           If FACELEC% = 1 Then
'             If TRIM$(Label12) = TRIM$(Control("FACTURA", "PUVEFAEX")) Then
'                Call COMUNI("Imposible Realizar Factura Electrónica por Texto Libre\ Para Cliente de Exportacion ")
'                Label6.Caption = "(Mat.Codif)": MODODOC% = 0
'             End If
'           End If
   ElseIf MODODOC% = 1 Then
        Label6.Caption = "(Mat.Codif)": MODODOC% = 0
   End If
   'End If
   '
   Call Label7_Change ' ACTUALIZA LOS CASILLEROS DE NETO E IVA AL 21% Y 10,5% PARA PODER MODIFICARLOS
   '
   NCLI = XVALO(Text1.TEXT)
   If NCLI > 0 Then
     If MODODOC% > 0 Or TIFADOC% > 2 Then
         PORDESCU.TEXT = ""
       Else
         PORDESCU.TEXT = DESCUCLI$(NCLI)
     End If
   End If
   '
   If MODODOC% = 1 And TIFADOC% = 2 Then
     Check2.Visible = True  ' factura anticipo
     Check2.ZOrder 0
   Else
     Check2.Visible = False
   End If
   '
End Sub

Sub Label7_Click()
   If ULTREG&("!CARDEFAC") > 0 Then Exit Sub
   If Label7.Caption = "Factura:" Then
       TIFADOC% = 3: Label7.Caption = "N.Débito:"
       Label6.Caption = "(T.Libre)": MODODOC% = 1
       '
' validacion suprimida porque s/flavio sí se puede hacer factura
' de texto libre para exportacion
'       If FACELEC% = 1 Then 'POR SI ES FACTURA EXPORTACION
'        If TRIM$(Label12) = TRIM$(Control("FACTURA", "PUVEFAEX")) Then
'          Call COMUNI("Imposible Realizar Factura / N.Debito / N.Credito Electrónica por Texto Libre\ Para Cliente de Exportacion ")
'          Label6.Caption = "(Mat.Codif)": MODODOC% = 0
'        End If
'       End If
'          'f
     ElseIf Label7.Caption = "N.Débito:" Then
       TIFADOC% = 6: Label7.Caption = "N.Crédito:"
       Label6.Caption = "(Mat.Codif)": MODODOC% = 0
       AnulaFacturaComprobante$ = ""
     Else
       TIFADOC% = 2: Label7.Caption = "Factura:"
       Label6.Caption = "(Mat.Codif)": MODODOC% = 0
   End If
   Call NUFACTURA
   '
   NCLI = XVALO(Text1.TEXT)
   If NCLI > 0 Then
     If MODODOC% > 0 Or TIFADOC% > 2 Then
         PORDESCU.TEXT = ""
       Else
         PORDESCU.TEXT = DESCUCLI$(NCLI)
     End If
   End If
   '
   'SI ES TEXTO LIBRE Y FACTURA
   If MODODOC% = 1 And (TIFADOC% = 2 Or TIFADOC% = 3) Then
     Check2.Visible = True  ' factura anticipo
     Check2.ZOrder 0
   Else
     Check2.Visible = False
   End If
   
End Sub

Private Sub Label8_Click()
   On Error Resume Next
   Text13.SetFocus
   On Error GoTo 0
End Sub

Private Sub List1_DblClick()
   If Command14.Enabled = True Then
   
     If FACLOT07.Label15 = "ACTIVO" Then
        FACLOT07.Show 1  'CARGA POR BOBINA
     Else
        Call CARDETFAC
     End If
   End If
   
End Sub

Private Sub List3_DblClick()
   Call DIFPREC2
End Sub

Private Sub MONEMIS_DblClick()
   '
   Static CAMON%, DMONE$(), TCMONE()
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
10   If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         U2& = 0
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           If TRIM$(DMX$) <> "" Then
             VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
             U2& = U2& + 1
             DMONE$(U2&) = DMX$
             TCMONE(U2&) = VHX
           End If
         Next U&
         CAMON% = U2&
         If CAMON% < 1 Then GoTo 10
     End If
   End If
   '
   MONEMIA% = MONEMI%
   MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   MONEMIS = DMONE$(MONEMI%)
'   COEFICO = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   MONEMIS.ForeColor = &H0&
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = &H404080
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   TICAMBIS.Visible = False: Label24.Visible = False
   If MONEMI% > 1 Then
     TICAMBIS.Visible = True: Label24.Visible = True
   End If
   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBX# = TICAMONE(MONEMI%)
   TICAMBIS = FORMATNUM(TICAMBX#, "F9.4")
   FIN& = ULTREG&("!CARDEFAC")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDEFAC", U&)
     CAN = CVS(Mid$(X$, 1, 4))
     PRELI = COEFCON * CVS(Mid$(X$, 421, 4))
     PRECIO = COEFCON * CVS(Mid$(X$, 429, 4))
     IMPO# = CDbl(CAN * PRECIO)
     '
     Call REPLA(X$, MKS$(PRELI), 421, 4)
     Call REPLA(X$, MKS$(PRECIO), 429, 4)
     Call REPLA(X$, MKD$(IMPO#), 433, 8)
     Call GRAREG("!CARDEFAC", X$, U&)
   Next U&
   '
   Call CARLISFAC
   Call CIERRARCH("!CARDEFAC")
   Call CALTOFAC
End Sub

'Private Sub MONEMIS_dblclick()
'   If TRIM$(UCase$(MONEMIS)) = "PESOS" Then
'      MONEMIS = "Dólares"
'     Else
'       MONEMIS = "Pesos"
'   End If
'End Sub

Private Sub PORDESCU_Change()
   Call CALTOFAC
End Sub

Private Sub PORDESCU_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   If DERACCE%("MODESFA/NOMESS", "Modificacion descuento de Factura") < 2 Then
       Text14.SetFocus
       'If List1.ListCount > 0 Then AA1% = DERACCE%("MODESFA", "Modificacion descuento de Factura")
       If MSF.Rows > 1 Then AA1% = DERACCE%("MODESFA", "Modificacion descuento de Factura")
   End If
   '
End Sub

Private Sub PORIVA_Click(Index As Integer)
   PORDESCU.SetFocus
End Sub

Private Sub Text1_Change()
TXTRAZA$ = TXTRAZA$ & "PUNTO 201.1 = " & Now & vbCrLf
TXTRAZA$ = TXTRAZA$ & "ANTES DE BLANFORMU (Text1_Change)= " & Now & vbCrLf

   Call BLANFORMU
TXTRAZA$ = TXTRAZA$ & "DESPUES DE BLANFORMU (Text1_Change)= " & Now & vbCrLf
   
   TXTRAZA$ = TXTRAZA$ & "PUNTO 201.2 = " & Now & vbCrLf
   CLIPROACTIVO = XVALO(Text1)
   '
   If EsPorPedido& > 0 Then
       Call Text1_KeyPress(13)
       Call Command6_Click
   End If
   '
   If TRIM$(Text1) = "" Then A% = PIVACLI%(0)                 ' ES PARA REFRESCAR LAS RUTINAS DEL TIPO RASOCLI, PIVACLI, ETC.
   '
   txtFechaVencimiento = FECHATEX(CalculoVencimientoSegunCondicionPago)
   '
End Sub


Function CalculoVencimientoSegunCondicionPago() As Integer

   Dim FVI%(10), IM#(10)
   '
   NCI& = XVALO(Text1)
   If NCI& < 1 Then Exit Function
   '
   COPAGI% = CPAGCLI%(NCI&)
   Call BUCONPA(LU$, COPAGI%, CoPag$, FREF%, PLA$)
   
   IORIG# = TOBRUFAC#
   FF = FECHANUM(Text9)
   FFS$ = FECHATEX$(FF)
   PLA$ = Left$(PLA$ + String$(6, 0), 6)

   If FREF% = 2 Then FREF% = 1   ' AUN NO FUNCIONA SEMANA DE FACTURA
   If FREF% = 3 Then
     If XVALO(Left$(FFS$, 2)) <= 15 Then FEBAS$ = "15" + Mid$(FFS$, 3)
     If XVALO(Left$(FFS$, 2)) > 15 Then FEBAS$ = "31" + Mid$(FFS$, 3)
   End If
   If FREF% = 1 Then FEBAS$ = FFS$
   If FREF% = 4 Then FEBAS$ = "31" + Mid$(FFS$, 3)
   '
   For i% = 1 To 6: FVI%(i%) = 0: Next i%
   A11 = FECHANUM(FEBAS$): B22 = A11
   NUPAG% = 1
   For i% = 1 To 6
     PLAI% = Asc(Mid$(PLA$, i%, 1))
     If PLAI% > 0 Or i% = 1 Then
       While DIENTRE%(A11, B22) < PLAI%
         B22 = DIPOST(B22)
       Wend
       FVI%(i%) = B22
       NUPAG% = i%
     End If
   Next i%
   '
'   If NUPAG% < 1 Then NUPAG% = 1
'   IM#(1) = (Int(100 * IORIG# / NUPAG% + 0.5)) / 100
'   '
'   For h% = 2 To NUPAG% - 1
'     IM#(h%) = IM#(1)
'   Next h%
'   IM#(NUPAG%) = IORIG#
'   '
'   For h% = 1 To NUPAG% - 1
'     IM#(NUPAG%) = IM#(NUPAG%) - IM#(h%)
'   Next h%
'   '
'   ICANCE# = 0
'   For h% = 1 To NUPAG%
'     If ICANCE# > 0 Then
'       If ICANCE# >= IM#(h%) Then
'           ICANCE# = ICANCE# - IM#(h%)
'           IM#(h%) = 0
'         Else
'           IM#(h%) = IM#(h%) - ICANCE#
'           ICANCE# = 0
'       End If
'     End If
'   Next h%
'   '
'   For h% = 1 To 9
'     FEX = FVI(h%)
'     FEVEN(h% - 1) = FECHATEX$(FEX)
'     IVENCI(h% - 1) = FORMATNUM$(IM#(h%), "F11.2")
'   Next h%
   '
   CalculoVencimientoSegunCondicionPago = FVI(1)
   
End Function

'
Private Sub Text1_DblClick()

  NCLIE = ListaSeleccionClientes()
  If NCLIE > 0 Then
    Text1.TEXT = TRIM$(Str$(NCLIE))
    If XVALO(Text1.TEXT) > 0 Then
      If XVALO(Text1.TEXT) <= 999999 Then
        If REGICLI&(NCLIE) > 0 Then
          Call CARDACLI
          Call Label12_Change
        End If
      End If
    End If
  End If

'  Call SELECHO("DEUDOR1")
'  NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
'  If NCLIE > 0 Then
'    Text1.TEXT = TRIM$(Str$(NCLIE))
'    If XVALO(Text1.TEXT) > 0 Then
'      If XVALO(Text1.TEXT) <= 999999 Then
'        NCLIE = XVALO(Text1)
'        If REGICLI&(NCLIE) > 0 Then
'          Call CARDACLI
''TXTRAZA$ = TXTRAZA$ & "ANTES DE LABEL12_CHANGE= " & Now & vbCrLf
'          Call Label12_Change
''TXTRAZA$ = TXTRAZA$ & "DESPUES DE LABEL12_CHANGE= " & Now & vbCrLf
'        End If
'      End If
'    End If
'  End If
End Sub

Private Sub Text1_GotFocus()
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
   MARCONOTA.Visible = False
End Sub

Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If VALICLIE% > 0 Then
TXTRAZA$ = TXTRAZA$ & "PUNTO 202.1 = " & Now & vbCrLf

      Call BLANFORMU
TXTRAZA$ = TXTRAZA$ & "PUNTO 202.2 = " & Now & vbCrLf
      
      If XVALO(Text1.TEXT) > 0 Then
        If XVALO(Text1.TEXT) <= 999999 Then
          NCLIE = XVALO(Text1)
          If REGICLI&(NCLIE) > 0 Then
            If MARCONOTA.Visible = False Then
TXTRAZA$ = TXTRAZA$ & "PUNTO 202.3 = " & Now & vbCrLf
               Call CARDACLI
TXTRAZA$ = TXTRAZA$ & "PUNTO 202.4 = " & Now & vbCrLf
               Label12_Change 'TEMA REMITO
            End If
            On Error Resume Next
            Command6.SetFocus
            On Error GoTo 0
          End If
        End If
      End If
    End If
  End If
End Sub

Private Sub Text10_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Text10.TEXT = TRIM$(Text10.TEXT)
   Text10.SelStart = 0
   Text10.SelLength = Len(Text10.TEXT)
End Sub

Private Sub Text11_DblClick()
    'NUEVO PARA FACTURAR CON OTRO NUMERO
    If VALICLIE% = 1 Then
      If DERACCE%("FANULI", "Cambiar Manualmente Número de Factura") >= 2 Then
        '
10      NUF = InputBox("Ingrese el Número de Factura", "Cambiar Manualmente Número de Factura")
        NUF1$ = CStr(NUF)
        If TRIM$(NUF1$) = "" Then Exit Sub
        If IsNumeric(NUF) = True And XVALO(NUF) > 0 Then
            'MsgBox NUF
            Text11 = NUF
            NFCAM% = 1
        Else
           Call MENSERR(32, "Número Ingresado No Válido.")
           GoTo 10
        End If
        '
      End If
    Else
        On Error Resume Next
        Text1.SetFocus
        On Error GoTo 0
    End If
    '
End Sub
'
Private Sub Text11_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   If XVALO(Text11.TEXT) > 9 Then
      On Error Resume Next
      Text10.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text11_LostFocus()
   NUMEPRE$ = ""
   For KK% = 1 To Len(Text11.TEXT)
     CCX% = Asc(Mid$(Text11.TEXT, KK%, 1))
     If CCX% >= 48 Then
       If CCX% <= 57 Then
         NUMEPRE$ = NUMEPRE$ + Chr$(CCX%)
       End If
     End If
   Next KK%
   '
   Call NUFACTURA
   '
   If NUMEPRE$ <> TRIM$(Text11.TEXT) Then
     Text11.TEXT = NUMEPRE$
     Call MENSERR(24, "Número de Factura Ajustado")
   End If
   '
End Sub

Private Sub Text12_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text13_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text16_Change()
   Label16 = Left$(ECOARCH("ECUECON", Text16), 20)
End Sub

Private Sub Text16_DblClick()
   Call Command12_Click
End Sub

''///****NUEVO***//////////
Private Sub TEXT17_CHANGE()
   Call CALTOFA2
End Sub
'
Private Sub Text17_GotFocus()
   Text17.SelStart = 0
   Text17.SelLength = Len(Text17)
End Sub

Private Sub Text17_LostFocus()
    '
    Text17 = FORMATNUM$(XVALO(Text17), "F11.2")
    '
End Sub

Private Sub Text19_Change()
'  NTRA = XVALO(Text19)
'  If NTRA > 0 Then
'    If NTRA < 32767 Then
'      NTRAI% = NTRA
'      LENTREGA(2) = ECOARCH$("TRANSPOR", MKI$(NTRAI%))
'    End If
'  Else
'    LENTREGA(2) = "" ' MODIFICADO PARA QUE SI EL TRANSPORTE ES 0 LIMPIE EL TEXT
'  End If
   'LENTREGA(0) = ""
   'LENTREGA(1) = ""
   'LENTREGA(2) = ""
   TRANACT% = 0
   '\\\OT-07-0040-RG-FIN///
   '\\\OT-6-422-FG-05/10/06
   NTRA = XVALO(Text19)
   If NTRA > 0 Then
      If NTRA < 32767 Then
         NTRAI% = NTRA
         LENTREGA(2) = DENOTRANS$(NTRAI%)
         If LENTREGA(2) <> "" Then
            TRANACT% = NTRAI%
         End If
         '\\\OT-07-0040-RG-02/02/07///
         If TRIM$(LENTREGA(0)) = "" Then ' SOLO SI NO TIENE DOMICILIO DE ENTREGA, UTILIZA DOMICILIO Y LOCALIDAD DEL TRANSPORTE
            LENTREGA(0) = DOMITRANS$(NTRAI%)
            LENTREGA(1) = LOCATRANS$(NTRAI%)
         End If
         '\\\OT-07-0040-RG-FIN///
      End If     '
   End If

End Sub

Private Sub Text19_DblClick()
   '\\\OT-06-0570-RG-27/12/06///
   Call Command24_Click
   '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text2_gotfocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text20_Change()
  NSUC& = XVALO(Text20)
  Label27 = ""
  Text19 = ""
'  If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
'    If NSUC& > 0 Then
'      If NSUC& <= ULTREG&("!SUCENTRE") Then
'        XX$ = REGLEIDO$("!SUCENTRE", NSUC&)
'        Label27 = TRIM$(Mid$(XX$, 3, 62))
'        LENTREGA(0) = TRIM$(Mid$(XX$, 65, 48))
'        LENTREGA(1) = TRIM$(Mid$(XX$, 113, 48))
'      End If
'    End If
'    Text19 = NUTRANCLI%(XVALO(Text1), XVALO(Text20))
'  End If
  NCLIE = XVALO(Text1)
  CONDI$ = "NUME_SUC=" & NSUC& & " AND NUME_CLI=" & NCLIE
  Label27 = TRIM$(SAFETEXT$(VALOBADA("SUCENTRE", "DENO_SUC", CONDI$)))
  LENTREGA(0) = TRIM$(SAFETEXT$(VALOBADA("SUCENTRE", "DOMI_SUC", CONDI$)))
  LENTREGA(1) = TRIM$(SAFETEXT$(VALOBADA("SUCENTRE", "LOCA_SUC", CONDI$)))
  If Label27 <> "" Then Text19 = NUTRANCLI%(NCLIE, XVALO(Text20))
  '
End Sub

Private Sub Text20_DblClick()
   Call Command25_Click
End Sub

Private Sub TEXT21_DblClick(Index As Integer)
   '\\\OT-06-0570-RG-27/12/06///
   If Index = 0 Then
      Call Command21_Click(Index)
   ElseIf Index = 1 Then
      Call Command21_Click(Index)
   End If
   '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub TEXT21_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      Exit Sub
   End If
   TEXT21(Index).TEXT = TRIM$(TEXT21(Index).TEXT)
   TEXT21(Index).SelStart = 0
   TEXT21(Index).SelLength = Len(TEXT21(Index).TEXT)
End Sub

Private Sub Text22_LostFocus()
    If IsNumeric(Text22) = True Then
       TX# = XVALO(Text22)
       Text22 = FORMATNUM$(TX#, "F5.2")
    End If
End Sub

Private Sub Text23_Change()
   Call CALTOFAC
End Sub

Private Sub Text25_Change()
    Call CALTOFA2
End Sub
Private Sub Text26_Change()
    Call CALTOFA2
End Sub
Private Sub Text25_GotFocus()
   If VALICLIE% < 1 Then GoTo 99
   '
   If TIFADOC% < 3 Then GoTo 98
   Exit Sub
   '
98 On Error Resume Next
   PORDESCU.SetFocus
   On Error GoTo 0
   Exit Sub
   '
99 On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text26_GotFocus()
   If VALICLIE% < 1 Then GoTo 99
   '
   If TIFADOC% < 3 Then GoTo 98
   Exit Sub
   '
98 On Error Resume Next
   PORDESCU.SetFocus
   On Error GoTo 0
   Exit Sub
   '
99 On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub
Private Sub Text25_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Text25 = TRIM$(FORMATNUM$(XVALO(Text25), "F8.2"))
  End If
End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Text26 = TRIM$(FORMATNUM$(XVALO(Text26), "F8.2"))
  End If
End Sub
Private Sub Text25_LostFocus()
    Text25 = TRIM$(FORMATNUM$(XVALO(Text25), "F8.2"))
End Sub

Private Sub Text26_LostFocus()
    Text26 = TRIM$(FORMATNUM$(XVALO(Text26), "F8.2"))
End Sub
Private Sub Text3_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text4_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text5_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub TEXT21_Change(Index As Integer)
    '
    If XVALO(TEXT21(Index).TEXT) < 0 Or XVALO(TEXT21(Index).TEXT) > 255 Then
       TEXT21(Index).TEXT = ""
    End If
    '
    If Index = 0 Then
        If XVALO(TEXT21(Index).TEXT) < 1 Then
            Eco21(Index).TEXT = "CONCERTADOS FIJOS"
            ARCHIPRE$ = "MASTER": LIPRACTI% = 0
          Else
            Eco21(Index).TEXT = ECOARCH$("LISTAS", MKI$(XVALO(TEXT21(Index).TEXT)))
            ARCHIPRE$ = TRIM$(TEXT21(0).TEXT)
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$: LIPRACTI% = XVALO(Mid$(ARCHIPRE$, 6, 3))
        End If
        If NOCAMONE% < 1 Then
           If MONELISTA%(XVALO(TEXT21(Index))) = 2 Then
                MONEMIS = "Dólares"
                MONEMIS.ForeColor = &H404080
                MONEMI% = 2
                TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
                MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
                TICAMBIS.Visible = True
                Label24.Visible = True
              Else
                MONEMIS = "Pesos"
                MONEMIS.ForeColor = &H0&
                MONEMI% = 1
                TICAMBIS = "1.0000"
                MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
                TICAMBIS.Visible = False
                Label24.Visible = False
           End If
        End If
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(XVALO(TEXT21(Index).TEXT)))
      ElseIf Index = 2 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(XVALO(TEXT21(Index).TEXT)))
    End If
    '
End Sub

Private Sub Command21_Click(Index As Integer)
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
    If Index = 0 Then
        ARCHEC$ = "LISTAS": LOCA% = 2
      ElseIf Index = 1 Then
        ARCHEC$ = "CONPAG": LOCA% = 2
      ElseIf Index = 2 Then
        ARCHEC$ = "VENDEDOR": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    '\\\OT-06-0570-RG-27/12/06///
    Call CommSel(ARCHEC$, Command21(Index), TEXT21(Index))
    '\\\OT-06-0570-RG-FIN///
    'Call SELECHO(ARCHEC$)
    'TEXT21(Index).TEXT = VALACT1$(ARCHEC$)
    '
End Sub
   '
Function VALICLIE%()
   VCC% = 0
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 999999 Then
       NCLI = XVALO(Text1.TEXT)
       If rasocli$(NCLI) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function

Sub BLANFORMU()

   TASACLICK = 0
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
'''''   Text7.TEXT = "" 'se esta limpiando mas abajo
'''''   Text8.TEXT = "" 'se esta limpiando mas abajo
   Text10.TEXT = ""
   Text12.TEXT = ""
   Text14.TEXT = ""
   Text16 = ""
   Label16 = ""
   Label26 = ""
   Text17 = "" 'LIMPIA RET-IVA
   IEMBALA = ""
   Text19 = ""
   Text20 = ""
   Text22 = ""
   Label25 = ""
   Label28 = ""
   Ltota(10) = "Importe Total Neto"
   Text23 = ""
   '
   If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
     Call SETULTREG("!SUCENTRE", 0)
     Call FRESHECHO("!SUCENTRE")
   End If
   '
   '\\\OT-5-579-FG-16/08/05
   Ltota(3).Caption = "TOTAL DOCUMENTO"
   Ltota(3).Left = 1560
   Text18.Visible = False
   '\\\OT-5-579-FG-FIN
   '\\\OT-08-0672-OD-26/08/08
   TIFADOC% = 2: Label7.Caption = "Factura:"
   Label6.Caption = "(Mat.Codif)": MODODOC% = 0
   Label13.Caption = "A": PIVADOC% = 1: VARIDOC% = 1
   '\\\OT-08-0672-OD-FIN
   '
   ' PONE LETRA POR OMISIÓN
   Label13.Caption = "A"
   '\\\OT-5-506-FG-02/09/05
   ' PONE PREFIJO POR OMISION AL 0001
'   UPREFI$ = TRIM$(CONTROL$("FACTURA", "PREFIFA"))
    CLAVE$ = "PREFAC-" & Label13.Caption '  MODIFICADO PARA GUARDAR PREFIJO DE FACTURA POR ESTACION
    UPREFI$ = TRIM$(Control$(CLAVE$, IDENTER$))
   If UPREFI$ = "" Then UPREFI$ = "0001"
   'Label12.Caption = UPREFI$ 'TEMA REMITO
   '\\\OT-5-506-FG-FIN
   '
   Text6.TEXT = ""
   List1.Clear
   MSF.Rows = 1
   DOCUDIF = ""
   '
   Picture5.Visible = False
   List3.Clear
   '
   '\\\OT-8-0852-OD-02/09/05
   'SE INVIRTIO LA POSICION DE LOS DOS FOR QUE ESTAN A CONTINUACION POR LA ENTRADA
   'AL EVENTO CHANGE DEL TEXT
   PORDESCU.TEXT = ""
   For KK% = 6 To 0 Step -1
      Total(KK%).TEXT = ""
   Next KK%
   TLABEL28 = ""
   'EL TEXT15 QUE SE ENCONTRABA MAS ARRIBA EN EL CODIGO SE BAJO POR
   'QUE PROVOCABA QUE APARECIERA LA VENTANA CON EL MENSAJE
   '"IMPOSIBLE FACTURAR POR EL LIMITE DE CREDITO " ,
   Text15 = ""
   Text24 = ""
   Text25 = ""
   Text26 = ""
   'bajado aqui por que si no no tenia lectura del limite de credito en el evento change
   'de la caja de texto Total donde revisa el mismo
   Text7.TEXT = "" 'limite de credito
   Text8.TEXT = "" 'disponible
   '
   For KKI% = 0 To 2
      TEXT21(KKI%).TEXT = ""
      Eco21(KKI%) = ""
      BASIVA(KKI%) = ""
      LENTREGA(KKI%) = ""
   Next KKI%
  '\\\OT-8-0852-OD-FIN

   '
   MARCONOTA.Visible = False
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Call SETULTREG("!CARDEFAC", 0)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     On Error Resume Next
     Kill LUCOM$ + "1CARDEFA.WKF"
     On Error GoTo 0
   End If
   'Call NUFACTURA
   '
   FACTLISTAEMPAQUE& = 0
   FACTUANTICIPO% = 0
   ESPROFORMA& = 0
   '
   Label31 = ""
   '
   Check2.Value = 0
   Check2.Visible = False ' factura anticipo
   '
   If Control$("FACTURA", "SALIDA") = "FELECT" Then
      PTOVTA$ = TRIM$(Control("FACTURA", "PUVENELE"))
      If PTOVTA$ <> "" Then Label12 = PTOVTA$
   End If
   '
   ' *** PONE EN CERO TODOS LOS INDICES DEL VECTOR
   If NetoMultipleImputacion% = 1 Then
        For i% = 1 To UBound(VentaNetaPorCuentaContable#)
            VentaNetaPorCuentaContable#(i%) = 0
        Next i%
   End If
   '
   ' MODALIDAD FACTURA DE CREDITO
   chkModoCredito.Value = 0
   AplicacionFacturaCredito$ = ""
   TipoAplicacionFacturaCredito$ = ""
   '
   ' *** INICIALIZA EXPORTA SIMPLE
   ModoExportaSimple% = 0
   PuntoVentaExportaSimple$ = ""
   NumeroExportaSimple$ = ""
   CAEExportaSimple$ = ""
   VtoCAEExportaSimple% = 0
   '
   MONEMI% = 1  ' PESOS
   TICAMBIS = "1.0000"
   TICAMBIS.Visible = False
   
   PercepcionIBMisiones_BienDeUso = 0

End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    '
TXTRAZA$ = TXTRAZA$ & "ENTRO A CARGAR DATOS CLIENTE CARDACLI= " & Now & vbCrLf
    NCLIE = XVALO(Text1.TEXT): Text1.Refresh
    If NCLIE > 0 Then
      If rasocli$(NCLIE) <> "" Then
        If InStr(EMPREAC$, "DOSIVAC") > 0 Then Call MUESTRAVENCIDAS(NC%, 15)
      
        Text2.TEXT = rasocli$(NCLIE): Text2.Refresh
        Text3.TEXT = DOMICLI$(NCLIE): Text3.Refresh
        Text4.TEXT = CPOSCLI$(NCLIE): Text4.Refresh
        Text5.TEXT = LOCACLI$(NCLIE): Text5.Refresh
        Label26 = CODICLI$(NCLIE)
        Text7.TEXT = CSTRING$(MKS$(LIMICRE(NCLIE)), 3, 10, 2, 2)
          Text7.Refresh
        Text12.TEXT = TRIM$(POSIVAC$(NCLIE)) + " - CUIT " + CuitCli$(NCLIE)
          Text12.Refresh
        '
        VEND% = VENDCLI%(NCLIE)
        FIN& = ULTREG&("VENDEDOR")
        For i& = 1 To FIN&
            TX$ = REGLEIDO$("VENDEDOR", i&)
            If Asc(Mid$(TX$, 1, 1)) = VEND% Then
               COMIS# = CVS(Mid$(TX$, 44, 4))
               Text22 = FORMATNUM$(COMIS#, "F5.2")
            End If
        Next i&
        '
        '\\\OT-5-579-FG-16/08/05
        YY$ = TRIM$(POSIVAC$(NCLIE))
        If YY$ = "COM.EXT." Then
          Ltota(3).Caption = "Total:"
          Ltota(3).Left = 250
          Text18 = "FOB"
          Text18.Visible = True
        End If
        '\\\OT-5-579-FG-FIN
        '
        'CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
        CREDISPO# = XVALO(Text7) - SALDCLI#(NCLIE) - CHEPEACRE#(NCLIE)
        Text8.ForeColor = RGB(0, 0, 0)
        If CREDISPO# < 0 Then Text8.ForeColor = RGB(255, 0, 0)
        Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 10, 2, 2)
          Text8.Refresh
        '
        TEXT21(0).TEXT = Str$(LIPRCLI%(NCLIE))
        TEXT21(1).TEXT = Str$(CPAGCLI%(NCLIE))
        TEXT21(2).TEXT = Str$(VENDCLI%(NCLIE))
        '
        If XVALO(TEXT21(2).TEXT) < 1 Then
           Eco21(2).TEXT = "GERENCIA"
        End If
        '
        NCLIE = XVALO(Text1.TEXT)
        If NCLIE > 0 Then
          If MODODOC% > 0 Then
              PORDESCU.TEXT = ""
            Else
              PORDESCU.TEXT = DESCUCLI$(NCLIE)
          End If
        End If
        PIVADOC% = PIVACLI%(NCLIE)
        Call CARLETDOC: 'Call NUFACTURA '\\\OT-06-0051-WAR-15/02/06///
        Refresh
        '
        PPVI% = PIVADOC%
        If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
        Text16 = Control$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
        '
        'verificacion si es factura de exportacion valida si tiene token
        If (PPVI% = 5 Or PPVI% = 7) And Control$("FACTURA", "SALIDA") = "FELECT" Then
        CARPETOK$ = "C:\FLEXOFT\FACELEC\"
        ARCHIVO_TOKEN$ = "TOKEN_WSFEX.TXT"
            If EXISTE%(CARPETOK$ + ARCHIVO_TOKEN$) = 1 Then
                 FechArch = FileSystem.FileDateTime(CARPETOK$ + ARCHIVO_TOKEN$)
                 FECHARCHIVO% = CVINT(FechArch)
                 FECHHOY% = HOY(HO$)
                 If FECHARCHIVO% < FECHHOY% Then
                    Screen.MousePointer = 1
                    Call MENSERR(24, "Token Vencido. Vuelva a Generarlo")
                    Text1 = ""
                    Exit Sub
                 End If
            Else
                 Screen.MousePointer = 1
                 Call MENSERR(24, "Debe Generarse el Archivo 'TOKEN_WSFEX.TXT'")
                 Text1 = ""
                 Exit Sub
            End If
        End If
        '
        '\\\OT-5-506-FG-02/09/05
        ' PARA QUE TOME EN CUENTA EL PREFIJO DE LAS FACTURAS
        ' DE EXPORTACION
'        If (PPVI% = 5 Or PPVI% = 7) Then
'          Label12.Caption = PREFEXPO$()
'          CLAVE$ = "PREFAC-" & Label13.Caption
'          UPREFI$ = TRIM$(CONTROL$(CLAVE$, IDENTER$))
'          If UPREFI$ = "" Then UPREFI$ = PREFEXPO$()
'          If UPREFI$ = "" Then UPREFI$ = "0001"
'          Label12.Caption = UPREFI$
'        End If
        '\\\OT-5-506-FG-FIN
        'VALIDO SI TIENE CONFIGURADO P.VTA FACT.ELECT. EXPO, SI NO DEJA EL CONFIGURADO EN MODO HABITUAL
        ' If PPVI% = 5 And FACELEC% = 1 Then
        'SE AGREGO PARA QUE CONSIDERE LA POSICION DE IVA 7
        If (PPVI% = 5 Or PPVI% = 7) Then
          If Control$("FACTURA", "SALIDA") = "FELECT" Then
            PUVEX$ = TRIM$(Control("FACTURA", "PUVEFAEX"))
            If PUVEX$ <> "" Then
              Label12 = PUVEX$
            End If
          Else
            Label12.Caption = PREFEXPO$()
          End If
        End If
        '
        If (PPVI% = 0 Or PPVI% = 3 Or PPVI% = 4 Or PPVI% = 6) And Control$("FACTURA", "SALIDA") <> "FELECT" Then ' PARA EL CASO QUE LA FACTURA SEA B
            CLAVE$ = "PREFAC-" & Label13.Caption
            UPREFI$ = TRIM$(Control$(CLAVE$, IDENTER$))
            If UPREFI$ = "" Then UPREFI$ = "0001"
            Label12.Caption = UPREFI$
        End If
        '
        If PPVI% = 5 And InStr(1, EMPREAC$, "FERVI") > 0 Then ' PARA EL CASO QUE LA FACTURA SEA B
            Call MONEMIS_DblClick
        End If
        '
        'Call NUFACTURA 'TEMA REMITO'\\\OT-06-0051-WAR-15/02/06///
        '
        Call CARGALISEL("SUCENTRE", "SUCENTRE", "NUME_SUC/I5;DENO_SUC/A32", "NUME_CLI=" & NCLIE)
'        If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
'          '\\\OT-7-010-RG-12/01/07
'          Call FILTERFILE("SUCENTRE", "", 5, MKI$(NCLIE))
'          '\\\OT-7-010-RG-FIN
'        End If
        '
        For KKI% = 0 To 2
          LENTREGA(KKI%) = LUENTRECLI$(NCLIE, KKI%)
        Next KKI%
        '
        NTRAN% = NUTRANCLI%(NCLIE)
        If NTRAN% > 0 Then
          Text19 = TRIM$(Str$(NTRAN%))
        End If
        '
        Text6 = ANOTACLI$(NCLIE)
        MARCONOTA.Visible = True
        MARCONOTA.Refresh
        '
        LMCR# = LIMICRE(NCLIE)
        If VERLICRE% = 1 Or LMCR# < 0 Then 'CONSIDERA CLIENTE SUSPENDIDO,
'          LMCR# = LIMICRE(NC%)
          SDEU# = SALDCLI#(NCLIE)
          CHPA# = CHEPEACRE#(NCLIE)
          If SDEU# + CHPA# >= LMCR# Then
            Call COMUNI("Imposible Facturar a este Cliente.\  \Límite de Crédito Asignado ya está Superado.\  \Informe al Supervisor de Ventas.")
            Call BLANFORMU
            On Error Resume Next
            Text1.SetFocus
            On Error GoTo 0
            GoTo 99
          End If
        End If
        '
        If EXEDELIMCC%(NCLIE, 0, DIFER) > 0 Then
            LIMCCTE$ = TRIM$(FORMATNUM$(LCred_CCte(NCLIE), "f12.2"))
            TX$ = "Límite de Crédito Asignado en C.Cte. $ " & LIMCCTE$ & " ya está Superado."
            If DERACCE%("AUTOFACT", "Autorizar Facturas con Exceso Credito") < 2 Then
                Call COMUNI("Imposible Facturar a este Cliente.\  \Límite de Crédito Asignado en C.Cte. ya está Superado.\  \Informe al Supervisor de Ventas.")
                Call BLANFORMU
                On Error Resume Next
                Text1.SetFocus
                On Error GoTo 0
            Else
              If COMALTER%(TX$ + "\   \Puede Autorizar la Facturación\con Crédito Excedido.\\Cancelar\Autorizar Factura") < 2 Then
                Call BLANFORMU
                On Error Resume Next
                Text1.SetFocus
                On Error GoTo 0
              End If
            End If
        End If
        ' SE INTRODUCE UN XVALO PORQUE POREMBAL DEVUELVE UN TEXTO Y LA COMPARACION SE HACE COMO SI FUESE NUMERICO
        If XVALO(POREMBAL(NCLIE)) > 0 Then
          CUEMBA$ = Control$("SVD451S", "CUEVE-12")
          If CUEMBA$ = "" Or CUECOINT%(CUEMBA$) < 1 Then
            Call MENSERR(24, "Falta Definir Cuenta Contable\Correspondiente a Embalajes.\   \FLEXOFT Utilizará la Cuenta de Ventas Netas.")
          End If
        End If
        '
        PVCL% = PIVADOC%
        If PVCL% = 1 Or PVCL% = 2 Or PVCL% = 3 Or PVCL% = 4 Or PVCL% = 6 Then
           CUTT$ = CuitCli$(NCLIE)
           If VALICUIT%(CUTT$) < 1 Then
              Call MENSERR(24, "Imposible Emitir Documento Fiscal.\  \Posición de IVA del Cliente\Requiere C.U.I.T. Válido.")
              Text1 = ""
           End If
        End If
        '
'        ' *** VALIDA POSICION DE IVA CONTRA AFIP
'        If Control("FACTURA", "VAPOSIVA") = "SI" Then
'            Dim oValidaCliente As New DatosEmpresaAfip
'            condicionIvaAfip% = oValidaCliente.ObtenerIdCondicionIvaAfip(CUTT$)
'            If condicionIvaAfip% >= 0 And condicionIvaAfip% <> PIVADOC% Then
'                denominacionPosIvaCliente$ = ECOARCH$("CAPOSIV", Chr$(PIVADOC%))
'                denominacionPosIvaAfip$ = ECOARCH$("CAPOSIV", Chr$(condicionIvaAfip%))
'                opcionElegida% = COMALTER%("Se Ha Encontrado que la Posición de Iva del Cliente\" & Text2 & " \No Coincide con la Registrada en Afip\ \Actualmente registrado como: " & denominacionPosIvaCliente$ & "\Según Afip: " & denominacionPosIvaAfip$ & "\ \¿Desea Considerar la que Posee Afip?\\Si, Considerar Afip\No, Utilizar la del Cliente\Cancelar")
'                MARCONOTA.Visible = True
'                If opcionElegida% < 1 Or opcionElegida% > 3 Then Text1 = ""
'                If opcionElegida% = 1 Then
'                    Call ACTUREGISTRO("DEUDOR12", "PIVA_CLI", "NUME_CLI=" & CStr(NCLIE), condicionIvaAfip%, REGAF&)
'                    Text1 = ""
'                    Text1 = CStr(NCLIE)
'                    Call Text1_KeyPress(13)
'                ElseIf opcionElegida% = 3 Then
'                    Text1 = ""
'                End If
'            End If
'        End If
        '
      End If
    End If
99  Screen.MousePointer = 1
    '.
TXTRAZA$ = TXTRAZA$ & "SALIO DE CARGAR DATOS CLIENTE (CARDACLI) = " & Now & vbCrLf
    
End Sub
    
Sub CARLETDOC()
   '
   Label13.Caption = "A"
   VARIDOC% = 1
   RNIDOC% = 0
   '
   If PIVADOC% = 0 Or PIVADOC% = 3 Or PIVADOC% = 4 Or PIVADOC% = 6 Then
        Label13.Caption = "B"
        VARIDOC% = 2
        '
      ElseIf PIVADOC% = 2 Then
        RNIDOC% = 1
        '
      ElseIf PIVADOC% = 5 Or PIVADOC% = 7 Then
        Label13.Caption = "X"
        VARIDOC% = 3
        If TRIM$(LETREXPO$) <> "" Then
          Label13.Caption = LETREXPO$
        End If
        If LETREXPO$ = "A" Then
          VARIDOC% = 1
          Label13.Caption = "A"
        End If
    End If
    '
End Sub
'
Function FLINUMERA%()         ' FECHA DE NUEVA NUMERACION
   '
   Static FLINU%
   '
   If FLINU% = 0 Then
      FLINU% = FECHANUM("01/01/99")
      If ULTREG&("FORMFAC") > 0 Then
         FFS$ = REGLEIDO$("FORMFAC", 2&)
         FLINU% = CVI(Mid$(FFS$, 19, 2))
      End If
   End If
   '
   FLINUMERA% = FLINU%
   '
End Function
'
Function NUMECORR%()          ' NUMERACION CORRELATIVA O INDEPENDIENTE
   '
   Static NUMECO%
   '
   If NUMECO% = 0 Or NUMECO% = (-1) Then
      NUMECO% = 1
      QTRX$ = Control$("NUMEFACT", "NUMECORR")
      If QTRX$ = "INDEPEN" Then
           NUMECO% = (-1)
         ElseIf QTRX$ = "UNIFICA" Then
           NUMECO% = (-2)
      End If
   End If
   '
   'CASO QUE ES FACTURA ELECTRONICA Y DE EXPORTACION MARCA CORRELATIVIDAD INDEPENDIENTE
   If Control$("FACTURA", "SALIDA") = "FELECT" And Label13.Caption = "E" Then NUMECO% = (-1)
   '
   NUMECORR% = NUMECO%
   '
End Function
'
Function MONUMERA$()
   '
   Static MONUM$
   '
   If MONUM$ = "" Then
      MONUM$ = "AUTO"
      If TRIM$(Control$("", "NUMFACTU")) = "MANUAL" Then
         MONUM$ = "MANUAL"
      End If
   End If
   '
   MONUMERA$ = MONUM$
   '
End Function
'
Sub NUREMIFA_HAS(NROREMIT&, PUNTODEVENTA$)
   '
   VAMOVIX$ = Label13
   Call TOMANUMDELCONFIS("RT", VAMOVIX$, PUNTODEVENTA$, NumeComp$)
   NROREMIT& = XVALO(NumeComp$)
   '
End Sub
'
Sub NUFACTURA()
   '
    Call CREAR_STORED_PROCEDURE

   MsgBox "entre en  nufactura"
TXTRAZA$ = TXTRAZA$ & "ENTRO EN NUFACTURA = " & Now & vbCrLf

   DATEDOC% = FECHANUM(Text9.TEXT)
   If DATEDOC% < FLINUMERA% Then
      Call MENSERR(24, "Imposible Facturar en la Fecha Elegida.\  \La Fecha de Factura Debe Ser POSTERIOR\a la Fecha de Nueva Numeración.")
      Call CIERRARCH("*.*")
      Call FINAL("")
   End If
   '
   Screen.MousePointer = 11
   '
   
   OPMOVIX$ = "FC"
   If TIFADOC% = 3 Then OPMOVIX$ = "ND"
   If TIFADOC% = 6 Then OPMOVIX$ = "NC"
   VAMOVIX$ = Label13
   '
   
   ' SI ES IMPRESORA FISCAL, TOMA EL NUMERO DEL HASAR
   If CONTROFIS% = 2 Then
      If MUECOMHAS% < 2 Then
         If VAMOVIX$ = "A" Or VAMOVIX$ = "B" Then ' PARA EXCLUIR EXPORTACION
           Call TOMANUMDELCONFIS(OPMOVIX$, VAMOVIX$, PUNTODEVENTA$, NumeComp$)
           Label12 = PUNTODEVENTA$
           NPP& = XVALO(NumeComp$)
           Text11.TEXT = TRIM$(Str$(NPP&))
           '
           Call NUREMIFA_HAS(NROREMIT&, PUNTODEVENTA$)
           If NROREMIT& > XVALO(Text10.TEXT) Then
              Text10.TEXT = TRIM$(Str$(NROREMIT&))
           End If
           GoTo 999
         End If
      End If
   End If
      MsgBox "ANTES DE FLINUME"

   FLINUME = CDbl(CVDAT(FLINUMERA%))
   MsgBox "DESPUES DE FLINUME"

   '
   FEXUL = DATEDOC%
   FLIDIAS = CDbl(CVDAT(DIASANTES(FEXUL, 15)))
   MsgBox "DESPUES  DE FLIDIAS"

   '//////////*******NOTA*******////////
   'Con posterioridad se detecto un Error en la Impòrtacion
   '(VARIMOVIM(A,B), SE IMPORTO EN VEZ DE LA LETRA CON EL OPCIMOVIX$(ND,NC).
   'Por tal motivo el sistema no toma en cuenta los numero
   'de NC-ND inferiores al 01/11/04.
   'Pero si lo hace a futuro. es decir yo puedo reiniciar numeración
   'al 01/11/04
   
  ' FLIMI = FLIDIAS: MODORAPIDO% = 1
  ' If FLINUME > FLIMI Then
  '    FLIMI = FLINUME
  '    MODORAPIDO% = 0
  ' End If
   
   '
'22
   Call APERBASDAT("CABAN")
   'jandy sql esta rutina debe ser igual
   'ASELE$ = "SELECT NumeComp FROM CajaBanc "
   '
   FLIMI = FETEXCOR1$(FLINUMERA)
   '
TXTRAZA$ = TXTRAZA$ & "ANTES DE VALIDAR FACELEC > 0 = " & Now & vbCrLf
   
   'SI ES FACTURA ELECTRONICA ASIGNA NUMERO 99999999 A LA CAJA DE TEXTO
   If FACELEC% > 0 Then
      NPP& = "99999999"
      GoTo 150
   End If
   
   ' *** MODALIDAD EXPORTA SIMPLE
   If ModoExportaSimple% > 0 Then
17    OBX$ = OBJPLA$("CBTEXSIM", VDEF$)
      If OBX$ <> "" Then
        PuntoVentaExportaSimple$ = Right("0000" & TRIM$(Mid$(OBX, 1, 5)), 4)
        NumeroExportaSimple$ = Right("00000000" & TRIM$(Mid$(OBX, 6, 8)), 8)
        CAEExportaSimple$ = TRIM$(Mid$(OBX, 14, 20))
        VtoCAEExportaSimple% = CVI(Mid$(OBX, 34, 2))
        
        If XVALO(PuntoVentaExportaSimple$) = 0 Then
            Call MENSERR(24, "Punto de Venta Ingresado Inválido")
            VDEF$ = OBX$
            GoTo 17
        End If
        
        If XVALO(NumeroExportaSimple$) = 0 Then
            Call MENSERR(24, "Número de Comprobante Ingresado Inválido")
            VDEF$ = OBX$
            GoTo 17
        End If
        
        If XVALO(CAEExportaSimple$) = 0 Then
            Call MENSERR(24, "CAE Ingresado Inválido")
            VDEF$ = OBX$
            GoTo 17
        End If
        
        TipoComprobante$ = OPMOVIX$
        Comprobante$ = TipoComprobante$ + "-" + PuntoVentaExportaSimple$ + "-" + NumeroExportaSimple$ + "-" + Label13
        If CantRegistros("CAJABANC", "codicom_lar_exsimple='" & Comprobante$ & "'") > 0 Then
            Call COMUNI("Imposible Continuar.\ \Este Documento Ya se Encuentra Procesado.")
            PuntoVentaExportaSimple$ = ""
        
            NPP& = 1
        Else
            NPP& = XVALO(NumeroExportaSimple$)
        End If
        
        GoTo 150
      Else
        PuntoVentaExportaSimple$ = ""
        
        NPP& = 1
        
        GoTo 150
      End If
   End If
   '
   MsgBox "ANTES DE SQL CAJABANC"
   ASELE$ = "SELECT NumeComp FROM CajaBanc WITH(NOLOCK)"
   If OPMOVIX$ = "ND" Or OPMOVIX$ = "NC" Then
     If NUMECORR% = (-2) Then    ' Unificada debitos y creditos
       ASELE$ = ASELE$ + "WHERE ((TipoMovim = 'FVEN' AND (OpciMovim = 'ND' OR OpciMovim = 'NC')) "
       ASELE$ = ASELE$ + "OR (TipoMovim = 'FCOM' AND (OpciMovim = 'DP' OR OpciMovim = 'CP'))) "
       GoTo 25
     End If
   End If
   '
   MsgBox "SALE DE SQL CAJABANC"
   ASELE$ = ASELE$ + "WHERE TipoMovim = 'FVEN' "
   If NUMECORR% <> 1 Then
       ASELE$ = ASELE$ + "AND OpciMovim = '" & OPMOVIX$ & "' "
     Else
       '       ASELE$ = ASELE$ + "AND (OpciMovim = 'FC' OR OpciMovim = 'ND' OR OpciMovim = 'NC') "
       '       //////****NUEVA CONDICION PARA CALCULAR PROX. Nº FAC***////
       '       ESTA CONDICION LA AGREGUE  PARA QUE TOME EN CTA. LOS Nº
       '       DE COMPROBANTES ARRUINADOS: OR OpciMovim = 'ANUL'
       ASELE$ = ASELE$ + "AND (OpciMovim = 'FC' OR OpciMovim = 'ND' OR OpciMovim = 'NC' OR OpciMovim = 'ANUL') "
   End If
   '
25 ASELE$ = ASELE$ + "AND VariMovim = '" & VAMOVIX$ & "' "
   If XVALO(Label12) <= 1 Then
       ASELE$ = ASELE$ + "AND (NumeTalo = 0 OR NumeTalo = 1) "
     Else
       ASELE$ = ASELE$ + "AND NumeTalo = " & XVALO(Label12) & " "
   End If
   ASELE$ = ASELE$ + "AND FechEmisi >= '" & FLIMI & "' "
   ASELE$ = ASELE$ + " ORDER BY NUMECOMP DESC "
   '
   NUMAXI& = 0
   Dim MIRS As ADODB.Recordset
   Set MIRS = New ADODB.Recordset
   ' USAR EL ERROR AQUI CON CAPTURA DE ERROR PARA ACELERAR LA BUSQUEDA
   'Set MIRS = FLEXCONN.Execute(FILTSQL$(ASELE$))
   MIRS.CursorLocation = adUseServer
   Set MIRS = READSET(ASELE$)
'   If Not IsNull(MIRS!maximo) Then
    NUMAXI& = XVALO(MIRS!NumeComp)
'   End If
   MIRS.Close
   Set MIRS = Nothing
   '
99 If Visible = True Then
     If NUMAXI& < 1 Then
       NUMAXI& = XVALO(InputBox$("Ingrese Número del Próximo Comprobante del Tipo Elegido", "Incicialización de Talonarios")) - 1
       If NUMAXI& < 0 Then GoTo 99
     End If
   End If
   NPP& = NUMAXI& + 1
   
150 Text11.TEXT = TRIM$(Str$(NPP&))
TXTRAZA$ = TXTRAZA$ & "VIENE ACA POR QUE ES F.ELECTRONICA = " & Now & vbCrLf

   Text11.Refresh
   '
   If TIFADOC% = 2 Then
 TXTRAZA$ = TXTRAZA$ & "ENTRO EN IF TIFADOC% = 2 = " & Now & vbCrLf
  
      ' If TRIM$(Text10.TEXT) = "" Or FACELEC% = 1 Then       'Para facturas electronica actualiza el Nº de remito
      If TRIM$(Text10.TEXT) = "" Or FACELEC% > 0 Then      'Para facturas electronica actualiza el Nº de remito
        ' If VERIFSALIDA("NOMESS") < 3 Then NPVX$ = Label12                                     'segun el punto de venta seleccionado
        NPVX$ = Control$("", "PREFIREM") 'AGREGADO PARA QUE RESPETE EL NUMERO DE REMITO
        If NPVX$ = "" Then NPVX$ = "0001" 'CONFIGURADO EN SETUP REMITO, EN EL CASO FACTURA COMUN
        '
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        If FACELEC% > 0 And FACELEC% < 3 Then NPVX$ = Label12 ' SOLO FACTURA ELECTRONICA                                   'segun el punto de venta seleccionado
        If PIVADOC% = 5 Or PIVADOC% = 7 Then
          If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "TANIT") > 0 Or InStr(EPAC$, "GABAL") > 0 Or InStr(EPAC$, "FERVI") > 0 Or InStr(EPAC$, "AHP") > 0 Or InStr(EPAC$, "J Y M") > 0 Or InStr(EPAC$, "CEYEQ") > 0 Or InStr(EPAC$, "EYM") > 0 Then
            NPVX$ = "0001"
          End If
        End If
        '
        'VALIDACION AGREGADA PARA DOSIVAC PARA QUE MANTENGA EL PREFIJO 0001 PARA EL REMITO
        If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "TANIT") > 0 Or InStr(EPAC$, "GABAL") > 0 Or InStr(EPAC$, "FERVI") > 0 Or InStr(EPAC$, "AHP") > 0 Or InStr(EPAC$, "J Y M") > 0 Or InStr(EPAC$, "CEYEQ") > 0 Or InStr(EPAC$, "EYM") > 0 Or InStr(EPAC$, "DAVTECNICA") > 0 Then NPVX$ = "0001"
        If InStr(EPAC$, "ASEMBLI") > 0 Then NPVX$ = Control$("", "PREFIREM")
        '
        
TXTRAZA$ = TXTRAZA$ & "ANTES DE NUREMIFA EN NUFACTURA = " & Now & vbCrLf
      MsgBox "ANTES DE NUREMIFA"
     
        Call NUREMIFA(NROREMIT&, NPVX$)
           MsgBox "DESPUES DE NUREMIFA"

        ' If NROREMIT& > XVALO(Text10.TEXT) Or FACELEC% = 1 Then      'Para facturas electronica actualiza el Nº de remito
        If NROREMIT& > XVALO(Text10.TEXT) Or FACELEC% > 0 Then     'Para facturas electronica actualiza el Nº de remito
          Text10.TEXT = TRIM$(Str$(NROREMIT&))                      'segun el punto de venta seleccionado
        End If
        '
TXTRAZA$ = TXTRAZA$ & "DESPUES DE NUREMIFA EN NUFACTURA = " & Now & vbCrLf
        ' esto se puso por reclamo de dosivac de que a veces se 'traba'
        ' y repite el numero de remito anterior - Ana el 22 de setiembre de 2008
'995     If VERINURE% < 0 Then
'           NROREMIT& = NROREMIT& + 1
'           Text10.TEXT = TRIM$(Str$(NROREMIT&))
'           GoTo 995
'        End If
        '
      End If
      Screen.MousePointer = 1
      Exit Sub
     Else
      Text10.TEXT = ""
   End If
   '
   
999 Screen.MousePointer = 1
   '
End Sub
'
Sub TOMANUMDELCONFIS(TDC$, LTDC$, PREFIFA$, NumeComp$)
     Call FORHAS320.TOMANUMDELCONFIS(TDC$, LTDC$, PREFIFA$, NumeComp$)
End Sub
'
Function VERINUFA%(NFCY$)
   '
   Static FLINUMERA%
   Static NUMECORR%
   '
   If CONTROFIS% = 2 Then
       VERINUFA% = 1
       Exit Function
   End If
   '
   OPMOVIX$ = Left$(NFCY$, 2)
   '
   Screen.MousePointer = 11
   If FLINUMERA% = 0 Then
      FLINUMERA% = FECHANUM("01/01/99")
      If ULTREG&("FORMFAC") > 0 Then
         FFS$ = REGLEIDO$("FORMFAC", 2&)
         FLINUMERA% = CVI(Mid$(FFS$, 19, 2))
      End If
   End If
   '
   If NUMECORR% = 0 Then
      NUMECORR% = 1
      QTRX$ = Control$("NUMEFACT", "NUMECORR")
      If QTRX$ = "INDEPEN" Then
           NUMECORR% = (-1)
         ElseIf QTRX$ = "UNIFICA" Then
           NUMECORR% = (-2)
      End If
   End If
   '
   FLIMI = CDbl(CVDAT(FLINUMERA%))
22 Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CajaBanc "
   SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' and CODICOM_LAR = '" & NFCY$ & "') "
   If NUMECORR% = 1 Then
      NFCY1$ = "FC" + Mid$(NFCY$, 3)
      NFCY2$ = "ND" + Mid$(NFCY$, 3)
      NFCY3$ = "NC" + Mid$(NFCY$, 3)
      SQLX$ = SQLX$ + " OR (TipoMovim = 'FVEN' and (CODICOM_LAR = '" & NFCY1$ & "' OR CODICOM_LAR = '" & NFCY2$ & "' OR CODICOM_LAR = '" & NFCY3$ & "')) "
   End If
   If OPMOVIX$ = "ND" Or OPMOVIX$ = "NC" Then
     If NUMECORR% = (-2) Then    ' Unificada debitos y creditos
       SQLX$ = SQLX$ + "OR (TipoMovim = 'FCOM' AND CODICOM_LAR = '" & NFCY$ & "') "
     End If
   End If
   
   'jandy sql
   
   'Set NUFAXI = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   Dim NUFAXI As ADODB.Recordset
   Set NUFAXI = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   On Error Resume Next
   NUFAXI.MoveFirst
   If Err Then
       VERINUFA% = 1
     Else
       Do While Not NUFAXI.EOF
         'If CDbl(NUFAXI.FechEmisi) >= FLIMI Then
         If CDbl(NUFAXI!FECHEMISI) >= FLIMI Then
            VERINUFA% = (-1)
            GoTo 90
         End If
       NUFAXI.MoveNext
       Loop
       VERINUFA% = 1
   End If

90 On Error GoTo 0
NUFAXI.Close
Set NUFAXI = Nothing

   '
End Function
'
Function VERINURE%()
   '
   Static PREFIREM$, FLIM%, CMO$
   '
   If CONTROFIS% = 2 Then
       VERINURE% = 1
       Exit Function
   End If
   '
   If CMO$ = "" Then
     CMO$ = TRIM$(Control$("", "CODREFAC"))
     If CMO$ = "" Then CMO$ = "RT"
   End If
   '
   If PREFIREM$ = "" Then
     PREFIREM$ = Control$("", "PREFIREM")
     If PREFIREM$ = "" Then PREFIREM$ = "0001"
     'ESTO SE AGREGA PORQUE LOS REMITOS DE IMPRESORA FISCAL TIENEN EL MISMO
     'PUNTO DE VENTA QUE LA FACTURA
     If Control$("FACTURA", "SALIDA") = "IFISCAL" Then PREFIREM$ = Label12
   End If
   '
   'ESTO SE AGREGA PORQUE LOS REMITOS DE LAS
   'FACTURAS ELECTRONICAS TIENEN EL PUNTO DE VENTA DE LA FACTURA
   ' If FACELEC% = 1 And VERIFSALIDA("NOMESS") < 3 Then PREFIREM$ = Label12
   If FACELEC% > 0 And FACELEC% < 3 Then PREFIREM$ = Label12
   'AGREGADO PARA QUE RESPETE EL PUNTO DE VENTA EN DOSIVAC
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "TANIT") > 0 Or InStr(EPAC$, "GABAL") > 0 Or InStr(EPAC$, "FERVI") > 0 Or InStr(EPAC$, "AHP") > 0 Or InStr(EPAC$, "J Y M") > 0 Or InStr(EPAC$, "CEYEQ") > 0 Or InStr(EPAC$, "EYM") > 0 Or InStr(EPAC$, "DAVTECNICA") > 0 Then PREFIREM$ = "0001"
   If InStr(EPAC$, "ASEMBLI") > 0 Then PREFIREM$ = Control$("", "PREFIREM")
   'FIN DOSIVAC
   '
   If FLIM% < 1 Then
     FENUE$ = Control$("", "FENUREFA")
     If FENUE$ = "" Then FENUE$ = "18/04/03"
     FLIM% = FECHANUM(FENUE$)
     If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
   End If
   'FLINUME = CDbl(CVDAT(FLIM%))
   'FLINUME = FECHATEX(FLIM%)
   FLINUME = FETEXCOR1$(FLIM%)
   'cambio
   
   '
   NUPU = XVALO(PREFIREM$) ' TOMA EL PUNTO DE VENTA
   CONDI$ = "FechMov >= '" & FLINUME & "' "
   CONDI$ = CONDI$ + " AND CodiMovi = '" & CMO$ & "' "
   CONDI$ = CONDI$ + " AND NUMPVTA = " & NUPU
   If CantRegistros&("MOVISTO", CONDI$, "NOMESS") > 0 Then
     CONDI$ = CONDI$ + " AND DOPRINUM = " & XVALO(Text10) ' NUMERO DE REMITO DE PANTALLA
     If InStr(1, EMPREAC$, "GABAL") > 0 Then
        CONDI$ = CONDI$ & " AND Codiempr =" & CodiEmp%
     End If
     EXISTEREMI% = CantRegistros&("MOVISTO", CONDI$, " NOMESS")
     If EXISTEREMI% > 0 Then
       VERINURE% = (-1)
       Exit Function
     End If
      REMICOR$ = "RT." + TRIM$(Text10)
      REMILAR$ = Text10: While Len(REMILAR$) < 8: REMILAR$ = "0" + REMILAR$: Wend
      REMILAR$ = CMO$ + "-X-" + PREFIREM$ + "-" + REMILAR$
      GoTo 80 ' REMITO DE REPARACION
   End If
   '
'fin cambio
   '
   REMICOR$ = "RT." + TRIM$(Text10)
   REMILAR$ = Text10: While Len(REMILAR$) < 8: REMILAR$ = "0" + REMILAR$: Wend
   REMILAR$ = CMO$ + "-X-" + PREFIREM$ + "-" + REMILAR$
   '
   Call APERBASDAT("STOCKS")
   'JANDY SQL
   SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)" 'TEMA REMITO
   SQLX$ = SQLX$ + " WHERE FechMov >= '" & FLINUME & "' "
   SQLX$ = SQLX$ + " AND CodiMovi = '" & CMO$ & "' "
   SQLX$ = SQLX$ + " AND DoPriLar = '" & REMILAR$ & "' "
   '
   Dim MOVITMP As ADODB.Recordset
   Set MOVITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With MOVITMP
      If Not (.EOF And .BOF) Then
          VERINURE% = (-1)
          Exit Function
      End If
   End With
   MOVITMP.Close
   Set MOVITMP = Nothing
   '
80 SQLX$ = "SELECT * FROM MOVIREPA "
   SQLX$ = SQLX$ + " WHERE FechMov >= '" & FLINUME & "' "
   SQLX$ = SQLX$ + " AND CodiMovi = '*R' "
   SQLX$ = SQLX$ + " AND DoPriLar = '" & REMILAR$ & "' "
   '
   Set MOVITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With MOVITMP
      If Not (.EOF And .BOF) Then
          VERINURE% = (-1)
          Exit Function
      End If
   End With
   MOVITMP.Close
   Set MOVITMP = Nothing
   '
   VERINURE% = 1
   '
End Function
'
Function VERINUFAVIE%(NFCY$)
    '
    Static NUMECORR%
    '
    OPMOVIX$ = Left$(NFCY$, 2)
    '
    If NUMECORR% = 0 Then
      NUMECORR% = 1
      QTRX$ = Control$("NUMEFACT", "NUMECORR")
      If QTRX$ = "INDEPEN" Then
           NUMECORR% = (-1)
         ElseIf QTRX$ = "UNIFICA" Then
           NUMECORR% = (-2)
      End If
    End If
    '
    Call APERBASDAT("CABAN")
    With CajaBanc
       SQLX$ = "(TipoMovim = 'FVEN' and CODICOM_LAR = '" & NFCY$ & "') "
       If OPMOVIX$ = "ND" Or OPMOVIX$ = "NC" Then
         If NUMECORR% = (-2) Then    ' Unificada debitos y creditos
           SQLX$ = SQLX$ + "OR (TipoMovim = 'FCOM' AND CODICOM_LAR = '" & NFCY$ & "') "
         End If
       End If
       '
       VERINUFAVIE% = 1
       .FindFirst SQLX$
       If .NoMatch = False Then
         VERINUFAVIE = (-1)            'Nº DE FACTURA PREEXISTENTE
       End If
    End With
    '
End Function

Private Sub Text6_GotFocus()
   On Error Resume Next
   'Text1.SetFocus
   Command6.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text7_DblClick()
    '\\\OT-08-00732-OD-29/08/08///
    If DERACCE%("MODLICRE", "Modificar Límite de Crédito") >= 2 Then
       NC% = XVALO(Text1)
       LimCre$ = Text7
       LiCreNue = ROUND#(XVALO(InputBox$("Ingrese el Nuevo Límite", "Límite de Crédito", LimCre$)), 2)
       If LiCreNue < 0.01 Then
            Call MENSERR(20, "Ingreso No Válido")
            Exit Sub
         ElseIf LiCreNue >= 10000000 - 0.01 Then
            Call COMUNI("Importe Supera el Valor Máximo Permitido")
            Exit Sub
         Else
            SQLX$ = "SELECT  Lcre_Cli FROM DEUDOR12 "
            SQLX$ = SQLX$ + "WHERE Nume_Cli = " & NC% & ""
            '
            Dim LIMTMP As ADODB.Recordset
            Set LIMTMP = New ADODB.Recordset
25          On Error Resume Next
            LIMTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 25
            End If
            On Error GoTo 0
            '
            With LIMTMP
               If Not .EOF Then
                 !Lcre_Cli = LiCreNue
                 .Update
               End If
            End With
            LIMTMP.Close
            Set LIMTMP = Nothing
            Text7 = FORMATNUM(LiCreNue, "F10.2")
            Text7.Refresh
            '
            CREDISPO# = XVALO(Text7) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
            Text8.ForeColor = RGB(0, 0, 0)
            If CREDISPO# < 0 Then Text8.ForeColor = RGB(255, 0, 0)
            Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 10, 2, 2)
            Text8.Refresh
       End If
    End If
End Sub

Private Sub Text7_GotFocus()
   On Error Resume Next
   Text14.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text8_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text9_GotFocus()
   On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
End Sub

Private Sub TICAMBIS_Change()
   If TICAMBIS.Enabled = True Then
     TICACH(MONEMI%) = XVALO(TICAMBIS)
   End If
   TICAMBX# = 1
   If MONEMI% > 1 Then
     TICAMBX# = XVALO(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(XVALO(TICAMBIS), "F9.4")
End Sub

Private Sub TICAMBIS_DblClick()
   'Call CARTAMONED
   '
   '\\\OT-07-0086-RG-06/03/07///
   If DERACCE%("COTILOCA", "Ajustar Tipo de Cambio Local") >= 2 Then
      ENTRATICA.Frame4.Caption = "Cotización Local para " + MONEMIS.Caption
      ENTRATICA.Text4 = TICAMBIS.Caption
      ENTRATICA.Text4 = FORMATNUM(XVALO(ENTRATICA.Text4), "F10.4")
      ENTRATICA.Show 1
      If TRIM$(ENTRATICA.Text4) <> "" Then
        TICAMBIS = TRIM$(ENTRATICA.Text4)
      End If
   End If
   '\\\OT-07-0086-RG-FIN///
   '
End Sub

Private Sub Timer1_Timer()
   '
   If CAMBIOVALGRILLA% > 0 Then
      Screen.MousePointer = 11
      Call CALTOFAC
      CAMBIOVALGRILLA% = 0
      Screen.MousePointer = 1
   End If
   '
End Sub

Sub PRIFACTURA()
    '
    If (CONTROFIS% = 1 Or CONTROFIS% = 2) Then ' OBVIA LOS COMPROBANTES 'E' DE LA IMP. FISCAL
       'imprime por controlador / impresora fiscal
       Call PRICONFIS
       If VARIDOC% <> 3 Then   ' PARA QUE EN FACTURA DE EXPORTACION SIGA A
          Exit Sub             ' EMITIR FACTURA POR IMPRESORA COMÚN
       End If
    End If
    '
    Static AGREDESP%
    TICAMBX# = XVALO(TICAMBIS)
    '
    If REPRIFAC% = 1 Then GoTo 10
    '
    ' ESTO ES PARA ACELERAR BUSQUEDA DEL NUMERO DE SERIE
    CMO$ = TRIM$(Control$("", "CODREFAC")):   If Len(CMO$) <> 2 Then CMO$ = "RT"
    NCLIE = XVALO(Text1.TEXT)
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE & " "
    SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND (DoSecLar is null or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    SQLX$ = SQLX$ + " AND FechMov >= '" & DESELEC & " ' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HASELEC & "' "
    Dim CONSITEMP As ADODB.Recordset
    Set CONSITEMP = New ADODB.Recordset
    '
10  CAHOJAS% = 0
    FORIPRE$ = TRIM$(Control$("", "FORFAC-" + Label13.Caption))
    TIPOCOM$ = TRIM$(Label6)
    If TIPOCOM$ = "(T.Libre)" Then
      If Label13.Caption <> "E" Then   ' EXCLUYE FACTURAS DE EXPORTACION
        FORIPRE$ = TRIM$(Control$("", "FACTEXLI"))
        If UCase$(Label13) = "B" Then
          FORIPRE$ = TRIM$(Control$("", "FACTELIB"))
        End If
      End If
    End If
    '\\\OT-06-0038-WAR-01/02/06///
    FORIALT$ = ""
    If TIFADOC% = 3 Or TIFADOC% = 6 Then 'SI NC O ND
      If TIFADOC% = 3 Then ' SI NOTA DE DEBITO
        If UCase$(Label13.Caption) = "A" Then
          FORIALT$ = Control$("", "FORNDE-A") 'ND-A
          ElseIf UCase$(Label13.Caption) = "B" Then
          FORIALT$ = Control$("", "FORNDE-B") 'ND-B
        End If
      End If
      '
      If TIFADOC% = 6 Then ' SI NOTA DE CREDITO
        If UCase$(Label13.Caption) = "A" Then
          FORIALT$ = Control$("", "FORNCR-A") 'NC-A
          ElseIf UCase$(Label13.Caption) = "B" Then
          FORIALT$ = Control$("", "FORNCR-B") 'NC-B
        End If
      End If
    End If
    If TRIM$(FORIALT$) <> "" Then
      FORIPRE$ = FORIALT$
    End If
    ''\\\OT-06-0038-WAR-FIN///
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Controlador de Impresión")
       Exit Sub
    End If
    '
    'SE AGREGA IF PARA QUE VALIDE SOLO SI NO ES FACTURA ELECTRONICA
    If FACELEC% < 1 Then
       If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
         Call MENSERR(24, "Imposible Imprimir Documento.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
         Exit Sub
       End If
    End If
    '
    CODFOR$ = FORIPRE$
    '
    Screen.MousePointer = 11
    ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
    If ITPORHOJA% < 1 Then
      ITPORHOJA% = 32767
    End If
    '
    CANCELPRINT% = 0
    FECHDOC$ = Text9.TEXT
    NUMEDOC$ = TRIM$(Text11.TEXT)
    While Len(NUMEDOC$) < 8
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    NUMEDOC$ = Label12.Caption + "-" + NUMEDOC$
    '
    TIPODOC$ = TRIM$(REPLACAR$(Label7.Caption, ":", "")): CODOCUM$ = "FAC"
    If TIFADOC% = 3 Then TIPODOC$ = "Nota de Debito": CODOCUM$ = "DEB"
    If TIFADOC% = 6 Then TIPODOC$ = "Nota de Credito": CODOCUM$ = "CRE"
    CPAG% = XVALO(TEXT21(1).TEXT)
    '
    DESTIDOCNUE = XVALO(Text1.TEXT)
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    'CODTABU1$(4) = "PRE-LIST"
    'CODTABU1$(5) = "PORDESCU"
    'CODTABU1$(6) = "PRE-NETO"
    'CODTABU1$(7) = "NET-ITEM"
    'CODTABU1$(8) = "FECH-REC"
    CODTABU1$(9) = "ORD-ITEM"
    CODTABU1$(10) = "UNIMED"
    CODTABU1$(11) = "REF-MAT1"
    CODTABU1$(12) = "REF-MAT2"
    CODTABU1$(13) = "REF-MAT3"
    CODTABU1$(14) = "N-PLANO"
    CODTABU1$(15) = "N-ORIGIN"
    CODTABU1$(16) = "N-OCOMPR"
    CODTABU1$(17) = "N-REMITO"
    CODTABU1$(18) = "N-BOLREC"
    CODTABU1$(19) = "FECH-REC"
    CODTABU1$(20) = "PRE-LIST"
    CODTABU1$(21) = "PORDESCU"
    CODTABU1$(22) = "PRE-NETO"
    CODTABU1$(23) = "NET-ITEM"
    CODTABU1$(24) = "ALICUIVA"
    CODTABU1$(25) = "IVA-ITEM"
    CODTABU1$(26) = "BRU-ITEM"
    '
    CACOLTAB1% = 26
    '
    CAITAB1% = 0
    CAITCRU% = 0
    '
    SUMATOTA# = XVALO(Total(0).TEXT)
    IDESCUEN# = XVALO(Total(1).TEXT)
    '
    COPIASREM% = XVALO(FINFACTUR07.Text2(1))
    If COPIASREM% < 1 Then Text10.TEXT = ""
    '
    NroRemi$ = TRIM$(Text10.TEXT)
    If NroRemi$ <> "" Then
      While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
      PREFIREM$ = Control$("", "PREFIREM")
      If PREFIREM$ = "" Then PREFIREM$ = "0001"
      'ESTO SE AGREGA PORQUE LOS REMITOS DE LAS
      'FACTURAS ELECTRONICAS TIENEN EL PUNTO DE VENTA DE LA FACTURA
      ' If FACELEC% = 1 And VERIFSALIDA("NOMESS") < 3 Then PREFIREM$ = Label12
      If FACELEC% > 0 And FACELEC% < 3 Then PREFIREM$ = Label12
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "ASEMBLI") > 0 Then PREFIREM$ = Control$("", "PREFIREM")
      
      '
      NroRemi$ = PREFIREM$ + "-" + NroRemi$
    End If
    '
    If CONSICLI% = 1 Then
      NroRemi$ = NROREMICO$
    End If
    '
    COEFIVAINCLIEU# = 1
    If XVALO(Total(2).TEXT) > 0.005 Then
      COEFIVAINCLIEU# = (XVALO(Total(2).TEXT) + XVALO(Total(3).TEXT) + XVALO(Total(4).TEXT) + XVALO(Total(5).TEXT)) / XVALO(Total(2).TEXT)
    End If
    '
    If TRIM$(Text14.TEXT) <> "" Then
      NUOCOMPRA$ = TRIM$(Text14.TEXT)
    End If
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = Label13.Caption
    CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = NroRemi$
    CODPARAM$(3) = "NRO-OCOM":  DOCPARAM$(3) = NUOCOMPRA$
    CODPARAM$(4) = "NRO-AENT":  DOCPARAM$(4) = NROPEDIDO$
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = ""
      NUVENDEX$ = AJUSTI$(FORMATNUM$(XVALO(TEXT21(2)), "I3"), 4)
    CODPARAM$(6) = "VENDEDOR": DOCPARAM$(6) = NUVENDEX$ + Eco21(2).TEXT
    CODPARAM$(7) = "COND-PAG": DOCPARAM$(7) = Eco21(1).TEXT
    CODPARAM$(8) = "POR-IVA": DOCPARAM$(8) = PORIVA(0).Caption
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = ""
    CODPARAM$(10) = "POR-IVA1": DOCPARAM$(10) = PORIVA(1).Caption
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = ""
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = ""
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = ""
    'CODPARAM$(4) = "NRO-OCOM": DOCPARAM$(4) = Text10.TEXT
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = ""
    CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = TRIM$(PORDESCU.TEXT)
    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = ""
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = ""
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = ""
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = ""
    CODPARAM$(23) = "TOTA-DOL": DOCPARAM$(23) = ""    ' Valor en PESOS equivalente a dolares
    '
    CODPARAM$(20) = "DOMI-TRA": DOCPARAM$(20) = LENTREGA(0).TEXT
    CODPARAM$(21) = "DESTINO": DOCPARAM$(21) = LENTREGA(1).TEXT
    CODPARAM$(22) = "TRANSPOR": DOCPARAM$(22) = LENTREGA(2).TEXT
    CODPARAM$(24) = "MONEDEMI": DOCPARAM$(24) = MONEMIS
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = ""
    CODPARAM$(26) = "SUB-TOTA": DOCPARAM$(26) = ""
    CODPARAM$(27) = "PERC-IVA": DOCPARAM$(27) = "" 'PERCIVA
    CODPARAM$(28) = "FECH-VEN": DOCPARAM$(28) = Format$(FVENCIM, "dd/mm/yy") ' FECHA DE VENCIMIENTO
      On Error Resume Next
      FEXO = FECHANUM(FECHDOC$)
      FEXU = CVINT(FVENCIM)
      CADIVEN = DIENTRE(FEXO, FEXU)
      On Error GoTo 0
    CODPARAM$(32) = "FECH-LAN": DOCPARAM$(32) = FORMATNUM$(CADIVEN, "F5.0")
    '\\\OT-5-579-FG-16/08/05
    CODPARAM$(29) = "TIPO-VAL": DOCPARAM$(29) = TRIM$(Text18)          ' VALOR FOB
    CODPARAM$(30) = "TICAMBIO": DOCPARAM$(30) = TICAMBIS               ' TIPO DE CAMBIO
    CODPARAM$(31) = "CTAC-CLI": DOCPARAM$(31) = CONTACTCLI$(DESTIDOCNUE) ' CONTACTO
    '\\\OT-5-579-FG-FIN
    '\\\OT-06-0125-WAR-03/04/06///
    CODPARAM$(33) = "NUBULTOS": DOCPARAM$(33) = TRIM$(FINFACTUR07.Text2(2))          ' BULTOS
    CODPARAM$(34) = "PES-CONJ": DOCPARAM$(34) = TRIM$(FINFACTUR07.Text2(11)) + " Kg" ' PESO NETO
    CODPARAM$(35) = "PES-COMP": DOCPARAM$(35) = TRIM$(FINFACTUR07.Text2(4)) + " Kg"  ' PESO BRUTO
    CODPARAM$(36) = "BANC-CHE": DOCPARAM$(36) = TRIM$(FINFACTUR07.Text2(5))          'BANCO EXTERIOR
    CODPARAM$(37) = "DOCU-APL": DOCPARAM$(37) = TRIM$(FINFACTUR07.Text2(8))          ' BANCO LOCAL
    CODPARAM$(38) = "NUME-CHE": DOCPARAM$(38) = TRIM$(FINFACTUR07.Text2(10))         ' CREDITO
    CODPARAM$(39) = "DEN-CUEN": DOCPARAM$(39) = TRIM$(FINFACTUR07.Text2(9))          ' GUIA DE IMPORTACION
    CODPARAM$(40) = "REF-ASIE": DOCPARAM$(40) = TRIM$(FINFACTUR07.Text2(6))          ' PUERTO DE EMBARQUE
    CODPARAM$(41) = "DET-PASE": DOCPARAM$(41) = TRIM$(FINFACTUR07.Text2(7))          ' DESTINO
    CODPARAM$(42) = "MEDIDA-1": DOCPARAM$(42) = TRIM$(IEMBALA)                      ' gasto de embalaje
    CODPARAM$(43) = "ZONA-VEN": DOCPARAM$(43) = TRIM$(Str$(GRUCOCLI%(NCLIE)))
    '\\\OT-06-0125-WAR-FIN///
    'AGREGUEMH04-10
    CODPARAM$(44) = "NRO-PASE": DOCPARAM$(44) = TRIM$(Label27.Caption) 'Sucursal
    '
    CODPARAM$(45) = "ITO-GRAV": DOCPARAM$(45) = ""
    CODPARAM$(46) = "ITO-NOGR": DOCPARAM$(46) = ""
    CODPARAM$(47) = "ITO-GRA1": DOCPARAM$(47) = ""
    CODPARAM$(48) = "ITO-GRA2": DOCPARAM$(48) = ""
    '
    CODPARAM$(49) = "IMP-DEBE": DOCPARAM$(49) = FINFACTUR07.Text2(13)  ' FLETE
    CODPARAM$(50) = "IMP-HABE": DOCPARAM$(50) = FINFACTUR07.Text2(12)  ' SEGURO
    CODPARAM$(51) = "TOT-HABE": DOCPARAM$(51) = FORMATNUM(TOTATOTA#, "F13.2")
    CODPARAM$(52) = "REF-MAT4": DOCPARAM$(52) = FINFACTUR07.Label12
    CODPARAM$(53) = "COD-CUEN": DOCPARAM$(53) = Text23
        DESC$ = TRIM$(PORDESCU.TEXT)
        PODES = COEFDESCU(DESC$)
        DESC0$ = TRIM$(Text23) ' DECUENTO VARIOS REALIZADO PARA GABAL
        PODES0 = COEFDESCU(DESC0$)
        TTT1# = (XVALO(Total(0).TEXT) * PODES / 100)
        IMPONET# = XVALO(Total(0).TEXT) - TTT1#
        IMPODEGA# = (IMPONET# * PODES0 / 100)
    CODPARAM$(54) = "IMPO-PEN": DOCPARAM$(54) = TRIM$(FORMATNUM(IMPODEGA#, "F12.2"))
    '
    CAPARDOC% = 54 '32 '\\\OT-06-0125-WAR-03/04/06///
    '
    Call SETULTREG("!OBSERVOF", 0)
    '
    If FINFACTUR07.Text4 <> "" Then
      JR& = 0
      FORMOBSE$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
      ANCHOB% = XVALO(Mid$(FORMOBSE$, 2)): If ANCHOB% < 32 Then ANCHOB% = 32
      Call FRATEXTO(FINFACTUR07.Text4, ANCHOB%)
      For KKII% = 1 To CARETEX%
        JR& = JR& + 1
        Call GRAREG("!OBSERVOF", RETEX$(KKII%), JR&)
      Next KKII%
      Call SETULTREG("!OBSERVOF", JR&)
      Call CIERRARCH("!OBSERVOF")
    End If
    'SE MODIFICA VARIABLE POR QUE POVERTI% ES UNA VARIABLE PUBLICA
    POVERTI1% = XVALO(TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/POSINVER"))))
    '
    PRINTERPORT$ = PORTUFAC$
    CACOPDOC% = XVALO(FINFACTUR07.Text2(0).TEXT)
    '
    NOCONFIRM% = 0
    If TRIM$(UCase$(Control$("FACTUREM", "IMPRENUM"))) <> "AUTONUME" Then
      If Control$("FACTUREM", "FORMULAR") <> "HOJASUEL" Then
        NOCONFIRM% = 1
      End If
    End If
    '
    HOJASUEL% = 0
    If Control$("FACTUREM", "FORMULAR") = "HOJASUEL" Then
       HOJASUEL% = 1
    End If
    '
    ARCHILIB$ = "!1CARDEFA"
    FIN& = ULTREG&("!CARDEFAC")
    SUBTOTA# = 0: SUBTOTB# = 0
  '
  '\\\OT-06-0339-MH-24/08/06///
    'control es la funcion que devuelve en este caso SI o NO
    'y asi reemplaza el cod orig o no seg lo selecc en configuracion
    'remstan es el modulo en que esta el proyecto
    'replaori es una clave que refleja la configuracion
    REPLAORI% = 0
      If Control$("REMSTAN", "REPLAORI") = "SI" Then
        REPLAORI% = 1
      End If
     'va fuera del for para q no se ejecute innecesariamente
  '\\\OT-06-0339-MH-FIN///
    '
    NOJA% = 0
    If REPRIFAC% = 1 Then
       PROJA% = XVALO(InputBox$("Imprime desde Hoja Número", "Primera Hoja de Factura", "1"))
    End If
    If PROJA% < 1 Then PROJA% = 1
    '
    For U& = 1 To FIN&
      If (U& Mod ITPORHOJA%) = 1 Then
        ' PONE ENCABEZADO
        CAITAB1% = 0
        CAHOJAS% = CAHOJAS% + 1
      End If
      '
      XXX$ = REGLEIDO$("!CARDEFAC", U&)
      CIPRE% = CVI(Mid$(XXX$, 5, 2))
      CAPEX$ = Str$(CVS(Left$(XXX$, 4)))       ' cantidad
      '
      NPEDIDD& = XVALO(TRIM$(Mid$(XXX$, 371, 16)))
      SUBCO$ = SUBCODI$(NPEDIDD&, CIPRE%)
      '
      CAITAB1% = CAITAB1% + 1
      '
      '\\\OT-06-0339-MH-24/08/06///
      CODDX$ = CODEXT$(CIPRE%) + SUBCO$ ' en coddx$ trae el codigo externo a partir de cipre(int)
      If REPLAORI% = 1 Then
        CODDX$ = CODORIG$(CIPRE%, NCLIE) ' + SUBCO$
      ' se pasa el codigo interno, mas el nº cliente(NCLIE) y trae el codigo externo solicitado
      End If
      '\\\OT-06-0339-MH-FIN///
      '
      '
      PLANOX$ = VALOBADA("MASTER", "NPlano", "CODINT = " & CODINT%(CODDX$))
      '
      TETABU1$(1, CAITAB1%) = CODDX$
      TETABU1$(2, CAITAB1%) = Mid$(XXX$, 7, 64)
        ' SE AGREGA CONDICION PARA QUE ACTUE UNICAMENTE EN DOSIVAC
        ' Y MANTENER LA FUNCIONALIDAD PARA EL RESTO DE LOS USUARIOS
        If PIVADOC% = 5 And InStr(EMPREAC$, "DOSIVA") > 0 Then
          DEALTE$ = TRIM$(DESCRIT2$(CIPRE%))
          If DEALTE$ <> "" Then
             TETABU1$(2, CAITAB1%) = DEALTE$
          End If
        End If
      TETABU1$(3, CAITAB1%) = CAPEX$
      'TETABU1$(4, CAITAB1%) = PRELI$
      'TETABU1$(5, CAITAB1%) = DESCX$
      'TETABU1$(6, CAITAB1%) = PREUX$
      'TETABU1$(7, CAITAB1%) = ITOTX$
      'TETABU1$(8, CAITAB1%) = FENTX$
      TETABU1$(9, CAITAB1%) = Str$(U&)
      TETABU1$(10, CAITAB1%) = Mid$(XXX$, 71, 4)    ' unidad de medida
      TETABU1$(11, CAITAB1%) = Mid$(XXX$, 75, 64)   ' referencia 1
      'TETABU1$(12, CAITAB1%) = Mid$(XXX$, 139, 64)  ' referencia 1 ' SE AGREGO ABAJO
      TETABU1$(13, CAITAB1%) = Mid$(XXX$, 203, 64)  ' referencia 1
      TETABU1$(14, CAITAB1%) = PLANOX$
      'TETABU1$(14, CAITAB1%) = Mid$(XXX$, 315, 24)  ' numero de plano
      TETABU1$(15, CAITAB1%) = Mid$(XXX$, 339, 16)  ' numero original
      TETABU1$(16, CAITAB1%) = Mid$(XXX$, 355, 16)  ' numero de o.compra
      TETABU1$(17, CAITAB1%) = Mid$(XXX$, 387, 16)  ' numero de remito
      TETABU1$(18, CAITAB1%) = Mid$(XXX$, 403, 16)  ' numero de doc-recepcion
      TETABU1$(19, CAITAB1%) = FECHATEX$(CVI(Mid$(XXX$, 419, 2))) ' fecha de recepcion
      '
      FACIVITEM = 1: If Label13.Caption = "B" Then FACIVITEM = 1 + (CVS(Mid$(XXX$, 441, 4))) / 100
      LPRE% = XVALO(TEXT21(0).TEXT)
      If LPRE% > 0 Then
        TETABU1$(20, CAITAB1%) = Str$(FACIVITEM * CVS(Mid$(XXX$, 421, 4))) ' precio de lista
        TETABU1$(21, CAITAB1%) = Str$(CVS(Mid$(XXX$, 425, 4))) ' PORCENTAJE DE DESCUENTO
      End If
      '
      TETABU1$(22, CAITAB1%) = Str$(FACIVITEM * CVS(Mid$(XXX$, 429, 4))) ' precio unitario
      TETABU1$(23, CAITAB1%) = Str$(CVD(Mid$(XXX$, 433, 8))) ' total neto
      TETABU1$(24, CAITAB1%) = Str$(CVS(Mid$(XXX$, 441, 4))) ' alicuota iva
      TETABU1$(25, CAITAB1%) = Str$(CVD(Mid$(XXX$, 445, 8))) ' importe iva del item
      TETABU1$(26, CAITAB1%) = Str$(CVD(Mid$(XXX$, 453, 8))) ' total bruto
      '
      SUBTOTA# = SUBTOTA# + XVALO(TETABU1$(23, CAITAB1%))
      SUBTOTB# = SUBTOTB# + XVALO(TETABU1$(26, CAITAB1%))
      '
      NUSERRR$ = TRIM$(Mid$(XXX$, 139, 64)) ' SI TIENE NUMERO DE SERIE
      If NUSERRR$ <> "" And UCase$(NUSERRR$) <> "S/N" Then 'IMPRIME NUMERO DE SERIE ABAJO DE LA DESCRIPCION
        CAITAB1% = CAITAB1% + 1
        TETABU1$(12, CAITAB1%) = "S/Nro.: " + NUSERRR$ ' numero de serie
      End If
      '
      If Control("FACTUREM", "INCLIEUSPE") = "SI" Then ' IMPRIME ESPECIFICACIONES DEL ARTICULO
        XX$ = ESPECIFIT$(CIPRE%)
        If TRIM$(XX$) <> "" Then
          CAITAB1% = CAITAB1% - 1 ' DESCUENTA 1 PARA PISAR LA DESCRIPCION
          Call PRINTEXT(XX$, 2, "DES-MAT")
        End If
      End If
      '
      Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
      '
      If COLTELIB% > 0 Then
          TELIBRE$ = REGLEIDO$(ARCHILIB$, U&)
          FORTELIB$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "DES-MAT/FORMATO")))
          ATELIB% = XVALO(Mid$(FORTELIB$, 2))
          Call FRATEXTO(TELIBRE$, ATELIB%)
          TETABU1$(2, CAITAB1%) = RETEX$(1)
          For KKU1% = 2 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
        Else
          FORTELIB$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "DES-MAT/FORMATO")))
          ATELIB% = XVALO(Mid$(FORTELIB$, 2))
          If ATELIB% > 100 Then
            ATELIB% = ATELIB% Mod 100
            If CIPRE% > 0 Then
                 TELIBRE$ = DESCRIT0$(CIPRE%)
                 If PIVADOC% = 5 Then
                   DEALTE$ = TRIM$(DESCRIT2$(CIPRE%))
                   If DEALTE$ <> "" Then
                      TELIBRE$ = DEALTE$
                   End If
                 End If
               Else
                 TELIBRE$ = Mid$(XXX$, 7, 64)
            End If
            Call FRATEXTO(TELIBRE$, ATELIB%)
            TETABU1$(2, CAITAB1%) = RETEX$(1)
            For KKU1% = 2 To CARETEX%
              CAITAB1% = CAITAB1% + 1
              TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
            Next KKU1%
          End If
      End If
      '
      ' AQUI BUSCA EL NUMERO DE SERIE
      If CIPRE% > 0 Then
        DOPRIC$ = TRIM$(Mid$(XXX$, 387, 16))
        NORIT% = CVI(Mid$(XXX$, 503, 2))
        '
        SQLX$ = "SELECT * from MOVISTO "
        SQLX$ = SQLX$ + "WHERE DoPriCor = '" & DOPRIC$ & "' "
        SQLX$ = SQLX$ + "AND Cod_Inte = " & CIPRE% & " "
        SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
        
        Set CONSITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With CONSITEMP
          If Not (.EOF And .BOF) Then
            SERNU$ = !ReferCor
            If Left$(SERNU$, 5) = "S.No." Then
              CAITAB1% = CAITAB1% + 1
              TETABU1$(2, CAITAB1%) = SERNU$
            End If
          End If
        End With
      End If
      '
      ' AQUI AGREGA DESPACHO DE IMPORTACION
      If AGREDESP% = 0 Then
        AGREDESP% = (-1)
        If Control$("FACTURA", "AGREDESP") = "SI" Then AGREDESP% = 1
      End If
      If AGREDESP% = 1 Then
        DESPAI$ = TRIM$(ATRIB1$(CIPRE%))
        If DESPAI$ <> "" Then
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = "Despacho: " + DESPAI$
        End If
      End If
      'ESTO SE LLEVO PARA ARRIBA POR QUE SE GENERA UN DIFERENCIA EN EL CAITAB Y NO REALIZA LA SUMA
'      SUBTOTA# = SUBTOTA# + XVALO(TETABU1$(23, CAITAB1%))
'      SUBTOTB# = SUBTOTB# + XVALO(TETABU1$(26, CAITAB1%))
      '
      If (U& Mod ITPORHOJA%) = 0 Then
        If U& < FIN& Then
          If CACOPDOC% > 0 Then
            DOCPARAM$(26) = TRIM$(CSTRING$(MKD$(SUBTOTA#), 4, 12, 2, 2))
            NOJA% = NOJA% + 1
            If REPRIFAC% = 0 Or NOJA% >= PROJA% Then
                 ' ESTO QUE SIGUE ES PARA IMPRIMIR TRANSPORTE DEL IMPORTE
                 FORSUBTO$ = UCase$(TRIM$(UCase$(ATRIFORM$(CODFOR$, "SUB-TOTA/FORMATO"))))
                 If Left$(FORSUBTO$, 4) = "SUMA" Then
                     CAITAB1% = CAITAB1% + 1
                     For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
                     TETABU1$(23, CAITAB1%) = "=========================="
                     TETABU1$(26, CAITAB1%) = "=========================="
                     CAITAB1% = CAITAB1% + 1
                     For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
                     TETABU1$(2, CAITAB1%) = "TRANSPORTE .........................................................."
                     TETABU1$(23, CAITAB1%) = FORMATNUM$(SUBTOTA#, "F13.2")
                     TETABU1$(26, CAITAB1%) = FORMATNUM$(SUBTOTB#, "F13.2")
                 End If
                 ' HASTA AQUI AGREGADO EL 22/12/2009 PARA IMPRIMIR TRANSPORTE
                 Call PRINTDOC("@" + FORIPRE$, "NOCANOBS")
              Else
                 For KKI% = 0 To CACOLTAB1%
                   For KKJ% = 0 To CAITAB1%
                     TETABU1$(KKI%, KKJ%) = ""
                   Next KKJ%
                 Next KKI%
                 CAITAB1% = 0
            End If
          End If
        End If
      End If
109 Next U&
    '
    'IMPRIME GASTOS DE EXPORTACION
    If Label13.Caption = "E" And XVALO(FINFACTUR07.Text2(15)) > 0 Then
        CAITAB1% = CAITAB1% + 1
        For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
        TETABU1$(2, CAITAB1%) = TRIM$(FINFACTUR07.Text10)
        TETABU1$(3, CAITAB1%) = 1
        TETABU1$(22, CAITAB1%) = TRIM$(FINFACTUR07.Text2(15))
        TETABU1$(26, CAITAB1%) = TRIM$(FINFACTUR07.Text2(15))
        SUBTOTB# = SUBTOTB# + XVALO(TETABU1$(26, CAITAB1%))
    End If
    '
    FORSUBTO$ = UCase$(TRIM$(UCase$(ATRIFORM$(CODFOR$, "SUB-TOTA/FORMATO"))))
    If Left$(FORSUBTO$, 3) = "SUM" Then
      CAITAB1% = CAITAB1% + 1
      For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
      TETABU1$(23, CAITAB1%) = "=========================="
      TETABU1$(26, CAITAB1%) = "=========================="
      CAITAB1% = CAITAB1% + 1
      For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
      TETABU1$(2, CAITAB1%) = "TOTAL .........................................................."
      TETABU1$(23, CAITAB1%) = FORMATNUM$(SUBTOTA#, "F13.2")
      TETABU1$(26, CAITAB1%) = FORMATNUM$(SUBTOTB#, "F13.2")
    End If
    '
    'SE MODIFICA VARIABLE POR QUE POVERTI% ES UNA VARIABLE PUBLICA
    If POVERTI1% < 1 Then
      CAITAB1% = CAITAB1% + 1
      For KU& = 1 To ULTREG&("!OBSERVOF")
        TTXX$ = REGLEIDO$("!OBSERVOF", KU&)
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = RTrim$(TTXX$)
      Next KU&
      Call SETULTREG("!OBSERVOF", 0)
    End If
        
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = FORMATNUM$(XVALO(Total(0).TEXT) + XVALO(IEMBALA) + XVALO(FINFACTUR07.Text2(15)), "F13.2")
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = Total(3).TEXT
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = Total(4).TEXT
    'CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = Total(1).TEXT
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = TRIM$(FORMATNUM(XVALO(Total(1).TEXT) - IMPODEGA#, "F12.2"))
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = FORMATNUM$(XVALO(Total(2).TEXT) + XVALO(FINFACTUR07.Text2(15)), "F13.2")
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = FORMATNUM$(XVALO(Total(6).TEXT) + XVALO(FINFACTUR07.Text2(15)), "F13.2")
    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = Total(5).TEXT
       VADOLA# = XVALO(Total(6).TEXT) * TICAMBX#
    DOCPARAM$(23) = TRIM$(CSTRING$(MKD$(VADOLA#), 4, 12, 2, 2))
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = Text15
    DOCPARAM$(26) = ""
    DOCPARAM$(27) = Text17                          'PARA IMPRIMIR PERC-IVA
      ITOGRAVA# = XVALO(Total(2)) - XVALO(TLABEL28)
    DOCPARAM$(45) = TRIM$(FORMATNUM$(ITOGRAVA#, "F13.2"))
      INOGRAVA# = XVALO(TLABEL28)
    DOCPARAM$(46) = TRIM$(FORMATNUM$(INOGRAVA#, "F13.2"))
    DOCPARAM$(47) = TRIM$(BASIVA(0))
    DOCPARAM$(48) = TRIM$(BASIVA(1))
      TOTATOTA# = XVALO(Total(6)) + XVALO(FINFACTUR07.Text2(12)) + XVALO(FINFACTUR07.Text2(13))
    DOCPARAM$(51) = FORMATNUM$(TOTATOTA#, "F13.2")
    '
    If Label13.Caption = "B" Then
      DOCPARAM$(5) = FORMATNUM$(COEFIVAINCLIEU# * (XVALO(Total(0).TEXT) + XVALO(IEMBALA)), "F13.2")
      DOCPARAM$(12) = FORMATNUM$(COEFIVAINCLIEU# * (XVALO(Total(1).TEXT) - IMPODEGA#), "F13.2")
      DOCPARAM$(13) = FORMATNUM$(COEFIVAINCLIEU# * XVALO(Total(2).TEXT), "F13.2")
      DOCPARAM$(42) = FORMATNUM$(COEFIVAINCLIEU# * XVALO(IEMBALA), "F13.2")          ' gasto de embalaje
      DOCPARAM$(45) = TRIM$(FORMATNUM$(COEFIVAINCLIEU# * ITOGRAVA#, "F13.2"))
      DOCPARAM$(46) = TRIM$(FORMATNUM$(COEFIVAINCLIEU# * INOGRAVA#, "F13.2"))
      DOCPARAM$(47) = TRIM$(FORMATNUM$(COEFIVAINCLIEU# * XVALO(BASIVA(0)), "F13.2"))
      DOCPARAM$(48) = TRIM$(FORMATNUM$(COEFIVAINCLIEU# * XVALO(BASIVA(1)), "F13.2"))
      DOCPARAM$(54) = TRIM$(FORMATNUM$(COEFIVAINCLIEU# * XVALO(DOCPARAM$(54)), "F13.2"))
    End If
    '
    CODFOR$ = FORIPRE$
    LOIMPLET% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET"))
    If LOIMPLET% < 16 Then LOIMPLET% = 32
    TTTT# = XVALO(Total(6).TEXT) + XVALO(FINFACTUR07.Text2(15))
    TOTFAC$ = NUMTEX$(TTTT#)
    If MONECO$ <> "" Then
      PPXX% = InStr(TOTFAC$, "PESOS")
      If PPXX% > 0 Then
        TOTFAC$ = Left$(TOTFAC$, PPXX% - 1) + UCase$(MONECO$) + Mid$(TOTFAC$, PPXX% + 5)
      End If
    End If
    Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = NRX1$
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = NRX2$
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = NRX3$
    '
    If CACOPDOC% > 0 Then
      Call PRINTDOC("@" + FORIPRE$, "NOCANOBS")
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    If CONTROFIS% < 2 Then      ' OBVIA LOS COMPROBANTES 'E' DE LA IMP. FISCAL
       CACOPDOC% = XVALO(FINFACTUR07.Text2(1).TEXT)
       If CACOPDOC% > 0 Then
          If TRIM$(Text10) <> "" Then
           '
           Call LIMPIATETAB
           '
           DOCUORIG$ = NUMEDOC$
           'ESTO SE AGREGA PORQUE CUANDO SE EMITE SOLO REMITO NO TIENE QUE IMPRIMIR
           If XVALO(FINFACTUR07.Text2(0).TEXT) = 0 Then DOCUORIG$ = ""
           Call PRIREMIFA
           '
           Call LIMPIATETAB
           Call IMPRIMENOTAENTREGA
          End If
       End If
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    If NOCONFIRM% > 0 Then
       Call COMUNI("Se han Emitido los Documentos Solicitados")
    End If
    For KKI% = 0 To CACOLTAB1%
      For KKJ% = 0 To CAITAB1%
        TETABU1$(KKI%, KKJ%) = ""
      Next KKJ%
    Next KKI%
    Call SETULTREG("!OBSERVOF", 0)
    '
End Sub

Sub PRICONFIS()
' DISTINGUIR:
' A) fACTURA / DEBITO / CREDITO
' B) A O B
   '

   If YAIMPRIMIO% > 0 Then Exit Sub
   '
   If TIFADOC% = 2 Then                ' ticket factura
        CACOPDOCF% = XVALO(FINFACTUR07.Text2(0).TEXT) ' LEE CANTIDAD COPIAS FACTURA
        CACOPDOCR% = XVALO(FINFACTUR07.Text2(1).TEXT) ' LEE CANTIDAD COPIAS REMITO
        If VARIDOC% = 1 Then           ' factura a
            If CACOPDOCF% > 0 Then Call PRITIFACA
            If CONTROFIS% = 2 And CACOPDOCR% > 0 Then Call PRITIREMI
          ElseIf VARIDOC% = 2 Then     ' factura b
            If CACOPDOCF% > 0 Then Call PRITIFACB
            If CONTROFIS% = 2 And CACOPDOCR% > 0 Then Call PRITIREMI
          ElseIf VARIDOC% = 3 Then     ' FACTURA EXPORTACION
            If CONTROFIS% = 2 And CACOPDOCR% > 0 Then
               Call COMUNI("Se Imprimirá Remito de Exportacion por Impresora Fiscal.")
               Call PRITIREMI
            End If
            If CACOPDOCF% > 0 Then
               Call MENSERR(24, "Cliente de Exportacion.\  \Imposible Emitir Factura Exportacion\por Impresora Fiscal.")
            End If
          Else
            Call MENSERR(24, "Imposible Emitir Factura Variante '" & TRIM$(Str$(VARIDOC%)) & "'\por Impresora Fiscal.")
            GoTo 90
        End If
        Exit Sub
     ElseIf TIFADOC% = 3 Then          ' nota de debito
        If VARIDOC% = 1 Then           ' N/D A
            Call PRINODEBA
          ElseIf VARIDOC% = 2 Then     ' N/D B
            Call PRINODEBB
          Else
            GoTo 90
        End If
        Exit Sub
     ElseIf TIFADOC% = 6 Then           ' nota de credito
        If VARIDOC% = 1 Then            ' N/C A
            Call PRINOCREA
          ElseIf VARIDOC% = 2 Then        ' N/C B
            Call PRINOCREB
          Else
            GoTo 90
        End If
        Exit Sub
     Else
    
90     Call MENSERR(24, "Comprobante no Admisible por Controlador Fiscal. Consulte")
       Call FINAL("")
   End If
   '
End Sub
Sub PRITIREMI()
    '
    If CONTROFIS% = 1 Then
        Call MENSERR(24, "Imposible Emitir Remito por Controlador Fiscal")
      ElseIf CONTROFIS% = 2 Then
        Call FORHAS320.PRITIREMI
    End If
    '
End Sub

Sub PRITIFACA()
    '
    If CONTROFIS% = 1 Then
      Call FORHASAR.PRITIFACA
    ElseIf CONTROFIS% = 2 Then
      Call FORHAS320.PRITIFACA
    End If
    '
End Sub


Sub PRITIFACB()
    '
    If CONTROFIS% = 1 Then
      Call FORHASAR.PRITIFACB
    ElseIf CONTROFIS% = 2 Then
      Call FORHAS320.PRITIFACB
    End If
    '
End Sub
Sub PRINODEBA()
   '
   If CONTROFIS% = 1 Then
     Call FORHASAR.PRINODEBA
   ElseIf CONTROFIS% = 2 Then
     Call FORHAS320.PRINODEBA
   End If
   '
End Sub
Sub PRINODEBB()
   '
   If CONTROFIS% = 1 Then
     Call FORHASAR.PRINODEBB
   ElseIf CONTROFIS% = 2 Then
     Call FORHAS320.PRINODEBB
   End If
   '
End Sub
Sub PRINOCREA()
    '
    If CONTROFIS% = 1 Then
      Call FORHASAR.PRINOCREA
    ElseIf CONTROFIS% = 2 Then
      Call FORHAS320.PRINOCREA
    End If
    '
End Sub

Sub PRINOCREB()
   '
   If CONTROFIS% = 1 Then
     Call FORHASAR.PRINOCREB
   ElseIf CONTROFIS% = 2 Then
     Call FORHAS320.PRINOCREB
   End If
   '
End Sub

Sub PRIREMIFA()
    '
TXTRAZA$ = TXTRAZA$ & "PUNTO 50.1 = " & Now & vbCrLf
    
    CACOPDOC% = XVALO(FINFACTUR07.Text2(1).TEXT)
    'If CACOPDOC% < 1 Then Exit Sub
    'CAMBIO
    SOLOREPLICAENEXPEDICION% = 0
    If CACOPDOC% < 1 Then
       If FINFACTUR07.Check8.Value = 1 Then 'CONSIDERA EL CASO QUE CANTI DE COPIAS SEA 0 PERO ESTE MARCADO REPLICA EXPEDICION
          SOLOREPLICAENEXPEDICION% = 1      'EN ESTE CASO IGUAL IMIPRIME IGUAL EN EXPEDICION.
       Else
         Exit Sub
       End If
    End If
    'FIN CAMBIO
TXTRAZA$ = TXTRAZA$ & "PUNTO 50.2 = " & Now & vbCrLf
    
    FORIPRE$ = TRIM$(Control$("", "FOREMIFA"))
    If TRIM$(FORIPRE$) <> "" Then
       '
TXTRAZA$ = TXTRAZA$ & "PUNTO 50.3 = " & Now & vbCrLf
       
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    
    IMPRIMELOTE% = 0
    If XVALO(Control$("FACTURA", "IMPRLOTE")) > 0 Then IMPRIMELOTE% = 1
    
       Call LIMPIATETAB
       '
       '\\\OT-06-0114-WAR-31/03/06///
       Call SETULTREG("!OBSERVOF", 0)
       If FINFACTUR07.Text5 <> "" Then
         JR& = 0
         CODFOR$ = FORIPRE$
         FORMOBSE$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
         ANCHOB% = XVALO(Mid$(FORMOBSE$, 2)): If ANCHOB% < 32 Then ANCHOB% = 32
         Call FRATEXTO(FINFACTUR07.Text5, ANCHOB)
         For KKII% = 1 To CARETEX%
           JR& = JR& + 1
           Call GRAREG("!OBSERVOF", RETEX$(KKII%), JR&)
         Next KKII%
         Call SETULTREG("!OBSERVOF", JR&)
         Call CIERRARCH("!OBSERVOF")
       End If
       'SE MODIFICA VARIABLE POR QUE POVERTI% ES UNA VARIABLE PUBLICA
       POVERTI1% = XVALO(TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/POSINVER"))))
       '\\\OT-06-0114-WAR-31/03/06///
       '
TXTRAZA$ = TXTRAZA$ & "PUNTO 50.4 = " & Now & vbCrLf
       
       PREFIREM$ = Control$("", "PREFIREM")
       If PREFIREM$ = "" Then PREFIREM$ = "0001"
       'ESTO SE AGREGA PORQUE LOS REMITOS DE LAS
       'FACTURAS ELECTRONICAS TIENEN EL PUNTO DE VENTA DE LA FACTURA
       ' If FACELEC% = 1 And VERIFSALIDA("NOMESS") < 3 Then PREFIREM$ = Label12
       If FACELEC% > 0 And FACELEC% < 3 Then PREFIREM$ = Label12
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "TANIT") > 0 Or InStr(EPAC$, "GABAL") > 0 Or InStr(EPAC$, "FERVI") > 0 Or InStr(EPAC$, "AHP") > 0 Or InStr(EPAC$, "J Y M") > 0 Or InStr(EPAC$, "CEYEQ") > 0 Or InStr(EPAC$, "EYM") > 0 Or InStr(EPAC$, "DAVTECNICA") > 0 Then PREFIREM$ = "0001"
       If InStr(EPAC$, "ASEMBLI") > 0 Then PREFIREM$ = Control$("", "PREFIREM")
       '
       NroRemi$ = TRIM$(Text10.TEXT)
       While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
       NroRemi$ = PREFIREM$ + "-" + NroRemi$
       '
       If InStr(EPAC$, "EYM") > 0 Then
          NroRemi$ = CMO$ + "-X-" + NroRemi$
       End If
       
       '
       NCLIE = XVALO(Text1.TEXT)
       FECHDOC$ = Text9.TEXT
       NUMEDOC$ = NroRemi$
       TIPODOC$ = "Remito"
       '
       If TRIM$(Text14.TEXT) <> "" Then
         NUOCOMPRA$ = TRIM$(Text14.TEXT)
       End If
TXTRAZA$ = TXTRAZA$ & "PUNTO 50.5 = " & Now & vbCrLf
       '
         NTRA% = XVALO(Text19)
       CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = LENTREGA(2).TEXT ' TRIM$(Mid$(TRXP$, 1, 32))
       CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(FINFACTUR07.Text2(2).TEXT)
       CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = CUITTRANS$(NTRA%) ' TRIM$(Mid$(TRXP$, 65, 16))
       CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(FINFACTUR07.Text2(3).TEXT)
       CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(TRXP$, 97, 32))
       CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(TRXP$, 129, 24))
       CODPARAM$(7) = "DEPOSAL ": 'DOCPARAM$(7) = TRIM$(DEPOSAL$)
       CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NroRemi$)
       CODPARAM$(9) = "NRO-OCOM":  DOCPARAM$(9) = NUOCOMPRA$
       CODPARAM$(10) = "PATE-TRA":  DOCPARAM$(10) = TRIM$(Mid$(TRXP$, 81, 16))
       CODPARAM$(11) = "DOMI-TRA":  DOCPARAM$(11) = LENTREGA(0).TEXT
       CODPARAM$(12) = "LOCA-TRA":  DOCPARAM$(12) = TRIM$(Mid$(TRXP$, 159, 32))
       CODPARAM$(13) = "COND-PAG": DOCPARAM$(13) = Eco21(1).TEXT
       CODPARAM$(14) = "NRO-AENT":  DOCPARAM$(14) = NROPEDIDO$
       CODPARAM$(15) = "DESTINO": DOCPARAM$(15) = LENTREGA(1).TEXT
       CODPARAM$(16) = "TRANSPOR": DOCPARAM$(16) = LENTREGA(2).TEXT
       CODPARAM$(17) = "MEDIDA-1": DOCPARAM$(17) = TRIM$(IEMBALA)             ' gasto de embalaje
       CODPARAM$(18) = "ZONA-VEN": DOCPARAM$(18) = TRIM$(Str$(GRUCOCLI%(NCLIE)))
       CODPARAM$(19) = "TOTA-ITS": DOCPARAM$(19) = Label25
       '
       CODPARAM$(20) = "REF-ASIE": DOCPARAM$(20) = DOMITRANS$(NTRA%)
       CODPARAM$(21) = "DEN-CUEN": DOCPARAM$(21) = LOCATRANS$(NTRA%)
       CODPARAM$(22) = "MEDIDA-2": DOCPARAM$(22) = TELETRANS$(NTRA%)
       'AGREGUE 02-10MH
       CODPARAM$(23) = "NRO-PASE": DOCPARAM$(23) = TRIM$(Label27.Caption)
       '\\\OT-07-0007-RG-08/01/07///
       HORA$ = HORAENT$(NCLIE)
       CODPARAM$(24) = "Z-PLANO": DOCPARAM$(24) = TRIM$(HORA$)
          NUVENDEX$ = AJUSTI$(FORMATNUM$(XVALO(TEXT21(2)), "I3"), 4)
       CODPARAM$(25) = "VENDEDOR": DOCPARAM$(25) = NUVENDEX$ + Eco21(2).TEXT
       '\\\OT-07-0007-RG-FIN///
       '
       'ARMADO DE CADENA PARA CODIGO DE BARRA DE CLIENTE+REMITO
       NUCL$ = Right("000000" & TRIM$(Text1), 6)
       NUREMI$ = REPLACAR(NroRemi$, "-", "")
       Cadena$ = NUCL$ + NUREMI$
       BARRA$ = CODIBARRA$(Cadena$, 2)
       '
       CODPARAM$(26) = "REF-MAT3": DOCPARAM$(26) = BARRA$
       DEP% = XVALO(FINFACTUR07.Text7) 'DEPOSITO DE SALIDA
       CODPARAM$(27) = "DEPOSITO": DOCPARAM$(27) = Str$(DEP%) + " " + TRIM$(Mid$(ECOARCH$("TADEPOSI", Chr(DEP%)), 1, 16))
       CODPARAM$(28) = "TOTA-CAN"  'CANTIDAD DE BULTOS ASOCIADA AL  LOTE INDIVISIBLE
       CODPARAM$(29) = "MEDIDA-3": DOCPARAM$(29) = CTACTRANS$(NTRA%)
       '
       CAPARDOC% = 29
       CODTABU1$(1) = "COD-MAT"
       CODTABU1$(2) = "DES-MAT"
       CODTABU1$(3) = "CANTIDAD"
       CODTABU1$(4) = "ORD-ITEM"
       CODTABU1$(5) = "CAN-PLCO" 'CANTIDAD DE BULTOS
       CODTABU1$(6) = "UBI-MPRI"
       CODTABU1$(7) = "REF-MAT6"
       CODTABU1$(8) = "NRO-LOTE"
       CODTABU1$(9) = "N-OCOMPR"
       
       CACOLTAB1% = 9
       '
       DESTIDOCNUE = NCLIE
       CAITAB1% = 0
       '
       If FINFACTUR07.Text3 <> "" Then
         Call FRATEXTO(FINFACTUR07.Text3, 40)
         For KKII% = 1 To CARETEX%
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = ""
           TETABU1$(2, CAITAB1%) = RETEX$(KKII%)
           TETABU1$(3, CAITAB1%) = ""
           TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
         Next KKII%
         FINFACTUR07.Text3 = ""
         FINFACTUR07.Frame5.Visible = False
         GoTo 80
       End If
       '
       ARCHILIB$ = "!1CARDEFA"
      
      '\\\OT-06-0339-MH-24/08/06///
      REPLAORI% = 0
      If Control$("REMSTAN", "REPLAORI") = "SI" Then
        REPLAORI% = 1
      End If
      '\\\OT-06-0339-MH-FIN///
 TXTRAZA$ = TXTRAZA$ & "PUNTO 50.6 = " & Now & vbCrLf
      
       For U& = 1 To ULTREG&("!CARDEFAC")
         XXX$ = REGLEIDO$("!CARDEFAC", U&)
         CANS$ = Str$(CVS(Left$(XXX$, 4))) ' cantidad
         NCLIE = XVALO(Text1.TEXT) ' AGREGADO MH
         CIPRE% = CVI(Mid$(XXX$, 5, 2)) ' AGREGADO
         NUOCO$ = SAFETEXT$(TRIM$(Mid$(XXX$, 371, 16)))

         'If XVALO(CANS$) >= 0.05 Then
           If TRIM$(Mid$(XXX$, 387, 16)) = "" Then
             CAITAB1% = CAITAB1% + 1
             DDTXT$ = Mid$(XXX$, 7, 64)                  ' descripcion
             ' SE AGREGA CONDICION PARA QUE ACTUE UNICAMENTE EN DOSIVAC
             ' Y MANTENER LA FUNCIONALIDAD PARA EL RESTO DE LOS USUARIOS
             If PIVADOC% = 5 And InStr(EMPREAC$, "DOSIVA") > 0 Then
               DEALTE$ = TRIM$(DESCRIT2$(CIPRE%))
               If DEALTE$ <> "" Then
                  DDTXT$ = DEALTE$
               End If
             End If
             '
             CDXXS$ = CODEXT$(CVI(Mid$(XXX$, 5, 2)))     ' codigo
             '\\\OT-06-0339-MH-24/08/06///
             If REPLAORI% = 1 Then
               ' se pasa el codigo interno, mas el nº cliente(NCLIE) y trae el codigo externo solicitado
               CDXXS$ = CODORIG$(CIPRE%, NCLIE) ' + SUBCO$
             End If
             '
             '\\\OT-06-0339-MH-FIN///
             CINTE% = CODINT%(CDXXS$)
             Cantidad# = XVALO(CANS$)
             '
             If DESCOMPO%(CINTE%) > 0 Then
                'Imprimir componentes en factura
                If InStr(EPAC$, "AHP") > 0 Then GoTo 50 ' PARA QUE NO IMPRIMA LOS COMPONENTES EN AHP EN REMITO.
                Call AREXPLO(CINTE%, 1, 1)
                NUORIT% = 0
                For KKJ% = 1 To CAIT%
                  NUORIT% = NUORIT% + 1
                  CICO% = CIJ%(KKJ%)
                  CACICO = USOI(KKJ%)
                  
                  TETABU1$(1, CAITAB1%) = CODEXT$(CICO%)
                  TETABU1$(2, CAITAB1%) = "..." + DESCRIT0$(CICO%)
                  CANTI# = CACICO * Cantidad#
                  TETABU1$(3, CAITAB1%) = TRIM$(FORMATNUM$(CANTI#, "F18.6"))
                  TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
                  TETABU1$(6, CAITAB1%) = UBISTO$(CICO%)
                  If NUORIT% = 1 Then
                    'SOLO IMPRIME EN EL PRIMERO LA CANTIDAD DE BULTOS DEL PACK
                    CABU = CANBULTOS(CIPRE%, CANTI#, "VENTA")
                    TETABU1$(5, CAITAB1%) = TRIM$(FORMATNUM$(CABU, "F18.6"))
                    TOTABULTO = TOTABULTO + CABU
                  End If
                  If KKJ% < CAIT% Then CAITAB1% = CAITAB1% + 1
                Next KKJ%
              Else 'De lo Contrario Imprimir el Item
50              TETABU1$(1, CAITAB1%) = CDXXS$
                TETABU1$(2, CAITAB1%) = DDTXT$
                TETABU1$(3, CAITAB1%) = CANS$
                TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
                CABU = CANBULTOS(CIPRE%, Cantidad#, "VENTA")
                TETABU1$(5, CAITAB1%) = TRIM$(FORMATNUM$(CABU, "F18.6"))
                TOTABULTO = TOTABULTO + CABU
                TETABU1$(6, CAITAB1%) = UBISTO$(CINTE%)
                    CORIG$ = TRIM$(CODORIG$(CIPRE%, NCLIE))
                    If CORIG$ = TRIM$(CDXXS$) Then CORIG$ = ""
                TETABU1$(7, CAITAB1%) = CORIG$
                TETABU1$(8, CAITAB1%) = "" 'LO LIMPIO Y LUEGO VALIDO E IMPRIMO EL NSERIE
                NUSERRR$ = TRIM$(Mid$(XXX$, 139, 64)) 'NUMERO DE SERIE
                
                If NUSERRR$ <> "" And UCase$(NUSERRR$) <> "S/N" Then
                   TETABU1$(8, CAITAB1%) = NUSERRR$
                Else
                    If IMPRIMELOTE% > 0 Then
                       TETABU1$(8, CAITAB1%) = LOTEAIMPRIMIR$(CINTE%, U&)
                    End If
                End If
                
                TETABU1$(9, CAITAB1%) = NUOCO$ ' NPEDIDO POR EYM
                If Control("FACTUREM", "INCLUSPE") = "SI" Then ' IMPRIME ESPECIFICACIONES DEL ARTICULO
                  XX$ = ESPECIFIT$(CIPRE%)
                  If TRIM$(XX$) <> "" Then
                    CAITAB1% = CAITAB1% - 1 ' DESCUENTA 1 PARA PISAR LA DESCRIPCION
                    Call PRINTEXT(XX$, 2, "DES-MAT")
                  End If
                End If

             End If
             '

             If NUSERRR$ <> "" And UCase$(NUSERRR$) <> "S/N" And InStr(EPAC$, "EYM") <= 0 Then 'IMPRIME NUMERO DE SERIE ABAJO DE LA DESCRIPCION
               CAITAB1% = CAITAB1% + 1 ' SI NO ES EYM LO AGREGA ABAJO DE LA DESCRIPCION
               TETABU1$(2, CAITAB1%) = "S/Nro.: " + NUSERRR$ ' numero de serie
             End If
             '
             'CANTIDAD DE BULTOS SEGUN LOTE INDIVISIBLE
             DOCPARAM$(28) = TRIM$(FORMATNUM$(TOTABULTO, "F18.6"))
             
             'IMPRIME DETALLE DE LOTES REMITIDOS
'             Call DETLOTREM(CIPRE%, U&)
             '
             CODFOR$ = FORIPRE$
             Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
             '
             If COLTELIB% > 0 Then
               TELIBRE$ = REGLEIDO$(ARCHILIB$, U&)
               FORTELIB$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "DES-MAT/FORMATO")))
               ATELIB% = XVALO(Mid$(FORTELIB$, 2))
               Call FRATEXTO(TELIBRE$, ATELIB%)
               TETABU1$(2, CAITAB1%) = RETEX$(1)
               For KKU1% = 2 To CARETEX%
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
               Next KKU1%
             End If
             '\\\OT-06-0273-WAR-03/07/06///
             ' AQUI AGREGA DESPACHO DE IMPORTACION
             If AGREDESP% = 0 Then
               AGREDESP% = (-1)
               If Control$("FACTURA", "AGREDESP") = "SI" Then AGREDESP% = 1
             End If
             If AGREDESP% = 1 Then
               CIIY% = CODINT%(CDXXS$)
               DESPAI$ = TRIM$(ATRIB1$(CIIY%))
               If DESPAI$ <> "" Then
                 'CAITAB1% = CAITAB1% + 1
                 LONGDESCRI% = XVALO(Mid(TRIM$(ATRIFORM$(CODFOR$, "DES-MAT/FORMATO")), 2))
                 LONGDESPACHO% = Len(" Despacho: " + DESPAI$)
                 TETABU1$(2, CAITAB1%) = Mid$(TETABU1$(2, CAITAB1%), 1, LONGDESCRI% - LONGDESPACHO%) & " Despacho: " + DESPAI$
               End If
             End If
             '\\\OT-06-0273-WAR-FIN///
             '
           End If
         'End If
       Next U&
       '
TXTRAZA$ = TXTRAZA$ & "PUNTO 50.7 = " & Now & vbCrLf
       
       'SE MODIFICA VARIABLE POR QUE POVERTI% ES UNA VARIABLE PUBLICA
       If POVERTI1% < 1 Then
         CAITAB1% = CAITAB1% + 1
         For KU& = 1 To ULTREG&("!OBSERVOF")
           TTXX$ = REGLEIDO$("!OBSERVOF", KU&)
           CAITAB1% = CAITAB1% + 1
           TETABU1$(2, CAITAB1%) = RTrim$(TTXX$)
         Next KU&
         Call SETULTREG("!OBSERVOF", 0)
       End If
       '
       'AGREGADO PARA EL CASO QUE SEA 0 COPIAS DE REMITO PERO SE ENCUENTRE MARCADA OPCION DE REPLICA EN EXPEDICION
       If SOLOREPLICAENEXPEDICION% > 0 Then GoTo 90
      
80     CACOPDOC% = XVALO(FINFACTUR07.Text2(1).TEXT)
       PRINTERPORT$ = PORTUREM$
 TXTRAZA$ = TXTRAZA$ & "PUNTO 50.8 = " & Now & vbCrLf
       Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
  TXTRAZA$ = TXTRAZA$ & "PUNTO 50.9 = " & Now & vbCrLf
      
       '
       'SE COMENTA PORQUE AHORA EL CHECK DE REPLICA EN EXPEDICION
       'ESTA EN LA PANTALLA DE FINFACTURA07
       '
'       IMPREOR$ = CONTROL("FACTURA", "REPLIREM")
'       PORTURE$ = TRIM$(CONTROL$("ORDESPA", "PUERTOIM"))
       '
90    IMPREOR$ = "NO"
      '
      If FINFACTUR07.Check8.Value = 1 Then
         IMPREOR$ = "SI"
         PORTURE$ = TRIM$(Control$("ORDESPA", "PUERTOIM"))
         CACOPDOC% = 1
      End If
      '
      If IMPREOR$ = "SI" And PORTURE$ <> "" Then
         PRINTERPORT$ = PORTURE$
         Call SETPRINPORT(PORTURE$)   ' IMPRIME UNA COPIA DEL REMITO SI ESTA MARCADA LA
 TXTRAZA$ = TXTRAZA$ & "PUNTO 50.9 = " & Now & vbCrLf
                                       ' OPCION DEL SETUP 'REPLICA EN  DESPACHO'
         Call PRINTDOC("FOREMIFA")    ' Y EL MISMO SE IMPRIME EN LA IMPRESORA CONFIGURADA
  TXTRAZA$ = TXTRAZA$ & "PUNTO 50.10 = " & Now & vbCrLf
        
         Call SETPRINPORT("RESTORE")  ' COMO IMPRESION DE ORDEN DE DESPACHO (SETUP DE PEDIDOS)
  TXTRAZA$ = TXTRAZA$ & "PUNTO 50.11 = " & Now & vbCrLf
      End If
    End If
    '
    Call LIMPIATETAB
End Sub

Function LOTEAIMPRIMIR$(CI%, REG&)
    '
    LOTEAIMPRIMIR$ = ""
    If ESTRAZABLE%(CI%) = 1 Then
       TOTALOTE$ = ""
       For UY& = 0 To ASIGLOT07.List1.ListCount - 1
           YY$ = ASIGLOT07.List1.List(UY&)
           If XVALO(Mid$(YY$, 80, 5)) = REG& Then
               LOTEAIMPRIMIR$ = TRIM$(Mid$(YY$, 85))
           End If
       Next UY&
       If Right$(LOTEAIMPRIMIR$, 1) = ";" Then LOTEAIMPRIMIR$ = Left$(LOTEAIMPRIMIR$, Len(LOTEAIMPRIMIR$) - 1)
    End If
    
End Function
Private Sub Timer2_Timer()
'   Static PASADAS%
'   PASADAS% = PASADAS% + 1
'   If PASADAS% > 29 Then
'      Timer2.Enabled = False
'      TESTCAE09.YAGRABOCAE = 0
'      RECIBO08.Show 1
'   End If
'   '
'   If TESTCAE09.YAGRABOCAE% > 0 Then
'      Timer2.Enabled = False
'      TESTCAE09.YAGRABOCAE = 0
'      RECIBO08.Show 1
'   End If
End Sub

Private Sub Timer3_Timer()
        'AHP 11 ESTA EN DESHUSO
        Timer3.Enabled = False
        Exit Sub
        
        'ATAJA ERROR CUANDO ES SOLO FACTURA SOBRE REMITO YA REALIZADO
        If ASIGLOAHP.List2.ListCount < 1 Then
           Timer3.Enabled = False
           Exit Sub
        End If
        
        Set colProcessList = GetObject("Winmgmts:").ExecQuery("Select * from Win32_Process")
        Static VUELTA%
        Dim Encontrado As Boolean
        VUELTA% = VUELTA% + 1
        For Each objProcess In colProcessList
        Debug.Print UCase$(objProcess.Name)
        If UCase$(objProcess.Name) = "WSFE_CLIENTE.EXE" Or UCase$(objProcess.Name) = "WSBFE_CLIENTE.EXE" Or UCase$(objProcess.Name) = "WSFEX_CLIENTE.EXE" Then
           Encontrado = True
           Exit For
        End If
        Next
        '
        If Encontrado = True Then
          If VUELTA% > 10 Then
            Timer3.Enabled = False
            Exit Sub
          End If
        Else
        'LLAMAR A FRM Y EJECUTARLO
          VUELTA% = 0
          Timer3.Enabled = False
'          If ASIGLOAHP.List2.ListCount >= 1 Then
'             ASIGLOAHP.EJECUTAR = 1
'             ASIGLOAHP.Show 1
'          End If
        End If
        
End Sub

Private Sub TLABEL28_Change()
'   NOGRA# = XVALO(TLABEL28)
'   IGRAVA0# = XVALO(Total(2)) - NOGRA#
'   If TRIM$(BASIVA(0)) <> "" Then
'       BASIVA(0) = TRIM$(FORMATNUM$(IGRAVA0#, "F13.2"))
'       BASIVA(1) = ""
'     Else
'       BASIVA(0) = ""
'       BASIVA(1) = TRIM$(FORMATNUM$(IGRAVA0#, "F13.2"))
'   End If
'   Total(3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(0)) * XVALO(PORIVA(0).Caption) * APLIVA% / 100), 3, 12, 2, 2)
'   Total(4).TEXT = CSTRING$(MKD$(XVALO(BASIVA(1)) * XVALO(PORIVA(1).Caption) * APLIVA% / 100), 3, 12, 2, 2)
End Sub

Private Sub TLABEL28_GotFocus()
   '
   If VALICLIE% < 1 Then GoTo 99
   If TIFADOC% < 3 Then GoTo 98
   '
98 On Error Resume Next
   PORDESCU.SetFocus
   On Error GoTo 0
   Exit Sub
   '
99 On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0

End Sub

Private Sub TLABEL28_LostFocus()
   If XVALO(TLABEL28) > XVALO(Total(2)) Then
      Call MENSERR(24, "Importe no Gravado Supera Neto Comprobante")
      TLABEL28 = ""
      Exit Sub
   End If
   TLABEL28 = TRIM$(FORMATNUM$(TLABEL28, "F13.2"))
End Sub

Private Sub Total_Change(Index As Integer)
   If VALICLIE% < 1 Then Exit Sub
   If TIFADOC% < 3 Then GoTo 10
   '
   If Index >= 3 Then
     If Index <= 4 Then
       Call CALTOFA2
     End If
   End If
   '
10 If Index = 6 Then
     If XVALO(Total(6)) > 0.005 Then
       If VERLICRE% = 1 Then
         If TIFADOC% < 4 Then ' EXCLUYE CREDITOS
           NCLIE% = XVALO(Text1)
           LMCR# = XVALO(Text7)
           SDEU# = SALDCLI#(NCLIE%)
           CHPA# = CHEPEACRE#(NCLIE%)
           TPED# = XVALO(Total(6).TEXT) * TICAMBX#
           If SDEU# + CHPA# + TPED# > LMCR# Then
             TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR# - SDEU# - CHPA#), 3, 12, 2, 2))
             If TLIMI$ = "" Or XVALO(TLIMI$) < 0 Then TLIMI$ = "0.00"
             Call COMUNI("Imposible Emitir Este Documento.\  \Supera el Límite de Crédito Disponible\para esta Cuenta de $ " + TLIMI$ + ".\  \Informe al Supervisor de Ventas.")
             Exit Sub
           End If
         End If
       End If
     End If
   End If
   '
20 If Index = 2 Or (VARIDOC% = 2 And Index < 6) Then
     Call CALPERIB
   End If
   '
End Sub

'
Private Sub Total_GotFocus(Index As Integer)
   '
   'PARA AHP EDITAR IVA
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   PVTANORMALPRINT$ = Control("FACTURA", "PUVENPRI")
   PERMITEEDITAR% = 0
   If XVALO(Label12) = XVALO(PVTANORMALPRINT$) And InStr(EPAC$, "AHP") Then
      PERMITEEDITAR% = 1
   End If

   If VALICLIE% < 1 Then GoTo 99
   If Index < 3 Or Index > 4 Then GoTo 98
   If TIFADOC% < 3 Then
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "FRIOLA") < 1 And InStr(EPAC$, "EYM") < 1 And InStr(EPAC$, "DAI ICHI") < 1 And PERMITEEDITAR% < 1 Then GoTo 98
   End If
   Call CALTOFA2
   Exit Sub
   '
98 On Error Resume Next
   PORDESCU.SetFocus
   On Error GoTo 0
   Exit Sub
   '
99 On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub
'
Sub GRAMOCABE()
   '
   MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESVENTA", "MODOMONE"))), 8)
   If MODOMONE$ <> "BIMONETA" Then
      If MODOMONE$ <> "MULTIMON" Then
         Exit Sub
      End If
   End If
   '
   FEMI% = FECHANUM(Text9)
   NCLIE% = XVALO(Text1)
   Tipo% = TIFADOC%
   VA% = VARIDOC%
   NRO& = XVALO(Text11.TEXT)
   NUDOCU$ = CoComLar$
   '
   SUCUVEN% = 0
   On Error Resume Next
   SUCUVEN% = XVALO(Label12)
   On Error GoTo 0
   '
   TOTAPESOS# = TOTAFAC#(1, 8)
   TOTADOLAR# = TOTAPESOS# / TICAMBX#
   '
   RAPLI$ = REGBLAN$("CARBIMON")
   ' documento cabeza
   Call REPLA(RAPLI$, MKI$(FEMI%), 1, 2)
   Call REPLA(RAPLI$, MKI$(NCLIE%), 3, 2)
   Call REPLA(RAPLI$, Chr$(Tipo%), 5, 1)
   Call REPLA(RAPLI$, MKI$(VA%), 6, 1)
   Call REPLA(RAPLI$, MKS$(NRO), 7, 4)
   Call REPLA(RAPLI$, MKI$(SUCUVEN%), 11, 2)
   Call REPLA(RAPLI$, NUDOCU$, 13, 18)
   Call REPLA(RAPLI$, MKI$(MONEMI%), 31, 2)
   '
   ' documento aplicado
   Call REPLA(RAPLI$, MKI$(FEMI%), 33, 2)
   Call REPLA(RAPLI$, MKI$(NCLIE%), 35, 2)
   Call REPLA(RAPLI$, Chr$(Tipo%), 37, 1)
   Call REPLA(RAPLI$, MKI$(VA%), 38, 1)
   Call REPLA(RAPLI$, MKS$(NRO), 39, 4)
   Call REPLA(RAPLI$, MKI$(SUCUVEN%), 43, 2)
   Call REPLA(RAPLI$, NUDOCU$, 45, 18)
   Call REPLA(RAPLI$, MKI$(MONEMI%), 63, 2)
   '
   'importe en pesos
   Call REPLA(RAPLI$, MKD$(TOTAPESOS#), 65, 8)
   Call REPLA(RAPLI$, MKD$(TICAMBX#), 89, 8)
   '
   'importe en dolares
   Call REPLA(RAPLI$, MKD$(TOTADOLAR#), 97, 8)
   '
   Call REGAPP("CARBIMON", RAPLI$)
   Call CIERRARCH("CARBIMON")
   '
End Sub
'
Sub APLIDICAM()
   '
   MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESVENTA", "MODOMONE"))), 8)
   If MODOMONE$ <> "BIMONETA" Then
      If MODOMONE$ <> "MULTIMON" Then
         Exit Sub
      End If
   End If
   '
   Screen.MousePointer = 11
   NCI% = XVALO(Text1)
   TIDOCAPLI% = 2
   NUDOCAPLI& = XVALO(Mid$(DOCUDIF, 9, 8))
   CANCE# = XVALO(Total(6)) * TICAMONE(MONEMI%)
   TICAMBX# = XVALO(TICAMBIS): If TICAMBX# < 0.05 Then TICAMBX# = 1
   '
   If TIFADOC% = 6 Then
     Call APERBASDAT("CABAN")
     DOCUCA$ = DOCUDIF
     ICAN# = CANCE#
     '
     NumeClie = NCI%
     DocOrCor$ = ""
     TotOrig# = 0
     IvaOrig# = 0
     '
     SQLX$ = " SELECT * FROM SaDeuPen "
     SQLX$ = SQLX$ + " WHERE NuClien = " & NCI% & " "
     SQLX$ = SQLX$ + " AND CODICOM_LAR = '" & DOCUCA$ & "' "
     SQLX$ = SQLX$ + " AND SALDdebe > 0"
     Dim SADDDEUPEN As ADODB.Recordset
     Set SADDDEUPEN = New ADODB.Recordset
25   On Error Resume Next
     SADDDEUPEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
     End If
     On Error GoTo 0
     '
     With SADDDEUPEN
       '.FindFirst "NuClien = " & NCI% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDdebe > 0"
       'If .NoMatch = False Then
       On Error Resume Next
       If Not (.EOF And .BOF) Then
         Do While (Not .EOF)
850        PARCAN# = ICAN#: If !SaldDebe < ICAN# Then PARCAN# = !SaldDebe
           '.Edit
           !SaldDebe = !SaldDebe - PARCAN#
           .Update
           '
           DocOrCor$ = !CodiCom_Cor
           DocOrLar$ = !codicom_lar
           TotOrig# = !IBruComp
           IvaOrig# = 0
           On Error Resume Next
           IvaOrig# = !IIvaComp
           On Error GoTo 0
           '
           ICAN# = ICAN# - PARCAN#
           If ICAN# < 0.005 Then
              Exit Do
  '           .FindNext "CODICOM_LAR = '" & DOCUCA$ & "' AND SALDdebe > 0"
  '           GoTo 850
           End If
           .MoveNext
         Loop
       End If
     End With
     SADDDEUPEN.Close
     Set SADDDEUPEN = Nothing
     '
     ImpApli# = CANCE#
     RefApli$ = "Diferencia de Cambio " + DOCUDIF
     Call GRAPLICOB
     '
     MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESVENTA", "MODOMONE"))), 8)
     If MODOMONE$ = "BIMONETA" Or MODOMONE$ = "MULTIMON" Then
       NCI% = NumeClie
       RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCI%))
       If Len(RFF$) > 0 Then
         If CANCE# >= 0.005 Then
           ABUSQUE$ = DOCUDIF
           For JKLX& = 1 To Len(RFF$) Step 4
             RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
             YCB$ = REGLEIDO$("CARBIMON", RECABI&)
             If TRIM$(Mid$(YCB$, 13, 18)) = ABUSQUE$ Or TRIM$(Mid$(YCB$, 45, 18)) = ABUSQUE$ Then
               DOCUCABE$ = TRIM$(Mid$(YCB$, 13, 18))
               For JKLY& = 1 To Len(RFF$) Step 4
                 RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
                 YCB$ = REGLEIDO$("CARBIMON", RECABI&)
                 If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
                     MONEMORI% = CVI(Mid$(YCB$, 31, 2))
                     TICAMBX# = TICAMONE(MONEMORI%)
                     FEDOAPLI% = FECHANUM(Text9)
                     NROO& = XVALO(Mid$(CoComLar$, 11))
                   Call REPLA(YCB$, MKI$(FEDOAPLI%), 33, 2)
                   Call REPLA(YCB$, MKI$(NCI%), 35, 2)
                   Call REPLA(YCB$, Chr$(TIDOAP%), 37, 1)
                   Call REPLA(YCB$, Chr$(VADOAP%), 38, 1)
                   Call REPLA(YCB$, MKS$(NROO&), 39, 4)
                   Call REPLA(YCB$, MKI$(SUC%), 43, 2)
                   Call REPLA(YCB$, CoComLar$, 45, 18)
                   Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
                   Call REPLA(YCB$, MKD$(0), 65, 8)
                   Call REPLA(YCB$, MKD$(CANCE#), 73, 8)
                   Call REPLA(YCB$, MKD$(0), 81, 8)
                   Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
                   Call REPLA(YCB$, MKD$(0), 97, 8)
                   Call REPLA(YCB$, MKD$(0), 105, 8)
                   Call REPLA(YCB$, MKD$(0), 113, 8)
                   '
                   Call REGAPP("CARBIMON", YCB$)
                   GoTo 835
                 End If
               Next JKLY&
             End If
           Next JKLX&
         End If
       End If
835    Call CIERRARCH("CARBIMON")
     End If
     '
   ElseIf TIFADOC% = 3 Then ' nota de debito
     '
     RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCI%))
     ABUSQUE$ = TRIM$(DOCUDIF)
     If Len(RFF$) > 0 Then
       For JKLX& = 1 To Len(RFF$) Step 4
         RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
         YCB$ = REGLEIDO$("CARBIMON", RECABI&)
         If Mid$(YCB$, 13, 18) = ABUSQUE$ Then
             FEDOAPLI% = FECHANUM(Text9)
           Call REPLA(YCB$, MKI$(FEDOAPLI%), 33, 2)
           Call REPLA(YCB$, MKI$(NCI%), 35, 2)
           Call REPLA(YCB$, Chr$(TIFADOC%), 37, 1)
           Call REPLA(YCB$, Chr$(VARIDOC%), 38, 1)
           Call REPLA(YCB$, MKS$(XVALO(Text11)), 39, 4)
           Call REPLA(YCB$, MKI$(XVALO(Label12)), 43, 2)
             DOCAPLI$ = CoComLar$
           Call REPLA(YCB$, DOCAPLI$, 45, 18)
           Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
           Call REPLA(YCB$, MKD$(CANCE#), 65, 8)
           Call REPLA(YCB$, MKD$(0), 73, 8)
           Call REPLA(YCB$, MKD$(0), 81, 8)
           Call REPLA(YCB$, MKD$(XVALO(TICAMBIS)), 89, 8)  ' NO PONE TIPO DE CAMBIO
           Call REPLA(YCB$, MKD$(0), 97, 8)
           Call REPLA(YCB$, MKD$(0), 105, 8)
           Call REPLA(YCB$, MKD$(0), 113, 8)
           '
           Call REGAPP("CARBIMON", YCB$)
           GoTo 39
         End If
       Next JKLX&
     End If
     '
39   Call CIERRARCH("CARBIMON")
     '
   End If
   DOCUDIF = ""
   '
End Sub
'
Sub GRADEUFA(Tipo%, VA%, NRO, NCL, FFAC%, COPAGI%, TOTFAC#)
   '
   Dim FVI%(6), PLAII%(6), IM#(6), FEX(3)
   '
   LU$ = LUDAT$
   IORIG# = TOTFAC#
   '
   If Tipo% = 0 Then
     If IORIG# < 0 Then
       Tipo% = 7
     End If
   End If
   '
   If Tipo% > 3 Then
     NUPAG% = 1
     FVI%(1) = FFAC%
     IM#(1) = TOTFAC#     ' IORIG#
     GoTo 41
   End If
   '
   ' busca condicion de pago segun COPAGI%
20 Call BUCONPA(LU$, COPAGI%, CoPag$, FREF%, PLA$)
   '
   ' determina vencimientos segun total factura y condicion de pago
30 FF = FFAC%: If FF < 1 Then FF = FECHANUM("01/01/80")
   FFS$ = FECHATEX$(FF)
   PLA$ = Left$(PLA$ + String$(6, 0), 6)

   If (FREF% < 1 Or FREF% > 4) Then
           MENSERR 24, "Mal Fecha Referencia en Cond.Pago Nro." + Str$(COPAGI%)
           FREF% = 1
   End If

   If FREF% = 2 Then FREF% = 1   ' AUN NO FUNCIONA SEMANA DE FACTURA
   If FREF% <> 3 Then GoTo 35
   If XVALO(Left$(FFS$, 2)) <= 15 Then FEBAS$ = "15" + Mid$(FFS$, 3)
   If XVALO(Left$(FFS$, 2)) > 15 Then FEBAS$ = "31" + Mid$(FFS$, 3)
35 If FREF% = 1 Then FEBAS$ = FFS$
   If FREF% = 4 Then FEBAS$ = "31" + Mid$(FFS$, 3)
   '
   For i% = 1 To 6
     FVI%(i%) = 0
   Next i%
   '
   FEFA = FECHANUM(FEBAS$)
   For i% = 1 To 6
     PLAI% = Asc(Mid$(PLA$, i%, 1))
     PLAII%(i%) = PLAI%
     If i% > 1 Then
       If (PLAII%(i%) < PLAII%(i% - 1) Or PLAII%(i%) <= 0) Then
         GoTo 31
       End If
     End If
     '
     FEXU = FEFA + PLAI%
     While DIENTRE(FEFA, FEXU) < PLAI%
       FEXU = DIPOST(FEXU)
     Wend
     '
     FVI%(i%) = FEXU
     NUPAG% = i%
   Next i%

31 FVIJ% = FVI%(NUPAG%)
   If NUPAG% < 2 Then FVIJ% = FFAC%
   '
   ' determina importes segun numero de pagos
40 If NUPAG% < 1 Then NUPAG% = 1
   IM#(1) = (Int(100 * IORIG# / NUPAG% + 0.5)) / 100

   For H% = 2 To NUPAG% - 1
     IM#(H%) = IM#(1)
   Next H%
   IM#(NUPAG%) = IORIG#

   For H% = 1 To NUPAG% - 1
     IM#(NUPAG%) = IM#(NUPAG%) - IM#(H%)
   Next H%
   '
   ICANCE# = IORIG# - TOTFAC#
   For H% = 1 To NUPAG%
     If ICANCE# > 0 Then
       If ICANCE# >= IM#(H%) Then
           ICANCE# = ICANCE# - IM#(H%)
           IM#(H%) = 0
         Else
           IM#(H%) = IM#(H%) - ICANCE#
           ICANCE# = 0
       End If
     End If
   Next H%
   '
41 NuorVen% = 0
   SIGX% = 1: If Tipo% > 3 Then SIGX% = (-1)
   For i% = 1 To NUPAG%
     If Abs(IM#(i%)) >= 0.005 Then
       FechEmi = CVDAT(FFAC%)
       IBruTot# = IORIG#
       NuorVen% = NuorVen% + 1
       FeVenc = CVDAT(FVI%(i%))
       ImVenc# = (Abs(IM#(i%))) * SIGX%
       Call GRADEUPEN
     End If
   Next i%
   FVENCIM = CVDAT(FVI%(1))
   If IsDate(FVENCIM) = False Then
     FVENCIM = ""
   End If
   '
End Sub
'

Sub GRAREMITO(OKX%)
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = XVALO(Text1.TEXT)
   '
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("CONSICLI")
   '
   AINDI% = 0
   CMO$ = TRIM$(Control$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   '
   REMI$ = "RT." + TRIM$(Text10.TEXT)
   NROO& = XVALO(Mid$(REMI$, 4))
   FF% = FECHANUM(Text9.TEXT)
   DEPOSALI% = DEPOEXPE%
   '
   PREFIREM$ = Control$("", "PREFIREM")
   If PREFIREM$ = "" Then PREFIREM$ = "0001"
   'ESTO SE AGREGA PORQUE LOS REMITOS DE LAS
   'FACTURAS ELECTRONICAS TIENEN EL PUNTO DE VENTA DE LA FACTURA
   ' If FACELEC% = 1 And VERIFSALIDA("NOMESS") < 3 Then PREFIREM$ = Label12
   If FACELEC% > 0 And FACELEC% < 3 Then PREFIREM$ = Label12
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "ASEMBLI") > 0 Then PREFIREM$ = Control$("", "PREFIREM")

   
   '
   NroRemi$ = TRIM$(Mid$(REMI$, 4))
   While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
   NroRemi$ = CMO$ + "." + PREFIREM$ + "-" + NroRemi$
   '
   HO% = HOY(HOS$)
   HP = 100 * XVALO(Left$(Time$, 2)) + XVALO(Mid$(Time$, 4, 2))
   User% = USNBR% Mod 100
   CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(User%) + String$(2, 0)
   '
   RECONSI$ = REGBLAN$("CONSICLI")
   Call REPLA(RECONSI$, String$(12, 0), 109, 12)
   Call REPLA(RECONSI$, MKI$(NCLIE%), 1, 2)
   Call REPLA(RECONSI$, MKI$(FF%), 3, 2)
   Call REPLA(RECONSI$, NroRemi$, 5, 18)
   Call REPLA(RECONSI$, String$(12, 0), 109, 12)
   Call REPLA(RECONSI$, CTRXX$, 122, 7)
   '
   REFE$ = rasocli$(NCLIE%)
   '
   YXI% = 0
   FIN& = ULTREG&("!CARDEFAC")
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", i&)
     CANTAA = CVS(Left$(X$, 4))
     If CANTAA > 0 Then
       CI% = CVI(Mid$(X$, 5, 2))
       If CI% > 0 Then
         NPEDID& = XVALO(Mid$(X$, 371, 8))
         COD$ = CODEXT$(CI%)
         PRE# = CDbl(CVS(Mid$(X$, 429, 4)))
         CMOV$ = "RT"
         If Mid$(X$, 461, 1) = "R" Then CMOV$ = "*R" ' REPARACION
         DOPRI$ = REMI$
         DOSEC$ = ""
         LOTE$ = ""
         SIG% = (-1)
         '
         'Call MOVISTO(FF%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, PRE#, "$", NCLIE%, DEPOSALI%, SIG% * CANTAA)
         '
         YXI% = YXI% + 1
         ZZ$ = RECONSI$
         Call REPLA(ZZ$, TRIM$(REMI$) + "-" + TRIM$(Str$(YXI%)), 23, 12)
         Call REPLA(ZZ$, MKI$(CI%), 35, 2)
         Call REPLA(ZZ$, MKD$(CDbl(Abs(CANTAA))), 37, 8)
         Call REPLA(ZZ$, MKS$(NPEDID&), 109, 4)
         Call REPLA(ZZ$, MKD$(PRE#), 113, 8)
         Call REPLA(ZZ$, Chr$(MONEMI%), 121, 1)
         Call REGAPP("CONSICLI", ZZ$)
         '
       End If
     End If
   Next i&
   '
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("CONSICLI")
   Screen.MousePointer = 1
   OKX% = 1
   Exit Sub
   '
   '
98 Call CIERRARCH("MOVISTO")
   Call CIERRARCH("CONSICLI")
   Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub
'
Sub APROCOPIAS(OKX%)
   '
   TIPCOM$ = TRIM$(REPLACAR$(Label7.Caption, ":", ""))
   CPAG% = XVALO(TEXT21(1).TEXT)
   CATIVAS$ = Label13.Caption
   EPAC$ = TRIM$(UCase$(EMPREAC$))

   '
   TCMP$ = TRIM$(TIPCOM$)
   For KKI% = 0 To 13
     FINFACTUR07.Text2(KKI%) = ""
   Next KKI%
   '
   TICOMPROBAN$ = TCMP$
   FINFACTUR07.Label2.Caption = TCMP$
   FINFACTUR07.Label1(0) = TCMP$
   FINFACTUR07.Label4.Caption = CATIVAS$
   '
   CODFOR$ = TRIM$(Control$("", "FORFAC-" + Label13.Caption))
   NROCON& = XVALO(Text11.TEXT)
   NCN$ = TRIM$(Str$(NROCON&))
   While Len(NCN$) < 8: NCN$ = "0" + NCN$: Wend
   PREFI$ = Label12.Caption
   NCN$ = PREFI$ + "-" + NCN$
   NUCOMPROBAN$ = NCN$
   FINFACTUR07.Label3.Caption = NCN$
   '
5  COPIASFAC% = XVALO(Control$("", "COPIAFAC"))
   COPIASREM% = XVALO(Control$("", "COPIAREM"))
   If COPIASFAC% < 1 Then COPIASFAC% = 1
   If COPIASREM% < 1 Then COPIASREM% = 1
   
'   If InStr(EPAC$, "DEL CIELO") < 1 Then ' SI NO ES DEL CIELO (PARA QUE NO PONGO 1 PARA HACER REMITO)
   If Control("FACTURA", "SISTMANT") = "SI" Then
      COPIASREM% = 0
   End If
   '
   If TIFADOC% > 2 Then
      'COPIASFAC% = 2: COPIASREM% = 0
      COPIASREM% = 0           ' POR TEMA CANTIDAD DE COPIAS FRIOLATINA
   End If
   '
   If TIFADOC% = 2 And MODODOC% = 1 Then
      COPIASREM% = 0
   End If
   '
   If CONSICLI% = 1 Or REPRIFAC% = 1 Then   ' SI REPITE IMPRESION
      COPIASREM% = 0
   End If
   '
   'VALIDACION DE PEDIDOS CON MOTIVO
   X1$ = REGLEIDO$("!CARDEFAC", 1)
   NPED& = XVALO(TRIM$(Mid$(X1$, 371, 16)))
   If NUMOTIVO%(NPED&) > 0 Then
      COPIASFAC% = 0
   End If
   '
   FINFACTUR07.Text2(0) = TRIM$(Str$(COPIASFAC%))
   FINFACTUR07.Text2(1) = TRIM$(Str$(COPIASREM%))
   FINFACTUR07.Text2(2) = " "
'   'VALIDACION DE PEDIDOS CON MOTIVO
'   X1$ = REGLEIDO$("!CARDEFAC", 1)
'   NPED& = XVALO(TRIM$(Mid$(X1$, 371, 16)))
'   If NUMOTIVO%(NPED&) > 0 Then FINFACTUR07.Text2(0) = ""
'   '
   'FINFACTUR07.Text2(3) = TRIM$(Str$(Int(TICAMBX# * XVALO(Total(2).TEXT)))) + ".="
   '\\\OT-05-0492-WAR-11/07/05///
     COVADE = XVALO(Control("", "COEVADEC")) / 100 'COEFICIENTE DE VALOR AGREGADO
     If COVADE < 0.01 Then COVADE = 1
     FINFACTUR07.Text2(3) = TRIM$(Str$(Int(COVADE * TICAMBX# * XVALO(Total(2).TEXT)))) + ".="
   '\\\OT-05-0492-WAR-FIN
   FINFACTUR07.Label5.Caption = ""
   For U& = 0 To 8
      FINFACTUR07.Text1(U&).TEXT = ""
   Next U&
   '\\\OT-06-0114-WAR-03/04/06///
   FINFACTUR07.Text4.TEXT = ""
   FINFACTUR07.Text5.TEXT = ""
   'LIMPIO LOS TEXT DE DATOS DE LA EXPO
   For JX% = 4 To 11
    FINFACTUR07.Text2(JX%) = ""
   Next JX%
   '\\\OT-06-0114-WAR-FIN///
   TTXX$ = ""
   '
   FINFACTUR07.Label6.Enabled = False
   FINFACTUR07.CLAREDOL = "0"
   If MONEMI% = 1 Then
     FINFACTUR07.Label6.Enabled = True
   End If
   '
   FINFACTUR07.Label6.Visible = False
   FINFACTUR07.Text3 = ""
   FINFACTUR07.Frame5.Visible = False
   '\\\OT-06-0125-WAR-03/04/06///
   FINFACTUR07.Height = 6690
   'SI FACTURA DE EXPO => SE MUESTRA DATOS DE EXPORTACIÓN
   NC11 = XVALO(FORFASQL.Text1.TEXT)
   If PIVACLI%(NC11) = 5 Or PIVACLI%(NC11) = 7 Then
     FINFACTUR07.Height = FINFACTUR07.Height + 2265 '(8955)
   End If
   'CACULA EL PESO EN KILOS NETO
   PST = 0
   P1 = 0
   PT = 0
   For UKK& = 1 To ULTREG&("!CARDEFAC")
     XKK$ = REGLEIDO$("!CARDEFAC", UKK&)
     CI% = CVI(Mid$(XKK$, 5, 2))
     CANTI1 = CVS(Mid$(XKK$, 1, 4))
     P1 = PESUNI(CI%)
     PST = P1 * CANTI1
     PT = PT + PST
   Next UKK&
   FINFACTUR07.Text2(11) = FORMATNUM$(PT, "F8.1")
   '\\\OT-06-0125-WAR-FIN///
   '18/04/07 (OT-07-0162) 'PARA VER LA COTIZACION DEL DOLAR CUANDO EMITIMOS UNA FACTURA EN PESOS
   If MONEMI% <= 1 Then
      NLIPRE% = XVALO(FORFASQL.TEXT21(0))
      If MONELISTA%(NLIPRE%) = 2 Then
         If InStr(TRIM$(UCase$(EMPREAC$)), "EYM") < 1 Then ' SI NO ES EYM INGRESA
           TTXX$ = "Emitido en " + DEMONECO$(MONEMI%) + " a cotización " + DEMONECO$(2) + " de " + TRIM$(FORMATNUM$(TICAMONE(2), "F9.4")) + ". "
           FINFACTUR07.Text4 = TTXX$
         End If
      End If
   End If
   
   FLEYE% = VARIDOC% - 1
   If TRIM(FINFACTUR07.Text4) <> "" Then FINFACTUR07.Text4 = FINFACTUR07.Text4 + Chr(10) + Chr(14)
   RFF$ = RECFILT$("LEYFAC", 1, MKI$(FLEYE%))
   'LEE OBSERVACIONES Y LAS CARGA EN LA FACTURA CORRESPONDIENTE
   For U& = 1 To Len(RFF$) Step 4
     RELE& = CVS(Mid$(RFF$, U&, 4))
     LEYY$ = REGLEIDO$("LEYFAC", RELE&)
     LEYY$ = DECODIFICA_LEYENDA$(LEYY$)
     FINFACTUR07.Text4 = FINFACTUR07.Text4 + " " + TRIM(Mid$(LEYY$, 3)) + vbCrLf
   Next U&
   
   'LEE OBSERVACIONES Y LAS CARGA EN EL REMITO
   RFF$ = RECFILT$("LEYFAC", 1, MKI$(3))
   For U& = 1 To Len(RFF$) Step 4
      RELE& = CVS(Mid$(RFF$, U&, 4))
      LEYY$ = REGLEIDO$("LEYFAC", RELE&)
      FINFACTUR07.Text5 = TRIM(FINFACTUR07.Text5) + " " + TRIM(Mid$(LEYY$, 3))
   Next U&
   
   If InStr(1, EMPREAC$, "GABAL") > 0 Then
        For U& = 1 To ULTREG&("!CARDEFAC")
            X$ = REGLEIDO$("!CARDEFAC", U&)
            NROPEDXX& = XVALO(TRIM$(Mid$(X$, 371, 16)))
            '
            TXTX$ = TRIM$(SAFETEXT$(VALOBADA("PEDENCA", "OBSERVA", "NROPED=" & NROPEDXX&)))
            If TXTX$ <> "" Then
                 If TRIM$(FINFACTUR07.Text4) <> "" Then FINFACTUR07.Text4 = FINFACTUR07.Text4 & vbCrLf
                 FINFACTUR07.Text4 = FINFACTUR07.Text4 & TXTX$
            End If
        Next U&
   End If
   '18/04/07 (FIN)
   '
' Se suprime esta vieja modalidad de obtencion de CAE ya que a partir
' de diciembre 2010 la factura se graba siempre con numero provisorio
' y recién despues de grabada se obtiene el CAE
'   FINFACTUR07.Label11 = ""
'   If FACELEC% = 1 Then
'      If VERIFSALIDA%("NOMESS") > 0 Then
'      If Label13 <> "E" Then
'         FINFACTUR07.Command1.Top = 5655
'         FINFACTUR07.Command1.Height = FINFACTUR07.Command2.Height + FINFACTUR07.Command2.Top - FINFACTUR07.Command1.Top
'         FINFACTUR07.Command1.Enabled = False
'         FINFACTUR07.Command2.Enabled = True
'
'       End If
'     End If
'   End If
   '
   If Control("FACTURA", "VERIVEN") = "SI" Then
    If XVALO(TEXT21(2).TEXT) < 1 Then ' VALIDACION DE VENDEDOR
     fech = Text9
     TIPFACT$ = "FC-"
     If Label7 = "N.Débito:" Then TIPFACT$ = "ND-"
     If Label7 = "N.Crédito:" Then TIPFACT$ = "NC-"
     
     NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend

     NFACT$ = TIPFACT$ & Label12 & "-" & NFCZ$ & "-" & Label13
     NC1& = XVALO(Text1)
     RazonSoc$ = TRIM$(Text2)
     Call REPLA(VDEF$, MKI$(CVINT(fech)), 1, 2) ' INFORMACION PARA EL OBJETO
     Call REPLA(VDEF$, NFACT$, 3, 18)
     Call REPLA(VDEF$, MKS$(NC1&), 21, 4)
     Call REPLA(VDEF$, RazonSoc$, 25, 30)
    '
30   OBX$ = OBJPLA$("ACTUVENDE", VDEF$) ' ABRE EL OBJETO
     If OBX$ = "" Then
        Call COMUNI("Debe Asignar un Vendedor a La Factura")
        GoTo 30
     End If
     '
     On Error Resume Next ' POR SI INGRESA UN VENDEDOR > 255
     VDDOR% = CVI(Mid$(OBX$, 55, 2)) ' VALIDA EL NUMERO DE VENDEDOR
     If VDDOR% < 1 Or ECOARCH$("VENDEDOR", Chr$(VDDOR%)) = "" Then
       On Error GoTo 0
       Call COMUNI("Vendedor No Válido")
       VDEF$ = OBX$
       GoTo 30
     End If
     On Error GoTo 0
     TEXT21(2) = VDDOR%
    End If
   End If
   '
   'Limpio campos de gastos de exportacion
   FINFACTUR07.Text2(15) = ""
   FINFACTUR07.Text10 = ""
   '
   
   ' ENVIO POR MAIL DE FACTURA
10 FINFACTUR07.Check1.Value = 0
   FINFACTUR07.Label1(14) = ""
   If Control("FACTURA", "ENVFAMAIL") = "SI" Then FINFACTUR07.Check1.Value = 1 'SI ESTA CONFIGURADO, ENTONCES MARCA COMO SIEMPRE ENVIA
   CONDI$ = "Nume_Cli =" & XVALO(FORFASQL.Text1)
   MAIX$ = TRIM$(VALOBADA("DEUDOR12", "MAILFACTURA_COB", CONDI$, "NOMESS"))
   If MAIX$ <> "" Then ' SI TIENE MAIL LO REGISTRA EN EL LABEL
      FINFACTUR07.Label1(14) = MAIX$
   Else ' SI NO TIENE, DESMARCA EL CHECK PARA ENVIO Y AVISA AL USUARIO QUE NO TIENE MAIL
      FINFACTUR07.Label1(14) = "Falta Correo Electrónico"
      FINFACTUR07.Label1(14).ToolTipText = "Configurar Correo de Mail de Cobranzas desde Maestros de Clientes"
      FINFACTUR07.Check1.Value = 0
   End If
   
   If InStr(EPAC$, "MEDITEA") > 0 Or InStr(EPAC$, "EYM") > 0 Then
     If XVALO(Label31) > 0 Then FINFACTUR07.Text2(1) = 1 'SI ES REPARACION POR DEFECTO REMITO 1
   End If
   
   FINFACTUR07.Show 1

   Refresh
   '
   If FINFACTUR07.Label5.Caption <> "APRO" Then
      If COMALTER%("Realmente Quiere CANCELAR este Documento\\No, Continuar\Si, Cancelar") <> 2 Then GoTo 10
      OKX% = (-1)
      Exit Sub
   End If
   '
' aqui es donde realmente se pide el CAE segun el punto de venta
   FACELECOMUN% = XVALO(Control("FACTURA", "PUVENELE"))
   FACELEBIENCAP% = XVALO(Control("FACTURA", "PUVENCAP"))
   FACELEEXPO% = XVALO(Control("FACTURA", "PUVEFAEX"))
   PVTMODOCREDITO% = XVALO(Control("FACTURA", "PUVECRED"))
   PTOVTA% = XVALO(Label12.Caption)
   '
   If FACELEC% > 0 And (PTOVTA% = FACELECOMUN% Or PTOVTA% = PVTMODOCREDITO%) Then
        If PTOVTA% = FACELECOMUN% Then
            If EXISTE%("WSFEV1_CLIENTE.EXE") = 1 Then
                Call SOLICITARCAE1_V1
            Else
                Call SOLICITARCAE1
            End If
        Else
            If EXISTE%("WSFEV2_CLIENTE.EXE") = 1 Then
                Call SOLICITARCAE1_V1
            End If
        End If
   End If
   '
   If FACELEC% > 0 And PTOVTA% = FACELEBIENCAP% Then Call SOLICITARCAE_WSBFE1
   '
   If FACELEC% > 0 And PTOVTA% = FACELEEXPO% Then
50      DATFAEX07.Show 1
        If DATFAEX07.Label17 = "OK" Then
           Call SOLICITARCAE_WSFEX1
        Else
           If COMALTER%("Realmente Quiere CANCELAR este Documento\\No, Continuar\Si, Cancelar") <> 2 Then GoTo 50
           OKX% = (-1)
           Exit Sub
        End If
   End If
   '
   COPIASFAC% = XVALO(FINFACTUR07.Text2(0))
   If COPIASFAC% < 0 Or COPIASFAC% > 9 Then GoTo 5
   COPIASREM% = XVALO(FINFACTUR07.Text2(1))
   If COPIASREM% > 9 Then GoTo 5
   If REPRIFAC% = 1 Then COPIASREM% = 0
   If COPIASREM% < 1 And FINFACTUR07.Check8.Value < 1 Then
      Text10.TEXT = ""
   End If
   '
End Sub
Function DECODIFICA_LEYENDA$(TEXTO1$)
    '
    DECODIFICA_LEYENDA$ = TEXTO1$
    If InStr(TEXTO1$, "[|") < 1 Then Exit Function
    If InStr(TEXTO1$, "|]") < 1 Then Exit Function
    '
    'SE LE DEBERA PASAR CON LOS PARAMETROS EL NUMERO DE OREDEN DE LA CUENTA QUE DESEAN QUE SE MUESTRE TANTO LA COTIZACION, COMO EL IMPOENDOLARES
    'EJEMPLO: [|COTIZADOLAR2|], SI NO TIENE PARAMETROS TOMA EL MISMO QUE MONEMI
    'SI NO SE LE PASAN LOS NUMEROS FINALES EN EL PARAMETRO TOMA LA MONEDA DE EMISION DE LA FACTURA
    'VALIDA SI VIENE CON NUMERACION DE PARAMETROS
    '
    Call TEXTOLOTES$(TEXTO1$, "[|", "|]")
    For i% = 1 To UBound(CADENATEX$)
       LOTE$ = CADENATEX$(i%)
       MON% = XVALO(Right$(LOTE$, 1))  'VALIDA SI VIENE CON NUMERO DE ORDEN EL PARAMETRO
       If MON% > 0 Then LOTE$ = Left$(LOTE$, Len(LOTE$) - 1) 'SI TIENE NUMERO TOMA SOLO LA LEYENDA PARA EL CASE
       '
       Select Case LOTE$
          Case "COTIZADOLAR"
            If MON% < 1 Then MON% = MONEMI%
            valor$ = TRIM$(FORMATNUM$(XVALO(TICAMONE(MON%)), "F12.2"))
            TEXTO1$ = REPLACAR(TEXTO1$, LOTE$, valor$)
          '
          Case "IMPOENDOLARES"
            If MON% < 1 Then MON% = MONEMI%
            TOTAL6 = XVALO(FORFASQL.Total(6))
            COTIZ = TICAMONE(MON%)
            If MON% <> MONEMI% Then
               ' SI LA MONEDA DE ORIGEN ES <> A LA MONEDA PASADA COMO PARAMETRO TRANSFORMA A $ Y LUEGO A LA MONEDA PARAMETRO
               valor$ = TOTAL6 * TICAMONE(MONEMI%)
               valor$ = TRIM$(FORMATNUM$(valor$ / COTIZ, "F12.2"))
            Else
               valor$ = TRIM$(FORMATNUM$(TOTAL6, "F12.2")) ' SI ES LA MISMA MONEDA DIRECTAMENTE NO RECALCULA
            End If
            SIMBO$ = SIMBOLPES$(MON%) 'CONCATENA EL SIMBOLO DE LA MONEDA PARAMETRO
            valor$ = SIMBO$ & " " & valor$
            TEXTO1$ = REPLACAR(TEXTO1$, LOTE$, valor$)
       End Select
    Next i%
    '
    TEXTO1$ = REPLACAR(TEXTO1$, "[|", "")
    TEXTO1$ = REPLACAR(TEXTO1$, "|]", "")
    DECODIFICA_LEYENDA$ = TEXTO1$
    '
End Function
'
Private Sub Total_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Index = 0 Then
       On Error Resume Next
       PORDESCU.SetFocus
       On Error GoTo 0
     ElseIf Index = 3 Then
      ' calcula importe no gravado
       If Label6.Caption = "(T.Libre)" Then
         BASIVA(0) = TRIM$(FORMATNUM(XVALO(Total(3)) / ((XVALO(PORIVA(0))) / 100), "F13.2"))
       End If
     ElseIf Index = 4 Then
      ' calcula importe no gravado
       If Label6.Caption = "(T.Libre)" Then
         BASIVA(1) = TRIM$(FORMATNUM(XVALO(Total(4)) / ((XVALO(PORIVA(1))) / 100), "F13.2"))
       End If
     Else
       On Error Resume Next
       Command14.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Total_LostFocus(Index As Integer)
   'If Index = 0 Then
     TTT1 = XVALO(Total(Index).TEXT)
     Total(Index) = CSTRING$(MKS$(TTT1), 3, 12, 2, 2)
   'End If
End Sub

Sub CARDETFAC()
   '
   Static AGRUPAENTREGA% 'ASIGNA A AGRUPAENTREGA% SEGUN LA CONFIGURACION DEL SETUP
   If AGRUPAENTREGA% = 0 Then
     AGRUPAENTREGA% = (-1)
     If Control$("FACTUREM", "AGRUPEDI") = "TRANSUCU" Then AGRUPAENTREGA% = 1
   End If

   HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC/ARCHECO(2)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle de Facturación (Mat.Codificado)"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left - 50))
   '
   If FACTLISTAEMPAQUE& > 0 Then
     ATRI$ = ATRI$ + "/POSEDIT(1)=2"
     ATRI$ = ATRI$ + "/POSEDIT(2)=2"
     ATRI$ = ATRI$ + "/POSEDIT(3)=2"
     ATRI$ = ATRI$ + "/POSEDIT(4)=2"
     ATRI$ = ATRI$ + "/POSEDIT(5)=2"
     ATRI$ = ATRI$ + "/POSEDIT(5)=2"
     ATRI$ = ATRI$ + "/POINDEL=NO" ' DESHABILITA EL INSERTAR Y EL DEL
   End If
   '
   'AGREGADO PARA CONSIDERAR NOTA DE DEBITO POR DIFERENCIA DE CAMBIO DESHABILITA EDICION
   ESNOTADEBITOAUTO% = 0
   TXX1$ = REGLEIDO$("!CARDEFAC", 1)
   TXX2$ = Mid$(TXX1$, 7, 64)
   If InStr(TXX1$, "Nota de Débito x Dif. Cambio (Recibo: ") > 0 Or InStr(TXX1$, "Nota de Crédito x Dif. Cambio (Recibo: ") > 0 Then
     ESNOTADEBITOAUTO% = 1 ' SE USA MAS ABAJO PARA QUE NO SE PISE 'ATRI$'
     ATRI$ = "/NC "
     ATRI$ = ATRI$ + "/TITUPAN=Detalle x Diferencia de Cambio"
     ATRI$ = ATRI$ + "/POSEDIT(1)=2"
     ATRI$ = ATRI$ + "/POSEDIT(2)=2"
     ATRI$ = ATRI$ + "/POSEDIT(3)=2"
     ATRI$ = ATRI$ + "/POSEDIT(4)=2"
     ATRI$ = ATRI$ + "/POSEDIT(5)=2"
     ATRI$ = ATRI$ + "/POINDEL=NO" ' DESHABILITA EL INSERTAR Y EL DEL
   End If
   '============================================================================================
''   'MODIFICACION N.CREDITO PARA QUE EL REMITO QUEDE COMO EN CONSIGNACION
''   Y1Y$ = REGLEIDO$("!CARDEFAC", 1)
''   LETX$ = Mid$(Y1Y$, 461, 1)
''   If LETX$ = "N" And TIFADOC% = 6 Then
''          ATRI$ = ATRI$ + "/POSEDIT(6)=2" ' DETECTA QUE ES NOTA DE CREDITO ANULA FACTURA PARA QUE VUELVA PARA ATRAS EL REMITO
''          Y1Y$ = REGLEIDO$("!CARDEFAC", ULTREG&("!CARDEFAC"))
''          NUFACTU$ = Mid$(Y1Y$, 7, 64)
''          If Left$(NUFACTU$, 3) = "FC-" Then Call SETULTREG("!CARDEFAC", ULTREG&("!CARDEFAC") - 1)
''          Label31 = NUFACTU$
''   End If
   '
   'FIN MODIFICACION N.CREDITO PARA QUE EL REMITO QUEDE COMO EN CONSIGNACION
10 Timer1.Enabled = True
   PRF$ = "FACONGUI"
   If CONSICLI% = 0 Then
      PRF$ = "FACDOSI"
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "FRIOLA") > 0 Then
        PRF$ = "FACFRIOL"
      End If
      If MODESCU% = 1 Then PRF$ = "FACPORDE"
      If InStr(EPAC$, "DOSIVAC") > 0 Then PRF$ = "FACDOSV" ' FACTURACION SOBRE PEDIDO PREPARADO
      If InStr(EPAC$, "J Y M") > 0 Then PRF$ = "FACDOSJ" ' FACTURACION SOBRE PEDIDO PREPARADO
   End If
   
   If MODODOC% = 1 Then
      PRF$ = "FAC-LIBR"
      If ESNOTADEBITOAUTO% < 1 Then
        ATRI$ = "/NC"
        ATRI$ = ATRI$ + "/TITUPAN=Detalle de Facturación (Texto Libre)"
        ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
        ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
        If XVALO(Label31) > 0 Then ' IMPIDE QUE SE MODIFIQUEN PRECIOS CUANDO SE FACTURA UNA REPARACION
          ATRI$ = ATRI$ + "/POSEDIT(3)=2"
          ATRI$ = ATRI$ + "/POSEDIT(4)=2"
        End If
      End If
   End If
   'List1.Clear
   MSF.Rows = 1
   '
   If EsPorPedido& < 1 Then
        XXX% = VENTABU%(PRF$ + ATRI$)
      Else
        XXX% = 0                   ' SI ES POR PEDIDO
   End If
   '
   Call CALTOFAC
   '
   Timer1.Enabled = False
   '
   NPEDID& = 0
   If XXX% < 0 Then
      Call CIERRARCH("!CARDEFAC")
      Call BLANARCH("!CARDEFAC")
      Call BLANFORMU
      Total(0).TEXT = ""
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   If ULTREG&("!CARDEFAC") < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   If CONSICLI% < 1 Then
     If InStr(EMPREAC$, "PREMOTEC") > 0 Then
         'SE COMENTA SEGUN MODIFICACION PRESUPUESTO 13/04/18
'       Call ASIGNALOTES_PREMOTEC("!CARDEFAC", RESUL$)
     Else
       Call ASIGNALOTES("!CARDEFAC", RESUL$)    ' SI NO EXISTEN ARTICULOS TRAZABLE NO INGRESA
       If RESUL$ <> "APRO" Then GoTo 10
     End If
   End If
   '
   If (InStr(EMPREAC$, "MEDITEA") > 0 Or InStr(EMPREAC$, "EYM") > 0) And CONSICLI% = 0 Then  ' EXCLUYE LA FACTURACION DE CONSIGNACIONES
      Call CARNUMSER(OK%)                    ' CARGA NUMEROS DE SERIE
      If OK% <> 1 Then GoTo 10               ' CANCELA EN VENTANA N/SERIE VUELVE ANTERIOR
   End If
   '

   Screen.MousePointer = 11   ' Registrando FACTURA ...
   'List1.Clear
   MSF.Rows = 1
   FIN& = ULTREG&("!CARDEFAC")
   TXPEDI$ = "": Label28 = ""
   PORLETRA% = 0:  If Control$("ADMIPED", "MOCOMIS") = "LETRA" Then PORLETRA% = 1
   
   '
   PASOYA% = 0
   For U& = 1 To FIN&
      Y$ = REGLEIDO$("!CARDEFAC", U&)
      NUOC$ = Mid$(Y$, 355, 16)
      If NUOC$ <> "" Then
         NUOCOMPRA$ = NUOC$ + ","
      End If
      
      
      If MODODOC% < 1 Then
           CI% = CVI(Mid$(Y$, 5, 2))
           '
           CAN = CVS(Mid$(Y$, 1, 4))
           If CAN > 999999 Then
               Call MENSERR(24, "Cantidad Excesiva")
               GoTo 10
             ElseIf Abs(CAN) < 0.1 Then
               If CI% > 0 Then
                 Call MENSERR(24, "Falta Cantidad")
                 Call REPLA(Y$, MKS$(0), 1, 4)
                 Call GRAREG("!CARDEFAC", Y$, U&)
                 GoTo 10
               End If
           End If
           '
           If Abs(CAN - REDON_INDIVIS(CI%, CDbl(CAN), "VENTA")) >= 0.05 Then
              TTXX$ = "Cantidad Ajustada a Lotes Indivisibles\en Código '" + CODEXT$(CI%) + "'"
              Call MENSERR(24, TTXX$)
              Call REPLA(Y$, MKS$(REDON_INDIVIS(CI%, CDbl(CAN), "VENTA")), 1, 4)
              Call GRAREG("!CARDEFAC", Y$, U&)
              GoTo 10
           End If
           '
           NPEDID& = XVALO(Mid$(Y$, 371, 8))
           If NPEDID& > 0 Then
              TYPEDI$ = FORMATNUM$(NPEDID&, "F10.0")
              If InStr(TXPEDI$, TYPEDI$) < 1 Then
                 TXPEDI$ = TXPEDI$ + TYPEDI$
                 If PORLETRA% > 0 Then
                    If TRIM$(Label28) = "" Then
                      LETRAPE$ = Chr$(64 + XVALO(VALOBADA("PEDENCA", "PorComVend", "NroPed = " & NPEDID&)))
                      If Asc(LETRAPE$) > 64 Then Label28 = LETRAPE$
                    End If
                 End If
              End If
           End If
           '
         Else     ' FACTURA DE TEXTO LIBRE
           CAN = CVS(Mid$(Y$, 1, 4))
           If CAN < 0.1 Then
              CAN = 1
              Call REPLA(Y$, MKS$(1), 1, 4)
              Call GRAREG("!CARDEFAC", Y$, U&)
              Call MENSERR(24, "Faltaba Cantidad - Ajustada a '1,0'")
              GoTo 10
           End If
           '
           PREUNI# = CVS(Mid$(X$, 429, 4))
           TOT# = CVD(Mid$(X$, 433, 8))

           PREUNI# = CVS(Mid$(Y$, 429, 4))
           TOT# = CVD(Mid$(Y$, 433, 8))
'           If (PREUNI# > 0.005 And TOT# < 0.005) Or (PREUNI# < 0.005 And TOT# > 0.005) Then
'              Call COMUNI("Debe Ingresar Precio Unitario e Importe")
'              GoTo 10
'           End If
           
           VARIA1# = Abs(CAN * PREUNI# - TOT#)
           VARIA2# = XVALO(FORMATNUM(VARIA1#, "F12.3"))
           If VARIA2# > 0.005 Then
              'TOT# = ROUND(CAN * PREUNI)
              'Call REPLA(y$, MKD$(TOT#), 433, 8)
              'Call GRAREG("!CARDEFAC", y$, U&)
              'Call MENSERR(24, "Importe Incorrecto - Ajustado")
              'GoTo 10
              If CAN > 0 Then
                PRUN# = TOT# / CAN
                Call REPLA(Y$, MKS$(PRUN#), 429, 4)
                Call GRAREG("!CARDEFAC", Y$, U&)
              End If
           End If
      End If
      '
      ' VALIDACION PARA MODALIDAD DE FACTURACION SOBRE PEDIDO PREPARADO
      If InStr(1, EMPREAC$, "DOSIVAC") > 0 Then
            CIXI% = CVI(Mid$(Y$, 5, 2))
            NORITPE% = CVI(Mid$(Y$, 503, 2))
            NROPED& = XVALO(Mid$(Y$, 371, 16))
            If NROPED& > 0 And CONSICLI% < 1 Then
                '
                CANPEDIDA = CVS(Mid$(Y$, 505, 4)) ' VALIDA QUE REMITIDO SEA MENOR O IGUAL AL SALDO PENDIENTE
                If CAN > CANPEDIDA + 0.005 Then
                  Call MENSERR(24, "La Cantidad del Item de la Fila " & U& & "\Supera el Saldo Pendiente del Item")
                  GoTo 10
                End If
                '
                CONDI$ = "NROPED=" & NROPED& & " AND NUORITEM=" & NORITPE% & " AND CODIINT=" & CIXI%
                CANPRE = XVALO(VALOBADA("PEDDETA", "ACUPREPA", CONDI$))
                CANENT = XVALO(VALOBADA("PEDDETA", "CANTENT", CONDI$))
                SALDOPRE = CANPRE - CANENT        ' VALIDAD QUE REMITIDO SEA MENOR O IGUAL AL PENDIENTE PREPARADO
                If CAN > (SALDOPRE + 0.005) Then
                  OPX% = COMALTER%("La Cantidad del Item de la Fila " & U& & "\Supera el Saldo Preparado.\\Continuar\Cancelar")
                  If OPX% = 1 Then
                    If DERACCE%("REMAYPRE", "Remite Cantidad Mayor a la Preparada") < 1 Then
                       GoTo 10
                    End If
                  End If
                End If
            End If
            '
      End If
      
   Next U&
   '

   Call CARLISFAC
   '
   MALCOMBINA% = 0
   If VERIPEDMO% < 1 Then MALCOMBINA% = 1   ' DISTINTO MOTIVO DE PEDIDO

   ' CUANDO SE TRATA DE FACTURACION POR CONSIGNACION,
   ' VALIDA QUE LAS CONDICIONES DE LOS PEDIDOS ASOCIADOS A LOS REMITOS SEA LA MISMA
   If CONSICLI% = 1 Then AGRUPAENTREGA% = 1

   If AGRUPAENTREGA% > 0 Then ' VALIDA SI ESTA SELECCIONADA LA OPCION DE AGRUPACION DEL SETUP
    If Len(TXPEDI$) > 0 Then
       Dim PEDDDENCA As ADODB.Recordset
       'MALCOMBINA% = 0
       LIPENOCO07.List1.Clear
       PDXA$ = "": LPRXA$ = "": CPGXA$ = "": DMEXA$ = "": LMEXA$ = "": NTREXA$ = "": TREXA$ = "": NSUENA$ = "": NMOEMA$ = "": VENDEA% = (-1)
       For KKI% = 1 To Len(TXPEDI$) Step 10
         TENCA$ = ""
         NPEDID& = XVALO(TRIM$(Mid$(TXPEDI$, KKI%, 10)))
         SQLX$ = " SELECT * FROM PedEnca"
         SQLX$ = SQLX$ + " WHERE NroPed = " & NPEDID& & ""
         Set PEDDDENCA = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With PEDDDENCA
           If Not (.EOF And .BOF) Then
             '
             NPEDI$ = (TRIM$(SAFETEXT$(!NROPED)))
             TENCA$ = AJUSTD$(NPEDI$, 6)
             '
             'MODIFICADA LINEA DE ABAJO PARA QUE TRAIGA EL DESCUENTO DE RETIRA CLIENTE
             PDX$ = TRIM$(DESCUPEDI$(NPEDID&)) 'TRIM$(SAFETEXT$(!TDescApli))
             If PDXA$ <> "" Then
                If PDX$ <> PDXA$ Then
                    MALCOMBINA% = 1
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinto Descuento.")
'                  GoTo 10
                End If
             End If
             Call REPLA(TENCA$, AJUSTI$(PDX$, 12), 103, 12)
             PDXA$ = PDX$
             '
'             LPRX$ = TRIM$(Str$(XVALO(!LISTPRE)))
'             If LPRXA$ <> "" Then
'                If LPRX$ <> LPRXA$ Then
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinta Lista de Precios.")
'                  GoTo 10
'                End If
'             End If
'             LPRXA$ = LPRX$
             '
             CPGX$ = TRIM$(Str$(XVALO(!CondPag)))
             If CPGXA$ <> "" Then
                If CPGX$ <> CPGXA$ Then
                     MALCOMBINA% = 1
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinta Condicion de Pago.")
'                  GoTo 10
                End If
             End If
             Call REPLA(TENCA$, AJUSTD$(CPGX$, 5), 95, 5)
             CPGXA$ = CPGX$
             '
             DMEX$ = TRIM$(SAFETEXT$(!DOMIENT))
             If DMEXA$ <> "" Then
                If DMEX$ <> DMEXA$ Then
                   MALCOMBINA% = 1
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinto Domicilio de Entrega.")
'                  GoTo 10
                End If
             End If
             Call REPLA(TENCA$, AJUSTI$(DMEX$, 25), 9, 37)
             DMEXA$ = DMEX$
             '
             LMEX$ = TRIM$(SAFETEXT$(!LOCAENT))
             If LMEXA$ <> "" Then
                If LMEX$ <> LMEXA$ Then
                   MALCOMBINA% = 1
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinta Localidad de Entrega.")
'                  GoTo 10
                End If
             End If
             Call REPLA(TENCA$, AJUSTI$(LMEX$, 22), 37, 22)
             LMEXA$ = LMEX$
             '
             NTREX$ = TRIM$(Str$(XVALO(!NTranspo)))
             If NTREXA$ <> "" Then
                If NTREX$ <> NTREXA$ Then
                   MALCOMBINA% = 1
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinto Numero de Transporte.")
'                  GoTo 10
                End If
             End If
             Call REPLA(TENCA$, AJUSTD$(NTREX$, 4), 61, 4)
             NTREXA$ = NTREX$
             '
             TREX$ = TRIM$(SAFETEXT$(!TranspEnt))
             If TREXA$ <> "" Then
                If TREX$ <> TREXA$ Then
                   MALCOMBINA% = 1
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinto Transporte.")
'                  GoTo 10
                End If
             End If
             Call REPLA(TENCA$, AJUSTI$(TREX$, 22), 67, 30)
             TREXA$ = TREX$
             '
             VENDE% = XVALO(!VENDED)
             If VENDEA% > (-1) Then
                 If VENDE% <> VENDEA% Then MALCOMBINA% = 1
             End If
             VENDEA% = VENDE%
             Call REPLA(TENCA$, AJUSTI$(CStr(VENDE%), 4), 115, 4)
             '
             
             NOCAMONE% = 1            ' BANDERA PARA NO CAMBIAR LA MONEDA DE EMISION
             PORDESCU.TEXT = PDXA$
             If LPRXA$ <> "" Then TEXT21(0).TEXT = LPRXA$
             TEXT21(1).TEXT = CPGXA$
             NOCAMONE% = 0
             '
             If DMEXA$ <> "" Then LENTREGA(0).TEXT = DMEXA$
             If LMEXA$ <> "" Then LENTREGA(1).TEXT = LMEXA$
               If NTREXA$ <> "" Then Text19.TEXT = NTREXA$
             If TREXA$ <> "" Then LENTREGA(2).TEXT = TREXA$
             If VENDEA% > 0 Then TEXT21(2).TEXT = VENDEA%
           End If
         End With
         LIPENOCO07.List1.AddItem (TENCA$)
       Next KKI%
       '
       If MALCOMBINA% > 0 Then
          If COMALTER%("Los Pedidos Seleccionados Tienen Condiciones Comerciales\que no Pueden Combinarse.\\Ver Detalle\Corregir Seleccion") <> 2 Then
             Screen.MousePointer = 1
             LIPENOCO07.Show 1
             GoTo 10
          Else
             GoTo 10
          End If
       End If
       '
'       If CONTROL("FACTURA", "VERIVEN") = "SI" Then ' CONTROL DE VENDEDOR
'          If VENDEA% < 1 Then
'              Call COMUNI("Se Han Encontrado Pedidos Sin Vendedor\Actualice El Vendedor en el Pedido Y Vuelva a Facturar")
'              GoTo 10
'          End If
'       End If
       '
    End If
   End If
   '
   'Se comenta la condicion para que presente la la seleccion de lotes
   'independientemente de la configuracion de facturacion
   'If TRIM$(CONTROL("FACTURA", "PREFACTU")) = "SELELOTE" Then
   'If InStr(EPAC$, "PAPELERA NICARAGUA") > 0 Then
      Screen.MousePointer = 1
      '
      FIN& = ULTREG&("!CARDEFAC")
      For UI& = 1 To FIN&
        XX$ = REGLEIDO$("!CARDEFAC", UI&)
        CCII% = CVI(Mid$(XX$, 5, 2))
        If ESTRAZABLE%(CCII%) = 1 Then
            For i = 0 To ASIGLOT07.List1.ListCount - 1
                YY$ = ASIGLOT07.List1.List(i)
                CANTASI = XVALO(Mid$(YY$, 64, 10))
                REGIST& = XVALO(Mid$(YY$, 80, 5))
                If UI& = REGIST& And CANTASI > 0 Then
                    Call REPLA(XX$, MKS$(CANTASI), 1, 4)
                    Call GRAREG("!CARDEFAC", XX$, UI&)
                End If
            Next i
        End If
      Next UI&

' AQUI AJUSTAR CANTIDADES E IMPORTES
      Call CALTOFAC
      Call CARLISFAC
'   End If
'   End If
   '
   '
99 On Error Resume Next
   Command14.SetFocus
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub
'
Sub CARLISFAC()
'   List1.Clear
   MuestraPedido% = 0
   If TRIM$(Control("FACTURA", "PREFACTU")) = "REMITO" Then MuestraPedido% = 1

   MSF.Rows = 1
   FIN& = ULTREG&("!CARDEFAC")
   For U& = 1 To FIN&
      Y$ = REGLEIDO$("!CARDEFAC", U&)
      TTXX$ = Space$(128)
      CI% = CVI(Mid$(Y$, 5, 2))
      PRECIO = CVS(Mid$(Y$, 429, 4))
      PRELI = PRELISTA#(ARCHIPRE$, CI%)
      PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
      MAREPA$ = UCase$(Mid$(Y$, 461, 1)): If MAREPA$ <> "R" Then MAREPA$ = " "
      '
      If MAREPA$ = "R" Then
        PORDESCU = ""
      End If
      '
      CAN = CVS(Mid$(Y$, 1, 4))
      '
      If CAN <= 9999 Then
        Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 1, 2), 1, 6)
        If MODODOC% = 1 Then Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 2, 2), 1, 6)
      Else
        Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
      End If
      '
      CANDEC% = 4
      Call REPLA(TTXX$, CODEXT$(CI%), 8, 15)
      Call REPLA(TTXX$, Mid$(Y$, 7, 64), 24, 26)
      PRELI = CVS(Mid$(Y$, 421, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PRELI), 3, 8, CANDEC%, 2), 50, 8)
      PORDESCUIN = CVS(Mid$(Y$, 425, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 59, 5)
      PREUNI = CVS(Mid$(Y$, 429, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, CANDEC%, 2), 64, 8)
      TOT# = CVD(Mid$(Y$, 433, 8))
      Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 4, 2), 74, 10)
      Call REPLA(TTXX$, Mid$(Y$, 461, 2), 86, 2) ' MARCA REPARACION
'      List1.AddItem TTXX$

'      MSF.Rows = U& + 1
'      MSF.TextMatrix(U&, 1) = TRIM$(Mid$(TTXX$, 1, 6))
'      MSF.TextMatrix(U&, 2) = TRIM$(Mid$(TTXX$, 8, 15))
'      MSF.TextMatrix(U&, 3) = TRIM$(Mid$(TTXX$, 24, 26))
'      MSF.TextMatrix(U&, 4) = TRIM$(FORMATNUM$(PRELI, "F12." & CANDEC%))
'      MSF.TextMatrix(U&, 5) = TRIM$(FORMATNUM$(PORDESCUIN, "F4.1"))
'      MSF.TextMatrix(U&, 6) = TRIM$(FORMATNUM$(PREUNI, "F12." & CANDEC%))
'      MSF.TextMatrix(U&, 7) = TRIM$(FORMATNUM$(TOT#, "F14.2"))
'      MSF.TextMatrix(U&, 8) = TRIM$(Mid$(Y$, 461, 2))
      '
      MSF.Rows = U& + 1
      MSF.TextMatrix(U&, 1) = TRIM$(Mid$(TTXX$, 1, 6))
      MSF.TextMatrix(U&, 2) = TRIM$(Mid$(TTXX$, 8, 15))
      MSF.TextMatrix(U&, 3) = TRIM$(Mid$(TTXX$, 24, 26))
      MSF.TextMatrix(U&, 4) = TRIM$(FORMATNUM$(PRELI, "F12." & CANDEC%))
      If MuestraPedido% < 1 Then
        MSF.TextMatrix(U&, 5) = TRIM$(FORMATNUM$(PORDESCUIN, "F4.1"))
        MSF.TextMatrix(U&, 6) = TRIM$(FORMATNUM$(PREUNI, "F12." & CANDEC%))
        MSF.TextMatrix(U&, 7) = TRIM$(FORMATNUM$(TOT#, "F14.2"))
        MSF.TextMatrix(U&, 8) = TRIM$(Mid$(Y$, 461, 2))
      Else
        'SI ESTA CONFIGURADO FACTURACION POR REMITO MUESTRA EN LA GRILLA EL NUMERO DE REMITO
        NPED& = XVALO(TRIM$(Mid$(Y$, 371, 16)))

        MSF.TextMatrix(U&, 5) = TRIM$(FORMATNUM$(NPED&, "F6.0"))
        If NPED& < 1 Then Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, U&, 5)
        MSF.TextMatrix(U&, 6) = TRIM$(FORMATNUM$(PORDESCUIN, "F4.1"))
        MSF.TextMatrix(U&, 7) = TRIM$(FORMATNUM$(PREUNI, "F12." & CANDEC%))
        MSF.TextMatrix(U&, 8) = TRIM$(FORMATNUM$(TOT#, "F14.2"))
'        If NPED& = 1459 Then
'            MSF.TextMatrix(U&, 7) = "162185.4000"
'            MSF.TextMatrix(U&, 8) = "162185.40"
'        End If
        MSF.TextMatrix(U&, 9) = TRIM$(Mid$(Y$, 461, 2))
      End If
      If GRUPIVA%(CI%) > 1 Then
         Call MENSERR(24, "Grupo de IVA no Procesable\para Código '" + TRIM$(CODEXT$(CI%)) + "'.\   \Se Procesará a la Tasa Básica del " + TRIM$(FORMATNUM$(ALIVA, "F5.2")) + " %.")
      End If
      '
   Next U&
   
   'List1.Refresh
   '
   If MODODOC% = 0 Then   'FACTURA Y CODIFICADO
    If Control("ADMIPED", "ITEMSCGO") = "SI" Then
      For i& = 1 To MSF.Rows - 1
        If XVALO(MSF.TextMatrix(i&, 7)) <= 0.01 Then
           'Call COLOREAR_GRILLA_SOLAFILA(MSF, vbYellow, GLOBALFRM_FILA&, , , 1)
           Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, vbYellow, i&)
           MSF.TextMatrix(i&, 0) = "*"
        Else
           Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, vbWhite, i&)
           MSF.TextMatrix(i&, 0) = ""
        End If
      Next i&
    End If
   End If
End Sub
'
Sub CALTOFAC()
   '
   Static CLIA$, RNI%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE > 0 Then
       PPVI% = PIVACLI%(NCLIE)
       If PPVI% = 2 Then RNI% = 1
       If PPVI% = 5 Or PPVI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   CLIACT1 = XVALO(Text1)
   '
   TASA1# = XVALO(PORIVA(0).Caption)
   TASA2# = XVALO(PORIVA(1).Caption)
   TASA3# = XVALO(PORIVA(2).Caption)
   SUNEFAC# = 0: SUNEFA1# = 0: SUNEFA2# = 0
   ' *** PONE EN CERO TODOS LOS INDICES DEL VECTOR
   If NetoMultipleImputacion% = 1 Then
        For i% = 1 To UBound(VentaNetaPorCuentaContable#)
            VentaNetaPorCuentaContable#(i%) = 0
        Next i%
   End If
   '
   Label18 = TRIM$(Str$(ULTREG&("!CARDEFAC")))
   SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
   
   'MODIFICACION DE EYM ITEM CON 2 IVAS (CARGA EL VECTOR CON LOS CODIGOS)
   For J& = 1 To 256
       CI_VECTOR_TIVA(J&) = 0
   Next J&

   For U& = 1 To ULTREG&("!CARDEFAC")
     Y$ = REGLEIDO$("!CARDEFAC", U&)
     CI_VECTOR_TIVA(U&) = CVI(Mid$(Y$, 5, 2))
   Next U&
   Call ACTUALIZAVECTORDEIVA
   
   'FIN MODIFICACION DE EYM ITEM CON 2 IVAS
   
   
   
   
   SUMACAN# = 0
   '
   
   For U& = 1 To ULTREG&("!CARDEFAC")
     '
     X$ = REGLEIDO$("!CARDEFAC", U&)
     CIXXI% = CVI(Mid$(X$, 5, 2))
     CAN# = CVS(Mid$(X$, 1, 4))
     SUMACAN# = SUMACAN# + CAN#
     PORDESCUIN# = CVS(Mid$(X$, 425, 4))
     If PORDESCUIN# < 0 Or PORDESCUIN# > 100 Then PORDESCUIN# = 0
     '
     If Abs(CAN#) > 0.05 Or MODODOC% > 0 Then
       TASARE# = TASA1#
       If TASACLICK >= 0.005 Then TASARE# = TASACLICK
       If CIXXI% > 0 Then
         CILOCAL% = CIXXI%
         TASARE# = PORCEIVA(CILOCAL%)
       End If
       If TASARE# <> CVS(Mid$(X$, 441, 4)) Then
         Call REPLA(X$, MKS$(TASARE#), 441, 4)
         Call GRAREG("!CARDEFAC", X$, U&)
       End If
       '
       PRELI# = CVS(Mid$(X$, 421, 4))
       PREUNI# = CVS(Mid$(X$, 429, 4))
       '
       If MODESCU% = 1 Then
         If PRELI# < 0.00005 Then PRELI# = PREUNI#
         PREUNI# = PRELI# * (100 - PORDESCUIN#) / 100
         Call REPLA(X$, MKS$(PREUNI#), 429, 4)
       End If
       '
       'PREUNI# = ROUND(PREUNI#, DECILISTACT%)
       '
       If CAN = 1 Then
           TOT# = XVALO(PREUNI#)
         Else
           PR1# = ROUND(PREUNI#, 4)
           TOT# = ROUND(CAN#, 1) * PR1#
       End If
       TOT# = ROUND(TOT#, 2)
       '
       If MODODOC% = 1 Then
          PREUNI# = ROUND(CVS(Mid$(X$, 429, 4)), 4)
          TOT# = ROUND(CVD(Mid$(X$, 433, 8)), 2)
       End If
       '
       Call REPLA(X$, MKD$(TOT#), 433, 8)
       SUNEFAC# = SUNEFAC# + TOT#
       '
       TASARE# = CVS(Mid$(X$, 441, 4))
       IVAITEM1# = ROUND((TOT# * (TASARE# / 100)) * APLIVA%, 2)
       Call REPLA(X$, MKS$(PREUNI#), 429, 4)
       Call REPLA(X$, MKD$(TOT#), 433, 8)
       Call REPLA(X$, MKD$(IVAITEM1#), 445, 8)
       Call REPLA(X$, MKD$(TOT# + IVAITEM1#), 453, 8)
       Call GRAREG("!CARDEFAC", X$, U&)
       '
       If Abs(TASARE# - TASA2#) < 0.01 Then
           SUNEFA2# = SUNEFA2# + TOT#
         Else
           SUNEFA1# = SUNEFA1# + TOT#
       End If
       '
     End If
   Next U&
   '
   Label25 = FORMATNUM(SUMACAN, "F8.1")
   '
   'Total(0).TEXT = CSTRING$(MKD$(SUNEFAC#), 3, 12, 2, 2)
   'If Abs(SUNEFAC# - 162185.4) < 0.1 Then SUNEFAC# = 162185.4
   Total(0).TEXT = FORMATNUM$(SUNEFAC#, "F12.2")
   Total(0).Refresh
   
   'ITEMS SIN CARGO MEDIEA
  If Control("ADMIPED", "ITEMSCGO") = "SI" Then
      CANTI_SIN_CARGO = 0
      If XVALO(Total(0)) <= 0.01 Then
        For J& = 1 To MSF.Rows - 1
          If MSF.TextMatrix(J&, 0) = "*" Then
           CANTI_SIN_CARGO = CANTI_SIN_CARGO + 1
          End If
         Next J&
      End If
      '
      'SI ESTAN TODOS LOS REGISTROS EN 0 AL ULTIMO LE AGREGA EL VALOR DE 0.01
      If CANTI_SIN_CARGO = MSF.Rows - 1 And MSF.Rows > 1 Then 'SI ESTAN MARCADOS TODOS COMO SIN CARGO
         rr& = ULTREG&("!CARDEFAC")
         Z$ = REGLEIDO$("!CARDEFAC", rr&)
         CAN = CVS(Mid$(Z$, 1, 4)) 'LEO CANTIDAD
         PRECIO = 0.01 / CAN 'LO DIVIDE POR LA CANTIDAD PARA QUE SIEMPRE DE 0.01
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
         IMPO# = CDbl(PRECIO) * CAN
         IMPO# = ROUND(IMPO#, 2)
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call GRAREG("!CARDEFAC", Z$, rr&) 'GRABO EN CARDEFAC
         Total(0) = "0.01"
         SUNEFAC# = 0.01
         Total(0).Refresh
         APLIVA% = 0
      End If
   End If
   '''''
   '
   DESC$ = TRIM$(PORDESCU.TEXT)
   PODES# = COEFDESCU(DESC$)
   DESC0$ = TRIM$(Text23.TEXT) ' DECUENTO VARIOS REALIZADO PARA GABAL
   PODES0# = COEFDESCU(DESC0$)
   TTT1# = (SUNEFAC# * PODES# / 100)
   IMPONET# = SUNEFAC# - TTT1#
   TTT1# = TTT1# + (IMPONET# * PODES0# / 100)
   'Total(1).TEXT = CSTRING$(MKD$(TTT1#), 3, 12, 2, 2)
   Total(1).TEXT = FORMATNUM$(TTT1#, "F12.2")
   '
   'TTT2# = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
   TTT2# = SUNEFAC# - TTT1#
   '
   ' CALCULO DE EMBALAJE
   PEMBA = POREMBAL(CLIACT1)
   Ltota(10) = "Importe Total Neto"
   IEMBALA = "0.00"
   If TIFADOC% < 3 Then ' SOLO PARA FACTURAS, NO N/CRED Y N/DEB
     If XVALO(PEMBA) > 0 Then 'se agrega xvalo por que devolvia "" y daba error 13
       IEMBALA = TRIM$(FORMATNUM$(PEMBA * TTT2# / 100, "F12.2"))
       Ltota(10) = "Tot.Neto (Inc." + TRIM$(IEMBALA) + " Embal.)"
       TTT2# = TTT2# + XVALO(IEMBALA)
     End If
   End If
   '
   'Total(2).TEXT = CSTRING$(MKD$(TTT2#), 3, 12, 2, 2)
   Total(2).TEXT = FORMATNUM$(TTT2#, "F12.2")
   TLABEL28 = ""
   '
   TTTX# = SUNEFA2# * (100 - PODES#) / 100
   'BASIVA(1) = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   BASIVA(1) = TRIM$(FORMATNUM$(TTTX#, "F12.2"))
   'Total(4).TEXT = CSTRING$(MKD$(XVALO(BASIVA(1)) * XVALO(PORIVA(1).Caption) * APLIVA% / 100), 3, 12, 2, 2)
   Total(4).TEXT = FORMATNUM$((TTTX# * XVALO(PORIVA(1).Caption) * APLIVA% / 100), "F12.2")
   '
   'TTTX# = (XVALO(Total(2).TEXT) - XVALO(BASIVA(1)))
   NetoIva105# = ROUND#(TTTX#, 2)
   ImporteNeto# = ROUND#(TTT2#, 2)
   NetoIva21# = ImporteNeto# - NetoIva105#
   TTTX# = NetoIva21#
   'TTTX# = TTT2# - TTTX#
   '
   'BASIVA(0) = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   BASIVA(0) = TRIM$(FORMATNUM(TTTX#, "F12.2"))
   '
   'Total(3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(0)) * XVALO(PORIVA(0).Caption) * APLIVA% / 100), 3, 12, 2, 2)
   Total(3).TEXT = FORMATNUM$(TTTX# * XVALO(PORIVA(0).Caption) * APLIVA% / 100, "F12.2")
   '
   If APLIVA% < 1 Then
     TLABEL28 = Total(2).TEXT
     BASIVA(0) = ""
     BASIVA(1) = ""
   End If
   '
99 Call CALTOFA2
   '
End Sub

Sub ImputacionVentasPorGrupoContable()
   '
   cuentaContablePantalla% = CUEINT%(Text16)
   
   TASA1# = XVALO(PORIVA(0).Caption)
   TASA2# = XVALO(PORIVA(1).Caption)
   TASA3# = XVALO(PORIVA(2).Caption)
   SUNEFAC# = 0: SUNEFA1# = 0: SUNEFA2# = 0
   ' *** PONE EN CERO TODOS LOS INDICES DEL VECTOR
   If NetoMultipleImputacion% = 1 Then
        For i% = 1 To UBound(VentaNetaPorCuentaContable#)
            VentaNetaPorCuentaContable#(i%) = 0
        Next i%
   End If
   '
   Label18 = TRIM$(Str$(ULTREG&("!CARDEFAC")))
   SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
   SUMACAN# = 0
   '
   DESC$ = TRIM$(PORDESCU.TEXT)
   PODES# = COEFDESCU(DESC$)
   '
   For U& = 1 To ULTREG&("!CARDEFAC")
     '
     X$ = REGLEIDO$("!CARDEFAC", U&)
     CIXXI% = CVI(Mid$(X$, 5, 2))
     CAN# = CVS(Mid$(X$, 1, 4))
     SUMACAN# = SUMACAN# + CAN#
     PORDESCUIN# = CVS(Mid$(X$, 425, 4))
     If PORDESCUIN# < 0 Or PORDESCUIN# > 100 Then PORDESCUIN# = 0
     '
     If Abs(CAN#) > 0.05 Or MODODOC% > 0 Then
       TASARE# = TASA1#
       If TASACLICK >= 0.005 Then TASARE# = TASACLICK
       If CIXXI% > 0 Then
         CILOCAL% = CIXXI%
         TASARE# = PORCEIVA(CILOCAL%)
       End If
       If TASARE# <> CVS(Mid$(X$, 441, 4)) Then
         Call REPLA(X$, MKS$(TASARE#), 441, 4)
         Call GRAREG("!CARDEFAC", X$, U&)
       End If
       '
       PRELI# = CVS(Mid$(X$, 421, 4))
       PREUNI# = CVS(Mid$(X$, 429, 4))
       '
       If MODESCU% = 1 Then
         If PRELI# < 0.00005 Then PRELI# = PREUNI#
         PREUNI# = PRELI# * (100 - PORDESCUIN#) / 100
         Call REPLA(X$, MKS$(PREUNI#), 429, 4)
       End If
       '
       If CAN = 1 Then
           TOT# = XVALO(PREUNI#)
         Else
           PR1# = ROUND(PREUNI#, 4)
           TOT# = ROUND(CAN#, 1) * PR1#
       End If
       TOT# = ROUND(TOT#, 2)
       '
       If MODODOC% = 1 Then
          PREUNI# = ROUND(CVS(Mid$(X$, 429, 4)))
          TOT# = ROUND(CVD(Mid$(X$, 433, 8)), 2)
       End If
       '
       Call REPLA(X$, MKD$(TOT#), 433, 8)
       SUNEFAC# = SUNEFAC# + (TOT# * (1 - PODES# / 100))
       '
       If NetoMultipleImputacion% = 1 Then
            grupoVenta% = GRUCOVEI%(CIXXI%)
            If grupoVenta% > 0 Then CuentaContable% = CodigoCtaContablePorGrupoVenta(grupoVenta%)
            If CuentaContable% = 0 Then CuentaContable% = cuentaContablePantalla%
            VentaNetaPorCuentaContable#(CuentaContable%) = VentaNetaPorCuentaContable#(CuentaContable%) + (TOT# * (1 - PODES# / 100))
       End If
       '
     End If
   Next U&
   '
End Sub

'
Sub CALTOFA2()
   '
   Static CLIA$
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE = XVALO(CLIA$)
     RNI% = 0
     If NCLIE > 0 Then
       PPVI% = PIVACLI%(NCLIE)
       If PPVI% = 2 Then RNI% = 1
     End If
   End If
   '
   TTTX# = (XVALO(Total(3).TEXT) + XVALO(Total(4).TEXT)) * RNI%
   BASIVA(2) = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   Total(5).TEXT = CSTRING$(MKD$(XVALO(BASIVA(2)) * XVALO(PORIVA(2).Caption) / 100), 3, 12, 2, 2)
   '
   TTTX# = 0: For KKI% = 2 To 5
                 TTTX# = TTTX# + XVALO(Total(KKI%).TEXT)
              Next KKI%
              TTTX# = TTTX# + XVALO(Text15)
              TTTX# = TTTX# + XVALO(Text26)
              TTTX# = TTTX# + XVALO(Text25)
              TTTX# = TTTX# + XVALO(Text17)
              TTTX# = TTTX# + XVALO(Text24)
   Total(6).TEXT = CSTRING$(MKD$(TTTX#), 3, 12, 2, 2)
   '
End Sub

Sub LEETASIVA()
   '
   URETA& = ULTREG&("GRUPIVA")
   If URETA& < 2 Then
     Call MENSERR(24, "Falta Definir Grupos de I.V.A.")
     Call CIERRARCH("*.*")
     Call FINAL("")
   End If
   '
   For U& = 1 To 2
      X$ = REGLEIDO$("GRUPIVA", U&)
      TAIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
      PORIVA(U& - 1) = TRIM$(CSTRING$(MKS$(TAIVA(U& - 1)), 3, 6, 2, 2))
   Next U&
   PORIVA(2) = "50.00" ' responsable no inscripto
   Call CIERRARCH("GRUPIVA")
   ALIIVA = ALIVA
   '
End Sub

Sub GRACONSI()
   'jandy
   'ver ernesto LISTO 22/08
   CMO$ = TRIM$(Control$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   '
   NUFACLAR$ = CoComLar$
   NUFACORT$ = CoComCor$
   PEDIANT& = 0
   '
   NCLIE = XVALO(Text1.TEXT)
   Call APERBASDAT("STOCKS")
   '
   ' ESTO ES NUEVO
'   SQLX$ = " SELECT * FROM MOVISTO "
'   SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE & " "
'   SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
'   SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
'   SQLX$ = SQLX$ + " AND DoPriLar <> '' "
'   SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
'   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DESELEC# & " "
'   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HASELEC# & " "
   '
   
 '  Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   FIN& = ULTREG&("!CARDEFAC")
   '
   DESCX$ = PORDESCU
   PODES = COEFDESCU(DESCX$)
   COEFPRE = (100 - PODES) / 100
   '
   TTT1# = XVALO(Total(0).TEXT) * PODES / 100
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", i&)
     CANTAA = CVS(Left$(X$, 4))
     If Abs(CANTAA) > 0 Then
       CI% = CVI(Mid$(X$, 5, 2))
       DOPRIC$ = TRIM$(Mid$(X$, 387, 16))
       NORIT% = CVI(Mid$(X$, 503, 2))
       PREUNI = ROUND(CVS(Mid$(X$, 429, 4)) * COEFPRE, 4)
       DESCRILIBRE$ = Mid$(X$, 7, 64)
       
         FACIVITEM = 1: If Label13.Caption = "B" Then FACIVITEM = 1 + PORCEIVA(CI%) / 100
       'SE MODIFICA PORQUE EL CAMPO PRECONIVA SE CREO PARA QUE TENGA EL IMPORTE SIN DESCUENTO
       'PRECONIVA# = ROUND(PREUNI, DECILISTACT%)
'       PRECONIVA# = ROUND(CVS(Mid$(X$, 429, 4)), DECILISTACT%)
       PRECONIVA# = ROUND(FACIVITEM * CDbl(CVS(Mid$(X$, 429, 4))), 4)
       NPEDCAR& = XVALO(Mid$(X$, 371, 16))
       ESREPA$ = Mid$(X$, 461, 1)
       '
       If CI% > 0 Then
         '
         ORIMOVI$ = "MOVISTO"
         If ESREPA$ = "R" Then ORIMOVI$ = "MOVIREPA"
         '
         SQLX$ = " SELECT * FROM " + ORIMOVI$ + " "
         SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE & " "
         SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
         SQLX$ = SQLX$ + " OR CodiMovi =  'DV' "
         SQLX$ = SQLX$ + " OR CodiMovi =  '*R') "
         SQLX$ = SQLX$ + " AND DoPriLar <> '' "
         SQLX$ = SQLX$ + " AND ((DoSecLar is null) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
         SQLX$ = SQLX$ + " AND FechMov >= '" & DESELEC & "' "
         SQLX$ = SQLX$ + " AND FechMov <= '" & HASELEC & "' "
         SQLX$ = SQLX$ + " AND DoPriCor = '" & DOPRIC$ & "' "
         SQLX$ = SQLX$ + " AND Cod_Inte = " & CI% & " "
         SQLX$ = SQLX$ + " AND NuorItem = " & NORIT% & " "
         ' SQLX$ = SQLX$ + " AND SIGN(Cantsalid - CantIngre)= " & Sgn(CANTAA) & " "
         Dim CONSITMP As ADODB.Recordset
         Set CONSITMP = New ADODB.Recordset
25       On Error Resume Next
         CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
         End If
         On Error GoTo 0
         '
         With CONSITMP   ' Movisto
             If (.EOF And .BOF) Then
                Call MENSERR(24, "Imposible Continuar-Incoherencia")
                Exit Sub
             End If
             '
             ' ESTO REEMPLAZA LA CONDICION SQL POR PROBLEMA GUIDI 4/3/2009
26           If Sgn(XVALO(!CANTSALID) - XVALO(!CANTINGRE)) <> Sgn(CANTAA) Then
                If .EOF Then GoTo 29
                .MoveNext: GoTo 26
             End If
             '
             If IsNull(!NUPEDCLI) = False Then NPEDID& = !NUPEDCLI
             '
             !DoSecCor = NUFACORT$
             !DOSECLAR = NUFACLAR$
             !VALOUNIT = PREUNI
             !VALOUNITIVA = PRECONIVA#
             !MoneVal = MONEMI%         '  moneda de emision
             If MONEMI% <= 1 Then
                 !TIPOCAM = 1           '  SUPRIMIDO siempre cotizacion del dolar
               Else
                 !TIPOCAM = TICAMONE(MONEMI%)
             End If
             If CANTAA > 0 Then
                 !CANTSALID = CANTAA
                 !CANTINGRE = 0
               Else
                 !CANTSALID = 0
                 !CANTINGRE = Abs(CANTAA)
             End If
             CADEC% = XVALO(CNTDC$(CI%)) ' CANTIDAD DE DECIMALES CONFIGURADOS
             '
             !ValTotIt = ROUND(PREUNI * (ROUND(Abs(CANTAA), CADEC%)), 2)
             !ValTotItIva = ROUND(PRECONIVA# * (ROUND(Abs(CANTAA), CADEC%)), 2)

             If NPEDID& < 1 Then
               !NUPEDCLI = NPEDCAR&
               NPEDID& = NPEDCAR&
             End If
             NUMSERIE$ = SAFETEXT$(!idenlote)
             .Update
             GoTo 29
         End With
       End If
       '
       'ACA VIENE SI ES CODIGO 0
       FECHA7% = FECHANUM(Text9)
       VUNI7# = PREUNI
       Call MOVISTOK(FECHA7%, CI%, "", CMO$, NUFACLAR$, NUFACORT$, NUFACLAR$, NUFACORT$, rasocli$(NCLIE), NORIT%, VUNI7#, MONEMI%, NCLIE, DEPOEXPE%, (-1) * CANTAA, NPEDCAR&, 0, DESCRILIBRE$, "", "", PRECONIVA#)
       If DESCOMPO%(CI%) > 0 Then
          Call MOVISTOK(FECHA7%, CI%, "", CMO$, NUFACLAR$, NUFACORT$, NUFACLAR$, NUFACORT$, rasocli$(NCLIE), NORIT%, VUNI7#, MONEMI%, NCLIE, DEPOEXPE%, CANTAA, NPEDCAR&)
          Call AREXPLO(CI%, 1, 1)
          For KKJ% = 1 To CAIT%
             CIMO% = CIJ%(KKJ%)
             CAMO = USOI(KKJ%) * CANTAA
             Call MOVISTOK(FECHA7%, CIMO%, "", CMO$, NUFACLAR$, NUFACORT$, NUFACLAR$, NUFACORT$, rasocli$(NCLIE), NORIT%, VUNI7#, MONEMI%, NCLIE, DEPOEXPE%, (-1) * CAMO, NPEDCAR&)
             NORIT% = NORIT% + 1
          Next KKJ%
       End If
       '
     End If
     'SI TIENE NUMERO DE SERIE ESCRIBE EN EL TAMUMSE EL NUMERO DE FACTURA
29   If TRIM$(NUMSERIE$) <> "" And TRIM$(NUMSERIE$) <> "S/N" Then
        CONDI$ = "NUMEROSERIE = '" & NUMSERIE$ & "'"
        Call ACTUREGISTRO("TANUMSE", "NUMFACTURA", CONDI$, NUFACLAR$, REGAF&, "NOMESS")
     End If
     '
   Next i&
   '
   CONSITMP.Close
   Set CONSITMP = Nothing
   '
   Exit Sub
   '
End Sub

Sub GRAREMIFA()
   '
   
   CMO$ = TRIM$(Control$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   DEPOSALI% = DEPOEXPE%
   '
   PRONUMTRANSF& = 0
   NUFACLAR$ = CoComLar$
   NUFACORT$ = CoComCor$
   COPIASFAC% = XVALO(FINFACTUR07.Text2(0))
   If COPIASFAC% < 1 Then
      NUFACLAR$ = ""
      NUFACORT$ = ""
   End If
   '
   UTILIZA_ITEM_S_CARGO$ = TRIM$(UCase$(Control("ADMIPED", "ITEMSCGO") = "SI"))
   
   USASALDOLOTEDEP% = 0
   If TRIM$(Control("CONSLOTE", "SALOTEDE")) = "SI" Then USASALDOLOTEDEP% = 1
   
   PREFIREM$ = Control$("", "PREFIREM")
   If PREFIREM$ = "" Then PREFIREM$ = "0001"
   'ESTO SE AGREGA PORQUE LOS REMITOS DE IMPRESORA FISCAL TIENEN EL MISMO
   'PUNTO DE VENTA QUE LA FACTURA
   If Control$("FACTURA", "SALIDA") = "IFISCAL" Then PREFIREM$ = Label12
   '
   'ESTO SE AGREGA PORQUE LOS REMITOS DE LAS
   'FACTURAS ELECTRONICAS TIENEN EL PUNTO DE VENTA DE LA FACTURA
   ' If FACELEC% = 1 And VERIFSALIDA("NOMESS") < 3 Then PREFIREM$ = Label12
   
   If FACELEC% > 0 And FACELEC% < 3 Then PREFIREM$ = Label12
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "TANIT") > 0 Or InStr(EPAC$, "GABAL") > 0 Or InStr(EPAC$, "FERVI") > 0 Or InStr(EPAC$, "AHP") > 0 Or InStr(EPAC$, "J Y M") > 0 Or InStr(EPAC$, "CEYEQ") > 0 Or InStr(EPAC$, "EYM") > 0 Or InStr(EPAC$, "DAVTECNICA") > 0 Then PREFIREM$ = "0001"
   If InStr(EPAC$, "ASEMBLI") > 0 Then PREFIREM$ = Control$("", "PREFIREM")
   '
   REMICOR$ = "RT." + TRIM$(Text10)
   REMILAR$ = Text10: While Len(REMILAR$) < 8: REMILAR$ = "0" + REMILAR$: Wend
   REMILAR$ = CMO$ + "-X-" + PREFIREM$ + "-" + REMILAR$
   If COPIASREM% < 1 Then
     '\\\OT-8-965 - SE MODIFICA PORQUE LA ANULACION DE UNA NC REQUIERE ESTE DATO
     REMICOR$ = NUFACORT$
     REMILAR$ = NUFACLAR$
     '\\\OT-8-965-FIN
   End If
   '
   NCLI = XVALO(Text1.TEXT)
   FF% = FECHANUM(Text9)
   '
   DESCX$ = PORDESCU
   PODES = COEFDESCU(DESCX$)
   COEFPRE = (100 - PODES) / 100
   '
   FIN& = ULTREG&("!CARDEFAC")
   '\\\OT-07-0023-RG-24/01/07///
   DEPOSALI% = XVALO(FINFACTUR07.Text7)
   '\\\OT-07-0023-RG-FIN///
   PEDIANT& = 0
   '
   ANOTA$ = ""
   TTXYZ$ = Space$(76)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, "Factura: " & NUFACORT$ & " - Remito: " & REMICOR$, 11, 56) 'ASIGNO EL NUMERO DE FACUTURA Y REMITO AL PROSPECT
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
   '
   '*MODIFICACION PARA QUE EN NOTA DE DEB. CRED. SI SE MODIFICA LA BASE DE IVA SE CALCULE EN CONSECUENCIA PARA EL MOVISTO
   '**********************************************************************************************************************************
   CALCULAIVA_TL% = 0 ' VARIABLE BANDERA DE CALCULO DE IVA N.CRED. DEB TEXTO LIBRE
   If XVALO(TLABEL28) > 0.005 And Label6 = "(T.Libre)" And (Label7 = "N.Crédito:" Or Label7 = "N.Débito:") Then
        If XVALO(BASIVA(0)) + XVALO(BASIVA(1)) < 0.005 Then
            CALCULAIVA_TL% = -1 ' NO SUMA IVA POR QUE ES
        Else
            TOTNETO# = XVALO(Total(2))
            TOTAIMPOIVA# = XVALO(Total(3)) + XVALO(Total(4))  'SUMO LOS IMPORTES DE IVA
            TOTAX# = TOTNETO# + TOTAIMPOIVA#                  'TOTAL CON IVA
            PJEIVA_PROM = ((TOTAX# / TOTNETO#) - 1) * 100 ' CALCULO EL PORCENTAJE DE IVA PROMEDIO ENTRE CAJA IVA 10.5 Y 21.00
            CALCULAIVA_TL% = 1
        End If
   End If
   '*/*/*/*/**//*/*/*/*/*
   DEPOFACTU% = XVALO(Control$("", "DEPOEXPE"))
   PANOLFACTU% = XVALO(Control$("", "DPREPPAÑ"))
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", i&)
     CANTAA = CVS(Left$(X$, 4))
     If Abs(CANTAA) > 0 Then
       CI% = CVI(Mid$(X$, 5, 2))
       If CI% >= 0 Then                 ' para incluir remito texto libre
''            'MODIFICACION N.CREDITO PARA QUE EL REMITO QUEDE COMO EN CONSIGNACION
''            X$ = REGLEIDO$("!CARDEFAC", I&)
''            LETX$ = Mid$(X$, 461, 1)
''            If LETX$ = "N" And TIFADOC% = 6 Then
''               NUMORDIT% = CVI(Mid$(X$, 117, 2))
''               CONDI$ = "NUORITEM = " & NUMORDIT% & " AND  COD_INTE = " & CI% & " AND DOSECLAR = '" & TRIM$(Label31) & "'"
''               Call ACTUREGISTRO("MOVISTO", "DOSECLAR", CONDI$, "P/F", REGAF&, "NOMESS")
''               'FIN MODIFICACION N.CREDITO PARA QUE EL REMITO QUEDE COMO EN CONSIGNACION
''            End If
         NORIT% = i& - 1
         ' PREUNI# = CDbl(CVS(Mid$(X$, 429, 4)) * COEFPRE)
         PREUNI# = ROUND(CVS(Mid$(X$, 429, 4)) * COEFPRE, 4)
         FACIVITEM = 1: If Label13.Caption = "B" Then FACIVITEM = 1 + PORCEIVA(CI%) / 100
         '
         '*MODIFICACION PARA QUE EN NOTA DE DEB. CRED. SI SE MODIFICA LA BASE DE IVA SE CALCULE EN CONSECUENCIA PARA EL MOVISTO
         '**********************************************************************************************************************************
         If CALCULAIVA_TL% < 0 Then 'NO TIENE IVA GRABADO 'NOTA DE CREDITO O DEBITO POR TEXTO LIBRE
            FACIVITEM = 1: If Label13.Caption = "B" Then FACIVITEM = 1
         ElseIf CALCULAIVA_TL% = 1 Then
            FACIVITEM = 1: If Label13.Caption = "B" Then FACIVITEM = 1 + PJEIVA_PROM / 100
         End If
         '**********************************************************************************************************************************
         
         ' PRECONIVA# = FACIVITEM * CDbl(CVS(Mid$(X$, 429, 4)))
         PRECONIVA# = ROUND(FACIVITEM * CDbl(CVS(Mid$(X$, 429, 4))), 4)
         '\\\OT-05-0562-WAR-02/08/05///
         SIG% = (-1)
         If TIFADOC% = 6 Then SIG% = 1 'NC
         '\\\OT-05-0562-WAR-02/08/05///
         NUPEDIORI& = XVALO(Mid$(X$, 371, 16))
         NUORITPE% = 0: If NUPEDIORI& > 0 Then NUORITPE% = CVI(Mid$(X$, 503, 2))
         '
         If Mid$(X$, 461, 1) = "R" Then CMO$ = "*R" ' REPARACION
         TEXLIB$ = Mid$(X$, 7, 64)
         '
         NUSERRR$ = TRIM$(Mid$(X$, 139, 64))
         If TRIM$(NUSERRR$) = "S/N" Then NUSERRR$ = ""
TXTRAZA$ = TXTRAZA$ & "PUNTO 46.1 = " & Now & vbCrLf         '
         ' APLICA ASIGNACION DE LOTE
         If ESTRAZABLE%(CI%) = 1 Then
           If InStr(UCase$(EMPREAC$), "AHP") > 0 Then
              GoSub CONSUMO_AHP
           'SE COMENTA SEGUN MODIFICACION PRESUPUESTO 13/04/18
           'ElseIf EXISTE_TABLA%("RESERPED") And InStr(UCase$(EMPREAC$), "PREMOTEC") > 0 Then
              'GoSub CONSUMO_PREMOTEC
           Else
             ' BUSCA EL LOTE DEL ARTICULO
             GoSub MUEVEPORLOTE
           End If
             Call LIBARCH("DEUDOR2")
         Else
             If Label13.Caption = "E" And InStr(1, EMPREAC$, "GABAL") > 0 And CI% > 0 Then
                TEXLIB$ = SAFETEXT$(VALOBADA("MASTER", "DESCRIDOS", "CODINT=" & CI%))
             End If
             If PANOLFACTU% > 0 Then 'CONFIGURACION EN EL SETUP DE FACTURA
                If STODEPOS(CI%, DEPOFACTU%) <= 0 And NUSERRR$ = "" Then 'SI VIENE ACA P Q NO ES TRAZABLE Y SI ENTRA PQ NO TIENE STOCK EN 3
                   If PRONUMTRANSF& < 1 Then 'PARA QUE SOLO ENTRE UNA VEZ
                      CMOV$ = TRIM$(Control$("", "CODTRAST"))
                      If CMOV$ = "" Then CMOV$ = "TR"
                      PRONUMTRANSF& = ProNroComprobante("TR")
                      DOPRI$ = "TR." + TRIM$(Str$(PRONUMTRANSF&))
                   End If
                   'REALIZADO PARA MEDITEA (CASO NO TRAZABLE 1) SI NO HAY STOCK EN DEP FACTURACION TRANSFIERE DEL 44 AL 3 CONFIGURACION EN SETUP DE FACTURA
                   If TIFADOC% <> 6 Then 'CON EXEPCION DE LA NOTA DE CREDITO
                     Call MOVISTOK(FF%, CI%, NUSERRR$, CMOV$, DOPRI$, DOPRI$, NUFACLAR$, NUFACORT$, "Sal.Aut.P/Fact.", NORIT%, PREUNI#, MONEMI%, NCLI, PANOLFACTU%, -1 * CANTAA, NUPEDIORI&, , TEXLIB$, , , PRECONIVA#)       '(-1) * CANTAA )
                     Call MOVISTOK(FF%, CI%, NUSERRR$, CMOV$, DOPRI$, DOPRI$, NUFACLAR$, NUFACORT$, "Ing.Aut.P/Fact.", NORIT%, PREUNI#, MONEMI%, NCLI, DEPOFACTU%, 1 * CANTAA, NUPEDIORI&, , TEXLIB$, , , PRECONIVA#)   '(-1) * CANTAA )
                   End If
                End If
             End If
             '
             If MODODOC% = 0 Then 'SI ES MATERIAL CODIFICADO
               If UTILIZA_ITEM_S_CARGO$ = "SI" Then
                If PREUNI# <= 0.01 Then
                   TEXLIB$ = TEXLIB$ & " (S/C)"
                End If
               End If
             End If
             Call MOVISTOK(FF%, CI%, NUSERRR$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, rasocli$(NCLI), NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, SIG% * CANTAA, NUPEDIORI&, NUORITPE%, RTrim$(TEXLIB$) + DATOSIMPO$(CI%, NUSERRR$), , , PRECONIVA#)     '(-1) * CANTAA )
             Call LIBARCH("DEUDOR2")
             If DESCOMPO%(CI%) > 0 Then
                  Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, rasocli$(NCLI), NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, (-1) * SIG% * CANTAA, NUPEDIORI&, , TEXLIB$, , , PRECONIVA#)   '(-1) * CANTAA )
                  Call AREXPLO(CI%, 1, 1)
                  For KKJ% = 1 To CAIT%
                     CIMO% = CIJ%(KKJ%)
                     CAMO = USOI(KKJ%) * CANTAA
                     Call MOVISTOK(FF%, CIMO%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, rasocli$(NCLI), NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, SIG% * CAMO, NUPEDIORI&, , TEXLIB$, , , PRECONIVA#) '(-1) * CAMO)
                     NORIT% = NORIT% + 1
                  Next KKJ%
             End If
         End If
TXTRAZA$ = TXTRAZA$ & "PUNTO 46.2 = " & Now & vbCrLf         '
10       If NUSERRR$ <> "" And UCase$(NUSERRR$) <> "S/N" And TIFADOC% <> 6 Then  'GRABA NUERO DE FACTURA Y DE REMITO EN TABLA TANUMSE
            NUREM1$ = REMILAR$
            NUFAC1$ = NUFACLAR$
            NUCLI = NCLI
            FECHENT = CVDAT(FF%)
            FE1% = DIASPOST(FF%, 730) 'FECHA DE GARANTIA
            FECHENT1 = CVDAT(FE1%)
            '
'            Call ACTUREGISTRO("TANUMSE", "NumRemito", "NUMEROSERIE = '" & NUSERRR$ & "'", NUREM1$, NUAFEC&, "NOMESS")
'            Call ACTUREGISTRO("TANUMSE", "NumFACTURA", "NUMEROSERIE = '" & NUSERRR$ & "'", NUFAC1$, NUAFEC&, "NOMESS")
'            Call ACTUREGISTRO("TANUMSE", "NumeCli", "NUMEROSERIE = '" & NUSERRR$ & "'", NUCLI, NUAFEC&, "NOMESS")
'            Call ACTUREGISTRO("TANUMSE", "FechEntrega", "NUMEROSERIE = '" & NUSERRR$ & "'", FechEnt, NUAFEC&, "NOMESS")
'            Call ACTUREGISTRO("TANUMSE", "FeVenGar", "NUMEROSERIE = '" & NUSERRR$ & "'", FECHENT1, NUAFEC&, "NOMESS")
            '
            ''''''*/**********
            SQLX$ = "SELECT NumRemito, NumFACTURA, NumeCli, FechEntrega, FeVenGar FROM TANUMSE "
            SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NUSERRR$ & "'"
            
            Dim RSTNSE As ADODB.Recordset
            Set RSTNSE = New ADODB.Recordset
25          On Error Resume Next
            RSTNSE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                On Error GoTo 0
                Call CapturaErrorOpen
                GoTo 25
            End If
            On Error GoTo 0
            With RSTNSE
              Do While Not .EOF
               !NumRemito = NUREM1$
               !NumFACTURA = NUFAC1$
               !NUMECLI = NUCLI
               !FECHENTREGA = FECHENT
               !FEVENGAR = FECHENT1
               .Update
               .MoveNext
              Loop
            End With
            RSTNSE.Close
            Set RSTNSE = Nothing
TXTRAZA$ = TXTRAZA$ & "PUNTO 46.3 = " & Now & vbCrLf
            '////********
            'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
            CANCOPIASFAC% = XVALO(FINFACTUR07.Text2(0)) ' cantidad de copias de factura
            CANCOPIASREM% = XVALO(FINFACTUR07.Text2(1)) ' cantidad de copias de remitos
            CUERPO$ = "Venta de Equipos "
            If CANCOPIASFAC% > 0 Then
              CUERPO$ = CUERPO$ & "S/Factura: " & NUFAC1$ & vbCrLf
            End If
            If CANCOPIASREM% > 0 Then
               CUERPO$ = CUERPO$ + "Remito: " & NUREM1$
               '
               'ESCRIBO CON EL NUMERO DE REMITO QUE ESTE SALIO SIN CARGO Y EL MOTIVO
               NUMOT% = NUMOTIVO%(NUPEDIORI&)
               If NUMOT% > 0 Then
                 DESCMOTI$ = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
                 CUERPO$ = CUERPO$ + " - S/Cargo - Motivo: " & NUMOT%
                 CUERPO$ = CUERPO$ + " (" & TRIM$(DESCMOTI$) & ")"
               End If
            End If
TXTRAZA$ = TXTRAZA$ & "PUNTO 46.4 = " & Now & vbCrLf
            Call ANOTANUSE(NUSERRR$, CUERPO$)
    '
         End If
TXTRAZA$ = TXTRAZA$ & "PUNTO 46.5 = " & Now & vbCrLf
         
         ' NOTA DE CREDITO
         If NUSERRR$ <> "" And UCase$(NUSERRR$) <> "S/N" And TIFADOC% = 6 Then
            DEPX% = DEPOSALI% ' deposito configurado en FINFACTUR07
'            CONDI$ = "NUMEROSERIE = '" & NUSERRR$ & "'"
'            Call ACTUREGISTRO("Tanumse", "NumRemito", CONDI$, "", REG&, "NOMESS")
'            Call ACTUREGISTRO("Tanumse", "NumFactura", CONDI$, "", REG&, "NOMESS")
'            Call ACTUREGISTRO("Tanumse", "FechEntrega", CONDI$, "", REG&, "NOMESS")
'            Call ACTUREGISTRO("Tanumse", "NumeCli", CONDI$, 0, REG&, "NOMESS")
'            Call ACTUREGISTRO("Tanumse", "DepoNumse", CONDI$, DEPX%, REG&, "NOMESS")
'            Call ACTUREGISTRO("Tanumse", "FeVenGar", CONDI$, "", REG&, "NOMESS")
'            Call ACTUREGISTRO("PEDDETA", "NUMEROSERIE", CONDI$, "", REG&, "NOMESS")
            
            SQLX$ = "SELECT NumRemito, NumFACTURA, NumeCli, FechEntrega, FeVenGar,DepoNumse FROM TANUMSE "
            SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NUSERRR$ & "'"
            
            Dim RSTNSE1 As ADODB.Recordset
            Set RSTNSE1 = New ADODB.Recordset
35          On Error Resume Next
            RSTNSE1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                On Error GoTo 0
                Call CapturaErrorOpen
                GoTo 35
            End If
            On Error GoTo 0
            With RSTNSE1
              Do While Not .EOF
               !NumRemito = ""
               !NumFACTURA = ""
               !NUMECLI = 0
               !FECHENTREGA = CVDAT(0)
               !FEVENGAR = CVDAT(0)
               !DEPONUMSE = DEPX%
               .Update
               .MoveNext
              Loop
            End With
            RSTNSE1.Close
            Set RSTNSE1 = Nothing
            '
            CONDI$ = "NUMEROSERIE = '" & NUSERRR$ & "'"
            Call ACTUREGISTRO("PEDDETA", "NUMEROSERIE", CONDI$, "", REG&, "NOMESS")
            '
            'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
            CUERPO$ = "Reingreso de Equipo Según: " & NUFACLAR$
            Call ANOTANUSE(NUSERRR$, CUERPO$)
            '
         End If
         'EN EL CASO QUE EL PEDIDO SEA CON MOTIVO LO MARCA COMO SIN CARGO
         'CON EXCEPCION SI TIENE INGRESADA CANTIDAD DE COPIAS DE FACTURAS > 0
         If NUMOTIVO%(NUPEDIORI&) > 0 And COPIASFAC% < 1 Then
            'CONDI$ = "DOPRILAR = '" & REMILAR$ & "'"
'            Call ACTUREGISTRO("MOVISTO", "DOSECCOR", CONDI$, "S/Cargo", REGAF&, "NOMESS")
'            Call ACTUREGISTRO("MOVISTO", "DoSecLar", CONDI$, "S/Cargo", REGAF&, "NOMESS")
            
            SQLX$ = "SELECT DOSECCOR, DoSecLar FROM MOVISTO "
            SQLX$ = SQLX$ + " WHERE DOPRILAR = '" & REMILAR$ & "'"
            
            Dim RSMVTO As ADODB.Recordset
            Set RSMVTO = New ADODB.Recordset
45          On Error Resume Next
            RSMVTO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                On Error GoTo 0
                Call CapturaErrorOpen
                GoTo 45
            End If
            On Error GoTo 0
            With RSMVTO
              Do While Not .EOF
               !DoSecCor = "S/Cargo"
               !DOSECLAR = "S/Cargo"
               .Update
               .MoveNext
              Loop
            End With
            RSMVTO.Close
            Set RSMVTO = Nothing
            
            Call COMUNI("Movimiento " + REMICOR$ + " Marcado 'Sin Cargo'.")
         End If
         '
         TTXYZ$ = Space$(76) ' AGREGO DETALLE AL PROSPECT
         Call REPLA(TTXYZ$, "- " + CODEXT$(CI%), 11, 18)
         Call REPLA(TTXYZ$, DESCRIT0$(CI%), 29, 38)
         Call REPLA(TTXYZ$, CSTRING$(MKS$(CANTAA), 3, 10, 1, 2), 57, 10)
         ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)

         ' PARA DOSIVAC ESQUEMA DE PACKING-LIST
         EPAC$ = TRIM$(UCase$(EMPREAC$))
         If InStr(EPAC$, "DOSIVA") > 0 Then
           If Left$(REMICOR$, 2) = "RT" Then
             DOPRIQ$ = "PL." + TRIM$(Mid$(X$, 371, 16))
             If TRIM$(DOPRIQ$) = "PL." Then   ' REMITO DIRECTO SIN O/VENTA
               NRXLX& = 900000 + Val(Text10)
               DOPRIQ$ = DOPRIQ$ + TRIM$(Str$(NRXLX&))
             End If
             REFEQ$ = "Packing-List " + rasocli$(NCLI)
             If CMO$ = "*R" Then
                  CMOVQ$ = "*D"
               Else
                  CMOVQ$ = "RP"
             End If
             Call MOVISTOK(FF%, CI%, LOTE$, CMOVQ$, DOPRIQ$, DOPRIQ$, NUFACLAR$, NUFACORT$, REFEQ$, NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, (-1) * SIG% * CANTAA, NUPEDIORI&)  '(-1) * CANTAA)
             If DESCOMPO(CI%) > 0 Then
                  Call MOVISTOK(FF%, CI%, LOTE$, CMOVQ$, DOPRIQ$, DOPRIQ$, NUFACLAR$, NUFACORT$, REFEQ$, NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, SIG% * CANTAA, NUPEDIORI&) '(-1) * CANTAA)
                  Call AREXPLO(CI%, 1, 1)
                  For KKJ% = 1 To CAIT%
                     CIMO% = CIJ%(KKJ%)
                     CAMO = USOI(KKJ%) * CANTAA
                     Call MOVISTOK(FF%, CIMO%, LOTE$, CMOVQ$, DOPRIQ$, DOPRIQ$, NUFACLAR$, NUFACORT$, REFEQ$, NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, (-1) * SIG% * CAMO, NUPEDIORI&) '(-1) * CAMO)
                     NORIT% = NORIT% + 1
                 Next KKJ%
             End If
           End If
         End If
         '
         
         If CANTAA > 0 Then ' PARA EVITAR CON NOTAS DE CREDITO
           CANAX# = CANTAA
           NUORITPE% = CVI(Mid$(X$, 503, 2))
           If NUPEDIORI& > 0 Then
TXTRAZA$ = TXTRAZA$ & "PUNTO 46.6 = " & Now & vbCrLf
           
             Call BLANPEDIPE(NCLI, FF%, REMILAR$, CI%, CANAX#, NUPEDIORI&, NUORITPE%)
TXTRAZA$ = TXTRAZA$ & "PUNTO 46.7 = " & Now & vbCrLf
             GoSub ANOTAPEDI
TXTRAZA$ = TXTRAZA$ & "PUNTO 46.8 = " & Now & vbCrLf
           End If
         End If
         '
       End If
     End If
   Next i&
   '
   If FACTLISTAEMPAQUE& > 0 Then
     CONDI$ = "NUMLTAEMP = " & FACTLISTAEMPAQUE& & " AND CODI_CLI = " & NCLI
     Call ACTUREGISTRO("ENCEMPAQ", "NUMFACTURA", CONDI$, NUFACLAR$, REGAF&, "NOMESS")
   End If
   '
   GASTEXP# = XVALO(FINFACTUR07.Text2(15))
   If Label13.Caption = "E" And GASTEXP# > 0 Then
       'SE AGREGA A PARAMETRO DESCRILA LA REFERENCIA DEL GASTO DE EXPORTACION
       Call MOVISTOK(FF%, 0, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, TRIM$(FINFACTUR07.Text10), NORIT% + 1, GASTEXP#, MONEMI%, NCLI, DEPOSALI%, -1, NUPEDIORI&, 0, TRIM$(FINFACTUR07.Text10), , , GASTEXP#)
       GRABGASEXP% = 1
   End If
   '
   If PEDIANT& > 0 Then
      Call AGREGA_ANOTAPED(PEDIANT&, Space$(10) + String$(56, "-"))
      Call AGREGA_ANOTAPED(PEDIANT&, String$(88, "="))
   End If
   '
   Call ABRECLIEN
   '
   NPRO$ = VALOBADA("BSEGCOM", "CODI_CLI", "CODI_CLI = '" & TRIM$(Label26) & "'") 'BUSCO EL NUMERO DE CODIGO DE CLIENTE
   If NPRO$ <> "" Then
     Call AGREGA_ANOTAPROSPECT(NPRO$, Mid$(ANOTA$, 1, Len(ANOTA$) - 2)) ' AGREGA AL ANOTADOR DEL PROSPECT
     Call AGREGA_ANOTAPROSPECT(NPRO$, String$(88, "="))
   End If
   '
   'GRABO SI ES REPARACION
TXTRAZA$ = TXTRAZA$ & "PUNTO 46.9 = " & Now & vbCrLf
   If XVALO(Label31) > 0 Then
     NUREP& = XVALO(Label31)
     CONDI$ = "NUMREPA = " & NUREP&
     IMPOR# = XVALO(Total(6)): IMPOR# = IMPOR# * TICAMONE(MONEMI%)
     Call ACTUREGISTRO("REPARA", "FACTURA", CONDI$, NUFACLAR$, REGAF&, "NOMESS")
     Call ACTUREGISTRO("REPARA", "IMPORTE", CONDI$, FORMATNUM$(IMPOR#, "F16.5"), REGAF&, "NOMESS")
     Label31 = ""
   End If
TXTRAZA$ = TXTRAZA$ & "PUNTO 46.10 = " & Now & vbCrLf
   '
Call GRABAR_ASIENTO_IMPORTADO(REMILAR$, NUFACLAR$)

Exit Sub
   '
ANOTAPEDI:                ' ANOTADOR DE PEDIDOS
   '
   NPEDIDX& = NUPEDIORI&
   If NPEDIDX& <> PEDIANT& Then
      If PEDIANT& <> 0 Then
          Call AGREGA_ANOTAPED(PEDIANT&, Space$(10) + String$(56, "-"))
          Call AGREGA_ANOTAPED(PEDIANT&, String$(88, "="))
      End If
      PEDIANT& = NPEDIDX&
      COPIASFAC% = Val(FINFACTUR07.Text2(0))
      COPIASREM% = Val(FINFACTUR07.Text2(1))
      TTXX$ = ""
      If COPIASREM% > 0 Then TTXX$ = "Remito " + TRIM$(Text10)
      'If COPIASFAC% > 0 Then TTXX$ = TTXX$ + " - Factura " + TRIM$(Label12) + "-" + TRIM$(Text11)
      TTXX$ = TTXX$ + ":"
      '
      TTXYZ$ = Space$(76)
      HOII% = HOY(HOS$)
      Call REPLA(TTXYZ$, HOS$, 1, 8)
      Call REPLA(TTXYZ$, TTXX$, 11, 56)
      Call REPLA(TTXYZ$, USERTER$, 69, 24)
      Call AGREGA_ANOTAPED(NPEDIDX&, TTXYZ$)
      Call REPLA(TTXYZ$, Space$(88), 1, 88)
      Call AGREGA_ANOTAPED(NPEDIDX&, TTXYZ$)
      '
   End If
   '
   TTXX$ = Space$(76)
   Call REPLA(TTXX$, "- " + CODEXT$(CI%), 11, 18)
   Call REPLA(TTXX$, DESCRIT0$(CI%), 29, 38)
   Call REPLA(TTXX$, FORMATNUM$(CANAX#, "F10.1"), 57, 10)
   Call AGREGA_ANOTAPED(NPEDIDX&, TTXX$)
   '
Return

'***********************************************************************************************
MUEVEPORLOTE:
        AcumulaImporteAsiento = 0
        For II& = 0 To ASIGLOT07.List1.ListCount
            XX$ = ASIGLOT07.List1.List(II&)
            If XVALO(Mid$(XX$, 80, 5)) = i& Then
               If CODINT%(TRIM$(Mid$(XX$, 1, 16))) = CI% Then
                  CANTLOT$ = Mid$(XX$, 85)
                  CANTI = XVALO(Mid$(XX$, 64, 10)) 'CANTIDAD A ENTREGAR
                  GoTo 5
               End If
            End If
        Next II&
    '
    
    ' MUEVE EL STOCK POR CADA LOTE
5   NN% = InStr(CANTLOT$, ";")
    If NN% > 0 Then
        LOTE$ = Mid$(CANTLOT$, 1, NN% - 1)
        If TRIM$(LOTE$) <> "" Then
            'CANTAA = XVALO(VALOBADA("LOTINGRE", "CANALTA", "IDENLOTE='" & LOTE$ & "'", "NOMESS"))
            'ENTREGA LA MENOR ENTRE LA CANTIDAD A ENTREGAR Y EL SALDO DEL LOTE
            'CANLOTE = XVALO(VALOBADA("LOTINGRE", "CANALTA", "IDENLOTE='" & LOTE$ & "'", "NOMESS"))
            'SE DEBERIA REEMPLAZAR LA CONSULTA AL LOTINGRE POR EL SALDOLOTE Y ASIGNAR ESTE VALOR A  CANLOTE
            'ver si es nota de credito en ese caso si deberia validar que no se ingrese mas del ingreso total del lote
            
            If TIFADOC% = 6 Then 'si es nota de credito
               CANLOTE = XVALO(VALOBADA("LOTINGRE", "CANALTA", "IDENLOTE='" & LOTE$ & "'", "NOMESS"))
            Else
               CANLOTE = SaldoLote(CI%, LOTE$)
            End If
            
            CANTAA = CANTI: If CANTI > CANLOTE Then CANTAA = CANLOTE
            'SE AGREGA EL PRECONIVA PARA QUE IMPRIMA EL PRECIO
            '
            If PRONUMTRANSF& < 1 And PANOLFACTU% > 0 Then  'PARA QUE SOLO ENTRE UNA VEZ
                CMOV$ = TRIM$(Control$("", "CODTRAST"))
                If CMOV$ = "" Then CMOV$ = "TR"
                PRONUMTRANSF& = ProNroComprobante(CMOV$)
                DOPRI$ = "TR." + TRIM$(Str$(PRONUMTRANSF&))
            End If
            'TRANSFERIR DEL 44 AL 3 MEDITEA
            If PANOLFACTU% > 0 Then 'CONFIGURACION EN EL SETUP DE FACTURA
                Call MOVISTOK(FF%, CI%, LOTE$, CMOV$, DOPRI$, DOPRI$, NUFACLAR$, NUFACORT$, "Sal.Aut.P/Fact.", NORIT%, PREUNI#, MONEMI%, NCLI, PANOLFACTU%, -1 * CANTAA, NUPEDIORI&, , TEXLIB$, , , PRECONIVA#)    '(-1) * CANTAA )
                Call MOVISTOK(FF%, CI%, LOTE$, CMOV$, DOPRI$, DOPRI$, NUFACLAR$, NUFACORT$, "Ing.Aut.P/Fact.", NORIT%, PREUNI#, MONEMI%, NCLI, DEPOFACTU%, 1 * CANTAA, NUPEDIORI&, , TEXLIB$, , , PRECONIVA#)   '(-1) * CANTAA )
            End If
            Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, rasocli$(NCLI), NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, SIG% * CANTAA, NUPEDIORI&, , RTrim$(TEXLIB$) + DATOSIMPO$(CI%, LOTE$), , , PRECONIVA#)       '(-1) * CANTAA )
            CANTI = CANTI - CANTAA 'RECALCULA CANTIDAD A ENTREGAR
            If CANTI < 0.05 Then
               GoTo 10
            End If
'                Call ASIGLOT07.List1.Clear comentado por que limpiaba el list y dejaba en cantidad 0 los siguientes
'                Call ASIGLOT07.List2.Clear
'                GoTo 10
'            End If
            CANTLOT$ = Mid$(CANTLOT$, NN% + 1)
            GoTo 5
        End If
    End If
'    Call ASIGLOT07.List1.Clear
'    Call ASIGLOT07.List2.Clear
    GoTo 10
    
    '
Return
 'SE COMENTA SEGUN MODIFICACION PRESUPUESTO 13/04/18
'CONSUMO_PREMOTEC:
'        CANTISALDO = CANTAA
'        II& = 0
'        Do While II& < ASIGLOPRM.MSF.Rows
'          With ASIGLOPRM.MSF
'            II& = II& + 1
'            CI1% = XVALO(.TextMatrix(II&, 1))
'            If CI1% = CI% Then
'              LOTE$ = .TextMatrix(II&, 6)
'              CANTIRESERVA = XVALO(.TextMatrix(II&, 7))
'              NUMPED& = XVALO(.TextMatrix(II&, 2))
'              If CANTIRESERVA > CANTAA Then CANTIRESERVA = CANTAA
'              Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, rasocli$(NCLI), NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, SIG% * CANTIRESERVA, NUPEDIORI&, , TEXLIB$, , , PRECONIVA#)   '(-1) * CANTAA )
'              Call CONSUMERESERVAS(CI%, NUMPED&, LOTE$, NCLI, CANTIRESERVA)
'              If CANTIRESERVA = XVALO(.TextMatrix(II&, 7)) Then
'                   Call cmdEliminaItem(ASIGLOPRM.MSF, II&)
'                    II& = II& - 1
'              Else
'                   .TextMatrix(II&, 7) = CANTISALDO
'              End If
'              CANTISALDO = CANTISALDO - CANTIRESERVA
'              If CANTISALDO <= 0.005 Then Exit Do
'            End If
'          End With
'        Loop
''    GoTo 10
'
'    '
'Return

CONSUMO_AHP:
    '
    'LEO EL LOTE
    LOTE$ = ""
    NPROVX& = CVS(Mid(X$, 83, 16))
    
    If TIFADOC% = 6 And AnulaFacturaComprobante$ <> "" Then LOTE$ = Mid(X$, 139, 64) 'si es nota de credito
    Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, rasocli$(NCLI), NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, SIG% * CANTAA, NUPEDIORI&, NUORITPE%, TEXLIB$, , , PRECONIVA#) '
    Call RecuperarLote(CI%, REMILAR$, NPROVX&)
'    If USASALDOLOTEDEP% < 1 Then
'    'ESTE CONSUMO TIENE EL CRITERIO QUE SI TIENE LOTES Y LA CANTIDAD NECESARIA GENERA UNA SALIDA CON LOS LOTES
'        'EN CASO QUE LOS MISMOS TENGAN ALGUN PARCIAL GENERA 2 MOVIMIENTOS UNA SALIDA CON LOTE X EL SALDO DEL LOTE
'        'Y OTRA POR LA CANTIDAD RESTANTE PERO SIN LOTE.
'        'SI NO POSEE LOTE DIRECTAMENTE GENERA UNA SALIDA POR EL TOTAL DE LA CANTIDAD.
'        '
'        'LEO EL LOTE
'        If MODODOC% < 1 And TIFADOC% > 2 Then  ' caso nota de credito material codificado
'           LOTE$ = TRIM$(Mid(X$, 139, 64)) ' TRAE EL QUE CARGA EN EL LA RUTINA CARDETCRE
'        Else
'           LOTE$ = TRIM$(Mid(X$, 267, 16)) ' SI ES FACTURA TOMA EL DEL PEDIDO.
'        End If
'        NPROVX = CVS(Mid(X$, 83, 16))
'        ASIGLOAHP.lblfactura.Caption = REMILAR$
'        ASIGLOAHP.LBLCLIENTE.Caption = NCLI
'        CANTIDAD = CANTAA
'        CANTI = 0
'        If LOTE$ <> "" Then ' SI TIENE LOTE
'          CANLOTE = SaldoLote(CI%, LOTE$): If CANLOTE < 0 Then CANLOTE = 0  'VEO EL SALDO DEL LOTE
'          If CANLOTE > 0 Then             'SI TIENE SALDO
'            CANTI = CANTIDAD                'GUARDO LA CANTIDAD ORIGINAL
'            If CANTI > CANLOTE Then CANTI = CANLOTE 'SI LA CANTIDAD ORIGINAL ES MAYOR AL SALDO DEL LOTE ASIGNO LA CANTIDAD DEL LOTE
'            'GENERA MOVIMIENTO DE  SALIDA CON LOTE POR EL TOTAL DE LA CANTIDAD SI ALCANZA O POR EL SALDO DEL LOTE
'           Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, rasocli$(NCLI), NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, SIG% * CANTI, NUPEDIORI&, NUORITPE%, TEXLIB$, , , PRECONIVA#) '
'           If Abs(CANTAA) - Abs(CANTI) > 0 Then GoTo 60
'          End If
'        Else
'60          CANTI = CANTIDAD - CANTI 'RECALCULA CANTIDAD A ENTREGAR
'            LOTE$ = ""
'            'SI FUE POR EL TOTAL SE VA
'            If CANTI < 0.05 Then
'               Return
'            Else
'               'SI QUEDA SALDO DE CANTIDAD GENERA UN NUEVO REGISTRO POR EL SALDO A REGISTRAR CON LOTE VACIO.
'               CANTIDAD = CANTI
'               Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, rasocli$(NCLI), NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, SIG% * CANTI, NUPEDIORI&, NUORITPE%, TEXLIB$, , , PRECONIVA#) '
'    'COMENTADA ESTAS 11 FILAS POR QUE FERNADO DIJO QUE NO SE CERRABA LA VENTANA DE ASIGNACION DE LOTES ASIGLOAHP
'                        If CI% > 0 Then
'                            TXTX$ = ""
'                            Call REPLA(TXTX$, CODEXT$(CI%), 1, 20)
'                            Call REPLA(TXTX$, DESCRIT0$(CI%), 21, 40)
'                            Call REPLA(TXTX$, FORMATNUM$(CANTIDAD, "F10.1"), 62, 11)
'                            Call REPLA(TXTX$, Space$(16), 74, 16)
'                            Call REPLA(TXTX$, FORMATNUM$(PRECONIVA#, "F12.2"), 91, 12)
'                            Call REPLA(TXTX$, FORMATNUM$(XVALO(NPROVX), "F7.0"), 104, 8)
'                            Call REPLA(TXTX$, FORMATNUM$(NORIT%, "F4"), 120, 4)
'                            ASIGLOAHP.List2.AddItem (TXTX$)
'                        End If
'                        'LUEGO SE EJECUTA CON ESTE CODIGO QUE ESTA DESPUES DE LA LINEA: Call TESTCAE09.PIDENUEVOCAE
'                        'HABILITANDO UN TIMER Y EN LA MEDIDA QUE YA HAYA CERRADO EL SERVICIO DE CAE GENERA EL PREOCESO QUE ESTA DENTRO DEL FRM ASIGLOAHP.
'    '                    If InStr(EPAC$, "AHP") > 0 Then
'    '                      Timer3.Enabled = True ' AHP 11
'    '                    End If
'
'
'            End If
'        End If
'        '
'
'Else
'        '
'        'LEO EL LOTE
'        If MODODOC% < 1 And TIFADOC% > 2 Then  ' caso nota de credito material codificado
'           LOTE$ = TRIM$(Mid(X$, 139, 64)) ' TRAE EL QUE CARGA EN EL LA RUTINA CARDETCRE
'        Else
'           LOTE$ = TRIM$(Mid(X$, 267, 16)) ' SI ES FACTURA TOMA EL DEL PEDIDO.
'        End If
'
'        NPROVX = CVS(Mid(X$, 83, 16))
'        If NPROVX < 1 Then NPROVX = 0
''        ASIGLOAHP.lblfactura.Caption = REMILAR$
''        ASIGLOAHP.LBLCLIENTE.Caption = NCLI
'        CANTIDAD = CANTAA
'        CANTI = 0
'        If LOTE$ <> "" Then ' SI TIENE LOTE
'          CANLOTE = SaldoLote(CI%, LOTE$, , , NPROVX): If CANLOTE < 0 Then CANLOTE = 0 'VEO EL SALDO DEL LOTE
'          If CANLOTE > 0 Then             'SI TIENE SALDO
'            CANTI = CANTIDAD                'GUARDO LA CANTIDAD ORIGINAL
'            If CANTI > CANLOTE Then CANTI = CANLOTE 'SI LA CANTIDAD ORIGINAL ES MAYOR AL SALDO DEL LOTE ASIGNO LA CANTIDAD DEL LOTE
'            'GENERA MOVIMIENTO DE  SALIDA CON LOTE POR EL TOTAL DE LA CANTIDAD SI ALCANZA O POR EL SALDO DEL LOTE
'           Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, rasocli$(NCLI), NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, SIG% * CANTI, NUPEDIORI&, NUORITPE%, TEXLIB$, , , PRECONIVA#) '
'           If Abs(CANTAA) - Abs(CANTI) > 0 Then GoTo 62
'          End If
'        Else
'62          CANTI = CANTIDAD - CANTI 'RECALCULA CANTIDAD A ENTREGAR
'            LOTE$ = ""
'            'SI FUE POR EL TOTAL SE VA
'            If CANTI < 0.05 Then
'               Return
'            Else
'               CANTISALDO = CANTI
'            'ACA UTILIZAR LA FUNCION DispLoteFifo$(CI%,NPROVX)
'            If NPROVX < 1 Then NPROVX = 0
'65               LOTEYCANTI$ = TRIM$(DispLoteFifo$(CI%, NPROVX))
'               POS1% = InStr(LOTEYCANTI$, "|")
'               If LOTEYCANTI$ <> "" Then
'                 CANTIDISPO = XVALO(Mid$(LOTEYCANTI$, POS1% + 1))
'                 LOTE$ = TRIM$(Mid$(LOTEYCANTI$, 1, POS1% - 1))
'                 If CANTISALDO > CANTIDISPO Then CANTI = CANTIDISPO
'               End If
'               'SI QUEDA SALDO DE CANTIDAD GENERA UN NUEVO REGISTRO POR EL SALDO A REGISTRAR CON LOTE VACIO.
'               Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, rasocli$(NCLI), NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, SIG% * CANTI, NUPEDIORI&, NUORITPE%, TEXLIB$, , , PRECONIVA#) '
'               CANTISALDO = CANTISALDO - CANTI
'               If CANTISALDO > 0.05 Then GoTo 65
'
'            End If
'        End If
'    '
'    End If
    Return
End Sub

Sub CONSUMERESERVAS(CI2%, NUMPED&, NORFA$, NC, CANTIRESERVA)

             SQLX$ = "SELECT * FROM RESERPED "
             SQLX$ = SQLX$ + " WHERE IdenLote = '" & NORFA$ & "' "
             SQLX$ = SQLX$ + " AND Cod_Inte = " & CI2% & " "
             SQLX$ = SQLX$ + " AND NroPed = " & NUMPED&
             SQLX$ = SQLX$ + " AND NroClie = " & NC
             
             Dim RESETEMP As ADODB.Recordset
             Set RESETEMP = New ADODB.Recordset
             RESETEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
             On Error Resume Next
             With RESETEMP
                 Do While Not .EOF
                    CARESE# = XVALO(!CANTRES)
                    CACOANT# = XVALO(!CANTCONS) + CANTIRESERVA
                    If CARESE# - CACOANT# >= SALCON# Then
                         !CANTCONS = CACOANT# + SALCON#
                         SALCON# = 0
                       Else
                         SALCON# = SALCON# - (CARESE# - CACOANT#)
                         !CANTCONS = XVALO(!CANTRES)
                    End If
                    .Update
                    .MoveNext
                 Loop
               
               On Error GoTo 0
             End With
             RESETEMP.Close
             Set RESETEMP = Nothing
 

End Sub

'Sub CARGARFRM_ASIGLOAHP(REMILAR$, NCLI, CI%, CANTIX, NORIT%, LOTE$, NUPROV)
'    '
'    CAMPOS$ = "Remito/A20;Código/A20;Descripción/A32;Cant./F9.O;Lote/A18;Proveedor/F7;NUORD/F0"
'    Call CARGA_ENCABEZADO(ASIGLOAHP.MSF, CAMPOS$, ASIGLOAHP, , , ANTOT)
'    '
'    If TRIM$(ASIGLOAHP.Label8.Caption) <> TRIM$(REMILAR$) Then
'       ASIGLOAHP.Label8.Caption = REMILAR$
'       ASIGLOAHP.MSF.Rows = 1
'       CAMPOS$ = "Remito/A20;Código/A20;Descripción/A32;Cant./F9.O;Lote/A18;Proveedor/F7;NUORD/F0"
'       Call CARGA_ENCABEZADO(ASIGLOAHP.MSF, CAMPOS$, ASIGLOAHP, , , ANTOT)
'    End If
'    '
'    'VER EL TEMA DEL ORDEN EN EL CASO QUE SE ABRAN LOS MOVIMIENTOS
'    With ASIGLOAHP
'     .MSF.Rows = MSF.Rows + 1
'     .MSF.TextMatrix(.MSF.Rows - 1, 1) = REMILAR$
'     .MSF.TextMatrix(.MSF.Rows - 1, 2) = CODEXT$(CI%)
'     .MSF.TextMatrix(.MSF.Rows - 1, 3) = DESCRIT0$(CI%)
'     .MSF.TextMatrix(.MSF.Rows - 1, 4) = FORMATNUM$(CANTIX, "F11,1")
'     .MSF.TextMatrix(.MSF.Rows - 1, 5) = LOTE$
'     .MSF.TextMatrix(.MSF.Rows - 1, 6) = NUPROV
'     .MSF.TextMatrix(.MSF.Rows - 1, 7) = NORIT%
'   End With
'   '
'End Sub
Sub AJUPREVE()
   '
   NC1% = XVALO(Text1.TEXT)
   Call APERBASDAT("STOCKS")
   
   FIN& = ULTREG&("!DIFPRECO")
   '
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     X$ = REGLEIDO$("!DIFPRECO", i&)
     PRENUEVO# = CVD(Mid$(X$, 45, 8))
     If PRENUEVO# > 0.00005 Then
       CI% = CVI(Mid$(X$, 1, 2))
       If CI% > 0 Then
         DOPRIL$ = TRIM$(Mid$(X$, 81, 18))
         NORIT% = CVI(Mid$(X$, 117, 2))
         PREUNI = CVD(Mid$(X$, 45, 8))
         '
         SQLX$ = "SELECT * FROM MOVISTO "
         SQLX$ = SQLX$ + " WHERE DoPriLar = '" & DOPRIL$ & "' "
         SQLX$ = SQLX$ + "AND Cod_Inte = " & CI% & " "
         SQLX$ = SQLX$ + "AND Nume_Clie = " & NC1% & " "
         SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
         Dim MOVVVISTO As ADODB.Recordset
         Set MOVVVISTO = New ADODB.Recordset
24       On Error Resume Next
         MOVVVISTO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 24
         End If
         On Error GoTo 0
         '
         With MOVVVISTO
           '.FindFirst SQLX$
25         'If .NoMatch = False Then
             '.Edit
           If Not (.EOF And .BOF) Then
             !VALOUNIT = PREUNI
             .Update
           End If
         End With
       End If
     End If
   Next i&
   MOVVVISTO.Close
   Set MOVVVISTO = Nothing
   '
End Sub

Sub CARTAMONED()
   '
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     Exit Sub
   End If
   '
   If ULTREG&("TAMONED") > 6 Then
      Call SETULTREG("TAMONED", 6)
      Call COMUNI("Tabla de Monedas Truncada\al Máximo de 6 Registros")
   End If
   '
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
   TICAMBIS = FORMATNUM(TICAMONE(MONEMI%), "F9.4")
   TICAMBX# = TICAMONE(MONEMI%)
   '
End Sub

Sub SHOWPEDIPEN()
     
     If YAINDI% = 0 Then
         '
         Screen.MousePointer = 11
         
         Call ABREPEDCLI
         'jandy sql     '

         Dim PediModifTemp As ADODB.Recordset
         Set PediModifTemp = New ADODB.Recordset
         PediModifTemp.CursorLocation = adUseClient
         
         SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
         SQLX$ = SQLX$ + " FROM PEDENCA "
         SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
         SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
         SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
         SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
         SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
         SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
         SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
         SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
         
         'Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
25       On Error Resume Next
         PediModifTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
         End If
         On Error GoTo 0
         '
         JJ& = 0
         With PediModifTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
             Do While Not .EOF
                 NRODES& = !NROPED
               TTXX$ = MKS$(NRODES&)
                 NOCO$ = !NROOCOM
               If NOCO$ <> "" Then
                 TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
               End If
                 FEX = CVINT(!FECHEMIS)
               TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
                 NCLIE% = !NROCLIE
               TTXX$ = TTXX$ + TRIM$(rasocli$(NCLIE%)) + Space$(52)
               JJ& = JJ& + 1
               Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
             .MoveNext
             Loop
           End If
         End With
         PediModifTemp.Close
         Set PediModifTemp = Nothing
         '
         Call SETULTREG("INDIPEP", JJ&)
         Call CIERRARCH("INDIPEP")
         Call FRESHECHO("INDIPEP")
         '
         Screen.MousePointer = 1
         YAINDI% = 1
     End If
     '
2112 If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay O/V's Pendientes\que Puedan Consultarse")
        GoTo 9999
     End If
     '
2110 X$ = REGSEL$("INDIPEP/Ordenes de Venta Pendientes de Entrega")
     If Len(X$) <= 4 Then
        GoTo 9999
     End If
     '
     NPX& = CVS(Left$(X$, 4))
     If NPX& < 1 Then GoTo 9999
     '
     TIDOCUM$ = "Pedido de Cliente"
     TIDOCUX$ = "Pedido ABIERTO"
     TIDOCUY$ = "Pedido Interno"
     If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
        TIDOCUM$ = "Orden de Despacho"
        TIDOCUX$ = "Orden de Despacho"
        TIDOCUY$ = "Orden de Despacho"
     End If
     '
     NUDOCU$ = TRIM$(Str$(NPX&))
     While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
     NUDOCU$ = NUDOCU$ + " "
     Screen.MousePointer = 11
     '
4    For UI& = ULTREG&("PDOCLST") To 1 Step -1
       XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
       If InStr(XX$, TIDOCUM$) > 4 Or InStr(XX$, TIDOCUX$) > 4 Or InStr(XX$, TIDOCUY$) > 4 Then
         PPXX% = InStr(XX$, "Nro.")
         If PPXX% > 0 Then
           XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
         End If
         If InStr(XX$, NUDOCU$) > 4 Then
           DOCUNU& = CVS(Left$(XX$, 4))
           GoTo 5
         End If
       End If
     Next UI&
     '
     Screen.MousePointer = 1
     Call COMUNI("Documento no Encontrado")
     Exit Sub
     '
5    Call MUESTRADOC(DOCUNU&)
     '
     GoTo 2110
     '
9999 End Sub
'
Sub DIFPREC()
    '
    CONSICLI% = 0
    NCLIE% = XVALO(Text1.TEXT)
    If NCLIE% <= 0 Then
        Call COMUNI("Numero de Cliente no Válido")
        Exit Sub
    End If
    '
    If Des% < 1 Then
2     FEX = HOY(HO$)
      Has% = FEX
      HASS$ = FECHATEX$(FEX)
      DESS$ = "01" + Mid$(HASS$, 3)
      Des% = FECHANUM(DESS$)
    End If
    '
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
       GoTo 2
    End If
    '
    'DES1# = CDbl(CVDAT(DES%))
    'HAS1# = CDbl(CVDAT(HAS%))
    DESDE1 = CVDAT(Des%)
    HASTA1 = CVDAT(Has%)
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM MOVISTO  "
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & DESDE1 & "' "
    SQLX$ = SQLX$ + "AND FechMov <= '" & HASTA1 & "' "
    SQLX$ = SQLX$ + "AND CodiMovi= 'RT' "
    SQLX$ = SQLX$ + "AND LEFT(DoSecCor,2)='FC' "
    SQLX$ = SQLX$ + "AND Nume_Clie= " & NCLIE% & " "
    SQLX$ = SQLX$ + "ORDER BY FechMov ASC "
    'Set DIFPRETMP = Stocks.OpenRecordset(SQLX$, 4)
    Dim DIFPRETMP As ADODB.Recordset
    Set DIFPRETMP = New ADODB.Recordset
25  On Error Resume Next
    DIFPRETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    With DIFPRETMP
       On Error Resume Next
       .MoveFirst
       If Err Then
            Call COMUNI("No Existen Remitos Generados")
            On Error GoTo 0
            Exit Sub
       End If
       On Error GoTo 0
       YYY$ = REGBLAN("!DIFPRECO")
       JJ& = 0
       '
       If TRIM$(TEXT21(0)) = "0" Then
         Call COMUNI("Debe Selecionar una Lista de Precio")
         Screen.MousePointer = 1
         Exit Sub
       End If
       NULIST$ = TRIM(TEXT21((0)))
       Do While Len(NULIST$) < 3
         NULIST$ = "0" + NULIST$
       Loop
            '
       NOLIST$ = "LIPRE" + NULIST$
       Do While .EOF = False
            '
            PRENUE# = PRELISTA(NOLIST$, !cod_inte)
            '
            Call REPLA(YYY$, MKI(!cod_inte), 1, 2)
            Call REPLA(YYY$, !DoPriCor, 3, 12)
            Call REPLA(YYY$, !DoSecCor, 15, 12)
            Call REPLA(YYY$, MKI(CVINT(!FechMov)), 27, 2)
            Call REPLA(YYY$, MKD(!CANTSALID), 29, 8)
            Call REPLA(YYY$, MKD(!VALOUNIT), 37, 8)
            Call REPLA(YYY$, !DOPRILAR, 81, 18)
            Call REPLA(YYY$, !DOSECLAR, 99, 18)
            Call REPLA(YYY$, MKI(!NuOrItem), 117, 2)
            Call REPLA(YYY$, MKD(PRENUE#), 45, 8)
            JJ& = JJ& + 1
            Call GRAREG("!DIFPRECO", YYY$, JJ&)
       .MoveNext
       Loop
    End With
    Set DIFPRETMP = Nothing
    Call SETULTREG("!DIFPRECO", JJ&)
    Call CIERRARCH("!DIFPRECO")
    '
    Call DIFPREC2
End Sub

Sub DIFPREC2()
    '
    List3.Clear
    '
    PRF$ = "DIFPREC"
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Diferencia de Cotización"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
    '
    VTB% = VENTABU%(PRF$ + ATRI$)
    '
    If VTB% >= 0 Then 'Si esta aprobado
      JJJ& = 0
      For U& = 1 To ULTREG&("!DIFPRECO")
        XX$ = REGLEIDO$("!DIFPRECO", U&)
        PRENUE# = CVD(Mid$(XX$, 45, 8))
        'If PRENUE# >= 0.00005 Then
          If PRENUE# < 0.00005 Then PRENUE# = CVD(Mid$(XX$, 37, 8))
          DIFUNI# = PRENUE - CVD(Mid$(XX$, 37, 8))
          DIFTOT# = DIFUNI# * CVD(Mid$(XX$, 29, 8))
            Call REPLA(XX$, MKD(DIFUNI#), 53, 8)
            Call REPLA(XX$, MKD(DIFTOT#), 61, 8)
            JJJ& = JJJ& + 1
            Call GRAREG("!DIFPRECO", XX$, JJJ&)
         'End If
      Next U&
       Call SETULTREG("!DIFPRECO", JJJ&)
      '
    Else
        Call BLANFORMU ' SI CANCELA EL WINTABU
        Exit Sub
    End If
    '
    Call COMPLETITU
    Call COMPLEDETA
    '
End Sub
Sub COMPLETITU()
    T1$ = "Cant."
    T2$ = "Código"
    T3$ = "Remito "
    T4$ = "Factura"
    T5$ = "Pr.Fact"
    T51$ = "Pr.Ajus"
    T6$ = "Difer."
    T7$ = "Importe"
    '
    T1$ = AJUSTD$(T1$, 7)
    T2$ = AJUSTI$(T2$, 12)
    T3$ = AJUSTI$(T3$, 12)
    T4$ = AJUSTI$(T4$, 12)
    T5$ = AJUSTD$(T5$, 7)
    T51$ = AJUSTD$(T51$, 7)
    T6$ = AJUSTD$(T6$, 7)
    T7$ = AJUSTD$(T7$, 7)
    '
    ADLIS$ = T1 & Space$(2) & T2 & Space$(1) & T3 & Space$(2) & T4 & Space$(1) & T5 & Space$(1) & T51 & Space$(1) & T6 & Space$(3) & T7 & Space$(1)
    Label17 = ADLIS$
    List3.Visible = True
    '
    MARSELPENDE.Visible = False
    '
End Sub
'
Sub COMPLEDETA()
    '
    Z$ = REGBLAN("!CARDEFAC")
    JJ& = 0
    '
    For U& = 1 To ULTREG&("!DIFPRECO")
        XX$ = REGLEIDO$("!DIFPRECO", U&)
        CANTI1# = CVD(Mid$(XX$, 29, 8))
        CODI1% = CVI(Mid$(XX$, 1, 2))
        RTO1$ = TRIM$(Mid$(XX$, 3, 12))
        FACT1$ = TRIM$(Mid$(XX$, 15, 12))
        PRVTA1# = CVD(Mid$(XX$, 37, 8))
        PRVCO1# = CVD(Mid$(XX$, 45, 8))
        DIFCOT1# = CVD(Mid$(XX$, 53, 8))
        IMPORDIF# = CVD(Mid$(XX$, 61, 8))
        FECHFACT% = CVI(Mid$(XX$, 27, 2))
        '
        RTOLAR$ = TRIM$(Mid$(XX$, 81, 18))
        FACTLAR$ = TRIM$(Mid$(XX$, 99, 18))
        NUMORDIT% = CVI(Mid$(XX$, 117, 2))
        '
        D1$ = FORMATNUM$(CANTI1#, "F7.1")
        CODI2$ = CODEXT(CODI1%)
        D2$ = AJUSTI(CODI2$, 16)
        D3$ = AJUSTI(RTO1$, 13)
        D4$ = AJUSTI(FACT1$, 13)
        D5$ = FORMATNUM$(PRVTA1#, "F11.4")
        D6$ = FORMATNUM$(PRVCO1#, "F11.4")
        D7$ = FORMATNUM$(DIFCOT1#, "F11.4") 'DIF
        D8$ = FORMATNUM$(IMPORDIF#, "F11.2") 'IMPORTE
        If Abs(PRVCO1#) < Abs(0.00005) Or TRIM$(D5$) = TRIM$(D6$) Then
            GoTo 3
        Else
            ADDLIS$ = D1$ & "  " & D2$ & D3$ & D4$ & D5$ & D6$ & D7$ & Space(1) & D8$ & Space(3)
            List3.AddItem ADDLIS$
            '
            DESCRI$ = AJUSTI$(FACTLAR$, 18)
            DESCRI$ = DESCRI$ + FORMATNUM(PRVCO1#, "F12.4")
            '
            Call REPLA(Z$, MKI(CODI1%), 5, 5)
            Call REPLA(Z$, DESCRI$, 7, 64)
            Call REPLA(Z$, MKI(CVINT(Text9)), 419, 2)
            Call REPLA(Z$, MKS(CANTI1#), 1, 4)
            Call REPLA(Z$, MKS(DIFCOT1#), 421, 4)
            Call REPLA(Z$, MKS(DIFCOT1#), 429, 4)
            Call REPLA(Z$, MKD(IMPORDIF#), 433, 8)
            Call REPLA(Z$, RTOLAR$, 387, 16)
            Call REPLA(Z$, MKI(NUMORDIT%), 503, 2)
            JJ& = JJ& + 1
            Call GRAREG("!CARDEFAC", Z$, JJ&)
        End If
        '
3   Next U&
    '
    Picture5.Visible = True
    '
    Call CALTOFAC
    Call CALTOFA2
    '
End Sub

Sub CALPERIB()
   '
   '
   ' *** CALCULO PERCEPCION IIBB PROVINCIA DE BUENOS AIRES
   NCLIE0 = XVALO(Text1)
   TASAX# = ALIPERBRU#(NCLIE0, MINIPER#, VARIDOC%)
   '
   Text15 = ""
   BIMPON# = XVALO(Total(2))
   If BIMPON# * TICAMONE(MONEMI%) >= MINIPER Then
       PERSALTA# = BIMPON# * TASAX# / 100
       Text15 = TRIM$(FORMATNUM$(PERSALTA#, "F9.2"))
   End If
   '
   ' *** CALCULO PERCEPCION IIBB PROVINCIA DE MISIONES
   ' *** SI ESTA CONFIGURADA LA CUENTA CONTABLE, ENTONCES SE HABILITA LA PERCEPCION DE IIBB
   CUEPERI$ = Control$("PERIBMI", "CUECONTA")
   CUENPER% = CUECOINT%(CUEPERI$)
   If TRIM$(CUEPERI$) <> "" And CUENPER% > 0 Then
        TASAX# = ALIPERBRU_MISIONES#(NCLIE0, MINIPER#, VARIDOC%)
        '
        Text25 = ""
        BIMPON# = XVALO(Total(2)): If PercepcionIBMisiones_BienDeUso = 1 Then BIMPON# = 0
        If BIMPON# * TICAMONE(MONEMI%) >= MINIPER Then
            PERSALTA# = BIMPON# * TASAX# / 100
            Text25 = TRIM$(FORMATNUM$(PERSALTA#, "F9.2"))
        End If
   End If
   
   ' *** CALCULO PERCEPCION IIBB PROVINCIA DE TUCUMAN
   ' *** SI ESTA CONFIGURADA LA CUENTA CONTABLE, ENTONCES SE HABILITA LA PERCEPCION DE IIBB
   CUEPERI$ = Control$("PERIBTU", "CUECONTA")
   CUENPER% = CUECOINT%(CUEPERI$)
   If TRIM$(CUEPERI$) <> "" And CUENPER% > 0 Then
        CodigoProvinciaTucuman$ = Control("IIBBTUCU", "CODPROVI")
        If PROVCLI$(NCLIE0) = CodigoProvinciaTucuman$ Then
            TASAX# = ALIPERBRU_TUCUMAN#(NCLIE0, MINIPER#, VARIDOC%)
            '
            Text26 = ""
            BIMPON# = XVALO(Total(2))
            PERSALTA# = BIMPON# * TICAMONE(MONEMI%) * (TASAX# / 100)
            If PERSALTA# >= MINIPER Then
                PERSALTA# = BIMPON# * TASAX# / 100
                Text26 = TRIM$(FORMATNUM$(PERSALTA#, "F9.2"))
            End If
        End If
   End If
   
   '
   ' *** CALCULO PERCEPCION IIBB CABA
   If Label7.Caption <> "N.Débito:" Then    ' NO APLICA A LAS NOTAS DE DEBITO
        If EXISTE%(LUDAT$ + "PAIBCABA.TXT") > 0 Or EXISTE%(LUDAT$ + "IIBBCABA.TXT") > 0 Then
             TASAX# = ALIPERBRU_CABA#(NCLIE0, MINIPER#, VARIDOC%)
             Text24 = ""
             BIMPON# = XVALO(Total(2))
             If BIMPON# * TICAMONE(MONEMI%) >= MINIPER Then
                 PERSALTA# = BIMPON# * TASAX# / 100
                 Text24 = TRIM$(FORMATNUM$(PERSALTA#, "F8.2"))
             End If
             If (Label7.Caption = "N.Crédito:" And AnulaFacturaComprobante$ = "") Then
                Text24 = ""
             End If
        End If
   End If
   '
   Call CALTOFA2
   '
End Sub


'
Sub CARDETPEDIPEN()
   '
   MODODOSI% = 0
   If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then MODODOSI% = 1
   '
   Screen.MousePointer = 1
   NC% = XVALO(Text1)
   Call SELECHO("SELFECHA/Entrega Solicitada Hasta:")
   FEHAS$ = TRIM$(VALACT1$("SELFECHA"))
   If FEHAS$ <> "" Then
       HAS1# = CDbl(CVDAT(FECHANUM(FEHAS$)))
     Else
       Call COMUNI("Sin Fecha de Corte.\   \Selecciona TODOS los Pedidos\Pendientes de Entrega.")
   End If
   '
   Call APERBASDAT("PEDCLI")
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PEDDETA "
   SQLX$ = SQLX$ + " WHERE NroClie = " & NC% & " "
   SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
   SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
   '
   If HAS1# > 0 Then
      'sqlx$ = sqlx$ + " AND CDBL(FeSolEnt) <= " & HAS1# & " "
      SQLX$ = SQLX$ + " AND FeSolEnt <= '" & FEHAS$ & "' "
   End If
   '
   'sqlx$ = sqlx$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
   SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= getdate()))"
   
   SQLX$ = SQLX$ + " ORDER BY NroPed, NuOrItem;" 'ESTA LINEA LA AGREGUE
   '
   'Set PEDTEMP = Ventas.OpenRecordset(sqlx$, dbOpenSnapshot)
   Dim PEDTEMP As ADODB.Recordset
   Set PEDTEMP = New ADODB.Recordset
   PEDTEMP.CursorLocation = adUseClient
25 On Error Resume Next
   PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With PEDTEMP
     On Error Resume Next
     .MoveFirst
     '
     JJ& = 0: CAITCAR% = 0
     Do While (.EOF = False)
       CODIMAT% = !CODIINT
       If CODIMAT% > 0 And CODIMAT% <= NUMAS% Then
         CAN = !CANPED - !CANTENT
           If CAPEN < 0 Then CAPEN = 0
         If !TipoPed = "ABIERTO" Then CAPEN = 999999
         NPEDID& = !NROPED
         FEMIPED% = CVINT(!FECHEMIS)
         FENTRESOL% = CVINT(!FeSolEnt)
         PRECI1# = !PreUnid
         MONEMA% = !MoneEmis
            If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
         COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         '
         If MODODOSI% = 1 Then COEFCON = 1 ' DOSIVAC CONSERVA LA MONEDA DEL PEDIDO
         '
         PRECIO = PRECI1# * COEFCON
           If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
         PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
           If PRELI < PRECIO Then PRELI = PRECIO
         PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         PRECIO = ROUND(PRECIO, 4) '
         '
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
         Call REPLA(Z$, MKS$(XVALO(!Num_Prov)), 83, 16) '(INT15-372)
           NREMIT$ = ""
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = 0
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(Z$, NUPLANO$(CODIMAT%), 315, 24)
         '
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
         '  IMPO# = CDbl(PRECIO) * CAN
         IMPO# = ROUND(PRECIO * ROUND(CAN, 1), 2)
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
         '
         Call REPLA(Z$, " ", 461, 1)
         If Left$(UCase$(SAFETEXT$(!VenRep)), 1) = "R" Then
            Call REPLA(Z$, "R", 461, 1)
         End If
           '
           NORITPE% = !NuOrItem
         Call REPLA(Z$, MKI$(NORITPE%), 503, 2)
         Call REPLA(Z$, MKS$(0), 505, 4)
         Call REPLA(Z$, MKS$(0), 509, 4)
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(Z$, NUOCOX$, 355, 16)
         '
         JJ& = JJ& + 1
         Call GRAREG("!CARDEFAC", Z$, JJ&)
         CAITCAR% = CAITCAR% + 1
         '
       End If
     .MoveNext
     Loop
   End With
   PEDTEMP.Close
   Set PEDTEMP = Nothing
   '
   Call SETULTREG("!CARDEFAC", JJ&)
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     Z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(Z$, 1, 4))) >= 0.01 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", Z$, JJ&)
     End If
   Next U&
   Call SETULTREG("!CARDEFAC", JJ&)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   Call CIERRARCH("!CARDEFAC")
   '
End Sub

Sub SELPEDIP()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   MODODOSI% = 0: If InStr(EPAC$, "DOSIVA") > 0 Then MODODOSI% = 1
   ' DETERMINA LA MODALIDAD DE PRESENTACION DEL PRECIO DE LOS ARTICULOS: RESPETA EL PEDIDO O PESIFICA
   PesificaPrecioPedido$ = TRIM$(Control$("FACTURA", "PESPREPE"))
   '
   NC = XVALO(Text1)
   Screen.MousePointer = 11
   NCLI0 = XVALO(Text1)
   Call ABREPEDCLI
   '
   'INTERVENCION SP
   SQLX$ = " EXECUTE PEDIDOS_PENDIENTES " & XVALO(NCLI0)
   SQLX$ = SQLX$ + "," & CodiEmp%
   SQLX$ = SQLX$ + ", " & EsPorPedido&
   Set PediModifTemp = FLEXCONN.Execute(SQLX$)
   'FIN 'INTERVENCION SP
'   SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
'   SQLX$ = SQLX$ + " FROM PEDENCA "
'   SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
'   SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
'   SQLX$ = SQLX$ + " WHERE PEDENCA.NroClie = " & CStr(NCLI0) & " "
'   SQLX$ = SQLX$ + " AND PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
'   SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
'   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
'   SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
'   SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
'   If EsPorPedido& > 0 Then
'      SQLX$ = SQLX$ + " AND PEDDETA.NroPed = " & EsPorPedido&
''      GoTo 251
'   End If
'
'   SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
'   'JANDY SQL
'   Dim PediModifTemp As ADODB.Recordset
'   Set PediModifTemp = New ADODB.Recordset
'   'Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
'   '
'   PediModifTemp.CursorLocation = adUseClient
'25 On Error Resume Next
'   PediModifTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'   End If
'   On Error GoTo 0
   '
   U& = 0
   SELPEDIPEN.LICOPEN.Clear
   With PediModifTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         Do While Not .EOF
           U& = U& + 1
             NRODES& = !NROPED
           TTXX$ = FORMATNUM$(NRODES&, "F6.0") + "  "
             NOCO$ = TRIM$(!NROOCOM)
           If NOCO$ <> "" Then
             TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
           End If
             FEX = CVINT(!FECHEMIS)
           TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
             NCLI0 = !NROCLIE
           TTXX$ = TTXX$ + TRIM$(rasocli$(NCLI0)) + Space$(52)
           SELPEDIPEN.LICOPEN.AddItem TTXX$
           If EsPorPedido& > 0 Then
             SELPEDIPEN.LICOPEN.Selected(U& - 1) = True
           End If
         .MoveNext
         Loop
      End If
   End With
   PediModifTemp.Close
   Set PediModifTemp = Nothing
   '
   Screen.MousePointer = 1
   '
   If EsPorPedido& < 1 Then
        If SELPEDIPEN.LICOPEN.ListCount < 1 Then ' se modifico el nombre del list por que no era el correcto
           SELPEDIPEN.Label5 = "0"
           GoTo 12                ' para que no muestre el form de seleccion
        End If
        SELPEDIPEN.Show 1
      Else
        Call SELPEDIPEN.BOTREC_Click
   End If
   '
12 CANSELE% = XVALO(SELPEDIPEN.Label5)
   If CANSELE% < 1 Then
      JJ& = 0
      GoTo 20     ' PARA BLANQUEAR ARCHIVO DE CARGA
   End If
   '

   Call APERBASDAT("PEDCLI")
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PEDDETA "
   SQLX$ = SQLX$ + " WHERE NroClie = " & NCLI0 & " "
   SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
   SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
   'SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
   SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05)"
   SQLX$ = SQLX$ + " OR (TipoPed = 'ABIERTO' AND FinVigen >= GETDATE()))"
   '
   ATXX% = 0
   For KKU& = 1 To SELPEDIPEN.LICOPEN.ListCount
      If SELPEDIPEN.LICOPEN.Selected(KKU& - 1) = True Then
         NPESEL& = XVALO(Left$(SELPEDIPEN.LICOPEN.List(KKU& - 1), 6))
         If ATXX% < 1 Then
             SQLX$ = SQLX$ + "AND (NroPed = " & NPESEL& & " "
             CondicionSql$ = CondicionSql$ + "(NroPed = " & NPESEL& & " "
             ATXX% = 1
           Else
             SQLX$ = SQLX$ + "OR NroPed = " & NPESEL & " "
             CondicionSql$ = CondicionSql$ + "OR NroPed = " & NPESEL & " "
         End If
      End If
   Next KKU&
   SQLX$ = SQLX$ + ") "
   CondicionSql$ = CondicionSql$ + ") "
   '
   SQLX$ = SQLX$ + " ORDER BY NroPed, NuOrItem;" 'ESTA LINEA LA AGREGUE
   '
   'LEO SI SE ESTA TRABAJANDO CON PEDIDOS CON LOTES
   VENCARPE$ = TRIM$(Control$("", "VENCARPE"))
   '
   ' VERIFICA SI LOS PEDIDOS ELEGIDOS TIENEN LA MISMA MONEDA. EN CASO AFIRMATIVO, PRESENTA LOS PRECIOS
   ' EN LA MONEDA ORIGINAL CAMBIANDO LA VARIABLE MONEMI% Y LA MONEDA EN PANTALLA. CASO CONTRARIO,
   ' ADVIERTE Y LOS PESIFICA CUANDO ARMA MAS ADELANTE EL DETALLE.
   CantidadMonedasDistintas% = VALOBADA("PEDDETA", "COUNT(DISTINCT(MONEEMIS))", CondicionSql$) 'READSET("SELECT DISTINCT() FROM PEDENCA WITH(NOLOCK) WHERE " & CondicionSql$)
   If CantidadMonedasDistintas% > 1 Then
        Call COMUNI("Se Ha Encontrado que los Pedidos Elegidos No Tienen la Misma Moneda.\ \Los Precios serán Unificados Conviertiéndolos a Pesos")
   Else
        If PesificaPrecioPedido$ <> "SI" Then
            EncontroMoneda% = 0
            MonedaPedidos% = XVALO(VALOBADA("PEDDETA", "MONEEMIS", CondicionSql$))
            For iMon% = 1 To ULTREG&("TAMONED")
                If MONEMI% = MonedaPedidos% Then
                    EncontroMoneda% = 1: Exit For
                Else
                    Call MONEMIS_DblClick
                End If
            Next iMon%
            If EncontroMoneda% = 0 Then
                Call COMUNI("Imposible Continuar con la Facturación.\ \Falta Moneda Número '" & CStr(MonedaPedidos%) & "' en Tabla de Cotizaciones.")
                GoTo 21
            End If
        End If
   End If
   '
   Dim PEDTEMP As ADODB.Recordset
   Set PEDTEMP = New ADODB.Recordset
   PEDTEMP.CursorLocation = adUseClient
26 On Error Resume Next
   PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
   End If
   On Error GoTo 0
   '
   With PEDTEMP
     On Error Resume Next
     .MoveFirst
     '
     JJ& = 0: CAITCAR% = 0: PEDIDOX$ = ""
     Do While (.EOF = False)
       CODIMAT% = !CODIINT
       If CODIMAT% > 0 And CODIMAT% <= NUMAS% Then
         CAN = XVALO(!CANPED) - XVALO(!CANTENT)
         CANPREP = XVALO(!AcuPrepa) - XVALO(!CANTENT)
           If CAN < 0 Then CAN = 0
           If CANPREP < 0 Then CANPREP = 0
         If !TipoPed = "ABIERTO" Then CAPEN = 999999
         NPEDID& = XVALO(!NROPED)
         'GUARDO EN UNA VARIABLE LOS PEDIDOS
         If NROPED_ANT& <> NPEDID& Then
            If JJ& = 0 Then
               PEDIDOX$ = SAFETEXT$(NPEDID&)
            Else
              PEDIDOX$ = PEDIDOX$ & ";" & SAFETEXT$(NPEDID&)
            End If
         End If
         NROPED_ANT& = NPEDID&
         
         FEMIPED% = CVINT(!FECHEMIS)
         FENTRESOL% = CVINT(!FeSolEnt)
         PRECI1# = XVALO(!PreUnid)
         'PRECI1# = ROUND(XVALO(!PreUnid), 3)'propuesta de modificacion ej ahp que pedido en u$s 1.65 x 50 importe 82.48 en vez de 82.50
         LOTEX1$ = SAFETEXT$(!idenlote)
         '
         'CALCULO DE RECARGO POR CONDICION DE PAGO
         RECXPAGO$ = VALOBADA("PEDENCA", "Rec_ConPa", "NROPED = " & NPEDID&, "NOMESS")
         CODESC = COEFDESCU(RECXPAGO$)
         If Abs(XVALO(CODESC)) < 0.05 Then CODESC = 0
         PRECI1# = PRECI1# + (PRECI1# * CODESC / 100)
         '
         MONEMA% = !MoneEmis
            If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
         COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         
         If MODODOSI% = 1 Then COEFCON = 1 ' DOSIVAC CONSERVA LA MONEDA DEL PEDIDO
         '

         PRECIO = PRECI1# * COEFCON
           'If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
         'PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
         PRELI = XVALO(!PreList) ' MODIFICADO PARA QUE TRAIGA EL PRECIO DE LISTA CON EL CUAL SE CREO EL PEDIDO
           If PRELI < PRECIO Then PRELI = PRECIO
           
         PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         PRECIO = ROUND(PRECIO, 4) '
'         PRECIO = ROUND(PRECIO, 3) 'propuesta de modificacion ej ahp que pedido en u$s 1.65 x 50 importe 82.48 en vez de 82.50
         '
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
         If InStr(1, EMPREAC$, "DOSIVAC") > 0 Then Call REPLA(Z$, MKS$(CANPREP), 1, 4)
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = ""
         Call REPLA(Z$, MKS$(XVALO(!Num_Prov)), 83, 16) '(INT15-372)
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = 0
           If InStr(1, EMPREAC$, "J Y M") > 0 Then FEREMIT% = FENTRESOL%
         '
         ESTRAX% = ESTRAZABLE%(CODIMAT%)
         If ESTRAX% = 2 Then
            NUMSE$ = "" ' NUMERO DE SERIE ASIGNADO AL PEDIDO
            ORDENITEM% = XVALO(!NuOrItem)
            CONDICION$ = " NroPed = " & NPEDID& & " AND NuOrItem = " & ORDENITEM% & " AND CodiInt = " & CODIMAT% & ""
            NUMSE$ = VALOBADA("PEDDETA", "NUMEROSERIE", CONDICION$)
            Call REPLA(Z$, NUMSE$, 139, 64)
         'AGREGADO AHP
         ElseIf ESTRAX% = 1 Then
'           If UCase$(VENCARPE$) = "CARDETLO" Then ' CONFIGURACION DE VENTANA EN PEDIDOS PARA TRABAJAR CON LOTES AHP
                'CASO AHP QUE VIENE EL LOTE DEL  PEDIDO DE LOS TRAZABLES 1
                Call REPLA(Z$, LOTEX1$, 267, 16)
'                'VALIDACAION DE LOTES EN EL CASO QUE SE TRABAJE DE ESTE MODO DESDE PEDIDOS
'           End If
         End If
         'FIN AGREGADO AHP
         '
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(Z$, NUPLANO$(CODIMAT%), 315, 24)
         '
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           ' IMPO# = CDbl(PRECIO) * CAN
           IMPO# = ROUND(PRECIO * ROUND(CAN, 1), 2)
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
         '
         Call REPLA(Z$, " ", 461, 1)
         If Left$(UCase$(SAFETEXT$(!VenRep)), 1) = "R" Then
            Call REPLA(Z$, "R", 461, 1)
         End If
           '
           NORITPE% = !NuOrItem
         Call REPLA(Z$, MKI$(NORITPE%), 503, 2)
'         Call REPLA(Z$, MKS$(0), 505, 4)
         Call REPLA(Z$, MKS$(CAN), 505, 4)
         Call REPLA(Z$, MKS$(0), 509, 4)
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(Z$, NUOCOX$, 355, 16)
           If TRIM$(NUOCOX$) <> "" Then
             If InStr(Text14, TRIM$(NUOCOX$) + ",") < 1 Then
               Text14 = Text14 + NUOCOX$ + ","
             End If
           End If
           '
         JJ& = JJ& + 1
         Call GRAREG("!CARDEFAC", Z$, JJ&)
         CAITCAR% = CAITCAR% + 1
         '
       End If
     .MoveNext
     Loop
   End With
   PEDTEMP.Close
   Set PEDTEMP = Nothing
   '.
   Text14 = TRIM$(Text14)
   If Right(Text14, 1) = "," Then Text14 = Left(Text14, Len(Text14) - 1)
'   Call AppProformasFacturadasDePedidoActivo(NC, PEDIDOX$)
20 Call SETULTREG("!CARDEFAC", JJ&)
   '
21 JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     Z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(Z$, 1, 4))) >= 0.01 Or InStr(1, EMPREAC$, "DOSIVAC") > 0 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", Z$, JJ&)
     End If
   Next U&
   Call SETULTREG("!CARDEFAC", JJ&)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   Call CIERRARCH("!CARDEFAC")
   '
End Sub
'
Sub GRAFACTURA(OKX%)
   '
   GLOB_POSIVA% = 0
   GLOB_CUITCLI$ = ""
   MonedaListaPrecios% = 0
   If InStr(1, EMPREAC$, "TANIT") > 0 Then
        MonedaListaPrecios% = MONELISTA%(XVALO(TEXT21(0)))
        If MonedaListaPrecios% <> 7 Then MonedaListaPrecios% = 0
   End If
   '
   ACU_COSTOX = 0
   If MODODOC% = 0 Then
    If Control("ADMIPED", "ITEMSCGO") = "SI" Then 'CONFIGURACION DE MEDITEA ITEMS SIN CARGO
       For J& = 1 To MSF.Rows - 1
           If MSF.TextMatrix(J&, 0) = "*" Then
              If XVALO(MSF.TextMatrix(J&, 8)) <= 0.01 Then ' IMPORTE
                 CAN = XVALO(MSF.TextMatrix(J&, 1))
                 CI1% = CODINT%(MSF.TextMatrix(J&, 2))
                 COSTOX = COSCALC#(CI1%)
                 If COSTOX <= 0 Then
                   COSTOX = COSPESOS(CI1%) 'COSTO REPOSICION
                 End If
                 COSTOX = COSTOX * CAN
                 ACU_COSTOX = COSTOX + ACU_COSTOX
              End If
           End If
       Next J&
    End If
   End If
   
   
TXTRAZA$ = TXTRAZA$ & "PUNTO 18 = " & Now & vbCrLf

   Dim NIVA7#(10), NIVA7DOL#(10), CUEIVX$(10)
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))

   COTIMONEM# = XVALO(TICAMBIS) ' para prever ajuste de moneda de emision
   If COTIMONEM# < 0.05 Then COTIMONEM# = 1
   '
   OKX% = 0
   Tipo% = TIFADOC%
TXTRAZA$ = TXTRAZA$ & "PUNTO 19 = " & Now & vbCrLf
   
   GoSub APRUEBACOPIAS       ' APROBACION FINAL Y CANTIDAD DE COPIAS
TXTRAZA$ = TXTRAZA$ & "PUNTO 20 = " & Now & vbCrLf
   
     If REPRIFAC% = 1 Then
        OKX% = 1
        Exit Sub
     End If
TXTRAZA$ = TXTRAZA$ & "PUNTO 20 = " & Now & vbCrLf
     
   GoSub SOLOREMITO          ' PARA EL CASO DE QUE LAS COPIAS DE FACTURA SEAN CERO
TXTRAZA$ = TXTRAZA$ & "PUNTO 21 = " & Now & vbCrLf
   GoSub ICONTABLE           ' IMPUTACIÓN CONTABLE DE DEBITOS Y CREDITOS
TXTRAZA$ = TXTRAZA$ & "PUNTO 22 = " & Now & vbCrLf
   GoSub VALIDATCLI          ' VALIDACION DE DATOS DEL CLIENTE
TXTRAZA$ = TXTRAZA$ & "PUNTO 23 = " & Now & vbCrLf
   GoSub ARMANUMERO          ' ARMA NUMERO DE FACTURA CORTO Y LARGO
TXTRAZA$ = TXTRAZA$ & "PUNTO 24 = " & Now & vbCrLf
   GoSub VERILIMCRE          ' VERIFICA LIMITE DE CREDITO
TXTRAZA$ = TXTRAZA$ & "PUNTO 25 = " & Now & vbCrLf
   GoSub RECALTOTA           ' RECALCULA Y VERIFICA TOTALES
TXTRAZA$ = TXTRAZA$ & "PUNTO 26 = " & Now & vbCrLf
   GoSub ACSALCUEN           ' ACTUALIZA SALDO DE LA CUENTA
TXTRAZA$ = TXTRAZA$ & "PUNTO 27 = " & Now & vbCrLf
   GoSub RENUMHOJAS          ' RENUMERA SEGUN CANTIDAD DE HOJAS A CONSUMIR
TXTRAZA$ = TXTRAZA$ & "PUNTO 28 = " & Now & vbCrLf
   GoSub VERINUMERO          ' BLOQUEA Y VERIFICA NUMERO DE FACTURA LIBRE
TXTRAZA$ = TXTRAZA$ & "PUNTO 29 = " & Now & vbCrLf
   '
   If CONTROFIS% = 2 Then    ' PARA EL CASO DE CONTROLADOR FISCAL IMPRIME Y LUEGO GRABA
      Call LIBARCH("CUECORR")
      YAIMPRIMIO% = 0
      '
      ' se agrega esto para verificar si grabó en la impresora fiscal
      OPMOVIX$ = "FC"
         If TIFADOC% = 3 Then OPMOVIX$ = "ND"
         If TIFADOC% = 6 Then OPMOVIX$ = "NC"
      VAMOVIX$ = Label13
      Call TOMANUMDELCONFIS(OPMOVIX$, VAMOVIX$, PUNTODEVENTA$, NumeComp$)
      ESTAFAC$ = PUNTODEVENTA$ + "-" + NumeComp$
      CACOPDOCF% = XVALO(FINFACTUR07.Text2(0).TEXT) ' LEE CANTIDAD COPIAS FACTURA
      '
      YAIMPRIMIO% = 0
      Call PRIFACTURA
      YAIMPRIMIO% = 1
      '
      ' ahora verifica que la factura quedó registrada en la impresora fiscal
      If VAMOVIX$ = "E" Then GoTo 29
      Call TOMANUMDELCONFIS(OPMOVIX$, VAMOVIX$, PUNTODEVENTA$, NumeComp$)
      PROXIFAC$ = PUNTODEVENTA$ + "-" + NumeComp$
      '
      If CACOPDOCF% > 0 Then
        If PROXIFAC$ = ESTAFAC$ Then
           TTXX$ = "Este Comprobante no Pudo Ser Emitido por Impresora Fiscal !!!"
           TTXX$ = TTXX$ + "\   \Causa Probable: Datos Incorrectos del Cliente."
           TTXX$ = TTXX$ + "\   \1.- Razón Social muy Larga (max. 30 pos)"
           TTXX$ = TTXX$ + "\2.- Domicilio muy Largo (max 30 pos)"
           TTXX$ = TTXX$ + "\3.- Posicion IVA o C.U.I.T. no Válidos"
           TTXX$ = TTXX$ + "\4.- Caracteres Extraños (acentos o eñes)"
           TTXX$ = TTXX$ + "\   \El Comprobante NO se Registró en la B.Datos FLEXOFT."
           TTXX$ = TTXX$ + "\   \Corrija los Datos y Re-Intente."
           Call MENSERR(24, TTXX$)
           '
           If MUECOMHAS% < 2 Then
              Call FINAL("")
           End If
        End If
      End If
      '
   End If
   '
29 Call APERBASDAT("CABAN")
   NOCLOSE% = 1
TXTRAZA$ = TXTRAZA$ & "PUNTO 30 = " & Now & vbCrLf
   GoSub GRATOTFAC           ' GRABA CAJABANC PARA FACTURA
TXTRAZA$ = TXTRAZA$ & "PUNTO 31 = " & Now & vbCrLf
   GoSub GRAVENETA           ' GRABA VENTA NETA
TXTRAZA$ = TXTRAZA$ & "PUNTO 31 = " & Now & vbCrLf
   GoSub GRAEMBALA           ' GRABA IMPORTE EMBALAJES
TXTRAZA$ = TXTRAZA$ & "PUNTO 32 = " & Now & vbCrLf
   GoSub GRABAIVA            ' graba iva inscripto Y NO INSCRIPTO
TXTRAZA$ = TXTRAZA$ & "PUNTO 33 = " & Now & vbCrLf
   GoSub GRAPERCEP           ' graba PERCEPCIONES IVA E I.BRUTOS
TXTRAZA$ = TXTRAZA$ & "PUNTO 34 = " & Now & vbCrLf
   GoSub GRAGASEXP           ' graba GASTOS DE EXPORTACION
   If REDOCEL.MODORECU > 0 Then
      Call APLICACOBRA
   End If
 TXTRAZA$ = TXTRAZA$ & "PUNTO 35 = " & Now & vbCrLf
 '
 If FACTUANTICIPO% < 1 Then GoSub ACTUFACTUANTICIPO
 
 Call GRABAR_ASIENTO_ITEM_SIN_CARGO(NFCY$, ACU_COSTOX) 'ACU_COSTOX LO CALCULA AL INICIO DE ESTA RUTINA
   
   If FACELEC% > 0 Then ' GRABA EL ARCHIVO GENERADO A PARTIR DEL CONSULTA CAE
      Call GRABACONCAE(NFCY$)
   End If
   '
   'RELACIONA LA NOTA DE DEBITO GENERADA CON LA LISTA DE NOTAS DE DEBITO A GENERAR TABLA HojaTrabajoDifCambio
   GoSub GrabaNotaDebitoDifCambio
   
   CARPETOK$ = "C:\FLEXOFT\FACELEC\"
   ' Call FILEKILL(LUCOM$ + "IDSOLCAE.TXT")  ' SI LLEGO HASTA ACA ELIMINA EL ARCHIVO IDSOLCAE.TXT
 TXTRAZA$ = TXTRAZA$ & "PUNTO 40 = " & Now & vbCrLf
   
   Call LIBARCH("CUECORR")
 TXTRAZA$ = TXTRAZA$ & "PUNTO 41 = " & Now & vbCrLf
   
   '
   
   'GRABA LA LETRA CORRESPONDIENTE A LA COMISION DEL VENDEDOR Y EL N. de VENDEDOR
   CONDI$ = "CodiCom_Lar = '" & TRIM$(NFCY$) & "'"
   If Label28.Visible = True Then
     Call ACTUREGISTRO("CAJABANC", "LETRACOMIS", CONDI$, Label28.Caption, REG&)
   End If
   '
   If Text22.Visible = True Then
     Call ACTUREGISTRO("CAJABANC", "LETRACOMIS", CONDI$, TRIM$(Text22), REGAF&, "NOMESS")
   End If
   '
   VENDXX = XVALO(TEXT21(2))
   Call ACTUREGISTRO("CAJABANC", "Vend_Cli", CONDI$, VENDXX, REGAF&, "NOMESS")
   '
 TXTRAZA$ = TXTRAZA$ & "PUNTO 42 = " & Now & vbCrLf
   
   'Call BLIMESS("Calculando Vencimientos ...")
   If TIFADOC% <> 6 Or TRIM$(DOCUDIF) = "" Then
     Call GRADEUFA(Tipo%, VA%, NRO, NC0, FFI%, CPAG%, TOTAGEN#)
   End If
 TXTRAZA$ = TXTRAZA$ & "PUNTO 43 = " & Now & vbCrLf
   
   ' APLICA AUTOMATICAMENTE LA NOTA DE CREDITO CON LA FACTURA
   ' EN LA MODALIDAD 'ANULA FACTURA'
   If TIFADOC% = 6 And AnulaFacturaComprobante <> "" Then
        CodiSuc% = 0
        NumeClie = XVALO(Text1.TEXT)
        CoComCor$ = NFCX$
        CoComLar$ = NFCY$
        ' APLICA EL IMPORTE MENOR ENTRE EL TOTAL DE LA NOTA DE CREDITO Y EL SALDO PENDIENTE DE LA FACTURA
        ImporteTotalFactura# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", "CODICOM_LAR='" & AnulaFacturaComprobante$ & "' AND TIPOMOVIM='FVEN' AND CLI_REAL=" & NumeClie))
        ImporteTotalAplicado# = XVALO(VALOBADA("APLICOBRA", "SUM(IMPOAPLI)", "CODOCOR_LAR='" & AnulaFacturaComprobante$ & "' AND NUCLIEN=" & NumeClie))
        ImporteTotalFactura# = ImporteTotalFactura# - ImporteTotalAplicado#
        If ImporteTotalFactura# > 0.005 Then
            ImpApli# = TOTAGEN#
            If (TOTAGEN# - ImporteTotalFactura#) > 0.005 Then
                 ImpApli# = ImporteTotalFactura#
            End If
            RefApli$ = "Aplicación Anula Factura"
            DocOrLar$ = AnulaFacturaComprobante
            DocOrCor$ = Mid$(AnulaFacturaComprobante, 1, 3) & Mid$(AnulaFacturaComprobante, 13, 6) & Mid$(AnulaFacturaComprobante, 3, 2)
            TotOrig# = TOTAGEN#
            IvaOrig# = 0
            Call GRAPLICOB
        Else
            Call COMUNI("Imposible Aplicar Nota de Crédito'" & NFCY$ & "'.\con la Factura '" & AnulaFacturaComprobante$ & "'.\Esta Última se Encuentra Cancelada en su Totalidad.")
        End If
   End If
   
   'CajaBanc.Close
   'SaDeuPen.Close
   NOCLOSE% = 0
   '
   'Call BLIMESS("Movimiento de Stock ...")
   NRO1 = XVALO(Text11.TEXT)
   If TIFADOC% = 6 Then NRO1 = (-1) * XVALO(Text11.TEXT)
   '
 TXTRAZA$ = TXTRAZA$ & "PUNTO 44 = " & Now & vbCrLf
   
   REMI = XVALO(Text10.TEXT)
   REMI1 = REMI: If TIFADOC% = 3 Then REMI1 = (-1) * NRO1
   If TIFADOC% = 3 Then   ' nota de debito
     If MODODOC% = 0 Then ' por diferencia de precios
       If DEBIMANUAL% = 0 Then
            Call AJUPREVE
            GoTo 80
       End If
     End If
   End If
   '
 TXTRAZA$ = TXTRAZA$ & "PUNTO 45 = " & Now & vbCrLf
   
   ' VALIDA QUE NO ESTE OCUPADO EL NUMERO DE REMITO
   ' SE TRASLADA ACA ANTES DE QUE GRABE PARA GARANTIZAR QUE NO REPITA EL NUMERO
   If TIFADOC% < 3 Then
     If COPIASREM% > 0 Then
       'Call BLOQARCH("DEUDOR2")
       If VERINURE% < 0 Then
         Call LIBARCH("CUECORR")
         NREMIRE$ = TRIM$(Text10)
         NROREMIT& = 0
         NPVX$ = Label12
         If PIVADOC% = 5 Or PIVADOC% = 7 Then
           If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "TANIT") > 0 Or InStr(EPAC$, "GABAL") > 0 Or InStr(EPAC$, "FERVI") > 0 Or InStr(EPAC$, "AHP") > 0 Or InStr(EPAC$, "J Y M") > 0 Or InStr(EPAC$, "CEYEQ") > 0 Or InStr(EPAC$, "EYM") > 0 Then
             NPVX$ = "0001"
           End If
         End If
         'VALIDACION AGREGADA PARA DOSIVAC PARA QUE MANTENGA EL PREFIJO 0001 PARA EL REMITO
        
         If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "TANIT") > 0 Or InStr(EPAC$, "GABAL") > 0 Or InStr(EPAC$, "FERVI") > 0 Or InStr(EPAC$, "AHP") > 0 Or InStr(EPAC$, "J Y M") > 0 Or InStr(EPAC$, "CEYEQ") > 0 Or InStr(EPAC$, "EYM") > 0 Or InStr(EPAC$, "DAVTECNICA") > 0 Then NPVX$ = "0001"
        
         If CONTROFIS% = 2 Then
              Call NUREMIFA_HAS(NROREMIT&, NPVX$)
            Else
              Call NUREMIFA(NROREMIT&, NPVX$)
         End If
         Text10.TEXT = TRIM$(Str$(NROREMIT&))
33       If VERINURE% < 0 Then
            NROREMIT& = NROREMIT& + 1
            Text10.TEXT = TRIM$(Str$(NROREMIT&))
            Call MENSERR(24, "Número de Remito '" + NREMIRE$ + "' Ocupado.\   \FLEXOFT ha Re-Numerado el Remito\como Número '" + TRIM$(Text10.TEXT) + "'.\  \Verifique y Vuelva a Grabar.")
            GoTo 33
         End If
       End If
     End If
   End If
   ' FIN VALIDACION DEL NUMERO DE REMITO
   '
 TXTRAZA$ = TXTRAZA$ & "PUNTO 46 = " & Now & vbCrLf
   
   ' SE AGREGA PARA LA MODALIDAD DE CREDITO PORQUE JUNTO CON LA INFO DE LA FACTURA
   ' SE ENVIA EL REMITO CON EL CUAL SE DESPACHA LA MERCADERIA, Y RECIEN EN ESTE
   ' PUNTO ESTA GARANTIZADO EL NUMERO DEL REMITO.
   If FACELEC% > 0 And chkModoCredito.Value = 1 Then
        Call SOLICITARCAE1_V1
   End If
   
   If CONSICLI% < 1 Then
       If REDOCEL.MODORECU < 1 Then Call GRAREMIFA
     Else
       Call GRACONSI
   End If
   'Call LIBARCH("DEUDOR2")
   '
 TXTRAZA$ = TXTRAZA$ & "PUNTO 47 = " & Now & vbCrLf
   
   If TIFADOC% = 3 Or TIFADOC% = 6 Then
     If TRIM$(DOCUDIF) <> "" Then
       Call APLIDICAM
     End If
   End If
 TXTRAZA$ = TXTRAZA$ & "PUNTO 48 = " & Now & vbCrLf
   '
   If MONEMI% > 1 Then
     If TIFADOC% <= 3 Then   ' graba noviniento cabeza en caso de factura o nota
       Call GRAMOCABE        ' de debito en moneda extranjera
     End If
   End If
   '
80 OKX% = 1
   Screen.MousePointer = 1
   Exit Sub
   '
APRUEBACOPIAS:          ' PRIMERA VALIDACIÓN DE DATOS *******************************
   NCLI = XVALO(Text1.TEXT)
   If MARCONOTA.Visible = True Then Exit Sub
   If ULTREG&("!CARDEFAC") < 1 Then Exit Sub
   '
   Call APROCOPIAS(OKX%)
   COPIASFAC% = XVALO(FINFACTUR07.Text2(0))
   COPIASREM% = XVALO(FINFACTUR07.Text2(1))
   If COPIASREM% < 1 And FINFACTUR07.Check8.Value < 1 Then
      Text10.TEXT = ""
   End If
   If OKX% < 0 Then Exit Sub
   If COPIASFAC% + COPIASREM% < 1 Then Exit Sub
   '
   If COPIASFAC% > 0 Then
     If XVALO(Total(6).TEXT) < 0.01 Then
       Call MENSERR(24, "Imposible Emitir Factura SIN CARGO !!!")
       Exit Sub
     End If
   End If
Return                  ' ***********************************************************
'
SOLOREMITO:             ' CASO DE QUE SE EMITE REMITO SOLO Y NO FACTURA
   If COPIASFAC% < 1 Then
      '
      ' verificar que no sea de texto libre
      FIN& = ULTREG&("!CARDEFAC")
      For i& = 1 To FIN&
         Call TRACE(20, i&, FIN&)
         X$ = REGLEIDO$("!CARDEFAC", i&)
         CI% = CVI(Mid$(X$, 5, 2))
         If CI% < 1 Then
            Call MENSERR(24, "Imposible Emitir Solo Remito\de Material No Codificado.")
            OKX% = 0
            Exit Sub
         End If
      Next i&
      '
      GoSub VERINUMERO
      Call GRAREMIFA
      Call LIBARCH("CUECORR")
      YAIMPRIMIO% = 0
      Call PRIFACTURA
      Screen.MousePointer = 1
      OKX% = (-1)
      Exit Sub
   End If
Return                  ' ***********************************************************
'
ICONTABLE:              ' IMPUTACIÓN CONTABLE DE DEBITOS Y CREDITOS
   PPVI% = PIVACLI%(NCLI)
   If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
   CUEVE$ = Control$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
   QQQ% = CUECOINT%(CUEVE$)
   '
   If TIFADOC% = 3 Or TIFADOC% = 6 Then
      If XVALO(Label31) > 0 And TIFADOC% = 6 Then ' nota de credito por reparacion
         Text16 = TRIM$(Control("SVD451S", "CUEVE-15"))
         Return
      End If
      '
      OPIMPUTA.Option1.Value = True
      Call OPIMPUTA.Option1_Click
      '
      If TRIM$(DOCUDIF) <> "" Then
        OPIMPUTA.Option2.Value = True
      End If
      '
      OPIMPUTA.Show 1
   End If
Return                  ' ***********************************************************
'
VALIDATCLI:             ' VALIDACION DE DATOS DE CLIENTE
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLI = XVALO(Text1.TEXT)
   '
   If XVALO(TEXT21(0).TEXT) > 255 Then
      MERRO$ = "Lista de Precios no Válida"
      GoTo 98
   End If
   '
   If XVALO(TEXT21(1).TEXT) > 255 Then
      MERRO$ = "Condición de Pago no Válida"
      GoTo 98
   End If
   '
   If XVALO(TEXT21(2).TEXT) > 255 Then
      MERRO$ = "Vendedor no Válido"
      GoTo 98
   End If
   '
   If XVALO(Total(6).TEXT) * TICAMBX# < 0.01 Then
      MERRO$ = "Importe no Válido"
      GoTo 98
   End If
   '

   HOII% = HOY(HOS$): FEACT = HOII%
   FEMI% = FECHANUM(Text9.TEXT)
   ' LA SIGUIENTE VALIDACION SE SUPRIME A PEDIDO DE S.IGLESIAS DE I.GUIDI
   'If FEMI% < DIANTE(FEACT) Or FEMI% > DIPOST(FEACT) Then
   '   MERRO$ = "Fecha de Emisión no Válida"
   '   GoTo 98
   'End If
   GoTo 99                 ' PASÓ LA VALIDACION DE DATOS
   '
98 Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   Exit Sub
   '
99 Return                  ' ***********************************************************
'
ARMANUMERO:                ' ARMA NUMERO DE FACTURA CORTO Y LARGO
   TICOMX$ = "FC"
   If Tipo% = 3 Then TICOMX$ = "ND"
   If Tipo% = 6 Then TICOMX$ = "NC"
   '
   'ARMADO DE NUMERO PROVISORIO DE FACTURA
   CoComLar_ExSimple$ = "": NumeroCae$ = "": FechVtoCae% = 0
   If FACELEC% > 0 Then
      PuntoVenta$ = "XXX" & XVALO(Label12): PuntoVenta$ = Right(PuntoVenta$, 4)
20    ALEATO$ = RANDSTRING(-6)
      NFCZ$ = "00000000" & ALEATO$: NFCZ$ = Right(NFCZ$, 8)
      NFCY$ = TICOMX$ & "-" & PuntoVenta$ & "-" & NFCZ$ & "-" & Label13
      
      'VALIDO QUE NO EXISTA EN NINGUNA DE LAS TABLAS SIGUIENTES
'      CJBNC& = CantRegistros("CAJABANC", "CODICOM_LAR =  '" & NFCY$ & "'", "NOMESS")
'      MVST& = CantRegistros("MOVISTO", "DOSECLAR =  '" & NFCY$ & "'", "NOMESS")
'      SDPN& = CantRegistros("SADEUPEN", "CODICOM_LAR =  '" & NFCY$ & "'", "NOMESS")
'      TNMSE& = CantRegistros("TANUMSE", "NUMFACTURA =  '" & NFCY$ & "'", "NOMESS")
'      If CJBNC& > 0 Or MVST& > 0 Or SDPN& > 0 Or TNMSE& > 0 Then
'        GoTo 20
'      End If
      'VALIDO QUE NO EXISTA
      CJBNC& = CantRegistros("CAJABANC", "CODICOM_LAR =  '" & NFCY$ & "'", "NOMESS")
      If CJBNC& > 0 Then
        GoTo 20
      End If
      Label30 = NFCY$
      '
   Else
      'ARMADO DE NUMERO DE FACTURA TRADICIONAL
      NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
      NFCY$ = TICOMX$ + "-" + Label12 + "-" + NFCZ$ + "-" + Label13
      
      If ModoExportaSimple% = 1 Then
        CoComLar_ExSimple$ = TICOMX$ + "-" + PuntoVentaExportaSimple$ + "-" + NFCZ$ + "-" + Label13
        NumeroCae$ = CAEExportaSimple$
        FechVtoCae% = VtoCAEExportaSimple%
      End If
   End If


Return                     ' ***********************************************************
'
VERILIMCRE:                 ' VERIFICA LIMITE DE CREDITO      '
   If VERLICRE% = 1 Then
     If TIFADOC% < 4 Then   ' EXCLUYE CREDITOS
       '
       For UKK& = 1 To ULTREG&("!CARDEFAC")
         XKK$ = REGLEIDO$("!CARDEFAC", UKK&)
         NROPEDXX& = XVALO(TRIM$(Mid$(XKK$, 371, 16)))
         If NROPEDXX& < 1 Then GoTo 22
         RFFKK$ = RECFILT$("AUTORENT", 1, MKS$(NROPEDXX&))
         If Len(RFFKK$) < 4 Then GoTo 22 ' NO AUTORIZADO
         REGAUTKK& = CVS(RFFKK$)
         XAUT$ = REGLEIDO$("AUTORENT", REGAUTKK&)
         If UCase$(Mid$(XAUT$, 31, 2)) <> "SI" Then GoTo 22
       Next UKK&
       GoTo 23
       '
22     LMCR# = XVALO(Text7)     ' LIMICRE(NCLI)
       SDEU# = SALDCLI#(NCLI)
       CHPA# = CHEPEACRE#(NCLI)
       TPED# = XVALO(Total(6).TEXT) * TICAMBX#
       If SDEU# + CHPA# + TPED# > LMCR# Then
         TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2))
         If TLIMI$ = "" Then TLIMI$ = "0.00"
         Call COMUNI("Imposible Emitir Este Documento.\  \Supera el Límite de Crédito Asignado\para esta Cuenta de $ " + TLIMI$ + ".\  \Informe al Supervisor de Ventas.")
         Exit Sub
       End If
     End If
   End If
   '
23 Return
   '
Return                      ' ***********************************************************
'
RECALTOTA:                  ' RE-CALCULO TOTALES DE FACTURA
   DESC$ = TRIM$(PORDESCU.TEXT)
   PODES = COEFDESCU(DESC$)
   '
   For KK% = 1 To 10: TOTAFAC#(1, KK%) = 0: Next KK%
   TOTAFAC#(1, 1) = (XVALO(BASIVA(0)) + XVALO(TLABEL28)) / (1 - PODES / 100)
   TOTAFAC#(1, 2) = XVALO(BASIVA(1)) / (1 - PODES / 100)
   TOTAFAC#(1, 3) = XVALO(Total(0).TEXT)
   TOTAFAC#(1, 4) = XVALO(Total(1).TEXT)
   TOTAFAC#(1, 5) = XVALO(BASIVA(0)) + XVALO(TLABEL28)
   TOTAFAC#(1, 6) = XVALO(Total(3).TEXT)
   TOTAFAC#(1, 7) = XVALO(Total(5).TEXT)
   TOTAFAC#(1, 8) = XVALO(Total(6).TEXT)
   TOTAFAC#(1, 9) = XVALO(BASIVA(1))
   TOTAFAC#(1, 10) = XVALO(Total(4).TEXT)
   '
   TICA = 1
   For KK% = 1 To 10
     TOTAFAC#(1, KK%) = ROUND#(TOTAFAC#(1, KK%) * TICAMBX#, 2)
     TOTAFAC#(2, KK%) = TOTAFAC#(1, KK%) / TICA
   Next KK%
   '
   TOTAGEN# = 0: TOTAGENDOL# = 0: TOTGASEXP# = 0
   NENETO# = XVALO(Total(2)) - XVALO(IEMBALA)
   NEFA7# = ROUND#(NENETO# * TICAMBX#, 2)
   NEFA7DOL# = ROUND#(NENETO#, 2)
   TOTAGEN# = TOTAGEN# + NEFA7#
   TOTAGENDOL# = TOTAGENDOL# + NEFA7DOL#
   NEMBAL7# = ROUND#(XVALO(IEMBALA) * TICAMBX#, 2)
   NEMBAL7DOL# = ROUND#(XVALO(IEMBALA), 2)
   TOTAGEN# = TOTAGEN# + NEMBAL7#
   TOTAGENDOL# = TOTAGENDOL# + NEMBAL7DOL#
   For KKI% = 0 To 2
     NIVA7#(KKI%) = ROUND#(XVALO(Total(KKI% + 3)) * TICAMBX#, 2)
     NIVA7DOL#(KKI%) = ROUND#(XVALO(Total(KKI% + 3)), 2)
     TOTAGEN# = TOTAGEN# + NIVA7#(KKI%)
     TOTAGENDOL# = TOTAGENDOL# + NIVA7DOL#(KKI%)
   Next KKI%
   PERIB7# = ROUND#(XVALO(Text15) * TICAMBX#, 2)
   PERIB7DOL# = ROUND#(XVALO(Text15), 2)
   TOTAGEN# = TOTAGEN# + PERIB7#
   TOTAGENDOL# = TOTAGENDOL# + PERIB7DOL#
   
   PERIBCABA7# = ROUND#(XVALO(Text24) * TICAMBX#, 2)
   PERIBCABA7DOL# = ROUND#(XVALO(Text24), 2)
   TOTAGEN# = TOTAGEN# + PERIBCABA7#
   TOTAGENDOL# = TOTAGENDOL# + PERIBCABA7DOL#
   
   PERIBMISIONES7# = ROUND#(XVALO(Text25) * TICAMBX#, 2)
   PERIBMISIONES7DOL# = ROUND#(XVALO(Text25), 2)
  
   PERIBTUCUMAN7# = ROUND#(XVALO(Text26) * TICAMBX#, 2)
   PERIBTUCUMAN7DOL# = ROUND#(XVALO(Text26), 2)
   
   TOTAGEN# = TOTAGEN# + PERIBMISIONES7# + PERIBTUCUMAN7#
   TOTAGENDOL# = TOTAGENDOL# + PERIBMISIONES7DOL# + PERIBTUCUMAN7DOL#
      
   PERIVA7# = ROUND#(XVALO(Text17) * TICAMBX#, 2)
   PERIVA7DOL# = ROUND#(XVALO(Text17), 2)
   TOTAGEN# = TOTAGEN# + PERIVA7#
   TOTAGENDOL# = TOTAGENDOL# + PERIVA7DOL#
   TOTGASEXP# = ROUND#(XVALO(FINFACTUR07.Text2(15)) * TICAMBX#, 2)
   TOTGASEXPDOL# = ROUND#(XVALO(FINFACTUR07.Text2(15)), 2)
   TOTAGEN# = TOTAGEN# + TOTGASEXP#
   TOTAGENDOL# = TOTAGENDOL# + TOTGASEXPDOL#
   '
Return                      ' ***********************************************************
'
ACSALCUEN:                  ' ACTUALIZA SALDO DE LA CUENTA
   LU$ = LUDAT$
   ARCHIDEU$ = "SVD202"
   ARCHICLI$ = "DEUDOR"
   ARCHIFAC$ = "FACTUR"
   rd% = 1
   Tipo% = TIFADOC%
   VA% = VARIDOC%
   NRO = XVALO(Text11.TEXT)
   NC0 = NCLI
   FFI% = FECHANUM(Text9.TEXT)
   CPAG% = XVALO(TEXT21(1).TEXT)
   '
   '
   'Call BLIMESS("Actualizando Saldo de Cuenta ...")
   SIGX% = 1: If Tipo% = 6 Then SIGX% = (-1)
       TTF# = TOTAGEN# * SIGX%
   Call GRACLI(LU$, ARCHICLI$, NC0, TTF#, DATCLI$(), 2)
   'Call CIERRARCH("DEUDOR2")
   '
   FLIM% = FECHANUM("01/01/99")
   If ULTREG&("FORMFAC") > 0 Then
      FFS$ = REGLEIDO$("FORMFAC", 2&)
      FLIM% = CVI(Mid$(FFS$, 19, 2))
   End If
Return                      ' ***********************************************************
'
RENUMHOJAS:                 ' RENUMERA PARA CONSIDERAR LAS HOJAS QUE SE CONSUMIRAN
   FORIPRE$ = TRIM$(Control$("", "FORFAC-" + Label13.Caption))
   CARENGLO& = ULTREG&("!CARDEFAC")
   ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
   If ITPORHOJA% < 1 Then
     ITPORHOJA% = 80
   End If
   CAHOJNEC% = Int(0.99 + CARENGLO& / ITPORHOJA%)
   '
'   If NFCAM% < 1 Then
'     Call NUFACTURA
'   End If
   '
27 NPP& = XVALO(Text11) + CAHOJNEC% - 1
   Text11.TEXT = TRIM$(Str$(NPP&))
   Text11.Refresh
   '

   TICOMX$ = "FC": OPMOVIX$ = "FAC"
   If Tipo% = 3 Then TICOMX$ = "ND": OPMOVIX$ = "DEB"
   If Tipo% = 6 Then TICOMX$ = "NC": OPMOVIX$ = "CRE"
   '
   If FACELEC% < 1 Then
      NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
      NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + Label13
      NFCY$ = TICOMX$ + "-" + Label12 + "-" + NFCZ$ + "-" + Label13
    Else
    'AGREGADA
     NFCX$ = TICOMX$ + "-" + Right(NFCZ$, 6) + "-" + Label13
   End If
   '
   OPMOVIX$ = "FC"
   If TIFADOC% = 3 Then OPMOVIX$ = "ND"
   If TIFADOC% = 6 Then OPMOVIX$ = "NC"
   VAMOVIX$ = Label13
   '
Return                      ' ***********************************************************
'
VERINUMERO:
   ' ESTO ES PARA VERIFICAR QUE EL ARCHIVO ESTÉ BLOQUEADO
   Call BLOQARCH("CUECORR")
   'CajaBanc.Close
   If ISLOCKED%("CUECORR") < 1 Then
      Call LIBARCH("CUECORR")
      Call MENSERR(24, "No fue Posible Abrir la Base en Modo Exclusivo.\  \Espere un Instante y Vuelva a Grabar.")
      Exit Sub
   End If
   '
   If COPIASFAC% > 0 Then
     If VERINUFA%(NFCY$) < 0 Then
       If FACELEC% > 0 Then GoTo 25
       Call LIBARCH("CUECORR")
       Call NUFACTURA
       NROREMIT& = XVALO(Text10)
       NPVX$ = Label12
       If PIVADOC% = 5 Or PIVADOC% = 7 Then
          EPAC$ = TRIM$(UCase$(EMPREAC$))
          If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "TANIT") > 0 Or InStr(EPAC$, "GABAL") > 0 Or InStr(EPAC$, "FERVI") > 0 Or InStr(EPAC$, "AHP") > 0 Or InStr(EPAC$, "J Y M") > 0 Or InStr(EPAC$, "CEYEQ") > 0 Or InStr(EPAC$, "EYM") > 0 Or InStr(EPAC$, "DAVTECNICA") > 0 Then
            NPVX$ = "0001"
          End If
       End If
       If CONTROFIS% = 2 Then
            Call NUREMIFA_HAS(NROREMIT&, NPVX$)
          Else
            Call NUREMIFA(NROREMIT&, NPVX$)
       End If
       Text10.TEXT = TRIM$(Str$(NROREMIT&))
25     If FACELEC% > 0 Then 'si ya esta en uso genera otro
         PuntoVenta$ = "XXX" & XVALO(Label12): PuntoVenta$ = Right(PuntoVenta$, 4)
31       ' NFCZ$ = RANDSTRING(-6)
         NFCZ$ = "00000000" & ALEATO$: NFCZ$ = Right(NFCZ$, 8)
         NFCY$ = TICOMX$ & "-" & PuntoVenta$ & "-" & NFCZ$ & "-" & Label13
         ' NFCY$ = TICOMX$ & "-" & "9" & PUNTOVENTA$ & "-" & NFCZ$ & "-" & Label13
         If CantRegistros("CAJABANC", "CODICOM_LAR =  '" & NFCY$ & "'", "NOMESS") > 0 Then
          ALEATO$ = RANDSTRING(-6)
          GoTo 31
         End If
       Else
         Call MENSERR(24, "Número de Comprobante '" + NFCY$ + "' Ocupado.\   \FLEXOFT ha Re-Numerado el Documento\como Número '" + TRIM$(Text11.TEXT) + "'.\  \Verifique y Vuelva a Grabar.")
         Exit Sub
      End If
     End If
   End If
'   If TIFADOC% < 3 Then
'     If COPIASREM% > 0 Then
'       If VERINURE% < 0 Then
'         Call LIBARCH("CUECORR")
'         NREMIRE$ = TRIM$(Text10)
'         NROREMIT& = 0
'         NPVX$ = Label12
'         If PIVADOC% = 5 Or PIVADOC% = 7 Then
'           EPAC$ = TRIM$(UCase$(EMPREAC$))
'           If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "TANIT") > 0 Then
'             NPVX$ = "0001"
'           End If
'         End If
'         If CONTROFIS% = 2 Then
'              Call NUREMIFA_HAS(NROREMIT&, NPVX$)
'            Else
'              Call NUREMIFA(NROREMIT&, NPVX$)
'         End If
'         Text10.TEXT = TRIM$(Str$(NROREMIT&))
'995      If VERINURE% < 0 Then
'            NROREMIT& = NROREMIT& + 1
'            Text10.TEXT = TRIM$(Str$(NROREMIT&))
'            GoTo 995
'         End If
'         Call MENSERR(24, "Número de Remito '" + NREMIRE$ + "' Ocupado.\   \FLEXOFT ha Re-Numerado el Remito\como Número '" + TRIM$(Text10.TEXT) + "'.\  \Verifique y Vuelva a Grabar.")
'         Exit Sub
'       End If
'     End If
'   End If
   '
Return                      ' ***********************************************************
'
GRATOTFAC:                  ' GRABA TOTAL DE FACTURA
   'SE AGREGA PARA ASEGURAR QUE NO SE PUEDA CONTINUAR SINO SE PUEDE OBTENER ID CAE
   If FACELEC% > 0 Then
     NumeCom = (-1)
     'SE UTILIZA UN CAST POR QUE ES UN CAMPO DE TEXTO VER SI QUEDA ASI
     
28   'IdentiCae$ = TRIM$(Str$(1 + XVALO(VALOBADA("CAJABANC", "MAX(cast(substring(IDCAE,1,LEN(IDCAE))as INT))", " TIPOMOVIM = 'FVEN'", "NOMESS"))))
     IDCAENUX& = XVALO(VALOBADA("CAJABANC", "MAX(IDCAENUM)", " TIPOMOVIM = 'FVEN'", "NOMESS"))
     If IDCAENUX& < 1 Then
        IdentiCae$ = TRIM$(Str$(1 + XVALO(VALOBADA("CAJABANC", "MAX(cast(substring(IDCAE,1,LEN(IDCAE))as INT))", " TIPOMOVIM = 'FVEN'", "NOMESS"))))
     Else
        IdentiCae$ = SAFETEXT$(1 + IDCAENUX&)
     End If
     
     If IdentiCae$ = "1" Then
        OPX% = COMALTER%("Imposible Obtener Proximo Id. \Atención !!!  Solo Inicializar  Si es la Primer Factura Electrónica.\\Re-Intentar\Inicializar Idcae\Cancelar")
        If OPX% = 1 Then GoTo 28
        If OPX% = 3 Then Screen.MousePointer = 1: Exit Sub
        'SI ES INICIALIZAR CONTINUA CON EL VALOR 1
     End If
   End If
   '
30 CoComCor$ = NFCX$
   CoComLar$ = NFCY$
'   CodiEmp% = 0
'   CodiSuc% = 0
   NumeClie = Int(GRUDIVCLI(NCLI))
   NCliReal = NCLI
   NumeProv% = 0
   FechEmi = CVDAT(FFI%)
   Referen$ = rasocli$(NC0)
   FeContab = CVDAT(FFI%)
     If REDOCEL.MODORECU > 0 Then FeContab = CVDAT(FECHANUM("31/12/10"))
   FeAcredi = CVDAT(FFI%): If chkModoCredito.Value = 1 Then FeAcredi = CVDAT(FECHANUM(txtFechaVencimiento))
   CuentaCo% = CUECOCLI%(NC0)
   If FACTUANTICIPO% > 0 Then
     CuentaCo1% = CUECOINT%(Control$("SVD451S", "CUEVE-" + TRIM$(Str$(17))))
     If CuentaCo1% > 0 Then CuentaCo% = CuentaCo1%
   End If
   TiMovi$ = "FVEN"
   OpMovi$ = OPMOVIX$
   VaMovi$ = Label13
   ModoOpMovi$ = ""
   If chkModoCredito.Value = 1 Then ModoOpMovi$ = "CR"
   CBUAlias$ = ""
   If chkModoCredito.Value = 1 Then CBUAlias$ = TRIM$(Control$("FACTURA", "CBUCREDI"))
   
   PPVI% = PIVACLI%(NCLI)
   GLOB_CUITCLI$ = CuitCli$(NCLI)
   GLOB_POSIVA% = PPVI% ' POSICION DE IVA DEL CLIENTE
   ' GUARDA EN UN CAMPO INDEPENDIENTE EL NUMERO DE CUIT DEL CLIENTE
   ' CREADO PRINCIPALMENTE PARA CONSUMIDORES FINALES
   CUITCP$ = ""
   POS1% = InStr(Text12, " - CUIT ")
   If POS1% > 0 Then
      POS1% = POS1% + 8
      CUITCP$ = TRIM$(Mid$(Text12, POS1%))
   End If
   If PPVI% = 5 Then
        CUITCP$ = Left$(TRIM$(DATFAEX07.Text3), 11)
        CUITCP$ = Mid$(CUITCP$, 1, 2) & "-" & Mid$(CUITCP$, 3, 8) & "-" & Mid$(CUITCP$, 11, 1)
   End If
       PuntoVenta$ = Label12.Caption
       If FACELEC% > 0 Then PuntoVenta$ = "XXX" & XVALO(Label12): PuntoVenta$ = Right(PuntoVenta$, 4)
   Descrip$ = Label7 + " " + Label13.Caption + " " + Label6.Caption + " " + Label12.Caption + "-" + Text11.TEXT
   NumeTal = XVALO(Label12)
   NumeCom = CDbl(Text11)
   'MODIFICADO 09
   If FACELEC% > 0 Then
     NumeCom = (-1)
     'SE UTILIZA UN CAST POR QUE ES UN CAMPO DE TEXTO VER SI QUEDA ASI
'     IdentiCae$ = TRIM$(Str$(1 + XVALO(VALOBADA("CAJABANC", "MAX(cast(substring(IDCAE,1,LEN(IDCAE))as INT))", " TIPOMOVIM = 'FVEN'", "NOMESS"))))
     IDCAENUX& = XVALO(VALOBADA("CAJABANC", "MAX(IDCAENUM)", " TIPOMOVIM = 'FVEN'", "NOMESS"))
     If IDCAENUX& < 1 Then
      IdentiCae$ = TRIM$(Str$(1 + XVALO(VALOBADA("CAJABANC", "MAX(cast(substring(IDCAE,1,LEN(IDCAE))as INT))", " TIPOMOVIM = 'FVEN'", "NOMESS"))))
     Else
      IdentiCae$ = 1 + IDCAENUX&
     End If
   End If
   'FIN MODIFICADO 09
   StatInit = 0
   ImpoDeb = TOTAGEN# * SIGX%
   ValoDola = TOTAGENDOL# * SIGX%
   Observa$ = rasocli$(NC0)
   'SE UTILIZA UN CAST POR QUE ES UN CAMPO DE TEXTO VER SI QUEDA ASI
   'VER POR QUE ESTA REPETIDO !!!!
   IDCAENUX& = XVALO(VALOBADA("CAJABANC", "MAX(IDCAENUM)", " TIPOMOVIM = 'FVEN'", "NOMESS"))
   If IDCAENUX& < 1 Then
      IdentiCae$ = TRIM$(Str$(1 + XVALO(VALOBADA("CAJABANC", "MAX(cast(substring(IDCAE,1,LEN(IDCAE))as INT))", " TIPOMOVIM = 'FVEN'", "NOMESS"))))
   Else
      IdentiCae$ = 1 + IDCAENUX&
   End If
   '
   ObservaFactu$ = (FINFACTUR07.Text4)
   INCOTERM$ = ""
   If FACELEC% > 0 And Label13 = "E" Then INCOTERM$ = TRIM$(Text18)
   CondPag% = XVALO(TEXT21(1))
   NroPedi$ = NROPEDIDO$
'   OrdCompra$ = NUOCOMPRA$
   OrdCompra$ = TRIM$(Text14)
   '(descuentos Gabal)
   Descue1$ = TRIM$(PORDESCU.TEXT)
   Descue2$ = TRIM$(Text23.TEXT)
      DESC$ = TRIM$(PORDESCU.TEXT)
      PODES = COEFDESCU(DESC$)
      DESC0$ = TRIM$(Text23.TEXT) ' DECUENTO VARIOS REALIZADO PARA GABAL
      PODES0 = COEFDESCU(DESC0$)
      TOTAL0# = XVALO(Total(0).TEXT)
      COEFIVAINCLU# = 1
      If XVALO(Total(2).TEXT) > 0.005 Then
        COEFIVAINCLU# = (XVALO(Total(2).TEXT) + XVALO(Total(3).TEXT) + XVALO(Total(4).TEXT) + XVALO(Total(5).TEXT)) / XVALO(Total(2).TEXT)
      End If
      If Label13 = "B" Then
        TOTAL0# = COEFIVAINCLU# * TOTAL0#
      End If
   TTT1# = (TOTAL0# * PODES / 100)
   IMPONET# = TOTAL0# - TTT1#
   TTT2# = (IMPONET# * PODES0 / 100)
   Impodesc1# = TTT1#
   Impodesc2# = TTT2#
   '
   COTI_DOLAR# = TICAMONE(2)
   '
   Call GRACABAN
   If FACTUANTICIPO% = 1 Then GoSub GENERA_F_A
   If ESPROFORMA& > 0 Then Call GRABAFACTURAPROFORMA(ESPROFORMA&, NFCY$)
   
'   If FACTUANTICIPO% = 2 Then ' ES UNA NOTA DE CREDITO APLICADA A UNA FACTURA.
'        Call AplicaNotaDeCredito(FacturaSeleccionada$, NFCY$)
'        FacturaSeleccionada$ = ""
'   End If

Return                      ' ***********************************************************
'
GRAVENETA:                  ' GRABA VENTA NETA

   If NetoMultipleImputacion% = 1 And TIFADOC% <> 3 And MODODOC% = 0 Then
        
        Call ImputacionVentasPorGrupoContable
        
        For iCueContable% = 1 To UBound(VentaNetaPorCuentaContable#)
            
            If VentaNetaPorCuentaContable#(iCueContable%) > 0.005 Then
                CuentaCo% = iCueContable%
                TiMovi$ = "VENTA"
                OpMovi$ = OPMOVIX$
                VaMovi$ = Label13
                Descrip$ = "Mercaderia Vendida " + NFCX$
                StatInit = 0
                multiImpu_NENETO# = VentaNetaPorCuentaContable#(iCueContable%)
                multiImpu_NEFA7# = ROUND#(multiImpu_NENETO# * TICAMBX#, 2)
                multiImpu_NEFA7DOL# = ROUND#(multiImpu_NENETO#, 2)
                ImpoDeb = multiImpu_NEFA7# * SIGX% * (-1)
                ValoDola = multiImpu_NEFA7DOL# * SIGX% * (-1)
                Call GRACABAN
            End If
        
        Next iCueContable%

   Else
        'graba venta neta - Verificar que QQQ sea la IC de combinacion Cliente-Tipo de Comprobante
        TOTANETX# = NEFA7#
        If TOTANETX# >= 0.005 Then
          QQ1% = CUECOINT%(Text16)
          If FACTUANTICIPO% > 0 Then
               QQ2% = CUECOINT%(Control$("SVD451S", "CUEVE-" + TRIM$(Str$(18))))
               If QQ2% > 0 Then QQ1% = QQ2%
          End If
          If QQ1% < 1 Then QQ1% = QQQ%
          CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
          TiMovi$ = "VENTA"
          OpMovi$ = OPMOVIX$
          VaMovi$ = Label13
          Descrip$ = "Mercaderia Vendida " + NFCX$
          StatInit = 0
          ImpoDeb = TOTANETX# * SIGX% * (-1)
          ValoDola = NEFA7DOL# * SIGX% * (-1)
          Call GRACABAN
          If FACTUANTICIPO% > 0 Then
             GoSub GENERA_F_A
          End If
        End If
   End If
   
Return                      ' ***********************************************************
'
GRAGASEXP:                  ' GRABA GASTOS DE EXPORTACION
   TOTANETX# = TOTGASEXP#
   If TOTANETX# >= 0.005 Then
     CUEMBA$ = Control$("SVD451S", "CUEVE-13")
     If CUEMBA$ = "" Or CUECOINT%(CUEMBA$) < 1 Then CUEMBA$ = Text16
     QQ1% = CUECOINT%(CUEMBA$)
     If QQ1% < 1 Then QQ1% = QQQ%
     CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
     TiMovi$ = "VENTA"
     OpMovi$ = OPMOVIX$
     VaMovi$ = Label13
     Descrip$ = TRIM$(FINFACTUR07.Text10) + " " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = TOTGASEXPDOL# * SIGX% * (-1)
     Call GRACABAN
   End If
Return                      ' ***********************************************************

GRAEMBALA:                   ' GRABA IMPORTE EMBALAJES
   TOTANETX# = NEMBAL7#
   If TOTANETX# >= 0.005 Then
     CUEMBA$ = Control$("SVD451S", "CUEVE-12")
     If CUEMBA$ = "" Or CUECOINT%(CUEMBA$) < 1 Then CUEMBA$ = Text16
     QQ1% = CUECOINT%(CUEMBA$)
     If QQ1% < 1 Then QQ1% = QQQ%
     CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
     TiMovi$ = "EMBAL"
     OpMovi$ = OPMOVIX$
     VaMovi$ = Label13
     Descrip$ = "Gasto Embalajes " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = NEMBAL7DOL# * SIGX% * (-1)
     Call GRACABAN
   End If
Return                      ' ***********************************************************
'
GRABAIVA:                   ' GRABA IVA INSCRIPTO Y NO INSCRIPTO
   'graba iva inscripto
   CUEIVX$(0) = Control$("SVD451S", "CUEIVA")
   CUEIVX$(1) = Control$("SVD451S", "CUEIVB")
   If TRIM$(CUEIVX$(1)) = "" Or CUECOINT%(CUEIVX$(1)) < 1 Then
     CUEIVX$(1) = CUEIVX$(0)
   End If
   For KKI% = 0 To 1
     TOTANETX# = NIVA7#(KKI%)
     If TOTANETX# >= 0.005 Then
       CuentaCo% = CUECOINT%(CUEIVX$(KKI%))
       TiMovi$ = "IVA"
       OpMovi$ = "DeFi"
       VaMovi$ = PORIVA(KKI%) + " %"
       Descrip$ = "Iva Debito Fiscal " + NFCX$
       StatInit = 0
       ImpoDeb = TOTANETX# * SIGX% * (-1)
       ValoDola = NIVA7DOL#(KKI%) * SIGX% * (-1)
       BasImpo# = ROUND#(XVALO(BASIVA(KKI%)) * TICAMBX#, 2)
       If KKI% = 0 Then
          BasImpo# = BasImpo# + ROUND#(XVALO(TLABEL28) * TICAMBX#, 2)
       End If
       Call GRACABAN
       If FACTUANTICIPO% > 0 Then
        GoSub GENERA_F_A
       End If
     End If
   Next KKI%
   BasImpo# = 0
   '
   'graba iva NO inscripto
   TOTANETX# = NIVA7#(2)
   If TOTANETX# >= 0.005 Then
     CUEIVA$ = Control$("SVD451S", "CUEADU")
     CuentaCo% = CUECOINT%(CUEIVA$)
     TiMovi$ = "IVA"
     OpMovi$ = "DeFi"
     VaMovi$ = "RNI"
     Descrip$ = "IVA Resp.No Inscripto " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = NIVA7DOL#(2) * SIGX% * (-1)
     BasImpo# = ROUND#(XVALO(BASIVA(2)) * TICAMBX#, 2)
     If FACTUANTICIPO% > 0 Then
        GoSub GENERA_F_A
     End If
     Call GRACABAN
   End If
   BasImpo# = 0
Return                      ' ***********************************************************
'
GRAPERCEP:                  ' PERCEPCIONES I.BRUTOS E IVA

   ' *** GRABA PERCEPCION IIBB PROVINCIA DE BUENOS AIRES
   TOTANETX# = PERIB7#
   If TOTANETX# >= 0.005 Then
     CUEIVA$ = Control$("PERIBRU", "CUECONTA")
     CuentaCo% = CUECOINT%(CUEIVA$)
     TiMovi$ = "IBRU"
     OpMovi$ = "IB-B"
     VaMovi$ = "PERC"
     Descrip$ = "Perc.IB.Bs.As. " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = PERIB7DOL# * SIGX% * (-1)
     BasImpo# = ROUND#(XVALO(Total(2)) * TICAMBX#, 2)
     Call GRACABAN
   End If
   BasImpo# = 0
   '
   ' *** GRABA PERPCEPCION IIBB CABA
   TOTANETX# = PERIBCABA7#
   If TOTANETX# >= 0.005 Then
     CUEIVA$ = Control$("PERIBCA", "CUECONTA")
     CuentaCo% = CUECOINT%(CUEIVA$)
     TiMovi$ = "IBRU"
     OpMovi$ = "IB-C"
     VaMovi$ = "PERC"
     Descrip$ = "Perc.IB. CABA. " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = PERIBCABA7DOL# * SIGX% * (-1)
     BasImpo# = ROUND#(XVALO(Total(2)) * TICAMBX#, 2)
     Call GRACABAN
   End If
   BasImpo# = 0
   '
   ' *** GRABA PERPCEPCION IIBB MISIONES
   TOTANETX# = PERIBMISIONES7#
   If TOTANETX# >= 0.005 Then
     CUEIVA$ = Control$("PERIBMI", "CUECONTA")
     CuentaCo% = CUECOINT%(CUEIVA$)
     TiMovi$ = "IBRU"
     OpMovi$ = "IB-MI"
     VaMovi$ = "PERC"
     Descrip$ = "Perc.IB. MISIONES. " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = PERIBMISIONES7DOL# * SIGX% * (-1)
     BasImpo# = ROUND#(XVALO(Total(2)) * TICAMBX#, 2)
     Call GRACABAN
   End If
   BasImpo# = 0
   '
   ' *** GRABA PERPCEPCION IIBB TUCUMAN
   TOTANETX# = PERIBTUCUMAN7#
   If TOTANETX# >= 0.005 Then
     CUEIVA$ = Control$("PERIBTU", "CUECONTA")
     CuentaCo% = CUECOINT%(CUEIVA$)
     TiMovi$ = "IBRU"
     OpMovi$ = "IB-TU"
     VaMovi$ = "PERC"
     Descrip$ = "Perc.IB. TUCUMAN. " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = PERIBTUCUMAN7DOL# * SIGX% * (-1)
     BasImpo# = ROUND#(XVALO(Total(2)) * TICAMBX#, 2)
     Call GRACABAN
   End If
   BasImpo# = 0
   
   '
   ' *** GRABA PERCEPCION IVA
   TOTANETX# = PERIVA7#
   If TOTANETX# >= 0.005 Then
     CUEPER$ = Control$("SVD451S", "CUERET")
     CuentaCo% = CUECOINT%(CUEPER$)
     TiMovi$ = "IVA"
     OpMovi$ = "DeFi"
     VaMovi$ = "PER-IVA"
     Descrip$ = "Percep.IVA " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = PERIVA7DOL# * SIGX% * (-1)
     BasImpo# = ROUND#(XVALO(Total(2)) * TICAMBX#, 2)
     Call GRACABAN
   End If
   BasImpo# = 0
Return                      ' ***********************************************************
'
GENERA_F_A:
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  SQLX$ = "SELECT * FROM FACTUANTI WHERE DOCU_ORIG = ''"
  rs.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With rs
     .AddNew
     !TIPOMOVIM = TiMovi$
     !FECHEMISI = FechEmi
     !DOCU_ORIG = NFCY$
     !NUCLIEN = Int(GRUDIVCLI(NCLI))
     !Cli_Real = NCLI
     !refecom = rasocli$(NCLI)
     On Error Resume Next
     !NumeCom = XVALO(Text11)
     On Error GoTo 0
     If FACELEC% > 0 Then
       !NumeCom = (-1)
     End If
     !Status = 0

     '!IMPODEBE = Abs(ImpoDeb / COTI_DOLAR#) 'corregido 13/01/20
     !IMPODEBE = Abs(ImpoDeb)
     !IMPOHABE = 0
     VaMovi1$ = VaMovi$
     If TiMovi$ = "VENTA" And XVALO(TLABEL28) Then VaMovi1$ = "EXENTO"
     !VARIMOVIM = VaMovi1$
     !COTIZA_DOLAR = COTI_DOLAR#
     If MONEMI% > 1 Then
        BAIMP# = Abs(BasImpo# / COTI_DOLAR#)
        !IMPODEBE = Abs(ImpoDeb / COTI_DOLAR#)
     Else
        BAIMP# = Abs(BasImpo#)
        '!IMPODEBE = Abs(ImpoDeb / COTI_DOLAR#)'corregido 13/01/20
        !IMPODEBE = Abs(ImpoDeb)
     End If
     'TOMA EL VALOR DE LA CAJA DE TEXTO PARA EL CASO QUE HAYA IMPORTES EXCENTOS
     If TiMovi$ = "VENTA" Then BAIMP# = XVALO(TLABEL28)
     If XVALO(VaMovi$) = XVALO(PORIVA(0)) And XVALO(BASIVA(0)) > 0 Then BAIMP# = XVALO(BASIVA(0))
     If XVALO(VaMovi$) = XVALO(PORIVA(1)) And XVALO(BASIVA(1)) > 0 Then BAIMP# = XVALO(BASIVA(1))
     !BASEIMPODEBE = Abs(BAIMP#)
     !BASEIMPOHABE = 0
     !Mone_Emis = MONEMI%

     .Update
   End With
   '
   rs.Close
   Set rs = Nothing
   
   Unload ANTSINAP09
   '
Return
ACTUFACTUANTICIPO:
      J& = 1
      'VALIDAR SI ESTA VACIO NO HAY NADA PARA ACTUALIZAR
      On Error Resume Next
      NFACT1$ = TRIM$(ANTSINAP09.MSF2.TextMatrix(1, 2))
      On Error GoTo 0
      If TRIM$(NFACT1$) = "" Then GoTo 889
      
      For i& = 1 To ANTSINAP09.MSF2.Rows - 1
           NFACT$ = TRIM$(ANTSINAP09.MSF2.TextMatrix(i&, 2))
           SALDOX# = XVALO(ANTSINAP09.MSF2.TextMatrix(i&, 3))
           IMPOR105# = XVALO(ANTSINAP09.MSF2.TextMatrix(i&, 5))
           GRAV105# = XVALO(ANTSINAP09.MSF2.TextMatrix(i&, 6))
           IMPOR210# = XVALO(ANTSINAP09.MSF2.TextMatrix(i&, 7))
           GRAV210# = XVALO(ANTSINAP09.MSF2.TextMatrix(i&, 8))
           EXENTO# = XVALO(ANTSINAP09.MSF2.TextMatrix(i&, 9))
           
           IMPOAPLICAR# = GRAV210# + IMPOR210# + GRAV105# + IMPOR105# + EXENTO#
           TiMovi1$ = "FVEN"
           BASEIMPODE# = 0
           BASEIMPOHA# = 0
           VaMovi1$ = Right$(NFACT$, 1)
880        If J& = 2 Then
                  IMPOAPLICAR# = 0
                  TiMovi1$ = "VENTA"
                  If XVALO(TLABEL28) > 0 Then VaMovi1$ = "EXENTO"
                  IMPOAPLICAR# = GRAV105# + GRAV210# + EXENTO#
                  BASEIMPOHA# = XVALO(ANTSINAP09.MSF2.TextMatrix(i&, 9))
           
           ElseIf J& = 3 Then
                If IMPOR105# > 0.005 Then
                  IMPOAPLICAR# = 0
                  TiMovi1$ = "IVA"
                  IMPOAPLICAR# = XVALO(ANTSINAP09.MSF2.TextMatrix(i&, 5))
                  VaMovi1$ = "10.50 %"
                  BASEIMPOHA# = XVALO(ANTSINAP09.MSF2.TextMatrix(i&, 6))
                Else
                  GoTo 888
                End If
           ElseIf J& = 4 Then
                If IMPOR210# > 0.005 Then
                  IMPOAPLICAR# = 0
                  TiMovi1$ = "IVA"
                  IMPOAPLICAR# = XVALO(ANTSINAP09.MSF2.TextMatrix(i&, 7))
                  VaMovi1$ = "21.00 %"
                  BASEIMPOHA# = XVALO(ANTSINAP09.MSF2.TextMatrix(i&, 8))
                Else
                  GoTo 888
                End If
           End If
           Set RS10 = New ADODB.Recordset
           SQLX$ = "SELECT * FROM FACTUANTI WHERE DOCU_ORIG = '' "
333        On Error Resume Next
           RS10.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 333
           End If
           On Error GoTo 0
           '
           With RS10
                  .AddNew
                  On Error Resume Next
                  !CODIEMPR = CodiEmp%
                  On Error GoTo 0

                  !DOCU_ORIG = NFACT$
                  !VARIMOVIM = VaMovi1$
                  !DOCU_APLI = NFCY$
                  !IMPODEBE = 0
                  !IMPOHABE = IMPOAPLICAR#
                  !NUCLIEN = Int(GRUDIVCLI(NCLI))
                  !Cli_Real = NCLI
                  !refecom = rasocli$(NCLI)
                  !NumeCom = XVALO(Mid$(NFACT$, 9, 8))
                  !FECHEMISI = FechEmi
                  !Status = 0
                  !TIPOMOVIM = TiMovi1$
                  !COTIZA_DOLAR = COTI_DOLAR#
                  !BASEIMPODEBE = 0
                  !BASEIMPOHABE = BASEIMPOHA#
                  !USUARIO = USERNAME$
                  !OBSERVACIONES = "APLICACION FACTURA ANTICIPO"
                  !Mone_Emis = MONEMI%
                  .Update
888                J& = J& + 1
                  If J& <= 4 Then
                       GoTo 880
                  End If
            End With
     Next i&
889  On Error Resume Next
     '
     IMPOX2# = XVALO(ANTSINAP09.BASIVA(2))
     IMPOX3# = XVALO(ANTSINAP09.BASIVA(3))
     TOTIMPOX# = IMPOX2# + IMPOX3#
     If TOTIMPOX# > 0 Then
        Call GRABA_ASIENTO_FACTURA_ANTICIPO(NCLI, TOTIMPOX#, NFACT$)
     End If
     Unload ANTSINAP09
     '
     On Error GoTo 0
     Return
GrabaNotaDebitoDifCambio:
   If TIFADOC% = 3 Or TIFADOC% = 6 Then
    ESNOTADEBITOAUTO% = 0
    TXX1$ = REGLEIDO$("!CARDEFAC", 1)
    TXX2$ = Mid$(TXX1$, 7, 64)
    If InStr(TXX1$, "Dif Cambio s/ Rec: ") > 0 Then
       RECIBOX$ = TRIM$(Mid$(TRIM$(TXX1$), 22, 18))
       SOLNOTADEB.TXTNOTADEBITO = NFCY$
       SOLNOTADEB.GRABAHOJADETRABAJO
    End If
  End If
  Return
End Sub
'
Sub RECONCLI(NCX7, hubore%)
   '
   hubore% = 0
   '
   Dim SALCONSI(999999), SALCONUE(999999)
   Dim DOCUSA$(2048)
   If NCX7 > 0 Then
     If rasocli$(NCX7) <> "" Then
       Screen.MousePointer = 11
'       FIN& = ULTREG&("CONSICLI")
'       For KLK& = 1 To 32767: SALCONSI(KLK&) = 0: SALCONUE(KLK&) = 0: Next KLK&
'       RFF$ = RECFILT$("CONSICLI", 1, MKI$(NCX7%))
'       For U2& = 1 To Len(RFF$) Step 4   ' FIN&
'         If U2& > Len(RFF$) Then Exit For
'         U& = CVS(Mid$(RFF$, U2&, 4))
'         Call TRACE(20, U&, FIN&)
'         XX$ = REGLEIDO$("CONSICLI", U&)
'         NCX8% = CVI(Left$(XX$, 2))
'         If NCX8% = NCX7% Then
'           CIXI% = CVI(Mid$(XX$, 35, 2))
'           If CIXI% > 0 Then
'             If CIXI% <= NUMAS% Then
'               CACON = CVD(Mid$(XX$, 37, 8)) - CVD(Mid$(XX$, 45, 8))
'               SALCONSI(CIXI%) = SALCONSI(CIXI%) + CACON
'             End If
'           End If
'         End If
'       Next U2&
       '
'NMS% = NUMAS%
'For UI% = 1 To NMS%
'If SALCONSI(UI%) > 0 Then
'AA1 = CODEXT$(UI%)
'BBB = CCC
'End If
'Next UI%

       FEXI = FETEXCOR1$(FECHANUM("22/08/08"))
       CMO$ = TRIM$(Control$("", "CODREFAC"))
       If Len(CMO$) <> 2 Then CMO$ = "RT"
       '
       SQLX$ = " SELECT * FROM MOVISTO "
       SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCX7 & " "
       SQLX$ = SQLX$ + " AND CodiMovi = '" & CMO$ & "' "
       SQLX$ = SQLX$ + " AND DoPriLar <> '' "
       SQLX$ = SQLX$ + " AND ((DoSecLar is null) OR (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
       SQLX$ = SQLX$ + " AND FechMov < '" & FEXI & "' "
       '
       Dim CONSITMP As ADODB.Recordset
       Set CONSITMP = New ADODB.Recordset
       CONSITMP.CursorLocation = adUseClient
25     On Error Resume Next
       CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
       End If
       On Error GoTo 0
       '
       FIN& = 0
       With CONSITMP
         If Not (.BOF And .EOF) Then
           .MoveFirst
           Do While Not .EOF
              CIXI% = !cod_inte
              SALCONUE(CIXI%) = SALCONUE(CIXI%) + !CANTSALID - !CANTINGRE
              FIN& = FIN& + 1
           .MoveNext
           Loop
         End If
       End With
       '
       CANCASOS% = 0
       For CII% = 1 To NUMAS%
          If SALCONSI(CII%) > 0 Then CANCASOS% = CANCASOS% + 1
          If SALCONSI(CII%) <> SALCONUE(CII%) Then
             hubore% = 1
          End If
       Next CII%
       If hubore% < 1 Then
          If CANCASOS% = FIN& Then
             Exit Sub
          End If
       End If
       '
       CONSITMP.Close
       Set CONSITMP = Nothing
       SQLX$ = " UPDATE MOVISTO SET DoSecCor = 'S/Cargo', DOSECLAR = 'Sin Cargo'"
       SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCX7 & " "
       SQLX$ = SQLX$ + " AND CodiMovi = '" & CMO$ & "' "
       SQLX$ = SQLX$ + " AND DoPriLar <> '' "
       SQLX$ = SQLX$ + " AND ((DoSecLar is null) OR (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
       SQLX$ = SQLX$ + " AND FechMov < '" & FEXI & "' "
       FLEXCONN.Execute SQLX$
       '
       For CII% = 1 To NUMAS%
          If SALCONSI(CII%) > 0 Then
             AA1 = CODEXT$(CII%)
             SQLX$ = " SELECT * FROM MOVISTO "
             SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCX7 & " "
             SQLX$ = SQLX$ + "AND Cod_Inte = " & CII% & " "
             SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
             SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
             SQLX$ = SQLX$ + " AND DoPriLar <> '' "
             SQLX$ = SQLX$ + " AND ((DoSecLar is null) OR (DoSecLar = 'Sin Cargo')) "
             'SQLX$ = SQLX$ + " AND FechMov < '" & FEXI & "' "
             '
             Set CONSITMP = New ADODB.Recordset
             CONSITMP.CursorLocation = adUseClient
26              On Error Resume Next
             CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText
             If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                On Error GoTo 0
                Call CapturaErrorOpen
                GoTo 26
             End If
             On Error GoTo 0
             '
             JJ& = 0
             With CONSITMP
               If Not (.BOF And .EOF) Then
                 .MoveFirst
                 Do While Not .EOF
                    If !CANTSALID > 0 Then
                      JJ& = JJ& + 1
                      DOCUSA$(JJ&) = !DOPRILAR
                    End If
                 .MoveNext
                 Loop
               End If
             End With
             '
             For JJ1& = 1 To JJ&
               For JJ2& = JJ1& + 1 To JJ&
                 If DOCUSA$(JJ1&) < DOCUSA$(JJ2&) Then
                   Call SWAP(DOCUSA$(JJ1&), DOCUSA$(JJ2&))
                 End If
               Next JJ2&
             Next JJ1&
             '
             For JJ1& = 1 To JJ&
               SQLX$ = " SELECT * FROM MOVISTO "
               SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCX7 & " "
               SQLX$ = SQLX$ + "AND Cod_Inte = " & CII% & " "
               SQLX$ = SQLX$ + " AND CodiMovi = '" & CMO$ & "' "
               SQLX$ = SQLX$ + " AND DoPriLar = '" & DOCUSA$(JJ1&) & "' "
               SQLX$ = SQLX$ + " AND CANTSALID > 0 "
               SQLX$ = SQLX$ + " AND FechMov < '" & FEXI & "' "
               '
               Set CONSITMP = New ADODB.Recordset
               CONSITMP.CursorLocation = adUseClient
27                On Error Resume Next
               CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText
               If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                    On Error GoTo 0
                    Call CapturaErrorOpen
                    GoTo 27
               End If
               On Error GoTo 0
               '
               With CONSITMP
                 On Error Resume Next
                 .MoveFirst
                 On Error GoTo 0
                 If Not (.BOF And .EOF) Then
                   Do While Not .EOF
                     If XVALO(!CANTSALID) > 0 And XVALO(!CANTSALID) <= SALCONSI(CII%) Then
                       !DoSecCor = ""
                       !DOSECLAR = ""
                       On Error Resume Next
                       .Update
                       If Err Then
                          On Error GoTo 0
                          GoTo 49
                       End If
                       On Error GoTo 0
                       SALCONSI(CII%) = SALCONSI(CII%) - XVALO(!CANTSALID)
                       If SALCONSI(CII%) < 0.5 Then GoTo 49
                     End If
                   .MoveNext
                   Loop
                 End If
               End With
               '
             Next JJ1&
          End If
          '
49     Next CII%
       '
     End If
   End If
   '
End Sub
'
Public Function VERIFSALIDA%(Optional MODO$)
   '
   'DEVUELVE(0 NINGUNO)-(1 SI ES PUNTO DE VTA NORMAL)-( 2 B.CAP.)
   '
   VRFSLIDA% = 0
   If Control$("FACTURA", "SALIDA") = "FELECT" Then
      If TRIM$(Label12) = Control("FACTURA", "PUVENELE") Or TRIM$(Label12) = Control("FACTURA", "PUVECRED") Then
           VRFSLIDA% = 1
         ElseIf TRIM$(Label12) = Control("FACTURA", "PUVENCAP") Then
           VRFSLIDA% = 2
         ElseIf TRIM$(Label12) = Control("FACTURA", "PUVEFAEX") Then
           VRFSLIDA% = 3
         Else
           If TRIM$(UCase$(MODO$)) <> "NOMESS" Then
              Call MENSERR(24, "Punto de Venta '" + TRIM$(Label12) + "' No Coincide\con los Configurados para Factura Electrónica")
           End If
           VRFSLIDA% = 0
      End If
   End If
   '
99 VERIFSALIDA% = VRFSLIDA%
   '
End Function

Function VALIDAPUNVEN%()
   ' VERIFICA DE ACUERDO AL PUNTO DE VENTA SI ES EL CONFIGURADO PARA FACTURAR BIENES DE CAPITALES
   ' QUE LOS ARTICULOS SELECCIONADOS ESTEN COMO TALES
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   VALIPUVEN% = 1
   If Control$("FACTURA", "SALIDA") <> "FELECT" Then GoTo 99
    
   VERISAL% = VERIFSALIDA%("NOMESS") 'DEVUELVE 1 SI ES PUNTO DE VTA NORMAL- 2 B.CAP. - 3 FACTURA DE EXPORTACION 0 NINGUNO DE LOS 3
   If VERISAL% < 1 Then GoTo 99
   '
   PVTMODOCREDITO$ = Control("FACTURA", "PUVECRED")
   If VERISAL% = 1 Then TIPODEPROD% = 0: TIPROX$ = "(Comun)"  ' comunes
   'If VERISAL% = 2 Or Label12 = PVTMODOCREDITO$ Then TIPODEPROD% = 1: TIPROX$ = "(Bien.Cap.)"  ' bienes de capital
   If VERISAL% = 2 Then TIPODEPROD% = 1: TIPROX$ = "(Bien.Cap.)"  ' bienes de capital
   '
   POXIVA% = PIVACLI%(TRIM$(Text1))
   
   ESEXCEPCION_EYM% = 0
   If InStr(1, EMPREAC$, "EYM") > 0 Then
    If VERISAL% = 1 Then 'SOLO VALIDA SI ES PUNTO DE VENTA COMUN SI ES P.VTA B.CAP VALIDA IGUAL QUE ANTES
       'SI ES PUNTO DE VENTA COMUN Y NO INCLUYE EQUIPOS (CONFIGURADOS EN LA TABLA 'COMBINA_IVA') POR MAS QUE ESTEN CONFIGURADOS
       ESEXCEPCION_EYM% = EXCEPCION_EYM_ACCESORIOS%(VERISAL%) 'COMO B.CAPITAL DEBE PERMITIR FACTURARSE COMO COMUNES
                                                              'EN EL PUNTO DE VENTA COMUN
    End If
   End If
   '
   
   If VERISAL% = 2 Then                     ' bienes de capital
     '
     If POXIVA% < 1 Or POXIVA% = 5 Then
        Call MENSERR(24, "Imposible Facturar como Bienes de Capital\a Consumidor Final o Extranjero.")
        VALIPUVEN% = 0
        GoTo 99
     End If
   
     NUMERO_CUIT$ = CuitCli$(XVALO(Text1))
     If VALICUIT%(NUMERO_CUIT$) <= 0 Then
       Call MENSERR(24, "Debe Tener un C.U.I.T. Válido\Para la Modalidad de Facturación\Bienes de Capitales")
       VALIPUVEN% = 0
       GoTo 99
     End If
     '
'     If IsNumeric(PORDESCU) = True Then
'        Call MENSERR(24, "No se Podran Aplicar Descuentos Generales\a Facturación de Bienes de Capital\Aplicar los Mismos Manualmente en Ventana de Carga")
'        VALIPUVEN% = 0
'        GoTo 99
'     End If
   End If
   '
   If ESEXCEPCION_EYM% > 0 Then GoTo 99
   
   If POXIVA% > 0 And POXIVA% <> 5 And POXIVA% <> 7 Then ' si no es consumidor final
     FIN& = ULTREG&("!CARDEFAC")
     For UI& = 1 To FIN&
       XX$ = REGLEIDO$("!CARDEFAC", UI&)
       CI% = CVI(Mid$(XX$, 5, 2))
       If BIENCAP%(CI%) <> TIPODEPROD% Then
            Call MENSERR(24, "El Artículo " & CODEXT$(CI%) & " \No Puede Facturarse Por el Punto de Venta Activo " & TIPROX$)
            VALIPUVEN% = 0
            GoTo 99
       End If
    Next UI&
   End If
   '
99 VALIDAPUNVEN% = VALIPUVEN%
  
End Function

Sub CARNUMSER(OK%)
   '
   OK% = 1
   REF2$ = ""
   FIN& = ULTREG&("!CARDEFAC")
'   FORNUMSE.GRID.Rows = FIN& + 1
   For U& = 1 To FIN&
      Y$ = REGLEIDO$("!CARDEFAC", U&)
      CAN# = CVS(Mid$(Y$, 1, 4))
      CI% = CVI(Mid$(Y$, 5, 2))
      If ESTRAZABLE%(CI%) = 2 Then
        If CAN# > 1 Then ' VALIDA QUE LA CANTIDAD NO SEA MAYOR A 1 PARA INGRESAR EL NUMERO DE SERIE
            Call MENSERR(24, "Cantidad No Valida - El Máximo Permitido es 1 (uno)\en los Articulos Trazables")
            OK% = 2
            Exit Sub
        End If
      Else
        REF2$ = "S/N"
        Call REPLA(Y$, REF2$, 139, 64)
        Call GRAREG("!CARDEFAC", Y$, U&)
      End If
   Next U&
   '
10 With FORNUMSE
        .Visible = False
        .Caption = "CARGA DE NUMEROS DE SERIE"
        .GRID.Clear
        .GRID.Row = 0

        .GRID.RowHeight(0) = 500
        .GRID.COL = 0: .GRID.TEXT = "Codigo"
        .GRID.COL = 1: .GRID.TEXT = "Descripcion"
        .GRID.COL = 2: .GRID.TEXT = "Nro. Serie": .GRID.CellAlignment = 4: .GRID.ColAlignment(2) = 1
        
        .GRID.ColWidth(0) = 1600
        .GRID.ColWidth(1) = 5000
        .GRID.ColWidth(2) = 2200
        For II& = 1 To FIN&
            Y$ = REGLEIDO$("!CARDEFAC", II&)
            '
            If II& + 1 > .GRID.Rows Then .GRID.Rows = II& + 1
            .GRID.Row = II&: .GRID.COL = 0
            CI% = CVI(Mid$(Y$, 5, 2)): .GRID.TEXT = CODEXT$(CI%)
            .GRID.Row = II&: .GRID.COL = 1
            .GRID.TEXT = DESCRIT0$(CI%)
            .GRID.Row = II&: .GRID.COL = 2
            .GRID.TEXT = TRIM$(Mid$(Y$, 139, 64))
        Next II&
        .Refresh
        .GRID.Row = 1
        .GRID.COL = 2
        Call FORNUMSE.GRID_Click
   End With
   '
   FORNUMSE.Show 1
   RSP$ = FORNUMSE.RESPUESTA
   Unload FORNUMSE
   '
   If RSP$ = "NOOK" Then
     OK% = 0
     Exit Sub
   Else
     FIN& = ULTREG&("!CARDEFAC")
     For U& = 1 To FIN&
       Y$ = REGLEIDO$("!CARDEFAC", U&)
       NUMSE$ = TRIM$(Mid$(Y$, 139, 64))
       If NUMSE$ <> "S/N" Then
         CI% = CVI(Mid$(Y$, 5, 2))
         '
         Condi1$ = "NumeroSerie='" & NUMSE$ & "'"
         '
         REGI1$ = VALOBADA("TANUMSE", "NumeroSerie", Condi1$)
         If REGI1$ = "" Then
           Call MENSERR(24, "Numero de Serie " & NUMSE$ & " Inexistente")
           GoTo 10
         End If
         '
         REGI2% = XVALO(VALOBADA("TANUMSE", "Cod_Inte", Condi1$))
         If REGI2% <> CI% Then
           Call MENSERR(24, "Numero de Serie " & NUMSE$ & " Corresponde al Código: " & CODEXT$(REGI2%) & "")
           GoTo 10
         End If
         '
         REGI3$ = TRIM$(VALOBADA("TANUMSE", "NumRemito", Condi1$))
         If TIFADOC% <= 2 And REGI3$ <> "" Then ' solo en caso de facturas
              Call MENSERR(24, "Numero de Serie " & NUMSE$ & " ya Entregado (Nro Remito: " & REGI3$ & ")")
              GoTo 10
         ElseIf TIFADOC% = 6 And REGI3$ = "" Then ' solo en caso de notas de credito
             Call MENSERR(24, "Numero de Serie " & NUMSE$ & " No Se Encuentra Entregado")
             GoTo 10
         End If
       End If
     Next U&
   End If
   '
End Sub
Function VERIPEDMO%()
   'VALIDA QUE NO HAYAN PEDIDOS CON DIFERENTES MOTIVOS
   ' Y SI LOS PEDIDOS SELECCIONADOS ESTAN AUTORIZADOS
   
   '0 = NO
   '1 = OK
   '
   VERI% = 1
   '
   'RECORRO LOS PEDIDOS Y VALIDO SI TODOS TIENEN EL MISMO MOTIVO
   For KKU& = 1 To ULTREG&("!CARDEFAC")
        XX$ = REGLEIDO$("!CARDEFAC", KKU&)
        NPED& = XVALO(TRIM$(Mid$(XX$, 371, 16)))
        If KKU& = 1 Then
          NUMOT1% = NUMOTIVO%(NPED&)
        End If
        NUMOT2% = NUMOTIVO%(NPED&)
        If NUMOT2% <> NUMOT1% Then
           'Call COMUNI("Se Han Detectado Pedidos con Diferente Motivo Asignado")
           VERI% = 0
           GoTo 99
        End If
    Next KKU&
    '
    'RECORRO LOS PEDIDOS Y VALIDO SI TODOS ESTAN AUTORIZADOS
    For KKU& = 1 To ULTREG&("!CARDEFAC")
        XX$ = REGLEIDO$("!CARDEFAC", KKU&)
        NPED& = XVALO(TRIM$(Mid$(XX$, 371, 16)))
        If VALIDAOV%(NPED&) < 1 Then
           VERI% = 0
           GoTo 99
        End If
    Next KKU&
    '
99 VERIPEDMO% = VERI%
End Function
Function LISTADEMPAQ%(NC%)
    '
    LISEMPAQ% = 0
    CONDI$ = "CODI_CLI = " & NC% & " AND EmpaCerra > 0 and (NumFactura = '' or NumFactura is Null)"
    HAYREG& = CantRegistros("ENCEMPAQ", CONDI$)
    If HAYREG& > 0 Then
       LISEMPAQ% = 1
    End If
    '
    LISTADEMPAQ% = LISEMPAQ%
End Function


'LA RUTINA SE PASO AL MODULO MOVIST07_SQL
'Sub ASIGNALOTES()
'   'SELECCION DE LOTES A ASIGNAR AL PEDIDO
'   REGIS% = 0
'   ASIGLOT07.List1.Clear
'   ASIGLOT07.List2.Clear
'   For U& = 1 To ULTREG&("!CARDEFAC")
'       XX$ = REGLEIDO$("!CARDEFAC", U&)
'       CCII% = CVI(Mid$(XX$, 5, 2))
'       If CCII% > 0 And ESTRAZABLE%(CCII%) = 1 Then
'         TXTX$ = ""
'         Call REPLA(TXTX$, CODEXT$(CCII%), 1, 16)
'         Call REPLA(TXTX$, DESCRIT$(CCII%), 18, 33)
'         SALPEND = CVS(Mid$(XX$, 1, 4))
'         Call REPLA(TXTX$, FORMATNUM$(SALPEND, "F10.1"), 54, 12)
'         Call REPLA(TXTX$, TRIM$(FORMATNUM$(U&, "I5")), 80, 5)
'         ASIGLOT07.List1.AddItem (TXTX$)
'         REGIS% = REGIS% + 1
'       End If
'   Next U&
'   If REGIS% > 0 Then
'      ASIGLOT07.Show 1
'   End If
'   '
'End Sub
'
'LA RUTINA SE PASO AL MODULO MOVIST07_SQL
'Sub DETLOTREM(CI%, REG&)
'    '
'    If ESTRAZABLE%(CI%) = 1 Then
'       TOTALOTE$ = ""
'       For UY& = 0 To ASIGLOT07.List1.ListCount - 1
'           YY$ = ASIGLOT07.List1.List(UY&)
'           If XVALO(Mid$(YY$, 80, 5)) = REG& Then
'               TOTALOTE$ = TRIM$(Mid$(YY$, 85))
'           End If
'       Next UY&
'       '
'5      NN% = InStr(TOTALOTE$, ";")
'       If NN% > 0 Then
'           LOTTE$ = Mid$(TOTALOTE$, 1, NN% - 1)
'           '
'           CONDI$ = "IDENLOTE = '" & LOTTE$ & "'"
'           FEALT$ = (VALOBADA("LOTINGRE", "FECHALTA", CONDI$, "NOMESS"))
'           FEEXX = CVINT(FEALT$)
'           FEE$ = FECHATEX$(FEEXX)
'           ANCHOO = XVALO(VALOBADA("LOTINGRE", "ANCHOMILI", CONDI$, "NOMESS"))
'           CANTI = XVALO(VALOBADA("LOTINGRE", "CANALTA", CONDI$, "NOMESS"))
'           '
'           LOTTE$ = AJUSTI$(LOTTE$, 12) & " - " & TRIM$(FORMATNUM$(CANTI, "F10.1")) & " " & UNIMED$(CI%)
'           If ANCHOO > 0 Then
'             LOTTE$ = LOTTE$ & " - " & TRIM$(FORMATNUM$(ANCHOO, "F10.1")) & " mm"
'           End If
'           '
'           LOTTE$ = LOTTE$ & " - " & FEE$
'           Call AGREGATEXTO(LOTTE$, 2)
'           TOTALOTE$ = Mid$(TOTALOTE$, NN% + 1)
'           GoTo 5
'       End If
'    End If
'    '
'End Sub

Sub SOLICITARCAE_WSFEX1()
    Flag% = 0
    CONT% = 0
    '
    ' AQUI RECALCULA EL IMPORTE TOTAL
    TONEFA# = 0
    For i& = 1 To ULTREG&("!CARDEFAC")
      XX$ = REGLEIDO$("!CARDEFAC", i&)
      CI% = CVI(Mid$(XX$, 5, 2))
      DESCRI$ = DESCRIT$(CI%)
      If CI% = 0 Then DESCRI$ = TRIM$(Mid$(XX$, 7, 64))
      CAN = CVS(Mid$(XX$, 1, 4))
      PRECI# = CVS(Mid$(XX$, 429, 4))
      PORDESCUEN = PORCEDESCU(FORFACPRINT.PORDESCU) 'MKS$(Mid$(XX$, 425, 4))
      PRECI# = ROUND(PRECI# * (100 - PORDESCUEN) / 100, 4)
      TOTIT# = ROUND(CDbl(ROUND(CAN, 1) * PRECI#), 2)
      TONEFA# = TONEFA# + TOTIT#
    Next i&
    'GASTOS DE EXPORTACIÓN
    GASEXPOR# = XVALO(FINFACTUR07.Text2(15))
    If GASEXPOR# > 0 Then TONEFA# = TONEFA# + GASEXPOR#
    '
    TONEFAX$ = TRIM$(FORMATNUM$(TONEFA#, "F15.2"))
    '
    '
    NCLI = XVALO(Text1)
    'TOMO LOS DATOS DEL FRM DE INGRESO DE DATOS
    
    TIPO_EXPO$ = Left$(TRIM$(DATFAEX07.Combo1.TEXT), 1)
    DEST_COMP$ = Left$(TRIM$(DATFAEX07.Text2.TEXT), 3)
    CUIT_PAIS$ = Left$(TRIM$(DATFAEX07.Text3.TEXT), 11)
    If DEST_COMP$ = "250" Then CUIT_PAIS$ = TRIM$(REPLACAR$(CuitCli$(NCLI), "-", ""))
    INCOTS$ = TRIM$(DATFAEX07.Combo5.TEXT)
    IDIOMA$ = TRIM$(Left(DATFAEX07.Combo6.TEXT, 1))
    COMP_ASOCIAD% = DATFAEX07.List1.ListCount
    S_N$ = "N" 'VARIABLE QUE DEFINE SI TIENE PERMIOS O NO
    
    PERMISO% = DATFAEX07.List2.ListCount: If PERMISO% > 0 Then S_N$ = "S"
    ' EN EL CASO QUE SEA NOTA DE CREDITO O DEBITO ESTE DEBE ESTAR VACIO
    If FORFACPRINT.Label7.Caption = "N.Débito:" Then S_N$ = "" 'faltaban los ':'
    If FORFACPRINT.Label7.Caption = "N.Crédito:" Then S_N$ = ""
    ' EN EL CASO QUE EL TIPO DE EXPORTACION NO SEA DE BIENES ESTE DEBE ESTAR VACIO
    If TIPO_EXPO$ <> "1" Then S_N$ = ""
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    If EXISTE%(CARPETOK$ + "token_WSFEX.txt") = 0 Then
      TXT$ = REPLACAR$(CARPETOK$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'token_WSFEX.txt'\en Carpeta " & TXT$)
      Exit Sub
    End If
    If EXISTE%(LUDAT$ + "consulta_cae_WSFEX_original.txt") = 0 Then
      TXT$ = REPLACAR$(LUDAT$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'consulta_cae_WSFEX_original.txt'\en Carpeta " & TXT$)
      Exit Sub
    End If
    NARCHI1% = FILEOPEN%(LUDAT$ + "consulta_cae_WSFEX_original.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    NARCHI2% = FILEOPEN%(CARPETOK$ + "consulta_cae_WSFEX.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    NCLI = XVALO(FORFACPRINT.Text1)
    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       CONT% = CONT% + 1
       Pos% = InStr(XX$, ">")
       ETIQUETACIERRE$ = TRIM$(Mid$(XX$, Pos% + 1))
       TX$ = Mid$(XX$, 1, Pos%)
       
          Select Case CONT%
           Case 2
            FECHA$ = FECHATEXLAR$(FECHANUM(FORFACPRINT.Text9))
            DIA$ = Mid(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 4)
            Dato$ = ANIO$ + MES$ + DIA$
           Case 3: Dato$ = TIPOX_CBTEX$ ' TIPO DE COMPROBANTE
           Case 4: Dato$ = TRIM$(Str(XVALO(FORFACPRINT.Label12))) ' Punto de Venta
           Case 5: Dato$ = TRIM$(FORFACPRINT.Text11)            ' Numero de comprobante inicial
           Case 6: Dato$ = TIPO_EXPO$ ' TIPO DE EXPORTACION
           Case 7: Dato$ = S_N$ ' PERMISO EXISTENTE
           Case 9:
           If S_N$ = "S" Then
              For J% = 0 To DATFAEX07.List2.ListCount - 1
                PERMI$ = DATFAEX07.List2.List(J%)
                Print #NARCHI2%, "<permiso>"
                Print #NARCHI2%, "<id_permiso>" & TRIM$(PERMI$) & "</id_permiso>"
                Print #NARCHI2%, "<dst_merc>" & TRIM$(DEST_COMP$) & "</dst_merc>"
                Print #NARCHI2%, "</permiso>"
              Next J%
           End If
           Case 10: Dato$ = DEST_COMP$ ' DESTINO PAIS
           Case 11: Dato$ = TRIM$(rasocli$(NCLI))
           Case 12: Dato$ = CUIT_PAIS$ ' CUIT DE PAIS
           Case 13: Dato$ = TRIM$(DOMICLI$(NCLI)) 'DOMICILIO
           Case 14: Dato$ = TRIM$(REPLACAR$(CuitCli$(NCLI), "-", "")) 'CUIT
           Case 15: Dato$ = MONEDA_CAE$ ' TIPO DE MONEDA
           Case 16: Dato$ = TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) 'COTIZACION
           Case 17: Dato$ = Left$(Text4, 2000) ' OBSERVACIONES
           ' Case 18: DATO$ = TRIM$(FORFACPRINT.Total(6)) ' Importe Total
           Case 18: Dato$ = TRIM$(TONEFAX$)               ' Importe Total
           Case 19: Dato$ = "cte(" & Text1 & ") " & Left$(Text4, 2000) ' OBSERVACIONES
           Case 21:
              If COMP_ASOCIAD% > 0 Then
                For J% = 0 To DATFAEX07.List1.ListCount - 1
                 RELEI$ = DATFAEX07.List1.List(J%)
                 CBTE$ = Mid$(TRIM$(RELEI$), 1, 2)
                 PVTA$ = Mid$(RELEI$, 20, 4)
                 NURO$ = Mid$(RELEI$, 32, 8)
                 Print #NARCHI2%, "<cmp_asoc>"
                 Print #NARCHI2%, "<cbte_tipo>" & TRIM$(CBTE$) & "</cbte_tipo>"
                 Print #NARCHI2%, "<cbte_punto_vta>" & TRIM$(PVTA$) & "</cbte_punto_vta>"
                 Print #NARCHI2%, "<cbte_nro>" & TRIM$(NURO$) & "</cbte_nro>"
                 Print #NARCHI2%, "</cmp_asoc>"
                Next J%
              End If
           Case 22: Dato$ = TRIM$(Left$(FORFACPRINT.Eco21(1), 50)) ' FORMA DE PAGO
           Case 23: Dato$ = INCOTS$ ' TIPO DE VENTA FOB,CIF,ETC
           FORFACPRINT.Text18 = INCOTS$ ' ACTUALIZO EL FRM
           Case 25: Dato$ = IDIOMA$ '1-ESPAÑOL 2-INGLES 3-PORTUGUES
           
          End Select
        TX$ = TX$ + Dato$ + ETIQUETACIERRE$ ' CONCATENA EL RENGLON A IMPRIMIR
        Dato$ = ""
        Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
     Loop
    '
    Print #NARCHI2%, "<Items>"

    descrilar$ = ""
    For i& = 1 To ULTREG&("!CARDEFAC")
      XX$ = REGLEIDO$("!CARDEFAC", i&)
      CI% = CVI(Mid$(XX$, 5, 2))
      DESCRI$ = DESCRIT$(CI%)
      If CI% = 0 Then DESCRI$ = TRIM$(Mid$(XX$, 7, 64))
      CAN = CVS(Mid$(XX$, 1, 4))
      PRECI# = CVS(Mid$(XX$, 429, 4))
      '
      PORDESCUEN = PORCEDESCU(FORFACPRINT.PORDESCU) 'MKS$(Mid$(XX$, 425, 4))
      '
      PRECI# = ROUND(PRECI# * (100 - PORDESCUEN) / 100, 3)
      PRCI$ = FORMATNUM$(PRECI#, "F12.3")
      IMPORTE# = ROUND(CDbl(ROUND(CAN, 1) * PRECI#), 2)
      IMPRT$ = FORMATNUM$(IMPORTE#, "F12.2")
      CONDI$ = "Codint = " & CI%
      NOMENC$ = VALOBADA("MASTER", "CODIGO_NCM", CONDI$, "NOMESS")
      UNIMEX$ = TRIM$(VALOBADA("MASTER", "PRO_UMED", CONDI$, "NOMESS")) 'REVISA SI ESTA EN BASE DE DATOS
      If UNIMEX$ = "" Or CI% = 0 Then UNIMEX$ = Str$(UNIMEAFIP%(CI%)) ' SI NO LO TRAE DEL ARCHIVO DE UNIMED
      '
      If PRECI# > 0.0005 Then
        Print #NARCHI2%, "<item>"
        Print #NARCHI2%, "<pro_codigo>" & TRIM$(NOMENC$) & "</pro_codigo>"
            If descrilar$ <> "" Then DESCRI$ = descrilar$ + DESCRI$
        Print #NARCHI2%, "<pro_ds>" & TRIM$(DESCRI$) & "</pro_ds>"
        Print #NARCHI2%, "<pro_qty>" & CAN & "</pro_qty>"
        Print #NARCHI2%, "<pro_umed>" & TRIM$(UNIMEX$) & "</pro_umed>"
        Print #NARCHI2%, "<pro_precio_uni>" & TRIM$(PRCI$) & "</pro_precio_uni>"
        Print #NARCHI2%, "<imp_total>" & TRIM$(IMPRT$) & "</imp_total>"
        Print #NARCHI2%, "</item>"
        descrilar$ = ""
      Else
        descrilar$ = descrilar$ + DESCRI$
      End If
    Next i&
    '
    'GASTOS DE EXPORTACION
    If GASEXPOR# > 0 Then
        Print #NARCHI2%, "<item>"
        REFER$ = Left$(FINFACTUR07.Text10, 32)
        Print #NARCHI2%, "<pro_codigo></pro_codigo>"
        Print #NARCHI2%, "<pro_ds>" & TRIM$(REFER$) & "</pro_ds>"
        Print #NARCHI2%, "<pro_qty>" & "1" & "</pro_qty>"
        Print #NARCHI2%, "<pro_umed>" & "7" & "</pro_umed>"
        Print #NARCHI2%, "<pro_precio_uni>" & TRIM$(FORMATNUM$(GASEXPOR#, "F12.3")) & "</pro_precio_uni>"
        Print #NARCHI2%, "<imp_total>" & TRIM$(FORMATNUM$(GASEXPOR#, "F12.2")) & "</imp_total>"
        Print #NARCHI2%, "</item>"
    End If
    '
    Print #NARCHI2%, "</Items>"
    Print #NARCHI2%, "</Cmp>"
    
    Close #NARCHI1%
    Close #NARCHI2%
    '
99
End Sub

Sub SOLICITARCAE_WSBFE1()
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    If EXISTE%(CARPETOK$ + "token_WSBFE.txt") = 0 Then
      TXT$ = REPLACAR$(CARPETOK$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'Token_WSBFE.txt'\en Carpeta " & TXT$)
      Exit Sub
    End If
    If EXISTE%(LUDAT$ + "consulta_cae_WSBFE_original.txt") = 0 Then
      TXT$ = REPLACAR$(LUDAT$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'consulta_cae_WSBFE_original.txt'\en Carpeta " & TXT$)
      Exit Sub
    End If
    NARCHI1% = FILEOPEN%(LUDAT$ + "consulta_cae_WSBFE_original.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    NARCHI2% = FILEOPEN%(CARPETOK$ + "consulta_cae_WSBFE.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA

    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       CONT% = CONT% + 1
       Pos% = InStr(XX$, ">")
       ETIQUETACIERRE$ = TRIM$(Mid$(XX$, Pos% + 1))
       TX$ = Mid$(XX$, 1, Pos%)
       
          Select Case CONT%
           
           Case 2: Dato$ = FINFACTUR07.TIPO_DOC      ' Tipo Documento (Cuit,Dni, etc) Identificador del comprador
           Case 3: Dato$ = FINFACTUR07.NRO_DOC      ' Numero Dni o Cuit
           Case 4: Dato$ = "1"   ' Numero de zona
           Case 5: Dato$ = FINFACTUR07.TIPO_CBTE$        ' Tipo de Comprobante(Factura A o B ,Nota de Credito,Nota de Debito)
           Case 6: Dato$ = TRIM$(Str(XVALO(FORFACPRINT.Label12))) ' Punto de Venta
           Case 7: Dato$ = TRIM$(FORFACPRINT.Text11)            ' Numero de comprobante inicial
           Case 8: Dato$ = TRIM$(FORFACPRINT.Total(6))          ' Importe Total
           Case 9: Dato$ = FINFACTUR07.IMP_TOT_CONC            ' Importe no gravados
           Case 10: Dato$ = FINFACTUR07.IMP_NETO                ' Importe Neto (IMPORTE NETO = IMPORTE NETO - IVA NO GRAVADO)
           Case 11: Dato$ = FINFACTUR07.IMPTO_LIQ            ' Impuesto Liquidado(la suma de los importes de iva  aplicados)
           Case 12: Dato$ = "0"           ' Impto_Liq_rni - siempre es cero  o percepcion a no categorizados
           Case 13: Dato$ = "0"           ' importe operaciones externas (No dieron informacion de que se trata)
           Case 14: Dato$ = FINFACTUR07.IMP_PERC
           Case 15: Dato$ = FINFACTUR07.IMP_IIBB
           Case 16: Dato$ = "0"
           Case 17: Dato$ = "0"
           Case 18: Dato$ = FINFACTUR07.MONEDA_CAE ' QUEDA PARA AGREGAR UNA FUNCION POR EL TIPO DE CAMBIO
           Case 19: Dato$ = TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4"))
           
           Case 20:           '1)Fecha Comprobante,2)Fecha de comienzo de Servicio, 3)Fecha Fin de Servicio, Fecha de Vencimento
                    FECHA$ = FECHATEXLAR$(FECHANUM(FORFACPRINT.Text9))
                    DIA$ = Mid(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 4)
                    Dato$ = ANIO$ + MES$ + DIA$
           Case 21:
                     If TIFADOC% = 2 And chkModoCredito.Value = 1 Then
                        FECHA$ = FECHATEXLAR$(FECHANUM(txtFechaVencimiento))
                        DIA$ = Mid(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 4)
                        Dato$ = ANIO$ + MES$ + DIA$
                     End If
          End Select
        TX$ = TX$ + Dato$ + ETIQUETACIERRE$ ' CONCATENA EL RENGLON A IMPRIMIR
        Dato$ = ""
        Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
     Loop
    '
    
    ' OPCIONALES
    If TIFADOC% = 2 And chkModoCredito.Value = 1 Then
        TX$ = "<Opcionales>": Print #NARCHI2%, TX$
        
        TX$ = "<Opcional>": Print #NARCHI2%, TX$
        TX$ = "<Id>2101</Id>": Print #NARCHI2%, TX$
        TX$ = "<Valor>" & TRIM$(Control$("FACTURA", "CBUCREDI")) & "</Valor>": Print #NARCHI2%, TX$
        TX$ = "</Opcional>": Print #NARCHI2%, TX$
                
        TX$ = "</Opcionales>": Print #NARCHI2%, TX$

    ElseIf (TIFADOC% = 3 Or TIFADOC% = 6) And chkModoCredito.Value = 1 Then
        
        TX$ = "<Opcionales>": Print #NARCHI2%, TX$
        
        TX$ = "<Opcional>": Print #NARCHI2%, TX$
        TX$ = "<Id>22</Id>": Print #NARCHI2%, TX$
        TX$ = "<Valor>" & TipoAplicacionFacturaCredito$ & "</Valor>": Print #NARCHI2%, TX$
        TX$ = "</Opcional>": Print #NARCHI2%, TX$
                
        TX$ = "</Opcionales>": Print #NARCHI2%, TX$
    End If
    
    
    ' ITEMS
    Print #NARCHI2%, "<Items>"

    For i& = 1 To ULTREG&("!CARDEFAC")
      XX$ = REGLEIDO$("!CARDEFAC", i&)
      CI% = CVI(Mid$(XX$, 5, 2))
      DESCRI$ = DESCRIT$(CI%)
      CAN = CVS(Mid$(XX$, 1, 4))
      PRECI# = CVS(Mid$(XX$, 429, 4))
      '
      PORDESCUEN = PORCEDESCU(FORFACPRINT.PORDESCU) 'MKS$(Mid$(XX$, 425, 4))
      IMPDESC = (PRECI# * PORDESCUEN) / 100
      PRECI# = ROUND(PRECI# - IMPDESC)
      '
      PRCI$ = FORMATNUM$(PRECI#, "F12.2")
      IMPORTE# = CAN * PRECI#
      IMPRT$ = FORMATNUM$(IMPORTE#, "F12.2")
      CONDI$ = "Codint = " & CI%
      NOMENC$ = VALOBADA("MASTER", "CODIGO_NCM", CONDI$, "NOMESS")
      COSEC$ = VALOBADA("MASTER", "CODIGO_SEC", CONDI$, "NOMESS")
      UNIMEX$ = VALOBADA("MASTER", "PRO_UMED", CONDI$, "NOMESS")
      COTIVA$ = VALOBADA("MASTER", "IVA_ID", CONDI$, "NOMESS")
      '
      Print #NARCHI2%, "<item>"
      Print #NARCHI2%, "<pro_codigo_ncm>" & TRIM$(NOMENC$) & "</pro_codigo_ncm>"
      Print #NARCHI2%, "<pro_codigo_sec>" & TRIM$(COSEC$) & "</pro_codigo_sec>"
      Print #NARCHI2%, "<pro_ds>" & TRIM$(DESCRI$) & "</pro_ds>"
      Print #NARCHI2%, "<pro_qty>" & CAN & "</pro_qty>"
      Print #NARCHI2%, "<pro_umed>" & TRIM$(UNIMEX$) & "</pro_umed>"
      Print #NARCHI2%, "<pro_precio_uni>" & TRIM$(PRCI$) & "</pro_precio_uni>"
      Print #NARCHI2%, "<imp_bonif>0</imp_bonif>"
      Print #NARCHI2%, "<imp_total>" & TRIM$(IMPRT$) & "</imp_total>"
      Print #NARCHI2%, "<iva_id>" & TRIM$(COTIVA$) & "</iva_id>"
      Print #NARCHI2%, "</item>"
    Next i&
    Print #NARCHI2%, "</Items>"
    
    
    ' COMPROBANTES ASOCIADOS
    If TIFADOC% = 2 And chkModoCredito.Value = 1 Then
        NPVX$ = Label12
        EPAC$ = EMPREAC$
        If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "TANIT") > 0 Or InStr(EPAC$, "GABAL") > 0 Or InStr(EPAC$, "FERVI") > 0 Or InStr(EPAC$, "AHP") > 0 Or InStr(EPAC$, "J Y M") > 0 Or InStr(EPAC$, "CEYEQ") > 0 Or InStr(EPAC$, "EYM") > 0 Or InStr(EPAC$, "DAVTECNICA") > 0 Then NPVX$ = "0001"
        
        'OBTENGO CUIT DE EMPRESA ACTIVA
        n1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
        Dim X As String * 41
        Get #n1%, 6, X
        CUITT$ = TRIM$(REPLACAR$(DESENCRI$(X), "-", ""))
        Close #n1%
             
        If CONSICLI% > 0 Then
            ' FACTURA REMITOS EN CONSIGNACION
            TX$ = "<CbtesAsoc>": Print #NARCHI2%, TX$
            For IIX& = 1 To ULTREG&("!CARDEFAC")
            
                XX$ = REGLEIDO$("!CARDEFAC", IIX&)
                REMITO$ = TRIM$(Mid$(XX$, 387, 16))
                
                If InStr(1, AgrupaRemitos$, REMITO$) > 0 Then GoTo 55
                
                AgrupaRemitos$ = AgrupaRemitos$ & REMITO$
                
                TX$ = "<CbteAsoc>": Print #NARCHI2%, TX$
                TX$ = "<Tipo_cbte>91</Tipo_cbte>": Print #NARCHI2%, TX$
                TX$ = "<Punto_vta>" & CStr(Val(NPVX$)) & "</Punto_vta>": Print #NARCHI2%, TX$
                NROREMITO$ = TRIM$(Mid$(REMITO$, 4))
                TX$ = "<Cbte_nro>" & NROREMITO$ & "</Cbte_nro>": Print #NARCHI2%, TX$
                TX$ = "<Cuit>" & CUITT$ & "</Cuit>": Print #NARCHI2%, TX$
                FechaRemito% = CVINT(VALOBADA("movisto", "fechmov", "dopricor='" & REMITO$ & "'"))
                FechaRemitoUSA$ = Format(CVDAT(FechaRemito%), "yyyymmdd")
                TX$ = "<Fecha_cbte>" & FechaRemitoUSA$ & "</Fecha_cbte>": Print #NARCHI2%, TX$
                TX$ = "</CbteAsoc>": Print #NARCHI2%, TX$
                
55          Next IIX&
            TX$ = "</CbtesAsoc>": Print #NARCHI2%, TX$
        Else
            COPIASREM% = XVALO(FINFACTUR07.Text2(1))
            If COPIASREM% > 0 Then
                ' EMITE JUNTOS FACTURA Y REMITO
                TX$ = "<CbtesAsoc>": Print #NARCHI2%, TX$
                
                TX$ = "<CbteAsoc>": Print #NARCHI2%, TX$
                TX$ = "<Tipo>91</Tipo>": Print #NARCHI2%, TX$
                TX$ = "<PtoVta>" & CStr(Val(NPVX$)) & "</PtoVta>": Print #NARCHI2%, TX$
                NROREMITO$ = TRIM$(Text10)
                TX$ = "<Nro>" & NROREMITO$ & "</Nro>": Print #NARCHI2%, TX$
                TX$ = "<Cuit>" & CUITT$ & "</Cuit>": Print #NARCHI2%, TX$
                FechaRemito% = HOY(HO$)
                FechaRemitoUSA$ = Format(CVDAT(FechaRemito%), "yyyymmdd")
                TX$ = "<CbteFch>" & FechaRemitoUSA$ & "</CbteFch>": Print #NARCHI2%, TX$
                TX$ = "</CbteAsoc>": Print #NARCHI2%, TX$
                
                TX$ = "</CbtesAsoc>": Print #NARCHI2%, TX$
            End If
        End If
    
    ElseIf (TIFADOC% = 3 Or TIFADOC% = 6) And chkModoCredito.Value = 1 Then
           
        'OBTENGO CUIT DE EMPRESA ACTIVA
        n1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
        Dim regi As String * 41
        Get #n1%, 6, regi
        CUITT$ = TRIM$(REPLACAR$(DESENCRI$(regi), "-", ""))
        Close #n1%
             
        ' APLICA A FACTURA
        PuntoVenta$ = Mid$(AplicacionFacturaCredito$, 4, 4)
        NumeroFactura$ = Mid$(AplicacionFacturaCredito$, 9, 8)
        Tipo$ = Mid$(AplicacionFacturaCredito$, 1, 2)
        If Tipo$ = "FC" Then TipoComprobanteAfip$ = "201"
        If Tipo$ = "ND" Then TipoComprobanteAfip$ = "202"
        If Tipo$ = "NC" Then TipoComprobanteAfip$ = "203"
        TX$ = "<CbtesAsoc>": Print #NARCHI2%, TX$
        
        TX$ = "<CbteAsoc>": Print #NARCHI2%, TX$
        TX$ = "<Tipo>" & TipoComprobanteAfip$ & "</Tipo>": Print #NARCHI2%, TX$
        TX$ = "<PtoVta>" & CStr(Val(PuntoVenta$)) & "</PtoVta>": Print #NARCHI2%, TX$
        TX$ = "<Nro>" & CStr(Val(NumeroFactura$)) & "</Nro>": Print #NARCHI2%, TX$
        TX$ = "<Cuit>" & CUITT$ & "</Cuit>": Print #NARCHI2%, TX$
        FechaFactura% = CVINT(VALOBADA("cajabanc", "fechconta", "nuclien=" & Text1 & " and codicom_lar='" & AplicacionFacturaCredito$ & "' and tipomovim='fven'"))
        FechaFacturaUSA$ = Format(CVDAT(FechaFactura%), "yyyymmdd")
        TX$ = "<CbteFch>" & FechaFacturaUSA$ & "</CbteFch>": Print #NARCHI2%, TX$
        TX$ = "</CbteAsoc>": Print #NARCHI2%, TX$
    End If
    
    Print #NARCHI2%, "</Cmp>"
    
    Close #NARCHI1%
    Close #NARCHI2%
    '
99 End Sub
Function MONEDA_CAE$()
    '
    MONE$ = ""
    If MONEMI% = 1 Then
         MONE$ = "PES"
      ElseIf MONEMI% = 2 Then
         MONE$ = "DOL"
      Else
         DEMONE$ = UCase(DEMONECO$(MONEMI%))
         If Left$(DEMONE$, 4) = "EURO" Then
              MONE$ = "060"
           ElseIf Left$(DEMONE$, 4) = "REAL" Then
              MONE$ = "012"
           ElseIf Left$(DEMONE$, 4) = "DOLA" Then
              MONE$ = "DOL"
         End If
    End If
    '
99  MONEDA_CAE$ = MONE$
    '
End Function

Function TIPOX_CBTEX$()

    TIPOCOMP$ = UCase$(FORFACPRINT.Label7 + FORFACPRINT.Label13)
    
    Select Case TIPOCOMP$
        Case "FACTURA:E": valor$ = "19"
        Case "N.DÉBITO:E": valor$ = "20"
        Case "N.CRÉDITO:E": valor$ = "21"
    End Select
    '
    TIPOX_CBTEX$ = valor$
    '
End Function

Sub SOLICITARCAE1()
    Flag% = 0
    CONT% = 0
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    If EXISTE%(CARPETOK$ + "token.txt") = 0 Then
      TXT$ = REPLACAR$(CARPETOK$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'Token.txt'\en Carpeta " & TXT$)
      GoTo 99
    End If
    If EXISTE%(LUDAT$ + "consulta_cae_original.txt") = 0 Then
      TXT$ = REPLACAR$(LUDAT$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'consulta_cae_original.txt'\en Carpeta " & TXT$)
      GoTo 99
    End If
    '
    'TIPO DE CAMBIO EL QUE CONFIGURA EN PANTALLLA
    TIPOCAM# = TICAMBX#

    NARCHI1% = FILEOPEN%(LUDAT$ + "consulta_cae_original.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    NARCHI2% = FILEOPEN%(CARPETOK$ + "consulta_cae.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    
    'SE CALCULA POR FUERA PORQUE EN EL CASO DE FACTURAS EN DOLARES EL TOTAL NO COINCIDIA CON LA SUMA DEL NETO, IVA E IIBB
    DATO16# = ROUND#(XVALO(FINFACTUR07.IMP_TOT_CONC$) * TIPOCAM#, 2)
    DATO17# = ROUND#(XVALO(FINFACTUR07.IMP_NETO$) * TIPOCAM#, 2)
    DATO18# = ROUND#(XVALO(FINFACTUR07.IMPTO_LIQ$) * TIPOCAM#, 2)
    DATO15# = ROUND#(DATO16# + DATO17# + DATO18#, 2)

    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       CONT% = CONT% + 1
       Pos% = InStr(XX$, ">")
       ETIQUETACIERRE$ = TRIM$(Mid$(XX$, Pos% + 1))
       TX$ = Mid$(XX$, 1, Pos%)
       
          Select Case CONT%
           'Case 4: DATO$ = TRIM$(FORFACPRINT.Label18)          ' Cantidad de registros del detalle de la factura
           Case 4: Dato$ = "1"            ' corregido porque debe ser siempre 1
           Case 5: Dato$ = "1"            ' Tipo de Presentacion dada (0 servicios , 1 Bienes de Capital)
           Case 9: Dato$ = FINFACTUR07.TIPO_DOC$      ' Tipo Documento (Cuit,Dni, etc) Identificador del comprador
           Case 10: Dato$ = FINFACTUR07.NRO_DOC$      ' Numero Dni o Cuit
           Case 11: Dato$ = FINFACTUR07.TIPO_CBTE$        ' Tipo de Comprobante(Factura A o B ,Nota de Credito,Nota de Debito)
           Case 12: Dato$ = TRIM$(Str(XVALO(FORFACPRINT.Label12))) ' Punto de Venta
           Case 13: Dato$ = "0"                                  ' Numero de comprobante inicial
           Case 14: Dato$ = "0"                                  ' Numero de comprobante hasta
           'CALCULO DE IMPORTE POR TIPO DE CAMBIO
           Case 15: Dato$ = TRIM$(FORMATNUM$(DATO15#, "F12.2"))           ' Importe Total
           Case 16: Dato$ = TRIM$(FORMATNUM$(DATO16#, "F12.2"))            ' Importe no gravados = C/Iva, entre ellos neto no gravado, las percepciones de iva e IIBB
                    If Dato$ = "" Then Dato$ = "0" 'SE AGREGA PORQUE CUANDO ES CERO EL FORMATNUM DEVUELVE VACIO
           Case 17: Dato$ = TRIM$(FORMATNUM$(DATO17#, "F12.2"))                ' Importe Neto (IMPORTE NETO = IMPORTE NETO - IVA NO GRAVADO)
                    If Dato$ = "" Then Dato$ = "0" 'SE AGREGA PORQUE CUANDO ES CERO EL FORMATNUM DEVUELVE VACIO
           Case 18: Dato$ = TRIM$(FORMATNUM$(DATO18#, "F12.2"))            ' Impuesto Liquidado(la suma de los importes de iva  aplicados)
                    If Dato$ = "" Then Dato$ = "0" 'SE AGREGA PORQUE CUANDO ES CERO EL FORMATNUM DEVUELVE VACIO
           Case 19: Dato$ = "0"           ' Impto_Liq_rni - siempre es cero  o percepcion a no categorizados
           Case 20: Dato$ = "0"           ' importe operaciones externas (No dieron informacion de que se trata)
           Case 21, 22, 23, 24:           '1)Fecha Comprobante,2)Fecha de comienzo de Servicio, 3)Fecha Fin de Servicio, Fecha de Vencimento
                    FECHA$ = FECHATEXLAR$(FECHANUM(FORFACPRINT.Text9))
                    DIA$ = Mid(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 4)
                    Dato$ = ANIO$ + MES$ + DIA$
          End Select
          ' NOTA !!!!************************************************
          ' EN LA FILA 24 SE DEBE PONER LA FECHA DE VENCIMIENTO DEL COMPROBANTE
          ' PROVISORIAMENTE SE PUSO LA MISMA DEL DIA
          '*********************************************************************
        TX$ = TX$ + Dato$ + ETIQUETACIERRE$ ' CONCATENA EL RENGLON A IMPRIMIR
        Dato$ = ""
        Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
     Loop
    '
    Close #NARCHI1%
    Close #NARCHI2%
    '
99 End Sub

Sub SOLICITARCAE1_V1()
    Flag% = 0
    CONT% = 0
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    If EXISTE%(CARPETOK$ + "token.txt") = 0 Then
      TXT$ = REPLACAR$(CARPETOK$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'Token.txt'\en Carpeta " & TXT$)
      GoTo 99
    End If
    If EXISTE%(LUDAT$ + "consulta_cae_v1_original.txt") = 0 Then
      TXT$ = REPLACAR$(LUDAT$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'consulta_cae_v1_original.txt'\en Carpeta " & TXT$)
      GoTo 99
    End If
    '
    'TIPO DE CAMBIO EL QUE CONFIGURA EN PANTALLLA
    TIPOCAM# = TICAMBX#
    
    On Error GoTo ControlError
    Call Kill(CARPETOK$ + "consulta_cae_v1.txt")
    On Error GoTo 0
    
    NARCHI1% = FILEOPEN%(LUDAT$ + "consulta_cae_v1_original.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    NARCHI2% = FILEOPEN%(CARPETOK$ + "consulta_cae_v1.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    
    'SE CALCULA POR FUERA PORQUE EN EL CASO DE FACTURAS EN DOLARES EL TOTAL NO COINCIDIA CON LA SUMA DEL NETO, IVA E IIBB
    DATO16# = XVALO(FORFASQL.TLABEL28)                            ' NO GRAVADO
    DATO17# = XVALO(FORFASQL.Total(2)) - XVALO(FORFASQL.TLABEL28) ' NETO
    DATO18# = XVALO(FORFASQL.Total(3)) + XVALO(FORFASQL.Total(4)) + XVALO(FORFASQL.Total(5))    ' IVA (TODOS)
    DATO19# = XVALO(FORFASQL.Text26) + XVALO(FORFASQL.Text25) + XVALO(FORFASQL.Text24) + XVALO(FORFASQL.Text15) + XVALO(FORFASQL.Text17) ' TRIBUTOS (PERCEPCIONES)
    DATO15# = XVALO(Total(6))

    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       CONT% = CONT% + 1
       Pos% = InStr(XX$, ">")
       ETIQUETACIERRE$ = TRIM$(Mid$(XX$, Pos% + 1))
       TX$ = Mid$(XX$, 1, Pos%)
       
          Select Case CONT%
          
            Case 4: Dato$ = "1"            ' corregido porque debe ser siempre 1
            Case 5: Dato$ = "1"            ' Tipo de Presentacion dada (0 servicios , 1 Bienes de Capital)
            Case 9: Dato$ = FINFACTUR07.TIPO_DOC$      ' Tipo Documento (Cuit,Dni, etc) Identificador del comprador
            Case 10: Dato$ = FINFACTUR07.NRO_DOC$      ' Numero Dni o Cuit
            Case 11: Dato$ = FINFACTUR07.TIPO_CBTE$        ' Tipo de Comprobante(Factura A o B ,Nota de Credito,Nota de Debito)
            Case 12: Dato$ = TRIM$(Str(XVALO(FORFACPRINT.Label12))) ' Punto de Venta
            Case 13: Dato$ = "0"                                    ' Numero de comprobante inicial
            Case 14: Dato$ = "0"                                    ' Numero de comprobante hasta
            'CALCULO DE IMPORTE POR TIPO DE CAMBIO
            Case 15: Dato$ = TRIM$(FORMATNUM$(DATO15#, "F12.2"))            ' Importe Total
            Case 16: Dato$ = TRIM$(FORMATNUM$(DATO16#, "F12.2"))            ' Importe no gravado
                     If Dato$ = "" Then Dato$ = "0" 'SE AGREGA PORQUE CUANDO ES CERO EL FORMATNUM DEVUELVE VACIO
            Case 17: Dato$ = TRIM$(FORMATNUM$(DATO17#, "F12.2"))            ' Importe Neto (IMPORTE NETO = IMPORTE NETO - IVA NO GRAVADO)
                     If Dato$ = "" Then Dato$ = "0" 'SE AGREGA PORQUE CUANDO ES CERO EL FORMATNUM DEVUELVE VACIO
            Case 18: Dato$ = TRIM$(FORMATNUM$(DATO18#, "F12.2"))            ' Impuesto Liquidado(la suma de los importes de iva  aplicados)
                     If Dato$ = "" Then Dato$ = "0" 'SE AGREGA PORQUE CUANDO ES CERO EL FORMATNUM DEVUELVE VACIO
            Case 19: Dato$ = TRIM$(FORMATNUM$(DATO19#, "F12.2"))            ' Tributos (Percepciones)
                     If Dato$ = "" Then Dato$ = "0" 'SE AGREGA PORQUE CUANDO ES CERO EL FORMATNUM DEVUELVE VACIO
                     
            Case 20: Dato$ = "0"           ' importe operaciones externas (No dieron informacion de que se trata)
            Case 21:           ' 1)Fecha Comprobante, 2)Fecha de Vencimento
                     FECHA$ = FECHATEXLAR$(FECHANUM(FORFACPRINT.Text9))
                     DIA$ = Mid(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 4)
                     Dato$ = ANIO$ + MES$ + DIA$
            
            Case 22, 23:           ' 1)Fecha de comienzo de Servicio, 2)Fecha Fin de Servicio
                     Dato$ = ""
            Case 24:
                     If TIFADOC% = 2 And chkModoCredito.Value = 1 Then
                        FECHA$ = FECHATEXLAR$(FECHANUM(txtFechaVencimiento))
                        DIA$ = Mid(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 4)
                        Dato$ = ANIO$ + MES$ + DIA$
                     End If

            Case 25: Dato$ = MONEDA_CAE$                                ' TIPO DE MONEDA
            Case 26: Dato$ = TRIM$(FORMATNUM$(TIPOCAM#, "F12.6"))       ' COTIZACION
            
          End Select
          
          ' NOTA !!!!************************************************
          ' EN LA FILA 24 SE DEBE PONER LA FECHA DE VENCIMIENTO DEL COMPROBANTE
          ' PROVISORIAMENTE SE PUSO LA MISMA DEL DIA
          '*********************************************************************
        TX$ = TX$ + Dato$ + ETIQUETACIERRE$ ' CONCATENA EL RENGLON A IMPRIMIR
        Dato$ = ""
        Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
    Loop
    '
    ' COMPROBANTES ASOCIADOS
    If TIFADOC% = 2 And chkModoCredito.Value = 1 Then
        TX$ = "<CbtesAsoc>": Print #NARCHI2%, TX$
        
        NPVX$ = Label12
        EPAC$ = EMPREAC$
        If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "TANIT") > 0 Or InStr(EPAC$, "GABAL") > 0 Or InStr(EPAC$, "FERVI") > 0 Or InStr(EPAC$, "AHP") > 0 Or InStr(EPAC$, "J Y M") > 0 Or InStr(EPAC$, "CEYEQ") > 0 Or InStr(EPAC$, "EYM") > 0 Or InStr(EPAC$, "DAVTECNICA") > 0 Then NPVX$ = "0001"
        
        'OBTENGO CUIT DE EMPRESA ACTIVA
        n1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
        Dim X As String * 41
        Get #n1%, 6, X
        CUITT$ = TRIM$(REPLACAR$(DESENCRI$(X), "-", ""))
        Close #n1%
             
        If CONSICLI% > 0 Then
            ' FACTURA REMITOS EN CONSIGNACION
            TX$ = "<CbteAsoc>": Print #NARCHI2%, TX$
            For IIX& = 1 To ULTREG&("!CARDEFAC")
                XX$ = REGLEIDO$("!CARDEFAC", IIX&)
                REMITO$ = TRIM$(Mid$(XX$, 387, 16))
                
                If InStr(1, AgrupaRemitos$, REMITO$) > 0 Then GoTo 55
                
                AgrupaRemitos$ = AgrupaRemitos$ & REMITO$
                
                TX$ = "<Tipo>91</Tipo>": Print #NARCHI2%, TX$
                TX$ = "<PtoVta>" & CStr(Val(NPVX$)) & "</PtoVta>": Print #NARCHI2%, TX$
                NROREMITO$ = TRIM$(Mid$(REMITO$, 4))
                TX$ = "<Nro>" & NROREMITO$ & "</Nro>": Print #NARCHI2%, TX$
                TX$ = "<Cuit>" & CUITT$ & "</Cuit>": Print #NARCHI2%, TX$
                FechaRemito% = CVINT(VALOBADA("movisto", "fechmov", "dopricor='" & REMITO$ & "'"))
                FechaRemitoUSA$ = Format(CVDAT(FechaRemito%), "yyyymmdd")
                TX$ = "<CbteFch>" & FechaRemitoUSA$ & "</CbteFch>": Print #NARCHI2%, TX$
55          Next IIX&
            TX$ = "</CbteAsoc>": Print #NARCHI2%, TX$
        Else
            COPIASREM% = XVALO(FINFACTUR07.Text2(1))
            If COPIASREM% > 0 Then
                ' EMITE JUNTOS FACTURA Y REMITO
                TX$ = "<CbteAsoc>": Print #NARCHI2%, TX$
                TX$ = "<Tipo>91</Tipo>": Print #NARCHI2%, TX$
                TX$ = "<PtoVta>" & CStr(Val(NPVX$)) & "</PtoVta>": Print #NARCHI2%, TX$
                NROREMITO$ = TRIM$(Text10)
                TX$ = "<Nro>" & NROREMITO$ & "</Nro>": Print #NARCHI2%, TX$
                TX$ = "<Cuit>" & CUITT$ & "</Cuit>": Print #NARCHI2%, TX$
                FechaRemito% = HOY(HO$)
                FechaRemitoUSA$ = Format(CVDAT(FechaRemito%), "yyyymmdd")
                TX$ = "<CbteFch>" & FechaRemitoUSA$ & "</CbteFch>": Print #NARCHI2%, TX$
                TX$ = "</CbteAsoc>": Print #NARCHI2%, TX$
            End If
        End If
        TX$ = "</CbtesAsoc>": Print #NARCHI2%, TX$
    
    ElseIf (TIFADOC% = 3 Or TIFADOC% = 6) And chkModoCredito.Value = 1 Then
           
        'OBTENGO CUIT DE EMPRESA ACTIVA
        n1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
        Dim regi As String * 41
        Get #n1%, 6, regi
        CUITT$ = TRIM$(REPLACAR$(DESENCRI$(regi), "-", ""))
        Close #n1%
             
        ' APLICA A FACTURA
        PuntoVenta$ = Mid$(AplicacionFacturaCredito$, 4, 4)
        NumeroFactura$ = Mid$(AplicacionFacturaCredito$, 9, 8)
        Tipo$ = Mid$(AplicacionFacturaCredito$, 1, 2)
        If Tipo$ = "FC" Then TipoComprobanteAfip$ = "201"
        If Tipo$ = "ND" Then TipoComprobanteAfip$ = "202"
        If Tipo$ = "NC" Then TipoComprobanteAfip$ = "203"
        TX$ = "<CbtesAsoc>": Print #NARCHI2%, TX$
        
        TX$ = "<CbteAsoc>": Print #NARCHI2%, TX$
        TX$ = "<Tipo>" & TipoComprobanteAfip$ & "</Tipo>": Print #NARCHI2%, TX$
        TX$ = "<PtoVta>" & CStr(Val(PuntoVenta$)) & "</PtoVta>": Print #NARCHI2%, TX$
        TX$ = "<Nro>" & CStr(Val(NumeroFactura$)) & "</Nro>": Print #NARCHI2%, TX$
        TX$ = "<Cuit>" & CUITT$ & "</Cuit>": Print #NARCHI2%, TX$
        FechaFactura% = CVINT(VALOBADA("cajabanc", "fechconta", "nuclien=" & Text1 & " and codicom_lar='" & AplicacionFacturaCredito$ & "' and tipomovim='fven'"))
        FechaFacturaUSA$ = Format(CVDAT(FechaFactura%), "yyyymmdd")
        TX$ = "<CbteFch>" & FechaFacturaUSA$ & "</CbteFch>": Print #NARCHI2%, TX$
        TX$ = "</CbteAsoc>": Print #NARCHI2%, TX$
        
        TX$ = "</CbtesAsoc>": Print #NARCHI2%, TX$
    End If
    
    ' ATRIBUTOS
    If DATO19# > 0.005 Then
        TX$ = "<Tributos>": Print #NARCHI2%, TX$
        If XVALO(Text26) > 0.005 Then
            TX$ = "<Tributo>": Print #NARCHI2%, TX$
            TX$ = "<Id>7</Id>": Print #NARCHI2%, TX$
            TX$ = "<Desc>Perc de IIBB Tucuman</Desc>": Print #NARCHI2%, TX$
            TX$ = "<BaseImp>" & TRIM$(Total(2)) & "</BaseImp>": Print #NARCHI2%, TX$
                ALICUOTA$ = TRIM$(FORMATNUM$((XVALO(Text26) / XVALO(Total(2))) * 100, "F10.2"))
            TX$ = "<Alic>" & ALICUOTA$ & "</Alic>": Print #NARCHI2%, TX$
            TX$ = "<Importe>" & TRIM$(Text26) & "</Importe>": Print #NARCHI2%, TX$
            TX$ = "</Tributo>": Print #NARCHI2%, TX$
        End If
        
        If XVALO(Text25) > 0.005 Then
            TX$ = "<Tributo>": Print #NARCHI2%, TX$
            TX$ = "<Id>7</Id>": Print #NARCHI2%, TX$
            TX$ = "<Desc>Perc de IIBB Misiones</Desc>": Print #NARCHI2%, TX$
            TX$ = "<BaseImp>" & TRIM$(Total(2)) & "</BaseImp>": Print #NARCHI2%, TX$
                ALICUOTA$ = TRIM$(FORMATNUM$((XVALO(Text25) / XVALO(Total(2))) * 100, "F10.2"))
            TX$ = "<Alic>" & ALICUOTA$ & "</Alic>": Print #NARCHI2%, TX$
            TX$ = "<Importe>" & TRIM$(Text25) & "</Importe>": Print #NARCHI2%, TX$
            TX$ = "</Tributo>": Print #NARCHI2%, TX$
        End If
        
        If XVALO(Text24) > 0.005 Then
            TX$ = "<Tributo>": Print #NARCHI2%, TX$
            TX$ = "<Id>7</Id>": Print #NARCHI2%, TX$
            TX$ = "<Desc>Perc de IIBB CABA</Desc>": Print #NARCHI2%, TX$
            TX$ = "<BaseImp>" & TRIM$(Total(2)) & "</BaseImp>": Print #NARCHI2%, TX$
                ALICUOTA$ = TRIM$(FORMATNUM$((XVALO(Text24) / XVALO(Total(2))) * 100, "F10.2"))
            TX$ = "<Alic>" & ALICUOTA$ & "</Alic>": Print #NARCHI2%, TX$
            TX$ = "<Importe>" & TRIM$(Text24) & "</Importe>": Print #NARCHI2%, TX$
            TX$ = "</Tributo>": Print #NARCHI2%, TX$
        End If
        
        If XVALO(Text15) > 0.005 Then
            TX$ = "<Tributo>": Print #NARCHI2%, TX$
            TX$ = "<Id>7</Id>": Print #NARCHI2%, TX$
            TX$ = "<Desc>Perc de IIBB BSAS</Desc>": Print #NARCHI2%, TX$
            TX$ = "<BaseImp>" & TRIM$(Total(2)) & "</BaseImp>": Print #NARCHI2%, TX$
                ALICUOTA$ = TRIM$(FORMATNUM$((XVALO(Text15) / XVALO(Total(2))) * 100, "F10.2"))
            TX$ = "<Alic>" & ALICUOTA$ & "</Alic>": Print #NARCHI2%, TX$
            TX$ = "<Importe>" & TRIM$(Text15) & "</Importe>": Print #NARCHI2%, TX$
            TX$ = "</Tributo>": Print #NARCHI2%, TX$
        End If
    
        TX$ = "</Tributos>": Print #NARCHI2%, TX$
    End If
     
    ' IVA
    If DATO18# > 0.005 Then
        TX$ = "<Iva>": Print #NARCHI2%, TX$
        If XVALO(Total(3)) > 0.005 Then
            TX$ = "<AlicIva>": Print #NARCHI2%, TX$
            TX$ = "<Id>5</Id>": Print #NARCHI2%, TX$
            TX$ = "<BaseImp>" & TRIM$(BASIVA(0)) & "</BaseImp>": Print #NARCHI2%, TX$
            TX$ = "<Importe>" & TRIM$(Total(3)) & "</Importe>": Print #NARCHI2%, TX$
            TX$ = "</AlicIva>": Print #NARCHI2%, TX$
        End If
            
        If XVALO(Total(4)) > 0.005 Then
            TX$ = "<AlicIva>": Print #NARCHI2%, TX$
            TX$ = "<Id>4</Id>": Print #NARCHI2%, TX$
            TX$ = "<BaseImp>" & TRIM$(BASIVA(1)) & "</BaseImp>": Print #NARCHI2%, TX$
            TX$ = "<Importe>" & TRIM$(Total(4)) & "</Importe>": Print #NARCHI2%, TX$
            TX$ = "</AlicIva>": Print #NARCHI2%, TX$
        End If
        
        TX$ = "</Iva>": Print #NARCHI2%, TX$
    End If
    
    ' OPCIONALES
    If TIFADOC% = 2 And chkModoCredito.Value = 1 Then
        TX$ = "<Opcionales>": Print #NARCHI2%, TX$
        
        TX$ = "<Opcional>": Print #NARCHI2%, TX$
        TX$ = "<Id>2101</Id>": Print #NARCHI2%, TX$
        TX$ = "<Valor>" & TRIM$(Control$("FACTURA", "CBUCREDI")) & "</Valor>": Print #NARCHI2%, TX$
        TX$ = "</Opcional>": Print #NARCHI2%, TX$
                
        TX$ = "</Opcionales>": Print #NARCHI2%, TX$

    ElseIf (TIFADOC% = 3 Or TIFADOC% = 6) And chkModoCredito.Value = 1 Then
        
        TX$ = "<Opcionales>": Print #NARCHI2%, TX$
        
        TX$ = "<Opcional>": Print #NARCHI2%, TX$
        TX$ = "<Id>22</Id>": Print #NARCHI2%, TX$
        TX$ = "<Valor>" & TipoAplicacionFacturaCredito$ & "</Valor>": Print #NARCHI2%, TX$
        TX$ = "</Opcional>": Print #NARCHI2%, TX$
                
        TX$ = "</Opcionales>": Print #NARCHI2%, TX$
    
    End If
    
    ' CIERRE DE ETIQUETAS
    Print #NARCHI2%, "</FEDetalleRequest>"
    Print #NARCHI2%, "</Fedr>"
    Print #NARCHI2%, "</Fer>"
    '
    '
    Close #NARCHI1%
    Close #NARCHI2%
    '
99  Exit Sub
    
ControlError:
    
    Call MENSERR(24, "Problemas al Eliminar 'consulta_cae_v1.txt'.\ \Contacte a su Soporte de Sistemas.")
    
End Sub

Function DECILISTACT%()
   '
   Static NLISANT%, DLISANT%
   '
   'CANTIDAD DE DECIMALES DE LA LISTA
   NLIS% = XVALO(TEXT21(0)) 'NUMERO DE LISTA
   If NLIS% = NLISANT% Then
      DECILISTACT% = DLISANT%
      Exit Function
   End If
   '
   DECILISTACT% = 4 'CANTIDAD DE DECIMALES -- SI ES MENOR A 2 DEVUELVE 2
   '
   If InStr(1, EMPREAC$, "LAFE") > 0 Then GoTo 90
   If FACELEC% > 0 Then
      If DECILISTACT% > 3 Then
         Call COMUNI("Factura Electronica Lleva Precios Unitarios\con Tres Decimales Unicamente")
         DECILISTACT% = 3
      End If
   End If
   '
90 DLISANT% = DECILISTACT%
   NLISANT% = NLIS%
   '
End Function

Sub Controla_Idcaenum()
   '
   'CONTROLAR SI EL CAMPO IDCAE ES > A IDCAENUM ACTUAIZA ESTE ULTIMO
   '
   'TOMA LOS ULTIMOS 180 DIAS PARA QUE NO DEMORE DEMASIADO TIEMPO
   FE% = HOY(HO$)
   FE_180% = DIANTE(FE%, 180)
   DESDE = FETEXCOR1$(FE_180%)
   '
   IdentiCaeMax$ = TRIM$(Str$(XVALO(VALOBADA("CAJABANC", "MAX(cast(substring(IDCAE,1,LEN(IDCAE))as INT))", " TIPOMOVIM = 'FVEN' AND FECHEMISI >= '" & DESDE & "'", "NOMESS"))))
   CAENUM = XVALO(IdentiCaeMax$)
   '
   IDCAENUM1 = XVALO(VALOBADA("CAJABANC", "MAX(IDCAENUM)", " TIPOMOVIM = 'FVEN' AND FECHEMISI >= '" & DESDE & "'", "NOMESS"))
   '
   ' SI EL CAE (FORMATO TEXTO) ES > AL IDCAENUM ACTUALIZA ESTE ULTIMO
   If CAENUM > 0 And CAENUM > IDCAENUM1 Then
      CONDI$ = "IDCAE = '" & IdentiCaeMax$ & "'"
      Call ACTUREGISTRO("CAJABANC", "IDCAENUM", CONDI$, CAENUM, REAF&, "NOMESS")
   End If
   '
End Sub
Sub Controla_NroRemito(NPVX$)
   '
   'TOMA LOS ULTIMOS 180 DIAS PARA QUE NO DEMORE DEMASIADO TIEMPO EN EL CASO QUE NO TENGA FECHA DE
   FE% = FECHANUM(Control$("", "FENUREFA"))
   If FE% < 11000 Then
      FE% = HOY(HO$)
      FE_180% = DIANTE(FE%, 180)
      DESDE = FETEXCOR1$(FE_180%)
   Else
      DESDE = FETEXCOR1$(FE%)
   End If
   '
   Call ABRECONEXION
   PARTIZ$ = TRIM$(NPVX$): While Len(PARTIZ$) < 4: PARTIZ$ = "0" + PARTIZ$: Wend
   PARTIZ$ = "RT-X-" + Left$(PARTIZ$, 4)

   CONDI$ = " CODIMOVI = 'RT' AND FECHMOV >= '" & DESDE & "' AND DOPRILAR LIKE '" + PARTIZ$ + "%'"
   DOPRITEX$ = TRIM$(VALOBADA("MOVISTO", "max(DoPriLar)", CONDI$, "NOMESS"))
   
   DOPRINUM& = XVALO(Mid$(DOPRITEX$, 11, 8))
   '
   DOPRINUM_EXISTENTE& = XVALO(VALOBADA("MOVISTO", "MAX(DOPRINUM)", CONDI$, "NOMESS"))
   '
   ' SI EL CAE (FORMATO TEXTO) ES > AL IDCAENUM ACTUALIZA ESTE ULTIMO
   If DOPRINUM& > 0 And DOPRINUM& > DOPRINUM_EXISTENTE& Then
      'CONDI$ = "DOPRINUM = '" & DOPRITEX$ & "'"
       CONDI$ = " CODIMOVI = 'RT' AND FECHMOV >= '" & DESDE & "' AND DOPRILAR = '" + DOPRITEX$ + "'"
      Call ACTUREGISTRO("MOVISTO", "DOPRINUM", CONDI$, DOPRINUM&, REAF&, "NOMESS")
      Call ACTUREGISTRO("MOVISTO", "NUMPVTA", CONDI$, XVALO(NPVX$), REAF&, "NOMESS")
   End If
   '
   
End Sub

Sub SELPEDIP1()
   '
   ' AGREGADO
     '\\\OT-07-0371-FL-01/10/07///
'   Static AGRUPAENTREGA%
'   If AGRUPAENTREGA% = 0 Then
'      AGRUPAENTREGA% = (-1)
'      If CONTROL$("FACTUREM", "AGRUPEDI") = "TRANSUCU" Then AGRUPAENTREGA% = 1
'   End If
'     '\\\OT-07-0371-FL-FIN///
'   ' FIN AGREGADO
'   '
   Screen.MousePointer = 11
   NC% = XVALO(Text1)
   Call ABREPEDCLI
   Set PediModifTemp = New ADODB.Recordset
   SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom,PEDENCA.REFERENCIA, PEDENCA.FeSolEnt "
   SQLX$ = SQLX$ + " FROM PEDENCA "
   SQLX$ = SQLX$ + " INNER JOIN  PEDDETA with(nolock)"
   SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
   SQLX$ = SQLX$ + " WHERE PEDENCA.NroClie = " & NC% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.STATUS >= 0 "   ' APROBADOS
   SQLX$ = SQLX$ + " AND PEDDETA.Status < 8 "    ' para que no este anulada ni cumplida
   SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "    ' para que no esté anulado
   SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
   'Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   Set PediModifTemp = READSET(SQLX$)
   '
   SELPEDIPEN2.LICOPEN.Clear
   SELPEDIPEN2.Label5 = "0"
   HAYPEDIPE% = 0
   With PediModifTemp
      On Error Resume Next
      .MoveFirst
        If Err < 1 Then
                Do While Not .EOF
                    NRODES& = !NROPED
                  TTXX$ = FORMATNUM$(NRODES&, "F6.0") + "  "
                    REFE$ = TRIM$(!REFERENCIA)
                    FEX = CVINT(!FECHEMIS)
                    FEX1 = CVINT(!FeSolEnt)
                  TTXX$ = TTXX$ + FECHATEX$(FEX) + " - " + FECHATEX$(FEX1) + " - "
                  TTXX$ = TTXX$ + TRIM$(REFE$) + Space$(52)
                  Debug.Print TTXX$
                  SELPEDIPEN2.LICOPEN.AddItem TTXX$
                  HAYPEDIPE% = 1
                .MoveNext
                Loop
             End If
          End With
   '
   Screen.MousePointer = 1
   If HAYPEDIPE% > 0 Then
        SELPEDIPEN2.Show 1
      Else
        Call COMUNI("No Hay Pedidos APROBADOS Pendientes\de Entrega para el Cliente Elegido.")
   End If
   '
   CANSELE% = XVALO(SELPEDIPEN2.Label5)
   If CANSELE% < 1 Then
      JJ& = 0
      GoTo 20     ' PARA BLANQUEAR ARCHIVO DE CARGA
   End If
   '
'   ' Recorro pedidos seleccionados y llamo a funcion revisanumero
'   ' por error generado en silcar al modificar pedido repetia nuemro de orden con esto corrige
'   For KKL& = 1 To SELPEDIPEN2.LICOPEN.ListCount
'      If SELPEDIPEN2.LICOPEN.Selected(KKL& - 1) = True Then
'         NPESELEC& = XVALO(Left$(SELPEDIPEN2.LICOPEN.LIST(KKL& - 1), 6))
'         Call REVISANUMERO(NPESELEC&)
'      End If
'   Next KKL&
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   '
   Call APERBASDAT("PEDCLI")
'     '\\\OT-07-0371-FL-01/10/07///
   SQLX$ = " SELECT PEDDETA.codiint, PEDDETA.CanPed, PEDDETA.CantEnt, PEDDETA.NroPed,PEDDETA.FechEmis, "
   SQLX$ = SQLX$ + "PEDDETA.FeSolEnt, PEDDETA.PreUnid, PEDDETA.MoneEmis,PEDDETA.TipoPed,PEDDETA.NuOrItem, PEDENCA.CondPag, "
   SQLX$ = SQLX$ + "PEDDETA.CBASE , PEDDETA.TTALLE , PEDDETA.TCOLOR, "
   SQLX$ = SQLX$ + "PEDDETA.PreList, PEDDETA.DescLista, PEDDETA.DescCombi, "
   SQLX$ = SQLX$ + "PEDENCA.DomiEnt, PEDENCA.LocaEnt, PEDENCA.TranspEnt, "
   SQLX$ = SQLX$ + "PEDENCA.NTranspo, PEDENCA.TDescApli "
   SQLX$ = SQLX$ + " FROM PEDDETA with(nolock)"
   SQLX$ = SQLX$ + "INNER JOIN PEDENCA ON (PEDDETA.NroPed = PEDENCA.NroPed) "
   SQLX$ = SQLX$ + " WHERE PEDDETA.NroClie = " & NC% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
   SQLX$ = SQLX$ + " AND PEDDETA.STATUS >= 0 "   ' APROBADOS
   'SQLX$ = SQLX$ + " AND (((PEDDETA.CanPed - PEDDETA.CantEnt) >= 0.05) OR (PEDDETA.TipoPed = 'ABIERTO' AND CDBL(PEDDETA.FinVigen) >= " & Int(CDbl(Now)) & ")) "
   SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05)"
   SQLX$ = SQLX$ + " OR (PEDDETA.TipoPed = 'ABIERTO' AND PEDDETA.FinVigen >= GETDATE()))"   '
   ATXX% = 0
   '
   '
   For KKU& = 1 To SELPEDIPEN2.LICOPEN.ListCount
      If SELPEDIPEN2.LICOPEN.Selected(KKU& - 1) = True Then
         NPESEL& = XVALO(Left$(SELPEDIPEN2.LICOPEN.List(KKU& - 1), 6))
         If ATXX% < 1 Then
             SQLX$ = SQLX$ + "AND (PEDDETA.NroPed = " & NPESEL& & " "
             ATXX% = 1
           Else
             SQLX$ = SQLX$ + "OR PEDDETA.NroPed = " & NPESEL & " "
         End If
      End If
   Next KKU&
   SQLX$ = SQLX$ + ") "
   '
   SQLX$ = SQLX$ + " ORDER BY PEDDETA.NroPed, PEDDETA.NuOrItem" 'ESTA LINEA LA AGREGUE
   '
   Set PEDTEMP = READSET(SQLX$)
   '
   With PEDTEMP
     On Error Resume Next
     .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("Ningun Registro Coincide con Condición SQL.")
        GoTo 20
      End If
      On Error GoTo 0
     DESCARTAPED% = 0
     '
     JJ& = 0: CAITCAR% = 0
     Do While (.EOF = False)
       CODIMAT% = !CODIINT
       If CODIMAT% > 0 And CODIMAT% <= NUMAS% Then
         CAN = !CANPED - !CANTENT
           If CAPEN < 0 Then CAPEN = 0
         If !TipoPed = "ABIERTO" Then CAPEN = 999999
         NPEDID& = !NROPED
         FEMIPED% = CVINT(!FECHEMIS)
         FENTRESOL% = CVINT(!FeSolEnt)
         PRECI1# = !PreUnid
         MONEMA% = !MoneEmis
         
            If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
         COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         PRECIO = PRECI1# * COEFCON
'        ESTO QUE SIGUE ES ANTERIOR A LA IMPLEMENTACION DE DAKOT
'           If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
'         PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
'           If PRELI < PRECIO Then PRELI = PRECIO
'         PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         '
         ' ESTO CORRESPONDE A DAKOT
'         PRELI = !PreList: If PRELI < 0.00005 Then PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
'         PDLISTA = XVALO(!DescLista)
'         DCOMBI0$ = TRIM$(SAFETEXT$(!DescCombi))
'         DPEDIDO$ = TRIM$(SAFETEXT$(!TDescApli))
'         DCOMBI$ = DCOMBI0$ + "+" + DPEDIDO$
'         If Left$(DCOMBI$, 1) = "+" Then DCOMBI$ = Mid$(DCOMBI$, 2)
'         If Right$(DCOMBI$, 1) = "+" Then DCOMBI$ = Left$(DCOMBI$, Len(DCOMBI$) - 1)
'         If Len(DCOMBI$) > 16 Then
'            Call MENSERR(24, "Imposible Aplicar Descuento Combinado - Consulte")
'         End If
         PDCOMBI = PORCEDESCU(DCOMBI$)
         '
'         'VALIDACION PARA FUINITER SALTEA DESCUENTO DE PEDIDO
'         If InStr(EPAC$, "FUINYTER") < 1 Then
'           If DPEDIDO$ <> "" Then PRECIO = PRECIO * (1 - PORCEDESCU(DPEDIDO$) / 100)
'         End If
'         '
'         PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         '
         ' AGREGADO
'         If AGRUPAENTREGA% > 0 Then
'           DOMIENPE$ = TRIM$(UCase$(!DOMIENT))
'           LOCAENPE$ = TRIM$(UCase$(!LOCAENT))
'           TRANENPE$ = TRIM$(UCase$(!TRANSPENT))
'           CONDPAGO% = !CONDPAG
'           NTRANS% = XVALO(!NTranspo)
'           '
'           If XVALO(Label32) < 1 Then
'              If DOMIENPE$ <> "" Then LENTREGA(0) = DOMIENPE$: Label32 = "1"
'              If LOCAENPE$ <> "" Then LENTREGA(1) = LOCAENPE$: Label32 = "1"
'              If NTRANS% <> 0 Then Text19.TEXT = TRIM$(Str$(NTRANS%)): Label32 = "1"
'              If TRANENPE$ <> "" Then LENTREGA(2) = TRANENPE$: Label32 = "1"
'              If CONDPAGO% <> 0 Then TEXT21(1).TEXT = CONDPAGO%: Label32 = "1"
'
'           End If
'           If DOMIENPE$ <> "" Then If DOMIENPE$ <> LENTREGA(0) Then GoTo 28
'           If LOCAENPE$ <> "" Then If LOCAENPE$ <> LENTREGA(1) Then GoTo 28
'           If NTRANS% <> 0 Then If NTRANS% <> XVALO(Text19.TEXT) Then GoTo 28
'           If TRANENPE$ <> "" Then If TRANENPE$ <> LENTREGA(2) Then GoTo 28
'           If CONDPAGO% <> 0 Then If CONDPAGO% <> XVALO(TEXT21(1).TEXT) Then GoTo 28
'
'         End If
         ' FIN AGREGADO
         '
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = ""
         Call REPLA(Z$, NREMIT$, 387, 14) ' numero de remito
           FEREMIT% = 0
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(Z$, MKI$(0), 401, 2)
         Call REPLA(Z$, NUPLANO$(CODIMAT%), 315, 24)
         '
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
           NORITPE% = !NuOrItem
         Call REPLA(Z$, MKI$(NORITPE%), 503, 2)
         Call REPLA(Z$, DCOMBI$, 487, 16)
         Call REPLA(Z$, MKS$(PDLISTA), 505, 4)
         Call REPLA(Z$, MKS$(PDCOMBI), 509, 4)
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(Z$, NUOCOX$, 355, 16)
         '
       
         COMPDESCRI$ = ""
'         If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then
'            COLORBA$ = SAFETEXT(!CBASE)
'            TELA$ = SAFETEXT(!TTALLE)
'            COLTELA$ = SAFETEXT(!TCOLOR)
'            Call REPLA(Z$, COLORBA$, 283, 10)
'            Call REPLA(Z$, TELA$, 293, 10)
'            Call REPLA(Z$, COLTELA$, 315, 24)
'            If TRIM$(COLOBA$) + TRIM$(TELA$) + TRIM$(COLTELA$) <> "" Then
'              COMPDESCRI$ = "CB: " + COLORBA$ + " - T: " + TELA$ + " - CT: " + COLTELA$
'            End If
'         End If
         Call REPLA(Z$, COMPDESCRI$, 75, 64)
         '
         JJ& = JJ& + 1
         Call GRAREG("!CARDEFAC", Z$, JJ&)
         CAITCAR% = CAITCAR% + 1
         '
       End If
       ' AGREGADO
         '\\\OT-07-0371-FL-01/10/07///
       GoTo 29
'28     DESCARTAPED% = 1
        '\\\OT-07-0371-FL-FIN///
       ' FIN AGREGADO
29     .MoveNext
     Loop
   End With
   '
'   If InStr(EPAC$, "FUINYTER") > 0 Then
'     If DPEDIDO$ <> "" Then
'        Call COMUNI("El o Los Pedidos Seleccionados Tienen Descuentos\Dichos Descuentos han sido Suprimidos en la Factura Actual")
'     End If
'   End If
20 Call SETULTREG("!CARDEFAC", JJ&)
   '
   ' AGREGADO
     '\\\OT-07-0371-FL-01/10/07///
'   If DESCARTAPED% = 1 Then
'      Call MENSERR(24, "Se han Descartado Pedidos que no Corresponden\con la Sucursal o con la Condición de Venta\del Pedido más Antiguo")
'   End If
     '\\\OT-07-0371-FL-FIN///
   ' FIN AGREGADO
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     Z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(Z$, 1, 4))) >= 0.01 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", Z$, JJ&)
     End If
   Next U&
   Call SETULTREG("!CARDEFAC", JJ&)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   Call CIERRARCH("!CARDEFAC")
   '
End Sub




