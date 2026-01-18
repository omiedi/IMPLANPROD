VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORPED07 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos de Clientes - Formato Lineal"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   11820
   Icon            =   "FORPEDCOPIA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   11820
   Begin VB.ListBox List3 
      Height          =   450
      Left            =   11040
      Sorted          =   -1  'True
      TabIndex        =   152
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fact"
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
      Height          =   360
      Left            =   11000
      TabIndex        =   143
      ToolTipText     =   "Genera Factura en Forma Directa"
      Top             =   6780
      Width           =   945
   End
   Begin VB.CommandButton Command14 
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
      Height          =   720
      Left            =   11000
      Picture         =   "FORPEDCOPIA.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   142
      ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
      Top             =   7200
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFFFFF&
      Height          =   6735
      Left            =   10950
      ScaleHeight     =   6675
      ScaleWidth      =   945
      TabIndex        =   127
      Top             =   0
      Width           =   1000
      Begin VB.CommandButton Command22 
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
         Height          =   580
         Left            =   105
         Picture         =   "FORPEDCOPIA.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Anotador de Pedidos"
         Top             =   3380
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command12 
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
         Height          =   580
         Left            =   105
         Picture         =   "FORPEDCOPIA.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   139
         ToolTipText     =   "Consultas y Re-Impresión"
         Top             =   3380
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   580
         Left            =   105
         Picture         =   "FORPEDCOPIA.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   105
         Width           =   650
      End
      Begin VB.CommandButton Command4 
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
         Height          =   580
         Left            =   105
         Picture         =   "FORPEDCOPIA.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   2790
         Width           =   650
      End
      Begin VB.CommandButton Command5 
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
         Height          =   580
         Left            =   105
         Picture         =   "FORPEDCOPIA.frx":11B4
         Style           =   1  'Graphical
         TabIndex        =   135
         ToolTipText     =   "Configuración de Funcionamiento de Pedidos"
         Top             =   4650
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   105
         Picture         =   "FORPEDCOPIA.frx":15F6
         Style           =   1  'Graphical
         TabIndex        =   134
         ToolTipText     =   "Acceso a Condiciones Comerciales"
         Top             =   5280
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   590
         Left            =   105
         Picture         =   "FORPEDCOPIA.frx":1900
         Style           =   1  'Graphical
         TabIndex        =   133
         ToolTipText     =   "Acceso a Maestro de Artículos"
         Top             =   2000
         Width           =   650
      End
      Begin VB.CommandButton Command11 
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
         Height          =   580
         Left            =   105
         Picture         =   "FORPEDCOPIA.frx":1C0A
         Style           =   1  'Graphical
         TabIndex        =   132
         ToolTipText     =   "Acceso a Base de Datos de Clientes"
         Top             =   1420
         Width           =   650
      End
      Begin VB.CommandButton Command13 
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
         Height          =   580
         Left            =   105
         Picture         =   "FORPEDCOPIA.frx":204C
         Style           =   1  'Graphical
         TabIndex        =   131
         ToolTipText     =   "Opciones Varias - Modificaciones y Consultas"
         Top             =   4000
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Height          =   590
         Left            =   105
         Picture         =   "FORPEDCOPIA.frx":2196
         Style           =   1  'Graphical
         TabIndex        =   130
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   685
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   128
         Top             =   6000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   129
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -315
         Picture         =   "FORPEDCOPIA.frx":24A0
         Top             =   6280
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2730
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   125
      Top             =   630
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C0C0C0&
      Height          =   330
      Left            =   7465
      ScaleHeight     =   270
      ScaleWidth      =   3225
      TabIndex        =   113
      Top             =   550
      Visible         =   0   'False
      Width           =   3280
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
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
         Left            =   840
         MaxLength       =   2
         TabIndex        =   115
         Text            =   " "
         Top             =   0
         Width           =   350
      End
      Begin VB.CommandButton Command27 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1180
         TabIndex        =   114
         Top             =   0
         Width           =   180
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   230
         Left            =   60
         TabIndex        =   117
         Top             =   15
         Width           =   750
      End
      Begin VB.Label Label28 
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
         Height          =   300
         Left            =   1330
         TabIndex        =   116
         Top             =   0
         Width           =   1895
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C0C0&
      Height          =   330
      Left            =   8930
      ScaleHeight     =   270
      ScaleWidth      =   1725
      TabIndex        =   111
      Top             =   550
      Width           =   1785
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Venta"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   900
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   0
         TabIndex        =   112
         Top             =   0
         Visible         =   0   'False
         Width           =   1785
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3300
      Left            =   100
      TabIndex        =   32
      Top             =   4650
      Width           =   10700
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resumen Final"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1560
         Left            =   5570
         TabIndex        =   48
         ToolTipText     =   "Habilita Descuento Retira Cliente"
         Top             =   105
         Width           =   5055
         Begin VB.TextBox Text141 
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
            Height          =   330
            Left            =   3405
            TabIndex        =   146
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   1050
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Dt.Rt.Cte."
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
            Left            =   1570
            TabIndex        =   145
            Top             =   1155
            Visible         =   0   'False
            Width           =   1178
         End
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
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   140
            Top             =   720
            Width           =   1455
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "I.V.A."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   105
            TabIndex        =   92
            Top             =   2205
            Width           =   1485
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "IVA 2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   420
               TabIndex        =   98
               Top             =   620
               Width           =   1005
            End
            Begin VB.Label Label19 
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "T-2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   97
               Top             =   620
               Width           =   405
            End
            Begin VB.Label Label15 
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "T-0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   96
               Top             =   190
               Width           =   400
            End
            Begin VB.Label Label16 
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "T-1"
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
               Left            =   60
               TabIndex        =   95
               Top             =   405
               Width           =   400
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "IVA 1"
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
               Left            =   420
               TabIndex        =   94
               Top             =   405
               Width           =   1005
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "IVA 0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   420
               TabIndex        =   93
               Top             =   190
               Width           =   1005
            End
         End
         Begin VB.TextBox Text14 
            BackColor       =   &H00E0E0E0&
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
            Height          =   330
            Index           =   2
            Left            =   3400
            TabIndex        =   90
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   1050
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text15 
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
            Height          =   330
            Left            =   3405
            TabIndex        =   84
            Top             =   2250
            Width           =   1500
         End
         Begin VB.TextBox Total 
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
            Height          =   315
            Index           =   3
            Left            =   3405
            TabIndex        =   83
            Text            =   " "
            Top             =   2625
            Width           =   1500
         End
         Begin VB.TextBox TOTIVA 
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
            Height          =   330
            Index           =   3
            Left            =   3405
            TabIndex        =   82
            Text            =   " "
            Top             =   1870
            Width           =   1500
         End
         Begin VB.TextBox TICAMBIS 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   430
            MaxLength       =   14
            TabIndex        =   74
            Top             =   1080
            Width           =   1065
         End
         Begin VB.TextBox Text14 
            BackColor       =   &H00E0E0E0&
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
            Height          =   330
            Index           =   1
            Left            =   3405
            TabIndex        =   72
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   680
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text14 
            BackColor       =   &H00E0E0E0&
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
            Height          =   330
            Index           =   0
            Left            =   3405
            TabIndex        =   71
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   315
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Total 
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
            Height          =   300
            Index           =   0
            Left            =   3405
            TabIndex        =   51
            Text            =   " "
            Top             =   315
            Width           =   1500
         End
         Begin VB.TextBox Total 
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
            Height          =   330
            Index           =   1
            Left            =   3405
            TabIndex        =   50
            TabStop         =   0   'False
            Text            =   " "
            Top             =   680
            Width           =   1500
         End
         Begin VB.TextBox PORDESCU 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1995
            MaxLength       =   14
            TabIndex        =   23
            Top             =   680
            Width           =   1065
         End
         Begin VB.TextBox Total 
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
            Height          =   330
            Index           =   2
            Left            =   3405
            TabIndex        =   49
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1050
            Width           =   1500
         End
         Begin VB.Label Label31 
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
            Height          =   330
            Left            =   2755
            TabIndex        =   147
            Top             =   1050
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda:"
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
            Index           =   8
            Left            =   -240
            TabIndex        =   141
            Top             =   420
            Width           =   1065
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   540
            Left            =   1755
            Picture         =   "FORPEDCOPIA.frx":43C2
            Top             =   -75
            Width           =   540
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Impuestos (I.V.A.):"
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
            Index           =   1
            Left            =   1365
            TabIndex        =   87
            Top             =   1950
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Percep.I.Brutos:"
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
            Index           =   6
            Left            =   1785
            TabIndex        =   86
            Top             =   2310
            Width           =   1485
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto Pedido:"
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
            Index           =   7
            Left            =   1365
            TabIndex        =   85
            Top             =   2730
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   230
            Index           =   4
            Left            =   1155
            TabIndex        =   77
            Top             =   420
            Width           =   1065
         End
         Begin VB.Label MONEMIS 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H00005500&
            Height          =   285
            Left            =   430
            TabIndex        =   76
            ToolTipText     =   "Doble-Click para 'Dólares'"
            Top             =   720
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "T-C:"
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
            Left            =   60
            TabIndex        =   75
            Top             =   1125
            Width           =   435
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Neto:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   230
            Index           =   0
            Left            =   2250
            TabIndex        =   54
            Top             =   390
            Width           =   1020
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Pedido:"
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
            Left            =   1785
            TabIndex        =   53
            Top             =   1155
            Width           =   1500
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "%:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   230
            Left            =   3150
            TabIndex        =   52
            Top             =   780
            Width           =   225
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   1590
         Index           =   0
         Left            =   5520
         TabIndex        =   78
         Top             =   720
         Visible         =   0   'False
         Width           =   2220
         Begin VB.TextBox TOTIVA 
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
            Height          =   300
            Index           =   0
            Left            =   210
            TabIndex        =   81
            Text            =   " "
            Top             =   210
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TOTIVA 
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
            Height          =   300
            Index           =   1
            Left            =   210
            TabIndex        =   80
            Text            =   " "
            Top             =   525
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TOTIVA 
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
            Height          =   300
            Index           =   2
            Left            =   210
            TabIndex        =   79
            Text            =   " "
            Top             =   840
            Visible         =   0   'False
            Width           =   870
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Condiciones Comerciales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   3120
         Left            =   110
         TabIndex        =   40
         Top             =   105
         Width           =   5330
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0C0C0&
            Height          =   420
            Left            =   1200
            TabIndex        =   155
            Top             =   1770
            Width           =   615
            Begin VB.OptionButton Option1 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Option1"
               Height          =   200
               Left            =   60
               TabIndex        =   157
               Top             =   160
               Width           =   255
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Option2"
               Height          =   200
               Left            =   330
               TabIndex        =   156
               Top             =   160
               Width           =   255
            End
         End
         Begin VB.TextBox TXTFLETE 
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
            Left            =   1890
            TabIndex        =   153
            Text            =   " "
            ToolTipText     =   "Doble Click Para Borrar"
            Top             =   1860
            Width           =   3270
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   1575
            TabIndex        =   123
            Top             =   1500
            Width           =   180
         End
         Begin VB.TextBox TEXT21 
            Alignment       =   1  'Right Justify
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
            Index           =   3
            Left            =   1155
            TabIndex        =   122
            Text            =   " "
            Top             =   1500
            Width           =   435
         End
         Begin VB.TextBox Eco21 
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
            Index           =   3
            Left            =   1890
            TabIndex        =   121
            Text            =   " "
            Top             =   1500
            Width           =   3270
         End
         Begin VB.TextBox Text44 
            Alignment       =   2  'Center
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
            Left            =   4620
            MaxLength       =   4
            TabIndex        =   70
            Top             =   1110
            Width           =   540
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   1155
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   2280
            Width           =   4005
         End
         Begin VB.TextBox Eco21 
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
            Left            =   1890
            TabIndex        =   43
            Text            =   " "
            Top             =   1110
            Width           =   1800
         End
         Begin VB.TextBox TEXT21 
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
            Left            =   1155
            TabIndex        =   17
            Text            =   " "
            Top             =   1110
            Width           =   435
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   1575
            TabIndex        =   18
            Top             =   1110
            Width           =   180
         End
         Begin VB.TextBox Eco21 
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
            Left            =   1890
            TabIndex        =   42
            Text            =   " "
            Top             =   705
            Width           =   2085
         End
         Begin VB.TextBox TEXT21 
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
            TabIndex        =   15
            Text            =   " "
            Top             =   720
            Width           =   435
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1575
            TabIndex        =   16
            Top             =   705
            Width           =   180
         End
         Begin VB.TextBox Eco21 
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
            Left            =   1890
            TabIndex        =   41
            Text            =   " "
            Top             =   300
            Width           =   3270
         End
         Begin VB.TextBox TEXT21 
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
            TabIndex        =   13
            Text            =   " "
            Top             =   300
            Width           =   435
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1575
            TabIndex        =   14
            Top             =   300
            Width           =   180
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Flete:"
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
            Index           =   7
            Left            =   75
            TabIndex        =   154
            Top             =   1890
            Width           =   1050
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Ing.:"
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
            Index           =   6
            Left            =   75
            TabIndex        =   124
            Top             =   1575
            Width           =   1050
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   4650
            TabIndex        =   110
            Top             =   720
            Width           =   510
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Des. Rec."
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
            Index           =   5
            Left            =   4005
            TabIndex        =   109
            Top             =   765
            Width           =   645
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Items:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   210
            TabIndex        =   89
            Top             =   2535
            Width           =   645
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Left            =   420
            TabIndex        =   88
            Top             =   2805
            Width           =   540
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "% Comis:"
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
            Index           =   4
            Left            =   3675
            TabIndex        =   69
            Top             =   1185
            Width           =   855
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Observac:"
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
            Index           =   3
            Left            =   75
            TabIndex        =   47
            Top             =   2220
            Width           =   1050
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
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
            Index           =   2
            Left            =   75
            TabIndex        =   46
            Top             =   1185
            Width           =   1050
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.Pago:"
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
            Index           =   1
            Left            =   75
            TabIndex        =   45
            Top             =   780
            Width           =   1050
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L.Precios:"
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
            Index           =   0
            Left            =   75
            TabIndex        =   44
            Top             =   375
            Width           =   1050
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1560
         Left            =   5570
         TabIndex        =   37
         Top             =   1650
         Width           =   5055
         Begin VB.CommandButton Command24 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1800
            TabIndex        =   105
            Top             =   1200
            Width           =   195
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
            Index           =   2
            Left            =   2040
            MaxLength       =   48
            TabIndex        =   104
            Top             =   1200
            Width           =   2865
         End
         Begin VB.TextBox Text19 
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
            Left            =   1155
            TabIndex        =   103
            Top             =   1200
            Width           =   645
         End
         Begin VB.CommandButton Command25 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1800
            TabIndex        =   100
            Top             =   840
            Width           =   195
         End
         Begin VB.TextBox Text20 
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
            Left            =   1155
            TabIndex        =   99
            Top             =   870
            Width           =   645
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
            Index           =   1
            Left            =   1155
            MaxLength       =   32
            TabIndex        =   25
            Top             =   555
            Width           =   3750
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
            Index           =   0
            Left            =   1155
            MaxLength       =   32
            TabIndex        =   24
            Top             =   240
            Width           =   3750
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
            Left            =   1155
            MaxLength       =   32
            TabIndex        =   60
            Top             =   555
            Visible         =   0   'False
            Width           =   3750
         End
         Begin VB.Label Label24 
            Height          =   225
            Left            =   105
            TabIndex        =   107
            Top             =   840
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label23 
            Height          =   225
            Left            =   105
            TabIndex        =   106
            Top             =   525
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
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
            Index           =   13
            Left            =   150
            TabIndex        =   102
            Top             =   975
            Width           =   880
         End
         Begin VB.Label Label27 
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
            Height          =   300
            Left            =   2040
            TabIndex        =   101
            Top             =   870
            Width           =   2865
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Transporte:"
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
            Index           =   5
            Left            =   30
            TabIndex        =   39
            Top             =   1305
            Width           =   1005
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Domicilio:"
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
            Left            =   150
            TabIndex        =   38
            Top             =   360
            Width           =   915
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   8050
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   10950
      Begin VB.CommandButton Command28 
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
         Left            =   7320
         TabIndex        =   149
         Top             =   1280
         Width           =   160
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
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
         Left            =   6360
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   1280
         Width           =   975
      End
      Begin VB.PictureBox MARSELPRESU 
         BackColor       =   &H00FFFF00&
         Height          =   2940
         Left            =   210
         ScaleHeight     =   2880
         ScaleWidth      =   10455
         TabIndex        =   61
         Top             =   1680
         Visible         =   0   'False
         Width           =   10520
         Begin VB.ListBox List2 
            BackColor       =   &H00F0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Left            =   -30
            Style           =   1  'Checkbox
            TabIndex        =   67
            Top             =   420
            Width           =   9780
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFF00&
            Height          =   470
            Left            =   -30
            ScaleHeight     =   405
            ScaleWidth      =   9720
            TabIndex        =   65
            Top             =   -30
            Width           =   9780
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Selección de Presupuestos en Curso"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   345
               Left            =   0
               TabIndex        =   66
               Top             =   70
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
            Picture         =   "FORPEDCOPIA.frx":46CC
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Continúa Generación del Pedido de Cliente"
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
            Picture         =   "FORPEDCOPIA.frx":49D6
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Cancela Pedido de Cliente"
            Top             =   1050
            Width           =   570
         End
         Begin VB.CommandButton Command17 
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
            Picture         =   "FORPEDCOPIA.frx":4CE0
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Imprime Ficha de Anotaciones"
            Top             =   1890
            Width           =   570
         End
         Begin VB.Line Line5 
            X1              =   9750
            X2              =   9750
            Y1              =   0
            Y2              =   2940
         End
         Begin VB.Line Line4 
            X1              =   10575
            X2              =   10575
            Y1              =   0
            Y2              =   3150
         End
      End
      Begin VB.PictureBox MARCONOTA 
         BackColor       =   &H00FFFFC0&
         Height          =   2940
         Left            =   210
         ScaleHeight     =   2880
         ScaleWidth      =   10455
         TabIndex        =   55
         Top             =   1680
         Visible         =   0   'False
         Width           =   10520
         Begin VB.CommandButton Command16 
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
            Picture         =   "FORPEDCOPIA.frx":534A
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Imprime Ficha de Anotaciones"
            Top             =   2205
            Width           =   570
         End
         Begin VB.CommandButton Command15 
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
            Picture         =   "FORPEDCOPIA.frx":59B4
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Cancela Pedido de Cliente"
            Top             =   1050
            Width           =   570
         End
         Begin VB.CommandButton Command6 
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
            Picture         =   "FORPEDCOPIA.frx":5CBE
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Continúa Generación del Pedido de Cliente"
            Top             =   420
            Width           =   570
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2540
            Left            =   -30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   58
            Top             =   370
            Width           =   9780
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFC0&
            Height          =   435
            Left            =   -30
            ScaleHeight     =   375
            ScaleWidth      =   9720
            TabIndex        =   56
            Top             =   -30
            Width           =   9780
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Anotaciones de Clientes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   225
               Left            =   0
               TabIndex        =   57
               Top             =   105
               Width           =   9825
            End
         End
         Begin VB.Line Line2 
            X1              =   10575
            X2              =   10575
            Y1              =   0
            Y2              =   3150
         End
         Begin VB.Line Line1 
            X1              =   9750
            X2              =   9750
            Y1              =   0
            Y2              =   3200
         End
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Command26"
         Height          =   495
         Left            =   8760
         TabIndex        =   120
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command23 
         Caption         =   "!"
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
         Left            =   1250
         TabIndex        =   108
         Top             =   890
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   735
         Top             =   0
      End
      Begin VB.CommandButton Command20 
         Height          =   720
         Left            =   4995
         Picture         =   "FORPEDCOPIA.frx":5FC8
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Búsqueda por CUIT, Domicilio, Nombre de Fantasía, etc"
         Top             =   890
         Width           =   540
      End
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
         Height          =   2535
         Left            =   210
         TabIndex        =   59
         Top             =   2100
         Width           =   10515
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.CommandButton Command10 
         Height          =   260
         Left            =   420
         TabIndex        =   36
         Top             =   6855
         Width           =   2825
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   210
         ScaleHeight     =   375
         ScaleWidth      =   10455
         TabIndex        =   34
         Top             =   1680
         Width           =   10515
         Begin VB.Label Label22 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Cant Código          Descripción            Fecha   Lista  Desc  P.Unit     Importe"
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
            Height          =   380
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   10515
         End
      End
      Begin VB.TextBox Text12 
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
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1280
         Width           =   3435
      End
      Begin VB.TextBox Text11 
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
         Left            =   9615
         TabIndex        =   11
         Text            =   " "
         Top             =   920
         Width           =   1120
      End
      Begin VB.TextBox Text10 
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
         Left            =   8820
         MaxLength       =   16
         TabIndex        =   12
         Top             =   1280
         Width           =   1920
      End
      Begin VB.TextBox Text9 
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
         Left            =   9615
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   180
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
         Height          =   300
         Left            =   10560
         TabIndex        =   10
         Top             =   180
         Width           =   180
      End
      Begin VB.TextBox Text7 
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
         Left            =   5775
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   330
         Width           =   1695
      End
      Begin VB.TextBox Text8 
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
         Left            =   5775
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   900
         Width           =   1695
      End
      Begin VB.TextBox Text5 
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
         Left            =   2310
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   920
         Width           =   2595
      End
      Begin VB.TextBox Text4 
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
         Left            =   1470
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   920
         Width           =   750
      End
      Begin VB.TextBox Text3 
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
         Left            =   1470
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   550
         Width           =   4065
      End
      Begin VB.TextBox Text2 
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
         Left            =   1470
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   180
         Width           =   4065
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
         Height          =   300
         Left            =   1050
         TabIndex        =   1
         Top             =   890
         Width           =   180
      End
      Begin VB.TextBox Text1 
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   890
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid GRID2 
         Height          =   765
         Left            =   4800
         TabIndex        =   144
         Top             =   3600
         Visible         =   0   'False
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1349
         _Version        =   393216
         Rows            =   23
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   8433840
         BackColorSel    =   -2147483643
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         Redraw          =   -1  'True
         FocusRect       =   2
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
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F.Ent:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   270
         Left            =   5640
         TabIndex        =   150
         Top             =   1425
         Width           =   735
      End
      Begin VB.Image Image4 
         Height          =   450
         Left            =   7500
         Picture         =   "FORPEDCOPIA.frx":732E
         Stretch         =   -1  'True
         ToolTipText     =   "Doble Click para Acceso a Pedidos Abiertos"
         Top             =   150
         Width           =   400
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         Height          =   450
         Left            =   7900
         Picture         =   "FORPEDCOPIA.frx":7770
         Stretch         =   -1  'True
         ToolTipText     =   "Doble Click Para Importar desde B.Datos MDB"
         Top             =   150
         Width           =   400
      End
      Begin VB.Label CIMP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C. Imp:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   7485
         TabIndex        =   119
         Top             =   855
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label30 
         Height          =   255
         Left            =   7620
         TabIndex        =   118
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image Image6 
         Height          =   450
         Left            =   8310
         Picture         =   "FORPEDCOPIA.frx":78BA
         Stretch         =   -1  'True
         ToolTipText     =   "Doble Click Para Importar desde Planilla Excel"
         Top             =   150
         Width           =   400
      End
      Begin VB.Label Label26 
         BackColor       =   &H00EEEEEE&
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
         Height          =   300
         Left            =   210
         TabIndex        =   91
         Top             =   1280
         Width           =   1035
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   11865
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "OV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   615
         Left            =   150
         TabIndex        =   68
         Top             =   75
         Width           =   1020
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   315
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   315
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   210
         Shape           =   4  'Rounded Rectangle
         Top             =   165
         Width           =   960
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   10710
         Y1              =   4640
         Y2              =   4640
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   210
         Left            =   5775
         TabIndex        =   33
         Top             =   690
         Width           =   1695
      End
      Begin VB.Line Line8 
         X1              =   10725
         X2              =   10725
         Y1              =   1700
         Y2              =   4640
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido Nro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   230
         Left            =   8250
         TabIndex        =   31
         Top             =   975
         Width           =   1320
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/C o N/P:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   230
         Left            =   7515
         TabIndex        =   30
         Top             =   1335
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   8820
         TabIndex        =   29
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Límite Crédito:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   230
         Left            =   5775
         TabIndex        =   28
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   230
         Left            =   50
         TabIndex        =   27
         Top             =   680
         Width           =   960
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FORPEDCOPIA.frx":8904
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "FORPEDCOPIA.frx":8B38
      TabIndex        =   126
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   11160
      TabIndex        =   151
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   105
      Picture         =   "FORPEDCOPIA.frx":8BB2
      ToolTipText     =   "Recepción Pedidos Electrónicos"
      Top             =   8715
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Menu Archi 
      Caption         =   "Archivos"
      Begin VB.Menu nueped 
         Caption         =   "Nuevo Pedido"
      End
      Begin VB.Menu MoPeCur 
         Caption         =   "Modificar Pedido en Curso"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu LiPrVe 
         Caption         =   "Listas de Precios de Venta"
      End
      Begin VB.Menu php0a 
         Caption         =   "-"
      End
      Begin VB.Menu maeart 
         Caption         =   "Maestro de Articulos"
      End
      Begin VB.Menu maecli 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu sali 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Edic 
      Caption         =   "Edición"
      Begin VB.Menu MoPeCu 
         Caption         =   "Modificación de Pedido en Curso"
      End
      Begin VB.Menu AnuPed 
         Caption         =   "Anulación de Pedidos"
      End
      Begin VB.Menu RevAnu 
         Caption         =   "Revertir Anulación de Pedidos"
      End
      Begin VB.Menu php1a 
         Caption         =   "-"
      End
      Begin VB.Menu AnuItem 
         Caption         =   "Anulación por Items de Pedidos"
      End
   End
   Begin VB.Menu ConLis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu pedistzona 
         Caption         =   "Pedidos Pendientes C/Stock y Zona"
      End
      Begin VB.Menu PePeCli 
         Caption         =   "Pedidos Pendientes por Cliente"
         Begin VB.Menu PeClPan 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu PeCliEx 
            Caption         =   "E"
         End
      End
      Begin VB.Menu Proyepedidos 
         Caption         =   "Visualizar Proyección de Pedidos"
      End
      Begin VB.Menu PePeCo 
         Caption         =   "Pedidos Pendientes por Código de Producto"
         Begin VB.Menu PeCoPa 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu PeCoEx 
            Caption         =   "Imprimir/Exportar"
         End
      End
      Begin VB.Menu php1b 
         Caption         =   "-"
      End
      Begin VB.Menu EncaPe 
         Caption         =   "Encabezados de Pedidos"
      End
      Begin VB.Menu DeNumPe 
         Caption         =   "Detalle por Número de Pedido"
      End
      Begin VB.Menu php2 
         Caption         =   "-"
      End
      Begin VB.Menu mnusitstock 
         Caption         =   "Sit.Stocks"
      End
      Begin VB.Menu ConRimp 
         Caption         =   "Consultas y Re-Impresión"
      End
   End
   Begin VB.Menu AccDirec 
      Caption         =   "Accesos Directos"
      Begin VB.Menu abmReserva 
         Caption         =   "ABM Reservas"
      End
      Begin VB.Menu mnuVerSituacionpedidoproyeccion 
         Caption         =   "Ver Situación de Pedidos y Proyección"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu LiPreVen 
         Caption         =   "Listas de Precios de Ventas"
      End
      Begin VB.Menu php2a 
         Caption         =   "-"
      End
      Begin VB.Menu MaesArt 
         Caption         =   "Maestro de Articulos"
      End
      Begin VB.Menu MaesCli 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu php3 
         Caption         =   "-"
      End
      Begin VB.Menu RCtaCte 
         Caption         =   "Resumen de Cuenta Corriente"
      End
      Begin VB.Menu php4 
         Caption         =   "-"
      End
      Begin VB.Menu pedidossuc 
         Caption         =   "Pedidos de Reposición de Sucursal"
         Visible         =   0   'False
         Begin VB.Menu Verocomsuc 
            Caption         =   "Ver Pedido de Reposición de Sucursal"
         End
         Begin VB.Menu descargarocsuc 
            Caption         =   "Descargar Archivos de Pedidos de Reposición de Sucursal"
         End
      End
      Begin VB.Menu otr 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu Configura 
      Caption         =   "Configuración"
      Begin VB.Menu ConGral 
         Caption         =   "Configuración General"
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu ConCom 
         Caption         =   "Condiciones Comerciales"
         Begin VB.Menu LiPrVta 
            Caption         =   "Listas de Precios "
         End
         Begin VB.Menu CoPag 
            Caption         =   "Condiciones de Pago"
         End
      End
   End
   Begin VB.Menu ayuda 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "FORPED07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$
Dim STATUPC%
Dim CAMON%, DMONE$(), TCMONE()
Public PORCODBARRA


Sub AGREGABONIFICACION(U&, IKITEM%)
       'LEO EL REGISTRO
       X$ = REGLEIDO$("!CARDETPE", U&)
       
       CIXI% = CVI(Left$(X$, 2))
       CAN = CVS(Mid$(X$, 89, 4))
       NROPE& = XVALO(Text11.TEXT)
       FEMI% = FECHANUM(Text9.TEXT)
       NCLIE = XVALO(Text1.TEXT)
       STATUPC% = (-1)
       If Left$(CONTROL$("PEDCLI", "MODOAPRO"), 6) = "DIRECT" Then STATUPC% = 0
       CAVER$ = TRIM$(Str$(NROPE&))
       While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
       While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend

       IK$ = TRIM$(Str$(IKITEM%))
       CAVERI$ = CAVER$ + "-" + IK$
       '
       'CREO UN NUEVO REGISTRO
       strsql = "Select * from PedDeta where NroPed=" & NROPE&
       Dim PedDeta1 As ADODB.Recordset
       Set PedDeta1 = New ADODB.Recordset
26     On Error Resume Next
       PedDeta1.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 26
       End If
       On Error GoTo 0
       With PedDeta1
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CODIEMPR = CodiEmp%
         On Error GoTo 0
           !NROPED = NROPE&
           !FECHEMIS = CVDAT(FEMI%)
           !NROCLIE = NCLIE
           !RaSoClie = RASOCLI$(NCLIE)
           !NROOCOM = TRIM$(Text10.TEXT)
             FENTRESOL% = CVI(Mid$(X$, 63, 2))
           !FESOLeNT = CVDAT(FENTRESOL%)
           !Status = STATUPC%              ' Status  de Aprobación
           !ListPre = XVALO(TEXT21(0).TEXT)
           !CondPag = XVALO(TEXT21(1).TEXT)
           !VENDED = XVALO(TEXT21(2).TEXT)
           !PorComVend = XVALO(Text44.TEXT)
           If Left$(UCase$(Label21(4)), 5) = "LETRA" Then
               VALOR% = 0
               If TRIM$(Text44) <> "" Then VALOR% = Asc(Text44.TEXT + "A") - 64
               !PorComVend = VALOR%
           End If

           !NuorItem = IKITEM%
           !CODIINT = CIXI%
           !CODIEXT = CODEXT$(CIXI%)
           !SubCod = Mid$(X$, 109, 8)
           !Descrip = DESCRIT0$(CIXI%)
           !TColor = Mid$(X$, 3, 16)
           !TTalle = Mid$(X$, 19, 8)
           !TLargo = CVS(Mid$(X$, 27, 4))
           !TAncho = CVS(Mid$(X$, 31, 4))
           !TEspesor = CVS(Mid$(X$, 35, 4))
           !VenRep = Left$(Label11, 1)   ' Venta o Reparacion
           !MoneEmis = MONEMI%
           !TipoCam = XVALO(TICAMBIS)
           !PreList = CVS(Mid$(X$, 39, 4))
           !CanAsig = 0
           !CANPED = CAN
           !PREUNID = 0   ' CVS(Mid$(X$, 51, 4))
           !ImpoTot = 0
           !CANTENT = 0
           !CanSaldo = CAN
           !FeUltEnt = CVDAT(0)
           !NroRemi = ""
           !TPedItem = CAVERI$
           '
           '\\\OT-8-766-19/09/08
           !ImpoAnti = 0
           !FApro_Nope = CVDAT(FECHANUM("01/01/80")) 'PARA EVITAR PROBLEMAS CON FORMATO REGIONAL
           !AcuPrepa = 0
           !AcuDespa = 0
           !TipoPed = "" ' se corroboro que este CAMPO es de  TEXTO
           !FinVigen = CVDAT(FECHANUM("01/01/80"))
           !StaDespa = 0
           !CBase = ""
           !DescLista = 0
           !DescCombi = ""
           !DescRubro = 0
           !DescTotal = 0
           !NUMEROSERIE = "" ' SE ESTABLECE VALOR POR DEFAULT PARA EL NUMERO DE SERIE
           '\\\OT-8-766-FIN
           '
           TTXYZ$ = Space$(76)
           Call REPLA(TTXYZ$, "- " + CODEXT$(CIXI%), 11, 18)
           Call REPLA(TTXYZ$, DESCRIT0$(CIXI%), 29, 38)
           Call REPLA(TTXYZ$, CSTRING$(MKS$(CAN), 3, 10, 1, 2), 57, 10)
           ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
           '
         .Update
       End With
       
PedDeta1.Close
Set PedDeta1 = Nothing

End Sub

Sub APP_PREMOTEC()
  On Error Resume Next
  'si es el formulario esta cargado lo cierra
  If IsFormLoaded(STOCMOLD) Then Unload STOCMOLD
  On Error GoTo 0
  
  FIN& = ULTREG&("!CARDETPE")
  CAMPOS$ = "Código/F16;Descripción/A24;Cant.Sol./F10.1;Cant.Unid./F10.1;Sol.Entr./D8;CI1/F0;LPRE/F0"
  Call CARGA_ENCABEZADO(STOCMOLD.MSF1, CAMPOS$, STOCMOLD, 0, 0)
  '
  For U& = 1 To FIN&
    X$ = REGLEIDO$("!CARDETPE", U&)
    CI1% = CVI(Left$(X$, 2))
    CAN = CVS(Mid$(X$, 47, 4))
    LPREXX% = CVS(Mid$(X$, 85, 4))
    FENTRESOL% = CVI(Mid$(X$, 63, 2))
    CODI$ = CODEXT$(CI1%)
    DESCR$ = DESCRIT$(CI1%)
    UNME$ = UNIMED$(CI1%)
    '
    CANTIMPORM2 = CAN ' POR DEFECTO , SI ENTRA EN EL IF CAMBIA
    CANTI_UNIDADES = CAN
    If UNME$ = "M2" Then
      If LPREXX < 1 Then
        Call COMUNI("Imposible Continuar !!  Código: " & CODI$ & " Falta Inofrmar Lista de Precios\Sin este Dato No se Puede Calcular La Cantidad de Unidades a Fabricar\Si no Corrige esta Situación No se Generará la O.Fabricación Correspondiente.")
        Exit For
      End If
      '
      If UNME$ = "M2" Then
        CANTIMPORM2 = CANTIXM2(CI1%, LPREXX%)
      End If
      '
      If CANTIMPORM2 < 0 Then
        Call COMUNI("Imposible Continuar !!  Código: " & CODI$ & " No tiene Relación Con Lista de Precios\Sin este Dato No se Puede Calcular La Cantidad de Unidades a Fabricar\Si no Corrige esta Situación No se Generará la O.Fabricación Correspondiente.")
        Exit For
      End If
      '
      CANTI_UNIDADES = CAN * CANTIMPORM2
      'VERIFICO SI REDONDEO PARA ABAJO LE AGREGO 1 SI NO NADA
      If (CANTI_UNIDADES Mod 1) <> 0 Then
        If ROUND(CANTI_UNIDADES) < CANTI_UNIDADES Then 'REDONDEO PARA ABAJO
          CANTI_UNIDADES = ROUND(CANTI_UNIDADES) + 1
        End If
      End If
      '
      'AQUI CARGAR LA GRILLA.
    End If
    '
    CAMPOS$ = "Código/F16;Descripción/A24;Cant.Sol./F10.1;Cant.Unid./F10.1;Sol.Entr./D8;CI1/F0;LPRE/F0"
    With STOCMOLD
      .MSF1.Rows = .MSF1.Rows + 1
      .MSF1.TextMatrix(.MSF1.Rows - 1, 1) = CODI$
      .MSF1.TextMatrix(.MSF1.Rows - 1, 2) = DESCR$
      .MSF1.TextMatrix(.MSF1.Rows - 1, 3) = CAN
      .MSF1.TextMatrix(.MSF1.Rows - 1, 4) = CANTI_UNIDADES
      .MSF1.TextMatrix(.MSF1.Rows - 1, 5) = FECHATEX$(FENTRESOL%)
      .MSF1.TextMatrix(.MSF1.Rows - 1, 6) = CI1%
      .MSF1.TextMatrix(.MSF1.Rows - 1, 7) = LPREXX%
    End With
  Next U&
    '
    STOCMOLD.MSF1.COL = 5 'xx Columan por la q quieres ordenar
    STOCMOLD.MSF1.Sort = 3 'Aplicar orden al grid
    Screen.MousePointer = 1
    STOCMOLD.Show 1
End Sub

 Function CANTIDADBONIFICADA(CANPEDIDA, DESCU)
     'DEVUELVE LA CANTIDAD QUE SE DEBERIA BONIFICAR SEGUN EL DESCUENTO
     '
     CANTNUE = 0
     If DESCU < 0 Then Exit Function
     
     CANTX = 100 / DESCU
     '
     If CANPEDIDA >= CANTX Then
        CANTNUE = Int(CANPEDIDA / CANTX)
     End If
     '
     CANTIDADBONIFICADA = CANTNUE
     '
 End Function

 Sub CargaBonificacion()
    '
    'CARGA LA BONIFICACION POR PRODUCTO
    '
    NCLIEN = XVALO(Text1)
    FIN& = ULTREG&("!CARDETPE")
    If FIN& > 0 Then
      For U& = 1 To FIN&
        Y$ = REGLEIDO$("!CARDETPE", U&)
        CI% = CVI(Left$(Y$, 2))
        If CI% > 0 And CI% <= NUMAS% Then
            MODO% = 0
            CANTIBONIF = 0
            CANPEDIDA = CVS(Mid$(Y$, 47, 4))
            DESCUXX$ = DESCOMBINA$(NCLIEN, CI%, MODODESC%): MODO% = MODODESC%
    ' MsgBox DESCUXX$ & "   " & MODO%
            If MODO% > 0 Then
               DESCU = ROUND#(PORCEDESCU(DESCUXX$), 2)
               CANTIBONIF = CANTIDADBONIFICADA(CANPEDIDA, DESCU)
            End If
            Call REPLA(Y$, MKS(CANTIBONIF), 89, 4)
            Call GRAREG("!CARDETPE", Y$, U&)
        End If
      Next U&
    End If
 End Sub
Sub DEPURAPED()
    '
    'DEPURACION DE PEDIDOS SEGUN CONFIGURACION DE SETUP DE PEDIDOS
    'SI TIENE MARCA DE CONTROL DE DEPURACION Y EL CASILLERO DE TO LERANCIA ES > 0
    'RECORRE LOS PEDIDOS PENDIENTES SI ESTOS TIENEN FECHA DE ENTREGA MENOR A LA DE HOY Y SUPERAN
    'LA TOLERANCIA CONFIGURADA LE CAMBIA EL STATUS A 9
    CantiDias% = XVALO(CONTROL("ADMIPED", "DEPEPEN"))
    If CONTROL("ADMIPED", "DEPUPPEN") = "SI" And CantiDias% > 0 Then
       'FECHA DE HOY + LOS DIAS DE TOLERANCIA
       FEHOY% = HOY(HO$)
       '
       OPX% = COMALTER%("Atención Rutina De Depuración de Pedidos\Se Anularan Los Items Pendientes de Entrega\Que posean Fecha de Entrega Menor al: " & FECHATEX$(FEHOY% - CantiDias%) & "\\ Cancelar Depuración\ Si Depurar")
       If OPX% <> 2 Then Exit Sub
       
       SQLX$ = "SELECT * FROM PEDDETA "
       SQLX$ = SQLX$ + " WHERE STATUS < 8 "
       SQLX$ = SQLX$ + " AND (CANPED - CANTENT) > 0.05 "
       
       Dim PEDDETTEMP As ADODB.Recordset
       Set PEDDETTEMP = New ADODB.Recordset
25     On Error Resume Next
       PEDDETTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
       End If
       On Error GoTo 0
       '
        With PEDDETTEMP
           Do While Not .EOF
             FESOL% = CVINT(!FESOLeNT)
             'SI LA CANTIDAD DE DIAS DE LA FECHA SOLICITADA COMPARADA CON LA DE HOY ES > A TOLERANCIA
             If DIENTRE(FESOL%, FEHOY%) > CantiDias% Then
                !Status = 9
                .Update
             End If
             .MoveNext
           Loop
        End With
        PEDDETTEMP.Close
        Set PEDDETEMP = Nothing
    End If
End Sub

Sub DESCUMONTO(FORM1 As Form)

    'CALCULA EL DESCUENTO SEGUN VALIDACION DE SETUP DE DESUENTO POR MONTO
    'SE PASA COMO PARAMETRO EL FORM PARA UTILIZAR LA MISMA EN LA MODIFICACION DEL PEDIDO
    'SI TIENE VENDEDOR DECUELVE EL DESCUENTO 1 CASO CONTRARIO DEVUELVE EL DESCEUENTO 2
    'ESTO TOMA SOLO PARA LA LISTA 1
    DESC1 = 0: DESC2 = 0
    '
    TOTPEDIDO = XVALO(FORM1.Total(0))
    'DEVUELVE EN DESC1 Y DESC2 EL DESCUENTO DE ACUERDO AL IMPORTE
    For I& = 1 To ULTREG&("DESMONTO")
       X$ = REGLEIDO$("DESMONTO", I&)
       DESDE1 = CVS(Mid$(X$, 1, 4))
       HASTA1 = CVS(Mid$(X$, 5, 4))
       If TOTPEDIDO >= DESDE1 And TOTPEDIDO <= HASTA1 Then
            DESC1 = CVS(Mid$(X$, 9, 4))
            DESC2 = CVS(Mid$(X$, 13, 4))
            Exit For
       End If
    Next I&
    '
    'POR SI NO COINCIDE  EN NINGUNO Y EL TOTAL DEL PEDIDO SUPERA EL  HASTA SE QUEDA CON EL ULTIMO
    If TOTPEDIDO > HASTA1 Then
       DESC1 = CVS(Mid$(X$, 9, 4))
       DESC2 = CVS(Mid$(X$, 13, 4))
    End If
    '
    VEND% = XVALO(FORM1.TEXT21(2))
    VENDED$ = TRIM$(ECOARCH$("VENDEDOR", Chr$(VEND%)))
    If VEND% = 0 Or InStr(UCase$(VENDED$), "GABAL") Then
       FORM1.PORDESCU = DESC2
    Else
       FORM1.PORDESCU = DESC1
    End If
    '
End Sub

Sub IMPRIMEBONIFICACION(U&)

      XX$ = REGLEIDO$("!CARDETPE", U&)
      CIPRE% = CVI(Left$(XX$, 2))
      CAPEX$ = Str$(CVS(Mid$(XX$, 89, 4)))
      '
      If CIPRE% > 0 Then
        If CIPRE% <= NUMAS% Then
          If XVALO(CAPEX$) > 0 Then
            '
            '\\\OT-07-0107-RG-15-03-07///
            PRELIX = PRELISTA#(ARCHIPRE$, CIPRE%)
            PRELI$ = Str$(PRELIX)
            '\\\OT-07-0107-RG-FIN///
            'PRELI$ = Str$(CVS(Mid$(XX$, 39, 4)))
            PREUX$ = 0
            PRESINRED = 0
            DESCX$ = "0": On Error Resume Next
                          DESCX$ = Str$(100 * (XVALO(PRELI$) - PRESINRED) / XVALO(PRELI$))
                          On Error GoTo 0
            PIVA# = 1
            If PRIVA$ = "SI" Then PIVA# = 1 + PORCEIVA(CIPRE%) / 100 ' VERIFICA QUE TIENE SETEADO EN EL SETUP PARA VER SI IMPRIME CON IVA INCLUIDO
            ITOTX$ = FORMATNUM$((XVALO(CAPEX$)) * (XVALO(PREUX$)) * PIVA#, "F15.4")
            '
            FEX = CVI(Mid$(XX$, 63, 2))
            FENTX$ = FECHATEX$(FEX)
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(CIPRE%)
            TETABU1$(3, CAITAB1%) = CAPEX$
            TETABU1$(4, CAITAB1%) = 0
            TETABU1$(5, CAITAB1%) = 0
            PRENETO# = XVALO(PREUX$) * PIVA# ' PRECIO CON IVA
            PREUX$ = FORMATNUM$(PRENETO#, "F15.4")
            TETABU1$(6, CAITAB1%) = 0
            TETABU1$(7, CAITAB1%) = 0
            TETABU1$(8, CAITAB1%) = FENTX$
            TETABU1$(9, CAITAB1%) = Str$(CAITAB1%)
            TETABU1$(10, CAITAB1%) = NUPLANO$(CIPRE%)
            '
            SUBTOCONIVA = SUBTOCONIVA + XVALO(ITOTX$)
            ' LO QUE SIGUE ES PARA IMPRIMIR O/DESPACHO
            If OPCI% = 2 Then
               TETABU1$(4, CAITAB1%) = ""
               TETABU1$(5, CAITAB1%) = ""
               TETABU1$(6, CAITAB1%) = ""
               TETABU1$(7, CAITAB1%) = ""
               TETABU1$(12, CAITAB1%) = UBISTO$(CIPRE%)
            End If
            ' FIN O/DESPACHO
            '
            PHCODI% = XVALO(ATRIFORM$(FORIPRE$, "COD-MAT/POSINHOR"))
            PHDESI% = XVALO(ATRIFORM$(FORIPRE$, "DES-MAT/POSINHOR"))
            PVCODI% = XVALO(ATRIFORM$(FORIPRE$, "COD-MAT/POSINVER"))
            PVDESI% = XVALO(ATRIFORM$(FORIPRE$, "DES-MAT/POSINVER"))
            If PHDESI% <= PHCODI% + 3 Then
               If PVDESI% <= PVCODI% + 1 Then
                  CAITAB1% = CAITAB1% + 1
               End If
            End If
            DECRI$ = DESCRIT0$(CIPRE%)
            TETABU1$(2, CAITAB1%) = DECRI$
            '
            If PVDESI% <= PVCODI% + 1 Then
               FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
               ATELIB% = XVALO(Mid$(FORTELIB$, 2))
               Call FRATEXTO(DECRI$, ATELIB%)
               TETABU1$(2, CAITAB1%) = RETEX$(1)
               For KKU1% = 2 To CARETEX%
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
               Next KKU1%
            End If
            '
            If OPCI% = 2 Then ' IMPRIME LOS COMPONENTES
               If DESCOMPO%(CIPRE%) > 0 Then
                   CANTI = CVS(Mid$(XX$, 47, 4))
                   Call PRINCOMPO(CIPRE%, CANTI)
               End If
            End If
            '
            If CVS(Mid$(X$, 89, 4)) > 0 Then
               Call IMPRIMEBONIFICACION(U&)
            End If

            
          End If
        End If
      End If

End Sub

Sub REVERTIR_CONTROL()
       CantiDias% = XVALO(CONTROL("ADMIPED", "DEPEPEN"))
       If CONTROL("ADMIPED", "DEPUPPEN") = "SI" And CantiDias% > 0 Then
       
       FEHOY% = HOY(HO$)

       SQLX$ = "SELECT * FROM PEDDETA WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE STATUS > 8 " 'SOLO LAS ANULDAS
       SQLX$ = SQLX$ + " AND (CANPED - CANTENT) > 0.05 "
       SQLX$ = SQLX$ + " "
       '
       Set PEDDETTEMP = READSET(SQLX$)
       With PEDDETTEMP
           Do While Not .EOF
             FESOL% = CVINT(!FESOLeNT)
             'SI LA CANTIDAD DE DIAS DE LA FECHA SOLICITADA COMPARADA CON LA DE HOY ES > A TOLERANCIA
             If DIENTRE(FESOL%, FEHOY%) > CantiDias% Then
                NP& = XVALO(!NROPED)
                NOR% = XVALO(!NuorItem)
                TEXTCONSULTA$ = TEXTCONSULTA$ + " OR NROPED = " & NP& & " AND NuOrItem = " & NOR%
'                !Status = 9
'                .Update
             End If
             .MoveNext
           Loop
       End With
       PEDDETTEMP.Close
       Set PEDDETEMP = Nothing
    End If

End Sub

Sub Verifica_Impresora_Despacho()
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(UCase(EPAC$), "MEDITEA") > 0 Then
       If TRIM$(CONTROL(IDENTER$, "PUERDESP")) = "" Or TRIM$(CONTROL$(IDENTER$, "PUERORVE")) = "" Then
          IMPREDESPA.Show 1
       End If
    End If

End Sub

Private Sub abmReserva_Click()
  RESEPEDI.Show 1
End Sub

Private Sub AnuItem_Click()
Call OPPED07.Command4_Click
End Sub

Private Sub AnuPed_Click()
Call OPPED07.Command45_Click
End Sub

Private Sub ayuda_Click()
Call Command3_Click
End Sub

Private Sub Check2_Click()
   VEND% = XVALO(TEXT21(2))
   VENDED$ = TRIM$(ECOARCH$("VENDEDOR", Chr$(VEND%)))
   If VEND% = 0 Or InStr(UCase$(VENDED$), "GABAL") Then
      If Check2.Value = 1 Then
       Label31 = XVALO(CONTROL("ADMIPED", "DESRTCTE"))
      Else
       GoTo 20
      End If
   Else
20     Label31 = ""
       Check2.Value = 0
   End If
End Sub

Private Sub Combo1_Click()
   '
   MONEMIA% = MONEMI%
   '
   MONEMI% = Combo1.ListIndex + 1
   MONEMIS = DMONE$(MONEMI%)
   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   '
   If MONEMI% > 1 Then
     TICAMBIS.Visible = True
     TICAMBIS.Enabled = False
   End If
   '
   FIN& = ULTREG&("!CARDETPE")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = COEFCON * CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     Call REPLA(X$, MKS$(PRECARGA), 39, 4)
     Call REPLA(X$, MKS$(PRECARGA), 51, 4)
     Call REPLA(X$, MKD$(TOT#), 55, 8)
     Call REPLA(X$, MKS$(PRECARGA), 65, 4)
     Call REPLA(X$, MKS$(PREUNI), 73, 4)
     Call REPLA(X$, MKD$(TOT#), 77, 8)
     Call GRAREG("!CARDETPE", X$, U&)
   Next U&
   '
   Call CARLISPE
   Call CIERRARCH("!CARDETPE")
   Call CALTOPED
   '
End Sub

Private Sub Command1_Click()
Call SELECHO("DEUDOR1")
   NCLIEN = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   Text1.TEXT = TRIM$(Str$(NCLIEN))
   If NCLIEN > 0 Then
     If REGICLI&(NCLIEN) > 0 Then
       Call CARDACLI
       '\\\OT-08-0179-OD-21/04/08///
       FAVEN% = HAYFACVEN%(NCLIEN)
       If FAVEN% > 0 Then
        Text8 = "Fac.Vencidas"
        Text8.ForeColor = RGB(255, 0, 0)
       End If
       '\\\OT-08-0179-OD-FIN///
       '\\\OT-05-0399-WAR-14/06/05///
       Image6.Enabled = True 'MODULO IMPORTADOR
       ''\\\OT-05-0399-WAR-FIN///
       On Error Resume Next
       Command6.SetFocus
       On Error GoTo 0
     Else
        Image6.Enabled = False 'MODULO IMPORTADOR
     End If
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07/AMACLI"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    '
    Command12.Enabled = False
    Screen.MousePointer = 11
    'Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "TANIT") < 1 Then
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
    Else
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 199000 ORDER BY NROPED DESC")
    End If
    
    '
'    SQLX$ = "SELECT * FROM PEDENCA ORDER BY NroPed ASC"
'    'Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
'
'    Dim PEDENCTEMP As ADODB.Recordset
'    Set PEDENCTEMP = New ADODB.Recordset
'    PEDENCTEMP.CursorLocation = adUseClient
'25  On Error Resume Next
'    PEDENCTEMP.Open (SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'    End If
'    On Error GoTo 0
'
'    '
'    JJ& = 0
'    With PEDENCTEMP
'      On Error Resume Next
'      .MoveFirst
'      If Err < 1 Then
'        Do While Not .EOF
'             NRODES& = !NroPed
'             TTXX$ = MKS$(NRODES&)
'               NOCO$ = !NroOcom
'             If NOCO$ <> "" Then
'               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
'             End If
'               FEX = CVINT(!FechEmis)
'             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
'               NCLIE% = !NroClie
'               ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
'             TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + ANULX$ 'Space$(52)
'             JJ& = JJ& + 1
'             Call GRAREG("INDIPEDI", TTXX$, JJ&)
'             '\\\OT-05-0682-WAR-FIN///
'        .MoveNext
'        Loop
'      End If
'    End With
'    PEDENCTEMP.Close
'    Set PEDENCTEMP = Nothing
'    '
'    Call SETULTREG("INDIPEDI", JJ&)
'    Call CIERRARCH("INDIPEDI")
'    Call FRESHECHO("INDIPEDI")
'    Screen.MousePointer = 1
'    '
3   Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
    NPX& = XVALO(VALACT1$("INDIPEDI"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Pedido de Cliente"
    TIDOCUX$ = "Pedido ABIERTO"
    TIDOCUY$ = "Pedido Interno"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = NUDOCU$ + " "
    Screen.MousePointer = 11
    '
    'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
    DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
    If DOCUNU& < 1 Then
         Screen.MousePointer = 1
         Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
         On Error Resume Next
         OPPED07.Command46.SetFocus
         OPPED07.Show 1
         If Err Then
          On Error GoTo 0
         End If
         GoTo 99
    End If
    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    
'4   For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCUM$) > 4 Or InStr(XX$, TIDOCUX$) > 4 Or InStr(XX$, TIDOCUY$) > 4 Then
'        PPXX% = InStr(XX$, "Nro.")
'        If PPXX% > 0 Then
'          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
'        End If
'        If InStr(XX$, NUDOCU$) > 4 Then
'          DOCUNU& = CVS(Left$(XX$, 4))
'          GoTo 5
'        End If
'      End If
'    Next UI&
'    '
'    Screen.MousePointer = 1
'    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'    On Error Resume Next
'    OPPED07.Command46.SetFocus
'    On Error GoTo 0
'    OPPED07.Show 1
'    GoTo 99
'    '
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    GoTo 99
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7170
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'10  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Visible = False
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
'    'EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
'    '
'    GoTo 3
'    '
99  Command12.Enabled = True
    '
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   PROLLAMA$ = "CERR"
   OPPED07.Show 1
   If MONEMI% < 1 Then MONEMI% = 1
   MONEMIS = DEMONECO$(MONEMI%)
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   Call REVERTIR_CONTROL
   
   ' VALIDACION SI SE INGRESO UN NUMERO DE VENDEDOR VALIDO
   VENDED$ = TRIM$(TEXT21(2))
   If XVALO(VENDED$) > 0 Then
      ECOVEN$ = TRIM$(ECOARCH$("VENDEDOR", Chr$(XVALO(VENDED$))))
      If ECOVEN$ = "" Then
        Call MENSERR(24, "Vendedor Ingresado No Válido")
        On Error Resume Next
           TEXT21(2).SelStart = 0
           TEXT21(2).SelLength = Len(TEXT21(2).TEXT)
           TEXT21(2).SetFocus
        On Error GoTo 0
        GoTo 99
      End If
   End If
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
     If TRIM$(TEXT21(2)) = "" Or TRIM$(Text44) = "" Then
       Call MENSERR(24, "Falta Vendedor / Comisión")
       On Error Resume Next
       TEXT21(2).SetFocus
       On Error GoTo 0
       GoTo 99
     End If
   End If
   '
   'validacion si la orden de compra es existente
   If CONTROL$("ADMIPED", "NOREPOC") = "SI" And TRIM$(Text10) <> "" Then
      NCLIE = XVALO(Text1)
      NORCOM$ = UCase$(TRIM$(Text10))
      If OCOMEXISTE%(NCLIE, NORCOM$) Then
         Call COMUNI("Orden de Compra Existente")
         Text10 = ""
         GoTo 99
      End If
   End If
   '
   If CONTROL$("ADMIPED", "EXICOPA") = "SI" Then
      If ECOARCH$("CONPAG", MKI$(XVALO(TEXT21(1).TEXT))) = "" Then
        Call COMUNI("Imposible Continuar. Debe Elegir Una Condición de Pago Válida.")
        GoTo 99
      End If
   End If
'''   'VALIDO LOTES EN CASO QUE VENGAN DE PRESUPUESTOS AHP.
'''   VENCARPE$ = TRIM$(Control$("", "VENCARPE"))
'''   If UCase$(VENCARPE$) = "CARDETLO" Then
'''    For U& = 1 To ULTREG&("!CARDETPE")
'''      X$ = REGLEIDO$("!CARDETPE", U&)
'''      CAN = CVS(Mid$(X$, 47, 4))
'''      If CAN > 0.05 Then
'''        '
'''        CIXI% = CVI(Left$(X$, 2))
'''        If CIXI% > 0 Then
'''           LOTEX$ = TRIM$(Mid$(X$, 3, 16))
'''           If LOTEX$ <> "" Then
'''            SLP = SALDO_LOTE_PEDIDO(CIXI%, LOTEX$)
'''            SLP = SaldoLote(CIXI%, LOTEX$) - SLP
'''            If SLP < CAN Then
'''              Call COMUNI("Lote: " & LOTEX$ & " Para Código: " & CODEXT$(CIXI%) & "\Cantidad Pedida: " & FORMATNUM$(CAN, "F9.1") & " Supera el Saldo (" & SLP & ")")
'''              Call CARGALOTE_PED(XVALO(Text11))
'''              Call CARDETPEDI(1)
'''              GoTo 99
'''            End If
'''           Else
'''              Call COMUNI("Para Código: " & CODEXT$(CIXI%) & "\Cantidad Pedida: " & FORMATNUM$(CAN, "F9.1") & " No Posee Lote")
'''              Call CARGALOTE_PED(XVALO(Text11))
'''              Call CARDETPEDI(1)
'''              GoTo 99
'''           End If
'''        End If
'''
'''      End If
'''    Next U&
'''   End If

   'VALIDA QUE SE HAYA ELEGID UN TIPO INGRESO SIEMPRE QUE SEA OBLIGATORIO
   If CONTROL$("ADMIPED", "VALTIPIN") = "SI" Then
        If XVALO(TEXT21(3)) < 1 Then
            Call COMUNI("Falta Elegir un Tipo de Ingreso")
            GoTo 99
        End If
   End If
   '
   MODOAPROB$ = "" ' VARIABLE PARA FORZAR LA APROBACION POSTERIOR EN CASO QUE NO PASE LAS VALIDACIONES
   'CONTROL DE VALIDACION DE CONDICIONES COMERCIALES
   NC = XVALO(Text1)
   If TRIM$(CONTROL("ADMIPED", "VALCONPA")) = "SI" Then
       CONDIP% = CPAGCLI%(NC)
       If XVALO(TEXT21(1)) <> CONDIP% Then
          OPX% = COMALTER%("Atención !! \Se Ha Cambiado La Condición de Pago Original " & CONDIP% & " Por la " & TEXT21(1) & "\De Aprobarse el Presente Pedido\El Mismo Quedará Pendiente de Aprobación \\Modificar \Continuar ")
          If OPX% <= 1 Then
             GoTo 99
           Else
             MODOAPROB$ = "CAMBIO_PAGO"
          End If
       End If
   End If
   '
   'CONTROL DE VALIDACION DE CONDICIONES COMERCIALES
''   If TRIM$(Control("ADMIPED", "VALLISPR")) = "SI" Then
''       LISTAPRECIO% = Str$(LIPRCLI%(NC))
''       If XVALO(TEXT21(0)) <> LISTAPRECIO% Then
''          OPX% = COMALTER%("Atención !!\Se Ha Cambiado La Lista de Precios Original " & LISTAPRECIO% & " Por la " & TEXT21(0) & "\De Aprobarse el Presente Pedido \El Mismo Quedará Registrado Pendiente de Aprobación \\Modificar\Continuar ")
''          If OPX% <= 1 Then
''             GoTo 99
''          Else
''             MODOAPROB$ = MODOAPROB$ + " CAMBIO_LISTA"
''          End If
''       End If
''   End If
   '
    'CONTROL DE VALIDACION DE FACTURAS VENCIDAS
   If TRIM$(CONTROL("ADMIPED", "FACVENCI")) = "SI" Then
       FIMISQL$ = FETEXCOR1$(HOY(HO$))
       CONDI$ = "Nuclien = " & NC & " AND FEVENCIM <= '" & FIMISQL$ & "' AND SaldDebe >= 0.005"
       FACTUVENCIDAS& = CantRegistros&("SADEUPEN", CONDI$, "NOMESS")
       If FACTUVENCIDAS& > 0 Then
          OPX% = COMALTER%("Atención !! Se Han Detectado Facturas Vencidas Para este Cliente\De Aprobarse el Presente Pedido \El Mismo Quedará Registrado Pendiente de Aprobación \\Modificar\Continuar ")
          If OPX% <= 1 Then
             GoTo 99
          Else
             MODOAPROB$ = MODOAPROB$ + " FACTURA_VENCIDA"
          End If
       End If
   End If
   '
   DIAS_ULT_COMPRA% = XVALO(CONTROL("ADMIPED", "QULTICOM"))
   If DIAS_ULT_COMPRA% > 0 Then
      CANTIFACTURAS& = TIENE_FACTURAS&(NC, DIAS_ULT_COMPRA%, FEULT_FACT%)
      If CANTIFACTURAS& < 1 And FEULT_FACT% > 0 Then
         OPX% = COMALTER%("Atención !! Se Han Detectado que el Cliente Activo\Hace Más de: " & DIAS_ULT_COMPRA% & " Días\Que No Registra Facturas\Última Compra: " & FECHATEX$(FEULT_FACT%) & "\De Aprobarse el Presente Pedido \El Mismo Quedará Registrado Pendiente de Aprobación \\Modificar\Continuar ")
          If OPX% <= 1 Then
             GoTo 99
          Else
             MODOAPROB$ = MODOAPROB$ + " Fecha Última Factura: " & FECHATEX$(FEULT_FACT%)
          End If
      End If
        
   End If
   '
   'CONFIGURACION PARA QUE SEA MANDATARIA EL DOMICILIO DE ENTREGA
   If CONTROL("ADMIPED", "DOMIMAND") = "SI" Then
     If TRIM$(LENTREGA(0).TEXT) = "" Then
        Call COMUNI("Falta Ingresar Domicilio de Entrega")
        GoTo 99
     End If
   End If
   '
   'VALIDO SI ESTA CONFIGURADA EL CARGO DE FLETE Y SI ES ASI LO EXIGE
   If CONTROL("ADMIPED", "FLETMAND") = "SI" Then
        If TRIM$(Me.TXTFLETE) = "" Then
           Call COMUNI("Falta Ingresar a Quien Corresponde el Cargo del Flete")
           GoTo 99
        End If
   End If
   '
   
   Call GRAPEDIDO(OKX%, MODOAPROB$)
   '
   Call DATMOTIVO
   
   If OKX% > 0 Then
     '
     CONPRIPE.Check7.Value = 1
     If CONTROL$("PEDIDO", "PRINOVEN") = "NO" Then CONPRIPE.Check7.Value = 0
     CONPRIPE.Check8.Value = 0
     If CONTROL$("PEDIDO", "PRINODES") = "SI" Then CONPRIPE.Check8.Value = 1
     CONPRIPE.Check2.Value = 0
     If CONTROL$("PEDCLIEN", "SUPRINT") = "SI" Then CONPRIPE.Check2.Value = 1
40   CONPRIPE.Show 1
   
     If CONPRIPE.Check7.Value = 0 And CONPRIPE.Check1.Value = 1 Then
        Call COMUNI("Imposible Enviar Mail.\No Se Ha Elegido la opción Orden de Venta.")
        GoTo 40
     End If
     '
     If CONPRIPE.Label1.Caption = "OK" Then
        '
        CANCELPRINT% = 0
        If CONPRIPE.Check2.Value = 1 Then CANCELPRINT% = 1  ' IMPRESION FANTASMA
          
        If CONPRIPE.Check7.Value = 1 Then
           Call PRIPEDIDO(1)     ' Imprime Orden de Venta
        End If
        '
        'LA IMPRESION FANTASMA ES SOLO PARA LA ORDEN DE DESPACHO
        If CONPRIPE.Check8.Value = 1 Then
           Call PRIPEDIDO(2)     ' Imprime Orden de Despacho
        End If
        '
        ' VERIFICA SI TIENE QUE REPLICAR LA ORDEN DE VENTA
        If CONPRIPE.Check1.Value = 1 And CONPRIPE.Check7.Value = 1 Then
           '
           NCX = XVALO(Text1)
           '
           'PERMITE ENVIAR A PARTIR DE LA LISTA DE MAILS DE CONTACTOS
           COCLI$ = VALOBADA("DEUDOR12", "CODI_CLI", "NUME_CLI = " & NCX, "NOMESS")
           Call CONTACTO_MAIL(COCLI$, NCX) 'AGREGA A SENDBYMAIL LOS MAIL DE LOS CONTACTOS SELECCIONADOS
           If SENDBYMAIL$ = "" Then
              SENDBYMAIL$ = TRIM$(EMAILCLI$(NCX))
           End If
           '
           If SENDBYMAIL$ = "" Or InStr(SENDBYMAIL$, "@") < 1 Then 'TIENE DIRECCION VALIDA
             Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Direccion de Email o la Existente es Inválida.")
             GoTo 30
           End If
           '
25         CANCELPRINT% = 0
           Call PRIPEDIDO(1)
        End If
        '
30      CONPRIPE.Label1 = ""
        CANCELPRINT% = 0
     End If
'\\\OT-08-0687-FL-FIN///
     Call GENERODES      ' genera orden de despacho
     '
     GENORFAB% = 0
     If CONTROL$("ADMIPED", "GENORFAB") = "SI" Then GENORFAB% = 1
     If STATUPC% = (-1) Then
          Call COMUNI("Pedido Registado como\'PENDIENTE DE APROBACIÓN'")
        Else
          ' GENERA ORDEN DE FABRICACION SOLO SI PEDIDO CONFIRMADO
          If GENORFAB% = 1 Then
             Call GENORFA
          End If
     End If
     '
     Screen.MousePointer = 11
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "TANIT") < 1 Then
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
     Else
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 199000 ORDER BY NROPED DESC")
     End If
     Screen.MousePointer = 1
     '
     NPED& = XVALO(Text11)
     Call BLANFORMU
     Call SETULTREG("!CARDETPE", 0)
     Call CIERRARCH("*.*")
     Text1.TEXT = ""
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     '
     If EXISTE%("VENGES07.EXE") > 0 And Check7.Value = 1 Then ' CHECK DE PEDIDO A FACUTURA
       If CONTROL("PEDCLI", "MODOAPRO") = "DIRECT" Then       ' APROBACION DIRECTA
        Call CONECRUN("VENGES07/PEDIDO" & NPED&, "MODULO DE FACTURACION")
       End If
     End If
     '
   End If
   '
   If SENDBYMAIL$ <> "" Then ' EN CASO DE HABER REPLICADO POR MAIL, RE-EJECUTA EL MODULO
      Call RECONECTA
   End If
   
   'VUELVA A 0 EL MOTIVO
   Text16 = "0"
   '
   If Label30 <> "" Then
             Call FINAL("")
   End If
   '
99 Command14.Enabled = True
End Sub

Sub DATMOTIVO()
    '
    'SI EL PEDIDO ES CON MOTIVO PRESENTA LA VENTANA DE CARGA DE DATOS
    NPED& = XVALO(Text11)
    NUMOT% = XVALO(Text16)
    If NUMOT% < 1 Then Exit Sub
    
    DESCMOTI$ = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
    '
10  FILTX$ = TRIM$(Str$(NPED&)) & ";" & TRIM$(Str$(NUMOT%)) & ";" & TRIM$(DESCMOTI$)
    Call STDFORM("DESMO1", FILTX$)
    If STDSCREEN.APRO <> "OK" Then
       If OPX% = 2 Then GoTo 15
       Exit Sub
    End If
    '

15  ' VER QUE SE VA A VALIDAR ACA

End Sub
Sub PRINCOMPO(CI%, CANTI)
        CNTDC1$ = CNTDC(CI%)
        Call AREXPLO(CI%, 1, 1)
        For KKJ% = 1 To CAIT%
          CICO% = CIJ%(KKJ%)
          CACICO = USOI(KKJ%)
          '
          CODI$ = CODEXT$(CICO%)
          DESCRI$ = "..." + DESCRIT0$(CICO%)
          CANTX$ = FORMATNUM$(CACICO * CANTI, "F16." & CNTDC1$ & "")
          '
          Call REPLA(TX$, CODI$, 1, 16)
          Call REPLA(TX$, DESCRI$, 18, 34)
          Call REPLA(TX$, CANTX$, 54, 16)
          Call REPLA(TX$, UBISTO$(CICO%), 80, 16)
          CAITAB1% = CAITAB1% + 1
          TETABU1$(11, CAITAB1%) = TX$
          '
        Next KKJ%
        '
        Call CARLINSEP
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

Private Sub Command17_Click()
   Call MENSERR(24, "No Disponible por el Momento")
End Sub

Private Sub Command18_Click()
   MARSELPRESU.Visible = False
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

''' Sub Command19_Click()
'''   '
'''   NOFRESHPRE% = 0
'''   If Left$(UCase$(Label11), 6) = "REPARA" Then GoTo 50
'''   '
'''   NPX& = 0
'''   Call ABREPRESVEN
'''   HOII% = HOY(HOS$)
'''   NCLIEN = XVALO(Text1)
'''   JJ& = 0: ATX% = 0: COEFCON = 1
'''   For KKI% = 1 To List2.ListCount
'''     If List2.Selected(KKI% - 1) = True Then
'''       NUPRESU& = XVALO(Mid$(List2.List(KKI% - 1), 251, 6))
'''       If NUPRESU& > 0 Then
'''         NOFRESHPRE% = 1
'''         SQLX$ = "SELECT * FROM PRESUVEN "
'''         SQLX$ = SQLX$ + "WHERE NroClie = " & NCLIEN & " "
'''         SQLX$ = SQLX$ + "AND NroPres = " & NUPRESU& & " "
'''         SQLX$ = SQLX$ + "AND NroPedi < 1 "
'''         SQLX$ = SQLX$ + "ORDER BY NroPres ASC, NuOrdIt ASC"
'''         '
'''         'Set IPRESPEN = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
'''         Dim IPRESPEN As ADODB.Recordset
'''         Set IPRESPEN = New ADODB.Recordset
'''         IPRESPEN.CursorLocation = adUseClient
'''25       On Error Resume Next
'''         IPRESPEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'''         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'''            On Error GoTo 0
'''            Call CapturaErrorOpen
'''            GoTo 25
'''         End If
'''         On Error GoTo 0
'''         With IPRESPEN
'''           On Error Resume Next
'''           .MoveFirst
'''           If Err < 1 Then
'''             Do While Not .EOF
'''               If !NroPedi < 1 Then
'''                 ZZ$ = REGBLAN$("!CARDETPE")
'''                   CIXI% = !CODIINT        ' CVI(Mid$(XX$, 69, 2))
'''                 Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
'''                   PRELI = !PreList        ' CVS(Mid(XX$, 71, 4))
'''                   DESCUTEX$ = !TDESCUEN   ' TRIM$(Mid$(XX$, 75, 14))
'''                   CPPP = (100 - COEFDESCU(DESCUTEX$)) / 100
'''                   If CPPP < 1 Then ATX% = 1
'''                   '
'''                   MONEMA% = !MonedEmi     ' Asc(Mid$(XX$, 68, 1))
'''                   If JJ& < 1 Then
'''                     MONEMI% = MONEMA%     ' Asc(Mid$(XX$, 68, 1))
'''                     If MONEMI% < 1 Then MONEMI% = 1
'''                     MONEMIS = DEMONECO$(MONEMI%)
'''                     TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
'''                   End If
'''                   COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
'''                   CAPRESU& = !Cantit      ' CVS(Mid$(XX$, 101, 4))
'''                   If CAPRESU& < 1 Then CAPRESU& = 1
'''                   '
'''                 Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 39, 4)
'''                 Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 51, 4)
'''                 Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 65, 4)
'''                 Call REPLA(ZZ$, MKS$(CAPRESU&), 47, 4)   ' CANTIDAD
'''                   DESCUTEX$ = !TDESCUEN   ' TRIM$(Mid$(XX$, 75, 14))
'''                   'Call REPLA(ZZ$, MKS$(COEFDESCU(DESCUTEX$)), 69, 4)
'''                 Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 73, 4)
'''                 Call REPLA(ZZ$, MKI$(HOII%), 63, 2)
'''                   NUORIT% = NuOrdIt
'''                 'JANDY SQL
'''                 'ACA FALTABA EL !
'''                   NUORIT% = !NuOrdIt
'''
'''                 Call REPLA(ZZ$, MKI$(NUORIT%), 123, 2)
'''                 Call REPLA(ZZ$, MKS$(NUPRESU&), 125, 4)
'''                 JJ& = JJ& + 1
'''                 Call GRAREG("!CARDETPE", ZZ$, JJ&)
'''               End If
'''             .MoveNext
'''             Loop
'''           End If
'''         End With
'''         IPRESPEN.Close
'''         Set IPRESPEN = Nothing
'''       End If
'''     End If
'''19 Next KKI%
'''   Call SETULTREG("!CARDETPE", JJ&)
'''   Call CIERRARCH("!CARDETPE")
'''   GoTo 90
'''   '
'''50 If InStr(1, EMPREAC$, "ARCON") > 0 Then
'''        Call SELECREPARA
'''        GoTo 90
'''   End If
'''
'''   UREPRE& = ULTREG&("PRESREPA")
'''   HOII% = HOY(HOS$)
'''   TEXT21(0) = "0"
'''   JJ& = 0: ATX% = 0
'''   For KKI% = 1 To List2.ListCount
'''     If List2.Selected(KKI% - 1) = True Then
'''       REPRESU& = XVALO(Mid$(List2.List(KKI% - 1), 121, 8))
'''       If REPRESU& > 0 Then
'''         If REPRESU& <= UREPRE& Then
'''           XX$ = REGLEIDO$("PRESREPA", REPRESU&)
'''           ZZ$ = REGBLAN$("!CARDETPE")
'''           CIXI% = CVI(Mid$(XX$, 69, 2))
'''           Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
'''           PRELI = CVS(Mid(XX$, 117, 4))
'''           'DESCUTEX$ = TRIM$(Mid$(XX$, 75, 14))
'''           CPPP = 1  ' (100 - COEFDESCU(DESCUTEX$)) / 100
'''           'If CPPP < 1 Then ATX% = 1
'''           '
'''           MONEMA% = 1   ' REPARACIONES SIEMPRE EN PESOS
'''           MONEMI% = 1
'''           MONEMIS = DEMONECO$(MONEMI%)
'''           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
'''           '
'''           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 39, 4)
'''           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 51, 4)
'''           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 65, 4)
'''           Call REPLA(ZZ$, MKS$(1), 47, 4)   ' CANTIDAD
'''           DESCUTEX$ = TRIM$(Mid$(XX$, 75, 14))
'''           'Call REPLA(ZZ$, MKS$(COEFDESCU(DESCUTEX$)), 69, 4)
'''           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 73, 4)
'''           Call REPLA(ZZ$, MKI$(HOII%), 63, 2)
'''           Call REPLA(ZZ$, MKS$((-1) * REPRESU&), 125, 4)
'''           JJ& = JJ& + 1
'''           Call GRAREG("!CARDETPE", ZZ$, JJ&)
'''         End If
'''       End If
'''     End If
'''   Next KKI%
'''   Call SETULTREG("!CARDETPE", JJ&)
'''   Call CIERRARCH("!CARDETPE")
'''   '
'''90 MARSELPRESU.Visible = False
'''   '
'''   If ATX% > 0 Then
'''      PORDESCU.TEXT = ""
'''      Call COMUNI("ATENCION: Descuento Aplicado según Presupuesto Previo !!!")
'''   End If
'''
'''   PORCODBARRA = 0
'''   EPAC$ = UCase$(TRIM$(EMPREAC$))
'''   If ULTREG&("!CARDETPE") < 1 And InStr(EPAC$, "GABAL") > 0 Then
'''      If COMALTER%("Opciones de Carga de Pedido\\En Forma Manual\Por Código de Barras") > 1 Then
'''         PORCODBARRA = 1
'''         Call CARGAPORCODBAR
'''         Exit Sub
'''      End If
'''   End If
'''   Call CARDETPEDI(1)
'''   '
'''End Sub
'''
Sub Command19_Click()
   '
   NOFRESHPRE% = 0
   If Left$(UCase$(Label11), 6) = "REPARA" Then GoTo 50
   '
   NPX& = 0
   Call ABREPRESVEN
   HOII% = HOY(HOS$)
   NCLIE = XVALO(Text1)
   JJ& = 0: ATX% = 0: COEFCON = 1
   For KKI% = 1 To List2.ListCount
     If List2.Selected(KKI% - 1) = True Then
       NUPRESU& = XVALO(Mid$(List2.LIST(KKI% - 1), 251, 6))
       If NUPRESU& > 0 Then
         NOFRESHPRE% = 1
         SQLX$ = "SELECT * FROM PRESUVEN "
         SQLX$ = SQLX$ + "WHERE NroClie = " & NCLIE & " "
         SQLX$ = SQLX$ + "AND NroPres = " & NUPRESU& & " "
         SQLX$ = SQLX$ + "AND NroPedi < 1 "
         SQLX$ = SQLX$ + "ORDER BY NroPres ASC, NuOrdIt ASC"
         '
         'Set IPRESPEN = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
         Dim IPRESPEN As ADODB.Recordset
         Set IPRESPEN = New ADODB.Recordset
         IPRESPEN.CursorLocation = adUseClient
25       On Error Resume Next
         IPRESPEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
            On Error GoTo 0
            Call CapturaErrorOpen
            GoTo 25
         End If
         On Error GoTo 0
         With IPRESPEN
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
             Do While Not .EOF
               If !NroPedi < 1 Then
                   ZZ$ = REGBLAN$("!CARDETPE")
                   CIXI% = !CODIINT        ' CVI(Mid$(XX$, 69, 2))
                   Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
                   PRELI = !PreList        ' CVS(Mid(XX$, 71, 4))
                   DESCUTEX$ = !TDESCUEN   ' TRIM$(Mid$(XX$, 75, 14))
                   CPPP = (100 - COEFDESCU(DESCUTEX$)) / 100
                   If CPPP < 1 Then ATX% = 1
                   '
                   MONEMA% = XVALO(!MonedEmi)     ' Asc(Mid$(XX$, 68, 1))
                   If JJ& < 1 Then
                     MONEMI% = MONEMA%     ' Asc(Mid$(XX$, 68, 1))
                     If MONEMI% < 1 Then MONEMI% = 1
                     MONEMIS = DEMONECO$(MONEMI%)
                     TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
                   End If
                   COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
                   CAPRESU& = XVALO(!Cantit)      ' CVS(Mid$(XX$, 101, 4))
                   If CAPRESU& < 1 Then CAPRESU& = 1
                   '
                   Call REPLA(ZZ$, MKS$(XVALO(!Num_Prov)), 31, 4) ' NUMERO DE PROVEEDOR '(INT15-372)
                   Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 39, 4)
                   Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 51, 4)
                   Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 65, 4)
                   Call REPLA(ZZ$, MKS$(CAPRESU&), 47, 4)   ' CANTIDAD
                   DESCUTEX$ = !TDESCUEN   ' TRIM$(Mid$(XX$, 75, 14))
                   'Call REPLA(ZZ$, MKS$(COEFDESCU(DESCUTEX$)), 69, 4)
                   Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 73, 4)
                   Call REPLA(ZZ$, MKI$(HOII%), 63, 2)
                   NUORIT% = XVALO(!NuOrdIt)
                 
'                  Call REPLA(ZZ$, MKI$(NUORIT%), 123, 2)
                   Call REPLA(ZZ$, MKS$(NUPRESU&), 125, 4)
                   'CASO NORMAL SIN LOTES CARGADOS EN PRESUPUESTO
                   If TRIM$(SAFETEXT$(!LOTES_Reservados)) = "" Then
                     Call REPLA(ZZ$, "", 3, 16)
                     Call REPLA(ZZ$, MKI$(NUORIT%), 123, 2)
                     JJ& = JJ& + 1
                     Call GRAREG("!CARDETPE", ZZ$, JJ&)
                   Else
                     'CASO CON LOTES RESERVADOS EN EL PRESUPURESTO AHP
                     'SI TIENE LOTES SE VALIDAN QUE TENGAN SALDO SI ES ASI LOS RESERVA.
                     LOTE$ = TRIM$(SAFETEXT$(!LOTES_Reservados))
                     SALLOTE = SaldoLote(CIXI%, LOTE$)
                     ST_RESERVADO = SALDO_LOTE_PEDIDO(CIXI%, LOTE$, XVALO(Text11))
                     SALLOTE = SALLOTE - ST_RESERVADO
                     If SALLOTE < 0 Then SALLOTE = 0
                     If SALLOTE < CAPRESU& Then
                        CantiConlote = SALLOTE
                     Else
                        CantiConlote = CAPRESU&
                     End If
                       
                     CAPRESU& = CAPRESU& - CantiConlote
                     If CantiConlote > 0 Then
                      Call REPLA(ZZ$, LOTE$, 3, 16)
                      Call REPLA(ZZ$, MKS$(CantiConlote), 47, 4)
                      Call REPLA(ZZ$, MKI$(NUORIT%), 123, 2)
                      NUORIT% = NUORIT% + 1
                      '
                      JJ& = JJ& + 1
                      Call GRAREG("!CARDETPE", ZZ$, JJ&)
                      X11$ = REGLEIDO$("!CARDETPE", JJ&)
                      NUORIT% = CVI(Mid$(X11$, 132, 2))
                     End If
                     If CAPRESU& > 0 Then
                         LOTE$ = ""
                         Call REPLA(ZZ$, "", 3, 16)
                         Call REPLA(ZZ$, MKS$(CAPRESU&), 47, 4)
                         Call REPLA(ZZ$, MKI$(NUORIT%), 123, 2)
                         NUORIT% = NUORIT% + 1
                
                         JJ& = JJ& + 1
                         Call GRAREG("!CARDETPE", ZZ$, JJ&)
                         X11$ = REGLEIDO$("!CARDETPE", JJ&)
                         NUORIT% = CVI(Mid$(X11$, 132, 2))
                     End If
                   End If
               End If
             .MoveNext
             Loop
           End If
         End With
         IPRESPEN.Close
         Set IPRESPEN = Nothing
       End If
     End If
19 Next KKI%
   Call SETULTREG("!CARDETPE", JJ&)
   Call CIERRARCH("!CARDETPE")
   GoTo 90
   '
50 If InStr(1, EMPREAC$, "ARCON") > 0 Then
        Call SELECREPARA
        GoTo 90
   End If

   UREPRE& = ULTREG&("PRESREPA")
   HOII% = HOY(HOS$)
   TEXT21(0) = "0"
   JJ& = 0: ATX% = 0
   For KKI% = 1 To List2.ListCount
     If List2.Selected(KKI% - 1) = True Then
       REPRESU& = XVALO(Mid$(List2.LIST(KKI% - 1), 121, 8))
       If REPRESU& > 0 Then
         If REPRESU& <= UREPRE& Then
           XX$ = REGLEIDO$("PRESREPA", REPRESU&)
           ZZ$ = REGBLAN$("!CARDETPE")
           CIXI% = CVI(Mid$(XX$, 69, 2))
           Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
           PRELI = CVS(Mid(XX$, 117, 4))
           'DESCUTEX$ = TRIM$(Mid$(XX$, 75, 14))
           CPPP = 1  ' (100 - COEFDESCU(DESCUTEX$)) / 100
           'If CPPP < 1 Then ATX% = 1
           '
           MONEMA% = 1   ' REPARACIONES SIEMPRE EN PESOS
           MONEMI% = 1
           MONEMIS = DEMONECO$(MONEMI%)
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           '
           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 39, 4)
           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 51, 4)
           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 65, 4)
           Call REPLA(ZZ$, MKS$(1), 47, 4)   ' CANTIDAD
           DESCUTEX$ = TRIM$(Mid$(XX$, 75, 14))
           'Call REPLA(ZZ$, MKS$(COEFDESCU(DESCUTEX$)), 69, 4)
           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 73, 4)
           Call REPLA(ZZ$, MKI$(HOII%), 63, 2)
           Call REPLA(ZZ$, MKS$((-1) * REPRESU&), 125, 4)
           JJ& = JJ& + 1
           Call GRAREG("!CARDETPE", ZZ$, JJ&)
         End If
       End If
     End If
   Next KKI%
   Call SETULTREG("!CARDETPE", JJ&)
   Call CIERRARCH("!CARDETPE")
   '
90 MARSELPRESU.Visible = False
   '
   If ATX% > 0 Then
      PORDESCU.TEXT = ""
      Call COMUNI("ATENCION: Descuento Aplicado según Presupuesto Previo !!!")
   End If
   
   PORCODBARRA = 0
   EPAC$ = UCase$(TRIM$(EMPREAC$))
   If ULTREG&("!CARDETPE") < 1 And InStr(EPAC$, "GABAL") > 0 Then
      If COMALTER%("Opciones de Carga de Pedido\\En Forma Manual\Por Código de Barras") > 1 Then
         PORCODBARRA = 1
         Call CARGAPORCODBAR
         Exit Sub
      End If
   End If
   Call CARDETPEDI(1)
   '
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_LostFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub


Private Sub Command20_Click()
   Static ARGU$
   Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     PUNCLI$ = ""
     Call COPYSTRU("DEUDOR1", "SELECLI")
     '
     Call ABRECLIEN
     
     SQLX$ = "SELECT  sysobjects.name AS table_name, syscolumns.name AS column_name,"
     SQLX$ = SQLX$ & " systypes.name AS datatype, syscolumns.LENGTH AS LENGTH"
     SQLX$ = SQLX$ & " FROM sysobjects INNER JOIN syscolumns ON sysobjects.id = syscolumns.id INNER JOIN"
     SQLX$ = SQLX$ & " systypes ON syscolumns.xtype = systypes.xtype where sysobjects.name='deudor12'"
     Dim RST As ADODB.Recordset
     Set RST = FLEXCONN.Execute(SQLX$)
     CONDI$ = ""
     With RST
        Do While Not .EOF
            CONDI$ = CONDI$ & !column_name & " like '%" & ARGU$ & "%' or "
            .MoveNext
        Loop
     End With
     On Error Resume Next
     RST.Close
     Set RST = Nothing
     On Error GoTo 0
     CONDI$ = Mid$(CONDI$, 1, Len(CONDI$) - 4)
     '
     Call CARGALISEL("BUSCLIEN", "deudor12", "nume_cli/I6;raso_cli/A48", CONDI$)
     Call SELECHO("BUSCLIEN/Lista de clientes que coinciden con '" & ARGU$ & "'")
     '
     Screen.MousePointer = 1
     NCLIX$ = VALACT1$("BUSCLIEN")
     If NCLIX$ <> "" Then
        Text1 = NCLIX$
        Call CARDACLI
     End If
  End If
'     Dim Deudor12 As ADODB.Recordset
'     Set Deudor12 = New ADODB.Recordset
'     Deudor12.CursorLocation = adUseClient
'     strsql = "Select * from Deudor12 "
'25   On Error Resume Next
'     Deudor12.Open FILTSQL$(strsql), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'        On Error GoTo 0
'        Call CapturaErrorOpen
'        GoTo 25
'     End If
'     On Error GoTo 0
'
'     With Deudor12
'      On Error Resume Next
'       .MoveFirst
'       If Err < 1 Then
'         On Error GoTo 0
'         Do While Not .EOF
'           For KKI% = 1 To .Fields.Count
'             On Error Resume Next
'             VATEX$ = ""
'             VATEX$ = .Fields(KKI% - 1).Value
'             On Error GoTo 0
'             If InStr(VATEX$, ARGU$) > 0 Then
'               NCLIE% = !nume_cli
'               RACLI$ = !Raso_Cli
'               XX$ = MKI$(NCLIE%) + RACLI$
'               JJ& = JJ& + 1
'               Call GRAREG("!SELECLI", XX$, JJ&)
'               Exit For
'             End If
'           Next KKI%
'         .MoveNext
'         Loop
'       End If
'     End With
'     Deudor12.Close
'     Set Deudor12 = Nothing
'
'     Call SETULTREG("!SELECLI", JJ&)
'     '
'     Screen.MousePointer = 1
'     If JJ& < 1 Then
'         Call MENSERR(24, "No se Encontraron Coincidencias")
'         GoTo 10
'       ElseIf JJ& = 1 Then
'         Text1 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELECLI", 1), 2))))
'         Call CARDACLI
'       Else
'         Call FRESHECHO("!SELECLI")
'         Call SELECHO("!SELECLI")
'         NCLIX$ = VALACT1$("!SELECLI")
'         If NCLIX$ <> "" Then
'           Text1 = NCLIX$
'           Call CARDACLI
'         End If
'     End If
'   End If
   Command20.Enabled = True
End Sub

Private Sub Command22_Click()
    '
    Command22.Enabled = False
    Screen.MousePointer = 11
    '
    'SE COMENTAN LAS LINEAS SIGUIENTES DEBIDO A QUE EL INDICE DE PEDIDOS (INDIPEDI)
    'SE ACTUALIZA CADA VEZ QUE SE EMITE UN NUEVO PEDIDO MEDIANTE UN CARGALISEL
    '
'    SQLX$ = "SELECT NroPed, Referencia, Status FROM PEDENCA "
'    SQLX$ = SQLX$ + " WHERE NROPED > 0 ORDER BY NroPed DESC"
'    '
'    Dim PEDENCTEMP As ADODB.Recordset
'    Set PEDENCTEMP = New ADODB.Recordset
'25  On Error Resume Next
'    PEDENCTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'       On Error GoTo 0
'       Call CapturaErrorOpen
'       GoTo 25
'    End If
'    On Error GoTo 0
'    '
'    TTXX$ = REGLEIDO$("INDIPEDI", 1)
'    ULPEDIN& = CVS(Left$(TTXX$, 4))
'    JJ& = 0
'    With PEDENCTEMP
'      On Error Resume Next
'      .MoveLast
'      If Err < 1 Then
'        FIN& = .RecordCount
'        .MoveFirst
'        If ULTREG&("INDIPEDI") = FIN& Then
'          If XVALO(!NROPED) = ULPEDIN& Then
'            GoTo 80 ' NO HAY NUEVOS PEDIDOS
'          End If
'        End If
'        JJ& = 0
'        Screen.MousePointer = 11
'        Do While Not .EOF
'           NRODES& = XVALO(!NROPED)
'           ANULX$ = "": If XVALO(!Status) = 9 Then ANULX$ = "(Anul)"
'           TTXX$ = MKS$(NRODES&)
'           TTXX$ = TTXX$ + ANULX$ + SAFETEXT$(!Referencia)
'           JJ& = JJ& + 1
'           Call GRAREG("#INDIPEDI", TTXX$, JJ&)
'           Call TRACE(20, JJ&, FIN&)
'        .MoveNext
'        Loop
'      End If
'    End With
'    Call FILEDUMP("INDIPEDI")
'    '
'80  PEDENCTEMP.Close
'    Set PEDENCTEMP = Nothing
    '
    Screen.MousePointer = 1
    CONPED07.Show 1
    '
    Command22.Enabled = True
    '
End Sub


Private Sub Command23_Click()
    '
    Call SELECHO("INDIPROS/Indice de Prospectos") ' PERMITE CARGA PRESUPUESTO A PARTIR DE PROSPECTOS
    NCLIE$ = TRIM$(VALACT1$("INDIPROS"))
    If NCLIE$ <> "" Then
      Label26.Caption = TRIM$(NCLIE$)
      NC& = XVALO(VALOBADA("DEUDOR12", "Nume_Cli", "Codi_Cli = '" & NCLIE$ & "'"))
      If NC& = 0 Then ' SI EL CLIENTE NO EXISTE LO DA DE ALTA
        Call ALTAPROSCLI(NCLIE$)
        NC& = XVALO(VALOBADA("DEUDOR12", "Nume_Cli", "Codi_Cli = '" & NCLIE$ & "'"))
        If NC& = 0 Then
          'SE COMENTA LA SIGUIENTE LINEA PORQUE EN ALTAPROSCLI EXISTE UNA VALIDACION RESPECTO A SI EL ALTA NO FUE SATISFACTORIA
          'EL NC& = 0 SE ASUME COMO UNA CANCELACION POR PARTE DEL USUARIO
          'Call MENSERR(24, "Problemas Al Dar de Alta Prospecto '" & NCLIE$ & "'")
          Exit Sub
        End If
      End If
      Text1 = TRIM$(FORMATNUM$(NC&, "I5"))
      Call CARDACLI
    End If
End Sub

Private Sub Command24_Click()
  '\\\OT-6-422-FG-05/10/06
'''   Call SELECHO("TRANSPOR")
'''   NTRAN% = XVALO(VALACT1$("TRANSPOR"))
'''   Text19 = TRIM$(FORMATNUM(NTRAN%, "I4"))
   Call COPYSTRU("TRANSPOR", "TRANSPO1")
   Call CARGALISEL("TRANSPO1", "TRANSPORTES", "CODI_TRA/F5;DENO_TRA/A32", "CODI_TRA > 0 ORDER BY DENO_TRA", "NOMESS")
   Call SELECHO("TRANSPO1")
   NTRAN% = XVALO(VALACT1$("TRANSPO1"))
   If NTRAN% > 0 Then
      Text19 = TRIM$(FORMATNUM(NTRAN%, "I4"))
   End If
  '\\\OT-6-422-FG-FIN
End Sub

Private Sub Command25_Click()
    ' Lista Generada de tabla SQL
    CONDI$ = "NUME_CLI = " & XVALO(Text1)
    Call CARGALISEL("SUCENTRE", "SUCENTRE", "NUME_SUC/f10;DENO_SUC/A24", CONDI$, "NOMESS")
    Call SELECHO("SUCENTRE")
    
    VDEVU$ = VALACT1$("SUCENTRE")
    Text20 = VDEVU$
'    CODIGO ANTERIOR
'    Call SELECHO("!SUCENTRE")
'
'    VDEVU$ = VALACT1$("!SUCENTRE")
'    Text20 = VDEVU$

End Sub

Private Sub Command26_Click()
    Call ABRECONEXION
    Dim CLACT%(32767)
    'TA$ = "MASUMCLI": CA$ = "Nume_Cli": GoSub 100
    ta$ = "CAJABANC": ca$ = "NUCliEN": GoSub 100
    'TA$ = "ENCAREMI": CA$ = "Nume_Clie": GoSub 100
    ta$ = "APLICOBRA": ca$ = "NuClien": GoSub 100
    'TA$ = "SUCENTRE": CA$ = "Nume_Cli": GoSub 100
    ta$ = "SACREPEN": ca$ = "NuClien": GoSub 100
    ta$ = "SADEUPEN": ca$ = "NuClien": GoSub 100
    'TA$ = "REPARA": CA$ = "Nume_Cli": GoSub 100
    'TA$ = "ORDETRAB": CA$ = "NCLIE_OTRAB": GoSub 100
    ta$ = "MOVIREPA": ca$ = "Nume_CliE": GoSub 100
    ta$ = "MOVISTO": ca$ = "Nume_CliE": GoSub 100
    'TA$ = "MOVIVIE": CA$ = "Nume_CliE": GoSub 100
    ta$ = "PEDENCA": ca$ = "NroCliE": GoSub 100
    ta$ = "PEDDETA": ca$ = "NroCliE": GoSub 100
    ta$ = "PRESUVEN": ca$ = "NroCliE": GoSub 100
    ta$ = "DEUDOR12": ca$ = "Nume_Cli": GoSub 100
    '
    JJ& = 0
    For I& = 1 To 32767
      If CLACT%(I&) < 1 Then
        JJ& = JJ& + 1
      End If
    Next I&
    
    MsgBox "Lugares Libres " & JJ&
Stop

    
100 Dim CLIMO As ADODB.Recordset
    Set CLIMO = New ADODB.Recordset
    SQLX$ = "SELECT " + ca$ + " AS NCX FROM " + ta$ + " WHERE " + ca$ + " > 0"
    Set CLIMO = FLEXCONN.Execute(SQLX$)
    With CLIMO
       Do While Not .EOF
         NCY& = XVALO(!NCX)
         If NCY& > 0 Then
           If NCY& <= 32767 Then
             CLACT%(NCY&) = 1
           End If
         End If
       .MoveNext
       Loop
    End With
    CLIMO.Close
    Set CLIMO = Nothing
Return
'
End Sub

Private Sub Command27_Click()
    'BOTON DE SELECCION DE MOTIVO
    Command27.Enabled = False
    '
'    'CARGO TABLA DE MOTIVOS SI NO ESTABA CARGADA PREVIAMENTE
'    If ULTREG&("TAMOTPRE") < 6 Then
'        Call GRAREG("TAMOTPRE", Chr$(0) + "Ventas", 1)
'        Call GRAREG("TAMOTPRE", Chr$(1) + "Giras/Congresos", 2)
'        Call GRAREG("TAMOTPRE", Chr$(2) + "Demostracion", 3)
'        Call GRAREG("TAMOTPRE", Chr$(3) + "Prestamo", 4)
'        Call GRAREG("TAMOTPRE", Chr$(4) + "Cambio", 5)
'        Call GRAREG("TAMOTPRE", Chr$(5) + "Donaciones", 6)
'        Call CIERRARCH("TAMOTPRE")
'    End If
    '
    'MUESTRO VENTANA DE SELECCION DE TABLA DE MOTIVOS
    Call SELECHO("TAMOTPRE")
    NUMOT% = XVALO(VALACT1$("TAMOTPRE"))
    'VALIDO LOS DATOS PARA SABER SI ELIGIO SALIR O ELIGIO 0 QUE ES VENTAS
    DESCRIMOT$ = TRIM$(VALACT2$("TAMOTPRE"))
    If NUMOT% <= 0 And DESCRIMOT$ = "" Then GoTo 99
    '
    Text16 = Str$(NUMOT%)
    '
99  Command27.Enabled = True
End Sub

Private Sub Command28_Click()
    Command28.Enabled = False
    If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Command28.Enabled = True
      Exit Sub
    End If
    Call SELECHO("SELFECHA")
    VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
    If VDEVU$ <> "" Then
      Text17.TEXT = VALACT1$("SELFECHA")
    End If
    Command28.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP01")
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
    Call CIERRARCH("*.*")
    If (USNBR% Mod 100) <> 1 Then
      Call COMUNI("La Opción Elegida Requiere\Privilegios de Supervisor.")
      Exit Sub
    End If
    CFGPED07.Show 1
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   MARCONOTA.Visible = False
   Call PEDIPEND
   Call SELPRESUPEN
End Sub
'
Sub SELPRESUPEN()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   'UREPRE& = ULTREG&("PRESUPU")
   Screen.MousePointer = 11
   List2.Clear
   NCLIEN = XVALO(Text1.TEXT)
   If UCase$(Left$(Label11, 6)) = "REPARA" Then GoTo 50
   '
   SQLX$ = "SELECT * FROM PRESUVEN WITH(NOLOCK) "
   SQLX$ = SQLX$ + "WHERE NroClie = " & NCLIEN & " "
   SQLX$ = SQLX$ + "AND NroPedi < 1 "
   SQLX$ = SQLX$ + "ORDER BY NroPres DESC"
   '
   NPX& = 0
   Call ABREPRESVEN
   'Set IPRESPEN = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   'jandy sql
   Dim IPRESPEN As ADODB.Recordset
   Set IPRESPEN = New ADODB.Recordset
   IPRESPEN.CursorLocation = adUseClient
25 On Error Resume Next
   IPRESPEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   With IPRESPEN
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       Do While Not .EOF
         If IsNull(!MarNoVig) = True Or !MarNoVig = "" Then
           NRO& = !NroPres
             TTXX$ = "Presup.No: " + TRIM$(Str$(NRO&))
           FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + " - " + FECHATEX$(FEX)
           If TRIM$(!Referen) <> "" Then
             TTXX$ = TTXX$ + " - Ref: " + TRIM$(!Referen)
           End If
          'End If
           CIXI% = !CODIINT
           '
           If NRO& = NPX& Then
             TTXX$ = TTXX$ + " - VARIOS"
             List2.RemoveItem (List2.ListCount - 1)
             GoTo 15
           End If
           '
           NPX& = NRO&
           CAPRESU& = !Cantit
           If CAPRESU& < 1 Then CAPRESU& = 1
           TTXX$ = TTXX$ + " - " + TRIM$(CODEXT$(CIXI%)) + " (" + TRIM$(Str$(CAPRESU&)) + ") " + DESCRIT0$(CIXI%)
           '
15         TTXX$ = AJUSTI$(TTXX$, 256)
           Call REPLA(TTXX$, TRIM$(Str$(NRO&)), 251, 6)
           List2.AddItem TTXX$
         End If
         '
19     .MoveNext
       Loop
     End If
   End With
   IPRESPEN.Close
   Set IPRESPEN = Nothing
   GoTo 90
   '
50 If InStr(1, EMPREAC$, "ARCON") > 0 Then
        Call CARGAREPARA
        GoTo 90
   End If

   UREPRE& = ULTREG&("PRESREPA")
   Screen.MousePointer = 11
   RFF$ = RECFILT$("PRESREPA", 4, MKI$(NCLIEN))
   For KKI% = 1 To Len(RFF$) Step 4
      REPRESU& = CVS(Mid$(RFF$, KKI%, 4))
      If REPRESU& > 0 Then
        If REPRESU& <= UREPRE& Then
          XX$ = REGLEIDO$("PRESREPA", REPRESU&)
          If CVI(Mid$(XX$, 121, 2)) > 0 Then
            If Abs(CVS(Mid$(XX$, 125, 4))) < 0.5 Then
              TTXX$ = "Reparación No: " + TRIM$(CSTRING$(Left$(XX$, 4), 3, 8, 0, 2))
              TTXX$ = TTXX$ + " - " + Mid$(XX$, 5, 60)
              TTXX$ = AJUSTI$(TTXX$, 128)
              Call REPLA(TTXX$, Str$(REPRESU&), 121, 8)
              List2.AddItem TTXX$
            End If
          End If
        End If
      End If
   Next KKI%
   '
90 Screen.MousePointer = 1
   '
   If List2.ListCount < 1 Then
      Call SETULTREG("!CARDETPE", 0)
      Call CIERRARCH("!CARDETPE")
      EPAC$ = UCase$(TRIM$(EMPREAC$))
      If InStr(EPAC$, "GABAL") > 0 Then
         If COMALTER%("Opciones de Carga de Pedido\\En Forma Manual\Por Código de Barras") > 1 Then
            PORCODBARRA = 1
            Call CARGAPORCODBAR
            Exit Sub
         End If
      End If
      PORCODBARRA = 0
      Call CARDETPEDI(1)

      Exit Sub
   End If

   '
99 MARSELPRESU.Visible = True
   Command19.SetFocus
   '
End Sub
Sub CARGAPORCODBAR()

        PEDCOBAR07.List1.Clear
        PEDCOBAR07.Show 1
        '
        If XVALO(PEDCOBAR07.List1.ListCount) > 0 Then
           PEDCOBAR07.List1.Clear
           Call CARDETPEDI(2)
        End If

End Sub
 Private Sub Command7_Click()
    Command7.Enabled = False
    Call CIERRARCH("*.*")
    MODX% = ALTERNA%("Listas de Precios\Condiciones de Pago")
    If MODX% = 1 Then
         Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
       ElseIf MODX% = 2 Then
         Call CARCONPAGOS
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "ARCHIG07/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text9.TEXT)
End Sub



Private Sub ConGral_Click()
Call Command5_Click
End Sub

Private Sub ConRimp_Click()
Call Command12_Click
End Sub

Private Sub CoPag_Click()
MODO% = 2
Call MODOCON(MODO%)
End Sub

Private Sub DeNumPe_Click()
Call OPPED07.Command63_Click
End Sub

Private Sub descargarocsuc_Click()

   Call CREAR_TABLA_OC_DE_SUCURSAL
   Call CREAR_TABLA_MENSAJES_DE_SUCURSAL
   '
   Call DESCARGAR_ARCHIVO_COMPRAS_SUCURSAL("Ocompra_Sucursal")
End Sub
Sub DESCARGAR_ARCHIVO_COMPRAS_SUCURSAL(archivo$)

        CLIENTE11 = CONTROL("RETRASLA", "SUC1CLIE")
        CLIENTE22 = CONTROL("RETRASLA", "SUC2CLIE")
        If XVALO(CLIENTE11) + XVALO(CLIENTE22) < 1 Then
           Call COMUNI("Falta Configurar Rutas de Descarga de Archivos de Sucursal")
           Exit Sub
        End If
        
        OPX% = COMALTER%("Seleccione Sucursal\\" & Left$(TRIM$(RASOCLI$(XVALO(CLIENTE11))), 16) & "\" & Left$(TRIM$(RASOCLI$(XVALO(CLIENTE22))), 16) & "")
        If OPX% < 1 Or OPX% > 2 Then Exit Sub
        If OPX% = 1 Then CLIENTESEL = XVALO(CLIENTE1): Ruta$ = CONTROL("RETRASLA", "SUC1RUTA")
        If OPX% = 2 Then CLIENTESEL = XVALO(CLIENTE2): Ruta$ = CONTROL("RETRASLA", "SUC2RUTA")

        'RUTA$ = TRIM$(CONTROL("", "CARPEINT"))
        If Ruta$ <> "" Then
            'ASIGNO PROPIEDADES AL FILE1
                On Error Resume Next
                File1.Path = "C:\"
                On Error GoTo 0
                File1.Path = Ruta$
                File1.Pattern = "*.txt"
                TX$ = ""
                List3.Clear
                'RECORRO EL FILE1 Y LE PONGO LA FECHA DEL ARCHIVO POR DELANTE PARA QUE ORDENE CORRECTAMENTE
                For n = 0 To File1.ListCount - 1
                  NombreArchivo$ = TRIM$(Left$(File1.LIST(n), Len(File1.LIST(n)) - 4))
                  FechaArchivo$ = TRIM$(FileDateTime(Ruta$ + "\" + NombreArchivo$ + ".txt"))
                  Call REPLA(TX$, AJUSTI$(FechaArchivo$ + " | ", 48), 1, 48)
                  Call REPLA(TX$, AJUSTI$(NombreArchivo$, 48), 50, 48)
                  List3.AddItem TX$
                Next n
                '
                'SEGUN EL ARCHIVO TOMA LA TABLA CORRESPONDIENTE
                If InStr(archivo$, "Ocompra_Sucursal") > 0 Then TABLAX$ = "OC_SUCURSAL"
                
                '
                'GRABA CADA ARCHIVO EN SQL Y SI SI GRABA CORRECTAMENTE SE ELIMINA.
                For I& = 0 To List3.ListCount - 1
                  POS1% = InStr(TX$, "|")
                  FechaArchivo$ = Mid$(List3.LIST(I&), 1, POS1% - 2)
                  NombreArchivo$ = TRIM$(Mid$(List3.LIST(I&), POS1% + 2))
                  If UCase("Ocompra_Sucursal") = Left$(TRIM$(UCase$(NombreArchivo$)), Len(archivo$)) Then
                    TABLAX$ = "OC_SUCURSAL"
                    Call GrabarArchivo(Ruta$ + "\" + NombreArchivo$ + ".txt", TABLAX$) 'GRABA EN LA BASE DE DATOS SQL
                  End If
                Next I&

        Else
            Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio")
            Exit Sub
        End If
        Call COMUNI("Proceso Terminado")
        '
End Sub
Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub


Private Sub EncaPe_Click()
Call OPPED07.Command61_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  While WindowState <> 0
    WindowState = 0
    Openforms = DoEvents
    AppActivate Caption
    SendKeys "%" & Chr$(32) & Chr$(13), True
  Wend
End Sub

Private Sub Form_Load()
    '
    Frame4.Height = 1440
    Openforms = DoEvents
    VERANTER$ = "PEDIDO01"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
    If TERMINA% > 0 Then Call FINAL("")
    '
    CMO$ = TRIM$(CONTROL$("", "CODPECLI"))
    If CMO$ <> "" Then Label10.Caption = CMO$

    Call TRANSLABELS(Name)
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture2)
    DoEvents
'    EPAC$ = TRIM$(EMPREAC$)
'    If EPAC$ <> "" Then
'        Caption = Caption + "  -  " + EPAC$
'    End If
    '
    EPAC$ = UCase$(TRIM$(EMPREAC$))
     If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    '
    If InStr(EPAC$, "PREMOTEC") > 0 Then
      Proyepedidos.Visible = True
    End If
    '
    Label11.Visible = True
    'If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "MEDITEA") > 0 Or InStr(EPAC$, "GABAL") > 0 Or InStr(EPAC$, "ARCON") > 0 Or InStr(EPAC$, "FERVI") > 0 Or InStr(EPAC$, "TANIT") Or InStr(EPAC$, "BICI") > 0 Or InStr(EPAC$, "AMBEST") > 0 Or InStr(EPAC$, "EYM") > 0 Or InStr(EPAC$, "AHP") > 0 Then
        Command22.Visible = True
   ' End If
   
    If InStr(EPAC$, "ARCON") > 0 Then
       pedidossuc.Visible = True
    End If
    
    If CONTROL$("ADMIPED", "MOCOMIS") = "LETRA" Then
       Label21(4) = "Letra:"
    End If
    '
       
    'DERECHO DE MODIFICACION DE DESCUENTO EN PANTALLA
    If DERACCE%("MODIDESC/NOMESS", "MODIFICAR DESCUENTO EN PANTALLA") < 2 Then
       PORDESCU.Enabled = False
    End If
    '
    Combo1.Clear
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
    '
    Call DEPURAPED
    Show
    Openforms = DoEvents
    '
    ORIPED$ = TRIM$(UCase$(CONTROL$("PEDIDO", "ORIDATOS")))
    If InStr(ORIPED$, "ORDENEMPAQUE.MDB") > 0 Then
       Image5.Visible = True
       Image6.Visible = False
    End If
    '
    If EXISTE%("REPARA02.EXE") > 0 Then
      Label11.Visible = True
    End If
    Call GRATITFOR(Caption)
    '
    If IDENTER$ <> "" Then
      If CONTROL$("CAPRINT", IDENTER$) = "SI" Then
        Timer2.Enabled = True
      End If
    End If
    '
    If EXISTE%("REPEDREM.EXE") > 0 Then
      Image2.Visible = True
    End If
    '
    
    'CARGO TABLA DE MOTIVOS SI NO ESTABA CARGADA PREVIAMENTE
    If EXISTE(LUDAT$ + "\TAMOTPRE.RFS") > 0 Then
        If ULTREG&("TAMOTPRE") < 6 Then
            Call GRAREG("TAMOTPRE", Chr$(0) + "Ventas", 1)
            Call GRAREG("TAMOTPRE", Chr$(1) + "Giras/Congresos", 2)
            Call GRAREG("TAMOTPRE", Chr$(2) + "Demostracion", 3)
            Call GRAREG("TAMOTPRE", Chr$(3) + "Prestamo", 4)
            Call GRAREG("TAMOTPRE", Chr$(4) + "Cambio", 5)
            Call GRAREG("TAMOTPRE", Chr$(5) + "Donaciones", 6)
            Call CIERRARCH("TAMOTPRE")
        End If
    End If
    '
    'PARA MEDITEA MUESTRA BOTON DE SELECCION DE MOTIVO DEL PEDIDO
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "MEDITEA") > 0 Then
       If EXISTE(LUDAT$ + "\TAMOTPRE.RFS") > 0 Then
         Picture6.Visible = True
         Text16 = "0"
       End If
    End If
    '
    MONEMI% = 1  ' PESOS
    MONEMIS.Caption = DEMONECO(MONEMI%)
    MONEMIS.Visible = True
    CONVERPRECIO% = 1
    TICAMBIS.Visible = True
    Label12.Visible = True
    Image1.Visible = True
    TICAMBIS = FORMATNUM$(TICAMONE(1), "F9.4")
    '
    ECOPRECIO% = 0
    If CONTROL$("ADMIPED", "MUEPREC") <> "SI" Then
      For KKI% = 0 To 2      ' TAPA IMPORTES TOTALES
        Text14(KKI%).Visible = True
      Next KKI%
      ECOPRECIO% = (-1)
    End If
    '
    If ECOPRECIO% < 0 Then
      Label22.Caption = "  Cant Código          Descripción                                           Entrega "
      Image1.Visible = False
    End If
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    Text17.TEXT = HOS$
    Text11.TEXT = TRIM$(Str$(NUPEDIDO&))
    '
    DoEvents
    '
    If EXISTE%("SEGCOM07.EXE") > 0 Or INTERPRE% = 1 Then
        Command23.Visible = True ' Si el sistema tiene la aplicacion seguimiento comercial muestra el botón
    End If
    '
    Call SETULTREG("!CARDETPE", 0)
    '
    Call Verifica_Impresora_Despacho
    ' RECUPERA PEDIDOS CON STATUS >= 100
    strsql = "SELECT * FROM PEDDETA "
    strsql = strsql & " WHERE sTATUS >=100 "
    Dim PEDDETTEMP As ADODB.Recordset
    Set PEDDETTEMP = New ADODB.Recordset
27  On Error Resume Next
    PEDDETTEMP.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 27
    End If
    On Error GoTo 0
    '
    With PEDDETTEMP
       Do While Not .EOF
         If !Status >= 100 Then
            STATUI% = !Status - 100
            !Status = STATUI%
            .Update
         End If
       .MoveNext
       Loop
    End With
    PEDDETTEMP.Close
    Set PEDDETEMP = Nothing
    '
    If Mid$(APLINVO$, 1, 9) = "NUEVAOVEN" Then
        If CODCARPE$ <> "" Then
           CONDI$ = "CodcarpImp ='" & CODCARPE$ & "' AND STATUS < 9 AND MODO = 1"
           If CantRegistros("CARPEIMP", CONDI$) > 0 Then
              NCLI% = XVALO(VALOBADA("CARPEIMP", "NClien_Oven", CONDI$, "NOMESS"))
              Text1 = NCLI%
              Label30 = CODCARPE$
              Command1.Enabled = False
              Text1.Enabled = False
              Label26.Enabled = False
              Call CARDACLI
           End If
        End If
    End If
    '
End Sub

Private Sub LENTREGA_DblClick(Index As Integer)
    If Index = 0 Then
       CAMPO$ = "Domicilio de Entrega": CAMPOFILTRO$ = "DomiEnt"
    ElseIf Index = 1 Then
       CAMPO$ = "Localidad de Entrega": CAMPOFILTRO$ = "LocaEnt"
    ElseIf Index = 2 And XVALO(Text19) < 1 Then: CAMPOFILTRO$ = "TranspEnt"
        CAMPO$ = "Transporte"
    Else
        Exit Sub
    End If
    
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPO$ & "/A99;CONSERPARADOR/A0", FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPO$ & "/A99;CONSERPARADOR/A0", FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1100
    FRMZELECHO.Caption = CAMPO$
    FRMZELECHO.MSF1.Rows = 1
    
    SQLX$ = "SELECT DISTINCT  " & CAMPOFILTRO$ & " AS VALOR1"
    If Index < 2 Then SQLX$ = SQLX$ + " , LocaEnt  AS VALOR2"
    SQLX$ = SQLX$ + " FROM PEDENCA WITH(NOLOCK) WHERE NroClie = " & XVALO(Text1)
    SQLX$ = SQLX$ + " ORDER BY " & CAMPOFILTRO$
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        I& = I& + 1
        FRMZELECHO.msf2.Rows = I& + 1
        FRMZELECHO.msf2.TextMatrix(I&, 1) = SAFETEXT$(!VALOR1)
        If Index <= 1 Then
          FRMZELECHO.msf2.TextMatrix(I&, 1) = SAFETEXT$(!VALOR1) & " " & SAFETEXT$(!VALOR2)
          If TRIM$(SAFETEXT$(!VALOR2)) <> "" Then
             FRMZELECHO.msf2.TextMatrix(I&, 2) = SAFETEXT$(!VALOR1) & "|" & SAFETEXT$(!VALOR2)
          End If
        End If
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    FRMZELECHO.Show 1
    LECTURAX$ = TRIM$(RESP_ZELE_VECT(2))
    If LECTURAX$ = "" Then LECTURAX$ = TRIM$(RESP_ZELE_VECT(1))
    If LECTURAX$ <> "" Then
     If Index <= 1 Then
       Call TEXTOLOTES$(LECTURAX$, "|")
       On Error Resume Next
       LECTURA1$ = CADENATEX$(1)
       LECTURA2$ = CADENATEX$(2)
       On Error GoTo 0
       If Index = 0 Then LENTREGA(0).TEXT = Left$(LECTURA1$, 48)
       LENTREGA(1).TEXT = Left$(LECTURA2$, 48)
     Else
       LENTREGA(Index).TEXT = Left$(LECTURAX$, 48)
     End If
    End If
    
End Sub

Private Sub mnusitstock_Click()
  SITUSTO.Show 1
End Sub

Private Sub mnuVerSituacionpedidoproyeccion_Click()
   DIAGPREM.Show 1
End Sub

Private Sub Option1_Click()
   If Option1.Value = True Then
     TXTFLETE = "Flete A Cargo del Cliente"
   Else
    TXTFLETE = ""
   End If

End Sub

Private Sub Option2_Click()
  If Option2.Value = True Then
     EMPRESAX$ = EMPREAC$
     If Len(EMPREAC$) > 6 Then
       POS1% = InStr(EMPREAC$, " ")
       EMPRESAX$ = Mid$(EMPREAC$, 1, POS1%)
     End If
     TXTFLETE.TEXT = "Flete A Cargo de " & EMPRESAX$
   Else
     TXTFLETE = ""
   End If

End Sub


Private Sub pedistzona_Click()
    CAMPOS$ = "Nro Ped./F10;Cliente/F7;R.Social/A20;Codint/F0;Código/F16;Descripción/A32;Cant.Ped./f10.1;Cant.Ent/F10.1;Zona/F0;Zona/A18;Saldo/F10.1;Stock/F10.1;F.Emis./D8;F.Entrega/D8;FEMISNUM/F0;FENTREGANUM/F0"
    Call CARGA_ENCABEZADO(MUESTRADATOS11.GRID, CAMPOS$, MUESTRADATOS11, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(MUESTRADATOS11.MSF_2, CAMPOS$, MUESTRADATOS11, 1, 0, ANTOT, , 1)
    MUESTRADATOS11.GRID.Rows = 1
    'PRESENTACION DE LOS COMPROBANTES QUE NO TIENEN CUIT
   SQLX$ = " SELECT * FROM PEDDETA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " INNER JOIN PEDENCA WITH(NOLOCK) ON PEDDETA.NroPed = PEDENCA.NroPed"
   SQLX$ = SQLX$ + " WHERE PEDENCA.STATUS < 9"
   SQLX$ = SQLX$ + " AND PEDDETA.STATUS < 8 "
   SQLX$ = SQLX$ + " AND PEDDETA.CANPED > PEDDETA.CANTENT "
   SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND PEDdeta.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " ORDER BY PEDDETA.NROPED,PEDDETA.NROCLIE, PEDDETA.FESOLENT"
   '
   Set RS = READSET(SQLX$)
   With RS
    Do While Not .EOF
        CI1% = XVALO(!CODIINT)
        NC = XVALO(!NROCLIE)
        NPE& = XVALO(!NROPED)
        CANPEX = XVALO(!CANPED)
        CANTENTX = XVALO(!CANTENT)
        SALDO = XVALO(!CANPED) - XVALO(!CANTENT): If SALDO < 0 Then SALDO = 0
        CODIGO$ = SAFETEXT$(!CODIEXT)
        DESCRI$ = SAFETEXT$(!Descrip)
        ZONAX$ = SAFETEXT$(ZONACLI$(NC))
        DESCRIZONA$ = ECOARCH$("ZONAREP", ZONAX$)
        ESTOCKS = FORMATNUM$(STOCKDIS(CI1%), "F10.1")
        FEEMI% = CVINT(!FECHEMIS)
        FESOL% = CVINT(!FESOLeNT)
        '
        j& = j& + 1
        MUESTRADATOS11.GRID.Rows = j& + 1
        MUESTRADATOS11.GRID.TextMatrix(j&, 1) = NPE&
        MUESTRADATOS11.GRID.TextMatrix(j&, 2) = NC
        MUESTRADATOS11.GRID.TextMatrix(j&, 3) = RASOCLI$(NC)
        MUESTRADATOS11.GRID.TextMatrix(j&, 4) = CI1%
        MUESTRADATOS11.GRID.TextMatrix(j&, 5) = CODIGO$
        MUESTRADATOS11.GRID.TextMatrix(j&, 6) = DESCRI$
        MUESTRADATOS11.GRID.TextMatrix(j&, 7) = FORMATNUM$(CANPEX, "F10.1")
            
        MUESTRADATOS11.GRID.TextMatrix(j&, 8) = FORMATNUM$(CANTENTX, "F10.1")
        MUESTRADATOS11.GRID.TextMatrix(j&, 9) = ZONAX$
        MUESTRADATOS11.GRID.TextMatrix(j&, 10) = DESCRIZONA$
        MUESTRADATOS11.GRID.TextMatrix(j&, 11) = FORMATNUM$(SALDO, "F10.1")
        MUESTRADATOS11.GRID.TextMatrix(j&, 12) = FORMATNUM$(ESTOCKS, "F10.1")

        MUESTRADATOS11.GRID.TextMatrix(j&, 13) = FECHATEX$(FEEMI%)
        MUESTRADATOS11.GRID.TextMatrix(j&, 14) = FECHATEX$(FESOL%)
        MUESTRADATOS11.GRID.TextMatrix(j&, 15) = FEEMI%
        MUESTRADATOS11.GRID.TextMatrix(j&, 16) = FESOL%
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    If j& > 0 Then
       Screen.MousePointer = 1
       MUESTRADATOS11.GRID.Visible = True
       MUESTRADATOS11.GRID.ZOrder 0
       MUESTRADATOS11.Caption = "Asignación de Cuenta Contable por Código"
       MUESTRADATOS11.Show 1
       
    Else
       Call COMUNI("No se Han Encontrado Pedidos Pendientes")
       Screen.MousePointer = 1
    End If

End Sub

Private Sub Proyepedidos_Click()
   DIAGPREM.Show 1
End Sub

'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call FINAL("")
End Sub









Function VALIDAR_OCOMP_GENERADA%(NROPEDIDO&)
    CONDI$ = "NroPED = " & NROPEDIDO&
    If CantRegistros&("PEDICOMPRA", CONDI$, "NOMESS") > 0 Then
            VALIDAR_OCOMP_GENERADA% = 1
            OPX% = COMALTER%("Este Pedido Ya Genero Solicitudes de Compras\Si Necesita Modificar:\Anule la-s Solicitud-es y Vuelva a Generar\\Ver Solicitudes Generadas\Cancelar")
            If OPX% < 1 Or OPX% > 1 Then Exit Function
            Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cta.Prov/F9;R.Social/A40;O.Compra/F9;Sol.Compra/F12", FRMZELECHO, , , ANTOT)
            Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Cta.Prov/F9;R.Social/A40;O.Compra/F9;Sol.Compra/F12", FRMZELECHO, , , ANTOT)
            If ANTOT < 5500 Then ANTOT = 5500
            FRMZELECHO.Width = ANTOT + 1100
            FRMZELECHO.Caption = "Lista de Ordenes de Compra - Pedido N°: " & NROPEDIDO&
            '
            FRMZELECHO.MSF1.Rows = 1
            SQLX$ = "SELECT DISTINCT(NroPed),Num_Prov, NumSolCompra FROM PediCOMPRA WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE NroPed = " & NROPEDIDO&
            SQLX$ = SQLX$ + " ORDER BY Num_Prov"
            Set RS = READSET(SQLX$)
            With RS
              Do While Not .EOF
                I& = I& + 1
                FRMZELECHO.msf2.Rows = I& + 1
                NP% = XVALO(!Num_Prov)
                FRMZELECHO.msf2.TextMatrix(I&, 1) = NP%
                FRMZELECHO.msf2.TextMatrix(I&, 2) = RASOCLI$(-1 * NP%)
                   NSOLIC& = XVALO(!NumSolCompra)
                   OCOMPRAX& = OrdenCompraSegunSolicitud&(NSOLIC&)
                FRMZELECHO.msf2.TextMatrix(I&, 3) = OCOMPRAX&
                If OCOMPRAX& > 0 Then Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.msf2, vbYellow, I&)
                FRMZELECHO.msf2.TextMatrix(I&, 4) = NSOLIC&
                .MoveNext
              Loop
            End With
            RS.Close
            Set RS = Nothing
            FRMZELECHO.Show 1
            
    End If
    '
End Function

Private Sub GRID2_Click()
   NGRID = GRID2.Row
   NLIST = List1.ListCount
   If (NLIST) >= NGRID - 1 Then
     List1.ListIndex = NGRID - 1
     List1_DblClick
   End If

End Sub

Private Sub Image1_Click()
    If ECOPRECIO >= 0 Then
      EPAC$ = TRIM$(UCase$(EMPREAC$))
        If Frame4.Height = 3120 Then
          Frame4.Height = 1440
          If InStr(EPAC$, "GABAL") > 0 Then
            Check2.Visible = False
            Label31.Visible = False
            Text141.Visible = False
            Ltota(3).Top = 1155
            Text14(2).Top = 1050
            Total(2).Top = 1050
          End If
        Else
          Frame4.Height = 3120
          If InStr(EPAC$, "GABAL") > 0 Then
            Check2.Visible = True
            Label31.Visible = True
            Text141.Visible = True
            Ltota(3).Top = 1575
            Text14(2).Top = 1500
            Total(2).Top = 1500
          End If
        End If
    End If


End Sub

Private Sub Image1_DblClick()
   Static CTXX%
   '
   If ULTREG&("DAPIBCLI") < 1 Then
     Call MENSERR(24, "Imposible desde esta Aplicación")
     Exit Sub
   End If
   '
   NCLIE = XVALO(TRIM$(Text1.TEXT))
   If NCLIE > 0 Then
     If RASOCLI$(NCLIE) <> "" Then
       '
       If CTXX% < 1 Then
         Call GRAREG("TACATIBC", Chr$(0) + "Exento Percepcion IB", 1)
         Call GRAREG("TACATIBC", Chr$(1) + "Sujeto Percepcion IB", 2)
         Call SETULTREG("TACATIBC", 2)
         CTXX% = 1
       End If
       '
       'RECORD$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE%))
       RECORD$ = FILTERGETRECORD$("DAPIBCLN", 1, MKI$(NCLIE))
       Call REPLA(RECORD$, MKI$(NCLIE), 1, 2)
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
         If ULTREG&("DAPIBCLN") < 1 Then
            Call REPLA(RECORD$, MKI$(NCLIE), 1, 2)
            Call REPLA(RECORD$, USERNAME$, 41, 20)
            Call REPLA(RECORD$, MKI$(HOY(HOS$)), 61, 2)
            Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 63, 2)
            Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIE), RECORD$)
            Call CIERRARCH("DAPIBCLI")
         Else
            '
            Call REPLA(RECORD$, MKS$(NCLIE), 1, 4)
            Call REPLA(RECORD$, USERNAME$, 42, 20)
            Call REPLA(RECORD$, MKI$(HOY(HOS$)), 62, 2)
            Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 64, 2)
            Call FILTERPUTRECORD("DAPIBCLN", 1, MKS$(NCLIE), RECORD$)
            Call CIERRARCH("DAPIBCLN")
         End If
         '
         Call CALPERIB
         '
       End If
     End If
   End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If ECOPRECIO >= 0 Then
'
'     Frame4.Height = 3120
'
'   End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Frame4.Height = 1440
End Sub

Private Sub Image2_DblClick()
   Call CIERRARCH("*.*")
   Call CONECRUN("REPEDREM", "Recepción Pedidos Remotos por Mail")
End Sub

Sub Image4_DblClick()
   PEDABIE07.Show 1
End Sub

Private Sub Image5_DblClick()
  '
  ORIPED$ = TRIM$(UCase$(CONTROL$("PEDIDO", "ORIDATOS")))
  If InStr(ORIPED$, "ORDENEMPAQUE.MDB") > 0 Then
     Call IMPORZET(ORIPED$)
     Exit Sub
  End If
  '
End Sub

Private Sub Image6_DblClick()
  '
  With IMPPED07
    .Label26 = Text1
    .Text2 = Text2
    .Text3 = Text3
    .Text4 = Text4
    .Text5 = Text5
    .Text12 = Text12
    .Text10 = Text10
    .Text9 = Text9
    .Label11 = XVALO(TEXT21(0)) 'LISTA DE PRECIOS
    .Show 1
  End With
End Sub

Private Sub Label11_DblClick()
   If TRIM$(UCase$(Label11)) = "VENTA" Then
       Label11.ForeColor = RGB(255, 0, 0)
       Label11 = "Reparación"
       Label11.ToolTipText = "Doble-Click para 'Venta'"
     Else
       Label11.ForeColor = RGB(0, 0, 255)
       Label11 = "Venta"
       Label11.ToolTipText = "Doble-Click para 'Reparación'"
   End If
End Sub

Private Sub Label25_Change()
   Call CALTOPED
End Sub

Private Sub Label26_DblClick()
   Call SELECHO("INDICLIE")
   CPROX$ = TRIM$(VALACT1$("INDICLIE"))
   If CPROX$ <> "" Then
     NCXI = NUMECLI(CPROX$)
     If NCXI > 0 Then
       If REGICLI&(NCXI) > 0 Then
         Label26 = CPROX$
         'NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NCXI))
         Call CARDACLI
         On Error Resume Next
         Command6.SetFocus
         On Error GoTo 0
       End If
     End If
   End If
End Sub

Private Sub Label31_Change()
    If ROUND#(Label31) < 100 Then
      Label31 = FORMATNUM(Label31, "F5.2")
      Call CALTOPED
    Else
      Call COMUNI("Descuento No Válido !!! Verifique")
      Label31 = ""
    End If
End Sub


Private Sub Label8_Click()
   On Error Resume Next
   Text13.SetFocus
   On Error GoTo 0
End Sub

Private Sub LENTREGA_GotFocus(Index As Integer)
   '\\\OT-6-422-FG-05/10/06
   If Index = 2 Then
     If XVALO(Text19) > 0 Then
       Call MENSERR(24, "Campo no Editable")
       On Error Resume Next
       Text19.SetFocus
       On Error GoTo 0
       Exit Sub ' LE AGREGUE ESTA LINEA PARA QUE NO SE EJECUTE EL RESTO DEL CODIGO
     End If
   End If
   '\\\OT-6-422-FG-FIN
   '
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   LENTREGA(Index).TEXT = TRIM$(LENTREGA(Index).TEXT)
   LENTREGA(Index).SelStart = 0
   LENTREGA(Index).SelLength = Len(LENTREGA(Index).TEXT)
End Sub

Private Sub LiPreVen_Click()
    Call CIERRARCH("*.*")
    Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
End Sub

Private Sub LiPrVe_Click()
    Call CIERRARCH("*.*")
    Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
End Sub

Private Sub LiPrVta_Click()
MODO% = 1
Call MODOCON(MODO%)
End Sub

Sub List1_DblClick()
   NOFRESHPRE% = 1
   If PORCODBARRA > 0 Then Call CARGALISTBARRA: Exit Sub
   
   Call CARDETPEDI(2)
   NOFRESHPRE% = 0
End Sub
Sub CARGALISTBARRA()

   PEDCOBAR07.List1.Clear
   FIN& = ULTREG&("!CARDETPE")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     TTXX$ = Space$(96)
     CIXI% = CVI(Left$(X$, 2))
     PRECARGA = CVS(Mid$(X$, 39, 4))
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     '
     CODI2$ = TRIM$(CODEXT$(CIXI%))
     Call REPLA(TTXX$, CODI2$, 1, 16)
     Call REPLA(TTXX$, Left(DESCRIT$(CIXI%), 30), 18, 36)
     Call REPLA(TTXX$, FORMATNUM$(CAN, "F10.1"), 57, 10)
     Call REPLA(TTXX$, FORMATNUM$(PRECARGA, "F14.4"), 69, 14)
     Call REPLA(TTXX$, FECHATEX$(FENTRESOL%), 90, 8)
     '
     PEDCOBAR07.List1.AddItem TTXX$
     
   Next U&
   PEDCOBAR07.Label6 = Total(2)
   PEDCOBAR07.Label2 = "OK"
   PEDCOBAR07.Show 1
   If PEDCOBAR07.Label2 <> "" Then Call CARDETPEDI(2)
   NOFRESHPRE% = 0
End Sub


Private Sub maeart_Click()
'\\\OT-5-869-28/12/05///
Call Command8_Click
'\\\OT-5-869-FIN///
End Sub

Private Sub maecli_Click()
'\\\OT-5-869-28/12/05///
Call Command11_Click
'\\\OT-5-869-FIN///
End Sub

Private Sub MaesArt_Click()
'\\\OT-5-869-28/12/05///
Call Command8_Click
'\\\OT-5-869-FIN///
End Sub

Private Sub MaesCli_Click()
'\\\OT-5-869-28/12/05///
Call Command11_Click
'\\\OT-5-869-///
End Sub

Private Sub MONEMIS_Change()
   MONEMIS.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
   End If
   On Error Resume Next
   Combo1.ListIndex = MONEMI% - 1
   Combo1.TEXT = MONEMIS
   If Err Then
      On Error GoTo 0
      Call Combo1_Click
   End If
   On Error GoTo 0
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
Exit Sub
   MONEMA% = MONEMI%
   If Left$(MONEMIS, 1) = "D" Then
        MONEMI% = 2
        MONEMIS.ForeColor = RGB(150, 0, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Pesos'"
        'Label12.Visible = True
        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = True
        TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
      Else
        MONEMI% = 1
        MONEMIS.ForeColor = RGB(0, 100, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Dólares'"
        'Label12.Visible = False
        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = False
   End If
   '
   COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
   FIN& = ULTREG&("!CARDETPE")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = COEFCON * CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     Call REPLA(X$, MKS$(PRECARGA), 39, 4)
     Call REPLA(X$, MKS$(PRECARGA), 51, 4)
     Call REPLA(X$, MKD$(TOT#), 55, 8)
     Call REPLA(X$, MKS$(PRECARGA), 65, 4)
     Call REPLA(X$, MKS$(PREUNI), 73, 4)
     Call REPLA(X$, MKD$(TOT#), 77, 8)
     Call GRAREG("!CARDETPE", X$, U&)
   Next U&
   '
   Call CARLISPE
   Call CIERRARCH("!CARDETPE")
   Call CALTOPED
   '
End Sub

Private Sub MONEMIS_DblClick()
   If CAMON% < 1 Then
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
         Next U&
     End If
   End If
   '
   MONEMIA% = MONEMI%
15 MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   If DMONE$(MONEMI%) = "" Then GoTo 15
   '
   MONEMIS = DMONE$(MONEMI%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   '
   If MONEMI% > 1 Then
     TICAMBIS.Visible = True
     TICAMBIS.Enabled = False
   End If
   '
   FIN& = ULTREG&("!CARDETPE")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = COEFCON * CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     Call REPLA(X$, MKS$(PRECARGA), 39, 4)
     Call REPLA(X$, MKS$(PRECARGA), 51, 4)
     Call REPLA(X$, MKD$(TOT#), 55, 8)
     Call REPLA(X$, MKS$(PRECARGA), 65, 4)
     Call REPLA(X$, MKS$(PREUNI), 73, 4)
     Call REPLA(X$, MKD$(TOT#), 77, 8)
     Call GRAREG("!CARDETPE", X$, U&)
   Next U&
   '
   Call CARLISPE
   Call CIERRARCH("!CARDETPE")
   Call CALTOPED
   '
End Sub

Private Sub MoPeCu_Click()
Call OPPED07.Command46_Click
End Sub

Private Sub MoPeCur_Click()
'\\\OT-5-869-28/12/05///
Call MoPeCu_Click
'\\\OT-5-869-FIN///
End Sub

Private Sub nueped_Click()
Call Command1_Click
End Sub

Private Sub otr_Click()
Call MENSERR(24, "No Disponible por Ahora")
End Sub

Private Sub PeCliEx_Click()
OPPED07.Option2 = True
Call OPPED07.Command62_Click
End Sub

Private Sub PeClPan_Click()
OPPED07.Option1 = True
Call OPPED07.Command62_Click
End Sub

Private Sub PeCoEx_Click()
OPPED07.Option2 = True
Call OPPED07.Command64_Click
End Sub

Private Sub PeCoPa_Click()
OPPED07.Option1 = True
Call OPPED07.Command64_Click
End Sub

Private Sub PORDESCU_Change()
   Call CALTOPED
End Sub

Private Sub PORDESCU_GotFocus()
   
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   If DERACCE%("MODESPE/NOMESS", "Modificacion descuento de pedido") < 2 Then
      On Error Resume Next
      Text13.SetFocus
      On Error GoTo 0
      If List1.ListCount > 0 Then AA1% = DERACCE%("MODESPE", "Modificacion descuento de pedido")
   End If
   '
End Sub

Private Sub RCtaCte_Click()
Call Command4_Click
End Sub

Private Sub RevAnu_Click()
Call OPPED07.Command43_Click
End Sub

Private Sub Sali_Click()
Call Command2_Click
End Sub

Private Sub Text1_Change()
   If XVALO(Text1.TEXT) < 999999 Then
      Call BLANFORMU
      CLIPROACTIVO = XVALO(Text1)
   End If
End Sub

Private Sub Text1_DblClick()
  Call SELECHO("DEUDOR1")
  NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
  Text1.TEXT = TRIM$(Str$(NCLIE))
  If NCLIE > 0 Then
    If REGICLI&(NCLIE) > 0 Then
      Call CARDACLI
      '\\\OT-08-0179-OD-21/04/08///
      FAVEN% = HAYFACVEN%(NCLIE)
      If FAVEN% > 0 Then
        Text8 = "Fac.Vencidas"
        Text8.ForeColor = RGB(255, 0, 0)
      End If
      '\\\OT-08-0179-OD-FIN///
      '\\\OT-05-0399-WAR-14/06/05///
      Image6.Enabled = True 'MODULO IMPORTADOR
      '\\\OT-05-0399-WAR-FIN///
      On Error Resume Next
      Command6.SetFocus
      On Error GoTo 0
    Else
      Image6.Enabled = False 'MODULO IMPORTADOR
    End If
  End If
End Sub

Private Sub Text1_GotFocus()
   If MARCONOTA.Visible = True Then
      On Error Resume Next
        Command6.SetFocus
      On Error GoTo 0
      Exit Sub
    ElseIf MARSELPRESU.Visible = True Then
      On Error Resume Next
        Command19.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
   MARCONOTA.Visible = False
   MARSELPRESU.Visible = False
   '
   If TRIM$(Text1) = "" Then
     Label11.ForeColor = RGB(0, 0, 255)
     Label11 = "Venta"
     Label11.ToolTipText = "Doble-Click para 'Reparación'"
   End If
   '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If VALICLIE% > 0 Then
      Call BLANFORMU
      If XVALO(Text1.TEXT) > 0 Then
        If XVALO(Text1.TEXT) <= 999999 Then
          NCLIE = XVALO(Text1)
          If REGICLI&(NCLIE) > 0 Then
            If MARCONOTA.Visible = False Then
              If MARSELPRESU.Visible = False Then
                Call CARDACLI
                Image6.Enabled = True 'MODULO IMPORTADOR
              End If
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
   '
End Sub
Private Sub Text10_LostFocus()
' comentado por que si no cuando se valida en el boton tilde
' aparece dos veces el mensaje

''   If VALICLIE% < 1 Then
''      On Error Resume Next
''      Text1.SetFocus
''      On Error GoTo 0
''      Exit Sub
''   End If
''   '
''   'VALIDACION DE ORDEN DE COMPRA PREEXISTENTE
''   '
''   If CONTROL$("ADMIPED", "NOREPOC") = "SI" And TRIM$(Text10) <> "" Then
''         NC% = XVALO(Text1)
''         NORCOM$ = UCase$(TRIM$(Text10))
''         If OCOMEXISTE%(NC%, NORCOM$) Then
''            Call COMUNI("Orden de Compra Existente")
''            Text10 = ""
''         End If
''   End If
   
End Sub
Private Sub Text11_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   If ULTREG&("PEDENCA") > 2 Then
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
   NPP& = NUPEDIDO&
   If XVALO(NUMEPRE$) < NPP& Then
     Call MENSERR(24, "Número no Válido")
     Text11.TEXT = TRIM$(Str$(NPP&))
     Exit Sub
   End If
   '
   If NUMEPRE$ <> TRIM$(Text11.TEXT) Then
     Text11.TEXT = NUMEPRE$
     Call MENSERR(24, "Número de Presupuesto Ajustado")
   End If
   '
End Sub

Private Sub Text12_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text13_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text16_Change()
   '
   'VALIDO QUE ESTE VISIBLE EL PICTURE EN ESE CASO EXISTE EL TAMOTPRE
   If Picture6.Visible = False Then Exit Sub
   '
   'TEXT DE INGESO DE MOTIVOS
   If Not IsNumeric(Text16) Then Text16 = ""
   
   NUMOT% = XVALO(Text16)
   If NUMOT% >= 0 Then
      Label28 = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
   Else
      Label28 = ""
   End If
   

End Sub

Private Sub Text17_GotFocus()
Text1.SetFocus
End Sub


Private Sub Text19_Change()
   '\\\OT-07-0040-RG-02/02/07///
   Label23.Caption = ""
   Label24.Caption = ""
   LENTREGA(2) = ""
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
         Label23.Caption = DOMITRANS$(NTRAI%)
         Label24.Caption = LOCATRANS$(NTRAI%)
         '\\\OT-07-0040-RG-FIN///
      End If     '
   End If
   '\\\OT-6-422-FG-FIN
End Sub

Private Sub Text19_DblClick()
   '
   Call Command24_Click
   '
End Sub

Private Sub Text2_gotfocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text20_Change()

   'AGREGADO MH04/10
    NSUC& = XVALO(Text20)
      Label27 = ""
      Text19 = ""
      If NSUC& > 0 Then
        'If NSUC& <= ULTREG&("!SUCENTRE") Then
        'Lectura desde la Tabla SQL.
        CONDI$ = "Nume_Cli = " & Text1 & " and Nume_Suc = " & NSUC
        If CantRegistros("SUCENTRE", CONDI$) > O Then
            Label27 = VALOBADA("SUCENTRE", "Deno_Suc", CONDI$, "NOMES")
            LENTREGA(0) = VALOBADA("SUCENTRE", "Domi_Suc", CONDI$, "NOMES")
            LENTREGA(1) = VALOBADA("SUCENTRE", "Loca_Suc", CONDI$, "NOMES")
            Text19 = XVALO(VALOBADA("SUCENTRE", "Codi_Tra", CONDI$, "NOMES"))
            ' Fin de Lectura
        ElseIf NSUC& <= ULTREG&("!SUCENTRE") Then
            XX$ = REGLEIDO$("!SUCENTRE", NSUC&)
            Label27 = TRIM$(Mid$(XX$, 3, 62))
            LENTREGA(0) = TRIM$(Mid$(XX$, 65, 48))
            LENTREGA(1) = TRIM$(Mid$(XX$, 113, 48))
            Text19 = CVI(Mid$(XX$, 195, 2))
        End If
      End If
      
      If XVALO(Text19) < 1 Then
        Text19 = NUTRANCLI%(XVALO(Text1), XVALO(Text20))
      End If
      '
End Sub

Private Sub Text20_DblClick()
    '
    Call Command25_Click
    '
End Sub

Private Sub TEXT21_DblClick(Index As Integer)
    '
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
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
    '
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

Private Sub Text44_LostFocus()
   If Left$(UCase$(Label21(4)), 5) = "LETRA" Then
      Text44 = UCase$(Left$(Text44, 1))
      Exit Sub
   End If
   '
   If XVALO(Text44) > 99.9 Then Text42 = "99.9"
   If TRIM$(Text44) <> "" Then Text44 = Format$(XVALO(Text44), "#0.0")
End Sub

Private Sub Text5_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text21_Change(Index As Integer)
    '
    If XVALO(TEXT21(Index).TEXT) < 0 Or XVALO(TEXT21(Index).TEXT) > 255 Then
       TEXT21(Index).TEXT = ""
    End If
    '
    If Index = 0 Then
        If XVALO(TEXT21(Index).TEXT) < 1 Then
            Eco21(Index).TEXT = "CONCERTADOS FIJOS"
            ARCHIPRE$ = "MASTER"
            LIPRACTI% = 0
          Else
            Eco21(Index).TEXT = ECOARCH$("LISTAS", MKI$(XVALO(TEXT21(Index).TEXT)))
            ARCHIPRE$ = TRIM$(TEXT21(0).TEXT): LIPRACTI% = XVALO(TRIM$(ARCHIPRE$))
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        End If
        MNLX% = MONELISTA%(XVALO(TEXT21(Index)))
        '
        'CONFIGURACION DE MONEDA POR DEFECTO
        MOX% = XVALO(CONTROL("ADMIPED", "MONEDEFE"))
        If MOX% > 0 Then
          MNLX% = MOX%
        End If
        '
        CAINTEN% = 0
        While MONEMI% <> MNLX%
           CAINTEN% = CAINTEN% + 1
           If CAINTEN% > 128 Then
              MONEMI% = 1
              Call MONEMIS_Change
              GoTo 19
           End If
           Call MONEMIS_DblClick
        Wend
        '
19      If CONTROL("PEDIDO", "MODODESC") = "MONTO" And XVALO(TEXT21(0)) = 1 Then
           Call DESCUMONTO(Me)
        Else
           PORDESCU.TEXT = DESCUCLI$(XVALO(Text1))
        End If
        '
        'If InStr(1, EMPREAC$, "FERVI") > 0 Then
        '    NLIST% = XVALO(TEXT21(Index))
        '    If NLIST% > 0 Then Call GENERAR_DAT(NLIST%)
        'End If
        '
      ElseIf Index = 1 Then
      
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(XVALO(TEXT21(Index).TEXT)))
        DESXREX = DESCURECA(XVALO(TEXT21(1)))
        '
        If ROUND#(DESXREX) < 100 Then
          Label25 = FORMATNUM(DESXREX, "F5.1")
        Else
          Call COMUNI("Descuento No Válido !!! Verifique")
          Label25 = ""
        End If
        '
      ElseIf Index = 2 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(XVALO(TEXT21(Index).TEXT)))
        If CONTROL("PEDIDO", "MODODESC") = "MONTO" And XVALO(TEXT21(0)) = 1 Then
           Call DESCUMONTO(Me)
        End If
        Call Check2_Click ' llama al evento donde se valida el vendedor para definir descuento

      ElseIf Index = 3 Then
        If EXISTE%(LUDAT$ + "TATINGRE.RFS") > 0 Then
           Eco21(Index).TEXT = ECOARCH$("TATINGRE", Chr$(XVALO(TEXT21(Index).TEXT)))
        End If
        
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
      ElseIf Index = 3 Then
        If EXISTE%(LUDAT$ + "TATINGRE.RFS") < 1 Then
          Call COMUNI("Tabla de Motivos de Ingreso de Pedidos \No Disponible en la Presente Versión")
          Exit Sub
        End If
        
        ARCHEC$ = "TATINGRE/Tabla de Ingresos": LOCA% = 1
    End If
    '
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
    '
End Sub
   '
Function VALICLIE%()
   VCC% = 0
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 999999 Then
       NCLIEN = XVALO(Text1.TEXT)
       If RASOCLI$(NCLIEN) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function

Sub BLANFORMU()
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   Text7.TEXT = ""
   Text8.TEXT = ""
   Text10.TEXT = ""
   Text12.TEXT = ""
   Text13.TEXT = ""
   Text15 = ""
   Text16 = ""
   Text6.TEXT = ""
   Text17.TEXT = HOS$
   Label18 = ""
   Text19 = ""
   Text20 = ""
   Text8.ForeColor = &H808000
   List1.Clear
   Call SETULTREG("!CARDETPE", 0)
   'AGREGADOMH04/10
   Call SETULTREG("!SUCENTRE", 0)
   Call FRESHECHO("!SUCENTRE")
   '
   TEXT21(0).TEXT = "": Eco21(0) = ""
   TEXT21(1).TEXT = "": Eco21(1) = ""
   TEXT21(2).TEXT = "": Eco21(2) = ""
   TEXT21(3).TEXT = "": Eco21(3) = ""  'text de tipo de ingreso
   Text44 = ""
   '
   PORDESCU.TEXT = ""
   For KK% = 0 To 3
      Total(KK%).TEXT = ""
      TOTIVA(KK%).TEXT = ""
      LENTREGA(KK%).TEXT = ""
   Next KK%
   Label13 = ""
   Label14 = ""
   Label20 = ""
   Label26 = "" ' SE AGREGA PARA LIMPIAR EL LABEL DEL CODIGO DE CUENTA
   Label31 = "" 'descuento de retira cliente
   Check2.Value = 0
   Text141 = ""
   Frame4.Height = 1440
   '
   MARCONOTA.Visible = False
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Text11.TEXT = TRIM$(Str$(NUPEDIDO&))
   '\\\OT-05-0399-WAR-14/06/05///
   Image6.Enabled = False 'MODULO IMPORTADOR
   '\\\OT-05-0399-WAR-FIN///
   GRID2.Visible = False
   Option1.Value = False
   Option2.Value = False
End Sub
'
Sub CARDACLI()
   '
   Screen.MousePointer = 11
   MARCONOTA.Visible = False
   MARSELPRESU.Visible = False
   If CONTROL$("ADMIPED", "VERLICRE") = "SI" Then VERLICRE% = 1
     '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   NCLIEN = XVALO(Text1.TEXT)
   If NCLIEN > 0 Then
     '
     If InStr(EPAC$, "DOSIVA") > 0 Then
       PVCL% = PIVACLI%(NCLIEN)
       If PVCL% <> 0 And PVCL% <> 5 Then
         PRVC$ = TRIM$(PROVCLI$(NCLIEN))
         If PRVC$ = "" Or DEPROVI$(PRVC$) = "" Then
           Call COMUNI("Imposible Registrar O/Venta.\Falta Código de Provincia del Cliente.")
           Text1 = ""
           GoTo 99
         End If
       End If
     End If
     '
     If RASOCLI$(NCLIEN) <> "" Then
       Text2.TEXT = RASOCLI$(NCLIEN)
       Text3.TEXT = DOMICLI$(NCLIEN)
       Text4.TEXT = CPOSCLI$(NCLIEN)
       Text5.TEXT = LOCACLI$(NCLIEN)
       Label26 = CODICLI$(NCLIEN) 'Carga el numero de cliente/prospect
       '
       '27/03/07 (OT-07-0133)
       LMCR# = LIMICRE(NCLIEN)  'PONE EN 'LIMITE DE CREDITO' LA LEYENDA 'SUSPENDIDO'
        If LMCR# < 0 Then
           Text7 = "Suspendido"
        Else
           Text7.TEXT = CSTRING$(MKS$(LIMICRE(NCLIEN)), 3, 13, 2, 2)
        End If
       '27/03/07 (FIN)
       '
       Text12.TEXT = TRIM$(POSIVAC$(NCLIEN)) + " - CUIT " + CUITCLI$(NCLIEN)
       CREDISPO# = LIMICRE(NCLIEN) - SALDCLI#(NCLIEN) - CHEPEACRE#(NCLIEN) - VALOPEDPEN1#(NCLIEN)
       Text8.ForeColor = RGB(0, 0, 0)
       If CREDISPO# < 0 Then
          Text8.ForeColor = RGB(255, 0, 0)
          Call COMUNI("Atención : Cuenta sin Crédito Disponible ")
       End If
       Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 13, 2, 2)
       '
       TEXT21(0).TEXT = Str$(LIPRCLI%(NCLIEN))
       TEXT21(1).TEXT = Str$(CPAGCLI%(NCLIEN))
       If CONTROL$("ADMIPED", "COPAVARI") = "SI" Then TEXT21(1).TEXT = ""
       TEXT21(2).TEXT = Str$(VENDCLI%(NCLIEN))
       '
       If XVALO(TEXT21(2).TEXT) < 1 Then
          Eco21(2).TEXT = "GERENCIA"
       End If
       '
       PORDESCU.TEXT = DESCUCLI$(NCLIEN)
       '
       'AGREGUE MH02-10
       '\\\OT-07-0010-RG-12/01/07///
        ' CARGA LISTA DE SUCURSALES DEL CLIENTE ACTIVO
        CONDI$ = "Nume_Cli = " & NCLIEN
        If CantRegistros("SUCENTRE", CONDI$) > 0 Then
            Call CARGALISEL("!SUCENTRE", "SUCENTRE", "Nume_Suc/I6;Deno_Suc/A62", CONDI$, "NOMESS")
'        Else
'            Call FILTERFILE("SUCENTRE", "", 5, MKI$(NCLIEN))
        End If
       '\\\OT-07-0010-RG-FIN///
       '
       'CONTROL PARA PRESENTAR DATOS DE ENTREGA
       If CONTROL("NOPREENT", "CONFPEDI") <> "SI" Then
        For KKI% = 0 To 2
          LENTREGA(KKI%) = LUENTRECLI$(NCLIEN, KKI%)
          Text19.TEXT = NUTRANCLI%(XVALO(Text1))
        Next KKI%
       End If
       '
       If REGICLI(NCLIEN) > 0 Then
          Text6 = ANOTACLI(NCLIEN)
          MARCONOTA.Visible = True
       End If
       '
       '23/03/07 (OT-07-0121)'NO PERMITE GENERAR UN PEDIDO A UN CLIENTE CON LA
       LMCR# = LIMICRE(NCLIEN)  'CUENTA SUSPENDIDA
       If LMCR# < 0 Then
          Call COMUNI("Imposible Generar Pedido.\ \La Cuenta esta Suspendida")
          Text1 = ""
          Text1.SetFocus
          GoTo 99
       End If
       '23/03/07 (FIN)
       '
       If VERLICRE% = 1 Then
         SDEU# = SALDCLI#(NCLIEN)
         CHPA# = CHEPEACRE#(NCLIEN)
         If SDEU# + CHPA# > LMCR# Then
           Call COMUNI("Este Pedido no Podrá Facturarse.\  \Límite de Crédito Asignado ya está Superado.\  \Informe al Supervisor de Ventas.")
         End If
       End If
       '
    End If
   End If
99 Screen.MousePointer = 1
   '
End Sub
   '
Function NUPEDIDO&()
   '
   NPP& = 1
   Call ABREPEDCLI
'
'   strsql = "SELECT max(NroPed) as maximo FROM Pedenca "
'   Dim MIRS As ADODB.Recordset
'   Set MIRS = New ADODB.Recordset
'   Set MIRS = FLEXCONN.Execute(FILTSQL$(strsql))
'   On Error Resume Next
'   If Not IsNull(MIRS!maximo) Then
'      NPP& = 1 + MIRS!maximo
'   End If
'   NUPEDIDO& = NPP&
'   MIRS.Close
'   Set MIRS = Nothing
   '
   NPXX& = 1 + XVALO(VALOBADA("PEDENCA", "MAX(NROPED)", "NROPED > 0", "NOMESS"))
   If NPXX& > NPP& Then NPP& = NPXX&
   NUPEDIDO& = NPP&
   '
End Function
'
Private Sub Text6_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text7_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text8_DblClick()
    HOII% = HOY(HOS$)
    DIAHOY = FECHATEX$(HOII%)
    If Text8 <> "Fac.Vencidas" Then GoTo 99
    NCLIE = XVALO(Text1)
    If NCLIE > 0 Then
      Call APERBASDAT("CABAN")
      SQLX$ = "SELECT * FROM SADEUPEN "
      SQLX$ = SQLX$ + " Where NuClien = " & NCLIE
      SQLX$ = SQLX$ + " AND SaldDebe > SaldAcre "
      SQLX$ = SQLX$ + " AND FEVENCIM < '" & DIAHOY & "' "
      JJ& = 0
      Dim DEUDATMP As ADODB.Recordset
      Set DEUDATMP = New ADODB.Recordset
25    On Error Resume Next
      DEUDATMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
      End If
      On Error GoTo 0
      '
      With DEUDATMP
        On Error Resume Next
        .MoveFirst
         If Err > 0 Then
           Call COMUNI(" El Cliente no Registra Facturas Vencidas ")
           GoTo 99
          Else
           On Error GoTo 0
           Do While Not .EOF
              SALDEB# = XVALO(!SaldDebe)
              SALACR# = XVALO(!SALDAcre)
              DEUDA$ = SALDEB# - SALACR#
              IMPORIG$ = SAFETEXT(!IBruComp)
              FEVE = CVINT(!FEVENCIM)
              NUMFACT$ = SAFETEXT(!CODICOM_LAR)
              FECHOR = CVINT(!FechEmisi)
              JJ& = JJ& + 1
              Z$ = Space$(128)
              Call REPLA(Z$, NUMFACT$, 1, 18)
              Call REPLA(Z$, FECHATEX$(FECHOR), 23, 8)
              Call REPLA(Z$, FORMATNUM(IMPORIG$, "F12.4"), 32, 13)
              Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
              Call REPLA(Z$, DEUDA$, 57, 13)
              Call GRAREG("SELDOCAP", Z$, JJ&)
              '
              .MoveNext
           Loop
         End If
         On Error GoTo 0
      End With
      DEUDATMP.Close
      Set DEUDATMP = Nothing
      Call SETULTREG("SELDOCAP", JJ&)
      SPDX$ = REGSEL$("SELDOCAP")
    End If
99
    

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
     VH(MONEMI%) = XVALO(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(XVALO(TICAMBIS), "F9.4")
End Sub

Private Sub Timer1_Timer()
   '
   Call CALTOPED
   '
End Sub

Private Sub Timer2_Timer()
  '
  Static FIDAPRIN#
  '
  FDP# = FIDATIM#("IPEDPRIN")
  If FDP# <> FIDAPRIN# Then
    FIDAPRIN# = FDP#
    RFF$ = RECFILT$("IPEDPRIN", 3, MKI$(0))
    '
    Screen.MousePointer = 11
    For UL& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UL&, 4))
      YY$ = REGLEIDO$("IPEDPRIN", U&)
      TDOC$ = UCase$(Left$(YY$, 2))
      NPX& = XVALO(Mid$(YY$, 3, 6))
      '
      If NPX& > 0 Then
        TIDOCUM$ = "Pedido de Cliente"
        If TDOC$ = "OD" Then TIDOCUM$ = "Orden de Despacho"
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        '
        ' esto para forzar que haya bajado al disco
        Call BLOQARCH("PDOCSPL")
        Call LIBARCH("PDOCSPL")
        '
        For UI& = ULTREG&("PDOCLST") To 1 Step -1
          XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
          If InStr(XX$, TIDOCUM$) > 4 Then
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
        GoTo 49
        '
5       REDOCU$ = REGACT$("PDOCLST")
        URE& = ULTREG&("PDOCSPL")
        LI& = 0: LS& = URE&
6       E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 49
        L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 49
        DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
        If DCC& < DOCUNU& Then LI& = L&: GoTo 6
        If DCC& > DOCUNU& Then LS& = L&: GoTo 6
        '
8       TPSP$ = ""
        Screen.MousePointer = 11
        While L& > 1
          DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
          If DCC& < DOCUNU& Then GoTo 9
          L& = L& - 1
        Wend
        '
9       For KKL& = L& To URE&
          TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
          If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 10
          If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
            TPSP$ = TPSP$ + Mid$(TBUF$, 5)
          End If
        Next KKL&
        '
10      PPXX% = InStr(TPSP$, Chr$(255))
        If PPXX% > 0 Then
          COPRIAT$ = Left$(TPSP$, PPXX% - 1)
          TPSP$ = Mid$(TPSP$, PPXX% + 1)
          '
          If Len(COPRIAT$) > 0 Then
            COPRI% = Asc(Left$(COPRIAT$, 1))
            ATRIP$ = Mid$(COPRIAT$, 2)
            Call PRINTCOMMAND(COPRI%, ATRIP$)
          End If
          GoTo 10
        End If
        '
      End If
      '
      YY$ = REGLEIDO$("IPEDPRIN", U&)
      HOII% = HOY(HOS$)
      Call REPLA(YY$, MKI$(HOII%), 57, 2)
      Call REPLA(YY$, Time$, 59, 6)
      Call GRAREG("IPEDPRIN", YY$, U&)
      Call CIERRARCH("IPEDPRIN")
      '
49  Next UL&
    Screen.MousePointer = 1
    '
  End If
  '
End Sub

Private Sub Total_Change(Index As Integer)
   If Index = 0 Then
     Call CALTOPED
'      DESC$ = PORDESCU.TEXT
'      TTT1 = XVALO(Total(0).TEXT) * COEFDESCU(DESC$) / 100
'      Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
'      TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
'      Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
   End If
'   If Index < 3 Then
'      TTT1 = XVALO(Total(2).TEXT)
'      TTT2 = XVALO(TOTIVA(3).TEXT)
'      Total(3).TEXT = Str$(TTT1 + TTT2)
'   End If
End Sub
'
Sub PRIPEDIDO(OPCI%)
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    
    NCLIEN = XVALO(Text1)
    FORIPRE$ = TRIM$(CONTROL$("", "FORPECLI"))
    If FORIPRE$ = "" Then Exit Sub
    '
    PUERDESPA$ = ""
    PUERORVE = ""
    If OPCI% = 1 Then
       PUEROVTA$ = CONTROL("*", "PUERORVE")
       If CONTROL(IDENTER$, "PUERORVE") <> "" Then PUEROVTA$ = CONTROL(IDENTER$, "PUERORVE")
    ElseIf OPCI% = 2 Then ' ORDEN DE DESPACHO
       PUERDESPA$ = CONTROL("*", "PUERDESP")
       If CONTROL(IDENTER$, "PUERDESP") <> "" Then PUERDESPA$ = CONTROL(IDENTER$, "PUERDESP")
       FORODES$ = TRIM$(CONTROL$("", "FORDESPA"))
       If FORODES$ <> "" Then
          FORIPRE$ = FORODES$
       End If
    End If
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Pedido de Cliente.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
'    CANCELPRINT = 0    Se comenta para respetar el Check2 del formulario CONPRIPE opcion de impresión
    If CONTROL$("PEDCLIEN", "SUPRINT") = "SI" Then
       CANCELPRINT% = 1
    End If
    '
'    EPAC$ = TRIM$(UCase$(EMPREAC$))
'    If InStr(EPAC$, "DOSIVAC") > 0 Then
'      CANCELPRINT% = 0
'      Screen.MousePointer = 1
'      OPRIPEDI.Show 1
'      If OPRIPEDI.Check1(0).Value <> 1 Then
'        CANCELPRINT% = 1
'      End If
'    End If
'    '
    Screen.MousePointer = 11
    NCLIEN = XVALO(Text1)
    FECHDOC$ = Text9.TEXT
    NUMEDOC$ = TRIM$(Text11.TEXT)
    While Len(NUMEDOC$) < 6
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    '
    TIPODOC$ = "Pedido de Cliente"
    If NCLIEN < 1 Then
      TIPODOC$ = "Pedido Interno"
    End If
    If OPCI% = 2 Then
       TIPODOC$ = "Orden de Despacho"
    End If
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "PRE-LIST"
    CODTABU1$(5) = "PORDESCU"
    CODTABU1$(6) = "PRE-NETO"
    CODTABU1$(7) = "NET-ITEM"
    CODTABU1$(8) = "FECH-REC"
    CODTABU1$(9) = "ORD-ITEM"
    CODTABU1$(10) = "Z-PLANO"
    CODTABU1$(11) = "DES-MPRI"
    CODTABU1$(12) = "UBI-MPRI"
    CODTABU1$(13) = "CAN-PLCO"
    CODTABU1$(14) = "UNIMED" 'SE AGREGA POR QUE ESTABA FALTANDO
    CODTABU1$(15) = "LAR-MATE"
    CODTABU1$(16) = "NRO-LOTE"
    CACOLTAB1% = 16
    '
    'DESTIDOC% = NCLIEN
    DESTIDOCNUE = NCLIEN
    CAITAB1% = 0
    CAITCRU% = 0
    '
    PRIVA$ = CONTROL$("PEDCLIEN", "PRECOIVA")
    DESCPRI% = 0: DESCOMP% = 0: Especif% = 0
    If CONTROL("ADMIPED", "IMPDESPR") = "SI" Then DESCPRI% = 1
    If CONTROL("ADMIPED", "IMPDESCO") = "SI" Then DESCOMP% = 1
    If CONTROL("ADMIPED", "IMPRESPE") = "SI" Then Especif% = 1
    If DESCOMP% + Especif% < 1 Then DESCPRI% = 1 'POR SI NO SE SELECCIONO NADA
    '
    For U& = 1 To ULTREG&("!CARDETPE")
      XX$ = REGLEIDO$("!CARDETPE", U&)
      CIPRE% = CVI(Left$(XX$, 2))
      CAPEX$ = Str$(CVS(Mid$(XX$, 47, 4)))
      '
      If CIPRE% > 0 Then
        If CIPRE% <= NUMAS% Then
          If XVALO(CAPEX$) > 0 Then
            '
            '\\\OT-07-0107-RG-15-03-07///
            PRELIX = PRELISTA#(ARCHIPRE$, CIPRE%)
            PRELI$ = Str$(PRELIX)
            '\\\OT-07-0107-RG-FIN///
            'PRELI$ = Str$(CVS(Mid$(XX$, 39, 4)))
            PREUX$ = FORMATNUM$(CVS(Mid$(XX$, 51, 4)), "F15.4")
            PRESINRED = CVS(Mid$(XX$, 73, 4))
            LOTEX$ = Mid$(XX$, 3, 16)
            DESCX$ = "0": On Error Resume Next
                          DESCX$ = Str$(100 * (XVALO(PRELI$) - PRESINRED) / XVALO(PRELI$))
                          On Error GoTo 0
                          
            PIVA# = 1
            If PIVACLI%(NCLIEN) <> 5 And PIVACLI%(NCLIEN) <> 7 Then
            'SI LA POSICION DEL IVA NO ES DE EXPORTACION VERIIFICA SITUACION
              If PRIVA$ = "SI" Then PIVA# = 1 + PORCEIVA(CIPRE%) / 100 ' VERIFICA QUE TIENE SETEADO EN EL SETUP PARA VER SI IMPRIME CON IVA INCLUIDO
            End If
            ITOTX$ = FORMATNUM$((XVALO(CAPEX$)) * (XVALO(PREUX$)) * PIVA#, "F15.4")
            '
            FEX = CVI(Mid$(XX$, 63, 2))
            FENTX$ = FECHATEX$(FEX)
            'LECTURA DE NRO DE PRESUPUESTO Y NORDEN PARA AHP PARA TOMAR EL NUMERO DE PROVEEDOR DE LA TABLA DE PRESUPUESTO
            NUORIT% = CVI(Mid$(XX$, 123, 2))
            NUPRESU& = CVS(Mid$(XX$, 125, 4))
            'FIN AHP
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(CIPRE%)
            TETABU1$(3, CAITAB1%) = CAPEX$
            TETABU1$(4, CAITAB1%) = PRELI$
            TETABU1$(5, CAITAB1%) = DESCX$
            PRENETO# = XVALO(PREUX$) * PIVA# ' PRECIO CON IVA
            PREUX$ = FORMATNUM$(PRENETO#, "F15.4")
            TETABU1$(6, CAITAB1%) = PREUX$
            TETABU1$(7, CAITAB1%) = ITOTX$
            TETABU1$(8, CAITAB1%) = FENTX$
            TETABU1$(9, CAITAB1%) = Str$(CAITAB1%)
            TETABU1$(10, CAITAB1%) = NUPLANO$(CIPRE%)
            TETABU1$(14, CAITAB1%) = UNIMED$(CIPRE%) 'SE AGREGA POR QUE ESTABA FALTANDO
            '
            SUBTOCONIVA = SUBTOCONIVA + XVALO(ITOTX$)
            ' LO QUE SIGUE ES PARA IMPRIMIR O/DESPACHO
            If OPCI% = 2 Then
               TETABU1$(4, CAITAB1%) = ""
               TETABU1$(5, CAITAB1%) = ""
               TETABU1$(6, CAITAB1%) = ""
               TETABU1$(7, CAITAB1%) = ""
               'CANTIDAD DE BULTOS
               CANTI# = XVALO(CAPEX$)
               CABU = CANBULTOS(CIPRE%, CANTI#, "VENTA")
               TETABU1$(13, CAITAB1%) = TRIM$(FORMATNUM$(CABU, "F7"))
               'AHP TRAE EL NUMERO DE PROVEEDOR DE LA TABLA DE PRESUPUESTO
               
               If InStr(EPAC$, "AHP") > 0 Then
                   'UBICACION DE AHP POR CODIGO Y PROVEEDOR
                   'xxxACA TOMAR SEGUN NROPRES Y ORDEN PARA GUARDAR EN LA TABLA DE LOS COMPONENTES EL NRO DE PEDIDO.
                   CONDI$ = " NroPres = " & NUPRESU& & " and NuOrdIt= " & NUORIT%
                   NPROV& = XVALO(VALOBADA("PRESUVEN", "Num_Prov", CONDI$, "NOMESS"))
                   MARCAX$ = CONDICOM$(-1 * NPROV&)
                   UBIX$ = TRIM$(UBISTOSQL$(CIPRE%, NPROV&))
                   If UBIX$ = "" Then UBIX$ = UBISTO$(CIPRE%)
               Else
                   UBIX$ = UBISTO$(CIPRE%)
               End If
               TETABU1$(12, CAITAB1%) = UBIX$
               TETABU1$(15, CAITAB1%) = TRIM$(MARCAX$)
               TETABU1$(16, CAITAB1%) = TRIM$(LOTEX$)
               'ACUMULO EL PESO
               PESOUNITARIOX = PESUNI(CIPRE%)
               ACUMUPESO = ACUMUPESO + (PESOUNITARIOX * CANTI#)
               'FIN AHP
            End If
            ' FIN O/DESPACHO
            '
            PHCODI% = XVALO(ATRIFORM$(FORIPRE$, "COD-MAT/POSINHOR"))
            PHDESI% = XVALO(ATRIFORM$(FORIPRE$, "DES-MAT/POSINHOR"))
            PVCODI% = XVALO(ATRIFORM$(FORIPRE$, "COD-MAT/POSINVER"))
            PVDESI% = XVALO(ATRIFORM$(FORIPRE$, "DES-MAT/POSINVER"))
            If PHDESI% <= PHCODI% + 3 Then
               If PVDESI% <= PVCODI% + 1 Then
                  CAITAB1% = CAITAB1% + 1
               End If
            End If
            '
            DESPRIN$ = DESCRIT0$(CIPRE%)
            DESCOMPLEM$ = DESCRIT2$(CIPRE%)
            ESPECIFICAC$ = ESPECIFIT$(CIPRE%)
            '
            'DESCRIPCION
            DIMPRE$ = ""
            If DESCPRI% > 0 Then
               If DESPRIN$ <> "" Then
                  DIMPRE$ = DESPRIN$
               End If
            End If
            '
            'DESCRIPCION COMPLEMENTARIA
            If DESCOMP% > 0 Then
               If DESCOMPLEM$ <> "" Then
                 If DIMPRE$ <> "" Then
                   DIMPRE$ = DIMPRE$ + vbCrLf
                 End If
                 DIMPRE$ = DIMPRE$ + DESCOMPLEM$
               End If
            End If
            '
            'ESPECIFICACION
            If Especif% > 0 Then
               If ESPECIFICAC$ <> "" Then
                 If DIMPRE$ <> "" Then
                   DIMPRE$ = DIMPRE$ + vbCrLf
                 End If
                 DIMPRE$ = DIMPRE$ + ESPECIFICAC$
               End If
            End If
            '
            'IMPRESION DE LA DESCRIPCION CON TODAS SUS VARIANTES CONFIGURADAS
            FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
            ATELIB% = XVALO(Mid$(FORTELIB$, 2))
            Call FRATEXTO(DIMPRE$, ATELIB%)
            For KKU1% = 1 To CARETEX%
                  TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
                  CAITAB1% = CAITAB1% + 1
            Next KKU1%
            'SE RESTA CAITAB1% PORQUE DEL FOR SALE EN LA PROXIMA LINEA
            'LO CUAL YA SE HACE CUANDO PASA AL PROXIMO CODIGO (VER MAS ARRIBA)
            CAITAB1% = CAITAB1% - 1
            'LINEA DE SEPARACION
            Call CARLINSEP
            '
            ' IMPRIME LOS COMPONENTES
            If OPCI% = 2 Then
               If DESCOMPO%(CIPRE%) > 0 Then
                   CANTI = CVS(Mid$(XX$, 47, 4))
                   Call PRINCOMPO(CIPRE%, CANTI)
               End If
            End If
            '
            If CVS(Mid$(XX$, 89, 4)) > 0 Then
               Call IMPRIMEBONIFICACION(U&)
            End If

            
          End If
        End If
      End If
109 Next U&
    '
    SUMATOTA# = XVALO(Total(0).TEXT)
    IDESCUEN# = XVALO(Total(1).TEXT)
    '\\\OT-05-0853-WAR-24/11/05///
      If PIVACLI%(NCLIEN) = 5 Or PIVACLI%(NCLIEN) = 7 Then
        Label13 = "" 'SI CLIENTE EXTERIOR O PROMOVIDO
        Label14 = "" 'NO DISCRIMINA IVA
        Label20 = ""
      End If
    '\\\OT-05-0853-WAR-FIN///
    CODPARAM$(1) = "IMP-NETO": DOCPARAM$(1) = Str$(SUMATOTA#)
    CODPARAM$(2) = "VENDEDOR":     DOCPARAM$(2) = Eco21(2).TEXT
    CODPARAM$(3) = "COND-PAG":     DOCPARAM$(3) = Eco21(1).TEXT
    CODPARAM$(4) = "NRO-OCOM":     DOCPARAM$(4) = Text10.TEXT
    CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = LENTREGA(0).TEXT
    CODPARAM$(6) = "DESTINO": DOCPARAM$(6) = LENTREGA(1).TEXT
    CODPARAM$(7) = "TRANSPOR": DOCPARAM$(7) = LENTREGA(2).TEXT
    CODPARAM$(8) = "IMP-DESC": DOCPARAM$(8) = Total(1).TEXT
    CODPARAM$(9) = "IMPNEDES": DOCPARAM$(9) = Total(2).TEXT
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = Label13 ' TOTIVA(3).TEXT
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = Total(3).TEXT
    '
    SIGNO$ = ""
    If XVALO(PORDESCU.TEXT) > 0 Then SIGNO$ = "+"
    PORDXXTO$ = TRIM$(Label25) & SIGNO$ & TRIM$(PORDESCU.TEXT)
    '
    CODPARAM$(12) = "POR-DESC": DOCPARAM$(12) = PORDXXTO$ 'PORDESCU.TEXT
    '\\\OT-5-869-28/12/05///
    'CODPARAM$(13) = "MONEDEMI": DOCPARAM$(13) = MONEMIS
    '\\\OT-5-869-FIN///
    CODPARAM$(13) = "PERC-IVA": DOCPARAM$(13) = Text15
    CODPARAM$(14) = "IMP-IVA1": DOCPARAM$(14) = Label14
    'AGREGADOMH04-10
    CODPARAM$(15) = "NRO-PASE": DOCPARAM$(15) = Label27
    '\\\OT-07-0007-RG-08/01/07///
    HORA$ = HORAENT$(NCLIEN)
    CODPARAM$(16) = "Z-PLANO": DOCPARAM$(16) = TRIM$(HORA$)
    '\\\OT-07-0007-RG-FIN///
    '
    '\\\OT-07-0010-RG-12/01/07///
    AISUC% = XVALO(Text20)
    TELSUC$ = TELESUC$(NCLIEN, AISUC%)
    CODPARAM$(17) = "REF-MAT1": DOCPARAM$(17) = TRIM$(TELSUC$)
    '\\\OT-07-0010-RG-FIN///
    '
    CODPARAM$(18) = "IMPO-CHE": DOCPARAM$(18) = TRIM$(FORMATNUM$(SUBTOCONIVA, "F12,2"))
    DESCONIVA = SUBTOCONIVA * (COEFDESCU(PORDXXTO$) / 100)
    CODPARAM$(19) = "IMPO-PEN": DOCPARAM$(19) = TRIM$(FORMATNUM$(DESCONIVA, "F12,2"))
    '
    'imprime domicilio de transportista (expreso)
    NTRA% = XVALO(Text19)
    If Text19.TEXT = "" Then
       If TRIM(NTRANCLI$(NCLIEN)) = TRIM(LENTREGA(2)) Then NTRA% = NUTRANCLI%(NCLIEN)
    End If
    '
    CODPARAM$(20) = "REF-MAT2": DOCPARAM$(20) = DOMITRANS$(NTRA%)
    '
    REG$ = FILTERGETRECORD("CONPAG", 1, Chr$(XVALO(TEXT21(1))))
    
    PORCONPAG = CVS(Mid$(REG$, 44, 4))
    CODPARAM$(21) = "CANT-DIS": DOCPARAM$(21) = TRIM$(FORMATNUM$(PORCONPAG, "F5.1"))
    CODPARAM$(22) = "DES-OPER": DOCPARAM$(22) = TRIM$(Label31) 'DESCUENTO RETIRA CLIENTE GABAL
    CODPARAM$(23) = "REF-PLCO": DOCPARAM$(23) = TRIM$(Text141) 'IMPORTE DE DESCTO RETIRA CLIENTE GABAL
    CODPARAM$(24) = "LAR-CUTE": DOCPARAM$(24) = TRIM$(TELETRANS$(NTRA%))  'Telefono del transportista
    CODPARAM$(25) = "PES-CONJ": DOCPARAM$(25) = TRIM$(FORMATNUM$(ACUMUPESO, "F15.4")) 'PESO ACUMULADO
    NTRAI% = NUTRANCLI%(NCLIEN)
    CODPARAM$(26) = "TELE-EMP": DOCPARAM$(26) = TRIM$(TELETRANS$(NTRAI%))
    CODPARAM$(27) = "LOCA-TRA": DOCPARAM$(27) = LENTREGA(1).TEXT
    CODPARAM$(28) = "LAR-CUCO": DOCPARAM$(28) = TRIM$(Me.TXTFLETE)
    '
    CAPARDOC% = 28
    '
    ' LO QUE SIGUE ES PARA QUE EN CASO DE QUE HAYA UN SOLO CASILLERO
    ' DE IVA, TOTALICE TODOS LOS IMPORTES PARCIALES DE IVA EN UNO SOLO
    FORIVASEC$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "IMP-IVA1/FORMATO")))
    If FORIVASEC$ = "" Then
      DOCPARAM$(10) = TOTIVA(3).TEXT
      DOCPARAM$(14) = ""
    End If
    '
    ' LO QUE SIGUE ES PARA ORDEN DE DESPACHO
    If OPCI% = 2 Then
       CODPARAM$(1) = ""
       CODPARAM$(8) = ""
       CODPARAM$(9) = ""
       CODPARAM$(10) = ""
       CODPARAM$(11) = ""
       CODPARAM$(12) = ""
       CODPARAM$(13) = ""
       CODPARAM$(14) = ""
    End If
    ' FIN ORDEN DE DESPACHO
    '
    Call BLANARCH("!OBSERVOF")
'    Call FRATEXTO(Text13.TEXT, 32)
'    If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
'    If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
'    If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
'    URETE& = CARETEX%
    POOBSE% = XVALO(Mid$(ATRIFORM$(FORIPRE, "OBSERVA/FORMATO"), 2))
    Call FRATEXTO(Text13.TEXT, POOBSE%)
    For I = 1 To CARETEX%
        III& = I
        Call GRAREG("!OBSERVOF", RETEX$(I), III&)
    Next I
    URETE& = CARETEX%
    Call SETULTREG("!OBSERVOF", URETE&)
    Call CIERRARCH("!OBSERVOF")
    '
    Call SETPRINPORT("RESTORE")
     If FORIPRE$ <> "" Then
      'Call PRINTDOC("FORPECLI")   ' PEDIDO DE CLIENTE
      If OPCI% = 1 Then Call SETPRINPORT(PUEROVTA$)
      If OPCI% = 2 Then Call SETPRINPORT(PUERDESPA$)
      Call PRINTDOC("@" + FORIPRE$) ' PEDIDO DE CLIENTE
'       If PUERDESPA$ <> "" Then Call SETPRINPORT("RESTORE")
    End If
    Call SETPRINPORT("RESTORE")
    PUERDESPA$ = ""
    PUEROVTA$ = ""
    '
    For KK1% = 1 To CACOLTAB1%
      For KK2% = 1 To CAITAB1%
        TETABU1$(KK1%, KK2%) = ""
      Next KK2%
    Next KK1%
    '
    'FORIPRE$ = TRIM$(CONTROL$("", "FORDESPA"))
    'If FORIPRE$ <> "" Then
    '  If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
    '    TIPODOC$ = "Orden de Despacho"
    '    CANCELPRINT = 0
    '    If CONTROL$("ORDESPA", "SUPRINT") = "SI" Then
    '       CANCELPRINT% = 1
    '    End If
    '    '
    '    If InStr(EPAC$, "DOSIVAC") > 0 Then
    '      CANCELPRINT% = 0
    '      If OPRIPEDI.Check1(1).Value <> 1 Then
    '        CANCELPRINT% = 1
    '      End If
    '    End If
    '    '
    '    If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
    '    If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
    '    If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
    '    URETE& = CARETEX%
    '    Call SETULTREG("!OBSERVOF", URETE&)
    '    Call CIERRARCH("!OBSERVOF")
    '    '
    '    Call PRINTDOC("FORDESPA")   ' ORDEN DE DESPACHO
    '  End If
    'End If
    '
End Sub

Private Sub Total_GotFocus(Index As Integer)
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Sub GRAPEDIDO(OKX%, Optional FUERZAPROBACION$)
   '
   If CONTROL$("ADMIPED", "VERLICRE") = "SI" Then VERLICRE% = 1
   '
   OKX% = 0
   If MARCONOTA.Visible = True Then Exit Sub
   If MARSELPRESU.Visible = True Then Exit Sub
   If XVALO(Total(2).TEXT) < 0.01 Then
     If COMALTER%("Este Pedido es SIN CARGO !!!\Está Seguro de Registrarlo Igual ?\\Cancelar Pedido\Registrar Igual") <> 2 Then
       Exit Sub
     End If
   End If
   If ULTREG&("!CARDETPE") < 1 Then Exit Sub
   '
   Call CIERRARCH("PEDENCA")
   Call CIERRARCH("PEDDETA")
   Call CIERRARCH("PEDDETP")
   Call CIERRARCH("DESPAPED")
   '
   Call BLOQARCH("PEDENCA")
   Call BLOQARCH("PEDDETA")
   Call BLOQARCH("PEDDETP")
   Call BLOQARCH("DESPAPED")
   '
   NPP& = NUPEDIDO&
   If NPP& > XVALO(Text11.TEXT) Then
      Call CIERRARCH("PEDENCA")
      Call CIERRARCH("PEDDETA")
      Call CIERRARCH("PEDDETP")
      Call CIERRARCH("DESPAPED")
      Call COMUNI("FLEXOFT ha Re-Numerado el Pedido\como Pedido Número '" + TRIM$(Str$(NPP&)) + "'.\  \Verifique y Vuelva a Grabar.")
      Text11.TEXT = TRIM$(Str$(NPP&))
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIEN = XVALO(Text1.TEXT)
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
   If XVALO(TEXT21(3).TEXT) > 255 Then
      MERRO$ = "Tipo de Ingreso No Válido"
      GoTo 98
   End If
   '
   '\\\OT-06-0512-RG-27/11/06///
   If CONTROL$("ADMIPED", "OBLICOM") = "SI" Then
   '\\\OT-06-0512-RG-FIN///
    If TRIM$(Text10) = "" Then
       MERRO$ = "Falta Orden de Compra Cliente"
       GoTo 98
    End If
   End If
   '
   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   '
   NROPE& = XVALO(Text11.TEXT)
   REFE$ = " - " + TRIM$(Text10.TEXT)
   If REFE$ = " - " Then REFE$ = ""
   FEX = FEMI%
   REFE$ = REFE$ + " - " + FECHATEX$(FEX)
   LOMA% = 48 - Len(REFE$)
   REFE$ = Left$(TRIM$(RASOCLI$(NCLIEN)), LOMA%) + REFE$
   '
   STATUPC% = (-1)
   If Left$(CONTROL$("PEDCLI", "MODOAPRO"), 6) = "DIRECT" Then STATUPC% = 0
   'FUERZA A APROBACION POSTERIOR EN EL CASO QUE ALGUNA VALIDACION NO SEA SUPERADA SE GRABA COMO NO APROBADO
   If FUERZAPROBACION$ <> "" Then STATUPC% = (-1)
   '
20 Call ABREPEDCLI
   If CantRegistros("PEDENCA", "NroPed = " & NROPE&) > 0 Then
      Call CIERRARCH("PEDENCA")
      Call CIERRARCH("PEDDETA")
      Call CIERRARCH("PEDDETP")
      Call CIERRARCH("DESPAPED")
      NPP& = NUPEDIDO&
      Call COMUNI("FLEXOFT ha Re-Numerado el Pedido\como Pedido Número '" + TRIM$(Str$(NPP&)) + "'.\  \Verifique y Vuelva a Grabar.")
      Text11.TEXT = TRIM$(Str$(NPP&))
      Exit Sub
   End If
   '
      'ACA ENTRA SI ESTA MARCADALA OPCION DE VALIDAR SI SE REALIZARON CAMBIOS EN LISTAS DE PRECIOS AL CARGAR EL PEDIDO
   OKLISTAPRE% = 1
   If TRIM$(CONTROL("ADMIPED", "VALLISPR")) = "SI" Then
      U& = 0
      VUELTA% = 0
      For U& = 1 To ULTREG&("!CARDETPE")
         X$ = REGLEIDO$("!CARDETPE", U&)
         CIXI% = CVI(Left$(X$, 2))
         LISTA_ITEM% = CVS(Mid$(X$, 85, 4))
         PRE_VENTA# = CVS(Mid$(X$, 73, 4))
         RETORNO% = Vali_Derechos_Lpre(NCLIEN, CIXI%, LISTA_ITEM%, PRE_VENTA#)
         If RETORNO% < 1 Then
            
            If VUELTA% = 0 Then
              OKLISTAPRE% = 0
              Call REPLA(TXTLPRE$, "********* Item Sin Autorización *********", 11, 56): ACUTXTLPRE$ = ACUTXTLPRE$ + TXTLPRE$ + Chr$(10) + Chr$(13)
              VUELTA% = 1
            End If
            TXTLP$ = Left$(DESCRIT0(CIXI%), 24)
         
            Select Case RETORNO%
              Case 0 'SE CAMBIO LISTA DE PRECIOS SIN DERECHOS
                 TXTLP$ = TXTLP$ & " L.P.(" & LISTA_ITEM% & ")-Camb./Lst. S/Der."
              Case -1 'SE CAMBIO PRECIO SIN DERECHOS
                 TXTLP$ = TXTLP$ & " L.P.(" & LISTA_ITEM% & ")-Camb.Prc. S/Der."
              Case -2 'SE CAMBIO LISTA Y PRECIO SIN DERECHO
                 TXTLP$ = TXTLP$ & " L.P.(" & LISTA_ITEM% & ")-Camb./Lst y Prc. S/Der."
              Case -3  'EL ITEM NO PERTENECE A LA LISTA GRABADA
                 TXTLP$ = TXTLP$ & " L.P.(" & LISTA_ITEM% & ")-No Existe en L.Prc."
            End Select
            Call REPLA(TXTLPRE$, TXTLP$, 11, 56): ACUTXTLPRE$ = ACUTXTLPRE$ + TXTLPRE$ + Chr$(10) + Chr$(13)
         End If

      Next U&
      If OKLISTAPRE% < 1 Then STATUPC% = (-1)
   End If

   strsql = "Select * from Pedenca where NroPed = " & NROPE&
   Dim PedEnca As ADODB.Recordset
   Set PedEnca = New ADODB.Recordset
25 On Error Resume Next
   PedEnca.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   
   With PedEnca
     If Not (.EOF And .BOF) Then 'si lo encuentra
       PedEnca.Close
       Set PedEnca = Nothing
       MERRO$ = "Número Ocupado"
       GoTo 98
     End If
     ' aca se inserta uno nuevo
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CODIEMPR = CodiEmp%
     On Error GoTo 0
     !NROPED = NROPE&
     .Update
   End With
   PedEnca.Close
   Set PedEnca = Nothing
   '
   Set PedEnca = New ADODB.Recordset
35 On Error Resume Next
   PedEnca.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 35
   End If
   On Error GoTo 0
 
   With PedEnca
     If (.EOF And .BOF) Then 'si NO lo encuentra
        PedEnca.Close
        Set PedEnca = Nothing
        GoTo 20
     End If
     !REFERENCIA = REFE$
     !FECHEMIS = CVDAT(FEMI%)
     !NROCLIE = NCLIEN
     !RaSoClie = RASOCLI$(NCLIEN)
     !NROOCOM = TRIM$(Text10.TEXT)
     !DOMIENT = LENTREGA(0)
     !LOCAENT = LENTREGA(1)
     !TranspEnt = LENTREGA(2)
     !NTranspo = XVALO(Text19)
     !NSucent = XVALO(Text20)
     !Observa = Text13
     !FESOLeNT = Now
       '\\\OT-05-0236-WAR-02/05/05///
     !Status = STATUPC%              ' Status  de Aprobación
       '\\\OT-05-0236-FIN///
     !ListPre = XVALO(TEXT21(0))
     !CondPag = XVALO(TEXT21(1))
     !VENDED = XVALO(TEXT21(2))
     !TIPOINGRE = XVALO(TEXT21(3)) ' tipo de ingreso (telefono , fax etc)
     !OV_Motivo = XVALO(Text16)
     !OV_AUTORIZA2 = ""
     !OV_AUTORIZA1 = ""
     !PorComVend = XVALO(Text44.TEXT)
        If Left$(UCase$(Label21(4)), 5) = "LETRA" Then
           VALOR% = 0
           If TRIM$(Text44) <> "" Then VALOR% = Asc(Text44.TEXT + "A") - 64
           !PorComVend = VALOR%
        End If
     !Rec_ConPa = TRIM$(Label25)
     !TDescApli = TRIM$(PORDESCU.TEXT)
     !Dto_Ret_Cte = TRIM$(Label31)
     !VenRep = Left$(Label11, 1)   ' Venta o Reparacion
     '
     !CODCARPIMP = CODCARPE$ ' Guarda Nº de Carpeta de Exportaciones
     If CODCARPE$ <> "" Then ' Anotador de Carpeta de Exportaciones
        Call REPLA(TTXX$, HOS$, 1, 8)
        Call REPLA(TTXX$, USERNAME$, 69, 24)
        NR$ = SAFETEXT(Text11)
        Call REPLA(TTXX$, "Generacion de Orden de Venta Nº: " + NR$, 11, 56)
        Call AGREGA_ANOTACARPI(CODCARPE$, TTXX$)
        TTXX$ = Space(76)
        Call REPLA(TTXX$, "Cliente: " + Left$(RASOCLI$(NCLIEN), 64), 11, 56)
        Call AGREGA_ANOTACARPI(CODCARPE$, TTXX$)
        Call AGREGA_ANOTACARPI(CODCARPE$, String$(88, "="))
     End If
     '
     !FLETE_TEX = TRIM$(Left$(TXTFLETE, 64))
     .Update
   End With
   PedEnca.Close
   Set PedEnca = Nothing
   '
   ANOTA$ = ""
   TTXYZ$ = Space$(76)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, "Generación de Orden de Venta:", 11, 56)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
   '
   CAVER$ = TRIM$(Str$(NROPE&))
   While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
   While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
   IKITEM% = 0
   ORDITEM% = 0
   '
   '
   strsql = "Select * from PedDeta where NroPed=" & NROPE&
   Dim PedDeta As ADODB.Recordset
   Set PedDeta = New ADODB.Recordset
26 On Error Resume Next
   PedDeta.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
   End If
   On Error GoTo 0
   '
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     If CAN > 0.05 Then
       '
       IKITEM% = IKITEM% + 1
       IK$ = TRIM$(Str$(IKITEM%))
       CAVERI$ = CAVER$ + "-" + IK$
       '
       ' ESTO ES PARA PONER EL PRECIO DE LISTA EN LA B.DATOS
       CIXI% = CVI(Left$(X$, 2))
       If CIXI% > 0 Then
          If Left$(UCase$(ARCHIPRE$), 5) = "LIPRE" Then
             PLISTA = PRELISTA#(ARCHIPRE$, CIXI%)
             Call REPLA(X$, MKS$(PLISTA), 39, 4)
          End If
       End If
       '
       With PedDeta
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CODIEMPR = CodiEmp%
         On Error GoTo 0
           !NROPED = NROPE&
           !FECHEMIS = CVDAT(FEMI%)
           !NROCLIE = NCLIEN
           !RaSoClie = RASOCLI$(NCLIEN)
           !NROOCOM = TRIM$(Text10.TEXT)
             FENTRESOL% = CVI(Mid$(X$, 63, 2))
           !FESOLeNT = CVDAT(FENTRESOL%)
           !Status = STATUPC%              ' Status  de Aprobación
           !ListPre = XVALO(TEXT21(0).TEXT)
           !CondPag = XVALO(TEXT21(1).TEXT)
           !VENDED = XVALO(TEXT21(2).TEXT)
           !PorComVend = XVALO(Text44.TEXT)
           If Left$(UCase$(Label21(4)), 5) = "LETRA" Then
               VALOR% = 0
               If TRIM$(Text44) <> "" Then VALOR% = Asc(Text44.TEXT + "A") - 64
               !PorComVend = VALOR%
           End If

           !NuorItem = IKITEM%
           !CODIINT = CIXI%
           !CODIEXT = CODEXT$(CIXI%)
           !SubCod = Mid$(X$, 109, 8)
           !Descrip = DESCRIT0$(CIXI%)
           !TColor = Mid$(X$, 3, 16)
           !TTalle = Mid$(X$, 19, 8)
           !TLargo = CVS(Mid$(X$, 27, 4))
           !TAncho = CVS(Mid$(X$, 31, 4))
           !TEspesor = CVS(Mid$(X$, 35, 4))
           !VenRep = Left$(Label11, 1)   ' Venta o Reparacion
           !MoneEmis = MONEMI%
           !TipoCam = XVALO(TICAMBIS)
           !PreList = CVS(Mid$(X$, 39, 4))
           !CanAsig = CVS(Mid$(X$, 43, 4))
           !CANPED = CAN
           !PREUNID = CVS(Mid$(X$, 73, 4))     ' CVS(Mid$(X$, 51, 4))
           !ImpoTot = CVD(Mid$(X$, 55, 8))
           !CANTENT = 0
           !CanSaldo = CAN
           !FeUltEnt = CVDAT(0)
           !NroRemi = ""
           !TPedItem = CAVERI$
           '
           '\\\OT-8-766-19/09/08
           !ImpoAnti = 0
           !FApro_Nope = CVDAT(FECHANUM("01/01/80")) 'PARA EVITAR PROBLEMAS CON FORMATO REGIONAL
           !AcuPrepa = 0
           !AcuDespa = 0
           !TipoPed = "" ' se corroboro que este CAMPO es de  TEXTO
           !FinVigen = CVDAT(FECHANUM("01/01/80"))
           !StaDespa = 0
           !CBase = ""
           !DescLista = 0
           !DescCombi = ""
           !DescRubro = 0
           !DescTotal = 0
           !NUMEROSERIE = "" ' SE ESTABLECE VALOR POR DEFAULT PARA EL NUMERO DE SERIE
           '\\\OT-8-766-FIN
           'INTERVENCION AHP
           LOTEX1$ = TRIM$(Mid$(X$, 3, 16))
           !IDENLOTE = LOTEX1$
            NPROVX = CVS(Mid$(X$, 31, 4)) '(INT15-372)
           !Num_Prov = NPROVX '(INT15-372)
           'FIN AHP
           
           NLISST% = CVS(Mid$(X$, 85, 4))
           !LISTA_DE_ITEM = NLISST%
           
           TTXYZ$ = Space$(76)
           Call REPLA(TTXYZ$, "- " + CODEXT$(CIXI%), 11, 18)
           Call REPLA(TTXYZ$, DESCRIT0$(CIXI%), 29, 38)
           Call REPLA(TTXYZ$, CSTRING$(MKS$(CAN), 3, 10, 1, 2), 57, 10)
           ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
           '
           'INTERVENCION AHP
           If LOTEX1$ <> "" Then
             'GUARDO EN EL ANTOADOR EL NUMERO DE LOTE
             TELOTE$ = "Lote Reservado: " + LOTEX1$
             TELOTEXX$ = Space$(76)
             Call REPLA(TELOTEXX$, TELOTE$, 29, 38)
             ANOTA$ = ANOTA$ + TELOTEXX$ + Chr$(10) + Chr$(13)
           Else '(INT15-372)
             'GUARDO EN EL ANTOADOR EL NUMERO DE PROV.
             NUMEPROVTEX$ = "Proveedor N°: " + TRIM$(Str$(NPROVX)) '(INT15-372)
             TELOTEXX$ = Space$(76)
             Call REPLA(TELOTEXX$, NUMEPROVTEX$, 29, 38)
             ANOTA$ = ANOTA$ + TELOTEXX$ + Chr$(10) + Chr$(13)
           
           End If
           'FIN AHP
         .Update
         '
       End With
       '
       NUORIT% = CVI(Mid$(X$, 123, 2))
       NUPRESU& = CVS(Mid$(X$, 125, 4))
       If NUPRESU& > 0 Then
         Call ABREPRESVEN
'         'jandy
'         strsql = "Update PresuVen set FechVent='" & FETEXCOR1$(FEMI%) & "'"
'         strsql = strsql & ", NroPedi= " & NROPE&
'         strsql = strsql & " where NroPres= " & NUPRESU&
'         strsql = strsql & " and NuOrdIt= " & NUORIT%
'         FLEXCONN.Execute (FILTSQL$(strsql))
            VENCARPE$ = CONTROL$("", "VENCARPE")
            'AHP
            'SE AGREGA ESTO POR QUE CUANDO SE ASIGNAN LOTES SE PUEDE PERDER EL ORDEN ORIGINAL
            If UCase$(TRIM$(VENCARPE$)) = "CARDETLO" Then
                  CONDI$ = "NroPres = " & NUPRESU& & " and CodiInt= " & CIXI%
                  Call ACTUREGISTRO("PRESUVEN", "FECHVENT", CONDI$, FETEXCOR1$(FEMI%), REGAF&, "NOMESS")
                  Call ACTUREGISTRO("PRESUVEN", "NroPedi", CONDI$, NROPE&, REGAF&, "NOMESS")
                  GoTo 85  ' SALTA
            'FIN AHP
            Else
                '
                strsql = "select FechVent, NroPedi from PresuVen where NroPres = " & NUPRESU& & " and NuOrdIt= " & NUORIT%
                '
                Dim presup As ADODB.Recordset
                Set presup = New ADODB.Recordset
80              On Error Resume Next
                presup.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                   On Error GoTo 0
                   Call CapturaErrorOpen
                   GoTo 80
                End If
                On Error GoTo 0
                With presup
                  If .EOF = False Then
                     !FechVent = CVDAT(FEMI%) ' SE CAMBIO POR CVDAT POR QUE EN TANIT DIO UN ERROR DE TIPO INCORRECTO (02/12/13)
                     !NroPedi = NROPE&
                     .Update
                  End If
                End With
                presup.Close
                Set presup = Nothing
            End If
       End If
       '
85     If NUPRESU& < 0 Then
         REPRESU& = Abs(NUPRESU&)
         UREPREP& = ULTREG&("PRESREPA")
         If REPRESU& <= UREPREP& Then
           YY$ = REGLEIDO$("PRESREPA", REPRESU&)
           If CVI(Mid$(YY$, 67, 2)) = NCLIEN Then    ' Coincide cliente
             If CVI(Mid$(YY$, 69, 2)) = CIXI% Then   ' Coincide Articulo
               Call REPLA(YY$, MKS$((-1) * NROPE&), 125, 4)
               Call GRAREG("PRESREPA", YY$, REPRESU&)
             End If
           End If
         End If
         '
         'LE ASIGNA EL NUMERO DE PEDIDO A LA REPARACION
         Call CREATABLA_OFABRIC
         CONDI$ = "NUMREPA=" & Abs(NUPRESU&)
         Call ACTUREGISTRO("REPARA", "NROPED", CONDI$, NROPE&, REG&, "NOMESS")
         '
       End If
       '
     End If
   Next U&
   '
   PedDeta.Close
   Set PedDeta = Nothing
   '
   U& = 0
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     If CAN > 0.05 Then
        If CVS(Mid$(X$, 89, 4)) > 0 Then
            IKITEM% = IKITEM% + 1
            Call AGREGABONIFICACION(U&, IKITEM%)
        End If
     End If
   Next U&
   '
   
   
   
   
'   'GRABACION EN BALAREQMO (CELSIUS)
'   FORMATPEDI$ = Right$("000000" + TRIM$(Str$(NROPE&)), 6)
'   SQLX$ = "SELECT * FROM PEDDETA"
'   SQLX$ = SQLX$ + " WHERE NROPED = " & NROPE&
'   Set PEDTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'   With PEDTMP
'   Do While Not .EOF
'     CI% = XVALO(!CodiInt)
'     If CI% > 0 Then
'       CANTI = XVALO(!CanPed)
'       If CANTI > 0.05 Then
'        ORDITEM% = XVALO(!NUORITEM)
'        DESTIN$ = "OV-" & FORMATPEDI$ & "-1"
'        IDENREQ$ = DESTIN$ & "/" & TRIM$(Str$(ORDITEM%))
'        '
'        CONDI$ = "IDENREQUE = '" & IDENREQ$ & "'"
'        '
'        Call CreaRegistro("BALAREQMO", "IDENREQUE", IDENREQ$, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "NUMECOMPRO", CONDI$, NROPE&, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "CODSUCU", CONDI$, 0, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "NUORITEM", CONDI$, ORDITEM%, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "COD_INTE", CONDI$, CI%, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "IDENLOTE", CONDI$, "", REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "CANTIDAD", CONDI$, CANTI, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "KILO_REQUE", CONDI$, CANTI, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "DESTINO", CONDI$, DESTIN$, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "CODICOMPRO", CONDI$, "OV", REGAF&, "NOMESS")
'       End If
'       '
'     End If
'     .MoveNext
'
'   Loop
'   End With
'   PEDTMP.Close
'   Set PEDTMP = Nothing
   '
   
   'FIN GRABACION BALAREQMO

   'VALIDO SI TIENE MOTIVO DE PEDIDO (MEDITEA) AGREGO EN EL ANOTADOR
   NUMOT% = NUMOTIVO%(NROPE&)
   If NUMOT% > 0 Then
      DESCMOTI$ = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
      ANOTA$ = ANOTA$ + Space$(10) + String$(56, "-") + Chr$(10) + Chr$(13)
      TTXYZ$ = ""
      Call REPLA(TTXYZ$, "Motivo: " & NUMOT% & " (" & TRIM$(DESCMOTI$) & ")", 11, 66)
      ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
   End If
   '
   ANOTA$ = ANOTA$ + Space$(10) + String$(56, "-") + Chr$(10) + Chr$(13)
   Call FRATEXTO(Text13.TEXT, 32)
   For KKI% = 1 To CARETEX%
     TTXX$ = Space$(92)
     Call REPLA(TTXX$, RETEX$(KKI%), 16, 32)
     If KKI% = 1 Then Call REPLA(TTXX$, "Observaciones:", 1, 15)
     ANOTA$ = ANOTA$ + Space$(10) + TTXX$ + Chr$(10) + Chr$(13)
   Next KKI%
   '
   If InStr(UCase$(FUERZAPROBACION$), "CAMBIO_PAGO") > 0 Then
      CONDIP% = CPAGCLI%(NCLIEN)
      CONDIPEDIDO% = XVALO(TEXT21(1))
      TEXTOPAGO$ = "Se Cambió Condición de Pago " & CONDIP% & "  Por la " & CONDIPEDIDO%
      Call REPLA(TXX1$, TEXTOPAGO$, 11, 56): ANOTA$ = ANOTA$ + TXX1$ + Chr$(10) + Chr$(13)
   End If
   '
   If InStr(UCase$(FUERZAPROBACION$), "CAMBIO_LISTA") > 0 Then
      LISTAX% = Str$(LIPRCLI%(NCLIEN))
      LISTAXPEDI% = XVALO(TEXT21(0))
      TEXTOLISTAP$ = "Se Cambió Lista de Precio " & LISTAX% & "  Por la " & LISTAXPEDI%
      Call REPLA(TXX1$, TEXTOLISTAP$, 11, 56): ANOTA$ = ANOTA$ + TXX1$ + Chr$(10) + Chr$(13)
   End If
   '
   If InStr(UCase$(FUERZAPROBACION$), "FACTURA_VENCIDA") > 0 Then Call REPLA(TXX1$, "Se Detectaron Facturas Vencidas", 11, 56): ANOTA$ = ANOTA$ + TXX1$ + Chr$(10) + Chr$(13)
   If InStr(UCase$(FUERZAPROBACION$), "FECHA ÚLTIMA FACTURA:") > 0 Then Call REPLA(TXX1$, "Supera Fecha Máxima de Última Compra", 11, 56): ANOTA$ = ANOTA$ + TXX1$ + Chr$(10) + Chr$(13)
   '
   'ANOTADOR DE DETALLES DE VALIDACION DE DERECHOS SOBRE LISTA DE PRECIOS
   If OKLISTAPRE% < 1 Then
      ANOTA$ = ANOTA$ + ACUTXTLPRE$
   End If
   '
   ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
   '
   'jandy sql
   SQLX$ = " SELECT * FROM PEDENCA "
   SQLX$ = SQLX$ + "WHERE  NroPed = " & NROPE & ""
   Dim PEDDDENCA As ADODB.Recordset
   Set PEDDDENCA = New ADODB.Recordset
27 On Error Resume Next
   PEDDDENCA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 27
   End If
   On Error GoTo 0
   On Error Resume Next
'  With Pedenca
   With PEDDDENCA
      If Err < 1 Then
        On Error GoTo 0
        If IsNull(!AnotaPed) = True Or !AnotaPed = "" Then
            !AnotaPed = ANOTA$
          Else
            VarAnotaped$ = !AnotaPed
            !AnotaPed = VarAnotaped$ + ANOTA$
         End If
          .Update
        End If
    End With
   '
   CODICL$ = CODICLI$(NCLIEN)
   ANOTA1$ = Mid$(ANOTA$, 1, Len(ANOTA$) - 2)
   If CODICL$ <> "" Then Call AGREGA_ANOTAPROSPECT(CODICL$, ANOTA1$) 'AGREGA ANOTACION EN EL PROSPECT
   Call ABRECLIEN
   '
   'VALIDACION DE SETUP SI SE GUARDA LOS DATOS DE ENTREGA (ACTUALIZANDO LOS DATOS DEL CLIENTE)
   If CONTROL("PEDIDO", "ACTULEN") <> "NO" Then
'      strsql = "Update Deudor12 set Lentre0_Cli ='" & TRIM$(LENTREGA(0).TEXT) & "'"
'      strsql = strsql & ", Lentre1_Cli ='" & TRIM$(LENTREGA(1).TEXT) & "'"
'      strsql = strsql & ", Lentre2_Cli ='" & TRIM$(LENTREGA(2).TEXT) & "'"
'      strsql = strsql & " where Nume_Cli = " & NCLIEN
'      FLEXCONN.Execute (FILTSQL$(strsql))
       strsql = "select Lentre0_Cli,Lentre1_Cli,Lentre2_Cli from Deudor12 where Nume_Cli = " & NCLIEN
       Dim cliente As ADODB.Recordset
       Set cliente = New ADODB.Recordset
90     On Error Resume Next
       cliente.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 90
       End If
       On Error GoTo 0
       With cliente
            !Lentre0_Cli = TRIM$(LENTREGA(0).TEXT)
            !Lentre1_Cli = TRIM$(LENTREGA(1).TEXT)
            !Lentre2_Cli = TRIM$(LENTREGA(2).TEXT)
            .Update
       End With
       cliente.Close
       Set cliente = Nothing
   End If
   '
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
   TPED# = XVALO(Total(3).TEXT)
   OKX% = 1
   If VERLICRE% = 1 Then
     LMCR# = LIMICRE(NCLIEN)
     CHPA# = CHEPEACRE#(NCLIEN)
     SDEU# = SALDCLI#(NCLIEN)

     If SDEU# + CHPA# + TPED# > LMCR# Then
       TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2))
       If TLIMI$ = "" Then TLIMI$ = "0.00"
       ATTX$ = "ATENCION: Este Pedido Supera\el Límite de Crédito Asignado\para esta Cuenta de $ " + TLIMI$
       If DERACCE%("AUTORENT/NOMESS", "Autorizar Pedidos con Exceso Credito") < 2 Then
           Call COMUNI(ATTX$ + ".\  \Informe al Supervisor de Ventas\para Autorizar Facturación y Despacho.")
         Else
           AUTORENT$ = "NO"
           If COMALTER%(ATTX$ + "\   \Puede Autorizar la Entrega\con Crédito Excedido.\\Continuar\Autorizar Entrega") = 2 Then AUTORENT$ = "SI"
           RFF$ = RECFILT$("AUTORENT", 1, MKS$(NROPE&))
           REGAUT& = 0: If Len(RFF$) >= 4 Then REGAUT& = CVS(Left$(RFF$, 4))
           AUTORI$ = MKS$(NROPE&) + AJUSTI$(USERNAME$, 26) + AUTORENT$
           If REGAUT& > 0 Then
               Call GRAREG("AUTORENT", AUTORI$, REGAUT&)
             Else
               Call REGAPP("AUTORENT", AUTORI$)
           End If
           Call CIERRARCH("AUTORENT")
       End If
     End If
   End If
   '
   If EXEDELIMCC%(NCLIEN, TPED#, DIFER) > 0 Then
       LIMCCTE$ = TRIM$(FORMATNUM$(LCred_CCte(NCLIEN), "f12.2")): If LIMCCTE$ = "" Then LIMCCTE$ = "0.00"
       DifADep$ = TRIM$(FORMATNUM$(DIFER, "f12.2")): If DifADep$ = "" Then DifADep$ = "0.00"
       TOTAPE$ = TRIM$(FORMATNUM$(TPED#, "f12.2")): If TOTAPE$ = "" Then TOTAPE$ = "0.00"
       TX$ = "ATENCION: Este Pedido Supera\el Límite de Crédito Asignado en C.Cte.\para esta Cuenta de $ " + LIMCCTE$
       TX$ = TX$ + "\Puede Ahora Enviar un Mail Solicitando la Cancelación\ del Excedido Incluyendo el Presente Pedido \Por La Diferencia ($ " + DifADep$ + ") \\Enviar Mail\Continuar"
110    If COMALTER%(TX$) <> 2 Then
            MAI_DESTI$ = EMAILCLI$(NCLIEN)
            If TRIM$(MAI_DESTI$) = "" Or InStr(MAI_DESTI$, "@") < 1 Then
120           MAI_DESTI$ = LCase$(TRIM$(InputBox$(" Mail de Cliente", "Intgrese Mail de Destino", MAI_DESTI$)))
              If TRIM$(MAI_DESTI$) = "" Then
                GoTo 110
              ElseIf InStr(MAI_DESTI$, "@") < 1 Then
                  Call COMUNI("Mail Incorrecto")
                  GoTo 120
              End If
            End If
            TXT$ = "Solicitamos realizar el Depósito de $ " & DifADep$
            TXT$ = TXT$ + " Para cubrir la diferencia de su límite de cuenta corriente "
            TXT$ = TXT$ + "de $ " & LIMCCTE$ & " y poder autorizar la entrega del pedido Nro.: " & NROPE&
            TXT$ = TXT$ + " por  importe de $ " & TOTAPE$
            MAI_DIREL$ = MAI_DESTI$
            MAI_ASUNT$ = "Solicitud de Interdepósito para Pedido Número: " & NROPE&
            MAI_TEXTO$ = TXT$
            '
            Call GENERAMAIL
       End If

   End If
   '
   Screen.MousePointer = 1
   Exit Sub
   
   '
98 Call CIERRARCH("PEDENCA")
   Call CIERRARCH("PEDDETA")
   Call CIERRARCH("PEDDETP")
   Call CIERRARCH("DESPAPED")
   '
   Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub

Private Sub Total_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Index = 0 Then
       On Error Resume Next
       PORDESCU.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Total_LostFocus(Index As Integer)
   If Index = 0 Then
     TTT1 = XVALO(Total(0).TEXT)
     Total(0) = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   End If
End Sub

Sub CARDETPEDI(MODO%)
   '
   HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC/ARCHECO(1)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle Pedido de Cliente"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top + 400))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left))
   '
   Timer1.Enabled = True
   CONLOTIND$ = UCase(TRIM$(CONTROL$("ADMIPED", "VALOIND")))  'CONTROL DE VALIDACION DE VERIFICACION DE LOTE INDIVISIBLE
10 VENCARPE$ = CONTROL$("", "VENCARPE")
   'AHP
   If UCase$(TRIM$(VENCARPE$)) = "CARDETLO" Then
     ATRI$ = ATRI$ + "/POSEDIT(1)=2" ' VENTANA CON LOTES AHP
   End If
   'FIN AHP
   '
   If TRIM$(VENCARPE$) = "" Then VENCARPE$ = "CARDETP"
   PRF$ = VENCARPE$    ' If FORMAPED% = 2 Then PRF$ = "CARDETQ"
   XXX% = VENTABU%(PRF$ + ATRI$)
   NOFRESHPRE% = 0
   Call CALTOPED
   If ECOPRECIO% >= 0 Then
     If XVALO(Total(2).TEXT) < 0.005 Then
       If TRIM$(Text13.TEXT) = "" Then
         Text13.TEXT = "SIN CARGO !!!"
       End If
     End If
   End If
   Timer1.Enabled = False
   '
   If XXX% < 0 Then
      Call CIERRARCH("!CARDETPE")
      Call BLANARCH("!CARDETPE")
      Total(0).TEXT = ""
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
20 If ULTREG&("!CARDETPE") < 1 Then
      Screen.MousePointer = 1
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   Call CargaBonificacion
   Screen.MousePointer = 11   ' Registrando Pedido ...
   FIN& = ULTREG&("!CARDETPE")
   If FIN& > 0 Then
     Y$ = REGLEIDO$("!CARDETPE", FIN&)
     CI% = CVI(Left$(Y$, 2))
     If CI% < 1 Or CI% > NUMAS% Then
       Call SETULTREG("!CARDETPE", FIN& - 1)
       Call CIERRARCH("!CARDETPE")
       Call MENSERR(24, "Código No Reconocido o No Válido\Debe Correr Control de Consistencia de Maestro.")
       GoTo 20
     End If
   End If
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   CODSINSTO$ = ""
   VALIREPE% = 1 ' BANDEREA DE VALIDACION DE PRECIOS
   If CONTROL("ADMIPED", "VALIREPE") = "NO" Then VALIREPE% = 0
   '
   VALICODIENCLIENTE% = 0 'BANDERA PARA VALIDAR SI EL CODIGO FUE PEDIDO POR EL CLIENTE
   If CONTROL("ADMIPED", "VACOPED") = "SI" Then VALICODIENCLIENTE% = 1
   '
   FENTREGEN% = FECHANUM(Text17)
   For U& = 1 To FIN&
     Y$ = REGLEIDO$("!CARDETPE", U&)
     CI% = CVI(Left$(Y$, 2))
     If CI% < 1 Or CI% > NUMAS% Then
       Call MENSERR(24, "Código no Válido")
       Call REPLA(Y$, MKI$(0), 1, 2)
       Call GRAREG("!CARDETPE", Y$, U&)
       GoTo 10
     End If
     '
     FENTRESOL% = CVI(Mid$(Y$, 63, 2))
     If FENTRESOL% < FENTREGEN% Then
        FENTRESOL% = FENTREGEN%
        Call REPLA(Y$, MKI$(FENTRESOL%), 63, 2)
        Call GRAREG("!CARDETPE", Y$, U&)
     End If
     If FENTRESOL% < HOII% Then
        FENTRESOL% = HOII%
        Call REPLA(Y$, MKI$(FENTRESOL%), 63, 2)
        Call GRAREG("!CARDETPE", Y$, U&)
     End If
     '
     '  verifica codigos repetidos
      If VALIREPE% > 0 Then
       For W& = 1 To U& - 1
         Y$ = REGLEIDO$("!CARDETPE", W&)
         If CVI(Left$(Y$, 2)) = CI% Then
           If CVI(Mid$(Y$, 63, 2)) = FENTRESOL% Then
             Screen.MousePointer = 1
             TTXX$ = "Código " + TRIM$(CODEXT$(CI%)) + "\Duplicado en el Pedido."
             If COMALTER%(TTXX$ + "\\Corregir\Continuar") <> 2 Then
               GoTo 10
             End If
           End If
         End If
       Next W&
     End If
     '
     If VALICODIENCLIENTE% > 0 Then
        NCLIE = XVALO(Text1)
        CONDI$ = "NROCLIE = " & NCLIE & " AND CODIINT = " & CI%
        CODPEDIDO& = CantRegistros&("PEDDETA", CONDI$, "NOMESS")
        If CODPEDIDO& < 1 Then
           TTXX$ = "Código " + TRIM$(CODEXT$(CI%)) + "\" & DESCRIT0$(CI%) & "\"
           TTXX$ = TTXX$ + "Nunca Fue Solicitado Por Este Cliente.\ "
           TTXX$ = TTXX$ + "¿Desea Continuar?"
           If COMALTER%(TTXX$ + "\\Sí, Continuar\No, Cancelar") > 1 Then
              GoTo 10
           End If
        End If
     End If
     X$ = REGLEIDO$("!CARDETPE", U&)
     CI% = CVI(Left$(X$, 2))
     TTXX$ = Space$(96)
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     '
     CAN = CVS(Mid$(X$, 47, 4))
     If CAN > 999999 Then
       Call MENSERR(24, "Cantidad Excesiva - Max 999999")
       Call REPLA(X$, MKS$(0), 47, 4)
       Call GRAREG("!CARDETPE", X$, U&)
       GoTo 10
     End If
     '
'     If UCase$(TRIM$(VENCARPE$)) <> "CARDETME" Then ' PREMOTEC
        If CAN < 0.5 Then
          Call MENSERR(24, "Falta Cantidad")
          Call REPLA(X$, MKS$(0), 47, 4)
          Call GRAREG("!CARDETPE", X$, U&)
          GoTo 10
        End If
'     End If
     '
     If CONLOTIND$ = "SI" Then      'SEGUN LA CONFIGRACION REALIZA EL CONTROL DEL LOTE INDIVISIBLE
       If Abs(CAN - REDON_INDIVIS(CI%, CDbl(CAN), "VENTA")) >= 0.05 Then
         TTXX$ = "Cantidad Ajustada a Lotes Indivisibles\en Código '" + CODEXT$(CI%) + "'"
         Call MENSERR(24, TTXX$)
         Call REPLA(X$, MKS$(REDON_INDIVIS(CI%, CDbl(CAN), "VENTA")), 47, 4)
         Call GRAREG("!CARDETPE", X$, U&)
         GoTo 10
       End If
     End If
     '
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 99.9 Then
       Call MENSERR(24, "Descuento no Válido")
       Call REPLA(X$, MKS$(0), 69, 4)
       Call GRAREG("!CARDETPE", X$, U&)
       GoTo 10
     End If
     '
     If ECOPRECIO% < 0 Then
       PRECARGA = PRELISTA#(ARCHIPRE$, CI%)
       Call REPLA(X$, MKS$(PRECARGA), 39, 4)
       Call GRAREG("!CARDETPE", X$, U&)
     End If
     PRECARGA = CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     CADECI% = DECILISTA%(XVALO(Mid$(ARCHIPRE$, 6, 3)))
     VAMINI1 = 0.5 / (10 ^ CADECI%)
     If PRECARGA < VAMINI1 Or PRECARGA > 99999.99 Then
       If MODO% < 2 Then
         If ECOPRECIO% >= 0 Then
           Call REPLA(X$, MKS$(0), 39, 4)
           Call GRAREG("!CARDETPE", X$, U&)
           MODO% = 2
           If COMALTER%("Artículo '" + TRIM$(CODEXT$(CI%)) + "' sin Precio.\\Corregir\Continuar") <> 2 Then
              NOFRESHPRE% = 1
              GoTo 10
           End If
         End If
       End If
     End If
     '
     CANTPEDIDA# = CVS(Mid$(X$, 47, 4))
     If CONTROL$("PEDIDO", "VERISTO") = "SI" Then
        If DESCOMPO%(CI%) > 0 Then
          Call AREXPLO(CI%, 1, 1)
          For KKJ% = 1 To CAIT%
             CICO% = CIJ%(KKJ%)
             CACICO = USOI(KKJ%) * CANTPEDIDA#
             If CACICO > STOCKACT(CICO%) Then
                If CODSINSTO$ <> "" Then
                    CODSINSTO$ = CODSINSTO$ & " ; " & CODEXT(CICO%)
                  Else
                    CODSINSTO$ = CODEXT(CICO%)
                End If
             End If
          Next KKJ%
          GoTo 30
        End If
        '
        If CANTPEDIDA# > STOCKACT(CI%) Then
           If CODSINSTO$ <> "" Then
              CODSINSTO$ = CODSINSTO$ & " ; " & CODEXT(CI%)
              Else
              CODSINSTO$ = CODEXT(CI%)
           End If
        End If
     End If
     '
30   Call REPLA(X$, MKS$(PRECARGA), 39, 4)
     Call REPLA(X$, MKS$(PRECARGA), 51, 4)
     Call REPLA(X$, MKD$(TOT#), 55, 8)
     Call REPLA(X$, MKS$(PRECARGA), 65, 4)
     Call REPLA(X$, MKS$(PREUNI), 73, 4)
     Call REPLA(X$, MKD$(TOT#), 77, 8)
     
     'GUARDA LA LISTA DE PRECIOS SI YA TIENE MANTIENE SI NO ESCRIBE
     '
     If Abs(CVS(Mid$(X$, 85, 4))) = 0 Then
        LPRE = XVALO(Mid$(ARCHIPRE$, 6, 3))
        If LPRE = 0 Then LPRE = -1
     Else
        LPRE = CVS(Mid$(X$, 85, 4))
     End If
     Call REPLA(X$, MKS$(LPRE), 85, 4)
     '
'     If UCase$(TRIM$(VENCARPE$)) <> "CARDETME" Then
'      LPREXX = CVS(Mid$(X$, 85, 4))
'      'CANTM2 = CANTIXM2( LPREXX)
'      If CANTM2 > 0 Then
'        ANTIU = CANTI_INGRE_EN_M2 / CANTM2
'        If Int(CANTIU) < CANTIU Then CANTIU = Int(CANTIU) + 1
'        Call REPLA(X$, MKS$(CANTIU), 47, 4)
'      Else
'        Call REPLA(X$, MKS$(CANTIU), 47, 4)
'      End If
'     End If
     Call GRAREG("!CARDETPE", X$, U&)
     '
     
   Next U&
   '
   If InStr(UCase(EPAC$), "PREMOTEC") > 0 Then
      Call APP_PREMOTEC
   End If
   If CODSINSTO$ <> "" Then
      OPX% = COMALTER("Los Siguientes Codigos Tienen Stock Insuficiente \" & CODSINSTO$ & "\\Corregir\Continuar")
      If OPX% <> 2 Then GoTo 10
   End If
   '
   'ACA LLAMAR A NUEVA VENTANA Y CARGAR EL RES
   Call CARLISPE
   Call CIERRARCH("!CARDETPE")
   Call CALTOPED
   'VALIDACION DE CONTROL DE MARCA DE MODO DE DESCUENTO
   'CAMBIA EL DESCUENTO DE ACUERDO A LA TABAL DESMONTO EN EL CASO QUE SEA LISTA 1
   If CONTROL("PEDIDO", "MODODESC") = "MONTO" And XVALO(TEXT21(0)) = 1 Then
       Call DESCUMONTO(Me)
   End If
   '
   List1.Refresh
'   Call REMPLAGRID(Me, GRID2, MARCONOTA, Label22, List1)

   On Error Resume Next
   Command14.SetFocus
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub

Sub CARLISPE()
   List1.Clear
   FIN& = ULTREG&("!CARDETPE")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     TTXX$ = Space$(96)
     CI% = CVI(Left$(X$, 2))
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = CVS(Mid$(X$, 39, 4))
       DLX% = DECILISTA(XVALO(Mid$(ARCHIPRE$, 6, 3)))    ' DECIMALES DE LA LISTA DE PRECIOS
     PREUNI = ROUND((PRECARGA * (100 - PORDESCUIN) / 100), DLX%)
     PREUNI_SINRED = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = ROUND(CDbl(CAN * PREUNI))
     '
     If CAN <= 9999 Then
       Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 1, 2), 1, 6)
     Else
       Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
     End If
     '
     Call REPLA(TTXX$, CODEXT$(CI%), 8, 15)
     FEX = FENTRESOL%
     If ECOPRECIO% < 0 Then
          Call REPLA(TTXX$, DESCRIT0$(CI%), 24, 52)
          Call REPLA(TTXX$, FECHATEX$(FEX), 78, 8)
        Else
          Call REPLA(TTXX$, DESCRIT0$(CI%), 24, 22)
          Call REPLA(TTXX$, Left$(FECHATEX$(FEX), 5), 47, 5)
     End If
     '
     PRELIX = PRELISTA#(ARCHIPRE$, CI%)
     If PRELIX < PRECARGA Then PRELIX = PRECARGA
     If PRELIX < PREUNI Then PRELIX = PREUNI
       DLX% = DECILISTA(XVALO(Mid$(ARCHIPRE$, 6, 3)))    ' DECIMALES DE LA LISTA DE PRECIOS
     PORDESCUIN = 0: If PRELIX > 0.005 Then PORDESCUIN = 100 * (PRELIX - PREUNI_SINRED) / PRELIX
     ' ESTO SE MODIFICO POR PROBLEMA CON CARGA DE PEDIDOS DE GABAL
     ' PORDESCUIN = 0: If PRELIX > 0.005 Then PORDESCUIN = 100 * (PRELIX - PREUNI) / PRELIX
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PRELIX), 3, 8, DLX%, 2), 52, 8)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 61, 5)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, DLX%, 2), 66, 8)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 2, 2), 76, 10)
     List1.AddItem TTXX$
     '
   Next U&

End Sub
'
Sub CALTOPED()
   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIEN = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIEN > 0 Then
       PVCLI% = PIVACLI%(NCLIEN)
       If PVCLI% = 2 Then RNI% = 1
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   SUTOTPED# = 0: SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
     Dim IMPTASIVA#()
     ReDim IMPTASIVA#(ULTREG&("GRUPIVA"))
     '
     DESC0$ = TRIM$(PORDESCU.TEXT)
     'SI TIENE DESCUENTO DE RETIRA CLIENTE LO CONCATENA PARA EL CALCULO IMPOSITIVO
     If XVALO(Label31) > 0 Then DESC0$ = TRIM$(Str$(XVALO(Label31))) & "+" & DESC0$
     COEF1# = 1 - COEFDESCU(DESC0$) / 100
   TOTITEMS% = 0
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CIXXI% = CVI(Left$(X$, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 99 Then PORDESCUIN = 0
     If CAN > 0.05 Then
       PRELI = CVS(Mid$(X$, 39, 4))
       CDX% = DECILISTA(LIPRACTI%)
       PREUNI = ROUND(PRELI * (100 - PORDESCUIN) / 100, CDX%)
       TOT# = ROUND(CDbl(CAN * PREUNI), 2)
       SUTOTPED# = SUTOTPED# + TOT#
       SUIVA# = SUIVA# + (Int(TOT# * PORCEIVA(CIXXI%) + 0.5)) / 100
       '\\\OT-05-0560-WAR-03/08/05///
       TODE1# = TOT# * COEF1#
       GRIVA% = GRUPIVA%(CIXXI%)
       IMPTASIVA#(GRIVA%) = IMPTASIVA#(GRIVA%) + (Int(TODE1# * PORCEIVA(CIXXI%) + 0.5)) / 100
       Label13 = CSTRING$(MKD$(IMPTASIVA#(0)), 3, 11, 2, 2)
       Label14 = CSTRING$(MKD$(IMPTASIVA#(1)), 3, 11, 2, 2)
       Label20 = CSTRING$(MKD$(IMPTASIVA#(2)), 3, 11, 2, 2)
       '\\\OT-05-0560-WAR-FIN///
       Call REPLA(X$, MKS$(PREUNI) + MKD$(TOT#), 51, 12)
       Call GRAREG("!CARDETPE", X$, U&)
       TOTITEMS% = TOTITEMS% + 1
     End If
   Next U&
   '
   Label18 = TRIM$(Str$(TOTITEMS%))
   '
   Total(0).TEXT = CSTRING$(MKD$(SUTOTPED#), 3, 11, 2, 2)
   Total(0).Refresh
   '
   'CALCULO DE RECARGO POR CONDICION DE PAGO
   SIGNO$ = ""
   RECPORPAGO = 0
   RECACON = COEFDESCU(Label25)
   If Abs(XVALO(RECACON)) > 0 Then RECPORPAGO = XVALO(RECACON) * (-1)
   If XVALO(PORDESCU.TEXT) > 0 Then SIGNO$ = "+"
   DESC$ = RECPORPAGO & SIGNO$ & TRIM$(PORDESCU.TEXT)
   '
   'DESC$ = TRIM$(Label25) & SIGNO$ & TRIM$(PORDESCU.TEXT)
   TTT1 = XVALO(Total(0).TEXT) * COEFDESCU(DESC$) / 100
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   '
   'AGREGADO PARA GABAL DESDE ACA
   DESC1 = (XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)) * COEFDESCU(Label31) / 100
   Text141 = CSTRING$(MKS$(DESC1), 3, 11, 2, 2)
   
   TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT) - XVALO(Text141) ' ESTA ULTIMA RESTA GABAL
   Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
   '
   On Error Resume Next
   TOTIVA(0).TEXT = CSTRING$(MKD$(APLIVA% * SUIVA# * XVALO(Total(2).TEXT) / XVALO(Total(0).TEXT)), 4, 11, 2, 2)
   On Error GoTo 0
   TOTIVA(2).TEXT = CSTRING$(MKD$(CDbl(XVALO(TOTIVA(0).TEXT) * 0.5 * RNI%)), 4, 11, 2, 2)
   TOTIVA(3).TEXT = CSTRING$(MKD$(CDbl(XVALO(TOTIVA(0).TEXT) + XVALO(TOTIVA(2).TEXT))), 4, 11, 2, 2)
   '
   Call CALPERIB
   '
   Total(3).TEXT = CSTRING$(MKD$(CDbl(XVALO(Total(2).TEXT) + XVALO(TOTIVA(3).TEXT)) + XVALO(Text15)), 4, 11, 2, 2)
   '
End Sub

Private Sub TOTIVA_Change(Index As Integer)
   'If Index = 3 Then
   '  Total(3).TEXT = Str$(XVALO(Total(2).TEXT) + XVALO(TOTIVA(3).TEXT))
   'End If
End Sub

Sub GENERODES()
   '
   If CONTROL$("ADMIPED", "GENDESPA") = "REVISTOCK" Then
     Call GENODESPA(XVALO(Text11), NODES&, 1, OKKY%)
     If OKKY% = 1 Then
        Call PRIODESPA(NODES&)
     End If
   End If
   '
End Sub
'
Sub CALPERIB()
   '
   'SE MODIFICA PARA PASARLE EL VALOR DE VARIDOC POR QUE LO ESTABA PASANDO SIN VALOR
   NCLIEN = XVALO(Text1)
   PIVADOC% = PIVACLI%(NCLIEN)
   '
   'SE ASIGNA VALOR DE VARIDOC
   VARIDOC% = 1
   If PIVADOC% = 0 Or PIVADOC% = 3 Or PIVADOC% = 4 Or PIVADOC% = 6 Then
        VARIDOC% = 2
      ElseIf PIVADOC% = 5 Or PIVADOC% = 7 Then
        VARIDOC% = 3
        If LETREXPO$ = "A" Then
          VARIDOC% = 1
        End If
   End If
   '
   TASAX = ALIPERBRU#(NCLIEN, MINIPER#, VARIDOC%)
   '
   Text15 = ""
   BIMPON# = XVALO(Total(2))
   If BIMPON# * TICAMONE(MONEMI%) >= MINIPER# Then
      PERSALTA# = BIMPON# * (TASAX / 100)
      Text15 = FORMATNUM$(PERSALTA#, "F11.2")
   End If
   '
End Sub
'
Sub CALPERIB_VIE()
   '
   Static CTTXX%, TASAA, TASAB, TASAEXP, MINIPER
   '
   'If TIFADOC% > 2 Then
   '   Exit Sub
   'End If
   '
   If CTTXX% < 1 Then
     TASAA = XVALO(CONTROL$("PERIBRU", "TASA-A"))
     TASAB = XVALO(CONTROL$("PERIBRU", "TASA-B"))
     TASAEXP = XVALO(CONTROL$("PERIBRU", "TASA-EXP"))
     MINIPER = XVALO(CONTROL$("PERIBRU", "MINIPER"))
     CTTXX% = 1
   End If
   '
   TASAX = TASAA
   NCLIE% = XVALO(Text1)
   If NCLIE < 1 Then Exit Sub
   '
   VARIDC% = 1: PVCLX% = PIVACLI%(NCLIE%)
   If PVCLX% = 0 Or PVCLX% = 4 Then VARIDC% = 2
   If PVCLX% = 5 Or PVCLX% = 7 Then VARIDC% = 3
   '
   If VARIDC% = 2 Then TASAX = TASAB
   If VARIDC% = 3 Then TASAX = TASAEXP
   '
   Text15 = ""
   REPERCI$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE%))
   CATIB1% = Asc(Mid$(REPERCI$, 3, 1))
   PORCONVE = CVS(Mid$(REPERCI$, 5, 4))
   If PORCONVE > 100 Then PORCONVE = 100
   '
   If CATIB1% > 0 Then         ' NO ESTA EXENTO
     If PORCONVE >= 10.1 Then  ' PARTICIPACION CONVENIO MULTILATERAL
       BIMPON# = XVALO(Total(2))
       If VARIDC% = 2 Then BIMPON# = BIMPON# + XVALO(TOTIVA(3))
       If BIMPON# * TICAMONE(MONEMI%) >= MINIPER Then
         PERSALTA# = (BIMPON# * PORCONVE / 100) * (TASAX / 100)
         Text15 = FORMATNUM$(PERSALTA#, "F11.2")
       End If
     End If
   End If
   '
End Sub
'
Private Sub MODOCON(MODO%)
    Call CIERRARCH("*.*")
    MODX% = MODO%
    If MODX% = 1 Then
         Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
       ElseIf MODX% = 2 Then
         Call CARCONPAGOS
    End If
End Sub

Sub IMPORZET(ORIPED$)
   '
   Dim NUPEDI&(), CLIPEDI%(), FEPEDI%()
   Dim NPEDIT&(), NORIT%(), CODIT%()
   CAENCA% = 0
   CADETA% = 0
   '
   If EXISTE%(ORIPED$) < 1 Then
     MERRO$ = REPLACAR$(ORIPED$, "\", "/")
     Call MENSERR(24, "Imposible Importar.\No Existe la B.D. Referenciada\'" + MERRO$ + "'.")
     Exit Sub
   End If
   '
   FECORTE% = FECHANUM(CONTROL$("PEDIDO", "FECORTE"))
   '
   Dim OEMPAQ As Database
   Dim ENCAOEM As Recordset
   Dim DETAOEM As Recordset
   Dim PARCIOEM As Recordset
   '
   Set OEMPAQ = OpenDatabase(ORIPED$)
   Set ENCAOEM = OEMPAQ.OpenRecordset("ORDENESEMPAQUE", dbOpenDynaset)
   Set DETAOEM = OEMPAQ.OpenRecordset("DETALLESOE", dbOpenDynaset)
   Set PARCIOEM = OEMPAQ.OpenRecordset("OEPARCIALES", dbOpenDynaset)
   '
   On Error Resume Next
   ENCAOEM.MoveLast
   If Err Then
     On Error GoTo 0
     Call MENSERR(24, "Imposible Importar. Tabla 'OrdenesEmpaque' Vacía.")
     Exit Sub
   End If
   On Error GoTo 0
   FIN& = ENCAOEM.RecordCount
   '
   Screen.MousePointer = 11
   With ENCAOEM
      .MoveFirst
      KUK& = 0
      PRIPEDI& = 999999
      Do While Not .EOF
        KUK& = KUK& + 1
        Call TRACE(20, KUK&, FIN&)
        If CVINT(!FechaHoraEmisión) >= FECORTE% Then
           ' validar cliente
           If !CodClien < 1 Or !CodClien > 32767 Then
              MERRO$ = "OE " & !NumOE & " - Cliente Número " & !CodClien & " no Válido."
              Call ADDMERR(MERRO$)
              GoTo 12
           End If
           If STATCLI%(!CodClien) < 1 Then
              MERRO$ = "OE " & !NumOE & " - Cliente Número " & !CodClien & " Inexistente o Inactivo (de Baja)."
              Call ADDMERR(MERRO$)
           End If
           ' fin validar cliente
           '
12         CAENCA% = CAENCA% + 1
           ReDim Preserve NUPEDI&(CAENCA%), CLIPEDI%(CAENCA%), FEPEDI%(CAENCA%)
           NUPEDI&(CAENCA%) = !NumOE
             If !NumOE < PRIPEDI& Then
               PRIPEDI& = !NumOE
             End If
           CLIPEDI%(CAENCA%) = !CodClien
           FEPEDI%(CAENCA%) = CVINT(!FechaHoraEmisión)
        End If
      .MoveNext
      Loop
   End With
   '
   On Error Resume Next
   DETAOEM.MoveLast
   If Err Then
     On Error GoTo 0
     Call MENSERR(24, "Imposible Importar. Tabla 'DetallesOE' Vacía.")
     Exit Sub
   End If
   On Error GoTo 0
   FIN& = DETAOEM.RecordCount
   '
   With DETAOEM
      .MoveFirst
      KUK& = 0
      Do While Not .EOF
        KUK& = KUK& + 1
        Call TRACE(20, KUK&, FIN&)
        NUPED& = !NumOE
        If NUPED& >= PRIPEDI& Then
           NUIT% = !NumItem
           CODEX$ = !CodProd
           CIXI% = CODINT%(CODEX$)
           ' validar articulo
           If CIXI% < 1 Then
              MERRO$ = "Codigo '" & CODEX$ & "' Inexistente o no Válido."
              Call ADDMERR(MERRO$)
           End If
           ' fin validar articulo
           '
           CANTU = !CANTIDAD
           NCLIE% = 0: FPED% = 0
           For KKU& = 1 To CAENCA%
             If NUPEDI&(KKU&) = NUPED& Then
               NCLIE% = CLIPEDI%(KKU&)
               FPED% = FEPEDI%(KKU&)
               GoTo 15
             End If
           Next KKU&
           MERRO$ = "OE " & NUPED& & " Código '" & CODEX$ & "' Falta Encabezado de OE."
           Call ADDMERR(MERRO$)
           GoTo 19      ' NO ENCONTRO PEDIDO
           '
15         If NCLIE% > 0 Then
             CADETA% = CADETA% + 1
             ReDim Preserve NPEDIT&(CADETA%), NORIT%(CADETA%), CODIT%(CADETA%)
             NPEDIT&(CADETA%) = NUPED&
             NORIT%(CADETA%) = NUIT%
             CODIT%(CADETA%) = CIXI%
             Call AGREPED(NCLIE%, NUPED&, NUIT%, FPED%, CIXI%, CANTU)
           End If
        End If
        '
19    .MoveNext
      Loop
   End With
   '
   ' FIN DE IMPORTACION DE PEDIDOS
   On Error Resume Next
   PARCIOEM.MoveLast
   If Err Then
     On Error GoTo 0
     Call MENSERR(24, "Imposible Importar. Tabla 'OEParciales' Vacía.")
     Exit Sub
   End If
   On Error GoTo 0
   FIN& = PARCIOEM.RecordCount
   '
   With PARCIOEM
      .MoveFirst
      KUK& = 0
      Do While Not .EOF
        KUK& = KUK& + 1
        Call TRACE(20, KUK&, FIN&)
        NUPED& = !NumOE
        FENTREG% = CVINT(!FechaEntrega)
        CANTU = !CANTIDAD
        '
        If NUPED& >= PRIPEDI& Then
           NUIT% = !NumItem
           NORE% = !NumeroEntrega
           '
           CIXI% = 0
           For KKU& = 1 To CADETA%
             If NPEDIT&(KKU&) = NUPED& Then
               If NORIT%(KKU&) = NUIT% Then
                 CIXI% = CODIT%(KKU&)
                 GoTo 25
               End If
             End If
           Next KKU&
           MERRO$ = "Parcial OE " & NUPED& & " Item " & NUIT% & " Orden " & NORE% & " Falta Encabezado de OE."
           Call ADDMERR(MERRO$)
           GoTo 29      ' NO ENCONTRO PEDIDO
           '
25         If CIXI% > 0 Then
             Call AGREMOVST(NUPED&, NUIT%, NORE%, FENTREG%, CIXI%, CANTU)
           End If
        End If
        '
29    .MoveNext
      Loop
   End With
   '
   Call CIERRARCH("LOGIMPOR")
   Screen.MousePointer = 1
   '
   If ULTREG&("LOGIMPOR") < 1 Then
       Call COMUNI("Fin de Importación de Ordenes de Entrega.")
     Else
       OPX% = COMALTER%("Fin de Importación de OE's con Errores.\¿ Desea Consultar LOG-File Correspondiente ?\\Si, Consultar\No, Terminar")
       If OPX% = 1 Then
         Call FINAL("*LOGIMPOR")
       End If
   End If
   '
End Sub

Sub AGREPED(NCLIE, NUPED&, NUIT%, FPED%, CIXI%, CANTU)
   '
   Call ABREPEDCLI
   With PedEnca
     .FindFirst "NroPed = " & NUPED&
     If .NoMatch = False Then
        ' ENCONTRO ENCABEZADO DE PEDIDO
        If !NROCLIE <> NCLIE Then
          MERRO$ = "OE Número " + TRIM$(Str$(NUPED&)) + " no Coincide Cliente"
          Call ADDMERR(MERRO$)
        End If
        GoTo 20
     End If
     '
     STATUPC% = (-1)
     If Left$(CONTROL$("PEDCLI", "MODOAPRO"), 6) = "DIRECT" Then STATUPC% = 0
     '
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CODIEMPR = CodiEmp%
     On Error GoTo 0
     !NROPED = NUPED&
     '!Referencia = REFE$
     !FECHEMIS = CVDAT(FPED%)
     !NROCLIE = NCLIE
     !RaSoClie = RASOCLI$(NCLIE)
     !NROOCOM = "OE " + TRIM$(Str$(NUPED&))
     !DOMIENT = LUENTRECLI$(NCLIE, 0)
     !LOCAENT = LUENTRECLI$(NCLIE, 1)
     !TranspEnt = LUENTRECLI$(NCLIE, 2)
     !NTranspo = 0
     !NSucent = 0
     !Observa = ""
     !FESOLeNT = CVDAT(FPED%)
     !Status = STATUPC%              ' Status  de Aprobación
     !ListPre = LIPRCLI%(NCLIE)
     !CondPag = CPAGCLI%(NCLIE)
     !VENDED = VENDCLI%(NCLIE)
     !TDescApli = DESCUCLI$(NCLIE)
     .Update
   End With
   '
20 With PedDeta
      .FindFirst "NroPed = " & NUPED& & " AND NuOrItem = " & NUIT%
      If .NoMatch = False Then
         ' VERIFICAR QUE COINCIDAN DATOS
         If !NROCLIE = NCLIE Then
           If !FECHEMIS = CVDAT(FPED%) Then
             If !CODIINT = CIXI% Then
               If Abs(!CANPED - CANTU) < 0.05 Then
                 GoTo 99
               End If
             End If
           End If
         End If
         MERRO$ = "OE " & NUPED& & " Código '" & CODEX$ & "' Pre-Importada no Coinciden Datos."
         Call ADDMERR(MERRO$)
         GoTo 99
      End If
      '
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CODIEMPR = CodiEmp%
      On Error GoTo 0
         !NROPED = NUPED&
         !FECHEMIS = CVDAT(FPED%)
         !NROCLIE = NCLIE
         !RaSoClie = RASOCLI$(NCLIE)
         !NROOCOM = "OE " + TRIM$(Str$(NUPED&))
         !FESOLeNT = CVDAT(FPED%)
         !Status = STATUPE%   'idem anterior.
         !ListPre = LIPRCLI%(NCLIE)
         !CondPag = CPAGCLI%(NCLIE)
         !VENDED = VENDCLI%(NCLIE)
         !PorComVend = 0
         !NuorItem = NUIT%
         !CODIINT = CIXI%
         !CODIEXT = CODEXT$(CIXI%)
         !SubCod = ""
         !Descrip = DESCRIT0$(CIXI%)
         !TColor = ""
         !TTalle = ""
         !TLargo = 0
         !TAncho = 0
         !TEspesor = 0
         !VenRep = "V"
         !MoneEmis = 1
         !TipoCam = 1
         !PreList = 0 ' CVS(Mid$(X$, 39, 4))
         !CanAsig = 0
         !CANPED = CANTU
         !PREUNID = 0  ' CVS(Mid$(X$, 51, 4))
         !ImpoTot = 0  ' CVD(Mid$(X$, 55, 8))
         '   CAENTAN = CVS(Mid$(X$, 117, 4))   ' CANTIDAD ENTREGAS ANTERIORES
         '   FULENTAN% = CVI(Mid$(X$, 121, 2)) ' FECHA ULTIMA ENTREGA
         !CANTENT = 0 ' CAENTAN
         !CanSaldo = CANTU   ' CAN - CAENTAN
         !FeUltEnt = CVDAT(0)
         !NroRemi = ""
            CAVER$ = TRIM$(Str$(NUPED&))
            While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
            While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
            IK$ = TRIM$(Str$(NUIT%))
            CAVERI$ = CAVER$ + "-" + IK$
         !TPedItem = CAVERI$
         '
      .Update
   End With
   '
99 End Sub

Sub AGREMOVST(NUPED&, NUIT%, NORE%, FECHA%, CICONJ%, CANTU)
   '
   CMOV$ = "OE"
   IDMO$ = TRIM$(Str$(10 * NUIT% + NORE%))
   While Len(IDMO$) < 3: IDMO$ = "0" + IDMO$: Wend
   DOPRILA$ = "OE-" + TRIM$(Str$(NUPED&)) + "-" + IDMO$
   NUPED1& = NUPED&: While NUPED1& > 9999: NUPED1& = NUPED1& - 10000: Wend
   DOPRI$ = "OE" + TRIM$(Str$(NUPED1&)) + "-" + IDMO$
   DOSELA$ = DOPRILA$
   DOSEC$ = DOPRI$
   LOTE$ = DOPRILA$
   REFE$ = "Orden Empaque " + TRIM$(Str$(NUPED&))
   VUNI# = 0
   MONEII% = 1
   NCX% = 0
   DEPX% = DEPOEXPE%
   '
   Call ABRESTOCKS
   With Movisto
     .FindFirst "CodiMovi = '" & CMOV$ & "' AND DoPriLar = '" & DOPRILA$ & "'"
     If .NoMatch = False Then
        ' ENCONTRO MOVIMIENTO - FALTA VALIDAR
        If CVINT(!FechMov) = FECHA% Then
          If !COD_INTE = CICONJ% Then
            If Abs(!CantIngre - CANTU) < 0.05 Then
              GoTo 20
            End If
          End If
        End If
        MERRO$ = "Parciales OE " & NUPED& & " Código '" & CODEXT$(CICONJ%) & "' Pre-Importada no Coinciden Datos."
        Call ADDMERR(MERRO$)
        GoTo 20
     End If
     '
     Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, 0, VUNI#, MONEII%, NCX%, DEPX%, CANTU)
   End With
   '
20 End Sub

Sub ADDMERR(MERRO$)

   Static CTX1%
   If CTX1% < 1 Then
     Call SETULTREG("LOGIMPOR", 0)
     CTX1% = 1
   End If
   '
   Call REGAPP("LOGIMPOR", MERRO$)
   '
End Sub

'

Sub GENORFA_VIEJA()
   '
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     CIXI% = CVI(Left$(X$, 2))
     CODE$ = TRIM$(CODEXT$(CIXI%))
     NCLIE = XVALO(Text1)     'toma numero del cliente
     NUMPEDI% = XVALO(Text11) ' toma numero de pedido
     ORCOCLI$ = TRIM$(Text10)   'toma orden de compra del cliente
     If CONTROL$("ADMIPED", "NOVALSTK") <> "SI" Then
       If CAN > 0.05 Then
          CANSTO = STOCKACT(CIXI%)
          If CAN > CANSTO + 0.05 Then
             Call LEEEXPLO(CODE$, 1)
             If CAIT% > 0 Then
               TTXX$ = "Stock Insuficiente para\'" + CODE$ + " - " + TRIM$(DESCRIT0$(CIXI%)) + "'"
               TTXX$ = TTXX$ + "\   \¿ Desea Generar una Orden de Fabricación ?"
               If COMALTER%(TTXX$ + "\\Si, Generar O/F\No, Continuar") = 1 Then
                  ACONEC$ = "ORFABR07/ORFNUE" + Chr$(124) + CODE$ + Chr$(124) + TRIM$(Str$(CAN)) + Chr$(124) + TRIM$(Str$(NCLIE)) + Chr$(124) + TRIM$(Str$(NUMPEDI%)) + Chr$(124) + ORCOCLI$
                  Call CONECRUN(ACONEC$, "Ordenes de Fabricación")
               End If
             End If
          End If
       End If
     Else
       Call LEEEXPLO(CODE$, 1)
       If CAIT% > 0 Then
           TTXX$ = "¿ Desea Generar una Orden de Fabricación ?"
           TTXX$ = TTXX$ + "\Para el Articulo " & CODE$
           If COMALTER%(TTXX$ + "\\Si, Generar O/F\No, Continuar") = 1 Then
               ACONEC$ = "ORFABR07/ORFNUE" + Chr$(124) + CODE$ + Chr$(124) + TRIM$(Str$(CAN)) + Chr$(124) + TRIM$(Str$(NCLIE)) + Chr$(124) + TRIM$(Str$(NUMPEDI%)) + Chr$(124) + ORCOCLI$
               Call CONECRUN(ACONEC$, "Ordenes de Fabricación")
           End If
        End If
     End If
   Next U&
   '
End Sub

Sub GENORFA()
   '
   'SI LA CONFIGURACION DE PEDIDO ES QUE VALIDE POR STOCKS EN CASO QUE NO HAYA STOCKS
   'PRESENTA MENSAJE PARA PREGUNTAR SI SE DESE REALIZAR LA OF.
   'SI NO ESTA EST CONFIGURACION  Y ESTA CONFIGURADA  ORDEN DE FABRICACION DIRECTA
   'ENTRA EN EL ELSE Y NO PREGUNTA SI SE DESEA GENERAR LA OF Y LA REALIZA DIRECTAMENTE.
   'EN EL CASO DE JYM SE DEBERAN TILDAR QUE NO VALIDE STOCKS Y QUE SEA GENERACION AUTOMATICA
   GENERAORDEFA% = 0: ACONEC1$ = ""
TXTRAZA$ = ""
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     CIXI% = CVI(Left$(X$, 2))
     CODE$ = TRIM$(CODEXT$(CIXI%))
     NCLIE = XVALO(Text1)     'toma numero del cliente
     NUMPEDI% = XVALO(Text11) ' toma numero de pedido
     ORCOCLI$ = TRIM$(Text10)   'toma orden de compra del cliente
     ORCOCLI$ = REPLACAR$(ORCOCLI$, "/", "-")
     If CONTROL$("ADMIPED", "NOVALSTK") <> "SI" Then
       If CAN > 0.05 Then
          CANSTO = STOCKACT(CIXI%)
          If CAN > CANSTO + 0.05 Then
             Call LEEEXPLO(CODE$, 1)
             If CAIT% > 0 Then
               TTXX$ = "Stock Insuficiente para\'" + CODE$ + " - " + TRIM$(DESCRIT0$(CIXI%)) + "'"
               TTXX$ = TTXX$ + "\   \¿ Desea Generar una Orden de Fabricación ?"
               If COMALTER%(TTXX$ + "\\Si, Generar O/F\No, Continuar") = 1 Then
                  GENERAORDEFA% = 1
                  ACONEC1$ = ACONEC1$ + Chr$(124) + CODE$ + Chr$(124) + TRIM$(Str$(CAN)) + Chr$(124) + TRIM$(Str$(NCLIE)) + Chr$(124) + TRIM$(Str$(NUMPEDI%)) + Chr$(124) + ORCOCLI$ + "!"
               End If
             End If
          End If
       End If
     Else
       
       Call LEEEXPLO(CODE$, 1)
       If CAIT% > 0 Then
           If UCase$(CONTROL("ADMIPED", "OFDIRECT")) = "SI" Then GoTo 50
           TTXX$ = "¿ Desea Generar una Orden de Fabricación ?"
           TTXX$ = TTXX$ + "\Para el Articulo " & CODE$
           If COMALTER%(TTXX$ + "\\Si, Generar O/F\No, Continuar") = 1 Then
50             GENERAORDEFA% = 1
               ACONEC1$ = ACONEC1$ + Chr$(124) + CODE$ + Chr$(124) + TRIM$(Str$(CAN)) + Chr$(124) + TRIM$(Str$(NCLIE)) + Chr$(124) + TRIM$(Str$(NUMPEDI%)) + Chr$(124) + ORCOCLI$ + "!"
               TXTRAZA$ = TXTRAZA$ + " " & ACONEC1$ + vbCrLf
           End If
        End If
     End If
   Next U&
   '
   ACONEC1$ = REPLACAR$(ACONEC1$, "/", "ß") 'PARA EVITAR CODIGOS CON BARRA POR QUE EN  EL PASAJE LLEGA TRUNCADO POR ESTE CARACTER (JYM) (ALT-225)
   If GENERAORDEFA% > 0 Then
      ACONEC$ = TRIM$("ORFABR07/ORFNUE" + ACONEC1$)
      'QUITO EL ULTIMO PUNTO Y COMA
      
      If Right(ACONEC$, 1) = "!" Then ACONEC$ = Left$(ACONEC$, Len(ACONEC$) - 1)
      
      'TRAZA
         LX& = 1
1999     If EXISTE%(LUCOM$ + "TRPED" & SAFETEXT$(LX&) & ".TXT") = 1 Then
           LX& = LX& + 1
           GoTo 1999
         End If
         TXTRAZA$ = TXTRAZA$ + " ULTIMO " & ACONEC$
         NARCHI2% = FILEOPEN%(LUCOM$ + "TRPED" & SAFETEXT$(LX&) & ".TXT", 2, 256) ' ABRE EL ARCHIVO DE ESCRITURA
         Print #NARCHI2%, TXTRAZA$
         Close #NARCHI2%
      'FIN TRAZA
      Call CONECRUN(ACONEC$, "Ordenes de Fabricación")
   End If
   '
End Sub


Sub CONAPLI(APLINVO$, TERMINA%)
   '
   CODCARPE$ = BUSCAVALOR(PARAPLI$, "CODCARPE =") ' codigo de carpeta de import / export
   '
   TERMINA% = 1
   Select Case APLINVO$
       '
       'OPCIONES VARIAS DE PEDIDOS
       '
       Case "OPVARPED"
           Call Command13_Click             ' OPCIONES VARIAS - PEDIDOS
       Case "MODPEDIDO"
           Call OPPED07.Command46_Click     ' MODIFICACION DE PEDIDO EN CURSO
       Case "ANULAPED"
           Call OPPED07.Command45_Click     ' ANULACION DE PEDIDO EN CURSO
       Case "REVANUPED"
           Call OPPED07.Command13_Click     ' REVERTIR ANULACION DE PEDIDO
       Case "ANUITEMPED"
           Call OPPED07.Command4_Click      ' ANULA ITEM DE PEDIDO
       Case "APROBAPED"
           Call OPPED07.Command5_Click      ' APROBAR PEDIDOS DE CLIENTES
       Case "GENPEDCOPI"
           Call OPPED07.Command11_Click     ' GENERA PEDIDO POR COPIA
       Case "GENPEDABI"
           Call Image4_DblClick             ' GENERA PEDIDO ABIERTO
       '
       'CONSULTAS Y LISTADOS DE PEDIDOS
       '
       Case "CONLISPED"
           Call Command13_Click             ' CONSULTAS Y LISTADOS
       Case "ENCAPED"
           Call OPPED07.Command61_Click     ' ENCABEZA LISTADO DE PEDIDOS
       Case "DETNUMPED"
           OPPED07.Option2.Value = True
           Call OPPED07.Command63_Click     ' DETALLE POR NUMERO DE PEDIDO
       Case "PEDPENCLI"
           OPPED07.Option2.Value = True
           Call OPPED07.Command62_Click     ' PEDIDOS PENDIENTES POR CLIENTE
       Case "PEDPENART"
           OPPED07.Option2.Value = True
           Call OPPED07.Command64_Click     ' PEDIDOS PENDIETNES POR ARTICULO
       Case "ESTVENPED"
           Call OPPED07.Command12_Click     ' ESTADISTICA DE VENTAS Y PEDIDOS
       Case "HISPREPED"
           Call OPPED07.Command61_Click     ' HISTORICO DEPRECIOS POR PEDIDO
       Case "RANKPEDI"
           Call OPPED07.Command8_Click      ' RANKING DE PEDIDO
       Case "RASFACPED"
           Call OPPED07.Command8_Click      ' RASTREABILIDAD FACTURAS / PEDIDOS
       'CONFIGURACION DE PEDIDOS
       '
       Case "CONGENPED"
           Call Command5_Click              ' CONFIGURACION GENERAL DE PEDIDO
      
       Case Else
           TERMINA% = 0
   End Select
End Sub

Sub PEDIPEND()
    '
    NCLIEN = XVALO(Text1.TEXT)
    CONDICION$ = " NroClie =  " & NCLIEN & " AND Status < 8 AND CantEnt < (CanPed - 0.005) AND  Status <> 3 ORDER BY NroPed ASC"
    '
    Call COPYSTRU("INDIPEP", "INDIPNN")
    '
    Call CARGALISEL("INDIPNN", "PEDDETA", "NroPed/F5.0;RaSoClie/A20;FechEmis/D8", CONDICION$, "DISTINCT")
    '
    If ULTREG&("INDIPNN") < 1 Then Exit Sub
    '
10  Call SELECHO("INDIPNN")
    NPX& = XVALO(VALACT1("INDIPNN"))
    If NPX& < 1 Then Exit Sub
   '
    If NPX& > 0 Then
      TIDOCUM$ = "Pedido de Cliente"
      '
      NUDOCU$ = TRIM$(Str$(NPX&))

      While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
      NUDOCU$ = NUDOCU$ + " "
      Screen.MousePointer = 11
      '
      For UI& = ULTREG&("PDOCLST") To 1 Step -1
        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
        If InStr(XX$, TIDOCUM$) > 4 Then
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
      Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
      GoTo 10
5     Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    End If
    GoTo 10
    
End Sub


Function PORCRECAR(CPAGO) 'CALCULA PORCENTAJE DE RECARGO SEGUN MEDIO DE PAGO
   '
   PORECA = 0
   For JJ& = 1 To ULTREG&("CONPAG")
     XXT$ = REGLEIDO$("CONPAG", JJ&)
     COPA% = Asc(Left$(XXT$, 1))
     If COPA% = XVALO(CPAGO) Then
        PORECA = CVS(Mid$(XXT$, 44, 4))
        PORCRECAR = PORECA
        Exit Function
     End If
   Next JJ&
   '
End Function

Sub CARGAREPARA()

   Screen.MousePointer = 11
   Dim RS As ADODB.Recordset
   NCLIE = XVALO(Text1)
   
   Call CREATABLA_OFABRIC

   SQLX$ = "SELECT * FROM REPARA WHERE (FECHPRES <> '' OR FECHPRES IS NOT NULL)"
   SQLX$ = SQLX$ + " AND NUME_CLI=" & NCLIE
   SQLX$ = SQLX$ + " AND NROPED < 0.5"
   
   Set RS = FLEXCONN.Execute(SQLX$)
   With RS
        On Error Resume Next
        .MoveFirst
        On Error GoTo 0
        Do While .EOF = False
            NROREP& = XVALO(!NUMREPA)
            REFE$ = SAFETEXT$(!Referen)
            TTXX$ = "Reparación No: " + TRIM$(FORMATNUM$(NROREP&, "F8.0"))
            TTXX$ = TTXX$ + " - " + Mid$(REFE$, 1, 60)
            TTXX$ = AJUSTI$(TTXX$, 128)
            Call REPLA(TTXX$, Str$(NROREP&), 121, 8)
            List2.AddItem TTXX$
            .MoveNext
        Loop
   End With
   
   On Error Resume Next
   RS.Close
   Set RS = Nothing
   On Error GoTo 0
   
   Screen.MousePointer = 1
End Sub

Sub SELECREPARA()

   Screen.MousePointer = 11
     
   Call CREATABLA_OFABRIC
   
   HOII% = HOY(HOS$)
   TEXT21(0) = "0"
   JJ& = 0: ATX% = 0
   For KKI% = 1 To List2.ListCount
     If List2.Selected(KKI% - 1) = True Then
       REPRESU& = XVALO(Mid$(List2.LIST(KKI% - 1), 121, 8))
       If REPRESU& > 0 Then
            CONDI$ = "NUMREPA=" & REPRESU&
            CODI$ = CONTROL$("REPARA", "CODGENER")
            CIXI% = CODINT%(CODI$)
            Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
            PRELI = XVALO(VALOBADA("REPARA", "Imp_Tot", CONDI$))
            CPPP = 1  ' (100 - COEFDESCU(DESCUTEX$)) / 100
            '
            MONEMA% = 1   ' REPARACIONES SIEMPRE EN PESOS
            MONEMI% = 1
            MONEMIS = DEMONECO$(MONEMI%)
            COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
            '
            Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 39, 4)
            Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 51, 4)
            Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 65, 4)
            Call REPLA(ZZ$, MKS$(1), 47, 4)   ' CANTIDAD
            Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 73, 4)
            Call REPLA(ZZ$, MKI$(HOII%), 63, 2)
            Call REPLA(ZZ$, MKS$((-1) * REPRESU&), 125, 4)
            JJ& = JJ& + 1
            Call GRAREG("!CARDETPE", ZZ$, JJ&)
       End If
     End If
   Next KKI%
   Call SETULTREG("!CARDETPE", JJ&)
   Call CIERRARCH("!CARDETPE")

   Screen.MousePointer = 11
End Sub

Sub GENERAR_DAT(n%)
   '
   If INCLUALIAS% = 0 Then
     INCLUALIAS% = (-1)
     If CONTROL$("", "IDSECUN") = "ALIAS" Then INCLUALIAS% = 1
   End If
   '
   Screen.MousePointer = 11
   ARCHIPRE$ = TRIM$(Str$(n%))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   CADECI% = XVALO(VALOBADA("LISTAPRE", "DECIMALES", "cod_lista=" & n%))
   If CADECI% > 4 Then CADECI% = 4
   '
   Call COPYSTRU("LIPREVEN", ARCHIPRE$)
   LF2% = LENFIELD(ARCHIPRE$, 2)
   '
   Call APERBASDAT("PEDCLI")
   CONDI$ = "Cod_Lista = " & n% & " "
   SQLX$ = "SELECT * FROM Preventa WITH(NOLOCK)"
   SQLX$ = SQLX$ + "WHERE " & CONDI$ & " "
   SQLX$ = SQLX$ + "ORDER BY Num_Orden;"
   'Set ARTSELEC = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim ARTSELEC As ADODB.Recordset
   Set ARTSELEC = New ADODB.Recordset
   Set ARTSELEC = FLEXCONN.Execute(FILTSQL$(SQLX$))

   FIN& = CantRegistros("PREVENTA", CONDI$)
   U& = 0
   On Error Resume Next
   With ARTSELEC
     If FIN& > 0 Then
        Do While (.EOF = False)
            U& = U& + 1
            Call TRACE(20, U&, FIN&)
            TTXX$ = REGBLAN$(ARCHIPRE$)
            CII% = XVALO(!Cod_Pieza)
            Call REPLA(TTXX$, MKI$(CII%), 1, 2)
            Call REPLA(TTXX$, DESCRIT$(CII%), 3, LF2%)
               MRX$ = SAFETEXT$(!Carac_Control)
            Call REPLA(TTXX$, MRX$, 3 + LF2%, 1)
            Call REPLA(TTXX$, !Num_Original, 4 + LF2%, 16)
               PRUNI# = XVALO(!Precio_Unitario)
            Call REPLA(TTXX$, MKD$(PRUNI#), 20 + LF2%, 8)
            Call REPLA(TTXX$, MKS$(!Cant_Limite), 28 + LF2%, 2)
            Call REPLA(TTXX$, MKD$(!Porc_Limite), 30 + LF2%, 1)
            '
            PRELI# = PRUNI#
            CDC$ = TRIM$(Str$(CADECI%))
            PREX$ = " " + FORMATNUM$(PRELI#, "F12." + CDC$) + "  "
            Call REPLA(TTXX$, PREX$, 1 + LF2% - Len(PREX$), 2 + Len(PREX$))
            Call REPLA(TTXX$, MRX$, 2 + LF2%, 1)
            '
            If INCLUALIAS% > 0 Then
              ALX$ = TRIM$(NUPLANO$(CII%))
              If ALX$ <> "" Then
                Call REPLA(TTXX$, " (" + ALX$ + ")", LF2% - 16 - Len(ALX$), 3 + Len(ALX$))
              End If
            End If
            JJ& = JJ& + 1
            Call GRAREG(ARCHIPRE$, TTXX$, JJ&)
            '
        .MoveNext
        Loop
        Call SETULTREG(ARCHIPRE$, JJ&)
        Call CIERRARCH(ARCHIPRE$)
        Call FRESHECHO(ARCHIPRE$)
      End If
   End With
   Set ARTSELEC = Nothing
   Screen.MousePointer = 1
   '
End Sub




Private Sub TXTFLETE_DblClick()
   TXTFLETE = ""
End Sub

Private Sub Verocomsuc_Click()
    Call CREAR_TABLA_OC_DE_SUCURSAL
    '
10  OPX% = COMALTER%("Pedido de Reposición de Sucursal\\Todas\Solo Pendientes de Generar Nota de Venta\Solo con Nota de Venta Generada")
    If OPX% < 1 Then Exit Sub
    '
    ENCA$ = "Sucursal/F10;R.Social/A32;P.Repos./F9;Emis.P.R./D8;NC.Pedido/F12;R.Social Cliente de Sucursal/A32;P.Avellaneda/F12.0;O.C.Cte.Suc./A12"
    Call CARGA_ENCABEZADO(OCOMSUCU.GRID, ENCA$, OCOMSUCU)
    Call CARGA_ENCABEZADO(OCOMSUCU.MSF_2, ENCA$, OCOMSUCU)
    ENCA$ = "Código/F16;Descripción/A32;Can.Compra/F11.1;Sol.Ent./D8;Ped.Avell../F9;Cant.Ped./F9.1;Cant.Ent./F9.1;Saldo/F9.1"
    Call CARGA_ENCABEZADO(OCOMSUCU.MSFDET, ENCA$, OCOMSUCU)
    ENCA$ = "O.Compra/F9;Mensaje/A112;NROPED/F0"
    Call CARGA_ENCABEZADO(OCOMSUCU.MSFMEN, ENCA$, OCOMSUCU)
    OCOMSUCU.Caption = "Lista de Pedidos de Reposición de Sucursales"
    '
    OCOMSUCU.GRID.Rows = 1
    SQLX$ = "SELECT SUCURSAL,O_COMPRA_SUC, NCLIEN_PEDIDO,VERSION_COMPRA_SUC, NPED_GENERADO_EN_AVELLANEDA,FECHEMISI, OComCliente,RAZ_CTE_SUC FROM OC_SUCURSAL WITH(NOLOCK)"
    SQLX$ = SQLX$ + "GROUP BY SUCURSAL,O_COMPRA_SUC, NCLIEN_PEDIDO,VERSION_COMPRA_SUC,NPED_GENERADO_EN_AVELLANEDA,FECHEMISI,OComCliente,RAZ_CTE_SUC"
    SQLX$ = SQLX$ + " ORDER BY VERSION_COMPRA_SUC,O_COMPRA_SUC DESC" 'SUCURSAL,O_COMPRA_SUC, NCLIEN_PEDIDO "
    Set RS = READSET(SQLX$)
    I& = 0
    With RS
      Do While Not .EOF
        P_Avellaneda& = XVALO(!NPED_GENERADO_EN_AVELLANEDA)
        If OPX% = 1 Then GoTo 30 'SI ES TODOS MUESTRA SIEMPRE
        If P_Avellaneda& <= 0 And OPX% = 2 Then GoTo 30 'SI NO TIENE PEDIDO Y SE ELIGIO ESA OPCION INGRESA
        If P_Avellaneda& > 0 And OPX% = 3 Then GoTo 30  'SI TIENE PEDIDO Y SE ELIGIO ESA OPCION ENTRA
        GoTo 50
        '
30      I& = I& + 1
        OCOMSUCU.GRID.Rows = I& + 1
        SUCUR = XVALO(!SUCURSAL)
        OCOMSUCU.GRID.TextMatrix(I&, 1) = SUCUR
        OCOMSUCU.GRID.TextMatrix(I&, 2) = RASOCLI$(SUCUR)
           OCOMPRAX& = XVALO(!O_COMPRA_SUC)
           P_Avellaneda& = XVALO(!NPED_GENERADO_EN_AVELLANEDA)
        OCOMSUCU.GRID.TextMatrix(I&, 3) = OCOMPRAX&
'        FECHA11 = FECHATEX$(CVINT(!FECHEMISI))
        OCOMSUCU.GRID.TextMatrix(I&, 4) = FECHATEX$(CVINT(!FechEmisi))
        OCOMSUCU.GRID.TextMatrix(I&, 5) = XVALO(!NCLIEN_PEDIDO)
        OCOMSUCU.GRID.TextMatrix(I&, 6) = SAFETEXT$(!RAZ_CTE_SUC)
        OCOMSUCU.GRID.TextMatrix(I&, 7) = P_Avellaneda&
        OCOMSUCU.GRID.TextMatrix(I&, 8) = SAFETEXT$(!OComCliente)
        '
50      .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    '
    'CARGA LA LISTA DE MENSAJES.
'    SQLX$ = "SELECT * FROM MENSA_SUCURSAL WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " WHERE STATUS = 0 "
'    Set RS = READSET(SQLX$)
'    With RS
'      Do While Not .EOF
'        OCOMPRAX& = XVALO(!O_COMPRA_SUC)
'        USUARIOX% = XVALO(!USUARIO)
'        MENSAX$ = SAFETEXT$(!MENSAJE)
'        j& = j& + 1
'        OCOMSUCU.MSFMEN.Rows = j& + 1
'        If TRIM$(MENSAX$) <> "" Then
'           OCOMSUCU.Label75 = MENSAX$
'           OCOMSUCU.Label75.Height = 100
'           OCOMSUCU.MSFMEN.RowHeight(j&) = OCOMSUCU.Label75.Height + 100
'           OCOMSUCU.MSFMEN.TextMatrix(j&, 2) = MENSAX$
'        End If
'        .MoveNext
'      Loop
'    End With
    OCOMSUCU.GRID.COL = 1
    OCOMSUCU.GRID.Sort = 3

    OCOMSUCU.GRID.COL = 3
    OCOMSUCU.GRID.Sort = 3

    If I& + j& > 0 Then
       OCOMSUCU.Show 1
       GoTo 10
    Else
       Call COMUNI("No Se Detectaron Pedidos de Reposición Segun Filtro Seleccionado")
    End If
End Sub
