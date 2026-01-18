VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FORPED07 
   BackColor       =   &H00624E28&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos de Clientes - Formato Lineal"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   12150
   Icon            =   "FORPED07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   12150
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
      Height          =   690
      Left            =   11400
      Picture         =   "FORPED07.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   154
      ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
      Top             =   7440
      Width           =   650
   End
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   11280
      Sorted          =   -1  'True
      TabIndex        =   129
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   11290
      ScaleHeight     =   6705
      ScaleWidth      =   975
      TabIndex        =   107
      Top             =   0
      Width           =   1000
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   157
         Top             =   6000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   158
            Top             =   0
            Width           =   12000
         End
      End
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
         Picture         =   "FORPED07.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   116
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
         Picture         =   "FORPED07.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   117
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
         Left            =   120
         Picture         =   "FORPED07.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   120
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
         Picture         =   "FORPED07.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   2760
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
         Picture         =   "FORPED07.frx":11B4
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Configuración de Funcionamiento de Pedidos"
         Top             =   4650
         Visible         =   0   'False
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
         Picture         =   "FORPED07.frx":15F6
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Acceso a Condiciones Comerciales"
         Top             =   5280
         Visible         =   0   'False
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
         Picture         =   "FORPED07.frx":1900
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Acceso a Maestro de Artículos"
         Top             =   2000
         Visible         =   0   'False
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
         Picture         =   "FORPED07.frx":1C0A
         Style           =   1  'Graphical
         TabIndex        =   110
         ToolTipText     =   "Acceso a Base de Datos de Clientes"
         Top             =   1420
         Visible         =   0   'False
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
         Picture         =   "FORPED07.frx":204C
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Opciones Varias - Modificaciones y Consultas"
         Top             =   3960
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
         Picture         =   "FORPED07.frx":2196
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   685
         Visible         =   0   'False
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2D3C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8000
      ScaleHeight     =   330
      ScaleWidth      =   3285
      TabIndex        =   94
      Top             =   550
      Visible         =   0   'False
      Width           =   3280
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
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
         Left            =   840
         MaxLength       =   2
         TabIndex        =   96
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
         TabIndex        =   95
         Top             =   0
         Width           =   180
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
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
         Height          =   225
         Left            =   0
         TabIndex        =   98
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   1330
         TabIndex        =   97
         Top             =   0
         Width           =   1895
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2D3C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9450
      ScaleHeight     =   330
      ScaleWidth      =   1785
      TabIndex        =   92
      Top             =   520
      Width           =   1785
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Venta"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         TabIndex        =   93
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      BorderStyle     =   0  'None
      Height          =   3540
      Left            =   50
      TabIndex        =   28
      Top             =   4800
      Width           =   11175
      Begin VB.CheckBox CheckAhpRetiraDep 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Retira Nuestro Depósito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   175
         Top             =   2520
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Modo Test AHP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   174
         Top             =   0
         Visible         =   0   'False
         Width           =   2775
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
         Height          =   1560
         Left            =   5630
         TabIndex        =   42
         ToolTipText     =   "Habilita Descuento Retira Cliente"
         Top             =   120
         Width           =   5550
         Begin VB.CommandButton Command29 
            Caption         =   "más datos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text141 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   3885
            TabIndex        =   123
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   1080
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Dt.Rt.Cte."
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
            Left            =   2040
            TabIndex        =   122
            Top             =   1200
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   720
            Width           =   1455
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E2D3C0&
            Caption         =   "IVA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   105
            TabIndex        =   74
            Top             =   2205
            Width           =   1830
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
               Height          =   255
               Left            =   720
               OleObjectBlob   =   "FORPED07.frx":24A0
               TabIndex        =   168
               Top             =   160
               Width           =   1005
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
               Height          =   255
               Left            =   720
               OleObjectBlob   =   "FORPED07.frx":24F6
               TabIndex        =   169
               Top             =   365
               Width           =   1005
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
               Height          =   255
               Left            =   720
               OleObjectBlob   =   "FORPED07.frx":254C
               TabIndex        =   170
               Top             =   580
               Width           =   1005
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
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
               Left            =   2000
               TabIndex        =   167
               Top             =   200
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
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
               Left            =   2000
               TabIndex        =   79
               Top             =   620
               Visible         =   0   'False
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
               TabIndex        =   78
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
               TabIndex        =   77
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
               TabIndex        =   76
               Top             =   405
               Width           =   400
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
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
               Left            =   2000
               TabIndex        =   75
               Top             =   405
               Visible         =   0   'False
               Width           =   1005
            End
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
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
            Left            =   3880
            TabIndex        =   72
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   1080
            Visible         =   0   'False
            Width           =   1500
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
            Height          =   330
            Left            =   3885
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   2250
            Width           =   1500
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
            Height          =   315
            Index           =   3
            Left            =   3885
            TabIndex        =   65
            Text            =   " "
            Top             =   2625
            Width           =   1500
         End
         Begin VB.TextBox TOTIVA 
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
            Height          =   330
            Index           =   3
            Left            =   3885
            Locked          =   -1  'True
            TabIndex        =   64
            Text            =   " "
            Top             =   1870
            Width           =   1500
         End
         Begin VB.TextBox TICAMBIS 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   840
            MaxLength       =   14
            TabIndex        =   56
            Top             =   1145
            Width           =   1065
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
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
            Left            =   3885
            TabIndex        =   54
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   680
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
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
            Index           =   0
            Left            =   3885
            TabIndex        =   53
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   315
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Total 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   3885
            TabIndex        =   45
            Text            =   " "
            Top             =   315
            Width           =   1500
         End
         Begin VB.TextBox Total 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   3885
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   " "
            Top             =   680
            Width           =   1500
         End
         Begin VB.TextBox PORDESCU 
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
            Left            =   2520
            MaxLength       =   14
            TabIndex        =   20
            Top             =   680
            Width           =   1065
         End
         Begin VB.TextBox Total 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   3885
            TabIndex        =   43
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1080
            Width           =   1500
         End
         Begin VB.Label Label10 
            Caption         =   "Label10"
            Height          =   255
            Left            =   1800
            TabIndex        =   160
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Left            =   3240
            TabIndex        =   124
            Top             =   1080
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Ltota 
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
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
            Index           =   8
            Left            =   120
            TabIndex        =   119
            Top             =   510
            Width           =   1065
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   540
            Left            =   120
            Picture         =   "FORPED07.frx":25A2
            Top             =   240
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Impuestos (IVA)"
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
            Left            =   1845
            TabIndex        =   69
            Top             =   2010
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Percep. IIBB"
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
            Index           =   6
            Left            =   2265
            TabIndex        =   68
            Top             =   2390
            Width           =   1485
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto Pedido"
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
            Left            =   1845
            TabIndex        =   67
            Top             =   2730
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento"
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
            Index           =   4
            Left            =   1635
            TabIndex        =   59
            Top             =   480
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
            TabIndex        =   58
            ToolTipText     =   "Doble-Click para 'Dólares'"
            Top             =   720
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de cambio"
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
            Left            =   120
            TabIndex        =   57
            Top             =   1050
            Width           =   795
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Neto"
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
            Index           =   0
            Left            =   2730
            TabIndex        =   48
            Top             =   390
            Width           =   1020
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Pedido"
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
            Index           =   3
            Left            =   1785
            TabIndex        =   47
            Top             =   1210
            Width           =   1500
         End
         Begin VB.Label Label7 
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
            Height          =   225
            Left            =   3630
            TabIndex        =   46
            Top             =   780
            Width           =   225
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   1590
         Index           =   0
         Left            =   6840
         TabIndex        =   60
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
            TabIndex        =   63
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
            TabIndex        =   62
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
            TabIndex        =   61
            Text            =   " "
            Top             =   840
            Visible         =   0   'False
            Width           =   870
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
         Height          =   3240
         Left            =   0
         TabIndex        =   34
         Top             =   120
         Width           =   5330
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
            Index           =   4
            Left            =   735
            TabIndex        =   173
            Text            =   " "
            Top             =   0
            Visible         =   0   'False
            Width           =   3270
         End
         Begin VB.TextBox TEXT21 
            Alignment       =   1  'Right Justify
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
            Index           =   4
            Left            =   0
            TabIndex        =   172
            Text            =   " "
            Top             =   0
            Visible         =   0   'False
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
            Height          =   315
            Index           =   4
            Left            =   420
            TabIndex        =   171
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.CommandButton cmdtext 
            Caption         =   "Test"
            Height          =   255
            Left            =   3120
            TabIndex        =   162
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Modo Test"
            Height          =   195
            Left            =   4080
            TabIndex        =   161
            Top             =   135
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   1155
            TabIndex        =   132
            Top             =   1920
            Width           =   615
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00E2D3C0&
               Caption         =   "Option1"
               ForeColor       =   &H80000008&
               Height          =   200
               Left            =   60
               TabIndex        =   134
               Top             =   160
               Width           =   255
            End
            Begin VB.OptionButton Option2 
               Appearance      =   0  'Flat
               BackColor       =   &H00E2D3C0&
               Caption         =   "Option2"
               ForeColor       =   &H80000008&
               Height          =   200
               Left            =   330
               TabIndex        =   133
               Top             =   160
               Width           =   255
            End
         End
         Begin VB.TextBox TXTFLETE 
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
            Left            =   1890
            TabIndex        =   130
            Text            =   " "
            ToolTipText     =   "Doble Click Para Borrar"
            Top             =   1980
            Width           =   3270
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
            Height          =   315
            Index           =   3
            Left            =   1575
            TabIndex        =   104
            Top             =   1620
            Width           =   180
         End
         Begin VB.TextBox TEXT21 
            Alignment       =   1  'Right Justify
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
            Index           =   3
            Left            =   1155
            TabIndex        =   103
            Text            =   " "
            Top             =   1620
            Width           =   435
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
            Index           =   3
            Left            =   1890
            TabIndex        =   102
            Text            =   " "
            Top             =   1620
            Width           =   3270
         End
         Begin VB.TextBox Text44 
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
            Left            =   4620
            MaxLength       =   4
            TabIndex        =   52
            Top             =   1230
            Width           =   540
         End
         Begin VB.TextBox Text13 
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
            Height          =   750
            Left            =   1155
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   2400
            Width           =   4005
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
            Index           =   2
            Left            =   1890
            TabIndex        =   37
            Text            =   " "
            Top             =   1230
            Width           =   1800
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
            Index           =   2
            Left            =   1155
            TabIndex        =   17
            Text            =   " "
            Top             =   1230
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
            TabIndex        =   18
            Top             =   1230
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
            Left            =   1890
            TabIndex        =   36
            Text            =   " "
            Top             =   825
            Width           =   2085
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
            Left            =   1155
            TabIndex        =   15
            Text            =   " "
            Top             =   840
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
            Index           =   1
            Left            =   1575
            TabIndex        =   16
            Top             =   825
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
            Left            =   1890
            TabIndex        =   35
            Text            =   " "
            Top             =   420
            Width           =   3270
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
            Left            =   1155
            TabIndex        =   13
            Text            =   " "
            Top             =   420
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
            Index           =   0
            Left            =   1575
            TabIndex        =   14
            Top             =   420
            Width           =   180
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Flete"
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
            Index           =   7
            Left            =   75
            TabIndex        =   131
            Top             =   2115
            Width           =   1050
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Ingreso"
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
            Index           =   6
            Left            =   75
            TabIndex        =   105
            Top             =   1725
            Width           =   1050
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Left            =   4650
            TabIndex        =   91
            Top             =   825
            Width           =   510
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dto. Rec"
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
            Index           =   5
            Left            =   3975
            TabIndex        =   90
            Top             =   945
            Width           =   645
         End
         Begin VB.Label Label17 
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
            Height          =   465
            Left            =   120
            TabIndex        =   71
            Top             =   2760
            Width           =   645
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   585
            TabIndex        =   70
            Top             =   2880
            Width           =   540
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "% Comisión"
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
            Index           =   4
            Left            =   3720
            TabIndex        =   51
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Observa- ciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   3
            Left            =   315
            TabIndex        =   41
            Top             =   2400
            Width           =   810
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor"
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
            Index           =   2
            Left            =   75
            TabIndex        =   40
            Top             =   1320
            Width           =   1050
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cond. de Pago"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   75
            TabIndex        =   39
            Top             =   945
            Width           =   1050
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
            Height          =   465
            Index           =   0
            Left            =   75
            TabIndex        =   38
            Top             =   315
            Width           =   1050
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   1560
         Left            =   5640
         TabIndex        =   31
         Top             =   1770
         Width           =   5550
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
            Left            =   1800
            TabIndex        =   86
            Top             =   1200
            Width           =   195
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
            Left            =   2040
            MaxLength       =   48
            TabIndex        =   85
            Top             =   1200
            Width           =   3345
         End
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
            TabIndex        =   84
            Top             =   1200
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
            Left            =   1800
            TabIndex        =   81
            Top             =   840
            Width           =   195
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
            TabIndex        =   80
            Top             =   870
            Width           =   645
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
            TabIndex        =   22
            Top             =   555
            Width           =   4230
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
            TabIndex        =   21
            Top             =   240
            Width           =   4230
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
            TabIndex        =   50
            Top             =   555
            Visible         =   0   'False
            Width           =   3750
         End
         Begin VB.Label Label24 
            Height          =   225
            Left            =   105
            TabIndex        =   88
            Top             =   840
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label23 
            Height          =   225
            Left            =   105
            TabIndex        =   87
            Top             =   525
            Visible         =   0   'False
            Width           =   960
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
            Left            =   150
            TabIndex        =   83
            Top             =   975
            Width           =   880
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
            Left            =   2040
            TabIndex        =   82
            Top             =   870
            Width           =   3345
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
            Index           =   5
            Left            =   30
            TabIndex        =   33
            Top             =   1305
            Width           =   1005
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
            Index           =   2
            Left            =   150
            TabIndex        =   32
            Top             =   360
            Width           =   915
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      BorderStyle     =   0  'None
      Height          =   8535
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton Command31 
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
         Left            =   0
         Picture         =   "FORPED07.frx":28AC
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Continúa Generación del Pedido de Cliente"
         Top             =   0
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.PictureBox Picture999 
         Height          =   540
         Index           =   0
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   1110
         TabIndex        =   165
         Top             =   0
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton Command30 
         Caption         =   ".."
         Height          =   300
         Left            =   0
         TabIndex        =   164
         ToolTipText     =   "CLICK - Para Editar Observación por Item"
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox MARSELPRESU 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   3060
         Left            =   45
         ScaleHeight     =   3030
         ScaleWidth      =   11145
         TabIndex        =   145
         Top             =   1680
         Visible         =   0   'False
         Width           =   11175
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C4F9FD&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2325
            Left            =   0
            Style           =   1  'Checkbox
            TabIndex        =   151
            Top             =   480
            Width           =   10500
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   470
            Left            =   -20
            ScaleHeight     =   435
            ScaleWidth      =   10500
            TabIndex        =   149
            Top             =   -30
            Width           =   10530
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "SELECCIÓN DE PRESUPUESTO EN CURSO"
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
               Height          =   345
               Left            =   0
               TabIndex        =   150
               Top             =   120
               Width           =   10425
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
            Left            =   10545
            Picture         =   "FORPED07.frx":2BB6
            Style           =   1  'Graphical
            TabIndex        =   148
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
            Left            =   10545
            Picture         =   "FORPED07.frx":2EC0
            Style           =   1  'Graphical
            TabIndex        =   147
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
            Left            =   10545
            Picture         =   "FORPED07.frx":31CA
            Style           =   1  'Graphical
            TabIndex        =   146
            ToolTipText     =   "Imprime Ficha de Anotaciones"
            Top             =   1890
            Width           =   570
         End
      End
      Begin VB.PictureBox MARCONOTA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   3060
         Left            =   50
         ScaleHeight     =   3030
         ScaleWidth      =   11145
         TabIndex        =   138
         Top             =   1680
         Visible         =   0   'False
         Width           =   11175
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   -30
            ScaleHeight     =   405
            ScaleWidth      =   10455
            TabIndex        =   143
            Top             =   -30
            Width           =   10480
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
               TabIndex        =   144
               Top             =   105
               Width           =   10425
            End
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C4F9FD&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2640
            Left            =   -30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   142
            Top             =   370
            Width           =   10485
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
            Left            =   10545
            Picture         =   "FORPED07.frx":3834
            Style           =   1  'Graphical
            TabIndex        =   141
            ToolTipText     =   "Continúa Generación del Pedido de Cliente"
            Top             =   420
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
            Left            =   10545
            Picture         =   "FORPED07.frx":3B3E
            Style           =   1  'Graphical
            TabIndex        =   140
            ToolTipText     =   "Cancela Pedido de Cliente"
            Top             =   1050
            Width           =   570
         End
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
            Left            =   10545
            Picture         =   "FORPED07.frx":3E48
            Style           =   1  'Graphical
            TabIndex        =   139
            ToolTipText     =   "Imprime Ficha de Anotaciones"
            Top             =   2205
            Width           =   570
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   3060
         Left            =   50
         TabIndex        =   137
         Top             =   1680
         Width           =   11170
         _ExtentX        =   19711
         _ExtentY        =   5398
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   210
         ScaleHeight     =   375
         ScaleWidth      =   10935
         TabIndex        =   135
         Top             =   1680
         Width           =   11000
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
            TabIndex        =   136
            Top             =   0
            Width           =   11000
         End
      End
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
         TabIndex        =   126
         Top             =   1280
         Width           =   160
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
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
         Left            =   6360
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   1280
         Width           =   975
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Command26"
         Height          =   495
         Left            =   7200
         TabIndex        =   101
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
         Left            =   1200
         TabIndex        =   89
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
         Picture         =   "FORPED07.frx":44B2
         Style           =   1  'Graphical
         TabIndex        =   55
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
         Height          =   2310
         Left            =   240
         TabIndex        =   49
         Top             =   2100
         Width           =   10950
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
         TabIndex        =   30
         Top             =   6855
         Width           =   2825
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
         Left            =   9615
         TabIndex        =   11
         Text            =   " "
         Top             =   920
         Width           =   1120
      End
      Begin VB.TextBox Text10 
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
         Left            =   9300
         MaxLength       =   16
         TabIndex        =   12
         Top             =   1280
         Width           =   1920
      End
      Begin VB.TextBox Text9 
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
         Left            =   10095
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
         Left            =   11040
         TabIndex        =   10
         Top             =   180
         Width           =   180
      End
      Begin VB.TextBox Text7 
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
         Left            =   5775
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   330
         Width           =   1695
      End
      Begin VB.TextBox Text8 
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
         Left            =   5760
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   900
         Width           =   1695
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
         Height          =   300
         Left            =   2310
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   920
         Width           =   2595
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
         Height          =   300
         Left            =   1470
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   920
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
         Height          =   300
         Left            =   1470
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   550
         Width           =   4065
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
         Left            =   960
         TabIndex        =   1
         Top             =   890
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
         Height          =   300
         Left            =   120
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   890
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid GRID2 
         Height          =   765
         Left            =   4800
         TabIndex        =   121
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
      Begin VB.TextBox Text28 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   50
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   163
         Top             =   4080
         Visible         =   0   'False
         Width           =   11100
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10880
         TabIndex        =   166
         ToolTipText     =   "Número de Revisión"
         Top             =   920
         Width           =   315
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pedidos de Clientes"
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
         Left            =   0
         TabIndex        =   159
         Top             =   100
         Width           =   1500
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
         Height          =   225
         Left            =   9240
         TabIndex        =   155
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Entrega"
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
         Height          =   390
         Left            =   5640
         TabIndex        =   127
         Top             =   1230
         Width           =   735
      End
      Begin VB.Image Image4 
         Height          =   450
         Left            =   7980
         Picture         =   "FORPED07.frx":5818
         Stretch         =   -1  'True
         ToolTipText     =   "Doble Click para Acceso a Pedidos Abiertos"
         Top             =   150
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         Height          =   450
         Left            =   8385
         Picture         =   "FORPED07.frx":5C5A
         Stretch         =   -1  'True
         ToolTipText     =   "Doble Click Para Importar desde B.Datos MDB"
         Top             =   150
         Visible         =   0   'False
         Width           =   405
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
         Left            =   7965
         TabIndex        =   100
         Top             =   855
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label30 
         Height          =   255
         Left            =   8100
         TabIndex        =   99
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image Image6 
         Height          =   450
         Left            =   8760
         Picture         =   "FORPED07.frx":5DA4
         Stretch         =   -1  'True
         ToolTipText     =   "Doble Click Para Importar desde Planilla Excel"
         Top             =   120
         Visible         =   0   'False
         Width           =   405
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
         Left            =   120
         TabIndex        =   73
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   11865
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible"
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
         Left            =   5775
         TabIndex        =   29
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido Nro."
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
         Height          =   225
         Left            =   8730
         TabIndex        =   27
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label Label4 
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
         Height          =   225
         Left            =   7755
         TabIndex        =   26
         Top             =   1400
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Límite Crédito"
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
         Height          =   230
         Left            =   5775
         TabIndex        =   25
         Top             =   120
         Width           =   1695
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
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   675
         Width           =   960
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FORPED07.frx":6DEE
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "FORPED07.frx":7022
      TabIndex        =   106
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   11160
      TabIndex        =   128
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox Check7 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   360
      Left            =   11280
      TabIndex        =   120
      ToolTipText     =   "Genera Factura en Forma Directa"
      Top             =   6900
      Width           =   225
   End
   Begin VB.Label Label100 
      BackColor       =   &H00624E28&
      Caption         =   "Genera Factura"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   11500
      TabIndex        =   153
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   152
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   105
      Picture         =   "FORPED07.frx":709C
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
      Caption         =   "Transacciones"
      Begin VB.Menu nueped 
         Caption         =   "Nuevo Pedido"
      End
      Begin VB.Menu mnuNuevoPedidoCopia 
         Caption         =   "Nuevo Pedido Por Copia"
      End
      Begin VB.Menu MoPeCur 
         Caption         =   "Modificar Pedido en Curso"
      End
      Begin VB.Menu AnuPed 
         Caption         =   "Anulación de Pedidos"
      End
      Begin VB.Menu RevAnu 
         Caption         =   "Revertir Anulación de Pedidos"
      End
      Begin VB.Menu separa1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAprobarPedidos 
         Caption         =   "Aprobar Pedidos "
      End
      Begin VB.Menu AnuItem 
         Caption         =   "Anulación por Items de Pedidos"
      End
      Begin VB.Menu mnuObservaItemDePedido 
         Caption         =   "Agregar Observaciones a Item de Pedido"
      End
      Begin VB.Menu mnuVisPlantform 
         Caption         =   "Visualizar Ficha Técnica de Items"
         Visible         =   0   'False
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu pedidossuc 
         Caption         =   "Pedidos de Reposición de Sucursal"
         Begin VB.Menu Verocomsuc 
            Caption         =   "Ver Pedido de Reposición de Sucursal"
         End
         Begin VB.Menu descargarocsuc 
            Caption         =   "Descargar Archivos de Pedidos de Reposición de Sucursal"
         End
      End
      Begin VB.Menu separa7 
         Caption         =   "-"
      End
      Begin VB.Menu menuCumpPP 
         Caption         =   "Test de Cumplimiento de Pedidos Pendientes"
         Visible         =   0   'False
      End
      Begin VB.Menu sali 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu ConLis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu RCtaCte 
         Caption         =   "Resumen de Cuenta Corriente"
      End
      Begin VB.Menu S2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAnotadorPedidos 
         Caption         =   "Anotador de Pedidos"
      End
      Begin VB.Menu EncaPe 
         Caption         =   "Encabezados de Pedidos"
      End
      Begin VB.Menu DeNumPe 
         Caption         =   "Detalle por Número de Pedido"
      End
      Begin VB.Menu PePeCli 
         Caption         =   "Pedidos Pendientes por Cliente"
         Begin VB.Menu PeClPan 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu PeCliEx 
            Caption         =   "Imprimir/Exportar"
         End
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
      Begin VB.Menu php2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRankingPedidos 
         Caption         =   "Ranking de Pedidos"
      End
      Begin VB.Menu mnuRastreabilidadPedido 
         Caption         =   "Rastreabilidad Pedido - Factura"
      End
      Begin VB.Menu mnuHistoricoPreciosPedido 
         Caption         =   "Histórico de Precios por Pedido"
      End
      Begin VB.Menu mnuEstadisticaDeVentasyPedido 
         Caption         =   "Estadísticos de Ventas y Pedidos"
      End
      Begin VB.Menu separa3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPedidosAprobVsEntregas 
         Caption         =   "Con Fecha de Aprobación y Entregas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusitstock 
         Caption         =   "Situación de Stocks"
      End
      Begin VB.Menu mnupresuppendientes 
         Caption         =   "Presupuestos Pendientes (Sin Pedido Generado)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPedidoSinOt 
         Caption         =   "Reporte Pedidos Pendientes C/Items Sin O.Trabajo"
         Visible         =   0   'False
         Begin VB.Menu mnupedpenImprimible 
            Caption         =   "Formato Imprimible"
         End
         Begin VB.Menu mnuPedPendinteractivo 
            Caption         =   "Reporte Interactivo"
         End
      End
      Begin VB.Menu mnuItCPedPenSinAprobar 
         Caption         =   "Items C/Pedidos Pendientes Sin Aprobar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPedPendStOf 
         Caption         =   "Consultar Pedidos Pendientes - Orden Fabricación Relacionada "
         Visible         =   0   'False
      End
      Begin VB.Menu Proyepedidos 
         Caption         =   "Visualizar Proyección de Pedidos"
         Visible         =   0   'False
      End
      Begin VB.Menu pedistzona 
         Caption         =   "Pedidos Pendientes C/Stock y Zona"
         Visible         =   0   'False
      End
      Begin VB.Menu separa6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPedPendAprob 
         Caption         =   "Pedidos Pendientes de Aprobación"
      End
      Begin VB.Menu mnuPedAFab 
         Caption         =   "Informe de Pedidos a Fabricar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuPedEmbalar 
         Caption         =   "Informe de Pedidos a Embalar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuBuscaModelo 
         Caption         =   "Buscador de Modelos"
         Shortcut        =   ^M
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPedidosyFabricacion 
         Caption         =   "Pedidos y Fabricación"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepVtasEstadistica 
         Caption         =   "Reportes de ventas para estadísticas"
      End
      Begin VB.Menu mnuRepVtasEstadisticaItemAnulCentrega 
         Caption         =   "Reporte de ventas con items cumplidos con entregas"
      End
      Begin VB.Menu mnuVentasVendedores 
         Caption         =   "Reporte de Vendedores (Fecha y Vended.)"
      End
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuver 
      Caption         =   "Ver"
      Begin VB.Menu mnuverpedidos 
         Caption         =   "Visualizar Pedidos / Despacho - Re-Impresión"
      End
      Begin VB.Menu mnupendienteaprobacion 
         Caption         =   "Pedidos Pendientes de Aprobación"
      End
      Begin VB.Menu mnuEtiqDespacho 
         Caption         =   "Etiquetas de Despacho"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuNuevoItem 
         Caption         =   "Nuevo Item "
         Visible         =   0   'False
      End
      Begin VB.Menu mnuControPedidos 
         Caption         =   "Control de Pedidos "
      End
      Begin VB.Menu CoPag 
         Caption         =   "Condiciones de Pago"
      End
      Begin VB.Menu mnuImportadorPedidos 
         Caption         =   "Importador de Detalle de Pedidos"
      End
      Begin VB.Menu mnuRegDespXLectora 
         Caption         =   "Módulo de  Despachos x Lectora"
      End
   End
   Begin VB.Menu Configura 
      Caption         =   "Configuración"
      Begin VB.Menu mnuAsociarReporte 
         Caption         =   "Vincular Reporte con Formulario"
      End
      Begin VB.Menu separa5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAtributosComercialesPedidos 
         Caption         =   "Atributos Comerciales de Pedidos"
      End
      Begin VB.Menu ConWaldbott 
         Caption         =   "Ruta de conexion a Base de datos Waldbott"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSncrnzCPWlbt 
         Caption         =   "Sincronizar Condición de Pagos Waldbott"
         Visible         =   0   'False
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu ConGral 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu AccDirec 
      Caption         =   "Accesos Directos"
      Begin VB.Menu MaesCli 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu MaesArt 
         Caption         =   "Maestro de Articulos"
      End
      Begin VB.Menu LiPrVe 
         Caption         =   "Listas de Precios de Venta"
      End
      Begin VB.Menu separa4 
         Caption         =   "-"
      End
      Begin VB.Menu abmReserva 
         Caption         =   "ABM Reservas"
         Visible         =   0   'False
      End
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
Dim FilaActiva As Long
Dim RetiraDomicilioAHP%
Dim PercepcionIBMisiones_BienDeUso As Integer



Sub AGREGABONIFICACION(U&, IKITEM%)
       'LEO EL REGISTRO
       X$ = REGLEIDO$("!CARDETPE", U&)
       
       CIXI% = CVI(Left$(X$, 2))
       CAN = CVS(Mid$(X$, 89, 4))
       NROPE& = XVALO(Text11.TEXT)
       FEMI% = FECHANUM(Text9.TEXT)
       NCLIE = XVALO(Text1.TEXT)
       STATUPC% = (-1)
       If Left$(Control$("PEDCLI", "MODOAPRO"), 6) = "DIRECT" Then STATUPC% = 0
       CAVER$ = TRIM$(Str$(NROPE&))
       While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
       While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend

       IK$ = TRIM$(Str$(IKITEM%))
       CAVERI$ = CAVER$ + "-" + IK$
       '
       'CREO UN NUEVO REGISTRO
       strSQL = "Select * from PedDeta where NroPed=" & NROPE&
       Dim PedDeta1 As ADODB.Recordset
       Set PedDeta1 = New ADODB.Recordset
26     On Error Resume Next
       PedDeta1.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
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
           !RasoClie = rasocli$(NCLIE)
           !NROOCOM = TRIM$(Text10.TEXT)
             FENTRESOL% = CVI(Mid$(X$, 63, 2))
           !FeSolEnt = CVDAT(FENTRESOL%)
           !Status = STATUPC%              ' Status  de Aprobación
           !ListPre = XVALO(TEXT21(0).TEXT)
           !CondPag = XVALO(TEXT21(1).TEXT)
           !VENDED = XVALO(TEXT21(2).TEXT)
           !PorComVend = XVALO(Text44.TEXT)
           If Left$(UCase$(Label21(4)), 5) = "LETRA" Then
               valor% = 0
               If TRIM$(Text44) <> "" Then valor% = Asc(Text44.TEXT + "A") - 64
               !PorComVend = valor%
           End If

           !NuOrItem = IKITEM%
           !CODIINT = CIXI%
           !CODIEXT = CODEXT$(CIXI%)
           !SubCod = Mid$(X$, 109, 8)
           !Descrip = DESCRIT0$(CIXI%)
           !TCOLOR = Mid$(X$, 3, 16)
           !TTALLE = Mid$(X$, 19, 8)
           !TLargo = CVS(Mid$(X$, 27, 4))
           !TAncho = CVS(Mid$(X$, 31, 4))
           !TEspesor = CVS(Mid$(X$, 35, 4))
           !VenRep = Left$(Label11, 1)   ' Venta o Reparacion
           !MoneEmis = MONEMI%
           !TIPOCAM = XVALO(TICAMBIS)
           !PreList = CVS(Mid$(X$, 39, 4))
           !CanAsig = 0
           !CanPed = CAN
           !PreUnid = 0   ' CVS(Mid$(X$, 51, 4))
           !ImpoTot = 0
           !CANTENT = 0
           !CanSaldo = CAN
           !FeUltEnt = CVDAT(0)
           !NroRemi = ""
           !TPedItem = CAVERI$
           '
           '\\\OT-8-766-19/09/08
           !ImpoAnti = 0
           !fapro_nope = CVDAT(FECHANUM("01/01/80")) 'PARA EVITAR PROBLEMAS CON FORMATO REGIONAL
           !AcuPrepa = 0
           !AcuDespa = 0
           !TipoPed = "" ' se corroboro que este CAMPO es de  TEXTO
           !finvigen = CVDAT(FECHANUM("01/01/80"))
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
  Campos$ = "Código/F16;Descripción/A24;Cant.Sol./F10.1;Cant.Unid./F10.1;Sol.Entr./D8;CI1/F0;LPRE/F0"
  Call CARGA_ENCABEZADO(STOCMOLD.MSF1, Campos$, STOCMOLD, 0, 0)
  '
  For U& = 1 To FIN&
    X$ = REGLEIDO$("!CARDETPE", U&)
    CI1% = CVI(Left$(X$, 2))
    CAN = CVS(Mid$(X$, 47, 4))
    LPREXX% = CVS(Mid$(X$, 85, 4))
    FENTRESOL% = CVI(Mid$(X$, 63, 2))
    CODI$ = CODEXT$(CI1%)
    Descr$ = DESCRIT$(CI1%)
    UNME$ = Unimed$(CI1%)
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
    Campos$ = "Código/F16;Descripción/A24;Cant.Sol./F10.1;Cant.Unid./F10.1;Sol.Entr./D8;CI1/F0;LPRE/F0"
    With STOCMOLD
      .MSF1.Rows = .MSF1.Rows + 1
      .MSF1.TextMatrix(.MSF1.Rows - 1, 1) = CODI$
      .MSF1.TextMatrix(.MSF1.Rows - 1, 2) = Descr$
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
        Ci% = CVI(Left$(Y$, 2))
        If Ci% > 0 And Ci% <= NUMAS% Then
            MODO% = 0
            CANTIBONIF = 0
            CANPEDIDA = CVS(Mid$(Y$, 47, 4))
            DESCUXX$ = DESCOMBINA$(NCLIEN, Ci%, MODODESC%): MODO% = MODODESC%
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
Function CARGAVENDEDORENVA%()
        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "CODIGO/F8;REFERENCIA/A48", FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "CODIGO/F8;REFERENCIA/A48", FRMZELECHO)
        FRMZELECHO.Caption = "CONDICIONES DE PAGO"

        Dim obj As New RECORXET
        Set rs1 = obj.RS_CONDIPAGO()
        Call Carga_MSF_Recordset(rs1, "CODIGO/F8;REFERENCIA/A48", FRMZELECHO.msf2)
        Set obj = Nothing
        rs1.Close
        Set rs1 = Nothing
        FRMZELECHO.Width = 9500
        FRMZELECHO.Show 1
        CARGAVENDEDORENVA% = XVALO(RESP_ZELE_VECT(1))
End Function

Sub DEPURAPED()
    '
    'DEPURACION DE PEDIDOS SEGUN CONFIGURACION DE SETUP DE PEDIDOS
    'SI TIENE MARCA DE CONTROL DE DEPURACION Y EL CASILLERO DE TO LERANCIA ES > 0
    'RECORRE LOS PEDIDOS PENDIENTES SI ESTOS TIENEN FECHA DE ENTREGA MENOR A LA DE HOY Y SUPERAN
    'LA TOLERANCIA CONFIGURADA LE CAMBIA EL STATUS A 9
    CantiDias% = XVALO(Control("ADMIPED", "DEPEPEN"))
    If Control("ADMIPED", "DEPUPPEN") = "SI" And CantiDias% > 0 Then
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
             FESOL% = CVINT(!FeSolEnt)
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

Sub DESCUMONTO(Form1 As Form)

    'CALCULA EL DESCUENTO SEGUN VALIDACION DE SETUP DE DESUENTO POR MONTO
    'SE PASA COMO PARAMETRO EL FORM PARA UTILIZAR LA MISMA EN LA MODIFICACION DEL PEDIDO
    'SI TIENE VENDEDOR DECUELVE EL DESCUENTO 1 CASO CONTRARIO DEVUELVE EL DESCEUENTO 2
    'ESTO TOMA SOLO PARA LA LISTA 1
    DESC1 = 0: DESC2 = 0
    '
    TOTPEDIDO = XVALO(Form1.Total(0))
    'DEVUELVE EN DESC1 Y DESC2 EL DESCUENTO DE ACUERDO AL IMPORTE
    For i& = 1 To ULTREG&("DESMONTO")
       X$ = REGLEIDO$("DESMONTO", i&)
       DESDE1 = CVS(Mid$(X$, 1, 4))
       HASTA1 = CVS(Mid$(X$, 5, 4))
       If TOTPEDIDO >= DESDE1 And TOTPEDIDO <= HASTA1 Then
            DESC1 = CVS(Mid$(X$, 9, 4))
            DESC2 = CVS(Mid$(X$, 13, 4))
            Exit For
       End If
    Next i&
    '
    'POR SI NO COINCIDE  EN NINGUNO Y EL TOTAL DEL PEDIDO SUPERA EL  HASTA SE QUEDA CON EL ULTIMO
    If TOTPEDIDO > HASTA1 Then
       DESC1 = CVS(Mid$(X$, 9, 4))
       DESC2 = CVS(Mid$(X$, 13, 4))
    End If
    '
    Vend% = XVALO(Form1.TEXT21(2))
    VENDED$ = TRIM$(ECOARCH$("VENDEDOR", Chr$(Vend%)))
    If Vend% = 0 Or InStr(UCase$(VENDED$), "GABAL") Then
       Form1.PORDESCU = DESC2
    Else
       Form1.PORDESCU = DESC1
    End If
    '
End Sub

Sub GENERAOF_ENVAPLAST(NPED&)
    '
'   Screen.MousePointer = 11
'   Avance.PROCESOX$ = "GENERAOF_ENVAPLAST" & "/" & NPED&
'   On Error Resume Next
'   Avance.Show 1
'   On Error GoTo 0
'   Screen.MousePointer = 1
'   Exit Sub
   

    If NPED& > 0 Then
        CODI$ = ""
        For U& = 1 To ULTREG&("!CARDETPE")
         X$ = REGLEIDO$("!CARDETPE", U&)
         CAN = CVS(Mid$(X$, 47, 4))
         If CAN > 0.05 Then
           '
           CIXI% = CVI(Left$(X$, 2))
           NUOR% = CVI(Mid$(X$, 123, 2))
           If CIXI% > 0 Then
              If TIMATE%(CIXI%) < 2 Then GoTo 50
              FENTRESOL% = CVI(Mid$(X$, 63, 2))
              If TRIM$(APROBO$(CIXI%)) = "" Then
                 If CODI$ = "" Then
                   CODI$ = CODI$ + CODEXT$(CIXI%)
                   Articulo$ = " El "
                   GoTo 50
                 Else
                   CODI$ = CODI$ & " , " & CODEXT$(CIXI%)
                   Articulo$ = " Los "
                   GoTo 50
                 End If
              End If
              Call FORPED07.GENERAR_OF_SIMPLE(CIXI%, XVALO(Text1), CAN, FENTRESOL%, NPED&, TRIM$(Text10), RETORNO_NORFA$, NUOR%)
              FILTX$ = TRIM$(Str$(CIXI%)) & ";" & TRIM$(Str$(NPED&)) & ";" & TRIM$(RETORNO_NORFA$) & ";" & TRIM$(MEDIDA_ENVAPLAST$(CIXI%, RETCORTE#))
              Call STDFORM("ORTR-ENV", FILTX$, "PRINT")
           End If
         End If
50       Next U&
         '
         If CODI$ <> "" Then
            TXX1$ = "" & Articulo$ & "Código" & Mid$(Articulo$, 3, 1) & " " & CODI$ & " Faltan ser Aprobados No se Generará la Orden de Trabajo Para Estos"
            Call COMUNI(TXX1$)
         End If
         '
     End If
End Sub

Sub GENERAR_OF_SIMPLE(CIXI%, NC, CAN, FENTRESOL%, NPED&, ORCOMP$, RETORNO_NORFA$, Optional NORD%)
   Call GENEOF(CIXI%, NC, CAN, FENTRESOL%, NPED&, ORCOMP$, RET_NORFA$, NORD%)
   NORFA$ = RET_NORFA$
   Call GENERES(CIXI%, CAN, NORFA$)
   Call ACTUECO(CIXI%, CAN, RET_NORFA$)
   RETORNO_NORFA$ = RET_NORFA$
End Sub

Sub GENERAR_TABLA_VERSION_IMAGEN()
   Call CREACAMPO("", "IMAG_VERSION", "NroOrig", 4, 0, 1)
   Call CREACAMPO("", "IMAG_VERSION", "NVERSION", 4, 0)
   Call CREACAMPO("", "IMAG_VERSION", "Referencia", 10, 128) 'EL TIDOCUM$ la misma que en el pdoc tbl ej: 'Pedido de Cliente'
   Call CREACAMPO("", "IMAG_VERSION", "NumDocPdoc", 4, 0, 1)
   Call CREACAMPO("", "IMAG_VERSION", "FechEmis", 8, 0)
   Call CREACAMPO("", "IMAG_VERSION", "CODIEMPR", 3, 2, 1)
End Sub

Sub GENERES(CIXI%, CAN, NORFA$)
    If NORFA$ = "" Then Exit Sub
    'USADA EN ENVAPLAST
    NORFX$ = TRIM$(NORFA$)
    SQLX$ = "DELETE FROM RESERVA WHERE IdSubOr = '" & NORFX$ & "'"
    FLEXCONN.Execute (FILTSQL$(SQLX$))
    '
    NORIT% = 0
    CODD$ = CODEXT$(CIXI%)
    Call LEEEXPLO(CODD$, 1, 1)


    For U& = 1 To CAIT%
      XXX$ = EXPLOLIN$(U&)
      CODI$ = Mid$(XXX$, 1, 16)
      Descr$ = Mid$(XXX$, 17, 28)
      un$ = Mid$(XXX$, 47, 2)
      Uso = XVALO(Mid$(XXX$, 49, 12))
      '
      cirf% = CODINT(CODI$)
      CANTRX# = Uso * CAN
      USOIX = Uso
      IDLOT$ = ""
      DEPOX% = 1
      '
      SQLX$ = " SELECT * FROM Reserva WHERE COD_INTE < 1 "
      Dim RESERVVVA As ADODB.Recordset
      Set RESERVVVA = New ADODB.Recordset
      RESERVVVA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      'With Reserva
      With RESERVVVA
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CODIEMPR = CodiEmp%
        On Error GoTo 0
        !cod_inte = cirf%
        !Cod_Exte = Left$(CODEXT$(cirf%), 16)
        !Descrip = Left$(DESCRIT0$(cirf%), 64)
        !idsubor = Left$(NORFX$, 12)
        !FechRes = Now
        !CANTRES = CANTRX#
        !CANTCONS = 0
        !Saldo_Entre = CANTRX#
        !idenlote = Left$(IDLOT$, 16)
        !DEPOSITO = DEPOX%
        !UsoUnit = USOIX
           NORIT% = NORIT% + 1
        !NuOrItem = NORIT%
        !NUME_SOLINT = 0 ' INICIALIZACION DE CAMPO
        !MarBorra = 0
        .Update
      End With
    Next U&
    RESERVVVA.Close
    Set RESERVVVA = Nothing
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    
    'se comenta lo siguiente dado que no andaba correctamente, se cambio en situacion de stock
    'para que muestre acumulado las cantidades de una misma of mismo item.
    
'    If InStr(EPAC$, "ENVAPLAST") > 0 Then
'      'ESTO ES PARA QUE REALICE LA RESERVA DE LOS ITEMS DUPLICADOS EN 1 SOLO REGISTRO
'      NORIT% = 0
'      SQLX$ = " SELECT * FROM Reserva WHERE COD_INTE < 1 "
'      Dim RS1 As ADODB.Recordset
'      Set RS1 = New ADODB.Recordset
'      RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'
'      SQLX$ = " SELECT COD_INTE,IDENLOTE,SUM(CANTRES) AS CANRES,COD_EXTE,Descrip,IDSUBOR,FechRes"
'      SQLX$ = SQLX$ + ",CANTCONS,Saldo_Entre,DEPOSITO,NUME_SOLINT FROM Reserva WITH(NOLOCK) "
'      SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFX$ & "' "
'      SQLX$ = SQLX$ + " GROUP BY COD_INTE,idenlote,COD_EXTE,Descrip,IDSUBOR,FechRes "
'      SQLX$ = SQLX$ + " ,CANTRES,CANTCONS,Saldo_Entre,DEPOSITO,NUME_SOLINT "
'      SQLX$ = SQLX$ + " ORDER BY COD_INTE,IDENLOTE"
'      Set RS = READSET(SQLX$)
'      With RS
'        Do While Not .EOF
'          RS1.AddNew
'          RS1!CodiEmpr = CodiEmp%
'          RS1!COD_INTE = XVALO(!COD_INTE)
'          RS1!IDENLOTE = SAFETEXT$(!IDENLOTE)
'          RS1!CANTRES = XVALO(!CANRES)
'          RS1!COD_EXTE = !COD_EXTE
'          RS1!Descrip = !Descrip
'          RS1!IDSUBOR = !IDSUBOR
'          RS1!FechRes = !FechRes
'          RS1!CANTCONS = !CANTCONS
'          RS1!Saldo_Entre = !Saldo_Entre
'          RS1!IDENLOTE = !IDENLOTE
'          RS1!Deposito = !Deposito
'          CIX = CIXI%
'          RS1!USOUNIT = USOUNITARIO#(CIX, !COD_INTE)
'           NORIT% = NORIT% + 1
'          RS1!NuOrItem = NORIT%
'          RS1!NUME_SOLINT = 0 ' INICIALIZACION DE CAMPO
'          RS1!MarBorra = 0
'          RS1.Update
'        .MoveNext
'      Loop
'      End With
'      RS.Close
'      Set RS = Nothing
'      RS1.Close
'      Set RS1 = Nothing
'      Call BORRAREGISTROS("RESERVA", "IDSUBOR = '" & NORFX$ & "' AND MARBORRA > 0 ", REGAF&, "NOMESS")
'      Call ACTUREGISTRO("RESERVA", "MARBORRA", "IDSUBOR = '" & NORFX$ & "' AND MARBORRA > 0", 0, REGAF&, "NOMESS")
'    End If
'      SQLX$ = " SELECT * FROM Reserva WHERE COD_INTE < 1 "
'      Dim RS1 As ADODB.Recordset
'      Set RS1 = New ADODB.Recordset
'      RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'      With RS
'          .AddNew
'          !CodiEmpr = CodiEmp%
'          !COD_INTE = XVALO(!COD_INTE)
'          !IDENLOTE = SAFETEXT$(!IDENLOTE)
'          !CANTRES = XVALO(!CANRES)
'          !COD_EXTE = !COD_EXTE
'          !Descrip = !Descrip
'          !IDSUBOR = !IDSUBOR
'          !FechRes = !FechRes
'          !CANTRES = !CANTRES
'          !CANTCONS = !CANTCONS
'          !Saldo_Entre = !Saldo_Entre
'          !IDENLOTE = !IDENLOTE
'          !DEPOSITO = !DEPOSITO
'          CIX = CIXI%
'          !USOUNIT = USOUNITARIO#(CIX, !COD_INTE)
'           NORIT% = NORIT% + 1
'          !NuOrItem = NORIT%
'          !NUME_SOLINT = 1 ' INICIALIZACION DE CAMPO
'          .Update
'          .MoveNext
'      End With
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
       CantiDias% = XVALO(Control("ADMIPED", "DEPEPEN"))
       If Control("ADMIPED", "DEPUPPEN") = "SI" And CantiDias% > 0 Then
       
       FEHOY% = HOY(HO$)

       SQLX$ = "SELECT * FROM PEDDETA WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE STATUS > 8 " 'SOLO LAS ANULDAS
       SQLX$ = SQLX$ + " AND (CANPED - CANTENT) > 0.05 "
       SQLX$ = SQLX$ + " "
       '
       Set PEDDETTEMP = READSET(SQLX$)
       With PEDDETTEMP
           Do While Not .EOF
             FESOL% = CVINT(!FeSolEnt)
             'SI LA CANTIDAD DE DIAS DE LA FECHA SOLICITADA COMPARADA CON LA DE HOY ES > A TOLERANCIA
             If DIENTRE(FESOL%, FEHOY%) > CantiDias% Then
                NP& = XVALO(!NROPED)
                NOR% = XVALO(!NuOrItem)
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


Function ValidaActualizaPresupVsCargadePedido%()
   Maximo& = 0
   ValidaActualizaPresupVsCargadePedido% = 1
   Dim VectorCotizacion()
   ReDim Preserve VectorCotizacion(1)
   'CARGO EN UN VECTOR CON INDICE DEL CI% COMO VALOR EL NRO DE COTIZACION
   For KKI% = 1 To List2.ListCount
      If List2.Selected(KKI% - 1) = True Then
        XX$ = List2.List(KKI% - 1)
        POS1% = InStr(XX$, "Cot:")
        POS2% = InStr(XX$, "F.T:")
        POS1% = InStr(XX$, "Presup.No:") '10
        POS2% = InStr(XX$, "-")
        
        POSCOD1% = InStr(XX$, "-|")
        POSCOD2% = InStr(XX$, "|-")
        If POSCOD1% > 0 And POSCOD2% > 0 Then
          Codigo$ = Mid$(XX$, POSCOD1% + 2, POSCOD2% - (POSCOD1% + 2))
          CI1% = CODINT(Codigo$)
          If POS1% > 0 And POS2% > 0 Then
            NROPRES& = XVALO(Mid$(XX$, POS1% + 10, POS2% - (POS1% + 1)))
            If NROPRES& > 0 Then
              If Maximo& < CI1% Then Maximo& = CI1%: ReDim Preserve VectorCotizacion(CI1%)
              VectorCotizacion(CI1%) = NROPRES&
            Else
              Call COMUNI("No se Pudo Registrar el Nro de Cotización Imposible Continuar")
              ValidaActualizaPresupVsCargadePedido% = -1
              Exit Function
            End If
          End If
         End If
      End If
   Next KKI%
   '
   'AHORA RECORRO EL CARDETPE
'  Y LE ASIGNO A CADA CODIGO EL NRO DE COTIZACION CORRESpNDIENTE
   'SI HAY ITEMS CARDADOS QUE NO POSEEN COTIZACION MENSAJE Y SALE
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CI1% = CVI(Left$(X$, 2))
     If CI1% > 0 Then
        If UBound(VectorCotizacion) < CI1% Then GoTo 80
        NUPRESU& = VectorCotizacion(CI1%) 'LE ASIGNO EL NRO DE COTIZACION SEGUN EL CODIGO LEIDO DEL LIST2 PARRAFO DE ARRIBA
        If NUPRESU& > 0 Then   'EXCEPCION POR CODIGO DE FAMILIA DE POLIMEROS
          Call REPLA(X$, MKS$(NUPRESU&), 125, 4)
          Call GRAREG("!CARDETPE", X$, U&)
        Else
        
80      FAMI$ = CODFAMIL$(CI1%)
'        If (FAMI$ <> "15" And FAMI$ <> "CT" And FAMI$ <> "TE") Then 'EXCEPCION POLIMEROS Y CORTES.
         If (FAMI$ = "BO" Or FAMI$ = "BB") Then  'PRESENTA MENSAJE SOLO CUANDO SON BOLSAS O BIGBAG
            OPX% = COMALTER%("Este Pedido el Código: " & CODEXT$(CI1%) & "\No está Enlazado Con una Cotización\\Cancelar\Grabar Igual (Solo Con Derechos)")
            If OPX% < 2 Then
               ValidaActualizaPresupVsCargadePedido% = -1
               Exit Function
            Else
               FIRMAELEC.Text1 = ""
               FIRMAELEC.Text2 = ""
               FIRMAELEC.Label1 = ""
               FIRMAELEC.Show 1
               If FIRMAELEC.ESTADO > 0 Then
                   If NIVCONFI%(XVALO(FIRMAELEC.Label1)) < 3 Then
                      Call COMUNI("Usuario Con Nivel de Clasificación Inválido Para Esta Operación")
                      ValidaActualizaPresupVsCargadePedido% = -1
                      Exit Function
                   End If
               Else
                    ValidaActualizaPresupVsCargadePedido% = -1
                    Exit Function
               End If
            End If
         End If
        End If
     End If
   Next U&

End Function

Sub Verifica_Impresora_Despacho()
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(UCase(EPAC$), "MEDITEA") > 0 Then
       If TRIM$(Control(IDENTER$, "PUERDESP")) = "" Or TRIM$(Control$(IDENTER$, "PUERORVE")) = "" Then
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

Private Sub Check1_Click()
   Me.cmdtext.Visible = False
   If Check1.Value = 1 Then Me.cmdtext.Visible = True
End Sub

Private Sub Check2_Click()
   Vend% = XVALO(TEXT21(2))
   VENDED$ = TRIM$(ECOARCH$("VENDEDOR", Chr$(Vend%)))
   If Vend% = 0 Or InStr(UCase$(VENDED$), "GABAL") Then
      If Check2.Value = 1 Then
       Label31 = XVALO(Control("ADMIPED", "DESRTCTE"))
      Else
       GoTo 20
      End If
   Else
20     Label31 = ""
       Check2.Value = 0
   End If
End Sub

Private Sub Check3_Click()
    Dim estadoVisible As Boolean
    Dim estadoVisible2 As Boolean
    estadoVisible = Frame3.Visible
    estadoVisible2 = CheckAhpRetiraDep.Visible
    ' Llamar a la función Conmutar para cambiar el valor entre true o false segun posea al momento
     Frame3.Visible = Conmutar(estadoVisible)
     CheckAhpRetiraDep.Visible = Conmutar(estadoVisible2)
End Sub

Private Sub cmdtext_Click()
   Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0ORDER BY NROPED DESC", "NOMESS")

'    NCLIEN = XVALO(Text1.TEXT)
'    CONDICION$ = " NroClie =  " & NCLIEN & " AND Status < 8 AND CantEnt < (CanPed - 0.005) AND  Status <> 3 "
'
'    Call COPYSTRU("INDIPEP", "INDIPNN")
'    '
'
'    Call CARGALISEL("INDIPNN", "PEDDETA", "NroPed/F5.0;RaSoClie/A20;FechEmis/D8", CONDICION$, "DISTINCT")
'    '
'    If ULTREG&("INDIPNN") < 1 Then Exit Sub
'
'    '
'10  Call SELECHO("INDIPNN")

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
   'Call SELECHO("DEUDOR1")
   'NCLIEN = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   'Text1.TEXT = TRIM$(Str$(NCLIEN))
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "CONTEPLAST") > 0 Then
      'FILTRA LOS CLIENTES POR VENDEDOR
      USNU% = USNBR% Mod 100
      VENDESEGUNUSUARIO% = XVALO(Control("LOGVEND", SAFETEXT(USNU%)))
      '
      NCLIEN = NCLLICONTEPLAST(VENDESEGUNUSUARIO%)
    Else
      'MODO TRADICIONAL
      NCLIEN = ListaSeleccionClientes()
    End If
   
    If NCLIEN > 0 Then
     Text1.TEXT = TRIM$(Str$(NCLIEN))
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
   EPAC$ = TRIM$(UCase$(EMPREAC$)) '
   If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
    If DerechoSupervisor% < 1 Then
     If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then
       GoTo 99
     End If
    End If
   End If
   
   PROLLAMA$ = "CERR"
   OPPED07.Show 1
   If MONEMI% < 1 Then MONEMI% = 1
   MONEMIS = DEMONECO$(MONEMI%)
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   
99 Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   Call REVERTIR_CONTROL
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   '
   RetiraDomicilioAHP% = 0
   If CheckAhpRetiraDep.Value > 0 Then RetiraDomicilioAHP% = 1
   
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
   
   If InStr(EPAC$, "CONTEPLAST") > 0 Then
      VENDEDENVAP$ = TRIM$(TEXT21(4))
     If XVALO(VENDED$) > 1 Then '> 1 por que si es gerencia no hace falta
        ECOVEN$ = TRIM$(ECOARCH$("VENDEENV", Chr$(XVALO(VENDEDENVAP$))))
        If ECOVEN$ = "" Then
          Call MENSERR(24, "Debe Ingresar El Tipo de Vendedor Envaplast")
           On Error Resume Next
           TEXT21(4).SelStart = 0
           TEXT21(4).SelLength = Len(TEXT21(4).TEXT)
           TEXT21(4).SetFocus
        On Error GoTo 0

          GoTo 99
        End If
     End If
   End If
   '
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
   If Control$("ADMIPED", "NOREPOC") = "SI" And TRIM$(Text10) <> "" Then
      NCLIE = XVALO(Text1)
      NORCOM$ = UCase$(TRIM$(Text10))
      If OCOMEXISTE%(NCLIE, NORCOM$) Then
         Call COMUNI("Orden de Compra Existente")
         Text10 = ""
         GoTo 99
      End If
   End If
   '
   If Control$("ADMIPED", "EXICOPA") = "SI" Then
      If ECOARCH$("CONPAG", MKI$(XVALO(TEXT21(1).TEXT))) = "" Then
        Call COMUNI("Imposible Continuar. Debe Elegir Una Condición de Pago Válida.")
        GoTo 99
      End If
   End If
   If Control$("ADMIPED", "VALIMONE") <> "NO" Then
    If MONELISTA%(XVALO(TEXT21(0))) <> MONEMI% Then
         OPX% = COMALTER%("Atención !!!\ Está Intentando Grabar el Pedido\En una Moneda Diferente a la de la Lista de Precios Activa\\Volver y Modificar\Grabar Igual")
         If OPX% < 2 Then GoTo 99
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
   If Control$("ADMIPED", "VALTIPIN") = "SI" Then
        If XVALO(TEXT21(3)) < 1 Then
            Call COMUNI("Falta Elegir un Tipo de Ingreso")
            GoTo 99
        End If
   End If
   '
   'VALIDACIÓN QUE LOS  ITEMS TRAZABLES VAYAN DE A UNO
   For U& = 1 To ULTREG&("!CARDETPE")
      X$ = REGLEIDO$("!CARDETPE", U&)
      CAN = CVS(Mid$(X$, 47, 4))
      If CAN > 1 Then
        CIXI% = CVI(Left$(X$, 2))
        If ESTRAZABLE%(CIXI%) = 2 Then
           Call MENSERR(24, "Cantidad No Valida (En Item: " & CODEXT$(CIXI%) & ") - El Máximo Permitido es 1 (uno)\en los Articulos Trazables")
           GoTo 99
        End If
      End If
   Next U&
   
   MODOAPROB$ = "" ' VARIABLE PARA FORZAR LA APROBACION POSTERIOR EN CASO QUE NO PASE LAS VALIDACIONES
   'CONTROL DE VALIDACION DE CONDICIONES COMERCIALES
   NC = XVALO(Text1)
   If TRIM$(Control("ADMIPED", "VALCONPA")) = "SI" Then
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
   If TRIM$(Control("ADMIPED", "FACVENCI")) = "SI" Then
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
   DIAS_ULT_COMPRA% = XVALO(Control("ADMIPED", "QULTICOM"))
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
   If Control("ADMIPED", "DOMIMAND") = "SI" Then
     If TRIM$(LENTREGA(0).TEXT) = "" Then
        Call COMUNI("Falta Ingresar Domicilio de Entrega")
        GoTo 99
     End If
   End If
   '
   'VALIDO SI ESTA CONFIGURADA EL CARGO DE FLETE Y SI ES ASI LO EXIGE
   If Control("ADMIPED", "FLETMAND") = "SI" Then
        If TRIM$(Me.TXTFLETE) = "" Then
           Call COMUNI("Falta Ingresar a Quien Corresponde el Cargo del Flete")
           GoTo 99
        End If
   End If
   '
   ' PERCEPCION IIBB MISIONES (NO APLICA A CONSUMIDORES FINALES)
   ' SI LOS ARTICULOS DEL DETALLE SERAN UTILIZADOS COMO PARTE DE LA ACTIVIDAD (BIEN DE USO) NO SE DEBE PERCIBIR,
   ' SI EN CAMBIO SE TRATA DE UN ARTICULO DE REVENTA (BIEN DE CAMBIO) SI DEBE PERCIBIRSE
   With PERCEIIBB
        NCLIE = XVALO(Text1)
        If .ImportePercepcionMisiones > 0.005 And PIVACLI%(NCLIE) <> 0 Then
            OpcionUsoDelBien% = COMALTER%("El Presente Pedido cuenta con Percepción de IIBB Misiones.\Consultar si los Artículos del Detalle\serán Utilizados como Bien de Uso o de Cambio.\\Bien de Cambio\Bien de Uso")
            If OpcionUsoDelBien% = 2 Then
                PercepcionIBMisiones_BienDeUso = 1
                Call CALTOPED
            End If
        End If
   End With
   '
    CTRLSINCARGO% = ControlItemsSinCargo% ' FUNCION
    '0= TODO BIEN
    '1= TIENE ITEMS SIN CARGO PERO ESTAN APROBADOS
    '-1 TIENE ITEMS SIN CARGO QUE ESTAN EN LA LISTA Y NECESITAN APROBACION
    '-2 TIENE ITEMS  QUE NO ESTAN EN LA LISTA Y LA DIFERENCIA ES MAYOR A LA TOLERADA CONTRA LA LISTA
    '3 TIENE ITEMS CON VALOR 0 QUE NO ESTAN EN LA LISTA (SE CONSIDERA POR SI NO HAY PORCENTAJE DE TOLERANCIA
    If CTRLSINCARGO% < 0 Then
        If CTRLSINCARGO% = (-1) Then
            OPX% = COMALTER%("Control de Importes.\ \Se Han Detectado\" & TEXTOMENSACTRLSCARGO$(CTRLSINCARGO%) & "\\Cancelar\Continuar y Grabar Pedido Pendiente de Aprobación")
            If OPX% < 2 Then
              GoTo 99
            End If
            MODOAPROB$ = "PENDIENTE DE APROBACION POR MATERIALES SIN CARGO"
            'EN GRAPEDIDO NO PRESENTA MENSAJE YA LO PRESENTO EN ESTE IF
        ElseIf CTRLSINCARGO% = (-3) Then
            Call COMUNI("Control de Importes.\ \Imposible Continuar con el Presente Pedido,\Se Han Detectado\" & TEXTOMENSACTRLSCARGO$(CTRLSINCARGO%))
            GoTo 99
        End If
    End If
'    If CTRLSINCARGO% < 0 Then
'        OPX% = COMALTER%("Control de Importes\Se Han Detectado \" & TEXTOMENSACTRLSCARGO$(CTRLSINCARGO%) & "\\Cancelar\Continuar y Grabar Pedido Pendiente de Aprobación")
'        If OPX% < 2 Then
'          GoTo 99
'        End If
'        MODOAPROB$ = "PENDIENTE DE APROBACION POR MATERIALES SIN CARGO"
'        'EN GRAPEDIDO NO PRESENTA MENSAJE YA LO PRESENTO EN ESTE IF
'    End If
   If InStr(EPAC$, "CONTEPLAST") > 0 Then
     If ValidaActualizaPresupVsCargadePedido% <= 0 Then GoTo 99
   End If
   
   If InStr(EPAC$, "AHP") > 0 Then
      If Check3.Value > 0 And RetiraDomicilioAHP% < 1 Then
         CONDI$ = "Nume_Cli =" & XVALO(Text1) & " AND STATUS >=0"  'SI NO TIENE SUCURSALES NO ENTRA
         cantidadSucursales& = CantRegistros&("SUCENTRE", CONDI$, "NOMESS")
         If cantidadSucursales& > 0 Then
            Entrega.SucuEntrega = "" 'SI SOLO TIENE UNA SUCURSAL, MUESTRA ESA
            Entrega.NroCliente = XVALO(Text1)
            Entrega.NroPedido = ""
            If cantidadSucursales& = 1 Then Entrega.SucuEntrega = XVALO(VALOBADA("SUCENTRE", "Nume_Suc", CONDI$, "NOMESS"))
            Entrega.Show 1
            If Entrega.Label1.Caption <> "OK" And Entrega.Label1.Caption <> "OK-SIN DATOS" Then GoTo 99
            LENTREGA(2) = Entrega.DenoTransporte
            Text20.TEXT = XVALO(Entrega.SucuEntrega)
            Text19.TEXT = XVALO(Entrega.NumeTransporte)
            LENTREGA(0).TEXT = Entrega.DomiTransporte.TEXT
            LENTREGA(1).TEXT = Entrega.LocTransporte.TEXT
         End If
      End If
   End If
   Call GRAPEDIDO(OKX%, MODOAPROB$)
   If InStr(EPAC$, "AHP") > 0 And OKX% > 0 Then
     If RetiraDomicilioAHP% > 0 Then
       Call Entrega.BLANFORMU
       Entrega.SucuEntrega = ""
       Call Entrega.grabarDatosEntregaAhp(XVALO(Text11), XVALO(Text1))
       CONDI$ = "NROPED = " & XVALO(Text11)
       CONDI$ = CONDI$ + " AND NROCLIE = " & XVALO(Text1)
       Call ACTUREGISTRO("DATOENTREGA", "EntEnDomi", CONDI$, 1, REGAF&)
       Unload Entrega
     End If
     If Check3.Value > 0 And RetiraDomicilioAHP% < 1 Then
        If Entrega.Label1.Caption = "OK" Then
           Call Entrega.grabarDatosEntregaAhp(XVALO(Text11), XVALO(Text1))
        End If
     End If
   End If
   '
   Call Entrega.BLANFORMU
   Call DATMOTIVO
   
   If OKX% > 0 Then
     '
     CONPRIPE.Check7.Value = 1
     If Control$("PEDIDO", "PRINOVEN") = "NO" Then CONPRIPE.Check7.Value = 0
     CONPRIPE.Check8.Value = 0
     If Control$("PEDIDO", "PRINODES") = "SI" Then CONPRIPE.Check8.Value = 1
     CONPRIPE.Check2.Value = 0
     If Control$("PEDCLIEN", "SUPRINT") = "SI" Then CONPRIPE.Check2.Value = 1
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
        ElseIf CONPRIPE.Check2.Value = 0 Or CONPRIPE.Check7.Value = 0 Then
          RestauraCancelPrint% = CANCELPRINT%
          CANCELPRINT% = 1
          Call PRIPEDIDO(1)
          CANCELPRINT% = RestauraCancelPrint%
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
     If Control$("ADMIPED", "GENORFAB") = "SI" Then GENORFAB% = 1
     If STATUPC% = (-1) Then
          Call COMUNI("Pedido Registado como\'PENDIENTE DE APROBACIÓN'")
        Else
          ' GENERA ORDEN DE FABRICACION SOLO SI PEDIDO CONFIRMADO
          If GENORFAB% = 1 Then
             Call GENORFA
          End If
     End If
     '
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "ENVAPLAST") > 0 And STATUPC% >= 0 Then
        Call GENERAOF_ENVAPLAST(XVALO(Text11))
     End If
     '

     
     Screen.MousePointer = 11
     If InStr(EPAC$, "TANIT") < 1 Then
'      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")

     ' *** LA SIGUIENTE LINEA SE COMENTA PORQUE APARENTEMENTE SE UTILIZO ANTE UN PROBLEMA EN LA LISTA DE SELECCION
     ' *** DEL FORMULARIO CONPED07. LUEGO DE APLICAR LA NUEVA LISTA TIPO GRILLA ESTA LLAMADA NO SE COMENTO.
'      Call ACTUALIZA_DAT_INDIPEDI
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
       If Control("PEDCLI", "MODOAPRO") = "DIRECT" Then       ' APROBACION DIRECTA
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
   'Call SETULTREG&("!CARDETPE", 0)
   
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
Sub PRINCOMPO(Ci%, CANTI)
        CNTDC1$ = CNTDC(Ci%)
        Call AREXPLO(Ci%, 1, 1)
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
   EPAC$ = UCase$(TRIM$(EMPREAC$))
   If InStr(EPAC$, "CONTEPLAST") > 0 Then
      SeleccionoPresupuesto% = 0
      For KKI% = 1 To List2.ListCount
         If List2.Selected(KKI% - 1) = True Then
            NUPRESU& = XVALO(Mid$(List2.List(KKI% - 1), 251, 6))
           If NUPRESU& > 0 Then
              SeleccionoPresupuesto% = 1
              Exit For
           End If
         End If
      Next KKI%
      If SeleccionoPresupuesto% < 1 Then
         If DERACCE%("GEPESICO", "Generación de Pedidos Sin Cotización Previa") < 2 Then
           Call COMUNI("Imposible Realizar Pedidos Sin Cotización")
           Call FINAL("")
         End If
      End If
   End If

   
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
       NUPRESU& = XVALO(Mid$(List2.List(KKI% - 1), 251, 6))
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
                   CAPRESU& = XVALO(!CANTIT)      ' CVS(Mid$(XX$, 101, 4))
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
       REPRESU& = XVALO(Mid$(List2.List(KKI% - 1), 121, 8))
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
   If InStr(EPAC$, "CONTEPLAST") > 0 And SeleccionoPresupuesto% > 0 Then
      Call CARDETPEDI(1, 1) 'SI ES CONTEPLAST Y SE SELECCIONO PRESUPUESTO ENVIA 1 PARA QUE EL CARDETPE SE CARGUE A MODO NO EDITABLE
   Else
      Call CARDETPEDI(1, 0) ' SI NO MANTIENE OPCION POR DEFECTO
   End If
   '
End Sub

Private Sub Command2_Click()
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        If InStr(EPAC$, "ENVALAST") > 0 Then
            On Error Resume Next
            BASEWALD.Close
            On Error GoTo 0
        End If
        
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
     Dim Rst As ADODB.Recordset
     Set Rst = FLEXCONN.Execute(SQLX$)
     CONDI$ = ""
     With Rst
        Do While Not .EOF
            CONDI$ = CONDI$ & !column_name & " like '%" & ARGU$ & "%' or "
            .MoveNext
        Loop
     End With
     On Error Resume Next
     Rst.Close
     Set Rst = Nothing
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
    
    EPAC$ = TRIM$(UCase$(EMPREAC$)) '
    If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
      If DerechoSupervisor% < 1 Then
        If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then
          GoTo 99
        End If
      End If
    End If
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
99  Command22.Enabled = True
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
    ta$ = "CAJABANC": CA$ = "NUCliEN": GoSub 100
    'TA$ = "ENCAREMI": CA$ = "Nume_Clie": GoSub 100
    ta$ = "APLICOBRA": CA$ = "NuClien": GoSub 100
    'TA$ = "SUCENTRE": CA$ = "Nume_Cli": GoSub 100
    ta$ = "SACREPEN": CA$ = "NuClien": GoSub 100
    ta$ = "SADEUPEN": CA$ = "NuClien": GoSub 100
    'TA$ = "REPARA": CA$ = "Nume_Cli": GoSub 100
    'TA$ = "ORDETRAB": CA$ = "NCLIE_OTRAB": GoSub 100
    ta$ = "MOVIREPA": CA$ = "Nume_CliE": GoSub 100
    ta$ = "MOVISTO": CA$ = "Nume_CliE": GoSub 100
    'TA$ = "MOVIVIE": CA$ = "Nume_CliE": GoSub 100
    ta$ = "PEDENCA": CA$ = "NroCliE": GoSub 100
    ta$ = "PEDDETA": CA$ = "NroCliE": GoSub 100
    ta$ = "PRESUVEN": CA$ = "NroCliE": GoSub 100
    ta$ = "DEUDOR12": CA$ = "Nume_Cli": GoSub 100
    '
    JJ& = 0
    For i& = 1 To 32767
      If CLACT%(i&) < 1 Then
        JJ& = JJ& + 1
      End If
    Next i&
    
    MsgBox "Lugares Libres " & JJ&
Stop

    
100 Dim CLIMO As ADODB.Recordset
    Set CLIMO = New ADODB.Recordset
    SQLX$ = "SELECT " + CA$ + " AS NCX FROM " + ta$ + " WHERE " + CA$ + " > 0"
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

Private Sub Command29_Click()
    If Command29.Caption = "mas datos" Then
        Command29.Caption = "menos datos"
    Else
        Command29.Caption = "mas datos"
    End If
        
    Call Image1_Click
    
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP01")
    Command3.Enabled = True
End Sub

Private Sub Command30_Click()
   MSF.Row = FilaActiva
   MSF.COL = J&
   '
  
   Codigo$ = MSF.TextMatrix(FilaActiva, 2)
   DESCRIX$ = MSF.TextMatrix(FilaActiva, 3)
   '
   
   OBSERVAC.Caption = "Observaciones del Item"
   OBSERVAC.Label24.Caption = Codigo$ & " - - " & DESCRIX$
   OBSERVAC.Label2 = 256
   OBSERVAC.Text28.MaxLength = 256
   OBSERVAC.Text28 = MSF.TextMatrix(FilaActiva, 9)

   OBSERVAC.Show 1
   '
   If OBSERVAC.Label1 = "OK" Then
     Text28 = OBSERVAC.Text28
     MSF.TextMatrix(FilaActiva, 9) = Text28
   End If




    'call proceso de validacion de simetria de facturas debe vs haber.
    Command31.Enabled = False
    VENTANA.Inicializar = 0
    VENTANA.Formula(4) = "2;-;3"
    VENTANA.TOTALES(4) = "S"
    VENTANA.AgregaRegistro = 0
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = ""
    VENTANA.Inicializar = 1
    VENTANA.Campox = "CODICOMLAR/A24;IDEBECOMP/F12.2;IHABECOMP/F12.2;DIFERENCIA/F12.2"
    
    
    '****************************************************************
    
      Dim obj As New RECORXET
      Set rs1 = obj.RS_VALIDAR()
      Call Carga_MSF_Recordset(rs1, "CODICOMLAR/A24;IDEBECOMP/F12.2;IHABECOMP/F12.2;DIFERENCIA/F12.2", VENTABNEW.MSF)
      Set obj = Nothing
      rs1.Close
      Set rs1 = Nothing
    '*****************************************************'
      VENTABNEW.Show 1
      Set VENTANA = Nothing
  
99 Command31.Enabled = True

End Sub




Private Sub Command4_Click()
    Command4.Enabled = False
    
    EPAC$ = TRIM$(UCase$(EMPREAC$)) '
    If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
      If DerechoSupervisor% < 1 Then
        If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then
          GoTo 99
        End If
      End If
    End If
    
    RECUEC07.Text1.TEXT = Text1.TEXT
    RECUEC07.Show 1
99  Command4.Enabled = True
End Sub
Sub Command5_Click()
    Command5.Enabled = False
    Call CIERRARCH("*.*")
    If (USNBR% Mod 100) = 1 Then
        CFGPED07.Show 1
    ElseIf DERACCE%("CONFPEDI", "Configuración de Pedidos") >= 1 Then
        CFGPED07.Show 1
    End If
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   MARCONOTA.Visible = False
   
   If Control("ADMIPED", "OBSXITEM") = "SI" Then
     Text28.Visible = True
     MSF.Height = 2400
     Text28.ZOrder 0
   Else
     Text28.Visible = False
     MSF.Height = 3060
   End If
   
   Call PEDIPEND
   Call SELPRESUPEN
End Sub
'
Sub SELPRESUPEN()
On Error GoTo ERROR_TRAZAR
   EPAC$ = UCase$(TRIM$(EMPREAC$))
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
   'TOLERANCIA CONFIGURADA EN SETUP DE PRESUPUESTO
   TOLERANCIAPRESUP% = XVALO(Control("ADMIPED", "VENCPRES"))
   SQLX$ = "SELECT * FROM PRESUVEN WITH(NOLOCK) "
   SQLX$ = SQLX$ + "WHERE NroClie = " & NCLIEN & " "
   SQLX$ = SQLX$ + "AND NroPedi < 1 "
   If TOLERANCIAPRESUP% > 0 Then
      'SI TIENE CARGADO DIAS, SOLO MIRA CONSIDERANDO LA TOLERANCIA ES DECIR SI TIENE 15
      'CONSIDERA LOS PRESUPUESTOS DE LOS ULTIMOS 15 DIAS.
      FE% = DIANTE(HOY(HO$), TOLERANCIAPRESUP%)
      SQLX$ = SQLX$ + " AND FECHEMIS >= '" & FETEXCOR1$(FE%) & "'"
   End If
   SQLX$ = SQLX$ + " ORDER BY NroPres DESC"
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
           NRO& = !NROPRES
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
           CAPRESU& = !CANTIT
           If CAPRESU& < 1 Then CAPRESU& = 1
                 
           If InStr(EPAC$, "CONTEPLAST") > 0 Then
              'esto es para poder  identificar el codigo al grabar va a estar entre -| y |-
              TTXX$ = TTXX$ + " -| " + TRIM$(CODEXT$(CIXI%)) + " |-" + " (" + TRIM$(Str$(CAPRESU&)) + ") " + DESCRIT0$(CIXI%)
 
           Else
              TTXX$ = TTXX$ + " - " + TRIM$(CODEXT$(CIXI%)) + " (" + TRIM$(Str$(CAPRESU&)) + ") " + DESCRIT0$(CIXI%)
           End If
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
      If InStr(EPAC$, "CONTEPLAST") > 0 Then
         If DERACCE%("GEPESICO", "Generación de Pedidos Sin Cotización Previa") < 2 Then
           Call COMUNI("Imposible Realizar Pedidos Sin Cotización")
           Call FINAL("")
         End If
      End If
      Call SETULTREG("!CARDETPE", 0)
      Call CIERRARCH("!CARDETPE")
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
   Exit Sub
ERROR_TRAZAR:
  MsgBox MENSAX$ + "ERROR DENTRO DE SELPRESUPEN  " & " " & Err.Description
  Call FINAL("")
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
   If DERACCE%("DESCPESU", "Descargas de Pedidos de Sucursal") < 2 Then
          Exit Sub
   End If
   EPAC$ = UCase$(TRIM$(EMPREAC$))
   Call CREAR_TABLA_OC_DE_SUCURSAL
   Call CREAR_TABLA_MENSAJES_DE_SUCURSAL
   '
   'ENJUAGUE PARA QUE ENTRE DIRECTAMENTE SI ES ARCON
   If InStr(EPAC$, "METALES ARCON") > 0 Then
      Ruta$ = RutaCarpetaLocalCompraInterna$ 'SI LA RUTA CONFIGURADA EN TABLA AUXILIAR (CTEPROCOMPRA) EXISTE
      Call DESCARGAR_ARCHIVO_COMPRAS_SUCURSAL("Ocompra_Sucursal") '
      Exit Sub
   End If
   
   
   
   EPAC$ = UCase$(TRIM$(EMPREAC$))
   If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "QUILMES") > 0 Then
       Call DESCARGAR_ARCHIVO_COMPRAS_SUCURSAL_COMPRAS_INTERNAS("Ocompra_Sucursal") 'modalidad conteplast
   Else
      'ESTO SOLO APLICA PARA DESCARGAR LOS ARCHIVOS DE LA ORDEN DE COMPRA GUARDADA EN EL DROP HECHA POR UNA SUCURSAL
      Ruta$ = RutaCarpetaLocalCompraInterna$ 'SI LA RUTA CONFIGURADA EN TABLA AUXILIAR (CTEPROCOMPRA) EXISTE
      If Dir(Ruta$, vbDirectory) <> "" Then  'ENTONCES BUSCA LA RUTA DEL MODO CONTEPLAST POR AHORA SOLO LA  RUTA
        'If (GetAttr(Ruta$) And vbDirectory) = vbDirectory And InStr(UCase$(EMPREAC$), "METALES ARCON") < 1 Then
        If (GetAttr(Ruta$) And vbDirectory) = vbDirectory Then  'QUE FIGURA COMO LOCAL
            Call DESCARGAR_ARCHIVO_COMPRAS_SUCURSAL_COMPRAS_INTERNAS("Ocompra_Sucursal") 'modalidad conteplast AHORA TAMBIEN COMO SE CONFIGURO COMO LOCAL EN CTEPROCOMPRA ENTRA ACA
            'SOLO SE CONFIGURO EN TABLA AUXILIAR CON REFERENCIA LOCAL Y LA RUTA DEL DROP "\\server\Users\Administrador\Dropbox\NEUQUEN"
            'ESTO SIRVE EN ARCON POR QUE ES SOLO 1 (AVELLANEDA) QUE RECIBE PEDIDOS DE REPOSICION
            'POR QUE SI NO LA T.AUXILIAR HABRIA QUE CARGARLA COMPLETA COMO EN CONTEPLAST.
        Else 'SI ES METALES ARCON ENTRA ACA
           'SI SE QUITA LA CONFIGURACION EN TABLA AUXILIAR (CTEPROCOMPRA) VUELVE A INGRESAR AQUI COMO ANTES.
           Call DESCARGAR_ARCHIVO_COMPRAS_SUCURSAL("Ocompra_Sucursal") 'MODALIDAD TRADICIONAL DE ARCON
       End If
     End If
   End If
   
End Sub
Sub DESCARGAR_ARCHIVO_COMPRAS_SUCURSAL(Archivo$)

        CLIENTE11 = Control("RETRASLA", "SUC1CLIE")
        CLIENTE22 = Control("RETRASLA", "SUC2CLIE")
        If XVALO(CLIENTE11) + XVALO(CLIENTE22) < 1 Then
           Call COMUNI("Falta Configurar Rutas de Descarga de Archivos de Sucursal")
           Exit Sub
        End If
        
        OPX% = COMALTER%("Seleccione Sucursal\\" & Left$(TRIM$(rasocli$(XVALO(CLIENTE11))), 16) & "\" & Left$(TRIM$(rasocli$(XVALO(CLIENTE22))), 16) & "")
        If OPX% < 1 Or OPX% > 2 Then Exit Sub
        If OPX% = 1 Then CLIENTESEL = XVALO(CLIENTE11): Ruta$ = Control("RETRASLA", "SUC1RUTA")
        If OPX% = 2 Then CLIENTESEL = XVALO(CLIENTE22): Ruta$ = Control("RETRASLA", "SUC2RUTA")
        'RUTA$ = TRIM$(CONTROL("", "CARPEINT"))
        Ruta$ = RutaCarpetaLocalCompraInterna$
        MsgBox Ruta$
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
                  NombreArchivo$ = TRIM$(Left$(File1.List(n), Len(File1.List(n)) - 4))
                  FechaArchivo$ = TRIM$(FileDateTime(Ruta$ + "\" + NombreArchivo$ + ".txt"))
                  Call REPLA(TX$, AJUSTI$(FechaArchivo$ + " | ", 48), 1, 48)
                  Call REPLA(TX$, AJUSTI$(NombreArchivo$, 48), 50, 48)
                  List3.AddItem TX$
                Next n
                '
                'SEGUN EL ARCHIVO TOMA LA TABLA CORRESPONDIENTE
                If InStr(Archivo$, "Ocompra_Sucursal") > 0 Then TABLAX$ = "OC_SUCURSAL"
                '
                'GRABA CADA ARCHIVO EN SQL Y SI SI GRABA CORRECTAMENTE SE ELIMINA.
                For i& = 0 To List3.ListCount - 1
                  POS1% = InStr(TX$, "|")
                  FechaArchivo$ = Mid$(List3.List(i&), 1, POS1% - 2)
                  NombreArchivo$ = TRIM$(Mid$(List3.List(i&), POS1% + 2))
                  If UCase("Ocompra_Sucursal") = Left$(TRIM$(UCase$(NombreArchivo$)), Len(Archivo$)) Then
                    TABLAX$ = "OC_SUCURSAL"
                    MsgBox "previo a grabar archivo"
                    Call GrabarArchivo(Ruta$ + "\" + NombreArchivo$ + ".txt", TABLAX$) 'GRABA EN LA BASE DE DATOS SQL
                  End If
                Next i&

        Else
            Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio")
            Exit Sub
        End If
        Call COMUNI("Proceso Terminado")
        '
End Sub
Sub DESCARGAR_ARCHIVO_COMPRAS_SUCURSAL_COMPRAS_INTERNAS(Archivo$)

'        CLIENTE11 = CONTROL("RETRASLA", "SUC1CLIE")
'        CLIENTE22 = CONTROL("RETRASLA", "SUC2CLIE")
'        If XVALO(CLIENTE11) + XVALO(CLIENTE22) < 1 Then
'           Call COMUNI("Falta Configurar Rutas de Descarga de Archivos de Sucursal")
'           Exit Sub
'        End If
         
        Ruta$ = RutaCarpetaLocalCompraInterna$

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
                  NombreArchivo$ = TRIM$(Left$(File1.List(n), Len(File1.List(n)) - 4))
                  FechaArchivo$ = TRIM$(FileDateTime(Ruta$ + "\" + NombreArchivo$ + ".txt"))
                  Call REPLA(TX$, AJUSTI$(FechaArchivo$ + " | ", 48), 1, 48)
                  Call REPLA(TX$, AJUSTI$(NombreArchivo$, 48), 50, 48)
                  List3.AddItem TX$
                Next n
                '
                'SEGUN EL ARCHIVO TOMA LA TABLA CORRESPONDIENTE
                If InStr(Archivo$, "Ocompra_Sucursal") > 0 Then TABLAX$ = "OC_SUCURSAL"
                
                '
                'GRABA CADA ARCHIVO EN SQL Y SI SI GRABA CORRECTAMENTE SE ELIMINA.
                For i& = 0 To List3.ListCount - 1
                  POS1% = InStr(TX$, "|")
                  FechaArchivo$ = Mid$(List3.List(i&), 1, POS1% - 2)
                  NombreArchivo$ = TRIM$(Mid$(List3.List(i&), POS1% + 2))
                  If UCase("Ocompra_Sucursal") = Left$(TRIM$(UCase$(NombreArchivo$)), Len(Archivo$)) Then
                    TABLAX$ = "OC_SUCURSAL"
                    Call GrabarArchivo(Ruta$ + "\" + NombreArchivo$ + ".txt", TABLAX$) 'GRABA EN LA BASE DE DATOS SQL
                  End If
                Next i&

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

Private Sub ConWaldbott_Click()
   If DERACCE%("MODRUTWA", "Acceso a Ruta de B.Datos Waldbott") >= 2 Then
     RUTAWALD.Show 1
   End If
End Sub
'
Private Sub Form_Load()
    '

    Frame4.Height = 1560
    Openforms = DoEvents
    VERANTER$ = "PEDIDO01"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
    If TERMINA% > 0 Then Call FINAL("")
    '
    CMO$ = TRIM$(Control$("", "CODPECLI"))
    If CMO$ <> "" Then Label10.Caption = CMO$

    Call TRANSLABELS(Name)
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    DoEvents
    '
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
    If InStr(EPAC$, "MINEPARTS") > 0 Then
      UPDATMASTER% = 1
      Call GENECOMAS
      UPDATMASTER% = 0
    End If
    mnuBuscaModelo.Visible = False
    mnuPedEmbalar.Visible = False
    mnuBuscaModelo.Visible = False
    mnuRegDespXLectora.Visible = False
    mnuRegDespXLectora.Visible = False
    If InStr(EPAC$, "AHP") > 0 Then
       Frame3.Visible = True
       Check3.Visible = True
       mnuRegDespXLectora.Visible = True
       mnuRegDespXLectora.Visible = True
    End If
    If InStr(EPAC$, "PREMOTEC") > 0 Then
      Proyepedidos.Visible = True
      mnuBuscaModelo.Visible = True
      mnuPedEmbalar.Visible = True
      mnuBuscaModelo.Visible = True
    End If
    '
    If InStr(UCase$(EMPREAC$), "ENVAPLAST") > 0 Then
        mnuNuevoItem.Visible = True
        mnuVisPlantform.Visible = True
        mnupresuppendientes.Visible = True
        mnuPedidoSinOt.Visible = True
        mnuItCPedPenSinAprobar.Visible = True
        mnuPedPendStOf.Visible = True
        'mnuSincClientes.Visible = True
        RutaEnvap$ = Control("CONWALDB", "RUTAWALD")
        mnupresuppendientes.Visible = True
        ConWaldbott.Visible = True
        mnuSncrnzCPWlbt.Visible = True
        Call EnvaplastAgregaItemsAlista1(1) ' INTEGRA TODOS LOS ITEMS A LA LISTA 1
        If RutaEnvap$ <> "" Then
'           On Error Resume Next
'           Call abreBaDaWald
'           On Error GoTo 0
        End If
    End If
    
    If InStr(UCase$(EMPREAC$), "CONTEPLAST") > 0 Then
'        mnuVisPlantform.Visible = True
        mnupresuppendientes.Visible = True
        mnuPedidoSinOt.Visible = True
        mnuItCPedPenSinAprobar.Visible = True
        mnuPedPendStOf.Visible = True
        ConWaldbott.Visible = True
        If DERACCE%("MODIDESC/NOMESS", "MODIFICAR DESCUENTO EN PANTALLA") < 2 Then
           TEXT21(1).Enabled = False
           TEXT21(2).Enabled = False
           Command21(1).Enabled = False
           Command21(2).Enabled = False
        End If
        If DERACCE%("VECLICTP/NOMESS", "VISUALIZAR TODOS LOS CLIENTES") < 2 Then
           Text1.Enabled = False
        End If
        TEXT21(4).Visible = True
        Command21(4).Visible = True
        Eco21(4).Visible = True
        Label21(7).Caption = "Vend.Envap."
        Label21(7).ToolTipText = "Selecciòn Vendedor Envaplast (Modalidad Comisiòn)"
        '
        TEXT21(4).Top = 2010
        TEXT21(4).Left = 1155
        Command21(4).Top = 2010
        Command21(4).Left = 1575
        Frame6.Visible = False
        Eco21(4).Top = 1980
        Eco21(4).Left = 1890
        
'        'mnuSincClientes.Visible = True
'        RutaEnvap$ = Control("CONWALDB", "RUTAWALD")
'        mnupresuppendientes.Visible = True
'        ConWaldbott.Visible = True
'        mnuSncrnzCPWlbt.Visible = True
'        Call EnvaplastAgregaItemsAlista1(1) ' INTEGRA TODOS LOS ITEMS A LA LISTA 1
'        If RutaEnvap$ <> "" Then
''           On Error Resume Next
''           Call abreBaDaWald
''           On Error GoTo 0
'        End If
    End If
    If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Then
       mnuPedidosAprobVsEntregas.Visible = True
    End If
    '
    Call GENERAR_ESTRUCTURAS_SQL
    Call GENERAR_TABLA_VERSION_IMAGEN
    Label11.Visible = True
    'If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "MEDITEA") > 0 Or InStr(EPAC$, "GABAL") > 0 Or InStr(EPAC$, "ARCON") > 0 Or InStr(EPAC$, "FERVI") > 0 Or InStr(EPAC$, "TANIT") Or InStr(EPAC$, "BICI") > 0 Or InStr(EPAC$, "AMBEST") > 0 Or InStr(EPAC$, "EYM") > 0 Or InStr(EPAC$, "AHP") > 0 Then
        Command22.Visible = True
   ' End If
    If InStr(EPAC$, "ARCON") > 0 Then
       pedidossuc.Visible = True
    End If
    
    If Control$("ADMIPED", "MOCOMIS") = "LETRA" Then
       Label21(4) = "Letra"
    End If
    '
    Campos$ = "Cant./F6;Código/A16;Descripción/A22;Fecha/D5;Lista/F10.2;Desc/F5.2;P.Unit/F10.2;Importe/F12.2;OBSERVACIONES/A0"
    Call CARGA_ENCABEZADO(Me.MSF, Campos$, Me)
    
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
    ORIPED$ = TRIM$(UCase$(Control$("PEDIDO", "ORIDATOS")))
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
      If Control$("CAPRINT", IDENTER$) = "SI" Then
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
    'Image1.Visible = True
    TICAMBIS = FORMATNUM$(TICAMONE(1), "F9.4")
    '
    ECOPRECIO% = 0
    If Control$("ADMIPED", "MUEPREC") <> "SI" Then
      For KKI% = 0 To 2      ' TAPA IMPORTES TOTALES
        Text14(KKI%).Visible = True
      Next KKI%
      ECOPRECIO% = (-1)
    End If
    '
    If ECOPRECIO% < 0 Then
      Label22.Caption = "  Cant Código          Descripción                                           Entrega "
      Image1.Visible = False
      Command29.Visible = False
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
    strSQL = "SELECT * FROM PEDDETA "
    strSQL = strSQL & " WHERE sTATUS >=100 "
    Dim PEDDETTEMP As ADODB.Recordset
    Set PEDDETTEMP = New ADODB.Recordset
27  On Error Resume Next
    PEDDETTEMP.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
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
    If Control("ADMIPED", "ITEMSCGO") = "SI" Then
        Call CargaTablaUsuariosConDerechoParaEntregarMaterialSinCargo
    End If
    '
    Call CREATABLA_EQUIPOS
    
End Sub

Sub CargaTablaUsuariosConDerechoParaEntregarMaterialSinCargo()

    Call BLOQARCH("FLEXUSR")
    FUSER$ = LOADFILE$("USER.DAT")
    JJI& = 0: KKI% = 0
    IDLIST& = ObtenerIdTablaAuxiliar("USUASINCAR")
    For i% = 1 To Len(FUSER$) Step 64
       F1$ = Mid$(FUSER$, i%, 64)
       If Asc(F1$) = 0 Then GoTo 10
       NOM$ = Left$(F1$, 30)
       If Left$(NOM$, 1) = "@" Then NOM$ = DESENCRI$(Mid$(F1$, 2, 29))
       '
       KKI% = KKI% + 1
       'CondicionSql$ = "codiempr=" & CStr(CodiEmpr%) & " and id_lista=" & CStr(IDLIST&) & " and campo1='" & CStr(KKI%) & "'"
       CondicionSql$ = "id_lista=" & CStr(IDLIST&) & " and campo1='" & CStr(KKI%) & "'"
       If CantRegistros&("DATASQL", CondicionSql$) = 0 Then
            ConsultaSql$ = "insert into datasql (codiempr, id_lista, lista, campo1, campo2, campo3) values (" _
            & CStr(CODIEMPR%) & "," & CStr(IDLIST&) & "," & "'USUASINCAR'" & "," & CStr(KKI%) & ",'" & TRIM$(NOM$) & "'," & "'')"
            FLEXCONN.Execute ConsultaSql$
       Else
            'CondicionSql$ = "codiempr=" & CStr(CodiEmpr%) & " and id_lista=" & CStr(IDLIST&) & " and campo1='" & CStr(KKI%) & "'"
            CondicionSql$ = "id_lista=" & CStr(IDLIST&) & " and campo1='" & CStr(KKI%) & "'"
            ConsultaSql$ = "update datasql set campo2='" & TRIM$(NOM$) & "' where " & CondicionSql$
            FLEXCONN.Execute ConsultaSql$
       End If
       
10  Next i%
    Call LIBARCH("FLEXUSR")
End Sub

Private Sub Frame3_DblClick()
   Entrega.Show 1
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
        i& = i& + 1
        FRMZELECHO.msf2.Rows = i& + 1
        FRMZELECHO.msf2.TextMatrix(i&, 1) = SAFETEXT$(!VALOR1)
        If Index <= 1 Then
          FRMZELECHO.msf2.TextMatrix(i&, 1) = SAFETEXT$(!VALOR1) & " " & SAFETEXT$(!VALOR2)
          If TRIM$(SAFETEXT$(!VALOR2)) <> "" Then
             FRMZELECHO.msf2.TextMatrix(i&, 2) = SAFETEXT$(!VALOR1) & "|" & SAFETEXT$(!VALOR2)
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

Private Sub menuCumpPP_Click()


 Exit Sub
 
 
'LOS NROS DE PEDIDO SON LOS QUE PASO VICTORIA COMO PENDIENTES
'ENTONCES EXCEPTUA TODOS LOS PENDIENTES Y ADEMAS COMO CONDICION QUE SEAN MENORES AL NRO MAYOR DE PEDIDO
SQLX$ = "Update PEDDETA set Status=8, tcolor = 'A M -1'"
SQLX$ = SQLX$ + " Where CanPed > CantEnt"
SQLX$ = SQLX$ + " AND NROPED NOT IN(7711,7817,7961,8013,8035,8039,8040,8041,8042,8149,8150,8319,8352,8365,8384,8452,8481,8510,8511,8512,8513,8520,8581,8595,8611,8623,8625,8630,8632,8636,8638,8641,8647,8652,8653,8669,8670,8671,"
SQLX$ = SQLX$ + " 8697,8700,8702,8703,8704,8705,8706,8707,8711,8713,8722,8723,8724,8731,8732,8733,8734,8736,8759,8757,8760,8762,8774,8768,8771,8772,8773,8775,8781,8778,8786,8787,8788,8789,8790,8791,8794,8793,"
SQLX$ = SQLX$ + " 8796,8798,8803,8804,8805,8806,8808,8811,8814,8815,8816,8818,8817,8819,8820,8821,8823,8824,8825,8829,8822,8830,8831,8832,8833,8835,8834,8836,8838,8839,8842,8844,8845,8846,8843,8841,8847,8848,"
SQLX$ = SQLX$ + " 8849,8850,8852,8853,8854,8855,8856,8857,8858,8859,8860,8861,8862,8863,8864,8866,8867,8868,8869,8870,8871,8872,8873,8874,8875,8877,8878,8879,8880,8882,8883,8884,8885,8886,8887,8888,8889,8890,"
SQLX$ = SQLX$ + " 8891,8892,8893,8895,8896,8897,8898,8899,8900,8901,8902,8903,8905,8906)"
SQLX$ = SQLX$ + " AND NROPED < 8906"
SQLX$ = SQLX$ + " AND Status = -1"
FLEXCONN.Execute (SQLX$)
 
 

SQLX$ = "Update PEDDETA set Status=8, tcolor = 'A M 0'"
SQLX$ = SQLX$ + " Where CanPed > CantEnt"
SQLX$ = SQLX$ + " AND NROPED NOT IN(7711,7817,7961,8013,8035,8039,8040,8041,8042,8149,8150,8319,8352,8365,8384,8452,8481,8510,8511,8512,8513,8520,8581,8595,8611,8623,8625,8630,8632,8636,8638,8641,8647,8652,8653,8669,8670,8671,"
SQLX$ = SQLX$ + " 8697,8700,8702,8703,8704,8705,8706,8707,8711,8713,8722,8723,8724,8731,8732,8733,8734,8736,8759,8757,8760,8762,8774,8768,8771,8772,8773,8775,8781,8778,8786,8787,8788,8789,8790,8791,8794,8793,"
SQLX$ = SQLX$ + " 8796,8798,8803,8804,8805,8806,8808,8811,8814,8815,8816,8818,8817,8819,8820,8821,8823,8824,8825,8829,8822,8830,8831,8832,8833,8835,8834,8836,8838,8839,8842,8844,8845,8846,8843,8841,8847,8848,"
SQLX$ = SQLX$ + " 8849,8850,8852,8853,8854,8855,8856,8857,8858,8859,8860,8861,8862,8863,8864,8866,8867,8868,8869,8870,8871,8872,8873,8874,8875,8877,8878,8879,8880,8882,8883,8884,8885,8886,8887,8888,8889,8890,"
SQLX$ = SQLX$ + " 8891,8892,8893,8895,8896,8897,8898,8899,8900,8901,8902,8903,8905,8906)"
SQLX$ = SQLX$ + " AND NROPED < 8906"
SQLX$ = SQLX$ + " AND Status = 0"
FLEXCONN.Execute (SQLX$)
 
SQLX$ = "Update PEDDETA set Status=8, tcolor = 'A M 1'"
SQLX$ = SQLX$ + " Where CanPed > CantEnt"
SQLX$ = SQLX$ + " AND NROPED NOT IN(7711,7817,7961,8013,8035,8039,8040,8041,8042,8149,8150,8319,8352,8365,8384,8452,8481,8510,8511,8512,8513,8520,8581,8595,8611,8623,8625,8630,8632,8636,8638,8641,8647,8652,8653,8669,8670,8671,"
SQLX$ = SQLX$ + " 8697,8700,8702,8703,8704,8705,8706,8707,8711,8713,8722,8723,8724,8731,8732,8733,8734,8736,8759,8757,8760,8762,8774,8768,8771,8772,8773,8775,8781,8778,8786,8787,8788,8789,8790,8791,8794,8793,"
SQLX$ = SQLX$ + " 8796,8798,8803,8804,8805,8806,8808,8811,8814,8815,8816,8818,8817,8819,8820,8821,8823,8824,8825,8829,8822,8830,8831,8832,8833,8835,8834,8836,8838,8839,8842,8844,8845,8846,8843,8841,8847,8848,"
SQLX$ = SQLX$ + " 8849,8850,8852,8853,8854,8855,8856,8857,8858,8859,8860,8861,8862,8863,8864,8866,8867,8868,8869,8870,8871,8872,8873,8874,8875,8877,8878,8879,8880,8882,8883,8884,8885,8886,8887,8888,8889,8890,"
SQLX$ = SQLX$ + " 8891,8892,8893,8895,8896,8897,8898,8899,8900,8901,8902,8903,8905,8906)"
SQLX$ = SQLX$ + " AND NROPED < 8906"
SQLX$ = SQLX$ + "  AND Status = 1"

FLEXCONN.Execute (SQLX$)

End Sub

Private Sub mnuAnotadorPedidos_Click()
    Call Command22_Click
End Sub

Private Sub mnuAprobarPedidos_Click()
    Call OPPED07.Command5_Click
End Sub

Private Sub mnuAsociarReporte_Click()
    Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnuAtributosComercialesPedidos_Click()

    Dim da As New DatosAuxiliares
    
    Call da.MostrarTablaAuxilar("ATRICOPE")

End Sub
Private Sub Archi_Click()
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
    If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then
       'sale solo por que al perder el foco no entra en el menu
    End If
  End If
End Sub

Private Sub mnuEtiqDespacho_Click()
    TituloLista$ = "Pedidos Con Etiqueta de Despacho"
    Campos$ = "Fecha/D8;Pedido/F8;Cta./F7;R.Social/A48;Sucursal/A48"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    TABLA$ = "DatoEntrega DA INNER JOIN DEUDOR12 D WITH(NOLOCK) ON NROCLIE = D.Nume_Cli"
    TABLA$ = TABLA$ + " INNER JOIN PEDENCA P ON P.NroPed = DA.NroPed"
    Campos1$ = " P.FechEmis, DA.NROPED,DA.NroClie,D.Raso_Cli,DENO_SUC "
    Set rs1 = obj.RS_GENERAL(TABLA$, Campos1$, " EntEnDomi < 1 ORDER BY DA.NROPED DESC")

    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 12000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(2)) > 0 Then
       NROPED& = XVALO(RESP_ZELE_VECT(2))
    End If
    If NROPED& < 1 Then Exit Sub
    PrintEtiDes.NroPedido = NROPED&
    PrintEtiDes.Show 1
End Sub

Private Sub mnuPedidosAprobVsEntregas_Click()
   Call ReportePedidosAprobEntregas
End Sub

Private Sub mnuRegDespXLectora_Click()
DespaCbrra.Show 1
End Sub

Private Sub mnuRepVtasEstadistica_Click()
   Call ReportePedidosTodosXRangoFecha
End Sub

Private Sub mnuRepVtasEstadisticaItemAnulCentrega_Click()
  Call ReportePedidosItemCumpCEntregas
End Sub


Private Sub pedidossuc_Click()
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
    If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then
       'sale solo por que al perder el foco no entra en el menu
    End If
  End If
End Sub
Private Sub ConLis_Click()
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
    If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then
       'sale solo por que al perder el foco no entra en el menu
    End If
  End If
End Sub
Private Sub mnuver_Click()
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
    If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then
       'sale solo por que al perder el foco no entra en el menu
    End If
  End If
End Sub
Private Sub mnuUtilidades_Click()
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
    If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then
       'sale solo por que al perder el foco no entra en el menu
    End If
  End If
End Sub
Private Sub Configura_Click()
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
    If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then
              'sale solo por que al perder el foco no entra en el menu
    End If
  End If
End Sub
Private Sub AccDirec_Click()
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
    If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then
       'sale solo por que al perder el foco no entra en el menu
    End If
  End If
End Sub

Private Sub mnuBuscaModelo_Click()
'-Buscador de Modelos: NUEVA PLANILLA:
'a)  Esta planilla es un buscador de stock y contendrá 2 casilleros
'CANTIDAD DE M2              FECHA.
'B ) Resultado:  Que códigos voy a tener disponibles y secos según la fecha ingresada y los m2.
'c) CANTIDAD DE M2: en este casillero se ingresan los m2 y la planilla muestra solo los articulos cuyo stock disponible iguala o supera la cantidad pedida por el cliente.
'd) FECHA: En este casillero figurara por defecto la fecha del día pero podrá ingresarse otras fechas y la planilla mostrara los artículos que están fabricados pero en etapa de secado cuyas cantidades igualen o superen a la cantidad pedida por el cliente (Ingresada en el primer casillero).
'e) AGREGAR FILTRO POR FAMILIA. (Si no se selecciona ninguna familia no discriminara por este

    'ingreso de datos de cantidad y fecha
10  VDEF$ = MKS$(0) + MKI$(HOY(HO$)) + Chr$(0)
15    OBX$ = OBJPLA$("FECHAM2", VDEF$)
    If OBX$ = "" Then
      Exit Sub
    End If
    '
    CANTI = CVS(Left$(OBX$, 4))
    FE% = CVI(Mid$(OBX$, 5, 2))
    
    If CANTI < 1 Or FE% < 33 Then
      Call MENSERR(24, "Datos no Válidos")
      GoTo 15
    End If
    '
'    MOVISTO WHERE CANTIDAD INGRE - CANTSALID > CANTI
'    DESPUES LEER EL STOCK DE ESTAS Y DESCONTAR LA CANTIDAD COMPROMETIDA POR PEDIDOS
'    DESCUENTO LA CANTIDAD EN PROCESO DE SECADO
'    SI LA CANTIDAD SUPERA A CANTI MOSTRAR EN LA GRILLA.
     Campos$ = "CODIGOINTERNO/A0;Código/A20;Descripción/A30;C.Disponible/F12.1"
     Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
     Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
     Screen.MousePointer = 11
     
     SQLX$ = "SELECT COD_INTE,SUM(CANTINGRE)AS INGRE , sum(CANTSALID) AS SALI FROM MOVISTO WITH(NOLOCK) GROUP BY COD_INTE "
     'SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP"
     SQLX$ = SQLX$ + " HAVING SUM(CANTINGRE) - SUM(CANTSALID) > " & CANTI
     SQLX$ = SQLX$ + " ORDER BY COD_INTE"
     Set RS = READSET(SQLX$)
     With RS
        Do While Not .EOF
          CI1% = XVALO(!cod_inte)
          ST = XVALO(!INGRE) - XVALO(!sali)
          CANPED_PEND = CANTIDAD_PEDIDA_PENDIENTE(CI1%)
          T_SECADOX% = XVALO(COLOR$(CI1%))
          EnProcesoDeSecado = CANTI_ENPROCESO_SECADO(CI1%, T_SECADOX%, FE%)
          ST_DISPONIBLE = ST - (CANPED_PEND + EnProcesoDeSecado)
          If ST_DISPONIBLE >= CANTI Then
            J& = J& + 1
            'AGREGO A LA GRILLA
            GRILLAGRAL.GRID.Rows = J& + 1
            GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI1%
            GRILLAGRAL.GRID.TextMatrix(J&, 2) = CODEXT$(CI1%)
            GRILLAGRAL.GRID.TextMatrix(J&, 3) = DESCRIT0$(CI1%)
            GRILLAGRAL.GRID.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(ST_DISPONIBLE, "F12.1"))
          End If
          .MoveNext
        Loop
     End With
     RS.Close
     Set RS = Nothing
    
     GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
     GRILLAGRAL.mnuMenu1.Visible = True
     GRILLAGRAL.mnuMenu2.Visible = False
     GRILLAGRAL.mnuMenu3.Visible = False

     GRILLAGRAL.GRID.ZOrder 0
     GRILLAGRAL.TXTBUSCAR = ""
     TX$ = "Modelos Con Stock Disponible Superior a la Cantidad: " & TRIM$(FORMATNUM$(CANTI, "F12.1"))
     'Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
     Screen.MousePointer = 1
     If J& > 0 Then
        GRILLAGRAL.Show 1
     Else
        Call COMUNI("No se Detectaron Items Que Cumplan Con la Cantidad Solicitada")
        GoTo 99
     End If
     '
     Dim MSF As msFlexGrid
     Set MSF = GRILLAGRAL.GRID
    
     If GRILLAGRAL.MENU1 > 0 Then
         Screen.MousePointer = 11
         If TRIM$(GRILLAGRAL.TXTBUSCAR) <> "" Then MSF = GRILLAGRAL.MSF_2
         Call CARGA_TABLA_IMPRE(MSF, Campos$, "BUSCMODE", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
         Call FINAL("?" & RET_NAMECONS$)
         'Call EliminarSQF(RET_NAMECONS$)
         GRILLAGRAL.MENU1 = 0
         Screen.MousePointer = 1
     End If
     Set MSF = Nothing
99   Screen.MousePointer = 1
End Sub

Private Sub mnuControPedidos_Click()
   CONTROLPEDI.Show 1
End Sub

Private Sub mnuEstadisticaDeVentasyPedido_Click()
   OPPED07.Command12_Click
End Sub

Private Sub mnuHistoricoPreciosPedido_Click()
    Call OPPED07.Command10_Click
End Sub

Private Sub mnuImportadorPedidos_Click()
  Call Image6_DblClick
End Sub

'
Private Sub mnuItCPedPenSinAprobar_Click()
    Screen.MousePointer = 11
    Dim obj As New RECORXET
    Campos$ = "Cta/F6;R.Social/A24;F.Ped./D8;Nro.Ped./F7;Código/A16;Descripición/A24;Cantidad/f12;F.Sol./D8;Ord.Compra/A16"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
10  FRMZELECHO.MSF1.Rows = 1
    Dim APROBADO1%()
    ReDim Preserve APROBADO1%(1)
    Set RS11 = obj.RS_PEDIPEN2
    With RS11
      Do While Not .EOF
          A = SAFETEXT$(!CODIEXT)
          CIXI% = CODINT%(SAFETEXT$(!CODIEXT))
          If UBound(APROBADO1%) < CIXI% Then ReDim Preserve APROBADO1%(CIXI%)
          If APROBO$(CIXI%) = "" Then APROBADO1%(CIXI%) = 1
          .MoveNext
      Loop
    End With
    '
    On Error Resume Next
    Set obj = Nothing
    RS11.Close
    Set RS11 = Nothing
    On Error GoTo 0
    '
    Dim obj2 As New RECORXET
    Set rs1 = obj2.RS_PEDIPEN2 ' FUNCION QUE DEVUELVE UN RECORDSET
    i& = 0
    With rs1
      Do While Not .EOF
          CIXI% = CODINT%(SAFETEXT$(!CODIEXT))
          If CIXI% < 1 Then GoTo 30
          If APROBADO1%(CIXI%) < 1 Then GoTo 30 ' NEXT
          NC = XVALO(!NROCLIE)
          RASO$ = rasocli$(NC)
          NPED& = XVALO(!NROPED)
          FEMI% = CVINT(!FECHEMIS)
          DESC$ = DESCRIT0$(CIXI%)
          CantSol = XVALO(!CanPed)
          FentSol% = CVINT(!FeSolEnt)
          i& = i& + 1
          FRMZELECHO.msf2.Rows = i& + 1
          FRMZELECHO.msf2.TextMatrix(i&, 1) = XVALO(!NROCLIE)
          FRMZELECHO.msf2.TextMatrix(i&, 2) = SAFETEXT$(!RasoClie)
          FRMZELECHO.msf2.TextMatrix(i&, 3) = FECHATEX$(FEMI%)
          FRMZELECHO.msf2.TextMatrix(i&, 4) = NPED&
          FRMZELECHO.msf2.TextMatrix(i&, 5) = CODEXT$(CIXI%)
          FRMZELECHO.msf2.TextMatrix(i&, 6) = DESC$
          FRMZELECHO.msf2.TextMatrix(i&, 7) = FORMATNUM$(CantSol, "f10.1")
          FRMZELECHO.msf2.TextMatrix(i&, 8) = FECHATEX$(FentSol%)
          FRMZELECHO.msf2.TextMatrix(i&, 9) = OrdenCompraPed$(NPED&)
30       .MoveNext
      Loop
    End With
    On Error Resume Next
    Set obj2 = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
    
    FRMZELECHO.Width = 14500
    FRMZELECHO.Caption = "Items C/Pedidos Pendientes Sin Aprobar - (Doble Click Sobre Items Para Ver Ficha Técnica y Aprobar)"
    Screen.MousePointer = 1
    FRMZELECHO.Show 1
    '
    CIXI% = CODINT%(SAFETEXT$(RESP_ZELE_VECT(5)))
    If CIXI% < 1 Then Exit Sub
    CODI$ = SAFETEXT$(RESP_ZELE_VECT(5))
    
    FRMENVAP.TXTCODIGO = CODI$
    FRMENVAP.Show 1
End Sub

Private Sub mnuNuevoItem_Click()
       If DERACCE%("ALTAITEM", "Alta de un Nuevo Item de Stock") < 2 Then
         Exit Sub
       End If
       Call FRMENVAP.DARDEALTA

End Sub

Private Sub mnuNuevoPedidoCopia_Click()
    Call OPPED07.Command11_Click
End Sub

Private Sub mnuObservaItemDePedido_Click()
    Call OPPED07.agregarobservaciondeitem_Click
End Sub

Private Sub mnuPedAFab_Click()
        '- Pedidos a fabricar: NUEVA PLANILLA:
        'a)Esta planilla contiene un listado de los pedidos de los clientes cuyas cantidades excedan al
        'stock total. Ej Pedido 40 m2 Adoquín y si el stock total es inferior a 40 m2 entonces este pedido permanecerá en la lista hasta que se genere una orden de fabricación que sumada al stock total igualen o superen las cantidades del pedido, ni bien se cumpla esta condición el pedido desaparecerá de la lista. Esta lista se podrá ordenar por código o fecha de Entrega del pedido.
        'b)El reporte tendrá los siguientes campos:
        'NRO DE PEDIDO , CLIENTE, CODIGO, DESCRIPCION, CANTID. STOCK, CANTIDAD A FABRICAR.
        'c)Se depura en la medida que se vayan cerrando las o.f. o bien con la entrega
        'de los pedidos (movimiento que se registra con la realización del remito).
        '
        '================================================================================================
        'CARGO UN VECTOR CON EL RESULTADO DEL STOCK ACTUAL + LOS SALDOS DE ORDEN DE FABRICACION PENDIENTE
        '================================================================================================
        OPX% = COMALTER%("Opciones de Reporte\\Todos Los Items Pedidos\Solo Items Pedidos Con Estructura")
        Dim TOTAL_STOCK_FABPEND(), STOCKAXT(), FABPEND()
        
        SQLX$ = "SELECT CODIINT FROM PEDDETA P WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE P.CANPED > P.CANTENT  AND P.STATUS < 8 AND P.STATUS >= 0"
        If OPX% = 2 Then
        'solo items con estrucutura y que no tengan proveedor habitual(para no mostrar los que se consignan al proveedor (criterio mrp))
          SQLX$ = SQLX$ + " AND exists (select 1 from FORMULAS F where P.CODIINT  = F.CODICONJ)   "
          SQLX$ = SQLX$ + " AND  NOT exists (select 1 from MASTER M where P.CODIINT  = M.CODINT AND (M.PROVHABI > 0))   "
        End If
        SQLX$ = SQLX$ + " GROUP BY P.CODIINT ORDER BY P.CODIINT"
        Set RS = READSET(SQLX$)
        With RS
            Do While Not .EOF
              CI1% = XVALO(!CODIINT)
              COD$ = CODEXT$(CI1%)
              'lo saco p q si embalo lo descuenta y aparece a fabricar y esta mal
              'YADESPA = CANTI_INFO_DESPA(CI1%)
              ReDim Preserve TOTAL_STOCK_FABPEND(CI1%), STOCKAXT(CI1%), FABPEND(CI1%)
              STOCKAXT(CI1%) = STOCKACT(CI1%)
              FABPEND(CI1%) = XVALO(VALOBADA("ORDEFABR", "SUM(CANPROY - CANAPRO)", " STATUORF < 3  AND cod_inte =  " & CI1%, "NOMESS"))
              TOTAL_STOCK_FABPEND(CI1%) = (STOCKAXT(CI1%) + FABPEND(CI1%)) '- YADESPA
              
              .MoveNext
            Loop
        End With
        RS.Close
        Set RS = Nothing
        'FIN CARGA DE VECTORES
        '===============================================================================================
        '
        '==============================================================================================================
        'CARGO LA INFO EN LA GRILLA ARMO UN RECORDSET DE LOS PEDIDOS PENDIENTES Y LE ASIGNO LOS VALORES DE LOS VECTORES
        '==============================================================================================================
        J& = 0: JJ& = 0
20      Campos$ = "CODIGOINTERNO/A0;N°.Pedido/F8;F.Ent./D8;Cta./F8;R.Social/A24;Código/A20;Descripción/A30;C.Pend./F10.1;Stock/F9.1;A Fabricar/F11.2"
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
        CONDI$ = "CANPED > CANTENT  AND STATUS < 8 AND STATUS >= 0"
        FIN& = CantRegistros&("peddeta", CONDI$)
        SQLX$ = "SELECT P.NROPED, P.NROCLIE, P.RASOCLIE, P.CODIINT, P.CODIEXT , P.DESCRIP,P.FeSolEnt "
        SQLX$ = SQLX$ + " , SUM(P.CANPED - P.CANTENT) AS CANTIPEND FROM PEDDETA P WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE " & CONDI$
        If OPX% = 2 Then
          'solo items con estrucutura y que no tengan proveedor habitual(para no mostrar los que se consignan al proveedor (criterio mrp))
          SQLX$ = SQLX$ + " AND  exists (select 1 from FORMULAS F where P.CODIINT  = F.CODICONJ)"
          SQLX$ = SQLX$ + " AND  NOT exists (select 1 from MASTER M where P.CODIINT  = M.CODINT AND (M.PROVHABI > 0 ))   "
        End If

        'SQLX$ = SQLX$ + " AND CODIINT= 2"
        SQLX$ = SQLX$ + " GROUP BY P.NROPED, P.NROCLIE, P.RASOCLIE, P.CODIINT, P.CODIEXT , P.DESCRIP,P.FeSolEnt"
        SQLX$ = SQLX$ + " ORDER BY P.NROPED, P.CODIINT,P.FeSolEnt  "
        Set RS = READSET(SQLX$)
        With RS
            Do While Not .EOF
                JJ& = JJ& + 1
                Call TRACE(20, JJ&, FIN&)
                CI1% = XVALO(!CODIINT)
                '
                'VOY DESCONTANDO LAS CANTIDADES PEDIDAS
                TOTAL_STOCK_FABPEND(CI1%) = ROUND(TOTAL_STOCK_FABPEND(CI1%) - XVALO(!CANTIPEND), 1)
                '
                If TOTAL_STOCK_FABPEND(CI1%) + XVALO(!CANTIPEND) < XVALO(!CANTIPEND) Then
                    TOTAL_STOCK_FABPEND(CI1%) = TOTAL_STOCK_FABPEND(CI1%) + XVALO(!CANTIPEND)
                    J& = J& + 1
                    GRILLAGRAL.GRID.Rows = J& + 1
                    GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI1%
                    GRILLAGRAL.GRID.TextMatrix(J&, 2) = XVALO(!NROPED)
                    GRILLAGRAL.GRID.TextMatrix(J&, 3) = FECHATEX$(CVINT(!FeSolEnt))
                    GRILLAGRAL.GRID.TextMatrix(J&, 4) = XVALO(!NROCLIE)
                    GRILLAGRAL.GRID.TextMatrix(J&, 5) = SAFETEXT$(!RasoClie)
                    GRILLAGRAL.GRID.TextMatrix(J&, 6) = SAFETEXT$(!CODIEXT)
                    GRILLAGRAL.GRID.TextMatrix(J&, 7) = SAFETEXT$(!Descrip)
                    GRILLAGRAL.GRID.TextMatrix(J&, 8) = FORMATNUM$(XVALO(!CANTIPEND), "F10.1")
                    If XVALO(TOTAL_STOCK_FABPEND(CI1%)) < 0 Then TOTAL_STOCK_FABPEND(CI1%) = 0
                    GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(TOTAL_STOCK_FABPEND(CI1%), "F10.1")
                    GRILLAGRAL.GRID.TextMatrix(J&, 10) = FORMATNUM$(XVALO(!CANTIPEND) - TOTAL_STOCK_FABPEND(CI1%), "F10.1")
                    TOTAL_STOCK_FABPEND(CI1%) = 0 ' LO PONGO EN 0 POR QUE SIEMPRE VA A CUBRIR CON TODO SI NO NO ENTRARIA
                End If
                .MoveNext
            Loop
        End With
        RS.Close
        Set RS = Nothing
        Screen.MousePointer = 1
        GRILLAGRAL.Caption = "Pedidos a Fabricar"
        GRILLAGRAL.mnuMenu2.Visible = False
        GRILLAGRAL.mnuMenu3.Visible = False
        GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
        GRILLAGRAL.mnuMenu1.Visible = True
        GRILLAGRAL.GRID.ZOrder 0
        GRILLAGRAL.TXTBUSCAR = ""
        'Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
        If J& > 0 Then
           GRILLAGRAL.Show 1
        Else
           Call COMUNI("No se Detectaron Pedidos Pendientes Que No Tengan Stocks Suficiente")
           Exit Sub
        End If
        '
        Dim MSF As msFlexGrid
        Set MSF = GRILLAGRAL.GRID

        If GRILLAGRAL.MENU1 > 0 Then
            If TRIM$(GRILLAGRAL.TXTBUSCAR) <> "" Then MSF = GRILLAGRAL.MSF_2
            Call CARGA_TABLA_IMPRE(MSF, Campos$, "PEDIAFAB", "Pedidos de Items a Fabricar", "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
            
        End If
       Set MSF = Nothing

End Sub

Private Sub mnuPedEmbalar_Click()
        '4- Pedidos a embalar: NUEVA PLANILLA (hermana de la anterior):
        'a) Contiene un listado de los pedidos que sí tienen stock entonces se mostraran en una lista que contiene los datos del cliente y las cantidades expresadas en unidades.(posibilidad de imprimir)
        'PODER TILIDAR LAS QUE YA ESTAN PREPARADOS. NO SE DEPURAN POR REMITO, SI NO POR LA MARCA.
        '
        '=================================================================================================
        'CARGO UN VECTOR CON EL RESULTADO DEL STOCK ACTUAL + LOS SALDOS DE ORDEN DE FABRICACION PENDIENTE
        '=================================================================================================
        Dim TOTAL_STOCK_FABPEND(32767), TENGOSTOCK()
        
        SQLX$ = "SELECT CODIINT,NROPED FROM PEDDETA WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE CANPED > CANTENT  AND STATUS < 8 AND STATUS >= 0"
        SQLX$ = SQLX$ + " GROUP BY CODIINT,NROPED ORDER BY CODIINT,NROPED"
        Set RS = READSET(SQLX$)
        With RS
            Do While Not .EOF
              CI1% = XVALO(!CODIINT)
'              NPED& = XVALO(!NROPED)
'              If NPED& > UBound(TENGOSTOCK) Then ReDim Preserve TENGOSTOCK(NPED&)
              'T_SECADOX% = XVALO(COLOR$(CI1%))
              YADESPA = CANTI_INFO_DESPA(CI1%)
              STOCKAXT = STOCKACT(CI1%)
              'FABPEND = XVALO(VALOBADA("ORDEFABR", "SUM(CANPROY - CANAPRO)", " STATUORF < 3 AND Cod_Inte = " & CI1%, "NOMESS"))
              
              TOTAL_STOCK_FABPEND(CI1%) = (STOCKAXT - YADESPA)    '(TOTAL DEL STOCK + FABRICADO SIN CERRAR) - YA INFORMADA COMO DESPACHADA
'              If TOTAL_STOCK_FABPEND(CI1%) > 0 Then
'                 TOTAL_STOCK_FABPEND (CI1%)
'                 TENGOSTOCK(NPED&) = 1
'              End If
              .MoveNext
            Loop
        End With
        RS.Close
        Set RS = Nothing
        'FIN CARGA DE VECTORES
        '===============================================================================================
        '
        '==============================================================================================================
        'CARGO LA INFO EN LA GRILLA ARMO UN RECORDSET DE LOS PEDIDOS PENDIENTES Y LE ASIGNO LOS VALORES DE LOS VECTORES
        '==============================================================================================================
        J& = 0: JJ& = 0
20      Campos$ = "CODIGOINTERNO/A0;N°.Pedido/F8;F.Ent./D8;Cta./F8;R.Social/A24;Código/A20;Descripción/A30;C.Pend./F10.1;Stock/F12.1"
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
        CONDI$ = "CANPED > CANTENT  AND STATUS < 8 AND STATUS >= 0"
        FIN& = CantRegistros&("peddeta", CONDI$)
        SQLX$ = "SELECT NROPED, NROCLIE, RASOCLIE, CODIINT, CODIEXT , DESCRIP,FeSolEnt "
        SQLX$ = SQLX$ + " , SUM(CANPED - CANTENT) AS CANTIPEND FROM PEDDETA WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE " & CONDI$
        SQLX$ = SQLX$ + " AND STADESPA < 1"
        SQLX$ = SQLX$ + " GROUP BY NROPED, NROCLIE, RASOCLIE, CODIINT, CODIEXT , DESCRIP,FeSolEnt"
        SQLX$ = SQLX$ + " ORDER BY  NROPED,FeSolEnt,CODIINT  "
        Set RS = READSET(SQLX$)
        With RS
            Do While Not .EOF
                JJ& = JJ& + 1
                Call TRACE(20, JJ&, FIN&)
                CI1% = XVALO(!CODIINT)
                '
                GRILLAGRAL.GRID.Redraw = False
                If NROPED_Ant <> XVALO(!NROPED) And J& > 1 Then
                   J& = J& + 1
                End If
                J& = J& + 1
                GRILLAGRAL.GRID.Rows = J& + 1
                GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI1%
                GRILLAGRAL.GRID.TextMatrix(J&, 2) = XVALO(!NROPED)
                GRILLAGRAL.GRID.TextMatrix(J&, 3) = FECHATEX$(CVINT(!FeSolEnt))
                GRILLAGRAL.GRID.TextMatrix(J&, 4) = XVALO(!NROCLIE)
                GRILLAGRAL.GRID.TextMatrix(J&, 5) = SAFETEXT$(!RasoClie)
                GRILLAGRAL.GRID.TextMatrix(J&, 6) = SAFETEXT$(!CODIEXT)
                GRILLAGRAL.GRID.TextMatrix(J&, 7) = SAFETEXT$(!Descrip)
                GRILLAGRAL.GRID.TextMatrix(J&, 8) = FORMATNUM$(XVALO(!CANTIPEND), "F10.1")
                If TOTAL_STOCK_FABPEND(CI1%) > XVALO(!CANTIPEND) Then
                    GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(TOTAL_STOCK_FABPEND(CI1%), "F10.1")
                    TOTAL_STOCK_FABPEND(CI1%) = TOTAL_STOCK_FABPEND(CI1%) - XVALO(!CANTIPEND)
                    Call COLOREAR_GRILLA_SOLO_UNA_AFILA(GRILLAGRAL.GRID, &HC0C0C0, J&)
                    If XVALO(TOTAL_STOCK_FABPEND(CI1%)) < 0 Then TOTAL_STOCK_FABPEND(CI1%) = 0
                End If
                NROPED_Ant = XVALO(!NROPED)
                .MoveNext
            Loop
        End With
        RS.Close
        Set RS = Nothing
        GRILLAGRAL.Caption = "PEDIDOS PARA EMBALAR (Resta en el Stock los Items que han sido Despachados)"
        GRILLAGRAL.GRID.Redraw = True
        Screen.MousePointer = 1
        GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
        GRILLAGRAL.mnuMenu1.Visible = True
        GRILLAGRAL.mnuMenu2.Caption = "Despachar"
        GRILLAGRAL.mnuMenu2.Visible = True
        GRILLAGRAL.mnuMenu3.Caption = "Seleccionar Todo/Revertir Selección"
        GRILLAGRAL.mnuMenu3.Visible = True

        GRILLAGRAL.GRID.ZOrder 0
        GRILLAGRAL.TXTBUSCAR = ""
        'Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
        If J& > 0 Then
           GRILLAGRAL.Show 1
        Else
           Call COMUNI("No se Detectaron Pedidos Pendientes Que No Tengan Stocks Suficiente")
           Exit Sub
        End If
        '
        Dim MSF As msFlexGrid
        Set MSF = GRILLAGRAL.GRID
        If TRIM$(GRILLAGRAL.TXTBUSCAR) <> "" Then Set MSF = GRILLAGRAL.MSF_2
        If GRILLAGRAL.MENU1 > 0 Then
            OPX% = COMALTER%("Opciones de Impresión\\Imprimir Toda la Grilla\Imprimir Solo los Items Seleccionados")
            SoloSelec% = 0: If OPX% = 2 Then SoloSelec% = 1
            Call CARGA_TABLA_IMPRE(MSF, Campos$, "PEDIEMBA", "Pedidos a Embalar", "TABLA_IMPRE", RET_NAMECONS$, 1, SoloSelec%)
            Call FINAL("?" & RET_NAMECONS$)
'            Call EliminarSQF(RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
        ElseIf GRILLAGRAL.MENU2 > 0 Then ' EDICIÓN PARA QUE PINTE LAS FILAS
           FIN& = MSF.Rows - 1
           MSF.Redraw = False
           For i& = 1 To FIN&
              If MSF.TextMatrix(i&, 0) = "*" Then
                  CI1% = XVALO(MSF.TextMatrix(i&, 1))
                  NP& = XVALO(MSF.TextMatrix(i&, 2))
                  Call ACTUREGISTRO("PEDDETA", "STADESPA", "NROPED = " & NP& & "AND CODIINT = " & CI1%, 1, REGAF&, "NOMESS")
              End If
           Next i&
           'EN ESTE CASO IMPRIMO DE TODOS MODOS SOLO LO SELECCIONADO
           Call CARGA_TABLA_IMPRE(MSF, Campos$, "PEDIEMBA", "Pedidos a Embalar", "TABLA_IMPRE", RET_NAMECONS$, 1, 1)
           Call FINAL("?" & RET_NAMECONS$)
           'Call EliminarSQF(RET_NAMECONS$)
           GRILLAGRAL.MENU2 = 0
           Screen.MousePointer = 1

        End If
        Set MSF = Nothing
'
End Sub



Private Sub mnuPedidosyFabricacion_Click()
20      Campos$ = "OT N°/A12;" '1
        Campos$ = Campos$ + " Fecha/D8;" '2
        Campos$ = Campos$ + " Cod.BB./A8;" '3
        Campos$ = Campos$ + " F.Técnica./A16;" '4
        Campos$ = Campos$ + " Medidas/A12;" '5
        Campos$ = Campos$ + " N°Ped./f6;" '6
        Campos$ = Campos$ + " Razón Social/A32;" '7
        Campos$ = Campos$ + " Orden de Compra/A12;" '8
        Campos$ = Campos$ + " Vendedor/F8;" '9
        Campos$ = Campos$ + " F.Ped./D8;" '10
        Campos$ = Campos$ + " F.Ent/D8;" '111
        Campos$ = Campos$ + " Esp.Cliente/A12;" '12
        Campos$ = Campos$ + " Cant.A Fab./F12;" '13
        Campos$ = Campos$ + " Mon./A6;" '14
        Campos$ = Campos$ + " P.Unit./F12.2;" '15
        Campos$ = Campos$ + " Importe/F12.2" '16
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)

        SQLX$ = "SELECT O.IDSUBOR,O.FechGen ,R.CODBIGBAG,O.Cod_Exte"
        SQLX$ = SQLX$ + " , RTRIM(CONVERT(CHAR,CPOANCHOCM)) + 'x' + RTRIM(CONVERT(CHAR,CPOLARGOCM)) + 'x' +  RTRIM(CONVERT(CHAR,CPOALTOCM))  + ' cm' AS MEDIDAS"
        SQLX$ = SQLX$ + " ,P.nroped,P.RaSoClie,P.NroOcom,P.VENDED,P.FechEmis,P.FeSolEnt,r.CODESPTECTE,o.CanProy"
        SQLX$ = SQLX$ + " ,Case (p.MoneEmis)"
        SQLX$ = SQLX$ + " when '1' then '$'"
        SQLX$ = SQLX$ + " when '2' then 'U$S'"
        SQLX$ = SQLX$ + " end Moneda"
        SQLX$ = SQLX$ + " ,p.PreUnid,(p.PreUnid*CanProy) as importe"
        
        SQLX$ = SQLX$ + " FROM PEDDETA P WITH(NOLOCK) LEFT JOIN  RECETACTP R"
        SQLX$ = SQLX$ + " ON P.CODIINT = R.CODINT LEFT JOIN ORDEFABR O ON R.CODINT = O.Cod_Inte"
        SQLX$ = SQLX$ + " Where o.Cod_Inte > 0 And p.codiint > 0 And r.CODINT > 0 "
        SQLX$ = SQLX$ + " AND P.NROPED = O.NROPED " 'SI NO INCLUIA OTROS PEDIDOS
        SQLX$ = SQLX$ + " AND O.NuOrNroPed = P.NuOrItem" 'SI NO DUPLICA  EN LOS PEDIDOS CON MAS DE UN ITEM
        SQLX$ = SQLX$ + " order by o.IdSubOr "
        Set RS = READSET(SQLX$)
        Call Carga_MSF_Recordset(RS, Campos$, GRILLAGRAL.GRID)

        RS.Close
        Set RS = Nothing

        GRILLAGRAL.Caption = "Pedidos Y Fabricación"
        GRILLAGRAL.mnuMenu2.Visible = False
        GRILLAGRAL.mnuMenu3.Visible = False
        GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
        GRILLAGRAL.mnuMenu1.Visible = True
        GRILLAGRAL.GRID.ZOrder 0
        GRILLAGRAL.TXTBUSCAR = ""
        '
        GRILLAGRAL.Show 1
        Dim MSF As msFlexGrid
        Set MSF = GRILLAGRAL.GRID

        If GRILLAGRAL.MENU1 > 0 Then
            If TRIM$(GRILLAGRAL.TXTBUSCAR) <> "" Then Set MSF = GRILLAGRAL.MSF_2
            Call CARGA_TABLA_IMPRE(MSF, Campos$, "PEDIDFAB", "Pedidos y Fabricación", "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
        End If
       Set MSF = Nothing



End Sub

Private Sub mnuPedPendAprob_Click()
    APRONP07.Command6_Click
End Sub

Private Sub mnuPedPendinteractivo_Click()
    Dim obj As New RECORXET
    Campos$ = "Cta/F6.0;R.Social/A24;F.Ped./D8;Nro.Ped./F7.0;Código/A16;Descripición/A24;Cantidad/f12.0;F.Sol./D8;Ord.Compra/A16;OrdPed/F6.0"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
10  FRMZELECHO.MSF1.Rows = 1
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   

    CAMPOS_BD$ = "NROCLIE,RaSoClie, FECHEMIS, NROPED,  CodiExt, Descrip, CANPED, FESOLeNT "
    Set rs1 = obj.RS_Pedidos_Sin_OT()  ' FUNCION QUE DEVUELVE UN RECORDSET
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2, 1)
    

    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
    FRMZELECHO.Width = 14500
    If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
        FRMZELECHO.Caption = "Items de Pedidos Sin O.Trabajo Generada - (Doble Click Sobre Items Ver Pedido)"
    Else
       FRMZELECHO.Caption = "Items de Pedidos Sin O.Trabajo Generada - (Doble Click Sobre Items Ver Pedido - Ver Ficha Técnica - Generar Orden de Trabajo"
    End If
    FRMZELECHO.Show 1
    '
    NROCIE = XVALO(RESP_ZELE_VECT(1))
    NPED& = XVALO(RESP_ZELE_VECT(4))
    CIXI% = CODINT%(SAFETEXT$(RESP_ZELE_VECT(5)))
    CODI$ = SAFETEXT$(RESP_ZELE_VECT(5))
    CAN = XVALO(RESP_ZELE_VECT(7))
    FentSol% = FECHANUM(SAFETEXT$(RESP_ZELE_VECT(8)))
    OCOM$ = SAFETEXT$(RESP_ZELE_VECT(9))
    NORD% = XVALO(RESP_ZELE_VECT(10))
    '
    TX1$ = "Cliente: " & SAFETEXT$(RESP_ZELE_VECT(2))
    TX1$ = TX1$ + " - Pedido: " & NPED&
    TX1$ = TX1$ + " - Item: " & DESC$
    '
    If NROCIE < 1 Then Exit Sub
    'OPCIONES
    If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
       GoTo 25
    End If
20  OPX% = COMALTER%(TX1$ & "\\Ver Pedido\Ver Ficha Técnica\Generar Orden de Trabajo")
    If OPX% < 1 Then Exit Sub
    '
    '*************************************/
    If OPX% = 1 Then ' VER IMAGEN DE PEDIDO
25    DOCUNU& = 0
      NUDOCU$ = SAFETEXT$(NPED&)
30    OPX2% = ALTERNA%("Visualizar Pedidos - Orden de Venta\Visualizar Orden de Despacho")
      If OPX2% < 1 Or OPX% > 2 Then Exit Sub
      '
      If OPX2% <> 2 Then
        If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
           Call MENSERR(24, "Usuario No Autorizado para Visualizar Pedidos.\Seleccione Orden de Despacho")
           GoTo 30
        End If
      End If
      If OPX2% = 1 Then
         TIDOCUM$ = "Pedido de Cliente"
      ElseIf OPX2% = 2 Then
         TIDOCUM$ = "Orden de Despacho"
      End If
      While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
      NUDOCU$ = NUDOCU$ + " "

      Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPED&), NUDOCU$)
      GoTo 10
      '*******************************************
      '
    ElseIf OPX% = 2 Then
       FRMENVAP.TXTCODIGO = CODI$
       FRMENVAP.Show 1
       GoTo 20
    ElseIf OPX% = 3 Then ' GENERA ORDEN DE TRABAJO
      If TRIM$(APROBO$(CIXI%)) = "" Then
         Call COMUNI("Imposible Generar Orden de Trabajo\Items Seleccionado No Está Aprobado")
         GoTo 20
      Else
         Call GENERAR_OF_SIMPLE(CIXI%, NROCIE, CAN, FENTRESOL%, NPED&, OCOM$, RETORNO_NORFA$, NORD%)
         FILTX$ = TRIM$(Str$(CIXI%)) & ";" & TRIM$(Str$(NPED&)) & ";" & TRIM$(RETORNO_NORFA$) & ";" & TRIM$(MEDIDA_ENVAPLAST$(CIXI%, RETCORTE#))
         Call STDFORM("ORTR-ENV", FILTX$)
      End If
    End If
    
End Sub

Private Sub mnuPedPendStOf_click()
    Call PDPNDTYSTOCK
End Sub
Sub PDPNDTYSTOCK()

10    DES1% = 0
      HAS1% = 0
      Call DESHASFECHA(DES1%, HAS1%, PERIO1$)
      If PERIO1$ <> "" Then
            Desde = FECHATEX$(DES1%)
            Hasta = FECHATEX$(HAS1%)
            If DES1% > HAS1% Then
               Call COMUNI("Rango de Fechas Incorrecto")
               GoTo 10
            End If
      End If

      Campos$ = "Cliente/f7;R.Social/A24;Pedido/F7;Codigo/a16;Descripcion/A24;Cant.Ped./f9.1;Cant.Ent/f8.1;Saldo/F9.1;O.F./A12;Cant.Proy./F9;Pend.Fab./F9.1;Emisión/D8"
      Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
      Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
      FRMZELECHO.Caption = "PEDIDOS PENDIENTES "
      'agregue esta select para cargar los vectores y no traer con funciones en la select de abajo por qeu demora mucho
      Screen.MousePointer = 11
      SQLX$ = "SELECT O.NroPed,O.IDSUBOR,O.CANPROY,O.CanApro,NumeOrFa,O.Cod_Inte,O.Cod_Exte ,p.Status"
      SQLX$ = SQLX$ + " , p.Fechemis FROM PEDDETA P WITH(NOLOCK)"
      SQLX$ = SQLX$ + " Inner Join"
      SQLX$ = SQLX$ + " ORDEFABR O ON P.NroPed = O.NroPed"
    
      SQLX$ = SQLX$ + " Where P.CanPed > P.CantEnt"
      If DES1% > 0 Then SQLX$ = SQLX$ + " and P.Fechemis >= '" & Desde & "' AND P.Fechemis <= '" & Hasta & "'"
      
      SQLX$ = SQLX$ + " GROUP BY O.IDSUBOR,O.CANPROY,O.CanApro,NumeOrFa,O.Cod_Inte,O.Cod_Exte ,p.Status,o.NroPed,p.Fechemis"
      SQLX$ = SQLX$ + "  ORDER BY o.NumeOrFa  "
      i& = 0
      Set RS = READSET(SQLX$)
      With RS
        Do While Not .EOF
          i& = i& + 1
          FRMZELECHO.msf2.Rows = i& + 1
 
          NUMEORFA1 = XVALO(!NUMEORFA)
          NROPED1& = XVALO(!NROPED)
          NORFA$ = SAFETEXT$(!idsubor)
          CANPROY1 = XVALO(!CANPROY)
          CANAPRO1 = XVALO(!CanApro)
          SaldoAFabricar = XVALO(!CANPROY) - XVALO(!CanApro)
          CI1% = XVALO(!cod_inte)
          Codigo$ = SAFETEXT(!Cod_Exte)
          FE% = CVINT(!FECHEMIS)
            CONDI$ = "CODIMOVI = 'RT'  and NuPedCli = " & NROPED1& & " AND IDENLOTE  = '" & NORFA$ & "'"
          CantEntregSegunLote = XVALO(VALOBADA("MOVISTO", "Sum(cantsalid)", CONDI$, "NOMESS"))
          NC = NroClienteSegunPedido&(NROPED1&)
          FRMZELECHO.msf2.TextMatrix(i&, 1) = NC
          FRMZELECHO.msf2.TextMatrix(i&, 2) = rasocli$(NC)
          FRMZELECHO.msf2.TextMatrix(i&, 3) = NROPED1&
          FRMZELECHO.msf2.TextMatrix(i&, 4) = Codigo$
          FRMZELECHO.msf2.TextMatrix(i&, 5) = DESCRIT0$(CI1%)
          CANTIPED = XVALO(VALOBADA("PEDDETA", "canped", "NROPED =" & NROPED1& & " AND CODIINT = " & CI1%, "NOMESS"))
          FRMZELECHO.msf2.TextMatrix(i&, 6) = CANTIPED
          
          FRMZELECHO.msf2.TextMatrix(i&, 7) = FORMATNUM$(XVALO(CantEntregSegunLote), "F10.1")
          PendEntrega = CANTIPED - CantEntregSegunLote
          FRMZELECHO.msf2.TextMatrix(i&, 8) = FORMATNUM$(XVALO(PendEntrega), "F10.1")


          FRMZELECHO.msf2.TextMatrix(i&, 9) = NORFA$
          FRMZELECHO.msf2.TextMatrix(i&, 10) = FORMATNUM$(CantiProyectadaOF(NORFA$), "F10.1")
          SALDOOF = FORMATNUM$(CANPROY1 - CANAPRO1, "F10.1")
          FRMZELECHO.msf2.TextMatrix(i&, 11) = FORMATNUM$(SALDOOF, "F10.1")
          FRMZELECHO.msf2.TextMatrix(i&, 12) = FECHATEX$(CVINT(!FECHEMIS))
          .MoveNext
        Loop
      End With
      RS.Close
      Set RS = Nothing
      Screen.MousePointer = 1
      '****************************************************************
'      Screen.MousePointer = 11
'      Dim obj As New RECORXET
'      Set rs1 = obj.RS_Situación_de_Pendientes_de_Entrega(DES1%, HAS1%)
'      With rs1
'        i& = 0
'        Do While Not .EOF
'    '      SQLX$ = "SELECT PEDDETA.CODIINT,PEDDETA.NroClie, PEDENCA.RasoClie, PEDDETA.NroPed, PEDDETA.CodiExt, PEDDETA.Descrip, PEDDETA.CanPed, PEDDETA.CantEnt FROM PEDENCA WITH(NOLOCK) INNER JOIN PEDDETA ON PEDENCA.NroPed = PEDDETA.NroPed"
'          i& = i& + 1
'          PEDIX& = XVALO(!NROPED)
'          FRMZELECHO.MSF2.Rows = i& + 1
'          FRMZELECHO.MSF2.TextMatrix(i&, 1) = XVALO(!NROCLIE)
'          FRMZELECHO.MSF2.TextMatrix(i&, 2) = SAFETEXT$(!RasoClie)
'          FRMZELECHO.MSF2.TextMatrix(i&, 3) = PEDIX&
'          FRMZELECHO.MSF2.TextMatrix(i&, 4) = SAFETEXT$(!CODIEXT)
'          FRMZELECHO.MSF2.TextMatrix(i&, 5) = SAFETEXT$(!Descrip)
'          FRMZELECHO.MSF2.TextMatrix(i&, 6) = FORMATNUM$(XVALO(!CanPed), "F10.1")
'          FRMZELECHO.MSF2.TextMatrix(i&, 7) = FORMATNUM$(XVALO(!CANTENT), "F10.1")
'          FRMZELECHO.MSF2.TextMatrix(i&, 8) = FORMATNUM$(XVALO(!CanPed) - XVALO(!CANTENT), "F10.1")
'          NORFA$ = ""
'          SALDOOF = 0
'
'          For J& = 1 To UBound(NROPED)
'              If NROPED(J&) = PEDIX& Then 'CARGO EL MSF SEGUN EL VECTOR CARGADO ARRIBA
'                If CII1%(J&) = XVALO(!CODIINT) Then
'                  NORFA$ = idsubor(J&)
'                  SALDOOF = SALDOAFABRIX(J&)
'                  Exit For
'                End If
'              End If
'          Next J&
'          FRMZELECHO.MSF2.TextMatrix(i&, 9) = NORFA$
'          FRMZELECHO.MSF2.TextMatrix(i&, 10) = FORMATNUM$(SALDOOF, "F10.1")
'          FRMZELECHO.MSF2.TextMatrix(i&, 11) = FECHATEX$(CVINT(!FECHEMIS))
'
'         .MoveNext
'        Loop
'       End With
'       Set obj = Nothing
'       rs1.Close
'       Set rs1 = Nothing
       FRMZELECHO.Width = 16500
       FRMZELECHO.msf2.Font.Size = 9
       FRMZELECHO.MSF1.Font.Size = 9
       FRMZELECHO.mnuImprimir.Visible = True
       Screen.MousePointer = 1
        
60     FRMZELECHO.Show 1
       ARGUBUS$ = TRIM$(FRMZELECHO.TXTBUSCAR)
       If FRMZELECHO.IMPRIMIR > 0 Then
          Dim MSF As msFlexGrid
          Set MSF = FRMZELECHO.msf2
          If ARGUBUS$ <> "" Then Set MSF = FRMZELECHO.MSF1
          Call CARGA_TABLA_IMPRE(MSF, Campos$, "PEDPENOF", "PEDIDOS PENDIENTES ", "TABLA_IMPRE", RET_NAMECONS$)
         '*****************************************************'
         Call FINAL("?" & RET_NAMECONS$)
         Exit Sub
          FILTX$ = TRIM$("PEDIDOS PENDIENTES CON STOCK Y OF")
          Call STDFORM("IMPR-STD", FILTX$)
          GoTo 60
       End If
       
      NPED& = XVALO(RESP_ZELE_VECT(3))
      If NPED& < 1 Then Exit Sub
      
      DOCUNU& = 0
      NUDOCU$ = SAFETEXT$(NPED&)
30    OPX2% = ALTERNA%("Visualizar Pedidos - Orden de Venta\Visualizar Orden de Despacho")
      If OPX2% < 1 Or OPX2% > 2 Then GoTo 60
      '
      If OPX2% <> 2 Then
        If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
           Call MENSERR(24, "Usuario No Autorizado para Visualizar Pedidos.\Seleccione Orden de Despacho")
           GoTo 30
        End If
      End If
      If OPX2% = 1 Then
         TIDOCUM$ = "Pedido de Cliente"
      ElseIf OPX2% = 2 Then
         TIDOCUM$ = "Orden de Despacho"
      End If
      While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
      NUDOCU$ = NUDOCU$ + " "

      Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPED&), NUDOCU$)
      GoTo 60

'10    DES1% = 0
'      HAS1% = 0
'      Call DESHASFECHA(DES1%, HAS1%, PERIO1$)
'      If PERIO1$ <> "" Then
'            Desde = FECHATEX$(DES1%)
'            Hasta = FECHATEX$(HAS1%)
'            If DES1% > HAS1% Then
'               Call COMUNI("Rango de Fechas Incorrecto")
'               GoTo 10
'            End If
'      End If
'
'      Campos$ = "Cliente/f7;R.Social/A24;Pedido/F7;Codigo/a16;Descripcion/A24;Cant.Ped./f9.1;Cant.Ent/f8.1;Saldo/F9.1;O.F./A12;Pend.Fab./F9.1;Emisión/D8"
'      Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
'      Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
'      FRMZELECHO.Caption = "PEDIDOS PENDIENTES "
'      'agregue esta select para cargar los vectores y no traer con funciones en la select de abajo por qeu demora mucho
'
'      SQLX$ = "SELECT O.NroPed,O.IDSUBOR,O.CANPROY,O.CanApro,NumeOrFa,O.Cod_Inte,O.Cod_Exte ,p.Status"
'      SQLX$ = SQLX$ + " , p.Fechemis FROM PEDDETA P WITH(NOLOCK)"
'      SQLX$ = SQLX$ + " Inner Join"
'      SQLX$ = SQLX$ + " ORDEFABR O ON P.NroPed = O.NroPed"
'
'      SQLX$ = SQLX$ + " Where P.CanPed > P.CantEnt"
'      If DES1% > 0 Then SQLX$ = SQLX$ + " and P.Fechemis >= '" & Desde & "' AND P.Fechemis <= '" & Hasta & "'"
'
'      SQLX$ = SQLX$ + " GROUP BY O.IDSUBOR,O.CANPROY,O.CanApro,NumeOrFa,O.Cod_Inte,O.Cod_Exte ,p.Status,o.NroPed,p.Fechemis"
'      SQLX$ = SQLX$ + " ORDER BY o.NumeOrFa  "
'      Dim NROPED(), SALDOAFABRIX(), idsubor(), CII1%()
'      i& = 0
'      Set RS = READSET(SQLX$)
'      With RS
'        Do While Not .EOF
'          i& = XVALO(!NUMEORFA) 'TOMO COMO INDICE EL N. DE OF P Q EL PEDIDO SE PUEDE REPETIR CUANDO EL MISMO TIENE MAS DE UN ITEM
'          ReDim Preserve NROPED(i&), SALDOAFABRIX(i&), idsubor(i&), CII1%(i&)
'          'If XVALO(!Cod_Inte) > UBound(STX) Then ReDim Preserve STX(XVALO(!Cod_Inte))
'          NROPED(i&) = XVALO(!NROPED)
'          idsubor(i&) = SAFETEXT$(!idsubor)
'          SALDOAFABRIX(i&) = XVALO(!CANPROY) - XVALO(!CanApro)
'          CII1%(i&) = XVALO(!cod_inte)
'          'STX(XVALO(!Cod_Inte)) = STOCKACT(XVALO(!Cod_Inte)) 'lo comente por que demora mucho
'          .MoveNext
'        Loop
'      End With
'      RS.Close
'      Set RS = Nothing
'
'      '****************************************************************
'      Screen.MousePointer = 11
'      Dim obj As New RECORXET
'      Set rs1 = obj.RS_Situación_de_Pendientes_de_Entrega(DES1%, HAS1%)
'      With rs1
'        i& = 0
'        Do While Not .EOF
'    '      SQLX$ = "SELECT PEDDETA.CODIINT,PEDDETA.NroClie, PEDENCA.RasoClie, PEDDETA.NroPed, PEDDETA.CodiExt, PEDDETA.Descrip, PEDDETA.CanPed, PEDDETA.CantEnt FROM PEDENCA WITH(NOLOCK) INNER JOIN PEDDETA ON PEDENCA.NroPed = PEDDETA.NroPed"
'          i& = i& + 1
'          PEDIX& = XVALO(!NROPED)
'          FRMZELECHO.MSF2.Rows = i& + 1
'          FRMZELECHO.MSF2.TextMatrix(i&, 1) = XVALO(!NROCLIE)
'          FRMZELECHO.MSF2.TextMatrix(i&, 2) = SAFETEXT$(!RasoClie)
'          FRMZELECHO.MSF2.TextMatrix(i&, 3) = PEDIX&
'          FRMZELECHO.MSF2.TextMatrix(i&, 4) = SAFETEXT$(!CODIEXT)
'          FRMZELECHO.MSF2.TextMatrix(i&, 5) = SAFETEXT$(!Descrip)
'          FRMZELECHO.MSF2.TextMatrix(i&, 6) = FORMATNUM$(XVALO(!CanPed), "F10.1")
'          FRMZELECHO.MSF2.TextMatrix(i&, 7) = FORMATNUM$(XVALO(!CANTENT), "F10.1")
'          FRMZELECHO.MSF2.TextMatrix(i&, 8) = FORMATNUM$(XVALO(!CanPed) - XVALO(!CANTENT), "F10.1")
'          NORFA$ = ""
'          SALDOOF = 0
'
'          For J& = 1 To UBound(NROPED)
'              If NROPED(J&) = PEDIX& Then 'CARGO EL MSF SEGUN EL VECTOR CARGADO ARRIBA
'                If CII1%(J&) = XVALO(!CODIINT) Then
'                  NORFA$ = idsubor(J&)
'                  SALDOOF = SALDOAFABRIX(J&)
'                  Exit For
'                End If
'              End If
'          Next J&
'          FRMZELECHO.MSF2.TextMatrix(i&, 9) = NORFA$
'          FRMZELECHO.MSF2.TextMatrix(i&, 10) = FORMATNUM$(SALDOOF, "F10.1")
'          FRMZELECHO.MSF2.TextMatrix(i&, 11) = FECHATEX$(CVINT(!FECHEMIS))
'
'         .MoveNext
'        Loop
'       End With
'       Set obj = Nothing
'       rs1.Close
'       Set rs1 = Nothing
'       FRMZELECHO.Width = 16500
'       FRMZELECHO.MSF2.Font.Size = 9
'       FRMZELECHO.MSF1.Font.Size = 9
'       FRMZELECHO.mnuimprimir.Visible = True
'       Screen.MousePointer = 1
'
'60     FRMZELECHO.Show 1
'       ARGUBUS$ = TRIM$(FRMZELECHO.TXTBUSCAR)
'       If FRMZELECHO.IMPRIMIR > 0 Then
'          Dim MSF As msFlexGrid
'          Set MSF = FRMZELECHO.MSF2
'          If ARGUBUS$ <> "" Then Set MSF = FRMZELECHO.MSF1
'          Call CARGA_TABLA_IMPRE(MSF, Campos$, "PEDPENOF", "PEDIDOS PENDIENTES ", "TABLA_IMPRE", RET_NAMECONS$)
'         '*****************************************************'
'         Call FINAL("?" & RET_NAMECONS$)
'         Exit Sub
'          FILTX$ = TRIM$("PEDIDOS PENDIENTES CON STOCK Y OF")
'          Call STDFORM("IMPR-STD", FILTX$)
'          GoTo 60
'       End If
'
'      NPED& = XVALO(RESP_ZELE_VECT(3))
'      If NPED& < 1 Then Exit Sub
'
'      DOCUNU& = 0
'      NUDOCU$ = SAFETEXT$(NPED&)
'30    OPX2% = ALTERNA%("Visualizar Pedidos - Orden de Venta\Visualizar Orden de Despacho")
'      If OPX2% < 1 Or OPX2% > 2 Then GoTo 60
'      '
'      If OPX2% <> 2 Then
'        If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
'           Call MENSERR(24, "Usuario No Autorizado para Visualizar Pedidos.\Seleccione Orden de Despacho")
'           GoTo 30
'        End If
'      End If
'      If OPX2% = 1 Then
'         TIDOCUM$ = "Pedido de Cliente"
'      ElseIf OPX2% = 2 Then
'         TIDOCUM$ = "Orden de Despacho"
'      End If
'      While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
'      NUDOCU$ = NUDOCU$ + " "
'
'      Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPED&), NUDOCU$)
'      GoTo 60


  End Sub



Private Sub mnupedpenImprimible_Click()
'ESTA CONSULTA DEVUELVE LOS ITEMS  QUE NO EXISTEN EN LA TABLA ORDEFABR Y SI EN EL PEDDETA
    Dim obj As New RECORXET
    Campos$ = "Cta/F6;R.Social/A24;F.Ped./D8;Nro.Ped./F7;Código/A16;Descripición/A24;Cantidad/f12;F.Sol./D8"
    Set rs1 = obj.RS_Pedidos_Sin_OT ' FUNCION QUE DEVUELVE UN RECORDSET
    
    With rs1
      Do While Not .EOF
        'GUARDO EN UN VECTOR LOS CLIENTES QUE NO ESTAN EN PEDENCA
          NC = XVALO(!NROCLIE)
          RASO$ = rasocli$(NC)
          NPED& = XVALO(!NROPED)
          FEMI% = CVINT(!FECHEMIS)
          'CIXI% = XVALO(!CODIINT)
          Codigo$ = SAFETEXT$(!CODIEXT)
          DESC$ = SAFETEXT$(!Descrip)
          CantSol = XVALO(!CanPed)
          FentSol% = CVINT(!FeSolEnt)
TEXTOGRILLA$ = NC & "|" & RASO$ & "|" & FECHATEX$(FEMI%) & "|" & NPED& & "|"
TEXTOGRILLA$ = TEXTOGRILLA$ + Codigo$ & "|" & DESC$ & "|" & CantSol & "|" & FECHATEX$(FentSol%)
'Call CARGA_GRILLA_SHOWREP(SHOWREP07, SHOWREP07.GRID, Campos$, TEXTOGRILLA$, INICIALIZO%)
         .MoveNext
      Loop
    End With
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
    
'MUESTRA FRM DE GENREP07
SHOWREP07.Label7 = SHOWREP07.GRID.Rows - 1
SHOWREP07.Frame1.Visible = False
SHOWREP07.Label3 = "Items de Pedidos Sin O.Trabajo Generada"
Screen.MousePointer = 1
Call COLOREAR_GRILLASQL2(SHOWREP07.GRID, &HC0C0C0, 1)

SHOWREP07.Show 1

End Sub


Private Sub mnupendienteaprobacion_Click()
    APRONP07.Command4_Click
End Sub

Private Sub mnupresuppendientes_Click()
    'ARMO LISTA DE TODOS LOS PRESUPUESTOS AUN PENDIENTES DE GENERAR PEDIDOS
    Campos$ = "Cta./F7;Razon Social/A24;Nro.Presup./F10;Cant./F8.1;Código/A16;Descripción/A32"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    FRMZELECHO.Width = 10000
    FRMZELECHO.Caption = "Presupuestos Que Aún No se Transformaron en Pedidos"
    FRMZELECHO.MSF1.Rows = 1
    
    SQLX$ = "SELECT * from presuven  WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NROPEDI < 1 AND MARNOVIG IS NULL OR MARNOVIG <> '*' "
    SQLX$ = SQLX$ + " ORDER BY NROPRES"
    
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.msf2.Rows = i& + 1
        FRMZELECHO.msf2.TextMatrix(i&, 1) = XVALO(!NROCLIE)
        FRMZELECHO.msf2.TextMatrix(i&, 2) = SAFETEXT$(!RasoClie)
        FRMZELECHO.msf2.TextMatrix(i&, 3) = XVALO(!NROPRES)
        FRMZELECHO.msf2.TextMatrix(i&, 4) = XVALO(!CANTIT)
        FRMZELECHO.msf2.TextMatrix(i&, 5) = SAFETEXT$(!CODIEXT)
        FRMZELECHO.msf2.TextMatrix(i&, 6) = SAFETEXT$(!Descrip)
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    FRMZELECHO.Show 1
    '
    NROCIE = XVALO(RESP_ZELE_VECT(1))
    Text1 = NROCIE
    'AUTOMATIZA - DIRECTAMENTE MUESTRA EL LIST CON TODOS LOS PRESUPUESTO DEL CLIENTE SELECCIONADO
    If XVALO(Text1) > 0 Then
       Call Text1_KeyPress(13)
       Command6_Click
    End If
End Sub




Private Sub mnuRankingPedidos_Click()
    Call OPPED07.RANKPED
End Sub

Private Sub mnuRastreabilidadPedido_Click()
    
    OPX2% = ALTERNA%("De Factura a Pedido\De Pedido a Factura")
    Select Case OPX2%
       Case 1
        Call OPPED07.FILARTCLI
       Case 2
        Call OPPED07.PEDAFACT
    End Select
    
End Sub


Private Sub mnuReportesVarios_Click()
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnusitstock_Click()
  SITUSTO.Show 1
End Sub

Private Sub mnuSncrnzCPWlbt_Click()
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "ENVAPLAST") > 0 Then
       BDX$ = Control("CONWALDB", "RUTAWALD")
       If BDX$ <> "" Then
        Dim env As New Envap
        Call env.SINCRONIZA_CondVenta_Waldbott
        Set env = Nothing
      End If
    End If

End Sub

Private Sub mnuverpedidos_Click()
   Call CARGAPEDIDOS(0, 1)
   NPX& = XVALO(RESP_ZELE_VECT(1))
   CONPED07.Label7 = TRIM$(Str$(NPX&))
   Call CONPED07.Command6_Click
End Sub

Private Sub mnuVerSituacionpedidoproyeccion_Click()
   DIAGPREM.Show 1
End Sub

Private Sub mnuVisPlantform_Click()
   FRMENVAP.Show 1
End Sub

Private Sub MSF_Click()
   COL& = MSF.MouseCol
   FilaActiva = MSF.Row
   If COL& = 0 Then
      If Control("ADMIPED", "OBSXITEM") = "SI" Then
        With MSF
            .COL = COL&
            Command30.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
            Command30.Visible = True
            Command30.ZOrder 0
        End With
    End If
   End If
End Sub
Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
'  If Button = 2 Then
'    If CONTROL("ADMIPED", "ITEMSCGO") = "SI" Then
'        Set MSFGLOBALFRM = MSF
'        GLOBALFRM_INDIX% = Index
'        REG& = MSFGLOBALFRM.MouseRow
'        COL% = MSFGLOBALFRM.MouseCol
'        Call MENU_CELDA(MSFGLOBALFRM, REG&, COL%)
'    End If
'  End If

End Sub
Private Sub MSF_RowColChange()
    Fila& = MSF.Row
    On Error Resume Next
    Text28 = MSF.TextMatrix(Fila&, 9)
    Command30.Visible = False
    On Error GoTo 0
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
    Campos$ = "Nro Ped./F10;Cliente/F7;R.Social/A20;Codint/F0;Código/F16;Descripción/A32;Cant.Ped./f10.1;Cant.Ent/F10.1;Zona/F0;Zona/A18;Saldo/F10.1;Stock/F10.1;F.Emis./D8;F.Entrega/D8;FEMISNUM/F0;FENTREGANUM/F0"
    Call CARGA_ENCABEZADO(MUESTRADATOS11.GRID, Campos$, MUESTRADATOS11, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(MUESTRADATOS11.MSF_2, Campos$, MUESTRADATOS11, 1, 0, ANTOT, , 1)
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
        CANPEX = XVALO(!CanPed)
        CANTENTX = XVALO(!CANTENT)
        SALDO = XVALO(!CanPed) - XVALO(!CANTENT): If SALDO < 0 Then SALDO = 0
        Codigo$ = SAFETEXT$(!CODIEXT)
        DESCRI$ = SAFETEXT$(!Descrip)
        ZONAX$ = SAFETEXT$(ZONACLI$(NC))
        DESCRIZONA$ = ECOARCH$("ZONAREP", ZONAX$)
        ESTOCKS = FORMATNUM$(STOCKDIS(CI1%), "F10.1")
        FEEMI% = CVINT(!FECHEMIS)
        FESOL% = CVINT(!FeSolEnt)
        '
        J& = J& + 1
        MUESTRADATOS11.GRID.Rows = J& + 1
        MUESTRADATOS11.GRID.TextMatrix(J&, 1) = NPE&
        MUESTRADATOS11.GRID.TextMatrix(J&, 2) = NC
        MUESTRADATOS11.GRID.TextMatrix(J&, 3) = rasocli$(NC)
        MUESTRADATOS11.GRID.TextMatrix(J&, 4) = CI1%
        MUESTRADATOS11.GRID.TextMatrix(J&, 5) = Codigo$
        MUESTRADATOS11.GRID.TextMatrix(J&, 6) = DESCRI$
        MUESTRADATOS11.GRID.TextMatrix(J&, 7) = FORMATNUM$(CANPEX, "F10.1")
            
        MUESTRADATOS11.GRID.TextMatrix(J&, 8) = FORMATNUM$(CANTENTX, "F10.1")
        MUESTRADATOS11.GRID.TextMatrix(J&, 9) = ZONAX$
        MUESTRADATOS11.GRID.TextMatrix(J&, 10) = DESCRIZONA$
        MUESTRADATOS11.GRID.TextMatrix(J&, 11) = FORMATNUM$(SALDO, "F10.1")
        MUESTRADATOS11.GRID.TextMatrix(J&, 12) = FORMATNUM$(ESTOCKS, "F10.1")

        MUESTRADATOS11.GRID.TextMatrix(J&, 13) = FECHATEX$(FEEMI%)
        MUESTRADATOS11.GRID.TextMatrix(J&, 14) = FECHATEX$(FESOL%)
        MUESTRADATOS11.GRID.TextMatrix(J&, 15) = FEEMI%
        MUESTRADATOS11.GRID.TextMatrix(J&, 16) = FESOL%
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    If J& > 0 Then
       Screen.MousePointer = 1
       MUESTRADATOS11.GRID.Visible = True
       MUESTRADATOS11.GRID.ZOrder 0
       MUESTRADATOS11.Caption = "Pedidos Pendientes x Zona"
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
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "ENVALAST") > 0 Then
        On Error Resume Next
        BASEWALD.Close
        On Error GoTo 0
    End If
    
    Call CIERRARCH("*.*")
    Call LIBARCH("*.*")
    Call FINAL("")
End Sub

Function VALIDAR_OCOMP_GENERADA%(NroPedido&)
    CONDI$ = "NroPED = " & NroPedido&
    If CantRegistros&("PEDICOMPRA", CONDI$, "NOMESS") > 0 Then
            VALIDAR_OCOMP_GENERADA% = 1
            OPX% = COMALTER%("Este Pedido Ya Genero Solicitudes de Compras\Si Necesita Modificar:\Anule la-s Solicitud-es y Vuelva a Generar\\Ver Solicitudes Generadas\Cancelar")
            If OPX% < 1 Or OPX% > 1 Then Exit Function
            Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cta.Prov/F9;R.Social/A40;O.Compra/F9;Sol.Compra/F12", FRMZELECHO, , , ANTOT)
            Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Cta.Prov/F9;R.Social/A40;O.Compra/F9;Sol.Compra/F12", FRMZELECHO, , , ANTOT)
            If ANTOT < 5500 Then ANTOT = 5500
            FRMZELECHO.Width = ANTOT + 1100
            FRMZELECHO.Caption = "Lista de Ordenes de Compra - Pedido N°: " & NroPedido&
            '
            FRMZELECHO.MSF1.Rows = 1
            SQLX$ = "SELECT DISTINCT(NroPed),Num_Prov, NumSolCompra FROM PediCOMPRA WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE NroPed = " & NroPedido&
            SQLX$ = SQLX$ + " ORDER BY Num_Prov"
            Set RS = READSET(SQLX$)
            With RS
              Do While Not .EOF
                i& = i& + 1
                FRMZELECHO.msf2.Rows = i& + 1
                NP% = XVALO(!Num_Prov)
                FRMZELECHO.msf2.TextMatrix(i&, 1) = NP%
                FRMZELECHO.msf2.TextMatrix(i&, 2) = rasocli$(-1 * NP%)
                   NSOLIC& = XVALO(!NumSolCompra)
                   OCOMPRAX& = OrdenCompraSegunSolicitud&(NSOLIC&)
                FRMZELECHO.msf2.TextMatrix(i&, 3) = OCOMPRAX&
                If OCOMPRAX& > 0 Then Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.msf2, vbYellow, i&)
                FRMZELECHO.msf2.TextMatrix(i&, 4) = NSOLIC&
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
        If Frame4.Height = 3240 Then
          Frame4.Height = 1560
          If InStr(EPAC$, "GABAL") > 0 Then
            Check2.Visible = False
            Label31.Visible = False
            Text141.Visible = False
            Ltota(3).Top = 1155
            Text14(2).Top = 1050
            Total(2).Top = 1050
          End If
        Else
          Frame4.Height = 3240
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
     If rasocli$(NCLIE) <> "" Then
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
  ORIPED$ = TRIM$(UCase$(Control$("PEDIDO", "ORIDATOS")))
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
Sub MSF_DblClick()
    
    If XVALO(Text1) < 1 Then Exit Sub
    
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
      'If List1.ListCount > 0 Then AA1% = DERACCE%("MODESPE", "Modificacion descuento de pedido")
      If MSF.Rows > 1 Then AA1% = DERACCE%("MODESPE", "Modificacion descuento de pedido")
   End If
   '
End Sub

Private Sub RCtaCte_Click()
Call Command4_Click
End Sub

Private Sub RevAnu_Click()
Call OPPED07.Command13_Click
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
  'Call SELECHO("DEUDOR1")
  'NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
  'Text1.TEXT = TRIM$(Str$(NCLIE))
  NCLIE = ListaSeleccionClientes()
  If NCLIE > 0 Then
    Text1.TEXT = TRIM$(Str$(NCLIE))
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
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "ENVAPLAST") > 0 Then
      If XVALO(Text1.TEXT) > 0 Then
        Call env.ALTA_CLIENTE_WALDBOTCLIENTE(XVALO(Text1), 1)
      End If
    End If
    '
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


Private Sub Text14_Change(Index As Integer)

End Sub

Private Sub Text141_Change()

End Sub

Private Sub Text15_DblClick()
  
    PERCEIIBB.Show 1
        
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
   LENTREGA(0) = ""
   LENTREGA(1) = ""

   LENTREGA(2) = ""
   TRANACT% = 0
   EPAC$ = TRIM$(UCase$(EMPREAC$))
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
         If InStr(EPAC$, "EYM") > 0 Then
           LENTREGA(0) = DOMITRANS$(NTRAI%)
           LENTREGA(1) = LOCATRANS$(NTRAI%)
         End If
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
        If CantRegistros("SUCENTRE", CONDI$) > o Then
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

Private Sub Text21_GotFocus(Index As Integer)
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
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If XVALO(TEXT21(Index).TEXT) < 0 Or XVALO(TEXT21(Index).TEXT) > 255 Then
       TEXT21(Index).TEXT = ""
    End If
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
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
            If InStr(EPAC$, "J Y M") > 0 Then
              Call ACTUALZACION_LISTAS_SELECCION(ARCHIPRE$, "LISTAS DE PRECIOS") ' ACTUALIZA LISTA DE PRECIOS EN SISTEMAS MULTIEMPRESA
            End If
        End If
        MNLX% = MONELISTA%(XVALO(TEXT21(Index)))
        '
        'CONFIGURACION DE MONEDA POR DEFECTO
        MOX% = XVALO(Control("ADMIPED", "MONEDEFE"))
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
19      If Control("PEDIDO", "MODODESC") = "MONTO" And XVALO(TEXT21(0)) = 1 Then
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
        If InStr(EPAC$, "ENVAPLAST") > 0 Then
           Eco21(Index).TEXT = VALOBADA("LICONPAG", "REFERENCIA", "CODICONPAG =  " & XVALO(TEXT21(Index)), "NOMESS")
        Else
           Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(XVALO(TEXT21(Index).TEXT)))
        End If
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
        If Control("PEDIDO", "MODODESC") = "MONTO" And XVALO(TEXT21(0)) = 1 Then
           Call DESCUMONTO(Me)
        End If
        Call Check2_Click ' llama al evento donde se valida el vendedor para definir descuento

      ElseIf Index = 3 Then
        If EXISTE%(LUDAT$ + "TATINGRE.RFS") > 0 Then
           Eco21(Index).TEXT = ECOARCH$("TATINGRE", Chr$(XVALO(TEXT21(Index).TEXT)))
        End If
      ElseIf Index = 4 Then
        If EXISTE%(LUDAT$ + "VENDEENV.RFS") > 0 Then
           Eco21(Index).TEXT = ECOARCH$("VENDEENV", Chr$(XVALO(TEXT21(Index).TEXT)))
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
    EPAC$ = TRIM$(UCase$(EMPREAC$))
     '
    If Index = 0 Then
        ARCHEC$ = "LISTAS": LOCA% = 2
      ElseIf Index = 1 Then
        If InStr(EPAC$, "ENVAPLAST") > 0 Then
          CONDIPAX% = CARGAVENDEDORENVA%
          If CONDIPAX% > 0 Then
             TEXT21(Index).TEXT = CONDIPAX%
          End If
          Exit Sub
        Else
          ARCHEC$ = "CONPAG": LOCA% = 2
        End If
      ElseIf Index = 2 Then
        ARCHEC$ = "VENDEDOR": LOCA% = 1
      ElseIf Index = 3 Then
        If EXISTE%(LUDAT$ + "TATINGRE.RFS") < 1 Then
          Call COMUNI("Tabla de Motivos de Ingreso de Pedidos \No Disponible en la Presente Versión")
          Exit Sub
        End If
        
        ARCHEC$ = "TATINGRE/Tabla de Ingresos": LOCA% = 1
      ElseIf Index = 4 Then
        If EXISTE%(LUDAT$ + "VENDEENV.RFS") < 1 Then
          Call COMUNI("Tabla Extra Vendedores\No Disponible en la Presente Versión")
          Exit Sub
        End If
        
        ARCHEC$ = "VENDEENV/Tabla de Extra Vendedores": LOCA% = 1

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
       If rasocli$(NCLIEN) <> "" Then
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
   MSF.Rows = 1
   Text28 = ""
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
   Frame4.Height = 1560
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
   Command30.Visible = False 'BOTON DE OBSERVACIONES POR ITEM
   Erase CI_VECTOR_TIVA()
   Erase TASAIVAX_CI()
   
   PercepcionIBMisiones_BienDeUso = 0
   On Error Resume Next
   Check3.Value = 0
   CheckAhpRetiraDep.Value = 0
   On Error GoTo 0
End Sub
'
Sub CARDACLI()
   '
   Screen.MousePointer = 11
   MARCONOTA.Visible = False
   MARSELPRESU.Visible = False
   If Control$("ADMIPED", "VERLICRE") = "SI" Then VERLICRE% = 1
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
     If rasocli$(NCLIEN) <> "" Then
       Text2.TEXT = rasocli$(NCLIEN)
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
       Text12.TEXT = TRIM$(POSIVAC$(NCLIEN)) + " - CUIT " + CuitCli$(NCLIEN)
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
       If Control$("ADMIPED", "COPAVARI") = "SI" Then TEXT21(1).TEXT = ""
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
       If Control("NOPREENT", "CONFPEDI") <> "SI" Then
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
    NCX7 = XVALO(Text1)
    If NCX7 < 1 Then Exit Sub
    
    SQLX$ = "SELECT * FROM SADEUPEN WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX7 & " "
    SQLX$ = SQLX$ + "AND (ABS(SaldAcre) >= 0.005 OR ABS(SaldDebe) >= 0.005 "
    If MODOMONEI% = 2 Then
         SQLX$ = SQLX$ + "OR ABS(SaldAcre_Dola) >= 0.005 OR ABS(SaldDebe_Dola) >= 0.005) "
       Else
         SQLX$ = SQLX$ + ") "       ' PARA CERRAR EL PARENTESIS
    End If
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    Set FacPenTemp = READSET(SQLX$)

'    If NCX7 <> XVALO(Text1) Then GoTo 10   ' PARA ATAJAR ERROR CUANDO SE  INGRESA NUMERO DE CUENTA A MANO
    '
    SUDEUDOL# = 0
    REG& = 0
    Call CARGA_ENCABEZADO(VENTABNEW.MSF, "T.Doc:/A20;Documento/F18;F.Emisión/D12;Importe/F12.2;F.Vencimiento./D12;Saldo/F12.2;Sal.Dolar/F12.2", VENTABNEW)
    With FacPenTemp
        Do While Not .EOF
          TIX$ = UCase$(Left$(!codicom_lar, 2))
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "AJ" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "RB" Then TIPOS$ = "Recibo Cobranza"
                              If TIX$ = "VC" Then TIPOS$ = "Venta Contado"
                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                              If TIX$ = "DC" Then TIPOS$ = "Descuento Caja"
                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito"
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1 ''\\\OT-06-0112-WAR-30/03/06///
                              '
          IMPO# = (-1) * ROUND(!IBruComp) * SG%
          SUMAFAC# = SUMAFAC# + IMPO#
          '
          SALDO# = ROUND(!SaldDebe - !SaldAcre)
          SUMADEU# = SUMADEU# + SALDO#    ' (Int(100 * SALDO# + 0.5)) / 100
          '
          NROL& = XVALO(Mid$(!CODICOM_COR, 4, 6))
          'NROX$ = CSTRING$(MKS$(NROL&), 3, 8, 0, 2)
          NROX$ = FORMATNUM$(NROL&, "F8.0")
          FEMI = CVINT(!FECHEMISI)
          FEVE = CVINT(!FEVENCIM)
          '
          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
          SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
          SALTEXDOL$ = ""
          '
          COCOLAR$ = !codicom_lar
          '
          If MODOMONEI% = 2 Then
             SALDOLA# = ROUND(!SaldDebe_DOLA - !SaldAcre_DOLA)
             SALTEXDOL$ = FORMATNUM$(SALDOLA#, "F13.2")
             GoTo 20
          End If
          '
          ' rutina incluida para procesamiento bimonetario
          RECABIN& = 0

20        With VENTABNEW
          REG& = REG& + 1
          .MSF.Rows = REG& + 1
          .MSF.TextMatrix(REG&, 1) = TIPOS$
          .MSF.TextMatrix(REG&, 2) = NROX$
          .MSF.TextMatrix(REG&, 3) = FECHATEX$(FEMI)
          .MSF.TextMatrix(REG&, 4) = IMPTEX$
          .MSF.TextMatrix(REG&, 5) = FECHATEX$(FEVE)
          .MSF.TextMatrix(REG&, 6) = SALTEX$
          .MSF.TextMatrix(REG&, 6) = SUDEUDOL#
          End With
          '
        .MoveNext
        Loop
'      End If
'      On Error GoTo 0
    End With
    VENTABNEW.Caption = "Saldos Pendientes"
    VENTABNEW.Width = 12200
    VENTABNEW.Show 1
    '
999 Screen.MousePointer = 1
    '
'    If InStr(UCase$(EMPREAC$), "FERVI") > 0 Then GoTo 900 'provisorio FFFERVI
'
'    If MONRESU% <= 1 Then
'       If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
'          ATENTI% = 1
'          TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta"
'          If Visible = True Then
'              TTXX$ = TTXX$ + " Elegida."
'            Else
'              TTXX$ = TTXX$ + ":\   \'" + TRIM$(rasocli$(NCX7)) + "'"
'          End If
'          TTXX$ = TTXX$ + "\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Créditos y Cobranzas\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
'          '
'          If COMALTER%(TTXX$) = 2 Then
'              Call REAPLICRE(NCX7)
'              GoTo 100
'            Else
'              ATENTI% = 0
'              Call CIERRARCH("*.*")
'              Call FINAL("")
'          End If
'       End If
'    End If
'    '
'900 DISPONI = XVALO(Text11.TEXT) - XVALO(Text10.TEXT) - XVALO(Text6.TEXT)
'    Text12.ForeColor = &H404080
'    If DISPONI < 0 Then Text12.ForeColor = &HFF&
'    Text12.TEXT = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
'    '
'    Screen.MousePointer = 1
    '
''''''End Sub

''''''    HOII% = HOY(HOS$)
''''''    DIAHOY = FECHATEX$(HOII%)
''''''    If Text8 <> "Fac.Vencidas" Then GoTo 99
''''''    NCLIE = XVALO(Text1)
''''''    If NCLIE > 0 Then
''''''      Call APERBASDAT("CABAN")
''''''      SQLX$ = "SELECT * FROM SADEUPEN "
''''''      SQLX$ = SQLX$ + " Where NuClien = " & NCLIE
''''''      SQLX$ = SQLX$ + " AND SaldDebe > SaldAcre "
''''''      SQLX$ = SQLX$ + " AND FEVENCIM < '" & DIAHOY & "' "
''''''      JJ& = 0
''''''      Dim DEUDATMP As ADODB.Recordset
''''''      Set DEUDATMP = New ADODB.Recordset
''''''25    On Error Resume Next
''''''      DEUDATMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
''''''      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
''''''         On Error GoTo 0
''''''         Call CapturaErrorOpen
''''''         GoTo 25
''''''      End If
''''''      On Error GoTo 0
''''''      '
''''''      With DEUDATMP
''''''        On Error Resume Next
''''''        .MoveFirst
''''''         If Err > 0 Then
''''''           Call COMUNI(" El Cliente no Registra Facturas Vencidas ")
''''''           GoTo 99
''''''          Else
''''''           On Error GoTo 0
''''''           Do While Not .EOF
''''''              SALDEB# = XVALO(!SaldDebe)
''''''              SALACR# = XVALO(!SALDAcre)
''''''              DEUDA$ = SALDEB# - SALACR#
''''''              IMPORIG$ = SAFETEXT(!IBruComp)
''''''              FEVE = CVINT(!FEVENCIM)
''''''              NUMFACT$ = SAFETEXT(!CODICOM_LAR)
''''''              FECHOR = CVINT(!FECHEMISI)
''''''              JJ& = JJ& + 1
''''''              Z$ = Space$(128)
''''''              Call REPLA(Z$, NUMFACT$, 1, 18)
''''''              Call REPLA(Z$, FECHATEX$(FECHOR), 23, 8)
''''''              Call REPLA(Z$, FORMATNUM(IMPORIG$, "F12.4"), 32, 13)
''''''              Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
''''''              Call REPLA(Z$, DEUDA$, 57, 13)
''''''              Call GRAREG("SELDOCAP", Z$, JJ&)
''''''              '
''''''              .MoveNext
''''''           Loop
''''''         End If
''''''         On Error GoTo 0
''''''      End With
''''''      DEUDATMP.Close
''''''      Set DEUDATMP = Nothing
''''''      Call SETULTREG("SELDOCAP", JJ&)
''''''      SPDX$ = REGSEL$("SELDOCAP")
''''''    End If
''''''99
    

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
    Call LIMPIATETAB

    NCLIEN = XVALO(Text1)
    NROPED& = XVALO(Text11)
    FORIPRE$ = TRIM$(Control$("", "FORPECLI"))
    If FORIPRE$ = "" Then Exit Sub
    '
    'SI TIENE FRM DE DISEÑADOR DE FRM VA POR AHI
    If OPCI% = 1 Then
        ARCSQR$ = FORIPRE$
        MODO$ = "PRINT"
        If CANCELPRINT% > 0 Then MODO$ = ""
        If EXISTE%(LUDAT$ + ARCSQR$ + ".SQR") > 0 Then
           FILTX$ = TRIM$(Str$(XVALO(Text11))) & ";" & TRIM$(Str$(NCLIEN))
           Call STDFORM(ARCSQR$, FILTX$, MODO$)
           Exit Sub
        End If
    End If
    
    PUERDESPA$ = ""
    PUERORVE = ""
    If OPCI% = 1 Then
       PUEROVTA$ = Control("*", "PUERORVE")
       If Control(IDENTER$, "PUERORVE") <> "" Then PUEROVTA$ = Control(IDENTER$, "PUERORVE")
    ElseIf OPCI% = 2 Then ' ORDEN DE DESPACHO
       PUERDESPA$ = Control("*", "PUERDESP")
       If Control(IDENTER$, "PUERDESP") <> "" Then PUERDESPA$ = Control(IDENTER$, "PUERDESP")
       FORODES$ = TRIM$(Control$("", "FORDESPA"))
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
    If Control$("PEDCLIEN", "SUPRINT") = "SI" And SENDBYMAIL$ = "" Then
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
    CODTABU1$(17) = "LAR-MACO"
    
    CACOLTAB1% = 17
    '
    'DESTIDOC% = NCLIEN
    DESTIDOCNUE = NCLIEN
    CAITAB1% = 0
    CAITCRU% = 0
    '
    PRIVA$ = Control$("PEDCLIEN", "PRECOIVA")
    DESCPRI% = 0: DESCOMP% = 0: Especif% = 0: ImprimeObser% = 0
    If Control("ADMIPED", "IMPDESPR") = "SI" Then DESCPRI% = 1
    If Control("ADMIPED", "IMPDESCO") = "SI" Then DESCOMP% = 1
    If Control("ADMIPED", "IMPRESPE") = "SI" Then Especif% = 1
    If Control("ADMIPED", "PRINOBSE") = "SI" Then ImprimeObser% = 1

    If DESCOMP% + Especif% < 1 Then DESCPRI% = 1 'POR SI NO SE SELECCIONO NADA
    '
    ORD% = 0
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
            ORD% = ORD% + 1
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
            TETABU1$(9, CAITAB1%) = TRIM$(Str$(ORD%))
            TETABU1$(10, CAITAB1%) = NUPLANO$(CIPRE%)
            TETABU1$(14, CAITAB1%) = Unimed$(CIPRE%) 'SE AGREGA POR QUE ESTABA FALTANDO
            TETABU1$(17, CAITAB1%) = ""
            If InStr(EPAC$, "CONTEPLAST") > 0 Then
               CONDI$ = " NroPres = " & NUPRESU& & " and NuOrdIt= " & NUORIT% 'COMISION REAL GRABADA EN LA COTIZACION
               PorComisReal$ = FORMATNUM$(XVALO(VALOBADA("PRESUVEN", "PresComReal", CONDI$, "NOMESS")), "F5.1")
               TETABU1$(17, CAITAB1%) = TRIM$(PorComisReal$)
            End If
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
            OBSERVACX$ = TRIM$(MSF.TextMatrix(U&, 9))
            
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
            'OBSERVACIONES DEL ITEM.
            If ImprimeObser% > 0 Then
              If OBSERVACX$ <> "" Then
                If DIMPRE$ <> "" Then
                  DIMPRE$ = DIMPRE$ + vbCrLf
                End If
                DIMPRE$ = DIMPRE$ + OBSERVACX$
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
            '
            
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
    CODPARAM$(5) = "DOMI-TRA": If RetiraDomicilioAHP% < 1 Then DOCPARAM$(5) = LENTREGA(0).TEXT Else DOCPARAM$(5) = ""
    CODPARAM$(6) = "DESTINO": If RetiraDomicilioAHP% < 1 Then DOCPARAM$(6) = LENTREGA(1).TEXT Else DOCPARAM$(6) = ""
    CODPARAM$(7) = "TRANSPOR": If RetiraDomicilioAHP% < 1 Then DOCPARAM$(7) = LENTREGA(2).TEXT Else DOCPARAM$(7) = "Retira de Nuestro Depósito"
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
    CODPARAM$(26) = "TELE-EMP": If RetiraDomicilioAHP% < 1 Then DOCPARAM$(26) = TRIM$(TELETRANS$(NTRAI%)) Else DOCPARAM$(26) = ""
    CODPARAM$(27) = "LOCA-TRA": If RetiraDomicilioAHP% < 1 Then DOCPARAM$(27) = LENTREGA(1).TEXT Else DOCPARAM$(27) = ""
    CODPARAM$(28) = "LAR-CUCO": DOCPARAM$(28) = TRIM$(Me.TXTFLETE)
    CODPARAM$(29) = "NUOPERS": DOCPARAM$(29) = FORMATOCON0$(XVALO(Me.Label34), 3)
    CODPARAM$(31) = "REF-MAT3": DOCPARAM$(31) = LOCATRANS$(NTRA%)
    
    'LIMPIO LOS DATOS POR SI NO ENTRA POR QUE SE ESTA QUEDANDO CON LA INFO DEL ANTERIOR
    CODPARAM$(32) = "OBS-OPER": DOCPARAM$(32) = ""
    CODPARAM$(33) = "ZONA-VEN": DOCPARAM$(33) = ""
    CODPARAM$(34) = "DOMI-EMP": DOCPARAM$(34) = ""
    CODPARAM$(35) = "LOCA-EMP": DOCPARAM$(35) = ""
    CODPARAM$(36) = "ACH-CUTE": DOCPARAM$(36) = ""
    CODPARAM$(37) = "CPOS-EMP": DOCPARAM$(37) = ""
    CODPARAM$(38) = "RASO-EMP": DOCPARAM$(38) = ""
    CODPARAM$(39) = "ACH-MATE": DOCPARAM$(39) = ""

    
    If RetiraDomicilioAHP% < 1 Then
      CONDI$ = "NROPED = " & NROPED&
      CONDI$ = CONDI$ + " AND NROCLIE = " & NCLIEN
      INSTRUCCIONTRAN$ = VALOBADA("DatoEntrega", "InstrucTransporte", CONDI$, "NOMESS")
      INSTRUCCIONDESTINO$ = VALOBADA("DatoEntrega", "InstruccionSuc", CONDI$, "NOMESS")
      DOMIDESTINATARIO$ = VALOBADA("DatoEntrega", "DOMI_SUC", CONDI$, "NOMESS")
      LOCALIDADESTINATARIO$ = VALOBADA("DatoEntrega", "LOCA_SUC", CONDI$, "NOMESS")
      LETRAPROVINCIA$ = TRIM$(VALOBADA("DatoEntrega", "ProvEntre_cli", CONDI$, "NOMESS"))
      PROVINCIAX$ = TRIM$(ECOARCH$("PROVINC", LETRAPROVINCIA$))
      CP_DESTINATARIO$ = VALOBADA("DatoEntrega", "CodPostal_Suc", CONDI$, "NOMESS")
      NOMBREDESTINATARIO$ = VALOBADA("DatoEntrega", "DENO_SUC", CONDI$, "NOMESS")
      HORARIOX$ = VALOBADA("DatoEntrega", "HoraTransporte", CONDI$, "NOMESS")
      
      CODPARAM$(32) = "OBS-OPER": DOCPARAM$(32) = INSTRUCCIONTRAN$
      CODPARAM$(33) = "ZONA-VEN": DOCPARAM$(33) = INSTRUCCIONDESTINO$
      CODPARAM$(34) = "DOMI-EMP": DOCPARAM$(34) = DOMIDESTINATARIO$
      CODPARAM$(35) = "LOCA-EMP": DOCPARAM$(35) = LOCALIDADESTINATARIO$
      CODPARAM$(36) = "ACH-CUTE": DOCPARAM$(36) = PROVINCIAX$
      CODPARAM$(37) = "CPOS-EMP": DOCPARAM$(37) = CP_DESTINATARIO$
      CODPARAM$(38) = "RASO-EMP": DOCPARAM$(38) = NOMBREDESTINATARIO$
      CODPARAM$(39) = "ACH-MATE": DOCPARAM$(39) = HORARIOX$
      
    End If
    CAPARDOC% = 39
    '
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
    For i = 1 To CARETEX%
        III& = i
        Call GRAREG("!OBSERVOF", RETEX$(i), III&)
    Next i
    URETE& = CARETEX%
    Call SETULTREG("!OBSERVOF", URETE&)
    Call CIERRARCH("!OBSERVOF")
    '
    Call SETPRINPORT("RESTORE")
     If FORIPRE$ <> "" Then
      'Call PRINTDOC("FORPECLI")   ' PEDIDO DE CLIENTE
      If OPCI% = 1 Then Call SETPRINPORT(PUEROVTA$)
      If OPCI% = 2 Then Call SETPRINPORT(PUERDESPA$)
      RETORNO_NRODOCU& = 0
      Call PRINTDOC("@" + FORIPRE$, "", RETORNO_NRODOCU&) ' PEDIDO DE CLIENTE
      If RETORNO_NRODOCU& > 0 Then Call GRABAR_VERSION_IMAGEN(XVALO(NUMEDOC$), RETORNO_NRODOCU&, TIPODOC$)
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
    Call LIMPIATETAB
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
   If Control$("ADMIPED", "VERLICRE") = "SI" Then VERLICRE% = 1
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
   If Control$("ADMIPED", "OBLICOM") = "SI" Then
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
   REFE$ = Left$(TRIM$(rasocli$(NCLIEN)), LOMA%) + REFE$
   '
   STATUPC% = (-1)
   If Left$(Control$("PEDCLI", "MODOAPRO"), 6) = "DIRECT" Then STATUPC% = 0
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
   If TRIM$(Control("ADMIPED", "VALLISPR")) = "SI" Then
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
   

   strSQL = "Select * from Pedenca where NroPed = " & NROPE&
   Dim PedEnca As ADODB.Recordset
   Set PedEnca = New ADODB.Recordset
25 On Error Resume Next
   PedEnca.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
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
   
   ' *** ELIGE ATRIBUTO COMERCIAL
   If InStr(1, EMPREAC$, "CONTEPLAST") > 0 Then 'USO ESTE CAMPO PARA GUARDAR EL VENDEDOR QUE SE VA A PASAR A ENVAPLAST
      idAtributo% = XVALO(TEXT21(4))
   Else
        idAtributo% = 0
        If CantRegistros("datasql", "lista='ATRICOPE'") > 0 Then
             Dim da As New DatosAuxiliares
             Screen.MousePointer = 1
             Call da.ListaSeleccionTablaAuxiliar("ATRICOPE", "Atributos Comerciales")
             Screen.MousePointer = 11
             idAtributo% = XVALO(RESP_ZELE_VECT(1))
        End If
   End If
35 On Error Resume Next
   PedEnca.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
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
     !REFERENCIA = Left$(REFE$, 128)
     !FECHEMIS = CVDAT(FEMI%)
     !NROCLIE = NCLIEN
     
     !RasoClie = Left$(rasocli$(NCLIEN), 64)
     !NROOCOM = Left$(TRIM$(Text10.TEXT), 16)
     !DOMIENT = Left$(LENTREGA(0), 48)
     !LOCAENT = Left$(LENTREGA(1), 48)
     !TranspEnt = Left$(LENTREGA(2), 48)
     !NTranspo = XVALO(Text19)
     !NSucent = XVALO(Text20)
     !Observa = Text13
     !FeSolEnt = Now
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
           valor% = 0
           If TRIM$(Text44) <> "" Then valor% = Asc(Text44.TEXT + "A") - 64
           !PorComVend = valor%
        End If
     !Rec_ConPa = Left$(TRIM$(Label25), 14)
     !TDescApli = Left$(TRIM$(PORDESCU.TEXT), 14)
     !Dto_Ret_Cte = Left$(TRIM$(Label31), 14)
     !VenRep = Left$(Label11, 1)   ' Venta o Reparacion
     '
     !CODCARPIMP = Left$(CODCARPE$, 12) ' Guarda Nº de Carpeta de Exportaciones
     If CODCARPE$ <> "" Then ' Anotador de Carpeta de Exportaciones
        Call REPLA(TTXX$, HOS$, 1, 8)
        Call REPLA(TTXX$, USERNAME$, 69, 24)
        NR$ = SAFETEXT(Text11)
        Call REPLA(TTXX$, "Generacion de Orden de Venta Nº: " + NR$, 11, 56)
        Call AGREGA_ANOTACARPI(CODCARPE$, TTXX$)
        TTXX$ = Space(76)
        Call REPLA(TTXX$, "Cliente: " + Left$(rasocli$(NCLIEN), 64), 11, 56)
        Call AGREGA_ANOTACARPI(CODCARPE$, TTXX$)
        Call AGREGA_ANOTACARPI(CODCARPE$, String$(88, "="))
     End If
     '
     !FLETE_TEX = TRIM$(Left$(TXTFLETE, 64))
     !Revision = 0
     !AtributoComercial = idAtributo%
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
   'pedidoss con items sin cargo MEDITEA
   Dim PED_S_CARGO%()
   ReDim PED_S_CARGO%(MSF.Rows)
   If Control("ADMIPED", "ITEMSCGO") = "SI" Then
      For i& = 1 To MSF.Rows - 1
        If MSF.TextMatrix(i&, 0) = "*" And XVALO(MSF.TextMatrix(i&, 7)) <= 0.005 Then
           PED_S_CARGO%(i&) = 1
        End If
      Next i&
   End If

   strSQL = "Select * from PedDeta where NroPed=" & NROPE&
   Dim PedDeta As ADODB.Recordset
   Set PedDeta = New ADODB.Recordset
26 On Error Resume Next
   PedDeta.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
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
          NLISST% = CVS(Mid$(X$, 85, 4))
          'If Left$(UCase$(ARCHIPRE$), 5) = "LIPRE" Then
             'PLISTA = PRELISTA#(ARCHIPRE$, CIXI%)
             PLISTA = PRELISTA#(SAFETEXT$(NLISST%), CIXI%)
             Call REPLA(X$, MKS$(PLISTA), 39, 4)
          'End If
       End If
       '
       NCOTI& = 0
       NUPRESU& = CVS(Mid$(X$, 125, 4))
       If NUPRESU& > 0 Then
          NCOTI& = XVALO(VALOBADA("PRESUVEN", "NROCOTI", "NROPRES = " & NUPRESU&, "NOMESS"))
       End If

       OBSERVAX$ = MSF.TextMatrix(U&, 9)
       With PedDeta
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CODIEMPR = CodiEmp%
         On Error GoTo 0
           !NROPED = NROPE&
           !FECHEMIS = CVDAT(FEMI%)
           !NROCLIE = NCLIEN
           !RasoClie = Left$(rasocli$(NCLIEN), 64)
           !NROOCOM = TRIM$(Text10.TEXT)
             FENTRESOL% = CVI(Mid$(X$, 63, 2))
           !FeSolEnt = CVDAT(FENTRESOL%)
           !Status = STATUPC%              ' Status  de Aprobación
           !ListPre = XVALO(TEXT21(0).TEXT)
           !CondPag = XVALO(TEXT21(1).TEXT)
           !VENDED = XVALO(TEXT21(2).TEXT)
           !PorComVend = XVALO(Text44.TEXT)
           If Left$(UCase$(Label21(4)), 5) = "LETRA" Then
               valor% = 0
               If TRIM$(Text44) <> "" Then valor% = Asc(Text44.TEXT + "A") - 64
               !PorComVend = valor%
           End If

           !NuOrItem = IKITEM%
           !CODIINT = CIXI%
           !CODIEXT = CODEXT$(CIXI%)
           !SubCod = Mid$(X$, 109, 8)
           !Descrip = DESCRIT0$(CIXI%)
           !TCOLOR = Mid$(X$, 3, 16)
           !TTALLE = Mid$(X$, 19, 8)
           !TLargo = CVS(Mid$(X$, 27, 4))
           !TAncho = CVS(Mid$(X$, 31, 4))
           !TEspesor = CVS(Mid$(X$, 35, 4))
           !VenRep = Left$(Label11, 1)   ' Venta o Reparacion
           !MoneEmis = MONEMI%
           !TIPOCAM = XVALO(TICAMBIS)
           
           NLISST% = CVS(Mid$(X$, 85, 4))
           !lista_de_item = NLISST%
           !PreList = CVS(Mid$(X$, 39, 4))
           !CanAsig = CVS(Mid$(X$, 43, 4))
           !CanPed = CAN
           !PreUnid = CVS(Mid$(X$, 73, 4))     ' CVS(Mid$(X$, 51, 4))
           !ImpoTot = CVD(Mid$(X$, 55, 8))
           !CANTENT = 0
           !CanSaldo = CAN
           !FeUltEnt = CVDAT(0)
           !NroRemi = ""
           !TPedItem = CAVERI$
           '
           '\\\OT-8-766-19/09/08
           !ImpoAnti = 0
           !fapro_nope = CVDAT(FECHANUM("01/01/80")) 'PARA EVITAR PROBLEMAS CON FORMATO REGIONAL
           !AcuPrepa = 0
           !AcuDespa = 0
           !TipoPed = "" ' se corroboro que este CAMPO es de  TEXTO
           !finvigen = CVDAT(FECHANUM("01/01/80"))
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
           !idenlote = LOTEX1$
            NPROVX = CVS(Mid$(X$, 31, 4)) '(INT15-372)
           !Num_Prov = NPROVX '(INT15-372)
           'FIN AHP
           !NROCOTI = SAFETEXT$(NCOTI&)
           !SIN_CARGO = PED_S_CARGO%(U&) 'PEDIDO SIN CARGO  MEDITEA
           
           !ObservaItem = OBSERVAX$
           TTXYZ$ = Space$(76)
           SCG$ = "": If PED_S_CARGO%(U&) > 0 Then SCG$ = "(S/C) "
           Call REPLA(TTXYZ$, "- " + CODEXT$(CIXI%), 11, 18)
           Call REPLA(TTXYZ$, SCG$ & DESCRIT0$(CIXI%), 29, 38)
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
            VENCARPE$ = Control$("", "VENCARPE")
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
                strSQL = "select FechVent, NroPedi,nrocoti from PresuVen where NroPres = " & NUPRESU& & " and NuOrdIt= " & NUORIT%
                '
                Dim presup As ADODB.Recordset
                Set presup = New ADODB.Recordset
80              On Error Resume Next
                presup.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
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
   If UCase(FUERZAPROBACION$) = "PENDIENTE DE APROBACION POR MATERIALES SIN CARGO" Then Call REPLA(TXX1$, "Items Con Importes Que Difieren de Política S/Cargo", 11, 56): ANOTA$ = ANOTA$ + TXX1$ + Chr$(10) + Chr$(13)
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
   If Control("PEDIDO", "ACTULEN") <> "NO" Then
'      strsql = "Update Deudor12 set Lentre0_Cli ='" & TRIM$(LENTREGA(0).TEXT) & "'"
'      strsql = strsql & ", Lentre1_Cli ='" & TRIM$(LENTREGA(1).TEXT) & "'"
'      strsql = strsql & ", Lentre2_Cli ='" & TRIM$(LENTREGA(2).TEXT) & "'"
'      strsql = strsql & " where Nume_Cli = " & NCLIEN
'      FLEXCONN.Execute (FILTSQL$(strsql))
       strSQL = "select Lentre0_Cli,Lentre1_Cli,Lentre2_Cli from Deudor12 where Nume_Cli = " & NCLIEN
       Dim Cliente As ADODB.Recordset
       Set Cliente = New ADODB.Recordset
90     On Error Resume Next
       Cliente.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 90
       End If
       On Error GoTo 0
       With Cliente
            !Lentre0_Cli = TRIM$(LENTREGA(0).TEXT)
            !Lentre1_Cli = TRIM$(LENTREGA(1).TEXT)
            !Lentre2_Cli = TRIM$(LENTREGA(2).TEXT)
            .Update
       End With
       Cliente.Close
       Set Cliente = Nothing
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

Sub CARDETPEDI(MODO%, Optional NOEDITABLE%)
   '
On Error GoTo ERROR_TRAZAR
   
   
   HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC/ARCHECO(1)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle Pedido de Cliente"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top + 400))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left))
   '
   If NOEDITABLE% > 0 Then
     ATRI$ = ATRI$ + "/POSEDIT(1)=2"
     ATRI$ = ATRI$ + "/POSEDIT(3)=2"
     ATRI$ = ATRI$ + "/POSEDIT(4)=2"
     ATRI$ = ATRI$ + "/POSEDIT(5)=2"
     ATRI$ = ATRI$ + "/POSEDIT(6)=2"
     ATRI$ = ATRI$ + "/POINDEL=NO" ' DESHABILITA EL INSERTAR Y EL DEL
     NOEDITABLE% = 0
   End If
   
   Timer1.Enabled = True
   CONLOTIND$ = UCase(TRIM$(Control$("ADMIPED", "VALOIND")))  'CONTROL DE VALIDACION DE VERIFICACION DE LOTE INDIVISIBLE
10 VENCARPE$ = Control$("", "VENCARPE")
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
'         Text13.TEXT = "SIN CARGO !!!"
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
     Ci% = CVI(Left$(Y$, 2))
     If Ci% < 1 Or Ci% > NUMAS% Then
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
   If Control("ADMIPED", "VALIREPE") = "NO" Then VALIREPE% = 0
   '
   VALICODIENCLIENTE% = 0 'BANDERA PARA VALIDAR SI EL CODIGO FUE PEDIDO POR EL CLIENTE
   If Control("ADMIPED", "VACOPED") = "SI" Then VALICODIENCLIENTE% = 1
   '
   FENTREGEN% = FECHANUM(Text17)
   For U& = 1 To FIN&
     Y$ = REGLEIDO$("!CARDETPE", U&)
     Ci% = CVI(Left$(Y$, 2))
     If Ci% < 1 Or Ci% > NUMAS% Then
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
         If CVI(Left$(Y$, 2)) = Ci% Then
           If CVI(Mid$(Y$, 63, 2)) = FENTRESOL% Then
             Screen.MousePointer = 1
             TTXX$ = "Código " + TRIM$(CODEXT$(Ci%)) + "\Duplicado en el Pedido."
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
        CONDI$ = "NROCLIE = " & NCLIE & " AND CODIINT = " & Ci%
        CODPEDIDO& = CantRegistros&("PEDDETA", CONDI$, "NOMESS")
        If CODPEDIDO& < 1 Then
           TTXX$ = "Código " + TRIM$(CODEXT$(Ci%)) + "\" & DESCRIT0$(Ci%) & "\"
           TTXX$ = TTXX$ + "Nunca Fue Solicitado Por Este Cliente.\ "
           TTXX$ = TTXX$ + "¿Desea Continuar?"
           If COMALTER%(TTXX$ + "\\Sí, Continuar\No, Cancelar") > 1 Then
              GoTo 10
           End If
        End If
     End If
     X$ = REGLEIDO$("!CARDETPE", U&)
     Ci% = CVI(Left$(X$, 2))
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
       If Abs(CAN - REDON_INDIVIS(Ci%, CDbl(CAN), "VENTA")) >= 0.05 Then
         TTXX$ = "Cantidad Ajustada a Lotes Indivisibles\en Código '" + CODEXT$(Ci%) + "'"
         Call MENSERR(24, TTXX$)
         Call REPLA(X$, MKS$(REDON_INDIVIS(Ci%, CDbl(CAN), "VENTA")), 47, 4)
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
       PRECARGA = PRELISTA#(ARCHIPRE$, Ci%)
       Call REPLA(X$, MKS$(PRECARGA), 39, 4)
       Call GRAREG("!CARDETPE", X$, U&)
     End If
     PRECARGA = CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     CADECI% = DECILISTA%(XVALO(Mid$(ARCHIPRE$, 6, 3)))
     VAMINI1 = 0.5 / (10 ^ CADECI%)
     If PRECARGA < VAMINI1 Or PRECARGA > 9999999999.99 Then
       If MODO% < 2 Then
         If ECOPRECIO% >= 0 Then
           Call REPLA(X$, MKS$(0), 39, 4)
           Call GRAREG("!CARDETPE", X$, U&)
           MODO% = 2
           If COMALTER%("Artículo '" + TRIM$(CODEXT$(Ci%)) + "' sin Precio.\\Corregir\Continuar") <> 2 Then
              NOFRESHPRE% = 1
              GoTo 10
           End If
         End If
       End If
     End If
     '
     CANTPEDIDA# = CVS(Mid$(X$, 47, 4))
     If Control$("PEDIDO", "VERISTO") = "SI" Then
        If DESCOMPO%(Ci%) > 0 Then
          Call AREXPLO(Ci%, 1, 1)
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
        If CANTPEDIDA# > STOCKACT(Ci%) Then
           If CODSINSTO$ <> "" Then
              CODSINSTO$ = CODSINSTO$ & " ; " & CODEXT(Ci%)
              Else
              CODSINSTO$ = CODEXT(Ci%)
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
   'SE COMENTA SEGUN MODIFICACION PRESUPUESTO 13/04/18
'   If InStr(UCase(EPAC$), "PREMOTEC") > 0 Then
'      Call APP_PREMOTEC
'   End If
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
   If Control("PEDIDO", "MODODESC") = "MONTO" And XVALO(TEXT21(0)) = 1 Then
       Call DESCUMONTO(Me)
   End If
   '
   List1.Refresh
'   Call REMPLAGRID(Me, GRID2, MARCONOTA, Label22, List1)
  'pedidoss con items sin cargo MEDITEA
   If Control("ADMIPED", "ITEMSCGO") = "SI" Then
      For i& = 1 To MSF.Rows - 1
        If XVALO(MSF.TextMatrix(i&, 7)) <= 0 Then
           'Call COLOREAR_GRILLA_SOLAFILA(MSF, vbYellow, GLOBALFRM_FILA&, , , 1)
           Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, vbYellow, i&)
           MSF.TextMatrix(i&, 0) = "*"
        Else
           Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, vbWhite, i&)
           MSF.TextMatrix(i&, 0) = ""
        End If
      Next i&
   End If
   
   On Error Resume Next
   Command14.SetFocus
   On Error GoTo 0
   Screen.MousePointer = 1
   '
   Exit Sub
ERROR_TRAZAR:
  MsgBox MENSAX$ + " ERROR DENTRO DE CARDETPEDI " & MODO% & " " & Err.Description
  Call FINAL("")
End Sub

Sub CARLISPE()
   List1.Clear
   'MSF.Rows = 1
   FIN& = ULTREG&("!CARDETPE")
   '
   'ESTO PARA CONTROLAR LAS OBSERVACIONES
   If Text28.Visible = True Then
    Dim MATRIZ$()
    ReDim MATRIZ$(1)
    For J& = 1 To FIN&
       X$ = REGLEIDO$("!CARDETPE", J&)
       Ci% = CVI(Left$(X$, 2))
       COD$ = TRIM$(CODEXT$(Ci%))
       For i% = 1 To MSF.Rows - 1
          If MSF.TextMatrix(i%, 2) = COD$ Then
             If UBound(MATRIZ$) < J& Then ReDim Preserve MATRIZ(J&)
             MATRIZ(J&) = MSF.TextMatrix(i%, 9)
             Exit For
          End If
        Next i%
    Next J&
    End If
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     TTXX$ = Space$(96)
     Ci% = CVI(Left$(X$, 2))
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = CVS(Mid$(X$, 39, 4))
       DLX% = DECILISTA(XVALO(Mid$(ARCHIPRE$, 6, 3)))    ' DECIMALES DE LA LISTA DE PRECIOS
     If Abs(PORDESCUIN) > 0 Then
       PREUNI = ROUND((PRECARGA * (100 - PORDESCUIN) / 100), DLX%)
       PREUNI_SINRED = PRECARGA * (100 - PORDESCUIN) / 100
     Else
       PREUNI = ROUND(PRECARGA, DLX%)
       PREUNI_SINRED = PRECARGA
     End If
     
     TOT# = ROUND(CDbl(CAN * PREUNI))
     '
     If CAN <= 9999 Then
       Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 1, 2), 1, 6)
     Else
       Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
     End If
     '
     Call REPLA(TTXX$, CODEXT$(Ci%), 8, 15)
     FEX = FENTRESOL%
     If ECOPRECIO% < 0 Then
          Call REPLA(TTXX$, DESCRIT0$(Ci%), 24, 52)
          Call REPLA(TTXX$, FECHATEX$(FEX), 78, 8)
        Else
          Call REPLA(TTXX$, DESCRIT0$(Ci%), 24, 22)
          Call REPLA(TTXX$, Left$(FECHATEX$(FEX), 5), 47, 5)
     End If
     '
     PRELIX = PRELISTA#(ARCHIPRE$, Ci%)
     If PRELIX < PRECARGA Then PRELIX = PRECARGA
     If PRELIX < PREUNI Then PRELIX = PREUNI
       DLX% = DECILISTA(XVALO(Mid$(ARCHIPRE$, 6, 3)))     ' DECIMALES DE LA LISTA DE PRECIOS
     PORDESCUIN = 0: If PRELIX > 0.005 Then PORDESCUIN = 100 * (PRELIX - PREUNI_SINRED) / PRELIX
     ' ESTO SE MODIFICO POR PROBLEMA CON CARGA DE PEDIDOS DE GABAL
     ' PORDESCUIN = 0: If PRELIX > 0.005 Then PORDESCUIN = 100 * (PRELIX - PREUNI) / PRELIX
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PRELIX), 3, 8, DLX%, 2), 52, 8)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 61, 5)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, DLX%, 2), 66, 8)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 2, 2), 76, 10)
     '
     
'     List1.AddItem TTXX$
     MSF.Rows = U& + 1
     MSF.TextMatrix(U&, 1) = TRIM$(Mid$(TTXX$, 1, 6))
     MSF.TextMatrix(U&, 2) = TRIM$(Mid$(TTXX$, 8, 15))
     MSF.TextMatrix(U&, 3) = TRIM$(Mid$(TTXX$, 24, 22))
     MSF.TextMatrix(U&, 4) = TRIM$(Mid$(TTXX$, 47, 5))
     MSF.TextMatrix(U&, 5) = TRIM$(FORMATNUM$(PRELIX, "F12." & SAFETEXT$(DLX%)))
     MSF.TextMatrix(U&, 6) = TRIM$(FORMATNUM$(PORDESCUIN, "F7.2"))
     MSF.TextMatrix(U&, 7) = TRIM$(FORMATNUM$(PREUNI, "F12." & SAFETEXT$(DLX%)))
     MSF.TextMatrix(U&, 8) = TRIM$(FORMATNUM$(TOT#, "F14.2"))
     If Text28.Visible = True Then
        'TEXT28 DE LAS OBSERVACIONES SI NO ESTA VISIBLE ES POR QUE NO ESTA CONFIGURADO
        On Error Resume Next
        MSF.TextMatrix(U&, 9) = MATRIZ(U&)
        On Error GoTo 0
     End If
     '
   Next U&
   Text28 = ""
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
   
   'MODIFICACION DE EYM ITEM CON 2 IVAS (CARGA EL VECTOR CON LOS CODIGOS)
   For i& = 1 To 256
       CI_VECTOR_TIVA(i&) = 0
   Next i&

   For U& = 1 To ULTREG&("!CARDETPE")
     Y$ = REGLEIDO$("!CARDETPE", U&)
     CI_VECTOR_TIVA(U&) = CVI(Left$(Y$, 2))
   Next U&
   Call ACTUALIZAVECTORDEIVA
   
   'FIN MODIFICACION DE EYM ITEM CON 2 IVAS
   
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CIXXI% = CVI(Left$(X$, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 99 Then PORDESCUIN = 0
     If CAN > 0.05 Then
       PRELI = CVS(Mid$(X$, 39, 4))
       CDX% = DECILISTA(LIPRACTI%)
       If Abs(PORDESCUIN) > 0 Then
          PREUNI = ROUND(PRELI * (100 - PORDESCUIN) / 100, CDX%)
       Else
          PREUNI = ROUND(PRELI, CDX%)
       End If
       TOT# = ROUND(CDbl(CAN * PREUNI), 2)
       SUTOTPED# = SUTOTPED# + TOT#
       SUIVA# = SUIVA# + (Int(TOT# * PORCEIVA(CIXXI%) + 0.5)) / 100
       '\\\OT-05-0560-WAR-03/08/05///
       TODE1# = TOT# * COEF1#
       GRIVA% = GRUPIVA%(CIXXI%)
       IMPTASIVA#(GRIVA%) = IMPTASIVA#(GRIVA%) + (Int(TODE1# * PORCEIVA(CIXXI%) + 0.5)) / 100
       Label13 = CSTRING$(MKD$(IMPTASIVA#(0)), 3, 13, 2, 2): SkinLabel13.Caption = Label13
       Label14 = CSTRING$(MKD$(IMPTASIVA#(1)), 3, 13, 2, 2): SkinLabel14.Caption = Label14
       Label20 = CSTRING$(MKD$(IMPTASIVA#(2)), 3, 13, 2, 2): SkinLabel20.Caption = Label20
       '\\\OT-05-0560-WAR-FIN///
       Call REPLA(X$, MKS$(PREUNI) + MKD$(TOT#), 51, 12)
       Call GRAREG("!CARDETPE", X$, U&)
       TOTITEMS% = TOTITEMS% + 1
     End If
   Next U&
   '
   Label18 = TRIM$(Str$(TOTITEMS%))
   '
   Total(0).TEXT = CSTRING$(MKD$(SUTOTPED#), 3, 13, 2, 2)
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
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 13, 2, 2)
   '
   'AGREGADO PARA GABAL DESDE ACA
   DESC1 = (XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)) * COEFDESCU(Label31) / 100
   Text141 = CSTRING$(MKS$(DESC1), 3, 13, 2, 2)
   
   TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT) - XVALO(Text141) ' ESTA ULTIMA RESTA GABAL
   Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 13, 2, 2)
   '
   On Error Resume Next
   TOTIVA(0).TEXT = TRIM$(CSTRING$(MKD$(APLIVA% * SUIVA# * XVALO(Total(2).TEXT) / XVALO(Total(0).TEXT)), 4, 13, 2, 2))
   On Error GoTo 0
   TOTIVA(2).TEXT = TRIM$(CSTRING$(MKD$(CDbl(XVALO(TOTIVA(0).TEXT) * 0.5 * RNI%)), 4, 13, 2, 2))
   TOTIVA(3).TEXT = TRIM$(CSTRING$(MKD$(CDbl(XVALO(TOTIVA(0).TEXT) + XVALO(TOTIVA(2).TEXT))), 4, 13, 2, 2))
   '
   Call CALPERIB
   '
   Total(3).TEXT = TRIM$(CSTRING$(MKD$(CDbl(XVALO(Total(2).TEXT) + XVALO(TOTIVA(3).TEXT)) + XVALO(Text15)), 4, 13, 2, 2))
   '
'   Call T_Espera(5)
End Sub

Private Sub TOTIVA_Change(Index As Integer)
   'If Index = 3 Then
   '  Total(3).TEXT = Str$(XVALO(Total(2).TEXT) + XVALO(TOTIVA(3).TEXT))
   'End If
End Sub

Sub GENERODES()
   '
   If Control$("ADMIPED", "GENDESPA") = "REVISTOCK" Then
     Call GENODESPA(XVALO(Text11), NODES&, 1, OKKY%)
     If OKKY% = 1 Then
        Call PRIODESPA(NODES&)
     End If
   End If
   '
End Sub
'
Sub CALPERIB()

    Dim Percepciones As Double
    
    With PERCEIIBB
        .Cliente = XVALO(Text1)
        .BaseCABA = 0
        .BaseBuenosAires = XVALO(Total(2))
        .BaseMisiones = XVALO(Total(2)): If PercepcionIBMisiones_BienDeUso = 1 Then .BaseMisiones = 0
        .BaseTucuman = 0
        .Moneda = MONEMI%
        
        If .Bases > 0.005 Then
            Call .CalcularPercepciones
            Percepciones = .TotalPercepciones
            Text15 = FORMATNUM$(Percepciones, "F12.2")
        End If
    End With
        
'   '
'   '
'   'SE MODIFICA PARA PASARLE EL VALOR DE VARIDOC POR QUE LO ESTABA PASANDO SIN VALOR
'   NCLIEN = XVALO(Text1)
'   PIVADOC% = PIVACLI%(NCLIEN)
'   '
'   'SE ASIGNA VALOR DE VARIDOC
'   VARIDOC% = 1
'   If PIVADOC% = 0 Or PIVADOC% = 3 Or PIVADOC% = 4 Or PIVADOC% = 6 Then
'        VARIDOC% = 2
'      ElseIf PIVADOC% = 5 Or PIVADOC% = 7 Then
'        VARIDOC% = 3
'        If LETREXPO$ = "A" Then
'          VARIDOC% = 1
'        End If
'   End If
'   '
'   TASAX = ALIPERBRU#(NCLIEN, MINIPER#, VARIDOC%)
'   '
'   Text15 = ""
'   BIMPON# = XVALO(Total(2))
'   If BIMPON# * TICAMONE(MONEMI%) >= MINIPER# Then
'      PERSALTA# = BIMPON# * (TASAX / 100)
'      Text15 = FORMATNUM$(PERSALTA#, "F11.2")
'   End If
'   '

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
     TASAA = XVALO(Control$("PERIBRU", "TASA-A"))
     TASAB = XVALO(Control$("PERIBRU", "TASA-B"))
     TASAEXP = XVALO(Control$("PERIBRU", "TASA-EXP"))
     MINIPER = XVALO(Control$("PERIBRU", "MINIPER"))
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
   FECORTE% = FECHANUM(Control$("PEDIDO", "FECORTE"))
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
           NUIT% = !NUMITEM
           CODEX$ = !CodProd
           CIXI% = CODINT%(CODEX$)
           ' validar articulo
           If CIXI% < 1 Then
              MERRO$ = "Codigo '" & CODEX$ & "' Inexistente o no Válido."
              Call ADDMERR(MERRO$)
           End If
           ' fin validar articulo
           '
           CANTU = !Cantidad
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
        CANTU = !Cantidad
        '
        If NUPED& >= PRIPEDI& Then
           NUIT% = !NUMITEM
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
     If Left$(Control$("PEDCLI", "MODOAPRO"), 6) = "DIRECT" Then STATUPC% = 0
     '
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CODIEMPR = CodiEmp%
     On Error GoTo 0
     !NROPED = NUPED&
     '!Referencia = REFE$
     !FECHEMIS = CVDAT(FPED%)
     !NROCLIE = NCLIE
     !RasoClie = rasocli$(NCLIE)
     !NROOCOM = "OE " + TRIM$(Str$(NUPED&))
     !DOMIENT = LUENTRECLI$(NCLIE, 0)
     !LOCAENT = LUENTRECLI$(NCLIE, 1)
     !TranspEnt = LUENTRECLI$(NCLIE, 2)
     !NTranspo = 0
     !NSucent = 0
     !Observa = ""
     !FeSolEnt = CVDAT(FPED%)
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
               If Abs(!CanPed - CANTU) < 0.05 Then
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
         !RasoClie = rasocli$(NCLIE)
         !NROOCOM = "OE " + TRIM$(Str$(NUPED&))
         !FeSolEnt = CVDAT(FPED%)
         !Status = STATUPE%   'idem anterior.
         !ListPre = LIPRCLI%(NCLIE)
         !CondPag = CPAGCLI%(NCLIE)
         !VENDED = VENDCLI%(NCLIE)
         !PorComVend = 0
         !NuOrItem = NUIT%
         !CODIINT = CIXI%
         !CODIEXT = CODEXT$(CIXI%)
         !SubCod = ""
         !Descrip = DESCRIT0$(CIXI%)
         !TCOLOR = ""
         !TTALLE = ""
         !TLargo = 0
         !TAncho = 0
         !TEspesor = 0
         !VenRep = "V"
         !MoneEmis = 1
         !TIPOCAM = 1
         !PreList = 0 ' CVS(Mid$(X$, 39, 4))
         !CanAsig = 0
         !CanPed = CANTU
         !PreUnid = 0  ' CVS(Mid$(X$, 51, 4))
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
          If !cod_inte = CICONJ% Then
            If Abs(!CANTINGRE - CANTU) < 0.05 Then
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
     If Control$("ADMIPED", "NOVALSTK") <> "SI" Then
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
     If Control$("ADMIPED", "NOVALSTK") <> "SI" Then
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
           If UCase$(Control("ADMIPED", "OFDIRECT")) = "SI" Then GoTo 50
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
         Lx& = 1
1999     If EXISTE%(LUCOM$ + "TRPED" & SAFETEXT$(Lx&) & ".TXT") = 1 Then
           Lx& = Lx& + 1
           GoTo 1999
         End If
         TXTRAZA$ = TXTRAZA$ + " ULTIMO " & ACONEC$
         NARCHI2% = FILEOPEN%(LUCOM$ + "TRPED" & SAFETEXT$(Lx&) & ".TXT", 2, 256) ' ABRE EL ARCHIVO DE ESCRITURA
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
    On Error GoTo ERROR_TRAZAR
    
    NCLIEN = XVALO(Text1.TEXT)
    CONDICION$ = " NroClie =  " & NCLIEN & " AND Status < 8 AND CantEnt < (CanPed - 0.005) AND  Status <> 3 "
    '
    JJ& = 0
    FRMZELECHO.msf2.Clear
    FRMZELECHO.MSF1.Clear
   Campos$ = "N° Pedido/F10;Razon Social/A32;Emitido/D8"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , 10000)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , 10000)
   '
   FRMZELECHO.Caption = "Visualizar Pedidos Pendientes de: " & Text2
   '
   SQLX$ = "SELECT NROPED,RaSoClie,FechEmis FROM PEDDETA WITH(NOLOCK) WHERE " & CONDICION$
   SQLX$ = SQLX$ + " GROUP BY NROPED,RASOCLIE,FECHEMIS ORDER BY NROPED "
   Set RS = READSET(SQLX$)
   MENSAX$ = "1"
   With RS
      Do While Not .EOF
            JJ& = JJ& + 1
            TEXTOGRILLA$ = XVALO(!NROPED) & "|" & SAFETEXT$(!RasoClie) & "|" & FECHATEX$(CVINT(!FECHEMIS))
            Call CARGA_GRILLA_GENERICO(FRMZELECHO, FRMZELECHO.msf2, Campos$, TEXTOGRILLA$, 1)
        .MoveNext
      Loop
   End With
   If JJ& > 0 Then
10   FRMZELECHO.Show 1
   Else
     Exit Sub
   End If
   


'    Call COPYSTRU("INDIPEP", "INDIPNN")
'    '
'
'    Call CARGALISEL("INDIPNN", "PEDDETA", "NroPed/F5.0;RaSoClie/A20;FechEmis/D8", CONDICION$, "DISTINCT")
'    '
'    If ULTREG&("INDIPNN") < 1 Then Exit Sub
'
'    '
'10  Call SELECHO("INDIPNN")
'If Check1.Value = 1 Then MsgBox "7"

'    NPX& = XVALO(VALACT1("INDIPNN"))
'    If NPX& < 1 Then Exit Sub
   '
    NPX& = XVALO(RESP_ZELE_VECT(1))
    MENSAX$ = "3"
    If NPX& < 1 Then Exit Sub

    If NPX& > 0 Then
    ''''CAMBIO PARA QUE MUESTRE DISEÑADOR DE FRM SI EXISTE
         FORIPRE$ = TRIM$(Control$("", "FORPECLI"))
         ARCSQR$ = FORIPRE$
         If EXISTE%(LUDAT$ + ARCSQR$ + ".SQR") > 0 Then
            NPSEL& = NPX&
            NCLIEN = XVALO(Text1)
            FILTX$ = TRIM$(Str$(NPSEL&)) & ";" & TRIM$(Str$(NCLIEN))
            Call STDFORM(ARCSQR$, FILTX$)
            GoTo 10
         End If
   ''''FIN CAMBIO PARA QUE MUESTRE DISEÑADOR DE FRM SI EXISTE
    
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
    Exit Sub
ERROR_TRAZAR:
  MsgBox MENSAX$ + "ERROR DENTRO DE PEDIPEND  " & " " & Err.Description
  Call FINAL("")
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
       REPRESU& = XVALO(Mid$(List2.List(KKI% - 1), 121, 8))
       If REPRESU& > 0 Then
            CONDI$ = "NUMREPA=" & REPRESU&
            CODI$ = Control$("REPARA", "CODGENER")
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
     If Control$("", "IDSECUN") = "ALIAS" Then INCLUALIAS% = 1
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
    If DERACCE%("PEDISUCU", "Pedidos de Sucursal (Neuquén o Comodoro)") < 2 Then
          Exit Sub
    End If
    
    Call CREAR_TABLA_OC_DE_SUCURSAL
    '
10  OPX% = COMALTER%("Pedido de Reposición de Sucursal\\Todas\Solo Pendientes de Generar Nota de Venta\Solo con Nota de Venta Generada")
    If OPX% < 1 Then Exit Sub
    '
    ENCA$ = "Sucursal/F10;R.Social/A32;P.Repos./F9;Emis.P.R./D8;NC.Pedido/F12;R.Social Cliente de Sucursal/A32;P.Generado/F12.0;O.C.Cte.Suc./A12"
    Call CARGA_ENCABEZADO(OCOMSUCU.GRID, ENCA$, OCOMSUCU)
    Call CARGA_ENCABEZADO(OCOMSUCU.MSF_2, ENCA$, OCOMSUCU)
    ENCA$ = "Código/F16;Descripción/A32;Can.Compra/F11.1;Sol.Ent./D8;Ped.Gen../F9;Cant.Ped./F9.1;Cant.Ent./F9.1;Saldo/F9.1"
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
    i& = 0
    With RS
      Do While Not .EOF
        P_Avellaneda& = XVALO(!NPED_GENERADO_EN_AVELLANEDA)
        If OPX% = 1 Then GoTo 30 'SI ES TODOS MUESTRA SIEMPRE
        If P_Avellaneda& <= 0 And OPX% = 2 Then GoTo 30 'SI NO TIENE PEDIDO Y SE ELIGIO ESA OPCION INGRESA
        If P_Avellaneda& > 0 And OPX% = 3 Then GoTo 30  'SI TIENE PEDIDO Y SE ELIGIO ESA OPCION ENTRA
        GoTo 50
        '
30      i& = i& + 1
        OCOMSUCU.GRID.Rows = i& + 1
        SUCUR = XVALO(!SUCURSAL)
        OCOMSUCU.GRID.TextMatrix(i&, 1) = SUCUR
        OCOMSUCU.GRID.TextMatrix(i&, 2) = rasocli$(SUCUR)
           OCOMPRAX& = XVALO(!O_COMPRA_SUC)
           P_Avellaneda& = XVALO(!NPED_GENERADO_EN_AVELLANEDA)
        OCOMSUCU.GRID.TextMatrix(i&, 3) = OCOMPRAX&
'        FECHA11 = FECHATEX$(CVINT(!FECHEMISI))
        OCOMSUCU.GRID.TextMatrix(i&, 4) = FECHATEX$(CVINT(!FECHEMISI))
        OCOMSUCU.GRID.TextMatrix(i&, 5) = XVALO(!NCLIEN_PEDIDO)
        OCOMSUCU.GRID.TextMatrix(i&, 6) = SAFETEXT$(!RAZ_CTE_SUC)
        OCOMSUCU.GRID.TextMatrix(i&, 7) = P_Avellaneda&
        OCOMSUCU.GRID.TextMatrix(i&, 8) = SAFETEXT$(!OComCliente)
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

    If i& + J& > 0 Then
       OCOMSUCU.Show 1
       GoTo 10
    Else
       Call COMUNI("No Se Detectaron Pedidos de Reposición Segun Filtro Seleccionado")
    End If
End Sub
Private Sub mnuVentasVendedores_Click()
    If DERACCE%("REPVTA", "REPORTE DE VENTAS") < 2 Then Exit Sub
    HOI% = HOY(HO$)
    FE$ = "01" + Mid$(HO$, 3)
    FF1% = FECHANUM(FE$): FF2% = HOI%
    '
    LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
    If LONCODI% < 16 Or LONCODI% > 24 Then LONCODI% = 16

    XX$ = Left$(REGLEIDO$("ECOMAST", 1), LONCODI%): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), LONCODI%): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(FF1%) + MKI$(FF2%)
    '
22  RACOD$ = OBJPLA$("DESHASCOD2", VDEF$)
    If RACOD$ = "" Then
        GoTo 99
    End If
    CI1% = CVI(Left$(RACOD$, 2)): COD1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): COD2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TMT% = 0 Or TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    FF1% = CVI(Mid$(RACOD$, 6, 2)): Desde = FETEXCOR1$(FF1%)
    FEX = FF1%: FLI1$ = FECHATEX$(FEX)
    FF2% = CVI(Mid$(RACOD$, 8, 2)): Hasta = FETEXCOR1$(FF2%)
    FEX = FF2%: FLI2$ = FECHATEX$(FEX)
    '
    '
    Rango$ = TRIM$(COD1$) + " - " + TRIM$(COD2$)
    Rango$ = Rango$ + " / " + TRIM$(TIMAT$)
    Rango$ = Rango$ + " / " + FLI1$ + " - " + FLI2$
    '
    Screen.MousePointer = 11
    CONDI$ = ""
    If CI1% > 0 And CI2% Then
       CONDI$ = CONDI$ + " AND codiext >= '" & COD1$ & "' "
       CONDI$ = CONDI$ + " AND codiext <= '" & COD2$ & "' "
    End If
    If FF1% > 0 Then
       CONDI$ = CONDI$ + " AND FECHEMIS >=  '" & Desde & "' AND FECHEMIS <= '" & Hasta & "'"
    End If

    FIN& = CantRegistros&("PEDDETA", "CANPED >0 " & CONDI$, "NOMESS")
    SQLX$ = " select nroped,fechemis,Vended, nroclie,rasoclie,codiext, canped,preunid"
    SQLX$ = SQLX$ + " ,sum(canped * preunid) as importe,MoneEmis,fesolent,codiint"
    SQLX$ = SQLX$ + " From PedDeta WITH(NOLOCK) WHERE CANPED >0 "
    SQLX$ = SQLX$ + CONDI$
    SQLX$ = SQLX$ + " group by nroped,fechemis,Vended, nroclie,rasoclie,codiext, canped,preunid,MoneEmis,fesolent,CodiInt"
    SQLX$ = SQLX$ + " order by NroPed"
    Set RS = READSET(SQLX$)
    
    Campos$ = "Pedido N°/F8;Emisión/D8;N° Vend/f7;Nomb.Vend./A20;Cta/F8;R.Social/A24;Código/A16;Descrip/A24;Cantidad/F8.1"
    Campos$ = Campos$ + "; Precio U./F10.2;Importe/F12.2;Mon/f4;Ref./A12;Sol.Ent./D8"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    J& = 0
    Set RS0 = READSET(SQLX$)
    With RS0
          Do While Not .EOF
            CIXI% = XVALO(!CODIINT)
            If TMT% = 0 Or TIMATE%(CIXI%) = TMT% Then
                J& = J& + 1
                Call TRACE(20, J&, FIN&)
    
                GRILLAGRAL.GRID.Rows = J& + 1
                GRILLAGRAL.GRID.TextMatrix(J&, 1) = XVALO(!NROPED)
                'GRILLAGRAL.GRID.TextMatrix(J&, 2) = FECHATEX$(CVINT(!FECHEMIS))
                GRILLAGRAL.GRID.TextMatrix(J&, 2) = CVDAT(CVINT(!FECHEMIS))
                VENDE% = XVALO(!VENDED)
                GRILLAGRAL.GRID.TextMatrix(J&, 3) = VENDE%
                GRILLAGRAL.GRID.TextMatrix(J&, 4) = ECOARCH$("VENDEDOR", Chr$(VENDE%))
                GRILLAGRAL.GRID.TextMatrix(J&, 5) = XVALO(!NROCLIE)
                GRILLAGRAL.GRID.TextMatrix(J&, 6) = SAFETEXT$(!RasoClie) '
                GRILLAGRAL.GRID.TextMatrix(J&, 7) = SAFETEXT$(!CODIEXT)  '
                GRILLAGRAL.GRID.TextMatrix(J&, 8) = DESCRIT0$(XVALO(CODINT%(SAFETEXT$(!CODIEXT))))
                GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(XVALO(!CanPed), "F8.1")
                GRILLAGRAL.GRID.TextMatrix(J&, 10) = FORMATNUM$(XVALO(!PreUnid), "F8.2")
                GRILLAGRAL.GRID.TextMatrix(J&, 11) = FORMATNUM$(XVALO(!IMPORTE), "F12.2")
                GRILLAGRAL.GRID.TextMatrix(J&, 12) = XVALO(!MoneEmis)
                GRILLAGRAL.GRID.TextMatrix(J&, 13) = DEMONECO$(XVALO(!MoneEmis))
                GRILLAGRAL.GRID.TextMatrix(J&, 14) = FECHATEX$(CVINT(!FeSolEnt))
            End If
            .MoveNext
            
          Loop
    End With

        RS0.Close
        Set RS0 = Nothing

        Screen.MousePointer = 1
        If GRILLAGRAL.GRID.Rows < 2 Then
          Call COMUNI("No se Detectaron Registros segun Filtros Realizados")
          Exit Sub
        End If
        TX$ = "Ventas (" & Rango$ & ")"
        GRILLAGRAL.Caption = TX$
        '*******************************************
    
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""

    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "DEPUNEGA", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
    Set obj = Nothing
'EL ANCHO Y EL CORTE VIENE DE LA FICHA TECNICA TABLA ESPECIFIC CON EL ID = 1 EL PRIMER VALOR DEL CAMPO Descripcion Y  EL CORTE TAMBIEN FIGURA EN MISMA TABLA CAMPO CORTE.
99
End Sub
Sub ReportePedidosAprobEntregas()

    If DERACCE%("RENTAPRO", "REPORTE DE PEDIDOS APROBACIONES ENTREGAS") < 2 Then Exit Sub
    '

20 Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
    If PERIO$ = "" Then Exit Sub

    DESFE = FETEXCOR1$(DESDEX%)
    HASFE = FETEXCOR1$(HASTAX%, "HASTA")

    '
    Screen.MousePointer = 11
    If DESDEX% > 0 Then
       CONDI$ = CONDI$ + " AND FECHEMIS >=  '" & Desde & "' AND FECHEMIS <= '" & Hasta & "'"
    End If

    
    SQLX$ = "  select pc.NroPed AS NROPEDIDO, pc.FechEmis AS EMISION,m.DOCUMENTO,    pd.CodiExt,M.IDENLOTE,"
    SQLX$ = SQLX$ + " pc.fapro_nope AS APROBACION ,sum(pd.CanPed) AS PEDIDO,m.ENTREGADO,m.codinterno"
    SQLX$ = SQLX$ + " ,m.ENTREGA "
    
    SQLX$ = SQLX$ + " FROM pedenca pc  with(nolock)"
    SQLX$ = SQLX$ + " inner join peddeta pd on pd.nroped= pc.nroped"
    SQLX$ = SQLX$ + " LEFT JOIN (select m.cod_inte as codinterno,  sum(m.CantSalid) AS ENTREGADO ,m.FechMov AS ENTREGA,m.DoPriLar AS DOCUMENTO,m.NuPedCli,m.Cod_Inte,m.IDENLOTE"
    
    SQLX$ = SQLX$ + " from MOVISTO m WHERE m.cantsalid > 0  group by m.NuPedCli,m.Cod_Inte,m.FechMov,m.DoPriLar,m.IDENLOTE) m on m.NuPedCli = pc.NroPed and m.Cod_Inte = pd.CodiInt"
    
'    SQLX$ = SQLX$ + " LEFT JOIN(SELECT SUM(M.CANTINGRE) AS SumaCantIngre, MAX(M.FEREMOVI) AS FechaMasReciente, O.CANPROY"
'    SQLX$ = SQLX$ + " FROM MOVISTO MOV"
'    SQLX$ = SQLX$ + " INNER JOIN ORDEFABR O ON O.IDSUBOR = MOV.IDENLOTE"
'    SQLX$ = SQLX$ + " Where MOV.CANTINGRE > 0 And O.StatuOrf = 3"
'    SQLX$ = SQLX$ + " GROUP BY O.CANPROY) MOV on MOV.NuPedCli = pc.NroPed and MOV.Cod_Inte = pd.CodiInt "
    
    SQLX$ = SQLX$ + " WHERE PC.FECHEMIS > = ' " & DESFE & "' AND PC.FECHEMIS <= '" & HASFE & "'"
    SQLX$ = SQLX$ + " group by m.codinterno,pc.NroPed,pc.FechEmis,pd.CodiExt,pc.FApro_Nope,m.entregado,pc.NroPed,m.DOCUMENTO,m.Cod_Inte,m.ENTREGA,M.IDENLOTE"
    SQLX$ = SQLX$ + " order by pc.NroPed  ,pd.CodiExt"

'ESTA CONSULTA ESTA PROBADA ES LA MISMA PERO AGREGA LA FECHA DEL CIERRE DE LA OF Y EL SUM(CANTINGRE)
'select pc.NroPed AS NROPEDIDO, pc.FechEmis AS EMISION,m.DOCUMENTO,    pd.CodiExt,M.IDENLOTE,
'pc.fapro_nope AS APROBACION ,sum(pd.CanPed) AS PEDIDO,m.ENTREGADO,m.codinterno ,m.ENTREGA  ,MOV.SumaCantIngre,MOV.FechaMasReciente
'FROM pedenca pc  with(nolock) inner join peddeta pd on pd.nroped= pc.nroped
'LEFT JOIN (select m.cod_inte as codinterno,  sum(m.CantSalid) AS ENTREGADO ,m.FechMov AS ENTREGA,m.DoPriLar AS DOCUMENTO,m.NuPedCli,m.Cod_Inte,m.IDENLOTE
'from MOVISTO m WHERE m.cantsalid > 0
'group by m.NuPedCli,m.Cod_Inte,m.FechMov,m.DoPriLar,m.IDENLOTE) m
'on m.NuPedCli = pc.NroPed and m.Cod_Inte = pd.CodiInt
'LEFT JOIN(SELECT MOV.Cod_Inte,MOV.NuPedCli,SUM(MOV.CANTSALID) AS SumaCantIngre,
'MAX(MOV.FEREMOVI) AS FechaMasReciente, O.CANPROY
'FROM MOVISTO MOV INNER JOIN ORDEFABR O ON O.IDSUBOR = MOV.IDENLOTE
'Where MOV.CANTSALID > 0 And O.StatuOrf = 3
'GROUP BY O.CANPROY,MOV.NUPEDCLI,MOV.Cod_Inte) MOV
'on MOV.NuPedCli = pc.NroPed and MOV.Cod_Inte = pd.CodiInt
'WHERE PC.FECHEMIS > = ' 01/OCT/2022' AND PC.FECHEMIS <= '25/OCT/2023 23:59:00'
'group by m.codinterno,pc.NroPed,pc.FechEmis,pd.CodiExt,pc.FApro_Nope,m.entregado,MOV.SumaCantIngre,MOV.FechaMasReciente
',pc.NroPed,m.DOCUMENTO,m.Cod_Inte,m.ENTREGA,M.IDENLOTE order by pc.NroPed  ,pd.CodiExt
    
    Set RS = READSET(SQLX$)
    Dim CTP As New Conteplast
    
    Campos$ = "Pedido N°/F8;Emisión/D8;Documento/A18;Código/A35;Aprobación/D8;Ent.Sol./A20;Cant.Pedida/F12.1;Entregado/F12.1;Fecha Entregado/A12;O.Fabric/A12"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    J& = 0
    Set RS0 = READSET(SQLX$)
    With RS0
          Do While Not .EOF
                J& = J& + 1
                GRILLAGRAL.GRID.Rows = J& + 1
                If XVALO(!NroPedido) <> XVALO(PedAnterior) Then
                   Call NEGRITA_SOLO_UNA_FILA(GRILLAGRAL.GRID, J&)
                End If
                NP& = XVALO(!NroPedido)
                GRILLAGRAL.GRID.TextMatrix(J&, 1) = NP&
                GRILLAGRAL.GRID.TextMatrix(J&, 2) = FECHATEX$(CVINT(!EMISION))
                GRILLAGRAL.GRID.TextMatrix(J&, 3) = SAFETEXT$(!DOCUMENTO)
                GRILLAGRAL.GRID.TextMatrix(J&, 4) = SAFETEXT$(!CODIEXT) '
                GRILLAGRAL.GRID.TextMatrix(J&, 5) = FECHATEX$(CVINT(!APROBACION)) '
                CONDI$ = " NROPED = " & NP& & " AND CODIINT = '" & XVALO(!codinterno) & "'"
                GRILLAGRAL.GRID.TextMatrix(J&, 6) = StringAggFlex$("PEDDETA", CONDI$, "FeSolEnt", ",")
                'GRILLAGRAL.GRID.TextMatrix(j&, 6) = FECHATEX$(CVINT(!SOLICITUD_ENTREGA)) '
                GRILLAGRAL.GRID.TextMatrix(J&, 7) = XVALO(!PEDIDO)
                GRILLAGRAL.GRID.TextMatrix(J&, 8) = XVALO(!ENTREGADO)
                GRILLAGRAL.GRID.TextMatrix(J&, 9) = SAFETEXT$(!Entrega)
                'NORFA$ = CTP.NroOrdenFabricacionSegunPedido$(NP&)

                GRILLAGRAL.GRID.TextMatrix(J&, 10) = SAFETEXT$(!idenlote)
                
                
                PedAnterior = XVALO(!NroPedido)
            .MoveNext
            
          Loop
    End With

    RS0.Close
    Set RS0 = Nothing

    Screen.MousePointer = 1
    If GRILLAGRAL.GRID.Rows < 2 Then
      Call COMUNI("No se Detectaron Registros segun Filtros Realizados")
      Exit Sub
    End If
    TX$ = "Rango Seleccionado (" & PERIO$ & ")"
    GRILLAGRAL.Caption = TX$
        '*******************************************
    
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    GRILLAGRAL.SinEfectoClickEnGrilla% = 1

    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "PEDAPREN", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
    Set obj = Nothing
'EL ANCHO Y EL CORTE VIENE DE LA FICHA TECNICA TABLA ESPECIFIC CON EL ID = 1 EL PRIMER VALOR DEL CAMPO Descripcion Y  EL CORTE TAMBIEN FIGURA EN MISMA TABLA CAMPO CORTE.
99
End Sub


Sub ReportePedidosItemCumpCEntregas()
20 Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
   If PERIO$ = "" Then Exit Sub
'NROPED,CÓDIGO CLIENTE   CLIENTE CODIGO PRODUCTO Descripcion F.PED.  NRO.OV  CANT
   DESFE = FETEXCOR1$(DESDEX%)
   HASFE = FETEXCOR1$(HASTAX%, "HASTA")
   Campos$ = "PEDIDO/F8;Nro.Cte./F7;CLIENTE/A32;CÓDIGO/A16;DESCRIPCIÓN/A36;F.PED./D8;NRO.OV./A16;Ped/F10;Ent./F10.0"
   Screen.MousePointer = 11
   Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
   Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
   CONDI$ = "FechEmis >= '" & DESFE & "' AND FechEmis <='" & HASFE & "'"
   FIN& = CantRegistros("PEDDETA", CONDI$, "NOMESS")
   
   SQLX$ = "SELECT PD.NROPED,PD.NroClie,PD.RaSoClie,PD.CodiInt,PD.CodiExt,PD.Descrip,PD.FechEmis,PD.NROOCOM,PD.CanPed, PD.cantent "
   SQLX$ = SQLX$ + " FROM PEDDETA PD WITH(NOLOCK)"
'   SQLX$ = SQLX$ + "  INNER JOIN PEDENCA PC  ON PC.NROPED = PD.NROPED "
   SQLX$ = SQLX + " WHERE PD.FechEmis >= '" & DESFE & "' AND PD.FechEmis <='" & HASFE & "'"
'   SQLX$ = SQLX + "  AND PC.Status < 8 AND PD.STATUS = 8"
   SQLX$ = SQLX + " AND PD.CANPED > PD.CANTENT   AND PD.STATUS = 8 AND PD.CANTENT >  0"
   SQLX$ = SQLX + " ORDER BY PD.NroPed"
   
   Set RS = READSET(SQLX$)
   With RS
   Do While Not .EOF
        J& = J& + 1
        Call TRACE(20, J&, FIN&)
        '
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = XVALO(!NROPED)
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = XVALO(!NROCLIE)
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = SAFETEXT$(!RasoClie)
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = SAFETEXT$(!CODIEXT)
        
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = SAFETEXT$(!Descrip)
        GRILLAGRAL.GRID.TextMatrix(J&, 6) = FECHATEX$(CVINT(!FECHEMIS))
        GRILLAGRAL.GRID.TextMatrix(J&, 7) = SAFETEXT$(!NROOCOM)
        GRILLAGRAL.GRID.TextMatrix(J&, 8) = XVALO(!CanPed)
        GRILLAGRAL.GRID.TextMatrix(J&, 9) = XVALO(!CANTENT)
        
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    
    Screen.MousePointer = 1
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    GRILLAGRAL.Caption = "Entregas de pedidos Anulados por Items Cumplidos"

    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    Set MSF = GRILLAGRAL.GRID
    If TRIM$(GRILLAGRAL.TXTBUSCAR) <> "" Then Set MSF = GRILLAGRAL.MSF_2

    If GRILLAGRAL.MENU1 > 0 Then
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "PEDCUMEN", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If


   
   

End Sub

Sub ReportePedidosTodosXRangoFecha()
20 Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
   If PERIO$ = "" Then Exit Sub
'NROPED,CÓDIGO CLIENTE   CLIENTE CODIGO PRODUCTO Descripcion F.PED.  NRO.OV  CANT
   DESFE = FETEXCOR1$(DESDEX%)
   HASFE = FETEXCOR1$(HASTAX%, "HASTA")
   Campos$ = "PEDIDO/F8;Nro.Cte./F7;CLIENTE/A32;CÓDIGO/A16;DESCRIPCIÓN/A36;F.PED./D8;NRO.OV./A16;CANT/F10"
   Screen.MousePointer = 11
   Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
   Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
   CONDI$ = "FechEmis >= '" & DESFE & "' AND FechEmis <='" & HASFE & "'"
   FIN& = CantRegistros("PEDDETA", CONDI$, "NOMESS")
   
'   SQLX$ = "SELECT PD.NROPED,PD.NroClie,PD.RaSoClie,PD.CodiInt,PD.CodiExt,PD.Descrip,PD.FechEmis,PD.NROOCOM,PD.CanPed "
'   SQLX$ = SQLX$ + " FROM PEDDETA PD WITH(NOLOCK)"
'   SQLX$ = SQLX$ + "  INNER JOIN PEDENCA PC  ON PC.NROPED = PD.NROPED "
'   SQLX$ = SQLX + " WHERE PD.FechEmis >= '" & DESFE & "' AND PD.FechEmis <='" & HASFE & "'"
'   SQLX$ = SQLX + "  AND PC.Status < 8 AND PD.STATUS < 8"
'   SQLX$ = SQLX + " ORDER BY PD.NroPed"

    SQLX$ = "SELECT PD.NROPED, PD.NroClie, PD.RaSoClie, PD.CodiInt, PD.CodiExt, PD.Descrip, PD.FechEmis, PD.NROOCOM, PD.CanPed,pd.Status,pd.CantEnt "
    SQLX$ = SQLX$ + "FROM PEDDETA PD WITH(NOLOCK) "
'    SQLX$ = SQLX$ + "INNER JOIN PEDENCA PC ON PC.NROPED = PD.NROPED "
    SQLX$ = SQLX$ + " WHERE (PD.STATUS < 8 OR (PD.STATUS >= 8 AND PD.NROPED IN (SELECT NROPED FROM PEDDETA WHERE CANTENT > 0.5))) "
    SQLX$ = SQLX$ + " AND PD.FechEmis >= '" & DESFE & "' AND PD.FechEmis <= '" & HASFE & "' "
    SQLX$ = SQLX$ + " ORDER BY PD.NROPED"
   
   Set RS = READSET(SQLX$)
   With RS
   Do While Not .EOF
        J& = J& + 1
        Call TRACE(20, J&, FIN&)
        '
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = XVALO(!NROPED)
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = XVALO(!NROCLIE)
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = SAFETEXT$(!RasoClie)
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = SAFETEXT$(!CODIEXT)
        
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = SAFETEXT$(!Descrip)
        GRILLAGRAL.GRID.TextMatrix(J&, 6) = FECHATEX$(CVINT(!FECHEMIS))
        GRILLAGRAL.GRID.TextMatrix(J&, 7) = SAFETEXT$(!NROOCOM)
        GRILLAGRAL.GRID.TextMatrix(J&, 8) = XVALO(!CanPed)

        
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    
    Screen.MousePointer = 1
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""

    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    Set MSF = GRILLAGRAL.GRID
    If TRIM$(GRILLAGRAL.TXTBUSCAR) <> "" Then Set MSF = GRILLAGRAL.MSF_2

    If GRILLAGRAL.MENU1 > 0 Then
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "ASIGNASEC", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If


   
   

End Sub


