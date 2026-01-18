VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPMASPR07 
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Proveedores - Configuración"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   10365
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
   ScaleHeight     =   5250
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   2760
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   5600
      Left            =   9500
      ScaleHeight     =   5565
      ScaleWidth      =   1170
      TabIndex        =   56
      Top             =   0
      Width           =   1200
      Begin VB.CommandButton Command6 
         Caption         =   "IVA"
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
         Picture         =   "OPMASPR07.frx":0000
         TabIndex        =   70
         ToolTipText     =   "Configuración Funcionamiento Retención de IVA"
         Top             =   2970
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Height          =   690
         Left            =   120
         Picture         =   "OPMASPR07.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Listado Proveedores para I.Brutos"
         Top             =   3660
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Picture         =   "OPMASPR07.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Configuracion Funcionamiento Percepciones IB Pcia de Bs.As"
         Top             =   2280
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command5 
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
         Picture         =   "OPMASPR07.frx":0EEE
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Configuracion Funcionamiento Retenciones de Ganancias"
         Top             =   1590
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
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
         Left            =   120
         Picture         =   "OPMASPR07.frx":1330
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Guardar "
         Top             =   900
         Width           =   650
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
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
         Left            =   120
         Picture         =   "OPMASPR07.frx":163A
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
   End
   Begin VB.CommandButton BOTSEL 
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
      Index           =   7
      Left            =   9030
      TabIndex        =   51
      Top             =   420
      Width           =   175
   End
   Begin VB.TextBox COTEXT 
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
      Index           =   7
      Left            =   8505
      TabIndex        =   52
      Text            =   " "
      Top             =   420
      Width           =   570
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C7D9CC&
      Caption         =   "VALORES PREDETERMINADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5030
      Left            =   5040
      TabIndex        =   23
      Top             =   105
      Width           =   4320
      Begin VB.TextBox DETEXT 
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
         Height          =   285
         Index           =   7
         Left            =   150
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   735
         Width           =   4035
      End
      Begin VB.TextBox DETEXT 
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
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1575
         Width           =   4035
      End
      Begin VB.TextBox DETEXT 
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
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   2415
         Width           =   4035
      End
      Begin VB.CommandButton BOTSEL 
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
         Index           =   6
         Left            =   3990
         TabIndex        =   31
         Top             =   4140
         Width           =   175
      End
      Begin VB.TextBox DETEXT 
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
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   3255
         Width           =   4035
      End
      Begin VB.CommandButton BOTSEL 
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
         Index           =   4
         Left            =   3990
         TabIndex        =   53
         Top             =   2835
         Width           =   175
      End
      Begin VB.TextBox COTEXT 
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
         Left            =   3465
         TabIndex        =   54
         Text            =   " "
         Top             =   2835
         Width           =   570
      End
      Begin VB.CommandButton BOTSEL 
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
         Index           =   0
         Left            =   3990
         TabIndex        =   24
         Top             =   1995
         Width           =   175
      End
      Begin VB.TextBox COTEXT 
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
         Index           =   6
         Left            =   1785
         TabIndex        =   33
         Text            =   " "
         Top             =   4140
         Width           =   2250
      End
      Begin VB.TextBox DETEXT 
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
         Height          =   285
         Index           =   6
         Left            =   150
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   4560
         Width           =   4035
      End
      Begin VB.CommandButton BOTSEL 
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
         Index           =   2
         Left            =   3990
         TabIndex        =   27
         Top             =   1155
         Width           =   175
      End
      Begin VB.TextBox COTEXT 
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
         Index           =   2
         Left            =   3465
         TabIndex        =   26
         Text            =   " "
         Top             =   1155
         Width           =   570
      End
      Begin VB.TextBox COTEXT 
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
         Index           =   0
         Left            =   3465
         TabIndex        =   25
         Text            =   "  "
         Top             =   1995
         Width           =   570
      End
      Begin VB.Label Label5 
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
         Height          =   225
         Left            =   525
         TabIndex        =   55
         Top             =   2960
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Madre"
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
         TabIndex        =   34
         Top             =   4245
         Width           =   1305
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Proveedor"
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
         Left            =   1050
         TabIndex        =   30
         Top             =   420
         Width           =   2325
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría de Ingresos Brutos"
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
         Left            =   630
         TabIndex        =   29
         Top             =   1260
         Width           =   2745
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
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
         Left            =   2205
         TabIndex        =   28
         Top             =   2100
         Width           =   1185
      End
   End
   Begin VB.TextBox COTEXT 
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
      Index           =   1
      Left            =   6090
      TabIndex        =   37
      Text            =   " "
      Top             =   945
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox DETEXT 
      BackColor       =   &H00E0E0E0&
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
      Index           =   1
      Left            =   5775
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   2625
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.CommandButton BOTSEL 
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
      Left            =   6615
      TabIndex        =   35
      Top             =   945
      Visible         =   0   'False
      Width           =   175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C7D9CC&
      Caption         =   "CONTROLES OPERATIVOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   105
      TabIndex        =   18
      Top             =   2720
      Width           =   4750
      Begin VB.Frame Frame12 
         BackColor       =   &H00C7D9CC&
         Caption         =   "FORMATO DE NÚMERO DE CUIT"
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
         Left            =   160
         TabIndex        =   63
         Top             =   1560
         Width           =   4430
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Suprime Validación de Formato (Argentina)"
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
            Height          =   255
            Left            =   105
            TabIndex        =   64
            Top             =   300
            Width           =   3960
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C7D9CC&
         Caption         =   "ALTA DE NUEVA CUENTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   160
         TabIndex        =   19
         Top             =   300
         Width           =   4425
         Begin VB.Frame Frame6 
            BackColor       =   &H00C7D9CC&
            Caption         =   "PRF"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   105
            TabIndex        =   41
            Top             =   315
            Width           =   810
            Begin VB.TextBox Text4 
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
               Left            =   105
               TabIndex        =   42
               Top             =   315
               Width           =   620
            End
         End
         Begin VB.TextBox Text3 
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
            Height          =   285
            Left            =   3600
            MaxLength       =   4
            TabIndex        =   22
            Top             =   600
            Width           =   645
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Primer Número Libre Mayor a"
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
            Height          =   750
            Left            =   2160
            TabIndex        =   21
            Top             =   315
            Width           =   1425
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Número Siguiente al Ultimo"
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
            Height          =   750
            Left            =   940
            TabIndex        =   20
            Top             =   315
            Value           =   -1  'True
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C7D9CC&
      Caption         =   "TABLAS DE VALIDACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4750
      Begin VB.CommandButton Command20 
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
         Left            =   3570
         TabIndex        =   1
         Top             =   285
         Width           =   175
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C7D9CC&
         Caption         =   "OPCIONES DE TRABAJO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   160
         TabIndex        =   16
         Top             =   1125
         Width           =   4425
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Grupos de Proveedores"
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
            Index           =   9
            Left            =   2520
            TabIndex        =   15
            Top             =   735
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Cuentas Contables"
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
            Index           =   8
            Left            =   2520
            TabIndex        =   14
            Top             =   315
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Condiciones de Pago"
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
            Index           =   6
            Left            =   320
            TabIndex        =   12
            Top             =   315
            Width           =   1590
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Provincias y Zonas de Venta"
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
            Index           =   5
            Left            =   320
            TabIndex        =   13
            Top             =   735
            Width           =   1800
         End
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Abrir"
         Height          =   700
         Left            =   3885
         Picture         =   "OPMASPR07.frx":1784
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   285
         Width           =   690
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
         Height          =   285
         Left            =   160
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   705
         Width           =   3590
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   285
         Width           =   1905
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   2
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command21 
            Caption         =   "Análisis A-B-C"
            Height          =   435
            Left            =   2205
            TabIndex        =   8
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Estadísticas"
            Height          =   435
            Left            =   210
            TabIndex        =   7
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command23 
            Caption         =   "Críticos"
            Height          =   435
            Left            =   2205
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command24 
            Caption         =   "Consignación"
            Height          =   435
            Left            =   210
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Negativos"
            Height          =   435
            Left            =   2205
            TabIndex        =   4
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command26 
            Caption         =   "Inventarios"
            Height          =   435
            Left            =   210
            TabIndex        =   3
            Top             =   420
            Width           =   1590
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TABLA"
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
         Left            =   -735
         TabIndex        =   17
         Top             =   390
         Width           =   2325
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "NO VISIBLES"
      Height          =   1170
      Left            =   3720
      TabIndex        =   38
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
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
         Index           =   3
         Left            =   2835
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   0
         Width           =   3615
      End
      Begin VB.TextBox COTEXT 
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
         Index           =   3
         Left            =   1995
         TabIndex        =   47
         Text            =   " "
         Top             =   0
         Width           =   570
      End
      Begin VB.TextBox COTEXT 
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
         Index           =   5
         Left            =   1995
         TabIndex        =   46
         Text            =   " "
         Top             =   840
         Width           =   570
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   5
         Left            =   2835
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   840
         Width           =   3615
      End
      Begin VB.CommandButton BOTSEL 
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
         Index           =   3
         Left            =   2520
         TabIndex        =   44
         Top             =   0
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
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
         Index           =   5
         Left            =   2520
         TabIndex        =   43
         Top             =   840
         Width           =   175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Vendedores"
         Height          =   345
         Index           =   4
         Left            =   0
         TabIndex        =   40
         Top             =   420
         Width           =   1800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listas de Precios"
         Height          =   345
         Index           =   7
         Left            =   0
         TabIndex        =   39
         Top             =   735
         Width           =   1800
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   330
         Left            =   0
         TabIndex        =   50
         Top             =   945
         Width           =   1905
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L.Precios:"
         Height          =   225
         Left            =   210
         TabIndex        =   49
         Top             =   105
         Width           =   1710
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "OPMASPR07.frx":1A8E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPMASPR07.frx":1CC2
      TabIndex        =   66
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utiildades"
      Begin VB.Menu mnuSincrprov 
         Caption         =   "Sincronización de Código Alfa de Provincias"
      End
   End
End
Attribute VB_Name = "OPMASPR07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$

Sub CATEGANDEF()
    '
    Call COMUNI("Importante: Solo se Asignará la Categoría Elegida\a Aquellos Proveedores que No Posean Una")
    Call SELECHO("tacretg")
    categan% = XVALO(VALACT1$("tacretg"))
    If categan% <> 0 Then
            SQLX$ = "select * from proved12 with(nolock) where stat_cli>0"
            Set Rst = READSET(SQLX$)
            FIN& = CantRegistros("proved12", "stat_cli>0")
            i& = 0
            With Rst
                Do While Not .EOF
                    Call TRACE(24, i&, FIN&)
                    NPRO% = XVALO(!nume_cli)
                    RECORD$ = FILTERGETRECORD$("dargapro", 1, MKI$(NPRO%))
                    If CVI(Mid$(RECORD$, 1, 2)) > 0 Then
                        cat% = Asc(Mid$(RECORD$, 3, 1))
                        If cat% = 0 Then
                            Call REPLA(RECORD$, Chr$(categan%), 3, 1)
                            Call FILTERPUTRECORD("DARGAPRO", 1, MKI$(NPRO%), RECORD$)
                        End If
                    Else
                        Call REPLA(RECORD$, MKI$(NPRO%), 1, 2)
                        Call REPLA(RECORD$, Chr$(categan%), 3, 1)
                        Call REPLA(RECORD$, USERNAME$, 36, 20)
                        Call REPLA(RECORD$, MKI$(HOY(HOS$)), 56, 2)
                        Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 58, 2)
                        Call FILTERPUTRECORD("DARGAPRO", 1, MKI$(NPRO%), RECORD$)
                    End If
                    i& = i& + 1
                    .MoveNext
                Loop
            End With
            On Error Resume Next
            Rst.Close
            Set Rst = Nothing
            On Error GoTo 0
            Call COMUNI("Fin del Proceso de Asignación de Categoría")
    End If
    '
End Sub

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

 Sub Command16_Click()
    '
    Command16.Enabled = False
    Call COPYSTRU("GRUCOCLI", "GRUCOPRO")
    If TRIM$(MODUCONE$) <> "" Then
        If TRIM$(MODUCONE$) = "GRUCOPRO" Then
            LDT$ = LUDAT$
            Call CIERRARCH("GRUCOPRO")
            Call CIERRARCH("!GRUCOPRO")
            'Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
            Call TRALOCAL("GRUCOPRO", "")
            VTB% = VENTABU%("GRUCOPRO")
            If VTB% < 0 Then GoTo 99
            J& = 0
            For i& = 1 To ULTREG&("!GRUCOPRO")
                X$ = REGLEIDO$("!GRUCOPRO", i&)
                If CVI(Left$(X$, 2)) > 0 Then
                    J& = J& + 1
                    Call GRAREG("GRUCOPRO", X$, J&)
                End If
            Next i&
            Call SETULTREG("GRUCOPRO", J&)
            Call CIERRARCH("GRUCOPRO")
          ElseIf MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TATIMAT" Then
            Call CARTATIMA
          ElseIf MODUCONE = "CONPAGO" Then
            Call CARCONPAGOS
          ElseIf MODUCONE = "TAPROVIN" Then
            Call CARTAPROVIN
          ElseIf MODUCONE = "CUEINGAS" Then
            Call CARCUEINGAS
          Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
        End If
    End If
99  Command16.Enabled = True
    '
End Sub

Private Sub Command2_Click()
   OPXX$ = "-1"
   If Option3.Value = True Then OPXX$ = TRIM$(Str$(XVALO(Text3.TEXT)))
   Call GRACONTROL("", "PRONUPRO", OPXX$)
   Call GRACONTROL("ALTAPROV", "CODIPROV", Text4)
   '
   '\\\OT-05-0198-WAR-19/04/05-///
   VALICUIT1$ = "NO"
   If Check1.Value = 1 Then VALICUIT1$ = "SI"
   Call GRACONTROL("FORMCUIT", "SUPRIVAL", VALICUIT1$)
   '\\\OT-05-0198-WAR-FIN///
   '
   For Index = 0 To 7
       TV11% = TIPFIELD(ARCHE$(Index), 1)
       LF11% = LENFIELD(ARCHE$(Index), 1)
       ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
       VAPRE$ = "": ECO$ = ""
       If TRIM$(COTEXT(Index).TEXT) <> "" Then
          ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
          If ECO$ <> "" Then
             VAPRE$ = TRIM$(COTEXT(Index).TEXT)
          End If
       End If
       COTEXT(Index).TEXT = VAPRE$
       DETEXT(Index).TEXT = ECO$
       Call GRACONTROL("AMAPRO", ARCHE$(Index), VAPRE$)
   Next Index
   '
End Sub

Sub Command3_Click()
   Command3.Enabled = False
   If DERACCE%("ACPAIBRU", "Actualizar Parámetros Ingresos Brutos") = 2 Then
   
        OpcionesIB% = COMALTER%("Jurisdiccion Ingresos Brutos:\\Buenos Aires\C.A.B.A\Misiones\Tucuman")
        If OpcionesIB% < 1 Or OpcionesIB% > 4 Then GoTo 99
        
        If OpcionesIB% = 1 Then
            If ULTREG&("TACATIBP") < 5 Then
              Call GRAREG("TACATIBP", Chr$(0) + AJUSTI$("EXENTO", 31) + String$(32, 0), 1)
              Call GRAREG("TACATIBP", Chr$(1) + AJUSTI$("BIENES MUEBLES", 31) + String$(32, 0), 2)
              Call GRAREG("TACATIBP", Chr$(2) + AJUSTI$("LOCACIONES SERVICIOS", 31) + String$(32, 0), 3)
              Call GRAREG("TACATIBP", Chr$(3) + AJUSTI$("EVENTUALES", 31) + String$(32, 0), 4)
              Call GRAREG("TACATIBP", Chr$(4) + AJUSTI$("PARTIDAS ESPECIALES", 31) + String$(32, 0), 5)
              Call CIERRARCH("TACATIBP")
            End If
            '
            Call TRALOCAL("TACATIBP", "")
            VTB% = VENTABU%("TACATIBP")
            If VTB% >= 0 Then
              For U& = 1 To ULTREG&("!TACATIBP")
                XX$ = REGLEIDO$("!TACATIBP", U&)
                Call REPLA(XX$, Chr$(U& - 1), 1, 1)
                Call GRAREG("TACATIBP", XX$, U&)
              Next U&
              Call CIERRARCH("TACATIBP")
            End If
            
        ElseIf OpcionesIB% = 2 Then
            If ULTREG&("TACAIBCA") < 5 Then
              Call GRAREG("TACAIBCA", Chr$(0) + AJUSTI$("EXENTO", 31) + String$(32, 0), 1)
              Call GRAREG("TACAIBCA", Chr$(1) + AJUSTI$("BIENES MUEBLES", 31) + String$(32, 0), 2)
              Call GRAREG("TACAIBCA", Chr$(2) + AJUSTI$("LOCACIONES SERVICIOS", 31) + String$(32, 0), 3)
              Call GRAREG("TACAIBCA", Chr$(3) + AJUSTI$("EVENTUALES", 31) + String$(32, 0), 4)
              Call GRAREG("TACAIBCA", Chr$(4) + AJUSTI$("PARTIDAS ESPECIALES", 31) + String$(32, 0), 5)
              Call CIERRARCH("TACAIBCA")
            End If
            '
            Call TRALOCAL("TACAIBCA", "")
            VTB% = VENTABU%("TACAIBCA")
            If VTB% >= 0 Then
              For U& = 1 To ULTREG&("!TACAIBCA")
                XX$ = REGLEIDO$("!TACAIBCA", U&)
                Call REPLA(XX$, Chr$(U& - 1), 1, 1)
                Call GRAREG("TACAIBCA", XX$, U&)
              Next U&
              Call CIERRARCH("TACAIBCA")
            End If
            
         ElseIf OpcionesIB% = 3 Then
            Call COPYSTRU("TACATIBP", "TACAIBMI")
            '
            If ULTREG&("TACAIBMI") < 5 Then
              Call GRAREG("TACAIBMI", Chr$(0) + AJUSTI$("No disponible", 31) + String$(32, 0), 1)
              Call GRAREG("TACAIBMI", Chr$(1) + AJUSTI$("Exento", 31) + String$(32, 0), 2)
              Call GRAREG("TACAIBMI", Chr$(2) + AJUSTI$("Pago de comisiones o análogos", 31) + String$(32, 0), 3)
              Call GRAREG("TACAIBMI", Chr$(3) + AJUSTI$("C.M. c/sede en otras provincias", 31) + String$(32, 0), 4)
              Call GRAREG("TACAIBMI", Chr$(4) + AJUSTI$("C.M. c/sede en Misiones", 31) + String$(32, 0), 5)
              Call GRAREG("TACAIBMI", Chr$(5) + AJUSTI$("Bs/Serv no insc. en Misiones", 31) + String$(32, 0), 6)

              Call CIERRARCH("TACAIBMI")
            End If
            '
            Call TRALOCAL("TACAIBMI", "")
            VTB% = VENTABU%("TACAIBMI")
            If VTB% >= 0 Then
              For U& = 1 To ULTREG&("!TACAIBMI")
                XX$ = REGLEIDO$("!TACAIBMI", U&)
                Call REPLA(XX$, Chr$(U& - 1), 1, 1)
                Call GRAREG("TACAIBMI", XX$, U&)
              Next U&
              Call CIERRARCH("TACAIBMI")
            End If
            
         Else
            OpcionIBTucuman% = COMALTER%("Opciones de Configuración:\\Tabla de Categorías\Código Provincia")
            If OpcionIBTucuman% = 1 Then
                Call COPYSTRU("TACATIBP", "TACAIBTU")
                '
                If ULTREG&("TACAIBTU") < 3 Then
                  Call GRAREG("TACAIBTU", Chr$(0) + AJUSTI$("Inscripto en padrón", 31) + String$(32, 0), 1)
                  Call GRAREG("TACAIBTU", Chr$(1) + AJUSTI$("No inscripto en padrón 5%", 31) + String$(32, 0), 2)
                  Call GRAREG("TACAIBTU", Chr$(2) + AJUSTI$("No inscripto en padrón 7%", 31) + String$(32, 0), 3)
    
                  Call CIERRARCH("TACAIBTU")
                End If
                '
                Call TRALOCAL("TACAIBTU", "")
                VTB% = VENTABU%("TACAIBTU")
                If VTB% >= 0 Then
                  For U& = 1 To ULTREG&("!TACAIBTU")
                    XX$ = REGLEIDO$("!TACAIBTU", U&)
                    Call REPLA(XX$, Chr$(U& - 1), 1, 1)
                    Call GRAREG("TACAIBTU", XX$, U&)
                  Next U&
                  Call CIERRARCH("TACAIBTU")
                End If
            
            Else
                Call SELECHO("PROVINC")
                CodigoProvincia$ = TRIM$(VALACT1$("PROVINC"))
                PROVINCIA$ = TRIM$(ECOARCH$("PROVINC", CodigoProvincia$))
                If CodigoProvincia$ = "" Then Exit Sub
                
                Call GRACONTROL("IIBBTUCU", "CODPROVI", CodigoProvincia$)
                
                Call COMUNI("Se Ha Configurado Correctamente\'" & CodigoProvincia$ & " - " & PROVINCIA$ & "'")
            
            End If
         End If
   End If
99 Command3.Enabled = True
   '
End Sub

Sub Command4_Click()
   Command4.Enabled = False
   Call FILESORT("DARIBPRO", "DARIBPRJ", 1, 1, "ASC")
   Screen.MousePointer = 11
   FIN& = ULTREG&("DARIBPRJ")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("DARIBPRJ", U&)
     NCLIX% = CVI(Left$(XX$, 2))
     CATIBX% = XVALO(CATIBRU$((-1) * NCLIX%))
     NIBR$ = NIBRCLI$((-1) * NCLIX%)
     Call REPLA(XX$, MKI$(CATIBX%), 65, 2)
     Call REPLA(XX$, NIBR$, 67, 16)
     If CATIBX% > 0 Then
       If TRIM$(NIBR$) = "" Then
         Call REPLA(XX$, "Falta Nro. Inscripción", 83, 24)
       End If
     End If
     Call GRAREG("DARIBPRJ", XX$, U&)
   Next U&
   Call CIERRARCH("DARIBPRJ")
   Screen.MousePointer = 1
   '
   Call FINAL("*LIDARIBP")
   Command4.Enabled = True
End Sub

Sub Command5_Click()
  Command5.Enabled = False
  '\\\OT-05-0635-WAR-22/09/05///
  RTA% = ALTERNA%("Tabla de Categorias\Escalas de Retención\Asignar Categoría por Default")
  If RTA% = 1 Then
      Call CARTARET
    ElseIf RTA% = 2 Then
      Call ESCALRET
    ElseIf RTA% = 3 Then
      Call CATEGANDEF
  End If
  Command5.Enabled = True
  '\\\OT-05-0635-WAR-22/09/05///
End Sub

Sub ESCALRET() 'ESCALAS DE RETENCION DE GANANCIAS
  '
  Call SELECHO("TACRETG")
  If TRIM$(VALACT1$("TACRETG")) <> "" Then
    '3.- SI NO TIENE MINIMO O TASA => AFUERA
    NUCA% = XVALO(VALACT1$("TACRETG")) 'CATEGORIA ELEGIDA
    X$ = FILTERGETRECORD$("TACRETG", 1, Chr(NUCA%))
    DESCRICAT$ = Mid$(X$, 2, 31)
    MINEX1 = CVD(Mid$(X$, 33, 8))
    TASA1 = CVS(Mid$(X$, 41, 4))
    If MINEX1 <= 0.005 And TASA1 <= 0.005 Then
      Call COMUNI("No se puede Cargar Escala a Categoria Exento")
      Exit Sub
    End If
    '6.- CARGA LA TABLA C/REGISTROS Q' MACHEEN CON LA CATEGORIA SELECIONADA
    Call FILTERFILE("ESCALGAN", "!ESCALGAN", 1, MKI$(NUCA%))
    '7.-GRABACION
    VTB% = VENTABU%("ESCALGAN/NC/TITUPAN=Escala de Retencion - " & DESCRICAT$ & "")
    If VTB% >= 0 Then
      For J& = 1 To ULTREG&("!ESCALGAN")
        YY$ = REGLEIDO$("!ESCALGAN", J&)
        Call REPLA(YY$, MKI$(NUCA%), 1, 2)
        Call GRAREG("!ESCALGAN", YY$, J&)
      Next J&
      Call FILTERUPDATE("ESCALGAN", "!ESCALGAN", 1, MKI$(NUCA%))
    End If
    '
  End If
  '
End Sub

Sub CARTARET()
    Call TRALOCAL("TACRETG", "")
     VTB% = VENTABU%("TACRETG")
     If VTB% >= 0 Then
       For U& = 1 To ULTREG&("!TACRETG")
         XX$ = REGLEIDO$("!TACRETG", U&)
         ' Graba el numero de orden automaticamente
         Call REPLA(XX$, Chr$(U& - 1), 1, 1)
         Call GRAREG("TACRETG", XX$, U&)
       Next U&
       Call CIERRARCH("TACRETG")
     End If
End Sub

Sub Command6_Click()

    If DERACCE%("PARETIVA", "Actualizar Parametros Retencion IVA") = 2 Then
      
      MinimoImponible# = XVALO(Control$("RETENIVA", "MINIMPO"))
      CuentaContable$ = TRIM$(Control$("RETENIVA", "CUECONTA"))
      VDEF$ = MKS$(MinimoImponible#) & Space(12) & CuentaContable$
      '
10    OBX$ = OBJPLA$("SETRETENCIONIVA", VDEF$)
      If OBX$ <> "" Then
        MinimoImponible# = CVS(Mid$(OBX$, 1, 4))
        If MinimoImponible# < 0.005 Then
           Call MENSERR(24, "Falta Completar el Mínimo Imponible.")
           GoTo 10
        End If
        Call GRACONTROL$("RETENIVA", "MINIMPO", TRIM$(FORMATNUM$(MinimoImponible#, "F10.2")))
        
        CuentaContable$ = TRIM$(Mid$(OBX$, 17, 12))
        If CuentaContable$ = "" Or ECOARCH("ECUECON", CuentaContable$) = "" Then
           Call MENSERR(24, "Cuenta Contable no Válida.")
           GoTo 10
        End If
        Call GRACONTROL$("RETENIVA", "CUECONTA", CuentaContable$)
      End If
      '
    End If

End Sub

'
Private Sub Form_Load()
  '
  Screen.MousePointer = 11
  
  UTILIZA_SKIN% = 1
  Call APLICACIONSKINS(Me)

  Call CARTAVALI
  ARCHE$(7) = "GRUCOPRO"   ' ESTO ES PORQUE ANTES NO TRAÍA LA TABLA
                           ' DE PROVEEDORES SINO LA DE CLIENTES
  Screen.MousePointer = 1
  '
  TABLA$(0) = "TATIMAT"       ' tabla de tipos de material
  TABLA$(1) = "*UNIMED"
  TABLA$(2) = "TADEPOSI"
  TABLA$(3) = ""       ' grupos contables por articulo
  TABLA$(4) = "*VENDEDOR"
  TABLA$(5) = "TAPROVIN"
  TABLA$(6) = "CONPAGO"
  TABLA$(7) = "LPREAC07"
  TABLA$(8) = "CUEINGAS"
  '
  TABLA(9) = "GRUCOPRO" 'NUEVO
  '
  PRONUCLI$ = Control$("", "PRONUPRO")
  If XVALO(PRONUCLI$) >= 0 Then
    Option3.Value = True
    Text3 = PRONUCLI$
  End If
  '
  Text4 = Left$(TRIM$(Control("ALTAPROV", "CODIPROV")), 4)
  '
  For Index = 0 To 7
    VAPRE$ = TRIM$(Control$("AMAPRO", ARCHE$(Index)))
    COTEXT(Index) = "": DETEXT(Index) = ""
    If VAPRE$ <> "" Then
      TV11% = TIPFIELD(ARCHE$(Index), 1)
      LF11% = LENFIELD(ARCHE$(Index), 1)
      ARGU$ = CBIN$(VAPRE$, TV11%, LF11%)
      ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
      If ECO$ <> "" Then
          COTEXT(Index).TEXT = VAPRE$
          DETEXT(Index).TEXT = ECO$
        Else
          COTEXT(Index).TEXT = ""
          DETEXT(Index).TEXT = ""
       End If
    End If
    DETEXT(Index).Refresh
    DETEXT(Index).TabStop = False
  Next Index
  '
'  Frame3.Visible = True
  ' '\\\OT-05-0198-WAR-19/04/05-///
  If Control$("FORMCUIT", "SUPRIVAL") = "SI" Then
      Check1.Value = 1
  End If
  
  '''\\\OT-05-0198-WAR-FIN///
End Sub

Private Sub mnuSincrprov_Click()
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 0
        VENTANA.HabilitaInsertar = 0
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = ""
        VENTANA.Inicializar = 1
        VENTANA.CampEditables = "4"
        VENTANA.Campox = "Cod.Oficial/A10;Provincia/A32;ID/F8;CodFlexoft/A10"

        Dim obj As New RECORXET
        Set rs1 = obj.RS_Tabla_Provincias_Config()
        Call Carga_MSF_Recordset(rs1, "CodOficial/A10;Provincia/A32;ID/F8;CodFlex/A10", VENTABNEW.MSF)
        Set obj = Nothing
        rs1.Close
        Set rs1 = Nothing
        '*****************************************************'
        VENTABNEW.Show 1
        If VENTABNEW.APROBADO > 0 Then
            Dim MSF As msFlexGrid
            Set MSF = VENTABNEW.MSF
    
    
            APARTIR% = 1
            SQLX$ = "SELECT * FROM PROVINCIA ORDER BY ID"
            Dim rs As ADODB.Recordset
            Set rs = New ADODB.Recordset
25          On Error Resume Next
            rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 25
            End If
              With rs
                Do While Not .EOF
                  IDX = XVALO(!ID)
                  For REG& = APARTIR% To MSF.Rows - 1
                   If XVALO(MSF.TextMatrix(REG&, 3)) = IDX Then
                      !CODFLEX = TRIM$(Left(MSF.TextMatrix(REG&, 4), 24))
                      .Update
                      APARTIR% = REG& + 1
                      GoTo 30
                    End If
                   Next REG&
30                 .MoveNext
                 Loop
              End With
            
              rs.Close
              Set rs = Nothing
        End If
        Set VENTANA = Nothing


End Sub

'
Private Sub Option1_Click(Index As Integer)
    Option1(Index).Value = True
    Text2.TEXT = Option1(Index).Caption
    MODUCONE$ = TABLA$(Index)
    TETABLA$ = MODUCONE$
    If Left$(TETABLA$, 1) = "#" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(#)"
      ElseIf Left$(TETABLA$, 1) = "*" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(*)"
    End If
    Text1.TEXT = TETABLA$
    '
    For k% = 4 To 9
       If k% <> Index Then
          Option1(k%).Value = False
       End If
    Next k%
    On Error Resume Next
    Command16.SetFocus
    On Error GoTo 0
End Sub

Sub Option1_DblClick(Index As Integer)
    '
    If TRIM$(MODUCONE$) = "" Then
       Call Form_Load
       MODUCONE$ = TABLA$(Index)
    End If
    '
    If TRIM$(MODUCONE$) <> "" Then
        If TRIM$(MODUCONE$) = "GRUCOPRO" Then
            Call COPYSTRU("GRUCOCLI", "GRUCOPRO") 'NUEVO
            LDT$ = LUDAT$
            Call CIERRARCH("GRUCOPRO")
            Call CIERRARCH("!GRUCOPRO")
            'Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
            Call TRALOCAL("GRUCOPRO", "")
            VTB% = VENTABU%("GRUCOPRO")
            If VTB% < 0 Then Exit Sub
            J& = 0
            For i& = 1 To ULTREG&("!GRUCOPRO")
                X$ = REGLEIDO$("!GRUCOPRO", i&)
                If CVI(Left$(X$, 2)) > 0 Then
                    J& = J& + 1
                    Call GRAREG("GRUCOPRO", X$, J&)
                End If
            Next i&
            Call SETULTREG("GRUCOPRO", J&)
            Call CIERRARCH("GRUCOPRO")
          ElseIf MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TATIMAT" Then
            Call CARTATIMA
          ElseIf MODUCONE = "CONPAGO" Then
            Call CARCONPAGOS
          ElseIf MODUCONE = "TAPROVIN" Then
            Call CARTAPROVIN
          ElseIf MODUCONE = "CUEINGAS" Then
            Call CARCUEINGAS
          Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
        End If
    End If
End Sub


Private Sub Option2_Click()
   Option2.Value = True
   Text3.TEXT = ""
End Sub

Private Sub Option3_Click()
   Option3.Value = True
End Sub

Private Sub Text1_GotFocus()
   Command16.SetFocus
End Sub

Private Sub Text2_gotfocus()
   Command16.SetFocus
End Sub

Private Sub Text3_Click()
   Option3.Value = True
End Sub

Private Sub COTEXT_DblClick(Index As Integer)
    Call SELECHO(ARCHE$(Index))
    If TRIM$(VALACT1$(ARCHE$(Index))) <> "" Then
        COTEXT(Index).TEXT = VALACT1$(ARCHE$(Index))
        COTEXT(Index).Refresh
        DETEXT(Index).TEXT = VALACT2$(ARCHE$(Index))
        DETEXT(Index).Refresh
        DETEXT(Index).TabStop = False
    End If
    COTEXT(Index).SetFocus
End Sub
'
Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO(ARCHE$(Index))
    If TRIM$(VALACT1$(ARCHE$(Index))) <> "" Then
        COTEXT(Index).TEXT = VALACT1$(ARCHE$(Index))
        COTEXT(Index).Refresh
        DETEXT(Index).TEXT = VALACT2$(ARCHE$(Index))
        DETEXT(Index).Refresh
        DETEXT(Index).TabStop = False
    End If
    COTEXT(Index).SetFocus
End Sub
'
Private Sub COTEXT_LOSTFOCUS(Index As Integer)
    If TRIM$(COTEXT(Index).TEXT) <> "" Then
        TV11% = TIPFIELD(ARCHE$(Index), 1)
        LF11% = LENFIELD(ARCHE$(Index), 1)
        ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
        ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
        If ECO$ <> "" Then
             COTEXT(Index).TEXT = TRIM$(COTEXT(Index).TEXT)
             DETEXT(Index).TEXT = ECO$
             DETEXT(Index).Refresh
             DETEXT(Index).TabStop = False
           Else
             DETEXT(Index).TEXT = ""
             Exit Sub
        End If
      Else
        DETEXT(Index).TEXT = ""
    End If
End Sub
'

Private Sub COTEXT_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
       If TRIM$(COTEXT(Index).TEXT) <> "" Then
          TV11% = TIPFIELD(ARCHE$(Index), 1)
          LF11% = LENFIELD(ARCHE$(Index), 1)
          ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
          ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
          If ECO$ <> "" Then
               DETEXT(Index).TEXT = ECO$
               DETEXT(Index).Refresh
               DETEXT(Index).TabStop = False
             Else
               DETEXT(Index).TEXT = ""
               COTEXT(Index).SetFocus
          End If
       End If
    End If
End Sub

Private Sub DETEXT_CLICK(Index As Integer)
    COTEXT(Index).SetFocus
End Sub


