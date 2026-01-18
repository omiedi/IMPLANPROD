VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CotPVent 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Planilla de Ventas"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   13620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   13620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   6975
      Left            =   10080
      TabIndex        =   77
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7100
      Index           =   1
      Left            =   0
      TabIndex        =   34
      ToolTipText     =   "Son 4 fuelles x bolsa"
      Top             =   2040
      Width           =   10050
      Begin VB.CommandButton Command12 
         Caption         =   "."
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
         Left            =   6720
         TabIndex        =   74
         Top             =   840
         Width           =   175
      End
      Begin VB.TextBox TXTRODILLO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   4740
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox TXTMAXIMOCOLORES 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2115
         Width           =   420
      End
      Begin VB.TextBox TXTLARGODECORTE2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   4740
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox TXTLARGODECORTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   4740
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   66
         TabStop         =   0   'False
         Text            =   "Máximo de Colores"
         Top             =   2115
         Width           =   1845
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Adjuntar Archivos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   120
         TabIndex        =   59
         Top             =   3840
         Width           =   9015
         Begin VB.CommandButton Command45 
            Caption         =   "Ver"
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
            Left            =   7080
            TabIndex        =   63
            Top             =   360
            Width           =   850
         End
         Begin VB.CommandButton Command43 
            Caption         =   "Quitar"
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
            Left            =   7965
            TabIndex        =   62
            Top             =   360
            Width           =   850
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   360
            Width           =   5565
         End
         Begin VB.CommandButton Command47 
            Caption         =   "Agregar"
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
            Left            =   240
            TabIndex        =   60
            Top             =   360
            Width           =   850
         End
      End
      Begin VB.TextBox TXTMEDFUELLE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   6980
         TabIndex        =   13
         Top             =   2595
         Width           =   1440
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   17
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "Medida Fuelle(x1)Cm."
         Top             =   2595
         Width           =   1800
      End
      Begin VB.CommandButton Command11 
         Caption         =   "."
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
         Left            =   6720
         TabIndex        =   14
         Top             =   3360
         Width           =   175
      End
      Begin VB.CommandButton Command10 
         Caption         =   "."
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
         Left            =   6720
         TabIndex        =   12
         Top             =   1995
         Width           =   175
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
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
         Left            =   6720
         TabIndex        =   11
         Top             =   1605
         Width           =   175
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
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
         Left            =   1920
         TabIndex        =   8
         Top             =   3360
         Width           =   175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "."
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
         Left            =   1920
         TabIndex        =   7
         Top             =   2595
         Width           =   175
      End
      Begin VB.CommandButton Command6 
         Caption         =   "."
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
         Left            =   1920
         TabIndex        =   5
         Top             =   1605
         Width           =   175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "."
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
         Left            =   1920
         TabIndex        =   4
         Top             =   885
         Width           =   175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
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
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   8
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "Observaciones:"
         Top             =   4800
         Width           =   1800
      End
      Begin VB.TextBox TXTOBSERVA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   5160
         Width           =   9015
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "Material (Bolsa)"
         Top             =   360
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "Ancho total Cm."
         ToolTipText     =   "Se contempla la medida total del ancho incluyendo fuelles"
         Top             =   885
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   4
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "Impresión"
         Top             =   1605
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   5
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "Cant. Colores"
         ToolTipText     =   "hasta  4 colores"
         Top             =   2115
         Width           =   1200
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   6
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "Fuelle"
         Top             =   2595
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   9
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "Microperforado"
         Top             =   3360
         Width           =   1800
      End
      Begin VB.TextBox TXTMATERIALBOLSA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox TXTANCHOTOTAL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   885
         Width           =   1455
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   10
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "Producto"
         Top             =   360
         Width           =   1800
      End
      Begin VB.TextBox TXTIMPRESION 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1605
         Width           =   1455
      End
      Begin VB.TextBox TXTPRODUCTO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   6980
         TabIndex        =   9
         ToolTipText     =   "Se debe ingresar el futuro producto interno, en la bolsa ej.:  Arroz, Maiz, etc."
         Top             =   360
         Width           =   2640
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   13
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "Alto Util Cm."
         ToolTipText     =   "Se contempla desde boca abierta hasta costura de hilo"
         Top             =   840
         Width           =   1800
      End
      Begin VB.TextBox TXTALTOUTIL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   6980
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Se contempla desde boca abierta hasta costura de hilo"
         Top             =   840
         Width           =   1440
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   14
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "Pleno"
         Top             =   1605
         Width           =   1800
      End
      Begin VB.TextBox TXTPLENO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   6980
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1605
         Width           =   1440
      End
      Begin VB.TextBox TXTCANTCOLORES 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   3840
         TabIndex        =   6
         Top             =   2115
         Width           =   495
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   16
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "Cant. Caras"
         Top             =   1995
         Width           =   1800
      End
      Begin VB.TextBox TXTCANTCARAS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   6980
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1995
         Width           =   1440
      End
      Begin VB.TextBox TXTFUELLE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2595
         Width           =   1455
      End
      Begin VB.TextBox TXTMICROPERFORADO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   7
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "Liner"
         Top             =   3360
         Width           =   1800
      End
      Begin VB.TextBox TXTLINER 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   6980
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3360
         Width           =   1440
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rodillo"
         Height          =   285
         Left            =   6840
         TabIndex        =   73
         Top             =   4740
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Largo de corte2"
         Height          =   285
         Left            =   4320
         TabIndex        =   71
         Top             =   4740
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Largo de corte"
         Height          =   285
         Left            =   2040
         TabIndex        =   70
         Top             =   4740
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.TextBox Text18 
      Height          =   495
      Left            =   6480
      TabIndex        =   33
      Text            =   "Text18"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Height          =   495
      Left            =   6480
      TabIndex        =   32
      Text            =   "Text17"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame lblNuevaCotizacion 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   0
      TabIndex        =   16
      Top             =   -120
      Width           =   10050
      Begin VB.CommandButton Command13 
         Caption         =   "Cotizar"
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
         Left            =   8760
         TabIndex        =   76
         Top             =   1680
         Width           =   1155
      End
      Begin VB.CommandButton Command39 
         Caption         =   "Solic. Cotiz."
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
         Left            =   7080
         TabIndex        =   75
         Top             =   1680
         Width           =   1155
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   2040
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2940
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   "Ref/Código/Cliente"
         Top             =   1680
         Width           =   1800
      End
      Begin VB.TextBox TXTNUMPLAVENTA 
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
         Height          =   330
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1200
         Width           =   900
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   18
         Left            =   7040
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "Planilla de Ventas N°"
         Top             =   1200
         Width           =   1920
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   0
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Fecha"
         Top             =   1200
         Width           =   900
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   12
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Cantidad de Pedido"
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Cantidad Mínima x 1 Pedido 15.000 unidades."
         Top             =   1200
         Width           =   1035
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   11
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "R.Social"
         Top             =   240
         Width           =   840
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   31
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Cliente"
         Top             =   240
         Width           =   840
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
         Left            =   4950
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Limpiar Casillero"
         Top             =   690
         Width           =   4200
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   15
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Vendedor"
         Top             =   690
         Width           =   900
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   4470
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   690
         Width           =   400
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
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
         Left            =   4260
         TabIndex        =   24
         Top             =   690
         Width           =   175
      End
      Begin VB.CommandButton CMDFECHA 
         Caption         =   "."
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
         Left            =   4260
         TabIndex        =   23
         Top             =   1200
         Width           =   175
      End
      Begin VB.TextBox TXTFECHA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   4470
         TabIndex        =   22
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TXTRAZOCLI 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   3000
         TabIndex        =   0
         ToolTipText     =   "Ingresar Razon Social Completa ej.: Arcor S.A. , Nestlé S.R.L."
         Top             =   240
         Width           =   6150
      End
      Begin VB.TextBox TXTCAPACIDAD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   2040
         TabIndex        =   1
         Top             =   690
         Width           =   1035
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   40
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Capacidad en Kgs."
         Top             =   690
         Width           =   1600
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
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
         TabIndex        =   20
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox TEXT1 
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
         Height          =   330
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   900
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   9240
         Picture         =   "CotPVent.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   240
         Width           =   650
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label LBLNUEVAPVTA 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva Planilla de Venta"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1185
         Left            =   4080
         TabIndex        =   58
         Top             =   960
         Visible         =   0   'False
         Width           =   3255
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "CotPVent.frx":014A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CotPVent.frx":037E
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuNewPV 
         Caption         =   "Nueva Planilla de Ventas"
      End
      Begin VB.Menu mnuAPVE 
         Caption         =   "Abrir Planilla de Ventas Existente"
      End
      Begin VB.Menu MnuGuardar 
         Caption         =   "Guardar Planilla Activa"
      End
      Begin VB.Menu mnuPrintPV 
         Caption         =   "Imprimir Planilla Activa"
      End
   End
   Begin VB.Menu mnuCL 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
   End
   Begin VB.Menu mnuUt 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuraciones"
   End
   Begin VB.Menu mnuAccesosDirectos 
      Caption         =   "Accesos Directos"
   End
End
Attribute VB_Name = "CotPVent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Public MODIFIC%
    Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
    Dim WinRar As String, WinZip As String
    Dim BRTEMP As Recordset
    Public NOPRESENTARMENSAJE%

Sub BLANFORMU()
        TEXT1 = ""
        TXTRAZOCLI = ""
        TXTCAPACIDAD = ""
        Text4 = ""
        TXTFECHA = ""
        Text2 = ""
        TXTMATERIALBOLSA = ""
        TXTANCHOTOTAL = ""
        TEXT5 = ""
        TXTIMPRESION = ""
        Text12 = ""
        TXTFUELLE = ""
        TXTMICROPERFORADO = ""
        Text10 = ""
        Text11 = ""
        TXTPLENO = ""
        TXTCANTCARAS = ""
        TXTLINER = ""
        TXTOBSERVA = ""
        TXTCANTCOLORES = ""
        Me.TXTPRODUCTO = ""
        Me.TXTALTOUTIL = ""
        Me.TXTMEDFUELLE = ""
        TXTMAXIMOCOLORES = ""
        TXTCANTCOLORES = ""
End Sub
Sub GUARDARARCHIVOSADJUNTOS_PVTA(NroPvta&)

   'MARCA DE BORRADO
   Call ACTUREGISTRO("DOCURELPVTA", "MARBORRA", "NROPVTA = " & NroPvta&, 1, REG&)
   '
   'RECORRO EL LIST PARA GRABAR
   For i& = 1 To Combo2.ListCount
      Z$ = Combo2.List(i& - 1)
      DESCRI$ = UCase$(TRIM$(Mid$(Z, 1, 24)))
      RUTAX$ = TRIM$(Mid$(Z, 32, 512))
      SQLX$ = "INSERT INTO  DOCURELPVTA (NROPVTA,CODINT ,Descripcion, Ruta, Marborra)"
      SQLX$ = SQLX$ + "VALUES ( " & NroPvta&
      'SQLX$ = SQLX$ + "," & CODINT(TXTCODIGO) 'CI%
      SQLX$ = SQLX$ + ",0"
      SQLX$ = SQLX$ + ",'" & DESCRI$ & "'"
      SQLX$ = SQLX$ + ",'" & TRIM$(RUTAX$) & "'"
      SQLX$ = SQLX$ + ",0"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
   Next i&
   '
   'BORRO LOS REGISTROS MARCADOS CON 1
   Call BORRAREGISTROS("DOCURELPVTA", "MARBORRA = 1 AND NROPVTA = " & NroPvta&, REGAF&, "NOMESS")
   


End Sub

Function VALIDACIONGRABACION%(Optional Grabar%)
   VALIDACIONGRABACION% = 1
   If XVALO(Me.TXTNUMPLAVENTA) < 1 Then
      Call COMUNI("Falta Número de Planilla Venta")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTNUMPLAVENTA)
      Exit Function
   End If

   If TRIM$(TXTRAZOCLI) = "" Then
      Call COMUNI("Falta Ingresar Razón Social")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTRAZOCLI)
      Exit Function
   End If

   If TRIM$(TEXT5) = "" Then
      Call COMUNI("Falta Ingresar Referencia-Código")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TEXT5)
      Exit Function
   End If

   If XVALO(TXTCAPACIDAD) <= 0 Then
      Call COMUNI("Falta Ingresar Capacidad")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTCAPACIDAD)
      Exit Function
   End If
   
   If XVALO(Me.Text4) < 1 Then
      Call COMUNI("Falta Ingresar Vendedor")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(Text4)
      Exit Function
   End If
    If XVALO(Me.Text2) < 1 Then
      Call COMUNI("Falta Ingresar Cantidad")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(Text2)
      Exit Function
   End If
   
   If TRIM$(Me.TXTMATERIALBOLSA) = "" Then
      Call COMUNI("Falta Ingresar Material")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTMATERIALBOLSA)
      Exit Function
   End If

   If XVALO(Me.TXTANCHOTOTAL) <= 0 Then
      Call COMUNI("Falta Ancho Total")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTANCHOTOTAL)
      Exit Function
   End If
   
   If TRIM$(Me.TXTIMPRESION) = "" Then
      Call COMUNI("Falta Informar si Lleva Impresión")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTIMPRESION)
      Exit Function
   End If
  
   If TRIM(Me.TXTPRODUCTO) = "" Then
      Call COMUNI("Falta Informar Producto")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTPRODUCTO)
      Exit Function
   End If
   
   If XVALO(Me.TXTALTOUTIL) <= 0 Then
      Call COMUNI("Falta Informar Alto Util")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTALTOUTIL)
      Exit Function
   End If

   If TRIM$(UCase$(Me.TXTIMPRESION)) = "SI" Then
        If XVALO(TXTCANTCOLORES) <= 0 Then
         Call COMUNI("Falta Informar Cantidad de Colores")
         VALIDACIONGRABACION% = -1
         Call PINTATEXT(TXTCANTCOLORES)
         Exit Function
        End If
        If TRIM(Me.TXTPLENO) = "" Then
         Call COMUNI("Falta Informar SI/NO Pleno")
         VALIDACIONGRABACION% = -1
         Call PINTATEXT(TXTPLENO)
         Exit Function
        End If
        If XVALO(Me.TXTCANTCARAS) <= 0 Then
         Call COMUNI("Falta Informar Cantidad de Caras")
         VALIDACIONGRABACION% = -1
         Call PINTATEXT(TXTCANTCARAS)
         Exit Function
        End If
        If XVALO(TXTMAXIMOCOLORES) < XVALO(TXTCANTCOLORES) Then
         Call COMUNI("Cantidad Informada de Colores Supera el Máximo Sugerido")
         VALIDACIONGRABACION% = -1
         Call PINTATEXT(TXTCANTCOLORES)
         Exit Function
        End If
    End If
     
    If TRIM(Me.TXTFUELLE) = "" Then
      Call COMUNI("Falta Informar SI/NO Fuelle")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTFUELLE)
      Exit Function
    End If
    
    If TRIM(UCase$(Me.TXTFUELLE)) = "SI" Then
       If XVALO(TXTMEDFUELLE) <= 0 Then
         Call COMUNI("Falta Informar Medida Fuelle")
         VALIDACIONGRABACION% = -1
         Call PINTATEXT(TXTMEDFUELLE)
         Exit Function
       End If
    End If
    
    If TRIM(Me.TXTMICROPERFORADO) = "" Then
      Call COMUNI("Falta Informar SI/NO MICROPERFORADO")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTMICROPERFORADO)
      Exit Function
    End If
    
    If TRIM(Me.TXTLINER) = "" Then
      Call COMUNI("Falta Informar SI/NO LINER")
      VALIDACIONGRABACION% = -1
      Call PINTATEXT(TXTLINER)
      Exit Function
    End If
    If Grabar% > 0 Then
      Call MnuGuardar_Click
    End If

    
End Function

Function VALIDASISEMODIFICO%()
    If XVALO(TXTNUMPLAVENTA) > 0 Then
     If MODIFIC% = 1 Then
          If COMALTER%("Se han Modificado Datos de La Planilla de Venta Activa\\Registrar Cambios\Continuar Sin Grabar") = 1 Then
            Call MnuGuardar_Click
          End If
          MODIFIC% = 0
     End If
    End If
   VALIDASISEMODIFICO% = 0
   MODIFIC% = 0
End Function

Sub GENINDIPRES(Optional INIPROFO&)
   '
   Call ABREPRESVEN
   ULTINUM& = XVALO(VALOBADA("PRESUVEN", "NroPres", "NroPres > 0 ORDER BY NroPres DESC"))
   '
   UNUMINDI& = 0
   URE& = ULTREG&("INDIPRES")
   If URE& > 0 Then
      XX$ = REGLEIDO$("INDIPRES", URE&)
      UNUMINDI& = CVS(Left$(XX$, 4))
   End If
   '
   If UNUMINDI& <> ULTINUM& Then
      Screen.MousePointer = 11
      SQLX$ = " SELECT * FROM PRESUVEN WITH(NOLOCK)"
      SQLX$ = SQLX$ + " ORDER BY NroPres ASC "
      '
      Set LISPRETEMP = New ADODB.Recordset
      'LISPRETEMP.CursorLocation = adUseClient
      Set LISPRETEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))    ', FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
      
      With LISPRETEMP
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           JJ& = 0: NPXA& = 0: FIN& = ULTINUM&
           REBLA$ = REGBLAN("INDIPRES")
           Do While (.EOF = False)
             NPX& = !NroPres
             NCLIE% = !NROCLIE
             CIXI% = !CODIINT
             FEX = CVINT(!FECHEMIS)
             XXX$ = REBLA$
             '
             '\\\OT-06-0168-WAR-09/05/06///
             CN1$ = "": MARNOVI$ = ""
             On Error Resume Next
             CR1$ = !Cod_Rubro
             NCR1 = !NumeCRubro
             CA1$ = !Cod_Anio
             '\\\OT-06-0497-RG-21/11/06///
             MARNOVI$ = !MarNoVig
             '\\\OT-06-0497-RG-FIN///
             On Error GoTo 0
             '
             If TRIM$(CR1$) <> "" And XVALO(NCR1) > 0 Then
               NCR2$ = TRIM$(Str$(NCR1))
               While Len(NCR2$) < 4: NCR2$ = "0" + NCR2$: Wend
               CN1$ = CR1$ + "-" + NCR2$ + "-" + CA1$ + "   "
             End If
             '\\\OT-06-0168-WAR-FIN///
                NomCli$ = TRIM$(rasocli$(NCLIE%))
                If NCLIE% < 1 Then 'RASON SOCIAL DEL PROSPECT
                    NUMCLI$ = VALOBADA("PRESUVEN", "CODI_CLI", "NroPres = " & NPX& & "")
                    NomCli$ = VALOBADA("BSEGCOM", "RASO_CLI", "CODI_CLI = '" & NUMCLI$ & "'")
                End If

             YY2$ = TRIM$(AJUSTI$(CN1$ + NomCli$, 30)) + " - "
             If NPX& <> NPXA& Then
                  YY2$ = YY2$ + TRIM$(CODEXT$(CIXI%))
                Else
                  YY2$ = YY2$ + "VARIOS"
             End If
             YY2$ = YY2$ + " - " + FECHATEX$(FEX)
             '\\\OT-06-0497-RG-21/11/06///
             If MARNOVI$ <> "" Then
                YY2$ = "(N/Vig) " + YY2$
             End If
             '\\\OT-06-0497-RG-FIN///
             '
             Call REPLA(XXX$, MKS$(NPX&), 1, 4)
             Call REPLA(XXX$, YY2$, 5, 60)
             If NPX& <> NPXA& Then
               JJ& = JJ& + 1
               Call TRACE(20, JJ&, FIN&)
             End If
             Call GRAREG("INDIPRES", XXX$, JJ&)
             '
             NPXA& = NPX&
           .MoveNext
           Loop
         End If
      End With
      LISPRETEMP.Close
      Set LISPRETEMP = Nothing
      '
      Call SETULTREG("INDIPRES", JJ&)
      Call CIERRARCH("INDIPRES")
      '
   End If
   '
   Call FRESHECHO("INDIPRES")
   Screen.MousePointer = 1
   
End Sub
'
Sub ACTINDIPRES(NPX&)
   '
   Call ABREPRESVEN
   SQLX$ = " SELECT * FROM PRESUVEN WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NroPres = " & NPX&
   '
   Dim LISPRETEMP As ADODB.Recordset
   Set LISPRETEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))    ', FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   '
   With LISPRETEMP
      If Not (.BOF And .EOF) Then
        XXX$ = REGBLAN("INDIPRES")
        JJ& = 0: NPXA& = 0: YY2$ = ""
        Do While (.EOF = False)
           NCLIEX = !NROCLIE
           CIXI% = !CODIINT
           FEX = CVINT(!FECHEMIS)
           '
           CN1$ = "": MARNOVI$ = ""
           On Error Resume Next
           CR1$ = !Cod_Rubro
           NCR1 = !NumeCRubro
           CA1$ = !Cod_Anio
           '\\\OT-06-0497-RG-21/11/06///
           MARNOVI$ = !MarNoVig
           '\\\OT-06-0497-RG-FIN///
           On Error GoTo 0
           '
           If TRIM$(CR1$) <> "" And XVALO(NCR1) > 0 Then
             NCR2$ = TRIM$(Str$(NCR1))
             While Len(NCR2$) < 4: NCR2$ = "0" + NCR2$: Wend
             CN1$ = CR1$ + "-" + NCR2$ + "-" + CA1$ + "   "
           End If
           '\\\OT-06-0168-WAR-FIN///
           NomCli$ = TRIM$(rasocli$(NCLIEX))
           If NCLIEX < 1 Then 'RASON SOCIAL DEL PROSPECT
               CDXCLI$ = VALOBADA("PRESUVEN", "CODI_CLI", "NroPres = " & NPX& & "")
               NomCli$ = VALOBADA("BSEGCOM", "RASO_CLI", "CODI_CLI = '" & CDXCLI$ & "'")
           End If

           YY2$ = TRIM$(AJUSTI$(CN1$ + NomCli$, 30)) + " - "
           If NPX& <> NPXA& Then
                YY2$ = YY2$ + TRIM$(CODEXT$(CIXI%))
              Else
                YY2$ = YY2$ + "VARIOS"
           End If
           YY2$ = YY2$ + " - " + FECHATEX$(FEX)
           '\\\OT-06-0497-RG-21/11/06///
           If MARNOVI$ <> "" Then
              YY2$ = "(N/Vig) " + YY2$
           End If
           NPXA& = NPX&
        .MoveNext
        Loop
        '
        If YY2$ <> "" Then
           Call REPLA(XXX$, MKS$(NPX&), 1, 4)
           Call REPLA(XXX$, YY2$, 5, 60)
           Call FILTERPUTRECORD("INDIPRES", 1, MKS$(NPX&), XXX$)
        End If
        Call CIERRARCH("INDIPRES")
      End If
   End With
   '
End Sub

Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

'
Sub Command1_Click()
    Command1.Enabled = False
    NCLIE = ListaSeleccionClientes("", True)
    If NCLIE < 1 Then Exit Sub
    '
    NC = NCLIE
    TEXT1.TEXT = TRIM$(Str$(NC))
    '
99  Command1.Enabled = True
    '
End Sub

Private Sub Command10_Click()
   X = LISTASELECGENERAL("CTP.CARAS", , , 1)
   Me.TXTCANTCARAS = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO
End Sub

Private Sub Command11_Click()
   
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   Me.TXTLINER = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub


Private Sub Command12_Click()
   X = LISTASELECGENERAL("ALTOUTIL", , , 1)
   Me.TXTALTOUTIL = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    
    If DERACCE("COTIPLVE", "COTIZAR PLANILLA DE VENTAS ACTIVA") < 2 Then GoTo 99
'   Campos$ = "N° Planilla Vta./f16;N° Vend./F9;Ref-Código/A48;Cta./f6;Raz.Social/A32;ID Sol./f8"

    NroPvta& = NROSOLICITUDPVTA(1, 0)
    If NroPvta& < 1 Then Exit Sub
    
    NVEN% = XVALO(RESP_ZELE_VECT(2))
    REFECODI$ = RESP_ZELE_VECT(3)
    NC = XVALO(RESP_ZELE_VECT(4))
    RS$ = XVALO(RESP_ZELE_VECT(5))
    IDSOL& = XVALO(RESP_ZELE_VECT(6))
    
99  Command13.Enabled = True
   
End Sub

 Function NROSOLICITUDPVTA(Optional SOLOSINCOTIZAR%, Optional NC) As Long
    Erase RESP_ZELE_VECT
    Campos$ = "N° Planilla Vta./f16;N° Vend./F9;Ref-Código/A48;Cta./f6;Raz.Social/A32;ID Sol./f8"
    
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "NROPVTA,NUVEND,REFECLIENTE,NUCLIEN,RASO_CLI,SOLICOTIPVTA_ID"
    TABLA$ = "SOLICOTIPVTA"
    CONDI$ = "SOLICOTIPVTA_ID > 0 "
    If SOLOSINCOTIZAR% > 0 Then  'SOLO SIN COTIZAR
       CONDI$ = CONDI$ + " AND (NROCOTIZACION = '' OR NROCOTIZACION IS NULL)"
    End If
    NC = XVALO(NC)
    If NC > 0 Then
       CONDI$ = CONDI$ + " AND NUCLIEN = " & NC
    End If
    '
    '
    CONDI$ = CONDI$ + " ORDER BY SOLICOTIPVTA_ID "
    Set rs1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 15705
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(6)) > 0 Then ' 7 ES EL ID QUE SI SELECCIONO ES > 0 POR QUE ES LA CONDICION
       NROSOLICITUDPVTA = XVALO(RESP_ZELE_VECT(6))
    End If


End Function


Private Sub Command2_Click()
    '
    Call CIERRARCH("*.*")
    Unload Me
    '
End Sub

Sub Command3_Click()
    '
    Call SELECHO("VENDEDOR")
    VDEVU$ = TRIM$(VALACT1$("VENDEDOR"))
    If VDEVU$ <> "" Then
       Text4.TEXT = VDEVU$
    End If

    '
End Sub

Private Sub Command39_Click()
   Command39.Enabled = False
   NroPvta& = XVALO(TXTNUMPLAVENTA)
   If NroPvta& < 1 Then
     Call COMUNI("No hay una Planilla de Venta Activa")
     GoTo 99
   End If
   
   If CantRegistros&("P_VTA_BOLSAS", "NROPVTA = " & NroPvta&) < 1 Then
     Call COMUNI("Debe Haber uan Planilla de Ventas Grabada Correctamente\Guarde la Misma y Vuelva a Intenar")
     GoTo 99
   End If
   NOPRESENTARMENSAJE% = 1
   If VALIDACIONGRABACION%(1) < 0 Then GoTo 99 'valida datos y graba
   NOPRESENTARMENSAJE% = 0

   '
   'AVISA SI YA TIENE UNA SOLICITUD PENDIENTE DEL MISMO VENDEDOR
   CONDI$ = "NUVEND = " & XVALO(Text4)
   CONDI$ = CONDI$ + " and NROPVTA = " & NroPvta&
   CONDI$ = CONDI$ + " and COTIZAR = 1"
   If CantRegistros&("SOLICOTIPVTA", CONDI$, "NOMESS") > 0 Then
      OPX% = COMALTER%("Se Ha Detectado una Solicitud de Cotización De esta Planilla de Ventas aún Pendiente\\Cancelar\Generar Nueva Solicitud")
      If OPX% < 2 Then GoTo 99
   End If
   
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   
    SQLX$ = "INSERT INTO  SOLICOTIPVTA (NUVEND ,CANTIDAD,  REFECLIENTE, COTIZAR, NROPVTA, NROCOTIZACION  "
    SQLX$ = SQLX$ + ", NUCLIEN, CodiEmpr, MARBORRA, FECHEMISI,RASO_CLI,DESCRIMASTER)"
    SQLX$ = SQLX$ + "VALUES ( " & XVALO(Text4) 'NUVEND
    SQLX$ = SQLX$ + "," & XVALO(Text2) 'CANTIDAD
    SQLX$ = SQLX$ + ",'" & TRIM$(TEXT5) & "'" 'REFECLIENTE
    SQLX$ = SQLX$ + ",1" 'COTIZAR
    SQLX$ = SQLX$ + " ," & NroPvta 'nrocotizacion
    SQLX$ = SQLX$ + " ,0" 'NROCOTIZACION
    SQLX$ = SQLX$ + " ," & XVALO(TEXT1) 'NUCLIEN
    SQLX$ = SQLX$ + " ,0" 'CodiEmpr
    SQLX$ = SQLX$ + " ,0" 'MARBORRA
    SQLX$ = SQLX$ + " ,'" & FETEXCOR1$(HOY(HO$)) & "'" 'FECHEMISI
    SQLX$ = SQLX$ + " ,'" & TRIM$(Left$(TXTRAZOCLI, 64)) & "'" 'RASO_CLI
    SQLX$ = SQLX$ + " ,''" '
    SQLX$ = SQLX$ + ")"
    FLEXCONN.Execute (SQLX$)
    FLEXCONN.CommitTrans
    Call COMUNI("Grabación Realizado Exitosamente")
    GoTo 99
    
ERROR_GUARDAR:
        If Err <> 0 Then
           FLEXCONN.RollbackTrans
           MsgBox Err.Description
           On Error GoTo 0
           Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
           Exit Sub
        End If

99 Command39.Enabled = True
   
End Sub

Private Sub Command4_Click()
   X = LISTASELECGENERAL("MATEBOLS", , , 1)
   Me.TXTMATERIALBOLSA = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub




Private Sub Command43_Click()
    Command43.Enabled = False
    NUMPVTA& = XVALO(TXTNUMPLAVENTA)
    If NUMPVTA& < 1 Then
          Call COMUNI("Debe Haber una Planilla de Venta Activa")
          Call PINTATEXT(TXTNUMPLAVENTA)
          GoTo 99
    End If
      
    If CantRegistros&("P_VTA_BOLSAS", "NROPVTA = " & NUMPVTA&) < 1 Then
        Call COMUNI("Debe Seleccionar Una Planilla de Venta Que esté Correctamente Guardada en al Base de Datos")
        GoTo 99
    End If
    '
    If Combo2.ListIndex >= 0 Then
      OPX% = COMALTER%("Confirmación de Eliminación Archivo Adjunto Seleccionado\\Cancelar\Si Eliminar")
      If OPX% = 2 Then
       Combo2.RemoveItem (Combo2.ListIndex)
       Call GUARDARARCHADJUNTOS_PVTA(NUMPVTA&)
      End If
    Else
       Call COMUNI("Atención No Hay Ningún Archivo Seleccionado")
       GoTo 99
    End If
    
99  Command43.Enabled = True
End Sub
Sub GUARDARARCHADJUNTOS_PVTA(NUMPVTA&)
   'MARCA DE BORRADO
   Call ACTUREGISTRO("DOCURELPVTA", "MARBORRA", "NROPVTA = " & NUMPVTA&, 1, REG&)
   '
   'RECORRO EL LIST PARA GRABAR
   For i& = 1 To Combo2.ListCount
      Z$ = Combo2.List(i& - 1)
      DESCRI$ = UCase$(TRIM$(Mid$(Z, 1, 24)))
      RUTAX$ = TRIM$(Mid$(Z, 32, 512))
      SQLX$ = "INSERT INTO DOCURELPVTA (NROPVTA,CODINT ,Descripcion, Ruta, Marborra)"
      SQLX$ = SQLX$ + "VALUES ( " & NUMPVTA&
      'SQLX$ = SQLX$ + "," & CODINT(TXTCODIGO) 'CI%
      SQLX$ = SQLX$ + ",0"
      SQLX$ = SQLX$ + ",'" & DESCRI$ & "'"
      SQLX$ = SQLX$ + ",'" & TRIM$(RUTAX$) & "'"
      SQLX$ = SQLX$ + ",0"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
   Next i&
   '
   'BORRO LOS REGISTROS MARCADOS CON 1
   Call BORRAREGISTROS("DOCURELPVTA", "MARBORRA = 1 AND NROPVTA = " & NUMPVTA&, REGAF&, "NOMESS")
   


End Sub
Private Sub Command47_Click()
    Command47.Enabled = False
    NUMPVTA& = XVALO(Me.TXTNUMPLAVENTA)
    If NUMPVTA& < 1 Then
          Call COMUNI("Debe Haber una Planilla de Venta Activa")
          Call PINTATEXT(TXTNUMPLAVENTA)
          GoTo 99
    End If
      
    If CantRegistros&("P_VTA_BOLSAS", "NROPVTA = " & XVALO(TXTNUMPLAVENTA)) < 1 Then
        Call COMUNI("Debe Seleccionar Una Planilla de Venta Que esté Correctamente Guardada en al Base de Datos")
        GoTo 99
    End If

    'BOTON ADD
    CARPETAX$ = CONTROL("CARPIMAG", "UBICIMAG")
    If CARPETAX$ = "" Then
       Call COMUNI("Imposible Continuar\Falta Seleccionar Carpeta de Imagenes")
       GoTo 99
     End If
     
    On Error Resume Next
    X = GetAttr(CARPETAX$)
    If Err Then
        Call COMUNI("Imposible Continuar\Falta Seleccionar Carpeta de Imagenes o Ruta Inexistente")
        GoTo 99
    End If
    '
    On Error GoTo 0
   
    Call AGREGARRUTAIMAG_PVTA
    Call GUARDARARCHIVOSADJUNTOS_PVTA(NUMPVTA&)

   '
   
   
99 Command47.Enabled = True
End Sub


Sub AGREGARRUTAIMAG_PVTA()

      CD1.FILTER = "(*.*)|*.*|"
      CD1.ShowOpen
      Ruta$ = CD1.FileName
      If TRIM(Ruta$) = "" Then Exit Sub
      '
20    DESCRI$ = TRIM$(InputBox$("Ingrese Texto ", "Guardar Como", CODI$))
      If DESCRI$ = "" Then Exit Sub
      DESCRI$ = UCase$(DESCRI$)
      If Len(DESCRI$) > 24 Then
        Call COMUNI("Longitud de Descripción \Supera La Cantidad Máxima de Caracteres (24)")
        GoTo 20
      End If
      On Error Resume Next
      CARPETA_ORIGEN$ = Ruta$
      POS1% = InStrRev(Ruta$, "\")
      POS2% = InStrRev(Ruta$, ".")
      CARPETA_DESTINO$ = CONTROL("CARPIMAG", "UBICIMAG") & "\" & TRIM$(DESCRI$) & Mid$(Ruta$, POS2%)
      FileCopy CARPETA_ORIGEN$, CARPETA_DESTINO$
      If Err > 0 Then
       A = 1
      End If
      On Error GoTo 0
      'AGREGAR INFORMACION
      Call REPLA(TX$, DESCRI$, 1, 24)
      Call REPLA(TX$, CARPETA_DESTINO$, 32, 512)
      Combo2.AddItem TX$

End Sub

Private Sub Command45_Click()
   Command45.Enabled = False

   If Combo2.ListIndex >= 0 Then
     Linea$ = Combo2.List(Combo2.ListIndex)
     Archi$ = TRIM$(Mid$(Linea$, 25))
     If (EXISTE%(Archi$)) <> 1 Then
       Call COMUNI("Archivo Inexistente")
       GoTo 99
     Else
        Call ShellExecute(Me.hwnd, "Open", Archi$, "", "", 1)
        GoTo 99
     End If
   Else
      Call COMUNI("Atención No Hay Ningún Archivo Seleccionado")
      GoTo 99
   End If
    
99 Command45.Enabled = True
End Sub

Sub Command5_Click()
   X = LISTASELECGENERAL("ANCHOTOT", , , 1)
   Me.TXTANCHOTOTAL = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO
End Sub

 Sub Command6_Click()
    Command6.Enabled = False
    X = LISTASELECGENERAL("TABLASI_NO", , , 1)
    Me.TXTIMPRESION = TRIM$(RESP_ZELE_VECT(2))
    Call VALORESPORDEFECTO
99  Command6.Enabled = True
    '
End Sub




Private Sub Command7_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   Me.TXTFUELLE = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO

End Sub

Private Sub Command8_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   Me.TXTMICROPERFORADO = TRIM$(RESP_ZELE_VECT(2))
   Call VALORESPORDEFECTO



End Sub


Private Sub Command9_Click()
    X = LISTASELECGENERAL("TABLASI_NO", , , 1)
    Me.TXTPLENO = TRIM$(RESP_ZELE_VECT(2))
    Call VALORESPORDEFECTO
End Sub

Private Sub Form_Load()
    '
    TXTFECHA = FECHATEX$(HOY(HO$))
    Call CREATABLA_P_VTA_BOLSAS
    Call mnuNewPV_Click 'CARGA COMO NUEVA PLANILLA DE VENTAS
    LBLNUEVAPVTA.Visible = True
    '
End Sub

Private Sub mnuAPVE_Click()
   If VALIDASISEMODIFICO% < 0 Then Exit Sub
   NUS% = USNBR% Mod 100
   VENDE% = LOGVEND.VendeVsUsuario(NUS%)
   NRO = Sel_P_Vta&()
   If NRO > 0 Then
     TXTNUMPLAVENTA = ""
     LBLNUEVAPVTA.Visible = False
     TXTNUMPLAVENTA = NRO
     Call VALORESPORDEFECTO
   End If
   MODIFIC% = 0
   'SIEMPRE QUEDA HABILITADA LA POSIBILIDAD DE AGREGAR UN ARCHIVO

End Sub
Function Sel_P_Vta&()
    Erase RESP_ZELE_VECT
    Campos$ = "PlanVta./F9;Fecha/D8;Cuenta/A12;R.Social/A32;Cant./F8.1;Vend/F5;ID"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    
    
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "NROPVTA,FECHEMISI,NUCLIEN,RASOCLI,CANTIDAD,NUVEND,P_VTA_BOLSAS_ID"
    TABLA$ = "P_VTA_BOLSAS"
    CONDI$ = "NROPVTA > 0 "
    '
    CONDI$ = CONDI$ + " ORDER BY NROPVTA DESC"
    Set rs1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 12000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(7)) > 0 Then ' 7 ES EL ID QUE SI SELECCIONO ES > 0 POR QUE ES LA CONDICION
       Sel_P_Vta& = XVALO(RESP_ZELE_VECT(1))
    End If
End Function

Private Sub MnuGuardar_Click()

   If VALIDACIONGRABACION% < 0 Then Exit Sub
   ESNUEVO% = 0
   If LBLNUEVAPVTA.Visible = True Then
      ESNUEVO% = 1
   End If
   '
   SQLX$ = "SELECT * FROM P_VTA_BOLSAS "
   If ESNUEVO% < 1 Then
     SQLX$ = SQLX$ + " WHERE NROPVTA = " & XVALO(TXTNUMPLAVENTA)
   Else
     SQLX$ = SQLX$ + " WHERE NROPVTA < 1 "
   End If
   
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With RS
     If ESNUEVO% > 0 Then
         NEWCOTIZ = NumNuevaPVta&
         .AddNew
         !NroPvta = NEWCOTIZ
         Else
           NEWCOTIZ = XVALO(TXTNUMPLAVENTA)
     End If
     
        NC = XVALO(TEXT1)
        !rasocli = Left$(TXTRAZOCLI, 128)
        !NUCLIEN = NC
        !observaciones = Left$(TXTOBSERVA, 1024)
        !Cantidad = XVALO(Text2)
        !NUVEND = XVALO(Text4)
        !FECHEMISI = CVDAT(CVINT(TXTFECHA))
        !CAPKGS = XVALO(TXTCAPACIDAD)
        !IMPRESION = Left$(Me.TXTIMPRESION, 24)
        !CANTICARAS = XVALO(TXTCANTCARAS)
        !Status = 0
        !MAT_BOLSA = Left(Me.TXTMATERIALBOLSA, 128)
        !ALTOUTILCM = XVALO(TXTALTOUTIL)
        !ANCHOTOTCM = XVALO(Me.TXTANCHOTOTAL)
        !Producto = Left(Me.TXTPRODUCTO, 64)
        !CANTCOLORES = XVALO(Me.TXTCANTCOLORES)
        !Fuelle = Left(Me.TXTFUELLE, 24)
        !MICROPERFO = Left(TXTMICROPERFORADO, 24)
        !Pleno = Left(Me.TXTPLENO, 24)
        !MEDIFUELLE = XVALO(Me.TXTMEDFUELLE)
        !LINER = Left(Me.TXTLINER, 24)
        !REFECLIENTE = Left(Me.TEXT5, 64)
    .Update
   End With
   
   TXTNUMPLAVENTA = NEWCOTIZ
    
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       
    Else
      FLEXCONN.CommitTrans
      If NOPRESENTARMENSAJE% < 1 Then
         Call COMUNI("Transacción Realizada")
      End If
      Me.LBLNUEVAPVTA.Visible = False
      MODIFIC% = 0
      RS.Close
      Set RS = Nothing
    End If



End Sub


Private Sub mnuNewPV_Click()
    Call VALIDASISEMODIFICO
    Call BLANFORMU
    LBLNUEVAPVTA.Visible = True
    TXTNUMPLAVENTA = SAFETEXT$(NumNuevaPVta&)
    TXTFECHA = FECHATEX$(HOY(HO$))
    MODIFIC% = 0
    NUS% = USNBR% Mod 100
    Text4 = LOGVEND.VendeVsUsuario(NUS%)

End Sub

Private Sub mnuPrintPV_Click()
   Call PrintPlanillaVenta
End Sub

Private Sub Text1_Change()
   NCLIE = XVALO(TEXT1.TEXT)
   If NCLIE > 0 And NCLIE < 99000 Then
      Me.TXTRAZOCLI.BackColor = &HC0C0C0
      Me.TXTRAZOCLI.TEXT = rasocli$(NCLIE)
   Else
       If NCLIE = 99000 And Me.TXTRAZOCLI.BackColor = &HC0C0C0 Then
         Me.TXTRAZOCLI = ""
       Else
         Me.TXTRAZOCLI.BackColor = &H80FFFF
       End If
   End If
   MODIFIC% = 1
End Sub


Private Sub Text2_Change()
MODIFIC% = 1
End Sub

Private Sub Text2_gotfocus()
Call PINTATEXT(Text2)
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If
End Sub

Private Sub Text4_Change()
  MODIFIC% = 1
  Text3 = ECOARCH$("VENDEDOR", Chr$(XVALO(Text4.TEXT)))
  MODIFIC% = 1
End Sub


Private Sub TXTALTOUTIL_LostFocus()
    Call VALORESPORDEFECTO
End Sub


Private Sub TXTCAPACIDAD_Change()
   MODIFIC% = 1
End Sub


Private Sub TXTCAPACIDAD_GotFocus()
    Call PINTATEXT(TXTCAPACIDAD)
End Sub


Private Sub TXTCAPACIDAD_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If
End Sub


Private Sub TXTCAPACIDAD_LostFocus()
    Call VALORESPORDEFECTO
    TXTCAPACIDAD = TRIM$(FORMATNUM$(XVALO(TXTCAPACIDAD), "f9.0"))
End Sub

Sub VALORESPORDEFECTO()
    TXTLABEL(14) = "Pleno"
    Text6 = "Máximo de Colores"
    If TRIM(UCase$(TXTIMPRESION)) = "SI" Then
       TXTLABEL(14) = "Pleno"
    ElseIf TRIM(TXTIMPRESION) <> "" Then
       TXTLABEL(14) = ""
       Text6 = ""
       TXTMAXIMOCOLORES = ""
       TXTCANTCOLORES = ""
    End If
    '
    If TRIM(UCase$(TXTIMPRESION)) = "SI" Then
       TXTLABEL(16) = "Cant. Caras"
    ElseIf TRIM(TXTIMPRESION) <> "" Then
       TXTLABEL(16) = ""
    End If
    '
    If TRIM(UCase$(TXTIMPRESION)) = "SI" Then
       TXTLABEL(5) = "Cant. Colores"
    ElseIf TRIM(TXTIMPRESION) <> "" Then
       TXTLABEL(5) = ""
    End If


    If TRIM(UCase$(TXTFUELLE)) = "SI" Then
       TXTLABEL(17) = "Medida Fuelle(x1)Cm."
    ElseIf TRIM(TXTFUELLE) <> "" Then
       TXTLABEL(17) = ""
    End If
    '
    'INFROMACION SEGUN ARCHIVO ENVIADO POR RAUL X:\CONTEPLAST\FLEXOFT\EXCEL\AA PLANILLA PARA VENTAS FINAL.xlsx
    If TRIM(UCase$(TXTMATERIALBOLSA)) = "BOLSA POLIPAPEL" Then
       TXTLARGODECORTE = TXTALTOUTIL
    Else
       TXTLARGODECORTE = XVALO(TXTALTOUTIL) + 4
    End If
    
    
    If TXTIMPRESION = "NO" Then
       RODILLOS = ""
    Else
        'TABLA DATASQL > RODICOLO >
        'CAMPO1= CODIGO
        'CAMPO2 =Rod -LPAP
        'CAMPO3 =EXTRA/F12
        'CAMPO4 =CantColor

        'BUSCA EL VALOR DEL LIMITE INFERIOR ES DECIR SI BUSCA EL 88 Y NO HAY VA A TRARE EL 86 EN CASO Q DE 86 SALTE A 90
        CONDI$ = "LISTA = 'RODICOLO' AND CONVERT(FLOAT, campo2)<= " & XVALO(TXTLARGODECORTE) & " ORDER BY CONVERT(FLOAT, campo2) DESC"
        VARLORLARGOCORTE2_INFERIOR = XVALO(VALOBADA("DATASQL", "TOP 1 CAMPO2", CONDI$, "NOMESS"))
        LC2 = VARLORLARGOCORTE2_INFERIOR
        If LC2 <= 0 Then 'SI NO ENCUENTRA <= BUSCA EL VALOR SIGUIENTE MAYOR
          CONDI$ = "LISTA = 'RODICOLO' AND CONVERT(FLOAT, campo2)>= " & XVALO(TXTLARGODECORTE) & " ORDER BY CONVERT(FLOAT, campo2) "
          VARLORLARGOCORTE2_SUPERIOR = XVALO(VALOBADA("DATASQL", "TOP 1 CAMPO2", CONDI$, "NOMESS"))
          LC2 = VARLORLARGOCORTE2_SUPERIOR
        End If
        TXTLARGODECORTE2 = TRIM$(FORMATNUM$(LC2, "F12.2"))
        '
        'CON LOS DATOS DE ARRIBA BUSCA EL VALOR PARA ASIGNAR A RODILLO
        If XVALO(TXTLARGODECORTE) = LC2 Then
           TXTRODILLO = TRIM$(FORMATNUM$(LC2, "F12.2"))
        Else
           'EN ESTE CASO LA CONDI ES POR IGUAL POR QUE ES UN DATO QUE VIENE EXISTENTE EN LA TABLA EN LAS CONSULTAS ANTERIORES
           CONDI$ = "LISTA = 'RODICOLO' AND CONVERT(FLOAT, campo2)= " & XVALO(LC2)
           TXTRODILLO = LC2 + XVALO(VALOBADA("DATASQL", " CAMPO3", CONDI$, "NOMESS"))
        End If
      End If
      
      CONDI$ = "LISTA = 'RODICOLO' AND CONVERT(FLOAT, campo2)= " & XVALO(TXTRODILLO)
      TXTMAXIMOCOLORES = Int(XVALO(VALOBADA("DATASQL", " CAMPO4", CONDI$, "NOMESS")))
  
    
End Sub


Private Sub TXTLABEL_DblClick(Index As Integer)
    If Label1.Visible = False Then
        Label1.Visible = True
        TXTLARGODECORTE.Visible = True
        Label2.Visible = True
        TXTLARGODECORTE2.Visible = True
        Label3.Visible = True
        TXTRODILLO.Visible = True
     Else
        Label1.Visible = False
        TXTLARGODECORTE.Visible = False
        Label2.Visible = False
        TXTLARGODECORTE2.Visible = False
        Label3.Visible = False
        TXTRODILLO.Visible = False
      End If

End Sub


Private Sub TXTNUMPLAVENTA_Change()
    Call BLANFORMU
    MODIFIC% = 0
    If XVALO(TXTNUMPLAVENTA) > 0 Then
      If CantRegistros&("P_VTA_BOLSAS", "NROPVTA = " & XVALO(TXTNUMPLAVENTA)) < 1 Then
         ESNUEVO% = 1
         Exit Sub
      End If
   Else
      Exit Sub
   End If

   SQLX$ = "SELECT * FROM P_VTA_BOLSAS  WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NROPVTA = " & XVALO(TXTNUMPLAVENTA)
   Set RS = READSET(SQLX$)
   With RS
  
        TEXT1 = XVALO(!NUCLIEN)
        TXTRAZOCLI = SAFETEXT$(!rasocli)
        TXTOBSERVA = SAFETEXT$(!observaciones)
        Text2 = XVALO(!Cantidad)
        Text4 = XVALO(!NUVEND)
        TXTFECHA = CVDAT(CVINT(!FECHEMISI))
        TXTCAPACIDAD = FORMATNUM$(XVALO(!CAPKGS), "F12.2")
        TXTIMPRESION = SAFETEXT$(!IMPRESION)
        TXTCANTCARAS = XVALO(!CANTICARAS)
        TXTMATERIALBOLSA = SAFETEXT$(!MAT_BOLSA)
        TXTALTOUTIL = XVALO(!ALTOUTILCM)
        TXTANCHOTOTAL = XVALO(!ANCHOTOTCM)
        TXTPRODUCTO = SAFETEXT$(!Producto)
        TXTCANTCOLORES = XVALO(!CANTCOLORES)
        TXTFUELLE = SAFETEXT$(!Fuelle)
        TXTMICROPERFORADO = SAFETEXT$(!MICROPERFO)
        TXTPLENO = SAFETEXT$(!Pleno)
        TXTMEDFUELLE = XVALO(!MEDIFUELLE)
        TXTLINER = SAFETEXT$(!LINER)
        TEXT5 = SAFETEXT$(!REFECLIENTE)
        
   End With
   LBLNUEVAPVTA.Visible = False
   MODIFIC% = 0
   RS.Close
   Set RS = Nothing
   Call VALORESPORDEFECTO
   Call CARGARARCHIVOSENCOMBOPVTA(XVALO(TXTNUMPLAVENTA))
End Sub
Sub CARGARARCHIVOSENCOMBOPVTA(NUMPVTA&)
      Combo2.Clear
      SQLX$ = "SELECT * FROM DOCURELPVTA WITH(NOLOCK) WHERE NROPVTA  = " & NUMPVTA&
      Dim DOCTMP As ADODB.Recordset
      Set DOCTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With DOCTMP
        Do While Not .EOF
           DESCRI$ = SAFETEXT$(!Descripcion)
           RUTAX$ = SAFETEXT$(!Ruta$)
           Call REPLA(TX$, DESCRI$, 1, 24)
           Call REPLA(TX$, RUTAX$, 32, 512)
           Combo2.AddItem TX$
           .MoveNext
        Loop
      End With
End Sub

Private Sub TXTRAZOCLI_Change()
MODIFIC% = 1
End Sub


Private Sub TXTRAZOCLI_DblClick()
    If TXTRAZOCLI.BackColor = &HE0E0E0 Then
       TXTRAZOCLI.BackColor = &H80FFFF
       TXTRAZOCLI.Locked = False
       'Call PINTATEXT(TXTRAZOCLI)
       Command1.Enabled = False
       If XVALO(TEXT1.TEXT) <> 99000 Then TEXT1.TEXT = ""
    Else
       TXTRAZOCLI.BackColor = &HE0E0E0
       TXTRAZOCLI.Locked = True
       Command1.Enabled = True
    End If
    MODIFIC% = 1
End Sub


Private Sub TXTRAZOCLI_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If
End Sub

Sub PrintPlanillaVenta()
    Call SETPRINPORT("RESTORE")
    PUERTOIMPRE$ = CONTROL$(IDENTER$, "PRINTSER")
    If PUERTOIMPRE$ = "" Then PUERTOIMPRE$ = CONTROL$("*", "PRINTSER")
    If PUERTOIMPRE$ <> "" Then
       Call SETPRINPORT(PUERTOIMPRE$)
    End If
    '
    RUTAEXCEL$ = CONTROL("COTICONT", "EXCEPVTA")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Ficha Técnica")
       Exit Sub
    End If
    '

    With VENTABNEW.MSF
       .Rows = 22: .Cols = 3
       '.TextMatrix(1, 1) = TXTRAZOCLI 'nombre cliente
       .TextMatrix(1, 2) = TXTRAZOCLI 'nombre cliente
       '.TextMatrix(2, 1) = Text2
       .TextMatrix(2, 2) = Text2
       .TextMatrix(3, 2) = Text3
       .TextMatrix(3, 1) = Text3
       .TextMatrix(4, 1) = TXTLABEL(1)
       .TextMatrix(4, 2) = TXTMATERIALBOLSA
       .TextMatrix(5, 1) = TXTLABEL(2)
       .TextMatrix(5, 2) = TXTANCHOTOTAL
       .TextMatrix(6, 1) = TXTLABEL(40)
       .TextMatrix(6, 2) = TXTCAPACIDAD
       
       .TextMatrix(7, 1) = TXTLABEL(4)
       .TextMatrix(7, 2) = TXTIMPRESION
       .TextMatrix(8, 1) = TXTLABEL(5)
       .TextMatrix(8, 2) = TXTCANTCOLORES
       .TextMatrix(9, 1) = TXTLABEL(6)  'Fuelle
       .TextMatrix(9, 2) = TXTFUELLE  'Fuelle
       .TextMatrix(10, 1) = TXTLABEL(9) 'Micro Perforado
       .TextMatrix(10, 2) = TXTMICROPERFORADO  'Micro Perforado
       .TextMatrix(11, 1) = TXTLABEL(10) 'Producto
       .TextMatrix(11, 2) = TXTPRODUCTO 'Producto
       .TextMatrix(12, 1) = TXTLABEL(13) 'Alto Util Cm.
       .TextMatrix(12, 2) = TXTALTOUTIL  'Alto Util Cm.
       .TextMatrix(13, 1) = TXTLABEL(14)  'Pleno
       .TextMatrix(13, 2) = TXTPLENO  'Pleno
       .TextMatrix(14, 1) = TXTLABEL(16) 'Cant.Caras
       .TextMatrix(14, 2) = TXTCANTCARAS 'Cant.Caras
       .TextMatrix(15, 1) = TXTLABEL(17) 'Medida Fuelle(x1) cm.
       .TextMatrix(15, 2) = TXTMEDFUELLE 'Medida Fuelle(x1) cm.
       .TextMatrix(16, 1) = TXTLABEL(7) 'Liner
       .TextMatrix(16, 2) = TXTLINER 'Liner
       '.TextMatrix(17, 1) = TXTOBSERVA  'Observaciones
       .TextMatrix(17, 2) = TXTOBSERVA  'Observaciones
       .TextMatrix(18, 1) = TXTNUMPLAVENTA   'nroplanilla
       .TextMatrix(18, 2) = TXTNUMPLAVENTA   'nroplanilla
       .TextMatrix(19, 1) = TXTFECHA   'fecha
       .TextMatrix(19, 2) = TXTFECHA   'fecha
       
       .TextMatrix(20, 2) = "SI"
       If Combo2.ListCount < 1 Then
          .TextMatrix(20, 2) = "NO"
       End If
    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
        
    Call Param_Print_Excel(RUTAEXCEL$, "SIMULADOR", VENTABNEW.MSF, PUERTOIMPRE$, "1;2", 1, 10, "", 1, 0)                      'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW

    Call SETPRINPORT("RESTORE")





End Sub

