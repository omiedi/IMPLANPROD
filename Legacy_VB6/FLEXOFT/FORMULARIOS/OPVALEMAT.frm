VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPVALEMAT 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración - Transferencia de Materiales y Servicios"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture16 
      Height          =   6585
      Left            =   10680
      ScaleHeight     =   6525
      ScaleWidth      =   1035
      TabIndex        =   87
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   6700
         Left            =   0
         Picture         =   "OPVALEMAT.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   9360
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   82
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   2100
      Style           =   2  'Dropdown List
      TabIndex        =   73
      Top             =   6720
      Width           =   3100
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Orden de Fabricación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   5400
      TabIndex        =   59
      Top             =   5020
      Width           =   5160
      Begin VB.CommandButton Command24 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4545
         Picture         =   "OPVALEMAT.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   1450
         Width           =   495
      End
      Begin VB.CommandButton Command23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4545
         Picture         =   "OPVALEMAT.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1050
         Width           =   495
      End
      Begin VB.TextBox Text26 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   77
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton Command22 
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
         Left            =   2040
         TabIndex        =   76
         Top             =   1440
         Width           =   175
      End
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   75
         Text            =   " "
         Top             =   1440
         Width           =   1140
      End
      Begin VB.TextBox Text24 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   72
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1080
         Width           =   2175
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
         Height          =   360
         Left            =   2040
         TabIndex        =   71
         Top             =   1080
         Width           =   175
      End
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   70
         Text            =   " "
         Top             =   1080
         Width           =   1140
      End
      Begin VB.CommandButton Command20 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4545
         Picture         =   "OPVALEMAT.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   650
         Width           =   495
      End
      Begin VB.CommandButton Command19 
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
         Left            =   2040
         TabIndex        =   66
         Top             =   720
         Width           =   175
      End
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   67
         Text            =   " "
         Top             =   720
         Width           =   1140
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   65
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command18 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4545
         Picture         =   "OPVALEMAT.frx":2CB6
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   " "
         Top             =   360
         Width           =   2175
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
         Height          =   360
         Left            =   2040
         TabIndex        =   61
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   60
         Text            =   " "
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O-Comp:"
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
         Left            =   -100
         TabIndex        =   81
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-Consig:"
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
         Left            =   -90
         TabIndex        =   80
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cierre OF:"
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
         Left            =   -90
         TabIndex        =   68
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Etiqueta:"
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
         Left            =   -90
         TabIndex        =   64
         Top             =   480
         Width           =   1005
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Remito Consignación - Armador Externo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1400
      Left            =   5400
      TabIndex        =   52
      Top             =   105
      Width           =   5160
      Begin VB.CheckBox Check18 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Imprimir Remito en Consignación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   88
         Top             =   1120
         Width           =   4110
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   4185
         Picture         =   "OPVALEMAT.frx":2FC0
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   315
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3750
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
         Height          =   360
         Left            =   3885
         TabIndex        =   54
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   53
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   57
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.TextBox Text19 
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
      Left            =   1440
      TabIndex        =   50
      Text            =   " "
      Top             =   6360
      Width           =   3780
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Orden de Compra / Pedido de Suministro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1430
      Left            =   5400
      TabIndex        =   42
      Top             =   3360
      Width           =   5160
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   " Suprime  la  Impresión  al  Generar  el  Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   49
         Top             =   1155
         Width           =   4740
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   46
         Text            =   " "
         Top             =   330
         Width           =   1380
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
         Height          =   360
         Left            =   3885
         TabIndex        =   45
         Top             =   330
         Width           =   175
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   315
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3750
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   4185
         Picture         =   "OPVALEMAT.frx":32CA
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   47
         Top             =   435
         Width           =   2325
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Orden de Servicios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   5400
      TabIndex        =   36
      Top             =   1800
      Width           =   5160
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   " Suprime  la  Impresión  al  Generar  el  Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   48
         Top             =   1050
         Width           =   4740
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
         Height          =   360
         Left            =   2520
         TabIndex        =   40
         Text            =   " "
         Top             =   225
         Width           =   1380
      End
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
         Left            =   3885
         TabIndex        =   39
         Top             =   225
         Width           =   175
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   315
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   " "
         Top             =   630
         Width           =   3750
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   4185
         Picture         =   "OPVALEMAT.frx":35D4
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   105
         TabIndex        =   41
         Top             =   330
         Width           =   2325
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Depósito de Armadores Externos"
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
      Left            =   120
      TabIndex        =   31
      Top             =   4200
      Width           =   5145
      Begin VB.TextBox Text11 
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
         Left            =   1000
         TabIndex        =   34
         Text            =   " "
         Top             =   350
         Width           =   450
      End
      Begin VB.CommandButton Command8 
         Caption         =   "..."
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
         Left            =   1440
         Picture         =   "OPVALEMAT.frx":38DE
         TabIndex        =   33
         Top             =   350
         Width           =   175
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
         Left            =   1680
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   " "
         Top             =   350
         Width           =   3240
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito: "
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
         Left            =   50
         TabIndex        =   35
         Top             =   400
         Width           =   960
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Controles Varios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   120
      TabIndex        =   29
      Top             =   5160
      Width           =   5145
      Begin VB.TextBox Text29 
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
         Left            =   1800
         TabIndex        =   96
         TabStop         =   0   'False
         Text            =   " "
         Top             =   600
         Width           =   3240
      End
      Begin VB.CommandButton Command27 
         Caption         =   "..."
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
         Left            =   1515
         Picture         =   "OPVALEMAT.frx":3BE8
         TabIndex        =   95
         Top             =   600
         Width           =   175
      End
      Begin VB.TextBox Text28 
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
         Left            =   1065
         TabIndex        =   94
         Text            =   " "
         Top             =   600
         Width           =   450
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Bloqueo de Negativos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   93
         Top             =   240
         Width           =   4305
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito Scrap "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         TabIndex        =   97
         Top             =   600
         Width           =   840
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Depósito Destino (Taller)"
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
      TabIndex        =   24
      Top             =   2400
      Width           =   5150
      Begin VB.TextBox Text9 
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
         Left            =   1660
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   " "
         Top             =   350
         Width           =   3255
      End
      Begin VB.TextBox Text8 
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
         Left            =   1000
         TabIndex        =   26
         Text            =   " "
         Top             =   350
         Width           =   450
      End
      Begin VB.CommandButton Command7 
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
         Left            =   1440
         TabIndex        =   25
         Top             =   350
         Width           =   175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito:"
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
         Left            =   50
         TabIndex        =   28
         Top             =   400
         Width           =   885
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Depósito Cierre (Entra P.Term)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   120
      TabIndex        =   19
      Top             =   3200
      Width           =   5145
      Begin VB.TextBox Text27 
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
         Left            =   1020
         TabIndex        =   91
         Text            =   " "
         Top             =   600
         Width           =   450
      End
      Begin VB.CommandButton Command26 
         Caption         =   "..."
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
         Left            =   1470
         Picture         =   "OPVALEMAT.frx":3EF2
         TabIndex        =   90
         Top             =   600
         Width           =   175
      End
      Begin VB.TextBox Text77 
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
         Left            =   1680
         TabIndex        =   89
         TabStop         =   0   'False
         Text            =   " "
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox Text7 
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
         Left            =   1665
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   " "
         Top             =   250
         Width           =   3255
      End
      Begin VB.CommandButton command10 
         Caption         =   "..."
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
         Left            =   1470
         Picture         =   "OPVALEMAT.frx":41FC
         TabIndex        =   20
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox Text5 
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
         Left            =   1005
         TabIndex        =   21
         Text            =   " "
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep.Import."
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
         Left            =   50
         TabIndex        =   92
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito: "
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
         Left            =   50
         TabIndex        =   22
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture14 
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
      Height          =   6600
      Left            =   10680
      ScaleHeight     =   6540
      ScaleWidth      =   1005
      TabIndex        =   13
      Top             =   120
      Width           =   1065
      Begin VB.CommandButton Command25 
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
         Left            =   120
         Picture         =   "OPVALEMAT.frx":4506
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Definición de Usuarios por Sector"
         Top             =   3400
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command16 
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
         Left            =   120
         Picture         =   "OPVALEMAT.frx":4650
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Tablas T.Canje - Configuración "
         Top             =   2625
         Width           =   650
      End
      Begin VB.CommandButton Command62 
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
         Left            =   120
         Picture         =   "OPVALEMAT.frx":479A
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Tabla de Armadores"
         Top             =   1850
         Width           =   650
      End
      Begin VB.CommandButton Command6 
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
         Left            =   120
         Picture         =   "OPVALEMAT.frx":48E4
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1080
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
         Height          =   645
         Left            =   120
         Picture         =   "OPVALEMAT.frx":4BEE
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Guardar Configuracion de Vale de Materiales"
         Top             =   4515
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         Picture         =   "OPVALEMAT.frx":4EF8
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   300
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   14
         Top             =   5355
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "OPVALEMAT.frx":5042
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5670
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   810
      Left            =   4200
      TabIndex        =   11
      Top             =   1560
      Width           =   1065
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
         Left            =   315
         TabIndex        =   12
         Text            =   "TR"
         Top             =   270
         Width           =   450
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Vale de Materiales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   120
      TabIndex        =   5
      Top             =   105
      Width           =   5160
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFC0&
         Caption         =   " Suprime  la  Impresión  al  Generar  el  Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   86
         Top             =   1020
         Width           =   4740
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
         Height          =   360
         Left            =   2415
         TabIndex        =   10
         Text            =   " "
         Top             =   195
         Width           =   1380
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
         TabIndex        =   9
         Top             =   200
         Width           =   175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   615
         Width           =   3750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   4095
         Picture         =   "OPVALEMAT.frx":6F64
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   195
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   1320
         TabIndex        =   6
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Depósito Central (Pañol)"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   4005
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Stock Neto en O.F. - Considera solo pañol "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   84
         Top             =   575
         Width           =   3705
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
         Height          =   300
         Left            =   1440
         TabIndex        =   3
         Top             =   225
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Left            =   1000
         TabIndex        =   2
         Text            =   " "
         Top             =   225
         Width           =   450
      End
      Begin VB.TextBox Text4 
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
         Left            =   1660
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   " "
         Top             =   225
         Width           =   2205
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito:"
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
         Left            =   50
         TabIndex        =   4
         Top             =   285
         Width           =   885
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "OPVALEMAT.frx":726E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPVALEMAT.frx":74A2
      TabIndex        =   83
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pto.O/Etiqueta OF:"
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
      TabIndex        =   74
      Top             =   6840
      Width           =   1710
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email Compras"
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
      Left            =   120
      TabIndex        =   51
      Top             =   6480
      Width           =   1245
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu tabla 
      Caption         =   "Tablas "
      Begin VB.Menu armadores 
         Caption         =   "Tabla de Armadores"
         Enabled         =   0   'False
      End
      Begin VB.Menu tablascanje 
         Caption         =   "Tablas T.Canje - Configuración "
         Enabled         =   0   'False
      End
      Begin VB.Menu usuarioxsector 
         Caption         =   "Definición de Usuarios por Sector"
      End
   End
   Begin VB.Menu mnuMigracion 
      Caption         =   "MIGRACION"
      Begin VB.Menu mnuMigra 
         Caption         =   "Creacion de Tablas y Migracion en Base Nueva"
      End
      Begin VB.Menu mnuPruebatablabasesdistintas 
         Caption         =   "Prueba de transaccion entre tablas de bases distintas"
      End
   End
End
Attribute VB_Name = "OPVALEMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub armadores_Click()
   Call Command62_Click
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Text5 = TRIM$(VALACT1$("TADEPOSI"))
    End If
End Sub

Private Sub Command11_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        TEXT13.TEXT = VDEVU$
        Text12.TEXT = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Command12_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text15.TEXT), "Formularios de Impresion")

End Sub

Private Sub Command13_Click()
 Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text15.TEXT = VDEVU$
        Text14.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command14_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text16.TEXT = VDEVU$
        Text17.TEXT = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Command15_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text16.TEXT), "Formularios de Impresion")

End Sub

Private Sub Command16_Click()
    
    Command16.Enabled = False
    '
    'valido si existe el archivo para evitar el error en otras empresas
    If EXISTE%(LUDAT$ + "TATICANJ.RFS") < 1 Then
      Call COMUNI("Tabla de Tipos de Canje \No Disponible en la Presente Versión")
      GoTo 99
    End If
    '
    'CREO LOS CAMPOS
    Call CREACAMPO("", "CANJPREDEF", "CODICONJ", 3, 0) 'CODIGO INTERNO DEL PRODUCTO
    Call CREACAMPO("", "CANJPREDEF", "TIPOCANJE", 3, 0)
    Call CREACAMPO("", "CANJPREDEF", "CODIELEM", 3, 0) ' CODIGO INTERNO DEL COMPONENTE
    Call CREACAMPO("", "CANJPREDEF", "AGREGAR", 7, 0)
    Call CREACAMPO("", "CANJPREDEF", "QUITAR", 7, 0)
    Call CREACAMPO("", "CANJPREDEF", "MARBORRA", 3, 0)
    '
10  OPX% = COMALTER%("Opciones de Trabajo\\Tabla de Tipo de Canje\Configuración de Canje")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    '
    If OPX% = 1 Then
      Call TABTIPOCANJE
      GoTo 10
    ElseIf OPX% = 2 Then
      Call CONFIGUCANJE
      GoTo 10
    End If
    
99  Command16.Enabled = True
   
End Sub
Sub TABTIPOCANJE()
    '
    ATRI$ = "Tabla de Tipos de Canje"
    Call TRALOCAL("TATICANJ", "")
10  VTB% = VENTABU%("TATICANJ/" + ATRI$)
    '
    If VTB% < 0 Then Exit Sub
    '
    'VALIDO
    For U& = 1 To ULTREG&("!TATICANJ")
         X$ = REGLEIDO$("!TATICANJ", U&)
         CODI% = CVI(Mid$(X$, 1, 2))
         DENOMI$ = TRIM$(Mid$(X$, 3, 62))
         If CODI% > 0 And DENOMI$ = "" Then
            Call COMUNI("Falta Denominación Para Código: " & CODI%)
            GoTo 10
         End If
    Next U&
    '
    Call TRACENTRAL("TATICANJ", "")
    
End Sub
Sub CONFIGUCANJE()
  
10    Call SELECHO("TATICANJ")
      TIPOCAN% = XVALO(VALACT1$("TATICANJ"))
      If TIPOCAN% < 1 Then Exit Sub
      '
      Call SELECHO("MASTER")
      CI% = CODINT%(VALACT1$("MASTER"))
      If CI% < 1 Then GoTo 10
      '
      CONFCANJ.Label11 = CODEXT$(CI%)
      CONFCANJ.Label1 = TRIM$(str$(TIPOCAN%))
      '
      CONFCANJ.Show 1
      '
End Sub

Private Sub Command17_Click()
 Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text18.TEXT = VDEVU$
        Text20.TEXT = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Command18_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text18.TEXT), "Formularios de Impresion")

End Sub

Private Sub Command19_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text22.TEXT = VDEVU$
        Text21.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORVALMA", Text1.TEXT)
    Call GRACONTROL("", "DEPOCEN", Text3.TEXT)
    Call GRACONTROL("", "CODVALMA", Text6.TEXT)
    Call GRACONTROL("", "DEPOCIE", Text5.TEXT)
    Call GRACONTROL("", "DEPOIMPO", Text27.TEXT)
    Call GRACONTROL("", "DEPODEST", Text8.TEXT)
    Call GRACONTROL("", "DEPOARMA", Text11.TEXT)
    Call GRACONTROL("", "ETIORFA", Text18.TEXT)
    Call GRACONTROL("", "VALCIEOF", Text22.TEXT)
    Call GRACONTROL("", "RCONSIOF", Text23.TEXT)
    Call GRACONTROL("", "ORCOMOF", Text25.TEXT)
    '
    Call GRACONTROL("VANEDECE", "DEPOCEN", INFO$)  'GRABA OPCION DE IMPRSION CON IVA
    '
    SUPRINT$ = "NO"
    If Check1.Value = 1 Then SUPRINT$ = "SI"
    Call GRACONTROL("ORSERVI1", "SUPRINT", SUPRINT$)
    '
    MONUME$ = "NO"
    If Check3.Value = 1 Then
      If XVALO(Text3) < 1 Then 'valido que se ingrese un deposito de pañol en el caso que este tildada la opcion
         Call COMUNI("Falta Ingresar Depósito Pañol")
         Exit Sub
      End If
      MONUME$ = "SI"
      Call GRACONTROL("", "STOPAÑOL", MONUME$)
    End If
    '
    SUPRINT$ = "NO"
    If Check2.Value = 1 Then SUPRINT$ = "SI"
    Call GRACONTROL("OCOMPRA1", "SUPRINT", SUPRINT$)
    '
    Call GRACONTROL("", "FORSERVI", TEXT13.TEXT)
    Call GRACONTROL("", "FORMOCOM", Text15.TEXT)
    Call GRACONTROL("", "FOREMICO", Text16.TEXT) 'FRM REMITO EN CONSIGNACION ARMADOR EXTERNO
    'E-mail para Gestion de control final
    Call GRACONTROL("OPVALEMA", "CONTMAIL", Text19)
    'Call GRACONTROL("ORFAETIQ", "PORTETIQ", Combo2.TEXT) 'GRABA LA IMPRESORA SELECCIONADA
    Call GRACONTROL(IDENTER$, "PORTETIQ", Combo2.TEXT) 'GRABA LA IMPRESORA SELECCIONADA

    SUPRINT$ = "NO": If Check10.Value = 1 Then SUPRINT$ = "SI"
    Call GRACONTROL("VANEDECE", "DEPOCEN", SUPRINT$)
    
    SUPRINT$ = "NO"
    If Check4.Value = 1 Then SUPRINT$ = "SI"
    Call GRACONTROL(IDENTER$, "SUIMPVAL", SUPRINT$)
    
    Call GRACONTROL("", "DEPSCRAP", Text28.TEXT)
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command20_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text22.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command21_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text23.TEXT = VDEVU$
        Text24.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command22_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text25.TEXT = VDEVU$
        Text26.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command23_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text23.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command24_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text25.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command25_Click()
    '
    Command24.Enabled = False
    '
    If DERACCE%("USUASEC", "USUARIOS POR SECTOR PRODUCTIVO") < 2 Then GoTo 99
    
    Call SETULTREG("!USUASEC", 0)
    J& = 0
    For U& = 1 To ULTREG&("USUASEC")
        X$ = REGLEIDO$("USUASEC", U&)
        COSECC = Asc(Left$(X$, 1))
        If COSECC >= 1 Then
            Call REGAPP("!USUASEC", X$)
        End If
    Next U&
    Call CIERRARCH("USUASEC")
    Call CIERRARCH("!USUASEC")
    '
    VTB% = VENTABU%("USUASEC")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!USUASEC")
        X$ = REGLEIDO$("!USUASEC", U&)
        COSECC = Asc(Left$(X$, 1))
        If COSECC >= 1 Then
            J& = J& + 1
            Call GRAREG("USUASEC", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("USUASEC", J&)
    Call CIERRARCH("USUASEC")
    Call CIERRARCH("!USUASEC")
    '
99  Command25.Enabled = True
    '
End Sub

Private Sub Command26_Click()
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Text27 = TRIM$(VALACT1$("TADEPOSI"))
    End If
End Sub

Private Sub Command27_Click()
    '
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Text28 = TRIM$(VALACT1$("TADEPOSI"))
      Text29 = TRIM$(VALACT2$("TADEPOSI"))
    End If
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
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("TADEPOSI")
    End If
End Sub

Private Sub Command5_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command62_Click()
    Command62.Enabled = False
  '
    Call TRALOCAL("ARMADORS", "")
    ATRI$ = "Tabla de Armadores "
    
10  VTB% = VENTABU%("ARMADORS1/" + ATRI$)
    If VTB% < 0 Then GoTo 99
    JJ& = 0
    '
    'VALIDACIONES
    '
    For U& = 1 To ULTREG&("!ARMADORS")
         X$ = REGLEIDO$("!ARMADORS", U&)
         CODARMAD% = CVI(Mid$(X$, 1, 2))
         ARMAD$ = Mid$(X$, 3, 64)
         If CODARMAD% > 0 Then
            TIPO_ARMADOR$ = Mid$(X$, 67, 1)
            COD_PROV% = CVI(Mid$(X$, 68, 2))
            ' SI TIENE MARCA DE ARMADOR EXTERNO DEBE ASIGNAR UN PROVEEDOR
            If TRIM$(UCase$(TIPO_ARMADOR$)) = "E" Then
               If ECOARCH$("PROVED1", MKI$(COD_PROV%)) = "" Then
                  Call COMUNI("Armador: " & CODARMAD% & " (" & TRIM$(ARMAD$) & ") Falta Asignar Proveedor")
                  GoTo 10
               End If
            Else
            ' SI NO TIENE MARCA DE ARMADOR EXTERNO NO PUEDE ASIGNAR PROVEEDOR
               If ECOARCH$("PROVED1", MKI$(COD_PROV%)) <> "" Then
                  Call COMUNI("Armador: " & CODARMAD% & " (" & TRIM$(ARMAD$) & ") No Válido Asignar Proveedor")
                  GoTo 10
               End If
            End If
            
         End If

         JJ& = JJ& + 1
         Call GRAREG("ARMADORS", X$, JJ&)
    Next U&
    '
    'GRABO
    '
    JJ& = 0
    For U& = 1 To ULTREG&("!ARMADORS")
         X$ = REGLEIDO$("!ARMADORS", U&)
         JJ& = JJ& + 1
         Call GRAREG("ARMADORS", X$, JJ&)
     Next U&
    Call SETULTREG("ARMADORS", JJ&)
    Call CIERRARCH("ARMADORS")
    Call TRACENTRAL("ARMADORS", "ARMADORS")

    
99  Command62.Enabled = True

End Sub

Private Sub Command7_Click()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("TADEPOSI")
    End If

End Sub

Private Sub Command8_Click()
    '
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Text11 = TRIM$(VALACT1$("TADEPOSI"))
    End If
    '
End Sub

Private Sub Command9_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(TEXT13.TEXT), "Formularios de Impresion")

End Sub

Private Sub Form_Load()
    '
If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Me.BackColor = &H8000000F
      Me.Frame1.BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Frame5.BackColor = &H8000000F
      Me.Frame6.BackColor = &H8000000F
      Me.Frame7.BackColor = &H8000000F
      Me.Frame8.BackColor = &H8000000F
      Me.Frame9.BackColor = &H8000000F
      Me.Frame10.BackColor = &H8000000F
      Me.Frame11.BackColor = &H8000000F
      Me.Frame12.BackColor = &H8000000F
      
      
      '
      Me.Check1.BackColor = &H8000000F
      Me.Check2.BackColor = &H8000000F
      Me.Check3.BackColor = &H8000000F
      Me.Check4.BackColor = &H8000000F
      Me.Check10.BackColor = &H8000000F
      '
      MARCOBARRA.Top = 6400
      
      Set MARCOBARRA.Container = Me.Picture16
      
    End If
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)

    Screen.MousePointer = 11
    J& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
        X$ = REGLEIDO$("SUCURSAL", U&)
        CODEPO% = Asc(Mid$(X$, 64, 1))
        If CODEPO% > 0 Then
           If CODEPO% <> 99 Then
              If J& < 62 Then
                 J& = J& + 1
                 Y$ = REGBLAN$("TADEPOSI")
                 Call REPLA(Y$, Chr$(CODEPO%), 1, 1)
                 Call REPLA(Y$, Left$(X$, 63), 2, 63)
                 Call GRAREG("TADEPOSI", Y$, J&)
              End If
           End If
        End If
    Next U&
    Call SETULTREG("TADEPOSI", J&)
    Call CIERRARCH("SUCURSAL")
    Call CIERRARCH("TADEPOSI")
    Screen.MousePointer = 1
    '
    Text1.TEXT = Control$("", "FORVALMA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = Control$("", "DEPOCEN")
    Text4.TEXT = ECOARCH$("TADEPOSI", Chr$(Val(Text3.TEXT)))
    Text5.TEXT = Control("", "DEPOCIE")
    Text27.TEXT = Control("", "DEPOIMPO")
    Text8.TEXT = Control("", "DEPODEST")
    Text11.TEXT = Control("", "DEPOARMA") ' DEPOSITO DE ARMADOR
    Check10.Value = 0
    If Control("VANEDECE", "DEPOCEN") = "SI" Then Check10.Value = 1 ' CONTROL DE NEGATIVOS SOBRE DEPOSITO CENTRAL
    '
    If Control$("ORSERVI1", "SUPRINT") = "SI" Then Check1.Value = 1 ' CONTROLES DE IMPRESION
    If Control$("OCOMPRA1", "SUPRINT") = "SI" Then Check2.Value = 1
    '
    TEXT13.TEXT = Control$("", "FORSERVI")
    Text12.TEXT = ECOARCH$("INDIFRM", TEXT13.TEXT)
    Text15.TEXT = Control$("", "FORMOCOM")
    Text14.TEXT = ECOARCH$("INDIFRM", Text15.TEXT)
    Text16.TEXT = Control$("", "FOREMICO")
    Text17.TEXT = ECOARCH$("INDIFRM", Text16.TEXT)
    Text18 = Control$("", "ETIORFA")
    Text20.TEXT = ECOARCH$("INDIFRM", Text18.TEXT)
    Text22 = Control$("", "VALCIEOF")
    Text21.TEXT = ECOARCH$("INDIFRM", Text22.TEXT)
    Text23 = Control$("", "RCONSIOF")
    Text24.TEXT = ECOARCH$("INDIFRM", Text23.TEXT)
    Text25 = Control$("", "ORCOMOF")
    Text26.TEXT = ECOARCH$("INDIFRM", Text25.TEXT)
    Text28.TEXT = Control("", "DEPSCRAP")
    Text19 = Control("OPVALEMA", "CONTMAIL")
    CMO$ = TRIM$(Control$("", "CODVALMA"))
    If CMO$ = "" Then CMO$ = "TR"
    Text6.TEXT = CMO$
    '
    Check3.Value = 0
    If Control("", "STOPAÑOL") = "SI" Then Check3.Value = 1
    '
    PFAC$ = TRIM$(Control(IDENTER$, "PORTETIQ"))
    Combo2.Clear ' CARGA EL COMBO DE LAS IMPRESORAS
    Dim PRX As Printer
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo2.AddItem PRX.DeviceName
      If PRX.DeviceName = PFAC$ Then
        Combo2.TEXT = PFAC$
      End If
    Next
    Check4.Value = 0
    IMPR$ = TRIM$(Control$(IDENTER$, "SUIMPVAL"))
    If IMPR$ = "SI" Then Check4.Value = 1

    If InStr(1, EMPREAC$, "ARCON") > 0 Then Command25.Visible = True
    If InStr(1, EMPREAC$, "ENVAPLAST") > 0 Then Command25.Visible = True
End Sub

Private Sub grabar_Click()
   Call Command2_Click
End Sub

Private Sub MNUMIGRA_Click()
    If CLACONTRA% < 1 Then Exit Sub
    'MIGRACION
    T11$ = "ESTE PROCESO GENERA LA MIGRACION DE DATOS DESDE EL SISTEMA DONDE SE ESTA PARADO A OTRA BASE NUEVA (dentro de la misma instancia) "
    T11$ = T11$ + "\Ingresandole Manualmente el nombre de la Nueva Base (donde se migraran los datos (debe estar creada la nueva base))"
    T11$ = T11$ + "\Primero recorre todas las tablas y las crea en la base nueva"
    T11$ = T11$ + "\Luego Comienza a Migar la informacion - Pregunta tabla por tabla si se desea pasar los datos"
    T11$ = T11$ + "\Se puede saltear tablas o bien cancelar la migracion)"
    Call COMUNI(T11$)
    
    BASE_DESTINO$ = InputBox$("Ingrese Base de Destino (" & BASEACTIVA$ & ")", "Migración de Base de Datos")
    If TRIM$(BASE_DESTINO$) = "" Then GoTo 99
    '
    Call ABRECONEXION_DESTINO(BASE_DESTINO$)
    '
    Call ABRECONEXION
    '
    Me.Caption = "Creando Tablas/Campos en B.Destino"
    'CREO LAS TABLAS Y CAMPOS
    SQLX$ = "SELECT CAST(table_name as varchar) AS TABLANAME FROM INFORMATION_SCHEMA.TABLES ORDER BY TABLANAME"
    
    Dim RSTORIGEN As New ADODB.Recordset
    Set RSTORIGEN = FLEXCONN.Execute(SQLX$)
    With RSTORIGEN
       
       Do While Not .EOF
         NOMBRE_TABLA$ = SAFETEXT$(!TABLANAME)
         If UCase$(Left$(NOMBRE_TABLA$, 3)) = "SYS" Or UCase$(Left$(NOMBRE_TABLA$, 3)) = "DTP" Then GoTo 35
         Call CFG_CREARTABLA(NOMBRE_TABLA$, BASE_DESTINO$)
35       .MoveNext
       Loop
    End With
    '
    Call GENERAR_ESTRUCTURAS_SQL
    
    RSTORIGEN.Close
    Set RSTORIGEN = Nothing
    '
    
    SQLX$ = "SELECT COUNT (*) AS CANTI FROM INFORMATION_SCHEMA.TABLES"
    Dim CANTMP As New ADODB.Recordset
    Set CANTMP = FLEXCONN_DESTINO.Execute(SQLX$)
    Reg& = CANTMP!CANTI
    OPX% = COMALTER%("Se han Creado " & Reg& & " Tablas\\Continuar Migración \Cancelar Migración")
    If OPX% <> 1 Then GoTo 99

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Me.Caption = "Transfiriendo Datos"
    SQLX$ = "SELECT CAST(table_name as varchar) AS TABLANAME FROM INFORMATION_SCHEMA.TABLES ORDER BY TABLE_NAME"
    Dim RSTORIGEN2 As New ADODB.Recordset
    Set RSTORIGEN2 = FLEXCONN.Execute(SQLX$)
    With RSTORIGEN2
       Do While Not .EOF
         'RECORRO LAS TABLAS
         NOMBRE_TABLA$ = SAFETEXT$(!TABLANAME)
         If UCase$(Left$(NOMBRE_TABLA$, 3)) = "SYS" Or UCase$(Left$(NOMBRE_TABLA$, 3)) = "DTP" Then GoTo 75
         OPX% = COMALTER%("Migración de Base de Datos\\Continuar Migración de Tabla: " & NOMBRE_TABLA$ & "\Cancelar Migración\Continuar Con Tabla Siguiente")
         If OPX% = 2 Or OPX% < 1 Then
            Call COMUNI("Migración Cancelada")
            GoTo 99
         ElseIf OPX% = 3 Then
            GoTo 75
         Else
             'VERIFICA SI LA TABLA TIENE INFORMACION,
45           SQLX$ = "SELECT COUNT (*) AS CANTI FROM " & NOMBRE_TABLA$
             Dim CANTREGIS As New ADODB.Recordset
             On Error Resume Next
             Set CANTREGIS = FLEXCONN_DESTINO.Execute(SQLX$)
             Reg& = CANTREGIS!CANTI
             CANTREGIS.Close
             Set CANTREGIS = Nothing
             On Error GoTo 0
             '
             'SI TIENE INFORMACION PREGUNTA SI SOBREESCRIBE O CONTINUA
             If Reg& > 0 Then
                OPX% = COMALTER%("La Tabla " & NOMBRE_TABLA$ & " Contiene Información (" & Reg& & " Registros) \\Sobre-Escribir \Continuar Con Tabla Siguiente")
                If OPX% < 1 Or OPX% = 2 Then GoTo 75
                If OPX% = 1 Then
                   SQLX$ = "DELETE FROM " & NOMBRE_TABLA$
                   FLEXCONN_DESTINO.Execute (SQLX$)
                   GoTo 45
                End If
             End If
         End If
         
         'BUSCO EN LA TABLA ORIGINAL
         SQLX$ = " SELECT * FROM " & NOMBRE_TABLA$
         Dim Rst As ADODB.Recordset
         Set Rst = FLEXCONN.Execute(SQLX$)
         '
         sqlx2$ = " SELECT * FROM " & NOMBRE_TABLA$
         Dim RSTDESTI As New ADODB.Recordset
25       RSTDESTI.Open sqlx2$, FLEXCONN_DESTINO, adOpenKeyset, adLockPessimistic, adCmdText
         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 25
         End If
         U& = 0
         FIN& = Rst.RecordCount
         With Rst
           Do While Not .EOF
                Screen.MousePointer = 11
                U& = U& + 1
                Me.Caption = "Transfiriendo Datos de Tabla " & NOMBRE_TABLA$ & "(" & U& & ")"
                Me.Refresh
                Call TRACE(20, U&, FIN&)
                '
                For KKI% = 1 To .Fields.Count
                    NOCAM$ = UCase$(.Fields(KKI% - 1).Name)
                    VALAC = .Fields(KKI% - 1).Value
                    '
                    With RSTDESTI
                        If KKI% = 1 Then
                           .AddNew
                        End If
                        For KKJ% = 1 To .Fields.Count
                           If UCase$(.Fields(KKJ% - 1).Name) = NOCAM$ Then
                              .Fields(KKJ% - 1).Value = VALAC
                           End If
                        Next KKJ%
                    End With
                Next KKI%
                RSTDESTI.Update
                '
29              .MoveNext
           Loop
           RSTDESTI.Close: Set RSTDESTI = Nothing
         End With
75       .MoveNext
       Loop
    End With
    Rst.Close: Set Rst = Nothing
'    RSTDESTI.Close: Set RSTDESTI = Nothing
    RSTORIGEN2.Close: Set RSTORIGEN2 = Nothing
    Call COMUNI("PROCESO TERMINADO")
    
99
    Screen.MousePointer = 1
     
     
    

End Sub

Sub CFG_CREARTABLA(NOMBRE_TABLA$, BASE_DESTINO$)

         SQLXY$ = "SELECT * FROM " & NOMBRE_TABLA$
         Dim RSTCAM As New ADODB.Recordset
         Set RSTCAM = FLEXCONN.Execute(SQLXY$)
         FIN& = RSTCAM.Fields.Count
            KU& = 0
            For KKI% = 1 To FIN&
              KU& = KU& + 1
              Call TRACE(20, KU&, FIN&)
              NOCAM$ = RSTCAM.Fields(KKI% - 1).Name
              TCAMPO = TIPODECAMPO(NOMBRE_TABLA$, NOCAM$)
              LOCAM = LONCAMPO(NOMBRE_TABLA$, NOCAM$)
              CONIN% = INDICEDECAMPO%(NOMBRE_TABLA$, NOCAM$)
              '
              TICAM = 10
              Select Case TCAMPO
                 Case 104
                   TICAM = 1
                 Case 173
                   TICAM = 2
                 Case 52
                   TICAM = 3
                 Case 56
                   TICAM = 4
                 Case 60
                   TICAM = 5
                 Case 62
                   TICAM = 7
                 Case 58
                   TICAM = 8
                 Case 231
                   TICAM = 10
                 Case 99
                   TICAM = 12
                 Case Else
                   Call MENSERR(24, "Imposible Crear Campo " + NCAMPO$ + " en Tabla: " + NOMBRE_TABLA$ + " con Tipo = " & TCAMPO)
                   Exit Sub
              End Select
              '
              Call CREACAMPO_DESTINO(BASE_DESTINO$, NOMBRE_TABLA$, NOCAM$, TICAM, LOCAM, CONIN%)
            Next KKI%
End Sub


Private Sub mnuPruebatablabasesdistintas_Click()
   If CLACONTRA% > 0 Then
       Call PruebaAgregaRegistro("AHP", "MASTER", "MINEPARTS", "MASTER")
   End If

End Sub
Sub PruebaAgregaRegistro(BaseOrigen$, TablaOrigen$, BaseDestino$, TablaDestino$)
    'esto realiza una prueba
    Origen$ = BaseOrigen$ & ".dbo." & TablaOrigen$
    DESTINO$ = BaseDestino$ & ".dbo." & TablaDestino$
    'ES PARA REALIZAR UNA PRUEBA POR QUE LA CONDI ES FIJA
    'VA A TOMAR UN REGISTRO (WHERE CODINT=1) DE LA BASE DE DATOS Y LA TABLA ORIGEN: "AHP.dbo.MASTER
    'Y LO VA A INSERTAR EN LA NUEVA BASE Y TABLA (MINEPARTS.dbo.MASTER), ANTES POR LAS DUDA LO BORRA
    'ESTO SE APLICARIA AL CASO DE SI MODIFICO PARADO EN LA BASE AHP REPLICO EN LA BASE MINEPARTS
    'ME SERVIRIA TANTO SI EDITA O AGREGA.
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    CI1X = 2
    SQLX$ = "DELETE FROM " & DESTINO$ & " WHERE CODINT=1"
    FLEXCONN.Execute (SQLX$)
    
    SQLX$ = "INSERT INTO " & DESTINO$ & " (Ferevis,CodiEmpr,Observa,Especificacion,DesCompAut,DescriComp,DepoEntra,DepoCon,Codigo,Descripcion,Descridos,Status,Codint,Alias,Umedida,Timate,Critrep,Monecos,Fecosto,PorDescCos,Cosuni,Pornacio,PorComVen,Provhabi,Stomin,Stomax,Lorepos,Lindivis,Pesuni,PesMetro,Color,Talle,Descompo,Grupiva,Grucoven,CLinea,Cfamilia,NPlano,Revision,Trazable,Revi_Formu,FechApFor,UsuApFor,Cod_Interna,Bien_Cap,Codigo_Ncm,Codigo_Sec,Pro_Umed,Iva_Id,ArchiPdf,Familia,Lindiven,CanDeci,ConsuProm,CosCal,FeCosCal,Stat_Movi,Clas_ATC,Vers_BC,FeVig_BC,FeRev_BC,CantApro,LeTiCom,PiezasXCaja,M3Caja,KiloNetoCaja,KiloBrutoCaja,Autono_Proy,UsuPrePes,UsuPreDol,MANO_OBRA_REAL,Tamano_Etiqueta,CueContaI,Molde,VALOBADA_1,VALOBADA_2,VALOBADA_3,Consumo_X_Corte,GRAMAJE,ANCHO,GR_ACTU,Cod_Arba,VALOBADA_4,VALOBADA_5,MICRONES)"
    SQLX$ = SQLX$ + " SELECT Ferevis, " & CI1X & " ,Observa,Especificacion,DesCompAut,DescriComp,DepoEntra,DepoCon,Codigo,Descripcion,Descridos,Status,Codint,Alias,Umedida,Timate,Critrep,Monecos,Fecosto,PorDescCos,Cosuni,Pornacio,PorComVen,Provhabi,Stomin,Stomax,Lorepos,Lindivis,Pesuni,PesMetro,Color,Talle,Descompo,Grupiva,Grucoven,CLinea,Cfamilia,NPlano,Revision,Trazable,Revi_Formu,FechApFor,UsuApFor,Cod_Interna,Bien_Cap,Codigo_Ncm,Codigo_Sec,Pro_Umed,Iva_Id,ArchiPdf,Familia,Lindiven,CanDeci,ConsuProm,CosCal,FeCosCal,Stat_Movi,Clas_ATC,Vers_BC,FeVig_BC,FeRev_BC,CantApro,LeTiCom,PiezasXCaja,M3Caja,KiloNetoCaja,KiloBrutoCaja,Autono_Proy,UsuPrePes,UsuPreDol,MANO_OBRA_REAL,Tamano_Etiqueta,CueContaI,Molde,VALOBADA_1,VALOBADA_2,VALOBADA_3,Consumo_X_Corte,GRAMAJE,ANCHO,GR_ACTU,Cod_Arba,VALOBADA_4,VALOBADA_5,MICRONES"
    SQLX$ = SQLX$ + " From " & Origen$ & " Where CODINT = 1"
    FLEXCONN.Execute (SQLX$)
    
    CANT111 = CantRegistros("" & Origen$ & "", "CODINT=1") '
    MsgBox "ahp " & CANT111
    cant222 = CantRegistros("" & DESTINO$ & "", "CODINT=1 and codiempr = 2") '
    MsgBox "MINEPARTS " & cant222
    
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Proceso Realizado Exitosamente")
    End If
    FLEXCONN.Close
    Call ABRECONEXION
    
End Sub
Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub tablascanje_Click()
   Call Command16_Click
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text11_Change()
  If XVALO(Text11) > 0 Then ' DEPOSITO DE ARMADOR
    If ECOARCH$("TADEPOSI", Chr$(XVALO(Text11))) <> "" Then
     Text10 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text11)))
    End If
  Else
     Text10 = ""
  End If
  
End Sub

Private Sub Text12_Click()
   TEXT13.SetFocus
End Sub

Private Sub Text13_DblClick()
    Call Command11_Click
End Sub

Private Sub Text14_Click()
  Text15.SetFocus
End Sub

Private Sub Text15_DblClick()
   Call Command13_Click
End Sub

Private Sub Text16_DblClick()
    Call Command14_Click
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text22_DblClick()
    Call Command19_Click
End Sub

Private Sub Text23_DblClick()
    Call Command21_Click
End Sub

Private Sub Text25_DblClick()
    Call Command22_Click
End Sub

Private Sub Text27_Change()
  Text77 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text27)))
End Sub

Private Sub Text28_Change()
  If XVALO(Text28) > 0 Then ' DEPOSITO DE SCRAP
    If ECOARCH$("TADEPOSI", Chr$(XVALO(Text28))) <> "" Then
     Text29 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text28)))
    End If
  Else
     Text29 = ""
  End If
End Sub

Private Sub TEXT3_Change()
  Text4 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text3)))
End Sub

'
Private Sub Text3_DblClick()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("TADEPOSI")
    End If
End Sub

Private Sub Text4_CLICK()
    Text3.SetFocus
End Sub

Private Sub Text5_Change()
  Text7 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text5)))
End Sub

Private Sub Text5_DblClick()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text5.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("TADEPOSI")
    End If

End Sub

Private Sub Text7_CLICK()
   Text5.SetFocus
End Sub

Private Sub Text8_Change()
   Text9 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text8)))
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("TADEPOSI")
    End If

End Sub

Private Sub Text9_CLICK()
   Text8.SetFocus
End Sub

Private Sub usuarioxsector_Click()
    Call Command25_Click
End Sub
