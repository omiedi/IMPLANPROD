VERSION 5.00
Begin VB.Form FORFACNEU 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8670
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11910
   Begin VB.CommandButton Command12 
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
      Height          =   275
      Left            =   11590
      Picture         =   "FORFACNEU.frx":0000
      TabIndex        =   60
      Top             =   25
      Width           =   275
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2325
      Left            =   0
      TabIndex        =   22
      Top             =   6300
      Width           =   10800
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   1635
         Left            =   5565
         ScaleHeight     =   1575
         ScaleWidth      =   60
         TabIndex        =   72
         Top             =   525
         Width           =   120
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   5000
            Left            =   0
            TabIndex        =   73
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.TextBox Total 
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
         Left            =   9030
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1485
         Width           =   1650
      End
      Begin VB.TextBox PORIVA 
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
         Left            =   7560
         TabIndex        =   68
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1485
         Width           =   1065
      End
      Begin VB.TextBox Total 
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
         Left            =   9030
         TabIndex        =   65
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1185
         Width           =   1650
      End
      Begin VB.TextBox PORIVA 
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
         Left            =   7560
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1185
         Width           =   1065
      End
      Begin VB.TextBox PORIVA 
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
         Left            =   7560
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   " "
         Top             =   885
         Width           =   1065
      End
      Begin VB.TextBox PORDESCU 
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
         Left            =   7560
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   495
         Width           =   1065
      End
      Begin VB.TextBox Total 
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
         Left            =   9030
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   " "
         Top             =   885
         Width           =   1650
      End
      Begin VB.TextBox Total 
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
         Left            =   9030
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   " "
         Top             =   495
         Width           =   1650
      End
      Begin VB.TextBox Total 
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
         Left            =   9030
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   1650
      End
      Begin VB.TextBox Total 
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
         Left            =   9030
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1890
         Width           =   1650
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
         Left            =   315
         TabIndex        =   34
         Text            =   " "
         Top             =   1890
         Width           =   4635
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
         Left            =   4410
         TabIndex        =   32
         Text            =   " "
         Top             =   1550
         Width           =   435
      End
      Begin VB.CommandButton Command21 
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
         Index           =   2
         Left            =   4830
         TabIndex        =   31
         Top             =   1550
         Width           =   180
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
         Index           =   1
         Left            =   315
         TabIndex        =   30
         Text            =   " "
         Top             =   1155
         Width           =   4635
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
         Index           =   1
         Left            =   4410
         TabIndex        =   28
         Text            =   " "
         Top             =   820
         Width           =   435
      End
      Begin VB.CommandButton Command21 
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
         Index           =   1
         Left            =   4830
         TabIndex        =   27
         Top             =   820
         Width           =   180
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
         Index           =   0
         Left            =   315
         TabIndex        =   26
         Text            =   " "
         Top             =   440
         Width           =   4635
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
         Index           =   0
         Left            =   4410
         TabIndex        =   24
         Text            =   " "
         Top             =   105
         Width           =   435
      End
      Begin VB.CommandButton Command21 
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
         Index           =   0
         Left            =   4830
         TabIndex        =   23
         Top             =   105
         Width           =   180
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   10710
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A.-R.N.I.:"
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
         Left            =   6090
         TabIndex        =   71
         Top             =   1590
         Width           =   1380
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "%:"
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
         Left            =   8715
         TabIndex        =   70
         Top             =   1590
         Width           =   225
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A.-2:"
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
         Left            =   6405
         TabIndex        =   67
         Top             =   1290
         Width           =   1065
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "%:"
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
         Left            =   8715
         TabIndex        =   66
         Top             =   1290
         Width           =   225
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "%:"
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
         Left            =   8715
         TabIndex        =   46
         Top             =   945
         Width           =   225
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "%:"
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
         Left            =   8715
         TabIndex        =   45
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Documento:"
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
         Left            =   7035
         TabIndex        =   42
         Top             =   1950
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A.-1:"
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
         Left            =   6405
         TabIndex        =   41
         Top             =   990
         Width           =   1065
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento:"
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
         Left            =   5565
         TabIndex        =   40
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Neto:"
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
         Left            =   7035
         TabIndex        =   39
         Top             =   170
         Width           =   1905
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor Asignado:"
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
         Left            =   210
         TabIndex        =   33
         Top             =   1640
         Width           =   4110
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Condición de Pago:"
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
         Left            =   210
         TabIndex        =   29
         Top             =   905
         Width           =   4110
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Precios:"
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
         Left            =   210
         TabIndex        =   25
         Top             =   180
         Width           =   4110
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   315
      Width           =   12000
      Begin VB.CommandButton Command14 
         Caption         =   "I.V.A."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   10920
         Picture         =   "FORFACNEU.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Tabla de Alícuotas de I.V.A."
         Top             =   7480
         Width           =   855
      End
      Begin VB.CommandButton Command13 
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
         Height          =   800
         Left            =   10920
         Picture         =   "FORFACNEU.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Configurar"
         Top             =   4305
         Width           =   855
      End
      Begin VB.CommandButton Command11 
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
         Height          =   800
         Left            =   10920
         Picture         =   "FORFACNEU.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Pendien"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   10920
         Picture         =   "FORFACNEU.frx":0CD8
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Remitos Pendientes de Facturación"
         Top             =   5200
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "C.Pago"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   10920
         Picture         =   "FORFACNEU.frx":0FE2
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Condiciones de Pago"
         Top             =   6790
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Precios"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   10920
         Picture         =   "FORFACNEU.frx":12EC
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Listas de Precios de Venta"
         Top             =   6100
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "C.Cte."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   10920
         Picture         =   "FORFACNEU.frx":1436
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   2550
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   10920
         Picture         =   "FORFACNEU.frx":1740
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Notas de Clientes"
         Top             =   3420
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
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
         Left            =   10920
         Picture         =   "FORFACNEU.frx":1A4A
         TabIndex        =   48
         Top             =   315
         Width           =   855
      End
      Begin VB.CommandButton BOTNEXT 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   10920
         Picture         =   "FORFACNEU.frx":1D54
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   735
         Width           =   855
      End
      Begin VB.TextBox Text11 
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
         Left            =   9240
         TabIndex        =   20
         Text            =   " "
         Top             =   1155
         Width           =   1440
      End
      Begin VB.TextBox Text10 
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
         Left            =   9030
         TabIndex        =   18
         Text            =   " "
         Top             =   735
         Width           =   1485
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
         Left            =   10500
         TabIndex        =   17
         Top             =   735
         Width           =   180
      End
      Begin VB.TextBox Text9 
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
         Left            =   9345
         TabIndex        =   15
         Text            =   " "
         Top             =   315
         Width           =   1170
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
         Left            =   10500
         TabIndex        =   14
         Top             =   315
         Width           =   180
      End
      Begin VB.TextBox Text7 
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
         Left            =   5565
         TabIndex        =   12
         Text            =   " "
         Top             =   735
         Width           =   1695
      End
      Begin VB.TextBox Text8 
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
         Left            =   5565
         TabIndex        =   11
         Text            =   " "
         Top             =   1155
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
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
         Left            =   7035
         TabIndex        =   10
         Top             =   315
         Width           =   180
      End
      Begin VB.TextBox Text6 
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
         Left            =   6615
         TabIndex        =   9
         Text            =   " "
         Top             =   315
         Width           =   435
      End
      Begin VB.TextBox Text5 
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
         Left            =   2310
         TabIndex        =   7
         Text            =   " "
         Top             =   1155
         Width           =   2850
      End
      Begin VB.TextBox Text4 
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
         Left            =   1470
         TabIndex        =   6
         Text            =   " "
         Top             =   1155
         Width           =   750
      End
      Begin VB.TextBox Text3 
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
         Left            =   1470
         TabIndex        =   5
         Text            =   " "
         Top             =   735
         Width           =   3690
      End
      Begin VB.TextBox Text2 
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
         Left            =   1470
         TabIndex        =   4
         Text            =   " "
         Top             =   315
         Width           =   3690
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
         Left            =   1050
         TabIndex        =   3
         Top             =   1155
         Width           =   180
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   2
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   4530
         Left            =   210
         ScaleHeight     =   4470
         ScaleWidth      =   10455
         TabIndex        =   49
         Top             =   1680
         Width           =   10515
         Begin VB.Frame Frame4 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   540
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   10515
            Begin VB.Line Line6 
               X1              =   8820
               X2              =   8820
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line5 
               X1              =   7560
               X2              =   7560
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line4 
               X1              =   2310
               X2              =   2310
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line1 
               X1              =   840
               X2              =   840
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   $"FORFACNEU.frx":205E
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   330
               Left            =   0
               TabIndex        =   52
               Top             =   140
               Width           =   10410
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5580
            Left            =   0
            TabIndex        =   50
            Top             =   -1050
            Width           =   10515
            Begin VB.Line Line9 
               X1              =   0
               X2              =   10500
               Y1              =   1590
               Y2              =   1590
            End
            Begin VB.Line Line2 
               Index           =   3
               X1              =   8820
               X2              =   8820
               Y1              =   525
               Y2              =   5565
            End
            Begin VB.Line Line2 
               Index           =   2
               X1              =   7560
               X2              =   7560
               Y1              =   525
               Y2              =   5565
            End
            Begin VB.Line Line2 
               Index           =   1
               X1              =   2310
               X2              =   2310
               Y1              =   525
               Y2              =   5565
            End
            Begin VB.Line Line2 
               Index           =   0
               X1              =   840
               X2              =   840
               Y1              =   525
               Y2              =   5565
            End
         End
      End
      Begin VB.Line Line8 
         X1              =   10720
         X2              =   10720
         Y1              =   1680
         Y2              =   5985
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0001"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   8520
         TabIndex        =   63
         Top             =   1155
         Width           =   585
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   9090
         TabIndex        =   62
         Top             =   1100
         Width           =   150
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Picture         =   "FORFACNEU.frx":20EF
         ToolTipText     =   "Click Aquí para cambiar Modalidad de Facturación"
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Doc.Nro:"
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
         Left            =   7560
         TabIndex        =   21
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rto./ Resum.:"
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
         Left            =   7245
         TabIndex        =   19
         Top             =   735
         Width           =   1800
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
         Height          =   330
         Left            =   8610
         TabIndex        =   16
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label2 
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
         Left            =   5670
         TabIndex        =   8
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   315
         TabIndex        =   1
         Top             =   840
         Width           =   960
      End
   End
   Begin VB.Label Label10 
      BackColor       =   &H00800000&
      Caption         =   "  Facturacion - Debitos y Creditos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   0
      TabIndex        =   59
      ToolTipText     =   "Doble Click Aquí para Cambiar Modalidad de Facturación "
      Top             =   0
      Width           =   11895
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
      TabIndex        =   13
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "FORFACNEU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTNEXT_Click()
     '
     If InStr(FORFACNEU.Label10.Caption, "Anulacion de Comprobantes") > 0 Then
         Call ANUFAC
         Exit Sub
     End If
     '
     NCII% = Val(Text1.TEXT)
     '
     If NC% < 1 Or NCII% <> NC% Then
        Text1.SetFocus
        Exit Sub
     End If
     '
     If TRIM$(Text2.TEXT) <> RASOCLI$(NC%) Then
        Text1.SetFocus
        Exit Sub
     End If
     '
     FFI1 = FECHANUM(TRIM$(FORFACNEU.Text9.TEXT))
     FECHALIM% = Val(CONTROL$("", "CIEMEVEN"))
     If FFI1 < 1 Or FFI1 <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        FORFACNEU.Text9.SetFocus
        Exit Sub
     End If
     '
     Call SIGUEFAC1
     '
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("DEUDOR1")
    NC% = Val(TRIM$(VALACT1$("DEUDOR1")))
    Text1.TEXT = TRIM$(Str$(NC%))
    Call CARDACLI
    '
End Sub

Private Sub COMMAND11_CLICK()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("AMACLI/NEW", "Alta de Nueva Cuenta de Cliente")
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    Call BLANPEDP
    Call FINAL("")
End Sub


Private Sub Command13_Click()
   Command13.Enabled = False
   If USNBR% Mod 100 <> 1 Then
      Call MENSERR(24, "La Opcion Elegida Requiere\Privilegios de Supervisor.")
      Command13.Enabled = True
      Exit Sub
   End If
   '
   MODOPER% = ALTERNA%("Tipos de Documento\Sucursales Remotas")
   If MODOPER% = 1 Then
        Call TRALOCAL("TIDOCFAC", "TIDOCFAC")
        VTB% = VENTABU%("TIDOCFA1")
        If VTB% >= 0 Then
           VTB% = VENTABU%("TIDOCFA2")
           If VTB% >= 0 Then
              Call TRACENTRAL("TIDOCFAC", "TIDOCFAC")
           End If
        End If
      ElseIf MODOPER% = 2 Then
        Call TRALOCAL("SUCREMOT", "SUCREMOT")
        VTB% = VENTABU%("SUCREMOT")
        If VTB% >= 0 Then
           Call TRACENTRAL("SUCREMOT", "SUCREMOT")
        End If
   End If
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call CARTASIVA
   Command14.Enabled = True
End Sub

Private Sub Command2_Click()
    Call BLANPEDP
    Call FINAL("")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NC% = Val(TRIM$(Text1.TEXT))
    If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
       Call NOTACLI(NC%, 2)
    End If
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RESCUDEU.Text1.TEXT = Text1.TEXT
    Hide
    RESCUDEU.Show
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call CONECRUN("ADMIPRE", "Listas de Precios de Venta")
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Call SELECHO("CAPOSIV")
    If TRIM$(VALACT1$("CAPOSIV")) <> "" Then
        Text6.TEXT = TRIM$(VALACT1$("CAPOSIV"))
      Else
        Text6.TEXT = "1"
    End If
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call CARCONPAGOS 'CONECRUN("#SVD151", "Condiciones de Pago Clientes")
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Screen.MousePointer = 11
    V& = 0
    For U& = 1 To ULTREG&("REMPEN")
        X$ = REGLEIDO$("REMPEN", U&)
        If CVI(Left$(X$, 2)) > 0 Then
            If CVS(Mid$(X$, 27, 4)) >= 0.05 Then
                V& = V& + 1
                Call GRAREG("REMPEN", X$, V&)
            End If
        End If
    Next U&
    Call SETULTREG("REMPEN", V&)
    Call CIERRARCH("REMPEN")
    Screen.MousePointer = 1
    '
    MDX% = ALTERNA%("Archivo Completo\Por Cliente\Por Número de Remito")
    If MDX% = 1 Then
        Call CONECRUN("*COREPEN", "Remitos Pendientes de Facturación")
      ElseIf MDX% = 2 Then
        Screen.MousePointer = 11
        Call COPYSTRU("DEUDOR1", "DEUDOREM")
        Dim NCAC%(16384)
        For U& = 1 To ULTREG&("REMPEN")
            X$ = REGLEIDO$("REMPEN", U&)
            NCX% = CVI(Left$(X$, 2))
            If NCX% > 0 Then
                If NCX% <= 16384 Then
                    NCAC%(NCX%) = 1
                End If
            End If
        Next U&
        '
        V& = 0
        For U& = 1 To ULTREG&("DEUDOR1")
            X$ = REGLEIDO$("DEUDOR1", U&)
            NCX% = CVI(Left$(X$, 2))
            If NCX% > 0 Then
                If NCX% <= 16384 Then
                    If NCAC%(NCX%) > 0 Then
                        V& = V& + 1
                        Call GRAREG("DEUDOREM", X$, V&)
                    End If
                End If
            End If
        Next U&
        '
        Call SETULTREG&("DEUDOREM", V&)
        Call CIERRARCH("DEUDOREM")
        Call FRESHECHO("DEUDOREM")
        Screen.MousePointer = 1
        '
        Call CONECRUN("*REPECLI", "Remitos Pendientes por Cliente")
      ElseIf MDX% = 3 Then
        Call CONECRUN("*REMPEN", "Remitos Pendientes por Numero")
    End If
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    Text9.TEXT = VALACT1$("SELFECHA")
    FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   TEXT21(Index).SetFocus
End Sub

Private Sub Form_Load()
    '
    'EPAC$ = TRIM$(EMPREAC$)
    'If EPAC$ <> "" Then
    '    Caption = Caption + "  -  " + EPAC$
    'End If
    '
    'Call VERJOBID
    '
    J& = 0
    TDF% = 0
10  FIN& = ULTREG&("TIDOCFAC")
    For U& = 1 To FIN&
        X$ = REGLEIDO$("TIDOCFAC", U&)
        If Asc(Left$(X$, 1)) = TDF% Then
            J& = J& + 1
            Y$ = Space$(128)
            X1$ = REPLACAR$(TRIM$(Mid$(X$, 2, 23)) + " - " + TRIM$(Mid$(X$, 33, 32)), Chr$(160), "a")
            X2$ = REPLACAR$(X1$, Chr$(130), "e")
            X3$ = REPLACAR$(X2$, Chr$(162), "o")
            Call REPLA(Y$, X3$, 1, 56)
            Call REPLA(Y$, Mid$(X$, 65, 64), 63, 64)
            Call REPLA(Y$, TRIM$(Str$(TDF%)), 127, 1)
            Call REPLA(Y$, TRIM$(Str$(Asc(Mid$(X$, 129, 1)))), 128, 1)
            Call GRAREG("SELTIFAC", Y$, J&)
            If J& = 1 Then
                TDFAC$ = Y$
            End If
        End If
    Next U&
    If TDF% < 3 Then
        TDF% = TDF% + 1
        GoTo 10
    End If
    J& = J& + 1
    Y$ = REGBLAN$("SELTIFAC")
    Call REPLA(Y$, "Anulacion de Comprobantes", 1, 56)
    Call GRAREG("SELTIFAC", Y$, J&)
    '
    Call SETULTREG("SELTIFAC", J&)
    Call CIERRARCH("SELTIFAC")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
          ENCABE$ = "  " + EPAC$ + "  -  "
        Else
          ENCABE$ = "  "
    End If
    ENCABE$ = ENCABE$ + TRIM$(Left$(TDFAC$, 56))
    FORFACNEU.Label10.Caption = ENCABE$
    '
    RGTIFA& = 1
    CONTROFA$ = REGLEIDO$("TIDOCFAC", RGTIFA&)
    Call CIERRARCH("TIDOCFAC")
    '
    Call VAPUBLI
    '
    If ULTREG&("REMPEN") < 1 Then
        Command8.Enabled = False
    End If
    '
    Call LEETASIVA
    '
End Sub
'

Private Sub Image1_Click()
    '
10  Call SELTIFAC
    If InStr(FORFACNEU.Label10.Caption, "Anulacion de Comprobantes") = 0 Then
       Call VAPUBLI
     Else
       Call ANUFAC
       Call SELTIFAC
       GoTo 10
    End If
    '
End Sub

Private Sub Label10_DblClick()
10  Call SELTIFAC
    If InStr(FORFACNEU.Label10.Caption, "Anulacion de Comprobantes") = 0 Then
       Call VAPUBLI
     Else
       Call ANUFAC
       Call SELTIFAC
       GoTo 10
    End If
End Sub

Private Sub Label12_DblClick()
   '
   If ULTREG&("SUCREMOT") > 1 Then
      XXX$ = REGSEL$("SUCREMOT")
      If Len(XXX$) > 4 Then
         COSUCU$ = TRIM$(Str$(CVI(Left$(XXX$, 2))))
         While Len(COSUCU$) < 4: COSUCU$ = "0" + COSUCU$: Wend
         Label12.Caption = COSUCU$
         Label12.Refresh
      End If
   End If
   NC% = Val(Text1.TEXT)
   Call CARDACLI
   '
End Sub

Private Sub PORDESCU_Click()
   If DERACCE%("CAMBDESC", "Modificar Descuento de Cliente") < 2 Then
      Text1.SetFocus
   End If
End Sub

Private Sub PORDESCU_LostFocus()
    If TRIM$(PORDESCU.TEXT) <> TRIM$(DESCUCLI$(NC%)) Then
        If ALTERNA%("Registra este Valor como Standard") = 1 Then
            DESCUNUE0$ = TRIM$(PORDESCU.TEXT)
            DESCUNUE$ = REPLACAR$(DESCUNUE0$, "%", " ")
            For U& = 1 To ULTREG("DESCUCLI")
                X1$ = REGLEIDO$("DESCUCLI", U&)
                If CVI(Left$(X1$, 2)) = NC% Then
                    Call REPLA(X1$, DESCUNUE$, 3, 14)
                    Call GRAREG("DESCUCLI", X1$, U&)
                    Call CIERRARCH("DESCUCLI")
                    Exit Sub
                End If
            Next U&
            X1$ = MKI$(NC%) + DESCUNUE$
            Call REGAPP("DESCUCLI", X1$)
            Call CIERRARCH("DESCUCLI")
        End If
    End If
End Sub
'

Private Sub PORIVA_GotFocus(Index As Integer)
   Command14.SetFocus
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("DEUDOR1")
    NC% = Val(TRIM$(VALACT1$("DEUDOR1")))
    Text1.TEXT = TRIM$(Str$(NC%))
    Call CARDACLI
    '
End Sub

Private Sub Text1_GotFocus()
    If TICA < 0.5 Then
        Call FORMPRIN(LU$, FLIM%, ALI1, ALI2, TC1, TLET$, INTER%, TFOR%, ANFO%, LARFO%, DIFER, CAFOR%, MARIZ%, MARINT%, REDON%, MPRI$, MCEN$, FACTO, CANCON%, COMA%, PARDEC%)
        TICA = TC1
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Val(Text1.TEXT) > 0 Then
          If Val(Text1.TEXT) <= 32767 Then
             NC% = Val(Text1.TEXT)
             If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
                Call CARDACLI
                BOTNEXT.SetFocus
                Exit Sub
             End If
          End If
       End If
       '
       Beep
       Text1.TEXT = ""
       '
    End If
    '
End Sub

Private Sub Text11_LostFocus()
    '
    LU$ = LUDAT$
    CLI$ = RASOCLI$(NC%)
    FLIM% = FECHANUM("01/01/99")
    If ULTREG&("FORMFAC") > 0 Then
        FFS$ = REGLEIDO$("FORMFAC", 2&)
        FLIM% = CVI(Mid$(FFS$, 19, 2))
    End If
    '
    If Val(Text11.TEXT) < 1 Then
        Call MENSERR(24, "Numero de Documento no Permitido")
        Call GRAFAC98(LU$, ARCHIFAC$, 1, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDE%, TOTAFAC#())
        '
        NROPRO = NRO
        FORFACNEU.Text11.TEXT = NROPRO
        Exit Sub
    End If
    '
    NRO = Val(Text11.TEXT)
    FLIM% = FECHANUM("01/01/99")
    If ULTREG&("FORMFAC") > 0 Then
        FFS$ = REGLEIDO$("FORMFAC", 2&)
        FLIM% = CVI(Mid$(FFS$, 19, 2))
    End If
    '
    Call GRAFAC98(LU$, ARCHIFAC$, 2, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDE%, TOTAFAC#())
    '
    If NRO <> Val(Text11.TEXT) Then
        Call MENSERR(24, "Numero de Documento Pre-existente")
        Text11.TEXT = NRO
        Exit Sub
    End If
    '
End Sub

Private Sub Text6_Change()
    PIVA% = Val(Text6.TEXT)
    If ECOARCH$("CAPOSIV", Chr$(PIVA%)) <> "" Then
        XX$ = REGACT$("CAPOSIV")
        Text7.TEXT = UCase$(Mid$(XX$, 33, 8))
        FACIVA = 1
        CATIVA$ = "A"
        VA% = 1
        If PIVA% = 0 Or PIVA% = 3 Or PIVA% = 4 Or PIVA% = 6 Then
           'FACIVA = 1 + ALICU / 100
           CATIVAS$ = "B"
           VA% = 2
           '
          ElseIf PIVA% = 2 Then
           'ALICU = ALI1 + ALI2
           RNI% = 1
           '
          ElseIf PIVA% = 5 Then
           'ALICU = 0
           CATIVAS$ = " "
           VA% = 0
           If LETREXPO$ = "A" Then
              VA% = 1
              CATIVAS$ = "A"
           End If
        End If
        Call VAPUBLI
      Else
        Text6.TEXT = "1"
    End If
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("CAPOSIV")
    If TRIM$(VALACT1$("CAPOSIV")) <> "" Then
        Text6.TEXT = TRIM$(VALACT1$("CAPOSIV"))
      Else
        Text6.TEXT = "1"
    End If
End Sub

Private Sub Text21_Change(Index As Integer)
    '
    If Index = 0 Then
        If Val(TEXT21(Index).TEXT) < 1 Then
            Eco21(Index).TEXT = "CONCERTADOS FIJOS"
          Else
            Eco21(Index).TEXT = ECOARCH$("LISTAS", MKI$(Val(TEXT21(Index).TEXT)))
        End If
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(Val(TEXT21(Index).TEXT)))
      ElseIf Index = 2 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(Val(TEXT21(Index).TEXT)))
    End If
    '
End Sub

Private Sub Command21_Click(Index As Integer)
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
    TEXT21(Index).TEXT = VALACT1$(ARCHEC$)
    '
End Sub

Sub CARTASIVA()
   Call TRALOCAL("GRUPIVA", "")
   VTB% = VENTABU%("GRUPIVA/NC/TITUPAN=Tasa IVA por Categoria")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!GRUPIVA")
         X$ = REGLEIDO$("!GRUPIVA", U&)
         Call REPLA(X$, Chr$(U& - 1), 1, 1)
         Call GRAREG("GRUPIVA", X$, U&)
      Next U&
      FIN& = ULTREG&("!GRUPIVA")
      Call SETULTREG("GRUPIVA", FIN&)
   End If
   If ULTREG&("GRUPIVA") > 3 Then
      Call SETULTREG("GRUPIVA", 3)
   End If
   Call CIERRARCH("GRUPIVA")
End Sub
'
Sub LEETASIVA()
   For U& = 1 To 2
      X$ = REGLEIDO$("GRUPIVA", U&)
      TAIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
      PORIVA(U& - 1) = CSTRING$(MKS$(TAIVA(U& - 1)), 3, 6, 2, 2)
   Next U&
   Call CIERRARCH("GRUPIVA")
End Sub

