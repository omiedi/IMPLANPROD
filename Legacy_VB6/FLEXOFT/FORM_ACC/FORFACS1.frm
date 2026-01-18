VERSION 5.00
Begin VB.Form FORFACS1 
   Caption         =   "Facturación en Cuenta Corriente"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8730
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12735
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFFF80&
         Height          =   545
         Left            =   1470
         Picture         =   "FORFACS1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Pedidos pendientes de Despacho, Remito o Factura"
         Top             =   630
         Width           =   550
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
         Height          =   300
         Left            =   3840
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   960
      End
      Begin VB.TextBox Text11 
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
         Left            =   4365
         TabIndex        =   68
         Text            =   " "
         Top             =   1440
         Width           =   1120
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00400000&
         Height          =   7215
         Left            =   5775
         ScaleHeight     =   7155
         ScaleWidth      =   840
         TabIndex        =   56
         Top             =   0
         Width           =   900
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
            Height          =   640
            Left            =   105
            Picture         =   "FORFACS1.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Acceso al Control General de Gestión Comercial"
            Top             =   3510
            Width           =   650
         End
         Begin VB.CommandButton Command20 
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
            Height          =   640
            Left            =   105
            Picture         =   "FORFACS1.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Opciones Varias - Anulación / Control Diferencia de Cambio"
            Top             =   4905
            Width           =   650
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
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
            Height          =   690
            Left            =   105
            Picture         =   "FORFACS1.frx":0896
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   100
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
            Height          =   640
            Left            =   105
            Picture         =   "FORFACS1.frx":09E0
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   2860
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
            Height          =   640
            Left            =   105
            Picture         =   "FORFACS1.frx":0CEA
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   5670
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
            Height          =   640
            Left            =   105
            Picture         =   "FORFACS1.frx":112C
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Acceso a Listas de Precio, Condiciones de Pago y Tasas de IVA"
            Top             =   4265
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
            Height          =   640
            Left            =   105
            Picture         =   "FORFACS1.frx":1436
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Remitos Pendientes de Facturación"
            Top             =   2220
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
            Height          =   640
            Left            =   105
            Picture         =   "FORFACS1.frx":1740
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Acceso a Bases de Datos Maestras"
            Top             =   1580
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
            Height          =   690
            Left            =   105
            Picture         =   "FORFACS1.frx":1A4A
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   790
            Width           =   650
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   120
            ScaleHeight     =   75
            ScaleWidth      =   540
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   6510
            Width           =   600
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   58
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -315
            Picture         =   "FORFACS1.frx":1D54
            Top             =   6760
            Width           =   1515
         End
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
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
         Height          =   855
         Left            =   5775
         Picture         =   "FORFACS1.frx":3C76
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Aprueba, Graba e Imprime Documento"
         Top             =   7455
         Width           =   800
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
         Height          =   300
         Left            =   4800
         TabIndex        =   54
         Top             =   840
         Width           =   180
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2745
         Left            =   210
         TabIndex        =   40
         Top             =   1785
         Width           =   5325
         Begin VB.TextBox Text12 
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
            Left            =   105
            MaxLength       =   40
            TabIndex        =   49
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1890
            Width           =   5055
         End
         Begin VB.TextBox Text5 
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
            Left            =   1785
            TabIndex        =   48
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1470
            Width           =   3375
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
            Height          =   300
            Left            =   105
            TabIndex        =   47
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1470
            Width           =   1485
         End
         Begin VB.TextBox Text3 
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
            Left            =   105
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1050
            Width           =   5055
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
            Height          =   300
            Left            =   105
            TabIndex        =   45
            TabStop         =   0   'False
            Text            =   " "
            Top             =   630
            Width           =   5055
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
            Left            =   1680
            TabIndex        =   44
            Top             =   210
            Width           =   180
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
            Height          =   300
            Left            =   840
            MaxLength       =   5
            TabIndex        =   43
            Text            =   " "
            Top             =   210
            Width           =   855
         End
         Begin VB.TextBox Text7 
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
            Left            =   1365
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2310
            Width           =   1380
         End
         Begin VB.TextBox Text8 
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
            Left            =   3990
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2310
            Width           =   1170
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   330
            Left            =   105
            TabIndex        =   53
            Top             =   315
            Width           =   960
         End
         Begin VB.Label Label26 
            BackColor       =   &H00EEEEEE&
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
            Left            =   2100
            TabIndex        =   52
            Top             =   210
            Width           =   1035
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Lim.Crédito:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   330
            Left            =   60
            TabIndex        =   51
            Top             =   2310
            Width           =   1380
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Disponible:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   435
            Left            =   2700
            TabIndex        =   50
            Top             =   2310
            Width           =   1380
         End
      End
      Begin VB.Frame Frame7 
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
         ForeColor       =   &H00004040&
         Height          =   1485
         Left            =   210
         TabIndex        =   26
         Top             =   5500
         Width           =   5330
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
            Height          =   270
            Left            =   1155
            TabIndex        =   34
            Top             =   775
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
            Height          =   270
            Left            =   1785
            TabIndex        =   33
            Top             =   775
            Width           =   180
         End
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
            Height          =   270
            Left            =   1785
            TabIndex        =   32
            Top             =   1100
            Width           =   180
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
            Height          =   270
            Left            =   1155
            TabIndex        =   31
            Top             =   1090
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
            Index           =   3
            Left            =   1575
            MaxLength       =   32
            TabIndex        =   30
            Top             =   1800
            Visible         =   0   'False
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
            Height          =   270
            Index           =   0
            Left            =   1155
            MaxLength       =   48
            TabIndex        =   29
            Top             =   210
            Width           =   3990
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
            Height          =   270
            Index           =   1
            Left            =   1155
            MaxLength       =   48
            TabIndex        =   28
            Top             =   490
            Width           =   3990
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
            Height          =   270
            Index           =   2
            Left            =   2050
            MaxLength       =   48
            TabIndex        =   27
            Top             =   1090
            Width           =   3095
         End
         Begin VB.Label Label32 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   420
            TabIndex        =   39
            Top             =   525
            Visible         =   0   'False
            Width           =   540
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
            Height          =   270
            Left            =   2055
            TabIndex        =   38
            Top             =   780
            Width           =   3095
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
            Left            =   -840
            TabIndex        =   37
            Top             =   860
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
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
            Index           =   12
            Left            =   -840
            TabIndex        =   36
            Top             =   315
            Width           =   1905
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
            Index           =   11
            Left            =   -840
            TabIndex        =   35
            Top             =   1155
            Width           =   1905
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Condiciones Comerciales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   990
         Left            =   210
         TabIndex        =   13
         Top             =   4515
         Width           =   5330
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   1130
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
            Height          =   270
            Index           =   1
            Left            =   1890
            TabIndex        =   19
            Text            =   " "
            Top             =   600
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
            Height          =   285
            Index           =   1
            Left            =   1155
            TabIndex        =   18
            Text            =   " "
            Top             =   600
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
            Height          =   270
            Index           =   1
            Left            =   1575
            TabIndex        =   17
            Top             =   600
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
            Height          =   270
            Index           =   0
            Left            =   1890
            TabIndex        =   16
            Text            =   " "
            Top             =   315
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
            Height          =   270
            Index           =   0
            Left            =   1155
            TabIndex        =   15
            Text            =   " "
            Top             =   315
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
            Height          =   270
            Index           =   0
            Left            =   1575
            TabIndex        =   14
            Top             =   315
            Width           =   180
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
            TabIndex        =   25
            Top             =   1205
            Width           =   4110
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.Pago:"
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
            Index           =   1
            Left            =   -3045
            TabIndex        =   24
            Top             =   675
            Width           =   4110
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L.Precios:"
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
            Index           =   0
            Left            =   -525
            TabIndex        =   23
            Top             =   390
            Width           =   1590
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   660
         Left            =   210
         TabIndex        =   4
         Top             =   6980
         Width           =   5330
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
            TabIndex        =   9
            Top             =   240
            Width           =   180
         End
         Begin VB.TextBox Text16 
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
            Left            =   1155
            TabIndex        =   8
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
            TabIndex        =   7
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
            TabIndex        =   6
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
            TabIndex        =   5
            Text            =   " "
            Top             =   1130
            Width           =   3270
         End
         Begin VB.Label Label16 
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
            Left            =   2600
            TabIndex        =   12
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta:"
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
            Left            =   -525
            TabIndex        =   11
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
            TabIndex        =   10
            Top             =   1205
            Width           =   4110
         End
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Continuar"
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
         Left            =   840
         TabIndex        =   3
         Top             =   7875
         Width           =   1485
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Cancelar"
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
         Left            =   3360
         TabIndex        =   2
         Top             =   7875
         Width           =   1485
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   315
         TabIndex        =   77
         Top             =   105
         Width           =   855
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   355
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   315
         Shape           =   4  'Rounded Rectangle
         Top             =   290
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   210
         Shape           =   4  'Rounded Rectangle
         Top             =   210
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   -5
         X2              =   5775
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   2887
         X2              =   2887
         Y1              =   780
         Y2              =   1260
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   2887
         X2              =   2887
         Y1              =   0
         Y2              =   100
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
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   3045
         TabIndex        =   75
         Top             =   910
         Width           =   750
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   3690
         TabIndex        =   74
         ToolTipText     =   "Doble-Click Para Cambiar Punto de Venta"
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Factura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   3150
         TabIndex        =   73
         Top             =   105
         Width           =   1590
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   7665
         TabIndex        =   72
         Top             =   885
         Width           =   1275
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H000000C0&
         Height          =   750
         Left            =   2570
         TabIndex        =   71
         Top             =   0
         Width           =   645
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   705
         Left            =   2570
         Top             =   105
         Width           =   645
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº - COMPROBANTE: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   315
         TabIndex        =   70
         Top             =   1365
         Width           =   3270
      End
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FFFF80&
      Height          =   545
      Left            =   2340
      Picture         =   "FORFACS1.frx":3F80
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Pedidos pendientes de Despacho, Remito o Factura"
      Top             =   840
      Width           =   555
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
End
Attribute VB_Name = "FORFACS1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TIFADOC%    ' 2=FACTURA / 3=N.DEBITO / 6=N.CREDITO
Dim MODODOC%    ' 0=MAT.CODIFICADO / 1=TEXTO LIBRE
Dim VARIDOC%
Dim RNIDOC%
Dim DATEDOC%
Dim PIVADOC%
Dim VERLICRE%
Dim MODESCU%
Dim FVENCIM
'
Dim CONSICLI%, PREFACTU$
'
Dim DESELEC#, HASELEC#
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
Dim APLIVA%, RNI%
'
Dim COPIASREM%
'
Dim TASACLICK


Private Sub Command1_Click()
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% > 0 Then
     Text1.TEXT = TRIM$(Str$(NCLIE%))
     If XVALO(Text1.TEXT) > 0 Then
       If XVALO(Text1.TEXT) <= 32767 Then
         NC% = XVALO(Text1)
         'If REGICLI&(NC%) > 0 Then
           Call CARDACLI
         'End If
       End If
     End If
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG04", "Archivos Maestros")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Call SELECHO("ECUECON")
   VDEVU$ = TRIM$(VALACT1$("ECUECON"))
   If VDEVU$ <> "" Then
     Text16 = VDEVU$
   End If
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   COGEVE04.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
'   If VALICLIE% < 1 Then Exit Sub
'   'If Abs(XVALO(Total(6))) < 0.005 Then Exit Sub
'   '
'   Command14.Enabled = False
'   '\\\OT-07-0023-RG-24/01/07///
'    FINFACTURA.Text7 = TRIM$(Str$(DEPOEXPE%))
'   '\\\OT-07-0023-RG-FIN///
'   If XVALO(TLABEL28) < 0 Then
'      Call MENSERR(24, "Imposible Registrar este Documento.\Importe No Gravado no Admisible.")
'      Command14.Enabled = True
'      Exit Sub
'   End If
'   '
'   ' 5/12/07 - SILCAR - ERROR DE IMPRESION DE IMPORTE EN NOTA DE CREITO TIPO B POR CH. RECH.
'   If TIFADOC% > 2 Then         ' solo débitos y creditos
'     If VARIDOC% = 2 Then       ' solo tipo 'B'
'       If MODODOC% = 1 Then     ' solo texto libre
'         ALIAPLI = XVALO(PORIVA(0))
'         If XVALO(Total(2)) > 0 Then
'           ALIAPLI = 100 * (XVALO(Total(3)) + XVALO(Total(4))) / XVALO(Total(2))
'         End If
'         For KU& = 1 To ULTREG&("!CARDEFAC")
'           XXX$ = REGLEIDO$("!CARDEFAC", KU&)
'           INETO1# = CVD(Mid$(XXX$, 433, 8))
'           Call REPLA(XXX$, MKS$(ALIAPLI), 441, 4)
'           Call REPLA(XXX$, MKD$(INETO1# * ALIAPLI / 100), 445, 8)
'           Call REPLA(XXX$, MKD$(INETO1# * (1 + ALIAPLI / 100)), 453, 8)
'           Call GRAREG("!CARDEFAC", XXX$, KU&)
'         Next KU&
'       End If
'     End If
'   End If
'   '
'   If XVALO(Text15) >= 0.005 Then
'     If XVALO(Total(4)) >= 0.005 Then
'       Call MENSERR(24, "Imposible Registrar este Documento.\Incompatibilidad entre Percepcion IB\y Tasa IVA Secundaria - Consulte.")
'       Command14.Enabled = True
'       Exit Sub
'     End If
'     CUEPERI$ = Control$("PERIBRU", "CUECONTA")
'     CUENPER% = CUECOINT%(CUEPERI$)
'     If TRIM$(CUEPERI$) = "" Or CUENPER% < 1 Then
'       Call MENSERR(24, "Imposible Registrar este Documento.\Falta Cuenta Contable Correspondiente\a Percepción de Ingresos Brutos.")
'       Command14.Enabled = True
'       Exit Sub
'     End If
'   End If
'   '
'   FECHALIM% = XVALO(Control$("", "CIEMEVEN"))
'   If FECHANUM(Text9) <= FECHALIM% Then
'     Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
'     Command14.Enabled = True
'     Exit Sub
'   End If
'   '
'18 EPAC$ = TRIM$(UCase$(EMPREAC$))
'   '
'   FINFACEX$ = ""
'   MONECO$ = MONEMIS
'   '
'20 NCLIE% = XVALO(Text1.TEXT)
'   If VERLICRE% = 1 Then
'     If TIFADOC% < 4 Then ' EXCLUYE CREDITOS
'       '
'       For U& = 1 To ULTREG&("!CARDEFAC")
'         X$ = REGLEIDO$("!CARDEFAC", U&)
'         NROPEDXX& = XVALO(TRIM$(Mid$(X$, 371, 16)))
'         If NROPEDXX& < 1 Then GoTo 22
'         RFF$ = RECFILT$("AUTORENT", 1, MKS$(NROPEDXX&))
'         If Len(RFF$) < 4 Then GoTo 22 ' NO AUTORIZADO
'         REGAUT& = CVS(RFF$)
'         XAUT$ = REGLEIDO$("AUTORENT", REGAUT&)
'         If UCase$(Mid$(XAUT$, 31, 2)) <> "SI" Then GoTo 22
'       Next U&
'       GoTo 25
'       '
'22     LMCR# = LIMICRE(NCLIE%)
'       SDEU# = SALDCLI#(NCLIE%)
'       CHPA# = CHEPEACRE#(NCLIE%)
'       TPED# = XVALO(Total(6).TEXT) * TICAMBX#
'       If SDEU# + CHPA# + TPED# > LMCR# Then
'         TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2))
'         If TLIMI$ = "" Then TLIMI$ = "0.00"
'         Call COMUNI("Imposible Emitir Este Documento.\  \Supera el Límite de Crédito Asignado\para esta Cuenta de $ " + TLIMI$ + ".\  \Informe al Supervisor de Ventas.")
'         Command14.Enabled = True
'         Exit Sub
'       End If
'     End If
'   End If
'   '
'' ojo aca grabar lugares de entrega
'25 Call APERBASDAT("CLIEN")
'   With Deudor12
'     .FindFirst "Nume_Cli = " & NCLIE%
'     If .NoMatch = False Then
'       .Edit
'       !Lentre0_Cli = Left$(LENTREGA(0) + " ", 48)
'       !Lentre1_Cli = Left$(LENTREGA(1) + " ", 48)
'       !Lentre2_Cli = Left$(LENTREGA(2) + " ", 48)
'       .Update
'     End If
'   End With
'   '
'   NUOCOMPRA$ = ""
'   NROPEDIDO$ = ""
'   NROREMICO$ = ""
'   '
'   For U& = 1 To ULTREG&("!CARDEFAC")
'     X$ = REGLEIDO$("!CARDEFAC", U&)
'     NUOCOPAR$ = TRIM$(Mid$(X$, 355, 16))
'     NROPEDIPAR$ = TRIM$(Mid$(X$, 371, 16))
'     NROREMIPAR$ = TRIM$(Mid$(X$, 390, 13))
'     '
'     If NUOCOPAR$ <> "" Then
'       If NUOCOPAR$ <> NUOCOPARA$ Then
'         If NUOCOMPRA$ <> "" Then NUOCOMPRA$ = NUOCOMPRA$ + ", "
'         NUOCOMPRA$ = NUOCOMPRA$ + NUOCOPAR$
'         NUOCOPARA$ = NUOCOPAR$
'       End If
'     End If
'     '
'     If NROPEDIPAR$ <> "" Then
'       If NROPEDIPAR$ <> NROPEDIPARA$ Then
'         If NROPEDIDO$ <> "" Then NROPEDIDO$ = NROPEDIDO$ + ", "
'         NROPEDIDO$ = NROPEDIDO$ + NROPEDIPAR$
'         NROPEDIPARA$ = NROPEDIPAR$
'       End If
'     End If
'     '
'     If CONSICLI% = 1 Then
'       If NROREMIPAR$ <> "" Then
'         If NROREMIPAR$ <> NROREMIPARA$ Then
'           If NROREMICO$ <> "" Then NROREMICO$ = NROREMICO$ + ", "
'           NROREMICO$ = NROREMICO$ + NROREMIPAR$
'           NROREMIPARA$ = NROREMIPAR$
'         End If
'       End If
'     End If
'     '
'   Next U&
'   '
'30 If CONSICLI% = 1 Then
'     REMI = 0
'     '
'     If TRIM$(Text14.TEXT) = "" Then
'       Text14.TEXT = NUOCOMPRA$
'     End If
'     '
'   End If
'   '
'   FORIPRE$ = FORELFAC$
'   If FORIPRE$ = Null Then Exit Sub
'   '
'   ' control de renglones por hoja
'   If Control$("FACTUREM", "FORMULAR") = "HOJASUEL" Then
'       HOJASUEL% = 1
'       CARENGLO& = ULTREG&("!CARDEFAC")
'       ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
'       If CARENGLO& > ITPORHOJA Then
'          Call MENSERR(24, "Imposible Emitir este Documento.\Supera Máximo de Renglones por Hoja.")
'          Command14.Enabled = True
'          Exit Sub
'       End If
'   End If
'   '
'   If MODODOC% = 0 Then    ' Material Codificado
'     If PREFACTU$ = "REMITO" Then     ' QUE ORDENE SOLO EN GUIDI
'       ARCHILOC$ = "!CARDEFAC"
'       Call CIERRARCH(ARCHILOC$)
'       If INTERFI$ = "" Then
'         INTERFI$ = "TRAN" + RANDSTRING$(-4)   ' negativo para que sean numeros
'       End If
'       ALOCDA$ = Mid$(ARCHILOC$, 2)
'       Call COPYSTRU(ALOCDA$, INTERFI$)
'       XY1% = COPYFILE(LUCOM$ + ALOCDA$ + ".WKF", LUDAT$ + INTERFI$ + ".DAT")
'       '
'       Call FILESORT(INTERFI$, "", 30, 30, "ASC")   ' POR NUMERO ORDEN ITEM
'       Call FILESORT(INTERFI$, "", 18, 18, "ASC")   ' POR REMITO
'       Call FILESORT(INTERFI$, "", 20, 20, "ASC")   ' POR FECHA REMITO
'       XY1% = COPYFILE(LUDAT$ + INTERFI$ + ".DAT", LUCOM$ + ALOCDA$ + ".WKF")
'       Call BAJALDISCO
'     End If
'   End If
'   '
'   CARENGLO& = ULTREG&("!CARDEFAC")
'   ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
'   If ITPORHOJA% < 1 Then
'     ITPORHOJA% = 32767
'   End If
'   '
'   CAHOJNEC% = Int(0.98 + CARENGLO& / ITPORHOJA%)
'   Call GRAFACTURA(OKX%)
'   Call CIERRARCH("*.*")
'   Call LIBARCH("*.*")
'   Call BAJALDISCO
'   '
'   If OKX% < 0 Then
'       GoTo 10     ' ABORTADA O SOLO REMITO
'     ElseIf OKX% = 0 Then
'       GoTo 18
'   End If
'   '
'   If OKX% > 0 Then
'       FF% = FECHANUM(Text9.TEXT)
'       REMI = XVALO(Text10.TEXT)
'       Call PRIFACTURA
'   End If
'   Call CIERRARCH("*.*")
'   Call LIBARCH("*.*")
'   Call BAJALDISCO
'   '
'   If OKX% > 0 Then
'     If InStr(UCase$(Eco21(1)), "CONTADO") > 0 Then
'       RECIBO04.Text1 = Text1
'       RECIBO04.Text2 = Text2
'       RECIBO04.Text3 = Text3
'       RECIBO04.Text4 = Text4
'       RECIBO04.Text5 = Text5
'       RECIBO04.Text9 = Text9
'       If ULTREG&("COBRADOR") > 0 Then
'         XXX$ = REGLEIDO$("COBRADOR", 1)
'         RECIBO04.Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
'       End If
'       RECIBO04.Show 1
'     End If
'   End If
'   '
'10 Call BLANFORMU
'   Call NUFACTURA
'   Call CIERRARCH("*.*")
'   Text1.TEXT = ""
'   On Error Resume Next
'   Text1.SetFocus
'   On Error GoTo 0
'   '
'   Command14.Enabled = True
End Sub



Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command20_Click()
   Command20.Enabled = False
   OPVARFA04.Show 1
   Command20.Enabled = True
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
End Sub

Private Sub Command24_Click()
   '\\\OT-06-0570-RG-27/12/06///
   Call CommSel("TRANSPOR", Command24, Text19)
   '\\\OT-06-0570-RG-FIN///
   'Call SELECHO("TRANSPOR")
   'NTRAN% = XVALO(VALACT1$("TRANSPOR"))
   'Text19 = TRIM$(FORMATNUM(NTRAN%, "I4"))
End Sub

Private Sub Command25_Click()
   If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
      '\\\OT-06-0570-RG-27/12/06///
      Call CommSel("!SUCENTRE", Command25, Text20)
      '\\\OT-06-0570-RG-FIN///
     'Call SELECHO("!SUCENTRE")
     'VDEVU$ = VALACT1$("!SUCENTRE")
     'Text20 = VDEVU$
   End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("FACTURA")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEC04.Text1.TEXT = Text1.TEXT
    RECUEC04.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call CONECRUN("FSETUP04/OPFACTUR", "Configuración de Funcionamiento")
    Call FINAL("")
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
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

Private Sub Command7_Click()
    Command7.Enabled = False
    Call CIERRARCH("*.*")
    MODX% = ALTERNA%("Listas de Precios\Condiciones de Pago\Alícuotas de I.V.A.")
    If MODX% = 1 Then
         Call CONECRUN("LISPRE04", "Listas de Precios de Venta")
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

Private Sub Command8_Click()
    '
    Dim RESP%
    Dim NUMEOCXX$
    '
    Command8.Enabled = False
    MODO% = COMALTER%("Opciones de Salida\\Pantalla\Impresora\Cancelar")
    If MODO% < 1 Or MODO% > 2 Then GoTo 99
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
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
      .Text9 = RASOCLI$(NCMIN%)
      .Text8 = TRIM$(Str$(NCMAX%))
      .Text7 = RASOCLI$(NCMAX%)
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
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
    'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    'FILTRO CLIENTE
    If MODO% <> 1 Then
      SQLX$ = SQLX$ + " AND Nume_Clie >= " & CLINFE & ""
      SQLX$ = SQLX$ + "AND Nume_Clie<= " & CLISUP & ""
    End If
    '
    Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
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
         NCLIE% = !Nume_Clie
         Call REPLA(RECONSI$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECONSI$, MKI$(CVINT(!FechMov)), 3, 2)
         Call REPLA(RECONSI$, "", 5, 18)
         Call REPLA(RECONSI$, !DoPriCor, 23, 12)
         Call REPLA(RECONSI$, MKI$(!Cod_Inte), 35, 2)
         Call REPLA(RECONSI$, MKD$(!CantSalid), 37, 8)
         Call REPLA(RECONSI$, MKD$(!CantIngre), 45, 8)
         Call REPLA(RECONSI$, MKD$(!CantSaldo), 53, 8)
         Call REPLA(RECONSI$, !ReferCor, 61, 48)
         Call REPLA(RECONSI$, String$(12, 0), 109, 12)
         Call REPLA(RECONSI$, MKS$(XVALO(Mid(!DoPriCor, 5))), 109, 4)
         Call REPLA(RECONSI$, MKI$(CVINT(!FeReMovi)), 122, 2)
          HOR = Format(!FeReMovi, "HH:MM:SS")
         Call REPLA(RECONSI$, MKI$(CVINT(HOR)), 124, 2)
         Call REPLA(RECONSI$, MKI$(!NumUsuar), 127, 2)
         '
         Call REPLA(RECONSJ$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECONSJ$, MKI$(!Cod_Inte), 3, 2)
         Call REPLA(RECONSJ$, !DoPriCor, 5, 12)
         Call REPLA(RECONSJ$, !DoPriLar, 17, 18)
         Call REPLA(RECONSJ$, MKI$(CVINT(!FechMov)), 35, 2)
         If IsNull(!NUPEDCLI) Then
            Call REPLA(RECONSJ$, MKS$(0), 37, 4)
         Else
            Call REPLA(RECONSJ$, MKS$(!NUPEDCLI), 37, 4)
         End If
         If IsNull(!NUPEDCLI) Then
             NUMEOCXX$ = ""
           Else
             NUMEOCXX$ = NUOCOCLI$(!NUPEDCLI, !Cod_Inte)
         End If
         Call REPLA(RECONSJ$, NUMEOCXX$, 41, 24)
         Call REPLA(RECONSJ$, MKD$(!CantSalid), 65, 8)
         Call REPLA(RECONSJ$, MKD$(!CantIngre), 73, 8)
         Call REPLA(RECONSJ$, MKD$(!CantSaldo), 81, 8)
         Call REPLA(RECONSJ$, MKD$(!ValoUnit), 89, 8)
           TOTI# = !ValoUnit * (!CantSalid - !CantIngre)
         Call REPLA(RECONSJ$, MKD$(TOTI#), 97, 8)
         '
         J& = J& + 1
         Call GRAREG("CONSICLI", RECONSI$, J&)
         Call GRAREG("REPENFAC", RECONSJ$, J&)
      .MoveNext
      Loop
    End With
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
                Call HEADERS("REPENFAC", "Cliente: " + RASOCLI$(NCXX%))
                Call CONECRUN("*LREPENFB", "Remitos Sin Facturar (Clientes)")
              Else
                Call CONECRUN("*LREPENFA", "Remitos Sin Facturar (Clientes)")
            End If
        End If
    End If
    '
99  Command8.Enabled = True
End Sub
'
Private Sub Eco21_Change(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub
'
Private Sub Form_Load()
    HASAR1.Puerto = 1
    HASAR1.Modelo = MODELO_615                              '//Este OCX no es 100% para 715F
    HASAR1.Comenzar
    HASAR1.PrecioBase = False
    HASAR1.TratarDeCancelarTodo
End Sub



'
Private Sub Label26_Click()
   Call SELECHO("INDICLIE")
   CPROX$ = TRIM$(VALACT1$("INDICLIE"))
   If CPROX$ <> "" Then
     NCXI% = NUMECLI%(CPROX$)
     If NCXI% > 0 Then
       If REGICLI&(NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
         Call CARDACLI
       End If
     End If
   End If
End Sub

Private Sub Label7_Click()
   If ULTREG&("!CARDEFAC") > 0 Then Exit Sub
   If Label7.Caption = "Factura:" Then
       TIFADOC% = 3: Label7.Caption = "N.Débito:"
       Label6.Caption = "(T.Libre)": MODODOC% = 1
     ElseIf Label7.Caption = "N.Débito:" Then
       TIFADOC% = 6: Label7.Caption = "N.Crédito:"
       Label6.Caption = "(Mat.Codif)": MODODOC% = 0
     Else
       TIFADOC% = 2: Label7.Caption = "Factura:"
       Label6.Caption = "(Mat.Codif)": MODODOC% = 0
   End If
   Call NUFACTURA
   '
   NCLIE% = XVALO(Text1.TEXT)
' If NCLIE% > 0 Then
'    If MODODOC% > 0 Then
'        PORDESCU.TEXT = ""
'      Else
'        PORDESCU.TEXT = DESCUCLI$(NCLIE%)
'    End If
'  End If
End Sub

Private Sub LENTREGA_Change(Index As Integer)
   If Index = 2 Then
     If XVALO(Text19) > 0 Then
       Call MENSERR(24, "Campo no Editable")
       On Error Resume Next
       Text19.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Text1_Change()
   Call BLANFORMU
End Sub

Private Sub Text12_Change()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text16_Change()
   Label16 = Left$(ECOARCH("ECUECON", Text16), 20)
End Sub

Private Sub Text19_Change()
  NTRA = XVALO(Text19)
  If NTRA > 0 Then
    If NTRA < 32767 Then
      NTRAI% = NTRA
      LENTREGA(2) = ECOARCH$("TRANSPOR", MKI$(NTRAI%))
    End If
  End If
End Sub

Private Sub Text2_Change()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text20_Change()
  NSUC& = XVALO(Text20)
  Label27 = ""
  If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
    If NSUC& > 0 Then
      If NSUC& <= ULTREG&("!SUCENTRE") Then
        XX$ = REGLEIDO$("!SUCENTRE", NSUC&)
        Label27 = TRIM$(Mid$(XX$, 3, 62))
        LENTREGA(0) = TRIM$(Mid$(XX$, 65, 48))
        LENTREGA(1) = TRIM$(Mid$(XX$, 113, 48))
      End If
    End If
  End If
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
          Else
            Eco21(Index).TEXT = ECOARCH$("LISTAS", MKI$(XVALO(TEXT21(Index).TEXT)))
            ARCHIPRE$ = TRIM$(TEXT21(0).TEXT)
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        End If
        If MONELISTA%(XVALO(TEXT21(Index))) = 2 Then
             MONEMIS = "Dólares"
             MONEMIS.ForeColor = RGB(255, 0, 0)
             MONEMI% = 2
             TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
             MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
             TICAMBIS.Visible = True
             Label24.Visible = True
           Else
             MONEMIS = "Pesos"
'            MONEMIS.ForeColor = RGB(0, 0, 255)
             MONEMI% = 1
             TICAMBIS = "1.0000"
'            MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
'            TICAMBIS.Visible = False
'            Label24.Visible = False
        End If
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(XVALO(TEXT21(Index).TEXT)))
      ElseIf Index = 2 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(XVALO(TEXT21(Index).TEXT)))
    End If
    '
End Sub

Private Sub Text3_Change()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text4_Change()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text5_Change()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text7_Change()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text8_Change()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text9_Change()
   On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
End Sub
'
Sub CARCONSIPEN()
   '
   Static DES%, HAS%
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   '
   If PREFACTU$ = "PEDIDO" Or PREFACTU$ = "ORDESPA" Then
     DES% = 33: HAS% = 32767: GoTo 3
   End If
   '
   If DES% < 1 Then
2    FEX = HOY(HO$)
     HAS% = FEX
     HASS$ = FECHATEX$(FEX)
     DESS$ = "01" + Mid$(HASS$, 3)
     DES% = FECHANUM(DESS$)
   End If
   '
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   OBX$ = OBJPLA$("DESHASFECHA/REMITOS Emitidos DESDE - HASTA Fecha", VDEF$)
   If OBX$ = "" Then
     CONSICLI% = 0
     Exit Sub
   End If
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If DES% > HAS% Then
      GoTo 2
   End If
   FEX = DES%: PERIO$ = FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " al " + FECHATEX$(FEX)
   '
3  DES1# = CDbl(CVDAT(DES%))
     DESELEC# = DES1#
   HAS1# = CDbl(CVDAT(HAS%))
     HASELEC# = HAS1#
   '
   LPRE% = XVALO(TEXT21(0).TEXT)
   NCLIE% = XVALO(Text1.TEXT)
   ARCHIPRE$ = TRIM$(Str$(LPRE%))
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
   CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
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
   SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DES1# & " "
   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
   '
   Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   With CONSITMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       RECONSI$ = REGBLAN("CONSICLI")
       J& = 0
       Do While .EOF = False
         CODIMAT% = !Cod_Inte
         CAN = !CantSalid - !CantIngre
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
           If InStr(EPAC$, "GUIDI") > 0 Then GoTo 7 ' para forzar que traiga el precio del remito
           '
           If NPEDID& > 0 Then
             ' busca precio y moneda del pedido
             PRECI1# = PRECIOPED#(NPEDID&, CODIMAT%, MONEMA%)
             If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
             COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
             PRECIO = PRECI1# * COEFCON
             GoTo 8
           End If
           '
7          MONEMA% = !MoneVal
           If MONEMA% <= 1 Then
             MONEPE% = MONEMIPED%(NPEDID&)    ' RECUPERA MONEDA DE PEDIDO
             If MONEPE% > MONEMA% Then MONEMA% = MONEPE%
           End If
           If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
           '
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           PRECIO = !ValoUnit * COEFCON
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
Sub CARCONSIPED(NPED&)
   '
   Static DES%, HAS%
   '
   DES% = 0: HAS% = 32767
   '
3  DES1# = CDbl(CVDAT(DES%))
   HAS1# = CDbl(CVDAT(HAS%))
   '
   LPRE% = XVALO(TEXT21(0).TEXT)
   NCLIE% = XVALO(Text1.TEXT)
   ARCHIPRE$ = TRIM$(Str$(LPRE%))
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
   CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
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
   SQLX$ = SQLX$ + " AND (TRIM(DoSECLar) = '' "
   SQLX$ = SQLX$ + " OR DoSecLar LIKE 'P/F*' ) "
   SQLX$ = SQLX$ + " AND NuPedCli = " & NPED& & " "
   '
   On Error GoTo 0
   Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   With CONSITMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       RECONSI$ = REGBLAN("CONSICLI")
       J& = 0
       Do While .EOF = False
         CODIMAT% = !Cod_Inte
         CAN = !CantSalid - !CantIngre
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
           PRECIO = !ValoUnit * COEFCON
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
Sub RECARPED(NCLIE%, CIPD%, FEREMIT%, NUMEPED&, PREPED, TITUCON$)
   '
   FEREMIX = CDbl(CVDAT(FEREMIT%))
   Call APERBASDAT("PEDCLI")
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PEDDETA "
   SQLX$ = SQLX$ + " WHERE NroClie = " & NCLIE% & " "
   SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
   SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
   SQLX$ = SQLX$ + "AND CodiInt = " & CIPD% & " "
   'SQLX$ = SQLX$ + "AND CDBL(INT(FechEmis)) <= " & FEREMIX & " "
   SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & FEREMIX & ")) "
   SQLX$ = SQLX$ + "ORDER BY NroPed ASC"
   '
   Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
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
         NUMEPED& = !NroPed
         PREPED = !PreUnid
         ENCUPE% = 1
       ElseIf CAREGIS& > 1 Then
         .MoveFirst
         JJ& = 0
         Do While (.EOF = False)
           TTXX$ = MKS$(!NroPed)
           TTXX$ = TTXX$ + AJUSTI$(!NroOcom, 16) + "  "
              FEX = CVINT(!FECHEMIS)
           TTXX$ = TTXX$ + FECHATEX$(FEX) + "  "
              FEX = CVINT(!FESOLENT)
           TTXX$ = TTXX$ + FECHATEX$(FEX) + "  "
           TTXX$ = TTXX$ + FORMATNUM$(!PreUnid, "F12.4")
              CAPEN = !CanPed - !CantEnt
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
    '
End Sub
'
Sub CARDICAMBIO()
   '
   DOCUDIF = ""
   MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
   If MODOMONE$ <> "BIMONETA" Then
      If MODOMONE$ <> "MULTIMON" Then
         Exit Sub
      End If
   End If
   '
   Dim DOCABX$(2048), SALPES#(2048), SALDOL#(2048), REDOCA&(2048)
   '
   NC% = XVALO(Text1)
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
'   '
'   Screen.MousePointer = 1
'   If ULTREG&("SELDOCAP") < 1 Then
'     Exit Sub
'   End If
'   '
'   Call SELECHO("SELDOCAP")
'   DOCABY$ = Left$(VALACT1$("SELDOCAP"), 18)
'   If TRIM$(DOCABY$) = "" Then Exit Sub
'   DOCUDIF = DOCABY$
'   BBXX$ = VALACT2$("SELDOCAP")
'   '
'   MONEMIS = "Pesos"
'   IDIFE = XVALO(Mid$(BBXX$, 12, 11))
'   If TIFADOC% = 3 Then
'     IDIFE = XVALO(Mid$(BBXX$, 28, 12)) * TICAMONE(2)
'   End If
'   If PIVADOC% = 1 Then
'     TASA1 = XVALO(PORIVA(0).Caption)
'     IDIFE = IDIFE / (1 + TASA1 / 100)
'   End If
'   '
'   ZZ$ = REGBLAN$("CARDEFAC")
'   Call REPLA(ZZ$, MKS$(1), 1, 4)
'   Call REPLA(ZZ$, "Por Diferencia de Cambio", 7, 64)
'   Call REPLA(ZZ$, MKS$(IDIFE), 429, 4)
'   Call REPLA(ZZ$, MKD$(CDbl(IDIFE)), 433, 8)
'   Call REGAPP("!CARDEFAC", ZZ$)
'   ZZ$ = REGBLAN$("CARDEFAC")
'   Call REPLA(ZZ$, "s/Factura " + DOCABY$, 7, 64)
'   Call REGAPP("!CARDEFAC", ZZ$)
'   Call CIERRARCH("!CARDEFAC")
'   '
End Sub


Sub CARGAPUERTOS()
    '
    Printer.TrackDefault = False
    PORTUFAC$ = TRIM$(CONTROL$(IDENTER$, "PORTFAC"))
      If PORTUFAC$ = "" Then PORTUFAC$ = CONTROL$("*", "PORTFAC")
      If PORTUFAC$ = "" Then GoTo 20
      NEWPORT$ = PORTUFAC$    ' + "/NOMESS"
      Call SETPRINPORT(NEWPORT$)
      If NEWPORT$ = "" Then GoTo 20
      Call SETPRINPORT("RESTORE")
    PORTUREM$ = TRIM$(CONTROL$(IDENTER$, "PORTREM"))
      If PORTUREM$ = "" Then PORTUREM$ = CONTROL$("*", "PORTREM")
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
   '
Function VALICLIE%()
   VCC% = 0
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE% = XVALO(Text1.TEXT)
       If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function

Sub BLANFORMU()
'   TASACLICK = 0
'   Text2.TEXT = ""
'   Text3.TEXT = ""
'   Text4.TEXT = ""
'   Text5.TEXT = ""
'   Text7.TEXT = ""
'   Text8.TEXT = ""
'   Text10.TEXT = ""
'   Text12.TEXT = ""
'   Text14.TEXT = ""
'   Text15 = ""
'   Text16 = ""
'   Label16 = ""
'   Label26 = ""
'   Text17 = "" 'LIMPIA RET-IVA
'   IEMBALA = ""
'   Text19 = ""
'   Text20 = ""
'   Label25 = ""
'   Ltota(10) = "Importe Total Neto"
'   '
'   If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
'     Call SETULTREG("!SUCENTRE", 0)
'     Call FRESHECHO("!SUCENTRE")
'   End If
'   '
'   '\\\OT-5-579-FG-16/08/05
'   Ltota(3).Caption = "Total Documento:"
'   Ltota(3).Left = 1260
'   Text18.Visible = False
'   '\\\OT-5-579-FG-FIN
'   '
'   ' PONE LETRA POR OMISIÓN
'   Label13.Caption = "A"
'   '\\\OT-5-506-FG-02/09/05
'   ' PONE PREFIJO POR OMISION AL 0001
'   UPREFI$ = TRIM$(Control$(IDENTER$, "PREFIFA"))
'   If UPREFI$ = "" Then UPREFI$ = "0001"
'   Label12.Caption = UPREFI$
'   '\\\OT-5-506-FG-FIN
'   '
'   Text6.TEXT = ""
'   List1.Clear
'   DOCUDIF = ""
'   '
'   Picture5.Visible = False
'   List3.Clear
'   '
'   For KKI% = 0 To 2
'      TEXT21(KKI%).TEXT = ""
'      Eco21(KKI%) = ""
'      BASIVA(KKI%) = ""
'      LENTREGA(KKI%) = ""
'   Next KKI%
'   '
'   PORDESCU.TEXT = ""
'   For KK% = 0 To 6
'      Total(KK%).TEXT = ""
'   Next KK%
'   TLABEL28 = ""
'   '
'   MARCONOTA.Visible = False
'   HOII% = HOY(HOS$)
'   Text9.TEXT = HOS$
'   Call SETULTREG("!CARDEFAC", 0)
'   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
'     On Error Resume Next
'     Kill LUCOM$ + "1CARDEFA.WKF"
'     On Error GoTo 0
'   End If
'   'Call NUFACTURA
'   '
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    '
    NC% = XVALO(Text1.TEXT): Text1.Refresh
    If NC% > 0 Then
      If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
        Text2.TEXT = RASOCLI$(NC%): Text2.Refresh
        Text3.TEXT = DOMICLI$(NC%): Text3.Refresh
        Text4.TEXT = CPOSCLI$(NC%): Text4.Refresh
        Text5.TEXT = LOCACLI$(NC%): Text5.Refresh
        Label26 = CODICLI$(NC%)
        Text7.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 10, 2, 2)
          Text7.Refresh
        Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CUITCLI$(NC%)
          Text12.Refresh
        '
        '\\\OT-5-579-FG-16/08/05
        YY$ = TRIM$(POSIVAC$(NC%))
        If YY$ = "COM.EXT." Then
          Ltota(3).Caption = "Total:"
          Ltota(3).Left = 250
          Text18 = "FOB"
          Text18.Visible = True
        End If
        '\\\OT-5-579-FG-FIN
        '
        CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
        Text8.ForeColor = RGB(0, 0, 0)
        If CREDISPO# < 0 Then Text8.ForeColor = RGB(255, 0, 0)
        Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 10, 2, 2)
          Text8.Refresh
        '
        TEXT21(0).TEXT = Str$(LIPRCLI%(NC%))
        TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
        TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
        '
        If XVALO(TEXT21(2).TEXT) < 1 Then
           Eco21(2).TEXT = "GERENCIA"
        End If
        '
        NCLIE% = XVALO(Text1.TEXT)
        If NCLIE% > 0 Then
          If MODODOC% > 0 Then
'             PORDESCU.TEXT = ""
            Else
'             PORDESCU.TEXT = DESCUCLI$(NCLIE%)
          End If
        End If
        PIVADOC% = PIVACLI%(NCLIE%)
        Call CARLETDOC: 'Call NUFACTURA '\\\OT-06-0051-WAR-15/02/06///
        Refresh
        '
        PPVI% = PIVADOC%
        If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
        Text16 = CONTROL$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
        '
        '\\\OT-5-506-FG-02/09/05
        ' PARA QUE TOME EN CUENTA EL PREFIJO DE LAS FACTURAS
        ' DE EXPORTACION
        If PPVI% = 5 Or PPVI% = 7 Then
          Label12.Caption = PREFEXPO$()
        End If
        '\\\OT-5-506-FG-FIN
        Call NUFACTURA '\\\OT-06-0051-WAR-15/02/06///
        '
        If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
          '\\\OT-7-010-RG-12/01/07
          Call FILTERFILE("SUCENTRE", "", 5, MKI$(NCLIE%))
          '\\\OT-7-010-RG-FIN
        End If
        '
        For KKI% = 0 To 2
          LENTREGA(KKI%) = LUENTRECLI$(NCLIE%, KKI%)
        Next KKI%
        '
        NTRAN% = NUTRANCLI%(NCLIE%)
        If NTRAN% > 0 Then
          Text19 = TRIM$(Str$(NTRAN%))
        End If
        '
        Label32 = "0"      ' significa que los datos son los de la base de datos de clientes
        '
        Text6 = ANOTACLI$(NC%)
'       MARCONOTA.Visible = True
'       MARCONOTA.Refresh
        '
        If VERLICRE% = 1 Then
          LMCR# = LIMICRE(NC%)
          SDEU# = SALDCLI#(NC%)
          CHPA# = CHEPEACRE#(NC%)
          If SDEU# + CHPA# >= LMCR# Then
            Call COMUNI("Imposible Facturar a este Cliente.\  \Límite de Crédito Asignado ya está Superado.\  \Informe al Supervisor de Ventas.")
            Call BLANFORMU
            On Error Resume Next
            Text1.SetFocus
            On Error GoTo 0
          End If
        End If
        '
        If POREMBAL(NCLIE%) > 0 Then
          CUEMBA$ = CONTROL$("SVD451S", "CUEVE-12")
          If CUEMBA$ = "" Or CUECOINT%(CUEMBA$) < 1 Then
            Call MENSERR(24, "Falta Definir Cuenta Contable\Correspondiente a Embalajes.\   \FLEXOFT Utilizará la Cuenta de Ventas Netas.")
          End If
        End If
        '
      End If
    End If
    Screen.MousePointer = 1
    '
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
        '
        If LETREXPO$ = "A" Then
          VARIDOC% = 1
          Label13.Caption = "A"
        End If
    End If
    '
End Sub
   '
Sub NUFACTURA()
   '
   Static FLINUMERA%
   Static NUMECORR%
   Static MONUMERA$
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
      QTRX$ = CONTROL$("NUMEFACT", "NUMECORR")
      If QTRX$ = "INDEPEN" Then
           NUMECORR% = (-1)
         ElseIf QTRX$ = "UNIFICA" Then
           NUMECORR% = (-2)
      End If
   End If
   '
   If MONUMERA$ = "" Then
      MONUMERA$ = "AUTO"
      If TRIM$(CONTROL$("", "NUMFACTU")) = "MANUAL" Then
         MONUMERA$ = "MANUAL"
      End If
   End If
   '
   OPMOVIX$ = "FC"
   If TIFADOC% = 3 Then OPMOVIX$ = "ND"
   If TIFADOC% = 6 Then OPMOVIX$ = "NC"
   VAMOVIX$ = Label13
   '
   DATEDOC% = FECHANUM(Text9.TEXT)
   If DATEDOC% < FLINUMERA% Then
      Call MENSERR(24, "Imposible Facturar en la Fecha Elegida.\  \La Fecha de Factura Debe Ser POSTERIOR\a la Fecha de Nueva Numeración.")
      Call CIERRARCH("*.*")
      Call FINAL("")
   End If
   FLINUME = CDbl(CVDAT(FLINUMERA%))
   '
   FEXUL = DATEDOC%
   FLIDIAS = CDbl(CVDAT(DIASANTES(FEXUL, 15)))
   '//////////*******NOTA*******////////
   'Con posterioridad se detecto un Error en la Impòrtacion
   '(VARIMOVIM(A,B), SE IMPORTO EN VEZ DE LA LETRA CON EL OPCIMOVIX$(ND,NC).
   'Por tal motivo el sistema no toma en cuenta los numero
   'de NC-ND inferiores al 01/11/04.
   'Pero si lo hace a futuro. es decir yo puedo reiniciar numeración
   'al 01/11/04
   
   FLIMI = FLIDIAS: MODORAPIDO% = 1
   If FLINUME > FLIMI Then
      FLIMI = FLINUME
      MODORAPIDO% = 0
   End If
   '
22 Call APERBASDAT("CABAN")
   ASELE$ = "SELECT NumeComp FROM CajaBanc "
   '
   If OPMOVIX$ = "ND" Or OPMOVIX$ = "NC" Then
     If NUMECORR% = (-2) Then    ' Unificada debitos y creditos
       ASELE$ = ASELE$ + "WHERE ((TipoMovim = 'FVEN' AND (OpciMovim = 'ND' OR OpciMovim = 'NC')) "
       ASELE$ = ASELE$ + "OR (TipoMovim = 'FCOM' AND (OpciMovim = 'DP' OR OpciMovim = 'CP'))) "
       GoTo 25
     End If
   End If
   '
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
   'ASELE$ = ASELE$ + "AND (CDBL(FechEmisi) >= " & FLINUME & ") "
   ASELE$ = ASELE$ + "AND FechEmisi >= " & FLIMI & " "
   ASELE$ = ASELE$ & "ORDER BY NumeComp DESC;"
   '
   Set NUFAXI = CueCorri.OpenRecordset(ASELE$, dbOpenDynaset, dbReadOnly)
   '
   NUMAXI& = 0
   On Error Resume Next
   NUFAXI.MoveFirst
   If Err < 1 Then
       On Error GoTo 0
       NUMAXI& = NUFAXI!NumeComp
     Else
       On Error GoTo 0
       If MODORAPIDO% = 1 Then
          FLIMI = FLINUME
          MODORAPIDO% = 0
          GoTo 22
       End If
   End If
   '
99 If Visible = True Then
     If NUMAXI& < 1 Then
       NUMAXI& = XVALO(InputBox$("Ingrese Número del Próximo Comprobante del Tipo Elegido", "Incicialización de Talonarios")) - 1
       If NUMAXI& < 0 Then GoTo 99
     End If
   End If
   NPP& = NUMAXI& + 1
   Text11.TEXT = TRIM$(Str$(NPP&))
   Text11.Refresh
   '
   If TIFADOC% = 2 Then
'     If TRIM$(Text10.TEXT) = "" Then
        Call NUREMIFA(NROREMIT&)
'       If NROREMIT& > XVALO(Text10.TEXT) Then
'         Text10.TEXT = TRIM$(Str$(NROREMIT&))
'       End If
'     End If
      Screen.MousePointer = 1
      Exit Sub
'     Else
'      Text10.TEXT = ""
   End If
   '
   Screen.MousePointer = 1
   '
End Sub
'
Function VERINUFA%(NFCY$)
   '
   Static FLINUMERA%
   Static NUMECORR%
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
      QTRX$ = CONTROL$("NUMEFACT", "NUMECORR")
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
   Set NUFAXI = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   '
   On Error Resume Next
   NUFAXI.MoveFirst
   If Err Then
       VERINUFA% = 1
     Else
       Do While Not NUFAXI.EOF
         If CDbl(NUFAXI.FechEmisi) >= FLIMI Then
            VERINUFA% = (-1)
            GoTo 90
         End If
       NUFAXI.MoveNext
       Loop
       VERINUFA% = 1
   End If
90 On Error GoTo 0
   '
End Function
'
Function VERINURE%()
   '
   Static PREFIREM$, FLIM%, CMO$
   '
   If CMO$ = "" Then
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If CMO$ = "" Then CMO$ = "RT"
   End If
   '
   If PREFIREM$ = "" Then
     PREFIREM$ = CONTROL$("", "PREFIREM")
     If PREFIREM$ = "" Then PREFIREM$ = "0001"
   End If
   '
   If FLIM% < 1 Then
     FENUE$ = CONTROL$("", "FENUREFA")
     If FENUE$ = "" Then FENUE$ = "18/04/03"
     FLIM% = FECHANUM(FENUE$)
     If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
   End If
   FLINUME = CDbl(CVDAT(FLIM%))
   '
  ' REMICOR$ = "RT." + TRIM$(Text10)
   REMILAR$ = Text10: While Len(REMILAR$) < 8: REMILAR$ = "0" + REMILAR$: Wend
   REMILAR$ = CMO$ + "-X-" + PREFIREM$ + "-" + REMILAR$
   '
   Call APERBASDAT("STOCKS")
   SQLX$ = "SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + " WHERE FechMov >= " & FLINUME & " "
   SQLX$ = SQLX$ + " AND CodiMovi = '" & CMO$ & "' "
   SQLX$ = SQLX$ + " AND DoPriLar = '" & REMILAR$ & "' "
   '
   Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
          VERINURE% = 1
        Else
          VERINURE% = (-1)
      End If
   End With
   On Error GoTo 0
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
      QTRX$ = CONTROL$("NUMEFACT", "NUMECORR")
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

Sub PRIFACTURA()
    '
    Static AGREDESP%
    '
    TICAMBX# = XVALO(TICAMBIS)
    '
    ' ESTO ES PARA ACELERAR BUSQUEDA DEL NUMERO DE SERIE
    CMO$ = TRIM$(CONTROL$("", "CODREFAC")):   If Len(CMO$) <> 2 Then CMO$ = "RT"
    NC1% = XVALO(Text1.TEXT)
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NC1% & " "
    SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DESELEC# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HASELEC# & " "
    Set CONSITEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    CAHOJAS% = 0
    FORIPRE$ = FORELFAC$            ' BUSCA POR LA FUNCION ASOCIADA AL TIPO DE COMPROBANTE
    If FORIPRE$ = "NULL" Then Exit Sub
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
'    If InStr(UCase$(TIPODOC$), "FACTURA") > 0 Then '\\\OT-06-0110-WAR-29/03/06///
'      If InStr(UCase$(TIPODOC$), "CONTADO") < 1 Then
'        If CPAG% < 1 Or InStr(UCase$(DECONPA$(CPAG%)), "CONTADO") > 0 Then
'          TIPODOC$ = TIPODOC$ + " CONTADO"
'        End If
'      End If
'    End If
    '
    DESTIDOC% = XVALO(Text1.TEXT)
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
    CODTABU1$(27) = "N-AUTENT"
    '
    CACOLTAB1% = 27
    '
    CAITAB1% = 0
    CAITCRU% = 0
    '
   ' SUMATOTA# = XVALO(Total(0).TEXT)
    'IDESCUEN# = XVALO(Total(1).TEXT)
    '
    COPIASREM% = XVALO(FINFACTURA.Text2(1))
    If COPIASREM% < 1 Then Text10.TEXT = ""
    '
    NroRemi$ = TRIM$(Text10.TEXT)
    If NroRemi$ <> "" Then
      While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
      PREFIREM$ = CONTROL$("", "PREFIREM")
      If PREFIREM$ = "" Then PREFIREM$ = "0001"
      NroRemi$ = PREFIREM$ + "-" + NroRemi$
    End If
    '
    If CONSICLI% = 1 Then
      NroRemi$ = NROREMICO$
    End If
    '
    COEFIVAINCLU# = 1
'    If XVALO(Total(2).TEXT) > 0.005 Then
'      COEFIVAINCLU# = (XVALO(Total(2).TEXT) + XVALO(Total(3).TEXT) + XVALO(Total(4).TEXT) + XVALO(Total(5).TEXT)) / XVALO(Total(2).TEXT)
'    End If
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
   ' CODPARAM$(8) = "POR-IVA": DOCPARAM$(8) = PORIVA(0).Caption
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = ""
   ' CODPARAM$(10) = "POR-IVA1": DOCPARAM$(10) = PORIVA(1).Caption
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
  '  CODPARAM$(29) = "TIPO-VAL": DOCPARAM$(29) = TRIM$(Text18)          ' VALOR FOB
    CODPARAM$(30) = "TICAMBIO": DOCPARAM$(30) = TICAMBIS               ' TIPO DE CAMBIO
    CODPARAM$(31) = "CTAC-CLI": DOCPARAM$(31) = CONTACTCLI$(DESTIDOC%) '  CONTACTO
    '\\\OT-5-579-FG-FIN
    '\\\OT-06-0125-WAR-03/04/06///
    CODPARAM$(33) = "NUBULTOS": DOCPARAM$(33) = TRIM$(FINFACTURA.Text2(2))          ' BULTOS
    CODPARAM$(34) = "PES-CONJ": DOCPARAM$(34) = TRIM$(FINFACTURA.Text2(11)) + " Kg" ' PESO NETO
    CODPARAM$(35) = "PES-COMP": DOCPARAM$(35) = TRIM$(FINFACTURA.Text2(4)) + " Kg"  ' PESO BRUTO
    CODPARAM$(36) = "BANC-CHE": DOCPARAM$(36) = TRIM$(FINFACTURA.Text2(5))          'BANCO EXTERIOR
    CODPARAM$(37) = "DOCU-APL": DOCPARAM$(37) = TRIM$(FINFACTURA.Text2(8))          ' BANCO LOCAL
    CODPARAM$(38) = "NUME-CHE": DOCPARAM$(38) = TRIM$(FINFACTURA.Text2(10))         ' CREDITO
    CODPARAM$(39) = "DEN-CUEN": DOCPARAM$(39) = TRIM$(FINFACTURA.Text2(9))          ' GUIA DE IMPORTACION
    CODPARAM$(40) = "REF-ASIE": DOCPARAM$(40) = TRIM$(FINFACTURA.Text2(6))          ' PUERTO DE EMBARQUE
    CODPARAM$(41) = "DET-PASE": DOCPARAM$(41) = TRIM$(FINFACTURA.Text2(7))          ' DESTINO
   ' CODPARAM$(42) = "MEDIDA-1": DOCPARAM$(42) = TRIM$(IEMBALA)                      ' gasto de embalaje
    CODPARAM$(43) = "ZONA-VEN": DOCPARAM$(43) = TRIM$(Str$(GRUCOCLI%(NC1%)))
    '\\\OT-06-0125-WAR-FIN///
    'AGREGUEMH04-10
    CODPARAM$(44) = "NRO-PASE": DOCPARAM$(44) = TRIM$(Label27.Caption) 'Sucursal
    '
    CODPARAM$(45) = "ITO-GRAV": DOCPARAM$(45) = ""
    CODPARAM$(46) = "ITO-NOGR": DOCPARAM$(46) = ""
    CODPARAM$(47) = "ITO-GRA1": DOCPARAM$(47) = ""
    CODPARAM$(48) = "ITO-GRA2": DOCPARAM$(48) = ""
    '
    CODPARAM$(49) = "IMP-DEBE": DOCPARAM$(49) = FINFACTURA.Text2(13)  ' FLETE
    CODPARAM$(50) = "IMP-HABE": DOCPARAM$(50) = FINFACTURA.Text2(12)  ' SEGURO
    CODPARAM$(51) = "TOT-HABE": DOCPARAM$(51) = FORMATNUM(TOTATOTA#, "F13.2")
    '
    CAPARDOC% = 51 '32 '\\\OT-06-0125-WAR-03/04/06///
    '
'MsgBox "PRIFACTURA-3    "
    Call SETULTREG("!OBSERVOF", 0)
    If FINFACTURA.Text4 <> "" Then
      JR& = 0
      FORMOBSE$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
      ANCHOB% = XVALO(Mid$(FORMOBSE$, 2)): If ANCHOB% < 32 Then ANCHOB% = 32
      Call FRATEXTO(FINFACTURA.Text4, ANCHOB%)
      For KKII% = 1 To CARETEX%
        JR& = JR& + 1
        Call GRAREG("!OBSERVOF", RETEX$(KKII%), JR&)
      Next KKII%
      Call SETULTREG("!OBSERVOF", JR&)
      Call CIERRARCH("!OBSERVOF")
    End If
    POVERTI% = XVALO(TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/POSINVER"))))
    '
    PRINTERPORT$ = PORTUFAC$
    CACOPDOC% = XVALO(FINFACTURA.Text2(0).TEXT)
    '
    NOCONFIRM% = 0
    If TRIM$(UCase$(CONTROL$("FACTUREM", "IMPRENUM"))) <> "AUTONUME" Then
      If CONTROL$("FACTUREM", "FORMULAR") <> "HOJASUEL" Then
        NOCONFIRM% = 1
      End If
    End If
    '
    HOJASUEL% = 0
    If CONTROL$("FACTUREM", "FORMULAR") = "HOJASUEL" Then
       HOJASUEL% = 1
    End If
    '
'MsgBox "PRIFACTURA-4    "
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
      If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
        REPLAORI% = 1
      End If
     'va fuera del for para q no se ejecute innecesariamente
  '\\\OT-06-0339-MH-FIN///
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
        CODDX$ = CODORIG$(CIPRE%, NC1%) ' + SUBCO$
      ' se pasa el codigo interno, mas el nº cliente(NC1%) y trae el codigo externo solicitado
      End If
      '\\\OT-06-0339-MH-FIN///
      '
      TETABU1$(1, CAITAB1%) = CODDX$
      TETABU1$(2, CAITAB1%) = Mid$(XXX$, 7, 64)
      TETABU1$(3, CAITAB1%) = CAPEX$
      'TETABU1$(4, CAITAB1%) = PRELI$
      'TETABU1$(5, CAITAB1%) = DESCX$
      'TETABU1$(6, CAITAB1%) = PREUX$
      'TETABU1$(7, CAITAB1%) = ITOTX$
      'TETABU1$(8, CAITAB1%) = FENTX$
      TETABU1$(9, CAITAB1%) = Str$(U&)
      TETABU1$(10, CAITAB1%) = Mid$(XXX$, 71, 4)    ' unidad de medida
      TETABU1$(11, CAITAB1%) = Mid$(XXX$, 75, 64)   ' referencia 1
      TETABU1$(12, CAITAB1%) = Mid$(XXX$, 139, 64)  ' referencia 1
      TETABU1$(13, CAITAB1%) = Mid$(XXX$, 203, 64)  ' referencia 1
      TETABU1$(14, CAITAB1%) = Mid$(XXX$, 315, 24)  ' numero de plano
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
      TETABU1$(27, CAITAB1%) = Mid$(XXX$, 371, 16) ' numero de pedido cliente
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
'        call APERbasdat("STOCKS")
'        DOPRIC$ = TRIM$(Mid$(XXX$, 387, 16))
'        NORIT% = CVI(Mid$(XXX$, 503, 2))
'        '
'        SQLX$ = "DoPriCor = '" & DOPRIC$ & "' "
'        SQLX$ = SQLX$ + "AND Cod_Inte = " & CIPRE% & " "
'        SQLX$ = SQLX$ + "AND Nume_Clie = " & DESTIDOC% & " "
'        SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
'        '
'        With Movisto
'          .FindFirst SQLX$
'25        If .NoMatch = False Then
'            SERNU$ = !ReferCor
'            If Left$(SERNU$, 5) = "S.No." Then
'              CAITAB1% = CAITAB1% + 1
'              TETABU1$(2, CAITAB1%) = SERNU$
'            End If
'          End If
'        End With
        DOPRIC$ = TRIM$(Mid$(XXX$, 387, 16))
        NORIT% = CVI(Mid$(XXX$, 503, 2))
        '
        SQLX$ = "DoPriCor = '" & DOPRIC$ & "' "
        SQLX$ = SQLX$ + "AND Cod_Inte = " & CIPRE% & " "
        SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
        '
        With CONSITEMP
          .FindFirst SQLX$
25        If .NoMatch = False Then
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
        If CONTROL$("FACTURA", "AGREDESP") = "SI" Then AGREDESP% = 1
      End If
      If AGREDESP% = 1 Then
        DESPAI$ = TRIM$(ATRIB1$(CIPRE%))
        If DESPAI$ <> "" Then
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = "Despacho: " + DESPAI$
        End If
      End If
      '
      SUBTOTA# = SUBTOTA# + XVALO(TETABU1$(23, CAITAB1%))
      SUBTOTB# = SUBTOTB# + XVALO(TETABU1$(26, CAITAB1%))
      '
      If (U& Mod ITPORHOJA%) = 0 Then
        If U& < FIN& Then
          If CACOPDOC% > 0 Then
            DOCPARAM$(26) = TRIM$(CSTRING$(MKD$(SUBTOTA#), 4, 12, 2, 2))
            Call PRINTDOC("@" + FORIPRE$)
            Call CIERRARCH("*.*")
            Call BAJALDISCO
          End If
        End If
      End If
109 Next U&
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
'MsgBox "PRIFACTURA-5    "
    If POVERTI% < 1 Then
      CAITAB1% = CAITAB1% + 1
      For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
      For KU& = 1 To ULTREG&("!OBSERVOF")
        TTXX$ = REGLEIDO$("!OBSERVOF", KU&)
        CAITAB1% = CAITAB1% + 1
        For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
        TETABU1$(2, CAITAB1%) = RTrim$(TTXX$)
      Next KU&
      Call SETULTREG("!OBSERVOF", 0)
    End If
        
  ' CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = FORMATNUM$(XVALO(Total(0).TEXT) + XVALO(IEMBALA), "F13.2")
  ' CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = Total(3).TEXT
  ' CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = Total(4).TEXT
  ' CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = Total(1).TEXT
  ' CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = Total(2).TEXT
  ' CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = Total(6).TEXT
  ' CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = Total(5).TEXT
  '       VADOLA# = XVALO(Total(6).TEXT) * TICAMBX#
    DOCPARAM$(23) = TRIM$(CSTRING$(MKD$(VADOLA#), 4, 12, 2, 2))
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = Text15
    DOCPARAM$(26) = ""
    DOCPARAM$(27) = Text17                          'PARA IMPRIMIR PERC-IVA
   '   ITOGRAVA# = XVALO(Total(2)) - XVALO(TLABEL28)
    DOCPARAM$(45) = TRIM$(FORMATNUM$(ITOGRAVA#, "F13.2"))
      INOGRAVA# = XVALO(TLABEL28)
    DOCPARAM$(46) = TRIM$(FORMATNUM$(INOGRAVA#, "F13.2"))
  '  DOCPARAM$(47) = TRIM$(BASIVA(0))
  '  DOCPARAM$(48) = TRIM$(BASIVA(1))
  '    TOTATOTA = XVALO(Total(6)) + XVALO(FINFACTURA.Text2(12)) + XVALO(FINFACTURA.Text2(13))
    DOCPARAM$(51) = FORMATNUM$(TOTATOTA#, "F13.2")
    '
    If Label13.Caption = "B" Then
    '  DOCPARAM$(5) = FORMATNUM$(COEFIVAINCLU# * (XVALO(Total(0).TEXT) + XVALO(IEMBALA)), "F13.2")
    '  DOCPARAM$(12) = FORMATNUM$(COEFIVAINCLU# * XVALO(Total(1).TEXT), "F13.2")
'      DOCPARAM$(13) = FORMATNUM$(COEFIVAINCLU# * XVALO(Total(2).TEXT), "F13.2")
'      DOCPARAM$(42) = FORMATNUM$(COEFIVAINCLU# * XVALO(IEMBALA), "F13.2")          ' gasto de embalaje
'      DOCPARAM$(45) = TRIM$(FORMATNUM$(COEFIVAINCLU# * ITOGRAVA#, "F13.2"))
'      DOCPARAM$(46) = TRIM$(FORMATNUM$(COEFIVAINCLU# * INOGRAVA#, "F13.2"))
'      DOCPARAM$(47) = TRIM$(FORMATNUM$(COEFIVAINCLU# * XVALO(BASIVA(0)), "F13.2"))
'      DOCPARAM$(48) = TRIM$(FORMATNUM$(COEFIVAINCLU# * XVALO(BASIVA(1)), "F13.2"))
    End If
    '
    CODFOR$ = FORIPRE$
    LOIMPLET% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET"))
    If LOIMPLET% < 16 Then LOIMPLET% = 32
'    TTTT# = XVALO(Total(6).TEXT)
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
      Call PRINTDOC("@" + FORIPRE$)
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
'MsgBox "PRIFACTURA-6    "
    CACOPDOC% = XVALO(FINFACTURA.Text2(1).TEXT)
    If CACOPDOC% > 0 Then
      If TRIM$(Text10.TEXT) <> "" Then
        For KKI% = 0 To CACOLTAB1%
          For KKJ% = 0 To CAITAB1%
            TETABU1$(KKI%, KKJ%) = ""
          Next KKJ%
        Next KKI%
        '
        DOCUORIG$ = NUMEDOC$
        Call PRIREMIFA
      End If
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
'MsgBox "PRIFACTURA-7    "
    If NOCONFIRM% > 0 Then
       Call COMUNI("Se han Emitido los Documentos Solicitados")
    End If
    For KKI% = 0 To CACOLTAB1%
      For KKJ% = 0 To CAITAB1%
        TETABU1$(KKI%, KKJ%) = ""
      Next KKJ%
    Next KKI%
    '
End Sub

Sub PRIREMIFA()
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    If TRIM$(FORIPRE$) <> "" Then
       '
       '\\\OT-06-0114-WAR-31/03/06///
       Call SETULTREG("!OBSERVOF", 0)
       If FINFACTURA.Text5 <> "" Then
         JR& = 0
         CODFOR$ = FORIPRE$
         FORMOBSE$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
         ANCHOB% = XVALO(Mid$(FORMOBSE$, 2)): If ANCHOB% < 32 Then ANCHOB% = 32
         Call FRATEXTO(FINFACTURA.Text5, ANCHOB)
         For KKII% = 1 To CARETEX%
           JR& = JR& + 1
           Call GRAREG("!OBSERVOF", RETEX$(KKII%), JR&)
         Next KKII%
         Call SETULTREG("!OBSERVOF", JR&)
         Call CIERRARCH("!OBSERVOF")
       End If
       POVERTI% = XVALO(TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/POSINVER"))))
       '\\\OT-06-0114-WAR-31/03/06///
       '
       PREFIREM$ = CONTROL$("", "PREFIREM")
       If PREFIREM$ = "" Then PREFIREM$ = "0001"
       NroRemi$ = TRIM$(Text10.TEXT)
       While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
       NroRemi$ = PREFIREM$ + "-" + NroRemi$
       '
       NC% = XVALO(Text1.TEXT)
       FECHDOC$ = Text9.TEXT
       NUMEDOC$ = NroRemi$
       TIPODOC$ = "Remito"
       '
       If TRIM$(Text14.TEXT) <> "" Then
         NUOCOMPRA$ = TRIM$(Text14.TEXT)
       End If
       '
       CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(TRXP$, 1, 32))
       CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(FINFACTURA.Text2(2).TEXT)
       CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(TRXP$, 65, 16))
       CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(FINFACTURA.Text2(3).TEXT)
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
'       CODPARAM$(17) = "MEDIDA-1": DOCPARAM$(17) = TRIM$(IEMBALA)             ' gasto de embalaje
       CODPARAM$(18) = "ZONA-VEN": DOCPARAM$(18) = TRIM$(Str$(GRUCOCLI%(NC%)))
       CODPARAM$(19) = "TOTA-ITS": DOCPARAM$(19) = Label25
         NTRA% = XVALO(Text19)
       CODPARAM$(20) = "REF-ASIE": DOCPARAM$(20) = DOMITRANS$(NTRA%)
       CODPARAM$(21) = "DEN-CUEN": DOCPARAM$(21) = LOCATRANS$(NTRA%)
       CODPARAM$(22) = "MEDIDA-2": DOCPARAM$(22) = TELETRANS$(NTRA%)
       'AGREGUE 02-10MH
       CODPARAM$(23) = "NRO-PASE": DOCPARAM$(23) = TRIM$(Label27.Caption)
       '\\\OT-07-0007-RG-08/01/07///
       HORA$ = HORAENT$(NC%)
       CODPARAM$(24) = "Z-PLANO": DOCPARAM$(24) = TRIM$(HORA$)
          NUVENDEX$ = AJUSTI$(FORMATNUM$(XVALO(TEXT21(2)), "I3"), 4)
       CODPARAM$(25) = "VENDEDOR": DOCPARAM$(25) = NUVENDEX$ + Eco21(2).TEXT
       '\\\OT-07-0007-RG-FIN///
       '
       CAPARDOC% = 25
       '
       CODTABU1$(1) = "COD-MAT"
       CODTABU1$(2) = "DES-MAT"
       CODTABU1$(3) = "CANTIDAD"
       CODTABU1$(4) = "ORD-ITEM"
       CODTABU1$(5) = "N-AUTENT"
       CACOLTAB1% = 5
       '
       DESTIDOC% = NC%
       CAITAB1% = 0
       '
       If FINFACTURA.Text3 <> "" Then
         Call FRATEXTO(FINFACTURA.Text3, 40)
         For KKII% = 1 To CARETEX%
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = ""
           TETABU1$(2, CAITAB1%) = RETEX$(KKII%)
           TETABU1$(3, CAITAB1%) = ""
           TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
         Next KKII%
         FINFACTURA.Text3 = ""
         FINFACTURA.Frame5.Visible = False
         GoTo 80
       End If
       '
       ARCHILIB$ = "!1CARDEFA"
      
      '\\\OT-06-0339-MH-24/08/06///
      REPLAORI% = 0
      If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
        REPLAORI% = 1
      End If
      '\\\OT-06-0339-MH-FIN///
       
       For U& = 1 To ULTREG&("!CARDEFAC")
         XXX$ = REGLEIDO$("!CARDEFAC", U&)
         CANS$ = Str$(CVS(Left$(XXX$, 4))) ' cantidad
         NC1% = XVALO(Text1.TEXT) ' AGREGADO MH
         CIPRE% = CVI(Mid$(XXX$, 5, 2)) ' AGREGADO
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
             TETABU1$(5, CAITAB1%) = Mid$(XXX$, 371, 16) 'numero de pedido cliente
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
               If CONTROL$("FACTURA", "AGREDESP") = "SI" Then AGREDESP% = 1
             End If
             If AGREDESP% = 1 Then
               CIIY% = CODINT%(CDXXS$)
               DESPAI$ = TRIM$(ATRIB1$(CIIY%))
               If DESPAI$ <> "" Then
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(2, CAITAB1%) = "Despacho: " + DESPAI$
                 TETABU1$(3, CAITAB1%) = "" 'Para que no Imprima Q
               End If
             End If
             '\\\OT-06-0273-WAR-FIN///
             '
           End If
         'End If
       Next U&
       '
       If POVERTI% < 1 Then
         CAITAB1% = CAITAB1% + 1
         For KU& = 1 To ULTREG&("!OBSERVOF")
           TTXX$ = REGLEIDO$("!OBSERVOF", KU&)
           CAITAB1% = CAITAB1% + 1
           TETABU1$(2, CAITAB1%) = RTrim$(TTXX$)
         Next KU&
         Call SETULTREG("!OBSERVOF", 0)
       End If
       '
80     CACOPDOC% = XVALO(FINFACTURA.Text2(1).TEXT)
       PRINTERPORT$ = PORTUREM$
       Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
       '
    End If
    '
End Sub

Sub GRAFACTURA(OKX%)
   '
''MsgBox "Grabando en Factura_Nue"
'Call GRAFACTURA_NUE(OKX%)
'Exit Sub
'   Dim NIVA7#(10)
'   '
'   OKX% = 0
'   NCLIE% = XVALO(Text1.TEXT)
'   If MARCONOTA.Visible = True Then Exit Sub
'   If ULTREG&("!CARDEFAC") < 1 Then Exit Sub
'   '
'   Call APROCOPIAS(OKX%)
'   COPIASFAC% = XVALO(FINFACTURA.Text2(0))
'   COPIASREM% = XVALO(FINFACTURA.Text2(1))
'   If OKX% < 0 Then Exit Sub
'   If COPIASFAC% + COPIASREM% < 1 Then Exit Sub
'   '
''   If COPIASFAC% > 0 Then
''     If XVALO(Total(6).TEXT) < 0.01 Then
''       Call MENSERR(24, "Imposible Emitir Factura SIN CARGO !!!")
''       Exit Sub
''     End If
''   End If
'   '
'   If COPIASFAC% < 1 Then
'      Call GRAREMIFA
'      Call PRIFACTURA
'      Screen.MousePointer = 1
'      Exit Sub
'   End If
'   '
'   PPVI% = PIVACLI%(NCLIE%)
'   If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
'   CUEVE$ = Control$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
'   QQQ% = CUECOINT%(CUEVE$)
'   '
'   If TIFADOC% = 3 Or TIFADOC% = 6 Then
'      OPIMPUTA.Option1.Value = True
'      Call OPIMPUTA.Option1_Click
'      '
'      If TIFADOC% = 6 Then
'        'Option1(5).Value = True
'        If Control$("", "IMPUNCRE") = "DIFERENCIACAMBIO" Then
'          OPIMPUTA.Option2.Value = True
'          Call OPIMPUTA.Option1_Click
'        ElseIf Control$("", "IMPUNCRE") = "VARIOS" Then
'           OPIMPUTA.Option3.Value = True
'           Call OPIMPUTA.Option1_Click
'        End If
'      End If
'
'      If TRIM$(DOCUDIF) <> "" Then
'        OPIMPUTA.Option2.Value = True
'      End If
'      '
'      OPIMPUTA.Show 1
'   End If
'   '
''   'NUEVO PARA INGRESAR UN NUMERO DE FACTURA MANUAL
'   '
'   NUPANTALLA& = XVALO(Text11.TEXT)
'   If NFCAM% < 1 Then
'       Call NUFACTURA
'       If XVALO(Text11) <> NUPANTALLA& Then GoTo 16
'     Else
'       TICOMX$ = "FC"
'       If TIPO% = 3 Then TICOMX$ = "ND"
'       If TIPO% = 6 Then TICOMX$ = "NC"
'       NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
'       NFCY$ = TICOMX$ + "-" + Label12 + "-" + NFCZ$ + "-" + Label13
'       '
'       If VERINUFA%(NFCY$) < 0 Then
'16       Call MENSERR(24, "Número de Comprobante no Válido")
'         Exit Sub
'       End If
'   End If
'   '
'21 If XVALO(Text11.TEXT) > NUPANTALLA& Then
'      Call CIERRARCH("SVD202")
'      Call CIERRARCH("DEUDOR2")
'      Call CIERRARCH("FACTUR")
'      Call CIERRARCH("CAJABANC")
'      Call CIERRARCH("MOVISTO")
'      Call CIERRARCH("SALSTOCK")
'      Call COMUNI("FLEXOFT ha Re-Numerado el Documento\como Número '" + TRIM$(Text11.TEXT) + "'.\  \Verifique y Vuelva a Grabar.")
'      Exit Sub
'   End If
'   Text11.TEXT = TRIM$(Str$(NUPANTALLA&))
'   '
'   COPIASREM% = XVALO(FINFACTURA.Text2(1))
'   If COPIASREM% < 1 Then
'      Text10.TEXT = ""
'   End If
'   '
'   Screen.MousePointer = 11
'   If VALICLIE% < 1 Then
'      MERRO$ = "Número de Cliente no Válido"
'      GoTo 98
'   End If
'   NCLIE% = XVALO(Text1.TEXT)
'   '
'   If XVALO(TEXT21(0).TEXT) > 255 Then
'      MERRO$ = "Lista de Precios no Válida"
'      GoTo 98
'   End If
'   '
'   If XVALO(TEXT21(1).TEXT) > 255 Then
'      MERRO$ = "Condición de Pago no Válida"
'      GoTo 98
'   End If
'   '
'   If XVALO(TEXT21(2).TEXT) > 255 Then
'      MERRO$ = "Vendedor no Válido"
'      GoTo 98
'   End If
'   '
'   If XVALO(Total(6).TEXT) * TICAMBX# < 0.01 Then
'      MERRO$ = "Importe no Válido"
'      GoTo 98
'   End If
'   '
'   HOII% = HOY(HOS$): FEACT = HOII%
'   FEMI% = FECHANUM(Text9.TEXT)
'   ' LA SIGUIENTE VALIDACION SE SUPRIME A PEDIDO DE S.IGLESIAS DE I.GUIDI
'   'If FEMI% < DIANTE(FEACT) Or FEMI% > DIPOST(FEACT) Then
'   '   MERRO$ = "Fecha de Emisión no Válida"
'   '   GoTo 98
'   'End If
'   '
'   If VERLICRE% = 1 Then
'     If TIFADOC% < 4 Then ' EXCLUYE CREDITOS
'       '
'       For UKK& = 1 To ULTREG&("!CARDEFAC")
'         XKK$ = REGLEIDO$("!CARDEFAC", UKK&)
'         NROPEDXX& = XVALO(TRIM$(Mid$(XKK$, 371, 16)))
'         If NROPEDXX& < 1 Then GoTo 22
'         RFFKK$ = RECFILT$("AUTORENT", 1, MKS$(NROPEDXX&))
'         If Len(RFFKK$) < 4 Then GoTo 22 ' NO AUTORIZADO
'         REGAUTKK& = CVS(RFFKK$)
'         XAUT$ = REGLEIDO$("AUTORENT", REGAUTKK&)
'         If UCase$(Mid$(XAUT$, 31, 2)) <> "SI" Then GoTo 22
'       Next UKK&
'       GoTo 25
'       '
'22     LMCR# = LIMICRE(NCLIE%)
'       SDEU# = SALDCLI#(NCLIE%)
'       CHPA# = CHEPEACRE#(NCLIE%)
'       TPED# = XVALO(Total(6).TEXT) * TICAMBX#
'       If SDEU# + CHPA# + TPED# > LMCR# Then
'         TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2))
'         If TLIMI$ = "" Then TLIMI$ = "0.00"
'         Call COMUNI("Imposible Emitir Este Documento.\  \Supera el Límite de Crédito Asignado\para esta Cuenta de $ " + TLIMI$ + ".\  \Informe al Supervisor de Ventas.")
'         Exit Sub
'       End If
'     End If
'   End If
'   '
'25 DESC$ = TRIM$(PORDESCU.TEXT)
'   PODES = COEFDESCU(DESC$)
'   '
'   For KK% = 1 To 10: TOTAFAC#(1, KK%) = 0: Next KK%
'   TOTAFAC#(1, 1) = XVALO(BASIVA(0)) / (1 - PODES / 100)
'   TOTAFAC#(1, 2) = XVALO(BASIVA(1)) / (1 - PODES / 100)
'   TOTAFAC#(1, 3) = XVALO(Total(0).TEXT)
'   TOTAFAC#(1, 4) = XVALO(Total(1).TEXT)
'   TOTAFAC#(1, 5) = XVALO(BASIVA(0))
'   TOTAFAC#(1, 6) = XVALO(Total(3).TEXT)
'   TOTAFAC#(1, 7) = XVALO(Total(5).TEXT)
'   TOTAFAC#(1, 8) = XVALO(Total(6).TEXT)
'   TOTAFAC#(1, 9) = XVALO(BASIVA(1))
'   TOTAFAC#(1, 10) = XVALO(Total(4).TEXT)
'   '
'   TICA = 1
'   For KK% = 1 To 10
'     TOTAFAC#(1, KK%) = ROUND#(TOTAFAC#(1, KK%) * TICAMBX#, 2)
'     TOTAFAC#(2, KK%) = TOTAFAC#(1, KK%) / TICA
'   Next KK%
'   '
'   LU$ = LUDAT$
'   ARCHIDEU$ = "SVD202"
'   ARCHICLI$ = "DEUDOR"
'   ARCHIFAC$ = "FACTUR"
'   RD% = 1
'   TIPO% = TIFADOC%
'   VA% = VARIDOC%
'   NRO = XVALO(Text11.TEXT)
'   NC% = NCLIE%
'   FFI% = FECHANUM(Text9.TEXT)
'   CPAG% = XVALO(TEXT21(1).TEXT)
'   '
'   TOTAGEN# = 0
'   NENETO# = XVALO(Total(2)) - XVALO(IEMBALA)
'   NEFA7# = ROUND#(NENETO# * TICAMBX#, 2)
'   TOTAGEN# = TOTAGEN# + NEFA7#
'   NEMBAL7# = ROUND#(XVALO(IEMBALA) * TICAMBX#, 2)
'   TOTAGEN# = TOTAGEN# + NEMBAL7#
'   For KKI% = 0 To 2
'     NIVA7#(KKI%) = ROUND#(XVALO(Total(KKI% + 3)) * TICAMBX#, 2)
'     TOTAGEN# = TOTAGEN# + NIVA7#(KKI%)
'   Next KKI%
'   PERIB7# = ROUND#(XVALO(Text15) * TICAMBX#, 2)
'   TOTAGEN# = TOTAGEN# + PERIB7#
'   PERIVA7# = ROUND#(XVALO(Text17) * TICAMBX#, 2)
'   TOTAGEN# = TOTAGEN# + PERIVA7#
'   ' TOTAGEN# = XVALO(Total(6).TEXT) * TICAMBX#
'   '
'   'Call BLIMESS("Actualizando Saldo de Cuenta ...")
'   SIGX% = 1: If TIPO% = 6 Then SIGX% = (-1)
'   TTF# = TOTAGEN# * SIGX%
'   Call GRACLI(LU$, ARCHICLI$, NC%, TTF#, DATCLI$(), 2)
'   Call CIERRARCH("DEUDOR2")
'   '
'   FLIM% = FECHANUM("01/01/99")
'   If ULTREG&("FORMFAC") > 0 Then
'      FFS$ = REGLEIDO$("FORMFAC", 2&)
'      FLIM% = CVI(Mid$(FFS$, 19, 2))
'   End If
'   '
'   'Call BLIMESS("Grabando Factura ...")
'   ' LO QUE SIGUE ES PARA CONSIDERAR LAS HOJAS QUE SE CONSUMIRAN
'   'FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
'   FORIPRE$ = FORELFAC$
'   CARENGLO& = ULTREG&("!CARDEFAC")
'   ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
'   If ITPORHOJA% < 1 Then
'     ITPORHOJA% = 80
'   End If
'   CAHOJNEC% = Int(0.99 + CARENGLO& / ITPORHOJA%)
'   '
'   If NFCAM% < 1 Then
'     Call NUFACTURA
'   End If
'   '
'27 NPP& = XVALO(Text11) + CAHOJNEC% - 1
'   Text11.TEXT = TRIM$(Str$(NPP&))
'   Text11.Refresh
'   '
'   TICOMX$ = "FC": OPMOVIX$ = "FAC"
'   If TIPO% = 3 Then TICOMX$ = "ND": OPMOVIX$ = "DEB"
'   If TIPO% = 6 Then TICOMX$ = "NC": OPMOVIX$ = "CRE"
'   NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
'   NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + Label13
'   NFCY$ = TICOMX$ + "-" + Label12 + "-" + NFCZ$ + "-" + Label13
'   '
'   OPMOVIX$ = "FC"
'   If TIFADOC% = 3 Then OPMOVIX$ = "ND"
'   If TIFADOC% = 6 Then OPMOVIX$ = "NC"
'   VAMOVIX$ = Label13
'   '
'   SIGX% = 1: If TIPO% = 6 Then SIGX% = (-1)
'   '
'   'NUEVO PARA INGRESO MANUAL DE NUM FAC
'   If NFCAM% = 1 Then
'      If VERINUFA%(NFCY$) < 0 Then
'        OKX% = (-1)
'        Exit Sub
'      End If
'    End If
'   'HASTA ACA NUERO PARA INGRESO MANUAL DE NUM FAC
'   '
'' AQUI GRABAR CAJABANC PARA FACTURA
'   Call APERBASDAT("CABAN")
'   NOCLOSE% = 1
'30 CoComCor$ = NFCX$
'   CoComLar$ = NFCY$
'   CodiEmp% = 0
'   CodiSuc% = 0
'   NumeClie% = GRUDIVCLI%(NC%)
'   NumeProv% = 0
'   FechEmi = CVDAT(FFI%)
'   Referen$ = RASOCLI$(NC%)
'   FeContab = CVDAT(FFI%)
'   FeAcredi = CVDAT(FFI%)
'   CuentaCo% = CUECOCLI%(NC%)
'   TiMovi$ = "FVEN"
'   OpMovi$ = OPMOVIX$
'   VaMovi$ = Label13
'   Descrip$ = Label7 + " " + Label13.Caption + " " + Label6.Caption + " " + Label12.Caption + "-" + Text11.TEXT
'   NumeTal = XVALO(Label12)
'   NumeCom = CDbl(Text11)
'   StatInit = 0
'   ImpoDeb = TOTAGEN# * SIGX%
'   Observa$ = RASOCLI$(NC%)
'   Call GRACABAN
'   '
'   'graba venta neta - Verificar que QQQ sea la IC de combinacion Cliente-Tipo de Comprobante
'   TOTANETX# = NEFA7#
'   If TOTANETX# >= 0.005 Then
'     QQ1% = CUECOINT%(Text16)
'     If QQ1% < 1 Then QQ1% = QQQ%
'     CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
'     TiMovi$ = "VENTA"
'     OpMovi$ = OPMOVIX$
'     VaMovi$ = Label13
'     Descrip$ = "Mercaderia Vendida " + NFCX$
'     StatInit = 0
'     ImpoDeb = TOTANETX# * SIGX% * (-1)
'     Call GRACABAN
'   End If
'   '
'   ' GRABA IMPORTE EMBALAJES
'   TOTANETX# = NEMBAL7#
'   If TOTANETX# >= 0.005 Then
'     CUEMBA$ = Control$("SVD451S", "CUEVE-12")
'     If CUEMBA$ = "" Or CUECOINT%(CUEMBA$) < 1 Then CUEMBA$ = Text16
'     QQ1% = CUECOINT%(CUEMBA$)
'     If QQ1% < 1 Then QQ1% = QQQ%
'     CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
'     TiMovi$ = "EMBAL"
'     OpMovi$ = OPMOVIX$
'     VaMovi$ = Label13
'     Descrip$ = "Gasto Embalajes " + NFCX$
'     StatInit = 0
'     ImpoDeb = TOTANETX# * SIGX% * (-1)
'     Call GRACABAN
'   End If
'     '
'   'graba iva inscripto
'   For KKI% = 0 To 1
'     TOTANETX# = NIVA7#(KKI%)
'     If TOTANETX# >= 0.005 Then
'       CUEIVA$ = Control$("SVD451S", "CUEIVA")
'       CuentaCo% = CUECOINT%(CUEIVA$)
'       TiMovi$ = "IVA"
'       OpMovi$ = "DeFi"
'       VaMovi$ = PORIVA(KKI%) + " %"
'       Descrip$ = "Iva Debito Fiscal " + NFCX$
'       StatInit = 0
'       ImpoDeb = TOTANETX# * SIGX% * (-1)
'       BasImpo# = ROUND#(XVALO(BASIVA(KKI%)) * TICAMBX#, 2)
'       Call GRACABAN
'     End If
'   Next KKI%
'   BasImpo# = 0
'   '
'   'graba iva NO inscripto
'   TOTANETX# = NIVA7#(2)
'   If TOTANETX# >= 0.005 Then
'     CUEIVA$ = Control$("SVD451S", "CUEADU")
'     CuentaCo% = CUECOINT%(CUEIVA$)
'     TiMovi$ = "IVA"
'     OpMovi$ = "DeFi"
'     VaMovi$ = "RNI"
'     Descrip$ = "IVA Resp.No Inscripto " + NFCX$
'     StatInit = 0
'     ImpoDeb = TOTANETX# * SIGX% * (-1)
'     BasImpo# = ROUND#(XVALO(BASIVA(2)) * TICAMBX#, 2)
'     Call GRACABAN
'   End If
'   BasImpo# = 0
'   '
'   'graba percepcion I.Brutos
'   TOTANETX# = PERIB7#
'   If TOTANETX# >= 0.005 Then
'     CUEIVA$ = Control$("PERIBRU", "CUECONTA")
'     CuentaCo% = CUECOINT%(CUEIVA$)
'     TiMovi$ = "IBRU"
'     OpMovi$ = "IB-B"
'     VaMovi$ = "PERC"
'     Descrip$ = "Perc.IB.Bs.As. " + NFCX$
'     StatInit = 0
'     ImpoDeb = TOTANETX# * SIGX% * (-1)
'     BasImpo# = ROUND#(XVALO(Total(2)) * TICAMBX#, 2)
'     Call GRACABAN
'   End If
'   BasImpo# = 0
'   '
'   'graba percepcion IVA
'   TOTANETX# = PERIVA7#
'   If TOTANETX# >= 0.005 Then
'     CUEPER$ = Control$("SVD451S", "CUERET")
'     CuentaCo% = CUECOINT%(CUEPER$)
'     TiMovi$ = "IVA"
'     OpMovi$ = "DeFi"
'     VaMovi$ = "PER-IVA"
'     Descrip$ = "Percep.IVA " + NFCX$
'     StatInit = 0
'     ImpoDeb = TOTANETX# * SIGX% * (-1)
'     BasImpo# = ROUND#(XVALO(Total(2)) * TICAMBX#, 2)
'     Call GRACABAN
'   End If
'   BasImpo# = 0
'   '
'   'Call BLIMESS("Calculando Vencimientos ...")
'   If TIFADOC% <> 6 Or TRIM$(DOCUDIF) = "" Then
'     Call GRADEUFA(TIPO%, VA%, NRO, NC%, FFI%, CPAG%, TOTAGEN#)
'   End If
'   'CajaBanc.Close
'   'SaDeuPen.Close
'   NOCLOSE% = 0
'   '
'   'Call BLIMESS("Movimiento de Stock ...")
'   NRO1 = XVALO(Text11.TEXT)
'   If TIFADOC% = 6 Then NRO1 = (-1) * XVALO(Text11.TEXT)
'   '
'   REMI = XVALO(Text10.TEXT)
'   REMI1 = REMI: If TIFADOC% = 3 Then REMI1 = (-1) * NRO1
'   If TIFADOC% = 3 Then   ' nota de debito
'     If MODODOC% = 0 Then ' por diferencia de precios
'       Call AJUPREVE
'       GoTo 80
'     End If
'   End If
'   '
'   If CONSICLI% < 1 Then
'       Call GRAREMIFA
'     Else
'       Call GRACONSI
'   End If
'   '
'   If TIFADOC% = 3 Or TIFADOC% = 6 Then
'     If TRIM$(DOCUDIF) <> "" Then
'       Call APLIDICAM
'     End If
'   End If
'   '
'   If MONEMI% > 1 Then
'     If TIFADOC% < 3 Then   ' graba noviniento cabeza
'       Call GRAMOCABE
'     End If
'   End If
'   '
'80 OKX% = 1
'   Screen.MousePointer = 1
'   Exit Sub
'
'98 Screen.MousePointer = 1
'   Call MENSERR(24, MERRO$)
'   '
End Sub
'
Sub GRAMOCABE()
   '
   MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
   If MODOMONE$ <> "BIMONETA" Then
      If MODOMONE$ <> "MULTIMON" Then
         Exit Sub
      End If
   End If
   '
   FEMI% = FECHANUM(Text9)
   NCLIE% = XVALO(Text1)
   TIPO% = TIFADOC%
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
   Call REPLA(RAPLI$, Chr$(TIPO%), 5, 1)
   Call REPLA(RAPLI$, MKI$(VA%), 6, 1)
   Call REPLA(RAPLI$, MKS$(NRO), 7, 4)
   Call REPLA(RAPLI$, MKI$(SUCUVEN%), 11, 2)
   Call REPLA(RAPLI$, NUDOCU$, 13, 18)
   Call REPLA(RAPLI$, MKI$(MONEMI%), 31, 2)
   '
   ' documento aplicado
   Call REPLA(RAPLI$, MKI$(FEMI%), 33, 2)
   Call REPLA(RAPLI$, MKI$(NCLIE%), 35, 2)
   Call REPLA(RAPLI$, Chr$(TIPO%), 37, 1)
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
'   '
'   MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESVENTA", "MODOMONE"))), 8)
'   If MODOMONE$ <> "BIMONETA" Then
'      If MODOMONE$ <> "MULTIMON" Then
'         Exit Sub
'      End If
'   End If
'   '
'   Screen.MousePointer = 11
'   NCI% = XVALO(Text1)
'   TIDOCAPLI% = 2
'   NUDOCAPLI& = XVALO(Mid$(DOCUDIF, 9, 8))
'   CANCE# = XVALO(Total(6)) * TICAMONE(MONEMI%)
'   TICAMBX# = XVALO(TICAMBIS): If TICAMBX# < 0.05 Then TICAMBX# = 1
'   '
'   If TIFADOC% = 6 Then
'     Call APERBASDAT("CABAN")
'     DOCUCA$ = DOCUDIF
'     ICAN# = CANCE#
'     '
'     NumeClie% = NCI%
'     DocOrCor$ = ""
'     TotOrig# = 0
'     IvaOrig# = 0
'     '
'     With SaDeuPen
'       .FindFirst "NuClien = " & NCI% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDdebe > 0"
'       If .NoMatch = False Then
'850      PARCAN# = ICAN#: If !SaldDebe < ICAN# Then PARCAN# = !SaldDebe
'         .Edit
'         !SaldDebe = !SaldDebe - PARCAN#
'         .Update
'         '
'         DocOrCor$ = !CodiCom_Cor
'         DocOrLar$ = !CodiCom_Lar
'         TotOrig# = !IBruComp
'         IvaOrig# = 0
'         On Error Resume Next
'         IvaOrig# = !IIvaComp
'         On Error GoTo 0
'         '
'         ICAN# = ICAN# - PARCAN#
'         If ICAN# >= 0.005 Then
'           .FindNext "CODICOM_LAR = '" & DOCUCA$ & "' AND SALDdebe > 0"
'           GoTo 850
'         End If
'       End If
'     End With
'     '
'     ImpApli# = CANCE#
'     RefApli$ = "Diferencia de Cambio " + DOCUDIF
'     Call GRAPLICOB
'     '
'     MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESVENTA", "MODOMONE"))), 8)
'     If MODOMONE$ = "BIMONETA" Or MODOMONE$ = "MULTIMON" Then
'       NCI% = NumeClie%
'       RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCI%))
'       If Len(RFF$) > 0 Then
'         If CANCE# >= 0.005 Then
'           ABUSQUE$ = DOCUDIF
'           For JKLX& = 1 To Len(RFF$) Step 4
'             RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
'             YCB$ = REGLEIDO$("CARBIMON", RECABI&)
'             If TRIM$(Mid$(YCB$, 13, 18)) = ABUSQUE$ Or TRIM$(Mid$(YCB$, 45, 18)) = ABUSQUE$ Then
'               DOCUCABE$ = TRIM$(Mid$(YCB$, 13, 18))
'               For JKLY& = 1 To Len(RFF$) Step 4
'                 RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
'                 YCB$ = REGLEIDO$("CARBIMON", RECABI&)
'                 If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
'                     MONEMORI% = CVI(Mid$(YCB$, 31, 2))
'                     TICAMBX# = TICAMONE(MONEMORI%)
'                     FEDOAPLI% = FECHANUM(Text9)
'                     NROO& = XVALO(Mid$(CoComLar$, 11))
'                   Call REPLA(YCB$, MKI$(FEDOAPLI%), 33, 2)
'                   Call REPLA(YCB$, MKI$(NCI%), 35, 2)
'                   Call REPLA(YCB$, Chr$(TIDOAP%), 37, 1)
'                   Call REPLA(YCB$, Chr$(VADOAP%), 38, 1)
'                   Call REPLA(YCB$, MKS$(NROO&), 39, 4)
'                   Call REPLA(YCB$, MKI$(SUC%), 43, 2)
'                   Call REPLA(YCB$, CoComLar$, 45, 18)
'                   Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
'                   Call REPLA(YCB$, MKD$(0), 65, 8)
'                   Call REPLA(YCB$, MKD$(CANCE#), 73, 8)
'                   Call REPLA(YCB$, MKD$(0), 81, 8)
'                   Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
'                   Call REPLA(YCB$, MKD$(0), 97, 8)
'                   Call REPLA(YCB$, MKD$(0), 105, 8)
'                   Call REPLA(YCB$, MKD$(0), 113, 8)
'                   '
'                   Call REGAPP("CARBIMON", YCB$)
'                   GoTo 835
'                 End If
'               Next JKLY&
'             End If
'           Next JKLX&
'         End If
'       End If
'835    Call CIERRARCH("CARBIMON")
'     End If
'     '
'   ElseIf TIFADOC% = 3 Then ' nota de debito
'     '
'     RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCI%))
'     ABUSQUE$ = TRIM$(DOCUDIF)
'     If Len(RFF$) > 0 Then
'       For JKLX& = 1 To Len(RFF$) Step 4
'         RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
'         YCB$ = REGLEIDO$("CARBIMON", RECABI&)
'         If Mid$(YCB$, 13, 18) = ABUSQUE$ Then
'             FEDOAPLI% = FECHANUM(Text9)
'           Call REPLA(YCB$, MKI$(FEDOAPLI%), 33, 2)
'           Call REPLA(YCB$, MKI$(NCI%), 35, 2)
'           Call REPLA(YCB$, Chr$(TIFADOC%), 37, 1)
'           Call REPLA(YCB$, Chr$(VARIDOC%), 38, 1)
'           Call REPLA(YCB$, MKS$(XVALO(Text11)), 39, 4)
'           Call REPLA(YCB$, MKI$(XVALO(Label12)), 43, 2)
'             DOCAPLI$ = CoComLar$
'           Call REPLA(YCB$, DOCAPLI$, 45, 18)
'           Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
'           Call REPLA(YCB$, MKD$(CANCE#), 65, 8)
'           Call REPLA(YCB$, MKD$(0), 73, 8)
'           Call REPLA(YCB$, MKD$(0), 81, 8)
'           Call REPLA(YCB$, MKD$(XVALO(TICAMBIS)), 89, 8)  ' NO PONE TIPO DE CAMBIO
'           Call REPLA(YCB$, MKD$(0), 97, 8)
'           Call REPLA(YCB$, MKD$(0), 105, 8)
'           Call REPLA(YCB$, MKD$(0), 113, 8)
'           '
'           Call REGAPP("CARBIMON", YCB$)
'           GoTo 39
'         End If
'       Next JKLX&
'     End If
'     '
'39   Call CIERRARCH("CARBIMON")
'     '
'   End If
'   DOCUDIF = ""
'   '
End Sub
'
Sub GRADEUFA(TIPO%, VA%, NRO, NC%, FFAC%, COPAGI%, TOTFAC#)
   '
   Dim FVI%(6), PLAII%(6), IM#(6), FEX(3)
   '
   LU$ = LUDAT$
   IORIG# = TOTFAC#
   '
   If TIPO% = 0 Then
     If IORIG# < 0 Then
       TIPO% = 7
     End If
   End If
   '
   If TIPO% > 3 Then
     NUPAG% = 1
     FVI%(1) = FFAC%
     IM#(1) = TOTFAC#     ' IORIG#
     GoTo 41
   End If
   '
   ' busca condicion de pago segun COPAGI%
20 Call BUCONPA(LU$, COPAGI%, COPAG$, FREF%, PLA$)
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
   For I% = 1 To 6
     FVI%(I%) = 0
   Next I%
   '
   FEFA = FECHANUM(FEBAS$)
   For I% = 1 To 6
     PLAI% = Asc(Mid$(PLA$, I%, 1))
     PLAII%(I%) = PLAI%
     If I% > 1 Then
       If (PLAII%(I%) < PLAII%(I% - 1) Or PLAII%(I%) <= 0) Then
         GoTo 31
       End If
     End If
     '
     FEXU = FEFA + PLAI%
     While DIENTRE(FEFA, FEXU) < PLAI%
       FEXU = DIPOST(FEXU)
     Wend
     '
     FVI%(I%) = FEXU
     NUPAG% = I%
   Next I%

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
   SIGX% = 1: If TIPO% > 3 Then SIGX% = (-1)
   For I% = 1 To NUPAG%
     If Abs(IM#(I%)) >= 0.005 Then
       FechEmi = CVDAT(FFAC%)
       IBruTot# = IORIG#
       NuorVen% = NuorVen% + 1
       FeVenc = CVDAT(FVI%(I%))
       ImVenc# = (Abs(IM#(I%))) * SIGX%
       Call GRADEUPEN
     End If
   Next I%
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
   CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   '
   REMI$ = "RT." + TRIM$(Text10.TEXT)
   NROO& = XVALO(Mid$(REMI$, 4))
   FF% = FECHANUM(Text9.TEXT)
   DEPOSALI% = DEPOEXPE%
   '
   PREFIREM$ = CONTROL$("", "PREFIREM")
   If PREFIREM$ = "" Then PREFIREM$ = "0001"
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
   REFE$ = RASOCLI$(NCLIE%)
   '
   YXI% = 0
   FIN& = ULTREG&("!CARDEFAC")
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", I&)
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
   Next I&
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
'   TIPCOM$ = TRIM$(REPLACAR$(Label7.Caption, ":", ""))
'   CPAG% = XVALO(TEXT21(1).TEXT)
'   CATIVAS$ = Label13.Caption
'   '
'   TCMP$ = TRIM$(TIPCOM$)
'   For KKI% = 0 To 13
'     FINFACTURA.Text2(KKI%) = ""
'   Next KKI%
'   '
'   TICOMPROBAN$ = TCMP$
'   FINFACTURA.Label2.Caption = TCMP$
'   FINFACTURA.Label1(0) = TCMP$ + ":"
'   FINFACTURA.Label4.Caption = CATIVAS$
'   '
'   CODFOR$ = TRIM$(Control$("", "FORFAC-" + Label13.Caption))
'   NROCON& = XVALO(Text11.TEXT)
'   NCN$ = TRIM$(Str$(NROCON&))
'   While Len(NCN$) < 8: NCN$ = "0" + NCN$: Wend
'   PREFI$ = Label12.Caption
'   NCN$ = PREFI$ + "-" + NCN$
'   NUCOMPROBAN$ = NCN$
'   FINFACTURA.Label3.Caption = NCN$
'   '
'5  COPIASFAC% = XVALO(Control$("", "COPIAFAC"))
'   COPIASREM% = XVALO(Control$("", "COPIAREM"))
'   If COPIASFAC% < 1 Then COPIASFAC% = 1
'   If COPIASREM% < 1 Then COPIASREM% = 1
'   '
'   If TIFADOC% > 2 Then
'      'COPIASFAC% = 2: COPIASREM% = 0
'      COPIASREM% = 0           ' POR TEMA CANTIDAD DE COPIAS FRIOLATINA
'   End If
'   '
'   If CONSICLI% = 1 Then
'      COPIASREM% = 0
'   End If
'   '
'   FINFACTURA.Text2(0) = TRIM$(Str$(COPIASFAC%))
'   FINFACTURA.Text2(1) = TRIM$(Str$(COPIASREM%))
'   FINFACTURA.Text2(2) = " "
'   'FINFACTURA.Text2(3) = TRIM$(Str$(Int(TICAMBX# * XVALO(Total(2).TEXT)))) + ".="
'   '\\\OT-05-0492-WAR-11/07/05///
'     COVADE = XVALO(Control("", "COEVADEC")) / 100 'COEFICIENTE DE VALOR AGREGADO
'     If COVADE < 0.01 Then COVADE = 1
'     FINFACTURA.Text2(3) = TRIM$(Str$(Int(COVADE * TICAMBX# * XVALO(Total(2).TEXT)))) + ".="
'   '\\\OT-05-0492-WAR-FIN
'   FINFACTURA.Label5.Caption = ""
'   For U& = 0 To 8
'      FINFACTURA.Text1(U&).TEXT = ""
'   Next U&
'   '\\\OT-06-0114-WAR-03/04/06///
'   FINFACTURA.Text4.TEXT = ""
'   FINFACTURA.Text5.TEXT = ""
'   'LIMPIO LOS TEXT DE DATOS DE LA EXPO
'   For JX% = 4 To 11
'    FINFACTURA.Text2(JX%) = ""
'   Next JX%
'   '\\\OT-06-0114-WAR-FIN///
'   EPAC$ = TRIM$(UCase$(EMPREAC$))
'   TTXX$ = ""
'   '
'   FINFACTURA.Label6.Enabled = False
'   FINFACTURA.CLAREDOL = "0"
'   If MONEMI% = 1 Then
'     FINFACTURA.Label6.Enabled = True
'   End If
'   '
'   FINFACTURA.Label6.Visible = False
'   FINFACTURA.Text3 = ""
'   FINFACTURA.Frame5.Visible = False
'   '\\\OT-06-0125-WAR-03/04/06///
'   FINFACTURA.Height = 6690
'   'SI FACTURA DE EXPO => SE MUESTRA DATOS DE EXPORTACIÓN
'   NC1% = XVALO(FORFACCS.Text1.TEXT)
'   If PIVACLI%(NC1%) = 5 Or PIVACLI%(NC1%) = 7 Then
'     FINFACTURA.Height = FINFACTURA.Height + 1935 '(8600)
'   End If
'   'CACULA EL PESO EN KILOS NETO
'   PST = 0
'   P1 = 0
'   PT = 0
'   For UKK& = 1 To ULTREG&("!CARDEFAC")
'     XKK$ = REGLEIDO$("!CARDEFAC", UKK&)
'     CI% = CVI(Mid$(XKK$, 5, 2))
'     CANTI1 = CVS(Mid$(XKK$, 1, 4))
'     P1 = PESUNI(CI%)
'     PST = P1 * CANTI1
'     PT = PT + PST
'   Next UKK&
'   FINFACTURA.Text2(11) = FORMATNUM$(PT, "F8.1")
'   '\\\OT-06-0125-WAR-FIN///
'   '18/04/07 (OT-07-0162) 'PARA VER LA COTIZACION DEL DOLAR CUANDO EMITIMOS UNA FACTURA EN PESOS
'   If MONEMI% <= 1 Then
'      NLIPRE% = XVALO(FORFACCS.TEXT21(0))
'      If MONELISTA%(NLIPRE%) = 2 Then
'         TTXX$ = "Emitido en " + DEMONECO$(MONEMI%) + " a cotización " + DEMONECO$(2) + " de " + TRIM$(FORMATNUM$(TICAMONE(2), "F9.4")) + ". "
'          FINFACTURA.Text4 = TTXX$
'      End If
'   End If
'   '18/04/07 (FIN)
'10 FINFACTURA.Show 1
'   Refresh
'   '
'   If FINFACTURA.Label5.Caption <> "APRO" Then
'      If COMALTER%("Realmente Quiere CANCELAR este Documento\\No, Continuar\Si, Cancelar") <> 2 Then GoTo 10
'      OKX% = (-1)
'      Exit Sub
'   End If
'   '
'   COPIASFAC% = XVALO(FINFACTURA.Text2(0))
'   If COPIASFAC% < 0 Or COPIASFAC% > 9 Then GoTo 5
'   COPIASREM% = XVALO(FINFACTURA.Text2(1))
'   If COPIASREM% > 9 Then GoTo 5
'   If COPIASREM% < 1 Then
'      Text10.TEXT = ""
'   End If
'   '
End Sub


Sub CARDETFAC()
   '
   HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC/ARCHECO(2)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle de Facturación (Mat.Codificado)"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left - 50))
   '
   Timer1.Enabled = True
10 PRF$ = "FACONGUI"
   NOFRESHPRE% = 1
   If CONSICLI% = 0 Then
      PRF$ = "FACDOSI"
      NOFRESHPRE% = 0
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "FRIOLA") > 0 Then
        PRF$ = "FACFRIOL"
      End If
      If MODESCU% = 1 Then PRF$ = "FACPORDE"
   End If
   If MODODOC% = 1 Then
      'Call SETTELIB
      PRF$ = "FAC-LIBR"
      ATRI$ = "/NC"
      ATRI$ = ATRI$ + "/TITUPAN=Detalle de Facturación (Texto Libre)"
      ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
      ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   End If
   List1.Clear
   '
   XXX% = VENTABU%(PRF$ + ATRI$)
   Call CALTOFAC
   Timer1.Enabled = False
   '
   NPEDID& = 0
   If XXX% < 0 Then
      Call CIERRARCH("!CARDEFAC")
      Call BLANARCH("!CARDEFAC")
      Call BLANFORMU
      'Total(0).TEXT = ""
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
   Screen.MousePointer = 11   ' Registrando FACTURA ...
   List1.Clear
   FIN& = ULTREG&("!CARDEFAC")
   For U& = 1 To FIN&
      Y$ = REGLEIDO$("!CARDEFAC", U&)
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
      End If
   Next U&
   '
   Call CARLISFAC
   '
   If NPEDID& > 0 Then
      For UK& = ULTREG&("PEDENCA") To 1 Step -1
        YX$ = REGLEIDO$("PEDENCA", UK&)
        If CVS(Left$(YX$, 4)) = NPEDID& Then GoTo 1300
      Next UK&
       GoTo 99
      '
1300  TEXT21(0).TEXT = Str$(Asc(Mid$(YX$, 268, 1)))
      TEXT21(1).TEXT = Str$(Asc(Mid$(YX$, 269, 1)))
      TEXT21(2).TEXT = Str$(Asc(Mid$(YX$, 270, 1)))
      '
      If XVALO(TEXT21(2).TEXT) < 1 Then
        Eco21(2).TEXT = "GERENCIA"
      End If
      '
      LENTREG$ = Mid$(YX$, 73, 96)
      If TRIM$(LENTREG$) <> "" Then
        LENTREGA(0).TEXT = Mid$(LENTREG$, 1, 32)
        LENTREGA(1).TEXT = Mid$(LENTREG$, 33, 32)
        LENTREGA(2).TEXT = Mid$(LENTREG$, 65, 32)
      End If
      '
      If Mid$(YX$, 275, 14) <> String$(14, 0) Then
         PORDESCU.TEXT = TRIM$(Mid$(YX$, 275, 14))
      End If
   End If
   '
99 On Error Resume Next
   Command14.SetFocus
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub
'
Sub CARLISFAC()
    List1.Clear
   FIN& = ULTREG&("!CARDEFAC")
   For U& = 1 To FIN&
      Y$ = REGLEIDO$("!CARDEFAC", U&)
      TTXX$ = Space$(128)
      CI% = CVI(Mid$(Y$, 5, 2))
      PRECIO = CVS(Mid$(Y$, 429, 4))
      PRELI = PRELISTA#(ARCHIPRE$, CI%)
      PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
      MAREPA$ = UCase$(Mid$(Y$, 461, 1)): If MAREPA$ <> "R" Then MAREPA$ = " "
      NPEDID& = XVALO(Mid$(Y$, 371, 8))
      CAN = CVS(Mid$(Y$, 1, 4))
      '
      Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
      Call REPLA(TTXX$, CODEXT$(CI%), 8, 15)
      Call REPLA(TTXX$, Mid$(Y$, 7, 64), 24, 26)
      PRELI = CVS(Mid$(Y$, 421, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PRELI), 3, 8, 2, 2), 50, 8)
      PORDESCUIN = CVS(Mid$(Y$, 425, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 59, 5)
      PREUNI = CVS(Mid$(Y$, 429, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, 2, 2), 64, 8)
      TOT# = CVD(Mid$(Y$, 433, 8))
      Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 2, 2), 74, 10)
      Call REPLA(TTXX$, Mid$(Y$, 461, 2), 86, 2) ' MARCA REPARACION
      List1.AddItem TTXX$
      '
      If GRUPIVA%(CI%) > 1 Then
         Call MENSERR(24, "Grupo de IVA no Procesable\para Código '" + TRIM$(CODEXT$(CI%)) + "'.\   \Se Procesará a la Tasa Básica del " + TRIM$(FORMATNUM$(ALIVA, "F5.2")) + " %.")
      End If
      '
   Next U&
   List1.Refresh
   '
End Sub
'
Sub CALTOFAC()
   '
   Static CLIA$, RNI%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PPVI% = PIVACLI%(NCLIE%)
       If PPVI% = 2 Then RNI% = 1
       If PPVI% = 5 Or PPVI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   CLIACT% = XVALO(Text1)
   '
'   TASA1 = XVALO(PORIVA(0).Caption)
'   TASA2 = XVALO(PORIVA(1).Caption)
'   TASA3 = XVALO(PORIVA(2).Caption)
'   SUNEFAC# = 0: SUNEFA1# = 0: SUNEFA2# = 0
   '
   Label18 = TRIM$(Str$(ULTREG&("!CARDEFAC")))
   SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
   SUMACAN = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     '
     X$ = REGLEIDO$("!CARDEFAC", U&)
     CIXXI% = CVI(Mid$(X$, 5, 2))
     CAN = CVS(Mid$(X$, 1, 4))
     SUMACAN = SUMACAN + CAN
     PORDESCUIN = CVS(Mid$(X$, 425, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 100 Then PORDESCUIN = 0
     '
     If Abs(CAN) > 0.05 Or MODODOC% > 0 Then
       TASARE = TASA1
       If TASACLICK >= 0.005 Then TASARE = TASACLICK
       If CIXXI% > 0 Then
         CILOCAL% = CIXXI%
         TASARE = PORCEIVA(CILOCAL%)
       End If
       If TASARE <> CVS(Mid$(X$, 441, 4)) Then
         Call REPLA(X$, MKS$(TASARE), 441, 4)
         Call GRAREG("!CARDEFAC", X$, U&)
       End If
       '
       PRELI = CVS(Mid$(X$, 421, 4))
       PREUNI = CVS(Mid$(X$, 429, 4)) '
       '
       If MODESCU% = 1 Then
         If PRELI < 0.00005 Then PRELI = PREUNI
         PREUNI = PRELI * (100 - PORDESCUIN) / 100
         Call REPLA(X$, MKS$(PREUNI), 429, 4)
       End If
       '
       If CAN = 1 Then
           TOT# = ROUND(CDbl(PREUNI), 2)
         Else
           TOT# = ROUND(CDbl(ROUND(CAN, 1) * ROUND(PREUNI, 4)))
       End If
       '
       If MODODOC% = 1 Then
          PREUNI = CVS(Mid$(X$, 429, 4))
          TOT# = CVD(Mid$(X$, 433, 8))
       End If
       '
       Call REPLA(X$, MKD$(TOT#), 433, 8)
       SUNEFAC# = SUNEFAC# + TOT#
       '
       TASARE = CVS(Mid$(X$, 441, 4))
       IVAITEM1# = (Int(TOT# * TASARE * APLIVA% + 0.5)) / 100
       Call REPLA(X$, MKD$(IVAITEM1#), 445, 8)
       Call REPLA(X$, MKD$(TOT# + IVAITEM1#), 453, 8)
       Call GRAREG("!CARDEFAC", X$, U&)
       '
       If Abs(TASARE - TASA2) < 0.01 Then
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
'   Total(0).TEXT = CSTRING$(MKD$(SUNEFAC#), 3, 12, 2, 2)
'   Total(0).Refresh
   '
   DESC$ = TRIM$(PORDESCU.TEXT)
   PODES = COEFDESCU(DESC$)
  ' TTT1# = XVALO(Total(0).TEXT) * PODES / 100
  ' Total(1).TEXT = CSTRING$(MKD$(TTT1#), 3, 12, 2, 2)
   '
 '  TTT2# = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
   '
   ' CALCULO DE EMBALAJE
   PEMBA = POREMBAL(CLIACT%)
   Ltota(10) = "Importe Total Neto"
   IEMBALA = "0.00"
   If TIFADOC% < 3 Then ' SOLO PARA FACTURAS, NO N/CRED Y N/DEB
     If PEMBA > 0 Then
       IEMBALA = TRIM$(FORMATNUM$(PEMBA * TTT2# / 100, "F12.2"))
      ' Ltota(10) = "Tot.Neto (Inc." + TRIM$(IEMBALA) + " Embal.)"
       TTT2# = TTT2# + XVALO(IEMBALA)
     End If
   End If
   '
  ' Total(2).TEXT = CSTRING$(MKD$(TTT2#), 3, 12, 2, 2)
   TLABEL28 = ""
   '
   TTTX# = SUNEFA2# * (100 - PODES) / 100
   'BASIVA(1) = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   'Total(4).TEXT = CSTRING$(MKD$(XVALO(BASIVA(1)) * XVALO(PORIVA(1).Caption) * APLIVA% / 100), 3, 12, 2, 2)
   '
'   TTTX# = (XVALO(Total(2).TEXT) - XVALO(BASIVA(1)))
'   BASIVA(0) = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
'   Total(3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(0)) * XVALO(PORIVA(0).Caption) * APLIVA% / 100), 3, 12, 2, 2)
   '
'   If APLIVA% < 1 Then
'     TLABEL28 = Total(2).TEXT
'     BASIVA(0) = ""
'     BASIVA(1) = ""
'   End If
   '
   Call CALTOFA2
   '
End Sub
'
Sub CALTOFA2()
   '
   Static CLIA$
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0
     If NCLIE% > 0 Then
       PPVI% = PIVACLI%(NCLIE%)
       If PPVI% = 2 Then RNI% = 1
     End If
   End If
   '
'   TTTX# = (XVALO(Total(3).TEXT) + XVALO(Total(4).TEXT)) * RNI%
'   BASIVA(2) = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
'   Total(5).TEXT = CSTRING$(MKD$(XVALO(BASIVA(2)) * XVALO(PORIVA(2).Caption) / 100), 3, 12, 2, 2)
   '
'   TTTX# = 0: For KKI% = 2 To 5
'                 TTTX# = TTTX# + XVALO(Total(KKI%).TEXT)
'              Next KKI%
'              TTTX# = TTTX# + XVALO(Text15)
'              TTTX# = TTTX# + XVALO(Text17)
'   Total(6).TEXT = CSTRING$(MKD$(TTTX#), 3, 12, 2, 2)
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
'   For U& = 1 To 2
'      X$ = REGLEIDO$("GRUPIVA", U&)
'      TAIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
'      PORIVA(U& - 1) = TRIM$(CSTRING$(MKS$(TAIVA(U& - 1)), 3, 6, 2, 2))
'   Next U&
'   PORIVA(2) = "50.00" ' responsable no inscripto
'   Call CIERRARCH("GRUPIVA")
'   ALIIVA = ALIVA
   '
End Sub

Sub GRACONSI()
   '
'   CMO$ = TRIM$(Control$("", "CODREFAC"))
'   If Len(CMO$) <> 2 Then
'     CMO$ = "RT"
'   End If
'   '
'   NUFACLAR$ = CoComLar$
'   NUFACORT$ = CoComCor$
'   '
'   NC1% = XVALO(Text1.TEXT)
'   Call APERBASDAT("STOCKS")
'   '
'   ' ESTO ES NUEVO
'   SQLX$ = " SELECT * FROM MOVISTO "
'   SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NC1% & " "
'   SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
'   SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
'   SQLX$ = SQLX$ + " AND DoPriLar <> '' "
'   SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
'   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DESELEC# & " "
'   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HASELEC# & " "
'   '
'   Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
'   '
'   FIN& = ULTREG&("!CARDEFAC")
'   '
'   DESCX$ = PORDESCU
'   PODES = COEFDESCU(DESCX$)
'   COEFPRE = (100 - PODES) / 100
'   '
'   TTT1# = XVALO(Total(0).TEXT) * PODES / 100
'   For I& = 1 To FIN&
'     Call TRACE(20, I&, FIN&)
'     X$ = REGLEIDO$("!CARDEFAC", I&)
'     CANTAA = CVS(Left$(X$, 4))
'     If Abs(CANTAA) > 0 Then
'       CI% = CVI(Mid$(X$, 5, 2))
'       DOPRIC$ = TRIM$(Mid$(X$, 387, 16))
'       NORIT% = CVI(Mid$(X$, 503, 2))
'       PREUNI = CVS(Mid$(X$, 429, 4)) * COEFPRE
'       NPEDCAR& = XVALO(Mid$(X$, 371, 16))
'       '
'       If CI% > 0 Then
'         '
'         SQLX$ = "DoPriCor = '" & DOPRIC$ & "' "
'         SQLX$ = SQLX$ + "AND Cod_Inte = " & CI% & " "
'         SQLX$ = SQLX$ + "AND Nume_Clie = " & NC1% & " "
'         SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
'         '
'         With CONSITMP   ' Movisto
'           .FindFirst SQLX$
'25         If .NoMatch = False Then
'             NPEDID& = 0
'             If IsNull(!NUPEDCLI) = False Then NPEDID& = !NUPEDCLI
'             If NPEDID& > 0 And NPEDID& <> NPEDCAR& Then
'               .FindNext SQLX$
'               GoTo 25
'             End If
'             '
'             ' ESTO ES PARA QUE NO CONFUNDA DEVOLUCIONES CON REMITOS - 23/05/07
'             If Sgn(!CantSalid - !CantIngre) <> Sgn(CANTAA) Then
'               .FindNext SQLX$
'               GoTo 25
'             End If
'             '
'             ' ESTO ES PARA QUE NO CONFUNDA LOS QUE YA ESTÁN PROCESADOS
'             If !DoSecLar = NUFACLAR$ Then
'               .FindNext SQLX$
'               GoTo 25
'             End If
'
'             .Edit
'             !DoSecCor = NUFACORT$
'             !DoSecLar = NUFACLAR$
'             !ValoUnit = PREUNI
'             !MoneVal = MONEMI%     ' moneda de emision
'             If MONEMI% <= 1 Then
'                 !TipoCam = TICAMONE(2) ' siempre cotizacion del dolar
'               Else
'                 !TipoCam = TICAMONE(MONEMI%)
'             End If
'             If CANTAA > 0 Then
'                 !CantSalid = CANTAA
'                 !CantIngre = 0
'               Else
'                 !CantSalid = 0
'                 !CantIngre = Abs(CANTAA)
'             End If
'             If NPEDID& < 1 Then
'               !NUPEDCLI = NPEDCAR&
'             End If
'             .Update
'             GoTo 29
'           End If
'         End With
'       End If
'       '
'       FECHA7% = FECHANUM(Text9)
'       VUNI7# = PREUNI
'       Call MOVISTOK(FECHA7%, CI%, "", CMO$, NUFACLAR$, NUFACORT$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NORIT%, VUNI7#, MONEMI%, NC1%, DEPOEXPE%, (-1) * CANTAA, NPEDCAR&)
'       '
'     End If
'29 Next I&
'   '
'   CONSITMP.Close
   '
End Sub

Sub GRAREMIFA()
   '
   CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   DEPOSALI% = DEPOEXPE%
   '
   NUFACLAR$ = CoComLar$
   NUFACORT$ = CoComCor$
   COPIASFAC% = XVALO(FINFACTURA.Text2(0))
   If COPIASFAC% < 1 Then
      NUFACLAR$ = ""
      NUFACORT$ = ""
   End If
   '
   PREFIREM$ = CONTROL$("", "PREFIREM")
   If PREFIREM$ = "" Then PREFIREM$ = "0001"
   'REMICOR$ = "RT." + TRIM$(Text10)
   REMILAR$ = Text10: While Len(REMILAR$) < 8: REMILAR$ = "0" + REMILAR$: Wend
   REMILAR$ = CMO$ + "-X-" + PREFIREM$ + "-" + REMILAR$
   If COPIASREM% < 1 Then
     REMICOR$ = NUFACORT$
     REMILAR$ = NUFACLAR$
   End If
   '
   NC1% = XVALO(Text1.TEXT)
   FF% = FECHANUM(Text9)
   '
   DESCX$ = PORDESCU
   PODES = COEFDESCU(DESCX$)
   COEFPRE = (100 - PODES) / 100
   '
   FIN& = ULTREG&("!CARDEFAC")
   '\\\OT-07-0023-RG-24/01/07///
   DEPOSALI% = XVALO(FINFACTURA.Text7)
   '\\\OT-07-0023-RG-FIN///
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", I&)
     CANTAA = CVS(Left$(X$, 4))
     If Abs(CANTAA) > 0 Then
       CI% = CVI(Mid$(X$, 5, 2))
       If CI% >= 0 Then                 ' para incluir remito texto libre
         NORIT% = I& - 1
         PREUNI# = CDbl(CVS(Mid$(X$, 429, 4)) * COEFPRE)
         '\\\OT-05-0562-WAR-02/08/05///
         SIG% = (-1)
         If TIFADOC% = 6 Then SIG% = 1 'NC
         '\\\OT-05-0562-WAR-02/08/05///
         NUPEDIORI& = XVALO(Mid$(X$, 371, 16))
         '
         Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NORIT%, PREUNI#, MONEMI%, NC1%, DEPOSALI%, SIG% * CANTAA, NUPEDIORI&) '(-1) * CANTAA)
         '
         If DESCOMPO%(CI%) > 0 Then
              Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NORIT%, PREUNI#, MONEMI%, NC1%, DEPOSALI%, CANTAA, NUPEDIORI&)  '(-1) * CANTAA)            MOVIMIENTO SALIDA PARA COMPONENTES DEL KIT
              Call AREXPLO(CI%, 1, 1)
              NUORIT% = 0
              For KKJ% = 1 To CAIT%
                NUORIT% = NUORIT% + 1
                CICO% = CIJ%(KKJ%)
                CACICO = USOI(KKJ%) * CANTAA
                Call MOVISTOK(FF%, CICO%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NUORIT%, PREUNI#, MONEMI%, NC1%, DEPOSALI%, SIG% * CACICO, NUPEDIORI&) '(-1) * CANTAA)
              Next KKJ%
         End If
         '
         If CANTAA > 0 Then ' PARA EVITAR CON NOTAS DE CREDITO
           CANAX# = CANTAA
           NUORITPE% = CVI(Mid$(X$, 503, 2))
           If NUPEDIORI& > 0 Then
             Call BLANPEDIPE(NC1%, FF%, REMILAR$, CI%, CANAX#, NUPEDIORI&, NUORITPE%)
           End If
         End If
         '
       End If
     End If
   Next I&
   '
End Sub

Sub AJUPREVE()
   '
   NC1% = XVALO(Text1.TEXT)
   Call APERBASDAT("STOCKS")
   
   FIN& = ULTREG&("!DIFPRECO")
   '
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!DIFPRECO", I&)
     PRENUEVO# = CVD(Mid$(X$, 45, 8))
     If PRENUEVO# > 0.00005 Then
       CI% = CVI(Mid$(X$, 1, 2))
       If CI% > 0 Then
         DOPRIL$ = TRIM$(Mid$(X$, 81, 18))
         NORIT% = CVI(Mid$(X$, 117, 2))
         PREUNI = CVD(Mid$(X$, 45, 8))
         '
         SQLX$ = "DoPriLar = '" & DOPRIL$ & "' "
         SQLX$ = SQLX$ + "AND Cod_Inte = " & CI% & " "
         SQLX$ = SQLX$ + "AND Nume_Clie = " & NC1% & " "
         SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
         '
         With Movisto
           .FindFirst SQLX$
25         If .NoMatch = False Then
             .Edit
             !ValoUnit = PREUNI
             .Update
           End If
         End With
       End If
     End If
   Next I&
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
     '
     If YAINDI% = 0 Then
         '
         Screen.MousePointer = 11
         
         Call ABREPEDCLI
         SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
         SQLX$ = SQLX$ + " FROM PEDENCA "
         SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
         SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
         SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
         SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
         SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
         SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
         Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
         '
         JJ& = 0
         With PediModifTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
             Do While Not .EOF
                 NRODES& = !NroPed
               TTXX$ = MKS$(NRODES&)
                 NOCO$ = !NroOcom
               If NOCO$ <> "" Then
                 TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
               End If
                 FEX = CVINT(!FECHEMIS)
               TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
                 NCLIE% = !NroClie
               TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
               JJ& = JJ& + 1
               Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
             .MoveNext
             Loop
           End If
         End With
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
5    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Cls
       EDITFORM.Picture1(KKI%).DrawWidth = 3
       EDITFORM.Picture1(KKI%).DrawStyle = 0
       EDITFORM.Picture1(KKI%).FillStyle = 1
     Next KKI%
     '
     REDOCU$ = REGACT$("PDOCLST")
     URE& = ULTREG&("PDOCSPL")
     LI& = 0: LS& = URE&
6    E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
     L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
     If DCC& < DOCUNU& Then LI& = L&: GoTo 6
     If DCC& > DOCUNU& Then LS& = L&: GoTo 6
     GoTo 8
     '
7    Screen.MousePointer = 1
     Call MENSERR(24, "Documento no Registrado")
     GoTo 9999
     '
8    TPSP$ = ""
     Screen.MousePointer = 11
     While L& > 1
       DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
       If DCC& < DOCUNU& Then GoTo 9
       L& = L& - 1
     Wend
     '
9    CAPAGINAS% = 0
     For KKL& = L& To URE&
       TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
       If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
       If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
         TPSP$ = TPSP$ + Mid$(TBUF$, 5)
       End If
     Next KKL&
     '
95   SPOOLSTRING$ = TPSP$
     For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Cls
       EDITFORM.Picture1(KKI%).Top = 0
       EDITFORM.Picture1(KKI%).Refresh
       EDITFORM.Picture1(KKI%).Height = 7170
       EDITFORM.Picture1(KKI%).Width = 11700
       EDITFORM.Picture1(KKI%).DrawWidth = 3
       EDITFORM.Picture1(KKI%).DrawStyle = 0
       EDITFORM.Picture1(KKI%).FillStyle = 1
     Next KKI%
     ALTUPAGINA = 7170
     TOPEPAGINA = 0
     PAGINACTIVA% = 0
     CAPAGINAS% = 0
     '
10   PPXX% = InStr(TPSP$, Chr$(255))
     If PPXX% > 0 Then
       COPRIAT$ = Left$(TPSP$, PPXX% - 1)
       TPSP$ = Mid$(TPSP$, PPXX% + 1)
       Call TRACE(20, FIN& - Len(TPSP$), FIN&)
       '
       If Len(COPRIAT$) > 0 Then
         COPRI% = Asc(Left$(COPRIAT$, 1))
         ATRIP$ = Mid$(COPRIAT$, 2)
         Call SHOWCOMMAND(COPRI%, ATRIP$)
       End If
       GoTo 10
       '
     End If
     '
     Screen.MousePointer = 1
     PAGINACTIVA% = 0
     For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
     Next KKI%
     EDITFORM.Picture1(PAGINACTIVA%).Visible = True
     EDITFORM.Picture1(PAGINACTIVA%).Refresh
     EDITFORM.VScroll1.Min = 0
     EDITFORM.VScroll1.Value = 0
     EDITFORM.VScroll1.Max = 1
     If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
     End If
     EDITFORM.Command3.Enabled = False
     'EDITFORM.Command5.Enabled = False
     EDITFORM.Show 1
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
    If DES% < 1 Then
2     FEX = HOY(HO$)
      HAS% = FEX
      HASS$ = FECHATEX$(FEX)
      DESS$ = "01" + Mid$(HASS$, 3)
      DES% = FECHANUM(DESS$)
    End If
    '
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
       GoTo 2
    End If
    '
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM MOVISTO  "
    SQLX$ = SQLX$ + " WHERE CDBL(FechMov)>= " & DES1# & " "
    SQLX$ = SQLX$ + "AND CDBL(FechMov)<= " & HAS1# & " "
    SQLX$ = SQLX$ + "AND CodiMovi= 'RT' "
    SQLX$ = SQLX$ + "AND LEFT(DoSecCor,2)='FC' "
    SQLX$ = SQLX$ + "AND Nume_Clie= " & NCLIE% & " "
    SQLX$ = SQLX$ + "ORDER BY FechMov ASC "
    Set DIFPRETMP = Stocks.OpenRecordset(SQLX$, 4)
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
            PRENUE# = PRELISTA(NOLIST$, !Cod_Inte)
            '
            Call REPLA(YYY$, MKI(!Cod_Inte), 1, 2)
            Call REPLA(YYY$, !DoPriCor, 3, 12)
            Call REPLA(YYY$, !DoSecCor, 15, 12)
            Call REPLA(YYY$, MKI(CVINT(!FechMov)), 27, 2)
            Call REPLA(YYY$, MKD(!CantSalid), 29, 8)
            Call REPLA(YYY$, MKD(!ValoUnit), 37, 8)
            Call REPLA(YYY$, !DoPriLar, 81, 18)
            Call REPLA(YYY$, !DoSecLar, 99, 18)
            Call REPLA(YYY$, MKI(!NuOrItem), 117, 2)
            Call REPLA(YYY$, MKD(PRENUE#), 45, 8)
            JJ& = JJ& + 1
            Call GRAREG("!DIFPRECO", YYY$, JJ&)
       .MoveNext
       Loop
    End With
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
   NCLIE% = XVALO(Text1)
   TASAX = ALIPERBRU#(NCLIE%, MINIPER#, VARIDOC%)
   '
   Text15 = ""
'   BIMPON# = XVALO(Total(2))
'   If BIMPON# * TICAMONE(MONEMI%) >= MINIPER# Then
'      PERSALTA# = BIMPON# * (TASAX / 100)
'      Text15 = FORMATNUM$(PERSALTA#, "F11.2")
'   End If
   '
   Call CALTOFA2
   '
End Sub
'
Sub CARDETPEDIPEN()
   '
   Static AGRUPAENTREGA%
   If AGRUPAENTREGA% = 0 Then
      AGRUPAENTREGA% = (-1)
      If CONTROL$("FACTUREM", "AGRUPEDI") = "TRANSUCU" Then AGRUPAENTREGA% = 1
   End If
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
    '\\\OT-07-0371-FL-01/10/07///
   SQLX$ = " SELECT PEDDETA.codiint, PEDDETA.CanPed, PEDDETA.CantEnt, PEDDETA.NroPed,PEDDETA.FechEmis,"
   SQLX$ = SQLX$ + "PEDDETA.FeSolEnt, PEDDETA.PreUnid, PEDDETA.MoneEmis, "
   SQLX$ = SQLX$ + "PEDDETA.TipoPed, PEDDETA.NuOrItem, PEDENCA.CondPag, "
   SQLX$ = SQLX$ + "PEDENCA.DomiEnt, PEDENCA.LocaEnt, PEDENCA.TranspEnt, PEDENCA.NTranspo "
   SQLX$ = SQLX$ + "FROM PEDDETA "
   SQLX$ = SQLX$ + "INNER JOIN PEDENCA ON (PEDDETA.NroPed = PEDENCA.NroPed) "
   SQLX$ = SQLX$ + "WHERE PEDDETA.NroClie = " & NC% & " "
   SQLX$ = SQLX$ + "AND PEDDETA.Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
   SQLX$ = SQLX$ + "AND PEDDETA.STATUS >= 0 "   ' APROBADOS
   '\\\OT-07-371-FL-FIN///
   '
   If HAS1# > 0 Then
      SQLX$ = SQLX$ + " AND CDBL(PEDDETA.FeSolEnt) <= " & HAS1# & " "
   End If
   '
   SQLX$ = SQLX$ + " AND (((PEDDETA.CanPed - PEDDETA.CantEnt) >= 0.05) OR (PEDDETA.TipoPed = 'ABIERTO' AND CDBL(PEDDETA.FinVigen) >= " & Int(CDbl(Now)) & ")) "
   SQLX$ = SQLX$ + " ORDER BY PEDDETA.NroPed, PEDDETA.NuOrItem;" 'ESTA LINEA LA AGREGUE
   '
   Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With PEDTEMP
     On Error Resume Next
     .MoveFirst
     '
     JJ& = 0: CAITCAR% = 0
     '\\\OT-07-0371-FL-01/10/07///
     DESCARTAPED% = 0
     '\\\OT-07-0371-FL-FIN///
     Do While .EOF = False
       CODIMAT% = !CodiInt
       If CODIMAT% > 0 And CODIMAT% <= NUMAS% Then
         CAN = !CanPed - !CantEnt
           If CAPEN < 0 Then CAPEN = 0
         If !TipoPed = "ABIERTO" Then CAPEN = 999999
         NPEDID& = !NroPed
         FEMIPED% = CVINT(!FECHEMIS)
         FENTRESOL% = CVINT(!FESOLENT)
         PRECI1# = !PreUnid
         MONEMA% = !MoneEmis
            If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
         COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         PRECIO = PRECI1# * COEFCON
           If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
         PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
           If PRELI < PRECIO Then PRELI = PRECIO
         PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         '
         '\\\OT-07-0371-FL-01/10/07///
         '
         
         If AGRUPAENTREGA% > 0 Then
            DOMIENPE$ = TRIM$(UCase$(!DOMIENT))
            LOCAENPE$ = TRIM$(UCase$(!LOCAENT))
            TRANENPE$ = TRIM$(UCase$(!TRANSPENT))
            CONDPAGO% = !CONDPAG
            NTRANS% = XVALO(!NTranspo)
           '
           If XVALO(Label32) < 1 Then
              If DOMIENPE$ <> "" Then LENTREGA(0) = DOMIENPE$: Label32 = "1"
              If LOCAENPE$ <> "" Then LENTREGA(1) = LOCAENPE$: Label32 = "1"
              If NTRANS% <> 0 Then Text19.TEXT = TRIM$(Str$(NTRANS%)): Label32 = "1"
              If TRANENPE$ <> "" Then LENTREGA(2) = TRANENPE$: Label32 = "1"
              If CONDPAGO% <> 0 Then TEXT21(1).TEXT = CONDPAGO%: Label32 = "1"
              
           End If
           If DOMIENPE$ <> "" Then If DOMIENPE$ <> LENTREGA(0) Then GoTo 28
           If LOCAENPE$ <> "" Then If LOCAENPE$ <> LENTREGA(1) Then GoTo 28
           If NTRANS% <> 0 Then If NTRANS% <> XVALO(Text19.TEXT) Then GoTo 28
           If TRANENPE$ <> "" Then If TRANENPE$ <> LENTREGA(2) Then GoTo 28
           If CONDPAGO% <> 0 Then If CONDPAGO% <> XVALO(TEXT21(1).TEXT) Then GoTo 28
         End If
         '
         '\\\OT-07-0371-FL-FIN///
         '
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = ""
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = 0
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
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
       '\\\OT-07-0371-FL-01/10/07///
       GoTo 29
28     DESCARTAPED% = 1
29   .MoveNext
       '\\\OT-07-0371-FL-FIN///
     Loop
   End With
   '
   Call SETULTREG("!CARDEFAC", JJ&)
   '
    '\\\OT-07-0371-FL-01/10/07///
   If DESCARTAPED% = 1 Then
      Call MENSERR(24, "Se han Descartado Pedidos que no Corresponden\con la Sucursal o con la Condición de Venta\del Pedido más Antiguo")
   End If
    '\\\OT-07-0371-FL-FIN///
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
   ' AGREGADO
     '\\\OT-07-0371-FL-01/10/07///
   Static AGRUPAENTREGA%
   If AGRUPAENTREGA% = 0 Then
      AGRUPAENTREGA% = (-1)
      If CONTROL$("FACTUREM", "AGRUPEDI") = "TRANSUCU" Then AGRUPAENTREGA% = 1
   End If
     '\\\OT-07-0371-FL-FIN///
   ' FIN AGREGADO
   '
   Screen.MousePointer = 11
   NC% = XVALO(Text1)
   Call ABREPEDCLI
   SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
   SQLX$ = SQLX$ + " FROM PEDENCA "
   SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
   SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
   SQLX$ = SQLX$ + " WHERE PEDENCA.NroClie = " & NC% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.STATUS >= 0 "   ' APROBADOS
   SQLX$ = SQLX$ + " AND PEDDETA.Status < 8 "    ' para que no este anulada ni cumplida
   SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "    ' para que no esté anulado
   SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
   Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   SELPEDIPEN.LICOPEN.Clear
   SELPEDIPEN.Label5 = "0"
   HAYPEDIPE% = 0
   With PediModifTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         Do While Not .EOF
             NRODES& = !NroPed
           TTXX$ = FORMATNUM$(NRODES&, "F6.0") + "  "
             NOCO$ = TRIM$(!NroOcom)
           If NOCO$ <> "" Then
             TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
           End If
             FEX = CVINT(!FECHEMIS)
           TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
             NCLIE% = !NroClie
           TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
           SELPEDIPEN.LICOPEN.AddItem TTXX$
           HAYPEDIPE% = 1
         .MoveNext
         Loop
      End If
   End With
   '
   Screen.MousePointer = 1
   If HAYPEDIPE% > 0 Then
        SELPEDIPEN.Show 1
      Else
        Call COMUNI("No Hay Pedidos APROBADOS Pendientes\de Entrega para el Cliente Elegido.")
   End If
   '
   CANSELE% = XVALO(SELPEDIPEN.Label5)
   If CANSELE% < 1 Then
      JJ& = 0
      GoTo 20     ' PARA BLANQUEAR ARCHIVO DE CARGA
   End If
   '
   Call APERBASDAT("PEDCLI")
     '\\\OT-07-0371-FL-01/10/07///
   SQLX$ = " SELECT PEDDETA.codiint, PEDDETA.CanPed, PEDDETA.CantEnt, PEDDETA.NroPed,PEDDETA.FechEmis,"
   SQLX$ = SQLX$ + "PEDDETA.FeSolEnt, PEDDETA.PreUnid, PEDDETA.MoneEmis,PEDDETA.TipoPed,PEDDETA.NuOrItem, PEDENCA.CondPag,"
   SQLX$ = SQLX$ + "PEDENCA.DomiEnt, PEDENCA.LocaEnt, PEDENCA.TranspEnt, PEDENCA.NTranspo"
   SQLX$ = SQLX$ + " FROM PEDDETA "
   SQLX$ = SQLX$ + "INNER JOIN PEDENCA ON (PEDDETA.NroPed = PEDENCA.NroPed) "
   SQLX$ = SQLX$ + " WHERE PEDDETA.NroClie = " & NC% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
   SQLX$ = SQLX$ + " AND PEDDETA.STATUS >= 0 "   ' APROBADOS
   SQLX$ = SQLX$ + " AND (((PEDDETA.CanPed - PEDDETA.CantEnt) >= 0.05) OR (PEDDETA.TipoPed = 'ABIERTO' AND CDBL(PEDDETA.FinVigen) >= " & Int(CDbl(Now)) & ")) "
   '
   ATXX% = 0
   For KKU& = 1 To SELPEDIPEN.LICOPEN.ListCount
      If SELPEDIPEN.LICOPEN.Selected(KKU& - 1) = True Then
         NPESEL& = XVALO(Left$(SELPEDIPEN.LICOPEN.List(KKU& - 1), 6))
         If ATXX% < 1 Then
             SQLX$ = SQLX$ + "AND (PEDDETA.NroPed = " & NPESEL& & " "
             ATXX% = 1
           Else
             SQLX$ = SQLX$ + "OR PEDDETA.NroPed = " & NPESEL & " "
         End If
      End If
   Next KKU&
   SQLX$ = SQLX$ + ") "
   SQLX$ = SQLX$ + " ORDER BY PEDDETA.NroPed, PEDDETA.NuOrItem;" 'ESTA LINEA LA AGREGUE
   '
   Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With PEDTEMP
     On Error Resume Next
     .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("Ningun Registro Coincide con Condición SQL.")
      End If
      On Error GoTo 0
     DESCARTAPED% = 0
     '
     JJ& = 0: CAITCAR% = 0
     Do While (.EOF = False)
       CODIMAT% = !CodiInt
       If CODIMAT% > 0 And CODIMAT% <= NUMAS% Then
         CAN = !CanPed - !CantEnt
           If CAPEN < 0 Then CAPEN = 0
         If !TipoPed = "ABIERTO" Then CAPEN = 999999
         NPEDID& = !NroPed
         FEMIPED% = CVINT(!FECHEMIS)
         FENTRESOL% = CVINT(!FESOLENT)
         PRECI1# = !PreUnid
         MONEMA% = !MoneEmis
         
            If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
         COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         PRECIO = PRECI1# * COEFCON
           If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
         PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
           If PRELI < PRECIO Then PRELI = PRECIO
         PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         '
         ' AGREGADO
         If AGRUPAENTREGA% > 0 Then
           DOMIENPE$ = TRIM$(UCase$(!DOMIENT))
           LOCAENPE$ = TRIM$(UCase$(!LOCAENT))
           TRANENPE$ = TRIM$(UCase$(!TRANSPENT))
           CONDPAGO% = !CONDPAG
           NTRANS% = XVALO(!NTranspo)
           '
           If XVALO(Label32) < 1 Then
              If DOMIENPE$ <> "" Then LENTREGA(0) = DOMIENPE$: Label32 = "1"
              If LOCAENPE$ <> "" Then LENTREGA(1) = LOCAENPE$: Label32 = "1"
              If NTRANS% <> 0 Then Text19.TEXT = TRIM$(Str$(NTRANS%)): Label32 = "1"
              If TRANENPE$ <> "" Then LENTREGA(2) = TRANENPE$: Label32 = "1"
              If CONDPAGO% <> 0 Then TEXT21(1).TEXT = CONDPAGO%: Label32 = "1"
              
           End If
           If DOMIENPE$ <> "" Then If DOMIENPE$ <> LENTREGA(0) Then GoTo 28
           If LOCAENPE$ <> "" Then If LOCAENPE$ <> LENTREGA(1) Then GoTo 28
           If NTRANS% <> 0 Then If NTRANS% <> XVALO(Text19.TEXT) Then GoTo 28
           If TRANENPE$ <> "" Then If TRANENPE$ <> LENTREGA(2) Then GoTo 28
           If CONDPAGO% <> 0 Then If CONDPAGO% <> XVALO(TEXT21(1).TEXT) Then GoTo 28
           
         End If
         ' FIN AGREGADO
         '
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = ""
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = 0
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
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
       ' AGREGADO
         '\\\OT-07-0371-FL-01/10/07///
       GoTo 29
28     DESCARTAPED% = 1
        '\\\OT-07-0371-FL-FIN///
       ' FIN AGREGADO
29     .MoveNext
     Loop
   End With
   '
20 Call SETULTREG("!CARDEFAC", JJ&)
   '
   ' AGREGADO
     '\\\OT-07-0371-FL-01/10/07///
   If DESCARTAPED% = 1 Then
      Call MENSERR(24, "Se han Descartado Pedidos que no Corresponden\con la Sucursal o con la Condición de Venta\del Pedido más Antiguo")
   End If
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
'
Sub GRAFACTURA_NUE(OKX%)
'   '
'   Dim NIVA7#(10), NIVA7DOL#(10), CUEIVX$(10)
'   '
'   COTIMONEM# = XVALO(TICAMBIS) ' para prever ajuste de moneda de emision
'   If COTIMONEM# < 0.05 Then COTIMONEM# = 1
'   '
'   OKX% = 0
'   GoSub APRUEBACOPIAS       ' APROBACION FINAL Y CANTIDAD DE COPIAS
'   GoSub SOLOREMITO          ' PARA EL CASO DE QUE LAS COPIAS DE FACTURA SEAN CERO
'   GoSub ICONTABLE           ' IMPUTACIÓN CONTABLE DE DEBITOS Y CREDITOS
'   GoSub VALIDATCLI          ' VALIDACION DE DATOS DEL CLIENTE
'   GoSub ARMANUMERO          ' ARMA NUMERO DE FACTURA CORTO Y LARGO
'   GoSub VERILIMCRE          ' VERIFICA LIMITE DE CREDITO
'   GoSub RECALTOTA           ' RECALCULA Y VERIFICA TOTALES
'   GoSub ACSALCUEN           ' ACTUALIZA SALDO DE LA CUENTA
'   GoSub RENUMHOJAS          ' RENUMERA SEGUN CANTIDAD DE HOJAS A CONSUMIR
'   GoSub VERINUMERO          ' BLOQUEA Y VERIFICA NUMERO DE FACTURA LIBRE
'   '
'   Call APERBASDAT("CABAN")
'   NOCLOSE% = 1
'   GoSub GRATOTFAC           ' GRABA CAJABANC PARA FACTURA
'   GoSub GRAVENETA           ' GRABA VENTA NETA
'   GoSub GRAEMBALA           ' GRABA IMPORTE EMBALAJES
'   GoSub GRABAIVA            ' graba iva inscripto Y NO INSCRIPTO
'   GoSub GRAPERCEP           ' graba PERCEPCIONES IVA E I.BRUTOS
'   '
'   'Call BLIMESS("Calculando Vencimientos ...")
''   If TIFADOC% <> 6 Or TRIM$(DOCUDIF) = "" Then
''     Call GRADEUFA(TIPO%, VA%, NRO, NC%, FFI%, CPAG%, TOTAGEN#)
''   End If
'   CajaBanc.Close
'   SaDeuPen.Close
'   NOCLOSE% = 0
'   Call LIBARCH("CUECORR")
'   '
'   'Call BLIMESS("Movimiento de Stock ...")
'   NRO1 = XVALO(Text11.TEXT)
'   If TIFADOC% = 6 Then NRO1 = (-1) * XVALO(Text11.TEXT)
'   '
'   REMI = XVALO(Text10.TEXT)
'   REMI1 = REMI: If TIFADOC% = 3 Then REMI1 = (-1) * NRO1
'   If TIFADOC% = 3 Then   ' nota de debito
'     If MODODOC% = 0 Then ' por diferencia de precios
'       Call AJUPREVE
'       GoTo 80
'     End If
'   End If
'   '
'   If CONSICLI% < 1 Then
'       Call GRAREMIFA
'     Else
'       Call GRACONSI
'   End If
'   '
'   If TIFADOC% = 3 Or TIFADOC% = 6 Then
''     If TRIM$(DOCUDIF) <> "" Then
''       Call APLIDICAM
''     End If
'   End If
'   '
'   If MONEMI% > 1 Then
'     If TIFADOC% <= 3 Then   ' graba noviniento cabeza en caso de factura o nota
'       Call GRAMOCABE        ' de debito en moneda extranjera
'     End If
'   End If
'   '
'80 OKX% = 1
'   Screen.MousePointer = 1
'   Exit Sub
'
'   '
'APRUEBACOPIAS:          ' PRIMERA VALIDACIÓN DE DATOS *******************************
'   NCLIE% = XVALO(Text1.TEXT)
'   If MARCONOTA.Visible = True Then Exit Sub
'   If ULTREG&("!CARDEFAC") < 1 Then Exit Sub
'   '
'   Call APROCOPIAS(OKX%)
'   COPIASFAC% = XVALO(FINFACTURA.Text2(0))
'   COPIASREM% = XVALO(FINFACTURA.Text2(1))
'   If COPIASREM% < 1 Then
'      Text10.TEXT = ""
'   End If
'   If OKX% < 0 Then Exit Sub
'   If COPIASFAC% + COPIASREM% < 1 Then Exit Sub
'   '
''   If COPIASFAC% > 0 Then
''     If XVALO(Total(6).TEXT) < 0.01 Then
''       Call MENSERR(24, "Imposible Emitir Factura SIN CARGO !!!")
''       Exit Sub
''     End If
''   End If
'Return                  ' ***********************************************************
''
'SOLOREMITO:             ' CASO DE QUE SE EMITE REMITO SOLO Y NO FACTURA
'   If COPIASFAC% < 1 Then
'      '
'      ' verificar que no sea de texto libre
'      FIN& = ULTREG&("!CARDEFAC")
'      For I& = 1 To FIN&
'         Call TRACE(20, I&, FIN&)
'         X$ = REGLEIDO$("!CARDEFAC", I&)
'         CI% = CVI(Mid$(X$, 5, 2))
'         If CI% < 1 Then
'            Call MENSERR(24, "Imposible Emitir Solo Remito\de Material No Codificado.")
'            OKX% = 0
'            Exit Sub
'         End If
'      Next I&
'      '
'      GoSub VERINUMERO
'      Call GRAREMIFA
'      Call LIBARCH("CUECORR")
'      Call PRIFACTURA
'      Screen.MousePointer = 1
'      OKX% = (-1)
'      Exit Sub
'   End If
'Return                  ' ***********************************************************
''
'ICONTABLE:              ' IMPUTACIÓN CONTABLE DE DEBITOS Y CREDITOS
'   PPVI% = PIVACLI%(NCLIE%)
'   If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
'   CUEVE$ = Control$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
'   QQQ% = CUECOINT%(CUEVE$)
'   '
'   If TIFADOC% = 3 Or TIFADOC% = 6 Then
'      OPIMPUTA.Option1.Value = True
'      Call OPIMPUTA.Option1_Click
'      '
'      '
'       If TIFADOC% = 6 Then
'        'Option1(5).Value = True
'        If Control$("", "IMPUNCRE") = "DIFERENCIACAMBIO" Then
'          OPIMPUTA.Option2.Value = True
'          Call OPIMPUTA.Option2_Click
'        ElseIf Control$("", "IMPUNCRE") = "VARIOS" Then
'          OPIMPUTA.Option3.Value = True
'          Call OPIMPUTA.Option3_Click
'        End If
'      End If
'      '
'      '
''      If TRIM$(DOCUDIF) <> "" Then
''        OPIMPUTA.Option2.Value = True
''      End If
'      '
'      OPIMPUTA.Show 1
'   End If
'Return                  ' ***********************************************************
''
'VALIDATCLI:             ' VALIDACION DE DATOS DE CLIENTE
'   If VALICLIE% < 1 Then
'      MERRO$ = "Número de Cliente no Válido"
'      GoTo 98
'   End If
'   NCLIE% = XVALO(Text1.TEXT)
'   '
'   If XVALO(TEXT21(0).TEXT) > 255 Then
'      MERRO$ = "Lista de Precios no Válida"
'      GoTo 98
'   End If
'   '
'   If XVALO(TEXT21(1).TEXT) > 255 Then
'      MERRO$ = "Condición de Pago no Válida"
'      GoTo 98
'   End If
'   '
'   If XVALO(TEXT21(2).TEXT) > 255 Then
'      MERRO$ = "Vendedor no Válido"
'      GoTo 98
'   End If
'   '
''   If XVALO(Total(6).TEXT) * TICAMBX# < 0.01 Then
''      MERRO$ = "Importe no Válido"
''      GoTo 98
''   End If
'   '
'   HOII% = HOY(HOS$): FEACT = HOII%
'   FEMI% = FECHANUM(Text9.TEXT)
'   ' LA SIGUIENTE VALIDACION SE SUPRIME A PEDIDO DE S.IGLESIAS DE I.GUIDI
'   'If FEMI% < DIANTE(FEACT) Or FEMI% > DIPOST(FEACT) Then
'   '   MERRO$ = "Fecha de Emisión no Válida"
'   '   GoTo 98
'   'End If
'   GoTo 99                 ' PASÓ LA VALIDACION DE DATOS
'   '
'98 Screen.MousePointer = 1
'   Call MENSERR(24, MERRO$)
'   Exit Sub
'   '
'99 Return                  ' ***********************************************************
''
'ARMANUMERO:                ' ARMA NUMERO DE FACTURA CORTO Y LARGO
'   TICOMX$ = "FC"
'   If TIPO% = 3 Then TICOMX$ = "ND"
'   If TIPO% = 6 Then TICOMX$ = "NC"
'   NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
'   NFCY$ = TICOMX$ + "-" + Label12 + "-" + NFCZ$ + "-" + Label13
'Return                     ' ***********************************************************
''
'VERILIMCRE:                 ' VERIFICA LIMITE DE CREDITO      '
'   If VERLICRE% = 1 Then
'     If TIFADOC% < 4 Then   ' EXCLUYE CREDITOS
'       '
'       For UKK& = 1 To ULTREG&("!CARDEFAC")
'         XKK$ = REGLEIDO$("!CARDEFAC", UKK&)
'         NROPEDXX& = XVALO(TRIM$(Mid$(XKK$, 371, 16)))
'         If NROPEDXX& < 1 Then GoTo 22
'         RFFKK$ = RECFILT$("AUTORENT", 1, MKS$(NROPEDXX&))
'         If Len(RFFKK$) < 4 Then GoTo 22 ' NO AUTORIZADO
'         REGAUTKK& = CVS(RFFKK$)
'         XAUT$ = REGLEIDO$("AUTORENT", REGAUTKK&)
'         If UCase$(Mid$(XAUT$, 31, 2)) <> "SI" Then GoTo 22
'       Next UKK&
'       GoTo 23
'       '
''22     LMCR# = LIMICRE(NCLIE%)
''       SDEU# = SALDCLI#(NCLIE%)
''       CHPA# = CHEPEACRE#(NCLIE%)
''       TPED# = XVALO(Total(6).TEXT) * TICAMBX#
''       If SDEU# + CHPA# + TPED# > LMCR# Then
''         TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2))
''         If TLIMI$ = "" Then TLIMI$ = "0.00"
''         Call COMUNI("Imposible Emitir Este Documento.\  \Supera el Límite de Crédito Asignado\para esta Cuenta de $ " + TLIMI$ + ".\  \Informe al Supervisor de Ventas.")
''         Exit Sub
''       End If
'     End If
'   End If
'   '
'23 Return
'   '
'Return                      ' ***********************************************************
''
'RECALTOTA:                  ' RE-CALCULO TOTALES DE FACTURA
'   DESC$ = TRIM$(PORDESCU.TEXT)
'   PODES = COEFDESCU(DESC$)
'   '
'   For KK% = 1 To 10: TOTAFAC#(1, KK%) = 0: Next KK%
'   TOTAFAC#(1, 1) = (XVALO(BASIVA(0)) + XVALO(TLABEL28)) / (1 - PODES / 100)
'   TOTAFAC#(1, 2) = XVALO(BASIVA(1)) / (1 - PODES / 100)
'   TOTAFAC#(1, 3) = XVALO(Total(0).TEXT)
'   TOTAFAC#(1, 4) = XVALO(Total(1).TEXT)
'   TOTAFAC#(1, 5) = XVALO(BASIVA(0)) + XVALO(TLABEL28)
'   TOTAFAC#(1, 6) = XVALO(Total(3).TEXT)
'   TOTAFAC#(1, 7) = XVALO(Total(5).TEXT)
'   TOTAFAC#(1, 8) = XVALO(Total(6).TEXT)
'   TOTAFAC#(1, 9) = XVALO(BASIVA(1))
'   TOTAFAC#(1, 10) = XVALO(Total(4).TEXT)
'   '
'   TICA = 1
'   For KK% = 1 To 10
'     TOTAFAC#(1, KK%) = ROUND#(TOTAFAC#(1, KK%) * TICAMBX#, 2)
'     TOTAFAC#(2, KK%) = TOTAFAC#(1, KK%) / TICA
'   Next KK%
'   '
'   TOTAGEN# = 0: TOTAGENDOL# = 0
'   NENETO# = XVALO(Total(2)) - XVALO(IEMBALA)
'   NEFA7# = ROUND#(NENETO# * TICAMBX#, 2)
'   NEFA7DOL# = ROUND#(NENETO#, 2)
'   TOTAGEN# = TOTAGEN# + NEFA7#
'   TOTAGENDOL# = TOTAGENDOL# + NEFA7DOL#
'   NEMBAL7# = ROUND#(XVALO(IEMBALA) * TICAMBX#, 2)
'   NEMBAL7DOL# = ROUND#(XVALO(IEMBALA), 2)
'   TOTAGEN# = TOTAGEN# + NEMBAL7#
'   TOTAGENDOL# = TOTAGENDOL# + NEMBAL7DOL#
'   For KKI% = 0 To 2
'     NIVA7#(KKI%) = ROUND#(XVALO(Total(KKI% + 3)) * TICAMBX#, 2)
'     NIVA7DOL#(KKI%) = ROUND#(XVALO(Total(KKI% + 3)), 2)
'     TOTAGEN# = TOTAGEN# + NIVA7#(KKI%)
'     TOTAGENDOL# = TOTAGENDOL# + NIVA7DOL#(KKI%)
'   Next KKI%
'   PERIB7# = ROUND#(XVALO(Text15) * TICAMBX#, 2)
'   PERIB7DOL# = ROUND#(XVALO(Text15), 2)
'   TOTAGEN# = TOTAGEN# + PERIB7#
'   TOTAGENDOL# = TOTAGENDOL# + PERIB7DOL#
'   PERIVA7# = ROUND#(XVALO(Text17) * TICAMBX#, 2)
'   PERIVA7DOL# = ROUND#(XVALO(Text17), 2)
'   TOTAGEN# = TOTAGEN# + PERIVA7#
'   TOTAGENDOL# = TOTAGENDOL# + PERIVA7DOL#
'   '
'Return                      ' ***********************************************************
''
'ACSALCUEN:                  ' ACTUALIZA SALDO DE LA CUENTA
'   LU$ = LUDAT$
'   ARCHIDEU$ = "SVD202"
'   ARCHICLI$ = "DEUDOR"
'   ARCHIFAC$ = "FACTUR"
'   RD% = 1
'   TIPO% = TIFADOC%
'   VA% = VARIDOC%
'   NRO = XVALO(Text11.TEXT)
'   NC% = NCLIE%
'   FFI% = FECHANUM(Text9.TEXT)
'   CPAG% = XVALO(TEXT21(1).TEXT)
'   '
'   '
'   'Call BLIMESS("Actualizando Saldo de Cuenta ...")
'   SIGX% = 1: If TIPO% = 6 Then SIGX% = (-1)
'   TTF# = TOTAGEN# * SIGX%
'   Call GRACLI(LU$, ARCHICLI$, NC%, TTF#, DATCLI$(), 2)
'   Call CIERRARCH("DEUDOR2")
'   '
'   FLIM% = FECHANUM("01/01/99")
'   If ULTREG&("FORMFAC") > 0 Then
'      FFS$ = REGLEIDO$("FORMFAC", 2&)
'      FLIM% = CVI(Mid$(FFS$, 19, 2))
'   End If
'Return                      ' ***********************************************************
''
'RENUMHOJAS:                 ' RENUMERA PARA CONSIDERAR LAS HOJAS QUE SE CONSUMIRAN
'   'FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
'   FORIPRE$ = FORELFAC$
'   CARENGLO& = ULTREG&("!CARDEFAC")
'   ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
'   If ITPORHOJA% < 1 Then
'     ITPORHOJA% = 80
'   End If
'   CAHOJNEC% = Int(0.99 + CARENGLO& / ITPORHOJA%)
'   '
''   If NFCAM% < 1 Then
''     Call NUFACTURA
''   End If
'   '
'27 NPP& = XVALO(Text11) + CAHOJNEC% - 1
'   Text11.TEXT = TRIM$(Str$(NPP&))
'   Text11.Refresh
'   '
'   TICOMX$ = "FC": OPMOVIX$ = "FAC"
'   If TIPO% = 3 Then TICOMX$ = "ND": OPMOVIX$ = "DEB"
'   If TIPO% = 6 Then TICOMX$ = "NC": OPMOVIX$ = "CRE"
'   NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
'   NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + Label13
'   NFCY$ = TICOMX$ + "-" + Label12 + "-" + NFCZ$ + "-" + Label13
'   '
'   OPMOVIX$ = "FC"
'   If TIFADOC% = 3 Then OPMOVIX$ = "ND"
'   If TIFADOC% = 6 Then OPMOVIX$ = "NC"
'   VAMOVIX$ = Label13
'   '
'Return                      ' ***********************************************************
''
'VERINUMERO:
'   ' ESTO ES PARA VERIFICAR QUE EL ARCHIVO ESTÉ BLOQUEADO
'   Call BLOQARCH("CUECORR")
'   CajaBanc.Close
'   If ISLOCKED%("CUECORR") < 1 Then
'      Call LIBARCH("CUECORR")
'      Call MENSERR(24, "No fue Posible Abrir la Base en Modo Exclusivo.\  \Espere un Instante y Vuelva a Grabar.")
'      Exit Sub
'   End If
'   '
'   If COPIASFAC% > 0 Then
'     If VERINUFA%(NFCY$) < 0 Then
'       Call LIBARCH("CUECORR")
'       Call NUFACTURA
'       NROREMIT& = XVALO(Text10)
'       Call NUREMIFA(NROREMIT&)
'       Text10.TEXT = TRIM$(Str$(NROREMIT&))
'       Call MENSERR(24, "Número de Comprobante '" + NFCY$ + "' Ocupado.\   \FLEXOFT ha Re-Numerado el Documento\como Número '" + TRIM$(Text11.TEXT) + "'.\  \Verifique y Vuelva a Grabar.")
'       Exit Sub
'     End If
'   End If
'   If TIFADOC% < 3 Then
'     If COPIASREM% > 0 Then
'       If VERINURE% < 0 Then
'         Call LIBARCH("CUECORR")
'         NREMIRE$ = TRIM$(Text10)
'         NROREMIT& = 0
'         Call NUREMIFA(NROREMIT&)
'         Text10.TEXT = TRIM$(Str$(NROREMIT&))
'         Call MENSERR(24, "Número de Remito '" + NREMIRE$ + "' Ocupado.\   \FLEXOFT ha Re-Numerado el Remito\como Número '" + TRIM$(Text10.TEXT) + "'.\  \Verifique y Vuelva a Grabar.")
'         Exit Sub
'       End If
'     End If
'   End If
'   '
'Return                      ' ***********************************************************
''
'GRATOTFAC:                  ' GRABA TOTAL DE FACTURA
'30 CoComCor$ = NFCX$
'   CoComLar$ = NFCY$
'   CodiEmp% = 0
'   CodiSuc% = 0
'   NumeClie% = GRUDIVCLI%(NC%)
'   NumeProv% = 0
'   FechEmi = CVDAT(FFI%)
'   Referen$ = RASOCLI$(NC%)
'   FeContab = CVDAT(FFI%)
'   FeAcredi = CVDAT(FFI%)
'   CuentaCo% = CUECOCLI%(NC%)
'   TiMovi$ = "FVEN"
'   OpMovi$ = OPMOVIX$
'   VaMovi$ = Label13
'   Descrip$ = Label7 + " " + Label13.Caption + " " + Label6.Caption + " " + Label12.Caption + "-" + Text11.TEXT
'   NumeTal = XVALO(Label12)
'   NumeCom = CDbl(Text11)
'   StatInit = 0
'   ImpoDeb = TOTAGEN# * SIGX%
'   ValoDola = TOTAGENDOL# * SIGX%
'   Observa$ = RASOCLI$(NC%)
'   Call GRACABAN
'Return                      ' ***********************************************************
''
'GRAVENETA:                  ' GRABA VENTA NETA
'   'graba venta neta - Verificar que QQQ sea la IC de combinacion Cliente-Tipo de Comprobante
'   TOTANETX# = NEFA7#
'   If TOTANETX# >= 0.005 Then
'     QQ1% = CUECOINT%(Text16)
'     If QQ1% < 1 Then QQ1% = QQQ%
'     CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
'     TiMovi$ = "VENTA"
'     OpMovi$ = OPMOVIX$
'     VaMovi$ = Label13
'     Descrip$ = "Mercaderia Vendida " + NFCX$
'     StatInit = 0
'     ImpoDeb = TOTANETX# * SIGX% * (-1)
'     ValoDola = NEFA7DOL# * SIGX% * (-1)
'     Call GRACABAN
'   End If
'Return                      ' ***********************************************************
''
'GRAEMBALA:                   ' GRABA IMPORTE EMBALAJES
'   TOTANETX# = NEMBAL7#
'   If TOTANETX# >= 0.005 Then
'     CUEMBA$ = Control$("SVD451S", "CUEVE-12")
'     If CUEMBA$ = "" Or CUECOINT%(CUEMBA$) < 1 Then CUEMBA$ = Text16
'     QQ1% = CUECOINT%(CUEMBA$)
'     If QQ1% < 1 Then QQ1% = QQQ%
'     CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
'     TiMovi$ = "EMBAL"
'     OpMovi$ = OPMOVIX$
'     VaMovi$ = Label13
'     Descrip$ = "Gasto Embalajes " + NFCX$
'     StatInit = 0
'     ImpoDeb = TOTANETX# * SIGX% * (-1)
'     ValoDola = NEMBAL7DOL# * SIGX% * (-1)
'     Call GRACABAN
'   End If
'Return                      ' ***********************************************************
''
'GRABAIVA:                   ' GRABA IVA INSCRIPTO Y NO INSCRIPTO
'   'graba iva inscripto
'   CUEIVX$(0) = Control$("SVD451S", "CUEIVA")
'   CUEIVX$(1) = Control$("SVD451S", "CUEIVB")
'   If TRIM$(CUEIVX$(1)) = "" Or CUECOINT%(CUEIVX$(1)) < 1 Then
'     CUEIVX$(1) = CUEIVX$(0)
'   End If
'   For KKI% = 0 To 1
'     TOTANETX# = NIVA7#(KKI%)
'     If TOTANETX# >= 0.005 Then
'       CuentaCo% = CUECOINT%(CUEIVX$(KKI%))
'       TiMovi$ = "IVA"
'       OpMovi$ = "DeFi"
'       VaMovi$ = PORIVA(KKI%) + " %"
'       Descrip$ = "Iva Debito Fiscal " + NFCX$
'       StatInit = 0
'       ImpoDeb = TOTANETX# * SIGX% * (-1)
'       ValoDola = NIVA7DOL#(KKI%) * SIGX% * (-1)
'       BasImpo# = ROUND#(XVALO(BASIVA(KKI%)) * TICAMBX#, 2)
'       If KKI% = 0 Then
'          BasImpo# = BasImpo# + ROUND#(XVALO(TLABEL28) * TICAMBX#, 2)
'       End If
'       Call GRACABAN
'     End If
'   Next KKI%
'   BasImpo# = 0
'   '
'   'graba iva NO inscripto
'   TOTANETX# = NIVA7#(2)
'   If TOTANETX# >= 0.005 Then
'     CUEIVA$ = Control$("SVD451S", "CUEADU")
'     CuentaCo% = CUECOINT%(CUEIVA$)
'     TiMovi$ = "IVA"
'     OpMovi$ = "DeFi"
'     VaMovi$ = "RNI"
'     Descrip$ = "IVA Resp.No Inscripto " + NFCX$
'     StatInit = 0
'     ImpoDeb = TOTANETX# * SIGX% * (-1)
'     ValoDola = NIVA7DOL#(2) * SIGX% * (-1)
'     BasImpo# = ROUND#(XVALO(BASIVA(2)) * TICAMBX#, 2)
'     Call GRACABAN
'   End If
'   BasImpo# = 0
'Return                      ' ***********************************************************
''
'GRAPERCEP:                  ' PERCEPCIONES I.BRUTOS E IVA
'   'graba percepcion I.Brutos
'   TOTANETX# = PERIB7#
'   If TOTANETX# >= 0.005 Then
'     CUEIVA$ = Control$("PERIBRU", "CUECONTA")
'     CuentaCo% = CUECOINT%(CUEIVA$)
'     TiMovi$ = "IBRU"
'     OpMovi$ = "IB-B"
'     VaMovi$ = "PERC"
'     Descrip$ = "Perc.IB.Bs.As. " + NFCX$
'     StatInit = 0
'     ImpoDeb = TOTANETX# * SIGX% * (-1)
'     ValoDola = PERIB7DOL# * SIGX% * (-1)
'     BasImpo# = ROUND#(XVALO(Total(2)) * TICAMBX#, 2)
'     Call GRACABAN
'   End If
'   BasImpo# = 0
'   '
'   'graba percepcion IVA
'   TOTANETX# = PERIVA7#
'   If TOTANETX# >= 0.005 Then
'     CUEPER$ = Control$("SVD451S", "CUERET")
'     CuentaCo% = CUECOINT%(CUEPER$)
'     TiMovi$ = "IVA"
'     OpMovi$ = "DeFi"
'     VaMovi$ = "PER-IVA"
'     Descrip$ = "Percep.IVA " + NFCX$
'     StatInit = 0
'     ImpoDeb = TOTANETX# * SIGX% * (-1)
'     ValoDola = PERIVA7DOL# * SIGX% * (-1)
'     BasImpo# = ROUND#(XVALO(Total(2)) * TICAMBX#, 2)
'     Call GRACABAN
'   End If
'   BasImpo# = 0
'Return                      ' ***********************************************************
''


End Sub
'
Function FORELFAC$()   ' DEVUELVE EL NOMBRE DEL FORMULARIO ELECTRONICO DE FACTURACIÓN
    '
    FEFAC$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
    TIPOCOM$ = TRIM$(Label6)
    If TIPOCOM$ = "(T.Libre)" Then
      If Label13.Caption <> "E" Then   ' EXCLUYE FACTURAS DE EXPORTACION
        FEFAC$ = TRIM$(CONTROL$("", "FACTEXLI"))
        If UCase$(Label13) = "B" Then
          FEFAC$ = TRIM$(CONTROL$("", "FACTELIB"))
        End If
      End If
    End If
    '\\\OT-06-0038-WAR-01/02/06///
    FORIALT$ = ""
    If TIFADOC% = 3 Or TIFADOC% = 6 Then 'SI NC O ND
      If TIFADOC% = 3 Then ' SI NOTA DE DEBITO
        If UCase$(Label13.Caption) = "A" Then
          FORIALT$ = CONTROL$("", "FORNDE-A") 'ND-A
          ElseIf UCase$(Label13.Caption) = "B" Then
          FORIALT$ = CONTROL$("", "FORNDE-B") 'ND-B
        End If
      End If
      '
      If TIFADOC% = 6 Then ' SI NOTA DE CREDITO
        If UCase$(Label13.Caption) = "A" Then
          FORIALT$ = CONTROL$("", "FORNCR-A") 'NC-A
          ElseIf UCase$(Label13.Caption) = "B" Then
          FORIALT$ = CONTROL$("", "FORNCR-B") 'NC-B
        End If
      End If
    End If
    If TRIM$(FORIALT$) <> "" Then
      FEFAC$ = FORIALT$
    End If
    ''\\\OT-06-0038-WAR-FIN///
    '
    ' VALIDACIONES
    If FEFAC$ = "" Then
       Call MENSERR(24, "Falta Definir Controlador de Impresión")
       FEFAC$ = "NULL"
    End If
    '
    If EXISTE%(LUDAT$ + FEFAC$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Documento.\No Existe Formulario Electrónico '" + FEFAC$ + ".FRM'.")
       FEFAC$ = "NULL"
    End If
    '
    FORELFAC$ = FEFAC$
    '
End Function

Sub SETTELIB()
    FORIPRE$ = FORELFAC$
    If FORIPRE$ = "NULL" Then
       Call FINAL("")
       Exit Sub
    End If
    AA1$ = ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")
    ACOLDE = Val(Mid$(AA1$, 2))
    If ACOLDE < 1 Or ACOLDE > 64 Then ACOLDE = 64 ' PARA PODER MANDAR LO AL !CARDEFAC
    '
    TEPRUE.Font = Text28.Font
    TEPRUE.FontSize = Text28.FontSize
    TEPRUE.FontBold = Text28.FontBold
    TEPRUE = String$(ACOLDE + 3, "*")
    ATEX = TEPRUE.Width + 50
    '
          ANCHOTO = Label14.Width + Text28.Width + Label15.Width
          If ATEX > ANCHOTO - 50 Then ATEX = ANCHOTO - 50
          Label14.Width = (ANCHOTO - ATEX) / 2
          Label15.Width = (ANCHOTO - ATEX) / 2
          Text28.Left = Label15.Left + Label15.Width
          Text28.Width = ATEX
          Label14.Left = Text28.Left + Text28.Width
          '
    '
    Frame12.Top = Picture5.Top
    Frame12.Left = Picture5.Left
    Frame12.Visible = True
    '

End Sub
