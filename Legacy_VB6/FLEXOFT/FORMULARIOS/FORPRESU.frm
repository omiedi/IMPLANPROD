VERSION 5.00
Begin VB.Form FORPRESU 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuestos - Generación y Emisión"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11910
   Icon            =   "FORPRESU.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11910
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2325
      Left            =   0
      TabIndex        =   39
      Top             =   5880
      Width           =   10800
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otros"
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
         Height          =   1065
         Left            =   5880
         TabIndex        =   60
         Top             =   1260
         Width           =   4740
         Begin VB.TextBox Total 
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
            Index           =   6
            Left            =   2205
            TabIndex        =   30
            Text            =   " "
            Top             =   250
            Width           =   2385
         End
         Begin VB.TextBox Total 
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
            Index           =   5
            Left            =   2205
            TabIndex        =   31
            Text            =   " "
            Top             =   630
            Width           =   2385
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Validez de Oferta:"
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
            Index           =   6
            Left            =   265
            TabIndex        =   62
            Top             =   315
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimiento / Garantia:"
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
            Left            =   -50
            TabIndex        =   61
            Top             =   705
            Width           =   2220
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
         Height          =   1065
         Left            =   210
         TabIndex        =   57
         Top             =   1260
         Width           =   5580
         Begin VB.TextBox Total 
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
            TabIndex        =   29
            Text            =   " "
            Top             =   630
            Width           =   4275
         End
         Begin VB.TextBox Total 
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
            TabIndex        =   28
            Text            =   " "
            Top             =   250
            Width           =   4275
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar:"
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
            Left            =   -840
            TabIndex        =   59
            Top             =   680
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Plazo:"
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
            Left            =   -840
            TabIndex        =   58
            Top             =   315
            Width           =   1905
         End
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
         Index           =   1
         Left            =   9080
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   " "
         Top             =   405
         Width           =   1500
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
         Left            =   9080
         TabIndex        =   24
         Text            =   " "
         Top             =   0
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
         Index           =   3
         Left            =   9080
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   1500
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
         Left            =   2100
         TabIndex        =   45
         Text            =   " "
         Top             =   810
         Width           =   3690
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
         Left            =   1365
         TabIndex        =   19
         Text            =   " "
         Top             =   810
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
         Left            =   1785
         TabIndex        =   20
         Top             =   810
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
         Left            =   2100
         TabIndex        =   43
         Text            =   " "
         Top             =   405
         Width           =   3690
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
         Left            =   1365
         TabIndex        =   17
         Text            =   " "
         Top             =   405
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
         Left            =   1785
         TabIndex        =   18
         Top             =   405
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
         Left            =   2100
         TabIndex        =   41
         Text            =   " "
         Top             =   0
         Width           =   3690
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
         Left            =   1365
         TabIndex        =   15
         Text            =   " "
         Top             =   0
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
         Left            =   1785
         TabIndex        =   16
         Top             =   0
         Width           =   180
      End
      Begin VB.TextBox TICAMBIS 
         Enabled         =   0   'False
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
         Left            =   6510
         MaxLength       =   14
         TabIndex        =   78
         Top             =   840
         Visible         =   0   'False
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
         MaxLength       =   14
         TabIndex        =   25
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   5985
         TabIndex        =   88
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label TOTPER 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   5985
         TabIndex        =   85
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label TOTIVA 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   5985
         TabIndex        =   84
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label11 
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
         Left            =   6050
         TabIndex        =   77
         Top             =   915
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label MONEMIS 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pesos"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   5985
         TabIndex        =   76
         ToolTipText     =   "Doble-Click para 'Dólares'"
         Top             =   420
         Visible         =   0   'False
         Width           =   1380
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
         Height          =   330
         Left            =   8715
         TabIndex        =   49
         Top             =   510
         Width           =   225
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Neto:"
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
         Left            =   7560
         TabIndex        =   48
         Top             =   915
         Width           =   1380
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento:"
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
         Index           =   1
         Left            =   5985
         TabIndex        =   47
         Top             =   175
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
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
         Left            =   7035
         TabIndex        =   46
         Top             =   75
         Width           =   1905
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
         Left            =   -2835
         TabIndex        =   44
         Top             =   880
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
         Left            =   -2835
         TabIndex        =   42
         Top             =   480
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
         Left            =   -315
         TabIndex        =   40
         Top             =   75
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   12000
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000009&
         Height          =   3970
         Left            =   210
         ScaleHeight     =   3915
         ScaleWidth      =   10470
         TabIndex        =   80
         Top             =   1785
         Visible         =   0   'False
         Width           =   10530
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
            Height          =   3435
            Left            =   -30
            TabIndex        =   83
            Top             =   450
            Width           =   10480
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00E0E0E0&
            Height          =   490
            Left            =   -30
            ScaleHeight     =   435
            ScaleWidth      =   10410
            TabIndex        =   81
            Top             =   -30
            Width           =   10470
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "  Cant Código          Descripción                    Lista  Desc  P.Unit     Importe"
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
               Left            =   0
               TabIndex        =   82
               ToolTipText     =   "Haga Doble-Click para Cargar Detalle."
               Top             =   120
               Width           =   10515
            End
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   10500
            Y1              =   3900
            Y2              =   3900
         End
         Begin VB.Line Line1 
            X1              =   10450
            X2              =   10450
            Y1              =   0
            Y2              =   3920
         End
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
         Height          =   480
         Left            =   10920
         Picture         =   "FORPRESU.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Vista Previa de Impresion"
         Top             =   7245
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6950
         Left            =   10920
         TabIndex        =   65
         Top             =   105
         Width           =   850
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
            Height          =   770
            Left            =   0
            Picture         =   "FORPRESU.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
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
            Height          =   770
            Left            =   0
            Picture         =   "FORPRESU.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   3090
            Width           =   855
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "C-Com"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORPRESU.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Condiciones Comerciales"
            Top             =   5410
            Width           =   855
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Copia"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORPRESU.frx":0BB2
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Nuevo Presupuesto por Copia"
            Top             =   6170
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Articulos"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORPRESU.frx":0EBC
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   1540
            Width           =   855
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
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
            Height          =   770
            Left            =   0
            Picture         =   "FORPRESU.frx":11C6
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   2315
            Width           =   855
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00E0E0E0&
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
            Height          =   770
            Left            =   0
            Picture         =   "FORPRESU.frx":1608
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   4640
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Help"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORPRESU.frx":1A4A
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   770
            Width           =   855
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "List"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORPRESU.frx":1D54
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Consultas y Listados"
            Top             =   3860
            Width           =   855
         End
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
         Height          =   480
         Left            =   10920
         Picture         =   "FORPRESU.frx":205E
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Aprueba, Graba e Imprime Presupuesto"
         Top             =   7730
         Width           =   855
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Artículo"
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
         Height          =   700
         Left            =   210
         TabIndex        =   63
         Top             =   5040
         Width           =   8520
         Begin VB.TextBox Text13 
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
            Left            =   210
            TabIndex        =   21
            Text            =   " "
            Top             =   280
            Width           =   2325
         End
         Begin VB.CommandButton Command18 
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
            Left            =   2520
            TabIndex        =   22
            Top             =   280
            Width           =   180
         End
         Begin VB.Label Label8 
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
            Left            =   2940
            TabIndex        =   64
            Top             =   285
            Width           =   5415
         End
      End
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2745
         Left            =   1470
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Top             =   2215
         Width           =   7890
      End
      Begin VB.CommandButton Command10 
         Height          =   260
         Left            =   420
         TabIndex        =   53
         Top             =   6855
         Width           =   2825
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   210
         ScaleHeight     =   375
         ScaleWidth      =   10455
         TabIndex        =   51
         Top             =   1785
         Width           =   10515
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción General - Especificaciones"
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
            TabIndex        =   52
            Top             =   105
            Width           =   10515
         End
      End
      Begin VB.TextBox Text12 
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
         MaxLength       =   40
         TabIndex        =   6
         Text            =   " "
         Top             =   1280
         Width           =   3690
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
         TabIndex        =   13
         Text            =   " "
         ToolTipText     =   "Haga un Doble-Click para Ingresar el Número del Próximo Presupuesto"
         Top             =   890
         Width           =   1500
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
         Left            =   9240
         TabIndex        =   14
         Text            =   " "
         Top             =   1280
         Width           =   1500
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
         Left            =   9615
         TabIndex        =   11
         Text            =   " "
         Top             =   105
         Width           =   960
      End
      Begin VB.CommandButton Command9 
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
         Left            =   10560
         TabIndex        =   12
         Top             =   105
         Width           =   180
      End
      Begin VB.TextBox Text7 
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
         Left            =   5565
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   10
         Text            =   " "
         Top             =   1280
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
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
         Left            =   7060
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   105
         Width           =   180
      End
      Begin VB.TextBox Text6 
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
         Left            =   6650
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   435
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
         Left            =   2310
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   2850
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
         Left            =   1470
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   750
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
         Left            =   1470
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   3690
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
         Left            =   1470
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   3690
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1050
         TabIndex        =   1
         Top             =   1280
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
         TabIndex        =   0
         Text            =   " "
         Top             =   1280
         Width           =   855
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad"
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
         Height          =   700
         Left            =   8820
         TabIndex        =   86
         Top             =   5040
         Width           =   1905
         Begin VB.TextBox Text14 
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
            Left            =   240
            TabIndex        =   23
            Text            =   " "
            Top             =   280
            Width           =   1500
         End
         Begin VB.CommandButton Command16 
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
            Left            =   2520
            TabIndex        =   87
            Top             =   280
            Width           =   180
         End
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item Unico"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   7770
         TabIndex        =   79
         ToolTipText     =   "Doble-Click para 'Dólares'"
         Top             =   480
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   2730
         Left            =   9355
         TabIndex        =   56
         Top             =   2220
         Width           =   1365
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   2730
         Left            =   210
         TabIndex        =   55
         Top             =   2215
         Width           =   1275
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   10710
         Y1              =   4970
         Y2              =   4970
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono:"
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
         Left            =   5565
         TabIndex        =   50
         Top             =   975
         Width           =   1695
      End
      Begin VB.Line Line8 
         X1              =   10730
         X2              =   10730
         Y1              =   1785
         Y2              =   4950
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Picture         =   "FORPRESU.frx":2368
         ToolTipText     =   "Click Aquí para cambiar Modalidad de Facturación"
         Top             =   315
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Presup.Nro:"
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
         Left            =   7350
         TabIndex        =   38
         Top             =   945
         Width           =   1800
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia:"
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
         Left            =   7350
         TabIndex        =   37
         Top             =   1365
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
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   8820
         TabIndex        =   36
         Top             =   150
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
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   5670
         TabIndex        =   35
         Top             =   150
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         Left            =   315
         TabIndex        =   34
         Top             =   945
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
End
Attribute VB_Name = "FORPRESU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$
Dim ALPORCOPIA%

Private Sub Label12_DblClick()
   If TRIM$(UCase$(Label12)) = "ITEM UNICO" Then
       Label12.ForeColor = RGB(200, 0, 0)
       Label12 = "Lista Repuestos"
       Label12.ToolTipText = "Doble-Click para 'Item Unico'"
       Picture1.Visible = True
     Else
       Label12.ForeColor = RGB(0, 80, 80)
       Label12 = "Item Unico"
       Label12.ToolTipText = "Doble-Click para 'Lista Repuestos'"
       Picture1.Visible = False
   End If
End Sub


Private Sub Command1_Click()
   Call SELECHO("DEUDOR1")
   NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% > 0 Then
      Text1.TEXT = TRIM$(Str$(NCLIE%))
      Call CARDACLI
      Text12.SetFocus
   End If
   End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   If EXISTE%("AMACLI02.EXE") > 0 Then
       Call CONECRUN("AMACLI02", "Padron Maestro de Clientes")
     Else
       Call CONECRUN("AMACLI01", "Padron Maestro de Clientes")
   End If
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   OPLIPRESU.Show 1
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   OPPRESUP.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
15 ANEXPRESU.Show 1
   Refresh
   If ANEXPRESU.CONTROLABEL.Caption = "OK" Then
     SENDBYMAIL$ = ""
     If ANEXPRESU.Option2.Value = True Then
       PORMAIL% = 1
       NCI% = Val(Text1)
       emai$ = TRIM$(EMAILCLI$(NCI%))
       If emai$ = "" Or InStr(emai$, "@") < 1 Then
         Call MENSERR(24, "Imposible Enviar por Correo Electrónico.\  \Dirección de Correo del Cliente es\Inexistente o no Válida.")
         GoTo 15
       End If
       SENDBYMAIL$ = emai$
       '
       For Each prx1 In Printers
         If InStr(UCase$(prx1.DeviceName), "PDFWRITER") > 0 Then GoTo 20
       Next
       Call MENSERR(24, "Imposible Enviar por Correo Electrónico.\  \Falta Emulador PDF-Writer.")
       GoTo 15
     End If
     '
20   Call GRAPRESU(OKX%)
     ''\\\OT-05-0649-WAR-30/09/05///
     ALPORCOPIA% = 0
     ''\\\OT-05-0649-WAR-FIN///
     If OKX% > 0 Then
       CANCELPRINT% = 0
       Call PRIPRESU(0) ' ES IMPRESION DEFINITIVA
       Call BLANFORMU
       Text1.TEXT = ""
       '
       If SENDBYMAIL$ <> "" Then
         Call CONECRUN(App.EXEName, "")
         Call FINAL("")
         Exit Sub
       End If
       '
       Text1.SetFocus
     End If
   End If
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   If Val(Text11.TEXT) < 1 Then Exit Sub
   Command15.Enabled = False
   CANCELPRINT% = 1
   Screen.MousePointer = 11
   Call PRIPRESU(1)     ' ES VISTA PREVIA
   '
   TIDOCUM$ = "Presup"
   TIDOCUM2$ = "F. Proforma"
   ' NUDOCU$ = " " + TRIM$(Text11.TEXT) + " -"
   NUDOCU$ = "(VP-" + TRIM$(Text11.TEXT) + ")"
   Screen.MousePointer = 11
   '
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Or InStr(XX$, TIDOCUM2$) > 4 Then
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
   Screen.MousePointer = 1
   Call COMUNI("Documento no Encontrado")
   GoTo 99
   '
5  For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
   Next KKI%
   '
   REDOCU$ = REGACT$("PDOCLST")
   URE& = ULTREG&("PDOCSPL")
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Screen.MousePointer = 1
   Call MENSERR(24, "Documento no Registrado")
   GoTo 99
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
   Next KKL&
   '
95 SPOOLSTRING$ = TPSP$
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
10 PPXX% = InStr(TPSP$, Chr$(255))
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
   UI& = UI& + 1
   XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
   If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
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
   EDITFORM.Command5.Enabled = False
   EDITFORM.Show 1
   '
99 Command15.Enabled = True
   '
End Sub

Private Sub Command18_Click()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
   ARCHIPRE$ = "MASTER"
   LPREX% = Val(TEXT21(0))
   If LPREX% > 0 Then
     If LPREX% < 256 Then
       If ECOARCH$("LISTAS", Chr$(LPREX%)) <> "" Then
         ARCHIPRE$ = TRIM$(Str$(LPREX%))
         While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
         ARCHIPRE$ = "LIPRE" + ARCHIPRE$
         DELISECO$ = DENOLISTA$(LPREX%)
         ARCHIPRE$ = ARCHIPRE$ + "/" + DELISECO$
       End If
     End If
   End If
   Call SELECHO(ARCHIPRE$)
   Text13 = VALACT1$(ARCHIPRE$)
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP01")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RESCUDEU.Text1.TEXT = Text1.TEXT
    PROLLAMA$ = "FORPRESU"
    Hide
    RESCUDEU.Show
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    '\\\OT-05-0649-WAR-30/09/05///
    RTA% = ALTERNA%("Listas de Precios de Venta\Condiciones de Pago")
    If RTA% = 1 Then
      Call CIERRARCH("*.*")
      LPREU% = Val(TEXT21(0))
      If LPREU% < 1 Then LPRE% = 1
      ARCHIPRU$ = TRIM$(Str$(LPREU%))
      While Len(ARCHIPRU$) < 3: ARCHIPRU$ = "0" + ARCHIPRU$: Wend
      ARCHIPRU$ = "LIPRE" + ARCHIPRU$
      Call FRESHECHO(ARCHIPRU$)
      '
      If EXISTE%("LISPRE02.EXE") > 0 Then
          Call CONECRUN("LISPRE02", "Listas de Precios de Venta")
        Else
          Call CONECRUN("ADMIPRE", "Listas de Precios de Venta")
      End If
    ElseIf RTA% = 2 Then
      Call CARCONPAGOS
    End If
    '\\\OT-05-0649-WAR-FIN///
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Text1.SetFocus
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    If ALPORCOPIA% = 1 Then
      Call MENSERR(24, "Imposible Continuar.Generación por Copia en Proceso.")
      GoTo 99
    End If
    Screen.MousePointer = 11
    JJ& = 0: NPX& = 0
    FIN& = ULTREG&("PRESUPU")
    If FIN& < 1 Then
      Call MENSERR(24, "Archivo Vacío")
      GoTo 99
    End If
    '
    FINI1& = ULTREG&("PRESUPU")
    XX$ = REGLEIDO$("PRESUPU", FINI1&)
    NPX1& = CVS(Left$(XX$, 4))
    FINI2& = ULTREG&("INDIPRES")
    YY$ = REGLEIDO$("INDIPRES", FINI2&)
    NPX2& = CVS(Left$(YY$, 4))
    If NPX1& = NPX2& Then GoTo 3
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
2     XX$ = REGLEIDO$("PRESUPU", U&)
      NPY& = CVS(Left$(XX$, 4))
      If NPY& > 0 Then
        If NPY& <> NPX& Then
          NPX& = NPY&
          NCLIE% = CVI(Mid$(XX$, 7, 2))
          CIXI% = CVI(Mid$(XX$, 69, 2))
            '
            'Call RECODOSI(CIXI%, CIXJ%, NPX&)
            'If CIXJ% <> CIXI% Then
            '  Call REPLA(XX$, MKI$(CIXJ%), 69, 2)
            '  Call GRAREG("PRESUPU", XX$, U&)
            '  GoTo 2
            'End If
            '
          FEX = CVI(Mid$(XX$, 5, 2))
          YY$ = MKS$(NPX&) + TRIM$(RASOCLI$(NCLIE%)) + " - "
          If CVS(Left$(REGLEIDO$("PRESUPU", U& + 1), 4)) <> NPX& Then
              YY$ = YY$ + TRIM$(CODEXT$(CIXI%))
            Else
              YY$ = YY$ + "VARIOS"
          End If
          YY$ = YY$ + " - " + FECHATEX$(FEX)
          JJ& = JJ& + 1
          Call GRAREG("INDIPRES", YY$, JJ&)
        End If
      End If
    Next U&
    Call SETULTREG("INDIPRES", JJ&)
    '
3   Screen.MousePointer = 1
    '\\\OT-05-0758-WAR-18/11/05///
4    OBJ1$ = OBJPLA$("SELECLIPRE", VDEF$)
    If TRIM$(OBJ1$) <> "" Then
      OBJANT$ = OBJ1$
      NPX& = CVS(Mid$(OBJ1$, 1, 4))
      NCLIE% = CVI(Mid$(OBJ1$, 5, 2))
      'VALIDACIONES
      If NPX& <= 0 Then
        Call COMUNI("Numero de Presupuesto no Válido")
        VDEF$ = OBJANT$
        GoTo 4
      End If
      If NCLIE% <= 0 Then
        Call COMUNI("Numero de Cliente no Válido o Inexistente")
        VDEF$ = OBJANT$
        GoTo 4
      End If
      '
      Call PREPORCOP(NPX&, NCLIE%)
    End If
    '\\\OT-05-0758-WAR-FIN///
    'CODIGO ANTIGUO
'    NPX& = Val(VALACT1$("INDIPRES"))
'    If NPX& < 1 Then GoTo 99
'    '
'    Call PREPORCOP(NPX&)
'    '
99  Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    If EXISTE%("AMASTO02.EXE") > 0 Then
        Call CONECRUN("AMASTO02", "Maestro de Artículos de Stock")
      Else
        Call CONECRUN("AMASTO01", "Maestro de Artículos de Stock")
    End If
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
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
    Show
    Refresh
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    Text11.TEXT = TRIM$(Str$(NUPRESUP&))
    '
    ' Call RECLAPAG   ' RECLAMO DE PAGO - DESHABILITADO 8/12/05 PARA DOMIZI
    '
    ARCHIPRE$ = "MASTER"
    '
    MONEMI% = 1  ' PESOS
    CONVERPRECIO% = 0
    'If InStr(EPAC$, "DOSIVA") > 0 Then
       MONEMIS.Visible = True
       CONVERPRECIO% = 1
       TICAMBIS.Visible = True
       Label11.Visible = True
    'End If
    '
    Label12.Visible = True
    TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
    '
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
   Call FINAL("")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Label13_Click()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
   Call CARDETPRES
End Sub

Private Sub Label8_Click()
   Text13.SetFocus
End Sub

Private Sub List1_DblClick()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
   Call CARDETPRES
End Sub

Private Sub MONEMIS_Change()
   MONEMA% = MONEMI%
   If Left$(MONEMIS, 1) = "D" Then
        MONEMIS.ForeColor = RGB(255, 0, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Pesos'"
        Label11.Visible = True
        TICAMBIS.Enabled = True
        TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
        MONEMI% = 2
      Else
        MONEMIS.ForeColor = RGB(0, 0, 255)
        MONEMIS.ToolTipText = "Doble-Click para 'Dólares'"
        Label11.Visible = False
        TICAMBIS.Enabled = False
        MONEMI% = 1
   End If '
   Total(0) = FORMATNUM$(Val(Total(0)) * TICAMONE(MONEMA%) / TICAMONE(MONEMI%), "F11.2")
   '
   If Picture1.Visible = True Then
      COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
      FIN& = ULTREG&("!CARDETPR")
      For U& = 1 To FIN&
        X$ = REGLEIDO$("!CARDETPR", U&)
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
        Call GRAREG("!CARDETPR", X$, U&)
      Next U&

      Call CARLISPRES
      Call CIERRARCH("!CARDETPR")
      Call CALTOPRES
   End If
   '
End Sub

Private Sub MONEMIS_DblClick()
   If TRIM$(UCase$(MONEMIS)) = "PESOS" Then
       MONEMIS = "Dólares"
     Else
       MONEMIS = "Pesos"
   End If
End Sub

Private Sub PORDESCU_Change()
   DESC$ = PORDESCU.TEXT
   TTT1 = Val(Total(0).TEXT) * COEFDESCU(DESC$) / 100
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   TTT3 = Val(Total(0).TEXT) - Val(Total(1).TEXT)
   Total(3).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
End Sub

Private Sub PORDESCU_GotFocus()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
End Sub

Private Sub Text1_Change()
  Call BLANFORMU
  '\\\OT-05-0649-WAR-30/09/05///
  ALPORCOPIA% = 0
  '\\\OT-05-0649-WAR-FIN///
End Sub

Private Sub Text1_DblClick()
  Call SELECHO("DEUDOR1")
  NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
  If NCLIE% > 0 Then
    Text1.TEXT = TRIM$(Str$(NCLIE%))
    Call CARDACLI
    Text12.SetFocus
  End If
End Sub

Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If VALICLIE% > 0 Then
      Call CARDACLI
      Text12.SetFocus
    End If
  End If
End Sub

Private Sub Text10_GotFocus()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
End Sub

Private Sub Text11_DblClick()
   Text11 = InputBox$("Número del Próximo Presupuesto", "Numeración de Presupuestos", Text11)
   If Val(Text11) < 1 Then
     Call FINAL("")
   End If
End Sub

Private Sub Text11_GotFocus()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
   If ULTREG&("PRESUPU") > 2 Then
      Text10.SetFocus
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
   NPP& = NUPRESUP&
   If Val(NUMEPRE$) < NPP& Then
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
   If VALICLIE% < 1 Then Text1.SetFocus
End Sub

Private Sub Text13_Change()
   If CODINT%(Text13) > 0 Then
      CI% = CODINT%(Text13)
      Label8.Caption = DESCRIT0$(CI%)
      LPREX% = Val(TEXT21(0))
      If LPREX% > 0 Then
        If LPREX% < 256 Then
          If ECOARCH$("LISTAS", Chr$(LPREX%)) <> "" Then
            ARCHIPRE$ = TRIM$(Str$(LPREX%))
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$
            PREUNI = PRELISTA(ARCHIPRE$, CI%)
            Total(0).TEXT = CSTRING$(MKS$(PREUNI), 3, 11, 2, 2)
            '\\\OT-5-851-FG-26/12/05///
            Label14.Caption = Total(0).TEXT ' GUARDA EL PRECIO UNITARIO
            Text14 = FORMATNUM$("1", "F11.1") ' ESTABLECE CANTIDAD A 1
            '\\\OT-5-851-FG-FIN///
          End If
        End If
      End If
      Call CARESPECI
    End If
End Sub

Private Sub Text13_DblClick()
   ARCHIPRE$ = "MASTER"
   LPREX% = Val(TEXT21(0))
   If LPREX% > 0 Then
     If LPREX% < 256 Then
       If ECOARCH$("LISTAS", Chr$(LPREX%)) <> "" Then
         ARCHIPRE$ = TRIM$(Str$(LPREX%))
         While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
         ARCHIPRE$ = "LIPRE" + ARCHIPRE$
         DELISECO$ = DENOLISTA$(LPREX%)
         ARCHIPRE$ = ARCHIPRE$ + "/" + DELISECO$
       End If
     End If
   End If
   Call SELECHO(ARCHIPRE$)
   Text13 = VALACT1$(ARCHIPRE$)
End Sub

Private Sub Text13_GotFocus()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
End Sub

'\\\OT-5-851-FG-26/12/05///
Private Sub Text14_Change()
   '
   CANTI = Val(Text14.TEXT)
   SUBTOT = Val(Label14) * CANTI
   Total(0) = FORMATNUM$(SUBTOT, "F11.2")
   '
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
'
If KeyAscii = 13 Then
  CANTI = Val(Text14.TEXT)
  Text14.TEXT = FORMATNUM$(CANTI, "F11.1")
End If
'
End Sub

Private Sub Text14_LostFocus()
   '
   CANTI = Val(Text14.TEXT)
   Text14.TEXT = FORMATNUM$(CANTI, "F11.1")
   '
End Sub
'\\\OT-5-851-FG-FIN///

Private Sub Text2_GotFocus()
   Text1.SetFocus
End Sub

Private Sub TEXT21_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
End Sub

Private Sub Text3_GotFocus()
   Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
   Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
   Text1.SetFocus
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
           FACIVA = 1 + ALICU / 100
           CATIVAS$ = "B"
           VA% = 2
           '
          ElseIf PIVA% = 2 Then
           ALICU = ALI1 + ALI2
           RNI% = 1
           '
          ElseIf PIVA% = 5 Then
           ALICU = 0
           CATIVAS$ = " "
           VA% = 0
           If LETREXPO$ = "A" Then
              VA% = 1
              CATIVAS$ = "A"
           End If
        End If
      Else
        Text6.TEXT = "1"
    End If
End Sub

Private Sub Text21_Change(Index As Integer)
    '
    If Val(TEXT21(Index).TEXT) < 0 Or Val(TEXT21(Index).TEXT) > 255 Then
       TEXT21(Index).TEXT = ""
    End If
    '
    If Index = 0 Then
        If Val(TEXT21(Index).TEXT) < 1 Then
            Eco21(Index).TEXT = "CONCERTADOS FIJOS"
            ARCHIPRE$ = "MASTER"
          Else
            Eco21(Index).TEXT = ECOARCH$("LISTAS", MKI$(Val(TEXT21(Index).TEXT)))
            ARCHIPRE$ = TRIM$(TEXT21(Index))
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        End If
        If MONELISTA%(Val(TEXT21(Index))) = 2 Then
             MONEMIS = "Dólares"
           Else
             MONEMIS = "Pesos"
        End If
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(Val(TEXT21(Index).TEXT)))
      ElseIf Index = 2 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(Val(TEXT21(Index).TEXT)))
    End If
    '
End Sub

Private Sub Command21_Click(Index As Integer)
    If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
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
   '
Function VALICLIE%()
   VCC% = 0
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE% = Val(Text1.TEXT)
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
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   Text6.TEXT = ""
   Text7.TEXT = ""
   Text8.TEXT = ""
   Text10.TEXT = ""
   Text12.TEXT = ""
   Text13.TEXT = ""
   Text14.TEXT = FORMATNUM$(1, "F11.1")
   Text28.TEXT = ""
   Label8.Caption = ""
   TOTIVA = ""
   TOTPER = ""
   '
   TEXT21(0).TEXT = "": Eco21(0) = ""
   TEXT21(1).TEXT = "": Eco21(1) = ""
   TEXT21(2).TEXT = "": Eco21(2) = ""
   '
   PORDESCU.TEXT = ""
   For KK% = 0 To 6
      Total(KK%).TEXT = ""
   Next KK%
   '
   List1.Clear
   Call SETULTREG("!CARDETPR", 0)
   Label12 = "Item Unico"
   Picture1.Visible = False
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Text11.TEXT = TRIM$(Str$(NUPRESUP&))
   '
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then
      If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Text6.TEXT = PIVACLI%(NC%)
        Text7.TEXT = POSIVAC$(NC%)
        Text8.TEXT = TEL1CLI$(NC%)
        '
        TEXT21(0).TEXT = Str$(LIPRCLI%(NC%))
        TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
        TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
        '
        If Val(TEXT21(2).TEXT) < 1 Then
           Eco21(2).TEXT = "GERENCIA"
        End If
        '
        PORDESCU.TEXT = DESCUCLI$(NC%)
        '
      End If
    End If
    Screen.MousePointer = 1
    '
End Sub
   '
Function NUPRESUP&()
   '
   NPP& = 1
   '
   FIN& = ULTREG&("PRESUPU")
   If FIN& > 0 Then
      For U& = FIN& To FIN& - 128 Step -1
         If U& < 1 Then GoTo 99
         XX$ = REGLEIDO$("PRESUPU", U&)
         NPPX& = CVS(Mid$(XX$, 1, 4))
         If NPPX& >= NPP& Then
            NPP& = 1 + NPPX&
         End If
      Next U&
   End If
   '
99 NUPRESUP& = NPP&
   '
End Function
'
Sub CARESPECI()
      '
    CODD$ = Text13.TEXT
    CI% = CODINT%(CODD$)
    If CI% < 1 Then Exit Sub
    '
    Text28.TEXT = "": TTYZ$ = ""
    If CI% > 0 Then
       If CI% <= NUMAS% Then
          Screen.MousePointer = 11
          RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI%))
          URN& = ULTREG&("ESPECIFI")
          '
          For U& = 1 To Len(RNC$) Step 4
             RENOTA& = CVS(Mid$(RNC$, U&, 4))
             If RENOTA& > 0 Then
                If RENOTA& <= URN& Then
                   TTYY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
                   If TTYZ$ <> "" Then
                      TTYZ$ = TTYZ$ + Chr$(13) + Chr$(10)
                   End If
                   TTYZ$ = TTYZ$ + RTrim$(TTYY$)
                End If
             End If
          Next U&
          Screen.MousePointer = 1
       End If
    End If
    LOCADE& = Len(TTYZ$)
    For U& = LOCADE& To 1 Step -1
      If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
    Next U&
    U& = 0
30  Text28.TEXT = Left$(TTYZ$, U&)
    Text28.Refresh
    '
End Sub
'


Private Sub Text6_GotFocus()
   Text1.SetFocus
End Sub

Private Sub Text7_GotFocus()
   Text1.SetFocus
End Sub

Private Sub Text8_GotFocus()
   'Text1.SetFocus
End Sub

Private Sub Text9_GotFocus()
   Command9.SetFocus
End Sub

Private Sub TICAMBIS_Change()
   If TICAMBIS.Enabled = True Then
     VH(MONEMI%) = Val(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(Val(TICAMBIS), "F9.4")
End Sub

Private Sub Timer1_Timer()
   '
   Call CALTOPRES
   '
End Sub

Private Sub Total_Change(Index As Integer)
   If Index = 0 Then
      DESC$ = PORDESCU.TEXT
      TTT1 = Val(Total(0).TEXT) * COEFDESCU(DESC$) / 100
      Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
      TTT3 = Val(Total(0).TEXT) - Val(Total(1).TEXT)
      Total(3).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
      CANTI = Val(Text14.TEXT)
      If CANTI < 0.1 Then CANTI = 1
      SUBTOT = Val(Total(0))
        PREUNI = Val(Total(0)) / CANTI ' CALCULA EL PRECIO UNITARIO
        If Val(Total(0)) >= 0.1 Then
          Label14 = FORMATNUM$(PREUNI, "F12.2") ' GUARDA EL ULTIMO P.UNITARIO
        End If
    ElseIf Index = 3 Then
      Call CALTOIVA
   End If
End Sub
'
Sub PRIPRESU(VIPREVIA%)
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORPRESU"))
    FPXX$ = "FORPRESU"
    If Picture1.Visible = True Then
       FORIPRE$ = TRIM$(CONTROL$("", "FORPRERE"))
       FPXX$ = "FORPRERE"
    End If

    If FORIPRE$ = "" Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto.\Falta Definir Formulario Electrónico.")
       Exit Sub
    End If
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
10  NC% = Val(Text1.TEXT)
    TIPODOC$ = "Presupuesto"
    '\\\OT-5-851-FG-26/12/05///
    If PIVACLI%(NC%) = 5 Then
      TIPODOC$ = "F. Proforma"
    End If
    '\\\OT-5-851-FG-FIN///
    DESTIDOC% = NC%
    NUMEDOC$ = TRIM$(Text11.TEXT)
      If VIPREVIA% = 1 Then NUMEDOC$ = "(VP-" + NUMEDOC$ + ")"
    FECHDOC$ = TRIM$(Text9.TEXT)
    '
    CODPARAM$(1) = "REF-MAT3": DOCPARAM$(1) = TRIM$(Text10.TEXT)
    CODPARAM$(2) = "REF-MAT2": DOCPARAM$(2) = TRIM$(Text12.TEXT)
    CODPARAM$(3) = "PROV-CLI": DOCPARAM$(3) = NPROCLI$(NC%)
    CODPARAM$(4) = "COD-MAT": DOCPARAM$(4) = TRIM$(Text13.TEXT)
    CODPARAM$(5) = "DES-MAT": DOCPARAM$(5) = TRIM$(Label8.Caption)
    CODPARAM$(6) = "IMP-NETO": DOCPARAM$(6) = Total(0).TEXT
    CODPARAM$(7) = "POR-DESC": DOCPARAM$(7) = TRIM$(PORDESCU.TEXT)
    CODPARAM$(8) = "IMP-DESC": DOCPARAM$(8) = Total(1).TEXT
    CODPARAM$(9) = "IMPNEDES": DOCPARAM$(9) = Total(3).TEXT
    CODPARAM$(10) = "FECH-LAN": DOCPARAM$(10) = TRIM$(Total(2).TEXT)
    CODPARAM$(11) = "COND-PAG": DOCPARAM$(11) = TRIM$(Eco21(1).TEXT)
    CODPARAM$(12) = "DESTINO": DOCPARAM$(12) = TRIM$(Total(4).TEXT)
    CODPARAM$(13) = "FECH-VEN": DOCPARAM$(13) = TRIM$(Total(6).TEXT)
    CODPARAM$(14) = "NRO-LOTE": DOCPARAM$(14) = TRIM$(Total(5).TEXT)
    CODPARAM$(15) = "VENDEDOR": DOCPARAM$(15) = TRIM$(Eco21(2).TEXT)
    CODPARAM$(16) = "EQUIPO1": DOCPARAM$(16) = TRIM$(Text8.TEXT)
    CODPARAM$(17) = "MONEDEMI": DOCPARAM$(17) = TRIM$(DEMONECO$(MONEMI%))
    CODPARAM$(18) = "TICAMBIO": DOCPARAM$(18) = TRIM$(TICAMBIS)
    CODPARAM$(19) = "IMP-IVA": DOCPARAM$(19) = TRIM$(TOTIVA)
    CODPARAM$(20) = "PERC-IVA": DOCPARAM$(20) = TRIM$(TOTPER)
    CODPARAM$(21) = "IMP-TOTA": DOCPARAM$(21) = TRIM$(FORMATNUM$((Val(Total(3)) + Val(TOTIVA)), "F12.2"))
    CODPARAM$(22) = "CAN-CONJ": DOCPARAM$(22) = TRIM$(Text14.TEXT)
    CAPARDOC% = 22
    '
    CODTABU1$(1) = "REF-MAT1"
    CACOLTAB1% = 1
    '
    Call FRATEXTO(Text28.TEXT, 62)
    CAITAB1% = 0
    '
    For YX& = 1 To CARETEX%
       CAITAB1% = CAITAB1% + 1
       TETABU1$(1, CAITAB1%) = "." + RETEX$(YX&)
    Next YX&
    '
    If Picture1.Visible = True Then
      CODPARAM$(4) = "": DOCPARAM$(4) = ""
      CODPARAM$(5) = "": DOCPARAM$(5) = ""
      '
      CODTABU1$(1) = "CANTIDAD"
      CODTABU1$(2) = "COD-MAT"
      CODTABU1$(3) = "DES-MAT"
      CODTABU1$(4) = "PRE-LIST"
      CODTABU1$(5) = "NET-ITEM"
      CACOLTAB1% = 5
      '
      CAITAB1% = 0
      For U& = 1 To ULTREG&("!CARDETPR")
        X$ = REGLEIDO$("!CARDETPR", U&)
        CIIX% = CVI(Mid$(X$, 1, 2))
        CAN = CVS(Mid$(X$, 47, 4))
        INETO = CVS(Mid$(X$, 39, 4)) * (100 - CVS(Mid$(X$, 69, 4))) / 100
        If CAN > 0.05 Then
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = Str$(CAN)
          TETABU1$(2, CAITAB1%) = CODEXT$(CIIX%)
          TETABU1$(3, CAITAB1%) = DESCRIT0$(CIIX%)
          TETABU1$(4, CAITAB1%) = FORMATNUM$(INETO, "F12.4")
          TETABU1$(5, CAITAB1%) = FORMATNUM$(INETO * CAN, "F12.2")
        End If
      Next U&
    End If
    '
    NOCONFIRM% = 1
    Call PRINTDOC(FPXX$)
    '
    If Picture1.Visible = False Then
        CIXI% = CODINT%(Text13.TEXT)
        DOCUORIG$ = "Presupuesto " + NUMEDOC$
        TIPODOC$ = "Anexo Presup."
        For KKI% = 0 To 3
          If ANEXPRESU.Check1(KKI%).Value > 0 Then
            Call PRIDOGRAF(CIXI%, KKI%)
          End If
        Next KKI%
      Else
        For U& = 1 To ULTREG&("!CARDETPR")
          X$ = REGLEIDO$("!CARDETPR", U&)
          CIXI% = CVI(Mid$(X$, 1, 2))
          DOCUORIG$ = "Presupuesto " + NUMEDOC$
          TIPODOC$ = "Anexo Presup."
          For KKI% = 0 To 3
            If ANEXPRESU.Check1(KKI%).Value > 0 Then
              Call PRIDOGRAF(CIXI%, KKI%)
            End If
          Next KKI%
        Next U&
    End If
    '
    If CANCELPRINT% < 1 Then
      If COMALTER%(DOCUORIG$ + " Registrado y Emitido por Impresora\\Repetir Impresión\Continuar") = 1 Then
        GoTo 10
      End If
    End If
    '
99 End Sub
   '
Private Sub Total_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
   If Index = 1 Or Index = 3 Then
     Total(0).SetFocus
   End If
   If Index = 0 Then
     CII% = CODINT%(Text13.TEXT)
     If CII% < 1 Or CII% > NUMAS% Then
       Text13.SetFocus
     End If
   End If
End Sub

Sub GRAPRESU(OKX%)
   '
   OKX% = 0
   Call BLOQARCH("PRESUPU")
   '
   NPP& = NUPRESUP&
   If NPP& > Val(Text11.TEXT) Then
      Call CIERRARCH("PRESUPU")
      Call COMUNI("Número Ocupado")
      Text11.TEXT = TRIM$(Str$(NPP&))
      Exit Sub
   End If
   '
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = Val(Text1.TEXT)
   '
   If Picture1.Visible = False Then
     CIIX% = CODINT%(Text13.TEXT)
     If CIIX% < 1 Or CIIX% > NUMAS% Then
       MERRO$ = "Código de Artículo no Válido"
       GoTo 98
     End If
   End If
   '
   If Val(Total(0).TEXT) < 1 Then
      MERRO$ = "Falta(n) Precio(s) de Lista"
      GoTo 98
   End If
   '
   If Val(Total(3).TEXT) < 1 Then
      MERRO$ = "Importe no Válido"
      GoTo 98
   End If
   '
   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   '\\\OT-05-0649-WAR-30/09/05///
   NPRESU& = Val(Text11.TEXT)
   Call GRAPRESU1(NPRESU&)
   '\\\OT-05-0649-WAR-FIN///
   '
   CANTI0 = Val(Text14): If CANTI0 < 1 Then CANTI0 = 1
   XX$ = REGBLAN$("PRESUPU")
   Call REPLA(XX$, MKS$(Val(Text11.TEXT)), 1, 4)
   Call REPLA(XX$, MKI$(FEMI%), 5, 2)
   Call REPLA(XX$, MKI$(NCLIE%), 7, 2)
   Call REPLA(XX$, TRIM$(Text10.TEXT), 9, 16)
   Call REPLA(XX$, TRIM$(Text12.TEXT), 25, 40)
   Call REPLA(XX$, Chr$(Val(TEXT21(0).TEXT)), 65, 1)
   Call REPLA(XX$, Chr$(Val(TEXT21(1).TEXT)), 66, 1)
   Call REPLA(XX$, Chr$(Val(TEXT21(2).TEXT)), 67, 1)
   Call REPLA(XX$, Chr$(MONEMI%), 68, 1)
   Call REPLA(XX$, MKI$(CIIX%), 69, 2)
   Call REPLA(XX$, MKS$(Val(Total(0).TEXT) / CANTI0), 71, 4)
   Call REPLA(XX$, TRIM$(PORDESCU.TEXT), 75, 14)
      DESC$ = PORDESCU.TEXT
      TTT1 = COEFDESCU(DESC$)
   Call REPLA(XX$, MKS$(TTT1), 89, 4)
   Call REPLA(XX$, MKS$(Val(Total(1).TEXT) / CANTI0), 93, 4)
   Call REPLA(XX$, MKS$(Val(Total(3).TEXT) / CANTI0), 97, 4)
   Call REPLA(XX$, MKS$(CANTI0), 101, 4)
   Call REPLA(XX$, MKD$(CDbl(1 * Val(Total(3).TEXT))), 105, 8)
   '
   TCM = 1
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") < 1 Then ' lo guarda en pesos
     TCM = TICAMONE(MONEMI%)
     Call REPLA(XX$, Chr$(1), 68, 1)
     Call REPLA(XX$, MKS$(TCM * Val(Total(0).TEXT)), 71, 4)
     Call REPLA(XX$, MKS$(TCM * Val(Total(1).TEXT)), 93, 4)
     Call REPLA(XX$, MKS$(TCM * Val(Total(3).TEXT)), 97, 4)
     Call REPLA(XX$, MKD$(CDbl(1 * TCM * Val(Total(3).TEXT))), 105, 8)
   End If
   '
   If Picture1.Visible = False Then
       Call REGAPP("PRESUPU", XX$)
     Else
       For U& = 1 To ULTREG&("!CARDETPR")
         X$ = REGLEIDO$("!CARDETPR", U&)
         CIIX% = CVI(Mid$(X$, 1, 2))
         CAN = CVS(Mid$(X$, 47, 4))
         INETO = CVS(Mid$(X$, 39, 4)) * (100 - CVS(Mid$(X$, 69, 4))) / 100
         If CAN > 0.05 Then
           Call REPLA(XX$, MKI$(CIIX%), 69, 2)
           Call REPLA(XX$, MKS$(TCM * INETO), 71, 4) ' !!!!!!
           Call REPLA(XX$, MKS$(TCM * INETO * TTT1 / 100), 93, 4)
           Call REPLA(XX$, MKS$(TCM * INETO * (1 - TTT1 / 100)), 97, 4)
           Call REPLA(XX$, MKS$(CAN), 101, 4)
           Call REPLA(XX$, MKD$(CDbl(CAN * TCM * INETO * (1 - TTT1 / 100))), 105, 8)
           Call REGAPP("PRESUPU", XX$)
         End If
       Next U&
   End If
   '
   Call CIERRARCH("PRESUPU")
   '
   FEX = FEMI%
   YY$ = MKS$(NPP&) + TRIM$(RASOCLI$(NCLIE%)) + " - "
   If Picture1.Visible = False Then
        YY$ = YY$ + TRIM$(CODEXT$(CIIX%))
      Else
        YY$ = YY$ + "VARIOS"
   End If
   YY$ = YY$ + " - " + FECHATEX$(FEX)
   Call REGAPP("INDIPRES", YY$)
   Call CIERRARCH("INDIPRES")
   '
   OKX% = 1
   Exit Sub
   '
98 Call CIERRARCH("PRESUPU")
   Call CIERRARCH("INDIPRES")
   Call MENSERR(24, MERRO$)
   '
End Sub

Private Sub Total_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Index = 0 Then
       PORDESCU.SetFocus
     End If
   End If
End Sub

Private Sub Total_LostFocus(Index As Integer)
   If Index = 0 Then
     TTT1 = Val(Total(0).TEXT)
     Total(0) = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   End If
End Sub

Sub CARDETPRES()
   '
   HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC/ARCHECO(1)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle Presupuesto (Repuestos)"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   '
   Timer1.Enabled = True
10 PRF$ = "CARPRESR"
   XXX% = VENTABU%(PRF$ + ATRI$)
   Call CALTOPRES
   Timer1.Enabled = False
   '
   If XXX% < 0 Then
      Call CIERRARCH("!CARDETPR")
      Call BLANARCH("!CARDETPR")
      Total(0).TEXT = ""
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
20 If ULTREG&("!CARDETPR") < 1 Then
      Screen.MousePointer = 1
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   Screen.MousePointer = 11   ' Registrando Pedido ...
   FIN& = ULTREG&("!CARDETPR")
   If FIN& > 0 Then
     Y$ = REGLEIDO$("!CARDETPR", FIN&)
     CI% = CVI(Left$(Y$, 2))
     If CI% < 1 Or CI% > NUMAS% Then
       Call SETULTREG("!CARDETPR", FIN& - 1)
       Call CIERRARCH("!CARDETPR")
       GoTo 20
     End If
   End If
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   For U& = 1 To FIN&
     Y$ = REGLEIDO$("!CARDETPR", U&)
     CI% = CVI(Left$(Y$, 2))
     If CI% < 1 Or CI% > NUMAS% Then
       Call MENSERR(24, "Código no Válido")
       Call REPLA(Y$, MKI$(0), 1, 2)
       Call GRAREG("!CARDETPR", Y$, U&)
       GoTo 10
     End If
     '
     FENTRESOL% = CVI(Mid$(Y$, 63, 2))
     If FENTRESOL% < HOII% Then
        FENTRESOL% = HOII%
        Call REPLA(Y$, MKI$(FENTRESOL%), 63, 2)
        Call GRAREG("!CARDETPR", Y$, U&)
     End If
     '
     ' para DOSIVAC no verifica codigos repetidos
     If InStr(EPAC$, "DOSIVA") < 1 Then
       For W& = 1 To U& - 1
         Y$ = REGLEIDO$("!CARDETPR", W&)
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
     X$ = REGLEIDO$("!CARDETPR", U&)
     CI% = CVI(Left$(X$, 2))
     TTXX$ = Space$(96)
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     '
     CAN = CVS(Mid$(X$, 47, 4))
     If CAN > 999999 Then
       Call MENSERR(24, "Cantidad Excesiva - Max 999999")
       Call REPLA(X$, MKS$(0), 47, 4)
       Call GRAREG("!CARDETPR", X$, U&)
       GoTo 10
     End If
     If CAN < 0.5 Then
       Call MENSERR(24, "Falta Cantidad")
       Call REPLA(X$, MKS$(0), 47, 4)
       Call GRAREG("!CARDETPR", X$, U&)
       GoTo 10
     End If
     '
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 99.9 Then
       Call MENSERR(24, "Descuento no Válido")
       Call REPLA(X$, MKS$(0), 69, 4)
       Call GRAREG("!CARDETPR", X$, U&)
       GoTo 10
     End If
     '
     If ECOPRECIO% < 0 Then
       PRECARGA = PRELISTA#(ARCHIPRE$, CI%)
       Call REPLA(X$, MKS$(PRECARGA), 39, 4)
       Call GRAREG("!CARDETPR", X$, U&)
     End If
     PRECARGA = CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     If PRECARGA < 0.005 Or PRECARGA > 99999.99 Then
       If modo% < 2 Then
         If ECOPRECIO% >= 0 Then
           Call MENSERR(24, "Artículo sin Precio")
           Call REPLA(X$, MKS$(0), 39, 4)
           Call GRAREG("!CARDETPR", X$, U&)
           modo% = 2
           GoTo 10
         End If
       End If
     End If
     '
     Call REPLA(X$, MKS$(PRECARGA), 39, 4)
     Call REPLA(X$, MKS$(PRECARGA), 51, 4)
     Call REPLA(X$, MKD$(TOT#), 55, 8)
     Call REPLA(X$, MKS$(PRECARGA), 65, 4)
     Call REPLA(X$, MKS$(PREUNI), 73, 4)
     Call REPLA(X$, MKD$(TOT#), 77, 8)
     Call GRAREG("!CARDETPR", X$, U&)
     '
   Next U&
   '
   Call CARLISPRES
   Call CIERRARCH("!CARDETPR")
   Call CALTOPRES
   '
   List1.Refresh
   On Error Resume Next
   Command14.SetFocus
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub

Sub CARLISPRES()
   '
   List1.Clear
   FIN& = ULTREG&("!CARDETPR")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPR", U&)
     TTXX$ = Space$(96)
     CI% = CVI(Left$(X$, 2))
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
     Call REPLA(TTXX$, CODEXT$(CI%), 8, 15)
     FEX = FENTRESOL%
     Call REPLA(TTXX$, DESCRIT0$(CI%), 24, 28)
     PRELIX = PRELISTA#(ARCHIPRE$, CI%)
     If PRELIX < PRECARGA Then PRELIX = PRECARGA
     If PRELIX < PREUNI Then PRELIX = PREUNI
     PORDESCUIN = 0: If PRELIX > 0.005 Then PORDESCUIN = 100 * (PRELIX - PREUNI) / PRELIX
     Call REPLA(TTXX$, CSTRING$(MKS$(PRELIX), 3, 8, 2, 2), 52, 8)
     Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 61, 5)
     Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, 2, 2), 66, 8)
     Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 2, 2), 76, 10)
     List1.AddItem TTXX$
     '
   Next U&
   '
End Sub
'
Sub CALTOPRES()
   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = Val(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PVCLI% = PIVACLI%(NCLIE%)
       If PVCLI% = 2 Then RNI% = 1
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   SUTOTPED# = 0: SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
   For U& = 1 To ULTREG&("!CARDETPR")
     X$ = REGLEIDO$("!CARDETPR", U&)
     CIXXI% = CVI(Left$(X$, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     If PORDESCUEIN < 0 Or PORDESCUEIN > 99 Then PORDESCUEIN = 0
     If CAN > 0.05 Then
       PRELI = CVS(Mid$(X$, 39, 4))
       'If PRELI < 0.01 Then
       '  PRELI = PRELISTA(ARCHIPRE$, CIXXI%)
       'End If
       PREUNI = PRELI * (100 - PORDESCUIN) / 100
       TOT# = CDbl(CAN * PREUNI)
       SUTOTPED# = SUTOTPED# + TOT#
       SUIVA# = SUIVA# + (Int(TOT# * PORCEIVA(CIXXI%) + 0.5)) / 100
       Call REPLA(X$, MKS$(PREUNI) + MKD$(TOT#), 51, 12)
       Call GRAREG("!CARDETPR", X$, U&)
     End If
   Next U&
   '
   Total(0).TEXT = CSTRING$(MKD$(SUTOTPED#), 3, 11, 2, 2)
   Total(0).Refresh
   '
   DESC$ = TRIM$(PORDESCU.TEXT)
   TTT1 = Val(Total(0).TEXT) * COEFDESCU(DESC$) / 100
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   '
   TTT3 = Val(Total(0).TEXT) - Val(Total(1).TEXT)
End Sub

Sub CALTOIVA()
   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = Val(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PVCLI% = PIVACLI%(NCLIE%)
       If PVCLI% = 2 Then RNI% = 1
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   SUTOTPED# = Val(Total(3))
   SUIVA# = APLIVA% * (Int(SUTOTPED# * ALIVA + 0.5)) / 100
   '
   If TRIM$(Text13.TEXT) <> "" Then
    CIXX% = CODINT%(TRIM$(Text13.TEXT))
    If CIXX% > 0 Then
      ALIVACOD = PORCEIVA(CIXX%)
      SUIVA# = APLIVA% * (Int(SUTOTPED# * ALIVACOD + 0.5)) / 100
    End If
   End If
   '
   SUIVARNI# = SUIVA# * RNI% / 2
   TOTIVA = CSTRING$(MKD$(SUIVA# + SUIVARNI#), 4, 12, 2, 2)
   '
   Call CALPERIB
   '
End Sub

Sub CALPERIB()
   '
   Static CTTXX%, TASAA, TASAB, TASAEXP, MINIPER
   '
   If CTTXX% < 1 Then
     TASAA = Val(CONTROL$("PERIBRU", "TASA-A"))
     TASAB = Val(CONTROL$("PERIBRU", "TASA-B"))
     TASAEXP = Val(CONTROL$("PERIBRU", "TASA-EXP"))
     MINIPER = Val(CONTROL$("PERIBRU", "MINIPER"))
     CTTXX% = 1
   End If
   '
   TASAX = TASAA
   NCLIE% = Val(Text1)
   If NCLIE% < 1 Then Exit Sub
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
       BIMPON# = Val(Total(3))
       If VARIDC% = 2 Then BIMPON# = BIMPON# + Val(TOTIVA)
       If BIMPON# * TICAMONE(MONEMI%) >= MINIPER Then
         PERSALTA# = (BIMPON# * PORCONVE / 100) * (TASAX / 100)
         TOTPER = FORMATNUM$(PERSALTA#, "F11.2")
       End If
     End If
   End If
   '
End Sub
'\\\OT-05-0649-WAR-30/09/05///
Sub PREPORCOP(NPRESU&, NCLIE%)
  '
  X$ = RECFILT$("PRESUPU", 1, MKS$(NPRESU&))
  NREG& = CVS(X$)
  XX$ = REGLEIDO$("PRESUPU", NREG&)
  'DATOS DEL PRESUPUESTO
  '\\\OT-05-0758-WAR-18/11/05///
  NUCLI% = NCLIE% 'CVI(Mid$(XX$, 7, 2))'CLIENTE QUE SE ELIGIO
  '\\\OT-05-0758-WAR-FIN///
  REFE$ = Mid$(XX$, 9, 16)
  ATESR$ = Mid$(XX$, 25, 40)
  LIPREC% = Asc(Mid$(XX$, 65, 1))
  COPA% = Asc(Mid$(XX$, 66, 1))
  VENDE1% = Asc(Mid$(XX$, 67, 1))
  MONEPRE% = Asc(Mid$(XX$, 68, 1))
  CODI1% = CVI(Mid$(XX$, 69, 2))
  PORDESCU1 = CVS(Mid$(XX$, 89, 4))
  PREUNI1 = CVS(Mid$(XX$, 71, 4))
  CANTI0 = CVS(Mid$(XX$, 101, 4))
  'CARGA DE ENCABEZADO DE PRESUPUESTO
  Text1 = TRIM$(Str$(NUCLI%))
  Call CARDACLI
  TEXT21(0) = TRIM$(Str$(LIPREC%))
  TEXT21(1) = TRIM$(Str$(COPA%))
  TEXT21(2) = TRIM$(Str$(VENDE1%))
  Text10 = REFE$
  Text12 = ATESR$
  'CARGO MONEDA DE ORIGEN
  If MONEPRE% = 2 Then
       MONEMI% = 2
       MONEMIS = "Dólares"
    Else
       MONEMI% = 1
       MONEMIS = "Pesos"
  End If
  'CARGO COMPLEMENTOS DEL PRESUPUESTO SI LOS TIENE
  Z$ = RECFILT$("PRESUPU1", 1, MKS$(NPRESU&))
  NREG1& = CVS(Z$)
  If NREG1& >= 1 Then
    ZZ$ = REGLEIDO$("PRESUPU1", NREG1&)
    Total(2) = Mid$(ZZ$, 5, 48)
    Total(4) = Mid$(ZZ$, 53, 48)
    Total(6) = Mid$(ZZ$, 101, 32)
    Total(5) = Mid$(ZZ$, 133, 32)
  End If
  'VERIFICO SI ES ITEM UNICO O LISTA DE REPUESTOS
  Call FILTERFILE("PRESUPU", "!PRESUPU", 1, MKS$(NPRESU&))
  NUIT1& = ULTREG("!PRESUPU")
  If NUIT1& = 1 Then
      'ITEM UNICO
      Label12 = "LISTA REPUESTOS"
      Call Label12_DblClick
      Text13 = CODEXT$(CODI1%)
      Text14 = FORMATNUM$(CANTI0, "F11.1")
      Total(0) = FORMATNUM$(PREUNI1 * CANTI0, "F11.2")
      PORDESCU = FORMATNUM$(PORDESCU1, "F5.2")
    Else
      'LISTA DE REPUESTOS
      Label12 = "ITEM UNICO"
      Call Label12_DblClick
      '
      XX$ = REGBLAN$("CARDETPR")
      JJ& = 0
      Call BLANARCH("!CARDETPR")
      For J& = 1 To ULTREG&("!PRESUPU")
        Y$ = REGLEIDO$("!PRESUPU", J&)
        CODI1% = CVI(Mid$(Y$, 69, 2))
        PRELI1 = CVS(Mid$(Y$, 71, 4))
        IMPOUN = CVS(Mid$(Y$, 97, 4))
        CAN1 = CVS(Mid$(Y$, 101, 4))
        '
        Call REPLA(XX$, MKI$(CODI1%), 1, 2)
        Call REPLA(XX$, MKS$(PRELI1), 39, 4)
        Call REPLA(XX$, MKS$(IMPOUN), 65, 4)
        Call REPLA(XX$, MKS$(CAN1), 47, 4)
        JJ& = JJ& + 1
        Call GRAREG("!CARDETPR", XX$, JJ&)
      Next J&
      Call SETULTREG("!CARDETPR", JJ&)
      '
      Call CARDETPRES
  End If
  ALPORCOPIA% = 1 'MARCA PARA SABER SI ESTA EN ALTA POR COPIA
  '
End Sub
Sub GRAPRESU1(NPRESU&) 'GRABO ARCHIVO CON COMPLEMENTOS DEL PRESUPUESTO
  '
  PZO$ = Total(2)
  LUG$ = Total(4)
  VALI$ = Total(6)
  VTO$ = Total(5)
  '
  X$ = REGBLAN$("PRESUPU1")
  Call REPLA(X$, MKS$(NPRESU&), 1, 4)
  Call REPLA(X$, PZO$, 5, 48)
  Call REPLA(X$, LUG$, 53, 48)
  Call REPLA(X$, VALI$, 101, 32)
  Call REPLA(X$, VTO$, 133, 32)
  Call REGAPP("PRESUPU1", X$)
  Call CIERRARCH("PRESUPU1")
  '
End Sub
'\\\OT-05-0649-WAR-FIN///


