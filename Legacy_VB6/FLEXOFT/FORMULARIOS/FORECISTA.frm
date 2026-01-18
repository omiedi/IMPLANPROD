VERSION 5.00
Begin VB.Form FORECISTA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0A0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibo de Cobranza en Cuenta Corriente"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11910
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0A0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8520
      Left            =   0
      TabIndex        =   7
      Top             =   -150
      Width           =   12000
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0A0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1335
         Left            =   3960
         TabIndex        =   65
         Top             =   7080
         Width           =   2560
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Aplicados:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -1080
            TabIndex        =   71
            Top             =   120
            Width           =   2220
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   810
            TabIndex        =   70
            Top             =   420
            Width           =   1695
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   810
            TabIndex        =   69
            Top             =   900
            Width           =   1695
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   495
            TabIndex        =   68
            Top             =   375
            Width           =   330
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "u$s"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   435
            Left            =   75
            TabIndex        =   67
            Top             =   855
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   2800
         Left            =   210
         ScaleHeight     =   2745
         ScaleWidth      =   10395
         TabIndex        =   22
         Top             =   4250
         Width           =   10450
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   10515
            Begin VB.Line Line1 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line2 
               X1              =   6720
               X2              =   6720
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   " Documento             Nro      Emision                Importe       Vencim.                  Saldo                Acredit.       "
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
               TabIndex        =   29
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line8 
               X1              =   1680
               X2              =   1680
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line9 
               X1              =   2520
               X2              =   2520
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line10 
               X1              =   3730
               X2              =   3730
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line11 
               X1              =   5460
               X2              =   5460
               Y1              =   0
               Y2              =   1050
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5630
            Left            =   0
            TabIndex        =   23
            Top             =   315
            Width           =   10515
            Begin VB.ListBox List2 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2085
               Left            =   -30
               TabIndex        =   24
               Top             =   0
               Width           =   10450
            End
            Begin VB.ListBox List3 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2085
               Left            =   -30
               TabIndex        =   43
               Top             =   0
               Width           =   10450
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   8295
               TabIndex        =   31
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   0
               TabIndex        =   30
               Top             =   0
               Width           =   1695
            End
            Begin VB.Line Line4 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   2415
            End
            Begin VB.Line Line12 
               X1              =   1680
               X2              =   1680
               Y1              =   1575
               Y2              =   3360
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   6615
               TabIndex        =   27
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3570
               TabIndex        =   26
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Total General"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   25
               TabIndex        =   25
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Line Line14 
               BorderWidth     =   2
               X1              =   0
               X2              =   10395
               Y1              =   2085
               Y2              =   2085
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   10450
               Y1              =   2415
               Y2              =   2415
            End
            Begin VB.Line Line16 
               X1              =   2520
               X2              =   2520
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line17 
               X1              =   3730
               X2              =   3730
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line18 
               X1              =   5460
               X2              =   5460
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line19 
               X1              =   6720
               X2              =   6720
               Y1              =   2100
               Y2              =   2415
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0C0A0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1335
         Left            =   2640
         TabIndex        =   66
         Top             =   7030
         Visible         =   0   'False
         Width           =   2600
         Begin VB.TextBox DESCUEN 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   840
            MaxLength       =   14
            TabIndex        =   73
            Top             =   480
            Width           =   1695
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
            Height          =   285
            Left            =   1260
            TabIndex        =   72
            Top             =   120
            Width           =   645
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "u$s"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   435
            Left            =   105
            TabIndex        =   77
            Top             =   915
            Width           =   855
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   525
            TabIndex        =   76
            Top             =   435
            Width           =   330
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento:             %"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   75
            Top             =   180
            Width           =   2220
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   840
            TabIndex        =   74
            Top             =   960
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Imprimir"
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
         Height          =   1250
         Left            =   10920
         Picture         =   "FORECISTA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Aprueba - Graba e Imprime Recibo"
         Top             =   6600
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   10920
         Picture         =   "FORECISTA.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Acceso a Calculadora 'CALC.EXE'"
         Top             =   6600
         Width           =   855
      End
      Begin VB.TextBox TICAMBIS 
         Enabled         =   0   'False
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
         Left            =   10700
         MaxLength       =   14
         TabIndex        =   56
         Top             =   7980
         Width           =   1065
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Opción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   10920
         Picture         =   "FORECISTA.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Anular, Re-Imprimir, Imputación de Créditos, Tablas"
         Top             =   4250
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
         Picture         =   "FORECISTA.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Continúa Carga de Recibo."
         Top             =   735
         Width           =   855
      End
      Begin VB.CommandButton Command7 
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
         TabIndex        =   48
         Top             =   735
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   435
         Left            =   10670
         TabIndex        =   49
         Top             =   670
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
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
         Height          =   780
         Left            =   10920
         Picture         =   "FORECISTA.frx":0DC8
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   5070
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   2430
         Left            =   210
         ScaleHeight     =   2370
         ScaleWidth      =   10395
         TabIndex        =   32
         Top             =   1680
         Width           =   10450
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   10515
            Begin VB.Line Line23 
               X1              =   8505
               X2              =   8505
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line22 
               X1              =   4935
               X2              =   4935
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "  M-P    Medio de Pago      Numero              P/T       Cl.         Emision           Vencim.               Importe"
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
               TabIndex        =   34
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line20 
               X1              =   5670
               X2              =   5670
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line13 
               X1              =   4200
               X2              =   4200
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line7 
               X1              =   2625
               X2              =   2625
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line6 
               X1              =   630
               X2              =   630
               Y1              =   0
               Y2              =   1155
            End
            Begin VB.Line Line5 
               X1              =   7140
               X2              =   7140
               Y1              =   0
               Y2              =   1050
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0F0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5580
            Left            =   -315
            TabIndex        =   35
            Top             =   0
            Width           =   10800
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
               Height          =   2085
               Left            =   285
               TabIndex        =   36
               Top             =   315
               Width           =   10500
            End
            Begin VB.Line Line21 
               X1              =   1155
               X2              =   1155
               Y1              =   1470
               Y2              =   3255
            End
         End
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   6195
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   315
         Width           =   435
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
         Height          =   780
         Left            =   10920
         Picture         =   "FORECISTA.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
         Top             =   1680
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
         Height          =   780
         Left            =   10920
         Picture         =   "FORECISTA.frx":164C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   2520
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
         Height          =   780
         Left            =   10920
         Picture         =   "FORECISTA.frx":1956
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Notas de Clientes"
         Top             =   3330
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
         Picture         =   "FORECISTA.frx":1C60
         TabIndex        =   16
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   315
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
         Left            =   9030
         TabIndex        =   0
         Text            =   " "
         Top             =   1155
         Width           =   1650
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
         Left            =   9450
         TabIndex        =   3
         Text            =   " "
         Top             =   315
         Width           =   1065
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
         TabIndex        =   4
         Top             =   315
         Width           =   180
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
         TabIndex        =   12
         TabStop         =   0   'False
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
         TabIndex        =   11
         TabStop         =   0   'False
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
         TabIndex        =   10
         TabStop         =   0   'False
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
         TabIndex        =   9
         TabStop         =   0   'False
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
         TabIndex        =   2
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
         TabIndex        =   1
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label TOCRE 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   5250
         TabIndex        =   63
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RNICRE 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   5250
         TabIndex        =   62
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label IVACRE 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   5250
         TabIndex        =   61
         Top             =   420
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label NECRE 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   5250
         TabIndex        =   60
         Top             =   210
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label35 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7140
         TabIndex        =   59
         Top             =   315
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label34 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         Height          =   330
         Left            =   6825
         TabIndex        =   58
         Top             =   315
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "T-C:"
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
         Left            =   10470
         TabIndex        =   57
         ToolTipText     =   "Doble-Click para Ajustar T-C"
         Top             =   7750
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "u$s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   105
         TabIndex        =   55
         Top             =   7930
         Width           =   855
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "u$s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   7875
         TabIndex        =   54
         Top             =   7930
         Width           =   855
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8295
         TabIndex        =   53
         Top             =   7455
         Width           =   330
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   525
         TabIndex        =   52
         Top             =   7455
         Width           =   330
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   840
         TabIndex        =   51
         Top             =   7980
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   8610
         TabIndex        =   50
         Top             =   7980
         Width           =   1695
      End
      Begin VB.Line Line28 
         X1              =   210
         X2              =   10660
         Y1              =   4110
         Y2              =   4110
      End
      Begin VB.Line Line27 
         X1              =   10660
         X2              =   10660
         Y1              =   1680
         Y2              =   4095
      End
      Begin VB.Line Line26 
         X1              =   10660
         X2              =   10660
         Y1              =   4250
         Y2              =   7035
      End
      Begin VB.Line Line25 
         X1              =   10650
         X2              =   12000
         Y1              =   6500
         Y2              =   6500
      End
      Begin VB.Line Line24 
         X1              =   10650
         X2              =   12000
         Y1              =   5970
         Y2              =   5970
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   10320
         Picture         =   "FORECISTA.frx":1F6A
         Top             =   5985
         Width           =   2010
      End
      Begin VB.Label Label22 
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
         Left            =   8265
         TabIndex        =   47
         Top             =   735
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Talonario:"
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
         Left            =   6840
         TabIndex        =   46
         Top             =   750
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   8610
         TabIndex        =   42
         Top             =   7500
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Anticipos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6720
         TabIndex        =   41
         Top             =   7200
         Width           =   2220
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3780
         TabIndex        =   40
         Top             =   5775
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Valores:"
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
         Left            =   1470
         TabIndex        =   39
         Top             =   5880
         Width           =   2220
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   840
         TabIndex        =   38
         Top             =   7500
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valores:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -1050
         TabIndex        =   37
         Top             =   7200
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Documento no Válido como Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5355
         TabIndex        =   21
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   1050
         Width           =   2115
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero:"
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
         TabIndex        =   15
         Top             =   1170
         Width           =   1380
      End
      Begin VB.Label Label3 
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
         Left            =   8640
         TabIndex        =   14
         Top             =   325
         Width           =   750
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
         TabIndex        =   8
         Top             =   735
         Width           =   960
      End
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
Attribute VB_Name = "FORECISTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$, AUTONUME%, TICAMBX#
Dim ACCECALC$

Private Sub DESCUEN_Change(Index As Integer)
   Label20(0).Caption = CSTRING$(MKD$(Val(Label14(0)) + Val(DESCUEN(0)) - Val(Label18(0))), 4, 14, 2, 2)
   If Index = 0 Then Call CONDOLA
End Sub

Private Sub DESCUEN_LostFocus(Index As Integer)
   DESCUEN(Index) = FORMATNUM$(Val(DESCUEN(Index)), "F12.2")
End Sub

Private Sub Label14_Change(Index As Integer)
   If Index = 0 Then
      TT77$ = REPLACAR$(TRIM$(Text7), ",", ".")
      PORDEX = Val(Text7): If PORDEX > 99 Then PORDEX = 99
      If PORDEX < 0 Then PORDEX = 0
      IDES# = Val(Label14(0)) * (100 / (100 - PORDEX) - 1)
      DESCUEN(0) = TRIM$(CSTRING$(MKD$(IDES#), 4, 14, 2, 2))
      Call CONDOLA
   End If
End Sub

Private Sub Label18_Change(Index As Integer)
   If Index = 0 Then Call CONDOLA
End Sub

Private Sub Label20_Change(Index As Integer)
   If Index = 0 Then Call CONDOLA
End Sub

Private Sub Text7_Change()
   TT77$ = REPLACAR$(TRIM$(Text7), ",", ".")
   PORDEX = Val(Text7): If PORDEX > 99 Then PORDEX = 99
   If PORDEX < 0 Then PORDEX = 0
   IDES# = Val(Label14(0)) * (100 / (100 - PORDEX) - 1)
   DESCUEN(0) = CSTRING$(MKD$(IDES#), 4, 14, 2, 2)
   Call CONDOLA
End Sub

Private Sub ticambis_GotFocus()
   TICAMBIS.SelStart = 0
   TICAMBIS.SelLength = Len(TICAMBIS.TEXT)
End Sub

Private Sub Label29_DblClick()
   TICAMBIS.Enabled = True
   TICAMBIS.SetFocus
End Sub

Private Sub TICAMBIS_Change()
   If TICAMBIS.Enabled = True Then
     VH(MONEMI%) = Val(TICAMBIS)
   End If
   Call CONDOLA
End Sub

Private Sub TICAMBIS_LostFocus()
If Val(TICAMBIS) < 1 Then TICAMBIS = "1"
   TICAMBIS = FORMATNUM(Val(TICAMBIS), "F9.4")
   TICAMBIS.Enabled = False
End Sub

Private Sub BOTNEXT_Click()
    '
    Call GRACONTROL("RECISTA", "COBRACTI", TRIM$(Label22.Caption))
    Call SIGUERECI1
    Command5.Enabled = True
    '
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Call NOTACLI(NC%, 1)
    End If
    '
End Sub

Private Sub Command10_Click()
    On Error Resume Next
    Shell ACCECALC$
    On Error GoTo 0
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "AMACLI02"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "AMACLI01"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "AMACLI/NEW"
    Call CONECRUN(ACONEC$, "Alta de Nueva Cuenta de Cliente")
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command2_Click()
    If List1.ListCount > 0 Or List2.ListCount > 0 Then
          Call CIERRARCH("!CARMECO")
          Call SETULTREG("!CARMECO", 0)
          Call FRESHGRID
          Call BLANARCH("!CARMECO")
          Call CIERRARCH("!APLIRECI")
          Call BLANARCH("!APLIRECI")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          For KKI% = 0 To 1
            Label14(KKI%).Caption = ""
            Label18(KKI%).Caption = ""
            Label20(KKI%).Caption = ""
            DESCUEN(0) = ""
            Label33 = ""
          Next KKI%
          BOTNEXT.Enabled = True
          Command5.Enabled = False
          Text1.SetFocus
        Else
          If TRIM$(CONTROL$("RECISTA", "AVAHOJA")) = "NEWPAGE" Then
              'Printer.EndDoc
              Call SETSPOOL("ENDDOC", "")
          End If
          Call FINAL("")
    End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    'Call CONECRUN("*NOTACLI/" + NCLI$, "Notas de Clientes")
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then Call NOTACLI(NC%, 2)
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    'If EXISTE%("COGEVEN.EXE") < 1 Then
    '    Call OPNOIN("")
    '    Command4.Enabled = True
    '    Exit Sub
    'End If
    RESCUDEU.Text1.TEXT = Text1.TEXT
    PROLLAMA$ = Name
    'Hide
    RESCUDEU.Show
    RESCUDEU.Refresh
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    If NC% <> Val(Text1.TEXT) Then
       Call MENSERR(24, "Imposible Grabar - Intente de Nuevo")
       Call CIERRARCH("*.*")
       Call FINAL("")
    End If
    '
    TOTAVAL# = CDbl(Val(FORECISTA.Label14(0).Caption)) + CDbl(Val(FORECISTA.DESCUEN(0)))
    TOTACAN# = CDbl(Val(FORECISTA.Label18(0).Caption))
    If TOTACAN# > TOTAVAL# + 0.005 Then
        Call COMUNI("Operación Incorrecta.\  \Las Cancelaciones no Pueden\Exceder la Suma de Valores Recibidos.")
        Command5.Enabled = True
        Exit Sub
    End If
    '
    If Val(Label20(0)) >= 0.005 Then
      If Val(DESCUEN(0)) >= 0.005 Then
        Call MENSERR(24, "Para Generar Descuento\Debe Aplicarse la Totalidad del Pago.")
        Command5.Enabled = True
        Exit Sub
      End If
    End If
    '
    FECHALIM% = Val(CONTROL$("", "CIEMEVEN"))
    FF% = FECHANUM(Text9.TEXT)
    If FF% < 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
    '
    If FF% <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        FORECISTA.Text9.SetFocus
        Exit Sub
    End If
    '
    Call GRARECISTA(OKEY%)
    If OKEY% = 0 Then
        Command5.Enabled = True
        Exit Sub
      ElseIf OKEY% <> 6 Then
        Call COMUNI("La Registración de este Recibo\no Pudo Completarse. (" + TRIM$(Str$(OKEY%)) + ")\  \Consulte a su Soporte de Sistemas FLEXOFT.")
        Call FINAL("")
    End If
    Call PRIRECISTA
    '
    If Val(DESCUEN(0)) >= 0.005 Then
      Call PRICREDESC
    End If
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    NROPRO& = 0
    Call NURECIBO(NROPRO&)
    Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
    Text11.SelStart = 3
    Text11.SelLength = Len(Text11.TEXT) - 3
    '
    Call CIERRARCH("!CARMECO")
    Call SETULTREG("!CARMECO", 0)
    Call FRESHGRID
    Call BLANARCH("!CARMECO")
    Call CIERRARCH("!APLIRECI")
    Call BLANARCH("!APLIRECI")
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Text7 = ""
    DESCUEN(0) = ""
    Label34 = ""
    Label35 = ""
    NECRE = ""
    IVACRE = ""
    RNICRE = ""
    TOCRE = ""
    '
    For KKI% = 0 To 1
      Label14(KKI%).Caption = ""
      Label18(KKI%).Caption = ""
      Label20(KKI%).Caption = ""
      DESCUEN(0) = ""
    Next KKI%
    BOTNEXT.Enabled = True
    Command5.Enabled = False
    Text11.SetFocus
    '
    If CONTROL$("RECISTA", "CLIENTE") = "REFRESH" Then
      Text1 = ""
      Text2 = ""
      Text3 = ""
      Text4 = ""
      Text5 = ""
      Text1.SetFocus
    End If
    '
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    OPRECISTA.Show 1
    Call COMUNI("Re-Configuracion Completa.\Re-Inicie la Aplicacion.")
    Call FINAL("")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    If CTX% = 0 Then
       CTX% = 1
       If ULTREG&("COBRADOR") > 1 Then
          XXX$ = REGSEL$("COBRADOR")
          If Len(XXX$) < 64 Then
             Call FINAL("")
          End If
          Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
          NROPRO& = 0
          Call NURECIBO(NROPRO&)
          Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
          Text11.SelStart = 3
          Text11.SelLength = Len(Text11.TEXT) - 3
       End If
    End If
    '
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   RECICOMP.Show 1
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Form_Load()
    '
    'Show
    'Refresh
    VERANTER$ = "RECISTA"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    COMMA$ = Command$
    '
    PPX% = InStr(COMMA$, "#" + JOBID$ + "#")
    If PPX% > 0 Then
        COMMA$ = Left$(COMMA$, PPX% - 1)
    End If
    If COMMA$ <> "" Then
        Text1.TEXT = COMMA$
    End If
    Show
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If InStr(EPAC$, "DOSIVA") > 0 Then
       Height = 8700
       Label29.Visible = True
    End If
    Top = (Screen.Height - 300 - Height) / 2
    DoEvents
    
    TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
    If Val(TICAMBIS) < 1 Then TICAMBIS = "   1.0000"
    '
10  FORIPRE$ = TRIM$(CONTROL$("", "FORECIST"))
    'If FORIPRE$ = "" Then
    '    Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Recibo.")
    '    OPRECISTA.Show 1
    '    GoTo 10
    'End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo Standard de Cobranza:\  \'" + RMCC$ + "'.")
            OPRECISTA.Show 1
            GoTo 10
        End If
    End If
    '
    If CONTROL$("RECISTA", "DESCAJA") = "CREDESCU" Then
       Frame7.Left = 5280
       Frame8.Visible = True
       DoEvents
    End If
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    If ULTREG&("COBRADOR") > 1 Then
        Label21.Visible = True
        Label22.Visible = True
        Command7.Visible = True
        Label22.Caption = CONTROL$("RECISTA", "COBRACTI")
      Else
        Label21 = "Cobrador:"
        Label21.Visible = True
        Label22.Visible = True
    End If
    '
    AUTONUME% = 0
    If CONTROL$("", "NURECIST") = "AUTONUME" Then
      AUTONUME% = 1
    End If
    '
    ACCECALC$ = CONTROL(IDENTER$, "RUT-CALC")
    ACCECALC$ = ACCECALC$ + "\" + CONTROL(IDENTER$, "EXE-CALC")
    If EXISTE%(ACCECALC$) > 0 Then
       Command5.Top = 7250
       Command5.Height = 600
    End If
    '
    NROPRO& = 0
    Call NURECIBO(NROPRO&)
    Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
    Text11.SelStart = 3
    Text11.SelLength = Len(Text11.TEXT) - 3
    '
    Text1.TEXT = COMMA$
    '
    If AUTONUME% <> 0 Then
      Text1.TEXT = "     "
      Text1.SetFocus
    End If
    '
    Refresh
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BLOQEXE("*")
     Call LIBARCH("*.*")
     Reset
     Call BLOQEXE(App.EXEName)
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label2_DblClick()
    OPRECISTA.Show 1
End Sub

Private Sub Label22_DblClick()
    '
    If CTX% = 0 Then
       CTX% = 1
       If ULTREG&("COBRADOR") > 1 Then
          XXX$ = REGSEL$("COBRADOR")
          If Len(XXX$) < 64 Then
             Call FINAL("")
          End If
          Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
          NROPRO& = 0
          Call NURECIBO(NROPRO&)
          Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
          Text11.SelStart = 3
          Text11.SelLength = Len(Text11.TEXT) - 3
       End If
    End If
    '
End Sub

Private Sub List1_DblClick()
    '
    Call SIGUERECI2
    '
End Sub

Private Sub List2_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    Z$ = List2.TEXT
    LIST2ANT$ = Z$
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    ACREMAX# = CDbl(Val(FORECISTA.Label14(0).Caption)) + CDbl(Val(FORECISTA.DESCUEN(0))) - CDbl(Val(FORECISTA.Label18(0).Caption))
    If ACREMAX# < 0 Then ACREMAX# = 0
    IPEN$ = Mid$(Z$, 56, 14)
    If Val(IPEN$) > ACREMAX# Then
        IPEN$ = CSTRING$(MKD$(ACREMAX#), 3, 14, 2, 2)
    End If
    SUMACRE# = Val(Label18(0).Caption)
    If IACRE$ = "" Then
         Call REPLA(Z$, IPEN$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(Val(IPEN$))
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(Val(IACRE$))
    End If
    VACON% = 1
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = ILISTA&
    VACON% = 0
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    Label18(0).Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label18(0).Refresh
    '
    ANTI# = CDbl(Val(Label14(0).Caption)) + CDbl(Val(DESCUEN(0))) - SUMACRE#
    'If ANTI# < 0 Then
    '      FORECISTA.Label19.Caption = "Descuentos:"
    '    Else
    '      FORECISTA.Label19.Caption = "Anticipos:"
    'End If
    Label20(0).Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    Label20(0).Refresh
    '
End Sub
'

Private Sub List2_DblClick()
    '
    VACON% = 1
    ILIST2& = List2.ListIndex
    TOLIST2& = List2.TopIndex
    '
    Z$ = List2.TEXT
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = Mid$(Z$, 56, 14)
    SUMACRE# = Val(Label18(0).Caption)
    SUMACRE# = SUMACRE# - CDbl(Val(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(Val(Mid$(LIST2ANT$, 70, 14)))
    List2.Visible = False
    List2.AddItem LIST2ANT$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    '
    Call SIGUERECI3
    '
    VACON% = 1
    List2.TopIndex = TOLIST2&
    List2.Visible = True
    If List2.ListCount > ILIST2& Then
          List2.ListIndex = ILIST2&
        Else
          Call CIERRARCH("!CARMECO")
          Call SETULTREG("!CARMECO", 0)
          Call FRESHGRID
          Call BLANARCH("!CARMECO")
          Call CIERRARCH("!APLIRECI")
          Call BLANARCH("!APLIRECI")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          For KKI% = 0 To 1
            Label14(KKI%).Caption = ""
            Label18(KKI%).Caption = ""
            Label20(KKI%).Caption = ""
            DESCUEN(0) = ""
          Next KKI%
          BOTNEXT.Enabled = True
          Command5.Enabled = False
          Text1.SetFocus
    End If
    VACON% = 0
    '
End Sub

Private Sub Text1_Change()
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    For KKI% = 0 To 1
      Label14(KKI%).Caption = ""
      Label18(KKI%).Caption = ""
      Label20(KKI%).Caption = ""
      DESCUEN(0) = ""
    Next KKI%
    NCLIE% = Val(Text1)
    If ULTREG&("COBRADOR") < 2 Then
      Label22 = DEVENDE$(NCLIE%)
    End If
    BOTNEXT.Enabled = True
    Command5.Enabled = False
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Call NOTACLI(NC%, 1)
    End If
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
        FORECISTA.Refresh
        CTXX% = 1
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Val(Text1.TEXT) > 0 Then
          If Val(Text1.TEXT) <= 32767 Then
             NC% = Val(Text1.TEXT)
             If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
                Text2.TEXT = RASOCLI$(NC%)
                Text3.TEXT = DOMICLI$(NC%)
                Text4.TEXT = CPOSCLI$(NC%)
                Text5.TEXT = LOCACLI$(NC%)
                Call NOTACLI(NC%, 1)
                On Error Resume Next
                BOTNEXT.SetFocus
                On Error GoTo 0
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

Private Sub Text10_DblClick()
    Call SELECHO("ECODEP")
    DPX1% = Val(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Text11_GotFocus()
    If AUTONUME% <> 0 Then
      NROPRO& = 0
      Call NURECIBO(NROPRO&)
      Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
    End If
    '
    Text11.SelStart = 3
    On Error Resume Next
    Text11.SelLength = Len(Text11.TEXT) - 3
    On Error GoTo 0
End Sub

Private Sub Text11_LostFocus()
    '
    VDEFF$ = TRIM$(Text11.TEXT)
    Text11.TEXT = VDEFF$
    If Left$(VDEFF$, 3) <> "RB." Then GoTo 90
    NRONUE& = Val(Mid$(VDEFF$, 4))
    If NRONUE& > 999999 Then
        Call MENSERR(24, "Número de Recibo Fuera de Rango")
        Call NURECIBO(NROPRO&)
        GoTo 95
    End If
    '
    NROPRO& = NRONUE&
    Call NURECIBO(NROPRO&)
    If NROPRO& <> NRONUE& Then
        Beep
        Call MENSERR(24, "Número no Válido o Pre-existente")
        GoTo 95
    End If
    Exit Sub
    '
90  Beep
    If FORMALTER.Visible = True Then Exit Sub
    If OPRECISTA.Visible = True Then Exit Sub
    Call MENSERR(24, "Número no Válido")
    NROPRO& = 0
    Call NURECIBO(NROPRO&)
95  Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
    Text11.SelStart = 3
    Text11.SelLength = Len(Text11.TEXT) - 3
    Text11.SetFocus
    '
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
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

Private Sub Text6_Click()
    Text1.SetFocus
End Sub

Private Sub Text6_DblClick()
    OPRECISTA.Show 1
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Text9_LostFocus()
    VDEVV$ = TRIM$(Text9.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FF% = FECHANUM(VDEVV$)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          Text9.TEXT = HO$
          Text9.SetFocus
    End If
End Sub

Sub CONDOLA()
   '
   TICAMBX# = Val(TICAMBIS)
   If TICAMBX# < 1 Then TICAMBX# = 1
   Label14(1) = FORMATNUM$(Val(Label14(0)) / TICAMBX#, "F12.2")
   Label18(1) = FORMATNUM$(Val(Label18(0)) / TICAMBX#, "F12.2")
   Label20(1) = FORMATNUM$(Val(Label20(0)) / TICAMBX#, "F12.2")
   Label33 = FORMATNUM$(Val(DESCUEN(0)) / TICAMBX#, "F12.2")
End Sub

Sub PRICREDESC()
    '
    CANCELPRINT% = 0
    If CONTROL$("CREDESCU", "SUPRINT") = "SI" Then CANCELPRINT% = 1
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + Label34.Caption))
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Controlador de Impresión")
       Exit Sub
    End If
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Documento.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    FECHDOC$ = Text9.TEXT
    NUMEDOC$ = TRIM$(Label35)
    While Len(NUMEDOC$) < 8
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    NUMEDOC$ = "0001-" + NUMEDOC$
    '
    If CANCELPRINT% = 0 Then
      Call COMUNI("Coloque Formulario de Nota de Credito\Número " + NUMEDOC$ + "\en Impresora Predeterminada")
    End If
    '
    Screen.MousePointer = 11
    '
    TIFADOC% = 6
    TIPODOC$ = "Nota de Credito"
    CPAG% = 0
    '
    DESTIDOC% = Val(Text1.TEXT)
    NCX12% = DESTIDOC%
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
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
    CAITCRU% = 0
    CAITAB1% = 1
      'TETABU1$(1, CAITAB1%) = CODDX$
    TETABU1$(2, CAITAB1%) = "Descuento " + TRIM$(Text7) + " % s/Recibo " + TRIM$(Text11)
    TETABU1$(3, CAITAB1%) = "1"
    '
    TETABU1$(23, CAITAB1%) = NECRE  ' total neto
    TETABU1$(24, CAITAB1%) = IVACRE ' alicuota iva
    TETABU1$(25, CAITAB1%) = RNICRE ' importe iva del item
    TETABU1$(26, CAITAB1%) = TOCRE  ' total bruto
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = "" 'Label13.Caption
    CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = "" 'NROREMI$
    CODPARAM$(3) = "NRO-OCOM":  DOCPARAM$(3) = "" 'NUOCOMPRA$
    CODPARAM$(4) = "NRO-AENT":  DOCPARAM$(4) = "" 'NROPEDIDO$
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = NECRE  'Total(0).TEXT
    CODPARAM$(6) = "VENDEDOR": DOCPARAM$(6) = Label22  'Eco21(2).TEXT
    CODPARAM$(7) = "COND-PAG": DOCPARAM$(7) = "Contado" 'Eco21(1).TEXT
    CODPARAM$(8) = "POR-IVA": DOCPARAM$(8) = "" 'PORIVA(0).Caption
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = IVACRE 'Total(3).TEXT
    CODPARAM$(10) = "POR-IVA1": DOCPARAM$(10) = "" 'PORIVA(1).Caption
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = "" 'Total(4).TEXT
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = "" 'Total(1).TEXT
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = NECRE  'Total(2).TEXT
    'CODPARAM$(4) = "NRO-OCOM": DOCPARAM$(4) = Text10.TEXT
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = TOCRE 'Total(6).TEXT
    CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = "" 'TRIM$(PORDESCU.TEXT)
    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = RNICRE 'Total(5).TEXT
    CODPARAM$(23) = "TOTA-DOL": DOCPARAM$(23) = ""
    '
    If Label34.Caption = "B" Then
      DOCPARAM$(5) = TOCRE
      DOCPARAM$(9) = ""
      DOCPARAM$(13) = TOCRE
    End If
    '
    CODFOR$ = FORIPRE$
    LOIMPLET% = Val(ATRIFORM$(CODFOR$, "LOIMPLET"))
    If LOIMPLET% < 16 Then LOIMPLET% = 32
    TTTT# = Val(TOCRE)
    TOTFAC$ = NUMTEX$(TTTT#)
    '
    Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = NRX1$
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = NRX2$
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = NRX3$
    '
    CODPARAM$(20) = "DOMI-TRA": DOCPARAM$(20) = "" 'LENTREGA(0).TEXT
    CODPARAM$(21) = "DESTINO": DOCPARAM$(21) = "" 'LENTREGA(1).TEXT
    CODPARAM$(22) = "TRANSPOR": DOCPARAM$(22) = "" 'LENTREGA(2).TEXT
    CODPARAM$(24) = "MONEDEMI": DOCPARAM$(24) = "" 'MONEMIS
    CODPARAM$(25) = "PERC-IVA": DOCPARAM$(25) = "" 'Text15
    CODPARAM$(26) = "FECH-VEN": DOCPARAM$(26) = Text9 'TRIM$(FVENCIM)
    CAPARDOC% = 26
    '
    Call SETULTREG("!OBSERVOF", 0)
    Call CIERRARCH("!OBSERVOF")
    '
    NOCONFIRM% = 0
    HOJASUEL% = 0
    If CONTROL$("FACTUREM", "FORMULAR") = "HOJASUEL" Then
       HOJASUEL% = 1
    End If
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    If CANCELPRINT% = 0 Then
      If NOCONFIRM% > 0 Then
        Call COMUNI("Se han Emitido los Documentos Solicitados")
      End If
    End If
    '
End Sub

Sub NURECIBO(NROPRO&)
    '
    Screen.MousePointer = 11
    '
    FENUE$ = CONTROL$("", "FENURECI")
    FLIM% = FECHANUM(FENUE$)
    If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
    '
    AUTONURE% = 0
    If CONTROL$("", "NURECIST") = "AUTONUME" Then
      AUTONURE% = 1
    End If
    '
    COBRACT% = 0
    On Error Resume Next
    COBRACT% = Val(FORECISTA.Label22.Caption)
    On Error GoTo 0
    '
    CAMOLE% = 0: CAMOCOBRA% = 0
    NROMA& = 0
    For U& = ULTREG&("COBRAN") To 1 Step -1
        XX$ = REGLEIDO$("COBRAN", U&)
        CAMOLE% = CAMOLE% + 1
        If CAMOLE% > 512 Then
          If CAMOCOBRA% > 16 Then
            GoTo 90
          End If
        End If
        '
        FEMO% = CVI(Mid$(XX$, 7, 2))
        COBR% = CVI(Mid$(XX$, 23, 2))
        If FEMO% >= FLIM% Then
          If (COBR% = COBRACT% Or COBRACT% = 0) Then
            NROLE& = CVS(Mid$(XX$, 3, 4))
            If NROLE& = NROPRO& Then
               NROPRO& = 0
            End If
            If NROLE& > NROMA& Then NROMA& = NROLE&
            CAMOCOBRA% = CAMOCOBRA% + 1
          End If
        End If
    Next U&
    '
90  If AUTONURE% = 1 Then
      If COBRACT% = UCOBRA% Then
        If NROMA& < UNUREC& Then
          NROMA& = UNUREC&
        End If
      End If
    End If
    '
    If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
    '
End Sub
'

Sub SIGUERECI1()
    '
    If MODORECI% = 0 Then
        MODORECI% = 1       ' por emision
        If CONTROL$("RECISTA", "MODORDE") = "VENCIM" Then
           MODORECI% = 2
        End If
    End If
    '
    NCXI = Val(FORECISTA.Text1.TEXT)
    If NCXI < 1 Or NCXI > 32767 Then
        Call MENSERR(24, "Numero de Cuenta no Válido")
        FORECISTA.Text1.SetFocus
        Exit Sub
    End If
    '
    NC% = NCXI
    If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
        Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Clientes")
        FORECISTA.Text1.SetFocus
        Exit Sub
    End If
    '
    FE$ = TRIM$(FORECISTA.Text9.TEXT)
    FF% = FECHANUM(FE$)
    HOI% = HOY(HO$)
    If FF% < 1 Or FF% > HOI% + 1 Then
        Call MENSERR(24, "Fecha no Válida")
        FORECISTA.Text9.SetFocus
        Exit Sub
    End If
    '
    If DIENTRE%(FF%, HOI%) > 30 Then
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "NEUMAN") < 1 Or Val(Left$(TRIM$(FORECISTA.Label22.Caption), 2)) <> 1 Then
        Call MENSERR(24, "Fecha no Válida")
        FORECISTA.Text9.SetFocus
        Exit Sub
      End If
    End If
   '
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 11
    NC% = Val(FORECISTA.Text1.TEXT)
    If NC% > 0 Then
        FORECISTA.Text2.TEXT = RASOCLI$(NC%)
        FORECISTA.Text3.TEXT = DOMICLI$(NC%)
        FORECISTA.Text4.TEXT = CPOSCLI$(NC%)
        FORECISTA.Text5.TEXT = LOCACLI$(NC%)
        'FORECISTA.Text7.TEXT = TEL1CLI$(NC%)
        'FORECISTA.Text10.TEXT = CSTRING$(MKD$(SALDCLI#(NC%)), 4, 13, 2, 2)
        FORECISTA.Refresh
    End If
    '
100 FORECISTA.List1.Clear
    FORECISTA.List1.Refresh
    FORECISTA.List2.Clear
    FORECISTA.List2.Refresh
    '
    MAXDEU& = ULTREG&("SVD202")
    RFD$ = RECFILT$("SVD202", 1, MKI$(NC%))
    '
    RESCUDEU.LINDICE.Clear
    For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("SVD202", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("SVD202", 1)
            Call GRAREG("SVD202", X$, 1)
            Call CIERRARCH("SVD202")
            GoTo 100
        End If
        IMPO# = CVD(Mid$(X$, 21, 8))
        SALDO# = CVD(Mid$(X$, 11, 8))
        If Abs(SALDO#) >= 0.005 Then
          If Sgn(IMPO#) <> Sgn(SALDO#) Or Abs(SALDO#) > Abs(IMPO#) Then
            ABUS$ = Left$(X$, 8)
            For UU& = 1 To Len(RFD$) Step 4
                REG& = CVS(Mid$(RFD$, UU&, 4))
                X$ = REGLEIDO$("SVD202", REG&)
                If Left$(X$, 8) = ABUS$ Then
                    Call GRAREG("SVD202", String$(32, 0), REG&)
                End If
            Next UU&
            GoTo 100
          End If
        End If
        '
        If MODORECI% = 1 Then
             FF% = CVI(Mid$(X$, 19, 2)) ' orden emision
           Else
             FF% = CVI(Mid$(X$, 9, 2)) ' orden vencimiento
        End If
        Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2)
        RESCUDEU.LINDICE.AddItem Z$
    Next U&
    '
    SUMADEU# = 0: SUMAFAC# = 0
    For U& = 1 To RESCUDEU.LINDICE.ListCount
        RESCUDEU.LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(RESCUDEU.LINDICE.TEXT, 7, 8))
        If REG& > 0 Then
            If REG& <= MAXDEU& Then
                X$ = REGLEIDO$("SVD202", REG&)
                TI% = Asc(Mid$(X$, 3, 1))
                VA% = Asc(Mid$(X$, 4, 1))
                TIPOS$ = "Factura": If TI% = 0 Then TIPOS$ = "Ajuste Saldo"
                                   If TI% = 1 Then TIPOS$ = "Venta Contado"
                                   If TI% = 3 Then TIPOS$ = "Nota Debito"
                                   If TI% = 4 Then TIPOS$ = "Anticipo"
                                   If TI% = 5 Then TIPOS$ = "Desc. Caja"
                                   If TI% = 6 Then TIPOS$ = "Nota Credito"
                                   If TI% = 7 Then TIPOS$ = "Ajuste Saldo"
                                   '
                SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
                IMPO# = CVD(Mid$(X$, 21, 8))
                If TI% <> 0 Then IMPO# = Abs(IMPO#)
                IMPO# = SG% * IMPO#
                '
                SALDO# = CVD(Mid$(X$, 11, 8))
                If TI% <> 0 Then SALDO# = Abs(SALDO#)
                SALDO# = SG% * SALDO#
                '
                If SALDO# >= 0.005 Then
                    SUMAFAC# = SUMAFAC# + IMPO#
                    SUMADEU# = SUMADEU# + SALDO#
                    NROL& = CVS(Mid$(X$, 5, 4))
                    FEMI = CVI(Mid$(X$, 19, 2))
                    FEVE = CVI(Mid$(X$, 9, 2))
                    NUSUCUVE% = CVI(Mid$(X$, 29, 2))
                    '
                    IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
                    SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
                    DOCU$ = FORMDOC$(NROL&, TI%, VA%, NUSUCUVE%)
                    '
                    Z$ = Space$(128)
                    Call REPLA(Z$, TIPOS$, 1, 14)
                    Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 13, 8)
                    Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
                    Call REPLA(Z$, IMPTEX$, 31, 14)
                    Call REPLA(Z$, FECHATEX$(FEVE), 48, 8)
                    Call REPLA(Z$, SALTEX$, 56, 14)
                    Call REPLA(Z$, DOCU$, 97, 18)
                    Call REPLA(Z$, TRIM$(Str$(REG&)), 121, 6)
                    '
                    FORECISTA.List2.AddItem Z$
                End If
                '
            End If
        End If
        '
    Next U&
    '
    FORECISTA.List2.ListIndex = -1
    FORECISTA.List2.Refresh
    'FORECISTA.LABEL14(0).Caption = CSTRING$(MKD$(SUMAFAC#), 4, 14, 2, 2)
    'FORECISTA.LABEL14(0).Refresh
    FORECISTA.Label8.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    FORECISTA.Label8.Refresh
    FORECISTA.Label10.Caption = ""
    FORECISTA.Label10.Refresh
    FORECISTA.List2.Visible = True
    FORECISTA.List2.SetFocus
    If FORECISTA.List2.ListCount > 0 Then
        VACON% = 1
        FORECISTA.List2.ListIndex = 0
        VACON% = 0
    End If
    '
    FORECISTA.BOTNEXT.Enabled = False
    Call SETULTREG("!CARMECO", 0)
    Call FRESHGRID
    Call BLANARCH("!CARMECO")
    Screen.MousePointer = 1
    '
    Call SIGUERECI2
    '
End Sub
'

Sub SIGUERECI2()
    '
    HOI% = HOY(HO$)
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Ingreso de Medios de Pago Codificados"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORECISTA.Top + FORECISTA.Picture1.Top + 100))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORECISTA.Left + FORECISTA.Picture1.Left))
    '
    VTB% = VENTABU%("CARMECOW" + ATRI$)
    Call CIERRARCH("!CARMECO")
    AAA = ULTREG&("!CARMECO")
    If VTB% < 0 Then
        FORECISTA.BOTNEXT.Enabled = True
        Exit Sub
    End If
    '
    ATXY% = 0
    For U& = 1 To ULTREG&("!CARMECO")
      X$ = REGLEIDO$("!CARMECO", U&)
      PROTER$ = UCase$(TRIM$(Mid$(X$, 43, 2)))
      If PROTER$ = "" Or Left$(PROTER$, 1) = "P" Then
           PROTER$ = "  "
         Else
           PROTER$ = "3 "
      End If
      If TRIM$(PROTER$) <> "" Then ATXY% = 1
    Next U&
    '
    If ATXY% > 0 Then
      ATRI$ = "/NC"
      ATRI$ = ATRI$ + "/TITUPAN=Ingreso CUIT Cheques de Terceros"
      ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORECISTA.Top + FORECISTA.Picture1.Top + 100))
      ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORECISTA.Left + FORECISTA.Picture1.Left))
      '
      VTB% = VENTABU%("CARMECOX" + ATRI$)
      Call CIERRARCH("!CARMECO")
      AAA = ULTREG&("!CARMECO")
      If VTB% < 0 Then
          FORECISTA.BOTNEXT.Enabled = True
          Exit Sub
      End If
    End If
    '
    ATXX% = 0
    FORECISTA.List1.Clear
    SUMAVA# = 0
    FORECISTA.Label14(0) = ""
    For U& = 1 To ULTREG&("!CARMECO")
        X$ = REGLEIDO$("!CARMECO", U&)
        If X$ = REBLA$ Then GoTo 290
        MECO% = CVI(Left$(X$, 2))
        BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(MECO%)))
        '
        If BANCO$ = "" Then
            Call MENSERR(24, "Medio de Pago no Válido")
            GoTo 200
        End If
        '
        NUVA$ = Mid$(X$, 19, 12)
        IMVA# = CVD(Mid$(X$, 31, 8))
        If IMVA# < 0.005 Then
            Call MENSERR(24, "Falta Importe")
            GoTo 200
        End If
        '
        FEMI = CVI(Mid$(X$, 39, 2))
        If FEMI = 0 Then
          Call REPLA(X$, MKI$(HOI%), 39, 2)
          Call GRAREG("!CARMECO", X$, U&)
          FEMI = HOI%
          ATXX% = 1
        End If
        F1 = FEMI: F2 = HOI%
        If FEMI < HOI% - 32 * 13 Or FEMI > HOI% + 30 Or DIENTRE%(F1, F2) > 365 Then
          EPAC$ = TRIM$(UCase$(EMPREAC$))
          If InStr(EPAC$, "NEUMAN") < 1 Or Val(Left$(TRIM$(FORECISTA.Label22.Caption), 2)) <> 1 Then
            Call MENSERR(24, "Fecha de Emisión no Válida")
            GoTo 200
          End If
        End If
        '
        FEVE = CVI(Mid$(X$, 41, 2))
        If FEVE = 0 Then
          FEMII% = FEMI
          Call REPLA(X$, MKI$(FEMII%), 41, 2)
          Call GRAREG("!CARMECO", X$, U&)
          FEVE = FEMII%
          ATXX% = 1
        End If
        F1 = FEVE: F2 = HOI%
        If FEVE < HOI% - 64 Or DIENTRE%(F1, F2) > 30 Then
          EPAC$ = TRIM$(UCase$(EMPREAC$))
          If InStr(EPAC$, "NEUMAN") < 1 Or Val(Left$(TRIM$(FORECISTA.Label22.Caption), 2)) <> 1 Then
            Call MENSERR(24, "Fecha de Vencimiento no Válida")
            GoTo 200
          End If
        End If
        '
        PROTER$ = UCase$(TRIM$(Mid$(X$, 43, 2)))
        If PROTER$ = "" Or Left$(PROTER$, 1) = "P" Then
             PROTER$ = "  "
           Else
             PROTER$ = "3 "
        End If
        Call REPLA(X$, PROTER$, 43, 2)
        Call GRAREG("!CARMECO", X$, U&)
        '
        CLEA% = CVI(Mid$(X$, 45, 2))
        If CLEA% < 24 Then
          CLEA% = 24
          Call REPLA(X$, MKI$(CLEA%), 45, 2)
          Call GRAREG("!CARMECO", X$, U&)
        End If
        If (CLEA% Mod 24) <> 0 Then
            Call MENSERR(24, "Clearing no Válido")
            GoTo 200
        End If
        Call REPLA(X$, MKI$(CLEA%), 45, 2)
        Call GRAREG("!CARMECO", X$, U&)
        '
        Z$ = Space$(96)
        SUMAVA# = SUMAVA# + IMVA#
        IMPTEX$ = CSTRING$(MKD$(IMVA#), 3, 14, 2, 2)
        Call REPLA(Z$, CSTRING$(MKI$(MECO%), 1, 4, 0, 2), 1, 4)
        Call REPLA(Z$, BANCO$, 7, 14)
        Call REPLA(Z$, NUVA$, 23, 12)
        Call REPLA(Z$, PROTER$, 38, 2)
        Call REPLA(Z$, CSTRING$(MKI$(CLEA%), 1, 4, 0, 2), 42, 4)
        Call REPLA(Z$, FECHATEX$(FEMI), 50, 8)
        Call REPLA(Z$, FECHATEX$(FEVE), 62, 8)
        Call REPLA(Z$, IMPTEX$, 70, 14)
        '
        FORECISTA.List1.AddItem Z$
        '
        If TRIM$(NUVA$) <> "" Then
          TBUS$ = Left$(Z$, 34)
          For JI% = 1 To List1.ListCount - 1
            If Left$(List1.List(JI% - 1), 34) = TBUS$ Then
              Call MENSERR(24, "Cheque o Valor Repetido")
              GoTo 200
            End If
          Next JI%
        End If
        '
290 Next U&
    '
    If ATXX% > 0 Then
      Call MENSERR(24, "Verifique el Ajuste de Fechas de Emisión y Vencimiento.")
      GoTo 200
    End If
    '
    If FORECISTA.List1.ListCount > 0 Then
        FORECISTA.List1.ListIndex = 0
    End If
    '
    FORECISTA.Label14(0).Caption = TRIM$(CSTRING$(MKD$(SUMAVA#), 4, 14, 2, 2))
    FORECISTA.Label14(0).Refresh
    ANTI# = SUMAVA# + Val(FORECISTA.DESCUEN(0)) - CDbl(Val(FORECISTA.Label18(0).Caption))
    'If ANTI# < 0 Then
    '      FORECISTA.Label19.Caption = "Descuentos:"
    '    Else
    '      FORECISTA.Label19.Caption = "Anticipos:"
    'End If
    FORECISTA.Label20(0).Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    FORECISTA.Label20(0).Refresh
    '
End Sub
'

Sub SIGUERECI3()
    '
    If FORECISTA.List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!APLIRECI")
    For U& = 1 To FORECISTA.List2.ListCount
        FORECISTA.List2.ListIndex = U& - 1
        Z$ = FORECISTA.List2.TEXT
        TIPOS$ = Left$(Z$, 12)
        NROL& = Val(Mid$(Z$, 13, 8))
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(Val(Mid$(Z$, 31, 14)))
        FEVEI% = FECHANUM(Mid$(Z$, 48, 8))
        SALD# = CDbl(Val(Mid$(Z$, 56, 14)))
        ACRE# = CDbl(Val(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 97, 18)
        REG2& = Val(Mid$(Z$, 121, 6))
        '
        Y$ = REGBLAN$("!APLIRECI")
        Call REPLA(Y$, TIPOS$, 1, 16)
        Call REPLA(Y$, MKS$(NROL&), 17, 4)
        Call REPLA(Y$, MKI$(FEMII%), 21, 2)
        Call REPLA(Y$, MKD$(IPED#), 23, 8)
        Call REPLA(Y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(Y$, MKD$(SALD#), 33, 8)
        Call REPLA(Y$, MKD$(ACRE#), 41, 8)
        Call REPLA(Y$, DOCU$, 49, 18)
        Call REPLA(Y$, MKS$(REG2&), 67, 4)
        '
        Call REGAPP("!APLIRECI", Y$)
    Next U&
    Call CIERRARCH("!APLIRECI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Cobranza"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORECISTA.Top + FORECISTA.Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORECISTA.Left + FORECISTA.Picture2.Left))
    '
    VTB% = VENTABU%("APLIRECI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        FORECISTA.List2.Clear
        For U& = 1 To ULTREG&("!APLIRECI")
            Y$ = REGLEIDO$("!APLIRECI", U&)
            TIPOS$ = Left$(Y$, 16)
            NROL& = CVS(Mid$(Y$, 17, 4))
            FEMI = CVI(Mid$(Y$, 21, 2))
            IPED# = CVD(Mid$(Y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 14, 2, 2)
            FEVE = CVI(Mid$(Y$, 31, 8))
            SALD# = CVD(Mid$(Y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 14, 2, 2)
            ACRE# = CVD(Mid$(Y$, 41, 8))
            '
            If ACRE# < 0 Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If ACRE# > SALD# Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                GoTo 200
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(Y$, 49, 18)
            REG2& = CVS(Mid$(Y$, 67, 4))
            '
            Z$ = Space$(128)
            Call REPLA(Z$, TIPOS$, 1, 14)
            Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 13, 8)
            Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(Z$, IMPTEX$, 31, 14)
            Call REPLA(Z$, FECHATEX$(FEVE), 48, 8)
            Call REPLA(Z$, SALTEX$, 56, 14)
            Call REPLA(Z$, ACRTEX$, 70, 14)
            Call REPLA(Z$, DOCU$, 97, 18)
            Call REPLA(Z$, TRIM$(Str$(REG2&)), 121, 6)
            '
            FORECISTA.List2.AddItem Z$
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To FORECISTA.List2.ListCount
        FORECISTA.List2.ListIndex = U& - 1
        Z$ = FORECISTA.List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(Val(Mid$(Z$, 70, 14)))
    Next U&
    '
    FORECISTA.Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    FORECISTA.Label10.Refresh
    FORECISTA.Label18(0).Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    FORECISTA.Label18(0).Refresh
    ANTI# = CDbl(Val(FORECISTA.Label14(0).Caption)) + CDbl(Val(FORECISTA.DESCUEN(0))) - SUMACRE#
    'If ANTI# < 0 Then
    '      FORECISTA.Label19.Caption = "Descuentos:"
    '    Else
    '      FORECISTA.Label19.Caption = "Anticipos:"
    'End If
    FORECISTA.Label20(0).Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    FORECISTA.Label20(0).Refresh
    '
    FORECISTA.List2.Refresh
    VACON% = 1
    '
End Sub
'

Sub GRARECISTA(OKEY%)
    '
700 OKEY% = 0
    '
    NC% = Val(FORECISTA.Text1.TEXT)
    NROO& = Val(Mid$(TRIM$(FORECISTA.Text11.TEXT), 4))
    FC% = FECHANUM(TRIM$(FORECISTA.Text9.TEXT))
    TOTAVAL# = CDbl(Val(FORECISTA.Label14(0).Caption))
    ANTI# = CDbl(Val(FORECISTA.Label20(0).Caption))
    IDESCU# = Val(FORECISTA.DESCUEN(0))
    URE202& = ULTREG&("SVD202")
    '
    COBRACT% = 0
    On Error Resume Next
    COBRACT% = Val(FORECISTA.Label22.Caption)
    On Error GoTo 0
    '
    HAYCREDES% = 0
    If IDESCU# >= 0.005 Then
      HAYCREDES% = 1
      VAX% = 1: FORECISTA.Label34 = "A"
      PVCL% = PIVACLI%(NC%)
      If PVCL% < 1 Or PVCL% > 2 Then
        VAX% = 2
        FORECISTA.Label34 = "B"
      End If
    End If
    '
    Call BLOQARCH("COBRAN")
    Call BLOQARCH("CHECAR")
    Call BLOQARCH("APLICOB")
    Call BLOQARCH("CARBIMON")
    If HAYCREDES% = 1 Then
      Call BLOQARCH("FACTUR")
      For U& = ULTREG&("FACTUR") To 1 Step -1
        YX$ = REGLEIDO$("FACTUR", U&)
        If Asc(Mid$(YX$, 13, 1)) = 6 Then ' es nota de credito
          If Asc(Mid$(YX$, 14, 1)) = VAX% Then
            NROCRED& = 1 + CVS(Mid$(YX$, 3, 4))
            GoTo 710
          End If
        End If
      Next U&
      NROCRED& = 1
    End If
    '
710 Screen.MousePointer = 11   ' Grabando Recibo - Espere un Momento, por Favor"
    FORECISTA.Label35 = TRIM$(Str$(NROCRED&))
    '
    NROPRES& = Val(Mid$(FORECISTA.Text11.TEXT, 4))
    NROPRO& = NROPRES&
    Call NURECIBO(NROPRO&)
    If NROPRO& <> NROPRES& Then
       Call CIERRARCH("COBRAN")
       Call CIERRARCH("CHECAR")
       Call CIERRARCH("APLICOB")
       Call CIERRARCH("CARBIMON")
       Call CIERRARCH("FACTUR")
       Screen.MousePointer = 1
       Call MENSERR(24, "Al Grabar este Recibo\el Número Propuesto Estaba Ocupado.\  \Verifique el Nuevo Número Propuesto\y Vuelva a Grabar e Imprimir.")
       FORECISTA.Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
       OKEY% = 0
       Exit Sub
    End If
    '
    URE202& = ULTREG&("SVD202")
    '
    If ANTI# >= 0.005 Then
        '
        Z$ = REGBLAN$("SVD202")
        Call REPLA(Z$, MKI$(NC%), 1, 2)
        Call REPLA(Z$, Chr$(4), 3, 1)
        Call REPLA(Z$, Chr$(0), 4, 1)
        Call REPLA(Z$, MKS$(NROO&), 5, 4)
        Call REPLA(Z$, MKI$(FC%), 9, 2)
        Call REPLA(Z$, MKD$(ANTI#), 11, 8)
        Call REPLA(Z$, MKI$(FC%), 19, 2)
        Call REPLA(Z$, MKD$(TOTAVAL#), 21, 8)
        Call REPLA(Z$, String$(2, 255), 31, 2)
        '
        Call BLOQARCH("SVD202")
        Call REGAPP("SVD202", Z$)
        Call CIERRARCH("SVD202")
        '
    End If
    OKEY% = 1
    '
800 ' *********************************** aqui blanquear facturas canceladas
    '
    XX% = 0
    VACON% = 1
    FORECISTA.List2.Visible = False
    ILISTA& = FORECISTA.List2.ListIndex
    TOLISTA& = FORECISTA.List2.TopIndex
    '
    For U& = 1 To FORECISTA.List2.ListCount
        'FORECISTA.List2.ListIndex = U& - 1
        X$ = FORECISTA.List2.List(U& - 1)
        CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
        If CANCE# >= 0.005 Then
            NREG& = Val(Mid$(X$, 121, 6))
            If NREG& > 0 Then
                If NREG& <= ULTREG&("SVD202") Then
                    '
                    SS$ = REGLEIDO$("SVD202", NREG&)
                    SALDEU# = CVD(Mid$(SS$, 11, 8)) - CANCE#
                    Call REPLA(SS$, MKD$(SALDEU#), 11, 8)
                    Call GRAREG("SVD202", SS$, NREG&)
                    '
                End If
            End If
        End If
    Next U&
    Call CIERRARCH("SVD202")
    OKEY% = 2
    '
750 RGX& = REGICLI&(NC%)
    If RGX& > 0 Then
        If RGX& <= ULTREG&("DEUDOR2") Then
           DX$ = REGLEIDO$("DEUDOR2", RGX&)
           SALDO# = CVD(Mid$(DX$, 109, 8)) - TOTAVAL# - IDESCU#
           Call REPLA(DX$, MKD$(SALDO#), 109, 8)
           Call GRAREG("DEUDOR2", DX$, RGX&)
        End If
    End If
    Call CIERRARCH("DEUDOR2")
    OKEY% = 3
    '
    Z$ = REGBLAN$("COBRAN")
    Call REPLA(Z$, MKI$(NC%), 1, 2)
    Call REPLA(Z$, MKS$(NROO&), 3, 4)
    Call REPLA(Z$, MKI$(FC%), 7, 2)
    Call REPLA(Z$, String$(6, 0), 9, 6)
    Call REPLA(Z$, MKD$(TOTAVAL#), 15, 8)
    Call REPLA(Z$, MKI$(COBRACT%), 23, 2)
    '
    Call REGAPP("COBRAN", Z$)
    Call CIERRARCH("COBRAN")
    UCOBRA% = COBRACT%
    UNUREC& = NROO&
    '
    FORECISTA.NECRE = ""
    FORECISTA.IVACRE = ""
    FORECISTA.RNICRE = ""
    FORECISTA.TOCRE = ""
    '
    If HAYCREDES% = 1 Then
      COEFIVA = 1 + ALIVA / 100
      If PVCL% = 2 Then COEFIVA = 1 + 1.5 * ALIVA / 100
      INECRE# = (Int(100 * IDESCU# / COEFIVA + 0.5)) / 100
      If PVCL% <> 2 Then
          IVACRED# = IDESCU# - INECRE#
          RNICRED# = 0
        Else
          IVACRED# = (Int(INECRE# * ALICU + 0.5)) / 100
          RNICRED# = IDESCU# - INECRE# - IVACRED#
      End If
      '
      Z$ = REGBLAN$("FACTUR")
      Call REPLA(Z$, MKI$(NC%), 1, 2)
      Call REPLA(Z$, MKS$(NROCRED&), 3, 4)
      Call REPLA(Z$, MKI$(FC%), 7, 2)
      Call REPLA(Z$, MKI$(FC%), 9, 2)
      Call REPLA(Z$, String$(2, 0), 11, 2)
      Call REPLA(Z$, Chr$(6) + Chr$(VAX%), 13, 2)
      Call REPLA(Z$, MKD$(INECRE#), 15, 8)
        FORECISTA.NECRE = FORMATNUM$(INECRE#, "F12.2")
      Call REPLA(Z$, MKD$(IVACRED#), 23, 8)
        FORECISTA.IVACRE = FORMATNUM$(IVACRED#, "F12.2")
      Call REPLA(Z$, MKD$(RNICRED#), 31, 8)
        FORECISTA.RNICRE = FORMATNUM$(RNICRED#, "F12.2")
      Call REPLA(Z$, MKD$(IDESCU#), 39, 8)
        FORECISTA.TOCRE = FORMATNUM$(IDESCU#, "F12.2")
      Call REPLA(Z$, String$(18, 0), 47, 18)
      Call REGAPP("FACTUR", Z$)
      Call CIERRARCH("FACTUR")
    End If
    '
    OKEY% = 4
    '
    'DESCAJ# = 0
    'ANTI# = CDbl(Val(FORECISTA.Label20(0).Caption))
    'If ANTI# < 0 Then
    '    DESCAJ# = Abs(ANTI#)
    '    ANTI# = 0
    'End If
    '
    'If DESCAJ# >= 0.005 Then
'aun no se permiten descuentos de caja
    'TIPO% = 5
     ' grabar como tipo 5 en 'FACTUR.DAT'
    'End If
    '
    OKEY% = 5

830 ' *************************************** aqui grabar archivo APLICOB.DAT
    '
    For U& = 1 To FORECISTA.List2.ListCount
      'FORECISTA.List2.ListIndex = U& - 1
      X$ = FORECISTA.List2.List(U& - 1)
      CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
      If CANCE# >= 0.005 Then
        CANCE1# = CANCE#
        CANCE2# = 0
        If HAYCREDES% = 1 Then
          CANCE1# = (Int(100 * CANCE# * TOTAVAL# / (TOTAVAL + IDESCU#) + 0.5)) / 100
          CANCE2# = CANCE# - CANCE1#
        End If
        NREG& = Val(Mid$(X$, 121, 6))
        If NREG& > 0 Then
          If NREG& <= URE202& Then
            SS$ = REGLEIDO$("SVD202", NREG&)
            NUSUCUVE% = CVI(Mid$(SS$, 29, 2))
            DOCORI$ = FORMDOC$(CVS(Mid$(SS$, 5, 4)), Asc(Mid$(SS$, 3, 1)), Asc(Mid$(SS$, 4, 1)), NUSUCUVE%)
            ITOTORI# = CVD(Mid$(SS$, 21, 8))
            IVAORI# = (Int(ALIVA * ITOTORI# / (1 + ALIVA / 100) + 0.5)) / 100
            '
            RefApli$ = "Recibo " + TRIM$(Str$(NROO&))
            Call APLIPAG(FC%, NC%, 4, 1, SUC%, NROO&, CANCE1#, RefApli$, DOCORI$, ITOTORI#, IVAORI#)
            '
            If CANCE2# >= 0.005 Then
              RefApli$ = "N/Cred." + TRIM$(Str$(NROCRED&))
              Call APLIPAG(FC%, NC%, 6, VAX%, SUC%, NROCRED&, CANCE2#, RefApli$, DOCORI$, ITOTORI#, IVAORI#)
            End If
            '
          End If
        End If
      End If
    Next U&
    VACON% = 0
    Call CIERRARCH("APLICOB")
    OKEY% = 6
    '
    RFF$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    If Len(RFF$) > 0 Then
      TICAMBX# = Val(FORECISTA.TICAMBIS.TEXT)
      If TICAMBX# < 1 Then TICAMBX# = 1
      For U& = 1 To FORECISTA.List2.ListCount
        X$ = FORECISTA.List2.List(U& - 1)
        CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
        If CANCE# >= 0.005 Then
          NREG& = Val(Mid$(X$, 121, 6))
          If NREG& > 0 Then
            If NREG& <= URE202& Then
              SS$ = REGLEIDO$("SVD202", NREG&)
              ABUSQUE$ = Mid$(SS$, 3, 6)
              '
              For JKLX& = 1 To Len(RFF$) Step 4
                RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
                YCB$ = REGLEIDO$("CARBIMON", RECABI&)
                If Mid$(YCB$, 5, 6) = ABUSQUE$ Or Mid$(YCB$, 37, 6) = ABUSQUE$ Then
                  Call REPLA(YCB$, MKI$(FC%), 33, 2)
                  Call REPLA(YCB$, MKI$(NC%), 35, 2)
                  Call REPLA(YCB$, Chr$(4), 37, 1)
                  Call REPLA(YCB$, Chr$(1), 38, 1)
                  Call REPLA(YCB$, MKS$(NROO&), 39, 4)
                  Call REPLA(YCB$, MKI$(SUC%), 43, 2)
                  Call REPLA(YCB$, FORMDOC$(NROO&, 4, 1, SUC%), 45, 18)
                  Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
                  Call REPLA(YCB$, MKD$(0), 65, 8)
                  Call REPLA(YCB$, MKD$(CANCE#), 73, 8)
                  Call REPLA(YCB$, MKD$(0), 81, 8)
                  Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
                  Call REPLA(YCB$, MKD$(0), 97, 8)
                  Call REPLA(YCB$, MKD$(CANCE# / TICAMBX#), 105, 8)
                  Call REPLA(YCB$, MKD$(0), 113, 8)
                  '
                  If Mid$(YCB$, 5, 6) <> ABUSQUE$ Then ' APLICADO AL DEBITO POR D.C.
                    Call REPLA(YCB$, MKD$(0), 105, 8)  ' NO AJUSTA SALDO EN DOLARES
                  End If
                  '
                  Call REGAPP("CARBIMON", YCB$)
                  GoTo 835
                End If
              Next JKLX&
            End If
          End If
        End If
835   Next U&
    End If
    Call CIERRARCH("CARBIMON")
    '
840 ' ************************************** aqui grabar archivo CHECAR.DAT
    '
    For U& = 1 To ULTREG&("!CARMECO")
        X$ = REGLEIDO$("!CARMECO", U&)
        TIPO% = CVI(Left$(X$, 2))
        BANCO$ = Mid$(X$, 3, 16)
        NUCHE$ = Mid$(X$, 19, 12)
        IMCHE# = CVD(Mid$(X$, 31, 8))
        FECHE% = CVI(Mid$(X$, 39, 2)): If FECHE% < 1 Then FECHE% = HO%
        FEVEN% = CVI(Mid$(X$, 41, 2)): If FEVEN% < FECHE% Then FEVEN% = FECHE%
        '
        If IMCHE# > 0 Then
            '
            Z$ = REGBLAN$("CHECAR")
            Call REPLA(Z$, MKI$(TIPO%), 1, 2)
            Call REPLA(Z$, MKI$(NC%), 3, 2)
            Call REPLA(Z$, NUCHE$, 5, 12)
            Call REPLA(Z$, BANCO$, 17, 16)
            Call REPLA(Z$, MKD$(IMCHE#), 33, 8)
            Call REPLA(Z$, MKI$(FEVEN%), 41, 2)
            Call REPLA(Z$, MKI$(FECHE%), 43, 2)
            Call REPLA(Z$, MKI$(FC%), 45, 2)
            Call REPLA(Z$, MKS$(NROO&), 47, 4)
            Call REPLA(Z$, MKI$(COBRACT%), 51, 2)
            Call REPLA(Z$, String$(12, 0), 53, 14)
            '
            Call REGAPP("CHECAR", Z$)
            '
        End If
    Next U&
    Call CIERRARCH("CHECAR")
    OKEY% = 6
    '
890 Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
End Sub
'


