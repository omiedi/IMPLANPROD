VERSION 5.00
Begin VB.Form FORPASTA 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Pago (Operaciones Comerciales de Compras)"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11910
   Icon            =   "FORPASTA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
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
      Height          =   8500
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command8 
         Caption         =   "Sort."
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
         Picture         =   "FORPASTA.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Cambia Ordenamiento de Valores (Fecha / Importe)"
         Top             =   3330
         Width           =   855
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   210
         TabIndex        =   55
         Top             =   0
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   540
         Left            =   6825
         TabIndex        =   52
         Top             =   140
         Width           =   3900
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Fondo Fijo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   540
            Left            =   1155
            TabIndex        =   54
            Top             =   0
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Orden de Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   540
            Left            =   2520
            TabIndex        =   53
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   10920
         ScaleHeight     =   75
         ScaleWidth      =   795
         TabIndex        =   50
         Top             =   6930
         Width           =   855
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Re-apli"
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
         Picture         =   "FORPASTA.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Re-Aplicacion de Creditos y Pagos a Cuenta"
         Top             =   5920
         Width           =   855
      End
      Begin VB.CommandButton Command6 
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
         Picture         =   "FORPASTA.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Anular, Re-Imprimir, Imputación de Créditos, Tablas"
         Top             =   5080
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
            BackColor       =   &H00C0C000&
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
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   " M-P    Referencia                  Numero               Emision        Vencim.           Importe            Selección"
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
               X1              =   4410
               X2              =   4410
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line7 
               X1              =   2835
               X2              =   2835
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
               X1              =   6930
               X2              =   6930
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
               Left            =   270
               TabIndex        =   36
               Top             =   315
               Width           =   10500
            End
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
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFC0C0&
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
               Left            =   -40
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
               TabIndex        =   45
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
               Left            =   8400
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
               Left            =   6510
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
               X2              =   10395
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
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
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
               Caption         =   " Documento  Número         Emision                Importe       Vencim.                  Saldo                Acredit.       "
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
         Height          =   645
         Left            =   10920
         Picture         =   "FORPASTA.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Aprueba - Graba e Imprime Orden de Pago"
         Top             =   7245
         Width           =   855
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
         Picture         =   "FORPASTA.frx":0F12
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Apertura de Nueva Cuenta de Proveedor"
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
         Picture         =   "FORPASTA.frx":121C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   2510
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
         Picture         =   "FORPASTA.frx":1526
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Notas de Proveedores"
         Top             =   4250
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
         Picture         =   "FORPASTA.frx":1830
         TabIndex        =   16
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   315
         Width           =   855
      End
      Begin VB.CommandButton BOTNEXT 
         Caption         =   "Next"
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
         Height          =   750
         Left            =   10920
         Picture         =   "FORPASTA.frx":1B3A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Carga de Valores de Pago Propios"
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
         Left            =   9030
         TabIndex        =   4
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
         Left            =   9345
         TabIndex        =   2
         Text            =   " "
         Top             =   735
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
         TabIndex        =   3
         Top             =   735
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
         TabIndex        =   1
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
         TabIndex        =   0
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label28 
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
         Left            =   105
         TabIndex        =   62
         Top             =   6930
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label27 
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
         Left            =   1785
         TabIndex        =   61
         Top             =   7035
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label26 
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
         Left            =   1050
         TabIndex        =   60
         Top             =   7035
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label25 
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
         Left            =   2625
         TabIndex        =   59
         Top             =   7035
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label24 
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
         Left            =   4620
         TabIndex        =   58
         Top             =   6930
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label23 
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
         Left            =   3360
         TabIndex        =   57
         Top             =   6930
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label22 
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
         Left            =   4305
         TabIndex        =   48
         Top             =   7245
         Width           =   2220
      End
      Begin VB.Label Label21 
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
         Left            =   5985
         TabIndex        =   47
         Top             =   7560
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
         Left            =   8610
         TabIndex        =   44
         Top             =   7560
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
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
         Left            =   6825
         TabIndex        =   43
         Top             =   7245
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
         Left            =   630
         TabIndex        =   42
         Top             =   7560
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Retenciones:"
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
         Left            =   2205
         TabIndex        =   41
         Top             =   7245
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
         Left            =   3360
         TabIndex        =   38
         Top             =   7560
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Creditos:"
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
         Left            =   -945
         TabIndex        =   37
         Top             =   7245
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
         Top             =   1155
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
         Left            =   8505
         TabIndex        =   14
         Top             =   735
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
Attribute VB_Name = "FORPASTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$

Private Sub BOTNEXT_Click()
    '
    If Val(Label18.Caption) < 0.005 Then
       Exit Sub
    End If
    '
    If Val(Label20.Caption) < 0 Then
         Call MENSERR(24, "Valores Cargados exceden Créditos")
         Exit Sub
       ElseIf Val(Label20.Caption) >= 0.005 Then
         Call CIERRARCH("!CARMEPA")
         Call BLANARCH("!CARMEPA")
         X$ = REGBLAN$("CARMEPA")
         HOI% = HOY(HO$)
         Call REPLA(X$, MKD$(Val(Label20.Caption)), 31, 8)
         Call REPLA(X$, MKI$(HOI%), 39, 2)
         Call REPLA(X$, MKI$(HOI%), 41, 2)
         Call REGAPP("!CARMEPA", X$)
         '
10       VTB% = VENTABU%("CARMEPA")
         If VTB% < 0 Then
            Call CIERRARCH("!CARMEPA")
            Call BLANARCH("!CARMEPA")
            Exit Sub
         End If
         SUVAPROP# = 0
         For U& = 1 To ULTREG&("!CARMEPA")
            X$ = REGLEIDO$("!CARMEPA", U&)
            IVAL# = CVD(Mid$(X$, 31, 8))
            If IVAL# < 0.005 Then
               Call MENSERR(24, "No se Admiten Valores Negativos")
               GoTo 10
            End If
            If CVI(Left$(X$, 2)) < 1 Or ECOARCH$("MEPAGO", Left$(X$, 2)) = "" Then
               Call MENSERR(24, "Tipo de Valor no Admitido")
               GoTo 10
            End If
            SUVAPROP# = SUVAPROP# + IVAL#
         Next U&
         If Abs(SUVAPROP# - Val(Label20.Caption)) >= 0.005 Then
            Call MENSERR(24, "Imposible Cerrar - Diferencia de Valores")
            GoTo 10
         End If
         '
         Call CIERRARCH("!CARMEPA")
         Call SIGUEORPA4
         '
       Else
         '
         Call CIERRARCH("!CARMEPA")
         Call BLANARCH("!CARMEPA")
         Call SIGUEORPA4
         '
    End If
    '
End Sub
'

Private Sub Command1_Click()
    '
    Call SELECHO("PROVED1")
    NCXI% = Val(TRIM$(VALACT1$("PROVED1")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        '\\\OT-05-0619-WAR-16/09/05///
        Call VERIRECHA(NC%)
        '\\\OT-05-0619-WAR-FIN///
        Call SIGUEORPA1
    End If
    '
End Sub

Private Sub Command10_Click()
    Call SELECHO("ECODEP")
    DPX1% = Val(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("AMAPRO/NEW", "Alta de Nueva Cuenta de Proveedor")
    Call FRESHECHO("PROVED1")
    Command11.Enabled = True
End Sub

Private Sub Command2_Click()
    Screen.MousePointer = 11
    If List1.ListCount > 0 Or List2.ListCount > 0 Then
          Call CIERRARCH("*.*")
          Call BLANARCH("!CARGAST")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          Label21.Caption = ""
          BOTNEXT.Enabled = False
          Command5.Enabled = False
          Screen.MousePointer = 1
          EPAC$ = UCase$(EMPREAC$)
          If InStr(EPAC$, "NEUMAN") > 0 Then
            Text1.TEXT = ""
            Text2.TEXT = ""
            Text3.TEXT = ""
            Text4.TEXT = ""
            Text5.TEXT = ""
          End If
          Text1.SetFocus
        Else
          If TRIM$(CONTROL$("ORPASTA", "AVAHOJA")) = "NEWPAGE" Then
              'Printer.EndDoc
              Call SETSPOOL("ENDDOC", "")
          End If
          Call FINAL("")
          Screen.MousePointer = 1
    End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If List1.ListCount > 0 Or List2.ListCount > 0 Then
        Command2.ToolTipText = "Limpia y Refresca Pantalla"
      Else
        Command2.ToolTipText = "Cierra y Abandona la Aplicación"
   End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    Call CONECRUN("*NOTAPRO/" + NCLI$, "Notas de Proveedores")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    'If EXISTE%("COGECOM.EXE") < 1 Then
    '    Call OPNOIN("")
    '    Command4.Enabled = True
    '    Exit Sub
    'End If
    RESCUPRO.Text1.TEXT = Text1.TEXT
    PROLLAMA$ = Name
    Hide
    RESCUPRO.Show
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    TOTAVAL# = CDbl(Val(FORPASTA.Label14.Caption) + Val(FORPASTA.Label21.Caption))
    TOTACAN# = CDbl(Val(FORPASTA.Label18.Caption))
    If Abs(TOTACAN# - TOTAVAL#) >= 0.005 Then
        Call COMUNI("Operación Incorrecta.\  \Las Cancelaciones Difieren\de la Suma de Valores a Entregar.")
        Command5.Enabled = True
        Exit Sub
    End If
    '
    VACON% = 1       ' PARA QUE NO BLANQUEE LISTADOS
    Call GRAORPASTA(OKX%)
    If OKX% < 0 Then GoTo 10   ' abortó en observaciones
    If OKX% = 0 Then           ' RE-NUMERO ORDEN DE PAGO
       Command5.Enabled = True
       Exit Sub
    End If
    '
    If OKX% <> 7 Then
        Call COMUNI("La Registración de esta Orden de Pago\no Pudo Completarse. (" + TRIM$(Str$(OKX%)) + ")\  \Consulte a su Soporte de Sistemas FLEXOFT.")
        Call FINAL("")
    End If
    Call PRIORPASTA
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "SABO") > 0 Or InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "FERVI") > 0 Then
      Call PRICORETIB
    End If
    VACON% = 0
    '
10  FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    NROPRO& = 0
    Call NUORPA(NROPRO&)
    Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
    '
    Call CIERRARCH("!CARMEPA")
    Call BLANARCH("!CARMEPA")
    Call CIERRARCH("!CARGAPLI")
    Call BLANARCH("!CARGAPLI")
    Call CIERRARCH("!CARGAST")
    Call BLANARCH("!CARGAST")
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label18.Caption = ""
    Label20.Caption = ""
    Label21.Caption = ""
    'BOTNEXT.Enabled = True
    Command5.Enabled = False
    EPAC$ = UCase$(EMPREAC$)
    If InStr(EPAC$, "NEUMAN") > 0 Then
      Text1.TEXT = ""
      Text2.TEXT = ""
      Text3.TEXT = ""
      Text4.TEXT = ""
      Text5.TEXT = ""
    End If
    Text1.SetFocus
    '
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    OPORPA00.Show 1
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    Command7.Enabled = False
    NC% = Val(Text1.TEXT)
    If NC% < 1 Or ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text1.SetFocus
       GoTo 99
    End If
    '
    If TICUEPRO%((-1) * NC%) <> 0 Then
       Call MENSERR(24, "Proceso Imposible - Cuenta de Gastos")
       GoTo 99
    End If
    '
    If COMALTER%("El Proceso Invocado Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar") <> 2 Then
       GoTo 99
    End If
    '
    Call REAPLICREPRO(NC%)
    '
99  Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   If List1.ListCount > 0 Then
     Screen.MousePointer = 11
     MODOSORVAL% = (MODOSORVAL% + 1) Mod 2
     List4.Clear
     '
     For U& = 1 To List1.ListCount
       Z1$ = List1.List(U& - 1)
       VALI1 = FECHANUM(Mid$(Z1$, 49, 8))
       If MODOSORVAL% = 1 Then VALI1 = Val(Mid$(Z1$, 57, 12))
       For U4& = 1 To List4.ListCount
         Z4$ = List4.List(U4& - 1)
         VALI4 = FECHANUM(Mid$(Z4$, 49, 8))
         If MODOSORVAL% = 1 Then VALI4 = Val(Mid$(Z4$, 57, 12))
         If VALI4 > VALI1 Then
           List4.AddItem Z1$, U4& - 1
           GoTo 19
         End If
       Next U4&
       List4.AddItem Z1$
19   Next U&
     '
     List1.Clear
     For U& = 1 To FORPASTA.List4.ListCount
       List1.AddItem List4.List(U& - 1)
     Next U&
     List1.Refresh
     Screen.MousePointer = 1
     '
   End If
   Command8.Enabled = True
   '
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
    Show
    Refresh
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
10  FORIPRE$ = TRIM$(CONTROL$("", "FORPASTA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Orden de Pago.")
        OPORPASTA.Show 1
        GoTo 10
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Pago Standard:\  \'" + RMCC$ + "'.")
            OPORPASTA.Show 1
            GoTo 10
        End If
    End If
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    NROPRO& = 0
    Call NUORPA(NROPRO&)
    Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
    '
    Refresh
    Call RECUMOVI
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If TRIM$(CONTROL$("ORPASTA", "AVAHOJA")) = "NEWPAGE" Then
        'Printer.EndDoc
       Call SETSPOOL("ENDDOC", "")
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Label10_Change()
   Call CALRETIBRU   ' CALCULO RETENCION I.BRUTOS PCIA.BS.AS
   If Val(Label10.Caption) >= 0.005 Then
        BOTNEXT.Enabled = True
      Else
        BOTNEXT.Enabled = False
   End If
End Sub

Private Sub Label14_Change()
   Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - Val(Label21.Caption)), 4, 14, 2, 2)
   Label20.Refresh
End Sub

Private Sub Label14_DblClick()
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "SABO") < 1 Then
     If InStr(EPAC$, "DOSIVA") < 1 Then
       If InStr(EPAC$, "FERVI") < 1 Then
         Exit Sub
       End If
     End If
   End If
   '
   If Val(Label18) > 0 Then
     VDEF$ = MKD$(Val(Label23)) + MKS$(Val(Label26)) + MKS$(Val(Label27))
10   BASIMPON# = CVD(Left$(VDEF$, 8))
     IRET# = 0
     If BASIMPON# >= Val(Label28) Then
       IRET# = CVD(Left$(VDEF$, 8)) * CVS(Mid$(VDEF$, 9, 4)) / 100 * CVS(Mid$(VDEF$, 13, 4)) / 100
     End If
     VDEF$ = VDEF$ + MKD$(IRET#)
     OBX$ = OBJPLA$("CARIRETIB", VDEF$)
     If OBX$ = "" Then Exit Sub
     IRET1# = 0
     If CVD(Left$(OBX$, 8)) >= Val(Label28) Then
       IRET1# = CVD(Left$(OBX$, 8)) * CVS(Mid$(OBX$, 9, 4)) / 100 * CVS(Mid$(OBX$, 13, 4)) / 100
     End If
     If Abs(IRET# - IRET1#) > 0.0005 Then
       VDEF$ = Left$(OBX$, 16)
       GoTo 10
     End If
     '
     Label23 = TRIM$(CSTRING$(Left$(OBX$, 8), 3, 12, 2, 2))
     Label26 = TRIM$(CSTRING$(Mid$(OBX$, 9, 4), 3, 12, 2, 2))
     Label27 = TRIM$(CSTRING$(Mid$(OBX$, 13, 4), 3, 12, 2, 2))
     Label14 = TRIM$(CSTRING$(Mid$(OBX$, 17, 8), 3, 12, 2, 2))
     '
     If Val(Label14) > 0.005 Then
       If Val(Label24) < 1 Then
         Dim ALIBRU(8), MINIMPO#(8), MEPAGX%(8)
         For U& = 1 To ULTREG&("TACATIBP")
           TCX$ = REGLEIDO$("TACATIBP", U&)
           TCY% = Asc(Mid$(TCX$, 1, 1))
           If TCY% > 0 Then
             If TCY% <= 8 Then
               ALIBRU(TCY%) = CVS(Mid$(TCX$, 33, 4))
               MINIMPO#(TCY%) = CVS(Mid$(TCX$, 37, 4))
               MEPAGX%(TCY%) = CVI(Mid$(TCX$, 41, 2))
             End If
           End If
         Next U&
         Label24 = Str$(MEPAGX%(1))
       End If
     End If
     '
   End If
End Sub

Private Sub Label18_Change()
   If Val(Label18.Caption) > 0 Then
        BOTNEXT.Enabled = True
      Else
        BOTNEXT.Enabled = False
   End If
   Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - Val(Label21.Caption)), 4, 14, 2, 2)
   Label20.Refresh
End Sub

Private Sub Label2_DblClick()
    OPORPASTA.Show 1
End Sub

Private Sub Label20_Change()
   If Val(Label20.Caption) > 0 Then
        Command5.Enabled = False
      Else
        Command5.Enabled = True
        BOTNEXT.Enabled = False
   End If

End Sub

Private Sub Label21_Change()
   Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - Val(Label21.Caption)), 4, 14, 2, 2)
   Label20.Refresh
End Sub

Private Sub List1_Click()
    '
    If VACON% <> 0 Then
       Exit Sub
    End If
    '
    Z$ = List1.TEXT
    LIST1ANT$ = Z$
    ILISTA& = List1.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = "  " + Mid$(Z$, 57, 12)
    SUMAVAL# = Val(Label21.Caption)
    MAXAPLI# = Val(Label18.Caption) - Val(Label21.Caption) - Val(Label14.Caption)
    If IACRE$ = "" Then
         If Val(IPEN$) <= MAXAPLI# + 0.005 Then
             Call REPLA(Z$, IPEN$, 70, 14)
             SUMAVAL# = SUMAVAL# + CDbl(Val(IPEN$))
           Else
             Call MENSERR(24, "Importe Valor Excede Saldo a Cancelar")
         End If
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMAVAL# = SUMAVAL# - CDbl(Val(IACRE$))
    End If
    VACON% = 1
    List1.AddItem Z$, ILISTA&
    List1.RemoveItem ILISTA& + 1
    List1.Refresh
    List1.ListIndex = ILISTA&
    VACON% = 0
    '
    Label21.Caption = CSTRING$(MKD$(SUMAVAL#), 4, 14, 2, 2)
    Label21.Refresh
    Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - SUMAVAL#), 4, 14, 2, 2)
    Label20.Refresh
    '
End Sub

Private Sub List1_DBLClick()
    '
    'Call SIGUEORPA2
    '
End Sub

Private Sub List2_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    If TICUEPRO%((-1) * NC%) = 1 Then   ' cuenta de gastos
       Call SIGUEORPA5
       Exit Sub
    End If
    '
    Z$ = List2.TEXT
    LIST2ANT$ = Z$
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = " " + Mid$(Z$, 56, 13)
    SUMACRE# = Val(Label18.Caption)
    If IACRE$ = "" Then
         Call REPLA(Z$, IPEN$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(Val(IPEN$))
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(Val(IACRE$))
         If Left$(Z$, 20) = "Anticipo no Aplicado" Then
            Call REPLA(Z$, Space$(13), 32, 13)
            Call REPLA(Z$, Space$(13), 56, 13)
         End If
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
    Label18.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label18.Refresh
    Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - Val(Label21.Caption)), 4, 14, 2, 2)
    Label20.Refresh
    '
End Sub
'

Private Sub List2_DblClick()
    '
    If TICUEPRO%((-1) * NC%) = 1 Then   ' cuenta de gastos
       Call SIGUEORPA5
       Exit Sub
    End If
    '
    VACON% = 1
    ILIST2& = List2.ListIndex
    TOLIST2& = List2.TopIndex
    '
    Z$ = List2.TEXT
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = Mid$(Z$, 56, 14)
    SUMACRE# = Val(Label18.Caption)
    SUMACRE# = SUMACRE# - CDbl(Val(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(Val(Mid$(LIST2ANT$, 70, 14)))
    List2.Visible = False
    List2.AddItem LIST2ANT$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    '
    Call SIGUEORPA3
    '
    VACON% = 1
    List2.TopIndex = TOLIST2&
    List2.Visible = True
    If List2.ListCount > ILIST2& Then
          List2.ListIndex = ILIST2&
        Else
          Call CIERRARCH("!CARMEPA")
          Call BLANARCH("!CARMEPA")
          Call CIERRARCH("!CARGAPLI")
          Call BLANARCH("!CARGAPLI")
          Call CIERRARCH("!CARGAST")
          Call BLANARCH("!CARGAST")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          Label21.Caption = ""
          'BOTNEXT.Enabled = True
          Command5.Enabled = False
          EPAC$ = UCase$(EMPREAC$)
          If InStr(EPAC$, "NEUMAN") > 0 Then
            Text1.TEXT = ""
            Text2.TEXT = ""
            Text3.TEXT = ""
            Text4.TEXT = ""
            Text5.TEXT = ""
          End If
          Text1.SetFocus
    End If
    VACON% = 0
    '
End Sub

Private Sub Option2_Click()
    Hide
    FORLIFOF.Show
    Call FRESHALL
    Call LIFOFIJ1
End Sub

Private Sub Text1_Change()
    Call CIERRARCH("!CARGAST")
    Call BLANARCH("!CARGAST")
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label18.Caption = ""
    Label20.Caption = ""
    Label21.Caption = ""
    Label23 = ""
    Label24 = ""
    Label25 = ""
    Label26 = ""
    Label27 = ""
    'BOTNEXT.Enabled = True
    Command5.Enabled = False
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("PROVED1")
    NCXI% = Val(TRIM$(VALACT1$("PROVED1")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        '\\\OT-05-0619-WAR-19/09/05///
        Call VERIRECHA(NC%)
        '\\\OT-05-0619-WAR-FIN///
        Call SIGUEORPA1
    End If
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
        FORPASTA.Refresh
        CTXX% = 1
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Val(Text1.TEXT) > 0 Then
          If Val(Text1.TEXT) <= 32767 Then
             NC% = Val(Text1.TEXT)
             If ECOARCH$("PROVED1", MKI$(NC%)) <> "" Then
                Text2.TEXT = RASOCLI$((-1) * NC%)
                Text3.TEXT = DOMICLI$((-1) * NC%)
                Text4.TEXT = CPOSCLI$((-1) * NC%)
                Text5.TEXT = LOCACLI$((-1) * NC%)
                '\\\OT-05-0619-WAR-19/09/05///
                Call VERIRECHA(NC%)
                '\\\OT-05-0619-WAR-FIN///
                Call SIGUEORPA1
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

Private Sub Text11_LostFocus()
    '
    VDEFF$ = TRIM$(Text11.TEXT)
    Text11.TEXT = VDEFF$
    If Left$(VDEFF$, 3) <> "OP." Then GoTo 90
    NRONUE& = Val(Mid$(VDEFF$, 4))
    NROPRO& = NRONUE&
    Call NUORPA(NROPRO&)
    If NROPRO& <> NRONUE& Then
        Beep
        Call MENSERR(24, "Número no Válido o Pre-existente")
        GoTo 95
    End If
    Exit Sub
    '
90  Beep
    Call MENSERR(24, "Número no Válido")
    NROPRO& = 0
    Call NUORPA(NROPRO&)
95  Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
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
    OPORPASTA.Show 1
End Sub

Private Sub Text7_GotFocus()
    Text10.SetFocus
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Text9_GotFocus()
   Screen.MousePointer = 1
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

Sub RECUMOVI()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
     If ULTREG&("BADACHE") < 64 Then
       Screen.MousePointer = 11
       ARECU$ = "AJUSTOCK": GoSub 10
       ARECU$ = "APLICOB": GoSub 10
       ARECU$ = "APLIPAG": GoSub 10
       ARECU$ = "ASIDIAR": GoSub 10
       ARECU$ = "BADACHE": GoSub 10
       ARECU$ = "ACCDIR": GoSub 10
       Call CIERRARCH("*.*")
       Screen.MousePointer = 1
     End If
   End If
   Exit Sub
   '
10 ULO$ = Left$(CurDir, 2)
   LOREG% = RECLEN%(ARECU$)
   If EXISTE%(ULO$ + "\DATOS\" + ARECU$ + ".DAT") = 1 Then
     N2% = FILEOPEN%(ULO$ + "\DATOS\" + ARECU$ + ".DAT", 4, 128)
     RELEI$ = String$(LOREG%, 0)
     UREA& = ULTREG&(ARECU$)
     For U& = 1 To LOF(N2%) / LOREG%
       POLEC& = (U& - 1) * LOREG% + 1
       Get #N2%, POLEC&, RELEI$
       '
       If ARECU$ = "BADACHE" Then
         For JJ& = 1 To UREA&
           XX$ = REGLEIDO$(ARECU$, JJ&)
           If Left$(XX$, 102) = Left$(RELEI$, 102) Then
             Call GRAREG(ARECU$, RELEI$, JJ&)
             GoTo 99
           End If
         Next JJ&
       End If
       '
       Call REGAPP(ARECU$, RELEI$)
99   Next U&
     Call CIERRARCH(ARECU$)
     Close #N2%: N2% = 0
   End If
   Return
   '
End Sub

Sub CALRETIBRU_vie()
   '
   Static CTCALCU%, ALIBRU(8), MINIMPO#(8), MEPAGX%(8)
   If CTCALCU% = 0 Then
     CTCALCU% = (-1)
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "SABO") > 0 Or InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "FERVI") > 0 Then
       CTCALCU% = 1
       For U& = 1 To ULTREG&("TACATIBP")
         TCX$ = REGLEIDO$("TACATIBP", U&)
         TCY% = Asc(Mid$(TCX$, 1, 1))
         If TCY% > 0 Then
           If TCY% <= 8 Then
             ALIBRU(TCY%) = CVS(Mid$(TCX$, 33, 4))
             MINIMPO#(TCY%) = CVS(Mid$(TCX$, 37, 4))
             MEPAGX%(TCY%) = CVI(Mid$(TCX$, 41, 2))
           End If
         End If
       Next U&
       Call CIERRARCH("TACATIBP")
     End If
   End If
   '
   If CTCALCU% < 0 Then Exit Sub
   '
   If FECHANUM(Text9) < FECHANUM("01/08/03") Then Exit Sub
   '
   ' determinar categoria de I.Brutos
   NPROVEI% = Val(Text1)
   RERET$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(NPROVEI%))
   CATIB1% = Asc(Mid$(RERET$, 3, 1))
   PORCONVE = CVS(Mid$(RERET$, 5, 4))
   If CATIB1% < 1 Then Exit Sub  ' exento
   '
   ' determinar monto imponible
   BASEIMTO# = 0
   DIVIAN = 1
   If PIVACLI%((-1) * NPROVEI%) = 1 Then
     DIVIAN = 1 + ALIVA / 100
   End If
   For UKI% = 1 To List2.ListCount
     Z$ = List2.List(UKI% - 1)
     IPARCI# = Val(Mid$(Z$, 70, 14))
     If IPARCI# >= 0.005 Then
       BASEIMPAR# = IPARCI# / DIVIAN
       RECUECO& = Val(Mid$(Z$, 119, 8))
       If RECUECO& > 0 Then
         X$ = REGLEIDO$("CUECORR", RECUECO&)
         IGRAEX# = Abs(CVD(Mid$(X$, 51, 8)) + CVD(Mid$(X$, 59, 8)))
         ITOTAX# = Abs(CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8)))
         COEFI = 1
         If ITOTAX# > 0.005 And IGRAEX# <= ITOTAX# Then
           COEFI = IGRAEX# / ITOTAX#
         End If
         BASEIMPAR# = IPARCI# * COEFI
       End If
       BASEIMTO# = BASEIMTO# + BASEIMPAR#
     End If
   Next UKI%
   '
   Label26 = "": Label27 = ""
   IRETIB# = 0
   If CATIB1% <= 8 Then
     Label28 = FORMATNUM$(MINIMPO#(CATIB1%), "F12.2")
     If BASEIMTO# >= MINIMPO#(CATIB1%) Then
       IRETIB# = BASEIMTO# * PORCONVE / 100 * ALIBRU(CATIB1%) / 100
       Label24 = Str$(MEPAGX%(CATIB1%))
       Label26 = FORMATNUM$(ALIBRU(CATIB1%), "F8.2")
       Label27 = FORMATNUM$(PORCONVE, "F8.2")
     End If
   End If
   '
   Label14 = FORMATNUM$(IRETIB#, "F12.2")
   Label23 = FORMATNUM$(BASEIMTO#, "F12.2")
   '
End Sub

Sub CALRETIBRU()
   '
   'Static CTCALCU%, ALIBRU(8), MINIMPO#(8), MEPAGX%(8)
   Static NPROVEA%, NPROVEB%, NOPAGAN$, MESANANT$, CATIB1%, PORCONVE
   Static ACUBASANT#, ACURETANT#
   '
   If CTCALCU% = 0 Then
     CTCALCU% = 1
   End If
   '
   If CTCALCU% < 0 Then
      Label23 = FORMATNUM$(NETOCANCE#, "F12.2")
      Exit Sub
   End If
   '
   If FECHANUM(Text9) < FECHANUM("01/08/03") Then Exit Sub
   '
   ' determinar categoria de I.Brutos
   NPROVEI% = Val(Text1)
   If TICUEPRO%((-1) * NPROVEI%) = 1 Then GoTo 99 ' RETENCION CERO EN CUENTAS DE GASTOS
   RERET$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(NPROVEI%))
   CATIB1% = Asc(Mid$(RERET$, 3, 1))
   PORCONVE = CVS(Mid$(RERET$, 5, 4))
   If CATIB1% < 1 Then Exit Sub  ' exento
   '
   ' determinar monto imponible
   BASEIMTO# = 0
   DIVIAN = 1
   If PIVACLI%((-1) * NPROVEI%) = 1 Then
     DIVIAN = 1 + ALIVA / 100
   End If
   For UKI% = 1 To List2.ListCount
     Z$ = List2.List(UKI% - 1)
     IPARCI# = Val(Mid$(Z$, 70, 14))
     If IPARCI# >= 0.005 Then
       BASEIMPAR# = IPARCI# / DIVIAN
       RECUECO& = Val(Mid$(Z$, 119, 8))
       If RECUECO& > 0 Then
         X$ = REGLEIDO$("CUECORR", RECUECO&)
         IGRAEX# = Abs(CVD(Mid$(X$, 51, 8)) + CVD(Mid$(X$, 59, 8)))
         ITOTAX# = Abs(CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8)))
         COEFI = 1
         If ITOTAX# > 0.005 And IGRAEX# <= ITOTAX# Then
           COEFI = IGRAEX# / ITOTAX#
         End If
         BASEIMPAR# = IPARCI# * COEFI
       End If
       BASEIMTO# = BASEIMTO# + BASEIMPAR#
     End If
   Next UKI%
   '
   Label26 = "": Label27 = ""
   IRETIB# = 0
   If CATIB1% <= 8 Then
     PORCEIBRU# = ALIRETBRU#(NPROVEI%, MINOIMPO#, MEDIOPAGO%)
     Label28 = FORMATNUM$(MINOIMPO#, "F12.2")
     If BASEIMTO# >= MINOIMPO# Then
       IRETIB# = BASEIMTO# * PORCEIBRU# / 100
       Label24 = Str$(MEDIOPAGO%)
       Label26 = FORMATNUM$(PORCEIBRU#, "F8.2")
       Label27 = FORMATNUM$(PORCONVE, "F8.2")
     End If
   End If
   '
   Label14 = FORMATNUM$(IRETIB#, "F12.2")
   Label23 = FORMATNUM$(BASEIMTO#, "F12.2")
   '
99 End Sub

Sub PRICORETIB()
   '
   If Val(Label14) >= 0.005 Then
      NURETBA& = Val(Label25)
      Screen.MousePointer = 11
      FECHDOC$ = FETEXTO$(Int(FECHANUM(Text9.TEXT)))
      NUMEDOC$ = TRIM$(Str$(NURETBA&))
      While Len(NUMEDOC$) < 8
         NUMEDOC$ = "0" + NUMEDOC$
      Wend
      NUMEDOC$ = "0001-" + NUMEDOC$
      TIPODOC$ = "Retencion I.Brutos"
      '
      CODPARAM$(1) = "IMPO-CHE"
      CODPARAM$(2) = "IMPO-PEN"
      CODPARAM$(3) = "IMPO-CAN"
      CODPARAM$(4) = "POR-DESC"
      CAPARDOC% = 4
      '
      DOCPARAM$(1) = TRIM$(Label23)
      DOCPARAM$(2) = TRIM$(Label26)
      DOCPARAM$(3) = TRIM$(Label14)
      DOCPARAM$(4) = TRIM$(Label27)
      '
      CACOLTAB1% = 0
      '
      DESTIDOC% = (-1) * Val(Text1)
      CAITAB1% = 0
      '
      Call PRINTDOC("FORRETIB")   ' Retencion I.Brutos
      '
   End If
   '
End Sub
'\\\OT-05-0619-WAR-16/09/05///
Sub VERIRECHA(NPRO%)
  '
  Call BLANARCH("!INDIVARE")
  For J& = 1 To ULTREG&("RECHLIN")
    X$ = REGLEIDO$("RECHLIN", J&)
    NPRO2% = CVI(Mid$(X$, 9, 2))
    If NPRO2% = NPRO% Then
      FE2% = CVI(Mid$(X$, 113, 2))
      USUPRO% = CVI(Mid$(X$, 115, 2))
      If FE2% <= 0 Then 'SI NO TIENE FECHA DE PROCESADO
        If USUPRO% <= 0 Then 'SI NO TIENE USUARIO DE PROCESADO
          'AGREGA LA LISTA DE SELECCION
          NROREC& = CVS(Mid$(X$, 1, 4))
          FEMI1$ = FECHATEX$(CVI(Mid$(X$, 7, 2)))
          REFERE2$ = FEMI1$ + " - " + Left$(RASOCLI$(NPRO% * (-1)), 30)
          'CODIGO VIEJO
'          ZZ$ = Space(64)
'          Call REPLA(ZZ$, MKS$(NROREC&), 1, 4)
'          Call REPLA(ZZ$, REFERE2$, 5, 60)
'          Call REGAPP("!INDIVARE", ZZ$)
          '\\\OT-05-0665-WAR-07/10/05///
          ZZ$ = Space(64)
          Call REPLA(ZZ$, MKS$(NROREC&), 1, 4)
             IMPORE# = CVD(Mid$(X$, 105, 8))
             IMPOR1$ = FORMATNUM$(IMPORE#, "F10.2")
          Call REPLA(ZZ$, REFERE2$ + IMPOR1$, 5, 60)
          Call REGAPP("!INDIVARE", ZZ$)
          '\\\OT-05-0665-WAR-FIN///
          HAYPEND% = 1
        End If
      End If
    End If
  Next J&
  Call CIERRARCH("!INDIVARE")
  If HAYPEND% = 0 Then Exit Sub 'NO HAY RECHAZOS PENDIENTES DE PROCESO
  Call FRESHECHO("!INDIVARE")
  '
  Call SELECHO("!INDIVARE")
  NPX& = Val(VALACT1$("!INDIVARE"))
  '
  If NPX& > 0 Then
    Call MUESTRARECH(NPX&)
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = " " + NUDOCU$ + " "
    If COMALTER%("¿Desea Procesar el Rechazo en Línea Número: " & NUDOCU$ & " ?\\Si, Procesar.\No, Continuar. ") = 1 Then
      If DERACCE%("PROCEREC", "Procesar Rechazos en Línea.") = 2 Then
        Call PROCERECH(NPX&) 'PROCESA RECHAZO EN LÍNEA
      End If
    End If
  End If
  '
99 End Sub
'\\\OT-05-0619-WAR-FIN//
'\\\OT-05-0619-WAR-19/09/05//
Sub PROCERECH(NPX&)
  '
  NPRX% = XVALO(Text1)
  RFF$ = RECFILT$("RECHLIN", 1, MKS$(NPX&))
  For KU& = 1 To Len(RFF$) Step 4
     REGO& = CVS(Mid$(RFF$, KU&, 4))
     XXX$ = REGLEIDO$("RECHLIN", REGO&)
     If CVI(Mid$(XXX$, 9, 2)) = NPRX% Then
       FEX1% = HOY(HO$)
       Call REPLA(XXX$, MKI$(USNBR%), 115, 2)
       Call REPLA(XXX$, MKI$(FEX1%), 113, 2)
       Call GRAREG("RECHLIN", XXX$, REGO&)
     End If
  Next KU&
  '
End Sub
'
Sub MUESTRARECH(NPX&)
 '
    TIDOCUM$ = "Rechazo en Línea"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = " " + NUDOCU$ + " "
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
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
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
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
99 End Sub
'\\\OT-05-0619-WAR-FIN//
