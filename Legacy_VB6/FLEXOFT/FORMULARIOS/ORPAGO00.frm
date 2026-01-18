VERSION 5.00
Begin VB.Form ORPAGO00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Pago (Operaciones Comerciales de Compras)"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFD0&
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
      Top             =   -30
      Width           =   12000
      Begin VB.ListBox LINDICE 
         Height          =   1035
         Left            =   1155
         Sorted          =   -1  'True
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   1170
      End
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
         Picture         =   "ORPAGO00.frx":0000
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
         BackColor       =   &H00E0FFD0&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   540
         Left            =   6825
         TabIndex        =   52
         Top             =   140
         Width           =   3900
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0FFD0&
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
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFD0&
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
         Height          =   750
         Left            =   10920
         Picture         =   "ORPAGO00.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Configuración de Funcionamiento"
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
         Picture         =   "ORPAGO00.frx":074C
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
         Height          =   3120
         Left            =   210
         ScaleHeight     =   3060
         ScaleWidth      =   10395
         TabIndex        =   22
         Top             =   4250
         Width           =   10450
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C000&
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
               Height          =   2310
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
               Left            =   8295
               TabIndex        =   31
               Top             =   2415
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
               Y1              =   210
               Y2              =   2730
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
               Top             =   2415
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
               Top             =   2415
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
               Left            =   30
               TabIndex        =   25
               Top             =   2415
               Width           =   1695
            End
            Begin VB.Line Line14 
               BorderWidth     =   2
               X1              =   0
               X2              =   10395
               Y1              =   2310
               Y2              =   2310
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   10395
               Y1              =   2730
               Y2              =   2730
            End
            Begin VB.Line Line16 
               X1              =   2520
               X2              =   2520
               Y1              =   2310
               Y2              =   2730
            End
            Begin VB.Line Line17 
               X1              =   3735
               X2              =   3735
               Y1              =   2310
               Y2              =   2730
            End
            Begin VB.Line Line18 
               X1              =   5460
               X2              =   5460
               Y1              =   2310
               Y2              =   2730
            End
            Begin VB.Line Line19 
               X1              =   6720
               X2              =   6720
               Y1              =   2310
               Y2              =   2730
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C000&
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
         Height          =   960
         Left            =   10920
         Picture         =   "ORPAGO00.frx":0896
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
         Picture         =   "ORPAGO00.frx":0F00
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
         Picture         =   "ORPAGO00.frx":120A
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
         Picture         =   "ORPAGO00.frx":1514
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
         Picture         =   "ORPAGO00.frx":181E
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
         Picture         =   "ORPAGO00.frx":1B28
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
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
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
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
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
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
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
      Begin VB.Line Line21 
         X1              =   210
         X2              =   10660
         Y1              =   4110
         Y2              =   4110
      End
      Begin VB.Line Line12 
         X1              =   10665
         X2              =   10660
         Y1              =   4230
         Y2              =   7330
      End
      Begin VB.Line Line8 
         X1              =   10660
         X2              =   10660
         Y1              =   1680
         Y2              =   4095
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
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4305
         TabIndex        =   48
         Top             =   7560
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
         Top             =   7875
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
         Top             =   7875
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
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6825
         TabIndex        =   43
         Top             =   7560
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
         Top             =   7875
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
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2205
         TabIndex        =   41
         Top             =   7560
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
         Top             =   7875
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
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -945
         TabIndex        =   37
         Top             =   7560
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
Attribute VB_Name = "ORPAGO00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$
Dim MODOSORVAL%
Dim YACARFO%

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
        Call SIGUEORPA1
    End If
    '
End Sub


Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCH_GES/AMAPRO", "Archivos Maestros")
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
              Call SETSPOOL("ENDDOC", "")
          End If
          Hide
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
    RESCUPRO.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    TOTAVAL# = CDbl(Val(Label14.Caption) + Val(Label21.Caption))
    TOTACAN# = CDbl(Val(Label18.Caption))
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
    '
    NROPRO& = 0
    Call NUORPA(NROPRO&)
    Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
    '
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    STORPA00.Show 1
    Command7.Enabled = True
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
     For U& = 1 To List4.ListCount
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

Private Sub Form_Activate()
   '
   If YACARFO% < 1 Then
      '
      YACARFO% = 1
      '
10    FORIPRE$ = TRIM$(CONTROL$("", "FORPASTA"))
      If FORIPRE$ = "" Then
          Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Orden de Pago.")
          STORPA00.Show 1
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          Command5.Enabled = False
          GoTo 10
      End If
      '
      If FORIPRE$ <> "" Then
          LU$ = LUDAT$
          If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
              RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
              Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Pago Standard:\  \'" + RMCC$ + "'.")
              STORPA00.Show 1
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
   End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call CENTRAFORM(Me)
    Refresh
    DoEvents
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If TRIM$(CONTROL$("ORPASTA", "AVAHOJA")) = "NEWPAGE" Then
       Call SETSPOOL("ENDDOC", "")
    End If
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    YACARFO% = 0
End Sub

Private Sub Label10_Change()
   If Val(Label10.Caption) >= 0.005 Then
        BOTNEXT.Enabled = True
      Else
        BOTNEXT.Enabled = False
   End If
End Sub

Private Sub Label18_Change()
   If Val(Label18.Caption) > 0 Then
        BOTNEXT.Enabled = True
      Else
        BOTNEXT.Enabled = False
   End If
End Sub

Private Sub Label2_DblClick()
    STORPA00.Show 1
End Sub

Private Sub Label20_Change()
   If Val(Label20.Caption) > 0 Then
        Command5.Enabled = False
      Else
        Command5.Enabled = True
        BOTNEXT.Enabled = False
   End If

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

Private Sub List1_DblClick()
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
    Stop 'Call LIFOFIJ1
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
    'BOTNEXT.Enabled = True
    Command5.Enabled = False
End Sub

Private Sub Text1_DBLCLICK()
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
        Call SIGUEORPA1
    End If
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
        Refresh
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

Private Sub Text4_GOTFOCUS()
    Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text6_Click()
    Text1.SetFocus
End Sub

Private Sub Text6_DblClick()
    STORPA00.Show 1
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

Sub NUORPA(NROPRO&)
    '
    Screen.MousePointer = 11
    '
    FENUE$ = CONTROL$("", "FENUORPA")
    FLIM% = FECHANUM(FENUE$)
    If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
    '
    CAMOLE% = 0
    NROMA& = 0
    For U& = ULTREG&("CUECORR") To 1 Step -1
        XX$ = REGLEIDO$("CUECORR", U&)
        CAMOLE% = CAMOLE% + 1
        If CAMOLE% > 512 Then GoTo 90
        '
        If Mid$(XX$, 7, 2) = "OP" Then
           FEMO% = CVI(Mid$(XX$, 1, 2))
           If FEMO% >= FLIM% Then
              NROLE& = Val(Mid$(XX$, 18, 7))
              If NROLE& = NROPRO& Then
                If Mid$(XX$, 12, 4) <> "anul" Then
                  NROPRO& = 0
                End If
              End If
              If NROLE& > NROMA& Then NROMA& = NROLE&
           End If
        End If
    Next U&
    '
90  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
    '
End Sub

Sub SIGUEORPA1()
    '
    Screen.MousePointer = 11
    If MODORPA% = 0 Then
        MODORPA% = 1       ' por emision
        If CONTROL$("ORPASTA", "MODORDE") = "VENCIM" Then
           MODORPA% = 2
        End If
    End If
    '
    NCXI = Val(Text1.TEXT)
    If NCXI < 1 Or NCXI > 32767 Then
        Call MENSERR(24, "Numero de Cuenta no Válido")
        Text1.SetFocus
        Exit Sub
    End If
    '
    NC% = NCXI
    If ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
        Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Proveedores")
        Text1.SetFocus
        Exit Sub
    End If
    '
    FE$ = TRIM$(Text9.TEXT)
    FF% = FECHANUM(FE$)
    HOI% = HOY(HO$)
    If FF% < 1 Or FF% > HOI% + 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
    '
    If DIENTRE%(FF%, HOI%) > 55 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
    '
    FECHALIM% = Val(CONTROL$("", "CIEMECOM"))
    If FF% <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        Text9.SetFocus
        Exit Sub
    End If
     '
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 11
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then
        Text2.TEXT = RASOCLI$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        'Text7.TEXT = TEL1CLI$(NC%)
        'Text10.TEXT = CSTRING$(MKD$(SALDCLI#(NC%)), 4, 13, 2, 2)
        Refresh
    End If
    '
100 List1.Clear
    List1.Refresh
    List2.Clear
    List2.Refresh
    '
    Z$ = Space$(128)
    Call REPLA(Z$, "Anticipo no Aplicado", 1, 20)
    Call REPLA(Z$, HO$, 23, 8)
    Call REPLA(Z$, "       0", 119, 8)
    Call REPLA(Z$, " 0", 127, 2)
    List2.AddItem Z$
    SUMADEU# = 0: SUMAFAC# = 0
    '
    If TICUEPRO%((-1) * NC%) = 1 Then        ' GASTOS
       Call BLOQARCH("!CARGAST")
       Call BLANARCH("!CARGAST")
       GoTo 200
    End If
    '
    MAXDEU& = ULTREG&("CUECORR")
    RFD$ = RECFILT$("CUECORR", 4, MKI$(NC%))
    '
    LINDICE.Clear
    For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("CUECORR", REG&)
        If CVI(Mid$(X$, 5, 2)) <> NC% Then
            X$ = REGLEIDO$("CUECORR", 1)
            Call GRAREG("CUECORR", X$, 1)
            Call CIERRARCH("CUECORR")
            GoTo 100
        End If
        '
        FEMI = CVI(Left$(X$, 2))
        IMPO# = CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8))
        SALDEU$ = Mid$(X$, 139, 60)
        For KKI% = 1 To 6
           FEVE = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
           SALDO# = CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8))
           If Abs(SALDO#) >= 0.005 Then
              If MODORPA% = 1 Then
                  FF% = FEMI      ' orden emision
                Else
                  FF% = FEVE      ' orden vencimiento
              End If
              Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2) + CSTRING$(MKI$(KKI%), 1, 4, 0, 2)
              LINDICE.AddItem Z$
           End If
        Next KKI%
    Next U&
    '
    SUMADEU# = 0: SUMAFAC# = 0
    For U& = 1 To LINDICE.ListCount
        LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(LINDICE.TEXT, 7, 8))
        KKI% = Val(Mid$(LINDICE.TEXT, 15, 4))
        If REG& > 0 Then
          If REG& <= MAXDEU& Then
            If KKI% > 0 Then
              If KKI% <= 6 Then
                X$ = REGLEIDO$("CUECORR", REG&)
                TIX$ = Mid$(X$, 7, 2)
                SG% = (-1)
                TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                                    If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
                                    If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
                                    If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                                    If TIX$ = "CO" Then TIPOS$ = "Comprobante"
                                    If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
                                    If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
                                    If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                                    If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
                                    '
                IMPO# = CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8))
                SALDO# = CVD(Mid$(X$, 139 + 10 * KKI% - 8, 8))
                '
                If TIX$ = "FC" Or TIX$ = "ND" Or TIX$ = "TF" Or TIX$ = "RH" Or TIX$ = "CO" Or (TIX$ = "AS" And IMPO# < 0) Or (TIX$ = "OP" And Mid$(X$, 12, 4) = "anul") Then
                  If SALDO# >= 0.005 Then
                    '
                    SUMAFAC# = SUMAFAC# + IMPO#
                    IMPO# = Abs(IMPO#)
                    SUMADEU# = SUMADEU# + SALDO#
                    '
                    NROX$ = Mid$(X$, 7, 18)
                    FEMI = CVI(Mid$(X$, 1, 2))
                    FEVE = CVI(Mid$(X$, 139 + 10 * KKI% - 10, 2))
                    '
                    IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 13, 2, 2)
                    SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 13, 2, 2)
                    REGTEX$ = CSTRING$(MKS$(REG&), 3, 8, 0, 2)
                    IRETEX$ = CSTRING$(MKI$(KKI%), 1, 2, 0, 2)
                    '
                    Z$ = Space$(128)
                    Call REPLA(Z$, NROX$, 1, 18)
                    Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
                    Call REPLA(Z$, IMPTEX$, 32, 13)
                    Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
                    Call REPLA(Z$, SALTEX$, 56, 13)
                    Call REPLA(Z$, REGTEX$, 119, 8)
                    Call REPLA(Z$, IRETEX$, 127, 2)
                    '
                    List2.AddItem Z$
                    '
                  End If
                End If
              End If
            End If
          End If
        End If
    Next U&
    '
200 List2.ListIndex = -1
    List2.Refresh
    Label8.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label8.Refresh
    Label10.Caption = ""
    Label10.Refresh
    List2.Visible = True
    List2.SetFocus
    If List2.ListCount > 0 Then
        VACON% = 1
        List2.ListIndex = 0
        VACON% = 0
    End If
    '
    Screen.MousePointer = 11
    List1.Clear
    '
    LINDICE.Clear
    If TRIM$(CONTROL$("ORPASTA", "UTICHECA")) = "SI" Then
      RFF$ = RECFILT$("CAJABANC", 16, Chr$(1))
      FIN& = Len(RFF$)
      For I& = 1 To FIN& Step 4
        Call TRACE(20, I&, FIN&)
        RECHE& = CVS(Mid$(RFF$, I&, 4))
        X$ = REGLEIDO$("CAJABANC", RECHE&)
        '
        If UCase$(Mid$(X$, 69, 2)) <> "CH" Then  ' no es un cheque
          GoTo 19
        End If
        '
        If CVI(Mid$(X$, 269, 2)) > 0 Then ' esta anulado
          Call REPLA(X$, Chr$(2), 121, 1)
          Call GRAREG("CAJABANC", X$, RECHE&)
          GoTo 19
        End If
        '
        If CVI(Mid$(X$, 275, 2)) > 0 Then ' salio de cartera
          Call REPLA(X$, Chr$(2), 121, 1)
          Call GRAREG("CAJABANC", X$, RECHE&)
          GoTo 19
        End If
        '
        FF% = CVI(Mid$(X$, 63, 2))
        ICHE# = CVD(Mid$(X$, 153, 8))
        If MODOSORVAL% = 0 Then
             Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(RECHE&), 4, 8, 0, 2) + CSTRING$(MKI$(KKI%), 1, 4, 0, 2)
           Else
             Z$ = CSTRING$(MKD$(ICHE#), 4, 6, 0, 2) + CSTRING$(MKS$(RECHE&), 4, 8, 0, 2) + CSTRING$(MKI$(KKI%), 1, 4, 0, 2)
         End If
         LINDICE.AddItem Z$
         '
19     Next I&
    End If
    '
    For U& = 1 To LINDICE.ListCount
      LINDICE.ListIndex = U& - 1
      REG& = Val(Mid$(LINDICE.TEXT, 7, 8))
      X$ = REGLEIDO$("CAJABANC", REG&)
      If CVI(Mid$(X$, 275, 2)) < 1 Then
         ICHE# = CVD(Mid$(X$, 153, 8))
         TIVA$ = TRIM$(Mid$(X$, 71, 18))
         NUCHE$ = TRIM$(Mid$(X$, 89, 14))
         Z$ = Space$(128)
         Call REPLA(Z$, "Ch.3", 1, 4)
         If Left$(UCase$(TIVA$), 3) = "DOC" Then
            Call REPLA(Z$, "Doc3", 1, 4)
         End If
         Call REPLA(Z$, TIVA$, 7, 18)
         Call REPLA(Z$, NUCHE$, 25, 14)
         FEMI = CVI(Mid$(X$, 111, 2))
         Call REPLA(Z$, FECHATEX$(FEMI), 39, 8)
         FEVE = CVI(Mid$(X$, 63, 2))
         Call REPLA(Z$, FECHATEX$(FEVE), 49, 8)
         Call REPLA(Z$, CSTRING$(MKD$(ICHE#), 3, 12, 2, 2), 57, 12)
         Call REPLA(Z$, CSTRING$(MKS$(REG&), 3, 8, 0, 2), 121, 8)
         List1.AddItem Z$
      End If
    Next U&
    '
    List1.Refresh
    Screen.MousePointer = 1
    '
    If TICUEPRO%((-1) * NC%) = 1 Then        ' GASTOS
       Call SIGUEORPA5
    End If
    '
End Sub
'
Sub SIGUEORPA3()
    '
    If List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!CARGAPLI")
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        NROX$ = Mid$(Z$, 1, 20)
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(Val(Mid$(Z$, 32, 13)))
        FEVEI% = FECHANUM(Mid$(Z$, 47, 8))
        SALD# = CDbl(Val(Mid$(Z$, 56, 13)))
        ACRE# = CDbl(Val(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 1, 18)
        REG2& = Val(Mid$(Z$, 119, 8))
        IRE% = Val(Mid$(Z$, 127, 2))
        '
        Y$ = REGBLAN$("!CARGAPLI")
        Call REPLA(Y$, NROX$, 1, 20)
        Call REPLA(Y$, MKI$(FEMII%), 21, 2)
        Call REPLA(Y$, MKD$(IPED#), 23, 8)
        Call REPLA(Y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(Y$, MKD$(SALD#), 33, 8)
        Call REPLA(Y$, MKD$(ACRE#), 41, 8)
        Call REPLA(Y$, DOCU$, 49, 18)
        Call REPLA(Y$, MKS$(REG2&), 67, 4)
        Call REPLA(Y$, MKI$(IRE%), 71, 2)
        '
        Call REGAPP("!CARGAPLI", Y$)
    Next U&
    Call CIERRARCH("!CARGAPLI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Pagos"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    '
    VTB% = VENTABU%("CARGAPLI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        List2.Clear
        For U& = 1 To ULTREG&("!CARGAPLI")
            Y$ = REGLEIDO$("!CARGAPLI", U&)
            NROX$ = Mid$(Y$, 1, 20)
            If TRIM$(NROX$) = "Anticipo no Aplicado" Then
               HOI% = HOY(HO$)
               ACRE# = CVD(Mid$(Y$, 41, 8))
               Call REPLA(Y$, MKI$(HOI%), 31, 2)
               Call REPLA(Y$, MKD$(ACRE#), 23, 8)
               Call REPLA(Y$, MKD$(ACRE#), 33, 8)
               Call GRAREG("!CARGAPLI", Y$, U&)
            End If
            FEMI = CVI(Mid$(Y$, 21, 2))
            IPED# = CVD(Mid$(Y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 13, 2, 2)
            FEVE = CVI(Mid$(Y$, 31, 2))
            SALD# = CVD(Mid$(Y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 13, 2, 2)
            ACRE# = CVD(Mid$(Y$, 41, 8))
            '
            If ACRE# < 0 Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If U& > 1 Then
              If ACRE# > SALD# Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                 GoTo 200
              End If
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(Y$, 49, 18)
            REG2& = CVS(Mid$(Y$, 67, 4))
            IRE% = CVI(Mid$(Y$, 71, 2))
            REGTEX$ = CSTRING$(MKS$(REG2&), 3, 8, 0, 2)
            IRETEX$ = CSTRING$(MKI$(IRE%), 1, 2, 0, 2)
            '
            Z$ = Space$(128)
            Call REPLA(Z$, NROX$, 1, 20)
            Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(Z$, IMPTEX$, 32, 13)
            Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
            Call REPLA(Z$, SALTEX$, 56, 13)
            Call REPLA(Z$, ACRTEX$, 70, 14)
            Call REPLA(Z$, REGTEX$, 119, 8)
            Call REPLA(Z$, IRETEX$, 127, 2)
            List2.AddItem Z$
            '
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(Val(Mid$(Z$, 70, 14)))
    Next U&
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label18.Refresh
    Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - Val(Label21.Caption)), 4, 14, 2, 2)
    Label20.Refresh
    '
    List2.Refresh
    VACON% = 1
    '
End Sub
'

Sub SIGUEORPA4()
   '
   VACON% = 1
   List4.Clear
10 For U& = 1 To List1.ListCount
     Z$ = List1.List(U& - 1)
     If Val(Mid$(Z$, 70, 14)) >= 0.005 Then
       List4.AddItem Z$
     End If
   Next U&
   List1.Clear
   For U& = 1 To List4.ListCount
     List1.AddItem List4.List(U& - 1)
   Next U&
   List1.Refresh
   '
   SUMAVAL# = Val(Label21.Caption)
   For U& = 1 To ULTREG&("!CARMEPA")
      X$ = REGLEIDO$("!CARMEPA", U&)
      ICHE# = CVD(Mid$(X$, 31, 8))
      COVA% = CVI(Left$(X$, 2))
      TIVA$ = ECOARCH$("MEPAGO", MKI$(COVA%))
      NUCHE$ = TRIM$(Mid$(X$, 19, 12))
      '
      Z$ = Space$(128)
      Call REPLA(Z$, "Prop", 1, 4)
      Call REPLA(Z$, TIVA$, 7, 18)
      Call REPLA(Z$, NUCHE$, 25, 14)
      FEMI = CVI(Mid$(X$, 39, 2))
      Call REPLA(Z$, FECHATEX$(FEMI), 39, 8)
      FEVE = CVI(Mid$(X$, 41, 2))
      Call REPLA(Z$, FECHATEX$(FEVE), 49, 8)
      Call REPLA(Z$, CSTRING$(MKD$(ICHE#), 3, 12, 2, 2), 57, 12)
      Call REPLA(Z$, CSTRING$(MKD$(ICHE#), 3, 14, 2, 2), 70, 14)
      Call REPLA(Z$, CSTRING$(MKI$(COVA%), 1, 6, 0, 2), 115, 6)
      List1.AddItem Z$
      SUMAVAL# = SUMAVAL# + ICHE#
      Label21.Caption = CSTRING$(MKD$(SUMAVAL#), 4, 14, 2, 2)
    '
   Next U&
   '
   List1.Refresh
   Label21.Refresh
   Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - SUMAVAL#), 4, 14, 2, 2)
   Label20.Refresh
   '
   U& = 0
20 If U& < List2.ListCount Then
      List2.ListIndex = U&
      Z$ = List2.TEXT
      If Val(Mid$(Z$, 70, 14)) < 0.005 Then
           List2.RemoveItem U&
         Else
           U& = U& + 1
      End If
      GoTo 20
   End If
   '
   VACON% = 0
   '
   Command5.Enabled = True
   '
End Sub
'

'
Sub SIGUEORPA5()
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Distribución de Pagos"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    '
305 VTB% = VENTABU%("CARGAST" + ATRI$)
    If VTB% < 0 Then
       If COMALTER%("Realmente Desea Cancelar\la Distribución Ingresada ?\\No, Continuar\Si, Cancelar") <> 2 Then
           GoTo 305
       End If
       Call CIERRARCH("!CARGAST")
       Call BLANARCH("!CARGAST")
    End If
    '
    SUMACRE# = 0
    V& = 0: UF& = ULTREG&("!CARGAST")
    For U& = 1 To UF&
       X$ = REGLEIDO$("!CARGAST", U&)
       If ECOARCH$("ECUECON", Left$(X$, 12)) <> "" Then
       CUE% = CVI(Mid$(REGACT$("ECUECON"), 41, 2))
       Call REPLA(X$, MKI$(CUE%), 21, 2)
       SUMACRE# = SUMACRE# + CVD(Mid$(X$, 13, 8))
       V& = V& + 1
       Call GRAREG("!CARGAST", X$, V&)
       End If
    Next U&
    '
    Call SETULTREG("!CARGAST", V&)
    If V& < UF& Then
       Call COMUNI("Se han Detectado Importes Mal Aplicados.\  \Revise y Confirme la Distribución de Pagos.")
       GoTo 305
    End If
    '
    Label18.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label18.Refresh
    Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - Val(Label21.Caption)), 4, 14, 2, 2)
    Label20.Refresh
    '
    List2.ListIndex = 0
    Z$ = List2.TEXT
    IACRE$ = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Call REPLA(Z$, IACRE$, 70, 14)
    VACON% = 1
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = 0
    VACON% = 0
    '
End Sub
'
Sub GRAORPASTA(OKX%)
    '
    Call CIERRARCH("!OBSERVOF")
    Call BLANARCH("!OBSERVOF")
    VALIDA% = VENTABU%("OBSERVOF/TITUPAN=Orden de Pagos Numero" + Str$(CVS(Left$(X3$, 4))))
    If VALIDA% < 0 Then
       OKX% = (-1): Exit Sub
    End If
    '
    HP% = HOY(HOS$)
    OKX% = 0
    Call BLOQARCH("CUECORR")
    Call BLOQARCH("PROVED2")
    Call BLOQARCH("DETPAGAS")
    Call BLOQARCH("APLIPAG")
    Call BLOQARCH("CAJABANC")
    OKX% = 1                    ' bloqueó archivos
    '
    NC% = Val(Text1.TEXT)
    NC1% = (-1) * NC%
    NRO& = Val(Mid$(Text11.TEXT, 4))
    NROPRO& = NRO&
    Call NUORPA(NROPRO&)
    If NROPRO& <> NRO& Then
       Call CIERRARCH("CUECORR")
       Call CIERRARCH("PROVED2")
       Call CIERRARCH("DETPAGAS")
       Call CIERRARCH("APLIPAG")
       Call CIERRARCH("CAJABANC")
       Call MENSERR(24, "Al Grabar esta Orden de Pago\el Número Propuesto Estaba Ocupado.\  \Verifique el Nuevo Número Propuesto\y Vuelva a Grabar e Imprimir.")
       Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
       OKX% = 0
       Exit Sub
    End If
    '
    TIPO% = 4: VA% = 1: MO$ = "$"
    TOFAC#(0) = Val(Label14.Caption) + Val(Label21.Caption)
    TOFAC#(1) = 0
    TOFAC#(2) = TOFAC#(0)
    '
    FC% = FECHANUM(Text9.TEXT)
    FVII%(1) = FC%
    IMV#(1) = 0
    List2.ListIndex = 0
    X$ = List2.TEXT
    DOCU$ = Mid$(X$, 1, 20)
    If DOCU$ = "Anticipo no Aplicado" Then
        CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
        If Abs(CANCE#) >= 0.005 Then
           IMV#(1) = (-1) * CANCE#
        End If
    End If
    '
    ' grabar orden de pago en CUECORR
    Call CUECORR(FC%, (-1) * NC%, TIPO%, VA%, SUC%, NRO&, REFE$, MO$, TOFAC#(), FVII%(), IMV#(), GRU%(), IPAR#(), 1, FC%)
    URECOBRA& = ULTREG&("CUECORR")
    ' grabar orden de pago en CAJABANC
    COBRACT% = 0
    NRECIB$ = "0"    ' TALONARIO
    While Len(NRECIB$) < 2: NRECIB$ = "0" + NRECIB$: Wend
    NRECIB$ = TRIM$(Str$(NRO&)) + "-" + NRECIB$
    While Len(NRECIB$) < 9: NRECIB$ = "0" + NRECIB$: Wend
    NRECIB$ = "OP-" + NRECIB$
    '
    Z$ = REGBLAN$("CAJABANC")
    Call REPLA(Z$, NRECIB$, 1, 12)
    Call REPLA(Z$, RASOCLI$(NC1%), 13, 48)
    Call REPLA(Z$, MKI$(FC%), 61, 2)
    Call REPLA(Z$, MKI$(FC%), 63, 2)
    Call REPLA(Z$, MKI$(CUECOCLI%(NC1%)), 65, 2)
    Call REPLA(Z$, MKI$(NC1%), 67, 2)
    Call REPLA(Z$, "OP", 69, 2)
    Call REPLA(Z$, "Orden de Pago " + NRECIB$, 71, 40)
    Call REPLA(Z$, MKI$(FC%), 111, 2)
    Call REPLA(Z$, MKI$(COBRACT%), 115, 2)
    Call REPLA(Z$, MKS$(NRO&), 117, 4)
    Call REPLA(Z$, Chr$(0), 121, 1)
    Call REPLA(Z$, MKD$(TOFAC#(0)), 145, 8)
    Call REPLA(Z$, MKD$(TOFAC#(0)), 153, 8)
    Call REPLA(Z$, RASOCLI$(NC1%), 177, 32)
    Call REPLA(Z$, "CUECORR", 245, 8)
    Call REPLA(Z$, MKS$(URECOBRA&), 253, 4)
    Call REPLA(Z$, IDACTU$, 257, 6)  ' FECHA - HORA - USUARIO - TERMINAL
    Call REGAPP("CAJABANC", Z$)
    OKX% = 2         ' grabó Orden de Pago
    '
800  ' *********************************** aqui blanquear facturas canceladas
    '
    If TICUEPRO%(NC1%) = 0 Then
       '
       XX% = 0
       For U& = 1 To List2.ListCount
         List2.ListIndex = U& - 1
         Z$ = List2.TEXT
         DOCU$ = Left$(Z$, 20)
         If DOCU$ <> "Anticipo no Aplicado" Then
           CANCE# = CDbl(Val(Mid$(Z$, 70, 14)))
           If CANCE# > 0 Then
              NREG& = Val(TRIM$(Mid$(Z$, 119, 8)))
              If NREG& > 0 Then
                If NREG& <= ULTREG&("CUECORR") Then
                  ORD% = Val(Mid$(Z$, 127, 2))
                  If ORD% > 0 Then
                    If ORD% <= 6 Then
                       '
                       SS$ = REGLEIDO$("CUECORR", NREG&)
                       SALDO$ = Mid$(SS$, 139, 60)
                       SALDEU# = CVD(Mid$(SALDO$, 10 * ORD% - 7, 8))
                       SALDEU# = SALDEU# - CANCE#
                       Call REPLA$(SALDO$, MKD$(SALDEU#), 10 * ORD% - 7, 8)
                       '
                       If Abs(SALDEU#) < 0.005 Then
                         SALDO$ = Left$(SALDO$, 10 * (ORD% - 1)) + Mid$(SALDO$, 10 * ORD% + 1) + String$(10, 0)
                       End If
                       '
                       Call REPLA(SS$, SALDO$, 139, 60)
                       Call GRAREG("CUECORR", SS$, NREG&)
                    End If
                  End If
                End If
              End If
           End If
         End If
       Next U&
       '
    End If
    'Call CIERRARCH("CUECORR")
    OKX% = 3         ' depuró facturas pendientes
    '
820 If TICUEPRO%(NC1%) = 1 Then
        RC& = REGICLI&(NC1%)
        DX$ = REGLEIDO$("PROVED2", RC&)
        DX$ = REPLACE$(DX$, MKD$(0), 109, 8)
        Call GRAREG("PROVED2", DX$, RC&)
        '
        For K1& = 1 To ULTREG&("!CARGAST")
           Z$ = REGLEIDO$("!CARGAST", K1&)
           ZX$ = REGBLAN$("DETPAGAS")
           ZX$ = REPLACE$(ZX$, MKI$(NC%), 1, 2)
           ZX$ = REPLACE$(ZX$, MKS$(NRO&), 3, 4)
           ZX$ = REPLACE$(ZX$, MKI$(FC%), 7, 2)
           ZX$ = REPLACE$(ZX$, Mid$(Z$, 13, 10), 9, 10)
           AA = CVD(Mid$(Z$, 3))
           Call REGAPP("DETPAGAS", ZX$)
        Next K1&
        'Call CIERRARCH("DETPAGAS")
        GoTo 840
    End If
    '
    '
830 ' *************************************** aqui grabar archivo APLIPAG.DAT
    '
       For U& = 1 To List2.ListCount
         List2.ListIndex = U& - 1
         Z$ = List2.TEXT
         DOCU$ = Left$(Z$, 20)
         If DOCU$ <> "Anticipo no Aplicado" Then
           CANCE# = CDbl(Val(Mid$(Z$, 70, 14)))
           If CANCE# > 0 Then
              NREG& = Val(TRIM$(Mid$(Z$, 119, 8)))
              If NREG& > 0 Then
                If NREG& <= ULTREG&("CUECORR") Then
                  SS$ = REGLEIDO$("CUECORR", NREG&)
                  DOCORI$ = Mid$(SS$, 7, 18)
                  SALDO$ = Mid$(SS$, 139, 60)
                  SS77$ = Mid$(SS$, 51)
                  ITOTORI# = Abs(CVD(Mid$(SS$, 115, 8)) - CVD(Mid$(SS$, 123, 8)))
                  IVAORI# = Abs(CVD(Mid$(SS77$, 25, 8)) + CVD(Mid$(SS77$, 33, 8)) + CVD(Mid$(SS77$, 41, 8)))
                  '
                  Call APLIPAG(FC%, NC1%, 4, 1, SUC%, NRO&, CANCE#, REFAPLI$, DOCORI$, ITOTORI#, IVAORI#)
                  '
                End If
              End If
           End If
         End If
       Next U&
       OKX% = 4       ' grabó aplicaciones
     '
840  'Call CIERRARCH("APLIPAG")
     'Call CIERRARCH("DETPAGAS")
     '
     ' ************************************** aqui grabar archivo CHEQUEMI.DAT
     '
     REBLAN$ = REGBLAN$("CAJABANC")
     NORPAG$ = NRECIB$
     Z0$ = REBLAN$
     Call REPLA(Z0$, NORPAG$, 1, 12)
       RASOP$ = RASOCLI$(NC1%)
     Call REPLA(Z0$, TRIM$(RASOP$), 13, 48)
     Call REPLA(Z0$, MKI$(FC%), 61, 2)
     '
     For U& = 1 To List1.ListCount
       Z$ = List1.List(U& - 1)
       TIDOS$ = Left$(Z$, 4)
       TIPO% = Val(Mid$(Z$, 115, 6))
       BANCO$ = Mid$(Z$, 7, 18)
       NUCHE$ = Mid$(Z$, 25, 14)
       IMCHE# = Val(Mid$(Z$, 70, 14))
       FECHE% = FECHANUM(Mid$(Z$, 39, 8)): If FECHE% < 1 Then FECHE% = HO%
       FEVEN% = FECHANUM(Mid$(Z$, 49, 8)): If FEVEN% < FECHE% Then FEVEN% = FECHE%
       '
       HOII% = HOY(HOS$)
       If IMCHE# > 0 Then
          REGBA& = Val(TRIM$(Mid$(Z$, 121, 8)))
          If REGBA& > 0 Then ' es un cheque de cartera
               '
               TTS$ = REGLEIDO$("CAJABANC", REGBA&)
               Call REPLA(TTS$, MKI$(FC%), 275, 2)
               Call REPLA(TTS$, NORPAG$, 277, 16)
               Call REPLA(TTS$, RASOCLI$(NC1%), 293, 30)
               Call REPLA(TTS$, "Orden de Pago Nro." + Str$(NRO&), 323, 40)
               Call REPLA(TTS$, HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 363, 40)
               Call REPLA(TTS$, "Usuario: " + TRIM$(USERNAME$), 403, 40)
               Call REPLA(TTS$, Chr$(2), 121, 1)  ' Marca Salida de Cartera
               Call REPLA(TTS$, IDACTU$, 263, 6)  ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
               Call GRAREG("CAJABANC", TTS$, REGBA&)
               '
               Z$ = TTS$
               Call REPLA(Z$, Mid$(REBLA$, 209), 209, 304)
               '
               Call REPLA(Z$, NORPAG$, 1, 12)
                 FEX = FC%
               Call REPLA(Z$, TRIM$(RASOCLI$(NC1%)) + " - " + FECHATEX$(FEX), 13, 48)
               Call REPLA(Z$, MKI$(FC%), 61, 2)
               Call REPLA(Z$, "C3", 69, 2)
                 IVAX$ = Mid$(Z$, 71, 40)
                 IDEVA$ = TRIM$(Left$(IVAX$, 18)) + " " + TRIM$(Mid$(IVAX$, 19, 14)) + " " + TRIM$(Mid$(IVAX$, 33))
               Call REPLA(Z$, MKI$(0), 115, 2)
               Call REPLA(Z$, MKS$(NRO&), 117, 4)
                 STATUI% = 0
               Call REPLA(Z$, Chr$(STATUI%), 121, 1) ' STATUS%=1 ES CHEQUE EN CARTERA
               Call REPLA(Z$, MKS$(0), 125, 4)
               Call REPLA(Z$, String$(16, 0), 129, 16)
               Call REPLA(Z$, MKD$(0), 153, 8)
               Call REPLA(Z$, MKD$(IMCHE#), 161, 8)
               Call REPLA(Z$, MKD$(0), 169, 8)
               Call REPLA(Z$, IDEVA$, 177, 32)
               If TRIM$(Mid$(Z$, 89, 12)) = "" Then
                 Call REPLA(Z$, RASOP$, 177, 32)
               End If
               Call REPLA(Z$, "CAJABANC", 245, 8)
               Call REPLA(Z$, MKS$(REGBA&), 253, 4)
               Call REPLA(Z$, IDACTU$, 257, 6)  ' FECHA - HORA - USUARIO - TERMINAL
               Call REGAPP("CAJABANC", Z$)
               '
            Else ' es un cheque propio
               Z$ = Z0$
               FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
               IDEVA$ = AJUSTI$(BANCO$, 18) + AJUSTI$(NUCHE$, 14) + FECHATEX$(FEX)
               '
               Call REPLA(Z$, MKI$(FEVEN%), 63, 2)
                 CUECO% = CUEMEPAG%(TIPO%)
               Call REPLA(Z$, MKI$(CUECO%), 65, 2)
               Call REPLA(Z$, MKI$(NC1%), 67, 2)
               Call REPLA(Z$, "CP", 69, 2)
               Call REPLA(Z$, IDEVA$, 71, 40)
               Call REPLA(Z$, MKI$(FECHE%), 111, 2)
               Call REPLA(Z$, MKI$(COBRACT%), 115, 2)
               Call REPLA(Z$, MKS$(NRO&), 117, 4)
                 STATUI% = 0
               Call REPLA(Z$, Chr$(STATUI%), 121, 1) ' STATUS%=1 ES CHEQUE EN CARTERA
               Call REPLA(Z$, MKD$(IMCHE#), 145, 8)
               Call REPLA(Z$, MKD$(IMCHE#), 161, 8)
               Call REPLA(Z$, RASOP$, 177, 32)
               If TRIM$(NUCHE$) <> "" Then
                 Call REPLA(Z$, TRIM$(BANCO$) + " " + TRIM$(NUCHE$) + " " + TRIM$(FEVECHE$), 177, 32)
               End If
               Call REPLA(Z$, IDACTU$, 257, 6)  ' FECHA - HORA - USUARIO - TERMINAL
               Call REGAPP("CAJABANC", Z$)
          End If
       End If
     Next U&
     '
     OKX% = 5          ' grabo cheques de terceros
     OKX% = 6          ' grabo cheques emitidos
     '
890  Call CIERRARCH("PROVED2")
     Call CIERRARCH("DETPAGAS")
     Call CIERRARCH("APLIPAG")
     Call CIERRARCH("CAJABANC")
     Call CIERRARCH("CUECORR")
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     OKX% = 7          ' LIBERO ARCHIVOS
     '
End Sub
'
Sub PRIORPASTA()
    '
    HOI% = HOY(HO$)
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORPASTA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Orden de Pago.")
        Exit Sub
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo de Cobranza:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    CODFOR$ = FORIPRE$
    '
    CONTA% = List1.ListCount
    If CONTA% < 1 Then
        Call COMUNI("No se ha Ingresado Detalle de Valores")
        Exit Sub
    End If
    '
    If Val(Label14.Caption) + Val(Label21.Caption) < 0.005 Then
        Call COMUNI("No se ha Ingresado Detalle de Valores")
        Exit Sub
    End If
    '
    AVAHOJA$ = TRIM$(CONTROL$("ORPASTA", "AVAHOJA"))
    Screen.MousePointer = 11
    Call BLANARCH("!PRINRECI")
    '
    ILISTA& = List1.ListIndex
    TOLISTA& = List1.TopIndex
    '
    For U& = 1 To CONTA%
        List1.ListIndex = U& - 1
        X$ = List1.TEXT
        If TRIM$(X$) = "" Then GoTo 19
        TIPOS$ = Left$(X$, 4)
        BANCO$ = Mid$(X$, 7, 18)
        NUVA$ = Mid$(X$, 25, 14)
        IMVA# = Val(Mid$(X$, 70, 14))
        If IMVA# < 0.005 Then GoTo 19
        '
        FEMII% = FECHANUM(Mid$(X$, 39, 8)): If FEMII% < 1 Then FEMII% = HOI%
        FEVEI% = FECHANUM(Mid$(X$, 49, 8)): If FEVEI% < 1 Then FEMII% = HOI%
        PROTER$ = ""
        CLEA% = 48
        '
10      Z$ = REGBLAN$("!PRINRECI")
        Call REPLA(Z$, TIPOS$, 1, 4)
        Call REPLA(Z$, NUVA$, 5, 12)
        Call REPLA(Z$, BANCO$, 17, 16)
        Call REPLA(Z$, MKD$(IMVA#), 33, 8)
        FEX = FEVEI%
        If FEX > 0 Then
            Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
        End If
        Call REGAPP("!PRINRECI", Z$)
        '
19  Next U&
    '
    V0& = ULTREG&("!PRINRECI")
    '
    ILISTA& = List2.ListIndex
    TOLISTA& = List2.TopIndex
    '
    V& = 0
    If TICUEPRO%((-1) * NC%) = 1 Then
       For U& = 1 To ULTREG&("!CARGAST")
          X$ = REGLEIDO$("!CARGAST", U&)
          DOCU$ = Left$(X$, 12)
          SALPE# = 0
          CANCE# = CVD(Mid$(X$, 13, 8))
          If Abs(CANCE#) >= 0.005 Then
             V& = V& + 1
             If V& <= V0& Then
                  Z$ = REGLEIDO$("!PRINRECI", V&)
                Else
                  Z$ = REGBLAN$("!PRINRECI")
             End If
             Call REPLA(Z$, DOCU$, 53, 18)
             Call REPLA(Z$, MKD$(CANCE#), 71, 8)
             Call REPLA(Z$, MKD$(SALPE#), 79, 8)
             Call GRAREG("!PRINRECI", Z$, V&)
          End If
       Next U&
       Call CIERRARCH("!CARGAST")
       GoTo 25
    End If
    '
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        X$ = List2.TEXT
        DOCU$ = Mid$(X$, 1, 20)
        SALPE# = CDbl(Val(Mid$(X$, 56, 13)))
        CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
        If Abs(CANCE#) >= 0.005 Then
            V& = V& + 1
            If V& <= V0& Then
                  Z$ = REGLEIDO$("!PRINRECI", V&)
               Else
                  Z$ = REGBLAN$("!PRINRECI")
            End If
            Call REPLA(Z$, DOCU$, 53, 18)
            Call REPLA(Z$, MKD$(CANCE#), 71, 8)
            Call REPLA(Z$, MKD$(SALPE#), 79, 8)
            Call GRAREG("!PRINRECI", Z$, V&)
        End If
    Next U&
    '
25  V0& = ULTREG&("!PRINRECI")
    '
    TOTAVAL# = CDbl(Val(Label14.Caption) + Val(Label21.Caption))
    TOTACAN# = CDbl(Val(Label18.Caption)) + CDbl(Val(Label20.Caption))
33  ITEXTO$ = NUMTEX$(TOTAVAL#) + " ctvs"
    LOTEX% = Val(ATRIFORM$(CODFOR$, "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = "": IMPLET2$ = "": IMPLET3$ = ""
    If TRIM$(NRX1$) <> "" Then
       IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    End If
    If TRIM$(NRX2$) <> "" Then
       IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
    End If
    If TRIM$(NRX3$) <> "" Then
       IMPLET3$ = TRIM$(NRX3$) + " " + "***************************************************************************************************************"
    End If
    '
    Call CIERRARCH("!PRINRECI")
    '
30  If TRIM$(FORIPRE$) <> "" Then
        '
        CANCELPRINT% = 0
        If CONTROL$("ORPASTA", "SUPRINT") = "SI" Then CANCELPRINT% = 1
        '
        CODFOR$ = TRIM$(UCase$(FORIPRE$))
        MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
        MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
        '
        INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
        If INTERLI < 1 Then
            Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
            INTERLI = 1
        End If
        '
        DESMM = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
        HASMM = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
        '
        CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
        If CACOPI% < 1 Or CACOPI% > 6 Then
            Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
            CACOPI% = 1
        End If
        '
        ORIENTA% = 1
        PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
        If PORI$ = "PAISAJE" Or PORI$ = "2" Then
            ORIENTA% = 2
        End If
        '
        PORIAN = Printer.Orientation  ' retiene orientacion
        '
        NRORPA$ = TRIM$(Mid$(Text11.TEXT, 4))
        '
        Screen.MousePointer = 11
        If AVAHOJA$ <> "NEWPAGE" Then
            'Printer.Orientation = 1
            Call SETSPOOL("ORIENTATION", "1")
            If ORIENTA% = 2 Then
                'Printer.Orientation = 2
                Call SETSPOOL("ORIENTATION", "2")
            End If
        End If
        '
        NC1% = (-1) * NC%
        For COP% = 1 To CACOPI%
            '
            If AVAHOJA$ = "NEWPAGE" Then
               If PRINTERCONTROL% = 1 Then
                  'Printer.NewPage
                  Call SETSPOOL("NEWPAGE", "")
               End If
            End If
            PRINTERCONTROL% = 1
            '
            SUMAVA# = 0: SUMACAN# = 0
            Call CARFORWIN
            '
            Call PRINTFORWIN(TRIM$(Str$(NC%)), "CUEN-CLI", 0, 0)
            Call PRINTFORWIN(TRIM$(RASOCLI$(NC1%)), "RASO-CLI", 0, 0)
            Call PRINTFORWIN(DOMICLI$(NC1%), "DOMI-CLI", 0, 0)
            Call PRINTFORWIN(CPOSCLI$(NC1%), "CPOS-CLI", 0, 0)
            Call PRINTFORWIN(LOCACLI$(NC1%), "LOCA-CLI", 0, 0)
            Call PRINTFORWIN(CUITCLI$(NC1%), "CUIT-CLI", 0, 0)
            Call PRINTFORWIN(POSIVAL$(NC1%), "PIVA-CLI", 0, 0)
            FF% = FECHANUM(TRIM$(Text9.TEXT))
            Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
            Call PRINTFORWIN(NRORPA$, "NUME-COM", 0, 0)
            '
            YAC = 0
            For U& = 1 To ULTREG&("!OBSERVOF")
               TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", U&))
               Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
               YAC = YAC + INTERLI
            Next U&
            Call CIERRARCH("!OBSERVOF")
            '
            YAC = DESMM
            For YX& = 1 To ULTREG&("!PRINRECI")
               '
               YAC = YAC + INTERLI
               '
               If YAC > HASMM - 2 * INTERLI Then
                   YAC = YAC + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   '
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-VAL", 0, 0)
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-CAN", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET1", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET2", 0, 0)
                   '
                   If AVAHOJA$ = "NEWPAGE" Then
                        'Printer.NewPage
                        Call SETSPOOL("NEWPAGE", "")
                      Else
                       'Printer.EndDoc
                       Call SETSPOOL("ENDDOC", "")
                   End If
                   '
                   Call CARFORWIN
                   Call PRINTFORWIN(TRIM$(Str$(NC1%)), "CUEN-CLI", 0, 0)
                   Call PRINTFORWIN(TRIM$(RASOCLI$(NC1%)), "RASO-CLI", 0, 0)
                   Call PRINTFORWIN(DOMICLI$(NC1%), "DOMI-CLI", 0, 0)
                   Call PRINTFORWIN(CPOSCLI$(NC1%), "CPOS-CLI", 0, 0)
                   Call PRINTFORWIN(LOCACLI$(NC1%), "LOCA-CLI", 0, 0)
                   Call PRINTFORWIN(CUITCLI$(NC1%), "CUIT-CLI", 0, 0)
                   Call PRINTFORWIN(POSIVAL$(NC1%), "PIVA-CLI", 0, 0)
                   FEX = FF%
                   Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
                   Call PRINTFORWIN(NRORPA$, "NUME-COM", 0, 0)
                   '
                   YAC = DESMM + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   YAC = YAC + 2 * INTERLI
                   '
                End If
                '
                YY$ = REGLEIDO$("!PRINRECI", YX&)
                Call PRINTFORWIN(Left$(YY$, 4), "TIPO-VAL", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 5, 12), "NUME-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 17, 16), "BANC-CHE", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 33, 8), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                SUMAVA# = SUMAVA# + CVD(Mid$(YY$, 33, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 41, 8), 6, 8, 0, 1), "FECH-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 53, 18), "DOCU-APL", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 71, 8), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                SUMACAN# = SUMACAN# + CVD(Mid$(YY$, 71, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 79, 8), 4, 12, 2, 2), "IMPO-PEN", 0, YAC)
                '
            Next YX&
            '
            Call PRINTFORWIN(CSTRING$(MKD$(TOTAVAL#), 4, 12, 2, 2), "TOTA-VAL", 0, 0)
            Call PRINTFORWIN(CSTRING$(MKD$(TOTACAN#), 4, 12, 2, 2), "TOTA-CAN", 0, 0)
            Call PRINTFORWIN(IMPLET1$, "IMP-LET1", 0, 0)
            Call PRINTFORWIN(IMPLET2$, "IMP-LET2", 0, 0)
                   '
            If AVAHOJA$ <> "NEWPAGE" Then
                 'Printer.EndDoc
                 Call SETSPOOL("ENDDOC", "")
            End If
            '
        Next COP%
        '
        Call CIERRARCH("*.*")
        DESCRIDOC$ = "Orden de Pago Nro. " + TRIM$(NRORPA$) + " - " + TRIM$(RASOCLI$(NC1%))
        FEX = FF%
        DESCRIDOC$ = DESCRIDOC$ + " - " + FECHATEX$(FEX)
        '
        Call SAVESPOOL(DESCRIDOC$, OKX%)
        Screen.MousePointer = 1
        If COMALTER%("Orden de Pago Numero\'" + NRORPA$ + "'\Emitida y Registrada\\Repetir Impresion\Otra Orden de Pago") = 1 Then GoTo 30
        '
        If AVAHOJA$ <> "NEWPAGE" Then
            Printer.Orientation = PORIAN
        End If
        '
    End If
    Call CIERRARCH("*.*")
    '
End Sub




