VERSION 5.00
Begin VB.Form FORCREDEV 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Notas de Crédito por Devoluciones"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   6
      Top             =   0
      Width           =   12000
      Begin VB.TextBox Eco21 
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
         Left            =   7620
         TabIndex        =   47
         Text            =   " "
         Top             =   7140
         Width           =   2640
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
         Left            =   6720
         TabIndex        =   46
         Text            =   " "
         Top             =   7140
         Width           =   435
      End
      Begin VB.CommandButton Command21 
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
         Index           =   2
         Left            =   7140
         TabIndex        =   45
         Top             =   7140
         Width           =   180
      End
      Begin VB.TextBox Eco21 
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
         TabIndex        =   44
         Text            =   " "
         Top             =   7170
         Width           =   3165
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
         Left            =   1260
         TabIndex        =   43
         Text            =   " "
         Top             =   7170
         Width           =   435
      End
      Begin VB.CommandButton Command21 
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
         Index           =   1
         Left            =   1680
         TabIndex        =   42
         Top             =   7170
         Width           =   180
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1905
         Left            =   10680
         TabIndex        =   39
         Top             =   1575
         Width           =   1170
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
            Left            =   210
            Picture         =   "FORCREDEV.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Notas de Clientes"
            Top             =   105
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
            Left            =   210
            Picture         =   "FORCREDEV.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   945
            Width           =   855
         End
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
         Left            =   5250
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   2640
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   4485
         Left            =   210
         ScaleHeight     =   4425
         ScaleWidth      =   11505
         TabIndex        =   17
         Top             =   1680
         Width           =   11565
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   1.11565e5
            Begin VB.Line Line10 
               X1              =   6700
               X2              =   6700
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line13 
               X1              =   10430
               X2              =   10430
               Y1              =   0
               Y2              =   315
            End
            Begin VB.Line Line1 
               X1              =   8925
               X2              =   8925
               Y1              =   0
               Y2              =   1155
            End
            Begin VB.Line Line2 
               X1              =   7840
               X2              =   7840
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   $"FORCREDEV.frx":0614
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
               TabIndex        =   24
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line8 
               X1              =   630
               X2              =   630
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line9 
               X1              =   2520
               X2              =   2520
               Y1              =   0
               Y2              =   1050
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5630
            Left            =   0
            TabIndex        =   18
            Top             =   315
            Width           =   1.11565e5
            Begin VB.CommandButton Command6 
               Caption         =   "Add."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10710
               Picture         =   "FORCREDEV.frx":069C
               TabIndex        =   36
               ToolTipText     =   "Agrega un Item al Final"
               Top             =   1680
               Width           =   580
            End
            Begin VB.CommandButton Command12 
               Caption         =   "Ins."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10710
               Picture         =   "FORCREDEV.frx":09A6
               TabIndex        =   35
               ToolTipText     =   "Intercala un Item Antes del Activo"
               Top             =   2235
               Width           =   580
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Del"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10710
               Picture         =   "FORCREDEV.frx":0CB0
               TabIndex        =   34
               ToolTipText     =   "Elimina el Item Activo "
               Top             =   2790
               Width           =   580
            End
            Begin VB.CommandButton Command8 
               Caption         =   "Edit"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10710
               Picture         =   "FORCREDEV.frx":0FBA
               TabIndex        =   33
               ToolTipText     =   "Eldición del Item Activo"
               Top             =   3345
               Width           =   580
            End
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
               Height          =   3660
               Left            =   -30
               TabIndex        =   19
               Top             =   0
               Width           =   10450
            End
            Begin VB.Line Line14 
               X1              =   0
               X2              =   10395
               Y1              =   3680
               Y2              =   3680
            End
            Begin VB.Line Line24 
               X1              =   10425
               X2              =   10425
               Y1              =   0
               Y2              =   4935
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
               TabIndex        =   26
               Top             =   4650
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
               TabIndex        =   25
               Top             =   0
               Width           =   1695
            End
            Begin VB.Line Line4 
               X1              =   8400
               X2              =   8400
               Y1              =   2520
               Y2              =   4935
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
               TabIndex        =   22
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
               TabIndex        =   21
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Subtotal"
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
               TabIndex        =   20
               Top             =   3770
               Width           =   1695
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   10420
               Y1              =   4095
               Y2              =   4095
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
         Left            =   6405
         MaxLength       =   1
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   315
         Width           =   500
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
         Picture         =   "FORCREDEV.frx":12C4
         TabIndex        =   15
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   315
         Width           =   855
      End
      Begin VB.CommandButton BOTNEXT 
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
         Height          =   1020
         Left            =   10875
         Picture         =   "FORCREDEV.frx":15CE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Aprueba - Registra - Imprime Factura"
         Top             =   6405
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
         TabStop         =   0   'False
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
         TabIndex        =   3
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
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   2745
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
         TabIndex        =   10
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
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3585
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
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   315
         Width           =   3585
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
      Begin VB.Label Label23 
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
         Left            =   6300
         TabIndex        =   50
         Top             =   6405
         Width           =   1380
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   10665
         X2              =   10665
         Y1              =   6090
         Y2              =   7665
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vended:"
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
         Left            =   5670
         TabIndex        =   49
         Top             =   7140
         Width           =   960
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L.Precio:"
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
         TabIndex        =   48
         Top             =   7170
         Width           =   960
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   0
         X2              =   10670
         Y1              =   6930
         Y2              =   6930
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IVA-RNI:"
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
         Left            =   3990
         TabIndex        =   37
         Top             =   6510
         Width           =   2220
      End
      Begin VB.Label Label20 
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
         Left            =   8925
         TabIndex        =   32
         Top             =   6405
         Width           =   1590
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tot.Bruto:"
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
         Left            =   6615
         TabIndex        =   31
         Top             =   6510
         Width           =   2220
      End
      Begin VB.Label Label18 
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
         TabIndex        =   30
         Top             =   6405
         Width           =   1380
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A.:"
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
         Left            =   2520
         TabIndex        =   29
         Top             =   6510
         Width           =   1170
      End
      Begin VB.Label Label14 
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
         Left            =   1260
         TabIndex        =   28
         Top             =   6405
         Width           =   1590
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tot,Neto:"
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
         Left            =   -1050
         TabIndex        =   27
         Top             =   6510
         Width           =   2220
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
         Height          =   435
         Left            =   7560
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   315
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
         TabIndex        =   7
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
      TabIndex        =   12
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "FORCREDEV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$


Private Sub BOTNEXT_Click()
   If List2.ListCount < 1 Then
       Call MENSERR(24, "Falta Ingresar Detalle de Nota de Crédito")
       Exit Sub
   End If
   '
   OKEY% = 0
   Call GRACREDEV(OKEY%)
   If OKEY% > 0 Then
      Call PRICREDEV
      Call CLEARFORCREDEV
      Text1.SetFocus
   End If
   '
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        Text1.TEXT = TRIM$(Str$(NCXI%))
        NC% = NCXI%
        OKEY% = 0
        Call CREDEVOL1(OKEY%)
        If OKEY% <> 1 Then
           Call CLEARFORCREDEV
           Exit Sub
        End If
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
    Call CONECRUN("AMACLI/NEW", "Alta de Nueva Cuenta de Cliente")
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Call MENSERR(24, "Opción Inserción no Disponible.\Debe Agregarse al Final de la Lista.")
End Sub

Private Sub Command2_Click()
    If List2.ListCount > 0 Then
          List2.Clear
          Text1.TEXT = ""
          Text2.TEXT = ""
          Text3.TEXT = ""
          Text4.TEXT = ""
          Text5.TEXT = ""
          HOII% = HOY(HOS$)
          Text9.TEXT = HOS$
          FEFAC% = HOII%
          '
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          '
          Text1.SetFocus
        Else
          OPFACMAI.Option1.Value = True
          VACON% = (-1)
          Unload FORCREDEV
          FORFACMAI.Show
          OPFACMAI.Show 1
    End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    Call CONECRUN("*NOTACLI/" + NCLI$, "Notas de Clientes")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RESCUDEU.Text1.TEXT = Text1.TEXT
    Hide
    PROLLAMA$ = "FORCREDEV"
    RESCUDEU.Show
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    '
    If List2.ListIndex < 0 Then Exit Sub
    If List2.ListIndex > List2.ListCount - 1 Then Exit Sub
    '
    ILISTA& = List2.ListIndex
    List2.RemoveItem ILISTA&
    List2.ListIndex = (-1)
    If List2.ListCount > ILISTA& Then
       List2.ListIndex = ILISTA&
    End If
    '
    TOPEDIDO# = 0
    For U& = 1 To FORCREDEV.List2.ListCount
       Z$ = FORCREDEV.List2.List(U& - 1)
       TOPEDIDO# = TOPEDIDO# + Val(Mid$(Z$, 74, 10))
    Next U&
    '
    ALITOTA = 0
    If PIVAFAC% <> 5 Then ALITOTA = ALIVA
    FORCREDEV.Label14.Caption = CSTRING$(MKD$(TOPEDIDO#), 3, 11, 2, 2)
    If PIVAFAC% <> 5 Then FORCREDEV.Label18.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVA / 100), 3, 11, 2, 2)
    If RENOIN% = 1 Then
       FORCREDEV.Label23.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVARNI / 100), 3, 11, 2, 2)
       ALITOTA = ALIVA + ALIVARNI
    End If
    FORCREDEV.Label20.Caption = CSTRING$(MKD$(TOPEDIDO# * (1 + ALITOTA / 100)), 3, 11, 2, 2)
    '
End Sub

Private Sub Command6_Click()
    '
    OKEY% = 0
    Call CREDEVOL1(OKEY%)
    If OKEY% <> 1 Then
        Call CLEARFORCREDEV
        Text1.SetFocus
        Exit Sub
    End If
    If List2.ListCount < 1 Then
        Call BLANARCH("!CARITPED")
    End If
    '
5   VDEF$ = REGBLAN$("CARITPED")
10  OBX$ = OBJPLA$("CARITPED1" + "/ARCHECO(1)=" + ARCHIPRE$, VDEF$)
    If OBX$ = "" Then Exit Sub
    VDEF$ = OBX$
    CIIT% = CVI(Mid$(OBX$, 3, 2))
    CAIT = CVS(Mid$(OBX$, 5, 4))
    If CODEXT$(CIIT%) = "" Then
       Call MENSERR(24, "Codigo no Válido")
       GoTo 10
    End If
    '
    If CAIT < 0.1 Then
       Call MENSERR(24, "Falta Cantidad")
       GoTo 10
    End If
    '
    PREBA = PREBASE(CIIT%)
    PRELI = PRELISTA#(ARCHIPRE$, CIIT%)
    If PRELI < 0.005 Then
       Call MENSERR(24, "Artículo sin Precio")
    End If
    '
    PRECO = PRELI
    If PRECO < PREBA Then PRECO = PREBA
20  IMPO# = CAIT * PRECO
    '
    Call REPLA(OBX$, MKS$(PRELI), 9, 4)
    Call REPLA(OBX$, MKS$(PREBA), 13, 4)
    Call REPLA(OBX$, MKS$(PRECO), 17, 4)
    Call REPLA(OBX$, MKD$(IMPO#), 21, 8)
    '
    VDEF$ = OBX$ + "/ARCHECO(3)"
    CAANT = CAIT: PREANT = PRECO
25  Y$ = OBJPLA$("CARITPED2", VDEF$)
    If Y$ = "" Then Exit Sub
    If CVI(Mid$(Y$, 3, 2)) <> CIIT% Then
       Call MENSERR(24, "Cambio de Artículo no Válido\en esta Pantalla.")
       GoTo 10
    End If
    '
    VDEF$ = Y$
    CIIT% = CVI(Mid$(Y$, 3, 2))
    CAIT = CVS(Mid$(Y$, 5, 4))
    PRELI = CVS(Mid$(Y$, 9, 4))
    PREBA = PREBASE(CIIT%)
    PRECO = CVS(Mid$(Y$, 17, 4))
    If PRECO < PREBA Then
       Call MENSERR(24, "Precio Cotizado Inferior a Precio de Base")
       GoTo 25
    End If
    '
    IMPO# = CAIT * PRECO
    '
    Z$ = Space$(96)
    Call REPLA(Z$, CSTRING$(MKS$(CAIT), 3, 4, 0, 2), 1, 4)
    Call REPLA(Z$, CODEXT$(CIIT%), 7, 15)
    Call REPLA(Z$, DESCRIT0$(CIIT%), 23, 33)
    Call REPLA(Z$, CSTRING$(MKS$(PRELI), 3, 9, 2, 2), 56, 9)
    Call REPLA(Z$, CSTRING$(MKS$(PRECO), 3, 9, 2, 2), 65, 9)
    Call REPLA(Z$, CSTRING$(MKD$(IMPO#), 3, 10, 2, 2), 74, 10)
    FORCREDEV.List2.AddItem Z$
    List2.ListIndex = List2.ListCount - 1
    '
    TOPEDIDO# = 0
    For U& = 1 To FORCREDEV.List2.ListCount
       Z$ = FORCREDEV.List2.List(U& - 1)
       TOPEDIDO# = TOPEDIDO# + Val(Mid$(Z$, 74, 10))
    Next U&
    '
    ALITOTA = 0
    If PIVAFAC% <> 5 Then ALITOTA = ALIVA
    FORCREDEV.Label14.Caption = CSTRING$(MKD$(TOPEDIDO#), 3, 11, 2, 2)
    If PIVAFAC% <> 5 Then FORCREDEV.Label18.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVA / 100), 3, 11, 2, 2)
    If RENOIN% = 1 Then
       FORCREDEV.Label23.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVARNI / 100), 3, 11, 2, 2)
       ALITOTA = ALIVA + ALIVARNI
    End If
    FORCREDEV.Label20.Caption = CSTRING$(MKD$(TOPEDIDO# * (1 + ALITOTA / 100)), 3, 11, 2, 2)
    '
    GoTo 5
    '
End Sub
'

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
        FEFAC% = FF%
    End If
End Sub

Private Sub Eco21_Click(Index As Integer)
   Text1.SetFocus
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    FEFAC% = FF%
    '
    Refresh
    VACON% = 0
    '
End Sub

Private Sub Label2_DblClick()
    OPREMIFAC.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 1
   If VACON% = (-1) Then
      Cancel = 0
   End If
End Sub

Private Sub List1_Click()
   Call CARITFAC1
End Sub

Private Sub List1_DblClick()
    '
    For J& = 1 To List1.ListCount
       Z$ = List1.List(J& - 1)
       CACON = Val(Mid$(Z$, 4, 7))
       RECON$ = Mid$(Z$, 13, 42)
       FFI% = FECHANUM(Mid$(Z$, 57, 8))
       PREUCO = Val(Mid$(Z$, 65, 9))
       ITOTCO# = Val(Mid$(Z$, 74, 10))
       ITID% = Val(Left$(Z$, 3))
       '
       X$ = REGBLAN$("CARCONPE")
       Call REPLA(X$, MKS$(CACON), 1, 4)
       Call REPLA(X$, RECON$, 5, 58)
       Call REPLA(X$, MKI$(FFI%), 63, 2)
       Call REPLA(X$, MKS$(PREUCO), 65, 4)
       Call REPLA(X$, MKD$(ITOTCO#), 69, 8)
       Call REPLA(X$, MKI$(ITID%), 127, 2)
       Call GRAREG("!CARCONPE", X$, J&)
    Next J&
    Call SETULTREG("!CARCONPE", List1.ListCount)
    '
    Call CARCONFAC
    '
End Sub

Private Sub List2_Click()
    '
End Sub
'

Private Sub List2_DblClick()
    '
End Sub

Private Sub Text1_Change()
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label18.Caption = ""
    Label20.Caption = ""
    Label23.Caption = ""
    BOTNEXT.Enabled = True
End Sub

Private Sub Text1_DBLCLICK()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        Text1.TEXT = TRIM$(Str$(NCXI%))
        NC% = NCXI%
        OKEY% = 0
        Call CREDEVOL1(OKEY%)
        If OKEY% <> 1 Then
           Call CLEARFORCREDEV
           Exit Sub
        End If
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
       OKEY% = 0
       Call CREDEVOL1(OKEY%)
       If OKEY% <> 1 Then
           Call CLEARFORCREDEV
           Exit Sub
       End If
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

Private Sub Text11_CLICK()
   Text1.SetFocus
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub

Private Sub TEXT21_Click(Index As Integer)
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
    OPFACTUR.Show 1
End Sub

Private Sub Text7_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FEFAC% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Text9_LostFocus()
    VDEVV$ = TRIM$(Text9.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FEFAC% = FECHANUM(VDEVV$)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          FEFAC% = FEX
          Text9.TEXT = HO$
          Text9.SetFocus
    End If
End Sub
