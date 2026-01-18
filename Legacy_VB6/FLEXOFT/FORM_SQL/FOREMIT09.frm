VERSION 5.00
Begin VB.Form FOREMIT09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Despacho de Materiales"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Rec. Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   11025
      Picture         =   "FOREMIT09.frx":0000
      TabIndex        =   58
      ToolTipText     =   "Recuperacion de Stocks"
      Top             =   1785
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   8200
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00400040&
         ForeColor       =   &H00FFFFFF&
         Height          =   8100
         Left            =   10880
         ScaleHeight     =   8040
         ScaleWidth      =   1185
         TabIndex        =   48
         Top             =   0
         Width           =   1245
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
            Height          =   700
            Left            =   105
            Picture         =   "FOREMIT09.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Continúa Carga de Remito."
            Top             =   990
            Width           =   800
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
            Height          =   700
            Left            =   120
            Picture         =   "FOREMIT09.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Aprueba - Graba e Imprime Remito"
            Top             =   6960
            Width           =   800
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Kardex"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   120
            Picture         =   "FOREMIT09.frx":0ABE
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Ficha de Movimientos de Material"
            Top             =   2625
            Width           =   800
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
            Height          =   700
            Left            =   120
            Picture         =   "FOREMIT09.frx":0DC8
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Opciones Complementarias - Anulaciones y Setup"
            Top             =   3480
            Width           =   800
         End
         Begin VB.CommandButton Command3 
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
            Height          =   700
            Left            =   120
            Picture         =   "FOREMIT09.frx":0F12
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   4200
            Width           =   800
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Trasla."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   120
            Picture         =   "FOREMIT09.frx":1354
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Generar Diskette de Transferencia"
            Top             =   4920
            Visible         =   0   'False
            Width           =   800
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
            Left            =   120
            Picture         =   "FOREMIT09.frx":165E
            TabIndex        =   51
            ToolTipText     =   "Ababdona la Aplicación"
            Top             =   300
            Width           =   800
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   150
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   49
            Top             =   6000
            Width           =   650
            Begin VB.Frame Frame10 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   50
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   -250
            Picture         =   "FOREMIT09.frx":1968
            Top             =   6240
            Width           =   1515
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0FF&
         Height          =   700
         Left            =   6930
         TabIndex        =   39
         Top             =   5250
         Width           =   3795
         Begin VB.TextBox Text17 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2205
            TabIndex        =   12
            Top             =   250
            Width           =   1380
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Declarado:"
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
            Left            =   0
            TabIndex        =   40
            Top             =   285
            Width           =   2115
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   6930
         TabIndex        =   38
         Top             =   5985
         Width           =   3795
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
            Index           =   3
            Left            =   210
            TabIndex        =   16
            Top             =   1365
            Width           =   3375
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
            Index           =   2
            Left            =   210
            TabIndex        =   15
            Top             =   1050
            Width           =   3375
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
            Index           =   1
            Left            =   210
            TabIndex        =   14
            Top             =   735
            Width           =   3375
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
            Index           =   0
            Left            =   210
            TabIndex        =   13
            Top             =   420
            Width           =   3375
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Height          =   2640
         Left            =   210
         TabIndex        =   33
         Top             =   5250
         Width           =   6525
         Begin VB.Frame Frame9 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Conductor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3885
            TabIndex        =   44
            Top             =   1575
            Width           =   2430
            Begin VB.TextBox Text13 
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
               Left            =   630
               TabIndex        =   46
               Top             =   420
               Width           =   1590
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "DNI:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   0
               TabIndex        =   45
               Top             =   465
               Width           =   540
            End
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   240
            ScaleHeight     =   75
            ScaleWidth      =   2580
            TabIndex        =   42
            Top             =   315
            Width           =   2640
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   43
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.TextBox Text15 
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
            Left            =   5250
            TabIndex        =   8
            Top             =   250
            Width           =   855
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Domicilio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   210
            TabIndex        =   36
            Top             =   1575
            Width           =   3585
            Begin VB.TextBox Text14 
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
               Left            =   210
               TabIndex        =   11
               Top             =   420
               Width           =   3165
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Transportista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   210
            TabIndex        =   34
            Top             =   630
            Width           =   6105
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
               Height          =   285
               Left            =   210
               TabIndex        =   9
               Top             =   420
               Width           =   3270
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
               Height          =   285
               Left            =   4305
               TabIndex        =   10
               Top             =   420
               Width           =   1590
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "CUIT:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3675
               TabIndex        =   35
               Top             =   465
               Width           =   540
            End
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Numero de Bultos:"
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
            Left            =   3045
            TabIndex        =   37
            Top             =   285
            Width           =   2115
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
         Left            =   9345
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   1335
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
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   315
         Width           =   435
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
         Top             =   735
         Width           =   1650
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
         Left            =   8715
         TabIndex        =   6
         Text            =   " "
         Top             =   1155
         Width           =   330
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
         Left            =   9030
         TabIndex        =   5
         Top             =   1155
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
         TabIndex        =   2
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   3480
         Left            =   210
         ScaleHeight     =   3420
         ScaleWidth      =   10455
         TabIndex        =   27
         Top             =   1680
         Width           =   10515
         Begin VB.Frame Frame4 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   540
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   10515
            Begin VB.Line Line7 
               X1              =   6300
               X2              =   6300
               Y1              =   0
               Y2              =   1260
            End
            Begin VB.Line Line6 
               X1              =   9345
               X2              =   9345
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line5 
               X1              =   7665
               X2              =   7665
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line4 
               X1              =   1785
               X2              =   1785
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line1 
               X1              =   5670
               X2              =   5670
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "  Codigo                  Descripcion                                           U-M     Cantidad      Lote / Color        N/Ped."
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
               TabIndex        =   30
               Top             =   140
               Width           =   10410
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5580
            Left            =   0
            TabIndex        =   28
            Top             =   -1050
            Width           =   10515
            Begin VB.ListBox List1 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2940
               Left            =   -40
               TabIndex        =   41
               Top             =   1575
               Visible         =   0   'False
               Width           =   10515
            End
            Begin VB.Line Line2 
               Index           =   1
               X1              =   5670
               X2              =   5670
               Y1              =   525
               Y2              =   5565
            End
            Begin VB.Line Line2 
               Index           =   4
               X1              =   1785
               X2              =   1785
               Y1              =   105
               Y2              =   5145
            End
            Begin VB.Line Line2 
               Index           =   3
               X1              =   9345
               X2              =   9345
               Y1              =   630
               Y2              =   5670
            End
            Begin VB.Line Line2 
               Index           =   2
               X1              =   7665
               X2              =   7665
               Y1              =   525
               Y2              =   5565
            End
            Begin VB.Line Line2 
               Index           =   0
               X1              =   6300
               X2              =   6300
               Y1              =   1260
               Y2              =   6300
            End
         End
      End
      Begin VB.Label NORVENTA 
         Caption         =   "Label8"
         Height          =   225
         Left            =   6720
         TabIndex        =   47
         Top             =   315
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Line Line9 
         X1              =   10720
         X2              =   10720
         Y1              =   1680
         Y2              =   5145
      End
      Begin VB.Line Line8 
         X1              =   210
         X2              =   10710
         Y1              =   5160
         Y2              =   5160
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
         TabIndex        =   32
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
         TabIndex        =   26
         Top             =   735
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito de Salida:"
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
         Left            =   7665
         TabIndex        =   25
         Top             =   1050
         Width           =   960
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
         TabIndex        =   24
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
         TabIndex        =   18
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
      TabIndex        =   23
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "FOREMIT09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTNEXT_Click()
    '
    Call SIGUERETRA1
    If ULTREG&("!BOLREDET") < 1 Then
        List1.Clear
        List1.Visible = False
        Command5.Enabled = False
        BOTNEXT.Enabled = True
        Text9.SetFocus
        Exit Sub
    End If
    List1.Clear
    List1.Visible = True
    V& = 0
    For U& = 1 To ULTREG&("!BOLREDET")
        X$ = REGLEIDO$("!BOLREDET", U&)
        RELI$ = Space$(80)
        CI% = CVI(Mid$(X$, 83, 2))
        If TRIM$(CODEXT$(CI%)) <> "" Then
            V& = V& + 1
            Call GRAREG("!BOLREDET", X$, V&)
            Call REPLA(RELI$, CODEXT$(CI%), 1, 15)
            Call REPLA(RELI$, DESCRIT0$(CI%), 16, 30)
            Call REPLA(RELI$, UNIMED$(CI%), 49, 4)
            CANS$ = CSTRING$(Mid$(X$, 89, 8), 3, 10, 1, 2)
            Call REPLA(RELI$, CANS$, 53, 10)
            Call REPLA(RELI$, Mid$(X$, 107, 12), 66, 12)
            NOPE$ = CSTRING$(Mid$(X$, 151, 4), 3, 7, 0, 2)
            Call REPLA(RELI$, NOPE$, 78, 7)
            List1.AddItem RELI$
        End If
    Next U&
    Call SETULTREG("!BOLREDET", V&)
    '
    Command5.Enabled = True
    Text15.SetFocus
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


Private Sub Command2_Click()
    '
    If ULTREG&("!BOLREDET") > 0 Then
       Call CIERRARCH("!BOLREDET")
       Call BLANARCH("!BOLREDET")
       List1.Clear
       List1.Visible = False
       Command5.Enabled = False
       BOTNEXT.Enabled = True
       Text9.SetFocus
       Exit Sub
    End If
    '
    Screen.MousePointer = 1
    Call FINAL("")
    '
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    OPREMIT09.Show 1
    Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    '
    NCXI = Val(Text1.TEXT)
    If NCXI < 1 Or NCXI > 32767 Then
        Call MENSERR(24, "Numero de Cuenta no Válido")
        Text1.SetFocus
        GoTo 99
    End If
    '
    NC% = NCXI
    If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
        Call MENSERR(24, "Destino no Válido.")
        Text1.SetFocus
        GoTo 99
    End If
    '
    If COMALTER%("Genera Archivo de Transferencia\a la Sucursal " + TRIM$(ECOARCH$("DEUDOR1", MKI$(NC%))) + " ?\\Cancelar\Generar Transferencia") = 2 Then
      Call GENARTRA(NC%)
    End If
    '
99  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call PRIREMITRA
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    DP1% = 99
    If DEPOEXPE% > 0 Then DP1% = DEPOEXPE%
    Text10.TEXT = TRIM$(Str$(DP1%))
    Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    '
    CMO$ = TRIM$(CONTROL$("", "CODRETRA"))
    If CMO$ = "" Then CMO$ = "RT"
    '
    NROPRO& = 0
    Call NUREMITRA(NROPRO&)
    Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '
    Call CIERRARCH("!BOLREDET")
    Call BLANARCH("!BOLREDET")
    List1.Clear
    List1.Visible = False
    Command5.Enabled = False
    BOTNEXT.Enabled = True
    '
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
    OPRETRA09.Show 1
    Command6.Enabled = True
    '
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
'   If DERACCE%("RECUPESTO", "Recuperacion de Stocks") < 1 Then
'        Exit Sub
'   End If
    Call SELECHO("TADEPOSI/Deposito Destino")
    NUDEPO% = Val(VALACT1$("TADEPOSI"))
    If NUDEPO% < 1 Then GoTo 99
    '
    Call SETULTREG("!MOVISTO", 0)
    Call SETULTREG("MOVISTO", 0)
    '
    Call ABRESTOCKS
    CONDI$ = "DEPOMOVI = " & NUDEPO%
    CANTREG& = CantRegistros("MOVISTO", CONDI$)
    
    I& = 0
    SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE DEPOMOVI = " & NUDEPO%
    SQLX$ = SQLX$ + " ORDER BY FechMov ASC"
    Dim MOVITMP As New ADODB.Recordset
    Set MOVITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

    With MOVITMP
     
      Do While Not .EOF
        I& = I& + 1
        Call TRACE(24, I&, CANTREG&)
        FE% = CVINT(!FeReMovi)
        NUMUSU% = XVALO(!NumUsuar)
        FEMOVI% = CVINT(!FechMov)
        CI% = XVALO(!COD_INTE)
        CODIGO$ = CODEXT$(CI%)
        LOTE$ = SAFETEXT$(!IDENLOTE)
        CODMOV$ = SAFETEXT$(!CodiMovi)
        DOCPRI$ = Left(SAFETEXT$(!DoPriCor), 10)
        DOCSEC$ = Left(SAFETEXT$(!DoSecCor), 10)
        Moneda% = XVALO(!MoneVal)
        REFER$ = Left(SAFETEXT$(!ReferCor), 30)
        VALORUN# = XVALO(!ValoUnit)
        NUMCLIEN% = XVALO(!Nume_Clie)
        NUMPROVE% = XVALO(!Nume_Prov)
        CANTIN# = XVALO(!CANTINGRE)
        CANSA# = XVALO(!CantSalid)
        CANSALD# = XVALO(!CantSaldo)
        '
        Call REPLA(TX$, MKI$(FEMOVI%), 1, 2)
        Call REPLA(TX$, MKI$(CI%), 3, 2)
        Call REPLA(TX$, LOTE$, 5, 16)
        Call REPLA(TX$, CODMOV$, 21, 2)
        Call REPLA(TX$, DOCPRI$, 23, 10)
        Call REPLA(TX$, DOCSEC$, 33, 10)
        Call REPLA(TX$, REFER$, 43, 30)
        Call REPLA(TX$, MKI$(VALORUN#), 73, 8)
        Call REPLA(TX$, MKI$(Moneda%), 81, 2)
        Call REPLA(TX$, Chr$(1), 83, 1)
        Call REPLA(TX$, MKI$(NUMCLIEN%), 84, 2)
        Call REPLA(TX$, MKI$(NUMPROVE%), 86, 2)
        Call REPLA(TX$, MKD$(CANTIN#), 88, 8)
        Call REPLA(TX$, MKD$(CANSA#), 96, 8)
        Call REPLA(TX$, MKD$(CANSALD#), 104, 8)
        Call REPLA(TX$, MKI$(FE%), 122, 2)
        Call REPLA(TX$, Chr$(NUMUSU%), 126, 1)
        '
        Call GRAREG("!MOVISTO", TX$, I&)
       .MoveNext
      Loop
    End With
    Call SETULTREG("!MOVISTO", I&)
    '
    Call TRACENTRAL("MOVISTO", "MOVISTO")
99  Call CIERRARCH("*.*")
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    KARDEX$ = "CONSTO07/FIKARDEX07"
    Call CONECRUN("*KARDEX", "Ficha de Movimientos por Pieza")
'   FIKARDEX07.Show 1
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
    Show
    Refresh
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    TEXPE$ = "Remito de Traslado"
    Caption = Caption + " - " + TEXPE$
    '
    EPAC$ = TRIM$(EMPREAC$)
    If InStr(EPAC$, "NEUMANN") > 0 Then
       Caption = "Remito de Traslado a Sucursal Cárdoba"
    End If
    '
    If EPAC$ <> "" Then
        FOREMIT09.Caption = FOREMIT09.Caption + "  -  " + EPAC$
    End If
    '
    If InStr(UCase(EPAC$), "ARCON") > 0 Then
       Command4.Visible = True
    End If
    '
10  If ULTREG("ECODEP") < 1 Then
        Call MENSERR(24, "Falta Definir Tabla de Depósitos")
        Call CARTADEPOSI
        GoTo 10
    End If
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMTRA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Remito de Traslado.")
        OPREMIT09.Show 1
        GoTo 10
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito de Facturación:\  \'" + RMCC$ + "'.")
            OPREMIT09.Show 1
            GoTo 10
        End If
    End If
    '
    FF% = HOY(HO$)
    FOREMIT09.Text9.TEXT = HO$
    '
    DP1% = Val(CONTROL$("", "DEPOEXPE"))
    If DP1% < 1 Then DP1% = 99
    FOREMIT09.Text10.TEXT = TRIM$(Str$(DP1%))
    FOREMIT09.Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    '
    CMO$ = TRIM$(CONTROL$("", "CODRETRA"))
    If CMO$ = "" Then CMO$ = "RT"
    NROPRO& = 0
    Call NUREMITRA(NROPRO&)
    FOREMIT09.Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '
    Refresh
    '
End Sub

Private Sub Label2_DblClick()
    OPREMIT09.Show 1
End Sub

Private Sub List1_DblClick()
    '
    'If Command12.Enabled = False Then
    '   Call MENSERR(24, "Imposible Re-Editar Despacho.\Cancele y Vuelva a Cargar.")
    '   Exit Sub
    'End If
    '
    Call SIGUERETRA2
    If ULTREG&("!BOLREDET") < 1 Then
        List1.Clear
        List1.Visible = False
        Command5.Enabled = False
        BOTNEXT.Enabled = True
        Text1.SetFocus
        Exit Sub
    End If
    List1.Clear
    List1.Visible = True
    V& = 0
    For U& = 1 To ULTREG&("!BOLREDET")
        X$ = REGLEIDO$("!BOLREDET", U&)
        RELI$ = Space$(80)
        CI% = CVI(Mid$(X$, 83, 2))
        If TRIM$(CODEXT$(CI%)) <> "" Then
            V& = V& + 1
            Call GRAREG("!BOLREDET", X$, V&)
            Call REPLA(RELI$, CODEXT$(CI%), 1, 15)
            Call REPLA(RELI$, DESCRIT0$(CI%), 16, 30)
            Call REPLA(RELI$, UNIMED$(CI%), 49, 4)
            CANS$ = CSTRING$(Mid$(X$, 89, 8), 3, 10, 1, 2)
            Call REPLA(RELI$, CANS$, 53, 10)
            Call REPLA(RELI$, Mid$(X$, 107, 12), 66, 12)
            NOPE$ = CSTRING$(Mid$(X$, 151, 4), 3, 7, 0, 2)
            Call REPLA(RELI$, NOPE$, 78, 7)
            List1.AddItem RELI$
        End If
    Next U&
    Call SETULTREG("!BOLREDET", V&)
    '
    Command5.Enabled = True
    Text15.SetFocus
    '
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
             If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
                Text2.TEXT = RASOCLI$(NC%)
                Text3.TEXT = DOMICLI$(NC%)
                Text4.TEXT = CPOSCLI$(NC%)
                Text5.TEXT = LOCACLI$(NC%)
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
    Static CMO$
    VDEFF$ = TRIM$(Text11.TEXT)
    Text11.TEXT = VDEFF$
    '
    If CMO$ = "" Then
       CMO$ = TRIM$(CONTROL$("", "CODRETRA"))
       If CMO$ = "" Then CMO$ = "RT"
    End If
    '
    If Left$(VDEFF$, 3) <> CMO$ + "." Then GoTo 90
    NRONUE& = Val(Mid$(VDEFF$, 4))
    NROPRO& = NRONUE&
    Call NUREMITRA(NROPRO&)
    If NROPRO& <> NRONUE& Then
        Beep
        GoTo 95
    End If
    Exit Sub
    '
90  Beep
    NROPRO& = 0
    Call NUREMITRA(NROPRO&)
95  Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    Text11.SetFocus
    '
End Sub

Private Sub Text2_gotfocus()
    Text9.SetFocus
End Sub

Private Sub Text3_GotFocus()
    Text9.SetFocus
End Sub

Private Sub Text4_GotFocus()
    Text9.SetFocus
End Sub

Private Sub Text5_GotFocus()
    Text9.SetFocus
End Sub

Private Sub Text6_Click()
    Text9.SetFocus
End Sub

Private Sub Text6_DblClick()
    OPREMIT09.Show 1
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

Sub NUREMITRA(NROPRO&)
    '
    NROPRO& = ProNroComprobante("RC")
    Exit Sub
    '
    Screen.MousePointer = 11
    '
    EPAC$ = EMPREAC$
    CMO$ = TRIM$(CONTROL$("", "CODRETRA"))
    If CMO$ = "" Then CMO$ = "RT"
    '
    FENUE$ = CONTROL$("", "FENURET")
    FLIM% = FECHANUM(FENUE$)
    If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
    FLID = FETEXCOR1$(FLIM%)
    '
    Call ABRESTOCKS
    strsql = "select max(DoPriLar) as maximo from movisto with(nolock) "
    strsql = strsql & "WHERE FechMov >= '" & FLID & "'"
    strsql = strsql & " AND CodiMovi = '" & CMO$ & "' "
    
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.CursorLocation = adUseClient
    Set MOVITMP = FLEXCONN.Execute(FILTSQL$(strsql))
    '
    On Error Resume Next
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Determinar el Próximo Número.\Debe Introducirlo Manualmente.")
       NROMA& = 0
       GoTo 84
    End If
    On Error GoTo 0
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "Imposible Determinar Número de Próximo Remito")
        On Error GoTo 0
        GoTo 85
      End If
      On Error GoTo 0
      '
      On Error Resume Next
      NROMA& = XVALO(Mid$(!maximo, 4))
      On Error GoTo 0
    End With
    MOVITMP.Close
    Set MOVITMP = Nothing
    '
84  If NROMA& < 0 Then NROMA& = 0
85  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
'    CAMOLE% = 0
'    NROMA& = 0
'    RETOLE& = 0
'    For U& = ULTREG&("MOVISTO") To 1 Step -1
'        XX$ = REGLEIDO$("MOVISTO", U&)
'        RETOLE& = RETOLE& + 1
'        If RETOLE& > 2048 Then
'           If CAMOLE% >= 1 Then
'              GoTo 90
'           End If
'        End If
'        '
'        If UCase$(Mid$(XX$, 21, 2)) = CMO$ Then
'            If InStr(EPAC$, "NEUMANN") > 0 Then
'              If Mid$(XX$, 23, 2) <> CMO$ Then GoTo 80
'              If Abs(Val(Mid$(XX$, 26, 7)) - Val(Mid$(XX$, 36, 7))) < 0.1 Then
'                GoTo 80
'              End If
'            End If
'            CAMOLE% = CAMOLE% + 1
'            If CAMOLE% > 256 Then GoTo 90
'            '
'            FEMO% = CVI(Left$(XX$, 2))
'            If FEMO% >= FLIM% Then
'                NROLE& = Val(Mid$(XX$, 26, 7))
'                If NROLE& = NROPRO& Then
'                    NROPRO& = 0
'                End If
'                If NROLE& > NROMA& Then NROMA& = NROLE&
'            End If
'        End If
'80  Next U&
'    '
'90  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
'    Screen.MousePointer = 1
    '
End Sub
'
Sub SIGUERETRA1()
    '
    Dim CAPRECON(32767)
    NCXI = Val(Text1.TEXT)
    If NCXI < 1 Or NCXI > 32767 Then
        Call MENSERR(24, "Numero de Cuenta no Válido")
        Text1.SetFocus
        Exit Sub
    End If
    '
    NC% = NCXI
    If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
        Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Clientes")
        Text1.SetFocus
        Exit Sub
    End If
    '
    FE$ = TRIM$(Text9.TEXT)
    FF% = FECHANUM(FE$)
    If FF% < 1 Or FF% > HOY(HO$) + 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Call CIERRARCH("*.*")
    Call BLANARCH("!BOLREDET")
    '
    Call APERBASDAT("PEDCLI")
    CONDI$ = " NroClie = " & NC% & " AND STATUS <= 1 AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= '" & HOS$ & "')) "
    CANTREG& = CantRegistros("PEDDETA", CONDI$)
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE NroClie = " & NC% & " "
'    SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
    SQLX$ = SQLX$ + " AND STATUS <= 1 "   ' APROBADOS
    SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= '" & HOS$ & "')) "
    SQLX$ = SQLX$ + " ORDER BY NroPed ASC, FeSolEnt ASC, NuOrItem ASC " 'ESTA LINEA LA AGREGUE
    '
    Dim PEDTEMP As ADODB.Recordset
    Set PEDTEMP = New ADODB.Recordset
    PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    
    DEPOSALI% = Val(Text10): If DEPOSALI% < 1 Then DEPOSALI% = DEPOEXPE%
    
    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        Call TRACE(24, JJ&, CANTREG&)
        CIPD% = !CODIINT
        CODEX$ = CODEXT$(CIPD%)
        If CIPD% > 0 And CIPD% <= NUMAS% Then
          CAPEN# = XVALO(!CANPED) - XVALO(!CANTENT)
            If CAPEN# < 0 Then CAPEN# = 0
          If !TipoPed = "ABIERTO" Then CAPEN = 999999
          FEMIPED% = CVINT(!FECHEMIS)
          FENTRESOL% = CVINT(!FESOLENT)
          PREUNI = !PreUnid * (100 - PORDESCLI(NC%)) / 100
          MONEMA% = !MoneEmis
          '
          CAPOSIB# = CAPEN#
          CAPOSI2# = STODEPOS(CIPD%, DEPOSALI%) - CAPRECON(CIPD%)
          If CAPOSI2# < CAPOSIB# Then CAPOSIB# = CAPOSI2#
          If CAPOSIB# < 0 Then CAPOSIB# = 0
          CAPRECON(CIPD%) = CAPRECON(CIPD%) + CAPOSIB#
          '
          Y$ = REGBLAN$("BOLREDET")
          Call REPLA(Y$, MKI$(FEMIPED%), 81, 2)
          Call REPLA(Y$, MKI$(CIPD%), 83, 2)
          Call REPLA(Y$, UNIMED$(CIPD%), 85, 4)
          Call REPLA(Y$, MKD$(CAPOSIB#), 89, 8)
          Call REPLA(Y$, MKS$(!NROPED), 151, 4)
          Call REPLA(Y$, MKI$(FENTRESOL%), 155, 2)
          Call REPLA(Y$, MKI$(MONEMA%), 157, 2)
          Call REPLA(Y$, MKS$(PREUNI), 159, 4)
          Call REPLA(Y$, MKD$(CDbl(CAPEN)), 163, 8)
            U& = !NuorItem
          Call REPLA(Y$, MKS$(U&), 189, 4)
          JJ& = JJ& + 1
          Call GRAREG("!BOLREDET", Y$, JJ&)
          '
        End If
      .MoveNext
      Loop
      On Error GoTo 0
    End With
    PEDTEMP.Close
    Set PEDTEMP = Nothing
    '
'    FIN& = ULTREG&("PEDDETA")
'    RPEDX$ = RECFILT$("PEDDETA", 3, MKI$(NC%))
'    NROPEDA& = 0
'    DEPOSALI% = Val(Text10): If DEPOSALI% < 1 Then DEPOSALI% = DEPOEXPE%
'    FIN& = Len(RPEDX$)
'    For UU& = 1 To Len(RPEDX$) Step 4
'        U& = CVS(Mid$(RPEDX$, UU&, 4))
'        If FIN& > 0 Then
'            Call TRACE(20, UU&, FIN&)
'        End If
'        X$ = REGLEIDO$("PEDDETA", U&)
'        CIPD% = CVI(Mid$(X$, 33, 2))
'        If CVI(Mid$(X$, 7, 2)) = NC% Then
'          Status% = Asc(Mid$(X$, 27, 1))
'          If Status% <= 1 Then
'            CAN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
'            If CAN > 0.05 Then
'              '
'              CAPOSIB# = CAN
'              CAPOSI2# = STODEPOS(CIPD%, DEPOSALI%) - CAPRECON(CIPD%)
'              If CAPOSI2# < CAPOSIB# Then CAPOSIB# = CAPOSI2#
'              If CAPOSIB# < 0 Then CAPOSIB# = 0
'              CAPRECON(CIPD%) = CAPRECON(CIPD%) + CAPOSIB#
'              '
'              NROPED& = CVS(Left$(X$, 4))
'              MEDIDAS$ = Mid$(X$, 59, 12)
'              Y$ = REGBLAN$("BOLREDET")
'              Call REPLA(Y$, MKI$(CIPD%), 83, 2)
'              Call REPLA(Y$, MKD$(CAPOSIB#), 89, 8)
'              Call REPLA(Y$, MKS$(NROPED&), 151, 4)
'              Call REPLA(Y$, "    ", 155, 4)
'              Call REPLA(Y$, MKD$(CDbl(CAN)), 163, 8)
'              Call REPLA(Y$, MKS$(U&), 189, 4)
'              Call REPLA(Y$, MEDIDAS$, 193, 12)
'              Call REGAPP("!BOLREDET", Y$)
'            End If
'          End If
'        End If
'    Next UU&
    '
    FIN& = ULTREG&("!BOLREDET")
    For U& = 1 To FIN&
5     X$ = REGLEIDO$("!BOLREDET", U&)
      NPX& = CVS(Mid$(X$, 151, 4))
      For V& = U& + 1 To FIN&
        Y$ = REGLEIDO$("!BOLREDET", V&)
        NPY& = CVS(Mid$(Y$, 151, 4))
        If NPY& < NPX& Then
          Call GRAREG("!BOLREDET", Y$, U&)
          Call GRAREG("!BOLREDET", X$, V&)
          GoTo 5
        End If
      Next V&
    Next U&
    '
    For U& = 1 To FIN&
      X$ = REGLEIDO$("!BOLREDET", U&)
      CIIX% = CVI(Mid$(X$, 83, 2))
      For V& = U& + 1 To FIN&
        Y$ = REGLEIDO$("!BOLREDET", V&)
        CIIY% = CVI(Mid$(Y$, 83, 2))
        If CIIY% = CIIX% Then
          Call REPLA(Y$, "****", 155, 4)
          Call GRAREG("!BOLREDET", Y$, V&)
        End If
      Next V&
    Next U&
    Call CIERRARCH("!BOLREDET")
    '
10  BOTNEXT.Enabled = False
    Screen.MousePointer = 1
    Call SIGUERETRA2
    '
End Sub
'

Sub SIGUERETRA2()
    '
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Remito de Traslado - Detalle de Material a Despachar"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
    '
    VERIST% = (-1)
    If CONTROL$("REMSTAN", "VERISTOK") = "VERISTOK" Then
       VERIST% = 1
    End If
    '
10  VTB% = VENTABU%("DETAREMI" + ATRI$)
    If VTB% < 0 Then
       If COMALTER%("Realmente Desea Cancelar Este Remito ?\\No, Continuar\Si, Cancelar") <> 2 Then GoTo 10
       Call BLANARCH("!BOLREDET")
       Exit Sub
    End If
    '
    ATXX% = 0
    FIN& = ULTREG&("!BOLREDET")
    For U& = 1 To FIN&
       YY$ = REGLEIDO$("!BOLREDET", U&)
       CI% = CVI(Mid$(YY$, 83, 2))
       CAN# = CVD(Mid$(YY$, 89, 8))
       CAPEPE# = CVD(Mid$(YY$, 163, 8))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If CAN# > 0.05 Then
             If CAN# > CAPEPE# + 0.05 Then
               ' redistribuir
               For JJ& = U& + 1 To FIN&
                 ZZ$ = REGLEIDO$("!BOLREDET", JJ&)
                 If CVI(Mid$(ZZ$, 83, 2)) = CI% Then
                   CANUEN# = CAN# - CAPEPE# + CVD(Mid$(ZZ$, 89, 8))
                   Call REPLA(ZZ$, MKD$(CANUEN#), 89, 8)
                   Call GRAREG("!BOLREDET", ZZ$, JJ&)
                   YY$ = REGLEIDO$("!BOLREDET", U&)
                   Call REPLA(YY$, MKD$(CAPEPE#), 89, 8)
                   Call GRAREG("!BOLREDET", YY$, U&)
                   ATXX% = 1
                   GoTo 19
                 End If
               Next JJ&
             End If
           End If
         End If
       End If
19  Next U&
    '
    If ATXX% > 0 Then
      Call COMUNI("FLEXOFT ha Re-Distribuido la Entrega\Según los Pedidos Pendientes.\  \Pulse 'Aceptar' para ver la Distribución.")
      GoTo 10
    End If
    '
    CAITPE% = 0: UJ& = 0: AJUCAN% = 0
    VALODECLA# = 0
    DEPOSALI% = Val(Text10): If DEPOSALI% < 1 Then DEPOSALI% = DEPOEXPE%
    NORVENTA = "": NORVEA& = 0
    For U& = 1 To ULTREG&("!BOLREDET")
       YY$ = REGLEIDO$("!BOLREDET", U&)
       REPED& = 0
       If Mid$(YY$, 189, 4) <> "    " Then
         REPED& = CVS(Mid$(YY$, 189, 4))
       End If
       CI% = CVI(Mid$(YY$, 83, 2))
       CAN# = CVD(Mid$(YY$, 89, 8))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If CAN# > 0.05 Then
             'Call VERISAL(CI%)
             Call VESARRAS(CI%)
             CADISPONI# = STODEPOS(CI%, DEPOSALI%) ' - STOCKASI(CI%)
             If CADISPONI# < 0 Then CADISPONI# = 0
             If CADISPONI# - CAN# < (-0.005) Then
               CONTI% = COMALTER%("ATENCION !!!\Stock Insuficiente en Código '" + TRIM$(CODEXT$(CI%)) + "'\" + DESCRIT0$(CI%) + "\\Ajustar\Continuar\Abortar")
               If CONTI% <= 1 Then ' ajusta cantidad
                    AJUCAN% = 1
                    Call REPLA(YY$, MKD$(CADISPONI#), 89, 8)
                   ElseIf CONTI% = 3 Then
                    Call SETULTREG("!BOLREDET", 0)
                    Call CIERRARCH("!BOLREDET")
                    Exit Sub
               End If
             End If
             UJ& = UJ& + 1
             Call GRAREG("!BOLREDET", YY$, UJ&)
             '
             VALUNI = 0.7 * PRELISTA("LIPRE001", CI%)
             If VALUNI < 0.001 Then VALUNI = COSUNI(CI%)
             VALODECLA# = VALODECLA# + VALUNI * CAN#
             '
             NORVE& = CVS(Mid$(YY$, 151, 4))
             If NORVE& > 0 Then
               If NORVE& <> NORVEA& Then
                 NORVEA& = NORVE&
                 If NORVENTA <> "" Then NORVENTA = NORVENTA + " - "
                 NORVENTA = NORVENTA + TRIM$(Str$(NORVE&))
               End If
             End If
             '
           End If
         End If
       End If
    Next U&
    '
    If UJ& < ULTREG&("!BOLREDET") Or AJUCAN% > 0 Then
       Call SETULTREG("!BOLREDET", UJ&)
       Call CIERRARCH("!BOLREDET")
       GoTo 10
    End If
    '
    FOREMIT09.Text17.TEXT = CSTRING$(MKD$(VALODECLA#), 3, 10, 2, 2)
    FOREMIT09.Text17.Refresh
    '
End Sub
'
Sub PRIREMITRA()
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMTRA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Remito de Traslado.")
        Exit Sub
    End If
    '
    ORCOP$(1) = "ORIGINAL"
    ORCOP$(2) = "DUPLICADO"
    ORCOP$(3) = "TRIPLICADO"
    ORCOP$(4) = "CUADRUPLICADO"
    ORCOP$(5) = "QUINTUPLICADO"
    ORCOP$(6) = "SEXTUPLICADO"
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito de Traslado:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    CONTA% = ULTREG&("!BOLREDET")
    If CONTA% < 1 Then
        Call COMUNI("No se ha Ingresado Detalle")
        Exit Sub
    End If
    '
    Screen.MousePointer = 11      ' Registrando Remito
    '
    AINDI% = 0
    CMO$ = TRIM$(CONTROL$("", "CODRETRA"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    NC% = Val(FOREMIT09.Text1.TEXT)
    REMI$ = TRIM$(FOREMIT09.Text11.TEXT)
    NROO& = Val(Mid$(REMI$, 4))
    FF% = FECHANUM(FOREMIT09.Text9.TEXT)
    DEPO% = Val(FOREMIT09.Text10.TEXT)
    '
    DEPOSUC% = 0
    EPAC$ = UCase$(TRIM$(EMPREAC$))
    If InStr(EPAC$, "NEUMANN") > 0 Then DEPOSUC% = 98
    If InStr(EPAC$, "ARCON") > 0 Then
       Call SELECHO("TADEPOSI/Deposito Destino")
       DEPOSUC% = Val(VALACT1$("TADEPOSI"))
       If DEPOSUC% < 1 Then GoTo 99
    End If
    '
    RTRASX$ = REGBLAN$("RTRASLA") 'genera encabez
    IDOCUM& = 10000& * (FF% Mod 1000) + 100& * (NROO& Mod 100)
    Call REPLA(RTRASX$, MKI$(FF%), 1, 2)
    Call REPLA(RTRASX$, REMI$, 3, 14)
    Call REPLA(RTRASX$, MKI$(NC%), 17, 2)
    Call REPLA(RTRASX$, MKI$(1), 113, 2)
    Call REPLA(RTRASX$, MKI$(0), 115, 2)
    '
    NUORIT% = 0
    For YX& = 1 To ULTREG&("!BOLREDET")
      YY$ = REGLEIDO$("!BOLREDET", YX&)
      CI% = CVI(Mid$(YY$, 83, 2))
      CAN# = CVD(Mid$(YY$, 89, 8))
      LOTE$ = Mid$(YY$, 107, 12)
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          If CAN# > 0.005 Then
            NUOC = CVS(Mid$(YY$, 151, 4))
            DOSEC$ = ""
            CANTI = (-1) * CAN#
            ' AGREGADO
            NUORIT% = NUORIT% + 1
            MONEVA% = CVI(Mid$(YY$, 157, 2)):  If MONEVA% < 1 Then MONEVA% = 1
            If MONEVA% > ULTREG&("TAMONED") Then MONEVA% = 1
            VUNIST# = CDbl(CVS(Mid$(YY$, 159, 4)))
            NUPEDIORI& = CVS(Mid$(YY$, 151, 4))
            '
            DEPX% = DEPO%: If DEPO% = 99 Then DEPX% = DEPOPRE%(CI%)
            Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, RASOCLI$(NC%), NUORIT%, VUNIST#, MONEVA%, NC%, DEPX%, CANTI, NUPEDIORI&)
            If DEPOSUC% > 0 Then
              Call MOVISTOK(FF%, CI%, LOTE$, "RS", REMI$, REMI$, DOSEC$, DOSEC$, "De Deposito Central", NUORIT%, VUNIST#, MONEVA%, NC%, DEPOSUC%, (-1) * CANTI, NUPEDIORI&)
            End If
            '
            
            RTRASY$ = RTRASX$
            IDOCUM& = IDOCUM& + 1
            Call REPLA(RTRASY$, MKS$(IDOCUM&), 19, 4)
            Call REPLA(RTRASY$, LOTE$, 23, 16)
            Call REPLA(RTRASY$, MKI$(CI%), 39, 2)
            Call REPLA(RTRASY$, CODEXT$(CI%), 41, 16)
            Call REPLA(RTRASY$, DESCRIT0$(CI%), 57, 48)
            Call REPLA(RTRASY$, MKD$(CAN#), 105, 8)
            Call REPLA(RTRASY$, MKS$(NUPEDIORI&), 117, 4)
            Call REGAPP("RTRASLA", RTRASY$)
            '
            
'            REGPEDPE& = CVS(Mid$(YY$, 189, 4))
'            If REGPEDPE& > 0 Then
'              If REGPEDPE& <= ULTREG&("PEDDETA") Then
'                CIENT% = CI%
'                CAENTRE = CAN#
'                Y$ = REGLEIDO$("PEDDETA", REGPEDPE&)
'                If CVI(Mid$(Y$, 33, 2)) = CIENT% Then ' debe coincidir codigo
'                  TOENTRE = CAENTRE + CVS(Mid$(Y$, 95, 4))
'                  SALPEN = CVS(Mid$(Y$, 79, 4)) - TOENTRE: If SALPEN < 0 Then SALPEN = 0
'                  '
'                  CASIG = 0: If Mid$(Y$, 75, 4) <> "    " Then CASIG = CVS(Mid$(Y$, 75, 4))
'                  CASIG = CASIG - CAENTRE: If CASIG < 0 Then CASIG = 0
'                  Call REPLA(Y$, MKS$(CASIG), 75, 4)
'                  '
'                  Call REPLA(Y$, MKS$(TOENTRE), 95, 4)
'                  Call REPLA(Y$, MKS$(SALPEN), 99, 4)
'                  Call REPLA(Y$, MKI$(FF%), 103, 2)
'                  Call REPLA(Y$, MKS$(NROO&), 105, 4)
'                  Call GRAREG("PEDDETA", Y$, REGPEDPE&)
'                  '
'                  For UK& = 1 To ULTREG&("PEDDETP")
'                    Z$ = REGLEIDO$("PEDDETP", UK&)
'                    If Left$(Z$, 30) = Left$(Y$, 30) Then
'                      If Mid$(Z$, 33, 42) = Mid$(Y$, 33, 42) Then
'                        Call REPLA(Z$, MKS$(CASIG), 75, 4)
'                        Call REPLA(Z$, MKS$(TOENTRE), 95, 4)
'                        Call REPLA(Z$, MKS$(SALPEN), 99, 4)
'                        Call REPLA(Z$, MKI$(FF%), 103, 2)
'                        Call REPLA(Z$, MKS$(NROO&), 105, 4)
'                        Call GRAREG("PEDDETP", Z$, UK&)
'                      End If
'                    End If
'                  Next UK&
'                  '
'                End If
'              End If
'            End If
          End If
        End If
      End If
    Next YX&
    '
    For YX& = 1 To ULTREG&("!BOLREDET")
       YY$ = REGLEIDO$("!BOLREDET", YX&)
       CI% = CVI(Mid$(YY$, 83, 2))
       CAN# = CVD(Mid$(YY$, 89, 8))
       NUPEDIORI& = CVS(Mid$(YY$, 151, 4))
       NUORITPE% = CVS(Mid$(YY$, 189, 4))
       '
       Call BLANPEDIPE(NC%, FF%, NROREMIT$, CI%, CAN#, NUPEDIORI&, NUORITPE%)
    Next YX&
    '
5   Call CIERRARCH("MOVISTO")
    Call CIERRARCH("RTRASLA")
'    Call CIERRARCH("PEDDETA")
'    Call CIERRARCH("PEDDETP")
    Call BAJALDISCO
    '
    If TRIM$(FORIPRE$) <> "" Then
        '
        CODFOR$ = TRIM$(UCase$(FORIPRE$))
        MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
        MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
        LARHOW = Val(ATRIFORM$(CODFOR$, "LARGOJMM"))
        If LARHOW < 80 Then LARHOW = 25.4 * 12
        '
        INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
        If INTERLI < 1 Then
            Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
            INTERLI = 1
        End If
        '
        DETDESMM = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
        DETHASMM = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
        If DETHASMM < DETDESMM + 4 + INTERLI Then
           DETHASMM = DETDESMM + 4 * INTERLI
        End If
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
        LARHOAN = Printer.Height
        ANCHOAN = Printer.Width
        '
        PREFIREM$ = CONTROL$("", "PREFIRET")
        If PREFIREM$ = "" Then PREFIREM$ = "0001"
        NroRemi$ = TRIM$(Mid$(REMI$, 4))
        While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
        NroRemi$ = PREFIREM$ + "-" + NroRemi$
        '
10      Screen.MousePointer = 11
        For COP% = 1 To CACOPI%
            '
            INIDET& = 1
            '
20          'Printer.Orientation = 1
            Call SETSPOOL("ORIENTATION", "1")
            If ORIENTA% = 2 Then
               'Printer.Orientation = 2
               Call SETSPOOL("ORIENTATION", "2")
            End If
            '
            'Printer.Height = LARHOW * 56.7
            Call SETSPOOL("HEIGHT", Str$(LARHOW * 56.7))
            'Printer.Width = ANCHOAN
            Call SETSPOOL("WIDTH", Str$(LARHOW * 56.7))
            '
            Call CARFORWIN
            '
            Call PRINTFORWIN(RASOCLI$(NC%), "RASO-CLI", 0, 0)
            Call PRINTFORWIN(DOMICLI$(NC%), "DOMI-CLI", 0, 0)
            Call PRINTFORWIN(CPOSCLI$(NC%), "CPOS-CLI", 0, 0)
            Call PRINTFORWIN(LOCACLI$(NC%), "LOCA-CLI", 0, 0)
            Call PRINTFORWIN(CUITCLI$(NC%), "CUIT-CLI", 0, 0)
            Call PRINTFORWIN(POSIVAL$(NC%), "PIVA-CLI", 0, 0)
            FEX = FF%
            Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
            Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
            Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
            Call PRINTFORWIN(NroRemi$, "NUME-COM", 0, 0)
            Call PRINTFORWIN(DEPOSAL$, "DEPOSITO", 0, 0)
            Call PRINTFORWIN(NroRemi$, "NRO-REMI", 0, 0)
            '
            Call PRINTFORWIN(TRIM$(FOREMIT09.Text15.TEXT$), "NUBULTOS", 0, 0)
            Call PRINTFORWIN(TRIM$(FOREMIT09.Text8.TEXT$), "TRANSPOR", 0, 0)
            Call PRINTFORWIN(TRIM$(FOREMIT09.Text14.TEXT$), "DOMI-TRA", 0, 0)
            Call PRINTFORWIN(TRIM$(FOREMIT09.Text12.TEXT$), "CUIT-TRA", 0, 0)
            Call PRINTFORWIN(TRIM$(FOREMIT09.Text13.TEXT$), "DOCUCOND", 0, 0)
            Call PRINTFORWIN(TRIM$(FOREMIT09.Text17.TEXT$), "VALO-DEC", 0, 0)
            Call PRINTFORWIN(TRIM$(NORVENTA), "NRO-OCOM", 0, 0)
            '
            Call PRINTFORWIN(ORCOP$(COP%), "ORD-COPI", 0, 0)
            '
            INTEROBS = INTERLI
            If INTERLI > 6 Then INTEROBS = INTERLI / 2
            For KKI% = 0 To 3
                Call PRINTFORWIN(TRIM$(FOREMIT09.Text16(KKI%).TEXT$), "OBSERVA", 0, KKI% * INTEROBS)
            Next KKI%
            '
            YAC = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
            '
            For YX& = INIDET& To ULTREG&("!BOLREDET")
               YY$ = REGLEIDO$("!BOLREDET", YX&)
               CI% = CVI(Mid$(YY$, 83, 2))
               CAN# = CVD(Mid$(YY$, 89, 8))
               'CAC# = CVD(Mid$(YY$, 163, 8))
               If CI% > 0 Then
                  If CI% <= NUMAS% Then
                     If CAN# > 0.005 Then
                        '
                        If YAC > DETHASMM Then
                           INIDET& = YX&
                           'Printer.EndDoc
                           Call SETSPOOL("ENDDOC", "")
                         GoTo 20
                        End If
                        '
                        UNID$ = UNIMED$(CI%)
                        If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                        CANS$ = CSTRING$(MKD$(CAN#), 3, 12, 1, 2)
                        'If CAC# >= 0.05 Then
                        '   CANS$ = CSTRING$(MKD$(CAC#), 3, 12, 1, 2)
                        '   Call REPLA(CANS$, "(*)", 1, 3)
                        'End If
                        NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
                        NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
                        '
                        Call PRINTFORWIN(TRIM$(CODEXT$(CI%)), "COD-MAT", 0, YAC)
                        Call PRINTFORWIN(UNID$, "UNIMED", 0, YAC)
                        Call PRINTFORWIN(NULOTE$, "NRO-LOTE", 0, YAC)
                        Call PRINTFORWIN(CANS$, "CANTIDAD", 0, YAC)
                        Call PRINTFORWIN(NUOCO$, "N-OCOMPR", 0, YAC)
                        Call PRINTFORWIN(DESCRIT0$(CI%), "DES-MAT", 0, YAC)
                        '
                        YAC = YAC + INTERLI
                        '
                        NCIY% = NC%: CIIY% = CI%
                        Call DATADI(NCIY%, CIIY%, ORCO$, NORI$, NPLA$)
                        '
                        If NPLA$ <> "" Then
                            Call PRINTFORWIN("Codigo Plano: " + NPLA$, "DES-MAT", 0, YAC)
                            YAC = YAC + INTERLI
                        End If
                        '
                        If ORCO$ <> "" Then
                            Call PRINTFORWIN("Nro.O.Compra: " + ORCO$, "DES-MAT", 0, YAC)
                            YAC = YAC + INTERLI
                        End If
                        '
                        If NORI$ <> "" Then
                            Call PRINTFORWIN("Nro.de Serie: " + NORI$, "DES-MAT", 0, YAC)
                            YAC = YAC + INTERLI
                        End If
                        '
                     End If
                  End If
               End If
            Next YX&
            '
            'Printer.EndDoc
            Call SETSPOOL("ENDDOC", "")           '
            '
        Next COP%
        '
        Call CIERRARCH("*.*")
        '
        DEDOC$ = "Remito Traslado " + TRIM$(NroRemi$) + " " + TRIM$(FOREMIT09.Text1.TEXT) + "-" + TRIM$(FOREMIT09.Text2.TEXT)
        Call SAVESPOOL(DEDOC$, OKX%)      ' guarda el spool de impresion
        '
        TTXX$ = "Remito Numero\'" + NroRemi$ + "'\Emitido y Registrado\\Repetir Impresion\Otro Remito"
        '
        Screen.MousePointer = 1
        If COMALTER%(TTXX$) = 1 Then GoTo 10
        '
        Printer.Orientation = PORIAN
        Printer.Height = LARHOAN
        Printer.Width = ANCHOAN
        '
    End If
    '
    If InStr(UCase(EPAC$), "ARCON") > 0 Then
      If COMALTER%("Puede Ahora Generar\la Transferencia Electrónica de Datos\\Generar Ahora\Continuar") <> 2 Then
        Call GENARTRA(NC%)
      End If
    End If
    '
99  Call CIERRARCH("*.*")
    '
End Sub
'
Sub GENARTRA(NCLIE%)
    '
    Screen.MousePointer = 11
    Call BLOQARCH("ATRASUC")
    Call BLANARCH("ATRASUC")
    '
    Call TRANTRA("MASTER")
    Call TRANTRA("INVERT")
    Call TRANTRA("ECOMAST")
    Call TRANTRA("PUNCODIG")
    Call TRANTRA("REPLACOC")
    Call TRANTRA("PROVINC")
    '
    Screen.MousePointer = 11
    FIN& = ULTREG("RTRASLA")
    For U& = 1 To ULTREG("RTRASLA")
      XX$ = REGLEIDO$("RTRASLA", U&)
      If CVI(Mid$(XX$, 17, 2)) = NCLIE% Then ' COINCIDE CLIENTE
        VENVI% = CVI(Mid$(XX$, 115, 2))     ' veces enviado
        REMI& = XVALO(Mid$(XX$, 6, 11))
        If VENVI% < 5 Then
          Call REPLA(XX$, MKI$(VENVI% + 1), 115, 2)
          Call GRAREG("RTRASLA", XX$, U&)
          '
          TTXX$ = Space$(8) + String$(128, 0)
          Call REPLA(TTXX$, "RTRASLA", 1, 8) 'esto es lo que
          Call REPLA(TTXX$, XX$, 9, 128)
          Call REGAPP("ATRASUC", TTXX$)
        Else
         A = B
        End If
      End If
    Next U&
    Call CIERRARCH("RTRASLA")
    Call CIERRARCH("ATRASUC")
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 1
    '
35  OPX% = COMALTER%("Archivo de Transferencia\\Enviar Vía Web\Descargar en Disco")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
        NC1% = XVALO(CONTROL("RETRASLA", "SUC1CLIE"))
        RUTA1$ = TRIM$(CONTROL("RETRASLA", "SUC1RUTA"))
        
        NC2% = XVALO(CONTROL("RETRASLA", "SUC2CLIE"))
        RUTA2$ = TRIM$(CONTROL("RETRASLA", "SUC2RUTA"))
        
        NCIX% = XVALO(Text1)
        Select Case NCIX%
            Case NC1%: ORIGARCH$ = RUTA1$
            Case NC2%: ORIGARCH$ = RUTA2$
            Case Else:
                Call COMUNI("El Cliente Actual No Se Corresponde con las Sucursales Configuradas")
                GoTo 35
        End Select
        GoTo 45
'        CARPETASINCRO$ = TRIM$(Control("", "CARPEINT"))
'        If CARPETASINCRO$ <> "" Then
'            ORIGARCH$ = CARPETASINCRO$
'            On Error GoTo 50
'            GoTo 45
'        Else
'            Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio")
'            GoTo 35
'        End If
    End If
        
    '03/04/07 (OT-07-0147)
    FILEORIG.Caption = "Destino del Archivo"
    FILEORIG.Show 1
    If FILEORIG.Label1.Caption = "" Then GoTo 99
    '
40  LU$ = LUDAT$
    On Error GoTo 50
    ORIGARCH$ = Left$(FILEORIG.Label1.Caption, 2)
    If Mid$(ORIGARCH$, 2, 1) <> ":" Then
       Call MENSERR(24, "Unidad '" + ORIGARCH$ + "' No Disponible.")
       GoTo 99
    End If
45  FileCopy LU$ + "ATRASUC.DAT", ORIGARCH$ + "\ATRASUC.DAT"
    On Error GoTo 0
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call COMUNI("Archivo de Transferencia Electrónica\para " + TRIM$(RASOCLI$(NC%)) + "\Generado Exitosamente en '" + REPLACAR$(ORIGARCH$, "\", "/") + "\'.")
    GoTo 99
    '
50  Resume 51
51  On Error GoTo 0
    Call MENSERR(24, "Error de Copiado a Unidad.\   \ Intente de Nuevo.")
    GoTo 99
    '
99  Screen.MousePointer = 1
End Sub

Sub TRANTRA(ATRAN$)
    '
    Dim BUFLE As String * 128
    '
    ATRAN1$ = ATRAN$
    Select Case ATRAN$
        Case "MASTER":
                        Dim RS As ADODB.Recordset
                        
                        REG& = 1
                        Call COPYSTRU("MASTER", "MASTER1")
                        Call SETULTREG("MASTER1", 0)
                        
                        XX$ = REGBLAN$("MASTER")
                        Call REPLA(XX$, MKI$(NUMAS%), 1, 2)
                        Call GRAREG("MASTER1", XX$, 1)
                    
                        SQL$ = "SELECT * FROM MASTER"
                        SQL$ = SQL$ + " ORDER BY Codint"
                        Set RS = FLEXCONN.Execute(SQL$)
                        With RS
                            Do While Not .EOF
                                CI% = XVALO(!CODINT)
                                DENOMINACION$ = SAFETEXT$(!Descripcion)
                                GRUIVA% = XVALO(!GRUPIVA)
                                CIEXT$ = SAFETEXT$(!CODIGO)
                                TIPOMATE% = XVALO(!TIMATE)
                                CRITREPOS% = XVALO(!CRITREP)
                                Monecos% = XVALO(!Monecos)
                                TCOLOR$ = SAFETEXT$(!Color)
                                TTALLE$ = SAFETEXT$(!TALLE)
                                    FECHA$ = FECHATEX$(CVINT(!FEREVIS))
                                CCOMPO% = Val(FECHA$)
                                PESU = XVALO(!PESUNI)
                                LOTEMAI% = XVALO(!STOMAX)
                                STOCKMIN# = XVALO(!STOMIN)
                                LOTEREP% = XVALO(!LOREPOS)
                                Unidad$ = SAFETEXT$(!UMEDIDA)
                                COSTO# = XVALO(!COSUNI)
                                FECHACOS% = CVINT(!FECOSTO)
                                Revision$ = SAFETEXT$(!Revision)
                                FECHREV% = CVINT(!FEREVIS)
                                PROBHAB% = XVALO(!PROVHABI)
                                GRUPOCON% = XVALO(!Grucoven)
                                NACIONA% = XVALO(!PORNACIO)
                                STAT% = XVALO(!Status)
                                '
                                If CI% = 1498 Or CI% = 1499 Then
                                AA = BB
                                End If
                                XX$ = REGBLAN$("MASTER")
                                Call REPLA(XX$, MKI$(CI%), 1, 2)
                                Call REPLA(XX$, Chr$(GRUIVA%), 3, 1)
                                Call REPLA(XX$, DENOMINACION$, 5, 64)
                                Call REPLA(XX$, Space$(8), 69, 8)
                                Call REPLA(XX$, CIEXT$, 77, 16)
                                If STAT% < 0 Then Call REPLA(XX$, Chr$(96), 92, 1)
                                Call REPLA(XX$, Chr$(TIPOMATE%), 93, 1)
                                Call REPLA(XX$, Chr$(CRITREPOS%), 94, 1)
                                Call REPLA(XX$, Chr$(Monecos%), 95, 1)
                                Call REPLA(XX$, Chr$(CCOLOR%(TCOLOR$)), 96, 1)
                                Call REPLA(XX$, Chr$(CTALLE%(TTALLE$)), 97, 1)
                                Call REPLA(XX$, Chr$(CCOMPO%), 98, 1)
                                Call REPLA(XX$, MKS$(PESU), 99, 4)
                                Call REPLA(XX$, MKI$(LOTEMAI%), 103, 2)
                                Call REPLA(XX$, MKI$(STOCKMIN#), 105, 2)
                                Call REPLA(XX$, MKI$(LOTEREP%), 107, 2)
                                Call REPLA(XX$, Unidad$, 109, 2)
                                Call REPLA(XX$, MKS$(COSTO#), 111, 4)
                                Call REPLA(XX$, MKI$(FECHACOS%), 115, 2)
                                Call REPLA(XX$, Revision$, 117, 4)
                                Call REPLA(XX$, MKI$(FECHREV%), 121, 2)
                                Call REPLA(XX$, MKI$(PROBHAB%), 123, 2)
                                Call REPLA(XX$, MKI$(GRUPOCON%), 125, 2)
                                Call REPLA(XX$, MKI$(NACIONA%), 127, 2)
                                REG& = REG& + 1
                                Call GRAREG("MASTER1", XX$, REG&)
                                '
                                .MoveNext
                            Loop
                            Call SETULTREG("MASTER1", REG&)
                            Call CIERRARCH("MASTER1")
                            '
                            Call GENECOMAS1
                            '
                            ATRAN1$ = ATRAN$ + "1"
                        End With
        Case "ECOMAST", "INVERT":
                        ATRAN1$ = ATRAN$ + "1"
    End Select
    '
    LU$ = LUDAT$
    N1% = FILEOPEN(LU$ + ATRAN1$ + ".DAT", 0, 128)
    For KU& = 1 To 1 + LOF(N1%) / 128
      Get #N1%, KU&, BUFLE$
      TTXX$ = Space$(8) + String$(128, 0)
      Call REPLA(TTXX$, ATRAN$, 1, 8)
      Call REPLA(TTXX$, BUFLE$, 9, 128)
      Call REGAPP("ATRASUC", TTXX$)
    Next KU&
    Close #N1%
    '
End Sub

Sub GENECOMAS1()

   Screen.MousePointer = 11
   '
   Call COPYSTRU("ECOMAST", "ECOMAST1")
   Call COPYSTRU("INVERT", "INVERT1")
   '
   FIN& = ULTREG&("MASTER1")
   If FIN& < 2 Then
     Screen.MousePointer = 1
     Exit Sub
   End If
   '
   JI& = 1  ' PARA INVERT1
   REBLI$ = Space$(16) + MKI$(0)
   JE& = 0  ' PARA ECOMAST1
   REBLE$ = Space$(62) + MKI$(0)
   '
   XI$ = REGLEIDO$("INVERT1", 1)
   Call REPLA(XI$, String$(16, 0), 1, 16)
   Call GRAREG("INVERT1", XI$, 1)
   '
   For U& = 2 To FIN&
     Call TRACE(20, U&, FIN&)
     XM$ = REGLEIDO$("MASTER1", U&)
     CII% = U& - 1
     If CVI(Left$(XM$, 2)) <> CII% Then
       Call REPLA(XM$, MKI$(CII%), 1, 2)
       Call GRAREG("MASTER1", XM$, U&)
     End If
     '
     CODE$ = Mid$(XM$, 77, 16)
     DEDE$ = Mid$(XM$, 5, 64)
     '
     XI$ = REBLI$
     Call REPLA(XI$, CODE$, 1, 16)
     Call REPLA(XI$, MKI$(CII%), 17, 2)
     JI& = JI& + 1
     Call GRAREG("INVERT1", XI$, JI&)
     '
     If Mid$(CODE$, 16, 1) <> Chr$(96) Then
       XE$ = REBLE$
       Call REPLA(XE$, CODE$, 1, 16)
       Call REPLA(XE$, DEDE$, 17, 46)
       Call REPLA(XE$, MKI$(CII%), 63, 2)
       JE& = JE& + 1
       Call GRAREG("ECOMAST1", XE$, JE&)
       '
     End If
   Next U&
   '
   NMS% = FIN& - 1
   XM$ = REGLEIDO$("MASTER1", 1)
   Call REPLA(XM$, MKI$(NMS%), 1, 2)
   Call GRAREG("MASTER1", XM$, 1)
   Call CIERRARCH("MASTER1")
   '
   NVI% = JI& - 1
   XI$ = REGLEIDO$("INVERT1", 1)
   Call REPLA(XI$, MKI$(NVI%), 17, 2)
   Call GRAREG("INVERT1", XI$, 1)
   Call SETULTREG("INVERT1", JI&)
   Call CIERRARCH("INVERT1")
   Call FILESORT("INVERT1", "", 1, 1, "ASC")
   '
   Call SETULTREG("ECOMAST1", JE&)
   Call CIERRARCH("ECOMAST1")
   Call FILESORT("ECOMAST1", "", 1, 1, "ASC")
   '
   BFC$ = ""
   NMSX& = NUMAS%
   LOPUCO& = 16 * NMSX&
   BFC$ = String$(LOPUCO&, 0)
   For U& = 1 To NUMAS%
      X$ = REGLEIDO$("MASTER1", U& + 1)
      BFD$ = AJUSTI$(Mid$(X$, 77, 16), 16)
      Mid$(BFC$, 16 * U& - 15, 16) = BFD$
   Next U&
   '
   LOPUCO1& = 2048 * (1 + Int(LOPUCO& / 2048))
   BFC1$ = Left$(BFC$ + String$(4096, 0), LOPUCO1&)
   Call SAVEFILE(LUDAT$ + "PUNCODIG.DAT", BFC1$)

   Screen.MousePointer = 1
End Sub
