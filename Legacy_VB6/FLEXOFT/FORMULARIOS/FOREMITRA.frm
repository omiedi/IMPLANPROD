VERSION 5.00
Begin VB.Form FOREMITRA 
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
      TabIndex        =   19
      Top             =   0
      Width           =   12000
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
         Height          =   830
         Left            =   10920
         Picture         =   "FOREMITRA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Generar Diskette de Transferencia"
         Top             =   5355
         Visible         =   0   'False
         Width           =   855
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
         Height          =   830
         Left            =   10920
         Picture         =   "FOREMITRA.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   4305
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
         Height          =   830
         Left            =   10920
         Picture         =   "FOREMITRA.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Opciones Complementarias - Anulaciones y Setup"
         Top             =   3360
         Width           =   855
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
         Height          =   830
         Left            =   10920
         Picture         =   "FOREMITRA.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Ficha de Movimientos de Material"
         Top             =   1680
         Width           =   855
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
         Height          =   855
         Left            =   10920
         Picture         =   "FOREMITRA.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Aprueba - Graba e Imprime Remito"
         Top             =   7035
         Width           =   855
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0FF&
         Height          =   700
         Left            =   6930
         TabIndex        =   43
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
            TabIndex        =   13
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
            TabIndex        =   44
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
         TabIndex        =   42
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   420
            Width           =   3375
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Height          =   2640
         Left            =   210
         TabIndex        =   37
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
            TabIndex        =   50
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
               TabIndex        =   52
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
               TabIndex        =   51
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
            TabIndex        =   46
            Top             =   315
            Width           =   2640
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   47
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
            TabIndex        =   9
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
            TabIndex        =   40
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
               TabIndex        =   12
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
            TabIndex        =   38
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
               TabIndex        =   10
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
               TabIndex        =   11
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
               TabIndex        =   39
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
            TabIndex        =   41
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
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   315
         Width           =   435
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
         Picture         =   "FOREMITRA.frx":120A
         TabIndex        =   29
         ToolTipText     =   "Ababdona la Aplicación"
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
         Picture         =   "FOREMITRA.frx":1514
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Continúa Carga de Remito."
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   30
         Top             =   1680
         Width           =   10515
         Begin VB.Frame Frame4 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   540
            Left            =   0
            TabIndex        =   32
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
               TabIndex        =   33
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
            TabIndex        =   31
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
               TabIndex        =   45
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
         TabIndex        =   53
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
         TabIndex        =   36
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   20
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
      TabIndex        =   25
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "FOREMITRA"
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
    OPREMITRA.Show 1
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
        Exit Sub
    End If
    '
    NC% = NCXI
    If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
        Call MENSERR(24, "Destino no Válido.")
        Text1.SetFocus
        Exit Sub
    End If
    '
    If COMALTER%("Genera Archivo de Transferencia\a la Sucursal " + TRIM$(ECOARCH$("DEUDOR1", MKI$(NC%))) + " ?\\Cancelar\Generar Transferencia") = 2 Then
      Call GENARTRA(NC%)
    End If
    '
    Command4.Enabled = True
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
    OPRETRAS.Show 1
    Command6.Enabled = True
    '
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CONECRUN("*KARDEX", "Ficha de Movimientos por Pieza")
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
        FOREMITRA.Caption = FOREMITRA.Caption + "  -  " + EPAC$
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
        OPREMITRA.Show 1
        GoTo 10
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito de Facturación:\  \'" + RMCC$ + "'.")
            OPREMITRA.Show 1
            GoTo 10
        End If
    End If
    '
    FF% = HOY(HO$)
    FOREMITRA.Text9.TEXT = HO$
    '
    DP1% = Val(CONTROL$("", "DEPOEXPE"))
    If DP1% < 1 Then DP1% = 99
    FOREMITRA.Text10.TEXT = TRIM$(Str$(DP1%))
    FOREMITRA.Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    '
    CMO$ = TRIM$(CONTROL$("", "CODRETRA"))
    If CMO$ = "" Then CMO$ = "RT"
    NROPRO& = 0
    Call NUREMITRA(NROPRO&)
    FOREMITRA.Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '
    Refresh
    '
End Sub

Private Sub Label2_DblClick()
    OPREMITRA.Show 1
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

Private Sub Text2_GotFocus()
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
    OPREMITRA.Show 1
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
    Screen.MousePointer = 11
    '
    EPAC$ = EMPREAC$
    CMO$ = TRIM$(CONTROL$("", "CODRETRA"))
    If CMO$ = "" Then CMO$ = "RT"
    '
    FENUE$ = CONTROL$("", "FENURET")
    FLIM% = FECHANUM(FENUE$)
    If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
    '
    CAMOLE% = 0
    NROMA& = 0
    RETOLE& = 0
    For U& = ULTREG&("MOVISTO") To 1 Step -1
        XX$ = REGLEIDO$("MOVISTO", U&)
        RETOLE& = RETOLE& + 1
        If RETOLE& > 2048 Then
           If CAMOLE% >= 1 Then
              GoTo 90
           End If
        End If
        '
        If UCase$(Mid$(XX$, 21, 2)) = CMO$ Then
            If InStr(EPAC$, "NEUMANN") > 0 Then
              If Mid$(XX$, 23, 2) <> CMO$ Then GoTo 80
              If Abs(Val(Mid$(XX$, 26, 7)) - Val(Mid$(XX$, 36, 7))) < 0.1 Then
                GoTo 80
              End If
            End If
            CAMOLE% = CAMOLE% + 1
            If CAMOLE% > 256 Then GoTo 90
            '
            FEMO% = CVI(Left$(XX$, 2))
            If FEMO% >= FLIM% Then
                NROLE& = Val(Mid$(XX$, 26, 7))
                If NROLE& = NROPRO& Then
                    NROPRO& = 0
                End If
                If NROLE& > NROMA& Then NROMA& = NROLE&
            End If
        End If
80  Next U&
    '
90  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
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
    FIN& = ULTREG&("PEDDETA")
    RPEDX$ = RECFILT$("PEDDETA", 3, MKI$(NC%))
    NROPEDA& = 0
    DEPOSALI% = Val(Text10): If DEPOSALI% < 1 Then DEPOSALI% = DEPOEXPE%
    FIN& = Len(RPEDX$)
    For UU& = 1 To Len(RPEDX$) Step 4
        U& = CVS(Mid$(RPEDX$, UU&, 4))
        If FIN& > 0 Then
            Call TRACE(20, UU&, FIN&)
        End If
        X$ = REGLEIDO$("PEDDETA", U&)
        CIPD% = CVI(Mid$(X$, 33, 2))
        If CVI(Mid$(X$, 7, 2)) = NC% Then
          STATUS% = Asc(Mid$(X$, 27, 1))
          If STATUS% <= 1 Then
            CAN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            If CAN > 0.05 Then
              '
              CAPOSIB# = CAN
              CAPOSI2# = STODEPOS(CIPD%, DEPOSALI%) - CAPRECON(CIPD%)
              If CAPOSI2# < CAPOSIB# Then CAPOSIB# = CAPOSI2#
              If CAPOSIB# < 0 Then CAPOSIB# = 0
              CAPRECON(CIPD%) = CAPRECON(CIPD%) + CAPOSIB#
              '
              NROPED& = CVS(Left$(X$, 4))
              MEDIDAS$ = Mid$(X$, 59, 12)
              Y$ = REGBLAN$("BOLREDET")
              Call REPLA(Y$, MKI$(CIPD%), 83, 2)
              Call REPLA(Y$, MKD$(CAPOSIB#), 89, 8)
              Call REPLA(Y$, MKS$(NROPED&), 151, 4)
              Call REPLA(Y$, "    ", 155, 4)
              Call REPLA(Y$, MKD$(CDbl(CAN)), 163, 8)
              Call REPLA(Y$, MKS$(U&), 189, 4)
              Call REPLA(Y$, MEDIDAS$, 193, 12)
              Call REGAPP("!BOLREDET", Y$)
            End If
          End If
        End If
    Next UU&
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
             Call VERISAL(CI%)
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
    FOREMITRA.Text17.TEXT = CSTRING$(MKD$(VALODECLA#), 3, 10, 2, 2)
    FOREMITRA.Text17.Refresh
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
    NC% = Val(FOREMITRA.Text1.TEXT)
    REMI$ = TRIM$(FOREMITRA.Text11.TEXT)
    NROO& = Val(Mid$(REMI$, 4))
    FF% = FECHANUM(FOREMITRA.Text9.TEXT)
    DEPO% = Val(FOREMITRA.Text10.TEXT)
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
    RTRASX$ = REGBLAN$("RTRASLA")
    IDOCUM& = 10000& * (FF% Mod 1000) + 100& * (NROO& Mod 100)
    Call REPLA(RTRASX$, MKI$(FF%), 1, 2)
    Call REPLA(RTRASX$, REMI$, 3, 14)
    Call REPLA(RTRASX$, MKI$(NC%), 17, 2)
    Call REPLA(RTRASX$, MKI$(1), 113, 2)
    Call REPLA(RTRASX$, MKI$(0), 115, 2)
    '
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
            '
            DEPX% = DEPO%: If DEPO% = 99 Then DEPX% = DEPOPRE%(CI%)
            Call MOVISTO(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, RASOCLI$(NC%), 0, "$", NC%, DEPX%, CANTI)
            If DEPOSUC% > 0 Then
              Call MOVISTO(FF%, CI%, LOTE$, "RS", REMI$, DOSEC$, "De Deposito Central", 0, "$", NC%, DEPOSUC%, (-1) * CANTI)
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
            Call REGAPP("RTRASLA", RTRASY$)
            '
            REGPEDPE& = CVS(Mid$(YY$, 189, 4))
            If REGPEDPE& > 0 Then
              If REGPEDPE& <= ULTREG&("PEDDETA") Then
                CIENT% = CI%
                CAENTRE = CAN#
                Y$ = REGLEIDO$("PEDDETA", REGPEDPE&)
                If CVI(Mid$(Y$, 33, 2)) = CIENT% Then ' debe coincidir codigo
                  TOENTRE = CAENTRE + CVS(Mid$(Y$, 95, 4))
                  SALPEN = CVS(Mid$(Y$, 79, 4)) - TOENTRE: If SALPEN < 0 Then SALPEN = 0
                  '
                  CASIG = 0: If Mid$(Y$, 75, 4) <> "    " Then CASIG = CVS(Mid$(Y$, 75, 4))
                  CASIG = CASIG - CAENTRE: If CASIG < 0 Then CASIG = 0
                  Call REPLA(Y$, MKS$(CASIG), 75, 4)
                  '
                  Call REPLA(Y$, MKS$(TOENTRE), 95, 4)
                  Call REPLA(Y$, MKS$(SALPEN), 99, 4)
                  Call REPLA(Y$, MKI$(FF%), 103, 2)
                  Call REPLA(Y$, MKS$(NROO&), 105, 4)
                  Call GRAREG("PEDDETA", Y$, REGPEDPE&)
                  '
                  For UK& = 1 To ULTREG&("PEDDETP")
                    Z$ = REGLEIDO$("PEDDETP", UK&)
                    If Left$(Z$, 30) = Left$(Y$, 30) Then
                      If Mid$(Z$, 33, 42) = Mid$(Y$, 33, 42) Then
                        Call REPLA(Z$, MKS$(CASIG), 75, 4)
                        Call REPLA(Z$, MKS$(TOENTRE), 95, 4)
                        Call REPLA(Z$, MKS$(SALPEN), 99, 4)
                        Call REPLA(Z$, MKI$(FF%), 103, 2)
                        Call REPLA(Z$, MKS$(NROO&), 105, 4)
                        Call GRAREG("PEDDETP", Z$, UK&)
                      End If
                    End If
                  Next UK&
                  '
                End If
              End If
            End If
          End If
        End If
      End If
    Next YX&
    '
5   Call CIERRARCH("MOVISTO")
    Call CIERRARCH("RTRASLA")
    Call CIERRARCH("PEDDETA")
    Call CIERRARCH("PEDDETP")
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
        NROREMI$ = TRIM$(Mid$(REMI$, 4))
        While Len(NROREMI$) < 8: NROREMI$ = "0" + NROREMI$: Wend
        NROREMI$ = PREFIREM$ + "-" + NROREMI$
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
            Call PRINTFORWIN(NROREMI$, "NUME-COM", 0, 0)
            Call PRINTFORWIN(DEPOSAL$, "DEPOSITO", 0, 0)
            Call PRINTFORWIN(NROREMI$, "NRO-REMI", 0, 0)
            '
            Call PRINTFORWIN(TRIM$(FOREMITRA.Text15.TEXT$), "NUBULTOS", 0, 0)
            Call PRINTFORWIN(TRIM$(FOREMITRA.Text8.TEXT$), "TRANSPOR", 0, 0)
            Call PRINTFORWIN(TRIM$(FOREMITRA.Text14.TEXT$), "DOMI-TRA", 0, 0)
            Call PRINTFORWIN(TRIM$(FOREMITRA.Text12.TEXT$), "CUIT-TRA", 0, 0)
            Call PRINTFORWIN(TRIM$(FOREMITRA.Text13.TEXT$), "DOCUCOND", 0, 0)
            Call PRINTFORWIN(TRIM$(FOREMITRA.Text17.TEXT$), "VALO-DEC", 0, 0)
            Call PRINTFORWIN(TRIM$(NORVENTA), "NRO-OCOM", 0, 0)
            '
            Call PRINTFORWIN(ORCOP$(COP%), "ORD-COPI", 0, 0)
            '
            INTEROBS = INTERLI
            If INTERLI > 6 Then INTEROBS = INTERLI / 2
            For kki% = 0 To 3
                Call PRINTFORWIN(TRIM$(FOREMITRA.Text16(kki%).TEXT$), "OBSERVA", 0, kki% * INTEROBS)
            Next kki%
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
        DEDOC$ = "Remito Traslado " + TRIM$(NROREMI$) + " " + TRIM$(FOREMITRA.Text1.TEXT) + "-" + TRIM$(FOREMITRA.Text2.TEXT)
        Call SAVESPOOL(DEDOC$, OKX%)      ' guarda el spool de impresion
        '
        TTXX$ = "Remito Numero\'" + NROREMI$ + "'\Emitido y Registrado\\Repetir Impresion\Otro Remito"
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
      If COMALTER%("Puede Ahora Generar el Diskette\de Transferencia Electrónica de Datos\\Generar Diskette\Continuar") <> 2 Then
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
 FIN& = ULTREG("RTRASLA")
    For U& = 1 To ULTREG("RTRASLA")
      XX$ = REGLEIDO$("RTRASLA", U&)
      If CVI(Mid$(XX$, 17, 2)) = NCLIE% Then ' COINCIDE CLIENTE
        VENVI% = CVI(Mid$(XX$, 115, 2))     ' veces enviado
        If VENVI% < 5 Then
          Call REPLA(XX$, MKI$(VENVI% + 1), 115, 2)
          Call GRAREG("RTRASLA", XX$, U&)
          '
          TTXX$ = Space$(8) + String$(128, 0)
          Call REPLA(TTXX$, "RTRASLA", 1, 8)
          Call REPLA(TTXX$, XX$, 9, 128)
          Call REGAPP("ATRASUC", TTXX$)
        End If
      End If
    Next U&
    Call CIERRARCH("RTRASLA")
    Call CIERRARCH("ATRASUC")
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 1
    '
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
    FileCopy LU$ + "ATRASUC.DAT", ORIGARCH$ + "\ATRASUC.DAT"
    On Error GoTo 0
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call COMUNI("Archivo de Transferencia Electrónica\para " + TRIM$(RASOCLI$(NC%)) + "\Generado Exitosamente en '" + ORIGARCH$ + "\'.")
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
    LU$ = LUDAT$
    N1% = FILEOPEN(LU$ + ATRAN$ + ".DAT", 0, 128)
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
