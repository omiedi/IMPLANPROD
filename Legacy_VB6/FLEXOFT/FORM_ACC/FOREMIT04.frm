VERSION 5.00
Begin VB.Form FOREMIT04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Despacho de Materiales "
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11940
   Begin VB.CommandButton Command12 
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
      Height          =   630
      Left            =   12915
      Picture         =   "FOREMIT04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Descompone Conjuntos en Partes Componentes"
      Top             =   420
      Width           =   650
   End
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
      Height          =   8415
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command15 
         Caption         =   "T L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   7875
         TabIndex        =   76
         ToolTipText     =   "Remito de Texto Libre"
         Top             =   290
         Width           =   300
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1890
         Top             =   0
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9360
         TabIndex        =   73
         Top             =   6000
         Width           =   1335
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
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
            Left            =   450
            TabIndex        =   74
            Top             =   270
            Width           =   540
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Orden  de Suministro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   6930
         TabIndex        =   70
         Top             =   5250
         Width           =   3795
         Begin VB.TextBox Text22 
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
            MaxLength       =   24
            TabIndex        =   71
            Top             =   300
            Width           =   3375
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00006000&
         Height          =   9045
         Left            =   11025
         ScaleHeight     =   8985
         ScaleWidth      =   3210
         TabIndex        =   57
         Top             =   -105
         Width           =   3270
         Begin VB.PictureBox Picture16 
            Height          =   8450
            Left            =   0
            ScaleHeight     =   8385
            ScaleWidth      =   1020
            TabIndex        =   77
            Top             =   0
            Visible         =   0   'False
            Width           =   1080
            Begin VB.Image Image3 
               Height          =   8450
               Left            =   0
               Picture         =   "FOREMIT04.frx":030A
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOREMIT04.frx":26A2
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Generación por Código de Barras"
            Top             =   5110
            Width           =   640
         End
         Begin VB.CommandButton Command4 
            Caption         =   "CCorr"
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
            Picture         =   "FOREMIT04.frx":29AC
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   2350
            Width           =   650
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
            Height          =   690
            Left            =   105
            Picture         =   "FOREMIT04.frx":2CB6
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Configuracion de Funcionamiento"
            Top             =   6550
            Width           =   650
         End
         Begin VB.CommandButton Command2 
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
            Picture         =   "FOREMIT04.frx":30F8
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Terminar - Salir de la Aplicación"
            Top             =   200
            Width           =   650
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
            Height          =   690
            Left            =   120
            Picture         =   "FOREMIT04.frx":3242
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Continúa Carga de Remito."
            Top             =   930
            Width           =   650
         End
         Begin VB.CommandButton Command5 
            Caption         =   "O.K."
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
            Height          =   690
            Left            =   105
            Picture         =   "FOREMIT04.frx":354C
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Aprueba - Graba e Imprime Remito"
            Top             =   7245
            Width           =   650
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Fact."
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
            Picture         =   "FOREMIT04.frx":3BB6
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Acceso Directo a Facturación"
            Top             =   4440
            Width           =   650
         End
         Begin VB.CommandButton Command6 
            Caption         =   "More"
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
            Picture         =   "FOREMIT04.frx":3FF8
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Opciones Complementarias - Listados y Reportes"
            Top             =   5860
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Clien."
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
            Picture         =   "FOREMIT04.frx":4142
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
            Top             =   1660
            Width           =   650
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Note"
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
            Picture         =   "FOREMIT04.frx":4584
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Notas de Clientes"
            Top             =   3040
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Kdx."
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
            Picture         =   "FOREMIT04.frx":488E
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Ficha de Movimientos de Material"
            Top             =   3730
            Width           =   650
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -300
            Picture         =   "FOREMIT04.frx":4B98
            Top             =   7980
            Width           =   1515
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   495
         Left            =   210
         TabIndex        =   55
         Top             =   1710
         Visible         =   0   'False
         Width           =   10515
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   $"FOREMIT04.frx":6ABA
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   0
            TabIndex        =   56
            Top             =   105
            Width           =   10410
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Variante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   5985
         TabIndex        =   43
         Top             =   210
         Width           =   1800
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   45
            Top             =   420
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Facturación"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   210
            TabIndex        =   44
            Top             =   210
            Value           =   -1  'True
            Width           =   1485
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Valor Declarado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   6930
         TabIndex        =   33
         Top             =   5985
         Width           =   2235
         Begin VB.TextBox Text17 
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
            Left            =   480
            TabIndex        =   9
            Top             =   250
            Width           =   1170
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0FFD0&
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
         Height          =   1635
         Left            =   6930
         TabIndex        =   32
         Top             =   6615
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
            Index           =   4
            Left            =   210
            TabIndex        =   54
            Top             =   2500
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
            Index           =   3
            Left            =   210
            TabIndex        =   13
            Top             =   1260
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
            TabIndex        =   12
            Top             =   945
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
            TabIndex        =   11
            Top             =   630
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
            TabIndex        =   10
            Top             =   315
            Width           =   3375
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0FFD0&
         Height          =   3020
         Left            =   210
         TabIndex        =   30
         Top             =   5250
         Width           =   6525
         Begin VB.Frame Frame14 
            BackColor       =   &H00E0FFD0&
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
            Height          =   1170
            Left            =   210
            TabIndex        =   46
            Top             =   1680
            Width           =   2325
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
               Height          =   285
               Left            =   105
               TabIndex        =   51
               Top             =   340
               Width           =   2115
            End
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
               TabIndex        =   47
               Top             =   735
               Width           =   1590
            End
            Begin VB.Label Label7 
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
               Height          =   330
               Left            =   105
               TabIndex        =   53
               Top             =   800
               Width           =   435
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Orden de Compra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   210
            TabIndex        =   40
            Top             =   210
            Width           =   2325
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
               Height          =   285
               Left            =   315
               TabIndex        =   41
               Top             =   315
               Width           =   1695
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Numero de Bultos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   210
            TabIndex        =   39
            Top             =   945
            Width           =   2325
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
               Left            =   525
               TabIndex        =   42
               Top             =   315
               Width           =   1275
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0FFD0&
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
            Height          =   2640
            Left            =   2625
            TabIndex        =   31
            Top             =   210
            Width           =   3690
            Begin VB.Frame Frame5 
               BackColor       =   &H00E0FFD0&
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
               Height          =   1130
               Left            =   105
               TabIndex        =   48
               Top             =   1365
               Width           =   3480
               Begin VB.TextBox Text21 
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
                  Left            =   105
                  TabIndex        =   52
                  Top             =   700
                  Width           =   3270
               End
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
                  Left            =   105
                  TabIndex        =   49
                  Top             =   315
                  Width           =   3270
               End
            End
            Begin VB.Frame Frame10 
               BackColor       =   &H00E0FFD0&
               Caption         =   "Patente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   620
               Left            =   1995
               TabIndex        =   37
               Top             =   735
               Width           =   1590
               Begin VB.TextBox Text18 
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
                  Left            =   105
                  TabIndex        =   38
                  Top             =   210
                  Width           =   1380
               End
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00E0FFD0&
               Caption         =   "C.U.I.T."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   620
               Left            =   105
               TabIndex        =   35
               Top             =   735
               Width           =   1800
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
                  Left            =   105
                  TabIndex        =   36
                  Top             =   210
                  Width           =   1590
               End
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
               Height          =   285
               Left            =   210
               TabIndex        =   8
               Top             =   315
               Width           =   3270
            End
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
         Left            =   5355
         TabIndex        =   28
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         Top             =   735
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
         Top             =   735
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   3480
         Left            =   210
         ScaleHeight     =   3420
         ScaleWidth      =   10455
         TabIndex        =   24
         Top             =   1680
         Width           =   10515
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   480
            Left            =   0
            TabIndex        =   26
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
               TabIndex        =   27
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
            TabIndex        =   25
            Top             =   -1050
            Width           =   10515
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
               Height          =   2985
               Left            =   -30
               TabIndex        =   34
               Top             =   1530
               Visible         =   0   'False
               Width           =   10515
            End
            Begin VB.Line Line10 
               X1              =   0
               X2              =   10500
               Y1              =   1530
               Y2              =   1530
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
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   8340
         TabIndex        =   75
         ToolTipText     =   "Doble Click  para Cambiar el Punto de Venta"
         Top             =   735
         Width           =   660
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
         Left            =   210
         TabIndex        =   72
         Top             =   1155
         Width           =   1035
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         X1              =   210
         X2              =   10710
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   10720
         X2              =   10720
         Y1              =   1680
         Y2              =   5160
      End
      Begin VB.Label NORVENTA 
         Height          =   225
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   1590
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
         TabIndex        =   29
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   1050
         Width           =   2115
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro:"
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
         TabIndex        =   23
         Top             =   735
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito:"
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
         Left            =   7350
         TabIndex        =   22
         Top             =   1155
         Width           =   1275
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
         TabIndex        =   21
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
         TabIndex        =   15
         Top             =   315
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
      TabIndex        =   20
      Top             =   0
      Width           =   960
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu grabar 
         Caption         =   "Grabar Remito"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu edicion 
      Caption         =   "Edición"
      Begin VB.Menu anularemito 
         Caption         =   "Anulación General de Remito / Orden de Despacho Emitidos"
      End
      Begin VB.Menu marcasincargo 
         Caption         =   "Eliminación Solo de Lista de Pendientes de Factura"
      End
      Begin VB.Menu devoluciondematerialenconsignacion 
         Caption         =   "Devolución Material en Consignación"
      End
      Begin VB.Menu corregirremitos 
         Caption         =   "Corrección Cantidades Despachadas"
      End
      Begin VB.Menu reasignacionderemitosapedidos 
         Caption         =   "Re-Asignación de Remitos a Pedidos"
      End
   End
   Begin VB.Menu consultasylistados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu entregasporarticulo 
         Caption         =   "Entregas por Artículo "
      End
      Begin VB.Menu entregaporcliente 
         Caption         =   "Entregas por Cliente"
      End
      Begin VB.Menu pornumerodecomprobante 
         Caption         =   "Secuencial por Número de Comprobante"
      End
      Begin VB.Menu pendientedefacturar 
         Caption         =   "Pendientes de Facturar "
      End
      Begin VB.Menu codigosequivalentesdeclientes 
         Caption         =   "Códigos Equivalentes de Clientes"
      End
   End
   Begin VB.Menu accesodirecto 
      Caption         =   "Accesos Directos"
      Begin VB.Menu facturacion 
         Caption         =   "Facturación"
      End
      Begin VB.Menu maestrodeclientes 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu FICHACARDEX 
         Caption         =   "Ficha Kardex (Movimientos de Mercadería)"
      End
      Begin VB.Menu importacionderemitos 
         Caption         =   "Importador de Remitos"
      End
   End
   Begin VB.Menu ver 
      Caption         =   "Ver"
      Begin VB.Menu reimprimirremitos 
         Caption         =   "Consulta y Re-Impresión de Remitos"
      End
      Begin VB.Menu notasdeclientes 
         Caption         =   "Notas de Clientes"
      End
   End
   Begin VB.Menu utiles 
      Caption         =   "Útiles"
      Begin VB.Menu generacionbarras 
         Caption         =   "Generación por Código de Barras"
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu depurarremitosporrangodefecha 
         Caption         =   "Depurar Remitos Por Rango de Fechas"
      End
      Begin VB.Menu controldeintegridaddebasededatos 
         Caption         =   "Control de Integridad de la B.Datos"
      End
   End
   Begin VB.Menu configuracion 
      Caption         =   "Configuración"
      Begin VB.Menu configremito 
         Caption         =   "Opciones de Configuracion de Remito"
      End
   End
End
Attribute VB_Name = "FOREMIT04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TOTITFO& 'CANTIDAD DE ITEMS DEL FORMULARIO DE IMPRESIÓN
Dim REPIREMI%

Private Sub anularemito_Click()
    Call OPREMSTA04.Command25_Click
End Sub

'
Private Sub BOTNEXT_Click()
    '\\\OT-05-420-WAR-26/06/05///
    If TRIM$(Label10) = "" Then
      Call MENSERR(24, "Numero de Punto de Venta no Valido")
      BOTNEXT.Enabled = True
      Exit Sub
    End If
    '\\\OT-05-420-WAR-FIN///
    NCXX% = Val(Text1)
    If RASOCLI(NCXX%) <> "" Then
       MODOPRE% = 1
       If Option2.Value = False Then
          RESP1% = COMALTER%("Formato de Pre-Carga:\\Abierto\Por Pedidos Pendientes")
          MODOPRE% = 2: If RESP1% = 1 Then MODOPRE% = 1
       End If
       Call SIGUEREMI1(MODOPRE%)
      Else
        Call MENSERR(24, "Numero de Cuenta no Valido")
    End If
    '
    If ULTREG&("!BOLREDET") < 1 Then
        List1.Clear
        Unload BCODE04 '\\\OT-06-0182-WAR-29/05/06///
        List1.Visible = False
        Command5.Enabled = False
        BOTNEXT.Enabled = True
        Text1.SetFocus
        Exit Sub
    End If
    List1.Clear
    List1.Visible = True
    V& = 0
    '
    If Frame15.Visible = True Then
      'TOTITEMS% = 0 '*
      For U& = 1 To ULTREG&("!BOLREDET")
        x$ = REGLEIDO$("!BOLREDET", U&)
        RELI$ = Space$(80)
        TTXX$ = Mid$(x$, 5, 64)
        Call REPLA(RELI$, TTXX$, 1, 52)
        CANS$ = CSTRING$(Mid$(x$, 89, 8), 3, 10, 1, 2)
        Call REPLA(RELI$, CANS$, 75, 10)
        List1.AddItem RELI$
        'TOTITEMS% = TOTITEMS% + 1 '*
      Next U&
      'Label18 = TRIM$(Str$(TOTITEMS%)) '*
      GoTo 20
    End If
    '
    'TOTITEMS% = 0 '*
    NUOCOMPRA$ = ""
    For U& = 1 To ULTREG&("!BOLREDET")
        x$ = REGLEIDO$("!BOLREDET", U&)
        MEDIDAS$ = Mid$(x$, 193, 12)
        RELI$ = Space$(80)
        CI% = CVI(Mid$(x$, 83, 2))
        If TRIM$(CODEXT$(CI%)) <> "" Then
            V& = V& + 1
            Call GRAREG("!BOLREDET", x$, V&)
            Call REPLA(RELI$, CODEXT$(CI%), 1, 15)
            Call REPLA(RELI$, DESCRIT0$(CI%), 16, 30)
            TXMD$ = TEXTOMEDI$(MEDIDAS$)
            If TXMD$ <> "" Then
               Call REPLA(RELI$, "P.C. " + TXMD$, 16, 30)
            End If
            Call REPLA(RELI$, UNIMED$(CI%), 49, 4)
            CANS$ = CSTRING$(Mid$(x$, 89, 8), 3, 10, 1, 2)
            Call REPLA(RELI$, CANS$, 53, 10)
            Call REPLA(RELI$, Mid$(x$, 107, 12), 66, 12)
            NOPE$ = CSTRING$(Mid$(x$, 151, 4), 3, 7, 0, 2)
            Call REPLA(RELI$, NOPE$, 78, 7)
            NPEDID& = CVS(Mid$(x$, 151, 4))
            NUOCO$ = TRIM$(NUOCOCLI$(NPEDID&, CI%))
            If NUOCO$ <> "" Then
              If InStr(NUOCOMPRA$, " " + NUOCO$ + ",") < 1 Then
                NUOCOMPRA$ = NUOCOMPRA$ + " " + NUOCO$ + ","
              End If
            End If
            List1.AddItem RELI$
            'TOTITEMS% = TOTITEMS% + 1 '*
        End If
    Next U&
    'Label18 = TRIM$(Str$(TOTITEMS%)) '*
    NUOCOMPRA$ = TRIM$(NUOCOMPRA$)
    If Right$(NUOCOMPRA$, 1) = "," Then
      NUOCOMPRA$ = Left$(NUOCOMPRA$, Len(NUOCOMPRA$) - 1)
    End If
    Text20 = NUOCOMPRA$
    Call SETULTREG("!BOLREDET", V&)
    '
20  Command5.Enabled = True
    Command12.Enabled = True
    Text15.SetFocus
    '
End Sub

Private Sub codigosequivalentesdeclientes_Click()
    Call OPREMSTA04.Command19_Click
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
        Label26 = CODICLI$(NC%)
        Call CARTRANSPOR
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
    Call CONECRUN("ARCHIG04/AMACLI", "Archivos Maestros")
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    Dim CIRF%(512), CARF(512)
    V& = 0
    For U& = 1 To ULTREG&("!BOLREDET")
        x$ = REGLEIDO$("!BOLREDET", U&)
        CI% = CVI(Mid$(x$, 83, 2))
        COBA$ = CODEXT$(CI%)
        CABA = CVD(Mid$(x$, 89, 8))
        If TRIM$(COBA$) <> "" Then
           Call LEEEXPLO(COBA$, CABA)
           '
           If CAIT% > 0 Then
              Call REPLA(x$, MKD$(CDbl(CABA)), 163, 8)
              Call REPLA(x$, String$(8, 0), 89, 8)
              Call GRAREG("!BOLREDET", x$, U&)
           End If
           '
           For J% = 1 To CAIT%
             CIY% = CIJ%(J%)
             CANY = CABA * USOI(J%)
             For W% = 1 To V&
               If CIRF%(W%) = CIY% Then
                  CARF(W%) = CARF(W%) + CANY
                  GoTo 19
               End If
             Next W%
             V& = V& + 1
             CIRF%(V&) = CIY%
             CARF(V&) = CANY
19         Next J%
        End If
    Next U&
    '
    For J% = 1 To V&
       x$ = REGBLAN$("BOLREDET")
       Call REPLA(x$, MKI$(CIRF%(J%)), 83, 2)
       Call REPLA(x$, MKD$(CDbl(CARF(J%))), 89, 8)
       Call REGAPP("!BOLREDET", x$)
    Next J%
    Call CIERRARCH("!BOLREDET")
    Command12.Enabled = False
    '
    List1.Clear
    V& = 0
    For U& = 1 To ULTREG&("!BOLREDET")
        x$ = REGLEIDO$("!BOLREDET", U&)
        RELI$ = Space$(80)
        CI% = CVI(Mid$(x$, 83, 2))
        If TRIM$(CODEXT$(CI%)) <> "" Then
            V& = V& + 1
            Call GRAREG("!BOLREDET", x$, V&)
            Call REPLA(RELI$, CODEXT$(CI%), 1, 15)
            Call REPLA(RELI$, DESCRIT0$(CI%), 16, 30)
            Call REPLA(RELI$, UNIMED$(CI%), 49, 4)
            CANS$ = CSTRING$(Mid$(x$, 89, 8), 3, 10, 1, 2)
            Call REPLA(RELI$, CANS$, 53, 10)
            Call REPLA(RELI$, Mid$(x$, 107, 12), 66, 12)
            NOPE$ = CSTRING$(Mid$(x$, 151, 4), 3, 7, 0, 2)
            Call REPLA(RELI$, NOPE$, 78, 7)
            List1.AddItem RELI$
        End If
    Next U&
    Call SETULTREG("!BOLREDET", V&)
    Call CIERRARCH("!BOLREDET")
    '
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    '\\\OT-05-0299-WAR-18/05/05///
    If DERACCE%("SETREMIT", "Configuración - Setup de Remitos.") >= 2 Then
      OPREMIT04.Show 1
    End If
    '\\\OT-05-0299-WAR-FIN///
    Command13.Enabled = True
End Sub

Private Sub Command14_Click()
  '
  NCXX% = Val(Text1)
  If RASOCLI$(NCXX%) <> "" Then
       BCODE04.Show 1
       If BCODE04.List1.ListCount > 0 Then
         '1.- DESHABILITO EL BOTON NEXT
         BOTNEXT.Enabled = False
         '2.- CARGAR BOLREDET
         JJ& = 0
         REBLA$ = REGBLAN$("!BOLREDET")
         For JX% = 0 To BCODE04.List1.ListCount - 1
           z$ = BCODE04.List1.List(JX%)
           CI1$ = Mid$(z$, 1, 17)
           CI2% = CODINT%(CI1$)
           Q1# = XVALO(Mid$(z$, 39, 6))
           y$ = REBLA$
           Call REPLA(y$, MKI$(CI2%), 83, 2)
           Call REPLA(y$, Left$(UNIMED$(CI2%), 4), 85, 4)
           Call REPLA(y$, MKD$(Q1#), 89, 8)
           JJ& = JJ& + 1
           Call GRAREG("!BOLREDET", y$, JJ&)
         Next JX%
         Call SETULTREG("!BOLREDET", JJ&)
         '3.- LLAMAR A SIGUEREMI2 DE FOREMIT04
         If JJ& > 0 Then Command12.Enabled = True
         Call List1_DblClick   'SIGUEREMI2
       ElseIf BCODE04.List1.ListCount <= 0 Then
         Call SETULTREG("!BOLREDET", 0)
         List1.Clear
       End If
       Exit Sub
     Else
       Call COMUNI("Debe Seleccionar un Cliente")
       Call Command1_Click
  End If
  '
End Sub

Private Sub Command15_Click()
    '
    Command15.Enabled = False
    LIBREMIT04.Show 1
    Command15.Enabled = True
    '
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
        Text1.SetFocus
       Else
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Call FINAL("")
    End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NC% = Val(TRIM$(Text1.TEXT))
    If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
         Call NOTACLI(NC%, 2)
       Else
         Call MENSERR(24, "Ingrese el Número de Cliente.")
    End If
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
AA1 = ULTREG&("!BOLREDET")
    '
    ' TRANSFIERE LAS OBSERVACIONES
    For U& = 1 To 4
       TTXXX$ = TRIM$(Text16(U& - 1).TEXT)
       Call GRAREG("!OBSERVOF", TTXXX$, U&)
    Next U&
    Call SETULTREG("!OBSERVOF", 4)
    Call CIERRARCH("!OBSERVOF")
    '
    ' VALIDA LOS DATOS INGRESADOS
    'If VALIDATOS% < 1 Then Exit Sub
    '
    '\\\OT-05-0323-EPB-24/05/05///
    Call BLOQARCH("MOVISTO")
    NPRESENTADO& = Val(Mid(Text11, 4))
    'OT-05-0419-WAR-26/06/05
    NUPUVE$ = TRIM(Label10)
    NROPRO& = NPRESENTADO&
    If REPIREMI% > 0 Then GoTo 10
    '
    Call NUREMIFA(NROPRO&, NUPUVE$)
    'Call NUREMIFA(NROPRO&)
    'FIN OT
    If NPRESENTADO& <> NROPRO& Then
        Call MENSERR(24, "Número de Remito ya Utilizado.\FLEXOFT Asignara el Próximo Número libre, vuelva a pedir la Impresión.")
        CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
        If CMO$ = "" Then CMO$ = "RT"
        Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
        Command5.Enabled = True
        Exit Sub
    End If
    '
10  Call PRIREMIFAC
    Call LIBARCH("MOVISTO")
    '\\\OT-05-0323-EPB-FIN///
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    DP1% = Val(CONTROL$("", "DEPOEXPE"))
    Text10.TEXT = TRIM$(Str$(DP1%))
    Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    NROPRO& = 0
    'OT-05-0419-WAR-26/06/05///
    NUPUVE$ = TRIM(Label10)
    Call NUREMIFA(NROPRO&, NUPUVE$)
    'Call NUREMIFA(NROPRO&)
    'FIN OT
    Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '
    Call CIERRARCH("!BOLREDET")
    Call BLANARCH("!BOLREDET")
    List1.Clear
    List1.Visible = False
    Command5.Enabled = False
    Command12.Enabled = False
    BOTNEXT.Enabled = True
    Frame15.Visible = False
    Text15 = ""
    Text17 = ""
    '
    '\\\OT-06-0182-WAR-29-05-06///
    '1.- Descargo el Formulario
    Unload BCODE04
    '\\\OT-06-0182-WAR-FIN///
    '
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
    OPREMSTA04.Show 1
    Command6.Enabled = True
    '
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call CONECRUN("VENGES04/FACTURA", "Gestión de Ventas y Deudores")
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    FIKARDEX.Show 1
'    Call GENKARDEX("")
'    AppActivate Caption
'    SendKeys Chr$(1)
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

Private Sub configremito_Click()
    Call Command13_Click
End Sub

Private Sub controldeintegridaddebasededatos_Click()
    Call OPREMSTA04.Command13_Click
End Sub

Private Sub corregirremitos_Click()
    Call OPREMSTA04.Command10_Click
End Sub

Private Sub depurarremitosporrangodefecha_Click()
    Call OPREMSTA04.Command15_Click
End Sub

Private Sub devoluciondematerialenconsignacion_Click()
   Call OPREMSTA04.Command9_Click
End Sub

Private Sub entregaporcliente_Click()
    Call OPREMSTA04.Command3_Click
End Sub

Private Sub entregasporarticulo_Click()
   Call OPREMSTA04.Command4_Click
End Sub

Private Sub facturacion_Click()
   Call Command7_Click
End Sub


Private Sub FICHACARDEX_Click()
      Call Command8_Click

End Sub

Private Sub Form_Load()
    '
    
    Call CENTRAFORM(Me)
    Me.Visible = False
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
      Me.Frame10.BackColor = &HFCD7B6
      Me.Frame11.BackColor = &HFCD7B6
      Me.Frame12.BackColor = &HFCD7B6
      Me.Frame13.BackColor = &HFCD7B6
      Me.Frame14.BackColor = &HFCD7B6
      Me.Frame15.BackColor = &HFCD7B6
      Me.Frame16.BackColor = &HFCD7B6
      Me.Frame17.BackColor = &HFCD7B6
      
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      BOTNEXT.Top = 1760
      
      Set BOTNEXT.Container = Me.Picture16
      BOTNEXT.ZOrder 0
      Command5.Top = 7500
      
      Set Command5.Container = Me.Picture16
      Command5.ZOrder 0
      Me.Refresh
      ' &H8000000F
    End If

    TEXPE$ = "Remito de Ventas"
    If CONTROL$("", "MODOEXPE") = "ODESPACH" Then
        TEXPE$ = "Orden de Despacho"
    End If
    Caption = Caption + " - " + TEXPE$
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
10  If ULTREG("ECODEP") < 1 Then
        Call MENSERR(24, "Falta Definir Tabla de Depósitos")
        Call CARTADEPOSI
        GoTo 10
    End If
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Remito.")
        OPREMIT04.Show 1
        GoTo 10
    End If
    '\\\OT-06-0001-WAR-02/01/06///
    CODFOR$ = TRIM$(UCase$(FORIPRE$))
    TOTITFO& = Val(ATRIFORM$(CODFOR$, "MAX-ITEM"))
    '\\\OT-06-0001-WAR-FIN///
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito de Facturación:\  \'" + RMCC$ + "'.")
            OPREMIT04.Show 1
            GoTo 10
        End If
    End If
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    DP1% = Val(CONTROL$("", "DEPOEXPE"))
    If DP1% < 1 Or ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
       Call COMUNI("Error de Configuración:\  \Falta Definir Depósito de Expedición.")
       OPREMIT04.Show 1
       GoTo 10
    End If
    '
    Text10.TEXT = TRIM$(Str$(DP1%))
    Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    '
    Label10 = ""
    PREFIREM$ = TRIM$(CONTROL$("", "PREFIREM"))
    If PREFIREM$ <> "*" Then
      Label10 = PREFIREM$
    End If
    '
    NROPRO& = 0
    Call NUREMIFA(NROPRO&)
    Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '\\\OT-05-420-WAR-26/06/05///
    '\\\OT-05-420-WAR-FIN///
    Timer1.Enabled = True
    Refresh
    '
    REPIREMI% = 0
    Call COPYSTRU("BOLREDET", "GUARREMI")
    If ULTREG&("!GUARREMI") > 0 Then
30      TTXX$ = "El Último Remito Generado\Se ha Grabó o Imprimió en Forma Defectuosa."
        TTXX$ = TTXX$ + "\   \Desea Recuperar y Re-Imprimir ese Remito ?"
        OPX% = COMALTER%(TTXX$ + "\\Si, Recuperar\No, Cancelar")
        If OPX% = 1 Then
             Call RECUREMI
           ElseIf OPX% = 2 Then
             Call BORRAGUARDA
           Else
             GoTo 30
        End If
    End If
    '
    Me.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Cancel = 0
   Screen.MousePointer = 1
   Call FINAL("")
End Sub

Private Sub generacionbarras_Click()
   Call Command14_Click
End Sub

Private Sub grabar_Click()
  Call Command5_Click
End Sub

Private Sub importacionderemitos_Click()
   Call OPREMSTA04.Command21_Click
End Sub

'\\\OT-05-0419-WAR-26/06/05///
Private Sub Label10_dblClick()
  TALO$ = TRIM$(Label10)
'  Select Case TALO$
'    Case ""
'        Label10 = "0001"
'    Case "0001"
'        Label10 = "0002"
'    Case "0002"
'        Label10 = "0001"
'  End Select
  Select Case TALO$
    Case ""
        Label10 = "0001"
    Case "0001"
        Label10 = "0002"
    Case "0002"
        Label10 = "0003"
    Case "0003"
        Label10 = "0004"
    Case "0004"
        Label10 = "0001"
  End Select
  NUPUVE$ = TRIM(Label10)
  CMO$ = "RT"
  NROPRO& = 0
  Call NUREMIFA(NROPRO&, NUPUVE$)
  Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
End Sub

Private Sub Label18_Change()
   '\\\OT-06-0001-WAR-02/01/06///
   TOTITEMS% = Val(Label18)
   If TOTITEMS% > TOTITFO& Then
     Call MENSERR(24, "Renglones de Este Remito (" & TRIM$(Str$(TOTITEMS%)) & ") Superan\Máximo Permitido (" & TRIM$(Str$(TOTITFO&)) & ") del Formulario\de Impresión de Remito")
   End If
   '
End Sub

'\\\OT-05-0419-WAR-FIN///
Private Sub Label2_DblClick()
'    OPREMIFAC.Show 1
End Sub

Private Sub Label26_DblClick()
   Call SELECHO("INDICLIE")
   CPROX$ = TRIM$(VALACT1$("INDICLIE"))
   If CPROX$ <> "" Then
     NCXI% = NUMECLI%(CPROX$)
     If NCXI% > 0 Then
       If REGICLI&(NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
         Text2.TEXT = RASOCLI$(NC%)
         Text3.TEXT = DOMICLI$(NC%)
         Text4.TEXT = CPOSCLI$(NC%)
         Text5.TEXT = LOCACLI$(NC%)
         Call CARTRANSPOR
       End If
     End If
   End If
End Sub

Private Sub Label9_DblClick()
   Frame15.Visible = True
   List1.Visible = True
End Sub

Private Sub List1_DblClick()
    '
    If Command12.Enabled = False Then
       Call MENSERR(24, "Imposible Re-Editar Despacho.\Cancele y Vuelva a Cargar.")
       Exit Sub
    End If
    '
    Call SIGUEREMI2
    If ULTREG&("!BOLREDET") < 1 Then
        List1.Clear
        Unload BCODE04 '\\\OT-06-0182-WAR-29/05/06///
        List1.Visible = False
        Command5.Enabled = False
        Command12.Enabled = False
        BOTNEXT.Enabled = True
        Text1.SetFocus
        Label18 = "" '*
        Exit Sub
    End If
    List1.Clear
    List1.Visible = True
    V& = 0
    'TOTITEMS% = 0 '*
    NUOCOMPRA$ = ""
    For U& = 1 To ULTREG&("!BOLREDET")
        x$ = REGLEIDO$("!BOLREDET", U&)
        RELI$ = Space$(80)
        CI% = CVI(Mid$(x$, 83, 2))
        If TRIM$(CODEXT$(CI%)) <> "" Then
            V& = V& + 1
            Call GRAREG("!BOLREDET", x$, V&)
            Call REPLA(RELI$, CODEXT$(CI%), 1, 15)
            Call REPLA(RELI$, DESCRIT0$(CI%), 16, 30)
            Call REPLA(RELI$, UNIMED$(CI%), 49, 4)
            CANS$ = CSTRING$(Mid$(x$, 89, 8), 3, 10, 1, 2)
            Call REPLA(RELI$, CANS$, 53, 10)
            Call REPLA(RELI$, Mid$(x$, 107, 12), 66, 12)
            NOPE$ = CSTRING$(Mid$(x$, 151, 4), 3, 7, 0, 2)
            Call REPLA(RELI$, NOPE$, 78, 7)
            NPEDID& = CVS(Mid$(x$, 151, 4))
            NUOCO$ = TRIM$(NUOCOCLI$(NPEDID&, CI%))
            If NUOCO$ <> "" Then
              If InStr(NUOCOMPRA$, " " + NUOCO$ + ",") < 1 Then
                NUOCOMPRA$ = NUOCOMPRA$ + " " + NUOCO$ + ","
              End If
            End If
            List1.AddItem RELI$
            'TOTITEMS% = TOTITEMS% + 1 '*
        End If
    Next U&
    'Label18 = TRIM(Str$(TOTITEMS%))
    NUOCOMPRA$ = TRIM$(NUOCOMPRA$)
    If Right$(NUOCOMPRA$, 1) = "," Then
      NUOCOMPRA$ = Left$(NUOCOMPRA$, Len(NUOCOMPRA$) - 1)
    End If
    Text20 = NUOCOMPRA$
    Call SETULTREG("!BOLREDET", V&)
    Screen.MousePointer = 1
    '
    Command5.Enabled = True
    Text15.SetFocus
    '
End Sub

Private Sub maestrodeclientes_Click()
   Call Command11_Click
End Sub

Private Sub marcasincargo_Click()
   Call OPREMSTA04.Command6_Click
End Sub

Private Sub notasdeclientes_Click()
   Call Command3_Click
End Sub

Private Sub pendientedefacturar_Click()
   Call OPREMSTA04.Command8_Click
End Sub

Private Sub pornumerodecomprobante_Click()
     Call OPREMSTA04.Command5_Click
End Sub

Private Sub reasignacionderemitosapedidos_Click()
    Call OPREMSTA04.Command12_Click
End Sub

Private Sub reimprimirremitos_Click()
   Call OPREMSTA04.Command14_Click
End Sub

Private Sub salir_Click()
   Call Command2_Click
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
        Label26 = CODICLI$(NC%)
        Call CARTRANSPOR
        BOTNEXT.SetFocus
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
                Label26 = CODICLI$(NC%)
                List1.Clear
                Unload BCODE04 '\\\OT-06-0182-WAR-29/05/06///
                Call CARTRANSPOR
                BOTNEXT.Enabled = True
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
    If CMO$ = "" Then
       CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
       If CMO$ = "" Then CMO$ = "RT"
    End If
    '
    VDEFF$ = TRIM$(Text11.TEXT)
    Text11.TEXT = VDEFF$
    If Left$(VDEFF$, 3) <> CMO$ + "." Then GoTo 90
    NRONUE& = Val(Mid$(VDEFF$, 4))
    NROPRO& = NRONUE&
    '\\\OT-05-0419-WAR-26/06/05///
    'Call NUREMIFA(NROPRO&)
    NUPUVE$ = TRIM(Label10)
    Call NUREMIFA(NROPRO&, NUPUVE$)
    'FIN
    If NROPRO& <> NRONUE& Then
        Call MENSERR(24, "Número de Remito ya Utilizado.\FLEXOFT Asignara el Próximo Número libre.")
        Beep
        GoTo 95
    End If
    Exit Sub
    '
90  Beep
    NROPRO& = 0
    'OT-05-0419-WAR-26/06/05///
    NUPUVE$ = TRIM(Label10)
    Call NUREMIFA(NROPRO&, NUPUVE$)
    'Call NUREMIFA(NROPRO&)
    'FIN OT
95  Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    Text11.SetFocus
    '
End Sub

Private Sub Text2_gotfocus()
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
    Call Command13_Click
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

Sub CARTRANSPOR()
    '
    NCLIE% = Val(Text1.TEXT)
    LENTREG$ = Space$(96)
    RFEX$ = RECFILT$("LENTRECL", 1, MKI$(NC%))
    If Len(RFEX$) >= 4 Then
       RELENT& = CVS(Left$(RFEX$, 4))
       If RELENT& > 0 Then
         If RELENT& <= ULTREG&("LENTRECL") Then
           LENTREG$ = Mid$(REGLEIDO$("LENTRECL", RELENT&), 3)
         End If
       End If
    End If
    Text8.TEXT = TRIM$(Mid$(LENTREG$, 65, 32))
    '
End Sub

Sub SIGUEREMI1(MODOPRE%)
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
    Call NOTACLI(NC%, 1)
    '
    Call CIERRARCH("*.*")
    Call BLANARCH("!BOLREDET")
    '
    If MODOPRE% = 1 Then GoTo 10
    If Frame15.Visible = True Then GoTo 10
    '\\\OT-05-0240-WAR-02/05/05///
    Call DESHASFECHA(DES%, HAS%, PERIO$, 1) ' SE MODIFICA PARA PODER
    If PERIO$ = "" Then Exit Sub            ' INGRESAR UN DESDE Y HASTA
    
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
    
    '\\\OT-05-0240-WAR-FIN///
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE NroClie = " & NC% & " "
    SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
    SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
    '\\\OT-05-0240-WAR-02/05/05///
    SQLX$ = SQLX$ + " AND CDBL(FeSolEnt) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND CDBL(FeSolEnt) <= " & HAS1# & " "
    
    '\\\OT-05-0240-WAR-FIN///
    SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
    SQLX$ = SQLX$ + " ORDER BY NroPed ASC, FeSolEnt ASC, NuOrItem ASC " 'ESTA LINEA LA AGREGUE
    '
    Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      '
      JJ& = 0
      Do While (.EOF = False)
        CIPD% = !CodiInt
        If CIPD% > 0 And CIPD% <= NUMAS% Then
          CAPEN = !CanPed - !CantEnt
            If CAPEN < 0 Then CAPEN = 0
          If !TipoPed = "ABIERTO" Then CAPEN = 999999
          FEMIPED% = CVINT(!FechEmis)
          FENTRESOL% = CVINT(!FESOLENT)
          PREUNI = !PreUnid * (100 - PORDESCLI(NC%)) / 100
          MONEMA% = !MoneEmis
          '
          y$ = REGBLAN$("BOLREDET")
          Call REPLA(y$, MKI$(FEMIPED%), 81, 2)
          Call REPLA(y$, MKI$(CIPD%), 83, 2)
          Call REPLA(y$, UNIMED$(CIPD%), 85, 4)
          Call REPLA(y$, MKD$(0), 89, 8)
          Call REPLA(y$, MKS$(!NroPed), 151, 4)
          Call REPLA(y$, MKI$(FENTRESOL%), 155, 2)
          Call REPLA(y$, MKI$(MONEMA%), 157, 2)
          Call REPLA(y$, MKS$(PREUNI), 159, 4)
          Call REPLA(y$, MKD$(CDbl(CAPEN)), 163, 8)
            U& = !NuOrItem
          Call REPLA(y$, MKS$(U&), 189, 4)
          JJ& = JJ& + 1
          Call GRAREG("!BOLREDET", y$, JJ&)
          '
        End If
      .MoveNext
      Loop
      On Error GoTo 0
    End With
    '
10  BOTNEXT.Enabled = False
    Call SIGUEREMI2
    '
End Sub
'

Sub SIGUEREMI2()
    '
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Remito de Facturación - Detalle de Material a Despachar"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
    '
    If Frame15.Visible = True Then
      Call BLANARCH("!BOLREDET")
      VTB% = VENTABU%("REMILIB" + ATRI$)
      GoTo 11
    End If
    '
    ARCHIPRE$ = "MASTER"
    LPREX% = Val(CONTROL$("", "LIPREXPE"))
    If LPREX% > 0 Then
      If LPREX% < 256 Then
        If ECOARCH$("LISTAS", Chr$(LPREX%)) <> "" Then
          ARCHIPRE$ = TRIM$(Str$(LPREX%))
          While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
          ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        End If
      End If
    End If
    '
    VERIST% = (-1)
    If CONTROL$("REMSTAN", "VERISTOK") = "VERISTOK" Then
       VERIST% = 1
    End If
    '
    ATRI$ = ATRI$ + "/ARCHECO(1)=" + ARCHIPRE$
10  VTB% = VENTABU%("DETAREMI" + ATRI$)
11  If VTB% < 0 Then
       If COMALTER%("Realmente Desea Cancelar Este Remito ?\\No, Continuar\Si, Cancelar") <> 2 Then GoTo 10
       Call BLANARCH("!BOLREDET")
       Frame15.Visible = False
       Exit Sub
    End If
    '
    NCLIE% = Val(Text1)
    ATXX% = 0
    For KU& = 1 To ULTREG&("!BOLREDET")
      y$ = REGLEIDO$("!BOLREDET", KU&)
      CIPD% = CVI(Mid$(y$, 83, 2))
      If CIPD% > 0 Then
        NUPED& = CVS(Mid$(y$, 151, 4))
        If NUPED& < 1 Then
          If Option2.Value = False Then
            ENCUPE% = 0
            Call BUSCAPEDAC(CIPD%, NCLIE%, NUPED&, FEMIPE%, FENTSOL%, PREUNI, CAPEN, NORITU&, ENCUPE%)
            '      FEMIPED% = CVINT(!FechEmis)
            '  FENTRESOL% = CVINT(!FeSolEnt)
            '  PREUNI = !PREUNID * (100 - PORDESCLI(NC%)) / 100
              '
            If ENCUPE% > 0 Then
              ATXX% = 1
              Call REPLA(y$, MKI$(FEMIPE%), 81, 2)
              Call REPLA(y$, UNIMED$(CIPD%), 85, 4)
              Call REPLA(y$, MKS$(NUPED&), 151, 4)
              Call REPLA(y$, MKI$(FENTSOL%), 155, 2)
              Call REPLA(y$, MKS$(PREUNI), 159, 4)
              Call REPLA(y$, MKD$(CDbl(CAPEN)), 163, 8)
              Call REPLA(y$, MKS$(NORITU&), 189, 4)
              Call GRAREG("!BOLREDET", y$, KU&)
            End If
          End If
        End If
      End If
    Next KU&
    '
    If ATXX% > 0 Then GoTo 10
    '
    JJ& = 0
    For KU& = 1 To ULTREG&("!BOLREDET")
      y$ = REGLEIDO$("!BOLREDET", KU&)
      CIPD% = CVI(Mid$(y$, 83, 2))
      If CIPD% > 0 Then
        CANENT = CVD(Mid$(y$, 89, 8))
        If CANENT > 0.05 Then
          JJ& = JJ& + 1
          Call GRAREG("!BOLREDET", y$, JJ&)
        End If
      End If
    Next KU&
    Call SETULTREG("!BOLREDET", JJ&)
    If JJ& < 1 Then
      Frame15.Visible = False
      Exit Sub
    End If
    '
    If CONTROL$("REMSTAN", "NUSERIE") = "SI" Then
      VTB% = VENTABU%("DETAREMS" + ATRI$)
      If VTB% < 1 Then
        Frame15.Visible = False
        Exit Sub
      End If
      '
    End If
End Sub
'
Sub BUSCAPEDAC(CIPD%, NCLIE%, NUPED&, FEMIPE%, FENTSOL%, PREUNI, CAPEN, NORITU&, ENCUPE%)
    '
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE NroClie = " & NCLIE% & " "
    SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
    SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
    SQLX$ = SQLX$ + "AND CodiInt = " & CIPD% & " "
    SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
    SQLX$ = SQLX$ + "ORDER BY NroPed " 'ESTA LINEA LA AGREGUE YO 17/02/05
    '
    Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    NUPED& = 0
    FEMIPE% = 0
    FENTSOL% = HOY(HOS$)
    PREUNI = 0
    CAPEN = 0
    NORITU& = 0
    ENCUPE% = 0
    '
    With PEDTEMP
      On Error Resume Next
      .MoveLast
      CAREGIS& = .RecordCount
      '
      If CAREGIS& = 1 Then
          NUPED& = !NroPed
          FEMIPE% = CVINT%(!FechEmis)
          FENTSOL% = CVINT(!FESOLENT)
          PREUNI = !PreUnid
          CAPEN = !CanPed - !CantEnt
            If CAPEN < 0 Then CAPEN = 0
          If !TipoPed = "ABIERTO" Then CAPEN = 999999
          NORITU& = !NuOrItem
          ENCUPE% = 1
        ElseIf CAREGIS& > 1 Then
          .MoveFirst
          JJ& = 0
          Do While (.EOF = False)
            TTXX$ = MKS$(!NroPed)
            TTXX$ = TTXX$ + AJUSTI$(!NroOcom, 16) + "  "
               FEX = CVINT(!FechEmis)
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
          '////***////
          Call FRESHECHO("!SELECPE") 'ESTO SE AGREGO EL 16/02/05 PARA REFRESCAR LA LISTA DE PEDIDOS.
          '////***/////
          '
          TITUCON$ = "Cod: " + TRIM$(CODEXT$(CIPD%)) + " - " + DESCRIT$(CIPD%)
20        TTXX$ = REGSEL$("!SELECPE/" + TITUCON$)
          If Len(TTXX$) <= 4 Then GoTo 20
          '
          NUPED& = CVS(Left$(TTXX$, 4))
          FEMIPE% = FECHANUM(Mid$(TTXX$, 23, 8))
          FENTSOL% = FECHANUM(Mid$(TTXX$, 23, 8))
          PREUNI = Val(Mid$(TTXX$, 45, 12))
          CAPEN = Val(Mid$(TTXX$, 57, 10))
          NORITU& = Val(Mid$(TTXX$, 67, 6))
          ENCUPE% = 1
      End If
    End With
    '

End Sub

Sub PRIREMIFAC()
    '
    Call LIMPIATETAB
    NC% = XVALO(Text1.TEXT)
    DESTIDOC% = NC%


    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    If FORIPRE$ = "" Then
      Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Remito.")
      Exit Sub
    End If
    '
    If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
        RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
        Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito de Facturación:\  \'" + RMCC$ + "'.")
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
    VERIST% = (-1)
    If CONTROL$("REMSTAN", "VERISTOK") = "VERISTOK" Then
      VERIST% = 1
    End If
    '\\\OT-05-0216-WAR-25/04/05///
    REPLAORI% = 0
    If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
       REPLAORI% = 1
    End If
    '\\\OT-05-0216-WAR-FIN///
    '
    DEPORIG% = Val(Text10.TEXT)
    CAITMO% = 0
    For YX& = 1 To ULTREG&("!BOLREDET")
      YY$ = REGLEIDO$("!BOLREDET", YX&)
      CI% = CVI(Mid$(YY$, 83, 2))
      CAN# = CVD(Mid$(YY$, 89, 8))
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          If CAN# > 0.005 Then
            CAITMO% = CAITMO% + 1
            CIMOVIM%(CAITMO%) = CI%
            CAMOVIM(CAITMO) = (-1) * CAN#
          End If
        End If
      End If
    Next YX&
    If REPIREMI% > 0 Then GoTo 20
    '
    If PUEDEMOVI%(DEPORIG%) < 0 Then Exit Sub       ' CANCELA SALIDA
    '
    If Frame15.Visible = True Then GoTo 20
    '
    If VERIST% = 1 Then
      For YX& = 1 To ULTREG&("!BOLREDET")
        YY$ = REGLEIDO$("!BOLREDET", YX&)
        CI% = CVI(Mid$(YY$, 83, 2))
        CAN# = CVD(Mid$(YY$, 89, 8))
        LOTE$ = Mid$(YY$, 107, 12)
        If CI% > 0 Then
          If CI% <= NUMAS% Then
            If CAN# > 0.005 Then
              If STODEPOS(CI%, DEPORIG%) - CAN# < (-0.005) Then
                CONTI% = COMALTER%("ATENCION !!!\Stock Insuficiente en Código '" + TRIM$(CODEXT$(CI%)) + "'\" + DESCRIT0$(CI%) + "\\Cancelar\Continuar")
                If CONTI% <> 2 Then
                  Exit Sub
                End If
              End If
            End If
          End If
        End If
      Next YX&
    End If
    '
20  HOJASUEL% = 0
    If CONTROL$("", "TIPAREMI") = "HOJASUEL" Then HOJASUEL% = 1
    '
    Screen.MousePointer = 11      ' Registrando Remito
    '
    AINDI% = 0
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    '
    NC% = XVALO(Text1.TEXT)
    REMI$ = TRIM$(Text11.TEXT)
    NROO& = XVALO(Mid$(REMI$, 4))
    FF% = FECHANUM(Text9.TEXT)
    DEPO% = Val(Text10.TEXT)
    '\\\OT-05-0419-WAR-26/06/05///
        'PREFIREM$ = CONTROL$("", "PREFIREM")
        'If PREFIREM$ = "" Then PREFIREM$ = "0001"
        PREFIREM$ = TRIM(Label10)
    '\\\OT-05-0419-WAR-FIN///
    NROREMIT$ = TRIM$(Mid$(REMI$, 4))
    While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
    NROREMIT$ = CMO$ + "-X-" + PREFIREM$ + "-" + NROREMIT$
    '
    HO% = HOY(HOS$)
    HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
    User% = USNBR% Mod 100
    CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(User%) + String$(2, 0)
    '
Call GUARDAREMI
    YXI% = 0
    NUORIT% = 0
    Call APERBASDAT("STOCKS")
    Call APERBASDAT("PEDCLI")
    NOCLOSE = 1: NOARRAS% = 1
    '
    For YX& = 1 To ULTREG&("!BOLREDET")
      YY$ = REGLEIDO$("!BOLREDET", YX&)
      CI% = CVI(Mid$(YY$, 83, 2))
      CAN# = CVD(Mid$(YY$, 89, 8))
      LOTE$ = Mid$(YY$, 107, 12)
      NPEDI& = CVS(Mid$(YY$, 151, 4))
      MEDIDAS$ = Mid$(YY$, 193, 12)
      '
      SUCOD$ = SUBCODI$(NPEDI&, CI%)
      Call REPLA(YY$, SUCOD$, 97, 8)
      Call GRAREG("!BOLREDET", YY$, YX&)
      '
      If CI% > 0 Or Frame15.Visible = True Then
        If CI% <= NUMAS% Or Frame15.Visible = True Then
          If CAN# > 0.005 Or Frame15.Visible = True Then
            NUOC& = CVS(Mid$(YY$, 151, 4))
            DOSEC$ = ""
            CANTI = (-1) * CAN#
            VUNIST# = CDbl(CVS(Mid$(YY$, 159, 4)))
            MONEVA% = CVI(Mid$(YY$, 157, 2)):  If MONEVA% < 1 Then MONEVA% = 1
            If MONEVA% > ULTREG&("TAMONED") Then MONEVA% = 1
            '
            DOSELA$ = "P/F - PC." + TRIM$(Str$(NUOC&))
            DOSEC$ = "PC." + TRIM$(Str$(NUOC&))
            '
            SERINU$ = TRIM$(Mid$(YY$, 209, 16)) + TRIM$(Mid$(YY$, 225, 8))
            If SERINU$ <> "" Then
              If TRIM$(Mid$(YY$, 233, 8)) <> "" Then
                SERINU$ = SERINU$ + " - "
              End If
            End If
            SERINU$ = SERINU$ + TRIM$(Mid$(YY$, 233, 8))
            '
            REFEMO$ = RASOCLI$(NC%)
            If SERINU$ <> "" Then
              REFEMO$ = "S.No." + SERINU$
            End If
            '
            NUORIT% = NUORIT% + 1
            DEPX% = DEPO%: If DEPO% = 99 Then DEPX% = DEPOPRE%(CI%)
            '\\\OT-06-0083-WAR-07/03/06///
            REMICOR$ = Mid$(Text11.TEXT, 1, 3)
            NURE1& = Val(Mid$(Text11.TEXT, 4))
            NURE2$ = TRIM$(Str$(NURE1&))
            SUC1% = Val(Label10)
            SUC2$ = TRIM$(Str$(SUC1%))
            REMICOR1$ = REMICOR$ + SUC2$ + "-" + NURE2$
            '\\\OT-06-0083-WAR-FIN///
            Call MOVISTOK(FF%, CI%, LOTE$, CMO$, NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NC%, DEPO%, CANTI)
            CIMOVI%(CI%) = 1
            '\\\OT-06-0105-WAR-27/03/06///
            'MOVIMIENTO COMPENSACION INTERNA PARA CI% KIT
            If DESCOMPO%(CI%) > 0 Then
              Call MOVISTOK(FF%, CI%, LOTE$, "CI", NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NC%, DEPO%, CANTI * (-1))
            'MOVIMIENTO SALIDA PARA COMPONENTES DEL KIT
              Call AREXPLO(CI%, 1, 1)
              NUORIT% = 0
              For KKJ% = 1 To CAIT%
                NUORIT% = NUORIT% + 1
                CICO% = CIJ%(KKJ%)
                CACICO = USOI(KKJ%) * CANTI
                Call MOVISTOK(FF%, CICO%, LOTE$, "CI", NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NC%, DEPO%, CACICO)
                '
              Next KKJ%
            End If
            '\\\OT-06-0105-WAR-FIN///
          End If
        End If
      End If
    Next YX&
    NOCLOSE% = 0: NOARRAS% = 0
    '
    ' blanquear pedidos pendientes
    Call APERBASDAT("PEDCLI")
    NC% = Val(Text1.TEXT)
    For YX& = 1 To ULTREG&("!BOLREDET")
      YY$ = REGLEIDO$("!BOLREDET", YX&)
      CI% = CVI(Mid$(YY$, 83, 2))
      CAN# = CVD(Mid$(YY$, 89, 8))
      NUPEDIORI& = CVS(Mid$(YY$, 151, 4))
      NUORITPE% = CVS(Mid$(YY$, 189, 4))
      '
      Call BLANPEDIPE(NC%, FF%, NROREMIT$, CI%, CAN#, NUPEDIORI&, NUORITPE%)
      '
    Next YX&
    '
    PedDeta.Close
    Movisto.Close
    '
5   Call BAJALDISCO
    '
    GoSub PRIDOCREMI
    Call VESARRGLO
    '
Call BORRAGUARDA
'
GoTo 999
    '
    '
    '
PRIDOCREMI:
    '
    PRINTERPORT$ = CONTROL$(IDENTER$, "PORTREM")
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    If TRIM$(FORIPRE$) <> "" Then
      '
      'PREFIREM$ = CONTROL$("", "PREFIREM")
      PREFIREM$ = TRIM(Label10)
      If PREFIREM$ = "" Then PREFIREM$ = "0001"
      NROREMIT$ = TRIM$(Mid$(REMI$, 4))
      While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
      NROREMIT$ = PREFIREM$ + "-" + NROREMIT$
      '
      FEX = FF%
      FECHDOC$ = FECHATEX$(FEX)
      NUMEDOC$ = NROREMIT$
      TIPODOC$ = "Remito"
      If Option2.Value = True Then
        TIPODOC$ = "Remito (en Consignación)"
      End If
      '
      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Text8.TEXT)
      CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(Text15.TEXT)
      CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Text12.TEXT)
      CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(Text17.TEXT)
      CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = TRIM$(Text14.TEXT)
      CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Text13.TEXT)
      CODPARAM$(7) = "DEPOSAL ": DOCPARAM$(7) = TRIM$(DEPOSAL$)
      CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NROREMIT$)
      CODPARAM$(9) = "NRO-OCOM": DOCPARAM$(9) = TRIM$(Text20.TEXT)
      CODPARAM$(10) = "PATE-TRA": DOCPARAM$(10) = TRIM$(Text18.TEXT)
      CODPARAM$(11) = "N-AUTENT": DOCPARAM$(11) = TRIM$(NORVENTA)
      CODPARAM$(12) = "COND-TRA": DOCPARAM$(12) = TRIM$(Text19)
      CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = TRIM$(Text21)
      CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = TRIM$(Text22)
      '\\\OT-07-0007-RG-08/01/07///
      HORA$ = HORAENT$(NC%)
      CODPARAM$(15) = "Z-PLANO": DOCPARAM$(15) = TRIM$(HORA$)
      '\\\OT-07-0007-RG-FIN///
      CAPARDOC% = 15
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "CANTIDAD"
      CODTABU1$(4) = "NRO-LOTE"
      CODTABU1$(5) = "UNIMED"
      CODTABU1$(6) = "N-OCOMPR"
      CODTABU1$(7) = "ORD-ITEM"
      CODTABU1$(8) = "REF-MAT1"
      CODTABU1$(9) = "REF-MAT2"
      CACOLTAB1% = 9
      '
'      DESTIDOC% = NC%
      CAITAB1% = 0
      '
      For YX& = 1 To ULTREG&("!BOLREDET")
        YY$ = REGLEIDO$("!BOLREDET", YX&)
        CI% = CVI(Mid$(YY$, 83, 2))
        CAN# = CVD(Mid$(YY$, 89, 8))
        CAC# = 0
        If CI% > 0 Or Frame15.Visible = True Then
          If CI% <= NUMAS% Or Frame15.Visible = True Then
            If CAN# + CAC# > 0.005 Or Frame15.Visible = True Then
              '
              UNID$ = UNIMED$(CI%)
              If UNID$ = "" Then UNID$ = UNIMED$(CI%)
              CANS$ = CSTRING$(MKD$(CAN#), 3, 12, 1, 2)
              If CAC# >= 0.05 Then
                CANS$ = CSTRING$(MKD$(CAC#), 3, 12, 1, 2)
                Call REPLA(CANS$, "(*)", 1, 3)
              End If
              NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
              NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
              '
              SERINU$ = TRIM$(Mid$(YY$, 209, 16)) + TRIM$(Mid$(YY$, 225, 8))
              If SERINU$ <> "" Then
                If TRIM$(Mid$(YY$, 233, 8)) <> "" Then
                  SERINU$ = SERINU$ + " - "
                End If
              End If
              SERINU$ = SERINU$ + TRIM$(Mid$(YY$, 233, 8))
              '
              CAITAB1% = CAITAB1% + 1
              CODDX$ = CODEXT$(CI%) + TRIM$(Mid$(YY$, 97, 8)) ' agrega subcodigo
              If REPLAORI% = 1 Then
                CODDX$ = CODORIG$(CI%, NC%)
              End If
              TETABU1$(1, CAITAB1%) = CODDX$
              TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
              If Frame15.Visible = True Then
                TETABU1$(2, CAITAB1%) = Mid$(YY$, 5, 64)
              End If
              TETABU1$(3, CAITAB1%) = CANS$
              TETABU1$(4, CAITAB1%) = NULOTE$
              TETABU1$(5, CAITAB1%) = UNID$
              TETABU1$(6, CAITAB1%) = NUOCO$
              TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
              TETABU1$(8, CAITAB1%) = NUOCOCLI$(Val(NUOCO$), CI%)
              TETABU1$(9, CAITAB1%) = NUPLANO$(CI%)
              '
              If SERINU$ <> "" Then
                CAITAB1% = CAITAB1% + 1
                For KI1% = 1 To 20
                  TETABU1$(KI1%, CAITAB1%) = ""
                Next KI1%
                TETABU1$(2, CAITAB1%) = "S.No." + SERINU$
              End If
              '
              NCIY% = NC%: CIIY% = CI%
              'Call DATADI(NCIY%, CIIY%, ORCO$, NORI$, NPLA$)
            End If
          End If
        End If
      Next YX&
      '
      Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
      '
      
    End If
    '
Return
    '
    '
999 End Sub
'
Sub GUARDAREMI()
    Call COPYSTRU("BOLREDET", "GUARREMI")
    Call BLANARCH("!GUARREMI")
    Call GRAREG("!GUARREMI", Text1, 1)
    Call GRAREG("!GUARREMI", Text2, 2)
    Call GRAREG("!GUARREMI", Text3, 3)
    Call GRAREG("!GUARREMI", Text4, 4)
    Call GRAREG("!GUARREMI", Text5, 5)
    Call GRAREG("!GUARREMI", Text6, 6)
    Call GRAREG("!GUARREMI", Text7, 7)
    Call GRAREG("!GUARREMI", Text8, 8)
    Call GRAREG("!GUARREMI", Text9, 9)
    Call GRAREG("!GUARREMI", Text10, 10)
    Call GRAREG("!GUARREMI", Text11, 11)
    Call GRAREG("!GUARREMI", Text12, 12)
    Call GRAREG("!GUARREMI", Text13, 13)
    Call GRAREG("!GUARREMI", Text14, 14)
    Call GRAREG("!GUARREMI", Text15, 15)
    Call GRAREG("!GUARREMI", "", 16)
    Call GRAREG("!GUARREMI", Text17, 17)
    Call GRAREG("!GUARREMI", Text18, 18)
    Call GRAREG("!GUARREMI", Text19, 19)
    Call GRAREG("!GUARREMI", Text20, 20)
    Call GRAREG("!GUARREMI", Text21, 21)
    Call GRAREG("!GUARREMI", Text22, 22)
    Call GRAREG("!GUARREMI", "", 23)
    Call GRAREG("!GUARREMI", "", 24)
    Call GRAREG("!GUARREMI", "", 25)
    Call GRAREG("!GUARREMI", Label10, 26)
    Call GRAREG("!GUARREMI", Label18, 27)
    Call GRAREG("!GUARREMI", Label26, 28)
    Call GRAREG("!GUARREMI", "", 29)
    Call GRAREG("!GUARREMI", "", 30)
    Call GRAREG("!GUARREMI", Text16(0), 31)
    Call GRAREG("!GUARREMI", Text16(1), 32)
    Call GRAREG("!GUARREMI", Text16(2), 33)
    Call GRAREG("!GUARREMI", Text16(3), 34)
    Call GRAREG("!GUARREMI", "", 35)
    Call GRAREG("!GUARREMI", "", 36)
    Call GRAREG("!GUARREMI", "", 37)
    Call GRAREG("!GUARREMI", "", 38)
    Call GRAREG("!GUARREMI", "", 39)
    Call GRAREG("!GUARREMI", "", 40)
    '
    URE& = 40
    FIN& = ULTREG&("!BOLREDET")
    For U1& = 1 To FIN&
       XX1$ = REGLEIDO$("!BOLREDET", U1&)
       Call GRAREG("!GUARREMI", XX1$, 40 + U1&)
       URE& = 40 + U1&
    Next U1&
    Call SETULTREG("!GUARREMI", URE&)
    Call CIERRARCH("!GUARREMI")
    '
End Sub
'
Sub BORRAGUARDA()
    '
    Call SETULTREG("!GUARREMI", 0)
    Call BLANARCH("!GUARREMI")
    REPIREMI% = 0
    '
End Sub
'
Private Sub Timer1_Timer()
   '
   TOTITEMS% = ULTREG&("!BOLREDET")
   Label18 = TRIM$(Str$(TOTITEMS%))
   '
End Sub

Sub RECUREMI()
    '
    REPIREMI% = 1
    '
    Text1 = TRIM$(REGLEIDO$("!GUARREMI", 1))
    Text2 = TRIM$(REGLEIDO$("!GUARREMI", 2))
    Text3 = TRIM$(REGLEIDO$("!GUARREMI", 3))
    Text4 = TRIM$(REGLEIDO$("!GUARREMI", 4))
    Text5 = TRIM$(REGLEIDO$("!GUARREMI", 5))
    Text6 = TRIM$(REGLEIDO$("!GUARREMI", 6))
    Text7 = TRIM$(REGLEIDO$("!GUARREMI", 7))
    Text8 = TRIM$(REGLEIDO$("!GUARREMI", 8))
    Text9 = TRIM$(REGLEIDO$("!GUARREMI", 9))
    Text10 = TRIM$(REGLEIDO$("!GUARREMI", 10))
    Text11 = TRIM$(REGLEIDO$("!GUARREMI", 11))
    Text12 = TRIM$(REGLEIDO$("!GUARREMI", 12))
    Text13 = TRIM$(REGLEIDO$("!GUARREMI", 13))
    Text14 = TRIM$(REGLEIDO$("!GUARREMI", 14))
    Text15 = TRIM$(REGLEIDO$("!GUARREMI", 15))
    'Text16 = TRIM$(REGLEIDO$("!GUARREMI", 1))
    Text17 = TRIM$(REGLEIDO$("!GUARREMI", 17))
    Text18 = TRIM$(REGLEIDO$("!GUARREMI", 18))
    Text19 = TRIM$(REGLEIDO$("!GUARREMI", 19))
    Text20 = TRIM$(REGLEIDO$("!GUARREMI", 20))
    Text21 = TRIM$(REGLEIDO$("!GUARREMI", 21))
    Text22 = TRIM$(REGLEIDO$("!GUARREMI", 22))
    '
    Label10 = TRIM$(REGLEIDO$("!GUARREMI", 26))
    Label18 = TRIM$(REGLEIDO$("!GUARREMI", 27))
    Label26 = TRIM$(REGLEIDO$("!GUARREMI", 28))
    '
    Text16(0) = TRIM$(REGLEIDO$("!GUARREMI", 31))
    Text16(1) = TRIM$(REGLEIDO$("!GUARREMI", 32))
    Text16(2) = TRIM$(REGLEIDO$("!GUARREMI", 33))
    Text16(3) = TRIM$(REGLEIDO$("!GUARREMI", 34))
    '
    FIN& = ULTREG&("!GUARREMI")
    For U1& = 41 To FIN&
       XX1$ = REGLEIDO$("!GUARREMI", U1&)
       Call GRAREG("!BOLREDET", XX1$, U1& - 40)
       URE& = U1& - 40
    Next U1&
    Call SETULTREG("!BOLREDET", URE&)
    Call CIERRARCH("!BOLREDET")
    '
    ' ARMAR NUMERO DE REMITO Y BORRAR MOVIMIENTOS DE STOCK
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    REMI$ = TRIM$(Text11.TEXT)
    FF% = FECHANUM(Text9.TEXT)
    PREFIREM$ = TRIM(Label10)
    NROREMIT$ = TRIM$(Mid$(REMI$, 4))
    While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
    NROREMIT$ = CMO$ + "-X-" + PREFIREM$ + "-" + NROREMIT$
    '
    Call COMUNI("Coloque el Formulario de Impresión\Correspondiente al Remito '" + NROREMIT$ + "' en Posición\en la Impresora y Pulse 'Aceptar Mensaje'.")
    '
    Call APERBASDAT("STOCKS")
    FEMOVD = Int(CDbl(CVDAT(FF%)))
    CLISEL% = XVALO(Text1)
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE DoPrilar = '" & NROREMIT$ & "' "
    SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechMov)) = " & FEMOVD
    Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With MoviTemp
20    On Error Resume Next
      .MoveFirst
      If Err > 1 Then
        On Error GoTo 0
        GoTo 30
        'Call MENSERR(24, "No se Encontró Detalle del Remito")
      End If
      On Error GoTo 0
      '
      ' VERIFICA SI EL REMITO TIENE ASOCIADO UN PEDIDO
      NROPEDID& = MoviTemp!NUPEDCLI
      If NROPEDID& > 0 Then
        ' ACTUALIZA LOS PEDIDOS QUE FUERON ENTREGADOS CON EL REMITO
          CODINTE% = MoviTemp!COD_INTE
          CANTSALD# = MoviTemp!CANTSALID
          Call APERBASDAT("PEDCLI")
          SQLY$ = "NroPed = " & NROPEDID& & " "
          SQLY$ = SQLY$ + "AND CodiInt = " & CODINTE% & " "
          SQLY$ = SQLY$ + "AND NroClie = " & CLISEL% & " "
          SQLY$ = SQLY$ + "AND CantEnt >= " & FORMATNUM$((CANTSALD# - 0.05), "F15.2") & " "
          SQLY$ = SQLY$ & " AND Status < 9 "
          '
          PedDeta.FindFirst SQLY$
          If PedDeta.NoMatch = False Then
            PedDeta.Edit
            ENTREGA# = PedDeta!CantEnt - CANTSALD#
            SALDO# = PedDeta!CanSaldo + CANTSALD#
            PedDeta!CantEnt = ENTREGA#
            PedDeta!CanSaldo = SALDO#
            PedDeta!NroRemi = ""
            PedDeta.Update
          End If
      End If
      '
      ' ANULACION DEL REMITO
      .Delete
      If .RecordCount > 0 Then
        GoTo 20
      End If
    End With
    ' FIN DE BORRAR MOVIMIENTOS DE STOCK
    '
30  Call Command5_Click
    '
End Sub
