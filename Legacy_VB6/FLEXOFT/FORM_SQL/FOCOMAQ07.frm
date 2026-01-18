VERSION 5.00
Begin VB.Form FOCOMAQ07 
   BackColor       =   &H00C0C0A0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Compra - Proveedores Externos"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11880
   Icon            =   "FOCOMAQ07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text14 
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
      Height          =   240
      Left            =   3570
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   8265
      Width           =   960
   End
   Begin VB.CommandButton Command7 
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
      Height          =   220
      Left            =   4545
      TabIndex        =   89
      Top             =   8265
      Width           =   160
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
      Height          =   240
      Left            =   4305
      TabIndex        =   88
      Top             =   7980
      Width           =   400
   End
   Begin VB.CommandButton Command18 
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
      Height          =   285
      Left            =   2730
      TabIndex        =   83
      Top             =   7875
      Width           =   180
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1575
      TabIndex        =   82
      Top             =   7875
      Width           =   1170
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1575
      TabIndex        =   81
      Top             =   8190
      Width           =   1340
   End
   Begin VB.CommandButton Command15 
      Caption         =   "T.Lib"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   8910
      TabIndex        =   79
      Top             =   8715
      Width           =   645
   End
   Begin VB.CommandButton Command21 
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
      Height          =   250
      Index           =   1
      Left            =   4725
      TabIndex        =   78
      Top             =   8715
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
      Height          =   250
      Index           =   1
      Left            =   4305
      TabIndex        =   77
      Text            =   " "
      Top             =   8715
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
      Height          =   250
      Index           =   1
      Left            =   5040
      TabIndex        =   76
      Text            =   " "
      Top             =   8715
      Width           =   4215
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0A0&
      Caption         =   "Valorizar en"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   105
      TabIndex        =   72
      Top             =   8505
      Width           =   3060
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
         Height          =   240
         Left            =   630
         TabIndex        =   75
         Top             =   280
         Width           =   150
      End
      Begin VB.TextBox Text15 
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
         Height          =   240
         Left            =   210
         MaxLength       =   2
         TabIndex        =   74
         TabStop         =   0   'False
         Text            =   " "
         Top             =   280
         Width           =   435
      End
      Begin VB.TextBox Text16 
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
         Height          =   240
         Left            =   900
         TabIndex        =   73
         TabStop         =   0   'False
         Text            =   " "
         Top             =   280
         Width           =   2010
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0C0A0&
      Caption         =   "Pendientes de Recepción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   2280
      Left            =   0
      TabIndex        =   63
      Top             =   8925
      Width           =   10580
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10500
         TabIndex        =   70
         Top             =   315
         Width           =   10565
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Código          Descripción                           Cantid.   F.Entreg  Atr    O/C"
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
            Left            =   30
            TabIndex        =   71
            Top             =   105
            Width           =   10515
         End
      End
      Begin VB.ListBox List2 
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
         Height          =   1185
         Left            =   0
         MultiSelect     =   1  'Simple
         TabIndex        =   69
         Top             =   750
         Width           =   10565
      End
      Begin VB.Frame Frame10 
         Caption         =   "Frame10"
         Height          =   1730
         Left            =   1890
         TabIndex        =   68
         Top             =   210
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame11"
         Height          =   1730
         Left            =   5985
         TabIndex        =   67
         Top             =   210
         Width           =   15
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame12"
         Height          =   1730
         Left            =   7665
         TabIndex        =   66
         Top             =   210
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame13"
         Height          =   1730
         Left            =   8715
         TabIndex        =   65
         Top             =   210
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "Frame14"
         Height          =   1730
         Left            =   9345
         TabIndex        =   64
         Top             =   210
         Width           =   15
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10560
         Y1              =   1940
         Y2              =   1940
      End
      Begin VB.Line Line2 
         X1              =   10565
         X2              =   10565
         Y1              =   315
         Y2              =   1940
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0A0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7575
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame16 
         BackColor       =   &H00C0C0A0&
         Caption         =   "C.U.I.T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   5670
         TabIndex        =   86
         Top             =   30
         Width           =   1905
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   105
            TabIndex        =   87
            TabStop         =   0   'False
            Text            =   " "
            Top             =   315
            Width           =   1650
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00004000&
         Height          =   9000
         Left            =   10965
         ScaleHeight     =   8940
         ScaleWidth      =   1005
         TabIndex        =   52
         Top             =   0
         Width           =   1065
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cotiz."
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
            Left            =   130
            Picture         =   "FOCOMAQ07.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Pedido de Cotización"
            Top             =   5225
            Width           =   650
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "More"
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
            Left            =   130
            Picture         =   "FOCOMAQ07.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Opciones Varias - Modificaciones y Consultas"
            Top             =   4535
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ver"
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
            Left            =   130
            Picture         =   "FOCOMAQ07.frx":0896
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
            Top             =   3845
            Width           =   650
         End
         Begin VB.CommandButton Command3 
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
            Left            =   130
            Picture         =   "FOCOMAQ07.frx":0BA0
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   790
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Prove."
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
            Left            =   130
            Picture         =   "FOCOMAQ07.frx":0EAA
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Acceso a Base de Datos de Proveedores"
            Top             =   1550
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Rep."
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
            Left            =   130
            Picture         =   "FOCOMAQ07.frx":11B4
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Acceso a Maestro de Repuestos e Insumos"
            Top             =   2240
            Width           =   650
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Setup"
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
            Left            =   130
            Picture         =   "FOCOMAQ07.frx":14BE
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Configuración de Funcionamiento de Compras"
            Top             =   3045
            Width           =   650
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Exit"
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
            Left            =   130
            Picture         =   "FOCOMAQ07.frx":1900
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   100
            Width           =   650
         End
         Begin VB.CommandButton Command22 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Check"
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
            Left            =   105
            Picture         =   "FOCOMAQ07.frx":1A4A
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Aprobar Ordenes de Compra"
            Top             =   8160
            Width           =   650
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "O.K."
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
            Left            =   130
            Picture         =   "FOCOMAQ07.frx":1D54
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Aprueba, Graba e Imprime Orden de Compra"
            Top             =   5985
            Width           =   650
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -315
            Picture         =   "FOCOMAQ07.frx":205E
            Top             =   6765
            Width           =   1515
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   47
         ToolTipText     =   "Lista de Proveedores con Entregas Programadas"
         Top             =   8040
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Frame Frame15 
         Caption         =   "Frame15"
         Height          =   3400
         Left            =   10450
         TabIndex        =   46
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0A0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2050
         Left            =   0
         TabIndex        =   22
         Top             =   5170
         Width           =   10800
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0A0&
            Caption         =   "Resumen Final"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   1965
            Left            =   6615
            TabIndex        =   23
            Top             =   60
            Width           =   4130
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
               Left            =   1365
               MaxLength       =   14
               TabIndex        =   28
               Top             =   1260
               Width           =   750
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
               Height          =   285
               Index           =   4
               Left            =   2415
               TabIndex        =   49
               Top             =   1260
               Width           =   1500
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
               Height          =   285
               Left            =   735
               MaxLength       =   14
               TabIndex        =   26
               Top             =   630
               Width           =   1380
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
               Height          =   285
               Index           =   1
               Left            =   2415
               TabIndex        =   30
               TabStop         =   0   'False
               Text            =   " "
               Top             =   630
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
               Height          =   285
               Index           =   0
               Left            =   2415
               TabIndex        =   29
               Top             =   315
               Width           =   1500
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
               Left            =   1365
               MaxLength       =   14
               TabIndex        =   27
               Top             =   945
               Width           =   750
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
               Height          =   285
               Index           =   2
               Left            =   2415
               TabIndex        =   25
               Top             =   945
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
               Height          =   285
               Index           =   3
               Left            =   2415
               TabIndex        =   24
               Top             =   1575
               Width           =   1500
            End
            Begin VB.Label Label4 
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
               Left            =   2205
               TabIndex        =   51
               Top             =   1320
               Width           =   225
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Percepcion:"
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
               Left            =   -1680
               TabIndex        =   50
               Top             =   1320
               Width           =   2955
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
               Left            =   2205
               TabIndex        =   37
               Top             =   720
               Width           =   225
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Desc:"
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
               Left            =   -1260
               TabIndex        =   36
               Top             =   675
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
               Left            =   1365
               TabIndex        =   35
               Top             =   375
               Width           =   960
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "I.V.A.:"
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
               Left            =   -630
               TabIndex        =   34
               Top             =   1050
               Width           =   1905
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total Orden de Compra:"
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
               Left            =   -630
               TabIndex        =   33
               ToolTipText     =   "Haga un Doble-Click para Emitir CON / SIN Precios"
               Top             =   1620
               Width           =   2955
            End
            Begin VB.Label Label9 
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
               Left            =   2205
               TabIndex        =   32
               Top             =   1020
               Width           =   225
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Pesos"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   12
                  Charset         =   0
                  Weight          =   900
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   330
               Left            =   105
               TabIndex        =   31
               Top             =   260
               Visible         =   0   'False
               Width           =   1275
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C0A0&
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
            ForeColor       =   &H00404000&
            Height          =   1965
            Left            =   160
            TabIndex        =   38
            Top             =   60
            Width           =   6270
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
               Height          =   250
               Left            =   1155
               TabIndex        =   40
               Top             =   645
               Width           =   4950
            End
            Begin VB.CommandButton Command21 
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
               Height          =   250
               Index           =   0
               Left            =   1575
               TabIndex        =   43
               Top             =   315
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
               Height          =   250
               Index           =   0
               Left            =   1155
               TabIndex        =   42
               Text            =   " "
               Top             =   315
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
               Height          =   250
               Index           =   0
               Left            =   1890
               TabIndex        =   41
               Top             =   315
               Width           =   4215
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
               Height          =   780
               Left            =   1155
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   39
               Top             =   1020
               Width           =   4950
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "L.Entrega:"
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
               TabIndex        =   48
               Top             =   690
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
               Index           =   0
               Left            =   -525
               TabIndex        =   45
               Top             =   370
               Width           =   1590
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Observac:"
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
               TabIndex        =   44
               Top             =   1020
               Width           =   4110
            End
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   3400
         Index           =   6
         Left            =   8665
         TabIndex        =   21
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   3400
         Index           =   5
         Left            =   7465
         TabIndex        =   20
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   3400
         Index           =   4
         Left            =   9400
         TabIndex        =   19
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   3400
         Index           =   3
         Left            =   5725
         TabIndex        =   18
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   3400
         Index           =   2
         Left            =   6250
         TabIndex        =   17
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   3400
         Index           =   0
         Left            =   2050
         TabIndex        =   16
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0A0&
         Caption         =   "Ident"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   750
         Left            =   7665
         TabIndex        =   15
         Top             =   30
         Width           =   3060
         Begin VB.CommandButton Command9 
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
            Height          =   285
            Left            =   2790
            TabIndex        =   94
            Top             =   315
            Width           =   195
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
            Height          =   285
            Left            =   1845
            TabIndex        =   93
            TabStop         =   0   'False
            Text            =   " "
            Top             =   315
            Width           =   960
         End
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   420
            TabIndex        =   92
            Text            =   " "
            Top             =   315
            Width           =   915
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "F:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   330
            Left            =   1050
            TabIndex        =   96
            Top             =   345
            Width           =   750
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "N:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   330
            Left            =   -840
            TabIndex        =   95
            Top             =   315
            Width           =   1170
         End
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
         Height          =   285
         Left            =   6195
         MaxLength       =   30
         TabIndex        =   14
         Text            =   " "
         Top             =   945
         Width           =   4530
      End
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
         Left            =   160
         TabIndex        =   12
         Top             =   1995
         Width           =   10565
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   120
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6950
         Left            =   11025
         TabIndex        =   11
         Top             =   105
         Width           =   850
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   160
         ScaleHeight     =   375
         ScaleWidth      =   10500
         TabIndex        =   9
         Top             =   1575
         Width           =   10565
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Código          Descripción                    U/S      Cant   Pr.Unit % Dc  F.Entreg"
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
            Left            =   30
            TabIndex        =   10
            ToolTipText     =   "Haga Doble-Click para Blanquear Pre-Carga y Habilitar Carga Manual"
            Top             =   105
            Width           =   10515
         End
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
         Height          =   285
         Left            =   2310
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
         Width           =   3225
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
         Height          =   285
         Left            =   1470
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
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
         Height          =   285
         Left            =   1470
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   520
         Width           =   4065
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
         Height          =   285
         Left            =   1470
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   4065
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
         Left            =   945
         TabIndex        =   1
         Top             =   945
         Width           =   195
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
         Height          =   285
         Left            =   210
         MaxLength       =   5
         TabIndex        =   0
         Text            =   " "
         Top             =   945
         Width           =   750
      End
      Begin VB.Line Line4 
         X1              =   -240
         X2              =   11760
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line5 
         X1              =   180
         X2              =   10720
         Y1              =   4980
         Y2              =   4980
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "OC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   750
         Left            =   210
         TabIndex        =   13
         Top             =   0
         Width           =   1380
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   200
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   220
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   315
         Shape           =   4  'Rounded Rectangle
         Top             =   165
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   400
         Left            =   210
         Shape           =   4  'Rounded Rectangle
         Top             =   105
         Width           =   960
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "At.Sr:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   5670
         TabIndex        =   8
         Top             =   945
         Width           =   1695
      End
      Begin VB.Line Line8 
         X1              =   10725
         X2              =   10725
         Y1              =   1575
         Y2              =   5000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   105
         TabIndex        =   7
         Top             =   630
         Width           =   960
      End
   End
   Begin VB.Label Label8 
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
      Height          =   285
      Left            =   5565
      TabIndex        =   97
      Top             =   8085
      Width           =   1170
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dias:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3825
      TabIndex        =   91
      Top             =   8040
      Width           =   435
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Dest:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1050
      TabIndex        =   85
      ToolTipText     =   "Dobe-Click para Asignar Destino por Item"
      Top             =   7935
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sect:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1050
      TabIndex        =   84
      ToolTipText     =   "Dobe-Click para Asignar Sector por Item"
      Top             =   8265
      Width           =   1065
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Embalaje:"
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
      Left            =   105
      TabIndex        =   80
      Top             =   8775
      Width           =   4110
   End
   Begin VB.Menu Archima 
      Caption         =   "Archivos Maestros"
      Begin VB.Menu Amastock 
         Caption         =   "Articulos de Stock"
      End
      Begin VB.Menu Provedor 
         Caption         =   "Proveedores y Contratistas"
      End
      Begin VB.Menu plh3 
         Caption         =   "-"
      End
      Begin VB.Menu Nuorig 
         Caption         =   "Números Originales"
      End
      Begin VB.Menu plh4 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Orcom 
      Caption         =   "Ordenes de Compra"
      Begin VB.Menu Nue_Orco 
         Caption         =   "Nueva Orden de Compra"
      End
      Begin VB.Menu phl1 
         Caption         =   "-"
      End
      Begin VB.Menu Ver_Ocom 
         Caption         =   "Ver y Re-Imprimir"
      End
      Begin VB.Menu plh2 
         Caption         =   "-"
      End
      Begin VB.Menu MarItCum 
         Caption         =   "Marcar Items Cumplidos"
      End
      Begin VB.Menu plh22 
         Caption         =   "-"
      End
      Begin VB.Menu Anul_Ocom 
         Caption         =   "Anular O/Compra"
      End
   End
   Begin VB.Menu ConLis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu CoCoPen 
         Caption         =   "Control de Compras Pendientes"
         Begin VB.Menu ConuOC 
            Caption         =   "Por Numero de O/C"
            Begin VB.Menu Screen1 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu Print1 
               Caption         =   "Imprimir / Exportar"
            End
         End
         Begin VB.Menu CoProv 
            Caption         =   "Por Proveedor"
            Begin VB.Menu Screen2 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu Print2 
               Caption         =   "Imprimir / Exportar"
            End
         End
         Begin VB.Menu CoMat 
            Caption         =   "Por Material"
            Begin VB.Menu Screen3 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu Print3 
               Caption         =   "Imprimir / Exportar"
            End
         End
      End
   End
   Begin VB.Menu Config 
      Caption         =   "Configuración"
      Begin VB.Menu MnuTabl 
         Caption         =   "Tablas"
         Begin VB.Menu MnuConPa 
            Caption         =   "Condiciones de Pago"
         End
         Begin VB.Menu MnuCueCon 
            Caption         =   "Cuentas Contables"
         End
         Begin VB.Menu MnuTas 
            Caption         =   "Tasas de I.V.A."
         End
         Begin VB.Menu MnuNumOr 
            Caption         =   "Numeros Originales"
         End
      End
      Begin VB.Menu Nuorig1 
         Caption         =   "Formularios de Impresión"
      End
   End
End
Attribute VB_Name = "FOCOMAQ07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$
Dim MUECOPEN%
Dim EMICONSI%
Dim MODOULPRE%
Dim RAI$ '\\\OT-05-0676-WAR-20/10/05///'REGLAMENTO DE COMPRA ANEXOS

Private Sub Anul_Ocom_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call LISCOMAQ04.Command14_Click
End Sub

Private Sub command1_Click()
   Call SELECHO("ACREDOR")
   NCLIE% = Val(TRIM$(VALACT1$("ACREDOR")))
   If NCLIE% > 0 Then
     If REGICLI&((-1) * NCLIE%) > 0 Then
       Call BLANFORMU
       Text1.TEXT = TRIM$(Str$(NCLIE%))
       Call CARDACLI
       Call CARCOPEN
       Call CARCOANT
       If ULTREG&("!OCOMDETA") > 0 Then
            Call CARLISCO
            Call CALTOCOM
         Else
            Call CARDETCOM
       End If
     End If
   End If
End Sub

Private Sub Command10_Click()
   Call SELECHO("PROVEPLA")
   NCLIE% = Val(TRIM$(VALACT1$("PROVEPLA")))
   If NCLIE% > 0 Then
     If REGICLI&((-1) * NCLIE%) > 0 Then
       Call BLANFORMU
       Text1.TEXT = TRIM$(Str$(NCLIE%))
       Call CARDACLI
       Call CARCOPEN
       Call CARCOANT
       If ULTREG&("!OCOMDETA") > 0 Then
            Call CARLISCO
            Call CALTOCOM
         Else
            Call CARDETCOM
       End If
     End If
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   ACRE_MAQ04.Show 1
'   Call CIERRARCH("*.*")
'   ACONEC$ = "ARCHIG04/AMAPRO"
'   Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    '
    Command12.Enabled = False
    '
3   Call SELECHO("TAMONED")
    MMI% = Val(VALACT1$("TAMONED"))
    If MMI% > 0 Then
      Text15 = TRIM$(Str$(MMI%))
    End If
    '
99  Command12.Enabled = True
    '
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   List1.Clear
   Call SETULTREG("!OCOMDETA", 0)
   LISCOMAQ04.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    '
    Command14.Enabled = False
    '
    '\\\OT-05-0693-WAR-FIN///
'10  FORSELREG.Label3 = Text11 '\\**//
'    Call SELEC_REGLACOM(OKX%)
'    If OKX% < 1 Then GoTo 99
    '
    NCI% = Val(Text1)
    NCIX% = NCI% * (-1)
    EMAI$ = TRIM$(EMAILCLI$(NCIX%))
    If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
       GoTo 20 'NO TIENE DIRECCION O DIRECCION NO VALIDA
    End If
    RESP22% = COMALTER%("¿Envia por Mail?\\Si, Enviar.\No Enviar por Mail")
    If RESP22% = 1 Then GoTo 15
    If RESP22% = 2 Then GoTo 20
    '
15  SENDBYMAIL$ = ""
    PORMAIL% = 1
    SENDBYMAIL$ = EMAI$
    '
    For Each PRX1 In Printers
       If InStr(UCase$(PRX1.DeviceName), "PDFWRITER") > 0 Then GoTo 20
    Next
    Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Impresora Virtual PDFWRITER.")
    SENDBYMAIL = ""
    '
20  Call GRACOMPRA(OKX%)
    If OKX% > 0 Then
     Call PRICOMPRA
     If RAI$ <> "" Then
       Call PRIREGLACOM 'IMPRIME REGLAMENTO DE COMPRA
     End If
     Call BLANFORMU
     Call SETULTREG("!OCOMDETA", 0)
     Call CIERRARCH("*.*")
     Text1.TEXT = ""
     'AGREGADO PARA ENVIAR POR MAIL
     If SENDBYMAIL$ <> "" Then
         Call RECONECTA
         Exit Sub
     End If
     'HASTA ACA AGREGADO
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
   End If
99 Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   FORCOLIB04.Show 1
   Command15.Enabled = True
End Sub

Private Sub Command18_Click()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
   '
   If ULTREG&("ECUECON") < 1 Then
      Call MENSERR(24, "Falta Definir Plan de Cuentas Contables")
      Call CARCUEINGAS
      Exit Sub
   End If
   '
   RG1$ = REGSEL$("ECUECON")
   If Len(RG1$) > 4 Then
      Text7.TEXT = Left$(RG1$, 12)
   End If
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

'Private Sub Command2_LostFocus()
'   On Error Resume Next
'   Text1.SetFocus
'   On Error GoTo 0
'End Sub
Private Sub Command22_Click()
    Command22.Enabled = False
  FOAPROC04.Show 1
    'Screen.MousePointer = 11
    '
    'JJ& = 0
    'For U& = 1 To ULTREG&("PEDENCA")
2   '  XX$ = REGLEIDO$("PEDENCA", U&)
    '  If CVS(Left$(XX$, 4)) > 0 Then
    '    NPX& = CVS(Left$(XX$, 4))
    '    NCLIE% = CVI(Mid$(XX$, 55, 2))
    '    FEX = CVI(Mid$(XX$, 53, 2))
    '    ANULX$ = "": If Mid$(XX$, 267, 1) = Chr$(9) Then ANULX$ = "(Anul) "
    '    YY$ = MKS$(NPX&) + ANULX$ + TRIM$(Mid$(XX$, 5, 48))
    '    JJ& = JJ& + 1
    '    Call GRAREG("INDIPEDI", YY$, JJ&)
    '  End If
    'Next U&
    'Call SETULTREG("INDIPEDI", JJ&)
    'Screen.MousePointer = 1
    '
    'CONPEDID.Show 1
    '
    Command22.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    '
    Call HELPONLINE("FOCOMAQ")
    '
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call LISCOMAQ04.Command22_Click
99  Command4.Enabled = True
   End Sub

Sub Command5_Click()
    Command5.Enabled = False
    Call CIERRARCH("*.*")
    If DERACCE%("STPOCOM", "Configuracion de Funcionamiento Compras") = 2 Then
      OPCOMAQ04.Show 1
    End If
    Command5.Enabled = True
End Sub


Private Sub Command6_Click()
  '
  FOPCMAQ04.Show 1
  '
End Sub

'
Private Sub Command7_Click()
    Command7.Enabled = False
    If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
    End If
    Call SELECHO("SELFECHA")
    VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
    If VDEVU$ <> "" Then
      Text14.TEXT = VALACT1$("SELFECHA")
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    AMAS_MAQ04.Show 1
'    ACONEC$ = "ARCHIG04/AMASTO"
'    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Desanuoc_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call OPCICOM04.Command10_Click
End Sub

Private Sub DevoConsi_Click()
   OPCICOMP.Command15_Click
End Sub


Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub


Private Sub Entrecon_Click()
   OPCICOM04.Command2_Click
End Sub

Private Sub Exit_Click()
   Call Command2_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  While WindowState <> 0
    WindowState = 0
    openforms = DoEvents
    AppActivate Caption
    SendKeys "%" & Chr$(32) & Chr$(13), True
  Wend
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    '
    EPAC$ = UCase$(TRIM$(EMPREAC$))
    If EPAC$ <> "" Then
10     Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If ULTREG&("PROVEPLA") > 0 Then
      Command10.Visible = True
    End If
    '
    If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
      Nuorig.Enabled = True
      Nuorig1.Enabled = True
    End If
    '
    'Show
    DoEvents
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    Text14.TEXT = HOS$
    Text11.TEXT = TRIM$(Str$(PRONUOCO&))
    Text12.TEXT = FORMATNUM$(ALIVA, "F5.2")
    '
    Text15.TEXT = TRIM$(CONTROL$("OCOMPRA", "MONECOS"))
    If Text15 = "" Then Text15 = "1"      ' MONEDA DE VALUACIóN
    MONEMI% = 1
    '
    MODOULPRE% = 1
    If CONTROL("OCOMPRA", "PRECIDEF") = "COSTOREP" Then MODOULPRE% = 2
    '
    DoEvents
    '
    Call SETULTREG("!OCOMDETA", 0)
    Call ABRECOMPRAS
    '
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
End Sub

Private Sub Label11_DblClick()
'   '
'   Static CAMON%, DMONE$(), TCMONE()
'   If CAMON% < 1 Then
'     CAMON% = ULTREG&("TAMONED")
'     If CAMON% < 1 Then
'         CAMON% = 1
'         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
'         DMONE$(1) = "Pesos"
'         TCMONE(1) = 1
'       Else
'         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
'         For U& = 1 To CAMON%
'           XX$ = REGLEIDO$("TAMONED", U&)
'           DMX$ = TRIM$(Left$(XX$, 8)) + " "
'           PPXX% = InStr(DMX$, " ")
'           DMX$ = Left$(DMX$, PPXX% - 1)
'           '
'           VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
'           DMONE$(U&) = DMX$
'           TCMONE(U) = VHX
'         Next U&
'     End If
'   End If
'   '
'   MONEMIA% = MONEMI%
'   MONEMI% = MONEMI% + 1
'   If MONEMI% > CAMON% Then MONEMI% = 1
'   Label11 = DMONE$(MONEMI%)
'   Label11.ForeColor = RGB(0, 0, 255)
'   If MONEMI% > 1 Then
'      Label11.ForeColor = RGB(255, 0, 0)
'   End If
'   Label11.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
'   '
'   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
'   TICAMBX# = TICAMONE(MONEMI%)
'   '
'   For U& = 1 To ULTREG&("!OCOMDETA")
'     YX$ = REGLEIDO$("!OCOMDETA", U&)
'     PREDOL# = CVD(Mid$(YX$, 67, 8)) * COEFCON
'     Call REPLA(YX$, MKD$(PREDOL#), 67, 8)
'     Call GRAREG("!OCOMDETA", YX$, U&)
'   Next U&
'   Call CIERRARCH("!OCOMDETA")
'   Call CARLISCO
'   Call CALTOCOM
'   '
End Sub

Private Sub Label14_dblClick()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
   '
   Call CARDETCOM1
End Sub

Private Sub Label2_dblClick()
   Call Label14_dblClick
End Sub

Private Sub Label22_DBLClick()
   List1.Clear
   Call SETULTREG("!OCOMDETA", 0)
   Unload FORVENTB
   Call CARDETCOM
End Sub

Private Sub Label8_Click()
   Call SELECHO("INDIACRE")
   CPROX$ = TRIM$(VALACT1$("INDIACRE"))
   If CPROX$ <> "" Then
     NCLIE% = NUMEPRO%(CPROX$)
     If NCLIE% > 0 Then
       If REGICLI&((-1) * NCLIE%) > 0 Then
         Call BLANFORMU
         Text1.TEXT = TRIM$(Str$(NCLIE%))
         Call CARDACLI
         Call CARCOPEN
         Call CARCOANT
         If ULTREG&("!OCOMDETA") > 0 Then
              Call CARLISCO
              Call CALTOCOM
           Else
              Call CARDETCOM
         End If
       End If
     End If
   End If
End Sub

Private Sub Label8_DblClick()
   Call Label8_Click
End Sub

Private Sub List1_Click()
   List1.ToolTipText = DESCRIT$(CODINT%(Left$(List1.TEXT, 15)))
   MUECOPEN% = 1
   For U& = 1 To List2.ListCount
     List2.Selected(U& - 1) = False
     If Left$(List2.List(U& - 1), 15) = Left$(List1.TEXT, 15) Then
       List2.Selected(U& - 1) = True
     End If
   Next U&
   MUECOPEN% = 0
End Sub

Private Sub List1_DblClick()
    Call CARDETCOM
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
     Call MUESTRADATCOD(Left$(List1.TEXT, 15))
   End If
   '
End Sub

Private Sub List2_Click()
   If MUECOPEN% < 1 Then
     For U& = 1 To List2.ListCount
       List2.Selected(U& - 1) = False
     Next U&
   End If
End Sub

Private Sub List2_DblClick()
    NPX& = Val(Mid$(List2.TEXT, 79, 6))
    If NPX& > 0 Then
      Call SHOWOCOM(NPX&)
    End If
End Sub

Private Sub Ltota_DblClick(Index As Integer)
   For KKI% = 0 To 4
     Total(KKI%).Enabled = Not (Total(KKI%).Enabled)
   Next KKI%
End Sub

Private Sub MarItCum_Click()
  '
  Call LISCOMAQ04.Command7_Click
  '
End Sub

Private Sub Amastock_Click()
   '
   AMAS_MAQ04.Show 1
   '
End Sub

Private Sub MnuConPa_Click()
  '
  OPCOMAQ04.Option1(10).Value = True
  Call OPCOMAQ04.Option1_DblClick(10)
  '
End Sub

Private Sub MnuCueCon_Click()
  '
  OPCOMAQ04.Option1(11).Value = True
  Call OPCOMAQ04.Option1_DblClick(11)
  '
End Sub

Private Sub MnuNumOr_Click()
  '
  OPCOMAQ04.Option1(13).Value = True
  Call OPCOMAQ04.Option1_DblClick(13)
  '
End Sub

Private Sub MnuTas_Click()
  '
  OPCOMAQ04.Option1(12).Value = True
  Call OPCOMAQ04.Option1_DblClick(12)
  '
End Sub

Private Sub Modi_Ocom_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call OPCICOM04.Command13_Click
End Sub

Private Sub Nue_Orco_Click()
 Call command1_Click
End Sub

Private Sub Nuorig_Click()
  '
  OPCOMAQ04.Option1(13).Value = True
  Call OPCOMAQ04.Option1_DblClick(13)
  '
End Sub

Private Sub Nuorig1_Click()
   '
   OPCOMAQ04.Show 1
   '
End Sub

Private Sub Por_MRP_Click()
   Call Command10_Click
End Sub

Private Sub PORDESCU_Change()
   Call CALTOCOM
End Sub

Private Sub PORDESCU_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub

Private Sub Print1_Click()
   OPLISCO04.Option2.Value = True
   Call OPLISCO04.Command3_Click
End Sub

Private Sub Print2_Click()
   OPLISCO04.Option2.Value = True
   Call OPLISCO04.Command4_Click
End Sub

Private Sub Print3_Click()
   OPLISCO04.Option2.Value = True
   Call OPLISCO04.Command5_Click
End Sub

Private Sub Print4_Click()
   OPLISCO04.Option2.Value = True
   Call OPLISCO04.Command9_Click
End Sub

Private Sub Provedor_Click()
   Command11_Click
End Sub

Private Sub Screen1_Click()
   LISCOMAQ04.Option1.Value = True
   Call LISCOMAQ04.Command3_Click
End Sub

Private Sub Screen2_Click()
   LISCOMAQ04.Option1.Value = True
   Call LISCOMAQ04.Command4_Click
End Sub

Private Sub Screen3_Click()
   LISCOMAQ04.Option1.Value = True
   Call LISCOMAQ04.Command5_Click
End Sub

Private Sub Screen4_Click()
   LISCOMAQ04.Option1.Value = True
   Call LISCOMAQ04.Command9_Click
End Sub

Private Sub Setup_Click()
  OPCICOM04.Command18_Click
End Sub

Private Sub Text1_DblClick()
'  Call SELECHO("PROVEPLA")
'  NCLIE% = Val(TRIM$(VALACT1$("PROVEPLA")))
'  If NCLIE% > 0 Then
'    If REGICLI&((-1) * NCLIE%) > 0 Then
'      Call BLANFORMU
'      Text1.TEXT = TRIM$(Str$(NCLIE%))
'      Call CARDACLI
'      Call CARCOPEN
'      Call CARCOANT
'      If ULTREG&("!OCOMDETA") > 0 Then
'           Call CARLISCO
'           Call CALTOCOM
'        Else
'           Call CARDETCOM
'      End If
'    End If
'  End If
  Call command1_Click
End Sub

Private Sub Text1_GotFocus()
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
   '
   If MONEMI% < 1 Then
     Label11.ForeColor = RGB(0, 0, 255)
     Label11 = "Pesos"
     MONEMI% = 1
     Label11.ToolTipText = "Doble-Click para 'Dólares'"
   End If
   '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If VALICLIE% > 0 Then
      Call BLANFORMU
      Call CARDACLI
      Call CARCOPEN
      Call CARCOANT
      If ULTREG&("!OCOMDETA") > 0 Then
           Call CARLISCO
           Call CALTOCOM
         Else
           Call CARDETCOM
      End If
    End If
  End If
End Sub

Private Sub Text11_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   If ULTREG&("PEDENCA") > 2 Then
      On Error Resume Next
      Text10.SetFocus
      On Error GoTo 0
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
   NPP& = PRONUOCO&
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

Private Sub Text12_Change()
   Call CALTOCOM
End Sub

Private Sub Text12_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text13_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text15_Change()
   Text16 = ECOARCH$("TAMONED", Chr$(Val(Text15)))
End Sub

Private Sub Text17_Change()
   Call CALTOCOM
End Sub

Private Sub Text17_GotFocus()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
End Sub

Private Sub Text19_Change()
    '
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text19 = ""
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
'    FEMIS% = FECHANUM(Text9)
'    FENTREG% = FEMIS%
'    '
'    If Text19 <> "" Then
'     If IsNumeric(Text19) = True Then
'      While DIENTRE%(FEMIS%, FENTREG%) < CInt(Text19)
'         FENTREG% = DIPOST(FENTREG)
'      Wend
'      Text14 = FECHATEX$(FENTREG%)
'     End If
'    End If
    '\\\OT-05-0690-WAR-24/10/05///
    '
    FEMIS% = FECHANUM(Text9)
    FENTREG% = FEMIS%
    '
    If Text19 <> "" Then
      If IsNumeric(Text19) = True And Val(Text19) < 7300 Then 'MENOR A 20 AÑOS EN DIAS
        While DIENTRE%(FEMIS%, FENTREG%) < CInt(Text19)
          FENTREG% = DIPOST(FENTREG)
        Wend
        Text14 = FECHATEX$(FENTREG%)
      Else
        Call MENSERR(24, "Cantidad de Días no Válida.")
      End If
    Else
      HOII% = HOY(HOS$)
      Text14.TEXT = HOS$
    End If
    '\\\OT-05-0690-WAR-FIN///
    '
End Sub
Private Sub Text2_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub TEXT21_DblClick(Index As Integer)
    '
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
    If Index = 0 Then
        ARCHEC$ = "CONPAG": LOCA% = 1
      ElseIf Index = 1 Then
        ARCHEC$ = "TAEMBALA": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
    '
End Sub

Private Sub TEXT21_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      Exit Sub
   End If
   TEXT21(Index).TEXT = TRIM$(TEXT21(Index).TEXT)
   TEXT21(Index).SelStart = 0
   TEXT21(Index).SelLength = Len(TEXT21(Index).TEXT)
End Sub

Private Sub Text3_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text4_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text5_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text21_Change(Index As Integer)
    '
    If Val(TEXT21(Index).TEXT) < 0 Or Val(TEXT21(Index).TEXT) > 255 Then
       TEXT21(Index).TEXT = ""
    End If
    '
    If Index = 0 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", Chr$(Val(TEXT21(Index).TEXT)))
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("TAEMBALA", Chr$(Val(TEXT21(Index).TEXT)))
    End If
    '
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
        ARCHEC$ = "CONPAG": LOCA% = 1
      ElseIf Index = 1 Then
        ARCHEC$ = "TAEMBALA": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
    '
End Sub
   '
Function VALICLIE%()
   VCC% = 0
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE% = Val(Text1.TEXT)
       If ECOARCH$("PROVED1", MKI$(NCLIE%)) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function

Sub BLANFORMU()
   Call SETULTREG("!OCOMDETA", 0)
   Call CIERRARCH("!OCOMDETA")
   Call BLANARCH("!OCOMDETA")
   Call SETULTREG("!1OCOMDET", 0)
   Call CIERRARCH("!1OCOMDET")
   Call BLANARCH("!1OCOMDET")
   On Error Resume Next
   Unload FORVENTB
   On Error GoTo 0
   Label8 = ""
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   'Label2 = ""
   Text8.TEXT = ""
   Text10 = ""
   Text12.TEXT = FORMATNUM$(ALIVA, "F5.2")
   Text13.TEXT = ""
   Text6.TEXT = ""
   Text7.TEXT = ""
   Text18.TEXT = ""
   Text19.TEXT = ""
   List1.Clear: List1.ToolTipText = ""
   List2.Clear
   '
   TEXT21(0).TEXT = "": Eco21(0) = ""
   TEXT21(1).TEXT = "": Eco21(1) = ""
   'TEXT21(2).TEXT = "": Eco21(2) = ""
   '
   PORDESCU.TEXT = ""
   Text17 = ""
   For KKI% = 0 To 4
      Total(KKI%).Enabled = True
      Total(KKI%).TEXT = ""
   Next KKI%
   '
   'MARCONOTA.Visible = False
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Text11.TEXT = TRIM$(Str$(PRONUOCO&))
   '
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    '
    NC% = (-1) * Val(Text1.TEXT)
    If NC% < 0 Then
      If STATCLI%(NC%) > 0 Then
        Label8 = CODICLI$(NC%)
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        'Label2 = TRIM$(POSIVAC$(NC%))
        Text8.TEXT = TRIM$(CUITCLI$(NC%))
        '
        TEXT21(0).TEXT = Str$(CPAGCLI%(NC%))
        TEXT21(1).TEXT = ""
        '
        Text10 = CONTROL$("OCOMPRA", "LENTREGA")
        '
        Text15 = TRIM$(Str$(MONEVACO%(NC%)))
        If MONEMI% < 1 Then
          Label11.ForeColor = RGB(0, 0, 255)
          Label11 = "Pesos"
          MONEMI% = 1
          Label11.ToolTipText = "Doble-Click para 'Dólares'"
        End If
        '
        If MONEMICO%(NC%) = 2 Then
          Label11.ForeColor = RGB(255, 0, 0)
          Label11 = "Dólares"
          MONEMI% = 2
          Label11.ToolTipText = "Doble-Click para 'Pesos'"
        End If
        '
        Call ABRECOMPRAS
        SQLX$ = "SELECT Nume_Ocom, Femi_Ocom, AtSr_Ocom, Pdes_Ocom FROM OcomEnca "
        SQLX$ = SQLX$ + "WHERE Npro_Ocom = " & Abs(NC%) & " "
        SQLX$ = SQLX$ + "ORDER BY Femi_Ocom DESC;"
        'Set OEncaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
        Dim OENCATEMP As ADODB.Recordset
        Set OENCATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        
        On Error Resume Next
        With OENCATEMP
          .MoveFirst
          If Err < 1 Then
            ATEN$ = !AtSr_Ocom: Text6 = ATEN$
            PODESC$ = !Pdes_Ocom: PORDESCU = PODESC$
          End If
        End With
        'CARGO ULTIMA IMP.CONTABLE
        If NC% < 0 Then
          NCXX1% = NC% * (-1)
        End If
        Call ABRECOMPRAS
        SQLX$ = " SELECT IContab  FROM OCOMDETA "
        SQLX$ = SQLX$ + " WHERE Npro_Ocom = " & NCXX1% & " "
        SQLX$ = SQLX$ + " AND Stat_Ocom < 9 "
        SQLX$ = SQLX$ + " AND IContab <> 0 "
        SQLX$ = SQLX$ + " ORDER BY Femi_Ocom DESC "
        'Set ICONTMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
        Dim ICONTMP As ADODB.Recordset
        Set ICONTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        
        With ICONTMP
            On Error Resume Next
            .MoveFirst
            If Err Then
              On Error GoTo 0
              GoTo 20
            End If
            On Error GoTo 0
            Text7 = CODCUE(!IContab)
        End With
        'CARGO ULTIMO SECTOR
20      Call ABRECOMPRAS
        SQLX$ = " SELECT Sect_Ocom  FROM OCOMDETA "
        SQLX$ = SQLX$ + " WHERE Npro_Ocom = " & NCXX1% & " "
        SQLX$ = SQLX$ + " AND Stat_Ocom < 9 "
        SQLX$ = SQLX$ + " AND Sect_Ocom <> ' ' "
        SQLX$ = SQLX$ + " ORDER BY Femi_Ocom DESC "
        'Set SECTORTMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
        Dim SECTORTMP As ADODB.Recordset
        Set SECTORTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        
        With SECTORTMP
            On Error Resume Next
            .MoveFirst
            If Err Then
              On Error GoTo 0
              GoTo 99
            End If
            On Error GoTo 0
            Text18 = !Sect_Ocom
        End With
        '
      End If
    End If
    
99  Screen.MousePointer = 1
    '
End Sub
   '
Function PRONUOCO&()
   '
   PRONU& = 1
   Call ABRECOMPRAS
   '
   SQLX$ = "SELECT Nume_Ocom FROM OcomEnca "
   SQLX$ = SQLX$ + "ORDER BY Nume_Ocom DESC;"
   'Set OENCATEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim OENCATEMP As ADODB.Recordset
   Set OENCATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   
   On Error Resume Next
   With OENCATEMP
     .MoveFirst
     If Err < 1 Then
       PRONU& = 1 + !Nume_Ocom
     End If
   End With
   On Error GoTo 0
   '
   PRONUOCO& = PRONU&
   OENCATEMP.Close
   Set OENCATEMP = Nothing
   '
End Function

Private Sub Text6_GotFocus()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
   '
   On Error Resume Next
   Text6.SelStart = 0
   Text6.SelLength = Len(Text6.TEXT)
   On Error GoTo 0
End Sub

Private Sub Text7_DblClick()
   Call Command18_Click
End Sub

Private Sub Text8_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text9_GotFocus()
   On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
End Sub

Private Sub Timer1_Timer()
   '
   'Call CALTOCOM
   '
End Sub

Private Sub Total_Change(Index As Integer)
   If Index = 0 Then
     Call CALTOCOM
'      DESC$ = PORDESCU.TEXT
'      TTT1 = Val(Total(0).TEXT) * COEFDESCU(DESC$) / 100
'      Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
'      TTT3 = Val(Total(0).TEXT) - Val(Total(1).TEXT)
'      Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
   End If
'   If Index < 3 Then
'      TTT1 = Val(Total(2).TEXT)
'      TTT2 = Val(TOTIVA(3).TEXT)
'      Total(3).TEXT = Str$(TTT1 + TTT2)
'   End If
End Sub

Private Sub Total_DblClick(Index As Integer)
   For KKI% = 0 To 4
     Total(KKI%).Enabled = Not (Total(KKI%).Enabled)
   Next KKI%
End Sub

'
Private Sub Total_GotFocus(Index As Integer)
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Sub GRACOMPRA(OKX%)
   '
   OKX% = 0
   If ULTREG&("!OCOMDETA") < 1 Then Exit Sub
   '
   COPRECIO = 1
   If Total(0).Enabled = False Then COPRECIO = 0
   '
   Call CIERRARCH("OCOMENCA")
   Call CIERRARCH("OCOMDETA")
   '
   Call BLOQARCH("OCOMENCA")
   Call BLOQARCH("OCOMDETA")
   Call BLOQARCH("MACONSIG")
   '
   NPP& = PRONUOCO&
   If NPP& > Val(Text11.TEXT) Then
      Call CIERRARCH("OCOMENCA")
      Call CIERRARCH("OCOMDETA")
      Call CIERRARCH("MACONSIG")
      Call COMUNI("FLEXOFT ha Re-Numerado la O/Compra\como O/C Número '" + TRIM$(Str$(NPP&)) + "'.\  \Verifique y Vuelva a Grabar.")
      Text11.TEXT = TRIM$(Str$(NPP&))
      Exit Sub
   End If
   '
   MONECO% = Val(Text15)
   If MONECO% < 1 Then MONECO% = Val(TRIM$(CONTROL$("OCOMPRA", "MONECOS")))
   If MONECO% < 1 Then MONECO% = 1
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Proveedor no Válido"
      GoTo 98
   End If
   NPRO% = Val(Text1.TEXT)
   '
   If Val(TEXT21(0).TEXT) > 255 Then
      MERRO$ = "Condición de Pago no Válida"
      GoTo 98
   End If
   '
   If Val(TEXT21(1).TEXT) > 255 Then
      MERRO$ = "Embalaje no Válido"
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
123 If COPRECIO < 0.99 Then
      PORDESCU.TEXT = ""
      Text17 = ""
      For KK% = 0 To 4
        Total(KK%).TEXT = ""
      Next KK%
    End If
    '
    Call ABRECOMPRAS
    With OCOMENCA
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
         NRO& = Val(Text11)
       !Nume_Ocom = NRO&
    'Call REPLA(ENCAO$, MKS$(NRO&), 1, 4)
         FEX = FEMI%
         REFE$ = FECHATEX$(FEX) + " - " + RASOCLI$((-1) * NPRO%)
       !Refe_Ocom = REFE$
    'Call REPLA(ENCAO$, REFE$, 5, 44)
       !Femi_Ocom = CVDAT(FEMI%)
       !Npro_Ocom = NPRO%
       !RASO_PROV = RASOCLI$((-1) * NPRO%)
       !TSum_Ocom = 0               ' materiales
          STATUOC% = (-1)
          If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) = "DIRECT" Then STATUOC% = 0
       !Stat_Ocom = STATUOC%              ' Status Pendiente de Aprobación
       
       !Fgen_Ocom = Now
       !Usua_Gen = USNBR%
         '
         MONECOI% = Val(Text15)
       !Mone_Cos = MONECOI%        ' antes pos 83,1
       !Mone_Emis = MONEMI%        ' antes pos 84,1
         TICAMBX = 1
         If MONEMI% <= 1 Then
           If MONECOI% > 1 Then
             TICAMBX = TICAMONE(MONECOI%)
           End If
         End If
       !Tcam_Ocom = TICAMBX       ' antes pos 85, 4
         '
         CPAG% = Val(TEXT21(0))
       !Cpag_Ocom = CPAG%
         EMBAL% = Val(TEXT21(1))
       On Error Resume Next    ' POR SI NO ESTA DEFINIDO
       !Embala_Ocom = EMBAL%
       On Error GoTo 0
       !AtSr_Ocom = TRIM$(Text6)
         '
         PORDESCUENTO$ = TRIM$(PORDESCU)
         DESC$ = PORDESCUENTO$
         TOTANE# = Val(Total(0)) * COPRECIO
         IDESCUEN# = Val(Total(1)) * COPRECIO
         TOTANEDES# = TOTANE# - IDESCUEN#
         TOTIVA# = Val(Total(2)) * COPRECIO
         TOTPER# = Val(Total(4)) * COPRECIO
         TOTATO# = TOTANEDES# + TOTIVA# + TOTPER
         '
       !ToBru_Ocom = TOTANE#
       !Pdes_Ocom = PORDESCU
       !IDes_Ocom = IDESCUEN#
       !Iiva_Ocom = TOTIVA#
       !Iotr1_Ocom = TOTPER#
       !Total_Ocom = TOTATO#
       !Lentrega = Text10
       !Obse_Ocom = Text13
       On Error Resume Next
       !Anex_Ocom = RAI$
       On Error GoTo 0
       '
       .Update
    End With
    '
    EMICONSI% = 0
    IDENOC$ = MKS$(NRO&) + MKI$(FEMI%) + MKI$(NPRO%)
    '
    ORDIT% = 0
    APRODIR% = 1 'Aprobación Directa en Gestion de Mantenimiento
    If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) = "DIRECT" Then
      APRODIR% = 1
    End If
    ARCHILIB$ = "!1OCOMDET"
    For J& = 1 To ULTREG&("!OCOMDETA")
      '
      ZZ$ = REGLEIDO$("!OCOMDETA", J&)
      If COPRECIO < 0.99 Then
        Call REPLA(ZZ$, String$(28, 0), 67, 28)  ' BORRA PRECIO UNITARIO E IMPORTE
        Call GRAREG("!OCOMDETA", ZZ$, J&)
      End If
      '
      ZZ$ = REGLEIDO$("!OCOMDETA", J&)
      If TRIM$(Mid$(ZZ$, 105, 12)) = "" Then Call REPLA(ZZ$, Text7, 105, 12)
      If TRIM$(Mid$(ZZ$, 15, 12)) = "" Then Call REPLA(ZZ$, Text18, 15, 12)
      Call GRAREG("!OCOMDETA", ZZ$, J&)
      '
      CI0% = CVI(Mid$(ZZ$, 9, 2))
      COEFUNI = CVS(Mid$(ZZ$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
      PREUNPE# = CVD(Mid$(ZZ$, 67, 8)) * TICAMONE(MONEMI%) / COEFUNI
      PREUNCO# = PREUNPE# / TICAMONE(MONECOI%)
      Monecos$ = "$ ": If MONECOI% > 1 Then Monecos$ = Left$(Text16, 2)
      VAUNI$ = Monecos$ + Mid$(FORMATNUM$(PREUNCO#, "F12.4"), 3)
      '
      ORDIT% = ORDIT% + 1
      With OCOMDETA
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
         !Nume_Ocom = NRO&
         !NUORD_ITEM = ORDIT%
         !Femi_Ocom = CVDAT(FEMI%)
         !Npro_Ocom = NPRO%
         !RASO_PROV = RASOCLI$((-1) * NPRO%)
         !Cod_Inte = CI0%
         !Cod_Exte = CODEXT$(CI0%)
         !DESCRITEM = DESCRIT0$(CI0%)
         !Delibre = ""
         If COLTELIB% >= 0 Then
           !Delibre = TRIM$(REGLEIDO$(ARCHILIB$, J&))
         End If
         If CI0% < 1 Then
           If TRIM$(!Delibre) <> "" Then
             !DESCRITEM = Left$(!Delibre, 64)
           End If
         End If
         !Valo_Unid = VAUNI$
         !Uni_Mas = Mid$(ZZ$, 11, 4)
           CAUNIS = CVS(Mid$(ZZ$, 51, 4))  ' CANT. EN UNIDADES DE STOCK
         !Unids_Ocom = CAUNIS
           CANTU = CVS(Mid$(ZZ$, 55, 4))   ' CANT EN UNIDADES DE COMPRA
         !Cant_Ocom = CANTU
         !Uni_Com = Mid$(ZZ$, 59, 4) ' unidad de compra
         !Coef_Unids = COEFUNI
         !Pre_Unit = CVD(Mid$(ZZ$, 67, 8))
         !Desc_Item = CVS(Mid$(ZZ$, 75, 4))
         !PrUn_Neto = CVD(Mid$(ZZ$, 79, 8))
         !INet_Item = CVD(Mid$(ZZ$, 87, 8))
         !FENTRE_SOL = CVDAT(CVI(Mid$(ZZ$, 95, 2)))
         !CONSIMAT = TRIM$(Mid$(ZZ$, 97, 2))
         !Cant_Rec = 0
         !Cod_Oper = 0
         !Doc_Orig = ""
         !N_OrServ = 0
         !Mone_Cos = MONECOI%        ' antes pos 83,1
         !Mone_Emis = MONEMI%        ' antes pos 84,1
         !TSum_Ocom = 0              ' materiales
         !Stat_Ocom = APRODIR% - 1  ' status Pendiente de aprobacion
         '
         !IContab = CUECOINT(TRIM$(Mid$(ZZ$, 105, 12))) ' Imputacion contable
         !Sect_Ocom = TRIM$(Mid$(ZZ$, 15, 12))
        .Update
      End With
      '
      NOCOM$ = TRIM$(Str$(NRO&))
        While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
      ITNU$ = TRIM$(Str$(ORDIT%))
        While Len(ITNU$) < 2: ITNU$ = "0" + ITNU$: Wend
      NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-" + ITNU$
      Call LEEEXPLO(CODEXT$(CI0%), CAUNIS)
      '
      For I% = 1 To CAIT%
        CII% = CIJ%(I%)
        CANTI# = CDbl(CAUNIS * USOI(I%))
        '
        If TRIM$(CODEXT$(CII%)) <> "" Then
          If CANTI# > 0.05 Then
          Call ABRESTOCKS
            With MaConsig
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next
              !CodiEmpr = CodiEmp%
              On Error GoTo 0
               !Femi_Ocom = CVDAT(FEMI%)
               !Npro_Ocom = NPRO%
               !RASO_PROV = RASOCLI$((-1) * NPRO%)
               !Nro_Ocom = NOCOM$
                '
               !Cod_Inte = CII%
               !Cod_Exte = CODEXT$(CII%)
               !DESCRITEM = DESCRIT0$(CII%)
               !Unidad = UNIMED$(CII%)
               !Cant_Asig = CANTI#
              .Update
              EMICONSI% = 1
            End With
          End If
        End If
      Next I%
      '
   Next J&
   Call BAJALDISCO
   '
   If COPRECIO < 0.99 Then
     PORDESCU.TEXT = ""
     Text17 = ""
     For KK% = 0 To 4
       Total(KK%).TEXT = ""
     Next KK%
   End If
   '
   ' actualizar costo standard y kilos por pieza
   ACOSREP$ = UCase$(CONTROL$("OCOMPRA", "ACOSREP"))
   '
   HOII% = HOY(HOS$)
   For J& = 1 To ULTREG&("!OCOMDETA")
     ZZ$ = REGLEIDO$("!OCOMDETA", J&)
     NUORSE& = CVS(Mid$(ZZ$, 117, 4))
     If NUORSE& < 1 Then   ' EXCLUYE ORDEN DE SERVICIO
        CI0% = CVI(Mid$(ZZ$, 9, 2))
        COEFUNI = CVS(Mid$(ZZ$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
        COSPESOSD# = CVD(Mid$(ZZ$, 79, 8)) * TICAMONE(MONEMI%) / COEFUNI
        '
        'REMAL& = CI0% + 1
        'REMAS$ = REGLEIDO$("MASTER", REMAL&)
        PRHBL% = PROVHABI%(CI0%)
        If PRHBL% < 1 Or PRHBL% = NPRO% Or ACOSREP$ = "SIEMPRE" Then
            COSMONCO = COSPESOSD# / TICAMONE(MONECO%)
            If COSMONCO >= 0.000005 Or PRHBL% < 1 Then  ' PARA QUE ACTUALICE P.HABIT AUN SI COSTO REPOS. ES CERO
              If ACOSREP$ <> "NUNCA" Then
                Call ABREMASTER
                Master.FindFirst ("CODINT = " & CI0%)
                If Master.NoMatch = False Then
                  Master.Edit
                  If COSMONCO >= 0.000005 Then
                    Master!COSUNI = COSMONCO
                    Master!FECOSTO = CVDAT(FEMI%)
                  End If
                  If PRHBL% < 1 Then
                    Master!PROVHABI = NPRO%
                  End If
                  Master!Monecos = MONECO%
                  Master.Update
                End If
              End If
            End If
            '
          Else
            '
            If PRHBL% < 1 Then GoTo 164 ' no hay proveedor habitual anterior
            '
            TTXX$ = "Para el Artículo " + TRIM$(CODEXT$(CI0%)) + "\el Proveedor Habitual Registrado es\"
            TTXX$ = TTXX$ + TRIM$(RASOCLI$((-1) * PRHBL%)) + "\  \"
            TTXX$ = TTXX$ + "Desea Cambiar el nuevo Proveedor Habitual a\"
            TTXX$ = TTXX$ + RASOCLI$((-1) * NPRO%)
            If COMALTER%(TTXX$ + "\\No, Conservar Anterior\Si, Cambiar") = 2 Then
              '
164           Call ABREMASTER
              Master.FindFirst "CODINT =" & CI0%
              If Master.NoMatch = False Then
                Master.Edit
                Master!PROVHABI = NPRO%
                COSMONCO = COSPESOSD# / TICAMONE(MONECO%)
                If COSMONCO >= 0.000005 Then
                  Master!COSUNI = COSMONCO
                  Master!FECOSTO = CVDAT(FEMI%)
                  Master!Monecos = MONECO%
                End If
                Master.Update
              End If
              '
              ZZX$ = REGBLAN$("RECAMPH")
              Call REPLA(ZZX$, MKI$(CI0%), 1, 2)
              HOII% = HOY(HOS$)
              Call REPLA(ZZX$, MKI$(HOII%), 3, 2)
              Call REPLA(ZZX$, MKI$(PRHBL%), 5, 2)
              Call REPLA(ZZX$, MKI$(NPRO%), 7, 2)
              Call REPLA(ZZX$, USERTER$, 9, 24)
              Call REPLA(ZZX$, "Emis.O/C " + TRIM$(Str$(NRO&)), 33, 32)
              Call REGAPP("RECAMPH", ZZX$)
              Call CIERRARCH("RECAMPH")
              GoTo 169
              '
            End If
            '
            CUNIREG# = COSUNI(CI0%) * TICAMONE(MONECOSTO(CI0%))
            If CUNIREG# >= 0.00005 Then
              If COSPESOSD# > CUNIREG# + 0.00005 Then
                Call COMUNI("ATENCION: Precio de Compra\es Superior al Costo Standard\en Código '" + TRIM$(CODEXT$(CI0%)) + "' !!!\  \Informe al Supervisor de Compras.")
              End If
            End If
169     End If
        '
        If PESUNI(CI0%) < 0.000005 Then
           UNIST$ = UCase$(TRIM$(Mid$(ZZ$, 11, 4)))
           If Left$(UNIST$, 1) = "U" Then
             UNICO$ = UCase$(TRIM$(Mid$(ZZ$, 59, 4)))
             If UNICO$ = "KG" Then
               On Error Resume Next
               PSXX = 1 / COEFUNI
               On Error GoTo 0
               Call ABREMASTER
               With Master
                 .FindFirst "CodInt = " & CI0%
                 If .NoMatch = False Then
                   .Edit
                   !PESUNI = PSXX
                   .Update
                 End If
               End With
             End If
           End If
        End If
        '
      End If
    Next J&
    '
    '
    'depurar PLACOMP
    For J& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", J&)
      NUORSE& = CVS(Mid$(ZZ$, 117, 4))
      If NUORSE& < 1 Then   ' EXCLUYE ORDEN DE SERVICIO
        CI0% = CVI(Mid$(ZZ$, 9, 2))
        COEFUNI = CVS(Mid$(ZZ$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
        CANOCOM = CVS(Mid$(ZZ$, 51, 4))
        FENTRESO% = CVI(Mid$(ZZ$, 95, 2))
        If FENTRESO% < FEMI% Then FENTRESO% = FEMI%
        '
        RFF$ = RECFILT$("PLACOMP", 1, MKI$(CI0%))
        For UJ& = 1 To Len(RFF$) Step 4
          RPLA& = CVS(Mid$(RFF$, UJ&, 4))
          XX$ = REGLEIDO$("PLACOMP", RPLA&)
          FENTREPLA% = CVI(Mid$(XX$, 3, 2))
            If FENTREPLA% < FEMI% Then FENTREPLA% = FEMI%
          CANPLA = CVS(Mid$(XX$, 21, 4))
          '
          If FENTREPLA% >= FENTRESO% Then
            If CANPLA <= CANOCOM Then
                CANOCOM = CANOCOM - CANPLA
                CANPLA = 0
              Else
                CANPLA = CANPLA - CANOCOM
                CANOCOM = 0
            End If
            Call REPLA(XX$, MKS$(CANPLA), 21, 4)
            ' depurar registro de plan de compras
            If CANPLA < 0.005 Then XX$ = String$(128, 0)
            Call GRAREG("PLACOMP", XX$, RPLA&)
          End If
        Next UJ&
      End If
    Next J&
    Call CIERRARCH("PLACOMP")
    '
    JJ& = 0
    For U& = 193 To 512 Step 64
      TTXX$ = Mid$(ENCAO$, U&, 64)
      JJ& = JJ& + 1
      Call GRAREG("!OBSERVOF", TTXX$, JJ&)
    Next U&
    Call SETULTREG("!OBSERVOF", JJ&)
    Call CIERRARCH("!OBSERVOF")
    '
    CPAG% = CVI(Mid$(ENCAO$, 153, 2))
    FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    '
    Screen.MousePointer = 1
    OKX% = 1
    OCOMENCA.Close
    OCOMDETA.Close
    Exit Sub
    '
98  OCOMENCA.Close
    OCOMDETA.Close
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    Call MENSERR(24, MERRO$)
    '
End Sub

Private Sub Total_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Index = 0 Then
       On Error Resume Next
       PORDESCU.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Total_LostFocus(Index As Integer)
   If Index = 0 Then
     TTT1 = Val(Total(0).TEXT)
     Total(0) = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   End If
End Sub
'\\\OT-05-0688-WAR-20/10/05///
Sub CARDESCOM()
   '
   Call CIERRARCH("!OCOMDETA")
   List2.ListIndex = (-1)
   HOII% = HOY(HOS$)
   '
10 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Detalle de Orden de Compra en " + TRIM$(Label11)
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(140 + Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   '
   Timer1.Enabled = True
   PRF$ = "CARDESCO"
   XXX% = VENTABU%(PRF$ + ATRI$)
   List1.Clear
   Call CALTOCOM
   Timer1.Enabled = False
   '
   If XXX% < 0 Or ULTREG&("!OCOMDETA") < 1 Then
      Call CIERRARCH("!OCOMDETA")
      Call BLANARCH("!OCOMDETA")
      Call BLANFORMU
      Total(0).TEXT = ""
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   'ATK% = 0
   'DESACTIVE LA BANDERA DE CAMBIOS XQ SIEMPRE CAMBIA
   'YA QUE SE ESCRIBE EL REGISTRO DE UNIDAD/C X EJEMPLO.
   NPRO% = Val(Text1)
   For U& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", U&)
      'Z0$ = ZZ$
      CI0% = CVI(Mid$(ZZ$, 9, 2))
      If TRIM$(Mid$(ZZ$, 11, 4)) = "" Then Call REPLA(ZZ$, UNIMED$(CI0%), 11, 4)
      Call REPLA(ZZ$, UNIMED$(CI0%), 59, 4)
      Call REPLA(ZZ$, MKS$(1), 63, 4)
      '
      CANSTOK = CVS(Mid$(ZZ$, 51, 4))
      CANCOMP = CANSTOK
      Call REPLA(ZZ$, MKS$(CANCOMP), 55, 4)
      Call REPLA(ZZ$, MKS$(COEFI), 63, 4)
      '
      PreUnid# = CVD(Mid$(ZZ$, 67, 8))
      If PreUnid# < 0.00005 Then
        If UC$ = "" Then UC$ = ULTICOM$(NPRO%, CI0%)
        MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
        PreUnid# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
        PORDESCUEI = CVS(Mid$(UC$, 75, 4))
        If PreUnid# < 0.00005 Or MODOULPRE% = 2 Then
          PreUnid# = COSUNI(CI0%) * TICAMONE(MONECOSTO%(CI0%)) * COEFI
          PORDESCUEI = 0
        End If
        PreUnid# = PreUnid# / TICAMONE(MONEMI%)
        Call REPLA(ZZ$, MKD$(PreUnid#), 67, 8)
        Call REPLA(ZZ$, MKS$(PORDESCUEI), 75, 4)
      End If
      '
      If TRIM$(Mid$(ZZ$, 105, 12)) = "" Then Call REPLA(ZZ$, Text7, 105, 12)
      If TRIM$(Mid$(ZZ$, 15, 12)) = "" Then Call REPLA(ZZ$, Text18, 15, 12)
    '  If ZZ$ <> Z0$ Then ATK% = 1
      Call GRAREG("!OCOMDETA", ZZ$, U&)
   Next U&
   '
   Screen.MousePointer = 1
   'If ATK% = 1 Then GoTo 10
   '
   Call CARLISCO
   Call CALTOCOM
   Screen.MousePointer = 1
   '
  '
End Sub
'\\\OT-05-0688-WAR-FIN///
Sub CARDETCOM()
   '
   '\\\OT-05-0688-WAR-20/10/05///
   If CONTROL$("OCOMPRA", "MODOIMPU") = "DES" Then
     Call CARDESCOM
     Screen.MousePointer = 1
     Exit Sub
   End If
   '\\\OT-05-0688-WAR-FIN///
   '
   Call CIERRARCH("!OCOMDETA")
   List2.ListIndex = (-1)
   HOII% = HOY(HOS$)
   '
10 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Detalle de Orden de Compra en " + TRIM$(Label11)
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(140 + Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   '
   Timer1.Enabled = True
   PRF$ = "CADECOMA"
   XXX% = VENTABU%(PRF$ + ATRI$)
   List1.Clear
   Call CALTOCOM
   Timer1.Enabled = False
   '
   If XXX% < 0 Or ULTREG&("!OCOMDETA") < 1 Then
      Call CIERRARCH("!OCOMDETA")
      Call BLANARCH("!OCOMDETA")
      Call BLANFORMU
      Total(0).TEXT = ""
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   ATK% = 0
   NPRO% = Val(Text1)
   For U& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", U&)
      Z0$ = ZZ$
      CI0% = CVI(Mid$(ZZ$, 9, 2))
      '
      If CI0% <= 0 Then GoTo 15 'SOLO CODIFICADO
      '
      If TRIM$(Mid$(ZZ$, 11, 4)) = "" Then Call REPLA(ZZ$, UNIMED$(CI0%), 11, 4)
      If TRIM$(Mid$(ZZ$, 59, 4)) = "" Then Call REPLA(ZZ$, UNICOM$(CI0%, NPRO%), 59, 4)
      If TRIM$(Mid$(ZZ$, 11, 4)) = TRIM$(Mid$(ZZ$, 59, 4)) Then Call REPLA(ZZ$, MKS$(1), 63, 4)
      '
      UC$ = ""
      COEFI = CVS(Mid$(ZZ$, 63, 4))
      CANCOMP = CVS(Mid$(ZZ$, 55, 4))
      CANSTOK = CVS(Mid$(ZZ$, 51, 4))
      If COEFI < 0.0001 Then
        If CANCOMP > 0.05 Then
          If CANSTOK > 0.05 Then
            COEFI = CANSTOK / CANCOMP
          End If
        End If
      End If
      '
      If COEFI < 0.0001 Then
        UC$ = ULTICOM$(NPRO%, CI0%)
        If TRIM$(Mid$(ZZ$, 11, 4)) = TRIM$(Mid$(UC$, 11, 4)) Then
          If TRIM$(Mid$(ZZ$, 59, 4)) = TRIM$(Mid$(UC$, 59, 4)) Then
            COEFI = CVS(Mid$(UC$, 63, 4))
          End If
        End If
      End If
      If COEFI < 0.0001 Then COEFI = 1
      '
      If CANCOMP < 0.05 Then
          CANCOMP = CANSTOK / COEFI
        Else
          CANSTOK = CANCOMP * COEFI
      End If
      CANCOMP = CANSTOK / COEFI
      Call REPLA(ZZ$, MKS$(CANSTOK), 51, 4)
      Call REPLA(ZZ$, MKS$(CANCOMP), 55, 4)
      Call REPLA(ZZ$, MKS$(COEFI), 63, 4)
      '
      PreUnid# = CVD(Mid$(ZZ$, 67, 8))
      If PreUnid# < 0.00005 Then
        If UC$ = "" Then UC$ = ULTICOM$(NPRO%, CI0%)
        MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
        PreUnid# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
        PORDESCUEI = CVS(Mid$(UC$, 75, 4))
        If PreUnid# < 0.00005 Or MODOULPRE% = 2 Then
          PreUnid# = COSUNI(CI0%) * TICAMONE(MONECOSTO%(CI0%)) * COEFI
          PORDESCUEI = 0
        End If
        PreUnid# = PreUnid# / TICAMONE(MONEMI%)
        Call REPLA(ZZ$, MKD$(PreUnid#), 67, 8)
        Call REPLA(ZZ$, MKS$(PORDESCUEI), 75, 4)
      End If
      '
      If ZZ$ <> Z0$ Then ATK% = 1
      Call REPLA(ZZ$, Text7, 105, 12)
      Call REPLA(ZZ$, Text18, 15, 12)
      Call GRAREG("!OCOMDETA", ZZ$, U&)
15 Next U&
   '
   Screen.MousePointer = 1
   If ATK% = 1 Then GoTo 10
   '
   Call CARLISCO
   Call CALTOCOM
   Screen.MousePointer = 1
   '
End Sub

Sub CALTOCOM()
   '
   NC1% = (-1) * Val(Text1)
   TOTANE# = 0
   For J& = 1 To ULTREG&("!OCOMDETA")
       X$ = REGLEIDO$("!OCOMDETA", J&)
       CI0% = CVI(Mid$(X$, 9, 2))
       '
       If CI0% <= 0 Then GoTo 10 'SOLO MAT CODIFICADO
       '
       UNIMAS$ = Mid$(X$, 11, 4)
       CANTU = CVS(Mid$(X$, 55, 4))
       PreUnid# = CVD(Mid$(X$, 67, 8))
       PORDE = CVS(Mid$(X$, 75, 4))
       FENTRE% = CVI(Mid$(X$, 95, 2))
       '
       PREUNINE# = (Int(PreUnid# * 100 * (100 - PORDE) + 0.5)) / 10000
       INETITEM# = CDbl((Int(100 * PREUNINE# * CANTU + 0.5)) / 100)
       TOTANE# = TOTANE# + INETITEM#
       '
       Call REPLA(X$, MKD$(PREUNINE#), 79, 8)
       Call REPLA(X$, MKD$(INETITEM#), 87, 8)
       Call GRAREG("!OCOMDETA", X$, J&)
       '
       UNIST$ = UCase$(TRIM$(Mid$(X$, 11, 4)))
       UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
       CANSTXX = CVS(Mid$(X$, 51, 4))
       '
       If CANTU > 0.05 Then
         COEFI = CANSTXX / CANTU
         Call REPLA(X$, MKS$(COEFI), 63, 4)
         Call GRAREG("!OCOMDETA", X$, J&)
       End If
       '
10 Next J&
   '
   Call CIERRARCH("!OCOMDETA")
   '
' presentar registro de totales y condicion de pago
   '
   If MONEMI% > 1 Then
     MONECOI% = MONEMI%
     TICAMBX = TICAEMI
   End If
   '
   PORDESCUENTO$ = TRIM$(PORDESCU)
   DESC$ = PORDESCUENTO$
   IDESCUEN# = TOTANE# * COEFDESCU(DESC$) / 100
   TOTANEDES# = TOTANE# - IDESCUEN#
   TOTIVA# = 0
   ALIVX = Val(Text12)
   If ALIVX < 0 Or ALIVX > 99 Then
      Text12 = TRIM$(Str$(ALIVA))
      ALIVX = ALIVA
   End If
   If Abs(NC1%) > 0 Then
     PVCL% = PIVACLI%(NC1%)
     If PVCL% <> 2 Then
       If PVCL% <> 4 Then
         If PVCL% <> 5 Then
           TOTIVA# = (Int(ALIVX * TOTANEDES# + 0.5)) / 100
         End If
       End If
     End If
   End If
   '
   TOTPER# = 0
   PERCX = Val(Text17)
   If PERCX < 0 Or PERCX > 99 Then
      Text17 = ""
      PERCX = 0
   End If
   If Abs(NC1%) > 0 Then
     PVCL% = PIVACLI%(NC1%)
     If PVCL% <> 2 Then
       If PVCL% <> 4 Then
         If PVCL% <> 5 Then
           TOTPER# = (Int(PERCX * TOTANEDES# + 0.5)) / 100
         End If
       End If
     End If
   End If
   '
   
   TOTATO# = TOTANEDES# + TOTIVA# + TOTPER
   '
   Total(0) = FORMATNUM$(TOTANE#, "F11.2")
   Total(1) = FORMATNUM$(IDESCUEN#, "F11.2")
   Total(2) = FORMATNUM$(TOTIVA#, "F11.2")
   Total(4) = FORMATNUM$(TOTPER#, "F11.2")
   Total(3) = FORMATNUM$(TOTATO#, "F11.2")
   '
End Sub

Sub CARCOPEN()
   '
   Screen.MousePointer = 11
   NPRO% = Val(Text1)
   HO% = HOY(HOS$)
   '
   JJ& = 0
   REBLA$ = REGBLAN$("OCOMDETA")
   Call SETULTREG("!OCOMDETA", 0)
   ARCHILIB$ = "!1OCOMDET"
   Call SETULTREG(ARCHILIB$, 0)
   '
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   PCI$ = "": UFECHA = HOY(HOS$)
   RFF$ = RECFILT$("PLACOMP", 8, MKI$(NPRO%))
   If RFF$ = "" Then Exit Sub
   '
   For KKI% = 1 To Len(RFF$) Step 4
      REPX& = CVS(Mid$(RFF$, KKI%, 4))
      XX$ = REGLEIDO$("PLACOMP", REPX&)
      CI0% = CVI(Left$(XX$, 2))
      MCI0$ = MKI$(CI0%)
      FENTREIT% = CVI(Mid$(XX$, 3, 2))
      If FENTREIT% > UFECHA Then UFECHA = FENTREIT%
      For UKI% = 1 To Len(PCI$) Step 2
        If Mid$(PCI$, UKI%, 2) = MCI0$ Then GoTo 9
      Next UKI%
      PCI$ = PCI$ + MCI0$
9  Next KKI%
   '
   FORSELMAT.LIMATSEL.Clear
   For UKI% = 1 To Len(PCI$) Step 2
     CI0% = CVI(Mid$(PCI$, UKI%, 2))
     TTXX$ = AJUSTI$(CODEXT$(CI0%), 16) + DESCRIT0$(CI0%)
     FORSELMAT.LIMATSEL.AddItem TTXX$
   Next UKI%
   FORSELMAT.Caption = "Materiales a Comprar - " + Text2
   FORSELMAT.Label2 = FECHATEX$(UFECHA)
   For U& = 1 To FORSELMAT.LIMATSEL.ListCount
     FORSELMAT.LIMATSEL.Selected(U& - 1) = True
   Next U&
   '
   Screen.MousePointer = 1
   On Error Resume Next
   FORSELMAT.BOTREC.SetFocus
   FORSELMAT.LIMATSEL.TopIndex = 0
   On Error GoTo 0
   FORSELMAT.Show 1
   '
   DoEvents
   If TRIM$(FORSELMAT.Label2 = "") Then Exit Sub
   '
   Screen.MousePointer = 11
   PCI$ = "":
   UFECHAI% = FECHANUM(FORSELMAT.Label2)
   For U& = 1 To FORSELMAT.LIMATSEL.ListCount
     If FORSELMAT.LIMATSEL.Selected(U& - 1) = True Then
       PCI$ = PCI$ + MKI$(CODINT%(TRIM$(Left$(FORSELMAT.LIMATSEL.List(U& - 1), 15))))
     End If
   Next U&
   '
   RFF$ = RECFILT$("PLACOMP", 8, MKI$(NPRO%))
   For KKI% = 1 To Len(RFF$) Step 4
      REPX& = CVS(Mid$(RFF$, KKI%, 4))
      XX$ = REGLEIDO$("PLACOMP", REPX&)
      CI0% = CVI(Left$(XX$, 2))
      MCI0$ = MKI$(CI0%)
      For UKI% = 1 To Len(PCI$) Step 2
        If Mid$(PCI$, UKI%, 2) = MCI0$ Then GoTo 24
      Next UKI%
      GoTo 29
      
24    FENTREIT% = CVI(Mid$(XX$, 3, 2))
      If FENTREIT% < HO% Then FENTREIT% = HO%
      If FENTREIT% <= UFECHAI% Then
        '
        CANTU = CVS(Mid$(XX$, 21, 4))
        '
        For KKJ& = 1 To JJ&
          ZZ$ = REGLEIDO$("!OCOMDETA", KKJ&)
          If CVI(Mid$(ZZ$, 9, 2)) = CI0% Then
            If CVI(Mid$(ZZ$, 95, 2)) = FENTREIT% Then
              CANTU = CANTU + CVS(Mid$(ZZ$, 51, 4))
              Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
              Call GRAREG("!OCOMDETA", ZZ$, KKJ&)
              GoTo 29
            End If
          End If
        Next KKJ&
        '
        ZZ$ = REBLA$
        Call REPLA(ZZ$, MKI$(CI0%), 9, 2)
        Call REPLA(ZZ$, UNIMED$(CI0%), 11, 4)
        Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
        Call REPLA(ZZ$, UNIMED$(CI0%), 59, 4)
        Call REPLA(ZZ$, MKS$(1), 63, 4)
        Call REPLA(ZZ$, MKI$(FENTREIT%), 95, 2)
        '
        UC$ = ULTICOM$(NPRO%, CI0%)
        MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
        PreUnid# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
        COEFIX = CVS(Mid$(UC$, 63, 4)): If COEFIX < 0.0005 Then COEFIX = 1
        If PreUnid# < 0.00005 Or MODOULPRE% = 2 Then
          PreUnid# = COSUNI(CI0%) * TICAMONE(MONECOSTO%(CI0%)) * COEFIX
        End If
        PreUnid# = PreUnid# / TICAMONE(MONEMI%)
        '
        UNICOMPRA$ = TRIM$(Mid$(UC$, 59, 4))
        If UNICOMPRA$ <> "" Then Call REPLA(ZZ$, UNICOMPRA$, 59, 4)
        COEFUNI = CVS(Mid$(UC$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
        If UNICOMPRA$ = TRIM$(UNIMED$(CI0%)) Then COEFUNI = 1
        Call REPLA(ZZ$, MKS$(COEFUNI), 63, 4)
        
        Call REPLA(ZZ$, MKD$(PreUnid#), 67, 8)
        '
        JJ& = JJ& + 1
        Call GRAREG("!OCOMDETA", ZZ$, JJ&)
          DELIBX$ = DESCRIT0$(CI0%)
        Call GRAREG(ARCHILIB$, DELIBX$, JJ&)
        '
      End If
29 Next KKI%
   '
   Call SETULTREG("!OCOMDETA", JJ&)
   Screen.MousePointer = 1
   '
End Sub

Sub CARCOANT()
   '
   Screen.MousePointer = 11
   List2.Clear
   NPRO% = Val(Text1)
   HOU = HOY(HOS$)
   '
   Call ABRECOMPRAS
   SQLX$ = " SELECT * FROM OCOMDETA "
   SQLX$ = SQLX$ + " WHERE Stat_Ocom < 8 "
   SQLX$ = SQLX$ + " AND Stat_Ocom >=0 "
   SQLX$ = SQLX$ + " AND Npro_Ocom = " & NPRO% & " "
   'Set CARCOANTMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim CARCOANTMP As ADODB.Recordset
   Set CARCOANTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   
   With CARCOANTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       Exit Sub
     End If
     Do While Not .EOF
       NUORSE& = !N_OrServ
             If NUORSE& < 1 Then
         COEFCAN = !Coef_Unids
         If COEFCAN < 0.005 Then COEFCAN = 1
         CACOM = !Cant_Ocom * COEFCAN
         CAREC = !Cant_Rec * COEFCAN
         CAPEN = CACOM - CAREC
         If CAPEN / COEFCAN >= 0.05 Then
           CIXI% = !Cod_Inte
           NROPED& = !Nume_Ocom
           UNICO$ = !Uni_Com
           FEX = CVINT(!FENTRE_SOL)
           ATRAU = 0
           If FEX < HOU Then
             ATRAU = (DIENTRE(FEX, HOU)) / 7
             If ATRAU > 99.9 Then ATRAU = 99.9
           End If
           ATRAX$ = FORMATNUM$(ATRAU, "F4.1")
           '
           LI2TEX$ = Space$(80)
           Call REPLA(LI2TEX$, CODEXT$(CIXI%), 1, 16)
           Call REPLA(LI2TEX$, DESCRIT0$(CIXI%), 17, 33)
           Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN / COEFCAN), 4, 8, 1, 2), 51, 8)
           Call REPLA(LI2TEX$, UNICO$, 60, 4)
           Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 8)
           Call REPLA(LI2TEX$, ATRAX$, 74, 4)
           Call REPLA(LI2TEX$, FORMATNUM$(NROPED&, "I6"), 79, 6)
           List2.AddItem LI2TEX$
           '
         End If
       End If
       .MoveNext
     Loop
   End With
   Screen.MousePointer = 1
   CARCOANTMP.Close
   Set CARCOANTMP = Nothing
   
   '
'   '
'   Screen.MousePointer = 11
'   List2.Clear
'   NPRO% = Val(Text1)
'   HOU = HOY(HOS$)
'   '
'   RFF$ = RECFILT$("OCOMDETA", 3, MKI$(NPRO%))
'   For UK& = 1 To Len(RFF$) Step 4
'     U& = CVS(Mid$(RFF$, UK&, 4))
'     X$ = REGLEIDO$("OCOMDETA", U&)
'     STAT% = Asc(Mid$(X$, 124, 1))
'     If STAT% < 8 Then      ' no esta anulada ni marcada cumplida
'       NUORSE& = CVS(Mid$(X$, 117, 4))
'       If NUORSE& < 1 Then
'         COEFCAN = CVS(Mid$(X$, 63, 4))
'         If COEFCAN < 0.005 Then COEFCAN = 1
'         CACOM = CVS(Mid$(X$, 55, 4)) * COEFCAN
'         CAREC = CVS(Mid$(X$, 99, 4)) * COEFCAN
'         CAPEN = CACOM - CAREC
'         If CAPEN / COEFCAN >= 0.05 Then
'           CIXI% = CVI(Mid$(X$, 9, 2))
'           NroPed& = CVS(Left$(X$, 4))
'           UNICO$ = TRIM$(Mid$(X$, 59, 4))
'           FEX = CVI(Mid$(X$, 95, 2))
'           ATRAU = 0
'           If FEX < HOU Then
'             ATRAU = (DIENTRE(FEX, HOU)) / 7
'             If ATRAU > 99.9 Then ATRAU = 99.9
'           End If
'           ATRAX$ = FORMATNUM$(ATRAU, "F4.1")
'           '
'           LI2TEX$ = Space$(80)
'           Call REPLA(LI2TEX$, CODEXT$(CIXI%), 1, 16)
'           Call REPLA(LI2TEX$, DESCRIT0$(CIXI%), 17, 33)
'           Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN / COEFCAN), 4, 8, 1, 2), 51, 8)
'           Call REPLA(LI2TEX$, UNICO$, 60, 4)
'           Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 8)
'           Call REPLA(LI2TEX$, ATRAX$, 74, 4)
'           Call REPLA(LI2TEX$, FORMATNUM$(NroPed&, "I6"), 79, 6)
'           List2.AddItem LI2TEX$
'           '
'         End If
'       End If
'     End If
'   Next UK&
'   '
'   Screen.MousePointer = 1
'   '
End Sub

Sub CARLISCO()
   '
   Screen.MousePointer = 11   ' cargando Lista de Compra
   HOII% = HOY(HOS$)
   FENTREGEN% = FECHANUM(Text14)
   List1.Clear: List1.ToolTipText = ""
   FIN& = ULTREG&("!OCOMDETA")
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   ARCHILIB$ = "!1OCOMDET"
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!OCOMDETA", U&)
     CI% = CVI(Mid$(X$, 9, 2))
     '
     If CI% <= 0 Then GoTo 10 'SOLO MATERIAL CODIFICADO
     '
     FENTRESOL% = CVI(Mid$(X$, 95, 2))
     If FENTRESOL% < FENTREGEN% Then FENTRESOL% = FENTREGEN%
     If FENTRESOL% < HOII% Then FENTRESOL% = HOII%
     If FENTRESOL% <> CVI(Mid$(X$, 95, 2)) Then
        Call REPLA(X$, MKI$(FENTRESOL%), 95, 2)
        Call GRAREG("!OCOMDETA", X$, U&)
     End If
     '
     UNIST$ = UCase$(TRIM$(Mid$(X$, 11, 4)))
     UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
     CANSTOK = CVS(Mid$(X$, 51, 4))
     CANCOMP = CVS(Mid$(X$, 55, 4))
     'If CANSTOK < 0.05 Then
     '   Call MENSERR(24, "Falta Cantidad en " + UNIST$ + "\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
     'End If
     '
     COEFI = CVS(Mid$(X$, 63, 4))
     If UNIST$ = UNICO$ Then COEFI = 1
     If COEFI < 0.0001 Then COEFI = 1
     '
     If CANCOMP < 0.05 Then
         CANCOMP = CANSTOK / COEFI
       Else
         CANSTOK = CANCOMP * COEFI
     End If
     CANCOMP = CANSTOK / COEFI
     Call REPLA(X$, MKS$(CANSTOK), 51, 4)
     Call REPLA(X$, MKS$(CANCOMP), 55, 4)
     Call REPLA(X$, MKS$(COEFI), 63, 4)
     Call GRAREG("!OCOMDETA", X$, U&)
     '
     UNISTO$ = Mid$(X$, 11, 4)
     UNICOX$ = Mid$(X$, 59, 4)
     COEFUNI = CVS(Mid$(X$, 63, 4))
     PREUNI = CVD(Mid$(X$, 67, 8))
     CANSTO = CVS(Mid$(X$, 51, 4))
     CANCOM = CVS(Mid$(X$, 55, 4))
     '
     PORDESCUIN = CVS(Mid$(X$, 75, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 99.9 Then PORDESCUIN = 0
     '
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, CODEXT$(CI%), 1, 15)
     Call REPLA(TTXX$, DESCRIT0$(CI%), 17, 27)
       If COLTELIB% > 0 Then
         Call REPLA(TTXX$, RTrim$(REGLEIDO$(ARCHILIB$, U&)), 17, 15)
       End If
     Call REPLA(TTXX$, UNISTO$, 48, 4)
     Call REPLA(TTXX$, FORMATNUM$(CANSTO, "F9.1"), 52, 9)
     Call REPLA(TTXX$, FORMATNUM$(PREUNI, "F10.4"), 61, 10)
     Call REPLA(TTXX$, FORMATNUM$(PORDESCUIN, "F5.1"), 72, 5)
     Call REPLA(TTXX$, FORMATNUM$(FENTRESOL%, "D8"), 78, 8)
     List1.AddItem TTXX$
     '
     'If CAN > 999999 Then
     '  Call MENSERR(24, "Cantidad Excesiva - Max 999999")
     '  Call REPLA(X$, MKS$(0), 47, 4)
     '  Call GRAREG("!CARDETPE", X$, U&)
     '  GoTo 10
     'End If
     'If CAN < 0.5 Then
     '  Call MENSERR(24, "Falta Cantidad")
     '  Call REPLA(X$, MKS$(0), 47, 4)
     '  Call GRAREG("!CARDETPE", X$, U&)
     '  GoTo 10
     'End If
     '
     '
10 Next U&
   'Call CIERRARCH("!CARDETPE")
   'Call CALTOPED
   List1.Refresh
   Screen.MousePointer = 1
End Sub

Function ULTICOM$(NPRO%, CIIX%)
   '
   ' trae el registro de OCOMDETA correspondiente a la
   ' última compra no anulada del artículo y proveedor
   ' con precio registrado mayor que cero
   '
   VUELTA% = 0
   '
20 UCO$ = REGBLAN$("OCOMDETA")
   Call REPLA(UCO$, MKI$(NPRO%), 7, 2)
   Call REPLA(UCO$, MKI$(CIIX%), 9, 2)
   '
   SQLX$ = "SELECT * FROM OCOMDETA "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIIX% & " "
   SQLX$ = SQLX$ + "AND Npro_Ocom = " & NPRO% & " "
   If VUELTA% = 0 Then SQLX$ = SQLX$ + "AND Pre_Unit >= 0.00005 "
   SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
   SQLX$ = SQLX$ + "ORDER BY Femi_Ocom ASC, Nume_Ocom ASC"
   'Set OcomdetaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim OCOMDETATEMP As ADODB.Recordset
   Set OCOMDETATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   
   On Error Resume Next
   With OCOMDETATEMP
     .MoveLast
     If Err < 1 Then
         Call REPLA(UCO$, MKS$(!Nume_Ocom), 1, 4)
         Call REPLA(UCO$, MKI$(FECHANUM(Format(!Femi_Ocom, "dd/mm/yy"))), 5, 2)
         Call REPLA(UCO$, !Uni_Mas, 11, 4)
         Call REPLA(UCO$, !Uni_Com, 59, 4)
         Call REPLA(UCO$, MKS$(!Coef_Unids), 63, 4)
         Call REPLA(UCO$, Chr$(!Mone_Emis), 122, 1)
         Call REPLA(UCO$, MKD$(!Pre_Unit), 67, 8)
         Call REPLA(UCO$, MKS$(!Desc_Item), 75, 4)
       Else
         If VUELTA% < 1 Then
           VUELTA% = VUELTA% + 1
           GoTo 20
         End If
     End If
   End With
   OCOMDETATEMP.Close
   Set OCOMDETATEMP = Nothing
   '
   ULTICOM$ = UCO$
   '
End Function

Sub SHOWOCOM(NPX&)
    '
    TIDOCUM$ = "Nota de Pedido"
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
    Exit Sub
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
    Exit Sub
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
    EDITFORM.Command1.Enabled = False
    EDITFORM.Show 1
    '
End Sub

Sub PRICOMPRA()
    '
    CPAG% = Val(TEXT21(0))
    FORPAG$ = TRIM$(Eco21(0))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    TIEM% = Val(TEXT21(1))
    EMBALA$ = TRIM$(Eco21(1))
    '
    Call SETULTREG("!OBSERVOF", 0)
    Call FRATEXTO(Text13, 52)
    For KKU& = 1 To CARETEX%
      Call REGAPP("!OBSERVOF", RETEX$(KKU&))
    Next KKU&
    Call CIERRARCH("!OBSERVOF")
    '
    PRIPLANOI% = 0
    If CONTROL$("OCOMPRA", "PRIPLANO") = "SI" Then PRIPLANOI% = 1
    PRIESPEC% = 0
    If CONTROL$("OCOMPRA", "PRIESPEC") = "SI" Then PRIESPEC% = 1
    PRIHRINS% = 0
    If CONTROL$("OCOMPRA", "PRIHRINS") = "SI" Then PRIHRINS% = 1
    '
    ' IMPRESION DE LA O/COMPRA
    '
    FORIPRE$ = CONTROL$("", "FORMOCOM")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHEMII% = FECHANUM(Text9)
      FECHDOC$ = FETEXTO$(FECHEMII%)
      NUMEDOC$ = TRIM$(Text11)
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Nota de Pedido"
      '
      CODPARAM$(1) = "COND-PAG"
      CODPARAM$(2) = "NUBULTOS"
      CODPARAM$(3) = "IMP-NETO"
      CODPARAM$(4) = "POR-IVA"
      CODPARAM$(5) = "IMP-IVA"
      CODPARAM$(6) = "IMP-TOTA"
      CODPARAM$(7) = "VENDEDOR"
      CODPARAM$(8) = "LIST-PRE"
      CODPARAM$(9) = "ZONA-VEN"
      CODPARAM$(10) = "FECH-LAN"
      CODPARAM$(11) = "IMPNEDES"
      CODPARAM$(12) = "TRANSPOR"
      CODPARAM$(13) = "POR-DESC"
      CODPARAM$(14) = "IMP-DESC"
      CODPARAM$(15) = "FECH-REC"
      CODPARAM$(16) = "TIPO-VAL"
      CODPARAM$(17) = "IMPO-CHE"
      CODPARAM$(18) = "REF-CONJ"
      CODPARAM$(19) = "DOMI-TRA"
      CODPARAM$(20) = "PERC-IVA"
      '\\\OT-05-0676-WAR-20/10/05///
      CODPARAM$(21) = "DOCU-APL" 'REGLAMENTOS DE COMPRA
      '\\\OT-05-0676-WAR-FIN///
      CAPARDOC% = 21
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = EMBALA$
      DOCPARAM$(3) = Total(0)
      DOCPARAM$(4) = Text12
      DOCPARAM$(5) = Total(2)
      DOCPARAM$(6) = Total(3)
      DOCPARAM$(7) = "" 'Mid$(X1$, 147, 20)
      DOCPARAM$(8) = "" 'Mid$(X1$, 67, 40)
      DOCPARAM$(9) = "" 'Mid$(X1$, 167, 40)
      DOCPARAM$(10) = "" ' Mid$(X1$, 107, 40)
      DOCPARAM$(11) = Str$(Val(Total(0)) - Val(Total(1)))
      DOCPARAM$(12) = TRIM$(Text6) ' ATENCION SR.
      DOCPARAM$(13) = TRIM$(PORDESCU)
      DOCPARAM$(14) = Total(1)
      HOII% = HOY(HOS$)
      FECHENT% = FECHANUM(Text14)
      DOCPARAM$(15) = Text14
      If FECHENT% = HOII% Then
          DOCPARAM$(15) = DOCPARAM$(15) + " (inmediato)"
        ElseIf FECHENT% < HOII% + 4 Then
          DOCPARAM$(15) = DOCPARAM$(15) + " (urgente!!)"
      End If
      DOCPARAM$(16) = TRIM$(Label11)
      DOCPARAM$(17) = ""
      If MONEMI% > 1 Then DOCPARAM$(17) = Str$(TICAMONE(MONEMI%))
      DOCPARAM$(18) = Eco21(1)
      DOCPARAM$(19) = Text10
      DOCPARAM$(20) = Total(4)
      '\\\OT-05-0676-WAR-20/10/05///
      DOCPARAM$(21) = RAI$ 'REGLAMENTO COMPRA ANEXO
      '\\\OT-05-0676-WAR-FIN///
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "FECH-VEN"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "PRE-LIST"
      CODTABU1$(8) = "PORDESCU"
      CODTABU1$(9) = "PRE-NETO"
      CODTABU1$(10) = "NET-ITEM"
      CODTABU1$(11) = "CAKILOS"
      CODTABU1$(12) = "F-PLANO"
      CODTABU1$(13) = "ORD-ITEM"
      CODTABU1$(14) = "DESTINO"
      CODTABU1$(15) = "EQUIPO1"   ' para sector solicitante
      CACOLTAB1% = 15
      '
      DESTIDOC% = (-1) * Val(Text1)
      CAITAB1% = 0
      CAITCRU% = 0
      INCLUSPE% = 0
      If CONTROL$("OCOMPRA", "INCLUSPE") = "SI" Then
         INCLUSPE% = 1
      End If
      '
      REPLAORI% = 0
      If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
         REPLAORI% = 1
      End If
      '
      RECOPROP% = 0
      If CONTROL$("OCOMPRA", "RECOPROP") = "SI" Then
         RECOPROP% = 1
      End If
      '
      REVICODX% = 0
      If CONTROL$("OCOMPRA", "REVICOD") = "SI" Then
         REVICODX% = 1
      End If
      '
      NC1% = DESTIDOC%
      ARCHILIB$ = "!1OCOMDET"
      For J& = 1 To ULTREG&("!OCOMDETA")
         X$ = REGLEIDO$("!OCOMDETA", J&)
         CI0% = CVI(Mid$(X$, 9, 2))
         CANTU = CVS(Mid$(X$, 55, 4))
         PreUnid# = CVD(Mid$(X$, 67, 8))
         PORDE = CVS(Mid$(X$, 75, 4))
         FEX = CVI(Mid$(X$, 95, 2))
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 97, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
         CODDX$ = TRIM$(CODEXT$(CI0%))
         CPROP$ = CODDX$
         If REPLAORI% = 1 Then
            CODDX$ = TRIM$(CODORIG$(CI0%, NC1%))
         End If
         If REVICODX% = 1 Then
            FEREVX = FEREVI%(CI0%)
            RVCDDX$ = TRIM$(REVICOD$(CI0%))
            If RVCDDX$ <> "" Then
              CODDX$ = CODDX$ + " (Letra: " + RVCDDX$ + " - " + FECHATEX$(FEREVX) + ")"
            End If
         End If
         TETABU1$(1, CAITAB1%) = CODDX$
           DECRIT$ = TRIM$(DESCRIT0$(CI0%))
           If REPLAORI% = 1 Then
             If CODDX$ <> CPROP$ Then
               If RECOPROP% > 0 Then
                 DECRIT$ = DECRIT$ + " (" + CPROP$ + ")"
               End If
             End If
           End If
         TETABU1$(3, CAITAB1%) = Mid$(X$, 59, 4)
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 97, 2)  'consigna m.prima
         TETABU1$(7, CAITAB1%) = Str$(PreUnid#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PreUnid# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PreUnid# * (1 - PORDE / 100))
           KILOX$ = FORMATNUM$(CVS(Mid$(X$, 51, 4)) * PESUNI(CI0%), "F12.3")
         TETABU1$(11, CAITAB1%) = KILOX$
         TETABU1$(12, CAITAB1%) = FEPLANO$(CI0%)
         TETABU1$(13, CAITAB1%) = TRIM$(Str$(J&))
         TETABU1$(14, CAITAB1%) = TRIM$(Mid$(X$, 105, 12))   ' Imputacion contable
         TETABU1$(15, CAITAB1%) = TRIM$(Mid$(X$, 15, 12))   ' Sector
         '
         POCODI% = Val(ATRIFORM$(FORIPRE$, "COD-MAT/POSINHOR"))
         PODESI% = Val(ATRIFORM$(FORIPRE$, "DES-MAT/POSINHOR"))
         If PODESI% <= POCODI% + 3 Then
           If TRIM$(CODDX$) <> "" Then
             CAITAB1% = CAITAB1% + 1
           End If
         End If
         '
         TETABU1$(2, CAITAB1%) = DECRIT$
         If COLTELIB% > 0 Then
           TELIBRE$ = RTrim$(REGLEIDO$(ARCHILIB$, J&))
           If REPLAORI% = 1 Then
             If CODDX$ <> CPROP$ Then
               If RECOPROP% > 0 Then
                 TELIBRE$ = TELIBRE$ + " (" + CPROP$ + ")"
               End If
             End If
           End If
           'FORTELIB$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "DES-MAT/FORMATO")))'\\**//
           FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
           ATELIB% = Val(Mid$(FORTELIB$, 2))
           Call FRATEXTO(TELIBRE$, ATELIB%)
           TETABU1$(2, CAITAB1%) = RETEX$(1)
           For KKU1% = 2 To CARETEX%
             CAITAB1% = CAITAB1% + 1
             TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
           Next KKU1%
         End If
         '
         UNIST$ = Mid$(X$, 11, 4)
         UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
         If TRIM$(Left$(UCase$(UNIST$), 2)) <> UNICO$ Then
           DEADI$ = "(equiv. " + TRIM$(FORMATNUM$(CVS(Mid$(X$, 51, 4)), "F12.1")) + " " + TRIM$(ECOARCH$("UNIMED", UNIST$)) + ".)"
           CAITAB1% = CAITAB1% + 1
           TETABU1$(2, CAITAB1%) = DEADI$
         End If
         '
         If INCLUSPE% > 0 Then
           RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI0%))
           If Len(RNC$) > 0 Then
             INTERLI = Val(ATRIFORM$(FORIPRE$, "INTERLIN"))
             '
             COLUTA% = 0
             URN& = ULTREG&("ESPECIFI")
             '
             For U& = 1 To Len(RNC$) Step 4
               RENOTA& = CVS(Mid$(RNC$, U&, 4))
               If RENOTA& > 0 Then
                 If RENOTA& <= URN& Then
                   TTXY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
                   TTYY$ = REPLACAR$(TTXY$, Chr$(9), "     ")
                   If INTERLI <= 6 Or InStr(UCase$(FORIPRE$), "MZZ") > 0 Then
                       CAITAB1% = CAITAB1% + 1
                       TETABU1$(2, CAITAB1%) = TTYY$
                     Else
                       COLUTA% = COLUTA% + 1
                       If COLUTA% > 2 Then COLUTA% = 1
                       If COLUTA% = 1 Then CAITAB1% = CAITAB1% + 1
                       TETABU1$(COLUTA%, CAITAB1%) = TTYY$
                   End If
                 End If
               End If
             Next U&
             '
           End If
         End If
         '
109   Next J&
      '
      Call PRINTDOC("FORMOCOM")   ' Orden de Compra
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
      DOCUORIG$ = "N/Ped-" + NUMEDOC$
      NUMEDOC$ = ""
      '
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "ARCON") > 0 Then
         Call CIERRARCH("*.*")
         Call BAJALDISCO
         ANEXOCOM.Show 1
         GoTo 199
      End If
      '
      For J& = 1 To ULTREG&("!OCOMDETA")
         X$ = REGLEIDO$("!OCOMDETA", J&)
         CI0% = CVI(Mid$(X$, 9, 2))
         If CI0% > 0 Then
           If CI0% <= NUMAS% Then
             If PRIPLANOI% = 1 Then Call PRIPLANO(CI0%)
             If PRIESPEC% = 1 Then Call PRIHOJESP(CI0%)
             If PRIHRINS% = 1 Then Call PRIHOJRINSP(CI0%)
           End If
         End If
      Next J&
      '
     End If
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    If EMICONSI% = 1 Then
      If COMALTER%("Puede Ahora Generar\el Remito de Consignacion de Materiales\\Consignar Materiales\Continuar") = 1 Then
        NPRO% = Val(Text1)
        'Call ENTRECONSIACC(NPRO%)
        Call NUENTRECONSI(NPRO%)
      End If
    End If
    Call CIERRARCH("*.*")
    '
    Exit Sub
    '
999 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub
'\\\OT-05-0676-WAR-14/10/05///
Sub PRIREGLACOM() 'IMPRIME REGLAMENTO DE COMPRA
  '
  With FORSELREG
    '
    For JJ% = 0 To .LIMATSEL.ListCount - 1
      If .LIMATSEL.Selected(JJ%) = True Then
        NREGLA& = Mid$(.LIMATSEL.List(JJ%), 1, 5)
        If NREGLA& > 0 Then
          'IMPRESIÓN DEL REGLAMENTO
          FORIPRE$ = CONTROL$("", "REGLACOM")
          If TRIM$(FORIPRE$) <> "" Then
           '
           Screen.MousePointer = 11
           FECHDOC$ = Text9
           NumOC$ = TRIM$(FORSELREG.Label3) '\\**//
           While Len(NumOC$) < 6
             NumOC$ = "0" + NumOC$
           Wend
           '
           TIPODOC$ = "Reglamento de Compra"
           '
           CODPARAM$(1) = "NUME-COM"
           CAPARDOC% = 1
           '
           DOCPARAM$(1) = NumOC$
           '
           CODTABU1$(1) = "REF-MAT1"
           CACOLTAB1% = 1
           '
           CAITAB1% = 0
           DESREGLA$ = DESCRIREGLA$(NREGLA&) 'DESCRIPCION DEL REGLAMENTO
           Call FRATEXTO(DESREGLA$, 92)
           For U& = 1 To CARETEX%
             Z$ = "." + RETEX$(U&)
             CAITAB1% = CAITAB1% + 1
             TETABU1$(1, CAITAB1%) = Z$
           Next U&
           '
           Call PRINTDOC("REGLACOM")
           '
          End If
        End If
      End If
    Next JJ%
    Screen.MousePointer = 1
  End With
  '
  Unload FORSELREG
  '
End Sub
  
Sub SELEC_REGLACOM(OKX%) 'selecciona REGLAMENTO DE COMPRA
   '
   OKX% = 0
   RAI$ = "" 'REGLAMENTO COMPRA ANEXO
   FOREGLA$ = TRIM$(CONTROL$("", "REGLACOM"))
   If FOREGLA$ <> "" Then
     If ECOARCH$("INDIFRM", FOREGLA$) <> "" Then
       If ULTREG&("INDIRCOM") > 0 Then
         GoTo 10
       End If
     End If
   End If
   OKX% = 1 'SI BASE DE DATOS NO TIENE  REGLAMENTOS O NO USA
   Exit Sub
   '
10 With FORSELREG
    .LIMATSEL.Clear
    FIN& = ULTREG&("INDIRCOM")
    For J& = 1 To FIN&
      X$ = REGLEIDO$("INDIRCOM", J&)
      NR& = CVI(Mid$(X$, 1, 2))
      DRC$ = Mid$(X$, 3, 46)
      RC$ = FORMATNUM(NR&, "F5.0") & Chr(9) & DRC$
      .LIMATSEL.AddItem RC$
    Next J&
    .Show 1
    If .Label1 = "OK" Then
      OKX% = 1
    End If
    Call FRESHALL
    DoEvents
    '
    RAI$ = "" 'REGLAMENTOS ADJUNTOS IMPRESOS
    For JJ% = 0 To .LIMATSEL.ListCount - 1
      If .LIMATSEL.Selected(JJ%) = True Then
        NREGLA& = Val(Mid$(.LIMATSEL.List(JJ%), 1, 5))
        If NREGLA& > 0 Then
          RAI$ = RAI$ + TRIM$(FORMATNUM(NREGLA&, "F6.0")) + "-"
        End If
      End If
    Next JJ%
    Screen.MousePointer = 1
    '
    If Len(RAI$) > 0 Then
      RAI$ = Mid$(RAI$, 1, Len(RAI$) - 1)
    End If
    '
  End With
  '
End Sub
Private Function DESCRIREGLA$(NREGLA&)
   '
   CIIINSTRU% = Val(NREGLA&)
   If CIIINSTRU% < 1 Then Exit Function
   '
   REGI& = CIIINSTRU%
   URN& = ULTREG&("REGLACOM")
   '
   DESCRIREGLA$ = "": TTYZ$ = ""
   If CIIINSTRU% > 0 Then
     If CIIINSTRU% <= URN& Then
       Screen.MousePointer = 11
       RNC$ = RECFILT$("REGLACOM", 1, MKI$(CIIINSTRU%))
       '
       For U& = 1 To Len(RNC$) Step 4
         RENOTA& = CVS(Mid$(RNC$, U&, 4))
         If RENOTA& > 0 Then
           If RENOTA& <= URN& Then
             TTYY$ = RTrim$(Mid$(REGLEIDO$("REGLACOM", RENOTA&), 3))
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
   '
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
   Next U&
   U& = 0
   '
30 DESCRIREGLA$ = Left$(TTYZ$, U&)
  '
End Function
'\\\OT-05-0676-WAR-FIN///
Private Sub UcoMat_Click()
Call OPLISCO04.Command20_Click
End Sub

Private Sub UcoPro_Click()
Call OPLISCO04.Command12_Click
End Sub

Private Sub Ver_Ocom_Click()
    Call LISCOMAQ04.Command22_Click
End Sub
'///////******RUTINA VIEJA DE  CONSIGNACION
'Sub ENTRECONSIACC(NPROVE%)
'        '
'        LLAMACOMPRA% = 0
'        If NPROVE% > 0 Then LLAMACOMPRA% = 1
'        '
'        FORIPRE$ = TRIM$(CONTROL$("", "FOREMICO"))
'        If FORIPRE$ <> "" Then
'            LU$ = LUDAT$
'            If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
'                RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
'                Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
'                GoTo 9999
'            End If
'        End If
'        '
'        DEPOCON% = Val(CONTROL$("", "DEPOCONS"))
'        DEPORIG% = DEPOCON%
'        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
'        '
'        'Call BLIMESS("Revisando Archivo de Movimientos de Stock")
'1100    Screen.MousePointer = 11
'        CMO$ = TRIM$(CONTROL$("", "CODRECON"))
'        If CMO$ = "" Then CMO$ = "RC"
'        FIN& = ULTREG&("MOVISTO")
'        For U& = FIN& To 1 Step -1
'            Call TRACE(20, FIN& - U& + 1, FIN&)
'            XX$ = REGLEIDO$("MOVISTO", U&)
'            If UCase$(Mid$(XX$, 21, 2)) = CMO$ Then
'                UNRO& = Val(Mid$(XX$, 26, 7))
'                GoTo 1002
'            End If
'        Next U&
'        UNRO& = 0
'        '
'1002    REMI$ = CMO$ + "-" + TRIM$(Str$(UNRO& + 1))
'        Call FRESHALL '¿?
'        '
'1105    Call CIERRARCH("*.*")
'        '
'        Screen.MousePointer = 1
'        If NPROVE% > 0 Then
'             EBOL$ = MKI$(NPROVE%)
'          Else
'             EBOL$ = OBJPLA$("SELPROVE", "")
'             If EBOL$ = "" Then
'                GoTo 9999
'             End If
'        End If
'        '
'        NC% = CVI(EBOL$) 'Obtengo el numero de prov
'        If NC% < 1 Or ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
'            Call MENSERR(24, "Proveedor Inexistente")
'            NPROVE% = 0
'            GoTo 1105
'        End If
'        '
'        NC1% = (-1) * NC%
'        '
'        Y$ = PEDY$: If Y$ = "" Then Y$ = Space$(128)
'        Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
'        Call REPLA(Y$, EBOL$, 7, 2)
'        Call REPLA(Y$, MKI$(HOY(HO$)), 5, 2)
'        Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
'        Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
'        CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
'        Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
'1110    Call REPLA(Y$, REMI$, 105, 16)
'        Call REPLA(Y$, Chr$(DEPORIG%), 121, 1)
'        '
'1111    PEDY$ = OBJPLA$("REMICON-ENCB", Y$)
'        If PEDY$ = "" Then GoTo 9998
'        FF% = CVI(Mid$(PEDY$, 5, 2))
'        REMI$ = TRIM$(Mid$(PEDY$, 105, 16))
'        If TRIM$(REMI$) = "" Then
'            Y$ = PEDY$
'            Call MENSERR(24, "Falta Numero de Remito")
'            GoTo 1111
'        End If
'        '
'        DEPORIG% = Asc(Mid$(PEDY$, 121, 1))
'        If DEPOCON% > 0 Then
'          If DEPORIG% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPORIG%)) = "" Then
'            Call MENSERR(24, "Depósito de Salida no Válido")
'            DEPORIG% = Val(CONTROL$("", "DEPOCONS"))
'            GoTo 1110
'          End If
'        End If
'        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
'        '
'1112    Call CIERRARCH("*.*")
'        Call BLANARCH("!BOLREDET")
'        '
'        Screen.MousePointer = 11
'        HAYOC% = 0
'        For I& = 1 To ULTREG&("MACONSIG") 'Aca empieza con maconsig
'            X$ = REGLEIDO$("MACONSIG", I&)
'            If CVI(Mid$(X$, 3, 2)) = NC% Then
'                CAPENE# = CVD(Mid$(X$, 55, 8)) - CVD(Mid$(X$, 63, 8)) ' Cant asignada - Cant entregada
'                If CAPENE# > 0 Then
'                    Nuoc& = Val(Mid$(X$, 24, 6))
'                    CI% = CVI(Mid$(X$, 33, 2))
'                    STODISPO# = CDbl(STOCKACT(CI%))
'                    If STODISPO# < 0 Then STODISPO# = 0
'                    If STODISPO# < CAPENE# Then CAPENE# = STODISPO# ' Chequea el stock disponible
'                    '
'                    Y$ = REGBLAN$("!BOLREDET") 'Nueva boleta recep. prov
'                    Call REPLA(Y$, MKI$(CI%), 83, 2) 'cod mat
'                    Call REPLA(Y$, UNIMED$(CI%), 85, 4) 'unidad de medida
'                    Call REPLA(Y$, MKD$(CAPENE#), 89, 8) 'Cantidad nominal
'                    Call REPLA(Y$, MKS$(Nuoc&), 151, 4) 'Nº O/C
'                    Call REPLA(Y$, MKS$(I&), 189, 4) ' Relleno
'                    HAYOC% = 1
'                    Call REGAPP("!BOLREDET", Y$)
'                End If
'            End If
'        Next I&
'        '
'        If HAYOC% = 0 Then
'            Call COMUNI("No hay Consignaciones Pendientes\para este Proveedor.")
'        End If
'        '
'1113    VTB% = VENTABU%("ECONSIG/NC/TITUPAN=Consignación de Materiales - " + TRIM$(RASOCLI$(NC1%)))
'        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
'            X$ = "": Y$ = "": YY$ = ""
'            EBOL$ = "": PEDY$ = "": DESCRI$ = ""
'            GoTo 1105
'        End If
'        '
'        For YX& = 1 To ULTREG&("!BOLREDET")
'          YY$ = REGLEIDO$("!BOLREDET", YX&)
'          CI% = CVI(Mid$(YY$, 83, 2)) 'Asigna el cod mat
'          CAN# = CVD(Mid$(YY$, 89, 8)) ' Asigna la cant nominal
'          If CI% > 0 Then
'            If CI% <= NUMAS% Then
'              If CAN# > 0.005 Then
'                GoTo 1114
'              End If
'            End If
'          End If
'        Next YX&
'        Call MENSERR(24, "Consignación Cancelada !!!\Faltan Cantidades o Códigos.")
'        GoTo 1113
'        '
'1114    If TRIM$(FORIPRE$) <> "" Then
'           DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
'           If DTP$ = "" Then GoTo 1113         ' CANCELADO
'           DATRANSPO$ = DTP$
'        End If
'        '
'11141   CONTA% = ULTREG&("!BOLREDET")
'        '
'1115    If CONTA% > 0 Then
'          '
'          Call CIERRARCH("MOVISTO")
'          Call BLOQARCH("MOVISTO")
'          Call BLOQARCH("MACONSIG")
'          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
'          '
'          AINDI% = 0
'          For YX& = 1 To ULTREG&("!BOLREDET")
'             YY$ = REGLEIDO$("!BOLREDET", YX&)
'             CI% = CVI(Mid$(YY$, 83, 2))  'Asigna cod mat de boleta recep. prov.
'             CAN# = CVD(Mid$(YY$, 89, 8)) 'Asigna cant nominal
'             LOTE$ = Mid$(YY$, 69, 12)    'Asigna identif. lote
'             If CI% > 0 Then
'               If CI% <= NUMAS% Then
'                 If CAN# > 0.005 Then
'                    Nuoc& = CVS(Mid$(YY$, 151, 4)) ' Asigna Nº O/C
'                    REGOC& = CVS(Mid$(YY$, 189, 4)) 'Asigna filler
'                    DOSEC$ = ""
'                    If REGOC& > 0 Then
'                       If REGOC& <= ULTREG&("MACONSIG") Then
'                          ZZ$ = REGLEIDO$("MACONSIG", REGOC&)
'                          If CVI(Mid$(ZZ$, 3, 2)) = NC% Then ' Si son = PROV
'                             If CVI(Mid$(ZZ$, 33, 2)) = CI% Then ' Si son = MAT
'                                DOSEC$ = Mid$(ZZ$, 21, 12)       ' asigna Nº O/C
'                                TOENTRE# = CVD(Mid$(ZZ$, 63, 8)) + CAN# 'cant entregada + cant nominal
'                                Call REPLA(ZZ$, MKD$(TOENTRE#), 63, 8)  'Escribe como: nueva cantidad entregada
'                                If FF% > CVI(Mid$(ZZ$, 71, 2)) Then
'                                    Call REPLA(ZZ$, MKI$(FF%), 71, 2)   'Escribe la fecha de entrega
'                                End If
'                                Call GRAREG("MACONSIG", ZZ$, REGOC&) ' Graba el nuevo registro
'                                GoTo 1120   ' Va al 1120
'                             End If
'                          End If
'                       End If
'                    End If
'                    '
'                    ZZ$ = REGBLAN$("MACONSIG") 'Hace un blanco en el archivo maconsig
'                    Call REPLA(ZZ$, MKI$(FF%), 1, 2) 'fecha de reserva
'                    Call REPLA(ZZ$, MKI$(NC%), 3, 2) ' Nº de prov.
'                       NOCOM$ = TRIM$(Str$(Nuoc&)) ' Calcula Nº O/C.
'                       While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
'                       NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-00"
'                    Call REPLA(ZZ$, NOCOM$, 21, 12) ' Nº O/C
'                    Call REPLA(ZZ$, MKI$(CI%), 33, 2) ' Codigo de Mat
'                    Call REPLA(ZZ$, UNIMED$(CI%), 51, 4) 'Unidad de medida
'                    Call REPLA(ZZ$, MKD$(CAN#), 55, 8) 'Cant asignada
'                    Call REPLA(ZZ$, MKD$(CAN#), 63, 8) ' Cant entregada
'                    Call REPLA(ZZ$, MKI$(FF%), 71, 2)   'Fecha de entrega
'                    Call REGAPP("MACONSIG", ZZ$)
'                    '
'1120                CANTI = (-1) * CAN# ' Hace negativa la cantidad nominal
'                    'LLAMA A LA FUNCION MOVISTO
'                    'Call Movisto(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, RASOCLI$(NC1%), 0, "$", NC1%, DEPORIG%, CANTI)
'                 End If
'               End If
'             End If
'             '
'          Next YX&
'          Call CIERRARCH("MOVISTO")
'          Call CIERRARCH("MACONSIG")
'          '
'          Call FRESHECHO("IREPRIPE")
'          Screen.MousePointer = 1
'          '
'          If TRIM$(FORIPRE$) <> "" Then
'            '
'            FEX = FF%
'            FECHDOC$ = FECHATEX$(FEX)
'            NUMEDOC$ = TRIM$(REMI$)
'            TIPODOC$ = "Remito de Traslado"
'            '
'            CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
'            CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
'            CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
'            CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
'            CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
'            CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
'            CODPARAM$(7) = "LOCA-TRA": DOCPARAM$(7) = TRIM$(Mid$(DTP$, 159, 32))
'            CODPARAM$(8) = "NUBULTOS": DOCPARAM$(8) = TRIM$(Mid$(DTP$, 153, 6))
'            CAPARDOC% = 8
'            '
'            CODTABU1$(1) = "COD-MAT"
'            CODTABU1$(2) = "DES-MAT"
'            CODTABU1$(3) = "CANTIDAD"
'            CODTABU1$(4) = "CAKILOS"
'            CODTABU1$(5) = "UNIMED"
'            CODTABU1$(6) = "N-OCOMPR"
'            CACOLTAB1% = 6
'            '
'            DESTIDOC% = NC1%
'            CAITAB1% = 0
'            '
'            For YX& = 1 To ULTREG&("!BOLREDET")
'               YY$ = REGLEIDO$("!BOLREDET", YX&)
'               CI% = CVI(Mid$(YY$, 83, 2))
'               CAN# = CVD(Mid$(YY$, 89, 8))
'               If CI% > 0 Then
'                  If CI% <= NUMAS% Then
'                     If CAN# > 0.005 Then
'                        UNID$ = TRIM$(Mid$(YY$, 85, 4))
'                        If UNID$ = "" Then UNID$ = UNIMED$(CI%)
'                        CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
'                        KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
'                        NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
'                        NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
'                        '
'                        CAITAB1% = CAITAB1% + 1
'                        TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
'                        TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
'                        TETABU1$(3, CAITAB1%) = CANS$
'                        TETABU1$(4, CAITAB1%) = KILO$
'                        TETABU1$(5, CAITAB1%) = UNID$
'                        TETABU1$(6, CAITAB1%) = NUOCO$
'                        '
'                     End If
'                  End If
'               End If
'            Next YX&
'            '
'            Call PRINTDOC("FOREMICO")   ' Remito en Consignacion
'            '
'          End If
'          '
'1199    End If
'        '
'9998    Call CIERRARCH("*.*")
'        '
'        X$ = "": Y$ = "": YY$ = ""
'        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
'        '
'        If LLAMACOMPRA% = 0 Then
'           GoTo 1100
'        End If
'        '
'9999    Call CIERRARCH("*.*")
'        '
'End Sub
''

Sub CARDETCOM1()
   '
   Call CIERRARCH("!OCOMDETA")
   List2.ListIndex = (-1)
   HOII% = HOY(HOS$)
   '
   For U& = 1 To ULTREG&("!OCOMDETA")
     ZZ$ = REGLEIDO$("!OCOMDETA", U&)
     If TRIM$(Mid$(ZZ$, 105, 12)) = "" Then Call REPLA(ZZ$, Text7, 105, 12)
     If TRIM$(Mid$(ZZ$, 15, 12)) = "" Then Call REPLA(ZZ$, Text18, 15, 12)
     Call GRAREG("!OCOMDETA", ZZ$, U&)
   Next U&
   '
10 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Detalle de Orden de Compra en " + TRIM$(Label11)
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(140 + Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   '
   Timer1.Enabled = True
   PRF$ = "CARDETC1"
   XXX% = VENTABU%(PRF$ + ATRI$)
   '
   If XXX% < 0 Or ULTREG&("!OCOMDETA") < 1 Then
      Call CIERRARCH("!OCOMDETA")
      Call BLANARCH("!OCOMDETA")
      Call BLANFORMU
      Total(0).TEXT = ""
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
End Sub
