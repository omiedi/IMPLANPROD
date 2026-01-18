VERSION 5.00
Begin VB.Form FOCAMOD07 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Compra Abiertas - Proveedores Externos"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "FOCAMOD07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "O/C's Abiertas en Curso"
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
         Height          =   1960
         Left            =   160
         TabIndex        =   70
         Top             =   6030
         Width           =   10580
         Begin VB.Frame Frame14 
            Caption         =   "Frame14"
            Height          =   1730
            Left            =   10290
            TabIndex        =   71
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame13 
            Caption         =   "Frame13"
            Height          =   1730
            Left            =   9345
            TabIndex        =   72
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame12 
            Caption         =   "Frame12"
            Height          =   1730
            Left            =   8085
            TabIndex        =   73
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame10 
            Caption         =   "Frame10"
            Height          =   1730
            Left            =   1890
            TabIndex        =   74
            Top             =   210
            Width           =   15
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00C0C0C0&
            Height          =   435
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   10500
            TabIndex        =   76
            Top             =   315
            Width           =   10565
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Código          Descripción                                         F.Vigen.    O/C"
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
               TabIndex        =   77
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
            TabIndex        =   75
            Top             =   750
            Width           =   10565
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2050
         Left            =   0
         TabIndex        =   51
         Top             =   3990
         Width           =   10800
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   57
            Top             =   0
            Width           =   7845
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
               Height          =   330
               Left            =   1155
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   65
               Top             =   1470
               Width           =   6420
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
               Left            =   1890
               TabIndex        =   64
               Text            =   " "
               Top             =   600
               Width           =   5685
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
               Left            =   1155
               TabIndex        =   63
               Text            =   " "
               Top             =   600
               Width           =   435
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
               Left            =   1575
               TabIndex        =   62
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
               Height          =   250
               Index           =   0
               Left            =   1890
               TabIndex        =   61
               Top             =   315
               Width           =   5685
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
               TabIndex        =   60
               Text            =   " "
               Top             =   315
               Width           =   435
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
               TabIndex        =   59
               Top             =   315
               Width           =   180
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
               Height          =   250
               Left            =   1155
               TabIndex        =   58
               Top             =   890
               Width           =   6420
            End
            Begin VB.Label Label13 
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
               Height          =   255
               Left            =   1155
               TabIndex        =   81
               Top             =   1185
               Width           =   5580
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Anexos:"
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
               Index           =   7
               Left            =   -3045
               TabIndex        =   80
               Top             =   1205
               Width           =   4110
            End
            Begin VB.Label Label15 
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
               Height          =   255
               Left            =   7140
               TabIndex        =   79
               Top             =   1185
               Width           =   420
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "St:"
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
               Left            =   2940
               TabIndex        =   78
               Top             =   1200
               Width           =   4110
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
               TabIndex        =   69
               Top             =   1470
               Width           =   4110
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
               Left            =   -3045
               TabIndex        =   68
               Top             =   655
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
               TabIndex        =   67
               Top             =   370
               Width           =   1590
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
               TabIndex        =   66
               Top             =   930
               Width           =   4110
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Moneda"
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
            Left            =   8190
            TabIndex        =   52
            Top             =   0
            Width           =   2550
            Begin VB.TextBox Text122 
               Alignment       =   1  'Right Justify
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
               Left            =   630
               TabIndex        =   53
               TabStop         =   0   'False
               Text            =   " "
               Top             =   1470
               Width           =   1440
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
               Left            =   735
               TabIndex        =   56
               ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
               Top             =   420
               Width           =   1275
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cotización al "
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
               Left            =   210
               TabIndex        =   55
               Top             =   1050
               Width           =   1170
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " "
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
               Index           =   5
               Left            =   1260
               TabIndex        =   54
               Top             =   1005
               Width           =   1170
            End
         End
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
         Left            =   6195
         TabIndex        =   47
         Top             =   440
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
         Left            =   6880
         TabIndex        =   46
         Top             =   840
         Visible         =   0   'False
         Width           =   645
      End
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
         Left            =   5880
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1180
         Width           =   1650
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
         Left            =   6195
         TabIndex        =   43
         Top             =   120
         Width           =   1170
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
         Left            =   7350
         TabIndex        =   42
         Top             =   120
         Width           =   180
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00004000&
         Height          =   9000
         Left            =   10965
         ScaleHeight     =   8940
         ScaleWidth      =   1005
         TabIndex        =   31
         Top             =   0
         Width           =   1065
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
            Left            =   105
            Picture         =   "FOCAMOD07.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
            Top             =   4155
            Width           =   650
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "List"
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
            Picture         =   "FOCAMOD07.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Consultas y Listados"
            Top             =   4845
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
            Left            =   105
            Picture         =   "FOCAMOD07.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   38
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
            Left            =   105
            Picture         =   "FOCAMOD07.frx":0D60
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Acceso a Base de Datos de Proveedores"
            Top             =   1650
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Art."
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
            Picture         =   "FOCAMOD07.frx":106A
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   2340
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
            Left            =   105
            Picture         =   "FOCAMOD07.frx":1374
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Configuración de Funcionamiento de Compras"
            Top             =   3255
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
            Left            =   105
            Picture         =   "FOCAMOD07.frx":17B6
            Style           =   1  'Graphical
            TabIndex        =   34
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
            Picture         =   "FOCAMOD07.frx":1900
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Aprobar Ordenes de Compra"
            Top             =   5535
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
            Height          =   1005
            Left            =   105
            Picture         =   "FOCAMOD07.frx":1C0A
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Aprueba, Graba e Imprime Orden de Compra"
            Top             =   6405
            Width           =   650
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -315
            Picture         =   "FOCAMOD07.frx":1F14
            Top             =   7500
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
         Left            =   1155
         TabIndex        =   30
         ToolTipText     =   "Lista de Proveedores con Entregas Programadas"
         Top             =   870
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Frame Frame15 
         Caption         =   "Frame15"
         Height          =   2250
         Left            =   10450
         TabIndex        =   29
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2250
         Index           =   6
         Left            =   8085
         TabIndex        =   28
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2250
         Index           =   4
         Left            =   8820
         TabIndex        =   27
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2280
         Index           =   0
         Left            =   2050
         TabIndex        =   26
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0FF&
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
         Left            =   7665
         TabIndex        =   22
         Top             =   840
         Width           =   3060
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
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   " "
            Top             =   280
            Width           =   2010
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
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   " "
            Top             =   280
            Width           =   435
         End
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
            TabIndex        =   23
            Top             =   280
            Width           =   150
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Vigencia"
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
         Height          =   800
         Left            =   7680
         TabIndex        =   19
         Top             =   15
         Width           =   1380
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
            Height          =   285
            Left            =   1050
            TabIndex        =   21
            Top             =   420
            Width           =   195
         End
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
            Height          =   285
            Left            =   105
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   " "
            Top             =   420
            Width           =   960
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
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   18
         Text            =   " "
         Top             =   1180
         Width           =   4065
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
         Height          =   1860
         Left            =   160
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   105
         Width           =   850
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   160
         ScaleHeight     =   375
         ScaleWidth      =   10500
         TabIndex        =   13
         Top             =   1575
         Width           =   10565
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Código          Descripción                                        U/S   Pr.Unitario "
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
            TabIndex        =   14
            ToolTipText     =   "Haga Doble-Click para Blanquear Pre-Carga y Habilitar Carga Manual"
            Top             =   105
            Width           =   10515
         End
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
         Left            =   9450
         TabIndex        =   7
         Top             =   105
         Width           =   810
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
         Left            =   9615
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   525
         Width           =   960
      End
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
         Left            =   10560
         TabIndex        =   6
         Top             =   525
         Width           =   195
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
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   770
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
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   770
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
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   440
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
         TabIndex        =   1
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
         Height          =   285
         Left            =   920
         TabIndex        =   0
         Top             =   870
         Width           =   195
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   10290
         X2              =   10395
         Y1              =   245
         Y2              =   245
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   10450
         TabIndex        =   50
         ToolTipText     =   "Número de Revisión"
         Top             =   105
         Width           =   315
      End
      Begin VB.Label LText1 
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
         Left            =   165
         TabIndex        =   49
         Top             =   870
         Width           =   750
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
         Left            =   5670
         TabIndex        =   48
         ToolTipText     =   "Dobe-Click para Asignar Sector por Item"
         Top             =   510
         Width           =   1065
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
         Left            =   5670
         TabIndex        =   44
         ToolTipText     =   "Dobe-Click para Asignar Destino por Item"
         Top             =   180
         Width           =   1100
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
         Left            =   165
         TabIndex        =   41
         Top             =   1185
         Width           =   1170
      End
      Begin VB.Line Line5 
         X1              =   180
         X2              =   10720
         Y1              =   3850
         Y2              =   3850
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
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
         TabIndex        =   17
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
         Caption         =   "C.U.I.T.:"
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
         TabIndex        =   12
         Top             =   945
         Width           =   1695
      End
      Begin VB.Line Line8 
         X1              =   10725
         X2              =   10725
         Y1              =   1575
         Y2              =   3860
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
         Left            =   8255
         TabIndex        =   11
         Top             =   180
         Width           =   1170
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
         Left            =   8820
         TabIndex        =   10
         Top             =   550
         Width           =   750
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
         TabIndex        =   9
         Top             =   630
         Width           =   960
      End
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
         Enabled         =   0   'False
      End
      Begin VB.Menu plh4 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Orcom 
      Caption         =   "Transacciones"
      Begin VB.Menu Nue_Orco 
         Caption         =   "Nueva Orden de Compra"
         Begin VB.Menu Directa 
            Caption         =   "O/C Directa"
         End
         Begin VB.Menu Por_MRP 
            Caption         =   "O/C por MRP"
         End
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
      Caption         =   "Configuracion "
      Begin VB.Menu Setup 
         Caption         =   "Configuración General"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu Nuorig1 
         Caption         =   "Números Originales"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FOCAMOD07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$
Dim MUECOPEN%
Dim EMICONSI%
Dim MODOULPRE%
Dim RAI$

Private Sub Anul_Ocom_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call OPOCAB07.Command14_Click
End Sub

Private Sub Command1_Click()
   Call MENSERR(24, "No Válido Cambiar Proveedor")
End Sub

Private Sub Command10_Click()
   Call MENSERR(24, "No Válido Cambiar Proveedor")
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07/AMAPRO"
   Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")
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

Private Sub Command14_Click()
   Command14.Enabled = False
   '
10 FOSERE07.Label3 = Text11 '\\**//
   Call SELEC_REGLACOM(OKX%)
   If OKX% < 1 Then
     Command14.Enabled = True
     Exit Sub
   End If
   '
    NCI% = Val(LText1)
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
20 Call GRACOMPRA(OKX%)
   If OKX% > 0 Then
     Call PRICOMPRA
     If RAI$ <> "" Then
       Call FOROCAB07.PRIREGLACOM 'IMPRIME REGLAMENTO DE COMPRA
     End If
     Call BLANFORMU
     Call SETULTREG("!OCOMDETA", 0)
     Call CIERRARCH("*.*")
   End If
   Command14.Enabled = True
   Call CIERRARCH("*.*")
   '
   'AGREGADO PARA ENVIAR POR MAIL
   If SENDBYMAIL$ <> "" Then
     Call RECONECTA
     Exit Sub
   End If
   'HASTA ACA AGREGADO
   Unload Me
   '
End Sub


Private Sub Command15_Click()
   Command15.Enabled = False
   FORCOLIB07.Show 1
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
    Unload Me
End Sub

Private Sub Command2_LostFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
  FAPOCAB07.Show 1
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
    Call HELPONLINE("PRESUP01")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
     'agrego todas las O/C de ocomenca
      Call ABRECOMPRAS
      SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM ENCOCABI"
      SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  DESC "
      'Set VEROCTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
      Dim VEROCTEMP As ADODB.Recordset
      Set VEROCTEMP = New ADODB.Recordset
      VEROCTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
      With VEROCTEMP
        On Error Resume Next ' si hay algun error msgbox err
         .MoveFirst
         If Err Then
           Call MENSERR(24, "No hay Ordenes de Compra Abiertas Emitidas")
           On Error GoTo 0
           GoTo 99
         End If
        On Error GoTo 0
         '
         JJ& = 0
         Do While (.EOF = False)
             XXY$ = REGBLAN$("INDIOCOM")
             Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4)
             Call REPLA(XXY$, !Refe_OcoM, 5, 44)
             JJ& = JJ& + 1
             Call GRAREG("!INDIOCOM", XXY$, JJ&)
             .MoveNext
         Loop
      End With
      Set VEROCTEMP = Nothing
      Call SETULTREG("!INDIOCOM", JJ&)
      Call CIERRARCH("!INDIOCOM")
      '
3     Call SELECHO("!INDIOCOM/Indice General de Ordenes de Compra Abiertas (Pendientes-Cumplidas-Anuladas)")
      
      NPX& = Val(VALACT1$("!INDIOCOM"))
      If NPX& > 0 Then
        Call SHOWOCOM(NPX&)
        GoTo 3
      End If
      '
99  Command4.Enabled = True
   End Sub

Sub Command5_Click()
    Command5.Enabled = False
    Call CIERRARCH("*.*")
    If DERACCE%("STPOCOM", "Configuracion de Funcionamiento Compras") = 2 Then
      Call CONECRUN("FSETUP04/SETUPCOM", "Configuración de Funcionamiento")
      Call FINAL("")
    End If
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   LISOCAB07.Show 1
   Command6.Enabled = True
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
    ACONEC$ = "ARCHIG07/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Call MENSERR(24, "No Válido Cambiar Fecha de Emisión")
End Sub

Private Sub Desanuoc_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call OPCICOM07.Command10_Click
End Sub

Private Sub DevoConsi_Click()
   OPCICOMP07.Command15_Click
End Sub

Private Sub Directa_Click()
   Call Command1_Click
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub


Private Sub Entrecon_Click()
   OPCICOM07.Command2_Click
End Sub

Private Sub Exit_Click()
   Call Command2_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  While WindowState <> 0
    WindowState = 0
    Openforms = DoEvents
    AppActivate Caption
    SendKeys "%" & Chr$(32) & Chr$(13), True
  Wend
End Sub

Private Sub Form_Load()
    '
    If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
      Nuorig.Enabled = True
      Nuorig1.Enabled = True
    End If
    '
    MODOULPRE% = 1
    If CONTROL("OCOMPRA", "PRECIDEF") = "COSTOREP" Then MODOULPRE% = 2
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
   Call LIBARCH("*.*")
End Sub

 Sub Label11_DblClick()
   '
   Static CAMON%, DMONE$(), TCMONE()
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
     If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
           DMONE$(U&) = TRIM$(DMX$)
           TCMONE(U) = VHX
         Next U&
     End If
   End If
   '
   MONEMIA% = MONEMI%
15 MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   If DMONE$(MONEMI%) = "" Then GoTo 15
   Label11 = DMONE$(MONEMI%)
   Label11.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      Label11.ForeColor = RGB(255, 0, 0)
   End If
   Label11.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"

   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBX# = TICAMONE(MONEMI%)
   '\\\OT-06-0149-WAR-19/04/06///
   Text122 = FORMATNUM(TICAMBX#, "F11.4")
   '\\\OT-06-0149-WAR-FIN///
   '
   For U& = 1 To ULTREG&("!OCOMDETA")
     YX$ = REGLEIDO$("!OCOMDETA", U&)
     PREDOL# = CVD(Mid$(YX$, 67, 8)) * COEFCON
     Call REPLA(YX$, MKD$(PREDOL#), 67, 8)
     Call GRAREG("!OCOMDETA", YX$, U&)
   Next U&
   Call CIERRARCH("!OCOMDETA")
   Call CARLISCO
   Call CALTOCOM
   '
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

Private Sub Label2_DblClick()
   Call Label14_dblClick
End Sub

Private Sub Label22_DblClick()
   Call CARDETCOM
End Sub

Private Sub Label8_Click()
   '
   Call MENSERR(24, "No Válido Cambiar Proveedor")
   Exit Sub
   '
   Call SELECHO("INDIACRE")
   CPROX$ = TRIM$(VALACT1$("INDIACRE"))
   If CPROX$ <> "" Then
     NCLIE% = NUMEPRO%(CPROX$)
     If NCLIE% > 0 Then
       If REGICLI&((-1) * NCLIE%) > 0 Then
         Call BLANFORMU
         Text1.TEXT = TRIM$(Str$(NCLIE%))
         Call CARDACLI
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

'Private Sub Ltota_DblClick(Index As Integer)
'   For KKI% = 0 To 4
'     Total(KKI%).Enabled = Not (Total(KKI%).Enabled)
'   Next KKI%
'End Sub

Private Sub LText1_Change()
   Call CARDACLI
End Sub

Private Sub MarItCum_Click()
  Call OPCOCAB07.Command4_Click
End Sub

Private Sub Amastock_Click()
   Call Command8_Click
End Sub

Private Sub Nuorig_Click()
   Call OPCICOM07.Command19_Click
End Sub

Private Sub Nuorig1_Click()
   Call OPCICOM07.Command19_Click
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
   LISOCAB07.Option2.Value = True
   Call LISOCAB07.Command3_Click
End Sub

Private Sub Print2_Click()
   LISOCAB07.Option2.Value = True
   Call LISOCAB07.Command4_Click
End Sub

Private Sub Print3_Click()
   LISOCAB07.Option2.Value = True
   Call LISOCAB07.Command5_Click
End Sub

Private Sub Print4_Click()
   LISOCAB07.Option2.Value = True
   Call LISOCAB07.Command9_Click
End Sub

Private Sub Provedor_Click()
   Command11_Click
End Sub

Private Sub Screen1_Click()
   LISOCAB07.Option1.Value = True
   Call LISOCAB07.Command3_Click
   
End Sub

Private Sub Screen2_Click()
   LISOCAB07.Option1.Value = True
   Call LISOCAB07.Command4_Click
End Sub

Private Sub Screen3_Click()
   LISOCAB07.Option1.Value = True
   Call LISOCAB07.Command5_Click
End Sub

Private Sub Screen4_Click()
   LISOCAB07.Option1.Value = True
   Call LISOCAB07.Command9_Click
End Sub

Private Sub Setup_Click()
  OPOCAB07.Command18_Click
End Sub

Private Sub Text11_GotFocus()
   Text6.SetFocus
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



Private Sub Text16_Click()
   On Error Resume Next
   Text6.SetFocus
   On Error GoTo 0
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

Private Sub Text2_GotFocus()
   On Error Resume Next
   Text6.SetFocus
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
   Text6.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text4_GotFocus()
   On Error Resume Next
   Text6.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text5_GotFocus()
   On Error Resume Next
   Text6.SetFocus
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
   If Val(LText1) > 0 Then
     If Val(LText1) <= 32767 Then
       NCLIE% = Val(LText1)
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
'   Text12.TEXT = FORMATNUM$(ALIVA, "F5.2")
   Text13.TEXT = ""
   Text6.TEXT = ""
   Text7.TEXT = ""
   Text18.TEXT = ""
   List1.Clear: List1.ToolTipText = ""
   List2.Clear
   '
   TEXT21(0).TEXT = "": Eco21(0) = ""
   TEXT21(1).TEXT = "": Eco21(1) = ""
   'TEXT21(2).TEXT = "": Eco21(2) = ""
   '
'   PORDESCU.TEXT = ""
   Text17 = ""
'   For KKI% = 0 To 4
'      Total(KKI%).Enabled = True
'      Total(KKI%).TEXT = ""
'   Next KKI%
   '
   'MARCONOTA.Visible = False
   HOII% = HOY(HOS$)
   Text9.TEXT = ""
   Text11.TEXT = ""
   '
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    '
    NC% = (-1) * Val(LText1)
    If NC% < 0 Then
      If STATCLI%(NC%) > 0 Then
        Label8 = CODICLI$(NC%)
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        'Label2 = TRIM$(POSIVAC$(NC%))
        Text8.TEXT = TRIM$(CUITCLI$(NC%))
      End If
    End If
    '
    Screen.MousePointer = 1
    '
End Sub
   '
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
   Text6.SetFocus
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
   End If
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
   'COPRECIO = 1
   'If Total(0).Enabled = False Then COPRECIO = 0
   '
   Call CIERRARCH("OCOMENCA")
   Call CIERRARCH("OCOMDETA")
   '
   Call BLOQARCH("OCOMENCA")
   Call BLOQARCH("OCOMDETA")
   Call BLOQARCH("MACONSIG")
   '
   NRO& = Val(Text11)
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
   NPRO% = Val(LText1)
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
      Text17 = ""
    End If
    '
    Call ABRECOMPRAS
    '*********************
     Dim ENKOKABI As ADODB.Recordset
     Set ENKOKABI = New ADODB.Recordset
     SQLX$ = "SELECT * FROM EncoCabi "
     SQLX$ = SQLX$ + "WHERE Nume_Ocom = " & NRO& & " "
     SQLX$ = SQLX$ + "AND Stat_Ocom < 9"
     ENKOKABI.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     'With EncoCabi
     '.FindFirst "Nume_Ocom = " & NRO& & "AND Stat_Ocom < 9 "
     'If .NoMatch = True Then
         'MERRO$ = "No se Ha Encontrado Registro\de Encabezado de esta O/Compra."
         'GoTo 98
     'End If

     With ENKOKABI
       On Error Resume Next
       .MoveFirst
       If Err > 0 Then
          On Error GoTo 0
          MERRO$ = "No se Ha Encontrado Registro\de Encabezado de esta O/Compra."
          GoTo 98
        Else
          On Error GoTo 0
       FEX = FEMI%
       REFE$ = FECHATEX$(FEX) + " - " + RASOCLI$((-1) * NPRO%)
       !Refe_OcoM = Left$(REFE$, 64)
       !Femi_Ocom = CVDAT(FEMI%)
       !Npro_Ocom = NPRO%
       !RASO_PROV = Left$(RASOCLI$((-1) * NPRO%), 48)
       !TSum_Ocom = 0               ' materiales
       '
       !FMod_Ocom = Now
       !Usua_Mod = USNBR%
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
       '\\\OT-06-0082-WAR-06/03/06///
       !OCOMREVIS = XVALO(Label16)
       '\\\OT-06-0082-WAR-FIN///
       On Error GoTo 0
       !AtSr_Ocom = Left$(TRIM$(Text6), 64)
       !ToBru_Ocom = TOTANE#
       !Pdes_Ocom = PORDESCU
       !IDes_Ocom = IDESCUEN#
       !Iiva_Ocom = TOTIVA#
       !Iotr1_Ocom = TOTPER#
       !Total_Ocom = TOTATO#
       !Lentrega = Left$(Text10, 64)
       !Obse_Ocom = Text13
     If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) <> "DIRECT" Then
       !Stat_Ocom = (-1)       ' status Pendiente de aprobacion
     End If
       On Error Resume Next
       !Anex_Ocom = RAI$
       On Error GoTo 0
       '
       .Update
       End If
    End With
    Set ENKOKABI = Nothing

    '
    Label13 = RAI$
    ' ANULAR ITEMS
    
'    SQLX$ = "SELECT * FROM DETOCABI "
'    SQLX$ = SQLX$ + "WHERE Nume_Ocom = " & NRO& & " "
'    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
    'Set COMDET = Compras.OpenRecordset(SQLX$, dbOpenDynaset)
    SQLX$ = "UPDATE DETOCABI SET Stat_Ocom = 9"
    SQLX$ = SQLX$ + "WHERE Nume_Ocom = " & NRO& & " "
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
     FLEXCONN.Execute (FILTSQL$(SQLX$))
    '
'    With COMDET
'      On Error Resume Next
'      .MoveFirst
'      If Err Then
'        On Error GoTo 0
'        GoTo 126
'      End If
'      On Error GoTo 0
'      Do While Not .EOF
'        .Edit
'        !Stat_Ocom = 9
'        .Update
'      .MoveNext
'      Loop
'    End With
    '
126 EMICONSI% = 0
    IDENOC$ = MKS$(NRO&) + MKI$(FEMI%) + MKI$(NPRO%)
    '
    Dim DETOKABI As ADODB.Recordset
    Set DETOKABI = New ADODB.Recordset
    SQLX$ = "SELECT * FROM DetoCabi where Nume_Ocom = " & NRO&
    DETOKABI.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        '
    ORDITMA% = 0
    ARCHILIB$ = "!1OCOMDET"
    For J& = 1 To ULTREG&("!OCOMDETA")
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
      With DETOKABI
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
         !Nume_Ocom = NRO&
           ORDIT% = CVS(Mid$(XX$, 125, 4))
           If ORDIT% < 1 Then ORDIT% = ORDITMA% + 1
           If ORDIT% > ORDITMA% Then ORDITMA% = ORDIT%
         !NuOrd_Item = ORDIT%
         !Femi_Ocom = CVDAT(FEMI%)
         !Npro_Ocom = NPRO%
         !RASO_PROV = Left$(RASOCLI$((-1) * NPRO%), 64)
         !Cod_Inte = CI0%
         !COD_EXTE = CODEXT$(CI0%)
         !Descritem = DESCRIT0$(CI0%)
         !Delibre = ""
         If COLTELIB% >= 0 Then
           !Delibre = TRIM$(REGLEIDO$(ARCHILIB$, J&))
         End If
         If CI0% < 1 Then
           If TRIM$(!Delibre) <> "" Then
             !Descritem = Left$(!Delibre, 64)
           End If
         End If
         !Valo_Unid = VAUNI$
         !Uni_Mas = Mid$(ZZ$, 11, 4)
         !Unids_Ocom = CVS(Mid$(ZZ$, 51, 4))
           CANTU = CVS(Mid$(ZZ$, 55, 4))
         !Cant_Ocom = CANTU
         !Uni_Com = Mid$(ZZ$, 59, 4) ' unidad de compra
         !Coef_Unids = COEFUNI
         !Pre_Unit = CVD(Mid$(ZZ$, 67, 8))
         !Desc_Item = CVS(Mid$(ZZ$, 75, 4))
         !PrUn_Neto = CVD(Mid$(ZZ$, 79, 8))
         !INet_Item = CVD(Mid$(ZZ$, 87, 8))
         !FVig_Ocom = CVDAT(CVI(Mid$(ZZ$, 95, 2)))
         !CONSIMAT = TRIM$(Mid$(ZZ$, 97, 2))
            CFXUNI# = !Coef_Unids
            If CFXUNI# < 0.005 Then CFXUNI# = 1
         !Cant_Rec = CANTORECI(CI0%, NRO&) / CFXUNI#
         !Cod_Oper = 0
         !Doc_Orig = ""
         !N_OrServ = 0
         !Mone_Cos = MONECOI%        ' antes pos 83,1
         !Mone_Emis = MONEMI%        ' antes pos 84,1
         !TSum_Ocom = 0              ' materiales
            STATUOC% = Val(Label15)
            If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) <> "DIRECT" Then STATUOC% = (-1)
         !Stat_Ocom = STATUOC%       ' status Pendiente de aprobacion
            NUECUEIMPO$ = TRIM$(Mid$(ZZ$, 105, 12))
         !IContab = CUECOINT(TRIM$(NUECUEIMPO$)) ' Imputacion contable
            SECTOROC$ = TRIM$(Mid$(ZZ$, 15, 12))
         !Sect_Ocom = SECTOROC$   ' Sector
            '
        .Update
      End With
   Next J&
   '
   Call BAJALDISCO
   '
   If COPRECIO < 0.99 Then
     Text17 = ""
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
        PRHBL% = PROVHABI%(CI0%)
        If PRHBL% < 1 Or PRHBL% = NPRO% Or ACOSREP$ = "SIEMPRE" Then
            COSMONCO = COSPESOSD# / TICAMONE(MONECO%)
            If COSMONCO >= 0.000005 Or PRHBL% < 1 Then
              If ACOSREP$ <> "NUNCA" Then
                Call ABREMASTER
                'Master.FindFirst ("CODINT = " & CI0%)
                'If Master.NoMatch = False Then
                'Master.Edit
                Dim MASTTTER As ADODB.Recordset
                Set MASTTTER = New ADODB.Recordset
                SQLX$ = " SELECT * FROM Master "
                SQLX$ = SQLX$ + " WHERE CODINT = " & CI0% & ""
                MASTTTER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
                  With MASTTTER
                    On Error Resume Next
                    .MoveFirst
                    If Err < 1 Then
                      If COSMONCO >= 0.000005 Then
                        !COSUNI = COSMONCO
                        !FECOSTO = CVDAT(FEMI%)
                      End If
                      If PRHBL% < 1 Then
                        !PROVHABI = NPRO%
                      End If
                      !Monecos = MONECO%
                      .Update
                    End If
                    On Error GoTo 0
                  End With
                  Set MASTTTER = Nothing
                'End If
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
              'Master.FindFirst "CODINT =" & CI0%
              'If Master.NoMatch = False Then
                'Master.Edit
              'Dim MASTTTER As ADODB.Recordset
              Set MASTTTER = New ADODB.Recordset
              SQLX$ = "SELECT * FROM MASTER"
              SQLX$ = SQLX$ + "WHERE CODINT = " & CI0% & ""
              MASTTTER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
              With MASTTTER
                On Error Resume Next
                .MoveFirst
                On Error GoTo 0
                !PROVHABI = NPRO%
                COSMONCO = COSPESOSD# / TICAMONE(MONECO%)
                If COSMONCO >= 0.000005 Then
                  !COSUNI = COSMONCO
                  !FECOSTO = CVDAT(FEMI%)
                  !Monecos = MONECO%
                End If
                .Update
                On Error GoTo 0
                End With
                Set MASTTTER = Nothing
              'End If
              '
              ZZX$ = REGBLAN$("RECAMPH")
              Call REPLA(ZZX$, MKI$(CI0%), 1, 2)
              FEMI% = FECHANUM(Text9)
              Call REPLA(ZZX$, MKI$(FEMI%), 3, 2)
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
               'With Master
                 '.FindFirst "CodInt = " & CI0%
                 'If .NoMatch = False Then
                   '.Edit
                   '!PESUNI = PSXX
                   '.Update
               SQLX$ = " UPDATE Master SET  PESUNI = PSXX "
               SQLX$ = SQLX$ + " WHERE CodInt = " & CI0% & ""
                FLEXCONN.Execute (FILTSQL$(SQLX$))
                 'End If
               'End With
             End If
           End If
        End If
        '
      End If
    Next J&
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
    Exit Sub
    '
98  Call CIERRARCH("*.*")
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

'Private Sub Total_LostFocus(Index As Integer)
'   If Index = 0 Then
'     TTT1 = Val(Total(0).TEXT)
'     Total(0) = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
'   End If
'End Sub

Sub CARDETCOM()
   '
   '\\\OT-05-0690-WAR-24/10/05///
   'PRIMER SWTICH = SI FORMATO = DESTINO
   If CONTROL$("OCOMPRA", "MODOIMPU") = "DES" Then
     Call CARDESCOM
     Screen.MousePointer = 1
     Exit Sub
   End If
   '\\\OT-05-0690-WAR-FIN///
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
   PRF$ = "CARMODCO"
   If Val(Label15) < 0 Then ' no esta aprobada - puede editarse todo
     PRF$ = "CARDETCO"
   End If
   XXX% = VENTABU%(PRF$ + ATRI$)
   List1.Clear
   Call CALTOCOM
   Timer1.Enabled = False
   '
   If XXX% < 0 Or ULTREG&("!OCOMDETA") < 1 Then
      Call CIERRARCH("!OCOMDETA")
      Call BLANARCH("!OCOMDETA")
      Call BLANFORMU
'      Total(0).TEXT = ""
      LText1 = ""
      Hide
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
      If TRIM$(Mid$(ZZ$, 11, 4)) = "" Then Call REPLA(ZZ$, UNIMED$(CI0%), 11, 4)
      If TRIM$(Mid$(ZZ$, 59, 4)) = "" Then Call REPLA(ZZ$, UNICOM$(CI0%), 59, 4)
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
      PREUNID# = CVD(Mid$(ZZ$, 67, 8))
      If PREUNID# < 0.00005 Then
        If UC$ = "" Then UC$ = ULTICOM$(NPRO%, CI0%)
        MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
        PREUNID# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
        PORDESCUEI = CVS(Mid$(UC$, 75, 4))
        If PREUNID# < 0.00005 Or MODOULPRE% = 2 Then
          PREUNID# = COSUNI(CI0%) * TICAMONE(MONECOSTO%(CI0%)) * COEFI
          PORDESCUEI = 0
        End If
        PREUNID# = PREUNID# / TICAMONE(MONEMI%)
        Call REPLA(ZZ$, MKD$(PREUNID#), 67, 8)
        Call REPLA(ZZ$, MKS$(PORDESCUEI), 75, 4)
      End If
      '
      If ZZ$ <> Z0$ Then ATK% = 1
      Call REPLA(ZZ$, Text7, 105, 12)
      Call REPLA(ZZ$, Text18, 15, 12)
      Call GRAREG("!OCOMDETA", ZZ$, U&)
   Next U&
   '
   Screen.MousePointer = 1
   If ATK% = 1 Then GoTo 10
   '
   Call CARLISCO
   Call CALTOCOM
   Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0690-WAR-24/10/05///
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
   PRF$ = "CAMOCADE"
   If Val(Label15) < 0 Then ' no esta aprobada - puede editarse todo
     PRF$ = "CARDEOCA"
   End If
   XXX% = VENTABU%(PRF$ + ATRI$)
   List1.Clear
   Call CALTOCOM
   Timer1.Enabled = False
   '
   If XXX% < 0 Or ULTREG&("!OCOMDETA") < 1 Then
      Call CIERRARCH("!OCOMDETA")
      Call BLANARCH("!OCOMDETA")
      Call BLANFORMU
'      Total(0).TEXT = ""
      LText1 = ""
      Hide
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   NPRO% = Val(Text1)
   For U& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", U&)
      CI0% = CVI(Mid$(ZZ$, 9, 2))
      If TRIM$(Mid$(ZZ$, 11, 4)) = "" Then Call REPLA(ZZ$, UNIMED$(CI0%), 11, 4)
      CANSTOK = CVS(Mid$(ZZ$, 51, 4))
      CANCOMP = CANSTOK
      Call REPLA(ZZ$, MKS$(CANSTOK), 51, 4)
      Call REPLA(ZZ$, MKS$(CANCOMP), 55, 4)
      Call REPLA(ZZ$, MKS$(COEFI), 63, 4)
      '
      PREUNID# = CVD(Mid$(ZZ$, 67, 8))
      If PREUNID# < 0.00005 Then
        If CI0% > 0 Then    ' para que no lo traiga con material libre
          If UC$ = "" Then UC$ = ULTICOM$(NPRO%, CI0%)
          MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
          PREUNID# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
          PORDESCUEI = CVS(Mid$(UC$, 75, 4))
          If PREUNID# < 0.00005 Or MODOULPRE% = 2 Then
            PREUNID# = COSUNI(CI0%) * TICAMONE(MONECOSTO%(CI0%)) * COEFI
            PORDESCUEI = 0
          End If
          PREUNID# = PREUNID# / TICAMONE(MONEMI%)
          Call REPLA(ZZ$, MKD$(PREUNID#), 67, 8)
          Call REPLA(ZZ$, MKS$(PORDESCUEI), 75, 4)
        End If
      End If
      '
      Call GRAREG("!OCOMDETA", ZZ$, U&)
   Next U&
   '
   Screen.MousePointer = 1
   '
   Call CARLISCO
   Call CALTOCOM
   Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0690-WAR-FIN///
Sub CALTOCOM()
   '
   NC1% = (-1) * Val(LText1)
   TOTANE# = 0
   For J& = 1 To ULTREG&("!OCOMDETA")
       X$ = REGLEIDO$("!OCOMDETA", J&)
       CI0% = CVI(Mid$(X$, 9, 2))
       UNIMAS$ = Mid$(X$, 11, 4)
       CANTU = CVS(Mid$(X$, 55, 4))
       PREUNID# = CVD(Mid$(X$, 67, 8))
       PORDE = CVS(Mid$(X$, 75, 4))
       FENTRE% = CVI(Mid$(X$, 95, 2))
       '
       PREUNINE# = (Int(PREUNID# * 100 * (100 - PORDE) + 0.5)) / 10000
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
   Next J&
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
   'PORDESCUENTO$ = TRIM$(PORDESCU)
   'DESC$ = PORDESCUENTO$
   'IDESCUEN# = TOTANE# * COEFDESCU(DESC$) / 100
   'TOTANEDES# = TOTANE# - IDESCUEN#
   'TOTIVA# = 0
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
   
   TOTATO# = TOTANEDES# + TOTIVA# + TOTPER#
   '
   'Total(0) = FORMATNUM$(TOTANE#, "F11.2")
   'Total(1) = FORMATNUM$(IDESCUEN#, "F11.2")
   'Total(2) = FORMATNUM$(TOTIVA#, "F11.2")
   'Total(4) = FORMATNUM$(TOTPER#, "F11.2")
   'Total(3) = FORMATNUM$(TOTATO#, "F11.2")
   '
End Sub

Sub CARCOANT()
   '
   Screen.MousePointer = 11
   List2.Clear
   NPRO% = Val(LText1)
   NOCOA& = Val(Text11)
   HOU = HOY(HOS$)
   '
   Call ABRECOMPRAS
   SQLX$ = " SELECT * FROM DETOCABI "
   SQLX$ = SQLX$ + " WHERE Stat_Ocom < 8 "
   SQLX$ = SQLX$ + " AND Stat_Ocom >=0 "
   SQLX$ = SQLX$ + " AND Npro_Ocom = " & NPRO% & " "
   SQLX$ = SQLX$ + " AND Nume_Ocom <> " & NOCOA& & " "
   'Set CARCOANTMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim CARCOANTMP As ADODB.Recordset
   Set CARCOANTMP = New ADODB.Recordset
   CARCOANTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   '
   With CARCOANTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       Exit Sub
     End If
     On Error GoTo 0
     Do While Not .EOF
       NUORSE& = !N_OrServ
             If NUORSE& < 1 Then
         COEFCAN = !Coef_Unids
         If COEFCAN < 0.005 Then COEFCAN = 1
         CACOM = !Cant_Ocom * COEFCAN
         CAREC = !Cant_Rec * COEFCAN
         'CAPEN = CACOM - CAREC
         'If CAPEN / COEFCAN >= 0.05 Then
           CIXI% = !Cod_Inte
           NROPED& = !Nume_Ocom
           UNICO$ = !Uni_Com
           FEX = CVINT(!FVig_Ocom)
           ATRAU = 0
           If FEX < HOU Then
             ATRAU = (DIENTRE(FEX, HOU)) / 7
             If ATRAU > 99.9 Then ATRAU = 99.9
           End If
           ATRAX$ = FORMATNUM$(ATRAU, "F4.1")
           '
           LI2TEX$ = Space$(80)
           Call REPLA(LI2TEX$, CODEXT$(CIXI%), 1, 16)
           Call REPLA(LI2TEX$, DESCRIT0$(CIXI%), 17, 45)
             If CIXI% < 1 Then        ' MATERIAL LIBRE
               DECRI$ = TRIM$(!Delibre)
               If DECRI$ <> "" Then
                 Call REPLA(LI2TEX$, DECRI$, 17, 45)
               End If
             End If
           'Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN / COEFCAN), 4, 8, 1, 2), 51, 8)
           'Call REPLA(LI2TEX$, UNICO$, 60, 4)
           Call REPLA(LI2TEX$, FECHATEX$(FEX), 69, 8)
           'Call REPLA(LI2TEX$, ATRAX$, 74, 4)
           Call REPLA(LI2TEX$, FORMATNUM$(NROPED&, "I6"), 78, 6)
           List2.AddItem LI2TEX$
           '
         'End If
       End If
       .MoveNext
     Loop
   End With
   Screen.MousePointer = 1
   '
End Sub

Sub CARCOANTvie()
   '
   Screen.MousePointer = 11
   List2.Clear
   NPRO% = Val(LText1)
   NOCOA& = Val(Text11)
   HOU = HOY(HOS$)
   '
   Call ABRECOMPRAS
   SQLX$ = " SELECT * FROM OCOMDETA "
   SQLX$ = SQLX$ + " WHERE Stat_Ocom < 8 "
   SQLX$ = SQLX$ + " AND Stat_Ocom >=0 "
   SQLX$ = SQLX$ + " AND Npro_Ocom = " & NPRO% & " "
   SQLX$ = SQLX$ + " AND Nume_Ocom <> " & NOCOA& & " "
   'Set CARCOANTMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim CARCOANTMP As ADODB.Recordset
   Set CARCOANTMP = New ADODB.Recordset
   CARCOANTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
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
           FEX = CVINT(!Fentre_Sol)
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
             If CIXI% < 1 Then        ' MATERIAL LIBRE
               DECRI$ = TRIM$(!Delibre)
               If DECRI$ <> "" Then
                 Call REPLA(LI2TEX$, DECRI$, 17, 33)
               End If
             End If
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
   Set CARCOANTMP = Nothing
   Screen.MousePointer = 1
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
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!OCOMDETA", U&)
     CI% = CVI(Mid$(X$, 9, 2))
     FENTRESOL% = CVI(Mid$(X$, 95, 2))
     If FENTRESOL% < FENTREGEN% Then FENTRESOL% = FENTREGEN%
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
     '\\\OT-05-0690-WAR-24/10/05///
     Call REPLA(TTXX$, DESCRIT0$(CI%), 17, 65)
     If CI% <= 0 Then
       DESCRILI$ = TRIM$(REGLEIDO$("!1OCOMDET", U&))
       Call REPLA(TTXX$, DESCRILI$, 17, 15)
     End If
     '\\\OT-05-0690-WAR-FIN///
     Call REPLA(TTXX$, UNISTO$, 68, 4)
'     Call REPLA(TTXX$, FORMATNUM$(CANSTO, "F8.1"), 38, 8)
'     Call REPLA(TTXX$, UNICOX$, 48, 4)
'     Call REPLA(TTXX$, FORMATNUM$(CANCOM, "F9.1"), 52, 9)
     MASCA$ = "F10.4"
     If PREUNI > 9999 Then MASCA$ = "F10.2"
     Call REPLA(TTXX$, FORMATNUM$(PREUNI, MASCA$), 75, 10)
'     Call REPLA(TTXX$, FORMATNUM$(PORDESCUIN, "F5.1"), 72, 5)
'     Call REPLA(TTXX$, FORMATNUM$(FENTRESOL%, "D8"), 78, 8)
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
   Next U&
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
   FULCO% = 0
   UCO$ = REGBLAN$("OCOMDETA")
   Call REPLA(UCO$, MKI$(NPRO%), 7, 2)
   Call REPLA(UCO$, MKI$(CIIX%), 9, 2)
   '
   SQLX$ = "SELECT * FROM OCOMDETA "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIIX% & " "
   SQLX$ = SQLX$ + "AND Npro_Ocom = " & NPRO% & " "
   SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
   SQLX$ = SQLX$ + "AND Pre_Unit >= 0.00005 "
   SQLX$ = SQLX$ + "ORDER BY Femi_Ocom ASC"
   'Set OcomdetaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim OcomdetaTemp As ADODB.Recordset
   Set OcomdetaTemp = New ADODB.Recordset
   OcomdetaTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   On Error Resume Next
   With OcomdetaTemp
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
     End If
   End With
   Set OcomdetaTemp = Nothing
   '
   ULTICOM$ = UCO$
   '
End Function

Sub SHOWOCOM(NPX&)
    '
    TIDOCUM$ = "Nota de Pedido"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU1$ = " " + NUDOCU$ + " "
    NUDOCU2$ = " " + NUDOCU$ + "-"
    '
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
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
      '\\\OT-06-0082-WAR-06/03/06///
       NUMEDOC$ = NUMEDOC$ + "-" + TRIM$(Label16)
      '\\\OT-06-0082-WAR-FIN///
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
      'DOCPARAM$(3) = Total(0)
      DOCPARAM$(4) = Text12
      'DOCPARAM$(5) = Total(2)
      'DOCPARAM$(6) = Total(3)
      DOCPARAM$(7) = "" 'Mid$(X1$, 147, 20)
      DOCPARAM$(8) = "" 'Mid$(X1$, 67, 40)
      DOCPARAM$(9) = "" 'Mid$(X1$, 167, 40)
      DOCPARAM$(10) = "" ' Mid$(X1$, 107, 40)
      'DOCPARAM$(11) = Str$(Val(Total(0)) - Val(Total(1)))
      DOCPARAM$(12) = TRIM$(Text6) ' ATENCION SR.
      'DOCPARAM$(13) = TRIM$(PORDESCU)
      'DOCPARAM$(14) = Total(1)
      HOII% = HOY(HOS$)
      FECHENT% = FECHANUM(Text14)
      DOCPARAM$(15) = Text14
      If FECHENT% <= HOII% Then
          DOCPARAM$(15) = DOCPARAM$(15) + " (inmediato)"
        ElseIf FECHENT% < HOII% + 4 Then
          DOCPARAM$(15) = DOCPARAM$(15) + " (urgente!!)"
      End If
      DOCPARAM$(16) = TRIM$(Label11)
      DOCPARAM$(17) = ""
      If MONEMI% > 1 Then DOCPARAM$(17) = Str$(TICAMONE(MONEMI%))
      DOCPARAM$(18) = Eco21(1)
      DOCPARAM$(19) = Text10
      'DOCPARAM$(20) = Total(4)
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
      DESTIDOC% = (-1) * Val(LText1)
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
         PREUNID# = CVD(Mid$(X$, 67, 8))
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
           '
         TETABU1$(3, CAITAB1%) = Mid$(X$, 59, 4) 'UNIDAD DE COMPRA
         If TRIM$(Mid$(X$, 59, 4)) = "" Then TETABU1$(3, CAITAB1%) = Mid$(X$, 11, 4) 'UNIDAD DE STOCK
         '
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 97, 2)  'consigna m.prima
         TETABU1$(7, CAITAB1%) = Str$(PREUNID#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PREUNID# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PREUNID# * (1 - PORDE / 100))
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
           If CI0% > 0 Then  ' que no lo haga en texto libre
             DEADI$ = "(equiv. " + TRIM$(FORMATNUM$(CVS(Mid$(X$, 51, 4)), "F12.1")) + " " + TRIM$(ECOARCH$("UNIMED", UNIST$)) + ".)"
             CAITAB1% = CAITAB1% + 1
             TETABU1$(2, CAITAB1%) = DEADI$
           End If
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
    Exit Sub
    '
999 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub

Private Sub UcoMat_Click()
Call LISOCAB07.Command20_Click
End Sub

Private Sub UcoPro_Click()
Call LISOCAB07.Command12_Click
End Sub

Private Sub Ver_Ocom_Click()
    Call LISOCAB07.Command22_Click
End Sub

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
'      Total(0).TEXT = ""
      'Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
End Sub

Function CANTORECI(CIXI%, NUOCO&)
    '
    CANTO = 0
    '
    Call ABRESTOCKS
    SQLX$ = "SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
    SQLX$ = SQLX$ + "AND NumOC = " & NUOCO& & " "
    'Set BOLDET = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim BOLDET As ADODB.Recordset
    Set BOLDET = New ADODB.Recordset
    BOLDET.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    With BOLDET
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        GoTo 126
      End If
      On Error GoTo 0
      Do While Not .EOF
        CANTO = CANTO + !CantNom - CanToRech - CanToFalta
      .MoveNext
      Loop
    End With
    Set BOLDET = Nothing
    '
126 CANTORECI = CANTO
    '
End Function

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
10 RAI$ = "-" + Label13 + "-"
   With FOSERE07
     .LIMATSEL.Clear
     FIN& = ULTREG&("INDIRCOM")
     For J& = 1 To FIN&
       X$ = REGLEIDO$("INDIRCOM", J&)
       NR& = CVI(Mid$(X$, 1, 2))
       DRC$ = Mid$(X$, 3, 46)
       RC$ = FORMATNUM(NR&, "F5.0") & Chr(9) & DRC$
       .LIMATSEL.AddItem RC$
       NR1$ = "-" + TRIM$(Str$(NR&)) + "-"
       If InStr(RAI$, NR1$) > 0 Then
         .LIMATSEL.Selected(J& - 1) = True
       End If
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
