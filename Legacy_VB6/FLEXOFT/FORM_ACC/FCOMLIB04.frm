VERSION 5.00
Begin VB.Form FCOMLIB04 
   BackColor       =   &H00E0E0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Compra - Proveedores Externos"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "FCOMLIB04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame9 
         Caption         =   "NRO DE PLANO"
         Height          =   2535
         Left            =   600
         TabIndex        =   90
         Top             =   240
         Visible         =   0   'False
         Width           =   4395
         Begin VB.Label Label32 
            Caption         =   "NRO. DE ORDEN"
            Height          =   435
            Left            =   105
            TabIndex        =   102
            Top             =   2100
            Width           =   1590
         End
         Begin VB.Label Label31 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LABEL28"
            Height          =   330
            Left            =   1680
            TabIndex        =   101
            Top             =   1995
            Width           =   1905
         End
         Begin VB.Label Label30 
            Caption         =   "CODIGO EXTERNO"
            Height          =   225
            Left            =   420
            TabIndex        =   100
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label Label29 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LABEL29"
            Height          =   330
            Left            =   1680
            TabIndex        =   99
            Top             =   1365
            Width           =   1905
         End
         Begin VB.Label Label28 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LABEL28"
            Height          =   330
            Left            =   1680
            TabIndex        =   98
            Top             =   1680
            Width           =   1905
         End
         Begin VB.Label Label27 
            Caption         =   "NRO DE PLANO"
            Height          =   435
            Left            =   525
            TabIndex        =   97
            Top             =   1680
            Width           =   1170
         End
         Begin VB.Label Label26 
            Caption         =   "Nro de Orden de Compra"
            Height          =   435
            Left            =   525
            TabIndex        =   96
            Top             =   945
            Width           =   1170
         End
         Begin VB.Label Label25 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label25"
            Height          =   330
            Left            =   1680
            TabIndex        =   95
            Top             =   945
            Width           =   1905
         End
         Begin VB.Label Label24 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label24"
            Height          =   330
            Left            =   1680
            TabIndex        =   94
            Top             =   630
            Width           =   1905
         End
         Begin VB.Label Label23 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label23"
            Height          =   330
            Left            =   1680
            TabIndex        =   93
            Top             =   315
            Width           =   1905
         End
         Begin VB.Label Label20 
            Caption         =   "Nro de Grupo"
            Height          =   225
            Left            =   420
            TabIndex        =   92
            Top             =   630
            Width           =   1275
         End
         Begin VB.Label Label19 
            Caption         =   "Nro de EV"
            Height          =   225
            Left            =   630
            TabIndex        =   91
            Top             =   315
            Width           =   1275
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   160
         ScaleHeight     =   375
         ScaleWidth      =   10485
         TabIndex        =   67
         Top             =   1680
         Width           =   10550
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
            TabIndex        =   68
            Top             =   105
            Width           =   10515
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
         Height          =   2640
         Left            =   1470
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   2115
         Width           =   7890
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0F0&
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
         Height          =   1020
         Left            =   160
         TabIndex        =   50
         Top             =   4830
         Width           =   10560
         Begin VB.CommandButton Command15 
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
            Left            =   1680
            TabIndex        =   89
            Top             =   600
            Width           =   180
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
            Left            =   3675
            TabIndex        =   88
            Top             =   600
            Width           =   180
         End
         Begin VB.TextBox Text22 
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
            Left            =   1155
            TabIndex        =   86
            Text            =   " "
            Top             =   600
            Width           =   540
         End
         Begin VB.TextBox Text20 
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
            Left            =   315
            TabIndex        =   52
            Text            =   " "
            Top             =   600
            Width           =   540
         End
         Begin VB.TextBox Text19 
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
            Left            =   8860
            TabIndex        =   55
            Top             =   600
            Width           =   1500
         End
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
            Height          =   570
            Left            =   4200
            MultiLine       =   -1  'True
            TabIndex        =   54
            Top             =   315
            Width           =   4320
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
            Left            =   2205
            TabIndex        =   53
            Top             =   600
            Width           =   1485
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "U-M:"
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
            Left            =   945
            TabIndex        =   87
            Top             =   375
            Width           =   540
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Pr. Unitario:"
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
            Left            =   8715
            TabIndex        =   83
            Top             =   375
            Width           =   1485
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
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
            Left            =   3045
            TabIndex        =   82
            Top             =   315
            Width           =   1065
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Destino:"
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
            Left            =   1995
            TabIndex        =   81
            Top             =   375
            Width           =   1065
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant:"
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
            Left            =   105
            TabIndex        =   80
            Top             =   370
            Width           =   540
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00E0E0F0&
         Caption         =   "CUIT"
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
         Height          =   900
         Left            =   5670
         TabIndex        =   47
         Top             =   580
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
            TabIndex        =   48
            TabStop         =   0   'False
            Text            =   " "
            Top             =   315
            Width           =   1695
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Left            =   105
            TabIndex        =   49
            Top             =   630
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00004000&
         Height          =   9000
         Left            =   10965
         ScaleHeight     =   8940
         ScaleWidth      =   1005
         TabIndex        =   46
         Top             =   0
         Width           =   1065
         Begin VB.CommandButton Command14 
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
            Picture         =   "FCOMLIB04.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Aprueba, Graba e Imprime Orden de Compra"
            Top             =   6720
            Width           =   650
         End
         Begin VB.CommandButton Command22 
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
            Picture         =   "FCOMLIB04.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Anotador de Ordenes de Compra"
            Top             =   5180
            Width           =   650
         End
         Begin VB.CommandButton Command2 
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
            Picture         =   "FCOMLIB04.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   105
            Width           =   650
         End
         Begin VB.CommandButton Command5 
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
            Picture         =   "FCOMLIB04.frx":0BA0
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Configuración de Funcionamiento de Compras"
            Top             =   3045
            Width           =   650
         End
         Begin VB.CommandButton Command8 
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
            Height          =   630
            Left            =   105
            Picture         =   "FCOMLIB04.frx":0FE2
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   2200
            Width           =   650
         End
         Begin VB.CommandButton Command11 
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
            Picture         =   "FCOMLIB04.frx":12EC
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Acceso a Base de Datos de Proveedores"
            Top             =   1560
            Width           =   650
         End
         Begin VB.CommandButton Command13 
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
            Picture         =   "FCOMLIB04.frx":15F6
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Opciones Varias - Modificaciones y Consultas"
            Top             =   5820
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
            Picture         =   "FCOMLIB04.frx":1740
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   745
            Width           =   650
         End
         Begin VB.CommandButton Command6 
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
            Height          =   630
            Left            =   105
            Picture         =   "FCOMLIB04.frx":1A4A
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Consultas y Listados"
            Top             =   4540
            Width           =   650
         End
         Begin VB.CommandButton Command4 
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
            Picture         =   "FCOMLIB04.frx":1D54
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
            Top             =   3900
            Width           =   650
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -315
            Picture         =   "FCOMLIB04.frx":205E
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   41
         ToolTipText     =   "Lista de Proveedores con Entregas Programadas"
         Top             =   870
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2050
         Left            =   0
         TabIndex        =   23
         Top             =   5985
         Width           =   10800
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0F0&
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
            TabIndex        =   24
            Top             =   0
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
               TabIndex        =   65
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
               TabIndex        =   43
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
               TabIndex        =   63
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
               TabIndex        =   28
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
               TabIndex        =   27
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
               TabIndex        =   64
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
               TabIndex        =   26
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
               TabIndex        =   25
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
               TabIndex        =   45
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
               TabIndex        =   44
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
               TabIndex        =   34
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
               TabIndex        =   33
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
               TabIndex        =   32
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
               TabIndex        =   31
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
               TabIndex        =   30
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
               TabIndex        =   29
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
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   330
               Left            =   105
               TabIndex        =   62
               ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
               Top             =   260
               Width           =   1275
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E0E0F0&
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
            TabIndex        =   35
            Top             =   0
            Width           =   6270
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
               Height          =   540
               Left            =   1155
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   61
               Top             =   1260
               Width           =   4950
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
               TabIndex        =   60
               Top             =   890
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
               TabIndex        =   57
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
               TabIndex        =   56
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
               TabIndex        =   37
               Top             =   315
               Width           =   4215
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
               TabIndex        =   59
               Top             =   600
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
               Left            =   1155
               TabIndex        =   58
               Text            =   " "
               Top             =   600
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
               Left            =   1890
               TabIndex        =   36
               Text            =   " "
               Top             =   600
               Width           =   4215
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
               TabIndex        =   42
               Top             =   930
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
               TabIndex        =   40
               Top             =   370
               Width           =   1590
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
               TabIndex        =   39
               Top             =   655
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
               TabIndex        =   38
               Top             =   1260
               Width           =   4110
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0F0&
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
         TabIndex        =   19
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   280
            Width           =   150
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0F0&
         Caption         =   "Entrega"
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
         Left            =   7665
         TabIndex        =   16
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
            TabIndex        =   18
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
            TabIndex        =   17
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
         TabIndex        =   15
         Text            =   " "
         Top             =   1180
         Width           =   4065
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
         TabIndex        =   14
         Top             =   105
         Width           =   850
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
         Left            =   9615
         TabIndex        =   8
         Text            =   " "
         Top             =   105
         Width           =   1120
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         Height          =   285
         Left            =   920
         TabIndex        =   1
         Top             =   870
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
         Left            =   160
         MaxLength       =   5
         TabIndex        =   0
         Text            =   " "
         Top             =   870
         Width           =   750
      End
      Begin VB.Label Label17 
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
         Left            =   160
         TabIndex        =   85
         Top             =   1180
         Width           =   1200
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
         Left            =   105
         TabIndex        =   84
         Top             =   0
         Width           =   1380
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   315
         Shape           =   4  'Rounded Rectangle
         Top             =   225
         Width           =   540
      End
      Begin VB.Line Line2 
         X1              =   180
         X2              =   10710
         Y1              =   4750
         Y2              =   4750
      End
      Begin VB.Line Line1 
         X1              =   10710
         X2              =   10710
         Y1              =   1680
         Y2              =   4750
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   2640
         Left            =   165
         TabIndex        =   70
         Top             =   2115
         Width           =   1350
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   2640
         Left            =   9360
         TabIndex        =   69
         Top             =   2115
         Width           =   1350
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Texto Libre"
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
         Height          =   330
         Left            =   5775
         TabIndex        =   13
         Top             =   160
         Width           =   1695
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
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   8400
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   630
         Width           =   960
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   210
         Shape           =   4  'Rounded Rectangle
         Top             =   165
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   105
         Shape           =   4  'Rounded Rectangle
         Top             =   105
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
      Caption         =   "Ordenes de Compra"
      Begin VB.Menu Nue_Orco 
         Caption         =   "Nueva Orden de Compra"
         Begin VB.Menu Directa 
            Caption         =   "O/C Directa"
         End
         Begin VB.Menu Por_MRP 
            Caption         =   "O/C por MRP"
         End
      End
      Begin VB.Menu Modi_Ocom 
         Caption         =   "Modificar O/Compra"
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
      Begin VB.Menu Desanuoc 
         Caption         =   "Revertir Anulación"
      End
      Begin VB.Menu plh6 
         Caption         =   "-"
      End
      Begin VB.Menu Maconsi1 
         Caption         =   "Materiales Consignados"
         Begin VB.Menu Entrecon 
            Caption         =   "Entregar Materiales"
         End
         Begin VB.Menu DevoConsi 
            Caption         =   "Devoluciones"
            Enabled         =   0   'False
         End
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
         Begin VB.Menu plh7 
            Caption         =   "-"
         End
         Begin VB.Menu Maconsi 
            Caption         =   "Material en Consignación"
            Begin VB.Menu Screen4 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu Print4 
               Caption         =   "Imprimir / Exportar"
            End
         End
      End
      Begin VB.Menu CoUlCom 
         Caption         =   "Control Ultimas Compras"
         Begin VB.Menu UcoPro 
            Caption         =   "Control por Proveedor"
         End
         Begin VB.Menu UcoMat 
            Caption         =   "Control por Material"
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
Attribute VB_Name = "FCOMLIB04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$
Dim MUECOPEN%
Dim EMICONSI%
Dim MODOULPRE%

Private Sub Anul_Ocom_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call OPCICOM04.Command14_Click
End Sub

Private Sub Command1_Click()
   Call SELECHO("ACREDOR")
   NCLIE% = Val(TRIM$(VALACT1$("ACREDOR")))
   If NCLIE% > 0 Then
     If REGICLI&((-1) * NCLIE%) > 0 Then
       Call BLANFORMU
       Text1.TEXT = TRIM$(Str$(NCLIE%))
       Call CARDACLI
       'Call CARCOPEN
       'Call CARCOANT
       'If ULTREG&("!OCOMDETA") > 0 Then
       '     Call CARLISCO
       '     Call CALTOCOM
       '  Else
            Call CARDETCOM
       'End If
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
       'Call CARCOPEN
       'Call CARCOANT
       'If ULTREG&("!OCOMDETA") > 0 Then
       '     Call CARLISCO
       '     Call CALTOCOM
       '  Else
            Call CARDETCOM
       'End If
     End If
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG04/AMAPRO"
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

Private Sub Command13_Click()
   Command13.Enabled = False
   'List1.Clear
   'Call SETULTREG("!OCOMDETA", 0)
   OPCICOM04.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   'DESDE ACA NUEVO PARA ENVIAR POR MAIL
    SENDBYMAIL$ = ""
    RESP22% = COMALTER%("Elija Modo de Envío\\Impresión\Por E-Mail")
    If RESP22% <> 2 Then GoTo 10
   '
5   PORMAIL% = 1
    NCI% = Val(Text1)
    NCIX% = NCI% * (-1)
    EMAI$ = TRIM$(EMAILCLI$(NCIX%))
    If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
       Call MENSERR(24, "Imposible Enviar por Correo Electrónico.\  \Dirección de Correo del Proveedor es\Inexistente o no Válida.")
       GoTo 10
    End If
    SENDBYMAIL$ = EMAI$
     '
    For Each prx1 In Printers
       If InStr(UCase$(prx1.DeviceName), "PDFWRITER") > 0 Then GoTo 10
    Next
    Call MENSERR(24, "Imposible Enviar por Mail.\Falta Impresora Virtual 'PDFWRITER'.")
    PORMAIL% = 0
    EMAI$ = ""
    SENDBYMAIL$ = ""
    'HASTA ACA AGREGADO PARA ENVIAR POR MAIL
   '
10 Label25 = ""
   Call GRACOMPRA(OKX%)
   If OKX% > 0 Then
     CUDI$ = CurDir
     Call PRICOMPRA
     '
     NUMEDO$ = TRIM$(Text11)
     While Len(NUMEDO$) < 6: NUMEDO$ = "0" + NUMEDO$: Wend
     NUMEDO$ = "OC-" + NUMEDO$ + "-00"
     '
     Call BLANFORMU
     Call SETULTREG("!OCOMDETA", 0)
     Call CIERRARCH("*.*")
     Text1.TEXT = ""
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Label25 = NUMEDO$ ' señal de que grabo OK
     '
     'AGREGADO PARA ENVIAR POR MAIL
     If SENDBYMAIL$ <> "" Then
       Call RECONECTA              ' CONECTA CON EL MISMO MODULO
       Exit Sub
     End If
     '
     Command14.Enabled = True
     Hide
   End If
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
   '
   RG1$ = REGSEL$("UNIMED")
   If Len(RG1$) > 4 Then
      Text22.TEXT = Left$(RG1$, 4)
   End If
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

Private Sub Command2_LostFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub


Private Sub Command22_Click()
    Command22.Enabled = False
    Call MENSERR(24, "No Disponible por Ahora.")
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
      SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM OCOMENCA"
      SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  DESC "
      Set VEROCTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
      With VEROCTEMP
        On Error Resume Next ' si hay algun error msgbox err
         .MoveFirst
         If Err Then
           Call MENSERR(24, "No hay Ordenes de Compra Emitidas")
           GoTo 99
         End If
         '
         JJ& = 0
         Do While (.EOF = False)
             XXY$ = REGBLAN$("INDIOCOM")
             Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4)
             Call REPLA(XXY$, !Refe_OcoM, 5, 44)
             JJ& = JJ& + 1
             Call GRAREG("INDIOCOM", XXY$, JJ&)
             .MoveNext
         Loop
      End With
      Set VEROCTEMP = Nothing
      Call SETULTREG("INDIOCOM", JJ&)
      Call CIERRARCH("INDIOCOM")
      '
3     Call SELECHO("INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
      
      NPX& = Val(VALACT1$("INDIOCOM"))
      If NPX& > 0 Then
        Call SHOWOCOM(NPX&)
        GoTo 3
      End If
      '
99  Command4.Enabled = True
   End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call FORCOMP04.Command5_Click
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   OPLISCO04.Show 1
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
    ACONEC$ = "ARCHIG04/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
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

Private Sub Directa_Click()
   Call Command1_Click
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
    OPENFORMS = DoEvents
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

Private Sub Label11_DblClick()
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
           DMONE$(U&) = DMX$
           TCMONE(U) = VHX
         Next U&
     End If
   End If
   '
   MONEMIA% = MONEMI%
   MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   Label11 = DMONE$(MONEMI%)
   Label11.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      Label11.ForeColor = RGB(255, 0, 0)
   End If
   Label11.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBX# = TICAMONE(MONEMI%)
   '
   For U& = 1 To ULTREG&("!OCOMDETA")
     YX$ = REGLEIDO$("!OCOMDETA", U&)
     PREDOL# = CVD(Mid$(YX$, 67, 8)) * COEFCON
     Call REPLA(YX$, MKD$(PREDOL#), 67, 8)
     Call GRAREG("!OCOMDETA", YX$, U&)
   Next U&
   Call CIERRARCH("!OCOMDETA")
   'Call CARLISCO
   Call CALTOCOM
   '
End Sub

Private Sub Label17_Click()
   Call SELECHO("INDIACRE")
   CPROX$ = TRIM$(VALACT1$("INDIACRE"))
   If CPROX$ <> "" Then
     NCLIE% = NUMEPRO%(CPROX$)
     If NCLIE% > 0 Then
       If REGICLI&((-1) * NCLIE%) > 0 Then
         Call BLANFORMU
         Text1.TEXT = TRIM$(Str$(NCLIE%))
         Call CARDACLI
         'Call CARCOPEN
         'Call CARCOANT
         'If ULTREG&("!OCOMDETA") > 0 Then
         '     Call CARLISCO
         '     Call CALTOCOM
          ' Else
              Call CARDETCOM
         'End If
       End If
     End If
   End If
End Sub

Private Sub Label17_DblClick()
   Call Label17_Click
End Sub

Private Sub Label22_DblClick()
   List1.Clear
   Call SETULTREG("!OCOMDETA", 0)
   Unload FORVENTB
   Call CARDETCOM
End Sub



Private Sub Ltota_DblClick(Index As Integer)
   For KKI% = 0 To 4
     Total(KKI%).Enabled = Not (Total(KKI%).Enabled)
   Next KKI%
End Sub

Private Sub MarItCum_Click()
  Call OPCICOM04.Command4_Click
End Sub

Private Sub Amastock_Click()
   Call Command8_Click
End Sub

Private Sub Modi_Ocom_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call OPCICOM04.Command13_Click
End Sub

Private Sub Nuorig_Click()
   Call OPCICOMP.Command19_Click
End Sub

Private Sub Nuorig1_Click()
   Call OPCICOMP.Command19_Click
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
   OPLISCO04.Option1.Value = True
   Call OPLISCO04.Command3_Click
   
End Sub

Private Sub Screen2_Click()
   OPLISCO04.Option1.Value = True
   Call OPLISCO04.Command4_Click
End Sub

Private Sub Screen3_Click()
   OPLISCO04.Option1.Value = True
   Call OPLISCO04.Command5_Click
End Sub

Private Sub Screen4_Click()
   OPLISCO04.Option1.Value = True
   Call OPLISCO04.Command9_Click
End Sub

Private Sub Setup_Click()
  OPCICOM04.Command18_Click
End Sub

Private Sub Text1_Change()
   Call CARDACLI
End Sub

Private Sub Text1_DblClick()
  Call SELECHO("PROVEPLA")
  NCLIE% = Val(TRIM$(VALACT1$("PROVEPLA")))
  Text1.TEXT = TRIM$(Str$(NCLIE%))
End Sub

Private Sub Text1_GotFocus()
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
   '
   Label11.ForeColor = RGB(0, 0, 255)
   Label11 = "Pesos"
   MONEMI% = 1
   Label11.ToolTipText = "Doble-Click para 'Dólares'"
   '
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

Private Sub Text18_GotFocus()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
End Sub

Private Sub Text19_Change()
   Call CALTOCOM
End Sub

Private Sub Text19_GotFocus()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
End Sub

Private Sub TEXT19_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call Text19_LostFocus
     Call CALTOCOM
   End If
End Sub

Private Sub Text19_LostFocus()
   Text19 = FORMATNUM$(Val(Text19), "F11.2")
End Sub

Private Sub Text2_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text20_Change()
   Call CALTOCOM
End Sub

Private Sub Text20_GotFocus()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
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

Private Sub Text22_DblClick()
   Call Command15_Click
End Sub

Private Sub Text22_GotFocus()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
End Sub

Private Sub Text28_GotFocus()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
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
   On Error Resume Next
   Unload FORVENTB
   On Error GoTo 0
   Label17 = ""
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   Label2 = ""
   Text8.TEXT = ""
   Text10 = ""
   Text12.TEXT = FORMATNUM$(ALIVA, "F5.2")
   Text13.TEXT = ""
   Text6.TEXT = ""
   '
   Text22 = "U."
   Text20 = "1"
   Text7 = ""
   Text18 = ""
   Text19 = ""
   Text28 = ""
   '
   TEXT21(0).TEXT = "": Eco21(0) = ""
   TEXT21(1).TEXT = "": Eco21(1) = ""
   '
   PORDESCU.TEXT = ""
   Text17 = ""
   For KKI% = 0 To 4
      Total(KKI%).Enabled = True
      Total(KKI%).TEXT = ""
   Next KKI%
   '
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
        Label17 = CODICLI$(NC%)
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Label2 = TRIM$(POSIVAC$(NC%))
        Text8.TEXT = TRIM$(CUITCLI$(NC%))
        '
        TEXT21(0).TEXT = Str$(CPAGCLI%(NC%))
        TEXT21(1).TEXT = ""
        '
        Text15 = TRIM$(Str$(MONEVACO%(NC%)))
        Label11.ForeColor = RGB(0, 0, 255)
        Label11 = "Pesos"
        MONEMI% = 1
        Label11.ToolTipText = "Doble-Click para 'Dólares'"
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
        Set OEncaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
        On Error Resume Next
        With OEncaTemp
          .MoveFirst
          If Err < 1 Then
            ATEN$ = !AtSr_Ocom: Text6 = ATEN$
            PODESC$ = !Pdes_Ocom: PORDESCU = PODESC$
          End If
        End With
        '
      End If
    End If
    Screen.MousePointer = 1
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
   Set OEncaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   With OEncaTemp
     .MoveFirst
     If Err < 1 Then
       PRONU& = 1 + !Nume_Ocom
     End If
   End With
   On Error GoTo 0
   '
   PRONUOCO& = PRONU&
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

Private Sub Text7_GotFocus()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
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
   'If ULTREG&("!OCOMDETA") < 1 Then Exit Sub
   '
   COPRECIO = 1
   If Total(0).Enabled = False Then COPRECIO = 0
   '
   NPP& = PRONUOCO&
   If NPP& > Val(Text11.TEXT) Then
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
   If Val(Text20) < 0.99 Then
      MERRO$ = "Falta Cantidad"
      GoTo 98
   End If
   '
   If TRIM$(Text7) = "" Or CUEINT%(Text7) < 1 Then
      MERRO$ = "Falta Destino / Imputación Contable"
      GoTo 98
   End If
   '
   If TRIM$(Text18) = "" Then
      MERRO$ = "Falta Descripción Corta"
      GoTo 98
   End If
   '
   If Val(Text19) < 0.01 Then
      MERRO$ = "Falta Precio Unitario"
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
    With OcomEnca
      .AddNew
          NRO& = Val(Text11)
       !Nume_Ocom = NRO&
         FEX = FEMI%
         REFE$ = FECHATEX$(FEX) + " - " + RASOCLI$((-1) * NPRO%)
       !Refe_OcoM = REFE$
       !Femi_Ocom = CVDAT(FEMI%)
       !Npro_Ocom = NPRO%
       !RaSo_Prov = RASOCLI$((-1) * NPRO%)
       !TSum_Ocom = 0              ' materiales
          STATUOC% = (-1)
          If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) = "DIRECT" Then STATUOC% = 0
       !Stat_Ocom = STATUOC%       ' status abierta Y NO APROBADA
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
       !Ides_Ocom = IDESCUEN#
       !Iiva_Ocom = TOTIVA#
       !Iotr1_Ocom = TOTPER#
       !Total_Ocom = TOTATO#
       !Lentrega = Text10
       !Obse_Ocom = Text13
         '
      .Update
    End With
    '
    ORDIT% = 1
    CI0% = 0                          ' ES DE TEXTO LIBRE  ' CVI(Mid$(ZZ$, 9, 2))
    COEFUNI = 1                       ' CVS(Mid$(ZZ$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
    PREUNPE# = Val(Text19)            ' CVD(Mid$(ZZ$, 67, 8)) * TICAMONE(MONEMI%) / COEFUNI
    PREUNCO# = PREUNPE# / TICAMONE(MONECOI%)
    MONECOS$ = "$ ": If MONECOI% > 1 Then MONECOS$ = Left$(Text16, 2)
    VAUNI$ = MONECOS$ + Mid$(FORMATNUM$(PREUNCO#, "F12.4"), 3)
    '
    With OcomDeta
      .AddNew
       !Nume_Ocom = NRO&
       !NuOrd_Item = ORDIT%
       !Femi_Ocom = CVDAT(FEMI%)
       !Npro_Ocom = NPRO%
       !RaSo_Prov = RASOCLI$((-1) * NPRO%)
       !Cod_Inte = CI0%
       !Cod_Exte = ""                 ' TEXTO LIBRE   CODEXT$(CI0%)
       !Descritem = Left$(Text18, 64) ' DESCRIT0$(CI0%)
       !Delibre = Text28
       !ICONTAB = CUEINT%(Text7)
       !Valo_Unid = VAUNI$
       !Uni_Mas = Text22              ' Mid$(ZZ$, 11, 4)
         CANTU = Val(Text20)
       !Unids_Ocom = CANTU            ' CVS(Mid$(ZZ$, 51, 4))
       !Cant_Ocom = CANTU
       !Uni_Com = Text22              ' Mid$(ZZ$, 59, 4) ' unidad de compra
       !Coef_Unids = COEFUNI
       !Pre_Unit = PREUNPE#           ' CVD(Mid$(ZZ$, 67, 8))
       !Desc_Item = 0                 ' CVS(Mid$(ZZ$, 75, 4))
       !PrUn_Neto = PREUNPE#          ' CVD(Mid$(ZZ$, 79, 8))
       !INet_Item = Val(Total(0))     ' CVD(Mid$(ZZ$, 87, 8))
       !Fentre_Sol = CVDAT(FECHANUM(Text14))    ' CVDAT(CVI(Mid$(ZZ$, 95, 2)))
       !CONSIMAT = ""                 ' TRIM$(Mid$(ZZ$, 97, 2))
       !Cant_Rec = 0
       !Cod_Oper = 0
       !Doc_Orig = ""
       !N_OrServ = 0
       !Mone_Cos = MONECOI%           ' antes pos 83,1
       !Mone_Emis = MONEMI%           ' antes pos 84,1
       !TSum_Ocom = 0                 ' materiales
          STATUOC% = (-1)
          If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) = "DIRECT" Then STATUOC% = 0
       !Stat_Ocom = STATUOC%          ' status abierta
       '
       !Nro_Otrab = Label23
       !Nro_Grup = Label24
       '
      .Update
    End With
    '
    CPAG% = Val(TEXT21(0))
    FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    '
    MERRO$ = ""
    OKX% = 1
    '
98  OcomEnca.Close
    OcomDeta.Close
    Screen.MousePointer = 1
    If MERRO$ <> "" Then Call MENSERR(24, MERRO$)
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

Sub CARDETCOM()
   '
   Text28.SetFocus
   'Call CALTOCOM
   Screen.MousePointer = 1
   '
End Sub

Sub CALTOCOM()
   '
   NC1% = (-1) * Val(Text1)
   TOTANE# = 0
   CI0% = 0
   UNIMAS$ = ""  ' Mid$(X$, 11, 4)
   CANTU = Val(Text20)
   PREUNID# = Val(Text19)
   PORDE = 0
   FENTRE% = FECHANUM(Text14)
   '
   PREUNINE# = (Int(PREUNID# * 100 * (100 - PORDE) + 0.5)) / 10000
   INETITEM# = CDbl((Int(100 * PREUNINE# * CANTU + 0.5)) / 100)
   TOTANE# = TOTANE# + INETITEM#
   '
   UNIST$ = UNIMAS$
   UNICO$ = UNIMAS$
   CANSTXX = CANTU
   '
   If CANTU > 0.05 Then
     COEFI = CANSTXX / CANTU
   End If
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
    EDITFORM.Command5.Enabled = False
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
    PRIESPEC% = 0
    PRIHRINS% = 0
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
      CODPARAM$(21) = "REF-MAT2"  ' Para EV
      CODPARAM$(22) = "REF-MAT3" ' PARA GRUPO
      CODPARAM$(23) = "NRO-LOTE" ' Nro de orden del material en la lista de materiales
      CAPARDOC% = 23
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = EMBALA$
      DOCPARAM$(3) = Total(0)
      DOCPARAM$(4) = Text12
      DOCPARAM$(5) = Total(2)
      DOCPARAM$(6) = Total(3)
      DOCPARAM$(7) = ""                           ' Mid$(X1$, 147, 20)
      DOCPARAM$(8) = ""                           ' Mid$(X1$, 67, 40)
      DOCPARAM$(9) = ""                           ' Mid$(X1$, 167, 40)
      DOCPARAM$(10) = ""                          ' Mid$(X1$, 107, 40)
      DOCPARAM$(11) = Str$(Val(Total(0)) - Val(Total(1)))
      DOCPARAM$(12) = TRIM$(Text6)                ' ATENCION SR.
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
      DOCPARAM$(21) = Label23
      DOCPARAM$(22) = Label24
      DOCPARAM$(23) = Label31
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
      CACOLTAB1% = 13
      '
      DESTIDOC% = (-1) * Val(Text1)
      CAITAB1% = 0
      CAITCRU% = 0
      INCLUSPE% = 1
      REPLAORI% = 0
      RECOPROP% = 0
      REVICODX% = 0
      '
      FORDESCRI$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
      ANCHODES% = Val(Mid$(FORDESCRI$, 2))
      If ANCHODES% < 16 Then ANCHODES% = 16
      '
      NC1% = DESTIDOC%
         CANTU = Val(Text20)                      ' CVS(Mid$(X$, 55, 4))
         PREUNID# = Val(Text19)
         PORDE = 0                                ' CVS(Mid$(X$, 75, 4))
         FEX = FECHANUM(Text14)                   ' CVI(Mid$(X$, 95, 2))
         CONSIMAT% = 0                            ' If TRIM$(Mid$(X$, 97, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
           CODDX$ = TRIM$(Label29)
           If CODDX$ = "" Then CODDX$ = Label28
         TETABU1$(1, CAITAB1%) = CODDX$
           DECRIT$ = Text18                       ' TRIM$(DESCRIT0$(CI0%))
         TETABU1$(2, CAITAB1%) = DECRIT$
         TETABU1$(3, CAITAB1%) = Text22           ' Mid$(X$, 59, 4)
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 97, 2)  ' consigna m.prima
         TETABU1$(7, CAITAB1%) = Str$(PREUNID#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PREUNID# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PREUNID# * (1 - PORDE / 100))
           KILOX$ = ""                            ' FORMATNUM$(CVS(Mid$(X$, 51, 4)) * PESUNI(CI0%), "F12.3")
         TETABU1$(11, CAITAB1%) = KILOX$
         TETABU1$(12, CAITAB1%) = ""              ' FEPLANO$(CI0%)
         TETABU1$(13, CAITAB1%) = "1"             ' TRIM$(Str$(J&))
         '
         If INCLUSPE% > 0 Then
           CAITAB1% = CAITAB1% + 1
           TETABU1$(2, CAITAB1%) = "   " ' PARA DEJAR UN RENGLON EN BLANCO
           '
           TXTX$ = Text28
           Call FRATEXTO(TXTX$, ANCHODES%)
           INTERLI = Val(ATRIFORM$(FORIPRE$, "INTERLIN"))
           '
           COLUTA% = 0
           URN& = CARETEX%
           '
           For U& = 1 To CARETEX%
             TTXY$ = RTrim$(RETEX$(U&))
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
           Next U&
           '
         End If
         '
      Call PRINTDOC("FORMOCOM")                   ' Orden de Compra
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    End If
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
999 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub

Private Sub UcoMat_Click()
Call OPLISCO04.Command20_Click
End Sub

Private Sub UcoPro_Click()
Call OPLISCO04.Command12_Click
End Sub

Private Sub Ver_Ocom_Click()
    Call OPLISCO04.Command22_Click
End Sub

Sub ENTRECONSIACC(NPROVE%)
        '
        LLAMACOMPRA% = 0
        If NPROVE% > 0 Then LLAMACOMPRA% = 1
        '
        FORIPRE$ = TRIM$(CONTROL$("", "FOREMICO"))
        If FORIPRE$ <> "" Then
            LU$ = LUDAT$
            If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
                RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
                Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
                GoTo 9999
            End If
        End If
        '
        DEPOCON% = Val(CONTROL$("", "DEPOCONS"))
        DEPORIG% = DEPOCON%
        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
        '
        'Call BLIMESS("Revisando Archivo de Movimientos de Stock")
1100    Screen.MousePointer = 11
        CMO$ = TRIM$(CONTROL$("", "CODRECON"))
        If CMO$ = "" Then CMO$ = "RC"
        FIN& = ULTREG&("MOVISTO")
        For U& = FIN& To 1 Step -1
            Call TRACE(20, FIN& - U& + 1, FIN&)
            XX$ = REGLEIDO$("MOVISTO", U&)
            If UCase$(Mid$(XX$, 21, 2)) = CMO$ Then
                UNRO& = Val(Mid$(XX$, 26, 7))
                GoTo 1002
            End If
        Next U&
        UNRO& = 0
        '
1002    REMI$ = CMO$ + "-" + TRIM$(Str$(UNRO& + 1))
        Call FRESHALL '¿?
        '
1105    Call CIERRARCH("*.*")
        '
        Screen.MousePointer = 1
        If NPROVE% > 0 Then
             EBOL$ = MKI$(NPROVE%)
          Else
             EBOL$ = OBJPLA$("SELPROVE", "")
             If EBOL$ = "" Then
                GoTo 9999
             End If
        End If
        '
        NC% = CVI(EBOL$) 'Obtengo el numero de prov
        If NC% < 1 Or ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
            Call MENSERR(24, "Proveedor Inexistente")
            NPROVE% = 0
            GoTo 1105
        End If
        '
        NC1% = (-1) * NC%
        '
        Y$ = PEDY$: If Y$ = "" Then Y$ = Space$(128)
        Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
        Call REPLA(Y$, EBOL$, 7, 2)
        Call REPLA(Y$, MKI$(HOY(HO$)), 5, 2)
        Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
        Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
        CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
        Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
1110    Call REPLA(Y$, REMI$, 105, 16)
        Call REPLA(Y$, Chr$(DEPORIG%), 121, 1)
        '
1111    PEDY$ = OBJPLA$("REMICON-ENCB", Y$)
        If PEDY$ = "" Then GoTo 9998
        FF% = CVI(Mid$(PEDY$, 5, 2))
        REMI$ = TRIM$(Mid$(PEDY$, 105, 16))
        If TRIM$(REMI$) = "" Then
            Y$ = PEDY$
            Call MENSERR(24, "Falta Numero de Remito")
            GoTo 1111
        End If
        '
        DEPORIG% = Asc(Mid$(PEDY$, 121, 1))
        If DEPOCON% > 0 Then
          If DEPORIG% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPORIG%)) = "" Then
            Call MENSERR(24, "Depósito de Salida no Válido")
            DEPORIG% = Val(CONTROL$("", "DEPOCONS"))
            GoTo 1110
          End If
        End If
        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
        '
1112    Call CIERRARCH("*.*")
        Call BLANARCH("!BOLREDET")
        '
        Screen.MousePointer = 11
        HAYOC% = 0
        For I& = 1 To ULTREG&("MACONSIG") 'Aca empieza con maconsig
            X$ = REGLEIDO$("MACONSIG", I&)
            If CVI(Mid$(X$, 3, 2)) = NC% Then
                CAPENE# = CVD(Mid$(X$, 55, 8)) - CVD(Mid$(X$, 63, 8)) ' Cant asignada - Cant entregada
                If CAPENE# > 0 Then
                    NUOC& = Val(Mid$(X$, 24, 6))
                    CI% = CVI(Mid$(X$, 33, 2))
                    STODISPO# = CDbl(STOCKACT(CI%))
                    If STODISPO# < 0 Then STODISPO# = 0
                    If STODISPO# < CAPENE# Then CAPENE# = STODISPO# ' Chequea el stock disponible
                    '
                    Y$ = REGBLAN$("!BOLREDET") 'Nueva boleta recep. prov
                    Call REPLA(Y$, MKI$(CI%), 83, 2) 'cod mat
                    Call REPLA(Y$, UNIMED$(CI%), 85, 4) 'unidad de medida
                    Call REPLA(Y$, MKD$(CAPENE#), 89, 8) 'Cantidad nominal
                    Call REPLA(Y$, MKS$(NUOC&), 151, 4) 'Nº O/C
                    Call REPLA(Y$, MKS$(I&), 189, 4) ' Relleno
                    HAYOC% = 1
                    Call REGAPP("!BOLREDET", Y$)
                End If
            End If
        Next I&
        '
        If HAYOC% = 0 Then
            Call COMUNI("No hay Consignaciones Pendientes\para este Proveedor.")
        End If
        '
1113    VTB% = VENTABU%("ECONSIG/NC/TITUPAN=Consignación de Materiales - " + TRIM$(RASOCLI$(NC1%)))
        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
            X$ = "": Y$ = "": YY$ = ""
            EBOL$ = "": PEDY$ = "": DESCRI$ = ""
            GoTo 1105
        End If
        '
        For YX& = 1 To ULTREG&("!BOLREDET")
          YY$ = REGLEIDO$("!BOLREDET", YX&)
          CI% = CVI(Mid$(YY$, 83, 2)) 'Asigna el cod mat
          CAN# = CVD(Mid$(YY$, 89, 8)) ' Asigna la cant nominal
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If CAN# > 0.005 Then
                GoTo 1114
              End If
            End If
          End If
        Next YX&
        Call MENSERR(24, "Consignación Cancelada !!!\Faltan Cantidades o Códigos.")
        GoTo 1113
        '
1114    If TRIM$(FORIPRE$) <> "" Then
           DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
           If DTP$ = "" Then GoTo 1113         ' CANCELADO
           DATRANSPO$ = DTP$
        End If
        '
11141   CONTA% = ULTREG&("!BOLREDET")
        '
1115    If CONTA% > 0 Then
          '
          Call CIERRARCH("MOVISTO")
          Call BLOQARCH("MOVISTO")
          Call BLOQARCH("MACONSIG")
          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
          '
          AINDI% = 0
          For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             CI% = CVI(Mid$(YY$, 83, 2))  'Asigna cod mat de boleta recep. prov.
             CAN# = CVD(Mid$(YY$, 89, 8)) 'Asigna cant nominal
             LOTE$ = Mid$(YY$, 69, 12)    'Asigna identif. lote
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# > 0.005 Then
                    NUOC& = CVS(Mid$(YY$, 151, 4)) ' Asigna Nº O/C
                    REGOC& = CVS(Mid$(YY$, 189, 4)) 'Asigna filler
                    DOSEC$ = ""
                    If REGOC& > 0 Then
                       If REGOC& <= ULTREG&("MACONSIG") Then
                          ZZ$ = REGLEIDO$("MACONSIG", REGOC&)
                          If CVI(Mid$(ZZ$, 3, 2)) = NC% Then ' Si son = PROV
                             If CVI(Mid$(ZZ$, 33, 2)) = CI% Then ' Si son = MAT
                                DOSEC$ = Mid$(ZZ$, 21, 12)       ' asigna Nº O/C
                                TOENTRE# = CVD(Mid$(ZZ$, 63, 8)) + CAN# 'cant entregada + cant nominal
                                Call REPLA(ZZ$, MKD$(TOENTRE#), 63, 8)  'Escribe como: nueva cantidad entregada
                                If FF% > CVI(Mid$(ZZ$, 71, 2)) Then
                                    Call REPLA(ZZ$, MKI$(FF%), 71, 2)   'Escribe la fecha de entrega
                                End If
                                Call GRAREG("MACONSIG", ZZ$, REGOC&) ' Graba el nuevo registro
                                GoTo 1120   ' Va al 1120
                             End If
                          End If
                       End If
                    End If
                    '
                    ZZ$ = REGBLAN$("MACONSIG") 'Hace un blanco en el archivo maconsig
                    Call REPLA(ZZ$, MKI$(FF%), 1, 2) 'fecha de reserva
                    Call REPLA(ZZ$, MKI$(NC%), 3, 2) ' Nº de prov.
                       NOCOM$ = TRIM$(Str$(NUOC&)) ' Calcula Nº O/C.
                       While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
                       NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-00"
                    Call REPLA(ZZ$, NOCOM$, 21, 12) ' Nº O/C
                    Call REPLA(ZZ$, MKI$(CI%), 33, 2) ' Codigo de Mat
                    Call REPLA(ZZ$, UNIMED$(CI%), 51, 4) 'Unidad de medida
                    Call REPLA(ZZ$, MKD$(CAN#), 55, 8) 'Cant asignada
                    Call REPLA(ZZ$, MKD$(CAN#), 63, 8) ' Cant entregada
                    Call REPLA(ZZ$, MKI$(FF%), 71, 2)   'Fecha de entrega
                    Call REGAPP("MACONSIG", ZZ$)
                    '
1120                CANTI = (-1) * CAN# ' Hace negativa la cantidad nominal
                    'LLAMA A LA FUNCION MOVISTO
                    'Call Movisto(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, RASOCLI$(NC1%), 0, "$", NC1%, DEPORIG%, CANTI)
                 End If
               End If
             End If
             '
          Next YX&
          Call CIERRARCH("MOVISTO")
          Call CIERRARCH("MACONSIG")
          '
          Call FRESHECHO("IREPRIPE")
          Screen.MousePointer = 1
          '
          If TRIM$(FORIPRE$) <> "" Then
            '
            FEX = FF%
            FECHDOC$ = FECHATEX$(FEX)
            NUMEDOC$ = TRIM$(REMI$)
            TIPODOC$ = "Remito de Traslado"
            '
            CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
            CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
            CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
            CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
            CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
            CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
            CODPARAM$(7) = "LOCA-TRA": DOCPARAM$(7) = TRIM$(Mid$(DTP$, 159, 32))
            CODPARAM$(8) = "NUBULTOS": DOCPARAM$(8) = TRIM$(Mid$(DTP$, 153, 6))
            CAPARDOC% = 8
            '
            CODTABU1$(1) = "COD-MAT"
            CODTABU1$(2) = "DES-MAT"
            CODTABU1$(3) = "CANTIDAD"
            CODTABU1$(4) = "CAKILOS"
            CODTABU1$(5) = "UNIMED"
            CODTABU1$(6) = "N-OCOMPR"
            CACOLTAB1% = 6
            '
            DESTIDOC% = NC1%
            CAITAB1% = 0
            '
            For YX& = 1 To ULTREG&("!BOLREDET")
               YY$ = REGLEIDO$("!BOLREDET", YX&)
               CI% = CVI(Mid$(YY$, 83, 2))
               CAN# = CVD(Mid$(YY$, 89, 8))
               If CI% > 0 Then
                  If CI% <= NUMAS% Then
                     If CAN# > 0.005 Then
                        UNID$ = TRIM$(Mid$(YY$, 85, 4))
                        If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                        CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
                        KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
                        NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                        NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
                        '
                        CAITAB1% = CAITAB1% + 1
                        TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
                        TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
                        TETABU1$(3, CAITAB1%) = CANS$
                        TETABU1$(4, CAITAB1%) = KILO$
                        TETABU1$(5, CAITAB1%) = UNID$
                        TETABU1$(6, CAITAB1%) = NUOCO$
                        '
                     End If
                  End If
               End If
            Next YX&
            '
            Call PRINTDOC("FOREMICO")   ' Remito en Consignacion
            '
          End If
          '
1199    End If
        '
9998    Call CIERRARCH("*.*")
        '
        X$ = "": Y$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
        '
        If LLAMACOMPRA% = 0 Then
           GoTo 1100
        End If
        '
9999    Call CIERRARCH("*.*")
        '
End Sub
'


