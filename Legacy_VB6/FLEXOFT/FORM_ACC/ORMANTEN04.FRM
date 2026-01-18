VERSION 5.00
Begin VB.Form ORMANTEN04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0E0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden de Mantenimiento - Gestión de Mantenimiento Industrial"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "ORMANTEN04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command17 
      Caption         =   "Cortar"
      Height          =   260
      Left            =   5940
      TabIndex        =   44
      ToolTipText     =   "Click para Cortar"
      Top             =   5980
      Width           =   1170
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Copiar"
      Height          =   260
      Left            =   7125
      TabIndex        =   43
      ToolTipText     =   "Click para Copiar"
      Top             =   5980
      Width           =   1170
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Pegar"
      Height          =   260
      Left            =   8295
      TabIndex        =   42
      ToolTipText     =   "Click para Pegar"
      Top             =   5980
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Selecc Todo"
      Height          =   260
      Left            =   9570
      TabIndex        =   41
      ToolTipText     =   "Click para Seleccionar Todo"
      Top             =   5980
      Width           =   1320
   End
   Begin VB.CommandButton Command7 
      Height          =   260
      Left            =   1200
      TabIndex        =   40
      Top             =   5980
      Width           =   4740
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
      Left            =   1200
      TabIndex        =   67
      Text            =   " "
      ToolTipText     =   "Click para Ingresar Referencia de la Orden de Mantenimiento"
      Top             =   930
      Width           =   6940
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0E0D0&
      Caption         =   "Cronograma Proyectado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   7920
      TabIndex        =   58
      Top             =   1470
      Width           =   2985
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
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
         Left            =   1890
         TabIndex        =   66
         Text            =   " "
         ToolTipText     =   "Hora Proyectada de Finalización"
         Top             =   1155
         Width           =   675
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
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
         Left            =   1890
         TabIndex        =   65
         Text            =   " "
         ToolTipText     =   "Hora Proyectada de Comienzo"
         Top             =   585
         Width           =   675
      End
      Begin VB.CommandButton Command13 
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
         Left            =   1515
         Picture         =   "ORMANTEN04.frx":0442
         TabIndex        =   63
         ToolTipText     =   "Click para Seleccionar Fecha"
         Top             =   1110
         Width           =   175
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
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
         TabIndex        =   62
         Text            =   " "
         ToolTipText     =   "Fecha Proyectada de Finalización"
         Top             =   1110
         Width           =   1020
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
         Height          =   285
         Left            =   1515
         Picture         =   "ORMANTEN04.frx":074C
         TabIndex        =   60
         ToolTipText     =   "Click para Seleccionar Fecha"
         Top             =   585
         Width           =   175
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
         TabIndex        =   59
         ToolTipText     =   "Fecha Proyectada de Comienzo"
         Top             =   585
         Width           =   1020
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "-              hs"
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
         Left            =   1750
         TabIndex        =   73
         Top             =   1155
         Width           =   2505
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "-              hs"
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
         Left            =   1750
         TabIndex        =   72
         Top             =   630
         Width           =   2505
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
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
         Left            =   -210
         TabIndex        =   64
         Top             =   880
         Width           =   930
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         Left            =   105
         TabIndex        =   61
         Top             =   360
         Width           =   2505
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0D0&
      Caption         =   "Servicios de Terceros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   4305
      TabIndex        =   53
      Top             =   1470
      Width           =   3540
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
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
         Left            =   2020
         TabIndex        =   71
         Text            =   " "
         Top             =   1110
         Width           =   1380
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
         Height          =   285
         Left            =   2460
         MaxLength       =   5
         TabIndex        =   55
         Text            =   " "
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command11 
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
         Left            =   3210
         TabIndex        =   54
         Top             =   315
         Width           =   195
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Costo Estimado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   260
         TabIndex        =   70
         Top             =   1155
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   150
         TabIndex        =   57
         Top             =   660
         Width           =   3255
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1440
         TabIndex        =   56
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.TextBox Text24 
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
      Left            =   9720
      TabIndex        =   47
      Text            =   " "
      Top             =   570
      Width           =   1020
   End
   Begin VB.CommandButton Command5 
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
      Left            =   10760
      Picture         =   "ORMANTEN04.frx":0A56
      TabIndex        =   46
      ToolTipText     =   "Click para Seleccionar Fecha"
      Top             =   570
      Width           =   175
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
      Height          =   2445
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   45
      ToolTipText     =   "Click para Ingresar Descripción de la Orden de Mantenimiento"
      Top             =   3560
      Width           =   9705
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   1200
      ScaleHeight     =   375
      ScaleWidth      =   9645
      TabIndex        =   38
      Top             =   3320
      Width           =   9705
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción General - Especificaciones"
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
         TabIndex        =   39
         Top             =   20
         Width           =   9600
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   1545
      Left            =   1200
      ScaleHeight     =   1485
      ScaleWidth      =   9675
      TabIndex        =   32
      Top             =   6360
      Width           =   9735
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         ItemData        =   "ORMANTEN04.frx":0D60
         Left            =   -30
         List            =   "ORMANTEN04.frx":0D62
         TabIndex        =   33
         ToolTipText     =   "Doble-Click para Cargar Materiales"
         Top             =   210
         Width           =   9735
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   330
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   10515
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   " Código                    Descripción                                                   U/M     Cantidad   Observaciones"
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
            Left            =   0
            TabIndex        =   35
            Top             =   15
            Width           =   10995
         End
      End
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
      Index           =   1
      Left            =   2355
      Picture         =   "ORMANTEN04.frx":0D64
      TabIndex        =   18
      ToolTipText     =   "Click para Seleccionar Equipo"
      Top             =   210
      Width           =   175
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
      Left            =   1205
      MaxLength       =   15
      TabIndex        =   19
      Text            =   " "
      ToolTipText     =   "Equipo"
      Top             =   210
      Width           =   1200
   End
   Begin VB.CommandButton CmdLiPre 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3285
      TabIndex        =   13
      Top             =   5610
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
      Height          =   285
      Left            =   2865
      TabIndex        =   12
      Text            =   " "
      Top             =   5370
      Width           =   435
   End
   Begin VB.TextBox Text11 
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
      Left            =   9975
      TabIndex        =   7
      Top             =   8595
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   10000
      Left            =   11025
      ScaleHeight     =   9945
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command15 
         Caption         =   "Ver"
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
         Picture         =   "ORMANTEN04.frx":106E
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
         Top             =   3270
         Width           =   650
      End
      Begin VB.CommandButton Command10 
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
         Picture         =   "ORMANTEN04.frx":1378
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Aprueba, Graba e Imprime Orden de Mantenimiento"
         Top             =   6825
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   640
         Left            =   105
         Picture         =   "ORMANTEN04.frx":1682
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Opciones Varias"
         Top             =   2040
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Left            =   120
         Picture         =   "ORMANTEN04.frx":17CC
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   4920
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Copy"
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
         Picture         =   "ORMANTEN04.frx":1C0E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Copia Interactiva "
         Top             =   3960
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   5985
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   6
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command99 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "ORMANTEN04.frx":1D58
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   9000
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Picture         =   "ORMANTEN04.frx":2062
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit"
         Top             =   105
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Help"
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
         Picture         =   "ORMANTEN04.frx":21AC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ayuda en Línea"
         Top             =   800
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "ORMANTEN04.frx":24B6
         Stretch         =   -1  'True
         Top             =   6300
         Width           =   1410
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00D0E0D0&
      Caption         =   "Personal Propio Asignado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   105
      TabIndex        =   50
      Top             =   1470
      Width           =   4140
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   120
         TabIndex        =   51
         ToolTipText     =   "Doble-Click para Ingresar Empleados"
         Top             =   630
         Width           =   3900
      End
      Begin VB.Label Label92 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Leg.           Apellido y Nombre       "
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
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   3900
      End
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia:"
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
      Left            =   30
      TabIndex        =   69
      Top             =   1000
      Width           =   1110
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11025
      Y1              =   1370
      Y2              =   1370
   End
   Begin VB.Line Line6 
      X1              =   10910
      X2              =   10910
      Y1              =   3320
      Y2              =   6220
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle:"
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
      Left            =   0
      TabIndex        =   37
      Top             =   3380
      Width           =   1140
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Materiales:"
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
      Left            =   120
      TabIndex        =   36
      Top             =   6440
      Width           =   1020
   End
   Begin VB.Label Label24 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   9120
      TabIndex        =   31
      Top             =   930
      Width           =   1815
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   10080
      TabIndex        =   30
      Top             =   210
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   7410
      TabIndex        =   29
      Top             =   570
      Width           =   735
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Posición:"
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
      Left            =   6240
      TabIndex        =   28
      Top             =   690
      Width           =   1110
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   4410
      TabIndex        =   27
      Top             =   570
      Width           =   2055
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Linea:"
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
      Left            =   3255
      TabIndex        =   26
      Top             =   690
      Width           =   1110
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   2280
      TabIndex        =   25
      Top             =   10840
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   2805
      TabIndex        =   24
      Top             =   570
      Width           =   930
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   2040
      TabIndex        =   23
      Top             =   10840
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   1200
      TabIndex        =   22
      Top             =   570
      Width           =   690
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   2640
      TabIndex        =   21
      Top             =   210
      Width           =   5505
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Equipo:"
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
      Index           =   39
      Left            =   360
      TabIndex        =   20
      Top             =   285
      Width           =   795
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Left            =   5115
      TabIndex        =   16
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Sección:"
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
      Left            =   1860
      TabIndex        =   15
      Top             =   690
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Planta:"
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
      Left            =   30
      TabIndex        =   14
      Top             =   690
      Width           =   1110
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11025
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Resp.:"
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
      Left            =   8520
      TabIndex        =   11
      Top             =   1050
      Width           =   585
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de O/M:"
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
      Left            =   8955
      TabIndex        =   10
      Top             =   300
      Width           =   1050
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "F.Emisión:"
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
      Left            =   8760
      TabIndex        =   9
      Top             =   660
      Width           =   930
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Resp.:"
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
      Left            =   9450
      TabIndex        =   8
      Top             =   8700
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Line Line1 
      X1              =   -420
      X2              =   11580
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu MnuArch 
      Caption         =   "Archivos"
      Begin VB.Menu MnuNuOM 
         Caption         =   "Nueva O/M"
         Begin VB.Menu MnuDire 
            Caption         =   "Directa"
         End
         Begin VB.Menu php0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuPreExi 
            Caption         =   "Copia Pre-Existente"
         End
         Begin VB.Menu MnuPlanti 
            Caption         =   "Copia Plantilla"
         End
      End
      Begin VB.Menu MnuModi 
         Caption         =   "Modificar O/M"
      End
      Begin VB.Menu MnuCerr 
         Caption         =   "Cerrar O/M"
      End
      Begin VB.Menu MnuDel 
         Caption         =   "Anulación O/M"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu php2 
         Caption         =   "Ver y Re-Imprimir"
      End
      Begin VB.Menu php3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFin 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu MnuEdi 
      Caption         =   "Edición"
      Begin VB.Menu MnuBusc 
         Caption         =   "Buscar"
      End
      Begin VB.Menu MnuDesCam 
         Caption         =   "Des-Hacer Cambios"
      End
      Begin VB.Menu php10 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVeryRe 
         Caption         =   "Ver y Re-Imprimir"
      End
   End
   Begin VB.Menu MnuStp 
      Caption         =   "Configuración"
   End
   Begin VB.Menu MnuConLis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu MnuNum 
         Caption         =   "Por Número"
         Begin VB.Menu MnuPorPan 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu MnuImp1 
            Caption         =   "Por Impresión"
         End
      End
      Begin VB.Menu MnuEqui 
         Caption         =   "Por Equipo"
         Begin VB.Menu MnuPant1 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu PorImp2 
            Caption         =   "Por Impresión"
         End
      End
      Begin VB.Menu php20 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVeryRei 
         Caption         =   "Ver y Re-Imprimir"
      End
   End
   Begin VB.Menu MnuAyu 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "ORMANTEN04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%, MODIFIC%
Dim GRABATODO%
Dim ARCHIPRE$
Private Sub Command1_Click()
    '
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
    '
End Sub

Private Sub AyuCli_Click()
  Call Command3_Click
End Sub

Private Sub Command10_Click()
  '
  Command10.Enabled = False
  'VALIDACIONES
  '
  If DERACCE%("NUORMAN", "Nueva Orden de Mantenimiento") > 1 Then GoTo 99
  '
  If TRIM$(Text7.TEXT) = "" Then
    Call MENSERR(24, "Referencia no Válida.")
    GoTo 99
  End If
  '
  NPRO = Val(Text3.TEXT) * -1
  If REGICLI&(NPRO) <= 0 And List4.ListCount <= 0 Then
    Call MENSERR(24, "Es Necesario Asignar Empleados\o Tercerizar la O/M.")
    GoTo 99
  End If
  '
  If TRIM(Text1) = "" Or ECOARCH$("INDIEQUI", Text1.TEXT) = "" Then
    Call MENSERR(24, "Equipo no Válido.")
    GoTo 99
  End If
  '
  If TRIM(Text28) = "" Then
    Call MENSERR(24, "Descripción no Válida.")
    GoTo 99
  End If
  '
  HOII% = HOY(HOS$)
  FEMI% = FECHANUM(Text24.TEXT)
  If FEMI% < 1 Or FEMI% > HOII% Then
    Call MENSERR(24, "Fecha de Emisión no Válida")
    GoTo 99
  End If
  '
  FEPROCOM% = FECHANUM(Text2.TEXT)
  If FEPROCOM% < FEMI% Then
    Call MENSERR(24, "Fecha de Comienzo Proyectada no Válida")
    GoTo 99
  End If
  '
  FEPROFIN% = FECHANUM(Text4.TEXT)
    If FEPROFIN% < FEPROCOM% Then
    Call MENSERR(24, "Fecha de Finalización Proyectada NO puede ser\ Inferior a la Fecha de Comienzo Proyectada.")
    GoTo 99
  End If
  '
  'GRABACIÓN
  Call ABREOMAN
  NRO& = Val(Label22)
  With OrdeMant
    .AddNew
      !NumeOrMa_OM = NRO&
      !Referen_OM = Text7.TEXT 'REFERENCIA CORTA
      !FechGen_OM = CVDAT(FEMI%)
      !Codigo_Equ = Text1.TEXT
      !Descri_OM = Text28.TEXT
      !Status_OM = 1 'O/M generada
      !UsuGen_OM = Label24
      !Proved_OM = Val(Text3.TEXT)
      !CostoTer_OM = Val(Text8.TEXT)
      !FeProCom_OM = CVDAT(FECHANUM(Text2.TEXT))
      !HoProCom_OM = CDate(Text5.TEXT)
      !FeProFin_OM = CVDAT(FECHANUM(Text4.TEXT))
      !HoProFin_OM = CDate(Text6.TEXT)
      !UsuCie_OM = " "
    .Update
  End With
  '
  'CARGO EMPLEADOS ASIGNADOS A LA O/M
  If List4.ListCount > 0 Then
    For JX% = 0 To List4.ListCount - 1
      L1$ = List4.List(JX%)
      LEG1% = Val(Left$(L1$, 8))
      With RecuProp
        .AddNew
          !NumeOrMa_OM = NRO&
          !LegaEmp_OM = LEG1%
        .Update
      End With
    Next JX%
  End If
  '
  'RESERVAS
  CMO$ = TRIM$(CONTROL$("", "CODORMAN"))
  If CMO$ = "" Then CMO$ = "OM"
  Call ABRESERVA
  NBRS$ = TRIM$(Str$(NRO&)): While Len(NBRS$) < 6
                                  NBRS$ = "0" + NBRS$
                             Wend
  NBRS$ = CMO$ + "-" + NBRS$ + "-"
  For UX% = 0 To List1.ListCount - 1
    YY$ = List1.List(UX%)
    NORMA$ = NBRS$ + TRIM$(Str$(UX%))
    NORFX$ = NORMA$
    '
    CIRF0$ = Mid$(YY$, 1, 16)
    CIRF% = CODINT%(CIRF0$)
    CANTRX# = Val(Mid$(YY$, 62, 10))
    USOIX = Val(Mid$(YY$, 62, 10))
    OBS2$ = Mid$(YY$, 73, 25)
    IDLOT$ = ""
    DEPOX% = DEPOPRE%(CIRF%) 'DEPOSITO PREDETERMINADO
    '
    With Reserva
      .AddNew
      !Cod_Inte = CIRF%
      !Cod_Exte = CODEXT$(CIRF%)
      !DESCRIP = DESCRIT0$(CIRF%)
      !IdSubOr = NORFX$
      !FechRes = Now
      !CantRes = CANTRX#
      !CantCons = 0
      !Saldo_Entre = CANTRX#
      !IdenLote = IDLOT$
      !Deposito = DEPOX%
      !UsoUnit = USOIX
      On Error Resume Next
      !Observa = OBS2$
      On Error GoTo 0
      .Update
    End With
  Next UX%
  Reserva.Close
  '
  Call INDIMAN 'INDICE DE O/MANTENIMIENTO
  '
  'IMPRESION
  Call PRIOMAN 'IMPRESION DE LA O/MANTENIMIENTO
  '
  Call BLANFORMU
  Text1.TEXT = ""
  '
  MODIFIC% = 0
  '
99 Command10.Enabled = True
  '
End Sub

Private Sub Command11_Click()
   '
   Call SELECHO("ACREDOR")
   NCLIE% = Val(TRIM$(VALACT1$("ACREDOR")))
   If NCLIE% > 0 Then
     If REGICLI&((-1) * NCLIE%) > 0 Then
       Text3.TEXT = TRIM$(Str$(NCLIE%))
       Label18 = RASOCLI(NCLIE% * (-1))
     End If
   End If
   '
End Sub

Private Sub Command12_Click()
   '
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text2.TEXT = VALACT1$("SELFECHA")
   End If
   '
End Sub

Private Sub Command13_Click()
   '
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text4.TEXT = VALACT1$("SELFECHA")
   End If
   '
End Sub

Private Sub Command14_Click()
  Command14.Enabled = False
    OPOMAQ04.Show 1
  Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
 '
   Command15.Enabled = False
   Call COPYSTRU("INDIOMAN", "INPEOMAN")
   Call BLANARCH("!INPEOMAN")
   '
   Call ABREOMAN
   '
   SQLX$ = "SELECT * FROM ORDEMANT "
   SQLX$ = SQLX$ + " ORDER BY NumeOrMa_OM ASC "
   Set OMTMP = Omantin.OpenRecordset(SQLX$, 4)
   With OMTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       GoTo 99
     End If
     On Error GoTo 0
    '
    Y$ = REGBLAN$("INPEOMAN")
    JJ& = 0
    ANU1$ = " (A)"
    CER1$ = " (C)"
    Do While .EOF = False
      NRO& = !NumeOrMa_OM
      Call REPLA(Y$, MKS$(NRO&), 1, 4)
      DES1$ = !Codigo_Equ
      DES1$ = DES1$ + " - " + Mid$(!Referen_OM, 1, 20)
      DES1$ = DES1$ + " - " + FECHATEX$(CVINT(!FechGen_OM))
      If !Status_OM < 0 Then DES1$ = DES1$ + ANU1$
      If !Status_OM = 3 Then DES1$ = DES1$ + CER1$
      Call REPLA(Y$, DES1$, 5, 44)
      JJ& = JJ& + 1
      Call GRAREG("!INPEOMAN", Y$, JJ&)
      .MoveNext
    Loop
    '
  End With
  '
  Call SETULTREG("!INPEOMAN", JJ&)
  Call CIERRARCH("!INPEOMAN")
  '
3 YYY$ = REGSEL$("!INPEOMAN/Indice General de Ordenes de Mantenimiento (Pendientes-Cumplidas-Anuladas)")
  '
  If Len(YYY$) <= 4 Then GoTo 99
  '
  NPX& = CVS(Left$(YYY$, 4))
  '
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Orden de Mantenimiento"
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
    GoTo 3
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
    EDITFORM.command1.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command15.Enabled = True
    '
90 End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command2_Click()
  Command2.Enabled = False
    LISOMAQ04.Show 1
  Command2.Enabled = True
End Sub

Private Sub Command22_Click()
  '
  Command22.Enabled = False
  RTA1% = COMALTER%("Puede Realizar una Copia de una O/M Pre-Existente\ o Usar una Plantilla de O/M.\\Alta por Copia\Alta por Plantilla\Cancelar")
  Select Case RTA1%
    Case 1
      Call ALCOPCOM
    Case 2
      'Call ALPLANOM
  End Select
  Command22.Enabled = True
  '
End Sub

Private Sub Command4_Click()
  '
  Call HELPONLINE("OMANTEN")
  '
End Sub

Private Sub ConPagCli_Click()
With OPMASCL04
  .Option1(6).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub CuCoCli_Click()
With OPMASCL04
  .Option1(8).Value = True
  Call .Command16_Click
End With
End Sub


Private Sub DeLiCreCli_Click()
    Call AMACLIS04.Command3_Click
End Sub

Private Sub EditCli_Click()
   'Call Command9_Click
End Sub

Private Sub ConGenCli_Click()
  '
  Call Command14_Click
  '
End Sub

Private Sub Exit_Click()
   Call Command3_Click
End Sub

Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub

Private Sub Command9_Click(Index As Integer)
   '
   Call SELECHO("INDIEQUI")
   COD$ = TRIM$(VALACT1$("INDIEQUI"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

'
Private Sub Form_Load()
  '
  Call CENTRAFORM(Me)
  '
  EPAC$ = TRIM$(EMPREAC$)
  If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
  End If
  Call GRATITFOR(Caption)
  '
  HOII% = HOY(HOS$)
  Text24.TEXT = HOS$
  Text2.TEXT = HOS$
  Text5.TEXT = Left$(Time$, 5)
  '
  Label24.Caption = USERNAME$
  '
  'NUMERO DE PROXIMA OM
  NRO& = 1
  Call ABREOMAN
  '
  SQLX$ = "SELECT NumeOrMa_OM FROM ORDEMANT "
  SQLX$ = SQLX$ + "ORDER BY NumeOrMa_OM DESC;"
  Set ORMANTEMP = Omantin.OpenRecordset(SQLX$, dbOpenSnapshot)
  On Error Resume Next
  With ORMANTEMP
     .MoveFirst
     If Err < 1 Then
       NRO& = 1 + !NumeOrMa_OM
     End If
  End With
  On Error GoTo 0
  '
  Label22 = FORMATNUM$(NRO&, "F9.0")
  '
  MODIFIC% = 0
  '
End Sub
'
Private Sub Form_Activate()
  WindowState = 0
  OPENFORMS = DoEvents
End Sub
'
Private Sub Form_Resize()
  If WindowState = 1 Then
    Call CIERRARCH("*.*")
    Call BAJALDISCO
  End If
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

'\\\OT-05-0396-WAR-13/06/05///
'Private Sub COMMAND15_CLICK()
'    '
'    If DERACCE%("GRACOTCA", "Grabar Cotización") >= 2 Then
'      NUPRESU1$ = TRIM(Label31)
'      'RUBRO =2
'      For I = 0 To List1.ListCount - 1
'          Y$ = List1.List(I)
'          CODEXT1$ = TRIM(Mid$(Y$, 1, 16))
'          CI% = CODINT(CODEXT1$)
'          PREU1# = Val(Mid$(Y$, 65, 12))
'          CANTI1# = Val(Mid$(Y$, 78, 7))
'          IMPITEM# = Val(Mid$(Y$, 86, 12))
'          Call ABREPRECAM
'          With DetaPresu
'            SQLX$ = " Nro_Presu = '" & NUPRESU1$ & "' "
'            SQLX$ = SQLX$ + " AND  Rubro_Det = 3 "
'            'SQLX$ = SQLX$ + " AND  SubRub_Det = 1 "
'            SQLX$ = SQLX$ + " AND CodINte_Det = " & CI% & " "
'            .FindFirst SQLX$
'10              If .NoMatch = False Then
'              .Edit
'                !PreUni_Det = PREU1#
'                !Cantid_Det = CANTI1#
'                !Impo_Det = IMPITEM#
'              .Update
'              .FindNext SQLX$
'              GoTo 10
'            End If
'          End With
'      Next I
'    End If
'End Sub
'\\\OT-05-0396-WAR-FIN///
Sub CAMGRIS()
   '
   COLOBAK = RGB(230, 230, 230)
   Text6.BackColor = COLOBAK
   Text13.BackColor = COLOBAK
   Text38.BackColor = COLOBAK
   TxtLiPre.BackColor = COLOBAK
   TICAMBIS.BackColor = COLOBAK
   '
End Sub

Private Sub Command25_Click()
   'Call PRILISMAT
  PRINTFORM
End Sub

Private Sub Command3_Click()
   '
   If MODIFIC% = 1 Then
    If COMALTER%("¿Desea Realmente Cancelar la \Generación de la O/M?\\No, Continuar\Si, Cancelar.") = 1 Then
      MODIFIC% = 1
      Exit Sub
    End If
   End If
   '
   Call BLANFORMU
   List1.Clear
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Text1.TEXT = ""
   MODIFIC% = 0
   Hide
   '
End Sub

Private Sub Command5_Click()
   '
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text24.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text24.TEXT)
   '
End Sub

Private Sub LiCuInCli_Click()
    Call AMACLIS04.Command5_Click
End Sub

Private Sub Frame3_DBLCLICK()
'    Call List1_DblClick
End Sub


Private Sub Label17_dblClick()
    'MUESTRA AL O/C
     NPX& = Val(Mid$(Label17, 4, 7))
     If NPX& > 0 Then
        Call SHOWOCOM(NPX&)
     End If
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
    EDITFORM.command1.Enabled = False
    EDITFORM.Show 1
    '
End Sub

Private Sub Label18_dblClick()
    '
    NPX& = Val(Mid$(Label18, 4, 7))
    If NPX& > 0 Then
     Call SHOWBORE(NPX&)
    End If
    '
End Sub
Sub SHOWBORE(NPX&)
    '
    TIDOCUM$ = "Boleta de Recepción"
    NUDOCU$ = TRIM$(Str$(NPX&))
    NUDOCU$ = " " + NUDOCU$ + " -"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
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
5  For KKI% = 0 To 5
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
    EDITFORM.command1.Enabled = False
    EDITFORM.Show 1
    '
End Sub
'Private Sub Label3_Change()
'  '
'  Label6 = TRIM$(ECOARCH$("PLANTPRO", Label3.Caption))
'  '
'End Sub

'Private Sub Label7_Change()
'  '
'  CS1 = Val(Label7)
'  Label13.Caption = TRIM$(ECOARCH$("SECPROD", MKS$(CS1)))
'  '
'End Sub
Private Sub LiGeClaCli_Click()
 Call Command25_Click
End Sub

Private Sub SuspOper_Click()
    Call Command3_Click
End Sub

Sub GENEPLACO() 'GENERAR CAPLACO.DAT
    '
    NUOTRAB$ = TRIM$(Text1.TEXT)
    Call ABREORTRA
    SQLX$ = "SELECT * FROM PlanCob "
    SQLX$ = SQLX$ + " WHERE Nro_Otrab= '" & NUOTRAB$ & "' "
    Set PLACOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
    
    J& = 0
    With PLACOTMP
      '
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do Until .EOF = False
          Y$ = REGBLAN("CAPLACO")
          Call REPLA(Y$, !Nro_Otrab, 1, 16)
          Call REPLA(Y$, !NuorItem, 17, 2)
          Call REPLA(Y$, !Refe_Pcob, 19, 64)
          Call REPLA(Y$, !Impo_Pcob, 83, 8)
          Call REPLA(Y$, !FeVen_Pcob, 91, 2)
          Call REPLA(Y$, !NroFac_Pcob, 93, 16)
          Call REPLA(Y$, !MONEDA, 109, 2)
          Call REPLA(Y$, !Coti_Pcob, 111, 8)
          J& = J& + 1
          Call GRAREG("!CAPLACO", Y$, J&)
        .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
    Call SETULTREG("!CAPLACO", J&)
    Call CIERRARCH("!CAPLACO")
    
End Sub



Private Sub Label9_DblClick()
'
  If TRIM(Text1) <> "" Then
    '
    Call BLANARCH("!REPUESTO")
    JJ& = 0
    For JX% = 0 To List1.ListCount - 1
      L1$ = List1.List(JX%)
      CODI0$ = Mid$(L1$, 1, 16)
        CODI1% = CODINT%(CODI0$)
      DESC1$ = Mid$(L1$, 18, 40)
      UM2$ = Mid$(L1$, 59, 2)
      USO0$ = Mid$(L1$, 62, 10)
        USO1 = Val(USO0$)
      OBS1$ = Mid$(L1$, 73, 25)
      '
      YY$ = REGBLAN$("REPUESTO")
      Call REPLA(YY$, MKI$(CODI1%), 1, 2)
      Call REPLA(YY$, UM2$, 3, 2)
      Call REPLA(YY$, MKS$(USO1), 5, 4)
      Call REPLA(YY$, OBS1$, 9, 25)
      JJ& = JJ& + 1
      Call GRAREG("!REPUESTO", YY$, JJ&)
      '
    Next JX%
    Call SETULTREG("!REPUESTO", JJ&)
    Call CIERRARCH("!REPUESTO")
    '
    PRF$ = "CAREPUES"
    ATRI$ = ATRI$ + "/TITUPAN=Lista de Repuestos /"
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + AMAEQUI04.Left + 170))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + AMAEQUI04.Top + 2250))
    VTB% = VENTABU%(PRF$ + ATRI$)
    If VTB% >= 0 Then
      '
      List1.Clear
      '
      FIN& = ULTREG&("!REPUESTO")
      For J& = 1 To FIN&
        X$ = REGLEIDO$("!REPUESTO", J&)
        CODI1% = CVI(Mid$(X$, 1, 2))
        CODI2$ = CODEXT$(CODI1%)
        DESC2$ = DESCRIT0(CODI1%)
        UM2$ = Mid$(X$, 3, 2)
        USO1 = CVS(Mid$(X$, 5, 4))
        USO2$ = FORMATNUM$(USO1, "F9.4")
        OBS2$ = Mid$(X$, 9, 25)
        'AGREGO AL LIST
        L1$ = Space(128)
        Call REPLA(L1$, CODI2$, 1, 16)
        Call REPLA(L1$, DESC2$, 18, 40)
        Call REPLA(L1$, UM2$, 59, 2)
        Call REPLA(L1$, USO2$, 62, 10)
        Call REPLA(L1$, OBS2$, 73, 25)
        List1.AddItem L1$
      Next J&
      '
    End If
  End If
  '
End Sub

Private Sub Label92_dblClick()
'
  If TRIM(Text1) <> "" Then
    '
    Call BLANARCH("!RECPROP")
    JJ& = 0
    For JX% = 0 To List4.ListCount - 1
      L1$ = List4.List(JX%)
      LEG1% = Mid$(L1$, 1, 8)
      '
      YY$ = REGBLAN$("RECPROP")
      Call REPLA(YY$, MKI$(LEG1%), 1, 2)
      JJ& = JJ& + 1
      Call GRAREG("!RECPROP", YY$, JJ&)
      '
    Next JX%
    Call SETULTREG("!RECPROP", JJ&)
    Call CIERRARCH("!RECPROP")
    '
    PRF$ = "RECPROP"
    ATRI$ = ATRI$ + "/TITUPAN=Lista de Empleados /"
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + AMAEQUI04.Left + 170))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + AMAEQUI04.Top + 2250))
    VTB% = VENTABU%(PRF$ + ATRI$)
    If VTB% >= 0 Then
      '
      List4.Clear
      '
      FIN& = ULTREG&("!RECPROP")
      For J& = 1 To FIN&
        X$ = REGLEIDO$("!RECPROP", J&)
        LEG1% = CVI(Mid$(X$, 1, 2))
          LEG2$ = FORMATNUM$(LEG1%, "F5.0")
        EMP1$ = ECOARCH$("PADRON", MKI$(LEG1%))
        'AGREGO AL LIST
        L1$ = Space(128)
        Call REPLA(L1$, LEG2$, 1, 8)
        Call REPLA(L1$, EMP1$, 10, 20)
        List4.AddItem L1$
      Next J&
      '
    End If
  End If
  '
End Sub

Private Sub List1_DblClick()
  Call Label9_DblClick
End Sub

Private Sub List4_dblClick()
  '
  Call Label92_dblClick
  '
End Sub

 Sub MnuAyu_Click()
  '
  Call Command4_Click
  '
End Sub

 Sub MnuCerr_Click()
  '
  Call LISOMAQ04.Command46_Click
  '
End Sub

 Sub MnuDel_Click()
  '
  Call LISOMAQ04.Command14_Click
  '
End Sub

 Sub MnuDire_Click()
  '
  Call BLANFORMU
  Text1.TEXT = ""
  '
End Sub

 Sub MnuFin_Click()
  '
  Call Command3_Click
  '
End Sub

 Sub MnuImp1_Click()
   '
   With LISOMAQ04
      .Option2.Value = True
      Call .Command3_Click
   End With
   '

End Sub

 Sub MnuModi_Click()
  '
  Call LISOMAQ04.Command13_Click
  '
End Sub

 Sub MnuPant1_Click()
   '
   With LISOMAQ04
      .Option1.Value = True
      Call .Command4_Click
   End With
   '
End Sub

 Sub MnuPorPan_Click()
   '
   With LISOMAQ04
      .Option1.Value = True
      Call .Command3_Click
   End With
   '
End Sub

 Sub MnuPreExi_Click()
 '
 Call ALCOPCOM
 '
End Sub

 Sub MnuStp_Click()
  '
  Call Command14_Click
  '
End Sub

Sub MnuVeryRe_Click()
  '
  Call LISOMAQ04.Command22_Click
  '
End Sub

Sub MnuVeryRei_Click()
  '
  Call LISOMAQ04.Command22_Click
  '
End Sub

 Sub php2_Click()
  '
  Call LISOMAQ04.Command22_Click
  '
End Sub

 Sub PorImp2_Click()
   '
   With LISOMAQ04
      .Option2.Value = True
      Call .Command4_Click
   End With
   '
End Sub

Private Sub Text1_Change()
  '
  Call BLANFORMU
  If TRIM$(Text1) <> "" Then
    If ECOARCH$("INDIEQUI", Text1.TEXT) <> "" Then
      Label1 = ECOARCH$("INDIEQUI", Text1.TEXT)
      Call CARGAEQUI(Text1.TEXT)
      MODIFIC% = 1
    End If
  End If
  '
End Sub
Private Sub Text10_GOTFOCUS()
  '
  Command3.SetFocus
  '
End Sub

Private Sub Text12_GotFocus()
  '
  Command3.SetFocus
  '
End Sub
Private Sub Text13_GotFocus()
  '
  Command3.SetFocus
  '
End Sub
Private Sub Text18_GOTFOCUS()
  '
  Command3.SetFocus
  '
End Sub

Private Sub Text2_change()
  '
  Static HOX%
  If HOX% = 0 Then
    HOX% = HOY(HO$)
  End If
  '
  If FECHANUM(Text2.TEXT) > HOX% Then
    Text5.TEXT = "0:00"
  ElseIf FECHANUM(Text2.TEXT) = HOX% Then
    Text5.TEXT = Left$(Time$, 5)
  End If
  '
  MODIFIC% = 1
  '
End Sub
Private Sub Text21_Change()
'    '
'    If Val(TEXT21.TEXT) < 0 Or Val(TEXT21.TEXT) > 255 Then
'       TEXT21.TEXT = ""
'    End If
'    '
'    If Val(TEXT21) < 1 Then
'        TxtLiPre.TEXT = "CONCERTADOS FIJOS"
'        ARCHIPRE$ = "MASTER"
'      Else
'        TxtLiPre.TEXT = ECOARCH$("LISTAS", MKI$(Val(TEXT21.TEXT)))
'        ARCHIPRE$ = TRIM$(TEXT21)
'        While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
'        ARCHIPRE$ = "LIPRE" + ARCHIPRE$
'    End If
'    '
End Sub


Sub BLANFORMU()
   '
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   '
   HOII% = HOY(HOS$)
   Text24.TEXT = HOS$
   Text2.TEXT = HOS$
   Text5.TEXT = Left$(Time$, 5)
   '
   List1.Clear
   List4.Clear
   Text28.TEXT = ""
   '
   Text23 = ""
   Text35 = ""
   Text36 = ""
   Text37 = ""
   Text3 = ""
   Text18 = ""
   Text7 = ""
   Text8 = ""
   '
   Label1 = ""
   Label3 = ""
   Label6 = ""
   Label7 = ""
   Label13 = ""
   Label15 = ""
   Label20 = ""
   '
   Text9 = ""
   Text10 = ""
   Text25 = ""
   Text11 = ""
   Text13 = ""
   Text12 = ""
   Text17 = ""
   Text6 = ""
   '
   'NUMERO DE PROXIMA OM
   NRO& = 1
   Call ABREOMAN
   '
   SQLX$ = "SELECT NumeOrMa_OM FROM ORDEMANT "
   SQLX$ = SQLX$ + "ORDER BY NumeOrMa_OM DESC;"
   Set ORMANTEMP = Omantin.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   With ORMANTEMP
      .MoveFirst
      If Err < 1 Then
        NRO& = 1 + !NumeOrMa_OM
      End If
   End With
   On Error GoTo 0
   '
   Label22 = FORMATNUM$(NRO&, "F9.0")
   '
   YACAR% = 0
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
   Unload Me
   '
End Sub

Private Sub GruCli_Click()
With OPMASCL04
  .Option1(9).Value = True
  Call .Command16_Click
End With
End Sub



Private Sub LiPreCli_Click()
With OPMASCL04
  .Option1(7).Value = True
  Call .Command16_Click
End With
End Sub




Private Sub Print_Click()
   Call Command25_Click
End Sub

Private Sub PrZoVeCli_Click()
With OPMASCL04
  .Option1(5).Value = True
  Call .Command16_Click
End With
End Sub

'Sub PRILISMAT()
'    '
'    'ACA VA CODIGO PARA IMPRIMIR OBSERVACIONES
'    'Call SETULTREG("!OBSERVOF", 0)
'    '
'    ' IMPRESION DE LA LISTA DE MATERIALES
'    '
'    FORIPRE$ = CONTROL$("", "LISTAMAT")
'    '
'    If TRIM$(FORIPRE$) <> "" Then
'      '
'      Screen.MousePointer = 11
'' VER ESTO CON WALTER
'' TODA LA PARTE DE LOS PARAMETROS TENES QUE HACERLA VOS, FLAVIO
'      FECHDOC$ = Text9
'      NUMEDOC$ = TRIM$(Label31) + " / " + TRIM$(Label2)
'      TIPODOC$ = "Lista de Materiales "
'      '
'      DESTIDOC% = Val(Label22) 'CLIENTE
'      '
'      CODPARAM$(1) = "DOCUORIG" 'NUMERO DE OBRA
'      CODPARAM$(2) = "REF-MAT1" 'CODIGO DE GRUPO
'      CODPARAM$(3) = "RASO-CLI" 'CLIENTE
'      CODPARAM$(4) = "FECH-EMI" 'FECHA DE EMISION
'      CODPARAM$(5) = "REF-MAT2" 'DENOMINACION DE LA OBRA
'      CODPARAM$(6) = "REF-MAT3" 'DENOMINACION DEL GRUPO
'      CODPARAM$(7) = "TELE-CLI" 'RESPONSABLE
'
'      CAPARDOC% = 7
'
'      DOCPARAM(1) = Label31
'      DOCPARAM(2) = Label2
'      DOCPARAM(3) = Label5
'      DOCPARAM(4) = Text9
'      DOCPARAM(5) = Label7
'      DOCPARAM(6) = Label8
'      DOCPARAM(7) = Text11
'      '
' 'A PARTIR DE AQUI LO MODIFIQUE YO - EPB
'      CAITAB1% = 0
'      CAITCRU% = 0
'      '
' 'AQUI PONER LOS CODIGOS DE COLUMNAS A LISTAR
'      CODTABU1$(1) = "ORD-ITEM"
'      CODTABU1$(2) = "CANTIDAD"
'      CODTABU1$(3) = "DES-MAT"
'      CODTABU1$(4) = "N-PLANO"
'      CODTABU1$(5) = "MEDIDA-1"
'      CODTABU1$(6) = "MEDIDA-2"
'      CODTABU1$(7) = "COD-MAT"
'      CODTABU1$(8) = "MEDIDA-3"
'
'      CACOLTAB1% = 8
'      '
'      For U& = 1 To List3.ListCount
'        '
'        TTXX$ = List3.List(U& - 1)
'        CAITAB1% = CAITAB1% + 1
'        TETABU1$(1, CAITAB1%) = Left$(TTXX$, 2)
'        TETABU1$(2, CAITAB1%) = Mid$(TTXX$, 4, 6)
'        TETABU1$(3, CAITAB1%) = Mid$(TTXX$, 11, 36)
'        TETABU1$(4, CAITAB1%) = Mid$(TTXX$, 48, 14)
'        TETABU1$(5, CAITAB1%) = Mid$(TTXX$, 63, 5)
'        TETABU1$(6, CAITAB1%) = Mid$(TTXX$, 69, 12)
'        TETABU1$(7, CAITAB1%) = Mid$(TTXX$, 82, 16)
'        TETABU1$(8, CAITAB1%) = Mid$(TTXX$, 98, 2)
'        '
'        'AQUI AGREGA LA DESCRIPCION LARGA
'        DELARGA$ = Mid$(TTXX$, 129)
'        Call FRATEXTO$(DELARGA$, 36)
'        For UI% = 1 To CARETEX%
'          CAITAB1% = CAITAB1% + 1
'          TETABU1$(1, CAITAB1%) = ""
'          TETABU1$(2, CAITAB1%) = ""
'          TETABU1$(3, CAITAB1%) = RETEX$(UI%)
'          TETABU1$(4, CAITAB1%) = ""
'          TETABU1$(5, CAITAB1%) = ""
'          TETABU1$(6, CAITAB1%) = ""
'          TETABU1$(7, CAITAB1%) = ""
'          TETABU1$(8, CAITAB1%) = ""
'        Next UI%
'        '
'        'ESTO ES PARA QUE EL PROXIMO EMPIECE EN UN NUEVO RENGLON TRIPLE
'        While (CAITAB1% Mod 3) <> 0
'          CAITAB1% = CAITAB1% + 1
'          TETABU1$(1, CAITAB1%) = ""
'          TETABU1$(2, CAITAB1%) = ""
'          TETABU1$(3, CAITAB1%) = ""
'          TETABU1$(4, CAITAB1%) = ""
'          TETABU1$(5, CAITAB1%) = ""
'          TETABU1$(6, CAITAB1%) = ""
'          TETABU1$(7, CAITAB1%) = ""
'          TETABU1$(8, CAITAB1%) = ""
'        Wend
'        '
'      Next U&
'      '
'      NOCONFIRM% = 1
'      Call PRINTDOC("LISTAMAT")                   ' LISTA DE MATERIALES
'      '
'      'LIMPIA LA LISTA
'      For KK1% = 1 To CACOLTAB1%
'        For KK2% = 1 To CAITAB1%
'          TETABU1$(KK1%, KK2%) = ""
'        Next KK2%
'      Next KK1%
'      '
'    End If
'    '
'199 Call CIERRARCH("*.*")
'    Call BAJALDISCO
'    Screen.MousePointer = 1
'
'End Sub
Sub CARGAEQUI(EQ1$)
  '
  Call ABREMAQUI
  With Equipos
    .FindFirst "Codigo_Equ = '" & EQ1$ & "' "
    If .NoMatch = False Then
     On Error Resume Next
      Label1 = !Descri_Equ
      SEPRO1 = !Ubica_Equ
      Label7 = FORMATNUM$(SEPRO1, "F7.3")
      '
      Label15 = !Linea_Equ
      Label20 = !Posi_Equ
      Label3 = !Planta_Equ
      '
     On Error GoTo 0
    End If
  End With
  '
End Sub
Sub INDIMAN() 'INDICE DE ORDENES DE MANTENIMIENTO
  '
  Call BLANARCH("!INDIOMAN")
  Call ABREOMAN
  '
  SQLX$ = "SELECT * FROM ORDEMANT "
  SQLX$ = SQLX$ + " ORDER BY NumeOrMa_OM ASC "
  Set OMTMP = Omantin.OpenRecordset(SQLX$, 4)
  With OMTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      GoTo 99
    End If
    On Error GoTo 0
    '
    Y$ = REGBLAN$("INDIOMAN")
    JJ& = 0
    Do While .EOF = False
      NRO& = !NumeOrMa_OM
      Call REPLA(Y$, MKS$(NRO&), 1, 4)
      DES1$ = !Codigo_Equ
      DES1$ = DES1$ + " - " + Mid$(!Referen_OM, 1, 20)
      DES1$ = DES1$ + " - " + FECHATEX$(CVINT(!FechGen_OM))
      Call REPLA(Y$, DES1$, 5, 44)
      JJ& = JJ& + 1
      Call GRAREG("INDIOMAN", Y$, JJ&)
      .MoveNext
    Loop
    '
  End With
  '
  Call SETULTREG("INDIOMAN", JJ&)
  Call CIERRARCH("INDIOMAN")
  '
99 End Sub

Private Sub Text24_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text28_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text3_Change()
  '
  NCLIE% = Val(TRIM$(Text3.TEXT))
  Label18 = RASOCLI(NCLIE% * (-1))
  MODIFIC% = 1
  '
End Sub

Private Sub Text4_Change()
  '
  Static HOX%
  If HOX% = 0 Then
    HOX% = HOY(HO$)
  End If
  '
  If FECHANUM(Text4.TEXT) > HOX% Then
    Text6.TEXT = "0:00"
  ElseIf FECHANUM(Text4.TEXT) = HOX% Then
    Text6.TEXT = Left$(Time$, 5)
  End If
  '
  MODIFIC% = 1
  '
End Sub
Sub PRIOMAN()
    ' IMPRESION DE LA O/MANTENIMIENTO
    '
    FORIPRE$ = CONTROL$("", "FORMOMAN")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = Text24.TEXT
      NUMEDOC$ = TRIM$(Label22)
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Orden de Mantenimiento"
      '
      CODPARAM$(1) = "COD-MAT"
      CODPARAM$(2) = "DES-MAT"
      CODPARAM$(3) = "REF-MAT1"
      CODPARAM$(4) = "REF-MAT2"
      CODPARAM$(5) = "REF-MAT3"
      CODPARAM$(6) = "MEDIDA-1"
      CODPARAM$(7) = "USUARIO"
      CODPARAM$(8) = "RASO-CLI"
      CODPARAM$(9) = "FECH-LAN"
      CODPARAM$(10) = "TIE-PREP"
      CODPARAM$(11) = "FECH-VEN"
      CODPARAM$(12) = "TIE-PROD"
      CAPARDOC% = 12
      '
      DOCPARAM$(1) = Text1.TEXT
      DOCPARAM$(2) = Label1
      DOCPARAM$(3) = Label3
      DOCPARAM$(4) = Label7
      DOCPARAM$(5) = Label15
      DOCPARAM$(6) = Label20
      DOCPARAM$(7) = Label24
      DOCPARAM$(8) = Label18
      DOCPARAM$(9) = Text2.TEXT
      DOCPARAM$(10) = Text5.TEXT
      DOCPARAM$(11) = Text4.TEXT
      DOCPARAM$(12) = Text6.TEXT
      '
      CODTABU2$(1) = "COD-MPRI"
      CODTABU2$(2) = "DES-MPRI"
      CODTABU2$(3) = "UNI-MPRI"
      CODTABU2$(4) = "CAN-MPRI"
      CODTABU2$(5) = "PAR-MPRI"
      CODTABU2$(6) = "NUOPERS" 'OPERARIOS
      '
      CACOLTAB2% = 6
      '
      CAITAB2% = 0
      CAITCRU% = 0
      '
      'ME FIJO SI OPERARIOS O  REPUESTOS TIENE MAS ITEMS PARA DEFINIR LA LA CACOLTAB2%
      NUOPE1& = List4.ListCount
      NUREP1& = List1.ListCount
      MAXIT1& = NUREP1&
      If NUOPE1& > NUREP1& Then MAXIT1& = NUOPE1&
      '
      For J% = 0 To MAXIT1& 'List1.ListCount - 1
         On Error Resume Next
         X$ = List1.List(J%)
         '
         CODI2$ = Mid$(X$, 1, 16)
         CI% = CODINT%(CODI2$)
         DESC2$ = DESCRIT0$(CI%)
         UM2$ = UNIMED$(CI%) ' 59, 2)
         USO2$ = Mid$(X$, 62, 10)
         CANTI = Val(USO2$)
         OBS2$ = Mid$(X$, 73, 25)
         '
         CAITAB2% = CAITAB2% + 1
         TETABU2$(1, CAITAB2%) = CODEXT$(CI%)
         TETABU2$(2, CAITAB2%) = DESC2$
         TETABU2$(3, CAITAB2%) = UM2$
         TETABU2$(4, CAITAB2%) = CANTI
         TETABU2$(5, CAITAB2%) = OBS2$
         TETABU2$(6, CAITAB2%) = List4.List(J%)
         On Error GoTo 0
         '
109   Next J%
      '
      CODTABU1$(1) = "MEDIDA-2" 'DESCRIPCION
      CACOLTAB1% = 1
      '
      CAITAB1% = 0
      DESGRAL$ = Text28.TEXT 'DESCRIPCION GENERAL
      Call FRATEXTO(DESGRAL$, 92)
      For U& = 1 To CARETEX%
        Z$ = "." + RETEX$(U&)
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = Z$
      Next U&
      '
      Call PRINTDOC("FORMOMAN")   ' Orden de Mantenimiento
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
      For KK1% = 1 To CACOLTAB2%
        For KK2% = 1 To CAITAB2%
          TETABU2$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
      DOCUORIG$ = "N/Oman-" + NUMEDOC$
      NUMEDOC$ = ""
      '
      '
     End If
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
End Sub

Private Sub Text5_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text5_LOSTFOCUS()
  '
  If Val(Text5.TEXT) > 0 Then
    Text5.TEXT = REPLACAR$(Text5.TEXT, ".", ":")
    If InStr(1, Text5.TEXT, ".") = 0 Then
      If InStr(1, Text5.TEXT, ":") = 0 Then
        Text5.TEXT = Text5.TEXT + ":0"
      End If
    End If
    '
    On Error Resume Next
      D1 = CDate(Text5.TEXT)
    On Error GoTo 0
    Text5.TEXT = Format$(D1, "hh:mm")
  Else
    Text5.TEXT = Format$(Now, "hh:mm")
  End If
  '
End Sub

Private Sub Text6_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text6_lostfocus()
  '
  If Val(Text6.TEXT) > 0 Then
    Text6.TEXT = REPLACAR$(Text6.TEXT, ".", ":")
    If InStr(1, Text6.TEXT, ".") = 0 Then
      If InStr(1, Text6.TEXT, ":") = 0 Then
        Text6.TEXT = Text6.TEXT + ":0"
      End If
    End If
    '
    On Error Resume Next
      D2 = CDate(Text6.TEXT)
    On Error GoTo 0
    Text6.TEXT = Format$(D2, "hh:mm")
  Else
    Text6.TEXT = "00.00"
  End If
  '
End Sub
Sub ALCOPCOM()
   '
   Call COPYSTRU("INDIOMAN", "INPEOMAN")
   '
   Call ABREOMAN
   '
   SQLX$ = "SELECT * FROM ORDEMANT "
   SQLX$ = SQLX$ + " ORDER BY NumeOrMa_OM ASC "
   Set OMTMP = Omantin.OpenRecordset(SQLX$, 4)
   With OMTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       GoTo 99
     End If
     On Error GoTo 0
    '
    Y$ = REGBLAN$("INPEOMAN")
    JJ& = 0
    ANU1$ = " (A)"
    CER1$ = " (C)"
    Do While .EOF = False
      NRO& = !NumeOrMa_OM
      Call REPLA(Y$, MKS$(NRO&), 1, 4)
      DES1$ = !Codigo_Equ
      DES1$ = DES1$ + " - " + Mid$(!Referen_OM, 1, 20)
      DES1$ = DES1$ + " - " + FECHATEX$(CVINT(!FechGen_OM))
      If !Status_OM < 0 Then DES1$ = DES1$ + ANU1$
      If !Status_OM = 3 Then DES1$ = DES1$ + CER1$
      Call REPLA(Y$, DES1$, 5, 44)
      JJ& = JJ& + 1
      Call GRAREG("INPEOMAN", Y$, JJ&)
      .MoveNext
    Loop
    '
  End With
  '
  Call SETULTREG("INPEOMAN", JJ&)
  Call CIERRARCH("INPEOMAN")
  '
3 YYY$ = REGSEL$("INPEOMAN/Indice General de Ordenes de Mantenimiento (Pendientes-Cumplidas-Anuladas)")
  '
  If Len(YYY$) <= 4 Then GoTo 99
  '
  Call BLANFORMU
  '
  NRO& = CVS(Left$(YYY$, 4))
  If NRO& > 0 Then
    '
    Call ABREOMAN
    '
    With OrdeMant
      .FindFirst " NumeOrMa_OM = " & NRO& & ""
      If .NoMatch = False Then
        Text1.TEXT = "" ' PARA FORZAR EL CHANGE
        Text1.TEXT = !Codigo_Equ
        NOM1& = !NumeOrMa_OM
        Text28 = !Descri_OM
        Text3 = !Proved_OM
        Text7 = !Referen_OM
        COSTER2 = !CostoTer_OM
        Text8.TEXT = FORMATNUM$(COSTER2, "F9.2")
      End If
    End With
    'CARGO LOS OPERARIOS
    List4.Clear
    With RecuProp
      SQLX$ = " NumeOrMa_OM = " & NRO& & ""
      .FindFirst SQLX$
10    If .NoMatch = False Then
          LEG1% = !LegaEmp_OM
          LEG2$ = FORMATNUM$(LEG1%, "F5.0")
          EMP1$ = ECOARCH$("PADRON", MKI$(LEG1%))
          'AGREGO AL LIST
          L1$ = Space(128)
          Call REPLA(L1$, LEG2$, 1, 8)
          Call REPLA(L1$, EMP1$, 10, 20)
          List4.AddItem L1$
        .FindNext SQLX$
        GoTo 10
      End If
    End With
    'CARGO LAS RESERVAS
    BLANARCH ("!REPUESTO")
    '
    Call ABRESERVA
    With Reserva
      SQLX$ = " VAL(MID(IdSubOr,4,6)) = " & NRO& & ""
      .FindFirst SQLX$
      JJ& = 0
15    If .NoMatch = False Then
          'AGREGO AL LIST
           CODI1% = !Cod_Inte
           CODI2$ = CODEXT$(CODI1%)
           DESC2$ = DESCRIT0(CODI1%)
           UM2$ = UNIMED$(CODI1%)
           USO1 = !UsoUnit
           USO2$ = FORMATNUM$(USO1, "F9.4")
           On Error Resume Next
           OBS2$ = !Observa
           On Error GoTo 0
           'AGREGO AL LIST
           L1$ = Space(128)
           Call REPLA(L1$, CODI2$, 1, 16)
           Call REPLA(L1$, DESC2$, 18, 40)
           Call REPLA(L1$, UM2$, 59, 2)
           Call REPLA(L1$, USO2$, 62, 10)
           Call REPLA(L1$, OBS2$, 73, 25)
           List1.AddItem L1$
           'AGREGO AL ARCHIVO
           Y2$ = REGBLAN$("REPUESTO")
           Call REPLA(Y2$, MKI$(CODI1%), 1, 2)
           Call REPLA(Y2$, UM2$, 3, 2)
           Call REPLA(Y2$, MKS(USO1), 5, 4)
           Call REPLA(Y2$, OBS2$, 9, 25)
           JJ& = JJ& + 1
           Call GRAREG("!REPUESTO", Y2$, JJ&)
        .FindNext SQLX$
        GoTo 15
      End If
    End With
    Call SETULTREG("!REPUESTO", JJ&)
    Call CIERRARCH("!REPUESTO")
  End If
  '
99 End Sub

Private Sub text7_change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text8_LOSTFOCUS()
  COESTI1 = Val(Text8.TEXT)
  Text8 = FORMATNUM$(COESTI1, "F9.2")
  MODIFIC% = 1
  '
End Sub
