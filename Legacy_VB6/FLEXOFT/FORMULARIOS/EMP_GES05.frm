VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EMP_GES05 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8D8C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empleados - A-B-M-L Base de Datos de Empleados"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "EMP_GES05.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   154
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
      Height          =   300
      Left            =   7275
      Picture         =   "EMP_GES05.frx":0442
      TabIndex        =   4
      Top             =   315
      Visible         =   0   'False
      Width           =   175
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6915
      TabIndex        =   62
      Top             =   315
      Width           =   405
   End
   Begin VB.CommandButton Command10 
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
      Height          =   315
      Left            =   3480
      Picture         =   "EMP_GES05.frx":074C
      TabIndex        =   3
      Top             =   315
      Width           =   175
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      MaxLength       =   4
      TabIndex        =   2
      Top             =   315
      Width           =   405
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D8D8C0&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   880
      Left            =   8820
      TabIndex        =   57
      Top             =   210
      Width           =   1990
      Begin VB.CommandButton Command20 
         Height          =   500
         Left            =   1280
         Picture         =   "EMP_GES05.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Búsqueda por  Apellido, Nombre, Documento, etc"
         Top             =   250
         Width           =   550
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   155
         Picture         =   "EMP_GES05.frx":1DBC
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   250
         Width           =   550
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   725
         Picture         =   "EMP_GES05.frx":20C6
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Restablece los Valores Anteriores - Descarta Cambios Efectuados"
         Top             =   250
         Width           =   550
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   8100
      Index           =   0
      Left            =   11025
      ScaleHeight     =   8040
      ScaleWidth      =   1215
      TabIndex        =   53
      Top             =   0
      Width           =   1275
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   55
         Top             =   7140
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   56
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Print"
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
         Picture         =   "EMP_GES05.frx":23D0
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Imprime Ficha del Empleado"
         Top             =   5015
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "List"
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
         Picture         =   "EMP_GES05.frx":2A3A
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Emisión de Listados"
         Top             =   4320
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Picture         =   "EMP_GES05.frx":2D44
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   6315
         Width           =   650
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
         Picture         =   "EMP_GES05.frx":3186
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8100
         Width           =   650
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Baja"
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
         Picture         =   "EMP_GES05.frx":3490
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Baja/Restauración Lógica del Legajo"
         Top             =   3085
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Picture         =   "EMP_GES05.frx":379A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command4 
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
         Picture         =   "EMP_GES05.frx":38E4
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ayuda en Línea"
         Top             =   935
         Width           =   650
      End
      Begin VB.CommandButton Command21 
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
         Height          =   690
         Left            =   105
         Picture         =   "EMP_GES05.frx":3BEE
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Alta de Nuevo Legajo"
         Top             =   2400
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "EMP_GES05.frx":3D38
         Stretch         =   -1  'True
         Top             =   7455
         Width           =   1410
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   840
      Left            =   10710
      Sorted          =   -1  'True
      TabIndex        =   46
      Top             =   3780
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1155
      TabIndex        =   5
      Top             =   735
      Width           =   7470
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1155
      MaxLength       =   5
      TabIndex        =   0
      Text            =   " "
      Top             =   315
      Width           =   780
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
      Height          =   330
      Left            =   1935
      Picture         =   "EMP_GES05.frx":5C5A
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6795
      Left            =   120
      TabIndex        =   35
      Top             =   1155
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11986
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   14211264
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "EMP_GES05.frx":5F64
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Anotaciones"
      TabPicture(1)   =   "EMP_GES05.frx":5F80
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(1)=   "Text28"
      Tab(1).Control(2)=   "Command17"
      Tab(1).Control(3)=   "Command8"
      Tab(1).Control(4)=   "Command6"
      Tab(1).Control(5)=   "Command5"
      Tab(1).Control(6)=   "Command7"
      Tab(1).Control(7)=   "Picture4"
      Tab(1).Control(8)=   "Line3"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Datos Familiares"
      TabPicture(2)   =   "EMP_GES05.frx":5F9C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame18"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Conceptos Opcionales"
      TabPicture(3)   =   "EMP_GES05.frx":5FB8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Line14"
      Tab(3).Control(1)=   "Line15"
      Tab(3).Control(2)=   "Picture3"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Transferencia DGI"
      TabPicture(4)   =   "EMP_GES05.frx":5FD4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame22"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame22 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   -74790
         TabIndex        =   120
         Top             =   525
         Width           =   10215
         Begin VB.Frame Frame24 
            Caption         =   "Otros Datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5430
            Left            =   210
            TabIndex        =   121
            Top             =   315
            Width           =   9810
            Begin VB.TextBox Text35 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7770
               MaxLength       =   1
               TabIndex        =   146
               Text            =   " "
               Top             =   4305
               Width           =   1185
            End
            Begin VB.TextBox Text34 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7770
               MaxLength       =   4
               TabIndex        =   144
               Text            =   " "
               Top             =   3510
               Width           =   1200
            End
            Begin VB.TextBox Text33 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7770
               MaxLength       =   9
               TabIndex        =   142
               Top             =   2565
               Width           =   1185
            End
            Begin VB.TextBox Text32 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7770
               MaxLength       =   4
               TabIndex        =   140
               Text            =   " "
               Top             =   1575
               Width           =   1185
            End
            Begin VB.TextBox Text21 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7770
               MaxLength       =   4
               TabIndex        =   138
               Text            =   " "
               Top             =   675
               Width           =   1185
            End
            Begin VB.TextBox Text31 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2100
               MaxLength       =   4
               TabIndex        =   136
               Top             =   4305
               Width           =   1185
            End
            Begin VB.TextBox Text30 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2100
               MaxLength       =   6
               TabIndex        =   134
               Text            =   " "
               Top             =   3825
               Width           =   1185
            End
            Begin VB.TextBox Text24 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2100
               MaxLength       =   4
               TabIndex        =   132
               Text            =   " "
               Top             =   3300
               Width           =   1185
            End
            Begin VB.TextBox Text23 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2100
               MaxLength       =   4
               TabIndex        =   130
               Text            =   " "
               Top             =   2775
               Width           =   1185
            End
            Begin VB.TextBox Text18 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2100
               MaxLength       =   4
               TabIndex        =   128
               Text            =   " "
               Top             =   2250
               Width           =   1185
            End
            Begin VB.TextBox Text27 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2100
               MaxLength       =   4
               TabIndex        =   124
               Top             =   1200
               Width           =   1185
            End
            Begin VB.TextBox Text22 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2100
               MaxLength       =   4
               TabIndex        =   123
               Text            =   " "
               Top             =   1725
               Width           =   1185
            End
            Begin VB.TextBox Text20 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2100
               MaxLength       =   4
               TabIndex        =   122
               Text            =   " "
               Top             =   675
               Width           =   1185
            End
            Begin VB.Line Line18 
               X1              =   4850
               X2              =   4850
               Y1              =   450
               Y2              =   5040
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Operación:"
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
               Left            =   6195
               TabIndex        =   147
               Top             =   4230
               Width           =   1380
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Trabajador Convencionado:"
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
               Left            =   6195
               TabIndex        =   145
               Top             =   3465
               Width           =   1380
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Capital  Recomposición de LRT:"
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
               Left            =   6195
               TabIndex        =   143
               Top             =   2310
               Width           =   1380
            End
            Begin VB.Label Label51 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Marca Corresponde Reducción:"
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
               Left            =   6195
               TabIndex        =   141
               Top             =   1320
               Width           =   1380
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Código de Siniestrado:"
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
               Left            =   6195
               TabIndex        =   139
               Top             =   630
               Width           =   1380
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "Adherentes:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   630
               TabIndex        =   137
               Top             =   4425
               Width           =   1290
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Código de  Obra Social:"
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
               Left            =   525
               TabIndex        =   135
               Top             =   3780
               Width           =   1380
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Modalidad de Contratación:"
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
               Left            =   525
               TabIndex        =   133
               Top             =   3255
               Width           =   1380
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "% Aporte Adicional SS:"
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
               Left            =   525
               TabIndex        =   131
               Top             =   2730
               Width           =   1380
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Código de Condición:"
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
               Left            =   525
               TabIndex        =   129
               Top             =   2205
               Width           =   1380
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "Hijos:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   630
               TabIndex        =   127
               Top             =   1320
               Width           =   1290
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Código de Situación:"
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
               Left            =   525
               TabIndex        =   126
               Top             =   1680
               Width           =   1380
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cónyuge:"
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
               Left            =   0
               TabIndex        =   125
               Top             =   780
               Width           =   1905
            End
         End
         Begin VB.Line Line17 
            BorderColor     =   &H00E0E0E0&
            X1              =   5280
            X2              =   5280
            Y1              =   240
            Y2              =   3600
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   4220
         Left            =   -74160
         ScaleHeight     =   4155
         ScaleWidth      =   8895
         TabIndex        =   66
         Top             =   1520
         Width           =   8955
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Index           =   0
            Left            =   0
            TabIndex        =   67
            Top             =   0
            Width           =   10515
            Begin VB.Line Line12 
               X1              =   720
               X2              =   720
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line10 
               X1              =   6000
               X2              =   6000
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line9 
               X1              =   5280
               X2              =   5280
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line8 
               X1              =   4680
               X2              =   4680
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "  Cod.   Descripción                                             U/M      Can.  Forma de Cálculo"
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
               TabIndex        =   68
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line6 
               X1              =   8520
               X2              =   8520
               Y1              =   0
               Y2              =   1050
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0F0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5460
            Left            =   0
            TabIndex        =   69
            Top             =   -1050
            Width           =   10515
            Begin VB.Frame Frame21 
               Caption         =   "Frame11"
               Height          =   4000
               Left            =   8520
               TabIndex        =   101
               Top             =   1260
               Width           =   15
            End
            Begin VB.Frame Frame20 
               Caption         =   "Frame11"
               Height          =   4000
               Left            =   6000
               TabIndex        =   100
               Top             =   1260
               Width           =   15
            End
            Begin VB.Frame Frame19 
               Caption         =   "Frame11"
               Height          =   4000
               Left            =   5280
               TabIndex        =   99
               Top             =   1260
               Width           =   15
            End
            Begin VB.Frame Frame15 
               Caption         =   "Frame11"
               Height          =   4000
               Left            =   4680
               TabIndex        =   98
               Top             =   1260
               Width           =   15
            End
            Begin VB.Frame Frame14 
               Caption         =   "Frame10"
               Height          =   4000
               Left            =   720
               TabIndex        =   97
               Top             =   1260
               Width           =   15
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
               Height          =   3660
               Left            =   -30
               TabIndex        =   70
               Top             =   1365
               Width           =   8955
            End
            Begin VB.Line Line11 
               X1              =   1155
               X2              =   1155
               Y1              =   1470
               Y2              =   3255
            End
         End
      End
      Begin VB.Frame Frame18 
         BorderStyle     =   0  'None
         Caption         =   "Frame18"
         Height          =   5900
         Left            =   -74600
         TabIndex        =   47
         Top             =   740
         Width           =   9990
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00E0E0E0&
            Height          =   4220
            Left            =   950
            ScaleHeight     =   4155
            ScaleWidth      =   8055
            TabIndex        =   48
            Top             =   900
            Width           =   8120
            Begin VB.Frame Frame16 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               Height          =   330
               Left            =   0
               TabIndex        =   49
               Top             =   0
               Width           =   10515
               Begin VB.Line Line5 
                  X1              =   7680
                  X2              =   7680
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Label Label37 
                  BackStyle       =   0  'Transparent
                  Caption         =   " Nombre y Apellido                         F. Nacim.      Parentesco    Documento"
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
                  TabIndex        =   50
                  Top             =   50
                  Width           =   10410
               End
               Begin VB.Line Line7 
                  X1              =   3250
                  X2              =   3250
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line13 
                  X1              =   4775
                  X2              =   4775
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line20 
                  X1              =   6200
                  X2              =   6200
                  Y1              =   0
                  Y2              =   1050
               End
            End
            Begin VB.Frame Frame17 
               BackColor       =   &H00C0F0FF&
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   5460
               Left            =   0
               TabIndex        =   51
               Top             =   -1050
               Width           =   10515
               Begin VB.Frame Frame13 
                  Caption         =   "Frame11"
                  Height          =   4000
                  Left            =   7680
                  TabIndex        =   96
                  Top             =   1260
                  Width           =   15
               End
               Begin VB.Frame Frame12 
                  Caption         =   "Frame11"
                  Height          =   4000
                  Left            =   6200
                  TabIndex        =   95
                  Top             =   1260
                  Width           =   15
               End
               Begin VB.Frame Frame11 
                  Caption         =   "Frame11"
                  Height          =   4000
                  Left            =   4770
                  TabIndex        =   94
                  Top             =   1260
                  Width           =   15
               End
               Begin VB.Frame Frame10 
                  Caption         =   "Frame10"
                  Height          =   4000
                  Left            =   3255
                  TabIndex        =   93
                  Top             =   1260
                  Width           =   15
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
                  Height          =   3660
                  Left            =   -30
                  TabIndex        =   52
                  Top             =   1365
                  Width           =   8115
               End
               Begin VB.Line Line21 
                  X1              =   1155
                  X2              =   1155
                  Y1              =   1470
                  Y2              =   3255
               End
            End
         End
         Begin VB.Line Line16 
            X1              =   945
            X2              =   9080
            Y1              =   5110
            Y2              =   5110
         End
         Begin VB.Line Line4 
            X1              =   9070
            X2              =   9070
            Y1              =   900
            Y2              =   5100
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   -74550
         ScaleHeight     =   375
         ScaleWidth      =   9720
         TabIndex        =   44
         Top             =   1455
         Width           =   9780
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Anotaciones - Novedades - Comentarios"
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
            Left            =   105
            TabIndex        =   45
            Top             =   105
            Width           =   9570
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6000
         Left            =   200
         TabIndex        =   42
         Top             =   600
         Width           =   10320
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6015
            Left            =   120
            TabIndex        =   43
            Top             =   -120
            Width           =   10215
            Begin VB.TextBox Text14 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   5775
               TabIndex        =   152
               Text            =   " "
               Top             =   5775
               Visible         =   0   'False
               Width           =   3585
            End
            Begin VB.Frame Frame9 
               Caption         =   "Datos Gráficos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2955
               Left            =   5520
               TabIndex        =   87
               Top             =   240
               Width           =   4575
               Begin VB.CommandButton Command222 
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   280
                  Left            =   3990
                  TabIndex        =   92
                  ToolTipText     =   "Click Aqui para Eliminar Imagen"
                  Top             =   525
                  Width           =   280
               End
               Begin VB.Frame Frame1 
                  Caption         =   "Fotografía"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2340
                  Index           =   2
                  Left            =   630
                  TabIndex        =   88
                  Top             =   420
                  Width           =   3300
                  Begin VB.FileListBox File1 
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1260
                     Index           =   2
                     Left            =   210
                     Pattern         =   "*.JPG"
                     TabIndex        =   89
                     Top             =   630
                     Visible         =   0   'False
                     Width           =   2755
                  End
                  Begin VB.PictureBox Picture1 
                     AutoRedraw      =   -1  'True
                     Height          =   1380
                     Index           =   2
                     Left            =   210
                     ScaleHeight     =   1320
                     ScaleWidth      =   2850
                     TabIndex        =   90
                     Top             =   735
                     Width           =   2905
                  End
                  Begin VB.CommandButton Command53 
                     Caption         =   "."
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
                     Left            =   2950
                     TabIndex        =   19
                     Top             =   315
                     Width           =   175
                  End
                  Begin VB.Image Image1 
                     Height          =   19995
                     Index           =   2
                     Left            =   0
                     Stretch         =   -1  'True
                     Top             =   105
                     Visible         =   0   'False
                     Width           =   19995
                  End
                  Begin VB.Label Label34 
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
                     Index           =   2
                     Left            =   210
                     TabIndex        =   91
                     Top             =   315
                     Width           =   2755
                  End
               End
            End
            Begin VB.Frame Frame8 
               Caption         =   "Datos Personales"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3375
               Left            =   110
               TabIndex        =   77
               Top             =   240
               Width           =   5295
               Begin VB.TextBox Text38 
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
                  Left            =   1365
                  TabIndex        =   156
                  Text            =   " "
                  Top             =   2880
                  Width           =   3690
               End
               Begin VB.TextBox Text36 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1365
                  TabIndex        =   148
                  Top             =   2520
                  Width           =   1170
               End
               Begin VB.CommandButton Command34 
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
                  Left            =   1680
                  Picture         =   "EMP_GES05.frx":5FF0
                  TabIndex        =   13
                  Top             =   1800
                  Width           =   150
               End
               Begin VB.CommandButton Command23 
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
                  Left            =   1695
                  Picture         =   "EMP_GES05.frx":62FA
                  TabIndex        =   8
                  Top             =   720
                  Width           =   150
               End
               Begin VB.TextBox Text29 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1365
                  TabIndex        =   11
                  Text            =   " "
                  Top             =   1440
                  Width           =   3690
               End
               Begin VB.TextBox Text26 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1365
                  MaxLength       =   1
                  TabIndex        =   7
                  Text            =   " "
                  Top             =   720
                  Width           =   330
               End
               Begin VB.TextBox Text25 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1365
                  TabIndex        =   6
                  Top             =   360
                  Width           =   3660
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
                  Height          =   315
                  Left            =   1365
                  MaxLength       =   1
                  TabIndex        =   12
                  Text            =   " "
                  Top             =   1800
                  Width           =   345
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
                  Height          =   315
                  Left            =   1365
                  TabIndex        =   14
                  Text            =   " "
                  Top             =   2160
                  Width           =   3690
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
                  Height          =   315
                  Left            =   2625
                  TabIndex        =   15
                  Text            =   " "
                  Top             =   2520
                  Width           =   2430
               End
               Begin VB.CommandButton Command16 
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
                  Left            =   2475
                  Picture         =   "EMP_GES05.frx":6604
                  TabIndex        =   10
                  Top             =   1080
                  Width           =   150
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
                  Height          =   315
                  Left            =   1365
                  TabIndex        =   9
                  Top             =   1080
                  Width           =   1125
               End
               Begin VB.Label Label44 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Teléfono:"
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
                  Left            =   -620
                  TabIndex        =   157
                  Top             =   3000
                  Width           =   1905
               End
               Begin VB.Label Label22 
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
                  Height          =   315
                  Left            =   2040
                  TabIndex        =   86
                  Top             =   1800
                  Width           =   2985
               End
               Begin VB.Label Label30 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nacionalidad:"
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
                  Left            =   -600
                  TabIndex        =   85
                  Top             =   1560
                  Width           =   1905
               End
               Begin VB.Label Label27 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sexo:"
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
                  Left            =   -630
                  TabIndex        =   84
                  Top             =   830
                  Width           =   1905
               End
               Begin VB.Label Label26 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Documento:"
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
                  Left            =   -630
                  TabIndex        =   83
                  Top             =   465
                  Width           =   1905
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Estado Civil:"
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
                  Left            =   -600
                  TabIndex        =   82
                  Top             =   1920
                  Width           =   1905
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "CP-Localidad:"
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
                  Left            =   -600
                  TabIndex        =   81
                  Top             =   2640
                  Width           =   1905
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFF80&
                  BackStyle       =   0  'Transparent
                  Caption         =   "F. Nacim.:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   0
                  TabIndex        =   80
                  Top             =   1200
                  Width           =   1290
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Domicilio:"
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
                  Left            =   -600
                  TabIndex        =   79
                  Top             =   2280
                  Width           =   1905
               End
               Begin VB.Label Label19 
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
                  Height          =   315
                  Left            =   2040
                  TabIndex        =   78
                  Top             =   720
                  Width           =   2985
               End
            End
            Begin VB.Frame Frame7 
               Caption         =   "Datos Laborales"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2175
               Left            =   110
               TabIndex        =   72
               Top             =   3690
               Width           =   5295
               Begin VB.TextBox Text37 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   3360
                  TabIndex        =   150
                  Text            =   " "
                  Top             =   420
                  Width           =   1485
               End
               Begin VB.CommandButton Command26 
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
                  Height          =   315
                  Left            =   2245
                  Picture         =   "EMP_GES05.frx":690E
                  TabIndex        =   117
                  Top             =   420
                  Width           =   175
               End
               Begin VB.TextBox Text17 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1380
                  MaxLength       =   7
                  TabIndex        =   116
                  Top             =   420
                  Width           =   930
               End
               Begin VB.TextBox Text8 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1380
                  TabIndex        =   16
                  Text            =   " "
                  Top             =   1230
                  Width           =   3480
               End
               Begin VB.TextBox Text16 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1380
                  MaxLength       =   3
                  TabIndex        =   17
                  Top             =   1620
                  Width           =   405
               End
               Begin VB.CommandButton Command19 
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
                  Height          =   315
                  Left            =   1785
                  Picture         =   "EMP_GES05.frx":6C18
                  TabIndex        =   18
                  Top             =   1620
                  Width           =   175
               End
               Begin VB.Label Label39 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Horario:"
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
                  Left            =   1365
                  TabIndex        =   149
                  Top             =   495
                  Width           =   1905
               End
               Begin VB.Label Label31 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1380
                  TabIndex        =   119
                  Top             =   840
                  Width           =   3480
               End
               Begin VB.Label Label29 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFF80&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sector:"
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
                  TabIndex        =   118
                  Top             =   525
                  Width           =   1110
               End
               Begin VB.Label Label9 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tarea:"
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
                  Left            =   -585
                  TabIndex        =   75
                  Top             =   1305
                  Width           =   1905
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFF80&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Categoría:"
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
                  TabIndex        =   74
                  Top             =   1725
                  Width           =   1110
               End
               Begin VB.Label Label23 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   2085
                  TabIndex        =   73
                  Top             =   1620
                  Width           =   2775
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "Otros"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2595
               Left            =   5520
               TabIndex        =   71
               Top             =   3255
               Width           =   4575
               Begin VB.TextBox Text142 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   2730
                  MaxLength       =   13
                  TabIndex        =   107
                  Top             =   840
                  Width           =   1665
               End
               Begin VB.TextBox Text141 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   840
                  TabIndex        =   108
                  Text            =   " "
                  Top             =   1260
                  Width           =   3585
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
                  Height          =   300
                  Left            =   1785
                  Picture         =   "EMP_GES05.frx":6F22
                  TabIndex        =   103
                  Top             =   420
                  Width           =   150
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
                  Height          =   300
                  Left            =   1500
                  Picture         =   "EMP_GES05.frx":722C
                  TabIndex        =   110
                  Top             =   2100
                  Width           =   150
               End
               Begin VB.TextBox Text12 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   390
                  TabIndex        =   109
                  Top             =   2085
                  Width           =   1125
               End
               Begin VB.TextBox Text9 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   855
                  TabIndex        =   102
                  Top             =   420
                  Width           =   1020
               End
               Begin VB.CommandButton Command22 
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
                  Height          =   315
                  Left            =   1785
                  Picture         =   "EMP_GES05.frx":7536
                  TabIndex        =   105
                  Top             =   825
                  Width           =   150
               End
               Begin VB.TextBox Text5 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   855
                  TabIndex        =   104
                  Top             =   825
                  Width           =   1020
               End
               Begin VB.TextBox Text15 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   2730
                  MaxLength       =   13
                  TabIndex        =   106
                  Top             =   420
                  Width           =   1665
               End
               Begin VB.TextBox Text13 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1965
                  TabIndex        =   112
                  Text            =   " "
                  Top             =   2100
                  Width           =   2430
               End
               Begin VB.Label Label42 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "AFJP:"
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
                  Left            =   735
                  TabIndex        =   153
                  Top             =   945
                  Width           =   1905
               End
               Begin VB.Label Label15 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Contr:"
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
                  Left            =   -120
                  TabIndex        =   151
                  Top             =   1320
                  Width           =   930
               End
               Begin VB.Line Line23 
                  BorderColor     =   &H00FFFFFF&
                  X1              =   0
                  X2              =   4515
                  Y1              =   1740
                  Y2              =   1740
               End
               Begin VB.Line Line19 
                  BorderColor     =   &H80000003&
                  X1              =   0
                  X2              =   4620
                  Y1              =   1720
                  Y2              =   1720
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFF80&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Egreso:"
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
                  Left            =   -255
                  TabIndex        =   115
                  Top             =   1845
                  Width           =   1050
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFF80&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ingreso:"
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
                  Left            =   -255
                  TabIndex        =   114
                  Top             =   510
                  Width           =   1050
               End
               Begin VB.Label Label25 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFF80&
                  BackStyle       =   0  'Transparent
                  Caption         =   "F.Antig:"
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
                  Left            =   -120
                  TabIndex        =   113
                  Top             =   915
                  Width           =   930
               End
               Begin VB.Label Label28 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "CUIL:"
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
                  Left            =   735
                  TabIndex        =   111
                  Top             =   525
                  Width           =   1905
               End
               Begin VB.Label Label14 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Motivo:"
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
                  Left            =   1365
                  TabIndex        =   76
                  Top             =   1845
                  Width           =   1050
               End
            End
            Begin VB.TextBox Text11 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   4245
               TabIndex        =   61
               Top             =   5775
               Visible         =   0   'False
               Width           =   1125
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
               Height          =   300
               Left            =   5355
               Picture         =   "EMP_GES05.frx":7840
               TabIndex        =   60
               Top             =   5760
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.Line Line22 
               BorderColor     =   &H00E0E0E0&
               X1              =   5280
               X2              =   5280
               Y1              =   240
               Y2              =   3600
            End
            Begin VB.Line Line2 
               X1              =   5280
               X2              =   5280
               Y1              =   240
               Y2              =   3600
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "F.Cat.:"
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
               Left            =   3360
               TabIndex        =   59
               Top             =   5880
               Visible         =   0   'False
               Width           =   810
            End
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
         Height          =   4215
         Left            =   -74550
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   1875
         Width           =   9780
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         Height          =   260
         Left            =   -69690
         TabIndex        =   31
         Top             =   6090
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         Height          =   260
         Left            =   -68505
         TabIndex        =   32
         Top             =   6090
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pegar"
         Height          =   260
         Left            =   -67335
         TabIndex        =   33
         Top             =   6090
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Left            =   -66045
         TabIndex        =   34
         Top             =   6090
         Width           =   1290
      End
      Begin VB.CommandButton Command7 
         Height          =   260
         Left            =   -74550
         TabIndex        =   41
         Top             =   6090
         Width           =   4830
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         Height          =   1590
         Left            =   -72690
         ScaleHeight     =   1530
         ScaleWidth      =   2055
         TabIndex        =   40
         Top             =   2175
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Line Line15 
         X1              =   -74160
         X2              =   -65210
         Y1              =   5730
         Y2              =   5730
      End
      Begin VB.Line Line14 
         X1              =   -65200
         X2              =   -65200
         Y1              =   1520
         Y2              =   5760
      End
      Begin VB.Line Line3 
         X1              =   -64770
         X2              =   -64770
         Y1              =   1455
         Y2              =   6300
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   3795
         TabIndex        =   39
         Top             =   4905
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   4215
         TabIndex        =   38
         Top             =   3960
         Width           =   15
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "EMP_GES05.frx":7B4A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "EMP_GES05.frx":7D7E
      TabIndex        =   155
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7530
      TabIndex        =   65
      Top             =   315
      Width           =   1095
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   64
      Top             =   315
      Width           =   2415
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
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
      Left            =   5760
      TabIndex        =   63
      Top             =   420
      Width           =   1110
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo:"
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
      Left            =   1965
      TabIndex        =   58
      Top             =   420
      Width           =   1110
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido y Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   -315
      TabIndex        =   37
      Top             =   680
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de Legajo:"
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
      Left            =   -105
      TabIndex        =   36
      Top             =   240
      Width           =   1140
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu NuevEmp 
         Caption         =   "Nuevo Legajo"
      End
      Begin VB.Menu EditEmp 
         Caption         =   "Abrir Existente"
      End
      Begin VB.Menu ImpoEmpleados 
         Caption         =   "Importacion de Empleados"
      End
      Begin VB.Menu Impofamiliares 
         Caption         =   "Importacion de Familiares"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu GuaEmp 
         Caption         =   "Guardar"
      End
      Begin VB.Menu php1a 
         Caption         =   "-"
      End
      Begin VB.Menu Baj_Emp 
         Caption         =   "Baja/Restauración Lógica del Legajo"
      End
      Begin VB.Menu PLh2 
         Caption         =   "-"
      End
      Begin VB.Menu Print 
         Caption         =   "Imprimir Ficha"
      End
      Begin VB.Menu plh3 
         Caption         =   "-"
      End
      Begin VB.Menu List 
         Caption         =   "Consultas y Listados"
      End
      Begin VB.Menu plh4 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu EdiEmp 
      Caption         =   "Edición"
      Begin VB.Menu DeHaCaEmp 
         Caption         =   "Deshacer Cambios"
      End
      Begin VB.Menu php5a 
         Caption         =   "-"
      End
      Begin VB.Menu BusEmp 
         Caption         =   "Buscar"
      End
   End
   Begin VB.Menu ConLiEmp 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu LiGeClaEmp 
         Caption         =   "Listados"
      End
   End
   Begin VB.Menu ConfEmp 
      Caption         =   "Configuración"
      Begin VB.Menu ConGenEmp 
         Caption         =   "Configuración General"
      End
      Begin VB.Menu PHP35 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGrupEmp 
         Caption         =   "Grupos de Empleados"
      End
      Begin VB.Menu mnuCatego 
         Caption         =   "Categorías de Empleados"
      End
      Begin VB.Menu mnuSalNomBas 
         Caption         =   "Salarios Nominales Básicos"
      End
      Begin VB.Menu mnuParInd 
         Caption         =   "Parámetros Individuales"
      End
      Begin VB.Menu mnuRemOrd 
         Caption         =   "Remuneraciones Ordinarias"
      End
      Begin VB.Menu mnuRetLegApo 
         Caption         =   "Retenciones Legales y Aportes"
      End
      Begin VB.Menu mnuAsiSubFam 
         Caption         =   "Asignaciones y Subsidios Familiares"
      End
   End
   Begin VB.Menu AyuCli 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "EMP_GES05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MODIFIC% 'MARCA DE MODIFICACION
Public LEGACT% 'LEGAJO ACTUAL
Dim TOPEINI(3), IZQUINI(3), ANCHINI(3), ALTUINI(3)
Private objexcel As Excel.Application
Private Libro As Excel.Workbook
Private HOJA As Excel.Worksheet

Sub ActualizaTablaDatosAuxiliaresDeOperarios()

    Dim da As New DatosAuxiliares
    
    Call da.CrearTablasAuxiliares
    idTabla& = da.ObtenerIdTablaAuxiliar("TABOPERA")
    Call BORRAREGISTROS("datasql", "id_lista=" & CStr(idTabla&), REGAF&)
    
    For i& = 1 To ULTREG&("PADREMP")
        RegOperario$ = REGLEIDO$("PADREMP", i&)
        LEGAJO% = CVI(Mid$(RegOperario$, 1, 2))
        nombre$ = TRIM$(Mid$(RegOperario$, 3, 30))
        
        ConsultaSql$ = "insert into datasql (id_lista, lista, codiempr, campo1, campo2)" _
        & " values (" & CStr(idTabla&) & ",'TABOPERA'," & CStr(CODIEMPR%) & ",'" & CStr(LEGAJO%) & "','" & nombre$ & "')"
        FLEXCONN.Execute ConsultaSql$
    Next i&
    
    Call CIERRARCH("PADREMP")
    
End Sub

Private Sub AyuCli_Click()
  '
  Call Command4_Click
  '
End Sub

Private Sub Baj_Emp_Click()
  '
  Call Command24_Click
  '
End Sub

Private Sub BusEmp_Click()
  '
  Call Command20_Click
  '
End Sub

Private Sub Command1_Click()
  '
  If Val(Text1) > 0 Then
    Call GRAEMP(Val(Text1))
    Call GRANOTEMP 'GRABA ANOTACIONES DE EMPLEADOS
    MODIFIC% = 0
  End If
  '
End Sub

Private Sub Command10_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("DATGREM1")
     VDEVU$ = TRIM$(VALACT1$("DATGREM1"))
     If VDEVU$ <> "" Then
        Text6.TEXT = VALACT1$("DATGREM1")
     End If
  End If
End Sub

Private Sub Command11_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("SELFECHA")
     VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
     If VDEVU$ <> "" Then
        Text9.TEXT = VALACT1$("SELFECHA")
     End If
  End If
End Sub

Private Sub Command12_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("SELFECHA")
     VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
     If VDEVU$ <> "" Then
        Text11.TEXT = VALACT1$("SELFECHA")
     End If
  End If
End Sub

 Sub Command13_Click()
  '
  Command13.Enabled = False
  RESP1% = COMALTER%("Seleccione el Listado:\\Listado de Padrón\Listado de Familiares")
  If RESP1% = 1 Then
    Call FINAL("*PADREMP")
  ElseIf RESP1% = 2 Then
    Call LISFAMIL
  End If
  Command13.Enabled = True
  '
End Sub
Sub LISFAMIL()
  '
  Call REGBLAN$("LISFAMIL")
  Y$ = REGBLAN("LISFAMIL")
  J2& = 0
  For J& = 1 To ULTREG("PADREMP")
    X$ = REGLEIDO$("PADREMP", J&)
    NELGUI% = CVI(Mid$(X$, 1, 2))
    If NELGUI% <= 0 Then Exit Sub
    '
    Call TRALOFIL("FAMCAR", 1, MKI$(NELGUI%))
    '
    For J1& = 1 To ULTREG&("!FAMCAR")
        XX$ = REGLEIDO$("!FAMCAR", J1&)
        NOAP$ = Mid$(XX$, 3, 25)
        FENA% = CVI(Mid$(XX$, 28, 2))
        PAREN$ = Mid$(XX$, 30, 10)
        DOCUFA$ = Mid$(XX$, 40, 11)
        '
        Call REPLA(Y$, MKI$(NELGUI%), 1, 2)
        Call REPLA(Y$, NOAP$, 3, 25)
        Call REPLA(Y$, MKI$(FENA%), 28, 2)
        Call REPLA(Y$, PAREN$, 30, 10)
        Call REPLA(Y$, DOCUFA$, 40, 11)
        J2& = J2& + 1
        Call GRAREG("LISFAMIL", Y$, J2&)
    Next J1&
  Next J&
  Call SETULTREG("LISFAMIL", J2&)
  '
  Call FINAL("*LISFAMIL")
  '
End Sub
Sub Command14_Click()
  '
  OPSUEL05.Show 1
  '
End Sub

Private Sub Command16_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("SELFECHA")
     VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
     If VDEVU$ <> "" Then
        Text10.TEXT = VALACT1$("SELFECHA")
     End If
  End If
End Sub

Private Sub Command18_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("SELFECHA")
     VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
     If VDEVU$ <> "" Then
        Text12.TEXT = VALACT1$("SELFECHA")
     End If
  End If
End Sub

Private Sub Command19_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("CATEGO")
     VDEVU$ = TRIM$(VALACT1$("CATEGO"))
     If VDEVU$ <> "" Then
        Text16.TEXT = VALACT1$("CATEGO")
     End If
  End If
End Sub

Private Sub Command2_Click()
  '
  NEMP1$ = Text1.TEXT
  If COMALTER%("Esta Seguro de Des-Hacer los Cambios del Empleado\'" + TRIM$(Text3.TEXT) + "'\\Des-Hacer Cambios\Cancelar") = 1 Then
     Screen.MousePointer = 11
     MODIFIC% = 0
     Call CARDATEMP(Val(Text1))
     MODIFIC% = 0
     Screen.MousePointer = 1
  End If
  Text1.TEXT = NEMP1$
  '
End Sub

Private Sub Command20_Click()
  '
  PALBU$ = InputBox("Ingrese el Texto a Buscar", "Busqueda de Empleados")
  If TRIM(PALBU$ = "") Then Exit Sub
  PALBU$ = UCase$(PALBU$)
  ENCO1% = 0
  Call BLANARCH("!PADREMP")
  For J& = 1 To ULTREG&("PADREMP")
    X$ = REGLEIDO$("PADREMP", J&)
    NLEG% = CVI(Mid$(X$, 1, 2))
    X1$ = UCase$(X$)
    Y$ = REGBLAN$("PADREMP")
    If InStr(1, X1$, PALBU$) > 0 Then
      Call REPLA$(Y$, X$, 1, 256)
      Call REGAPP("!PADREMP", Y$)
      ENCO1% = 1
    End If
  Next J&
  '
  If ENCO1% = 1 Then
       '
    If MODIFIC% = 1 And Val(Text1) > 0 And Val(Text1) < 32767 Then
     If COMALTER%("Se han Modificado Datos del Empleado\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") = 1 Then
       Screen.MousePointer = 11
       MODIFIC% = 0
       Call GRAEMP(Val(Text1))
       Screen.MousePointer = 1
     End If
    End If
    '
    If MODIFIC% = 1 Then
     Call GRANOTEMP 'GRABA ANOTACIONES DE EMPLEADOS
    End If
    '
    Call FRESHECHO("!PADREMP")
    Call SELECHO("!PADREMP")
    NLEG% = Val(TRIM$(VALACT1$("!PADREMP")))
    If NLEG% < 1 Then Exit Sub
    '
    NC% = NLEG%
    Text1.TEXT = TRIM$(str$(NC%))
    Call CARDATEMP(NLEG%) 'CARGA DATOS DEL EMPLEADO
    Call CARFAMCAR(NLEG%) 'CARGA FAMILIARES A CARGO
    Call CARCONCEOP(NLEG%) 'CARGA DE CONCEPTOS OPCIONALES
    Call CARNOTAEMP 'CARGA ANOTACIONES DEL EMPLEADO
    MODIFIC% = 0
   
    '
  End If
  '
End Sub

 Sub Command21_Click()
  '
  Command21.Enabled = False
  'VDEF$ TIENE QUE DAR EL PROX. NUM DE LEGAJO
  VDEF$ = Space(32)
  For JJ& = 1 To ULTREG("PADREMP")
  XX$ = REGLEIDO$("PADREMP", JJ&)
    NLEG% = CVI(Mid$(XX$, 1, 2))
    If NLEG% > PRONU% Then
      PRONU% = NLEG%
    End If
  Next JJ&
  Call REPLA(VDEF$, MKI$(PRONU% + 1), 1, 2)
9 obj$ = OBJPLA$("ALTEMP", VDEF$)
  If obj$ = "" Then GoTo 99
  VDEF$ = obj$
  '
  NULEG% = CVI(Mid$(obj$, 1, 2))
  DESCRI$ = TRIM$(Mid$(obj$, 3, 30))
  NUGRU% = CVI(Mid$(obj$, 33, 2))
  '//////////////VALIDACIONES////////////
  If Val(NULEG%) <= 0 Then
    Call COMUNI("Número de Legajo no válido.")
    GoTo 9
  End If
  '
  If TRIM(DESCRI$) = "" Then
    Call COMUNI("Apellido y Nombre no válido.")
    GoTo 9
  End If
  '
  For J& = 1 To ULTREG("PADREMP")
    XX$ = REGLEIDO$("PADREMP", J&)
    LEG% = CVI(Mid$(XX$, 1, 2))
    NOAP$ = TRIM(Mid$(XX$, 3, 30))
    If NULEG% = LEG% Then
      Call COMUNI("Número de Legajo no válido o Pre-Existente.")
      GoTo 9
    End If
    If NOAP$ = DESCRI$ Then
      Call COMUNI("Empleado Pre-Existente.")
      GoTo 9
    End If
  Next J&
  '
  If NUGRU% <= 0 Then
    Call COMUNI("Grupo de Empleado no Válido.")
    GoTo 9
  End If
  'LIMPIO LOS TEXT
  Text1 = ""
  'GRABO EL REGISTRO
  X$ = REGBLAN("PADREMP")
  ULTRE& = ULTREG&("PADREMP")
  Call REPLA(X$, MKI$(NULEG%), 1, 2)
  Call REPLA(X$, DESCRI$, 3, 30)
  Call REPLA(X$, Chr$(NUGRU%), 49, 1)
  Call REPLA(X$, Chr$(1), 48, 1) 'ESTADO ACTIVO
  Call GRAREG("PADREMP", X$, ULTRE& + 1)
  Call SETULTREG("PADREMP", ULTRE& + 1)
  '\\\OT-05-0577-WAR-25/07/06///
  YZ$ = REGBLAN("DATEMDGI")
  ULTRE1& = ULTREG&("DATEMDGI")
  Call REPLA(YZ$, MKI$(NULEG%), 1, 2)
  Call GRAREG("DATEMDGI", YZ$, ULTRE1& + 1)
  Call SETULTREG("DATEMDGI", ULTRE1& + 1)
  '\\\OT-05-0577-WAR-FIN///
  'GRABO EN ARCHIVOS AUXILIARES
  PADA1$ = Space(64)
  Call REPLA(PADA1$, X$, 1, 64)
  Call GRAREG("PADRON", PADA1$, ULTRE& + 1)
  PADA2$ = Mid$(X$, 65, 96)
  Call GRAREG("PADAUX", PADA2$, ULTRE& + 1)
  PADA3$ = Mid$(X$, 161, 48)
  Call GRAREG("PADAUX2", PADA3$, ULTRE& + 1)
  '
  RENUE$ = String$(128, 0)
  Call GRAREG("ACTIVA", RENUE$, ULTRE& + 1) 'RENGLON SEGUN EMPLEADO
  '
  Text1 = NULEG%
  '
99 Command21.Enabled = True
End Sub
Private Sub Command22_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("SELFECHA")
     VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
     If VDEVU$ <> "" Then
        Text5.TEXT = VALACT1$("SELFECHA")
     End If
  End If
End Sub

Private Sub Command222_Click()
  '
  If TRIM$(Label34(2).Caption) <> "" Then
      '
      NLEG% = Val(Text1.TEXT)
      '
      Screen.MousePointer = 11
      RNC$ = RECFILT$("INDIGRAF", 1, MKI$(NLEG%))
      REPLANO& = 0
      If Len(RNC$) >= 4 Then REPLANO& = CVS(RNC$)
      '
      If REPLANO& < 1 Or REPLANO > ULTREG&("INDIGRAF") Then
         REPLANO& = 1 + ULTREG&("INDIGRAF")
      End If
      RPLA$ = REGBLAN$("INDIGRAF")
      SPACIO1$ = Space$(999)
      Call REPLA(RPLA$, MKI$(NLEG%), 1, 2)
      Call REPLA(RPLA$, SPACIO1$, 3, 999)
      Call GRAREG("INDIGRAF", RPLA$, REPLANO&)
      Call CIERRARCH("INDIGRAF")
      '
      Label34(2).Caption = ""
      Screen.MousePointer = 1
      '
  End If
  '
End Sub

Private Sub Command23_Click()
   '
   If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
   If TRIM(Text1) <> "" Then
     Call SELECHO("SEXTAB")
     NSEX$ = TRIM$(VALACT1$("SEXTAB"))
     If TRIM$(NSEX$) = "" Then Exit Sub
     Text26.TEXT = NSEX$
   End If
   '
End Sub

Private Sub Command24_Click()
  '
  If Val(Text1) > 0 Then
   If Val(Text19) = 1 Then
    If COMALTER%("¿Esta Seguro de Dar de Baja el Empleado: '" + TRIM$(Text3.TEXT) + "'?\\Dar de Baja\Cancelar") = 1 Then
     Screen.MousePointer = 11
     MODIFIC% = 0
     Call BAJEMP(Val(Text1)) 'BAJA DE EMPLEADO
     Screen.MousePointer = 1
     Exit Sub
    End If
   Else
    If COMALTER%("¿Esta Seguro de Restaurar el Empleado: '" + TRIM$(Text3.TEXT) + "'?\\Restaurar\Cancelar") = 1 Then
     Screen.MousePointer = 11
     MODIFIC% = 0
     Call ACTEMP(Val(Text1)) 'ALTA DE EMPLEADO
     Screen.MousePointer = 1
     Exit Sub
    End If
   End If
  End If
  '
End Sub

Private Sub Command25_Click()
  '
  Command25.Enabled = False
  EMP_GES05.PRINTFORM
  Command25.Enabled = True
  '
End Sub

Private Sub Command26_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("SECPROD")
     VDEVU$ = TRIM$(VALACT1$("SECPROD"))
     If VDEVU$ <> "" Then
        Text17.TEXT = VALACT1$("SECPROD")
     End If
  End If
End Sub


Private Sub Command3_Click()
  '
  If Val(Text1) > 32767 Then
    GoTo 10
  End If
  If TRIM(RECFILT("PADREMP", 1, MKI$(Val(Text1)))) <> "" Then
    If MODIFIC% = 1 Then
      If COMALTER%("Se han Modificado Datos del Empleado\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") = 1 Then
        Screen.MousePointer = 11
        MODIFIC% = 0
       Call GRAEMP(Val(Text1))
       Screen.MousePointer = 1
      End If
    End If
    If MODIFIC% = 1 Then
      Call GRANOTEMP 'GRABA ANOTACIONES DE EMPLEADOS
    End If
  End If
  '
10 Call CIERRARCH("*.*")
   Unload Me
  '
End Sub

Private Sub Command34_Click()
  '
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("ESTCIV")
     VDEVU$ = TRIM$(VALACT1$("ESTCIV"))
     If VDEVU$ <> "" Then
        Text2.TEXT = VALACT1$("ESTCIV")
     End If
  End If
  '
End Sub

Private Sub Command4_Click()
  '
  Call HELPONLINE("HELPEMP")
  '
End Sub

Private Sub Command53_Click(Index As Integer)
   '
   NLEG% = Val(TRIM$(Text1.TEXT))
   If NLEG% < 1 Then Exit Sub
   If TRIM(RECFILT$("PADREMP", 1, MKI$(NLEG%))) = "" Then Exit Sub
   '
   Command53(Index).Enabled = False
   '
10 CommonDialog1.InitDir = ULODATOS$ & "\"
   CommonDialog1.DialogTitle = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Archivo de Imagen(JPG; BMP; WMF)|*.JPG; *.GIF; *.BMP; *.WMF|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   '
   RUELE$ = Left$(CommonDialog1.FileName, 2)
   If RUELE$ <> ULODATOS$ Then
      Call MENSERR(24, "Ruta de Archivo no Válida.\ Los Archivos Gráficos Deben \ Alojarse en el Disco de Trabajo del Sistema.")
      CommonDialog1.InitDir = ULODATOS$ & "\"
      GoTo 10
   End If
   '
   Label34(Index) = UCase$(CommonDialog1.FileName)
   '
   On Error GoTo 0
   '
99 Screen.MousePointer = 1
   Command53(Index).Enabled = True
   On Error GoTo 0
   '
End Sub

 Sub Command9_Click()
   '
   If MODIFIC% = 1 And Val(Text1) > 0 And Val(Text1) < 32767 Then
    If COMALTER%("Se han Modificado Datos del Empleado\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") = 1 Then
      Screen.MousePointer = 11
      MODIFIC% = 0
      Call GRAEMP(Val(Text1))
      Screen.MousePointer = 1
    End If
   End If
   '
   If MODIFIC% = 1 Then
    Call GRANOTEMP 'GRABA ANOTACIONES DE EMPLEADOS
   End If
   '
   Call SELECHO("PADREMP")
   NLEG% = Val(TRIM$(VALACT1$("PADREMP")))
   If NLEG% < 1 Then Exit Sub
   '
   NC% = NLEG%
   Text1.TEXT = TRIM$(str$(NC%))
   Call CARDATEMP(NLEG%) 'CARGA DATOS DEL EMPLEADO
   Call CARFAMCAR(NLEG%) 'CARGA FAMILIARES A CARGO
   Call CARCONCEOP(NLEG%) 'CARGA DE CONCEPTOS OPCIONALES
   Call CARNOTAEMP 'CARGA ANOTACIONES DEL EMPLEADO
   MODIFIC% = 0
 
   '
End Sub
Sub CARFAMCAR(LEGA%) 'CARGA DE FAMILIARES X Nº LEG.
  '
  List3.Clear
  '
  Call TRALOFIL("FAMCAR", 1, MKI$(LEGA%))
  For J& = 1 To ULTREG&("!FAMCAR")
      XX$ = REGLEIDO$("!FAMCAR", J&)
      NOAP$ = Mid$(XX$, 3, 25)
      FENA = CVI(Mid$(XX$, 28, 2))
      PAREN$ = Mid$(XX$, 30, 10)
      DOCUFA$ = Mid$(XX$, 40, 11)
      YY$ = Space(68)
      Call REPLA(YY$, NOAP$, 1, 27)
      Call REPLA(YY$, FECHATEX$(FENA), 30, 10)
      Call REPLA(YY$, PAREN$, 41, 10)
      Call REPLA(YY$, DOCUFA$, 53, 11)
      List3.AddItem YY$
  Next J&
  '
End Sub
Sub CARDATEMP(LEGA%) 'CARGA DE EMPLEADOS
  '
  XX1$ = RECFILT$("PADREMP", 1, MKI$(LEGA%))
  J1& = CVS(XX1$)
  '
  XX$ = REGLEIDO$("PADREMP", J1&)
  NOAP$ = TRIM$(Mid$(XX$, 3, 30))
  TIDO$ = TRIM$(Mid$(XX$, 33, 13))
  FEIN = CVI(Mid$(XX$, 46, 2))
  STA = Asc(Mid$(XX$, 48, 1))
  GRU = Asc(Mid$(XX$, 49, 1))
  SECP1 = CVS(Mid$(XX$, 50, 4)) 'SECTOR
  CATE$ = TRIM$(Mid$(XX$, 54, 9))
  FENAC = CVI(Mid$(XX$, 63, 2))
    DOMI$ = TRIM$(Mid$(XX$, 65, 30)) 'DOMICILIO
  LOCA$ = TRIM$(Mid$(XX$, 95, 30)) 'LOCALIDAD
  NACIO$ = TRIM$(Mid$(XX$, 125, 10)) 'NACIONALIDAD
  ECIV$ = TRIM$(Mid$(XX$, 135, 10)) 'EST.CIVIL
  FANTIG = CVI(Mid$(XX$, 145, 2)) 'F.ANTIGUEDAD
  FEGRE = CVI(Mid$(XX$, 147, 2)) 'F.EGRESO
  SEX$ = TRIM$(Mid$(XX$, 149, 1)) 'SEXO
    TARE$ = TRIM$(Mid$(XX$, 161, 20)) 'TAREA
  CAJU$ = TRIM$(Mid$(XX$, 181, 16)) 'CAJA DE JUBILACION
  NUAFI$ = TRIM$(Mid$(XX$, 197, 12)) 'NUMERO DE AFILIADO
  CUIL1$ = TRIM$(Mid$(XX$, 209, 13)) 'NUMERO DE CUIL
  CPOSTA$ = TRIM$(Mid$(XX$, 222, 10)) ' CODIGO POSTAL
  HORTRA$ = TRIM$(Mid$(XX$, 232, 15))
  NAFJP$ = TRIM$(Mid$(XX$, 247, 15))
  TICONTRA$ = TRIM$(Mid$(FILTERGETRECORD$("TICONTRA", 1, MKI$(LEGA%)), 3))
  '
  SECP2$ = FORMATNUM$(SECP1, "F7.3")
  Text17 = SECP2$
  Text15 = CUIL1$
  Text3 = NOAP$
  Text25 = TIDO$
  Text9 = FECHATEX(FEIN)
  Text5 = FECHATEX(FANTIG)
  Text6 = CStr(GRU)
  Text10 = FECHATEX(FENAC)
  Text16 = CATE$
  Text19 = CStr(STA)
    Text4 = DOMI$
  Text7 = LOCA$
  Text36 = CPOSTA$
  Text37 = HORTRA$
  Text29 = NACIO$
  Text2 = ECIV$
  Text12 = FECHATEX(FEGRE)
  Text26 = SEX$
    Text13 = CAJU$
  Text14 = NUAFI$
  Text141 = TICONTRA$
  Text142 = NAFJP$
  Text8 = TARE$
  LEGACT% = LEGA%
  'CARGO IMAGEN DEL EMPLEADO
  Screen.MousePointer = 11
  RNC$ = RECFILT$("INDIGRAF", 1, MKI$(LEGA%))
  REPLANO& = 0
  If Len(RNC$) >= 4 Then REPLANO& = CVS(RNC$)
  X$ = REGLEIDO$("INDIGRAF", REPLANO&)
  Label34(2).Caption = UCase$(TRIM$(Mid$(X$, 3)))
  Call CIERRARCH("INDIGRAF")
  '
  '\\\OT-05-0577-WAR-25/07/06///
  For J2& = 1 + ULTREG&("DATEMDGI") To J1&
     XY1$ = REGLEIDO$("PADREMP", J2&)
     LEGA2% = CVI(Left$(XX1$, 2))
     XY2$ = REGBLAN$("PADREMP")
     Call REPLA(XY2$, MKI$(LEGA2%), 1, 2)
     Call GRAREG("DATEMDGI", XY2$, J2&)
  Next J2&
  'XX1$ = RECFILT$("DATEMDGI", 1, MKI$(LEGA%))
  'J1& = CVS(XX1$)
  '
  XY$ = REGLEIDO$("DATEMDGI", J1&)
  CAHI% = Asc(Mid$(XY$, 20, 1)) 'GRUPO FLIAR.-CANTIDAD DE HIJOS
    Text27.TEXT = Chr$(CAHI%)
  CAAD% = Asc(Mid$(XY$, 21, 1)) 'ADHERENTES
    Text31.TEXT = Chr$(CAAD%)
  COOS = CVS(Mid$(XY$, 22, 4)) 'CODIGO DE O.S.
    Text30.TEXT = TRIM$(FORMATNUM$(COOS, "F6.0"))
  APAD = CVS(Mid$(XY$, 26, 4)) 'APORTE ADICIONAL EN PORCENTAJE
    Text23.TEXT = TRIM$(FORMATNUM$(APAD, "F6.0"))
  COSI$ = Mid$(XY$, 30, 2) 'CODIGO DE SITUACIÓN
    Text22.TEXT = COSI$
  COCI$ = Mid$(XY$, 32, 2) 'CODIGO DE CONDICIÓN
    Text18.TEXT = COCI$
  COCON$ = Mid$(XY$, 34, 3) 'CODIGO DE MODALIDAD DE CONTRATACION
    Text24.TEXT = COCON$
  COSIN$ = Mid$(XY$, 37, 2) 'CODIGO DEL SINIESTRADO
    Text21.TEXT = COSIN$
  MARCO$ = Mid$(XY$, 39, 1) 'CODIGO DE CORRESPONDE REDUCCIÓN
    Text32.TEXT = MARCO$
  CAPRECO$ = Mid$(XY$, 40, 9) 'CAPITAL DE RECOMPOSICIÓN LRT
    Text33.TEXT = CAPRECO$
  TRACON$ = Mid$(XY$, 49, 1) 'TRABAJADOR CONVENCIONADO
    Text34.TEXT = TRACON$
  TIPOP$ = Mid$(XY$, 50, 1) 'TIPO DE OPERACION
    Text35.TEXT = TIPOP$
  CONYU$ = Mid$(XY$, 51, 1) 'CONYUGE
    Text20.TEXT = CONYU$
  '
  '\\\OT-05-0577-WAR-FIN///
  Screen.MousePointer = 1
  '
End Sub

Sub GRAEMP(NLEGA%) 'GRABA EMPLEADO
  '
  Call COPYSTRU("DEUDOR1", "TICONTRA")
  '
  XX1$ = RECFILT$("PADREMP", 1, MKI$(NLEGA%))
  J1& = CVS(XX1$)
  '
  ' ESTO ES PARA COMPLETAR LOS DATOS DE LAS TABLAS AUXILIARES
  For J2& = 1 + ULTREG&("PADRON") To J1& - 1
     XX$ = REGLEIDO$("PADREMP", J2&)
     PADA1$ = Mid$(XX$, 1, 64)
     Call GRAREG("PADRON", PADA1$, J2&)
  Next J2&
  '
  For J2& = 1 + ULTREG&("PADAUX") To J1& - 1
     XX$ = REGLEIDO$("PADREMP", J2&)
     PADA2$ = Mid$(XX$, 65, 96)
     Call GRAREG("PADAUX", PADA2$, J2&)
  Next J2&
  '
  For J2& = 1 + ULTREG&("PADAUX2") To J1& - 1
     XX$ = REGLEIDO$("PADREMP", J2&)
     PADA3$ = Mid$(XX$, 161, 48)
     Call GRAREG("PADAUX2", PADA3$, J2&)
  Next J2&
  '
  XX$ = REGLEIDO$("PADREMP", J1&)
  If XX$ <> "" Then
    '
    NLEG% = LEGACT% 'Val(Text1)
    NOAP$ = Text3
    TIDO$ = Text25
    'validación del numero de documento
    FEIN% = FECHANUM(TRIM(Text9))
    FANTIG% = FECHANUM(TRIM(Text5))
      If FANTIG% = 0 Then FANTIG% = FEIN%
      If FANTIG% > FEIN% Then FANTIG% = FEIN%
    GRU% = Val(Text6)
    If TRIM$(ECOARCH$("DATGREM1", MKI$(GRU%))) = "" Then
      Call MENSERR(24, "Grupo de Empleados no Válido.")
      Exit Sub
    End If
    CATE$ = TRIM(Text16)
    If TRIM$(ECOARCH$("CATEGO", CATE$)) = "" Then
      Call MENSERR(24, "Categoria no Válida.")
      Exit Sub
    End If
    FENA% = FECHANUM(TRIM(Text10))
    DOMI$ = TRIM$(Text4)
    LOCA$ = TRIM$(Text7)
    NACIO$ = TRIM$(Text29)
    ECIV$ = TRIM$(Text2)
    If TRIM$(ECOARCH("ESTCIV", MKI$(Val(ECIV$)))) = "" Then
      Call MENSERR(24, "Estado Civil no Válido.")
      Exit Sub
    End If
    FEGRE% = FECHANUM(TRIM(Text12))
    SEX$ = TRIM$(Text26)
    If TRIM$(ECOARCH("SEXTAB", SEX$)) = "" Then
      Call MENSERR(24, "Sexo no Válido.")
      Exit Sub
    End If
    '
    SECP1 = XVALO(Text17.TEXT)
    If TRIM$(ECOARCH$("SECPROD", MKS$(SECP1))) = "" Then
      Call MENSERR(24, "Sector no Válido.")
      Exit Sub
    End If
    '
    TARE$ = TRIM$(Text8)
    CAJUB$ = TRIM$(Text13)
    NUAFI$ = TRIM$(Text14)
    CUIL1$ = TRIM$(Text15)
    CPOSTA$ = Text36
    HORTRA$ = Text37
    TCONTRA$ = Text141
    NAFJP$ = Text142
    '
    Call REPLA(XX$, MKI$(NLEG%), 1, 2)
    Call REPLA(XX$, NOAP$, 3, 30)
    Call REPLA(XX$, TIDO$, 33, 13)
    Call REPLA(XX$, MKI$(FEIN%), 46, 2)
    Call REPLA(XX$, Chr$(1), 48, 1) 'ESTADO ACTIVO
    Call REPLA(XX$, Chr$(GRU%), 49, 1)
    Call REPLA(XX$, MKS$(SECP1), 50, 4) 'SECTOR
    Call REPLA(XX$, CATE$, 54, 9)
    Call REPLA(XX$, MKI$(FENA%), 63, 2)
    Call REPLA(XX$, DOMI$, 65, 30)
    Call REPLA(XX$, LOCA$, 95, 30)
    Call REPLA(XX$, NACIO$, 125, 10)
      ECOECIV$ = TRIM(Label22)
    Call REPLA(XX$, ECIV$, 135, 10)
    Call REPLA(XX$, MKI$(FANTIG%), 145, 2)
    Call REPLA(XX$, MKI$(FEGRE%), 147, 2)
    Call REPLA(XX$, SEX$, 149, 1)
    Call REPLA(XX$, TARE$, 161, 20)
    Call REPLA(XX$, CAJUB$, 181, 16)
    Call REPLA(XX$, NUAFI$, 197, 12)
    Call REPLA(XX$, CUIL1$, 209, 13)
    Call REPLA(XX$, CPOSTA$, 222, 10)
    Call REPLA(XX$, HORTRA$, 232, 15)
    Call REPLA(XX$, NAFJP$, 247, 10)
     'GRABO EN PADREMP
    Call GRAREG("PADREMP", XX$, J1&)
    'GRABO EN ARCHIVOS AUXILIARES
    PADA1$ = Mid$(XX$, 1, 64)
    Call GRAREG("PADRON", PADA1$, J1&)
    PADA2$ = Mid$(XX$, 65, 96)
    Call REPLA(PADA2$, ECOECIV$, 71, 10)
    Call GRAREG("PADAUX", PADA2$, J1&)
    PADA3$ = Mid$(XX$, 161, 48)
    Call GRAREG("PADAUX2", PADA3$, J1&)
    'GRABO TIPO DE CONTRATO
    Call FILTERPUTRECORD("TICONTRA", 1, MKI$(NLEG%), MKI$(NLEG%) + AJUSTI$(TCONTRA$, 30))
    'GRABO SI TIENE IMAGEN
    If TRIM$(Label34(2).Caption) <> "" Then
        '
        NLEG% = LEGACT% 'Val(Text1.TEXT)
        '
        Screen.MousePointer = 11
        RNC$ = RECFILT$("INDIGRAF", 1, MKI$(NLEG%))
        REPLANO& = 0
        If Len(RNC$) >= 4 Then REPLANO& = CVS(RNC$)
        '
        If REPLANO& < 1 Or REPLANO > ULTREG&("INDIGRAF") Then
           REPLANO& = 1 + ULTREG&("INDIGRAF")
        End If
        RPLA$ = REGBLAN$("INDIGRAF")
        Call REPLA(RPLA$, MKI$(NLEG%), 1, 2)
        Call REPLA(RPLA$, TRIM$(Label34(2).Caption), 3, 999)
        Call GRAREG("INDIGRAF", RPLA$, REPLANO&)
        Call CIERRARCH("INDIGRAF")
        '
        Screen.MousePointer = 1
        '
    End If
    Call FRESHECHO("PADREMP")
    '\\\OT-05-0577-WAR-25/07/06///
    '
    For J2& = 1 + ULTREG&("DATEMDGI") To J1&
       XY1$ = REGLEIDO$("PADREMP", J2&)
       LEGA2% = CVI(Left$(XX1$, 2))
       XY2$ = REGBLAN$("PADREMP")
       Call REPLA(XY2$, MKI$(LEGA2%), 1, 2)
       Call GRAREG("DATEMDGI", XY2$, J2&)
    Next J2&
    '
    XY$ = REGLEIDO$("DATEMDGI", J1&)
    CAHI% = Asc(XVALO(Text27.TEXT)) 'GRUPO FLIAR.-CANTIDAD DE HIJOS
      Call REPLA(XY$, Chr$(CAHI%), 20, 1)
    CAAD% = Asc(XVALO(Text31.TEXT)) 'ADHERENTES
      Call REPLA(XY$, Chr$(CAAD%), 21, 1)
    COOS = XVALO(Text30.TEXT) 'CODIGO DE O.S.
      Call REPLA(XY$, MKS$(COOS), 22, 4)
    APAD = XVALO(Text23.TEXT) 'APORTE ADICIONAL EN PORCENTAJE
      Call REPLA(XY$, MKS(APAD), 26, 4)
    COSI$ = Text22.TEXT 'CODIGO DE SITUACIÓN
      Call REPLA(XY$, COSI$, 30, 2)
    COCI$ = Text18.TEXT 'CODIGO DE CONDICIÓN
      Call REPLA(XY$, COCI$, 32, 2)
    COCON$ = Text24.TEXT 'CODIGO DE MODALIDAD DE CONTRATACION
      Call REPLA(XY$, COCON$, 34, 3)
    COSIN$ = Text21.TEXT 'CODIGO DEL SINIESTRADO
      Call REPLA(XY$, COSIN$, 37, 2)
    MARCO$ = Text32.TEXT 'CODIGO DE CORRESPONDE REDUCCIÓN
      Call REPLA(XY$, MARCO$, 39, 1)
    CAPRECO$ = Text33.TEXT 'CAPITAL DE RECOMPOSICIÓN LRT
      Call REPLA(XY$, CAPRECO$, 40, 9)
    TRACON$ = Text34.TEXT 'TRABAJADOR CONVENCIONADO
      Call REPLA(XY$, TRACON$, 49, 1)
    TIPOP$ = Text35.TEXT 'TIPO DE OPERACION
      Call REPLA(XY$, TIPOP$, 50, 1)
    CONYU$ = Text20.TEXT 'CONYUGE
      Call REPLA(XY$, CONYU$, 51, 1)
    Call GRAREG("DATEMDGI", XY$, J1&)
    '
    '\\\OT-05-0577-WAR-FIN///
  End If
  '
  Call ActualizaTablaDatosAuxiliaresDeOperarios
  '
End Sub

Sub ACTEMP(NLEGA%) 'ESTADO ACTIVO AL  EMPLEADO
  '
  XX1$ = RECFILT$("PADREMP", 1, MKI$(NLEGA%))
  J1& = CVS(XX1$)
  '
  XX$ = REGLEIDO$("PADREMP", J1&)
  If XX$ <> "" Then
    '
    Call REPLA(XX$, Chr$(1), 48, 1) 'ESTADO ACTIVO
    '
    Call GRAREG("PADREMP", XX$, J1&)
    Text19 = 1
    'GRABO ESTADO EN ARCHIVO AUXILIAR
    XX1$ = REGLEIDO$("PADRON", J1&)
    Call REPLA(XX1$, Chr$(1), 48, 1) 'ESTADO ACTIVO
    Call GRAREG("PADRON", XX1$, J1&)
    '
  End If
  '
End Sub
Sub BAJEMP(NLEGA%) 'BAJA DE EMPLEADO
  '
  XX1$ = RECFILT$("PADREMP", 1, MKI$(NLEGA%))
  J1& = CVS(XX1$)
  '
  XX$ = REGLEIDO$("PADREMP", J1&)
  If XX$ <> "" Then
    '
    Call REPLA(XX$, Chr$(0), 48, 1) 'ESTADO BAJA
    '
    Call GRAREG("PADREMP", XX$, J1&)
    Text19 = 0
    'GRABO ESTADO EN ARCHIVO AUXILIAR
    XX1$ = REGLEIDO$("PADRON", J1&)
    Call REPLA(XX1$, Chr$(0), 48, 1) 'ESTADO BAJA
    Call GRAREG("PADRON", XX1$, J1&)
    '
  End If
  '
End Sub

Private Sub ConGenEmp_Click()
  '
  Call Command14_Click
  '
End Sub

Private Sub DeHaCaEmp_Click()
  '
  Call Command2_Click
  '
End Sub

Private Sub EditEmp_Click()
  '
  Call Command9_Click
  '
End Sub

Private Sub Exit_Click()
  '
  Call Command3_Click
  '
End Sub

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
  'Call APLICACIONSKINS(Me, Picture1)

  Call BLANFORMU
  '
  Call CREATAB
  Call APACTPADRO 'APAREA PADRON DE EMPLEADOS CON ACTIVA.DAT
  Call APAPADREMP 'APAREA PADREMP.DAT CON PADRON.DAT, PADAUX.DAT, PADAUX2.DAT.
  Call APAREMBA   'APAREA PADRON DE EMPLEADOS CON REMBAS.DAT
  Call APAREDGI   'APAREA ARHIVO DE TRANSF. DE DGI
  '
  TOPEINI(2) = Picture1(2).Top
  IZQUINI(2) = Picture1(2).Left
  ANCHINI(2) = Picture1(2).Width
  ALTUINI(2) = Picture1(2).Height
  '
  Call APLICACIONSKINS(Me)
  Call COPYSTRU("DEUDOR1", "TICONTRA")
  '
End Sub
'DESDE ACA COPIADO - ¿PARA QUE SIRVE?
Private Sub Form_Activate()
   WindowState = 0
   Openforms = DoEvents
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
     Openforms = DoEvents
   Wend
End Sub

'HASTA ACA
'Private Sub Form_Unload(Cancel As Integer)
'  '
'  If TRIM(RECFILT("PADREMP", 1, MKI$(Val(Text1)))) <> "" Then
'    If MODIFIC% = 1 Then
'      If COMALTER%("Se han Modificado Datos del Empleado\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") = 1 Then
'        Screen.MousePointer = 11
'        MODIFIC% = 0
'       Call GRAEMP(Val(Text1))
'       Screen.MousePointer = 1
'      End If
'    End If
'    If MODIFICNOT% = 1 Then
'      Call GRANOTEMP 'GRABA ANOTACIONES DE EMPLEADOS
'    End If
'  End If
'  '
'  Call CIERRARCH("*.*")
'  '
'End Sub
Private Sub GuaEmp_Click()
  '
  Call Command1_Click
  '
End Sub

Private Sub ImpoEmpleados_Click()
    '
    Set objexcel = New Excel.Application
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    '
    DoEvents
    Screen.MousePointer = 11
    '
    Ruta$ = UCase$(CommonDialog1.FileName)
    Set Libro = objexcel.Workbooks.Open(Ruta$)
    Set HOJA = Libro.Sheets(1)
    '
    URECORD& = UltimoRegistro(1)
    Call CIERRARCH("PADREMP")
    For Fila& = 2 To URECORD&
        Call TRACE(24, Fila&, URECORD&)
        
        X$ = REGBLAN("PADREMP")
        ULTRE& = ULTREG&("PADREMP")
            NULEG% = XVALO(HOJA.Cells(Fila&, 1))
        Call REPLA(X$, MKI$(NULEG%), 1, 2)
            DESCRI$ = SAFETEXT$(HOJA.Cells(Fila&, 2))
        Call REPLA(X$, DESCRI$, 3, 30)
            NUGRU% = XVALO(HOJA.Cells(Fila&, 6))
        Call REPLA(X$, Chr$(NUGRU%), 49, 1)
        Call REPLA(X$, Chr$(1), 48, 1) 'ESTADO ACTIVO
        Call GRAREG("PADREMP", X$, ULTRE& + 1)
        Call SETULTREG("PADREMP", ULTRE& + 1)
        '\\\OT-05-0577-WAR-25/07/06///
        YZ$ = REGBLAN("DATEMDGI")
        ULTRE1& = ULTREG&("DATEMDGI")
        Call REPLA(YZ$, MKI$(NULEG%), 1, 2)
        Call GRAREG("DATEMDGI", YZ$, ULTRE1& + 1)
        Call SETULTREG("DATEMDGI", ULTRE1& + 1)
        '\\\OT-05-0577-WAR-FIN///
        'GRABO EN ARCHIVOS AUXILIARES
        PADA1$ = Space(64)
        Call REPLA(PADA1$, X$, 1, 64)
        Call GRAREG("PADRON", PADA1$, ULTRE& + 1)
        PADA2$ = Mid$(X$, 65, 96)
        Call GRAREG("PADAUX", PADA2$, ULTRE& + 1)
        PADA3$ = Mid$(X$, 161, 48)
        Call GRAREG("PADAUX2", PADA3$, ULTRE& + 1)
        '
        'RENUE$ = String$(128, 0)
        'Call GRAREG("ACTIVA", RENUE$, ULTRE& + 1) 'RENGLON SEGUN EMPLEADO
        '
        Text1.TEXT = CStr(NULEG%)
        '
        SECP2$ = FORMATNUM$(XVALO(HOJA.Cells(Fila&, 7)), "F7.3")
        Text17 = SECP2$
        Text15 = CStr(HOJA.Cells(Fila&, 20))
        'Text3 = TRIM$(Hoja.Cells(FILA&, 1))
        Text25 = TRIM$(HOJA.Cells(Fila&, 3))
            FEIN = CVINT(CDate(HOJA.Cells(Fila&, 4)))
        Text9 = FECHATEX(FEIN)
            FANTIG = CVINT(CDate(HOJA.Cells(Fila&, 15)))
        Text5 = FECHATEX(FANTIG)
        Text6 = TRIM$(HOJA.Cells(Fila&, 6))
            FNAC = CVINT(CDate(HOJA.Cells(Fila&, 9)))
        Text10 = FECHATEX$(FNAC)
            STAT$ = CStr(HOJA.Cells(Fila&, 5))
            If STAT$ = "A" Then
                Text19 = "1"
            Else
                Text19 = "0"
            End If
        Text16 = CStr(HOJA.Cells(Fila&, 8))
        Text4 = TRIM$(HOJA.Cells(Fila&, 10))
        Text7 = TRIM$(HOJA.Cells(Fila&, 12))
        Text36 = TRIM$(HOJA.Cells(Fila&, 11))
        Text37 = TRIM$(HOJA.Cells(Fila&, 21))
        Text29 = TRIM$(HOJA.Cells(Fila&, 13))
        Text2 = TRIM$(HOJA.Cells(Fila&, 14))
            FEGRE = CVINT(CDate(HOJA.Cells(Fila&, 16)))
        Text12 = FECHATEX(FEGRE)
        Text26 = TRIM$(HOJA.Cells(Fila&, 17))
          Text13 = ""
        Text14 = ""
        Text141 = ""
        Text142 = ""
        Text8 = TRIM$(HOJA.Cells(Fila&, 18))
        '
        MODIFIC% = 0
        Call IMPORTAEMPLEADOS(Val(Text1))
   Next Fila&
   '
   Libro.Close
   objexcel.QUIT
   Set Libro = Nothing
   Set objexcel = Nothing
   '
   Screen.MousePointer = 1
   Call MENSERR(24, "Proceso de Importación Finalizada")
   
99 Exit Sub
End Sub

Sub IMPORTAEMPLEADOS(NLEGA%) 'GRABA EMPLEADO
  '
  Call COPYSTRU("DEUDOR1", "TICONTRA")
  '
  XX1$ = RECFILT$("PADREMP", 1, MKI$(NLEGA%))
  J1& = CVS(XX1$)
  '
  ' ESTO ES PARA COMPLETAR LOS DATOS DE LAS TABLAS AUXILIARES
  For J2& = 1 + ULTREG&("PADRON") To J1& - 1
     XX$ = REGLEIDO$("PADREMP", J2&)
     PADA1$ = Mid$(XX$, 1, 64)
     Call GRAREG("PADRON", PADA1$, J2&)
  Next J2&
  '
  For J2& = 1 + ULTREG&("PADAUX") To J1& - 1
     XX$ = REGLEIDO$("PADREMP", J2&)
     PADA2$ = Mid$(XX$, 65, 96)
     Call GRAREG("PADAUX", PADA2$, J2&)
  Next J2&
  '
  For J2& = 1 + ULTREG&("PADAUX2") To J1& - 1
     XX$ = REGLEIDO$("PADREMP", J2&)
     PADA3$ = Mid$(XX$, 161, 48)
     Call GRAREG("PADAUX2", PADA3$, J2&)
  Next J2&
  '
  XX$ = REGLEIDO$("PADREMP", J1&)
  If XX$ <> "" Then
    '
    NLEG% = LEGACT% 'Val(Text1)
    NOAP$ = Text3
    TIDO$ = Text25
    'validación del numero de documento
    FEIN% = FECHANUM(TRIM(Text9))
    FANTIG% = FECHANUM(TRIM(Text5))
      If FANTIG% = 0 Then FANTIG% = FEIN%
      If FANTIG% > FEIN% Then FANTIG% = FEIN%
    GRU% = Val(Text6)
    If TRIM$(ECOARCH$("DATGREM1", MKI$(GRU%))) = "" Then
      'Call MENSERR(24, "Grupo de Empleados no Válido.")
      'Exit Sub
      GRU% = 0
    End If
    CATE$ = TRIM(Text16)
    If TRIM$(ECOARCH$("CATEGO", CATE$)) = "" Then
      'Call MENSERR(24, "Categoria no Válida.")
      'Exit Sub
      CATE$ = ""
    End If
    FENA% = FECHANUM(TRIM(Text10))
    DOMI$ = TRIM$(Text4)
    LOCA$ = TRIM$(Text7)
    NACIO$ = TRIM$(Text29)
    ECIV$ = TRIM$(Text2)
    If TRIM$(ECOARCH("ESTCIV", MKI$(Val(ECIV$)))) = "" Then
      'Call MENSERR(24, "Estado Civil no Válido.")
      'Exit Sub
      ECIV$ = ""
    End If
    FEGRE% = FECHANUM(TRIM(Text12))
    SEX$ = TRIM$(Text26)
    If TRIM$(ECOARCH("SEXTAB", SEX$)) = "" Then
      'Call MENSERR(24, "Sexo no Válido.")
      'Exit Sub
      SEX$ = ""
    End If
    '
    SECP1 = XVALO(Text17.TEXT)
    If TRIM$(ECOARCH$("SECPROD", MKS$(SECP1))) = "" Then
      'Call MENSERR(24, "Sector no Válido.")
      'Exit Sub
      SECP1 = 0
    End If
    '
    TARE$ = TRIM$(Text8)
    CAJUB$ = TRIM$(Text13)
    NUAFI$ = TRIM$(Text14)
    CUIL1$ = TRIM$(Text15)
    CPOSTA$ = Text36
    HORTRA$ = Text37
    TCONTRA$ = Text141
    NAFJP$ = Text142
    '
    Call REPLA(XX$, MKI$(NLEG%), 1, 2)
    Call REPLA(XX$, NOAP$, 3, 30)
    Call REPLA(XX$, TIDO$, 33, 13)
    Call REPLA(XX$, MKI$(FEIN%), 46, 2)
    Call REPLA(XX$, Chr$(1), 48, 1) 'ESTADO ACTIVO
    Call REPLA(XX$, Chr$(GRU%), 49, 1)
    Call REPLA(XX$, MKS$(SECP1), 50, 4) 'SECTOR
    Call REPLA(XX$, CATE$, 54, 9)
    Call REPLA(XX$, MKI$(FENA%), 63, 2)
    Call REPLA(XX$, DOMI$, 65, 30)
    Call REPLA(XX$, LOCA$, 95, 30)
    Call REPLA(XX$, NACIO$, 125, 10)
      ECOECIV$ = TRIM(Label22)
    Call REPLA(XX$, ECIV$, 135, 10)
    Call REPLA(XX$, MKI$(FANTIG%), 145, 2)
    Call REPLA(XX$, MKI$(FEGRE%), 147, 2)
    Call REPLA(XX$, SEX$, 149, 1)
    Call REPLA(XX$, TARE$, 161, 20)
    Call REPLA(XX$, CAJUB$, 181, 16)
    Call REPLA(XX$, NUAFI$, 197, 12)
    Call REPLA(XX$, CUIL1$, 209, 13)
    Call REPLA(XX$, CPOSTA$, 222, 10)
    Call REPLA(XX$, HORTRA$, 232, 15)
    Call REPLA(XX$, NAFJP$, 247, 10)
     'GRABO EN PADREMP
    Call GRAREG("PADREMP", XX$, J1&)
    'GRABO EN ARCHIVOS AUXILIARES
    PADA1$ = Mid$(XX$, 1, 64)
    Call GRAREG("PADRON", PADA1$, J1&)
    PADA2$ = Mid$(XX$, 65, 96)
    Call REPLA(PADA2$, ECOECIV$, 71, 10)
    Call GRAREG("PADAUX", PADA2$, J1&)
    PADA3$ = Mid$(XX$, 161, 48)
    Call GRAREG("PADAUX2", PADA3$, J1&)
    'GRABO TIPO DE CONTRATO
    Call FILTERPUTRECORD("TICONTRA", 1, MKI$(NLEG%), MKI$(NLEG%) + AJUSTI$(TCONTRA$, 30))
    'GRABO SI TIENE IMAGEN
    If TRIM$(Label34(2).Caption) <> "" Then
        '
        NLEG% = LEGACT% 'Val(Text1.TEXT)
        '
        Screen.MousePointer = 11
        RNC$ = RECFILT$("INDIGRAF", 1, MKI$(NLEG%))
        REPLANO& = 0
        If Len(RNC$) >= 4 Then REPLANO& = CVS(RNC$)
        '
        If REPLANO& < 1 Or REPLANO > ULTREG&("INDIGRAF") Then
           REPLANO& = 1 + ULTREG&("INDIGRAF")
        End If
        RPLA$ = REGBLAN$("INDIGRAF")
        Call REPLA(RPLA$, MKI$(NLEG%), 1, 2)
        Call REPLA(RPLA$, TRIM$(Label34(2).Caption), 3, 999)
        Call GRAREG("INDIGRAF", RPLA$, REPLANO&)
        Call CIERRARCH("INDIGRAF")
        '
        Screen.MousePointer = 1
        '
    End If
    Call FRESHECHO("PADREMP")
    '\\\OT-05-0577-WAR-25/07/06///
    '
    For J2& = 1 + ULTREG&("DATEMDGI") To J1&
       XY1$ = REGLEIDO$("PADREMP", J2&)
       LEGA2% = CVI(Left$(XX1$, 2))
       XY2$ = REGBLAN$("PADREMP")
       Call REPLA(XY2$, MKI$(LEGA2%), 1, 2)
       Call GRAREG("DATEMDGI", XY2$, J2&)
    Next J2&
    '
    XY$ = REGLEIDO$("DATEMDGI", J1&)
    CAHI% = Asc(XVALO(Text27.TEXT)) 'GRUPO FLIAR.-CANTIDAD DE HIJOS
      Call REPLA(XY$, Chr$(CAHI%), 20, 1)
    CAAD% = Asc(XVALO(Text31.TEXT)) 'ADHERENTES
      Call REPLA(XY$, Chr$(CAAD%), 21, 1)
    COOS = XVALO(Text30.TEXT) 'CODIGO DE O.S.
      Call REPLA(XY$, MKS$(COOS), 22, 4)
    APAD = XVALO(Text23.TEXT) 'APORTE ADICIONAL EN PORCENTAJE
      Call REPLA(XY$, MKS(APAD), 26, 4)
    COSI$ = Text22.TEXT 'CODIGO DE SITUACIÓN
      Call REPLA(XY$, COSI$, 30, 2)
    COCI$ = Text18.TEXT 'CODIGO DE CONDICIÓN
      Call REPLA(XY$, COCI$, 32, 2)
    COCON$ = Text24.TEXT 'CODIGO DE MODALIDAD DE CONTRATACION
      Call REPLA(XY$, COCON$, 34, 3)
    COSIN$ = Text21.TEXT 'CODIGO DEL SINIESTRADO
      Call REPLA(XY$, COSIN$, 37, 2)
    MARCO$ = Text32.TEXT 'CODIGO DE CORRESPONDE REDUCCIÓN
      Call REPLA(XY$, MARCO$, 39, 1)
    CAPRECO$ = Text33.TEXT 'CAPITAL DE RECOMPOSICIÓN LRT
      Call REPLA(XY$, CAPRECO$, 40, 9)
    TRACON$ = Text34.TEXT 'TRABAJADOR CONVENCIONADO
      Call REPLA(XY$, TRACON$, 49, 1)
    TIPOP$ = Text35.TEXT 'TIPO DE OPERACION
      Call REPLA(XY$, TIPOP$, 50, 1)
    CONYU$ = Text20.TEXT 'CONYUGE
      Call REPLA(XY$, CONYU$, 51, 1)
    Call GRAREG("DATEMDGI", XY$, J1&)
    '
    '\\\OT-05-0577-WAR-FIN///
  End If
  '
End Sub
Function UltimoRegistro(COLUMNA) As Double
   Dim valor As Variant
   Dim CantFilasVacias&
   CantFilasVacias& = 0
   UltFila& = 0
   For Fila& = 2 To HOJA.Rows.Count
      If CantFilasVacias >= 100 Then
         GoTo 10
      End If
         valor = TRIM$(HOJA.Cells(Fila&, COLUMNA))
         If valor = "" Then 'Si no hay valor
            If UltFila& = 0 Then 'Si no hay registro de ultima fila se registra, si no se mantiene la almacenada
               UltFila& = Fila& - 1
            End If
            CantFilasVacias& = CantFilasVacias& + 1
           Else
            CantFilasVacias& = 0
            UltFila& = 0
          End If
   Next Fila&
10 UltimoRegistro = UltFila&
End Function

Private Sub Impofamiliares_Click()
    Set objexcel = New Excel.Application
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    '
    DoEvents
    Screen.MousePointer = 11
    '
    Ruta$ = UCase$(CommonDialog1.FileName)
    Set Libro = objexcel.Workbooks.Open(Ruta$)
    Set HOJA = Libro.Sheets(1)
    '
    URECORD& = UltimoRegistro(1)
    JJ& = 0
    For Fila& = 2 To URECORD&
        Call TRACE(24, Fila&, URECORD&)
        '
        LEGA% = XVALO(HOJA.Cells(Fila&, 1))
        PAREN$ = HOJA.Cells(Fila&, 2)
        NOAP$ = HOJA.Cells(Fila&, 3)
        FENA = CVINT(CDate(HOJA.Cells(Fila&, 4)))
        DOCUFA$ = HOJA.Cells(Fila&, 5)
        YY$ = REGBLAN$("FAMCAR")
        Call REPLA(YY$, MKI$(LEGA%), 1, 2)
        Call REPLA(YY$, NOAP$, 3, 25)
        Call REPLA(YY$, MKI$(FENA), 28, 2)
            PARE$ = "HIJO"
            If PAREN$ = "E" Then PARE$ = "ESPOSA"
        Call REPLA(YY$, PARE$, 30, 10)
        Call REPLA(YY$, DOCUFA$, 40, 11)
        '
        JJ& = JJ& + 1
        Call GRAREG("FAMCAR", YY$, JJ&)
        '
    Next Fila&
    '
    Call COMUNI("Proceso de Importacion de Familiares Finalizado")
99  '
End Sub

Private Sub Label2_dblClick()
  '
  NELGUI% = Val(Text1)
  If NELGUI% <= 0 Then Exit Sub
  '
  NR$ = RECFILT$("PADREMP", 1, MKI$(NELGUI%))
  NR1& = CVS(NR$) 'Nº DE REGISTRO DE EMPLEADO
  '
  X$ = REGLEIDO$("ACTIVA", NR1&)
  Call BLANARCH("!CONCOPCI")
  For J& = 1 To 96 Step 2
    C1% = CVI(Mid$(X$, J&, 2)) 'Nº DE CODIGO
    If C1% > 0 Then
      Y$ = REGBLAN$("!CONCOPCI")
      Call REPLA(Y$, MKI$(C1%), 1, 2)
      Call REGAPP("!CONCOPCI", Y$)
    End If
  Next J&
  '
  For J1& = 97 To 128 'LEO LA CANTIDAD
    CAN1 = Asc(Mid$(X$, J1&, 1))
    If CAN1 > 0 Then
      CODASI% = Int(400 + J1& - 96)
      Y$ = REGBLAN$("!CONCOPCI")
      Call REPLA(Y$, MKI$(CODASI%), 1, 2) 'CODIGO
      Call REPLA(Y$, MKI$(Int(CAN1)), 7, 2) 'CANTIDAD
      Call REGAPP("!CONCOPCI", Y$)
    End If
  Next J1&
  '
15 VTB% = VENTABU%("CONCOPCI")
  If VTB% >= 0 Then
    RENUE$ = "": RENUE2$ = String$(32, 0)
    For J& = 1 To ULTREG&("!CONCOPCI")
      X1$ = REGLEIDO$("!CONCOPCI", J&)
      '
      NUCO% = CVI(Mid$(X1$, 1, 2))
      CANTI1 = CVI(Mid$(X1$, 7, 2))
      'ADVERTENCIA
      If NUCO% < 400 And CANTI1 > 0 Then
        Call MENSERR(24, "Solo es Posible Ingresar Cantidad a las Asignaciones Familiares.")
        GoTo 15
      End If
      '
      If NUCO% <= 400 Then
            RENUE$ = RENUE$ + MKI$(NUCO%)
         Else
           If CANTI1 > 0 Then
             Call REPLA(RENUE2$, Chr$(CANTI1), NUCO% - 400, 1)
           End If
      End If
    Next J&
    RENUE$ = Left$(RENUE$ + String$(96, 0), 96) + RENUE2$
    Call GRAREG("ACTIVA", RENUE$, NR1&) 'RENGLON SEGUN EMPLEADO
    Call CARCONCEOP(NELGUI%) 'REFRESCO LIST1
  End If
  '
End Sub

Private Sub Label34_Change(Index As Integer)
   '
   Screen.MousePointer = 11
   '
   Picture1(Index).Visible = True
   Label34(Index).Visible = True
   Refresh
   '
   Image1(Index).Picture = LoadPicture("")
   Picture1(Index).Cls
   COPLANO$ = Label34(Index)
   If COPLANO$ <> "" Then
      If EXISTE%(COPLANO$) > 0 Then
         On Error GoTo 90
         Picture1(Index).Visible = False
         Image1(Index).Picture = LoadPicture(COPLANO$)
         Picture1(Index).Top = TOPEINI(Index)
         Picture1(Index).Left = IZQUINI(Index)
         Picture1(Index).Width = ANCHINI(Index)
         Picture1(Index).Height = ALTUINI(Index)
         Picture1(Index).Picture = LoadPicture("")
         Picture1(Index).Refresh
         '
         HIMAGEN = Image1(Index).Picture.Width
         VIMAGEN = Image1(Index).Picture.Height
         
         HCUADRO = Picture1(Index).Width
         VCUADRO = Picture1(Index).Height
         '
         REDUH = HIMAGEN / HCUADRO
         REDUV = VIMAGEN / VCUADRO
         REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
         Picture1(Index).Height = VCUADRO * REDUV / REDUC
         Picture1(Index).Width = HCUADRO * REDUH / REDUC
         Picture1(Index).Refresh
         '
         HCUADRO = Picture1(Index).Width
         VCUADRO = Picture1(Index).Height
         Picture1(Index).Left = (Frame1(Index).Width - HCUADRO) / 2
         '
         Picture1(Index).PaintPicture Image1(Index).Picture, 0, 0, HCUADRO, VCUADRO
         Picture1(Index).Visible = True
         Picture1(Index).Refresh
         MODIFIC% = 1
       Else
         'Call MENSERR(24, "No Existe Archivo '" + COPLANO$ + "'.")
      End If
   End If
   GoTo 98
   '
90 On Error GoTo 0
   Call MENSERR(24, "Imagen no Válida")
   '
98 Screen.MousePointer = 1
   '
End Sub

Private Sub Label34_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  '
  Label34(Index).ToolTipText = Label34(Index).Caption
  '
End Sub

Private Sub Label37_DBLClick()
  '
  NELGUI% = Val(Text1)
  If NELGUI% <= 0 Then Exit Sub
  '
  Call TRALOFIL("FAMCAR", 1, MKI$(NELGUI%))
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(str$(Left + 1200))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(str$(Top + SSTab1.Top + 1800))
  VTB% = VENTABU%("FAMCAR/" + ATRI$)
  If VTB% >= 0 Then
        List3.Clear
        '
        For J& = 1 To ULTREG&("!FAMCAR")
            XX$ = REGLEIDO$("!FAMCAR", J&)
            NOAP$ = Mid$(XX$, 3, 25)
            FENA = CVI(Mid$(XX$, 28, 2))
            PAREN$ = Mid$(XX$, 30, 10)
            DOCUFA$ = Mid$(XX$, 40, 11)
            YY$ = Space(68)
            Call REPLA(YY$, NOAP$, 1, 27)
            Call REPLA(YY$, FECHATEX$(FENA), 30, 10)
            Call REPLA(YY$, PAREN$, 41, 10)
            Call REPLA(YY$, DOCUFA$, 53, 11)
            List3.AddItem YY$
        Next J&
        '
        Call TRACENFIL("FAMCAR", 1, MKI$(NELGUI%))
  End If
  '
End Sub



Private Sub LiCuInEmp_Click()
  '
  Call COMUNI("Listados. No Disponible por el Momento.")
  '
End Sub

Private Sub LiGeClaEmp_Click()
  '
  Call Command13_Click
  '
End Sub

Private Sub List_Click()
  '
  Call Command13_Click
  '
End Sub

Private Sub List1_DblClick()
  Call Label2_dblClick
End Sub

Private Sub List3_dblClick()
  Call Label37_DBLClick
End Sub

Private Sub mnuAsiSubFam_Click()
With OPSUEL05
  .Option1(7).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub mnuCatego_Click()
With OPSUEL05
  .Option1(3).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub mnuGrupEmp_Click()
With OPSUEL05
  .Option1(0).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub MnuParInd_Click()
With OPSUEL05
  .Option1(8).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub mnuRemOrd_Click()
With OPSUEL05
  .Option1(5).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub mnuRetLegApo_Click()
With OPSUEL05
  .Option1(6).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub mnuSalNomBas_Click()
With OPSUEL05
  .Option1(4).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub NuevEmp_Click()
  '
  Call Command21_Click
  '
End Sub

Private Sub Print_Click()
  '
  Call Command25_Click
  '
End Sub


Private Sub Text1_Change()
  '
  '
   If MODIFIC% = 1 And Val(Text1) > 0 And Val(Text1) < 32767 And TRIM$(Text3.TEXT) <> "" Then
    If COMALTER%("Se han Modificado Datos del Empleado\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") = 1 Then
      Screen.MousePointer = 11
      MODIFIC% = 0
      Call GRAEMP(LEGACT%)
      Screen.MousePointer = 1
    End If
   End If
  '
  Call BLANFORMU
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  NLEG% = Val(Text1)
  Call CARDATEMP(NLEG%) 'CARGA DATOS DEL EMPLEADO
  Call CARFAMCAR(NLEG%) 'CARGA FAMILIARES A CARGO
  Call CARCONCEOP(NLEG%) 'CARGA DE CONCEPTOS OPCIONALES
  Call CARNOTAEMP 'CARGA ANOTACIONES DEL EMPLEADO
  MODIFIC% = 0

  '
End Sub

Private Sub Text1_DblClick()
  '
  Call Command9_Click
  '
End Sub

Private Sub Text10_Change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  MODIFIC% = 1
  '
End Sub

Private Sub Text11_Change()
  '
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  MODIFIC% = 1
  '
End Sub

Private Sub Text12_Change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  MODIFIC% = 1
  '
End Sub

Private Sub Text13_Change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  MODIFIC% = 1
  '
End Sub

Private Sub Text14_Change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  MODIFIC% = 1
  '
End Sub

Private Sub Text15_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text16_Change()
  '
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  Label23 = ECOARCH$("CATEGO", Text16)
  MODIFIC% = 1
  '
End Sub

Private Sub Text17_Change()
  '
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  Label31 = ECOARCH$("SECPROD", MKS$(XVALO(Text17)))
  MODIFIC% = 1
  '
End Sub

Private Sub Text17_LostFocus()
  '
  SEC1 = XVALO(Text17.TEXT)
  Text17 = FORMATNUM$(SEC1, "F7.3")
  '
End Sub

Private Sub Text19_Change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If Val(Text19) = 1 Then
    Label21 = "ACTIVO"
  Else
    Label21 = "BAJA"
  End If
  '
End Sub

Private Sub Text19_gotfocus()
  '
  Text1.SetFocus
  '
End Sub

Private Sub Text2_Change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  Label22 = ECOARCH("ESTCIV", MKI$(Val(Text2)))
  MODIFIC% = 1
  '
End Sub

Private Sub Text25_Change()
  '
  On Error Resume Next
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then
    On Error GoTo 0
    Exit Sub
  End If
  On Error GoTo 0
  MODIFIC% = 1
  '
End Sub

Private Sub Text26_Change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  Label19 = ECOARCH$("SEXTAB", TRIM$(Text26))
  MODIFIC% = 1
  '
End Sub
Sub text26_lostfocus()
  '
  XX$ = Text26
  Text26 = UCase(XX$)
  '
End Sub
Sub BLANFORMU()
  '
  Text17 = ""
  Text25 = ""
  Text6 = ""
  Text19 = ""
  Text3 = ""
  Text16 = ""
  Text26 = ""
  Text10 = ""
  Text29 = ""
  Label19 = ""
  Label20 = ""
  Label21 = ""
  Label22 = ""
  Label23 = ""
  Label31 = ""
  Text2 = ""
  Text4 = ""
  Text7 = ""
  Text36 = ""
  Text37 = ""
  Text8 = ""
  Text13 = ""
  Text14 = ""
  Text9 = ""
  Text5 = ""
  Text11 = ""
  Text12 = ""
  Text6 = ""
  Text19 = ""
  Text28 = ""
  '
End Sub
Sub CREATAB()
  'TABLA DE TIPO DE ACCESO
  If ULTREG&("TABACCE") < 1 Then
    Call BLANARCH("TABACCE")
    Call REGAPP("TABACCE", Chr$(0) + "Libre")
    Call REGAPP("TABACCE", Chr$(1) + "Reservado")
    Call CIERRARCH("TABACCE")
  End If
  'TABLA  DE FRECUENCIA
  If ULTREG&("TABFREQ") < 1 Then
    Call BLANARCH("TABFREQ")
    Call REGAPP("TABFREQ", Chr$(1) + "Mensual")
    Call REGAPP("TABFREQ", Chr$(2) + "Quincenal")
    Call CIERRARCH("TABFREQ")
  End If
  'TABLA DE MODO
  If ULTREG&("TABMODO") < 1 Then
    Call BLANARCH("TABMODO")
    Call REGAPP("TABMODO", Chr$(0) + "Directo")
    Call REGAPP("TABMODO", Chr$(1) + "Formula")
    Call CIERRARCH("TABMODO")
  End If
  'TABLA DE TABLA CARÁCTER
  If ULTREG&("TACARAC") < 1 Then
    Call BLANARCH("TACARAC")
    Call REGAPP("TACARAC", Chr$(0) + "Opcional")
    Call REGAPP("TACARAC", Chr$(1) + "Mandatorio Mensuales")
    Call REGAPP("TACARAC", Chr$(2) + "Mandatorio Quincenales")
    Call REGAPP("TACARAC", Chr$(3) + "Mandatorio General")
    Call CIERRARCH("TACARAC")
  End If
  'TABLA DE SUJETO A RETENCION
  If ULTREG&("TASURET") < 1 Then
    Call BLANARCH("TASURET")
    Call REGAPP("TASURET", Chr$(0) + "Exento")
    Call REGAPP("TASURET", Chr$(1) + "Gravado")
    Call CIERRARCH("TASURET")
  End If
  'TABLA DE QUINCENA DE LIQUIDACION
  If ULTREG&("TAQUINLI") < 1 Then
    Call BLANARCH("TAQUINLI")
    Call REGAPP("TAQUINLI", Chr$(0) + "Todas las Liquidaciones")
    Call REGAPP("TAQUINLI", Chr$(1) + "Solo 1ra Quincena")
    Call REGAPP("TAQUINLI", Chr$(2) + "Solo 2da Quincena")
    Call CIERRARCH("TAQUINLI")
  End If
  'TABLA DE ESTADOS
  If ULTREG&("TASTATUS") < 1 Then
    Call BLANARCH("TASTATUS")
    Call REGAPP("TASTATUS", "A" + "Activo")
    Call REGAPP("TASTATUS", "S" + "Suspendido")
    Call REGAPP("TASTATUS", "F" + "Fantasma")
    Call CIERRARCH("TASTATUS")
  End If
  'TABLA DE TIPOS DE REMUNERACION
  If ULTREG&("TATIREM") < 1 Then
    Call BLANARCH("TATIREM")
    Call REGAPP("TATIREM", Chr$(1) + "Normal")
    Call REGAPP("TATIREM", Chr$(2) + "Gratificacion")
    Call CIERRARCH("TATIREM")
  End If
  'TABLA DE CALCULO DE REMUNERACIONES
  If ULTREG&("TAMOCREM") < 1 Then
    Call BLANARCH("TAMOCREM")
    Call REGAPP("TAMOCREM", Chr$(0) + "Calculo Interno")
    Call REGAPP("TAMOCREM", Chr$(1) + "Informe Importe")
    Call REGAPP("TAMOCREM", Chr$(2) + "Informe Horas")
    Call REGAPP("TAMOCREM", Chr$(3) + "Cantidad")
    Call REGAPP("TAMOCREM", Chr$(4) + "Escalonamiento")
    Call CIERRARCH("TAMOCREM")
  End If
  'TABLA DE RECALCULO OFF / ON
  If ULTREG&("TARECALC") < 1 Then
    Call BLANARCH("TARECALC")
    Call REGAPP("TARECALC", Chr$(0) + "Recalculo off")
    Call REGAPP("TARECALC", Chr$(1) + "Recalculo on")
    Call CIERRARCH("TARECALC")
  End If
  'TABLA DE MODO DE CALCULO DE RETENCIONES
  If ULTREG&("TAMOCRET") < 1 Then
    Call BLANARCH("TAMOCRET")
    Call REGAPP("TAMOCRET", Chr$(0) + "Por Formula")
    Call REGAPP("TAMOCRET", Chr$(1) + "Informe Importe")
    Call REGAPP("TAMOCRET", Chr$(2) + "Porcentajes")
    Call REGAPP("TAMOCRET", Chr$(3) + "Sumas Fijas")
    Call CIERRARCH("TAMOCRET")
  End If
  'TABLA DE TIPO DE ACTUALIZACION (SALARIOS NOMINALES)
  If ULTREG&("TIACTSAL") < 1 Then
    Call BLANARCH("TIACTSAL")
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(1), 1, 2)
    Call REPLA(Y$, "Actualizacion", 3, 30)
    Call REGAPP("TIACTSAL", Y$)
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(2), 1, 2)
    Call REPLA(Y$, "Listado", 3, 30)
    Call REGAPP("TIACTSAL", Y$)
    Call CIERRARCH("TIACTSAL")
  End If
  'TABLA DE SEXO
  If ULTREG&("SEXTAB") < 1 Then
    Call BLANARCH("SEXTAB")
    Y$ = Space(32)
    Call REPLA(Y$, "F", 1, 1)
    Call REPLA(Y$, "FEMENINO", 2, 31)
    Call REGAPP("SEXTAB", Y$)
    Y$ = Space(32)
    Call REPLA(Y$, "M", 1, 1)
    Call REPLA(Y$, "MASCULINO", 2, 31)
    Call REGAPP("SEXTAB", Y$)
    Call CIERRARCH("SEXTAB")
  End If
  'TABLA DE ESTADO CIVIL
  If ULTREG&("ESTCIV") < 8 Then
    Call BLANARCH("ESTCIV")
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(1), 1, 2)
    Call REPLA(Y$, "SOLTERO/A", 3, 30)
    Call REGAPP("ESTCIV", Y$)
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(2), 1, 2)
    Call REPLA(Y$, "CASADO/A", 3, 30)
    Call REGAPP("ESTCIV", Y$)
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(3), 1, 2)
    Call REPLA(Y$, "SEPARADO/A", 3, 30)
    Call REGAPP("ESTCIV", Y$)
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(4), 1, 2)
    Call REPLA(Y$, "DIVORCIADO/A", 3, 30)
    Call REGAPP("ESTCIV", Y$)
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(5), 1, 2)
    Call REPLA(Y$, "VIUDO/A", 3, 30)
    Call REGAPP("ESTCIV", Y$)
    Call CIERRARCH("ESTCIV")
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(6), 1, 2)
    Call REPLA(Y$, "APARENTE MATRIMONIO", 3, 30)
    Call REGAPP("ESTCIV", Y$)
    Call CIERRARCH("ESTCIV")
    '\\\OT-06-0472-RG-07/11/06///
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(7), 1, 2)
    Call REPLA(Y$, "CONVIENCIA", 3, 30)
    Call REGAPP("ESTCIV", Y$)
    Call CIERRARCH("ESTCIV")
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(8), 1, 2)
    Call REPLA(Y$, "CONCUBINATO", 3, 30)
    Call REGAPP("ESTCIV", Y$)
    Call CIERRARCH("ESTCIV")
    '\\\OT-06-0472-RG-FIN///
  End If
  '\\\LIQUIDACIONES///
  'TABLA DE TIPOS DE LIQUIDACION
  If ULTREG&("TATIPLIQ") < 1 Then
    Call BLANARCH("TATIPLIQ")
    Call REGAPP("TATIPLIQ", Chr$(1) + "Primera Quincena")
    Call REGAPP("TATIPLIQ", Chr$(2) + "Segunda Quincena")
    Call REGAPP("TATIPLIQ", Chr$(3) + "Retroactivo")
    Call CIERRARCH("TATIPLIQ")
  End If
  'TABLA DE ASIGNACIONES FLIARES.
  If ULTREG&("TASIGFAM") < 1 Then
    Call BLANARCH("TASIGFAM")
    Call REGAPP("TASIGFAM", Chr$(0) + "No se Liquida")
    Call REGAPP("TASIGFAM", Chr$(1) + "Si se Liquida")
    Call CIERRARCH("TASIGFAM")
  End If
  'TABLA DE MESES
  If ULTREG&("TAMESE") < 1 Then
    Call BLANARCH("TAMESE")
    Call REGAPP("TAMESE", Chr$(1) + "Enero")
    Call REGAPP("TAMESE", Chr$(2) + "Febrero")
    Call REGAPP("TAMESE", Chr$(3) + "Marzo")
    Call REGAPP("TAMESE", Chr$(4) + "Abril")
    Call REGAPP("TAMESE", Chr$(5) + "Mayo")
    Call REGAPP("TAMESE", Chr$(6) + "Junio")
    Call REGAPP("TAMESE", Chr$(7) + "Julio")
    Call REGAPP("TAMESE", Chr$(8) + "Agosto")
    Call REGAPP("TAMESE", Chr$(9) + "Septiembre")
    Call REGAPP("TAMESE", Chr$(10) + "Octubre")
    Call REGAPP("TAMESE", Chr$(11) + "Noviembre")
    Call REGAPP("TAMESE", Chr$(12) + "Diciembre")
    Call REGAPP("TAMESE", Chr$(13) + "Aguinaldo")
    Call REGAPP("TAMESE", Chr$(14) + "Vacaciones")
    Call CIERRARCH("TAMESE")
  End If
End Sub

Private Sub Text28_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text29_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text3_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text4_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text5_DblClick()
   Call Command22_Click
End Sub

Private Sub Text6_Change()
  '
  If Val(Text1) > 32767 Then
    Call COMUNI("Número de Legajo no Válido.")
    Text1.SelStart = 0
    Text1.SelLength = (Len(Text1))
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  '
  Label20 = ECOARCH("DATGREM1", MKI$(Val(Text6)))
  MODIFIC% = 1
  '
End Sub

Private Sub Text7_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text8_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text9_Change()
  MODIFIC% = 1
End Sub

Private Sub Text5_Change()
  MODIFIC% = 1
End Sub

Sub CARCONCEOP(NELG%)
  '
  If NELG% <= 0 Then Exit Sub
  '
  NR$ = RECFILT$("PADREMP", 1, MKI$(NELG%))
  NR1& = CVS(NR$) 'Nº DE REGISTRO DE EMPLEADO
  '
  X$ = REGLEIDO$("ACTIVA", NR1&)
  Call BLANARCH("!CONCOPCI")
  For J& = 1 To 96 Step 2
    C1% = CVI(Mid$(X$, J&, 2)) 'Nº DE CODIGO
    If C1% > 0 Then
      Y$ = REGBLAN$("!CONCOPCI")
      Call REPLA(Y$, MKI$(C1%), 1, 2)
      Call REGAPP("!CONCOPCI", Y$)
      '
    End If
  Next J&
  '
  For J1& = 97 To 128 'LEO LA CANTIDAD
    CAN1 = Asc(Mid$(X$, J1&, 1))
    If CAN1 > 0 Then
      CODASI% = Int(400 + J1& - 96)
      Y$ = REGBLAN$("!CONCOPCI")
      Call REPLA(Y$, MKI$(CODASI%), 1, 2) 'CODIGO
      Call REPLA(Y$, MKI$(Int(CAN1)), 7, 2) 'CANTIDAD
      Call REGAPP("!CONCOPCI", Y$)
    End If
  Next J1&
  'AGREGO A LA LISTA
  List1.Clear
  For J& = 1 To ULTREG&("!CONCOPCI")
    XX$ = REGLEIDO$("!CONCOPCI", J&)
    Z$ = Space$(64)
    COD2% = CVI(Mid$(XX$, 1, 2))
    DESCRI$ = ECOARCH("ECONCEP", MKI$(COD2%))
    UM$ = Mid$(XX$, 3, 4)
    CANT2% = CVI(Mid$(XX$, 7, 2))
    FORMU$ = Mid$(XX$, 9, 56)
    '
    Call REPLA(Z$, FORMATNUM(COD2%, "F5"), 1, 5)
    Call REPLA(Z$, DESCRI$, 7, 30)
    Call REPLA(Z$, UM$, 39, 4)
    Call REPLA(Z$, FORMATNUM(CANT2%, "F4"), 45, 4)
    Call REPLA(Z$, FORMU$, 52, 56)
    '
    List1.AddItem Z$
  Next J&
  '
End Sub
Sub APACTPADRO() 'APAREA PADREMP = ACTIVA
  '
  PANU& = ULTREG("PADREMP")
  CONU& = ULTREG("ACTIVA")
  DIF& = PANU& - CONU&
  If DIF& > 0 Then
    For J& = 1 To DIF&
      YYY$ = String$(128, Chr$(0))
      'PARA Q' NO SEA VÁLIDO EL NUMERO DE CONCEPTO OPCIONAL
      Call REPLA(YYY$, MKI$(-1), 1, 2)
      Call REGAPP("ACTIVA", YYY$)
    Next J&
  End If
  '
  Call CIERRARCH("PADREMP")
  Call CIERRARCH("ACTIVA")
  '
End Sub
Sub APAREMBA() 'APAREA PADREMP = REMBAS
  '
  PANU& = ULTREG("PADREMP")
  CONU& = ULTREG("REMBAS")
  DIF& = PANU& - CONU&
  If DIF& > 0 Then
    For J& = 1 To DIF&
      YYY$ = String$(128, Chr$(0))
      Call REPLA(YYY$, Space(48), 1, 48)
      Call REGAPP("REMBAS", YYY$)
    Next J&
  End If
  '
  Call CIERRARCH("PADREMP")
  Call CIERRARCH("REMBAS")
  '
End Sub
Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub
Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub
Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub
Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub
'
Sub GRANOTEMP() 'GRABA ANOTACIONES DE EMPLEADOS
   '
   NLEG% = Val(TRIM$(Text1.TEXT))
   If NLEG% < 1 Then Exit Sub
   If TRIM(RECFILT$("PADREMP", 1, MKI$(NLEG%))) = "" Then Exit Sub
   '
   If NLEG% > 0 Then
         If MODIFIC% < 1 Then
           Exit Sub
         End If
         '
         If DERACCE%("ANOTACLI", "Anotaciones de Empleado") < 2 Then
            Exit Sub
         End If
         '
         Screen.MousePointer = 11
         URN& = ULTREG&("NOTAEMP")
         RNC$ = RECFILT$("NOTAEMP", 1, MKI$(NLEG%))
         KK& = (-3)
         '
         TEXTO1$ = Text28.TEXT + Chr$(13) + Chr$(10)
         Call FRATEXTO(TEXTO1$, 78)
         For LL% = 1 To CARETEX%
            TTYY$ = MKI$(NLEG%) + RETEX$(LL%)
            KK& = KK& + 4
            If KK& <= Len(RNC$) - 3 Then
                RENOTA& = CVS(Mid$(RNC$, KK&, 4))
              Else
                URN& = URN& + 1
                RENOTA& = URN&
            End If
            Call GRAREG("NOTAEMP", TTYY$, RENOTA&)
         Next LL%
         '
20       While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Call GRAREG("NOTAEMP", MKI$(0) + Space$(78), RENOTA&)
         Wend
         '
         Call CIERRARCH("NOTAEMP")
        
         Screen.MousePointer = 1
   End If
End Sub
Sub CARNOTAEMP() 'CARGA NOTAS DEL EMPLEADO
   '
   NLEG% = Val(Text1.TEXT)
   Text28.TEXT = ""
   '
   If NLEG% > 0 Then
      If TRIM$(RECFILT$("PADREMP", 1, MKI$(NLEG%))) <> "" Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("NOTAEMP", 1, MKI$(NLEG%))
         URN& = ULTREG&("NOTAEMP")
         '
         For U& = 1 To Len(RNC$) Step 4
            RENOTA& = CVS(Mid$(RNC$, U&, 4))
            If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                  TTYY$ = RTrim$(Mid$(REGLEIDO$("NOTAEMP", RENOTA&), 3))
                  If Text28.TEXT <> "" Then
                     Text28.TEXT = Text28.TEXT + Chr$(13) + Chr$(10)
                  End If
                  Text28.TEXT = Text28.TEXT + RTrim$(TTYY$)
               End If
            End If
         Next U&
         Screen.MousePointer = 1
      End If
   End If
   Text28.Refresh
   '
End Sub
Sub APAPADREMP()
  'APAREA SI NO EXISTE PADREMP O NO TIENE REGISTROS
  If ULTREG&("PADREMP") <= 0 Then
    Call BLANARCH("PADREMP")
    For J& = 1 To ULTREG("PADRON")
      PA1$ = REGLEIDO$("PADRON", J&)
      PA2$ = REGLEIDO$("PADAUX", J&)
      PA3$ = REGLEIDO$("PADAUX2", J&)
      NULE% = CVI(Mid$(PA1$, 1, 2))
      If NULE% > 0 Then 'PARA QUE NO GRABE REGISTROS EN BLANCO
        PADREM$ = Space(256)
        Call REPLA(PADREM$, PA1$, 1, 64)
        'CATEGO1$ = Mid$(PA1$, 54, 9)
        Call REPLA(PADREM$, PA2$, 65, 96)
        Call REPLA(PADREM$, PA3$, 161, 48)
        Call REGAPP("PADREMP", PADREM$)
      End If
    Next J&
  End If
  '
End Sub
'\\\OT-05-0577-WAR-25/07/06///
Sub APAREDGI()
  'APAREA SI NO EXISTE DATEMDGI
  '
  If ULTREG&("DATEMDGI") < ULTREG&("PADREMP") Then
    XBLA$ = REGBLAN$("DATEMDGI")
    For J& = 1 To ULTREG&("PADREMP")
       X$ = REGLEIDO$("PADREMP", J&)
       NLEGA1% = CVI(Mid$(X$, 1, 2))
       NCUI1$ = Mid$(X$, 209, 13)
       Z$ = XBLA$
       Call REPLA(Z$, MKI$(NLEGA1%), 1, 2)
       Call REPLA(Z$, NCUI1$, 3, 13)
       Call REGAPP("DATEMDGI", Z$)
    Next J&
    Call SETULTREG("DATEMDGI", J&)
  End If
  '
  '
End Sub
'\\\OT-05-0577-WAR-FIN///
Private Sub Text9_DblClick()
   Call Command11_Click
End Sub
