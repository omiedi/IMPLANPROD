VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AORTRA_GES04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8D8C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes - A-B-M-L Base de Datos de Clientes"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "AORTRA_GES04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
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
      Height          =   330
      Left            =   9240
      TabIndex        =   143
      Top             =   840
      Width           =   1335
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
      Height          =   315
      Left            =   10395
      Picture         =   "AORTRA_GES04.frx":0442
      TabIndex        =   141
      Top             =   420
      Width           =   175
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9240
      TabIndex        =   140
      Top             =   430
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
      Height          =   315
      Left            =   10395
      Picture         =   "AORTRA_GES04.frx":074C
      TabIndex        =   138
      Top             =   105
      Width           =   175
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
      Height          =   330
      Left            =   9240
      TabIndex        =   137
      Top             =   105
      Width           =   1020
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
      Height          =   330
      Left            =   6510
      TabIndex        =   135
      Top             =   840
      Width           =   1755
   End
   Begin VB.TextBox Text7 
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
      Left            =   5250
      TabIndex        =   133
      Top             =   840
      Width           =   705
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
      Height          =   390
      Left            =   5250
      TabIndex        =   131
      Top             =   430
      Width           =   3015
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
      Height          =   330
      Left            =   5250
      TabIndex        =   124
      Top             =   105
      Width           =   2805
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
      Left            =   8085
      Picture         =   "AORTRA_GES04.frx":0A56
      TabIndex        =   123
      Top             =   105
      Width           =   175
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
      Left            =   9975
      TabIndex        =   118
      Top             =   1365
      Visible         =   0   'False
      Width           =   1990
      Begin VB.CommandButton Command20 
         Height          =   500
         Left            =   1280
         Picture         =   "AORTRA_GES04.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Búsqueda por CUIT, Domicilio, Nombre de Fantasía, etc"
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
         Picture         =   "AORTRA_GES04.frx":20C6
         Style           =   1  'Graphical
         TabIndex        =   120
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
         Picture         =   "AORTRA_GES04.frx":23D0
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "Restablece los Valores Anteriores - Descarta Cambios Efectuados"
         Top             =   250
         Width           =   550
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   8100
      Left            =   11025
      ScaleHeight     =   8040
      ScaleWidth      =   1215
      TabIndex        =   105
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command23 
         Caption         =   "Parar"
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
         Picture         =   "AORTRA_GES04.frx":26DA
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Suspensión de Operaciones"
         Top             =   2580
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   116
         Top             =   7140
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   117
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
         Picture         =   "AORTRA_GES04.frx":29E4
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Imprime Ficha del Cliente"
         Top             =   5625
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
         Picture         =   "AORTRA_GES04.frx":304E
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Emisión de Listados"
         Top             =   4935
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
         Picture         =   "AORTRA_GES04.frx":3358
         Style           =   1  'Graphical
         TabIndex        =   113
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
         Picture         =   "AORTRA_GES04.frx":379A
         Style           =   1  'Graphical
         TabIndex        =   112
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
         Picture         =   "AORTRA_GES04.frx":3AA4
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Baja Lógica de la Cuenta"
         Top             =   3270
         Width           =   650
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Sort"
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
         Picture         =   "AORTRA_GES04.frx":3DAE
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Ordena Base de Datos de Clientes"
         Top             =   3960
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
         Picture         =   "AORTRA_GES04.frx":40B8
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Exit"
         Top             =   270
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
         Picture         =   "AORTRA_GES04.frx":4202
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Ayuda en Línea"
         Top             =   960
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
         Picture         =   "AORTRA_GES04.frx":450C
         Style           =   1  'Graphical
         TabIndex        =   110
         ToolTipText     =   "Alta de Nuevo Cliente"
         Top             =   1890
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "AORTRA_GES04.frx":4656
         Stretch         =   -1  'True
         Top             =   7455
         Width           =   1410
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   10710
      Sorted          =   -1  'True
      TabIndex        =   80
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
      Height          =   615
      Left            =   1155
      TabIndex        =   2
      Top             =   630
      Width           =   2925
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Top             =   105
      Width           =   1725
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
      Left            =   2940
      Picture         =   "AORTRA_GES04.frx":6578
      TabIndex        =   1
      Top             =   105
      Width           =   175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6675
      Left            =   150
      TabIndex        =   36
      Top             =   1365
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11774
      _Version        =   327681
      Tab             =   1
      TabHeight       =   520
      BackColor       =   14211264
      TabCaption(0)   =   "Especificacion de Venta"
      TabPicture(0)   =   "AORTRA_GES04.frx":6882
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Hs."
      TabPicture(1)   =   "AORTRA_GES04.frx":689E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Line3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Picture4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command6"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command8"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command17"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text28"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Picture2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Cuenta Corriente"
      TabPicture(2)   =   "AORTRA_GES04.frx":68BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame18"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame18 
         BorderStyle     =   0  'None
         Caption         =   "Frame18"
         Height          =   6000
         Left            =   -74600
         TabIndex        =   81
         Top             =   440
         Width           =   9990
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00E0E0E0&
            Height          =   2800
            Left            =   105
            ScaleHeight     =   2745
            ScaleWidth      =   9615
            TabIndex        =   87
            Top             =   3120
            Width           =   9675
            Begin VB.Frame Frame14 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   5630
               Left            =   0
               TabIndex        =   90
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
                  Height          =   2085
                  Left            =   -30
                  TabIndex        =   91
                  Top             =   0
                  Width           =   9675
               End
               Begin VB.Line Line19 
                  X1              =   7350
                  X2              =   7350
                  Y1              =   2100
                  Y2              =   2415
               End
               Begin VB.Line Line18 
                  X1              =   6195
                  X2              =   6195
                  Y1              =   2100
                  Y2              =   2415
               End
               Begin VB.Line Line17 
                  X1              =   4200
                  X2              =   4200
                  Y1              =   2100
                  Y2              =   2415
               End
               Begin VB.Line Line16 
                  X1              =   3045
                  X2              =   3045
                  Y1              =   2100
                  Y2              =   2415
               End
               Begin VB.Line Line15 
                  X1              =   0
                  X2              =   9660
                  Y1              =   2420
                  Y2              =   2420
               End
               Begin VB.Line Line14 
                  BorderWidth     =   2
                  X1              =   0
                  X2              =   9660
                  Y1              =   2085
                  Y2              =   2085
               End
               Begin VB.Label Label22 
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
                  Left            =   25
                  TabIndex        =   94
                  Top             =   2130
                  Width           =   1695
               End
               Begin VB.Label Label25 
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
                  Left            =   4330
                  TabIndex        =   93
                  Top             =   2130
                  Width           =   1695
               End
               Begin VB.Label Label28 
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
                  Left            =   7455
                  TabIndex        =   92
                  Top             =   2130
                  Width           =   1695
               End
               Begin VB.Line Line12 
                  X1              =   2205
                  X2              =   2205
                  Y1              =   1575
                  Y2              =   3360
               End
            End
            Begin VB.Frame Frame15 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               Height          =   330
               Left            =   0
               TabIndex        =   88
               Top             =   0
               Width           =   10515
               Begin VB.Line Line11 
                  X1              =   6195
                  X2              =   6195
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line10 
                  X1              =   4200
                  X2              =   4200
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line9 
                  X1              =   3045
                  X2              =   3045
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line8 
                  X1              =   2205
                  X2              =   2205
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Label Label36 
                  BackStyle       =   0  'Transparent
                  Caption         =   " Documento                     Nro      Emision                    Importe      Vencim.                      Saldo              "
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
                  TabIndex        =   89
                  Top             =   50
                  Width           =   10410
               End
               Begin VB.Line Line4 
                  X1              =   7350
                  X2              =   7350
                  Y1              =   0
                  Y2              =   1050
               End
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00E0E0E0&
            Height          =   2430
            Left            =   105
            ScaleHeight     =   2370
            ScaleWidth      =   9615
            TabIndex        =   82
            Top             =   525
            Width           =   9675
            Begin VB.Frame Frame16 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               Height          =   330
               Left            =   0
               TabIndex        =   83
               Top             =   0
               Width           =   10515
               Begin VB.Line Line5 
                  X1              =   7560
                  X2              =   7560
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Label Label37 
                  BackStyle       =   0  'Transparent
                  Caption         =   "   Fecha       Documento                      Nro                Debe                   Haber                  Saldo              "
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
                  TabIndex        =   84
                  Top             =   50
                  Width           =   10410
               End
               Begin VB.Line Line6 
                  X1              =   1155
                  X2              =   1155
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line7 
                  X1              =   3360
                  X2              =   3360
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line13 
                  X1              =   4200
                  X2              =   4200
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line20 
                  X1              =   5775
                  X2              =   5775
                  Y1              =   0
                  Y2              =   1050
               End
            End
            Begin VB.Frame Frame17 
               BackColor       =   &H00C0F0FF&
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   5580
               Left            =   0
               TabIndex        =   85
               Top             =   -1050
               Width           =   10515
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
                  TabIndex        =   86
                  Top             =   1365
                  Width           =   9670
               End
               Begin VB.Line Line21 
                  X1              =   1155
                  X2              =   1155
                  Y1              =   1470
                  Y2              =   3255
               End
            End
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Disp:"
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
            Index           =   1
            Left            =   7245
            TabIndex        =   102
            Top             =   150
            Width           =   645
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cheq:"
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
            Index           =   0
            Left            =   4935
            TabIndex        =   101
            Top             =   150
            Width           =   645
         End
         Begin VB.Label label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L.Cre:"
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
            Left            =   105
            TabIndex        =   100
            Top             =   150
            Width           =   645
         End
         Begin VB.Label label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.Cte:"
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
            Index           =   3
            Left            =   2520
            TabIndex        =   99
            Top             =   150
            Width           =   645
         End
         Begin VB.Label LTEXT36 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   780
            TabIndex        =   98
            Top             =   105
            Width           =   1500
         End
         Begin VB.Label LTEXT37 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3195
            TabIndex        =   97
            Top             =   105
            Width           =   1500
         End
         Begin VB.Label LTEXT35 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   5610
            TabIndex        =   96
            Top             =   105
            Width           =   1500
         End
         Begin VB.Label LTEXT23 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   7920
            TabIndex        =   95
            Top             =   105
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   450
         ScaleHeight     =   375
         ScaleWidth      =   9720
         TabIndex        =   70
         Top             =   1155
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
            TabIndex        =   71
            Top             =   105
            Width           =   9570
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6120
         Left            =   -74800
         TabIndex        =   44
         Top             =   420
         Width           =   10320
         Begin VB.Frame Frame12 
            Caption         =   "Contable"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1625
            Left            =   5145
            TabIndex        =   67
            Top             =   4420
            Width           =   5055
            Begin VB.TextBox Text4 
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
               Left            =   1995
               TabIndex        =   130
               Top             =   1155
               Width           =   2880
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
               Height          =   330
               Left            =   1725
               Picture         =   "AORTRA_GES04.frx":68D6
               TabIndex        =   128
               Top             =   1155
               Width           =   175
            End
            Begin VB.TextBox Text2 
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
               Left            =   945
               MaxLength       =   5
               TabIndex        =   127
               Text            =   " "
               Top             =   1155
               Width           =   780
            End
            Begin VB.CommandButton BOTSEL 
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
               Index           =   6
               Left            =   4680
               TabIndex        =   31
               Top             =   735
               Width           =   175
            End
            Begin VB.CommandButton BOTSEL 
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
               Index           =   7
               Left            =   1725
               TabIndex        =   29
               Top             =   315
               Width           =   175
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   7
               Left            =   1155
               TabIndex        =   28
               Text            =   " "
               Top             =   315
               Width           =   570
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   6
               Left            =   2415
               TabIndex        =   30
               Text            =   " "
               Top             =   735
               Width           =   2270
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo Divisional:"
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
               Left            =   -30
               TabIndex        =   129
               Top             =   1050
               Width           =   975
            End
            Begin VB.Label DETEXT 
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
               Index           =   7
               Left            =   1995
               TabIndex        =   79
               Top             =   315
               Width           =   2880
            End
            Begin VB.Label DETEXT 
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
               Index           =   6
               Left            =   420
               TabIndex        =   78
               Top             =   1470
               Visible         =   0   'False
               Width           =   4455
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
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
               Height          =   225
               Left            =   315
               TabIndex        =   69
               Top             =   420
               Width           =   765
            End
            Begin VB.Label Label42 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cuenta Madre:"
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
               Left            =   420
               TabIndex        =   68
               Top             =   840
               Width           =   1905
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Comerciales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2540
            Left            =   5145
            TabIndex        =   61
            Top             =   1780
            Width           =   5055
            Begin VB.TextBox Text20 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3150
               TabIndex        =   21
               Text            =   " "
               Top             =   2035
               Width           =   1710
            End
            Begin VB.TextBox Text21 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   420
               TabIndex        =   20
               Text            =   " "
               Top             =   2035
               Width           =   1815
            End
            Begin VB.CommandButton BOTSEL 
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
               Index           =   5
               Left            =   1730
               TabIndex        =   19
               Top             =   1275
               Width           =   175
            End
            Begin VB.CommandButton BOTSEL 
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
               Index           =   4
               Left            =   1730
               TabIndex        =   17
               Top             =   840
               Width           =   175
            End
            Begin VB.CommandButton BOTSEL 
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
               Index           =   3
               Left            =   1730
               TabIndex        =   15
               Top             =   420
               Width           =   175
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   3
               Left            =   1155
               TabIndex        =   14
               Text            =   " "
               Top             =   420
               Width           =   570
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   4
               Left            =   1155
               TabIndex        =   16
               Text            =   " "
               Top             =   840
               Width           =   570
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   5
               Left            =   1155
               TabIndex        =   18
               Text            =   " "
               Top             =   1260
               Width           =   570
            End
            Begin VB.Label DETEXT 
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
               Index           =   5
               Left            =   1995
               TabIndex        =   77
               Top             =   1260
               Width           =   2880
            End
            Begin VB.Label DETEXT 
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
               Index           =   4
               Left            =   1995
               TabIndex        =   76
               Top             =   840
               Width           =   2880
            End
            Begin VB.Label DETEXT 
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
               Index           =   3
               Left            =   1995
               TabIndex        =   75
               Top             =   420
               Width           =   2880
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Límite de Crédito:"
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
               Left            =   1890
               TabIndex        =   66
               Top             =   1770
               Width           =   1905
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Desc.Std:"
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
               Left            =   -630
               TabIndex        =   65
               Top             =   1770
               Width           =   1710
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Lista de Precios:"
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
               TabIndex        =   64
               Top             =   315
               Width           =   1185
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cond.Pago:"
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
               Left            =   0
               TabIndex        =   63
               Top             =   945
               Width           =   1080
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Vendedor:"
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
               Left            =   -630
               TabIndex        =   62
               Top             =   1365
               Width           =   1710
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Impuestos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2010
            Left            =   105
            TabIndex        =   56
            Top             =   4040
            Width           =   4845
            Begin VB.CommandButton Command22 
               Caption         =   "Cfg"
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
               Left            =   210
               TabIndex        =   126
               ToolTipText     =   "Datos Percepción I.Brutos Cliente Activo"
               Top             =   1540
               Width           =   540
            End
            Begin VB.CommandButton BOTSEL 
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
               Index           =   2
               Left            =   1260
               TabIndex        =   26
               Top             =   1155
               Width           =   175
            End
            Begin VB.CommandButton BOTSEL 
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
               Index           =   1
               Left            =   1260
               TabIndex        =   23
               Top             =   315
               Width           =   175
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   1
               Left            =   945
               MaxLength       =   2
               TabIndex        =   22
               Text            =   " "
               Top             =   315
               Width           =   360
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
               Left            =   2310
               MaxLength       =   15
               TabIndex        =   24
               Top             =   315
               Width           =   2355
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   2
               Left            =   945
               MaxLength       =   2
               TabIndex        =   25
               Text            =   " "
               Top             =   1155
               Width           =   360
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
               Left            =   2310
               MaxLength       =   15
               TabIndex        =   27
               Text            =   " "
               Top             =   1155
               Width           =   2355
            End
            Begin VB.Label DETEXT 
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
               Index           =   2
               Left            =   945
               TabIndex        =   74
               Top             =   1540
               Width           =   3720
            End
            Begin VB.Label DETEXT 
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
               Index           =   1
               Left            =   945
               TabIndex        =   73
               Top             =   700
               Width           =   3720
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "P-IVA:"
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
               Left            =   -525
               TabIndex        =   60
               Top             =   420
               Width           =   1380
            End
            Begin VB.Label Label34 
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
               Height          =   330
               Left            =   315
               TabIndex        =   59
               Top             =   420
               Width           =   1905
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cat I-B:"
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
               Left            =   -525
               TabIndex        =   58
               Top             =   1260
               Width           =   1380
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nro I-B:"
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
               Left            =   315
               TabIndex        =   57
               Top             =   1260
               Width           =   1905
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Agenda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2180
            Left            =   105
            TabIndex        =   51
            Top             =   1780
            Width           =   4845
            Begin VB.TextBox Text36 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   1365
               TabIndex        =   103
               Text            =   " "
               Top             =   1755
               Width           =   3300
            End
            Begin VB.TextBox Text19 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   1365
               TabIndex        =   13
               Text            =   " "
               Top             =   1395
               Width           =   3300
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
               Height          =   280
               Left            =   1365
               TabIndex        =   10
               Text            =   " "
               Top             =   315
               Width           =   3300
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
               Height          =   280
               Left            =   1365
               TabIndex        =   11
               Text            =   " "
               Top             =   675
               Width           =   3300
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
               Height          =   280
               Left            =   1365
               TabIndex        =   12
               Text            =   " "
               Top             =   1035
               Width           =   3300
            End
            Begin VB.Label Label71 
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
               Left            =   -630
               TabIndex        =   104
               Top             =   1840
               Width           =   1905
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Contacto:"
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
               TabIndex        =   55
               Top             =   1480
               Width           =   1905
            End
            Begin VB.Label Label31 
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
               Left            =   -630
               TabIndex        =   54
               Top             =   400
               Width           =   1905
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fax:"
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
               TabIndex        =   53
               Top             =   760
               Width           =   1905
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "E-Mail:"
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
               TabIndex        =   52
               Top             =   1120
               Width           =   1905
            End
         End
         Begin VB.Frame Frame5 
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
            Height          =   1695
            Left            =   105
            TabIndex        =   45
            Top             =   0
            Width           =   10095
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
               Left            =   7980
               TabIndex        =   4
               Text            =   " "
               Top             =   315
               Width           =   1935
            End
            Begin VB.CommandButton BOTSEL 
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
               Index           =   0
               Left            =   1785
               TabIndex        =   8
               Top             =   1155
               Width           =   175
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   0
               Left            =   1365
               TabIndex        =   7
               Text            =   "  "
               Top             =   1155
               Width           =   465
            End
            Begin VB.TextBox Text25 
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
               Left            =   1365
               TabIndex        =   3
               Top             =   315
               Width           =   5295
            End
            Begin VB.TextBox Text26 
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
               Left            =   1365
               TabIndex        =   5
               Text            =   " "
               Top             =   735
               Width           =   1305
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
               Left            =   3885
               TabIndex        =   6
               Text            =   " "
               Top             =   735
               Width           =   6030
            End
            Begin VB.TextBox Text29 
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
               Left            =   7035
               TabIndex        =   9
               Text            =   " "
               Top             =   1155
               Width           =   2880
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre de Fantasía:"
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
               Left            =   6750
               TabIndex        =   122
               Top             =   230
               Width           =   1110
            End
            Begin VB.Label DETEXT 
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
               Index           =   0
               Left            =   2100
               TabIndex        =   72
               Top             =   1155
               Width           =   4005
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Localidad:"
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
               Left            =   1890
               TabIndex        =   50
               Top             =   840
               Width           =   1905
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Provincia:"
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
               Top             =   1260
               Width           =   1185
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Calle y Nro:"
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
               TabIndex        =   48
               Top             =   420
               Width           =   1905
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cod.Postal:"
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
               TabIndex        =   47
               Top             =   840
               Width           =   1905
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "País:"
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
               Left            =   5040
               TabIndex        =   46
               Top             =   1260
               Width           =   1905
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
         Left            =   450
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   1575
         Width           =   9780
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         Height          =   260
         Left            =   5310
         TabIndex        =   43
         Top             =   5790
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         Height          =   260
         Left            =   6495
         TabIndex        =   33
         Top             =   5790
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pegar"
         Height          =   260
         Left            =   7665
         TabIndex        =   34
         Top             =   5790
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Left            =   8955
         TabIndex        =   35
         Top             =   5790
         Width           =   1290
      End
      Begin VB.CommandButton Command7 
         Height          =   260
         Left            =   450
         TabIndex        =   42
         Top             =   5790
         Width           =   4830
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         Height          =   1590
         Left            =   2310
         ScaleHeight     =   1530
         ScaleWidth      =   2055
         TabIndex        =   41
         Top             =   1875
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Line Line3 
         X1              =   10230
         X2              =   10230
         Y1              =   1155
         Y2              =   6000
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -71205
         TabIndex        =   40
         Top             =   4605
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -70785
         TabIndex        =   39
         Top             =   3660
         Width           =   15
      End
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
      Left            =   8610
      TabIndex        =   144
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Entreg:"
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
      Left            =   8505
      TabIndex        =   142
      Top             =   525
      Width           =   690
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   8505
      TabIndex        =   139
      Top             =   105
      Width           =   690
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Loc.:"
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
      Left            =   5775
      TabIndex        =   136
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "C.P.:"
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
      Left            =   4095
      TabIndex        =   134
      Top             =   840
      Width           =   1110
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
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
      Height          =   225
      Left            =   4095
      TabIndex        =   132
      Top             =   525
      Width           =   1110
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Cliente:"
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
      Left            =   4095
      TabIndex        =   125
      Top             =   0
      Width           =   1110
   End
   Begin VB.Line Line1 
      X1              =   -420
      X2              =   11580
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
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
      Left            =   -210
      TabIndex        =   38
      Top             =   735
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de E.V.:"
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
      Left            =   0
      TabIndex        =   37
      Top             =   135
      Width           =   1140
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu NueCli 
         Caption         =   "Nuevo Cliente"
      End
      Begin VB.Menu EditCli 
         Caption         =   "Abrir Existente"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu GuaCli 
         Caption         =   "Guardar"
      End
      Begin VB.Menu php1a 
         Caption         =   "-"
      End
      Begin VB.Menu Baj_Cli 
         Caption         =   "Baja Cuenta Activa"
      End
      Begin VB.Menu Susp_Cli 
         Caption         =   "Suspender Crédito"
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
   Begin VB.Menu EdiCli 
      Caption         =   "Edición"
      Begin VB.Menu DeHaCaCli 
         Caption         =   "Deshacer Cambios"
      End
      Begin VB.Menu php5a 
         Caption         =   "-"
      End
      Begin VB.Menu SuspOper 
         Caption         =   "Suspender Operación"
      End
      Begin VB.Menu php5b 
         Caption         =   "-"
      End
      Begin VB.Menu BusCli 
         Caption         =   "Buscar"
      End
   End
   Begin VB.Menu ConLiCli 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu LiGeClaCli 
         Caption         =   "Listado General y Clasificado"
      End
      Begin VB.Menu LiCuInCli 
         Caption         =   "Listado de Cuentas Inactivas"
      End
      Begin VB.Menu DeLiCreCli 
         Caption         =   "Deudas y Límites de Crédito"
      End
   End
   Begin VB.Menu ConfCli 
      Caption         =   "Configuración"
      Begin VB.Menu ConGenCli 
         Caption         =   "Configuración General"
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu ConPagCli 
         Caption         =   "Condiciones de Pago"
      End
      Begin VB.Menu LiPreCli 
         Caption         =   "Listas de Precios"
      End
      Begin VB.Menu VenCli 
         Caption         =   "Vendedores"
      End
      Begin VB.Menu PrZoVeCli 
         Caption         =   "Provincias y Zonas de Venta"
      End
      Begin VB.Menu GruCli 
         Caption         =   "Grupo de Clientes"
      End
      Begin VB.Menu CuCoCli 
         Caption         =   "Cuentas Contables"
      End
      Begin VB.Menu php6 
         Caption         =   "-"
      End
      Begin VB.Menu OrDBCli 
         Caption         =   "Ordenar Base de Datos"
      End
   End
   Begin VB.Menu AyuCli 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "AORTRA_GES04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%(8), MODIFIC%(8)
Dim GRABATODO%
Dim AGRECLI%
Dim RECONTANT$

Dim UPDATCLI%

Private Sub AyuCli_Click()
Call Command4_Click
End Sub

Private Sub Baj_Cli_Click()
   Call Command24_Click
End Sub

Private Sub BusCli_Click()
Call Command20_Click
End Sub

Private Sub CoInBDCli_Click()
Call Command99_Click
End Sub

Private Sub Command11_Click()
   '
   If Val(TRIM$(Text1)) > 0 Then
     Call SELECHO("DEUDOR1")
     NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
     If NCLIE% < 1 Then Exit Sub
     '
     NC% = NCLIE%
     Text2 = TRIM$(Str$(NC%))
     Text4 = RASOCLI$(NC%)
   End If
   '
End Sub

Private Sub Command22_Click()
   Static CTXX%
   Command22.Enabled = False
   '
   If ULTREG&("DAPIBCLI") < 1 Then
     Call GENDAPIBCLI
   End If
   '
   NCLIE% = Val(TRIM$(Text1.TEXT))
   If NCLIE% > 0 Then
     If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
       '
       If CTXX% < 1 Then
         Call GRAREG("TACATIBC", Chr$(0) + "Exento Percepcion IB", 1)
         Call GRAREG("TACATIBC", Chr$(1) + "Sujeto Percepcion IB", 2)
         Call SETULTREG("TACATIBC", 2)
         CTXX% = 1
       End If
       '
       RECORD$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE%))
       Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
       CATIB1% = Asc(Mid$(RECORD$, 3, 1))
       If CATIB1% < 0 Or CATIB1% > 1 Then
         Call REPLA(RECORD$, Chr$(1), 3, 1)
         Call REPLA(RECORD$, Space$(32), 9, 32)
       End If
       '
10     OBX$ = OBJPLA$("CARDAPERIB", RECORD$)
       If OBX$ <> "" Then
         RECORD$ = OBX$
         CATIB1% = Asc(Mid$(RECORD$, 3, 1))
         If CATIB1% < 0 Or CATIB1% > 1 Then
           Call MENSERR(24, "Mal Categoría I.Brutos")
           GoTo 10
         End If
         If CATIB1% = 0 Then
           If TRIM$(Mid$(RECORD$, 9, 32)) = "" Then
             Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
             GoTo 10
           End If
         End If
         If CATIB1% = 1 Then
           Call REPLA(RECORD$, Space$(32), 9, 32)
         End If
         PORCONVE = CVS(Mid$(RECORD$, 5, 4))
         If PORCONVE < 0 Or PORCONVE > 100 Then
           Call MENSERR(24, "% Convenio Multilateral\Fuera de Rango")
           GoTo 10
         End If
         '
         Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECORD$, USERNAME$, 41, 20)
         Call REPLA(RECORD$, MKI$(HOY(HOS$)), 61, 2)
         Call REPLA(RECORD$, MKI$(HORANUM%(Time$)), 63, 2)
         Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIE%), RECORD$)
         Call CIERRARCH("DAPIBCLI")
         '
       End If
     End If
   End If
   Command22.Enabled = True
End Sub

Private Sub Command4_Click()
   Call HELPONLINE("AMACLI")
End Sub

Private Sub Command99_Click()
   Command99.Enabled = False
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("DEUDOR1")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     R0$ = REGLEIDO$("DEUDOR1", U&)
     NCLIE% = CVI(Left$(R0$, 2))
     R1$ = Mid$(R0$, 3, 30)
     If Left$(R1$, 1) = Chr$(64) Then
       R1$ = AJUSTI$(TRIM$(Mid$(R1$, 2)), 30)
     End If
     R1$ = REPLACAR$(R1$, "@", " ")
     R2$ = AJUSTI$(RASOCLI$(NCLIE%), 30)
     If R1$ <> R2$ Then
       HAYINCO% = HAYINCO% + 1
       Call MENSERR(24, "Revisar Cuenta " + TRIM$(Str$(NCLIE%)))
     End If
   Next U&
   '
   If HAYINCO% < 1 Then
     Call COMUNI("Control Completo.\No se Detectaron Errores.")
   End If
   '
   Screen.MousePointer = 1
   Command99.Enabled = True
End Sub

Private Sub ConGenCli_Click()
Call Command14_Click
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

Private Sub DeHaCaCli_Click()
Call Command2_Click
End Sub

Private Sub DeLiCreCli_Click()
    Call AMACLIS04.Command3_Click
End Sub

Private Sub EditCli_Click()
   Call Command9_Click
End Sub

Private Sub Exit_Click()
   Call Command3_Click
End Sub

'
Private Sub Form_Load()
    '
    SSTab1.Tab = 0
    Call CENTRAFORM(Me)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    ARCHEX$(0) = "PROVINC"
    ARCHEX$(1) = "CAPOSIV"
    ARCHEX$(2) = "CARETIB"
    ARCHEX$(3) = "LISTAS"
    ARCHEX$(4) = "CONPAG"
    ARCHEX$(5) = "VENDEDOR"
    ARCHEX$(6) = "ECUECON"
    ARCHEX$(7) = "GRUCOCLI"
    '
    Call CARTAVALI
    '
    If ULTREG&("BKDATCLI") < 1 Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + "DEUDOR1.BAK") > 0 Then
        If EXISTE%(LU$ + "DEUDOR2.BAK") > 0 Then
          Screen.MousePointer = 11
          JJ& = 0
          Dim D1 As String * 32
          Dim D2 As String * 128
          N1% = FreeFile
          Open LU$ + "DEUDOR1.BAK" For Random Access Read Write Shared As #N1% Len = 32
          N2% = FreeFile
          Open LU$ + "DEUDOR2.BAK" For Random Access Read Write Shared As #N2% Len = 128
          For K& = 1 To LOF(N1%) / 32
            Get #N1%, K&, D1$
            NC% = CVI(Left$(D1$, 2))
            If NC% > 0 Then
              Get #N2%, K&, D2$
              D3$ = D1$ + D2$
              JJ& = JJ& + 1
              Call GRAREG("BKDATCLI", D3$, JJ&)
            End If
          Next K&
          Call SETULTREG("BKDATCLI", JJ&)
          Call CIERRARCH("BKDATCLI")
          Screen.MousePointer = 1
        End If
      End If
    End If
    '
End Sub
'
Private Sub Form_Activate()
   WindowState = 0
   openforms = DoEvents
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
     openforms = DoEvents
   Wend
End Sub

Private Sub BOTSEL_Click(Index As Integer)
   Call SELECHO(ARCHEX$(Index))
   VDEVU$ = VALACT1$(ARCHEX$(Index))
   If TRIM$(VDEVU$) <> "" Then
      COTEXT(Index).TEXT = VDEVU$
   End If
End Sub

Private Sub Command1_Click()
   GRABATODO% = 1
   Call GRAGENCLI
   Call GRANOTACLI
   GRABATODO% = 0
End Sub

Private Sub Command13_Click()
   AMACLIS04.Show 1
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   OPMASCL04.Show 1
   Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   Command15.Enabled = False
   Call MENSERR(24, "No Disponible en la Presente Versión.\ \Los Ordenamientos son Automáticos\al Generar los Listados.")
   Command15.Enabled = True
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command2_Click()
   Call GRAGENCLI
   Call GRANOTACLI
   For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call CARDATCLIE
End Sub

Private Sub Command20_Click()
   Static ARGU$
   Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     Call COPYSTRU("DEUDOR1", "SELECLI")
     '
     Call ABRECLIEN
     On Error Resume Next
     With Deudor12
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           For KKI% = 1 To .Fields.Count
             VATEX$ = .Fields(KKI% - 1).Value
             If InStr(VATEX$, ARGU$) > 0 Then
               NCLIE% = !Nume_Cli
               RACLI$ = !Raso_Cli
               XX$ = MKI$(NCLIE%) + RACLI$
               JJ& = JJ& + 1
               Call GRAREG("!SELECLI", XX$, JJ&)
               Exit For
             End If
           Next KKI%
         .MoveNext
         Loop
       End If
     End With
     Call SETULTREG("!SELECLI", JJ&)
     '
     Screen.MousePointer = 1
     If JJ& < 1 Then
         Call MENSERR(24, "No se Encontraron Coincidencias")
         GoTo 10
       ElseIf JJ& = 1 Then
         Text1 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELECLI", 1), 2))))
       Else
         Call FRESHECHO("!SELECLI")
         Call SELECHO("!SELECLI")
         NCLIX$ = VALACT1$("!SELECLI")
         If NCLIX$ <> "" Then
           Text1 = NCLIX$
         End If
     End If
   End If
   Command20.Enabled = True
End Sub

Private Sub Command21_Click()
   '
   If DERACCE%("ALTACLI", "Alta de Nuevo Cliente") < 2 Then
      Exit Sub
   End If
   '
   Call ABRECLIEN
   '
   Command21.Enabled = False
   Call GRAGENCLI
   Call GRANOTACLI
   '
   Screen.MousePointer = 11
   PRONUCLI& = Val(CONTROL$("", "PRONUCLI"))
   If PRONUCLI& > 9999 Then PRONUCLI& = 0
   NROPRO& = PRONUCLI&
   If NROPRO& < 0 Then NROPRO& = 0
   '
   Set Tranclie = Badaclie.OpenRecordset("SELECT Nume_Cli FROM Deudor12 WHERE Nume_Cli >= " & PRONUCLI& & " ORDER BY Nume_Cli DESC;", dbOpenSnapshot)
   On Error Resume Next
   '
3  NROPRO& = NROPRO& + 1
   Tranclie.FindFirst "Nume_Cli = " & NROPRO&
   If Tranclie.NoMatch = False Then GoTo 3
   '
   NCLIE% = NROPRO&
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(NCLIE%), 13, 2)
   Call REPLA(VDEF$, Space$(62), 15, 62)
   Call REPLA(VDEF$, Chr$(1), 77, 1)
   Call REPLA(VDEF$, Space$(15), 78, 15)
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("ALTACLIENACC", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   '
   COCLINU$ = TRIM$(Left$(OBX$, 12))
   If COCLINU$ = "" Then
     Call MENSERR(24, "Codigo de Cuenta no Válido")
     GoTo 10
   End If
   '
   NCLINUE% = CVI(Mid$(OBX$, 13, 2))
   If NCLINUE% < 1 Or NCLINUE% > 9999 Then
      Call MENSERR(24, "Número de Cuenta no Válido.")
      GoTo 10
   End If
   '
   Set Tranclie = Badaclie.OpenRecordset("SELECT Codi_Cli FROM Deudor12;", dbOpenSnapshot)
   On Error Resume Next
   Tranclie.FindFirst "Codi_Cli = '" & COCLINU$ & "'"
   If Tranclie.NoMatch = False Then
      Call MENSERR(24, "Ingreso no Válido.\Código de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   Set Tranclie = Badaclie.OpenRecordset("SELECT Nume_Cli FROM Deudor12;", dbOpenSnapshot)
   On Error Resume Next
   Tranclie.FindFirst "Nume_Cli = " & NCLINUE
   If Tranclie.NoMatch = False Then
      Call MENSERR(24, "Ingreso no Válido.\Número de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   DCLINUE$ = TRIM$(Mid$(OBX$, 15, 62))
   If Len(DCLINUE$) < 1 Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Razón Social del Cliente")
      GoTo 10
   End If
   '
   PIVCLI% = Asc(Mid$(OBX$, 77, 1))
   If PIVCLI% < 0 Or PIVCLI% > 6 Then
      Call MENSERR(24, "Ingreso no Válido.\Posición IVA no Figura en Tablas.")
      GoTo 10
   End If
   '
   CUTT$ = TRIM$(Mid$(OBX$, 78, 15))
   If VALICUIT%(CUTT$) <> 1 Then
     If PIVCLI% <> 0 Then
       If PIVCLI% <> 5 Then
         Call MENSERR(24, "Ingreso no Válido.\Número de C.U.I.T. Incorrecto.")
         GoTo 10
       End If
     End If
   End If
   '
   If CUTT$ <> "" Then
     Set Tranclie = Badaclie.OpenRecordset("SELECT Nume_Cli FROM Deudor12 WHERE Cuit_Cli = '" & CUTT$ & "';", dbOpenSnapshot)
     CAREI% = 0
     On Error Resume Next
     Tranclie.MoveLast
     CAREI% = Tranclie.RecordCount
     '
     If CAREI% > 0 Then
       If Tranclie!Nume_Cli > 0 Then
         NCLIX% = Tranclie!Nume_Cli
         If COMALTER%("Existe Registrada una Cuenta con el Mismo C.U.I.T.\  \'" + TRIM$(RASOCLI$(NCLIX%)) + "' (" + TRIM$(Str$(NCLIX%)) + ")\\Cancelar\Registrar Igual") <> 2 Then
           GoTo 10
         End If
       End If
     End If
   End If
   '
   Screen.MousePointer = 11
20 Deudor12.Close
   On Error Resume Next
   Call ABRECLIEN
   Deudor12.MoveFirst
   If Err < 1 Then
     On Error GoTo 0
     ABUS$ = "Nume_Cli = " & NCLINUE%
     Deudor12.FindFirst ABUS$
     If Deudor12.NoMatch = False Then
       NCLINUE% = NCLINUE% + 1
       GoTo 20
     End If
   End If
   '
   With Deudor12
      .AddNew
      !Nume_Cli = TRIM$(Str$(NCLINUE%))
      !Raso_Cli = DCLINUE$
      !Codi_Cli = COCLINU$
      !Cuit_Cli = CUTT$
      !Piva_Cli = PIVCLI%
      !Stat_Cli = 1
      '
      On Error Resume Next
      .Update
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
      End If
      On Error GoTo 0
      .Bookmark = .LastModified
      .Close
      UPDATCLI% = 1
      Call GENECOCLI
   End With
   '
   AGRECLI% = 1
   Text1.TEXT = TRIM$(Str$(NCLINUE%))
   SSTab1.Tab = 0
   Call CARDATCLIE
   Call CARTAVALI
   '
   For Index = 0 To 7
     If Index <> 1 Then
       VAPRE$ = TRIM$(CONTROL$("AMACLI", ARCHE$(Index)))
       COTEXT(Index) = "": DETEXT(Index) = ""
       If VAPRE$ <> "" Then
         TV11% = TIPFIELD(ARCHE$(Index), 1)
         LF11% = LENFIELD(ARCHE$(Index), 1)
         ARGU$ = CBIN$(VAPRE$, TV11%, LF11%)
         ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
         If ECO$ <> "" Then
             COTEXT(Index).TEXT = VAPRE$
             DETEXT(Index).Caption = ECO$
           Else
             COTEXT(Index).TEXT = ""
             DETEXT(Index).Caption = ""
         End If
       End If
     End If
   Next Index
   '
   Text3 = DCLINUE$
   COTEXT(1) = TRIM$(Str$(PIVCLI%))
   Text33 = TRIM$(CUTT$)
   '
   Screen.MousePointer = 1
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command23_Click()
   SSTab1.Tab = 0
   Text20.TEXT = " Suspendido"
   Text20.Refresh
End Sub

Private Sub Command24_Click()
   '
   Command24.Enabled = False
   Screen.MousePointer = 11
   '
   Call GRAGENCLI
   Call GRANOTACLI
   '
   If DERACCE%("BAJACLI", "Baja de Cuenta de Cliente") < 2 Then
     GoTo 99
   End If
   '
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE& = Val(Text1.TEXT)
       Call ABRECLIEN
       With Deudor12
       On Error Resume Next
       .MoveFirst
       .FindFirst "Nume_Cli = " & NCLIE&
       If .NoMatch = False Then
         '
         On Error GoTo 0
         NC% = NCLIE
         '
         If Abs(SALDCLI#(NC%)) >= 0.005 Then
           Call MENSERR(24, "Baja Imposible - Cuenta con Saldo.")
           GoTo 99
         End If
         '
         RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
         If RFF$ <> "" Then
           Call MENSERR(24, "Baja Imposible - Cuenta con Movimiento.")
           GoTo 99
         End If
         '
         RFF$ = RECFILT$("COBRAN", 1, MKI$(NC%))
         If RFF$ <> "" Then
           Call MENSERR(24, "Baja Imposible - Cuenta con Movimiento.")
           GoTo 99
         End If
         
         TTXX$ = "La Baja de un Cliente\Puede Producir Inconsistencias en la Base de Datos\en Caso de que la Cuenta Registre\Facturación, Cobranza o Pedidos."
         TTXX$ = TTXX$ + "\  \Puede Ahora Cancelar la Operación\o Confirmar la Baja de la Cuenta:\'" + TRIM$(Text3.TEXT) + "'"
         If COMALTER%(TTXX$ + "\\Cancelar Operación\Confirmar Baja") = 2 Then
           '
           .Edit
           !Stat_Cli = (-1)
           .Update
           '
           YACAR%(1) = 1
           Text28.TEXT = ""
           MODIFIC%(1) = 1
           GRABATODO% = 1
           Call GRANOTACLI
           '
           UPDATCLI% = 1
           Call GENECOCLI
           Text1.TEXT = ""
           Call BLANFORMU
           For KKI% = 0 To 4: MODIFIC%(KKI%) = 0: YACAR%(KKI%) = 0: Next KKI%
           GRABATODO% = 0
           Call BAJALDISCO
           '
         End If
         '
       End If
       End With
     End If
   End If
   '
99 Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Private Sub Command25_Click()
   PRINTFORM
End Sub

Private Sub Command3_Click()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   Call CARDATCLIE
   '
   Call CIERRARCH("*.*")
   If AGRECLI% > 0 Then
      Call SORTCLI(1)
      AGRECLI% = 0
   End If
   Call BAJALDISCO
   '
   Hide
   'If APLINVO$ <> "" Then
   '    Call FINAL("")
   '  Else
   '    Hide
   'End If
   '
End Sub

Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub

Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub

Private Sub Command9_Click()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   '
   Call SELECHO("DEUDOR1")
   NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% < 1 Then Exit Sub
   '
   NC% = NCLIE%
   Text1.TEXT = TRIM$(Str$(NC%))
   'Call CARDATCLIE
   '
End Sub

Private Sub LiCuInCli_Click()
    Call AMACLIS04.Command5_Click
End Sub

Private Sub LiGeClaCli_Click()
 Call AMACLIS04.Command8_Click
End Sub

Private Sub SuspOper_Click()
    Call Command3_Click
End Sub

Private Sub Text2_Change()
   Text4 = ""
   If Val(TRIM$(Text1)) > 0 Then
     NCLIE% = Val(TRIM$(Text2))
     If NCLIE% > 0 Then
       Text4 = RASOCLI(NCLIE%)
     End If
   End If
End Sub

Private Sub Text2_DblClick()
   Call Command11_Click
End Sub

Private Sub Text6_Change()
   MODIFIC%(0) = 1
   If NUMECLI%(Text6) > 0 Then
     Text1 = TRIM$(Str$(NUMECLI%(Text6)))
   End If
End Sub

Private Sub Text6_DblClick()
   Call Command10_Click
End Sub

Private Sub COTEXT_Change(Index As Integer)
   '
   ECO$ = ""
   If TRIM$(COTEXT(Index).TEXT) <> "" Then
      TV11% = TIPFIELD(ARCHEX$(Index), 1)
      LF11% = LENFIELD(ARCHEX$(Index), 1)
      ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
      ECO$ = TRIM$(ECOARCH$(ARCHEX$(Index), ARGU$))
   End If
   '
   If ECO$ <> "" Then
        DETEXT(Index).Caption = ECO$
      Else
        DETEXT(Index).Caption = "No Definido"
   End If
   '
   MODIFIC%(0) = 1
End Sub

Private Sub COTEXT_DblClick(Index As Integer)
   Call SELECHO(ARCHEX$(Index))
   VDEVU$ = VALACT1$(ARCHEX$(Index))
   If TRIM$(VDEVU$) <> "" Then
      COTEXT(Index).TEXT = VDEVU$
   End If
End Sub

Private Sub Command10_Click()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   '
   Screen.MousePointer = 1
   RSLX$ = REGSEL$("INDICLIE")
   If Len(RSLX$) > 4 Then
     NCLIE% = CVI(Mid$(RSLX$, 63, 2))
     If NCLIE% > 0 Then
       NC% = NCLIE%
       Text1.TEXT = TRIM$(Str$(NC%))
     End If
   End If
   '
End Sub

Sub CARDATCLIE()
    '
    NCLIE% = Val(TRIM$(Text1.TEXT))
    STCL% = STATCLI%(NCLIE)
    If NCLIE < 1 Or Abs(STCL%) < 1 Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text3.TEXT = RASOCLI$(NCLIE)
    Text24.TEXT = TEL2CLI$(NCLIE)
    Text23.TEXT = NPROCLI$(NCLIE)
    MODIFIC(0) = 0
    '
    Select Case SSTab1.Tab
      Case 0
        Call CARGENCLI
      Case 1
        Call CARNOTACLI
      Case 2
     '   Call CARESCUENTA
    End Select
    Screen.MousePointer = 1
    '
End Sub
'
Sub CARGENCLI()
    '
    If YACAR%(0) > 0 Then Exit Sub
    NC% = Val(TRIM$(Text1.TEXT))
    '
    Text6 = CODICLI$(NC%)
    Text25.TEXT = DOMICLI$(NC%)
    Text25.Refresh
    Text26.TEXT = CPOSCLI$(NC%)
    Text26.Refresh
    Text27.TEXT = LOCACLI$(NC%)
    Text27.Refresh
    If LIMICRE(NC%) < 0 Then
        Text20.TEXT = " Suspendido"
      Else
        Text20.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 12, 2, 2)
    End If
    Text20.Refresh
    Text23.TEXT = NPROCLI$(NC%)
    Text23.Refresh
    Text29.TEXT = PAISCLI$(NC%)
    Text29.Refresh
    Text30.TEXT = TEL1CLI$(NC%)
    Text30.Refresh
    Text31.TEXT = FAXCLI$(NC%)
    Text31.Refresh
    Text32.TEXT = EMAILCLI$(NC%)
    Text32.Refresh
    Text19.TEXT = CONTACTCLI$(NC%)
    Text19.Refresh
    Text36.TEXT = HORARCLI$(NC%)
    Text36.Refresh
    Text33.TEXT = CUITCLI$(NC%)
    Text33.Refresh
    Text34.TEXT = NIBRCLI$(NC%)
    Text34.Refresh
    Text21.TEXT = TRIM$(DESCUCLI$(NC%))
    Text21.Refresh
    '
    If ULTREG&("CAPOSIV") < 8 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(0) + AJUSTI$("IVA Consumidor Final            ", 31) + AJUSTI$("Cons.Fin", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(1) + AJUSTI$("IVA Responsable Inscripto   ", 31) + AJUSTI$("Resp.Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(2) + AJUSTI$("IVA Responsable no Inscripto", 31) + AJUSTI$("R.No Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(3) + AJUSTI$("IVA No Responsable          ", 31) + AJUSTI$("No Resp.", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(4) + AJUSTI$("IVA Responsable Exento      ", 31) + AJUSTI$("Exento  ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Operac.Comercio Exterior", 31) + AJUSTI$("Com.Ext. ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(6) + AJUSTI$("IVA Contrib.Monotributo     ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(7) + AJUSTI$("IVA Exento Zona Promovida   ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      Call SETULTREG("CAPOSIV", IQ&)
      Call CIERRARCH("CAPOSIV")
    End If
    '
    COTEXT(0).TEXT = TRIM$(PROVCLI$(NC%))
    COTEXT(0).Refresh
    COTEXT(1).TEXT = TRIM$(Str$(PIVACLI%(NC%)))
    COTEXT(1).Refresh
    COTEXT(2).TEXT = TRIM$(CATIBRU$(NC%))
    COTEXT(2).Refresh
    COTEXT(3).TEXT = TRIM$(Str$(LIPRCLI%(NC%)))
    COTEXT(3).Refresh
    COTEXT(4).TEXT = TRIM$(Str$(CPAGCLI%(NC%)))
    COTEXT(4).Refresh
    COTEXT(5).TEXT = TRIM$(Str$(VENDCLI%(NC%)))
    COTEXT(5).Refresh
    COTEXT(6).TEXT = TRIM$(CUEMADRE$(NC%))
    COTEXT(6).Refresh
    COTEXT(7).TEXT = TRIM$(Str$(GRUCOCLI%(NC%)))
    COTEXT(7).Refresh
    '
    'grupo divisional
    GRUDI% = GRUDIVCLI%(NC%)
    Text2 = ""
    If GRUDI% <> NC% Then Text2 = TRIM$(Str$(GRUDI%))
    '
    YACAR%(0) = 1: MODIFIC%(0) = 0
    '
End Sub
'
Sub GRAGENCLI()
    '
    NCLIE& = Val(TRIM$(Text1.TEXT))
    If NCLIE& < 1 Then
      If GRABATODO% = 1 Then
        Call MENSERR(24, "Imposible Registrar !!!\Número de Cliente no Válido.")
      End If
      Exit Sub
    End If
    '
    If Abs(STATCLI%(NCLIE)) < 1 Then
      If GRABATODO% = 1 Then
        Call MENSERR(24, "Imposible Registrar !!!\Numero de Cliente Inexistente")
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    If STATCLI(NCLIE&) < 1 Then
      If GRABATODO% = 1 Then
        If COMALTER%("El Cliente Elegido Figura\Marcado Como 'Dado de Baja'.\  \Desea Re-Activar Esta Cuenta ?\\No, Dejar de Baja\Si, Re-Activar") = 2 Then
          Call ABRECLIEN
          With Deudor12
            On Error Resume Next
            .FindFirst "Nume_Cli = " & Abs(NCLIE&)
            If .NoMatch = False Then
             .Edit
             !Stat_Cli = 1
             .Update
            End If
          End With
          UPDATCLI% = 1
          Call GENECOCLI
          Call CARGENCLI
        End If
      End If
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    If MODIFIC%(0) < 1 Then
      If GRABATODO% <> 1 Then
        Exit Sub
      End If
    End If
    '
    If GRABATODO% = 1 Then
      If SSTab1.Tab <> 0 Then
        Exit Sub
      End If
    End If
    '
    If DERACCE%("MODATCLI", "Modificar Datos de Clientes") < 2 Then
       Exit Sub
    End If
    '
    If GRABATODO% <> 1 Then
      Command9.SetFocus
      If COMALTER%("Se han Modificado Datos del Cliente\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
        YACAR%(0) = 0
        Screen.MousePointer = 11
        Text3.TEXT = RASOCLI$(NCLIE)
        Text24.TEXT = TEL2CLI$(NCLIE)
        MODIFIC(0) = 0
        Call CARGENCLI
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    Screen.MousePointer = 11
    '
    CUIT$ = TRIM$(Text33.TEXT)
    If CUIT$ <> "" Then
      If VALICUIT%(CUIT$) <> 1 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Número de C.U.I.T. no Válido")
        SSTab1.Tab = 0
        Text33.SetFocus
        Exit Sub
      End If
    End If
    '
    If NCLIE = 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar Cuenta Número '0'")
        Exit Sub
    End If
    '
    If TRIM$(Text6) = "" Then
      Call MENSERR(24, "Imposible Registrar !!!\Falta Código de Cuenta")
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    NPRII% = NUMECLI%(Text6)
    If NPRII% > 0 Then
      If NPRII% <> Abs(NCLIE) Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Codigo de Cuenta Repetido o no Válido.")
        Exit Sub
      End If
    End If
    '
    If TRIM$(Text3.TEXT) = "" Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Falta Razón Social del Cliente.")
        Exit Sub
    End If
    '
    '
    Call ABRECLIEN
    With Deudor12
      On Error Resume Next
      .MoveFirst
      .FindFirst "Nume_Cli = " & NCLIE&
      If .NoMatch = False Then GoTo 10
      On Error GoTo 0
      Call MENSERR(24, "Imposible Registrar")
      Exit Sub
    End With
    '
10  On Error GoTo 0
    With Deudor12
      On Error Resume Next
      ABC1$ = !Codi_Cli
      If Err Then
        ABC1$ = ""
      End If
      On Error GoTo 0
      .Edit
        If TRIM$(Text3) <> TRIM$(!Raso_Cli) Or UCase$(Text6) <> ABC1$ Then UPDATCLI% = 1
      !Codi_Cli = UCase$(Text6)
      !Raso_Cli = Text3
      !Stat_Cli = 1
      !Domi_Cli = Text25
      !Cpos_Cli = Text26
      !Loca_Cli = Text27
      !Tele_Cli = Text30
      !Fant_Cli = Text24
      !Cuit_Cli = TRIM$(CUIT$)
      !Nibr_Cli = TRIM$(Text34)
      !Caib_Cli = TRIM$(COTEXT(2))
      !Prov_Cli = TRIM$(COTEXT(0))
      '
      MODLIMCRE% = 0
      On Error Resume Next
      LMCRA = CDbl(!Lcre_Cli)
      If Err Then
        LMCRA = 0: !Lcre_Cli = 0
      End If
      On Error GoTo 0
      '
      If Abs(CDbl(!Lcre_Cli) - Val(Text20.TEXT)) > 0.05 Or InStr(UCase$(Text20.TEXT), "SUSPEN") > 0 Then
        If DERACCE%("MODLICRE", "MODIFICACION LIMITE DE CREDITO") >= 1 Then
            !Lcre_Cli = Val(Text20)
            If TRIM$(UCase$(Text20.TEXT)) = "SUSPENDIDO" Then
              !Lcre_Cli = (-1)
            End If
          Else
           Text20.TEXT = CSTRING$(MKS$(!Lmcr_Cli), 3, 12, 2, 2)
           If LMCR < 0 Then Text20.TEXT = " Suspendido"
           Text20.Refresh
           MODLIMCRE% = 1
        End If
      End If
      '
      PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
      !Piva_Cli = PIVA%
      '
      CATIB% = VALIMI(COTEXT(2).TEXT, 0, 2)
      !Caib_Cli = CATIB%
      '
      LIPRE% = VALIMI(COTEXT(3).TEXT, 0, 255)
      If LIPRE% > 0 Then
        If ECOARCH("LISTAS", Chr$(LIPRE%)) = "" Then
          Call MENSERR(24, "Lista de Precios no Válida")
          LIPRE% = 0
        End If
      End If
      !Lpre_Cli = LIPRE%
      '
      COPAGI% = VALIMI(COTEXT(4).TEXT, 0, 255)
      If COPAGI% > 0 Then
        If ECOARCH("CONPAG", Chr$(COPAGI%)) = "" Then
          Call MENSERR(24, "Condición de Pago no Válida")
          COPAGI% = 0
        End If
      End If
      !Cpag_Cli = COPAGI%
      '
      VENDE% = VALIMI(COTEXT(5).TEXT, 0, 255)
      If VENDE% > 0 Then
        If ECOARCH("VENDEDOR", Chr$(VENDE%)) = "" Then
          Call MENSERR(24, "Vendedor no Válido")
          VENDE% = 0
        End If
      End If
      !Vend_Cli = VENDE%
      '
      CUECO% = VALIMI(COTEXT(7).TEXT, 0, 32767)
      If ECOARCH$("GRUCOCLI", MKI$(CUECO%)) = "" Then
        CUECO% = 0
      End If
      !GrCo_Cli = CUECO%
      !Gru_Div = Val(Text2)
      '
      CCX% = REGBUS&("CUENTAS", 1, AJUSTI$(COTEXT(6).TEXT, 12))
      !Cuma_Cli = CCX%
      '
      DCC$ = TRIM$(Text21.TEXT)
      !Dsta_Cli = TRIM$(Text21)
      '
      !Pais_Cli = Text29
      !Faxi_Cli = Text31
      !Mail_Cli = Text32
      !Ctac_Cli = Text19
      !Prvd_Cli = Text23
      !Hora_Cli = Text36
      '
      .Update
    End With
    '
    Deudor12.Close
    Call GENECOCLI
    '
    MODIFIC%(0) = 0
    Screen.MousePointer = 1
    '
End Sub
'
Sub CARNOTACLI()
   '
   If YACAR%(1) > 0 Then Exit Sub
   NCLIE% = Val(Text1.TEXT)
   Text28 = ""
   '
   If NCLIE% > 0 Then
     If REGICLI(NCLIE%) > 0 Then
       Call ABRECLIEN
       Set ClienTemp = Badaclie.OpenRecordset("SELECT * FROM Deudor12 WHERE Nume_Cli = " & NCLIE%, dbOpenSnapshot)
       '
       On Error Resume Next '
       ClienTemp.MoveFirst
       If Err < 1 Then
         Text28 = ClienTemp!Nota_Cli
       End If
     End If
   End If
   Text28.Refresh
   YACAR%(1) = 1: MODIFIC%(1) = 0
   '
End Sub
'
Sub GRANOTACLI()
   '
   NCLIE% = Val(TRIM$(Text1.TEXT))
   If NCLIE% < 1 Then Exit Sub
   '
   If REGICLI(NCLIE%) < 1 Then Exit Sub
   If YACAR%(1) < 1 Then Exit Sub
   If MODIFIC%(1) < 1 Then Exit Sub
   '
   If DERACCE%("ANOTACLI", "Anotaciones de de Clientes") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     Command9.SetFocus
     If COMALTER%("Se han Modificado las Notas del Cliente\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(1) = 0
       Call CARNOTACLI
       Exit Sub
     End If
     Command9.SetFocus
   End If
   '
   Call ABRECLIEN
   Screen.MousePointer = 11
   TEXTO$ = Text28.TEXT + Chr$(13) + Chr$(10)
   '
   With Deudor12
     On Error Resume Next
     .MoveFirst
     .FindFirst "Nume_Cli = " & Abs(NCLIE%)
     If .NoMatch = False Then GoTo 10
     On Error GoTo 0
     Call MENSERR(24, "Imposible Registrar")
     Screen.MousePointer = 1
     Exit Sub
   End With
   '
10 On Error GoTo 0
   With Deudor12
     .Edit
     !Nota_Cli = TEXTO$
     .Update
   End With
   Deudor12.Close
   MODIFIC%(1) = 0
   Screen.MousePointer = 1
   '
End Sub
'
Sub BLANFORMU()
   '
   Text2 = ""
   Text3 = ""
   Text24 = ""
   Text25.TEXT = ""
   Text26.TEXT = ""
   Text27.TEXT = ""
   Text20.TEXT = ""
   Text29.TEXT = ""
   Text30.TEXT = ""
   Text31.TEXT = ""
   Text32.TEXT = ""
   Text19.TEXT = ""
   Text33.TEXT = ""
   Text34.TEXT = ""
   Text36.TEXT = ""
   Text21.TEXT = ""
   '
   For KK% = 0 To 7
     COTEXT(KK%).TEXT = ""
   Next KK%
   '
   Text28.TEXT = ""
   '
   Text23 = ""
   Text35 = ""
   Text36 = ""
   Text37 = ""
   List2.Clear
   List3.Clear
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
   Call GRAGENCLI
   Call GRANOTACLI
   Call CARDATCLIE
   '
   Call CIERRARCH("*.*")
   If AGRECLI% > 0 Then
      Call SORTCLI(1)
      AGRECLI% = 0
   End If
   Call BAJALDISCO
   '
   Hide
'   If APLINVO$ <> "" Then
'       Call FINAL("")
'     Else
'       Hide
'   End If
   '
End Sub

Private Sub GruCli_Click()
With OPMASCL04
  .Option1(9).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub GuaCli_Click()
Call Command1_Click
End Sub

Private Sub LiPreCli_Click()
With OPMASCL04
  .Option1(7).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub List_Click()
   Call Command13_Click
End Sub

Private Sub List3_DblClick()
   '
   VUELBUS% = 1
   '
3  RELIS$ = List3.TEXT
   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 16))
   If TIDOCUM$ = "Ajuste de Saldo" Then TIDOCUM$ = "Ajuste Saldo Cliente"
   NUDOCU$ = " " + TRIM$(Mid$(RELIS$, 27, 8)) + " "
   If InStr(UCase$(TIDOCUM$), "FACTURA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Then
     NUDOCU$ = TRIM$(NUDOCU$)
     While Len(NUDOCU$) < 8: NUDOCU$ = "0" + NUDOCU$: Wend
   End If
   RACLI$ = TRIM$(Left$(RASOCLI$(Val(Text1.TEXT)), 8))
   Screen.MousePointer = 11
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
     If InStr(XX$, TIDOCUM$) > 4 Then
       If InStr(XX$, NUDOCU$) > 4 Then
         If InStr(XX$, RACLI$) > 16 Or VUELBUS% = 2 Then
           DOCUNU& = CVS(Left$(XX$, 4))
           GoTo 5
         End If
       End If
     End If
   Next UI&
   '
   If VUELBUS% < 2 Then
     VUELBUS% = 2
     GoTo 3
   End If
   '
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
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Screen.MousePointer = 1
   Call MENSERR(24, "Documento no Registrado")
   Exit Sub
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = L& To URE&
     TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
     End If
   Next KKL&
   '
95 SPOOLSTRING$ = TPSP$
   For KKI% = 0 To 5
     EDITFORM.Picture1(KKI%).Cls
     EDITFORM.Picture1(KKI%).Top = 0
     EDITFORM.Picture1(KKI%).Refresh
     EDITFORM.Picture1(KKI%).Height = 7170
     EDITFORM.Picture1(KKI%).Width = 11700
     EDITFORM.Picture1(KKI%).DrawWidth = 3
     EDITFORM.Picture1(KKI%).DrawStyle = 0
     EDITFORM.Picture1(KKI%).FillStyle = 1
   Next KKI%
   ALTUPAGINA = 7170
   TOPEPAGINA = 0
   PAGINACTIVA% = 0
   CAPAGINAS% = 0
   '
   FIN& = Len(TPSP$)
10 PPXX% = InStr(TPSP$, Chr$(255))
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
   Screen.MousePointer = 1
   PAGINACTIVA% = 0
   For KKI% = 0 To 5
     EDITFORM.Picture1(KKI%).Visible = False
     If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
   Next KKI%
   '
   EDITFORM.Picture1(PAGINACTIVA%).Refresh
   EDITFORM.VScroll1.Min = 0
   EDITFORM.VScroll1.Value = 0
   EDITFORM.VScroll1.Max = 1
   If ALTUPAGINA > EDITFORM.Frame1.Height Then
      EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
   End If
   EDITFORM.Command3.Enabled = False
   EDITFORM.Command5.Enabled = False
   EDITFORM.Show 1
   '
End Sub

Private Sub LTEXT36_Change()
   DISPONI = Val(LTEXT36.Caption) - Val(LTEXT37.Caption) - Val(LTEXT35.Caption)
   LTEXT23.ForeColor = &H80000008
   If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
   LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
   LTEXT23.Refresh
End Sub

Private Sub NueCli_Click()
   Call Command21_Click
End Sub

Private Sub OrDBCli_Click()
Call COMMAND15_CLICK
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

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab <> PreviousTab Then
      Call GRAGENCLI
      Call GRANOTACLI
      Call FRESHALL
      Call CARDATCLIE
   End If
End Sub



Private Sub Susp_Cli_Click()
   Call Command23_Click
End Sub

Private Sub Text1_Change()
   For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call BLANFORMU
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE = Val(Text1.TEXT)
       If STATCLI(NCLIE) > 0 Then
         'NC% = NCLIE%
         Call CARDATCLIE
       End If
     End If
   End If
End Sub
'
Private Sub Text1_DblClick()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   '
   Call SELECHO("DEUDOR1")
   NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% < 1 Then Exit Sub
   '
   NC% = NCLIE%
   Text1.TEXT = TRIM$(Str$(NC%))
   '
End Sub

Private Sub Text1_GotFocus()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   '
End Sub

Private Sub Text19_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text20_Change()
   MODIFIC%(0) = 1
   If TRIM$(UCase$(Text20.TEXT)) = "SUSPENDIDO" Then
       LTEXT36.Caption = Text20.TEXT
     Else
       LTEXT36.Caption = CSTRING$(MKS$(Val(Text20.TEXT)), 4, 11, 2, 2)
   End If
End Sub

Private Sub Text20_LostFocus()
  If TRIM$(UCase$(Text20.TEXT)) <> "SUSPENDIDO" Then
    Text20.TEXT = CSTRING$(MKS$(Val(Text20.TEXT)), 3, 12, 2, 2)
  End If
End Sub

Private Sub Text21_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text23_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text23_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text24_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text24_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text25_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text26_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text27_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text28_Change()
   MODIFIC(1) = 1
End Sub

Private Sub Text29_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text3_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text3_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text30_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text31_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text32_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text33_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text34_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub LText35_DblClick()
   Screen.MousePointer = 11
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NC% = Val(Text1.TEXT)
       If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
         Call COPYSTRU("CHECAR", "LICHECLI")
         Call HEADERS("LICHECLI", "")
         Call HEADERS("LICHECLI", "Cliente: " + RASOCLI$(NC%))
         '
         JJL& = 0
         RFF$ = RECFILT$("CHECAR", 3, MKI$(NC%))
         HOII% = HOY(HOS$)
         For UI& = 1 To Len(RFF$) Step 4
            RECHE& = CVS(Mid$(RFF$, UI&, 4))
            If RECHE& > 0 Then
               If RECHE& <= ULTREG&("CHECAR") Then
                  XX$ = REGLEIDO$("CHECAR", RECHE&)
                  If CVI(Mid$(XX$, 3, 2)) = NC% Then
                     If CVI(Mid$(XX$, 41, 2)) >= HOII% Then
                        JJL& = JJL& + 1
                        Call GRAREG("LICHECLI", XX$, JJL&)
                     End If
                  End If
               End If
            End If
         Next UI&
         Call SETULTREG("LICHECLI", JJL&)
         Call CIERRARCH("LICHECLI")
         Call CIERRARCH("CHECAR")
         '
         Call CONECRUN("*LICHECLI", "Listado de Cheques por Cliente/WAIT")
         '
       End If
     End If
   End If
   '
   Screen.MousePointer = 1
   '
End Sub
Sub CARESCUENTA()
    '
    If YACAR%(2) > 0 Then Exit Sub
    '
50  Screen.MousePointer = 11
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then
        List3.Clear
        List3.Refresh
        List2.Clear
        List2.Refresh
        LTEXT23.Caption = ""
        LTEXT37.Caption = ""
        LTEXT36.Caption = ""
        LTEXT35.Caption = ""
        LTEXT37.Caption = CSTRING$(MKD$(SALDCLI#(NC%)), 4, 11, 2, 2)
        LMCR = LIMICRE(NC%)
        LTEXT36.Caption = CSTRING$(MKS$(LMCR), 4, 11, 2, 2)
        If LMCR < 0 Then LTEXT36.Caption = " Suspendido"
        Refresh
        LTEXT35.Caption = CSTRING$(MKD$(CHEPEACRE#(NC%)), 4, 11, 2, 2)
        DISPONI = Val(LTEXT36.Caption) - Val(LTEXT37.Caption) - Val(LTEXT35.Caption)
        LTEXT23.ForeColor = &H80000008
        If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
        LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
        Refresh
    End If
    '
100 MAXDIAC& = ULTREG&("COBRAN")
    MAXE& = ULTREG&("FACTUR")
    '
    RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
    RFC$ = RECFILT$("COBRAN", 1, MKI$(NC%))
    '
    LINDICE.Clear
    For U& = 1 To Len(RFF$) Step 4
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("FACTUR", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("FACTUR", 1)
            Call GRAREG("FACTUR", X$, 1)
            Call CIERRARCH("FACTUR")
            GoTo 100
        End If
        FF% = CVI(Mid$(X$, 7, 2))
        Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2)
        LINDICE.AddItem Z$
    Next U&
    '
    For U& = 1 To Len(RFC$) Step 4
        REG& = CVS(Mid$(RFC$, U&, 4))
        X$ = REGLEIDO$("COBRAN", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("COBRAN", 1)
            Call GRAREG("COBRAN", X$, 1)
            Call CIERRARCH("COBRAN")
            GoTo 100
        End If
        FF% = CVI(Mid$(X$, 7, 2))
        Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$((-1) * REG&), 4, 8, 0, 2)
        LINDICE.AddItem Z$
    Next U&
    '
    SALDO# = 0: SARRAS# = 0
    For U& = 1 To LINDICE.ListCount
        LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(LINDICE.TEXT, 7, 8))
        If REG& > 0 Then
            If REG& > MAXE& Then GoTo 990
            X$ = REGLEIDO$("FACTUR", REG&)
            TI% = Asc(Mid$(X$, 13, 1))
            TIPOS$ = "Factura": If TI% = 0 Then TIPOS$ = "Ajuste de Saldo"
                               If TI% = 1 Then TIPOS$ = "Venta Contado"
                               If TI% = 3 Then TIPOS$ = "Nota de Debito"
                               If TI% = 5 Then TIPOS$ = "Descuento Caja"
                               If TI% = 6 Then TIPOS$ = "Nota de Credito"
                               If TI% = 7 Then TIPOS$ = "Ajuste de Saldo"
                               '
            NROL& = CVS(Mid$(X$, 3, 4))
            If NROL& < 1 Then
               If TI% = 0 Or TI% = 7 Then
                  Call REPLA(X$, MKS$(REG&), 3, 4)
                  Call GRAREG("FACTUR", X$, REG&)
               End If
            End If
            '
            SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
            IMPO# = CVD(Mid$(X$, 15, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 23, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 31, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 49, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 57, 8))
            IMPO# = SG% * IMPO#
            GoTo 890
            '
         ElseIf REG& < 0 Then
            REG1& = Abs(REG&)
            If REG1& > MAXDIAC& Then GoTo 990
            X$ = REGLEIDO$("COBRAN", REG1&)
            TI% = 4
            TIPOS$ = "Recibo"
            SG% = (-1)
            IMPO# = (-1) * CVD(Mid$(X$, 15, 8))
            GoTo 890
        End If
        GoTo 990
        '
890     SALDO# = SALDO# + (Int(100 * IMPO# + 0.5)) / 100
        FEX = CVI(Mid$(X$, 7, 2))
        NROL& = CVS(Mid$(X$, 3, 4))
        Z$ = Space$(80)
        '
        Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
        Call REPLA(Z$, TIPOS$, 11, 16)
        Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 27, 8)
        IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 13, 2, 2)
        SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
        If IMPO# > 0 Then
            Call REPLA(Z$, IMPTEX$, 35, 13)
          ElseIf IMPO# < 0 Then
            Call REPLA(Z$, IMPTEX$, 50, 13)
        End If
        Call REPLA(Z$, SALTEX$, 63, 14)
        List3.AddItem Z$
        '
990 Next U&
    '
991 SARRAS# = SALDO#
    List3.ListIndex = List3.ListCount - 1
    List3.Refresh
    '
    MAXDEU& = ULTREG&("SVD202")
    RFD$ = RECFILT$("SVD202", 1, MKI$(NC%))
    '
    LINDICE.Clear
    For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("SVD202", REG&)
        NCII% = CVI(Left$(X$, 2))
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("SVD202", 1)
            Call GRAREG("SVD202", X$, 1)
            Call CIERRARCH("SVD202")
            GoTo 100
        End If
        IMPO# = CVD(Mid$(X$, 21, 8))
        SALDO# = CVD(Mid$(X$, 11, 8))
        If Abs(SALDO#) < 0.005 Then
          If Abs(SALDO#) > 0.0005 Then
            SALDO# = 0
            Call REPLA(X$, MKD$(0), 11, 8)
            Call GRAREG("SVD202", X$, REG&)
          End If
        End If
        '
        If Abs(SALDO#) >= 0.005 Then
          If Sgn(IMPO#) <> Sgn(SALDO#) Or Abs(SALDO#) > Abs(IMPO#) + 0.005 Then
            ABUS$ = Left$(X$, 8)
            For UU& = 1 To Len(RFD$) Step 4
                REG& = CVS(Mid$(RFD$, UU&, 4))
                X$ = REGLEIDO$("SVD202", REG&)
                If Left$(X$, 8) = ABUS$ Then
                    Call GRAREG("SVD202", String$(32, 0), REG&)
                End If
            Next UU&
            GoTo 100
          End If
        End If
        '
        If Abs(SALDO#) > 0 Then
            FF% = CVI(Mid$(X$, 9, 2))
            Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2)
            LINDICE.AddItem Z$
        End If
    Next U&
    '
    SUMADEU# = 0: SUMAFAC# = 0
    For U& = 1 To LINDICE.ListCount
        LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(LINDICE.TEXT, 7, 8))
        If REG& > 0 Then
            If REG& <= MAXDEU& Then
                X$ = REGLEIDO$("SVD202", REG&)
                TI% = Asc(Mid$(X$, 3, 1))
                TIPOS$ = "Factura": If TI% = 0 Then TIPOS$ = "Ajuste de Saldo"
                                   If TI% = 1 Then TIPOS$ = "Venta Contado"
                                   If TI% = 3 Then TIPOS$ = "Nota de Debito"
                                   If TI% = 4 Then TIPOS$ = "Anticipo"
                                   If TI% = 5 Then TIPOS$ = "Descuento Caja"
                                   If TI% = 6 Then TIPOS$ = "Nota de Credito"
                                   If TI% = 7 Then TIPOS$ = "Ajuste de Saldo"
                                   '
                SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
                IMPO# = CVD(Mid$(X$, 21, 8))
                If TI% <> 0 Then IMPO# = Abs(IMPO#)
                IMPO# = SG% * IMPO#
                SUMAFAC# = SUMAFAC# + IMPO#
                '
                SALDO# = CVD(Mid$(X$, 11, 8))
                If TI% <> 0 Then SALDO# = Abs(SALDO#)
                SALDO# = SG% * SALDO#
                SUMADEU# = SUMADEU# + SALDO#
                '
                NROL& = CVS(Mid$(X$, 5, 4))
                FEMI = CVI(Mid$(X$, 19, 2))
                FEVE = CVI(Mid$(X$, 9, 2))
                '
1234            FEVEX$ = FECHATEX$(FEVE)
                If Val(Mid$(FEVEX$, 4, 2)) = 2 Then
                  If Val(Mid$(FEVEX$, 1, 2)) > 28 Then
                    FEVE = DIPOST(FEVE)
                    GoTo 1234
                  End If
                End If
                If CVI(Mid$(X$, 9, 2)) <> FEVE Then
                  FEVEI% = FEVE
                  Call REPLA(X$, MKI$(FEVEI%), 9, 2)
                  Call GRAREG("SVD202", X$, REG&)
                End If
                '
                IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
                SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
                '
If Abs(Abs(IMPO#) - Abs(SALDO#)) < 0.005 Then ' NO ESTA CANCELADA
  MONEMX% = Asc(Mid$(X$, 31, 1))
  If MONEMX% > 2 Then
    MONEMX% = 0
    Call REPLA(X$, Chr$(0), 31, 1)
    Call GRAREG("SVD202", X$, REG&)
  End If
  If MONEMX% = 0 Then
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If CVS(Mid$(XXC$, 39, 4)) = NROL& Then
        If Asc(Mid$(XXC$, 37, 1)) = TI% Then
          MONEMX% = Asc(Mid$(X$, 31, 1))
          GoTo 15
        End If
      End If
    Next KUL&
15 If MONEMX% <> 2 Then MONEMX% = 1
    Call REPLA(X$, Chr$(MONEMX%), 31, 1)
    Call GRAREG("SVD202", X$, REG&)
  End If
  If MONEMX% = 2 Then
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If CVS(Mid$(XXC$, 39, 4)) = NROL& Then
        If Asc(Mid$(XXC$, 37, 1)) = TI% Then
          SALTEX$ = "u$s" + FORMATNUM$(CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8)), "F11.2")
          GoTo 20
        End If
      End If
    Next KUL&
  End If
End If
'
20              If TRIM$(SALTEX$) <> "" Then
                  Z$ = Space$(80)
                  Call REPLA(Z$, TIPOS$, 1, 16)
                  Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 17, 8)
                  Call REPLA(Z$, FECHATEX$(FEMI), 27, 8)
                  Call REPLA(Z$, IMPTEX$, 37, 14)
                  Call REPLA(Z$, FECHATEX$(FEVE), 53, 8)
                  Call REPLA(Z$, SALTEX$, 63, 14)
                  '
                  List2.AddItem Z$
                End If
                '
            End If
        End If
        '
    Next U&
    '
    List2.ListIndex = List2.ListCount - 1
    List2.Refresh
    Label28.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label28.Refresh
    Label25.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label25.Refresh
    '
    If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
       ATENTI% = 1
       TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta"
       If Visible = True Then
           TTXX$ = TTXX$ + " Elegida."
         Else
           TTXX$ = TTXX$ + ":\   \'" + TRIM$(RASOCLI$(NC%)) + "'"
       End If
       TTXX$ = TTXX$ + "\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Créditos y Cobranzas\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
       '
       If COMALTER%(TTXX$) = 2 Then
           Call REAPLICRE(NC%)
           GoTo 50
         Else
           ATENTI% = 0
           Call CIERRARCH("*.*")
           Call FINAL("")
       End If
    End If
    '
    If Abs(SARRAS# - SALDCLI#(NC%)) >= 0.005 Then
        RGXC& = REGICLI&(NC%)
        XC$ = REGLEIDO$("DEUDOR2", RGXC&)
        Call REPLA(XC$, MKD$(SARRAS#), 109, 8)
        Call GRAREG("DEUDOR2", XC$, RGXC&)
        LTEXT37.Caption = CSTRING$(MKD$(SARRAS#), 4, 11, 2, 2)
        Refresh
    End If
    YACAR%(2) = 1
    Screen.MousePointer = 1
    '
End Sub
'
Sub REAPLICRE(NC%)
    '
    ATENTI% = 1
    Dim DOCU$(8192), IMPO#(8192), REGDO&(8192), FFDOC%(8192), IORI#(8192)
    '
100 Screen.MousePointer = 11
    '
    SUMARRAS# = 0
    If List3.ListCount > 0 Then
      SUMARRAS# = Val(Mid$(List3.List(List3.ListCount - 1), 63, 14))
    End If
    SUMADEUP# = Val(Label28.Caption)
    CADOCPEN% = List2.ListCount
    '
    MAXDIAC& = ULTREG&("COBRAN")
    MAXE& = ULTREG&("FACTUR")
    '
    RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
    RFC$ = RECFILT$("COBRAN", 1, MKI$(NC%))
    '
    NUDOCU% = 0
    FIN& = Len(RFF$)
    For U& = 1 To Len(RFF$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("FACTUR", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("FACTUR", 1)
            Call GRAREG("FACTUR", X$, 1)
            Call CIERRARCH("FACTUR")
            GoTo 100
        End If
110     NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = Asc(Mid$(X$, 13, 1))
        If NUDO& < 1 Then
           If TI% = 0 Or TI% = 7 Then
              Call REPLA(X$, MKS$(REG&), 3, 4)
              Call GRAREG("FACTUR", X$, REG&)
              GoTo 110
           End If
        End If
        VA% = Asc(Mid$(X$, 14, 1))
        SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
115     IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 23, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 31, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 49, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 57, 8))
        IMPO#(NUDOCU%) = SG% * (Int(100 * IMPO#(NUDOCU%) + 0.5)) / 100
        NUSUCUVE% = CVI(Mid$(X$, 47, 2))
        DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, NUSUCUVE%), 2) + "-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
        '
        If TI% > 0 Then
          If TI% < 7 Then
            If CVD(Mid$(X$, 15, 8)) < (-0.005) Then
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 15, 8)))), 15, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 23, 8)))), 23, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 31, 8)))), 31, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 49, 8)))), 49, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 57, 8)))), 57, 8)
              Call GRAREG("FACTUR", X$, REG&)
              GoTo 115
            End If
          End If
        End If
    Next U&
    '
    FIN& = Len(RFC$)
    For U& = 1 To Len(RFC$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFC$, U&, 4))
        X$ = REGLEIDO$("COBRAN", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("COBRAN", 1)
            Call GRAREG("COBRAN", X$, 1)
            Call CIERRARCH("COBRAN")
            GoTo 100
        End If
        NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = 4
        VA% = 1
        SG% = (-1)
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
        IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = SG% * (Int(100 * IMPO#(NUDOCU%) + 0.5)) / 100
        DOCU$(NUDOCU%) = "RB-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = (-1) * REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    If NUDOCU% >= 1 Then
       If Left$(DOCU$(1), 2) = "AS" Then
          FIN& = ULTREG&("SALINID")
          For U& = 1 To FIN&
             X$ = REGLEIDO$("SALINID", U&)
             If CVI(Left$(X$, 2)) = NC% Then
               TI% = Asc(Mid$(X$, 3, 1))
               VA% = Asc(Mid$(X$, 4, 1))
               NUSUCUVE% = CVI(Mid$(X$, 29, 2))
               SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
               IPEN# = SG% * CVD(Mid$(X$, 11, 8))
               If Sgn(IPEN#) = Sgn(IMPO#(1)) Then
                 If Abs(IPEN#) <= Abs(IMPO#(1)) Then
                   IMPO#(1) = IMPO#(1) - IPEN#
                   NUDO& = CVS(Mid$(X$, 5, 4))
                   NUDOCU% = NUDOCU% + 1
                   If NUDOCU% > 8192 Then
                      Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
                      Call CIERRARCH("*.*")
                      Call FINAL("")
                   End If
                   IMPO#(NUDOCU%) = IPEN#
                   DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, NUSUCUVE%), 2) + "-" + TRIM$(Str$(NUDO&))
                   REGDO&(NUDOCU%) = 900000 + U&
                   FFDOC%(NUDOCU%) = CVI(Mid$(X$, 19, 2))
                 End If
               End If
             End If
          Next U&
       End If
    End If
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
' HASTA AQUI O.K
    KKI% = 0: SUMAPEX# = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          SUMAPEX# = SUMAPEX# + IMPO#(KKI%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    FIN& = ULTREG&("APLICOB")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("APLICOB", U&)
       If CVI(Mid$(X$, 3, 2)) = NC% Then
          RECI$ = Mid$(X$, 5, 18)
          RECI$ = Left$(RECI$, 2) + "-" + TRIM$(Val(Mid$(RECI$, 11)))
          FACU$ = Mid$(X$, 55, 18)
          PINI% = 11: If Mid$(FACU$, 17, 1) = "-" Then PINI% = 9
          FACU$ = Left$(FACU$, 2) + "-" + TRIM$(Val(Mid$(FACU$, PINI%, 8)))
          '
210       IAPLI# = CVD(Mid$(X$, 23, 8))
          IRECIMA# = 0: IFACMAX# = 0
          For KKI% = 1 To NUDOCU%
             If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
               If IMPO#(KKI%) < 0 Then
                 IRECIMA# = IRECIMA# - IMPO#(KKI%)
               End If
             End If
             If DOCU$(KKI%) = FACU$ Then
               If IMPO#(KKI%) > 0 Then
                 IFACMAX# = IFACMAX# + IMPO#(KKI%)
               End If
             End If
          Next KKI%
          '
          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
          '
          IRECIMA# = IAPLI#
          IFACMAX# = IAPLI#
          '
          If IAPLI# > 0 Then
             For KKI% = 1 To NUDOCU%
                If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
                   If IMPO#(KKI%) < 0 Then
                      IAP1# = IRECIMA#
                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) + IAP1#
                      IRECIMA# = IRECIMA# - IAP1#
                   End If
                End If
                If DOCU$(KKI%) = FACU$ Then
                   If IMPO#(KKI%) > 0 Then
                      IAP2# = IFACMAX#
                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) - IAP2#
                      IFACMAX# = IFACMAX# - IAP2#
                   End If
                End If
             Next KKI%
          End If
          '
       End If
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Or (FFDOC%(KKI%) = FFDOC%(KKJ%) And Sgn(REGDO&(KKI%)) = Sgn(REGDO&(KKJ%)) > 0 And Abs(REGDO&(KKI%)) > Abs(REGDO&(KKJ%))) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    KKI% = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    For KKI% = 1 To NUDOCU%
       For KKJ% = KKI% + 1 To NUDOCU%
          If Abs(IMPO#(KKI%) + IMPO#(KKJ%)) < 0.005 Then
             IMPO#(KKI%) = 0
             IMPO#(KKJ%) = 0
          End If
       Next KKJ%
    Next KKI%
    '
    KKI% = 0: SUMAPEX# = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          SUMAPEX# = SUMAPEX# + IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    Screen.MousePointer = 1
    If COMALTER%("Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(RASOCLI$(NC%)) + "'\   \Desea Grabar la Nueva Composición de la Deuda ?\\Cancelar\Grabar") <> 2 Then
       GoTo 999
    End If
    '
    ' para no enquilombar las cuentas que mas o menos estan
    If Abs(SUMARRAS# - SUMADEUP#) < 0.005 Then
      If NUDOCU% >= CADOCPEN% Then
        GoTo 999
      End If
    End If
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("SVD202")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("SVD202", U&)
       If CVI(Left$(X$, 2)) = NC% Then
          Call REPLA(X$, String$(30, 0), 3, 30)
          Call GRAREG("SVD202", X$, U&)
       End If
    Next U&
    '
    LU$ = LUDAT$
    ARCHIDEU$ = "SVD202"
    RDX% = 1
    FIN& = NUDOCU%
    YABLO% = 0
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       If Abs(IMPO#(KKI%)) >= 0.005 Then
          IORIG# = IMPO#(KKI%)
          If REGDO&(KKI%) > 0 Then
             If REGDO&(KKI%) > 900000 Then
                X$ = REGLEIDO$("SALINID", REGDO&(KKI%) - 900000)
                NRO = CVS(Mid$(X$, 5, 4))
                TIPO% = Asc(Mid$(X$, 3, 1))
                VA% = Asc(Mid$(X$, 4, 1))
                FFAC% = CVI(Mid$(X$, 9, 2))
                COPAGI% = 0
                GoTo 290
             End If
             X$ = REGLEIDO$("FACTUR", REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             TIPO% = Asc(Mid$(X$, 13, 1))
             VA% = Asc(Mid$(X$, 14, 1))
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = CPAGCLI%(NC%)
             If TIPO% < 1 Or TIPO% > 3 Then COPAGI% = 0
             SG% = 1: If TIPO% >= 4 Then If TIPO% <= 6 Then SG% = -1
             IORIG# = CVD(Mid$(X$, 15, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 23, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 31, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 49, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 57, 8))
             IORIG# = SG% * IORIG#
           Else
             X$ = REGLEIDO$("COBRAN", (-1) * REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             TIPO% = 4
             VA% = 0
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = 0
          End If
290       TOTFAC# = IMPO#(KKI%)
          If YABLO% = 0 Then
             YABLO% = 1
             Call BLOQARCH(ARCHIDEU$)
          End If
          ARDEX$ = ARCHIDEU$ + "/" + TRIM$(CSTRING$(MKD$(IORIG#), 4, 12, 2, 2))
          Call GRADEU(LU$, ARDEX$, RDX%, TIPO%, VA%, NRO, NC%, FFAC%, COPAGI%, TOTFAC#)
       End If
    Next KKI%
    '
999 Call CIERRARCH(ARCHIDEU$)
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    ATENTI% = 0
End Sub


Private Sub Text36_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub VenCli_Click()
With OPMASCL04
  .Option1(4).Value = True
  Call .Command16_Click
End With
End Sub

Sub GENECOCLI()
   If UPDATCLI% > 0 Then
      Screen.MousePointer = 11
      Call ABRECLIEN
      Set CLIENTECO = Badaclie.OpenRecordset("SELECT Codi_Cli, Nume_Cli,Raso_Cli FROM Deudor12 WHERE Stat_Cli > 0 ORDER BY Raso_Cli ASC;", dbOpenDynaset, dbReadOnly)
      On Error Resume Next
      CLIENTECO.MoveLast
      If Err Then Exit Sub
      '
      CAREC% = CLIENTECO.RecordCount
      If CAREC% < 1 Then Exit Sub
      '
      FIN& = CAREC%
      REBLA$ = REGBLAN$("DEUDOR1")
      REBLB$ = REGBLAN$("INDICLIE")
      JJ& = 0
      With CLIENTECO
        .MoveFirst
        '
10      Call TRACE(20, JJ&, FIN&)
        EL1$ = REBLA$
        Call REPLA(EL1$, MKI$(!Nume_Cli), 1, 2)
        Call REPLA(EL1$, !Raso_Cli, 3, 30)
        JJ& = JJ& + 1
        Call GRAREG("DEUDOR1", EL1$, JJ&)
        '
        EL2$ = REBLB$
        Call REPLA(EL2$, !Codi_Cli, 1, 12)
        Call REPLA(EL2$, !Raso_Cli, 13, 50)
        Call REPLA(EL2$, MKI$(!Nume_Cli), 63, 2)
        Call GRAREG("INDICLIE", EL2$, JJ&)
        '
        If JJ& < CAREC% Then
          CLIENTECO.MoveNext
          GoTo 10
        End If
      End With
      '
      CLIENTECO.Close
      Call SETULTREG("DEUDOR1", JJ&)
      Call CIERRARCH("DEUDOR1")
      Call FILESORT("INDICLIE", "", 1, 1, "ASC")
      Call FRESHECHO("DEUDOR1")
      Call FRESHECHO("INDICLIE")
      Call BAJALDISCO
      UPDATCLI% = 0
      Screen.MousePointer = 1
   End If
End Sub
'
Sub GENDAPIBCLI()
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("DEUDOR1")
    REBLA$ = REGBLAN$("DAPIBCLI")
    Call REPLA(REBLA$, USERNAME$, 41, 20)
    Call REPLA(REBLA$, MKI$(HOY(HOS$)), 61, 2)
    Call REPLA(REBLA$, MKI$(HORANUM%(Time$)), 63, 2)
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("DEUDOR1", U&)
      NCLIX% = CVI(Left$(XX$, 2))
      If NCLIX% > 0 Then
        PRC$ = UCase$(PROVCLI$(NCLIX%))
        RECORD$ = REBLA$
        Call REPLA(RECORD$, MKI$(NCLIX%), 1, 2)
        If PRC$ = "B" Then
            Call REPLA(RECORD$, Chr$(1), 3, 1)
            Call REPLA(RECORD$, MKS$(100), 5, 4)
          Else
            Call REPLA(RECORD$, Chr$(0), 3, 1)
            Call REPLA(RECORD$, MKS$(0), 5, 4)
            Call REPLA(RECORD$, "EXTRAZONA", 9, 32)
        End If
        Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIX%), RECORD$)
      End If
    Next U&
    Call CIERRARCH("DAPIBCLI")
    Screen.MousePointer = 1
    '
End Sub

