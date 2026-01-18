VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form ACLI_GES04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8D8C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes - A-B-M-L Base de Datos de Clientes"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "ACLI_GES04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   Begin VB.PictureBox Picture6 
      Height          =   8300
      Left            =   11025
      ScaleHeight     =   8235
      ScaleWidth      =   1215
      TabIndex        =   181
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
      Begin VB.Image Image3 
         Height          =   8415
         Left            =   0
         Picture         =   "ACLI_GES04.frx":0442
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Recuperacion de Datos de Clientes"
      Height          =   255
      Left            =   8160
      TabIndex        =   178
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00D8D8C0&
      Caption         =   "Destino Fabrica"
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
      Left            =   5700
      TabIndex        =   133
      Top             =   210
      Width           =   1170
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
      Left            =   3360
      TabIndex        =   126
      Top             =   315
      Width           =   1965
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
      Left            =   5295
      Picture         =   "ACLI_GES04.frx":27DA
      TabIndex        =   125
      Top             =   315
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
      Left            =   8820
      TabIndex        =   120
      Top             =   210
      Width           =   1990
      Begin VB.CommandButton Command20 
         Height          =   500
         Left            =   1280
         Picture         =   "ACLI_GES04.frx":2AE4
         Style           =   1  'Graphical
         TabIndex        =   123
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
         Picture         =   "ACLI_GES04.frx":3E4A
         Style           =   1  'Graphical
         TabIndex        =   122
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
         Picture         =   "ACLI_GES04.frx":4154
         Style           =   1  'Graphical
         TabIndex        =   121
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
      TabIndex        =   107
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
         Picture         =   "ACLI_GES04.frx":445E
         Style           =   1  'Graphical
         TabIndex        =   111
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
         TabIndex        =   118
         Top             =   7140
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   119
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
         Picture         =   "ACLI_GES04.frx":4768
         Style           =   1  'Graphical
         TabIndex        =   117
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
         Picture         =   "ACLI_GES04.frx":4DD2
         Style           =   1  'Graphical
         TabIndex        =   116
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
         Picture         =   "ACLI_GES04.frx":50DC
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   6300
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
         Picture         =   "ACLI_GES04.frx":551E
         Style           =   1  'Graphical
         TabIndex        =   114
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
         Picture         =   "ACLI_GES04.frx":5828
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Baja Lógica de la Cuenta / Restauración Cuentas Dadas de Baja"
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
         Picture         =   "ACLI_GES04.frx":5B32
         Style           =   1  'Graphical
         TabIndex        =   110
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
         Picture         =   "ACLI_GES04.frx":5E3C
         Style           =   1  'Graphical
         TabIndex        =   109
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
         Picture         =   "ACLI_GES04.frx":5F86
         Style           =   1  'Graphical
         TabIndex        =   108
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
         Picture         =   "ACLI_GES04.frx":6290
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Alta de Nuevo Cliente"
         Top             =   1890
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "ACLI_GES04.frx":63DA
         Stretch         =   -1  'True
         Top             =   7455
         Width           =   1410
      End
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
      Height          =   330
      Left            =   7035
      TabIndex        =   104
      Top             =   735
      Width           =   1425
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
      Height          =   330
      Left            =   1155
      TabIndex        =   2
      Top             =   735
      Width           =   5550
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
      Picture         =   "ACLI_GES04.frx":82FC
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6675
      Left            =   105
      TabIndex        =   36
      Top             =   1260
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11774
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14211264
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "ACLI_GES04.frx":8606
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(2)=   "Label17"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Anotaciones"
      TabPicture(1)   =   "ACLI_GES04.frx":8622
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
      TabCaption(2)   =   "Cuenta Corriente"
      TabPicture(2)   =   "ACLI_GES04.frx":863E
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame18"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Entrega"
      TabPicture(3)   =   "ACLI_GES04.frx":865A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame23"
      Tab(3).Control(1)=   "Frame22"
      Tab(3).Control(2)=   "Frame20"
      Tab(3).Control(3)=   "Frame1"
      Tab(3).ControlCount=   4
      Begin VB.Frame Frame1 
         Caption         =   "Sucursales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   -68595
         TabIndex        =   173
         Top             =   600
         Width           =   4005
         Begin VB.CommandButton Command26 
            Caption         =   "Editar"
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
            Left            =   2950
            Style           =   1  'Graphical
            TabIndex        =   176
            ToolTipText     =   "Modificar Datos de la Sucursal  Activa"
            Top             =   1260
            Width           =   840
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Nuevo"
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
            Left            =   2100
            Style           =   1  'Graphical
            TabIndex        =   175
            ToolTipText     =   "Agrega Nueva Sucursal"
            Top             =   1260
            Width           =   840
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
            Height          =   960
            ItemData        =   "ACLI_GES04.frx":8676
            Left            =   210
            List            =   "ACLI_GES04.frx":8678
            TabIndex        =   174
            Top             =   260
            Width           =   3585
         End
      End
      Begin VB.Frame Frame20 
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
         Height          =   3480
         Left            =   -74790
         TabIndex        =   144
         Top             =   2950
         Width           =   10200
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
            Height          =   330
            Left            =   6670
            TabIndex        =   168
            Top             =   525
            Width           =   175
         End
         Begin VB.TextBox Text45 
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
            Left            =   5985
            TabIndex        =   167
            Top             =   525
            Width           =   750
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Nuevo"
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
            Left            =   7770
            Style           =   1  'Graphical
            TabIndex        =   155
            ToolTipText     =   "Agrega Nuevo Transporte a la Base de Datos"
            Top             =   525
            Width           =   1050
         End
         Begin VB.Frame Frame21 
            Caption         =   "Teléfono - Fax - E-Mail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1380
            Left            =   210
            TabIndex        =   146
            Top             =   1890
            Width           =   9830
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "e-mail:"
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
               Left            =   525
               TabIndex        =   154
               Top             =   945
               Width           =   645
            End
            Begin VB.Label Label48 
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
               Height          =   225
               Left            =   5565
               TabIndex        =   153
               Top             =   525
               Width           =   645
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Teléfono(s):"
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
               TabIndex        =   152
               Top             =   525
               Width           =   1065
            End
            Begin VB.Label Label65 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   1260
               TabIndex        =   151
               Top             =   420
               Width           =   4215
            End
            Begin VB.Label Label66 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   6300
               TabIndex        =   150
               Top             =   420
               Width           =   3375
            End
            Begin VB.Label Label67 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   1260
               TabIndex        =   149
               Top             =   840
               Width           =   4215
            End
            Begin VB.Label Label69 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   6300
               TabIndex        =   148
               Top             =   840
               Width           =   3375
            End
            Begin VB.Label Label70 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Att.Sr:"
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
               Left            =   5565
               TabIndex        =   147
               Top             =   945
               Width           =   645
            End
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Editar"
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
            Left            =   8820
            Style           =   1  'Graphical
            TabIndex        =   145
            ToolTipText     =   "Modificar Datos del Transporte Activo"
            Top             =   525
            Width           =   1065
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.U.I.T:"
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
            Left            =   6405
            TabIndex        =   166
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.:"
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
            Left            =   5355
            TabIndex        =   165
            Top             =   630
            Width           =   540
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
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
            TabIndex        =   164
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label54 
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
            Height          =   225
            Left            =   105
            TabIndex        =   163
            Top             =   1470
            Width           =   1275
         End
         Begin VB.Label Label55 
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
            Height          =   225
            Left            =   5145
            TabIndex        =   162
            Top             =   1470
            Width           =   1275
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Razon Social:"
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
            Left            =   -105
            TabIndex        =   161
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label61 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   1470
            TabIndex        =   160
            Top             =   525
            Width           =   3900
         End
         Begin VB.Label Label62 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   1470
            TabIndex        =   159
            Top             =   945
            Width           =   5370
         End
         Begin VB.Label Label63 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   1470
            TabIndex        =   158
            Top             =   1365
            Width           =   3900
         End
         Begin VB.Label Label60 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   7770
            TabIndex        =   157
            Top             =   945
            Width           =   2115
         End
         Begin VB.Label Label64 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   6510
            TabIndex        =   156
            Top             =   1365
            Width           =   3375
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Domicilio de Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   -74790
         TabIndex        =   135
         Top             =   600
         Width           =   6105
         Begin VB.CommandButton Command27 
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
            Height          =   360
            Left            =   5720
            TabIndex        =   177
            Top             =   1680
            Width           =   175
         End
         Begin VB.TextBox Text5 
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
            Left            =   1470
            MaxLength       =   48
            TabIndex        =   169
            Text            =   " "
            Top             =   1260
            Width           =   4425
         End
         Begin VB.TextBox Text43 
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
            Left            =   1470
            MaxLength       =   48
            TabIndex        =   139
            Text            =   " "
            Top             =   420
            Width           =   4425
         End
         Begin VB.TextBox Text44 
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
            Left            =   1470
            MaxLength       =   48
            TabIndex        =   138
            Text            =   " "
            Top             =   840
            Width           =   4425
         End
         Begin VB.TextBox Text35 
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
            Left            =   1470
            TabIndex        =   137
            Top             =   1680
            Width           =   2640
         End
         Begin VB.TextBox Text37 
            BackColor       =   &H00FFFFFF&
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
            Left            =   4540
            TabIndex        =   136
            Top             =   1680
            Width           =   1170
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Transporte:"
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
            TabIndex        =   170
            Top             =   1365
            Width           =   1065
         End
         Begin VB.Label Label58 
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
            Height          =   225
            Left            =   315
            TabIndex        =   143
            Top             =   945
            Width           =   1065
         End
         Begin VB.Label Label59 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Calle y Nro.:"
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
            TabIndex        =   142
            Top             =   525
            Width           =   1065
         End
         Begin VB.Label Label68 
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
            Height          =   225
            Left            =   315
            TabIndex        =   141
            Top             =   1785
            Width           =   1065
         End
         Begin VB.Label Label72 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Z:"
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
            Left            =   3400
            TabIndex        =   140
            Top             =   1785
            Width           =   1065
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -68595
         TabIndex        =   134
         Top             =   2280
         Width           =   4005
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2835
            TabIndex        =   171
            Top             =   150
            Width           =   750
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Recargo por Embalaje:               %"
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
            TabIndex        =   172
            Top             =   250
            Width           =   3690
         End
      End
      Begin VB.Frame Frame18 
         BorderStyle     =   0  'None
         Caption         =   "Frame18"
         Height          =   6000
         Left            =   400
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
            Alignment       =   1  'Right Justify
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
            Alignment       =   1  'Right Justify
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
            Alignment       =   1  'Right Justify
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
            Alignment       =   1  'Right Justify
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
         Left            =   -74550
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
            Top             =   4470
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
               TabIndex        =   132
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
               Picture         =   "ACLI_GES04.frx":867A
               TabIndex        =   130
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
               TabIndex        =   129
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
               TabIndex        =   131
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
            Height          =   2775
            Left            =   5145
            TabIndex        =   61
            Top             =   1680
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
               Top             =   2235
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
               Top             =   2235
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
               Top             =   1970
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
               Top             =   1970
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
            Top             =   4100
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
               TabIndex        =   128
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
               MaxLength       =   24
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
            Height          =   2400
            Left            =   105
            TabIndex        =   51
            Top             =   1680
            Width           =   4845
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
               Left            =   1365
               TabIndex        =   179
               Text            =   " "
               ToolTipText     =   "Doble clic para abrir URL en explorador"
               Top             =   1340
               Width           =   3300
            End
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
               TabIndex        =   105
               Text            =   " "
               Top             =   2040
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
               Top             =   1680
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
               Height          =   270
               Left            =   1365
               TabIndex        =   12
               Text            =   " "
               Top             =   1005
               Width           =   3300
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Web:"
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
               TabIndex        =   180
               Top             =   1440
               Width           =   1905
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
               TabIndex        =   106
               Top             =   2090
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
               Top             =   1750
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
               Top             =   1100
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
               TabIndex        =   124
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
         Left            =   -74550
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   1575
         Width           =   9780
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         Height          =   260
         Left            =   -69690
         TabIndex        =   43
         Top             =   5790
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         Height          =   260
         Left            =   -68505
         TabIndex        =   33
         Top             =   5790
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pegar"
         Height          =   260
         Left            =   -67335
         TabIndex        =   34
         Top             =   5790
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Left            =   -66045
         TabIndex        =   35
         Top             =   5790
         Width           =   1290
      End
      Begin VB.CommandButton Command7 
         Height          =   260
         Left            =   -74550
         TabIndex        =   42
         Top             =   5790
         Width           =   4830
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         Height          =   1590
         Left            =   -72690
         ScaleHeight     =   1530
         ScaleWidth      =   2055
         TabIndex        =   41
         Top             =   1875
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Line Line3 
         X1              =   -64770
         X2              =   -64770
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Cuenta:"
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
      Left            =   2205
      TabIndex        =   127
      Top             =   240
      Width           =   1110
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor Nro:"
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
      Left            =   7245
      TabIndex        =   103
      Tag             =   " "
      Top             =   285
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "R.Social:"
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
      Left            =   -315
      TabIndex        =   38
      Top             =   840
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de Cuenta:"
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
      TabIndex        =   37
      Top             =   240
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
         Begin VB.Menu mnupantalla1 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu mnuimpresora1 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu LiCuInCli 
         Caption         =   "Listado de Cuentas Inactivas"
         Begin VB.Menu mnuPantalla2 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu munuimpresora2 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu DeLiCreCli 
         Caption         =   "Deudas y Límites de Crédito"
         Begin VB.Menu mnupantalla3 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu Mnuimpresora3 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu mnulistadodedomicilios 
         Caption         =   "Listado de Domicilios Postales"
         Begin VB.Menu mnupantalla4 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu mnuimpresora4 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu mnulistadomailclientes 
         Caption         =   "Listado de Direcciones de Mail"
         Begin VB.Menu mnupantalla5 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu mnuimpresora5 
            Caption         =   "Impresora"
         End
      End
   End
   Begin VB.Menu ConfCli 
      Caption         =   "Configuración"
      Begin VB.Menu ConGenCli 
         Caption         =   "Configuración General de Clientes"
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
      Begin VB.Menu linea 
         Caption         =   "-"
      End
      Begin VB.Menu mnuibrutos 
         Caption         =   "Configuracion Funcionamiento Percepciones IB Pcia de Bs.As"
      End
   End
   Begin VB.Menu AyuCli 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "ACLI_GES04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%(8), MODIFIC%(8)
Dim GRABATODO%
Dim AGRECLI%
Dim RECONTANT$

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



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

Private Sub Check1_Click()
   If Check1.Value = 1 Then
     If TRIM$(Text24) = "" Then
       Text24 = TRIM$(Left$(Text3, 8))
     End If
   End If
   MODIFIC%(0) = 1
End Sub

Private Sub Command11_Click()
   '
   If XVALO(TRIM$(Text1)) > 0 Then
     Call SELECHO("DEUDOR1")
     NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
     If NCLIE% < 1 Then Exit Sub
     '
     NC% = NCLIE%
     Text2 = TRIM$(Str$(NC%))
     Text4 = RASOCLI$(NC%)
   End If
   '
End Sub

Private Sub Command12_Click()
  
    '\\\OT-07-0010-RG-12/01/07///
    Command12.Enabled = False
    '
    NC% = XVALO(Text1)
    If NC% > 0 Then
       REG& = ULTREG&("!SUCENTRE") + 1
       URE% = REG&
       VDEF$ = MKI$(URE%) + Space$(190)
10     X$ = OBJPLA("SUCENTRE", VDEF$)
     If X$ <> "" Then
        'Estructura de objeto Sucentre
        'Numero de sucursal en  1 , 2
        XX1$ = TRIM$(Mid$(X$, 3, 62)) 'Sucursal
        XX2$ = TRIM$(Mid$(X$, 65, 48)) 'Domicilio
        XX3$ = TRIM$(Mid$(X$, 113, 48)) 'Localidad
        XX4$ = TRIM$(Mid$(X$, 161, 32)) 'Telefono
        XX5$ = Mid$(X$, 193, 2) 'Transporte
        XX6$ = Mid$(X$, 195, 8) 'Zona Reparto
        '
        If XX1$ = "" Or XX2$ = "" Or XX3$ = "" Or XX4$ = "" Then
           Call MENSERR(24, "Información Incompleta - Faltan Datos")
           VDEF$ = X$
           GoTo 10
        End If
         '
         'Estructura del archivo Sucentre
         
          Y$ = REGBLAN$("SUCENTRE")
          Call REPLA(Y$, X$, 1, 160) 'Sucursal,Domicilio,Localidad
          Call REPLA(Y$, MKI$(NC%), 161, 2) 'Numero de Cliente
          Call REPLA(Y$, Mid$(X$, 161, 32), 163, 32) 'Telefono
          Call REPLA(Y$, XX5$, 195, 2) 'Transporte
          Call REPLA(Y$, XX6$, 197, 8) 'Zona Reparto
         '
          Call REGAPP("!SUCENTRE", Y$)
         '
          Call CIERRARCH("!SUCENTRE")
         '
          Call FILTERUPDATE("SUCENTRE", "", 5, MKI$(NC%))
         '
          List1.Clear
         '
         'Estructura de list1
         'SUCURSAL (1,62)
         'DOMICILIO   (63,48)
         'LOCALIDAD (112,48)
         'TELEFONO (161,32)
         'NUMSUCURSAL =  (191, 4)
         'NUMERO TRANSPORTE 195, 4))
         'NUZONA$ = (199, 8)

        URE1& = ULTREG&("!SUCENTRE")
        List1.Clear
        For KKU& = 1 To URE1&
          NUMSUC% = CVI(Mid$(REGLEIDO$("!SUCENTRE", KKU&), 1, 2))
          NUMESUCURS$ = FORMATNUM(NUMSUC%, "I4")
          TTXX2$ = REGLEIDO$("!SUCENTRE", KKU&)
          ZONAREPA$ = TRIM$(Mid$(TTXX2, 197, 8))
          NUTRAI% = CVI(Mid$(REGLEIDO$("!SUCENTRE", KKU&), 195, 2))
          TTXX$ = Mid$(TTXX2$, 3, 158) + Mid$(TTXX2$, 163, 32) + NUMESUCURS$ + FORMATNUM(NUTRAI, "I4") + ZONAREPA$
          List1.AddItem TTXX$
        Next KKU&
          If URE1& > 0 Then
             List1.ListIndex = 0
             Call List1_Click
          End If
       End If
    End If
    '
    Command12.Enabled = True

    '
End Sub
Private Sub Command16_Click()
   '
   Command16.Enabled = False
   '
   If DERACCE%("ATRANSPO", "Actualizar Base de Datos de Transportes") < 2 Then
      GoTo 99
   End If
   '
   VDEF$ = REGBLAN$("TRANSPOR")
10 PROXI% = 1 + ULTREG&("TRANSPOR")
11 RFF$ = RECFILT$("TRANSPOR", 1, MKI$(PROXI%))
   If RFF$ <> "" Then
     PROXI% = PROXI% + 1
     GoTo 11
   End If
   Call REPLA(VDEF$, MKI$(PROXI%), 1, 2)
   '
   OBX$ = OBJPLA$("CARGA-TRANSPOR", VDEF$)
   If OBX$ <> "" Then
     If TRIM$(Mid$(OBX$, 3, 32)) = "" Then
       Call MENSERR(24, "Falta Nombre del Transporte")
       VDEF$ = OBX$
       GoTo 10
     End If
     '
     Call BLOQARCH("TRANSPOR")
21   RFF$ = RECFILT$("TRANSPOR", 1, MKI$(PROXI%))
     If RFF$ <> "" Then
       PROXI% = PROXI% + 1
       GoTo 21
     End If
     Call REPLA(OBX$, MKI$(PROXI%), 1, 2)
     Call REGAPP("TRANSPOR", OBX$)
     Call CIERRARCH("TRANSPOR")
   End If
   '
99 Command16.Enabled = True
   '
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   '
   If DERACCE%("ATRANSPO", "Actualizar Base de Datos de Transportes") < 2 Then
      GoTo 99
   End If
   '
   NTRA = XVALO(Text45)
   If NTRA > 0 Then
     If NTRA < 32767 Then
       NTRAI% = NTRA
       If ECOARCH$("TRANSPOR", MKI$(NTRAI%)) <> "" Then
         VDEF$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRAI%))
10       OBX$ = OBJPLA$("CARGA-TRANSPOR", VDEF$)
         If OBX$ <> "" Then
           If TRIM$(Mid$(OBX$, 3, 32)) = "" Then
             Call MENSERR(24, "Falta Nombre del Transporte")
             VDEF$ = OBX$
             GoTo 10
           End If
           '
           Call BLOQARCH("TRANSPOR")
           Call FILTERPUTRECORD("TRANSPOR", 1, MKI$(NTRAI%), OBX$)
           Call CIERRARCH("TRANSPOR")
           Call Text45_Change
           '
           NOTRANS$ = TRIM$(Mid$(OBX$, 3, 32))
           TTXX$ = "¿ Desea Aplicar el Domicilio de '" + NOTRANS$ + "'\como Domicilio de Entrega de TODOS los Clientes\Asociados a Este Transportista ?"
           If COMALTER%(TTXX$ + "\\No, Conservar los Anteriores\Si, Aplicar") = 2 Then
              Screen.MousePointer = 1
              NUTRAI% = XVALO(Text45)
              DOMINUE$ = TRIM$(Label63)
                Text43 = DOMINUE$
              LOCANUE$ = TRIM$(Label64)
                Text44 = LOCANUE$
              TRANNUE$ = TRIM$(Label61)
                Text5 = TRANNUE$
                '
              FIN& = ULTREG&("LENTRECL")
              For U& = 1 To FIN&
                Call TRACE(20, U&, FIN&)
                XY$ = REGLEIDO$("LENTRECL", U&)
                If CVI(Mid$(XY$, 99, 2)) = NUTRAI% Then
                   Call REPLA(XY$, DOMINUE$, 3, 32)
                   Call REPLA(XY$, LOCANUE$, 35, 32)
                   Call REPLA(XY$, TRANNUE$, 67, 32)
                   Call GRAREG("LENTRECL", XY$, U&)
                End If
              Next U&
              Call CIERRARCH("LENTRECL")
              Call FRESHECHO("LENTRECL")
              '
              Call APERBASDAT("CLIEN")
              SQLX$ = "SELECT * FROM DEUDOR12 "
              SQLX$ = SQLX$ + "WHERE NTrans_Cli = " & NUTRAI%
              Set CLISEL = Badaclie.OpenRecordset(SQLX$, dbOpenDynaset)
              With CLISEL
                 On Error Resume Next
                 .MoveFirst
                 If Err < 1 Then
                    On Error GoTo 0
                    Do While Not .EOF
                       .Edit
                       !Lentre0_Cli = DOMINUE$
                       !Lentre1_Cli = LOCANUE$
                       !Lentre2_Cli = TRANNUE$
                       .Update
                    .MoveNext
                    Loop
                 End If
                 On Error GoTo 0
              End With
              '
              Call COMUNI("Se ha Completado la Conversión")
              '
           End If
           '
         End If
         GoTo 99
       End If
     End If
   End If
   '
   Call MENSERR(24, "Imposible Editar Datos Transportista")
   '
99 Command18.Enabled = True
   '
End Sub

Private Sub Command19_Click()
   Call SELECHO("TRANSPOR")
   NTRAN% = XVALO(VALACT1$("TRANSPOR"))
   Text45 = TRIM$(Str$(NTRAN%))
   Label51 = Text45
End Sub

Private Sub Command22_Click()
   Static CTXX%
   Command22.Enabled = False
   '
   If ULTREG&("DAPIBCLI") < 1 Then
     Call GENDAPIBCLI
   End If
   '
   NCLIE% = XVALO(TRIM$(Text1.TEXT))
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
         Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 63, 2)
         Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIE%), RECORD$)
         Call CIERRARCH("DAPIBCLI")
         '
       End If
     End If
   End If
   Command22.Enabled = True
End Sub

Private Sub Command26_Click()
  
  '\\\OT-07-0010-RG-12/01/07///
  Command26.Enabled = False
  NC% = XVALO(Text1)
  If List1.ListIndex >= 0 Then
     VDEF$ = Space$(224)
     SUC$ = List1.List(List1.ListIndex)
     NUMSUC% = XVALO(Mid(SUC$, 191, 4))
     NUTRAI% = XVALO(Mid$(SUC$, 195, 4))
     NUZONA$ = Mid$(SUC$, 199, 4)
     If NUMSUC% = 0 Then NUMSUC% = List1.ListIndex + 1
     RODR$ = Mid$(SUC$, 190)
     VDEF$ = MKI$(NUMSUC%) + Mid$(SUC$, 1, 190) + MKI$(NUTRAI%) + NUZONA$
     VALANT = MODIFIC%(3): MODIFIC%(3) = 0
10   X$ = OBJPLA$("SUCENTRE", VDEF$)
     MODIFIC%(3) = VALANT
    'Estructura de objeto Sucentre
        'Numero de sucursal en  1 , 2
     '
     If X$ <> "" Then
        XX1$ = TRIM$(Mid$(X$, 3, 62)) 'Sucursal
        XX2$ = TRIM$(Mid$(X$, 65, 48)) 'Domicilio
        XX3$ = TRIM$(Mid$(X$, 113, 48)) 'Localidad
        XX4$ = TRIM$(Mid$(X$, 161, 32)) 'Telefono
        XX5$ = Mid$(X$, 193, 2) 'Transporte
        XX6$ = Mid$(X$, 195, 8) 'Zona Reparto
        '
        If XX1$ = "" Or XX2$ = "" Or XX3$ = "" Or XX4$ = "" Then
           Call MENSERR(24, "Información Incompleta - Faltan Datos")
           VDEF$ = X$
           GoTo 10
        End If
        '
        URE& = ULTREG&("!SUCENTRE")
        For KKU& = 1 To URE&
            XX$ = REGLEIDO$("!SUCENTRE", KKU&)
            A% = CVI(Mid$(X$, 1, 2))
            B% = CVI(Mid$(XX$, 1, 2))
            If A% = B% Then
              Call REPLA(XX$, X$, 1, 158)
              Call REPLA(XX$, Mid(X$, 161, 32), 163, 32)
              Call REPLA(XX$, XX5$, 195, 2)
              Call REPLA(XX$, XX6$, 197, 8)
              Call GRAREG("!SUCENTRE", XX$, KKU&)
              Exit For
            End If
        Next KKU&
        '
        Call CIERRARCH("!SUCENTRE")
        '
        Call FILTERUPDATE("SUCENTRE", "", 5, MKI$(NC%))
        '
        List1.Clear
        URE& = ULTREG&("!SUCENTRE")
        For KKU& = 1 To URE&
          NUMSUC% = CVI(Mid$(REGLEIDO$("!SUCENTRE", KKU&), 1, 2))
          NUTRAI% = CVI(Mid$(REGLEIDO$("!SUCENTRE", KKU&), 195, 2))
          NUZONA$ = Mid$(REGLEIDO$("!SUCENTRE", KKU&), 197, 8)
          NUMESUCURS$ = FORMATNUM(NUMSUC%, "I4")
          TTXX2$ = REGLEIDO$("!SUCENTRE", KKU&)
          NUZONA$ = TRIM$(Mid$(TTXX2, 197, 8))
          TTXX$ = Mid$(TTXX2$, 3, 158) + Mid$(TTXX2$, 163, 32) + NUMESUCURS$ + FORMATNUM(NUTRAI, "I4") + NUZONA$
          List1.AddItem TTXX$
        Next KKU&
        List1.ListIndex = 0
        Call List1_Click
     End If
  End If
  Command26.Enabled = True

  '
  '\\\OT-07-0010-RG-FIN///
End Sub


Private Sub Command27_Click()
   If List1.ListIndex >= 0 Then
     Command26_Click
   Else
     Call CommSel("ZONAREP", Command27, Text37)
   End If
End Sub

Private Sub Command28_Click()
    '\\\OT-08-0815-09/09/08///
    '
    ' PARA RECUPERAR LOS DATOS DE LOS CLIENTES DE METALMEDICA QUE  QUEDARON
    ' TODOS IMPORTADOS CON CODI-CLI = "130000"
    '
    Call APERBASDAT("CLIEN")
    SQLX$ = "SELECT * FROM DEUDOR12 "
    SQLX$ = SQLX$ + "WHERE Codi_Cli = '130001'"
    Set CLISEL = Badaclie.OpenRecordset(SQLX$, dbOpenDynaset)
    With CLISEL
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          On Error GoTo 0
          Do While Not .EOF
             NCLINUE% = !Nume_Cli
             CODCUENTA$ = "CLI-" + TRIM$(Str$(NCLINUE%))
             .Edit
             !Codi_Cli = CODCUENTA$
             .Update
          .MoveNext
          Loop
       End If
       On Error GoTo 0
    End With
    '
    Call COMUNI("Proceso Completado")
    '
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
    'Call AMACLIS04.Command3_Click
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
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      Frame2.BackColor = &HFCD7B6
      Check1.BackColor = &HFCD7B6
    End If

    SSTab1.Tab = 0
    Call CENTRAFORM(Me)
    Call TRANSLABELS(Name)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
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
   Call GRADAENTREGA
   GRABATODO% = 0
End Sub

Private Sub Command13_Click()
   AMACLIS04.Show 1
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If DERACCE%("STPCLIEN", "Configuración de Funcionamiento ABM Clientes") >= 2 Then
    OPMASCL04.Show 1
   End If
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
   Call GRADAENTREGA
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
     Call APERBASDAT("CLIEN")
     On Error Resume Next
     With Deudor12
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           For KKI% = 1 To .Fields.Count
             On Error Resume Next
             VATEX$ = ""
             VATEX$ = .Fields(KKI% - 1).Value
             On Error GoTo 0
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
   Call APERBASDAT("CLIEN")
   '
   Command21.Enabled = False
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
   '
   Screen.MousePointer = 11
   PRONUCLI& = XVALO(CONTROL$("", "PRONUCLI"))
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
   NCLINUE% = CVI(Mid$(OBX$, 13, 2))
   If NCLINUE% < 1 Or NCLINUE% > 29999 Then
      Call MENSERR(24, "Número de Cuenta no Válido.")
      GoTo 10
   End If
   '
   COCLINU$ = TRIM$(Left$(OBX$, 12))
   If COCLINU$ = "" Then
     COCLINU$ = "CLI-" + TRIM$(Str$(NCLINUE%))
     Call REPLA(VDEF$, COCLINU$, 1, 12)
     'Call MENSERR(24, "Codigo de Cuenta no Válido")
     'GoTo 10
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
   If PIVCLI% < 0 Or PIVCLI% > 7 Then
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
   Call APERBASDAT("CLIEN")
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
   Call GRADAENTREGA
   '
   If DERACCE%("BAJACLI", "Baja de Cuenta de Cliente") < 2 Then
     GoTo 99
   End If
   '
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE& = XVALO(Text1.TEXT)
       Call APERBASDAT("CLIEN")
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
'            RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
'            If RFF$ <> "" Then
'              Call MENSERR(24, "Baja Imposible - Cuenta con Movimiento.")
'              GoTo 99
'            End If
'            '
'            RFF$ = RECFILT$("COBRAN", 1, MKI$(NC%))
'            If RFF$ <> "" Then
'              Call MENSERR(24, "Baja Imposible - Cuenta con Movimiento.")
'              GoTo 99
'            End If
            
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
       GoTo 99
     End If
   End If
   '
   ' restauración de cuentas de baja
   Call GENCLIANU
   If ULTREG&("INDICLIA") < 1 Then
      Call MENSERR(24, "No Hay Cuentas de Clientes Anuladas\que Puedan Restaurarse.")
      GoTo 99
   End If
   Call FRESHECHO("INDICLIA")
   RPA$ = REGSEL$("INDICLIA/INDICE DE CLIENTES DE BAJA POR CODIGO")
   If Len(RPA$) > 4 Then
      NCLIE1% = CVI(Mid$(RPA$, 63, 2))
      If NCLIE1% > 0 Then
         NC% = NCLIE1%
         Text1.TEXT = TRIM$(Str$(NC%))
         Call Command1_Click
      End If
   End If
   '
99 Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Sub GENCLIANU()
   '
   Call COPYSTRU("INDIPROA", "INDICLIA")
   Call APERBASDAT("CLIEN")
   SQLX$ = "SELECT * FROM DEUDOR12 "
   SQLX$ = SQLX$ + " WHERE Stat_Cli = -1 "
   Set APRO1TMP = Badaclie.OpenRecordset(SQLX$, 4)
   With APRO1TMP
     On Error Resume Next
     .MoveFirst
     If Err Then
         On Error GoTo 0
         GoTo 99
     End If
     On Error GoTo 0
     REBLB$ = REGBLAN("INDICLIA")
     JJ& = 0
     Do While .EOF = False
       EL2$ = REBLB$
       Call REPLA(EL2$, !Codi_Cli, 1, 12)
       Call REPLA(EL2$, !Raso_Cli, 13, 50)
       Call REPLA(EL2$, MKI$(!Nume_Cli), 63, 2)
       JJ& = JJ& + 1
       Call GRAREG("INDICLIA", EL2$, JJ&)
       .MoveNext
      Loop
      Call SETULTREG("INDICLIA", JJ&)
      Call BAJALDISCO
      Call FILESORT("INDICLIA", "", 1, 1, "ASC")
      Call FRESHECHO("INDICLIA")
   '
99 End With
   '
End Sub
'
Private Sub Command25_Click()
    PRINTFORM
End Sub

Private Sub Command3_Click()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
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
   Call GRADAENTREGA
   '
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% < 1 Then Exit Sub
   '
   NC% = NCLIE%
   Text1.TEXT = TRIM$(Str$(NC%))
   'Call CARDATCLIE
   '
End Sub

Private Sub Frame1_DblClick()
   'Call List1_DblClick
End Sub

Private Sub LiCuInCli_Click()
    'Call AMACLIS04.Command5_Click
End Sub

Private Sub LiGeClaCli_Click()
' Call AMACLIS04.Command8_Click
End Sub

Private Sub List1_Click()
   TTXX$ = List1.TEXT
   Text43 = TRIM$(Mid$(TTXX$, 63, 48))
   Text44 = TRIM$(Mid$(TTXX$, 111, 48))
   Text37 = TRIM$(Mid$(TTXX$, 199, 8))
   If TRIM$(Mid$(TTXX$, 195, 4)) <> "" Then
      Text45 = TRIM$(Mid$(TTXX$, 195, 4))
   End If
   MODIFIC(3) = 1
End Sub

Private Sub List1_DblClick()
   '\\\OT-07-0010-RG-12/01/07///
   '
   Call Command26_Click
  '
  '\\\OT-07-0010-RG-FIN///
'***********************************************************************************
  
'   NC% = XVALO(Text1)
'   If NC% > 0 Then
'     If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
'10     VTB% = VENTABU%("SUCENTRE")
'       If VTB% >= 0 Then
'
'         URE& = ULTREG&("!SUCENTRE")
'         For KKU& = 1 To URE&
'           XX$ = REGLEIDO$("!SUCENTRE", KKU&)
'           XX1$ = TRIM$(Mid$(XX$, 3, 62))
'           XX2$ = TRIM$(Mid$(XX$, 65, 48))
'           XX3$ = TRIM$(Mid$(XX$, 113, 48))
'           If XX1$ = "" Or XX2$ = "" Or XX3$ = "" Then
'             Call MENSERR(24, "Información Incompleta - Faltan Datos")
'             GoTo 10
'           End If
'           KKI% = KKU&
'           Call REPLA(XX$, MKI$(KKI%), 1, 2)
'           Call GRAREG("!SUCENTRE", XX$, KKU&)
'         Next KKU&
'         Call CIERRARCH("!SUCENTRE")
'
'         Call FILTERUPDATE("SUCENTRE", "", 5, MKI$(NC%))
'
'         List1.Clear
'         If URE& > 0 Then
'           For KKU& = 1 To URE&
'
'             TTXX$ = Mid$(REGLEIDO$("!SUCENTRE", KKU&), 3, 160)
'             List1.AddItem TTXX$
'           Next KKU&
'           List1.ListIndex = 0
'           Call List1_Click
'         End If
'
'       End If
'     End If
'   End If
End Sub

Private Sub mnuibrutos_Click()
   Call Command3_Click
End Sub

Private Sub mnuimpresora1_Click()
    AMACLIS04.Option2 = True
    Call AMACLIS04.Command8_Click
End Sub

Private Sub Mnuimpresora3_Click()
   AMACLIS04.Option2 = True
   Call AMACLIS04.Command3_Click
End Sub

Private Sub mnuimpresora4_Click()
    AMACLIS04.Option2 = True
    Call AMACLIS04.Command4_Click
End Sub

Private Sub mnuimpresora5_Click()
    AMACLIS04.Option2 = True
    Call AMACLIS04.Command2_Click
End Sub

Private Sub mnupantalla1_Click()
    AMACLIS04.Option1(0) = True
    Call AMACLIS04.Command8_Click
End Sub

Private Sub mnuPantalla2_Click()
    AMACLIS04.Option1(0) = True
    Call AMACLIS04.Command5_Click
End Sub

Private Sub mnupantalla3_Click()
   AMACLIS04.Option1(0) = True
   Call AMACLIS04.Command3_Click
End Sub

Private Sub mnupantalla4_Click()
    AMACLIS04.Option1(0) = True
    Call AMACLIS04.Command4_Click
End Sub

Private Sub mnupantalla5_Click()
   AMACLIS04.Option1(0) = True
   Call AMACLIS04.Command2_Click
End Sub

Private Sub munuimpresora2_Click()
    AMACLIS04.Option2 = True
    Call AMACLIS04.Command5_Click

End Sub

Private Sub SuspOper_Click()
    Call Command3_Click
End Sub

Private Sub Text2_Change()
   Text4 = ""
   If XVALO(TRIM$(Text1)) > 0 Then
     NCLIE% = XVALO(TRIM$(Text2))
     If NCLIE% > 0 Then
       Text4 = RASOCLI(NCLIE%)
     End If
   End If
End Sub

Private Sub Text2_DblClick()
   Call Command11_Click
End Sub

Private Sub Text21_GotFocus()
  Static PASO%
  If PASO% = 1 Then Exit Sub
  PASO% = 1
  If DERACCE%("DESCSTAN", "Modificar Descuento Standard") < 2 Then
   PASO% = 0
   Text1.SetFocus
   End If
End Sub

Private Sub Text35_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text37_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text37_GotFocus()
   If ULTREG&("ZONAREP") > 0 Then
     Call MENSERR(24, "Campo no Editable")
     Command27.SetFocus
   End If
End Sub

Private Sub Text43_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text43_GotFocus()
   If List1.ListCount > 0 Then
     Call MENSERR(24, "Campo no Editable")
     Text45.SetFocus
   End If
End Sub

Private Sub Text44_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text44_GotFocus()
   If List1.ListCount > 0 Then
     Call MENSERR(24, "Campo no Editable")
     Text45.SetFocus
   End If
End Sub

Private Sub Text45_Change()
  '
  Label60 = ""
  Label61 = ""
  Label62 = ""
  Label63 = ""
  Label64 = ""
  Label65 = ""
  Label66 = ""
  Label67 = ""
  Label69 = ""
  '
  NTRA = XVALO(Text45)
  If NTRA > 0 Then
    If NTRA < 32767 Then
      NTRAI% = NTRA
      If ECOARCH$("TRANSPOR", MKI$(NTRAI%)) <> "" Then
        XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRAI%))
        Label60 = TRIM$(Mid$(XX$, 179, 15))
        Label61 = TRIM$(Mid$(XX$, 3, 32))
        Label62 = TRIM$(Mid$(XX$, 35, 48))
        Label63 = TRIM$(Mid$(XX$, 83, 48))
        Label64 = TRIM$(Mid$(XX$, 131, 48))
        Label65 = TRIM$(Mid$(XX$, 195, 48))
        Label66 = TRIM$(Mid$(XX$, 243, 48))
        Label67 = TRIM$(Mid$(XX$, 291, 48))
        Label69 = TRIM$(Mid$(XX$, 339, 48))
        '
        Text5 = Label61
        '
' ESTO QUE SIGUE SE HIZO PARA ACTUALIZAR EL TRANSPORTE POR SUCURSAL, PERO NO
' FUNCIONA CORRECTAMENTE POR ESO ESTÁ DESHABILITADO PARA REVISAR - EPB 26/11/08
'        If List1.ListIndex >= 0 Then
'          SUC$ = List1.List(List1.ListIndex)
'          NUMSUC% = XVALO(Mid(SUC$, 191, 4))
'          NCLIE% = XVALO(Text1)
'          URE& = ULTREG&("!SUCENTRE")
'          List1.Clear
'          For KKU& = 1 To URE&
'            TTXX2$ = REGLEIDO$("!SUCENTRE", KKU&)
'            If CVI(Mid$(TTXX2$, 1, 2)) = NUMSUC% Then
'              Call REPLA(TTXX2$, MKI$(NTRAI%), 195, 2)
'              Call GRAREG("!SUCENTRE", TTXX2$, KKU&)
'            End If
'            NUTRAI% = CVI(Mid$(TTXX2$, 195, 2))
'            ZONAREPA$ = TRIM$(Mid$(TTXX2, 197, 8))
'            TTXX1$ = FORMATNUM$(NUMSUC%, "I4")
'            TTXX$ = Mid$(TTXX2$, 3, 158) + Mid$(TTXX2$, 163, 32) + TTXX1$ + FORMATNUM(NUTRAI%, "I4") + ZONAREPA$
'            List1.AddItem TTXX$
'          Next KKU&
'          Call FILTERUPDATE("SUCURSAL", "", 5, MKI$(NCLIE%))
'          Call CIERRARCH("!SUCENTRE")
'        End If
      End If
    End If
  End If
  '
  MODIFIC%(3) = 1
  '
End Sub

Private Sub Text46_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text45_GotFocus()
   If List1.ListIndex > 0 Then
     Call MENSERR(24, "Imposible Editar")
     On Error Resume Next
     Command26.SetFocus
     On Error GoTo 0
   End If
End Sub

Private Sub Text5_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text5_GotFocus()
   If XVALO(Text45) > 0 Then
     Call MENSERR(24, "Imposible Editar")
     On Error Resume Next
     Text45.SetFocus
     On Error GoTo 0
   End If
End Sub

Private Sub Text6_Change()
   MODIFIC%(0) = 1
   If NUMECLI%(Text6) > 0 Then
     Text1 = TRIM$(Str$(NUMECLI%(Text6)))
   End If
End Sub

Private Sub Text6_DBLCLICK()
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
   Call GRADAENTREGA
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
    NCLIE% = XVALO(TRIM$(Text1.TEXT))
    STCL% = STATCLI%(NCLIE)
    If NCLIE < 1 Or Abs(STCL%) < 1 Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text3.TEXT = RASOCLI$(NCLIE)
    Text24.TEXT = TEL2CLI$(NCLIE)
    Text23.TEXT = NPROCLI$(NCLIE)
    Check1.Value = 0
    If TRIM$(Text24) <> "" Then
      If TRIM$(ECOARCH$("DESTFAB", Text24)) <> "" Then
        Check1.Value = 1
      End If
    End If
    MODIFIC(0) = 0
    '
    Select Case SSTab1.Tab
      Case 0
        Call CARGENCLI
      Case 1
        Call CARNOTACLI
      Case 2
        Call CARESCUENTA
      Case 3
        Call CARDAENTREGA
    End Select
    '
    Screen.MousePointer = 1
    '
End Sub
'
Sub CARGENCLI()
    '
    If YACAR%(0) > 0 Then Exit Sub
    NC% = XVALO(TRIM$(Text1.TEXT))
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
    Text8.TEXT = PAGWEBCLI$(NC%)
    Text8.Refresh
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
    NCLIE& = XVALO(TRIM$(Text1.TEXT))
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
          Call APERBASDAT("CLIEN")
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
          Text3 = RASOCLI$(NCLIE&)
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
      PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
      If PIVA% <> 5 Then  ' EXCLUYE COMERCIO EXTERIOR
        If VALICUIT%(CUIT$) <> 1 Then
          Screen.MousePointer = 1
          Call MENSERR(24, "Imposible Registrar !!!\Número de C.U.I.T. no Válido")
          SSTab1.Tab = 0
          Text33.SetFocus
          Exit Sub
        End If
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
    Call APERBASDAT("CLIEN")
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
      ABC1$ = Left$(!Codi_Cli, 12)
      If Err Then
        ABC1$ = ""
      End If
      On Error GoTo 0
      .Edit
        If TRIM$(Text3) <> TRIM$(!Raso_Cli) Or UCase$(Text6) <> ABC1$ Then UPDATCLI% = 1
      !Codi_Cli = Left$(UCase$(Text6), 12)
      !Raso_Cli = Left$(Text3, 64)
      !Stat_Cli = 1
      !Domi_Cli = Left$(Text25, 64)
      !Cpos_Cli = Left$(Text26, 16)
      !Loca_Cli = Left$(Text27, 48)
      !Tele_Cli = Left$(Text30, 32)
      !Fant_Cli = Left$(Text24, 32)
      !Cuit_Cli = Left$(TRIM$(CUIT$), 24)
      !Nibr_Cli = Left$(TRIM$(Text34), 24)
      !Caib_Cli = Left$(TRIM$(COTEXT(2)), 2)
      !Prov_Cli = Left$(TRIM$(COTEXT(0)), 2)
      '
      MODLIMCRE% = 0
      On Error Resume Next
      LMCRA = CDbl(!Lcre_Cli)
      If Err Then
        LMCRA = 0: !Lcre_Cli = 0
      End If
      On Error GoTo 0
      '
      If Abs(CDbl(!Lcre_Cli) - XVALO(Text20.TEXT)) > 0.05 Or InStr(UCase$(Text20.TEXT), "SUSPEN") > 0 Then
        If DERACCE%("MODLICRE", "MODIFICACION LIMITE DE CREDITO") >= 1 Then
            !Lcre_Cli = XVALO(Text20)
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
      !Gru_Div = XVALO(Text2)
      '
      CCX% = REGBUS&("CUENTAS", 1, AJUSTI$(COTEXT(6).TEXT, 12))
      !CuMa_Cli = CCX%
      '
      DCC$ = TRIM$(Text21.TEXT)
      If DCC$ <> "" Then
         If EXISTE%(LUDAT$ + "DESCOMBI.RFS") > 0 Then
            If ULTREG&("DESCOMBI") > 0 Then
               Call MENSERR(24, "Descuento Standard por Cliente\no Compatible con Descuentos Combinados.\   \Aplique Descuento Combinado a esta Cuenta.")
               DCC$ = ""
            End If
         End If
      End If
      !Dsta_Cli = Left$(TRIM$(DCC$), 32)
      '
      !Pais_Cli = Left$(Text29, 48)
      !Faxi_Cli = Left$(Text31, 32)
      !Mail_Cli = Left$(Text32, 64)
      !PagWeb = Left$(Text8, 64)
      !Ctac_Cli = Left$(Text19, 48)
      !Prvd_Cli = Left$(Text23, 24)
      !Hora_Cli = Left$(Text36, 32)
      '
      .Update
    End With
    '
    If Check1.Value = 1 Then
        If TRIM$(Text24) <> "" Then
          If ULTREG&("DESTFAB") < 2 Then Call VERIDEST
          NFANCLI$ = AJUSTI$(UCase$(Text24), 8)
          RASOCIA$ = AJUSTI$(Text3, 24)
          Call FILTERPUTRECORD("DESTFAB", 1, TRIM$(NFANCLI$), NFANCLI$ + RASOCIA$)
        End If
      Else
        NFANCLI$ = AJUSTI$(UCase$(Text24), 8)
        Call FILTERDELETE("DESTFAB", 1, NFANCLI$)
    End If
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
   NCLIE% = XVALO(Text1.TEXT)
   Text28 = ""
   '
   If NCLIE% > 0 Then
     If REGICLI(NCLIE%) > 0 Then
       Call APERBASDAT("CLIEN")
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
   NCLIE% = XVALO(TRIM$(Text1.TEXT))
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
   Call APERBASDAT("CLIEN")
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
Sub CARDAENTREGA()
   '
   If YACAR%(3) > 0 Then Exit Sub
   NCLIE% = XVALO(Text1.TEXT)
   Text45 = ""
   Text43 = ""
   Text44 = ""
   Text5 = ""
   Text35 = ""
   Text37 = ""
   Text46 = ""
   List1.Clear
   '
   If NCLIE% > 0 Then
      Text45 = TRIM$(Str$(NUTRANCLI%(NCLIE%)))
      If TRIM$(Text45) = "0" Then Text45 = ""
      Text43 = TRIM$(DOMIENT$(NCLIE%))
      Text44 = TRIM$(LOCAENT$(NCLIE%))
      Text5 = TRIM$(TRANENT$(NCLIE%))
      Text35 = TRIM$(HORAENT$(NCLIE%))
      Text37 = TRIM$(ZONACLI$(NCLIE%))
      Text46 = OBSENTRE$(NCLIE%)
      Text7 = FORMATNUM$(POREMBAL(NCLIE%), "F5.1")
      '
      Call FILTERFILE("SUCENTRE", "", 5, MKI$(NCLIE%))
      URE& = ULTREG&("!SUCENTRE")
      For KKU& = 1 To URE&
'          NUMSUCU$ = Mid$(REGLEIDO$("!SUCENTRE", KKU&), 1, 2)
'          SUCURSAL$ = Mid$(REGLEIDO$("!SUCENTRE", KKU&), 3, 62)
'          DOMICILIO$ = Mid$(REGLEIDO$("!SUCENTRE", KKU&), 65, 48)
'          LOCALIDAD$ = Mid$(REGLEIDO$("!SUCENTRE", KKU&), 113, 48)
'          TELEFONO$ = Mid$(REGLEIDO$("!SUCENTRE", KKU&), 163, 32)
'          NUTRANS$ = Mid$(REGLEIDO$("!SUCENTRE", KKU&), 195, 2)
'          NUZONA$ = Mid$(REGLEIDO$("!SUCENTRE", KKU&), 197, 8)
'          TTXX$ = SUCURSAL$ + DOMICILIO$ & LOCALIDAD$ & TELEFONO$ & NUTRANS$ & NUZONA$ & NUMSUC%

        NUMSUC% = CVI(Mid$(REGLEIDO$("!SUCENTRE", KKU&), 1, 2))
        TTXX1$ = FORMATNUM$(NUMSUC%, "I4")
        TTXX2$ = REGLEIDO$("!SUCENTRE", KKU&)
        NUTRAI% = CVI(Mid$(TTXX2$, 195, 2))
        ZONAREPA$ = TRIM$(Mid$(TTXX2, 197, 8))
        TTXX$ = Mid$(TTXX2$, 3, 158) + Mid$(TTXX2$, 163, 32) + TTXX1$ + FORMATNUM(NUTRAI, "I4") + ZONAREPA$
        List1.AddItem TTXX$
      Next KKU&
      If URE& > 0 Then
        List1.ListIndex = 0
        Call List1_Click
      End If
      '
   End If
   '
   YACAR%(3) = 1: MODIFIC%(3) = 0
   '
End Sub
'

Sub GRADAENTREGA()
   '
   NCLIE% = XVALO(TRIM$(Text1.TEXT))
   If NCLIE% < 1 Then Exit Sub
   '
   If REGICLI(NCLIE%) < 1 Then Exit Sub
   If YACAR%(3) < 1 Then Exit Sub
   If MODIFIC%(3) < 1 Then Exit Sub
   '
   If GRABATODO% <> 1 Then
     'Command9.SetFocus
     If COMALTER%("Se han Modificado los Datos de Entrega\del Cliente '" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(3) = 0
       Call CARDAENTREGA
       Exit Sub
     End If
     On Error Resume Next
     Command9.SetFocus
     On Error GoTo 0
   End If
   '
   ' LO QUE SIGUE ES PARA QUE EL LUGAR DE ENTREGA
   ' SEA EL DE LA PRIMERA SUCURSAL
   If List1.ListCount > 0 Then
     List1.ListIndex = 0
     Call List1_Click
   End If
   '
   Call APERBASDAT("CLIEN")
   Screen.MousePointer = 11
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
     !Lentre0_Cli = Left$(Text43, 48)
     !Lentre1_Cli = Left$(Text44, 48)
     !Lentre2_Cli = Left$(Text5, 48)
     If Abs(XVALO(Text45)) < 32767 Then !NTrans_Cli = XVALO(Text45)
     !Hora_Ent = Left$(Text35, 48)
     !Zona_Cli = Left$(Text37, 8)
     !Obser_Entre = Text46
     !Recar_Embal = XVALO(Text7)
     .Update
   End With
   Deudor12.Close
   MODIFIC%(3) = 0
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
   Check1.Value = 0
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
   Label60 = ""
   Label61 = ""
   Label62 = ""
   Label63 = ""
   Label64 = ""
   Label65 = ""
   Label66 = ""
   Label67 = ""
   Label69 = ""
   '
   Text45 = ""
   Text43 = ""
   Text44 = ""
   Text5 = ""
   Text35 = ""
   Text37 = ""
   Text46 = ""
   Text7 = ""
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
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
   If TIDOCUM$ = "Recibo Cobranza" Then TIDOCUM$ = "Recibo"
   If TIDOCUM$ = "Ajuste de Saldo" Then TIDOCUM$ = "Ajuste Saldo Cliente"
   NUDOCU$ = " " + TRIM$(Mid$(RELIS$, 27, 8)) + " "
   If InStr(UCase$(TIDOCUM$), "FACTURA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Then
     NUDOCU$ = TRIM$(NUDOCU$)
     While Len(NUDOCU$) < 8: NUDOCU$ = "0" + NUDOCU$: Wend
   End If
   RACLI$ = TRIM$(Left$(RASOCLI$(XVALO(Text1.TEXT)), 8))
   FECHADOC$ = Mid$(RELIS$, 1, 5)
   Screen.MousePointer = 11
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
     If InStr(XX$, TIDOCUM$) > 4 Then
       If InStr(XX$, NUDOCU$) > 4 Then
         'If InStr(XX$, RACLI$) > 16 Or VUELBUS% = 2 Then
         'If InStr(XX$, FECHADOC$) > 0 Then
           DOCUNU& = CVS(Left$(XX$, 4))
           GoTo 5
         'End If
       End If
     End If
   Next UI&
   '
'   If VUELBUS% < 2 Then ' COMENTADO PARA QUE TRAIGA SIEMPRE DONDE COINCIDA NUMERO DE DOCUMENTO
'     VUELBUS% = 2       ' Y NUMERO DE CLIENTE
'     GoTo 3
'   End If
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
'   EDITFORM.Command5.Enabled = False
   EDITFORM.Show 1
   '
End Sub

Private Sub LTEXT36_Change()
   DISPONI = XVALO(LTEXT36.Caption) - XVALO(LTEXT37.Caption) - XVALO(LTEXT35.Caption)
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
      Call GRADAENTREGA
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
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE = XVALO(Text1.TEXT)
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
   Call GRADAENTREGA
   '
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
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
   Call GRADAENTREGA
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
       LTEXT36.Caption = CSTRING$(MKS$(XVALO(Text20.TEXT)), 4, 11, 2, 2)
   End If
End Sub

Private Sub Text20_LostFocus()
  If TRIM$(UCase$(Text20.TEXT)) <> "SUSPENDIDO" Then
    Text20.TEXT = CSTRING$(MKS$(XVALO(Text20.TEXT)), 3, 12, 2, 2)
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
'\\\OT-06-0429-MH-11/10/06///
Private Sub text32_DblClick()
   If InStr(Text32.TEXT, "@") <> 0 Then
      FORMAI04.Text28 = ""
      FORMAI04.Text9 = ""
      FORMAI04.Label3.Caption = Text32.TEXT
      FORMAI04.Show 1
   End If
End Sub
'\\\OT-06-0429-MH-FIN///

Private Sub Text33_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text34_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub LText35_DblClick()
   Call RECUEC04.Text6_DBLCLICK
End Sub
'
Sub CARESCUENTA()
    '
    If DERACCE%("VERCCTE", "Consulta De Cuenta Corriente") < 1 Then
      Exit Sub
    End If
    '
    If YACAR%(2) > 0 Then Exit Sub
    '
50  Screen.MousePointer = 11
    NC% = XVALO(Text1.TEXT)
    If NC% > 0 Then
        List3.Clear
        List3.Refresh
        List2.Clear
        List2.Refresh
        LTEXT23.Caption = ""
        LTEXT37.Caption = ""
        LTEXT36.Caption = ""
        LTEXT35.Caption = ""
        LMCR = LIMICRE(NC%)
        LTEXT36.Caption = CSTRING$(MKS$(LMCR), 4, 11, 2, 2)
        If LMCR < 0 Then LTEXT36.Caption = " Suspendido"
        Refresh
    End If
    '
100 RECUEC04.Text1 = Text1
    List2.Clear: List3.Clear
    With RECUEC04
      For U& = 1 To .List1.ListCount
        List3.AddItem .List1.List(U& - 1)
      Next U&
      '
      For U& = 1 To .List2.ListCount
        List2.AddItem .List2.List(U& - 1)
      Next U&
      '
      Label25 = .Label6
      Label28 = .Label5
      LTEXT37 = .Text10
      LTEXT35 = .Text6
    End With
    '
    DISPONI = XVALO(LTEXT36.Caption) - XVALO(LTEXT37.Caption) - XVALO(LTEXT35.Caption)
    LTEXT23.ForeColor = &H80000008
    If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
    LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
    '
    YACAR%(2) = 1
    Screen.MousePointer = 1
    '
End Sub
'
Private Sub Text36_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text7_Change()
   MODIFIC(3) = 1
End Sub

Private Sub Text7_LostFocus()
   Text7 = FORMATNUM$(Text7, "F5.1")
End Sub

Private Sub Text8_DblClick()

If TRIM$(Text8) <> "" Then Call ShellExecute(Me.hwnd, "Open", TRIM$(Text8), "", "", 1)

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
      Call APERBASDAT("CLIEN")
      Set CLIENTECO = Badaclie.OpenRecordset("SELECT Codi_Cli, Nume_Cli,Raso_Cli FROM Deudor12 WHERE Stat_Cli > 0 ORDER BY Raso_Cli ASC;", dbOpenDynaset, dbReadOnly)
      On Error Resume Next
      CLIENTECO.MoveLast
      If Err Then Exit Sub
      '
      On Error GoTo 0
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
10      If .EOF Then GoTo 20
        Call TRACE(20, JJ&, FIN&)
        If !Nume_Cli >= 0.5 Then
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
        End If
        '
        If Not .EOF Then
          .MoveNext
          GoTo 10
        End If
      End With
      '
20    CLIENTECO.Close
      Call SETULTREG("DEUDOR1", JJ&)
      Call CIERRARCH("DEUDOR1")
      Call SETULTREG("INDICLIE", JJ&)
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
    Call REPLA(REBLA$, MKI$(Int(HORANUM(Time$))), 63, 2)
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

Sub VERIDEST()
    '
    STO1% = 0: STO2% = 0
    For I& = 1 To ULTREG&("DESTFAB")
       X$ = REGLEIDO$("DESTFAB", I&)
       CODEST$ = TRIM$(UCase$(Left$(X$, 8)))
       If CODEST$ = "STOCK" Then STO1% = 1
       If CODEST$ = "SEMIELA" Then STO2% = 1
    Next I&
    '
    If STO1% = 0 Then
       X$ = REGBLAN$("DESTFAB")
       Call REPLA(X$, "STOCK", 1, 8)
       Call REPLA(X$, "Stock de Producto Terminado", 9, 32)
       Call REGAPP("DESTFAB", X$)
    End If
    '
    If STO2% = 0 Then
       X$ = REGBLAN$("DESTFAB")
       Call REPLA(X$, "SEMIELA", 1, 8)
       Call REPLA(X$, "Stock de Semielaborados", 9, 32)
       Call REGAPP("DESTFAB", X$)
    End If
    '
    Call CIERRARCH("DESTFAB")
    '
End Sub


