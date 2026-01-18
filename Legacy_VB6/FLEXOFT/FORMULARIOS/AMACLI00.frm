VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AMACLI00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes - A-B-M-L Base de Datos de Clientes"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "AMACLI00.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
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
      TabIndex        =   3
      Top             =   735
      Width           =   5710
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFC0&
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
      Left            =   5775
      TabIndex        =   138
      Top             =   210
      Width           =   1170
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
      TabIndex        =   98
      Top             =   735
      Width           =   1425
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10815
      ScaleHeight     =   75
      ScaleWidth      =   930
      TabIndex        =   95
      Top             =   4670
      Width           =   990
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   96
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "A-B-M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3350
      Left            =   10815
      TabIndex        =   50
      Top             =   1260
      Width           =   990
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
         Height          =   700
         Left            =   105
         Picture         =   "AMACLI00.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Ordena Base de Datos de Clientes"
         Top             =   2520
         Width           =   765
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Stop"
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
         Picture         =   "AMACLI00.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Suspensión de Operaciones"
         Top             =   1050
         Width           =   765
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
         Height          =   700
         Left            =   105
         Picture         =   "AMACLI00.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Alta de Nuevo Cliente"
         Top             =   315
         Width           =   750
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
         Height          =   700
         Left            =   105
         Picture         =   "AMACLI00.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Baja Lógica de la Cuenta"
         Top             =   1785
         Width           =   765
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Varios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3320
      Left            =   10815
      TabIndex        =   83
      Top             =   4800
      Width           =   990
      Begin VB.CommandButton Command99 
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
         Height          =   700
         Left            =   105
         Picture         =   "AMACLI00.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   2485
         Width           =   765
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
         Height          =   700
         Left            =   105
         Picture         =   "AMACLI00.frx":11B4
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   280
         Width           =   765
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
         Height          =   700
         Left            =   105
         Picture         =   "AMACLI00.frx":15F6
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Emisión de Listados"
         Top             =   1015
         Width           =   765
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
         Height          =   700
         Left            =   105
         Picture         =   "AMACLI00.frx":1900
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Imprime Ficha del Cliente"
         Top             =   1750
         Width           =   765
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   10710
      Sorted          =   -1  'True
      TabIndex        =   94
      Top             =   3780
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   8610
      Picture         =   "AMACLI00.frx":1F6A
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Guarda los Cambios Realizados"
      Top             =   300
      Width           =   750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   9360
      Picture         =   "AMACLI00.frx":2274
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Restablece los Valores Anteriores - Descarta Cambios Efectuados"
      Top             =   300
      Width           =   750
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   10290
      Picture         =   "AMACLI00.frx":257E
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Ayuda en Línea"
      Top             =   300
      Width           =   750
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
      Height          =   770
      Left            =   11040
      Picture         =   "AMACLI00.frx":2888
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   300
      Width           =   750
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
      Height          =   330
      Left            =   3570
      MaxLength       =   10
      TabIndex        =   2
      Text            =   " "
      Top             =   315
      Width           =   1935
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
      Picture         =   "AMACLI00.frx":29D2
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   150
      TabIndex        =   47
      Top             =   1365
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   11880
      _Version        =   327681
      TabHeight       =   520
      BackColor       =   16777152
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "AMACLI00.frx":2CDC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Anotaciones"
      TabPicture(1)   =   "AMACLI00.frx":2CF8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line3"
      Tab(1).Control(1)=   "Picture4"
      Tab(1).Control(2)=   "Command7"
      Tab(1).Control(3)=   "Command5"
      Tab(1).Control(4)=   "Command6"
      Tab(1).Control(5)=   "Command8"
      Tab(1).Control(6)=   "Command17"
      Tab(1).Control(7)=   "Text28"
      Tab(1).Control(8)=   "Picture2"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Entrega"
      TabPicture(2)   =   "AMACLI00.frx":2D14
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame23"
      Tab(2).Control(1)=   "Frame22"
      Tab(2).Control(2)=   "Frame20"
      Tab(2).ControlCount=   3
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
         Left            =   -74850
         TabIndex        =   114
         Top             =   3045
         Width           =   10200
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
            TabIndex        =   125
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
            TabIndex        =   116
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
               TabIndex        =   124
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
               TabIndex        =   123
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
               TabIndex        =   122
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
               TabIndex        =   121
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
               TabIndex        =   120
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
               TabIndex        =   119
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
               TabIndex        =   118
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
               TabIndex        =   117
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
            TabIndex        =   115
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
            TabIndex        =   137
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
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
            Left            =   5985
            TabIndex        =   136
            Top             =   525
            Width           =   855
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
            TabIndex        =   135
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
            TabIndex        =   134
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
            TabIndex        =   133
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
            TabIndex        =   132
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
            TabIndex        =   131
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
            TabIndex        =   130
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
            TabIndex        =   129
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
            TabIndex        =   128
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
            TabIndex        =   127
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
            TabIndex        =   126
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
         Left            =   -74850
         TabIndex        =   104
         Top             =   735
         Width           =   6105
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
            TabIndex        =   109
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
            TabIndex        =   108
            Text            =   " "
            Top             =   840
            Width           =   4425
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
            Left            =   1470
            TabIndex        =   107
            Top             =   1260
            Width           =   4320
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
            TabIndex        =   106
            Top             =   1680
            Width           =   4425
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
            Height          =   330
            Left            =   5730
            TabIndex        =   105
            Top             =   1260
            Width           =   175
         End
         Begin VB.Label Label57 
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
            TabIndex        =   113
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
            TabIndex        =   112
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
            TabIndex        =   111
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
            TabIndex        =   110
            Top             =   1785
            Width           =   1065
         End
      End
      Begin VB.Frame Frame23 
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
         Height          =   2220
         Left            =   -68655
         TabIndex        =   102
         Top             =   735
         Width           =   4005
         Begin VB.TextBox Text46 
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
            Height          =   1590
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   103
            Top             =   420
            Width           =   3730
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   -74650
         ScaleHeight     =   375
         ScaleWidth      =   9720
         TabIndex        =   84
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
            TabIndex        =   85
            Top             =   105
            Width           =   9570
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6105
         Left            =   105
         TabIndex        =   56
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
            TabIndex        =   80
            Top             =   4480
            Width           =   5055
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
               TabIndex        =   93
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
               TabIndex        =   92
               Top             =   1155
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
               TabIndex        =   82
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
               TabIndex        =   81
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
            TabIndex        =   74
            Top             =   1840
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
               TabIndex        =   91
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
               TabIndex        =   90
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
               TabIndex        =   89
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
               TabIndex        =   79
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
               TabIndex        =   78
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
               TabIndex        =   77
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
               TabIndex        =   76
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
               TabIndex        =   75
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
            TabIndex        =   69
            Top             =   4095
            Width           =   4845
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
               TabIndex        =   88
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
               TabIndex        =   87
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
               TabIndex        =   73
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
               TabIndex        =   72
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
               TabIndex        =   71
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
               TabIndex        =   70
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
            TabIndex        =   64
            Top             =   1840
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
               TabIndex        =   99
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
               TabIndex        =   100
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
               TabIndex        =   68
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
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
            TabIndex        =   58
            Top             =   105
            Width           =   10095
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
               TabIndex        =   4
               Text            =   " "
               Top             =   315
               Width           =   8550
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
               TabIndex        =   86
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
               TabIndex        =   63
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
               TabIndex        =   62
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
               TabIndex        =   61
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
               TabIndex        =   60
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
               TabIndex        =   59
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
         Left            =   -74650
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   1575
         Width           =   9780
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         Height          =   260
         Left            =   -69800
         TabIndex        =   55
         Top             =   5790
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         Height          =   260
         Left            =   -68615
         TabIndex        =   44
         Top             =   5790
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pegar"
         Height          =   260
         Left            =   -67445
         TabIndex        =   45
         Top             =   5790
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Left            =   -66155
         TabIndex        =   46
         Top             =   5790
         Width           =   1290
      End
      Begin VB.CommandButton Command7 
         Height          =   260
         Left            =   -74650
         TabIndex        =   54
         Top             =   5790
         Width           =   4830
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         Height          =   1590
         Left            =   -72690
         ScaleHeight     =   1530
         ScaleWidth      =   2055
         TabIndex        =   53
         Top             =   1875
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Line Line3 
         X1              =   -64875
         X2              =   -64875
         Y1              =   1155
         Y2              =   6000
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   3795
         TabIndex        =   52
         Top             =   4605
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   4215
         TabIndex        =   51
         Top             =   3660
         Width           =   15
      End
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
      TabIndex        =   97
      Tag             =   " "
      Top             =   290
      Width           =   1005
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
      Left            =   2310
      TabIndex        =   57
      Top             =   240
      Width           =   1110
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
      TabIndex        =   49
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
      TabIndex        =   48
      Top             =   240
      Width           =   1140
   End
End
Attribute VB_Name = "AMACLI00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%(8), MODIFIC%(8)
Dim GRABATODO%
Dim AGRECLI%
Dim RECONTANT$

Private Sub Check1_Click()
   If Check1.Value = 1 Then
     If TRIM$(Text24) = "" Then
       Text24 = TRIM$(Left$(Text3, 8))
     End If
   End If
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

'
Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    SSTab1.Tab = 0
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
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
   Command13.Enabled = False
   AMACLIST.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   'OPAMACLI.Show 1
   Call CONECRUN("FLSETUP/OPAMACLI", "Configuración de Funcionamiento")
   Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   Command15.Enabled = False
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
   Call CIERRARCH("DEUDOR1")
   Call CIERRARCH("DEUDOR2")
   Call SORTCLI(1)
   Call CIERRARCH("DEUDOR1")
   Call CIERRARCH("DEUDOR2")
   Command15.Enabled = True
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
     PROXI% = 1 + ULTREG&("TRANSPOR")
     Call REPLA(OBX$, MKI$(PROXI%), 1, 2)
     Call REGAPP("TRANSPOR", OBX$)
     Call CIERRARCH("TRANSPOR")
   End If
   '
99 Command16.Enabled = True
   '
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command18_Click()
   '
   Command18.Enabled = False
   '
   If DERACCE%("ATRANSPO", "Actualizar Base de Datos de Transportes") < 2 Then
      GoTo 99
   End If
   '
   URE& = ULTREG&("TRANSPOR")
   RACT& = Val(Label51)
   If RACT& < 1 Or RACT& > URE& Then GoTo 99
   '
   VDEF$ = REGLEIDO$("TRANSPOR", RACT&)
   '
10 OBX$ = OBJPLA$("CARGA-TRANSPOR", VDEF$)
   If OBX$ <> "" Then
      '
      If TRIM$(Mid$(OBX$, 3, 32)) = "" Then
        Call MENSERR(24, "Falta Nombre del Transporte")
        VDEF$ = OBX$
        GoTo 10
      End If
      '
      Call BLOQARCH("TRANSPOR")
      PROXI% = RACT&
      Call REPLA(OBX$, MKI$(PROXI%), 1, 2)
      Call GRAREG("TRANSPOR", OBX$, RACT&)
      Call CIERRARCH("TRANSPOR")
      '
      XX$ = OBX$
      Label51 = TRIM$(Str$(CVI(Left$(XX$, 2))))
      Label61 = TRIM$(Mid$(XX$, 3, 32))
      Text45 = TRIM$(Mid$(XX$, 3, 32))
      Label60 = TRIM$(Mid$(XX$, 179, 15))
      Label62 = TRIM$(Mid$(XX$, 35, 48))
      Label63 = TRIM$(Mid$(XX$, 83, 48))
      Label64 = TRIM$(Mid$(XX$, 131, 48))
      Label65 = TRIM$(Mid$(XX$, 195, 48))
      Label66 = TRIM$(Mid$(XX$, 243, 48))
      Label67 = TRIM$(Mid$(XX$, 291, 48))
      Label69 = TRIM$(Mid$(XX$, 339, 48))
   End If
   '
99 Command18.Enabled = True
   '
End Sub

Private Sub Command19_Click()
    NCLIE% = Val(TRIM$(Text1.TEXT))
    If NCLIE% < 1 Then Exit Sub
    If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) = "" Then Exit Sub
    '
    XX$ = REGSEL$("TRANSPOR")
    If Len(XX$) > 4 Then
      Label51 = TRIM$(Str$(CVI(Left$(XX$, 2))))
      Label61 = TRIM$(Mid$(XX$, 3, 32))
      Text45 = TRIM$(Mid$(XX$, 3, 32))
      Label60 = TRIM$(Mid$(XX$, 179, 15))
      Label62 = TRIM$(Mid$(XX$, 35, 48))
      Label63 = TRIM$(Mid$(XX$, 83, 48))
      Label64 = TRIM$(Mid$(XX$, 131, 48))
      Label65 = TRIM$(Mid$(XX$, 195, 48))
      Label66 = TRIM$(Mid$(XX$, 243, 48))
      Label67 = TRIM$(Mid$(XX$, 291, 48))
      Label69 = TRIM$(Mid$(XX$, 339, 48))
    End If
End Sub

Private Sub Command2_Click()
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
   For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call CARDATCLIE
End Sub

Private Sub COMMAND21_CLICK()
   '
   If DERACCE%("ALTACLI", "Alta de Nuevo Cliente") < 2 Then
      Exit Sub
   End If
   '
   Command21.Enabled = False
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
   '
   Screen.MousePointer = 11
   PRONUCLI& = Val(CONTROL$("", "PRONUCLI"))
   If PRONUCLI& > 9999 Then PRONUCLI& = 0
   NROPRO& = 0
   PCLI$ = String$(32767, 0)
   For U& = 1 To ULTREG&("DEUDOR1")
      X$ = REGLEIDO$("DEUDOR1", U&)
      NCXI% = CVI(Left$(X$, 2))
      If NCXI% > 0 Then
        If NCXI% < 9998 Then
          Mid$(PCLI$, NCXI%, 1) = Chr$(1)
          If NCXI% > NROPRO& Then
            NROPRO& = NCXI%
          End If
        End If
      End If
   Next U&
   NCLIE% = NROPRO& + 1
   '
   If PRONUCLI& >= 0 Then
      For KKII% = PRONUCLI& + 1 To 9999
        If Asc(Mid$(PCLI$, KKII%, 1)) < 1 Then
           NCLIE% = KKII%
           NROPRO& = NCLIE% - 1
           GoTo 5
        End If
      Next KKII%
   End If
   '
5  Screen.MousePointer = 1
   '
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(NCLIE%), 1, 2)
   Call REPLA(VDEF$, Space$(62), 3, 62)
   Call REPLA(VDEF$, Chr$(1), 65, 1)
   Call REPLA(VDEF$, Space$(15), 66, 15)
   '
10 Screen.MousePointer = 1
   Call CIERRARCH("DEUDOR1")
   Call CIERRARCH("DEUDOR2")
   Call CIERRARCH("BADACLIE")
   OBX$ = OBJPLA$("ALTACLIEN", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   NCLINUE% = CVI(Left$(OBX$, 2))
   '
   If NCLINUE% < 1 Or NCLINUE% > 9999 Then
      Call MENSERR(24, "Número de Cuenta no Válido.")
      GoTo 10
   End If
   '
   If RECFILT$("DEUDOR1", 1, MKI$(NCLINUE%)) <> "" Then
      Call MENSERR(24, "Ingreso no Válido.\Número de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   DCLINUE$ = TRIM$(Mid$(OBX$, 3, 62))
   If Len(DCLINUE$) < 1 Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Razón Social del Cliente")
      GoTo 10
   End If
   '
   PIVCLI% = Asc(Mid$(OBX$, 65, 1))
   If PIVCLI% < 0 Or PIVCLI% > 6 Then
      Call MENSERR(24, "Ingreso no Válido.\Posición IVA no Figura en Tablas.")
      GoTo 10
   End If
   '
   If PIVCLI% <> 0 Then
     If PIVCLI% <> 5 Then
       CUTT$ = TRIM$(Mid$(OBX$, 66, 15))
       If VALICUIT%(CUTT$) <> 1 Then
         Call MENSERR(24, "Ingreso no Válido.\Número de C.U.I.T. Incorrecto.")
         GoTo 10
       End If
     End If
   End If
   '
   Screen.MousePointer = 11
   ARX1$ = "DEUDOR1": ARX2$ = "DEUDOR2"
   Call BLOQARCH(ARX1$)
   Call BLOQARCH(ARX2$)
   Call BLOQARCH("BADACLIE")
   '
   CUTT$ = TRIM$(CUTT$)
   For U& = 1 To ULTREG&(ARX1$)
     XX$ = REGLEIDO$(ARX1$, U&) + REGLEIDO$(ARX2$, U&)
     If CVI(Left$(XX$, 2)) = NCLINUE% Then
       Call MENSERR(24, "Ingreso no Válido.\Número de Cuenta Pre-existente.")
       GoTo 10
     End If
     '
     If CUTT$ <> "" Then
       If CVI(Left$(XX$, 2)) > 0 Then
         If TRIM$(Mid$(XX$, 107, 15)) = CUTT$ Then
           NCLIX% = CVI(Left$(XX$, 2))
           If COMALTER%("Existe Registrada una Cuenta con el Mismo C.U.I.T.\  \'" + TRIM$(RASOCLI$(NCLIX%)) + "' (" + TRIM$(Str$(NCLIX%)) + ")\\Cancelar\Registrar Igual") <> 2 Then
             GoTo 10
           End If
         End If
       End If
     End If
   Next U&
   '
   REGX& = 1 + ULTREG&(ARX1$)
   XX$ = REGBLAN$(ARX1$) + String$(128, 0)
   YY$ = REGBLAN$("BADACLIE")
   Call REPLA(XX$, MKI$(NCLINUE%), 1, 2)
   Call REPLA(XX$, DCLINUE$, 3, 30)
   Call REPLA(XX$, TRIM$(CUTT$), 107, 15)
   Call REPLA(XX$, Chr$(PIVCLI%), 136, 1)
   '
   Call REPLA(YY$, MKI$(NCLINUE%), 1, 2)
   Call REPLA(YY$, DCLINUE$, 3, 62)
   '
   Call REGAPP(ARX1$, Left$(XX$, 32))
   REGX& = ULTREG&(ARX1$)
   RGBLX$ = REGBLAN$(ARX2$)
   While ULTREG&(ARX2$) < REGX&
     Call REGAPP(ARX2$, RGBLX$)
   Wend
   '
   Call GRAREG(ARX1$, Left$(XX$, 32), REGX&)
   Call GRAREG(ARX2$, Mid$(XX$, 33), REGX&)
   Call CIERRARCH(ARX1$)
   Call CIERRARCH(ARX2$)
   '
   RFF$ = RECFILT$("BADACLIE", 1, MKI$(NCLINUE%))
   If Len(RFF$) >= 4 Then
     REBADA& = CVS(Left$(RFF$, 4))
     If REBADA& > 0 Then
       If REBADA& <= ULTREG&("BADACLIE") Then
         Call GRAREG("BADACLIE", YY$, REBADA&)
         GoTo 29
       End If
     End If
   End If
   Call REGAPP("BADACLIE", YY$)
   '
29 Call CIERRARCH("BADACLIE")
   Call BAJALDISCO
   '
   AGRECLI% = 1
   Text1.TEXT = ""
   Call BLANFORMU
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
             'DETEXT(Index).TabStop = False
          Next Index
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
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE% = Val(Text1.TEXT)
       If REGICLI&(NCLIE%) > 0 Then
         NC% = NCLIE%
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
           YACAR%(1) = 1
           Text28.TEXT = ""
           MODIFIC%(1) = 1
           GRABATODO% = 1
           Call GRANOTACLI
           '
           REG& = REGICLI&(NCLIE%)
           XX$ = REGLEIDO$("DEUDOR1", REG&)
           Call COPYSTRU("DEUDOR1", "DEUDOBAJ")
           Call REGAPP("DEUDOBAJ", XX$)
           Call CIERRARCH("DEUDOBAJ")
           '
           Call REPLA(XX$, MKI$(0), 1, 2)
           Call REPLA(XX$, Space$(30), 3, 30)
           Call GRAREG("DEUDOR1", XX$, REG&)
           Call CIERRARCH("DEUDOR1")
           '
           AGRECLI% = 1
           Text1.TEXT = ""
           Call BLANFORMU
           For KKI% = 0 To 4: MODIFIC%(KKI%) = 0: YACAR%(KKI%) = 0: Next KKI%
           GRABATODO% = 0
           Call BAJALDISCO
           '
         End If
         '
       End If
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
   Call GRADAENTREGA
   Call CARDATCLIE
   '
   Call CIERRARCH("*.*")
   If AGRECLI% > 0 Then
      Call SORTCLI(1)
      AGRECLI% = 0
   End If
   Call BAJALDISCO
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
   Call FRESHECHO("DEUDOR1")
   Call SELECHO("DEUDOR1")
   NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% < 1 Then Exit Sub
   '
   NC% = NCLIE%
   Text1.TEXT = TRIM$(Str$(NC%))
   'Call CARDATCLIE
   '
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

Sub CARDATCLIE()
    '
    NCLIE% = Val(TRIM$(Text1.TEXT))
    'Call CIERRARCH("DEUDOR1")
    'Call FRESHECHO("DEUDOR1")
    If NCLIE% < 1 Or ECOARCH$("DEUDOR1", MKI$(NCLIE%)) = "" Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text3.TEXT = RASOCLI$(NCLIE%)
    Text24.TEXT = TEL2CLI$(NCLIE%)
    Text23.TEXT = NPROCLI$(NCLIE%)
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
        Call CARDAENTREGA
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
    YACAR%(0) = 1: MODIFIC%(0) = 0
    '
End Sub
'
Sub GRAGENCLI()
    '
    NCLIE% = Val(TRIM$(Text1.TEXT))
    If NCLIE% < 1 Then
      If GRABATODO% = 1 Then
        Call MENSERR(24, "Imposible Registrar !!!\Número de Cliente no Válido.")
        Exit Sub
      End If
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
        Text3.TEXT = RASOCLI$(NCLIE%)
        Text24.TEXT = TEL2CLI$(NCLIE%)
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
    If NCLIE% = 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar Cuenta Número '0'")
        Exit Sub
    End If
    '
    If TRIM$(Text3.TEXT) = "" Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Falta Razón Social del Cliente.")
        Exit Sub
    End If
    '
    ARX1$ = "DEUDOR1": ARX2$ = "DEUDOR2"
    REGX& = REGICLI&(NCLIE%)
    XX11$ = REGLEIDO$("DEUDOR1", REGX&)
    If Abs(CVI(Left$(XX11$, 2))) <> Abs(NCLIE%) Then
       REGX& = 0
    End If
    '
    If REGX& < 1 Or REGX& > ULTREG&(ARX1$) Then
       REGX& = 0
       For UUL& = 1 To ULTREG&("DEUDOR1")
         XX11$ = REGLEIDO$("DEUDOR1", UUL&)
         If Abs(CVI(Left$(XX11$, 2))) = NCLIE% Then
           REGX& = UUL&
           GoTo 20
         End If
       Next UUL&
       REGX& = 1 + ULTREG&("DEUDOR1")
       XX11$ = REGBLAN$(ARX1$)
       Call REPLA(XX11$, MKI$(NCLIE%), 1, 2)
       Call REPLA(XX11$, TRIM$(Text3.TEXT), 3, 30)
       Call GRAREG(ARX1$, Left$(XX11$, 32), REGX&)
       RGBLX$ = REGBLAN$(ARX2$)
       While ULTREG&(ARX2$) < REGX&
         Call REGAPP(ARX2$, RGBLX$)
       Wend
       Call CIERRARCH(ARX1$): Call CIERRARCH(ARX2$)
       Call BAJALDISCO
    End If
    '
    If TRIM$(Text24) <> "" Then
      If ULTREG&("DESTFAB") < 2 Then Call VERIDEST
      RFF$ = RECFILT$("DESTFAB", 1, TRIM$(Text24))
      REDEST& = 1 + ULTREG&("DESTFAB"): REGIX$ = REGBLAN$("DESTFAB")
      If Len(RFF$) >= 4 Then
        REDEST& = CVS(Left$(RFF$, 4))
        REGIX$ = REGLEIDO$("DESTFAB", REDEST&)
      End If
      Call REPLA(REGIX$, Text24, 1, 8)
      Call REPLA(REGIX$, Text3, 9, 32)
      Call GRAREG("DESTFAB", REGIX$, REDEST&)
      Call CIERRARCH("DESTFAB")
    End If
    '
20  XX$ = REGLEIDO$(ARX1$, REGX&) + REGLEIDO$(ARX2$, REGX&)
    If Abs(CVI(Left$(XX$, 2))) <> Abs(NCLIE%) Then
       Screen.MousePointer = 1
       Call MENSERR(24, "Imposible Registrar.\Consulte Soporte de Sistemas.")
       Exit Sub
    End If
    '
    CARASO% = 0
    If TRIM$(Text3.TEXT) <> TRIM$(RASOCLI$(NCLIE%)) Then
       CARASO% = 1
    End If
    '
    If TRIM$(Left$(Text3.TEXT, 30)) <> TRIM$(Mid$(XX$, 3, 30)) Then
       AGRECLI% = 1
    End If
    '
    Call REPLA(XX$, TRIM$(Text3.TEXT), 3, 30)
    Call REPLA(XX$, TRIM$(Text25.TEXT), 33, 25)
    CPOST% = 0
    If Val(Text26.TEXT) > 0 Then
      If Val(Text26.TEXT) < 32768 Then
        CPOST% = Val(Text26.TEXT)
      End If
    End If
    Call REPLA(XX$, MKI$(CPOST%), 60, 2)
    Call REPLA(XX$, TRIM$(Text27.TEXT), 62, 25)
    Call REPLA(XX$, TRIM$(Text30.TEXT), 87, 10)
    Call REPLA(XX$, TRIM$(Text24.TEXT), 97, 10)
    Call REPLA(XX$, TRIM$(CUIT$), 107, 15)
    Call REPLA(XX$, TRIM$(Text34.TEXT), 122, 13)
    Call REPLA(XX$, TRIM$(COTEXT(2).TEXT), 135, 1)
    Call REPLA(XX$, TRIM$(COTEXT(0).TEXT), 149, 1)
    MODLIMCRE% = 0
    If Abs(CVS(Mid$(XX$, 153, 4)) - Val(Text20.TEXT)) > 0.05 Or InStr(UCase$(Text20.TEXT), "SUSPEN") > 0 Then
      If DERACCE%("MODLICRE", "MODIFICACION LIMITE DE CREDITO") >= 1 Then
          Call REPLA(XX$, MKS$(Val(Text20.TEXT)), 153, 4)
          If TRIM$(UCase$(Text20.TEXT)) = "SUSPENDIDO" Then
             Call REPLA(XX$, MKS$(-1&), 153, 4)
          End If
        Else
          LMCR = LIMICRE(NCLIE%)
          Text20.TEXT = CSTRING$(MKS$(LMCR), 3, 12, 2, 2)
          If LMCR < 0 Then Text20.TEXT = " Suspendido"
          Text20.Refresh
          MODLIMCRE% = 1
      End If
    End If
    '
    PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
    Call REPLA(XX$, Chr$(PIVA%), 136, 1)
    '
    CATIB% = VALIMI(COTEXT(2).TEXT, 0, 2)
    '
    LIPRE% = VALIMI(COTEXT(3).TEXT, 0, 255)
    If LIPRE% > 0 Then
      If ECOARCH("LISTAS", Chr$(LIPRE%)) = "" Then
        Call MENSERR(24, "Lista de Precios no Válida")
        LIPRE% = 0
      End If
    End If
    Call REPLA(XX$, Chr$(LIPRE%), 151, 1)
    '
    COPAGI% = VALIMI(COTEXT(4).TEXT, 0, 255)
    If COPAGI% > 0 Then
      If ECOARCH("CONPAG", Chr$(COPAGI%)) = "" Then
        Call MENSERR(24, "Condición de Pago no Válida")
        COPAGI% = 0
      End If
    End If
    Call REPLA(XX$, Chr$(COPAGI%), 150, 1)
    '
    VENDE% = VALIMI(COTEXT(5).TEXT, 0, 255)
    If VENDE% > 0 Then
      If ECOARCH("VENDEDOR", Chr$(VENDE%)) = "" Then
        Call MENSERR(24, "Vendedor no Válido")
        VENDE% = 0
      End If
    End If
    Call REPLA(XX$, Chr$(VENDE%), 152, 1)
    '
    CUECO% = VALIMI(COTEXT(7).TEXT, 0, 32767)
    If ECOARCH$("GRUCOCLI", MKI$(CUECO%)) = "" Then
      CUECO% = 0
    End If
    Call REPLA(XX$, MKI$(CUECO%), 157, 2)
    '
    CCX% = REGBUS&("CUENTAS", 1, AJUSTI$(COTEXT(6).TEXT, 12))
    Call REPLA(XX$, MKI$(CCX%), 159, 2)
    '
    Call GRAREG(ARX1$, Left$(XX$, 32), REGX&)
    Call GRAREG(ARX2$, Mid$(XX$, 33), REGX&)
    Call CIERRARCH(ARX1$)
    Call CIERRARCH(ARX2$)
    '
    VDL& = CVI(Mid$(XX$, 159, 2))
    If TRIM$(Mid$(XX$, 33, 25)) <> "" Then
      If VDL& > 0 Then
        For KK& = 1 To ULTREG&("BKDATCLI")
          XXYY$ = REGLEIDO$("BKDATCLI", KK&)
          If CVI(Left$(XXYY$, 2)) = NCLIE% Then
            Call GRAREG("BKDATCLI", XX$, KK&)
            GoTo 23
          End If
        Next KK&
        Call REGAPP("BKDATCLI", XX$)
23      Call CIERRARCH("BKDATCLI")
      End If
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    DCC$ = TRIM$(Text21.TEXT)
    For UJ& = 1 To ULTREG&("DESCUCLI")
      YYY$ = REGLEIDO$("DESCUCLI", UJ&)
      If CVI(Left$(YYY$, 2)) = NCLIE% Then
         Call REPLA(YYY$, DCC$, 3, 14)
         Call GRAREG("DESCUCLI", YYY$, UJ&)
         GoTo 99
      End If
    Next UJ&
    YYY$ = REGBLAN$("DESCUCLI")
    Call REPLA(YYY$, MKI$(NCLIE%), 1, 2)
    Call REPLA(YYY$, DCC$, 3, 14)
    Call REGAPP("DESCUCLI", YYY$)
    '
99  Call CIERRARCH("DESCUCLI")
    '
    If NCLIE% > 0 Then
       RFF$ = RECFILT$("BADACLIE", 1, MKI$(NCLIE%))
       REBACLI& = 0
       If Len(RFF$) >= 4 Then
          REBACLI& = CVS(Left$(RFF$, 4))
       End If
       If REBACLI& < 1 Or REBACLI& > ULTREG&("BADACLIE") Then
         REBACLI& = 0
         For UUL& = 1 To ULTREG&("BADACLIE")
           XX11$ = REGLEIDO$("BADACLIE", UUL&)
           If Abs(CVI(Left$(XX11$, 2))) = NCLIE% Then
             REBACLI& = UUL&
             GoTo 120
           End If
         Next UUL&
         REBACLI& = 1 + ULTREG&("BADACLIE")
       End If
       '
120    XX$ = REGBLAN$("BADACLIE")
       Call REPLA(XX$, MKI$(NCLIE%), 1, 2)
       Call REPLA(XX$, Text3.TEXT, 3, 62)
       Call REPLA(XX$, Text24.TEXT, 65, 32)
       Call REPLA(XX$, Text25.TEXT, 97, 64)
       Call REPLA(XX$, Text26.TEXT, 161, 16)
       Call REPLA(XX$, Text27.TEXT, 177, 48)
       Call REPLA(XX$, Text29.TEXT, 225, 32)
       Call REPLA(XX$, Text30.TEXT, 257, 24)
       Call REPLA(XX$, Text31.TEXT, 281, 24)
       Call REPLA(XX$, Text32.TEXT, 305, 48)
       Call REPLA(XX$, Text19.TEXT, 385, 48)
       Call REPLA(XX$, Text23.TEXT, 433, 16)
       Call REPLA(XX$, Text36.TEXT, 449, 32)
       Call GRAREG("BADACLIE", XX$, REBACLI&)
       Call CIERRARCH("BADACLIE")
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
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
   Text28.TEXT = ""
   '
   If NCLIE% > 0 Then
      If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("NOTACLI", 1, MKI$(NCLIE%))
         URN& = ULTREG&("NOTACLI")
         '
         For U& = 1 To Len(RNC$) Step 4
            RENOTA& = CVS(Mid$(RNC$, U&, 4))
            If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                  TTYY$ = RTrim$(Mid$(REGLEIDO$("NOTACLI", RENOTA&), 3))
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
   YACAR%(1) = 1: MODIFIC%(1) = 0
   '
End Sub
'
Sub GRANOTACLI()
   '
   NCLIE% = Val(TRIM$(Text1.TEXT))
   If NCLIE% < 1 Then Exit Sub
   If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) = "" Then Exit Sub
   If YACAR%(1) < 1 Then Exit Sub
   '
   If NCLIE% > 0 Then
      If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         '
         If MODIFIC%(1) < 1 Then
           Exit Sub
         End If
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
         Screen.MousePointer = 11
         URN& = ULTREG&("NOTACLI")
         RNC$ = RECFILT$("NOTACLI", 1, MKI$(NCLIE%))
         KK& = (-3)
         '
         TEXTO$ = Text28.TEXT + Chr$(13) + Chr$(10)
         Call FRATEXTO(TEXTO$, 78)
         For LL% = 1 To CARETEX%
            TTYY$ = MKI$(NCLIE%) + RETEX$(LL%)
            KK& = KK& + 4
            If KK& <= Len(RNC$) - 3 Then
                RENOTA& = CVS(Mid$(RNC$, KK&, 4))
              Else
                URN& = URN& + 1
                RENOTA& = URN&
            End If
            Call GRAREG("NOTACLI", TTYY$, RENOTA&)
         Next LL%
         '
20       While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Call GRAREG("NOTACLI", MKI$(0) + Space$(78), RENOTA&)
         Wend
         '
         Call CIERRARCH("NOTACLI")
         MODIFIC%(1) = 0
         Screen.MousePointer = 1
      End If
   End If

End Sub
'
Sub BLANFORMU()
   '
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
   Hide
   '
End Sub

'
Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab <> PreviousTab Then
      Call GRAGENCLI
      Call GRANOTACLI
      Call GRADAENTREGA
      Call FRESHALL
      Call CARDATCLIE
   End If
End Sub

Private Sub Text1_Change()
   For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call BLANFORMU
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE% = Val(Text1.TEXT)
       If REGICLI&(NCLIE%) > 0 Then
         NC% = NCLIE%
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
   Call GRADAENTREGA
   '
End Sub

Private Sub Text19_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text20_Change()
   MODIFIC%(0) = 1
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

Private Sub Text35_Change()
   MODIFIC%(4) = 1
End Sub

Private Sub Text43_Change()
   MODIFIC%(4) = 1
End Sub

Private Sub Text44_Change()
   MODIFIC%(4) = 1
End Sub

Private Sub Text45_Change()
   MODIFIC%(4) = 1
   If TRIM$(Text45) <> TRIM$(Label61) Then
      '
      Label51 = ""
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
   End If
End Sub

Private Sub Text46_GotFocus()
   On Error Resume Next
   Text43.SetFocus
   On Error GoTo 0
End Sub

Sub CARDAENTREGA()
   '
   If YACAR%(4) > 0 Then Exit Sub
   NCLIE% = Val(Text1.TEXT)
   Call BLADAENTREGA
   '
   If NCLIE% > 0 Then
      If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         Screen.MousePointer = 11
         RFF$ = RECFILT$("LENTRECL", 1, MKI$(NCLIE%))
         If Len(RFF$) > 0 Then
           REG& = CVS(Left$(RFF$, 4))
           If REG& > 0 Then
             If REG& <= ULTREG&("LENTRECL") Then
               XX$ = REGLEIDO$("LENTRECL", REG&)
               NTRAN& = CVI(Mid$(XX$, 99, 2))
               Text43 = Mid$(XX$, 3, 32)
               Text44 = Mid$(XX$, 35, 32)
               Text45 = Mid$(XX$, 67, 32)
               Text35 = Mid$(XX$, 101, 24)
               '
               If NTRAN& > 0 Then
                 If NTRAN& <= ULTREG&("TRANSPOR") Then
                   XX$ = REGLEIDO$("TRANSPOR", NTRAN&)
                   Label51 = TRIM$(Str$(CVI(Left$(XX$, 2))))
                   Label61 = TRIM$(Mid$(XX$, 3, 32))
                   Text45 = TRIM$(Mid$(XX$, 3, 32))
                   Label60 = TRIM$(Mid$(XX$, 179, 15))
                   Label62 = TRIM$(Mid$(XX$, 35, 48))
                   Label63 = TRIM$(Mid$(XX$, 83, 48))
                   Label64 = TRIM$(Mid$(XX$, 131, 48))
                   Label65 = TRIM$(Mid$(XX$, 195, 48))
                   Label66 = TRIM$(Mid$(XX$, 243, 48))
                   Label67 = TRIM$(Mid$(XX$, 291, 48))
                   Label69 = TRIM$(Mid$(XX$, 339, 48))
                 End If
               End If
             End If
           End If
         End If
      End If
   End If
   '
   YACAR%(4) = 1: MODIFIC%(4) = 0
   Screen.MousePointer = 1
   '
End Sub

Sub BLADAENTREGA()
   '
   Text43 = ""
   Text44 = ""
   Text45 = ""
   Text35 = ""
   '
   Label51 = ""
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
End Sub
'
Sub GRADAENTREGA()
   '
   NCLIE% = Val(TRIM$(Text1.TEXT))
   If NCLIE% < 1 Then Exit Sub
   If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) = "" Then Exit Sub
   If YACAR%(4) < 1 Then Exit Sub
   '
   If NCLIE% > 0 Then
      If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         '
         If MODIFIC%(4) < 1 Then
           Exit Sub
         End If
         '
         If DERACCE%("DAENTREG", "Actualizar Datos de Entrega Clientes") < 2 Then
            Exit Sub
         End If
         '
         If GRABATODO% <> 1 Then
           Command9.SetFocus
           If COMALTER%("Se han Modificado los Datos de Entrega del Cliente\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
             YACAR%(4) = 0
             Call CARDAENTREGA
             Exit Sub
           End If
         End If
         '
         Screen.MousePointer = 11
         URN& = ULTREG&("LENTRECL")
         RFF$ = RECFILT$("LENTRECL", 1, MKI$(NCLIE%))
         RELENT& = 1 + URN&
         If Len(RFF$) > 0 Then
           REG& = CVS(Left$(RFF$, 4))
           If REG& > 0 Then
             If REG& <= URN& Then
               RELENT& = REG&
             End If
           End If
         End If
         XX$ = REGBLAN$("LENTRECL")
         Call REPLA(XX$, MKI$(NCLIE%), 1, 2)
         Call REPLA(XX$, TRIM$(Text43), 3, 32)
         Call REPLA(XX$, TRIM$(Text44), 35, 32)
         Call REPLA(XX$, TRIM$(Text45), 67, 32)
         Call REPLA(XX$, MKI$(Val(Label51)), 99, 2)
         Call REPLA(XX$, TRIM$(Text35), 101, 24)
         Call GRAREG("LENTRECL", XX$, RELENT&)
         '
      End If
   End If
   '
   Call CIERRARCH("LENTRECL")
   MODIFIC%(4) = 0
   Screen.MousePointer = 1
   '
End Sub
'

