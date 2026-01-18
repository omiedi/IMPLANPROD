VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AMACLIEN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes - A-B-M-L Base de Datos de Clientes"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "AMACLIEN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11910
   Begin VB.CommandButton Command26 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   205
      Top             =   210
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command20 
      Height          =   770
      Left            =   8085
      Picture         =   "AMACLIEN.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   203
      ToolTipText     =   "Búsqueda por CUIT, Domicilio, Nombre de Fantasía, etc"
      Top             =   300
      Width           =   540
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
      Left            =   6510
      TabIndex        =   162
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
         Picture         =   "AMACLIEN.frx":17A8
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
         Picture         =   "AMACLIEN.frx":1AB2
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
         Picture         =   "AMACLIEN.frx":1DBC
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
         Picture         =   "AMACLIEN.frx":1F06
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
         Picture         =   "AMACLIEN.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   202
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
         Picture         =   "AMACLIEN.frx":251A
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
         Picture         =   "AMACLIEN.frx":295C
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
         Picture         =   "AMACLIEN.frx":2C66
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
      Left            =   8715
      Picture         =   "AMACLIEN.frx":32D0
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
      Left            =   9465
      Picture         =   "AMACLIEN.frx":35DA
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
      Picture         =   "AMACLIEN.frx":38E4
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
      Picture         =   "AMACLIEN.frx":3BEE
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
      TabIndex        =   2
      Text            =   " "
      Top             =   315
      Width           =   2805
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
      TabIndex        =   3
      Top             =   735
      Width           =   5235
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
      Picture         =   "AMACLIEN.frx":3D38
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
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   16777152
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "AMACLIEN.frx":4042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Anotaciones"
      TabPicture(1)   =   "AMACLIEN.frx":405E
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
      TabPicture(2)   =   "AMACLIEN.frx":407A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame18"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Contactos"
      TabPicture(3)   =   "AMACLIEN.frx":4096
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Entrega"
      TabPicture(4)   =   "AMACLIEN.frx":40B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame19"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame19 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5790
         Left            =   -74895
         TabIndex        =   163
         Top             =   735
         Width           =   10305
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
            Left            =   6255
            TabIndex        =   186
            Top             =   0
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
               TabIndex        =   187
               Top             =   420
               Width           =   3730
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
            Left            =   60
            TabIndex        =   177
            Top             =   0
            Width           =   6105
            Begin VB.TextBox Text37 
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
               Left            =   4830
               TabIndex        =   206
               Top             =   1680
               Width           =   1065
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
               TabIndex        =   181
               Top             =   1260
               Width           =   175
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
               TabIndex        =   182
               Top             =   1680
               Width           =   2955
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
               TabIndex        =   180
               Top             =   1260
               Width           =   4320
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
               TabIndex        =   179
               Text            =   " "
               Top             =   840
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
               TabIndex        =   178
               Text            =   " "
               Top             =   420
               Width           =   4425
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
               Left            =   3700
               TabIndex        =   207
               Top             =   1785
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
               TabIndex        =   197
               Top             =   1785
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
               TabIndex        =   185
               Top             =   525
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
               TabIndex        =   184
               Top             =   945
               Width           =   1065
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
               TabIndex        =   183
               Top             =   1365
               Width           =   1065
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
            Left            =   60
            TabIndex        =   164
            Top             =   2310
            Width           =   10200
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
               TabIndex        =   188
               ToolTipText     =   "Modificar Datos del Transporte Activo"
               Top             =   525
               Width           =   1065
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
               TabIndex        =   166
               Top             =   1890
               Width           =   9830
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
                  TabIndex        =   199
                  Top             =   945
                  Width           =   645
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
                  TabIndex        =   198
                  Top             =   840
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
                  TabIndex        =   196
                  Top             =   840
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
                  TabIndex        =   195
                  Top             =   420
                  Width           =   3375
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
                  TabIndex        =   194
                  Top             =   420
                  Width           =   4215
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
                  TabIndex        =   169
                  Top             =   525
                  Width           =   1065
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
                  TabIndex        =   168
                  Top             =   525
                  Width           =   645
               End
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
                  TabIndex        =   167
                  Top             =   945
                  Width           =   645
               End
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
               TabIndex        =   165
               ToolTipText     =   "Agrega Nuevo Transporte a la Base de Datos"
               Top             =   525
               Width           =   1050
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
               TabIndex        =   193
               Top             =   1365
               Width           =   3375
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
               TabIndex        =   192
               Top             =   945
               Width           =   2115
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
               TabIndex        =   191
               Top             =   1365
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
               TabIndex        =   190
               Top             =   945
               Width           =   5370
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
               TabIndex        =   189
               Top             =   525
               Width           =   3900
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
               TabIndex        =   176
               Top             =   1050
               Width           =   1485
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
               TabIndex        =   175
               Top             =   1470
               Width           =   1275
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
               TabIndex        =   174
               Top             =   1470
               Width           =   1275
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
               TabIndex        =   173
               Top             =   630
               Width           =   1065
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
               TabIndex        =   172
               Top             =   630
               Width           =   540
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
               TabIndex        =   171
               Top             =   525
               Width           =   855
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
               TabIndex        =   170
               Top             =   1050
               Width           =   1275
            End
         End
         Begin VB.Line Line22 
            X1              =   4620
            X2              =   2100
            Y1              =   3225
            Y2              =   3225
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6105
         Left            =   -74895
         TabIndex        =   119
         Top             =   525
         Width           =   10305
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
            Height          =   330
            Left            =   1365
            MaxLength       =   48
            TabIndex        =   156
            Text            =   " "
            Top             =   210
            Width           =   3795
         End
         Begin VB.Frame Frame2 
            Caption         =   "Empresa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Left            =   105
            TabIndex        =   146
            Top             =   735
            Width           =   7365
            Begin VB.TextBox Text2 
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
               MaxLength       =   48
               TabIndex        =   151
               Top             =   315
               Width           =   5895
            End
            Begin VB.TextBox Text4 
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
               MaxLength       =   48
               TabIndex        =   150
               Text            =   " "
               Top             =   735
               Width           =   5895
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
               Left            =   1260
               MaxLength       =   8
               TabIndex        =   149
               Text            =   " "
               Top             =   1155
               Width           =   1275
            End
            Begin VB.TextBox Text6 
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
               Left            =   2730
               MaxLength       =   48
               TabIndex        =   148
               Text            =   " "
               Top             =   1155
               Width           =   4425
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
               Height          =   330
               Left            =   1260
               TabIndex        =   147
               Text            =   " "
               Top             =   1575
               Width           =   5895
            End
            Begin VB.Label Label6 
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
               Height          =   225
               Left            =   315
               TabIndex        =   155
               Top             =   1680
               Width           =   855
            End
            Begin VB.Label Label5 
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
               Left            =   105
               TabIndex        =   154
               Top             =   1260
               Width           =   1065
            End
            Begin VB.Label Label4 
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
               Left            =   105
               TabIndex        =   153
               Top             =   840
               Width           =   1065
            End
            Begin VB.Label label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "R. Social:"
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
               Left            =   210
               TabIndex        =   152
               Top             =   420
               Width           =   960
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Teléfono"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1275
            Left            =   105
            TabIndex        =   135
            Top             =   2940
            Width           =   7365
            Begin VB.TextBox Text8 
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
               TabIndex        =   141
               Text            =   " "
               Top             =   315
               Width           =   2115
            End
            Begin VB.TextBox Text9 
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
               Left            =   3885
               TabIndex        =   140
               Text            =   " "
               Top             =   315
               Width           =   750
            End
            Begin VB.TextBox Text10 
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
               Left            =   5460
               TabIndex        =   139
               Text            =   " "
               Top             =   315
               Width           =   1695
            End
            Begin VB.TextBox Text11 
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
               TabIndex        =   138
               Text            =   " "
               Top             =   735
               Width           =   2115
            End
            Begin VB.TextBox Text12 
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
               Left            =   3885
               TabIndex        =   137
               Text            =   " "
               Top             =   735
               Width           =   750
            End
            Begin VB.TextBox Text13 
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
               Left            =   5460
               TabIndex        =   136
               Text            =   " "
               Top             =   735
               Width           =   1695
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Comercial:"
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
               TabIndex        =   145
               Top             =   420
               Width           =   1065
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Int:"
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
               Left            =   3150
               TabIndex        =   144
               Top             =   420
               Width           =   645
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Celular:"
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
               Left            =   4725
               TabIndex        =   143
               Top             =   840
               Width           =   645
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Partic.:"
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
               Left            =   4620
               TabIndex        =   142
               Top             =   420
               Width           =   750
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Fax - E-Mail"
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
            TabIndex        =   128
            Top             =   4305
            Width           =   5370
            Begin VB.TextBox Text17 
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
               Left            =   735
               TabIndex        =   132
               Text            =   " "
               Top             =   1155
               Width           =   4425
            End
            Begin VB.TextBox Text15 
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
               Left            =   735
               TabIndex        =   131
               Text            =   " "
               Top             =   735
               Width           =   4425
            End
            Begin VB.TextBox Text14 
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
               Left            =   735
               TabIndex        =   130
               Text            =   " "
               Top             =   315
               Width           =   2115
            End
            Begin VB.TextBox Text16 
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
               Left            =   3045
               TabIndex        =   129
               Text            =   " "
               Top             =   315
               Width           =   2115
            End
            Begin VB.Label Label45 
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
               Left            =   0
               TabIndex        =   134
               Top             =   840
               Width           =   645
            End
            Begin VB.Label Label44 
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
               Left            =   0
               TabIndex        =   133
               Top             =   420
               Width           =   645
            End
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Nuevo Contacto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   370
            Left            =   7665
            Style           =   1  'Graphical
            TabIndex        =   127
            Top             =   3840
            Width           =   2535
         End
         Begin VB.Frame Frame8 
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
            Height          =   1695
            Left            =   5565
            TabIndex        =   125
            Top             =   4305
            Width           =   4635
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
               Height          =   1170
               Left            =   105
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   126
               Top             =   315
               Width           =   4425
            End
         End
         Begin VB.PictureBox Picture1 
            Height          =   435
            Left            =   7665
            ScaleHeight     =   375
            ScaleWidth      =   2460
            TabIndex        =   123
            Top             =   210
            Width           =   2515
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               Caption         =   "Contactos Asociados"
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
               TabIndex        =   124
               Top             =   105
               Width           =   2220
            End
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   370
            Left            =   8925
            Style           =   1  'Graphical
            TabIndex        =   122
            Top             =   3450
            Width           =   1275
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   370
            Left            =   7665
            Style           =   1  'Graphical
            TabIndex        =   121
            Top             =   3450
            Width           =   1260
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2790
            ItemData        =   "AMACLIEN.frx":40CE
            Left            =   7665
            List            =   "AMACLIEN.frx":40D0
            TabIndex        =   120
            Top             =   630
            Width           =   2535
         End
         Begin VB.Label REGCONTA 
            Height          =   330
            Left            =   0
            TabIndex        =   160
            Top             =   0
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
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
            Left            =   420
            TabIndex        =   159
            Top             =   315
            Width           =   1065
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.:"
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
            Left            =   5985
            TabIndex        =   158
            Top             =   315
            Width           =   540
         End
         Begin VB.Label Label7 
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
            Left            =   6615
            TabIndex        =   157
            Top             =   210
            Width           =   855
         End
         Begin VB.Line Line2 
            X1              =   4620
            X2              =   2100
            Y1              =   3225
            Y2              =   3225
         End
         Begin VB.Line Line1 
            X1              =   10185
            X2              =   10185
            Y1              =   210
            Y2              =   630
         End
      End
      Begin VB.Frame Frame18 
         BorderStyle     =   0  'None
         Caption         =   "Frame18"
         Height          =   6000
         Left            =   -74685
         TabIndex        =   97
         Top             =   525
         Width           =   9990
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00E0E0E0&
            Height          =   2800
            Left            =   105
            ScaleHeight     =   2745
            ScaleWidth      =   9615
            TabIndex        =   103
            Top             =   3120
            Width           =   9675
            Begin VB.Frame Frame14 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   5630
               Left            =   0
               TabIndex        =   106
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
                  TabIndex        =   107
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
                  TabIndex        =   110
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
                  TabIndex        =   109
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
                  TabIndex        =   108
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
               TabIndex        =   104
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
                  TabIndex        =   105
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
            TabIndex        =   98
            Top             =   525
            Width           =   9675
            Begin VB.Frame Frame16 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               Height          =   330
               Left            =   0
               TabIndex        =   99
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
                  TabIndex        =   100
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
               TabIndex        =   101
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
                  TabIndex        =   102
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
            TabIndex        =   118
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
            TabIndex        =   117
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
            TabIndex        =   116
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
            TabIndex        =   115
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
            TabIndex        =   114
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
            TabIndex        =   113
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
            TabIndex        =   112
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
            TabIndex        =   111
            Top             =   105
            Width           =   1500
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
               TabIndex        =   204
               ToolTipText     =   "Datos Percepción I.Brutos Cliente Activo"
               Top             =   1575
               Visible         =   0   'False
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
               TabIndex        =   200
               Text            =   " "
               Top             =   1755
               Width           =   3300
            End
            Begin VB.CommandButton Command27 
               Caption         =   "Users"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   3780
               TabIndex        =   208
               Top             =   1755
               Width           =   855
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
               TabIndex        =   201
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
      Left            =   6720
      TabIndex        =   161
      Tag             =   " "
      Top             =   285
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
Attribute VB_Name = "AMACLIEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%(8), MODIFIC%(8)
Dim GRABATODO%
Dim AGRECLI%
Dim RECONTANT$

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

Private Sub Command26_Click()
   Command26.Enabled = False
   Screen.MousePointer = 11
   N1% = FILEOPEN%("C:\FLEXOFT\LPVENCLI.PRN", 1, 128)
   While Not EOF(N1%)
     Line Input #N1%, A$
     NC% = Val(Left$(A$, 8))
     If NC% > 0 Then
       LPRE% = Val(Mid$(A$, 17, 8)): If LPRE% < 1 Then LPRE% = 1
       VENDE% = Val(Mid$(A$, 9, 8))
       REGO& = REGICLI&(NC%)
       If REGO& > 0 Then
         If CVI(Left$(REGLEIDO$("DEUDOR1", REGO&), 2)) <> NC% Then
           AAA = BBB
         End If
         XX$ = REGLEIDO$("DEUDOR2", REGO&)
         Call REPLA(XX$, Chr$(LPRE%) + Chr$(VENDE%), 119, 2)
         Call GRAREG("DEUDOR2", XX$, REGO&)
       End If
     End If
   Wend
   Close #N1%
   Call CIERRARCH("DEUDOR2")
   Call CIERRARCH("*.*")
   On Error Resume Next
   Kill "C:\FLEXOFT\LPVENCLI.PRN"
   On Error GoTo 0
   Screen.MousePointer = 1
   Command26.Enabled = True
End Sub

Private Sub Command27_Click()
    '
    Command27.Enabled = False
    NC% = XVALO(Text1)
    If NC% < 1 Then GoTo 99
    If RASOCLI$(NC%) = "" Then GoTo 99
    '
    LANID$ = TRIM$(HORARCLI$(NC%))
    If LANID$ = "" Then
       Call MENSERR(24, "Imposible Cargar usuarios de WEB Services.\Debe Primero Identificar la LAN del Cliente\Mediante el Numero de Serie del Volumen.")
       GoTo 99
    End If
    '
    Call COPYSTRU("FLEXUSR", "WEBUSER")
    Call FILTERFILE("WEBUSER", "!FLEXUSR", 4, LANID$)
    '
320 XUX% = VENTABU%("FLEXUSR")
    If XUX% < 0 Then GoTo 99
    '
    Call FILTERUPDATE("WEBUSER", "!FLEXUSR", 4, LANID$)
    '
99  Call CIERRARCH("*.*")
    Command27.Enabled = True
    '
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

Private Sub COTEXT_GotFocus(Index As Integer)
   If Index = 1 Then
     If DERACCE%("MODIPIVA", "Modificar Posición IVA Cliente") < 2 Then
       On Error Resume Next
       Text33.SetFocus
       On Error GoTo 0
       Exit Sub
     End If
   End If
End Sub

'
Private Sub Form_Load()
    '
    SSTab1.Tab = 0
    Call CENTRAFORM(Me)
    Show
    Refresh
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If InStr(UCase$(App.Path), "COTELEF") > 0 Then
       Label71 = "LAN - ID:"
       VIFO = Text36.FontSize
       Text36.FontSize = 4
       Text36.Width = 2250
       Text36.FontSize = VIFO
    End If
    '
    '\\\OT-06-0559-RG-19/12/06///
    If InStr(EPAC$, "SABO") > 0 Or InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "FERVI") > 0 Then
    '\\\OT-06-0559-RG-FIN///
      Command22.Visible = True
      DoEvents
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
    If DERACCE%("MODIPIVA/NOMESS", "Modificar Posición IVA Cliente") < 2 Then
      COTEXT(1).Enabled = False
    End If
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
    If EXISTE%("C:\FLEXOFT\LPVENCLI.PRN") > 0 Then
      Command26.Visible = True
    End If
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

Private Sub BOTSEL_Click(Index As Integer)
   If Index = 1 Then
     If DERACCE%("MODIPIVA", "Modificar Posición IVA Cliente") < 2 Then
       On Error Resume Next
       Text33.SetFocus
       On Error GoTo 0
       Exit Sub
     End If
   End If
   '
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
   Call GRADACONTA
   Call GRADAENTREGA
   GRABATODO% = 0
End Sub

Private Sub Command10_Click()
   '
   If DERACCE%("CONTACLI", "Actualizar Contactos de Clientes") < 2 Then
      Exit Sub
   End If
   '
   Call GRADACONTA
   FORCONTAC.Show 1
   YACAR%(3) = 0
   Call CARDATCLIE
End Sub

Private Sub Command11_Click()
   Call GRADACONTA
   Call AGRECONTAC
End Sub

Private Sub Command12_Click()
   Call QUITACONTAC
End Sub

Private Sub Command13_Click()
   AMACLIST.Show 1
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   OPAMACLI.Show 1
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADACONTA
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
      '
      NOTRANS$ = TRIM$(Mid$(OBX$, 3, 32))
      TTXX$ = "¿ Desea Aplicar el Domicilio de '" + NOTRANS$ + "'\como Domicilio de Entrega de TODOS los Clientes\Asociados a Este Transportista ?"
      If COMALTER%(TTXX$ + "\\No, Conservar los Anteriores\Si, Aplicar") = 2 Then
         Screen.MousePointer = 1
         NUTRAI% = XVALO(Label51)
         DOMINUE$ = TRIM$(Label63)
           Text43 = DOMINUE$
         LOCANUE$ = TRIM$(Label64)
           Text44 = LOCANUE$
         TRANNUE$ = TRIM$(Label61)
           Text45 = TRANNUE$
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
         Call COMUNI("Se ha Completado la Conversión")
         '
      End If
      '
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
   Call GRADACONTA
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
     PUNCLI$ = ""
     Call COPYSTRU("DEUDOR1", "SELECLI")
     FIN& = ULTREG&("DEUDOR2")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = UCase$(REGLEIDO$("DEUDOR2", U&))
       If InStr(X$, ARGU$) > 0 Then
         X$ = REGLEIDO$("DEUDOR1", U&)
         PUNCLI$ = PUNCLI$ + Left$(X$, 2)
         JJ& = JJ& + 1
         Call GRAREG("!SELECLI", X$, JJ&)
       End If
     Next U&
     '
     FIN& = ULTREG&("BADACLIE")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = UCase$(REGLEIDO$("BADACLIE", U&))
       If InStr(X$, ARGU$) > 0 Then
         NCX$ = Left$(REGLEIDO$("BADACLIE", U&), 2)
         For KKI% = 1 To Len(PUNCLI$) Step 2
           If Mid$(PUNCLI$, KKI%, 2) = NCX$ Then GoTo 19 ' ya esta
         Next KKI%
         UK& = REGBUS&("DEUDOR1", 1, NCX$)
         If UK& > 0 Then
           X$ = REGLEIDO$("DEUDOR1", UK&)
           JJ& = JJ& + 1
           Call GRAREG("!SELECLI", X$, JJ&)
         End If
       End If
19   Next U&
     '
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
   Command21.Enabled = False
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADACONTA
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
   Call GRADACONTA
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
           Call MENSERR(24, "Atención !!! - Cuenta con Movimiento.")
           GoTo 85
         End If
         '
         RFF$ = RECFILT$("COBRAN", 1, MKI$(NC%))
         If RFF$ <> "" Then
           Call MENSERR(24, "Atención !!! -  Cuenta con Movimiento.")
           GoTo 85
         End If
         
85       TTXX$ = "La Baja de un Cliente\Puede Producir Inconsistencias en la Base de Datos\en Caso de que la Cuenta Registre\Facturación, Cobranza o Pedidos."
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
           Call FRESHECHO("DEUDOR1")
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
   Call GRADACONTA
   Call GRADAENTREGA
   Call CARDATCLIE
   '
   Call CIERRARCH("*.*")
   If AGRECLI% > 0 Then
      Call SORTCLI(1)
      AGRECLI% = 0
   End If
   Call FINAL("")
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
   Call GRADACONTA
   Call GRADAENTREGA
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
   If Index = 1 Then
     If DERACCE%("MODIPIVA", "Modificar Posición IVA Cliente") < 2 Then
       On Error Resume Next
       Text33.SetFocus
       On Error GoTo 0
       Exit Sub
     End If
   End If
   '
   Call SELECHO(ARCHEX$(Index))
   VDEVU$ = VALACT1$(ARCHEX$(Index))
   If TRIM$(VDEVU$) <> "" Then
      COTEXT(Index).TEXT = VDEVU$
   End If
End Sub

Sub CARDATCLIE()
    '
    NCLIE% = Val(TRIM$(Text1.TEXT))
    If NCLIE% < 1 Or ECOARCH$("DEUDOR1", MKI$(NCLIE%)) = "" Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text3.TEXT = RASOCLI$(NCLIE%)
    Text24.TEXT = TEL2CLI$(NCLIE%)
    Text23.TEXT = NPROCLI$(NCLIE%)
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
        Call CARLICONTAC
      Case 4
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
      On Error Resume Next
      Command9.SetFocus
      On Error GoTo 0
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
    If XVALO(COTEXT(1)) = 5 Then GoTo 15      ' NO VALIDA COM.EXTERIOR
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
15  If NCLIE% = 0 Then
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
       Call REPLA(XX$, Text37.TEXT, 481, 12)
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
Sub CARLICONTAC()
   '
   If YACAR%(3) > 0 Then Exit Sub
   NCLIE% = Val(Text1.TEXT)
   Call BLACONTAC
   '
   If NCLIE% > 0 Then
      If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         Screen.MousePointer = 11
         RFF$ = RECFILT$("CONTACT2", 3, MKI$(NCLIE%))
         List1.Clear
         For KL& = 1 To Len(RFF$) Step 4
           REG& = CVS(Mid$(RFF$, KL&, 4))
           XX$ = REGLEIDO$("CONTACT2", REG&)
           NCONI% = CVI(Left$(XX$, 2))
           If NCONI% > 0 Then
             NOCONTA$ = ECOARCH$("CONTACT1", MKI$(NCONI%))
             RECONTA& = REGNUM&("CONTACT1")
             TELIS$ = Space$(64)
             Call REPLA(TELIS$, NOCONTA$, 1, 56)
             Call REPLA(TELIS$, Str$(RECONTA&), 57, 8)
             List1.AddItem TELIS$
           End If
         Next KL&
      End If
   End If
   List1.Refresh
   '
   If List1.ListCount > 0 Then
     List1.TopIndex = 0
     List1.ListIndex = 0
     Call CARCONTAC
   End If
   '
   YACAR%(3) = 1: MODIFIC%(3) = 0
   Screen.MousePointer = 1
   '
End Sub

Sub CARCONTAC()
   '
   TELIS$ = List1.TEXT
   REG& = Val(Mid$(TELIS$, 57))
   Screen.MousePointer = 11
   '
   RECON$ = REGLEIDO$("CONTACT1", REG&)
   RECONTANT$ = RECON$
   REGCONTA.Caption = TRIM$(Str$(REG&))
   '
   Label7.Caption = TRIM$(Str$(CVI(Left$(RECON$, 2))))
   Text18.TEXT = TRIM$(Mid$(RECON$, 3, 48))
   Text2.TEXT = TRIM$(Mid$(RECON$, 51, 48))
   Text4.TEXT = TRIM$(Mid$(RECON$, 99, 48))
   Text5.TEXT = TRIM$(Mid$(RECON$, 147, 8))
   Text6.TEXT = TRIM$(Mid$(RECON$, 155, 48))
   Text7.TEXT = TRIM$(Mid$(RECON$, 203, 32))
   Text8.TEXT = TRIM$(Mid$(RECON$, 235, 16))
   Text9.TEXT = TRIM$(Mid$(RECON$, 251, 4))
   Text11.TEXT = TRIM$(Mid$(RECON$, 255, 16))
   Text12.TEXT = TRIM$(Mid$(RECON$, 271, 4))
   Text10.TEXT = TRIM$(Mid$(RECON$, 275, 16))
   Text13.TEXT = TRIM$(Mid$(RECON$, 291, 16))
   Text14.TEXT = TRIM$(Mid$(RECON$, 307, 16))
   Text16.TEXT = TRIM$(Mid$(RECON$, 323, 16))
   Text15.TEXT = TRIM$(Mid$(RECON$, 339, 48))
   Text17.TEXT = TRIM$(Mid$(RECON$, 387, 48))
   '
   NCONI% = Val(Label7.Caption)
   If NCONI% > 0 Then
      RNC$ = RECFILT$("CONTACT3", 1, MKI$(NCONI%))
      URN& = ULTREG&("CONTACT3")
      '
      TTYY$ = ""
      For U& = 1 To Len(RNC$) Step 4
        RENOTA& = CVS(Mid$(RNC$, U&, 4))
        If RENOTA& > 0 Then
          If RENOTA& <= URN& Then
            TTYY$ = TTYY$ + Mid$(REGLEIDO$("CONTACT3", RENOTA&), 3)
          End If
        End If
      Next U&
      Text22.TEXT = RTrim$(TTYY$)
   End If
   '
   MODIFIC%(3) = 0
   Screen.MousePointer = 1
   '
End Sub
'
Sub AGRECONTAC()
   '
   If DERACCE%("CONTACLI", "Actualizar Contactos de Clientes") < 2 Then
      Exit Sub
   End If
   '
   YACAR%(3) = 0
   NCLIE% = Val(Text1.TEXT)
   If NCLIE% > 0 Then
     If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
       Call SELECHO("CONTACT1")
       NCONI% = Val(VALACT1$("CONTACT1"))
       If NCONI% > 0 Then
         NOCONTA$ = AJUSTI$(ECOARCH$("CONTACT1", MKI$(NCONI%)), 56)
         For KKI% = 1 To List1.ListCount
           If Left$(List1.List(KKI% - 1), 56) = NOCONTA$ Then
             Call MENSERR(24, "Contacto Existente")
             Exit Sub
           End If
         Next KKI%
         RECONTAC$ = MKI$(NCONI%) + MKI$(NCLIE%)
         Call REGAPP("CONTACT2", RECONTAC$)
         Call CIERRARCH("CONTACT2")
         Call CARLICONTAC
       End If
     End If
   End If
   '
End Sub
'
Sub QUITACONTAC()
   '
   If DERACCE%("CONTACLI", "Actualizar Contactos de Clientes") < 2 Then
      Exit Sub
   End If
   '
   YACAR%(3) = 0
   NCLIE% = Val(Text1.TEXT)
   If NCLIE% > 0 Then
     If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
       If List1.ListIndex >= 0 Then
         If List1.ListIndex < List1.ListCount Then
           Screen.MousePointer = 11
           NOCONTA$ = TRIM$(Left$(List1.TEXT, 56))
           RFF$ = RECFILT$("CONTACT2", 3, MKI$(NCLIE%))
           For KL& = 1 To Len(RFF$) Step 4
             REG& = CVS(Mid$(RFF$, KL&, 4))
             XX$ = REGLEIDO$("CONTACT2", REG&)
             NCONI% = CVI(Left$(XX$, 2))
             If TRIM$(ECOARCH$("CONTACT1", MKI$(NCONI%))) = NOCONTA$ Then
               Call GRAREG("CONTACT2", String$(4, 0), REG&)
             End If
           Next KL&
           Call CIERRARCH("CONTACT2")
           Call CARLICONTAC
           Screen.MousePointer = 1
         End If
       End If
     End If
   End If
   '
End Sub
'
Sub BLACONTAC()
   '
   Label7.Caption = ""
   Text18.TEXT = ""
   Text2.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   Text6.TEXT = ""
   Text7.TEXT = ""
   Text8.TEXT = ""
   Text9.TEXT = ""
   Text11.TEXT = ""
   Text12.TEXT = ""
   Text10.TEXT = ""
   Text13.TEXT = ""
   Text14.TEXT = ""
   Text16.TEXT = ""
   Text15.TEXT = ""
   Text17.TEXT = ""
   '
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
   Call BLACONTAC
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADACONTA
   Call GRADAENTREGA
   Call CARDATCLIE
   '
   Call CIERRARCH("*.*")
   If AGRECLI% > 0 Then
      Call SORTCLI(1)
      AGRECLI% = 0
   End If
   Call FINAL("")
   '
End Sub

'
Private Sub List1_Click()
   Call GRADACONTA
   Call CARCONTAC
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
   '16/03/07 (OT-07-0113)
5  Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
   '16/03/07 (FIN)
   '
'5  For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Cls
'     EDITFORM.Picture1(KKI%).DrawWidth = 3
'     EDITFORM.Picture1(KKI%).DrawStyle = 0
'     EDITFORM.Picture1(KKI%).FillStyle = 1
'   Next KKI%
'   '
'   REDOCU$ = REGACT$("PDOCLST")
'   URE& = ULTREG&("PDOCSPL")
'   LI& = 0: LS& = URE&
'6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'   GoTo 8
'   '
'7  Screen.MousePointer = 1
'   Call MENSERR(24, "Documento no Registrado")
'   Exit Sub
'   '
'8  TPSP$ = ""
'   Screen.MousePointer = 11
'   While L& > 1
'     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'     If DCC& < DOCUNU& Then GoTo 9
'     L& = L& - 1
'   Wend
'   '
'9  CAPAGINAS% = 0
'   For KKL& = L& To URE&
'     TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'     End If
'   Next KKL&
'   '
'95 SPOOLSTRING$ = TPSP$
'   For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Cls
'     EDITFORM.Picture1(KKI%).Top = 0
'     EDITFORM.Picture1(KKI%).Refresh
'     EDITFORM.Picture1(KKI%).Height = 7170
'     EDITFORM.Picture1(KKI%).Width = 11700
'     EDITFORM.Picture1(KKI%).DrawWidth = 3
'     EDITFORM.Picture1(KKI%).DrawStyle = 0
'     EDITFORM.Picture1(KKI%).FillStyle = 1
'   Next KKI%
'   ALTUPAGINA = 7170
'   TOPEPAGINA = 0
'   PAGINACTIVA% = 0
'   CAPAGINAS% = 0
'   '
'   FIN& = Len(TPSP$)
'10 PPXX% = InStr(TPSP$, Chr$(255))
'   If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'   End If
'   '
'   Screen.MousePointer = 1
'   PAGINACTIVA% = 0
'   For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Visible = False
'     If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
'   Next KKI%
'   '
'   EDITFORM.Picture1(PAGINACTIVA%).Refresh
'   EDITFORM.VScroll1.Min = 0
'   EDITFORM.VScroll1.Value = 0
'   EDITFORM.VScroll1.Max = 1
'   If ALTUPAGINA > EDITFORM.Frame1.Height Then
'      EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'   End If
'   EDITFORM.Command3.Enabled = False
'   EDITFORM.Command5.Enabled = False
'   EDITFORM.Show
   '
End Sub

Private Sub LTEXT36_Change()
   DISPONI = Val(LTEXT36.Caption) - Val(LTEXT37.Caption) - Val(LTEXT35.Caption)
   LTEXT23.ForeColor = &H80000008
   If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
   LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
   LTEXT23.Refresh
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab <> PreviousTab Then
      Call GRAGENCLI
      Call GRANOTACLI
      Call GRADACONTA
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
Sub GRADACONTA()
   '
   If MODIFIC%(3) < 1 Then
      Exit Sub
   End If
   '
   If TRIM$(Text18.TEXT) = "" Then GoTo 99
   If TRIM$(Label7.Caption) = "" Then GoTo 99
   '
   If DERACCE%("CONTACLI", "Actualizar Contactos de Clientes") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     Command9.SetFocus
     If COMALTER%("Se han Modificado Datos del Contacto\'" + TRIM$(Text18.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(3) = 0
       Call CARCONTAC
       Exit Sub
     End If
   End If
   '
   NCON& = Val(Label7.Caption)
   If NCON& < 1 Or NCON& > 32767 Then
      Call MENSERR(24, "Imposible Registrar.\Número de Contacto Fuera de Rango.")
      Exit Sub
   End If
   NCONI% = NCON&
   '
   DECON$ = TRIM$(Text18.TEXT)
   If DECON$ = "" Then
      Call MENSERR(24, "Imposible Registrar.\Falta Nombre del Contacto.")
      Exit Sub
   End If
   '
   Call BLOQARCH("CONTACT1")
   Call BLOQARCH("CONTACT3")
   Screen.MousePointer = 11
   '
   RECON$ = REGBLAN$("CONTACT1")
   Call REPLA(RECON$, MKI$(NCONI%), 1, 2)
   Call REPLA(RECON$, DECON$, 3, 48)
   Call REPLA(RECON$, Text2.TEXT, 51, 48)
   Call REPLA(RECON$, Text4.TEXT, 99, 48)
   Call REPLA(RECON$, Text5.TEXT, 147, 8)
   Call REPLA(RECON$, Text6.TEXT, 155, 48)
   Call REPLA(RECON$, Text7.TEXT, 203, 32)
   Call REPLA(RECON$, Text8.TEXT, 235, 16)
   Call REPLA(RECON$, Text9.TEXT, 251, 4)
   Call REPLA(RECON$, Text11.TEXT, 255, 16)
   Call REPLA(RECON$, Text12.TEXT, 271, 4)
   Call REPLA(RECON$, Text10.TEXT, 275, 16)
   Call REPLA(RECON$, Text13.TEXT, 291, 16)
   Call REPLA(RECON$, Text14.TEXT, 307, 16)
   Call REPLA(RECON$, Text16.TEXT, 323, 16)
   Call REPLA(RECON$, Text15.TEXT, 339, 48)
   Call REPLA(RECON$, Text17.TEXT, 387, 48)
   '
   RFF$ = RECFILT$("CONTACT1", 1, MKI$(NCONI%))
   REG& = 0: If Len(RFF$) >= 4 Then REG& = CVS(Left$(RFF$, 4))
   If REG& < 1 Or REG& > ULTREG&("CONTACT1") Or REG& <> Val(REGCONTA.Caption) Then
      Call MENSERR(24, "Imposible Grabar Modificaciones - Vuelva a Intentar")
      GoTo 90
   End If
   '
   RCAA$ = REGLEIDO$("CONTACT1", REG&)
   If CVI(Left$(RCAA$, 2)) <> NCONI% Or RCAA$ <> RECONTANT$ Then
      Call MENSERR(24, "Imposible Grabar Modificaciones - Vuelva a Intentar")
      GoTo 90
   End If
   '
   Call GRAREG("CONTACT1", RECON$, REG&)
   '
   URN& = ULTREG&("CONTACT3")
   RNC$ = RECFILT$("CONTACT3", 1, MKI$(NCONI%))
   KK& = (-3)
   '
   TEXTO$ = RTrim$(Text22.TEXT)
   While Len(TEXTO$) > 0
      TTYY$ = MKI$(NCONI%) + Left$(TEXTO$, 254)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("CONTACT3", TTYY$, RENOTA&)
      TEXTO$ = Mid$(TEXTO$, 255)
   Wend
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("CONTACT3", MKI$(NCONI%) + String$(254, 0), RENOTA&)
   Wend
   '
90 Call CIERRARCH("CONTACT1")
   Call CIERRARCH("CONTACT3")
   Call BAJALDISCO
   '
99 MODIFIC%(3) = 0
   YACAR%(3) = 0
   Call CARLICONTAC
   Screen.MousePointer = 1

End Sub

Private Sub Text1_DblClick()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADACONTA
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
   NCLIE% = Val(Text1)
   If NCLIE% > 0 Then
     If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
       Call GRAGENCLI
       Call GRANOTACLI
       Call GRADACONTA
       Call GRADAENTREGA
     End If
   End If
   '
End Sub

Private Sub Text10_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text11_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text12_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text13_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text14_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text15_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text16_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text17_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text18_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text19_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text2_Change()
   MODIFIC%(3) = 1
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

Private Sub Text22_Change()
   MODIFIC%(3) = 1
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

Private Sub Text35_Change()
   MODIFIC%(4) = 1
End Sub

Private Sub Text37_Change()
   MODIFIC%(4) = 1
End Sub

Private Sub Text4_Change()
   MODIFIC%(3) = 1
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
   Text43.SetFocus
End Sub

Private Sub Text5_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text6_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text7_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text8_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text9_Change()
   MODIFIC%(3) = 1
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
  If MONEMX% <= 1 Then
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If CVS(Mid$(XXC$, 39, 4)) = NROL& Then
        If Asc(Mid$(XXC$, 37, 1)) = TI% Then
          MONEMX% = Asc(Mid$(XXC$, 31, 1))
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
          SALDOPES# = CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8))
          If Abs(SALDOPES# - SALDO#) <= 0.005 Then
            SALTEX$ = "u$s" + FORMATNUM$(CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8)), "F11.2")
            GoTo 20
          End If
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
               Text43 = TRIM$(Mid$(XX$, 3, 32))
               Text44 = TRIM$(Mid$(XX$, 35, 32))
               Text45 = TRIM$(Mid$(XX$, 67, 32))
               Text35 = TRIM$(Mid$(XX$, 101, 24))
               '
               Text37 = ZONACLI$(NCLIE%)
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
         ABX$ = FILTERGETRECORD$("BADACLIE", 1, MKI$(NCLIE%))
         Call REPLA(ABX$, Text37, 481, 12)
         Call FILTERPUTRECORD("BADACLIE", 1, MKI$(NCLIE%), ABX$)
      End If
   End If
   '
   Call CIERRARCH("LENTRECL")
   MODIFIC%(4) = 0
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
    If Left$(App.EXEName, 6) = "AMACLI" Then
         If AMACLIEN.List3.ListCount > 0 Then
           SUMARRAS# = Val(Mid$(AMACLIEN.List3.List(AMACLIEN.List3.ListCount - 1), 63, 14))
         End If
         SUMADEUP# = Val(AMACLIEN.Label28.Caption)
         CADOCPEN% = AMACLIEN.List2.ListCount
      Else
         If RESCUDEU.List1.ListCount > 0 Then
           SUMARRAS# = Val(Mid$(RESCUDEU.List1.List(RESCUDEU.List1.ListCount - 1), 63, 14))
         End If
         SUMADEUP# = Val(RESCUDEU.Label5.Caption)
         CADOCPEN% = RESCUDEU.List2.ListCount
    End If
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
