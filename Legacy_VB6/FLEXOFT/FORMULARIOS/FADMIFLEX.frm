VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form FADMIFLEX 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT Soluciones - Administracion de Actividad"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "FADMIFLEX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Interval        =   60000
      Left            =   11160
      Top             =   5280
   End
   Begin VB.CommandButton Command35 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   11025
      TabIndex        =   159
      ToolTipText     =   "Revisa Registraciones"
      Top             =   4305
      Width           =   750
   End
   Begin VB.CommandButton Command34 
      Caption         =   "R.Elec."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   11025
      Picture         =   "FADMIFLEX.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   7665
      Width           =   750
   End
   Begin VB.CommandButton Command26 
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
      Height          =   750
      Left            =   11025
      Picture         =   "FADMIFLEX.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   2625
      Width           =   750
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Config"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   11025
      Picture         =   "FADMIFLEX.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   1785
      Width           =   750
   End
   Begin VB.CommandButton Command11 
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
      Height          =   750
      Left            =   11025
      Picture         =   "FADMIFLEX.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
      Top             =   945
      Width           =   750
   End
   Begin VB.CommandButton Command6 
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
      Height          =   750
      Left            =   11025
      Picture         =   "FADMIFLEX.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   105
      Width           =   750
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   15214
      _Version        =   327681
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   882
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Consultas Telefonicas"
      TabPicture(0)   =   "FADMIFLEX.frx":12EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Timer1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame20"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command22"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame21"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command32"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command27"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame29"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Timer2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Timer3"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Ordenes de Trabajo"
      TabPicture(1)   =   "FADMIFLEX.frx":1308
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label24"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame9"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame12"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame14"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame15"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame16"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame18"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame19"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Command18"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Command19"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame13"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Frame11"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Command21"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Tareas Internas"
      TabPicture(2)   =   "FADMIFLEX.frx":1324
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command31"
      Tab(2).Control(1)=   "Frame28"
      Tab(2).Control(2)=   "Frame27"
      Tab(2).Control(3)=   "Command30"
      Tab(2).Control(4)=   "Command29"
      Tab(2).Control(5)=   "Frame24"
      Tab(2).Control(6)=   "Frame23"
      Tab(2).Control(7)=   "Frame22"
      Tab(2).Control(8)=   "Command23"
      Tab(2).Control(9)=   "Label34"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Clientes"
      TabPicture(3)   =   "FADMIFLEX.frx":1340
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.Timer Timer3 
         Interval        =   60000
         Left            =   1890
         Top             =   315
      End
      Begin VB.Timer Timer2 
         Interval        =   60000
         Left            =   1155
         Top             =   315
      End
      Begin VB.Frame Frame29 
         BorderStyle     =   0  'None
         Caption         =   "Frame29"
         Height          =   525
         Left            =   5670
         TabIndex        =   146
         Top             =   3120
         Width           =   3900
         Begin VB.OptionButton Option5 
            Caption         =   "Otros"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   3000
            TabIndex        =   150
            Top             =   220
            Width           =   1065
         End
         Begin VB.OptionButton Option5 
            Caption         =   "New Apli"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   2040
            TabIndex        =   149
            Top             =   0
            Width           =   1065
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Reclamo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1050
            TabIndex        =   148
            Top             =   220
            Width           =   1065
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Consulta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   147
            Top             =   0
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   9765
         Picture         =   "FADMIFLEX.frx":135C
         Style           =   1  'Graphical
         TabIndex        =   145
         ToolTipText     =   "Cierre general de CTel's"
         Top             =   4440
         Width           =   855
      End
      Begin VB.CommandButton Command32 
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
         Height          =   720
         Left            =   9765
         Picture         =   "FADMIFLEX.frx":1666
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Listado por Cliente y Rango de Fechas"
         Top             =   2960
         Width           =   855
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Search"
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
         Left            =   -65025
         Picture         =   "FADMIFLEX.frx":1970
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Anulación (Borrado) de Tareas Programadas"
         Top             =   3110
         Width           =   740
      End
      Begin VB.Frame Frame28 
         Caption         =   "Detalle de Pendientes por Responsable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4275
         Left            =   -74790
         TabIndex        =   130
         Top             =   4200
         Width           =   10515
         Begin VB.ListBox List3 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3480
            Left            =   120
            TabIndex        =   131
            Top             =   630
            Width           =   10260
         End
         Begin VB.Line Line5 
            X1              =   10080
            X2              =   10080
            Y1              =   315
            Y2              =   630
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            X1              =   130
            X2              =   130
            Y1              =   350
            Y2              =   630
         End
         Begin VB.Label Label26 
            BackColor       =   &H80000016&
            Caption         =   "  Nro.  Fecha    Detalle / Descripcion                                     Respons."
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   105
            TabIndex        =   132
            Top             =   360
            Width           =   9990
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   10080
            Y1              =   330
            Y2              =   330
         End
      End
      Begin VB.Frame Frame27 
         Caption         =   "Consulta de Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   -70065
         TabIndex        =   122
         Top             =   3150
         Width           =   4950
         Begin VB.OptionButton Option4 
            Caption         =   "Ezequiel"
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
            Index           =   4
            Left            =   2760
            TabIndex        =   127
            Top             =   630
            Width           =   1170
         End
         Begin VB.CommandButton Command25 
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
            Left            =   4200
            Picture         =   "FADMIFLEX.frx":1C7A
            Style           =   1  'Graphical
            TabIndex        =   129
            ToolTipText     =   "Listado Tareas Pendientes por Responsable"
            Top             =   210
            Width           =   630
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Todos"
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
            Index           =   9
            Left            =   3255
            TabIndex        =   136
            Top             =   315
            Width           =   1065
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Op-8"
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
            Index           =   8
            Left            =   3150
            TabIndex        =   135
            Top             =   0
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Op-7"
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
            Index           =   7
            Left            =   3150
            TabIndex        =   134
            Top             =   0
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Op-6"
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
            Index           =   6
            Left            =   3150
            TabIndex        =   133
            Top             =   0
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Omar"
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
            Index           =   5
            Left            =   1800
            TabIndex        =   128
            Top             =   630
            Width           =   1065
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Natalia"
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
            Index           =   3
            Left            =   2280
            TabIndex        =   126
            Top             =   315
            Width           =   960
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Francisco"
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
            Index           =   2
            Left            =   1050
            TabIndex        =   125
            Top             =   315
            Width           =   1275
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Flavio"
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
            Index           =   1
            Left            =   630
            TabIndex        =   124
            Top             =   630
            Width           =   960
         End
         Begin VB.OptionButton Option4 
            Caption         =   "EPB"
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
            Index           =   0
            Left            =   210
            TabIndex        =   123
            Top             =   315
            Value           =   -1  'True
            Width           =   750
         End
      End
      Begin VB.CommandButton Command30 
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
         Height          =   730
         Left            =   -65025
         Picture         =   "FADMIFLEX.frx":1F84
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Nueva Tarea Programada"
         Top             =   945
         Width           =   740
      End
      Begin VB.CommandButton Command29 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   730
         Left            =   -65025
         Picture         =   "FADMIFLEX.frx":20CE
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Graba las Modificaciones"
         Top             =   1680
         Width           =   740
      End
      Begin VB.Frame Frame24 
         Caption         =   "Referencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2250
         Left            =   -70065
         TabIndex        =   104
         Top             =   840
         Width           =   4950
         Begin VB.Frame Frame25 
            Caption         =   "Cierre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   3570
            TabIndex        =   119
            Top             =   1155
            Width           =   1275
            Begin VB.CommandButton Command33 
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
               Height          =   250
               Left            =   1010
               TabIndex        =   120
               ToolTipText     =   "Fecha Enviada por Mail"
               Top             =   525
               Width           =   160
            End
            Begin VB.TextBox Text31 
               BackColor       =   &H00FFFFFF&
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
               Left            =   105
               MaxLength       =   12
               TabIndex        =   121
               TabStop         =   0   'False
               ToolTipText     =   "Fecha Enviada por Mail"
               Top             =   525
               Width           =   990
            End
            Begin VB.Image Image3 
               Height          =   330
               Left            =   900
               Picture         =   "FADMIFLEX.frx":23D8
               Stretch         =   -1  'True
               Top             =   120
               Width           =   330
            End
         End
         Begin VB.Frame Frame26 
            Caption         =   "Responsable"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   130
            TabIndex        =   113
            Top             =   1155
            Width           =   3375
            Begin VB.OptionButton Option3 
               Caption         =   "Omar"
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
               Index           =   5
               Left            =   1440
               TabIndex        =   144
               Top             =   630
               Width           =   960
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Ezequiel"
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
               Index           =   4
               Left            =   2400
               TabIndex        =   118
               Top             =   630
               Width           =   900
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Natalia"
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
               Index           =   3
               Left            =   2280
               TabIndex        =   117
               Top             =   315
               Width           =   900
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Francisco"
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
               Index           =   2
               Left            =   1080
               TabIndex        =   116
               Top             =   315
               Width           =   1170
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Flavio"
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
               Index           =   1
               Left            =   525
               TabIndex        =   115
               Top             =   630
               Width           =   1275
            End
            Begin VB.OptionButton Option3 
               Caption         =   "EPB"
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
               Index           =   0
               Left            =   210
               TabIndex        =   114
               Top             =   315
               Value           =   -1  'True
               Width           =   750
            End
         End
         Begin VB.CommandButton Command28 
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
            Left            =   3045
            TabIndex        =   106
            Top             =   315
            Width           =   180
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
            Height          =   285
            Left            =   130
            TabIndex        =   105
            Top             =   735
            Width           =   4710
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "P-30"
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
            Left            =   4380
            TabIndex        =   142
            ToolTipText     =   "Posterga 1 mes"
            Top             =   315
            Width           =   435
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "P-7"
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
            Left            =   3885
            TabIndex        =   141
            ToolTipText     =   "Posterga 1 semana"
            Top             =   315
            Width           =   330
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "P-1"
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
            TabIndex        =   140
            ToolTipText     =   "Posterga 1 día"
            Top             =   315
            Width           =   330
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
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
            Left            =   420
            TabIndex        =   108
            Top             =   315
            Width           =   1035
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "F:"
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
            Left            =   -315
            TabIndex        =   110
            Top             =   375
            Width           =   645
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Nro:"
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
            Left            =   1365
            TabIndex        =   109
            Top             =   375
            Width           =   750
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Left            =   2205
            TabIndex        =   107
            Top             =   315
            Width           =   855
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1890
         Left            =   -74790
         TabIndex        =   102
         Top             =   840
         Width           =   4635
         Begin VB.TextBox Text28 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1380
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   103
            Top             =   315
            Width           =   4400
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   -74790
         TabIndex        =   100
         Top             =   2730
         Width           =   4635
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   101
            Top             =   210
            Width           =   4395
         End
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Anul."
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
         Left            =   -65025
         Picture         =   "FADMIFLEX.frx":26E2
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Anulación (Borrado) de Tareas Programadas"
         Top             =   2415
         Width           =   740
      End
      Begin VB.Frame Frame21 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1220
         Left            =   9765
         TabIndex        =   96
         Top             =   5250
         Width           =   855
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   60
            TabIndex        =   97
            Top             =   840
            Width           =   750
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   180
            Picture         =   "FADMIFLEX.frx":29EC
            Top             =   210
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   180
            Picture         =   "FADMIFLEX.frx":2CF6
            Top             =   210
            Width           =   480
         End
      End
      Begin VB.CommandButton Command22 
         Caption         =   "O.Trab."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   9765
         Picture         =   "FADMIFLEX.frx":3000
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Genera Orden de Trabajo"
         Top             =   3700
         Width           =   855
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Anul."
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
         Left            =   -65025
         Picture         =   "FADMIFLEX.frx":330A
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Informe Tareas Por Fecha"
         Top             =   4050
         Width           =   740
      End
      Begin VB.Frame Frame11 
         Caption         =   "Módulos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2520
         Left            =   -66810
         TabIndex        =   92
         Top             =   3045
         Width           =   1590
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2050
            IntegralHeight  =   0   'False
            Left            =   105
            TabIndex        =   93
            Top             =   315
            Width           =   1380
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Instalación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   -74895
         TabIndex        =   52
         Top             =   7770
         Width           =   4530
         Begin VB.CommandButton Command16 
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
            Height          =   250
            Left            =   4245
            TabIndex        =   84
            ToolTipText     =   "Fecha de Cierre"
            Top             =   315
            Width           =   160
         End
         Begin VB.CommandButton Command15 
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
            Height          =   250
            Left            =   2775
            TabIndex        =   83
            ToolTipText     =   "Fecha Verificada Instalación"
            Top             =   315
            Width           =   160
         End
         Begin VB.CommandButton Command13 
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
            Height          =   250
            Left            =   1305
            TabIndex        =   82
            ToolTipText     =   "Fecha Enviada por Mail"
            Top             =   315
            Width           =   160
         End
         Begin VB.TextBox Text27 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1845
            MaxLength       =   12
            TabIndex        =   73
            TabStop         =   0   'False
            Text            =   " "
            ToolTipText     =   "Fecha Verificada Instalación"
            Top             =   315
            Width           =   1100
         End
         Begin VB.TextBox Text26 
            BackColor       =   &H00FFFFFF&
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
            Left            =   375
            MaxLength       =   12
            TabIndex        =   72
            TabStop         =   0   'False
            ToolTipText     =   "Fecha Enviada por Mail"
            Top             =   315
            Width           =   1100
         End
         Begin VB.TextBox Text20 
            BackColor       =   &H00FFFFFF&
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
            Height          =   250
            Left            =   3315
            MaxLength       =   12
            TabIndex        =   71
            TabStop         =   0   'False
            Text            =   " "
            ToolTipText     =   "Doble Click para Cerrar"
            Top             =   315
            Width           =   1100
         End
         Begin VB.Label Label23 
            Caption         =   "I:"
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
            Left            =   1575
            TabIndex        =   76
            ToolTipText     =   "Fecha Verificada Instalación"
            Top             =   345
            Width           =   225
         End
         Begin VB.Label Label22 
            Caption         =   "E:"
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
            TabIndex        =   75
            ToolTipText     =   "Fecha Enviada por Mail"
            Top             =   345
            Width           =   225
         End
         Begin VB.Label Label19 
            Caption         =   "C:"
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
            Left            =   3045
            TabIndex        =   74
            ToolTipText     =   "Fecha de Cierre"
            Top             =   345
            Width           =   225
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   9765
         TabIndex        =   88
         Top             =   6490
         Width           =   855
         Begin VB.CommandButton Command20 
            Caption         =   "Attch"
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
            Picture         =   "FADMIFLEX.frx":3614
            Style           =   1  'Graphical
            TabIndex        =   90
            ToolTipText     =   "Adjuntar Archivos"
            Top             =   315
            Width           =   645
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Send"
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
            Picture         =   "FADMIFLEX.frx":375E
            Style           =   1  'Graphical
            TabIndex        =   89
            ToolTipText     =   "Envia por E-Mail"
            Top             =   1050
            Width           =   645
         End
      End
      Begin VB.CommandButton Command19 
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
         Left            =   -65025
         Picture         =   "FADMIFLEX.frx":3A68
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Informe Tareas Por Fecha"
         Top             =   3360
         Width           =   740
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65025
         Picture         =   "FADMIFLEX.frx":3D72
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Informe Tareas Por Fecha"
         Top             =   4725
         Width           =   740
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   105
         Top             =   315
      End
      Begin VB.Frame Frame19 
         Caption         =   "Liquidación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   -66915
         TabIndex        =   66
         Top             =   7770
         Width           =   2640
         Begin VB.TextBox Text25 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   1470
            TabIndex        =   68
            TabStop         =   0   'False
            Text            =   " "
            Top             =   315
            Width           =   1035
         End
         Begin VB.TextBox Text24 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   315
            TabIndex        =   67
            TabStop         =   0   'False
            Text            =   " "
            Top             =   315
            Width           =   720
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "$:"
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
            Left            =   1050
            TabIndex        =   70
            Top             =   350
            Width           =   330
         End
         Begin VB.Label Label20 
            Caption         =   "F:"
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
            TabIndex        =   69
            Top             =   350
            Width           =   750
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Horas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Left            =   -65130
         TabIndex        =   59
         Top             =   5670
         Width           =   855
         Begin VB.TextBox Text19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   105
            TabIndex        =   64
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1680
            Width           =   660
         End
         Begin VB.TextBox Text18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   105
            TabIndex        =   61
            TabStop         =   0   'False
            Text            =   " "
            Top             =   630
            Width           =   660
         End
         Begin VB.TextBox Text17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   105
            TabIndex        =   60
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1155
            Width           =   660
         End
         Begin VB.Label Label18 
            Caption         =   "Total:"
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
            TabIndex        =   65
            Top             =   1470
            Width           =   645
         End
         Begin VB.Label Label17 
            Caption         =   "EPB:"
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
            TabIndex        =   63
            Top             =   420
            Width           =   435
         End
         Begin VB.Label Label12 
            Caption         =   "Ots:"
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
            TabIndex        =   62
            Top             =   945
            Width           =   435
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Descripción / Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2520
         Left            =   -74895
         TabIndex        =   56
         Top             =   3045
         Width           =   7995
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2050
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   57
            Top             =   315
            Width           =   7755
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Descripcion / Síntoma / Fallo Reportado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -74895
         TabIndex        =   54
         Top             =   630
         Width           =   4635
         Begin VB.OptionButton Option1 
            Caption         =   "Varios"
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
            Index           =   9
            Left            =   3570
            TabIndex        =   156
            Top             =   630
            Width           =   1005
         End
         Begin VB.OptionButton Option1 
            Caption         =   "N/Desarr."
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
            Left            =   2205
            TabIndex        =   155
            Top             =   630
            Width           =   1245
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Modif."
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
            Index           =   2
            Left            =   1155
            TabIndex        =   154
            Top             =   630
            Width           =   930
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Repar."
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
            Left            =   105
            TabIndex        =   153
            Top             =   630
            Value           =   -1  'True
            Width           =   930
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
            Height          =   285
            Left            =   105
            TabIndex        =   152
            Top             =   315
            Width           =   4395
         End
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1380
            Left            =   105
            MaxLength       =   256
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   55
            Top             =   900
            Width           =   4400
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Desarrollo del Trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Left            =   -74895
         TabIndex        =   53
         Top             =   5670
         Width           =   9675
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
            Left            =   120
            TabIndex        =   81
            Top             =   630
            Width           =   9420
         End
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   130
            X2              =   130
            Y1              =   350
            Y2              =   630
         End
         Begin VB.Line Line2 
            X1              =   9240
            X2              =   9240
            Y1              =   315
            Y2              =   630
         End
         Begin VB.Line Line1 
            X1              =   115
            X2              =   9240
            Y1              =   325
            Y2              =   325
         End
         Begin VB.Label Label11 
            BackColor       =   &H80000016&
            Caption         =   " Fecha    Detalle / Descripcion                         Com   Fin   Hs. Rsp"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   105
            TabIndex        =   58
            Top             =   360
            Width           =   9150
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Imputación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   -70275
         TabIndex        =   48
         Top             =   7770
         Width           =   3270
         Begin VB.OptionButton Option2 
            Caption         =   "Vrs"
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
            Index           =   9
            Left            =   2625
            TabIndex        =   80
            Top             =   315
            Width           =   600
         End
         Begin VB.OptionButton Option2 
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
            Height          =   225
            Index           =   3
            Left            =   1785
            TabIndex        =   51
            Top             =   315
            Width           =   1100
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Gtía"
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
            Index           =   2
            Left            =   945
            TabIndex        =   50
            Top             =   315
            Width           =   1100
         End
         Begin VB.OptionButton Option2 
            Caption         =   "A-M"
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
            Left            =   105
            TabIndex        =   49
            Top             =   315
            Value           =   -1  'True
            Width           =   1100
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Referencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -70170
         TabIndex        =   38
         Top             =   630
         Width           =   4950
         Begin VB.Frame Frame17 
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1155
            Left            =   105
            TabIndex        =   157
            Top             =   1155
            Width           =   4740
            Begin VB.TextBox Text16 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   105
               MaxLength       =   126
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   158
               Top             =   210
               Width           =   4500
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Caption         =   "Motivo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   210
            TabIndex        =   91
            Top             =   1720
            Width           =   4635
         End
         Begin VB.CommandButton Command17 
            Caption         =   "P"
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
            Left            =   4515
            TabIndex        =   85
            ToolTipText     =   "Selección Ordenes de Trabajo Pendientes (No Cerradas)"
            Top             =   315
            Width           =   285
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
            Left            =   4200
            TabIndex        =   43
            Top             =   315
            Width           =   180
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H00FFFFFF&
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
            Left            =   2625
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   " "
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Usuario:"
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
            Left            =   2310
            TabIndex        =   47
            Top             =   620
            Width           =   750
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFC0&
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
            Left            =   210
            TabIndex        =   46
            Top             =   840
            Width           =   2280
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
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
            Left            =   2730
            TabIndex        =   45
            Top             =   315
            Width           =   1485
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
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
            Left            =   945
            TabIndex        =   44
            Top             =   315
            Width           =   1035
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Cliente:"
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
            TabIndex        =   42
            Top             =   620
            Width           =   645
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Nro:"
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
            Left            =   1890
            TabIndex        =   40
            Top             =   380
            Width           =   750
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Left            =   210
            TabIndex        =   39
            Top             =   380
            Width           =   645
         End
      End
      Begin VB.CommandButton Command10 
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
         Height          =   730
         Left            =   -65025
         Picture         =   "FADMIFLEX.frx":53A4
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Imprime Formulario de Consulta"
         Top             =   1470
         Width           =   740
      End
      Begin VB.CommandButton Command9 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   730
         Left            =   -65025
         Picture         =   "FADMIFLEX.frx":5A0E
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Graba las Modificaciones"
         Top             =   735
         Width           =   740
      End
      Begin VB.CommandButton Command8 
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
         Height          =   730
         Left            =   -65025
         Picture         =   "FADMIFLEX.frx":5D18
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Nueva Consulta"
         Top             =   2205
         Width           =   740
      End
      Begin VB.CommandButton Command7 
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
         Height          =   720
         Left            =   9765
         Picture         =   "FADMIFLEX.frx":5E62
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Nueva Consulta"
         Top             =   740
         Width           =   855
      End
      Begin VB.Frame Frame8 
         Caption         =   "Doc.Anexos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   7560
         TabIndex        =   30
         Top             =   3675
         Width           =   2010
         Begin VB.TextBox Text7 
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
            Left            =   210
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            TabStop         =   0   'False
            Text            =   "FADMIFLEX.frx":5FAC
            Top             =   315
            Width           =   1590
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Número"
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
         Left            =   7875
         TabIndex        =   27
         Top             =   630
         Width           =   1695
         Begin VB.VScrollBar VScroll1 
            Height          =   310
            Left            =   1365
            TabIndex        =   151
            Top             =   315
            Value           =   16384
            Width           =   175
         End
         Begin VB.CommandButton Command24 
            Caption         =   "P"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1260
            TabIndex        =   98
            ToolTipText     =   "Seleccion Consultas Pendientes"
            Top             =   735
            Width           =   285
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Buscar"
            Height          =   270
            Left            =   130
            TabIndex        =   32
            Top             =   735
            Width           =   915
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   105
            TabIndex        =   34
            Top             =   315
            Width           =   1275
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Origen"
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
         Left            =   5670
         TabIndex        =   22
         Top             =   630
         Width           =   2115
         Begin VB.TextBox Text6 
            BackColor       =   &H00FFFFC0&
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
            Left            =   735
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   " "
            Top             =   700
            Width           =   1035
         End
         Begin VB.TextBox Text12 
            BackColor       =   &H00FFFFC0&
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
            Left            =   735
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   " "
            Top             =   315
            Width           =   1035
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Hora:"
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
            Left            =   75
            TabIndex        =   26
            Top             =   795
            Width           =   645
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Left            =   75
            TabIndex        =   24
            Top             =   420
            Width           =   645
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   9765
         Picture         =   "FADMIFLEX.frx":5FAE
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Graba las Modificaciones"
         Top             =   1480
         Width           =   855
      End
      Begin VB.CommandButton Command5 
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
         Height          =   720
         Left            =   9765
         Picture         =   "FADMIFLEX.frx":62B8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprime Formulario de Consulta"
         Top             =   2220
         Width           =   855
      End
      Begin VB.Frame Frame5 
         Caption         =   "Distribuidor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1220
         Left            =   5670
         TabIndex        =   8
         Top             =   1820
         Width           =   3900
         Begin VB.ComboBox Combo1 
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
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   160
            Top             =   315
            Width           =   2955
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
            Height          =   295
            Left            =   735
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   735
            Width           =   2955
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "C/C:"
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
            Left            =   160
            TabIndex        =   11
            Top             =   800
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Para:"
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
            Left            =   160
            TabIndex        =   9
            Top             =   380
            Width           =   540
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "  Referencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   210
         TabIndex        =   6
         Top             =   2310
         Width           =   5370
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
            Height          =   855
            Left            =   210
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   315
            Width           =   4965
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "  Diagnóstico e Instrucciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3225
         Left            =   210
         TabIndex        =   4
         Top             =   5175
         Width           =   9360
         Begin VB.TextBox Text23 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2745
            Left            =   210
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   315
            Width           =   8955
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "  Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   210
         TabIndex        =   2
         Top             =   3675
         Width           =   7260
         Begin VB.TextBox Text22 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   210
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   315
            Width           =   6855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   210
         TabIndex        =   1
         ToolTipText     =   "Doble Click para Enviar Registro de Empresa"
         Top             =   630
         Width           =   5370
         Begin MSMAPI.MAPIMessages MAPIMessages1 
            Left            =   4620
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin MSMAPI.MAPISession MAPISession1 
            Left            =   3780
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
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
            TabIndex        =   18
            Text            =   " "
            Top             =   315
            Width           =   855
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
            TabIndex        =   17
            Top             =   315
            Width           =   180
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFC0&
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
         Begin VB.TextBox Text3 
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   " "
            Top             =   735
            Width           =   4950
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00FFFFC0&
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
            Left            =   560
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1155
            Width           =   2115
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H00FFFFC0&
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
            Left            =   3360
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1155
            Width           =   1800
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
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
            Left            =   2900
            TabIndex        =   29
            Top             =   1220
            Width           =   435
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Tel:"
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
            TabIndex        =   28
            Top             =   1220
            Width           =   435
         End
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -65025
         TabIndex        =   143
         Top             =   3885
         Width           =   750
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   -65025
         TabIndex        =   78
         Top             =   3045
         Visible         =   0   'False
         Width           =   750
      End
   End
End
Attribute VB_Name = "FADMIFLEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NUCETE$
Dim ABUSQUE$
Dim ENCARLISTA%

Private Sub Combo1_Click()
   Text10 = ""
   If Combo1.ListIndex > 0 Then
      Text10 = Combo1.List(0)
   End If
End Sub

Private Sub Command10_Click()
   Call PRIORTRA(TRIM$(Label15.Caption))
End Sub

Private Sub Command13_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = VALACT1$("SELFECHA")
   If VDEVU$ <> "" Then
     Text26 = VDEVU$
   End If
End Sub

Private Sub Command12_Click()
    '
    FORSELEC.ABUSQUE = Label15
    Call SELECHO("CORTRAB1")
    NUCOTE$ = AJUSTI$(VALACT1$("CORTRAB1"), 12)
    If TRIM$(NUCOTE$) <> "" Then
       Label15.Caption = NUCOTE$
    End If
    '
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call CONECRUN("FLSETUP", "Configuración de Funcionamiento")
   Call FINAL("")
   Command14.Enabled = True
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
      NC% = NCXI%
      Text1.TEXT = TRIM$(Str$(NC%))
      Text2.TEXT = RASOCLI$(NC%)
      Text3.TEXT = TRIM$(DOMICLI$(NC%)) + " - (" + TRIM$(CPOSCLI$(NC%)) + ") " + LOCACLI$(NC%)
      Text4.TEXT = TEL1CLI$(NC%)
      Text5.TEXT = FAXCLI$(NC%)
      Text10 = ""
      '
      For U& = ULTREG&("COTELEF1") To 1 Step -1
        XX$ = REGLEIDO$("COTELEF1", U&)
        If CVI(Mid$(XX$, 81, 2)) = NCXI% Then
          USDESTI$ = UCase$(TRIM$(Mid$(XX$, 83, 24)))
          For KKI% = 1 To Combo1.ListCount
             If TRIM$(Left$(Combo1.List(KKI% - 1), 24)) = USDESTI$ Then
                Combo1.ListIndex = KKI% - 1
                GoTo 15
             End If
          Next KKI%
          '
          If Combo1.ListCount < 1 Then
              Combo1.AddItem AJUSTI$(TRIM$(Mid$(XX$, 83, 24)), 64)
              Combo1.ListIndex = Combo1.ListCount - 1
            Else
              Combo1.ListIndex = 0
          End If
          '
15        Combo1.Refresh
          Text10 = ""
          If Combo1.ListIndex > 0 Then
             Text10 = Combo1.List(0)
          End If
          Exit Sub
          '
        End If
      Next U&
    End If
    '
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "AMACLI02"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
      ACONEC$ = "AMACLI01"
      If EXISTE%(ACONEC$ + ".EXE") < 1 Then
        ACONEC$ = "AMACLI"
      End If
    End If
    Call CONECRUN(ACONEC$, "Alta de Nueva Cuenta de Cliente")
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command15_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = VALACT1$("SELFECHA")
   If VDEVU$ <> "" Then
     Text27 = VDEVU$
   End If
End Sub

Private Sub Command16_Click()
   If USNBR% <> 101 Then Exit Sub
   Call SELECHO("SELFECHA")
   VDEVU$ = VALACT1$("SELFECHA")
   If VDEVU$ <> "" Then
     Text20 = VDEVU$
   End If
End Sub

Private Sub Command17_Click()
    '
    Command17.Enabled = False
    Screen.MousePointer = 11
    JJ& = 0
    Call COPYSTRU("CORTRAB1", "CORTRASE")
    '
    For U& = 1 To ULTREG("CORTRAB1")
      XX$ = REGLEIDO$("CORTRAB1", U&)
      If CVI(Mid$(XX$, 115, 2)) < 1 Then
        JJ& = JJ& + 1
        Call GRAREG("!CORTRASE", XX$, JJ&)
      End If
    Next U&
    Call SETULTREG("!CORTRASE", JJ&)
    Call FRESHECHO("!CORTRASE")
    Screen.MousePointer = 1
    '
    FORSELEC.ABUSQUE = Label15
    Call SELECHO("!CORTRASE")
    NUCOTE$ = AJUSTI$(VALACT1$("!CORTRASE"), 12)
    If TRIM$(NUCOTE$) <> "" Then
       Label15.Caption = NUCOTE$
    End If
    '
    Command17.Enabled = True
    '
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   Call FINAL("*TAROTRA")
   Command18.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   OPX% = ALTERNA%("OT's Pendientes Por Cliente\Detalle Actividad por Cliente y O/T\Detalle de OT's por Cliente y Rango de Fechas\Indice de OT's de UN Cliente\Control de Actividad por Técnico y R.Fechas\Actividad por Clientes (Uno o Todos) y R.Fechas\Actividad por UN Cliente y R.Fechas")
   If OPX% = 1 Then
       Call OTPENCLI
     ElseIf OPX% = 2 Then
       Call OTDETCLI
     ElseIf OPX% = 3 Then
       Call OTTOTCLI
     ElseIf OPX% = 4 Then
       Call INDIOTCLI
     ElseIf OPX% = 5 Then
       Call CONORTRA
     ElseIf OPX% = 6 Then
       Call CONORTRB
     ElseIf OPX% = 7 Then
       Call CONORTRD
   End If
   '
   Command19.Enabled = True
End Sub

Sub INDIOTCLI()
    '
    Call SELECHO("DEUDOR1")
    VDEVU$ = VALACT1$("DEUDOR1")
    If XVALO(VDEVU$) < 1 Then Exit Sub
    
    NCSEL% = XVALO(VDEVU$)
    Call COPYSTRU("CORTRAB1", "CORTRAB2")
    JJ& = 0
    For KU& = 1 To ULTREG&("CORTRAB1")
      XX$ = REGLEIDO$("CORTRAB1", KU&)
      If CVI(Mid$(XX$, 81, 2)) = NCSEL% Then   ' NO ESTA CERRADA
        DESCOT$ = Mid$(XX$, 129, 64)
        Call REPLA(XX$, DESCOT$, 13, 64)
        JJ& = JJ& + 1
        Call GRAREG("CORTRAB2", XX$, JJ&)
      End If
    Next KU&
    Call SETULTREG("CORTRAB2", JJ&)
    Call FILESORT("CORTRAB2", "", 1, 1, "ASC")
    '
    Call HEADERS("CORTRAB2", "")
    Call HEADERS("CORTRAB2", "Cliente: " + RASOCLI$(NCSEL%))
    Call FINAL("*INDIOTC")
    '
End Sub

Sub OTTOTCLI()
    '
    Dim RENGLO$(32)
    REBLA$ = REGBLAN$("LIOTPEN")
    '
    OPX% = COMALTER%("Opciones de Trabajo:\\Solo Pendientes\Todas por Rango de Fechas")
    If OPX% < 1 Then Exit Sub
    '
    DES% = FECHANUM("01/01/00"): HAS% = HOY(HO$): PERIO$ = "Solo O/T's Pendientes"
    '
    'If OPX% = 2 Then
      Call DESHASFECHA(DES1%, HAS1%, PERIO$)
      If PERIO$ = "" Then Exit Sub
      DES% = DES1%: HAS% = HAS1%
    'End If
    '
10  Call SELECHO("DEUDOR1")
    VDEVU$ = VALACT1$("DEUDOR1")
    NCSEL% = XVALO(VDEVU$)
    If NCSEL% < 1 Then Exit Sub
    '
    Call COPYSTRU("CORTRAB1", "CORTRAB2")
    JJ& = 0
    For KU& = 1 To ULTREG&("CORTRAB1")
      XX$ = REGLEIDO$("CORTRAB1", KU&)
      If CVI(Mid$(XX$, 81, 2)) = NCSEL% Then
        If OPX% = 2 Or CVI(Mid$(XX$, 115, 2)) < 33 Then
          FALTA% = CVI(Mid$(XX$, 107, 2))
          If FALTA% >= DES% Then
            If FALTA% <= HAS% Then
              JJ& = JJ& + 1
              Call GRAREG("CORTRAB2", XX$, JJ&)
            End If
          End If
        End If
      End If
    Next KU&
    Call SETULTREG("CORTRAB2", JJ&)
    '
    Call FILESORT("CORTRAB2", "", 4, 5, "ASC")
    JJ& = 0
    NORI& = 0
    For KU& = 1 To ULTREG&("CORTRAB2")
      XX$ = REGLEIDO$("CORTRAB2", KU&)
      NC% = CVI(Mid$(XX$, 81, 2))
      '
      If NC% > 0 Then
        DECLI$ = Left$(XX$, 12)
        DETA$ = TRIM$(Mid$(XX$, 321, 256))
        FALTA% = CVI(Mid$(XX$, 107, 2))
        FEBAJ% = CVI(Mid$(XX$, 115, 2))
        OBSX$ = TRIM$(Mid$(XX$, 193, 128))
        '
        For KKL& = 1 To 32
          RENGLO$(KKL&) = REBLA$
        Next KKL&
        '
        NORI& = NORI& + 1
        Call FRATEXTO(DETA$, 64)
        For LL% = 1 To CARETEX%
          TTYY$ = RENGLO$(LL%)
          Call REPLA(TTYY$, MKS$(NORI&), 1, 4)
          Call REPLA(TTYY$, DECLI$, 5, 12)
          Call REPLA(TTYY$, RETEX$(LL%), 17, 64)
          Call REPLA(TTYY$, MKI$(FALTA%), 81, 2)
          Call REPLA(TTYY$, MKI$(FEBAJ%), 123, 2)
          RENGLO$(LL%) = TTYY$
        Next LL%
        CAREMA% = CARETEX%
        '
        Call FRATEXTO(OBSX$, 40)
        For LL% = 1 To CARETEX%
          TTYY$ = RENGLO$(LL%)
          Call REPLA(TTYY$, MKS$(NORI&), 1, 4)
          Call REPLA(TTYY$, DECLI$, 5, 12)
          Call REPLA(TTYY$, MKI$(FALTA%), 81, 2)
          Call REPLA(TTYY$, RETEX$(LL%), 83, 40)
          RENGLO$(LL%) = TTYY$
        Next LL%
        If CARETEX% > CAREMA% Then CAREMA% = CARETEX%
        '
        For LL% = 1 To CAREMA%
          JJ& = JJ& + 1
          Call GRAREG("LIOTPEN", RENGLO$(LL%), JJ&)
        Next LL%
      End If
    Next KU&
    Call SETULTREG("LIOTPEN", JJ&)
    Call FILESORT("LIOTPEN", "", 2, 2, "ASC")
    Call CIERRARCH("LIOTPEN")
    Call HEADERS("LIOTPEN", "")
    Call HEADERS("LIOTPEN", "Corresponde a: " + TRIM$(RASOCLI$(NCSEL%)) + " - " + PERIO$)
    '
    Call FINAL("*LIOTTOT")
    '
    GoTo 10
    '
End Sub

Sub OTPENCLI()
    '
    Dim RENGLO$(32)
    NCSEL% = 0
    OPX% = COMALTER%("Opciones de Trabajo:\\Pendientes de Un Solo Cliente\Todas las Pendientes")
    If OPX% < 1 Or OPX% > 3 Then Exit Sub
    Call SELECHO("DEUDOR1")
    NCSEL% = XVALO(VALACT1$("DEUDOR1"))
    If NCSEL% < 1 Then Exit Sub
    '
    REBLA$ = REGBLAN$("LIOTPEN")
    Call COPYSTRU("CORTRAB1", "CORTRAB2")
    JJ& = 0
    For KU& = 1 To ULTREG&("CORTRAB1")
      XX$ = REGLEIDO$("CORTRAB1", KU&)
      If CVI(Mid$(XX$, 115, 2)) < 1 Then   ' NO ESTA CERRADA
        If NCSEL% < 1 Or CVI(Mid$(XX$, 81, 2)) = NCSEL% Then
          JJ& = JJ& + 1
          Call GRAREG("CORTRAB2", XX$, JJ&)
        End If
      End If
    Next KU&
    Call SETULTREG("CORTRAB2", JJ&)
    '
    Call FILESORT("CORTRAB2", "", 4, 5, "ASC")
    JJ& = 0
    NORI& = 0
    For KU& = 1 To ULTREG&("CORTRAB2")
      XX$ = REGLEIDO$("CORTRAB2", KU&)
      COTRA$ = Left$(XX$, 12)
      ' NORI& = NORI& + 1
      NC% = CVI(Mid$(XX$, 81, 2))
      '
      If NC% > 0 Then
        DECLI$ = TEL2CLI$(NC%)
           If NCSEL% > 0 Then DECLI$ = COTRA$
        DETA$ = TRIM$(Mid$(XX$, 321, 256))
        FALTA% = CVI(Mid$(XX$, 107, 2))
        FENVI% = CVI(Mid$(XX$, 111, 2)): If FENVI% <= 8224 Then FENVI% = 0
        OBSX$ = TRIM$(Mid$(XX$, 193, 128))
        '
        For KKL& = 1 To 32
          RENGLO$(KKL&) = REBLA$
        Next KKL&
        '
        NORI& = NORI& + 1
        Call FRATEXTO(DETA$, 60)
        For LL% = 1 To CARETEX%
          TTYY$ = RENGLO$(LL%)
          Call REPLA(TTYY$, MKS$(NORI&), 1, 4)
          Call REPLA(TTYY$, DECLI$, 5, 12)
          Call REPLA(TTYY$, RETEX$(LL%), 17, 60)
          Call REPLA(TTYY$, MKI$(FALTA%), 77, 2)
          Call REPLA(TTYY$, MKI$(FENVI%), 79, 2)
          RENGLO$(LL%) = TTYY$
        Next LL%
        CAREMA% = CARETEX%
        '
        Call FRATEXTO(OBSX$, 40)
        For LL% = 1 To CARETEX%
          TTYY$ = RENGLO$(LL%)
          Call REPLA(TTYY$, MKS$(NORI&), 1, 4)
          Call REPLA(TTYY$, DECLI$, 5, 12)
          Call REPLA(TTYY$, MKI$(FALTA%), 81, 2)
          Call REPLA(TTYY$, RETEX$(LL%), 83, 40)
          RENGLO$(LL%) = TTYY$
        Next LL%
        If CARETEX% > CAREMA% Then CAREMA% = CARETEX%
        '
        For LL% = 1 To CAREMA%
          JJ& = JJ& + 1
          Call GRAREG("LIOTPEN", RENGLO$(LL%), JJ&)
        Next LL%
      End If
    Next KU&
    Call SETULTREG("LIOTPEN", JJ&)
    Call FILESORT("LIOTPEN", "", 2, 2, "ASC")
    Call CIERRARCH("LIOTPEN")
    Call HEADERS("LIOTPEN", "")
    If NCSEL% > 0 Then
       Call HEADERS("LIOTPEN", "Corresponde a: " + RASOCLI$(NCSEL%))
    End If
    Call FINAL("*LIOTPEN")
    '
End Sub

Sub OTDETCLI()
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ <> "" Then
      Call SELECHO("DEUDOR1")
      VDEVU$ = VALACT1$("DEUDOR1")
      If XVALO(VDEVU$) >= 0 Then
        NCSEL% = XVALO(VDEVU$)
        JJ& = 0
        '
        Call COPYSTRU("DETOTRA", "LIDETOT")
        For U& = 1 To ULTREG&("DETOTRA")
          XX$ = REGLEIDO$("DETOTRA", U&)
          FEXI% = CVI(Mid$(XX$, 13, 2))
          If FEXI% >= DES% Then
            If FEXI% <= HAS% Then
              NOTRA$ = Left$(XX$, 12)
              TTXX$ = FILTERGETRECORD$("CORTRAB1", 1, NOTRA$)
              NCLIE% = CVI(Mid$(TTXX$, 81, 2))
              If NCLIE% = NCSEL% Then
              ' AGREGAR REGISTRO
                JJ& = JJ& + 1
                Call GRAREG("LIDETOT", XX$, JJ&)
  '              DOTRA$ = ""
  '              NUCOTE$ = AJUSTI$(NOTRA$, 12)
  '              CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
  '                             CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
  '                           Next UI%
  '                           '
  '              RNC$ = RECFILT$("CORTRAB1", 3, MKI$(CICOTE%))
  '              URN& = ULTREG&("CORTRAB1")
  '              For KKU& = 1 To Len(RNC$) Step 4
  '                 RENOTA& = CVS(Mid$(RNC$, KKU&, 4))
  '                 If RENOTA& > 0 Then
  '                    If RENOTA& <= URN& Then
  '                       TTXX$ = REGLEIDO$("CORTRAB1", RENOTA&)
  '                       If Left$(TTXX$, 12) = NUCOTE$ Then
  '                          NCLIE% = CVI(Mid$(TTXX$, 81, 2))
  '                          DCLI$ = TRIM$(TEL2CLI$(NCLIE%))
  '                          DOTRA$ = TRIM$(Mid$(TTXX$, 129, 64))
  '                          GoTo 29
  '                       End If
  '                    End If
  '                 End If
  '              Next KKU&
  '              '
              End If
            End If
          End If
        Next U&
        '
        Call SETULTREG("LIDETOT", JJ&)
        Call FILESORT("LIDETOT", "", 1, 3)
        Call FINAL("*LIDETOT")
        '
      End If
   End If
End Sub
'
Sub CONORTRA()
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ <> "" Then
      RESPOSE$ = TRIM$(UCase$(InputBox$("Responsable")))
      If RESPOSE$ <> "" Then
        REBLA$ = REGBLAN$("CONORTRA")
        JJ& = 0
        '
        For U& = 1 To ULTREG&("DETOTRA")
          XX$ = REGLEIDO$("DETOTRA", U&)
          RESPO$ = TRIM$(UCase$(Mid$(XX$, 73, 4)))
          If RESPO$ = RESPOSE$ Then
            FEXI% = CVI(Mid$(XX$, 13, 2))
            If FEXI% >= DES% Then
              If FEXI% <= HAS% Then
                NOTRA$ = Left$(XX$, 12)
                REFX$ = TRIM$(Mid$(XX$, 15, 50))
                COMIE% = CVI(Mid$(XX$, 65, 2))
                FINI% = CVI(Mid$(XX$, 67, 2))
                HORAX = CVS(Mid$(XX$, 69, 4))
                '
                DOTRA$ = ""
                NUCOTE$ = AJUSTI$(NOTRA$, 12)
                CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                               CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                             Next UI%
                             '
                RNC$ = RECFILT$("CORTRAB1", 3, MKI$(CICOTE%))
                URN& = ULTREG&("CORTRAB1")
                For KKU& = 1 To Len(RNC$) Step 4
                   RENOTA& = CVS(Mid$(RNC$, KKU&, 4))
                   If RENOTA& > 0 Then
                      If RENOTA& <= URN& Then
                         TTXX$ = REGLEIDO$("CORTRAB1", RENOTA&)
                         If Left$(TTXX$, 12) = NUCOTE$ Then
                            NCLIE% = CVI(Mid$(TTXX$, 81, 2))
                            DCLI$ = TRIM$(TEL2CLI$(NCLIE%))
                            DOTRA$ = TRIM$(Mid$(TTXX$, 129, 64))
                            GoTo 29
                         End If
                      End If
                   End If
                Next KKU&
                '
29              YY$ = REBLA$
                Call REPLA(YY$, MKI$(FEXI%), 1, 2)
                Call REPLA(YY$, MKI$(COMIE%), 3, 2)
                Call REPLA(YY$, MKI$(FINI%), 5, 2)
                Call REPLA(YY$, MKS$(HORAX), 7, 4)
                Call REPLA(YY$, NOTRA$, 17, 16)
                Call REPLA(YY$, DCLI$, 33, 12)
                Call REPLA(YY$, DOTRA$, 45, 64)
                Call REPLA(YY$, REFX$, 109, 64)
                JJ& = JJ& + 1
                Call GRAREG("CONORTRA", YY$, JJ&)
                '
              End If
            End If
          End If
        Next U&

        Call SETULTREG("CONORTRA", JJ&)
        Call FILESORT("CONORTRA", "", 1, 3)
        Call HEADERS("CONORTRA", "")
        Call HEADERS("CONORTRA", "Técnico: " + RESPOSE$)

        Call FINAL("*CONORTRA")
        '
      End If
   End If
End Sub
'
Sub CONORTRB()
   DES% = COMESACT%: HAS% = HOY(HOS$): VDEF$ = MKI$(DES%) + MKI$(HAS%) + MKI$(0)
   OBX$ = OBJPLA$("deshasfechcli", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   CLISEL% = CVI(Mid$(OBX$, 5, 2))
   PERIO$ = FECHATEX$(DES%) + " al " + FECHATEX$(HAS%)
   If CLISEL% > 0 Then PERIO$ = RASOCLI$(CLISEL%) + " - " + PERIO$
   '
   If PERIO$ <> "" Then
        REBLA$ = REGBLAN$("CONORTRB")
        JJ& = 0
        '
        For U& = 1 To ULTREG&("DETOTRA")
          XX$ = REGLEIDO$("DETOTRA", U&)
          RESPO$ = TRIM$(Mid$(XX$, 73, 4))
          'If RESPO$ = RESPOSE$ Then
            FEXI% = CVI(Mid$(XX$, 13, 2))
            If FEXI% >= DES% Then
              If FEXI% <= HAS% Then
                NOTRA$ = Left$(XX$, 12)
                REFX$ = TRIM$(Mid$(XX$, 15, 50))
                COMIE% = CVI(Mid$(XX$, 65, 2))
                FINI% = CVI(Mid$(XX$, 67, 2))
                HORAX = CVS(Mid$(XX$, 69, 4))
                '
                DOTRA$ = ""
                NUCOTE$ = AJUSTI$(NOTRA$, 12)
                CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                               CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                             Next UI%
                             '
                RNC$ = RECFILT$("CORTRAB1", 3, MKI$(CICOTE%))
                URN& = ULTREG&("CORTRAB1")
                For KKU& = 1 To Len(RNC$) Step 4
                   RENOTA& = CVS(Mid$(RNC$, KKU&, 4))
                   If RENOTA& > 0 Then
                      If RENOTA& <= URN& Then
                         TTXX$ = REGLEIDO$("CORTRAB1", RENOTA&)
                         If Left$(TTXX$, 12) = NUCOTE$ Then
                            DOTRA$ = TRIM$(Mid$(TTXX$, 129, 64))
                            NCLIE% = CVI(Mid$(TTXX$, 81, 2))
                            If CLISEL% > 0 Then
                               If NCLIE% <> CLISEL% Then
                                  GoTo 49
                               End If
                            End If
                            CLIE$ = TEL2CLI$(NCLIE%)
                            GoTo 29
                         End If
                      End If
                   End If
                Next KKU&
                GoTo 49
                '
29              YY$ = REBLA$
                Call REPLA(YY$, CLIE$, 1, 16)
                Call REPLA(YY$, NOTRA$, 17, 16)
                Call REPLA(YY$, DOTRA$, 33, 64)
                Call REPLA(YY$, REFX$, 97, 64)
                Call REPLA(YY$, MKI$(FEXI%), 161, 2)
                Call REPLA(YY$, MKI$(COMIE%), 163, 2)
                Call REPLA(YY$, MKI$(FINI%), 165, 2)
                Call REPLA(YY$, MKS$(HORAX), 167, 4)
                Call REPLA(YY$, RESPO$, 171, 8)
                JJ& = JJ& + 1
                Call GRAREG("CONORTRB", YY$, JJ&)
                '
              End If
            End If
          'End If
49      Next U&
        '
        Call SETULTREG("CONORTRB", JJ&)
        Call FILESORT("CONORTRB", "", 5, 6)
        Call FILESORT("CONORTRB", "", 1, 2)
        '
80      OPX% = COMALTER%("Formato de Salida:\\Detalle\Resumen")
        If OPX% = 1 Then
             Call HEADERS("CONORTRB", "")
             Call HEADERS("CONORTRB", "Período: " + PERIO$)
             Call FINAL("*CONORTRB")
             GoTo 80
          ElseIf OPX% = 2 Then
             Call COPYSTRU("CONORTRB", "CONORTRC")
             REBLA$ = REGBLAN$("CONORTRB")
             OTANT$ = "": SUMHOR = 0: JJ& = 0
             For II& = 1 To ULTREG&("CONORTRB")
               XX$ = REGLEIDO$("CONORTRB", II&)
               If Left$(XX$, 96) <> OTANT$ Then
                  ZZ$ = REBLA$
                  Call REPLA(ZZ$, OTANT$, 1, 96)
                  Call REPLA(ZZ$, MKS$(SUMHOR), 167, 4)
                  JJ& = JJ& + 1
                  Call GRAREG("CONORTRC", ZZ$, JJ&)
                  OTANT$ = Left$(XX$, 96)
                  SUMHOR = 0
               End If
               SUMHOR = SUMHOR + CVS(Mid$(XX$, 167, 4))
             Next II&
             Call SETULTREG("CONORTRC", JJ&)
             Call FILESORT("CONORTRC", "", 1, 2)
             Call HEADERS("CONORTRC", "")
             Call HEADERS("CONORTRC", "Período: " + PERIO$)
             Call FINAL("*CONORTRC")
             GoTo 80
        End If
        '
   End If
End Sub
'
Sub CONORTRD()
   '
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ <> "" Then
     Call SELECHO("DEUDOR1")
     VDEVU$ = VALACT1$("DEUDOR1")
     If XVALO(VDEVU$) >= 0 Then
        NCSEL% = XVALO(VDEVU$)
        REBLA$ = REGBLAN$("CONORTRA")
        JJ& = 0
        '
        For U& = 1 To ULTREG&("DETOTRA")
           XX$ = REGLEIDO$("DETOTRA", U&)
           RESPO$ = UCase$(TRIM$(Mid$(XX$, 73, 4)))
           If RESPO$ = "EPB" Then RESPO$ = "EB"
           FEXI% = CVI(Mid$(XX$, 13, 2))
           If FEXI% >= DES% Then
             If FEXI% <= HAS% Then
               NOTRA$ = Left$(XX$, 12)
               REFX$ = TRIM$(Mid$(XX$, 15, 50))
               COMIE% = CVI(Mid$(XX$, 65, 2))
               FINI% = CVI(Mid$(XX$, 67, 2))
               HORAX = CVS(Mid$(XX$, 69, 4))
               '
               DOTRA$ = ""
               NUCOTE$ = AJUSTI$(NOTRA$, 12)
               CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                              CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                            Next UI%
                            '
               RNC$ = RECFILT$("CORTRAB1", 3, MKI$(CICOTE%))
               URN& = ULTREG&("CORTRAB1")
               For KKU& = 1 To Len(RNC$) Step 4
                   RENOTA& = CVS(Mid$(RNC$, KKU&, 4))
                   If RENOTA& > 0 Then
                      If RENOTA& <= URN& Then
                         TTXX$ = REGLEIDO$("CORTRAB1", RENOTA&)
                         If Left$(TTXX$, 12) = NUCOTE$ Then
                            NCLIE% = CVI(Mid$(TTXX$, 81, 2))
                            DCLI$ = TRIM$(TEL2CLI$(NCLIE%))
                            DOTRA$ = TRIM$(Mid$(TTXX$, 129, 64))
                            GoTo 29
                         End If
                      End If
                   End If
               Next KKU&
               '
29             If NCLIE% = NCSEL% Then
                 YY$ = REBLA$
                 Call REPLA(YY$, MKI$(FEXI%), 1, 2)
                 Call REPLA(YY$, MKI$(COMIE%), 3, 2)
                 Call REPLA(YY$, MKI$(FINI%), 5, 2)
                 Call REPLA(YY$, MKS$(HORAX), 7, 4)
                 Call REPLA(YY$, RESPO$, 15, 2)
                 Call REPLA(YY$, NOTRA$, 17, 16)
                 Call REPLA(YY$, DCLI$, 33, 12)
                 Call REPLA(YY$, DOTRA$, 45, 64)
                 Call REPLA(YY$, REFX$, 109, 64)
                 JJ& = JJ& + 1
                 Call GRAREG("CONORTRA", YY$, JJ&)
               End If
               '
             End If
           End If
        Next U&
        '
        Call SETULTREG("CONORTRA", JJ&)
        Call FILESORT("CONORTRA", "", 1, 3)
        Call HEADERS("CONORTRA", "")
        Call HEADERS("CONORTRA", "Selección: " + TRIM$(RASOCLI$(NCSEL%)) + " - " + PERIO$)
        Call FINAL("*CONORTRD")
        '
      End If
   End If
End Sub
'
Private Sub Command2_Click()
   Call GRADATCOTE
   Call GRADETCOTE
   Call ENVIAPORMAIL
   '
   On Error Resume Next
   Unload ADJUNARCH
   On Error GoTo 0
   '
End Sub

Private Sub Command20_Click()
   Command20.Enabled = False
   ADJUNARCH.Show 1
   Command20.Enabled = True
End Sub

Private Sub Command21_Click()
   If USNBR% = 101 Then
     If COMALTER%("Está Seguro de Borrar\la Orden de Trabajo '" + TRIM$(Label15) + "' ?\\No, Conservar\Si, Borrar") = 2 Then
       Screen.MousePointer = 11
       JJ& = 0
       ABUS$ = TRIM$(Label15)
       For U& = 1 To ULTREG&("CORTRAB1")
         XX$ = REGLEIDO$("CORTRAB1", U&)
         If TRIM$(Left$(XX$, 12)) <> ABUS$ Then
           JJ& = JJ& + 1
           Call GRAREG("CORTRAB1", XX$, JJ&)
         End If
       Next U&
       Call SETULTREG("CORTRAB1", JJ&)
       Call FRESHECHO("CORTRAB1")
       Screen.MousePointer = 1
     End If
   End If
End Sub

Private Sub Command22_Click()
   Label15 = NEXTORTRA$
   HOII% = HOY(HOS$)
   Label14 = HOS$
   Label24 = Text1
   Label8 = Text2
   Text14 = Label7 + " - " + Text22
   Text8 = Combo1.TEXT    ' Text11
   Text9 = Text21
   Label24 = Text1
   SSTab1.Tab = 1
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   NTAR& = Val(Label28)
   If NTAR& > 0 Then
     TOPLIS = List3.TopIndex
     If COMALTER%("Confirma BORRAR la Tarea Activa ?\\No, Conservar\Si, Borrar") = 2 Then
       Call FILTERDELETE("TARINTER", 1, MKS$(NTAR&))
       Call CIERRARCH("TARINTER")
       IRESPO% = 0
       For KKI% = 1 To 5
          If Option3(KKI%).Value = True Then IRESPO% = KKI%
       Next KKI%
       Call Option4_Click(IRESPO%)
       Text29 = ""
       Text28 = ""
       Text13 = ""
       Text31 = ""
       Label28 = ""
     End If
     On Error Resume Next
     List3.TopIndex = TOPLIS
     Call List3_DblClick
   End If
   Command23.Enabled = True
End Sub

Private Sub Command24_Click()
   '
   FORSELEC.ABUSQUE = Label7
   Call SELECHO("COTELEF3")
   NUCOTE$ = AJUSTI$(VALACT1$("COTELEF3"), 12)
   If TRIM$(NUCOTE$) <> "" Then
       Label7.Caption = NUCOTE$
       Call CARDATCOTE
       Call CARDETCOTE
       Unload ADJUNARCH
    End If
    '
End Sub

Sub Command25_Click()
   Command25.Enabled = False
   Call BLANARCH("LITARPEN")
   '
   IRESPO% = 0
   For KKI% = 1 To 9
     If Option4(KKI%).Value = True Then IRESPO% = KKI%
   Next KKI%
   IRESX$ = Option4(IRESPO%).Caption
   Call HEADERS("LITARPEN", "")
   Call HEADERS("LITARPEN", "Responsable: " + IRESX$)
   '
   TODOS% = 0
   If Option4(9).Value = True Then ' SELECCIONA TODOS
      TODOS% = 1
   End If
   '
   REBLA$ = REGBLAN$("LITARPEN")
   '
   JJ& = 0
   For KU& = 1 To List3.ListCount
     YY$ = List3.List(KU& - 1)
     NTAR& = Val(Mid$(YY$, 1, 5))
     FETEXX$ = Mid$(YY$, 9, 8)
     DETEX$ = Mid$(YY$, 19, 64)
     OBSEX$ = Mid$(YY$, 97, 64)
     IRESPOX$ = ""
     If TODOS% > 0 Then
        IRESPOX$ = Mid$(YY$, 86, 10)
     End If
     '
     ZZ$ = REBLA$
     Call REPLA(ZZ$, MKS$(NTAR&), 1, 4)
     Call REPLA(ZZ$, FETEXX$, 5, 12)
     Call REPLA(ZZ$, DETEX$, 17, 80)
     Call REPLA(ZZ$, IRESPOX$, 97, 16)
     Call REPLA(ZZ$, OBSEX$, 113, 64)
     JJ& = JJ& + 1
     Call GRAREG("LITARPEN", ZZ$, JJ&)
   Next KU&
   '
   Call SETULTREG("LITARPEN", JJ&)
   Call CONECRUN("*LITARPE", "Listado de Tareas Pendientes por Responsable")
   '
   Command25.Enabled = True
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   OPCOTEL.Show 1
   Command26.Enabled = True
End Sub

Private Sub Command27_Click()
    '
    Command27.Enabled = False
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ <> "" Then
      Screen.MousePointer = 11
      JJ& = 0
      TOT& = 0: PEN& = 0
      CIERRCTEL.Label2 = TRIM$(Str$(DES%))
      CIERRCTEL.Label3 = TRIM$(Str$(HAS%))
      CIERRCTEL.List1.Clear     ' EPB - NECESARIO PORQUE SI NO AL VOLVER A ENTRAR DUPLICA LA CARGA
      For U& = 1 To ULTREG&("COTELEF1")
        XX$ = REGLEIDO$("COTELEF1", U&)
        FEMI% = CVI(Mid$(XX$, 107, 2))
        FECI% = CVI(Mid$(XX$, 251, 2))
        If FEMI% >= DES% Then
          If FEMI% <= HAS% Then
            TOT& = TOT& + 1
            If FECI% <= 8224 Then          ' NO ESTÁ CERRADA
              PEN& = PEN& + 1
              NUCOTE$ = Mid$(XX$, 1, 12)
              FECLIRE$ = Mid$(XX$, 13, 66)
              CIERRCTEL.List1.AddItem (NUCOTE$ + " - " + FECLIRE$)
            End If
          End If
        End If
      Next U&
      '
      CIERRCTEL.Text2 = TOT&
      CIERRCTEL.Text3 = PEN&
      Screen.MousePointer = 1
      '
      CIERRCTEL.Show 1
      '
    End If
    '
    Command27.Enabled = True
    '
End Sub

Private Sub Command28_Click()
    Call SELECHO("TARINTER")
    NTAR& = Val(VALACT1$("TARINTER"))
    If NTAR& > 0 Then
      Label28 = TRIM$(Str$(NTAR&))
    End If
End Sub

Private Sub Command29_Click()
    Command29.Enabled = False
    TTXX$ = REGBLAN$("TARINTER")
    TTYX$ = TTXX$
    UREX& = ULTREG&("TARINTER")
    If UREX& > 0 Then
      TTYX$ = REGLEIDO$("TARINTER", UREX&)
    End If
    UTAR& = CVS(Left$(TTYX$, 4))
    TOPLIS = List3.TopIndex
    '
    NTAR& = Val(Label28)
    If NTAR& > 0 Then
      If TRIM$(Text29) = "" Then
        Call FRATEXTO(Text28, 52)
        Text29 = RETEX$(1)
        Text29.Refresh
      End If
      Call REPLA(TTXX$, MKS$(NTAR&), 1, 4)
        TTYY$ = TRIM$(Label29) + " - " + TRIM$(Text29)
      Call REPLA(TTXX$, TTYY$, 5, 60)
        FEXI% = FECHANUM(Label29)
      Call REPLA(TTXX$, MKI$(FEXI%), 65, 2)
        FEXI% = FECHANUM(Text31)
      Call REPLA(TTXX$, MKI$(FEXI%), 67, 2)
        IRESPO% = 0
        For KKI% = 1 To 5
          If Option3(KKI%).Value = True Then IRESPO% = KKI%
        Next KKI%
      Call REPLA(TTXX$, Chr$(IRESPO%), 69, 1)
      Call REPLA(TTXX$, MKI$(0), 75, 2)
      Call REPLA(TTXX$, Text29, 77, 52)
      Call REPLA(TTXX$, Text28, 129, 512)
      Call REPLA(TTXX$, Text13, 641, 384)
      Call FILTERSETRECORD("TARINTER", 1, MKS$(NTAR&), TTXX$)
      '
      'Call GENTARPEN
      Call CIERRARCH("TARINTER")
      Call Option4_Click(IRESPO%)
      '
      If NTAR& <= UTAR& Then
        On Error Resume Next
        List3.TopIndex = TOPLIS
      End If
    End If
    '
    Command29.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    FORSELEC.ABUSQUE = Label7
    Call SELECHO("COTELEF1")
    NUCOTE$ = AJUSTI$(VALACT1$("COTELEF1"), 12)
    If TRIM$(NUCOTE$) <> "" Then
       Label7.Caption = NUCOTE$
       Call CARDATCOTE
       Call CARDETCOTE
       Unload ADJUNARCH
    End If
    '
End Sub

Private Sub Command30_Click()
   '
   Command30.Enabled = False
   Option3(0).Value = True
   Option4(0).Value = True
   '
   Label28 = TRIM$(Str$(PROTARIN&))
   HOII% = HOY(HOS$)
   Label29 = HOS$
   Text29 = ""
   Text28 = ""
   Text13 = ""
   Text31 = ""
   Label34 = ""
   '
   Command30.Enabled = True
   Text28.SetFocus
   '
End Sub

Function PROTARIN&()
   '
   PROTARIN& = 1
10 URE& = ULTREG&("TARINTER")
   If URE& > 0 Then
     F$ = REGLEIDO$("TARINTER", URE&)
     PROTARIN& = 1 + CVS(Left$(F$, 4))
   End If
   RFF$ = RECFILT$("TARINTER", 1, MKS$(PROTARIN&))
   If RFF$ <> "" Then
      Call MENSERR(24, "Error al Determinar Número Próxima Tarea")
      GoTo 10
   End If
   '
End Function

Private Sub Command31_Click()
   Command31.Enabled = False
   ABUSQUE$ = TRIM$(InputBox$("Argumento", "Búsqueda por Texto", ABUSQUE$))
   If ABUSQUE$ <> "" Then
     Option4(9).Value = True
     For KU& = 1 To List3.ListCount
       If InStr(UCase$(List3.List(KU& - 1)), UCase$(ABUSQUE$)) > 0 Then
         List3.ListIndex = KU& - 1
         List3.SetFocus
         GoTo 99
       End If
     Next KU&
     MsgBox "Argumento no Encontrado"
   End If
   '
99 Command31.Enabled = True
End Sub

Private Sub Command32_Click()
   Command32.Enabled = False
10 OPX% = ALTERNA%("Todas de un Unico Cliente\Solo Pendientes de Todos\Todas de Todos\-\Clasificacion por Motivo")
   If OPX% = 1 Then
       GoTo 20
     ElseIf OPX% = 2 Then
       Call LICOTELP(1)
     ElseIf OPX% = 3 Then
       Call LICOTELP(2)
     ElseIf OPX% = 5 Then
       Call CLASIMOT
       GoTo 10
   End If
   GoTo 99
   '
20 Call SELECHO("DEUDOR1")
   NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCXI% > 0 Then
      Call DESHASFECHA(DES%, HAS%, PERIO$)
      If PERIO$ <> "" Then
         REBLA$ = REGBLAN$("LICOTEL")
         JJ& = 0
         For U& = 1 To ULTREG&("COTELEF1")
           XX$ = REGLEIDO$("COTELEF1", U&)
           If CVI(Mid$(XX$, 81, 2)) = NCXI% Then
             FEMI% = CVI(Mid$(XX$, 107, 2))
             If FEMI% >= DES% Then
               If FEMI% <= HAS% Then
                 TTYY$ = REBLA$
                 Call REPLA(TTYY$, MKI$(NCXI%), 1, 2)
                   FEMI% = CVI(Mid$(XX$, 107, 2))
                 Call REPLA(TTYY$, MKI$(FEMI%), 3, 2)
                   HOMI% = CVI(Mid$(XX$, 109, 2))
                 Call REPLA(TTYY$, MKI$(HOMI%), 5, 2)
                   NUCOTE$ = Mid$(XX$, 1, 12)
                 Call REPLA(TTYY$, NUCOTE$, 7, 12)
                   ASUN$ = Mid$(XX$, 111, 64)
                 Call REPLA(TTYY$, ASUN$, 19, 64)
                   PARA$ = Mid$(XX$, 83, 24)
                 Call REPLA(TTYY$, PARA$, 83, 24)
                   FECI% = CVI(Mid$(XX$, 251, 2))
                 Call REPLA(TTYY$, MKI$(FECI%), 107, 2)
                 JJ& = JJ& + 1
                 Call GRAREG("LICOTEL", TTYY$, JJ&)
               End If
             End If
           End If
         Next U&
         Call SETULTREG("LICOTEL", JJ&)
         Call HEADERS("LICOTEL", "")
         Call HEADERS("LICOTEL", "Seleccion: Todas (" + TRIM$(RASOCLI$(NCXI%)) + ") - Período: " + PERIO$)
         Call CONECRUN("*LICOTEL", "Listado de Consultas Telefonicas")
      End If
   End If
99 Command32.Enabled = True
End Sub

Sub LICOTELP(MODO%)
   '
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ <> "" Then
      REBLA$ = REGBLAN$("LICOTEL")
      JJ& = 0
      For U& = 1 To ULTREG&("COTELEF1")
        XX$ = REGLEIDO$("COTELEF1", U&)
        FEMI% = CVI(Mid$(XX$, 107, 2))
        FECI% = CVI(Mid$(XX$, 251, 2))
        If FECI% <= 8224 Or MODO% = 2 Then        ' NO ESTÁ CERRADA
          If FEMI% >= DES% Then
            If FEMI% <= HAS% Then
              NCXI% = CVI(Mid$(XX$, 81, 2))
              TTYY$ = REBLA$
              Call REPLA(TTYY$, MKI$(NCXI%), 1, 2)
              Call REPLA(TTYY$, MKI$(FEMI%), 3, 2)
                HOMI% = CVI(Mid$(XX$, 109, 2))
              Call REPLA(TTYY$, MKI$(HOMI%), 5, 2)
                NUCOTE$ = Mid$(XX$, 1, 12)
              Call REPLA(TTYY$, NUCOTE$, 7, 12)
                ASUN$ = Mid$(XX$, 111, 64)
              Call REPLA(TTYY$, ASUN$, 19, 64)
                PARA$ = Mid$(XX$, 83, 24)
              Call REPLA(TTYY$, PARA$, 83, 24)
              JJ& = JJ& + 1
              Call GRAREG("LICOTEL", TTYY$, JJ&)
            End If
          End If
        End If
      Next U&
      Call SETULTREG("LICOTEL", JJ&)
      Call FILESORT("LICOTEL", "", 1, 3, "ASC")
      Call HEADERS("LICOTEL", "")
      If MODO% = 1 Then Call HEADERS("LICOTEL", "Pendientes - Período: " + PERIO$)
      If MODO% = 2 Then Call HEADERS("LICOTEL", "Todas - Período: " + PERIO$)
      Call CONECRUN("*LICOTELP", "Listado de Consultas Telefonicas")
   End If
   '
End Sub
'
Sub CLASIMOT()
   '
   Dim CANPORMOT%(4)
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ <> "" Then
      REBLA$ = REGBLAN$("LICOTEL")
      JJ& = 0: CANTOTA% = 0: FIN& = ULTREG("COTELEF1")
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        XX$ = REGLEIDO$("COTELEF1", U&)
        FEMI% = CVI(Mid$(XX$, 107, 2))
        If FEMI% >= DES% Then
          If FEMI% <= HAS% Then
            MOTI% = Asc(Mid$(XX$, 253, 1))
            If MOTI% > 3 Then MOTI% = 4         ' NO CLASIFICADOS
            CANPORMOT%(MOTI%) = CANPORMOT%(MOTI%) + 1
            CANTOTA% = CANTOTA% + 1
          End If
        End If
      Next U&
      '
      If CANTOTA% < 1 Then Exit Sub
      '
      Call COPYSTRU("DEUDOR1", "CLASMOT")
      REBLA$ = REGBLAN$("DEUDOR1")
      PORCEX$ = " (" + TRIM$(FORMATNUM$(100 * CANPORMOT%(0) / CANTOTA%, "F4.1")) + " %)"
      XX$ = MKI$(CANPORMOT%(0)) + "Consultas Puras" + PORCEX$
      Call GRAREG("!CLASMOT", XX$, 1)
      PORCEX$ = " (" + TRIM$(FORMATNUM$(100 * CANPORMOT%(1) / CANTOTA%, "F4.1")) + " %)"
      XX$ = MKI$(CANPORMOT%(1)) + "Reclamos Varios" + PORCEX$
      Call GRAREG("!CLASMOT", XX$, 2)
      PORCEX$ = " (" + TRIM$(FORMATNUM$(100 * CANPORMOT%(2) / CANTOTA%, "F4.1")) + " %)"
      XX$ = MKI$(CANPORMOT%(2)) + "Solicita Nueva Funcion" + PORCEX$
      Call GRAREG("!CLASMOT", XX$, 3)
      PORCEX$ = " (" + TRIM$(FORMATNUM$(100 * CANPORMOT%(3) / CANTOTA%, "F4.1")) + " %)"
      XX$ = MKI$(CANPORMOT%(3)) + "Otros Casos" + PORCEX$
      Call GRAREG("!CLASMOT", XX$, 4)
      PORCEX$ = " (" + TRIM$(FORMATNUM$(100 * CANPORMOT%(4) / CANTOTA%, "F4.1")) + " %)"
      XX$ = MKI$(CANPORMOT%(4)) + "No Clasificadas" + PORCEX$
      Call GRAREG("!CLASMOT", XX$, 5)
      XX$ = MKI$(0) + " "
      Call GRAREG("!CLASMOT", XX$, 6)
      XX$ = MKI$(CANTOTA%) + "Total del Periodo"
      Call GRAREG("!CLASMOT", XX$, 7)
      Call SETULTREG("!CLASMOT", 7)
      Call FRESHECHO("!CLASMOT")
      Call SELECHO("!CLASMOT/C-Tels por Motivo " + PERIO$)
      '
    End If
    '
End Sub

Private Sub Command33_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = VALACT1$("SELFECHA")
   If VDEVU$ <> "" Then
     Text31 = VDEVU$
   End If
End Sub

Private Sub Command34_Click_VIE()
   Command34.Enabled = False
   If DERACCE%("RECIELEC", "Recibo Electrónico por Mail") < 2 Then GoTo 90
   '
   OPX% = ALTERNA%("A.C.M.\TELEBIT\ESCORIAL\ELECTRO-DIEL\TANIT\RVC\SILCAR - TELSO\REGALIA\ASEMBLI - GRAMY\J Y M\COMPREFACIL")
   FEX = DIPOST(FINMESACT)
6  PROVEN$ = FECHATEX$(FEX)
   If Left$(PROVEN$, 2) <> "01" Then
      FEX = DIPOST(FEX)
      GoTo 6
   End If
   RUTAPAR$ = ""
   Select Case OPX%
     Case 1    ' ACM
         GoTo 90
         'IMPO# = 368.4
         'PROVEN$ = "01/12/06"
         'NCX% = 62
     Case 2     ' TELEBIT
         IMPO# = 274.54
         NCX% = 65
     Case 3   ' ESCORIAL
         IMPO# = 478.97
         NCX% = 70
     Case 4   ' ELECTRO-DIEL
         GoTo 90
         'IMPO# = 222#
         'NCX% = 68
     Case 5    ' TANIT
         IMPO# = 2130#
         NCX% = 66
     Case 6   ' RVC
         IMPO# = 374.06
         NCX% = 69
         RUTAPAR$ = "X:\RVC\FLEXOFT\WINDOWS\"
     Case 7   ' SILCAR
         IMPO# = 384.78
         NCX% = 72
     Case 8   ' REGALIA
         GoTo 90
         'IMPO# = 81.76
         'NCX% = 71
         'RUTAPAR$ = "X:\REGALIA\FLEXOFT\WINDOWS\"
     Case 9    ' ASEMBLI
         IMPO# = 1299.56
         NCX% = 53
     Case 10    ' J Y M
         IMPO# = 320
         NCX% = 76
     Case 11    ' COMPREFACIL
         IMPO# = 63.49
         NCX% = 88
     Case Else
         GoTo 90
   End Select
   '
   IMPOX$ = FORMATNUM$(IMPO#, "F12.2")
   CLIE$ = RASOCLI$(NCX%)
10 FVENCI$ = InputBox$("Fecha Próximo Vencimiento:", CLIE$ + " - $ " + TRIM$(IMPOX$), PROVEN$)
   If FVENCI$ = "" Then GoTo 90 ' PARA PODER CANCELAR
   If Len(FVENCI$) <> 8 Or FECHANUM(FVENCI$) < 1 Then
      Call MENSERR(24, "Fecha no Válida - Corrija")
      GoTo 10
   End If
   '
   AAX$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
   TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
   Mid$(AAX$, 1, 128) = TTXX$
   '
   ' ABONO MENSUAL
   BBX$ = Space$(128)
   Call REPLA(BBX$, "02", 1, 2)
   Call REPLA(BBX$, "ABONO MENSUAL", 3, 20)
   Call REPLA(BBX$, FVENCI$, 23, 8)
   Call REPLA(BBX$, "$", 31, 4)
   Call REPLA(BBX$, IMPOX$, 35, 12)
   SUCAR& = 0: For I% = 1 To 120
                   SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
               Next I%
   Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
   CCX$ = ENCRIPTA$(BBX$)
   Mid$(AAX$, 129, 128) = CCX$
   '
   RUTAX$ = "F:\ANOVE\"
   Call SAVEFILE(RUTAX$ + "MSTDRV04.EXE", AAX$)
   If RUTAPAR$ <> "" Then
      Call SAVEFILE(RUTAPAR$ + "MSTDRV04.EXE", AAX$)
   End If
   '
   ' CREACIÓN DEL ZIP-FILE
   HOII% = HOY(HOS$)
   AZIP0$ = "UP" + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 1, 2) + "-"
   IDX% = 0
12 AZIP$ = AZIP0$ + TRIM$(Str$(IDX%))
   If EXISTE%("F:\ANOVE\" + AZIP$ + ".ZIP") > 0 Then
     IDX% = IDX% + 1
     GoTo 12
   End If
   AZIPSH$ = AZIP$
   While InStr(AZIPSH$, "\") > 0
     AZIPSH$ = Mid$(AZIPSH$, 2)
   Wend
   '
   RUTAWIN$ = "C:\Archivos de programa\WinZip\WINZIP32.EXE"
   If EXISTE%(RUTAWIN$) < 1 Then
      RUTAWIN$ = "D:\Archivos de programa\WinZip\WINZIP32.EXE"
   End If
   On Error Resume Next
   RETURNVALUE = Shell(RUTAWIN$, 4)
   If Err Then
     On Error GoTo 0
     MsgBox "No Existe " + RUTAWIN$
     GoTo 90
   End If
   On Error GoTo 0
   On Error Resume Next
   AppActivate RETURNVALUE   ' : AppActivate RETURNVALUE
   If Err Then
      On Error GoTo 0
      MsgBox "Imposible Activar WINZIP32"
      GoTo 90
   End If
   On Error GoTo 0
   SendKeys "^N", True
   SendKeys "F:\ANOVE\" + AZIP$ + Chr$(13), True
   '
   TTXX$ = ""
   TTXX$ = TTXX$ + Chr$(34) + RUTAX$ + "MSTDRV04.EXE" + Chr$(34) + Chr$(32)
   SendKeys "%A", True
   SendKeys TRIM$(TTXX$) + Chr$(13), True
   '
   AppActivate "WinZip - " + AZIPSH$ + ".ZIP"
   AppActivate "WinZip - " + AZIPSH$ + ".ZIP"
   '
16 On Error Resume Next
   AppActivate "WINZIP"
   If Err Then
       Resume 20
     Else
       SendKeys "%{F4}", True
       GoTo 16
   End If
   '
20 On Error GoTo 0
   MADJUN$ = "F:\ANOVE\" + AZIP$ + ".ZIP"
   '
   
   '
   ' GENERACION DEL MAIL
   HOII% = HOY(HOS$)
   ASUNTO$ = "FLEXOFT Soluciones ("
   CLAX$ = RANDSTRING$(-6)
   ASUNTO$ = ASUNTO$ + CLAX$ + "/" + CONTRACLA$(CLAX$) + ")"
   '
   TEXTOMAI$ = "Estamos enviando adjunto al presente recibo electrónico" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "correspondiente a su pago recibido en la fecha." + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + CRLF$ + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "Para instalar el recibo electrónico, proceda así:" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "1) GUARDE el archivo adjunto 'MSTDRV04.EXE' en la carpeta \FLEXOFT\NOVEDAD" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "   del disco servidor (sobre-escribiendo el existente)" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "2) Puede también DES-COMPRIMIR el archivo adjunto '" + AZIP$ + "' en la carpeta" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "   \FLEXOFT\NOVEDAD. Esto es ALTERNATIVO a lo indicado en 1) ya que algunos servi-" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "   dores eliminan los '.EXE' adjuntos" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "3) Cerrar el FLEXOFT y volver a ingresar. Luego de introducir el" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "   usuario y la contraseña, aparecerá un mensaje de actualizaciones en" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "   el cual deberá cliquear sobre el botón 'Si'." + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "4) La instalación finaliza con el mensaje 'Actualizaciones" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "   Instaladas '." + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "En caso de dudas, solicite asistencia telefónica a nuestra mesa de" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "ayuda: (011)-4524-1284 / 4522-2048 / 4541-1515" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "Agradecemos su transferencia." + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "Quedamos a disposición para cualquier consulta adicional, aprovechando" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "para saludarlos" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "                       Cordialmente" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "===================================================================" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "FLEXOFT Soluciones Informáticas Integrales" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "Bauness 2170 - 5to. 'D' - 1431 Buenos Aires - Argentina" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "Tel (00-54-11) 4524-1284 / 4522-2048 - e-mail: info@flexoft.com.ar" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "===================================================================" + CRLF$
   '
   MAI_DESTI$ = RASOCLI$(NCX%)
     DESTIMA$ = TRIM$(EMAILCLI$(NCX%))
     If DESTIMA$ = "" Or InStr(DESTIMA$, "@") < 1 Then
       Call MENSERR(24, "Imposible Enviar.\Dirección de Mail Inexistente o no Válida.")
       GoTo 90
     End If
   MAI_DIREL$ = DESTIMA$
   MAI_ASUNT$ = ASUNTO$
   MAI_TEXTO$ = TEXTOMAI$
   MAI_ADJUN$(0) = "F:\ANOVE\MSTDRV04.EXE"
   MAI_ADJUN$(1) = MADJUN$
   '
   Call GENERAMAIL
   '
   ' REPLICA A LA DIRECCION DE CONTACTO
   ' para enviar aviso a la direccion de contacto
   DESTIMB$ = TRIM$(CONTACTCLI$(NCX%))
   If DESTIMB$ <> "" Then
      If InStr(DESTIMB$, "@") > 1 Then
        ASUNTO$ = "FLEXOFT Soluciones ("
        CLAX$ = RANDSTRING$(-6)
        ASUNTO$ = ASUNTO$ + CLAX$ + "/" + CONTRACLA$(CLAX$) + ")"
        '
        TEXTOMAI$ = "Informamos que desde nuestro sector de Cuentas Corrientes se ha enviado " + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Recibo Electrónico correspondiente a su transferencia." + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Teniendo en cuenta las características de configuración de esta vues-" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "tra cuenta de correo, la que bloquea la recepción de mensajes con" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "archivos adjuntos, el referido envío se ha remitido a la direccion" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + DESTIMA$ + " de donde pueden Uds. bajarlo." + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
        '
        TEXTOMAI$ = TEXTOMAI$ + "En caso de dudas, soliciten asistencia telefónica a nuestra mesa de" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "ayuda: (011)-4524-1284 / 4522-2048 - interno 32" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Agradecemos su Transferencia." + CRLF$
        '
        TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Quedamos a disposición para cualquier consulta adicional, aprovechando" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "para saludarlos" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "                       Cordialmente" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "=====================================================================" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "FLEXOFT Soluciones Informáticas Integrales" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Miembro de la Cámara de Empresas de Software y Servicios Informáticos" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Bauness 2170 - 5to. 'D' - 1431 Buenos Aires - Argentina" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Tel (00-54-11) 4524-1284 / 4522-2048 - e-mail: info@flexoft.com.ar" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "=====================================================================" + CRLF$
        '
        MAI_DIREL$ = DESTIMB$     ' "fgrasso@flexoft.com.ar"
        MAI_ASUNT$ = ASUNTO$
        MAI_TEXTO$ = TEXTOMAI$
        MAI_ADJUN$(0) = ""
        MAI_ADJUN$(1) = ""
        '
        Call GENERAMAIL
        '
     End If
   End If
   '
90 Command34.Enabled = True

End Sub


Private Sub Command34_Click()
   Command34.Enabled = False
   Call SELECHO("DEUDOR1")
   VDEVU$ = VALACT1$("DEUDOR1")
   If VDEVU$ <> "" Then
      NCLI% = XVALO(VDEVU$)
      LANID$ = TRIM$(HORARCLI$(NCLI%))
      If Len(TRIM$(LANID$)) <> 9 Then
         Call MENSERR(24, "Imposible Enviar Recibo Electronico.\LAN-Id Inexistente o no Válido.")
         GoTo 99
      End If
      CARELEC.Label6 = ""
      CARELEC.Label2 = RASOCLI$(NCLI%)
5     CARELEC.Show 1
      If CARELEC.Label6 <> "OK" Then GoTo 99
      IMPO# = XVALO(CARELEC.Text2)
      If IMPO# < 10 Then
         Call MENSERR(24, "Importe No Válido")
         GoTo 5
      End If
      IMPOX$ = FORMATNUM$(IMPO#, "F12.2")
      '
      FEPROVEN% = FECHANUM(CARELEC.Text1)
      If FEPROVEN < FECHANUM("30/06/08") Then
         Call MENSERR(24, "Fecha no Válida")
         GoTo 5
      End If
      HOII% = HOY(HO$)
      If FEPROVEN% < HOII% Then
         If COMALTER%("Esa Fecha Está Vencida.\   \¿ Está Seguro de Esa Fecha ?\\No, Corregir\Si, Confirmar") <> 2 Then GoTo 5
      End If
      If DIENTRE%(HOII%, FEPROVEN%) > 30 Then
         If COMALTER%("Esa Fecha es de Más de 30 Días.\   \¿ Está Seguro de Esa Fecha ?\\No, Corregir\Si, Confirmar") <> 2 Then GoTo 5
      End If
      '
      FEVENLI% = 0
      IMPOLI# = XVALO(CARELEC.Text4)
      If IMPOLI# > 0 Then
        IMPOLIX$ = FORMATNUM$(IMPOLI#, "F12.2")
        FEVENLI% = FECHANUM(CARELEC.Text3)
        If FEVENLI% < FECHANUM("30/06/08") Then
           Call MENSERR(24, "Fecha no Válida")
           GoTo 5
        End If
        HOII% = HOY(HO$)
        If FEVENLI% < HOII% Then
           If COMALTER%("Esa Fecha Está Vencida.\   \¿ Está Seguro de Esa Fecha ?\\No, Corregir\Si, Confirmar") <> 2 Then GoTo 5
        End If
        '
        CONCEPTO$ = UCase$(TRIM$(CARELEC.Text5))
        If Len(CONCEPTO$) < 6 Then
          Call MENSERR(24, "Concepto no Válido")
          GoTo 5
        End If
      End If
      '
      '
      ' GENERA RECIBO ELECTRONICO
      INI% = 129
      FVENCI$ = FECHATEX$(FEPROVEN%)
      AAX$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
      TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
      Mid$(AAX$, 1, 128) = TTXX$
      '
      ' ABONO MENSUAL o LICENCIA
      If IMPO# > 0 Then
         BBX$ = Space$(128)
         Call REPLA(BBX$, "02", 1, 2)
         Call REPLA(BBX$, "ABONO MENSUAL", 3, 20)
         Call REPLA(BBX$, FVENCI$, 23, 8)
         Call REPLA(BBX$, "$", 31, 4)
         Call REPLA(BBX$, IMPOX$, 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                         SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
         Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
         CCX$ = ENCRIPTA$(BBX$)
         Mid$(AAX$, INI%, 128) = CCX$
         INI% = INI% + 128
      End If
      '
      If IMPOLI# > 0 Then
         BBX$ = Space$(128)
         Call REPLA(BBX$, "01", 1, 2)
         Call REPLA(BBX$, CONCEPTO$, 3, 20)
           FVENCI$ = FECHATEX$(FEVENLI%)
         Call REPLA(BBX$, FVENCI$, 23, 8)
         Call REPLA(BBX$, "$", 31, 4)
         Call REPLA(BBX$, IMPOLIX$, 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                         SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
         Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
         CCX$ = ENCRIPTA$(BBX$)
         Mid$(AAX$, INI%, 128) = CCX$
         INI% = INI% + 128
      End If
      '
      RUTAX$ = "F:\ANOVE\"
      Call SAVEFILE(RUTAX$ + LANID$ + ".REL", AAX$)
      Call MANDALAWEB(RUTAX$ + LANID$ + ".REL", LANID$)
      '
      ' AQUI REPLICAR EN LA CARPETA DEL CLIENTE
      '
   End If
   '
99 Command34.Enabled = True
End Sub

Private Sub Command35_Click()
   Command35.Enabled = False
   Call REVIREGIS
   Command35.Enabled = True
End Sub

'
Private Sub Command4_Click()
   Call GRADATCOTE
   Call GRADETCOTE
End Sub

Private Sub Command5_Click()
   Call GRADATCOTE
   Call GRADETCOTE
   Call PRICOTEL(TRIM$(Label7.Caption))
End Sub

Private Sub Command6_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command7_Click()
   '
   On Error Resume Next
   Unload ADJUNARCH
   On Error GoTo 0
   '
   Label7.Caption = NEXTCOTELNU$
   HOII% = HOY(HOS$)
   Text12.TEXT = HOS$
   Text6.TEXT = Time$
   Text7.TEXT = ""
   Text21.TEXT = ""
   Text22.TEXT = ""
   Text23.TEXT = ""
   Label25 = ""
   Image2.Visible = True
   Image1.Visible = False
   Option5(0).Value = True
   Call Command1_Click
End Sub

Private Sub Command8_Click()
   '
   Call SELECHO("DEUDOR1")
   NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCXI% < 1 Then Exit Sub
   '
   Label24 = NCXI%
   Label8 = RASOCLI$(NCXI%)
   '
   Label15 = NEXTORTRA$
   HOII% = HOY(HOS$)
   Label14 = HOS$
   '
   Call SHOWDETHOR
End Sub

Private Sub Command9_Click()
   Call GRADATOTRA
   'Call GRADETOTRA
End Sub

Private Sub Form_Load()
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then
     Call FINAL("")
     Exit Sub
   End If
   '
   SSTab1.Tab = 0
   Label7.Caption = NEXTCOTELNU$
   HOII% = HOY(HOS$)
   Text12.TEXT = HOS$
   Text6.TEXT = Time$
   Text7.TEXT = ""
   Text21.TEXT = ""
   Text22.TEXT = ""
   Text23.TEXT = ""
   If USNBR% = 101 Then
     Text20.Enabled = True
     Call Option4_Click(0)
   End If
   '
   Call TARPENDE.CARLISTAPEN
   TARPENDE.Show
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Frame2_DblClick()
    NC% = XVALO(Text1)
    If NC% > 0 Then
      IDLAN$ = TRIM$(HORARCLI$(NC%))
      DSTX$ = TRIM$(Mid$(Combo1.TEXT, 65, 12))
      If IDLAN$ <> "" Then
        IDXX$ = ENCRIPTA$(IDLAN$)
        Call SAVEFILE(LUCOM$ + "BKSEVOL.DAT", IDXX$)
        Call MANDALAWEB(LUCOM$ + "BKSEVOL.DAT", IDLAN$ + ";" + DSTX$)
        Call FILEKILL(LUCOM$ + "BKSEVOL.DAT")
      End If
    End If
End Sub

Private Sub Image1_DblClick()
   Call Label25_DblClick
End Sub

Private Sub Image2_DblClick()
   If InStr(USERNAME$, "BAUER") > 0 Then
     If Image2.Visible = True Then
       If TRIM$(Label7) <> "" Then
         NUCOTE$ = AJUSTI$(Label7.Caption, 12)
         CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                        CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                      Next UI%
                      '
         RNC$ = RECFILT$("COTELEF1", 3, MKI$(CICOTE%))
         URN& = ULTREG&("COTELEF1")
         '
         For U& = 1 To Len(RNC$) Step 4
            RENOTA& = CVS(Mid$(RNC$, U&, 4))
            If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                  TTXX$ = REGLEIDO$("COTELEF1", RENOTA&)
                  If Left$(TTXX$, 12) = NUCOTE$ Then
                     GoTo 29
                  End If
               End If
            End If
         Next U&
         Exit Sub
         '
29       HOII% = HOY(HOS$)
         Call REPLA(TTXX$, MKI$(HOII%), 251, 2)
         Call GRAREG("COTELEF1", TTXX$, RENOTA&)
         Call CIERRARCH("COTELEF1")
         Call FRESHECHO("COTELEF1")
         Call FRESHECHO("COTELEF3")
         '
         FEX = HOII%
         Label25 = FECHATEX$(FEX)
         Image2.Visible = False
         Image1.Visible = True
         '
       End If
     End If
   End If
   '
End Sub

Private Sub Image3_DblClick()
   If USNBR% Mod 100 > 1 Then Exit Sub
   NTAR& = Val(Label28)
   If NTAR& > 0 Then
      Text31 = "30" + Mid$(Label29, 3)
      HOII% = HOY(HOS$)
      If FECHANUM(Text31) > HOII% Then
         Text31 = HOS$
      End If
      TOPLIS = List3.TopIndex
      Call Command29_Click
      On Error Resume Next
      List3.TopIndex = TOPLIS
      Call List3_DblClick
   End If
End Sub

Private Sub Label11_DblClick()
   If Label15 <> "" Then
     Call FILTERFILE("DETOTRA", "", 1, Label15)
     VTB% = VENTABU%("DETOTRA")
     If VTB% >= 0 Then
       Call FILTERUPDATE("DETOTRA", "", 1, Label15)
       Call SHOWDETHOR
     End If
   End If
End Sub

Private Sub Label15_Change()
   Call CARDATOTRA
   Call SHOWDETHOR
   'Call CARDETOTRA
End Sub

Sub SHOWDETHOR()
   List1.Clear
   Call FILTERFILE("DETOTRA", "", 1, Label15)
   SUMEPB = 0: SUMOTR = 0
   For U& = 1 To ULTREG&("!DETOTRA")
     XX$ = REGLEIDO$("!DETOTRA", U&)
     TTXX$ = Space$(128)
       FEX = CVI(Mid$(XX$, 13, 2))
     Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
       REFX$ = TRIM$(Mid$(XX$, 15, 50))
     Call REPLA(TTXX$, REFX$, 11, 50)
       COMI$ = HORATEX$(CVI(Mid$(XX$, 65, 2)))
     Call REPLA(TTXX$, COMI$, 63, 5)
       FINI$ = HORATEX$(CVI(Mid$(XX$, 67, 2)))
     Call REPLA(TTXX$, FINI$, 70, 5)
       HORAX$ = FORMATNUM$(CVS(Mid$(XX$, 69, 4)), "F6.2")
     Call REPLA(TTXX$, HORAX$, 75, 6)
       RESPO$ = Mid$(XX$, 73, 4)
     Call REPLA(TTXX$, RESPO$, 83, 4)
     If TRIM$(UCase$(RESPO$)) = "EPB" Or TRIM$(UCase$(RESPO$)) = "EB" Then
         SUMEPB = SUMEPB + Val(HORAX$)
       Else
         SUMOTR = SUMOTR + Val(HORAX$)
     End If
     List1.AddItem TTXX$
   Next U&
   '
   Text18 = FORMATNUM$(SUMEPB, "F6.2")
   Text17 = FORMATNUM$(SUMOTR, "F6.2")
   Text19 = FORMATNUM$(SUMEPB + SUMOTR, "F6.2")
   
End Sub

Private Sub Label25_DblClick()
   Label25 = ""
   Image2.Visible = True
   Image1.Visible = False
End Sub

Private Sub Label27_DblClick()
   NTAR& = Val(Label28)
   If NTAR& > 0 Then
      HOYX = HOY(HOS$)
      FPOSTI% = DIPOST(HOYX)
      TTXX$ = FILTERGETRECORD$("TARINTER", 1, MKS$(NTAR&))
      Call REPLA(TTXX$, MKI$(FPOSTI%), 75, 2)
      Call FILTERPUTRECORD("TARINTER", 1, MKS$(NTAR&), TTXX$)
      Label34 = FECHATEX$(FPOSTI)
      Call CARLISTAPEN
      Call List3_DblClick
    End If
End Sub

Private Sub Label28_Change()
   Text29 = ""
   Text28 = ""
   Text13 = ""
   Text31 = ""
   Label34 = ""
   Option3(0).Value = True
   '
   NTAR& = Val(Label28)
   If NTAR& > 0 Then
     TTXX$ = FILTERGETRECORD$("TARINTER", 1, MKS$(NTAR&))
     FEX = CVI(Mid$(TTXX$, 65, 2))
     Label29 = FECHATEX$(FEX)
     FEX = CVI(Mid$(TTXX$, 67, 2))
     Text31 = FECHATEX$(FEX)
     FEPOST = CVI(Mid$(TTXX$, 75, 2))
     If FEPOST > HOY(HOS$) Then
        Label34 = FECHATEX$(FEPOST)
     End If
     Text29 = TRIM$(Mid$(TTXX$, 77, 52))
     Text28 = RTrim(Mid$(TTXX$, 129, 512))
     Text13 = RTrim(Mid$(TTXX$, 641, 384))
     IRESPO% = Asc(Mid$(TTXX$, 69, 1))
     On Error Resume Next
     Option3(IRESPO%).Value = True
     Option4(IRESPO%).Value = True
     On Error GoTo 0
   End If
   '
End Sub

Private Sub Label28_DblClick()
   Call Command28_Click
End Sub

Private Sub Label30_DblClick()
   NTAR& = Val(Label28)
   If NTAR& > 0 Then
      HOYX = HOY(HOS$)
      FPOST = HOYX
      While DIENTRE(HOYX, FPOST) < 7
         FPOST = DIPOST(FPOST)
      Wend
      FPOSTI% = FPOST
      TTXX$ = FILTERGETRECORD$("TARINTER", 1, MKS$(NTAR&))
      Call REPLA(TTXX$, MKI$(FPOSTI%), 75, 2)
      Call FILTERPUTRECORD("TARINTER", 1, MKS$(NTAR&), TTXX$)
      Label34 = FECHATEX$(FPOSTI)
      Call CARLISTAPEN
      Call List3_DblClick
    End If
End Sub

Private Sub Label33_DblClick()
   NTAR& = Val(Label28)
   If NTAR& > 0 Then
      HOYX = HOY(HOS$)
      FPOST = HOYX
      While DIENTRE(HOYX, FPOST) < 30
         FPOST = DIPOST(FPOST)
      Wend
      FPOSTI% = FPOST
      TTXX$ = FILTERGETRECORD$("TARINTER", 1, MKS$(NTAR&))
      Call REPLA(TTXX$, MKI$(FPOSTI%), 75, 2)
      Call FILTERPUTRECORD("TARINTER", 1, MKS$(NTAR&), TTXX$)
      Label34 = FECHATEX$(FPOSTI)
      Call CARLISTAPEN
      Call List3_DblClick
    End If
End Sub

Private Sub Label34_DblClick()
   NTAR& = Val(Label28)
   If NTAR& > 0 Then
      TTXX$ = FILTERGETRECORD$("TARINTER", 1, MKS$(NTAR&))
      Call REPLA(TTXX$, MKI$(0), 75, 2)
      Call FILTERPUTRECORD("TARINTER", 1, MKS$(NTAR&), TTXX$)
      Label34 = ""
      Call CARLISTAPEN
    End If
End Sub

Private Sub Label7_Change()
   ADJUNARCH.Label3 = ""
End Sub

Private Sub Label8_DblClick()
   '
   Call SELECHO("DEUDOR1")
   NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCXI% > 0 Then
      Label24 = NCXI%
      Label8 = RASOCLI$(NCXI%)
   End If
   '
End Sub

Private Sub List1_DblClick()
   Call Label11_DblClick
End Sub

Private Sub List3_Click()
   If ENCARLISTA% <> 1 Then
      For KU& = List3.ListIndex To 0 Step -1
        NTAR1& = Val(Left$(List3.List(KU&), 5))
        If NTAR1& > 0 Then
          Label28 = TRIM$(Str$(NTAR1&))
          Exit Sub
        End If
      Next KU&
   End If
End Sub

Private Sub List3_DblClick()
   For KU& = List3.ListIndex To 0 Step -1
     NTAR1& = Val(Left$(List3.List(KU&), 5))
     If NTAR1& > 0 Then
       Label28 = TRIM$(Str$(NTAR1&))
       Exit Sub
     End If
   Next KU&
End Sub

Private Sub List3_KeyDown(KeyCode As Integer, Shift As Integer)
   '
   If KeyCode = 114 Then
     If Option4(9).Value = True Then
       If ABUSQUE$ <> "" Then
         If List3.ListIndex >= 0 Then
           For KU& = List3.ListIndex + 2 To List3.ListCount
             If InStr(UCase$(List3.List(KU& - 1)), UCase$(ABUSQUE$)) > 0 Then
                List3.ListIndex = KU& - 1
                Exit Sub
             End If
           Next KU&
           MsgBox "Argumento no Encontrado"
         End If
       End If
     End If
   End If
   '
End Sub

Private Sub Option3_Click(INDEX As Integer)
   ' Option4(INDEX).Value = True
End Sub

Private Sub Option4_Click(INDEX As Integer)
   Call CARLISTAPEN
   Call List3_Click
End Sub

Sub CARLISTAPEN()
   Dim RETEY$(1024)
   ENCARLISTA% = 1
   FILRESPO% = 0
   For KKI% = 1 To 8
     If Option4(KKI%).Value = True Then FILRESPO% = KKI%
   Next KKI%
   '
   TODOS% = 0
   If Option4(9).Value = True Then ' SELECCIONA TODOS
      TODOS% = 1
      RFF$ = ""
      URE& = ULTREG&("TARINTER")
      For KKU& = 1 To URE&
        RFF$ = RFF$ + MKS$(KKU&)
      Next KKU&
      GoTo 15
   End If
   
   RFF$ = RECFILT$("TARINTER", 5, Chr$(FILRESPO%))
   '
15 TP1 = List3.TopIndex
   TP2 = TP1
   For KU& = TP1 + 1 To TP1 + 20
      If KU& < List3.ListCount Then
         If List3.Selected(KU&) = True Then
            TP2 = KU&
            Exit For
         End If
      End If
   Next KU&
   '
   List3.Clear
   HOII% = HOY(HOS$)
   For KKU& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, KKU&, 4))
     XX$ = REGLEIDO$("TARINTER", REG&)
     NTAR& = CVS(Left$(XX$, 4))
     FEX = CVI(Mid$(XX$, 65, 2))
     FECIE% = CVI(Mid$(XX$, 67, 2))
     If FECIE% < 1 Then
        If CVI(Mid$(XX$, 75, 2)) <= HOII% Then ' PARA SUPRIMIR LAS POSTERGADAS
           DETAX$ = RTrim(Mid$(XX$, 129, 512))
           If TRIM$(DETAX$) = "" Then
             DETAX$ = RTrim(Mid$(XX$, 77, 52))
           End If
           Call FRATEXTO(DETAX$, 64)
           '
           CARETEY% = CARETEX%
           CARETEZ% = CARETEX%
           For KKI% = 1 To CARETEX%
              RETEY$(KKI%) = RETEX$(KKI%)
              RETEX$(KKI%) = ""
           Next KKI%
           '
           OBSEX$ = RTrim$(Mid$(XX$, 641, 384))
           Call FRATEXTO(OBSEX$, 64)
           If CARETEX% > CARETEZ% Then CARETEZ% = CARETEX%
           '
           YY$ = Space$(128 + 64)
           Call REPLA(YY$, FORMATNUM$(NTAR&, "F5.0"), 1, 5)
           Call REPLA(YY$, FECHATEX$(FEX), 9, 8)
           Call REPLA(YY$, RETEY$(1), 19, 64)
           Call REPLA(YY$, RETEX$(1), 97, 64)
           '
           If TODOS% > 0 Then
             IRESPO% = Asc(Mid$(XX$, 69, 1))
             IRESX$ = ""
             If IRESPO% < 9 Then
               IRESX$ = Option4(IRESPO%).Caption
             End If
             Call REPLA(YY$, IRESX$, 86, 10)
           End If
           '
           List3.AddItem YY$
           For KKI% = 2 To CARETEZ%
             YY$ = Space$(128 + 64)
             Call REPLA(YY$, RETEY$(KKI%), 19, 64)
             Call REPLA(YY$, RETEX$(KKI%), 97, 64)
             List3.AddItem YY$
           Next KKI%
           List3.AddItem String$(192, "-")
        End If
     End If
   Next KKU&
   '
   On Error Resume Next
   List3.TopIndex = TP1
   List3.Selected(TP2) = True
   On Error GoTo 0
   ENCARLISTA% = 0
   '
End Sub
'
Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab <> PreviousTab Then
     If SSTab1.Tab > 1 Then
       If USNBR% <> 101 Then
           SSTab1.Tab = PreviousTab
         Else
           Call Option4_Click(0)
       End If
     End If
   End If
End Sub

Private Sub Text1_Change()
    '
    Combo1.Clear
    NC% = Val(Text1)
    IDLAN$ = HORARCLI$(NC%)
    RFF$ = RECFILT$("WEBUSER", 4, IDLAN$)
    For KKU& = 1 To Len(RFF$) Step 4
       REGI& = CVS(Mid$(RFF$, KKU&, 4))
       XX$ = REGLEIDO$("WEBUSER", REGI&)
       TTXX$ = AJUSTI$(Mid$(XX$, 3, 30), 64) + TRIM$(Mid$(XX$, 33, 12))
       Combo1.AddItem TTXX$
    Next KKU&
    On Error Resume Next
    Combo1.ListIndex = 0
    Combo1.Refresh
    On Error GoTo 0
    '
End Sub

Private Sub Text1_DblClick()
    '
    Call Command1_Click
    '
End Sub


Sub CARDETCOTE()
   '
   NUCOTE$ = AJUSTI$(Label7.Caption, 12)
   CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                  CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                Next UI%
                '
   Screen.MousePointer = 11
   URN& = ULTREG&("COTELEF2")
   Text22.TEXT = "": Text23.TEXT = ""
   '
   Screen.MousePointer = 11
   RNC$ = RECFILT$("COTELEF2", 2, MKI$(CICOTE%))
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("COTELEF2", RENOTA&)
         If Left$(TTXX$, 12) = NUCOTE$ Then
           TIRE$ = Mid$(TTXX$, 15, 2)
           TTYY$ = Mid$(TTXX$, 17, 72)
           If TIRE$ = "CO" Then
                If Text22.TEXT <> "" Then
                  Text22.TEXT = Text22.TEXT + Chr$(13) + Chr$(10)
                End If
                Text22.TEXT = Text22.TEXT + RTrim$(Left$(TTYY$, 72))
              ElseIf TIRE$ = "DI" Then
                If Text23.TEXT <> "" Then
                  Text23.TEXT = Text23.TEXT + Chr$(13) + Chr$(10)
                End If
                Text23.TEXT = Text23.TEXT + RTrim$(Left$(TTYY$, 72))
           End If
         End If
       End If
     End If
   Next U&
   Screen.MousePointer = 1
   Text22.Refresh: Text23.Refresh
   '
End Sub

Sub GRADETCOTE()
   '
   NUCOTE$ = AJUSTI$(Label7.Caption, 12)
   CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                  CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                Next UI%
                '
   Screen.MousePointer = 11
   URN& = ULTREG&("COTELEF2")
   RNX$ = RECFILT$("COTELEF2", 2, MKI$(CICOTE%))
   RNC$ = ""
   For U& = 1 To Len(RNX$) Step 4
      RENOTA& = CVS(Mid$(RNX$, U&, 4))
      TTXX$ = REGLEIDO$("COTELEF2", RENOTA&)
      If Left$(TTXX$, 12) = NUCOTE$ Then
         RNC$ = RNC$ + MKS$(RENOTA&)
      End If
   Next U&
   KK& = (-3)
   '
   TEXTO$ = Text22.TEXT + Chr$(13) + Chr$(10)
   Call FRATEXTO(TEXTO$, 72)
   For LL% = 1 To CARETEX%
      TTYY$ = NUCOTE$ + MKI$(CICOTE%) + "CO" + RETEX$(LL%)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      TTXX$ = REGLEIDO$("COTELEF2", RENOTA&)
      Call GRAREG("COTELEF2", TTYY$, RENOTA&)
   Next LL%
   '
   TEXTO$ = Text23.TEXT + Chr$(13) + Chr$(10)
   Call FRATEXTO(TEXTO$, 71)
   For LL% = 1 To CARETEX%
      TTYY$ = NUCOTE$ + MKI$(CICOTE%) + "DI" + RETEX$(LL%)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      TTXX$ = REGLEIDO$("COTELEF2", RENOTA&)
      Call GRAREG("COTELEF2", TTYY$, RENOTA&)
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("COTELEF2", NUCOTE$ + MKI$(CICOTE%) + "DI" + Space$(72), RENOTA&)
   Wend
   '
   Call CIERRARCH("COTELEF2")
   Screen.MousePointer = 1
   '
End Sub
'

Sub CARDATCOTE()
   '
   NUCOTE$ = AJUSTI$(Label7.Caption, 12)
   CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                  CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                Next UI%
                '
   Screen.MousePointer = 11
   RNC$ = RECFILT$("COTELEF1", 3, MKI$(CICOTE%))
   URN& = ULTREG&("COTELEF1")
   '
   For U& = 1 To Len(RNC$) Step 4
      RENOTA& = CVS(Mid$(RNC$, U&, 4))
      If RENOTA& > 0 Then
         If RENOTA& <= URN& Then
            TTXX$ = REGLEIDO$("COTELEF1", RENOTA&)
            If Left$(TTXX$, 12) = NUCOTE$ Then
               GoTo 29
            End If
         End If
      End If
   Next U&
   TTXX$ = REGBLAN$("COTELEF1")
   '
29 Text21.TEXT = TRIM$(Mid$(TTXX$, 111, 64))
   NC% = CVI(Mid$(TTXX$, 81, 2))
   If NC% > 0 Then
      Text1.TEXT = TRIM$(Str$(NC%))
      Text2.TEXT = RASOCLI$(NC%)
      Text3.TEXT = TRIM$(DOMICLI$(NC%)) + " - (" + TRIM$(CPOSCLI$(NC%)) + ") " + LOCACLI$(NC%)
      Text4.TEXT = TEL1CLI$(NC%)
      Text5.TEXT = FAXCLI$(NC%)
   End If
   'Text11.TEXT = TRIM$(Mid$(TTXX$, 83, 24))
   Text12.TEXT = FECHATEX$(CVI(Mid$(TTXX$, 107, 2)))
   Text6.TEXT = HORATEX$(CVI(Mid$(TTXX$, 109, 2)))
   '
   WEBUS$ = TRIM$(Mid$(TTXX$, 83, 24))
   For KKI% = 1 To Combo1.ListCount
      If UCase$(TRIM$(Left$(Combo1.List(KKI% - 1), 24))) = UCase$(WEBUS$) Then
         Combo1.ListIndex = KKI% - 1
         Combo1.Refresh
         GoTo 34
      End If
   Next KKI%
   Combo1.AddItem AJUSTI$(WEBUS$, 64)
   Combo1.ListIndex = Combo1.ListCount - 1
   '
34 AAA$ = Mid$(TTXX$, 179, 72)
   Text10 = ""
   CONCO$ = TRIM$(Mid$(TTXX$, 179, 24))
   For KKI% = 1 To Combo1.ListCount
      If UCase$(TRIM$(Left$(Combo1.List(KKI% - 1), 24))) = UCase$(CONCO$) Then
         Text10.TEXT = Combo1.List(KKI% - 1)
      End If
   Next KKI%
   If Text10 = "" Then
     If Combo1.ListIndex > 0 Then
       Text10 = Combo1.List(0)
     End If
   End If
   '
   Label25 = ""
   Image2.Visible = True
   Image1.Visible = False
   If CVI(Mid$(TTXX$, 251, 2)) > 8224 Then
     FEX = CVI(Mid$(TTXX$, 251, 2))
     Label25 = FECHATEX$(FEX)
     Image2.Visible = False
     Image1.Visible = True
   End If
   '
   TIPOCO% = Asc(Mid$(TTXX$, 253, 1))
   If TIPOCO% > 3 Then TIPOCO% = 0
   Option5(TIPOCO%).Value = True
   '
   Screen.MousePointer = 1
   '
End Sub

Sub GRADATCOTE()
   '
   NUCOTE$ = AJUSTI$(Label7.Caption, 12)
   CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                  CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                Next UI%
                '
   Screen.MousePointer = 11
   NC% = Val(Text1.TEXT)
   REFE$ = TRIM$(TEL2CLI$(NC%)) + " - " + TRIM$(Text12.TEXT) + " - " + Text21.TEXT
   TTYY$ = REGBLAN$("COTELEF1")
   Call REPLA(TTYY$, NUCOTE$, 1, 12)
   Call REPLA(TTYY$, REFE$, 13, 66)
   Call REPLA(TTYY$, MKI$(CICOTE%), 79, 2)
   Call REPLA(TTYY$, MKI$(Val(Text1.TEXT)), 81, 2)
   Call REPLA(TTYY$, Combo1.TEXT, 83, 24)
   Call REPLA(TTYY$, MKI$(Int(FECHANUM(Text12.TEXT))), 107, 2)
   Call REPLA(TTYY$, MKI$(Int(HORANUM(Text6.TEXT))), 109, 2)
   Call REPLA(TTYY$, Text21.TEXT, 111, 64)
   RETEX$(1) = "": RETEX$(2) = "": RETEX$(3) = ""
   Call FRATEXTO(Text10.TEXT, 24)
   Call REPLA(TTYY$, AJUSTI$(TRIM$(RETEX$(1)), 24), 179, 24)
     FEXI% = FECHANUM(Label25)
   Call REPLA(TTYY$, MKI$(FEXI%), 251, 2)
   '
   TIPOCO% = 0
   If Option5(1).Value = True Then TIPOCO% = 1
   If Option5(2).Value = True Then TIPOCO% = 2
   If Option5(3).Value = True Then TIPOCO% = 3
   Call REPLA(TTYY$, Chr$(TIPOCO%), 253, 1)
   '
   RNC$ = RECFILT$("COTELEF1", 3, MKI$(CICOTE%))
   URN& = ULTREG&("COTELEF1")
   '
   For U& = 1 To Len(RNC$) Step 4
      RENOTA& = CVS(Mid$(RNC$, U&, 4))
      If RENOTA& > 0 Then
         If RENOTA& <= URN& Then
            TTXX$ = REGLEIDO$("COTELEF1", RENOTA&)
            If Left$(TTXX$, 12) = NUCOTE$ Then
               GoTo 29
            End If
         End If
      End If
   Next U&
   RENOTA& = 1 + ULTREG&("COTELEF1")
   '
29 Call GRAREG("COTELEF1", TTYY$, RENOTA&)
   '
   Call COPYSTRU("COTELEF1", "COTELEF3")
   If FEXI% > 8224 Then
       Call FILTERDELETE("COTELEF3", 1, NUCOTE$)
     Else
       RFF$ = RECFILT$("COTELEF3", 1, NUCOTE$)
       RENOTA3& = CVS(RFF$)
       If RENOTA3& > 0 Then
           Call GRAREG("COTELEF3", TTYY$, RENOTA3&)
         Else
           Call REGAPP("COTELEF3", TTYY$)
           Call CIERRARCH("COTELEF3")
           Call FILESORT("COTELEF3", "", 6, 7, "ASC")
       End If
   End If
'   JJ& = 0
'   Screen.MousePointer = 11
'   For U& = 1 To ULTREG("COTELEF1")
'     TTXX$ = REGLEIDO$("COTELEF1", U&)
'     If CVI(Mid$(TTXX$, 251, 2)) <= 8224 Then
'       JJ& = JJ& + 1
'       Call GRAREG("COTELEF3", TTXX$, JJ&)
'     End If
'   Next U&
   Call CIERRARCH("COTELEF1")
   Call FRESHECHO("COTELEF1")
   '
   'Call SETULTREG("COTELEF3", JJ&)
   'Call FRESHECHO("COTELEF3")
   '
   Screen.MousePointer = 1
   '
End Sub





Private Sub Text13_DblClick()
   Text13 = ULOCASE$(Text13)
End Sub

Private Sub Text14_DblClick()
   Text14 = ULOCASE$(Text14)
End Sub

Private Sub Text15_DblClick()
   Text15 = ULOCASE$(Text15)
End Sub

Private Sub Text16_DblClick()
   Text16 = ULOCASE$(Text16)
End Sub



Private Sub Text20_DblClick()
   If TRIM$(Text20) = "" Then
      HOII% = HOY(HOS$)
      Text20 = HOS$
      If TRIM$(Text27) = "" Then Text27 = HOS$
      If TRIM$(Text26) = "" Then Text26 = HOS$
      Command9.SetFocus
      Exit Sub
   End If
   '
   Call Command16_Click
End Sub

Private Sub Text20_GotFocus()
   If USNBR% <> 101 Then
     Text27.SetFocus
   End If
End Sub

Private Sub Text20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If USNBR% <> 101 Then
     Text27.SetFocus
   End If
End Sub

Private Sub Text21_DblClick()
   Text21 = ULOCASE$(Text21)
End Sub

Private Sub Text22_Change()
   LOMAX% = 72
   If Len(Text22.TEXT) > LOMAX% Then
      For KK% = Len(Text22.TEXT) To 1 Step -1
         If Mid$(Text22.TEXT, KK%, 2) = Chr$(13) + Chr$(10) Then
            PPXX% = KK%
            GoTo 10
         End If
      Next KK%
      PPXX% = 0
10    If Len(Text22.TEXT) - PPXX% > LOMAX% Then
         Call FRATEXTO(TRIM$(Text22.TEXT), LOMAX% - 1)
         TTXX$ = ""
         If CARETEX% > 0 Then
         TTXX$ = RETEX$(1)
         For KK% = 2 To CARETEX%
            TTXX$ = TTXX$ + Chr$(13) + Chr$(10) + RTrim(RETEX$(KK%))
         Next KK%
         End If
         Text22.TEXT = RTrim(TTXX$)
         Text22.SelStart = 1 + Len(Text22.TEXT)
      End If
   End If
End Sub

Sub ENVIAPORMAIL()
   '
   Command1.Enabled = False
   '
   IMPREDET$ = Printer.DeviceName
   Printer.TrackDefault = False
   For Each PRX1 In Printers
     If InStr(UCase$(PRX1.DeviceName), "PDFWRITER") > 0 Then
       On Error Resume Next
       Set Printer = PRX1
       If Err Then
         On Error GoTo 0
         GoTo 5
       End If
       On Error GoTo 0
       GoTo 10
     End If
   Next
   '
5  Call MENSERR(24, "Imposible Asignar a PDFWriter")
   GoTo 99
   '
10 NUCETE$ = Label7.Caption
   On Error Resume Next
   Kill "C:\FLEXOFT\" + TRIM$(Label7.Caption) + ".PDF"
   On Error GoTo 0
   Timer1.Enabled = True
   Call PRICOTEL(TRIM$(Label7.Caption))
   '
   NC% = XVALO(Text1)
   IDLAN$ = HORARCLI$(NC%)
   DSTX$ = TRIM$(Mid$(Combo1.TEXT, 65, 12))
   CONCO$ = TRIM$(Mid$(Text10.TEXT, 65, 12))
   '
   LISTARCH$ = "C:\FLEXOFT\" + TRIM$(Label7.Caption) + ".PDF"
   For KKLI% = 1 To ADJUNARCH.List1.ListCount
      LISTARCH$ = LISTARCH$ + "; " + ADJUNARCH.List1.List(KKLI% - 1)
   Next KKLI%
   '
   Call MANDALAWEB(LISTARCH$, IDLAN$ + ";" + DSTX$ + "+" + CONCO$)
   '
   HOII% = HOY(HOS$)
   ASUNTO$ = "FLEXOFT Soluciones ("
   CLAX$ = RANDSTRING$(-6)
   ASUNTO$ = ASUNTO$ + CLAX$ + "/" + CONTRACLA$(CLAX$) + ")"
   '
   TEXTOMAI$ = "Estamos enviando adjunto al presente respuesta a su" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "Consulta Telefonica Número " + Label7.Caption + ", en archivo" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + TRIM$(Label7.Caption) + ".PDF." + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "La consulta puede ser visualizada mediante ACROBAT Reader." + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "En caso de no estar instalado este utilitario, puede descar-" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "garlo de www.adobe.es/products/acrobat/readstep2.html" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
   If ADJUNARCH.Label3 <> "" Then
     If EXISTE%(ADJUNARCH.Label3) > 0 Then
       TEXTOMAI$ = TEXTOMAI$ + "Asimismo se adjunta ZIP-File conteniendo los siguientes ar-" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "chivos correspondientes a actualización de versión:" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
       For U& = 1 To ADJUNARCH.List1.ListCount
         FIFUNA$ = ReplaSkey$(ADJUNARCH.List1.List(U& - 1))
         FINA$ = FIFUNA$
15       PPXX% = InStr(FINA$, "\")
         If PPXX% > 0 Then
           FINA$ = Mid$(FINA$, PPXX% + 1)
           GoTo 15
         End If
         DATIM$ = ""
         On Error Resume Next
         DATIM$ = Format$(FileDateTime(FIFUNA$), "dd/mm/yy - hh:mm")
         If DATIM$ <> "" Then
           TEXTOMAI$ = TEXTOMAI$ + AJUSTI$(FINA$, 32) + " del " + DATIM$ + CRLF$
         End If
       Next U&
       TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "Para instalar las actualizaciones, proceda así:" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "1) Cerrar las aplicaciones a instalar en todas las estaciones de" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "   trabajo" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "2) Extraer el contenido del ZIP en la carpeta \FLEXOFT\NOVEDAD" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "   del disco servidor " + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "3) Cerrar el FLEXOFT y volver a ingresar. Luego de introducir el" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "   usuario y la contraseña, aparecerá un mensaje de actualizaciones en" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "   el cual deberá cliquear sobre el botón 'Si'." + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "4) La actualización finaliza con el mensaje 'Actualizaciones" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "   Instaladas '." + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "En caso de que no se auto-instalen los controladores tipo '*.RFS'" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "'*.frm', pueden copiarse manualmente desde la carpeta \FLEXOFT\NOVEDAD" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "hacia la(s) carpeta(s) de datos existentes." + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "En caso de dudas, solicite asistencia telefónica a nuestra mesa de" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "ayuda: (011)-4524-1284 / 4541-1515" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "Una vez verificado el correcto funcionamiento de las actualizaciones" + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "enviadas, le pedimos confirmar respondiendo el presente al remitente." + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "Esto nos ayudará a mejorar nuestro servicio. Desde ya, muchas gracias." + CRLF$
       '
     End If
   End If
   '
   'TEXTOMAI$ = TEXTOMAI$ + "Para Auto-Presentar, descomprima FLX_MSGR.EXE del UPFLEX00.ZIP" + CRLF$
   'TEXTOMAI$ = TEXTOMAI$ + "adjunto y cópielo a su carpeta \FLEXOFT\WINDOWS." + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "Quedamos a disposición para cualquier consulta adicional, aprovechando" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "para saludarlos" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "                       Cordialmente" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "===================================================================" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "FLEXOFT Soluciones Informáticas Integrales" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "Bauness 2170 - 5to. 'D' - 1431 Buenos Aires - Argentina" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "Tel (00-54-11) 4524-1284 / 4541-1515 - e-mail: info@flexoft.com.ar" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "===================================================================" + CRLF$
   '
   Call SAVEFILE("C:\FLEXOFT\PW" + CLAX$ + ".PWY", TEXTOSEND$)
   '
   MAI_DESTI$ = TRIM$(Text2)
     NCX% = Val(Text1)
     DESTIMA$ = TRIM$(EMAILCLI$(NCX%))
     If DESTIMA$ = "" Or InStr(DESTIMA$, "@") < 1 Then
       Call MENSERR(24, "Imposible Enviar.\Dirección de Mail Inexistente o no Válida.")
       GoTo 90
     End If
   MAI_DIREL$ = DESTIMA$     ' "fgrasso@flexoft.com.ar"
   MAI_ASUNT$ = ASUNTO$
   MAI_TEXTO$ = TEXTOMAI$
   MAI_ADJUN$(0) = "C:\FLEXOFT\" + TRIM$(Label7.Caption) + ".PDF"
   MAI_ADJUN$(1) = ""
   If ADJUNARCH.Label3 <> "" Then
     If EXISTE%(ADJUNARCH.Label3) > 0 Then
       MAI_ADJUN(1) = ADJUNARCH.Label3
     End If
   End If
   '
   Call GENERAMAIL
   ATENTI% = 1
   '
   ' para enviar aviso a la direccion de contacto
   DESTIMB$ = TRIM$(CONTACTCLI$(NC%))
   If DESTIMB$ <> "" Then
      If InStr(DESTIMB$, "@") > 1 Then
        ASUNTO$ = "FLEXOFT Soluciones ("
        CLAX$ = RANDSTRING$(-6)
        ASUNTO$ = ASUNTO$ + CLAX$ + "/" + CONTRACLA$(CLAX$) + ")"
        '
        TEXTOMAI$ = "Informamos que desde nuestro Servicio de Soporte de Usuarios se ha enviado " + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "respuesta a su Consulta Telefonica Número " + Label7.Caption + "." + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Teniendo en cuenta las características de configuración de esta vues-" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "tra cuenta de correo, la que bloquea la recepción de mensajes con" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "archivos adjuntos, el referido envío se ha remitido a la direccion" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + DESTIMA$ + " de donde pueden Uds. bajarlo." + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
        '
        TEXTOMAI$ = TEXTOMAI$ + "En caso de dudas, soliciten asistencia telefónica a nuestra mesa de" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "ayuda: (011)-4524-1284 / 4522-2048 - interno 32" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Una vez verificado el correcto funcionamiento de las actualizaciones" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "enviadas, les pedimos confirmar respondiendo el presente al remitente." + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Esto nos ayudará a mejorar nuestro servicio. Desde ya, muchas gracias." + CRLF$
        '
        TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Quedamos a disposición para cualquier consulta adicional, aprovechando" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "para saludarlos" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "                       Cordialmente" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "=====================================================================" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "FLEXOFT Soluciones Informáticas Integrales" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Miembro de la Cámara de Empresas de Software y Servicios Informáticos" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Bauness 2170 - 5to. 'D' - 1431 Buenos Aires - Argentina" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "Tel (00-54-11) 4524-1284 / 4522-2048 - e-mail: info@flexoft.com.ar" + CRLF$
        TEXTOMAI$ = TEXTOMAI$ + "=====================================================================" + CRLF$
        '
        MAI_DIREL$ = DESTIMB$     ' "fgrasso@flexoft.com.ar"
        MAI_ASUNT$ = ASUNTO$
        MAI_TEXTO$ = TEXTOMAI$
        MAI_ADJUN$(0) = ""
        MAI_ADJUN$(1) = ""
        '
        Call GENERAMAIL
        '
     End If
   End If
   '
90 For Each PRX1 In Printers
     If PRX1.DeviceName = IMPREDET$ Then
       On Error Resume Next
       Set Printer = PRX1
       On Error GoTo 0
     End If
   Next
   '
99 Command1.Enabled = True
   '
End Sub

Private Sub Text22_DblClick()
  Text22 = ULOCASE$(Text22)
'  TTXX$ = TRIM$(Text22)
'  For U& = 2 To Len(TTXX$)
'    If Asc(Mid$(TTXX$, U& - 1, 1)) >= 65 Then
'      Mid$(TTXX$, U&, 1) = LCase$(Mid$(TTXX$, U&, 1))
'    End If
'  Next U&
'  Text22 = TTXX$
End Sub

Private Sub Text23_DblClick()
   Text23 = ULOCASE$(Text23)
End Sub

Private Sub Text26_DblClick()
   Call Command13_Click
End Sub

Private Sub Text27_DblClick()
   Call Command15_Click
End Sub

Private Sub Text28_DblClick()
   Text28 = ULOCASE$(Text28)
End Sub

Private Sub Text28_GotFocus()
   If USNBR% Mod 100 > 1 Then
      Text13.SetFocus
      Exit Sub
   End If
   If Text28 = "" Then
      If Text29 = "" Then
         If Label28 < PROTARIN& Then
            Command30.SetFocus
         End If
      End If
   End If
End Sub

Private Sub Text29_GotFocus()
   If USNBR% Mod 100 > 1 Then Text13.SetFocus
End Sub

Private Sub Timer1_Timer()
   If NUCETE$ = "" Then
     Exit Sub
   End If
   '
   On Error Resume Next
   AppActivate "Save PDF"
   If Err Then
     On Error GoTo 0
     Exit Sub
   End If
   '
   SendKeys "%N" + "C:\FLEXOFT\" + NUCETE$ + ".PDF" + Chr$(13)
   NUCETE$ = ""
   Timer1.Enabled = False
   Exit Sub
   '
End Sub

Sub CARDATOTRA()
   '
   NUCOTE$ = AJUSTI$(Label15.Caption, 12)
   CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                  CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                Next UI%
                '
   Screen.MousePointer = 11
   RNC$ = RECFILT$("CORTRAB1", 3, MKI$(CICOTE%))
   URN& = ULTREG&("CORTRAB1")
   '
   For U& = 1 To Len(RNC$) Step 4
      RENOTA& = CVS(Mid$(RNC$, U&, 4))
      If RENOTA& > 0 Then
         If RENOTA& <= URN& Then
            TTXX$ = REGLEIDO$("CORTRAB1", RENOTA&)
            If Left$(TTXX$, 12) = NUCOTE$ Then
               GoTo 29
            End If
         End If
      End If
   Next U&
   TTXX$ = REGBLAN$("CORTRAB1")
   '
29 Text9.TEXT = TRIM$(Mid$(TTXX$, 129, 64))
   NC% = CVI(Mid$(TTXX$, 81, 2))
   If NC% > 0 Then
      Label24 = TRIM$(Str$(NC%))
      Label8 = RASOCLI$(NC%)
   End If
   Text8.TEXT = TRIM$(Mid$(TTXX$, 83, 24))
   Label14 = FECHATEX$(CVI(Mid$(TTXX$, 107, 2)))
     On Error Resume Next
     MOTI% = Asc(Mid$(TTXX$, 109, 1))
   Option1(MOTI%).Value = True
     On Error GoTo 0
     On Error Resume Next
     IMPU% = Asc(Mid$(TTXX$, 110, 1))
   Option2(IMPU%).Value = True
     On Error GoTo 0
   Text26 = "": If CVI(Mid$(TTXX$, 111, 2)) > 8224 Then Text26 = FECHATEX$(CVI(Mid$(TTXX$, 111, 2)))
   Text27 = "": If CVI(Mid$(TTXX$, 113, 2)) > 8224 Then Text27 = FECHATEX$(CVI(Mid$(TTXX$, 113, 2)))
   Text20 = "": If CVI(Mid$(TTXX$, 115, 2)) > 8224 Then Text20 = FECHATEX$(CVI(Mid$(TTXX$, 115, 2)))
   '
   Text16 = RTrim(Mid$(TTXX$, 193, 128))
   Text14 = RTrim(Mid$(TTXX$, 321, 256))
     DESOT$ = Mid$(TTXX$, 577, 448)
     RFF$ = RECFILT$("DESOTRA1", 1, NUCOTE$)
     For KU& = 1 To Len(RFF$) Step 4
       U& = CVS(Mid$(RFF$, KU&, 4))
       DESOT$ = DESOT$ + Mid$(REGLEIDO$("DESOTRA1", U&), 13, 116)
     Next KU&
     DESOT$ = RTrim(DESOT$)
   Text15 = DESOT$
   '
   Screen.MousePointer = 1
   '
End Sub

Sub GRADATOTRA()
   '
   NUCOTE$ = AJUSTI$(Label15.Caption, 12)
   CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                  CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                Next UI%
                '
   Screen.MousePointer = 11
   NC% = Val(Label24)
   REFE$ = TRIM$(TEL2CLI$(NC%)) + " - " + TRIM$(Text9.TEXT)
   If FECHANUM(Text20) > 0 Then
        REFE$ = "(c) " + REFE$
     ElseIf FECHANUM(Text27) > 0 Then
        REFE$ = "(i) " + REFE$
     ElseIf FECHANUM(Text26) > 0 Then
        REFE$ = "(e) " + REFE$
   End If
   '
   TTYY$ = REGBLAN$("CORTRAB1")
   Call REPLA(TTYY$, NUCOTE$, 1, 12)
   Call REPLA(TTYY$, REFE$, 13, 66)
   Call REPLA(TTYY$, MKI$(CICOTE%), 79, 2)
   Call REPLA(TTYY$, MKI$(Val(Label24)), 81, 2)   ' NUMERO DE CLIENTE
   Call REPLA(TTYY$, Text8.TEXT, 83, 24)          ' PERSONA DESTINATARIO
   Call REPLA(TTYY$, MKI$(Int(FECHANUM(Label14))), 107, 2)
   Call REPLA(TTYY$, Text9, 129, 64)
     MOTI% = 1
     On Error Resume Next
     For KKI% = 1 To 9
       If KKI% = 4 Then KKI% = 9
       If Option1(KKI%).Value = True Then MOTI% = KKI%
     Next KKI%
   Call REPLA(TTYY$, Chr$(MOTI%), 109, 1)
     IMPU% = 1
     On Error Resume Next
     For KKI% = 1 To 9
       If KKI% = 4 Then KKI% = 9
       If Option2(KKI%).Value = True Then IMPU% = KKI%
     Next KKI%
   Call REPLA(TTYY$, Chr$(IMPU%), 110, 1)
   '
   Call REPLA(TTYY$, MKI$(Int(FECHANUM(Text26))), 111, 2)
   Call REPLA(TTYY$, MKI$(Int(FECHANUM(Text27))), 113, 2)
   Call REPLA(TTYY$, MKI$(Int(FECHANUM(Text20))), 115, 2)
   '
   Call REPLA(TTYY$, Text16, 193, 128)
   Call REPLA(TTYY$, Text14, 321, 256)
   Call REPLA(TTYY$, Text15, 577, 448)
   '
   RNC$ = RECFILT$("CORTRAB1", 3, MKI$(CICOTE%))
   URN& = ULTREG&("CORTRAB1")
   '
   For U& = 1 To Len(RNC$) Step 4
      RENOTA& = CVS(Mid$(RNC$, U&, 4))
      If RENOTA& > 0 Then
         If RENOTA& <= URN& Then
            TTXX$ = REGLEIDO$("CORTRAB1", RENOTA&)
            If Left$(TTXX$, 12) = NUCOTE$ Then
               GoTo 29
            End If
         End If
      End If
   Next U&
   RENOTA& = 1 + ULTREG&("CORTRAB1")
   '
29 Call GRAREG("CORTRAB1", TTYY$, RENOTA&)
   '
   Call CIERRARCH("CORTRAB1")
   Call FRESHECHO("CORTRAB1")
   '
   TTYY$ = Mid$(Text15, 449)
   If Len(TTYY$) > 0 Then
     Call FILTERDELETE("DESOTRA1", 1, NUCOTE$)
     While Len(TTYY$) > 0
       TTZZ$ = Space$(128)
       Call REPLA(TTZZ$, NUCOTE$, 1, 12)
       Call REPLA(TTZZ$, Left$(TTYY$, 116), 13, 116)
       Call REGAPP("DESOTRA1", TTZZ$)
       TTYY$ = Mid$(TTYY$, 117)
     Wend
   End If
   Call CIERRARCH("DESOTRA1")
   '
   Screen.MousePointer = 1
   '
End Sub

Private Sub Timer2_Timer()
   Static SVLC$
   If SVLC$ = "" Then
      SVLC$ = SERIVOL$("C:")
   End If
   '
   If SVLC$ = "60D9-89EE" Then
      Call VERIFICALAWEB
   End If
End Sub

Private Sub Timer3_Timer()
   On Error Resume Next
   XXXX = Shell("FLX_BRWS.EXE", vbMinimizedFocus)
   YYYY = Shell("BABANEN.EXE", vbMinimizedFocus)
   On Error GoTo 0
End Sub

Private Sub Timer4_Timer()
    Static CAMINU%
    If CAMINU% Mod 10 = 0 Then
       CAMINU% = 0
       If SSTab1.Tab = 2 Then Exit Sub
       Call TARPENDE.CARLISTAPEN
       If TARPENDE.Visible <> True Then
            TARPENDE.Show
          Else
            TARPENDE.WindowState = 0
            TARPENDE.ZOrder 0
       End If
    End If
    CAMINU% = CAMINU% + 1
End Sub

Private Sub VScroll1_Change()
   Static VALANT%
   If VALANT% = 0 Then VALANT% = VScroll1.Value
   If VScroll1.Value > VALANT% Then
       RFF$ = RECFILT$("COTELEF1", 1, Label7)
       If Len(RFF$) >= 4 Then
          REGI& = CVS(Left$(RFF$, 4))
          If REGI& > 1 Then   '< ULTREG&("COTELEF1") Then
             REGI& = REGI& - 1
             XX$ = REGLEIDO$("COTELEF1", REGI&)
             NUCOTE$ = Left$(XX$, 12)
             Label7.Caption = NUCOTE$
             Call CARDATCOTE
             Call CARDETCOTE
             VALANT% = VALANT% + 1
          End If
       End If
      ElseIf VScroll1.Value < VALANT% Then
       RFF$ = RECFILT$("COTELEF1", 1, Label7)
       If Len(RFF$) >= 4 Then
          REGI& = CVS(Left$(RFF$, 4))
          If REGI& < ULTREG&("COTELEF1") Then
             REGI& = REGI& + 1
             XX$ = REGLEIDO$("COTELEF1", REGI&)
             NUCOTE$ = Left$(XX$, 12)
             Label7.Caption = NUCOTE$
             Call CARDATCOTE
             Call CARDETCOTE
             VALANT% = VALANT% - 1
          End If
       End If
    End If
End Sub

Sub VERIFICALAWEB()
   ' DETERMINA LA CARPETA PARA EL DOWNLOAD Y LA CREA
   CARPEDOWN$ = App.Path
   While Len(CARPEDOWN$) > 0 And Right$(CARPEDOWN$, 1) <> "\"
      CARPEDOWN$ = Left$(CARPEDOWN$, Len(CARPEDOWN$) - 1)
   Wend
   CARPEDOWN$ = CARPEDOWN$ + "DOWNLOAD\"
   CARPINS$ = Left$(CARPEDOWN$, Len(CARPEDOWN$) - 1)
   On Error Resume Next
   MkDir CARPINS$
   On Error GoTo 0
   '
   Call BORRACACHE
   JJ1& = URLDownloadToFile(0, "http://www.flexoft.com.ar/UPDATES/NOVETEST/IUPDATES.TXT", CARPEDOWN$ + "IUPDATES.TXT", 0, 0)
   '
   FULTACT$ = LOADFILE$("F:\WEBUPDATES\NOVEDAD\IUPDATES.TXT") ' LO LEE DEL DIRECTORIO DE TRABAJO
   FIDOWNL$ = LOADFILE$(CARPEDOWN$ + "IUPDATES.TXT")
   If FIDOWNL$ = FULTACT$ Then
      ' EL ARCHIVO BAJADO COINCIDE CON EL ANTERIOR - NO HAY NOVEDADES
      GoTo 99
   End If
   '
   TTZZ$ = ""
   For U& = 1 To Len(FULTACT$) Step 128
      AAA = Mid$(FULTACT$, U&, 128)
      If InStr(FIDOWNL$, AAA) < 1 Then
         TTZZ$ = TTZZ$ + vbCrLf + TRIM$(Left$(AAA, 64))
      End If
   Next U&
   TTXX$ = "Flavio, Hay que Actualizar Archivos en la WEB, Subiendo las Novedades de F:\WEBUPDATES\NOVEDAD al NOVETEST de la WEB. Gracias. Los Archivos Son:"
   MsgBox TTXX$ + TTZZ$, vbSystemModal
   '
   ' LIWEBUPD.SHOW 1
   '
99 End Sub

Sub REVIREGIS()
   openforms = DoEvents
   '
   Dim DIRMA$()
   Screen.MousePointer = 11
   '
   MAPISession1.DownLoadMail = False
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   MAPIMessages1.FetchUnreadOnly = True
   MAPIMessages1.Fetch
   CAMENSA& = MAPIMessages1.MsgCount
   Screen.MousePointer = 11
   '
   For U& = CAMENSA - 1000 To CAMENSA&
10   If U& > MAPIMessages1.MsgCount Then GoTo 99
     openforms = DoEvents
     MAPIMessages1.MsgIndex = U& - 1
     ORENVI$ = TRIM$(MAPIMessages1.MsgOrigDisplayName)
     ASUNTO$ = TRIM$(LCase$(MAPIMessages1.MsgSubject))
     If InStr(UCase$(ASUNTO$), "FLEXOFT WEB SERVICES") > 0 Then
       AAA = MAPIMessages1.MsgNoteText
       MsgBox AAA
     End If
89 Next U&
   '
99 MAPISession1.Action = 2  ' !!!
   '
End Sub
