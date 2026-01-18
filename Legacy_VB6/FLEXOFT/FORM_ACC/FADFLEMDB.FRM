VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Begin VB.Form FADFLEMDB 
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT Soluciones - Administracion de Actividad"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   12615
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0050765A&
      ForeColor       =   &H00404080&
      Height          =   8655
      Left            =   11000
      ScaleHeight     =   8595
      ScaleWidth      =   1635
      TabIndex        =   179
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   390
         ScaleHeight     =   225
         ScaleWidth      =   615
         TabIndex        =   192
         Top             =   7200
         Width           =   645
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   193
            Top             =   0
            Width           =   12000
         End
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
         Height          =   735
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   191
         Top             =   0
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
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   190
         ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
         Top             =   840
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
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   189
         Top             =   1680
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
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   2520
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   187
         Top             =   7560
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.CommandButton Command36 
         Caption         =   "M"
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
         Left            =   360
         TabIndex        =   186
         ToolTipText     =   "Migracion de la Base de Datos"
         Top             =   4215
         Width           =   735
      End
      Begin VB.CommandButton Command37 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   185
         ToolTipText     =   "Generador de Reportes - Consultas SQL"
         Top             =   6510
         Width           =   750
      End
      Begin VB.CommandButton Command35 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   184
         ToolTipText     =   "Apareo de Fuentes"
         Top             =   3375
         Width           =   735
      End
      Begin VB.Frame FRAME17 
         BackColor       =   &H0050765A&
         Caption         =   "Adjuntos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1500
         Index           =   7
         Left            =   0
         TabIndex        =   180
         Top             =   4935
         Width           =   1575
         Begin VB.ListBox List6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   60
            TabIndex        =   183
            Top             =   240
            Width           =   1450
         End
         Begin VB.CommandButton Command47 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   720
            TabIndex        =   182
            Top             =   1150
            Width           =   765
         End
         Begin VB.CommandButton Command47 
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   105
            TabIndex        =   181
            Top             =   1150
            Width           =   585
         End
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6000
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command391 
      Caption         =   "proceso"
      Height          =   195
      Left            =   9840
      TabIndex        =   169
      Top             =   550
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   10920
      Top             =   3960
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8625
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   15214
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   882
      ShowFocusRect   =   0   'False
      BackColor       =   13097420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Consultas Telefonicas"
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
      Tab(0).Control(20)=   "Command39"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Ordenes de Trabajo"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame33"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "AGenRep"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame30"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command38"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command21"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame11"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame13"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Command18"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame19"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame18"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame16"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame15"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame14"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Frame12"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame9"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Command10"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Command9"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Command8"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label24"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "Tareas Internas"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label34"
      Tab(2).Control(1)=   "Command23"
      Tab(2).Control(2)=   "Frame22"
      Tab(2).Control(3)=   "Frame23"
      Tab(2).Control(4)=   "Frame24"
      Tab(2).Control(5)=   "Command29"
      Tab(2).Control(6)=   "Command30"
      Tab(2).Control(7)=   "Frame27"
      Tab(2).Control(8)=   "Frame28"
      Tab(2).Control(9)=   "Command31"
      Tab(2).Control(10)=   "Frame34"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Control WEB"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame31"
      Tab(3).Control(1)=   "Command19"
      Tab(3).Control(2)=   "Frame32"
      Tab(3).ControlCount=   3
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -75000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   178
         Top             =   520
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.CommandButton Command39 
         Caption         =   "Pendiente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9765
         TabIndex        =   173
         Top             =   5160
         Width           =   855
      End
      Begin VB.Frame Frame34 
         Caption         =   "Prioridad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68010
         TabIndex        =   171
         Top             =   1860
         Width           =   1455
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Left            =   1030
            TabIndex        =   172
            Top             =   90
            Width           =   375
         End
      End
      Begin VB.Frame Frame33 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -65970
         TabIndex        =   167
         Top             =   3120
         Width           =   750
         Begin VB.Label Label37 
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
            Left            =   120
            TabIndex        =   168
            ToolTipText     =   "Doble Click para Cambiar"
            Top             =   315
            Width           =   500
         End
      End
      Begin VB.Frame Frame32 
         Caption         =   "Ordenamiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74760
         TabIndex        =   161
         Top             =   600
         Width           =   7935
         Begin VB.OptionButton Option6 
            Caption         =   "Por Tipo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   6120
            TabIndex        =   165
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Por Archivo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   4440
            TabIndex        =   164
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Por Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   163
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Por Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   1440
            TabIndex        =   162
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Revisar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   -66480
         TabIndex        =   160
         Top             =   680
         Width           =   2055
      End
      Begin VB.Frame Frame31 
         Caption         =   "Archivos en la WEB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7155
         Left            =   -74760
         TabIndex        =   157
         Top             =   1320
         Width           =   10395
         Begin VB.ListBox List5 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6180
            Left            =   120
            TabIndex        =   166
            Top             =   630
            Width           =   10020
         End
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
            Height          =   6180
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   158
            Top             =   630
            Visible         =   0   'False
            Width           =   10020
         End
         Begin VB.Line Line9 
            X1              =   120
            X2              =   10080
            Y1              =   330
            Y2              =   330
         End
         Begin VB.Label Label36 
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
            Left            =   240
            TabIndex        =   159
            Top             =   360
            Width           =   9990
         End
         Begin VB.Line Line8 
            BorderWidth     =   2
            X1              =   130
            X2              =   130
            Y1              =   350
            Y2              =   630
         End
         Begin VB.Line Line7 
            X1              =   10080
            X2              =   10080
            Y1              =   315
            Y2              =   630
         End
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   -65040
         Style           =   1  'Graphical
         TabIndex        =   156
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3050
         Width           =   740
      End
      Begin VB.Frame Frame30 
         Caption         =   "Resp."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -66810
         TabIndex        =   154
         Top             =   3120
         Width           =   750
         Begin VB.Label Label35 
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
            Left            =   120
            TabIndex        =   155
            ToolTipText     =   "Doble Click para Cambiar"
            Top             =   315
            Width           =   500
         End
      End
      Begin VB.CommandButton Command38 
         Caption         =   "Liq."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Left            =   -64815
         Style           =   1  'Graphical
         TabIndex        =   153
         ToolTipText     =   "Liquidacion por Cliente"
         Top             =   7400
         Width           =   540
      End
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
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Caption         =   "Frame29"
         Height          =   525
         Left            =   5670
         TabIndex        =   140
         Top             =   3120
         Width           =   3900
         Begin VB.CommandButton Command40 
            Caption         =   "No Conform."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   3000
            TabIndex        =   170
            Top             =   0
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   1320
            TabIndex        =   144
            Top             =   220
            Width           =   1065
         End
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Nuevo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   1320
            TabIndex        =   143
            Top             =   0
            Width           =   1665
         End
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   142
            Top             =   220
            Width           =   1065
         End
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   141
            Top             =   0
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   9765
         Style           =   1  'Graphical
         TabIndex        =   139
         ToolTipText     =   "Cierre general de CTel's"
         Top             =   4440
         Width           =   855
      End
      Begin VB.CommandButton Command32 
         Caption         =   "List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   9765
         Style           =   1  'Graphical
         TabIndex        =   133
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
         Style           =   1  'Graphical
         TabIndex        =   132
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
         TabIndex        =   125
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
            TabIndex        =   126
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
            TabIndex        =   127
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
         TabIndex        =   117
         Top             =   3150
         Width           =   4950
         Begin VB.OptionButton Option4 
            Caption         =   "Jorge Real"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   2520
            TabIndex        =   122
            Top             =   840
            Visible         =   0   'False
            Width           =   1410
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
            Style           =   1  'Graphical
            TabIndex        =   124
            ToolTipText     =   "Listado Tareas Pendientes por Responsable"
            Top             =   210
            Width           =   630
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   9
            Left            =   3840
            TabIndex        =   131
            Top             =   840
            Visible         =   0   'False
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
            TabIndex        =   130
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
            TabIndex        =   129
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
            TabIndex        =   128
            Top             =   0
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Omar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   1920
            TabIndex        =   123
            Top             =   360
            Width           =   1065
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Lorena"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   1560
            TabIndex        =   121
            Top             =   840
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Matias"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   720
            TabIndex        =   120
            Top             =   840
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Flavio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   600
            TabIndex        =   119
            Top             =   360
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton Option4 
            Caption         =   "EPB"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   118
            Top             =   840
            Visible         =   0   'False
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
         Style           =   1  'Graphical
         TabIndex        =   107
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
         Style           =   1  'Graphical
         TabIndex        =   106
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
         TabIndex        =   99
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
            TabIndex        =   114
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
               TabIndex        =   115
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
               TabIndex        =   116
               TabStop         =   0   'False
               ToolTipText     =   "Fecha Enviada por Mail"
               Top             =   525
               Width           =   990
            End
            Begin VB.Image Image3 
               Height          =   330
               Left            =   900
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
            TabIndex        =   108
            Top             =   1155
            Width           =   3375
            Begin VB.OptionButton Option3 
               Caption         =   "Jorge"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   4
               Left            =   2520
               TabIndex        =   113
               Top             =   840
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Omar"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   138
               Top             =   480
               Width           =   960
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Lorena"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   3
               Left            =   1680
               TabIndex        =   112
               Top             =   840
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Matias"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   720
               TabIndex        =   111
               Top             =   840
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Flavio"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   420
               TabIndex        =   110
               Top             =   480
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton Option3 
               Caption         =   "EPB"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   0
               TabIndex        =   109
               Top             =   840
               Visible         =   0   'False
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
            TabIndex        =   101
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
            TabIndex        =   100
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
            TabIndex        =   136
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
            TabIndex        =   135
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
            TabIndex        =   134
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
            TabIndex        =   103
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
            TabIndex        =   105
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
            TabIndex        =   104
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
            TabIndex        =   102
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
         TabIndex        =   97
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
            TabIndex        =   98
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
         TabIndex        =   95
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
            TabIndex        =   96
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
         Style           =   1  'Graphical
         TabIndex        =   94
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
         Height          =   1095
         Left            =   9765
         TabIndex        =   91
         Top             =   5580
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
            TabIndex        =   92
            Top             =   750
            Width           =   750
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   240
            Top             =   255
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   180
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.CommandButton Command22 
         Caption         =   "O.Trab."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   9765
         Style           =   1  'Graphical
         TabIndex        =   90
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
         Height          =   650
         Left            =   -65025
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Informe Tareas Por Fecha"
         Top             =   3700
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
         Height          =   1455
         Left            =   -66810
         TabIndex        =   87
         Top             =   3840
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
            Height          =   1050
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   88
            Top             =   240
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
         Height          =   1200
         Left            =   -74895
         TabIndex        =   50
         Top             =   7300
         Width           =   1590
         Begin VB.CommandButton Command16 
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
            Height          =   250
            Left            =   1305
            TabIndex        =   80
            ToolTipText     =   "Fecha de Cierre"
            Top             =   860
            Width           =   160
         End
         Begin VB.CommandButton Command15 
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
            Height          =   250
            Left            =   1305
            TabIndex        =   79
            ToolTipText     =   "Fecha Verificada Instalación"
            Top             =   580
            Width           =   160
         End
         Begin VB.CommandButton Command13 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   1305
            TabIndex        =   78
            ToolTipText     =   "Fecha Enviada por Mail"
            Top             =   300
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
            Left            =   375
            MaxLength       =   12
            TabIndex        =   71
            TabStop         =   0   'False
            Text            =   " "
            ToolTipText     =   "Fecha Verificada Instalación"
            Top             =   580
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
            Height          =   285
            Left            =   375
            MaxLength       =   12
            TabIndex        =   70
            TabStop         =   0   'False
            ToolTipText     =   "Fecha Enviada por Mail"
            Top             =   300
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
            Left            =   375
            MaxLength       =   12
            TabIndex        =   69
            TabStop         =   0   'False
            Text            =   " "
            ToolTipText     =   "Doble Click para Cerrar"
            Top             =   860
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
            Left            =   105
            TabIndex        =   74
            ToolTipText     =   "Fecha Verificada Instalación"
            Top             =   580
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
            TabIndex        =   73
            ToolTipText     =   "Fecha Enviada por Mail"
            Top             =   300
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
            Left            =   105
            TabIndex        =   72
            ToolTipText     =   "Fecha de Cierre"
            Top             =   860
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
         Height          =   1720
         Left            =   9765
         TabIndex        =   83
         Top             =   6680
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
            Style           =   1  'Graphical
            TabIndex        =   85
            ToolTipText     =   "Adjuntar Archivos"
            Top             =   240
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
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Envia por E-Mail"
            Top             =   960
            Width           =   645
         End
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
         Height          =   855
         Left            =   -65025
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Informe Tareas Por Fecha"
         Top             =   4350
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
         Height          =   1200
         Left            =   -67335
         TabIndex        =   64
         Top             =   7300
         Width           =   2430
         Begin VB.TextBox Text25 
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
            Height          =   260
            Left            =   1305
            TabIndex        =   66
            TabStop         =   0   'False
            Text            =   " "
            Top             =   735
            Width           =   1005
         End
         Begin VB.TextBox Text24 
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
            Height          =   260
            Left            =   1305
            TabIndex        =   65
            TabStop         =   0   'False
            Text            =   " "
            Top             =   420
            Width           =   1005
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Importe ($):"
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
            Left            =   100
            TabIndex        =   68
            Top             =   765
            Width           =   1100
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Nro.Factura:"
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
            Left            =   100
            TabIndex        =   67
            Top             =   450
            Width           =   1100
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
         Height          =   1200
         Left            =   -73215
         TabIndex        =   57
         Top             =   7300
         Width           =   1485
         Begin VB.TextBox Text19 
            Alignment       =   1  'Right Justify
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
            Height          =   260
            Left            =   735
            TabIndex        =   62
            TabStop         =   0   'False
            Text            =   " "
            Top             =   860
            Width           =   660
         End
         Begin VB.TextBox Text18 
            Alignment       =   1  'Right Justify
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
            Height          =   260
            Left            =   735
            TabIndex        =   59
            TabStop         =   0   'False
            Text            =   " "
            Top             =   300
            Width           =   660
         End
         Begin VB.TextBox Text17 
            Alignment       =   1  'Right Justify
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
            Height          =   260
            Left            =   735
            TabIndex        =   58
            TabStop         =   0   'False
            Text            =   " "
            Top             =   580
            Width           =   660
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   63
            Top             =   855
            Width           =   645
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   61
            Top             =   300
            Width           =   645
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   60
            Top             =   585
            Width           =   645
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Descripción / Detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         Left            =   -74895
         TabIndex        =   54
         Top             =   3120
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
            Height          =   1770
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   55
            Top             =   240
            Width           =   7755
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Descripcion / Síntoma / Fallo Reportado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -74880
         TabIndex        =   52
         Top             =   630
         Width           =   4635
         Begin VB.OptionButton Option1 
            Caption         =   "Varios"
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
            Index           =   9
            Left            =   3570
            TabIndex        =   150
            Top             =   630
            Width           =   1005
         End
         Begin VB.OptionButton Option1 
            Caption         =   "N/Desarr."
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
            Index           =   3
            Left            =   2205
            TabIndex        =   149
            Top             =   630
            Width           =   1245
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Modif."
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
            Left            =   1155
            TabIndex        =   148
            Top             =   630
            Width           =   930
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Repar."
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
            Index           =   1
            Left            =   105
            TabIndex        =   147
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
            TabIndex        =   146
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
            Height          =   1395
            Left            =   105
            MaxLength       =   256
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   53
            Top             =   900
            Width           =   4400
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Desarrollo del Trabajo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Left            =   -74895
         TabIndex        =   51
         Top             =   5250
         Width           =   10620
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
            TabIndex        =   77
            Top             =   630
            Width           =   10365
         End
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   130
            X2              =   130
            Y1              =   350
            Y2              =   630
         End
         Begin VB.Line Line2 
            X1              =   10185
            X2              =   10185
            Y1              =   315
            Y2              =   630
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   10185
            Y1              =   330
            Y2              =   330
         End
         Begin VB.Label Label11 
            BackColor       =   &H80000016&
            Caption         =   " Fecha    Detalle / Descripcion                            Com   Fin   Hs.  Rsp. L/C "
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
            TabIndex        =   56
            Top             =   360
            Width           =   10200
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
         Height          =   1200
         Left            =   -71640
         TabIndex        =   46
         Top             =   7300
         Width           =   4215
         Begin VB.OptionButton Option2 
            Caption         =   "Pres.Fijo"
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
            Left            =   2310
            TabIndex        =   152
            Top             =   735
            Width           =   1305
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Vrs.S.C.Ad."
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
            Left            =   2835
            TabIndex        =   76
            Top             =   420
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Liq.Horas."
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
            Left            =   1470
            TabIndex        =   49
            Top             =   420
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Garantía"
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
            Left            =   840
            TabIndex        =   48
            Top             =   735
            Width           =   1100
         End
         Begin VB.OptionButton Option2 
            Caption         =   "A-Manten."
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
            TabIndex        =   47
            Top             =   420
            Width           =   1305
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Referencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -70170
         TabIndex        =   36
         Top             =   630
         Width           =   4950
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
            Height          =   750
            Left            =   240
            MaxLength       =   126
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   175
            Top             =   1560
            Width           =   4545
         End
         Begin VB.ComboBox Combo2 
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
            Height          =   315
            Left            =   2630
            Style           =   2  'Dropdown List
            TabIndex        =   174
            Top             =   1200
            Width           =   2190
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
            TabIndex        =   86
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
            TabIndex        =   81
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
            TabIndex        =   41
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
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   " "
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            Caption         =   "Observaciones:"
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
            Left            =   80
            TabIndex        =   177
            Top             =   1320
            Width           =   1365
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            Left            =   2760
            TabIndex        =   43
            Top             =   320
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
            TabIndex        =   42
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
            TabIndex        =   40
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
            TabIndex        =   38
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
            TabIndex        =   37
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
         Height          =   650
         Left            =   -65025
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Imprime Formulario de Consulta"
         Top             =   1385
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
         Height          =   650
         Left            =   -65025
         Style           =   1  'Graphical
         TabIndex        =   34
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
         Height          =   650
         Left            =   -65025
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Nueva Consulta"
         Top             =   2035
         Width           =   740
      End
      Begin VB.CommandButton Command7 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   9765
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Nueva Consulta"
         Top             =   740
         Width           =   855
      End
      Begin VB.Frame Frame8 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   7080
         TabIndex        =   28
         Top             =   3675
         Width           =   2490
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
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
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   315
            Width           =   2295
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Número"
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1130
         Left            =   7875
         TabIndex        =   25
         Top             =   630
         Width           =   1695
         Begin VB.VScrollBar VScroll1 
            Height          =   360
            Left            =   1365
            TabIndex        =   145
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
            Left            =   1280
            TabIndex        =   93
            ToolTipText     =   "Seleccion Consultas Pendientes"
            Top             =   720
            Width           =   285
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Buscar"
            Height          =   270
            Left            =   105
            TabIndex        =   30
            Top             =   735
            Width           =   915
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   90
            TabIndex        =   32
            Top             =   315
            Width           =   1305
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Origen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1130
         Left            =   5670
         TabIndex        =   20
         Top             =   630
         Width           =   2115
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
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
            Left            =   735
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   " "
            Top             =   700
            Width           =   1035
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
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
            Left            =   735
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   " "
            Top             =   315
            Width           =   1035
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Hora"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   240
            TabIndex        =   24
            Top             =   795
            Width           =   405
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   22
            Top             =   420
            Width           =   525
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   9765
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Graba las Modificaciones"
         Top             =   1480
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   9765
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprime Formulario de Consulta"
         Top             =   2220
         Width           =   855
      End
      Begin VB.Frame Frame5 
         Caption         =   "Distribuidor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
            Appearance      =   0  'Flat
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
            TabIndex        =   151
            Top             =   315
            Width           =   2955
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
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
            Caption         =   "C/C"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   11
            Top             =   795
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Para"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   9
            Top             =   375
            Width           =   540
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "  Referencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
            Appearance      =   0  'Flat
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
            Name            =   "Arial"
            Size            =   9
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
            Appearance      =   0  'Flat
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
            Name            =   "Arial"
            Size            =   9
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
         Width           =   6765
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
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
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   315
            Width           =   6495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            Caption         =   "Fax"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2900
            TabIndex        =   27
            Top             =   1220
            Width           =   435
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Tel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   240
            TabIndex        =   26
            Top             =   1215
            Width           =   195
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
         TabIndex        =   137
         Top             =   3885
         Width           =   750
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   -65025
         TabIndex        =   75
         Top             =   2730
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin VB.Label Label39 
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
      Left            =   0
      TabIndex        =   176
      Top             =   0
      Width           =   645
   End
End
Attribute VB_Name = "FADFLEMDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NUCETE$
Dim ABUSQUE$
Dim ENCARLISTA%
Dim MIGRANDO%
Dim MODIFIC_CTE% 'BANDERA DE MODIFICACION DE REFERENCIA

Dim FORDEN%
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Sub CREATABLA_DOCUREL_CT()
   Call CREACAMPO("ADMIFLEX", "DOCUREL_CT", "CORTRAB", 10, 16, 1)
   Call CREACAMPO("ADMIFLEX", "DOCUREL_CT", "Descripcion", 10, 24)
   Call CREACAMPO("ADMIFLEX", "DOCUREL_CT", "Ruta", 10, 512)
   Call CREACAMPO("ADMIFLEX", "DOCUREL_CT", "Marborra", 3, 0)

End Sub

Private Sub AGenRep_Click()
    '
'   OPX% = ALTERNA%("OT's Pendientes Por Cliente\Detalle Actividad por Cliente y O/T\Detalle de OT's por Cliente y Rango de Fechas\Indice de OT's de UN Cliente\Control de Actividad por Técnico y R.Fechas\Actividad por Clientes (Uno o Todos) y R.Fechas\Actividad por UN Cliente y R.Fechas")
'   If OPX% = 1 Then
'       Call OTPENCLI
'     ElseIf OPX% = 2 Then
'       Call OTDETCLI
'     ElseIf OPX% = 3 Then
'       Call OTTOTCLI
'     ElseIf OPX% = 4 Then
'       Call INDIOTCLI
'     ElseIf OPX% = 5 Then
'       Call CONORTRA
'     ElseIf OPX% = 6 Then
'       Call CONORTRB
'     ElseIf OPX% = 7 Then
'       Call CONORTRD
'   End If
'   '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Combo1_Click()
   Text10 = ""
   If Combo1.ListIndex > 0 Then
      Text10 = Combo1.List(0)
   End If
End Sub

Private Sub Combo2_Click()
    '
    Text8 = ""
    If Combo2.ListIndex > 0 Then
       Text8 = Combo2.List(Combo2.ListIndex)
    End If
    '
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
    Call CARGALISEL("CORTRABS", "CORTRAB", "Nume_Ot/A12;Referencia/A66", "Nume_OT <> '' order by Nume_OT ASC")
    Call SELECHO("CORTRABS")
    NUCOTE$ = AJUSTI$(VALACT1$("CORTRABS"), 12)
    If TRIM$(NUCOTE$) <> "" Then
       Label15.Caption = NUCOTE$
    End If
    '
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call CONECRUN("FSETUP04", "Configuración de Funcionamiento")
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
      USDESTI$ = SAFETEXT$(VALOBADA("COTELEF", "Solicitante", "NroClie = " & NC% & " ORDER BY FECHEMIS DESC ", "NOMESS"))
      Call CliCombo(Combo1, USDESTI$)
      If Combo1.ListCount < 1 Then
           Combo1.AddItem AJUSTI$(USDESTI$, 64)
           Combo1.ListIndex = Combo1.ListCount - 1
         Else
           Combo1.ListIndex = 0
      End If
      '
15    Combo1.Refresh
      Text10 = ""
      If Combo1.ListIndex > 0 Then
          Text10 = Combo1.List(0)
      End If
      Exit Sub
      '
    End If
    '
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "ARCHIG07/AMACLI"
    Call CONECRUN(ACONEC$, "Alta de Nueva Cuenta de Cliente")
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
    CONDI$ = "(STATUS < 3 OR STATUS IS NULL) "
    CONDI$ = CONDI$ & " AND (Fech_Cierre < '01/01/90' OR Fech_Cierre IS NULL) "
    CONDI$ = CONDI$ + " ORDER BY Nume_OT ASC"
    Call CARGALISEL("!CORTRASE", "CORTRAB", "Nume_OT/A12;Referencia/A66", CONDI$)
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
   Call ITARFE09.SOLOPROPIO
   ITARFE09.Show 1
   Command18.Enabled = True
End Sub

Private Sub Command19_Click()
   '
   Command19.Enabled = False
   Screen.MousePointer = 11
   '
   ' DETERMINA LA CARPETA PARA EL DOWNLOAD Y LA CREA
   CARPEDOWN$ = App.Path
   While Len(CARPEDOWN$) > 0 And Right$(CARPEDOWN$, 1) <> "\"
      CARPEDOWN$ = Left$(CARPEDOWN$, Len(CARPEDOWN$) - 1)
   Wend
   CARPEDOWN$ = CARPEDOWN$ + "DOWNLOAD\"
   '
   Call BORRACACHE
   JJ1& = URLDownloadToFile(0, "http://www.flexoft.com.ar/UPDATES/NOVETEST/IUPDATES.TXT", CARPEDOWN$ + "IUPDATES.TXT", 0, 0)
   '
   FIDOWNL$ = LOADFILE$(CARPEDOWN$ + "IUPDATES.TXT")
   List4.Clear: List5.Clear
   Select Case FORDEN%
      Case 0    ' POR FECHA
        ENCA$ = "Fecha - Hora   Archivo                       Empresa       Usuario           "
      Case 1    ' POR EMPRESA
        ENCA$ = "Empresa        Usuario         Fecha - Hora    Archivo                           "
       Case 2    ' POR ARCHIVO
        ENCA$ = "Archivo                      Empresa         Usuario         Fecha - Hora    "
   End Select
   Label36 = ENCA$
   '
   For U& = 1 To Len(FIDOWNL$) Step 128
      AAA = Mid$(FIDOWNL$, U&, 128)
      ANAME$ = TRIM$(Left$(AAA, 64))
      fex = XVALO(Mid$(AAA, 65, 5))
      HOXI% = XVALO(Mid$(AAA, 70, 5))
      USDESTI$ = Mid$(AAA, 75, 12)
      SVDESTI$ = TRIM$(Mid$(AAA, 87, 16))
      If SVDESTI$ <> "" Then
         DCLI$ = TRIM$(VALOBADA("DEUDOR12", "FANT_CLI", "Hora_Cli = '" & SVDESTI$ & "'", "NOMESS"))
         If DCLI$ = "" Then
            DCLI$ = TRIM$(VALOBADA("DEUDOR12", "RASO_CLI", "Hora_Cli = '" & SVDESTI$ & "'", "NOMESS"))
         End If
         If DCLI$ = "" Then DCLI$ = SVDESTI$
      End If
      '
      TTXX$ = Space$(128)
      Select Case FORDEN%
         Case 0    ' POR FECHA
           FESS$ = FECHATEX$(fex)
           FESS$ = Mid$(FESS$, 7, 2) + "/" + Mid$(FESS$, 4, 2) + "/" + Left$(FESS$, 2) + " - " + HORATEX$(HOXI%)
           Mid$(TTXX$, 1, 16) = FESS$
           Mid$(TTXX$, 19, 32) = ANAME$
           Mid$(TTXX$, 53, 16) = DCLI$
           Mid$(TTXX$, 69, 16) = USDESTI$
         Case 1    ' POR EMPRESA
           Mid$(TTXX$, 1, 16) = DCLI$
           Mid$(TTXX$, 19, 16) = USDESTI$
           FESS$ = FECHATEX$(fex)
           Mid$(TTXX$, 37, 16) = FESS$ + " - " + HORATEX$(HOXI%)
           Mid$(TTXX$, 55, 32) = ANAME$
          Case 2    ' POR ARCHIVO
           Mid$(TTXX$, 1, 32) = ANAME$
           Mid$(TTXX$, 35, 16) = DCLI$
           Mid$(TTXX$, 53, 16) = USDESTI$
           FESS$ = FECHATEX$(fex)
           Mid$(TTXX$, 71, 16) = FESS$ + " - " + HORATEX$(HOXI%)
      End Select
      '
      List4.AddItem TTXX$
      '
   Next U&
   '
   For U& = 1 To List4.ListCount
      TTXX$ = List4.List(U& - 1)
      If FORDEN% = 0 Then
         AX$ = Left$(TTXX$, 2): DX$ = Mid$(TTXX$, 7, 2)
         Call REPLA(TTXX$, DX$, 1, 2)
         Call REPLA(TTXX$, AX$, 7, 2)
      End If
      List5.AddItem TTXX$
   Next U&
   Screen.MousePointer = 1
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
   NUCOTE$ = TRIM$(Label15)
   If NUCOTE$ <> "" Then
      If USNBR% = 101 Then
        If COMALTER%("Está Seguro de Borrar\la Orden de Trabajo '" + TRIM$(Label15) + "' ?\\No, Conservar\Si, Borrar") = 2 Then
          Screen.MousePointer = 11
          CONDI$ = "Nume_Ot = '" & NUCOTE$ & "'"
          Call BORRAREGISTROS("CORTRAB", CONDI$, RAFE&)
          Call BORRAREGISTROS("DETOTRA", CONDI$, RAFE&)
          Call COPYSTRU("CORTRAB1", "CORTRABS")
          Call CARGALISEL("CORTRABS", "CORTRAB", "Nume_Ot/A12;Referencia/A66", "Nume_OT <> '' order by Nume_OT ASC")
          Screen.MousePointer = 1
        End If
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
       CONDI$ = "Nume_Tar = " & NTAR&
       Call BORRAREGISTROS("TARINTER", CONDI$, RAFE&, "NOMESS")
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
   Call CIERRARCH("!COTELEF")
   Call COPYSTRU("COTELEF1", "COTELEF")
   CONDI$ = "Nume_Cotel <> '' "
   CONDI$ = CONDI$ + " AND (Fech_Cierre <= '30/NOV/1999' OR Fech_Cierre IS NULL) ORDER BY FECHEMIS ASC, NUME_COTEL ASC"
   Call CARGALISEL("!COTELEF", "COTELEF", "Nume_Cotel/A12;FECHEMIS/D0;Referencia/A66", CONDI$, "DISTINCT")
   '
   FORSELEC.ABUSQUE = Label7
   Call SELECHO("!COTELEF")
   NUCOTE$ = AJUSTI$(VALACT1$("!COTELEF"), 12)
   If TRIM$(NUCOTE$) <> "" Then
       Label7.Caption = NUCOTE$
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
   '
   Call CONECRUN("?TARINTPEN(" & IRESPO% & ")", "Tareas Pendientes por Responsable")
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
      CIERRCTEL07.Label2 = TRIM$(Str$(DES%))
         DESDE = FETEXCOR1$(DES%)
      CIERRCTEL07.Label3 = TRIM$(Str$(HAS%))
         HASTA = FETEXCOR1$(HAS%)
      CIERRCTEL07.List1.Clear     ' EPB - NECESARIO PORQUE SI NO AL VOLVER A ENTRAR DUPLICA LA CARGA
      CONDI0$ = "FechEmis >= '" & DESDE & "' "
      CONDI0$ = CONDI0$ & " AND FechEmis <= '" & HASTA & "' "
      CONDI$ = CONDI0$ & " AND (Fech_Cierre <= '30/NOV/1999' OR Fech_Cierre IS NULL)"
      Call CARGALISTA(CIERRCTEL07.List1, "COTELEF", "Nume_Cotel/A12;Referencia/A66", CONDI$)
      '
      CIERRCTEL07.Text2 = CantRegistros("COTELEF", CONDI0$)
      CIERRCTEL07.Text3 = CantRegistros("COTELEF", CONDI$)
      Screen.MousePointer = 1
      '
      CIERRCTEL07.Show 1
      Call ACTULISEL
      '
    End If
    '
    Command27.Enabled = True
    '
End Sub

Private Sub Command28_Click()
    Call CARGALISEL("TARINTER", "TARINTER", "Nume_Tar/F5.0;Des_Corta/A60", "Nume_Tar > 0")
    Call SELECHO("TARINTER")
    NTAR& = Val(VALACT1$("TARINTER"))
    If NTAR& > 0 Then
      Label28 = TRIM$(Str$(NTAR&))
    End If
End Sub

Private Sub Command29_Click()
    Command29.Enabled = False
    UTAR& = XVALO(VALOBADA("TARINTER", "MAX(Nume_Tar)", "Nume_Tar > 0", "NOMESS"))
    TOPLIS = List3.TopIndex
    '
    NTAR& = XVALO(Label28)
    If NTAR& > 0 Then
      If TRIM$(Text29) = "" Then
        Call FRATEXTO(Text28, 52)
        Text29 = RETEX$(1)
        Text29.Refresh
      End If
      CONDI$ = "Nume_Tar = " & NTAR&
      QTR& = CantRegistros("TARINTER", CONDI$)
      If QTR& < 1 Then
          Call CreaRegistro("TARINTER", "Nume_Tar", NTAR&)
      End If
        TTYY$ = TRIM$(Label29) + " - " + TRIM$(Text29)
      Call ACTUREGISTRO("TARINTER", "Des_Corta", CONDI$, TTYY$, RAFE&, "NOMESS")
        FEXI% = FECHANUM(Label29)
        FENUE$ = FETEXCOR1$(FEXI%)
      Call ACTUREGISTRO("TARINTER", "Fech_Emis", CONDI$, FENUE$, RAFE&, "NOMESS")
        FEXI% = FECHANUM(Text31)
        FENUE$ = FETEXCOR1$(FEXI%)
      Call ACTUREGISTRO("TARINTER", "Fech_Cier", CONDI$, FENUE$, RAFE&, "NOMESS")
        IRESPO% = 0
        For KKI% = 1 To 5
          If Option3(KKI%).Value = True Then IRESPO% = KKI%
        Next KKI%
      Call ACTUREGISTRO("TARINTER", "ITec_Resp", CONDI$, IRESPO%, RAFE&, "NOMESS")
      Call ACTUREGISTRO("TARINTER", "Fech_Prog", CONDI$, CVDAT(0), RAFE&, "NOMESS")
      Call ACTUREGISTRO("TARINTER", "Refe_Tar", CONDI$, RTrim(Text29), RAFE&, "NOMESS")
      Call ACTUREGISTRO("TARINTER", "Deta_Tar", CONDI$, Text28, RAFE&, "NOMESS")
      Call ACTUREGISTRO("TARINTER", "Obse_Tar", CONDI$, Text13, RAFE&, "NOMESS")
      Call ACTUREGISTRO("TARINTER", "Prio_Tar", CONDI$, Label38, RAFE&, "NOMESS")
      '
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
'    Call COPYSTRU("COTELEF1", "COTELEF")
'    Call CIERRARCH("COTELEF")
'    CONDI$ = "Nume_Cotel <> '' ORDER BY FECHEMIS ASC, NUME_COTEL ASC"
'    Call CARGALISEL("COTELEF", "COTELEF", "Nume_Cotel/A12;Referencia/A66;FECHEMIS/D0", CONDI$, "DISTINCT")
'    '

    ' Actualiza Listas de Seleccion
    Call COPYSTRU("COTELEF1", "COTELEF")
    Call CIERRARCH("COTELEF")
    CONDI$ = "Nume_Cotel <> '' ORDER BY FECHEMIS ASC, NUME_COTEL ASC"
    Call CARGALISEL("COTELEF", "COTELEF", "Nume_Cotel/A12;Referencia/A66;FECHEMIS/D0", CONDI$, "DISTINCT")

    FORSELEC.ABUSQUE = Label7
    Call SELECHO("COTELEF")
    NUCOTE$ = AJUSTI$(VALACT1$("COTELEF"), 12)
    If TRIM$(NUCOTE$) <> "" Then
       Label7.Caption = NUCOTE$
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
   UTAR& = XVALO(VALOBADA("TARINTER", "MAX(Nume_Tar)", "Nume_Tar > 0", "NOMESS"))
   If UTAR& > 0 Then
      PROTARIN& = 1 + UTAR&
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

Sub LICOTELP(modo%)
   '
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ <> "" Then
      REBLA$ = REGBLAN$("LICOTEL")
      JJ& = 0
      For U& = 1 To ULTREG&("COTELEF1")
        XX$ = REGLEIDO$("COTELEF1", U&)
        FEMI% = CVI(Mid$(XX$, 107, 2))
        FECI% = CVI(Mid$(XX$, 251, 2))
        If FECI% <= 8224 Or modo% = 2 Then        ' NO ESTÁ CERRADA
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
      If modo% = 1 Then Call HEADERS("LICOTEL", "Pendientes - Período: " + PERIO$)
      If modo% = 2 Then Call HEADERS("LICOTEL", "Todas - Período: " + PERIO$)
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
            moti% = Asc(Mid$(XX$, 253, 1))
            If moti% > 3 Then moti% = 4         ' NO CLASIFICADOS
            CANPORMOT%(moti%) = CANPORMOT%(moti%) + 1
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
   fex = DIPOST(FINMESACT)
6  PROVEN$ = FECHATEX$(fex)
   If Left$(PROVEN$, 2) <> "01" Then
      fex = DIPOST(fex)
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
        Concepto$ = UCase$(TRIM$(CARELEC.Text5))
        If Len(Concepto$) < 6 Then
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
         Call REPLA(BBX$, Concepto$, 3, 20)
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
   'Call CONECRUN("FLEX_ADM/COMPAFU", "Comparacion de Fuentes")
   APARFUEN2.Show 1
End Sub

Private Sub Command36_Click()
   Command36.Enabled = False
   Screen.MousePointer = 11
   '
   MIGRANDO% = 1
   '
   MsgBox "SOLO GRABA TARINTER"
   Call MIGRACLIEN
   Call MIGRACOTEL
   Call MIGRAOTRAB
   Call MIGRATARIN
   '
   MIGRANDO% = 0
   '
   Screen.MousePointer = 1
   Command36.Enabled = True
End Sub

Private Sub Command37_Click()
   Command37.Enabled = False
   Call CONECRUN("GENREP07", "Generador de Reportes SQL")
   Command37.Enabled = True
End Sub

Private Sub Command38_Click()
   Command38.Enabled = False
   OBXX$ = OBJPLA$("DESHASFECHCLI", "")
   If OBXX$ = "" Then GoTo 99
   '
   DES% = CVI(Left$(OBXX$, 2)): DESFE$ = FETEXCOR1$(DES%)
   HAS% = CVI(Mid$(OBXX$, 3, 2)): HASFE$ = FETEXCOR1$(HAS%)
   NCLIE& = CVI(Mid$(OBXX$, 5, 2))
   Screen.MousePointer = 11
   CONDI$ = "NroClie = " & NCLIE&
   CONDI$ = CONDI$ & " AND FechEjec >= '" & DESFE$ & "' "
   CONDI$ = CONDI$ & " AND FechEjec <= '" & HASFE$ & "'"
   SQLX$ = "SELECT * FROM DETOTRA "
   SQLX$ = SQLX$ + " INNER JOIN CORTRAB "
   SQLX$ = SQLX$ + " ON DETOTRA.Nume_OT = CORTRAB.Nume_OT "
   SQLX$ = SQLX$ + " Where " & CONDI$
   Dim DETRA As ADODB.Recordset
   Set DETRA = New ADODB.Recordset
   DETRA.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   
   With DETRA
      Do While Not .EOF
         NUOT$ = SAFETEXT$(!nume_ot)
         CONDJ$ = "Nume_OT = '" & NUOT$ & "' "
         moti% = XVALO(VALOBADA("CORTRAB", "Motivo", CONDJ$, "NOMESS"))
         IMPU% = XVALO(VALOBADA("CORTRAB", "Imputa", CONDJ$, "NOMESS"))
         VUNI = 0: VUNI2 = 0
         If IMPU% = 3 Then   ' LIQUIDACION DE HORAS
            VUNI = 75: VUNI2 = 90
            RESP$ = TRIM$(UCase$(Left$(SAFETEXT$(!RespEjec), 3)))
            If RESP$ = "FG" Then VUNI = 90: VUNI2 = 120
            If RESP$ = "EPB" Or RESP$ = "EB" Then VUNI = 120: VUNI2 = 150
              NC% = NCLIE&:
              If InStr(RASOCLI$(NC%), "ACHER") > 0 Then
                 VUNI = 135: VUNI2 = 135
              End If
            LUEJ$ = TRIM$(UCase$(Left$(SAFETEXT$(!LugEjec), 3)))
            If LUEJ$ = "COM" Then VUNI = VUNI2
         End If
         HOCOM% = XVALO(!HORACOM)
         HOFIN% = XVALO(!HORAFIN)
         TEJEC = DIFHORA(HOCOM%, HOFIN%)
         !TiemEjec = TEJEC
         !TarifHor = VUNI
         !Importe = VUNI * TEJEC
         .Update
      .MoveNext
      Loop
   End With
   Screen.MousePointer = 1
   Call CONECRUN("?LIQMENCLI(" & NCLIE & ";" & FECHATEX$(DES%) & ";" & FECHATEX$(HAS%) & ")", "Liquidacion Mensual por Cliente")
99 Command38.Enabled = True
End Sub

Private Sub Command39_Click()
      'HAY QUE AGREGAR EL ENCABEZADO EN CARGALISTA
      'HAY QUE AGREGAR LA CARPETA SKINS EN LA RUTA DEL EJECUTABLE
      Command39.Enabled = False
      PENDIENTES.Show 1
99    Command39.Enabled = True

End Sub

Private Sub Command391_Click()
   Command391.Enabled = False
   'PROCESO PARA PONER TODOS LOS REGISTROS CON CLASIFICACION A
   CONDI$ = "NUME_OT <> ''"
   A$ = "A"
   Call ACTUREGISTRO("CORTRAB", "CLASIFI", CONDI$, A$, RAFE&, "NOMESS")
   Command391.Enabled = True
End Sub


'
Private Sub Command4_Click()
   Call GRADATCOTE
'   Call GRADETCOTE
End Sub

Private Sub Command40_Click()
    
10  OPX% = COMALTER("Informes de No Conformidad:\\Nueva NC \Abrir Existente\Generar Nueva Independiente")
    If OPX% < 1 Or OPX% > 3 Then Exit Sub
    '
    'limpio las variables
    FREP = HOY(HO$)
    CTACTO$ = ""
    ENTRNO$ = ""
    DESCRI$ = ""
    REFERE$ = ""
    CT$ = "CT-"
    '
    'VALIDA OPCIONES
    If OPX% = 1 Then
        CT$ = AJUSTI$(Label7.Caption, 12) ' VERIFICA QUE LA CT ACTIVA ESTE GRABADA
'        CONDI$ = "Nume_Cotel = '" & CT$ & "'"
'        CANTCONSU& = CantRegistros("COTELEF", CONDI$)
'        If CANTCONSU& < 1 Then
'              Call COMUNI("Debe Primero Grabar La Consulta Activa")
'              GoTo 10
'        End If
    
        CONDI$ = "CTNRO = '" & CT$ & "'"  ' VALIDA SI LA CT ACTIVA NO TIENE NO CONFORMIDAD
        EXICT& = CantRegistros&("INOCONF", CONDI$)
        If EXICT& < 1 Then ' SI NO EXISTE TRANSFIERE LOS DATOS AL DISEÑADOR
                 CONDI$ = "NUMERO > 0 "
              NUMENOCONF& = 1 + XVALO(VALOBADA("INOCONF", "MAX(NUMERO)", CONDI$, "NOMESS"))
              Call CreaRegistro("INOCONF", "NUMERO", NUMENOCONF&)
                 CONDI$ = "NUMERO = " & NUMENOCONF&
              Call ACTUREGISTRO("INOCONF", "FeNoCon", CONDI$, CVDAT(HOY(HO$)), RAFE&)
              Call ACTUREGISTRO("INOCONF", "CTNRO", CONDI$, CT$, RAFE&)
              Call ACTUREGISTRO("INOCONF", "Detecto", CONDI$, TRIM$(Left$(Combo1.TEXT, 32)), RAFE&)
              Call ACTUREGISTRO("INOCONF", "Entorno", CONDI$, TRIM$(Text2), RAFE&)
              Call ACTUREGISTRO("INOCONF", "Descripcion", CONDI$, Text22, RAFE&)
              Call ACTUREGISTRO("INOCONF", "Referencia", CONDI$, TRIM$(Text21), RAFE&)
           Else ' SI EXISTE PRESENTA MENSAJE DE AVISO Y MUESTRA LA MISMA
             Call COMUNI("La Presente Consulta yá Tiene Generada No Conformidad")
             NUMENOCONF& = XVALO(VALOBADA("INOCONF", "NUMERO", CONDI$, "NOMESS"))
        End If
        '
        ' Abrir Existente
    ElseIf OPX% = 2 Then
        CONDI$ = "NUMERO > 0 "
        Call CARGALISEL("NOCONFOR", "INOCONF", "Numero/F10.0;FeNoCon/D8;Referencia/A48", CONDI$, "NOMESS")
        Call SELECHO("NOCONFOR/Indice de No Conformidades")
        NUMENOCONF& = XVALO(VALACT1$("NOCONFOR"))
        If NUMENOCONF& < 1 Then GoTo 10
        '
        ' Nueva INDEPENDIENTE
    ElseIf OPX% = 3 Then
        CONDI$ = "NUMERO > 0 "
        NUMENOCONF& = 1 + XVALO(VALOBADA("INOCONF", "MAX(NUMERO)", CONDI$, "NOMESS"))
        Call CreaRegistro("INOCONF", "NUMERO", NUMENOCONF&)
        CONDI$ = "NUMERO = " & NUMENOCONF&
        Call ACTUREGISTRO("INOCONF", "FeNoCon", CONDI$, CVDAT(HOY(HO$)), RAFE&)
    End If
    '
    'PRESENTA FORMULARIO DE DISEÑADOR
    FILTX$ = TRIM$(Str$(NUMENOCONF&))
    Call STDFORM("INFNOCON", FILTX$)
    GoTo 10
    '
End Sub






Private Sub Command47_Click(Index As Integer)
   'AGREGA / QUITA DOCUMENTO ANEXO O POE
   Command47(Index).Enabled = False
    
    Select Case Index
    'BOTON ADD
    Case 0: CARGARUTA
    'BOTON QUITAR
    Case 1
       If List6.ListIndex >= 0 Then
          List6.RemoveItem (List6.ListIndex)
       End If
   End Select
   '
99 Command47(Index).Enabled = True

End Sub
Sub CARGARUTA()
      'PERMITE AGREGAR NUEVOS POE'S
      '
      CD1.FILTER = "(*.*)|*.*|"
      CD1.ShowOpen
      Ruta$ = CD1.FileName
      If TRIM(Ruta$) = "" Then Exit Sub
      '
20    DESCRI$ = TRIM$(InputBox$("Ingrese Texto ", "Guardar Como", DESCRI$))
      If DESCRI$ = "" Then Exit Sub
      If Len(DESCRI$) > 24 Then
        Call COMUNI("Longitud de Descripción \Supera La Cantidad Máxima de Caracteres (24)")
        GoTo 20
      End If
      'AGREGAR INFORMACION
      Call REPLA(TX$, DESCRI$, 1, 24)
      Call REPLA(TX$, Ruta, 32, 512)
      List6.AddItem TX$
End Sub



Private Sub Command5_Click()
   Call GRADATCOTE
'   Call GRADETCOTE
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
'   If NCXI% > 0 Then
'      USDESTI$ = SAFETEXT$(VALOBADA("COTELEF", "Solicitante", "NroClie = " & NCXI% & " ORDER BY FECHEMIS DESC ", "NOMESS"))
'      Call CliCombo(Combo2, USDESTI$)
'      If Combo2.ListCount < 1 Then
'           Combo2.AddItem AJUSTI$(USDESTI$, 64)
'           Combo2.ListIndex = Combo2.ListCount - 1
'         Else
'           Combo2.ListIndex = 0
'      End If
'      '
'      Combo2.Refresh
'   End If
   '
   Label15 = NEXTORTRA$
   HOII% = HOY(HOS$)
   Label14 = HOS$
   Label37 = "A"
   '
   Call SHOWDETHOR
End Sub

Private Sub Command9_Click()
    If Text11.Height > 350 Then
      Text11.Height = 315
      Text11.Width = 525
      Text11.Visible = False
      Text15 = Text11
    End If
    Call GRADATOTRA
    Call COPYSTRU("CORTRAB1", "CORTRABS")
    'Call CARGALISEL("CORTRABS", "CORTRAB", "Nume_Ot/A12;Referencia/A66", "Nume_OT <> '' order by Nume_OT ASC")
End Sub

Private Sub Form_Load()
   SSTab1.Tab = 0
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then
     Call FINAL("")
     Exit Sub
   End If
   '
   Label36 = "Fecha - Hora   Archivo                       Empresa       Usuario           "
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
   'Call VERIPARTE               ' VERIFICA EL PARTE DIARIO DEL DIA ANTERIOR
   '
   'Call TARPEN07.CARLISTAPEN
   'TARPEN07.Show
   '
   Call ABRECONEXION
   Call CREATABLA_ORTRAB
   Call CREATABLA_DOCUREL_CT
   '
   ' RECUPERACION DE DATOS
   Call ACTUREGISTRO("DETOTRA", "RESPEJEC", "RESPEJEC = 'EB'", "EPB", RAFE&, "NOMESS")
   Call ACTUREGISTRO("DETOTRA", "RESPEJEC", "RESPEJEC = 'EP'", "EPB", RAFE&, "NOMESS")
   Call ACTUREGISTRO("DETOTRA", "RESPEJEC", "RESPEJEC = 'EPP'", "EPB", RAFE&, "NOMESS")
   Call ACTUREGISTRO("DETOTRA", "RESPEJEC", "RESPEJEC = ''", "EPB", RAFE&, "NOMESS")
   Call ACTUREGISTRO("DETOTRA", "RESPEJEC", "RESPEJEC = 'FR'", "FG", RAFE&, "NOMESS")
   Call ACTUREGISTRO("DETOTRA", "RESPEJEC", "RESPEJEC = 'MFR'", "FG", RAFE&, "NOMESS")
   Call ACTUREGISTRO("DETOTRA", "RESPEJEC", "RESPEJEC = 'MD'", "OD", RAFE&, "NOMESS")
   Call ACTUREGISTRO("DETOTRA", "RESPEJEC", "RESPEJEC = 'MH'", "OD", RAFE&, "NOMESS")
   MODIFIC_CTE% = 0
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
     If Image2.Visible = True Then
       If TRIM$(Label7) <> "" Then
         NUCOTE$ = TRIM$(Label7)
         CONDI$ = "Nume_Cotel = '" & NUCOTE$ & "'"
         FEXI% = HOY(HOS$)
         FENUE$ = FETEXCOR1$(FEXI%)
         Call ACTUREGISTRO("COTELEF", "Fech_Cierre", CONDI$, FENUE$, RAFE&, "NOMESS")
         Label25 = FECHATEX$(FEXI%)
         Image2.Visible = False
         Image1.Visible = True
         '
       End If
     End If
   '
   Call ACTULISEL
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

Private Sub Label15_Change()
   Call CARDATOTRA
   Call SHOWDETHOR
End Sub

Sub SHOWDETHOR()
   '
   List1.Clear: List1.Refresh
   NUCOTE$ = TRIM$(Label15.Caption)
   If MIGRANDO% > 0 Then
      Call SHOWDETHOR_DAT
      Exit Sub
   End If
   '
   CONDI$ = "Nume_Ot = '" & NUCOTE$ & "'"
   SQLX$ = "SELECT * FROM DETOTRA WITH(NOLOCK) WHERE " + CONDI$ + " ORDER BY fechejec asc, HOraCom ASC, horafin asc"
   '
   SUMEPB = 0: SUMOTR = 0
   Dim DETRA As ADODB.Recordset
   Set DETRA = New ADODB.Recordset
   Set DETRA = FLEXCONN.Execute(SQLX$)
   With DETRA
      Do While Not .EOF
            TTXX$ = Space$(128)
            fex = CVINT(!FechEjec)
          Call REPLA(TTXX$, FECHATEX$(fex), 1, 8)
            REFX$ = TRIM$(!Descrip)
          Call REPLA(TTXX$, REFX$, 11, 55)
            COMI$ = HORATEX$(!HORACOM)
          Call REPLA(TTXX$, COMI$, 67, 5)
            FINI$ = HORATEX$(!HORAFIN)
          Call REPLA(TTXX$, FINI$, 74, 5)
            HORAX$ = FORMATNUM$(!TiemEjec, "F6.2")
          Call REPLA(TTXX$, HORAX$, 79, 6)
            RESPO$ = UCase$(TRIM$(!RespEjec))
          Call REPLA(TTXX$, RESPO$, 87, 4)
          If TRIM$(UCase$(RESPO$)) = "EPB" Or TRIM$(UCase$(RESPO$)) = "EB" Then
              SUMEPB = SUMEPB + Val(HORAX$)
            Else
              SUMOTR = SUMOTR + Val(HORAX$)
          End If
             LUEJ$ = "Lab."
             If UCase$(Left$(SAFETEXT$(!LugEjec), 3)) = "COM" Then LUEJ$ = "Comp"
          Call REPLA(TTXX$, LUEJ$, 93, 4)
          List1.AddItem TTXX$
      .MoveNext
      Loop
   End With
   '
   Text18 = FORMATNUM$(SUMEPB, "F6.2")
   Text17 = FORMATNUM$(SUMOTR, "F6.2")
   Text19 = FORMATNUM$(SUMEPB + SUMOTR, "F6.2")
   '
End Sub

Sub GRADETHOR()
   '
   NUCOTE$ = TRIM$(Label15.Caption)
   If NUCOTE$ <> "" Then
      CONDI$ = "Nume_Ot = '" & NUCOTE$ & "'"
      Call BORRAREGISTROS("DETOTRA", CONDI$, RAFE&, "NOMESS")
      '
      SQLX$ = "SELECT * FROM DETOTRA WHERE " + CONDI$
      Dim DETRA As ADODB.Recordset
      Set DETRA = New ADODB.Recordset
      DETRA.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      With DETRA
         For U& = 1 To ULTREG&("!DETOTRA")
            XX$ = REGLEIDO$("!DETOTRA", U&)
            FEXI% = CVI(Mid$(XX$, 13, 2))
            REFX$ = TRIM$(Mid$(XX$, 15, 50))
            COMIX% = CVI(Mid$(XX$, 65, 2))
            FINIX% = CVI(Mid$(XX$, 67, 2))
            HORAX = CVS(Mid$(XX$, 69, 4))
            RESPO$ = Mid$(XX$, 73, 4)
            LUEJ$ = Mid$(XX$, 81, 4)
            '
            .AddNew
            !nume_ot = NUCOTE$
              FENUE$ = FETEXCOR1$(FEXI%)
            !FechEjec = FENUE$
            !Descrip = REFX$
            !HORACOM = COMIX%
            !HORAFIN = FINIX%
            !TiemEjec = HORAX
            !RespEjec = RESPO$
            !LugEjec = LUEJ$
            .Update
         Next U&
      End With
      DETRA.Close
      Set DETRA = Nothing
   End If
   '
End Sub
'
Sub SHOWDETHOR_DAT()
   List1.Clear
   Call FILTERFILE("DETOTRA", "", 1, Label15)
   SUMEPB = 0: SUMOTR = 0
   For U& = 1 To ULTREG&("!DETOTRA")
     XX$ = REGLEIDO$("!DETOTRA", U&)
     TTXX$ = Space$(128)
       fex = CVI(Mid$(XX$, 13, 2))
     Call REPLA(TTXX$, FECHATEX$(fex), 1, 8)
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
      FENUE$ = FETEXCOR1$(FPOSTI%)
      Call ACTUREGISTRO("TARINTER", "Fech_Prog", "Nume_Tar = " & NTAR&, FENUE$, RAFE&, "NOMESS")
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
   Label38 = ""
   Option3(0).Value = True
   '
   NTAR& = XVALO(Label28)
   If NTAR& > 0 Then
     If MIGRANDO% > 0 Then
          Call CARTARIN_DAT
       Else
          Call CARTARIN
     End If
   End If
   '
End Sub

Sub CARTARIN_DAT()
   '
   NTAR& = XVALO(Label28)
   TTXX$ = FILTERGETRECORD$("TARINTER", 1, MKS$(NTAR&))
   fex = CVI(Mid$(TTXX$, 65, 2))
   Label29 = FECHATEX$(fex)
   fex = CVI(Mid$(TTXX$, 67, 2))
   Text31 = FECHATEX$(fex)
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
   '
End Sub

Sub CARTARIN()
   '
   NTAR& = XVALO(Label28)
   CONDI$ = "Nume_Tar = " & NTAR&
     fex = CVINT(VALOBADA("TARINTER", "Fech_Emis", CONDI$, "NOMESS"))
   Label29 = FECHATEX$(fex)
     fex = CVINT(VALOBADA("TARINTER", "Fech_Cier", CONDI$, "NOMESS"))
   Text31 = FECHATEX$(fex)
     FEPOST = CVINT(VALOBADA("TARINTER", "Fech_Prog", CONDI$, "NOMESS"))
   If FEPOST > HOY(HOS$) Then
      Label34 = FECHATEX$(FEPOST)
   End If
   Text29 = TRIM$(VALOBADA("TARINTER", "Refe_Tar", CONDI$, "NOMESS"))
   Text28 = RTrim(VALOBADA("TARINTER", "Deta_Tar", CONDI$, "NOMESS"))
   Text13 = RTrim(VALOBADA("TARINTER", "Obse_Tar", CONDI$, "NOMESS"))
   Label38 = XVALO(VALOBADA("TARINTER", "Prio_Tar", CONDI$, "NOMESS"))
     IRESPO% = XVALO(VALOBADA("TARINTER", "Itec_Resp", CONDI$, "NOMESS"))
   On Error Resume Next
   Option3(IRESPO%).Value = True
   Option4(IRESPO%).Value = True
   On Error GoTo 0
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
      FENUE$ = FETEXCOR1$(FPOSTI%)
      Call ACTUREGISTRO("TARINTER", "Fech_Prog", "Nume_Tar = " & NTAR&, FENUE$, RAFE&, "NOMESS")
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
      FENUE$ = FETEXCOR1$(FPOSTI%)
      Call ACTUREGISTRO("TARINTER", "Fech_Prog", "Nume_Tar = " & NTAR&, FENUE$, RAFE&, "NOMESS")
      Label34 = FECHATEX$(FPOSTI)
      Call CARLISTAPEN
      Call List3_DblClick
    End If
End Sub

Private Sub Label34_DblClick()
   NTAR& = Val(Label28)
   If NTAR& > 0 Then
      Call ACTUREGISTRO("TARINTER", "Fech_Prog", "Nume_Tar = " & NTAR&, CVDAT(0), RAFE&, "NOMESS")
      Label34 = ""
      Call CARLISTAPEN
    End If
End Sub

Private Sub Label35_DBLClick()
   VANTE$ = TRIM$(UCase$(Label35))
   If VANTE$ = "" Then
        Label35 = "FG"
      ElseIf VANTE$ = "FG" Then
        Label35 = "OD"
      ElseIf VANTE$ = "OD" Then
        Label35 = ""
   End If
   Label35.Refresh
End Sub

Private Sub Label37_DblClick()
   If TRIM$(Label37) = "" Then
      Label37 = "V"
      Text16 = "VERIFICAR"
   ElseIf Label37 = "V" Then
     Label37 = "T"
     Text16 = "TESTEO"
   ElseIf Label37 = "T" Then
     Label37 = "F"
     Text16 = "FALLO"
   ElseIf Label37 = "F" Then
     Label37 = "M"
     Text16 = "MODIFICACION"
   Else
     Label37 = ""
     Text16 = ""
   End If
End Sub

Private Sub Label38_Click()
   Label38 = CStr(1 + XVALO(Label38))
   If Label38 > 9 Then Label38 = 0
End Sub

Private Sub Label7_Change()
   ADJUNARCH.Label3 = ""
   Call CARDATCOTE
   'Call CARDETCOTE
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
   TTXX$ = List1.TEXT
   FECHO$ = TRIM$(Left$(TTXX$, 8))
   RESPX$ = TRIM$(Mid$(TTXX$, 85, 4))
   ITARFE09.Label14 = FECHO$
   ITARFE09.Option3(0).Value = True
   If RESPX$ = "FG" Then
       ITARFE09.Option3(1).Value = True
     ElseIf RESPX$ = "OD" Then
       ITARFE09.Option3(5).Value = True
     ElseIf RESPX$ = "MC" Then
       ITARFE09.Option3(2).Value = True
     ElseIf RESPX$ = "EG" Then
       ITARFE09.Option3(4).Value = True
     ElseIf RESPX$ = "NC" Then
       ITARFE09.Option3(3).Value = True
   End If
   Call ITARFE09.SOLOPROPIO
   ITARFE09.Show 1
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

Private Sub List6_DblClick()
   LINEA$ = List6.List(List6.ListIndex)
   Archi$ = TRIM$(Mid$(LINEA$, 25))
   If (EXISTE%(Archi$)) <> 1 Then
       Call COMUNI("Archivo Inexistente")
       Exit Sub
     Else
        X = ShellExecute(Me.hwnd, "Open", Archi$, "", "", 1)
    End If

End Sub

Private Sub Option3_Click(Index As Integer)
   ' Option4(INDEX).Value = True
End Sub

Private Sub Option4_Click(Index As Integer)
   If MIGRANDO% < 1 Then
      Call CARLISTAPEN
      Call List3_Click
   End If
End Sub

Sub CARLISTAPEN()
   Dim RETEY$(1024)
   ENCARLISTA% = 1
   FILRESPO% = 9: TODOS% = 1
   For KKI% = 0 To 8
       If Option4(KKI%).Value = True Then
          FILRESPO% = KKI%
          TODOS% = 0
       End If
   Next KKI%
   '
   TODOS% = 0
   HOIX$ = FETEXCOR1$(HOY(HOS$))
   CONDI$ = "(Fech_Cier < '30/NOV/1990' OR Fech_Cier Is Null)"
   CONDI$ = CONDI$ & " AND (Fech_Prog <= '" & HOIX$ & "' OR Fech_Prog IS NULL)"
   If FILRESPO% < 9 Then
      CONDI$ = CONDI$ & " AND ITec_Resp = " & FILRESPO%
   End If
   SQLX$ = "SELECT * FROM TARINTER WITH(NOLOCK) WHERE " & CONDI$
   SQLX$ = SQLX$ & " ORDER BY Nume_Tar ASC"
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
   Dim TASEL As ADODB.Recordset
   Set TASEL = New ADODB.Recordset
   Set TASEL = FLEXCONN.Execute(SQLX$)
   '
   With TASEL
     Do While Not .EOF
       NTAR& = !Nume_Tar
       fex = CVINT(!Fech_Emis)
       DETAX$ = RTrim(SAFETEXT$(!Deta_Tar))
       If TRIM$(DETAX$) = "" Then
          DETAX$ = RTrim(SAFETEXT$(!Des_Corta))
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
       OBSEX$ = RTrim$(SAFETEXT$(!Obse_Tar))
       Call FRATEXTO(OBSEX$, 64)
       If CARETEX% > CARETEZ% Then CARETEZ% = CARETEX%
       '
       YY$ = Space$(128 + 64)
       Call REPLA(YY$, FORMATNUM$(NTAR&, "F5.0"), 1, 5)
       Call REPLA(YY$, FECHATEX$(fex), 9, 8)
       Call REPLA(YY$, RETEY$(1), 19, 64)
       Call REPLA(YY$, RETEX$(1), 97, 64)
       '
       If TODOS% > 0 Then
          IRESPO% = XVALO(!ITEC_RESP)
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
     .MoveNext
     Loop
   End With
   TASEL.Close
   Set TASEL = Nothing
   '
   On Error Resume Next
   List3.TopIndex = TP1
   List3.Selected(TP2) = True
   On Error GoTo 0
   ENCARLISTA% = 0
   '
End Sub

Private Sub Option6_Click(Index As Integer)
   If Index <> FORDEN% Then
      FORDEN% = Index
      Call Command19_Click
   End If
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
   If SSTab1.Tab = 1 Then
      FRAME17(7).Visible = True
   Else
      FRAME17(7).Visible = False
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
    MODIFIC_CTE% = 0
    On Error GoTo 0
    '
End Sub

Private Sub Text1_DblClick()
    '
    Call Command1_Click
    '
End Sub

Sub CARDETCOTE_DAT()
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

Sub CARDATCOTE()
   '
   NUCOTE$ = AJUSTI$(Label7.Caption, 12)
   If MIGRANDO% > 0 Then
      Call CARDATCOTE_DAT
      Call CARDETCOTE_DAT
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   CONDI$ = "Nume_Cotel = '" & NUCOTE$ & "'"
   '
29 Text21.TEXT = VALOBADA("COTELEF", "DesCorta", CONDI$, "NOMESS")
   NC% = XVALO(VALOBADA("COTELEF", "NroClie", CONDI$, "NOMESS"))
   If NC% > 0 Then
      Text1.TEXT = TRIM$(Str$(NC%))
      Text2.TEXT = RASOCLI$(NC%)
      Text3.TEXT = TRIM$(DOMICLI$(NC%)) + " - (" + TRIM$(CPOSCLI$(NC%)) + ") " + LOCACLI$(NC%)
      Text4.TEXT = TEL1CLI$(NC%)
      Text5.TEXT = FAXCLI$(NC%)
   End If
   '
   Text12.TEXT = FECHATEX$(CVINT(VALOBADA("COTELEF", "FechEmis", CONDI$, "NOMESS")))
   Text6.TEXT = HORATEX$(XVALO(VALOBADA("COTELEF", "HoraEmis", CONDI$, "NOMESS")))
   '
   WEBUS$ = TRIM$(SAFETEXT$(VALOBADA("COTELEF", "Solicitante", CONDI$, "NOMESS")))
   For KKI% = 1 To Combo1.ListCount
      If UCase$(TRIM$(Left$(Combo1.List(KKI% - 1), 24))) = UCase$(Left$(WEBUS$, 24)) Then
         Combo1.ListIndex = KKI% - 1
         Combo1.Refresh
         GoTo 34
      End If
   Next KKI%
   Combo1.AddItem AJUSTI$(WEBUS$, 64)
   Combo1.ListIndex = Combo1.ListCount - 1
   '
34 AAA$ = AJUSTI$(SAFETEXT$(VALOBADA("COTELEF", "ConCop-1", CONDI$, "NOMESS")), 24)
   AAA$ = AAA$ + AJUSTI$(SAFETEXT$(VALOBADA("COTELEF", "ConCop-2", CONDI$, "NOMESS")), 24)
   AAA$ = AAA$ + AJUSTI$(SAFETEXT$(VALOBADA("COTELEF", "ConCop-3", CONDI$, "NOMESS")), 24)
   Text10 = ""
   CONCO$ = TRIM$(Left$(AAA$, 24))
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
   FECIE% = CVINT(VALOBADA("COTELEF", "Fech_Cierre", CONDI$, "NOMESS"))
   If FECIE% > 8224 Then
      Label25 = FECHATEX$(FECIE%)
      Image2.Visible = False
      Image1.Visible = True
   End If
   '
   TIPOCO% = XVALO(VALOBADA("COTELEF", "TipoCon", CONDI$, "NOMESS"))
   If TIPOCO% > 3 Then TIPOCO% = 0
   Option5(TIPOCO%).Value = True
   '
   Text22 = VALOBADA("COTELEF", "DetaConsu", CONDI$, "NOMESS")
   Text23 = VALOBADA("COTELEF", "DiagInstru", CONDI$, "NOMESS")
   Text7 = VALOBADA("COTELEF", "Observa", CONDI$, "NOMESS")
   MODIFIC_CTE% = 0
   Screen.MousePointer = 1
   '
End Sub

Sub CARDATCOTE_DAT()
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
   RNC$ = RECFILT$("COTELEF1", 1, NUCOTE$)
   If Len(RNC$) >= 4 Then
      RENOTA& = CVS(Mid$(RNC$, U&, 4))
      If RENOTA& > 0 Then
         If RENOTA& <= URN& Then
            TTXX$ = REGLEIDO$("COTELEF1", RENOTA&)
            If Left$(TTXX$, 12) = NUCOTE$ Then
               GoTo 29
            End If
         End If
      End If
   End If
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
     fex = CVI(Mid$(TTXX$, 251, 2))
     Label25 = FECHATEX$(fex)
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
   Screen.MousePointer = 11
   NUCOTE$ = AJUSTI$(Label7.Caption, 12)
   CONDI$ = "Nume_Cotel = '" & NUCOTE$ & "'"
   QTR& = CantRegistros("COTELEF", CONDI$)
   If QTR& < 1 Then
       Call CreaRegistro("COTELEF", "Nume_Cotel", NUCOTE$)
       MODIFIC_CTE% = 1
   End If
      '
      NC% = Val(Text1.TEXT)
      REFE$ = TRIM$(TEL2CLI$(NC%)) + " - " + Text12 + " - " + Text21.TEXT
   Call ACTUREGISTRO("COTELEF", "Referencia", CONDI$, REFE$, RAFE&, "NOMESS")
   Call ACTUREGISTRO("COTELEF", "NroClie", CONDI$, NC%, RAFE&, "NOMESS")
   Call ACTUREGISTRO("COTELEF", "Solicitante", CONDI$, Left$(Combo1.TEXT, 32), RAFE&, "NOMESS")
       FENUE$ = FETEXCOR1$(FECHANUM(Text12.TEXT))
   Call ACTUREGISTRO("COTELEF", "FechEmis", CONDI$, FENUE$, RAFE&, "NOMESS")
   Call ACTUREGISTRO("COTELEF", "HoraEmis", CONDI$, Int(HORANUM(Text6)), RAFE&, "NOMESS")
   Call ACTUREGISTRO("COTELEF", "DesCorta", CONDI$, Left$(Text21, 128), RAFE&, "NOMESS")
      RETEX$(1) = "": RETEX$(2) = "": RETEX$(3) = ""
      Call FRATEXTO(Text10.TEXT, 24)
   Call ACTUREGISTRO("COTELEF", "ConCop-1", CONDI$, Left$(RETEX$(1), 32), RAFE&, "NOMESS")
   Call ACTUREGISTRO("COTELEF", "ConCop-2", CONDI$, Left$(RETEX$(2), 32), RAFE&, "NOMESS")
   Call ACTUREGISTRO("COTELEF", "ConCop-3", CONDI$, Left$(RETEX$(3), 32), RAFE&, "NOMESS")
       FENUE$ = FETEXCOR1$(FECHANUM(Label25))
   Call ACTUREGISTRO("COTELEF", "Fech_Cierre", CONDI$, FENUE$, RAFE&, "NOMESS")
      TIPOCO% = 0
      If Option5(1).Value = True Then TIPOCO% = 1
      If Option5(2).Value = True Then TIPOCO% = 2
      If Option5(3).Value = True Then TIPOCO% = 3
   Call ACTUREGISTRO("COTELEF", "TipoCon", CONDI$, TIPOCO%, RAFE&, "NOMESS")
   Call ACTUREGISTRO("COTELEF", "DetaConsu", CONDI$, Text22.TEXT, RAFE&, "NOMESS")
   Call ACTUREGISTRO("COTELEF", "DiagInstru", CONDI$, Text23.TEXT, RAFE&, "NOMESS")
    Call ACTUREGISTRO("COTELEF", "Observa", CONDI$, Text7.TEXT, RAFE&, "NOMESS")
   '
   If MODIFIC_CTE% > 0 Then Call ACTULISEL 'MODIFICA SOLO EN EL CASO QUE SE CAMBIE LA REFERENCIA
   MODIFIC_CTE% = 0
   Screen.MousePointer = 1
   '
End Sub

Sub ACTULISEL()
   '
'   Call COPYSTRU("COTELEF1", "COTELEF")
'   Call CIERRARCH("COTELEF")
'   CONDI$ = "Nume_Cotel <> '' ORDER BY FECHEMIS ASC, NUME_COTEL ASC"
'   Call CARGALISEL("COTELEF", "COTELEF", "Nume_Cotel/A12;Referencia/A66;FECHEMIS/D0", CONDI$, "DISTINCT")
   '
   CONDI$ = "Nume_Cotel <> '' "
   CONDI$ = CONDI$ + " AND (Fech_Cierre <= '30/NOV/1999' OR Fech_Cierre IS NULL) ORDER BY FECHEMIS ASC, NUME_COTEL ASC"
   Call CARGALISEL("!COTELEF", "COTELEF", "Nume_Cotel/A12;FECHEMIS/D0;Referencia/A66", CONDI$, "DISTINCT")
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or KeyCode = vbKeyEscape Then
    If Text11.Height > 350 Then
      Text11.Height = 315
      Text11.Width = 525
      Text11.Visible = False
      Text15 = Text11
      On Error Resume Next
      Text15.SetFocus
      On Error GoTo 0
    Else
     Text11 = Text15
     Text11.Height = 8000
     Text11.Width = 10000
     Text11.Visible = False
     On Error Resume Next
     Text11.SetFocus
     On Error GoTo 0
   End If
  End If
End Sub

Private Sub Text12_Change()
  MODIFIC_CTE% = 1
End Sub

'
Private Sub Text13_DblClick()
   Text13 = ULOCASE$(Text13)
End Sub

Private Sub Text14_DblClick()
   Text14 = ULOCASE$(Text14)
End Sub

Private Sub Text15_DblClick()
   '
   Text15 = ULOCASE$(Text15)
End Sub

Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    If Text11.Height > 350 Then
      Text11.Height = 315
      Text11.Width = 525
      Text11.Visible = False
      Text15 = Text11
      On Error Resume Next
      Text15.SetFocus
      On Error GoTo 0
    Else
     Text11 = Text15
     Text11.Height = 8000
     Text11.Width = 10000
     Text11.Visible = True
     On Error Resume Next
     Text11.SetFocus
     On Error GoTo 0
   End If
 End If
End Sub

Private Sub Text16_DblClick()
   Text16 = ULOCASE$(Text16)
End Sub

Private Sub Text2_Change()
  MODIFIC_CTE% = 1
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

Private Sub Text21_Change()
  MODIFIC_CTE% = 1
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
   NUCOTE$ = TRIM$(Label15.Caption)
   If MIGRANDO% > 0 Then
      Call CARDATOTRA_DAT
      Exit Sub
   End If
   '
   CONDI$ = "Nume_Ot = '" & NUCOTE$ & "'"
   '
   Screen.MousePointer = 11
29 Text9.TEXT = VALOBADA("CORTRAB", "DesCorta", CONDI$)
     NC% = XVALO(VALOBADA("CORTRAB", "NroClie", CONDI$))
     ' esto es para lar nuevas ordenes de trabajo
     If NC% > 0 Then
        Label24 = TRIM$(Str$(NC%))
        Label8 = RASOCLI$(NC%)
     End If
     '
'   If NC% > 0 Then
'      USDESTI$ = SAFETEXT$(VALOBADA("COTELEF", "Solicitante", "NroClie = " & NC% & " ORDER BY FECHEMIS DESC ", "NOMESS"))
'      Call CliCombo(Combo2, USDESTI$)
'      If Combo2.ListCount < 1 Then
'           Combo2.AddItem AJUSTI$(USDESTI$, 64)
'           Combo2.ListIndex = Combo2.ListCount - 1
'         Else
'           Combo2.ListIndex = 0
'      End If
'      '
'      Combo2.Refresh
'   End If
     
   Text8.TEXT = TRIM$(SAFETEXT$(VALOBADA("CORTRAB", "Solicitante", CONDI$)))
'   WEBUS$ = TRIM$(Text8.TEXT)
'   For KKI% = 1 To Combo2.ListCount
'      If UCase$(TRIM$(Left$(Combo2.List(KKI% - 1), 24))) = UCase$(Left$(WEBUS$, 24)) Then
'         Combo2.ListIndex = KKI% - 1
'         Combo2.Refresh
'         GoTo 34
'      End If
'   Next KKI%
'   Combo2.AddItem AJUSTI$(WEBUS$, 64)
'   Combo2.ListIndex = Combo2.ListCount - 1
   
34 Label14 = FECHATEX$(CVINT(VALOBADA("CORTRAB", "FechEmis", CONDI$)))
     On Error Resume Next
     moti% = XVALO(VALOBADA("CORTRAB", "Motivo", CONDI$))
   Option1(moti%).Value = True
     On Error GoTo 0
     On Error Resume Next
     IMPU% = XVALO(VALOBADA("CORTRAB", "Imputa", CONDI$))
   Option2(IMPU%).Value = True
     On Error GoTo 0
   Text26 = FECHATEX$(CVINT(VALOBADA("CORTRAB", "Fech_Enviada", CONDI$)))
   Text27 = FECHATEX$(CVINT(VALOBADA("CORTRAB", "Fech_Instala", CONDI$)))
   Text20 = FECHATEX$(CVINT(VALOBADA("CORTRAB", "Fech_Cierre", CONDI$)))
   Label35 = UCase$(TRIM$(VALOBADA("CORTRAB", "TecRespo", CONDI$)))
   '
   Text16 = RTrim(VALOBADA("CORTRAB", "Observa", CONDI$))
   Text14 = RTrim(VALOBADA("CORTRAB", "Descripcion", CONDI$))
   Text15 = RTrim(VALOBADA("CORTRAB", "Procedi", CONDI$))
   Label37 = TRIM(VALOBADA("CORTRAB", "Clasifi", CONDI$))
   '
    List6.Clear
    SQLX$ = "SELECT * FROM DOCUREL_CT WITH(NOLOCK) WHERE CORTRAB = '" & NUCOTE$ & "'"
    Dim DOCTMP As ADODB.Recordset
    Set DOCTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With DOCTMP
      Do While Not .EOF
         DESCRI$ = SAFETEXT$(!Descripcion)
         If DESCRI$ = "" Then GoTo 80
         RUTAX$ = SAFETEXT$(!Ruta$)
         Call REPLA(TX$, DESCRI$, 1, 24)
         Call REPLA(TX$, RUTAX$, 32, 512)
         List6.AddItem TX$
80      .MoveNext
      Loop
    End With

   Screen.MousePointer = 1
   '
End Sub

Sub GRADATOTRA()
   '
   Screen.MousePointer = 11
   NUCOTE$ = TRIM$(Label15.Caption)
   CONDI$ = "Nume_Ot = '" & NUCOTE$ & "'"
   QTR& = CantRegistros("CORTRAB", CONDI$)
   If QTR& < 1 Then
       'Call CreaRegistro("CORTRAB", "Nume_OT", NUCOTE$)
       FLEXCONN.Execute "insert into cortrab (nume_ot) values ('" & NUCOTE$ & "')"
   End If
   '
   If TRIM$(Text9) = "" Then Text9 = Text14
   NC% = Val(Label24)
   REFE$ = TRIM$(TEL2CLI$(NC%)) + " - " + Text9
   If FECHANUM(Text20) > 0 Then
        REFE$ = "(c) " + REFE$
     ElseIf FECHANUM(Text27) > 0 Then
        REFE$ = "(i) " + REFE$
     ElseIf FECHANUM(Text26) > 0 Then
        REFE$ = "(e) " + REFE$
   End If
   '
   moti% = 1
   On Error Resume Next
   For KKI% = 1 To 9
     If KKI% = 4 Then KKI% = 9
     If Option1(KKI%).Value = True Then moti% = KKI%
   Next KKI%
   'Call ACTUREGISTRO("CORTRAB", "Motivo", CONDI$, moti%, RAFE&, "NOMESS")
   IMPU% = 1
   On Error Resume Next
   For KKI% = 1 To 9
     If KKI% = 4 Then KKI% = 9
     If Option2(KKI%).Value = True Then IMPU% = KKI%
   Next KKI%
   '
   SQLX$ = "select * from cortrab where nume_ot='" & NUCOTE$ & "'"
   Dim RST As ADODB.Recordset
   Set RST = New ADODB.Recordset
   RST.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With RST
        If Not (.EOF And .BOF) Then
            !REFERENCIA = REFE$
            !NROCLIE = Val(Label24)
            !solicitante = Text8.TEXT
            !FECHEMIS = CVDAT(FECHANUM(Label14))
            !descorta = Text9.TEXT
            !motivo = moti%
            !imputa = IMPU%
            !tecrespo = TRIM$(UCase$(Label35))
            !fech_enviada = CVDAT(FECHANUM(Text26))
            !fech_instala = CVDAT(FECHANUM(Text27))
            !fech_cierre = CVDAT(FECHANUM(Text20))
            !Observa = Text16.TEXT
            !Descripcion = Text14.TEXT
            !procedi = Text15.TEXT
            !clasifi = TRIM$(Label37)
            .Update
        End If
   End With
   '
   RST.Close
   Set RST = Nothing
   
   'MARCA DE BORRADO
   Call ACTUREGISTRO("DOCUREL_CT", "MARBORRA", "CORTRAB = '" & NUCOTE$ & "'", 1, REG&)

   'RECORRO EL LIST PARA GRABAR
   For I& = 1 To List6.ListCount
      Z$ = List6.List(I& - 1)
      DESCRI$ = TRIM$(Mid$(Z, 1, 24))
      RUTAX$ = TRIM$(Mid$(Z, 32, 512))
10    RNDX$ = RANDSTRING$(8)
      If CantRegistros("DOCUREL_CT", "DESCRIPCION = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10

      Call CreaRegistro("DOCUREL_CT", "DESCRIPCION", RNDX$, "NOMESS")
      Call ACTUREGISTRO("DOCUREL_CT", "CORTRAB", "DESCRIPCION = '" & RNDX$ & "'", NUCOTE$, REG&)
      Call ACTUREGISTRO("DOCUREL_CT", "MARBORRA", "DESCRIPCION = '" & RNDX$ & "'", 0, REG&)
      Call ACTUREGISTRO("DOCUREL_CT", "RUTA", "DESCRIPCION = '" & RNDX$ & "'", RUTAX$, REG&)
      Call ACTUREGISTRO("DOCUREL_CT", "DESCRIPCION", "DESCRIPCION = '" & RNDX$ & "'", DESCRI$, REG&)
   Next I&
   Call BORRAREGISTROS("DOCUREL_CT", "MARBORRA = 1 AND CORTRAB = '" & NUCOTE$ & "'", REGAF&)

'   Call ACTUREGISTRO("CORTRAB", "Referencia", CONDI$, REFE$, RAFE&, "NOMESS")
'   Call ACTUREGISTRO("CORTRAB", "NroClie", CONDI$, Val(Label24), RAFE&, "NOMESS")
'   Call ACTUREGISTRO("CORTRAB", "Solicitante", CONDI$, Text8.TEXT, RAFE&, "NOMESS")
'   Call ACTUREGISTRO("CORTRAB", "FechEmis", CONDI$, FETEXCOR1$(Int(FECHANUM(Label14))), RAFE&, "NOMESS")
'   Call ACTUREGISTRO("CORTRAB", "DesCorta", CONDI$, Text9, RAFE&, "NOMESS")
'     MOTI% = 1
'     On Error Resume Next
'     For KKI% = 1 To 9
'       If KKI% = 4 Then KKI% = 9
'       If Option1(KKI%).Value = True Then MOTI% = KKI%
'     Next KKI%
'   Call ACTUREGISTRO("CORTRAB", "Motivo", CONDI$, MOTI%, RAFE&, "NOMESS")
'     IMPU% = 1
'     On Error Resume Next
'     For KKI% = 1 To 9
'       If KKI% = 4 Then KKI% = 9
'       If Option2(KKI%).Value = True Then IMPU% = KKI%
'     Next KKI%
'   Call ACTUREGISTRO("CORTRAB", "Imputa", CONDI$, IMPU%, RAFE&, "NOMESS")
'   Call ACTUREGISTRO("CORTRAB", "TecRespo", CONDI$, TRIM$(UCase$(Label35)), RAFE&, "NOMESS")
'   '
'   Call ACTUREGISTRO("CORTRAB", "Fech_Enviada", CONDI$, FETEXCOR1$(Int(FECHANUM(Text26))), RAFE&, "NOMESS")
'   Call ACTUREGISTRO("CORTRAB", "Fech_Instala", CONDI$, FETEXCOR1$(Int(FECHANUM(Text27))), RAFE&, "NOMESS")
'   Call ACTUREGISTRO("CORTRAB", "Fech_Cierre", CONDI$, FETEXCOR1$(Int(FECHANUM(Text20))), RAFE&, "NOMESS")
'   '
'   Call ACTUREGISTRO("CORTRAB", "Observa", CONDI$, Text16, RAFE&, "NOMESS")
'   Call ACTUREGISTRO("CORTRAB", "Descripcion", CONDI$, Text14, RAFE&, "NOMESS")
'   Call ACTUREGISTRO("CORTRAB", "Procedi", CONDI$, Text15, RAFE&, "NOMESS")
'   Call ACTUREGISTRO("CORTRAB", "Clasifi", CONDI$, TRIM$(Label37), RAFE&, "NOMESS")
   '
   Screen.MousePointer = 1
   '
End Sub

Private Sub Timer2_Timer()
   Static SVLC$
   If HOY(HOS$) <> FECHANUM("30/03/11") Then
     If SVLC$ = "" Then
       SVLC$ = SERIVOL$("C:")
     End If
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
       Call TARPEN07.CARLISTAPEN
       If TARPEN07.Visible <> True Then
            TARPEN07.Show
          Else
            TARPEN07.WindowState = 0
            TARPEN07.ZOrder 0
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
             'Call CARDATCOTE
             'Call CARDETCOTE
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
             'Call CARDATCOTE
             'Call CARDETCOTE
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
   Openforms = DoEvents
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
     Openforms = DoEvents
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

Sub CREATABLA_ORTRAB()
   '
   Call ABRECONEXION
   '
   Call CREACAMPO("ADMIFLEX", "COTELEF", "Nume_Cotel", 10, 16, 1)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "Referencia", 10, 128, 0)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "NroClie", 4, 0, 1)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "Solicitante", 10, 32, 0)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "ConCop-1", 10, 32, 0)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "ConCop-2", 10, 32, 0)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "ConCop-3", 10, 32, 0)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "FechEmis", 8, 0, 1)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "HoraEmis", 3, 0)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "Fech_Cierre", 8, 0)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "Status", 3, 0, 1)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "DesCorta", 10, 128, 0)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "TipoCon", 3, 0, 1)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "DetaConsu", 12, 0)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "TecRespo", 10, 16, 1)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "DiagInstru", 12, 0)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "FeHor_Respu", 8, 0)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "Fech_Aproba", 8, 0)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "Usua_Aproba", 10, 32)
   Call CREACAMPO("ADMIFLEX", "COTELEF", "Observa", 12, 0)
   '
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Nume_OT", 10, 16, 1)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "DesCorta", 10, 128, 0)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Referencia", 10, 128, 0)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Descripcion", 12, 0)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "NroClie", 4, 0, 1)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Solicitante", 10, 32, 0)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "FechEmis", 8, 0)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "FeSolEnt", 8, 0)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "TecRespo", 10, 16, 1)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Motivo", 3, 0, 1)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Imputa", 3, 0, 1)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Status", 3, 0, 1)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Fech_Enviada", 8, 0)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Fech_Instala", 8, 0)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Fech_Aproba", 8, 0)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Usua_Aproba", 10, 32)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Fech_Cierre", 8, 0)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Observa", 12, 0)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Procedi", 12, 0)
   Call CREACAMPO("ADMIFLEX", "CORTRAB", "Clasifi", 10, 1)
   '
   Call CREACAMPO("ADMIFLEX", "DETOTRA", "Nume_OT", 10, 16, 1)
   Call CREACAMPO("ADMIFLEX", "DETOTRA", "FechEjec", 8, 0)
   Call CREACAMPO("ADMIFLEX", "DETOTRA", "Descrip", 12, 0)
   Call CREACAMPO("ADMIFLEX", "DETOTRA", "HoraCom", 3, 0, 0)
   Call CREACAMPO("ADMIFLEX", "DETOTRA", "HoraFin", 3, 0, 0)
   Call CREACAMPO("ADMIFLEX", "DETOTRA", "TiemEjec", 6, 4)
   Call CREACAMPO("ADMIFLEX", "DETOTRA", "RespEjec", 10, 16, 1)
   Call CREACAMPO("ADMIFLEX", "DETOTRA", "LugEjec", 10, 16, 1)
   Call CREACAMPO("ADMIFLEX", "DETOTRA", "Concepto", 10, 16, 1)
   Call CREACAMPO("ADMIFLEX", "DETOTRA", "TarifHor", 6, 4)
   Call CREACAMPO("ADMIFLEX", "DETOTRA", "Importe", 7, 8)
   Call CREACAMPO("ADMIFLEX", "DETOTRA", "StatActi", 3, 0, 1)
   '
   Call CREACAMPO("ADMIFLEX", "TARINTER", "Nume_Tar", 4, 4, 1)
   Call CREACAMPO("ADMIFLEX", "TARINTER", "Des_Corta", 10, 128, 0)
   Call CREACAMPO("ADMIFLEX", "TARINTER", "Prio_Tar", 3, 0, 1)
   Call CREACAMPO("ADMIFLEX", "TARINTER", "Fech_Emis", 8, 0)
   Call CREACAMPO("ADMIFLEX", "TARINTER", "Fech_Cier", 8, 0)
   Call CREACAMPO("ADMIFLEX", "TARINTER", "ITec_Resp", 3, 0, 1)
   Call CREACAMPO("ADMIFLEX", "TARINTER", "Fech_Prog", 8, 0)
   Call CREACAMPO("ADMIFLEX", "TARINTER", "Refe_Tar", 10, 128)
   Call CREACAMPO("ADMIFLEX", "TARINTER", "Deta_Tar", 12, 0)
   Call CREACAMPO("ADMIFLEX", "TARINTER", "Obse_Tar", 12, 0)
   '
   Call CREACAMPO("ADMIFLEX", "INOCONF", "Numero", 4, 0, 1)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "CTNro", 10, 16)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "FeNoCon", 8, 0)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "Detecto", 10, 64)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "Entorno", 10, 64)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "Descripcion", 12, 0)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "FechDesc", 8, 0)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "Causa", 12, 0)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "Modulo", 10, 64)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "Formulario", 10, 64)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "Controlador", 10, 64)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "Otros", 10, 64)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "FechCausa", 8, 0)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "AcContingente", 12, 0)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "FechAcConti", 8, 0)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "Conforme", 10, 2)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "AcCorrectiva", 12, 0)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "VerFutura", 10, 2)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "Responsable", 10, 64)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "FechProg", 8, 0)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "OTrabajo", 10, 16)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "FechCierre", 8, 0)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "Referencia", 10, 64)
   Call CREACAMPO("ADMIFLEX", "INOCONF", "Status", 10, 16)
   
End Sub

Sub CARDATOTRA_DAT()
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
     moti% = Asc(Mid$(TTXX$, 109, 1))
   Option1(moti%).Value = True
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

Function NEXTCOTELNU$()
   '
   If COTERMI$ = "" Then
      ITERMI$ = IDENTER$
10    COTERMI$ = CONTROL$(ITERMI$, "COTERMI")
      If Len(COTERMI$) <> 2 Then
        OPCOTEL.Show 1
        GoTo 10
      End If
   End If
   '
   Call CREATABLA_ORTRAB
   OTX$ = "CT-______-" & COTERMI$
   CAMPOSE$ = "MAX(CAST(SUBSTRING(NUME_COTEL,4,6) AS Integer))"
   CONDI$ = "NUME_COTEL LIKE '" & OTX$ & "'"
   NUCONSU& = 1 + XVALO(VALOBADA("COTELEF", CAMPOSE$, CONDI$))
   If NUCONSU& < 100000 Then
      NUCONSU& = 100001
   End If
   '
20 NEXTNU$ = TRIM$(Str$(NUCONSU&))
   While Len(NEXTNU$) < 6: NEXTNU$ = "0" + NEXTNU$: Wend
   NEXTCOTELNU$ = "CT-" + NEXTNU$ + "-" + COTERMI$
   '
End Function

Function NEXTORTRA$()
   '
   fex = HOY(HOS$)
   COTERMJ$ = Mid$(FECHATEX$(fex), 7, 2)
   OTX$ = "OT-" + COTERMJ$ + "-"
   '
   NUCONSU& = 1 + XVALO(VALOBADA("CORTRAB", "MAX(CAST(SUBSTRING(NUME_OT,7,6) AS Integer))", "NUME_OT LIKE '" & OTX$ & "%'"))
   '
20 NEXTNU$ = TRIM$(Str$(NUCONSU&))
   While Len(NEXTNU$) < 4: NEXTNU$ = "0" + NEXTNU$: Wend
   NEXTORTRA$ = OTX$ + NEXTNU$
   '
End Function

Sub PRICOTEL(NUCOTE$)
   '
   Static CTXX%, RUTARGRA$
   Dim TENCABE$(3)
   '
   If CTXX% = 0 Then
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
     CTXX% = 1
   End If
   '
   PRIFORM$ = TRIM$(CONTROL$("", "FORCOTEL"))
   CAIMAGEN% = 0
   '
   CODPARAM$(1) = "NUME-COM"
   CODPARAM$(2) = "FECH-EMI"
   CODPARAM$(3) = "TIE-PREP"
   CODPARAM$(4) = "RASO-CLI"
   CODPARAM$(5) = "TELE-CLI"
   CODPARAM$(6) = "DOMI-CLI"
   CODPARAM$(7) = "MEDIDA-1"
   CODPARAM$(8) = "MEDIDA-2"
   CODPARAM$(9) = "MEDIDA-3"
   CODPARAM$(10) = "REF-MAT1"
   CODPARAM$(11) = "FECH-LAN"
   CAPARDOC% = 11
   '
   If PRIFORM$ <> "" Then
     If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
       Screen.MousePointer = 11
29     DOCPARAM$(10) = TRIM$(Text21)   '   Mid$(TTXX$, 111, 64))
       DOCPARAM$(1) = NUCOTE$
       NC% = XVALO(Text1)
       If NC% > 0 Then
          DOCPARAM$(4) = RASOCLI$(NC%)
          DOCPARAM$(5) = FAXCLI$(NC%)
       End If
       DOCPARAM$(6) = TRIM$(Left$(Combo1.TEXT, 24))
       DOCPARAM$(2) = Text12
       DOCPARAM$(3) = Text6
       DOCPARAM$(7) = TRIM$(Left$(Text10, 24))
       DOCPARAM$(8) = ""   ' TRIM$(Mid$(TTXX$, 203, 24))
       DOCPARAM$(9) = ""   ' TRIM$(Mid$(TTXX$, 227, 24))
       '
       HOII% = HOY(HOS$)
       DOCPARAM$(11) = HOS$ + " - " + Left$(Time$, 5) + " hs"
       '
       For KK1% = 0 To 32
         For KK2% = 0 To 1024
           TETABU1$(KK1%, KK2%) = ""
           TETABU2$(KK1%, KK2%) = ""
         Next KK2%
       Next KK1%
       '
       CODTABU1$(1) = "REF-MAT3"
       CACOLTAB1% = 1
       CAITAB1% = 0
       '
       CODTABU2$(1) = "REF-MAT2"
       CACOLTAB2% = 1
       CAITAB2% = 0
       '
       TTXX$ = Text22
       Call FRATEXTO(TTXX$, 72)
       For KKII% = 1 To CARETEX%
          CAITAB2% = CAITAB2% + 1
          TETABU2$(1, CAITAB2%) = RETEX$(KKII%)
       Next KKII%
       '
       TTXX$ = Text23
       Call FRATEXTO(TTXX$, 72)
       For KKII% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = RETEX$(KKII%)
       Next KKII%
       '
       DESTIDOC% = 0
       NUMEDOC$ = TRIM$(NUCOTE$)
       TIPODOC$ = "Consulta Telefonica"
       Call PRINTDOC("@" + PRIFORM$)
     End If
   End If
   '
End Sub

Sub PRIORTRA(NUCOTE$)
   '
   Static CTXX%, RUTARGRA$
   Dim TENCABE$(3)
   '
   If CTXX% = 0 Then
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
     CTXX% = 1
   End If
   '
   PRIFORM$ = TRIM$(CONTROL$("", "FORTRABA"))
   CAIMAGEN% = 0
   '
   CODPARAM$(1) = "REF-MAT1"
   CODPARAM$(2) = "FECH-EMI"
   CODPARAM$(3) = "DOCUORIG"
   CODPARAM$(4) = "RASO-CLI"
   CODPARAM$(5) = "TALLE"
   CODPARAM$(6) = "COLOR"
   CODPARAM$(7) = "FECH-LAN"
   CODPARAM$(8) = "FECH-REC"
   CODPARAM$(9) = "FECH-VEN"
   CODPARAM$(10) = "OBS-OPER"
   CODPARAM$(11) = "REF-MAT4"
   CODPARAM$(12) = "REF-MAT5"
   CODPARAM$(13) = "REF-MAT6"
   
   CAPARDOC% = 13
   '
   If PRIFORM$ <> "" Then
     If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
       CONDI$ = "Nume_OT = '" & NUCOTE$ & "'"
       If CantRegistros("CORTRAB", CONDI$) > 0 Then
         Screen.MousePointer = 11
         DOCPARAM$(1) = Text9
         DOCPARAM$(2) = Label14
         DOCPARAM$(3) = Text8
         DOCPARAM$(4) = Label8
           moti$ = "REPARACION"
           If Option1(2).Value = True Then moti$ = "MODIFICACION"
           If Option1(3).Value = True Then moti$ = "NUEVO DESARR."
           If Option1(9).Value = True Then moti$ = "VARIOS"
         DOCPARAM$(5) = moti$
           moti$ = "ABONO"
           If Option2(2).Value = True Then moti$ = "GARANTIA"
           If Option2(3).Value = True Then moti$ = "FACTURACION"
           If Option2(9).Value = True Then moti$ = "VARIOS"
         DOCPARAM$(6) = moti$
         DOCPARAM$(7) = Text26
         DOCPARAM$(8) = Text27
         DOCPARAM$(9) = Text20
         DOCPARAM$(10) = Text16
         '
         DOCPARAM$(11) = Text18
         DOCPARAM$(12) = Text17
         DOCPARAM$(13) = Text19
         '
         CODTABU1$(1) = "REF-MAT2"
         CODTABU1$(2) = "REF-MAT3"
         CODTABU1$(3) = "DES-MAT"
         CODTABU1$(4) = "DES-OPER"
         CACOLTAB1% = 4
         CAITAB1% = 0
         CAITAB2% = 0
         '
         For KK1% = 0 To 32
           For KK2% = 0 To 1024
             TETABU1$(KK1%, KK2%) = ""
             TETABU2$(KK1%, KK2%) = ""
           Next KK2%
         Next KK1%
         '
         For U& = 1 To List1.ListCount
           TTXX$ = TRIM$(Left$(List1.List(U& - 1), 62))
           TTYY$ = Mid$(List1.List(U& - 1), 63)
           If Left$(TTYY$, 1) = " " Then
             Mid$(TTYY$, 1, 1) = "_"
           End If
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = TTXX$
           TETABU1$(2, CAITAB1%) = TTYY$
         Next U&
         '
         Call FRATEXTO(Text15, 90)
         For KU% = 1 To CARETEX%
           TETABU1$(3, KU%) = RETEX$(KU%)
           If KU% > CAITAB1% Then CAITAB1% = KU%
         Next KU%
         '
         Call FRATEXTO(TRIM$(Text14), 41)
         If CARETEX% > 6 Then CARETEX% = 6
         For KU% = 1 To CARETEX%
           TETABU1$(4, KU%) = RETEX$(KU%)
           If KU% > CAITAB1% Then CAITAB1% = KU%
         Next KU%
         '
         DESTIDOC% = 0
         TIPODOC$ = "Orden de Trabajo"
         NUMEDOC$ = TRIM$(Label15)
         Call PRINTDOC("@" + PRIFORM$)
         '
       End If
     End If
   End If
   '
End Sub

Sub MIGRACLIEN()
  'Picture1.ToolTipText = "Depuración - Clientes"
  '
  Call APERBASDAT("CLIEN")
  On Error Resume Next
   FLEXCONN.Execute (FILTSQL$("DROP TABLE DEUDOR12"))
  On Error GoTo 0
  '
  Call APERBASDAT("CLIEN")
  Call GENERBASE("BADACLIEN")
  '
  Dim DEUDORRR As ADODB.Recordset
  Set DEUDORRR = New ADODB.Recordset
  SQLX$ = "SELECT * FROM DEUDOR12"
  DEUDORRR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  'Picture1.ToolTipText = "Conversión de Clientes"
  '
  'LECTURA
  UR& = ULTREG&("DEUDOR1")
  For U& = 1 To ULTREG&("DEUDOR1")
    '
    X1$ = REGLEIDO$("DEUDOR1", U&)
      NCI% = CVI(Mid$(X1$, 1, 2))
    X2$ = REGLEIDO$("DEUDOR2", U&)
    X3$ = FILTERGETRECORD$("BADACLIE", 1, MKI$(NCI%))
    'LECTURA
    RASO$ = Mid$(X3$, 3, 62)
    FANTA$ = Mid$(X3$, 65, 32)
    DOMI$ = Mid$(X3$, 97, 64)
    CPOSTX$ = Mid$(X3$, 161, 16)
    CPOST% = 0: If XVALO(CPOSTX$) < 32767 Then CPOST% = XVALO(CPOSTX$)
    LOCA$ = TRIM$(Mid$(X3$, 177, 48))
    PAISOK$ = Mid$(X3$, 225, 32)
    TELE$ = Mid$(X3$, 257, 24)
    FAXX$ = TRIM$(Mid$(X3$, 281, 24))
    CONTACTX$ = TRIM$(Mid$(X3, 385, 48))
    MAILPRO$ = TRIM$(Mid$(X3$, 305, 48))
    HORARIX$ = TRIM$(Mid$(X3$, 449, 32))
    ZONCLI$ = Mid$(X3$, 481, 12)
    CUIT$ = TRIM$(Mid$(X2$, 75, 15))
      If CUIT$ <> "" Then VCC% = VALICUIT%(CUIT$)
    GruCli% = CVI(Mid$(X2$, 125, 2))
    '
    NCITXT$ = CStr(NCI%)
    Do Until Len(NCITXT$) = 5
      NCITXT$ = "0" + NCITXT$
    Loop
    CODICUEN$ = "CL-" & NCITXT$
    '
    TELEF2$ = Mid$(X2$, 65, 10)
    PIVA% = Asc(Mid$(X2$, 104, 1))
    CONDI% = Asc(Mid$(X2$, 118, 1))
    PROVI$ = TRIM$(Mid$(X2$, 117, 1))
    LIPRE% = Asc(Mid$(X2$, 119, 1))
    COVEN% = Asc(Mid$(X2$, 120, 1))
    LICRE = CVS(Mid$(X2$, 121, 4))
    INGBRU$ = Mid$(X2$, 90, 13)
    CAIBRU$ = Mid$(X2$, 103, 1)
    CUEMA% = CVI(Mid$(X2$, 127, 2))
    GRUDIV% = 0
    '
    DOMIENT1$ = ""
    LOCAENT1$ = ""
    TRANSPOR$ = ""
    NTRANX% = 0
    HORARIENT$ = ""
    If NCI% > 0 Then
      XX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(NCI%))
      DOMIENT1$ = TRIM$(Mid$(XX$, 3, 32))
      LOCAENT1$ = TRIM$(Mid$(XX$, 35, 32))
      TRANSPOR$ = TRIM$(Mid$(XX$, 67, 32))
      NTRANX% = CVI(Mid$(XX$, 99, 2))
         If NTRANX% = 8224 Then NTRANX% = 0
      HORARIENT$ = TRIM$(Mid$(XX$, 101, 24))
    End If
    '
    DESSTAN$ = ""
    XXX$ = FILTERGETRECORD$("DESCUCLI", 1, MKI$(NCI%))
    DESSTAN$ = Mid$(XXX$, 3, 14)
    '
    ' anotaciones
    RNC$ = RECFILT$("NOTACLI", 1, MKI$(NCI%))
    URN& = ULTREG&("NOTACLI")
    '
    NOTAC$ = ""
    For UKL& = 1 To Len(RNC$) Step 4
       RENOTA1& = CVS(Mid$(RNC$, UKL&, 4))
       If RENOTA1& > 0 Then
          TTYY$ = TRIM$(Mid$(REGLEIDO$("NOTACLI", RENOTA1&), 3))
          If NOTAC$ <> "" Then
             NOTAC$ = NOTAC$ + Chr$(13) + Chr$(10)
          End If
          NOTAC$ = NOTAC$ + RTrim$(TTYY$)
       End If
    Next UKL&
    '
    'GRABACION
    On Error GoTo 0
    With DEUDORRR
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !Stat_Cli = 1                        'STATUS ACTIVO
      !Nume_Cli = NCI%                     'CODIGO ENTERO
      !Codi_Cli = CODICUEN$                'CODIGO TEXT
      !Tcpr_Cli = 0                        'TIPO CTA
      !GrCo_Cli = GruCli%                  'TIPO CLI  VA A GRUPO CLI
      !Raso_Cli = TRIM$(RASO$)             'RAZON SOCIAL
      !Fant_Cli = Left$(FANTA$, 10)        'FANTASIA
      !Domi_Cli = TRIM$(DOMI$)             'DIRECCION
      !Loca_Cli = LOCA$                    'LOCALIDAD
      !Prov_Cli = PROVI$                   'COD PROVINCIA
      !Pais_Cli = TRIM$(PAISOK$)           'PAIS
      !Cpos_Cli = TRIM$(CPOSTX$)           'CODIGO POSTAL
      !Tele_Cli = TRIM$(TELE$)             'TELEFONO)
'      !Tele2_Cli = TELEF2$                '2º TELEFONO
      !Faxi_Cli = FAXX$                    'FAX
      !Cuit_Cli = CUIT$                    'CUIT SIN VALIDAR
      !Cpag_Cli = CONDI%                   'CONDICION DE PAGO
      !Nibr_Cli = TRIM$(INGBRU$)           'IMP BRUTOS
      !Caib_Cli = CAIBRU$
      !Mail_Cli = MAILPRO$                 'MAIL
      !Piva_Cli = PIVA%                    'POS IVA
      '!Posivat_Cli = PIVATXT$             'POSICION DEL IVA EN TEXTO
      '!C_Ret = CRET$                      'RETENCION
      !Hora_Cli = HORARIX$
      !Lcre_Cli = LICRE                    'LIMITE CREDITO
      !Vend_Cli = COVEN%                   'CODIGO VENDEDOR
      !Ctac_Cli = CONTACTX$                'CONTACTO
      !Dsta_Cli = TRIM$(DESSTAN$)          'DESCUENTO
      !Lentre0_Cli = DOMIENT1$             'LUGAR DE ENTREGA
      !Lentre1_Cli = LOCAENT1$             'LOCALIDAD
      !Lentre2_Cli = TRANSPOR$             'TRANSPORTE
      !NTrans_Cli = NTRANX%
      !Hora_Ent = HORARIENT$
      !Zona_Cli = TRIM$(ZONCLI$)
      !Obser_Entre = ""
      !Recar_Embal = 0
      !Gru_Div = GRUDIV%                   'GRUPO DIVISIONAL
      !CuMa_Cli = CUEMA%                   'CTA CONTABLE MADRE
      !Lpre_Cli = LIPRE%                   'LISTA DE PRECIOS
      !Nota_Cli = NOTAC$                   'ANOTACION
      .Update
      Call TRACE(20, U&, UR&)
    End With
  Next U&
  '
  DEUDORRR.Close
  Set DEUDORRR = Nothing
  '
End Sub

Sub MIGRACOTEL()
   '
   FIN& = ULTREG&("COTELEF1")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("COTELEF1", U&)
      NUCOTE$ = TRIM$(Left$(XX$, 12))
      CONDI$ = "Nume_CoTel = '" & NUCOTE$ & "'"
      QTR& = CantRegistros("COTELEF", CONDI$)
      If QTR& < 1 Then                  ' PARA NO SOBRE-ESCRIBIR LO YA MIGRADO
         Label7 = NUCOTE$
         Openforms = DoEvents
         If TRIM$(Label7) <> "" Then
            Call GRADATCOTE
         End If
      End If
      '
   Next U&

End Sub
'
Sub MIGRAOTRAB()
   '
   FIN& = ULTREG&("CORTRAB1")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("CORTRAB1", U&)
      NUCOTE$ = TRIM$(Left$(XX$, 12))
      CONDI$ = "Nume_Ot = '" & NUCOTE$ & "'"
      QTR& = CantRegistros("CORTRAB", CONDI$)
      If QTR& < 1 Then                  ' PARA NO SOBRE-ESCRIBIR LO YA MIGRADO
         Label15 = TRIM$(Left$(XX$, 12))
         Openforms = DoEvents
         If TRIM$(Label15) <> "" Then
            Call GRADATOTRA
         End If
      End If
      '
      QTR& = CantRegistros("DETOTRA", CONDI$)
      If QTR& < 1 Then
         Label15 = TRIM$(Left$(XX$, 12))
         Openforms = DoEvents
         Call GRADETHOR
      End If
   Next U&

End Sub

Sub MIGRATARIN()
   '
   SSTab1.Tab = 2
   FIN& = ULTREG&("TARINTER")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("TARINTER", U&)
      NTAR& = CVS(Left$(XX$, 4))
      CONDI$ = "Nume_Tar = " & NTAR&
      QTR& = CantRegistros("TARINTER", CONDI$)
      If QTR& < 1 Then                  ' PARA NO SOBRE-ESCRIBIR LO YA MIGRADO
         Label28 = TRIM$(Str$(NTAR&))
         Openforms = DoEvents
         If XVALO(Label28) > 0 Then
            Call Command29_Click        ' PARA GRABAR
         End If
      End If
   Next U&

End Sub

Sub VERIPARTE()
   '
   Call ITARFE09.SOLOPROPIO
   '
   FECON% = DIANTE(HOY(HOS$))
   DSMN% = DIASEM%(FECON%)
   If DSMN% = 6 Or DSMN% = 0 Then Exit Sub    ' SABADO O DOMINGO
   '
   ITARFE09.Label14 = FECHATEX$(FECON%)
   '
   CANOR = XVALO(ITARFE09.Label12)
   If CANOR < 7 Then
     If USERID$ = "OD" Or USERID$ = "MC" Or USERID$ = "EG" Then
       ITARFE09.Label18 = "CONTROL"
       ITARFE09.Show 1
       Exit Sub
     End If
   End If
   '
   ITARFE09.Label14 = HOS$
   '
End Sub
