VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form AMAEQUI09 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maquinarias - A-B-M-L Maestro de Items de Maquinarias y Equipos"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   3885
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   117
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   5040
      OleObjectBlob   =   "AMAEQUI09.frx":0000
      TabIndex        =   116
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6615
      OleObjectBlob   =   "AMAEQUI09.frx":007A
      Top             =   0
   End
   Begin VB.TextBox Text21 
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
      Left            =   8760
      MaxLength       =   64
      MultiLine       =   -1  'True
      TabIndex        =   105
      Top             =   420
      Width           =   1980
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
      Index           =   1
      Left            =   1845
      Picture         =   "AMAEQUI09.frx":02AE
      TabIndex        =   31
      Top             =   420
      Width           =   175
   End
   Begin VB.TextBox Text3 
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
      Left            =   2520
      MaxLength       =   64
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   5730
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00006060&
      Height          =   8310
      Index           =   0
      Left            =   11040
      ScaleHeight     =   8250
      ScaleWidth      =   1005
      TabIndex        =   49
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command21 
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
         Picture         =   "AMAEQUI09.frx":05B8
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Alta de Nuevo Equipo - Maquinaria"
         Top             =   2940
         Width           =   650
      End
      Begin VB.CommandButton Command22 
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
         Picture         =   "AMAEQUI09.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Alta de Nuevo Equipo por Copia "
         Top             =   3585
         Width           =   650
      End
      Begin VB.CommandButton Command24 
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
         Picture         =   "AMAEQUI09.frx":084C
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Baja"
         Top             =   4230
         Width           =   650
      End
      Begin VB.CommandButton Command37 
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
         Left            =   120
         Picture         =   "AMAEQUI09.frx":0B56
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   5040
         Width           =   650
      End
      Begin VB.CommandButton Command38 
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
         Picture         =   "AMAEQUI09.frx":0F98
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Listado de Atributos Generales"
         Top             =   5685
         Width           =   650
      End
      Begin VB.CommandButton Command39 
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
         Picture         =   "AMAEQUI09.frx":12A2
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Imprime Ficha de Equipo"
         Top             =   6330
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "AMAEQUI09.frx":190C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   1480
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "AMAEQUI09.frx":1C16
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   2125
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   51
         Top             =   7140
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   52
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command3 
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
         Picture         =   "AMAEQUI09.frx":1F20
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   105
         Width           =   650
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Check"
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
         Picture         =   "AMAEQUI09.frx":206A
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8730
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
         Height          =   640
         Left            =   105
         Picture         =   "AMAEQUI09.frx":2374
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Ayuda en Línea"
         Top             =   750
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "AMAEQUI09.frx":267E
         Top             =   7455
         Width           =   1515
      End
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
      Left            =   720
      MaxLength       =   15
      TabIndex        =   0
      Text            =   " "
      Top             =   420
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6945
      Left            =   105
      TabIndex        =   30
      Top             =   945
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   12250
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "AMAEQUI09.frx":45A0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MARCOFICHA(0)"
      Tab(0).Control(1)=   "Frame22"
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(3)=   "Label17"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Componentes"
      TabPicture(1)   =   "AMAEQUI09.frx":45BC
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Imágenes"
      TabPicture(2)   =   "AMAEQUI09.frx":45D8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5(1)"
      Tab(2).Control(1)=   "Picture3"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Control"
      TabPicture(3)   =   "AMAEQUI09.frx":45F4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(0)"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   6420
         Left            =   105
         TabIndex        =   149
         Top             =   420
         Width           =   10620
         Begin VB.Frame Frame28 
            Caption         =   "Plan de Mantenimiento Calendario"
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
            Left            =   5565
            TabIndex        =   150
            Top             =   5
            Width           =   4860
            Begin VB.CommandButton Command14 
               Caption         =   "A.B.M."
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
               Left            =   240
               TabIndex        =   153
               Top             =   240
               Width           =   1455
            End
            Begin VB.CommandButton Command15 
               Caption         =   "Alta por Copia"
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
               Left            =   1705
               TabIndex        =   152
               Top             =   240
               Width           =   1455
            End
            Begin VB.CommandButton Command16 
               Caption         =   "Listados"
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
               Left            =   3160
               TabIndex        =   151
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "SubConjuntos y Partes Componentes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3375
            Left            =   315
            TabIndex        =   162
            Top             =   105
            Width           =   10095
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00C0C0C0&
               Height          =   435
               Index           =   0
               Left            =   120
               ScaleHeight     =   375
               ScaleWidth      =   9780
               TabIndex        =   169
               Top             =   735
               Width           =   9840
               Begin VB.Label Label4 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   $"AMAEQUI09.frx":4610
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   0
                  TabIndex        =   170
                  Top             =   0
                  Width           =   9795
               End
            End
            Begin VB.ListBox List2 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2220
               Left            =   120
               TabIndex        =   168
               Top             =   960
               Width           =   9840
            End
            Begin VB.Frame Frame31 
               Caption         =   "Frame2"
               Height          =   2400
               Left            =   570
               TabIndex        =   167
               Top             =   765
               Width           =   15
            End
            Begin VB.Frame Frame18 
               Caption         =   "Frame18"
               Height          =   2400
               Left            =   2800
               TabIndex        =   166
               Top             =   765
               Width           =   15
            End
            Begin VB.Frame Frame19 
               Caption         =   "Frame19"
               Height          =   2400
               Left            =   5320
               TabIndex        =   165
               Top             =   765
               Width           =   15
            End
            Begin VB.Frame Frame20 
               Caption         =   "Frame20"
               Height          =   2400
               Left            =   7400
               TabIndex        =   164
               Top             =   765
               Width           =   15
            End
            Begin VB.Frame Frame29 
               Caption         =   "Frame20"
               Height          =   2400
               Left            =   9480
               TabIndex        =   163
               Top             =   765
               Width           =   15
            End
            Begin VB.Line Line1 
               X1              =   9960
               X2              =   9960
               Y1              =   720
               Y2              =   3200
            End
            Begin VB.Line Line13 
               X1              =   120
               X2              =   9950
               Y1              =   3180
               Y2              =   3180
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Repuestos, Lubricantes y Consumibles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2730
            Left            =   315
            TabIndex        =   154
            Top             =   3570
            Width           =   10095
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00C0C0C0&
               Height          =   435
               Index           =   0
               Left            =   120
               ScaleHeight     =   375
               ScaleWidth      =   9780
               TabIndex        =   160
               Top             =   480
               Width           =   9840
               Begin VB.Label Label27 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Codigo                           Descripcion                                                U-M     Cantidad     Observaciones"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   0
                  TabIndex        =   161
                  Top             =   0
                  Width           =   9765
               End
            End
            Begin VB.ListBox List1 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1860
               Left            =   120
               TabIndex        =   159
               Top             =   740
               Width           =   9840
            End
            Begin VB.Frame Frame21 
               Caption         =   "Frame21"
               Height          =   2100
               Left            =   1920
               TabIndex        =   158
               Top             =   525
               Width           =   15
            End
            Begin VB.Frame Frame23 
               Caption         =   "Frame23"
               Height          =   2100
               Left            =   5040
               TabIndex        =   157
               Top             =   525
               Width           =   15
            End
            Begin VB.Frame Frame24 
               Caption         =   "Frame24"
               Height          =   2100
               Left            =   5640
               TabIndex        =   156
               Top             =   525
               Width           =   15
            End
            Begin VB.Frame Frame25 
               Caption         =   "Frame25"
               Height          =   2100
               Left            =   6600
               TabIndex        =   155
               Top             =   525
               Width           =   15
            End
            Begin VB.Line Line3 
               X1              =   9900
               X2              =   9900
               Y1              =   480
               Y2              =   960
            End
            Begin VB.Line Line11 
               X1              =   9960
               X2              =   9960
               Y1              =   525
               Y2              =   2600
            End
            Begin VB.Line Line12 
               X1              =   255
               X2              =   9900
               Y1              =   2590
               Y2              =   2590
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6500
         Index           =   0
         Left            =   -75000
         TabIndex        =   118
         Top             =   420
         Width           =   10800
         Begin VB.Frame Frame3 
            Caption         =   "Contadores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   105
            TabIndex        =   138
            Top             =   105
            Width           =   10620
            Begin VB.Frame Frame26 
               Caption         =   "Frame26"
               Height          =   1695
               Index           =   5
               Left            =   8000
               TabIndex        =   148
               Top             =   315
               Width           =   15
            End
            Begin VB.Frame Frame26 
               Caption         =   "Frame26"
               Height          =   1695
               Index           =   4
               Left            =   6820
               TabIndex        =   147
               Top             =   315
               Width           =   15
            End
            Begin VB.Frame Frame26 
               Caption         =   "Frame26"
               Height          =   1695
               Index           =   3
               Left            =   5865
               TabIndex        =   146
               Top             =   315
               Width           =   15
            End
            Begin VB.Frame Frame26 
               Caption         =   "Frame26"
               Height          =   1695
               Index           =   2
               Left            =   4700
               TabIndex        =   145
               Top             =   315
               Width           =   15
            End
            Begin VB.Frame Frame26 
               Caption         =   "Frame26"
               Height          =   1695
               Index           =   1
               Left            =   3750
               TabIndex        =   144
               Top             =   315
               Width           =   15
            End
            Begin VB.Frame Frame27 
               Caption         =   "Frame26"
               Height          =   1695
               Left            =   3250
               TabIndex        =   143
               Top             =   315
               Width           =   15
            End
            Begin VB.Frame Frame26 
               Caption         =   "Frame26"
               Height          =   1695
               Index           =   0
               Left            =   420
               TabIndex        =   142
               Top             =   315
               Width           =   15
            End
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
               Height          =   1380
               IntegralHeight  =   0   'False
               ItemData        =   "AMAEQUI09.frx":46C1
               Left            =   0
               List            =   "AMAEQUI09.frx":46C3
               TabIndex        =   141
               Top             =   630
               Width           =   10585
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00C0C0C0&
               Height          =   400
               Index           =   1
               Left            =   0
               ScaleHeight     =   345
               ScaleWidth      =   10530
               TabIndex        =   139
               Top             =   360
               Width           =   10585
               Begin VB.Label Label31 
                  BackStyle       =   0  'Transparent
                  Caption         =   $"AMAEQUI09.frx":46C5
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   0
                  TabIndex        =   140
                  Top             =   0
                  Width           =   10995
               End
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Anotaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2480
            Index           =   1
            Left            =   105
            TabIndex        =   131
            Top             =   2235
            Width           =   10620
            Begin VB.Frame Frame13 
               Caption         =   "Frame3"
               Height          =   4110
               Left            =   1155
               TabIndex        =   137
               Top             =   300
               Width           =   15
            End
            Begin VB.Frame Frame12 
               Caption         =   "Frame4"
               Height          =   4005
               Left            =   8085
               TabIndex        =   136
               Top             =   300
               Width           =   15
            End
            Begin VB.ListBox List4 
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
               Height          =   1860
               Left            =   0
               TabIndex        =   135
               Top             =   590
               Width           =   10575
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00C0C0C0&
               Height          =   750
               Index           =   1
               Left            =   0
               ScaleHeight     =   690
               ScaleWidth      =   10515
               TabIndex        =   133
               Top             =   315
               Width           =   10580
               Begin VB.Label Label2 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   0  'Transparent
                  Caption         =   $"AMAEQUI09.frx":475E
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   0
                  TabIndex        =   134
                  Top             =   0
                  Width           =   10515
               End
            End
            Begin VB.Frame Frame11 
               Caption         =   "Frame7"
               Height          =   3795
               Left            =   10290
               TabIndex        =   132
               Top             =   630
               Width           =   15
            End
            Begin VB.Line Line5 
               X1              =   10580
               X2              =   10580
               Y1              =   320
               Y2              =   6225
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   10605
               Y1              =   4410
               Y2              =   4410
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Novedades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   105
            TabIndex        =   119
            Top             =   4815
            Width           =   10620
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00E0E0E0&
               Height          =   960
               Left            =   0
               ScaleHeight     =   900
               ScaleWidth      =   1110
               TabIndex        =   128
               Top             =   590
               Width           =   1170
               Begin VB.Line Line10 
                  X1              =   0
                  X2              =   2205
                  Y1              =   340
                  Y2              =   340
               End
               Begin VB.Label Label19 
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
                  Height          =   345
                  Left            =   0
                  TabIndex        =   130
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.Label Label18 
                  BackStyle       =   0  'Transparent
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
                  Height          =   225
                  Left            =   0
                  TabIndex        =   129
                  Top             =   105
                  Width           =   2175
               End
            End
            Begin VB.PictureBox Picture7 
               BackColor       =   &H00E0E0E0&
               Height          =   960
               Left            =   8295
               ScaleHeight     =   900
               ScaleWidth      =   2205
               TabIndex        =   124
               Top             =   590
               Width           =   2265
               Begin VB.CommandButton Command5 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   530
                  Left            =   0
                  Picture         =   "AMAEQUI09.frx":4813
                  Style           =   1  'Graphical
                  TabIndex        =   125
                  Top             =   380
                  Width           =   750
               End
               Begin VB.Label Label23 
                  BackStyle       =   0  'Transparent
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
                  Height          =   350
                  Left            =   0
                  TabIndex        =   127
                  Top             =   0
                  Width           =   2250
               End
               Begin VB.Line Line2 
                  X1              =   0
                  X2              =   2205
                  Y1              =   340
                  Y2              =   340
               End
               Begin VB.Label Label20 
                  BackStyle       =   0  'Transparent
                  Caption         =   "  Registrar             Anotación"
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
                  Left            =   735
                  TabIndex        =   126
                  Top             =   380
                  Width           =   1550
               End
            End
            Begin VB.TextBox Text2 
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
               Left            =   1155
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   123
               Top             =   590
               Width           =   7155
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00C0C0C0&
               Height          =   735
               Index           =   2
               Left            =   0
               ScaleHeight     =   675
               ScaleWidth      =   10515
               TabIndex        =   120
               Top             =   315
               Width           =   10580
               Begin VB.Label Label9 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   0  'Transparent
                  Caption         =   $"AMAEQUI09.frx":4B1D
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
                  Left            =   0
                  TabIndex        =   122
                  Top             =   0
                  Width           =   10515
               End
               Begin VB.Label Label1 
                  Caption         =   "Label5"
                  Height          =   15
                  Left            =   0
                  TabIndex        =   121
                  Top             =   420
                  Width           =   1380
               End
               Begin VB.Line Line6 
                  X1              =   0
                  X2              =   1155
                  Y1              =   420
                  Y2              =   420
               End
            End
            Begin VB.Line Line9 
               X1              =   10500
               X2              =   10500
               Y1              =   315
               Y2              =   1560
            End
            Begin VB.Line Line7 
               X1              =   0
               X2              =   10665
               Y1              =   1550
               Y2              =   1550
            End
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   5655
         Left            =   -74760
         ScaleHeight     =   5595
         ScaleWidth      =   9435
         TabIndex        =   83
         Top             =   690
         Width           =   9495
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   99000
            Left            =   0
            ScaleHeight     =   99000
            ScaleWidth      =   9195
            TabIndex        =   85
            Top             =   0
            Width           =   9200
            Begin VB.PictureBox Picture99 
               AutoRedraw      =   -1  'True
               Height          =   1695
               Index           =   0
               Left            =   240
               ScaleHeight     =   1635
               ScaleWidth      =   2850
               TabIndex        =   91
               ToolTipText     =   "Doble-Click para Maximizar"
               Top             =   720
               Width           =   2905
            End
            Begin VB.TextBox Text99 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1395
               Index           =   0
               Left            =   3240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   90
               Top             =   720
               Width           =   5835
            End
            Begin VB.CommandButton Command99 
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   0
               Left            =   8880
               TabIndex        =   89
               Top             =   2190
               Width           =   175
            End
            Begin VB.PictureBox Picture99 
               AutoRedraw      =   -1  'True
               Height          =   1695
               Index           =   1
               Left            =   240
               ScaleHeight     =   1635
               ScaleWidth      =   2850
               TabIndex        =   88
               ToolTipText     =   "Dbl-Click para Maximizar"
               Top             =   3120
               Width           =   2905
            End
            Begin VB.TextBox Text99 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1395
               Index           =   1
               Left            =   3240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   87
               Top             =   3120
               Width           =   5835
            End
            Begin VB.CommandButton Command99 
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   1
               Left            =   8880
               TabIndex        =   86
               Top             =   4590
               Width           =   175
            End
            Begin VB.Label Label99 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   0
               Left            =   3240
               TabIndex        =   93
               Top             =   2190
               Width           =   5655
            End
            Begin VB.Label Label99 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   1
               Left            =   3240
               TabIndex        =   92
               Top             =   4590
               Width           =   5655
            End
            Begin VB.Image Image99 
               Height          =   10005
               Index           =   0
               Left            =   735
               Top             =   1260
               Visible         =   0   'False
               Width           =   10005
            End
            Begin VB.Image Image99 
               Height          =   10005
               Index           =   1
               Left            =   735
               Top             =   630
               Visible         =   0   'False
               Width           =   10005
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   5600
            Left            =   9200
            TabIndex        =   84
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ctrl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Index           =   1
         Left            =   -65160
         TabIndex        =   80
         Top             =   600
         Width           =   825
         Begin VB.CommandButton Command12 
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
            Left            =   85
            Picture         =   "AMAEQUI09.frx":4BD4
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Eliminar Imagen"
            Top             =   1320
            Width           =   650
         End
         Begin VB.CommandButton Command13 
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
            Left            =   85
            Picture         =   "AMAEQUI09.frx":4EDE
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Agregar Imagen"
            Top             =   370
            Width           =   650
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Height          =   6540
         Index           =   0
         Left            =   -74880
         TabIndex        =   36
         Top             =   360
         Width           =   10560
         Begin VB.Frame Frame30 
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
            Height          =   920
            Left            =   5520
            TabIndex        =   108
            Top             =   3360
            Width           =   4860
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
               Height          =   300
               Left            =   3465
               TabIndex        =   110
               Top             =   410
               Width           =   1185
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
               Height          =   300
               Left            =   2160
               TabIndex        =   109
               Top             =   400
               Width           =   175
            End
            Begin VB.TextBox Text24 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1200
               TabIndex        =   111
               Top             =   410
               Width           =   1065
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Equipo Alternativo:"
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
               Index           =   17
               Left            =   -165
               TabIndex        =   113
               Top             =   330
               Width           =   1335
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "$ Cuota Horaria:"
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
               Index           =   16
               Left            =   2400
               TabIndex        =   112
               Top             =   330
               Width           =   975
            End
         End
         Begin VB.Frame Frame10 
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
            Height          =   1920
            Left            =   5520
            TabIndex        =   74
            Top             =   4320
            Width           =   4860
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
               Height          =   1365
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   19
               Top             =   360
               Width           =   4635
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Datos Físicos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            Left            =   5520
            TabIndex        =   71
            Top             =   1200
            Width           =   4860
            Begin VB.Frame Frame17 
               Caption         =   "Medidas"
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
               Left            =   120
               TabIndex        =   101
               Top             =   1200
               Width           =   4635
               Begin VB.TextBox Text20 
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
                  Height          =   300
                  Left            =   3465
                  MaxLength       =   20
                  TabIndex        =   18
                  Top             =   315
                  Width           =   1035
               End
               Begin VB.TextBox Text19 
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
                  Height          =   300
                  Left            =   1890
                  MaxLength       =   20
                  TabIndex        =   17
                  Top             =   315
                  Width           =   1035
               End
               Begin VB.TextBox Text18 
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
                  Height          =   300
                  Left            =   315
                  MaxLength       =   20
                  TabIndex        =   16
                  Top             =   315
                  Width           =   1035
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "H:"
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
                  Index           =   15
                  Left            =   2100
                  TabIndex        =   104
                  Top             =   420
                  Width           =   1335
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "A:"
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
                  Index           =   12
                  Left            =   525
                  TabIndex        =   103
                  Top             =   420
                  Width           =   1335
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "L:"
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
                  Index           =   11
                  Left            =   -1050
                  TabIndex        =   102
                  Top             =   420
                  Width           =   1335
               End
            End
            Begin VB.TextBox Text5 
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
               Height          =   300
               Left            =   3450
               MaxLength       =   20
               TabIndex        =   15
               Top             =   825
               Width           =   1140
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
               Height          =   300
               Left            =   1560
               MaxLength       =   48
               TabIndex        =   14
               Top             =   825
               Width           =   1245
            End
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
               Height          =   300
               Left            =   1560
               MaxLength       =   48
               TabIndex        =   13
               Top             =   480
               Width           =   3030
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Kilos:"
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
               Index           =   9
               Left            =   2100
               TabIndex        =   100
               Top             =   930
               Width           =   1335
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Potencia:"
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
               Index           =   13
               Left            =   105
               TabIndex        =   73
               Top             =   930
               Width           =   1335
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Alimentación / Combustible:"
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
               Index           =   6
               Left            =   0
               TabIndex        =   72
               Top             =   390
               Width           =   1455
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Ubicación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1605
            Left            =   0
            TabIndex        =   37
            Top             =   4630
            Width           =   5445
            Begin VB.CommandButton Command17 
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
               Left            =   3360
               TabIndex        =   107
               Top             =   1125
               Width           =   175
            End
            Begin VB.CommandButton Command8 
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
               Left            =   2040
               TabIndex        =   75
               Top             =   375
               Width           =   175
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
               Height          =   300
               Left            =   2280
               TabIndex        =   54
               Top             =   750
               Width           =   175
            End
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
               Height          =   300
               Left            =   4560
               MaxLength       =   12
               TabIndex        =   10
               Top             =   1125
               Width           =   735
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
               Height          =   300
               Left            =   1410
               MaxLength       =   48
               TabIndex        =   9
               Top             =   1125
               Width           =   2055
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
               Height          =   300
               Left            =   1410
               MaxLength       =   4
               TabIndex        =   7
               Text            =   "  "
               Top             =   375
               Width           =   690
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
               Height          =   300
               Left            =   1410
               TabIndex        =   8
               Text            =   " "
               Top             =   750
               Width           =   930
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Index           =   14
               Left            =   3120
               TabIndex        =   79
               Top             =   1215
               Width           =   1335
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Lineas / Sistemas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Index           =   8
               Left            =   360
               TabIndex        =   78
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label16 
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
               Height          =   300
               Left            =   2370
               TabIndex        =   77
               Top             =   375
               Width           =   2910
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
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
               Height          =   540
               Left            =   0
               TabIndex        =   76
               Top             =   465
               Width           =   1335
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
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
               Height          =   540
               Left            =   0
               TabIndex        =   56
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label Label5 
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
               Height          =   300
               Left            =   2610
               TabIndex        =   55
               Top             =   750
               Width           =   2670
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Observaciones:"
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
               Index           =   3
               Left            =   240
               TabIndex        =   53
               Top             =   3240
               Width           =   1335
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
            Height          =   2505
            Left            =   0
            TabIndex        =   66
            Top             =   1995
            Width           =   5445
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
               Left            =   1965
               TabIndex        =   67
               Top             =   735
               Width           =   175
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
               Height          =   300
               Left            =   1425
               MaxLength       =   48
               TabIndex        =   5
               Top             =   360
               Width           =   3885
            End
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
               Height          =   300
               Left            =   1425
               TabIndex        =   6
               Text            =   "  "
               Top             =   735
               Width           =   570
            End
            Begin VB.Label Label30 
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
               Height          =   540
               Left            =   15
               TabIndex        =   99
               Top             =   2120
               Width           =   1335
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "CP - Localidad:"
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
               Left            =   15
               TabIndex        =   98
               Top             =   1685
               Width           =   1335
            End
            Begin VB.Label Label28 
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
               Height          =   540
               Left            =   15
               TabIndex        =   97
               Top             =   1250
               Width           =   1335
            End
            Begin VB.Label Label26 
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
               Height          =   300
               Left            =   1425
               TabIndex        =   96
               Top             =   1995
               Width           =   3885
            End
            Begin VB.Label Label25 
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
               Height          =   300
               Left            =   1425
               TabIndex        =   95
               Top             =   1575
               Width           =   3870
            End
            Begin VB.Label Label24 
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
               Height          =   300
               Left            =   1425
               TabIndex        =   94
               Top             =   1155
               Width           =   3885
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fabricante:"
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
               Index           =   10
               Left            =   0
               TabIndex        =   70
               Top             =   495
               Width           =   1335
            End
            Begin VB.Label Label13 
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
               Height          =   300
               Left            =   2280
               TabIndex        =   69
               Top             =   735
               Width           =   3030
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Representante:"
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
               Left            =   15
               TabIndex        =   68
               Top             =   840
               Width           =   1335
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Identificación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1620
            Left            =   0
            TabIndex        =   62
            Top             =   210
            Width           =   5445
            Begin VB.CommandButton Command20 
               Caption         =   "?"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   4850
               TabIndex        =   115
               ToolTipText     =   "Consulta Amortización"
               Top             =   1140
               Width           =   495
            End
            Begin VB.CommandButton Command18 
               Caption         =   "$"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   4320
               TabIndex        =   114
               ToolTipText     =   "Datos Contables"
               Top             =   1140
               Width           =   495
            End
            Begin VB.TextBox Text23 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1426
               MaxLength       =   48
               TabIndex        =   2
               Top             =   360
               Width           =   3870
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
               Height          =   300
               Index           =   1
               Left            =   1425
               MaxLength       =   48
               TabIndex        =   3
               Top             =   765
               Width           =   3870
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
               Height          =   300
               Left            =   1425
               MaxLength       =   48
               TabIndex        =   4
               Top             =   1140
               Width           =   2790
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Marca:"
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
               Index           =   2
               Left            =   0
               TabIndex        =   65
               Top             =   390
               Width           =   1335
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Modelo:"
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
               Index           =   0
               Left            =   0
               TabIndex        =   64
               Top             =   795
               Width           =   1335
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nº de Serie:"
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
               Index           =   1
               Left            =   0
               TabIndex        =   63
               Top             =   1155
               Width           =   1335
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Instalación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   920
            Left            =   5520
            TabIndex        =   57
            Top             =   210
            Width           =   4860
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
               Height          =   300
               Left            =   4440
               TabIndex        =   61
               Top             =   410
               Width           =   175
            End
            Begin VB.CommandButton Command6 
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
               Left            =   2280
               TabIndex        =   60
               Top             =   400
               Width           =   175
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
               Height          =   300
               Left            =   3465
               TabIndex        =   12
               Top             =   410
               Width           =   1065
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
               Height          =   300
               Left            =   1320
               TabIndex        =   11
               Top             =   410
               Width           =   1065
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fin de la Garantía:"
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
               Index           =   7
               Left            =   2040
               TabIndex        =   59
               Top             =   330
               Width           =   1335
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de Instalación:"
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
               Index           =   4
               Left            =   -50
               TabIndex        =   58
               Top             =   330
               Width           =   1335
            End
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "No Usados"
         Height          =   4215
         Left            =   -74475
         TabIndex        =   38
         Top             =   930
         Width           =   4320
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
            Height          =   330
            Left            =   5010
            TabIndex        =   42
            Top             =   1200
            Width           =   175
         End
         Begin VB.CommandButton Command35 
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
            Left            =   5010
            TabIndex        =   41
            Top             =   1710
            Width           =   175
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lote Máximo:"
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
            Left            =   3465
            TabIndex        =   47
            Top             =   630
            Width           =   1245
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Minimo:"
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
            Left            =   210
            TabIndex        =   46
            Top             =   630
            Width           =   1395
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Criterio de Reposicion:"
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
            Left            =   210
            TabIndex        =   45
            Top             =   0
            Width           =   1380
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "F.Des:"
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
            Left            =   210
            TabIndex        =   44
            Top             =   1890
            Width           =   855
         End
         Begin VB.Label Label39 
            BackColor       =   &H00E0E0E0&
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
            Height          =   330
            Left            =   1155
            TabIndex        =   43
            Top             =   2205
            Width           =   3375
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo I.V.A.:"
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
            Left            =   -105
            TabIndex        =   40
            Top             =   630
            Width           =   975
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo Contable:"
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
            Index           =   0
            Left            =   -420
            TabIndex        =   39
            Top             =   210
            Width           =   1290
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -71205
         TabIndex        =   35
         Top             =   4605
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -70785
         TabIndex        =   34
         Top             =   3660
         Width           =   15
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*."
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo Alias:"
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
      Index           =   5
      Left            =   8160
      TabIndex        =   106
      Top             =   180
      Width           =   1380
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   11865
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label COMMLAB 
      Height          =   225
      Left            =   105
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
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
      Index           =   38
      Left            =   1920
      TabIndex        =   33
      Top             =   180
      Width           =   1380
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
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
      Left            =   405
      TabIndex        =   32
      Top             =   180
      Width           =   795
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu Abri 
         Caption         =   "Abrir"
      End
      Begin VB.Menu phl5 
         Caption         =   "-"
      End
      Begin VB.Menu Anuevo 
         Caption         =   "Nuevo Equipo"
      End
      Begin VB.Menu Acopia 
         Caption         =   "Alta por Copia"
      End
      Begin VB.Menu plh0 
         Caption         =   "-"
      End
      Begin VB.Menu Guar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu Imp 
         Caption         =   "Imprimir Ficha"
      End
      Begin VB.Menu plh2 
         Caption         =   "-"
      End
      Begin VB.Menu Sal 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Edi 
      Caption         =   "Edición"
      Begin VB.Menu DesCam 
         Caption         =   "Deshacer Cambios"
      End
      Begin VB.Menu phl4 
         Caption         =   "-"
      End
      Begin VB.Menu BaLog 
         Caption         =   "Baja"
      End
      Begin VB.Menu php16 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAddImg 
         Caption         =   "Agregar Imagen"
      End
      Begin VB.Menu MnuMaxImg 
         Caption         =   "Maximizar Imagen"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuDelImg 
         Caption         =   "Eliminar Imagen"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu ListCon 
      Caption         =   "Listados y Consultas"
      Begin VB.Menu AtGral 
         Caption         =   "Atributos Generales"
      End
   End
   Begin VB.Menu AccDire 
      Caption         =   "Accesos Directos"
      Begin VB.Menu MaeRep 
         Caption         =   "Repuestos"
      End
      Begin VB.Menu php6a 
         Caption         =   "-"
      End
      Begin VB.Menu Otro 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu ConGen 
      Caption         =   "Configuración "
      Begin VB.Menu MnuTab 
         Caption         =   "Tablas"
         Begin VB.Menu MnuTiMa 
            Caption         =   "Tipo de Material"
         End
         Begin VB.Menu MnUniMedi 
            Caption         =   "Unidad de Medida"
         End
         Begin VB.Menu php10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlan 
            Caption         =   "Plantas"
         End
         Begin VB.Menu MnuSecc 
            Caption         =   "Secciones y Centro de Costos"
         End
         Begin VB.Menu MnuDep 
            Caption         =   "Dépositos y Almacenes"
         End
         Begin VB.Menu MnuEspec 
            Caption         =   "Especialidades"
         End
      End
      Begin VB.Menu MnuFoImp 
         Caption         =   "Formularios de Impresión"
      End
   End
   Begin VB.Menu ayu 
      Caption         =   "Ayuda"
      Begin VB.Menu mnuAyud 
         Caption         =   "Ayuda FLEXOFT en Línea"
      End
   End
End
Attribute VB_Name = "AMAEQUI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim INTERVER
Dim MODIFIC%
Dim YACAR%
Dim IMGACT%

Private Sub Abri_Click()
  '
  Call Command9_Click(1)
  '
End Sub

Private Sub Acopia_Click()
  '
  Call Command22_Click
  '
End Sub

Private Sub Anuevo_Click()
  '
  Call Command21_Click
  '
End Sub

Private Sub AtGral_Click()
  '
   Call FINAL("?EQUILIST") 'OK
  '
End Sub

Private Sub BaLog_Click()
  '
  Command24_Click
  '
End Sub
Private Sub Command1_Click()
    If TRIM$(Text1.TEXT) <> "" And CantRegistros("EQUIPOS", "Codigo_Equ = '" & TRIM$(Text1) & "'") > 0 Then
      Call GRABEQ
      Call INDIEQUI
      Call GRAIMG
      Call GRAREP
      Call GRACON
      Call GRASUB
    Else
      Call COMUNI("Debe Seleccionar un Equipo.")
    End If
End Sub

Private Sub Command10_Click()
  '
  If TRIM$(Text1) <> "" Then
    Call SELECHO("SECPROD")
    COD$ = VALACT1$("SECPROD")
    If TRIM$(COD$) <> "" Then
      Text12.TEXT = COD$
      CS1 = Val(Text12.TEXT)
      Label5.Caption = TRIM$(ECOARCH$("SECPROD", MKS$(CS1)))
    End If
  End If
  '
End Sub

Private Sub Command11_Click()
   '
   If TRIM$(Text1) <> "" Then
    Call SELECHO("ACREDOR")
    VALONUE$ = VALACT1$("ACREDOR")
    If VALONUE$ <> "" Then
       Text7.TEXT = VALONUE$
       Label13.Caption = ECOARCH("ACREDOR", MKI$(Int(Val(Text7.TEXT))))
    End If
   End If
   '
End Sub

Private Sub Command12_Click()
  'ELIMINAR BLOQUE DE IMAGENES
  'BLOQUE
'  For JX% = O To Text99.Count - 1
'    Unload Text99(JX%)  'falta poder identificar que imagen se quiere borrar
'  Next JX%
  '
End Sub

Private Sub Command13_Click()
  '
  Command13.Enabled = False
  'If TRIM$(Text1.TEXT) <> "" And ECOARCH$("INDIEQUI", Text1.TEXT) <> "" Then
  If CantRegistros("EQUIPOS", "codigo_equ= '" & TRIM$(Text1.TEXT) & "'") > 0 Then
    Call CARGABLOIMG   'AGREGA COMBO DE IMAGEN
    On Error Resume Next
      'SE POSICIONA EN LA ULTIMA IMAGEN
       VScroll1.Value = VScroll1.Max
    On Error GoTo 0
  End If
  Command13.Enabled = True
  '
End Sub
Sub CARGABLOIMG() 'CARGA BLOQUE DE IMAGEN
  '
  'Agrega un Conjunto de Imagen.
  'IMAGEN
  NUPIC% = Picture99.Count
  Load Picture99(NUPIC%)
  Picture99(NUPIC%).Top = Picture99(NUPIC% - 1).Top + Picture99(NUPIC% - 1).Height + INTERVER
  '
  If Picture99(NUPIC%).Top + Picture99(NUPIC%).Height > Picture4.Height - 500 Then
    Picture4.Height = Picture99(NUPIC%).Top + Picture99(NUPIC%).Height + 500
  End If
  '
  Picture99(NUPIC%).Visible = True
  VScroll1.Max = (Picture4.Height - Picture3.Height) / 10 + 1
  'TEXTO
  Load Text99(NUPIC%)
  Text99(NUPIC%).Top = Picture99(NUPIC%).Top
  Text99(NUPIC%).Visible = True
  Text99(NUPIC%).TEXT = ""
  'IMAGE
  Load Image99(NUPIC%)
  'LABEL
  Load Label99(NUPIC%)
  Label99(NUPIC%).Top = Text99(NUPIC%).Top + 1470
  Label99(NUPIC%).Visible = True
  Label99(NUPIC%) = ""
  'COMMAND
  Load Command99(NUPIC%)
  Command99(NUPIC%).Top = Text99(NUPIC%).Top + 1470
  Command99(NUPIC%).Visible = True
  '
  Picture99(NUPIC%).Top = Picture99(NUPIC% - 1).Top + Picture99(NUPIC% - 1).Height + INTERVER
  '
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   CODEQUI$ = Text1
   If CantRegistros("EQUIPOS", "Codigo_Equ = '" & CODEQUI$ & "'") < 1 Then
      Call MENSERR(24, "Debe Elegir Equipo")
      GoTo 99
   End If
   '
   If List2.ListIndex < 0 Or List2.ListIndex >= List2.ListCount Then
      Call MENSERR(24, "Debe Elegir Sub-Conjunto")
      GoTo 99
   End If
   '
   NUCOI& = XVALO(Left$(List2.TEXT, 3))
   If NUCOI& < 1 Then
      Call MENSERR(24, "Sub-Conjunto No Válido")
      GoTo 99
   End If
   '
   DESUB$ = TRIM$(Mid$(List2.TEXT, 4, 23)) + " - " + TRIM$(Mid$(List2.TEXT, 27, 18))
   CONDI$ = "Codigo_Equ = '" & CODEQUI$ & "' "
   CONDI$ = CONDI$ + " AND Nro_Sub = " & NUCOI& & " "
   '
   SQLX$ = "SELECT * FROM PLAMANPREV  WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE " & CONDI$ & " "
   SQLX$ = SQLX$ + " ORDER BY Orden_Tar ASC "
   '
   Dim TAREMAN As ADODB.Recordset
   Set TAREMAN = New ADODB.Recordset
   Set TAREMAN = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   REBLA$ = REGBLAN$("TAREMAN")
   JJ& = 0
   With TAREMAN
     Do While Not .EOF
       XX$ = REBLA$
       Call REPLA(XX$, SAFETEXT$(!descri_tar), 1, 96)
       Call REPLA(XX$, MKI$(XVALO(!Frecu_tar)), 97, 2)
       Call REPLA(XX$, SAFETEXT$(!Uni_Frecu), 99, 8)
       Call REPLA(XX$, SAFETEXT$(!Obse_Tar), 107, 16)
       JJ& = JJ& + 1
       Call GRAREG("!TAREMAN", XX$, JJ&)
     .MoveNext
     Loop
   End With
   Call SETULTREG("!TAREMAN", JJ&)
   '
   VTB% = VENTABU%("TAREMAN/TITUPAN=Plan de Mantenimiento Calendario: " + DESUB$)
   If VTB% >= 0 Then
      Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Mar_Borra", 3, 2)
      Call ACTUREGISTRO("PLAMANPREV", "Mar_Borra", "Mar_Borra = 9999", 0, RAFE&)
      Call ACTUREGISTRO("PLAMANPREV", "Mar_Borra", CONDI$, 1, RAFE&)
      For KKI& = 1 To ULTREG&("!TAREMAN")
         XX$ = REGLEIDO$("!TAREMAN", KKI&)
         Call CreaRegistro("PLAMANPREV", "Mar_Borra", 9999)
         Call ACTUREGISTRO("PLAMANPREV", "Codigo_Equ", "Mar_Borra =  9999", CODEQUI$, RAFE&)
         Call ACTUREGISTRO("PLAMANPREV", "CodLar_Equ", "Mar_Borra = 9999", TEXT21, RAFE&)
         Call ACTUREGISTRO("PLAMANPREV", "Nro_Sub", "Mar_Borra = 9999", NUCOI&, RAFE&)
         Call ACTUREGISTRO("PLAMANPREV", "Orden_Tar", "Mar_Borra = 9999", KKI&, RAFE&)
           DETAR$ = TRIM$(Mid$(XX$, 1, 96))
         Call ACTUREGISTRO("PLAMANPREV", "Descri_Tar", "Mar_Borra = 9999", DETAR$, RAFE&)
           FRECU% = CVI(Mid$(XX$, 97, 2))
         Call ACTUREGISTRO("PLAMANPREV", "Frecu_Tar", "Mar_Borra = 9999", FRECU%, RAFE&)
           UFRE$ = TRIM$(Mid$(XX$, 99, 8))
         Call ACTUREGISTRO("PLAMANPREV", "Uni_Frecu", "Mar_Borra = 9999", UFRE$, RAFE&)
           OBSX$ = TRIM$(Mid$(XX$, 107, 16))
         Call ACTUREGISTRO("PLAMANPREV", "Obse_Tar", "Mar_Borra = 9999", OBSX$, RAFE&)
         Call ACTUREGISTRO("PLAMANPREV", "Mar_Borra", "Mar_Borra = 9999", 0, RAFE&)
      Next KKI&
      '
      Call BORRAREGISTROS("PLAMANPREV", CONDI$ + " AND Mar_Borra = 1", RAFE&)
   End If
   '
99 Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
    Command15.Enabled = False
        COPIATA09.Show 1
    Command15.Enabled = True
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    '
    OPX% = ALTERNA%("Plan Mantenimiento Calendario\Cronograma Mantenimiento Calendario")
    If OPX% = 1 Then Call PRIPLAMACAL
    If OPX% = 2 Then Call PRICROMACAL
    '
    Command16.Enabled = True
End Sub

Sub PRICROMACAL()
'
    Call FRESHECHO("ABMSIS")
    Call SELECHO("ABMSIS")
    LINESIS$ = TRIM$(VALACT1$("ABMSIS"))
    If LINESIS$ = "" Then Exit Sub
    '
    HOII% = HOY(HOS$)
    FECHDOC$ = HOS$
    NUMEDOC$ = ""
    '
10  Call SELECHO("SELFECHA")
    FECC$ = TRIM$(VALACT1$("SELFECHA"))
    If FECC$ = "" Then
      Call MENSERR(24, "Debe Ingresar Fecha de Inicio")
      GoTo 10
    End If
    '
    For II& = 1 To ULTREG&("ABMSIS")
        XX$ = REGLEIDO$("ABMSIS", II&)
        If TRIM$(Mid$(XX$, 1, 64)) = LINESIS$ Then
            CRITICI$ = TRIM$(Mid$(XX$, 65, 32))
            'REVISIO$ = TRIM$(Mid$(XX$, 97, 16))
            NUMEPOE$ = TRIM$(Mid$(XX$, 113, 32))
            Exit For
        End If
    Next II&
    '
    For YY% = 1 To 4
        '
        CODPARAM$(1) = "MEDIDA-2": DOCPARAM$(1) = LINESIS$  'SISTEMA
        CODPARAM$(2) = "Z-PLANO": DOCPARAM$(2) = CRITICI$   'CRITICIDAD DEL SISTEMA
        CODPARAM$(3) = "DOCUORIG": DOCPARAM$(3) = NUMEPOE$  'POE
        CODPARAM$(4) = "FECH-LAN": DOCPARAM$(4) = FECC$       'FECHA DE INICIO
        CODPARAM$(5) = "DET-PASE": DOCPARAM$(5) = FORMATNUM$(YY%, "F1")      'Año de la revision
        CODPARAM$(6) = "REF-MAT4": DOCPARAM$(6) = STRIVENCI$("", 0, YY%, 1)  'Dias 30  60  90  120 ...
        CODPARAM$(7) = "REF-MAT6": DOCPARAM$(7) = STRIVENCI$("", 0, YY%, 2, FECHANUM(FECC$)) 'Dias 30  60  90  120 ...
        CAPARDOC% = 7
        '
        CODTABU1$(1) = "REF-MAT1"       'Descripcion del equipo
        CODTABU1$(2) = "REF-MAT2"       'descripcion del subconjunto
        CODTABU1$(3) = "REF-MAT3"       'Imprime las X
        CODTABU1$(4) = "REF-MAT5"       'CRIT
        CACOLTAB1% = 4
        DESTIDOC% = 0
        CAITAB1% = 0
        CAITCRU% = 0
        '
        Dim LINEQUTMP As ADODB.Recordset
        SQLX$ = " SELECT * FROM EQUIPOS "
        SQLX$ = SQLX$ + "where LINEA_EQU = '" & TRIM$(LINESIS$) & "'"
        Set LINEQUTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        '
        With LINEQUTMP
           Do While Not .EOF
              EQUIPP$ = SAFETEXT$(!codigo_equ)
              DEQUIPP$ = SAFETEXT$(!Descri_Equ)
              '
              CAITAB1% = CAITAB1% + 1
              TETABU1$(1, CAITAB1%) = String$(32, "-")
              '
              CAITAB1% = CAITAB1% + 1
              TETABU1$(1, CAITAB1%) = DEQUIPP$
              TETABU1$(2, CAITAB1%) = ""
              TETABU1$(3, CAITAB1%) = ""
              TETABU1$(4, CAITAB1%) = ""
              '
              CAITAB1% = CAITAB1% + 1
              TETABU1$(1, CAITAB1%) = String$(32, "-")
              '
              Dim EQUIPTMP As ADODB.Recordset
              SQLX$ = " SELECT * FROM SUBCONJUNTOS "
              SQLX$ = SQLX$ + "where CODIGO_EQU = '" & EQUIPP$ & "'"
              SQLX$ = SQLX$ + "ORDER BY NRO_SUB"
              Set EQUIPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
              '
              With EQUIPTMP
                 Do While Not EQUIPTMP.EOF
                    NUSUBCONJ& = XVALO(!NRO_SUB)
                    DESUBCONJ$ = SAFETEXT(!DESCRI_SUB)
                    CRITICIDA$ = SAFETEXT(!Criticidad)
                    '
                    CAITAB1% = CAITAB1% + 1
                    TETABU1$(1, CAITAB1%) = ""
                    TETABU1$(2, CAITAB1%) = DESUBCONJ$
                    TETABU1$(3, CAITAB1%) = STRIVENCI$(EQUIPP$, NUSUBCONJ&, YY%, 0) 'Esta funcion arma un string con las X segun fechas de la tarea
                    TETABU1$(4, CAITAB1%) = CRITICIDA$
                    '
                 .MoveNext
                 Loop
              End With
              EQUIPTMP.Close
              Set EQUIPTMP = Nothing
              '
           .MoveNext
           Loop
        End With
        '
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = String$(32, "-")
        '
        LINEQUTMP.Close
        Set LINEQUTMP = Nothing
        '
    'Next YY&
    Call PRINTDOC("FOR-CMC")   'Cronograma de Mantenimiento Calendario
    Next YY%                  'Se agrega aca para que imprima comience una hoja nueva por cada año

End Sub


 Sub PRIPLAMACAL()
    '
    Call FRESHECHO("ABMSIS")
    Call SELECHO("ABMSIS")
    LINESIS$ = TRIM$(VALACT1$("ABMSIS"))
    If LINESIS$ = "" Then Exit Sub
    '
    HOII% = HOY(HOS$)
    FECHDOC$ = HOS$
    NUMEDOC$ = ""
    '
    For II& = 1 To ULTREG&("ABMSIS")
        XX$ = REGLEIDO$("ABMSIS", II&)
        If TRIM$(Mid$(XX$, 1, 64)) = LINESIS$ Then
            CRITICI$ = TRIM$(Mid$(XX$, 65, 32))
            REVISIO$ = TRIM$(Mid$(XX$, 97, 16))
            NUMEPOE$ = TRIM$(Mid$(XX$, 113, 32))
            Exit For
        End If
    Next II&
    '
    CONDI$ = "LINEA_EQU ='" & LINESIS$ & "'"
    PLANT% = XVALO(VALOBADA("EQUIPOS", "PLANTA_EQU", CONDI$, "NOMESS"))
    CODPARAM$(1) = "MEDIDA-1": DOCPARAM$(1) = TRIM$(ECOARCH$("PLANTPRO", CStr(PLANT%)))        'EDIFICIO
    CODPARAM$(2) = "MEDIDA-2": DOCPARAM$(2) = LINESIS$  'SISTEMA
    CODPARAM$(3) = "Z-PLANO": DOCPARAM$(3) = CRITICI$   'CRITICIDAD
    CODPARAM$(4) = "DOCUORIG": DOCPARAM$(4) = NUMEPOE$  'POE
    CODPARAM$(5) = "R-PLANO": DOCPARAM$(5) = REVISIO$   'REVISION
    CODPARAM$(6) = "EQUIPO-1": DOCPARAM$(6) = ""        'EQUIPO
    CAPARDOC% = 6
    '
    CODTABU1$(1) = "REF-MAT1"
    CODTABU1$(2) = "REF-MAT2"
    CODTABU1$(3) = "REF-MAT3"
    CODTABU1$(4) = "REF-MAT4"
    CODTABU1$(5) = "REF-MAT5"
    CODTABU1$(6) = "REF-MAT6"
    CACOLTAB1% = 6
    DESTIDOC% = 0
    CAITAB1% = 0
    CAITCRU% = 0
    '
    Dim LINEQUTMP As ADODB.Recordset
    SQLX$ = " SELECT * FROM EQUIPOS "
    SQLX$ = SQLX$ + "where LINEA_EQU = '" & TRIM$(LINESIS$) & "'"
    Set LINEQUTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With LINEQUTMP
       Do While Not .EOF
          EQUIPP$ = SAFETEXT$(!codigo_equ)
          DEQUIPP$ = SAFETEXT$(!Descri_Equ)
          '
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = String$(32, "-")
          '
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = DEQUIPP$
          TETABU1$(2, CAITAB1%) = ""
          TETABU1$(3, CAITAB1%) = ""
          TETABU1$(4, CAITAB1%) = ""
          TETABU1$(5, CAITAB1%) = ""
          TETABU1$(6, CAITAB1%) = ""
          '
'          CAITAB1% = CAITAB1% + 1
'          TETABU1$(1, CAITAB1%) = ""
'          TETABU1$(2, CAITAB1%) = ""
'          TETABU1$(3, CAITAB1%) = ""
'          TETABU1$(4, CAITAB1%) = ""
'          TETABU1$(5, CAITAB1%) = ""
'          TETABU1$(6, CAITAB1%) = ""
          '
          Dim EQUIPTMP As ADODB.Recordset
          SQLX$ = " SELECT * FROM SUBCONJUNTOS "
          SQLX$ = SQLX$ + "where CODIGO_EQU = '" & EQUIPP$ & "'"
          SQLX$ = SQLX$ + "ORDER BY NRO_SUB"
          Set EQUIPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
          '
          With EQUIPTMP
             Do While Not EQUIPTMP.EOF
                NUSUBCONJ = XVALO(!NRO_SUB)
                DESUBCONJ$ = SAFETEXT(!DESCRI_SUB)
                '
                CAITAB1% = CAITAB1% + 1
                TETABU1$(1, CAITAB1%) = String$(32, "-")
                '
                CAITAB1% = CAITAB1% + 1
                TETABU1$(1, CAITAB1%) = ""
                TETABU1$(2, CAITAB1%) = DESUBCONJ$
                TETABU1$(3, CAITAB1%) = ""
                TETABU1$(4, CAITAB1%) = ""
                TETABU1$(5, CAITAB1%) = ""
                TETABU1$(6, CAITAB1%) = ""
                '
'                CAITAB1% = CAITAB1% + 1
'                TETABU1$(1, CAITAB1%) = ""
'                TETABU1$(2, CAITAB1%) = ""
'                TETABU1$(3, CAITAB1%) = ""
'                TETABU1$(4, CAITAB1%) = ""
'                TETABU1$(5, CAITAB1%) = ""
'                TETABU1$(6, CAITAB1%) = ""
                '
                Dim SUBCOTMP As ADODB.Recordset
                SQLX$ = " SELECT * FROM PLAMANPREV "
                SQLX$ = SQLX$ + "where NRO_SUB = " & NUSUBCONJ
                SQLX$ = SQLX$ + " AND CODIGO_EQU = '" & EQUIPP$ & "'"
                SQLX$ = SQLX$ + "ORDER BY orden_tar"
                Set SUBCOTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
                '
                CAITAB1% = CAITAB1% + 1
                TETABU1$(1, CAITAB1%) = String$(32, "-")
                '
                With SUBCOTMP
                  Do While Not SUBCOTMP.EOF
                    ORDENTAR = XVALO(!orden_tar)
                    DESCRITAR$ = SAFETEXT(!descri_tar)
                    FRECUTAR$ = XVALO(!Frecu_tar)
                    NUIFRECU = SAFETEXT(!Uni_Frecu)
                    OBSERTAR$ = SAFETEXT(!Obse_Tar)
                    '
                    CAITAB1% = CAITAB1% + 1
                    TETABU1$(1, CAITAB1%) = ""
                    TETABU1$(2, CAITAB1%) = ""
                    TETABU1$(3, CAITAB1%) = DESCRITAR$
                    TETABU1$(4, CAITAB1%) = FRECUTAR$
                    TETABU1$(5, CAITAB1%) = NUIFRECU
                    TETABU1$(6, CAITAB1%) = OBSERTAR$
                    '
                  SUBCOTMP.MoveNext
                  Loop
                End With
                SUBCOTMP.Close
                Set SUBCOTMP = Nothing
             .MoveNext
             Loop
          End With
          EQUIPTMP.Close
          Set EQUIPTMP = Nothing
          '
       .MoveNext
       Loop
    End With
    '
    CAITAB1% = CAITAB1% + 1
    TETABU1$(1, CAITAB1%) = String$(32, "-")
    '
    LINEQUTMP.Close
    Set LINEQUTMP = Nothing
    '
    Call PRINTDOC("FOR-PMC")   ' Plan mantenimiento calendario

 End Sub

Private Sub Command17_Click()
    If TRIM$(Text1) <> "" Then
        Call FRESHECHO("ABMSIS")
        Call SELECHO("ABMSIS")
        If VALACT1$("ABMSIS") <> "" Then
            Text15 = VALACT1$("ABMSIS")
        End If
    End If
End Sub


Function STRIVENCI$(CODEQUI$, NUCOI&, NANO%, MODO%, Optional FEBASE)
   '
   ' MODO%=0 --> dEVUELVE LAS EQUIS
   ' MODO%=1 DEVUELVE LOS DIAS
   ' MODO%=2 DEVUELVE LAS FECHAS
   '
   If XVALO(FEBASE) < 1 Then FEBASE = HOY(HO$)
   If NANO% < 1 Or NANO% > 4 Then NANO% = 1
   INTERMES% = 7     ' CANTIDAD DE CARACTERES PARA CADA PERIODO MENSUAL
   If MODO% > 0 Then INTERMES% = 8
   DIAPER% = 30      ' DIAS DE CADA PERIODO
   CANMES% = 48      ' CANTIDAD DE MESES - LUEGO PASAR A 48
   '
   STRIVE$ = String$(CANMES% * INTERMES%, " ")   ' INICIALIZAR - LUEGO PASAR A ESPACIOS
   Call REPLA(STRIVE$, ".", 1, 1)
   STRIVEFE$ = STRIVE$
   If MODO% > 0 Then
     FINI% = 0: FEBA = FEBASE
8    FINI% = FINI% + DIAPER%
     FEBA = DIASPOST(FEBA, DIAPER%)
     IMES% = FINI% / DIAPER%
     POSIX% = 1 + FINI% * INTERMES% / DIAPER% - INTERMES%
     If POSIX% > Len(STRIVE$) Then GoTo 90
     If POSIX% > 0 Then
        TTXX$ = TRIM$(Str$(FINI%))
        While Len(TTXX$) < INTERMES%: TTXX$ = " " + TTXX$ + " ": Wend
        TTXX$ = Left$(TTXX$, INTERMES%)
        Mid$(STRIVE$, POSIX%, INTERMES%) = TTXX$
        '
        TTXX$ = FECHATEX$(FEBA)
        While Len(TTXX$) < INTERMES%: TTXX$ = " " + TTXX$ + " ": Wend
        TTXX$ = Left$(TTXX$, INTERMES%)
        Mid$(STRIVEFE$, POSIX%, INTERMES%) = TTXX$
     End If
     GoTo 8
   End If
   '
   If MODO% = 0 Then
     CONDI$ = "Codigo_Equ = '" & CODEQUI$ & "' "
     CONDI$ = CONDI$ + " AND Nro_Sub = " & NUCOI& & " "
     '
     SQLX$ = "SELECT * FROM PLAMANPREV  WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE " & CONDI$ & " "
     SQLX$ = SQLX$ + " ORDER BY Orden_Tar ASC "
     '
     Dim TAREMAN As ADODB.Recordset
     Set TAREMAN = New ADODB.Recordset
     Set TAREMAN = FLEXCONN.Execute(FILTSQL$(SQLX$))
     '
     With TAREMAN
       Do While Not .EOF
         If Left$(UCase$(SAFETEXT$(!Uni_Frecu)), 3) = "DIA" Then
           If XVALO(!Frecu_tar) > 0 Then
             FINI% = 0
12              FINI% = FINI% + XVALO(!Frecu_tar)
             IMES% = FINI% / DIAPER%
             POSIX% = 1 + FINI% * INTERMES% / DIAPER% - INTERMES%
             If POSIX% > Len(STRIVE$) Then GoTo 19
             If POSIX% > 0 Then
               Mid$(STRIVE$, POSIX%, 1) = "X"
             End If
             GoTo 12
           End If
         End If
19        .MoveNext
       Loop
     End With
     '
   End If
   '
90 If MODO% = 2 Then STRIVE$ = STRIVEFE$
   STRIVENCI$ = Mid$(STRIVE$, 1 + (NANO% - 1) * 12 * INTERMES%, 12 * INTERMES%)
   '
End Function

Private Sub Command18_Click()
    Command18.Enabled = False
    '
    
    'VALIDACIONES
    CODIMAQ$ = TRIM$(Text1.TEXT)
    If CODIMAQ$ = "" Then
       Call COMUNI("Código Inexistente o No Válido")
       GoTo 99
    End If
    '
    'FALTA ARMAR EL ARVHIVO DE PERMISO DE USO
    If VERILICEN%("ACTIFIJO", "Sistema de Activo Fijo") > 0 Then
       
        CONDI$ = "CODIGO_EQU =  '" & CODIMAQ$ & "' AND ImpVta_Equ > 0"""
        If CantRegistros("EQUIPOS", CONDI$) > 0 Then
            Call COMUNI("Imposible Modificar, Bien Informado como Vendido \ Para Modificar Deberá Revertir el Informe de Venta")
        End If
        '
        If DERACCE%("ACTIVOFI", "Ingreso de Datos de Activos Fijos") > 1 Then
            '
            Call STDFORM("ACTIFIJO", CODIMAQ$)
            '
            If STDSCREEN.APRO = "OK" Then
              Call CALCUVALRES(CODIMAQ$)
            End If
        End If
            '
            '
    End If
    
    '
99  Command18.Enabled = True
End Sub

Private Sub Command19_Click()
   '
10 Call SELECHO("INDIEQUI")
   COD$ = TRIM$(VALACT1$("INDIEQUI"))
   If COD$ <> "" Then
      Text24.TEXT = COD$
      MODIFIC% = 1
   End If
   '
End Sub

'
Private Sub Command2_Click()
  '
  If COMALTER%("Realmente Desea Descartar los Cambios ?\\No, Continuar\Si, Descartar Cambios") <> 2 Then
     Exit Sub
  End If
  '
  Command2.Enabled = False
    A$ = TRIM$(Text1.TEXT)
    Text1.TEXT = ""
    Text1.TEXT = A$
  Command2.Enabled = True
  '
End Sub

Private Sub Command20_Click()
    '
    Command20.Enabled = False
    'VALIDACIONES
    CODIMAQ$ = TRIM$(Text1.TEXT)
    If CODIMAQ$ = "" Then
       Call COMUNI("Código Inexistente o No Válido")
       GoTo 99
    End If
    '
    'FALTA ARMAR EL ARVHIVO DE PERMISO DE USO
    If VERILICEN%("ACTIFIJO", "Sistema de Activo Fijo") > 0 Then
      
       If DERACCE%("VERVARES", "Ver Valor Residual de Activos Fijos") < 2 Then
          GoTo 99
       End If
       Call CALCUVALRES(CODIMAQ$)
       Call STDFORM("VALORESI", CODIMAQ$)
    End If
    '
99  Command20.Enabled = True
    '
End Sub

Private Sub Command21_Click()
    '
    Command21.Enabled = False
    VDEF$ = ""
10  OBJ1$ = OBJPLA$("NUEQUI", VDEF$)
    If TRIM$(OBJ1$) <> "" Then
      VDEF$ = OBJ1$
      'VALIDACIONES
      CODLAR$ = Mid$(OBJ1$, 47, 16)
      CODINUE$ = Left$(OBJ1$, 6)
      If Len(TRIM(CODINUE$)) < 3 Then
         Call MENSERR(24, "Código Alias no Válido.\Longitud Código Menor que 3 Caracteres.")
         GoTo 10
      End If
      '
      If Len(TRIM(CODLAR$)) < 3 Then
         Call MENSERR(24, "Código Principal no Válido.\Longitud Código Menor que 3 Caracteres.")
         GoTo 10
      End If
      '
      DESCRI1$ = Mid$(OBJ1$, 7, 40)
      If TRIM$(DESCRI1$) = "" Then
         Call MENSERR(24, "Descripción no Válida.")
         GoTo 10
      End If
      '
      CONDI$ = "Codigo_Equ = '" & CODINUE$ & "' "
      CONDI$ = CONDI$ + " OR CodLar_Equ = '" & CODLAR$ & "' """
      If CantRegistros("EQUIPOS", CONDI$) > 0 Then
         Call MENSERR(24, "Códigos Pre-Existentes.")
         GoTo 10
      End If
      '
      Call CreaRegistro("EQUIPOS", "Codigo_Equ", CODINUE$)
      Call ACTUREGISTRO("EQUIPOS", "CodLar_Equ", "Codigo_Equ = '" & CODINUE$ & "' ", CODLAR$, RAFE&)
      Call ACTUREGISTRO("EQUIPOS", "Descri_Equ", "Codigo_Equ = '" & CODINUE$ & "' ", DESCRI1$, RAFE&)
      '
      'REFRESCA LISTA INDICE DE EQUIPOS
      Call INDIEQUI
      Call FRESHECHO("INDIEQUI")
      '
      Text1.TEXT = CODINUE$
      '
    End If
    Command21.Enabled = True
    '
End Sub

Private Sub Command22_Click()
  '
   Dim FIELDVAL(256)
   '
   Command22.Enabled = False
   'VALIDACIONES PRE-COPIA
   If TRIM$(Text1.TEXT) = "" Then
     Call MENSERR(24, "Código de Equipo no Válido")
     GoTo 99
   End If
   If TRIM$(ECOARCH$("INDIEQUI", Text1.TEXT)) = "" Then
     Call MENSERR(24, "Código de Equipo no Válido")
     GoTo 99
   End If
   '
   CODIVIE$ = AJUSTI$(Text1.TEXT, 6)
   DESCRIVIE$ = ECOARCH$("INDIEQUI", Text1.TEXT)
   '
   VDEF$ = Space$(256)
   Call REPLA(VDEF$, " ", 1, 6)
   Call REPLA(VDEF$, DESCRIVIE$, 7, 56)
   '
10 OBJ1$ = OBJPLA$("NUEQUI", VDEF$)
   If OBJ1$ <> "" Then
     VDEF0$ = OBJ1$
     'VALIDACIONES
     CODINUE$ = Left$(OBJ1$, 6)
      If TRIM(CODINUE$) = "" Then
         Call MENSERR(24, "Código no Válido o Pre-Existente.")
         VDEF$ = VDEF0$
         GoTo 10
      End If
      '
      If Len(CODINUE$) < 3 Then
         Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
         VDEF$ = VDEF0$
         GoTo 10
      End If
      '
      If Len(CODINUE$) > 6 Then
         Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 6 Caracteres.")
         VDEF$ = VDEF0$
         GoTo 10
      End If
      '
      
      Call ABREMAQUI
      
'      Dim Equipos As ADODB.Recordset
'      Dim ConsSQL As String
'
'      Set Equipos = New ADODB.Recordset
'      ConsSQL = "SELECT * FROM EQUIPOS WHERE Codigo_Equ = '" & CODINUE$ & "' "
'      Equipos.Open FILTSQL$(ConsSQL), FLEXCONN, adOpenKeyset, adLockPessimistic
'
'
'      With Equipos ' COPIO TABLA EQUIPOS
'        If .BOF = False And .EOF = False Then
'          Call MENSERR(24, "Código no Válido o Pre-Existente.")
'          VDEF$ = VDEF0$
'          GoTo 10
'        End If
'      End With
      '
      CONDI$ = "Codigo_Equ = '" & CODINUE$ & "'"
      If CantRegistros&("EQUIPOS", CONDI$) > 0 Then
         Call MENSERR(24, "Código Pre-Existente.")
         VDEF$ = VDEF0$
         GoTo 10
      End If

      'SE AGREGO EL VALOR 40 POR QUE COPIABA EL ALIAS EN LA DESCRIPCION
      DESCRI1$ = Mid$(OBJ1$, 7, 40)
      If TRIM$(DESCRI1$) = "" Then
         Call MENSERR(24, "Descripción no Válida.")
         VDEF$ = VDEF0$
         GoTo 10
      End If
      '
      Alias$ = TRIM$(Mid$(OBJ1$, 47, 16))
      If Alias$ = "" Then
         Call MENSERR(24, "Alias no Válido.")
         GoTo 10
      End If
      '
      CONDI$ = "CodLar_Equ = '" & Alias & "'"
      If CantRegistros&("EQUIPOS", CONDI$) > 0 Then
         Call MENSERR(24, "Código Alias Pre-Existente.")
         VDEF$ = VDEF0$
         GoTo 10
      End If
      '
'      Call ABREMAQUI
'      CACAM% = Equipos.Fields.Count
      Dim Equipos As New ADODB.Recordset
      Dim ConsSQL As String

      ConsSQL = "SELECT * FROM EQUIPOS WHERE Codigo_Equ = '" & CODIVIE$ & "' "
'      Equipos.Close
      Equipos.Open FILTSQL$(ConsSQL), FLEXCONN, adOpenKeyset, adLockPessimistic
      
      With Equipos
      If .BOF = False And .EOF = False Then
        For KU% = 1 To .Fields.Count
          FIELDVAL(KU%) = .Fields(KU% - 1).Value
        Next KU%
      End If
      '
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !codigo_equ = CODINUE$
        !Descri_Equ = Left(DESCRI1$, 64)
        !codlar_equ = Left(Alias$, 16) ' AGREGADO PARA QUE TOME EL QUE SE INGRESA EN EL OBJETO
      .Update
      '
       ConsSQL = "SELECT * FROM EQUIPOS WHERE Codigo_Equ = '" & CODINUE$ & "' "
       Equipos.Close
       Equipos.Open FILTSQL$(ConsSQL), FLEXCONN, adOpenKeyset, adLockPessimistic
    
       
      If .BOF = False And .EOF = False Then
         'COPIO TABLA EQUIPOS
         For KU% = 1 To .Fields.Count
           NOCAMPO$ = TRIM$(UCase$(Equipos.Fields(KU% - 1).Name))
           If NOCAMPO$ <> "CODIGO_EQU" Then
            If NOCAMPO$ <> "DESCRI_EQU" Then
             If NOCAMPO$ <> "CODLAR_EQU" Then 'AGREGADO PARA QUE NO SOBRE ESCRIBA DEL VIEJO
                .Fields(KU% - 1).Value = FIELDVAL(KU%)
             End If
            End If
           End If
           
         Next KU%
         .Update
      End If
      '
      End With
      '
20    SQLX$ = "SELECT * FROM REPUMAQ "
      SQLX$ = SQLX$ + " WHERE Codigo_Equ = '" & CODIVIE$ & "' "
      Dim RSTREPU As ADODB.Recordset
      Set RSTREPU = New ADODB.Recordset
      RSTREPU.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
      
      Dim RepuMaq As ADODB.Recordset
      Set RepuMaq = New ADODB.Recordset
      SQLXAux$ = "SELECT * FROM REPUMAQ "
     ' SQLXAux$ = SQLXAux$ + " WHERE Codigo_Equ = '" & CODINUE$ & "' "
      RepuMaq.Open FILTSQL$(SQLXAux$), FLEXCONN, adOpenKeyset, adLockPessimistic
      
      With RSTREPU
        On Error Resume Next
        .MoveFirst
        On Error GoTo 0
        If Err Then
          On Error GoTo 0
          GoTo 30
        End If
        On Error GoTo 0
        Do While .EOF = False
          With RepuMaq 'COPIO TABLA DE REPUESTOS 'REVISAR NO TOMABA RepuMaq
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
              !codigo_equ = CODINUE$
              
              For J% = 1 To RSTREPU.Fields.Count
                NOCAMPO$ = TRIM$(UCase$(RepuMaq.Fields(J% - 1).Name))
                If NOCAMPO$ <> "CODIGO_EQU" Then
                 If NOCAMPO$ <> "CODLAR_EQU" Then
                  On Error Resume Next
                   
                    .Fields(J% - 1).Value = RSTREPU.Fields(J% - 1).Value
                  On Error GoTo 0
                 End If
                End If
              Next J%
            .Update
          End With
          .MoveNext
        Loop
      End With
      '
30    SQLX$ = "SELECT * FROM CONTADORES "
      SQLX$ = SQLX$ + " WHERE Codigo_Equ = '" & CODIVIE$ & "' "
      Dim RSTCONT As ADODB.Recordset
      Set RSTCONT = New ADODB.Recordset
      RSTCONT.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
      
      SQLXAux$ = "SELECT * FROM CONTADORES "
      'SQLXAux$ = SQLXAux$ + " WHERE Codigo_Equ = '" & CODINUE$ & "' "
      Dim Contadores As ADODB.Recordset
      Set Contadores = New ADODB.Recordset
      Contadores.Open FILTSQL$(SQLXAux$), FLEXCONN, adOpenKeyset, adLockPessimistic

      
      
      With RSTCONT
        On Error Resume Next
        .MoveFirst
        On Error GoTo 0
        If Err Then
          On Error GoTo 0
          GoTo 40
        End If
        On Error GoTo 0
        Do While .EOF = False
          With Contadores 'COPIO TABLA CONTADORES
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
              !codigo_equ = CODINUE$
              !Num_Con = RSTCONT.Fields("NUM_CON")
              !Referen_Con = RSTCONT.Fields("REFEREN_CON")
              !Freq_Con = RSTCONT.Fields("FREQ_CON")
              !Observa_Con = " "
            .Update
          End With
          .MoveNext
        Loop
      End With
      '
      
40    'Selecciono imagenes con el codigo viejo
      SQLX$ = "SELECT * FROM GALEIMAG "
      SQLX$ = SQLX$ + " WHERE Codigo_Equ = '" & CODIVIE$ & "' "
      Dim RSTIMG As ADODB.Recordset
      Set RSTIMG = New ADODB.Recordset
      RSTIMG.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
      
      'creo recordset para copiarlas con codigo Nuevo
      SQLXAux$ = "SELECT * FROM GALEIMAG "
      'SQLXAux$ = SQLXAux$ + " WHERE Codigo_Equ = '" & CODINUE$ & "' "
      Dim GaleImag As ADODB.Recordset
      Set GaleImag = New ADODB.Recordset
      GaleImag.Open FILTSQL$(SQLXAux$), FLEXCONN, adOpenKeyset, adLockPessimistic
      
      
      With RSTIMG
        On Error Resume Next
        .MoveFirst
        On Error GoTo 0
        If Err Then
          On Error GoTo 0
          GoTo 50
        End If
        On Error GoTo 0
        Do While .EOF = False
          With GaleImag 'COPIO TABLA DE IMAGENES
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
              !codigo_equ = CODINUE$
              !Norit_Imag = RSTIMG.Fields("NORIT_IMAG")
              !Ruta_Imag = RSTIMG.Fields("RUTA_IMAG")
              !Refe_Imag = RSTIMG.Fields("REFE_IMAG")
            .Update
          End With
          .MoveNext
        Loop
      End With
      '
50    Call INDIEQUI
      'Call FRESHECHO("INDIEQUI")
      '
   Else
     Text1.TEXT = CODIVIE$
     GoTo 99
   End If
   '
   Text1.TEXT = CODINUE$
   CARGAEQUI (CODINUE$)
   
99 Command22.Enabled = True
   '
End Sub

Private Sub Command24_Click()
  '
  Command24.Enabled = False
  If TRIM$(Text1.TEXT) <> "" Then
    If ECOARCH$("INDIEQUI", Text1.TEXT) <> "" Then
      Call BAJAMAQ(Text1.TEXT)
    Else
      Call MENSERR(24, "Código de Equipo no Válido")
    End If
  End If
  Command24.Enabled = True
  '
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   If TRIM$(ECOARCH$("INDIEQUI", Text1.TEXT)) <> "" Then
    If MODIFIC% = 1 Then
      If COMALTER%("Se han Modificado Datos del Equipo\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
        GoTo 99
      Else
        Call GRATODO
      End If
    End If
   End If
   '
99 MODIFIC% = 0
   Close: End
'   Unload Me
End Sub

Private Sub Command37_Click()
  '
  On Error Resume Next
  SETGEMA09.Show 1
  '400=El form ya esta mostrado.Imposible mostrar en forma modal
  If Err Then
    If Err.Number = 400 Then
      On Error GoTo 0
      Unload AMAEQUI09
    End If
  End If
  On Error GoTo 0
  '
End Sub

Private Sub Command38_Click()
  '
  Command38.Enabled = False
  Dim RESP%
  RESP% = ALTERNA%("Informacion Basica de Equipos\Listado de Sub-Conjuntos\Listado de Repuestos")
  
  Select Case RESP%
     Case Is = 1
       Call FINAL("?EQUILIST") 'OK
     Case Is = 2
       Call FINAL("?SUBCONJUNTOS") 'OK
     Case Is = 3
       Call FINAL("?EQUIREPU")
  End Select
  
  Command38.Enabled = True
  '
End Sub

Private Sub Command39_Click()
    '
    Command39.Enabled = False
    Me.PRINTFORM
99  Command39.Enabled = True
    '
End Sub

Private Sub Command4_Click()
  '
  Call HELPONLINE("AMAEQUI")
  '
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   NMAQ$ = ECOARCH$("INDIEQUI", Text1.TEXT)
   
   Dim Codigo As String * 6
   Codigo = Text1
   
   If NMAQ$ <> "" Then
     If TRIM$(Text2) <> "" Then
       Call FRATEXTO(Text2, 56)
       For KKI% = 1 To CARETEX%
         TTXX$ = Space$(76)
         Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTXX$, HOS$, 1, 8)
           Call REPLA(TTXX$, USERTER$, 69, 24)
         End If
         
         Call REGAPP("ANOTAMAQ", Left$(Codigo + TTXX$, 96))
       Next KKI%
       Call REGAPP("ANOTAMAQ", Codigo + String$(88, "="))
       Call CIERRARCH("ANOTAMAQ")
       Text2 = ""
       '
       List4.Clear
       RFF$ = RECFILT$("ANOTAMAQ", 1, Text1.TEXT)
       '
       For U& = 1 To Len(RFF$) Step 4
         RENOTA& = CVS(Mid$(RFF$, U&, 4))
         XX$ = REGLEIDO$("ANOTAMAQ", RENOTA&)
         List4.AddItem Mid$(XX$, 7, 92)
       Next U&
       '
     End If
   End If
   Command5.Enabled = True
End Sub
Private Sub Command7_Click()
 If TRIM$(Text1) <> "" Then
  Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text14.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text14.TEXT)
 End If
End Sub

Private Sub Command8_Click()
  '
  If TRIM$(Text1) <> "" Then
    Call SELECHO("PLANTPRO")
    COD$ = VALACT1$("PLANTPRO")
    If TRIM$(COD$) <> "" Then
      Text10.TEXT = COD$
      Label16.Caption = TRIM$(ECOARCH$("PLANTPRO", Text10.TEXT))
    End If
  End If
  '
End Sub

Private Sub Command9_Click(Index As Integer)
   '
   
   If TRIM$(ECOARCH$("INDIEQUI", Text1.TEXT)) <> "" Then
    If MODIFIC% = 1 Then
     If COMALTER%("Se han Modificado Datos del Equipo\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       GoTo 10
     Else
       Call GRATODO
     End If
    End If
   End If
   '
10 Call SELECHO("INDIEQUI")
   COD$ = TRIM$(VALACT1$("INDIEQUI"))
   If COD$ <> "" Then
    Text1.TEXT = COD$
    Call CARGAEQUI(COD$)
    MODIFIC% = 0
   End If
   '
End Sub

Private Sub Command99_Click(Index As Integer)
   '
   If TRIM$(Text1) = "" Then Exit Sub
   If TRIM$(ECOARCH$("INDIEQUI", Text1.TEXT)) = "" Then Exit Sub
   '
   Command99(Index).Enabled = False
   '
10 CommonDialog1.CancelError = True
   CommonDialog1.InitDir = ULODATOS$ & "\"
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Archivo de Imagen(JPG; BMP; WMF)|*.JPG; *.BMP; *.WMF|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   '
   APLANO$ = UCase$(CommonDialog1.FileName)
   If Left$(APLANO$, Len(LUDAT$)) = LUDAT$ Then
      APLANO$ = Mid$(APLANO$, 1 + Len(LUDAT$))
   End If
   Label99(Index) = APLANO$
   '
   On Error GoTo 0
   '
99 Screen.MousePointer = 1
   Command99(Index).Enabled = True
   On Error GoTo 0
End Sub



Private Sub DesCam_Click()
  '
  Call Command2_Click
  '
End Sub

Private Sub Form_LinkClose()
 Unload Me(AMAEQUI09)
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '
  If TRIM$(ECOARCH$("INDIEQUI", Text1.TEXT)) <> "" Then
    If MODIFIC% = 1 Then
      If COMALTER%("Se han Modificado Datos del Equipo\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
        Exit Sub
      Else
        Call GRATODO
      End If
    End If
  End If
  MODIFIC% = 0
  '
 
End Sub


Private Sub guar_Click()
  '
  Call Command1_Click
  '
End Sub

Private Sub Imp_Click()
  '
  Call Command39_Click
  '
End Sub

Private Sub Label27_DBLClick()
  '
  If TRIM(Text1) <> "" Then
    '
    Call BLANARCH("!REPUESTO")
    JJ& = 0
    TA& = List1.ListCount
    For JX% = 0 To List1.ListCount - 1
      L1$ = List1.List(JX%)
      CODI0$ = Mid$(L1$, 1, 16)
        CODI1% = CODINT%(CODI0$)
      DESC1$ = Mid$(L1$, 18, 30)
      UM2$ = Mid$(L1$, 49, 2)
      USO0$ = Mid$(L1$, 52, 10)
        USO1 = Val(USO0$)
      OBS1$ = Mid$(L1$, 63, 25)
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
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + AMAEQUI09.Left + 170))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + AMAEQUI09.Top + 4250))
    VTB% = VENTABU%(PRF$ + ATRI$)
    If VTB% >= 0 Then
      '
      List1.Clear
      '
      FIN& = ULTREG&("!REPUESTO")
      If TA& <> FIN& Then MODIFIC% = 1
      '
      For J& = 1 To FIN&
        X$ = REGLEIDO$("!REPUESTO", J&)
        CODI1% = CVI(Mid$(X$, 1, 2))
        CODI2$ = CODEXT$(CODI1%)
        DESC2$ = DESCRIT0(CODI1%)
        UM2$ = Mid$(X$, 3, 2)
        USO1 = CVS(Mid$(X$, 5, 4))
        USO2$ = FORMATNUM$(USO1, "F9.1")
        OBS2$ = Mid$(X$, 9, 25)
        'AGREGO AL LIST
        L1$ = Space(128)
        Call REPLA(L1$, CODI2$, 1, 16)
        Call REPLA(L1$, DESC2$, 18, 30)
        Call REPLA(L1$, UM2$, 49, 2)
        Call REPLA(L1$, USO2$, 52, 10)
        Call REPLA(L1$, OBS2$, 63, 25)
        List1.AddItem L1$
      Next J&
      '
    End If
  End If
  '
End Sub


Private Sub Label4_DblClick()
  '
  If TRIM(Text1) <> "" Then
   Call BLANARCH("!SUBCONJU")
    JJ& = 0
    For JX% = 0 To List2.ListCount - 1
      L1$ = List2.List(JX%)
      NRO1% = Val(Mid$(L1$, 1, 3))
      REFE1$ = Mid$(L1$, 5, 22)
      MAR1$ = Mid$(L1$, 27, 24)
      MOD1$ = Mid$(L1$, 51, 20)
      SER1$ = Mid$(L1$, 71, 20)
      CRIT$ = Mid$(L1$, 91, 2)
      '
      Y$ = REGBLAN$("SUBCONJU")
      Call REPLA(Y$, MKI$(NRO1%), 1, 2)
      Call REPLA(Y$, REFE1$, 3, 32)
      Call REPLA(Y$, MAR1$, 35, 24)
      Call REPLA(Y$, MOD1$, 59, 20)
      Call REPLA(Y$, SER1$, 79, 20)
      Call REPLA(Y$, CRIT$, 99, 2)
      JJ& = JJ& + 1
      Call GRAREG("!SUBCONJU", Y$, JJ&)
    Next JX%
    Call SETULTREG("!SUBCONJU", JJ&)
    Call CIERRARCH("!SUBCONJU")
    '
152 PRF$ = "SUBCONJU"
    ATRI$ = ATRI$ + "/TITUPAN=SubConjuntos de Equipo /"
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + AMAEQUI09.Left + 170))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + AMAEQUI09.Top + 2400))
    VTB% = VENTABU%(PRF$ + ATRI$)
    If VTB% >= 0 Then
      '
      List2.Clear
      '
      FIN& = ULTREG&("!SUBCONJU")
      For J& = 1 To FIN&
        X$ = REGLEIDO$("!SUBCONJU", J&)
        NRO% = CVI(Mid$(X$, 1, 2))
        If NRO% <= 0 Then
         Call MENSERR(24, "Número de SubConjunto no Válido")
         GoTo 152
        End If
          NRO2$ = FORMATNUM$(NRO%, "F3.0")
        REFE$ = Mid$(X$, 3, 32)
        MAR2$ = Mid$(X$, 35, 24)
        MOD2$ = Mid$(X$, 59, 20)
        SER2$ = Mid$(X$, 79, 20)
        CRIT$ = Mid$(X$, 99, 2)
        'AGREGO AL LIST
        L1$ = Space(256)
        Call REPLA(L1$, NRO2$, 1, 3)
        Call REPLA(L1$, REFE$, 5, 22)
        Call REPLA(L1$, MAR2$, 27, 24)
        Call REPLA(L1$, MOD2$, 51, 20)
        Call REPLA(L1$, SER2$, 71, 20)
        Call REPLA(L1$, CRIT$, 91, 4)
        List2.AddItem L1$
      Next J&
      '
    End If
  End If
  '
End Sub

Private Sub Label31_DblClick()
 If TRIM(Text1) <> "" Then
    '
    Call BLANARCH("!CONTADOR")
    JJ& = 0
    For JX% = 0 To List3.ListCount - 1
      L1$ = List3.List(JX%)
      NRO1% = Val(Mid$(L1$, 1, 3))
      REFE1$ = Mid$(L1$, 5, 22)
      FEQ1 = Val(Mid$(L1$, 28, 8))
      FUL1% = FECHANUM(Mid$(L1$, 37, 8))
      VAL1# = Val(Mid$(L1$, 46, 10))
      FPL1% = FECHANUM(Mid$(L1$, 57, 8))
      'CALCULO DE LA FECHA PROXIMA LECTURA
      If FUL1% > 0 And FEQ1 > 0 Then
         F1 = CDbl(CVDAT(FUL1%))
         FP1 = F1 + FEQ1
         FPL1% = CVINT%(FP1)
      End If
      '
      PROME1# = Val(Mid$(L1$, 66, 10))
      OBSE1$ = Mid$(L1$, 77, 64)
      '
      Y$ = REGBLAN$("CONTADOR")
      Call REPLA(Y$, MKI$(NRO1%), 1, 2)
      Call REPLA(Y$, REFE1$, 3, 32)
      Call REPLA(Y$, MKS(FEQ1), 33, 4)
      Call REPLA(Y$, MKI$(FUL1%), 37, 2)
      Call REPLA(Y$, MKD$(VAL1#), 39, 8)
      Call REPLA(Y$, MKI$(FPL1%), 47, 2)
      Call REPLA(Y$, MKD$(PROME1#), 49, 8)
      Call REPLA(Y$, OBSE1$, 57, 64)
      JJ& = JJ& + 1
      Call GRAREG("!CONTADOR", Y$, JJ&)
    Next JX%
    Call SETULTREG("!CONTADOR", JJ&)
    Call CIERRARCH("!CONTADOR")
    '
152 PRF$ = "CONTADOR"
    ATRI$ = ATRI$ + "/TITUPAN=Contadores de Equipo /"
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + AMAEQUI09.Left + 170))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + AMAEQUI09.Top + 2250))
    VTB% = VENTABU%(PRF$ + ATRI$)
    If VTB% >= 0 Then
      '
      List3.Clear
      '
      FIN& = ULTREG&("!CONTADOR")
      For J& = 1 To FIN&
        X$ = REGLEIDO$("!CONTADOR", J&)
        NRO% = CVI(Mid$(X$, 1, 2))
        If NRO% <= 0 Then
         Call MENSERR(24, "Número de Contador no Válido")
         GoTo 152
        End If
          NRO2$ = FORMATNUM$(NRO%, "F3.0")
        REFE$ = Mid$(X$, 3, 32)
        FEQ = CVS(Mid$(X$, 33, 4))
          FEQ2$ = FORMATNUM$(FEQ, "F7.0")
        FUL% = CVI(Mid$(X$, 37, 2))
          FUL2$ = FORMATNUM$(FUL%, "D8")
        'CALCULO DE LA FECHA PROXIMA LECTURA
        If FUL% > 0 And FEQ > 0 Then
          F1 = CDbl(CVDAT(FUL%))
          FP1 = F1 + FEQ
          FPL% = CVINT%(FP1)
        End If
        '
        VALO = CVD(Mid$(X$, 39, 8))
          VALO1$ = FORMATNUM$(VALO, "F10.1")
        'FPL% = CVI(Mid$(X$, 47, 2))
          FPL2$ = FORMATNUM$(FPL%, "D8")
        PROME = CVD(Mid$(X$, 49, 8))
          PROM1$ = FORMATNUM$(PROME, "F10.1")
        OBSE$ = Mid$(X$, 57, 64)
        'AGREGO AL LIST
        L1$ = Space(256)
        Call REPLA(L1$, NRO2$, 1, 3)
        Call REPLA(L1$, REFE$, 5, 22)
        Call REPLA(L1$, FEQ2$, 28, 8)
        Call REPLA(L1$, FUL2$, 37, 8)
        Call REPLA(L1$, VALO1$, 46, 10)
        Call REPLA(L1$, FPL2$, 57, 8)
        Call REPLA(L1$, PROM1$, 66, 10)
        Call REPLA(L1$, OBSE$, 77, 64)
        List3.AddItem L1$
      Next J&
      '
    End If
  End If
  '
End Sub


Private Sub Label99_Change(Index As Integer)
   '
   Static RUTARGRA$, TOPEINI, IZQUINI, ANCHINI, ALTUINI, INTERVER
   '
   Screen.MousePointer = 11
   '
   If INTERVER < 1 Then
     RUTARGRA$ = TRIM$(Control$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
        RUTARGRA$ = RUTARGRA$ + "\"
     End If
     '
     TOPEINI = Picture99(0).Top
     IZQUINI = Picture99(0).Left
     ANCHINI = Picture99(0).Width
     ALTUINI = Picture99(0).Height
     INTERVER = Picture99(1).Top - Picture99(0).Top
   End If
   '
   Picture99(Index).Visible = True
   Refresh
   '
   Image99(Index).Picture = LoadPicture("")
   Picture99(Index).Cls
   COPLANO$ = Label99(Index)
   If COPLANO$ = "" Then GoTo 98
   '
   If EXISTE%(COPLANO$) < 1 Then COPLANO$ = RUTARGRA$ + COPLANO$
   If EXISTE%(COPLANO$) < 1 Then GoTo 98
   '
   On Error GoTo 90
   Picture99(Index).Visible = False
   Image99(Index).Picture = LoadPicture(COPLANO$)
   Picture99(Index).Top = TOPEINI + INTERVER * Index
   Picture99(Index).Left = IZQUINI
   Picture99(Index).Width = ANCHINI
   Picture99(Index).Height = ALTUINI
   Picture99(Index).Picture = LoadPicture("")
   Picture99(Index).Refresh
   '
   HIMAGEN = Image99(Index).Picture.Width
   VIMAGEN = Image99(Index).Picture.Height
   
   HCUADRO = Picture99(Index).Width
   VCUADRO = Picture99(Index).Height
   '
   REDUH = HIMAGEN / HCUADRO
   REDUV = VIMAGEN / VCUADRO
   REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
   Picture99(Index).Height = VCUADRO * REDUV / REDUC
   Picture99(Index).Width = HCUADRO * REDUH / REDUC
   Picture99(Index).Refresh
   '
   HCUADRO = Picture99(Index).Width
   VCUADRO = Picture99(Index).Height
   Picture99(Index).Left = (ANCHINI - HCUADRO) / 2
   '
   Picture99(Index).PaintPicture Image99(Index).Picture, 0, 0, HCUADRO, VCUADRO
   Picture99(Index).Visible = True
   Picture99(Index).Refresh
   MODIFIC% = 1
   On Error GoTo 0
   GoTo 98
   '
90 On Error GoTo 0
   Call MENSERR(24, "Imagen no Válida")
   '
98 Screen.MousePointer = 1
   '
End Sub

Private Sub Label99_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  '
  Label99(Index).ToolTipText = Label99(Index)
  '
End Sub

Private Sub List1_DblClick()
  '
  Call Label27_DBLClick
  '
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'MUESTRA LA DESCRIPCIÓN DEL ITEM SELECCIONADO EN LA LISTA
  CODI1$ = Left$(List1.List(List1.ListIndex), 16)
  CI1% = CODINT%(CODI1$)
  List1.ToolTipText = DESCRIT0$(CI1%)
  '
End Sub

Private Sub List2_DblClick()
  Call Label4_DblClick
End Sub

Private Sub LIST3_DblClick()
  Call Label31_DblClick
End Sub

Private Sub MaeRep_Click()
  '
  Call CONECRUN("AMASTO07", "Lista de Equipos")
  '
End Sub

Private Sub MnuAddImg_Click()
  '
  If TRIM$(Text1) <> "" Then
    If ECOARCH$("INDIEQUI", Text1.TEXT) <> "" Then
      If SSTab1.Tab <> 2 Then SSTab1.Tab = 2
      Call Command13_Click
      On Error Resume Next
        VScroll1.Value = VScroll1.Max
      On Error GoTo 0
    End If
  End If
  '
End Sub

Private Sub mnuAyud_Click()
  '
  Call Command4_Click
  '
End Sub
Private Sub MnuDep_Click()
  '
  Call SETGEMA09.Option1_DblClick(3)
  '
End Sub

Private Sub MnuEspec_Click()
  '
  Call SETGEMA09.Option1_DblClick(4)
  '
End Sub

Private Sub MnuFoImp_Click()
  '
  Call Command37_Click
  '
End Sub

Private Sub MnUniMedi_Click()
  '
  Call SETGEMA09.Option1_DblClick(6)
  '
End Sub

Private Sub mnuPlan_Click()
  '
  Call SETGEMA09.Option1_DblClick(1)
  '
End Sub

Private Sub MnuSecc_Click()
  '
  Call SETGEMA09.Option1_DblClick(2)
  '
End Sub

Private Sub MnuTiMa_Click()
  '
  Call SETGEMA09.Option1_DblClick(0)
  '
End Sub

Private Sub Ordeman_Click()
  '
  ORMANTEN09.Show 1
  '
End Sub

Private Sub Picture99_DBLClick(Index As Integer)
    'MAXIMIZAR IMAGEN
    Screen.MousePointer = 11
    '
    ALTUTITU = SHOWIMG09.Height - SHOWIMG09.ScaleHeight
    ANCHOBOR = SHOWIMG09.Width - SHOWIMG09.ScaleWidth
    SHOWIMG09.Picture2.Top = 0
    SHOWIMG09.Picture2.Left = 0
    SHOWIMG09.Picture2.Width = Screen.Width - ANCHOBOR
    SHOWIMG09.Picture2.Height = Screen.Height - ALTUTITU - 300
    '
    HIMAGEN = Image99(Index).Picture.Width
    VIMAGEN = Image99(Index).Picture.Height
    If HIMAGEN < 1 Or VIMAGEN < 1 Then
       GoTo 99
    End If
    '
    HCUADRO = SHOWIMG09.Picture2.Width
    VCUADRO = SHOWIMG09.Picture2.Height
    '
    REDUH = HIMAGEN / HCUADRO
    REDUV = VIMAGEN / VCUADRO
    REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
    SHOWIMG09.Picture2.Height = VCUADRO * REDUV / REDUC
    SHOWIMG09.Picture2.Width = HCUADRO * REDUH / REDUC
    HCUADRO = SHOWIMG09.Picture2.Width
    VCUADRO = SHOWIMG09.Picture2.Height
    SHOWIMG09.Height = ALTUTITU + SHOWIMG09.Picture2.Height
    SHOWIMG09.Width = ANCHOBOR + SHOWIMG09.Picture2.Width
    If HORIMAGEN + SHOWIMG09.Width > Screen.Width Then HORIMAGEN = 0
    If HORIMAGEN < 1 Then HORIMAGEN = (Screen.Width - SHOWIMG09.Width) / 2
    SHOWIMG09.Left = HORIMAGEN
    If VERTIMAGEN + SHOWIMG09.Height > Screen.Height Then VERTIMAGEN = 0
    If VERTIMAGEN < 1 Then VERTIMAGEN = (Screen.Height - SHOWIMG09.Height) / 2
    SHOWIMG09.Top = VERTIMAGEN - 300
      If SHOWIMG09.Top < 0 Then SHOWIMG09.Top = 0
    SHOWIMG09.Caption = TRIM$(Text1.TEXT) + " - " + TRIM$(Text3) + " - " + TRIM$(Label99(Index))
    '
    SHOWIMG09.AGRAFIC = Label99(Index)
    SHOWIMG09.EQUIP = TRIM$(Text1.TEXT) + " - " + TRIM$(Text3) '+ " - " + TRIM$(Label99(Index))
    '
    Screen.MousePointer = 1
    '
    INDIMAGEN% = Index
    SHOWIMG09.Picture2.PaintPicture Image99(Index).Picture, 0, 0, 1 * HCUADRO, 1 * VCUADRO
    SHOWIMG09.Picture2.Visible = True
    SHOWIMG09.Show 1
    '
99  Screen.MousePointer = 1
    '
End Sub


Private Sub Picture99_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Call INVICTRL
    IMGACT% = Index
    PopupMenu Edi
    Call INVICTRL
  End If
End Sub
Sub INVICTRL()
  '
  DesCam.Visible = Not (DesCam.Visible)
  phl4.Visible = Not (phl4.Visible)
  BaLog.Visible = Not (BaLog.Visible)
  php16.Visible = Not (php16.Visible)
  MnuMaxImg.Visible = Not (MnuMaxImg.Visible)
  MnuDelImg.Visible = Not (MnuDelImg.Visible)
  '
End Sub
Private Sub MNUMAXIMG_CLICK()
  '
  Call Picture99_DBLClick(IMGACT%)
  '
End Sub
Private Sub MnuDelImg_Click()
  '
  'ELIMINAR LA IMAGEN DE LA BASE DE DATOS
  CE1$ = Text1.TEXT
  Call ABREMAQUI
  Call BORRAREGISTROS("GALEIMAG", "Norit_Imag = " & IMGACT%, REG&)
  'RENUMERO LAS IMAGENES
  SQLX$ = "SELECT * FROM GALEIMAG "
  SQLX$ = SQLX$ + "WHERE Codigo_Equ = '" & CE1$ & "' "
  SQLX$ = SQLX$ + "AND Norit_Imag > " & IMGACT% & " "
  Dim RSTIMG As New ADODB.Recordset
  RSTIMG.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
  With RSTIMG
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      GoTo 75
    End If
    On Error GoTo 0
    Do Until .EOF = True
        !Norit_Imag = !Norit_Imag - 1
      .Update
      .MoveNext
    Loop
    
  End With
  'DESCARGO Y VUELVO A CARGAR LAS IMAGENES
75 Call DECARIMG
   Call CARGAIMAG(CE1$)
  '
99 End Sub
Private Sub Sal_Click()
  '
  Call Command3_Click
  '
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  '
  If TRIM$(ECOARCH$("INDIEQUI", Text1.TEXT)) <> "" Then
    If MODIFIC% = 1 Then
      If COMALTER%("Se han Modificado Datos del Equipo\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
        Exit Sub
      Else
        Call GRATODO
      End If
    End If
  End If
  '
End Sub
Sub GRATODO()
  '
  Call GRABEQ
  Call INDIEQUI
  Call GRAIMG
  Call GRAREP
  Call GRACON
  Call GRASUB
  '
End Sub
Private Sub Text1_Change()
  '
  Call BLANFORMU
  Call DECARIMG 'DESCARGA IMAGENES
  If TRIM$(Text1) <> "" Then
    If CantRegistros("EQUIPOS", "Codigo_Equ = '" & TRIM$(Text1) & "'") > 0 Then
      MAX1& = 5
      Call TRACE(20, 1, MAX1&)
      Call CARGAEQUI(Text1.TEXT) 'CARGA EQUIP
      Call TRACE(20, 2, MAX1&)
      Call CARGAIMAG(UCase(Text1.TEXT)) 'CARGA IMAGENES
      Call TRACE(20, 3, MAX1&)
      Call CARREP(Text1.TEXT) 'CARGA REPUESTOS
      Call TRACE(20, 4, MAX1&)
      Call CARCONTA(Text1.TEXT) 'CARGA CONTADORES
      Call TRACE(20, 5, MAX1&)
      Call CARSUBCON(Text1.TEXT) 'CARGA SUBCONJUNTOS
      YACAR% = 1
      MODIFIC% = 0
    End If
  End If
  '
End Sub
Private Sub Form_Load()
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    Call APLICACIONSKINS(Me)

    SSTab1.Tab = 0
    Call CENTRAFORM(Me)
    Refresh
    Call CARTAVALISTO
    '
    Picture4.Height = Picture3.Height
    VScroll1.Min = 0
    VScroll1.Max = (Picture4.Height - Picture3.Height) / 10 + 1
    VScroll1.Value = 0
    '
    INTERVER = Picture99(1).Top - Picture99(0).Top - Picture99(0).Height
    '
    HOII% = HOY(HOS$)
    Label22 = " " + HOS$
    Label19 = "  " + HOS$
    Label23 = Left$(USERNAME, 17)
    MODIFIC% = 0
    '
    Call ABREMAQUI
    '
End Sub
Sub CARTAVALISTO()
    '
10  If ULTREG&("SUCURSAL") < 1 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Falta Definir Depósitos")
        SETGEMA04.Show 1
    End If
    '
    If ULTREG&("UMEDIDA") < 2 Then
        Call BLANARCH("UMEDIDA")
        Call REGAPP("UMEDIDA", "U.Unidades")
        Call REGAPP("UMEDIDA", "KgKilogramos")
        Call REGAPP("UMEDIDA", "grGramos")
        Call REGAPP("UMEDIDA", "gKGramos - Kilos")
        Call REGAPP("UMEDIDA", "M.Metros")
        Call REGAPP("UMEDIDA", "M2M.Cuadrados")
        Call REGAPP("UMEDIDA", "M3M.Cúbicos")
        Call REGAPP("UMEDIDA", "mmMilímetros")
        Call REGAPP("UMEDIDA", "cmCentímetros")
        Call REGAPP("UMEDIDA", "LtLitros")
        Call CIERRARCH("UMEDIDA")
    End If
    XX$ = REGLEIDO$("UMEDIDA", 1)
    '
    If ULTREG&("UNIMED") < 2 Then
        Call BLANARCH("UNIMED")
        Call REGAPP("UNIMED", "U.  Unidades" + String$(256, 0))
        Call REGAPP("UNIMED", "Kg  Kilogramos" + String$(256, 0))
        Call REGAPP("UNIMED", "gr  Gramos" + String$(256, 0))
        Call REGAPP("UNIMED", "gK  Gramos - Kilos" + String$(256, 0))
        Call REGAPP("UNIMED", "M.  Metros" + String$(256, 0))
        Call REGAPP("UNIMED", "M2  M.Cuadrados" + String$(256, 0))
        Call REGAPP("UNIMED", "M3  M.Cúbicos" + String$(256, 0))
        Call REGAPP("UNIMED", "mm  Milímetros" + String$(256, 0))
        Call REGAPP("UNIMED", "cm  Centímetros" + String$(256, 0))
        Call REGAPP("UNIMED", "Lt  Litros" + String$(256, 0))
        Call CIERRARCH("UNIMED")
    End If
    XX$ = REGLEIDO$("UNIMED", 1)
    '
    If ULTREG&("TATIMAT") < 1 Then
        Call BLANARCH("TATIMAT")
        Call REGAPP("TATIMAT", Chr$(1) + "Materia Prima")
        Call REGAPP("TATIMAT", Chr$(2) + "Pieza Comercial")
        Call REGAPP("TATIMAT", Chr$(3) + "Pieza Elaborada")
        Call REGAPP("TATIMAT", Chr$(4) + "Semielaborado")
        Call REGAPP("TATIMAT", Chr$(5) + "Prod.Terminado")
        Call REGAPP("TATIMAT", Chr$(8) + "Servicios")
        Call REGAPP("TATIMAT", Chr$(9) + "Improductivo")
        Call CIERRARCH("TATIMAT")
    End If
    XX$ = REGLEIDO$("TATIMAT", 1)
    '
    If ULTREG&("TABTIMA") < 1 Then
       DPPRE% = Asc(Left$(REGLEIDO$("ECODEP", 1), 1))
       For U& = 1 To ULTREG&("TATIMAT")
          X$ = REGLEIDO$("TATIMAT", U&)
          Y$ = REGBLAN$("TABTIMA")
          Call REPLA(Y$, X$, 1, 16)
          Call REPLA(Y$, Chr$(DPPRE%), 17, 1)
          Call REGAPP("TABTIMA", Y$)
       Next U&
       Call CIERRARCH("TABTIMA")
    End If
    XX$ = REGLEIDO$("TABTIMA", 1)
    '
    If ULTREG&("TACRIRE") < 4 Then
        Call BLANARCH("TACRIRE")
        Call REGAPP("TACRIRE", Chr$(1) + "Por Pedido")
        Call REGAPP("TACRIRE", Chr$(2) + "Por Mínimos")
        Call REGAPP("TACRIRE", Chr$(3) + "Cons.Estadístico")
        Call REGAPP("TACRIRE", Chr$(4) + "Programación")
        Call CIERRARCH("TACRIRE")
    End If
    XX$ = REGLEIDO$("TACRIRE", 1)
    '
    If ULTREG&("TAMONED") < 2 Then
        Call BLANARCH("TAMONED")
        Call REGAPP("TAMONED", Chr$(1) + AJUSTI$("Pesos", 27) + MKS$(1))
        Call REGAPP("TAMONED", Chr$(2) + AJUSTI$("Dolares U$S", 27) + MKS$(1))
        Call REGAPP("TAMONED", Chr$(3) + AJUSTI$("Coef.Est.Ref.", 27) + MKS$(1))
        Call CIERRARCH("TAMONED")
    End If
    XX$ = REGLEIDO$("TAMONED", 1)
    '
    If ULTREG&("TACOMPO") < 2 Then
        Call BLANARCH("TACOMPO")
        Call REGAPP("TACOMPO", Chr$(0) + "Manual")
        Call REGAPP("TACOMPO", Chr$(1) + "Automática")
        Call CIERRARCH("TACOMPO")
    End If
    XX$ = REGLEIDO$("TACOMPO", 1)
    '
    If ULTREG&("GRUPIVA") < 1 Then
        Call SETULTREG("GRUPIVA", 0)
        Call REGAPP("GRUPIVA", Chr$(0) + "IVA 21 %   " + MKS$(21))
        Call REGAPP("GRUPIVA", Chr$(1) + "IVA 10,5 % " + MKS$(10.5))
        Call REGAPP("GRUPIVA", Chr$(2) + "IVA 27 %   " + MKS$(27))
        Call CIERRARCH("GRUPIVA")
    End If
    XX$ = REGLEIDO$("GRUPIVA", 1)
    '
    Screen.MousePointer = 1
    '
End Sub
'
Sub INDIEQUI() 'INDICE DE EQUIPOS Y MAQUINARIAS
 Call ABREMAQUI
 Call CARGALISEL("INDIEQUI", "EQUIPOS", "CODIGO_EQU/A6;DESCRI_EQU/A40;CodLar_Equ/A16", "CODIGO_EQU <> '' ORDER BY CODIGO_EQU ASC")
 '
99 End Sub
'
Sub CARGAEQUI(EQ1$)
  '
  Dim ConsSQL As String
  Dim Equipos As ADODB.Recordset
  
  ConsSQL = "SELECT * FROM EQUIPOS WHERE codigo_equ = '" & EQ1$ & "' "
  Set Equipos = FLEXCONN.Execute(FILTSQL$(ConsSQL))

  Call ABREMAQUI
  With Equipos
  
   If Not (.BOF = True And .EOF = True) Then
     On Error Resume Next
      Text3 = !Descri_Equ
      TEXT21 = !codlar_equ
      Text13 = FECHATEX$(CVINT%(!Finst_Equ))
      Text14 = FECHATEX$(CVINT%(!Fgara_Equ))
      Text16 = FECHATEX$(CVINT%(!Ffuse_Equ))
      Text23 = !Marca_Equ
      Text4(1) = !Mode_Equ
      Text6 = !NuSeRi_Equ
      Text8 = !Fabri_Equ
      Text35 = !Represe_Equ
      SEPRO1 = !Ubica_Equ
      Text12 = FORMATNUM$(SEPRO1, "F7.3")
      Text10 = !Observa_Equ
      Text22 = FORMATNUM$(XVALO(!Cuot_Hora), "F10.2")
      Text24 = SAFETEXT$(!Equi_Alter)
      '
      Text3 = !Descri_Equ
      Text23 = !Marca_Equ
      Text4(1) = !Mode_Equ
      Text6 = !NuSeRi_Equ
      Text9 = !Fabri_Equ
      If Val(!Represe_Equ) > 0 Then
      Text7 = !Represe_Equ
      End If
        SEPRO1 = !Ubica_Equ
      Text12 = FORMATNUM$(SEPRO1, "F7.3") 'SECCIÓN PRODUCTIVA
      Text16 = !Observa_Equ
      Text13 = FECHATEX$(CVINT%(!Finst_Equ))
      Text14 = FECHATEX$(CVINT%(!Fgara_Equ))
      Text8 = !Alimenta_Equ
      Text11 = !Potenci_Equ
      Text15 = !Linea_Equ
      Text17 = !Posi_Equ
      Text10 = !Planta_Equ
        KI1 = !Kilos_Equ
      Text5.TEXT = FORMATNUM$(KI1, "F9.3")
        LA1 = !Largo_Equ
      Text18.TEXT = FORMATNUM$(LA1, "F9.3")
        AN1 = !Ancho_Equ
      Text19.TEXT = FORMATNUM$(AN1, "F9.3")
        AL1 = !Alto_Equ
      Text20.TEXT = FORMATNUM$(AL1, "F9.3")
      '
     On Error GoTo 0
    End If
  End With
  '
End Sub
Sub CARGAIMAG(EQ1$) 'CARGA IMAGENES ASOCIADAS AL EQUIPO
  '
  Call ABREMAQUI
  IMGTOT% = CantRegistros("GALEIMAG", "Codigo_Equ = '" & UCase(EQ1$) & "'")
  If IMGTOT% = 0 Then
     GoTo 99
  End If
  SQLX$ = " SELECT * FROM GALEIMAG "
  SQLX$ = SQLX$ + " WHERE Codigo_Equ = '" & UCase(EQ1$) & "' "
  SQLX$ = SQLX$ + " ORDER BY Norit_Imag ASC "
  Dim IMGTMP As ADODB.Recordset
  Set IMGTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
  
  With IMGTMP
  
   'On Error Resume Next
   '.MoveFirst
   ' If Err Then
   '   On Error GoTo 0
   '   GoTo 99
   ' End If
   
    On Error GoTo 0
    'CREO LAS IMAGENES SI FALTAN
    If IMGTOT% > 2 Then
      For JX% = 1 To (IMGTOT% - 2)
        Call Command13_Click
      Next JX%
    End If
    'CARGO LAS IMAGENES
    Do While .EOF = False
        NOIT% = !Norit_Imag
        Text99(NOIT%) = !Refe_Imag
        Label99(NOIT%) = !Ruta_Imag
        Label99(NOIT%).Refresh
      .MoveNext
    Loop
  End With
  '
99 End Sub
Sub CARREP(MAQ1$) 'CARGA LOS REPUESTOS ASOCIADOS AL EQUIPO
  '
  Call ABREMAQUI
  '
   SQLX$ = " SELECT * FROM REPUMAQ "
   SQLX$ = SQLX$ + " WHERE Codigo_Equ = '" & MAQ1$ & "' "
   
   Dim REPTMP As ADODB.Recordset
   Set REPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
  
   
   With REPTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Sub
    End If
    On Error GoTo 0
    '
    List1.Clear
    '
    Do While .EOF = False
        CODI1% = XVALO(!CodInte_Rep)
        CODI2$ = SAFETEXT$(!CodExte_Rep)
        DESC2$ = SAFETEXT$(!Descrip_Rep)
        UM2$ = SAFETEXT$(!UM_Rep)
        USO1 = XVALO(!Cant_Rep)
        USO2$ = FORMATNUM$(USO1, "F9.1")
        OBS2$ = SAFETEXT$(!Observa_Rep)
        'AGREGO AL LIST
        L1$ = Space(128)
        Call REPLA(L1$, CODI2$, 1, 16)
        Call REPLA(L1$, DESC2$, 18, 30)
        Call REPLA(L1$, UM2$, 49, 2)
        Call REPLA(L1$, USO2$, 52, 10)
        Call REPLA(L1$, OBS2$, 63, 25)
        List1.AddItem L1$
      .MoveNext
    Loop
   End With
  '
End Sub
Sub CARSUBCON(MAQ1$)
  'CARGO SUBCONJUNTOS DEL EQUIPO
  '
  Call ABREMAQUI
  '
  SQLX$ = " SELECT * FROM SUBCONJUNTOS "
  SQLX$ = SQLX$ + " WHERE Codigo_Equ = '" & MAQ1$ & "' "
  Dim REPTMP As ADODB.Recordset
  Set REPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
  
   With REPTMP
   On Error Resume Next
   .MoveFirst
   If Err Then
     On Error GoTo 0
     Exit Sub
   End If
   On Error GoTo 0
   '
   List2.Clear
   '
   Do While .EOF = False
       NRO% = XVALO(!NRO_SUB)
       NRO2$ = FORMATNUM$(NRO%, "F3.0")
       REFE$ = SAFETEXT$(!DESCRI_SUB)
       MAR2$ = SAFETEXT$(!Marca_Sub)
       MOD2$ = SAFETEXT$(!Modelo_Sub)
       SER2$ = SAFETEXT$(!NSerie_Sub)
       CRIT$ = SAFETEXT$(!Criticidad)
       'AGREGO AL LIST
       L1$ = Space(256)
       Call REPLA(L1$, NRO2$, 1, 3)
       Call REPLA(L1$, REFE$, 5, 22)
       Call REPLA(L1$, MAR2$, 27, 24)
       Call REPLA(L1$, MOD2$, 51, 20)
       Call REPLA(L1$, SER2$, 71, 20)
       Call REPLA(L1$, CRIT$, 91, 2)
       List2.AddItem L1$
     .MoveNext
   Loop
  End With
  '
End Sub
Sub CARCONTA(MAQ1$)
  'CARGO ANOTACIONES
   List4.Clear
   RFF$ = RECFILT$("ANOTAMAQ", 1, Text1.TEXT)
   '
   For U& = 1 To Len(RFF$) Step 4
    RENOTA& = CVS(Mid$(RFF$, U&, 4))
    XX$ = REGLEIDO$("ANOTAMAQ", RENOTA&)
    List4.AddItem Mid$(XX$, 7, 92)
   Next U&
   '
  'CARGO CONTADORES DEL EQUIPO
   '
  Call ABREMAQUI
  '
   SQLX$ = " SELECT * FROM CONTADORES "
   SQLX$ = SQLX$ + " WHERE Codigo_Equ = '" & MAQ1$ & "' "
   Dim REPTMP As ADODB.Recordset
   Set REPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
  
   With REPTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Sub
    End If
    On Error GoTo 0
    '
    List3.Clear
    '
    Do While .EOF = False
      NRO% = !Num_Con
        NRO1$ = FORMATNUM$(NRO%, "F3.0")
      REFE$ = !Referen_Con
      FEQ = !Freq_Con
        FEQ1$ = FORMATNUM$(FEQ, "F7.0")
      FUL% = CVINT(!FUL_Con)
        FUL1$ = FORMATNUM$(FUL%, "D8")
      VALO = !Val_Con
        VALO1$ = FORMATNUM$(VALO, "F10.1")
      FPL% = CVINT(!FPL_Con)
        FPL1$ = FORMATNUM$(FPL%, "D8")
      PROME = !Prom_Con
        PROM1$ = FORMATNUM$(PROME, "F10.1")
      OBSE$ = !Observa_Con
      'AGREGO AL LIST
      L1$ = Space(256)
      Call REPLA(L1$, NRO1$, 1, 3)
      Call REPLA(L1$, REFE$, 5, 22)
      Call REPLA(L1$, FEQ1$, 28, 8)
      Call REPLA(L1$, FUL1$, 37, 8)
      Call REPLA(L1$, VALO1$, 46, 10)
      Call REPLA(L1$, FPL1$, 57, 8)
      Call REPLA(L1$, PROM1$, 66, 10)
      Call REPLA(L1$, OBSE$, 77, 64)
      List3.AddItem L1$
      .MoveNext
    Loop
   End With
  '
99 End Sub
Sub GRABEQ()
  'VALIDACIONES
  '
'  If YACAR% < 1 Or MODIFIC% < 1 Then
'     Exit Sub
'  End If
  '
  CODIEQ$ = TRIM$(Text1.TEXT)
  If CODIEQ$ = "" Then
    Call MENSERR(24, "Código no Válido o Pre-Existente.")
    GoTo 99
  End If
  '
'  If Val(Text7.TEXT) > 32767 Then
'    Call MENSERR(24, "Código de Representante no Válido.")
'    GoTo 99
'  End If
'  If TRIM$(ECOARCH$("ACREDOR", MKI$(Int(Val(Text7.TEXT))))) = "" Then
'    Call MENSERR(24, "Código de Representante no Válido.")
'    GoTo 99
'  End If
'  '
'  If Val(Text10.TEXT) > 32767 Then
'    Call MENSERR(24, "Código de Planta no Válido.")
'    GoTo 99
'  End If
'  If TRIM$(ECOARCH$("PLANTPRO", Text10.TEXT)) = "" Then
'    Call MENSERR(24, "Código de Planta no Válido.")
'    GoTo 99
'  End If
'  '
'  If Val(Text12.TEXT) > 32767 Then
'    Call MENSERR(24, "Código de Sección no Válido.")
'    GoTo 99
'  End If
'  CS1 = Val(Text12.TEXT)
'  If TRIM$(ECOARCH$("SECPROD", MKS$(CS1))) = "" Then
'    Call MENSERR(24, "Código de Sección no Válido.")
'    GoTo 99
'  End If
'  '
  If TRIM$(Text14.TEXT) <> "" Then
    FF1% = FECHANUM(Text14.TEXT)
    FEX = FF1%: FECHAREC$ = FECHATEX$(FEX)
    If TRIM$(FECHAREC$) = "" Then
        Call MENSERR(24, "Fecha de Fin de la Garantía Incorrecta")
        GoTo 99
    End If
  End If
  '
  If TRIM$(Text13.TEXT) <> "" Then
    FF2% = FECHANUM(Text13.TEXT)
    FEX = FF2%: FECHAREC$ = FECHATEX$(FEX)
    If TRIM$(FECHAREC$) = "" Then
        Call MENSERR(24, "Fecha de Instalación Incorrecta")
        GoTo 99
    End If
  End If
  '
  If TRIM$(Text15.TEXT) <> "" Then
    If TRIM$(ECOARCH$("ABMSIS", TRIM$(Text15))) = "" Then
      Call MENSERR(24, "Sistema / Linea no Válido.")
      GoTo 99
    End If
  End If
  '
  Call ABREMAQUI
  '
  Call CREACAMPO("", "EQUIPOS", "Equi_Alter", 10, 16, 1)
  Call CREACAMPO("", "EQUIPOS", "Cuot_Hora", 7, 8)
  '
  'SE COMENTAN ESTAS DOS LINEAS PARA QUE NO BORREN LA INFORMACION INGRESADA
  'PARA LOS ACTIVOS FIJOS
'  Call BORRAREGISTROS("EQUIPOS", "Codigo_Equ = '" & Text1.TEXT & "' ", RAFE&)
'  Call CreaRegistro("EQUIPOS", "Codigo_Equ", Text1.TEXT)
  '
  'VERIFICA SI EXISTE EL CODIGO SI NO LO CREA (POR SI EN OTRA TERMINAL SE ELIMINO)
  CONDI$ = "Codigo_Equ = '" & CODIEQ$ & "' "
  If CantRegistros&("EQUIPOS", CONDI$, "NOMESS") < 1 Then
     Call CreaRegistro("EQUIPOS", "Codigo_Equ", CODIEQ$, "NOMESS")
  End If
  '
  Dim ConsSQL As String
  Dim Equipos As ADODB.Recordset
  
  ConsSQL = "SELECT * FROM EQUIPOS WHERE " & CONDI$
  Set Equipos = New ADODB.Recordset
  Equipos.Open FILTSQL$(ConsSQL), FLEXCONN, adOpenKeyset, adLockPessimistic
  
  With Equipos
    If .BOF = True And .EOF = True Then
      Call MENSERR(24, "Código no Válido o Pre-Existente.")
      GoTo 99
    Else 'GRABACIÓN
   
      !Descri_Equ = TRIM$(Text3)
      !codlar_equ = TRIM$(TEXT21)
      !Marca_Equ = Text23
      !Mode_Equ = Text4(1)
      !NuSeRi_Equ = Text6
      !Fabri_Equ = Text9
      !Represe_Equ = Val(Text7)
      !Ubica_Equ = Val(Text12) 'SECCIÓN PRODUCTIVA
      !Observa_Equ = Text16
      !Finst_Equ = CVDAT(FECHANUM(Text13))
      !Fgara_Equ = CVDAT(FECHANUM(Text14))
      !Alimenta_Equ = Text8
      !Potenci_Equ = Text11
      !Linea_Equ = Text15
      !Posi_Equ = Text17
      !Planta_Equ = Text10
      !Kilos_Equ = Val(Text5.TEXT)
      !Largo_Equ = Val(Text18.TEXT)
      !Ancho_Equ = Val(Text19.TEXT)
      !Alto_Equ = Val(Text20.TEXT)
      !Equi_Alter = Left$(Text24, 6)
      !Cuot_Hora = XVALO(Text22)
     .Update
    End If
    MODIFIC% = 0
  End With
  Equipos.Close
  Set Equipos = Nothing
  '
  Screen.MousePointer = 11
  SQLX$ = "SELECT * FROM EQUIPOS WITH(NOLOCK)"
  Dim EQSEL As ADODB.Recordset
  Set EQSEL = New ADODB.Recordset
  Set EQSEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
  With EQSEL
     EQAGRE% = 0
     Do While Not .EOF
        CODEQ$ = SAFETEXT$(!codigo_equ)
        DESEQ$ = SAFETEXT$(!Descri_Equ)
        SECEQ = XVALO(!Ubica_Equ)
        MAQALT$ = SAFETEXT$(!Equi_Alter)
        CUOTA = XVALO(!Cuot_Hora)
        '
10      XXX$ = TRIM$(FILTERGETRECORD("PADMAQ", 1, CODEQ$))
        If XXX$ = "" Then
           XXX$ = REGBLAN$("PADMAQ")
           Call REPLA(XXX$, CODEQ$, 1, 6)
           Call REGAPP("PADMAQ", XXX$)
           GoTo 10
        End If
        '
        Call REPLA(XXX$, DESEQ$, 7, 30)
        Call REPLA(XXX$, MKS$(SECEQ), 37, 4)
        Call REPLA(XXX$, MAQALT$, 43, 6)
        Call REPLA(XXX$, MKS$(CUOTA), 93, 4)
        Call FILTERPUTRECORD("PADMAQ", 1, CODEQ$, XXX$)
        EQAGRE% = 1
     .MoveNext
     Loop
  End With
  '
  If EQAGRE% > 0 Then
     Call CIERRARCH("PADMAQ")
     Call FILESORT("PADMAQ", "", 1, 1, "ASC")
  End If
  Screen.MousePointer = 1
  '
99 End Sub

Sub GRAIMG() 'GRABA IMAGENES ASOCIADAS AL EQUIPO
  'VALIDA QUE MAQUINA SEA EXISTENTE
  MAQ1$ = Text1.TEXT
  '
  Call ABREMAQUI
  
  Dim Equipos As ADODB.Recordset
  Dim ConsSQL As String
  
  ConsSQL = "SELECT * FROM  Equipos WHERE Codigo_Equ = '" & MAQ1$ & "' "
  Set Equipos = New ADODB.Recordset
  Equipos.Open FILTSQL$(ConsSQL), FLEXCONN, adOpenKeyset, adLockPessimistic
  
  With Equipos
   
    If (.BOF = True And .EOF = True) Then
      Call MENSERR(24, "Código no Válido o Pre-Existente.")
      GoTo 99
    End If
  End With
  
  
       
  'GUARDO IMAGENES
  For JX% = 0 To Image99.Count - 1
    If TRIM(Label99(JX%)) <> "" Then 'NO GRABA IMG EN BLANCO
      ConsSQL = "SELECT * FROM  GALEIMAG WHERE Codigo_Equ = '" & MAQ1$ & "' " & " AND " & " Norit_Imag = " & JX% & " "
      Dim GaleImag As ADODB.Recordset
      Set GaleImag = New ADODB.Recordset
      GaleImag.Open FILTSQL$(ConsSQL), FLEXCONN, adOpenKeyset, adLockPessimistic

      With GaleImag
         If .BOF = True And .EOF = True Then
         'GRABACION
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CodiEmpr = CodiEmp%
         On Error GoTo 0
           !codigo_equ = MAQ1$
           !Norit_Imag = JX%
           !Ruta_Imag = Label99(JX%)
           !Refe_Imag = Text99(JX%)
         .Update
        Else 'EDICIÓN
        
          !codigo_equ = MAQ1$
          !Norit_Imag = JX%
          !Ruta_Imag = Label99(JX%)
          !Refe_Imag = Text99(JX%)
        .Update
        End If
      End With
    End If
  Next JX%
  '
99 End Sub
Sub GRAREP() 'GUARDA REPUESTOS ASOCIADOS A LA MAQUINA
  '
  MAQ1$ = Text1.TEXT
  '
  Call ABREMAQUI
  'BORRO LOS REPUESTOS ANTERIORES
  Dim RepuMaq As ADODB.Recordset
  Dim ConsSQL As String
  Set RepuMaq = New ADODB.Recordset
  
  ConsSQL = "SELECT * FROM RepuMaq WHERE Codigo_Equ = '" & MAQ1$ & "' "
  RepuMaq.Open FILTSQL$(ConsSQL), FLEXCONN, adOpenKeyset, adLockPessimistic
  
  With RepuMaq
      If .BOF = False And .EOF = False Then
         .MoveFirst
      Do While Not (.EOF)
       .Delete
       .MoveNext
      Loop
   End If
  End With
  '
  With RepuMaq
    For JX% = 0 To List1.ListCount - 1
      L1$ = List1.List(JX%)
      CODI0$ = Mid$(L1$, 1, 16)
        CODI1% = CODINT%(CODI0$)
      DESC1$ = Mid$(L1$, 18, 20)
      UM2$ = Mid$(L1$, 39, 2)
      USO0$ = Mid$(L1$, 52, 10)
        USO1 = Val(USO0$)
      OBS1$ = Mid$(L1$, 63, 25)
      '
      .AddNew 'GRABO LOS REPUESTOS  ACTUALES   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !codigo_equ = MAQ1$
        !CodInte_Rep = CODI1%
        !CodExte_Rep = CODEXT$(CODI1%)
        !Descrip_Rep = DESCRIT0$(CODI1%)
        !UM_Rep = UNIMED$(CODI1%)
        !Cant_Rep = USO1
        !Observa_Rep = OBS1$
      .Update
    Next JX%
  End With
  '
End Sub
Sub GRASUB() 'GUARDA SUBCONJUNTOS DE LA MAQUINA
  '
  MAQ1$ = Text1.TEXT
  '
  Call ABREMAQUI
  
  'BORRO LOS SUBCONJUNTOS ANTERIORES
  Dim SubConjuntos As ADODB.Recordset
  Set SubConjuntos = New ADODB.Recordset
   '
  If CantRegistros("SubConjuntos", "Codigo_Equ = '" & MAQ1$ & "' ") > 0 Then Call BORRAREGISTROS("SUBCONJUNTOS", "Codigo_Equ='" & MAQ1$ & "'", REGAF&)
  '
  SQLX$ = "SELECT * FROM SUBCONJUNTOS WHERE Codigo_Equ = '" & MAQ1$ & "' "
  SubConjuntos.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
  
  With SubConjuntos
    For JX% = 0 To List2.ListCount - 1
      L1$ = List2.List(JX%)
      L1$ = List2.List(JX%)
      NRO1% = Val(Mid$(L1$, 1, 3))
      REFE1$ = Mid$(L1$, 5, 22)
      MAR1$ = Mid$(L1$, 27, 24)
      MOD1$ = Mid$(L1$, 51, 20)
      SER1$ = Mid$(L1$, 71, 20)
      CRIT$ = Mid$(L1$, 91, 2)
      '
      .AddNew 'GRABO LOS SUBCONJUNTOS ACTUALES   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !NRO_SUB = NRO1%
        !DESCRI_SUB = REFE1$
        !Marca_Sub = MAR1$
        !Modelo_Sub = MOD1$
        !NSerie_Sub = SER1$
        !codigo_equ = MAQ1$
        !Criticidad = CRIT$
      .Update
    Next JX%
  End With
  '
End Sub
Sub GRACON() 'GUARDA CONTADORES ASOCIADOS A LA MAQUINA
  '
  MAQ1$ = Text1.TEXT
  '
  Call ABREMAQUI
  '
  For JX% = 0 To List3.ListCount - 1
    L1$ = List3.List(JX%)
    NRO1% = Val(Mid$(L1$, 1, 3))
    REFE1$ = Mid$(L1$, 5, 22)
    FEQ1 = Val(Mid$(L1$, 28, 8))
    FUL1% = FECHANUM(Mid$(L1$, 37, 8))
    VAL1 = Val(Mid$(L1$, 46, 10))
    FPL1% = FECHANUM(Mid$(L1$, 57, 8))
    PROME1 = Val(Mid$(L1$, 66, 10))
    OBSE1$ = Mid$(L1$, 77, 64)
    '
    Dim Contadores As ADODB.Recordset
    Dim ConsSQL As String
    Set Contadores = New ADODB.Recordset
    
    ConsSQL = "SELECT * FROM CONTADORES WHERE CODIGO_EQU = '" & MAQ1$ & "'  AND Num_Con = " & NRO1% & " "
    Contadores.Open FILTSQL$(ConsSQL), FLEXCONN, adOpenKeyset, adLockPessimistic
    '
    With Contadores
      If (.BOF = False And .EOF = False) Then
          .MoveFirst
          !Freq_Con = FEQ1
          !Referen_Con = REFE1$
          !Observa_Con = OBSE1$
        .Update
      Else
        .AddNew 'GRABO LOS CONTADORES QUE NO EXISTEN   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
          !Num_Con = NRO1%
          !Referen_Con = REFE1$
          !Freq_Con = FEQ1
          !FUL_Con = CVDAT(FUL1%)
          !Val_Con = VAL1
          !FPL_Con = CVDAT(FPL1%)
          !Prom_Con = PROME1
          !Observa_Con = OBSE1$
          !codigo_equ = MAQ1$
        .Update
      End If
    End With
  Next JX%
  '
End Sub
Private Sub Text1_DblClick()
  '
  Call Command9_Click(1)
  '
End Sub

Private Sub Text1_CLICK()
   '
   If TRIM$(ECOARCH$("INDIEQUI", Text1.TEXT)) <> "" Then
    If MODIFIC% = 1 Then
      If COMALTER%("Se han Modificado Datos del Equipo\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
        Exit Sub
      Else
       Call GRATODO
      End If
    End If
   End If
   '
End Sub

Private Sub Text10_Change()
  '
  Label16 = TRIM$(ECOARCH$("PLANTPRO", Text10.TEXT))
  MODIFIC% = 1
  '
End Sub

Private Sub Text10_DblClick()
  '
  Call Command8_Click
  '
End Sub

Private Sub Text11_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text12_Change()
  '
  CS1 = Val(Text12)
  Label5.Caption = TRIM$(ECOARCH$("SECPROD", MKS$(CS1)))
  MODIFIC% = 1
  '
End Sub
Sub BLANFORMU()
  '
  List1.Clear
  List2.Clear
  List3.Clear
  List4.Clear
  '
  Text11 = ""
  Text13 = ""
  Text3 = ""
  Text14 = ""
  Text16 = ""
  Text23 = ""
  Text4(1) = ""
  Text6 = ""
  Text8 = ""
  Text35 = ""
  Text12 = ""
  Text10 = ""
  Text9 = ""
  Text15 = ""
  Text17 = ""
  Text12 = ""
  Label5 = ""
  Label16 = ""
  Text16 = ""
  Text7 = ""
  Text2.TEXT = ""
  Text20.TEXT = ""
  Text19.TEXT = ""
  Text18.TEXT = ""
  Text5.TEXT = ""
  TEXT21.TEXT = ""
  '
End Sub
Sub DECARIMG() 'DESCARGA LAS IMAGENES Y CONTROLES
  '
  'LIMPIO LOS DOS PICTURES INICIALES (0,1)
  For JX% = 0 To 1
    Set Image99(JX%).Picture = Nothing
    Set Picture99(JX%).Picture = Nothing
    Label99(JX%) = ""
    Text99(JX%) = ""
  Next JX%
  '
  'DESCARGO LOS DEMAS CONTROLES
  For JX% = 2 To Picture99.Count - 1
    Unload Picture99(JX%)
    Unload Label99(JX%)
    Unload Text99(JX%)
    Unload Command99(JX%)
    Unload Image99(JX%)
  Next JX%
  '
  'RESETEO LA BARRA DE DESPLAZ.
  Picture4.Height = Picture3.Height
  VScroll1.Min = 0
  VScroll1.Max = (Picture4.Height - Picture3.Height) / 10 + 1
  VScroll1.Value = 0
  '
End Sub
Private Sub Command6_Click()
 If TRIM$(Text1) <> "" Then
  Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text13.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text13.TEXT)
 End If
End Sub

Private Sub Text12_DblClick()
  '
  Call Command10_Click
  '
End Sub

Private Sub Text13_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text13_DBLCLICK()
  '
  Call Command6_Click
  '
End Sub

Private Sub Text14_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text14_DblClick()
  '
  Call Command7_Click
  '
End Sub

Private Sub Text15_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text15_DblClick()
    Call Command17_Click
End Sub

Private Sub Text16_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text17_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text18_Change()
   '
   MODIFIC% = 1
   '
End Sub
Private Sub Text19_Change()
   '
   MODIFIC% = 1
   '
End Sub

Private Sub Text20_Change()
   '
   MODIFIC% = 1
   '
End Sub

Private Sub Text18_LostFocus()
  Text18 = FORMATNUM(Val(Text18), "F9.3")
End Sub

Private Sub Text19_LostFocus()
  Text19 = FORMATNUM(Val(Text19), "F9.3")
End Sub
Private Sub Text20_LostFocus()
  Text20 = FORMATNUM(Val(Text20), "F9.3")
End Sub

Private Sub Text21_Change()
    '
    MODIFIC% = 1
    '
End Sub

Private Sub Text22_LostFocus()
   Text22 = FORMATNUM$(XVALO(Text22), "F10.2")
End Sub

Private Sub Text23_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text3_Change()
    '
    MODIFIC% = 1
    '
End Sub

Private Sub Text4_Change(Index As Integer)
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text5_Change()
    '
    MODIFIC% = 1
    '
End Sub

Private Sub Text5_LOSTFOCUS()
  Text5 = FORMATNUM(Val(Text5), "F9.3")
End Sub

Private Sub Text6_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text7_Change()
  '
  If Val(Text7) > 32767 Then
    Call COMUNI("Número de Representante no Válido.")
    Text1.SelStart = 0
    Text1.SelLength = (Len(Text7))
    Exit Sub
  End If
  PRO1% = Val(Text7)
  Label13.Caption = TRIM$(ECOARCH$("ACREDOR", MKI$(PRO1%)))
  PRO2% = PRO1% * (-1)
  Label24 = DOMICLI$(PRO2%)
  Label25 = CPLOCLI$(PRO2%)
  Label26 = TEL1CLI$(PRO2%)
  MODIFIC% = 1
  '
End Sub

Private Sub Text7_DblClick()
  '
  Call Command11_Click
  '
End Sub

Private Sub Text8_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text9_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text99_Change(Index As Integer)
  '
  MODIFIC% = 1
  '
End Sub

Private Sub VScroll1_Change()
   On Error Resume Next 'POR ERROR '6' DESBORDAMIENTO
     Picture4.Top = (-10) * VScroll1.Value
   On Error GoTo 0
End Sub

Private Sub VScroll1_Scroll()
  On Error Resume Next 'POR ERROR '6' DESBORDAMIENTO
   Picture4.Top = (-10) * VScroll1.Value
  On Error GoTo 0
End Sub
Private Sub BAJAMAQ(MAQ1$)
  '
  DEMAQ1$ = Text3.TEXT
  If COMALTER%(" Baja Fisica del Equipo '" + TRIM$(MAQ1$) + "' \" + TRIM$(DEMAQ1$) + "\\Cancelar\Confirmar") <> 2 Then GoTo 99
    Call ABREMAQUI
    CONDI$ = " Codigo_Equ = '" & MAQ1$ & "' "
    'BORRO EQUIPO
    Call BORRAREGISTROS("EQUIPOS", CONDI$, REG&)
    Call INDIEQUI
    Call BLANFORMU
    Text1.TEXT = ""
    'BORRO IMAGENES
    Call BORRAREGISTROS("GALEIMAG", CONDI$, REG&)
    'BORRO REPUESTOS
    Call BORRAREGISTROS("REPUMAQ", CONDI$, REG&)
    'BORRO CONTADORES
  '
99 End Sub
Sub LIEQUIAT()

 
  '
99 End Sub

