VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form ENSAMOT2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ensayo de Motores"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "ENSAMOT2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin VB.CommandButton Command4 
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
      Height          =   645
      Left            =   11130
      Picture         =   "ENSAMOT2.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   126
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   840
      Width           =   645
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8310
      Left            =   0
      TabIndex        =   1
      Top             =   105
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   14658
      _Version        =   327681
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   -2147483637
      TabCaption(0)   =   "Archivo Importación"
      TabPicture(0)   =   "ENSAMOT2.frx":18DC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Tabla Valores"
      TabPicture(1)   =   "ENSAMOT2.frx":18F8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(1)=   "Frame1(0)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Ensayo Completo"
      TabPicture(2)   =   "ENSAMOT2.frx":1914
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSChart1(0)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Inversión de Marcha"
      TabPicture(3)   =   "ENSAMOT2.frx":1930
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MSChart2(0)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Cupla - rpm"
      TabPicture(4)   =   "ENSAMOT2.frx":194C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4(2)"
      Tab(4).Control(1)=   "Frame4(3)"
      Tab(4).Control(2)=   "MSChart3"
      Tab(4).ControlCount=   3
      Begin MSChartLib.MSChart MSChart1 
         Height          =   7110
         Index           =   0
         Left            =   -74475
         OleObjectBlob   =   "ENSAMOT2.frx":1968
         TabIndex        =   2
         Top             =   840
         Width           =   9990
      End
      Begin MSChartLib.MSChart MSChart2 
         Height          =   6690
         Index           =   0
         Left            =   -74265
         OleObjectBlob   =   "ENSAMOT2.frx":405E
         TabIndex        =   3
         Top             =   945
         Width           =   9255
      End
      Begin MSChartLib.MSChart MSChart3 
         Height          =   4950
         Left            =   -74080
         OleObjectBlob   =   "ENSAMOT2.frx":6765
         TabIndex        =   4
         Top             =   735
         Width           =   9255
      End
      Begin VB.Frame Frame4 
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
         Height          =   1800
         Index           =   3
         Left            =   -69435
         TabIndex        =   109
         Top             =   6090
         Width           =   4635
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Carr:"
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
            Index           =   37
            Left            =   2415
            TabIndex        =   125
            Top             =   375
            Width           =   1020
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cmax:"
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
            Index           =   36
            Left            =   2415
            TabIndex        =   124
            Top             =   690
            Width           =   1020
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "V.S.(rpm):"
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
            Index           =   35
            Left            =   105
            TabIndex        =   123
            Top             =   670
            Width           =   1020
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Index           =   3
            Left            =   3570
            TabIndex        =   122
            Top             =   630
            Width           =   855
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Index           =   3
            Left            =   3570
            TabIndex        =   121
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Index           =   3
            Left            =   1260
            TabIndex        =   120
            Top             =   630
            Width           =   855
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Index           =   3
            Left            =   1260
            TabIndex        =   119
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "V.N.(rpm):"
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
            Index           =   34
            Left            =   105
            TabIndex        =   118
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label8 
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
            Index           =   3
            Left            =   1260
            TabIndex        =   117
            Top             =   1260
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "s (%):"
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
            Index           =   33
            Left            =   105
            TabIndex        =   116
            Top             =   1305
            Width           =   1020
         End
         Begin VB.Label Label7 
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
            Index           =   3
            Left            =   1260
            TabIndex        =   115
            Top             =   945
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "V(Cmax):"
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
            Index           =   32
            Left            =   105
            TabIndex        =   114
            Top             =   990
            Width           =   1020
         End
         Begin VB.Label Label6 
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
            Index           =   3
            Left            =   3570
            TabIndex        =   113
            Top             =   1260
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cmax / Cn:"
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
            Index           =   31
            Left            =   2415
            TabIndex        =   112
            Top             =   1305
            Width           =   1020
         End
         Begin VB.Label Label5 
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
            Index           =   3
            Left            =   3570
            TabIndex        =   111
            Top             =   945
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Carr / Cn:"
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
            Index           =   30
            Left            =   2415
            TabIndex        =   110
            Top             =   990
            Width           =   1020
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Alta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Index           =   2
         Left            =   -74160
         TabIndex        =   92
         Top             =   6090
         Width           =   4635
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Carr / Cn:"
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
            Index           =   29
            Left            =   2415
            TabIndex        =   108
            Top             =   990
            Width           =   1020
         End
         Begin VB.Label Label5 
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
            Index           =   2
            Left            =   3570
            TabIndex        =   107
            Top             =   945
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cmax / Cn:"
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
            Index           =   28
            Left            =   2415
            TabIndex        =   106
            Top             =   1305
            Width           =   1020
         End
         Begin VB.Label Label6 
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
            Index           =   2
            Left            =   3570
            TabIndex        =   105
            Top             =   1260
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "V(Cmax):"
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
            Index           =   27
            Left            =   105
            TabIndex        =   104
            Top             =   990
            Width           =   1020
         End
         Begin VB.Label Label7 
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
            Index           =   2
            Left            =   1260
            TabIndex        =   103
            Top             =   945
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "s (%):"
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
            Index           =   26
            Left            =   105
            TabIndex        =   102
            Top             =   1305
            Width           =   1020
         End
         Begin VB.Label Label8 
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
            Index           =   2
            Left            =   1260
            TabIndex        =   101
            Top             =   1260
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "V.N.(rpm):"
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
            Index           =   25
            Left            =   105
            TabIndex        =   100
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Index           =   2
            Left            =   1260
            TabIndex        =   99
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Index           =   2
            Left            =   1260
            TabIndex        =   98
            Top             =   630
            Width           =   855
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Index           =   2
            Left            =   3570
            TabIndex        =   97
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Index           =   2
            Left            =   3570
            TabIndex        =   96
            Top             =   630
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "V.S.(rpm):"
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
            Index           =   24
            Left            =   105
            TabIndex        =   95
            Top             =   670
            Width           =   1020
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cmax:"
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
            Index           =   23
            Left            =   2415
            TabIndex        =   94
            Top             =   690
            Width           =   1020
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Carr:"
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
            Index           =   22
            Left            =   2415
            TabIndex        =   93
            Top             =   375
            Width           =   1020
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Amplitud para Linealizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   5
         Left            =   210
         TabIndex        =   71
         Top             =   3570
         Width           =   5160
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
            Index           =   2
            Left            =   4200
            TabIndex        =   127
            Text            =   "50"
            Top             =   525
            Width           =   540
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
            Index           =   1
            Left            =   2415
            TabIndex        =   75
            Text            =   "10"
            Top             =   525
            Width           =   540
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
            Index           =   0
            Left            =   840
            TabIndex        =   72
            Text            =   "15"
            Top             =   525
            Width           =   540
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Intervalo de Clase:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   38
            Left            =   3150
            TabIndex        =   128
            Top             =   420
            Width           =   915
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Baja:"
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
            Index           =   13
            Left            =   1260
            TabIndex        =   74
            Top             =   570
            Width           =   1020
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Alta:"
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
            Index           =   10
            Left            =   -105
            TabIndex        =   73
            Top             =   570
            Width           =   810
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Resultados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3795
         Index           =   4
         Left            =   5565
         TabIndex        =   52
         Top             =   840
         Width           =   5160
         Begin VB.Frame Frame4 
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
            Height          =   3165
            Index           =   1
            Left            =   2625
            TabIndex        =   62
            Top             =   420
            Width           =   2325
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Carr:"
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
               Index           =   21
               Left            =   105
               TabIndex        =   91
               Top             =   1845
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cmax:"
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
               Index           =   20
               Left            =   105
               TabIndex        =   90
               Top             =   2160
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "V.S.(rpm):"
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
               Index           =   17
               Left            =   105
               TabIndex        =   87
               Top             =   670
               Width           =   1020
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Index           =   1
               Left            =   1260
               TabIndex        =   85
               Top             =   2100
               Width           =   855
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Index           =   1
               Left            =   1260
               TabIndex        =   84
               Top             =   1785
               Width           =   855
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Index           =   1
               Left            =   1260
               TabIndex        =   83
               Top             =   630
               Width           =   855
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Index           =   1
               Left            =   1260
               TabIndex        =   82
               Top             =   315
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "V.N.(rpm):"
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
               Index           =   15
               Left            =   105
               TabIndex        =   77
               Top             =   360
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Carr / Cn:"
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
               Index           =   12
               Left            =   105
               TabIndex        =   70
               Top             =   2460
               Width           =   1020
            End
            Begin VB.Label Label5 
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
               Index           =   1
               Left            =   1260
               TabIndex        =   69
               Top             =   2415
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cmax / Cn:"
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
               Index           =   11
               Left            =   105
               TabIndex        =   68
               Top             =   2775
               Width           =   1020
            End
            Begin VB.Label Label6 
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
               Index           =   1
               Left            =   1260
               TabIndex        =   67
               Top             =   2730
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "V(Cmax):"
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
               Index           =   10
               Left            =   105
               TabIndex        =   66
               Top             =   990
               Width           =   1020
            End
            Begin VB.Label Label7 
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
               Index           =   1
               Left            =   1260
               TabIndex        =   65
               Top             =   945
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "s (%):"
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
               Index           =   9
               Left            =   105
               TabIndex        =   64
               Top             =   1305
               Width           =   1020
            End
            Begin VB.Label Label8 
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
               Index           =   1
               Left            =   1260
               TabIndex        =   63
               Top             =   1260
               Width           =   855
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Alta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3165
            Index           =   0
            Left            =   210
            TabIndex        =   53
            Top             =   420
            Width           =   2325
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Carr:"
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
               Index           =   19
               Left            =   105
               TabIndex        =   89
               Top             =   1845
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cmax:"
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
               Index           =   18
               Left            =   105
               TabIndex        =   88
               Top             =   2160
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "V.S.(rpm):"
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
               Index           =   16
               Left            =   105
               TabIndex        =   86
               Top             =   670
               Width           =   1020
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Index           =   0
               Left            =   1260
               TabIndex        =   81
               Top             =   2100
               Width           =   855
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Index           =   0
               Left            =   1260
               TabIndex        =   80
               Top             =   1785
               Width           =   855
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Index           =   0
               Left            =   1260
               TabIndex        =   79
               Top             =   630
               Width           =   855
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Index           =   0
               Left            =   1260
               TabIndex        =   78
               Top             =   315
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "V.N.(rpm):"
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
               Index           =   14
               Left            =   105
               TabIndex        =   76
               Top             =   360
               Width           =   1020
            End
            Begin VB.Label Label8 
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
               Index           =   0
               Left            =   1260
               TabIndex        =   61
               Top             =   1260
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "s (%):"
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
               Index           =   8
               Left            =   105
               TabIndex        =   60
               Top             =   1305
               Width           =   1020
            End
            Begin VB.Label Label7 
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
               Index           =   0
               Left            =   1260
               TabIndex        =   59
               Top             =   945
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "V(Cmax):"
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
               Index           =   7
               Left            =   105
               TabIndex        =   58
               Top             =   990
               Width           =   1020
            End
            Begin VB.Label Label6 
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
               Index           =   0
               Left            =   1260
               TabIndex        =   57
               Top             =   2730
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cmax / Cn:"
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
               Index           =   6
               Left            =   105
               TabIndex        =   56
               Top             =   2775
               Width           =   1020
            End
            Begin VB.Label Label5 
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
               Index           =   0
               Left            =   1260
               TabIndex        =   55
               Top             =   2415
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Carr / Cn:"
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
               Left            =   105
               TabIndex        =   54
               Top             =   2460
               Width           =   1020
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000B&
         Caption         =   "Motor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Index           =   3
         Left            =   210
         TabIndex        =   40
         Top             =   840
         Width           =   5160
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
            Height          =   285
            Left            =   3570
            TabIndex        =   46
            Text            =   " "
            Top             =   525
            Width           =   855
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
            Height          =   285
            Left            =   1260
            TabIndex        =   42
            Text            =   " "
            Top             =   525
            Width           =   855
         End
         Begin VB.TextBox Text3 
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
            Left            =   1260
            TabIndex        =   41
            Text            =   " "
            Top             =   945
            Width           =   855
         End
         Begin VB.Label Label2 
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
            Left            =   3570
            TabIndex        =   51
            Top             =   945
            Width           =   855
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Kgm"
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
            Index           =   8
            Left            =   4515
            TabIndex        =   50
            Top             =   990
            Width           =   810
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.Nominal:"
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
            Index           =   4
            Left            =   2415
            TabIndex        =   49
            Top             =   990
            Width           =   1020
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "V.Nom:"
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
            Index           =   7
            Left            =   2625
            TabIndex        =   48
            Top             =   570
            Width           =   810
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "rpm"
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
            Index           =   6
            Left            =   4515
            TabIndex        =   47
            Top             =   570
            Width           =   810
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Inercia:"
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
            Index           =   3
            Left            =   105
            TabIndex        =   45
            Top             =   990
            Width           =   1020
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
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   315
            TabIndex        =   44
            Top             =   570
            Width           =   810
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "HP"
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
            Left            =   2205
            TabIndex        =   43
            Top             =   570
            Width           =   810
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Encoder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   2
         Left            =   210
         TabIndex        =   34
         Top             =   2415
         Width           =   5160
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
            Height          =   285
            Left            =   3570
            TabIndex        =   39
            Text            =   "11.71875"
            Top             =   525
            Width           =   1170
         End
         Begin VB.TextBox Text1 
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
            Left            =   1260
            TabIndex        =   37
            Text            =   "0.005"
            Top             =   525
            Width           =   855
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Seg"
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
            Left            =   2205
            TabIndex        =   38
            Top             =   570
            Width           =   810
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Intervalo:"
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
            Index           =   2
            Left            =   315
            TabIndex        =   36
            Top             =   570
            Width           =   810
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "K:"
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
            Index           =   2
            Left            =   2415
            TabIndex        =   35
            Top             =   570
            Width           =   1020
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ensayo de Baja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3585
         Index           =   1
         Left            =   -74630
         TabIndex        =   31
         Top             =   4410
         Width           =   10305
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
            Height          =   2760
            Index           =   1
            Left            =   210
            TabIndex        =   32
            Top             =   650
            Width           =   9885
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "   Ord   Puls        Seg        rpm(i)       rpm(p)     Cupla(i)     Cupla(p)"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   210
            TabIndex        =   33
            Top             =   410
            Width           =   9885
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ensayo de Alta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3585
         Index           =   0
         Left            =   -74630
         TabIndex        =   28
         Top             =   630
         Width           =   10305
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
            Height          =   2760
            Index           =   0
            Left            =   210
            TabIndex        =   29
            Top             =   650
            Width           =   9885
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "   Ord   Puls        Seg        rpm(i)       rpm(p)     Cupla(i)     Cupla(p)"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   30
            Top             =   410
            Width           =   9885
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Archivos de Importación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   210
         TabIndex        =   5
         Top             =   4725
         Width           =   10515
         Begin VB.CheckBox Check1 
            Caption         =   "Limite V.Max"
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
            Left            =   9450
            TabIndex        =   129
            Top             =   525
            Width           =   855
         End
         Begin VB.CommandButton Command2 
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
            Height          =   750
            Left            =   9450
            Picture         =   "ENSAMOT2.frx":87B7
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Inicia Importación de Datos"
            Top             =   1155
            Width           =   855
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   9430
            ScaleHeight     =   75
            ScaleWidth      =   825
            TabIndex        =   25
            Top             =   2100
            Width           =   885
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.PictureBox Picture1 
            Height          =   600
            Left            =   9250
            ScaleHeight     =   540
            ScaleWidth      =   1170
            TabIndex        =   24
            Top             =   2360
            Width           =   1230
            Begin VB.Line Line2 
               X1              =   1160
               X2              =   1160
               Y1              =   630
               Y2              =   0
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   1155
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Image Image2 
               Height          =   510
               Left            =   -400
               Picture         =   "ENSAMOT2.frx":8AC1
               ToolTipText     =   "Doble Click para Configurar Modalidad de Cálculo"
               Top             =   0
               Width           =   2010
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Ensayo de Baja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Index           =   1
            Left            =   4830
            TabIndex        =   15
            Top             =   420
            Width           =   4425
            Begin VB.OptionButton Option2 
               Caption         =   "TXT (delimitado por CR-LF)"
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
               Left            =   1050
               TabIndex        =   18
               Top             =   1995
               Width           =   3000
            End
            Begin VB.OptionButton Option1 
               Caption         =   "CSV (delimitado por comas)"
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
               Left            =   1050
               TabIndex        =   17
               Top             =   1680
               Value           =   -1  'True
               Width           =   3060
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
               Height          =   330
               Index           =   1
               Left            =   3990
               TabIndex        =   16
               Top             =   525
               Width           =   175
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Carpeta:"
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
               Left            =   105
               TabIndex        =   23
               Top             =   870
               Width           =   810
            End
            Begin VB.Label Label11 
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
               Height          =   580
               Index           =   1
               Left            =   1050
               TabIndex        =   22
               Top             =   885
               Width           =   3105
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Archivo:"
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
               Left            =   105
               TabIndex        =   21
               Top             =   570
               Width           =   810
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
               Height          =   300
               Index           =   1
               Left            =   1050
               TabIndex        =   20
               Top             =   525
               Width           =   2940
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Formato:"
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
               Left            =   105
               TabIndex        =   19
               Top             =   1680
               Width           =   810
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Ensayo de Alta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Index           =   0
            Left            =   210
            TabIndex        =   6
            Top             =   420
            Width           =   4425
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
               Height          =   330
               Index           =   0
               Left            =   3990
               TabIndex        =   14
               Top             =   525
               Width           =   175
            End
            Begin VB.OptionButton Option1 
               Caption         =   "CSV (delimitado por comas)"
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
               Left            =   1050
               TabIndex        =   12
               Top             =   1680
               Value           =   -1  'True
               Width           =   3060
            End
            Begin VB.OptionButton Option2 
               Caption         =   "TXT (delimitado por CR-LF)"
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
               Left            =   1050
               TabIndex        =   11
               Top             =   1995
               Width           =   3000
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Formato:"
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
               Index           =   0
               Left            =   105
               TabIndex        =   13
               Top             =   1680
               Width           =   810
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
               Height          =   300
               Index           =   0
               Left            =   1050
               TabIndex        =   10
               Top             =   525
               Width           =   2940
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Archivo:"
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
               Index           =   0
               Left            =   105
               TabIndex        =   9
               Top             =   570
               Width           =   810
            End
            Begin VB.Label Label11 
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
               Height          =   580
               Index           =   0
               Left            =   1050
               TabIndex        =   8
               Top             =   885
               Width           =   3105
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Carpeta:"
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
               Index           =   0
               Left            =   105
               TabIndex        =   7
               Top             =   870
               Width           =   810
            End
         End
      End
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
      Height          =   645
      Left            =   11130
      Picture         =   "ENSAMOT2.frx":C0AB
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   105
      Width           =   645
   End
End
Attribute VB_Name = "ENSAMOT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CARE&(1), CAREMA&(1), VEPROMA(1)
Dim VEPUL%(1, 32767)
Dim VETIE(1, 32767)
Dim VERPMI(1, 32767)
Dim VERPMP(1, 32767)
Dim VECUP(1, 32767)
Dim CUPROM(1, 32767)
Dim SIGNO%(1, 32767)
Dim PUINVE&(1)
Dim VCORTE
Dim CUPLAMAX

Dim ARCHIM$(1)

Private Sub Command1_Click(Index As Integer)
   '
   Command1(Index).Enabled = False
   '
   CommonDialog1.CancelError = True
   Label11(Index) = "": Label1(Index) = ""
   '
10 CommonDialog1.DialogTitle = "Lectura Datos de Ensayo - " + TRIM$(Mid$(Frame3(Index).Caption, 11))
   CommonDialog1.FILTER = "CSV - Delimitado por Comas|*.csv|"
   If Option2(Index).Value = True Then
     CommonDialog1.FILTER = "TXT - Texto|*.txt|"
   End If
   CommonDialog1.Flags = &H1000 + &H200000 + &H800& + &H4000
   CommonDialog1.FilterIndex = 1
   On Error Resume Next
   CommonDialog1.ShowOpen
   If Err < 1 Then
      ARXXX$ = CommonDialog1.filename
      If ARXXX$ <> "" Then
        For PPX% = Len(ARXXX$) To 1 Step -1
          If Mid$(ARXXX$, PPX%, 1) = Chr$(92) Then
            Label11(Index) = Left$(ARXXX$, PPX% - 1)
            Label1(Index) = Mid$(ARXXX$, PPX% + 1)
            If Label11(Index) = "" Or Label1(Index) = "" Then
              Call MENSERR(24, "Selección no Válida")
              GoTo 10
            End If
            GoTo 99
          End If
        Next PPX%
        Call MENSERR(24, "Selección no Válida")
        GoTo 10
      End If
   End If
   On Error GoTo 0
   '
99 Command1(Index).Enabled = True
   '
End Sub

'
Private Sub Command2_Click()
   '
   Command2.Enabled = False
   '
'Call LEARCHI
'GoTo 80
   HAYDATOS% = 0
   For IENSA% = 0 To 1
     ARCHIM$(IENSA%) = TRIM$(Label11(IENSA%)) + "\" + TRIM$(Label1(IENSA%))
     If EXISTE%(ARCHIM$(IENSA%)) > 0 Then
       HAYDATOS% = 1
     End If
   Next IENSA%
   If HAYDATOS% < 1 Then
     Call MENSERR(24, "Falta Seleccionar Archivo(s)")
     GoTo 99
   End If
   '
   INCREM = CVAL(Text1)
   If INCREM < 0.001 Then
     Call MENSERR(24, "Datos Encoder no Válidos")
     GoTo 99
   End If
   '
   CONSTA = CVAL(Text2)
   If CONSTA < 1 Then
     Call MENSERR(24, "Datos Encoder no Válidos")
     GoTo 99
   End If
   '
   JOTA = CVAL(Text3)
   If JOTA < 0.001 Then
     Call MENSERR(24, "Datos Inercia no Válidos")
     GoTo 99
   End If
   '
   POTEN = CVAL(Text4)
   If POTEN < 0.1 Then
     Call MENSERR(24, "Falta Potencia Nominal")
     GoTo 99
   End If
   '
   VELO = CVAL(Text6)
   If VELO < 300 Then
     Call MENSERR(24, "Falta Velocidad Nominal")
     GoTo 99
   End If
   PULNOMIN% = 1 + Int(VELO / CONSTA)
   '
   For IENSA% = 0 To 1
     If CVAL(Text7(IENSA%)) < 2 Then
       Call MENSERR(24, "Falta Amplitud para Linealizar")
       GoTo 99
     End If
   Next IENSA%
   '
   If CVAL(Text7(2)) < 50 Or CVAL(Text7(2)) > 200 Then
     Call MENSERR(24, "Intervalo de Clase no Válido")
     Text7(2) = "50"
     GoTo 99
   End If
   '
   For IENSA% = 0 To 1
     Label9(IENSA%) = ""
     Label13(IENSA%) = ""
     Label7(IENSA%) = ""
     Label8(IENSA%) = ""
     '
     Label9(IENSA% + 2) = ""
     Label13(IENSA% + 2) = ""
     Label7(IENSA% + 2) = ""
     Label8(IENSA% + 2) = ""
     '
     Label14(IENSA%) = ""
     Label15(IENSA%) = ""
     Label5(IENSA%) = ""
     Label6(IENSA%) = ""
     '
     Label14(IENSA% + 2) = ""
     Label15(IENSA% + 2) = ""
     Label5(IENSA% + 2) = ""
     Label6(IENSA% + 2) = ""
     '
     CARE&(IENSA%) = 0
     CAREMA&(IENSA%) = 0
     VEPROMA(IENSA%) = 0
     PUINVE&(IENSA%) = 0
     VCORTE = 0
     CUPLAMAX = 0
     '
     For KKI% = 0 To 32766
       VEPUL%(IENSA%, KKI%) = 0
       VETIE(IENSA%, KKI%) = 0
       VERPMI(IENSA%, KKI%) = 0
       VERPMP(IENSA%, KKI%) = 0
       VECUP(IENSA%, KKI%) = 0
       CUPROM(IENSA%, KKI%) = 0
       SIGNO%(IENSA%, KKI%) = 0
     Next KKI%
   Next IENSA%
   '
   CUPLAMAX = 0
   Screen.MousePointer = 11
   For IENSA% = 0 To 1
     If EXISTE%(ARCHIM$(IENSA%)) > 0 Then
       '
       VELOMIN = 999999: ARRANCO% = 0
       PULMAX% = 1
       TINI = 0
       VETIE(IENSA%, 0) = TINI
       '
       JJ& = 0
       N1% = FILEOPEN%(ARCHIM$(IENSA%), 1, 128)
       FIN& = LOF(N1%): CALE& = 0
       While Not EOF(N1%)
         Line Input #N1%, AAA$
         CALE& = CALE& + Len(AAA$) + 2
         Call TRACE(20, CALE&, FIN&)
         If JJ& < 32767 Then
           If CVAL(AAA$) > 0 Or JJ& > 0 Then
             JJ& = JJ& + 1
             VEPUL%(IENSA%, JJ&) = CVAL(AAA$)
             If Check1.Value = 1 Then
               If VEPUL%(IENSA%, JJ&) > PULNOMIN% Then
                 VEPUL%(IENSA%, JJ&) = PULNOMIN%
               End If
             End If
             If VEPUL%(IENSA%, JJ&) > PULMAX% Then PULMAX% = VEPUL%(IENSA%, JJ&)
           End If
         End If
       Wend
20     CARE&(IENSA%) = JJ&
       '
       While (CARE&(IENSA%) > 1 And VEPUL%(IENSA%, CARE&(IENSA%)) < 1)
         CARE&(IENSA%) = CARE&(IENSA%) - 1
       Wend
       '
       YAUMENTO% = 0
       UVELOMI& = 0
       For U& = 1 To CARE&(IENSA%)
         VETIE(IENSA%, U&) = VETIE(IENSA%, U& - 1) + INCREM
         VERPMI(IENSA%, U&) = VEPUL%(IENSA%, U&) * CONSTA
         If Abs(VERPMI(IENSA%, U&)) > VELO / 10 Then ARRANCO% = 1
         If ARRANCO% > 0 Then
           If YAUMENTO% = 0 Then
             If Abs(VERPMI(IENSA%, U&)) < VELOMIN Then
               VELOMIN = Abs(VERPMI(IENSA%, U&))
               UVELOMI& = U&
             End If
             If Abs(VERPMI(IENSA%, U&)) > 5 * (VELOMIN + CONSTA) Then
               YAUMENTO% = 1
             End If
           End If
         End If
       Next U&
       '
       If UVELOMI& < 50 Or YAUMENTO% < 1 Then
         Call MENSERR(24, "Ensayo no Válido\No Detectada Inversión de Marcha")
       End If
       '
       ' determinar sentido de giro
       UINVE& = 0
       For U& = 1 To CARE&(IENSA%)
         Call TRACE(20, U&, CARE&(IENSA%))
         'If U& = 1 Then
         '     SIGNO%(IENSA%, U&) = (-1)
         '   Else
         '     SIGNO%(IENSA%, U&) = SIGNO%(IENSA%, U& - 1)
              'If U& > UINVE& + 20 Then
              '  If Abs(VERPMI(IENSA%, U&)) < 2 * VELOMIN + 1.1 * CONSTA Then
              '    UINVE& = U&
              '    SIGNO%(IENSA%, U&) = (-1) * SIGNO%(IENSA%, U& - 1)
              '  End If
              'End If
         'End If
         SIGNO%(IENSA%, U&) = (-1)
         If U& > UVELOMI& Then
           SIGNO%(IENSA%, U&) = 1
         End If
29     Next U&
       '
       'pone signo a la velocidad instantanea
       For U& = 1 To CARE&(IENSA%)
         VERPMI(IENSA%, U&) = SIGNO%(IENSA%, U&) * Abs(VERPMI(IENSA%, U&))
       Next U&
       '
       'calcula velocidad promedio
       PUINVE&(IENSA%) = 0
       For U& = 1 To CARE&(IENSA%)
         Call TRACE(20, U&, CARE&(IENSA%))
         KKI% = 0
         SURPM = VERPMI(IENSA%, U&)
         SUMAPUL% = VEPUL%(IENSA%, U&)
         While SUMAPUL% < 3000
           KKI% = KKI% + 1
           LINF& = U& - KKI%: If LINF < 1 Then LINF = 1
           LSUP& = U& + KKI%: If LSUP& > CARE&(IENSA%) Then LSUP& = CARE&(IENSA%)
           SURPM = SURPM + VERPMI(IENSA%, LINF&) + VERPMI(IENSA%, LSUP&)
           SUMAPUL% = SUMAPUL% + VEPUL%(IENSA%, LINF&) + VEPUL%(IENSA%, LSUP&)
         Wend
         VERPMP(IENSA%, U&) = SURPM / (2 * KKI% + 1)
         If U& > 1 Then
           If Sgn(VERPMP(IENSA%, U&)) <> Sgn(VERPMP(IENSA%, U& - 1)) Then
             If PUINVE&(IENSA%) < 10 Then
               PUINVE&(IENSA%) = U&
             End If
           End If
         End If
       Next U&
       '
       CAREMA&(IENSA%) = 1: VEPROMA(IENSA%) = 0
       For U& = 1 To CARE&(IENSA%)
         If VERPMP(IENSA%, U&) > VEPROMA(IENSA%) * 1.0001 Then
           VEPROMA(IENSA%) = VERPMP(IENSA%, U&)
           CAREMA&(IENSA%) = U&
         End If
       Next U&
       '
       'calcula cupla instantanea
       For U& = 1 To CARE&(IENSA%)
         Call TRACE(20, U&, CARE&(IENSA%))
         LINF& = U& - 2: If LINF < 1 Then LINF = 1
         LSUP& = U& + 2: If LSUP& > CARE&(IENSA%) Then LSUP& = CARE&(IENSA%)
         VECUP(IENSA%, U&) = JOTA * (VERPMP(IENSA%, LSUP&) - VERPMP(IENSA%, LINF&)) / (VETIE(IENSA%, LSUP&) - VETIE(IENSA%, LINF))
       Next U&
       '
       AMPLI% = CVAL(Text7(IENSA%))
       VMX = VEPROMA(IENSA%)
       '
       For U& = 1 To CARE&(IENSA%)
         Call TRACE(20, U&, CARE&(IENSA%))
         AMPLIX% = AMPLI%
         '
         If Abs(VERPMP(IENSA%, U&)) < VMX / 2 Then
           AMPLIX% = 45 - (Abs(VERPMP(IENSA%, U&)) * (45 - AMPLI%)) / (VMX / 2)
         End If
         '
         LINF& = U& - AMPLIX%: If LINF < 1 Then LINF = 1
         LSUP& = U& + AMPLIX%: If LSUP& > CARE&(IENSA%) Then LSUP& = CARE&(IENSA%)
         SUCUP = 0
         ' ESTA ES LA FORMA DE CALCULO SEGUN VENUTOLO
         For JJ& = LINF& To LSUP&
           SUCUP = SUCUP + VECUP(IENSA%, JJ&)
         Next JJ&
         CUPROM(IENSA%, U&) = SUCUP / (LSUP& - LINF& + 1)
         '
         'ESTA ES POR CUADRADOS MINIMOS
         'SUMAVE& = 0: SUMATI = 0
         'For JJ& = LINF& To LSUP&
         '  SUMAVE = SUMAVE + Abs(VERPMP(JJ&) - VERPMP(U&))
         '  SUMATI = SUMATI + Abs(VETIE(JJ&) - VETIE(U&))
         'Next JJ&
         'CUPROM(U&) = SUMAVE / SUMATI
         If Abs(CUPROM(IENSA%, U&)) > CUPLAMAX Then CUPLAMAX = Abs(CUPROM(IENSA%, U&))
       Next U&
     '
     End If
   Next IENSA%


   ' aparear inversión de giro
   If PUINVE&(0) * PUINVE&(1) <> 0 Then
     DIFEPIN& = PUINVE&(1) - PUINVE&(0)
     If DIFEPIN& > 0 Then
         IENSAMO% = 0
       Else
         DIFEPIN& = Abs(DIFEPIN&)
         IENSAMO% = 1
     End If
     For U& = CARE&(IENSAMO%) To 1 Step -1
       If U& + DIFEPIN& < 32767 Then
         SIGNO%(IENSAMO%, U& + DIFEPIN&) = SIGNO%(IENSAMO%, U&)
         VEPUL%(IENSAMO%, U& + DIFEPIN&) = VEPUL%(IENSAMO%, U&)
         VETIE(IENSAMO%, U& + DIFEPIN&) = VETIE(IENSAMO%, U&)
         VERPMI(IENSAMO%, U& + DIFEPIN&) = VERPMI(IENSAMO%, U&)
         VERPMP(IENSAMO%, U& + DIFEPIN&) = VERPMP(IENSAMO%, U&)
         VECUP(IENSAMO%, U& + DIFEPIN&) = VECUP(IENSAMO%, U&)
         CUPROM(IENSAMO%, U& + DIFEPIN&) = CUPROM(IENSAMO%, U&)
       End If
     Next U&
     For U& = 1 To DIFEPIN&
       SIGNO%(IENSAMO%, U&) = SIGNO%(IENSAMO%, DIFEPIN& + 1)
       VEPUL%(IENSAMO%, U&) = VEPUL%(IENSAMO%, DIFEPIN& + 1)
       VETIE(IENSAMO%, U&) = VETIE(IENSAMO%, DIFEPIN& + 1)
       VERPMI(IENSAMO%, U&) = VERPMI(IENSAMO%, DIFEPIN& + 1)
       VERPMP(IENSAMO%, U&) = VERPMP(IENSAMO%, DIFEPIN& + 1)
       VECUP(IENSAMO%, U&) = VECUP(IENSAMO%, DIFEPIN& + 1)
       CUPROM(IENSAMO%, U&) = CUPROM(IENSAMO%, DIFEPIN& + 1)
     Next U&
     CARE&(IENSAMO%) = CARE&(IENSAMO%) + DIFEPIN&: If CARE&(IENSAMO%) > 32766 Then CARE&(IENSAMO%) = 32766
     CAREMA&(IENSAMO%) = CAREMA&(IENSAMO%) + DIFEPIN&
     PUINVE&(IENSAMO%) = PUINVE&(IENSAMO%) + DIFEPIN&
   End If
   '
   '
   For IENSA% = 0 To 1
      List1(IENSA%).Clear
      For U& = 1 To CARE&(IENSA%)
        Call TRACE(20, U&, CARE&(IENSA%))
        TTXX$ = AJUSTD$(Format(U&, "#####0"), 6)
        TTXX$ = TTXX$ + " " + AJUSTD$(Format(VEPUL%(IENSA%, U&), "#####0"), 6)
        TTXX$ = TTXX$ + "  " + AJUSTD$(Format(VETIE(IENSA%, U&), "#####0.000"), 10)
        TTXX$ = TTXX$ + " " + AJUSTD$(Format(VERPMI(IENSA%, U&), "########0.00"), 12)
        TTXX$ = TTXX$ + " " + AJUSTD$(Format(VERPMP(IENSA%, U&), "########0.00"), 12)
        TTXX$ = TTXX$ + " " + AJUSTD$(Format(VECUP(IENSA%, U&), "#######0.000"), 12)
        TTXX$ = TTXX$ + " " + AJUSTD$(Format(CUPROM(IENSA%, U&), "#######0.000"), 12)
        List1(IENSA%).AddItem (TTXX$)
      Next U&
      List1(IENSA%).Refresh
   Next IENSA%
   '
'Call GRARCHI

80 Call GRAFUNCI
   Screen.MousePointer = 1
   MsgBox "Importación Completa"
   '
99 Command2.Enabled = True
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   If Printer.ScaleWidth < 12000 Then
     WIDAN = Width
     Width = 11000
     PRINTFORM
     Width = WIDAN
   End If
   Command4.Enabled = True
End Sub

Private Sub Form_Load()
   Call AUTOEXEC
   abc$ = UCase$(LUDAT$)
   If InStr(abc$, "\DEVSTUDIO\VB\") > 0 Then ' \ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB") > 0 Then
     LUDAT$ = "F:\CLIENTES\ADELSUR\"
     MsgBox LUDAT$
   End If
   '
   HOII% = HOY(HOS$)
   '
   If CVAL(Text6) > 100 Then
     Label2 = Format(CVAL(Text4) * 716 / CVAL(Text6), "##0.00")
   End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Sub GRAFUNCI()
   '
   CAREX& = CARE&(0): If CARE&(1) > CAREX& Then CAREX& = CARE&(1)
   CAREMAX& = CAREMA&(0): If CAREMA&(1) > CAREMAX& Then CAREMAX& = CAREMA&(1)
   '
   MESCALA = 1
   If CUPLAMAX > 1 Then MESCALA = 2
   If CUPLAMAX > 2 Then MESCALA = 5
   If CUPLAMAX > 5 Then MESCALA = 8
   If CUPLAMAX > 8 Then MESCALA = 10
   If CUPLAMAX > 10 Then MESCALA = 20
   If CUPLAMAX > 20 Then MESCALA = 50
   If CUPLAMAX > 50 Then MESCALA = 100
   '
   MSChart1(0).Plot.Axis(2).ValueScale.Maximum = MESCALA
   MSChart1(0).Plot.Axis(2).ValueScale.Minimum = -1 * MESCALA
   MSChart2(0).Plot.Axis(2).ValueScale.Maximum = MESCALA
   MSChart2(0).Plot.Axis(2).ValueScale.Minimum = -1 * MESCALA
   '
   MSChart1(0).Repaint = True
   MSChart2(0).Repaint = True
   '
   For IENSA% = 0 To 1
      '
      INCREMI1% = Int(CAREX) / 500
      MSChart1(0).RowCount = 1 + CAREX& / INCREMI1%
      For UJU& = 1 To MSChart1(0).RowCount
        MSChart1(0).Row = UJU&
        MSChart1(0).Column = 1 + 2 * IENSA%
        MSChart1(0).Data = VERPMP(IENSA%, UJU& * INCREMI1%)
        MSChart1(0).Column = 2 + 2 * IENSA%
        MSChart1(0).Data = CUPROM(IENSA%, UJU& * INCREMI1%)
        MSChart1(0).Column = 5
        MSChart1(0).Data = 0
        If UJU& * INCREMI1% >= PUINVE&(0) Then
          If (UJU& - 1) * INCREMI1% < PUINVE&(0) Then
            MSChart1(0).Data = MESCALA
            MSChart1(0).Row = UJU& - 1
            MSChart1(0).Data = -1 * MESCALA
          End If
        End If
        If IENSA% = 0 Then
          VETIX = CVAL(Format$(VETIE(IENSA%, UJU& * INCREMI1%) / 10, "###0.00"))
          MSChart1(0).RowLabel = ""
          If Int(VETIX) = VETIX Then
            MSChart1(0).RowLabel = TRIM$(Format$(VETIE(IENSA%, UJU& * INCREMI1%), "###0"))
          End If
        End If
      Next UJU&
      '
      INCREMI2% = Int(CAREMAX / 500)
      If INCREMI2% < 1 Then INCREMI2% = 1
      MSChart2(0).RowCount = 1 + CAREMAX& / INCREMI2%
      For UJU& = 1 To MSChart2(0).RowCount
        MSChart2(0).Row = UJU&
        MSChart2(0).Column = 1 + 2 * IENSA%
        MSChart2(0).Data = VERPMP(IENSA%, UJU& * INCREMI2%)
        MSChart2(0).Column = 2 + 2 * IENSA%
        MSChart2(0).Data = CUPROM(IENSA%, UJU& * INCREMI2%)
        MSChart2(0).Column = 5
        MSChart2(0).Data = 0
        If UJU& * INCREMI2% >= PUINVE&(0) Then
          If (UJU& - 1) * INCREMI2% < PUINVE&(0) Then
            MSChart2(0).Data = MESCALA
            MSChart2(0).Row = UJU& - 1
            MSChart2(0).Data = -1 * MESCALA
          End If
        End If
        If IENSA% = 0 Then
          MSChart2(0).RowLabel = TRIM$(Format$(VETIE(IENSA%, UJU& * INCREMI2%), "###0"))
        End If
      Next UJU&
     '
19 Next IENSA%
   '
   Dim SUCUP(1, 40), CACAS%(1, 40), CUPLA(1, 40)
   Dim VSINC(1), VNOMI(1), VCUMA(1), CUPLAMA(1), CUPLARR(1)
   Dim UCLASI%(1)
   CUNOMI = CVAL(Label2)
   '
   INTERCLA = CVAL(Text7(2))
   If INTERCLA < 50 Or INTERCLA > 200 Then INTERCLA = 50
   CACLASI% = 1 + 2000 / INTERCLA
   If CACLASI% > 40 Then CACLASI% = 40
   '
   For IENSA% = 0 To 1
     VSINC(IENSA%) = 0
     VNOMI(IENSA%) = 0
     VCUMA(IENSA%) = 0
     CUPLAMA(IENSA%) = 0
     CUPLARR(IENSA%) = 0
     UCLASI%(IENSA%) = 0
     For CLASI% = 0 To 40
       SUCUP(IENSA%, CLASI%) = 0
       CACAS%(IENSA%, CLASI%) = 0
       CUPLA(IENSA%, CLASI%) = 0
     Next CLASI%
   Next IENSA%
   '
   For IENSA% = 0 To 1
     For UJU& = 1 To CAREX&
       If UJU& >= PUINVE&(IENSA%) Then ' Or Abs(VERPMP(UJU&)) < VEPROMA / 3 Then
         If VERPMP(IENSA%, UJU&) > VSINC(IENSA%) Then VSINC(IENSA%) = VERPMP(IENSA%, UJU&)
         If CUPROM(IENSA%, UJU&) <= 0 Then Exit For
         CLASI% = 1 + Int(Abs(VERPMP(IENSA%, UJU&)) / INTERCLA)
         If CLASI% > 0 Then
           If CLASI% <= CACLASI% Then
             SUCUP(IENSA%, CLASI%) = SUCUP(IENSA%, CLASI%) + Abs(CUPROM(IENSA%, UJU&))
             CACAS%(IENSA%, CLASI%) = CACAS%(IENSA%, CLASI%) + 1
             If Abs(CUPROM(IENSA%, UJU&)) > 0 Then
               If CLASI% > UCLASI%(IENSA%) Then
                 UCLASI%(IENSA%) = CLASI%
               End If
             End If
           End If
         End If
       End If
     Next UJU&
     '
     For CLASI% = UCLASI%(IENSA%) To CACLASI%
       SUCUP(IENSA%, CLASI%) = 0
       CACAS%(IENSA%, CLASI%) = 1
     Next CLASI%
   Next IENSA%
   '
   For IENSA% = 0 To 1
     VANT = 0: CUPLAMA(IENSA%) = 0: VCUMA(IENSA%) = 0
     For CLASI% = 1 To CACLASI%
       If CACAS%(IENSA%, CLASI%) < 1 Then
           CUPLA(IENSA%, CLASI%) = VANT
         Else
           CUPLA(IENSA%, CLASI%) = SUCUP(IENSA%, CLASI%) / CACAS%(IENSA%, CLASI%)
       End If
       If CUPLA(IENSA%, CLASI%) > CUPLAMA(IENSA%) Then
         CUPLAMA(IENSA%) = CUPLA(IENSA%, CLASI%)
         VCUMA(IENSA%) = INTERCLA * (CLASI% + 1)
       End If
       If CLASI% = 1 Then CUPLARR(IENSA%) = CUPLA(IENSA%, CLASI%)
       '
       If CLASI% > 1 Then
         If CUPLA(IENSA%, CLASI% - 1) > CUNOMI Then
           If CUPLA(IENSA%, CLASI%) <= CUNOMI Then
             VNOMI(IENSA%) = INTERCLA * (CLASI% - 1 + (CUPLA(IENSA%, CLASI% - 1) - CUNOMI) / (CUPLA(IENSA%, CLASI% - 1) - CUPLA(IENSA%, CLASI%)))
           End If
         End If
       End If
       '
       On Error Resume Next
       Label9(IENSA%) = TRIM$(Format$(VNOMI(IENSA%), "#####0"))
       Label13(IENSA%) = TRIM$(Format$(VSINC(IENSA%), "#####0"))
       Label7(IENSA%) = TRIM$(Format$(VCUMA(IENSA%), "#####0"))
       Label8(IENSA%) = TRIM$(Format$(100 * (VSINC(IENSA%) - VNOMI(IENSA%)) / VSINC(IENSA%), "#####0.0"))
       '
       Label9(IENSA% + 2) = TRIM$(Format$(VNOMI(IENSA%), "#####0"))
       Label13(IENSA% + 2) = TRIM$(Format$(VSINC(IENSA%), "#####0"))
       Label7(IENSA% + 2) = TRIM$(Format$(VCUMA(IENSA%), "#####0"))
       Label8(IENSA% + 2) = TRIM$(Format$(100 * (VSINC(IENSA%) - VNOMI(IENSA%)) / VSINC(IENSA%), "#####0.0"))
       '
       Label14(IENSA%) = TRIM$(Format$(CUPLARR(IENSA%), "####0.00"))
       Label15(IENSA%) = TRIM$(Format$(CUPLAMA(IENSA%), "####0.00"))
       Label5(IENSA%) = TRIM$(Format$(CUPLARR(IENSA%) / CUNOMI, "####0.00"))
       Label6(IENSA%) = TRIM$(Format$(CUPLAMA(IENSA%) / CUNOMI, "####0.00"))
       '
       Label14(IENSA% + 2) = TRIM$(Format$(CUPLARR(IENSA%), "####0.00"))
       Label15(IENSA% + 2) = TRIM$(Format$(CUPLAMA(IENSA%), "####0.00"))
       Label5(IENSA% + 2) = TRIM$(Format$(CUPLARR(IENSA%) / CUNOMI, "####0.00"))
       Label6(IENSA% + 2) = TRIM$(Format$(CUPLAMA(IENSA%) / CUNOMI, "####0.00"))
       On Error GoTo 0
       '
     Next CLASI%
     '
   Next IENSA%
   '
   MSChart3.RowCount = CACLASI%
   MSChart3.ColumnCount = 3
   For IENSA% = 0 To 1
     For CLASI% = 1 To CACLASI%
        MSChart3.Column = 1 + IENSA%
        MSChart3.Row = CLASI%
        MSChart3.Data = CUPLA(IENSA%, CLASI%)
      Next CLASI%
   Next IENSA%
   '
   For CLASI% = 1 To CACLASI%
     MSChart3.Column = 3
     MSChart3.Row = CLASI%
     MSChart3.Data = CVAL(Label2)
     MSChart3.RowLabel = TRIM$(Str$(Int(INTERCLA * CLASI%)))
   Next CLASI%
   '
End Sub

Private Sub Text3_Change()
   If CVAL(Text6) > 100 Then Label2 = Format(CVAL(Text4) * 716 / CVAL(Text6), "##0.00")
End Sub

Private Sub Text4_Change()
   If CVAL(Text6) > 100 Then Label2 = Format(CVAL(Text4) * 716 / CVAL(Text6), "##0.00")
End Sub

Private Sub Text6_Change()
   If CVAL(Text6) > 100 Then Label2 = Format(CVAL(Text4) * 716 / CVAL(Text6), "##0.00")
End Sub

Sub GRARCHI()
NA% = FILEOPEN%("ALTA", 2, 128)
For U& = 1 To List1(0).ListCount
Print #NA%, List1(0).List(U& - 1)
Next U&
Close #NA%
NB% = FILEOPEN%("BAJA", 2, 128)
For U& = 1 To List1(1).ListCount
Print #NA%, List1(1).List(U& - 1)
Next U&
Close #NB%
End Sub


Sub LEARCHI()
CUPLAMAX = 0: PUINVE&(0) = 0: PUINVE&(1) = 0
NA% = FILEOPEN%(LUDAT$ + "ALTA", 1, 128)
CARE&(0) = 0
CAREMA&(0) = 1: VEPROMA(0) = 0
While Not EOF(NA%)
   Line Input #NA%, TTXX$
   CARE&(0) = CARE&(0) + 1
   VEPUL%(0, CARE&(0)) = CVAL(Mid$(TTXX$, 8, 6))
   VETIE(0, CARE&(0)) = CVAL(Mid$(TTXX$, 16, 10))
   VERPMI(0, CARE&(0)) = CVAL(Mid$(TTXX$, 27, 12))
   VERPMP(0, CARE&(0)) = CVAL(Mid$(TTXX$, 40, 12))
   VECUP(0, CARE&(0)) = CVAL(Mid$(TTXX$, 53, 12))
   CUPROM(0, CARE&(0)) = CVAL(Mid$(TTXX$, 66, 12))
   If Abs(CUPROM(0, CARE&(0))) > CUPLAMAX Then CUPLAMAX = Abs(CUPROM(0, CARE&(0)))
   If VERPMP(0, CARE&(0)) > VEPROMA(0) * 1.0001 Then
     VEPROMA(0) = VERPMP(0, CARE&(0))
     CAREMA&(0) = CARE&(0)
   End If
   If CARE&(0) > 10 Then
     If PUINVE&(0) = 0 Then
       If VERPMP(0, CARE&(0)) > 0 Then
         PUINVE&(0) = CARE&(0)
       End If
     End If
   End If
Wend
Close #NA%
NB% = FILEOPEN%(LUDAT$ + "BAJA", 1, 128)
CARE&(1) = 0
CAREMA&(1) = 1: VEPROMA(1) = 0
While Not EOF(NB%)
   Line Input #NB%, TTXX$
   CARE&(1) = CARE&(1) + 1
   VEPUL%(1, CARE&(1)) = CVAL(Mid$(TTXX$, 8, 6))
   VETIE(1, CARE&(1)) = CVAL(Mid$(TTXX$, 16, 10))
   VERPMI(1, CARE&(1)) = CVAL(Mid$(TTXX$, 27, 12))
   VERPMP(1, CARE&(1)) = CVAL(Mid$(TTXX$, 40, 12))
   VECUP(1, CARE&(1)) = CVAL(Mid$(TTXX$, 53, 12))
   CUPROM(1, CARE&(1)) = CVAL(Mid$(TTXX$, 66, 12))
   If Abs(CUPROM(1, CARE&(1))) > CUPLAMAX Then CUPLAMAX = Abs(CUPROM(1, CARE&(1)))
   If VERPMP(1, CARE&(0)) > VEPROMA(1) * 1.0001 Then
     VEPROMA(1) = VERPMP(1, CARE&(0))
     CAREMA&(1) = CARE&(1)
   End If
   If CARE&(1) > 10 Then
     If PUINVE&(1) = 0 Then
       If VERPMP(1, CARE&(1)) > 0 Then
         PUINVE&(1) = CARE&(1)
       End If
     End If
   End If
Wend
Close #NB%
End Sub


Function CVAL(NTXCV$)
  CVAL = Val(REPLACAR$(NTXCV$, ",", "."))
End Function
