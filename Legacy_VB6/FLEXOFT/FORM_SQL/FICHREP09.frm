VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.Form FICHREP09 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Reparacion"
   ClientHeight    =   8310
   ClientLeft      =   1800
   ClientTop       =   1770
   ClientWidth     =   11910
   Icon            =   "FICHREP09.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command19 
      Caption         =   "Generar Reserva"
      Height          =   495
      Left            =   5880
      TabIndex        =   110
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Command18"
      Height          =   495
      Left            =   11040
      TabIndex        =   109
      Top             =   3600
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   1
      Top             =   -15
      Width           =   12000
      Begin TabDlg.SSTab SSTAB 
         Height          =   6915
         Left            =   75
         TabIndex        =   24
         Top             =   1335
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   12197
         _Version        =   393216
         Tabs            =   7
         Tab             =   5
         TabsPerRow      =   7
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Anotador"
         TabPicture(0)   =   "FICHREP09.frx":030A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame6"
         Tab(0).Control(1)=   "Frame2"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Galeria Imagenes"
         TabPicture(1)   =   "FICHREP09.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture4"
         Tab(1).Control(1)=   "Frame5(1)"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Doc. Anexos"
         TabPicture(2)   =   "FICHREP09.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame9"
         Tab(2).Control(1)=   "Frame8"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Inspeccion"
         TabPicture(3)   =   "FICHREP09.frx":035E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Text3"
         Tab(3).Control(1)=   "Frame11"
         Tab(3).Control(2)=   "Label36"
         Tab(3).Control(3)=   "Label35"
         Tab(3).ControlCount=   4
         TabCaption(4)   =   "Reparacion"
         TabPicture(4)   =   "FICHREP09.frx":037A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Text4"
         Tab(4).Control(1)=   "Frame12"
         Tab(4).Control(2)=   "Label38"
         Tab(4).Control(3)=   "Label37"
         Tab(4).ControlCount=   4
         TabCaption(5)   =   "Repuestos"
         TabPicture(5)   =   "FICHREP09.frx":0396
         Tab(5).ControlEnabled=   -1  'True
         Tab(5).Control(0)=   "Label33"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).Control(1)=   "Label34"
         Tab(5).Control(1).Enabled=   0   'False
         Tab(5).Control(2)=   "Frame13"
         Tab(5).Control(2).Enabled=   0   'False
         Tab(5).Control(3)=   "Picture6"
         Tab(5).Control(3).Enabled=   0   'False
         Tab(5).Control(4)=   "List3"
         Tab(5).Control(4).Enabled=   0   'False
         Tab(5).Control(5)=   "Frame3"
         Tab(5).Control(5).Enabled=   0   'False
         Tab(5).Control(6)=   "Frame10"
         Tab(5).Control(6).Enabled=   0   'False
         Tab(5).Control(7)=   "Command8"
         Tab(5).Control(7).Enabled=   0   'False
         Tab(5).ControlCount=   8
         TabCaption(6)   =   "Mat.Sumin.Cte."
         TabPicture(6)   =   "FICHREP09.frx":03B2
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Command17(2)"
         Tab(6).Control(1)=   "Command17(1)"
         Tab(6).Control(2)=   "Command17(0)"
         Tab(6).Control(3)=   "Command16"
         Tab(6).Control(3).Enabled=   0   'False
         Tab(6).Control(4)=   "Command15"
         Tab(6).Control(4).Enabled=   0   'False
         Tab(6).Control(5)=   "Command14"
         Tab(6).Control(5).Enabled=   0   'False
         Tab(6).Control(6)=   "MSF1"
         Tab(6).Control(7)=   "MSF2"
         Tab(6).Control(8)=   "MSF3"
         Tab(6).Control(9)=   "msf4"
         Tab(6).Control(10)=   "Label39(2)"
         Tab(6).Control(11)=   "Label39(1)"
         Tab(6).Control(12)=   "Label28(1)"
         Tab(6).Control(13)=   "Label39(0)"
         Tab(6).ControlCount=   14
         Begin VB.CommandButton Command17 
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
            Index           =   2
            Left            =   -67680
            Picture         =   "FICHREP09.frx":03CE
            Style           =   1  'Graphical
            TabIndex        =   108
            ToolTipText     =   "Visualizar Documento Adjunto"
            Top             =   6480
            Width           =   700
         End
         Begin VB.CommandButton Command17 
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
            Index           =   1
            Left            =   -73080
            Picture         =   "FICHREP09.frx":06D8
            Style           =   1  'Graphical
            TabIndex        =   107
            ToolTipText     =   "Visualizar Documento Adjunto"
            Top             =   6480
            Width           =   700
         End
         Begin VB.CommandButton Command17 
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
            Index           =   0
            Left            =   -68280
            Picture         =   "FICHREP09.frx":09E2
            Style           =   1  'Graphical
            TabIndex        =   106
            ToolTipText     =   "Visualizar Documento Adjunto"
            Top             =   4270
            Width           =   700
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Grabar Items Agregados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   -67440
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   4270
            Width           =   3135
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Grabar Consumo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   -72240
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   6480
            Width           =   2535
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Grabar Devolución"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   -66840
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   6480
            Width           =   2535
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Exportar a Excel"
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
            Left            =   8380
            TabIndex        =   92
            Top             =   6480
            Width           =   2115
         End
         Begin VB.Frame Frame10 
            Caption         =   "Frame3"
            Height          =   5070
            Left            =   8940
            TabIndex        =   79
            Top             =   1290
            Width           =   30
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame3"
            Height          =   5100
            Left            =   1995
            TabIndex        =   78
            Top             =   1290
            Width           =   45
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
            Height          =   5010
            ItemData        =   "FICHREP09.frx":0CEC
            Left            =   435
            List            =   "FICHREP09.frx":0CEE
            Sorted          =   -1  'True
            TabIndex        =   77
            Top             =   1320
            Width           =   9870
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5115
            Left            =   -74520
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   76
            Top             =   1140
            Width           =   9870
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5115
            Left            =   -74550
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   75
            Top             =   1110
            Width           =   9870
         End
         Begin VB.Frame Frame9 
            Caption         =   "Documentos Conectados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   -74325
            TabIndex        =   71
            Top             =   705
            Width           =   9450
            Begin VB.CommandButton Command24 
               Caption         =   "Delete"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   945
               Left            =   7875
               Picture         =   "FICHREP09.frx":0CF0
               Style           =   1  'Graphical
               TabIndex        =   74
               ToolTipText     =   "Baja del Archivo Adjunto"
               Top             =   285
               Width           =   700
            End
            Begin VB.CommandButton Command11 
               Caption         =   "View"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   945
               Left            =   8640
               Picture         =   "FICHREP09.frx":0FFA
               Style           =   1  'Graphical
               TabIndex        =   73
               ToolTipText     =   "Visualizar Documento Adjunto"
               Top             =   285
               Width           =   700
            End
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
               Height          =   1635
               Left            =   120
               Sorted          =   -1  'True
               TabIndex        =   72
               Top             =   285
               Width           =   7665
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Documentos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3855
            Left            =   -74325
            TabIndex        =   58
            Top             =   2940
            Width           =   9465
            Begin VB.DriveListBox Drive1 
               Height          =   315
               Left            =   1470
               TabIndex        =   64
               Top             =   255
               Width           =   2220
            End
            Begin VB.DirListBox Dir1 
               Height          =   1440
               Left            =   315
               TabIndex        =   63
               Top             =   675
               Width           =   3375
            End
            Begin VB.FileListBox File1 
               Height          =   1845
               Left            =   3780
               TabIndex        =   62
               Top             =   285
               Width           =   5520
            End
            Begin VB.CommandButton Command7 
               Caption         =   "View"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1000
               Left            =   8565
               Picture         =   "FICHREP09.frx":1304
               Style           =   1  'Graphical
               TabIndex        =   61
               ToolTipText     =   "Visualizar Documento Activo"
               Top             =   2205
               Width           =   750
            End
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
               Left            =   1050
               TabIndex        =   60
               Top             =   3315
               Width           =   7440
            End
            Begin VB.CommandButton Command6 
               Height          =   360
               Left            =   8565
               Picture         =   "FICHREP09.frx":160E
               Style           =   1  'Graphical
               TabIndex        =   59
               ToolTipText     =   "Agregar Documentos"
               Top             =   3285
               Width           =   750
            End
            Begin VB.Label Label30 
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
               Left            =   105
               TabIndex        =   70
               Top             =   2580
               Width           =   810
            End
            Begin VB.Label Label29 
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
               Height          =   615
               Left            =   1050
               TabIndex        =   69
               Top             =   2595
               Width           =   7440
            End
            Begin VB.Label Label28 
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
               TabIndex        =   68
               Top             =   2280
               Width           =   810
            End
            Begin VB.Label Label27 
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
               Left            =   1050
               TabIndex        =   67
               Top             =   2235
               Width           =   7440
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Unidad:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   330
               TabIndex        =   66
               Top             =   315
               Width           =   1065
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Conectar Como:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   105
               TabIndex        =   65
               Top             =   3210
               Width           =   810
            End
         End
         Begin VB.PictureBox Picture4 
            Height          =   5655
            Left            =   -74820
            ScaleHeight     =   5595
            ScaleWidth      =   9435
            TabIndex        =   47
            Top             =   810
            Width           =   9495
            Begin VB.VScrollBar VScroll1 
               Height          =   5600
               Left            =   9195
               TabIndex        =   57
               Top             =   -15
               Width           =   255
            End
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   99000
               Left            =   0
               ScaleHeight     =   99000
               ScaleWidth      =   9195
               TabIndex        =   48
               Top             =   0
               Width           =   9200
               Begin MSComDlg.CommonDialog CommonDialog1 
                  Left            =   75
                  Top             =   75
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
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
                  TabIndex        =   54
                  ToolTipText     =   "Buscar"
                  Top             =   4590
                  Width           =   175
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
                  TabIndex        =   53
                  Top             =   3120
                  Width           =   5835
               End
               Begin VB.PictureBox Picture99 
                  AutoRedraw      =   -1  'True
                  Height          =   1695
                  Index           =   1
                  Left            =   240
                  ScaleHeight     =   1635
                  ScaleWidth      =   2850
                  TabIndex        =   52
                  ToolTipText     =   "Dbl-Click para Maximizar"
                  Top             =   3120
                  Width           =   2905
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
                  TabIndex        =   51
                  ToolTipText     =   "Buscar"
                  Top             =   2190
                  Width           =   175
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
                  TabIndex        =   50
                  Top             =   720
                  Width           =   5835
               End
               Begin VB.PictureBox Picture99 
                  AutoRedraw      =   -1  'True
                  Height          =   1695
                  Index           =   0
                  Left            =   240
                  ScaleHeight     =   1635
                  ScaleWidth      =   2850
                  TabIndex        =   49
                  ToolTipText     =   "Doble-Click para Maximizar"
                  Top             =   720
                  Width           =   2905
               End
               Begin VB.Image Image99 
                  Height          =   10005
                  Index           =   1
                  Left            =   735
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   10005
               End
               Begin VB.Image Image99 
                  Height          =   10005
                  Index           =   0
                  Left            =   735
                  Top             =   1260
                  Visible         =   0   'False
                  Width           =   10005
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
                  TabIndex        =   56
                  Top             =   4590
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
                  Index           =   0
                  Left            =   3240
                  TabIndex        =   55
                  Top             =   2190
                  Width           =   5655
               End
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
            Left            =   -65295
            TabIndex        =   44
            Top             =   735
            Width           =   825
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
               Left            =   85
               Picture         =   "FICHREP09.frx":1758
               Style           =   1  'Graphical
               TabIndex        =   46
               ToolTipText     =   "Agregar Imagen"
               Top             =   370
               Width           =   650
            End
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
               Picture         =   "FICHREP09.frx":18A2
               Style           =   1  'Graphical
               TabIndex        =   45
               ToolTipText     =   "Eliminar Imagen"
               Top             =   1320
               Width           =   650
            End
         End
         Begin VB.Frame Frame6 
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
            Height          =   2060
            Left            =   -74930
            TabIndex        =   31
            Top             =   4785
            Width           =   10620
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00C0C0C0&
               Height          =   425
               Left            =   0
               ScaleHeight     =   360
               ScaleWidth      =   10515
               TabIndex        =   34
               Top             =   315
               Width           =   10580
               Begin VB.Label Label9 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fecha    Detalle                                                    Usuario"
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   225
                  TabIndex        =   36
                  Top             =   15
                  Width           =   10200
               End
               Begin VB.Label Label8 
                  Caption         =   "Label5"
                  Height          =   15
                  Left            =   0
                  TabIndex        =   35
                  Top             =   420
                  Width           =   1380
               End
               Begin VB.Line Line2 
                  X1              =   0
                  X2              =   1155
                  Y1              =   420
                  Y2              =   420
               End
            End
            Begin VB.CommandButton Command3 
               Caption         =   "O.K"
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
               Left            =   8370
               Picture         =   "FICHREP09.frx":1BAC
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   1125
               Width           =   750
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1275
               Left            =   1155
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   32
               Top             =   735
               Width           =   7155
            End
            Begin VB.Line Line9 
               X1              =   9120
               X2              =   10560
               Y1              =   1155
               Y2              =   1155
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "Registrar Anotación"
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
               Left            =   9420
               TabIndex        =   43
               Top             =   1380
               Width           =   960
            End
            Begin VB.Label Label23 
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
               Height          =   285
               Left            =   8370
               TabIndex        =   42
               Top             =   795
               Width           =   2175
            End
            Begin VB.Label Label22 
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
               Height          =   255
               Left            =   105
               TabIndex        =   41
               Top             =   795
               Width           =   885
            End
            Begin VB.Label Label21 
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
               Height          =   1280
               Left            =   8310
               TabIndex        =   40
               Top             =   735
               Width           =   2260
            End
            Begin VB.Label Label20 
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
               Height          =   1275
               Left            =   15
               TabIndex        =   39
               Top             =   735
               Width           =   1170
            End
            Begin VB.Line Line8 
               X1              =   0
               X2              =   10605
               Y1              =   2010
               Y2              =   2010
            End
            Begin VB.Line Line5 
               X1              =   8295
               X2              =   8295
               Y1              =   735
               Y2              =   1995
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Registrar Anotación"
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
               Left            =   9345
               TabIndex        =   38
               Top             =   1365
               Width           =   960
            End
            Begin VB.Line Line6 
               X1              =   8295
               X2              =   10605
               Y1              =   1130
               Y2              =   1130
            End
            Begin VB.Line Line7 
               X1              =   10580
               X2              =   10580
               Y1              =   315
               Y2              =   1995
            End
            Begin VB.Label Label10 
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
               Height          =   225
               Left            =   8400
               TabIndex        =   37
               Top             =   840
               Width           =   4000
            End
         End
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4140
            Left            =   -74930
            TabIndex        =   25
            Top             =   675
            Width           =   10620
            Begin VB.Frame Frame5 
               Caption         =   "Frame5"
               Height          =   3975
               Index           =   0
               Left            =   8115
               TabIndex        =   30
               Top             =   420
               Width           =   15
            End
            Begin VB.Frame Frame4 
               Caption         =   "Frame4"
               Height          =   4005
               Left            =   1230
               TabIndex        =   29
               Top             =   405
               Width           =   15
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00C0C0C0&
               Height          =   435
               Left            =   60
               ScaleHeight     =   375
               ScaleWidth      =   10440
               TabIndex        =   27
               Top             =   195
               Width           =   10500
               Begin VB.Label Label7 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fecha     Detalle                                                   Usuario/Terminal"
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   90
                  TabIndex        =   28
                  Top             =   30
                  Width           =   10200
               End
            End
            Begin VB.ListBox List1 
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
               Height          =   3435
               ItemData        =   "FICHREP09.frx":1EB6
               Left            =   60
               List            =   "FICHREP09.frx":1EB8
               TabIndex        =   26
               Top             =   630
               Width           =   10500
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   10605
               Y1              =   4410
               Y2              =   4410
            End
         End
         Begin VB.PictureBox Picture6 
            Height          =   360
            Left            =   435
            ScaleHeight     =   300
            ScaleWidth      =   9810
            TabIndex        =   80
            Top             =   990
            Width           =   9870
            Begin VB.Label Label31 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Codigo       Descripcion                                                Cantidad"
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   -15
               TabIndex        =   81
               Top             =   -15
               Width           =   11535
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Detalle de la Inspección"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5610
            Left            =   -74850
            TabIndex        =   83
            Top             =   825
            Width           =   10455
         End
         Begin VB.Frame Frame12 
            Caption         =   "Detalle de la Reparación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5610
            Left            =   -74850
            TabIndex        =   84
            Top             =   825
            Width           =   10455
         End
         Begin VB.Frame Frame13 
            Caption         =   "Detalle de Repuestos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5730
            Left            =   255
            TabIndex        =   85
            Top             =   705
            Width           =   10230
         End
         Begin MSFlexGridLib.MSFlexGrid MSF1 
            Height          =   1845
            Left            =   -74955
            TabIndex        =   94
            ToolTipText     =   "Doble Click Para Editar o Agregar"
            Top             =   610
            Width           =   10665
            _ExtentX        =   18812
            _ExtentY        =   3254
            _Version        =   393216
            BackColor       =   16777215
            BackColorFixed  =   12632256
            BackColorSel    =   64
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid MSF2 
            Height          =   1365
            Left            =   -74955
            TabIndex        =   95
            ToolTipText     =   "Doble Click Para Editar o Agregar"
            Top             =   5040
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   2408
            _Version        =   393216
            BackColor       =   16777215
            BackColorFixed  =   12632256
            BackColorSel    =   64
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid MSF3 
            Height          =   1365
            Left            =   -69600
            TabIndex        =   96
            ToolTipText     =   "Doble Click Para Subir el Item a Grilla Superior"
            Top             =   5040
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   2408
            _Version        =   393216
            BackColor       =   16777215
            BackColorFixed  =   12632256
            BackColorSel    =   64
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid msf4 
            Height          =   1365
            Left            =   -74955
            TabIndex        =   102
            ToolTipText     =   "Doble Click Para Editar o Agregar"
            Top             =   2820
            Width           =   10665
            _ExtentX        =   18812
            _ExtentY        =   2408
            _Version        =   393216
            BackColor       =   16777215
            BackColorFixed  =   12632256
            BackColorSel    =   64
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Materiales Yá Suministrados por el Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   -74910
            TabIndex        =   105
            Top             =   345
            Width           =   4275
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Items Seleccionados Para Agregar"
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
            Left            =   -74880
            TabIndex        =   103
            Top             =   2595
            Width           =   4890
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Items Seleccionados Para Devolución"
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
            Left            =   -69600
            TabIndex        =   101
            Top             =   4755
            Width           =   5370
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Items Seleccionados Para Consumo"
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
            Left            =   -74925
            TabIndex        =   100
            Top             =   4755
            Width           =   4890
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   -73005
            TabIndex        =   91
            Top             =   6540
            Width           =   2445
         End
         Begin VB.Label Label37 
            Caption         =   "Usuario Conectado:"
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
            Left            =   -74850
            TabIndex        =   90
            Top             =   6540
            Width           =   1815
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   -73005
            TabIndex        =   89
            Top             =   6540
            Width           =   2445
         End
         Begin VB.Label Label35 
            Caption         =   "Usuario Conectado:"
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
            Left            =   -74850
            TabIndex        =   88
            Top             =   6540
            Width           =   1800
         End
         Begin VB.Label Label34 
            Caption         =   "Usuario Conectado:"
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
            Left            =   150
            TabIndex        =   87
            Top             =   6510
            Width           =   1770
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   1995
            TabIndex        =   86
            Top             =   6510
            Width           =   2445
         End
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H00400000&
         Height          =   8325
         Left            =   10935
         ScaleHeight     =   8265
         ScaleWidth      =   1155
         TabIndex        =   17
         Top             =   15
         Width           =   1215
         Begin VB.CommandButton Command10 
            Height          =   660
            Left            =   105
            Picture         =   "FICHREP09.frx":1EBA
            Style           =   1  'Graphical
            TabIndex        =   93
            ToolTipText     =   "Compresion de imagenes guardadas"
            Top             =   4680
            Width           =   750
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
            Height          =   660
            Left            =   105
            Picture         =   "FICHREP09.frx":22D4
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Guarda los Cambios Realizados"
            Top             =   825
            Width           =   750
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
            Height          =   660
            Left            =   105
            Picture         =   "FICHREP09.frx":25DE
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Imprime Ficha "
            Top             =   6150
            Width           =   750
         End
         Begin VB.CommandButton Command13 
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
            Height          =   660
            Left            =   105
            Picture         =   "FICHREP09.frx":2C48
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   5400
            Width           =   750
         End
         Begin VB.CommandButton Command2 
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
            Height          =   660
            Left            =   105
            Picture         =   "FICHREP09.frx":308A
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   165
            Width           =   750
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   150
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   18
            Top             =   7590
            Width           =   645
            Begin VB.Frame Frame7 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   19
               Top             =   105
               Width           =   12000
            End
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   -195
            Picture         =   "FICHREP09.frx":31D4
            Top             =   7860
            Width           =   1515
         End
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
         Height          =   300
         Left            =   9120
         TabIndex        =   11
         Text            =   " "
         Top             =   885
         Width           =   1485
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10575
         TabIndex        =   0
         Top             =   510
         Width           =   180
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
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
         Left            =   0
         TabIndex        =   97
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label32 
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
         Height          =   315
         Left            =   9645
         TabIndex        =   82
         Top             =   105
         Width           =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   5340
         TabIndex        =   16
         Top             =   150
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N/Serie:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   7155
         TabIndex        =   15
         Top             =   945
         Width           =   1800
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   5325
         TabIndex        =   14
         Top             =   525
         Width           =   1695
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
         Left            =   5850
         TabIndex        =   13
         Top             =   105
         Width           =   2220
      End
      Begin VB.Label Label17 
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
         Left            =   5850
         TabIndex        =   12
         Top             =   490
         Width           =   2220
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   9645
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label15 
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
         Left            =   2310
         TabIndex        =   9
         Top             =   890
         Width           =   2850
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
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
         Left            =   1470
         TabIndex        =   8
         Top             =   890
         Width           =   750
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
         Left            =   1470
         TabIndex        =   7
         Top             =   490
         Width           =   3690
      End
      Begin VB.Label Label12 
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
         Left            =   1455
         TabIndex        =   6
         Top             =   105
         Width           =   3690
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
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
         Left            =   420
         TabIndex        =   5
         Top             =   885
         Width           =   750
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Talon Nro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   7770
         TabIndex        =   4
         Top             =   525
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   8820
         TabIndex        =   3
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   210
         TabIndex        =   2
         Top             =   570
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivo"
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Edi 
      Caption         =   "Edición"
      Visible         =   0   'False
      Begin VB.Menu phl4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAddImg 
         Caption         =   "Agregar Imagen"
         Visible         =   0   'False
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
   Begin VB.Menu matsumcli 
      Caption         =   "Material Suministrado por el Cliente"
      Begin VB.Menu histotalon 
         Caption         =   "Histórico de Mat.Suministrado x N° de Talón"
      End
      Begin VB.Menu matpencli 
         Caption         =   "Materiales Pendientes x Cliente"
      End
   End
End
Attribute VB_Name = "FICHREP09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim INTERVER
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim MODIFIC%
Dim YACAR%
Dim IMGACT%
Dim RUTATMP$
Private Sub BaLog_Click()
  '
  Command24_Click
  '
End Sub

Sub CARGAMATSUM()
   MSF1.Rows = 1: msf2.Rows = 1: MSF3.Rows = 1: msf4.Rows = 1
   TALOX& = XVALO(Label18)
   If TALOX& < 1 Then Exit Sub
   
   SQLX$ = "SELECT  COD_INTE,COD_EXTE,(SUM(CANTINGRE) - SUM(CANTSALID)) AS INGRE "
   SQLX$ = SQLX$ + " FROM DETAMASU WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE TALON =  " & TALOX&
   SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE"
   SQLX$ = SQLX$ + " HAVING (SUM(CANTINGRE) - SUM(CANTSALID)) > 0"
   Set rs = READSET(SQLX$)
   With rs
     Do While Not .EOF
      i& = i& + 1
      MSF1.Rows = i& + 1
      CIXI% = XVALO(!cod_inte)
      MSF1.TextMatrix(i&, 1) = CIXI%
      MSF1.TextMatrix(i&, 2) = CODEXT$(CIXI%)
      MSF1.TextMatrix(i&, 3) = DESCRIT0$(CIXI%)
      MSF1.TextMatrix(i&, 4) = FORMATNUM$(XVALO(!INGRE), "f10.1")
      .MoveNext
     Loop
   End With
  
   rs.Close
   Set rs = Nothing


End Sub

Sub ExportaExcelNet()
    'ARCHIVO PARA ENCABEZADO
    NARCHI2% = FILEOPEN%(LUCOM$ + "ArchivoEnca.TXT", 2, 512)     ' ABRE EL ARCHIVO DE ESCRITURA
    TX$ = AJUSTI$("CODIGO", 18): TX$ = TX$ + ";" 'CODIGO
    TX$ = TX$ + AJUSTI$("DESCRIPCION", 48): TX$ = TX$ + ";"  'DESCRIPCION
    TX$ = TX$ + AJUSTI$("CANTIDAD", 16) 'CANTIDAD
    Print #NARCHI2%, TX$ '
    
    TX$ = "I": TX$ = TX$ + ";" 'CODIGO
    TX$ = TX$ + "I": TX$ = TX$ + ";"  'DESCRIPCION
    TX$ = TX$ + "D" 'CANTIDAD
    Print #NARCHI2%, TX$
    Close #NARCHI2%
    'ARCHIVO PARA DETALLE
    
    NARCHI2% = FILEOPEN%(LUCOM$ + "ArchivoDetalle.TXT", 2, 512)      ' ABRE EL ARCHIVO DE ESCRITURA
    For YX& = 1 To ULTREG&("!BOLREDET")

     RENG1& = YX& + 1
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     Ci% = CVI(Mid$(YY$, 83, 2))
     CODIGO1$ = QuitarAcentos(CODEXT$(Ci%))
     DESCARIP1$ = QuitarAcentos(DESCRIT0$(Ci%))
     CAN = CVD(Mid$(YY$, 89, 8))
     CANTI$ = FORMATNUM$(CAN, "F9.1")
     TX$ = AJUSTI$(SAFETEXT$(CODIGO1$), 18): TX$ = TX$ + ";"
     TX$ = TX$ + AJUSTI$(DESCARIP1$, 48): TX$ = TX$ + ";"
     TX$ = TX$ + AJUSTD$(CANTI$, 16)
     Print #NARCHI2%, TX$
    Next YX&
    Close #NARCHI2%

    Call T_Espera(1)
    
    RUTAEJECUTABLE$ = App.Path + "\"
    On Error Resume Next
    XXX = Shell(RUTAEJECUTABLE$ + "ExportaExcelNet.EXE", vbMinimizedNoFocus)
    If Err Then
        On Error GoTo 0
        Call MENSERR(24, "Error al Generar Informe a Excel")
        Exit Sub
   End If
End Sub

Function QuitarAcentos(ByVal texto As String) As String
    ' Reemplazar caracteres acentuados por sus equivalentes sin acento
    texto = REPLACAR(texto, "á", "a")
    texto = REPLACAR(texto, "é", "e")
    texto = REPLACAR(texto, "í", "i")
    texto = REPLACAR(texto, "ó", "o")
    texto = REPLACAR(texto, "ú", "u")
    texto = REPLACAR(texto, "Á", "A")
    texto = REPLACAR(texto, "É", "E")
    texto = REPLACAR(texto, "Í", "I")
    texto = REPLACAR(texto, "Ó", "O")
    texto = REPLACAR(texto, "Ú", "U")
    texto = REPLACAR(texto, "ñ", "n")
    texto = REPLACAR(texto, "Ñ", "N")
    texto = REPLACAR(texto, "Ø", "O") 'reemplazo por o mayusc o minusc
    texto = REPLACAR(texto, "ø", "o")
    
    
    ' Retornar el texto sin acentos
    QuitarAcentos = texto
End Function
Private Sub Command1_Click()
    '
    Command1.Enabled = False
    If XVALO(Label18) < 1 Then GoTo 99
    Call GRABAIMG
    Call GRABANUMESERIE(Label18, Text10.TEXT, Label12, FECHANUM(Label32)) 'SE AGREGO ESTO PARA QUE EN ESTE SEGUNDO PASO SE PUEDA MODIFICAR EL N.SERIE SI HACE FALTA
    NPX& = XVALO(Label18)
    If PUEDEACTUREP%(NPX&) = 0 Then
        Call COMUNI("Imposible Guardar Cambios. O/R Ya Generada")
        MODIFIC% = 0: Exit Sub
    End If
    Call GRABATEX
    Call GRABAREPUESTOS
    MODIFIC% = 0
    '
99  Command1.Enabled = True
End Sub
Sub GRABANUMESERIE(NPX&, NUSER$, RAXOCLI$, FE%)
       CONDI$ = "NUMREPA = " & NPX&
       Call ACTUREGISTRO("REPARA", "NUMEROSERIE", CONDI$, Left$(NUSER$, 32), REG&, "NOMESS")
       REF2$ = VALOBADA("REPARA", "Referen2", CONDI$, "NOMESS")
       REFEPRE$ = "S/Nro." + TRIM$(Left$(NUSER$, 32) + " - " + Left$(TRIM$(RAXOCLI$), 30) + " - R: " + TRIM$(REF2$) + " - " + FECHATEX$(FE%))
       Call ACTUREGISTRO("REPARA", "Referen", CONDI$, Left$(REFEPRE$, 128), REG&, "NOMESS")
End Sub
Sub GRABAIMG()
    '
    NPX& = XVALO(Label18)
    For JX% = 0 To Image99.Count - 1
      If TRIM(Label99(JX%)) <> "" Then 'NO GRABA IMG EN BLANCO
        ' LECTURA DE LA IMAGEN
        FINAME$ = Label99(JX%)
        f$ = RUTATMP$ & "\" & FINAME$
        XXX$ = LOADFILE$(f$)
        '
        SQLX$ = "SELECT * FROM GALIMREP WHERE NumRepa = '" & NPX& & "' " & " AND " & " Norit_Imag = " & JX% & " "
        Dim GALIREP As ADODB.Recordset
        Set GALIREP = New ADODB.Recordset
        GALIREP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
        '
        With GALIREP
          If .BOF = True And .EOF = True Then
            'GRABACION
             .AddNew   ' YA CONVERTIDO
             On Error Resume Next
             !CodiEmpr = CodiEmp%
             On Error GoTo 0
             !NUMREPA = NPX&
             !Norit_Imag = JX%
             !Nom_Imag = FINAME$
             !Imagen = XXX$
             !Descripcion = Text99(JX%)
             .Update
          Else 'EDICIÓN
             !NUMREPA = NPX&
             !Norit_Imag = JX%
             !Nom_Imag = FINAME$
             !Imagen = XXX$
             !Descripcion = Text99(JX%)
             .Update
          End If
        End With
      End If
    Next JX%
End Sub
Sub GRABARH()
   '
   NPX& = XVALO(Label18)
   If TRIM$(Label18) <> "" Then
      f$ = File1.Path + "\" + Label27
      F2$ = SAFETEXT$(Label27)
      XXX$ = LOADFILE$(f$)
      If XXX$ <> "" Then
10       NUMAPRO& = XVALO(RANDSTRING$(-6))
         If CantRegistros("DOCANRE", "NUMREPA = " & NUMAPRO&, "NOMESS") > 0 Then GoTo 10
         Call CreaRegistro("DOCANRE", "NUMREPA", NUMAPRO&, "NOMESS")
         '
         CONDI$ = "NUMREPA = " & NUMAPRO&
         NORDE& = 1 + XVALO(VALOBADA("DOCANRE", "MAX(NUMREPA)", CONDI$, "NOMESS"))
         Call ACTUREGISTRO("DOCANRE", "FILE_NAME", "NUMREPA = " & NUMAPRO&, F2$, RAFE&, "NOMESS")
         Call ACTUREGISTRO("DOCANRE", "FILE_DATA", "NUMREPA = " & NUMAPRO&, XXX$, RAFE&, "NOMESS")
         Call ACTUREGISTRO("DOCANRE", "NUMREPA", "NUMREPA = " & NUMAPRO&, NPX&, RAFE&, "NOMESS")
       End If
   End If
   '
End Sub


Private Sub Command10_Click()
    
   Dim nombredeimagen(255) As String
   
   NPX& = XVALO(Label18)
   IMGTOT% = CantRegistros("GALIMREP", "NUMREPA = '" & UCase(NPX&) & "'")
   If IMGTOT% = 0 Then
     Call COMUNI("Falta Seleccionar Reparacion o La Misma No Posee Imágenes")
     GoTo 99
   End If
   Screen.MousePointer = 11
   SQLX$ = " SELECT * FROM GALIMREP "
   SQLX$ = SQLX$ + " WHERE NUMREPA = '" & NPX& & "' "
   SQLX$ = SQLX$ + " ORDER BY Norit_Imag ASC "
   Dim IMGTMP As ADODB.Recordset
   Set IMGTMP = New ADODB.Recordset
   IMGTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With IMGTMP
     On Error GoTo 0
     'CARGO LAS IMAGENES
     Do While .EOF = False
        NOIT% = !Norit_Imag
           FINAM$ = SAFETEXT$(!Nom_Imag)
           XXXX$ = SAFETEXT$(!Imagen)
           Call SaveFile(LUCOM$ + "\" + FINAM$, XXXX$)
           '
           ' Comprime el archivo llevándolo a JPG
           fMain.Label2 = LUCOM$ + "\" + FINAM$
           fMain.mnuOpen_Click
           Dim SaveForm As fSaveJPG
           Set SaveForm = New fSaveJPG
           SaveForm.SaveImage fMain.m_Image, LUCOM$ + "\1" + FINAM$
           SaveForm.cmdFinish_Click (1)
           Set SaveForm = Nothing
           '
           ' Guarda el archivo en el mismo registro pero esta vez como JPG
           XXX$ = LOADFILE(LUCOM$ + "\1" + FINAM$)
           !Imagen = XXX$
           .Update
      .MoveNext
     Loop
     MODIFIC% = 0
   End With
   
   Screen.MousePointer = 1
   Call COMUNI("Proceso de Compresion Finalizado")
   
99 Screen.MousePointer = 1

End Sub

Private Sub Command11_Click()

    Command11.Enabled = False
    If List2.ListIndex > -1 Then
        nombre$ = TRIM$(Mid$(List2.List(List2.ListIndex), 1, 64))
        NOMBREORIG$ = TRIM$(Mid$(List2.List(List2.ListIndex), 65))
        CONDI$ = "NUMREPA=" & Label18 & " AND NOMASIGNADO='" & nombre$ & "'"
        XXX$ = VALOBADA("DOCANRE", "FILE_DATA", CONDI$)
        ARCHIVOTMP$ = RUTATMP$ & "\" & NOMBREORIG$
        Call SaveFile(ARCHIVOTMP$, XXX$)
        Call MUESTRADOC3(ARCHIVOTMP$)
    End If
    Command11.Enabled = True
    
End Sub

Private Sub Command14_Click()
   'INFORMAR DEVOLUCION
   TALONX& = XVALO(Label18)
   If TALONX& < 1 Then Exit Sub
   If MSF3.Rows < 2 Then Exit Sub

   Command14.Enabled = False
   NEWID& = PROXID_DETAMASU
   VUELTA% = 0: CONTITULO% = 1
   For i& = 1 To MSF3.Rows - 1
       CIXI% = XVALO(MSF3.TextMatrix(i&, 1))
       Codigo$ = MSF3.TextMatrix(i&, 2)
       DESCRI$ = MSF3.TextMatrix(i&, 3)
       CANTI = XVALO(MSF3.TextMatrix(i&, 4))
       Call MOVI_SUMINISTRADOS(NEWID&, CIXI%, "DV", CANTI * (-1), XVALO(Label11), TALONX&, HOY(HO$), "Devolución de Suministros")
       Call ANOTAMATSUM(TALONX&, Codigo$, CANTI, DESCRI$, "Devolución de Materiales Suministrados por el Cliente:", CONTITULO%)
       If VUELTA% < 1 Then
          CONTITULO% = 0
        End If
   Next i&
   Call PRESRE09.AGREGA_ANOTAREPARA(TALONX&, String$(88, "="))

   FILTX$ = TRIM$(Str$(TALONX&)) & ";" & TRIM$(NEWID&)
   Call STDFORM("REDV-ARC", FILTX$)

   Call CARGAMATSUM ' para refrescar
   Command14.Enabled = True
End Sub

Sub ANOTAMATSUM(NROREP&, Codigo$, CANTI, DESCRI$, DESCRITIT$, CONTITULO%)
       
       If CONTITULO% > 0 Then
          TTXYZ$ = Space$(92)
'          NROREP& = XVALO(Label18)
          HOII% = HOY(HOS$)
          Call REPLA(TTXYZ$, HOS$, 1, 8)
          Call REPLA(TTXYZ$, USERTER$, 69, 24)
          Call REPLA(TTXYZ$, DESCRITIT$, 11, 56)
          Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
          
          TTXYZ$ = Space$(92)
          Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
          
          Call REPLA(TTXYZ$, AJUSTD(SAFETEXT$("Cant."), 6), 11, 6)
          Call REPLA(TTXYZ$, "Código", 18, 18)
          Call REPLA(TTXYZ$, "Descripción", 37, 20)
          Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
        End If
        
        TTXYZ$ = Space$(92)
        Call REPLA(TTXYZ$, AJUSTD(SAFETEXT$(CANTI), 6), 11, 6)
        Call REPLA(TTXYZ$, Codigo$, 18, 18)
        Call REPLA(TTXYZ$, DESCRI$, 37, 20)
        Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)

End Sub
Private Sub Command15_Click()
   'INFORMAR CONSUMO DE ELEMENTOS
   NROREP& = XVALO(Label18)
   If NROREP& < 1 Then Exit Sub
   If msf2.Rows < 2 Then Exit Sub
   
   Command15.Enabled = False
   VUELTA% = 0: CONTITULO% = 1
   NEWID& = PROXID_DETAMASU
   For i& = 1 To msf2.Rows - 1
       CIXI% = XVALO(msf2.TextMatrix(i&, 1))
       Codigo$ = msf2.TextMatrix(i&, 2)
       DESCRI$ = msf2.TextMatrix(i&, 3)
       CANTI = XVALO(msf2.TextMatrix(i&, 4))
       Call MOVI_SUMINISTRADOS(NEWID&, CIXI%, "CF", CANTI * (-1), XVALO(Label11), NROREP&, HOY(HO$), "Consumo de Fábrica")
       Call ANOTAMATSUM(NROREP&, Codigo$, CANTI, DESCRI$, "Detalle de Consumo de Material Suministrado por el Cliente:", CONTITULO%)
       If VUELTA% < 1 Then
          CONTITULO% = 0
        End If
   Next i&
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, String$(88, "="))
   FILTX$ = TRIM$(Str$(NROREP&)) & ";" & TRIM$(NEWID&)
   Call STDFORM("CONS-ARC", FILTX$)

   Call CARGAMATSUM ' para refrescar

   Command15.Enabled = True

End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   'ITEMS AGREGADOS
   NROREP& = XVALO(Label18)
   If NROREP& < 1 Then GoTo 99
   If msf4.Rows < 2 Then GoTo 99
   NEWID& = PROXID_DETAMASU
   VUELTA% = 0: CONTITULO% = 1

   For i& = 1 To msf4.Rows - 1
       CIXI% = XVALO(msf4.TextMatrix(i&, 1))
       Codigo$ = msf4.TextMatrix(i&, 2)
       DESCRI$ = msf4.TextMatrix(i&, 3)
       CANTI = XVALO(msf4.TextMatrix(i&, 4))
       Call MOVI_SUMINISTRADOS(NEWID&, CIXI%, "BR", CANTI, XVALO(Label11), XVALO(Label18), HOY(HO$), "Agregado Posterior")
       Call ANOTAMATSUM(NROREP&, Codigo$, CANTI, DESCRI$, "Agregado de Materiales Suministrados por el Cliente:", CONTITULO%)
       If VUELTA% < 1 Then
          CONTITULO% = 0
        End If
   Next i&
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, String$(88, "="))
   
   Call T_Espera(2) 'le agrego una espera por que no salia el detalle de los materiales sumin. cliente
   
   FILTX$ = TRIM$(Str$(NROREP&)) & ";" & TRIM$(NEWID&)
   Call STDFORM("AGRE-ARC", FILTX$)
   '
   Call CARGAMATSUM ' para refrescar
99 Command16.Enabled = True
End Sub

Private Sub Command17_Click(Index As Integer)
   Command17(Index).Enabled = False
   TALONX& = XVALO(Label18)
   If TALONX& < 1 Then GoTo 99
   
   If Index = 0 Then
     CONDI$ = "REFERENCIA = 'Agregado Posterior'"
     TITU$ = "Agregados Posterior a la Recepeción"
   ElseIf Index = 1 Then
     CONDI$ = "REFERENCIA = 'Consumo de Fábrica'"
     TITU$ = "Consumos de Fábrica"
   ElseIf Index = 2 Then
     CONDI$ = "REFERENCIA = 'Devolución de Suministros'"
     TITU$ = "Devoluciones de Suministros"
   End If
   '
   CONDI$ = CONDI$ + " and TALON = " & TALONX&
   Call CARGALISEL("IMPRDEMA", "DETAMASU", "ID/F7;TALON/F7;NUME_CLIE/F7+DEUDOR12/A32;FECHMOV/D8", CONDI$, "NOMESS/DISTINCT")
   Call SELECHO("IMPRDEMA/" & TITU$)
   '
   ID& = XVALO(VALACT1$("IMPRDEMA"))
   TXTX$ = TRIM$(VALACT2$("IMPRDEMA"))
   ACUMU$ = "" 'TOMO EL NUMERO DE VALACT2 DESDE EL COMIENZO HASTA EL PRIMER ESPACIO
   For i& = 1 To Len(TXTX$)
     LEIDO$ = Mid$(TXTX$, i&, 1)
     If IsNumeric(LEIDO$) Then
       ACUMU$ = ACUMU$ & LEIDO$
     Else
       Exit For
     End If
   Next i&
   '*********************
   
   TALONX& = XVALO(ACUMU$)
   If TALONX& > 0 And ID& > 0 Then
      FILTX$ = TRIM$(Str$(TALONX&)) & ";" & TRIM$(ID&)
      
      If Index = 0 Then Call STDFORM("AGRE-ARC", FILTX$)
      If Index = 1 Then Call STDFORM("CONS-ARC", FILTX$)
      If Index = 2 Then Call STDFORM("REDV-ARC", FILTX$)
   
   End If
   '
99 Command17(Index).Enabled = True
End Sub

   Sub HISTO_CONSU_MAT_SUM(NORDREP&)

      Campos$ = "ci/f0;IdTrans./F10;T.Mov./A5;Código/A14;Descripción/A32;Ingreso/f8.1;Salida/F8.1;Saldo/F8.1"
      Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
      Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
      If ANTOT < 5500 Then ANTOT = 5500
      FRMZELECHO.Width = ANTOT + 1500
      FRMZELECHO.Caption = "Facturas Anticipo de Cliente N°: " & NC
      TALOX& = XVALO(NORDREP&)
      FRMZELECHO.Caption = "Histórico de Materiales Suministrados por el Cliente (Talón: " & TALOX& & ")"
  
      SQLX$ = "SELECT  COD_INTE,COD_EXTE,CANTINGRE ,CANTSALID,ID,codimovi "
      SQLX$ = SQLX$ + " FROM DETAMASU WITH(NOLOCK)"
      SQLX$ = SQLX$ + " WHERE TALON =  " & TALOX&
      SQLX$ = SQLX$ + " ORDER BY ID "
      Set rs = READSET(SQLX$)
      With rs
        Do While Not .EOF
         i& = i& + 1
         
         FRMZELECHO.msf2.Rows = i& + 1
         CIXI% = XVALO(!cod_inte)
         FRMZELECHO.msf2.TextMatrix(i&, 1) = CIXI%
         FRMZELECHO.msf2.TextMatrix(i&, 2) = XVALO(!ID)
         FRMZELECHO.msf2.TextMatrix(i&, 3) = SAFETEXT$(!CODIMOVI)
         FRMZELECHO.msf2.TextMatrix(i&, 4) = CODEXT$(CIXI%)
         FRMZELECHO.msf2.TextMatrix(i&, 5) = DESCRIT0$(CIXI%)
         FRMZELECHO.msf2.TextMatrix(i&, 6) = FORMATNUM$(XVALO(!cantingre), "F10.1")
         FRMZELECHO.msf2.TextMatrix(i&, 7) = FORMATNUM$(XVALO(!CANTSALID), "F10.1")
         SALDO = SALDO + (XVALO(!cantingre) - XVALO(!CANTSALID))
         FRMZELECHO.msf2.TextMatrix(i&, 8) = FORMATNUM$(SALDO, "F10.1")
         .MoveNext
        Loop
      End With
     
      rs.Close
      Set rs = Nothing

      If FRMZELECHO.msf2.Rows <= 1 Then Exit Sub

50    FRMZELECHO.Show 1
      '
      ID& = XVALO(RESP_ZELE_VECT(2))
      If ID& > 0 Then
        CODMOV$ = RESP_ZELE_VECT(3)
        FILTX$ = TRIM$(Str$(NORDREP&)) & ";" & TRIM$(ID&)
        Select Case CODMOV$
          Case "BR"
             Call STDFORM("AGRE-ARC", FILTX$)
             GoTo 50
          Case "CF"
            Call STDFORM("CONS-ARC", FILTX$)
            GoTo 50
          Case "DV"
           Call STDFORM("REDV-ARC", FILTX$)
           GoTo 50
          Case Else
            Call COMUNI("No se Encontró el Documento Impreso")
            GoTo 50
        End Select
      End If
End Sub

  Sub MAT_PEN_X_CLIENTE(NC)

      Campos$ = "ci/f0;Talón/F9;Código/A14;Descripción/A36;Saldo/F8.1"
      Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
      Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
      If ANTOT < 5500 Then ANTOT = 5500
      FRMZELECHO.Width = ANTOT + 1500
      FRMZELECHO.Caption = "Facturas Anticipo de Cliente N°: " & NC
      TALOX& = XVALO(NORDREP&)
      FRMZELECHO.Caption = "Materiales Suministrados Pendientes del Cliente " & NC
      SQLX$ = "select (sum(cantingre)-sum(cantsalid)) as saldo ,cod_inte,TALON  "
      SQLX$ = SQLX$ + " from detamasu WITH(NOLOCK)"
      SQLX$ = SQLX$ + " Where nume_clie = " & NC
      SQLX$ = SQLX$ + " group by cod_inte,TALON having   (sum(cantingre)-sum(cantsalid)) > 0 "
      SQLX$ = SQLX$ + "  order by TALON,cod_inte"
      Set rs = READSET(SQLX$)
      With rs
        If rs.EOF Then
           Call COMUNI("No se Detectaron Materiales Suministrados Pendientes para el Cliente N°: " & NC & ""):  Exit Sub
        End If
        Do While Not .EOF
         i& = i& + 1
         
         FRMZELECHO.msf2.Rows = i& + 1
         CIXI% = XVALO(!cod_inte)
         
         FRMZELECHO.msf2.TextMatrix(i&, 1) = CIXI%
         FRMZELECHO.msf2.TextMatrix(i&, 2) = XVALO(!TALON)
         FRMZELECHO.msf2.TextMatrix(i&, 3) = CODEXT$(CIXI%)
         FRMZELECHO.msf2.TextMatrix(i&, 4) = DESCRIT0$(CIXI%)
         FRMZELECHO.msf2.TextMatrix(i&, 5) = FORMATNUM$(XVALO(!SALDO), "F10.1")
'         FRMZELECHO.msf2.TextMatrix(i&, 6) = XVALO(!ID)
'         FRMZELECHO.msf2.TextMatrix(i&, 7) = SAFETEXT$(!CODIMOVI)
         .MoveNext
        Loop
      End With
     
      rs.Close
      Set rs = Nothing

      If FRMZELECHO.msf2.Rows <= 1 Then Exit Sub
50    FRMZELECHO.Show 1
'      ID& = XVALO(RESP_ZELE_VECT(6))
'      NORDREP& = XVALO(RESP_ZELE_VECT(2))
'      If ID& > 0 Then
'        CODMOV$ = RESP_ZELE_VECT(7)
'        FILTX$ = TRIM$(Str$(NORDREP&)) & ";" & TRIM$(ID&)
'        Select Case CODMOV$
'          Case "BR"
'             Call STDFORM("AGRE-ARC", FILTX$)
'             GoTo 50
'          Case "CF"
'            Call STDFORM("CONS-ARC", FILTX$)
'            GoTo 50
'          Case "DV"
'           Call STDFORM("REDV-ARC", FILTX$)
'           GoTo 50
'          Case Else
'            Call COMUNI("No se Encontró el Documento Impreso")
'            GoTo 50
'        End Select
'      End If
      
End Sub


Private Sub Command19_Click()
    '
    NROREP$ = TRIM$(InputBox$("ORDEN DE REPARACION", "Ingrese OR (OR-000233)"))
    If NROREP$ = "" Then Exit Sub
    
    Call BORRAREGISTROS("RESERVA", "IDSUBOR = '" & NROREP$ & "'", REGAF&, "NOMESS")
    
    NPX& = XVALO(Label18)

    SQLX1$ = "SELECT * FROM RESERVA"
    SQLX1$ = SQLX1$ + " WHERE COD_INTE < 1 "
    Dim RESETMP As ADODB.Recordset
    Set RESETMP = New ADODB.Recordset
    RESETMP.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

    

    SQLX$ = "SELECT * FROM REPUREPA "
    SQLX$ = SQLX$ + " WHERE NUMREPA = '" & UCase(NPX&) & "' AND ORIGEN=0"
    Dim REPU As ADODB.Recordset
    Set REPU = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JJ& = 0
    With REPU
      On Error GoTo 0
      'CARGO LOS REPUESTOS
      Do While .EOF = False
         CAN# = XVALO(!CANT_SOL)
         CIIX% = XVALO(!cod_inte)
         Codigo$ = CODEXT$(CIIX%)
         Des$ = DESCRIT0$(CIIX%)
         
         If CIIX% > 0 Then
           If CIIX% <= NUMAS% Then
             RESETMP.AddNew
             On Error Resume Next
             RESETMP!CodiEmpr = CodiEmp%
             On Error GoTo 0
             RESETMP!cod_inte = CIIX%
             RESETMP!Cod_Exte = Codigo$
             RESETMP!Descrip = Des$
             RESETMP!IdSubOr = NROREP$
             RESETMP!FechRes = Now
             RESETMP!CANTRES = CAN#
             RESETMP!CantCons = 0
             RESETMP!Saldo_Entre = CAN#
             RESETMP!IDENLOTE = ""
             RESETMP!DEPOSITO = 0
             RESETMP!UsoUnit = CAN#
                KK& = KK& + 1
             RESETMP!NuOrItem = KK&
             RESETMP.Update
           End If
         End If
         .MoveNext
      Loop
      
   End With
 
   On Error Resume Next
   RESETMP.Close
   Set RESETMP = Nothing
   REPU.Close
   
   On Error GoTo 0
   '
99

End Sub

Private Sub Command2_Click()
    '
    If MODIFIC% = 1 Then
        OPX% = COMALTER%("Se Han Modificado Los Datos\¿Desea Guardar Los Cambios?\\Si, Guardar\No, Cancelar")
        If OPX% = 1 Then Call Command1_Click
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Label18 = ""
    Unload Me
    '
End Sub

Private Sub Command24_Click()
   '
   Command24.Enabled = False
   If List2.ListIndex > -1 Then
        NPX& = XVALO(Label18)
        'VALIDA SI SE PUEDE MODIFICAR
        If PUEDEACTUREP%(NPX&) = 0 Then
          Call COMUNI("Se Ha Detectado que la O/R Ya Generada")
          'MODIFIC% = 0: GoTo 99
        End If

        nombre$ = TRIM$(Mid$(List2.List(List2.ListIndex), 1, 64))
        CONDI$ = "NUMREPA=" & Label18 & " AND NOMASIGNADO='" & nombre$ & "'"
        Call BORRAREGISTROS("DOCANRE", CONDI$, REG&)
        List2.RemoveItem (List2.ListIndex)
        List2.Refresh
   End If
99 Command24.Enabled = True
    '
End Sub

Private Sub Command3_Click()
Command3.Enabled = False
   NPX& = Val(Label18)
   If NPX& > 0 Then
     If TRIM$(Text1) <> "" Then
       Call FRATEXTO(Text1, 56)
       For KKI% = 1 To CARETEX%
         TTXX$ = Space$(93)
         Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTXX$, HOS$, 1, 8)
           Call REPLA(TTXX$, USERTER$, 69, 24)
         End If
         Call PRESRE09.AGREGA_ANOTAREPARA(NPX&, TTXX$)
       Next KKI%
       '
       Call PRESRE09.AGREGA_ANOTAREPARA(NPX&, String$(88, "="))
       Text1 = ""
       '
       List1.Clear
       NPX& = XVALO(Label18)
       If NPX& < 1 Then Exit Sub
20     SQLX$ = "SELECT * FROM REPARA"
       SQLX$ = SQLX$ + " WHERE NumRepa = " & NPX& & ""
       Dim ANOTRE As New ADODB.Recordset
       Set ANOTRE = New ADODB.Recordset
25     On Error Resume Next
       ANOTRE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 25
       End If
       On Error GoTo 0
       '
       With ANOTRE
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
            On Error GoTo 0
            XX$ = SAFETEXT$(!AnotaRep): LONGI& = Len(XX$)
            INICIO& = 1
            Do While INICIO& < LONGI&
                FIN& = InStr(INICIO&, XX$, Chr$(10))
                If FIN& > 0 Then
                   TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
                   If FIN& < LONGI& Then A = Asc(Mid$(XX$, FIN& + 1, 1))
                   List1.AddItem TTXX$
                   b = 1: If A = 13 Then b = 2
                   INICIO& = FIN& + b
                Else
                   TTXX$ = Mid$(XX$, INICIO&)
                   List1.AddItem TTXX$
                   Exit Do
                End If
            Loop
         End If
         On Error GoTo 0
       End With
       ANOTRE.Close
       Set ANOTRE = Nothing
     End If
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
'    If CantRegistros("EQUIPOS", "codigo_equ= '" & TRIM$(Text1.TEXT) & "'") > 0 Then
      Call CARGABLOIMG   'AGREGA COMBO DE IMAGEN
      On Error Resume Next
        'SE POSICIONA EN LA ULTIMA IMAGEN
         VScroll1.Value = VScroll1.Max
      On Error GoTo 0
'    End If
    Command4.Enabled = True
End Sub
Sub CARGABLOIMG() 'CARGA BLOQUE DE IMAGEN
  '
  'Agrega un Conjunto de Imagen.
  'IMAGEN
  NUPIC% = Picture99.Count
  Load Picture99(NUPIC%)
  Picture99(NUPIC%).Top = Picture99(NUPIC% - 1).Top + Picture99(NUPIC% - 1).Height + INTERVER
  '
  If Picture99(NUPIC%).Top + Picture99(NUPIC%).Height > Picture5.Height - 500 Then
    Picture5.Height = Picture99(NUPIC%).Top + Picture99(NUPIC%).Height + 500
  End If
  '
  Picture99(NUPIC%).Visible = True
  VScroll1.Max = (Picture5.Height - Picture4.Height) / 10 + 1
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

Private Sub Command5_Click()
   Command5.Enabled = False
   '
   NROREP& = XVALO(Label18)
   If NROREP& < 1 Then GoTo 99
   '
   'REPORTE DE INSPECCION
   Call STDFORM("REP1-ARC", TRIM$(Str$(NROREP&)), "PRINT")
   
   'REPORTE DE IMAGENES
   Call STDFORM("REP3-ARC", TRIM$(Str$(NROREP&)), "PRINT")
   
   'REPORTE DE REPUESTOS
   Call STDFORM("REP2-ARC", TRIM$(Str$(NROREP&)), "PRINT")
   '
   'REPORTE DE REPARACION
   Call STDFORM("REP4-ARC", TRIM$(Str$(NROREP&)), "PRINT")
   '
99 Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
    NPX& = XVALO(Label18)
    If TRIM(Label18) <> "" Then
      f$ = File1.Path + "\" + Label27
      F2$ = SAFETEXT$(Label27)
      F3$ = TRIM$(Text2)
      XXX$ = LOADFILE$(f$)
      If XXX$ <> "" And F3$ <> "" Then
10       NUMAPRO& = XVALO(RANDSTRING$(-6))
         If CantRegistros("DOCANRE", "NUMREPA = " & NUMAPRO&, "NOMESS") > 0 Then GoTo 10
         Call CreaRegistro("DOCANRE", "NUMREPA", NUMAPRO&)
         '
         CONDI$ = "NUMREPA = " & NUMAPRO&
'         CONDI$ = CONDI$ + "AND FILE_NAME <> " & F2 & " "
         Call ACTUREGISTRO("DOCANRE", "NomAsignado", "NUMREPA = " & NUMAPRO&, F3$, RAFE&, "NOMESS")
         Call ACTUREGISTRO("DOCANRE", "FILE_NAME", "NUMREPA = " & NUMAPRO&, F2$, RAFE&, "NOMESS")
         Call ACTUREGISTRO("DOCANRE", "FILE_DATA", "NUMREPA = " & NUMAPRO&, XXX$, RAFE&, "NOMESS")
         Call ACTUREGISTRO("DOCANRE", "NUMREPA", "NUMREPA = " & NUMAPRO&, NPX&, RAFE&, "NOMESS")
         
         Call REPLA(XX$, AJUSTI$(Text2, 64), 1, 64)
         Call REPLA(XX$, TRIM$(Label27), 65, 128)
         List2.AddItem XX$
         List2.Refresh
       End If
   End If
   '
   DoEvents
   Command6.Enabled = True
   '
End Sub

Private Sub Command7_Click()
    '
    Command7.Enabled = False
    ARDOC$ = Label29 + "\" + Label27
    Call MUESTRADOC3(ARDOC$)
    Command7.Enabled = True
   
End Sub
Sub MUESTRADOC3(ARDOC$)
    '
    Call ShellExecute(Me.hwnd, "Open", ARDOC$, "", "", 1)
    '
End Sub


Sub CARLISPREF()
   '
   List2.Clear
   For U& = 1 To ULTREG&("DOCPREFE")
     XX$ = REGLEIDO$("DOCPREFE", U&)
     YY$ = Space$(320)
     Call REPLA(YY$, XX$, 1, 320)
     List2.AddItem YY$
   Next U&
   List2.Refresh
   DoEvents
   '
End Sub


'
Private Sub Command8_Click()
   Command8.Enabled = False
   'valido que haya informacion en el list
'   CAN = 100
'     If CAN <= 9999 Then
'          Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 1, 2), 73, 6)
'     Else
'          Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 73, 6)
'     End If
'     Call REPLA(TTXX$, 974000110, 1, 16)
'     Call REPLA(TTXX$, DESCRIT0$(CODINT("974000110")), 14, 52)
'     List3.AddItem TTXX$
'
   
   If List3.ListCount < 1 Then GoTo 99
   '
   Call ExportaExcelNet
   Exit Sub
   
   'presenta cuadro de dialogo
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
   ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   'CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FILTER = "Excel Files(*.xls)|*.xls|Excel Files(*.xlsx)|*.xlsx"

   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   ARCHIEX$ = UCase$(CommonDialog1.FileName)
   
   'si selecciona archivo genera excel
   If TRIM$(ARCHIEX$) <> "" Then
      Call GENERAXCEL(ARCHIEX$)
   End If

   
   
99 Command8.Enabled = True

End Sub
'Private Sub Command8_Click()
'   Command8.Enabled = False
'   'valido que haya informacion en el list
'   If List3.ListCount < 1 Then GoTo 99
'   '
'   Screen.MousePointer = 11
'   Campos$ = "Código/A24;Descripción/A48;Cantidad/F9.1"
'   Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
'   Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
'
'
'   For YX& = 1 To ULTREG&("!BOLREDET")
'     RENG1& = YX& + 1
'     YY$ = REGLEIDO$("!BOLREDET", YX&)
'     Ci% = CVI(Mid$(YY$, 83, 2))
'     CODIGO1$ = CODEXT$(Ci%)
'     DESCARIP1$ = DESCRIT0$(Ci%)
'       CAN = CVD(Mid$(YY$, 89, 8))
'     CANTI = FORMATNUM$(CAN, "F9.1")
'     GRILLAGRAL.GRID.Rows = YX& + 1
'     GRILLAGRAL.GRID.TextMatrix(YX&, 1) = CODIGO1$
'     GRILLAGRAL.GRID.TextMatrix(YX&, 2) = DESCARIP1$
'     GRILLAGRAL.GRID.TextMatrix(YX&, 3) = CANTI
'   Next YX&
'
'   Screen.MousePointer = 1
'
'   TX$ = "Detalle de Repuestos"
'   GRILLAGRAL.Caption = TX$
'    '*******************************************
'    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
'    GRILLAGRAL.mnuMenu1.Visible = True
'    GRILLAGRAL.GRID.ZOrder 0
'    GRILLAGRAL.TXTBUSCAR = ""
'
'    GRILLAGRAL.SinEfectoClickEnGrilla% = 1
'
'
'    GRILLAGRAL.Show 1
'    Dim MSF As msFlexGrid
'    If GRILLAGRAL.MENU1 > 0 Then
'        Set MSF = GRILLAGRAL.GRID
'        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
'        Call CARGA_TABLA_IMPRE(MSF, Campos$, "DETAREPU", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
'        Call FINAL("?" & RET_NAMECONS$)
'        GRILLAGRAL.MENU1 = 0
'        Screen.MousePointer = 1
'    End If
'99 Command8.Enabled = True
'
'End Sub
Sub GENERAXCEL(NombreDeArchivo$)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim texto As String
   '
   On Error GoTo 20
   'ver esto si soluciona el error cuando se inicializa el objeto
   'https://www.lawebdelprogramador.com/foros/Visual-Basic/1525733-Office-2016-elimina-la-referencia-a-Excel-en-Visual-Basic.html
   On Error Resume Next
   Set XLAPP = New Excel.Application
   Set XLAPP = CreateObject("Excel.Application")
   XLAPP.Visible = True ' Hacer visible la aplicación de Excel
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   XLSHEET.Cells(1, 1) = "Código"
   XLSHEET.Cells(1, 2) = "Descripción"
   XLSHEET.Cells(1, 3) = "Cantidad"

   For YX& = 1 To ULTREG&("!BOLREDET")
     RENG1& = YX& + 1
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     Ci% = CVI(Mid$(YY$, 83, 2))
     CODIGO1$ = CODEXT$(Ci%)
     DESCARIP1$ = DESCRIT0$(Ci%)
       CAN = CVD(Mid$(YY$, 89, 8))
     CANTI = FORMATNUM$(CAN, "F9.1")
     XLSHEET.Cells(RENG1&, 1) = CODIGO1$
     XLSHEET.Cells(RENG1&, 2) = DESCARIP1$
     XLSHEET.Cells(RENG1&, 3) = CANTI
   Next YX&

''   For I& = 1 To GRID.Rows
''       RENG1& = I&
''       For COLU1& = 1 To GRID.Cols
''          VALORX = GRID.TextMatrix(I& - 1, COLU1 - 1)
''          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
''             FEXU$ = VALORX
''             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
''          End If
''          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
''          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
''          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
''             If TRIM$(CStr(VALORX)) = "" Then
''                VALORX = 0
''             End If
''          End If
''          XLSHEET.Cells(RENG1&, COLU1&) = VALORX
''       Next COLU1&
''   Next I&
   '
   On Error GoTo 30
   XLSHEET.SaveAs NombreDeArchivo$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDeArchivo$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub


Private Sub Command9_Click()
    '
   If MODIFIC% = 1 Then
        OPX% = COMALTER%("Se Han Modificado Los Datos\¿Desea Guardar Los Cambios?\\Si, Guardar\No, Cancelar")
        If OPX% = 1 Then Call Command1_Click
        MODIFIC% = 0
   End If
   '
   CONDI$ = "NROPED < 0.5"
   JJ& = CantRegistros("REPARA", CONDI$)
   If JJ& < 1 Then
     Call COMUNI("No Hay Ordenes Pendientes de Reparación.")
     GoTo 99
   End If
   '
   CONDI$ = "NuOrdRep >= 0 AND STATUS < 9 ORDER BY NUMREPA"
   Call CARGALISEL("PREPAEMI", "REPARA", "NumRepa/I10;Referen/A128", CONDI$)
   Call FRESHECHO("PREPAEMI")
   Call SELECHO("PREPAEMI")
   '
   NREP% = Val(VALACT1$("PREPAEMI"))
   If NREP% > 0 Then
      Label18 = TRIM$(Str$(NREP%))
   End If
   '
   MODIFIC% = 0
   '
99 Command9.Enabled = True
End Sub

Private Sub Command99_Click(Index As Integer)
    '
     If TRIM$(Label18) = "" Then Exit Sub
     '
     Command99(Index).Enabled = False
     '
10   CommonDialog1.CancelError = True
     CommonDialog1.InitDir = ULODATOS$ & "\"
     CommonDialog1.DialogTitle = "Archivo Origen de Datos"
        ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
     CommonDialog1.FILTER = "Archivo de Imagen(JPG; BMP; WMF)|*.JPG; *.BMP; *.WMF|"
     CommonDialog1.FilterIndex = 1
     On Error GoTo 99
     CommonDialog1.ShowOpen
     On Error GoTo 0
     '
     MODIFIC% = 1
     DoEvents
     Screen.MousePointer = 11
     '
     APLANO$ = UCase$(CommonDialog1.FileName)
     XXX$ = LOADFILE$(APLANO$)
     i = 1
     While InStr(i, APLANO$, "\") > 0
        i = InStr(i, APLANO$, "\") + 1
     Wend
     APLANO$ = Mid$(APLANO$, i)
     Call SaveFile(RUTATMP$ & "\" & APLANO$, XXX$)
     Label99(Index) = ""
     Label99(Index) = APLANO$
     '
     On Error GoTo 0
     '
99   Screen.MousePointer = 1
     Command99(Index).Enabled = True
     On Error GoTo 0
     '
End Sub

Private Sub DesCam_Click()
  '
  Call Command2_Click
  '
End Sub
Private Sub Dir1_Change()
     '
     Label29.Caption = Dir1.Path
     File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
     '
     On Error GoTo 20
     Dir1.Path = Drive1.Drive
     GoTo 99
     '
20   Resume 21
21   On Error GoTo 0
     Call MENSERR(24, "Unidad no Disponible.")
     Drive1.Drive = Left$(CurDir$, 2)
     '
99   On Error GoTo 0
End Sub

Private Sub File1_Click()
     '
     Label27 = File1.FileName
     Label29.Caption = Dir1.Path
     '
End Sub

Private Sub Form_Load()
     '
     SSTAB.Tab = 0

     HOII% = HOY(HOS$)
     Label22 = " " + HOS$
     Label23 = Left$(USERNAME, 17)
     Label33 = Left$(USERNAME, 17)
     Label36 = Left$(USERNAME, 17)
     Label38 = Left$(USERNAME, 17)
     '
     Picture5.Height = Picture4.Height
     VScroll1.Min = 0
     VScroll1.Max = (Picture5.Height - Picture4.Height) / 10 + 1
     VScroll1.Value = 0
     '
     INTERVER = Picture99(1).Top - Picture99(0).Top - Picture99(0).Height
     '
     Label27.Caption = Dir1.Path
     File1.Path = Dir1.Path
     '
     TTXX$ = "CARPETMP"
     On Error Resume Next
     GetAttr (LUDAT$ & TTXX$)
     If Err Then
        MkDir (LUDAT$ & TTXX$)
     End If
     On Error GoTo 0
     RUTATMP$ = LUDAT$ & TTXX$
     '
     Call CARGA_ENCABEZADO(Me.MSF1, "CI/F0;Código/A16;Descripción/A30;Cantidad/F10.1", Me)
     Call CARGA_ENCABEZADO(Me.msf4, "CI/F0;Código/A16;Descripción/A30;Cantidad/F10.1", Me)
     Call CARGA_ENCABEZADO(Me.msf2, "CI/F0;Código/A6;Descripción/A12;Agregada/F5.1", Me)
     Call CARGA_ENCABEZADO(Me.MSF3, "CI/F0;Código/A6;Descripción/A12;Devol./F5.1", Me)
     Call CREATABLA_DETAMASU

End Sub

Private Sub histotalon_Click()
   NORDREP& = XVALO(Label18)
   If NORDREP& < 1 Then Exit Sub
   Call HISTO_CONSU_MAT_SUM(NORDREP&)

End Sub

Private Sub Label18_Change()
    '
   NPX& = Val(Label18)
   Label11 = ""
   Label12 = ""
   Label13 = ""
   Label14 = ""
   Label15 = ""
   Label16 = ""
   Label17 = ""
   Text10 = ""
   List2.Clear
   Label27 = ""
   Label29 = ""
   Text2 = ""
   '
   If NPX& > 0 Then
      CONDI$ = "NumRepa = " & NPX& & ""
      Label11 = VALOBADA("REPARA", "Nume_Cli", CONDI$, "NOMESS")
      Text10 = VALOBADA("REPARA", "NumeroSerie", CONDI$, "NOMESS")
      NC% = VALOBADA("REPARA", "Nume_Cli", CONDI$, "NOMESS")
      FF% = CVINT(VALOBADA("REPARA", "FechRecep", CONDI$, "NOMESS"))
      Label32 = FECHATEX$(FF%)
      '
     If NC% > 0 Then
       If rasocli$(NC%) <> "" Then
          Label12 = rasocli$(NC%)
          Label13 = DOMICLI$(NC%)
          Label14 = CPOSCLI$(NC%)
          Label15 = LOCACLI$(NC%)
          Label16 = TEL1CLI$(NC%)
          Label17 = FAXCLI$(NC%)
       End If
     End If
   End If
   '
   List1.Clear
   NPX& = XVALO(Label18)
   If NPX& < 1 Then Exit Sub
20 SQLX$ = "SELECT * FROM REPARA"
   SQLX$ = SQLX$ + " WHERE NumRepa = " & NPX& & ""
   Dim ANOTRE As New ADODB.Recordset
   Set ANOTRE = New ADODB.Recordset
25 On Error Resume Next
   ANOTRE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With ANOTRE
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
        On Error GoTo 0
        XX$ = SAFETEXT$(!AnotaRep): LONGI& = Len(XX$)
        INICIO& = 1
        Do While INICIO& < LONGI&
            FIN& = InStr(INICIO&, XX$, Chr$(10))
            If FIN& > 0 Then
               TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
               If FIN& < LONGI& Then A = Asc(Mid$(XX$, FIN& + 1, 1))
               List1.AddItem TTXX$
               b = 1: If A = 13 Then b = 2
               INICIO& = FIN& + b
            Else
               TTXX$ = Mid$(XX$, INICIO&)
               List1.AddItem TTXX$
               Exit Do
            End If
        Loop
     End If
     On Error GoTo 0
   End With
   ANOTRE.Close
   Set ANOTRE = Nothing
   '
   Call DECARIMG
   Call CARGAIMG
   Call CARARCH
   Call CARGATEX
   Call CARGAREP
   Call CARGAMATSUM
   '
End Sub
Sub CARGAIMG()
   '
   NPX& = XVALO(Label18)
   IMGTOT% = CantRegistros("GALIMREP", "NUMREPA = '" & UCase(NPX&) & "'")
   If IMGTOT% = 0 Then
     GoTo 99
   End If
   SQLX$ = " SELECT * FROM GALIMREP "
   SQLX$ = SQLX$ + " WHERE NUMREPA = '" & NPX& & "' "
   SQLX$ = SQLX$ + " ORDER BY Norit_Imag ASC "
   Dim IMGTMP As ADODB.Recordset
   Set IMGTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With IMGTMP
     On Error GoTo 0
     'CREO LAS IMAGENES SI FALTAN
     If IMGTOT% > 2 Then
       For JX% = 1 To (IMGTOT% - 2)
         Call Command4_Click
       Next JX%
     End If
     'CARGO LAS IMAGENES
     Do While .EOF = False
        NOIT% = !Norit_Imag
        Text99(NOIT%) = !Descripcion
           FINAM$ = SAFETEXT$(!Nom_Imag)
           XXXX$ = SAFETEXT$(!Imagen)
           Call SaveFile(RUTATMP$ + "\" + FINAM$, XXXX$)
        Label99(NOIT%) = !Nom_Imag
        Label99(NOIT%).Refresh
      .MoveNext
     Loop
     MODIFIC% = 0
   End With
99 End Sub





Private Sub Label27_Change()
   '
   PPXX% = InStr(Label27, ".")
   If PPXX% < 1 Then PPXX% = 1 + Len(Label27)
   Text2 = Left$(Label27, PPXX% - 1)
   '
End Sub

Private Sub Label27_DblClick()
    '
    ARDOC$ = Label27 + "\" + Label1
    Call MUESTRADOC3(ARDOC$)
    '
End Sub

Private Sub Label31_DblClick()
   '
   NPX& = XVALO(Label18)
   '
   'VALIDA SI SE PUEDE MODIFICAR
   If PUEDEACTUREP%(NPX&) = 0 Then
        Call COMUNI("Se Ha Detectado que la O/R Ya Generada")
        'Call COMUNI("Imposible Modificar O/R Ya Generada")
        'MODIFIC% = 0: Exit Sub
   End If
   '
   If NPX& > 0 Then
       ATRI$ = ATRI$ + "/TITUPAN=Detalle de Repuestos "
       ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(140 + Top + Frame1.Top + Picture3.Top))
       ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
       '
10     PRF$ = "REPUREP"
       XXX% = VENTABU%(PRF$ + ATRI$)
       List3.Clear
       '
       For YX& = 1 To ULTREG&("!BOLREDET")
        YY$ = REGLEIDO$("!BOLREDET", YX&)
         Ci% = CVI(Mid$(YY$, 83, 2))
         CAN = CVD(Mid$(YY$, 89, 8))
         If CAN = 0 Then
            Call COMUNI("Falta Ingresar Cantidad en Codigo " & CODEXT(Ci%))
            GoTo 10
         End If
       Next YX&
       MODIFIC% = 1
       '
       Call CARGALISTAREP
   End If
End Sub
Sub CARGALISTAREP()

   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     Ci% = CVI(Mid$(YY$, 83, 2))
     CAN = CVD(Mid$(YY$, 89, 8))
     If CAN <= 9999 Then
          Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 1, 2), 73, 6)
     Else
          Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 73, 6)
     End If
     Call REPLA(TTXX$, CODEXT$(Ci%), 1, 16)
     Call REPLA(TTXX$, DESCRIT0$(Ci%), 14, 52)
     List3.AddItem TTXX$
   Next YX&
End Sub

Private Sub Label99_Change(Index As Integer)
   '
   Static RUTARGRA$, TOPEINI, IZQUINI, ANCHINI, ALTUINI, INTERVER
   '
   Screen.MousePointer = 11
   '
   If INTERVER < 1 Then
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
   If COPLANO$ = "" Then
'       MODIFIC% = 0
       GoTo 98
   End If
   '
   If EXISTE%(COPLANO$) < 1 Then COPLANO$ = RUTATMP$ + "\" + COPLANO$
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

Private Sub List2_DblClick()

    Call Command11_Click

End Sub

Private Sub List3_DblClick()
    '
    Call Label31_DblClick
    '
End Sub

Private Sub matpencli_Click()
10  Call SELECHO("DEUDOR1")
    NC = Val(VALACT1$("DEUDOR1"))
    If NC < 1 Then Exit Sub

    Call MAT_PEN_X_CLIENTE(NC)
    GoTo 10
   
End Sub

Private Sub MnuAddImg_Click()
  '
  If TRIM$(Label18) <> "" Then
    If SSTAB.Tab <> 1 Then SSTAB.Tab = 1
      Call Command4_Click
      On Error Resume Next
        VScroll1.Value = VScroll1.Max
      On Error GoTo 0
  End If
  '
End Sub

Private Sub MnuDelImg_Click()
  '
  'ELIMINAR LA IMAGEN DE LA BASE DE DATOS
  NPX& = XVALO(Label18)
  'VALIDA SI SE PUEDE MODIFICAR
  If PUEDEACTUREP%(NPX&) = 0 Then
    Call COMUNI("Se Ha Detectado que la O/R Ya Generada")
    'Call COMUNI("Imposible Eliminar Imagen O/R Ya Generada")
    'MODIFIC% = 0: Exit Sub
  End If

  CONDI$ = "NumRepa = '" & NPX& & "' " & " AND " & " Norit_Imag = " & IMGACT% & " "
  Call BORRAREGISTROS("GALIMREP", CONDI$, REG&)
  'RENUMERO LAS IMAGENES
  SQLX$ = "SELECT * FROM GALIMREP "
  SQLX$ = SQLX$ + "WHERE NumRepa = " & NPX& & " "
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
   Call CARGAIMG
  '
99 End Sub

Private Sub MNUMAXIMG_CLICK()
  '
  Call Picture99_DblClick(IMGACT%)
  '
End Sub



Private Sub MSF1_DblClick()
   REG& = MSF1.MouseRow
   MSF1.COL = 0
   CANTRG& = MSF1.Rows
   If CANTRG& <= 1 Then Exit Sub
   CANTI_ORIG = XVALO(MSF1.TextMatrix(REG&, 4))
   FRMMODEDI.EDITABLES.Caption = "4" 'COLUMNA EDITABLE
   '
   Campos$ = "CI/F0;Código/A20;Descripción/A36;Cantidad/F15.1"
   OPX% = COMALTER%("Opciones de Trabajo\\Informar Consumo\Informar Devolución")
   If OPX% = 1 Then
     FRMMODEDI.Command1.Caption = "Informar Consumo"
     FRMMODEDI.Caption = "Consumo"
   ElseIf OPX% = 2 Then
     FRMMODEDI.Caption = "Devolución"
     FRMMODEDI.Command1.Caption = "Informar Devolución"
   Else
     Exit Sub
   End If
   
20 Call MODIFICARMSF(MSF1, REG&, FICHREP09, Campos$)
   'validaciones
   If FRMMODEDI.APROBADO = 1 Then
      CANTIMODIFICADA = XVALO(FRMMODEDI.MSF1.TextMatrix(1, 4))
      If CANTI_ORIG <= 0 Then
         Call COMUNI("Imposible Continuar La Cantidad informada Debe Ser Mayor a 0")
          GoTo 20
      End If
      If CANTI_ORIG < CANTIMODIFICADA Then
         Call COMUNI("Imposible Continuar La Cantidad informada Es Superior a La Cantidad Original")
         GoTo 20
      End If
      '
      Resultado = CANTI_ORIG - CANTIMODIFICADA
      If Resultado > 0 Then
         'SI ES UN PARCIAL MODIFICA LA CANTIDAD SIN BORRAR LA FILA
         MSF1.TextMatrix(REG&, 4) = FORMATNUM$(Resultado, "F11.1")
         GoSub APPFILA
      Else
         'SI LA CANTIDAD ES POR EL TOTAL AGREGA A LA GRILLA INFERIRO Y ELIMINA EL REGISTRO
         GoSub APPFILA
         Call cmdEliminaItem(Me.MSF1, REG&)
      End If
      
   End If
   Exit Sub
   
APPFILA:
   Dim MSF As msFlexGrid
   If OPX% = 1 Then 'CONSUMO
      Set MSF = msf2
   Else
      Set MSF = MSF3
   End If
   
   MSF.Rows = MSF.Rows + 1
   MSF.TextMatrix(MSF.Rows - 1, 1) = MSF1.TextMatrix(REG&, 1)
   MSF.TextMatrix(MSF.Rows - 1, 2) = MSF1.TextMatrix(REG&, 2)
   MSF.TextMatrix(MSF.Rows - 1, 3) = MSF1.TextMatrix(REG&, 3)
   MSF.TextMatrix(MSF.Rows - 1, 4) = FORMATNUM$(CANTIMODIFICADA, "F11.1")
   
   Return
   
   '
   
End Sub


Private Sub MSF2_Click()
   REG& = msf2.MouseRow
   If msf2.Rows <= 1 Then Exit Sub
   CIXI% = XVALO(msf2.TextMatrix(REG&, 1))
   CANTIMSF2 = XVALO(msf2.TextMatrix(REG&, 4))
   For i& = 1 To MSF1.Rows - 1
      If CIXI% = XVALO(MSF1.TextMatrix(i&, 1)) Then
         CANTIMSF1 = XVALO(MSF1.TextMatrix(i&, 4))
         MSF1.TextMatrix(i&, 4) = FORMATNUM$(CANTIMSF1 + CANTIMSF2, "F11.1")
         Call cmdEliminaItem(msf2, REG&)
         Exit Sub
      End If
   Next i&
   'SI LLEGA AQUI ES POR QUE NO LO ENCONTRO ENTONCES LO AGREGA
   MSF1.Rows = MSF1.Rows + 1
   MSF1.TextMatrix(MSF1.Rows - 1, 1) = msf2.TextMatrix(REG&, 1)
   MSF1.TextMatrix(MSF1.Rows - 1, 2) = msf2.TextMatrix(REG&, 2)
   MSF1.TextMatrix(MSF1.Rows - 1, 3) = msf2.TextMatrix(REG&, 3)
   MSF1.TextMatrix(MSF1.Rows - 1, 4) = FORMATNUM$(CANTIMSF2, "F11.1")
   Call cmdEliminaItem(msf2, REG&)

End Sub

Private Sub MSF3_DblClick()
   REG& = MSF3.MouseRow
   If MSF3.Rows <= 1 Then Exit Sub
   CIXI% = XVALO(MSF3.TextMatrix(REG&, 1))
   CANTIMSF3 = XVALO(MSF3.TextMatrix(REG&, 4))
   For i& = 1 To MSF1.Rows - 1
      If CIXI% = XVALO(MSF1.TextMatrix(i&, 1)) Then
         CANTIMSF1 = XVALO(MSF1.TextMatrix(i&, 4))
         MSF1.TextMatrix(i&, 4) = FORMATNUM$(CANTIMSF1 + CANTIMSF3, "F11.1")
         Call cmdEliminaItem(MSF3, REG&)
         Exit Sub
      End If
   Next i&
   'SI LLEGA AQUI ES POR QUE NO LO ENCONTRO ENTONCES LO AGREGA
   MSF1.Rows = MSF1.Rows + 1
   MSF1.TextMatrix(MSF1.Rows - 1, 1) = MSF3.TextMatrix(REG&, 1)
   MSF1.TextMatrix(MSF1.Rows - 1, 2) = MSF3.TextMatrix(REG&, 2)
   MSF1.TextMatrix(MSF1.Rows - 1, 3) = MSF3.TextMatrix(REG&, 3)
   MSF1.TextMatrix(MSF1.Rows - 1, 4) = FORMATNUM$(CANTIMSF3, "F11.1")
   Call cmdEliminaItem(MSF3, REG&)

End Sub


Private Sub msf4_DblClick()
   Call SETULTREG("!LIPECOTI", 0)
   
   For i& = 1 To msf4.Rows - 1
      CIXI% = XVALO(msf4.TextMatrix(i&, 1))
      Codigo$ = msf4.TextMatrix(i&, 2)
      DESCRI$ = msf4.TextMatrix(i&, 3)
      CANTI = XVALO(msf4.TextMatrix(i&, 4))
     '
      YY$ = REGBLAN$("LIPECOTI")
      Call REPLA(YY$, MKI$(CIXI%), 1, 2)
      Call REPLA(YY$, MKS$(CANTI), 3, 4)
      Call REPLA(YY$, MKI$(0), 7, 2)
      Call REPLA(YY$, Observa$, 9, 55)
      JJ& = JJ& + 1
      Call GRAREG("!LIPECOTI", YY$, JJ&)
      '
    Next i&
    Call SETULTREG("!LIPECOTI", JJ&)
    Call CIERRARCH("!LIPECOTI")
    '
    PRF$ = "MATESUMI"
    'ATRI$ = "/TITUPAN=MATERIALES SUMINISTRADOS POR EL CLIENTE"
    VTB% = VENTABU%("MATESUMI" + ATRI$)
    If VTB% >= 0 Then
      '
      
      FIN& = ULTREG&("!LIPECOTI")
      For J& = 1 To FIN&
        X$ = REGLEIDO$("!LIPECOTI", J&)
        CODI1% = CVI(Mid$(X$, 1, 2))
        If CODI1% < 1 Then
           Call COMUNI("Código en Fila: " & J& & " Inválido o Inexistente")
           Exit Sub
        End If
        
        CAN1 = CVS(Mid$(X$, 3, 4))
        If CAN1 <= 0 Then
           Call COMUNI("Falta Cantidad en Fila: " & J&)
           Exit Sub
        End If
      Next J&
      '
      FIN& = ULTREG&("!LIPECOTI")
      For J& = 1 To FIN&
        X$ = REGLEIDO$("!LIPECOTI", J&)
        CODI1% = CVI(Mid$(X$, 1, 2))
        CODI2$ = CODEXT$(CODI1%)
        DESC2$ = DESCRIT0(CODI1%)
        CAN1 = CVS(Mid$(X$, 3, 4))
        CAN2$ = FORMATNUM$(CAN1, "F9.1")
        OBS2$ = Mid$(X$, 9, 55)
        
        'AGREGO A LA GRILLA
        msf4.Rows = J& + 1
        msf4.TextMatrix(J&, 1) = CODI1%
        msf4.TextMatrix(J&, 2) = CODI2$
        msf4.TextMatrix(J&, 3) = DESC2$
        msf4.TextMatrix(J&, 4) = CAN2$
      Next J&
      '
    End If

End Sub

Private Sub Picture99_Click(Index As Integer)
   
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

Private Sub Picture99_DblClick(Index As Integer)
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
    '
    Call INVICTRL
    IMGACT% = Index
    PopupMenu Edi
    Call INVICTRL
  End If
End Sub

Sub INVICTRL()
  '
  'DesCam.Visible = Not (DesCam.Visible)
  'phl4.Visible = Not (phl4.Visible)
  'Balog.Visible = Not (Balog.Visible)
  'php16.Visible = Not (php16.Visible)
  MnuAddImg.Visible = Not (MnuAddImg.Visible)
  MnuMaxImg.Visible = Not (MnuMaxImg.Visible)
  MnuDelImg.Visible = Not (MnuDelImg.Visible)
  '
End Sub

Private Sub Salir_Click()
    '
    Call Command2_Click
    '
End Sub



Private Sub Text3_Change()
   MODIFIC% = 1
End Sub

Private Sub Text4_Change()
    MODIFIC% = 1
End Sub

Private Sub Text99_Change(Index As Integer)
    MODIFIC% = 1
End Sub

Private Sub VScroll1_Change()
   On Error Resume Next 'POR ERROR '6' DESBORDAMIENTO
     Picture5.Top = (-10) * VScroll1.Value
   On Error GoTo 0    '

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
  Picture5.Height = Picture4.Height
  VScroll1.Min = 0
  VScroll1.Max = (Picture4.Height - Picture4.Height) / 10 + 1
  VScroll1.Value = 0
  '
End Sub
Sub CARARCH()
   '
   NPX& = XVALO(Label18)
   IMGTOT% = CantRegistros("DOCANRE", "NUMREPA = '" & UCase(NPX&) & "'")
   If IMGTOT% = 0 Then
     GoTo 99
   End If
   SQLX$ = " SELECT * FROM DOCANRE "
   SQLX$ = SQLX$ + " WHERE NUMREPA = '" & UCase(NPX&) & "' "
   Dim Arch As ADODB.Recordset
   Set Arch = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With Arch
     On Error GoTo 0
     'CARGO LAS IMAGENES
     Do While .EOF = False
            Call REPLA(XX$, AJUSTI$(SAFETEXT$(!NomAsignado), 64), 1, 64)
            Call REPLA(XX$, TRIM$(SAFETEXT$(!FILE_NAME)), 65, 128)
            List2.AddItem XX$
     .MoveNext
     Loop
     List2.Refresh
   End With
99 End Sub

Sub GRABATEX()
   '
   NPX& = XVALO(Label18)
   If TRIM$(Label18) <> "" Then
       CONDI$ = "NUMREPA = " & NPX&
       If TRIM$(Text3) <> "" Then
           Call ACTUREGISTRO("REPARA", "INSPECCION", CONDI$, SAFETEXT$(Text3), REG&, "NOMESS")
       End If
       If TRIM$(Text4) <> "" Then
           Call ACTUREGISTRO("REPARA", "REPARACION", CONDI$, SAFETEXT$(Text4), REG&, "NOMESS")
       End If
   End If
   '
End Sub

Sub CARGATEX()
    '
    NPX& = XVALO(Label18)
    If TRIM$(Label18) <> "" Then
       CONDI$ = "NUMREPA = " & NPX&
       Text3 = VALOBADA("REPARA", "INSPECCION", CONDI$, "NOMESS")
       Text4 = VALOBADA("REPARA", "REPARACION", CONDI$, "NOMESS")
   End If
End Sub

Sub GRABAREPUESTOS()
    '
    NPX& = XVALO(Label18)
    If TRIM$(Label18) <> "" Then
       CONDI$ = "NUMREPA = " & NPX& & " AND ORIGEN=0"
       Call ACTUREGISTRO("REPUREPA", "MARBORRA", CONDI$, 1, REG&, "NOMESS")
       '
       SQLX$ = "SELECT * FROM REPUREPA"
       SQLX$ = SQLX$ + " WHERE NumRepa = " & NPX& & ""
       Dim REPUES As ADODB.Recordset
       Set REPUES = New ADODB.Recordset
       REPUES.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       '
       With REPUES
            '
            For YX& = 1 To ULTREG&("!BOLREDET")
                  YY$ = REGLEIDO$("!BOLREDET", YX&)
                  Ci% = CVI(Mid$(YY$, 83, 2))
                  CAN = CVD(Mid$(YY$, 89, 8))
                  '
                  .AddNew   ' YA CONVERTIDO
                  !NUMREPA = NPX&
                  !cod_inte = Ci%
                  !CANT_SOL = CAN
                  !CodiEmpr = CodiEmp%
                  !PRUN_NETO = 0
                  !IMPO_NETO = 0
                  !NUORD_ITEM = 0
                  !IDENLOTE = 0
                  !Origen = 0
                  .Update
            Next YX&
        End With
          REPUES.Close
          Set REPUES = Nothing
          '
          CONDI$ = CONDI$ + " AND MARBORRA = 1 AND ORIGEN=0"
          Call BORRAREGISTROS("REPUREPA", CONDI$, REG&, "NOMESS")
    End If
End Sub

Sub CARGAREP()
    '
    List3.Clear
    Call SETULTREG("!BOLREDET", 0)
    NPX& = XVALO(Label18)
    IMGTOT% = CantRegistros("REPUREPA", "NUMREPA = '" & UCase(NPX&) & "'")
    If IMGTOT% = 0 Then
      GoTo 99
    End If
    SQLX$ = "SELECT * FROM REPUREPA "
    SQLX$ = SQLX$ + " WHERE NUMREPA = '" & UCase(NPX&) & "' AND ORIGEN=0"
    Dim REPU As ADODB.Recordset
    Set REPU = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JJ& = 0
    With REPU
      On Error GoTo 0
      'CARGO LOS REPUESTOS
      Do While .EOF = False
         CAN# = XVALO(!CANT_SOL)
         Ci% = XVALO(!cod_inte)
         XX$ = REGBLAN$("BOLREDET")
         
         Call REPLA(XX$, MKI$(Ci%), 83, 2)
         Call REPLA(XX$, MKD$(CAN#), 89, 8)
         JJ& = JJ& + 1
         Call GRAREG("!BOLREDET", XX$, JJ&)
      .MoveNext
      Loop
      List3.Refresh
    End With
    
    Call CARGALISTAREP
    
99 End Sub

Function PUEDEACTUREP%(NPX&)

    PUEDEACTUREP% = 1
    OREPGEN% = XVALO(VALOBADA("REPARA", "NuOrdRep", "NUMREPA=" & NPX&))
    If OREPGEN% > 0 Then PUEDEACTUREP% = 0
    
End Function






Sub GENERARESERVA_VERSION2(NROREP$)
    '
    Call SETULTREG("!BOLREDET", 0)
    NPX& = XVALO(InputBox$("NUMERO DE TALON", "Ingrese el N° TALON"))
    If NPX& < 1 Then Exit Sub

    NPX& = XVALO(NPX&)
    IMGTOT% = CantRegistros("REPUREPA", "NUMREPA = '" & UCase(NPX&) & "'")
    If IMGTOT% = 0 Then
      GoTo 99
    End If

    SQLX$ = "SELECT * FROM REPUREPA "
    SQLX$ = SQLX$ + " WHERE NUMREPA = '" & UCase(NPX&) & "'"
    Dim REPU As ADODB.Recordset
    Set REPU = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JJ& = 0
    With REPU
      On Error GoTo 0
      'CARGO LOS REPUESTOS
      Do While .EOF = False
         CAN# = XVALO(!CANT_SOL)
         Ci% = XVALO(!cod_inte)
         XX$ = REGBLAN$("BOLREDET")

         Call REPLA(XX$, MKI$(Ci%), 83, 2)
         Call REPLA(XX$, MKD$(CAN#), 89, 8)
         JJ& = JJ& + 1
         Call GRAREG("!BOLREDET", XX$, JJ&)
      .MoveNext
      Loop
      List3.Refresh
    End With

    Call CARGALISTAREPVERSION2




    Call BORRAREGISTROS("RESERVA", "IDSUBOR = '" & NROREP$ & "'", REGAF&, "NOMESS")

    SQLX1$ = "SELECT * FROM RESERVA"
    SQLX1$ = SQLX1$ + " WHERE COD_INTE < 1 "
    Dim RESETMP As ADODB.Recordset
    Set RESETMP = New ADODB.Recordset
    RESETMP.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
    With RESETMP
      KK& = 0
      For UI% = 1 To List3.ListCount
         zz$ = List3.List(UI% - 1)
         CODD$ = TRIM$(Left$(zz$, 16))
         CIIX% = CODINT%(CODD$)
         If CIIX% > 0 Then
           If CIIX% <= NUMAS% Then
             CANTX# = Val(Mid$(zz$, 59, 10))
             .AddNew
             On Error Resume Next
             !CodiEmpr = CodiEmp%
             On Error GoTo 0
             !cod_inte = CIIX%
             !Cod_Exte = CODEXT$(CIIX%)
             !Descrip = DESCRIT$(CIIX%)
             !IdSubOr = NROREP$
             !FechRes = Now
             !CANTRES = CANTX#
             !CantCons = 0
             !Saldo_Entre = CANTX#
             !IDENLOTE = ""
             !DEPOSITO = 0
             !UsoUnit = CANTX#
                KK& = KK& + 1
             !NuOrItem = KK&
             .Update
           End If
         End If
      Next UI%
   End With
   '
   On Error Resume Next
   RESETMP.Close
   Set RESETMP = Nothing
   On Error GoTo 0
   '
99
End Sub
'
Sub CARGALISTAREPVERSION2()
   List1.Clear
   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     Ci% = CVI(Mid$(YY$, 83, 2))
     CAN = CVD(Mid$(YY$, 89, 8))
     Call REPLA(TTXX$, CODEXT$(Ci%), 1, 16)
     Call REPLA(TTXX$, DESCRIT0$(Ci%), 14, 52)
     Call REPLA(TTXX$, MKS$(CAN), 59, 10)
     List1.AddItem TTXX$
   Next YX&
End Sub







