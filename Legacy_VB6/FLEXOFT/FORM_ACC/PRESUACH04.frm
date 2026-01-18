VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PRESUACH04 
   BackColor       =   &H00EEEEDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuestación"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
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
      Height          =   285
      Left            =   3675
      Picture         =   "PRESUACH04.frx":0000
      TabIndex        =   161
      Top             =   210
      Width           =   175
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
      Left            =   1680
      TabIndex        =   160
      Top             =   225
      Width           =   2010
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
      Height          =   285
      Left            =   9555
      TabIndex        =   31
      Top             =   210
      Width           =   1005
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
      Height          =   285
      Left            =   10560
      TabIndex        =   30
      Top             =   210
      Width           =   175
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404000&
      Height          =   8415
      Left            =   11025
      ScaleHeight     =   8355
      ScaleWidth      =   1005
      TabIndex        =   0
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command5 
         Caption         =   "Calc"
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
         Picture         =   "PRESUACH04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   297
         ToolTipText     =   "Recalcula - Actualiza Presupesto"
         Top             =   3990
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
         Height          =   690
         Left            =   105
         Picture         =   "PRESUACH04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Guarda Cambios Realizados"
         Top             =   2955
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
         Picture         =   "PRESUACH04.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Ayuda en Línea"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command37 
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
         Picture         =   "PRESUACH04.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   6195
         Width           =   650
      End
      Begin VB.CommandButton Command38 
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
         Picture         =   "PRESUACH04.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Complementos, Consultas, Listados"
         Top             =   5355
         Width           =   650
      End
      Begin VB.CommandButton Command39 
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
         Picture         =   "PRESUACH04.frx":1374
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprime / Envía por e-mail"
         Top             =   4670
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
         Picture         =   "PRESUACH04.frx":19DE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8730
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
         Picture         =   "PRESUACH04.frx":1CE8
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Generación de Nuevo Presupuesto"
         Top             =   1575
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Copy"
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
         Picture         =   "PRESUACH04.frx":1E32
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Alta de Nuevo Presupuesto por Copia "
         Top             =   2265
         Width           =   650
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
         Picture         =   "PRESUACH04.frx":1F7C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   7035
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   2
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "PRESUACH04.frx":20C6
         Top             =   7350
         Width           =   1515
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6945
      Left            =   150
      TabIndex        =   11
      Top             =   840
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   12250
      _Version        =   327681
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   13693168
      TabCaption(0)   =   "Comercial"
      TabPicture(0)   =   "PRESUACH04.frx":3FE8
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MARCOFICHA(0)"
      Tab(0).Control(1)=   "Label17"
      Tab(0).Control(2)=   "Label8"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Formato"
      TabPicture(1)   =   "PRESUACH04.frx":4004
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Costos"
      TabPicture(2)   =   "PRESUACH04.frx":4020
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame10"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame11"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame12"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame13"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame5"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Res. Productivos"
      TabPicture(3)   =   "PRESUACH04.frx":403C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MARCOFICHA(1)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Res. Economicos"
      TabPicture(4)   =   "PRESUACH04.frx":4058
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "MARCOFICHA(2)"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame5 
         Caption         =   "Otros Costos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   5145
         TabIndex        =   234
         Top             =   5775
         Width           =   5370
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            TabIndex        =   91
            Top             =   460
            Width           =   1215
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3885
            TabIndex        =   92
            Top             =   460
            Width           =   1215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Amortización:"
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
            Index           =   1
            Left            =   105
            TabIndex        =   236
            Top             =   570
            Width           =   1290
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Varios:"
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
            Left            =   2520
            TabIndex        =   235
            Top             =   570
            Width           =   1290
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6450
         Index           =   2
         Left            =   -74895
         TabIndex        =   116
         Top             =   420
         Width           =   10530
         Begin VB.Frame Frame15 
            Caption         =   "Etiquetas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6315
            Left            =   105
            TabIndex        =   117
            Top             =   105
            Width           =   10305
            Begin VB.Label Label101 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   292
               Top             =   5355
               Width           =   1170
            End
            Begin VB.Label Label88 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   291
               Top             =   4200
               Width           =   1170
            End
            Begin VB.Label Label90 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   290
               Top             =   5040
               Width           =   1170
            End
            Begin VB.Label Label87 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   289
               Top             =   3885
               Width           =   1170
            End
            Begin VB.Label Label86 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   288
               Top             =   3570
               Width           =   1170
            End
            Begin VB.Label Label85 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   287
               Top             =   3045
               Width           =   1170
            End
            Begin VB.Label Label84 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   286
               Top             =   2730
               Width           =   1170
            End
            Begin VB.Label Label83 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   285
               Top             =   2415
               Width           =   1170
            End
            Begin VB.Label Label82 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   284
               Top             =   2100
               Width           =   1170
            End
            Begin VB.Label Label81 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   283
               Top             =   1785
               Width           =   1170
            End
            Begin VB.Label Label80 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   282
               Top             =   1470
               Width           =   1170
            End
            Begin VB.Label Label79 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   281
               Top             =   1155
               Width           =   1170
            End
            Begin VB.Label Label78 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   280
               Top             =   630
               Width           =   1170
            End
            Begin VB.Label Label89 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8820
               TabIndex        =   279
               Top             =   4725
               Width           =   1155
            End
            Begin VB.Label Label101 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   278
               Top             =   5355
               Width           =   1170
            End
            Begin VB.Label Label88 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   277
               Top             =   4200
               Width           =   1170
            End
            Begin VB.Label Label90 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   276
               Top             =   5040
               Width           =   1170
            End
            Begin VB.Label Label87 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   275
               Top             =   3885
               Width           =   1170
            End
            Begin VB.Label Label86 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   274
               Top             =   3570
               Width           =   1170
            End
            Begin VB.Label Label85 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   273
               Top             =   3045
               Width           =   1170
            End
            Begin VB.Label Label84 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   272
               Top             =   2730
               Width           =   1170
            End
            Begin VB.Label Label83 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   271
               Top             =   2415
               Width           =   1170
            End
            Begin VB.Label Label82 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   270
               Top             =   2100
               Width           =   1170
            End
            Begin VB.Label Label81 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   269
               Top             =   1785
               Width           =   1170
            End
            Begin VB.Label Label80 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   268
               Top             =   1470
               Width           =   1170
            End
            Begin VB.Label Label79 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   267
               Top             =   1155
               Width           =   1170
            End
            Begin VB.Label Label78 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   266
               Top             =   630
               Width           =   1170
            End
            Begin VB.Label Label89 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6825
               TabIndex        =   265
               Top             =   4725
               Width           =   1155
            End
            Begin VB.Label Label101 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   264
               Top             =   5355
               Width           =   1170
            End
            Begin VB.Label Label88 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   263
               Top             =   4200
               Width           =   1170
            End
            Begin VB.Label Label90 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   262
               Top             =   5040
               Width           =   1170
            End
            Begin VB.Label Label87 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   261
               Top             =   3885
               Width           =   1170
            End
            Begin VB.Label Label86 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   260
               Top             =   3570
               Width           =   1170
            End
            Begin VB.Label Label85 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   259
               Top             =   3045
               Width           =   1170
            End
            Begin VB.Label Label84 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   258
               Top             =   2730
               Width           =   1170
            End
            Begin VB.Label Label83 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   257
               Top             =   2415
               Width           =   1170
            End
            Begin VB.Label Label82 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   256
               Top             =   2100
               Width           =   1170
            End
            Begin VB.Label Label81 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   255
               Top             =   1785
               Width           =   1170
            End
            Begin VB.Label Label80 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   254
               Top             =   1470
               Width           =   1170
            End
            Begin VB.Label Label79 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   253
               Top             =   1155
               Width           =   1170
            End
            Begin VB.Label Label78 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   252
               Top             =   630
               Width           =   1170
            End
            Begin VB.Label Label89 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4830
               TabIndex        =   251
               Top             =   4725
               Width           =   1155
            End
            Begin VB.Label Label102 
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3780
               TabIndex        =   250
               Top             =   315
               Width           =   210
            End
            Begin VB.Label Label101 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   249
               Top             =   5355
               Width           =   1170
            End
            Begin VB.Label Label88 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   248
               Top             =   4200
               Width           =   1170
            End
            Begin VB.Label Label90 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   247
               Top             =   5040
               Width           =   1170
            End
            Begin VB.Label Label87 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   246
               Top             =   3885
               Width           =   1170
            End
            Begin VB.Label Label86 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   245
               Top             =   3570
               Width           =   1170
            End
            Begin VB.Label Label85 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   244
               Top             =   3045
               Width           =   1170
            End
            Begin VB.Label Label84 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   243
               Top             =   2730
               Width           =   1170
            End
            Begin VB.Label Label83 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   242
               Top             =   2415
               Width           =   1170
            End
            Begin VB.Label Label82 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   241
               Top             =   2100
               Width           =   1170
            End
            Begin VB.Label Label81 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   240
               Top             =   1785
               Width           =   1170
            End
            Begin VB.Label Label80 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   239
               Top             =   1470
               Width           =   1170
            End
            Begin VB.Label Label79 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   238
               Top             =   1155
               Width           =   1170
            End
            Begin VB.Label Label78 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   237
               Top             =   630
               Width           =   1170
            End
            Begin VB.Label Label117 
               Alignment       =   1  'Right Justify
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
               Left            =   5745
               TabIndex        =   144
               Top             =   5880
               Width           =   1155
            End
            Begin VB.Label Label169 
               Alignment       =   1  'Right Justify
               Caption         =   "Moneda:"
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
               Left            =   6300
               TabIndex        =   143
               Top             =   5965
               Width           =   2445
            End
            Begin VB.Label Label168 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Pesos"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   330
               Left            =   8715
               TabIndex        =   142
               ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
               Top             =   5820
               Width           =   1275
            End
            Begin VB.Label Label167 
               Alignment       =   1  'Right Justify
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
               Left            =   2835
               TabIndex        =   141
               Top             =   5880
               Width           =   1365
            End
            Begin VB.Label Label166 
               Alignment       =   1  'Right Justify
               Caption         =   "Importe Total por Cantidad:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   555
               TabIndex        =   140
               Top             =   5775
               Width           =   1920
            End
            Begin VB.Line Line7 
               X1              =   2835
               X2              =   9870
               Y1              =   5775
               Y2              =   5775
            End
            Begin VB.Label Label161 
               Alignment       =   1  'Right Justify
               Caption         =   "Precio CON CPI Mínimo:"
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
               Left            =   30
               TabIndex        =   139
               Top             =   5460
               Width           =   2445
            End
            Begin VB.Label Label156 
               Alignment       =   1  'Right Justify
               Caption         =   "Precio SIN  CPI Máximo:"
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
               Left            =   30
               TabIndex        =   138
               Top             =   3675
               Width           =   2445
            End
            Begin VB.Label Label151 
               Alignment       =   1  'Right Justify
               Caption         =   "Costo Total:"
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
               Left            =   30
               TabIndex        =   137
               Top             =   3150
               Width           =   2445
            End
            Begin VB.Label Label146 
               Alignment       =   1  'Right Justify
               Caption         =   "Costo Empaque:"
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
               Left            =   30
               TabIndex        =   136
               Top             =   2835
               Width           =   2445
            End
            Begin VB.Label Label145 
               Caption         =   "Etiqueta"
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
               Left            =   8820
               TabIndex        =   135
               Top             =   375
               Width           =   855
            End
            Begin VB.Label Label144 
               Caption         =   "Etiqueta"
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
               Left            =   6825
               TabIndex        =   134
               Top             =   375
               Width           =   855
            End
            Begin VB.Label Label143 
               Caption         =   "Etiqueta"
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
               Left            =   4830
               TabIndex        =   133
               Top             =   375
               Width           =   855
            End
            Begin VB.Label Label142 
               Caption         =   "Etiqueta"
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
               Left            =   2835
               TabIndex        =   132
               Top             =   375
               Width           =   855
            End
            Begin VB.Label Label141 
               Alignment       =   1  'Right Justify
               Caption         =   "Total Materia Prima:"
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
               Left            =   30
               TabIndex        =   131
               Top             =   1575
               Width           =   2445
            End
            Begin VB.Label Label140 
               Alignment       =   1  'Right Justify
               Caption         =   "Total Tinta:"
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
               Left            =   30
               TabIndex        =   130
               Top             =   1890
               Width           =   2445
            End
            Begin VB.Label Label139 
               Alignment       =   1  'Right Justify
               Caption         =   "Total Barniz:"
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
               Left            =   30
               TabIndex        =   129
               Top             =   2205
               Width           =   2445
            End
            Begin VB.Label Label138 
               Alignment       =   1  'Right Justify
               Caption         =   "Costo Proceso:"
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
               Left            =   30
               TabIndex        =   128
               Top             =   2520
               Width           =   2445
            End
            Begin VB.Label Label136 
               Caption         =   "B"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5820
               TabIndex        =   127
               Top             =   315
               Width           =   210
            End
            Begin VB.Label Label134 
               Caption         =   "C"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7830
               TabIndex        =   126
               Top             =   315
               Width           =   315
            End
            Begin VB.Label Label132 
               Caption         =   "D"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9765
               TabIndex        =   125
               Top             =   315
               Width           =   315
            End
            Begin VB.Label Label131 
               Alignment       =   1  'Right Justify
               Caption         =   "Cantidad:"
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
               Left            =   30
               TabIndex        =   124
               Top             =   735
               Width           =   2445
            End
            Begin VB.Label Label126 
               Alignment       =   1  'Right Justify
               Caption         =   "Mts. Cuadrados:"
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
               Left            =   30
               TabIndex        =   123
               Top             =   1260
               Width           =   2445
            End
            Begin VB.Label Label124 
               Alignment       =   1  'Right Justify
               Caption         =   "Precio SIN  CPI Sugerido:"
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
               Left            =   30
               TabIndex        =   122
               Top             =   3990
               Width           =   2445
            End
            Begin VB.Label Label122 
               Alignment       =   1  'Right Justify
               Caption         =   "Precio SIN  CPI Mínimo:"
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
               Left            =   30
               TabIndex        =   121
               Top             =   4305
               Width           =   2445
            End
            Begin VB.Label Label120 
               Alignment       =   1  'Right Justify
               Caption         =   "Precio CON  CPI Máximo:"
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
               Left            =   30
               TabIndex        =   120
               Top             =   4830
               Width           =   2445
            End
            Begin VB.Label Label118 
               Alignment       =   1  'Right Justify
               Caption         =   "Precio CON CPI Sugerido:"
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
               Left            =   30
               TabIndex        =   119
               Top             =   5145
               Width           =   2445
            End
            Begin VB.Line Line6 
               X1              =   2835
               X2              =   9870
               Y1              =   3450
               Y2              =   3450
            End
            Begin VB.Line Line5 
               X1              =   2835
               X2              =   9870
               Y1              =   4605
               Y2              =   4605
            End
            Begin VB.Line Line4 
               X1              =   2835
               X2              =   9870
               Y1              =   1020
               Y2              =   1020
            End
            Begin VB.Label Label89 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   118
               Top             =   4725
               Width           =   1155
            End
            Begin VB.Label Label119 
               Alignment       =   1  'Right Justify
               Caption         =   "CPI:"
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
               Left            =   2940
               TabIndex        =   145
               Top             =   5985
               Width           =   2445
            End
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6030
         Index           =   1
         Left            =   -74895
         TabIndex        =   95
         Top             =   735
         Width           =   10530
         Begin VB.Frame Frame14 
            Caption         =   "Etiquetas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5685
            Left            =   105
            TabIndex        =   96
            Top             =   210
            Width           =   10305
            Begin VB.OptionButton Option7 
               Caption         =   "Etiqueta"
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
               Left            =   8610
               TabIndex        =   115
               Top             =   690
               Width           =   1065
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Etiqueta"
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
               Left            =   6615
               TabIndex        =   114
               Top             =   690
               Width           =   1065
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Etiqueta"
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
               TabIndex        =   113
               Top             =   690
               Width           =   1065
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Etiqueta"
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
               Left            =   2730
               TabIndex        =   103
               Top             =   690
               Width           =   1065
            End
            Begin VB.Label Label100 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   3
               Left            =   8715
               TabIndex        =   209
               Top             =   5040
               Width           =   1170
            End
            Begin VB.Label Label99 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   3
               Left            =   8715
               TabIndex        =   208
               Top             =   4725
               Width           =   1170
            End
            Begin VB.Label Label98 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   3
               Left            =   8715
               TabIndex        =   207
               Top             =   4095
               Width           =   1170
            End
            Begin VB.Label Label97 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   3
               Left            =   8715
               TabIndex        =   206
               Top             =   3780
               Width           =   1170
            End
            Begin VB.Label Label96 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   3
               Left            =   8715
               TabIndex        =   205
               Top             =   3150
               Width           =   1170
            End
            Begin VB.Label Label95 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   3
               Left            =   8715
               TabIndex        =   204
               Top             =   2835
               Width           =   1170
            End
            Begin VB.Label Label94 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   3
               Left            =   8715
               TabIndex        =   203
               Top             =   2520
               Width           =   1170
            End
            Begin VB.Label Label93 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   3
               Left            =   8715
               TabIndex        =   202
               Top             =   2205
               Width           =   1170
            End
            Begin VB.Label Label92 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   3
               Left            =   8715
               TabIndex        =   201
               Top             =   1785
               Width           =   1170
            End
            Begin VB.Label Label91 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   8715
               TabIndex        =   200
               Top             =   1050
               Width           =   1155
            End
            Begin VB.Label Label100 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   2
               Left            =   6720
               TabIndex        =   199
               Top             =   5040
               Width           =   1170
            End
            Begin VB.Label Label99 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   2
               Left            =   6720
               TabIndex        =   198
               Top             =   4725
               Width           =   1170
            End
            Begin VB.Label Label98 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   2
               Left            =   6720
               TabIndex        =   197
               Top             =   4095
               Width           =   1170
            End
            Begin VB.Label Label97 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   2
               Left            =   6720
               TabIndex        =   196
               Top             =   3780
               Width           =   1170
            End
            Begin VB.Label Label96 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   2
               Left            =   6720
               TabIndex        =   195
               Top             =   3150
               Width           =   1170
            End
            Begin VB.Label Label95 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   2
               Left            =   6720
               TabIndex        =   194
               Top             =   2835
               Width           =   1170
            End
            Begin VB.Label Label94 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   2
               Left            =   6720
               TabIndex        =   193
               Top             =   2520
               Width           =   1170
            End
            Begin VB.Label Label93 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   2
               Left            =   6720
               TabIndex        =   192
               Top             =   2205
               Width           =   1170
            End
            Begin VB.Label Label92 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   2
               Left            =   6720
               TabIndex        =   191
               Top             =   1785
               Width           =   1170
            End
            Begin VB.Label Label91 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   6720
               TabIndex        =   190
               Top             =   1050
               Width           =   1155
            End
            Begin VB.Label Label100 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   1
               Left            =   4725
               TabIndex        =   189
               Top             =   5040
               Width           =   1170
            End
            Begin VB.Label Label99 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   1
               Left            =   4725
               TabIndex        =   188
               Top             =   4725
               Width           =   1170
            End
            Begin VB.Label Label98 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   1
               Left            =   4725
               TabIndex        =   187
               Top             =   4095
               Width           =   1170
            End
            Begin VB.Label Label97 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   1
               Left            =   4725
               TabIndex        =   186
               Top             =   3780
               Width           =   1170
            End
            Begin VB.Label Label96 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   1
               Left            =   4725
               TabIndex        =   185
               Top             =   3150
               Width           =   1170
            End
            Begin VB.Label Label95 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   1
               Left            =   4725
               TabIndex        =   184
               Top             =   2835
               Width           =   1170
            End
            Begin VB.Label Label94 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   1
               Left            =   4725
               TabIndex        =   183
               Top             =   2520
               Width           =   1170
            End
            Begin VB.Label Label93 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   1
               Left            =   4725
               TabIndex        =   182
               Top             =   2205
               Width           =   1170
            End
            Begin VB.Label Label92 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   1
               Left            =   4725
               TabIndex        =   181
               Top             =   1785
               Width           =   1170
            End
            Begin VB.Label Label91 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   4725
               TabIndex        =   180
               Top             =   1050
               Width           =   1155
            End
            Begin VB.Label Label100 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   0
               Left            =   2835
               TabIndex        =   179
               Top             =   5040
               Width           =   1170
            End
            Begin VB.Label Label99 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   0
               Left            =   2835
               TabIndex        =   178
               Top             =   4725
               Width           =   1170
            End
            Begin VB.Label Label98 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   0
               Left            =   2835
               TabIndex        =   177
               Top             =   4095
               Width           =   1170
            End
            Begin VB.Label Label97 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   0
               Left            =   2835
               TabIndex        =   176
               Top             =   3780
               Width           =   1170
            End
            Begin VB.Label Label96 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   0
               Left            =   2835
               TabIndex        =   175
               Top             =   3150
               Width           =   1170
            End
            Begin VB.Label Label95 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   0
               Left            =   2835
               TabIndex        =   174
               Top             =   2835
               Width           =   1170
            End
            Begin VB.Label Label94 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   0
               Left            =   2835
               TabIndex        =   173
               Top             =   2520
               Width           =   1170
            End
            Begin VB.Label Label93 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Index           =   0
               Left            =   2835
               TabIndex        =   172
               Top             =   2205
               Width           =   1170
            End
            Begin VB.Label Label92 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   0
               Left            =   2835
               TabIndex        =   171
               Top             =   1785
               Width           =   1170
            End
            Begin VB.Label Label29 
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3840
               TabIndex        =   98
               Top             =   630
               Width           =   210
            End
            Begin VB.Label Label91 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   2835
               TabIndex        =   97
               Top             =   1050
               Width           =   1155
            End
            Begin VB.Line Line3 
               X1              =   2835
               X2              =   9870
               Y1              =   1575
               Y2              =   1575
            End
            Begin VB.Line Line2 
               X1              =   2835
               X2              =   9870
               Y1              =   4560
               Y2              =   4560
            End
            Begin VB.Line Line1 
               X1              =   2835
               X2              =   9870
               Y1              =   3600
               Y2              =   3600
            End
            Begin VB.Label Label59 
               Alignment       =   1  'Right Justify
               Caption         =   "Horas Máquina:"
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
               Left            =   30
               TabIndex        =   112
               Top             =   5040
               Width           =   2445
            End
            Begin VB.Label Label57 
               Alignment       =   1  'Right Justify
               Caption         =   "Horas Preparación:"
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
               Left            =   30
               TabIndex        =   111
               Top             =   4725
               Width           =   2445
            End
            Begin VB.Label Label55 
               Alignment       =   1  'Right Justify
               Caption         =   "Kilogramos Barniz:"
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
               Left            =   30
               TabIndex        =   110
               Top             =   4200
               Width           =   2445
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               Caption         =   "Kilogramos Tinta:"
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
               Left            =   30
               TabIndex        =   109
               Top             =   3885
               Width           =   2445
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               Caption         =   "Mts. Cuadrados Totales:"
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
               Left            =   30
               TabIndex        =   107
               Top             =   1890
               Width           =   2440
            End
            Begin VB.Label Label42 
               Alignment       =   1  'Right Justify
               Caption         =   "Cantidad:"
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
               Left            =   30
               TabIndex        =   102
               Top             =   1155
               Width           =   2440
            End
            Begin VB.Label Label41 
               Caption         =   "D"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9720
               TabIndex        =   101
               Top             =   630
               Width           =   315
            End
            Begin VB.Label Label39 
               Caption         =   "C"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7725
               TabIndex        =   100
               Top             =   630
               Width           =   315
            End
            Begin VB.Label Label36 
               Caption         =   "B"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5715
               TabIndex        =   99
               Top             =   630
               Width           =   210
            End
            Begin VB.Label Label51 
               Alignment       =   1  'Right Justify
               Caption         =   "Mts. Lineales Totales:"
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
               Left            =   30
               TabIndex        =   108
               Top             =   3255
               Width           =   2445
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               Caption         =   "Mts. Lineales Descarte:"
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
               Left            =   30
               TabIndex        =   106
               Top             =   2940
               Width           =   2445
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
               Caption         =   "Mts. Lineales Producción:"
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
               Left            =   30
               TabIndex        =   105
               Top             =   2625
               Width           =   2445
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   "Mts. Lineales Preparación:"
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
               Left            =   30
               TabIndex        =   104
               Top             =   2310
               Width           =   2445
            End
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Costos de Pre-Impresión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   5145
         TabIndex        =   88
         Top             =   4725
         Width           =   5370
         Begin VB.TextBox Text24 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3885
            TabIndex        =   90
            Top             =   525
            Width           =   1215
         End
         Begin VB.TextBox Text19 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            TabIndex        =   89
            Top             =   525
            Width           =   1215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cortante:"
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
            Index           =   13
            Left            =   2520
            TabIndex        =   94
            Top             =   630
            Width           =   1290
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Planchas y Peliculas:"
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
            Index           =   12
            Left            =   105
            TabIndex        =   93
            Top             =   420
            Width           =   1290
         End
      End
      Begin VB.Frame Frame12 
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
         Height          =   3690
         Left            =   5145
         TabIndex        =   78
         Top             =   945
         Width           =   5370
         Begin VB.VScrollBar VScroll2 
            Height          =   285
            Left            =   5040
            Max             =   9
            Min             =   1
            TabIndex        =   215
            Top             =   525
            Value           =   9
            Width           =   175
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   285
            Left            =   1890
            Max             =   9
            Min             =   1
            TabIndex        =   214
            Top             =   525
            Value           =   9
            Width           =   175
         End
         Begin VB.TextBox Text26 
            Alignment       =   1  'Right Justify
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
            Left            =   1470
            MaxLength       =   1
            TabIndex        =   146
            Text            =   "1"
            Top             =   525
            Width           =   435
         End
         Begin VB.ComboBox Combo8 
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
            Height          =   345
            ItemData        =   "PRESUACH04.frx":4074
            Left            =   1470
            List            =   "PRESUACH04.frx":4081
            Style           =   2  'Dropdown List
            TabIndex        =   87
            Top             =   3150
            Width           =   1800
         End
         Begin VB.ComboBox Combo7 
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
            Height          =   345
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   2625
            Width           =   1800
         End
         Begin VB.TextBox Text17 
            Alignment       =   1  'Right Justify
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
            Left            =   4650
            TabIndex        =   82
            Text            =   "1"
            Top             =   525
            Width           =   435
         End
         Begin VB.ComboBox Combo5 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "PRESUACH04.frx":40B0
            Left            =   1470
            List            =   "PRESUACH04.frx":40B2
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   1575
            Width           =   645
         End
         Begin VB.Label Label77 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Costo/Kg:"
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
            Left            =   3360
            TabIndex        =   233
            Top             =   2730
            Width           =   945
         End
         Begin VB.Label Label76 
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
            Left            =   4335
            TabIndex        =   232
            Top             =   2625
            Width           =   870
         End
         Begin VB.Label Label75 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Costo/ml:"
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
            Left            =   3360
            TabIndex        =   231
            Top             =   3255
            Width           =   945
         End
         Begin VB.Label Label32 
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
            Left            =   4335
            TabIndex        =   230
            Top             =   3150
            Width           =   870
         End
         Begin VB.Label Label74 
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
            Left            =   4350
            TabIndex        =   229
            Top             =   2100
            Width           =   870
         End
         Begin VB.Label Label73 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mts Min:"
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
            Left            =   3675
            TabIndex        =   228
            Top             =   1995
            Width           =   630
         End
         Begin VB.Label Label72 
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
            Left            =   2865
            TabIndex        =   227
            Top             =   2100
            Width           =   870
         End
         Begin VB.Label Label71 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$ Hora:"
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
            Left            =   2205
            TabIndex        =   226
            Top             =   1995
            Width           =   630
         End
         Begin VB.Label Label70 
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
            Left            =   1470
            TabIndex        =   225
            Top             =   2100
            Width           =   765
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hs. Prep:"
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
            Left            =   735
            TabIndex        =   224
            Top             =   1995
            Width           =   630
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00808080&
            X1              =   0
            X2              =   5355
            Y1              =   1365
            Y2              =   1365
         End
         Begin VB.Label Label65 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mts PM:"
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
            Left            =   3675
            TabIndex        =   223
            Top             =   1470
            Width           =   630
         End
         Begin VB.Label Label68 
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
            Left            =   2860
            TabIndex        =   222
            Top             =   1575
            Width           =   870
         End
         Begin VB.Label Label67 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "% Desc:"
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
            Left            =   2205
            TabIndex        =   221
            Top             =   1470
            Width           =   630
         End
         Begin VB.Label Label66 
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
            Left            =   4350
            TabIndex        =   220
            Top             =   1575
            Width           =   870
         End
         Begin VB.Label Label64 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mts.PM:"
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
            Left            =   2100
            TabIndex        =   219
            Top             =   1050
            Width           =   1890
         End
         Begin VB.Label Label63 
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
            Left            =   4140
            TabIndex        =   218
            Top             =   945
            Width           =   1080
         End
         Begin VB.Label Label62 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hs Prep:"
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
            TabIndex        =   217
            Top             =   1050
            Width           =   1890
         End
         Begin VB.Label Label61 
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
            Left            =   1470
            TabIndex        =   216
            Top             =   945
            Width           =   1080
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Empaque:"
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
            TabIndex        =   86
            Top             =   3255
            Width           =   1890
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Acabado Superf.:"
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
            TabIndex        =   84
            Top             =   2520
            Width           =   1470
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad de Cambios:"
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
            Index           =   11
            Left            =   3255
            TabIndex        =   83
            Top             =   420
            Width           =   1290
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria:"
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
            TabIndex        =   80
            Top             =   1650
            Width           =   1890
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "  Cantidad de Colores:"
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
            Left            =   105
            TabIndex        =   79
            Top             =   420
            Width           =   1260
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Tintas"
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
         TabIndex        =   73
         Top             =   5355
         Width           =   4740
         Begin VB.ComboBox Combo3 
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
            Height          =   345
            ItemData        =   "PRESUACH04.frx":40B4
            Left            =   945
            List            =   "PRESUACH04.frx":40C1
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   840
            Width           =   1695
         End
         Begin VB.ComboBox Combo2 
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
            Height          =   345
            ItemData        =   "PRESUACH04.frx":40DC
            Left            =   945
            List            =   "PRESUACH04.frx":40DE
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   420
            Width           =   960
         End
         Begin VB.Label Label60 
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
            Height          =   345
            Left            =   3465
            TabIndex        =   213
            Top             =   840
            Width           =   1080
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$/Kg:"
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
            Left            =   1470
            TabIndex        =   212
            Top             =   945
            Width           =   1890
         End
         Begin VB.Label Label38 
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
            Height          =   345
            Left            =   3465
            TabIndex        =   211
            Top             =   420
            Width           =   1080
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kg/m2:"
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
            Left            =   1470
            TabIndex        =   210
            Top             =   525
            Width           =   1890
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo:"
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
            Left            =   -1050
            TabIndex        =   76
            Top             =   945
            Width           =   1890
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Porcent:"
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
            Left            =   -1050
            TabIndex        =   74
            Top             =   525
            Width           =   1890
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Troquelado"
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
         Left            =   210
         TabIndex        =   69
         Top             =   4200
         Width           =   4740
         Begin VB.OptionButton Option3 
            Caption         =   "Ninguno"
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
            Left            =   3255
            TabIndex        =   72
            Top             =   420
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Now Foil"
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
            Left            =   1680
            TabIndex        =   71
            Top             =   420
            Width           =   1380
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Flexo"
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
            Top             =   420
            Width           =   960
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Formato"
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
         TabIndex        =   61
         Top             =   945
         Width           =   4740
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   525
            Width           =   2745
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "cm 2"
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
            Left            =   3360
            TabIndex        =   170
            Top             =   2730
            Width           =   1890
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   "cm 2"
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
            Left            =   3360
            TabIndex        =   169
            Top             =   2310
            Width           =   1890
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
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
            Left            =   3360
            TabIndex        =   168
            Top             =   1890
            Width           =   1890
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
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
            Left            =   3360
            TabIndex        =   167
            Top             =   1470
            Width           =   1890
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Plancha:"
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
            TabIndex        =   166
            Top             =   2730
            Width           =   1890
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Superficie:"
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
            TabIndex        =   165
            Top             =   2310
            Width           =   1890
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Medida Mayor:"
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
            TabIndex        =   164
            Top             =   1890
            Width           =   1890
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " Medida Menor:"
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
            TabIndex        =   163
            Top             =   1470
            Width           =   1890
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Forma Básica:"
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
            TabIndex        =   162
            Top             =   1050
            Width           =   1890
         End
         Begin VB.Label Label25 
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
            Height          =   285
            Left            =   1680
            TabIndex        =   67
            Top             =   2625
            Width           =   1500
         End
         Begin VB.Label Label24 
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
            Height          =   285
            Left            =   1680
            TabIndex        =   66
            Top             =   2205
            Width           =   1500
         End
         Begin VB.Label Label23 
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
            Height          =   285
            Left            =   1680
            TabIndex        =   65
            Top             =   1785
            Width           =   1500
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "     Tipo de Formato:"
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
            TabIndex        =   64
            Top             =   600
            Width           =   1890
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
            Height          =   285
            Left            =   1680
            TabIndex        =   63
            Top             =   945
            Width           =   2730
         End
         Begin VB.Label Label19 
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
            Height          =   285
            Left            =   1680
            TabIndex        =   62
            Top             =   1365
            Width           =   1500
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cantidad de Etiquetas"
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
         Left            =   -74685
         TabIndex        =   52
         Top             =   5670
         Width           =   10095
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
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
            Index           =   3
            Left            =   8610
            TabIndex        =   59
            Top             =   525
            Width           =   1200
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
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
            Index           =   2
            Left            =   6195
            TabIndex        =   57
            Top             =   525
            Width           =   1200
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
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
            Index           =   1
            Left            =   3675
            TabIndex        =   55
            Top             =   525
            Width           =   1200
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   1155
            TabIndex        =   53
            Top             =   525
            Width           =   1200
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "D:"
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
            Left            =   7875
            TabIndex        =   60
            Top             =   630
            Width           =   630
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Height          =   330
            Left            =   5460
            TabIndex        =   58
            Top             =   630
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "B:"
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
            Left            =   2940
            TabIndex        =   56
            Top             =   630
            Width           =   630
         End
         Begin VB.Label Label4 
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
            Height          =   330
            Left            =   420
            TabIndex        =   54
            Top             =   630
            Width           =   630
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pie de Página"
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
         Left            =   -74685
         TabIndex        =   49
         Top             =   3675
         Width           =   10095
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Left            =   1185
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            Top             =   510
            Width           =   8670
         End
         Begin VB.Label Label10 
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
            Height          =   225
            Index           =   9
            Left            =   90
            TabIndex        =   51
            Top             =   525
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Encabezado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Left            =   -74685
         TabIndex        =   46
         Top             =   735
         Width           =   10095
         Begin VB.TextBox Text11 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1185
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   298
            Top             =   420
            Width           =   8670
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   1185
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   47
            Top             =   1470
            Width           =   8670
         End
         Begin VB.Label Label10 
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
            Height          =   225
            Index           =   5
            Left            =   90
            TabIndex        =   48
            Top             =   1470
            Width           =   1065
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6350
         Index           =   0
         Left            =   -74790
         TabIndex        =   12
         Top             =   525
         Width           =   10320
         Begin VB.Frame Frame7 
            Caption         =   "Destino"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3375
            Left            =   105
            TabIndex        =   21
            Top             =   0
            Width           =   4740
            Begin VB.TextBox Text14 
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
               Height          =   285
               Left            =   1155
               TabIndex        =   300
               Top             =   210
               Visible         =   0   'False
               Width           =   870
            End
            Begin VB.TextBox Text3 
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
               Left            =   1155
               TabIndex        =   159
               Top             =   2835
               Width           =   3200
            End
            Begin VB.TextBox Text30 
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
               Height          =   285
               Left            =   1155
               TabIndex        =   157
               Top             =   1890
               Width           =   3200
            End
            Begin VB.TextBox Text29 
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
               Height          =   285
               Left            =   1155
               TabIndex        =   152
               Top             =   1470
               Width           =   870
            End
            Begin VB.TextBox Text28 
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
               Height          =   285
               Left            =   1155
               TabIndex        =   151
               Top             =   630
               Width           =   3200
            End
            Begin VB.TextBox Text27 
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
               Height          =   285
               Left            =   1155
               TabIndex        =   150
               Top             =   1050
               Width           =   3200
            End
            Begin VB.TextBox Text18 
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
               Height          =   285
               Left            =   2100
               TabIndex        =   149
               Top             =   1470
               Width           =   2240
            End
            Begin VB.TextBox Text15 
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
               Left            =   1155
               TabIndex        =   148
               Top             =   2415
               Width           =   3200
            End
            Begin VB.Label Label105 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "Nro Cliente:"
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
               TabIndex        =   299
               Top             =   315
               Visible         =   0   'False
               Width           =   1110
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
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
               Left            =   210
               TabIndex        =   158
               Top             =   1995
               Width           =   900
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente:"
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
               TabIndex        =   156
               Top             =   735
               Width           =   1110
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
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
               Left            =   210
               TabIndex        =   155
               Top             =   1575
               Width           =   900
            End
            Begin VB.Label Label14 
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
               Left            =   30
               TabIndex        =   154
               Top             =   1155
               Width           =   1110
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
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
               Index           =   10
               Left            =   210
               TabIndex        =   153
               Top             =   2520
               Width           =   900
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "        Referencia:"
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
               Left            =   15
               TabIndex        =   25
               Top             =   2730
               Width           =   1080
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Detalle"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2650
            Left            =   105
            TabIndex        =   39
            Top             =   3465
            Width           =   10095
            Begin VB.TextBox Text10 
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
               Height          =   285
               Left            =   5355
               TabIndex        =   293
               Top             =   525
               Width           =   1425
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
               Height          =   285
               Left            =   1185
               TabIndex        =   43
               Text            =   " "
               Top             =   525
               Width           =   2380
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
               Height          =   285
               Left            =   3570
               TabIndex        =   42
               Top             =   525
               Width           =   175
            End
            Begin VB.TextBox Text20 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1065
               Left            =   1185
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   40
               Top             =   1455
               Width           =   8670
            End
            Begin VB.Label Label104 
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
               Height          =   285
               Left            =   8535
               TabIndex        =   296
               Top             =   525
               Width           =   1290
            End
            Begin VB.Label Label103 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo por Millar:"
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
               Left            =   7560
               TabIndex        =   295
               Top             =   420
               Width           =   945
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Unitario:"
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
               Index           =   2
               Left            =   4305
               TabIndex        =   294
               Top             =   420
               Width           =   975
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Materia Prima:"
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
               Left            =   315
               TabIndex        =   45
               Top             =   420
               Width           =   840
            End
            Begin VB.Label Label11 
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
               Height          =   285
               Left            =   1185
               TabIndex        =   44
               Top             =   945
               Width           =   8655
            End
            Begin VB.Label Label10 
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
               Height          =   225
               Index           =   32
               Left            =   90
               TabIndex        =   41
               Top             =   1470
               Width           =   1065
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Condiciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1905
            Left            =   5040
            TabIndex        =   15
            Top             =   1470
            Width           =   5160
            Begin VB.TextBox Text4 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1155
               MaxLength       =   2
               TabIndex        =   36
               Top             =   840
               Width           =   360
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
               Height          =   285
               Left            =   1515
               TabIndex        =   35
               Top             =   840
               Width           =   175
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
               Height          =   285
               Left            =   1515
               TabIndex        =   18
               Top             =   375
               Width           =   175
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
               Height          =   285
               Left            =   1155
               MaxLength       =   2
               TabIndex        =   17
               Text            =   "1"
               Top             =   375
               Width           =   360
            End
            Begin VB.TextBox Text21 
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
               Left            =   1155
               TabIndex        =   16
               Top             =   1365
               Width           =   3740
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " Condición de Pago:"
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
               Left            =   0
               TabIndex        =   38
               Top             =   735
               Width           =   1080
            End
            Begin VB.Label Label6 
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
               Height          =   285
               Left            =   1785
               TabIndex        =   37
               Top             =   840
               Width           =   3090
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda:"
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
               Left            =   -210
               TabIndex        =   34
               Top             =   420
               Width           =   1290
            End
            Begin VB.Label Label35 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Pesos"
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
               Left            =   1785
               TabIndex        =   20
               Top             =   375
               Width           =   3090
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Plazo de Entrega:"
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
               Index           =   10
               Left            =   -210
               TabIndex        =   19
               Top             =   1260
               Width           =   1290
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Confección"
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
            Left            =   5040
            TabIndex        =   13
            Top             =   0
            Width           =   5160
            Begin VB.TextBox Text8 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1605
               TabIndex        =   27
               Text            =   " "
               Top             =   840
               Width           =   360
            End
            Begin VB.CommandButton Command2 
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
               Left            =   1965
               TabIndex        =   26
               Top             =   840
               Width           =   175
            End
            Begin VB.Label Label13 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1605
               TabIndex        =   147
               Top             =   420
               Width           =   3300
            End
            Begin VB.Label Label5 
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
               Height          =   330
               Left            =   105
               TabIndex        =   29
               Top             =   915
               Width           =   1470
            End
            Begin VB.Label Label3 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2235
               TabIndex        =   28
               Top             =   840
               Width           =   2670
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Presupuestador:"
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
               Left            =   90
               TabIndex        =   14
               Top             =   465
               Width           =   1470
            End
         End
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -70785
         TabIndex        =   23
         Top             =   3960
         Width           =   15
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -71205
         TabIndex        =   22
         Top             =   4905
         Width           =   15
      End
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
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
      Height          =   435
      Index           =   19
      Left            =   8295
      TabIndex        =   32
      Top             =   315
      Width           =   1185
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Presupuesto Nro.:"
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
      Index           =   41
      Left            =   -315
      TabIndex        =   24
      Top             =   285
      Width           =   1950
   End
End
Attribute VB_Name = "PRESUACH04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim NOCALC%

Private Sub Combo1_Click()
   Text26.SetFocus
   Label20 = Mid$(Combo1.TEXT, 53, 24)
   Label19 = FORMATNUM$(Val(Mid$(Combo1.TEXT, 25, 8)), "f8.3")
   Label23 = FORMATNUM$(Val(Mid$(Combo1.TEXT, 33, 8)), "f8.3")
   Label24 = FORMATNUM$(Val(Mid$(Combo1.TEXT, 41, 12)), "f12.5")
   Label25 = FORMATNUM$(Val(Mid$(Combo1.TEXT, 77, 12)), "f12.5")
   '
   Call RECALCU
End Sub

Private Sub Combo2_Click()
   Label38 = FORMATNUM(Val(Mid$(Combo2.TEXT, 9)), "F7.4")
   Combo3.SetFocus
   Call RECALCU
End Sub

Private Sub Combo3_Click()
   Label60 = FORMATNUM(Val(Mid$(Combo3.TEXT, 17)), "F7.2")
   Text26.SetFocus
   Call RECALCU
End Sub

Private Sub Combo5_click()
  '
  Label68 = TRIM$(Mid$(Combo5.TEXT, 13, 8))
  Label66 = TRIM$(Mid$(Combo5.TEXT, 21, 8))
  Label70 = TRIM$(Mid$(Combo5.TEXT, 30, 8))
  Label72 = TRIM$(Mid$(Combo5.TEXT, 39, 8))
  Label74 = TRIM$(Mid$(Combo5.TEXT, 47, 8))
  Call RECALCU
  '
End Sub

Private Sub Combo6_Change()

End Sub

Private Sub Combo7_Click()
   Label76 = TRIM$(FORMATNUM(Val(Mid$(Combo7.TEXT, 17)), "F6.2"))
   Call RECALCU
End Sub

Private Sub Combo8_Click()
   Label32 = TRIM$(FORMATNUM(Val(Mid$(Combo8.TEXT, 17)), "F6.3"))
   Call RECALCU
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("CONPAG")
    VDEVU$ = TRIM$(VALACT1$("CONPAG"))
    If VDEVU$ <> "" Then
       Text4.TEXT = VDEVU$
    End If
    '
End Sub

Private Sub Command2_Click()
    Call SELECHO("VENDEDOR")
    VDEVU$ = TRIM$(VALACT1$("VENDEDOR"))
    If VDEVU$ <> "" Then
       Text8 = VDEVU$
    End If
End Sub

Private Sub Command21_Click()
   If DERACCE%("ALTANUPR", "Generar Nuevo Presupuesto") < 2 Then
      Exit Sub
   End If
   '
   Screen.MousePointer = 1
   '
   'FALTA PONER EL CODIGO PARA QUE TRAIGA LAS FECHAS POR DEFAULT
   '
   HOII% = HOY(HOS$)
   NPX$ = Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 1, 2)
   Call ABREPRECAM
   SQLX$ = "SELECT NRO_PRESU FROM ENCAPRESU "
   SQLX$ = SQLX$ + "WHERE NRO_PRESU LIKE 'PR-" + NPX$ + "-*' "
   SQLX$ = SQLX$ + "ORDER BY NRO_PRESU DESC;"
   Set EncaTemp = PresuCam.OpenRecordset(SQLX$, dbOpenSnapshot)
   NROMA& = 0
   On Error Resume Next
   With EncaTemp
     .MoveFirst
     If Err < 1 Then
       NROMA& = Val(Mid$(!Nro_PRESU, 11))
     End If
     On Error GoTo 0
   End With
   '
   NPRESX$ = "PR-" + NPX$ + "-" + TRIM$(Str$(NROMA& + 1))
   VDEF$ = Space$(64) + String$(6, 0)
   Call REPLA(VDEF$, NPRESX$, 1, 16)
   HOII% = HOY(HOS$)
   Call REPLA(VDEF$, MKI$(HOII%), 67, 2)
   '
9  OBj$ = OBJPLA$("NUEPREC", VDEF$)
   If OBj$ = "" Then GoTo 99
   VDEF$ = OBj$
   '
   NUPRE$ = TRIM$(Mid$(OBj$, 1, 16))
   DESCRIPRE$ = TRIM$(Mid$(OBj$, 17, 48))
   nuclipre% = CVI(Mid$(OBj$, 65, 2))
   RASOCLIPRE$ = RASOCLI$(NUCLIOTR%)
   FECPRE% = CVI(Mid$(OBj$, 67, 2))
   FEVAPRE% = CVI(Mid$(OBj$, 69, 2))
   
   '//////////////VALIDACIONES////////////
   If TRIM$(NUPRE$) = "" Then
      Call COMUNI("Número de Presupuesto no válido.")
      GoTo 9
   End If
   '
   Call ABREPRECAM
   EncaPresu.FindFirst "Nro_Presu = '" & NUPRE$ & "' "
   If EncaPresu.NoMatch = False Then
      Call COMUNI("Numero de Presupuesto no Válido o Pre-Existente")
      GoTo 9
   End If
   '
   If TRIM$(DESCRIPRE$) = "" Then
      Call COMUNI("Debe Ingresar una Descripcion para el Presupuesto")
      GoTo 9
   End If
   '
   If nuclipre% <= 0 Then
      Call COMUNI("Cliente no Válido.")
      GoTo 9
   End If
   '
   If FEVAPRE% < FECPRE% Then
      Call COMUNI("La Fecha de Validez Debe ser Superior o Igual a la Fecha de Generación")
      GoTo 9
   End If
   '
   jjj = CVINT(Date)
   If FECPRE% > jjj Then
      Call COMUNI("La Fecha de Generación no puede ser Superior al dia de la Fecha.")
      GoTo 9
   End If
   '
   Call ABREPRECAM
   With EncaPresu
      .AddNew
        !Nro_PRESU = NUPRE$
        !Refe_Presu = DESCRIPRE$
        !Nclie_Presu = nuclipre%
        !Femi_Presu = CVDAT(FECPRE%)
        !FVali_Presu = CVDAT(FEVAPRE%)
        !uSUCONF_Presu = USERNAME 'usuario de confección
        !UsuCont_Presu = "" 'usuario que controlo
        !UsuApro_Presu = "" 'usuario que aprobara
        !NroOcom_Presu = "" 'numero de orden de compra del cliente
        !FeOcom_Presu = Date 'fecha de pedido del cliente
        !PediCoti_Presu = "" 'Numero de Oferta
        !FePedCot_Presu = Date 'fecha de la oferta
        !Nuoritem_Presu = 0
        !CanPed_Presu = 0
        !Especi_Presu = ""
        !PreVeN_PRESU = 0
        !Mone_Presu = 1     ' MONEDA = PESOS
        !CotiMone_Presu = 0
        !LENTREGA_Presu = " "
        !Ajusta_Presu = 0
        !Vended_Presu = 0
        !ComiVend_Presu = 0
        !Observa_Presu = " "
      .Update
   End With
   '
   Call GENEINDIPRECA
   Text1 = TRIM$(NUPRE$)
   Text14 = TRIM$(Str$(nuclipre%))
   '
99 Screen.MousePointer = 1
   Command21.Enabled = True
   '\\\OT-05-0330-WAR-FIN///
End Sub

Private Sub Command24_Click()
   Call GUARDAPRE
End Sub

Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command37_Click()
  '
  OPREACH04.Show 1
  '
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command39_Click()
   If TRIM$(Text1) = "" Then
      Call COMUNI("Número de Presupuesto no Válido.")
      Exit Sub
   End If
   '
   If ECOARCH$("IPRESCAM", Text1) = "" Then
      Call COMUNI("Número de Presupuesto no Válido.")
      Exit Sub
   End If
   '
   If Text1 <> "" Then
        NCIX% = Val(Text14)
        EMAI$ = TRIM$(EMAILCLI$(NCIX%))
        If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
           GoTo 20 'NO TIENE DIRECCION O DIRECCION NO VALIDA
        End If
        RESP22% = COMALTER%("¿ Formato de Salida ?\\Por E-Mail\Impresora")
        If RESP22% <> 1 Then GoTo 20
        '
15      SENDBYMAIL$ = ""
        pormail% = 1
        SENDBYMAIL$ = EMAI$
        '
        For Each prx1 In Printers
          If InStr(UCase$(prx1.DeviceName), "PDFWRITER") > 0 Then
            For U& = 0 To 256
              MAI_ADJUN$(U&) = ""
            Next U&
            For KU& = 1 To List2.ListCount
              MAI_ADJUN$(KU&) = List2.List(KU& - 1)
            Next KU&
            GoTo 20
          End If
        Next
        Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Impresora Virtual PDFWRITER.")
        SENDBYMAIL = ""
        '
20      'Call PRIDETPRE ' IMPRESION DE DETALLE PRESUPUESTO
        pormail% = 0
        SENDBYMAIL$ = ""
     Else
        Call COMUNI("Debe Seleccionar un Presupuesto de Cámara.")
   End If
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   NOCALC% = 0
   CIXI% = CODINT%(Text12)
   Text10 = FORMATNUM$(COSUNI(CIXI%), "F9.4")
   Call RECALCU
   Command5.Enabled = True
End Sub

Private Sub Command8_Click()
   Call SELECHO("MASTER")
   VDEVU$ = VALACT1$("MASTER")
   If VDEVU$ <> "" Then
     Text12 = VDEVU$
   End If
End Sub

Private Sub Command9_Click()
   '
   Call SELECHO("IPRESCAM")
   NUPRECAM$ = TRIM$(VALACT1$("IPRESCAM"))
   Text1.TEXT = TRIM$(NUPRECAM$)
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
    ARCHEX$(4) = "CONPAG"
    ARCHEX$(5) = "VENDEDOR"
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Combo1.Clear
    Combo1.AddItem "Rect. 35 x 45           35.0    45.0    15.75000    Rectangulo              18.43500    "
    Combo1.AddItem "Rect. 22.3 x 58         22.3    58.0    12.93400    Rectangulo              16.3345     "
    Combo1.AddItem "Redon. diam 48          48.0    48.0    18.09500    Redondo                 25.435      "
    Combo1.AddItem "Ovalado 50 x 30         50.0    30.0    11.77500    Ovalado                 15.2345     "
    '
    Combo2.Clear
    Combo2.AddItem "  10    0.0007"
    Combo2.AddItem "  30    0.0013"
    Combo2.AddItem "  60    0.0020"
    Combo2.AddItem " 100    0.0030"
    '
    Combo3.Clear
    Combo3.AddItem "Agua            14"
    Combo3.AddItem "UV              38"
    Combo3.AddItem "Combinación     29"
    '
    Combo7.Clear                          ' acabado superficial
    Combo7.AddItem "Ninguno         00"
    Combo7.AddItem "Barniz          15"
    '
    Combo8.Clear
    Combo8.AddItem "Rollo Grande    0.015"
    Combo8.AddItem "Rollo Mediano   0.027"
    Combo8.AddItem "Rollo Chico     0.040"
    '
    Combo5.Clear
    '
    L1$ = Space(60)
    Call REPLA(L1$, "A", 1, 2)
    Call REPLA(L1$, "1.05", 13, 8)
    Call REPLA(L1$, "100", 21, 8)
    Call REPLA(L1$, "0.75", 30, 8)
    Call REPLA(L1$, "58", 39, 7)
    Call REPLA(L1$, "35", 47, 5)
    Combo5.AddItem L1$
    '
    L1$ = Space(60)
    Call REPLA(L1$, "B", 1, 2)
    Call REPLA(L1$, "1.10", 13, 8)
    Call REPLA(L1$, "200", 21, 8)
    Call REPLA(L1$, "1.5", 30, 8)
    Call REPLA(L1$, "70", 39, 7)
    Call REPLA(L1$, "35", 47, 5)
    Combo5.AddItem L1$
    '
    L1$ = Space(60)
    Call REPLA(L1$, "C", 1, 2)
    Call REPLA(L1$, "1.105", 13, 8)
    Call REPLA(L1$, "325", 21, 8)
    Call REPLA(L1$, "2", 30, 8)
    Call REPLA(L1$, "103", 39, 7)
    Call REPLA(L1$, "35", 47, 5)
    Combo5.AddItem L1$
    '
    L1$ = Space(60)
    Call REPLA(L1$, "D", 1, 2)
    Call REPLA(L1$, "1.15", 13, 8)
    Call REPLA(L1$, "450", 21, 8)
    Call REPLA(L1$, "2.5", 30, 8)
    Call REPLA(L1$, "121", 39, 7)
    Call REPLA(L1$, "35", 47, 5)
    Combo5.AddItem L1$
    '
    Text11 = "De acuerdo a vuestra solicitud, por medio de la presente tenemos el agrado de presupuestarle lo siguiente:"
    Call Text17_Change
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Sub GENEINDIPRECA() 'GENERA EL INDICE DE ORDENES DE TRABAJO
   'If UPDATOTR% > 0 Then
      Screen.MousePointer = 11
      Call ABREPRECAM
      Set ORDETRATECO = PresuCam.OpenRecordset("SELECT NRO_Presu, Nclie_Presu, REFE_Presu,FEMI_Presu FROM encapresu  ORDER BY REFE_Presu ASC;", dbOpenDynaset, dbReadOnly)
      '
      JJ& = 0
      On Error Resume Next
      ORDETRATECO.MoveLast
      If Err Then Exit Sub
      '
      FIN& = CAREC%
      REBLB$ = REGBLAN$("IPRESCAM")
      With ORDETRATECO
        .MoveFirst
        '
        Do While Not .EOF
          JJ& = JJ& + 1
          Call TRACE(20, JJ&, FIN&)
          NCLIE% = !Nclie_Presu
          EL2$ = REBLB$
          Call REPLA(EL2$, !Nro_PRESU, 1, 16)
            FEX = CVINT(!Femi_Presu)
            REFEINDI$ = FECHATEX$(FEX) + " - " + TRIM$(FANTCLI$(NCLIE%))
            REFEINDI$ = REFEINDI$ + " - " + !Refe_Presu
          Call REPLA(EL2$, REFEINDI$, 17, 64)
          Call GRAREG("IPRESCAM", EL2$, JJ&)
          '
        .MoveNext
        Loop
      End With
      '
90    ORDETRATECO.Close
      Call SETULTREG("IPRESCAM", JJ&)
      Call FILESORT("IPRESCAM", "", 1, 1, "ASC")
      Call FRESHECHO("IPRESCAM")
      Call BAJALDISCO
      UPDATOTR% = 0
      Screen.MousePointer = 1
End Sub

Private Sub Label101_DblClick(Index As Integer)
     ITOTA = XVALO(Label101(Index)) * XVALO(Label78(Index)) / 1000
     Label167 = TRIM$(FORMATNUM$(ITOTA, "F10.2"))
     Label166 = "Importe x " + TRIM$(Label78(Index)) + Mid$(Label161, 7)
End Sub

Private Sub Label19_Change()
   Call RECALCU
End Sub

Private Sub Label2_DblClick()
   Label2 = "C:"
End Sub

Private Sub Label28_DblClick()
   Label28 = "Empaque:"
End Sub

Private Sub Label35_DblClick()
   Text13 = "1"
   Label35 = "Pesos"
End Sub

Private Sub Label63_Change()
   Call RECALCU
End Sub

Private Sub Label86_DblClick(Index As Integer)
     ITOTA = XVALO(Label86(Index)) * XVALO(Label78(Index)) / 1000
     Label167 = TRIM$(FORMATNUM$(ITOTA, "F10.2"))
     Label166 = "Importe x " + TRIM$(Label78(Index)) + Mid$(Label156, 7)
End Sub

Private Sub Label87_DblClick(Index As Integer)
     ITOTA = XVALO(Label87(Index)) * XVALO(Label78(Index)) / 1000
     Label167 = TRIM$(FORMATNUM$(ITOTA, "F10.2"))
     Label166 = "Importe x " + TRIM$(Label78(Index)) + Mid$(Label124, 7)
End Sub

Private Sub Label88_DblClick(Index As Integer)
     ITOTA = XVALO(Label88(Index)) * XVALO(Label78(Index)) / 1000
     Label167 = TRIM$(FORMATNUM$(ITOTA, "F10.2"))
     Label166 = "Importe x " + TRIM$(Label78(Index)) + Mid$(Label122, 7)
End Sub

Private Sub Label89_DblClick(Index As Integer)
     ITOTA = XVALO(Label89(Index)) * XVALO(Label78(Index)) / 1000
     Label167 = TRIM$(FORMATNUM$(ITOTA, "F10.2"))
     Label166 = "Importe x " + TRIM$(Label78(Index)) + Mid$(Label120, 7)
End Sub

Private Sub Label90_DblClick(Index As Integer)
     ITOTA = XVALO(Label90(Index)) * XVALO(Label78(Index)) / 1000
     Label167 = TRIM$(FORMATNUM$(ITOTA, "F10.2"))
     Label166 = "Importe x " + TRIM$(Label78(Index)) + Mid$(Label118, 7)
End Sub

Private Sub Label91_Change(Index As Integer)
     Label78(Index) = Label91(Index)
End Sub

Private Sub Label92_Change(Index As Integer)
   Label79(Index) = Label92(Index)
End Sub

Private Sub Text1_Change()
   Call BLANFORMU
   If TRIM$(Text1.TEXT) <> "" Then
      NUOTR$ = TRIM$(Text1.TEXT)
      If ECOARCH$("IPRESCAM", NUOTR$) <> "" Then
        If NOCALC% < 1 Then
          Call CARGPRE
          Call CARGAPRE
        End If
      End If
   End If
End Sub

Sub BLANFORMU()
   '
   Text8 = ""
   
   
   
   Text2 = ""
   Text3 = ""
   Text24 = ""
   'Text25.TEXT = ""
   'Text36.TEXT = ""
   'Text38.TEXT = ""
   'Text44.TEXT = ""
   'Text16.TEXT = ""
   Text26.TEXT = ""
   Text20.TEXT = ""
   'Text23.TEXT = ""
   Text27.TEXT = ""
   'TICAMBIS.TEXT = ""
   Text4.TEXT = ""
   'Text14.TEXT = ""
   Text19.TEXT = ""
   Text29.TEXT = ""
   Text19.TEXT = ""
   'Text36.TEXT = ""
   Text21.TEXT = ""
   Text23 = ""
   Text36 = ""
   Text37 = ""
   Text3 = ""
   Text28 = ""
   Text27 = ""
   Text29 = ""
   Text18 = ""
   Text16 = ""
   Text10 = ""
   Text25 = ""
   Text24 = ""
   Label13 = ""
   Text13 = ""
   Text12 = ""
   Text17 = ""
   'List1.Clear
   '
   YACAR% = 0
   '
End Sub

Sub CARGPRE() 'CARGA LA FICHA DE UN PRESUPUESTO DE Cámara DETERMINADO
   '
   Static CAMON%, DMONE$(), TCMONE()
   '
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
     If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
          VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
          DMONE$(U&) = DMX$
          TCMONE(U&) = VHX
         Next U&
     End If
   End If
   '
   NUPRECA$ = TRIM$(Text1)
   If NUPRECA$ = "" Or ECOARCH("IPRESCAM", Text1) = "" Then Exit Sub
   '
   Call ABREPRECAM
   SQLX$ = "SELECT * FROM ENCAPRESU"
   Set OTTMP = PresuCam.OpenRecordset(SQLX$, 4)
   With OTTMP
     .FindFirst "  NRO_pRESU= '" & NUPRECA$ & "' "
     If .NoMatch = True Then
        On Error GoTo 0
        'Call COMUNI("No Existe la Orden de Trabajo")
        Exit Sub
     End If
     On Error GoTo 0
     '
     Text3 = !Refe_Presu
     '
     NCLIE% = !Nclie_Presu
     If NCLIE% > 0 Then
       Text28 = RASOCLI$(NCLIE%)
       Text27 = DOMICLI$(NCLIE%)
       Text29 = CPOSCLI$(NCLIE%)
       Text18 = LOCACLI$(NCLIE%)
       Text30 = FAXCLI$(NCLIE%)
       Text15 = CONTACTCLI$(NCLIE%)
     End If
     '
     Text16 = FECHATEX(CVINT(!Femi_Presu))
     COMIVE = Val(!ComiVend_Presu)
     
     
     
     
     Text25 = !NroOcom_Presu
     'Text24 = FECHATEX(CVINT(!FeOcom_Presu))
     Label13 = !uSUCONF_Presu
     Text13 = !PediCoti_Presu
     'Text12 = FECHATEX(CVINT(!FePedCot_Presu))
     Text17 = !Especi_Presu
     Text36 = !CanPed_Presu
     ALTPRE = !Alto_Presu: If ALTPRE < 1 Or ALTPRE > 99 Then ALTPRE = 1
     ANCPRE = !Ancho_Presu
     LARPRE = !Largo_Presu
     ESPRE = !Espesor_Presu
     Text4 = FORMATNUM$(ALTPRE, "I3")
       'VScroll1.Value = 100 - Val(Text4)
     Text14 = FORMATNUM$(ANCPRE, "F5.1")
     Text19 = FORMATNUM$(LARPRE, "F11.2") 'ESTABA CON F6.3
     Text20 = FORMATNUM$(!DesBon_Presu, "F6.1")
     Text22 = FORMATNUM$(ESPRE, "F6.1")
     Text26 = FORMATNUM$(!GGOba_Presu, "F11.2")
     '\\\OT-06-0145-WAR-17/04/06///
     Text31 = FORMATNUM$(!HorasCons_Presu, "F6.1")
     'Label28 = FORMATNUM$(!ImporCons_Presu, "F11.2")
     '\\\OT-06-0145-WAR-FIN///
     Text2 = !LENTREGA_Presu
     Text8 = !Vended_Presu
     Text21 = FORMATNUM(COMIVE, "F6.2")
     If IsNull(!ICONTAB) = True Then
        CUETEXT = ""
      Else
        CUETEXT = CODCUE(!ICONTAB) 'CARGO LA IMPUTACION CONTABLE
     End If
     '
     MONEMI% = !Mone_Presu
     If MONEMI% < 1 Then MONEMI% = 1
     Label168 = DMONE$(MONEMI%)
     Label168.ForeColor = RGB(0, 0, 255)
     If MONEMI% > 1 Then
        Label168.ForeColor = RGB(255, 0, 0)
     End If
     Label168.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
     '
     'Label81 = Label168
     'Label81.ForeColor = Label168.ForeColor
     'Label81.FontUnderline = Label81.FontUnderline
     PVBAS = !PreVeN_PRESU
     Text38 = FORMATNUM$(PVBAS, "F12.2")
     If IsNull(!ForPag_Presu) = False Then Text16 = !ForPag_Presu
     If IsNull(!Observa_Presu) = False Then Text29 = !Observa_Presu
     If IsNull(!Precios_Presu) = False Then Text6 = !Precios_Presu
     If IsNull(!PlaEnt_Presu) = False Then Text15 = !PlaEnt_Presu
     If IsNull(!ValOfer_Presu) = False Then Text32 = !ValOfer_Presu
     '
     'If TRIM$(Text6) = "" Then
     '  TTXX$ = "Los valores expresados son en moneda argentina (pesos) y no incluyen el IVA que deberá adicionarse a los mismos. "
     '  TTXX$ = TTXX$ + CRLF$ + CRLF$ + "Deberan cancelarse unicamente mediante cheque o transferencia bancaria. No se aceptan pagos en efectivo. "
     '  TTXX$ = TTXX$ + CRLF$ + CRLF$ + "El importe correspondiente al IVA, en su totalidad, debera abonarse en el momento de la instalación, con cheque propio al día 8 del mes siguiente, sin excepción."
     '  Text6 = TTXX$
     'End If
     '
   End With
   '
   'Call CARGOPLACO
   'Call SUMAHS
   MODIFIC% = 0
   YACAR% = 1
   '
End Sub
'

Private Sub Text10_Change()
   Call RECALCU
End Sub

Private Sub Text12_Change()
   Label11 = ""
   Text20 = ""
   CIXI% = CODINT%(Text12)
   If CIXI% > 0 Then
     Label11 = DESCRIT0$(CIXI%)
     Text20 = ESPECIFI$(CIXI%)
     If NOCALC% < 1 Then
       Text10 = FORMATNUM$(COSUNI(CIXI%), "F9.4")
     End If
   End If
End Sub

Private Sub Text12_DblClick()
   Call Command8_Click
End Sub

Private Sub Text13_DblClick()
   Text13 = "1"
   Label35 = "Pesos"
End Sub

Private Sub Text13_GotFocus()
   Text4.SetFocus
End Sub

Private Sub Text17_Change()
   If NOCALC% < 1 Then
     If Val(Text17) < 1 Then Text17 = 1
     If Val(Text17) > 9 Then Text17 = 9
     VScroll2.Value = 10 - Val(Text17)
     '
     CACAMB% = Val(Text17)
     CACOLO% = Val(Text26)
     If CACOLO% = 1 Then
         HOPREP = 0.33 * CACAMB%
         METPM = 75 * CACAMB%
       ElseIf CACOLO% <= 3 Then
         HOPREP = 0.44 * CACAMB%
         METPM = 200 * CACAMB%
       Else
         HOPREP = 1 * CACAMB%
         METPM = 325 * CACAMB%
     End If
     '
     Label61 = FORMATNUM(HOPREP, "F6.1")
     Label63 = FORMATNUM(METPM, "F7.1")
   End If
   '
End Sub

Private Sub Text19_Change()
   Call RECALCU
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call Text19_LostFocus
   End If
End Sub

Private Sub Text19_LostFocus()
   Text19 = FORMATNUM(Val(Text19), "f9.2")
End Sub

Private Sub Text24_Change()
   Call RECALCU
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call Text24_LostFocus
   End If
End Sub

Private Sub Text24_LostFocus()
   Text24 = FORMATNUM(Val(Text24), "f9.2")
End Sub

Private Sub Text26_Change()
   If NOCALC% < 1 Then
     If Val(Text26) < 1 Then Text26 = 1
     If Val(Text26) > 9 Then Text26 = 9
     VScroll1.Value = 10 - Val(Text26)
     Call Text17_Change
   End If
End Sub

Private Sub Text4_Change()
    Label6 = ECOARCH$("CONPAG", MKI$(XVALO(Text4.TEXT)))
End Sub

Private Sub Text4_DblClick()
   Call Command1_Click
End Sub

Private Sub Text6_Change()
   Call RECALCU
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call Text6_LostFocus
   End If
End Sub

Private Sub Text6_LostFocus()
   Text6 = FORMATNUM(Val(Text6), "f9.2")
End Sub

Private Sub Text7_Change()
   Call RECALCU
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call Text7_LostFocus
   End If
End Sub

Private Sub Text7_LostFocus()
   Text7 = FORMATNUM(Val(Text7), "f9.2")
End Sub

Private Sub Text8_Change()
   If NOCALC% < 1 Then
     Label3 = ECOARCH$("VENDEDOR", Chr$(Val(Text8)))
   End If
End Sub

Private Sub Text8_DblClick()
   Call Command2_Click
End Sub

Private Sub Text9_Change(Index As Integer)
   If NOCALC% < 1 Then
     Label91(Index) = FORMATNUM$(Val(Text9(Index)), "f8.0")
     Call RECALCU
   End If
End Sub

Private Sub Text9_LostFocus(Index As Integer)
   Text9(Index) = FORMATNUM$(Val(Text9(Index)), "f8.0")
End Sub

Sub RECALCU()
   '
   If NOCALC% < 1 Then
     SUPLANCHA = Val(Label25) / 10000
     MEDIMEN = Val(Label19)              ' MEDIDA MENOR EN MILIMETROS
     PORDESCARTE = Val(Label68)          ' COEFICIENTE DE DESCARTES
     VELOCI = Val(Label74)               ' VELOCIDAT MTS / MINUTO
       If VELOCI < 1 Then VELOCI = 1     ' PARA CAPTURA DE ERROR
     KILMETRO = Val(Label38)             ' KILOS POR METRO CUADRADO
     '
     COSPAPMIL = 1000 * SUPLANCHA * Val(Text10)
     Label104 = FORMATNUM(COSPAPMIL, "F9.2")
     '
     For KKI% = 0 To 3
       CANETIQ = Val(Text9(KKI%))
       '
       If CANETIQ > 0 Then
         '
         MECUADRA = CANETIQ * SUPLANCHA ' METROS CUADRADOS TOTALES
         Label92(KKI%).Caption = FORMATNUM(MECUADRA, "F8.2")
         '
         MEPREPA = Val(Label63) + Val(Label66)       ' METROS DE PREPARACIÓN
         Label93(KKI%).Caption = FORMATNUM(MEPREPA, "F8.2")
         '
         MELIPRO = PORDESCARTE * CANETIQ * (MEDIMEN + 3) / 1000 ' METROS LINEALES DE PRODUCCION - 3 ES POR LOS MILIMETROS DE SEPARACION ENTRE UNA Y OTRA ETIQUETA
         Label94(KKI%).Caption = FORMATNUM(MELIPRO, "F8.2")
         COSEMPAQ = 15 + MELIPRO * XVALO(Label32)
         Label84(KKI%).Caption = FORMATNUM(COSEMPAQ, "F8.2")
         '
         MEDESCAR = (PORDESCARTE - 1) * MELIPRO ' METROS LINEALES DE DESCARTE
         Label95(KKI%).Caption = FORMATNUM(MEDESCAR, "F8.2")
         '
         METOTA = MEPREPA + MELIPRO + MEDESCAR
         Label96(KKI%).Caption = FORMATNUM(METOTA, "F8.2")
         '
         KILOTINT = MECUADRA * KILMETRO
         Label97(KKI%).Caption = FORMATNUM(KILOTINT, "F8.2")
         COSTINT = KILOTINT * XVALO(Label60)
         Label81(KKI%).Caption = FORMATNUM(COSTINT, "F8.2")
         '
         KILOBARN = MECUADRA * 0.015
         ' 0.015 son los kilos de barniz por metro cuadrado
         Label98(KKI%).Caption = FORMATNUM(KILOBARN, "F8.2")
         COSBARN = KILOBARN * XVALO(Label76)
         Label82(KKI%).Caption = FORMATNUM(COSBARN, "F8.2")
         '
         HORASPREP = Val(Label61) + Val(Label70)
         Label99(KKI%).Caption = FORMATNUM(HORASPREP, "F8.2")
         '
         HORASMAQ = (MEPREPA + MELIPRO + MEDESCAR) / VELOCI / 60
         Label100(KKI%).Caption = FORMATNUM(HORASMAQ, "F8.2")
         '
         Label83(KKI%) = FORMATNUM((HORASPREP + HORASMAQ) * Val(Label72), "F8.2")
         Label80(KKI%) = FORMATNUM(COSPAPMIL * CANETIQ / 1000, "F9.2")
         '
         SUMAPAR = Val(Label79(KKI%)) + Val(Label80(KKI%))
         SUMAPAR = SUMAPAR + Val(Label81(KKI%)) + Val(Label82(KKI%))
         SUMAPAR = SUMAPAR + Val(Label83(KKI%)) + Val(Label84(KKI%))
         Label85(KKI%) = FORMATNUM(SUMAPAR, "F9.2")
         '
         PREMAX = 1.22 * 1000 * Val(Label85(KKI%)) / Val(Label78(KKI%))
         Label86(KKI%) = FORMATNUM(PREMAX, "F9.2")
         PRESUG = 1.2 * 1000 * Val(Label85(KKI%)) / Val(Label78(KKI%))
         Label87(KKI%) = FORMATNUM(PRESUG, "F9.2")
         PREMIN = 1.15 * 1000 * Val(Label85(KKI%)) / Val(Label78(KKI%))
         Label88(KKI%) = FORMATNUM(PREMIN, "F9.2")
         '
         SUMCPI = XVALO(Text19) + XVALO(Text7) + XVALO(Text24) + XVALO(Text6)
         Label117 = FORMATNUM(SUMCPI, "F9.2")
         PREMAX = 1.22 * 1000 * (Val(Label85(KKI%)) + SUMCPI) / Val(Label78(KKI%))
         Label89(KKI%) = FORMATNUM(PREMAX, "F9.2")
         PRESUG = 1.2 * 1000 * (Val(Label85(KKI%)) + SUMCPI) / Val(Label78(KKI%))
         Label90(KKI%) = FORMATNUM(PRESUG, "F9.2")
         PREMIN = 1.15 * 1000 * (Val(Label85(KKI%)) + SUMCPI) / Val(Label78(KKI%))
         Label101(KKI%) = FORMATNUM(PREMIN, "F9.2")
         '
       End If
     Next KKI%
   End If
End Sub

Private Sub VScroll1_Change()
   If NOCALC% < 1 Then
     Text26 = 10 - VScroll1.Value
   End If
End Sub

Private Sub VScroll2_Change()
   If NOCALC% < 1 Then
     Text17 = 10 - VScroll2.Value
   End If
End Sub

Sub GUARDAPRE()
   FIN& = Me.Controls.Count
   If TRIM$(Text1) <> "" Then
     Screen.MousePointer = 11
     N1% = FILEOPEN(LUDAT$ + Text1 + ".PRE", 2, 128)
     For U& = 1 To FIN&
       AB1$ = "": AB2$ = "": AB3$ = "": AB4$ = ""
       On Error Resume Next
       AB1$ = Controls(U& - 1).Caption
       AB2$ = Controls(U& - 1).TEXT
       AB3$ = Controls(U& - 1).Name
       AB4$ = TRIM$(Str$(Controls(U& - 1).Index))
       On Error GoTo 0
       '
       If AB3$ <> "" Then
         If InStr(UCase$(AB3$), "LABEL") > 0 Or InStr(UCase$(AB3$), "TEXT") > 0 Or InStr(UCase$(AB3$), "COMBO") > 0 Then
           TTXX$ = AJUSTI$(AB3$ + Chr$(124) + AB4$, 128) + AJUSTI$(AB1$, 128) + AB2$
           Print #N1%, TTXX$
         End If
       End If
     Next U&
     Close #N1%
   End If
   Screen.MousePointer = 1
   '
End Sub

Sub CARGAPRE()
   If TRIM$(Text1) <> "" Then
     If EXISTE%(LUDAT$ + Text1 + ".PRE") < 1 Then Exit Sub
     '
     On Error Resume Next
     N1% = FILEOPEN(LUDAT$ + Text1 + ".PRE", 1, 128)
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible abrir Presupuesto Guardado")
       Exit Sub
     End If
     On Error GoTo 0
     Screen.MousePointer = 11
     NOCALC% = 1
     '
     FIN& = Me.Controls.Count
     Do While Not EOF(N1%)
       Line Input #N1%, TTXX$
       AB4$ = ""
       AB3$ = TRIM$(Left$(TTXX$, 128))
         PPXX% = InStr(AB3$, Chr$(124))
         If PPXX% > 0 Then
           AB4$ = TRIM$(Mid$(AB3$, PPXX% + 1))
           AB3$ = TRIM$(Left$(AB3$, PPXX% - 1))
         End If
       AB1$ = RTrim(Mid$(TTXX$, 129, 128))
       AB2$ = Mid$(TTXX$, 257)
       '
       For U& = 1 To FIN&
         If Controls(U& - 1).Name = AB3$ Then
           AKXX$ = ""
           On Error Resume Next
           AKXX$ = TRIM$(Str$(Controls(U& - 1).Index))
           On Error GoTo 0
           If AKXX$ = AB4$ Then
             On Error Resume Next
             Controls(U& - 1).Caption = AB1$
             Controls(U& - 1).TEXT = AB2$
             On Error GoTo 0
             Exit For
           End If
         End If
       Next U&
     Loop
     Close #N1%
     NOCALC% = 0
     Screen.MousePointer = 1
   End If
   '
End Sub
