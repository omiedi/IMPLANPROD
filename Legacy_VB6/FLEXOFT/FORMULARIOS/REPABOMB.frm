VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form REPABOMB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reparaciones - Especificación Reparaciones Standard"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "REPABOMB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0FFE0&
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
      Height          =   1885
      Left            =   10740
      TabIndex        =   88
      Top             =   4260
      Width           =   1040
      Begin VB.CommandButton Command10 
         Caption         =   "Presup"
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
         Left            =   125
         Picture         =   "REPABOMB.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Generar Presupuesto por la Reparación Activa"
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O/Rep"
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
         Left            =   105
         Picture         =   "REPABOMB.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Generar Ordenes de Reparación"
         Top             =   1050
         Width           =   765
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0FFE0&
      Caption         =   "Reparación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   210
      TabIndex        =   79
      Top             =   210
      Width           =   8310
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
         Height          =   315
         Left            =   3045
         Picture         =   "REPABOMB.frx":075E
         TabIndex        =   82
         Top             =   700
         Width           =   175
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
         Height          =   315
         Left            =   1050
         MaxLength       =   15
         TabIndex        =   81
         Top             =   700
         Width           =   2040
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
         Height          =   645
         Left            =   3360
         MultiLine       =   -1  'True
         TabIndex        =   80
         Top             =   390
         Width           =   4785
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         TabIndex        =   83
         Top             =   380
         Width           =   1005
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0FFE0&
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
      Height          =   1885
      Left            =   10740
      TabIndex        =   14
      Top             =   6210
      Width           =   1040
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
         Left            =   125
         Picture         =   "REPABOMB.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Imprime Ficha de Reparación"
         Top             =   1050
         Width           =   765
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
         Left            =   125
         Picture         =   "REPABOMB.frx":10D2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Configuración - Formularios de Impresión"
         Top             =   315
         Width           =   765
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   8715
      ScaleHeight     =   1050
      ScaleWidth      =   60
      TabIndex        =   12
      Top             =   315
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5000
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0FFE0&
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
      Height          =   2620
      Left            =   10740
      TabIndex        =   4
      Top             =   1575
      Width           =   1040
      Begin VB.CommandButton Command22 
         Caption         =   "Copy"
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
         Left            =   125
         Picture         =   "REPABOMB.frx":1514
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Alta de Nueva Reparación por Copia Completa de Datos de Reparación Existente"
         Top             =   1050
         Width           =   765
      End
      Begin VB.CommandButton Command21 
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
         Height          =   690
         Left            =   125
         Picture         =   "REPABOMB.frx":165E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Alta de Nuevo Código de Reparación"
         Top             =   315
         Width           =   750
      End
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
         Height          =   690
         Left            =   125
         Picture         =   "REPABOMB.frx":17A8
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Baja Lógica Individual"
         Top             =   1785
         Width           =   765
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0FFE0&
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   9030
      TabIndex        =   1
      Top             =   210
      Width           =   2745
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
         Height          =   750
         Index           =   4
         Left            =   1830
         Picture         =   "REPABOMB.frx":1AB2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   315
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
         Height          =   750
         Left            =   1080
         Picture         =   "REPABOMB.frx":1BFC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ayuda en Línea"
         Top             =   315
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
         Height          =   750
         Left            =   210
         Picture         =   "REPABOMB.frx":1F06
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   315
         Width           =   750
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6420
      Left            =   150
      TabIndex        =   0
      Top             =   1680
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   11324
      _Version        =   327681
      TabHeight       =   520
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Descripción General"
      TabPicture(0)   =   "REPABOMB.frx":2210
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "MARCOFICHA(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Materiales / Repuestos"
      TabPicture(1)   =   "REPABOMB.frx":222C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MARCOFICHA(2)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tareas / Operaciones"
      TabPicture(2)   =   "REPABOMB.frx":2248
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MARCOFICHA(3)"
      Tab(2).ControlCount=   1
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Index           =   3
         Left            =   -74895
         TabIndex        =   38
         Top             =   420
         Width           =   10095
         Begin VB.Frame Frame4 
            Caption         =   "Detalle de la Operación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5580
            Left            =   4830
            TabIndex        =   45
            Top             =   210
            Width           =   5160
            Begin VB.Frame Frame10 
               BackColor       =   &H00C0C0C0&
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
               Height          =   2280
               Left            =   210
               TabIndex        =   84
               Top             =   3150
               Width           =   4770
               Begin VB.TextBox Text2 
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
                  Left            =   105
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   85
                  Top             =   315
                  Width           =   4530
               End
               Begin VB.Line Line5 
                  X1              =   105
                  X2              =   4640
                  Y1              =   2110
                  Y2              =   2110
               End
               Begin VB.Line Line4 
                  X1              =   4630
                  X2              =   4630
                  Y1              =   315
                  Y2              =   2100
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
               Height          =   330
               Index           =   0
               Left            =   2100
               TabIndex        =   54
               Top             =   1785
               Width           =   175
            End
            Begin VB.TextBox Text5 
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
               Left            =   1260
               TabIndex        =   53
               Top             =   885
               Width           =   3720
            End
            Begin VB.TextBox Text9 
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
               Index           =   0
               Left            =   1260
               MaxLength       =   6
               TabIndex        =   52
               Text            =   " "
               Top             =   1785
               Width           =   885
            End
            Begin VB.TextBox Text15 
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
               Left            =   3675
               TabIndex        =   51
               Top             =   525
               Width           =   465
            End
            Begin VB.TextBox Text17 
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
               Left            =   4200
               TabIndex        =   50
               Top             =   525
               Width           =   360
            End
            Begin VB.TextBox Text26 
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
               Left            =   4620
               TabIndex        =   49
               Top             =   525
               Width           =   360
            End
            Begin VB.TextBox Text30 
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
               Left            =   1260
               TabIndex        =   48
               Text            =   " "
               Top             =   2625
               Width           =   570
            End
            Begin VB.TextBox Text33 
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
               Left            =   2415
               TabIndex        =   47
               Text            =   " "
               Top             =   2625
               Width           =   360
            End
            Begin VB.CommandButton Command29 
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
               Left            =   2775
               TabIndex        =   46
               Top             =   2625
               Width           =   175
            End
            Begin VB.Label Label42 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Operación:"
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
               Left            =   -735
               TabIndex        =   69
               Top             =   630
               Width           =   1905
            End
            Begin VB.Label Label43 
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
               Left            =   1260
               TabIndex        =   68
               Top             =   2160
               Width           =   3720
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Equipo Principal:"
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
               Left            =   210
               TabIndex        =   67
               Top             =   1680
               Width           =   960
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tiempo:"
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
               Left            =   1680
               TabIndex        =   66
               Top             =   630
               Width           =   1905
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   4095
               TabIndex        =   65
               Top             =   525
               Width           =   120
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   4515
               TabIndex        =   64
               Top             =   525
               Width           =   120
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Operarios:"
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
               Left            =   -735
               TabIndex        =   63
               Top             =   2730
               Width           =   1905
            End
            Begin VB.Label Label51 
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
               Left            =   3045
               TabIndex        =   62
               Top             =   2625
               Width           =   1935
            End
            Begin VB.Label Label52 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cat:"
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
               Left            =   1785
               TabIndex        =   61
               Top             =   2730
               Width           =   540
            End
            Begin VB.Label Label54 
               BackStyle       =   0  'Transparent
               Caption         =   "    H.    M.    S."
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
               Left            =   3675
               TabIndex        =   60
               Top             =   300
               Width           =   1380
            End
            Begin VB.Label Label55 
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
               Left            =   1260
               TabIndex        =   59
               Top             =   525
               Width           =   570
            End
            Begin VB.Label Label60 
               Alignment       =   2  'Center
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
               Left            =   1260
               TabIndex        =   58
               Top             =   1260
               Width           =   885
            End
            Begin VB.Label Label61 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Revisión:"
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
               Left            =   -735
               TabIndex        =   57
               Top             =   1350
               Width           =   1905
            End
            Begin VB.Label Label62 
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
               Left            =   3675
               TabIndex        =   56
               Top             =   1260
               Width           =   1305
            End
            Begin VB.Label Label63 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Responsable:"
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
               TabIndex        =   55
               Top             =   1350
               Width           =   1905
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Operaciones Standard"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5580
            Left            =   105
            TabIndex        =   39
            Top             =   210
            Width           =   4530
            Begin VB.Frame Frame8 
               BorderStyle     =   0  'None
               Caption         =   "Frame8"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   80
               Left            =   630
               TabIndex        =   87
               Top             =   5450
               Width           =   3690
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
               Height          =   4560
               Left            =   90
               TabIndex        =   44
               Top             =   450
               Width           =   4320
            End
            Begin VB.CommandButton Command30 
               BackColor       =   &H00E0E0E0&
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
               Height          =   430
               Left            =   3465
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Copia la Secuencia de Operaciones de una Reparación a Otra"
               Top             =   5010
               Width           =   945
            End
            Begin VB.CommandButton Command32 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Del"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   430
               Left            =   1395
               Style           =   1  'Graphical
               TabIndex        =   42
               ToolTipText     =   "Elimina la Operación Activa"
               Top             =   5010
               Width           =   630
            End
            Begin VB.CommandButton Command33 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Add"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   430
               Left            =   105
               Style           =   1  'Graphical
               TabIndex        =   41
               ToolTipText     =   "Agrega una Operación al Final de la Lista"
               Top             =   5010
               Width           =   630
            End
            Begin VB.CommandButton Command43 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ins"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   430
               Left            =   750
               Style           =   1  'Graphical
               TabIndex        =   40
               ToolTipText     =   "Intercala una Operación Antes de la Activa"
               Top             =   5010
               Width           =   630
            End
            Begin VB.Line Line6 
               X1              =   2000
               X2              =   3465
               Y1              =   5430
               Y2              =   5430
            End
            Begin VB.Image Image2 
               Height          =   510
               Left            =   1785
               Picture         =   "REPABOMB.frx":2264
               Stretch         =   -1  'True
               Top             =   4980
               Width           =   2010
            End
            Begin VB.Line Line3 
               X1              =   4410
               X2              =   4410
               Y1              =   450
               Y2              =   5430
            End
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Index           =   2
         Left            =   -74895
         TabIndex        =   18
         Top             =   420
         Width           =   10200
         Begin VB.Frame Frame14 
            Caption         =   "Utilidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   8505
            TabIndex        =   19
            Top             =   3990
            Width           =   1590
            Begin VB.CommandButton Command40 
               Caption         =   "Incorporar"
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
               Left            =   210
               TabIndex        =   91
               ToolTipText     =   "Incorpora Lista de Repuestos a Partir de Fórmula de Conjunto"
               Top             =   800
               Width           =   1170
            End
            Begin VB.CommandButton Command40 
               Caption         =   "Copiar"
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
               Index           =   5
               Left            =   210
               TabIndex        =   20
               ToolTipText     =   "Nueva Lista de Repuestos por Copia de una Pre-Existente"
               Top             =   420
               Width           =   1170
            End
         End
         Begin VB.Frame Frame12 
            BorderStyle     =   0  'None
            Caption         =   "Frame11"
            Height          =   855
            Left            =   10080
            TabIndex        =   93
            Top             =   5040
            Width           =   225
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Actualizar Materiales / Repuestos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1740
            Left            =   210
            TabIndex        =   24
            Top             =   3990
            Width           =   8100
            Begin VB.CommandButton Command27 
               Caption         =   "Append"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   275
               Left            =   7035
               Picture         =   "REPABOMB.frx":584E
               TabIndex        =   31
               ToolTipText     =   "Agrega el Elemento al Final de la Lista"
               Top             =   410
               Width           =   855
            End
            Begin VB.TextBox Text32 
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
               Left            =   5670
               TabIndex        =   30
               Text            =   " "
               Top             =   420
               Width           =   1170
            End
            Begin VB.TextBox Text31 
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
               Left            =   1260
               TabIndex        =   29
               Top             =   420
               Width           =   1800
            End
            Begin VB.CommandButton Command26 
               Caption         =   "Insert"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   275
               Left            =   7035
               Picture         =   "REPABOMB.frx":5B58
               TabIndex        =   28
               ToolTipText     =   "Intercala el Elemento en la Lista Antes del Activo"
               Top             =   695
               Width           =   855
            End
            Begin VB.CommandButton Command19 
               Caption         =   "Delete"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   275
               Left            =   7035
               Picture         =   "REPABOMB.frx":5E62
               TabIndex        =   27
               ToolTipText     =   "Elimina el Elemento Activo de la Lista de Repuestos"
               Top             =   970
               Width           =   855
            End
            Begin VB.CommandButton Command31 
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
               Left            =   3045
               TabIndex        =   26
               Top             =   420
               Width           =   175
            End
            Begin VB.CommandButton Command18 
               Caption         =   "Edit"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   275
               Left            =   7035
               Picture         =   "REPABOMB.frx":616C
               TabIndex        =   25
               ToolTipText     =   "Edición del Elemento Activo de la Lista"
               Top             =   1260
               Width           =   855
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Uso:"
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
               TabIndex        =   37
               Top             =   525
               Width           =   855
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Código:"
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
               TabIndex        =   36
               Top             =   525
               Width           =   960
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "U-M:"
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
               TabIndex        =   35
               Top             =   525
               Width           =   855
            End
            Begin VB.Label Label23 
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
               Left            =   105
               TabIndex        =   34
               Top             =   945
               Width           =   1065
            End
            Begin VB.Label Label40 
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
               Height          =   570
               Left            =   1260
               TabIndex        =   33
               Top             =   945
               Width           =   5580
            End
            Begin VB.Label Label41 
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
               Height          =   360
               Left            =   4305
               TabIndex        =   32
               Top             =   420
               Width           =   435
            End
         End
         Begin VB.Frame Frame11 
            BorderStyle     =   0  'None
            Caption         =   "Frame11"
            Height          =   855
            Left            =   8295
            TabIndex        =   92
            Top             =   4935
            Width           =   225
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            ScaleHeight     =   375
            ScaleWidth      =   9810
            TabIndex        =   22
            Top             =   390
            Width           =   9870
            Begin VB.Label Label27 
               BackColor       =   &H00E0E0E0&
               Caption         =   $"REPABOMB.frx":6476
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   23
               Top             =   105
               Width           =   9780
            End
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2910
            Left            =   210
            TabIndex        =   21
            Top             =   840
            Width           =   9885
         End
         Begin VB.Shape Shape1 
            Height          =   590
            Left            =   8510
            Top             =   5150
            Width           =   1590
         End
         Begin VB.Image Image1 
            Height          =   510
            Left            =   8295
            Picture         =   "REPABOMB.frx":64FF
            Stretch         =   -1  'True
            Top             =   5200
            Width           =   2010
         End
         Begin VB.Line Line1 
            X1              =   10080
            X2              =   10080
            Y1              =   375
            Y2              =   835
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5685
         Index           =   0
         Left            =   105
         TabIndex        =   10
         Top             =   525
         Width           =   10215
         Begin VB.Frame Frame1 
            Caption         =   "Especificaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5265
            Left            =   105
            TabIndex        =   70
            Top             =   420
            Width           =   9990
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   1155
               ScaleHeight     =   375
               ScaleWidth      =   7725
               TabIndex        =   77
               Top             =   525
               Width           =   7780
               Begin VB.Label Label22 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Descripción General - Especificaciones"
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
                  Left            =   0
                  TabIndex        =   78
                  Top             =   105
                  Width           =   7680
               End
            End
            Begin VB.CommandButton Command7 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   260
               Left            =   1155
               TabIndex        =   76
               Top             =   4755
               Width           =   2825
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Selecc Todo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   260
               Left            =   7635
               TabIndex        =   75
               Top             =   4755
               Width           =   1320
            End
            Begin VB.CommandButton Command6 
               Caption         =   "Pegar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   260
               Left            =   6345
               TabIndex        =   74
               Top             =   4755
               Width           =   1275
            End
            Begin VB.CommandButton Command8 
               Caption         =   "Copiar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   260
               Left            =   5175
               TabIndex        =   73
               Top             =   4755
               Width           =   1170
            End
            Begin VB.CommandButton Command17 
               Caption         =   "Cortar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   260
               Left            =   3990
               TabIndex        =   72
               Top             =   4755
               Width           =   1170
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
               Height          =   3795
               Left            =   1155
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   71
               Top             =   945
               Width           =   7785
            End
            Begin VB.Line Line2 
               X1              =   8935
               X2              =   8935
               Y1              =   525
               Y2              =   5010
            End
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   105
            TabIndex        =   17
            Top             =   170
            Width           =   1905
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   3795
         TabIndex        =   9
         Top             =   4305
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   4215
         TabIndex        =   8
         Top             =   3360
         Width           =   15
      End
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   86
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "REPABOMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOPEINI(3), IZQUINI(3), ANCHINI(3), ALTUINI(3)
'
Dim CIACTI%(15)
Dim ARCHEX$(15)
Dim YACAR%(15), MODIFIC%(15)
Dim GRABATODO%
Dim AGREITEM%
Dim MODOEDIT%
'
Dim RESECO$(128), DESOPE$(128)
'
Private Sub Command1_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   GRABATODO% = 1
   Call GRADATITEM
   GRABATODO% = 0
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Call COMUNI("Para Generar un Presupuesto\Acceda Desde el Escritorio Principal\de Reparaciones.")
   GoTo 99
   '
   If TRIM$(Text1) = "" Then GoTo 99
   If ECOARCH("INDIREP", Text1) = "" Then GoTo 99
   '
10 Call SELECHO("DEUDOR1")
   NC% = Val(VALACT1$("DEUDOR1"))
   If NC% > 0 Then
     If RASOCLI$(NC%) <> "" Then
       PREREBOM.Label18 = TRIM$(Str$(NUPRESUP&))
       PREREBOM.Label11 = TRIM$(Str$(NC%))
       PREREBOM.Label8 = Text28
       PREREBOM.Label10 = Text1
       '
       COEFI = Val(CONTROL$("REPARA", "COEFPREA"))
       If COEFI < 1 Then COEFI = 1
       CUHOR = Val(CONTROL$("REPARA", "VALHORA"))
       PREREBOM.Label20 = CSTRING$(MKS$(CUHOR), 4, 8, 2, 2)
       '
       PREREBOM.List1.Clear
       SUREPUES = 0
       Call SETULTREG("!MATEROF", 0)
       '
       CODD$ = Text1
       Screen.MousePointer = 11
       RFF$ = RECFILT$("REPUESRE", 1, CODD$)
       For IK& = 1 To Len(RFF$) Step 4
         RERE& = CVS(Mid$(RFF$, IK&, 4))
         If RERE& > 0 Then
           If RERE& <= ULTREG&("REPUESRE") Then
             XX$ = REGLEIDO$("REPUESRE", RERE&)
             CIY% = CVI(Mid$(XX$, 17, 2))
             CODXX$ = CODEXT$(CIY%)
             DEXX$ = DESCRIT$(CIY%)
             COSUN = COEFI * COSUNI(CIY%)
             COSUX$ = CSTRING$(MKS$(COSUN), 4, 10, 2, 2)
             USO = CVS(Mid$(XX$, 19, 4))
             USOX$ = CSTRING$(MKS$(USO), 4, 10, 1, 2)
             COSTO = COSUN * USO
             COSTX$ = CSTRING$(MKS$(COSTO), 4, 12, 2, 2)
             SUREPUES = SUREPUES + COSTO
             '
             XXX$ = Space$(64)
             Call REPLA(XXX$, CODXX$, 1, 16)
             Call REPLA(XXX$, DEXX$, 17, 38)
             Call REPLA(XXX$, USOX$, 54, 10)
             Call REPLA(XXX$, COSUX$, 64, 10)
             Call REPLA(XXX$, COSTX$, 74, 12)
             PREREBOM.List1.AddItem XXX$
           End If
         End If
       Next IK&
       PREREBOM.SUMAREP = CSTRING$(MKS$(SUREPUES), 4, 12, 2, 2)
       '
       SUTIEREP = 0
       RFFX$ = RECFILT$("SECOPREP", 1, CODD$)
       For I1& = 1 To Len(RFFX$) Step 4
         I& = CVS(Mid$(RFFX$, I1&, 4))
         X$ = REGLEIDO$("SECOPREP", I&)
         '
         CAOPS = CVS(Mid$(X$, 97, 4)): If CAOPS < 1 Then CAOPS = 1
         SUTIEREP = SUTIEREP + CAOPS * CVS(Mid$(X$, 73, 4))
       Next I1&
       PREREBOM.Label19 = CSTRING$(MKS$(SUTIEREP), 4, 8, 1, 2)
       '
       Screen.MousePointer = 1
       '
       PREREBOM.Label18 = TRIM$(Str$(NUPRESUP&))
       PREREBOM.Show 1
       GoTo 10
     End If
   End If
   '
99 Command10.Enabled = True
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command18_Click()
    '
    If List1.ListIndex < 0 Then
       Call MENSERR(24, "Opción Edición no Válida")
       Exit Sub
    End If
    '
    Text31.TEXT = Left$(List1.TEXT, 16)
    CDYY% = CODINT%(Text31.TEXT)
    Label40.Caption = DESCRIT0$(CDYY%)
    Label41.Caption = UNIMED$(CDYY%)
    Text32.TEXT = TRIM$(Mid$(List1.TEXT, 67, 12))
    '
    ILISTA& = List1.ListIndex
    List1.RemoveItem ILISTA&
    Command18.Enabled = False
    Command19.Enabled = False
    '
    If ILISTA& < List1.ListCount Then
        List1.ListIndex = ILISTA&
      Else
        List1.ListIndex = List1.ListCount - 1
    End If
    Text32.SetFocus
    MODOEDIT% = 1
    MODIFIC%(1) = 1
    '
End Sub

Private Sub Command19_Click()
    '
    If List1.ListIndex < 0 Then
       Call MENSERR(24, "Opción Borrado no Válida")
       Exit Sub
    End If
    '
    ILISTA& = List1.ListIndex
    List1.RemoveItem ILISTA&
    '
    If ILISTA& < List1.ListCount Then
        List1.ListIndex = ILISTA&
      Else
        List1.ListIndex = List1.ListCount - 1
    End If
    MODIFIC%(1) = 1
    '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   '
   If DERACCE%("OREPARA", "Generar / Cerrar Orden de Reparación") < 2 Then
     GoTo 99
   End If
   '
   OPORREP.Show 1
   '
99 Command2.Enabled = True
End Sub

Private Sub COMMAND21_CLICK()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command21.Enabled = False
   Call GRADATITEM
   '
   VDEF$ = Space$(128)
   '
   OKX% = 0
   Call ALTACODREP(VDEF$, OKX%)
   If OKX% = 1 Then
      CODINUE$ = Mid$(VDEF$, 1, 16)
      Text1.TEXT = CODINUE$
      SSTab1.Tab = 0
   End If
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command22_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command22.Enabled = False
   Call GRADATITEM
   '
   CODIVIE$ = AJUSTI$(Text1.TEXT, 16)
   If TRIM$(CODIVIE$) = "" Then GoTo 99
   '
   VDEF$ = Space$(128)
   Call REPLA(VDEF$, Text3.TEXT, 17, 64)
   '
   OKX% = 0
   Call ALTACODREP(VDEF$, OKX%)
   If OKX% = 1 Then
      CODINUE$ = Mid$(VDEF$, 1, 16)
      Call COPYFORMU(CODIVIE$, CODINUE$)
      Call COPYRUTA(CODIVIE$, CODINUE$)
      Call COPYSPEC(CODIVIE$, CODINUE$)
      Text1.TEXT = CODINUE$
      SSTab1.Tab = 0
   End If
   '
99 Command22.Enabled = True
   '
End Sub

Private Sub Command24_Click()
   '
   Command24.Enabled = False
Call OPNOIN("")
GoTo 99
   Dim EL As String * 2048
   '
   If VALCODART% < 1 Then
      Text1.SetFocus
      GoTo 99
   End If
   '
   CII% = CODINT%(Text1.TEXT)
   COD$ = AJUSTI$(CODEXT$(CII%), 15)
   '
   If COMALTER%(" Baja Lógica del Codigo '" + TRIM$(COD$) + "' \" + TRIM$(DESCRIT$(CI%)) + "\\Cancelar\Confirmar") <> 2 Then GoTo 99
   '
   Screen.MousePointer = 11
   Call BLANFORMU
   Screen.MousePointer = 11
   Text1.TEXT = ""
   If CII% < 1 Or CII% > NUMAS% Then
      Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      Command25.SetFocus
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   REGMA& = CII% + 1
   X$ = REGLEIDO$("MASTER", REGMA&)
   If Mid$(X$, 77, 15) <> AJUSTI$(COD$, 15) Then
      Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      Command25.SetFocus
      GoTo 99
   End If
   '
   Call REPLA(X$, Chr$(96), 92, 1)
   Call GRAREG("MASTER", X$, REGMA&)
   Call CIERRARCH("MASTER")
   '
10 NVII& = 1 + NUINV%
   LOTOTE& = 18 * ULTREG&("INVERT")
   If LOTOTE& < 18 * NVII& Then LOTOTE& = 18 * NVII&
   FIECO% = FILEOPEN%(LUDAT$ + "INVERT.DAT", 0, 2048)
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   While Len(TTXX$) < LOTOTE&: TTXX$ = TTXX$ + Chr$(0): Wend
   '
   For U& = 19 To Len(TTXX$) Step 18
      If Mid$(TTXX$, U&, 15) = COD$ Then
         Mid$(TTXX$, U&, 16) = COD$ + Chr$(96)
         GoTo 20
      End If
   Next U&
   Close #FIECO%
   Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
   Command25.SetFocus
   GoTo 99
   '
20 JJ& = 0
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$, U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   For U& = 1 To Len(PUNCODIG$) Step 16
      If Mid$(PUNCODIG$, U&, 15) = COD$ Then
         Mid$(PUNCODIG$, U&, 16) = COD$ + Chr$(96)
         GoTo 25
      End If
   Next U&
   Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
   Command25.SetFocus
   GoTo 99
   '
25 PPCC$ = PUNCODIG$
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   JI% = 0
   While Len(PPCC$) > 0
      BUFERCOD$ = Left$(PPCC$ + String$(2048, 0), 2048)
      JI% = JI% + 1
      Put #NX%, JI%, BUFERCOD$
      PPCC$ = Mid$(PPCC$, 2049)
   Wend
   Close #NX%
   '
   LOTOTE& = 64 * ULTREG&("ECOMAST")
   FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 2048)
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   '
   For U& = 1 To Len(TTXX$) Step 64
      If Mid$(TTXX$, U&, 15) = COD$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + Mid$(TTXX$, U& + 64)
         GoTo 30
      End If
   Next U&
   '
30 While (Len(TTXX$) Mod 2048) <> 0: TTXX$ = TTXX$ + Chr$(0): Wend
   JJ& = 0
   
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$, U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   Call FRESHECHO("MASTER")
   '
99 Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Private Sub Command26_Click()
   '
   CODXX$ = TRIM$(Text31.TEXT)
   If CODXX$ = "" Or CODINT%(CODXX$) < 1 Then
      Call MENSERR(24, "Código no Válido")
      Text31.SetFocus
      Exit Sub
   End If
   '
   CIY% = CODINT%(CODXX$)
   DEXX$ = DESCRIT$(CIY%)
   UNXX$ = UNIMED$(CIY%)
   USO = Val(Text32.TEXT)
   If Abs(USO) < 0.0001 Then
      Call MENSERR(24, "Uso no Válido")
      Text32.SetFocus
      Exit Sub
   End If
   '
   MODOEDIT% = 0
   USOX$ = CSTRING$(MKS$(USO), 4, 12, 4, 2)
   '
   XXX$ = Space$(64)
   Call REPLA(XXX$, CODXX$, 1, 16)
   Call REPLA(XXX$, DEXX$, 17, 28)
   Call REPLA(XXX$, UNXX$, 65, 2)
   Call REPLA(XXX$, USOX$, 67, 12)
   CAIT% = CAIT% + 1
   ILISTA& = List1.ListIndex
   If ILISTA& < 0 Then ILISTA& = 0
   '
   List1.AddItem XXX$, ILISTA&
   List1.ListIndex = ILISTA&
   List1.Refresh
   '
   Text31.TEXT = ""
   CODE$ = ""
   Label40.Caption = ""
   Label41.Caption = ""
   Text32.TEXT = ""
   Text31.SetFocus
   MODIFIC%(1) = 1
   '
   Command18.Enabled = True
   Command19.Enabled = True
   '
End Sub

Private Sub Command27_Click()
   '
   CODXX$ = TRIM$(Text31.TEXT)
   If CODXX$ = "" Or CODINT%(CODXX$) < 1 Then
      Call MENSERR(24, "Código no Válido")
      Text31.SetFocus
      Exit Sub
   End If
   '
   CIY% = CODINT%(CODXX$)
   DEXX$ = DESCRIT$(CIY%)
   UNXX$ = UNIMED$(CIY%)
   USO = Val(Text32.TEXT)
   If Abs(USO) < 0.0001 Then
      Call MENSERR(24, "Uso no Válido")
      Text32.SetFocus
      Exit Sub
   End If
   '
   MODOEDIT% = 0
   USOX$ = CSTRING$(MKS$(USO), 4, 12, 4, 2)
   '
   XXX$ = Space$(64)
   Call REPLA(XXX$, CODXX$, 1, 16)
   Call REPLA(XXX$, DEXX$, 17, 46)
   Call REPLA(XXX$, UNXX$, 65, 2)
   Call REPLA(XXX$, USOX$, 67, 12)
   CAIT% = CAIT% + 1
   '
   List1.AddItem XXX$
   List1.ListIndex = List1.ListCount - 1
   '
   Text31.TEXT = ""
   CODE$ = ""
   Label40.Caption = ""
   Label41.Caption = ""
   Text32.TEXT = ""
   Text31.SetFocus
   MODIFIC%(1) = 1
   '
   Command18.Enabled = True
   Command19.Enabled = True
   '
End Sub

Private Sub Command28_Click(Index As Integer)
   Call SELECHO("PADMAQ")
   VDEVU$ = VALACT1$("PADMAQ")
   If VDEVU$ <> "" Then
      Text9(Index).TEXT = VDEVU$
   End If
End Sub

Private Sub Command29_Click()
   Call SELECHO("CATOPER")
   VALONUE$ = VALACT1$("CATOPER")
   If VALONUE$ <> "" Then
      Text33.TEXT = VALONUE$
   End If
End Sub

Private Sub Command3_Click(Index As Integer)
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Command30_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command30.Enabled = False
   Call GRADATITEM
   For Index = 0 To 15: YACAR%(Index) = 0: Next Index
   Call COPYRUTAS
   Command30.Enabled = True
End Sub

Private Sub COMMAND31_CLICK()
   '
   Call SELECHO("MASTER")
   COCOMPO$ = TRIM$(VALACT1$("MASTER"))
   If COCOMPO$ <> "" Then
     Text31.TEXT = COCOMPO$
     CICOMPO% = CODINT%(COCOMPO$)
     Label40.Caption = DESCRIT0$(CICOMPO%)
     Label41.Caption = UNIMED$(CICOMPO%)
     If TRIM$(Text32.TEXT) = "" Then Text32.TEXT = Str$(CABAS)
     Text32.SetFocus
   End If
    '
End Sub

Private Sub Command32_Click()
   CODD$ = Text1.TEXT
   '
   ILISTA& = List3.ListIndex
   If ILISTA& >= 0 Then
     List3.RemoveItem (ILISTA&)
     For UI& = ILISTA& To List3.ListCount - 1
       RESECO$(UI&) = RESECO$(UI& + 1)
       DESOPE$(UI&) = DESOPE$(UI& + 1)
     Next UI&
   End If
   '
   Label55.Caption = ""
   Text5.TEXT = ""
   Label43 = ""
   Text11 = ""
   On Error Resume Next
     For KKI% = 0 To 4
       Text9(KKI%).TEXT = ""
       Text15(KKI%).TEXT = ""
       Text17(KKI%).TEXT = ""
       Text26(KKI%).TEXT = ""
     Next KKI%
   On Error GoTo 0
   '
   Text2.TEXT = ""
   Text30.TEXT = ""
   Text33.TEXT = ""
   Label51.Caption = ""
   MODIFIC(2) = 1
   '
   If List3.ListCount > 0 Then
     List3.ListIndex = 0
   End If
   '
End Sub

Private Sub Command33_Click()
   CODD$ = AJUSTI$(Text1.TEXT, 16)
   '
   NOPY% = 10
   If List3.ListCount > 0 Then
     ULI& = List3.ListCount
     NOPY% = 10 * (1 + Int(Val(Left$(List3.List(ULI& - 1), 9) / 10)))
   End If
   RESECO$(List3.ListCount) = Left$(CODD$ + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) 'para acomodar hoja de procesos DE FABRICACION
   DESOPE$(List3.ListCount) = ""
   NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
   dope$ = AJUSTI$("Operación Agregada (" + TRIM$(NOPEX$) + ")", 48)
   REGO& = 0
   REGOX$ = AJUSTD$(Str$(REGO&), 6)
   SECO$ = NOPEX$ + "     " + dope$ + " " + REGOX$
   MODIFIC%(2) = 1
   List3.AddItem SECO$
   List3.ListIndex = List3.ListCount - 1
   Text5.SetFocus
End Sub

Private Sub Command37_Click()
   Command37.Enabled = False
   CONFREPA.Show 1
   Command37.Enabled = True
End Sub

Private Sub Command39_Click()
   Command39.Enabled = False
   PRINTFORM
   Command39.Enabled = True
End Sub

Private Sub Command40_Click(Index As Integer)
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   '
   Command40(Index).Enabled = False
   '
   Call GRADATITEM
   Select Case Index
      Case 0
        Call INCORSTRUC
      Case 1
        'Call LIMULTI
      Case 2
        'Call LIMULIN
      Case 3
        'Call REDESPIE
      Case 4
        'Call COMULTI
      Case 5
        Call COPYSTRUC
      Case 6
        'Call LIMPLOSI
   End Select
   '
   Command40(Index).Enabled = True
   '
End Sub

Sub INCORSTRUC()
   '
   CODESTI$ = AJUSTI$(Text1.TEXT, 16)
   If TRIM$(CODESTI$) = "" Then
     Call MENSERR(24, "Active la Reparación\DESTINO de la Copia.")
     GoTo 99
   End If
   '
   If List1.ListCount > 0 Then
     If COMALTER%("Opciones de Incorporación:\\Agregar al Final\Sobre-Escribir") = 2 Then
       List1.Clear
       MODIFIC%(1) = 1
     End If
   End If
   '
   Call SELECHO("MASTER/Seleccione Fórmula de ORIGEN.")
   CODORIGI$ = TRIM$(VALACT1$("MASTER"))
   If CODORIGI$ <> "" Then
     Call LEEEXPLO(CODORIGI$, 1)
     For UI% = 1 To CAIT%
       CIY% = CIJ%(UI%)
       If CIY% > 0 Then
         If CIY% <= NUMAS% Then
           CODXX$ = CODEXT$(CIY%)
           DEXX$ = DESCRIT$(CIY%)
           UNXX$ = UNIMED$(CIY%)
           USO = USOI(UI%)
           USOX$ = CSTRING$(MKS$(USO), 4, 12, 4, 2)
           '
           XXX$ = Space$(64)
           Call REPLA(XXX$, CODXX$, 1, 16)
           Call REPLA(XXX$, DEXX$, 17, 46)
           Call REPLA(XXX$, UNXX$, 65, 2)
           Call REPLA(XXX$, USOX$, 67, 12)
           List1.AddItem XXX$
           MODIFIC%(1) = 1
         End If
       End If
     Next UI%
   End If
   '
99 End Sub
'
Private Sub Command43_Click()
   CODD$ = AJUSTI$(Text1.TEXT, 16)
   '
   ILISTAX& = List3.ListIndex
   If ILISTAX& < 0 Or ILISTAX& >= List3.ListCount Then
      Call MENSERR(24, "Imposible Insertar")
      Exit Sub
   End If
   '
   NOPEP% = Val(Left$(List3.List(ILISTAX&), 4))
   NOPEA% = 0
   If ILISTAX& > 0 Then
     NOPEA% = Val(Left$(List3.List(ILISTAX& - 1), 4))
   End If
   NOPY% = (NOPEA% + NOPEP%) / 2
   If NOPY% <= NOPEA% Or NOPY% >= NOPEP% Then
      Call MENSERR(24, "Imposible Insertar")
      Exit Sub
   End If
   '
   For UI& = List3.ListCount To ILISTAX& Step -1
     If UI& > 0 Then
       RESECO$(UI&) = RESECO$(UI& - 1)
       DESOPE$(UI&) = DESOPE$(UI& - 1)
     End If
   Next UI&
   '
   RESECO$(ILISTAX&) = Left$(CODD$ + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) 'para acomodar hoja de procesos DE FABRICACION
   DESOPE$(ILISTAX&) = ""
   NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
   dope$ = AJUSTI$("Operación Insertada (" + TRIM$(NOPEX$) + ")", 48)
   REGO& = 0
   REGOX$ = AJUSTD$(Str$(REGO&), 6)
   SECO$ = NOPEX$ + "     " + dope$ + " " + REGOX$
   MODIFIC%(2) = 1
   List3.AddItem SECO$, ILISTAX&

   List3.ListIndex = ILISTAX&
   Text5.SetFocus
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
'OK
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   Call SELECHO("INDIREP")
   COD$ = TRIM$(VALACT1$("INDIREP"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    SSTab1.Tab = 0
    Screen.MousePointer = 11
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call CARTAVALISTO
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   IMPLOSI.Hide
   On Error GoTo 0
   DoEvents
   If WindowState = 1 Then
       Call CIERRARCH("*.*")
       Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If MODOEDIT% = 1 Then
     Cancel = 1
     Call MENSAEDIT
     Exit Sub
   End If
   Call GRADATITEM
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Cancel = 0
   Call FINAL("")
End Sub

Private Sub List3_Click()
   Call MUEDATOPREP
End Sub


Function VALCODART%()
   VALCODART% = 1
   If CODINT%(Text1.TEXT) < 1 Then
      Call MENSERR(24, "Falta Ingresar Código de Artículo")
      VALCODART% = 0
      Text1.SetFocus
   End If
End Function
'
Sub BLANFORMU()
   '
   Text3.TEXT = ""
   '
   Text28.TEXT = ""
   '
   List1.Clear
   Text31 = ""
   Text32 = ""
   Text27 = ""
   Label40 = ""
   Label41 = ""
   '
   List3.Clear
   Label55.Caption = ""
   Text5.TEXT = ""
   Label43 = ""
   Text11 = ""
   On Error Resume Next
     For KKI% = 0 To 4
       Text9(KKI%).TEXT = ""
       Text15(KKI%).TEXT = ""
       Text17(KKI%).TEXT = ""
       Text26(KKI%).TEXT = ""
     Next KKI%
   On Error GoTo 0
   '
   Text2.TEXT = ""
   Text30.TEXT = ""
   Text33.TEXT = ""
   Label51.Caption = ""
   '
End Sub

Sub CARDATREP()
    '
    CODD$ = AJUSTI$(Text1.TEXT, 16)
    REGRE& = REGBUS&("INDIREP", 1, CODD$)
    Label2 = TRIM$(Str$(REGRE&))
    '
    Screen.MousePointer = 11
    XX$ = REGLEIDO$("INDIREP", REGRE&)
    Text3.TEXT = TRIM$(Mid$(XX$, 17, 64))
    MODIFIC(0) = 0
    '
    Select Case SSTab1.Tab
      Case 0
        Call CARDESREP
      Case 1
        Call CAREPUES
      Case 2
        Call CAROPREP
    End Select
    Screen.MousePointer = 1
    '
End Sub

Sub CARDESREP()
   '
   If YACAR%(0) < 1 Then
      '
      CODD$ = Text1.TEXT
      If TRIM$(CODD$) = "" Then Exit Sub
      '
      Text28.TEXT = "": TTYZ$ = ""
      Screen.MousePointer = 11
      RNC$ = RECFILT$("DESCREP", 1, CODD$)
      URN& = ULTREG&("DESCREP")
      '
      For U& = 1 To Len(RNC$) Step 4
        RENOTA& = CVS(Mid$(RNC$, U&, 4))
        If RENOTA& > 0 Then
          If RENOTA& <= URN& Then
            TTYY$ = RTrim$(Mid$(REGLEIDO$("DESCREP", RENOTA&), 17))
            If TTYZ$ <> "" Then
              TTYZ$ = TTYZ$ + Chr$(13) + Chr$(10)
            End If
            TTYZ$ = TTYZ$ + RTrim$(TTYY$)
          End If
        End If
      Next U&
      '
      Screen.MousePointer = 1
      LOCADE& = Len(TTYZ$)
      For U& = LOCADE& To 1 Step -1
        If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
      Next U&
      U& = 0
30    Text28.TEXT = Left$(TTYZ$, U&)
      Text28.Refresh
      '
      YACAR%(0) = 1
      MODIFIC%(0) = 0
      '
   End If
   '
End Sub
'
Sub CAREPUES()
   '
   If YACAR%(1) < 1 Then
      '
      List1.Clear
      CODD$ = Text1.TEXT
      Screen.MousePointer = 11
      RFF$ = RECFILT$("REPUESRE", 1, CODD$)
      For IK& = 1 To Len(RFF$) Step 4
        RERE& = CVS(Mid$(RFF$, IK&, 4))
        If RERE& > 0 Then
          If RERE& <= ULTREG&("REPUESRE") Then
            XX$ = REGLEIDO$("REPUESRE", RERE&)
            CIY% = CVI(Mid$(XX$, 17, 2))
            CODXX$ = CODEXT$(CIY%)
            DEXX$ = DESCRIT$(CIY%)
            UNXX$ = UNIMED$(CIY%)
            USO = CVS(Mid$(XX$, 19, 4))
            USOX$ = CSTRING$(MKS$(USO), 4, 12, 4, 2)
            '
            XXX$ = Space$(64)
            Call REPLA(XXX$, CODXX$, 1, 16)
            Call REPLA(XXX$, DEXX$, 17, 46)
            Call REPLA(XXX$, UNXX$, 65, 2)
            Call REPLA(XXX$, USOX$, 67, 12)
            List1.AddItem XXX$
          End If
        End If
      Next IK&
      '
      YACAR%(1) = 1
      MODIFIC%(1) = 0
      Screen.MousePointer = 1
      '
   End If
   '
End Sub
'
Sub CAROPREP()
   '
   If YACAR%(2) < 1 Then
      '
      CODD$ = Text1.TEXT
      '
      Screen.MousePointer = 11
      List3.Clear
      Label55 = ""
      Text5 = ""
      Text7 = ""
      Text11 = ""
      Label43 = ""
      Text29 = ""
      Text30 = ""
      Text33 = ""
      Text34 = ""
      Label51 = ""
      Text15(0) = ""
      Text17(0) = ""
      Text26(0) = ""
      Text9(0) = ""
      '
      ULTOP% = 0
10    RFFX$ = RECFILT$("SECOPREP", 1, CODD$)
      For I1& = 1 To Len(RFFX$) Step 4
        I& = CVS(Mid$(RFFX$, I1&, 4))
        X$ = REGLEIDO$("SECOPREP", I&)
        '
        NOPY% = CVI(Mid$(X$, 17, 2))
        If NOPY% > ULTOP% Then ULTOP% = NOPY%
        NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
        dope$ = Mid$(X$, 19, 48)
        REGO& = I&
        REGOX$ = AJUSTD$(Str$(REGO&), 6)
        SECO$ = NOPEX$ + "     " + dope$ + " " + REGOX$
        List3.AddItem SECO$
        RESECO$(List3.ListCount - 1) = Left$(X$ + String$(128, 0), 128) + Left$(NUHOPROF$(CI%, NOPY%) + Space$(16), 16)
        DESOPE$(List3.ListCount - 1) = ""
      Next I1&
      '
      RNC$ = RECFILT$("OBSOPREP", 1, CODD$)
      URN& = ULTREG&("OBSOPREP")
      '
      For U& = 1 To Len(RNC$) Step 4
        RENOTA& = CVS(Mid$(RNC$, U&, 4))
        If RENOTA& > 0 Then
          If RENOTA& <= URN& Then
            TTZZ$ = REGLEIDO$("OBSOPREP", RENOTA&)
            NOPY% = CVI(Mid$(TTZZ$, 17, 2))
            For KKI% = 1 To List3.ListCount
              If CVI(Mid$(RESECO$(KKI% - 1), 17, 2)) = NOPY% Then
                DESOPE$(KKI% - 1) = DESOPE$(KKI% - 1) + Mid$(TTZZ$, 19)
              End If
            Next KKI%
          End If
        End If
      Next U&
      '
      If List3.ListCount > 0 Then
        List3.ListIndex = 0
      End If
      '
      YACAR%(2) = 1
      MODIFIC%(2) = 0
      Screen.MousePointer = 1
      '
   End If
   '
End Sub
'
Sub MUEDATOPREP()
   '
   CODD$ = Text1.TEXT
   RETEMODI3% = MODIFIC%(2)
   '
   CABAS = 1
   '
   REGI& = Val(Mid$(List3.TEXT, 59, 6))
   NOPY% = Val(Left$(List3.TEXT, 4))
   dope$ = Mid$(List3.TEXT, 10, 48)
   '
   VDEF$ = RESECO$(List3.ListIndex)
   '
   Label55.Caption = TRIM$(Str$(NOPY%))
   Text5.TEXT = dope$
   '
   Text9(0).TEXT = AJUSTI$(Mid$(VDEF$, 67, 6), 6)
   Label43.Caption = ECOARCH$("PADMAQ", Text9(0).TEXT)
   '
   TPREP = CVS(Mid$(VDEF$, 73, 4))
   TSTT$ = HORMINSE$(60 * 60 * TPREP)
   Text15(0).TEXT = Left$(TSTT$, 3)
   Text17(0).TEXT = Mid$(TSTT$, 5, 2)
   Text26(0).TEXT = Mid$(TSTT$, 8, 2)
   '
   Text30.TEXT = TRIM$(CSTRING$(Mid$(VDEF$, 97, 4), 3, 5, 2, 2))
   Text33.TEXT = TRIM$(Str$(Asc(Mid$(VDEF$, 101, 1))))
   Label51.Caption = ECOARCH$("CATOPER", Chr$(Text33.TEXT))
   '
   Text2.TEXT = RTrim$(DESOPE$(List3.ListIndex))
   '
   Label60 = "": FEX = CVI(Mid$(RESECO$(List3.ListIndex), 117, 2))
   If FEX > 0 Then Label60 = FECHATEX$(FEX)
   Label62 = TRIM$(Mid$(RESECO$(List3.ListIndex), 119, 10))
   '
   MODIFIC%(2) = RETEMODI3%
   '
End Sub
'
Sub GRADATITEM()
   '
   CODD$ = Text1.TEXT
   If TRIM$(CODD$) = "" Then Exit Sub
   '
   If MODIFIC%(0) > 0 Then
       Command9.SetFocus
       REGRE& = REGBUS&("INDIREP", 1, CODD$)
       If REGRE& > 0 Then
         If REGRE& <= ULTREG&("INDIREP") Then
           XX$ = REGLEIDO$("INDIREP", REGRE&)
           Call REPLA(XX$, TRIM$(Text3), 17, 64)
           Call GRAREG("INDIREP", XX$, REGRE&)
         End If
       End If
       Call GRADESREP
     ElseIf MODIFIC%(1) > 0 Or (SSTab1.Tab = 1 And GRABATODO% = 1) Then
       Command9.SetFocus
       Call GRAREPUES
     ElseIf MODIFIC%(2) > 0 Then
       Command9.SetFocus
       Call GRAOPREP
   End If
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
End Sub
'
Sub GRADESREP()
   '
   CODD$ = Text1.TEXT
   If TRIM$(CODD$) = "" Then Exit Sub
   If YACAR%(0) < 1 Then Exit Sub
   If MODIFIC%(0) < 1 Then Exit Sub
   '
   If DERACCE%("MODESREP", "Modificación de Descripcion Reparaciones") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Descripción de la Reparación\'" + TRIM$(CODD$) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(0) = 0
       MODIFIC(0) = 0
       Call CARDESREP
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   URN& = ULTREG&("DESCREP")
   RNC$ = RECFILT$("DESCREP", 1, CODD$)
   KK& = (-3)
   '
   TTYZ$ = Text28.TEXT
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
   Next U&
   U& = 0
   '
10 TEXTO$ = Left$(TTYZ$, U&) + Chr$(13) + Chr$(10)
   Call FRATEXTO(TEXTO$, 62)
   For LL% = 1 To CARETEX%
      TTYY$ = AJUSTI$(CODD$, 16) + RETEX$(LL%) + Space$(62)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("DESCREP", TTYY$, RENOTA&)
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("DESCREP", Space$(16) + Space$(62), RENOTA&)
   Wend
   '
   Call CIERRARCH("DESCREP")
   MODIFIC(0) = 0
   Screen.MousePointer = 1
   '
End Sub
'
Sub GRAREPUES()
   '
   If YACAR%(1) < 1 Then Exit Sub
   If MODIFIC%(1) < 1 And GRABATODO% < 1 Then Exit Sub
   '
   If DERACCE%("MODIREPU", "Modificación Lista de Repuestos") < 2 Then
      Exit Sub
   End If
   '
   CODD$ = Text1
   If GRABATODO% <> 1 Then
     RESPU% = COMALTER%("Se ha Modificado la Lista de Repuestos\de la Reparación '" + TRIM$(Text1) + " - " + TRIM$(Text3) + "'\\Registrar Cambios\Cancelar")
     If RESPU% <> 1 Then
       YACAR%(1) = 0
       MODIFIC(1) = 0
       Call CAREPUES
       GoTo 99
     End If
   End If
   '
   Screen.MousePointer = 11
   URN& = ULTREG&("REPUESRE")
   RNC$ = RECFILT$("REPUESRE", 1, CODD$)
   KK& = (-3)
   '
   CARETEX% = List1.ListCount
   For LL% = 1 To CARETEX%
      CIXI% = CODINT%(Left$(List1.List(LL% - 1), 16))
      CANTU = Val(Mid$(List1.List(LL% - 1), 67, 12))
      TTYY$ = AJUSTI$(CODD$, 16) + MKI$(CIXI%) + MKS$(CANTU) + String$(42, 0)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("REPUESRE", TTYY$, RENOTA&)
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("REPUESRE", Space$(16) + String$(48, 0), RENOTA&)
   Wend
   '
   Call CIERRARCH("REPUESRE")
   MODIFIC(1) = 0
   Screen.MousePointer = 1
   '
99 Text31 = ""
   Text32 = ""
   Text27 = ""
   Label40 = ""
   Label41 = ""
   Command18.Enabled = True
   Command19.Enabled = True
   Screen.MousePointer = 1
   '
End Sub
'
Sub GRAOPREP()
   '
   CODD$ = AJUSTI$(Text1.TEXT, 16)
   If YACAR%(2) < 1 Then MODIFIC%(2) = 0: Exit Sub
   If MODIFIC%(2) < 1 Then Exit Sub
   '
   If DERACCE%("MODOPREP", "Modificación Operaciones de Reparación") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Secuencia\de Operaciones de la Reparación\'" + TRIM$(CODD$) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(2) = 0
       MODIFIC(2) = 0
       Call CAROPREP
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   URN& = ULTREG&("SECOPREP")
   RNC$ = RECFILT$("SECOPREP", 1, CODD$)
   KK& = (-3)
   '
   HOII% = HOY(HOS$)
   For LL% = 1 To List3.ListCount
      TTYY$ = CODD$ + Mid$(RESECO$(LL% - 1), 17, 112)
      Call REPLA(TTYY$, MKI$(HOII%), 117, 2)
      Call REPLA(TTYY$, USERNAME$, 119, 10)
      Mid$(RESECO$(LL% - 1), 1, 128) = Left$(TTYY$, 128)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
           XY$ = REGLEIDO$("SECOPREP", RENOTA&)
           If Left$(XY$, 16) <> CODD$ Then
              Call MENSERR(24, "Error de Grabación de Secuencia\de Operaciones de Reparación.")
              URN& = URN& + 1
              RENOTA& = URN&
           End If
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("SECOPREP", TTYY$, RENOTA&)
      '
      NOPEX% = CVI(Mid$(TTYY$, 17, 2))
      '
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("SECOPREP", String$(128, 0), RENOTA&)
   Wend
   '
   URN& = ULTREG&("OBSOPREP")
   RNC$ = RECFILT$("OBSOPREP", 1, CODD$)
   KK& = (-3)
   '
   For LL% = 1 To List3.ListCount
     '
     If TRIM$(DESOPE$(LL% - 1)) <> "" Then
       TEXTO$ = RTrim$(DESOPE$(LL% - 1))
       While Len(TEXTO$) > 0
         TTYY$ = CODD$ + Mid$(RESECO$(LL% - 1), 17, 2) + Left$(TEXTO$, 238)
         KK& = KK& + 4
         If KK& <= Len(RNC$) - 3 Then
             RENOTA& = CVS(Mid$(RNC$, KK&, 4))
           Else
             URN& = URN& + 1
             RENOTA& = URN&
         End If
         Call GRAREG("OBSOPREP", TTYY$, RENOTA&)
         TEXTO$ = Mid$(TEXTO$, 253)
       Wend
     End If
   Next LL%
   '
   While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("OBSOPREP", String$(256, 0), RENOTA&)
   Wend
   '
   Call CIERRARCH("SECOPREP")
   Call CIERRARCH("OBSOPREP")
   '
   MODIFIC(2) = 0
   Screen.MousePointer = 1
   '
End Sub
'
Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab <> PreviousTab Then
      '
      If MODOEDIT% = 1 Then
        If PreviousTab = 1 Then
          SSTab1.Tab = PreviousTab
          Call MENSAEDIT
        End If
        Exit Sub
      End If
      '
      Call GRADATITEM
      Call FRESHALL
      Call CARDATREP
      '
   End If
End Sub

Private Sub Text1_Change()
   '
'OK
   For KKI% = 0 To 15: YACAR%(KKI%) = 0: Next KKI%
   Call BLANFORMU
   Call CARDATREP
   '
End Sub

Private Sub Text1_DblClick()
   '
'OK
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   Call SELECHO("INDIREP")
   COD$ = TRIM$(VALACT1$("INDIREP"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

Private Sub Text1_GotFocus()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
End Sub

Private Sub Text15_Change(Index As Integer)
   If Val(Text15(Index).TEXT) > 999 Then Text15(Index).TEXT = "999"
   If Val(Text15(Index).TEXT) < 0 Then Text15(Index).TEXT = "  0"
   If List3.ListIndex >= 0 Then
     RESECOA$ = RESECO$(List3.ListIndex)
     TPAR = Val(Text15(Index).TEXT) + Val(Text17(Index).TEXT) / 60 + Val(Text26(Index).TEXT) / 3600
     If Index = 2 Then TPAR = 60 * TPAR / Val(Text7.TEXT)
     If Index <= 2 Then
         Mid$(RESECO$(List3.ListIndex), 73 + 4 * Index, 4) = MKS$(TPAR)
       ElseIf Index = 3 Then
         Mid$(RESECO$(List3.ListIndex), 93, 4) = MKS$(TPAR)
       ElseIf Index = 4 Then
         Mid$(RESECO$(List3.ListIndex), 109, 4) = MKS$(TPAR)
     End If
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(2) = 1
   End If
End Sub

Private Sub Text15_GotFocus(Index As Integer)
   Text15(Index).SelStart = 0
   Text15(Index).SelLength = Len(Text15(Index).TEXT)
End Sub

Private Sub Text15_LostFocus(Index As Integer)
   TTTT$ = TRIM$(CSTRING$(MKS$(Val(Text15(Index).TEXT)), 3, 3, 0, 2))
   If TTTT$ = "" Then TTTT$ = "0"
   While Len(TTTT$) < 3: TTTT$ = " " + TTTT$: Wend
   Text15(Index).TEXT = TTTT$
End Sub

Private Sub Text17_Change(Index As Integer)
   If Val(Text17(Index).TEXT) > 59 Then Text17(Index).TEXT = "59"
   If Val(Text17(Index).TEXT) < 0 Then Text17(Index).TEXT = " 0"
   If List3.ListIndex >= 0 Then
     RESECOA$ = RESECO$(List3.ListIndex)
     TPAR = Val(Text15(Index).TEXT) + Val(Text17(Index).TEXT) / 60 + Val(Text26(Index).TEXT) / 3600
     If Index = 2 Then TPAR = 60 * TPAR / Val(Text7.TEXT)
     If Index <= 2 Then
         Mid$(RESECO$(List3.ListIndex), 73 + 4 * Index, 4) = MKS$(TPAR)
       ElseIf Index = 3 Then
         Mid$(RESECO$(List3.ListIndex), 93, 4) = MKS$(TPAR)
       ElseIf Index = 4 Then
         Mid$(RESECO$(List3.ListIndex), 109, 4) = MKS$(TPAR)
     End If
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(2) = 1
   End If
End Sub

Private Sub Text17_GotFocus(Index As Integer)
   Text17(Index).SelStart = 0
   Text17(Index).SelLength = Len(Text17(Index).TEXT)
End Sub

Private Sub Text17_LostFocus(Index As Integer)
   TTTT$ = TRIM$(CSTRING$(MKS$(Val(Text17(Index).TEXT)), 3, 2, 0, 2))
   While Len(TTTT$) < 2: TTTT$ = "0" + TTTT$: Wend
   Text17(Index).TEXT = TTTT$
End Sub

Private Sub Text2_Change()
   On Error Resume Next
   DESOPE$(List3.ListIndex) = Text2
   On Error GoTo 0
   MODIFIC%(2) = 1
End Sub

Private Sub Text26_Change(Index As Integer)
   If Val(Text26(Index).TEXT) > 59 Then Text26(Index).TEXT = "59"
   If Val(Text26(Index).TEXT) < 0 Then Text26(Index).TEXT = " 0"
   If List3.ListIndex >= 0 Then
     RESECOA$ = RESECO$(List3.ListIndex)
     TPAR = Val(Text15(Index).TEXT) + Val(Text17(Index).TEXT) / 60 + Val(Text26(Index).TEXT) / 3600
     If Index = 2 Then TPAR = 60 * TPAR / Val(Text7.TEXT)
     If Index <= 2 Then
         Mid$(RESECO$(List3.ListIndex), 73 + 4 * Index, 4) = MKS$(TPAR)
       ElseIf Index = 3 Then
         Mid$(RESECO$(List3.ListIndex), 93, 4) = MKS$(TPAR)
       ElseIf Index = 4 Then
         Mid$(RESECO$(List3.ListIndex), 109, 4) = MKS$(TPAR)
     End If
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(2) = 1
   End If
End Sub

Private Sub Text26_GotFocus(Index As Integer)
   Text26(Index).SelStart = 0
   Text26(Index).SelLength = Len(Text26(Index).TEXT)
End Sub

Private Sub Text26_LostFocus(Index As Integer)
   TTTT$ = TRIM$(CSTRING$(MKS$(Val(Text26(Index).TEXT)), 3, 2, 0, 2))
   While Len(TTTT$) < 2: TTTT$ = "0" + TTTT$: Wend
   Text26(Index).TEXT = TTTT$
End Sub

Private Sub Text28_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text3_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text30_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 97, 4) = MKS(Val(Text30.TEXT))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(2) = 1
End Sub

Private Sub Text30_GotFocus()
   Text30.SelStart = 0
   Text30.SelLength = Len(Text30.TEXT)
End Sub

Private Sub Text30_LostFocus()
   Text30.TEXT = CSTRING$(MKS$(Val(Text30.TEXT)), 3, 5, 2, 2)
End Sub

Private Sub Text33_GotFocus()
   Text33.SelStart = 0
   Text33.SelLength = Len(Text33.TEXT)
End Sub

'
Private Sub Text31_DblClick()
   '
   Call SELECHO("MASTER")
   COCOMPO$ = TRIM$(VALACT1$("MASTER"))
   If COCOMPO$ <> "" Then
     Text31.TEXT = COCOMPO$
     CICOMPO% = CODINT%(COCOMPO$)
     Label40.Caption = DESCRIT0$(CICOMPO%)
     Label41.Caption = UNIMED$(CICOMPO%)
     If TRIM$(Text32.TEXT) = "" Then Text32.TEXT = Str$(CABAS)
     Text32.SetFocus
   End If
   '
End Sub

Private Sub Text31_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 13 Then
      CODY$ = TRIM$(Text31.TEXT)
      CIY% = CODINT%(CODY$)
      If CIY% > 0 Then
         CODE$ = CODY$: CIE% = CIY%
         Text31.TEXT = CODY$
         CODE$ = CODY$: CIE% = CIY%
         Label40.Caption = DESCRIT0$(CIE%)
         Label41.Caption = UNIMED$(CIE%)
         Text32.SetFocus
        Else
         Call MENSERR(24, "Código no Válido")
         Text31.TEXT = CODE$
      End If
   End If
   '
End Sub

Private Sub Text32_GotFocus()
   Text32.SelStart = 0
   Text32.SelLength = Len(Text32.TEXT)
End Sub

Private Sub Text33_Change()
   VALONU = Val(Text33.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text33.TEXT = "": Exit Sub
   End If
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 101, 1) = Chr$(VALONU)
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(2) = 1
   Label51.Caption = ECOARCH$("CATOPER", Chr$(VALONU))
End Sub

Private Sub Text33_DblClick()
   Call SELECHO("CATOPER")
   VALONUE$ = VALACT1$("CATOPER")
   If VALONUE$ <> "" Then
      Text33.TEXT = VALONUE$
   End If
End Sub

Private Sub Text5_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 19, 48) = AJUSTI$(Text5.TEXT, 48)
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(2) = 1
End Sub

Private Sub Text5_GotFocus()
   Text5.SelStart = 0
   Text5.SelLength = Len(Text5.TEXT)
End Sub

Private Sub Text5_LostFocus()
   ILISTA& = List3.ListIndex
   For UI& = 1 To List3.ListCount
     If Val(Left$(List3.List(UI& - 1), 4)) = Val(Label55) Then
       TTXT$ = List3.List(UI& - 1)
       List3.RemoveItem (UI& - 1)
       Call REPLA(TTXT$, Text5.TEXT, 10, 48)
       List3.AddItem TTXT$, UI& - 1
     End If
   Next UI&
   On Error Resume Next
   List3.ListIndex = ILISTA&
   On Error GoTo 0
End Sub

Private Sub Text9_Change(Index As Integer)
   If List3.ListIndex >= 0 Then
     If TRIM$(Text9(Index).TEXT) = "" Or ECOARCH$("PADMAQ", Text9(Index).TEXT) <> "" Then
        If Index = 0 Then
            Label43.Caption = ECOARCH$("PADMAQ", Text9(Index).TEXT)
            Mid$(RESECO$(List3.ListIndex), 67, 6) = AJUSTI$(Text9(Index).TEXT, 6)
          Else
            Mid$(RESECO$(List3.ListIndex), 111 + 6 * Index, 6) = AJUSTI$(Text9(Index).TEXT, 6)
        End If
     End If
   End If
   MODIFIC(2) = 1
End Sub

Private Sub Text9_DblClick(Index As Integer)
   Call SELECHO("PADMAQ")
   VDEVU$ = VALACT1$("PADMAQ")
   If VDEVU$ <> "" Then
      Text9(Index).TEXT = VDEVU$
   End If
End Sub

Sub ALTACODREP(VDEF$, OKX%)
   '
10 OBX$ = OBJPLA$("ALTACODREP", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   CODINUE$ = TRIM$(Left$(OBX$, 16))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > 15 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
      GoTo 10
   End If
   '
   If TRIM$(Mid$(OBX$, 17, 64)) = "" Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("INDIREP")
     XX$ = REGLEIDO$("INDIREP", U&)
     If TRIM$(Left$(XX$, 16)) = CODINUE$ Then
       Screen.MousePointer = 1
       Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
       GoTo 10
     End If
   Next U&
   '
   Call REGAPP("INDIREP", OBX$)
   Call CIERRARCH("INDIREP")
   Call BAJALDISCO
   OKX% = 1
   '
99 Screen.MousePointer = 1
   '
End Sub
'
Sub COPYFORMU(CODIVIE$, CODINUE$)
   '
   CODOR$ = AJUSTI$(CODIVIE$, 16)
   CODES$ = AJUSTI$(CODINUE$, 16)
   '
   Screen.MousePointer = 11
   Call CIERRARCH("*.*")
   '
   URE& = ULTREG("REPUESRE")
   For I& = 1 To URE&
      EST1$ = REGLEIDO$("REPUESRE", I&)
      If Left$(EST1$, 16) = CODES$ Then
        EST1$ = String$(128, 0)
        Call GRAREG("REPUESRE", EST1$, I&)
      End If
   Next I&
   '
   For I& = 1 To URE&
      EST1$ = REGLEIDO$("REPUESRE", I&)
      If Left$(EST1$, 16) = CODOR$ Then
        Call REPLA(EST1$, CODES$, 1, 16)
        Call REGAPP("REPUESRE", EST1$)
      End If
   Next I&
   '
   Call CIERRARCH("REPUESRE")
   Call CIERRARCH("*.*")
   '
   Call BAJALDISCO
   '
   Screen.MousePointer = 1
   '
End Sub

Sub COPYRUTA(CODIVIE$, CODINUE$)
   '
   Dim REGTX$(256)
   CODOR$ = AJUSTI$(CODIVIE$, 16)
   CODES$ = AJUSTI$(CODINUE$, 16)
   '
   Screen.MousePointer = 11
   PRIREORI& = 0
   PRIREDEST& = 0
   '
   RFFX$ = RECFILT$("SECOPREP", 1, CODOR$)
   RFFY$ = RECFILT$("SECOPREP", 1, CODES$)
   '
   K1% = 0
   URE& = ULTREG&("SECOPREP")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("SECOPREP", RESE&)
         Call REPLA(X$, CODES$, 1, 16)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("SECOPREP")
     End If
     VALANT$ = REGLEIDO$("SECOPREP", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("SECOPREP", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("SECOPREP")
     Call GRAREG("SECOPREP", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("SECOPREP")
   '
   RFFX$ = RECFILT$("OBSOPREP", 1, CODOR$)
   RFFY$ = RECFILT$("OBSOPREP", 1, CODES$)
   '
   K1% = 0
   URE& = ULTREG&("OBSOPREP")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("OBSOPREP", RESE&)
         Call REPLA(X$, CODES$, 1, 16)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("OBSOPREP")
     End If
     VALANT$ = REGLEIDO$("OBSOPREP", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("OBSOPREP", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("OBSOPREP")
     Call GRAREG("OBSOPREP", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("OBSOPREP")
   '
   Screen.MousePointer = 1
   '
End Sub
'
Sub COPYSPEC(CODIVIE$, CODINUE$)
   '
   Dim REGTX$(1024)
   CODOR$ = AJUSTI$(CODIVIE$, 16)
   CODES$ = AJUSTI$(CODINUE$, 16)
   Screen.MousePointer = 11
   RFFX$ = RECFILT$("DESCREP", 1, CODOR$)
   RFFY$ = RECFILT$("DESCREP", 1, CODES$)
   '
   K1% = 0
   URE& = ULTREG&("DESCREP")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("DESCREP", RESE&)
         Call REPLA(X$, CODES$, 1, 16)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("DESCREP")
     End If
     VALANT$ = REGLEIDO$("DESCREP", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("DESCREP", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("DESCREP")
     Call GRAREG("DESCREP", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("DESCREP")
   Screen.MousePointer = 1
   '
End Sub
'
Sub CARTAVALISTO()
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
    Screen.MousePointer = 1
    '
End Sub
'


Sub COPYSTRUC()
   '
   CODORI$ = AJUSTI$(Text1.TEXT, 16)
   If TRIM$(CODORI$) = "" Then
     Call MENSERR(24, "Active la Reparación\ORIGEN de la Copia.")
     GoTo 99
   End If
   '
   Call SELECHO("INDIREP/Seleccione Reparación DESTINO de la Copia")
   CODEST$ = TRIM$(VALACT1$("INDIREP"))
   If CODEST$ = "" Then GoTo 99
   If TRIM$(CODEST$) = TRIM$(CODORI$) Then
     Call MENSERR(24, "Imposible Copiar Sobre SI MISMA")
     GoTo 99
   End If
   '
   Call COPYFORMU(CODORI$, CODEST$)
   '
   Text1.TEXT = CODEST$
   Text1.Refresh
   '
99 End Sub

Sub COPYRUTAS()
   '
   CODORI$ = AJUSTI$(Text1.TEXT, 16)
   If TRIM$(CODORI$) = "" Then GoTo 99
   '
   Call SELECHO("INDIREP")
   CODEST$ = TRIM$(VALACT1$("INDIREP"))
   If CODEST$ = "" Then GoTo 99
   If TRIM$(CODEST$) = TRIM$(CODORI$) Then GoTo 99
   '
   Call COPYRUTA(CODORI$, CODEST$)
   '
   Text1.TEXT = CODEST$
   Text1.Refresh
   Screen.MousePointer = 1
   '
99 End Sub
'
Private Sub Text9_GotFocus(Index As Integer)
   Text9(Index).SelStart = 0
   Text9(Index).SelLength = Len(Text9(Index).TEXT)
End Sub

Sub MENSAEDIT()
   Call MENSERR(24, "Operación no Válida.\  \Complete Previamente la Edición de la Lista de Repuestos\Haciendo Click sobre 'Append' o 'Insert'.\  \Luego Vuelva a Intentarlo.")
   Command26.SetFocus
End Sub

Function NUPRESUP&()
   '
   NPP& = 1
   '
   FIN& = ULTREG&("PRESREPA")
   If FIN& > 0 Then
      For U& = FIN& To FIN& - 128 Step -1
         If U& < 1 Then GoTo 99
         XX$ = REGLEIDO$("PRESREPA", U&)
         NPPX& = CVS(Mid$(XX$, 1, 4))
         If NPPX& >= NPP& Then
            NPP& = 1 + NPPX&
         End If
      Next U&
   End If
   '
99 NUPRESUP& = NPP&
   '
End Function
'

