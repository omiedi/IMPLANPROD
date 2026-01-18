VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPDISTK07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disponibilidad General de Stocks - Configuración"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   3120
      TabIndex        =   47
      Top             =   4075
      Width           =   175
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
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
      Left            =   4700
      MaxLength       =   2
      TabIndex        =   43
      Text            =   " "
      ToolTipText     =   "Rango de dias Proceso"
      Top             =   3720
      Width           =   540
   End
   Begin VB.TextBox Text6 
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
      Left            =   2145
      TabIndex        =   41
      Text            =   " "
      ToolTipText     =   "FORMATO: HH:MM:SS"
      Top             =   3720
      Width           =   1065
   End
   Begin VB.PictureBox Picture999 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   40
      Top             =   240
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00CDDADA&
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
      Left            =   105
      TabIndex        =   37
      Top             =   1300
      Width           =   5160
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Disponibilidad de Stocks con detalle de O.Compra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   105
         TabIndex        =   38
         Top             =   210
         Width           =   4950
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONFIGURACIÓN DE MAIL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      TabIndex        =   34
      Top             =   2880
      Width           =   5130
      Begin VB.TextBox Text19 
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
         Left            =   1440
         TabIndex        =   35
         Text            =   " "
         Top             =   260
         Width           =   3540
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email Ventas"
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
         TabIndex        =   36
         Top             =   360
         Width           =   1245
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONTROL"
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
      Left            =   4095
      TabIndex        =   32
      Top             =   1850
      Width           =   1170
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         Picture         =   "OPDISTK07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Realiza Control de Movimientos de Stock Repetidos o Duplicados"
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00CDDADA&
      Caption         =   "REGISTRO DE TRAZABILIDAD"
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
      TabIndex        =   29
      Top             =   1850
      Width           =   3900
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Por Declaración de Consumo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   1995
         TabIndex        =   31
         Top             =   315
         Width           =   1800
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Automática por F.I.F.O."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   420
         TabIndex        =   30
         Top             =   315
         Value           =   -1  'True
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4605
      Left            =   5460
      ScaleHeight     =   4575
      ScaleWidth      =   1140
      TabIndex        =   26
      Top             =   0
      Width           =   1170
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   49
         Top             =   4080
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   50
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Ubica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "OPDISTK07.frx":030A
         TabIndex        =   45
         Top             =   2160
         Width           =   645
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   105
         Picture         =   "OPDISTK07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   260
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "OPDISTK07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   870
         Width           =   645
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Presentación de Costo y Precios"
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
      Left            =   240
      TabIndex        =   22
      Top             =   8400
      Width           =   5160
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Calculado"
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
         Left            =   1920
         TabIndex        =   25
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Precio Base (C.Rep. + 15 %)"
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
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Informado"
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
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "FORMULARIO SITUACIÓN DE STOCK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   105
      TabIndex        =   16
      Top             =   105
      Width           =   5160
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
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
         Left            =   4095
         Picture         =   "OPDISTK07.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   315
         Width           =   750
      End
      Begin VB.TextBox Text5 
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
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   " "
         Top             =   680
         Width           =   3750
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3780
         TabIndex        =   18
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text4 
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
         Left            =   2380
         TabIndex        =   17
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario"
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
         Left            =   0
         TabIndex        =   21
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Origen de Stocks Remotos"
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
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   5160
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
         Height          =   360
         Index           =   3
         Left            =   735
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1995
         Width           =   4230
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
         Height          =   360
         Index           =   2
         Left            =   735
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4230
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
         Height          =   360
         Index           =   1
         Left            =   735
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   945
         Width           =   4230
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
         Height          =   360
         Index           =   0
         Left            =   735
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   420
         Width           =   4230
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-4:"
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
         Left            =   0
         TabIndex        =   15
         Top             =   2100
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-3:"
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
         Left            =   0
         TabIndex        =   13
         Top             =   1575
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-2:"
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
         Left            =   0
         TabIndex        =   11
         Top             =   1050
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-1:"
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
         Left            =   0
         TabIndex        =   9
         Top             =   525
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Destino Archivo de Transferencia"
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
      Left            =   210
      TabIndex        =   4
      Top             =   7440
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
         Height          =   360
         Left            =   735
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   420
         Width           =   4230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta:"
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
         TabIndex        =   6
         Top             =   525
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Identificador de Sede"
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
      Left            =   210
      TabIndex        =   1
      Top             =   5565
      Width           =   5160
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
         Height          =   360
         Left            =   3780
         MaxLength       =   8
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código Identificador de Sede:"
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
         TabIndex        =   3
         Top             =   525
         Width           =   3480
      End
   End
   Begin VB.CommandButton Command7 
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
      Height          =   855
      Left            =   5400
      Picture         =   "OPDISTK07.frx":0D72
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "OPDISTK07.frx":107C
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPDISTK07.frx":12B0
      TabIndex        =   39
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2145
      TabIndex        =   48
      Top             =   4075
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicio F.Kardex:"
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
      TabIndex        =   46
      Top             =   4185
      Width           =   1875
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Días Rango"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3720
      TabIndex        =   44
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comenzar Proceso a las:"
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
      Left            =   -120
      TabIndex        =   42
      Top             =   3840
      Width           =   2115
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5040
      Picture         =   "OPDISTK07.frx":132A
      Top             =   6950
      Width           =   2010
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuTablasAuxiliares 
         Caption         =   "Tablas Auxiliares"
      End
   End
End
Attribute VB_Name = "OPDISTK07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    RANPRO% = XVALO(Text7)
    If RANPRO% > 31 Then
       Call COMUNI("Rango Seleccionado No Válido Máximo 31")
       GoTo 99
    End If

    Call GRACONTROL("", "FORDISTO", Text4.TEXT)
    '
    MODOTRAZA$ = "FIFO"
    If Option5.Value = True Then
       If EXISTE%("TRAZAB04.EXE") < 1 Then
          Call MENSERR(24, "Trazabilidad por Declaración de Consumo\No Disponible en Presente Versión.\   \Falta Aplicativo 'TRAZAB04.EXE'.")
          Option4.Value = True
          GoTo 99
       End If
       MODOTRAZA$ = "DECONSU"
    End If
    Call GRACONTROL("", "RETRAZA", MODOTRAZA$)
    '
    '
    ' los que siguen están deshabilitados - son de versión 'DAT'
    IDENSEDE$ = TRIM$(Text1.TEXT)
    Call GRACONTROL("", "IDENSEDE", Text1.TEXT)
    '
    DESTITRA$ = TRIM$(Text2.TEXT)
    Call GRACONTROL("", "DESTITRA", Text2.TEXT)
    '
    For U& = 0 To 3
       ORITRA$(U&) = TRIM$(Text3(U&).TEXT)
    Next U&
    '
    Call GRACONTROL("", "ORITRA-1", Text3(0).TEXT)
    Call GRACONTROL("", "ORITRA-2", Text3(1).TEXT)
    Call GRACONTROL("", "ORITRA-3", Text3(2).TEXT)
    Call GRACONTROL("", "ORITRA-4", Text3(3).TEXT)
    '
    MODOCORE% = 1: MODOCOREP$ = "COREPOS"
    If Option2.Value = True Then
         MODOCORE% = 2: MODOCOREP$ = "PREBASE"
       ElseIf Option3.Value = True Then
         MODOCORE% = 3: MODOCOREP$ = "COSCALCU"
    End If
    Call GRACONTROL("SITUASTO", "MODOCORE", MODOCOREP$)
    '
    'E-mail para Ventas
    Call GRACONTROL("DIREMAI", "MAILVNTA", Text19)
    '
    'OPCION DE LISTADO DE GENERADOR DE REPORTES CON DETALLE DE ORDEN DE COMPRA
    DETAOCOM$ = ""
    If Check1.Value = 1 Then
       DETAOCOM$ = "SI"
    End If
    Call GRACONTROL("SITUASTO", "DISPONIB", DETAOCOM$)
    '
    INICIOPRO$ = ""
    HCOM% = Int(HORANUM(Text6))
    If HCOM% > 0 Then INICIOPRO$ = Text6
    Call GRACONTROL("SITUASTO", "HORAPROC", INICIOPRO$)
    '
    RANPRO% = XVALO(Text7)
    Call GRACONTROL("SITUASTO", "RANGPROC", Str$(RANPRO%))
    '
    Call GRACONTROL("", "FEINIKAR", Label8)

    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completada.\Re-Inicie Ahora la Aplicación '" + UCase$(App.EXEName) + "'.")
    Call FINAL("")
    '
99  Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   '
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT FeReMovi, FECHMOV, COD_INTE, DOPRILAR, CANTINGRE, CANTSALID, IdenLote, NUORITEM FROM MOVISTO WITH(NOLOCK) "
   SQLX$ = SQLX$ + "order by FeReMovi ASC"
   '
   'Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
'   Dim MoviTemp As ADODB.Recordset
'   Set MoviTemp = New ADODB.Recordset
'   MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   Set MoviTemp = READSET(SQLX$)
   With MoviTemp
'      On Error Resume Next
'      .MoveLast
'      If Err < 1 Then
'         On Error GoTo 0
'         FIN& = .RecordCount
'         .MoveFirst
         KU& = 0: APADUP% = 0
         Do While Not .EOF
           KU& = KU& + 1
           Call TRACE(20, KU&, FIN&)
           FEREMO = !FeReMovi
           FEMO = !FechMov
           CIXI% = !cod_inte
           DOPR$ = !DOPRILAR
           CAIN = !CANTINGRE
           CASA = !CANTSALID
           IDLO$ = SAFETEXT$(!idenlote)
           NORI% = !NuOrItem
           '
           If TRIM$(DOPR$) <> "" And CIXI% > 0 Then
              APOS = .AbsolutePosition
              '
              If APOS < FIN& - 1 Then
                 Do While !FeReMovi = FEREMO
                   .MoveNext
                   If FEMO = !FechMov Then
                     If CIXI% = !cod_inte Then
                       If DOPR$ = !DOPRILAR Then
                         If CAIN = !CANTINGRE Then
                           If CASA = !CANTSALID Then
                             If IDLO$ = !idenlote Then
                               If NORI% = !NuOrItem Then
                                 APADUP% = 1
                                 MsgBox "Aparente Duplicación en " + DOPR$ + " - " + CODEXT$(CIXI%)
                               End If
                             End If
                           End If
                         End If
                       End If
                     End If
                   End If
                 Loop
                 '
                 .AbsolutePosition = APOS
              End If
           End If
         .MoveNext
         Loop
'      End If
'      On Error GoTo 0
   End With
   Set MoviTemp = Nothing
   '
   Screen.MousePointer = 1
   TTXX$ = "Fin del Control"
   If APADUP% < 1 Then TTXX$ = TTXX$ + " sin Movimientos Duplicados."
   MsgBox TTXX$
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
10  OPX% = COMALTER%("Ubicaciones\\ABM de Sectores\ABM de Ubicaciones\Ubicacion de Codigos")
    If OPX% < 1 Or OPX% > 4 Then GoTo 99
    '
    If OPX% = 1 Then
        Call TRALOCAL("UBICASEC", "")
        VTB% = VENTABU%("UBICASEC")
        Call TRACENTRAL("UBICASEC", "")
        GoTo 10
    ElseIf OPX% = 2 Then
        ' LEE INFORMACION DE LA TABLA SQL
        SQLX$ = "SELECT * FROM UBICAVOL WITH(NOLOCK) WHERE SECTOR<>''"
        Set RSTR0 = READSET(SQLX$)
        With RSTR0
            JJ& = 0: REBLAN$ = REGBLAN$("UBICAVOL")
            Do While Not .EOF
                XX$ = REBLAN$
                Call REPLA(XX$, SAFETEXT$(!SECTOR), 1, 2)
                Call REPLA(XX$, SAFETEXT$(!ESTANTERIA), 3, 2)
                Call REPLA(XX$, SAFETEXT$(!familia), 5, 2)
                Call REPLA(XX$, SAFETEXT$(!Fila), 7, 2)
                Call REPLA(XX$, SAFETEXT$(!COLUMNA), 9, 2)
                Call REPLA(XX$, MKD$(XVALO(!VOLUMEN)), 11, 8)
                JJ& = JJ& + 1
                Call GRAREG("!UBICAVOL", XX$, JJ&)
                .MoveNext
            Loop
        End With
        On Error Resume Next
        RSTR0.Close
        Set RSTR0 = Nothing
        On Error GoTo 0
        Call SETULTREG("!UBICAVOL", JJ&)
        
        ' PRESENTA INFORMACION
        VTB% = VENTABU%("UBICAVOL")
        If VTB% < 0 Then GoTo 10
        
        Call ACTUREGISTRO("UBICAVOL", "MARBORRA", "SECTOR<>''", "1", REGAF&)
        
        SQLX$ = "SELECT * FROM UBICAVOL WHERE MARBORRA=2"
        Dim RSRW0 As ADODB.Recordset
        Set RSRW0 = New ADODB.Recordset
25      On Error Resume Next
        RSRW0.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        '
        With RSRW0
            For i& = 1 To ULTREG&("!UBICAVOL")
                XX$ = REGLEIDO$("!UBICAVOL", i&)
                .AddNew
                !SECTOR = Mid$(XX$, 1, 2)
                !ESTANTERIA = Mid$(XX$, 3, 2)
                !familia = Mid$(XX$, 5, 2)
                !Fila = Mid$(XX$, 7, 2)
                !COLUMNA = Mid$(XX$, 9, 2)
                !VOLUMEN = CVD(Mid$(XX$, 11, 8))
                !MARBORRA = 0
                .Update
            Next i&
        End With
        On Error Resume Next
        RSRW0.Close
        Set RSRW0 = Nothing
        On Error GoTo 0
        '
        Call BORRAREGISTROS("UBICAVOL", "MARBORRA=1", REGAF&)
        GoTo 10
        '
    ElseIf OPX% = 3 Then
        '
        Call CARGALISEL("SECUBICA", "UBICAVOL", "SECTOR/A2+UBICASEC/A15;ESTANTERIA/A3;FILA/A3;COLUMNA/A3;FAMILIA/A3+TAFAMIL/A15", "SECTOR<>''")
30      Call SELECHO("SECUBICA/UBICACION DE CODIGOS")
        SECTOR$ = TRIM$(VALACT1$("SECUBICA"))
        If SECTOR$ = "" Then GoTo 10
        
        TX$ = VALACT2$("SECUBICA")
        ESTANTERIA$ = TRIM$(Mid$(TX$, 17, 3))
        Fila$ = TRIM$(Mid$(TX$, 21, 3))
        Columna1$ = TRIM$(Mid$(TX$, 25, 3))
        familia$ = TRIM$(Mid$(TX$, 29, 3))
        '
        CONDI$ = ""
        CONDI$ = CONDI$ & " SECTOR='" & SECTOR$ & "'"
        CONDI$ = CONDI$ & " AND ESTANTERIA='" & ESTANTERIA$ & "'"
        CONDI$ = CONDI$ & " AND FILA='" & Fila$ & "'"
        CONDI$ = CONDI$ & " AND COLUMNA='" & Columna1$ & "'"
        SQLX$ = "SELECT CODIGO FROM UBICACOD WITH(NOLOCK) WHERE " & CONDI$
        '
        Set RSTR1 = READSET(SQLX$)
        With RSTR1
            JJ& = 0: REBLAN$ = REGBLAN$("MATEROF")
            Do While Not .EOF
                XX$ = REBLAN$
                Call REPLA(XX$, MKI$(SAFETEXT$(!Codigo)), 1, 2)
                JJ& = JJ& + 1
                Call GRAREG("!MATEROF", XX$, JJ&)
                .MoveNext
            Loop
        End With
        On Error Resume Next
        RSTR1.Close
        Set RSTR1 = Nothing
        On Error GoTo 0
        Call SETULTREG("!MATEROF", JJ&)
        
        Call COPYSTRU("MASTER", "CODFAMI")
        Call CARGALISEL("CODFAMI", "MASTER", "CODINT/C16+MASTER/A50", "FAMILIA='" & familia$ & "'")

        ' PRESENTA INFORMACION
        VTB% = VENTABU%("UBICACOD")
        If VTB% < 0 Then GoTo 30
        
        Call ACTUREGISTRO("UBICACOD", "MARBORRA", CONDI$, "1", REGAF&)
        
        SQLX$ = "SELECT * FROM UBICACOD WHERE MARBORRA=2"
        Dim RSRW1 As ADODB.Recordset
        Set RSRW1 = New ADODB.Recordset
35      On Error Resume Next
        RSRW1.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 35
        End If
        On Error GoTo 0
        '
        With RSRW1
            For i& = 1 To ULTREG&("!MATEROF")
                XX$ = REGLEIDO$("!MATEROF", i&)
                .AddNew
                !SECTOR = SECTOR$
                !ESTANTERIA = ESTANTERIA$
                !familia = familia$
                !Fila = Fila$
                !COLUMNA = Columna1$
                !Codigo = CVI(Mid$(XX$, 1, 2))
                !MARBORRA = 0
                .Update
            Next i&
        End With
        On Error Resume Next
        RSRW1.Close
        Set RSRW1 = Nothing
        On Error GoTo 0
        '
        Call BORRAREGISTROS("UBICACOD", "MARBORRA=1", REGAF&)
        '
        GoTo 30
    End If
    '
99  Command6.Enabled = True
    '
End Sub

Private Sub Command8_Click()
    Call SELECHO("SELFECHA")
    VDEVU$ = VALACT1$("SELFECHA")
    If VDEVU$ <> "" Then
        Label8 = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Form_Load()
    '

    Call TRANSLABELS(Name)
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
    Text4.TEXT = Control$("", "FORDISTO")
    Text5.TEXT = ECOARCH$("INDIFRM", Text4.TEXT)
    '
    MODOTRAZA$ = Control$("", "RETRAZA")
    If MODOTRAZA$ = "DECONSU" Then
       Option5.Value = True
    End If
    '
    '
    ' los que siguen están deshabilitados - son de versión 'DAT'
    Call CARUTRA
    '
    Text1.TEXT = IDENSEDE$
    Text2.TEXT = DESTITRA$
    For U& = 0 To 3
       Text3(U&).TEXT = ORITRA$(U&)
    Next U&
    '
    Option1.Value = True
    MODOCORE% = 1: MODOCOREP$ = "COREPOS"
    If TRIM$(UCase$(Control$("SITUASTO", "MODOCORE"))) = "PREBASE" Then
       Option2.Value = True
       MODOCORE% = 2: MODOCOREP$ = "PREBASE"
     ElseIf TRIM$(UCase$(Control$("SITUASTO", "MODOCORE"))) = "COSCALCU" Then
       Option3.Value = True
       MODOCORE% = 3: MODOCOREP$ = "COSCALCU"
    End If
    Text19 = Control$("DIREMAI", "MAILVNTA")
    '
    Check1.Value = 0
    If Control("SITUASTO", "DISPONIB") = "SI" Then
       Check1.Value = 1
    End If
    Text6 = TRIM$(Control("SITUASTO", "HORAPROC"))
    Text7 = TRIM$(Control("SITUASTO", "RANGPROC"))
    Label8 = Control("", "FEINIKAR")
    
End Sub

Private Sub Label8_DblClick()
10 OBX$ = OBJPLA$("ENTRAFECHA1", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   FE% = CVI(Mid$(OBX$, 1, 2))
   '
   Label8 = FECHATEX$(FE%)
   
End Sub

Private Sub mnuTablasAuxiliares_Click()
  Call GENERAR_ESTRUCTURAS_SQL
  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set RS1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(RS1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.msf2)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '
  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))
  If InStr(EPAC$, "ENVAPLAST") > 0 Then
    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
    If LIXTA$ = "MATERIALES ENVAPLAST" Then
       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
       Exit Sub
    End If
  End If

  Call ABM_TABLA(IDLIST&)

End Sub

Private Sub Option4_Click()
    Call COMUNI("Opción Solo Disponible como\Versión Beta (Experimental).")
End Sub

Private Sub Option5_Click()
    If EXISTE%("TRAZAB04.EXE") < 1 Then
       Call MENSERR(24, "No Disponible en Presente Versión.\   \Falta Aplicativo 'TRAZAB04.EXE'.")
       Option4.Value = True
    End If
End Sub

Private Sub Text4_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text5_Click()
   Text4.SetFocus
End Sub
Private Sub Text6_LostFocus()
    Text6 = HORATEX$(HORANUM(Text6))
End Sub
Private Sub Text6_GotFocus()
  Text6.SelStart = 0
  Text6.SelLength = Len(Text6)
End Sub

  
