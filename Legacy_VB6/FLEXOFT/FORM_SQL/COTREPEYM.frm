VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form COTREPEYM 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cotización Asistencia Técnica"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   16875
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   16875
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
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
      Height          =   300
      Index           =   12
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   79
      TabStop         =   0   'False
      Text            =   "Condición de Pago"
      Top             =   8880
      Width           =   1650
   End
   Begin VB.TextBox Eco21 
      Appearance      =   0  'Flat
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
      Height          =   300
      Left            =   9120
      TabIndex        =   78
      Text            =   " "
      Top             =   9210
      Width           =   3372
   End
   Begin VB.CommandButton Command21 
      Caption         =   "."
      Height          =   300
      Index           =   1
      Left            =   11280
      TabIndex        =   77
      Top             =   8880
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox TEXT211 
      Appearance      =   0  'Flat
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
      Height          =   300
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   76
      Text            =   " "
      Top             =   8880
      Width           =   435
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   300
      Index           =   11
      Left            =   14500
      Locked          =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Text            =   "Accesorios:"
      Top             =   8250
      Width           =   1280
   End
   Begin VB.TextBox TXTTOTALCONACCESORIOS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   300
      Left            =   15795
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   9195
      Width           =   1005
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   300
      Index           =   10
      Left            =   13665
      Locked          =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Text            =   "Total Con Accesorios:"
      Top             =   9195
      Width           =   2115
   End
   Begin VB.TextBox TXTIMPODESC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   300
      Left            =   15795
      Locked          =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   8880
      Width           =   1005
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   300
      Index           =   9
      Left            =   13665
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Text            =   "SubTotal:"
      Top             =   8565
      Width           =   2115
   End
   Begin VB.TextBox TXTFINAL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   300
      Left            =   15795
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   8565
      Width           =   1005
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   300
      Index           =   8
      Left            =   13680
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Text            =   "Descto.:"
      Top             =   8880
      Width           =   795
   End
   Begin VB.TextBox PORDESCU 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   300
      Left            =   14505
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   8880
      Width           =   1280
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
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
      Height          =   300
      Index           =   7
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Text            =   "Dolar Cotizado:"
      Top             =   8550
      Width           =   1650
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
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
      Height          =   300
      Index           =   6
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Text            =   "Cot.Actual Dolar:"
      Top             =   7920
      Width           =   1515
   End
   Begin VB.CommandButton Command14 
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
      Height          =   1560
      Left            =   12720
      Picture         =   "COTREPEYM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Aprueba, Graba"
      Top             =   7920
      Width           =   780
   End
   Begin VB.TextBox TXTTOTAL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   300
      Left            =   15795
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1005
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   300
      Index           =   4
      Left            =   13665
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Text            =   "Asistencia.Téc.Total:"
      Top             =   7920
      Width           =   2115
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2085
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   16700
      Begin VB.CommandButton BTNTIPOREP 
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
         Height          =   330
         Left            =   7560
         TabIndex        =   58
         Top             =   1680
         Width           =   180
      End
      Begin VB.TextBox TXTTIPOREP 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1920
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   3
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "M.O.Hora "
         Top             =   1680
         Width           =   1005
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   2
         Left            =   8205
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "Costo Rep."
         Top             =   1680
         Width           =   1155
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   1
         Left            =   10455
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "Tiempo M.O."
         Top             =   1680
         Width           =   1155
      End
      Begin VB.TextBox TXTVALORXHORA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   13720
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1005
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Ord."
         Top             =   1600
         Width           =   465
      End
      Begin VB.TextBox TXTORDEN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   580
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1600
         Width           =   465
      End
      Begin VB.CommandButton Command2 
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
         Height          =   300
         Left            =   14880
         MaskColor       =   &H0084962E&
         TabIndex        =   39
         ToolTipText     =   "Agregar"
         Top             =   1680
         Width           =   1350
      End
      Begin VB.TextBox TIEMPOMOB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   11600
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1005
      End
      Begin VB.TextBox TXTCOSTOREP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1005
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   9520
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   1005
      End
      Begin VB.TextBox TEXT21 
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   3
         Left            =   3500
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox TEXT21 
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   2
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Descripción de Tareas Cliente:"
         Top             =   240
         Width           =   3000
      End
      Begin VB.TextBox TEXT21 
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   1
         Left            =   11040
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Descripción de tareas E& M:"
         Top             =   240
         Width           =   2520
      End
      Begin VB.TextBox TEXT21 
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Falla Presentada:"
         Top             =   240
         Width           =   3000
      End
      Begin VB.CommandButton Command5 
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
         Height          =   345
         Left            =   16320
         MaskColor       =   &H0084962E&
         TabIndex        =   25
         ToolTipText     =   "Lista de Selección Datos Anteriromente Guardados"
         Top             =   1230
         Width           =   270
      End
      Begin VB.CommandButton Command4 
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
         Height          =   345
         Left            =   10800
         MaskColor       =   &H0084962E&
         TabIndex        =   24
         ToolTipText     =   "Tareas Valorizadas"
         Top             =   1230
         Width           =   270
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
         Height          =   345
         Left            =   5300
         MaskColor       =   &H0084962E&
         TabIndex        =   23
         ToolTipText     =   "Lista de Selección Datos Anteriromente Guardados"
         Top             =   1230
         Width           =   270
      End
      Begin VB.TextBox TXTDESCRIEYM 
         Appearance      =   0  'Flat
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
         Height          =   975
         Left            =   11060
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   600
         Width           =   5250
      End
      Begin VB.TextBox TXTTAREACTE 
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
         Height          =   975
         Left            =   5580
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   600
         Width           =   5205
      End
      Begin VB.TextBox TXTFALLAPRES 
         Appearance      =   0  'Flat
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
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   600
         Width           =   5205
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5280
         TabIndex        =   59
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   42
         Top             =   1680
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Accesorios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6015
         TabIndex        =   19
         Top             =   420
         Visible         =   0   'False
         Width           =   1245
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CODIGO - MODELO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2205
      Left            =   120
      TabIndex        =   7
      Top             =   80
      Width           =   16700
      Begin VB.TextBox TXTAPROBACION 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   80
         TabStop         =   0   'False
         Text            =   "APROBACIÓN"
         Top             =   315
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Historial"
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
         Left            =   15600
         Picture         =   "COTREPEYM.frx":030A
         TabIndex        =   64
         ToolTipText     =   "Muestra el Historial del Equipo Elegido"
         Top             =   680
         Width           =   850
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "Contacto:"
         Top             =   1830
         Width           =   850
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "R.Social:"
         Top             =   1440
         Width           =   850
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   5
         Left            =   15600
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "F.Cotiz:"
         Top             =   1440
         Width           =   795
      End
      Begin VB.TextBox TXTFECHCOT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   15600
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1005
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Defectos:"
         Top             =   315
         Width           =   1080
      End
      Begin VB.CommandButton BOTSELFORM 
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
         Height          =   330
         Index           =   5
         Left            =   15330
         TabIndex        =   29
         Top             =   675
         Width           =   180
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   14160
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Talón Nro:"
         Top             =   315
         Width           =   1140
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Accesorios:"
         Top             =   315
         Width           =   1080
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Equipo:"
         Top             =   315
         Width           =   795
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Nº Serie:"
         Top             =   1065
         Width           =   850
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   14160
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1005
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   0
         Left            =   14160
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "F.Alta:"
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label LBLCODIGO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1180
         TabIndex        =   56
         Top             =   315
         Width           =   1910
      End
      Begin VB.Label lblcontacto 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1250
         TabIndex        =   55
         Top             =   1830
         Width           =   3375
      End
      Begin VB.Label lblrasocli 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1250
         TabIndex        =   53
         Top             =   1455
         Width           =   3370
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Recepn: Fecha    y Hora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   6
         Left            =   5760
         TabIndex        =   35
         Top             =   195
         Width           =   1320
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7200
         TabIndex        =   34
         Top             =   315
         Width           =   2235
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3120
         TabIndex        =   33
         Top             =   315
         Width           =   1515
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   1485
         Index           =   0
         Left            =   9480
         TabIndex        =   30
         Top             =   675
         Width           =   4515
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   1485
         Left            =   4680
         TabIndex        =   11
         Top             =   675
         Width           =   4755
      End
      Begin VB.Label lblnserie 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1250
         TabIndex        =   10
         Top             =   1080
         Width           =   3380
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   14160
         TabIndex        =   9
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   330
         TabIndex        =   8
         Top             =   705
         Width           =   4300
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   3420
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Doble Click Editar/ Boton Derecho Eliminar -"
      Top             =   4440
      Width           =   16665
      _ExtentX        =   29395
      _ExtentY        =   6033
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
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
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   11160
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   9015
      Left            =   0
      TabIndex        =   1
      Top             =   -70
      Width           =   80
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "COTREPEYM.frx":0454
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "COTREPEYM.frx":0688
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   1452
      Left            =   120
      TabIndex        =   70
      ToolTipText     =   "Doble Click Editar/ Boton Derecho Eliminar -"
      Top             =   7920
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EFF1F1&
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
      Left            =   840
      ScaleHeight     =   315
      ScaleWidth      =   5175
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   5235
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   5200
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label MONEMIS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pesos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9120
      TabIndex        =   75
      ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
      Top             =   8230
      Width           =   1380
   End
   Begin VB.Label lbltotaccesorios 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   15795
      TabIndex        =   74
      ToolTipText     =   "Se Registra al Momento de La Impresión"
      Top             =   8250
      Width           =   1005
   End
   Begin VB.Label lblCOTIZADO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   10800
      TabIndex        =   63
      ToolTipText     =   "Se Registra al Momento de La Impresión"
      Top             =   8556
      Width           =   1032
   End
   Begin VB.Label TICAMBIS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   10680
      TabIndex        =   60
      Top             =   7920
      Width           =   1035
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   0
      Top             =   9500
      Width           =   750
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir Cotización (Solo Aprobada)"
      End
      Begin VB.Menu mnuImprimeIT 
         Caption         =   "Imprimir Informe Técnico"
      End
      Begin VB.Menu mnuAprobCot 
         Caption         =   "Aprobar Cotización"
      End
   End
   Begin VB.Menu mnuCl 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuCostos 
         Caption         =   "Tabla Costo Repuestos y Tiempo Mano de Obra"
      End
      Begin VB.Menu mnuConfigCotExcel 
         Caption         =   "Configuración Formulario Cotización"
      End
   End
   Begin VB.Menu mnuAccesDirect 
      Caption         =   "Accesos Directos"
   End
End
Attribute VB_Name = "COTREPEYM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$, ABUSQUE$
Public MODIFICATIEMMOBRA%
Dim SALIRX%
Sub ACOMODARMSF()
     MSF1.RowHeight(MSF1.Row) = Label7.Height + 100
     MSF1.TextMatrix(MSF1.Row, 2) = Label7
End Sub


Sub BLANFORMU()
    LBLCODIGO = ""
    Label3 = ""
    Label22 = ""
    lblnserie = ""
    lblrasocli = ""
    lblcontacto = ""
    Label28 = ""
    TEXT21(3) = ""
    TXTFALLAPRES = ""
    TXTTAREACTE = ""
    TXTDESCRIEYM = ""
    Label16 = ""
    Label2(0) = ""
    Text3 = ""
    TXTFECHCOT = ""
    TXTTOTAL = ""
    TXTFINAL = ""
    PORDESCU = ""
    lblCOTIZADO = ""
    TICAMBIS = ""
    TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")

End Sub

Sub CREATABLA_ACCESORIOSCOTIZ()
   Call CREACAMPO("OFABRIC", "ACCESORIOSCOTIZ", "NumRepa", 4, 0, 1)
   Call CREACAMPO("OFABRIC", "ACCESORIOSCOTIZ", "Cod_Inte", 3, 0)
   Call CREACAMPO("OFABRIC", "ACCESORIOSCOTIZ", "CODIGO", 10, 24)
   Call CREACAMPO("OFABRIC", "ACCESORIOSCOTIZ", "IDENLOTE", 10, 24)
   Call CREACAMPO("OFABRIC", "ACCESORIOSCOTIZ", "DESCRIPCION", 10, 128)
   Call CREACAMPO("OFABRIC", "ACCESORIOSCOTIZ", "CANTIDAD", 7, 0)
   Call CREACAMPO("OFABRIC", "ACCESORIOSCOTIZ", "PrUn_Neto", 7, 0)
   Call CREACAMPO("OFABRIC", "ACCESORIOSCOTIZ", "MarBorra", 3, 0)
   Call CREACAMPO("OFABRIC", "ACCESORIOSCOTIZ", "Nuord_Item", 3, 0, 1)
   Call CREACAMPO("OFABRIC", "ACCESORIOSCOTIZ", "PORDESC", 7, 0)

End Sub

Private Sub BOTSELFORM_Click(Index As Integer)
    Label25_DblClick
End Sub

Private Sub BTNTIPOREP_Click()
   X = LISTASELECGENERAL("TIPOREP")
   Me.Label4 = XVALO(RESP_ZELE_VECT(1))
'   Me.TXTTIPOREP = TRIM$(RESP_ZELE_VECT(2))
   
End Sub


Private Sub Command1_Click()
   Campos$ = "FALLA PRESENTADA/A160"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "FALLAS PRESENTADAS"
   SQLX$ = "SELECT DISTINCT FALLAPRE FROM COTIEYM WITH(NOLOCK) WHERE FALLAPRE <> '' ORDER BY FALLAPRE"
   Set RS1 = READSET(SQLX$)
   Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2)
   RS1.Close
   Set RS1 = Nothing
   '
   FRMZELECHO.Width = 16800
   FRMZELECHO.Show 1
   If RESP_ZELE_VECT(1) <> "" Then
     TXTFALLAPRES = RESP_ZELE_VECT(1)
   End If
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    ACUMULAMOBRA = 0
    ACUMULAREPU = 0
    SALIRX% = 0
    NROREPA& = XVALO(Label25)
    If TXTAPROBACION.Visible = False Then
        ESTADO% = XVALO(VALOBADA("REPARA", "STAT_APROBACION", "NUMREPA = " & NROREPA&, "NOMESS"))
        If ESTADO% = 1 Then
               OPX% = COMALTER%("Cotización  Aprobada, Si Graba La Misma Queda Pendiente de Aprobación\\Cancelar\Volver a Grabar")
               If OPX% < 2 Then
                 SALIRX% = 1
                 GoTo 99
               End If
        End If
    End If
    
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    CONDI$ = "NUMREPA = " & NROREPA&
    HORAVOLIRIZADA = XVALO(TXTVALORXHORA)
    If HORAVOLIRIZADA < 0.005 Then
      Call COMUNI("Falta Ingresar Valor Horario de Mano de Obra")
      GoTo 99
    End If
    Call ACTUREGISTRO("COTIEYM", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
    FECOT$ = VALOBADA("COTIEYM", "FECHCOT", CONDI$, "NOMESS") 'SI TIENE FECHA SE MANTIENE LA MISMA
    If FECOT$ = "" Then FECOT$ = FECHATEX$(HOY(HO$))
    
     strSQL = "Select * from COTIEYM WHERE " + CONDI$
     Dim rs As ADODB.Recordset
     Set rs = New ADODB.Recordset
25   On Error Resume Next
     rs.Open (strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     With rs

        For REG& = 1 To MSF1.Rows - 1
            ORDENX = MSF1.TextMatrix(REG&, 1)
            FALLAPREX$ = MSF1.TextMatrix(REG&, 2)
            DESTAREAX$ = MSF1.TextMatrix(REG&, 3)
            DESTAREA2X$ = MSF1.TextMatrix(REG&, 4)
            COSTOREP = XVALO(MSF1.TextMatrix(REG&, 5))
            COSTOMOB = XVALO(MSF1.TextMatrix(REG&, 6))
            .AddNew
            !numrepa = NROREPA&
            !ORDEN = ORDENX
            !FECHCOT = FECOT$
            !FALLAPRE = Left$(FALLAPREX$, 512)
            !DESTAREA = Left$(DESTAREAX$, 512)
            !DESTAREA2 = Left$(DESTAREA2X$, 512)
            !COSTOMAT = COSTOREP
            !COSTOMOB = COSTOMOB
            !CodiEmpr = CodiEmp%
            !MARBORRA = 0
            !TIPOREPA = XVALO(Label4)
            !PORDESCU = XVALO(PORDESCU)
            !VALORMOB = XVALO(Me.TXTVALORXHORA)
            !MONEEMIS = MONEMI%
            ACUMULAMOBRA = ACUMULAMOBRA + (COSTOMOB * HORAVOLIRIZADA)
            ACUMULAREPU = ACUMULAREPU + COSTOREP
            !COTIZA_DOLAR = XVALO(lblCOTIZADO)
            .Update
         Next REG&
     End With
     rs.Close
     Set rs = Nothing
     '
'    TotalRepConAccesorios = XVALO(lbltotaccesorios) + ACUMULAREPU 'le agrego los accesorios tambien como repuestos
     SQLX$ = "SELECT * FROM REPARA"
     SQLX$ = SQLX$ + " WHERE NUMREPA = " & NROREPA&
     Dim RSX As ADODB.Recordset
     Set RSX = New ADODB.Recordset
     RSX.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     With RSX
       !Tot_Mano_Obra = XVALO(ACUMULAMOBRA)
       !Tot_Repuestos = XVALO(ACUMULAREPU)
       !MONEEMIS = MONEMI%
       !DESCTO1 = XVALO(PORDESCU)
       !STAT_APROBACION = 0
       !CondPag = XVALO(TEXT211)
       .Update
     End With

     '
     CONDI$ = CONDI$ + " AND  MARBORRA = 1  "
     Call BORRAREGISTROS("COTIEYM", CONDI$, REGAF&, "NOMESS")
     '
     
ERROR_GUARDAR:
    If Err <> 0 Then
       MsgBox Err.Description
       
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
    End If
99  Command14.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    ORD% = XVALO(TXTORDEN)
    If ORD% > 0 Then 'si es edicion
       ORD% = ORD%
    Else
       ORD% = MSF1.Rows 'si esta agregando
       MSF1.Rows = MSF1.Rows + 1
    End If
    If Len(Me.TXTFALLAPRES) > Len(Me.TXTTAREACTE) Then
       DESCRI$ = TXTFALLAPRES
    Else
       DESCRI$ = TXTTAREACTE
    End If
    If Len(DESCRI$) < Len(Me.TXTDESCRIEYM) Then DESCRI$ = TXTDESCRIEYM
    Label7 = DESCRI$
    MSF1.RowHeight(MSF1.Rows - 1) = Label7.Height + 100

    MSF1.TextMatrix(ORD%, 1) = ORD%
    MSF1.TextMatrix(ORD%, 2) = RTrim$(LTrim$(Me.TXTFALLAPRES))
    MSF1.TextMatrix(ORD%, 3) = RTrim$(LTrim$(Me.TXTTAREACTE))
    MSF1.TextMatrix(ORD%, 4) = RTrim$(LTrim$(Me.TXTDESCRIEYM))
    MSF1.TextMatrix(ORD%, 5) = TRIM$(FORMATNUM$(XVALO(Me.TXTCOSTOREP), "F12.2"))
    MSF1.TextMatrix(ORD%, 6) = TRIM$(FORMATNUM$(XVALO(Me.TIEMPOMOB), "F12.2"))
    MSF1.TextMatrix(i&, 7) = XVALO(Label2(1).Caption)
    Call LIMPIA_EDICION
    TXTTOTAL = FORMATNUM$(CALCULARTOTAL#, "F12.2")
99  Command2.Enabled = True
End Sub
Function CALCULARTOTAL#()
    VALORHORA# = XVALO(TXTVALORXHORA)
    ACURESULTADO# = 0
    For REG& = 1 To MSF1.Rows - 1
            COSTOREP = XVALO(MSF1.TextMatrix(REG&, 5))
            COSTOMOB = XVALO(MSF1.TextMatrix(REG&, 6))
            Resultado = (VALORHORA# * COSTOMOB) + COSTOREP
            ACURESULTADO# = ACURESULTADO# + Resultado
    Next REG&
    CALCULARTOTAL# = ACURESULTADO#
End Function
Sub LIMPIA_EDICION()
    TXTFALLAPRES.TEXT = ""
    TXTTAREACTE.TEXT = ""
    TXTDESCRIEYM.TEXT = ""
    TXTCOSTOREP.TEXT = ""
    TIEMPOMOB.TEXT = ""
    Label2(1) = ""
    TXTORDEN = ""
End Sub

Private Sub Command21_Click(Index As Integer)
    Call SELECHO("CONPAG")
    VDEVU$ = TRIM$(VALACT1$("CONPAG"))
    If VDEVU$ <> "" Then
       TEXT211.TEXT = VDEVU$
    End If

End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    CONUSE07.Label7 = lblnserie
    CONUSE07.Show 1
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Campos$ = "CÓDIGO/F9;DESCRIPCIÓN/A128;COSTO/A0;TIEMPMOB/A0;EDITA/A0"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "TAREAS"
   SQLX$ = "SELECT CODINT,REFEREN,COSTOREP,HORAMOB,EDITABLE FROM REPACOS WITH(NOLOCK) WHERE REFEREN <> '' ORDER BY REFEREN"
   Set RS1 = READSET(SQLX$)
   Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2)
   RS1.Close
   Set RS1 = Nothing
   '
   FRMZELECHO.Width = 14800
   FRMZELECHO.Show 1
   Text5 = XVALO(RESP_ZELE_VECT(1))
   TXTTAREACTE = RESP_ZELE_VECT(2)
   Me.TXTCOSTOREP = RESP_ZELE_VECT(3)
   Me.TIEMPOMOB = RESP_ZELE_VECT(4)
   EDITA% = XVALO(RESP_ZELE_VECT(5))
   If EDITA% > 0 Then 'EN LA CONFIGURACION SE PUEDE MARCAR SI EL COSTO ES EDITABLE EN PANTALLA, SI ES ASI HABILITA EDICION EN CASILLERO.
      TXTCOSTOREP.Locked = False
      TXTCOSTOREP.BackColor = vbWhite
      
   Else
      TXTCOSTOREP.BackColor = &HE0E0E0
      TXTCOSTOREP.Locked = True
   End If
End Sub


Private Sub Command5_Click()
   Command5.Enabled = True
   Campos$ = "TAREA E&M/A160"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "TAREA E&M"
   SQLX$ = "SELECT DISTINCT DESTAREA2 FROM COTIEYM WITH(NOLOCK) WHERE DESTAREA2 <> '' ORDER BY DESTAREA2"
   Set RS1 = READSET(SQLX$)
   Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2)
   RS1.Close
   Set RS1 = Nothing
   '
   FRMZELECHO.Width = 16800
   FRMZELECHO.Show 1
   If RESP_ZELE_VECT(1) <> "" Then
     TXTDESCRIEYM = RESP_ZELE_VECT(1)
   End If
99 Command5.Enabled = True



End Sub




Private Sub Command8_Click()
   On Error Resume Next
   BUSCA_DES.TXTBUSCAR.SetFocus
   On Error GoTo 0

   BUSCA_DES.Show 1
   On Error Resume Next
   Command3.SetFocus
   On Error GoTo 0
   
End Sub



Private Sub mnuArchivo_Click()

End Sub

Private Sub Image2_Click()

End Sub

Private Sub Form_Load()
   Call CREATABLA_COTIEYM
   Call CREATABLA_ACCESORIOSCOTIZ
   Call CREATABLA_REPACOS
   Call CARGA_ENCABEZADO(Me.MSF1, "ORD/f3;Falla Presentada/A48;Descrip Tarea/A48;Descripción E&M/A48;Costo/F9.2;M.O./F6.2;ID/a0", Me)
   Call CARGA_ENCABEZADO(MSF2, "Código/A16;Descripción/A48;Importe/f12.2", Me)

10 If MONEMI% <> 2 Then
      VUELTA% = VUELTA% + 1
      If VUELTA% > 10 Then MsgBox "Alcanzó los 10 ": GoTo 20
      MONEMIS_DblClick
      GoTo 10
   End If

20 TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")

   Me.TXTVALORXHORA = TRIM$(FORMATNUM$(XVALO(Control$("REPARA", "VALHORA")), "F12.2"))
End Sub
Sub CREATABLA_REPACOS()
   'TABLA PARA REGISTRAR LAS TAREAS VALORIZADAS PARA LA REPARACION DE EYM LOS VALORES SON EN DOLARES
   Call CREACAMPO("OFABRIC", "REPACOS", "Referen", 10, 512, 1)
   If ValidaExisteCampo%("REPACOS_ID", "REPACOS") < 1 Then
    SQLX$ = "ALTER TABLE REPACOS ADD REPACOS_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("OFABRIC", "REPACOS", "COSTOREP", 7, 0)
   Call CREACAMPO("OFABRIC", "REPACOS", "HORAMOB", 7, 0)
   Call CREACAMPO("OFABRIC", "REPACOS", "EDITABLE", 3, 0) '(0 NO EDITABLE) (1 SI EDITABLE)
   
   Call CREACAMPO("OFABRIC", "REPACOS", "MARBORRA", 3, 0, 1)
   Call CREACAMPO("OFABRIC", "REPACOS", "CODINT", 4, 0, 1)
   
End Sub
Sub CREATABLA_COTIEYM()

   Call CREACAMPO("OFABRIC", "COTIEYM", "Referen", 10, 128)
   If ValidaExisteCampo%("COTIEYM_ID", "COTIEYM") < 1 Then
    SQLX$ = "ALTER TABLE COTIEYM ADD COTIEYM_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If

   Call CREACAMPO("OFABRIC", "COTIEYM", "NumRepa", 4, 0, 1)
   Call CREACAMPO("OFABRIC", "COTIEYM", "FECHCOT", 8, 0)
   Call CREACAMPO("OFABRIC", "COTIEYM", "FEAPROCOT", 8, 0)
   Call CREACAMPO("OFABRIC", "COTIEYM", "FALLAPRE", 10, 512)
   Call CREACAMPO("OFABRIC", "COTIEYM", "DESTAREA", 10, 512)
   Call CREACAMPO("OFABRIC", "COTIEYM", "DESTAREA2", 10, 512)
   Call CREACAMPO("OFABRIC", "COTIEYM", "COSTOMAT", 7, 0)
   Call CREACAMPO("OFABRIC", "COTIEYM", "COSTOMOB", 7, 0)
   Call CREACAMPO("OFABRIC", "COTIEYM", "VALORMOB", 7, 0) 'VALOR MANO DE OBRA
   
   Call CREACAMPO("OFABRIC", "COTIEYM", "OBSERVACIONES", 10, 512)
   Call CREACAMPO("OFABRIC", "COTIEYM", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("OFABRIC", "COTIEYM", "MARBORRA", 3, 0, 1)
   Call CREACAMPO("OFABRIC", "COTIEYM", "ORDEN", 3, 0)
   Call CREACAMPO("OFABRIC", "COTIEYM", "TIPOREPA", 3, 0)
   Call CREACAMPO("OFABRIC", "COTIEYM", "Nume_Cli", 4, 0, 1)
   Call CREACAMPO("OFABRIC", "COTIEYM", "Raso_Cli", 10, 128)
   Call CREACAMPO("OFABRIC", "COTIEYM", "Status", 3, 0)
   Call CREACAMPO("OFABRIC", "COTIEYM", "NumeroSerie", 10, 32)
   Call CREACAMPO("OFABRIC", "COTIEYM", "CODMODELO", 10, 32)
   Call CREACAMPO("OFABRIC", "COTIEYM", "Descripcion", 10, 128)
   Call CREACAMPO("OFABRIC", "COTIEYM", "COTIZA_DOLAR", 7, 0)
   Call CREACAMPO("OFABRIC", "COTIEYM", "PORDESCU", 10, 32)
   

   


End Sub

Private Sub Form_Resize()
   On Error Resume Next
'   Me.MSF1.Width = Me.Width - 300
'   Me.MSF1.Height = Me.Height - 1600
   On Error GoTo 0

End Sub

Private Sub Label25_Change()
    MSF1.Rows = 1
    MSF2.Rows = 1
    MSF2.Rows = 2
    Call BLANFORMU
    NUREP& = XVALO(Label25)
    CONDI$ = "NUMREPA = " & NUREP& & ""
    SQLX$ = "SELECT * FROM REPARA WITH(NOLOCK) WHERE " & CONDI$
    Set rs = READSET(SQLX$)
    With rs
      If Not .EOF Then
         NCLIEN = XVALO(!Nume_Cli)
         lblnserie = SAFETEXT$(!NUMEROSERIE)
         Label22 = SAFETEXT$(!Descripcion)
         Label2(0) = SAFETEXT$(!DEFEC_INFOR)
         Label28 = SAFETEXT$(!ACCESORIOS)
         Label28 = SAFETEXT$(!ACCESORIOS)
         lblrasocli = SAFETEXT$(!Raso_Cli)
         lblcontacto = SAFETEXT$(!Ctac_Cli)
         LBLCODIGO = CODEXT$(XVALO(!COD_INTE))
         FREP = SAFETEXT$(!FECHRECEP)
         Label16 = Format(FREP, "dd/mm/yy - hh:mm")
         MON% = XVALO(!MONEEMIS)
         GARANTIA% = XVALO(!STATUS_GAR)
         If GARANTIA% > 0 Then
            Label3 = "EN GARANTÍA"
         End If
         STATAPROB% = XVALO(!STAT_APROBACION)
         CPAG% = XVALO(!CondPag)
         If CPAG% > 0 Then
              TEXT211.TEXT = Str$(CPAG%)
           Else
              TEXT211.TEXT = Str$(CPAGCLI%(NCLIEN))
        End If

       Else
         lblnserie = ""
         Label22 = ""
'         label2 = ""
         Label28 = ""
         Label16 = ""
         
      End If
    End With
    rs.Close
    Set rs = Nothing '
    '
    NROREPA& = XVALO(Label25)
    SQLX$ = "SELECT * FROM COTIEYM WITH(NOLOCK) WHERE NUMREPA = " & NROREPA&
    SQLX$ = SQLX$ + " ORDER BY ORDEN"
    Set rs = READSET(SQLX$)
    REG& = 0
    With rs
      Do While Not .EOF
        REG& = REG& + 1
        If REG& = 1 Then
           FE% = CVINT(!FECHCOT)
           TXTFECHCOT = FECHATEX$(FE%)
           Label4 = XVALO(!TIPOREPA)
           lblCOTIZADO = TRIM$(FORMATNUM$(XVALO(!COTIZA_DOLAR), "F12.4"))
           DESXU$ = SAFETEXT$(!PORDESCU)
           VALORHORA = XVALO(!VALORMOB)
           If VALORHORA > 0 Then
             TXTVALORXHORA = FORMATNUM$(XVALO(!VALORMOB), "F12.2")
           End If
        End If
        
        MSF1.Rows = REG& + 1
        MSF1.TextMatrix(REG&, 1) = XVALO(!ORDEN)
        MSF1.TextMatrix(REG&, 2) = SAFETEXT$(!FALLAPRE)
        MSF1.TextMatrix(REG&, 3) = SAFETEXT$(!DESTAREA)
        MSF1.TextMatrix(REG&, 4) = SAFETEXT$(!DESTAREA2)
        MSF1.TextMatrix(REG&, 5) = TRIM$(FORMATNUM$(XVALO(!COSTOMAT), "F12.2"))
        MSF1.TextMatrix(REG&, 6) = TRIM$(FORMATNUM$(XVALO(!COSTOMOB), "F12.2"))
        MSF1.TextMatrix(REG&, 7) = XVALO(!COTIEYM_ID)
        .MoveNext
      Loop
    End With
    REG& = 0
    NLISTA% = XVALO(Control$("REPARA", "LIPREREP"))
    Campos$ = "Cod_Inte,CODIGO,DESCRIPCION,PrUn_Neto"
    SQLX$ = "SELECT " & Campos$
    SQLX$ = SQLX$ + " FROM ACCESORIOSCOTIZ WITH(NOLOCK) WHERE NUMREPA = " & XVALO(Label25)
    Set RS1 = READSET(SQLX$)
    With RS1
      Do While Not .EOF
         REG& = REG& + 1
         MSF2.Rows = REG& + 1
         CI1% = XVALO(!COD_INTE)
         MSF2.TextMatrix(REG&, 1) = SAFETEXT$(!Codigo)
         MSF2.TextMatrix(REG&, 2) = SAFETEXT$(!Descripcion)
         If XVALO(!PrUn_Neto) > 0.005 Then
            MSF2.TextMatrix(REG&, 3) = ROUND(XVALO(!PrUn_Neto), 2)
         Else
            PRELISTX = ROUND(PRELISTA#(TRIM$(Str$(NLISTA%)), CI1%), 2)
            MSF2.TextMatrix(REG&, 3) = PRELISTX
            If XVALO(PRELISTX) < 0.005 Then Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF2, vbYellow, REG&)

         End If
'         MSF2.TextMatrix(REG&, 4) = SAFETEXT$(!IDENLOTE)
            
        .MoveNext
      Loop
    End With
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    lbltotaccesorios = FORMATNUM$(SUMA_COLUMNA_SELECCION(MSF2, 3), "F12.2")


   If MON% < 1 Then MON% = 2
10 If MONEMI% <> MON% Then
      MONEMIS_DblClick
      GoTo 10
   End If

   
   TICAMBIS = FORMATNUM$(TICAMONE(MON%), "F9.4")

    '*****************************************************'
    
    Me.TXTTOTAL = FORMATNUM$(CALCULARTOTAL#, "F12.2")
    If TRIM$(DESXU$) <> "" Then
       PORDESCU = DESXU$
    Else
       PORDESCU = DESCUCLI$(NCLIEN)
    End If
    
    If STATAPROB% > 0 Then
       Call COMUNI("Atención!! La Cotización Activa Está Aprobada\Si la Misma se Vuelve a Grabar\Se Deberá Volver a Aprobar")
    End If
End Sub

Private Sub Label25_DblClick()
10  NroRep& = MREPARA09.LISTAREPARACION(" STATUS < 5")
    If NroRep& > 0 Then
      Label25 = NroRep&
    End If

End Sub


Private Sub Label4_Change()
      SQLX$ = "SELECT CAMPO2 FROM  DATASQL  WITH(NOLOCK) WHERE LISTA ='TIPOREP' AND CONVERT(int, campo1)= " & XVALO(Label4)
      Set rs = READSET(SQLX$)
      TXTTIPOREP = SAFETEXT$(rs!CAMPO2)
      rs.Close
      Set rs = Nothing

End Sub

Private Sub lbltotaccesorios_Change()
  TXTFINAL = FORMATNUM$(XVALO(TXTTOTAL) + XVALO(lbltotaccesorios), "f12.2")
End Sub

Sub SUMARTOTALCONACCESORIOS()
End Sub

Private Sub mnuAprobCot_Click()
    NroRep& = MREPARA09.LISTAREPARACION(" STAT_APROBACION <= 0 or STAT_APROBACION IS NULL")
    If NroRep& > 0 Then
     If DERACCE%("Aprobación de Cotización de Asistencia Técnica", "APCOTREP") >= 2 Then
       Label25 = NroRep&
       OPX% = COMALTER%("Desea Aprobar la Cotización Activa\\Cancelar\Aprobar")
       If OPX% = 2 Then
          TXTAPROBACION.Visible = True
          Call ACTUALIZACOTIZACION(NroRep&)
          Call ACTUREGISTRO("REPARA", "STAT_APROBACION", "NUMREPA = " & NroRep&, 1, REGAF&, "NOMESS")
          EstadoActual% = XVALO(VALOBADA("REPARA", "STATUS", "NUMREPA = " & NroRep&, "NOMESS"))
          If EstadoActual% < 2 Then
             Call ACTUREGISTRO("REPARA", "status", "NUMREPA = " & NroRep&, 2, REGAF&, "NOMESS")
          End If
          TXTAPROBACION.Visible = False

          If REGAF& > 0 Then
             Call COMUNI("Aprobación Realizada")
          Else
             Call COMUNI("Se Produjo un Error al Grabar, No se Realizó la Transacción")
          End If
       End If
      End If
    End If
End Sub

Sub ACTUALIZACOTIZACION(NROREPA&)
    NROREPA& = XVALO(Label25)
    If NROREPA& < 1 Then
       Call COMUNI("No Hay Ninguna Reparación Activa")
       Exit Sub
    End If
    '
    Call Command14_Click     'BOTON GRABACION
    
    'EL VALOR DEL DOLAR DE COTIZACION SE GUARDA CUANDO SE IMPRIME (ENTENDIENDOSE QUE ES CUANDO SE LE ENVIA AL CLIENTE9
    ACTUCOTIZACION% = 0
    If XVALO(lblCOTIZADO) > 0 Then
       If XVALO(lblCOTIZADO) <> XVALO(TICAMBIS) Then
          OPX% = COMALTER%("Esta Cotización Yá Fué Generada Con la Cotización del Dolar a: " & lblCOTIZADO & "\\Guardar Sin Actualizar Cotización\Guardar con Cotización Actual")
          If OPX% > 1 Then
             lblCOTIZADO = TRIM$(FORMATNUM$(XVALO(TICAMBIS), "F12.2"))
          End If
       End If
    Else
       lblCOTIZADO = TICAMBIS
    End If
    CONDI$ = "NUMREPA = " & NROREPA&
    Call ACTUREGISTRO("COTIEYM", "COTIZA_DOLAR", CONDI$, FORMATNUM$(XVALO(lblCOTIZADO), "F12.2"), REGAF&, "NOMESS")
    Call ACTUREGISTRO("REPARA", "TIPOCAM", CONDI$, FORMATNUM$(XVALO(lblCOTIZADO), "F12.2"), REGAF&, "NOMESS")

     '-----------------------------------
     NUMESER$ = TRIM$(lblnserie)
     CUERPO$ = "Cotización de Asist.Técnica" & vbCrLf
     CUERPO$ = CUERPO$ + "Número de Asist.Técnica: " & NROREPA& & vbCrLf
     CUERPO$ = CUERPO$ + "Cliente: " & lblrasocli & vbCrLf
     CUERPO$ = CUERPO$ + "Importe u$s: " & TRIM$(TXTFINAL.TEXT) & " ($" & TRIM$(FORMATNUM$(XVALO(TXTFINAL) * XVALO(TICAMBIS), "F12.2")) & ")" & vbCrLf
     CUERPO$ = CUERPO$ + " (Desc.Incluido: " & PORDESCU.TEXT & " %)" & vbCrLf
     If XVALO(TXTTOTALCONACCESORIOS) > 0 Then CUERPO$ = CUERPO$ + "+ Accesorios: " & TRIM$(FORMATNUM$(XVALO(TXTTOTALCONACCESORIOS), "F12.2")) & vbCrLf
     CUERPO$ = CUERPO$ + "Cotiz. Actual U$S: " & TRIM$(FORMATNUM$(TICAMONE(2), "F12.2")) & vbCrLf
     CUERPO$ = CUERPO$ + "Presup Cotizado C/U$S: " & TRIM$(lblCOTIZADO)
     
     Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")
     '-----------------------------------

End Sub
Private Sub mnuConfigCotExcel_Click()
  FSCOTEXCEL.Show 1
End Sub

Private Sub mnuCostos_Click()
    Unload VENTABNEW
    VENTANA.Inicializar = 0
    VENTANA.CampEditables = "2;3;4;5"
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = ""
'    CAMPOS$ = "FALLA PRESENTADA/A32;DESCR.TAREA/A32;DESC.TAREA E&M/A32;ID/A0;"
    Campos$ = "CÓD.INT/F8;DESCRIPCIÓN/A84;COSTO REP.U$S/F12.2;TIEMPO M.O./F12.2;Edit./F5;ID/A0;"
    VENTANA.Campox = Campos$
    VENTANA.TITULO = "Descripción de tareas"
    VENTANA.Inicializar = 1
    
    '****************************************************************
    Dim obj As New RECORXET
    SQLX$ = "SELECT CODINT,REFEREN,COSTOREP,HORAMOB,EDITABLE FROM REPACOS WITH(NOLOCK) WHERE REFEREN <> '' ORDER BY REFEREN"
    Set RS1 = READSET(SQLX$)
    Call Carga_MSF_Recordset(RS1, Campos$, VENTABNEW.MSF)
    RS1.Close
    Set RS1 = Nothing
    
30  VENTABNEW.Show 1
    Dim MSF As msFlexGrid
    Set MSF = VENTABNEW.MSF
        '
    If VENTABNEW.APROBADO% < 1 Then Exit Sub
    For REG& = 1 To MSF.Rows - 1
       Dato$ = REPLACAR(MSF.TextMatrix(REG&, 2), ",", " ")
       Dato$ = REPLACAR(MSF.TextMatrix(REG&, 2), "'", " ")
       MSF.TextMatrix(REG&, 2) = Dato$
    Next REG&
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans

    SQLX$ = "SELECT * FROM REPACOS"
    SQLX$ = SQLX$ + " WHERE REPACOS_ID < 1"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText



    With rs
          For REG& = 1 To MSF.Rows - 1
           If XVALO(MSF.TextMatrix(REG&, 1)) < 1 Then
                NUEVOCI& = 1 + XVALO(VALOBADA("REPACOS", "MAX(CODINT)", "REPACOS_ID > 0", "NOMESS"))
                SQLX$ = "INSERT INTO REPACOS "
                SQLX$ = SQLX$ + " (CODINT,REFEREN,COSTOREP,HORAMOB,EDITABLE,CODIEMPR,MARBORRA)"
                SQLX$ = SQLX$ + " VALUES(" & NUEVOCI&
                SQLX$ = SQLX$ + " ,'" & MSF.TextMatrix(REG&, 2) & "'"
                SQLX$ = SQLX$ + " ,'" & FORMATNUM$(XVALO(MSF.TextMatrix(REG&, 3)), "F12.2") & "'"
                SQLX$ = SQLX$ + " ,'" & FORMATNUM$(XVALO(MSF.TextMatrix(REG&, 4)), "F12.2") & "'"
                SQLX$ = SQLX$ + " ," & XVALO(MSF.TextMatrix(REG&, 5))
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ," & CodiEmp%
                SQLX$ = SQLX$ + ")"
                FLEXCONN.Execute (SQLX$)
           Else
                CI1 = XVALO(MSF.TextMatrix(REG&, 1))
                SQLX$ = " UPDATE REPACOS SET REFEREN = '" & MSF.TextMatrix(REG&, 2) & "'"
                SQLX$ = SQLX$ + " ,COSTOREP = '" & FORMATNUM$(XVALO(MSF.TextMatrix(REG&, 3)), "F12.2") & "'"
                SQLX$ = SQLX$ + " ,HORAMOB =  '" & FORMATNUM$(XVALO(MSF.TextMatrix(REG&, 4)), "F12.2") & "'"
                SQLX$ = SQLX$ + " ,EDITABLE = " & XVALO(MSF.TextMatrix(REG&, 5))
                SQLX$ = SQLX$ + " ,CODIEMPR = " & CodiEmp%
                SQLX$ = SQLX$ + " ,MARBORRA = 0"
                SQLX$ = SQLX$ + " WHERE CODINT  = " & CI1
                FLEXCONN.Execute (SQLX$)
           End If
          Next REG&
    End With
'
'    CONDI$ = CONDI$ + " AND  MARBORRA = 1  "
'    Call BORRAREGISTROS("REPACOS", CONDI$, REGAF&, "NOMESS")
    
    Set VENTANA = Nothing
        
ERROR_GUARDAR:
    If Err <> 0 Then
       MsgBox Err.Description
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
    End If
    Unload VENTABNEW

    
End Sub


Private Sub mnuImprimeIT_Click()
    NROREPA& = XVALO(Label25)
    If NROREPA& < 1 Then
       Call COMUNI("No Hay Ninguna Reparación Activa")
       Exit Sub
    End If
    '
    NSERIE$ = TRIM$(lblnserie)
    If NSERIE$ = "" Then
       Call COMUNI("Falta Seleccionar Un Número de Serie")
       Exit Sub
    End If
    '
    PDES$ = Control$(IDENTER$, "PRICOTEX") 'la configuracion ya esta.
    If PDES$ = "" Then PDES$ = Control$("*", "PRICOTEX")
    '
    RUTAEXCEL$ = Control("COTIEXCE", "PRININTE")
    If EXISTE(Control("COTIEXCE", "PRININTE")) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Informe Técnico")
       Exit Sub
    End If
    '

    With VENTABNEW.MSF
       .Rows = 11: .Cols = 4
       .TextMatrix(1, 1) = Me.lblcontacto
       .TextMatrix(2, 1) = Me.lblrasocli
       .TextMatrix(3, 1) = Me.Label25
       .TextMatrix(4, 1) = Me.LBLCODIGO
       .TextMatrix(5, 1) = Me.lblnserie
       .TextMatrix(6, 1) = Me.Label2(0) 'acA DEBERIA A TENER QUE IR FALLA INDICADA POR EL CLIENTE
       .TextMatrix(7, 1) = TRIM$(FORMATNUM$(XVALO(Me.TXTTOTAL) * XVALO(lblCOTIZADO), "F12.2")) 'LO PASO A $
       .TextMatrix(8, 1) = TRIM$(FORMATNUM$(XVALO(Me.TXTIMPODESC) * XVALO(lblCOTIZADO), "F12.2"))
       .TextMatrix(9, 1) = ""
       .TextMatrix(10, 1) = ""
       'AGREGA DETALLE DE REPARACION
       RetengoValorMSF1% = MSF1.Rows
       If MSF1.Rows < 12 Then
          MSF1.Rows = 12
       End If
       
       For J& = 1 To Me.MSF1.Rows - 1
              i& = J& + 10  'PARA QUE ARRANQUE A PARTIR DE LA FILA 11
              .Rows = i& + 1
              .TextMatrix(i&, 1) = Me.MSF1.TextMatrix(J&, 1)
              .TextMatrix(i&, 2) = Me.MSF1.TextMatrix(J&, 2)
              .TextMatrix(i&, 3) = Me.MSF1.TextMatrix(J&, 3)
        Next J&
        
    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
        
    Call Param_Print_Excel(RUTAEXCEL$, "INFOTECNICO", VENTABNEW.MSF, PDES$, "1;2;3", 1, 4, "", 1, 0)                'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW
    MSF1.Rows = RetengoValorMSF1%

End Sub

Private Sub mnuImprimir_Click()

    NROREPA& = XVALO(Label25)
    If NROREPA& < 1 Then
       Call COMUNI("No Hay Ninguna Reparación Activa")
       Exit Sub
    End If
    
    
    ESTADO% = XVALO(VALOBADA("REPARA", "STAT_APROBACION", "NUMREPA = " & NROREPA&, "NOMESS"))
    If ESTADO% < 1 Then
         Call COMUNI("Imposible Imprmir - Cotización No Aprobada")
         Exit Sub
'        If DERACCE%("Aprobación de Cotización de Asistencia Técnica", "APCOTREP") < 2 Then
'           Call COMUNI("Cotización sin Aprobar")
'           Exit Sub
'        Else
'           OPX% = COMALTER%("Cotización sin Aprobar\\Cancelar\Imprimir")
'           If OPX% < 2 Then Exit Sub
'
'        End If
    End If
     
    
    '
    NSERIE$ = TRIM$(lblnserie)
    If NSERIE$ = "" Then
       Call COMUNI("Falta Seleccionar Un Número de Serie")
       Exit Sub
    End If
    '
    PDES$ = Control$(IDENTER$, "PRICOTEX") 'la configuracion ya esta.
    If PDES$ = "" Then PDES$ = Control$("*", "PRICOTEX")
    '
    RUTAEXCEL$ = Control("COTIEXCE", "PRINCOTI")
    If EXISTE(Control("COTIEXCE", "PRINCOTI")) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Cotización")
       Exit Sub
    End If
    '
'    Call Command14_Click   'BOTON GRABACION
'    If SALIRX% = 1 Then
'       SALIRX% = 0
'      Exit Sub
'    End If
'
'    'EL VALOR DEL DOLAR DE COTIZACION SE GUARDA CUANDO SE IMPRIME (ENTENDIENDOSE QUE ES CUANDO SE LE ENVIA AL CLIENTE9
'    ACTUCOTIZACION% = 0
'    If XVALO(lblCOTIZADO) > 0 Then
'       If XVALO(lblCOTIZADO) <> XVALO(TICAMBIS) Then
'          OPX% = COMALTER%("Esta Cotización Yá Fué Generada Con la Cotización del Dolar a: " & lblCOTIZADO & "\\Guardar Sin Actualizar Cotización\Guardar con Cotización Actual")
'          If OPX% > 1 Then
'             lblCOTIZADO = TRIM$(FORMATNUM$(XVALO(TICAMBIS), "F12.2"))
'          End If
'       End If
'    Else
'       lblCOTIZADO = TICAMBIS
'    End If
'    Condi$ = "NUMREPA = " & NROREPA&
'    Call ACTUREGISTRO("COTIEYM", "COTIZA_DOLAR", Condi$, FORMATNUM$(XVALO(lblCOTIZADO), "F12.2"), REGAF&, "NOMESS")
'    Call ACTUREGISTRO("REPARA", "TIPOCAM", Condi$, FORMATNUM$(XVALO(lblCOTIZADO), "F12.2"), REGAF&, "NOMESS")

    '
'     '-----------------------------------
'     NUMESER$ = TRIM$(lblnserie)
'     CUERPO$ = "Cotización de Asist.Técnica" & vbCrLf
'     CUERPO$ = CUERPO$ + "Número de Asist.Técnica: " & NROREPA& & vbCrLf
'     CUERPO$ = CUERPO$ + "Cliente: " & lblrasocli & vbCrLf
'     CUERPO$ = CUERPO$ + "Importe u$s: " & TRIM$(TXTFINAL.TEXT) & " ($" & TRIM$(FORMATNUM$(XVALO(TXTFINAL) * XVALO(TICAMBIS), "F12.2")) & ")" & vbCrLf
'     CUERPO$ = CUERPO$ + " (Desc.Incluido: " & PORDESCU.TEXT & " %)" & vbCrLf
'     If XVALO(TXTTOTALCONACCESORIOS) > 0 Then CUERPO$ = CUERPO$ + "+ Accesorios: " & TRIM$(FORMATNUM$(XVALO(TXTTOTALCONACCESORIOS), "F12.2")) & vbCrLf
'     CUERPO$ = CUERPO$ + "Cotiz. Actual U$S: " & TRIM$(FORMATNUM$(TICAMONE(2), "F12.2")) & vbCrLf
'     CUERPO$ = CUERPO$ + "Presup Cotizado C/U$S: " & TRIM$(lblCOTIZADO)
'
'     Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")
'     Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NROREPA&, 2, REGAF&, "NOMESS")
'     '-----------------------------------

    With VENTABNEW.MSF
       .Rows = 11: .Cols = 4
       .TextMatrix(1, 1) = Me.lblcontacto
       .TextMatrix(2, 1) = Me.lblrasocli
       .TextMatrix(3, 1) = Me.Label25
       .TextMatrix(4, 1) = Me.LBLCODIGO
       .TextMatrix(5, 1) = Me.lblnserie
       .TextMatrix(6, 1) = "" 'acA DEBERIA A TENER QUE IR FALLA INDICADA POR EL CLIENTE
       .TextMatrix(7, 1) = TRIM$(FORMATNUM$(XVALO(Me.TXTTOTAL) * XVALO(lblCOTIZADO), "F12.2")) 'LO PASO A $
       .TextMatrix(8, 1) = TRIM$(FORMATNUM$(XVALO(Me.TXTIMPODESC) * XVALO(lblCOTIZADO), "F12.2"))
       .TextMatrix(9, 1) = ""
       .TextMatrix(10, 1) = ""
       'AGREGA DETALLE DE REPARACION
       RetengoValorMSF1% = MSF1.Rows
       If MSF1.Rows < 12 Then
          MSF1.Rows = 12
       End If
       
       For J& = 1 To Me.MSF1.Rows - 1
              i& = J& + 10  'PARA QUE ARRANQUE A PARTIR DE LA FILA 11
              .Rows = i& + 1
              .TextMatrix(i&, 1) = Me.MSF1.TextMatrix(J&, 1)
              .TextMatrix(i&, 2) = Me.MSF1.TextMatrix(J&, 2)
              .TextMatrix(i&, 3) = Me.MSF1.TextMatrix(J&, 3)
        Next J&
       RetengoValorMSF2% = MSF2.Rows
       If MSF2.Rows < 6 Then
          MSF2.Rows = 6
       End If
       
        
       'AGREGA DETALLE DE ACCESORIOS
        For k& = 1 To Me.MSF2.Rows - 1
               REG& = k& + i& + 10  'PARA QUE ARRANQUE A PARTIR DE LA FILA 11
              .Rows = REG& + 1
              .TextMatrix(REG&, 1) = Me.MSF2.TextMatrix(k&, 1)
              .TextMatrix(REG&, 2) = Me.MSF2.TextMatrix(k&, 2)
              IMPOACC# = XVALO(Me.MSF2.TextMatrix(k&, 3))
              IMPOACCESO$ = TRIM$(FORMATNUM$(IMPOACC# * XVALO(lblCOTIZADO), "F12.2")) 'LO PASO A $
              .TextMatrix(REG&, 3) = IMPOACCESO$
        Next k&
    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
        
'    Call Param_Print_Excel(RUTAEXCEL$, "COTIZACION", VENTABNEW.MSF, PDES$, "1", 1, 4, "", 1, 1)              'imprime en excel la primera columna de mi msf en la columna 4 a partir de la fila 1 en el excel (este graba pero no lo imprime)
    Call Param_Print_Excel(RUTAEXCEL$, "COTIZACION", VENTABNEW.MSF, PDES$, "1;2;3", 1, 4, "", 1, 0)                 'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW
    MSF1.Rows = RetengoValorMSF1%
    MSF2.Rows = RetengoValorMSF2%
End Sub



Private Sub MONEMIS_Change()
   '
   '\\\OT-07-0086-RG-06/03/07///

End Sub

Private Sub MONEMIS_DblClick()

   '
   Static CAMON%, DMONE$(), TCMONE()
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
10   If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         U2& = 0
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           If TRIM$(DMX$) <> "" Then
             VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
             U2& = U2& + 1
             DMONE$(U2&) = DMX$
             TCMONE(U2&) = VHX
           End If
         Next U&
         CAMON% = U2&
         If CAMON% < 1 Then GoTo 10
     End If
   End If
   '
   MONEMIA% = MONEMI%
   MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   MONEMIS = DMONE$(MONEMI%)
'   COEFICO = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   MONEMIS.ForeColor = &H0&
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = &H404080
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   TICAMBIS.Visible = False
   If MONEMI% > 1 Then
     TICAMBIS.Visible = True:
   End If
   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBX# = TICAMONE(MONEMI%)
   TICAMBIS = FORMATNUM(TICAMBX#, "F9.4")

End Sub





Private Sub MSF1_DblClick()
    i& = MSF1.MouseRow
    J& = MSF1.MouseCol
    If i& = 0 Or i& > MSF1.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF1.Row = i&
    MSF1.COL = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    '

    TXTORDEN.TEXT = MSF1.TextMatrix(REG&, 1)
    Me.TXTFALLAPRES = MSF1.TextMatrix(REG&, 2)
    Me.TXTTAREACTE = MSF1.TextMatrix(REG&, 3)
    Me.TXTDESCRIEYM = MSF1.TextMatrix(REG&, 4)
    Me.TXTCOSTOREP = MSF1.TextMatrix(REG&, 5)
    REXULTADO = MSF1.TextMatrix(REG&, 5)
    Me.TIEMPOMOB = MSF1.TextMatrix(REG&, 6)
    Label2(1).Caption = MSF1.TextMatrix(i&, 7)
    
    '
    Call SELECCIONARFILA(MSF1, REG&)
    
    


End Sub

Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
      REG& = MSF1.MouseRow
      If MSF1.Rows > 1 Then
        If REG& >= 1 Then
            Call SELECCIONARFILA(Me.MSF1, REG&)
            OPX% = COMALTER("Realmente Desea Eliminar el Registro Seleccionado\\Cancelar\Sí, Eliminar")
            If OPX% = 2 Then
              Call cmdEliminaItem(MSF1, REG&)
              TXTFINAL = FORMATNUM$(CALCULARTOTAL#, "F12.2")
            End If
        End If
      End If
    End If
    
End Sub





Private Sub MSF2_DblClick()
        'TENER EN CLARO QUE LOS ACCESORIOS QUE SE INGRESAN SON SOLO A MODO DE COTIZACION
        'EN EL CASO QUE SE DEBAN FACTURAR LOS MISMOS DEBERAN CARGARSE POR PEDIDO O BIEN FACTURAR DIRECTAMENTE.
        
        VENTANA.Campox = "Código/A16;Descripción/A48;Importe/f12.2"
        VENTANA.Inicializar = 0
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "COMBINACIÓN DE EQUIPOS CON ACCESORIOS (TASA DE IVA)"
        VENTANA.CampEditables = "1"
        VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER;STATUS >=0"
        VENTANA.Inicializar = 1
        
        Call CopiarMsfAOtroMsf(MSF2, VENTABNEW.MSF, 1)
        '****************************************************************
        Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
30      VENTABNEW.Show 1
       

        Dim MSF As Object
        Set MSF = VENTABNEW.MSF
        '
        If VENTABNEW.APROBADO% < 1 Then Exit Sub

        If VENTABNEW.APROBADO% > 0 Then
           If FilaVacia%(VENTABNEW.MSF, 1) < 1 Then
               VALIEC1& = VALIDAR_ECO(VENTABNEW.MSF, 1, 0)
               If VALIEC1& > 0 Then
                 Call COMUNI("Faltan Datos En Fila: " & VALIEC1&)
                 GoTo 30
               End If
          End If
        End If
        '
        NLISTA% = XVALO(Control$("REPARA", "LIPREREP"))
        If VENTABNEW.MSF.Rows < 2 Then
           MSF2.Rows = 1
           GoTo 20 'antes el 20 estaba al lado de Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF2, 1)
                   'y el acturegistro antes de esa fila o sea que si entraba aca lo salteaba
                   'por eso cuando grababa si vaciaba la grilla no hacia el acturegistro marborra = 1
        End If

        Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF2, 1)
        'TRANSACCIÓN GRABACION
20       Call ACTUREGISTRO("ACCESORIOSCOTIZ", "MARBORRA", "NUMREPA = " & XVALO(Label25), 1, REGAF&, "NOMESS")

        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans
        For REG& = 1 To MSF2.Rows - 1
          CI1% = CODINT%(MSF2.TextMatrix(REG&, 1))
          IMPORTE = FORMATNUM$(XVALO(MSF.TextMatrix(REG&, 3)), "F12.2")
          If CI1% > 0 Then
             SQLX$ = "INSERT INTO ACCESORIOSCOTIZ "
             SQLX$ = SQLX$ + " (NUMREPA,Cod_Inte,PrUn_Neto,CODIGO,DESCRIPCION,CANTIDAD) "
             SQLX$ = SQLX$ + " VALUES(" & XVALO(Label25)
             SQLX$ = SQLX$ + "," & CI1%
             If XVALO(IMPORTE) < 0.005 Then
               IMPORTE = FORMATNUM$(PRELISTA#(TRIM$(Str$(NLISTA%)), CI1%), "f12.2")
               MSF2.TextMatrix(REG&, 3) = FORMATNUM$(IMPORTE, "F12.2")
               If XVALO(IMPORTE) < 0.005 Then Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF2, vbYellow, REG&)
             End If
             SQLX$ = SQLX$ + " ,'" & IMPORTE & "'"
             SQLX$ = SQLX$ + " ,'" & MSF2.TextMatrix(REG&, 1) & "'"
             SQLX$ = SQLX$ + " ,'" & MSF2.TextMatrix(REG&, 2) & "'"
             SQLX$ = SQLX$ + " ,'1'"
             
             SQLX$ = SQLX$ + ")"
             FLEXCONN.Execute (SQLX$)
          End If
        Next REG&
        Set VENTANA = Nothing
        lbltotaccesorios = FORMATNUM$(SUMA_COLUMNA_SELECCION(MSF2, 3), "F12.2")
        Call BORRAREGISTROS("ACCESORIOSCOTIZ", "NUMREPA = " & XVALO(Label25) & " and MARBORRA = 1", REGAF&, "NOMESS")
        
'        VER SI SE ESTA GURADAND

ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       MsgBox Err.Description
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
    End If

End Sub


Private Sub PORDESCU_Change()
   DESC$ = TRIM$(PORDESCU.TEXT)
   PODES# = COEFDESCU(DESC$)
   TTT1# = (XVALO(Me.TXTFINAL) * PODES# / 100)
   TXTIMPODESC = FORMATNUM$(XVALO(TTT1#), "F12.2")

   TXTTOTALCONACCESORIOS = FORMATNUM$(XVALO(Me.TXTFINAL) - TTT1#, "F12.2")

End Sub
Private Sub TEXT211_Change()
     Eco21.TEXT = ECOARCH$("CONPAG", MKI$(XVALO(TEXT211.TEXT)))
'     HAY QUE GUARDAR LA CONDICION DE VENTA PARA ARRASTRAR A FACTURACION
End Sub


Private Sub TICAMBIS_DblClick()
   If DERACCE%("COTILOCA", "Ajustar Tipo de Cambio Local") >= 2 Then
      ENTRATICA.Frame4.Caption = "Cotización Local para " + MONEMIS.Caption
      ENTRATICA.Text4 = TICAMBIS.Caption
      ENTRATICA.Text4 = FORMATNUM(XVALO(ENTRATICA.Text4), "F10.4")
      ENTRATICA.Show 1
      If TRIM$(ENTRATICA.Text4) <> "" Then
        TICAMBIS = TRIM$(ENTRATICA.Text4)
      End If
   End If
   '
End Sub

Private Sub TXTFINAL_Change()
    Call PORDESCU_Change
End Sub

Private Sub TXTTAREACTE_Change()
   Text5 = ""
   Me.TIEMPOMOB = ""
   Me.TXTCOSTOREP = ""
   
End Sub






Sub GRABATIEMPMOBRA(CI1%)
   CONDI$ = "CODICONJ = " & CI1%
   '
   ULTIMAFEMOD% = 0
   'SI ES LA PRIMERAVEZ QUE SE GENERA LA LISTA LE AGREGO EN EL MASTER LA FECHA DE CREACION
   FEHO% = HOY(HO$)
   If CantRegistros("TIEMOBRA", CONDI$, "NOMESS") < 1 Then
      Call ACTUREGISTRO("MASTER", "FECHALTA_MO", "Codint = " & CI1%, FETEXCOR1$(FEHO%), REGAF&, "NOMESS")
   End If
   '
   'MARCO LOS REGISTROS CON MARBORRA
   Call ACTUREGISTRO("TIEMOBRA", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
   '
   SQLX$ = "SELECT * FROM TIEMOBRA "
   SQLX$ = SQLX$ + " WHERE CODICONJ < 1"
   '
   Set Rst = New ADODB.Recordset
   Rst.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   
   With Rst
       For i& = 1 To MSF1.Rows - 1
        NUOR% = TRIM$(MSF1.TextMatrix(i&, 1))
        DESCRIX$ = MSF1.TextMatrix(i&, 2)
        DESCRIX$ = LTrim(DESCRIX$)
        DESCRIX$ = RTrim(DESCRIX$)
        TIEM1# = XVALO(MSF1.TextMatrix(i&, 6))
        CANTI = XVALO(MSF1.TextMatrix(i&, 5))
        CODSECCION = XVALO(MSF1.TextMatrix(i&, 3))
        FEMODI% = CVINT(MSF1.TextMatrix(i&, 9))
        If ULTIMAFEMOD% < FEMODI% Then ULTIMAFEMOD% = FEMODI%

        '
        .AddNew
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        '
        !CODICONJ = CI1%
        !NUORD = NUOR%
        !Descripcion = DESCRIX$
        !TIEMPOUNITARIO = TIEM1#
        !CantiOps = CANTI
        !SecciPro = CODSECCION
        !FECHAMODIF = FETEXCOR1$(FEMODI%)
           User% = USNBR% Mod 100
           If User% < 1 Then User% = 1
        !Usua_Mod = User%
        !MARBORRA = 0
        !FECHA_VISTO = CVDAT(HOY(HO$))
        .Update
       Next i&
    End With
    Rst.Close
    Set Rst = Nothing
    TIEMPX = TRIM$(FORMATNUM$(XVALO(Text161) / 60, "F12.2"))
    Call ACTUREGISTRO("MASTER", "MANO_OBRA_REAL", "CODINT = " & CI1%, TIEMPX, REGAF&, "NOMESS")
    '
    Call BORRAREGISTROS("TIEMOBRA", CONDI$ & " AND MARBORRA = 1", REGAF&, "NOMESS")
    ANACOSI09.Text4 = TIEMPX
    '
    'ACUTALIZO ESTOS DATOS EN PANTALLA
'    Me.UsuarioModifico = USERNAM$(User%)
'    Me.TXTMODIFICADO = FECHATEX$(ULTIMAFEMOD%)
'    Me.TXTFECHAVISTO = FECHATEX$(HOY(HO$))
    
    '
End Sub

Private Sub MSF1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If ARBUS$ <> "" Then
          VUELTA% = 0
          J& = MSF1.Row

          h& = 0: If MSF1.COL < MSF1.Cols Then h& = MSF1.COL + 1
10        For U& = J& To MSF1.Rows - 1
             HI& = h&: If U& > J& Then HI& = 0
             For i& = HI& To MSF1.Cols - 1
              If InStr(UCase$(MSF1.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
                MSF1.Row = U&: MSF1.COL = i&
                On Error Resume Next
                 VISIROWS = (MSF1.Height / MSF1.RowHeight(J&)) - 4
                If Err > 0 Then
                 VISIROWS = 1
                End If
                On Error GoTo 0
                 If VISIROWS > 0 Then
                    If MSF1.TopRow < MSF1.Row - VISIROWS Then
                      MSF1.TopRow = MSF1.Row - VISIROWS
                    End If
                 End If
                MSF1.SetFocus
                Exit Sub
              End If
              Next i&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If

End Sub



Private Sub TXTTOTAL_Change()
   TXTFINAL = FORMATNUM$(XVALO(TXTTOTAL) + XVALO(lbltotaccesorios), "f12.2")
   
End Sub


